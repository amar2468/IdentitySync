# Retrieving the log file
$log_file = "$PSScriptRoot\Logs\log_onboard_users.log"

# Defining the env file
$env_file = "$PSScriptRoot\.env"

# Check if .env file exists
if (Test-Path $env_file) {
    # Extract the tenant_id from the .env file
    $tenant_id = (ConvertFrom-StringData (Get-Content -Raw $env_file)).tenant_id
}

# Stop script execution if .env file doesn't exist
else {
    Write-Warning "Failed to open .env file: $_"

    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] Failed to open .env file: $_"

    exit 1
}

# Terminating script execution if the tenant_id can't be found in the .env file
if ($null -eq $tenant_id) {
    Write-Warning "Tenant ID could not be found in the .env file."

    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] Tenant ID could not be found in the .env file: $_"

    exit 1
}

# Attempting to import the relevant Graph module
try {
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Imported the required module"
}

# Terminating script execution if the Graph module couldn't be imported
catch {
    Write-Warning "Failed to import the required module: $_"

    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] Failed to import the required module: $_"

    exit 1
}

# Connect to Microsoft Graph for the specified tenant and request permissions to manage users and group memberships.
try {
    Connect-MgGraph -TenantId $tenant_id -Scopes "User.ReadWrite.All","GroupMember.ReadWrite.All" -ErrorAction Stop

    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Successfully connected to Microsoft Graph"
}

# Terminating script execution if the connection to Microsoft Graph failed
catch {
    Write-Warning "Failed to connect to Microsoft Graph: $_"

    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] Failed to connect to Microsoft Graph: $_"

    exit 1
}

# Importing the CSV file with the employee data
try {
    $employee_data = Import-Csv -Path ".\employee-data-hr.csv" -ErrorAction Stop

    # Appending information to the log file
    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] CSV Imported"

    # Retrieving all users within the Entra ID tenant, and extracting the user principal name for each
    $all_users_in_entra = Get-MgUser -All -ErrorAction Stop | Select-Object -ExpandProperty UserPrincipalName
    
    # Appending information to the log file
    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] All users retrieved from Entra ID"
}

# Terminate script execution if CSV couldn't be imported or users couldn't be retrieved from Entra ID
catch {
    Write-Warning "Failed to import CSV OR retrieve Entra users: $_"

    # Appending information to the log file
    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] Error encountered: $_"

    exit 1
}

# Mapping departments against the groups that the employee should be assigned to
$department_mappings = @{
    "IT" = @{
        Groups = @("SG-IT-Users", "SG-Share-IT-Read", "SG-Share-IT-Write", "SG-App-IT")
    }

    "Sales" = @{
        Groups = @("SG-Sales-Users", "SG-Share-Sales-Read", "SG-Share-Sales-Write", "SG-App-Sales")
    }

    "Human Resources" = @{
        Groups = @("SG-HR-Users", "SG-Share-HR-Read", "SG-Share-HR-Write", "SG-App-HR")
    }

    "Management" = @{
        Groups = @("SG-Management-Users", "SG-Share-Management-Read", "SG-Share-Management-Write", "SG-App-Management")
    }

    "Finance" = @{
        Groups = @("SG-Finance-Users", "SG-Share-Finance-Read", "SG-Share-Finance-Write", "SG-App-Finance")
    }
}

# Initialising an empty lookup table for all employees
$employee_lookup = @{}

# Creating a lookup table using EmployeeID as the key to allow for quick retrieval of specific attributes of the employee
foreach ($employee in $employee_data) {
    $employee_lookup[$employee.EmployeeID] = $employee
}

Write-Host "Creating users in Entra ID..." -ForegroundColor Yellow

# Appending information to the log file
Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Creating users in Entra ID..."

# Read the default user password from the .env file and save it in a variable
try {
    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Reading .env file and looking for password..."

    # Read .env file and look for the password titled "DEFAULT_PASSWORD"
    $env_content = Get-Content -Path $env_file -ErrorAction Stop
    $match = $env_content | Select-String -Pattern "^DEFAULT_PASSWORD=(.*)$"
    
    # If an occurrence of the "DEFAULT_PASSWORD" key couldn't be found in .env file, throw an error
    if (-not $match) {
        Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] Password not found in the .env file"
        throw "DEFAULT_PASSWORD key was not found in .env file."
    }

    # Extract password value
    $default_password = $match.Matches.Groups[1].Value

    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Password retrieved from .env file"
}
catch {
    Write-Warning "Failed to load default password from .env: $_"

    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] Failed to load default password from .env: $_"

    exit 1
}

# Iterating through each employee, with the goal of adding them to Entra ID as a user.
foreach ($employee in $employee_data) {
    # Extracting the employee email from the CSV file
    $employee_email = $employee.Email

    # If the employee is not in Entra ID, create the user in Entra ID
    if ($all_users_in_entra -notcontains $employee_email) {

        Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $employee_email not in Entra ID. Creating user..."

        # Extracting the email prefix
        $upn_prefix = $employee_email.Split("@")[0]

        # Creating the mail nickname using only lowercase and uppercase letters, numbers, dot, underscore, and dash
        $mail_nickname = $upn_prefix -replace '[^a-zA-Z0-9._-]', ''

        # Building a hashtable with the parameters that should be present for the user in Entra ID
        $entra_params = @{
            EmployeeId = $employee.EmployeeID
            DisplayName = "$($employee.FirstName) $($employee.LastName)"
            GivenName = $employee.FirstName
            Surname = $employee.LastName
            PasswordProfile = @{
                Password = $default_password
                ForceChangePasswordNextSignIn = $true
            }
            JobTitle = $employee.JobTitle
            Department = $employee.Department
            EmployeeType = $employee.EmployeeType
            EmployeeHireDate = [DateTime]::ParseExact(
                $employee.HireDate,
                "dd/MM/yyyy",
                [System.Globalization.CultureInfo]::InvariantCulture
            )
            AccountEnabled = $employee.Status -eq "Active"
            Mail = $employee_email
            MailNickname = $mail_nickname
            UserPrincipalName = $employee_email
            UsageLocation = $employee.UsageLocation
            ErrorAction = "Stop"
        }

        # Creating a user with the specified parameters
        try {
            $newUser = New-MgUser @entra_params

            Write-Host "Successfully created the following user in Entra ID: $employee_email" -ForegroundColor Yellow

            Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Created the following user in Entra ID: $employee_email"
        }

        # Indicating what user accounts didn't get created in Entra ID
        catch {
            Write-Warning "Error encountered when creating the user account for $employee_email : $_"

            Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] Error encountered when creating user account for - $employee_email : $_"

            continue
        }

        # Check if there is a mapping for the department
        if ($department_mappings.ContainsKey($employee.Department)) {
            $department_profile = $department_mappings[$employee.Department]
            
            # Retrieving the security groups for the specific department
            $dept_groups = $department_profile.Groups

            # Iterating through the department groups and adding the user to them
            foreach ($group in $dept_groups) {

                # Retrieving the security group and adding the user to it
                try {
                    $group_id = Get-MgGroup -Filter "displayName eq '$group'" -ErrorAction Stop

                    New-MgGroupMember -GroupId $group_id -DirectoryObjectId $newUser.Id -ErrorAction Stop

                    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Adding $employee_email to the following security group: $group"
                }

                # Indicating what groups the user was not added to
                catch {
                    Write-Warning "User - $employee_email - was not added to this group - $group : $_"

                    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [WARNING] User - $employee_email - was not added to this group - $group : $_"
                }
            }
        }

        # If no mapping was found for the department, a warning will be displayed
        else {
            Write-Warning "No group mapping found for department: $($employee.Department)"

            Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [WARNING] No group mapping found for department: $($employee.Department)"
        }
    }

    # Notify if the user already exists
    else {
        Write-Warning "User - $employee_email - already exists in Entra ID"

        Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] User - $employee_email - already exists in Entra ID"
    }
}

# Iterating through each employee, with the goal of setting the manager for each employee.
foreach ($employee in $employee_data) {

    # Check if the manager ID is empty, null, or contains only spaces. 
    # If it does, it skips the step of setting the manager and moves on to new employee
    if ([string]::IsNullOrWhiteSpace($employee.ManagerID)) {
        Write-Warning "Manager ID in CSV is either empty, null, or contains only spaces - skipping the step to set the manager for $($employee.Email)"

        Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [WARNING] Manager ID in CSV is either empty, null, or contains only spaces - skipping the step to set the manager for $($employee.Email)"

        continue
    }

    # Extracting the employee email from the CSV file
    $employee_email = $employee.Email

    # Getting the user object from Entra ID
    try {
        $user = Get-MgUser -UserId $employee_email -ErrorAction Stop
    }

    # Terminating script execution if the user record couldn't be retrieved from Entra ID
    catch {
        Write-Warning "User - $employee_email - couldn't be retrieved from Entra ID to assign manager: $_"

        Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] User - $employee_email - couldn't be retrieved from Entra ID to assign manager: $_"

        exit 1
    }

    # Retrieving the current manager for the employee
    $check_manager_assigned = Get-MgUserManagerByRef -UserId $user.id -ErrorAction SilentlyContinue

    # If the manager is already assigned for this employee, no action is required and we can move to the next employee
    if ($check_manager_assigned) {
        Write-Warning "User - $employee_email - already has manager assigned - no action taken"

        Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] User - $employee_email - already has manager assigned - no action taken"

        continue
    }

    # Getting the manager ID for this employee
    $manager_record = $employee_lookup[$employee.ManagerID]

    # Retrieving the user record for the manager in Entra ID
    try {
        $manager_in_entra = Get-MgUser -UserId $manager_record.Email -ErrorAction Stop
    }

    # Terminating script execution if the manager record couldn't be retrieved from Entra ID
    catch {
        Write-Warning "The manager - $($manager_record.Email) - couldn't be retrieved from Entra ID: $_"

        Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] The manager - $($manager_record.Email) - couldn't be retrieved from Entra ID: $_"

        exit 1
    }

    $manager_params = @{
        UserID = $user.Id
        BodyParameter = @{
            "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($manager_in_entra.Id)"
        }
        ErrorAction = "Stop"        
    }

    # Setting the manager ID for this employee using the Entra ID GUID for the manager
    try {
        Set-MgUserManagerByRef @manager_params

        Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Successfully set the manager for the following user: $employee_email"
    }

    # Terminating script execution if there was an issue with setting the manager ID for the employee
    catch {
        Write-Warning "Issue with assigning the manager - $($manager_record.Email) - to the user - $employee_email : $_"

        Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] Issue with assigning the manager - $($manager_record.Email) - to the user - $employee_email : $_"

        exit 1
    }
}