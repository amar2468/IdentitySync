# Connect to Microsoft Graph for the specified tenant and request permissions to manage users and group memberships.
Connect-MgGraph -TenantId "3d9354b2-4be9-4ccd-b294-eefdbe22906a" -Scopes "User.ReadWrite.All","GroupMember.ReadWrite.All"

# Importing the CSV file with the employee data
$employee_data = Import-Csv -Path ".\employee-data-hr.csv"

# Retrieving all users within the Entra ID tenant, and extracting the user principal name for each
$all_users_in_entra = Get-MgUser | Select-Object -ExpandProperty UserPrincipalName

# Mapping departments against the group that the employee should be assigned to
$department_mappings = @{
    "IT" = @{
        Group = "IT"
    }

    "Sales" = @{
        Group = "Sales"
    }

    "Human Resources" = @{
        Group = "HR"
    }

    "Management" = @{
        Group = "Management"
    }

    "Finance" = @{
        Group = "Finance"
    }
}

# Initialising an empty lookup table for all employees
$employee_lookup = @{}

# Creating a lookup table using EmployeeID as the key, allowing employee records
# to be retrieved efficiently without iterating through the CSV each time.
foreach ($employee in $employee_data) {
    $employee_lookup[$employee.EmployeeID] = $employee
}

Write-Host "Creating users in Entra ID..." -ForegroundColor Yellow

# Iterating through each employee, with the goal of adding them to Entra ID as a user.
foreach ($employee in $employee_data) {
    # Extracting the employee email from the CSV file
    $employee_email = $employee.Email

    # Building a hashtable with the parameters that should be present for the user in Entra ID
    $entra_params = @{
        EmployeeId = $employee.EmployeeID
        DisplayName = "$($employee.FirstName) $($employee.LastName)"
        GivenName = $employee.FirstName
        Surname = $employee.LastName
        PasswordProfile = @{
            Password = "WelcomeToEntra2026#"
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
        MailNickname = ($employee_email -split "@")[0]
        UserPrincipalName = $employee_email
        UsageLocation = $employee.UsageLocation
    }

    # If the employee is not in Entra ID, create the user in Entra ID
    if ($all_users_in_entra -notcontains $employee_email) {
        New-MgUser @entra_params

        $department_profile = $department_mappings[$employee.Department]

        $user = Get-MgUser -UserId $employee_email

        $group = Get-MgGroup -Filter "displayName eq '$($department_profile.Group)'"

        New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $user.Id
    }
}

# Iterating through each employee, with the goal of setting the manager for each employee.
foreach ($employee in $employee_data) {

    # Check if the manager ID is empty, null, or contains only spaces. 
    # If it does, it skips the step of setting the manager and moves on to new employee
    if ([string]::IsNullOrWhiteSpace($employee.ManagerID)) {
        continue
    }

    # Extracting the employee email from the CSV file
    $employee_email = $employee.Email

    # Getting the user object from Entra ID
    $user = Get-MgUser -UserId $employee_email

    # Getting the manager ID for this employee
    $manager_record = $employee_lookup[$employee.ManagerID]

    # Retrieving the user record for the manager in Entra ID
    $manager_in_entra = Get-MgUser -UserId $manager_record.Email

    $manager_params = @{
        UserID = $user.Id
        BodyParameter = @{
            "@odata.id" = "https://graph.microsoft.com/v1.0/users/$($manager_in_entra.Id)"
        }
    }

    # Setting the manager ID for this employee using the Entra ID GUID for the manager
    Set-MgUserManagerByRef @manager_params
}