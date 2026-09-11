# Extract the tenant_id from the .env file
$tenant_id = (ConvertFrom-StringData (Get-Content -Raw .env)).tenant_id

# Terminating script execution if the tenant_id can't be found in the .env file
if ($null -eq $tenant_id) {
    Write-Warning "Tenant ID could not be found in the .env file."

    exit
}

# Attempting to import the relevant Graph modules
try {
    Import-Module Microsoft.Graph.Applications -ErrorAction Stop

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
}

# Terminating script execution if the Graph module couldn't be imported
catch {
    Write-Warning "$($_.Exception.Message)"

    exit
}

# Connect to Microsoft Graph for the specified tenant and request permissions to create and manage app registrations/enterprise apps
try {
    Connect-MgGraph -TenantId $tenant_id -Scopes "Application.ReadWrite.All","Group.Read.All" -ErrorAction Stop
}

# Terminating script execution if the connection to Microsoft Graph failed
catch {
    Write-Warning "$($_.Exception.Message)"

    exit
}

# Defining the application name, redirect URI, and security group name that should have access to the app
$application_name = "Finance Dashboard"
$web_redirect_uri = "http://localhost:5000/getAToken"
$assigned_security_group = "Finance"

# Populating the parameters that are necessary to create the app registration
$app_params = @{
    DisplayName = $application_name
    SignInAudience = "AzureADMyOrg"
    Web = @{
        RedirectUris = @($web_redirect_uri)
    }
    ErrorAction = "Stop"
}

# Creating the app registration using the specified parameters
try {
    $new_app = New-MgApplication @app_params
}

# Terminating the script execution if the app registration couldn't be created
catch {
    Write-Warning "$($_.Exception.Message)"

    exit
}

# Extracting the application (client) ID and the object ID for the newly created app registration
$app_id = $new_app.AppId
$app_obj_id = $new_app.Id

# Populating the parameters that are necessary to generate a client secret that should expire in 1 year.
$secret_params = @{
    ApplicationId = $app_obj_id
    PasswordCredential = @{
        DisplayName = "NewSecret"
        EndDateTime = (Get-Date).AddYears(1)
    }
    ErrorAction = "Stop"
}

# Adding the client secret to this app registration using the specified parameters
try {
    $client_secret_obj = Add-MgApplicationPassword @secret_params
}

# Terminating the script execution if the client secret couldn't be added to the app registration
catch {
    Write-Warning "$($_.Exception.Message)"

    exit
}

# Extracting the client secret text value into the variable
$client_secret = $client_secret_obj.SecretText

# Populating the app ID using the app registration's app ID, linking the service principal to the original app registration.
$service_principal_id = @{
    AppId = $app_id
    ErrorAction = "Stop"
}

# Creating the service principal using the app registration's app ID
try {
    $service_principal = New-MgServicePrincipal -BodyParameter $service_principal_id
}

# Terminating the script execution if the service principal couldn't be created
catch {
    Write-Warning "$($_.Exception.Message)"

    exit
}

# Extracting the object ID for this service principal
$service_principal_obj_id = $service_principal.Id

# Configuring the setting to only allow assigned members access to the app
try {
    Update-MgServicePrincipal -ServicePrincipalId $service_principal_obj_id -AppRoleAssignmentRequired:$true -ErrorAction Stop
}

# Terminating the script execution if the the setting couldn't be modified
catch {
    Write-Warning "$($_.Exception.Message)"

    exit
}

# Quering Entra ID to find the security group specified
try {
    $security_group_in_entra = Get-MgGroup -Filter "DisplayName eq '$assigned_security_group'" -ErrorAction Stop
}

# Terminating the script execution if the security group couldn't be found
catch {
    Write-Warning "$($_.Exception.Message)"

    exit
}

# Extracting the security group object ID from Entra ID
$security_group_obj_id = $security_group_in_entra.Id

# Using the Default Access Role ID to assign the group to an app role within the service principal
$app_role_id = "00000000-0000-0000-0000-000000000000"

# Populating the parameters that are necessary to assign the security group to the service principal
$app_role_params = @{
    GroupId = $security_group_obj_id

    PrincipalId = $security_group_obj_id

    ResourceId = $service_principal_obj_id

    AppRoleId = $app_role_id

    ErrorAction = "Stop"
}

# Assigning the security group to the service principal using the specified parameters
try {
    New-MgGroupAppRoleAssignment @app_role_params
}

# Terminating the script execution if there was an issue with assigning the security group to the service principal
catch {
    Write-Warning "$($_.Exception.Message)"

    exit
}

# Creating a ps custom object to define the tenant ID, app ID, and client secret, for easy retrieval.
$app_reg_credentials = [PSCustomObject]@{
    TenantId = $tenant_id
    AppId = $app_id
    ClientSecret = $client_secret
}

# Formatting the results in the form of a table, using the custom object
$app_reg_credentials | Format-Table

# Updating the credentials in the .env file
try {
    # Updating the client secret within the .env file
    (Get-Content .env -ErrorAction Stop) -replace "^(client_secret\s*=\s*).*", "`${1}$client_secret" | Set-Content .env -ErrorAction Stop

    # Updating the client ID within the .env file
    (Get-Content .env -ErrorAction Stop) -replace "^(client_id\s*=\s*).*", "`${1}$app_id" | Set-Content .env -ErrorAction Stop
}

# Showing the error message if the .env file couldn't be updated
catch {
    Write-Warning "$($_.Exception.Message)"
}