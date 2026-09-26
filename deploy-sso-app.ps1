# Defining the log file
$log_file = "$PSScriptRoot\Logs\log_deploy_sso_app.log"

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

# Attempting to import the relevant Graph modules
try {
    Import-Module Microsoft.Graph.Applications -ErrorAction Stop

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Imported the required Graph modules"
}

# Terminating script execution if the Graph module couldn't be imported
catch {
    Write-Warning "Failed to import required modules: $_"

    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] Failed to import the required Graph modules: $_"

    exit 1
}

# Attempting to import the relevant Azure Key Vault module
try {
    if (Get-Module -Name "Az.KeyVault" -ListAvailable) {
        Import-Module Az.KeyVault -ErrorAction Stop
    }

    else {
        Install-Module Az.KeyVault -ErrorAction Stop
    }

    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Successfully imported module: Az.KeyVault"
}

# Stop executing script if the Az.KeyVault module couldn't be imported
catch {
    Write-Warning "Failed to import module: Az.KeyVault: $_"

    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] Failed to import module: Az.KeyVault: $_"

    exit 1
}

# Connect to Azure account & Microsoft Graph for the specified tenant and
# request permissions to create and manage app registrations/enterprise apps
try {
    Connect-AzAccount

    Connect-MgGraph -Scopes "Application.ReadWrite.All","Group.Read.All" -ContextScope Process -ErrorAction Stop

    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Successfully connected to Microsoft Graph"
}

# Terminating script execution if the connection to Microsoft Graph failed
catch {
    Write-Warning "Failed to connect to Microsoft Graph: $_"

    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] Failed to connect to Microsoft Graph: $_"

    exit 1
}

# Defining the application name, redirect URI, and security group name that should have access to the app
$application_name = Read-Host "Enter the app registration name: "
$web_redirect_uri = Read-Host "Enter the redirect URI for the app registration: "
$assigned_security_group = Read-Host "Enter the security group name: "

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

    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] App Registration created successfully"
}

# Terminating the script execution if the app registration couldn't be created
catch {
    Write-Warning "Failed to create app registration: $_"

    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] Failed to create app registration: $_"

    exit 1
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

    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Created client secret for app registration"
}

# Terminating the script execution if the client secret couldn't be added to the app registration
catch {
    Write-Warning "Failed to add client secret to app registration: $_"

    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] Failed to add client secret to app registration: $_"

    exit 1
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

    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Created the service principal"
}

# Terminating the script execution if the service principal couldn't be created
catch {
    Write-Warning "Failed to create service principal: $_"

    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] Failed to create service principal: $_"

    exit 1
}

# Extracting the object ID for this service principal
$service_principal_obj_id = $service_principal.Id

# Configuring the setting to only allow assigned members access to the app
try {
    Update-MgServicePrincipal -ServicePrincipalId $service_principal_obj_id -AppRoleAssignmentRequired:$true -ErrorAction Stop
}

# Terminating the script execution if the the setting couldn't be modified
catch {
    Write-Warning "Failed to configure setting for group assignment: $_"

    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] Failed to configure setting for group assignment: $_"

    exit 1
}

# Quering Entra ID to find the security group specified
try {
    $security_group_in_entra = Get-MgGroup -Filter "DisplayName eq '$assigned_security_group'" -ErrorAction Stop
}

# Terminating the script execution if the security group couldn't be found
catch {
    Write-Warning "Failed to find the security group: $_"

    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] Failed to find the security group: $_"

    exit 1
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

    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Assigned security group to service principal"
}

# Terminating the script execution if there was an issue with assigning the security group to the service principal
catch {
    Write-Warning "Failed to assign security group to service principal: $_"

    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] Failed to assign security group to service principal: $_"

    exit 1
}

# Creating a ps custom object to define the tenant ID, app ID, and client secret, for easy retrieval.
$app_reg_credentials = [PSCustomObject]@{
    TenantId = $tenant_id
    AppId = $app_id
    ClientSecret = (ConvertTo-SecureString $client_secret -AsPlainText -Force)
}

# Formatting the results in the form of a table, using the custom object
$app_reg_credentials | Format-Table

# Only retrieve the key vault name from .env file if the .env file exists
if (Test-Path $env_file) {
    # Extracting the key vault name from the .env file
    $key_vault_name = (ConvertFrom-StringData (Get-Content -Raw $env_file)).KEY_VAULT_NAME
}

# Stop script execution if .env file doesn't exist
else {
    Write-Warning "Failed to open .env file: $_"

    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] Failed to open .env file: $_"

    exit 1
}

# Adding the app credentials to Azure Key Vault
try {
    Set-AzKeyVaultSecret -VaultName $key_vault_name -Name $app_reg_credentials.AppId -SecretValue $app_reg_credentials.ClientSecret -ErrorAction Stop

    Add-Content -Path $env_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Successfully added the app credentials to the key vault"
}

# Showing the error message if the .env file couldn't be updated
catch {
    Write-Warning "Failed to update key vault with app credentials: $_"

    Add-Content -Path $log_file -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [ERROR] Failed to update key vault with app credentials: $_"
}