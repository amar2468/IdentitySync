# Extract the tenant_id from the .env file
$tenant_id = (ConvertFrom-StringData (Get-Content -Raw .env)).tenant_id

Import-Module Microsoft.Graph.Applications

# Connect to Microsoft Graph for the specified tenant and request permissions to create and manage app registrations/enterprise apps
Connect-MgGraph -TenantId $tenant_id -Scopes "Application.ReadWrite.All","Group.Read.All"

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
}

# Creating the app registration using the specified parameters
$new_app = New-MgApplication @app_params

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

}

# Adding the client secret to this app registration using the specified parameters
$client_secret = Add-MgApplicationPassword @secret_params

# Populating the app ID using the app registration's app ID, linking the service principal to the original app registration.
$service_principal_id = @{
    AppId = $app_id
}

# Creating the service principal using the app registration's app ID
$service_principal = New-MgServicePrincipal -BodyParameter $service_principal_id

# Extracting the object ID for this service principal
$service_principal_obj_id = $service_principal.Id

# Configuring the setting to only allow assigned members access to the app
Update-MgServicePrincipal -ServicePrincipalId $service_principal_obj_id -AppRoleAssignmentRequired:$true

# Quering Entra ID to find the security group specified
$security_group_in_entra = Get-MgGroup -Filter "DisplayName eq '$assigned_security_group'"

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
}

# Assigning the security group to the service principal using the specified parameters
New-MgGroupAppRoleAssignment @app_role_params

# Creating a ps custom object to define the tenant ID, app ID, and client secret, for easy retrieval.
$app_reg_credentials = [PSCustomObject]@{
    TenantId = $tenant_id
    AppId = $app_id
    ClientSecret = $client_secret.SecretText
}

# Formatting the results in the form of a table, using the custom object
$app_reg_credentials | Format-Table

# Updating the client secret within the .env file
(Get-Content .env) -replace "^(client_secret\s*=\s*).*", "`${1}$client_secret" | Set-Content .env

# Updating the client ID within the .env file
(Get-Content .env) -replace "^(client_id\s*=\s*).*", "`${1}$app_id" | Set-Content .env