# HelloID-Task-SA-Target-AzureActiveDirectory-GroupGrantMembership
##################################################################
# Form mapping
$formObject = @{
    groupId     = $form.groupId
    membersToAdd = $form.membersToAdd  
}
try {
    Write-Information "Executing AzureActiveDirectory action: [GroupGrantMembership] for: [$($formObject.groupId)]"
    Write-Information "Retrieving Microsoft Graph AccessToken for tenant: [$AADTenantID]"
    $splatTokenParams = @{
        Uri         = "https://login.microsoftonline.com/$($AADTenantID)/oauth2/token"
        ContentType = 'application/x-www-form-urlencoded'
        Method      = 'POST'
        Body        = @{                                                                                                                         
            grant_type    = 'client_credentials'
            client_id     = $AADAppID
            client_secret = $AADAppSecret
            resource      = 'https://graph.microsoft.com'
        }
    }
    $accessToken = (Invoke-RestMethod @splatTokenParams).access_token

    $headers = [System.Collections.Generic.Dictionary[string, string]]::new()
    $headers.Add("Authorization", "Bearer $($accessToken)")
    $headers.Add("Content-Type", "application/json")

    foreach ($member in $formObject.membersToAdd){
        try {
            $splatAddMemberToGroup = @{
                Uri         = "https://graph.microsoft.com/v1.0/groups/$($formObject.groupId)/members/`$ref"
                ContentType = 'application/json'
                Method      = 'POST'
                Headers     = $headers
                Body        = @{ '@odata.id' = "https://graph.microsoft.com/v1.0/users/$($member.userPrincipleName)" } | ConvertTo-Json -Depth 10
            }
            $null = Invoke-RestMethod @splatAddMemberToGroup

            $auditLog = @{
                Action            = 'UpdateResource'
                System            = 'AzureActiveDirectory'
                TargetIdentifier  = $formObject.groupId 
                TargetDisplayName = ''
                Message           = "AzureActiveDirectory action: [GroupGrantMembership] to group [$($formObject.groupId)] for: [$($member.userPrincipleName)] executed successfully"
                IsError           = $false
            }
    
            Write-Information -Tags 'Audit' -MessageData $auditLog
            Write-Information "AzureActiveDirectory action: [GroupGrantMembership] to group [$($formObject.groupId)] for: [$($member.userPrincipleName)] executed successfully"
        }
        catch {
            $ex = $_
            $auditLog = @{
                Action            = 'UpdateResource'
                System            = 'AzureActiveDirectory'
                TargetIdentifier  = $formObject.groupId 
                TargetDisplayName = ''
                Message           = "Could not execute AzureActiveDirectory action: [GroupGrantMembership] to group [$($formObject.groupId)] for: [$($member.userPrincipleName)], error: $($ex.Exception.Message)"
                IsError           = $true
            }
            if ($ex.Exception.Response.StatusCode -eq 404) { 
                $auditLog.Message = "Could not execute AzureActiveDirectory action: [GroupGrantMembership] for [$($formObject.groupId)], the specified group does not exist in the Azure Active Directory." 
                Write-Information -Tags "Audit" -MessageData $auditLog
                Write-Error "$($auditLog.Message)"
                break
            }
            Write-Information -Tags "Audit" -MessageData $auditLog
            Write-Error "Could not execute AzureActiveDirectory action: [GroupGrantMembership] to group [$($formObject.groupId)] for: [$($member.userPrincipleName)], error: $($ex.Exception.Message)"
        }
    }
}
catch {
    $ex = $_
    $auditLog = @{
        Action            = 'UpdateResource'
        System            = 'AzureActiveDirectory'
        TargetIdentifier  = $formObject.groupId 
        TargetDisplayName = ''
        Message           = "Could not execute AzureActiveDirectory action: [GroupGrantMembership] for: [$($formObject.groupId)], error: $($ex.Exception.Message)"
        IsError           = $true
    }
    
    if ($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') {
        $auditLog.Message = "Could not execute AzureActiveDirectory action: [GroupGrantMembership] for [$($formObject.groupId)], error: $($ex.ErrorDetails)" 
    } elseif ($ex.Exception.Response.StatusCode -eq 404) { 
        $auditLog.Message = "Could not execute AzureActiveDirectory action: [GroupGrantMembership] for [$($formObject.groupId)], the specified group does not exist in the Azure Active Directory." 
    } else { 
        $auditLog.Message = "Could not execute AzureActiveDirectory action: [GroupGrantMembership] for [$($formObject.groupId)], error: $($ex.Exception.Message)" 
    }
    Write-Information -Tags "Audit" -MessageData $auditLog
    Write-Error "$($auditLog.Message)"
}
##################################################################