
# HelloID-Task-SA-Target-AzureActiveDirectory-GroupGrantMembership

## Prerequisites

Before using this snippet, verify you've met with the following requirements:

- [ ] AzureAD app registration
- [ ] The correct app permissions for the app registration
- [ ] User defined variables: `AADTenantID`, `AADAppID` and `AADAppSecret` created in your HelloID portal.
- [ ] Please see our documentation on how to create custom variables: (https://docs.helloid.com/en/variables/custom-variables.html)

## Description

This code snippet executes the following tasks:

1. Define a hash table `$formObject`. The keys of the hash table represent the properties to add a user account to a group, while the values represent the values entered in the form. 

> To view an example of the form output, please refer to the JSON code pasted below.

```json
{
    "GroupIdentity": "1def3157-d712-4ee8-a458-c7b3202431f2",
    "MembersToAdd": [
        {
            "userPrincipalName": "testuser1@mydomain.local"
        },
        {
            "userPrincipalName": "testuser2@mydomain.local"
        }
    ]
}
```

> :exclamation: It is important to note that the names of your form fields might differ. Ensure that the `$formObject` hashtable is appropriately adjusted to match your form fields.
> [See the Microsoft Docs page](https://learn.microsoft.com/en-us/graph/api/group-post-members?view=graph-rest-1.0&tabs=http#examples)

2. Receive a bearer token by making a POST request to: `https://login.microsoftonline.com/$AADTenantID/oauth2/token`, where `$AADTenantID` is the ID of your Azure Active Directory tenant.

3. update the attribues of a group using the: `Invoke-RestMethod` cmdlet. The hash table called: `$formObject` is passed to the body of the: `Invoke-RestMethod` cmdlet as a JSON object.

