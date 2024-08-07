# MSGraph authorization with OAuth2 flow

In this example, you will learn how to send emails using Microsoft Outlook Online using MSGraph. The
sending is the easy part, but now we do an authorization which relies on tokens. And that's usually done through the OAuth2.
[Authorization Code Grant](https://oauth.net/2/grant-types/authorization-code/) flow.


## Tasks

Before sending an e-mail with `Send Test Email`, you have to create the access and refresh tokens which are stored to the Control Room Vault by running `Init Microsoft OAuth` task first. This initializer step is required once, then you can send as many e-mails as you want with the tokens already configured.

### Microsoft Entra OAuth/Email tasks

1. `Init Microsoft OAuth`: Authenticate user, authorize app and have the tokens
   generated automatically in the Vault.
2. `Send Test Email`: Send an email using MSGraph.

## Client app setup

You need to register an app which will act on behalf of your account. The app
(Client) is the entity sending e-mails instead of you (User). But you need to
authenticate yourself and authorize the app first in order to allow it to send
e-mails for you. For this, certain settings are required:

### Microsoft Entra app registration

1. Go to Microsoft Entra (formerly Azure Active Directory) *[App registrations](https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps)*
   page and follow [these](https://learn.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
   app configuration instructions.
3. Ensure you created a "Web" app and have the following checked:
   - Is a *private* **single** or **multi-tenant** app.
   - The type of the application is **Web App**.
   - Redirect URI can be: `https://login.microsoftonline.com/common/oauth2/nativeclient`
   - Has at least the following **MSGraph** permission(s) enabled:
     - **Delegated**: `Mail.Send` (MSGraph -> Mail)

   - Create a client secret and take note of these credentials (client ID, client secret & tenant ID), as you need
     them later on.


## Variables setup

To use secrets from the Vault and send email you need to set certain variables. You can set them up by modifying `variables.py` file in project root. Set the `SECRET_NAME` variable to match your Control Room Vault `name`. Set also a working email addresse(s) to the `RECIPIENT` variable, separated by comma if more than one. 

## Vault setup

The client ID, client secret and tenant ID obtained from Entra needs to be stored securely in the Vault,
as they'll be used automatically by the automation in order to obtain the tokens. The
token entries in Vault will be created during the authentication flow and will be refreshed automatically later on by the automation. 

### Online Control Room Vault

Create a secret called e.g. `MSGraph` in Control Room's Vault with the
following entries (and make sure to connect **VSCode** to the online Vault):

- `tenant_id`: Your Microsoft Entra tenant ID.
- `client_id`: Your application/client ID.
- `client_secret`: Your app client secret.
- `access_token`: Optional, you can create this entry and leave it blank since this will be overridden by the task. 
- `refresh_token`: Optional, You can create this entry and leave it blank since this will be overridden by the task. 

## Task run

Run with **VSCode** or **rcc** the following tasks in order:

1. `Init Microsoft OAuth`: Opens a browser window for you to authenticate and
   finally getting a redirect response URL in the address bar. Once you get there, the
   browser closes and the token gets generated and updated in the Vault.
   - Now you should see your brand new `access_token` and `refresh_token` fields created/updated in the Vault.
     (keep it private as this is like a password which grants access into your e-mail)
   - This step is required to be run once, requires human intervention (attended) and
     once you get your token generated, it will stay valid (by refreshing itself)
     indefinitely.
2. `Send Test Email`: Sends a test e-mail to the recipient(s) listed in the `variables.py` given the credentials
   configured in the Vault during the previous step.
   - This step can be fully automated and doesn't require the first step run each time.
     As once the tokens are set, it remains available until you revoke the refresh
     token or remove the app.

## Remarks

- Access token lifetime:
  - While using this example the token refreshes itself automatically (internally handled by
    the libraries) and is automatically updated into the Vault as well.
- Learn more about OAuth2:
  - [Microsoft](https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-auth-code-flow)