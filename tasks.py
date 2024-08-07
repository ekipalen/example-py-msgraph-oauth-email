import requests
from robocorp import log
from robocorp.tasks import task

from oauth_utils import (
    REFRESH_TOKEN,
    app,
    build_headers,
    get_auth_code_using_browser,
    refresh_microsoft_token,
    update_vault,
)
from variables import BASE_GRAPH_URL, BODY, RECIPIENTS, SCOPES, SUBJECT

SCOPES = SCOPES.split(",")


@task
def initial_msgraph_authentication() -> None:
    """
    Perform initial Microsoft Graph authentication to acquire access and refresh tokens.
    """
    with log.suppress_variables():
        try:
            redirect_uri = (
                "https://login.microsoftonline.com/common/oauth2/nativeclient"
            )
            auth_url = app.get_authorization_request_url(
                SCOPES, redirect_uri=redirect_uri
            )
            auth_code = get_auth_code_using_browser(auth_url)
            result = app.acquire_token_by_authorization_code(
                auth_code, scopes=SCOPES, redirect_uri=redirect_uri
            )
            if "access_token" in result:
                access_token = result["access_token"]
                refresh_token = result.get("refresh_token", "")
                update_vault(access_token, refresh_token)
                print("OAuth flow completed successfully")
            else:
                error = result.get("error")
                error_description = result.get("error_description")
                print(f"Error acquiring token: {error}, {error_description}")
                raise Exception("Failed to acquire token")

        except requests.RequestException as e:
            print(f"Request failed: {e}")
            raise
        except Exception as e:
            print(f"Error: {e}")
            raise


@task
def send_email() -> None:
    """
    Sends an test email using the Microsoft Graph API.

    Raises:
        Exception: If sending the email fails.
    """
    try:
        if REFRESH_TOKEN:
            tokens = refresh_microsoft_token(REFRESH_TOKEN)
            new_access_token = tokens["access_token"]
        else:
            raise Exception(
                "No refresh token available. Please perform the initial OAuth flow."
            )

        headers = build_headers(new_access_token)
        email = {
            "message": {
                "subject": SUBJECT,
                "body": {"contentType": "HTML", "content": BODY},
                "toRecipients": [
                    {"emailAddress": {"address": email}} for email in RECIPIENTS
                ],
            }
        }

        response = requests.post(
            f"{BASE_GRAPH_URL}/me/sendMail", headers=headers, json=email
        )

        if response.status_code not in [200, 202]:
            raise Exception(f"Failed to send email: {response.text}")

        print("Email sent successfully.")

    except requests.RequestException as e:
        print(f"Request failed: {e}")
    except Exception as e:
        print(f"Error: {e}")
