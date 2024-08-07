from urllib.parse import parse_qs, urlparse

from msal import ConfidentialClientApplication
from robocorp import browser, log, vault

from variables import SCOPES, SECRET_NAME

MAIL_SECRETS = vault.get_secret(SECRET_NAME)
SCOPES = SCOPES.split(",")
MANDATORY_KEYS = ["tenant_id", "client_id", "client_secret"]


def check_mail_secrets(secrets: dict, mandatory_keys: list) -> None:
    """
    Checks for the presence of mandatory keys in the Control Room Vault.

    Args:
        secrets (dict): The dictionary containing the secrets.
        mandatory_keys (list): A list of keys that must be present in the secrets dictionary.

    Raises:
        KeyError: If any of the mandatory keys are missing from the secrets dictionary.
    """
    missing_keys = [key for key in mandatory_keys if key not in secrets]
    if missing_keys:
        raise KeyError(
            f"Missing mandatory keys in MAIL_SECRETS (Control Room Vault): {missing_keys}"
        )


try:
    check_mail_secrets(MAIL_SECRETS, MANDATORY_KEYS)

    TENANT_ID = MAIL_SECRETS["tenant_id"]
    CLIENT_ID = MAIL_SECRETS["client_id"]
    CLIENT_SECRET = MAIL_SECRETS["client_secret"]
    ACCESS_TOKEN = MAIL_SECRETS.get("access_token")
    REFRESH_TOKEN = MAIL_SECRETS.get("refresh_token")

except KeyError as e:
    print(f"Configuration error: {e}")

app = ConfidentialClientApplication(
    CLIENT_ID,
    authority=f"https://login.microsoftonline.com/{TENANT_ID}",
    client_credential=CLIENT_SECRET,
)


@log.suppress
def get_auth_code_using_browser(auth_url: str) -> str:
    """
    Retrieves the authorization code from the given authentication URL using a browser.

    Args:
        auth_url (str): The authentication URL to navigate to.

    Returns:
        str: The authorization code retrieved from the URL.

    Raises:
        TimeoutError: If the function times out waiting for 'code=' in the URL.
        Exception: If the authorization code is not found in the URL.
    """
    browser.configure(headless=False)
    browser.goto(auth_url)
    page = browser.page()
    page.wait_for_url(auth_url)

    try:
        page.wait_for_function("window.location.href.includes('code=')", timeout=60000)
        current_url = page.evaluate("window.location.href")
    except Exception as e:
        print(f"Error waiting for 'code=' in URL: {e}")
        raise TimeoutError("Timed out waiting for 'code=' in the URL.")

    parsed_url = urlparse(current_url)
    query_params = parse_qs(parsed_url.query)
    auth_code = query_params.get("code", [None])[0]

    if not auth_code:
        raise Exception("Authorization code not found in the URL.")

    return auth_code


@log.suppress
def refresh_microsoft_token(refresh_token) -> dict:
    """
    Refreshes the Microsoft token using the provided refresh token.

    Args:
        refresh_token (str): The refresh token.

    Returns:
        dict: A dictionary containing the new access token and refresh token.

    Raises:
        Exception: If the token refresh fails.
    """
    try:
        result = app.acquire_token_by_refresh_token(refresh_token, scopes=SCOPES)

        if "access_token" in result:
            access_token = result["access_token"]
            refresh_token = result.get("refresh_token", refresh_token)
            tokens = {"access_token": access_token, "refresh_token": refresh_token}
            update_vault(access_token, refresh_token)
            return tokens
        else:
            error_msg = (
                f"Error refreshing token: {result.get('error')}, "
                f"{result.get('error_description')}"
            )
            print(error_msg)
            raise Exception("Failed to refresh token")
    except Exception as e:
        print(f"An exception occurred: {e}")
        raise


@log.suppress
def build_headers(token: str) -> dict:
    """
    Builds the authorization headers for API requests.

    Args:
        token (str): The access token.

    Returns:
        dict: A dictionary containing the authorization headers.

    Raises:
        ValueError: If the token is missing or invalid.
    """
    try:
        if not token:
            raise ValueError("Token is missing or invalid")
        return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    except Exception as e:
        print(f"An exception occurred while building headers: {e}")
        raise


@log.suppress
def update_vault(access_token: str, refresh_token: str) -> None:
    """
    Updates the vault with new access and refresh tokens.

    Args:
        access_token (str): The new access token.
        refresh_token (str): The new refresh token.

    Raises:
        Exception: If updating the vault fails.
    """
    new_values = {
        "tenant_id": TENANT_ID,
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "access_token": access_token,
        "refresh_token": refresh_token,
    }

    try:
        vault.create_secret(
            name=SECRET_NAME,
            description="MSGraph Outlook Credentials",
            exist_ok=True,
            values=new_values,
        )
    except Exception as e:
        print(f"An exception occurred while updating the vault: {e}")
        raise
