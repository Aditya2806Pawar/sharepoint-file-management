

from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

USERNAME = "your_email@domain.com"
PASSWORD = "your_password"
SITE_URL = "https://yourdomain.sharepoint.com/personal/your_email_domain_com"

CTX = ClientContext(SITE_URL).with_credentials(UserCredential(USERNAME, PASSWORD))
