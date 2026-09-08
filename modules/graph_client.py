import msal
import requests

class GraphClient:
    def __init__(self, tenant_id, client_id, client_secret):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.authority = f"https://login.microsoftonline.com/{tenant_id}"
        self.scope = ["https://graph.microsoft.com/.default"]
        self.token = self._get_token()

    def _get_token(self):
        app = msal.ConfidentialClientApplication(
            self.client_id,
            authority=self.authority,
            client_credential=self.client_secret
        )
        result = app.acquire_token_for_client(scopes=self.scope)
        if "access_token" in result:
            return result["access_token"]
        else:
            raise Exception(f"Authentication Failed: {result.get('error_description', 'Unknown error')}")

    def get_users(self):
        """Fetches users and basic mailbox info from the tenant via Graph API."""
        headers = {"Authorization": f"Bearer {self.token}"}
        endpoint = "https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,mail"
        
        users = []
        response = requests.get(endpoint, headers=headers)
        if response.status_code == 200:
            data = response.json()
            users.extend(data.get("value", []))
            # Handle pagination if they have many users
            while "@odata.nextLink" in data:
                response = requests.get(data["@odata.nextLink"], headers=headers)
                if response.status_code == 200:
                    data = response.json()
                    users.extend(data.get("value", []))
                else:
                    break
        return users
