import requests
from urllib.parse import urlencode, urlparse


class Client(object):
    AUTHORITY_URL = 'https://login.microsoftonline.com/'
    AUTH_ENDPOINT = '/oauth2/v2.0/authorize?'
    TOKEN_ENDPOINT = '/oauth2/v2.0/token'

    RESOURCE = 'https://graph.microsoft.com/'

    def __init__(self, client_id, client_secret, api_version='v1.0', account_type='common'):
        self.client_id = client_id
        self.client_secret = client_secret
        self.api_version = api_version
        self.account_type = account_type

        self.base_url = self.RESOURCE + self.api_version + '/'
        self.token = None

    def authorization_url(self, redirect_uri, scope):
        """

        Args:
            redirect_uri:
            scope:

        Returns:

        """
        params = {
            'client_id': self.client_id,
            'redirect_uri': redirect_uri,
            'scope': ' '.join(scope),
            'response_type': 'code'
        }
        url = self.AUTHORITY_URL + self.account_type + self.AUTH_ENDPOINT + urlencode(params)
        return url

    def exchange_code(self, redirect_uri, code, scope):
        """Exchanges a code for a Token.

        Args:
            redirect_uri: A string with the redirect_uri set in the app config.
            code: A string containing the code to exchange.

        Returns:
            A dict.

        """
        params = {
            'client_id': self.client_id,
            'redirect_uri': redirect_uri,
            'client_secret': self.client_secret,
            'code': code,
            'grant_type': 'authorization_code',
        }
        return requests.post(self.AUTHORITY_URL + self.account_type + self.TOKEN_ENDPOINT, data=params).json()

    def set_token(self, token):
        """Sets the Access Token for its use in this library.

        Args:
            token: A string with the Access Token.

        """
        self.token = token

    def get_me(self):
        return self._get('me')

    def _get(self, endpoint, params=None):
        headers = {'Authorization': 'Bearer ' + self.token['access_token']}
        response = requests.get(self.base_url + endpoint, params=params, headers=headers)
        return self._parse(response)

    def _post(self, endpoint, params=None, data=None):
        response = requests.post(self.base_url + endpoint, params=params, data=data)
        return self._parse(response)

    def _delete(self, endpoint, params=None):
        response = requests.delete(self.base_url + endpoint, params=params)
        return self._parse(response)

    def _parse(self, response):
        return response.json()

