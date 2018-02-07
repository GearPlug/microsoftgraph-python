import base64
import mimetypes
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

    def authorization_url(self, redirect_uri, scope, state=None):
        """

        Args:
            redirect_uri: The redirect_uri of your app, where authentication responses can be sent and received by
            your app.  It must exactly match one of the redirect_uris you registered in the app registration portal

            scope: A list of the Microsoft Graph permissions that you want the user to consent to. This may also
            include OpenID scopes.

            state: A value included in the request that will also be returned in the token response.
            It can be a string of any content that you wish.  A randomly generated unique value is typically
            used for preventing cross-site request forgery attacks.  The state is also used to encode information
            about the user's state in the app before the authentication request occurred, such as the page or view
            they were on.

        Returns:

        """
        params = {
            'client_id': self.client_id,
            'redirect_uri': redirect_uri,
            'scope': ' '.join(scope),
            'response_type': 'code',
            'response_mode': 'query'
        }

        if state:
            params['state'] = None

        url = self.AUTHORITY_URL + self.account_type + self.AUTH_ENDPOINT + urlencode(params)
        return url

    def exchange_code(self, redirect_uri, code):
        """Exchanges a code for a Token.

        Args:
            redirect_uri: The redirect_uri of your app, where authentication responses can be sent and received by
            your app.  It must exactly match one of the redirect_uris you registered in the app registration portal

            code: The authorization_code that you acquired in the first leg of the flow.

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

    def refresh_token(self, redirect_uri, refresh_token):
        """

        Args:
            redirect_uri: The redirect_uri of your app, where authentication responses can be sent and received by
            your app.  It must exactly match one of the redirect_uris you registered in the app registration portal

            refresh_token: An OAuth 2.0 refresh token. Your app can use this token acquire additional access tokens
            after the current access token expires. Refresh tokens are long-lived, and can be used to retain access
            to resources for extended periods of time.

        Returns:

        """
        params = {
            'client_id': self.client_id,
            'redirect_uri': redirect_uri,
            'client_secret': self.client_secret,
            'refresh_token': refresh_token,
            'grant_type': 'refresh_token',
        }
        return requests.post(self.AUTHORITY_URL + self.account_type + self.TOKEN_ENDPOINT, data=params).json()

    def set_token(self, token):
        """Sets the Token for its use in this library.

        Args:
            token: A string with the Token.

        """
        self.token = token

    def get_me(self, params=None):
        """Retrieve the properties and relationships of user object.

        Note: Getting a user returns a default set of properties only (businessPhones, displayName, givenName, id,
        jobTitle, mail, mobilePhone, officeLocation, preferredLanguage, surname, userPrincipalName).
        Use $select to get the other properties and relationships for the user object.

        Args:
            params: A dict.

        Returns:


        """
        return self._get('me', params=None)

    def send_mail(self, subject=None, recipients=None, body='', content_type='HTML', attachments=None):
        """Helper to send email from current user.

        Args:
            subject: email subject (required)
            recipients: list of recipient email addresses (required)
            body: body of the message
            content_type: content type (default is 'HTML')
            attachments: list of file attachments (local filenames)

        Returns:
            Returns the response from the POST to the sendmail API.
        """

        # Verify that required arguments have been passed.
        if not all([subject, recipients]):
            raise ValueError('sendmail(): required arguments missing')

        # Create recipient list in required format.
        recipient_list = [{'EmailAddress': {'Address': address}} for address in recipients]

        # Create list of attachments in required format.
        attached_files = []
        if attachments:
            for filename in attachments:
                b64_content = base64.b64encode(open(filename, 'rb').read())
                mime_type = mimetypes.guess_type(filename)[0]
                mime_type = mime_type if mime_type else ''
                attached_files.append(
                    {'@odata.type': '#microsoft.graph.fileAttachment', 'ContentBytes': b64_content.decode('utf-8'),
                     'ContentType': mime_type, 'Name': filename})

        # Create email message in required format.
        email_msg = {'Message': {'Subject': subject,
                                 'Body': {'ContentType': content_type, 'Content': body},
                                 'ToRecipients': recipient_list,
                                 'Attachments': attached_files},
                     'SaveToSentItems': 'true'}

        # Do a POST to Graph's sendMail API and return the response.
        return self._post('me/microsoft.graph.sendMail', data=email_msg)

    def _get_headers(self):
        return {'Authorization': 'Bearer ' + self.token['access_token']}

    def _get(self, endpoint, params=None):
        response = requests.get(self.base_url + endpoint, params=params, headers=self._get_headers())
        return self._parse(response)

    def _post(self, endpoint, params=None, data=None):
        response = requests.post(self.base_url + endpoint, params=params, json=data, headers=self._get_headers())
        return self._parse(response)

    def _delete(self, endpoint, params=None):
        response = requests.delete(self.base_url + endpoint, params=params)
        return self._parse(response)

    def _parse(self, response):
        if 'application/json' in response.headers['Content-Type']:
            return response.json()
        return response.text
