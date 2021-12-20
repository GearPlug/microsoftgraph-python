from urllib.parse import urlencode

import requests

from microsoftgraph import exceptions
from microsoftgraph.calendar import Calendar
from microsoftgraph.contacts import Contacts
from microsoftgraph.files import Files
from microsoftgraph.mail import Mail
from microsoftgraph.notes import Notes
from microsoftgraph.response import Response
from microsoftgraph.users import Users
from microsoftgraph.webhooks import Webhooks
from microsoftgraph.workbooks import Workbooks


class Client(object):
    AUTHORITY_URL = "https://login.microsoftonline.com/"
    AUTH_ENDPOINT = "/oauth2/v2.0/authorize?"
    TOKEN_ENDPOINT = "/oauth2/v2.0/token"
    RESOURCE = "https://graph.microsoft.com/"

    def __init__(
        self,
        client_id: str,
        client_secret: str,
        api_version: str = "v1.0",
        account_type: str = "common",
        requests_hooks: dict = None,
        paginate: bool = True,
    ) -> None:
        """Instantiates library.

        Args:
            client_id (str): Application client id.
            client_secret (str): Application client secret.
            api_version (str, optional): v1.0 or beta. Defaults to "v1.0".
            account_type (str, optional): common, organizations or consumers. Defaults to "common".
            requests_hooks (dict, optional): Requests library event hooks. Defaults to None.

        Raises:
            Exception: requests_hooks is not a dict.
        """
        self.client_id = client_id
        self.client_secret = client_secret
        self.api_version = api_version
        self.account_type = account_type

        self.base_url = self.RESOURCE + self.api_version + "/"
        self.token = None
        self.workbook_session_id = None
        self.paginate = paginate

        self.calendar = Calendar(self)
        self.contacts = Contacts(self)
        self.files = Files(self)
        self.mail = Mail(self)
        self.notes = Notes(self)
        self.users = Users(self)
        self.webhooks = Webhooks(self)
        self.workbooks = Workbooks(self)

        if requests_hooks and not isinstance(requests_hooks, dict):
            raise Exception(
                'requests_hooks must be a dict. e.g. {"response": func}. http://docs.python-requests.org/en/master/user/advanced/#event-hooks'
            )
        self.requests_hooks = requests_hooks

    def authorization_url(self, redirect_uri: str, scope: list, state: str = None) -> str:
        """Generates an Authorization URL.

        The first step to getting an access token for many OpenID Connect (OIDC) and OAuth 2.0 flows is to redirect the
        user to the Microsoft identity platform /authorize endpoint. Azure AD will sign the user in and ensure their
        consent for the permissions your app requests. In the authorization code grant flow, after consent is obtained,
        Azure AD will return an authorization_code to your app that it can redeem at the Microsoft identity platform
        /token endpoint for an access token.

        https://docs.microsoft.com/en-us/graph/auth-v2-user#2-get-authorization

        Args:
            redirect_uri (str): The redirect_uri of your app, where authentication responses can be sent and received by
            your app. It must exactly match one of the redirect_uris you registered in the app registration portal.
            scope (list): A list of the Microsoft Graph permissions that you want the user to consent to. This may also
            include OpenID scopes.
            state (str, optional): A value included in the request that will also be returned in the token response.
            It can be a string of any content that you wish.  A randomly generated unique value is typically
            used for preventing cross-site request forgery attacks.  The state is also used to encode information
            about the user's state in the app before the authentication request occurred, such as the page or view
            they were on. Defaults to None.

        Returns:
            str: Url for OAuth 2.0.
        """
        params = {
            "client_id": self.client_id,
            "redirect_uri": redirect_uri,
            "scope": " ".join(scope),
            "response_type": "code",
            "response_mode": "query",
        }

        if state:
            params["state"] = state
        response = self.AUTHORITY_URL + self.account_type + self.AUTH_ENDPOINT + urlencode(params)
        return response

    def exchange_code(self, redirect_uri: str, code: str) -> Response:
        """Exchanges an oauth code for an user token.

        Your app uses the authorization code received in the previous step to request an access token by sending a POST
        request to the /token endpoint.

        https://docs.microsoft.com/en-us/graph/auth-v2-user#3-get-a-token

        Args:
            redirect_uri (str): The redirect_uri of your app, where authentication responses can be sent and received by
            your app.  It must exactly match one of the redirect_uris you registered in the app registration portal.
            code (str): The authorization_code that you acquired in the first leg of the flow.

        Returns:
            Response: Microsoft Graph Response.
        """
        data = {
            "client_id": self.client_id,
            "redirect_uri": redirect_uri,
            "client_secret": self.client_secret,
            "code": code,
            "grant_type": "authorization_code",
        }
        response = requests.post(self.AUTHORITY_URL + self.account_type + self.TOKEN_ENDPOINT, data=data)
        return self._parse(response)

    def refresh_token(self, redirect_uri: str, refresh_token: str) -> Response:
        """Exchanges a refresh token for an user token.

        Access tokens are short lived, and you must refresh them after they expire to continue accessing resources.
        You can do so by submitting another POST request to the /token endpoint, this time providing the refresh_token
        instead of the code.

        https://docs.microsoft.com/en-us/graph/auth-v2-user#5-use-the-refresh-token-to-get-a-new-access-token

        Args:
            redirect_uri (str): The redirect_uri of your app, where authentication responses can be sent and received by
            your app.  It must exactly match one of the redirect_uris you registered in the app registration portal.
            refresh_token (str): An OAuth 2.0 refresh token. Your app can use this token acquire additional access tokens
            after the current access token expires. Refresh tokens are long-lived, and can be used to retain access
            to resources for extended periods of time.

        Returns:
            Response: Microsoft Graph Response.
        """
        data = {
            "client_id": self.client_id,
            "redirect_uri": redirect_uri,
            "client_secret": self.client_secret,
            "refresh_token": refresh_token,
            "grant_type": "refresh_token",
        }
        response = requests.post(self.AUTHORITY_URL + self.account_type + self.TOKEN_ENDPOINT, data=data)
        return self._parse(response)

    def set_token(self, token: dict) -> None:
        """Sets the User token for its use in this library.

        Args:
            token (dict): User token data.
        """
        self.token = token

    def set_workbook_session_id(self, workbook_session_id: dict) -> None:
        """Sets the Workbook Session Id token for its use in this library.

        Args:
            token (dict): Workbook Session ID.
        """
        self.workbook_session_id = workbook_session_id

    def _paginate_response(self, response: dict, **kwargs) -> dict:
        """Some queries against Microsoft Graph return multiple pages of data either due to server-side paging or due to
        the use of the $top query parameter to specifically limit the page size in a request. When a result set spans
        multiple pages, Microsoft Graph returns an @odata.nextLink property in the response that contains a URL to the
        next page of results.

        https://docs.microsoft.com/en-us/graph/paging?context=graph%2Fapi%2F1.0&view=graph-rest-1.0

        Args:
            response (dict): Graph API Response.

        Returns:
            dict: Graph API Response.
        """
        if not self.paginate or not isinstance(response.data, dict):
            return response
        while "@odata.nextLink" in response.data:
            data = response.data["value"]
            response = self._get(response.data["@odata.nextLink"], **kwargs)
            response.data["value"] += data
        return response

    def _get(self, url, **kwargs):
        return self._paginate_response(self._request("GET", url, **kwargs), **kwargs)

    def _post(self, url, **kwargs):
        return self._request("POST", url, **kwargs)

    def _put(self, url, **kwargs):
        return self._request("PUT", url, **kwargs)

    def _patch(self, url, **kwargs):
        return self._request("PATCH", url, **kwargs)

    def _delete(self, url, **kwargs):
        return self._request("DELETE", url, **kwargs)

    def _request(self, method, url, headers=None, **kwargs):
        _headers = {
            "Accept": "application/json",
        }
        _headers["Authorization"] = "Bearer " + self.token["access_token"]
        if headers:
            _headers.update(headers)
        if self.requests_hooks:
            kwargs.update({"hooks": self.requests_hooks})
        if "files" not in kwargs:
            # If you use the 'files' keyword, the library will set the Content-Type to multipart/form-data
            # and will generate a boundary.
            _headers["Content-Type"] = "application/json"
        return self._parse(requests.request(method, url, headers=_headers, **kwargs))

    def _parse(self, response):
        status_code = response.status_code
        r = Response(original=response)
        if status_code in (200, 201, 202, 204, 206):
            return r
        elif status_code == 400:
            raise exceptions.BadRequest(r.data)
        elif status_code == 401:
            raise exceptions.Unauthorized(r.data)
        elif status_code == 403:
            raise exceptions.Forbidden(r.data)
        elif status_code == 404:
            raise exceptions.NotFound(r.data)
        elif status_code == 405:
            raise exceptions.MethodNotAllowed(r.data)
        elif status_code == 406:
            raise exceptions.NotAcceptable(r.data)
        elif status_code == 409:
            raise exceptions.Conflict(r.data)
        elif status_code == 410:
            raise exceptions.Gone(r.data)
        elif status_code == 411:
            raise exceptions.LengthRequired(r.data)
        elif status_code == 412:
            raise exceptions.PreconditionFailed(r.data)
        elif status_code == 413:
            raise exceptions.RequestEntityTooLarge(r.data)
        elif status_code == 415:
            raise exceptions.UnsupportedMediaType(r.data)
        elif status_code == 416:
            raise exceptions.RequestedRangeNotSatisfiable(r.data)
        elif status_code == 422:
            raise exceptions.UnprocessableEntity(r.data)
        elif status_code == 429:
            raise exceptions.TooManyRequests(r.data)
        elif status_code == 500:
            raise exceptions.InternalServerError(r.data)
        elif status_code == 501:
            raise exceptions.NotImplemented(r.data)
        elif status_code == 503:
            raise exceptions.ServiceUnavailable(r.data)
        elif status_code == 504:
            raise exceptions.GatewayTimeout(r.data)
        elif status_code == 507:
            raise exceptions.InsufficientStorage(r.data)
        elif status_code == 509:
            raise exceptions.BandwidthLimitExceeded(r.data)
        else:
            if r["error"]["innerError"]["code"] == "lockMismatch":
                # File is currently locked due to being open in the web browser
                # while attempting to reupload a new version to the drive.
                # Thus temporarily unavailable.
                raise exceptions.ServiceUnavailable(r.data)
            raise exceptions.UnknownError(r.data)
