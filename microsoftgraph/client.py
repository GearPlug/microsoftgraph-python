import base64
import mimetypes
import requests
import json
from microsoftgraph import exceptions
from microsoftgraph.decorators import token_required
from urllib.parse import urlencode, urlparse, quote_plus


class Client(object):
    AUTHORITY_URL = 'https://login.microsoftonline.com/'
    AUTH_ENDPOINT = '/oauth2/v2.0/authorize?'
    TOKEN_ENDPOINT = '/oauth2/v2.0/token'
    RESOURCE = 'https://graph.microsoft.com/'

    OFFICE365_AUTHORITY_URL = 'https://login.live.com'
    OFFICE365_AUTH_ENDPOINT = '/oauth20_authorize.srf?'
    OFFICE365_TOKEN_ENDPOINT = '/oauth20_token.srf'

    def __init__(self, client_id, client_secret, api_version='v1.0', account_type='common', office365=False):
        self.client_id = client_id
        self.client_secret = client_secret
        self.api_version = api_version
        self.account_type = account_type

        self.base_url = self.RESOURCE + self.api_version + '/'
        self.token = None
        self.office365 = office365
        self.office365_token = None

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
            A string.

        """
        params = {
            'client_id': self.client_id,
            'redirect_uri': redirect_uri,
            'scope': ' '.join(scope),
            'response_type': 'code',
            'response_mode': 'query'
        }

        if state:
            params['state'] = state
        if self.office365:
            response = self.OFFICE365_AUTHORITY_URL + self.OFFICE365_AUTH_ENDPOINT + urlencode(params)
        else:
            response = self.AUTHORITY_URL + self.account_type + self.AUTH_ENDPOINT + urlencode(params)
        return response

    def exchange_code(self, redirect_uri, code):
        """Exchanges a code for a Token.

        Args:
            redirect_uri: The redirect_uri of your app, where authentication responses can be sent and received by
            your app.  It must exactly match one of the redirect_uris you registered in the app registration portal

            code: The authorization_code that you acquired in the first leg of the flow.

        Returns:
            A dict.

        """
        data = {
            'client_id': self.client_id,
            'redirect_uri': redirect_uri,
            'client_secret': self.client_secret,
            'code': code,
            'grant_type': 'authorization_code',
        }
        if self.office365:
            response = requests.post(self.OFFICE365_AUTHORITY_URL + self.OFFICE365_TOKEN_ENDPOINT, data=data)
        else:
            response = requests.post(self.AUTHORITY_URL + self.account_type + self.TOKEN_ENDPOINT, data=data)
        return self._parse(response)

    def refresh_token(self, redirect_uri, refresh_token):
        """

        Args:
            redirect_uri: The redirect_uri of your app, where authentication responses can be sent and received by
            your app.  It must exactly match one of the redirect_uris you registered in the app registration portal

            refresh_token: An OAuth 2.0 refresh token. Your app can use this token acquire additional access tokens
            after the current access token expires. Refresh tokens are long-lived, and can be used to retain access
            to resources for extended periods of time.

        Returns:
            A dict.

        """
        data = {
            'client_id': self.client_id,
            'redirect_uri': redirect_uri,
            'client_secret': self.client_secret,
            'refresh_token': refresh_token,
            'grant_type': 'refresh_token',
        }
        if self.office365:
            response = requests.post(self.OFFICE365_AUTHORITY_URL + self.OFFICE365_TOKEN_ENDPOINT, data=data)
        else:
            response = requests.post(self.AUTHORITY_URL + self.account_type + self.TOKEN_ENDPOINT, data=data)
        return self._parse(response)

    def set_token(self, token):
        """Sets the Token for its use in this library.

        Args:
            token: A string with the Token.

        """
        if self.office365:
            self.office365_token = token
        else:
            self.token = token

    @token_required
    def get_me(self, params=None):
        """Retrieve the properties and relationships of user object.

        Note: Getting a user returns a default set of properties only (businessPhones, displayName, givenName, id,
        jobTitle, mail, mobilePhone, officeLocation, preferredLanguage, surname, userPrincipalName).
        Use $select to get the other properties and relationships for the user object.

        Args:
            params: A dict.

        Returns:
            A dict.

        """
        return self._get(self.base_url + 'me', params=params)

    @token_required
    def get_message(self, message_id, params=None):
        """Retrieve the properties and relationships of a message object.

        Args:
            message_id: A dict.
            params:

        Returns:
            A dict.

        """
        return self._get(self.base_url + 'me/messages/' + message_id, params=params)

    @token_required
    def create_subscription(self, change_type, notification_url, resource, expiration_datetime, client_state=None):
        """Creating a subscription is the first step to start receiving notifications for a resource.

        Args:
            change_type: The event type that caused the notification. For example, created on mail receive, or updated
            on marking a message read.
            notification_url:
            resource: The URI of the resource relative to https://graph.microsoft.com.
            expiration_datetime: The expiration time for the subscription.
            client_state: The clientState property specified in the subscription request.

        Returns:
            A dict.

        """
        data = {
            'changeType': change_type,
            'notificationUrl': notification_url,
            'resource': resource,
            'expirationDateTime': expiration_datetime,
            'clientState': client_state
        }
        return self._post(self.base_url + 'subscriptions', json=data)

    @token_required
    def renew_subscription(self, subscription_id, expiration_datetime):
        """The client can renew a subscription with a specific expiration date of up to three days from the time
        of request. The expirationDateTime property is required.


        Args:
            subscription_id:
            expiration_datetime:

        Returns:
            A dict.

        """
        data = {
            'expirationDateTime': expiration_datetime
        }
        return self._patch(self.base_url + 'subscriptions/{}'.format(subscription_id), json=data)

    @token_required
    def delete_subscription(self, subscription_id):
        """The client can stop receiving notifications by deleting the subscription using its ID.

        Args:
            subscription_id:

        Returns:
            None.

        """
        return self._delete(self.base_url + 'subscriptions/{}'.format(subscription_id))

    # Onenote
    @token_required
    def list_notebooks(self):
        """Retrieve a list of notebook objects.

        Returns:
            A dict.

        """
        return self._get(self.base_url + 'me/onenote/notebooks')

    @token_required
    def get_notebook(self, notebook_id):
        """Retrieve the properties and relationships of a notebook object.

        Args:
            notebook_id:

        Returns:
            A dict.

        """
        return self._get(self.base_url + 'me/onenote/notebooks/' + notebook_id)

    @token_required
    def get_notebook_sections(self, notebook_id):
        """Retrieve the properties and relationships of a notebook object.

        Args:
            notebook_id:

        Returns:
            A dict.

        """
        return self._get(self.base_url + 'me/onenote/notebooks/{}/sections'.format(notebook_id))

    @token_required
    def create_page(self, section_id, files):
        """Create a new page in the specified section.

        Args:
            section_id:
            files:

        Returns:
            A dict.

        """
        return self._post(self.base_url + '/me/onenote/sections/{}/pages'.format(section_id), files=files)

    @token_required
    def list_pages(self, params=None):
        """Create a new page in the specified section.

        Args:
            params:

        Returns:
            A dict.

        """
        return self._get(self.base_url + '/me/onenote/pages', params=params)

    # Calendar
    @token_required
    def get_me_events(self):
        """Get a list of event objects in the user's mailbox. The list contains single instance meetings and
        series masters.

        Currently, this operation returns event bodies in only HTML format.

        Returns:
            A dict.

        """
        return self._get(self.base_url + 'me/events')

    @token_required
    def create_calendar_event(self, subject, content, start_datetime, start_timezone, end_datetime, end_timezone,
                              location, calendar=None, **kwargs):
        """
        Create a new calendar event.

        Args:
            subject: subject of event, string
            content: content of event, string
            start_datetime: in the format of 2017-09-04T11:00:00, dateTimeTimeZone string
            start_timezone: in the format of Pacific Standard Time, string
            end_datetime: in the format of 2017-09-04T11:00:00, dateTimeTimeZone string
            end_timezone: in the format of Pacific Standard Time, string
            location:   string
            attendees: list of dicts of the form:
                        {"emailAddress": {"address": a['attendees_email'],"name": a['attendees_name']}
            calendar:

        Returns:
            A dict.

        """
        # TODO: attendees
        # attendees_list = [{
        #     "emailAddress": {
        #         "address": a['attendees_email'],
        #         "name": a['attendees_name']
        #     },
        #     "type": a['attendees_type']
        # } for a in kwargs['attendees']]
        body = {
            "subject": subject,
            "body": {
                "contentType": "HTML",
                "content": content
            },
            "start": {
                "dateTime": start_datetime,
                "timeZone": start_timezone
            },
            "end": {
                "dateTime": end_datetime,
                "timeZone": end_timezone
            },
            "location": {
                "displayName": location
            },
            # "attendees": attendees_list
        }
        url = 'me/calendars/{}/events'.format(calendar) if calendar is not None else 'me/events'
        return self._post(self.base_url + url, json=body)

    @token_required
    def create_calendar(self, name):
        """Create an event in the user's default calendar or specified calendar.

        You can specify the time zone for each of the start and end times of the event as part of these values,
        as the  start and end properties are of dateTimeTimeZone type.

        When an event is sent, the server sends invitations to all the attendees.

        Args:
            name:

        Returns:
            A dict.

        """
        body = {
            'name': '{}'.format(name)
        }
        return self._post(self.base_url + 'me/calendars', json=body)

    @token_required
    def get_me_calendars(self):
        """Get all the user's calendars (/calendars navigation property), get the calendars from the default
        calendar group or from a specific calendar group.

        Returns:
            A dict.

        """
        return self._get(self.base_url + 'me/calendars')

    # Mail
    @token_required
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
        return self._post(self.base_url + 'me/microsoft.graph.sendMail', json=email_msg)

    # Outlook
    @token_required
    def outlook_get_me_contacts(self, data_id=None, params=None):
        if data_id is None:
            url = "{0}me/contacts".format(self.base_url)
        else:
            url = "{0}me/contacts/{1}".format(self.base_url, data_id)
        return self._get(url, params=params)

    @token_required
    def outlook_create_me_contact(self, **kwargs):
        url = "{0}me/contacts".format(self.base_url)
        return self._post(url, **kwargs)

    @token_required
    def outlook_create_contact_in_folder(self, folder_id, **kwargs):
        url = "{0}/me/contactFolders/{1}/contacts".format(self.base_url, folder_id)
        return self._post(url, **kwargs)

    @token_required
    def outlook_get_contact_folders(self, params=None):
        url = "{0}me/contactFolders".format(self.base_url)
        return self._get(url, params=params)

    @token_required
    def outlook_create_contact_folder(self, **kwargs):
        url = "{0}me/contactFolders".format(self.base_url)
        return self._post(url, **kwargs)

    # Onedrive
    @token_required
    def drive_root_items(self, params=None):
        return self._get('https://graph.microsoft.com/beta/me/drive/root', params=params)

    @token_required
    def drive_root_children_items(self, params=None):
        return self._get('https://graph.microsoft.com/beta/me/drive/root/children', params=params)

    @token_required
    def drive_specific_folder(self, folder_id, params=None):
        url = "https://graph.microsoft.com/beta/me/drive/items/{0}/children".format(folder_id)
        return self._get(url, params=params)

    @token_required
    def drive_create_session(self, item_id, **kwargs):
        url = "https://graph.microsoft.com/v1.0/me/drive/items/{0}/workbook/createSession".format(item_id)
        # url = "https://graph.microsoft.com/beta/me/drive/items/{0}/workbook/createSession".format(item_id)
        return self._post(url, **kwargs)

    @token_required
    def drive_refresh_session(self, item_id, **kwargs):
        url = "https://graph.microsoft.com/beta/me/drive/items/{0}/workbook/refreshSession".format(item_id)
        return self._post(url, **kwargs)

    @token_required
    def drive_close_session(self, item_id, **kwargs):
        url = "https://graph.microsoft.com/beta/me/drive/items/{0}/workbook/closeSession".format(item_id)
        return self._post(url, **kwargs)

    @token_required
    def drive_download_contents(self, item_id, params=None, **kwargs):
        url = "https://graph.microsoft.com/beta/me/drive/items/{0}/content".format(item_id)
        return self._get(url, params=params, **kwargs)

    @token_required
    def drive_get_item(self, item_id, params=None, **kwargs):
        url = "https://graph.microsoft.com/beta/me/drive/items/{0}".format(item_id)
        return self._get(url, params=params, **kwargs)

    @token_required
    def drive_upload_item(self, item_id, params=None, **kwargs):
        url = "https://graph.microsoft.com/beta/me/drive/items/{0}/content".format(item_id)
        kwargs['headers'] = {'Content-Type': 'text/plain'}
        return self._put(url, params=params, **kwargs)

    # Excel
    @token_required
    def excel_get_worksheets(self, item_id, params=None, **kwargs):
        url = "https://graph.microsoft.com/beta/me/drive/items/{0}/workbook/worksheets".format(item_id)
        return self._get(url, params=params, **kwargs)

    @token_required
    def excel_get_names(self, item_id, params=None, **kwargs):
        url = "https://graph.microsoft.com/beta/me/drive/items/{0}/workbook/names".format(item_id)
        return self._get(url, params=params, **kwargs)

    @token_required
    def excel_add_worksheet(self, item_id, **kwargs):
        url = "https://graph.microsoft.com/beta/me/drive/items/{0}/workbook/worksheets/add".format(item_id)
        return self._post(url, **kwargs)

    @token_required
    def excel_get_specific_worksheet(self, item_id, worksheet_id, **kwargs):
        url = "https://graph.microsoft.com/beta/me/drive/items/{0}/workbook/worksheets/{1}".format(item_id, quote_plus(worksheet_id))
        return self._get(url, **kwargs)

    @token_required
    def excel_update_worksheet(self, item_id, worksheet_id, **kwargs):
        url = "https://graph.microsoft.com/beta/me/drive/items/{0}/workbook/worksheets/{1}".format(item_id, quote_plus(worksheet_id))
        return self._patch(url, **kwargs)

    @token_required
    def excel_get_charts(self, item_id, worksheet_id, params=None, **kwargs):
        url = "https://graph.microsoft.com/beta/me/drive/items/{0}/workbook/worksheets/{1}/charts".format(item_id, quote_plus(worksheet_id))
        return self._get(url, params=params, **kwargs)

    @token_required
    def excel_add_chart(self, item_id, worksheet_id, **kwargs):
        url = "https://graph.microsoft.com/beta/me/drive/items/{0}/workbook/worksheets/{1}/charts/add".format(item_id, quote_plus(worksheet_id))
        return self._post(url, **kwargs)

    @token_required
    def excel_get_tables(self, item_id, params=None, **kwargs):
        url = "https://graph.microsoft.com/beta/me/drive/items/{0}/workbook/tables".format(item_id)
        return self._get(url, params=params, **kwargs)

    @token_required
    def excel_add_table(self, item_id, **kwargs):
        url = "https://graph.microsoft.com/beta/me/drive/items/{0}/workbook/tables/add".format(item_id)
        return self._post(url, **kwargs)

    @token_required
    def excel_add_column(self, item_id, worksheets_id, table_id, **kwargs):
        url = "https://graph.microsoft.com/beta/me/drive/items/{0}/workbook/worksheets/{1}/tables/{2}/columns".format(item_id, quote_plus(worksheets_id), table_id)
        return self._post(url, **kwargs)

    @token_required
    def excel_add_row(self, item_id, worksheets_id, table_id, **kwargs):
        url = "https://graph.microsoft.com/beta/me/drive/items/{0}/workbook/worksheets/{1}/tables/{2}/rows".format(item_id, quote_plus(worksheets_id), table_id)
        return self._post(url, **kwargs)

    @token_required
    def excel_get_rows(self, item_id, table_id, params=None, **kwargs):
        url = "https://graph.microsoft.com/beta/me/drive/items/{0}/workbook/tables/{1}/rows".format(item_id, table_id)
        return self._get(url, params=params, **kwargs)

    # @token_required
    # def excel_get_cell(self, item_id, worksheets_id, params=None, **kwargs):
    #     url = "https://graph.microsoft.com/beta/me/drive/items/{0}/workbook/worksheets/{1}/Cell(row='1', column='A')".format(item_id, quote_plus(worksheets_id))
    #     return self._get(url, params=params, **kwargs)

    # @token_required
    # def excel_add_cell(self, item_id, worksheets_id, **kwargs):
    #     url = "https://graph.microsoft.com/beta/me/drive/items/{0}/workbook/worksheets/{1}/rows".format(item_id, worksheets_id)
    #     return self._patch(url, **kwargs)

    @token_required
    def excel_get_range(self, item_id, worksheets_id, **kwargs):
        url = "https://graph.microsoft.com/beta/me/drive/items/{0}/workbook/worksheets/{1}/range(address='A1:B2')".format(item_id, quote_plus(worksheets_id))
        return self._get(url, **kwargs)

    @token_required
    def excel_update_range(self, item_id, worksheets_id, **kwargs):
        url = "https://graph.microsoft.com/beta/me/drive/items/{0}/workbook/worksheets/{1}/range(address='A1:B2')".format(item_id, quote_plus(worksheets_id))
        return self._patch(url, **kwargs)

    def _get(self, url, **kwargs):
        return self._request('GET', url, **kwargs)

    def _post(self, url, **kwargs):
        return self._request('POST', url, **kwargs)

    def _put(self, url, **kwargs):
        return self._request('PUT', url, **kwargs)

    def _patch(self, url, **kwargs):
        return self._request('PATCH', url, **kwargs)

    def _delete(self, url, **kwargs):
        return self._request('DELETE', url, **kwargs)

    def _request(self, method, url, headers=None, **kwargs):
        _headers = {
            'Accept': 'application/json',
        }
        if self.office365:
            _headers['Authorization'] = 'Bearer ' + self.office365_token['access_token']
        else:
            _headers['Authorization'] = 'Bearer ' + self.token['access_token']
        if headers:
            _headers.update(headers)
        if 'files' not in kwargs:
            # If you use the 'files' keyword, the library will set the Content-Type to multipart/form-data
            # and will generate a boundary.
            _headers['Content-Type'] = 'application/json'
        return self._parse(requests.request(method, url, headers=_headers, **kwargs))

    def _parse(self, response):
        status_code = response.status_code
        if 'application/json' in response.headers['Content-Type']:
            r = response.json()
        else:
            r = response.content
        if status_code in (200, 201, 202):
            return r
        elif status_code == 204:
            return None
        elif status_code == 400:
            raise exceptions.BadRequest(r)
        elif status_code == 401:
            raise exceptions.Unauthorized(r)
        elif status_code == 403:
            raise exceptions.Forbidden(r)
        elif status_code == 404:
            raise exceptions.NotFound(r)
        elif status_code == 405:
            raise exceptions.MethodNotAllowed(r)
        elif status_code == 406:
            raise exceptions.NotAcceptable(r)
        elif status_code == 409:
            raise exceptions.Conflict(r)
        elif status_code == 410:
            raise exceptions.Gone(r)
        elif status_code == 411:
            raise exceptions.LengthRequired(r)
        elif status_code == 412:
            raise exceptions.PreconditionFailed(r)
        elif status_code == 413:
            raise exceptions.RequestEntityTooLarge(r)
        elif status_code == 415:
            raise exceptions.UnsupportedMediaType(r)
        elif status_code == 416:
            raise exceptions.RequestedRangeNotSatisfiable(r)
        elif status_code == 422:
            raise exceptions.UnprocessableEntity(r)
        elif status_code == 429:
            raise exceptions.TooManyRequests(r)
        elif status_code == 500:
            raise exceptions.InternalServerError(r)
        elif status_code == 501:
            raise exceptions.NotImplemented(r)
        elif status_code == 503:
            raise exceptions.ServiceUnavailable(r)
        elif status_code == 504:
            raise exceptions.GatewayTimeout(r)
        elif status_code == 507:
            raise exceptions.InsufficientStorage(r)
        elif status_code == 509:
            raise exceptions.BandwidthLimitExceeded(r)
        else:
            if r['error']['innerError']['code'] == 'lockMismatch':
                # File is currently locked due to being open in the web browser
                # while attempting to reupload a new version to the drive.
                # Thus temporarily unavailable.
                raise exceptions.ServiceUnavailable(r)
            raise exceptions.UnknownError(r)
