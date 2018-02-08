import base64
import mimetypes
import requests
from microsoftgraph import exceptions
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

        return self.AUTHORITY_URL + self.account_type + self.AUTH_ENDPOINT + urlencode(params)

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
        return requests.post(self.AUTHORITY_URL + self.account_type + self.TOKEN_ENDPOINT, data=data).json()

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
        data = {
            'client_id': self.client_id,
            'redirect_uri': redirect_uri,
            'client_secret': self.client_secret,
            'refresh_token': refresh_token,
            'grant_type': 'refresh_token',
        }
        return requests.post(self.AUTHORITY_URL + self.account_type + self.TOKEN_ENDPOINT, data=data).json()

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
        return self._get(self.base_url + 'me', params=params)

    def get_message(self, message_id, params=None):
        """Retrieve the properties and relationships of a message object.

        Args:
            message_id: A dict.

        Returns:


        """
        return self._get(self.base_url + 'me/messages/' + message_id, params=params)

    def create_subscription(self, change_type, notification_url, resource, expiration_datetime, client_state=None):
        """Creating a subscription is the first step to start receiving notifications for a resource.

        Args:
            change_type: The event type that caused the notification. For example, created on mail receive, or updated on marking a message read.
            notification_url:
            resource: The URI of the resource relative to https://graph.microsoft.com.
            expiration_datetime: The expiration time for the subscription.
            client_state: The clientState property specified in the subscription request.

        Returns:

        """
        data = {
            'changeType': change_type,
            'notificationUrl': notification_url,
            'resource': resource,
            'expirationDateTime': expiration_datetime,
            'clientState': client_state
        }
        return self._post('https://graph.microsoft.com/beta/' + 'subscriptions', json=data)

    def renew_subscription(self, subscription_id, expiration_datetime):
        """The client can renew a subscription with a specific expiration date of up to three days from the time
        of request. The expirationDateTime property is required.


        Args:
            subscription_id:

        Returns:

        """
        data = {
            'expirationDateTime': expiration_datetime
        }
        return self._patch('https://graph.microsoft.com/beta/' + 'subscriptions/{}'.format(subscription_id), json=data)

    def delete_subscription(self, subscription_id):
        """The client can stop receiving notifications by deleting the subscription using its ID.

        Args:
            subscription_id:

        Returns:

        """
        return self._delete('https://graph.microsoft.com/beta/' + 'subscriptions/{}'.format(subscription_id))

    def get_me_events(self):
        """
        Obtiene los eventos del usuario
        :return: dictionary of events.
        """
        try:
            response = self._get('me/events')
            return response
        except Exception as e:
            return False

    def create_calendar_event(
            self, subject, content,
            start_datetime, start_timezone, end_datetime,
            end_timezone, recurrence_type, recurrence_interval,
            recurrence_days_of_week, recurrence_range_type,
            recurrence_range_startdate, recurrence_range_enddate,
            location, attendees, calendar=None):
        """
        TODO: manual testing
        Create a new calendar event.
        Args:
            subject: subject of event, string
            content: content of event, string
            start_datetime: in the format of 2017-09-04T11:00:00, dateTimeTimeZone string
            start_timezone: in the format of Pacific Standard Time, string
            end_datetime: in the format of 2017-09-04T11:00:00, dateTimeTimeZone string
            end_timezone: in the format of Pacific Standard Time, string
            recurrence_type: daily, weekly, absoluteMonthly, relativeMonthly, absoluteYearly, relativeYearly
            recurrence_interval: The number of units between occurrences, can be in days, weeks, months, or years,
                                depending on the type. Required.
            recurrence_days_of_week: sunday, monday, tuesday, wednesday, thursday, friday, saturday
            recurrence_range_type: endDate, noEnd, numbered
            recurrence_range_startdate: The date to start applying the recurrence pattern. The first occurrence of the
                                        meeting may be this date or later, depending on the recurrence pattern of the
                                        event. Must be the same value as the start property of the recurring event.
                                        Required.
            recurrence_range_enddate:   Required if type is endDate, The date to stop applying the recurrence pattern.
                                        Depending on the recurrence pattern of the event, the last occurrence of the
                                        meeting may not be this date.
            location:   string
            attendees: list of dicts of the form:
                        {"emailAddress": {"address": a['attendees_email'],"name": a['attendees_name']}

        Returns:

        """
        attendees_list = [{
            "emailAddress": {
                "address": a['attendees_email'],
                "name": a['attendees_name']
            },
            "type": a['attendees_type']
        } for a in attendees]
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
            "recurrence": {
                "pattern": {
                    "type": recurrence_type,
                    "interval": recurrence_interval,
                    "daysOfWeek": recurrence_days_of_week
                },
                "range": {
                    "type": recurrence_range_type,
                    "startDate": recurrence_range_startdate,
                    "endDate": recurrence_range_enddate
                }
            },
            "location": {
                "displayName": location
            },
            "attendees": attendees_list
        }
        url = 'me/calendars/{}/events'.format(calendar) if calendar is not None else 'me/events'
        try:
            response = self._post(url, json=body)
            return response
        except Exception as e:
            return False

    def create_calendar(self, name):
        """
        Created a new calendar.
        Args:
            name: name of new calendar to be created, string.

        Returns:

        """
        body = {
            'name': '{}'.format(name)
        }
        try:
            response = self._post('me/calendars', json=body)
            return response
        except Exception as e:
            return False

    def get_me_calendar(self, id_cal=None):
        """
        TODO: manual test.
        Specific calendar.
        :return:
        """
        url = 'me/calendar/{}'.format(id_cal) if id_cal is not None else 'me/calendar'
        try:
            response = self._get(url)
            return response
        except Exception as e:
            return False

    def get_me_calendars(self):
        """
        All the calendars of user.
        :return:
        """
        try:
            response = self._get('me/calendars')
            return response
        except Exception as e:
            return False

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
            'Authorization': 'Bearer ' + self.token,
            'Accept': 'application/json',
            'Content-Type': 'application/json'
        }
        if headers:
            _headers.update(headers)
        return self._parse(requests.request(method, url, headers=_headers, **kwargs))

    def _parse(self, response):
        status_code = response.status_code
        if 'application/json' in response.headers['Content-Type']:
            r = response.json()
        else:
            r = response.text
        if status_code == 200 or status_code == 201:
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
            raise exceptions.UnknownError(r)
