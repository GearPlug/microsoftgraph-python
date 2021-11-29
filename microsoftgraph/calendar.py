from microsoftgraph.decorators import token_required
from microsoftgraph.response import Response


class Calendar(object):
    def __init__(self, client):
        self._client = client

    @token_required
    def get_me_events(self) -> Response:
        """Get a list of event objects in the user's mailbox. The list contains single instance meetings and
        series masters.

        Currently, this operation returns event bodies in only HTML format.

        Returns:
            A dict.

        """
        return self._client._get(self._client.base_url + "me/events")

    @token_required
    def create_calendar_event(
        self,
        subject: str,
        content: str,
        start_datetime: str,
        start_timezone: str,
        end_datetime: str,
        end_timezone: str,
        location: str,
        calendar: str = None,
        **kwargs,
    ) -> Response:
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
            "body": {"contentType": "HTML", "content": content},
            "start": {"dateTime": start_datetime, "timeZone": start_timezone},
            "end": {"dateTime": end_datetime, "timeZone": end_timezone},
            "location": {"displayName": location},
            # "attendees": attendees_list
        }
        url = "me/calendars/{}/events".format(calendar) if calendar is not None else "me/events"
        return self._client._post(self._client.base_url + url, json=body)

    @token_required
    def create_calendar(self, name: str) -> Response:
        """Create an event in the user's default calendar or specified calendar.

        You can specify the time zone for each of the start and end times of the event as part of these values,
        as the  start and end properties are of dateTimeTimeZone type.

        When an event is sent, the server sends invitations to all the attendees.

        Args:
            name:

        Returns:
            A dict.

        """
        body = {"name": "{}".format(name)}
        return self._client._post(self._client.base_url + "me/calendars", json=body)

    @token_required
    def get_me_calendars(self) -> Response:
        """Get all the user's calendars (/calendars navigation property), get the calendars from the default
        calendar group or from a specific calendar group.

        Returns:
            A dict.

        """
        return self._client._get(self._client.base_url + "me/calendars")
