import base64
import mimetypes

from microsoftgraph.decorators import token_required
from microsoftgraph.response import Response


class Mail(object):
    def __init__(self, client):
        self._client = client

    @token_required
    def get_message(self, message_id: str, params: dict = None) -> Response:
        """Retrieve the properties and relationships of a message object.

        https://docs.microsoft.com/en-us/graph/api/message-get?view=graph-rest-1.0&tabs=http

        Args:
            message_id (str): Message ID.
            params (dict, optional): Query. Defaults to None.

        Returns:
            dict: Message data.
        """
        return self._client._get(self._client.base_url + "me/messages/" + message_id, params=params)

    @token_required
    def send_mail(
        self,
        subject: str = None,
        recipients: list = None,
        body: str = "",
        content_type: str = "HTML",
        attachments: list = None,
    ) -> Response:
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
            raise ValueError("sendmail(): required arguments missing")

        # Create recipient list in required format.
        recipient_list = [{"EmailAddress": {"Address": address}} for address in recipients]

        # Create list of attachments in required format.
        attached_files = []
        if attachments:
            for filename in attachments:
                b64_content = base64.b64encode(open(filename, "rb").read())
                mime_type = mimetypes.guess_type(filename)[0]
                mime_type = mime_type if mime_type else ""
                attached_files.append(
                    {
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "ContentBytes": b64_content.decode("utf-8"),
                        "ContentType": mime_type,
                        "Name": filename,
                    }
                )

        # Create email message in required format.
        email_msg = {
            "Message": {
                "Subject": subject,
                "Body": {"ContentType": content_type, "Content": body},
                "ToRecipients": recipient_list,
                "Attachments": attached_files,
            },
            "SaveToSentItems": "true",
        }

        # Do a POST to Graph's sendMail API and return the response.
        return self._client._post(self._client.base_url + "me/microsoft.graph.sendMail", json=email_msg)
