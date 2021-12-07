import base64
import mimetypes

from microsoftgraph.decorators import token_required
from microsoftgraph.response import Response


class Mail(object):
    def __init__(self, client) -> None:
        """Use the Outlook mail REST API

        https://docs.microsoft.com/en-us/graph/api/resources/mail-api-overview?view=graph-rest-1.0

        Args:
            client (Client): Library Client.
        """
        self._client = client

    @token_required
    def get_message(self, message_id: str, params: dict = None) -> Response:
        """Retrieve the properties and relationships of a message object.

        https://docs.microsoft.com/en-us/graph/api/message-get?view=graph-rest-1.0&tabs=http

        Args:
            message_id (str): Unique identifier for the message.
            params (dict, optional): Query. Defaults to None.

        Returns:
            Response: Microsoft Graph Response.
        """
        return self._client._get(self._client.base_url + "me/messages/" + message_id, params=params)

    @token_required
    def send_mail(
        self,
        subject: str,
        content: str,
        to_recipients: list,
        cc_recipients: list = None,
        content_type: str = "HTML",
        attachments: list = None,
        save_to_sent_items: bool = True,
        **kwargs,
    ) -> Response:
        """Send the message specified in the request body using either JSON or MIME format.

        https://docs.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0&tabs=http

        Args:
            subject (str): The subject of the message.
            content (str): The body of the message.
            to_recipients (list, optional): The To: recipients for the message.
            cc_recipients (list, optional): The Cc: recipients for the message. Defaults to None.
            content_type (str, optional): It can be in HTML or text format. Defaults to "HTML".
            attachments (list, optional): The fileAttachment and itemAttachment attachments for the message. Defaults to None.
            save_to_sent_items (bool, optional): Indicates whether to save the message in Sent Items. Defaults to True.

        Returns:
            Response: Microsoft Graph Response.
        """
        # Create recipient list in required format.
        if isinstance(to_recipients, list):
            if all([isinstance(e, str) for e in to_recipients]):
                to_recipients = [{"EmailAddress": {"Address": address}} for address in to_recipients]
        elif isinstance(to_recipients, str):
            to_recipients = [{"EmailAddress": {"Address": to_recipients}}]
        else:
            raise Exception("to_recipients value is invalid.")

        if cc_recipients and isinstance(cc_recipients, list):
            if all([isinstance(e, str) for e in cc_recipients]):
                cc_recipients = [{"EmailAddress": {"Address": address}} for address in cc_recipients]
        elif cc_recipients and isinstance(cc_recipients, str):
            cc_recipients = [{"EmailAddress": {"Address": cc_recipients}}]
        else:
            cc_recipients = []

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
                "Body": {"ContentType": content_type, "Content": content},
                "ToRecipients": to_recipients,
                "ccRecipients": cc_recipients,
                "Attachments": attached_files,
            },
            "SaveToSentItems": save_to_sent_items,
        }
        email_msg.update(kwargs)

        # Do a POST to Graph's sendMail API and return the response.
        return self._client._post(self._client.base_url + "me/sendMail", json=email_msg)
