from microsoftgraph.decorators import token_required
from microsoftgraph.response import Response


class Webhooks(object):
    def __init__(self, client):
        self._client = client

    @token_required
    def create_subscription(
        self,
        change_type: str,
        notification_url: str,
        resource: str,
        expiration_datetime: str,
        client_state: str = None,
    ) -> Response:
        """Creates a subscription to start receiving notifications for a resource.

        https://docs.microsoft.com/en-us/graph/webhooks#creating-a-subscription

        Args:
            change_type (str): The event type that caused the notification. For example, created on mail receive, or
            updated on marking a message read.
            notification_url (str): Url to receive notifications.
            resource (str): The URI of the resource relative to https://graph.microsoft.com.
            expiration_datetime (str): The expiration time for the subscription.
            client_state (str, optional): The clientState property specified in the subscription request. Defaults to None.

        Returns:
            Response: Microsoft Graph Response.
        """
        data = {
            "changeType": change_type,
            "notificationUrl": notification_url,
            "resource": resource,
            "expirationDateTime": expiration_datetime,
            "clientState": client_state,
        }
        return self._client._post(self._client.base_url + "subscriptions", json=data)

    @token_required
    def renew_subscription(self, subscription_id: str, expiration_datetime: str) -> Response:
        """Renews a subscription to keep receiving notifications for a resource.

        https://docs.microsoft.com/en-us/graph/webhooks#renewing-a-subscription

        Args:
            subscription_id (str): Subscription ID.
            expiration_datetime (str): Expiration date.

        Returns:
            Response: Microsoft Graph Response.
        """
        data = {"expirationDateTime": expiration_datetime}
        return self._client._patch(self._client.base_url + "subscriptions/{}".format(subscription_id), json=data)

    @token_required
    def delete_subscription(self, subscription_id: str) -> Response:
        """Deletes a subscription to stop receiving notifications for a resource.

        https://docs.microsoft.com/en-us/graph/webhooks#deleting-a-subscription

        Args:
            subscription_id (str): Response ID.

        Returns:
            Response: Microsoft Graph Response.
        """
        return self._client._delete(self._client.base_url + "subscriptions/{}".format(subscription_id))
