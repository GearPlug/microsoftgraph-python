from microsoftgraph.decorators import token_required
from microsoftgraph.response import Response


class Users(object):
    def __init__(self, client) -> None:
        """Working with users in Microsoft Graph.

        https://docs.microsoft.com/en-us/graph/api/resources/users?view=graph-rest-1.0

        Args:
            client (Client): Library Client.
        """
        self._client = client

    @token_required
    def get_me(self, params: dict = None) -> Response:
        """Retrieve the properties and relationships of user object.

        Note: Getting a user returns a default set of properties only (businessPhones, displayName, givenName, id,
        jobTitle, mail, mobilePhone, officeLocation, preferredLanguage, surname, userPrincipalName).
        Use $select to get the other properties and relationships for the user object.

        https://docs.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0&tabs=http

        Args:
            params (dict, optional): Query. Defaults to None.

        Returns:
            Response: Microsoft Graph Response.
        """
        return self._client._get(self._client.base_url + "me", params=params)
