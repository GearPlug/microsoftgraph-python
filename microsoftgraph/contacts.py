from microsoftgraph.decorators import token_required
from microsoftgraph.response import Response


class Contacts(object):
    def __init__(self, client) -> None:
        """Working with Outlook Contacts.

        https://docs.microsoft.com/en-us/graph/api/resources/contact?view=graph-rest-1.0

        Args:
            client (Client): Library Client.
        """
        self._client = client

    @token_required
    def get_contact(self, contact_id: str, params: dict = None) -> Response:
        """Retrieve the properties and relationships of a contact object.

        https://docs.microsoft.com/en-us/graph/api/contact-get?view=graph-rest-1.0&tabs=http

        Args:
            contact_id (str): The contact's unique identifier.
            params (dict, optional): Query. Defaults to None.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = "{0}me/contacts/{1}".format(self._client.base_url, contact_id)
        return self._client._get(url, params=params)

    @token_required
    def list_contacts(self, params: dict = None) -> Response:
        """Get a contact collection from the default contacts folder of the signed-in user.

        https://docs.microsoft.com/en-us/graph/api/user-list-contacts?view=graph-rest-1.0&tabs=http

        Args:
            params (dict, optional): Query. Defaults to None.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = "{0}me/contacts".format(self._client.base_url)
        return self._client._get(url, params=params)

    @token_required
    def create_contact(self, **kwargs) -> Response:
        """Add a contact to the root Contacts folder.

        https://docs.microsoft.com/en-us/graph/api/user-post-contacts?view=graph-rest-1.0&tabs=http

        Returns:
            Response: Microsoft Graph Response.
        """
        url = "{0}me/contacts".format(self._client.base_url)
        return self._client._post(url, **kwargs)

    @token_required
    def create_contact_in_folder(self, folder_id: str, **kwargs) -> Response:
        """Add a contact to another contact folder.

        https://docs.microsoft.com/en-us/graph/api/user-post-contacts?view=graph-rest-1.0&tabs=http

        Args:
            folder_id (str): Unique identifier of the contact folder.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = "{0}me/contactFolders/{1}/contacts".format(self._client.base_url, folder_id)
        return self._client._post(url, **kwargs)

    @token_required
    def list_contact_folders(self, params: dict = None) -> Response:
        """Get the contact folder collection in the default Contacts folder of the signed-in user.

        https://docs.microsoft.com/en-us/graph/api/user-list-contactfolders?view=graph-rest-1.0&tabs=http

        Args:
            params (dict, optional): Query. Defaults to None.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = "{0}me/contactFolders".format(self._client.base_url)
        return self._client._get(url, params=params)

    @token_required
    def create_contact_folder(self, **kwargs) -> Response:
        """Create a new contactFolder under the user's default contacts folder.

        https://docs.microsoft.com/en-us/graph/api/user-post-contactfolders?view=graph-rest-1.0&tabs=http

        Returns:
            Response: Microsoft Graph Response.
        """
        url = "{0}me/contactFolders".format(self._client.base_url)
        return self._client._post(url, **kwargs)
