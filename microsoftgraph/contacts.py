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
        return self._client._get(self._client.base_url + "me/contacts/{}".format(contact_id), params=params)

    @token_required
    def list_contacts(self, folder_id: str = None, params: dict = None) -> Response:
        """Get a contact collection from the default contacts folder of the signed-in user.

        https://docs.microsoft.com/en-us/graph/api/user-list-contacts?view=graph-rest-1.0&tabs=http

        Args:
            folder_id (str): Folder ID.
            params (dict, optional): Query. Defaults to None.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = "me/contactfolders/{}/contacts".format(folder_id) if folder_id else "me/contacts"
        return self._client._get(self._client.base_url + url, params=params)

    @token_required
    def create_contact(
        self,
        given_name: str,
        surname: str,
        email_addresses: list,
        business_phones: list,
        folder_id: str = None,
        **kwargs
    ) -> Response:
        """Add a contact to the root Contacts folder or to the contacts endpoint of another contact folder.

        https://docs.microsoft.com/en-us/graph/api/user-post-contacts?view=graph-rest-1.0&tabs=http

        Args:
            given_name (str): The contact's given name.
            surname (str): The contact's surname.
            email_addresses (list): The contact's email addresses.
            business_phones (list): The contact's business phone numbers.
            folder_id (str, optional): Unique identifier of the contact folder. Defaults to None.

        Returns:
            Response: Microsoft Graph Response.
        """
        if isinstance(email_addresses, str):
            email_addresses = [{"address": email_addresses, "name": "{} {}".format(given_name, surname)}]

        if isinstance(business_phones, str):
            business_phones = [business_phones]

        body = {
            "givenName": given_name,
            "surname": surname,
            "emailAddresses": email_addresses,
            "businessPhones": business_phones,
        }
        body.update(kwargs)
        url = "me/contactFolders/{}/contacts".format(folder_id) if folder_id else "me/contacts"
        return self._client._post(self._client.base_url + url, json=body)

    @token_required
    def list_contact_folders(self, params: dict = None) -> Response:
        """Get the contact folder collection in the default Contacts folder of the signed-in user.

        https://docs.microsoft.com/en-us/graph/api/user-list-contactfolders?view=graph-rest-1.0&tabs=http

        Args:
            params (dict, optional): Query. Defaults to None.

        Returns:
            Response: Microsoft Graph Response.
        """
        return self._client._get(self._client.base_url + "me/contactFolders", params=params)

    @token_required
    def create_contact_folder(self, display_name: str, parent_folder_id: str, **kwargs) -> Response:
        """Create a new contactFolder under the user's default contacts folder.

        https://docs.microsoft.com/en-us/graph/api/user-post-contactfolders?view=graph-rest-1.0&tabs=http

        Args:
            display_name (str): The folder's display name.
            parent_folder_id (str): The ID of the folder's parent folder.

        Returns:
            Response: Microsoft Graph Response.
        """
        data = {"displayName": display_name, "parentFolderId": parent_folder_id}
        data.update(kwargs)
        return self._client._post(self._client.base_url + "me/contactFolders", json=data)
