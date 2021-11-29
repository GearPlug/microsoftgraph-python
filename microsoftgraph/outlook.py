from microsoftgraph.decorators import token_required
from microsoftgraph.response import Response


class Outlook(object):
    def __init__(self, client):
        self._client = client

    @token_required
    def outlook_get_me_contacts(self, data_id: str = None, params: dict = None) -> Response:
        if data_id is None:
            url = "{0}me/contacts".format(self._client.base_url)
        else:
            url = "{0}me/contacts/{1}".format(self._client.base_url, data_id)
        return self._client._get(url, params=params)

    @token_required
    def outlook_create_me_contact(self, **kwargs) -> Response:
        url = "{0}me/contacts".format(self._client.base_url)
        return self._client._post(url, **kwargs)

    @token_required
    def outlook_create_contact_in_folder(self, folder_id: str, **kwargs) -> Response:
        url = "{0}/me/contactFolders/{1}/contacts".format(self._client.base_url, folder_id)
        return self._client._post(url, **kwargs)

    @token_required
    def outlook_get_contact_folders(self, params: dict = None) -> Response:
        url = "{0}me/contactFolders".format(self._client.base_url)
        return self._client._get(url, params=params)

    @token_required
    def outlook_create_contact_folder(self, **kwargs) -> Response:
        url = "{0}me/contactFolders".format(self._client.base_url)
        return self._client._post(url, **kwargs)
