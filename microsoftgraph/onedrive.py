import base64

import requests

from microsoftgraph.decorators import token_required
from microsoftgraph.response import Response


class Onedrive(object):
    def __init__(self, client):
        self._client = client

    @token_required
    def drive_root_items(self, params: dict = None) -> Response:
        return self._client._get(self._client.base_url + "me/drive/root", params=params)

    @token_required
    def drive_root_children_items(self, params: dict = None) -> Response:
        return self._client._get(self._client.base_url + "me/drive/root/children", params=params)

    @token_required
    def drive_specific_folder(self, folder_id: str, params: dict = None) -> Response:
        url = self._client.base_url + "me/drive/items/{0}/children".format(folder_id)
        return self._client._get(url, params=params)

    @token_required
    def drive_create_session(self, item_id: str, **kwargs) -> Response:
        url = self._client.base_url + "me/drive/items/{0}/workbook/createSession".format(item_id)
        return self._client._post(url, **kwargs)

    @token_required
    def drive_refresh_session(self, item_id: str, **kwargs) -> Response:
        url = self._client.base_url + "me/drive/items/{0}/workbook/refreshSession".format(item_id)
        return self._client._post(url, **kwargs)

    @token_required
    def drive_close_session(self, item_id: str, **kwargs) -> Response:
        url = self._client.base_url + "me/drive/items/{0}/workbook/closeSession".format(item_id)
        return self._client._post(url, **kwargs)

    @token_required
    def drive_download_contents(self, item_id: str, params: dict = None, **kwargs) -> Response:
        url = self._client.base_url + "me/drive/items/{0}/content".format(item_id)
        return self._client._get(url, params=params, **kwargs)

    @token_required
    def drive_download_shared_contents(self, share_id: str, params: dict = None, **kwargs) -> Response:
        base64_value = base64.b64encode(share_id.encode()).decode()
        encoded_share_url = "u!" + base64_value.rstrip("=").replace("/", "_").replace("+", "-")
        url = self._client.base_url + "shares/{0}/driveItem".format(encoded_share_url)
        drive_item = self._client._get(url)
        file_download_url = drive_item["@microsoft.graph.downloadUrl"]
        return drive_item["name"], requests.get(file_download_url).content

    @token_required
    def drive_download_large_contents(self, downloadUrl: str, offset: int, size: int) -> Response:
        headers = {"Range": f"bytes={offset}-{size + offset - 1}"}
        return self._client._get(downloadUrl, headers=headers)

    @token_required
    def drive_get_item(self, item_id: str, params: dict = None, **kwargs) -> Response:
        url = self._client.base_url + "me/drive/items/{0}".format(item_id)
        return self._client._get(url, params=params, **kwargs)

    @token_required
    def drive_upload_item(self, item_id: str, params: dict = None, **kwargs) -> Response:
        url = self._client.base_url + "me/drive/items/{0}/content".format(item_id)
        kwargs["headers"] = {"Content-Type": "text/plain"}
        return self._client._put(url, params=params, **kwargs)
