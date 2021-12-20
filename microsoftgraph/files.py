import base64

import requests

from microsoftgraph.decorators import token_required
from microsoftgraph.response import Response


class Files(object):
    def __init__(self, client) -> None:
        """Working with files in Microsoft Graph

        https://docs.microsoft.com/en-us/graph/api/resources/onedrive?view=graph-rest-1.0

        Args:
            client (Client): Library Client.
        """
        self._client = client

    @token_required
    def drive_root_items(self, params: dict = None) -> Response:
        """Return a collection of DriveItems in the children relationship of a DriveItem.

        https://docs.microsoft.com/en-us/graph/api/driveitem-list-children?view=graph-rest-1.0&tabs=http

        Args:
            params (dict, optional): Query. Defaults to None.

        Returns:
            Response: Microsoft Graph Response.
        """
        return self._client._get(self._client.base_url + "me/drive/root", params=params)

    @token_required
    def drive_root_children_items(self, params: dict = None) -> Response:
        """Return a collection of DriveItems in the children relationship of a DriveItem.

        https://docs.microsoft.com/en-us/graph/api/driveitem-list-children?view=graph-rest-1.0&tabs=http

        Args:
            params (dict, optional): Query. Defaults to None.

        Returns:
            Response: Microsoft Graph Response.
        """
        return self._client._get(self._client.base_url + "me/drive/root/children", params=params)

    @token_required
    def drive_specific_folder(self, folder_id: str, params: dict = None) -> Response:
        """Return a collection of DriveItems in the children relationship of a DriveItem.

        https://docs.microsoft.com/en-us/graph/api/driveitem-list-children?view=graph-rest-1.0&tabs=http

        Args:
            folder_id (str): Unique identifier of the folder.
            params (dict, optional): Query. Defaults to None.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = "me/drive/items/{}/children".format(folder_id)
        return self._client._get(self._client.base_url + url, params=params)

    @token_required
    def drive_get_item(self, item_id: str, params: dict = None, **kwargs) -> Response:
        """Retrieve the metadata for a driveItem in a drive by file system path or ID. It may also be the unique ID of a
        SharePoint list item.

        https://docs.microsoft.com/en-us/graph/api/driveitem-get?view=graph-rest-1.0&tabs=http

        Args:
            item_id (str): ID of a driveItem.
            params (dict, optional): Query. Defaults to None.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = "me/drive/items/{}".format(item_id)
        return self._client._get(self._client.base_url + url, params=params, **kwargs)

    @token_required
    def drive_download_contents(self, item_id: str, params: dict = None, **kwargs) -> Response:
        """Download the contents of the primary stream (file) of a DriveItem. Only driveItems with the file property can
        be downloaded.

        https://docs.microsoft.com/en-us/graph/api/driveitem-get-content?view=graph-rest-1.0&tabs=http

        Args:
            item_id (str): ID of a driveItem.
            params (dict, optional): Extra params. Defaults to None.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = "me/drive/items/{}/content".format(item_id)
        return self._client._get(self._client.base_url + url, params=params, **kwargs)

    @token_required
    def drive_download_shared_contents(self, share_id: str, params: dict = None, **kwargs) -> Response:
        """Download the contents of the primary stream (file) of a DriveItem. Only driveItems with the file property can
        be downloaded.

        https://docs.microsoft.com/en-us/graph/api/driveitem-get-content?view=graph-rest-1.0&tabs=http

        Args:
            share_id (str): ID of a driveItem.
            params (dict, optional): Extra params. Defaults to None.

        Returns:
            Response: Microsoft Graph Response.
        """
        base64_value = base64.b64encode(share_id.encode()).decode()
        encoded_share_url = "u!" + base64_value.rstrip("=").replace("/", "_").replace("+", "-")
        url = self._client.base_url + "shares/{}/driveItem".format(encoded_share_url)
        drive_item = self._client._get(url)
        file_download_url = drive_item["@microsoft.graph.downloadUrl"]
        return drive_item["name"], requests.get(file_download_url).content

    @token_required
    def drive_download_large_contents(self, downloadUrl: str, offset: int, size: int) -> Response:
        """Download the contents of the primary stream (file) of a DriveItem. Only driveItems with the file property can
        be downloaded.

        https://docs.microsoft.com/en-us/graph/api/driveitem-get-content?view=graph-rest-1.0&tabs=http

        Args:
            downloadUrl (str): Url of the driveItem.
            offset (int): offset.
            size (int): size.

        Returns:
            Response: Microsoft Graph Response.
        """
        headers = {"Range": f"bytes={offset}-{size + offset - 1}"}
        return self._client._get(downloadUrl, headers=headers)

    @token_required
    def drive_upload_item(self, item_id: str, params: dict = None, **kwargs) -> Response:
        """The simple upload API allows you to provide the contents of a new file or update the contents of an existing
        file in a single API call. This method only supports files up to 4MB in size.

        https://docs.microsoft.com/en-us/graph/api/driveitem-put-content?view=graph-rest-1.0&tabs=http

        Args:
            item_id (str): Id of a driveItem.
            params (dict, optional): Extra params. Defaults to None.

        Returns:
            Response: Microsoft Graph Response.
        """
        kwargs["headers"] = {"Content-Type": "text/plain"}
        url = "me/drive/items/{}/content".format(item_id)
        return self._client._put(self._client.base_url + url, params=params, **kwargs)

    @token_required
    def search_items(self, q: str, params: dict = None, **kwargs) -> Response:
        """Search the hierarchy of items for items matching a query. You can search within a folder hierarchy, a whole
        drive, or files shared with the current user.

        https://docs.microsoft.com/en-us/graph/api/driveitem-search?view=graph-rest-1.0&tabs=http

        Args:
            q (str): The query text used to search for items. Values may be matched across several fields including
            filename, metadata, and file content.
            params (dict, optional): Query. Defaults to None.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = "me/drive/root/search(q='{}')".format(q)
        return self._client._get(self._client.base_url + url, params=params, **kwargs)
