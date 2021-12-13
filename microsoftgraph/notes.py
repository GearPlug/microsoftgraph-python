from microsoftgraph.decorators import token_required
from microsoftgraph.response import Response


class Notes(object):
    def __init__(self, client) -> None:
        """Use the OneNote REST API.

        https://docs.microsoft.com/en-us/graph/api/resources/onenote-api-overview?view=graph-rest-1.0

        Args:
            client (Client): Library Client.
        """
        self._client = client

    @token_required
    def list_notebooks(self, params: dict = None) -> Response:
        """Retrieve a list of notebook objects.

        https://docs.microsoft.com/en-us/graph/api/onenote-list-notebooks?view=graph-rest-1.0&tabs=http

        Args:
            params (dict, optional): Query. Defaults to None.

        Returns:
            Response: Microsoft Graph Response.
        """
        return self._client._get(self._client.base_url + "me/onenote/notebooks", params=params)

    @token_required
    def get_notebook(self, notebook_id: str, params: dict = None) -> Response:
        """Retrieve the properties and relationships of a notebook object.

        https://docs.microsoft.com/en-us/graph/api/notebook-get?view=graph-rest-1.0&tabs=http

        Args:
            notebook_id (str): The unique identifier of the notebook.
            params (dict, optional): Query. Defaults to None.

        Returns:
            Response: Microsoft Graph Response.
        """
        return self._client._get(self._client.base_url + "me/onenote/notebooks/" + notebook_id, params=params)

    @token_required
    def list_sections(self, notebook_id: str, params: dict = None) -> Response:
        """Retrieve a list of onenoteSection objects from the specified notebook.

        https://docs.microsoft.com/en-us/graph/api/notebook-list-sections?view=graph-rest-1.0&tabs=http

        Args:
            notebook_id (str): The unique identifier of the notebook.
            params (dict, optional): Query. Defaults to None.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = "me/onenote/notebooks/{}/sections".format(notebook_id)
        return self._client._get(self._client.base_url + url, params=params)

    @token_required
    def list_pages(self, params: dict = None) -> Response:
        """Retrieve a list of page objects.

        Args:
            params (dict, optional): Query. Defaults to None.

        Returns:
            Response: Microsoft Graph Response.
        """
        return self._client._get(self._client.base_url + "me/onenote/pages", params=params)

    @token_required
    def create_page(self, section_id: str, files: list) -> Response:
        """Create a new page in the specified section.

        https://docs.microsoft.com/en-us/graph/api/section-post-pages?view=graph-rest-1.0

        Args:
            section_id (str): The unique identifier of the section.
            files (list): Attachments.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = "me/onenote/sections/{}/pages".format(section_id)
        return self._client._post(self._client.base_url + url, files=files)
