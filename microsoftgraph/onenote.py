from microsoftgraph.decorators import token_required
from microsoftgraph.response import Response


class Onenote(object):
    def __init__(self, client):
        self._client = client

    @token_required
    def list_notebooks(self) -> Response:
        """Retrieve a list of notebook objects.

        Returns:
            A dict.

        """
        return self._client._get(self._client.base_url + "me/onenote/notebooks")

    @token_required
    def get_notebook(self, notebook_id: str) -> Response:
        """Retrieve the properties and relationships of a notebook object.

        Args:
            notebook_id:

        Returns:
            A dict.

        """
        return self._client._get(self._client.base_url + "me/onenote/notebooks/" + notebook_id)

    @token_required
    def get_notebook_sections(self, notebook_id: str) -> Response:
        """Retrieve the properties and relationships of a notebook object.

        Args:
            notebook_id:

        Returns:
            A dict.

        """
        return self._client._get(self._client.base_url + "me/onenote/notebooks/{}/sections".format(notebook_id))

    @token_required
    def create_page(self, section_id: str, files: list) -> Response:
        """Create a new page in the specified section.

        Args:
            section_id:
            files:

        Returns:
            A dict.

        """
        return self._client._post(
            self._client.base_url + "/me/onenote/sections/{}/pages".format(section_id), files=files
        )

    @token_required
    def list_pages(self, params: dict = None) -> Response:
        """Create a new page in the specified section.

        Args:
            params:

        Returns:
            A dict.

        """
        return self._client._get(self._client.base_url + "/me/onenote/pages", params=params)
