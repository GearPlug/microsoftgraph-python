from urllib.parse import quote_plus

from microsoftgraph.decorators import token_required, workbook_session_id_required
from microsoftgraph.response import Response


class Workbooks(object):
    def __init__(self, client) -> None:
        """Working with Excel in Microsoft Graph

        https://docs.microsoft.com/en-us/graph/api/resources/excel?view=graph-rest-1.0

        Args:
            client (Client): Library Client.
        """
        self._client = client

    @token_required
    def create_session(self, workbook_id: str, **kwargs) -> Response:
        """Create a new workbook session.

        https://docs.microsoft.com/en-us/graph/api/workbook-createsession?view=graph-rest-1.0&tabs=http

        Args:
            workbook_id (str): Excel file ID.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = "me/drive/items/{}/workbook/createSession".format(workbook_id)
        return self._client._post(self._client.base_url + url, **kwargs)

    @token_required
    @workbook_session_id_required
    def refresh_session(self, workbook_id: str, **kwargs) -> Response:
        """Refresh an existing workbook session.

        https://docs.microsoft.com/en-us/graph/api/workbook-refreshsession?view=graph-rest-1.0&tabs=http

        Args:
            workbook_id (str): Excel file ID.

        Returns:
            Response: Microsoft Graph Response.
        """
        headers = {"workbook-session-id": self._client.workbook_session_id}
        url = "me/drive/items/{}/workbook/refreshSession".format(workbook_id)
        return self._client._post(self._client.base_url + url, headers=headers, **kwargs)

    @token_required
    @workbook_session_id_required
    def close_session(self, workbook_id: str, **kwargs) -> Response:
        """Close an existing workbook session.

        https://docs.microsoft.com/en-us/graph/api/workbook-closesession?view=graph-rest-1.0&tabs=http

        Args:
            workbook_id (str): Excel file ID.

        Returns:
            Response: Microsoft Graph Response.
        """
        headers = {"workbook-session-id": self._client.workbook_session_id}
        url = "me/drive/items/{}/workbook/closeSession".format(workbook_id)
        return self._client._post(self._client.base_url + url, headers=headers, **kwargs)

    @token_required
    def list_names(self, workbook_id: str, params: dict = None, **kwargs) -> Response:
        """Retrieve a list of nameditem objects.

        https://docs.microsoft.com/en-us/graph/api/workbook-list-names?view=graph-rest-1.0&tabs=http

        Args:
            workbook_id (str): Excel file ID.
            params (dict, optional): Query. Defaults to None.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = "me/drive/items/{}/workbook/names".format(workbook_id)
        return self._client._get(self._client.base_url + url, params=params, **kwargs)

    @token_required
    def list_worksheets(self, workbook_id: str, params: dict = None, **kwargs) -> Response:
        """Retrieve a list of worksheet objects.

        https://docs.microsoft.com/en-us/graph/api/workbook-list-worksheets?view=graph-rest-1.0&tabs=http

        Args:
            workbook_id (str):  Excel file ID.
            params (dict, optional): Query. Defaults to None.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = "me/drive/items/{}/workbook/worksheets".format(workbook_id)
        return self._client._get(self._client.base_url + url, params=params, **kwargs)

    @token_required
    def get_worksheet(self, workbook_id: str, worksheet_id: str, **kwargs) -> Response:
        """Retrieve the properties and relationships of worksheet object.

        https://docs.microsoft.com/en-us/graph/api/worksheet-get?view=graph-rest-1.0&tabs=http

        Args:
            workbook_id (str): Excel file ID.
            worksheet_id (str): Excel worksheet ID.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = "me/drive/items/{}/workbook/worksheets/{}".format(workbook_id, quote_plus(worksheet_id))
        return self._client._get(self._client.base_url + url, **kwargs)

    @token_required
    def add_worksheet(self, workbook_id: str, **kwargs) -> Response:
        """Adds a new worksheet to the workbook.

        https://docs.microsoft.com/en-us/graph/api/worksheetcollection-add?view=graph-rest-1.0&tabs=http

        Args:
            workbook_id (str): Excel file ID.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = "me/drive/items/{}/workbook/worksheets/add".format(workbook_id)
        return self._client._post(self._client.base_url + url, **kwargs)

    @token_required
    def update_worksheet(self, workbook_id: str, worksheet_id: str, **kwargs) -> Response:
        """Update the properties of worksheet object.

        https://docs.microsoft.com/en-us/graph/api/worksheet-update?view=graph-rest-1.0&tabs=http

        Args:
            workbook_id (str): Excel file ID.
            worksheet_id (str): Excel worksheet ID.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = "me/drive/items/{}/workbook/worksheets/{}".format(workbook_id, quote_plus(worksheet_id))
        return self._client._patch(self._client.base_url + url, **kwargs)

    @token_required
    def list_charts(self, workbook_id: str, worksheet_id: str, params: dict = None, **kwargs) -> Response:
        """Retrieve a list of chart objects.

        https://docs.microsoft.com/en-us/graph/api/worksheet-list-charts?view=graph-rest-1.0&tabs=http

        Args:
            workbook_id (str): Excel file ID.
            worksheet_id (str): Excel worksheet ID.
            params (dict, optional): Query. Defaults to None.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = "me/drive/items/{}/workbook/worksheets/{}/charts".format(workbook_id, quote_plus(worksheet_id))
        return self._client._get(self._client.base_url + url, params=params, **kwargs)

    @token_required
    def add_chart(self, workbook_id: str, worksheet_id: str, **kwargs) -> Response:
        """Creates a new chart.

        https://docs.microsoft.com/en-us/graph/api/chartcollection-add?view=graph-rest-1.0&tabs=http

        Args:
            workbook_id (str): Excel file ID.
            worksheet_id (str): Excel worksheet ID.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = "me/drive/items/{}/workbook/worksheets/{}/charts/add".format(workbook_id, quote_plus(worksheet_id))
        return self._client._post(self._client.base_url + url, **kwargs)

    @token_required
    def list_tables(self, workbook_id: str, params: dict = None, **kwargs) -> Response:
        """Retrieve a list of table objects.

        https://docs.microsoft.com/en-us/graph/api/workbook-list-tables?view=graph-rest-1.0&tabs=http

        Args:
            workbook_id (str): Excel file ID.
            params (dict, optional): Query. Defaults to None.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = "me/drive/items/{}/workbook/tables".format(workbook_id)
        return self._client._get(self._client.base_url + url, params=params, **kwargs)

    @token_required
    def add_table(self, workbook_id: str, **kwargs) -> Response:
        """Create a new table.

        https://docs.microsoft.com/en-us/graph/api/tablecollection-add?view=graph-rest-1.0&tabs=http

        Args:
            workbook_id (str): Excel file ID.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = "me/drive/items/{}/workbook/tables/add".format(workbook_id)
        return self._client._post(self._client.base_url + url, **kwargs)

    @token_required
    def create_table_column(self, workbook_id: str, worksheet_id: str, table_id: str, **kwargs) -> Response:
        """Create a new TableColumn.

        https://docs.microsoft.com/en-us/graph/api/table-post-columns?view=graph-rest-1.0&tabs=http

        Args:
            workbook_id (str): Excel file ID.
            worksheet_id (str): Excel worksheet ID.
            table_id (str): Excel table ID.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = "me/drive/items/{}/workbook/worksheets/{}/tables/{}/columns".format(
            workbook_id, quote_plus(worksheet_id), table_id
        )
        return self._client._post(self._client.base_url + url, **kwargs)

    @token_required
    def create_table_row(self, workbook_id: str, worksheet_id: str, table_id: str, **kwargs) -> Response:
        """Adds rows to the end of a table.

        https://docs.microsoft.com/en-us/graph/api/table-post-rows?view=graph-rest-1.0&tabs=http

        Args:
            workbook_id (str): Excel file ID.
            worksheet_id (str): Excel worksheet ID.
            table_id (str): Excel table ID.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = "me/drive/items/{}/workbook/worksheets/{}/tables/{}/rows".format(
            workbook_id, quote_plus(worksheet_id), table_id
        )
        return self._client._post(self._client.base_url + url, **kwargs)

    @token_required
    def list_table_rows(self, workbook_id: str, table_id: str, params: dict = None, **kwargs) -> Response:
        """Retrieve a list of tablerow objects.

        https://docs.microsoft.com/en-us/graph/api/table-list-rows?view=graph-rest-1.0&tabs=http

        Args:
            workbook_id (str): Excel file ID.
            table_id (str): Excel table ID.
            params (dict, optional): Query. Defaults to None.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = "me/drive/items/{}/workbook/tables/{}/rows".format(workbook_id, table_id)
        return self._client._get(self._client.base_url + url, params=params, **kwargs)

    @token_required
    def get_range(self, workbook_id: str, worksheet_id: str, address: str, **kwargs) -> Response:
        """Gets the range object specified by the address or name.

        https://docs.microsoft.com/en-us/graph/api/worksheet-range?view=graph-rest-1.0&tabs=http

        Args:
            workbook_id (str): Excel file ID.
            worksheet_id (str): Excel worksheet ID.
            address (str): Address.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = "me/drive/items/{}/workbook/worksheets/{}/range(address='{}')".format(
            workbook_id, quote_plus(worksheet_id), address
        )
        return self._client._get(self._client.base_url + url, **kwargs)

    @token_required
    def get_used_range(self, workbook_id: str, worksheet_id: str, **kwargs) -> Response:
        """The used range is the smallest range that encompasses any cells that have a value or formatting assigned to
        them. If the worksheet is blank, this function will return the top left cell.

        https://docs.microsoft.com/en-us/graph/api/worksheet-usedrange?view=graph-rest-1.0&tabs=http

        Args:
            workbook_id (str): Excel file ID.
            worksheet_id (str): Excel worksheet ID.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = "me/drive/items/{}/workbook/worksheets/{}/usedRange".format(workbook_id, quote_plus(worksheet_id))
        return self._client._get(self._client.base_url + url, **kwargs)

    @token_required
    @workbook_session_id_required
    def update_range(self, workbook_id: str, worksheet_id: str, address: str, **kwargs) -> Response:
        """Update the properties of range object.

        https://docs.microsoft.com/en-us/graph/api/range-update?view=graph-rest-1.0&tabs=http

        Args:
            workbook_id (str): Excel file ID.
            worksheet_id (str): Excel worksheet ID.
            address (str): Address.

        Returns:
            Response: Microsoft Graph Response.
        """
        headers = {"workbook-session-id": self._client.workbook_session_id}
        url = "me/drive/items/{}/workbook/worksheets/{}/range(address='{}')".format(
            workbook_id, quote_plus(worksheet_id), address
        )
        return self._client._patch(self._client.base_url + url, headers=headers, **kwargs)
