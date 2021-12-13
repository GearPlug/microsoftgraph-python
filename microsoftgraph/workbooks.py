from urllib.parse import quote_plus

from microsoftgraph.decorators import token_required
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
        url = self._client.base_url + "me/drive/items/{0}/workbook/createSession".format(workbook_id)
        return self._client._post(url, **kwargs)

    @token_required
    def refresh_session(self, workbook_id: str, **kwargs) -> Response:
        """Refresh an existing workbook session.

        https://docs.microsoft.com/en-us/graph/api/workbook-refreshsession?view=graph-rest-1.0&tabs=http

        Args:
            workbook_id (str): Excel file ID.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = self._client.base_url + "me/drive/items/{0}/workbook/refreshSession".format(workbook_id)
        return self._client._post(url, **kwargs)

    @token_required
    def close_session(self, workbook_id: str, **kwargs) -> Response:
        """Close an existing workbook session.

        https://docs.microsoft.com/en-us/graph/api/workbook-closesession?view=graph-rest-1.0&tabs=http

        Args:
            workbook_id (str): Excel file ID.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = self._client.base_url + "me/drive/items/{0}/workbook/closeSession".format(workbook_id)
        return self._client._post(url, **kwargs)

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
        url = self._client.base_url + "me/drive/items/{0}/workbook/worksheets".format(workbook_id)
        return self._client._get(url, params=params, **kwargs)

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
        url = self._client.base_url + "me/drive/items/{0}/workbook/names".format(workbook_id)
        return self._client._get(url, params=params, **kwargs)

    @token_required
    def add_worksheet(self, workbook_id: str, **kwargs) -> Response:
        """Adds a new worksheet to the workbook.

        https://docs.microsoft.com/en-us/graph/api/worksheetcollection-add?view=graph-rest-1.0&tabs=http

        Args:
            workbook_id (str): Excel file ID.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = self._client.base_url + "me/drive/items/{0}/workbook/worksheets/add".format(workbook_id)
        return self._client._post(url, **kwargs)

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
        url = self._client.base_url + "me/drive/items/{0}/workbook/worksheets/{1}".format(
            workbook_id, quote_plus(worksheet_id)
        )
        return self._client._get(url, **kwargs)

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
        url = self._client.base_url + "me/drive/items/{0}/workbook/worksheets/{1}".format(
            workbook_id, quote_plus(worksheet_id)
        )
        return self._client._patch(url, **kwargs)

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
        url = self._client.base_url + "me/drive/items/{0}/workbook/worksheets/{1}/charts".format(
            workbook_id, quote_plus(worksheet_id)
        )
        return self._client._get(url, params=params, **kwargs)

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
        url = self._client.base_url + "me/drive/items/{0}/workbook/worksheets/{1}/charts/add".format(
            workbook_id, quote_plus(worksheet_id)
        )
        return self._client._post(url, **kwargs)

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
        url = self._client.base_url + "me/drive/items/{0}/workbook/tables".format(workbook_id)
        return self._client._get(url, params=params, **kwargs)

    @token_required
    def add_table(self, workbook_id: str, **kwargs) -> Response:
        """Create a new table.

        https://docs.microsoft.com/en-us/graph/api/tablecollection-add?view=graph-rest-1.0&tabs=http

        Args:
            workbook_id (str): Excel file ID.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = self._client.base_url + "me/drive/items/{0}/workbook/tables/add".format(workbook_id)
        return self._client._post(url, **kwargs)

    @token_required
    def create_column(self, workbook_id: str, worksheet_id: str, table_id: str, **kwargs) -> Response:
        """Create a new TableColumn.

        https://docs.microsoft.com/en-us/graph/api/table-post-columns?view=graph-rest-1.0&tabs=http

        Args:
            workbook_id (str): Excel file ID.
            worksheet_id (str): Excel worksheet ID.
            table_id (str): Excel table ID.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = self._client.base_url + "me/drive/items/{0}/workbook/worksheets/{1}/tables/{2}/columns".format(
            workbook_id, quote_plus(worksheet_id), table_id
        )
        return self._client._post(url, **kwargs)

    @token_required
    def create_row(self, workbook_id: str, worksheet_id: str, table_id: str, **kwargs) -> Response:
        """Adds rows to the end of a table.

        https://docs.microsoft.com/en-us/graph/api/table-post-rows?view=graph-rest-1.0&tabs=http

        Args:
            workbook_id (str): Excel file ID.
            worksheet_id (str): Excel worksheet ID.
            table_id (str): Excel table ID.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = self._client.base_url + "me/drive/items/{0}/workbook/worksheets/{1}/tables/{2}/rows".format(
            workbook_id, quote_plus(worksheet_id), table_id
        )
        return self._client._post(url, **kwargs)

    @token_required
    def list_rows(self, workbook_id: str, table_id: str, params: dict = None, **kwargs) -> Response:
        """Retrieve a list of tablerow objects.

        https://docs.microsoft.com/en-us/graph/api/table-list-rows?view=graph-rest-1.0&tabs=http

        Args:
            workbook_id (str): Excel file ID.
            table_id (str): Excel table ID.
            params (dict, optional): Query. Defaults to None.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = self._client.base_url + "me/drive/items/{0}/workbook/tables/{1}/rows".format(workbook_id, table_id)
        return self._client._get(url, params=params, **kwargs)

    # @token_required
    # def excel_get_cell(self, item_id, worksheet_id, params=None, **kwargs):
    #     url = self.base_url + "me/drive/items/{0}/workbook/worksheets/{1}/Cell(row='1', column='A')".format(item_id, quote_plus(worksheet_id))
    #     return self._get(url, params=params, **kwargs)

    # @token_required
    # def excel_add_cell(self, item_id, worksheet_id, **kwargs):
    #     url = self.base_url + "me/drive/items/{0}/workbook/worksheets/{1}/rows".format(item_id, worksheet_id)
    #     return self._patch(url, **kwargs)

    @token_required
    def get_range(self, workbook_id: str, worksheet_id: str, **kwargs) -> Response:
        """Gets the range object specified by the address or name.

        https://docs.microsoft.com/en-us/graph/api/worksheet-range?view=graph-rest-1.0&tabs=http

        Args:
            workbook_id (str): Excel file ID.
            worksheet_id (str): Excel worksheet ID.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = self._client.base_url + "me/drive/items/{0}/workbook/worksheets/{1}/range(address='A1:B2')".format(
            workbook_id, quote_plus(worksheet_id)
        )
        return self._client._get(url, **kwargs)

    @token_required
    def update_range(self, workbook_id: str, worksheet_id: str, **kwargs) -> Response:
        """Update the properties of range object.

        https://docs.microsoft.com/en-us/graph/api/range-update?view=graph-rest-1.0&tabs=http

        Args:
            workbook_id (str): Excel file ID.
            worksheet_id (str): Excel worksheet ID.

        Returns:
            Response: Microsoft Graph Response.
        """
        url = self._client.base_url + "me/drive/items/{0}/workbook/worksheets/{1}/range(address='A1:B2')".format(
            workbook_id, quote_plus(worksheet_id)
        )
        return self._client._patch(url, **kwargs)
