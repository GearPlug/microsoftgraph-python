from urllib.parse import quote_plus

from microsoftgraph.decorators import token_required
from microsoftgraph.response import Response


class Excel(object):
    def __init__(self, client):
        self._client = client

    @token_required
    def excel_get_worksheets(self, item_id: str, params: dict = None, **kwargs) -> Response:
        url = self._client.base_url + "me/drive/items/{0}/workbook/worksheets".format(item_id)
        return self._client._get(url, params=params, **kwargs)

    @token_required
    def excel_get_names(self, item_id: str, params: dict = None, **kwargs) -> Response:
        url = self._client.base_url + "me/drive/items/{0}/workbook/names".format(item_id)
        return self._client._get(url, params=params, **kwargs)

    @token_required
    def excel_add_worksheet(self, item_id: str, **kwargs) -> Response:
        url = self._client.base_url + "me/drive/items/{0}/workbook/worksheets/add".format(item_id)
        return self._client._post(url, **kwargs)

    @token_required
    def excel_get_specific_worksheet(self, item_id: str, worksheet_id: str, **kwargs) -> Response:
        url = self._client.base_url + "me/drive/items/{0}/workbook/worksheets/{1}".format(
            item_id, quote_plus(worksheet_id)
        )
        return self._client._get(url, **kwargs)

    @token_required
    def excel_update_worksheet(self, item_id: str, worksheet_id: str, **kwargs) -> Response:
        url = self._client.base_url + "me/drive/items/{0}/workbook/worksheets/{1}".format(
            item_id, quote_plus(worksheet_id)
        )
        return self._client._patch(url, **kwargs)

    @token_required
    def excel_get_charts(self, item_id: str, worksheet_id: str, params: dict = None, **kwargs) -> Response:
        url = self._client.base_url + "me/drive/items/{0}/workbook/worksheets/{1}/charts".format(
            item_id, quote_plus(worksheet_id)
        )
        return self._client._get(url, params=params, **kwargs)

    @token_required
    def excel_add_chart(self, item_id: str, worksheet_id: str, **kwargs) -> Response:
        url = self._client.base_url + "me/drive/items/{0}/workbook/worksheets/{1}/charts/add".format(
            item_id, quote_plus(worksheet_id)
        )
        return self._client._post(url, **kwargs)

    @token_required
    def excel_get_tables(self, item_id: str, params: dict = None, **kwargs) -> Response:
        url = self._client.base_url + "me/drive/items/{0}/workbook/tables".format(item_id)
        return self._client._get(url, params=params, **kwargs)

    @token_required
    def excel_add_table(self, item_id: str, **kwargs) -> Response:
        url = self._client.base_url + "me/drive/items/{0}/workbook/tables/add".format(item_id)
        return self._client._post(url, **kwargs)

    @token_required
    def excel_add_column(self, item_id: str, worksheets_id: str, table_id: str, **kwargs) -> Response:
        url = self._client.base_url + "me/drive/items/{0}/workbook/worksheets/{1}/tables/{2}/columns".format(
            item_id, quote_plus(worksheets_id), table_id
        )
        return self._client._post(url, **kwargs)

    @token_required
    def excel_add_row(self, item_id: str, worksheets_id: str, table_id: str, **kwargs) -> Response:
        url = self._client.base_url + "me/drive/items/{0}/workbook/worksheets/{1}/tables/{2}/rows".format(
            item_id, quote_plus(worksheets_id), table_id
        )
        return self._client._post(url, **kwargs)

    @token_required
    def excel_get_rows(self, item_id: str, table_id: str, params: dict = None, **kwargs) -> Response:
        url = self._client.base_url + "me/drive/items/{0}/workbook/tables/{1}/rows".format(item_id, table_id)
        return self._client._get(url, params=params, **kwargs)

    # @token_required
    # def excel_get_cell(self, item_id, worksheets_id, params=None, **kwargs):
    #     url = self.base_url + "me/drive/items/{0}/workbook/worksheets/{1}/Cell(row='1', column='A')".format(item_id, quote_plus(worksheets_id))
    #     return self._get(url, params=params, **kwargs)

    # @token_required
    # def excel_add_cell(self, item_id, worksheets_id, **kwargs):
    #     url = self.base_url + "me/drive/items/{0}/workbook/worksheets/{1}/rows".format(item_id, worksheets_id)
    #     return self._patch(url, **kwargs)

    @token_required
    def excel_get_range(self, item_id: str, worksheets_id: str, **kwargs) -> Response:
        url = self._client.base_url + "me/drive/items/{0}/workbook/worksheets/{1}/range(address='A1:B2')".format(
            item_id, quote_plus(worksheets_id)
        )
        return self._client._get(url, **kwargs)

    @token_required
    def excel_update_range(self, item_id: str, worksheets_id: str, **kwargs) -> Response:
        url = self._client.base_url + "me/drive/items/{0}/workbook/worksheets/{1}/range(address='A1:B2')".format(
            item_id, quote_plus(worksheets_id)
        )
        return self._client._patch(url, **kwargs)
