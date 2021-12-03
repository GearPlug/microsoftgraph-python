from datetime import datetime, timedelta


class Response:
    def __init__(self, original) -> None:
        self.original = original

    def __repr__(self) -> str:
        return "<Response [{}] ({})>".format(self.status_code, self.data)

    @property
    def data(self):
        if "application/json" in self.original.headers.get("Content-Type", ""):
            return self.original.json()
        return self.original.content

    @property
    def status_code(self):
        return self.original.status_code

    @property
    def throttling(self) -> datetime:
        """Microsoft Graph throttling

        https://docs.microsoft.com/en-us/graph/throttling

        Returns:
            datetime: Retry after.
        """
        if "Retry-After" in self.original.headers:
            return datetime.now() + timedelta(seconds=self.original.headers["Retry-After"])
        return None
