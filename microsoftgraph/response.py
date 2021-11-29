class Response:
    def __init__(self, original) -> None:
        self.original = original

    @property
    def data(self):
        if "application/json" in self.original.headers.get("Content-Type", ""):
            return self.original.json()
        else:
            return self.original.content

    @property
    def status_code(self):
        return self.original.status_code
