from datetime import datetime


def format_time(value: datetime, is_webhook: bool = False) -> str:
    if is_webhook:
        return value.strftime("%Y-%m-%dT%H:%M:%S.%fZ")
    return value.strftime("%Y-%m-%dT%H:%M:%S")
