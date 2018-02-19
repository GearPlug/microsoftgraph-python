from microsoftgraph.exceptions import TokenRequired
from functools import wraps


def token_required(func):
    @wraps(func)
    def helper(*args, **kwargs):
        client = args[0]
        if not client.token:
            raise TokenRequired('You must set the Token.')
        return func(*args, **kwargs)

    return helper
