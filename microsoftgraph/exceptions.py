class BaseError(Exception):
    pass


class UnknownError(BaseError):
    pass


class TokenRequired(BaseError):
    pass


class BadRequest(BaseError):
    pass


class Unauthorized(BaseError):
    pass


class Forbidden(BaseError):
    pass


class NotFound(BaseError):
    pass


class MethodNotAllowed(BaseError):
    pass


class NotAcceptable(BaseError):
    pass


class Conflict(BaseError):
    pass


class Gone(BaseError):
    pass


class LengthRequired(BaseError):
    pass


class PreconditionFailed(BaseError):
    pass


class RequestEntityTooLarge(BaseError):
    pass


class UnsupportedMediaType(BaseError):
    pass


class RequestedRangeNotSatisfiable(BaseError):
    pass


class UnprocessableEntity(BaseError):
    pass


class TooManyRequests(BaseError):
    pass


class InternalServerError(BaseError):
    pass


class NotImplemented(BaseError):
    pass


class ServiceUnavailable(BaseError):
    pass


class GatewayTimeout(BaseError):
    pass


class InsufficientStorage(BaseError):
    pass


class BandwidthLimitExceeded(BaseError):
    pass
