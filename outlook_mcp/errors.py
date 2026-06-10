"""Error types and COM error formatting."""


class ToolError(Exception):
    """User-facing tool failure (bad input, missing item, COM error)."""


def format_com_error(exc: BaseException) -> str:
    """Turn a pywintypes.com_error into a readable one-line message.

    com_error args are (hresult, strerror, excepinfo, argerror) where
    excepinfo is None or a 6-tuple whose third element is the source
    application's own description (usually the most useful text).
    """
    try:
        hresult = exc.args[0]
        strerror = exc.args[1]
        excepinfo = exc.args[2]
    except (IndexError, AttributeError):
        return f"Outlook COM error: {exc}"
    detail = None
    if excepinfo and len(excepinfo) > 2 and excepinfo[2]:
        detail = str(excepinfo[2]).strip()
    if detail:
        return f"Outlook error: {detail} (HRESULT {hresult})"
    return f"Outlook error: {strerror} (HRESULT {hresult})"
