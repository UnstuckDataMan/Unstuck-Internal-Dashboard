"""Read mailbox identities from optional Gmail Setup Tool workspace labels."""
import re


_HOST_LABEL = r'[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?'
_LABELLED_EMAIL = re.compile(
    r'(?P<email>[^\s()]+@[^\s()]+)\s+\((?:Master|Secondary:\s*'
    + _HOST_LABEL + r'(?:\.' + _HOST_LABEL + r')+ is Master)\)',
    re.IGNORECASE,
)


def email_cell_identity(value):
    """Strip only the supported labels; arbitrary trailing text stays invalid."""
    text = str(value or '').strip()
    match = _LABELLED_EMAIL.fullmatch(text)
    return (match.group('email') if match else text).lower()
