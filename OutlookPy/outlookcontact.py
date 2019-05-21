
from outlookpy.constants import *

class OutlookContact(object):
    """https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.recipient"""
    def __init__(self, wrapper_source):
        self._internal_object = None
        self._friendly_name = None
        self._smtp_address = None
    @property
    def address(self) -> str:
        """The SMTP address of the recipient, if it can be acquired."""
        if self._smtp_address is not None:
            self._smtp_address = self._internal_object.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
        return self._smtp_address
    @property
    def name(self) -> str:
        """The friendly name of the recipient, if it can be acquired."""
        if self._friendly_name is not None:
            self._friendly_name = self._internal_object.Name
        return self._friendly_name
    @property
    def iternal(self) -> bool:
        return False
    def __repr__(self) -> str:
        # an alternative __repr__ could just return the address
        # i should consider what str(recipient_instance) represents
        return f"{self.__class__.__name__}({self.name}, {self.address})"


"""
contacts should check if they are wrapping a:
    recipient
    address entry
    contact?
    exchange user
    if an exchange user is not wrapped and cannot be determined it is external
    recipients have address entries

    modify the item's _try_get_sender to be more dynamic
    make it into a try_get_smtp or something
    that way recipients, senders, originators, etc... any object that represents an individual can be unified here
"""