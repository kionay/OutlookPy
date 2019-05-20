"""All outlook item wrappers."""
from datetime import datetime
from typing import List, Tuple, TYPE_CHECKING

from pywintypes import com_error

from outlookpy.constants import *
if TYPE_CHECKING:
    from outlookpy.outlookfolder import OutlookFolder
from .outlookenumerations import OutlookItemType, OutlookResponse, OutlookItemImportance, OutlookItemBodyFormat


class OutlookItem(object):
    """
    Base wrapping class for outlook items.
    Represents the common functions of all other outlook item types.
    May need divided into AppointmentItem, JournalItem, MailItem, MeetingItem, and TaskItem.
    """
    def __init__(self, mail_item):
        self._mail_item = mail_item
        self._sender = None
        self._recipients = None
    @property
    def _local_id(self):
        """Closest thing to a unique ID we're going to get for an outlook item"""
        return self._mail_item.EntryID
    def delete(self):
        """moves the item to the Deleted Items folder, does not permanently delete unless it's already in that folder"""
        self._mail_item.Delete()
    def move(self, folder):
        self._mail_item = self._mail_item.Move(folder._folder) 
    @property
    def containing_folder(self):
        return OutlookFolder(self._mail_item.Parent)
    @property
    def recipients(self) -> List[str]:
        # recipeints might have to make an external call to get this information
        # if we already have it for this mail item, we don't need to call the server again
        # the recipients aren't going to spontaneously change
        if self._recipients is not None:
            return self._recipients
        recipient_addresses = []
        for recipient in self._mail_item.Recipients:
            recipient_addresses.append(recipient.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS))
        self._recipients = recipient_addresses
        return recipient_addresses
    @property
    def categories(self) -> List[str]:
        categories = self._mail_item.Categories.split(", ")
        if categories == ['']:
            return []
        return categories
    @property
    def read(self) -> bool:
        return not self._mail_item.UnRead
    @read.setter
    def read(self, read_status: bool):
        self._mail_item.UnRead = not read_status
    @property
    def unread(self) -> bool:
        return self._mail_item.UnRead
    @unread.setter
    def unread(self, unread_status: bool):
        self._mail_item.UnRead = unread_status
    def _try_get_sender_remote(self):
        remote_properties = [
            PR_SENT_REPRESENTING_EMAIL_ADDRESS_W,
            PR_SENT_REPRESENTING_SMTP_ADDRESS,
            PR_MEETING_SENDER_SMTP_ADDRESS,
            PR_SMTP_ADDRESS,
            PR_SENDER_SMTP_ADDRESS,
            PR_LAST_MODIFIER_NAME_W]
        core_mail_item = self._mail_item
        sender_sample = None
        for remote_property in remote_properties:
            try:
                sender_sample = core_mail_item.PropertyAccessor.GetProperty(remote_property)
            except Exception:
                sender_sample = None
            finally:
                if sender_sample is None or not sender_sample:
                    pass
                elif "@" not in sender_sample:
                    pass
                else:
                    break
        return sender_sample
    @property
    def sender(self) -> str:
        if self._sender is not None:
            return self._sender
        if self.type == OutlookItemType.Task:
            return None # Tasks are not "sent" so they can have no sender
        smtp = self._try_get_sender_remote()
        if smtp is not None and "@" in smtp:
            self._sender = smtp
            return smtp
        return None

    @property
    def body(self) -> str:
        return self._mail_item.Body
    @property
    def subject(self) -> str:
        return self._mail_item.Subject
    @property
    def external(self) -> bool:
        # Sender Email Type 'EX' stands for 'EXchange' not 'external
        # i have only ever seen the SenderEmailType be either "EX" or "SMTP"
        return self._mail_item.SenderEmailType != "EX"
    @property
    def internal(self) -> bool:
        # Sender Email Type 'EX' stands for 'EXchange' not 'external
        return self._mail_item.SenderEmailType == "EX"
    @property
    def type(self) -> OutlookItemType:
        for item_type in OutlookItemType:
            if self._mail_item.Class in item_type.value:
                return item_type
        return None
    @property
    def importance(self) -> OutlookItemImportance:
        for item_importance in OutlookItemImportance:
            if self._mail_item.Importance == item_importance.value:
                return item_importance
        return None
    @importance.setter
    def importance(self, importance: OutlookItemImportance):
        if not isinstance(importance, OutlookItemImportance):
            raise TypeError("importance must be of type OutlookItemImportance")
        self._mail_item.Importance = importance.value
    @property
    def received_datetime(self) -> datetime:
        return self._mail_item.ReceivedTime
    @received_datetime.setter
    def received_datetime(self, received_datetime: datetime):
        self._mail_item.ReceivedTime = received_datetime
    @property
    def body_format(self) -> str:
        return OutlookItemBodyFormat(self._mail_item.BodyFormat).name
    @body_format.setter
    def body_format(self, body_format: str):
        for possible_format in OutlookItemBodyFormat:
            if possible_format.name == body_format.lower():
                self._mail_item.BodyFormat = possible_format.value
                return
        raise ValueError(f"Body Format ({body_format}) is not a valid format for the body of this item.")
    def __repr__(self):
        return f"{self.__class__.__name__}({self.subject})"
    def __hash__(self):
        return hash(self._local_id)

class OutlookMailItem(OutlookItem):
    """https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.mailitem?view=outlook-pia"""

class OutlookAppointmentItem(OutlookItem):
    """https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.appointmentitem?view=outlook-pia"""

class OutlookMeetingItem(OutlookItem):
    """https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.meetingitem?view=outlook-pia"""
    MeetingResponse = Tuple[str, OutlookResponse]
    MeetingResponses = List[MeetingResponse]
    @property
    def responses(self) -> MeetingResponses:
        if self._responses is not None:
            return self._responses
        responses = []
        for recipient in self._mail_item.Recipients:
            responses.append(recipient.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS), OutlookResponse(recipient.MeetingResponseStatus))
        self._responses = responses
        return responses

class OutlookJournalItem(OutlookItem):
    """https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.journalitem?view=outlook-pia"""

class OutlookTaskItem(OutlookItem):
    """https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.taskitem?view=outlook-pia"""



