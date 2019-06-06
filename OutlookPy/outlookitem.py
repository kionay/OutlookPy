"""All outlook item wrappers."""
from datetime import datetime
from typing import List, Tuple, TYPE_CHECKING

from pywintypes import com_error

from outlookpy.constants import *
if TYPE_CHECKING:
    from outlookpy.outlookfolder import OutlookFolder
    from outlookpy.helpers import com_to_python
from .outlookenumerations import OutlookResponse, OutlookItemImportance, OutlookItemBodyFormat, OutlookTaskResponse, OutlookTaskStatus, OutlookRecipientType, OutlookShowAs


class OutlookItem(object):
    """
    Base wrapping class for outlook items.
    Represents the common functions of all other outlook item types.
    """
    def __init__(self, outlook_item):
        self._internal_item = outlook_item
        self._sender = None
        self._recipients = None
        self._parent = None
    @property
    def _local_id(self):
        """Closest thing to a unique ID we're going to get for an outlook item"""
        return self._internal_item.EntryID
    def delete(self):
        """moves the item to the Deleted Items folder, does not permanently delete unless it's already in that folder"""
        self._internal_item.Delete()
    def move(self, folder):
        self._internal_item = self._internal_item.Move(folder._folder) 
    @property
    def containing_folder(self):
        return OutlookFolder(self._internal_item.Parent)
    @property
    def parent(self):
        """
        The parent of an outlook item is that item that came before it in its conversation.
        """
        if self._parent is None:
            this_convo = self._internal_item.GetConversation()
            this_parent = this_convo.GetParent(self._internal_item)
            if this_parent is not None:
                self._parent = com_to_python(this_parent)
            else:
                self._parent = None
        return self._parent
    @property
    def children(self):
        """
        The children of an outlook item are those that came after it in its conversation.
        """
        this_convo = self._internal_item.GetConversation()
        children = [com_to_python(obj) for obj in this_convo.GetChildren(self._internal_item)._dispobj_]
        return children
    @property
    def recipients(self) -> List[str]:
        """
            A list of those that were intended to recieve this item.
            Currently a list of SMTP addresses, in the future an OutlookContact object.
        """
        # recipeints might have to make an external call to get this information
        # if we already have it for this mail item, we don't need to call the server again
        # the recipients aren't going to spontaneously change
        if self._recipients is not None:
            return self._recipients
        recipient_addresses = []
        for recipient in self._internal_item.Recipients:
            recipient_addresses.append(recipient.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS))
        self._recipients = recipient_addresses
        return recipient_addresses
    @property
    def categories(self) -> List[str]:
        categories = self._internal_item.Categories.split(", ")
        if categories == ['']:
            return []
        return categories
    @property
    def read(self) -> bool:
        return not self._internal_item.UnRead
    @read.setter
    def read(self, read_status: bool):
        self._internal_item.UnRead = not read_status
    @property
    def unread(self) -> bool:
        return self._internal_item.UnRead
    @unread.setter
    def unread(self, unread_status: bool):
        self._internal_item.UnRead = unread_status
    def _try_get_sender_remote(self):
        """
        Attempt to get the SMTP that sent this item.
        Sadly, this kludge of a solution works more reliably than the most sophisticated 'proper' log I could create.
        """
        # Any given property could be the one that has the SMTP we want.
        # I have tried to order these properties from most to least likely to be the one we need.
        # In the far future I imagine running some speed tests on these to see what the distribution of
        #   properties looks like, in order to order these in as optimal a way as possible.
        remote_properties = [
            PR_SENT_REPRESENTING_EMAIL_ADDRESS_W,
            PR_SENT_REPRESENTING_SMTP_ADDRESS,
            PR_MEETING_SENDER_SMTP_ADDRESS,
            PR_SMTP_ADDRESS,
            PR_SENDER_SMTP_ADDRESS,
            PR_LAST_MODIFIER_NAME_W]
        core_internal_item = self._internal_item
        sender_sample = None
        for remote_property in remote_properties:
            try:
                # try to get each property in our list
                sender_sample = core_internal_item.PropertyAccessor.GetProperty(remote_property)
            except Exception:
                # if there is an error, supress it.
                sender_sample = None
            finally:
                # if there was not an error, does its format meet our criteria for an 'acceptable' SMTP address?
                if sender_sample is None or not sender_sample:
                    pass
                elif "@" not in sender_sample:
                    pass
                else:
                    # if we have met all criteria (no error, contains an @ symbol) stop trying more properties
                    break
        return sender_sample
    @property
    def sender(self) -> str:
        if self._sender is not None:
            return self._sender
        if type(self) is OutlookTaskItem:
            return None # Tasks are not "sent" so they can have no sender
        smtp = self._try_get_sender_remote()
        if smtp is not None and "@" in smtp:
            self._sender = smtp
            return smtp
        return None

    @property
    def body(self) -> str:
        return self._internal_item.Body
    @property
    def subject(self) -> str:
        return self._internal_item.Subject
    @property
    def external(self) -> bool:
        # Sender Email Type 'EX' stands for 'EXchange' not 'external
        # i have only ever seen the SenderEmailType be either "EX" or "SMTP"
        return self._internal_item.SenderEmailType != "EX"
    @property
    def internal(self) -> bool:
        # Sender Email Type 'EX' stands for 'EXchange' not 'external
        return self._internal_item.SenderEmailType == "EX"
    @property
    def importance(self) -> OutlookItemImportance:
        for item_importance in OutlookItemImportance:
            if self._internal_item.Importance == item_importance.value:
                return item_importance
        return None
    @importance.setter
    def importance(self, importance: OutlookItemImportance):
        if not isinstance(importance, OutlookItemImportance):
            raise TypeError("importance must be of type OutlookItemImportance")
        self._internal_item.Importance = importance.value
    @property
    def received(self) -> datetime:
        return self._internal_item.ReceivedTime
    @received.setter
    def received(self, received: datetime):
        self._internal_item.ReceivedTime = received
    @property
    def body_format(self) -> str:
        return OutlookItemBodyFormat(self._internal_item.BodyFormat).name
    @body_format.setter
    def body_format(self, body_format: str):
        for possible_format in OutlookItemBodyFormat:
            if possible_format.name == body_format.lower():
                self._internal_item.BodyFormat = possible_format.value
                return
        raise ValueError(f"Body Format ({body_format}) is not a valid format for the body of this item.")
    def __repr__(self):
        return f"{self.__class__.__name__}({self.subject})"
    def __hash__(self):
        return hash(self._local_id)

class OutlookMailItem(OutlookItem):
    """https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.mailitem?view=outlook-pia"""
    @property
    def alternate_recipient_allowed(self) -> bool:
        return self._internal_item.AlternateRecipientAllowed
    @alternate_recipient_allowed.setter
    def alternate_recipient_allowed(self, alternate_allowed: bool):
        self._internal_item.AlternateRecipientAllowed = alternate_allowed
    

class OutlookAppointmentItem(OutlookItem):
    """https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.appointmentitem?view=outlook-pia"""
    @property
    def show_as(self) -> OutlookShowAs:
        return OutlookShowAs(self._internal_item.BusyStatus)
    @property
    def show_as(self) -> str:
        return OutlookShowAs(self._internal_item.BusyStatus).name
    @show_as.setter
    def show_as(self, busy_status: OutlookShowAs):
        self._internal_item.BusyStatus = busy_status.value
    @show_as.setter
    def show_as(self, busy_status: str):
        self._internal_item.BusyStatus = OutlookShowAs[busy_status.upper()].value

class OutlookReportItem(OutlookItem):
    """https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook._reportitem?view=outlook-pia
    usually a non-delivery report
    I can't find any special properties or members that seem to apply only to reports
    """

class OutlookMeetingItem(OutlookItem):
    """https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.meetingitem?view=outlook-pia"""
    MeetingResponse = Tuple[str, OutlookResponse]
    MeetingResponses = List[MeetingResponse]
    @property
    def responses(self) -> MeetingResponses:
        if self._responses is not None:
            return self._responses
        responses = []
        for recipient in self._internal_item.Recipients:
            responses.append(recipient.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS), OutlookResponse(recipient.MeetingResponseStatus))
        self._responses = responses
        return responses

class OutlookJournalItem(OutlookItem):
    """https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.journalitem?view=outlook-pia"""
    @property
    def posted(self) -> bool:
        return self._internal_item.DocPosted
    @posted.setter
    def posted(self, posted: bool):
        self._internal_item.DocPosted = posted
    @property
    def printed(self) -> bool:
        return self._internal_item.DocPrinted
    @printed.setter
    def printed(self, printed: bool):
        self._internal_item.DocPrinted = printed
    @property
    def routed(self) -> bool:
        return self._internal_item.DocPosted
    @posted.setter
    def routed(self, routed: bool):
        self._internal_item.DocRouted = routed
    @property
    def saved(self) -> bool:
        return self._internal_item.DocPosted
    @saved.setter
    def saved(self, saved: bool):
        self._internal_item.DocSaved = saved
    @property
    def duration(self) -> int:
        """integer duration in minutes"""
        return self._internal_item.Duration
    @duration.setter
    def duration(self, duration: int):
        self._internal_item.Duration = duration
    @property
    def start(self) -> datetime:
        return self._internal_item.Start
    @start.setter
    def start(self, start: datetime):
        self._internal_item.Start = start
    @property
    def end(self) -> datetime:
        return self._internal_item.End
    @end.setter
    def end(self, end: datetime):
        self._internal_item.End = end
    

class OutlookTaskItem(OutlookItem):
    """https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.taskitem?view=outlook-pia"""
    @property
    def due(self) -> datetime:
        return self._internal_item.DueDate
    @due.setter
    def due(self, due: datetime):
        self._internal_item.DueDate = due
    @property
    def card_data(self) -> str:
        return self._internal_item.CardData
    @card_data.setter
    def card_data(self, data: str):
        self._internal_item.CardData = data
    @property
    def actual_work(self) -> int:
        return self._internal_item.ActualWork
    @actual_work.setter
    def actual_work(self, work: int):
        self._internal_item.ActualWork = work
    @property
    def complete(self) -> bool:
        return self._internal_item.Complete
    @complete.setter
    def complete(self, complete: bool):
        self._internal_item.Complete = complete
    @property
    def date_completed(self) -> datetime:
        return self._internal_item.DateCompleted
    @date_completed.setter
    def date_completed(self, date_completed: datetime):
        self._internal_item.DateCompleted = date_completed
    @property
    def conflict(self) -> bool:
        return self._internal_item.IsConflict
    @property
    def recurring(self) -> bool:
        return self._internal_item.IsRecurring
    @property
    def owner(self) -> str:
        return self._internal_item.Owner
    @owner.setter
    def owner(self, owner: str):
        self._internal_item.Owner = owner
    @property
    def response(self) -> str:
        return OutlookTaskResponse(self._internal_item.ResponseState).name
    @property
    def role(self) -> str:
        return self._internal_item.Role
    @role.setter
    def role(self, role: str):
        self._internal_item.Role = role
    @property
    def schedule_plus_priority(self) -> str:
        return self._internal_item.SchedulePlusPriority
    @schedule_plus_priority.setter
    def schedule_plus_priority(self, priority: str):
        self._internal_item.SchedulePlusPriority = priority
    @property
    def status(self) -> str:
        return OutlookTaskStatus(self._internal_item.Status).name
    @status.setter
    def status(self, status: str):
        self._internal_item.Status = OutlookTaskStatus[status.upper()].value
    @status.setter
    def status(self, status: OutlookTaskStatus):
        self._internal_item.Status = status.value
    @property
    def team(self) -> bool:
        return self._internal_item.TeamTask
    @team.setter
    def team(self, isteamtask: bool):
        self._internal_item.TeamTask = isteamtask
    @property
    def todo_ordinal(self) -> datetime:
        return self._internal_item.ToDoTaskOrdinal
    @todo_ordinal.setter
    def todo_ordinal(self, ordinal: datetime):
        self._internal_item.ToDoTaskOrdinal = ordinal
