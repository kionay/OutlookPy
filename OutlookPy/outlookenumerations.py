"""Enumerations mimiced from Outlook."""
from enum import Enum

""" May need divided into 
AppointmentItem,    26
JournalItem,        42
MailItem,           43
MeetingItem,        54,181,53,55,56,57
and TaskItem        48, (49,51,52,50)
"""
class OutlookItemType(Enum):
    """
    imitates OlObjectClass Enumeration
    win32com has constants, but they're not divided up by enumeration class.
    """
    MailItem = [43]
    MeetingRequest = [53]
    MeetingResponse = [55,56,57]
    MeetingNotice = [54,181]
    Appointment = [26]
    DistributionList = [69]
    Task = [48]

class OutlookItemImportance(Enum):
    HIGH = 2
    NORMAL = 1
    LOW = 0

class OutlookItemBodyFormat(Enum):
    unspecified = 0
    plain = 1
    rich_text = 2
    html = 3

class OutlookRecipientType(Enum):
    """https://docs.microsoft.com/en-us/office/vba/api/outlook.recipient.type"""
    mail_originator = 0
    mail_to = 1
    mail_cc = 2
    mail_bcc = 3
    meeting_organizer = 0
    meeting_required = 1
    meeting_optional = 2
    meeting_resource = 3
    journal_contact = 1
    task_update = 2
    task_final_status = 3

class OutlookTaskResponse(Enum):
    """https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.oltaskresponse"""
    simple = 0
    assigned = 1
    accepted = 2
    declined = 3

class OutlookTaskStatus(Enum):
    """https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.oltaskstatus"""
    not_started = 0
    in_progress = 1
    completed = 2
    waiting = 3
    deferred = 4

class OutlookResponse(Enum):
    """
    If a Recipient is a recipient of a meeting, their MeetingResponseStatus will use this.
    Recipient.MeetingResponseStatus property (Outlook) and AppointmentItem.ResponseStatus property (Outlook)."""
    none = 0            # The appointment does not require a response (This is used for non-meeting recipients)
    organized = 1       # The appointment is on the organizer's calendar, or the recipient is the organizer
    tentative = 2       # Tentatively accepted
    accepted = 3        # Accepted
    declined = 4        # Declined
    notresponded = 5    # Recipient has not responded