"""Enumerations mimiced from Outlook."""
from enum import Enum

class OutlookItemImportance(Enum):
    HIGH = 2
    NORMAL = 1
    LOW = 0

class OutlookItemBodyFormat(Enum):
    UNSPECIFIED = 0
    PLAIN = 1
    RICH_TEXT = 2
    HTML = 3
    CLEAR_SIGNED = 4

class OutlookRecipientType(Enum):
    """https://docs.microsoft.com/en-us/office/vba/api/outlook.recipient.type"""
    MAIL_ORIGINATOR = 0
    MAIL_TO = 1
    MAIL_CC = 2
    MAIL_BCC = 3
    MEETING_ORGANIZER = 0
    MEETING_REQUIRED = 1
    MEETING_OPTIONAL = 2
    MEETING_RESOURCE = 3
    JOURNAL_CONTACT = 1
    TASK_UPDATE = 2
    TASK_FINAL_STATUS = 3

class OutlookTaskResponse(Enum):
    """https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.oltaskresponse"""
    SIMPLE = 0
    ASSIGNED = 1
    ACCEPTED = 2
    DECLINED = 3

class OutlookTaskStatus(Enum):
    """https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.oltaskstatus"""
    NOT_STARTED = 0
    IN_PROGRESS = 1
    COMPLETED = 2
    WAITING = 3
    DEFERRED = 4

class OutlookResponse(Enum):
    """
    If a Recipient is a recipient of a meeting, their MeetingResponseStatus will use this.
    Recipient.MeetingResponseStatus property (Outlook) and AppointmentItem.ResponseStatus property (Outlook).
    """
    NONE = 0            # The appointment does not require a response (This is used for non-meeting recipients)
    ORGANIZED = 1       # The appointment is on the organizer's calendar, or the recipient is the organizer
    TENTATIVE = 2       # Tentatively accepted
    ACCEPTED = 3        # Accepted
    DECLINED = 4        # Declined
    NOT_RESPONDED = 5    # Recipient has not responded

class OutlookShowAs(Enum):
    """https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.olbusystatus?view=outlook-pia"""
    FREE = 0
    TENTATIVE = 1
    BUSY = 2
    OUT_OF_OFFICE = 3
    WORKING_ELSEWHERE = 4
