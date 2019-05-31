"""Helper functions.
    Contains functions that are not logically bound to any particular class.
"""
from enum import Enum

from .outlookenumerations import *
from .outlookitem import *

CLASS_LOOKUP = {
    26 : OutlookAppointmentItem,
    42 : OutlookJournalItem,
    43 : OutlookMailItem,
    48 : OutlookTaskItem,
    49 : OutlookTaskItem,
    50 : OutlookTaskItem,
    51 : OutlookTaskItem,
    52 : OutlookTaskItem,
    53 : OutlookMeetingItem,
    54 : OutlookMeetingItem,
    55 : OutlookMeetingItem,
    56 : OutlookMeetingItem,
    57 : OutlookMeetingItem,
    162 : OutlookJournalItem,
    181 : OutlookMeetingItem
}


def com_to_python(COMObject):
    """Wraps the COM object in its associated python object.
        If the _ItemClass enumeration has the COM object's class number in it, we can wrap it with a class-specific object.
        If not, no wrapper has yet been written, so we attempt to fall back to the generic OutlookItem superclass.
        If even that does not work, the object is truly niche and I don't feel bad about not having gotten around to it yet.
    """
    if COMObject.Class in CLASS_LOOKUP:
        return CLASS_LOOKUP[COMObject.Class](COMObject)
    else:
        print(f"WARNING - Item Class {COMObject.Class} not a designated wrappable object.")
        return OutlookItem(COMObject)