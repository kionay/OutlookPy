#https://stackoverflow.com/questions/23341675/passing-additional-arguments-to-python-callback-object-win32com-client-dispatch?noredirect=1&lq=1

from win32com.client import Dispatch
from win32com.client import gencache
from win32com.client import getevents
from win32com.client import EventsProxy
import pythoncom

def _event_setattr_(self, attr, val):
    try:
        # Does the COM object have an attribute of this name?
        self.__class__.__bases__[0].__setattr__(self, attr, val)
    except AttributeError:
        # Otherwise just stash it away in the instance.
        self.__dict__[attr] = val

def DispatchWithEvents(clsid, user_event_class, arguments):
    # Create/Get the object.
    disp = Dispatch(clsid)
    if not disp.__class__.__dict__.get("CLSID"): # Eeek - no makepy support - try and build it.
        try:
            ti = disp._oleobj_.GetTypeInfo()
            disp_clsid = ti.GetTypeAttr()[0]
            tlb, index = ti.GetContainingTypeLib()
            tla = tlb.GetLibAttr()
            gencache.EnsureModule(tla[0], tla[1], tla[3], tla[4], bValidateFile=0)
            # Get the class from the module.
            disp_class = gencache.GetClassForProgID(str(disp_clsid))
        except pythoncom.com_error:
            raise TypeError("This COM object can not automate the makepy process - please run makepy manually for this object")
    else:
        disp_class = disp.__class__
    # If the clsid was an object, get the clsid
    clsid = disp_class.CLSID
    # Create a new class that derives from 3 classes - the dispatch class, the event sink class and the user class.
    # XXX - we are still "classic style" classes in py2x, so we need can't yet
    # use 'type()' everywhere - revisit soon, as py2x will move to new-style too...
    try:
        from types import ClassType as new_type
    except ImportError:
        new_type = type # py3k
    events_class = getevents(clsid)
    if events_class is None:
        raise ValueError("This COM object does not support events.")
    result_class = new_type("COMEventClass", (disp_class, events_class, user_event_class), {"__setattr__" : _event_setattr_})
    instance = result_class(disp._oleobj_) # This only calls the first base class __init__.
    events_class.__init__(instance, instance)
    args = [instance] + arguments
    if hasattr(user_event_class, "__init__"):
        user_event_class.__init__(*args)
    return EventsProxy(instance)

def WithEvents(disp, user_event_class, arguments):
    disp = Dispatch(disp)
    if not disp.__class__.__dict__.get("CLSID"): # Eeek - no makepy support - try and build it.
        try:
            ti = disp._oleobj_.GetTypeInfo()
            disp_clsid = ti.GetTypeAttr()[0]
            tlb, index = ti.GetContainingTypeLib()
            tla = tlb.GetLibAttr()
            gencache.EnsureModule(tla[0], tla[1], tla[3], tla[4], bValidateFile=0)
            # Get the class from the module.
            disp_class = gencache.GetClassForProgID(str(disp_clsid))
        except pythoncom.com_error as e:
            print(e)
            raise TypeError("This COM object can not automate the makepy process - please run makepy manually for this object")
    else:
        disp_class = disp.__class__
    # Get the clsid
    clsid = disp_class.CLSID
    # Create a new class that derives from 2 classes - the event sink
    # class and the user class.
    try:
        from types import ClassType as new_type
    except ImportError:
        new_type = type # py3k
    events_class = getevents(clsid)
    if events_class is None:
        raise ValueError("This COM object does not support events.")
    result_class = new_type("COMEventClass", (events_class, user_event_class), {})
    instance = result_class(disp) # This only calls the first base class __init__.
    args = [instance] + arguments
    if hasattr(user_event_class, "__init__"):
        user_event_class.__init__(*args)
    return instance