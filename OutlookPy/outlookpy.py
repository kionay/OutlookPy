"""
Class definition for OutlookPy
"""
import ctypes
import pythoncom
import win32com.client
from win32com.client import constants, DispatchBaseClass
from outlookpy.outlookfolder import OutlookFolder
from outlookpy.constants import *


class OutlookPy():
    """
    The master object for interacting with outlook.
    """
    def __init__(self):
        print("attaching to application")
        generate_cache = win32com.client.gencache.EnsureDispatch("Outlook.Application")
        self._outlook_application = win32com.client.DispatchEx("Outlook.Application")
        self._mapi_namespace = self._outlook_application.GetNamespace("MAPI")
        self._outlook_session = self._outlook_application.Session
        self._my_smtp_address = self.__get_property(self._outlook_session, PR_SMTP_ADDRESS)
        self._root_folder = OutlookFolder(self._outlook_session.Folders[self._my_smtp_address])
        print("application attached")
    def __get_property(self, session, property_string):
        return session.CurrentUser.PropertyAccessor.GetProperty(property_string)
    @property
    def root_folder(self):
        return self._root_folder
    @property
    def root(self):
        return self._root_folder
    @property
    def inbox(self):
        return self._root_folder.folders["Inbox"]
    @property
    def drafts(self):
        return self._root_folder.folders["Drafts"]
    @property
    def sent(self):
        return self._root_folder.folders["Sent Items"]
    @property
    def deleted(self):
        return self._root_folder.folders["Deleted Items"]
    @property
    def journal(self):
        return self._root_folder.folders["Journal"]
    @property
    def outbox(self):
        return self._root_folder.folders["Outbox"]
    @property
    def junk(self):
        return self._root_folder.folders["Junk Email"]
    @property
    def calendar(self):
        return self._root_folder.folders["Calendar"]
    def listen_for_events(self):
        # pumping messages will cause this python thread's event loop
        #  to listen to messages sent to outlook
        #  this is a blocking operation, so we won't have control over this
        #  python thread unless an error triggers PostQuitMessage (WM_QUIT)
        pythoncom.PumpMessages()