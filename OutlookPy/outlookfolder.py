from __future__ import annotations
from typing import List, Dict
import ctypes
import win32com.client
from win32com.client import Dispatch
import pythoncom

from outlookpy.alternatedispatch import WithEvents
from outlookpy.outlookitem import OutlookItem, com_to_python

class OutlookFolder(list):
    """
    Wrapper class for outlook folders. MAPIFolder
    Acts like an iterable of OutlookItem objects
    """
    def __init__(self, folder):
        self._folder = folder
        # when events are dispatched, the event OnItemAdd is actually on the .Items object
        # so a folder is listened to for additions to it by its Items property
        # OnItemAdd is not an event on the folder, but for the wrapper i'm binding a folder and its items
        # so our folder wrapper needs to understand what that item being listened to is, so it can make the user
        # seem like its the folder that has the event, beause that makes more sense (ya here that, microsoft!?)
        self._MAPI_items = folder.Items # used for attaching events
        # sub folders are a dictionary, keys being folder names and values being the folder objects
        self._folders = {sub_folder.Name:OutlookFolder(sub_folder) for sub_folder in folder.Folders}
        self._attached_handlers = {"add":[],"remove":[],"change":[]}
        self._internal_proxy = None
    def __eq__(self, other):
        return self._local_id == other._local_id
    def __ne__(self, other):
        return not (self == other)
    def __iter__(self):
        for item in self._MAPI_items:
            yield com_to_python(item)
    def __getitem__(self, key):
        return com_to_python(list(self._MAPI_items)[key])
    def __len__(self):
        return len(self._MAPI_items)
    def __repr__(self):
        return f"{self.__class__.__name__}({self.name})"
    def __hash__(self):
        return hash(self._local_id)
    @property
    def _local_id(self):
        """the closest thing to a unqiue ID we have"""
        return self._folder.EntryID
    @property
    def name(self) -> str:
        """given or well-known folder name, only unque amongst its parent folder"""
        return self._folder.Name     
    def OnItemAdd(self, mail):
        """mandatory event, name is hard-wired for exchange API"""
        # wrap the mail item, then use it
        mail = com_to_python(mail)
        for handler in self._attached_handlers["add"]:
            try:
                result = handler(mail)
                if not result: # if the response is falsey
                    break # stop processing more rules/handlers
            except Exception as e:
                print(e)
                ctypes.windll.user32.PostQuitMessage(0)
    def OnItemRemove(self):
        for handler in self._attached_handlers["remove"]:
            try:
                result = handler()
                if not result:
                    break
            except Exception as e:
                print(e)
                ctypes.windll.user32.PostQuitMessage(0)
    def OnItemChange(self, mail):
        mail = com_to_python(mail)
        for handler in self._attached_handlers["change"]:
            try:
                result = handler(mail)
                if not result:
                    break
            except Exception as e:
                print(e)
                ctypes.windll.user32.PostQuitMessage(0)
    def on_item_added(self):
        return self.on_item_received(self, config)
    def on_item_received(self):
        def decorator(callback):
            self._attached_handlers["add"].append(callback)
            if self._internal_proxy is not None:
                self._internal_proxy._attached_handlers["add"].append(callback)
            return callback
        return decorator
    def on_item_removed(self):
        def decorator(callback):
            self._attached_handlers["remove"].append(callback)
            if self._internal_proxy is not None:
                self._internal_proxy._attached_handlers["remove"].append(callback)
            return callback
        return decorator
    def on_item_changed(self):
        def decorator(callback):
            self._attached_handlers["change"].append(callback)
            if self._internal_proxy is not None:
                self._internal_proxy._attached_handlers["change"].append(callback)
            return callback
        return decorator
    def dispatch_events(self):
        client = Dispatch(self._folder.Items)
        return client
    def hook_events(self, client):
        proxy = WithEvents( client, OutlookFolder, [self._folder])
        self._internal_proxy = proxy
        # the init-ed things will all be the same in the proxy object (application, namespace, session)
        # anything in this object that is modified on the fly needs to be mirrored in the proxy object
        self._internal_proxy._attached_handlers = self._attached_handlers
    def dispatch_unread(self):
        for mail_item in self:
            if not mail_item.read:
                result = None
                for handler in self._attached_handlers["add"]:
                    result = handler(mail_item)
                    if not result:
                        break
    @property
    def folders(self) -> Dict[str,OutlookFolder]:
        return self._folders