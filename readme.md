# OutlookPy

## Requirements
- pypiwin32
- pywin32

These requirements are for interfacing with Microsoft Outlook's COM API.

## Features

### Attaching to outlook is as easy as instantiating an object.

```python
from OutlookPy import OutlookPy
my_outlook = OutlookPy()
```

### Outlook mail items are wrapped in a more pythonic object structure.

__The main OutlookPy object has built-in well-known folder names.__

```python
inbox_folder = my_outlook.inbox
unread_inbox_items = [item for item in inbox_folder.items if item.unread]
for junk_mail in my_outlook.junk.items:
    junk_mail.delete()
```

__Folders can contain subfolders as well as items.__

```python
from outlookfilter import OutlookItemImportance

root = my_outlook.root_folder
my_custom_folder = root.folders["My High Importance Items"]
for inbox_item in my_outlook.inbox.items:
    if inbox_item.importance == OutlookItemImportance.HIGH:
        inbox_item.move(my_custom_folder)
        print(f"moved {inbox_item.subject} to high importance folder")
```

__Mail items can fetch their attributes easily.__

```python
five_most_recent_inbox_items = sorted(my_outlook.inbox.items, 
                                      key=lambda i: i.received_datetime, 
                                      reverse=True)[:5]
for item in five_most_recent_inbox_items:
    print(item.sender)
    print(item.subject)
    print(item.received_datetime)
    print()
```

__Decorators are now used to define event handlers.__

__Messages received are events of the receiving folder.__

```python
@my_outlook.inbox.on_item_received()
def debug_handler(mail_item):
    print(f"""
Sender: {mail_item.sender}
Subject: {mail_item.subject}
""")
    return True
```

__Events as decorated are either ran manually or attached then activated.__

```python
# rules can be ran manually
outlook.inbox.dispatch_unread()
# or you can listen to events

# this requires setting up a client
inbox_client = outlook.inbox.dispatch_events()
# attaching hooks to the events
outlook.inbox.hook_events(inbox_client)
# and activating those hooks, listening on events and running any setup rules
outlook.listen_for_events()
# just keep in mind that the listening is a blocking operation
# a blocking operation means you will no longer be able to provide input to python
```

## API Object Details

### Currently available mail item attributes

__read__
 - **Get**/**Set** the read status of an outlook item.
 
__unread__\*
 - **Get**/**Set** the inverse of the read status of an outlook item.
 
__recipients__
 - **Get**/~~Set~~\*\* the list of SMTP addresses to recieve this outlook item.
 
__categories__
 - **Get**/~~Set~~\*\* the list of names outlook categories associated with this outlook item.
 
__sender__
 - **Get**/~~Set~~\*\* the name of a given outlook item's sender, if it can be obtained.
 
__body__
 - **Get**/~~Set~~\*\* the body of a given outlook item.
 
__subject__
 - **Get**/~~Set~~\*\* the subject line of a given outlook item.
 
__external__
 - **Get** to tell if the given outlook item came from within the same exchange group/forest.
 
__internal__\*
 - **Get** the inverse of external.
 
__type__
 - **Get** the class or category of outlook item. Currently this is an enumeration, but when this is split into subclasses it will provide a subclass type.
 
__received_datetime__
 - **Get**/**Set** the receieved datetime of an outlook item. The setter is only valid on drafts, which are themselves not fully implemented so I would advise pretending this was readonly.
 
__body_format__
 - **Get**/**Set** the format of a given outlook item. This can be used to change between richtext, and HTML formats.
 
__containing_folder__
 - **Get** the library folder object containing the given outlook item. Akin to a *Parent* attribute, as that is what is used internally.
 
__importance__
 - **Get**/**Set** the *importance* attribute on a given outlook item. Corresponds to an OutlookItemImportance enumeration of high/normal/low.

\*: QoL addition.

\*\*: Some setters, when harder to implement than their getters, will come later. Also, some attributes will be moved to subclasses in the future, and their respective getters/setters will change in response. Some setters will only apply to item drafts yet to be sent. As odd as it is, some attributes are read only some of the time, and read/write some of the time.


### Currently available mailitem methods

__delete__
 - **Void** - Moves the outlook item to the Deleted Items folder, does not permanently delete unless it's already in that folder.
 
__move__
 - **Folder** - Moves the outlook item to the specified folder. Uses this library's folder object, so we can pretend that MAPIFolders don't exist.



### TODO:
- mailitem properties
    - [x] recieved datetime (from ReceivedTime)
    - [x] body format (richtext vs html)
    - [ ] CC/BCC recipients
    - [ ] Conversation
        - [ ] Conversation Topic
        - [ ] Conversation Index
    - [ ] Creation time
    - [ ] Read Receipt Requested
    - [ ] Reminder
        - [ ] Reminder Set
        - [ ] Reminder Time
    - and more (https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.mailitem?view=outlook-pia)
    