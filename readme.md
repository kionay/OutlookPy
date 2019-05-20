# OutlookPy

## Requirements
- pypiwin32
- pywin32

These requirements are for interfacing with Microsoft Outlook's COM API.

## Features

### Attaching to outlook is as easy as instantiating an object.

```python
from outlookpy import OutlookPy
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

### Documentation will be created in a /docs/ folder, instead of in the readme.


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
    