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
unread_inbox_items = [item for item in inbox_folder if item.unread]
for junk_mail in my_outlook.junk:
    junk_mail.delete()
```

__Folders are iterable, check types to use properties unique to a type of item.__

```python
from outlookpy.outlookitem import OutlookMeetingItem
from outlookpy.enumerables import OutlookResponse

calendar_meetings = [meeting for meeting in my_outlook.calendar if type(meeting) is OutlookMeetingItem]

string_mapping = {
    OutlookResponse.organized: "{email} organized this meeting",
    OutlookResponse.tentative: "{email} is tentative",
    OutlookResponse.accepted: "{email} accepted this meeting",
    OutlookResponse.declined: "{email} declined this meeting",
    OutlookResponse.notresponded: "{email} has not responded"

}

for meeting in calendar_meetings:
    print(f"The meeting '{meeting.subject}' has the following responses:")
    for email, response in meeting.responses:
        if response != OutlookResponse.none:
            print(string_mapping[response].format(email=email))

```

__Folders can contain subfolders as well as items.__

```python
from outlookfilter import OutlookItemImportance

root = my_outlook.root_folder
my_custom_folder = root.folders["My High Importance Items"]
for inbox_item in my_outlook.inbox:
    if inbox_item.importance == OutlookItemImportance.HIGH:
        inbox_item.move(my_custom_folder)
        print(f"moved {inbox_item.subject} to high importance folder")
```

__Mail items can fetch their attributes easily.__

```python
five_most_recent_inbox_items = sorted(my_outlook.inbox, 
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
    # type check required, tasks might not have a sender
    if type(mail_item) is not OutlookTaskItem:
        print(f"Sender: {mail_item.sender}")
    # type check not required, all mail sub classes have a subject
    print(f"Subject: {mail_item.subject}")
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


### TODO items will be moved to Tasks/Tickets in github
    