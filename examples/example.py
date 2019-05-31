from outlookpy import OutlookPy
from outlookpy.outlookitem import *

my_outlook = OutlookPy()
print("Initializing handler functions")

@my_outlook.inbox.on_item_received()
def debug_handler(mail_item):
    # type check required, tasks might not have a sender
    if mail_item is not OutlookTaskItem:
        print(f"Sender: {mail_item.sender}")
    # type check not required, all mail sub classes have a subject
    print(f"Subject: {mail_item.subject}")
    return True

@my_outlook.deleted.on_item_received()
def deleted_test(mail_item):
    # type check not required, all mail sub classes have a subject
    print(f"mail item {mail_item.subject} deleted")
    

my_outlook.inbox.dispatch_unread()

print("dispatching inbox")
inbox_client = my_outlook.inbox.dispatch_events()
my_outlook.inbox.hook_events(inbox_client)


print("dispatching deleted")
deleted_client = my_outlook.deleted.dispatch_events()
my_outlook.deleted.hook_events(deleted_client)


print("listening for dispatched events")
my_outlook.listen_for_events()
