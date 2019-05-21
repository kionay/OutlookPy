from outlookpy import OutlookPy

my_outlook = OutlookPy()
print("Initializing handler functions")

@my_outlook.inbox.on_item_received()
def debug_handler(mail_item):
    print(f"""
Sender: {mail_item.sender}
Subject: {mail_item.subject}
""")
    return True

@my_outlook.deleted.on_item_received()
def deleted_test(mail_item):
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
