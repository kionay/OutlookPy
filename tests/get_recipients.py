from outlookpy import OutlookPy
from outlookpy.outlookitem import *

outlook = OutlookPy()

for folder in outlook.root_folder.folders:
    for item in folder:
        assert all(recipient.name for recipient in item.recipients)