' init_macro

logger "Init macro running"
use antispy_lib

' probably need to store it somewhere
myguid = generateguid()

Set e = new EAC 
Call init_ui_eapp(e, ProMail.get_mailbox("@AppInMail"), myguid, False)

resp = Dictionary
resp("eac_token") = myguid
resp("whole_xml") = base64encode(e.get_wholexml())
response.write(tojson(resp))