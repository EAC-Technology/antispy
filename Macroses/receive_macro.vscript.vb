' recive_macro

use eapp_xml
use antispy_lib

logger "Loading Antispy message"
args = event.data
token = args("additional")("eac_token") 

Try

  db_values = asjson(dbdictionary( token ))
  
  logger "EXISTS : " & token
  logger "db == " & dbdictionary( token )
  msg_time = db_values(1)
  msg_status = db_values(2)
  
  If msg_status = "closed" then
    msg_time = DateAdd("s", msg_time, Now)
    db_values(1) = msg_time
    db_values(2) = "opened"
    dbdictionary( token ) = tojson( db_values )
  End if
  
  msg_time = DateDiff("s", Now, TimeValue(msg_time))
  
  If msg_time > 0 then
    msg_secured = db_values(0)
  Else 
    msg_secured = "<i>Message deleted</i>"
    msg_time = 0
  End if
  
Catch
  msg_secured = "<i>Message deleted</i>"
  msg_time = 0
End try

Set e = new EAC 

xml_data = vdomxml_secured
xml_data = Replace(xml_data, "%img_spy%", resources.public_link("Icon.png"))
xml_data = Replace(xml_data, "%img_spy_bg%", resources.public_link("yelbg.png"))
xml_data = Replace(xml_data, "%SECURED_MSG%", msg_secured) ' embed msg 
xml_data = Replace(xml_data, "%remaining_time%", ConvertTime(msg_time)) ' embed msg 
xml_data = Replace(xml_data, "%remaining_time_ms%", msg_time) ' embed msg 

e.vdomxml_data = xml_data

e.dynamic = true               ' true/false whether it's dynamic content (boolean)
e.auth = "internal"             ' internal/external authentication (string)
e.session_token = ""         ' session token (string)

e.api_server = "https://" & Appinmail.utils.currentHost()      ' URL of backend server to handle EAC business
e.app_id = "7f459762-e1ba-42d3-a0e1-e74beda2eb85"            ' ID of the VDOM application running the business
e.eac_token = token
e.eac_method = "update"

response.write(e.get_wholexml())