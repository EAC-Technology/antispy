' antispy_lib

use eapp_xml

Function init_ui_eapp(e, mailbox, token, send_email)               
    
    myguid = token 
    host = Appinmail.utils.currentHost()
    
    xml_data = vdomxml_msg
    xml_data = Replace(xml_data, "%destination%", "destination")
    xml_data = Replace(xml_data, "%delete_time%", "delete after")
    xml_data = Replace(xml_data, "%send%", "send")
    xml_data = Replace(xml_data, "%img_spy%", resources.public_link("Icon.png"))
    'xml_data = Replace(xml_data, "%img_spy%", "https://" & host & "data:image/png;base64," & resources.open("Icon.png").getvalue)
    xml_data = Replace(xml_data, "%img_spy_bg%", resources.public_link("yelbg.png"))
    ' Replace(resources.open("yelbg_.b64").getvalue, "1", "\\"&vbCr))
    xml_data = Replace(xml_data, "%img_spy_bg_sent%", "")
    
    xml_data = Replace(xml_data, "%container_msg_sent%", "")    
    
    e.vdomxml_data = xml_data  ' initial ui
    e.dynamic = false               ' true/false whether it's dynamic content (boolean)
    e.auth = "internal"             ' internal/external authentication (string)
    e.session_token = ""         ' session token (string)
    e.login_container = "5073ff75-da99-44fb-a5d7-e44e5ab28598"   ' ID of the container where the login
                                ' method is (string)
    e.login_method = "login"     ' name of the login method (string)
    
    e.post_container = "5073ff75-da99-44fb-a5d7-e44e5ab28598"    ' ID of the container where to call the method
                                ' to post data (string)
    e.post_method = "call_macro"       ' name of the method to post data (string)
    e.post_data = "{""plugin_guid"": ""c0794b5f-35f5-4da8-a95a-d4c4b85e016b"", ""async"": 0, ""data"": {""eac_token"": """+ myguid +""", ""type"" : ""antispy""}, ""name"": ""sendeapp""}"        ' data to be sent as pattern for POST
                                ' request (string)

    e.api_server = "https://" & Appinmail.utils.currentHost()      ' URL of backend server to handle EAC business
                                ' logic (string)
    e.app_id = "7f459762-e1ba-42d3-a0e1-e74beda2eb85"            ' ID of the VDOM application running the business
                                ' logic (string)
    events = Dictionary    
	events("container1.form1:submit") = Array("form_send")
    e.events_data = tojson(events)        ' JSON structure that represents events (string)
    
    e.eac_token = myguid     ' EAC token (string)
    e.eac_method = "update"            ' EAC method (string, new/update/delete)
    
    Appinmail.acl.registerEac(myguid)

    Set u = ProAdmin.current_user
    user = Appinmail.users.resolve(u.email)
    Appinmail.acl.add(myguid, Array(user("email")), "r")   
    
    eacviewer_url = e.get_eacviewer_url(Appinmail.utils.currentHost(), user("email"))
    short_url = Appinmail.createShortUrl(eacviewer_url)
    
    logger "NOTSHORT_LINK: " & eacviewer_url
    
    If send_email = True Then
      try
        e.send(mailbox, user("login"), "", "", "AntiSPY form", "Use promail or link " + short_url)
      catch
        logger "ERROR !!!"
      end try
    'Else
    '  init_ui_eapp = e
    End If
                    
end function


' delete secured msg

'Sub delete_secured_msg(dbkey)
'   set t = GetTimer("msg_destruct")
'   t.delay = "00:00:00:10"
'   ActivateTimer(t.name)

'end Sub


Function ConvertTime(intTotalSecs)
  Dim intHours,intMinutes,intSeconds,Time
  intHours = intTotalSecs \ 3600
  intMinutes = (intTotalSecs Mod 3600) \ 60
  intSeconds = intTotalSecs Mod 60
  ConvertTime = LPad(intHours) & ":" & LPad(intMinutes) & ":" & LPad(intSeconds)
End Function


Function LPad(v) 
   LPad = Right("0" & v, 2) 
End Function