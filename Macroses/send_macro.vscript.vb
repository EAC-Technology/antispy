' send_macro

logger "Send macro running"

use eapp_xml
use eapp_html
use lib_utils

args = event.data

sender_guid = args("required")("user.guid")
sender_user = Appinmail.users.resolve(sender_guid)("email")

Function ReplaceAndSplit(Text, DelimChars)
    strTemp = Text
    Delim1 = Left(DelimChars, 1)
    DelimLen = Len(DelimChars)
    For Delim = 2 To DelimLen
        ThisDelim = Mid(DelimChars, Delim, 1)
        If InStr(strTemp, ThisDelim) <> 0 Then _
            strTemp = Replace(strTemp, ThisDelim, Delim1)
    Next
    
    ReplaceAndSplit = Split(strTemp, Delim1)    
End Function


Function get_valid_emails(emails_string)
    emails_array = ReplaceAndSplit(emails_string, ", ;")
    
    Set r = new RegExp
    r.IgnoreCase = True
    r.Pattern = "^(([^<>()\[\]\.,;:\s@""]+(\.[^<>()\[\]\.,;:\s@""]+)*)|("".+""))@(([^<>()[\]\.,;:\s@""]+\.)+[^<>()[\]\.,;:\s@""]{2,})$"
    
    valid_emails = Array
    
    For each email in emails_array    
        If r.Test(email) then            
            AppendToArray(valid_emails, email)
        End if    
    Next
    get_valid_emails = valid_emails
End Function


Function send_email_to_sender(e, mailbox, email_list)    
    
    e.dynamic = true               ' true/false whether it's dynamic content (boolean)
    e.auth = "external"             ' internal/external authentication (string)
    e.session_token = ""         ' session token (string)
    
    e.login_container = "5073ff75-da99-44fb-a5d7-e44e5ab28598"   ' ID of the container where the login
    e.login_method = "login"     ' name of the login method (string)
 
    e.get_container = "5073ff75-da99-44fb-a5d7-e44e5ab28598"     ' ID of the container where to call the method
    e.get_method = "call_macro"         ' name of the method to get data (string)
    
    
    e.api_server = "https://" & Appinmail.utils.currentHost()      ' URL of backend server to handle EAC business
                                ' logic (string)
    e.app_id = "7f459762-e1ba-42d3-a0e1-e74beda2eb85"            ' ID of the VDOM application running the business
    
    e.eac_method = "new"            ' EAC method (string, new/update/delete)       
    
    lang = args("language")
    secured_msg = args("trigger")("params")("formtexteditor1") 
    
    sender_email = receiver_email ' 
    
    sender_time = args("trigger")("params")("formtext2")
    sender_time_units = args("trigger")("params")("formlist1")
    
    
    if sender_time = "" Then
      sender_time = 0
    Else
      sender_time = CInt(sender_time)
    End If
    
    if sender_time > 0 then
      Select Case sender_time_units
        Case "m"
          sender_time = sender_time * 60 
        Case "h"
          sender_time = sender_time * 3600 
        Case "d"
          sender_time = sender_time * 3600 * 24
'        Case Else
'          sender_time = sender_time
      End Select
    End if
    
    data = Dictionary
    i1 = Dictionary
    i1("ctype") = "image/png"
    i1("content") = resources.open("Icon.png").getbinary
    i1("cid") = "<icon@appinmail.io>"
    i1("inline") = True
    data("Icon.png") = i1
    i2 = Dictionary
    i2("ctype") = "image/png"
    i2("content") = resources.open("incoming_en.png").getbinary
    i2("cid") = "<incomingen@appinmail.io>"
    i2("inline") = True
    data("incoming_en.png") = i2
    i3 = Dictionary
    i3("ctype") = "image/png"
    i3("content") = resources.open("incoming_fr.png").getbinary
    i3("cid") = "<incomingfr@appinmail.io>"
    i3("inline") = True
    data("incoming_fr.png") = i3
    i4 = Dictionary
    i4("ctype") = "image/png"
    i4("content") = resources.open("lock.png").getbinary
    i4("cid") = "<lock@appinmail.io>"
    i4("inline") = True
    data("lock.png") = i4
    i5 = Dictionary
    i5("ctype") = "image/png"
    i5("content") = resources.open("Appinmail.png").getbinary
    i5("cid") = "<appinmail@appinmail.io>"
    i5("inline") = True
    data("Appinmail.png") = i5
    i6 = Dictionary
    i6("ctype") = "image/png"
    i6("content") = resources.open("vDOM.png").getbinary
    i6("cid") = "<vdom@appinmail.io>"
    i6("inline") = True
    data("vDOM.png") = i6
    i7 = Dictionary
    i7("ctype") = "image/png"
    i7("content") = resources.open("chrome.png").getbinary
    i7("cid") = "<chrome@appinmail.io>"
    i7("inline") = True
    data("chrome.png") = i7
    i8 = Dictionary
    i8("ctype") = "image/png"
    i8("content") = resources.open("anty-spymessage.png").getbinary
    i8("cid") = "<spymessage@appinmail.io>"
    i8("inline") = True
    data("anty-spymessage.png") = i8
    i9 = Dictionary
    i9("ctype") = "image/png"
    i9("content") = resources.open("yelbg.png").getbinary
    i9("cid") = "<yelbg@appinmail.io>"
    i9("inline") = True
    data("yelbg.png") = i9
    
    Select Case lang
      Case "fr"
        email_html = locale_fr(regular_msg, sender_user)
    Case Else
        email_html = locale_en(regular_msg, sender_user)
    End Select
    
    
    email_html = Replace(email_html, "%plugin_link%", "https://chrome.google.com/webstore/detail/eac-plugin/jhpamlhmchhnapkegepdmkffplafgnla")
    email_html = Replace(email_html, "%promail_link%", "https://promail.appinmail.io")
    
    
    For each email in email_list
        
        new_eac_token = generateguid()
        
        logger "sending msg to " & email & "eac token " & new_eac_token
        
        e.eac_token = new_eac_token
        e.get_data = "{""plugin_guid"": ""c0794b5f-35f5-4da8-a95a-d4c4b85e016b"", ""async"": 0, ""data"": {""eac_token"": """ + new_eac_token +""", ""type"" : ""antispy""}, ""name"": ""load_secured""}"
        
        Appinmail.acl.registerEac(new_eac_token)
        
        Appinmail.acl.add(new_eac_token, email, "w")   
        user = Appinmail.users.resolve(email)
        
        eacviewer_url = e.get_eacviewer_url(Appinmail.utils.currentHost(), user("email"))
        short_url = Appinmail.createShortUrl(eacviewer_url)
          
        dbdictionary( new_eac_token ) = tojson(Array(secured_msg, sender_time, "closed", user("login"), user("email"), sender_guid))
        
        email_html_filled = Replace(email_html, "%browser_link%", short_url)
    
        try
            e.vdomxml_data = Replace(vdomxml_secured_preview_xml_updated, "%spy_img_bg%", resources.public_link("msgbg.png"))
            e.send(mailbox, user("email"), "", "", "Confidential Email", email_html_filled, data, sender_user)
            e.send(mailbox, user("login"), "", "", "Confidential Email", "Use promail or link " + short_url, Dictionary, sender_user)                        
        catch
          logger "Antispy email sending failed to " & email
        end try
    Next    
          
End function 

emails = args("trigger")("params")("formtext1")
emails = get_valid_emails(emails)
If len(emails) > 0 then
  Set eac_new = new EAC 
  Call send_email_to_sender(eac_new, ProMail.get_mailbox("@AppInMail"), emails)    
End if

eac_token = args("additional")("eac_token")

Set e = new EAC  

xml_data = vdomxml_msg
xml_data = Replace(xml_data, "%destination%", "destination")
xml_data = Replace(xml_data, "%delete_time%", "delete after")
xml_data = Replace(xml_data, "%send%", "send")
xml_data = Replace(xml_data, "%img_spy%", resources.public_link("Icon.png"))
xml_data = Replace(xml_data, "%img_spy_bg%", resources.public_link("yelbg.png"))
xml_data = Replace(xml_data, "%img_spy_bg_sent%", resources.public_link("msgbg.png"))


If len(emails) > 0 then    
    emails_string = Join(emails, ", ")
    emails_cut = SimpleIf((len(emails_string)>70), Left(emails_string, 70)&"...", emails_string)
    vdomxml_msg_sent_to_container_emails_filled = Replace(vdomxml_msg_sent_to_container, "%emails_cut%", emails_cut)
    vdomxml_msg_sent_to_container_all_filled = Replace(vdomxml_msg_sent_to_container_emails_filled, "%emails%", emails_string)
    vdomxml_msg_sent_container_filled = Replace(vdomxml_msg_sent_container, "%result_container%", vdomxml_msg_sent_to_container_all_filled)
    
Else
    vdomxml_msg_sent_container_filled = Replace(vdomxml_msg_sent_container, "%result_container%", vdomxml_msg_no_valid_emails)
End if



xml_data = Replace(xml_data, "%container_msg_sent%", vdomxml_msg_sent_container_filled)

e.vdomxml_data = xml_data

e.dynamic = true               ' true/false whether it's dynamic content (boolean)
e.auth = "internal"             ' internal/external authentication (string)
e.session_token = ""         ' session token (string)

e.login_container = "5073ff75-da99-44fb-a5d7-e44e5ab28598"   ' ID of the container where the login
e.login_method = "login"     ' name of the login method (string)

e.post_container = "5073ff75-da99-44fb-a5d7-e44e5ab28598"    ' ID of the container where to call the method
e.post_method = "call_macro"       ' name of the method to post data (string)
e.post_data = "{""plugin_guid"": ""c0794b5f-35f5-4da8-a95a-d4c4b85e016b"", ""async"": 0, ""data"": {""eac_token"": """+ eac_token +""", ""type"" : ""antispy""}, ""name"": ""sendeapp""}"        ' data to be sent as pattern for POST
                            ' request (string)
                            
e.api_server = "https://" & Appinmail.utils.currentHost()      ' URL of backend server to handle EAC business
                            ' logic (string)
e.app_id = "7f459762-e1ba-42d3-a0e1-e74beda2eb85"            ' ID of the VDOM application running the business
                            ' logic (string)
events = Dictionary    
events("container1.form1:submit") = Array("form_send")
events("container1.container2.formbutton1:click") = Array(Array("container1.container2:hide", "1000"))
e.events_data = tojson(events)        ' JSON structure that 
e.eac_token = eac_token
e.eac_method = "update"

response.write(e.get_wholexml())