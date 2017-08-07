'--- General libs


'#include(lib_utils) 
'#include(lib_dialog_2) 


use eapp_xml
use eapp_html

'Initial SetUp
args = xml_dialog.get_answer
logger ("Data sent :" & tojson(args))
logger Appinmail.utils.parseEmails(args)


set current_user = ProAdmin.current_user
sender_user = Appinmail.users.resolve(current_user.guid)("email")


if "step" in args then
  if args( "step" )="" then
    operation_step ="Init"
  else
    operation_step = args( "step" )
  end if
else
  operation_step ="Init"
end if


'###############################################################################################
'#                                Screen Definition                                            #
'###############################################################################################
'VDOMData = DynamicVDOM.render(vdomxml) 
set MyForms = New XMLDialogBuilder

'-------------------------------------- VDOM Screen TEST ---------------------------------------
' TEST to include VDOM XML in XML Dialog
'-----------------------------------------------------------------------------------------------    
    MyForms.addScreen("Show VDOM")
    MyForms.Screen("Show VDOM").addComponent("VDOMXML","hypertext")
    
    vdomxml =  "<TEXT name='test' value='test'/>"
   
    DynamicVDOM.left="0"
    DynamicVDOM.top="0"
    DynamicVDOM.width="600"
    DynamicVDOM.width="300"
    'VDOMData = DynamicVDOM.render(vdomxml) 
    'MyForms.Screen("Show VDOM").Component("VDOMXML").value(VDOMData)

    'logger Replace(VDOMData, "<", "&lt;")
    
'-------------------------------------- Std Composer -------------------------------------------
' Main error Window
'-----------------------------------------------------------------------------------------------    
  MyForms.addScreen("Std Composer")
  MyForms.Screen("Std Composer").width  = 800
  MyForms.Screen("Std Composer").Height = 750
  
  MyForms.Screen("Std Composer").Title  = "Antispy mail composer"  
  
  MyForms.Screen("Std Composer").addComponent("toemail","livesearch")
  MyForms.Screen("Std Composer").Component("toemail").label("To :")
  
  'MyForms.Screen("Std Composer").addComponent("delete_after","TextBox")
  'MyForms.Screen("Std Composer").Component("delete_after").label("Delete after :")
  'MyForms.Screen("Std Composer").addComponent("units", "DropDown")
   
  
  MyForms.Screen("Std Composer").addComponent("message","RichTextArea")
  MyForms.Screen("Std Composer").Component("message").label("Message :")
  
  MyForms.Screen("Std Composer").addComponent("btns","btngroup")
  MyForms.Screen("Std Composer").Component("btns").addBtn("Send","sendEmail")
  MyForms.Screen("Std Composer").Component("btns").addBtn("Cancel","Exit")

'-------------------------------------- Msg Email sent -----------------------------------------
' Msg box to show the email was sent
'-----------------------------------------------------------------------------------------------
    MyForms.addScreen("Email Sent Msg")
    MyForms.Screen("Email Sent Msg").addComponent("comment","Text")
    MyForms.Screen("Email Sent Msg").Component("comment").setCenter(true)
    MyForms.Screen("Email Sent Msg").Component("comment").value("<br>Your email was correctly sent.")
    MyForms.Screen("Email Sent Msg").Title = "Sending email ..."
    MyForms.Screen("Email Sent Msg").addComponent("timeout","Timer")
    MyForms.Screen("Email Sent Msg").Component("timeout").setTimer("Exit",1000)
    

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
      
      lang = "en"
      secured_msg = args("message")
      
      sender_email = receiver_email ' 
      
      'sender_time = args("trigger")("params")("formtext2")
      'sender_time_units = args("trigger")("params")("formlist1")
      
      
      'if sender_time = "" Then
      '  sender_time = 0
      'Else
      '  sender_time = CInt(sender_time)
      'End If
      
      'if sender_time > 0 then
      '  Select Case sender_time_units
      '    Case "m"
      '      sender_time = sender_time * 60 
      '    Case "h"
      '      sender_time = sender_time * 3600 
      '    Case "d"
      '      sender_time = sender_time * 3600 * 24
  '        Case Else
  '          sender_time = sender_time
       ' End Select
      'End if
      sender_time = 3600
      
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


    
Function tt(mailbox)
    

    toEmails = Appinmail.utils.parseEmails(args)
    Set eac_new = new EAC 
    Call send_email_to_sender(eac_new, ProMail.get_mailbox("@AppInMail"), toEmails)
    
end function 
'############################################################################################
'#                                                                                          #
'#                         State machine start with state / Init /                          #
'#                                                                                          #
'############################################################################################

if instr(operation_step,">")<>0 then
  mainStep = split(operation_step,">")(0)
  subStep = split(operation_step,">")(1)
else
  mainStep = operation_step
  subStep = ""
end if

logger ("operation step :" & operation_step)

Select case mainStep

  case "Init" 
    MyForms.ShowScreen("Std Composer")

  case "sendEmail"
    

    Call tt(ProMail.selected_mailbox)
  
    MyForms.ShowScreen("Email Sent Msg")  

  case "Exit"
    logger ("End of Processing Rules")

end select