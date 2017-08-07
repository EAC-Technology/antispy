' msg_delete


logger "Macro Delete Antispy run by timer"

Set mailbox = ProMail.get_mailbox("@AppInMail")


For Each eac_token in dbdictionary.keys
  val = asjson(dbdictionary(eac_token))  
  
  if UBound(val) = 5 then 
    
    if val(2) = "opened" then 
      if DateDiff("s", Now, TimeValue(val(1))) < 0 then
        logger "EAC is removing = " & tojson(val)
                
        AIM_user = Appinmail.users.resolve(val(5))
        proadmin_users = ProAdmin.users(AIM_user("login"))
        ProAdmin.set_user(proadmin_users(0))
                
        Set e = new EAC
        e.eac_token = eac_token
        e.eac_method = "delete"
        e.dynamic = false
        'e.api_server = "https://" & Appinmail.utils.currentHost()
        e.app_id = "7f459762-e1ba-42d3-a0e1-e74beda2eb85"
        
        logger e.send(mailbox, val(3), "", "", "Secured message", "Time is up. Message deleted.")
        logger e.send(mailbox, val(4), "", "", "Secured message", "Time is up. Message deleted.")
        
        dbdictionary.remove(eac_token)      
        
      end if
    end if
  Else
      dbdictionary.remove(eac_token)
  end if
Next

logger "Macro Delete Antispy finised"