' ====== FR ======= '
Function locale_fr(msg, username)
	description_str = "Ce message vous est envoyé par %USERNAME% et est protogé contre toute violation de la vie privée par la technologie EAC, pour pouvoir lire ce message vous devez activer la technologie EAC pour votre logiciel de messagerie."
	click_extension = "Installer l'extension EAC pour Chrome"
	click_promail = "Visualiser le message dans ProMail"
	click_browser = "Visualiser le message dans votre navigateur"
	footer = "© 2017 AppInMail S.A.S., Anti-Spy Plug-In pour ProMail. Les logo AppInMail et EAC sont des marques déposées par AppInMail S.A.S. Tous les droits sont réservés."

	descr = Replace(description_str, "%USERNAME%", username)
	msg = Replace(msg, "%description%", descr)
	msg = Replace(msg, "%click_extension%", click_extension)
	msg = Replace(msg, "%click_promail%", click_promail)
	msg = Replace(msg, "%click_browser%", click_browser)
	msg = Replace(msg, "%footer%", footer)
    
    msg = Replace(msg, "%middle_header_text_base64%", "cid:incomingfr@appinmail.io")'
    msg = Replace(msg, "%locker_icon%", "cid:lock@appinmail.io")
    msg = Replace(msg, "%eapp_icon%", "cid:appinmail@appinmail.io")
    msg = Replace(msg, "%promail_icon%", "cid:vdom@appinmail.io")
    msg = Replace(msg, "%browser_icon%", "cid:chrome@appinmail.io")
    msg = Replace(msg, "%spy_icon%", "cid:icon@appinmail.io")
    msg = Replace(msg, "%header_icon%", "cid:spymessage@appinmail.io")
    msg = Replace(msg, "%whole_background%", "cid:yelbg@appinmail.io")    

	locale_fr = msg
End function


' ====== EN ======= '
Function locale_en(msg, username)
	description_str = "This message from %USERNAME% is protected against privacy violation by EAC Technology in order to be able to read it, you must enable EAC Technology for your email client."
	click_extension = "Install EAC Plug-in form Chrome"
	click_promail = "See message in ProMail"
	click_browser = "See message in Web Browser"
	footer = "© 2017 AppInMail S.A.S., Anti-Spy Plug-In for ProMail. The AppInMail logo are trademark of AppInMail S.A.S. All rights reserved."

	descr = Replace(description_str, "%USERNAME%", username)
	msg = Replace(msg, "%description%", descr)
	msg = Replace(msg, "%click_extension%", click_extension)
	msg = Replace(msg, "%click_promail%", click_promail)
	msg = Replace(msg, "%click_browser%", click_browser)
	msg = Replace(msg, "%footer%", footer)

    msg = Replace(msg, "%middle_header_text_base64%", "cid:incomingen@appinmail.io")'
    msg = Replace(msg, "%locker_icon%", "cid:lock@appinmail.io")
    msg = Replace(msg, "%eapp_icon%", "cid:appinmail@appinmail.io")
    msg = Replace(msg, "%promail_icon%", "cid:vdom@appinmail.io")
    msg = Replace(msg, "%browser_icon%", "cid:chrome@appinmail.io")
    msg = Replace(msg, "%spy_icon%", "cid:icon@appinmail.io")
    msg = Replace(msg, "%header_icon%", "cid:spymessage@appinmail.io")
    msg = Replace(msg, "%whole_background%", "cid:yelbg@appinmail.io")
    
	locale_en = msg
    
End function


regular_msg_1 = ""&vbCrLf&_
"<style>"&vbCrLf&_
""&vbCrLf&_
"body {"&vbCrLf&_
"  font-family: Arial, ""Helvetica Neue"", Helvetica, sans-serif;"&vbCrLf&_
"  color: #333;"&vbCrLf&_
"  font-size: 14px;"&vbCrLf&_
"  line-height: 20px;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".section {"&vbCrLf&_
"  display: -webkit-box;"&vbCrLf&_
"  display: -webkit-flex;"&vbCrLf&_
"  display: -ms-flexbox;"&vbCrLf&_
"  display: flex;"&vbCrLf&_
"  height: 100%;"&vbCrLf&_
"  padding-bottom: 1px;"&vbCrLf&_
"  background-image: -webkit-linear-gradient(124deg, #000, #fff);"&vbCrLf&_
"  background-image: linear-gradient(326deg, #000, #fff);"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".div-block {"&vbCrLf&_
"  position: absolute;"&vbCrLf&_
"  left: 58px;"&vbCrLf&_
"  top: 98px;"&vbCrLf&_
"  right: 64px;"&vbCrLf&_
"  bottom: 62px;"&vbCrLf&_
"  display: inline-block;"&vbCrLf&_
"  margin-bottom: 57px;"&vbCrLf&_
"  padding-right: 103px;"&vbCrLf&_
"  float: none;"&vbCrLf&_
"  background-color: #136daa;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".body {"&vbCrLf&_
"  background-color: #fff;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".image {"&vbCrLf&_
"  position: relative;"&vbCrLf&_
"  left: 185px;"&vbCrLf&_
"  top: 217px;"&vbCrLf&_
"  opacity: 1;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".lightbox-link {"&vbCrLf&_
"  position: absolute;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".heading {"&vbCrLf&_
"  display: -webkit-box;"&vbCrLf&_
"  display: -webkit-flex;"&vbCrLf&_
"  display: -ms-flexbox;"&vbCrLf&_
"  display: flex;"&vbCrLf&_
"  margin-top: 50%;"&vbCrLf&_
"  margin-left: 1px;"&vbCrLf&_
"  padding-left: 0px;"&vbCrLf&_
"  -webkit-box-orient: vertical;"&vbCrLf&_
"  -webkit-box-direction: normal;"&vbCrLf&_
"  -webkit-flex-direction: column;"&vbCrLf&_
"  -ms-flex-direction: column;"&vbCrLf&_
"  flex-direction: column;"&vbCrLf&_
"  -webkit-justify-content: space-around;"&vbCrLf&_
"  -ms-flex-pack: distribute;"&vbCrLf&_
"  justify-content: space-around;"&vbCrLf&_
"  -webkit-box-align: start;"&vbCrLf&_
"  -webkit-align-items: flex-start;"&vbCrLf&_
"  -ms-flex-align: start;"&vbCrLf&_
"  align-items: flex-start;"&vbCrLf&_
"}"&vbCrLf

regular_msg_1_1 = ""&vbCrLf&_
".div-block-2 {"&vbCrLf&_
"  display: block;"&vbCrLf&_
"  padding-right: 4px;"&vbCrLf&_
"  padding-bottom: 258px;"&vbCrLf&_
"  background-color: #0098ff;"&vbCrLf&_
"  -webkit-transform: rotate(90deg);"&vbCrLf&_
"  -ms-transform: rotate(90deg);"&vbCrLf&_
"  transform: rotate(90deg);"&vbCrLf&_
"  -webkit-transition: border-radius 200ms cubic-bezier(.339, -.245, .626, 1.317);"&vbCrLf&_
"  transition: border-radius 200ms cubic-bezier(.339, -.245, .626, 1.317);"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".heading-2 {"&vbCrLf&_
"  margin-top: 31px;"&vbCrLf&_
"  margin-bottom: 25px;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".column {"&vbCrLf&_
"  position: static;"&vbCrLf&_
"  display: block;"&vbCrLf&_
"  overflow: visible;"&vbCrLf&_
"  height: 100px;"&vbCrLf&_
"  margin-top: 31px;"&vbCrLf&_
"  margin-bottom: -23px;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".row {"&vbCrLf&_
"  margin-bottom: -31px;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".button {"&vbCrLf&_
"  margin-top: 46px;"&vbCrLf&_
"  border: 0px solid transparent;"&vbCrLf&_
"  border-radius: 10px;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".paragraph {"&vbCrLf&_
"  margin-top: 72px;"&vbCrLf&_
"  padding-right: 46px;"&vbCrLf&_
"  padding-left: 15px;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".button-2 {"&vbCrLf&_
"  padding-top: 10px;"&vbCrLf&_
"  background-color: #fff;"&vbCrLf&_
"  color: #000;"&vbCrLf&_
"}"&vbCrLf

regular_msg_1_2 = ""&vbCrLf&_
".row-2 {"&vbCrLf&_
"  padding: 42px 42px 51px 41px;"&vbCrLf&_
"  background-color: #fd4c53;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".div-block-3 {"&vbCrLf&_
"  background-image: -webkit-linear-gradient(270deg, rgba(216, 76, 76, .5), rgba(216, 76, 76, .5));"&vbCrLf&_
"  background-image: linear-gradient(180deg, rgba(216, 76, 76, .5), rgba(216, 76, 76, .5));"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".rich-text-block {"&vbCrLf&_
"  display: block;"&vbCrLf&_
"  width: 70%;"&vbCrLf&_
"  margin-right: auto;"&vbCrLf&_
"  margin-left: auto;"&vbCrLf&_
"  padding-right: 117px;"&vbCrLf&_
"  padding-bottom: 1px;"&vbCrLf&_
"  padding-left: 0px;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".container {"&vbCrLf&_
"  background-color: #fff;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".div-block-4 {"&vbCrLf&_
"  height: 100vh;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".button-3 {"&vbCrLf&_
"  border-radius: 7px;"&vbCrLf&_
"  background-image: -webkit-radial-gradient(circle farthest-corner at 0% 0%, #501974, #da0750);"&vbCrLf&_
"  background-image: radial-gradient(circle farthest-corner at 0% 0%, #501974, #da0750);"&vbCrLf&_
"  opacity: 1;"&vbCrLf&_
"  -webkit-transform: scale(1.17);"&vbCrLf&_
"  -ms-transform: scale(1.17);"&vbCrLf&_
"  transform: scale(1.17);"&vbCrLf&_
"  -webkit-transition: background-color 317ms cubic-bezier(.298, -.338, .626, 1.252);"&vbCrLf&_
"  transition: background-color 317ms cubic-bezier(.298, -.338, .626, 1.252);"&vbCrLf&_
"  font-family: Verdana, Geneva, sans-serif;"&vbCrLf&_
"  text-transform: uppercase;"&vbCrLf&_
"}"

regular_msg_2 = ""&vbCrLf&_
".row-3 {"&vbCrLf&_
"  margin-top: 90px;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".section-2 {"&vbCrLf&_
"  height: 100vh;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".container-2 {"&vbCrLf&_
"  height: 100vh;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".div-block-5 {"&vbCrLf&_
"  width: 700px;"&vbCrLf&_
"  height: 500px;"&vbCrLf&_
"  margin-right: auto;"&vbCrLf&_
"  margin-left: auto;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".image-2 {"&vbCrLf&_
"  position: relative;"&vbCrLf&_
"  left: 108px;"&vbCrLf&_
"  top: 65px;"&vbCrLf&_
"  display: block;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".div-block-6 {"&vbCrLf&_
"  display: block;"&vbCrLf&_
"  height: auto;"&vbCrLf&_
"  margin-right: auto;"&vbCrLf&_
"  margin-left: auto;"&vbCrLf&_
"  padding-right: 13px;"&vbCrLf&_
"  padding-bottom: 13px;"&vbCrLf&_
"  padding-left: 13px;"&vbCrLf

regular_msg_2_1 = "  "&vbCrLf

regular_msg_2_2 = "  background-position: 0px 0px;"&vbCrLf&_
"  background-size: contain;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".div-block-7 {"&vbCrLf&_
"  position: static;"&vbCrLf&_
"  left: 0px;"&vbCrLf&_
"  top: 0px;"&vbCrLf&_
"  display: inline-block;"&vbCrLf&_
"  width: 100%;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".image-3 {"&vbCrLf&_
"  position: relative;"&vbCrLf&_
"  left: 15px;"&vbCrLf&_
"  top: 0px;"&vbCrLf&_
"  display: inline-block;"&vbCrLf&_
"  float: left;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".text-block {"&vbCrLf&_
"  max-width: 575px;"&vbCrLf&_
"  margin-right: auto;"&vbCrLf&_
"  margin-bottom: 0px;"&vbCrLf&_
"  margin-left: auto;"&vbCrLf&_
"  padding-right: 10px;"&vbCrLf&_
"  padding-bottom: 0px;"&vbCrLf&_
"  padding-left: 10px;"&vbCrLf&_
"  clear: left;"&vbCrLf&_
"  color: #fbf4c6;"&vbCrLf&_
"  text-align: center;"&vbCrLf&_
"}"&vbCrLf&_
".text-block-footer {"&vbCrLf&_
"  max-width: 575px;"&vbCrLf&_
"  margin-right: auto;"&vbCrLf&_
"  margin-bottom: 0px;"&vbCrLf&_
"  margin-left: auto;"&vbCrLf&_
"  padding-right: 10px;"&vbCrLf&_
"  padding-bottom: 0px;"&vbCrLf&_
"  padding-left: 10px;"&vbCrLf&_
"  clear: left;"&vbCrLf&_
"  color: #4b5668;"&vbCrLf&_
"  text-align: center;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".div-block-8 {"&vbCrLf&_
"  margin-bottom: 22px;"&vbCrLf&_
"  background-color: transparent;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".image-4 {"&vbCrLf&_
"  position: static;"&vbCrLf&_
"  left: 411px;"&vbCrLf&_
"  display: block;"&vbCrLf&_
"  max-width: auto;"&vbCrLf&_
"  margin-right: auto;"&vbCrLf&_
"  margin-left: auto;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf

regular_msg_2_3 = ".div-block-9 {"&vbCrLf&_
"  display: block;"&vbCrLf&_
"  margin-top: 54px;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".image-5 {"&vbCrLf&_
"  position: static;"&vbCrLf&_
"  left: 253px;"&vbCrLf&_
"  display: block;"&vbCrLf&_
"  margin-right: auto;"&vbCrLf&_
"  margin-left: auto;"&vbCrLf&_
"  border-bottom-color: #edd954;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".div-block-10 {"&vbCrLf&_
"  position: static;"&vbCrLf&_
"  left: 144px;"&vbCrLf&_
"  height: 25px;"&vbCrLf&_
"  max-width: 570px;"&vbCrLf&_
"  margin-right: auto;"&vbCrLf&_
"  margin-left: auto;"&vbCrLf&_
"  border-bottom: 1px solid #edd954;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".div-block-11 {"&vbCrLf&_
"  width: 570px;"&vbCrLf&_
"  margin-right: auto;"&vbCrLf&_
"  margin-left: auto;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".image-6 {"&vbCrLf&_
"  display: block;"&vbCrLf&_
"  margin-right: auto;"&vbCrLf&_
"  margin-left: auto;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".image-7 {"&vbCrLf&_
"  display: block;"&vbCrLf&_
"  margin-right: auto;"&vbCrLf&_
"  margin-left: auto;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".image-8 {"&vbCrLf&_
"  display: block;"&vbCrLf&_
"  margin-right: auto;"&vbCrLf&_
"  margin-left: auto;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".text-block-2 {"&vbCrLf&_
"  color: #fff;"&vbCrLf&_
"  text-align: center;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".row-4 {"&vbCrLf&_
"  margin-top: 25px;"&vbCrLf&_
"  margin-bottom: 9px;"&vbCrLf&_
"}"

regular_msg_3 = ".div-block-12 {"&vbCrLf&_
"  display: block;"&vbCrLf&_
"  margin-bottom: 21px;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".div-block-13 {"&vbCrLf&_
"  max-width: 158px;"&vbCrLf&_
"  min-width: 70px;"&vbCrLf&_
"  padding-right: 15px;"&vbCrLf&_
"  padding-left: 15px;"&vbCrLf&_
"  float: left;"&vbCrLf&_
"  clear: none;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".div-block-14 {"&vbCrLf&_
"  display: block;"&vbCrLf&_
"  max-width: 446px;"&vbCrLf&_
"  margin-top: 29px;"&vbCrLf&_
"  margin-right: auto;"&vbCrLf&_
"  margin-left: auto;"&vbCrLf&_
"  padding-bottom: 34px;"&vbCrLf&_
"  clear: none;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".div-block-15 {"&vbCrLf&_
"  position: static;"&vbCrLf&_
"  display: block;"&vbCrLf&_
"  max-width: 547px;"&vbCrLf&_
"  margin-right: auto;"&vbCrLf&_
"  margin-left: auto;"&vbCrLf&_
"  -webkit-justify-content: space-around;"&vbCrLf&_
"  -ms-flex-pack: distribute;"&vbCrLf&_
"  justify-content: space-around;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".paragraph-2 {"&vbCrLf&_
"  max-width: 575px;"&vbCrLf&_
"  margin-right: auto;"&vbCrLf&_
"  margin-left: auto;"&vbCrLf&_
"  color: #4b5668;"&vbCrLf&_
"  text-align: center;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf

regular_msg_3_1 = ".div-block-16 {"&vbCrLf&_
"  max-width: 575px;"&vbCrLf&_
"  margin-top: 32px;"&vbCrLf&_
"  margin-right: auto;"&vbCrLf&_
"  margin-left: auto;"&vbCrLf&_
"  background-color: transparent;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".div-block-17 {"&vbCrLf&_
"  width: 33%;"&vbCrLf&_
"  float: left;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".image-9 {"&vbCrLf&_
"  display: block;"&vbCrLf&_
"  width: 72px;"&vbCrLf&_
"  margin-right: auto;"&vbCrLf&_
"  margin-left: auto;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".paragraph-3 {"&vbCrLf&_
"  max-width: 150px;"&vbCrLf&_
"  margin-right: auto;"&vbCrLf&_
"  margin-left: auto;"&vbCrLf&_
"  padding-right: 9px;"&vbCrLf&_
"  padding-left: 9px;"&vbCrLf&_
"  color: #fff;"&vbCrLf&_
"  text-align: center; text-decoration: none; "&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".div-block-18 {"&vbCrLf&_
"  float: left;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".div-block-19 {"&vbCrLf&_
"  height: 100vh;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".image-10 {"&vbCrLf&_
"  position: relative;"&vbCrLf&_
"  left: 21px;"&vbCrLf&_
"  top: 17px;"&vbCrLf&_
"  float: left;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".div-block-20 {"&vbCrLf&_
"  width: 100%;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".div-block-21 {"&vbCrLf&_
"  position: static;"&vbCrLf&_
"  left: 0px;"&vbCrLf&_
"  right: 0px;"&vbCrLf&_
"  bottom: -12px;"&vbCrLf&_
"  width: 100%;"&vbCrLf&_
"  padding-top: 15px;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
".div-block-22 {"&vbCrLf&_
"  margin-top: 19px;"&vbCrLf&_
"  margin-bottom: 16px;"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
"@media (max-width: 991px) {"&vbCrLf&_
"  .div-block-10 {"&vbCrLf&_
"    width: 90%;"&vbCrLf&_
"  }"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
"@media (max-width: 767px) {"&vbCrLf&_
"  .div-block-6 {"&vbCrLf&_
"    display: block;"&vbCrLf&_
"  }"&vbCrLf&_
"}"&vbCrLf&_
""&vbCrLf&_
"@media (max-width: 479px) {"&vbCrLf&_
"  .image-5 {"&vbCrLf&_
"    max-width: 80%;"&vbCrLf&_
"  }"&vbCrLf&_
"  .image-10 {"&vbCrLf&_
"    overflow: hidden;"&vbCrLf&_
"    max-width: 80%;"&vbCrLf&_
"  }"&vbCrLf&_
"}"

regular_msg_4 = "}"&vbCrLf&_
"</style>"&vbCrLf&_
"<div class=""body"" style=""font-family: Arial, ""Helvetica Neue"", Helvetica, sans-serif;color: #333;font-size: 14px;line-height: 20px;background-color: #fff; "">"&vbCrLf&_
"  <div class=""div-block-19"" style=""height: 100vh;"">"&vbCrLf&_
"    <div class=""div-block-6 w-clearfix"" style=""display: block;height: auto;margin-right: auto;margin-left: auto;padding-right: 13px;padding-bottom: 13px;padding-left: 13px; background-position: 0px 0px; background-size: contain; background-image: url('%whole_background%'); "" >"&vbCrLf&_
"<img class=""image-3"" style=""position: relative; left: 15px; top: 0px; display: inline-block; float: left;"" src=""%spy_icon%"">"&_
"<img class=""image-10"" style=""overflow: hidden; max-width: 80%; margin-top: 16px; margin-left: 11px;"" src=""%header_icon%"">"

regular_msg_4_3 = "      <div class=""div-block-7"" style=""position: static;left: 0px;top: 0px;display: inline-block;width: 100%; background-color: #222d40"">"&vbCrLf&_
"        <div class=""div-block-9"" style=""display: block;margin-top: 54px;""><img class=""image-4"" src=""%locker_icon%"" style=""position: static;left: 411px;display: block;max-width: auto;margin-right: auto;margin-left: auto;"">"&vbCrLf

regular_msg_4_4 = "        </div>"&vbCrLf&_
"        <div class=""div-block-8"" style=""margin-bottom: 22px;background-color: transparent;""><img class=""image-5"" src=""%middle_header_text_base64%"" style=""position: static;left: 253px;display: block;margin-right: auto;margin-left: auto;border-bottom-color: #edd954;"">"&vbCrLf&_
"        </div>"&vbCrLf&_
"        <div>"&vbCrLf&_
"          <div class=""text-block"" style=""max-width: 575px;margin-right: auto;margin-bottom: 0px;margin-left: auto;padding-right: 10px;padding-bottom: 0px;padding-left: 10px;clear: left;color: #fbf4c6;text-align: center;"">%description%</div>"&vbCrLf&_
"        </div>"&vbCrLf&_
"        <div class=""div-block-10"" style=""position: static;left: 144px;height: 25px;max-width: 570px;margin-right: auto;margin-left: auto;border-bottom: 1px solid #edd954;""></div>"&vbCrLf&_
"        <div class=""div-block-20"" style=""width: 100%;"">"&vbCrLf&_
"          <div class=""div-block-16 w-clearfix"" style=""max-width: 575px;margin-top: 32px;margin-right: auto;margin-left: auto;background-color: transparent;"">"&vbCrLf&_
"            <div class=""div-block-17"" style=""width: 33%;float: left;"">"&vbCrLf&_
"              <div><img class=""image-9"" src=""%eapp_icon%"" style=""display: block;width: 72px;margin-right: auto;margin-left: auto;"">"&vbCrLf

regular_msg_4_5 = "              </div>"&vbCrLf&_
"              <div>"&vbCrLf&_
"                <a href=""%plugin_link%""><p class=""paragraph-3"" style=""text-decoration: none; max-width: 150px;margin-right: auto;margin-left: auto;padding-right: 9px;padding-left: 9px;color: #fff !important;text-align: center;"">%click_extension%</p></a>"&vbCrLf&_
"              </div>"&vbCrLf&_
"            </div>"&vbCrLf&_
"            <div class=""div-block-17"" style=""width: 33%;float: left;"">"&vbCrLf&_
"              <div><img class=""image-9"" src=""%promail_icon%"" style=""display: block;width: 72px;margin-right: auto;margin-left: auto;"">"&vbCrLf

regular_msg_4_6 = "              </div>"&vbCrLf&_
"              <div>"&vbCrLf&_
"                <a href=""%promail_link%""><p class=""paragraph-3"" style=""text-decoration: none; max-width: 150px;margin-right: auto;margin-left: auto;padding-right: 9px;padding-left: 9px;color: #fff !important;text-align: center;"">%click_promail%</p></a>"&vbCrLf&_
"              </div>"&vbCrLf&_
"            </div>"&vbCrLf&_
"            <div class=""div-block-17"" style=""width: 33%;float: left;"">"&vbCrLf&_
"              <div><img class=""image-9"" src=""%browser_icon%"" style=""display: block;width: 72px;margin-right: auto;margin-left: auto;"">"&vbCrLf&_
"              </div>"&vbCrLf&_
"              <div>"&vbCrLf&_
"                <a href=""%browser_link%""><p class=""paragraph-3"" style=""text-decoration: none; max-width: 150px;margin-right: auto;margin-left: auto;padding-right: 9px;padding-left: 9px;color: #fff !important;text-align: center;"">%click_browser%</p></a>"&vbCrLf&_
"              </div>"&vbCrLf&_
"            </div>"&vbCrLf&_
"          </div>"&vbCrLf&_
"        </div>"&vbCrLf&_
"        <div class=""div-block-22"" style=""margin-top: 19px;margin-bottom: 16px;"">"&vbCrLf&_
"          <div class=""text-block-footer"" style=""max-width: 575px;margin-right: auto;margin-bottom: 0px;margin-left: auto;padding-right: 10px;padding-bottom: 0px;padding-left: 10px;clear: left;color: #4b5668;text-align: center;"">%footer%</div>"&vbCrLf&_
"        </div>"&vbCrLf&_
"      </div>"&vbCrLf&_
"    </div>"&vbCrLf&_
"  </div>"&vbCrLf&_
"</div>"

regular_msg = regular_msg_1 & regular_msg_1_1 & regular_msg_1_2 &_
regular_msg_2 & regular_msg_2_1 & regular_msg_2_2 & regular_msg_2_3 &_
regular_msg_3 & regular_msg_3_1 &_
regular_msg_4 & regular_msg_4_1 & regular_msg_4_2 & regular_msg_4_3 & regular_msg_4_4 & regular_msg_4_5 & regular_msg_4_6