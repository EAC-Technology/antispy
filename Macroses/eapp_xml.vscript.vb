' vdomxml templates
 
' form to send
vdomxml_msg = "<CONTAINER name=""container1"" height=""550"" width=""844"" overflow=""3"" left=""0"" classname=""container_bg"">"&_
"%container_msg_sent%"&_
"  <FORM name=""form1"" top=""49"" height=""492"" width=""835"" meth=""event"" left=""2"" overflow=""3"">"&_
"    <CONTAINER name=""container3"" zindex=""3"" hierarchy=""1"" designcolor="""" height=""48"" width=""864"" backgroundcolor=""222D40"" left=""5"">"&_
"      <FORMBUTTON name=""formbutton1"" left=""730"" top=""7"" label=""%send%"" classname=""btn_ok"" width=""109"" height=""33"" />"&_
"      <HYPERTEXT name=""hypertext1"" zindex=""0"" top=""15"" height=""25"" classname=""css_text"" width=""169"" htmlcode=""%destination%"" overflow=""3"" left=""10""/>"&_
"      <FORMTEXT name=""formtext1"" tabindex=""1"" placeholder=""email"" top=""13"" height=""18"" width=""240"" left=""131""/>"&_
"      <HYPERTEXT name=""hypertext2"" visible=""1"" zindex=""0"" top=""15"" height=""25"" classname=""css_text"" width=""118"" htmlcode=""%delete_time%"" overflow=""3"" left=""401""/>"&_
"      <FORMTEXT name=""formtext2"" visible=""1"" tabindex=""2"" placeholder=""time"" value=""1"" top=""13"" height=""18"" width=""80"" left=""532""/>"&_
"    <FORMLIST name=""formlist1"" zindex=""0"" top=""13"" width=""79"" left=""633"" tabindex=""3"">"&_
"      <Attribute Name=""selectedvalue""><![CDATA[""m""]]></Attribute>"&_
"      <Attribute Name=""value""><![CDATA[{""m"": ""Minutes"", ""s"": ""Seconds"", ""h"": ""Hours"", ""d"": ""Days""}]]></Attribute>"&_
"    </FORMLIST>"&_
"    </CONTAINER>"&_
"    <FORMTEXTEDITOR name=""formtexteditor1"" tabindex=""4"" top=""44"" height=""451"" width=""860"" />"&_
"  </FORM>"&_
"<HYPERTEXT name=""hpt_js_css"" zindex=""10"" height=""42"" width=""75"" left=""629"">"&_
"  <Attribute Name=""htmlcode""><![CDATA[<style>"&_
".btn_ok {"&_
"  -webkit-border-radius: 3;"&_
"  -moz-border-radius: 3;"&_
"  border-radius: 3px;"&_
"  font-family: Courier New !important;"&_
"  color: #edd853;"&_
"  font-size: 16px !important;"&_
"  padding: 5px 20px 5px 20px;"&_
"  border: solid #edd853 1px;"&_
"  text-decoration: none;"&_
"  background: none;"&_
""&_
"  font-size: 16px !important;"&_
"  padding: 5px 20px 5px 20px;"&_
"  border: solid #edd853 1px;"&_
"  text-decoration: none;"&_
"  background: none;"&_
"}"&_
".btn_ok:hover {"&_
"  background: #354257;"&_
"  text-decoration: none;"&_
"  cursor: pointer;"&_
"}"&_
".css_text {"&_
"  font-family: Courier New !important;"&_
"  color: #edd853;"&_
"  font-size: 16px !important;"&_
"  text-decoration: none;"&_
"  background: none;"&_
"}"&_
""&_
".css_header {"&_
"  font-family: Courier New !important;"&_
"  color: #354257;"&_
"  font-size: 16px !important;"&_
"  text-decoration: none;"&_
"  background: none;"&_
"}"&_
""&_
".css_header_close {"&_
"  font-family: Courier New !important;"&_
"  color: #354257;"&_
"  font-size: 14px !important;"&_
"  font-weight: bold;"&_
"  background: none;"&_
"  text-align: right;"&_
"}"&_
".css_header_close:hover {"&_
"	text-decoration: underline;"&_
"}"&_
".container_bg {"&_
"  background-image: url(""%img_spy_bg%"");"&_
"  background-repeat:repeat;"&_
"  height: 100% !important;"&_
"  width: 100% !important;"&_
"}"&_
".container_bg_sent {"&_
"  background: url(""%img_spy_bg_sent%"");"&_
"  background-repeat:no-repeat;"&_
"  background-size: 100% 100%;"&_
"}"&_
".css_text_sent {"&_
"	font-size: 20pt !important;"&_
"	text-align: center;"&_
"}"&_
".css_text_sent_to {"&_
"	font-size: 12pt !important;"&_
"	text-align: center !important;"&_
"}"&_
".css_text_sent_dark {"&_
"	font-size: 20pt !important;"&_
"	text-align: center;"&_
"	color: #4D5971;"&_
"}"&_
"</style>"&_
"<script>"&_
"(function(doc) {"&_
"	"&_
"	function waitFormTextEditor() {"&_
"		var iframe = doc.querySelectorAll('.cleditorMain iframe')[0];"&_
""&_
"		if (!iframe) return;"&_
""&_
"		iframe.contentDocument.body.style.color = 'white';"&_
"		iframe.contentDocument.body.style.backgroundColor = '#354257';"&_
""&_
"		if (iframe.contentDocument.body.style.backgroundColor === '#354257') {doc.removeEventListener('DOMSubtreeModified', waitFormTextEditor);} "&_
"	}"&_
""&_
"	doc.addEventListener('DOMSubtreeModified', waitFormTextEditor);"&_
""&_
"})(document);"&_
"</script>"&_
"]]></Attribute>"&_
"</HYPERTEXT>"&_
"<HYPERTEXT name=""hypertext3"" zindex=""10"" top=""7"" height=""52"" width=""56"" left=""14"">"&_
"  <Attribute Name=""htmlcode""><![CDATA[<img src=""%img_spy%"" />]]></Attribute>"&_
"</HYPERTEXT>"&_
"  <HYPERTEXT name=""hypertext1"" zindex=""10"" top=""18"" height=""19"" classname=""css_header"" width=""418"" htmlcode=""Anti spy message"" left=""62""/>"&_
"</CONTAINER>"

' popup notification with form to send
vdomxml_msg_sent = "<CONTAINER name=""container1"" height=""515"" width=""567"">"&_
"  <CONTAINER name=""container2"" backgroundcolor=""FFFFFF"" zindex=""10"" height=""488"" width=""548"" left=""0"">"&_
"    <FORMBUTTON name=""formbutton1"" left=""229"" top=""366"" label=""Close""/>"&_
"    <TEXT name=""text1"" top=""330"" value=""Antispy email sent successfully"" fontsize=""14"" align=""center"" left=""74""/>"&_
"  </CONTAINER>"&_
"  <FORM name=""form1"" height=""515"" width=""548"" meth=""event"" overflow=""3"">"&_
"    <CONTAINER name=""container1"" designcolor=""FADCDC"" height=""60"" width=""548"" left=""0"">"&_
"      <FORMTEXT name=""formtext1"" top=""17"" width=""428"" left=""109""/>"&_
"      <TEXT name=""text1"" top=""22"" value=""Send to:"" width=""105"" left=""4""/>"&_
"    </CONTAINER>"&_
"    <CONTAINER name=""container3"" designcolor=""F28585"" top=""447"" height=""41"" width=""548"" left=""0"">"&_
"      <FORMBUTTON name=""formbutton1"" left=""442"" top=""10"" width=""106"" tabindex=""2""/>"&_
"    </CONTAINER>"&_
"    <FORMTEXTEDITOR name=""formtexteditor1"" top=""60"" height=""367"" width=""539"" tabindex=""1""/>"&_
"  </FORM>"&_
"</CONTAINER>"

vdomxml_msg_sent_container = "<CONTAINER name=""container2"" zindex=""10"" designcolor=""734646"" backgroundrepeat=""1"" top=""49"" height=""500"" classname=""container_bg_sent"" width=""864"" backgroundcolor=""FFFFFF"" left=""7"">"&_
"  <FORMBUTTON name=""formbutton1"" zindex=""2"" left=""355"" top=""336"" label=""OK"" classname=""btn_ok"" width=""140"" height=""40""/>"&_
"  <HYPERTEXT name=""hypertext1"" zindex=""2"" top=""233"" height=""40"" classname=""css_text css_text_sent_dark"" width=""366"" htmlcode=""antispy email"" left=""238""/>"&_
"%result_container%"&_
"</CONTAINER>"

vdomxml_msg_sent_to_container = "  <HYPERTEXT name=""hypertext2"" zindex=""2"" top=""274"" classname=""css_text css_text_sent"" width=""366"" htmlcode=""sent successfully to:"" left=""241""/>"&_
"  <TEXT name=""text_sent_emails""  align=""center"" color=""edd853"" zindex=""2"" value=""%emails_cut%"" hint=""%emails%"" top=""305"" classname=""css_text css_text_sent_to"" left=""70"" width=""724"" height=""25""/>"

vdomxml_msg_no_valid_emails = "  <HYPERTEXT name=""hypertext2"" zindex=""2"" top=""268"" classname=""css_text css_text_sent"" width=""366"" htmlcode=""please enter valid email address"" left=""240""/>"

vdomxml_secured_preview_xml = "<HYPERTEXT name=""hpt_show"" zindex=""0"" height=""488"" width=""630"" overflow=""3"">"&_
"  <Attribute Name=""htmlcode""><![CDATA["&_
"  %html_code%"&_
"  ]]></Attribute>"&_
"</HYPERTEXT>"


vdomxml_secured_preview_xml_updated = "<CONTAINER name=""container2"" zindex=""10"" designcolor=""B83737"" backgroundrepeat=""1"" height=""424"" classname=""container_bg_sent"" width=""583"" left=""0"">"&_
"  <HYPERTEXT name=""hypertext1"" zindex=""2"" top=""253"" height=""40"" classname=""css_text css_text_sent_dark"" width=""236"" htmlcode=""Loading..."" left=""167""/>"&_
"  <HYPERTEXT name=""hpt_js_css"" zindex=""10"" top=""52"" height=""42"" width=""75"" left=""247"">"&_
"    <Attribute Name=""htmlcode""><![CDATA[<style>"&_
".css_text {"&_
"  font-family: Courier New !important;"&_
"  color: #edd853;"&_
"  font-size: 16px !important;"&_
"  text-decoration: none;"&_
"  background: none; }"&_
".container_bg_sent {"&_
"  background: url(""%spy_img_bg%"");"&_
"  background-repeat:no-repeat;"&_
"  background-size: 100% 100%;"&_
"  position: relative!important;"&_
"  margin: 0 auto;"&_
"  top: auto !important;"&_
"  left: auto !important; }"&_
"#container { width: 100% !important;}"&_
".css_text_sent_dark {"&_
"	font-size: 20pt !important;"&_
"	text-align: center;"&_
"	color: #4D5971; }"&_
"</style>"&_
""&_
"]]></Attribute>"&_
"  </HYPERTEXT>"&_
"</CONTAINER>"


' Secured message
vdomxml_secured = "<CONTAINER name=""cnt_show"" height=""488"" classname=""container_bg"" width=""630"" overflow=""3"" left=""0"">"&_
"  <HYPERTEXT name=""hpt_title"" zindex=""10"" top=""15"" height=""19"" classname=""css_header"" width=""301"" htmlcode=""Anti spy message"" left=""62""/>"&_
"  <HYPERTEXT name=""hpt_remind"" zindex=""10"" top=""8"" height=""42"" classname=""css_time_label"" width=""79"" left=""414"">"&_
"    <Attribute Name=""htmlcode""><![CDATA[reminding<br/>time]]></Attribute>"&_
"  </HYPERTEXT>"&_
"  <HYPERTEXT name=""hpt_icon"" zindex=""0"" top=""4"" height=""52"" width=""56"" left=""14"">"&_
"    <Attribute Name=""htmlcode""><![CDATA[<img src=""%img_spy%"" />]]></Attribute>"&_
"  </HYPERTEXT>"&_
"  <HYPERTEXT name=""hpt_js_css"" zindex=""10"" height=""42"" width=""75"" left=""338"">"&_
"    <Attribute Name=""htmlcode""><![CDATA[<style>"&_
""&_
".css_text {"&_
" font-family: Courier New !important;"&_
" color: #edd853;"&_
" font-size: 16px !important;"&_
" text-decoration: none;"&_
" background: none;"&_
"}"&_
""&_
".css_header {"&_
" font-family: Courier New !important;"&_
" color: #354257;"&_
" font-size: 16px !important;"&_
" text-decoration: none;"&_
" background: none;"&_
"}"&_
""&_
".css_time_label {"&_
" font-family: Courier New !important;"&_
" color: #354257;"&_
" font-size: 14px !important;"&_
" font-weight: bold;"&_
" background: none;"&_
" text-align: right;"&_
" overflow: auto !important;    position: absolute !important;    z-index: 10 !important;    top: 8px !important;    left: auto !important; right: 130px !important;    width: 79px !important;    height: 42px !important;    display: block !important; }"&_
" .css_time {"&_
" font-family: Courier New !important;"&_
" color: #edd853;"&_
" background: #354257;"&_
" font-size: 14px !important;"&_
" font-weight: bold;"&_
" text-align: center;"&_
" vertical-align: middle;"&_
" display: inline-block;"&_
" line-height: 28px;"&_
" right: 6px; !important;"&_
" left: auto !important;"&_
" border: 1px solid rgba(0,0,0,0.15);"&_
"    border-radius: 2px;"&_
"    transition: all 0.3s ease-out;"&_
"    box-shadow:"&_
"        inset 0 1px 0 rgba(255,255,255,0.5),"&_
"        0 2px 2px rgba(0,0,0,0.3),"&_
"        0 0 4px 1px rgba(0,0,0,0.2);"&_
"}"&_
""&_
".css_msg {    font-size: 14px !important;    overflow: auto !important;    position: absolute !important;    z-index: 0 !important;    width: inherit !important;    bottom: 1% !important;    display: block !important;    padding: 10px 3px 9px 10px !important;    background-color: #ffffff !important;    overflow-y: auto !important;    Top: 47px !important;    Left: 0px !important; height: auto !important} "&_
".css_msg_container {    position: relative !important;    width: 99% !important;    margin: 0 auto !important;    height: inherit !important;    top: 0px !important;    Left: 0px !important;    overflow-x: hidden !important;}"&_
".container_bg {"&_
" background-image: url(""%img_spy_bg%""); "&_
" background-repeat:repeat;"&_
" position: absolute !important;    top: 0px !important;    left: 0px !important;    width: 100%  !important;    overflow: hidden !important;    height: 100% !important;"&_
"} "&_
""&_
"</style>"&_
" "&_
"<script> "&_
" "&_
"var countDownDate = new Date().getTime() + (%remaining_time_ms% * 1000); "&_
"function y() { "&_
" var x = setInterval(function() { "&_
"   var now = new Date().getTime(); "&_
"   var distance = countDownDate - now; "&_
"   var hours = Math.floor((distance % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60)); "&_
"   var minutes = Math.floor((distance % (1000 * 60 * 60)) / (1000 * 60)); "&_
"   var seconds = Math.floor((distance % (1000 * 60)) / 1000); "&_
"   document.getElementsByClassName(""css_time"")[0].innerHTML = (""0""+hours).slice(-2) + "":"" + (""0""+minutes).slice(-2) + "":"" + (""0""+seconds).slice(-2); "&_
"   if (distance < 0) { "&_
"     clearInterval(x); "&_
"     document.getElementsByClassName(""css_time"")[0].innerHTML = ""00:00:00""; "&_
"     document.getElementsByClassName(""css_msg"")[0].innerHTML = ""<i>Message deleted</i>""; "&_
"   } "&_
" }, 1000); "&_
"}; "&_
"y(); "&_
" "&_
"</script> "&_
""&_
"]]></Attribute>"&_
"  </HYPERTEXT>"&_
"  <CONTAINER name=""cnt_text"" top=""47"" height=""426"" classname=""css_msg_container"" width=""605"" left=""13"">"&_
"    <HYPERTEXT name=""hpt_msg"" zindex=""0"" top=""18"" height=""392"" classname=""css_msg"" width=""573"" left=""17"">"&_
"      <Attribute Name=""htmlcode""><![CDATA[%SECURED_MSG%]]></Attribute>"&_
"    </HYPERTEXT>"&_
"  </CONTAINER>"&_
"  <HYPERTEXT name=""hpt_time"" zindex=""10"" top=""9"" height=""27"" classname=""css_time"" width=""105"" htmlcode=""%remaining_time%"" overflow=""3"" left=""511""/>"&_
"</CONTAINER>"