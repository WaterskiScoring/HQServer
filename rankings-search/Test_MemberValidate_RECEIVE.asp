<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include16.asp"-->
<!--#include virtual="/rankings/tools_registration16.asp"-->
<%


Dim MemberID, MemberFound, MemberExpireDate, MembershipType, CanSkiTour
Dim RemHost, RefURL, TestValue
Dim MemberFirst, MemberLast
Dim xmlDoc



RefURL = Request.ServerVariables("HTTP_HOST")
RemHost = Request.ServerVariables("REMOTE_ADDR")
' RefURL = HttpContent.Current.Request.UrlReferrer.AbsolutePath





MemberID=""
MemberFound="no"
MemberExpireDate=""
MembershipType=""
CanSkiTour="no"


' --- Sets object ---
Set xmlDoc=CreateObject("Microsoft.XMLDOM")
xmlDoc.async="false"
xmlDoc.load(Request)
xml_posted = replace(replace(xmldoc.xml,"<","*"),">","$")

' --- Determines if a valid post was made - If so, then parse the XML the data ---
Set xmlList = xmlDoc.getElementsByTagName("MemberStatus")
NumNodes=xmlDoc.getElementsByTagName("MemberStatus").length

FOR Each xmlItem In xmlList
		' -- MemberFound = LCASE(xmlItem.childNodes(0).text)
		MemberID = xmlDoc.getElementsByTagName("MemberID")(0).text		
		MemberFound = xmlDoc.getElementsByTagName("MemberFound")(0).text
		MemberFirst = xmlDoc.getElementsByTagName("MemberFirst")(0).text
		MemberLast = xmlDoc.getElementsByTagName("MemberLast")(0).text
		MembershipType = xmlDoc.getElementsByTagName("MembershipType")(0).text
		MemberExpireDate = xmlDoc.getElementsByTagName("MemberExpireDate")(0).text
		CanSkiTour = xmlDoc.getElementsByTagName("CanSkiTour")(0).text
NEXT



SendMyEmail








' ----------------------
   SUB SendMyEmail
' ----------------------


Dim eMailSubj, eBody, eMailFrom, eMailBCC, sTest
Dim sTsEmail

eMailSubj="Test Listener RESPONSE from XML Member Validation Post"
eMailTo="mark@productdesign-biz.com"
eMailFrom="mark@productdesign-biz.com"
'eMailCC="paul@paulsantangelo.com"
eMailCC=""
eMailBCC=""

ebody = "<html>"
ebody = ebody & "<head>"
ebody = ebody & "<title>Notice to Mark Crone</title>"
ebody = ebody & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
ebody = ebody & "</head>"
ebody = ebody & "<body bgcolor=""#FFFFFF"" text=""#000000"">"
ebody = ebody & "<div align=""center"">"


ebody = ebody & "<TABLE BORDER=4 ALIGN=CENTER CELLPADDING=3 CELLSPACING=0 BGCOLOR="&TableColor1&" width=60% >"
ebody = ebody & "<TR>"
ebody = ebody & "<TD BGCOLOR=blue><center><font face="&font1&" color=#FFFFFF size=4><b>Notice to Mark Crone - RESPONSE</b></font></TD>"
ebody = ebody & "</TR>"
 
ebody = ebody & "<TR>"
ebody = ebody & "<td Align=center>"	
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>This message was generated because the Member Validation Posted.</b></font>"
ebody = ebody & "</td>"
ebody = ebody & "</TR>"

ebody = ebody & "<TR>"
ebody = ebody & "<td Align=center>"	
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>MemberID: "&MemberID&"</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "</td>"
ebody = ebody & "</TR>"

ebody = ebody & "<TR>"
ebody = ebody & "<td Align=center>"	
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>MemberFound: "&MemberFound&"</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "</td>"
ebody = ebody & "</TR>"

ebody = ebody & "<TR>"
ebody = ebody & "<td Align=center>"	
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>MemberFirst: "&MemberFirst&"</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "</td>"
ebody = ebody & "</TR>"

ebody = ebody & "<TR>"
ebody = ebody & "<td Align=center>"	
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>MemberLast: "&MemberLast&"</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "</td>"
ebody = ebody & "</TR>"

ebody = ebody & "<TR>"
ebody = ebody & "<td Align=center>"	
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>MemberExpireDate: "&MemberExpireDate&"</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "</td>"
ebody = ebody & "</TR>"

ebody = ebody & "<TR>"
ebody = ebody & "<td Align=center>"	
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>MembershipType: "&MembershipType&"</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "</td>"
ebody = ebody & "</TR>"

ebody = ebody & "<TR>"
ebody = ebody & "<td Align=center>"	
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>CanSkiTour: "&CanSkiTour&"</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "</td>"
ebody = ebody & "</TR>"

ebody = ebody & "<TR>"
ebody = ebody & "<td Align=center>"	
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>Referring URL: "&RefURL&"</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "</td>"
ebody = ebody & "</TR>"

ebody = ebody & "<TR>"
ebody = ebody & "<td Align=center>"	
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>Remote Host: "&RemHost&"</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "</td>"
ebody = ebody & "</TR>"

ebody = ebody & "<TR>"
ebody = ebody & "<td Align=center>"	
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>xml_posted (NOTE: < replaced by * and > replaced by &): "&xml_posted&"</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "</td>"
ebody = ebody & "</TR>"

ebody = ebody & "</TABLE>"







SetupEmailService

objMessage.To = eMailTo
objMessage.From = eMailFrom
IF LCASE(eMailBCC)<>"" THEN objMessage.bcc = eMailBCC
IF LCASE(eMailCC)<>"" THEN objMessage.cc = eMailCC

objMessage.Subject = eMailSubj
objMessage.HTMLBody = ebody
 
 ' --- Finally send the message, and then clear that object
objMessage.Send
set objMessage = Nothing


END SUB


%>