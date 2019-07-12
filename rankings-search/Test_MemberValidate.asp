<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include16.asp"-->
<!--#include virtual="/rankings/tools_registration16.asp"-->
<%

' Call Response.AddHeader("Access-Control-Allow-Origin", "http://usawaterski.org")




Dim First_Name, Last_Name, requestMemberID
Dim sMemberID, sEffectiveto, sTypeDesc, sCanSkiTour, sMemberFound 
Dim sFirstName, sLastName
Dim ClubAccessCode
Dim RefURL, RemHost, xml_posted, NumNodes
Dim AdminEmailAddress, SendEmailDuringTesting
Dim Member_DOB_ForValidation, IsMinor_AsOf_Date, sMinor, sDOB_validated


' --- Sends email to Mark Crone
SendEmailDuringTesting="N"
AdminEmailAddress = "mark@productdesign-biz.com"



RefURL = Request.ServerVariables("HTTP_HOST")
RemHost = Request.ServerVariables("REMOTE_ADDR")


IF RemHost = "50.62.177.105" THEN

		' --- Reads variables ---
		ReadValidationRequestXML

		' --- Looks up member record ---
		SearchMemberData

		' --- Posts back to requesting URL
		' Validation_PostBack

		' --- Writes XML to page ---
		Validation_WriteToBrowser
ELSE
		' --- Do nothing ---

END IF


' ---------------------------------------------------------------------------------------
' ---- BOTTOM OF MAIN CODE
' ---------------------------------------------------------------------------------------






' ------------------------------
  SUB ReadValidationRequestXML
' ------------------------------  


' MemberRequest = "<USAWaterski><MemberStatus><ClubAccessCode>"&ClubAccessCode&"</ClubAccessCode><MemberID>"&sMemberID&"</MemberID><Member_DOB_ForValidation>mm/dd/yyy</Member_DOB_ForValidation><IsMinor_AsOf_Date>mm/dd/yyyy</IsMinor_AsOf_Date></MemberStatus></USAWaterski>"

' --- Sets object ---
Set xmlDoc=CreateObject("Microsoft.XMLDOM")
xmlDoc.async="false"
xmlDoc.load(Request)
xml_posted = replace(replace(xmldoc.xml,"<","*"),">","$")

' --- Determines if a valid post was made - If so, then parse the XML the data ---
Set xmlList = xmlDoc.getElementsByTagName("MemberStatus")
NumNodes=xmlDoc.getElementsByTagName("MemberStatus").length
DOB_ValidationLength=xmlDoc.getElementsByTagName("Member_DOB_ForValidation").length
IsMinor_ValidationLength=xmlDoc.getElementsByTagName("IsMinor_AsOf_Date").length

requestMemberID=""
ClubAccessCode=""
Member_DOB_ForValidation=""
IsMinor_AsOf_Date=""

FOR Each xmlItem In xmlList
		requestMemberID = xmlDoc.getElementsByTagName("MemberID")(0).text
		ClubAccessCode = xmlDoc.getElementsByTagName("ClubAccessCode")(0).text		
		IF DOB_ValidationLength>0 THEN Member_DOB_ForValidation = xmlDoc.getElementsByTagName("Member_DOB_ForValidation")(0).text
		IF IsMinor_ValidationLength>0 THEN IsMinor_AsOf_Date = xmlDoc.getElementsByTagName("IsMinor_AsOf_Date")(0).text		
NEXT


END SUB



' ------------------------------
  SUB Validation_WriteToBrowser
' ------------------------------  

MemberResponse = "<USAWaterski><MemberStatus><MemberID>"&sMemberID&"</MemberID><MemberFound>"&sMemberFound&"</MemberFound><MemberFirst>"&sFirstName&"</MemberFirst><MemberLast>"&sLastName&"</MemberLast><MemberExpireDate>"&sEffectiveto&"</MemberExpireDate><MembershipType>"&sTypeDesc&"</MembershipType><CanSkiTour>"&sCanSkiTour&"</CanSkiTour><Minor>"&sMinor&"</Minor><DOB_validated>"&sDOB_validated&"</DOB_validated></MemberStatus></USAWaterski>"
'MemberResponse = "<USAWaterski><MemberStatus><MemberID>"&sMemberID&"</MemberID><MemberFound>"&sMemberFound&"</MemberFound><MemberFirst>"&sFirstName&"</MemberFirst><MemberLast>"&sLastName&"</MemberLast><MemberExpireDate>"&NumNodes&"</MemberExpireDate><MembershipType>"&sTypeDesc&"</MembershipType><CanSkiTour>"&sCanSkiTour&"</CanSkiTour><Minor>"&sMinor&"</Minor><DOB_validated>"&sDOB_validated&"</DOB_validated></MemberStatus></USAWaterski>"

Response.write(MemberResponse)


END SUB





' ------------------------
  SUB Validation_PostBack
' ------------------------  



' --- Test to USA Waterski --- 
url = "http://usawaterski.org/rankings/Test_MemberValidate_RECEIVE.asp"
MemberResponse = "<USAWaterski><MemberStatus><MemberID>"&sMemberID&"</MemberID><MemberFound>"&sMemberFound&"</MemberFound><MemberFirst>"&sFirstName&"</MemberFirst><MemberLast>"&sLastName&"</MemberLast><MemberExpireDate>"&sEffectiveto&"</MemberExpireDate><MembershipType>"&sTypeDesc&"</MembershipType><CanSkiTour>"&sCanSkiTour&"</CanSkiTour></MemberStatus></USAWaterski>"
Set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
xmlhttp.Open "POST", url, false
xmlhttp.setRequestHeader "Content-Type", "text/xml" 
xmlhttp.send MemberResponse


END SUB








' ------------------------
  SUB SearchMemberData 
' ------------------------

searchMemberID = 0
IF IsNumeric(requestMemberID) = true THEN searchMemberID=requestMemberID

	sSQL = "SELECT TOP 1 FirstName, LastName, City, State, Sex, Phone"
	sSQL = sSQL + ", MembershipTypeCode"
	sSQL = sSQL + ", Birthdate, Email, EffectiveTo"  
	sSQL = sSQL + ", Description"
	sSQL = sSQL + ", coalesce(MembershipTypeID,0) AS MembershipTypeID"
	sSQL = sSQL + ", coalesce(CanSkiInTournaments,0) AS CanSkiInTournaments"
	sSQL = sSQL + ", coalesce(CanSkiInGRTournaments,0) AS CanSkiInGRTournaments"
	sSQL = sSQL + ", coalesce(TypeCode,'XXX') AS TypeCode"

	sSQL = sSQL + " FROM "&MemberLiveTableName&" MT"
	sSQL = sSQL + " LEFT JOIN "&MemberTypeTableName&" MTT ON MTT.MembershipTypeID = MT.MembershipTypeCode"
	sSQL = sSQL + " WHERE PersonID = cast(right("&sqlclean(searchMemberID)&",8) AS INTEGER)"

	set rsMemb=Server.CreateObject("ADODB.recordset")
	rsMemb.open sSQL, sConnectionToTRATable, 3, 1

	
	sTypeDesc = ""
	sCanSkiTour = "no"
	sCanSkiGRTour = "no"
	sEffectiveto = ""
	sMemberFound = "no"
	sMemberID = requestMemberID	
	sFirstName = ""
	sLastName = ""
	sMinor=""
	sDOB_validated=""
	
	IF NOT rsMemb.eof THEN
			sMemberFound = "yes"
			sFirstName = SQLClean(rsMemb("FirstName"))
			sLastName = SQLClean(rsMemb("LastName"))
			sFullName = SQLClean(rsMemb("FirstName")&" "&rsMemb("LastName"))
			sMembCity = SQLClean(rsMemb("City"))
			sMembState = rsMemb("State")
			sMembSex = rsMemb("Sex")
			sMembPhone = rsMemb("Phone")
			sMembBirth = rsMemb("Birthdate")
			sMembEmail = rsMemb("Email")
			sEffectiveto = rsMemb("Effectiveto")
		
			sMembTypeID = rsMemb("MembershipTypeID")
			sCanSkiTour = rsMemb("CanSkiInTournaments")
		
			IF rsMemb("CanSkiInTournaments")=1 THEN sCanSkiTour="yes"
				
			IF rsMemb("CanSkiInGRTournaments")=1 THEN sCanSkiGRTour="yes"

			IF IsDate(Member_DOB_ForValidation)=true THEN
					IF Member_DOB_ForValidation=sMembBirth THEN sDOB_validated="yes"
			END IF
			sDOB_validated = IsDate(Member_DOB_ForValidation)
				
			IF IsDate(IsMinor_AsOf_Date)=true THEN
					IF DateDiff("d",sMembBirth,IsMinor_AsOf_Date) >= 18*365	THEN sMinor="yes"
					sMinor = DateDiff("d",sMembBirth,IsMinor_AsOf_Date)
			END IF			
			
			sMembTypeCode = rsMemb("TypeCode")
			sTypeDesc = rsMemb("Description")
	END IF	

	' --- Sends Email to Mark Crone during testing ---
	IF SendEmailDuringTesting="Y" THEN SendMyEmail

END SUB	  









' ----------------------
   SUB SendMyEmail
' ----------------------

Dim eMailSubj, eBody, eMailFrom, eMailBCC, sTest
Dim sTsEmail

OrgEmail = "competition@usawaterski.org"
OrgFriendlyFrom = "Member Validation"

eMailTo = AdminEmailAddress
eMailFrom = OrgEmail
' &" "&OrgFriendlyFrom
eMailCC=""
eMailBCC=""
eMailSubj="Member Validation Access Request"


ebody = "<html>"
ebody = ebody & "<head>"
ebody = ebody & "<title>Notice to Mark Crone</title>"
ebody = ebody & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
ebody = ebody & "</head>"
' ebody = ebody & "<body bgcolor=""#FFFFFF"" text=""#000000"">"
ebody = ebody & "<body style=""background-color=white;"">"
ebody = ebody & "<div align=""center"">"


' ebody = ebody & "<TABLE BORDER=4 ALIGN=CENTER CELLPADDING=3 CELLSPACING=0 BGCOLOR="&TableColor1&" width=100% >"
ebody = ebody & "<TABLE align=center style=""border:4px solid; padding:3px; background-color:"""&TableColor1&"""; width=100%; max-width:400px;"">"
ebody = ebody & "<TR>"
ebody = ebody & "<TD BGCOLOR=red><center><font face="&font1&" color=#FFFFFF size=4><b>Notice to Mark Crone - REQUEST</b></font></TD>"
ebody = ebody & "</TR>"
 
ebody = ebody & "<TR>"
ebody = ebody & "<td Align=center>"	
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>This message was generated by REQUEST FOR VALIDATION.</b></font>"
ebody = ebody & "</td>"
ebody = ebody & "</TR>"

ebody = ebody & "<TR>"
ebody = ebody & "<td Align=center>"	
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>MemberID: "&sMemberID&"</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "</td>"
ebody = ebody & "</TR>"

ebody = ebody & "<TR>"
ebody = ebody & "<td Align=center>"	
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>MemberFound: "&sMemberFound&"</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "</td>"
ebody = ebody & "</TR>"

ebody = ebody & "<TR>"
ebody = ebody & "<td Align=center>"	
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>MemberExpireDate: "&sEffectiveto&"</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "</td>"
ebody = ebody & "</TR>"

ebody = ebody & "<TR>"
ebody = ebody & "<td Align=center>"	
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>TypeDesc: "&sTypeDesc&"</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "</td>"
ebody = ebody & "</TR>"

ebody = ebody & "<TR>"
ebody = ebody & "<td Align=center>"	
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>CanSkiTour: "&sCanSkiTour&"</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "</td>"
ebody = ebody & "</TR>"

ebody = ebody & "<TR>"
ebody = ebody & "<td Align=center>"	
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>ClubAccessCode: "&ClubAccessCode&"</b></font>"
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
ebody = ebody & "<font face="&font1&" size=2><b>Date Time: "&NOW&"</b></font>"
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
