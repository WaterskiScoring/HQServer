<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->
<!--#include virtual="/rankings/tools_registration.asp"-->
<!--#include virtual="/rankings/tools_TourDefine.asp"-->
<!--#include virtual="/rankings/tools_Definitions.asp"-->
<!--#include virtual="/rankings/qualifications.asp"-->
<!--#include virtual="/rankings/RegFormDisplay.asp"-->
<!--#include virtual="/rankings/Register_Survey.asp"-->
<%


Dim sMemberID, sLastName, sFirstName, sFullName, sMembSex, sMembCity, sMembState, sMembAge, sMembPhone, sMembTypeID, sCanSkiTour, sMembTypeCode
Dim sMembEmail, sEffectiveTo, sMembBirth, sCostToUpgrade, sTypeDesc




' --- New 4-28-2013 - Gets SPECIAL WAIVER info from table based on SiteID rather than hard coding specific tournaments ---
'Dim swaiverSQL, sSpecialWaiverHeadline, sSpecialReleaseBannerText
swaiverSQL = "SELECT SpecialWaiverCode, SpecialWaiverHeadline, SpecialReleaseBannerText FROM usawsrank.TourExtras TE"
swaiverSQL = swaiverSQL + " JOIN sanctions.dbo.TSchedul AS TS"
swaiverSQL = swaiverSQL + "   ON SiteID=TS.TSiteID"
swaiverSQL = swaiverSQL + " WHERE LEFT(TS.TournAppID,6)='"&LEFT(sTourID,6)&"'"

Set rswaiver=Server.CreateObject("ADODB.recordset")
rswaiver.open swaiverSQL, sConnectionToTRATable, 3, 1

testwaiver=false
IF testwaiver=true AND sMemberID="000001151" THEN
		Response.write("<br>Found = ")
		response.write(NOT(rswaiver.eof))
		response.write("<br>rswaiver(SpecialWaiverHeadline) = "&rswaiver("SpecialWaiverHeadline"))
END IF

IF NOT(rswaiver.EOF) THEN
		sSpecialWaiverCode=rswaiver("SpecialWaiverCode")
		sSpecialWaiverHeadline=rswaiver("SpecialWaiverHeadline")
		sSpecialReleaseBannerText=rswaiver("SpecialReleaseBannerText")
END IF



' --- Now send the waiver ---
sTourID="13S999"
sSQL = "SELECT MemberID"
sSQL = sSQL + " FROM "&RegGenTableName 
sSQL = sSQL + " WHERE LEFT(TourID,6)='"&sTourID&"'"
sSQL = sSQL + " AND MemberID='000001151'"

SET rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable, 3, 3


CountSends=0
DO WHILE NOT rs.eof 

		sMemberID=rs("MemberID")
		
		' --- In Tools_Registration.asp ---
		SendSPECIALWaiverEmail_Tools sSpecialWaiverCode, sSpecialWaiverHeadline, sSpecialReleaseBannerText
		CountSends=CountSends+1
		rs.movenext
LOOP


response.write("<br><br>Waiver Resend COMPLETE")
response.write("<br>Total Messages Sent = "&CountSends)

%>