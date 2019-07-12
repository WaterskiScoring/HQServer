<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/Bio-Form-Print.asp"-->
<!--#include virtual="/rankings/Tools_Definitions.asp"-->
<!--#include virtual="/rankings/qualifications.asp"-->
<!--#include virtual="/rankings/Tools_include.asp"-->
<!--#include virtual="/rankings/tools_registration.asp"-->
<%

Server.ScriptTimeout = 3000 

Dim ThisFileName
ThisFileName="registration-email.asp"





' ---------------------------------------------------------------------------------------------
' --- This module displays various reports associated with REGISTRATION functions 
' --- Original module created by Mark Crone
' --- LAST updated: 7/4/2009
' ---------------------------------------------------------------------------------------------


Dim RegionSelected, EventSelected, DivSelected, StateSelected, WhatReport, WhatPayments, WhatNotify, WhatLetter
Dim SequenceSelected, sBioFilter, sQualFilter, sWaivFilter, sFeeFilter, sResendEmail, sSentBioEmail
Dim emailbuttonstatus, FileLetter, currentline, ECount, MaxECount, ebody
Dim EVT1_TIME, EVT2_TIME, EVT3_TIME

' --- Used in OLR Listing ---
Dim sSortBy, sIncludePast, sIncludeFuture
Dim PreviousDiv, PreviousEvent, SkiYearID

Dim sDiv1, sDiv2, sDiv3, sDiv4, sEvent1, sEvent2, sEvent3, sEvent4
Dim sMemberID, LastMemb, sFirstName, sLastName, sFullName

Dim sLast4
Dim EmailCount, sState, sNoEmail, sEmail

Dim MembStatusTitle, MembStatusText, MembStatuscolor, FeesText, Feescolor, FeesTitle, regstatuscolor, regstatusText, regstatusTitle
Dim QualStatusEvent1, QualStatusEvent2, QualStatusEvent3, QualStatusEvent4, sRequirePart

Dim WaiverTitle, Waivercolor, WaiverText, TrickTitle, Trickcolor, TrickText
Dim BioText, Biocolor, BioTitle, BioLink, sRefreshData, SeedCount

Dim StartCharSelected, EndCharSelected
Dim PrintButton
Dim TextDropcolor1, TextDropcolor2

Dim sSQL, sSkiYearID

Dim sShowSQL
Dim sTestMode




' --- For debugging ---
sShowSQL = Request("sShowSQL")


'sTestMode="<br>**TEST MODE**"



'Session("sTourID")="07W999A"
'sTourID="07W999A"

'Session("sTourID")="08S093"
'sTourID="08S093"





' --- Resets to blank for testing purposes
'IF TRIM(Request("process"))="reset" THEN Session("sTourID")=""
IF TRIM(Request("process"))="viewreg" THEN Session("sTourID")=""


adminmenulevel = Session("adminmenulevel")
IF adminmenulevel = "" THEN adminmenulevel = 0

sMemberID = TRIM(Request("sMemberID"))
RegionSelected = trim(Request("RegionSelected"))
EventSelected = trim(Request("EventSelected"))
DivSelected = trim(Request("DivSelected"))
StateSelected = trim(Request("StateSelected"))
WhatReport = LCASE(trim(Request("WhatReport")))





SequenceSelected = LCASE(trim(Request("SequenceSelected")))
WhatNotify=TRIM(Request("WhatNotify"))
WhatLetter=TRIM(Request("WhatLetter"))
sBioFilter=Request("sBioFilter")
sQualFilter=Request("sQualFilter")
sWaivFilter=Request("sWaivFilter")
sFeeFilter=Request("sFeeFilter")
sResendEmail=Request("sResendEmail")
sRefreshData=Request("sRefreshData")

sSkiYearID = Request("sSkiYearID")




StartCharSelected=Request("StartPulldown")
EndCharSelected=Request("EndPulldown")
IF StartCharSelected="" OR WhatReport = "seeding" THEN StartCharSelected="All"
IF EndCharSelected="" OR WhatReport = "seeding" THEN EndCharSelected="All"

sPrintDate=Request("sPrintDate")
PrintButton=Request("PrintButton")
ReturnButton=Request("ReturnButton")
IF ReturnButton="Main Menu" THEN
	Response.redirect("/rankings/defaultHQ.asp")
END IF


'----------------------------------------
' --- Sets Default values for report  ---
'----------------------------------------

'IF WhatReport="" THEN WhatReport="regstat"
IF WhatReport = "" THEN WhatReport = "noreportselected"
IF WhatReport = "seeding" THEN StateSelected="" 

IF TRIM(Request("DivSelected")) = "" THEN DivSelected = "ALL"
IF TRIM(RegionSelected) = "" THEN RegionSelected = 6
IF TRIM(Request("StateSelected")) = "" THEN StateSelected = "All"
IF TRIM(SkiYearSelected) = "" THEN SkiYearSelected = 1

' ---- This will need a condition depending on which sports division  ----
IF TRIM(Request("EventSelected")) = "" THEN EventSelected = "ALL"
IF SequenceSelected = "" THEN SequenceSelected = "seed"


DispData="N"
IF DispData="Y" THEN
		response.write("<br>EventSelected="&EventSelected)
		response.write("<br>DivSelected="&DivSelected)
		response.write("<br>WhatNotify="&WhatNotify)
		response.write("<br>WhatLetter="&WhatLetter)
		response.write("<br>sBioFilter="&sBioFilter)
		'response.end
END IF
	





IF WhatReport="tourstatus" THEN
		' --- Do nothing
ELSEIF TRIM(Request("sTourID"))<>"" THEN
		Session("sTourID")=TRIM(Request("sTourID"))
		sTourID=TRIM(Request("sTourID"))
ELSE
		IF TRIM(Session("sTourID"))="" THEN
				' - Go get tournament
				Session("sSendingPage") = "/rankings/"&ThisFileName&"?rid="&rid
				Session("sTourID") = ""

				sl=Request("sl")
				tr=Request("tr")
				ju=Request("ju")
				wb=Request("wb")
				ws=Request("ws")
				wu=Request("wu")
				bf=Request("bf")
				kb=Request("kb")
				hy=Request("hy")
				hf=Request("hf")
				jd=Request("jd")
				ad=Request("ad")

	   		IF sl="on" OR tr="on" OR ju="on" THEN		
						response.redirect("/rankings/view-tournamentsHQ.asp?sl=on&tr=on&ju=on&process=viewreg&rid="&rid)
				ELSEIF wb="on" OR ws="on" OR wu="on" THEN
						response.redirect("/rankings/view-tournamentsHQ.asp?wb=on&ws=on&wu=on&process=viewreg&rid="&rid)		
				ELSE
						response.redirect("http://www.usawaterski.org")	
				END IF		
		ELSE
				sTourID = Session("sTourID")
		END IF
END IF



' --- In SUB Tools_Registration.asp - Sets all tournament variables ---
IF WhatReport<>"tourstatus" THEN
		DefineTourVariables_New

		' --- SUB found in tools_include.asp - Defines what events this sTSptsGrpID offers ---
		RegistrationEventsOffered (sTSptsGrpID)
END IF








' http://www.usawaterski.org/rankings/registration-email.asp?WhatReport=notifications&WhatLetter=Bio&WhatNotify=PrintBio&sTourID=14C999&EventSelected=S&DivSelected=B3





'response.write("<br>WhatReport = "&WhatReport)
'response.end




SELECT CASE WhatReport

  CASE "viewbio"
	
			Session("sSendingPage") = "/rankings/"&ThisFileName
	
			rdState = "/rankings/bio-form.asp?FormStatus=new&EditStatus=disabled&sMemberID="&sMemberID&"&sTourID="&sTourID
			response.redirect("/rankings/bio-form.asp?FormStatus=new&BioStatus=disabled&sMemberID="&sMemberID&"&sTourID="&sTourID)


  CASE "notifications"
	

			BuildSeedingQuery
			MailNotices	


  CASE "noreportselected"

			'response.write("EXIT")
			'response.end
			IF PrintButton<>"Printer Friendly" THEN  WriteIndexPageHeader
			
			DropBoxFormat1
			%><br><br><center><font size=<% =fontsize3 %> color="<% =Textcolor3 %>"><b>Select Report and Other Settings <br>then Press 'Display Report'</b></FONT></center><%
			IF PrintButton<>"Printer Friendly" THEN  WriteIndexPageFooter

	CASE "endoffile"

			IF PrintButton<>"Printer Friendly" THEN  WriteIndexPageHeader
			%><br><br><br><center><font size=<% =fontsize3 %> color="<% =Textcolor3 %>"><b>No Data Matching Search Settings</b></FONT></center><%
			IF PrintButton<>"Printer Friendly" THEN  WriteIndexPageFooter

END SELECT



' ------------------------------------------------------------------------------------------------------------------------------
' ---------------------   END OF MAIN SECTION OF PROGRAM    --------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------------------







' ----------------------
  SUB DisplayNoDataLine
' ----------------------
	%>
	<table align=center width=100%>
	  <td align=center><font size=<% =fontsize3 %> color=<% =textcolor3 %>><br><b>No Data For These Settings.</b></font></td>
	</table><%


END SUB


' ----------------------
  SUB EndofReportLine
' ----------------------
	%>
	<center><font size=<% =fontsize3 %> color=<% =textcolor3 %>><b>Hold your cursor over certain fields or headings to view details about that item.</b></font></center><%


END SUB




' -------------------
  SUB SeedingHeading
' -------------------



	%>
	<TABLE class="innertable" align="Center" BORDER="1" CELLPADDING="3" CELLSPACING="0" style="background-color:<%=Tablecolor1%>;" width=100%>

	  <TR>	
	      <th align="Center" ColSpan="4" valign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Personal</FONT></th>      
	      <th align="Center" ColSpan="2" valign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Entry Selections</FONT></th>      
	      <th align="Center" ColSpan="5" valign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Seeding & Placement</FONT></th>
	  </TR>

	  <TR><%
		SeedColWidth=55	%>		
	      <th align="Center"  bgcolor="<%=Headcolor1%>" valign="top"><font size=<%=fontsize2%> color="#FFFFFF">Num</FONT></th>
	      <th align="Left"  bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Name</FONT></th>
	      <th align="center" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">MemberID</FONT></th>
	      <th align="Center" width="<%=SeedColWidth%>" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF"><a TITLE="State">ST</a></FONT></th>

	      <th align="Center" width="<%=SeedColWidth%>" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">
			<a TITLE="Entry Classification is coded G-Grassroots S-ClassC (or Base Class) and R-Record (Upgrade)">Class</a></FONT></th>
	      <th align="Center" width="<%=SeedColWidth%>" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">
			<a TITLE="Skill Level is used for ability-based grouping of competitors">Skill</a></FONT></th>
	      <th align="Center" width="<%=SeedColWidth%>" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">
			<a TITLE="Seeding Value - Same as Ranking Value - Also the order of participation, highest value skis last">Rank<br>Score</a></FONT></th>
	      <th align="Center" width="<%=SeedColWidth%>" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">
			<a TITLE="Current Qualification Level based on position on Rankings List">Rank<br>Level</a></FONT></th>
	      <th align="Center" width="<%=SeedColWidth%>" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">
			<a TITLE="Level from Ranking List to be Used as Qualification Method beginning in 2008">Rank<br>Pctl</a></FONT></th>
	      <th align="Center" width="<%=SeedColWidth%>" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">
			<a TITLE="Most Recent National Placement">Natl<br>Place</a></FONT></th>
	      <th align="Center" width="<%=SeedColWidth%>" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">
			<a TITLE="Most Recent Regional Placement">Regl<br>Place</a></FONT></th>
	  </TR><%

END SUB





' --------------------
  SUB DropBoxFormat1
' --------------------

	
	' --- Defines the Report Title ---
	SELECT CASE WhatReport
		CASE "seeding"
				ReportTitle = "Seeding Summary"&sTestMode	
		CASE "skierpayments"
				ReportTitle = "Payments By Type"&sTestMode	
		CASE "regstat"
				ReportTitle = "Registration Status"&sTestMode	
		CASE "bystate"
				ReportTitle = "Participants By State"&sTestMode	
		CASE "scratched"
				ReportTitle = "Scratch List"&sTestMode	
		CASE "othersales"
				ReportTitle = "Other Sales Report"&sTestMode	
		
	END SELECT


	' --- Looks up image to display in header based on link between TSiteID in TSchedul

	Set rs=Server.CreateObject("ADODB.recordset")

' --- Old ---
'	sSQL= " SELECT HeaderImage FROM usawsrank.TourExtras WHERE LEFT(TourID,6)='"&LEFT(sTourID,6)&"'"

' --- New ---
	sSQL = "SELECT HeaderImage FROM sanctions.dbo.TSchedul AS TS"
	sSQL = sSQL + " JOIN usawsrank.TourExtras AS TE"
	sSQL = sSQL + "   ON SiteID=TS.TSiteID"
	sSQL = sSQL + " WHERE TournAppID='"&sTourID&"'"

	rs.open sSQL, sConnectionToTRATable, 3, 1


	' --- Uses a default image if it does not find a site image ---
	IF (NOT rs.eof) AND (NOT PrintButton="Printer Friendly") THEN 
			MainImage="images\LOCSites\"&TRIM(rs("HeaderImage"))
			TextDropcolor1="#FFFFFF"
			TextDropcolor2=Textcolor2
	ELSE
			TextDropcolor1="#000000"	
			TextDropcolor2=Textcolor2
			MainImage="images\LOCSites\AMFog.jpg"			
	END IF


	%>
     <form action="/rankings/<%=ThisFileName%>" method="post">

	<TABLE align="center" class="droptable" WIDTH="740px" height=180px background="<%=MainImage%>">
	  <tr>
			<td align="left" colspan=4 width=60%>
				<font size=3 color="<% =TextDropcolor1 %>"><b>&nbsp;&nbsp;<% =sTourName %></b>&nbsp;&nbsp;&nbsp;&nbsp</font>
				<font size=<%=fontsize2%> color="<% =TextDropcolor2 %>"><br>&nbsp;&nbsp;<% =sTDateS %>-<% =sTDateE %> - <%=sTourID%></b></font>
			</td>	
			<td colspan=2 align="left" width=40%>
				<FONT color="<%=TextDropcolor1%>" size=3><B><% Response.Write(ReportTitle) %></B></FONT>
				<br><br>
			</td>
	 </tr>

	 <tr><%
		' --- Loads List of report options ---
		LoadReportPullDown  

		' --- Loads divisions offered in this event ---
	    	LoadDivPulldown %>
	    <td width=100px>&nbsp;</td>
	    <td width=100px>&nbsp;</td>
	 </tr>

	 <tr><%

	   ' --- Loads drop down to set sequence ---
	   LoadDropSequence

	   ' --- Loads Event Drop down ---	
	   LoadEventPulldownNew %>
	    <td>&nbsp;</td>
	    <td>&nbsp;</td>
	  </tr>

	 <tr><%

		' --- State or Region dropdown depending on WhatReport selected
		IF WhatReport="bystate" THEN 
			LoadStatePulldown 
		ELSE 	
			LoadRegionPulldown 
		END IF %>
	    <td>&nbsp;</td>
	    <td>&nbsp;</td><%

		IF AdminMenuLevel>=50 THEN  %>	
  			<td colspan=1 valign=top align="left">
				<FONT color="<% =Titlecolor %>" size="<%=fontsize2%>"><b>Show SQL</b></font>
				<input type=checkbox name="sShowSQL" <% IF sShowSQL="on" THEN response.write "checked" %>>

			</td><%
		ELSE  %>
			<td>&nbsp;</td><%
		END IF %>

	    <td>&nbsp;</td>
	  </tr>

	  <tr>	
	   <td align=center colspan=2><%
	      IF PrintButton="Printer Friendly" THEN %>	
		      <input type="submit" align="center" style="width:10em;"  value="Report Update"><%
	      ELSE %>
		      <input type="submit" align="center" style="width:10em;"  value="Display Report"><%
	      END IF %>	
	   </td>

	   <td align=center colspan=2><%
	      IF PrintButton="Printer Friendly" THEN %>	
			<a href='#' onclick='window.print()' title="Click here to Print"><input type=submit value="Print Now" style="width:9em"></a><%
	      ELSE %>
			<input type="submit" align="center" style="width:10em;" name="PrintButton" value="Printer Friendly"><%
	      END IF %>
	   </td>

	   <td align=center colspan=2><%
	      IF PrintButton="Printer Friendly" THEN 
			' --- Don't Display
	      ELSE %>
			<input type="submit" align="center" style="width:10em;" name="ReturnButton" value="Main Menu"><%
	      END IF %>
	   </td>


	 </tr>
	</table>
	</form>

	<%

END SUB





' -----------------
  SUB SendMarkEmail
' -----------------


Set objCDO = Server.CreateObject("CDO.Message")

objCDO.To = marksemailaddress
objCDO.From = "USA Water Ski<competition@usawaterski.org>"
objCDO.Subject = "Error in determining SkiYearID"

ebody="There was an error in determining the Ski Year ID - sTourID = "&sTourID  

ebody=ebody & "<br><br>"&sSQL
objCDO.HTMLBody = ebody	
objCDO.Send
Set objCDO = Nothing
ebody=""


END SUB




' ----------------------
  SUB BuildSeedingQuery
' ----------------------


	' ----------------------------------------------------------------------------------------------------------
	' -----------  Builds SQL string to define display values  -------------------------------------------------
	' ----------------------------------------------------------------------------------------------------------

	Set rs=Server.CreateObject("ADODB.recordset")

	' --- Set SkiYear first and if tournament is in 12 Month range then SkiYear=1 will be the top row of answerset ---
	Dim RankSkiYear
	sSQL = "SELECT SkiYearID FROM usawsrank.SkiYear"
	sSQL = sSQL + "	WHERE BeginDate<='"&sTDateS&"' AND EndDate>='"&sTDateS&"'"
	rs.open sSQL, sConnectionToTRATable, 3, 1

	IF rs.eof THEN
		SendMarkEmail		
		RankSkiYear=1
	ELSE
		RankSkiYear=rs("SkiYearID")
			
	END IF


	IF DateDiff("d",Date,sTDateS)<=365 AND DateDiff("d",Date,sTDateS)>=0 THEN RankSkiYear=1
	
'	response.write("<br>"& DateDiff("d",Date,sTDateS)) 
'	response.write("<br>RankSkiYear="&RankSkiYear)
'	response.write("<br>sSQL="&sSQL)
'	response.end

	rs.close


' --------------------------------
' --- Begin SQL Query contruct ---
' --------------------------------

sSQL = " SELECT EVT.MemberID, EVT.div, EVT.event, EVT.QfyOverride, EVT.FeeClass, EVT.Skill" 

sSQL = sSQL + ", RGEN.RegDate, RGEN.EntryType, RGEN.WaiverCode, RGEN.TotalEntry, RGEN.BanquetQty, RGEN.MembOverride, RGEN.RegionalOverride" 
sSQL = sSQL + ", RGEN.MoneyOverride, coalesce(RGEN.SentBioEmail,'N') AS SentBioEmail" 

'sSQL = sSQL + ", RT.Rank, coalesce(RT.RankPct,0) as RankPct, coalesce(RT.natl_plc,' ') AS natl_plc, coalesce(RT.regl_plc,' ') AS regl_plc, RT.Reg_Ski" 
'sSQL = sSQL + ", RT.Rating, RT.Reg_Ski" 

'sSQL = sSQL + ", coalesce(RQ.SkiedRegls,' ') AS SkiedRegls, RQ.QfyStatusTextNew" 
'sSQL = sSQL + ", coalesce(TP.Payments, 0) AS Payments" 

' *** FUTURE : Get Email from Members ***
'sSQL = sSQL + ", PW.Email" 

sSQL = sSQL + ", MEM.firstname, MEM.lastname, MEM.EffectiveTo, MEM.MemberShipTypeCode, UPPER(MEM.[state]) AS 'state', MEM.City, MEM.NoEmail" 
sSQL = sSQL + ", MEM.MembEmail" 

'sSQL = sSQL + ", MTT.MemberType, MTT.MemberShipTypeID, MTT.CanSkiInTournaments, MTT.CanSkiInGRTournaments, MTT.MembTypeDesc" 

sSQL = sSQL + ", BIO.BioMemberID" 

sSQL = sSQL + ", TGEN.Form1Name, TGEN.Form2Name, TGEN.Form3Name, TGEN.Form4Name, TGEN.Form5Name, TGEN.Form6Name, TGEN.EmailAddress, TGEN.QualLevel, TGEN.Bio_Reqd" 

'sSQL = sSQL + ", REGION.region" 

'sSQL = sSQL + ", coalesce(LT.RequirePart,'-') AS RequirePart, LT.LT_LeagueID" 


sSQL = sSQL + "	FROM "&RegDetailTableName&" AS EVT" 

sSQL = sSQL + "	JOIN" 
sSQL = sSQL + "		( SELECT MemberID, TourID, RegisterDate as 'RegDate', EntryType, WaiverCode, TotalEntry, BanquetQty, MembOverride, RegionalOverride," 
sSQL = sSQL + "				MoneyOverride, coalesce(SentBioEmail,'N') AS SentBioEmail" 
sSQL = sSQL + "			FROM "&RegGenTableName&") AS RGEN" 
sSQL = sSQL + "	ON EVT.MemberID = RGEN.MemberID AND LEFT(EVT.TourID,6) = LEFT(RGEN.TourID,6)" 

'sSQL = sSQL + "	LEFT JOIN" 
'sSQL = sSQL + "		( SELECT MemberID, Event, Div, SC_3, SkiYearID, RankScore as 'Rank', coalesce(RankPct,0) as RankPct," 
'sSQL = sSQL + "			coalesce(natl_plc,' ') AS natl_plc, coalesce(regl_plc,' ') AS regl_plc, Reg_Ski, AWSA_Rat AS 'Rating'" 
'sSQL = sSQL + "			FROM "&RankTableName
'sSQL = sSQL + "				WHERE SkiYearID='1' AND SC_3 IS NULL AND Event IS NOT NULL AND Div IS NOT NULL) AS RT" 
'sSQL = sSQL + "	ON EVT.MemberID=RT.MemberID AND EVT.Div=RT.Div AND EVT.Event=RT.Event"  

' --- Gets total amount paid by this member in this tournament
'sSQL = sSQL + "	LEFT JOIN" 
'sSQL = sSQL + "		( SELECT MemberID, SUM(Amount) AS Payments" 
'sSQL = sSQL + "				FROM "&RegPaymentTableName 
'sSQL = sSQL + "					WHERE LEFT(TourID,6) = '"&sTourID&"' and Result = '0'" 
'sSQL = sSQL + "				GROUP BY MemberID) AS TP" 
'sSQL = sSQL + "	ON TP.MemberID = EVT.MemberID"

'sSQL = sSQL + "	LEFT JOIN" 
'sSQL = sSQL + "		( SELECT MemberID, TourID, Event, Div, coalesce(SkiedRegls,' ') AS SkiedRegls, QfyStatus AS QfyStatusTextNew" 
'sSQL = sSQL + "			FROM "&RegQualifyTableName&" ) AS RQ" 
'sSQL = sSQL + "	ON EVT.MemberID=RQ.MemberID AND LEFT(EVT.TourID,6)=LEFT(RQ.TourID,6) AND EVT.Event=RQ.Event AND EVT.Div=RQ.Div" 

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		( SELECT PersonID, firstname, lastname, EffectiveTo, MemberShipTypeCode, UPPER(state) AS 'state', City, DoNotEmail AS NoEmail, Email AS MembEmail" 
sSQL = sSQL + "			FROM "&MemberLiveTableName&") AS MEM" 
sSQL = sSQL + "	ON CAST(RIGHT(RGEN.MemberID,8) AS INT) = MEM.PersonID" 

'sSQL = sSQL + "	LEFT JOIN" 
'sSQL = sSQL + "		( SELECT MembershipTypeID, TypeCode AS 'MemberType', CanSkiInTournaments, CanSkiInGRTournaments, Description AS MembTypeDesc" 
'sSQL = sSQL + "			FROM "&MemberTypeOLRTableName&") AS MTT" 
'sSQL = sSQL + "	ON MEM.MembershipTypeCode = MTT.MembershipTypeID" 
	
'sSQL = sSQL + "	LEFT JOIN" 
'sSQL = sSQL + "		( SELECT MemberID, Email"
'sSQL = sSQL + "			FROM "&RegPWTableName&" WHERE MemberID IS NOT NULL) AS PW" 
'sSQL = sSQL + "	ON PW.MemberID = RGEN.MemberID" 
	
'sSQL = sSQL + "	LEFT JOIN" 
'sSQL = sSQL + "		( SELECT State, region" 
'sSQL = sSQL + "			FROM "&RegionTableName&") AS REGION" 
'sSQL = sSQL + "	ON LOWER(MEM.[state]) = LOWER(REGION.[state])" 
	
sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		( SELECT MemberID AS BioMemberID"
sSQL = sSQL + "			FROM "&BioTableName&") AS BIO" 
sSQL = sSQL + "	ON RGEN.MemberID = BIO.BioMemberID" 

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		( SELECT TournAppID, Form1Name, Form2Name, Form3Name, Form4Name, Form5Name, Form6Name, EmailAddress, QualLevel, Bio_Reqd"
sSQL = sSQL + "			FROM "&TRegSetupTableName&") AS TGEN" 
sSQL = sSQL + "	ON LEFT(TGEN.TournAppID,6) = LEFT(RGEN.TourID,6)" 

'sSQL = sSQL + "	LEFT JOIN" 
'sSQL = sSQL + "		( SELECT QualifyTour, coalesce(RequirePart,'-') AS RequirePart, LeagueID AS LT_LeagueID"  
'sSQL = sSQL + "			FROM "&LeagueTableName&") AS LT" 
'sSQL = sSQL + "	ON TGEN.TournAppID = LEFT(LT.QualifyTour,6)" 




	' -----------------------------------
	' ------ Begin WHERE condition ------
	' -----------------------------------

	sSQL = sSQL + " WHERE LEFT(EVT.[TourID],6) = '"&LEFT(sTourID,6)&"'"

	IF DivSelected = "ALL" THEN
			'sSQL = sSQL + " AND (EVT.div IN ('MM', 'OM', 'OW', 'B1', 'B2','B3', 'G1', 'G2', 'G3', 'M1', 'M2', 'M3', 'M4', 'M5', 'M6', 'M7', 'M8', 'M9', 'MA', 'MB', 'W1', 'W2', 'W3','W4', 'W5', 'W6','W7', 'W8', 'W9', 'WA', 'WB'))"
	ELSE
			sSQL = sSQL + " AND EVT.div = '"&DivSelected&"'"
	END IF

	IF EventSelected = "ALL" THEN 
			'sSQL = sSQL + " AND (EVT.event IN ('S', 'T', 'J', 'WB', 'WS', 'WU', 'KS', 'KT', 'KF', 'KR'))"
	ELSE
			sSQL = sSQL + " AND (EVT.event = '"&EventSelected&"')"
	END IF



	' --- First and Last Characters when report gets too big
	IF StartCharSelected<>"All" THEN
		sSQL = sSQL + " AND (LEFT(MEM.LastName,1)>='"&StartCharSelected&"')"		
	END IF
	
	IF EndCharSelected<>"All" THEN
		sSQL = sSQL + " AND (LEFT(MEM.LastName,1)<='"&EndCharSelected&"')"		
	END IF




'	IF TRIM(sPrintDate)<>"" THEN
'		sSQL = sSQL + " AND RGEN.RegisterDate='"&sPrintDate&"'"			
'	END IF


'	IF RegionSelected <> "6" THEN sSQL = sSQL + " AND REGION.[region] = '"&RegionSelected&"'"
'	IF StateSelected <> "All" THEN sSQL = sSQL + " AND MEM.State = '"&StateSelected&"'"


	' ------------------------------------
	' ------ Sets ORDER of Display  ------
	' ------------------------------------

	IF WhatReport="notifications" THEN
		sSQL = sSQL + " ORDER BY MEM.LastName, MEM.FirstName, EVT.event"
'	ELSEIF WhatReport="bystate" THEN
'		sSQL = sSQL + " ORDER BY MEM.State, MEM.LastName, MEM.FirstName"
'	ELSEIF SequenceSelected="alpha" THEN
'		sSQL = sSQL + " ORDER BY EVT.div, EVT.event, MEM.LastName, MEM.FirstName"
'	ELSEIF SequenceSelected="seed" OR SequenceSelected="regstat" THEN
'			sSQL = sSQL + " ORDER BY EVT.div, EVT.event, RT.Rank DESC, MEM.LastName, MEM.FirstName"
'	ELSEIF SequenceSelected="regdate" THEN
'			sSQL = sSQL + " ORDER BY EVT.div, EVT.event, RegDate, MEM.LastName, MEM.FirstName"
'	ELSEIF SequenceSelected="regdateall" THEN
'			sSQL = sSQL + " ORDER BY RegDate, MEM.LastName, MEM.FirstName"
	ELSE
		' sSQL = sSQL + " ORDER BY EVT.div, EVT.event, MEM.LastName, MEM.FirstName"
	END IF

	Set rs=Server.CreateObject("ADODB.recordset")
	'rs.CommandTimeout = 90
	'rs.ConnectionTimeout = 90
	rs.open sSQL, sConnectionToTRATable, 3, 1

  'Set rs = Server.CreateObject("ADODB.Connection")
  'rs.ConnectionTimeout = 2000
  'rs.Open Application("sConnectionToTRATable")
  'rs.CommandTimeout = 2000

	' response.write("<br>"&sSQL)

	IF Session("adminmenulevel")>=50 AND TRIM(sShowSQL)<>"" THEN
		response.write("<br>"&sSQL)
'		response.end
	END IF

	IF rs.eof THEN WhatReport = "EndofFile"




END SUB












' -------------------------------
  SUB MailNotices
' -------------------------------

MaxECount=6


'IF NOT WhatNotify="PrintBio" THEN WriteIndexPageHeader
' --- Changed 8-9-2013 ---
IF NOT WhatNotify="PrintBio" THEN 
	
		WriteIndexPageHeader


		' ---- Heading Section for all reports ----
		ReportTitle = "Bios and Notifications"

		%>
		<TABLE class="innertable" WIDTH=750px  >
			<tr>
	   		<td align="left" colspan=4 width=60%>
					<font size=3 color="<% =Textcolor2 %>"><b><% =sTourName %></b>&nbsp;&nbsp;&nbsp;&nbsp</font>
					<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><br><% =sTDateS %>-<% =sTDateE %></b></font>
	   		</td>	
	   		<td colspan=2 align="left" width=40%>
		  		<font color="<%=Textcolor2%>" size=3><B><% Response.Write(ReportTitle) %></B></FONT>
					<br><br>
	   		</td>
	 		</tr>

	  	<form action="/rankings/<%=ThisFileName%>" method="post">

	  	<tr><%
			LoadReportPulldown

			' --- Loads divisions offered in this event  
			LoadDivPulldown 

			LoadStartPulldown %>

	  	</tr>
	  	<tr>	
	    	<td align="right">		
					<font size=<% =fontsize3 %> color="<% =Textcolor1 %>">Action:</font>
	    	</td>	
	    	<td align="left">		
    			<select name="WhatNotify">
						<option value=""<%IF WhatNotify = "" THEN Response.Write(" SELECTED ")%>>Not Selected</option>
						<option value="ViewNotices"<%IF WhatNotify = "ViewNotices" THEN Response.Write(" SELECTED ")%>>Preview Document</option>
	        	<option value="ViewList"<%IF WhatNotify = "ViewList" THEN Response.Write(" SELECTED ")%>>Show Target List</option>
	        	<%
						IF AdminMenuLevel>=30 THEN %>
	        		<option value="SendList"<%IF WhatNotify = "SendList" THEN Response.Write(" SELECTED ")%>>Send Emails</option><%
						END IF 	
						%>
	        	<option value="PrintBio"<%IF WhatNotify = "PrintBio" THEN Response.Write(" SELECTED ")%>>Print Bios</option>
        	</select>
	    	</td><%

				LoadEventPulldownNew  	

				LoadEndPulldown %>
	  	</tr>

	  	<tr>
	    	<td align="right">		
					<font size=<% =fontsize3 %> color="<% =Textcolor1 %>">Document:</font>
	    	</td><%

	    	IF AdminMenuLevel>=30 THEN %>
						<td align="left">		
							<select name="WhatLetter">
								<option value=""<%IF WhatLetter = "" THEN Response.Write(" SELECTED ")%>>Not Selected</option>
								<option value="reg_deficiency"<%IF WhatLetter = "reg_deficiency" THEN Response.Write(" SELECTED ")%>>Deficiency</option>
        				<option value="reg_bioincomplete"<%IF WhatLetter = "reg_bioincomplete" THEN Response.Write(" SELECTED ")%>>Bio Incomplete</option>
	      				<option value="reg_confirm"<%IF WhatLetter = "reg_confirm" THEN Response.Write(" SELECTED ")%>>Confirmation</option>
        				<option value="Bio"<%IF WhatLetter = "Bio" THEN Response.Write(" SELECTED ")%>>Personal Bio</option>
        				<%
								IF adminmenulevel>=50 THEN %>
	      					<option value="custom"<%IF WhatLetter = "custom" THEN Response.Write(" SELECTED ")%>>Custom</option><%
								END IF 
								%>
		    			</select>
						</td><%
	    	ELSE %>
					<input type="hidden" name="WhatLetter" value="Bio">		
					<td align="left">
						<font size=<% =fontsize3 %> color="<% =Textcolor2 %>">Bio Form</font>
					</td><%
	    	END IF %>
	
		
	    	<td>&nbsp;</td>	
	    	<td>&nbsp;</td><%

	    	IF AdminMenuLevel>=30 THEN %>
		    	<td align="right">
						<font size=<% =fontsize3 %> color="<% =Textcolor1 %>">Register Date:</font>
		    	</td>
		    	<td align="left">	
						<input type="text" name="sPrintDate" value= "<% =sPrintDate %>" maxlength="10" size="10" >
						<font size=<% =fontsize3 %> color="<% =Textcolor1 %>">mm/dd/yyyy</font>
		    	</td><%
	    	ELSE  %>
		    	<td>&nbsp;</td>
		    	<td>&nbsp;</td><%
	    	END IF  %>

	  	</tr><% 

	  	IF adminmenulevel >= 19  THEN %>
		  	<tr>
		    	<td align="right">
		        <font size=<% =fontsize3 %> color="<% =Textcolor1 %>">Scratch Filter&nbsp;</FONT>
		    	</td>
		    	<td align="left" colspan=3>	
						<font size=<% =fontsize3 %>>Bio(<%=MaxECount%>)</font>
						<input type=checkbox name="sBioFilter" <% IF sBioFilter = "on" THEN Response.Write("Checked") %>>
						<font size=<% =fontsize3 %>>Qualify</font>
						<input type=checkbox name="sQualFilter" <% IF sQualFilter = "on" THEN Response.Write("Checked") %>>
						<font size=<% =fontsize3 %>>Waiver</font>
						<input type=checkbox name="sWaivFilter" <% IF sWaivFilter = "on" THEN Response.Write("Checked") %>>
						<font size=<% =fontsize3 %>>Fee Due</font>
						<input type=checkbox name="sFeeFilter" <% IF sFeeFilter = "on" THEN Response.Write("Checked") %>>
		    	</td>

		    	<td align="right">
						<font size=<% =fontsize3 %>><b>Resend:</b></font>
		    	</td>
		    	<td align="left">	
						<input type=checkbox name="sResendEmail" <% IF sResendEmail = "on" THEN Response.Write("Checked") %>>
						<font size=<% =fontsize3 %>>If previously printed</font>
		    	</td>
		  	</tr><%
	  	END IF %>


	  	<tr>
	    	<td align=center colspan=6>
	      	<input type="submit" align="center" value="Submit Action">
	    	</td>
	   	</form>
	  	</tr>
		</TABLE>

		<br>

		<TABLE class="innertable" WIDTH=750px >
  		<tr>
   			<td width=10%>&nbsp;</td>	
   			<td align="left" colspan=4 width=60%>
					<br>
					<font size=4><b>NOTICE:</b></font>
					<br>
					<font size=<%=fontsize2%>>Authorized announcers may view Skier Bio online from links on the Registration Status report.</font>
					<br><br>
					<font size=<%=fontsize2%>>Bio printing follows three (3) steps. </font>
					<br>
					<font size=<%=fontsize2%>>&nbsp;&nbsp;&nbsp;1) Select the document design and Preview Document as the Action.  Confirm the bio format is what you desire to print.   </font>
					<br>
					<font size=<%=fontsize2%>>&nbsp;&nbsp;&nbsp;2) Display the target list of recipients. </font>
					<br>
					<font size=<%=fontsize2%>>&nbsp;&nbsp;&nbsp;3) Print the selected document or Send It to recipient list.  </font>
					<br><br>

					<font size=<%=fontsize2%>>Bios received for registrations on specified dates may be printed.  Larger tournaments may need to split the print run to <br>avoid buffer overflow by selecting ranges of names using the <b>Start With</b> and <b>End With</b> selections</font>
					<br><br>
					<font size=<%=fontsize2%>>It is recommended to print all events for a division as this makes the bio more versatile.</font>
					<br><br>
					<font size=<%=fontsize2%>>The scratch notice filter selections (Bio6, Qualify, Waiver and Fee Due) are used in conjunction with Deficiency <br>Notice document.</font>

					<br>
					<br>
	    	</td>
	  	</tr>
		</TABLE>



		<%
END IF   ' --- Bottom of IF-THEN for not displaying the drop down ---


' --- Performs function ---
If NOT rs.eof THEN
		
		' -------------------------------------------------------
		' --- Displays list of people receiving email message ---
		' -------------------------------------------------------
		GenerateRecordSet

ELSE  
		%>
		<br>
		<center><font size=<% =fontsize3 %> color=<% =textcolor3 %>><b>No Output For These Settings.</b></font></center>
		<%
END IF

IF NOT WhatNotify="PrintBio" THEN  WriteIndexPageFooter


END SUB







' ------------------------
   SUB GenerateRecordSet
' ------------------------

' --- Display list of members receiving email 
sDiv(1)=""
sDiv(2)=""
sDiv(3)=""
sDiv(4)=""

rs.movefirst

IF NOT rs.eof THEN 

	IF WhatNotify="ViewList" THEN  %>

		<br>
		<center><font size=5 font=<%=font1%> ><b>TARGET LIST</b></font></center>
		<br>
		<TABLE align="Center" class="innertable" width=100%>
 	    <TR>
	    	<th align="left"><font size=<%=fontsize2%> color="#FFFFFF"><b>&nbsp;Member</b></FONT></th>
	      <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>State</b></FONT></th>
		    <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>&nbsp;<%=sTEventName(1)%></b></FONT></th>
		    <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>&nbsp;<%=sTEventName(2)%></b></FONT></th>
		    <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>&nbsp;<%=sTEventName(3)%></b></FONT></th>
		    <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>&nbsp;<%=sTEventName(4)%></b></FONT></th>
		    <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>OK to<br>Email</b></FONT></th>
		    <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>Email Address</b></FONT></th>
		    <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>BioScore</b></FONT></th>
		    <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>Bio<br>Sent</b></FONT></th>
		  </TR><%	

	END IF

	IF WhatNotify="PrintBio" THEN
	    %>
	    <TABLE class="innertable" align=center height="120px" width="100%">
	    	<TR>
					<th align=center colspan=2 height="30px" style="border-style:none; vertical-align:middle;">	
						<font size="<%=fontsize4%>" color="#FFFFFF">Registration Bio Print Function</font>
		  		</th>
		  	</TR>
	    	<TR>
					<td align=center style="border-style:none; vertical-align:bottom">	
		  			<form action="" method="post"> 	
	          	<input type="submit" value="Print Displayed Bios" style="width:14em" title="Print From Form" onclick="window.print()">
		  			</form>
		  		</td>
					<td align=center style="border-style:none; vertical-align:bottom">	
		  			<form action="" method="post" > 	
	          	<input type="submit" value="Close Window" style="width:14em" title="Close this window" onclick="window.close()">
		  			</form>
		  		</td>
		  	</TR>
	    	<TR>
					<td align=center colspan=2 style="border-style:none;">	
						<font size=<%=fontsize3%> color="#000000">All Bios for the Division and Event selected are displayed below.  <br>Press the <b>Print Displayed Bios</b> button to send them to your printer. After printing, press the <b>Close Window</b> button to return to the Registration Status report and print another set of bios</font>
		  		</td>
		  	</TR>
		  </TABLE>	
			<%

	END IF

	LastMemb=rs("MemberID")

	EMailCount=0


	DO WHILE NOT rs.eof  

		  SELECT CASE TRIM(rs("Event"))
				CASE sTEvent(1) 
						sDiv(1)=rs("Div")
				CASE sTEvent(2)
						sDiv(2)=rs("Div")
				CASE sTEvent(3) 
						sDiv(3)=rs("Div")
				CASE sTEvent(4) 
						sDiv(4)=rs("Div")
		  END SELECT

		  sState=rs("State")
		  sFullName=rs("LastName")&", "&rs("FirstName")
		  sNoEmail=rs("NoEmail")
		  sEmail=rs("MembEmail")
		  'sTourEmail=rs("EmailAddress")
		  sSentBioEmail=TRIM(rs("SentBioEmail"))	


		  rs.movenext
		  IF NOT rs.eof THEN sMemberID=rs("MemberID")




		  ' ---- After collecting each event information then display this members info	---
		  IF rs.eof OR sMemberID<>LastMemb THEN 

			' --- View the list of selected ---
		 	IF WhatNotify="ViewList" THEN  
				
			    ' --- Checks whether bio form is complete ---
			    HowEmptyIsForm LastMemb, sTourID

					'response.write("INSIDE PONT 1")			
					'response.write("sBioFilter= "&sBioFilter)
					'response.write("ECount= "&ECount)
					'response.write("MaxECount= "&MaxECount)
			    ' --- Display info to screen ---
			    IF WhatLetter<>"reg_bioincomplete" OR (WhatLetter="reg_bioincomplete" AND sBioFilter="on" AND ECount >= MaxECount AND (sSentBioEmail="N" OR (TRIM(sSentBioEmail)="Y" AND sResendEmail="on"))) THEN
							'response.write("INSIDE PONT 2a")
				 			%>	 
			      	<TR>
			        	<TD align="left" >
									<font size=<%=fontsize2%>>
										<a title="MemberID: <%=sMemberID%>"><%=sFullName%></a>
									</font>
				      	</TD>
	        		  <TD align="Center" ><font size=<%=fontsize2%>><%=sState%></FONT></TD>
				      	<TD align="Center" ><font size=<%=fontsize2%>>&nbsp;<%=sDiv(1)%></FONT></TD>
				      	<TD align="Center" ><font size=<%=fontsize2%>>&nbsp;<%=sDiv(2)%></FONT></TD>
				      	<TD align="Center" ><font size=<%=fontsize2%>>&nbsp;<%=sDiv(3)%></FONT></TD>
				      	<TD align="Center" ><font size=<%=fontsize2%>>&nbsp;<%=sDiv(4)%></FONT></TD><%

		  	        IF sNoEmail="True" THEN  %>
										<TD align="Center" ><font size=<%=fontsize2%>>NO</FONT></TD><%
			         	ELSE %>	
										<TD>&nbsp;</TD><%
			         END IF	

		  	       IF sEmail<>"" THEN  %>
										<TD align="Center" ><font size=<%=fontsize2%>><%=sEmail%></FONT></TD><%
			         ELSE %>	
										<TD>&nbsp;</TD><%
			         END IF %>	
				      <TD align="Center" ><font size=<%=fontsize2%>>&nbsp;<%=ECount%></FONT></TD>
				      <TD align="Center" ><font size=<%=fontsize2%>>&nbsp;<%=sSentBioEmail%></FONT></TD>
				    </TR><%	

			    ELSE
							'response.write("INSIDE PONT 2b")
			    END IF

			
			' -----------------------------------------------------------------------------
			' --- View Samples of Email Letters or item to be displayed/printed in bulk ---
			' -----------------------------------------------------------------------------
			ELSEIF WhatNotify="ViewNotices" THEN

					'response.write("INSIDE PONT 3")
					'response.write("IN ELSEIF WhatLetter="&WhatLetter)

					SELECT CASE WhatLetter
						CASE "Bio"
								%><br><center><font size=5 font=<%=font1%> ><b>SAMPLE</b></font></center><br><%
								DisplayBioForm LastMemb, sTourID, sDiv(1), sDiv(2), sDiv(3), sDiv(4)
								EXIT DO

						CASE ELSE
								%><br><center><font size=5 font=<%=font1%> ><b>SAMPLE LETTER</b></font></center><br><%
								'markdebug("Inside")
								BuildMessageHTMLBody
								response.write(ebody)
								ebody=""
								EXIT DO
					END SELECT

				


			' ----------------------------
			' --- Send the next email
			' ----------------------------
		 	ELSEIF WhatNotify="SendList" THEN  
			
			    ' --- Checks whether bio form is complete ---
			    HowEmptyIsForm LastMemb, sTourID

			
			    ' --------------------------------------------------------------------------------------------
					' --- Letter must be the Bio Letter update
					' --- Checkbox for bio must be on
					' --- Analysis of missing field information must be greater than threshold
					' --- Member must not have NO EMAIL shown in Member file
					' --- Bio Letter must not have been previously sent or ResendEmail checkbox must be checked
			    ' --------------------------------------------------------------------------------------------
			    IF WhatLetter="reg_bioincomplete" AND sBioFilter="on" AND ECount >= MaxECount AND sNoEmail<>"True" AND (sSentBioEmail="N" OR (sSentBioEmail="Y" AND sReSendEmail="on")) THEN 	 
			
							' -------------------------------------------------------
							' --- Deploy email for the member from the record set ---
							' -------------------------------------------------------
							
							' SendTourEmail eMailSubj, eMailBody, eMailFrom, eMailBCC, sTest
							DeployEmailMessage


							' --- Updates the RegisterGenNew table with Email has been sent notice ---
							' --- Should be changed to a LastSentDate for several functions corresponding to type of notice or print run ---
							sSQL = "UPDATE "&RegGenTableName
							sSQL = sSQL + " SET SentBioEmail = 'Y'"
							sSQL = sSQL + " WHERE LEFT(TourID,6) = '"&LEFT(sTourID,6)&"' AND MemberID = '"&LastMemb&"'"
							con.execute(sSQL)


					' --- Custom letter or specific HQ letter from 2010 or so ---
					ELSEIF (WhatLetter="reg_westnile" OR WhatLetter="custom") AND sNoEmail<>"True" THEN

							' -------------------------------------------------------
							' --- Deploy email for the member from the record set ---
							' -------------------------------------------------------
							ebody=""
							
							eMailSubj
							'SendTourEmail eMailSubj, eMailBody, eMailFrom, eMailBCC, sTest
							DeployEmailMessage
					END IF



			' -----------------------------------------------------------
			' --- Write the next bio to the end of the display string
			' -----------------------------------------------------------
		 	ELSEIF WhatNotify="PrintBio" THEN  
					'response.write("INSIDE PONT 5")
					%><br><%

					DisplayBioForm LastMemb, sTourID, sDiv(1), sDiv(2), sDiv(3), sDiv(4)
			
			
			END IF

			'response.write("INSIDE PONT 6")
			sDiv(1)=""	
			sDiv(2)=""
			sDiv(3)=""	
			sDiv(4)=""	

		END IF

		'response.write("INSIDE PONT 7")

		IF NOT rs.eof THEN LastMemb=rs("MemberID")		  		    

	LOOP

	IF WhatNotify="ViewList" THEN %>
		</TABLE> <%

	ELSEIF WhatNotify="SendList" THEN
		
		%><br><br><center><font size=4 color="<% =Textcolor1 %>"><b><%=EmailCount%> - Emails Sent</b></center></font<%
	END IF

ELSE  %>

	<br>
	<center><font size=4 font=<%=font1%> ><b>No Data Meets Search Criteria</b></font></center>
	<br><%


END IF



END SUB




' ---------------------------
    SUB  DeployEmailMessage
' ---------------------------

	
	BuildMessageHTMLBody

	Set objCDO = Server.CreateObject("CDO.Message")

	ByPassEmails="NO"



	IF TRIM(sEmail)<>"" AND ByPassEmails <> "YES" THEN
			EmailCount=EmailCount+1			
	
			objCDO.To = sTourEmail&";"&sEmail
			objCDO.From = sTourEmail

			IF EmailCOunt=1 THEN
					objCDO.BCC = marksemailaddress
			END IF


			objCDO.Subject = "USA Waterski - Event Notification - "&sFullName
			objCDO.HTMLBody = ebody	

			ThisEmailTest="N"
			IF ThisEmailTest="Y" THEN
					response.write("sEmail="&sEmail)
					response.write("sTourEmail="&sTourEmail)
					'response.end
			END IF

			objCDO.Send
			Set objCDO = Nothing
			ebody=""
	END IF



END SUB



' ---------------------------
   SUB BuildMessageHTMLBody
' ---------------------------

	ebody = "<HTML><HEAD>"

	ebody = ebody & "<style>div.break {page-break-before:always}</style>" 

	ebody = ebody & "</HEAD><BODY>"

 	ebody = ebody & "<TABLE BORDER=4 align=CENTER CELLPADDING=5 CELLSPACING=0 BGcolor="&Tablecolor1&" width=75% >"


	' -----------------------------------------------------------------------------------------------
	' --- Reads and displays text from var (WhatLetter) file in communications folder ---
	' --- IF statements based on finding keywords on a line of file (ex: HEADLINE, EVENT, MEMBER)
	' -----------------------------------------------------------------------------------------------
	Set objfso = CreateObject("Scripting.FileSystemObject")
	IF objFSO.FileExists(PathToCommune & "\"&WhatLetter&".txt") THEN
			
			SET objstream=objFSO.opentextfile(PathToCommune & "\"&WhatLetter&".txt")

		  IF NOT objstream.atendofstream THEN
				DO WHILE not objstream.atendofstream
						currentline=" "&objstream.readline
						LenCurLine=Len(currentline)

						IF InStr(currentline, "+HEADLINE+") > 0 THEN
								headcolor=LEFT(RIGHT(currentline,LenCurLine-11),3)
								headcolor="red"
								ebody = ebody & "<TR>"
								IF LEFT(RIGHT(currentline,LenCurLine-11),3) ="RED" THEN
										ebody = ebody & "<TD BGcolor=red ><center><font face="&font1&" color=#FFFFFF size=4><b>"&RIGHT(currentline,LenCurLine-15)&"</b></font></TD>"
								ELSE
										ebody = ebody & "<TD BGcolor=blue ><center><font face="&font1&" color=#FFFFFF size=4><b>"&RIGHT(currentline,LenCurLine-15)&"</b></font></TD>"					
								END IF
	
								ebody = ebody & "</TR>"
								ebody = ebody & "<TR>"
								ebody = ebody & "<TD align=center Valign=top>"


						ELSEIF InStr(currentline, "+EVENT+") > 0 THEN
								ebody = ebody & "<br><br>"
								ebody = ebody & "<font face="&font1&" size=2><b>Events Entered</b></font>"
								ebody = ebody & "<br>"

								' --- Event names of up to four (4) events ---
								IF TRIM(sDiv(1)) <> "" THEN
										ebody = ebody & "<font color="&Textcolor2&" face="&font1&" size=2>"&sDiv(1)&" - "&sTEventName(1)&"</font>"
										ebody = ebody & "<br>"
								END IF
								IF TRIM(sDiv(2)) <> "" THEN
										ebody = ebody & "<font color="&Textcolor2&" face="&font1&" size=2>"&sDiv(2)&" - "&sTEventName(2)&"</font>"
										ebody = ebody & "<br>"
								END IF
								IF TRIM(sDiv(3)) <> "" THEN
										ebody = ebody & "<font color="&Textcolor2&" face="&font1&" size=2>"&sDiv(3)&" - "&sTEventName(3)&"</font>"
										ebody = ebody & "<br>"
								END IF
								IF TRIM(sDiv(4)) <> "" THEN
										ebody = ebody & "<font color="&Textcolor2&" face="&font1&" size=2>"&sDiv(4)&" - "&sTEventName(4)&"</font>"
										ebody = ebody & "<br>"
								END IF
		
						ELSEIF InStr(currentline, "+MEMBER+") > 0 THEN
								ebody = ebody & "<br>"
								ebody = ebody & "<font color="&Textcolor2&" face="&font1&" size=4><b>"&sFullName&"</b></font>"
								ebody = ebody & "<br>"
								ebody = ebody & "<font face="&font1&" size=2><b>MemberID = </b></font><font color="&Textcolor2&" face="&font1&" size=2>"&LastMemb&"</font>"

						ELSEIF InStr(currentline, "+TOURNAMENT+") > 0 THEN
								ebody = ebody & "<br>"
								ebody = ebody & "<font color=red face="&font1&" size=4><b>"&sTourName&"</b></font>"
								ebody = ebody & "<br>"
								ebody = ebody & "<font face="&font1&" size=2><b>SanctionID = </b></font><font color="&Textcolor2&" face="&font1&" size=2>"&sTourID&"</font>"
								ebody = ebody & "<br>"
								ebody = ebody & "<font face="&font1&" size=2><b>Date = </b></font><font color="&Textcolor2&" face="&font1&" size=2>"&sTDateS&" to "&sTDateE&"</font></b>"
								ebody = ebody & "<br><br>"


						ELSE
								ebody = ebody & "<br><font color="&Textcolor2&" face="&font1&" size=2>"&currentline&"</font>"
						END IF
				LOOP
			END IF
			objstream.close
	END IF


	ebody = ebody & "<br>"
	ebody = ebody & "</td></tr>"
	ebody = ebody & "</TABLE>"

	IF NOT rs.eof THEN
			ebody = ebody & "<div class=break />"
	END IF

	ebody = ebody & "</BODY></HTML>"

	' --- Bottom of 	
	

END SUB















' --------------------
  SUB LoadDropSequence
' --------------------

%>
  <td align=right>
	<font size=<% =fontsize3 %> color="<% =TextDropcolor2 %>">Order By:</font>
  </td>

  <td align=left>	
	<select name="SequenceSelected" style="width:11em">
	        <option value="alpha"<%IF SequenceSelected = "alpha" THEN Response.Write(" SELECTED ")%>>Alphabetic</option>
        	<option value="seed"<%IF SequenceSelected = "seed" THEN Response.Write(" SELECTED ")%>>Seeding Value</option>
	        <option value="regdate"<%IF SequenceSelected = "regdate" THEN Response.Write(" SELECTED ")%>>Register Date by Div</option>
	        <option value="regdateall"<%IF SequenceSelected = "regdateall" THEN Response.Write(" SELECTED ")%>>Register Date - All</option>
	</select>
  </td><%

END SUB


' ----------------------
   SUB LoadReportPullDown 
' ----------------------

%>
  <td align=right>
    <font size=<% =fontsize3 %> color="<% =TextDropcolor2 %>">Type:</font>
  </td>
  <td align=left>
     <select name="WhatReport" style="width:11em;">
			<option value="noreportselected"<%IF WhatReport = "noreportselected" THEN Response.Write(" SELECTED ")%>>Select Report</option>
      <option value="regstat"<%IF WhatReport = "regstat" THEN Response.Write(" SELECTED ")%>>Registration Status</option>
			<option value="seeding"<%IF WhatReport = "seeding" THEN Response.Write(" SELECTED ")%>>Seeding</option>
      <option value="scratched"<%IF WhatReport = "scratched" THEN Response.Write(" SELECTED ")%>>Not Ready To Ski</option>
			<% 
			IF adminmenulevel>=1 OR LCASE(Session("UserAdminPW"))=LCASE(Session("AdminCode")) THEN %> %>	
       	  <option value="skierpayments"<%IF WhatReport = "skierpayments" THEN Response.Write(" SELECTED ")%>>Payments by Type</option>
          <option value="financial"<%IF WhatReport = "financial" THEN Response.Write(" SELECTED ")%>>Payments Received</option>
					<option value="othersales" <%IF WhatReport = "othersales" THEN Response.Write(" SELECTED ")%>>Other Sales</option>
        	<% 
			END IF 
			IF adminmenulevel >= 19  THEN 
       	  %><option value="notifications"<%IF WhatReport = "notifications" THEN Response.Write(" SELECTED ")%>>Notifications</option><%
      END IF
			IF adminmenulevel >= 19 OR TestValidAdminCode=true  THEN 
					%><option value="notifications" <%IF WhatReport = "notifications" THEN Response.Write(" SELECTED ")%>>Print Bios</option><% 
			END IF 
			%>
			<option value="divisiontotals"<%IF WhatReport = "divisiontotals" THEN Response.Write(" SELECTED ")%>>Scheduling</option>
			<option value="bystate"<%IF WhatReport = "bystate" THEN Response.Write(" SELECTED ")%>>Skiers By State</option>
     </select>
  </td>
  <%


END SUB


'------------------
 SUB LoadDivPulldown
'------------------

' Loads applicable divisions into a division pulldown for each event selected

'    response.write("<br>sSptsGrpID = "&sSptsGrpID)

 SELECT CASE sSptsGrpID
  	CASE "AWS", "NCW"
		ThisDivTable = DivisionsTableName
	CASE "AKA", "USH", "USW", "ABC"
		ThisDivTable = DivisionsOtherTableName
 END SELECT


    opencon
    SET rsSelectFields=Server.CreateObject("ADODB.recordset")
    sSQL = "SELECT DISTINCT DT.div, DT.div_name FROM "&ThisDivTable&" AS DT"

    ' ///////  NOTE - Need to add filter to filte to current SkiYear

	

    SELECT CASE sSptsGrpID
  	CASE "AWS"
		sSQL = sSQL + " WHERE lower(left(DT.div,1)) <> 'i' AND lower(left(DT.div,1)) <> 'y' AND lower(left(DT.div,1)) <> 'x'"
		sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'n' AND lower(left(DT.div,1)) <> 'c'"
		sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'l' AND lower(left(DT.div,1)) <> 'e' AND lower(left(DT.div,1)) <> 's'"		
	CASE "AKA"
		sSQL = sSQL + " WHERE lower(left(DT.div,1)) = 'b' or lower(left(DT.div,1)) = 'g'"
	CASE "NCW"
		sSQL = sSQL + " WHERE lower(DT.div) = 'cm' or lower(DT.div) = 'cw'"
	CASE "USH"
		sSQL = sSQL + " WHERE SptsGrpID='USH'"
	CASE "ABC"
		sSQL = sSQL + " WHERE SptsGrpID='ABC'"

    END SELECT
    sSQL = sSQL + " order by DT.div"


'response.write("<br>"&sSQL)
'response.end
   	rsSelectFields.open sSQL, SConnectionToTRATable



%>
<td align=right>
  <font size=<% =fontsize3 %> color="<% =TextDropcolor2 %>">Division:</font>
</td>
<td align=left>
  <SELECT name="DivSelected" style="width:6em"><%


    IF NOT rsSelectFields.eof THEN 
  	rsSelectFields.movefirst

	Dim DivCounter
	DivCounter = 1

'SkipOver=false
'IF SkipOver=false THEN
	%>
	<option value ="ALL" <% IF DivSelected = "ALL" THEN Response.Write(" SELECTED ") %>>All</option><br>
  	<%
'END IF
  	
	DO WHILE NOT rsSelectFields.eof
		DivCounter = DivCounter + 1
		%>
		<option value = "<%=rsSelectFields("Div")%>"  <% IF DivSelected = rsSelectFields("Div") THEN Response.Write(" SELECTED ") %>> <%=rsSelectFields("Div")%></option><br>")
		<% 
		rsSelectFields.moveNEXT
	LOOP
    ELSE
	response.write("<option value =""None"" selected>None Available</option>")
    END IF 

    rsSelectFields.close  %>

  </select>
</td><%

END SUB



' ------------------------
   SUB LoadEventPulldown
' ------------------------

' ***************  OBSOLETE  ***************

Response.write("<br>Line EventSelected= "&EventSelected)
%>
<td align=right>
  <font size=<% =fontsize3 %> color="<% =TextDropcolor2 %>">Event:</font>
</td>
<td align=left>
<select name="EventSelected" style="width:6em">
<option value="ALL" <% IF EventSelected = "ALL" THEN Response.Write(" SELECTED ")%> >All</option><%

SELECT CASE sSptsGrpID
	CASE "AWS"
		%>
		<option value="S" <%IF EventSelected = "S" THEN Response.Write(" SELECTED ")%>>Slalom</option>
		<option value="T" <%IF EventSelected = "T" THEN Response.Write(" SELECTED ")%>>Tricks</option>
		<option value="J" <%IF EventSelected = "J" THEN Response.Write(" SELECTED ")%>>Jump</option>
		<%
	CASE "ABC"
		%>
		<option value="W" <%IF EventSelected = "W" THEN Response.Write(" SELECTED ")%>>Wake Cross</option>
		<option value="T" <%IF EventSelected = "T" THEN Response.Write(" SELECTED ")%>>Tricks</option>
		<option value="J" <%IF EventSelected = "J" THEN Response.Write(" SELECTED ")%>>Jump</option>
		<%
	CASE "USW"
		%>
		<option value="W" <%IF EventSelected = "W" THEN Response.Write(" SELECTED ")%>>Wakeboard</option>
		<option value="WS" <%IF EventSelected = "WS" THEN Response.Write(" SELECTED ")%>>Wake Skate</option>
		<option value="WU" <%IF EventSelected = "WU" THEN Response.Write(" SELECTED ")%>>Wake Surf</option>
		<%
	CASE "USH"
		%>
		<option value="W" <%IF EventSelected = "HB" THEN Response.Write(" SELECTED ")%>>Big Air</option>
		<option value="WS" <%IF EventSelected = "WS" THEN Response.Write(" SELECTED ")%>>Free Ride</option>
		<option value="WU" <%IF EventSelected = "WU" THEN Response.Write(" SELECTED ")%>>Jump Out</option>
		<%

	END SELECT
	%>
</select>
</td><%

END SUB



' ------------------------
   SUB LoadEventPulldownNew
' ------------------------

%>
<td align=right>
  <font size=<% =fontsize3 %> color="<% =TextDropcolor2 %>">Event:</font>
</td>
<td align=left>
<select name="EventSelected" style="width:6em">
<option value="ALL" <% IF EventSelected = "ALL" THEN Response.Write(" SELECTED ")%> >All</option><%

	FOR EvtNo = 1 TO TotEv 
		IF TRIM(sTEvent(EvtNo)) <> "" THEN %>
			<option value="<%=EventSelected%>" <% IF EventSelected = ""&EventSelected&"" THEN Response.Write(" SELECTED ")%>><%= EventSelected %></option><%
		END IF
	NEXT %>

</select>
</td><%

END SUB




' ------------------------
   SUB LoadRegionPulldown
' ------------------------

%>
<td align=right>
	<font size=<% =fontsize3 %> color="<% =TextDropcolor2 %>"><a TITLE="Region is based on State of Membership record.">Region:<a></font>
</td>
<td align=left>
  <select name="RegionSelected">
	<option value=""<%IF RegionSelected = "" THEN Response.Write(" SELECTED ")%>>All Regions</option>
	<option value="1"<%IF RegionSelected = "1" THEN Response.Write(" SELECTed ")%>>S. Central</option>
	<option value="2"<%IF RegionSelected = "2" THEN Response.Write(" SELECTED ")%>>Midwest</option>
	<option value="3"<%IF RegionSelected = "3" THEN Response.Write(" SELECTED ")%>>West</option>
	<option value="4"<%IF RegionSelected = "4" THEN Response.Write(" SELECTED ")%>>South</option>
	<option value="5"<%IF RegionSelected = "5" THEN Response.Write(" SELECTED ")%>>East</option>
  </select>
</td><%


END SUB


' ------------------------
   SUB LoadStatePulldown
' ------------------------

StateArray = Split(USStatesList2,",")  %>

<td align=right>
	<font size=<% =fontsize3 %> color="<% =Textcolor1 %>">State:</font>
</td>
<td align=left>
    <select name="StateSelected"><%
	  response.write("<option value = ""All"" SELECTED>All</option>")
	
	  FOR kvar = 0 TO UBOUND(StateArray)
	    IF TRIM(StateArray(kvar)) = "" THEN
		' - Blank do nothing	
	    ELSEIF TRIM(StateSelected) = TRIM(StateArray(kvar)) THEN
		response.write("<option value = """&StateSelected&""" SELECTED>"&StateSelected&"</option>")
	    ELSE
		response.write("<option value = """&StateArray(kvar)&""">"&StateArray(kvar)&"</option>")
	    END IF
	  NEXT  %>
    </select>
</td><%

END SUB




' ---------------------------------
   SUB LoadStartPulldown
' ---------------------------------

AlphaArray = Split(AlphaList,",")  %>
<td align="right">
	<font size=<% =fontsize3 %> color="<% =Textcolor1 %>">Start With:</font>
</td>
<td align="left">
	<select name="StartPulldown"><%
	    response.write("<option value = ""All"" SELECTED>All</option>")
	
	  FOR kvar = 0 TO UBOUND(AlphaArray)
	    IF TRIM(AlphaArray(kvar)) = "" THEN
		' - Blank do nothing	
	    ELSEIF TRIM(StartCharSelected) = TRIM(AlphaArray(kvar)) THEN
		response.write("<option value = """&StartCharSelected&""" SELECTED>"&StartCharSelected&"</option>")
	    ELSE
		response.write("<option value = """&AlphaArray(kvar)&""">"&AlphaArray(kvar)&"</option>")
	    END IF
	  NEXT  %>
	</select>
</td><%

END SUB


' ---------------------------------
   SUB LoadEndPulldown
' ---------------------------------

AlphaArray = Split(AlphaList,",")  %>
<td align="right">
	<font size=<% =fontsize3 %> color="<% =Textcolor1 %>">End With:</font>
</td>
<td align="left">
	 <select name="EndPulldown"><%
	    response.write("<option value = ""All"" SELECTED>All</option>")
	
	  FOR kvar = 0 TO UBOUND(AlphaArray)
	    IF TRIM(AlphaArray(kvar)) = "" THEN
		' - Blank do nothing	
	    ELSEIF TRIM(StartCharSelected) = TRIM(AlphaArray(kvar)) THEN
		response.write("<option value = """&EndCharSelected&""" SELECTED>"&EndCharSelected&"</option>")
	    ELSE
		response.write("<option value = """&AlphaArray(kvar)&""">"&AlphaArray(kvar)&"</option>")
	    END IF
	  NEXT  %>
</select>
</td><%

END SUB





' ------------------------------------------------------------
  SUB LoadTimePulldown (TimeAmount, MinTime, MaxTime, StepTime)
' ------------------------------------------------------------

Dim iCounter

TimeAmount = Cint(TimeAmount)

'response.write("<option value = 0 >NA</option>")

FOR iCounter = MinTime TO MaxTime STEP StepTime

	mymin=Fix(iCounter/60)

	mysec=iCounter - 60*mymin
	IF cdbl(mysec) = 0 THEN mySec = "00"
	myMinSec=mymin&":"&mysec

	IF iCounter = TimeAmount THEN
		response.write("<option value = """&iCounter&""" SELECTED>"&myMinSec&"</option>")
	ELSE
		response.write("<option value = """&iCounter&""">"&myMinSec&"</option>")
	END IF
NEXT


END SUB


%>




