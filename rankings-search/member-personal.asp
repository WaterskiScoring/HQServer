<!--#include file="settingsHQ.asp"-->

<%

' ---------------------------------------------------------------------------------------------
' --- This module is the start of a MEMBER PERSONAL PAGE
' --- Original module created by Mark Crone
' --- LAST updated: 7/3/2007  
' ---------------------------------------------------------------------------------------------


Dim RunByReport, RunByWhat, RegionSelected, EventSelected, DivSelected
Dim sTourID, sMemberID, SkiYearID, sSptsGrpID
Dim sTDateE, sTDateS, sTName 


sTourID="07W999A"


' ---------------------------------------------
' --- Read variables from form
' ---------------------------------------------

adminmenulevel = Session("adminmenulevel")
IF adminmenulevel = "" THEN adminmenulevel = 0

sMemberID = TRIM(Request("sMemberID"))

RegionSelected = trim(Request("RegionSelected"))
EventSelected = trim(Request("EventSelected"))
DivSelected = trim(Request("DivSelected"))
WhatReport = trim(Request("WhatReport"))
SequenceSelected = trim(Request("SequenceSelected"))



' --- Sets Default values for report  ---
IF TRIM(Request("DivSelected")) = "" THEN DivSelected = "ALL"
IF TRIM(RegionSelected) = "" THEN RegionSelected = 6
IF TRIM(SkiYearSelected) = "" THEN SkiYearSelected = 1

' ---- This will need a condition depending on which sports division  ----
IF TRIM(Request("EventSelected")) = "" THEN EventSelected = "ALL"
IF SequenceSelected = "" THEN SequenceSelected = "Seed"
IF WhatReport = "" THEN WhatReport = "Seeding"


' ------------------------------------
' Reads info from Sanction file
' ------------------------------------

Set rsTour=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT TOP 1 * FROM "&SanctionTableName&" AS ST WHERE LEFT(ST.TSanction,6) = '"&LEFT(sTourID,6)&"'"
rsTour.open sSQL, sConnectionToTRATable, 3, 1

sTDateS = rsTour("TDateS")
sTDateE = rsTour("TDateE")
sTName = rsTour("TName")
sSptsGrpID = rsTour("SptsGrpID")

rsTour.close






' ------------------------------------------------------------------------------------------------------------
' ----------------    Draws the filtering objects on the screen based on report type selected   --------------
' ------------------------------------------------------------------------------------------------------------


RunByWhat = "TEST"

SELECT CASE RunByWhat

   CASE "TEST"



END SELECT




' ----------------------------------------------------------------------------------------------------------
' -----------  Builds SQL string to define display values  -------------------------------------------------
' ----------------------------------------------------------------------------------------------------------


'markdebug("TRIM(sMemberID) = "&TRIM(sMemberID))

IF TRIM(sMemberID)="" THEN

	' --- Sends user to search-member routine to selected member
	Session("sSendingPage")="/member-personal.asp?pvar=FoundMember"
	Response.Redirect("/search-member.asp?rid="&rid&"&formstatus=search")
END IF


SET rs=Server.CreateObject("ADODB.recordset")	
sSQL = "SELECT * FROM "&MemberTableName
sSQL = sSQL + " WHERE PersonIDWithCheckDigit = '"&sMemberID&"'"
rs.open sSQL, sConnectionToTRATable, 3, 1

IF NOT rs.eof THEN 
	sFullName = rs("FirstName")&" "&rs("LastName")	  



Event1="S"
Event2="T"
Event3="J"
Event4="O"


Event1Name="Slalom"
Event2Name="Tricks"
Event3Name="Jump"
Event4Name="Overall"

'Div1="M4"
sProcessingYear=1

E12Date=DATE
B12Date=DateAdd("yyyy",-1,DATE)

BLastDate="08/18/2006"
ELastDate=DateAdd("yyyy",1,BLastDate)

BPrevDate="08/18/2005"
EPrevDate=DateAdd("yyyy",1,BPrevDate)


'MarkDebug(B12Date)
'MarkDebug(E12Date)
'MarkDebug(BLastDate)
'MarkDebug(ELastDate)
'MarkDebug(BPrevDate)
'MarkDebug(EPrevDate)

LastSkiYear="2006"
PrevSkiYear="2005"


' --- Coalesce to -1 to indicate when a value does not exist for this query.

SET rs=Server.CreateObject("ADODB.recordset")
sSQL = "  SELECT Coalesce(Ev1.Ev1_Cnt,-1) AS SL12Cnt, Coalesce(Ev1.Ev1_Max,-1) AS SL12Max, Coalesce(Ev1.Ev1_Min,-1) AS SL12Min"
sSQL = sSQL + ", Coalesce(Ev2.Ev2_Cnt,-1) AS TR12Cnt, Coalesce(Ev2.Ev2_Max,-1) AS TR12Max, Coalesce(Ev2.Ev2_Min,-1) AS TR12Min"
sSQL = sSQL + ", Coalesce(Ev3.Ev3_Cnt,-1) AS JU12Cnt, Coalesce(Ev3.Ev3_Max,-1) AS JU12Max, Coalesce(Ev3.Ev3_Min,-1) AS JU12Min"
sSQL = sSQL + ", Coalesce(Ev4.Ev4_Cnt,-1) AS OV12Cnt, Coalesce(Ev4.Ev4_Max,-1) AS OV12Max, Coalesce(Ev4.Ev4_Min,-1) AS OV12Min"

sSQL = sSQL + ", Coalesce(Ev5.Ev1_Cnt,-1) AS SLLastCnt, Coalesce(Ev5.Ev1_Max,-1) AS SLLastMax, Coalesce(Ev5.Ev1_Min,-1) AS SLLastMin"
sSQL = sSQL + ", Coalesce(Ev6.Ev2_Cnt,-1) AS TRLastCnt, Coalesce(Ev6.Ev2_Max,-1) AS TRLastMax, Coalesce(Ev6.Ev2_Min,-1) AS TRLastMin"
sSQL = sSQL + ", Coalesce(Ev7.Ev3_Cnt,-1) AS JULastCnt, Coalesce(Ev7.Ev3_Max,-1) AS JULastMax, Coalesce(Ev7.Ev3_Min,-1) AS JULastMin"
sSQL = sSQL + ", Coalesce(Ev8.Ev4_Cnt,-1) AS OVLastCnt, Coalesce(Ev8.Ev4_Max,-1) AS OVLastMax, Coalesce(Ev8.Ev4_Min,-1) AS OVLastMin"

sSQL = sSQL + ", Coalesce(Ev9.Ev1_Cnt,-1) AS SLPrevCnt, Coalesce(Ev5.Ev1_Max,-1) AS SLPrevMax, Coalesce(Ev5.Ev1_Min,-1) AS SLPrevMin"
sSQL = sSQL + ", Coalesce(Ev10.Ev2_Cnt,-1) AS TRPrevCnt, Coalesce(Ev6.Ev2_Max,-1) AS TRPrevMax, Coalesce(Ev10.Ev2_Min,-1) AS TRPrevMin"
sSQL = sSQL + ", Coalesce(Ev11.Ev3_Cnt,-1) AS JUPrevCnt, Coalesce(Ev7.Ev3_Max,-1) AS JUPrevMax, Coalesce(Ev11.Ev3_Min,-1) AS JUPrevMin"
sSQL = sSQL + ", Coalesce(Ev12.Ev4_Cnt,-1) AS OVPrevCnt, Coalesce(Ev12.Ev4_Max,-1) AS OVPrevMax, Coalesce(Ev12.Ev4_Min,-1) AS OVPrevMin"

sSQL = sSQL + ", Coalesce(Tour1.NoTour1_Cnt,-1) AS SL12TourCnt, Coalesce(Tour2.NoTour2_Cnt,-1) AS TR12TourCnt, Coalesce(Tour3.NoTour3_Cnt,-1) AS JU12TourCnt, Coalesce(Tour4.NoTour4_Cnt,-1) AS OV12TourCnt"
sSQL = sSQL + ", Coalesce(Tour5.NoTour1_Cnt,-1) AS SLLastTourCnt, Coalesce(Tour6.NoTour2_Cnt,-1) AS TRLastTourCnt, Coalesce(Tour7.NoTour3_Cnt,-1) AS JULastTourCnt, Coalesce(Tour8.NoTour4_Cnt,-1) AS OVLastTourCnt"
sSQL = sSQL + ", Coalesce(Tour9.NoTour1_Cnt,-1) AS SLPrevTourCnt, Coalesce(Tour10.NoTour2_Cnt,-1) AS TRPrevTourCnt, Coalesce(Tour11.NoTour3_Cnt,-1) AS JUPrevTourCnt, Coalesce(Tour12.NoTour4_Cnt,-1) AS OVPrevTourCnt"



sSQL = sSQL + " FROM " 
sSQL = sSQL + "  (SELECT COUNT(Score) AS Ev1_CNT, MAX(Score) AS Ev1_Max, Min(Score) AS Ev1_Min FROM "&RawScoresTableName&" WHERE MemberID='"&sMemberID&"' AND Event='"&Event1&"' AND EndDate BETWEEN '"&B12Date&"' and '"&E12Date&"') AS Ev1" 
sSQL = sSQL + ", (SELECT COUNT(Score) AS Ev2_CNT, MAX(Score) AS Ev2_Max, Min(Score) AS Ev2_Min FROM "&RawScoresTableName&" WHERE MemberID='"&sMemberID&"' AND Event='"&Event2&"' AND EndDate BETWEEN '"&B12Date&"' and '"&E12Date&"') AS Ev2" 
sSQL = sSQL + ", (SELECT COUNT(Score) AS Ev3_CNT, MAX(Score) AS Ev3_Max, Min(Score) AS Ev3_Min FROM "&RawScoresTableName&" WHERE MemberID='"&sMemberID&"' AND Event='"&Event3&"' AND EndDate BETWEEN '"&B12Date&"' and '"&E12Date&"') AS Ev3" 
sSQL = sSQL + ", (SELECT COUNT(TotalOverall) AS Ev4_CNT, MAX(TotalOverall) AS Ev4_Max, Min(TotalOverall) AS Ev4_Min FROM "&OverallScoresTableName&" WHERE MemberID='"&sMemberID&"' AND SkiYearID='1') AS Ev4" 

sSQL = sSQL + ", (SELECT COUNT(Score) AS Ev1_CNT, MAX(Score) AS Ev1_Max, Min(Score) AS Ev1_Min FROM "&RawScoresTableName&" WHERE MemberID='"&sMemberID&"' AND Event='"&Event1&"' AND EndDate BETWEEN '"&BLastDate&"' and '"&ELastDate&"') AS Ev5" 
sSQL = sSQL + ", (SELECT COUNT(Score) AS Ev2_CNT, MAX(Score) AS Ev2_Max, Min(Score) AS Ev2_Min FROM "&RawScoresTableName&" WHERE MemberID='"&sMemberID&"' AND Event='"&Event2&"' AND EndDate BETWEEN '"&BLastDate&"' and '"&ELastDate&"') AS Ev6" 
sSQL = sSQL + ", (SELECT COUNT(Score) AS Ev3_CNT, MAX(Score) AS Ev3_Max, Min(Score) AS Ev3_Min FROM "&RawScoresTableName&" WHERE MemberID='"&sMemberID&"' AND Event='"&Event3&"' AND EndDate BETWEEN '"&BLastDate&"' and '"&ELastDate&"') AS Ev7" 
sSQL = sSQL + ", (SELECT COUNT(TotalOverall) AS Ev4_CNT, MAX(TotalOverall) AS Ev4_Max, Min(TotalOverall) AS Ev4_Min FROM "&OverallScoresTableName&" WHERE MemberID='"&sMemberID&"' AND SkiYearID='9') AS Ev8" 

sSQL = sSQL + ", (SELECT COUNT(Score) AS Ev1_CNT, MAX(Score) AS Ev1_Max, Min(Score) AS Ev1_Min FROM "&RawScoresTableName&" WHERE MemberID='"&sMemberID&"' AND Event='"&Event1&"' AND EndDate BETWEEN '"&BPrevDate&"' and '"&EPrevDate&"') AS Ev9" 
sSQL = sSQL + ", (SELECT COUNT(Score) AS Ev2_CNT, MAX(Score) AS Ev2_Max, Min(Score) AS Ev2_Min FROM "&RawScoresTableName&" WHERE MemberID='"&sMemberID&"' AND Event='"&Event2&"' AND EndDate BETWEEN '"&BPrevDate&"' and '"&EPrevDate&"') AS Ev10" 
sSQL = sSQL + ", (SELECT COUNT(Score) AS Ev3_CNT, MAX(Score) AS Ev3_Max, Min(Score) AS Ev3_Min FROM "&RawScoresTableName&" WHERE MemberID='"&sMemberID&"' AND Event='"&Event3&"' AND EndDate BETWEEN '"&BPrevDate&"' and '"&EPrevDate&"') AS Ev11" 
sSQL = sSQL + ", (SELECT COUNT(TotalOverall) AS Ev4_CNT, MAX(TotalOverall) AS Ev4_Max, Min(TotalOverall) AS Ev4_Min FROM "&OverallScoresTableName&" WHERE MemberID='"&sMemberID&"' AND SkiYearID='8') AS Ev12" 

sSQL = sSQL + ", (SELECT COUNT(DISTINCT TourID) AS NoTour1_CNT FROM "&RawScoresTableName&" WHERE MemberID='"&sMemberID&"' AND Event='"&Event1&"' AND EndDate BETWEEN '"&B12Date&"' and '"&E12Date&"') AS Tour1" 
sSQL = sSQL + ", (SELECT COUNT(DISTINCT TourID) AS NoTour2_CNT FROM "&RawScoresTableName&" WHERE MemberID='"&sMemberID&"' AND Event='"&Event2&"' AND EndDate BETWEEN '"&B12Date&"' and '"&E12Date&"') AS Tour2" 
sSQL = sSQL + ", (SELECT COUNT(DISTINCT TourID) AS NoTour3_CNT FROM "&RawScoresTableName&" WHERE MemberID='"&sMemberID&"' AND Event='"&Event3&"' AND EndDate BETWEEN '"&B12Date&"' and '"&E12Date&"') AS Tour3" 
sSQL = sSQL + ", (SELECT COUNT(DISTINCT TourID) AS NoTour4_CNT FROM "&OverallScoresTableName&" WHERE MemberID='"&sMemberID&"' AND SkiYearID='1') AS Tour4" 

sSQL = sSQL + ", (SELECT COUNT(DISTINCT TourID) AS NoTour1_CNT FROM "&RawScoresTableName&" WHERE MemberID='"&sMemberID&"' AND Event='"&Event1&"' AND EndDate BETWEEN '"&BLastDate&"' and '"&ELastDate&"') AS Tour5" 
sSQL = sSQL + ", (SELECT COUNT(DISTINCT TourID) AS NoTour2_CNT FROM "&RawScoresTableName&" WHERE MemberID='"&sMemberID&"' AND Event='"&Event2&"' AND EndDate BETWEEN '"&BLastDate&"' and '"&ELastDate&"') AS Tour6" 
sSQL = sSQL + ", (SELECT COUNT(DISTINCT TourID) AS NoTour3_CNT FROM "&RawScoresTableName&" WHERE MemberID='"&sMemberID&"' AND Event='"&Event3&"' AND EndDate BETWEEN '"&BLastDate&"' and '"&ELastDate&"') AS Tour7" 
sSQL = sSQL + ", (SELECT COUNT(DISTINCT TourID) AS NoTour4_CNT FROM "&OverallScoresTableName&" WHERE MemberID='"&sMemberID&"' AND SkiYearID='9') AS Tour8" 

sSQL = sSQL + ", (SELECT COUNT(DISTINCT TourID) AS NoTour1_CNT FROM "&RawScoresTableName&" WHERE MemberID='"&sMemberID&"' AND Event='"&Event1&"' AND EndDate BETWEEN '"&BPrevDate&"' and '"&EPrevDate&"') AS Tour9" 
sSQL = sSQL + ", (SELECT COUNT(DISTINCT TourID) AS NoTour2_CNT FROM "&RawScoresTableName&" WHERE MemberID='"&sMemberID&"' AND Event='"&Event2&"' AND EndDate BETWEEN '"&BPrevDate&"' and '"&EPrevDate&"') AS Tour10" 
sSQL = sSQL + ", (SELECT COUNT(DISTINCT TourID) AS NoTour3_CNT FROM "&RawScoresTableName&" WHERE MemberID='"&sMemberID&"' AND Event='"&Event3&"' AND EndDate BETWEEN '"&BPrevDate&"' and '"&EPrevDate&"') AS Tour11" 
sSQL = sSQL + ", (SELECT COUNT(DISTINCT TourID) AS NoTour4_CNT FROM "&OverallScoresTableName&" WHERE MemberID='"&sMemberID&"' AND SkiYearID='8') AS Tour12" 

rs.open sSQL, sConnectionToTRATable, 3, 1	



'SET rsRT=Server.CreateObject("ADODB.recordset")
'sSQL = "SELECT Case RT.RankScore When SkiYearID=1 Then SL12Rank=RT.RankScore When SkiYearID=9 Then SLLastRank=RT.RankScore When SkiYearID=8 Then SLPrevRank=RT.RankScore END"
'sSQL = sSQL + " FROM "&RankTableName&" AS RT WHERE RT.MemberID='"&sMemberID&"'"
'rsRT.open sSQL, sConnectionToTRATable, 3, 1	







' --------------------------------------------------------------------------------------------------------
' ------------------------------- This section prints the headers on the reports -------------------------
' --------------------------------------------------------------------------------------------------------















WhatTab=TRIM(Request("WhatTab"))
IF WhatTab="" THEN WhatTab=1

SELECT CASE WhatTab
	CASE "1"
		Tab1Color="/images/buttons/Vertical_Shade_216x20.jpg"
		TabText1Color="#FFFFFF"
	CASE "2"
		Tab2Color="/images/buttons/Vertical_Shade_216x20.jpg"
		TabText2Color="#FFFFFF"
	CASE "3"
		Tab3Color="/images/buttons/Vertical_Shade_216x20.jpg"
		TabText3Color="#FFFFFF"
	CASE "4"
		Tab4Color="/images/buttons/Vertical_Shade_216x20.jpg"
		TabText4Color="#FFFFFF"

END SELECT


WriteIndexPageHeader
ReportTitle = "Personal Stats - "&sFullName


%>
<TABLE BORDER="0" CELLPADDING="6" CELLSPACING="0" WIDTH="100%" BORDERCOLOR="#C0C0C0" BORDERCOLORDARK="#C0C0C0" BORDERCOLORLIGHT="#C0C0C0" >
  <TR>
	<TD align="center" vAlign=bottom noWrap background="/images/buttons/Vertical_Shade_564x152_New.jpg">
		    <FONT face="Verdana, Arial, Helvetica, sans-serif" color=#ffffff size=5>
		    <B><% Response.Write(ReportTitle) %></B></FONT>&nbsp
	</TD>	
  </TR>
</TABLE>



<br>
<TABLE ALIGN="Center" BORDER="1" CELLPADDING="3" CELLSPACING="0" BGCOLOR="<%=TableColor1%>" width=100%>

  <TR>	

	<td WIDTH=25% ALIGN=center noWrap background=<%=Tab1Color%>><a href="/member-personal.asp?WhatTab=1&sMemberID=<%=sMemberID%>"><font size=<% =fontsize3 %> face=<% =font1 %> COlOR="<%=TabText1Color%>">Rankings & Performances</FONT></a></td>
	<td WIDTH=25% ALIGN=center noWrap background=<%=Tab2Color%>><a href="/member-personal.asp?WhatTab=2&sMemberID=<%=sMemberID%>"><font size=<% =fontsize3 %> face=<% =font1 %> COlOR="<%=TabText2Color%>">Score Details</FONT></a></td>
	<td WIDTH=25% ALIGN=center noWrap background="<%=Tab3Color%>"><a href="/member-personal.asp?WhatTab=3&sMemberID=<%=sMemberID%>"><font size=<% =fontsize3 %> face=<% =font1 %> COlOR=<%=TabText3Color%>>My Competition</FONT></a></td>
	<td WIDTH=25% ALIGN=center noWrap background="<%=Tab4Color%>"><a href="/member-personal.asp?WhatTab=4&sMemberID=<%=sMemberID%>"><font size=<% =fontsize3 %> face=<% =font1 %> COlOR="<%=TabText4Color%>">Future Tournaments</FONT></a></td>


  </TR>

</TABLE>
<%

SELECT CASE WhatTab

   CASE "1"

	%>
	<br>
        <TABLE ALIGN="Center" BORDER="1" CELLPADDING="3" CELLSPACING="0" BGCOLOR="<%=TableColor1%>" width=100%>
	  <TR>	
	    <TD bgcolor="<%=HeadColor1%>">&nbsp;</TD>
      	    <TD colspan=3 ALIGN="Center" vAlign="top" bgcolor="<%=HeadColor1%>"><font size=<% =fontsize3 %> face=<% =font1 %> COlOR="<% =TextColor1 %>"><b><%=Event1Name%></b></FONT></TD>
      	    <TD colspan=3 ALIGN="Center" vAlign="top" bgcolor="<%=HeadColor1%>"><font size=<% =fontsize3 %> face=<% =font1 %> COlOR="<% =TextColor1 %>"><b><%=Event2Name%></b></FONT></TD>
      	    <TD colspan=3 ALIGN="Center" vAlign="top" bgcolor="<%=HeadColor1%>"><font size=<% =fontsize3 %> face=<% =font1 %> COlOR="<% =TextColor1 %>"><b><%=Event3Name%></b></FONT></TD>
      	    <TD colspan=3 ALIGN="Center" vAlign="top" bgcolor="<%=HeadColor1%>"><font size=<% =fontsize3 %> face=<% =font1 %> COlOR="<% =TextColor1 %>"><b><%=Event4Name%></b></FONT></TD>
	  </TR>

	  <TR>	
	    <TD bgcolor="<%=TableColor1%>">&nbsp;</TD>
      	    <TD ALIGN="Center" vAlign="top" bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">Last12</FONT></TD>
      	    <TD ALIGN="Center" vAlign="top" bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>"><%=LastSkiYear%></FONT></TD>
      	    <TD ALIGN="Center" vAlign="top" bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>"><%=PrevSkiYear%></FONT></TD>

      	    <TD ALIGN="Center" vAlign="top" bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">Last12</FONT></TD>
      	    <TD ALIGN="Center" vAlign="top" bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>"><%=LastSkiYear%></FONT></TD>
      	    <TD ALIGN="Center" vAlign="top" bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>"><%=PrevSkiYear%></FONT></TD>

      	    <TD ALIGN="Center" vAlign="top" bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">Last12</FONT></TD>
      	    <TD ALIGN="Center" vAlign="top" bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>"><%=LastSkiYear%></FONT></TD>
      	    <TD ALIGN="Center" vAlign="top" bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>"><%=PrevSkiYear%></FONT></TD>

      	    <TD ALIGN="Center" vAlign="top" bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">Last12</FONT></TD>
      	    <TD ALIGN="Center" vAlign="top" bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>"><%=LastSkiYear%></FONT></TD>
      	    <TD ALIGN="Center" vAlign="top" bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>"><%=PrevSkiYear%></FONT></TD>
	  </TR>

	  <TR>	
      	    <TD ALIGN="left" vAlign="top" align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize3 %> face=<% =font1 %> COlOR="<% =TextColor1 %>"><b>Participation</b></FONT></TD>
	    <TD colspan=3 align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD colspan=3 align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD colspan=3 align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD colspan=3 align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	  </TR>

	  <TR>	
      	    <TD ALIGN="Center" vAlign="top" bgcolor="<%=TableColor1%>"><font size=<% =fontsize3 %> face=<% =font1 %> COlOR="<% =TextColor1 %>"># of Scores</FONT></TD>

	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("SL12CNT"),0)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("SLLastCNT"),0)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("SLPrevCNT"),0)%></font></TD>

	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("TR12CNT"),0)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("TRLastCNT"),0)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("TRPrevCNT"),0)%></font></TD>

	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("JU12CNT"),0)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("JULastCNT"),0)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("JUPrevCNT"),0)%></font></TD>

	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("OV12CNT"),0)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("OVLastCNT"),0)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("OVPrevCNT"),0)%></font></TD>
	  </TR>

	  <TR>	
      	    <TD ALIGN="Center" vAlign="top" align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize3 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">Tournaments</FONT></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("SL12TourCnt"),0)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("SLLastTourCnt"),0)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("SLPrevTourCnt"),0)%></font></TD>

	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("TR12TourCnt"),0)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("TRLastTourCnt"),0)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("TRPrevTourCnt"),0)%></font></TD>

	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("JU12TourCnt"),0)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("JULastTourCnt"),0)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("JUPrevTourCnt"),0)%></font></TD>

	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("OV12TourCnt"),0)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("OVLastTourCnt"),0)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("OVPrevTourCnt"),0)%></font></TD>

	  </TR>

	  <TR>	
      	    <TD ALIGN="left" vAlign="top" align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize3 %> face=<% =font1 %> COlOR="<% =TextColor1 %>"><b>Performances</b></FONT></TD>
	    <TD colspan=3 align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD colspan=3 align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD colspan=3 align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD colspan=3 align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	  </TR>

	  <TR>	
      	    <TD ALIGN="Center" vAlign="top" align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize3 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">Best</FONT></TD>

	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("SL12MAX"),2)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("SLLastMAX"),2)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("SLPrevMAX"),2)%></font></TD>

	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("TR12MAX"),0)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("TRLastMAX"),0)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("TRPrevMAX"),0)%></font></TD>

	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("JU12MAX"),1)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("JULastMAX"),1)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("JUPrevMAX"),1)%></font></TD>

	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("OV12MAX"),0)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("OVLastMAX"),0)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("OVPrevMAX"),0)%></font></TD>
	  </TR>

	  <TR>	
      	    <TD ALIGN="Center" vAlign="top" align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize3 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">Avg/Med</FONT></TD>

	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>

	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>

	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>

	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	  </TR>


	  <TR>	
      	    <TD ALIGN="Center" vAlign="top" align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize3 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">Worst</FONT></TD>

	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("SL12Min"),2)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("SLLastMin"),2)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("SLPrevMin"),2)%></font></TD>

	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("TR12Min"),0)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("TRLastMin"),0)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("TRPrevMin"),0)%></font></TD>

	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("JU12Min"),1)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("JULastMin"),1)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("JUPrevMin"),1)%></font></TD>

	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("OV12Min"),0)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("OVLastMin"),0)%></font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=formatNumber(rs("OVPrevMin"),0)%></font></TD>
	  </TR>


	  <TR>	
      	    <TD ALIGN="left" vAlign="top" align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize3 %> face=<% =font1 %> COlOR="<% =TextColor1 %>"><b>Ranking</b></FONT></TD>
	    <TD colspan=3 align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD colspan=3 align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD colspan=3 align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD colspan=3 align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	  </TR>

	  <TR>	
      	    <TD ALIGN="Center" vAlign="top" align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize3 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">Division 1 </FONT></TD>

	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>

	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>

	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>

	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	  </TR>

	  <TR>	
      	    <TD ALIGN="Center" vAlign="top" align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize3 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">Division 2 </FONT></TD>

	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>

	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>

	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>

	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	    <TD align=center bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
	  </TR>




	</TABLE>	
	<%

  CASE "2"

	%>
	<br>
        <TABLE ALIGN="Center" BORDER="1" CELLPADDING="3" CELLSPACING="0" BGCOLOR="<%=TableColor1%>" width=100%>
	  <TR>	
	    <TD bgcolor="<%=HeadColor1%>"><font size=<% =fontsize3 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
      	    <TD colspan=4 ALIGN="Center" vAlign="top" bgcolor="<%=HeadColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>"><b>Current Year</b></FONT></TD>
      	    <TD colspan=4 ALIGN="Center" vAlign="top" bgcolor="<%=HeadColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>"><b>Last Year</b></FONT></TD>
      	    <TD colspan=4 ALIGN="Center" vAlign="top" bgcolor="<%=HeadColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>"><b>2 Years Ago</b></FONT></TD>
	  </TR>
	</TABLE>	
	<%

  CASE "4"

	%>
	<br>
        <TABLE ALIGN="Center" BORDER="1" CELLPADDING="3" CELLSPACING="0" BGCOLOR="<%=TableColor1%>" width=100%>
	  <TR>	
	    <TD bgcolor="<%=HeadColor1%>"><font size=<% =fontsize3 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;</font></TD>
      	    <TD ALIGN="Center" vAlign="top" bgcolor="<%=HeadColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>"><b>Current Year</b></FONT></TD>
      	    <TD ALIGN="Center" vAlign="top" bgcolor="<%=HeadColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>"><b>Last Year</b></FONT></TD>
      	    <TD ALIGN="Center" vAlign="top" bgcolor="<%=HeadColor1%>"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>"><b>2 Years Ago</b></FONT></TD>
	  </TR>
	</TABLE>	
	<%


END SELECT



NewsPageNum="10RegRep"
WriteIndexPageFooter
	

rs.Close


ELSE  ' ---- If no name is found
	sFullName = "NameNotFound"
END IF




' ------------------------------------------------------------------------------------------------------------------------------
' ---------------------   END OF MAIN SECTION OF PROGRAM    --------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------------------









%>










