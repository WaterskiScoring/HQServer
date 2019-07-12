<!--#include file="settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->
<!--#include virtual="/rankings/tools_Definitions.asp"-->
<!--#include virtual="/rankings/tools_leagues.asp"-->
<!--#include virtual="/rankings/tools_mobile_version.asp"-->
<%



Dim ThisFileName

Dim sFullName, sRankDiv, sRankEvent, sRankTourID
Dim sRankScoreThis, sRankScoreLast, sNat_Qual, sRegl_Plc, sNatl_Plc
Dim sSummaryMetric, sSummaryEvent, sLast60dSummaryValue, sP60to120SummaryValue, sTotalSummaryValue
Dim MetricCount, EventCount, TourCount
Dim ThisYear12moStart, ThisYear12moEnd, LastYear12moStart, LastYear12moEnd
Dim sRawTournament, sRawTDateE, sRawEvent, sRawDiv, sRawPlace, sRawRound, sRawScore
Dim StatsMemberID, sWatchMemberIDs, SelectedMemberID
Dim sMembBirth, sMembAge, sMemberFed, sState

Dim UserMemberID, sName, sNum_Forwards, sCreated_Date, sModified_Date




ThisFileName = "mobile_monitor.asp"

' --- Date Range calcs for last 12 months This and Last --
ThisYear12moStart = DateAdd("d",-364,Date)
ThisYear12moEnd = Date
LastYear12moStart = DateAdd("d",-729,Date)
LastYear12moEnd = DateAdd("d",-365,Date)


sMemberID = TRIM(Request("sMemberID"))
StatsMemberID = TRIM(Request("StatsMemberID"))
sWatchMemberIDs = Request("sWatchMemberIDs")

' SelectedMemberID=sMemberID
IF StatsMemberID<>"" THEN SelectedMemberID=StatsMemberID ELSE SelectedMemberID=sMemberID


'response.write("<br>TEST")
'response.write("<br>sMemberID = "&sMemberID)
' response.write("<br>SelectedMemberID = "&SelectedMemberID)
'response.end 

OpenState="mystats"
DisplayHeadOpenBodyAndBannerTags OpenState


'IF sMemberID<>"" THEN 
		' --- Assembles and displays all the stats ---
		GetAllUserAdoptionStats
'ELSE
		' --- Display No User Set screen
'		DisplayNoUserSetScreen
'END IF	




' --- Writes the Closing tags for HTML - in tools_mobile_version.asp ---
DisplayCloseBodyAndHTMLTags




' ---------------------------------------------------------------------------------------------
' --- Bottom of MAIN PROGRAM ---
' ---------------------------------------------------------------------------------------------









' --------------------------------
  SUB BuildStatsMemberIDDropdown 
' --------------------------------
  
  

sSQL = "SELECT PersonID, FirstName, LastName, BirthDate, Sex"
sSQL = sSQL + " FROM "&MemberShortTableName
sSQL = sSQL + " WHERE PersonID = RIGHT("&sMemberID&",8)"

IF sWatchMemberIDs<>"" THEN 
		WatchArray=split(sWatchMemberIDs,",")

		FOR Each item IN WatchArray
				sSQL = sSQL + " OR PersonID = RIGHT("&Item&",8)"
		NEXT
END IF

'response.write("</div><div style=color:red>sMemberID = "&sMemberID)
'response.write("<br>SelectedMemberID = "&SelectedMemberID)
'response.write("<br>StatsMemberID = "&StatsMemberID)
'response.write("</div>")




SET rsMemb=Server.CreateObject("ADODB.recordset")
rsMemb.open sSQL, SConnectionToTRATable

%><SELECT id="StatsMemberID" name="StatsMemberID" style="width:15em; font-size:14pt" onchange="submit()"><%

DO WHILE NOT rsMemb.eof

		IF PersonIDwChkDgt(TRIM(rsMemb("PersonID"))) = SelectedMemberID THEN
				response.write("<option value =""" & PersonIDwChkDgt(rsMemb("PersonID")) &""" selected>")
				response.write(rsMemb("FirstName")&" "&rsMemb("LastName"))
				IF sMemberID = PersonIDwChkDgt(rsMemb("PersonID")) THEN response.write(" **")	
				response.write("</option><br>")
		ELSE
				response.write("<option value =""" & PersonIDwChkDgt(rsMemb("PersonID")) &""">")
				response.write(rsMemb("FirstName")&" "&rsMemb("LastName"))
				IF sMemberID = PersonIDwChkDgt(rsMemb("PersonID")) THEN response.write(" **")
				response.write("</option><br>")
		END IF 

		rsMemb.moveNEXT

LOOP

rsMemb.close 

%></SELECT><%



END SUB



' -----------------------------
  SUB GetAllUserAdoptionStats
' -----------------------------


%>
<div id="myteamlisting" style="padding:0px; border:0px solid white;">
	<form method="post">
		<input type="hidden" name="sWatchMemberIDs" value="<%=sWatchMemberIDs%>">
		<input type="hidden" name="sMemberID" value="<%=sMemberID%>">
	<div style="width:100%; margin-top:10px; padding-left:0px; text-align:left;">		
			<span class="span95" style="margin-left:0px; padding-left:0px; text-align:center; font-size:14px; color:yellow; border:0px solid white;">Mobile Statistics</span> 
	</div>	

	<div style="width:100%; margin-top:0px; padding-left:0px; text-align:center; display:inline-block" >
		<%
			
			' BuildStatsMemberIDDropdown 
		
		%>
		
	</div>			
	</form>
	<div class="scroll" style="margin-top:5px; padding:0px; margin-left:0px; height:435px; border:0px solid white;">
		<%   

		' --- Displays SUMMARY Information ---
		' ------------------------------------
		RunUserAdoptionQuery

		IF NOT rs.eof THEN 
				LoopThruMySummary
		ELSE
				DisplayUserSummaryTab
				DisplayNoListingFound "User Summary"
				DisplaySummaryBottom	
		END IF






' --- TEMP ---
RunThis1="Y"
IF RunThis1="Y" THEN

		' -------------------------------------			
		' --- Displays REGION Information ---
		' -------------------------------------		
		RunRegionSummaryQuery

		IF NOT rs.eof THEN 
				rs.movefirst
				LoopThruRegionSummary
		ELSE
				DisplayRegionTab
				DisplayNoListingFound "Region"
				DisplayRegionBottom
		END IF

END IF 		' --SkipAllThis logic



' --- TEMP ---
RunThis2="Y"
IF RunThis2="Y" THEN

		' -------------------------------------			
		' --- Displays DIVISION Information ---
		' -------------------------------------		
		RunDivisionSummaryQuery

		IF NOT rs.eof THEN 
				rs.movefirst
				LoopThruDivisionSummary
		ELSE
				DisplayDivisionTab
				DisplayNoListingFound "Division"
				DisplayRegionBottom
		END IF

END IF 		' --SkipAllThis logic





		' --- Displays SCORES Information ---
		RunUserDetailQuery

		IF NOT rs.eof THEN 
				rs.movefirst
				LoopThruRegisteredUsers
		ELSE
				DisplayRegUsersTab
				DisplayNoListingFound "Users"
				DisplayRegisteredUserBottom
		END IF

  
		
		%>
	</div> <! -- Bottom of scroll box -- ->
</div> <! -- Bottom of div for hiding and displaying -- ->
<%



END SUB




' -----------------------------
  SUB DisplayNoUserSetScreen
' -----------------------------

'response.write("</div><div style=""color:red;"">HERE lin 516</div>")
%>
<div class="tabrankings" style="height:50px; margin-top:50px; padding-top:4px; background-color:<%=DefineLevelcolor%>; text-color:<%=Textcolor1%>; font-size:16px;">
	<span class="span90" style="color:yellow; text-align:center;"><b>This function may not be accessed until an Authorized User has been set up on this device.</b></span>
</div>
<div class="span95" style="margin-top:50px; text-align:center;">
	<form action="/rankings/<%= MenuFileName %>" method="post">
		<input type="submit" value="Return to Main Menu" title="Go to Main Menu" style="font-size:12pt; size:12em;">
	</form>
</div>
<%

END SUB



' -----------------------------------------
  SUB DisplayNoListingFound (whichdisplay) 
' ------------------------------------------  

'response.write("</div><div style=""color:red;"">HERE lin 516</div>")
%>
<div class="rankingsbody" style="height:25px; margin-top:0px; padding-top:10px; background-color:white;">
	<span class="span95" style="color:black; text-align:center; font-size:14px;"><b>No <%= whichdisplay %> Found for these Settings</b></span>
</div>
<%

END SUB



' --------------------------
  SUB LoopThruMySummary
' --------------------------
 

SummaryCount = 1
MetricCount = 1
ThisMetric = rs("Metric")

DO WHILE NOT rs.eof


		GetCurrentMemberSummaryLine
		
		IF SummaryCount=1 THEN DisplayUserSummaryTab
		IF ThisMetric <> rs("Metric") THEN 
				response.write("<hr style=""width:96%; padding:0px 0px 0px 0px; margin:0px 0px 0px 2px; height:1px; background-color:#FFFFFF;"">")
				ThisMetric = rs("Metric")
				MetricCount = 1
		 END IF		
		DisplayUserSummaryLine
		
		rs.movenext
		SummaryCount = SummaryCount + 1
		MetricCount = MetricCount + 1		
LOOP

DisplaySummaryBottom

END SUB



' --------------------------
  SUB LoopThruRegionSummary
' --------------------------
 
MetricCount = 1
EventCount = 1
ThisRegion = rs("Metric")

DO WHILE NOT rs.eof

		GetRegionSummaryLine

		IF MetricCount=1 THEN DisplayRegionTab
		IF ThisEvent <> rs("Metric") THEN 
				response.write("<hr style=""width:96%; padding:0px 0px 0px 0px; margin:0px 0px 0px 2px; height:1px; background-color:#FFFFFF;"">")
				ThisRegion = rs("Metric")
				MetricCount = 1
		END IF		
		DisplayRegionSummaryLine
		
		rs.movenext
		MetricCount = MetricCount + 1
		EventCount = EventCount + 1
LOOP

DisplaySummaryBottom

END SUB



' --------------------------
  SUB LoopThruDivisionSummary
' --------------------------
 
MetricCount = 1
'EventCount = 1
ThisDivision = rs("Metric")

DO WHILE NOT rs.eof

		GetDivisionSummaryLine

		IF MetricCount=1 THEN DisplayDivisionTab
		IF ThisEvent <> rs("Metric") THEN 
				response.write("<hr style=""width:96%; padding:0px 0px 0px 0px; margin:0px 0px 0px 2px; height:1px; background-color:#FFFFFF;"">")
				ThisDivision = rs("Metric")
				MetricCount = 1
		END IF		
		DisplayDivisionSummaryLine
		
		rs.movenext
		MetricCount = MetricCount + 1
		'EventCount = EventCount + 1
LOOP

DisplaySummaryBottom

END SUB










' --------------------------
  SUB LoopThruRegisteredUsers
' --------------------------
 
Dim UserCount
UserCount = 1
TourCount = 1
' ThisMember = rs("Tournament")

DO WHILE NOT rs.eof

		GetCurrentRegisteredUserLine

		IF UserCount=1 THEN DisplayRegUsersTab
		' IF ThisTour <> sRawTournament THEN 
		'		response.write("<hr style=""padding:0px 0px 0px 0px; margin:0px 2px 0px 2px; height:2px; background-color:#FFFFFF;"">")
		'		ThisTour = sRawTournament
		'		TourCount = 1
		'END IF		
		DisplayThisUserLine
		
		rs.movenext
		UserCount = UserCount + 1
		'TourCount = TourCount + 1
LOOP

DisplayRegisteredUserBottom

END SUB









' --------------------------------
  SUB GetRegionSummaryLine
' --------------------------------  
  
'sFullName = rs("FirstName")&" "&rs("LastName")

sSummaryMetric = rs("Metric")
'sSummaryEvent = rs("Event")

sLast60dSummaryValue = rs("Last60dSummaryValue")
sP60to120SummaryValue = rs("P60to120SummaryValue")
sTotalSummaryValue = rs("TotalSummaryValue")


END SUB  




' --------------------------------
  SUB GetDivisionSummaryLine
' --------------------------------  
  
'sFullName = rs("FirstName")&" "&rs("LastName")

sSummaryMetric = rs("Metric")
'sSummaryEvent = rs("Event")

sLast60dSummaryValue = rs("Last60dSummaryValue")
sP60to120SummaryValue = rs("P60to120SummaryValue")
sTotalSummaryValue = rs("TotalSummaryValue")


END SUB  


' --------------------------------
  SUB GetCurrentMemberSummaryLine
' --------------------------------  
  
sSummaryMetric = rs("Metric")
'sSummaryEvent = rs("Event")

sLast60dSummaryValue = rs("Last60dSummaryValue")
sP60to120SummaryValue = rs("P60to120SummaryValue")
sTotalSummaryValue = rs("TotalSummaryValue")

END SUB  


' --------------------------------
  SUB GetCurrentRegisteredUserLine
' --------------------------------  

UserMemberID = rs("MemberID")
sName = rs("Name")
sNum_Forwards = rs("Num_Forwards")
sCreated_Date = FormatDateTime(rs("Created_Date"),2)
IF TRIM(rs("Modified_Date"))<>"" THEN 
		sModified_Date = FormatDateTime(rs("Modified_Date"),2)
ELSE
		sModified_Date = ""
END IF
sMembBirth = rs("BirthDate")
sMemberFed = rs("FederationCode")
sState = rs("State")
sMembAge = DateDiff("yyyy",sMembBirth,NOW())


END SUB  




' ----------------------------
  SUB DisplayUserSummaryTab
' ----------------------------

TabColor = scolor09
%> 
  <div class="tabrankings" style="height:32px; margin:0px 0px 0px 0px; padding:0px 0px 0px 5px; background-color:<% =TabColor %>;" >
		<span class="span95" style="font-size:12pt; font-weight:bold;">User Summary</span>
		<br>
		<span class="span25" style="text-align:left; font-size:9pt; font-weight:normal; border:0px solid red;">Metric</span>
		<span class="span20" style="text-align:right; font-size:9pt; font-weight:normal; padding:0px; margin:0px; border:0px solid yellow;">Last 60</span>
		<span class="span20" style="text-align:right; font-size:9pt; font-weight:normal; padding:0px; margin:0px; border:0px solid red;">61-90</span>
		<span class="span20" style="width:23%; text-align:right; font-size:9pt; font-weight:normal; padding:0px; margin:0px; border:0px solid red;">Total</span>		
	</div>
<%	

END SUB





' ----------------------
  SUB DisplayRegionTab
' ----------------------

TabColor = tcolor04
' TabColor = scolor09
%> 
  <div class="tabrankings" style="height:32px; margin:0px 0px 0px 0px; padding:0px 0px 0px 5px; background-color:<% =TabColor %>;"  >
		<span class="span95" style="font-size:12pt; font-weight:bold;">Region Adoption</span>
		<br>
		<span class="span25" style="text-align:left; font-size:9pt; font-weight:normal; border:0px solid red;">Region</span>
		<span class="span20" style="text-align:right; font-size:9pt; font-weight:normal; border:0px solid yellow;">Last 60</span>
		<span class="span20" style="text-align:right; font-size:9pt; font-weight:normal;">61-90</span>
		<span class="span20" style="width:23%; text-align:right; font-size:9pt; font-weight:normal;">Total</span>		
	</div>
<%	

END SUB




' ----------------------
  SUB DisplayDivisionTab
' ----------------------

TabColor = tcolor05
' TabColor = scolor09
%> 
  <div class="tabrankings" style="height:32px; margin:0px 0px 0px 0px; padding:0px 0px 0px 5px; background-color:<% =TabColor %>;"  >
		<span class="span95" style="font-size:12pt; font-weight:bold;">Division Summary</span>
		<br>
		<span class="span25" style="text-align:left; font-size:9pt; font-weight:normal; border:0px solid red;">Division</span>
		<span class="span20" style="text-align:right; font-size:9pt; font-weight:normal; border:0px solid yellow;">Last 60</span>
		<span class="span20" style="text-align:right; font-size:9pt; font-weight:normal; border:0px solid red;">61-90</span>
		<span class="span20" style="width:23%; text-align:right; font-size:9pt; font-weight:normal; border:0px solid yellow;">Total</span>		
	</div>
<%	

END SUB





' --------------------------
  SUB DisplayUserSummaryLine
' --------------------------

%> 
  <div class="rankingsbody" style="border:0px solid white; margin:0px 0px 0px 2px; padding:0px 0px 0px 5px;">
		<%
		IF MetricCount=1 THEN 
				%><span class="span25" style="text-align:left; font-size:9pt; font-weight:normal; margin:0px; border:0px solid red;"><%= sSummaryMetric %></span><%
		ELSE
				%><span class="span25" style="text-align:left; font-size:9pt; font-weight:normal; margin:0px; border:0px solid blue;">&nbsp;</span><%
		END IF
		%>
		
		<span class="span20" style="text-align:right; font-size:9pt; font-weight:normal; padding:0px; margin:0px; border:0px solid green;"><%= sLast60dSummaryValue %></span>
		<span class="span20" style="text-align:right; font-size:9pt; font-weight:normal; padding:0px; margin:0px; border:0px solid red;"><%= sP60to120SummaryValue %></span>
		<span class="span20" style="width:23%; text-align:right; font-size:9pt; font-weight:normal; padding:0px; margin:0px; border:0px solid yellow;"><%= sTotalSummaryValue %></span>
	</div>
<%	

END SUB




' ---------------------------
  SUB DisplayRegionSummaryLine
' ---------------------------

%> 
  <div class="rankingsbody" style="border:0px solid white; margin:0px 0px 0px 2px; padding:0px 0px 0px 5px;">
		<%
		IF MetricCount=1 THEN 
				%><span class="span25" style="text-align:left; font-size:9pt; font-weight:normal; margin:0px; border:0px solid red;"><%= sSummaryMetric %></span><%
		ELSE
				%><span class="span25" style="text-align:left; font-size:9pt; font-weight:normal; margin:0px;">&nbsp;</span><%
		END IF
		%>
		<span class="span20" style="text-align:right; font-size:9pt; font-weight:normal; padding:0px; margin:0px; border:0px solid green;"><%= sLast60dSummaryValue %></span>
		<span class="span20" style="text-align:right; font-size:9pt; font-weight:normal; padding:0px; margin:0px; border:0px solid red;"><%= sP60to120SummaryValue %></span>
		<span class="span20" style="width:23%; text-align:right; font-size:9pt; font-weight:normal; padding:0px; margin:0px; border:0px solid yellow;"><%= sTotalSummaryValue %></span>
	</div>
<%	


END SUB  



' -------------------------------
  SUB DisplayDivisionSummaryLine
' -------------------------------

%> 
  <div class="rankingsbody" style="border:0px solid white; margin:0px 0px 0px 2px; padding:0px 0px 0px 5px;">
		<%
		IF MetricCount=1 THEN 
				%><span class="span25" style="text-align:left; font-size:9pt; font-weight:normal; margin:0px; border:0px solid red;"><%= sSummaryMetric %></span><%
		ELSE
				%><span class="span25" style="text-align:left; font-size:9pt; font-weight:normal; margin:0px;">&nbsp;</span><%
		END IF
		%>
		
		<span class="span20" style="text-align:right; font-size:9pt; font-weight:normal; padding:0px; margin:0px; border:0px solid yellow;"><%= sLast60dSummaryValue %></span>
		<span class="span20" style="text-align:right; font-size:9pt; font-weight:normal; padding:0px; margin:0px; border:0px solid yellow;"><%= sP60to120SummaryValue %></span>
		<span class="span20" style="width:23%; text-align:right; font-size:9pt; font-weight:normal; padding:0px; margin:0px; border:0px solid yellow;"><%= sTotalSummaryValue %></span>
	</div>
<%	


END SUB  



' -------------------------
  SUB DisplayRegUsersTab
' -------------------------

TabColor = scolor06
%> 
  <div class="tabrankings" style="height:32px; margin:0px 0px 0px 2px; padding:0px 0px 0px 5px; background-color:<% =TabColor %>;"  >
		<span class="span80" style="font-size:12pt; font-weight:bold;">Registered Members</span>
		<br>
		<span class="span50" style="text-align:left; font-size:9pt; font-weight:normal;">Name</span>
		<span class="span15" style="text-align:center; font-size:9pt; font-weight:normal;">Forwards</span>	
	</div>
<%	




END SUB



' ---------------------------
  SUB DisplayThisUserLine
' ---------------------------

sbgcolor = scolor06
' scolor06

  %>
  <div class="tabrankings" style="height:17px; background-color:<%=sbgcolor%>; font-size:12pt; margin:0px 0px 0px 0px; padding:0px 0px 0px 5px;" >
  	<span class="span60" style="width:57%;"><b><%=LEFT(sName,20)%></b></span>
	  <span class="span20" style="width:20%; text-align:right;"><b><%= sNum_Forwards %></b></span>
 		<span style="width=5px; color:red; text-align:left;"><%=displaysuffix%></span>
	</div>
  
  <div class="rankingsbody" style="background-color:#FFFFFF; font-size:10pt; font-weight:normal; margin:0px 0px 0px 0px; padding:0px 0px 0px 5px; text-align:left;">
				<span class="span45">CR: <%= sCreated_Date %></span>
				<span class="span45" style="text-align:left;">MOD: <%= sModified_Date %></span>

	</div>
  
  <div class="tourbottom"  style="background-color:#FFFFFF; margin:0px 0px 0px 2px; padding:0px 0px 0px 5px;">
   	<span class="span25" style="font-size:13px;">Fed: <b><% =sMemberFed %></b></span>
		<span class="span25" style="text-align:left;">&nbsp;State: <b><%= sState %></b></span>
	</div>
	<%


END SUB





' -------------------------------
  SUB DisplaySummaryBottom
' -------------------------------  

%>
<div class="rankingsbottom" style="background-color:#FFFFFF; height:10px; margin:0px 0px 0px 2px; padding:0px 0px 0px 5px;">
		<span class="span95">&nbsp;</span>
</div>
<%

END SUB


' -------------------------------
  SUB DisplayRegisteredUserBottom
' -------------------------------  

%>
<div class="rankingsbottom" style="background-color:#FFFFFF; height:10px; margin:0px 0px 0px 4px; padding:0px 0px 0px 2px;">
		<span class="span95">&nbsp;</span>
</div>
<%

END SUB






' -------------------------
  SUB RunMyRankingsQuery
' -------------------------

'sMemberID = "800074054"
sSkiYearID="1"

sSQL = sSQL + " SELECT FirstName, LastName, RT1.Event, RT1.Div, Nat_Qual"
sSQL = sSQL + ", COALESCE(RT1.Regl_Plc,' ') AS Regl_Plc, COALESCE(RT1.Natl_Plc,' ') AS Natl_Plc"
sSQL = sSQL + " , RankScoreThis, RankScoreLast"
sSQL = sSQL + " FROM "

sSQL = sSQL + " ( SELECT MemberID, Event, Div, Regl_Plc, Natl_Plc, Nat_Qual, RankScore AS RankScoreThis"
sSQL = sSQL + "	    FROM "&RankTableName
sSQL = sSQL + "       WHERE MemberID='"&SelectedMemberID&"' AND SkiYearID='"&sSkiYearID&"'"
sSQL = sSQL + "         AND lower(left(div,1)) NOT IN ('e','s') ) RT1" 

sSQL = sSQL + "	LEFT JOIN"
sSQL = sSQL + "    ( SELECT MemberID, Event, Div, Regl_Plc, Natl_Plc, RankScore AS RankScoreLast"
sSQL = sSQL + "	      FROM "&RankTableName
sSQL = sSQL + "         WHERE MemberID='"&SelectedMemberID&"'"
'sSQL = sSQL + "   AND SkiYearID = (SELECT PrevYearID FROM usawsrank.SkiYear WHERE DefaultYear=1)"
sSQL = sSQL + "   AND SkiYearID = (SELECT SkiYearID FROM usawsrank.SkiYear WHERE SkiYear=(SELECT SkiYear-2 FROM usawsrank.SkiYear WHERE DefaultYear=1) )"
sSQL = sSQL + "           AND lower(left(div,1)) NOT IN ('e','s') ) RT2" 
sSQL = sSQL + " ON RT2.MemberID=RT1.MemberID AND RT2.Event=RT1.Event AND RT2.Div=RT1.Div"

sSQL = sSQL + "	JOIN "&MemberShortTableName&" AS MT ON CAST(RIGHT(RT1.MemberID,8) AS INT)=MT.PersonID"

sSQL = sSQL + " ORDER BY CASE when RT1.Event='S' then 1 when RT1.event='T' then 2 when RT1.Event='J' then 3 else 4 end, RT1.Div"

'response.write("</div><div style=color:red>"&sSQL&"</div>")
'response.end

Set rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable


END SUB



' --------------------------
  SUB RunUserAdoptionQuery
' --------------------------

sSQL = " SELECT *"
sSQL = sSQL + " 	FROM"
sSQL = sSQL + " 	("

sSQL = sSQL + " SELECT 1 AS MetricOrder, 'Reg User' AS Metric"

sSQL = sSQL + " , SUM(CASE WHEN Created_Date>=DateAdd(d,-60, GETDATE()) THEN 1 ELSE 0 END) AS Last60dSummaryValue"
sSQL = sSQL + " , SUM(CASE WHEN Created_Date<DateAdd(d,-60, GETDATE()) AND Created_Date>=DateAdd(d,-120, GETDATE()) THEN 1 ELSE 0 END) AS P60to120SummaryValue"
sSQL = sSQL + " , COUNT(*) AS TotalSummaryValue"
sSQL = sSQL + " 		FROM "&MobileAppUserTable

sSQL = sSQL + " UNION "

sSQL = sSQL + " SELECT 2 AS MetricOrder, 'Shares' AS Metric"
sSQL = sSQL + " , SUM(CASE WHEN Created_Date>=DateAdd(d,-60, GETDATE()) THEN Num_Forwards ELSE 0 END) AS Last60dSummaryValue"
sSQL = sSQL + " , SUM(CASE WHEN Created_Date<DateAdd(d,-60, GETDATE()) AND Created_Date>=DateAdd(d,-120, GETDATE()) THEN Num_Forwards ELSE 0 END) AS P60to120SummaryValue"
sSQL = sSQL + " , SUM(Num_Forwards) AS TotalSummaryValue"

sSQL = sSQL + " FROM "&MobileAppUserTable
sSQL = sSQL + ") A"
sSQL = sSQL + " 	ORDER BY MetricOrder"

'response.write("</div><div style='background-color:white; color:black;'>"&sSQL)

'response.end

Set rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable

END SUB



' --------------------------
  SUB RunRegionSummaryQuery
' --------------------------

sSQL = " SELECT *"
sSQL = sSQL + " 	FROM"
sSQL = sSQL + " 	("

sSQL = sSQL + " SELECT CASE WHEN Region=1 THEN 'S Central'"
sSQL = sSQL + " 		WHEN Region=2 THEN 'Midwest'"
sSQL = sSQL + " 		WHEN Region=3 THEN 'South'"
sSQL = sSQL + " 		WHEN Region=4 THEN 'West'"
sSQL = sSQL + " 		WHEN Region=5 THEN 'East'"
sSQL = sSQL + " 		ELSE 'International' END AS Metric"

sSQL = sSQL + " , SUM(CASE WHEN Created_Date>=DateAdd(d,-60, GETDATE()) THEN 1 ELSE 0 END) AS Last60dSummaryValue"
sSQL = sSQL + " , SUM(CASE WHEN Created_Date<DateAdd(d,-60, GETDATE()) AND Created_Date>=DateAdd(d,-120, GETDATE()) THEN 1 ELSE 0 END) AS P60to120SummaryValue"
sSQL = sSQL + " , COUNT(*) AS TotalSummaryValue"

sSQL = sSQL + " 		FROM "&MobileAppUserTable&" map"
sSQL = sSQL + "			JOIN "&MemberShortTableName& " m ON m.PersonID = RIGHT(map.MemberID,8)"

sSQL = sSQL + " 	GROUP BY CASE WHEN Region=1 THEN 'S Central'"
sSQL = sSQL + " 		WHEN Region=2 THEN 'Midwest'"
sSQL = sSQL + " 		WHEN Region=3 THEN 'South'"
sSQL = sSQL + " 		WHEN Region=4 THEN 'West'"
sSQL = sSQL + " 		WHEN Region=5 THEN 'East'"
sSQL = sSQL + " 		ELSE 'International' END"

'sSQL = sSQL + " UNION "

'sSQL = sSQL + " SELECT 2 AS MetricOrder, 'Shares' AS Metric"
'sSQL = sSQL + " , SUM(CASE WHEN Created_Date>=DateAdd(d,-60, GETDATE()) THEN Num_Forwards ELSE 0 END) AS Last60dSummaryValue"
'sSQL = sSQL + " , SUM(CASE WHEN Created_Date<DateAdd(d,-60, GETDATE()) AND Created_Date>=DateAdd(d,-120, GETDATE()) THEN Num_Forwards ELSE 0 END) AS P60to120SummaryValue"
'sSQL = sSQL + " , SUM(Num_Forwards) AS TotalSummaryValue"

' sSQL = sSQL + " FROM "&MobileAppUserTable

sSQL = sSQL + ") A"

sSQL = sSQL + " 	ORDER BY Metric"

'response.write("</div><div style='background-color:white; color:black;'>"&sSQL)

'response.end

Set rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable

END SUB




' --------------------------
  SUB RunDivisionSummaryQuery
' --------------------------

sSQL = " SELECT *"
sSQL = sSQL + " 	FROM"
sSQL = sSQL + " 	("

sSQL = sSQL + " SELECT Div AS Metric"

sSQL = sSQL + " , SUM(CASE WHEN Created_Date>=DateAdd(d,-60, GETDATE()) THEN 1 ELSE 0 END) AS Last60dSummaryValue"
sSQL = sSQL + " , SUM(CASE WHEN Created_Date<DateAdd(d,-60, GETDATE()) AND Created_Date>=DateAdd(d,-120, GETDATE()) THEN 1 ELSE 0 END) AS P60to120SummaryValue"
sSQL = sSQL + " , COUNT(*) AS TotalSummaryValue"

sSQL = sSQL + " 		FROM "&MobileAppUserTable&" map"
sSQL = sSQL + "			JOIN "&MemberShortTableName& " m ON m.PersonID = RIGHT(map.MemberID,8)"

sSQL = sSQL + "			JOIN usawsrank.division d ON LOW_AGE<=COALESCE(DateDiff(yy,BirthDate,GETDATE()),0)" 
sSQL = sSQL + "				AND UP_AGE>=COALESCE(DateDiff(yy,BirthDate,GETDATE()),0)"
sSQL = sSQL + "					AND ( (d.Sex='M' AND m.Sex='Male') OR (d.Sex='F' AND m.Sex='Female') )"
sSQL = sSQL + "					AND SkiYearID=1"
sSQL = sSQL + "					AND LOWER(LEFT(d.Div,1)) IN ('m','w','b','g')"
sSQL = sSQL + "					AND LOWER(d.Div) NOT IN ('mm','mw')"

sSQL = sSQL + " 	GROUP BY Div"

sSQL = sSQL + ") A"

sSQL = sSQL + " 	ORDER BY Metric"

'response.write("</div><div style='background-color:white; color:black;'>"&sSQL)


'SELECT
'COALESCE(DateDiff("yy",BirthDate,GETDATE()),0) AS Test
', Gender
'FROM usawaterski.dbo.membershort
'JOIN usawsrank.division d ON LOW_AGE>=COALESCE(DateDiff("yy",BirthDate,GETDATE()),0) AND UP_AGE<=COALESCE(DateDiff("yy",BirthDate,GETDATE()),0) AND Gender=Sex


'response.write("</div><div style='color:black; background-color:white;'>"&sSQL)
'response.end

Set rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable

END SUB



' -----------------------
  SUB RunUserDetailQuery
' -----------------------  
  
sSQL = "SELECT MemberID, LEFT(FirstName,1) + ' ' + LastName AS Name"
sSQL = sSQL + " , Email, FirstName, LastName"
sSQL = sSQL + " , Num_Forwards, Created_Date, Modified_Date"
sSQL = sSQL + " , State, BirthDate, FederationCode"
sSQL = sSQL + " , Address1, City, State, Zip"
sSQL = sSQL + " FROM "&MobileAppUserTable&" map"
sSQL = sSQL + "	JOIN "& MemberShortTableName & " m ON m.PersonID = RIGHT(map.MemberID,8)"
'sSQL = sSQL + "	WHERE MemberID = '"&SelectedMemberID&"'"
'sSQL = sSQL + "				AND TDateE>='"&ThisYear12moStart&"'"
sSQL = sSQL + "	ORDER BY Created_Date DESC"

Set rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable


' --- For running in SQL Server Studio --
' SELECT MemberID, LEFT(FirstName,1) + ' ' + LastName AS Name
'  , Email, FirstName, LastName
'  , Num_Forwards, Created_Date, Modified_Date
'  , State, BirthDate, FederationCode
'  , Address1, City, State, Zip
'   FROM usawsrank.mobile_Appusers map
'   JOIN usawaterski.dbo.membershort m ON m.PersonID = RIGHT(map.MemberID,8)
'   -- WHERE LastName='Reece'
'  ORDER BY Num_Forwards DESC, Created_Date DESC


 	
END SUB  

%>