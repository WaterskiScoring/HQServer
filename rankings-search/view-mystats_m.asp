<!--#include file="settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include16.asp"-->
<!--#include virtual="/rankings/tools_Definitions.asp"-->
<!--#include virtual="/rankings/tools_leagues.asp"-->
<!--#include virtual="/rankings/tools_mobile_version.asp"-->
<%



Dim ThisFileName

Dim sFullName, sRankDiv, sRankEvent, sRankTourID
Dim sRankScoreThis, sRankScoreLast, sNat_Qual, sRegl_Plc, sNatl_Plc
Dim sSummaryMetric, sSummaryEvent, sThisSummaryValue, sLastSummaryValue
Dim MetricCount, EventCount, TourCount
Dim ThisYear12moStart, ThisYear12moEnd, LastYear12moStart, LastYear12moEnd
Dim sRawTournament, sRawTDateE, sRawEvent, sRawDiv, sRawPlace, sRawRound, sRawScore, sMaxScore
Dim sFirstName, sLastName, sAddress1, sAddress2, sCity, sState, sZip, sTourID
Dim StatsMemberID, sWatchMemberIDs, SelectedMemberID

ThisFileName = "view-mystats_m.asp"



' --- Date Range calcs for last 12 months This and Last --
ThisYear12moStart = DateAdd("d",-364,Date)
ThisYear12moEnd = Date
LastYear12moStart = DateAdd("d",-729,Date)
LastYear12moEnd = DateAdd("d",-365,Date)


sMemberID = TRIM(Request("sMemberID"))
StatsMemberID = TRIM(Request("StatsMemberID"))
sWatchMemberIDs = Request("sWatchMemberIDs")

'response.write("<br>LINE 33 stats - sMemberID = "&sMemberID)
'response.write("<br> IsNull(sMemberID) = "&IsNull(sMemberID))

'response.end

' SelectedMemberID=sMemberID
IF StatsMemberID<>"" THEN SelectedMemberID=StatsMemberID ELSE SelectedMemberID=sMemberID


'response.write("<br>TEST")
'response.write("<br>sMemberID = "&sMemberID)
' response.write("<br>SelectedMemberID = "&SelectedMemberID)
'response.end 

OpenState="mystats"
DisplayHeadOpenBodyAndBannerTags OpenState


' --Custom functionality for this module --
BuildCustomJavascript




IF sMemberID<>"" THEN 
		'response.write("<br>LINE 52 stats - INSIDE sMemberID<> dbl quote = "&sMemberID)
	' response.end
		' --- Assembles and displays all the stats ---
		GetAllMyStats
ELSE
		' --- Display No User Set screen
		DisplayNoUserSetScreen
END IF	




' --- Writes the Closing tags for HTML - in tools_mobile_version.asp ---
DisplayCloseBodyAndHTMLTags




' ---------------------------------------------------------------------------------------------
' --- Bottom of MAIN PROGRAM ---
' ---------------------------------------------------------------------------------------------




' ---------------------------
  SUB BuildCustomJavascript
' ---------------------------  

%>
<script>
	function loadXMLDoc(stid, smid, sevt) {
  	
  	var stid = stid
  	var smid = smid
  	var sevt = sevt
  	
  	var PostURL = 'Personal_Best_Recording.asp?stid=' + stid + '&smid=' + smid + '&sevt=' + sevt;
  	// alert('URL = ' + PostURL);
  	var xhttp = new XMLHttpRequest();
  	xhttp.onreadystatechange = function() {
    if (this.readyState == 4 && this.status == 200) {
	
				// alert('this.responseText = ' + this.responseText);
				var responsetxxt =  this.responseText;
				// var xmlDoc = this.responseXML;
				// Set xmlList = xmlDoc.getElementsByTagName("result");

				parser = new DOMParser();
				xmlDoc = parser.parseFromString(responsetxxt,"text/xml");

				var sMemberID = xmlDoc.getElementsByTagName("memberid")[0].childNodes[0].nodeValue;
				var sTourID = xmlDoc.getElementsByTagName("tourid")[0].childNodes[0].nodeValue;
				var sEventName = xmlDoc.getElementsByTagName("eventname")[0].childNodes[0].nodeValue;
				var sScore = xmlDoc.getElementsByTagName("score")[0].childNodes[0].nodeValue;
				var sUnits = xmlDoc.getElementsByTagName("units")[0].childNodes[0].nodeValue;
				var scoreexists = xmlDoc.getElementsByTagName("scoreexists")[0].childNodes[0].nodeValue;
				var emailexists = xmlDoc.getElementsByTagName("emailexists")[0].childNodes[0].nodeValue;

				

				if (scoreexists == 'N') {
						if (emailexists == 'Y') {
							 messageBody = 'Your request has been recorded for the score of ' + sScore + ' ' + sUnits +' in ' + sEventName + ' from sanction ' + sTourID + '. Please look for a confirmation email.'
						}
						else {
								messageBody = 'Your request has been recorded for the score of ' + sScore + ' ' + sUnits +' in ' + sEventName + ' from sanction ' + sTourID + '.'
						}
				}
				
				else {
						messageBody = 'Thank you.  A previous request was received for the score of ' + sScore + ' ' + sUnits +' in ' + sEventName + ' from sanction ' + sTourID + '.'	
				}	
				
				// -- Displays notice to user --
				alert(messageBody);
				
				// alert('MemberID = ' + sMemberID);
				// alert('TourID = ' + sTourID);
				// alert('Event = ' + sEvent);
				// alert('Score = ' + sScore);
				// alert('scoreexists = ' + scoreexists);

    	}
  	};
  	xhttp.open("GET", "" + PostURL + "", true);
  	xhttp.send();
	}
</script>
<%
	

END SUB  





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



' -----------------------
  SUB GetAllMyStats
' -----------------------  


%>
<div id="myteamlisting" style="padding:0px; border:0px solid white;">
	<form method="post">
		<input type="hidden" name="sWatchMemberIDs" value="<%=sWatchMemberIDs%>">
		<input type="hidden" name="sMemberID" value="<%=sMemberID%>">
	<div style="width:100%; margin-top:10px; padding-left:0px; text-align:left;">		
			<span class="span95" style="margin-left:0px; padding-left:0px; text-align:center; font-size:14px; color:yellow; border:0px solid white;">Competition Statistics For</span> 
	</div>	

	<div style="width:100%; margin-top:0px; padding-left:0px; text-align:center; display:inline-block" >
		<%
			
			BuildStatsMemberIDDropdown 
		
		%>
		
	</div>			
	</form>
	<div class="scroll" style="margin-top:5px; padding:0px; margin-left:0px; border:0px solid white;">
		<%   

		' --- Displays SUMMARY Information ---
		RunSummaryQuery

		IF NOT rs.eof THEN 
				LoopThruMySummary
		ELSE
				DisplaySummaryTab
				DisplayNoListingFound "Summary"
				DisplayRankingBottomLine	
		END IF
		
		' --- Displays RANKINGS Information ---
		RunMyRankingsQuery

		IF NOT rs.eof THEN 
				rs.movefirst
				LoopThruMyRankings
		ELSE
				DisplayRankingTab
				DisplayNoListingFound "Rankings"
				DisplayRankingBottomLine
		END IF

		' --- Displays SCORES Information ---
		RunRawScoresQuery

		IF NOT rs.eof THEN 
				rs.movefirst
				LoopThruMyRawScores
		ELSE
				DisplayRawScoresTab
				DisplayNoListingFound "Scores"
				DisplayRankingBottomLine
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
<div class="rankingsbody" style="width:97%; border:0px solid white; padding:0px 3px 0px 2px; margin:0px 0px 0px 2px;>
	<span class="span90" style="color:black; text-align:center; font-size:14pt; font-weight:bold; margin:0px 0px 0px 30px;">No <%= whichdisplay %> Found for these Settings</span>
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
		
		IF SummaryCount=1 THEN DisplaySummaryTab
		IF ThisMetric <> rs("Metric") THEN 
				response.write("<hr style=""width:97%; padding:0px 0px 0px 3px; margin:0px 2px 0px 2px; height:2px; background-color:#FFFFFF;"">")
				ThisMetric = rs("Metric")
				MetricCount = 1
		 END IF		
		DisplayMySummaryLine
		
		rs.movenext
		SummaryCount = SummaryCount + 1
		MetricCount = MetricCount + 1		
LOOP

DisplayRankingBottomLine

END SUB



' --------------------------
  SUB LoopThruMyRankings
' --------------------------
 
RankCount = 1
EventCount = 1
ThisEvent = rs("Event")

DO WHILE NOT rs.eof

		GetCurrentMemberRankingLine

		IF RankCount=1 THEN DisplayRankingTab
		IF ThisEvent <> rs("Event") THEN 
				response.write("<hr style=""width:97%; padding:0px 0px 0px 3px; margin:0px 2px 0px 2px; height:2px; background-color:#FFFFFF;"">")
				ThisEvent = rs("Event")
				EventCount = 1
		END IF		
		DisplayMyRankingLine
		
		rs.movenext
		RankCount = RankCount + 1
		EventCount = EventCount + 1
LOOP

DisplayRankingBottomLine

END SUB



' --------------------------
  SUB LoopThruMyRawScores
' --------------------------
 
Dim ScoreCount
ScoreCount = 1
TourCount = 1
ThisTour = rs("Tournament")

DO WHILE NOT rs.eof

		GetCurrentMemberRawScoresLine

		IF ScoreCount=1 THEN DisplayRawScoresTab
		IF ThisTour <> sRawTournament THEN 
				response.write("<hr style=""width:97%; padding:0px 0px 0px 3px; margin:0px 2px 0px 2px; height:2px; background-color:#FFFFFF;"">")
				ThisTour = sRawTournament
				TourCount = 1
		END IF		
		DisplayMyRawScoresLine
		
		rs.movenext
		ScoreCount = ScoreCount + 1
		TourCount = TourCount + 1
LOOP

DisplayRankingBottomLine

END SUB









' --------------------------------
  SUB GetCurrentMemberRankingLine
' --------------------------------  
  
sFullName = rs("FirstName")&" "&rs("LastName")

sRankDiv = rs("Div")
sRankEvent = rs("Event")

sRankScoreThis = rs("RankScoreThis")
sRankScoreLast = rs("RankScoreLast")
sNat_Qual = rs("Nat_Qual")
sRegl_Plc = rs("Regl_Plc")
sNatl_Plc = rs("Natl_Plc")

END SUB  



' --------------------------------
  SUB GetCurrentMemberSummaryLine
' --------------------------------  
  
sSummaryMetric = rs("Metric")
sSummaryEvent = rs("Event")

sThisSummaryValue = rs("ThisSummaryValue")
sLastSummaryValue = rs("LastSummaryValue")

END SUB  


' ----------------------------------
  SUB GetCurrentMemberRawScoresLine
' -----------------------------------  

sSQL = "SELECT MemberID, TName AS Tournament, Event, Div, Round, Score, TDateE"  
sRawTournament = rs("Tournament")
sTourID = rs("TourID")
sRawTDateE = rs("TDateE")
sRawEvent = rs("Event")
sRawDiv = rs("Div")
sRawPlace = rs("Place")
sRawRound = rs("Round")
sRawScore = rs("Score")
sMaxScore = rs("MaxScore")
sFirstName = rs("FirstName")
sLastName = rs("LastName")
sMemberID = rs("MemberID")
sAddress1 = rs("Address1")
sAddress2 = rs("Address2")
sCity = rs("City")
sState = rs("State")
sZip = rs("Zip")


END SUB  




' ----------------------
  SUB DisplaySummaryTab
' ----------------------

TabColor = tcolor03
%> 
  <div class="tabrankings" style="width:97%; height:32px; background-color:<% =TabColor %>; padding:0px 1px 0px 4px; margin:0px 0px 0px 0px;" >
		<span class="span95" style="font-size:12pt; margin-left:5px; font-weight:bold;">Summary</span>
		<br>
		<span class="span35" style="text-align:left; margin-left:5px; font-size:9pt; font-weight:normal;">Metric</span>
		<span class="span10" style="text-align:center; font-size:9pt; font-weight:normal;">Event</span>
		<span class="span25" style="text-align:right; font-size:9pt; font-weight:normal;">This Year</span>
		<span class="span25" style="width:22%; text-align:right; font-size:9pt; font-weight:normal;">Last Year</span>		
	</div>
<%	

END SUB


' --------------------------
  SUB DisplayMySummaryLine
' --------------------------

%> 
  <div class="rankingsbody" style="width:97%; border:0px solid white; padding:0px 3px 0px 2px; margin:0px 0px 0px 2px;">
		<%
		IF MetricCount=1 THEN 
				%><span class="span35" style="text-align:left; font-size:9pt; font-weight:normal;"><%= sSummaryMetric %></span><%
		ELSE
				%><span class="span35" style="text-align:left; font-size:9pt; font-weight:normal;">&nbsp;</span><%
		END IF
		%>
		<span class="span10" style="text-align:center; font-size:9pt; font-weight:normal;"><%= sSummaryEvent %></span>
		<span class="span25" style="text-align:right; padding-left:3px; font-size:9pt; font-weight:normal;"><%= sThisSummaryValue %></span>
		<span class="span25" style="width:22%; text-align:right; font-size:9pt; font-weight:normal;"><%= sLastSummaryValue %></span>		
	</div>
<%	

END SUB



' ----------------------
  SUB DisplayRankingTab
' ----------------------

TabColor = tcolor02
%> 
  <div class="tabrankings" style="width:97%; height:32px; background-color:<% =TabColor %>; padding:0px 1px 0px 4px; margin:0px 0px 0px 0px;">
		<span class="span95" style="font-size:12pt; font-weight:bold;">Rankings</span>
		<br>
		<span class="span15" style="text-align:left; font-size:9pt; font-weight:normal;">Event</span>
		<span class="span10" style="text-align:center; font-size:9pt; font-weight:normal;">Div</span>
		<span class="span10" style="width:9%; text-align:right; font-size:9pt; font-weight:normal;">Regl</span>
		<span class="span10" style="width:8%; text-align:right; font-size:9pt; font-weight:normal;">Natl</span>
		<span class="span25" style="text-align:right; font-size:9pt; font-weight:normal;">Rank Score</span>
		<span class="span25" style="width:25%; font-size:9pt; text-align:right; font-weight:normal;">Last Year</span>
	</div>
<%	

END SUB



' ---------------------------
  SUB DisplayMyRankingLine
' ---------------------------

SELECT CASE TRIM(sRankEvent)
		CASE "S" 
			sRankEventText = "Slalom"
		CASE "T" 
			sRankEventText = "Tricks"
		CASE "J" 
			sRankEventText = "Jump"
		CASE "O" 
			sRankEventText = "Overall"
END SELECT			

IF IsNull(sRankScoreThis) THEN sRankScoreThisText="" ELSE sRankScoreThisText = FormatNumber(sRankScoreThis,2)
IF IsNull(sRankScoreLast) THEN sRankScoreLastText="" ELSE sRankScoreLastText = FormatNumber(sRankScoreLast,2)



%>
  <div class="rankingsbody" style="width:97%; border:0px solid black; padding:0px 3px 0px 2px; margin:0px 0px 0px 2px;">
		<%
		IF EventCount=1 THEN 
				%><span class="span15" style="text-align:left; font-size:9pt; font-weight:normal;"><%= sRankEventText %></span><%
		ELSE
				%><span class="span15" style="text-align:left; font-size:9pt; font-weight:normal;">&nbsp;</span><%
		END IF
		%>		
		<span class="span10" style="width:8%; text-align:center; font-size:9pt; font-weight:normal;"><%= sRankDiv %></span>
		<span class="span10" style="width:8%; text-align:right; font-size:9pt; font-weight:normal;"><%= sRegl_Plc %></span>
		<span class="span10" style="text-align:right; font-size:9pt; font-weight:normal;"><%= sNatl_Plc %></span>
		<span class="span25" style="text-align:right; font-size:9pt; font-weight:normal;"><%= sRankScoreThisText %></span>
		<span class="span25" style="text-align:right; font-size:9pt; font-weight:normal;"><%= sRankScoreLastText %></span>
	</div>
<%

END SUB  




' -------------------------
  SUB DisplayRawScoresTab
' -------------------------

TabColor = tcolor01
%> 
  <div class="tabrankings" style="width:97%; height:32px; background-color:<% =TabColor %>; padding:0px 0px 0px 5px; margin:0px 0px 0px 0px;">
		<span class="span80" style="font-size:12pt; font-weight:bold;">Scores - Last 12 Months</span>
		<br>
		<span class="span15" style="text-align:center; font-size:9pt; font-weight:normal;">Event</span>
		<span class="span15" style="text-align:center; font-size:9pt; font-weight:normal;">Div</span>
		<span class="span15" style="text-align:right; font-size:9pt; font-weight:normal;">Place</span>
		<span class="span15" style="text-align:center; font-size:9pt; font-weight:normal;">Round</span>		
		<span class="span30" style="text-align:right; font-size:9pt; font-weight:normal;">Score</span>
	</div>
<%	

END SUB



' ---------------------------
  SUB DisplayMyRawScoresLine
' ---------------------------

PersBest = ""
scolor="#FFFFFF"
IF sRawScore >= sMaxScore THEN 
		PersBest = "*" 
		scolor="#fff099"
		scolor="yellow"
END IF	

SELECT CASE sRawEvent
		CASE "S"
				sEventName = "Slalom"	
		CASE "T"
				sEventName = "Tricks"
		CASE "J"
				sEventName = "Jumping"
END SELECT
								
%> 
  <div class="rankingsbody" style="width:97%; border:0px solid white; padding:0px 2px 0px 3px; margin:0px 0px 0px 2px;">
		<%
		IF TourCount=1 THEN 
				%>
				<span class="span75" style="text-align:left; background-color:#FFFFFF; font-size:9pt; font-weight:normal;"><%= LEFT(sRawTournament,35) %></span>
				<span class="span20" style="text-align:right; background-color:#FFFFFF; font-size:8pt; font-weight:normal;"><%= sRawTDateE %></span>
				<br>
				<%
		END IF
		%>
	</div>
  <div class="rankingsbody" style="width:97%; background-color:<%= scolor %>; border:0px solid white; padding:0px 2px 0px 3px; margin:0px 0px 0px 2px;">
		<span class="span15" style="text-align:center; font-size:9pt; font-weight:normal;"><%= sRawEvent %></span>
		<span class="span15" style="text-align:center; font-size:9pt; font-weight:normal;"><%= sRawDiv %></span>
		<span class="span15" style="text-align:right; font-size:9pt; font-weight:normal;"><%= sRawPlace %></span>
		<span class="span15" style="text-align:center; font-size:9pt; font-weight:normal;"><%= sRawRound %></span>		
		<span class="span30" style="text-align:right; font-size:9pt; font-weight:normal;"><%= sRawScore %></span>
	</div>
		<%
		mailtoaddress = "j_surdej@yahoo.com"
		bccEmail = "cronemarka@gmail.com"
		IF PersBest = "*" THEN 
				%>
  			<div class="rankingsbody" style="width:97%; text-align:center; background-color:<%= scolor %>; border:0px solid white; padding:0px 2px 0px 3px; margin:0px 0px 0px 2px;">
					<a title="Request Personal Best Decal" href="javascript:loadXMLDoc('<%=sTourID%>','<%=sMemberID%>', '<%=sRawEvent%>')">
						<span class="span95" style="text-align:center; color:red; font-size:9pt; font-weight:normal;">** PERSONAL BEST **<br>Click here to request your <b>Personal Best</b> sticker</span>
					</a>
				</div>
				<%
		END IF
		%>
	
<%	

END SUB





' -------------------------------
  SUB DisplayRankingBottomLine
' -------------------------------  
' style="width:97%; border:0px solid white; padding:0px 3px 0px 2px; margin:0px 0px 0px 2px;"
' style="width:97%; border:0px solid white; padding:0px 2px 0px 3px; margin:0px 0px 0px 2px;"
' -- Orig style="width:97%; background-color:#FFFFFF; height:10px; margin:0px 0px 0px 2px; padding:0px 3px 0px 5px;"
%>
<div class="rankingsbottom" style="width:97%; background-color:#FFFFFF; border-right:1px solid #FFFFFF; height:10px; padding:0px 2px 0px 2px; margin:0px 0px 0px 2px;">
		<span class="span100">&nbsp;</span>
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



' --------------------
  SUB RunSummaryQuery
' --------------------




sSQL = " SELECT *"
sSQL = sSQL + " 	FROM"
sSQL = sSQL + " 	("
sSQL = sSQL + " SELECT THS.MemberID, 1 AS MetricOrder, THS.Metric, THS.Event, ThisScore AS ThisSummaryValue, LastScore AS LastSummaryValue"
sSQL = sSQL + " FROM"
sSQL = sSQL + " 	( SELECT MemberID, 'High Score' AS Metric, Event, MAX(Score) AS ThisScore"
sSQL = sSQL + " 		FROM usawsrank.Scores s"
sSQL = sSQL + " 		JOIN sanctions.dbo.TSchedul ts ON ts.TournAppID = LEFT(s.TourID,6)"
sSQL = sSQL + " 			WHERE MemberID = '"&SelectedMemberID&"'"
sSQL = sSQL + " 				AND TDateE>='"&ThisYear12moStart&"'"
sSQL = sSQL + " 		GROUP BY MemberID, Event ) THS"
sSQL = sSQL + " 	LEFT JOIN"
sSQL = sSQL + " 	( SELECT MemberID, 'High Score' AS Metric, Event, MAX(Score) AS LastScore"
sSQL = sSQL + " 		FROM usawsrank.Scores s"
sSQL = sSQL + " 		JOIN sanctions.dbo.TSchedul ts ON ts.TournAppID = LEFT(s.TourID,6)"
sSQL = sSQL + " 			WHERE MemberID = '"&SelectedMemberID&"'"
sSQL = sSQL + " 				AND TDateE<='"&LastYear12moEnd&"' AND TDateE>='"&LastYear12moStart&"'"
sSQL = sSQL + " 		GROUP BY MemberID, Event ) LHS"
sSQL = sSQL + " 	ON LHS.MemberID=THS.MemberID AND LHS.Event=THS.Event"
	
sSQL = sSQL + " 	UNION" 
	 
sSQL = sSQL + " 	 SELECT TSC.MemberID, 2 AS MetricOrder, TSC.Metric, TSC.Event, ThisScoreCount AS ThisSummaryValue, LastScoreCount AS LastSummaryValue"
sSQL = sSQL + " 		FROM"
sSQL = sSQL + " 	( SELECT MemberID, '# of Scores' AS Metric, Event, COUNT(Score) AS ThisScoreCount, Count(DISTINCT TourID) AS ThisTourCount"
sSQL = sSQL + " 		FROM usawsrank.Scores s"
sSQL = sSQL + " 		JOIN sanctions.dbo.TSchedul ts ON ts.TournAppID = LEFT(s.TourID,6)"
sSQL = sSQL + " 			WHERE MemberID = '"&SelectedMemberID&"'"
sSQL = sSQL + " 				AND TDateE>='"&ThisYear12moStart&"'"
sSQL = sSQL + " 		GROUP BY MemberID, Event ) TSC"
sSQL = sSQL + " 	LEFT JOIN"
sSQL = sSQL + " 	( SELECT MemberID, 'High Score' AS Metric, Event, COUNT(Score) AS LastScoreCount"
sSQL = sSQL + " 		FROM usawsrank.Scores s"
sSQL = sSQL + " 		JOIN sanctions.dbo.TSchedul ts ON ts.TournAppID = LEFT(s.TourID,6)"
sSQL = sSQL + " 			WHERE MemberID = '"&SelectedMemberID&"'"
sSQL = sSQL + " 				AND TDateE<='"&LastYear12moEnd&"' AND TDateE>='"&LastYear12moStart&"'"
sSQL = sSQL + " 		GROUP BY MemberID, Event ) LSC"
sSQL = sSQL + " 	ON LSC.MemberID=TSC.MemberID AND LSC.Event=TSC.Event"

sSQL = sSQL + " 	UNION"
	
sSQL = sSQL + " 	SELECT TTC.MemberID, '3' AS MetricOrder,  TTC.Metric, TTC.Event, ThisTourCount AS ThisSummaryValue, LastTourCount AS LastSummaryValue"
sSQL = sSQL + " 	FROM"
sSQL = sSQL + " 	( SELECT MemberID, '# of Tournaments' AS Metric, Event, Count(DISTINCT TourID) AS ThisTourCount"
sSQL = sSQL + " 		FROM usawsrank.Scores s"
sSQL = sSQL + " 		JOIN sanctions.dbo.TSchedul ts ON ts.TournAppID = LEFT(s.TourID,6)"
sSQL = sSQL + " 			WHERE MemberID = '"&SelectedMemberID&"'"
sSQL = sSQL + " 				AND TDateE>='"&ThisYear12moStart&"'"
sSQL = sSQL + " 		GROUP BY MemberID, Event ) TTC"
sSQL = sSQL + " 	LEFT JOIN"
sSQL = sSQL + " 	( SELECT MemberID, 'High Score' AS Metric, Event, Count(DISTINCT TourID) AS LastTourCount"
sSQL = sSQL + " 		FROM usawsrank.Scores s"
sSQL = sSQL + " 		JOIN sanctions.dbo.TSchedul ts ON ts.TournAppID = LEFT(s.TourID,6)"
sSQL = sSQL + " 			WHERE MemberID = '"&SelectedMemberID&"'"
sSQL = sSQL + " 				AND TDateE<='"&LastYear12moEnd&"' AND TDateE>='"&LastYear12moStart&"'"
sSQL = sSQL + " 		GROUP BY MemberID, Event ) LTC"
sSQL = sSQL + " 	ON LTC.MemberID=TTC.MemberID AND LTC.Event=TTC.Event" 
sSQL = sSQL + " 	) AllMet"

sSQL = sSQL + " 	ORDER BY MetricOrder, CASE WHEN Event='S' THEN 1 WHEN Event='T' THEN 2 ELSE 3 END"  

Set rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable

END SUB



' -----------------------
  SUB RunRawScoresQuery
' -----------------------  
  
sSQL = "SELECT s.MemberID, s.TourID, TName AS Tournament, s.Event, s.Div, Round, Place, Score, TDateE"
sSQL = sSQL + " , MaxScore, s.Perf_Qual1, s.Perf_Qual2"
sSQL = sSQL + " , FirstName, LastName, Address1, Address2, City, State, Zip"

sSQL = sSQL + " FROM usawsrank.Scores s"
sSQL = sSQL + "	JOIN sanctions.dbo.TSchedul ts ON ts.TournAppID = LEFT(s.TourID,6)"
sSQL = sSQL + "	JOIN usawaterski.dbo.MemberShort m ON m.PersonID = RIGHT(s.MemberID,8)"
sSQL = sSQL + "	LEFT JOIN"
sSQL = sSQL + "	  ( SELECT MemberID, Event, MAX(Score) AS MaxScore"
sSQL = sSQL + "	      FROM usawsrank.Scores s"
sSQL = sSQL + "	         WHERE MemberID = '"&SelectedMemberID&"'"
sSQL = sSQL + "	      GROUP BY MemberID, Event) mx"
sSQL = sSQL + "	ON mx.MemberID=s.MemberID AND mx.Event=s.Event"
  
sSQL = sSQL + "	WHERE s.MemberID = '"&SelectedMemberID&"'"
sSQL = sSQL + "				AND TDateE>='"&ThisYear12moStart&"'"
sSQL = sSQL + "	ORDER BY TDateE DESC, s.Event, Score DESC"

Set rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable
	
END SUB  

%>