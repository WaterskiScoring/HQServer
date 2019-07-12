<!--#include file="settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->
<!--#include virtual="/rankings/tools_Definitions.asp"-->
<!--#include virtual="/rankings/tools_leagues.asp"-->
<!--#include virtual="/rankings/tools_mobile_version.asp"-->
<%

' --- Last update 7-26-2015 ---


'DefineTRAStyles



Dim ThisFileName, sPriorYear, process, MainImage, AdminMenuLevel
Dim TabColor

Dim RatingLevel
Dim rsList

Dim TeamTypeIDSelected, EventSelected, EventName
'Dim sLeagueSelected 
Dim sShowSQL

Dim ThisTournAppID, LastTournAppID, ThisStartDate, LastStartDate, DiffBetweenStartDates

Dim Team_ID, Team_Name, Team_Level, Created_Date, TeamStatus, Manager_MemberID, Manager_FirstName, Manager_LastName
Dim FirstName, LastName, TeamMemberStatus, Team_Type_ID, Team_Type_Description
Dim TeamRank, Team_Improvement, Total_Score_Applied, Team_LSY_BM, No_Team_Members, BenchmarkApplied, No_Scoring_Members
Dim MemberRankScore, Rank_LSY_BM, MemberDeltaApplied, MemberRankonTeam, MemberScoreUsed



ThisFileName="view-vteamstatus.asp"
AdminMenuLevel=Session("AdminMenuLevel")


ReadFormVariables 


' -- Displays banner and Head tags --
OpenState="vteamrankings"
DisplayHeadOpenBodyAndBannerTags OpenState



Get_TeamStatus_RecordSet_ForMobile			' -- Queries database for records matching TeamTypeIDSelected --
		
Display_Teamstatus_Detail								' -- Displays League Rankings --




' --- Writes the Closing tags for HTML - in tools_mobile_version.asp ---
DisplayCloseBodyAndHTMLTags



' ----------------------------------------------------------------------------------------------------------------
' --- END OF MAIN PROGRAM ---
' ----------------------------------------------------------------------------------------------------------------









' -----------------------
  SUB ReadFormVariables
' -----------------------  
  
process=TRIM(LCASE(request("process")))

' -- Primary Filter --
sLeagueSelected=TRIM(Request("sLeagueSelected"))


' --- Event and League ---
EventSelected=TRIM(Request("EventSelected"))
IF EventSelected="" THEN EventSelected="S"



SELECT CASE EventSelected
	CASE "S"
		EventName="Slalom"
	CASE "T"
		EventName="Trick"
	CASE "J"
		EventName="Jump"
END SELECT

TeamTypeIDSelected = Request("TeamTypeIDSelected")
IF TRIM(TeamTypeIDSelected)="" THEN TeamTypeIDSelected=0

RatingLevel=TRIM(Request("RatingLevel"))

' --- Control execution ---
sShowSQL = Request("sShowSQL")


END SUB  





' ---------------------------------
  SUB Display_Teamstatus_Detail
' ---------------------------------

%>
<div id="myteamlisting" style="padding:0px; border:0px solid white;">
	<form method="post">
	<div style="width:100%; margin:10px 0px 0px 0px; padding:0px; text-align:left;">		
			<span class="span20" style="margin-left:0px; padding-left:0px; text-align:right; font-size:12pt; color:yellow; border:0px solid white;">League:</span> 
			<span class="span75" style="text-align:left;">
			<%
				' --- Builds Team Type (Virtual League) drop down - in tools_leagues.asp - parameters: thiswidth,thisfont,onchangeaction --
				BuildTeamType_DropDown 13,14, "submit()"
			%>
			</span>
	</div>	
	</form>
	<div class="scroll" style="width:100%; margin-top:5px; padding:0px; margin-left:0px; height:420px; border:0px solid white;">
		<%   
		
		' --- Displays the filter dropdowns inside ---
		IF NOT rs.eof THEN 
				rs.movefirst
				LoopThruMyTeams
		ELSE
				DisplayNoListingFound
		END IF
		
		%>
	</div> <! -- Bottom of scroll box -- ->
</div> <! -- Bottom of div for hiding and displaying -- ->

<%

END SUB





' --------------------------
  SUB DisplayNoListingFound
' --------------------------  

'response.write("</div><div style=""color:red;"">HERE lin 516</div>")
%>
<div class="tabrankings" style="height:20px; margin-top:30px; padding-top:4px; background-color:<%=DefineLevelcolor%>; text-color:<%=Textcolor1%>; font-size:16px;">
	<span class="span90" style="color:white; text-align:center;"><b>No Rankings Found for Selection</b></span>
</div>
<%

END SUB




' --------------------------
  SUB LoopThruMyTeams
' --------------------------
 
MemberCount = 1
TeamTypeCount = 1

ThisTeam_ID = rs("Team_ID")
ThisTeam_Type_ID = rs("Team_Type_ID")
DO WHILE NOT rs.eof

		GetCurrentMemberLine

		IF TeamTypeCount = 1 THEN DisplayNewTeamTypeTab
		IF MemberCount=1 THEN DisplayTeamTab

		DisplayTeamMemberLine
		
		MemberCount = MemberCount + 1
		TeamTypeCount = TeamTypeCount + 1
		
		rs.movenext
		IF NOT rs.eof THEN 
				Team_ID=rs("Team_ID")
				Team_Type_ID = rs("Team_Type_ID")
		END IF		
		IF ThisTeam_ID<>Team_ID THEN 
				MemberCount = 1			' -- Restart counter to display tab for team --	
				ThisTeam_ID = rs("Team_ID")
				DisplayTeamBottomLine
		END IF
		IF ThisTeam_Type_ID<>Team_Type_ID THEN
				TeamTypeCount = 1
				ThisTeam_Type_ID = rs("Team_Type_ID")
		END IF		
	

LOOP

DisplayTeamBottomLine

END SUB



' ---------------------------
  SUB DisplayNewTeamTypeTab
' ---------------------------

%>
  <div class="tabrankings" style="background-color:<%=TabColor%>; font-size:14pt; font-weight:bold; height:20px; margin-top:0px; padding-top:3px; padding-bottom:2px;" > 
	  <span class="span75" style="color:black; text-align:left; margin-left:2px;"><%= Team_Type_Description %></span>	  
	  <span class="span20" style="color:black; text-align:right;">Score</span>	  	  
	</div>
<%


END SUB






' -------------------
  SUB DisplayTeamTab
' -------------------
'

%> 
  <div class="tabrankings" style="height:34px; padding:0px 2px 0px 6px; margin:0px 0px 0px 0px; background-color:<% =TabColor %>;" >
		<span class="span80" style="font-size:12pt; font-weight:bold;"><%= Team_Name %></span>
		<span class="span15" style="text-align:right; font-size:12pt; font-weight:bold;"><%= FormatNumber(Team_Improvement,2) %></span>
	  <br>
  	<span style="width=13%; color:black; font-size:10pt; font-weight:normal; border:0px solid white;">Members:</span>
  	<span class="span10" style="width:8%; color:blue; font-size:10pt; font-weight:normal; text-align:left;"><%= No_Team_Members %></span>
  	<span class="span5" style="color:black; font-size:10pt; text-align:right; font-weight:normal;">ID:</span>
	  <span class="span15" style="color:blue; font-size:10pt; font-weight:normal; text-align:left;"><%= Team_ID %></span>

		<span class="span15" style="text-align:right; font-size:10pt; font-weight:normal;">Current</span>
		<span class="span15" style="text-align:right; font-size:10pt; font-weight:normal;">Base</span>
		<span class="span15" style="text-align:right; font-size:10pt; font-weight:normal; margin-left:2px;">Applied</span>
	</div>
<%	

END SUB



' ---------------------------
  SUB DisplayTeamMemberLine
' ---------------------------

MemberScoreUsed=""
IF CInt(MemberRankOnTeam)<=CInt(No_Scoring_Members) THEN MemberScoreUsed="*"
IF IsNull(MemberDeltaApplied) THEN MemberDeltaApplied = 0
IF IsNull(Rank_LSY_BM) THEN Rank_LSY_BM = 0
' response.write("</div><br><div style=""color:white;"">MemberDeltaApplied = "&MemberDeltaApplied)
' response.end		
%>
  <div class="rankingsbody" style="border:0px solid white; padding:0px 6px 0px 2px; margin:0px 0px 0px 2px;">
		<span class="span5" style="text-align:center; font-weight:normal;"><%= MemberRankOnTeam %></span>
		<span class="span45" style="text-align:left; font-weight:normal;"><%= FirstName %>&nbsp;<%= LastName %></span>
		<span class="span15" style="width:15%; text-align:right; font-weight:normal;"><%= FormatNumber(MemberRankScore,2) %></span>
		<span class="span15" style="width:15%; text-align:right; font-weight:normal;"><%= FormatNumber(Rank_LSY_BM,2) %></span>
		<span class="span15" style="width:11%; text-align:right; font-weight:normal;"><%= FormatNumber(MemberDeltaApplied,2) %></span>
		<span class="span5" style="width:5px; text-align:left; font-weight:normal; color:red;"><%= MemberScoreUsed %></span>
	</div>
<%

END SUB  




' ---------------------------
  SUB DisplayTeamBottomLine
' ---------------------------  

%>
<div class="tourbottom" style="width:97%; background-color:#FFFFFF; height:12px; margin:0px 0px 0px 2px; padding:0px 0px 2px 5px; ">
		<span class="span95">&nbsp;</span>
</div>
<%

END SUB







' --------------------------
  SUB GetCurrentMemberLine
' --------------------------

    Team_ID=rs("Team_ID")
		Team_Name=rs("Team_Name")
		'Team_Level=rs("Team_Level")
		' Created_Date=rs("Created_Date")
		' TeamStatus=rs("TeamStatus")
		'Manager_MemberID=rs("Manager_MemberID")
		'Manager_FirstName=rs("Manager_FirstName")
		'Manager_LastName=rs("Manager_LastName")
		FirstName=rs("FirstName")
		LastName=rs("LastName")
		'TeamMemberStatus=rs("TeamMemberStatus")
		MemberRankonTeam=rs("MemberRankonTeam")
		Team_Type_ID=rs("Team_Type_ID")
		Team_Type_Description=rs("Team_Type_Description")

		TeamRank=rs("TeamRank")
		Team_Improvement=rs("Team_Improvement")
		Total_Score_Applied=rs("Total_Score_Applied")
		Team_LSY_BM=rs("Team_LSY_BM") 
		No_Team_Members=rs("No_Team_Members")
		BenchmarkApplied=rs("BenchmarkApplied")

		MemberRankScore=rs("MemberRankScore")
		Rank_LSY_BM=rs("Rank_LSY_BM")
		MemberDeltaApplied=rs("MemberDeltaApplied")
		No_Scoring_Members = rs("No_Scoring_Members")



		SELECT CASE TRIM(Team_Type_ID)
				CASE 1 
						TabColor=scolor01
				CASE 2 
						TabColor=scolor02
				CASE 3 
						TabColor=scolor03
				CASE 4 
						TabColor=scolor04
				CASE 5 
						TabColor=scolor05
				CASE 6 
						TabColor=scolor06
				CASE 7 
						TabColor=scolor07
				CASE 8 
						TabColor=scolor08
				CASE 9 
						TabColor=scolor09
				CASE 10 
						TabColor=scolor10
				CASE ELSE
						TabColor=scolor05							
		END SELECT
END SUB








' ----------------------------------------
  SUB Get_TeamStatus_RecordSet_ForMobile 
' ----------------------------------------
  
  
  
sSQL = " SELECT DT.Team_ID, DT.Team_Name, DT.MemberID, DT.FirstName, DT.LastName, DT.State, DT.Event, DT.RankScore"
sSQL = sSQL + " 		, DT.Team_Type_ID, DT.Rank_LSY_BM, DT.Rank_2PYSY_BM, DT.MemberDeltaApplied"
sSQL = sSQL + " 		, DT.Team_Type_Description"
sSQL = sSQL + " 		, DT.MemberDeltaApplied"
sSQL = sSQL + " 		, COALESCE(DT.RankScore,0) AS MemberRankScore"
sSQL = sSQL + " 		, DT.Rank_LSY_BM"
sSQL = sSQL + " 		, TeamRankOverall AS TeamRank"
sSQL = sSQL + " 		, SU.No_Scoring_Members, SU.Team_Improvement, SU.BenchmarkApplied"
sSQL = sSQL + " 		, SU.No_Team_Members, SU.Total_Score_AllMembers, SU.Total_Score_Applied"
sSQL = sSQL + " 		, SU.TeamSumDeltaApplied, SU.Team_LSY_BM"
sSQL = sSQL + " 		, RANK() OVER(Partition By DT.Team_ID ORDER BY DT.MemberDelta DESC) AS MemberRankOnTeam"		

		
sSQL = sSQL + " 				FROM"
		
sSQL = sSQL + " 				( SELECT * FROM usawsrank.VTeam_Status_ForMobile WHERE Team_Type_ID = "&TeamTypeIDSelected&" ) DT"
sSQL = sSQL + " 				LEFT JOIN"
sSQL = sSQL + " 				( "
sSQL = sSQL + " 					SELECT RANK() OVER(Partition By Team_ID ORDER BY SUM(MemberDeltaApplied) DESC) AS TeamRankOverall"
sSQL = sSQL + " 						, Team_Type_ID, Team_Type_Description, Team_ID, Team_Name"
sSQL = sSQL + " 						, SUM(MemberDeltaApplied) AS Team_Improvement"
sSQL = sSQL + " 						, ROUND(SUM(MemberDeltaApplied),2) AS TeamSumDeltaApplied"
sSQL = sSQL + " 						, SUM(Rank_LSY_BM) AS Team_LSY_BM"
sSQL = sSQL + " 						, MAX(No_Scoring_Members) AS No_Scoring_Members"
sSQL = sSQL + " 						, SUM(BenchmarkApplied) AS BenchmarkApplied"
sSQL = sSQL + " 						, COUNT(MemberID) AS No_Team_Members"
sSQL = sSQL + " 						, SUM(MemberRankApplied) AS Total_Score_Applied"
sSQL = sSQL + " 						, SUM(RankScore) AS Total_Score_AllMembers"

sSQL = sSQL + " 					FROM usawsrank.VTeam_Status_ForMobile"
sSQL = sSQL + " 						WHERE Team_Type_ID = "&TeamTypeIDSelected
sSQL = sSQL + " 					GROUP BY Team_Type_ID, Team_Type_Description, Team_ID, Team_Name ) SU"
sSQL = sSQL + " 				ON SU.Team_ID=DT.Team_ID"
		
sSQL = sSQL + " 		ORDER BY DT.Team_Type_ID, DT.Team_Type_Description, TeamSumDeltaApplied DESC"


'response.write(sSQL)
'response.end
Set rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable

END SUB  




%>