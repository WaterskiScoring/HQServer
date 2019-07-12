<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_leagues.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->
<!--#include virtual="/rankings/tools_registration.asp"-->
<!--#include virtual="/rankings/Tools_TournamentListQuery.asp"-->
<!--#include virtual="/rankings/tools_mobile_version.asp"-->
<%



' ------------------------------------------------
' --- Dimensions variables used in this module ---
' ------------------------------------------------

Dim ThisFileName

Dim rowCount, i
Dim sMemberID, sFullName, sEventName

Dim EventSelected, DivSelected, FilterSelected, FederationSelected, SkiYearSelected

Dim Team_ID, Team_Name, Team_Level, Manager_MemberID, Manager_FirstName, Manager_LastName, TeamStatus
Dim MemberID, FirstName, LastName, TeamMemberStatus
Dim Team_Type_ID, Team_Type_Description, Created_Date, No_Team_Members
Dim MemberCount
Dim TeamMemberStatusText, TeamMemberStatusTextColor, TeamStatusText, TeamStatusTextColor
Dim TabColor





ThisFileName="virtualteam_m.asp"

' --- Names related programs for linking ---
SearchFileName = "search-memberHQ.asp"
RankingsMobileFilename="view-standings_m.asp"
TournamentsMobileFilename="view-tournaments_m.asp"
TeamsMobileFilename="virtualteam_m.asp"
LocalVarFileName="Test_localstorage_SET.asp"
MenuFileName = "mainmenu_m.asp"



' ------------------------------------
' --- Reads NVP's from querystring ---
' ------------------------------------
'ReadRankingsFormVariables



' --- Displays the html, head and opening body tag ---
OpenState="myteamlisting"
DisplayHeadOpenBodyAndBannerTags OpenState




' ------------------------------------------------------------------------------------------------           
' -------------------------------  BEGINS WRITING HEADERS AND RANKINGS  --------------------------
' ------------------------------------------------------------------------------------------------

	

' --- Displays the menu for view tournaments --- 
' DisplayMenuButtons_ViewTournaments
%>
<input type="hidden" id="sMemberID_Hidden_InRankingsSettings" name="sMemberID" value="">
<input type="hidden" id="sName_InRankingsSettings" name="sName_InRankingsSettings" value="">
<%


' *** LOGIC HERE ***
RunMyTeamListingQuery

' --- Displays listing of teams for this mobile user ---
DisplayMyTeamListing



'' --- Writes the Closing tags for HTML ---
DisplayCloseBodyAndHTMLTags







' ---------------------------------------------------------------------------------------------------------------
' ----------------------   END OF MAIN CODE HERE  ---------------------------------------------------------------
' ---------------------------------------------------------------------------------------------------------------
















' ----------------------------------------
   SUB DisplayMyTeamListing
' ----------------------------------------

Titlecolor=Textcolor2
TempManagerName = "Mark Crone Temp"

%>
<div id="myteamlisting" style="display:block-inline">
	<div class="filterdropdownline" style="height:25px; margin-top:8px; font-size:14px;">
		<span class="span30" style="text-align:right; color:white;">Manager:</span>
  	<span class="span65" style="text-align:left; color:#6495ED;"><%=TempManagerName%></span>
	</div>
	<div class="scroll" style="height:300px;">
		<%   
		
		' --- Displays the filter dropdowns inside ---
		LoopThruMyTeams
		
		%>
	</div> <! -- Bottom of scroll box -- ->
</div> <! -- Bottom of div for hiding and displaying -- ->
<%

DisplayMyTeamListingNavigationButtons


END SUB



' -----------------------------------------
  SUB DisplayMyTeamListingNavigationButtons
' -----------------------------------------  

%><div id="tourmenubuttons" class="menucell" style="padding:0px; margin:0px; background-color:black;">
	<TABLE align=center width=100% style="padding:0px; margin:0px; border:0px solid; border-color:yellow;"">
		<tr>
			<td width="46%" height="30px" background="images/buttons/ButtonBlank.png" style="background-position:center center; background-repeat:no-repeat; border:0px solid; border-color:#FFFFFF; background-size:75px; margin:0px; padding:0px; text-align:center;">
				<font size="2" color="blue"><a href="<%=TournamentsMobileFilename%>" style="text-decoration:none;">NEW</a></font>
			</td>
			<td width="46%" background="images/buttons/ButtonBlank.png" style="background-position:center center; background-repeat:no-repeat; border:0px solid; border-color:#FFFFFF; background-size:75px; margin:0px; padding:0px; text-align:center;">
				<font size="2" color="blue"><a href="" style="text-decoration:none;">RETURN</a></font>
			</td>				
		</tr>
	</TABLE>
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


' --------------------------
  SUB GetCurrentMemberLine
' --------------------------

    Team_ID=rs("Team_ID")
		Team_Name=rs("Team_Name")
		Team_Level=rs("Team_Level")
		Created_Date=rs("Created_Date")
		TeamStatus=rs("TeamStatus")
		Manager_MemberID=rs("Manager_MemberID")
		Manager_FirstName=rs("Manager_FirstName")
		Manager_LastName=rs("Manager_LastName")
		FirstName=rs("FirstName")
		LastName=rs("LastName")
		TeamMemberStatus=rs("TeamMemberStatus")
		Team_Type_ID=rs("Team_Type_ID")
		Team_Type_Description=rs("Team_Type_Description")
		No_Team_Members = rs("No_Team_Members")
		
		SELECT CASE TRIM(TeamMemberStatus)
				CASE "A" 
						TeamMemberStatusText="Accepted"
						TeamMemberStatusTextColor="green"
				CASE "I"
						TeamMemberStatusText="Invitation Pending"			
						TeamMemberStatusTextColor="orange"
				CASE "D"
						TeamMemberStatusText="Declined"			
						TeamMemberStatusTextColor="red"
				CASE ELSE
						TeamMemberStatusText="Invitation Needed"			
						TeamMemberStatusTextColor="red"
		END SELECT			

		SELECT CASE TRIM(TeamStatus)
				CASE "A" 
						TeamStatusText="Active"
						TeamStatusTextColor="green"
				CASE "P"
						TeamStatusText="Pending"			
						TeamStatusTextColor="orange"
				CASE "H"
						TeamStatusText="Inactive"			
						TeamStatusTextColor="red"
				CASE ELSE
						TeamStatusText="Unknown"			
						TeamStatusTextColor="red"
		END SELECT			

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




' ---------------------------
  SUB DisplayNewTeamTypeTab
' ---------------------------

%>
  <div class="tabrankings" style="background-color:<%=TabColor%>; font-size:14px; font-weight:bold; height:20px; margin-top:0px; padding-top:0px;" >
	  <span style="width=15%; font-weight:normal; color:black; text-align:left; border:0px solid white;">Team Type:</span>	  
	  <span class="span45" style="color:blue; font-weight:normal; text-align:left;"><%= Team_Type_Description %></span>	  
	</div>
<%


END SUB



' -------------------
  SUB DisplayTeamTab
' -------------------

%> 
  <div class="tabrankings" style="background-color:<%=TabColor%>; height:25px; margin-top:0px; padding-top:0px;" >
		<span class="span80" style="font-size:12px; font-weight:bold;"><%= Team_Name %></span>
		<span class="span15" style="text-align:right;"><b>Edit</b></span>
	  <br>
  	<span style="width=15%; color:black; font-weight:normal; border:0px solid white;"># Members:</span>
  	<span class="span10" style="color:blue; font-weight:normal; text-align:left;"><%= No_Team_Members %></span>
  	<span class="span10" style="color:black; text-align:right; font-weight:normal;">ID:</span>
	  <span class="span20" style="color:blue; font-weight:normal; text-align:left;"><%= Team_ID %></span>
	  <span class="span10" style="font-weight:normal; color:black; text-align:right;">Status:</span>	  
	  <span class="span20" style="color:<%=TeamStatusTextColor%>; font-weight:normal; text-align:left;"><%= TeamStatusText %></span>	  
	</div>
<%	

END SUB


' ---------------------------
  SUB DisplayTeamMemberLine
' ---------------------------
%>
  <div class="rankingsbody" >
		<span class="span65" style="color:black; font-weight:normal;"><%= FirstName %>&nbsp;<%= LastName %></span>
		<span class="span30" style="color:<%=TeamMemberStatusTextColor%>; font-weight:normal;"><%= TeamMemberStatusText %></span>
	</div>
<%

END SUB  


' ---------------------------
  SUB DisplayTeamBottomLine
' ---------------------------  

%>
<div class="tourbottom"  style="background-color:#FFFFFF; height:7px;">
		<span class="span95">&nbsp;</span>
</div>
<%

END SUB



' ---------------------------
  SUB RunMyTeamListingQuery
' ---------------------------


ThisUser = "600006433"
ThisUser = "000001151"

' --- Move these definitions to settingshq ---
'V_TeamTableName="usawsrank.V_Team"
'V_Team_MembersTableName="usawsrank.V_Team_Members"

sSQL = "SELECT t.Team_ID, Team_Name, t.Team_Type_ID, Team_Level, Manager_MemberID, t.Created_Date"
sSQL = sSQL + ", m.FirstName, m.LastName, tmem.MemberID"
sSQL = sSQL + ", m2.FirstName AS Manager_FirstName, m2.LastName AS Manager_LastName"
sSQL = sSQL + ", tmem.Status AS TeamMemberStatus"
sSQL = sSQL + ", t.Status AS TeamStatus, Team_Type_Description, tcnt.No_Team_Members"
sSQL = sSQL + " FROM "&V_TeamMembersTableName&" tmem"
sSQL = sSQL + " JOIN "&V_TeamTableName&" t ON t.Team_ID=tmem.Team_ID"
sSQL = sSQL + " JOIN "&V_TeamTypeTableName&" tt ON tt.Team_Type_ID=t.Team_Type_ID"
sSQL = sSQL + " LEFT JOIN "&MemberTableName&" m ON RIGHT(tmem.MemberID,8)=m.PersonID"
sSQL = sSQL + " LEFT JOIN "&MemberTableName&" m2 ON RIGHT(t.Manager_MemberID,8)=m.PersonID"

sSQL = sSQL + " LEFT JOIN "
sSQL = sSQL + "   ( SELECT Team_ID, COUNT(*) AS No_Team_Members FROM "&V_TeamMembersTableName
sSQL = sSQL + "      GROUP BY Team_ID ) tcnt"  
sSQL = sSQL + " ON tcnt.Team_ID=tmem.Team_ID"

sSQL = sSQL + " WHERE Manager_MemberID='"&ThisUser&"'"
sSQL = sSQL + " ORDER BY t.Team_Type_ID, tmem.Team_id, m.LastName, m.FirstName"

tt=2
IF tt=1 THEN
		%></div><div style="color:red;"><%
		response.write(sSQL)
		response.end
END IF

Set rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable


END SUB  






' ---------------------------
  SUB RunMyTeamListingQuery_NOT 
' ---------------------------

TeamTypeIDSelected=2


sSQL = "SELECT"
sSQL = sSQL + " RANK() OVER(Partition By Team_Type_ID ORDER BY SUM(MemberDeltaApplied) DESC) AS Rank"
sSQL = sSQL + ", Team_ID, Team_Name, Created_Date"

sSQL = sSQL + ", SUM(MemberDeltaApplied) AS Team_Improvement"
sSQL = sSQL + ", SUM(Rank_LSY_BM) AS Team_LSY_BM"
sSQL = sSQL + ", MAX(No_Scoring_Members) AS No_Scoring_Members"
sSQL = sSQL + ", SUM(BenchmarkApplied) AS BenchmarkApplied"
sSQL = sSQL + ", COUNT(MemberID) AS No_Team_Members"


sSQL = sSQL + ", SUM(MemberRankApplied) AS Total_Score_Applied"
sSQL = sSQL + ", SUM(RankScore) AS Total_Score_AllMembers"

sSQL = sSQL + " FROM"
sSQL = sSQL + " ("

sSQL = sSQL + " SELECT Team_Type_Description, ttype.Team_Type_ID"
sSQL = sSQL + ", tmem.Team_ID, Team_Name, t.Created_Date, tmem.MemberID, FirstName, LastName, State"
sSQL = sSQL + ", tmem.Event"
sSQL = sSQL + ", RankScore"
sSQL = sSQL + ", Rank_LSY_BM"
sSQL = sSQL + ", Rank_2PYSY_BM"
sSQL = sSQL + ", RankScore - Rank_LSY_BM AS MemberDelta"
sSQL = sSQL + ", RANK() OVER(Partition By tmem.Team_ID ORDER BY RankScore - Rank_LSY_BM DESC) AS TeamRank"
sSQL = sSQL + ", ttype.max_scoring AS No_Scoring_Members"

sSQL = sSQL + ", CASE WHEN RANK() OVER(Partition By tmem.Team_ID ORDER BY RankScore - Rank_LSY_BM DESC)<=ttype.max_scoring THEN RankScore ELSE 0 END AS MemberRankApplied"
sSQL = sSQL + ", CASE WHEN RANK() OVER(Partition By tmem.Team_ID ORDER BY RankScore - Rank_LSY_BM DESC)<=ttype.max_scoring THEN RankScore - Rank_LSY_BM ELSE 0 END AS MemberDeltaApplied"
sSQL = sSQL + ", CASE WHEN RANK() OVER(Partition By tmem.Team_ID ORDER BY RankScore - Rank_LSY_BM DESC)<=ttype.max_scoring THEN Rank_LSY_BM ELSE 0 END AS BenchmarkApplied"
	
sSQL = sSQL + " FROM usawsrank.V_Team_Members tmem"
sSQL = sSQL + " JOIN usawaterski.dbo.membershort AS memsht ON CAST(RIGHT(tmem.MemberID,8) AS INT)=memsht.PersonID"
	
sSQL = sSQL + " LEFT JOIN"
sSQL = sSQL + " ( SELECT *"
sSQL = sSQL + "	FROM usawsrank.V_Team) t"
sSQL = sSQL + " ON t.Team_ID=tmem.Team_ID"

sSQL = sSQL + " LEFT JOIN"
sSQL = sSQL + " ( SELECT MemberID, Event, MAX(RankScore) AS Rank_LSY_BM"
sSQL = sSQL + " FROM usawsrank.Rankings" 
sSQL = sSQL + " WHERE SkiYearID=20"
sSQL = sSQL + " GROUP BY MemberID, Event) rPY"
sSQL = sSQL + " ON rPY.MemberID=tmem.MemberID AND rPY.Event=tmem.Event"

sSQL = sSQL + " LEFT JOIN"
sSQL = sSQL + " ( SELECT MemberID, Event, MAX(RankScore) AS Rank_2PYSY_BM"
sSQL = sSQL + " FROM usawsrank.Rankings" 
sSQL = sSQL + " WHERE SkiYearID=17"
sSQL = sSQL + " GROUP BY MemberID, Event) r2PY" 
sSQL = sSQL + " ON r2PY.MemberID=tmem.MemberID AND r2PY.Event=tmem.Event"

sSQL = sSQL + " LEFT JOIN"
sSQL = sSQL + " ( SELECT MemberID, Event, MAX(RankScore) AS RankScore"
sSQL = sSQL + " FROM usawsrank.Rankings" 
sSQL = sSQL + " WHERE SkiYearID=1"
sSQL = sSQL + " GROUP BY MemberID, Event) r1" 
sSQL = sSQL + " ON r1.MemberID=tmem.MemberID AND r1.Event=tmem.Event"


sSQL = sSQL + " JOIN usawsrank.V_Team_Type ttype ON ttype.Team_Type_ID=t.Team_Type_ID"	

sSQL = sSQL + " WHERE t.Team_Type_ID="&TeamTypeIDSelected

sSQL = sSQL + " ) Summary"

sSQL = sSQL + " GROUP BY Team_Type_ID, Team_Type_Description, Team_ID, Team_Name"
sSQL = sSQL + " ORDER BY Team_Type_ID, Team_Type_Description, ROUND(SUM(MemberDeltaApplied),2) DESC"

'response.write(sSQL)
'response.end
Set rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable


END SUB











%>
