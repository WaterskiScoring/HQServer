<!--#include file="settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->
<!--#include virtual="/rankings/tools_Definitions.asp"-->
<!--#include virtual="/rankings/tools_leagues.asp"-->
<%

' --- Last update 7-26-2015 ---


DefineTRAStyles



Dim ThisFileName, sPriorYear, process, MainImage, AdminMenuLevel

Dim RatingLevel
Dim rsList

Dim TeamTypeIDSelected, EventSelected, EventName
'Dim sLeagueSelected 
Dim sShowSQL

Dim ThisTournAppID, LastTournAppID, ThisStartDate, LastStartDate, DiffBetweenStartDates



TourTableWidth=675
TabWidth = 1000  	' --- Used in case where report does not have specific parameters

ThisFileName="view_vteamstatus.asp"
AdminMenuLevel=Session("AdminMenuLevel")




ReadFormVariables 



sAction=Request("Action")
IF sAction="Return to Menu" THEN 
		process="return"
END IF



' --------------------------------------------------------------------------
' --- Defines the image to be displayed in the drop downs box background ---
' --------------------------------------------------------------------------

WhatDropdownImage EventSelected













SELECT CASE process
	CASE "return"
		response.redirect("/rankings/defaultHQ.asp")

	CASE "v_teammemberstatus"
		'WriteIndexPageHeader
		PageTitle="V-Team Member Status Detail"
		PageSubTitle="Current Rank Score vs 2014 Ski EOY Rank Score<br>&nbsp;&nbsp;Order By Member Improvement"

		v_TeamMemberStatus
		CreatePageHead 725
		IF NOT rs.eof THEN DisplayResult 900
		'WriteIndexPageFooter

	CASE "v_teamstatus"
		'WriteIndexPageHeader
		PageTitle="V-Team Scoring Status"
		PageSubTitle="Sum of Member Improvement as Team Improvement for Ranking<br>&nbsp;&nbsp;All & Top 2 Members Scoring"

		v_TeamStatus
		CreatePageHead 725
		IF NOT rs.eof THEN DisplayResult 900
		'WriteIndexPageFooter

	CASE "v_teamtypes"
		'WriteIndexPageHeader
		PageTitle="V-Team Type List"
		PageSubTitle="Defines Make-up of each Team Type"

		v_TeamType
		CreatePageHead 725
		IF NOT rs.eof THEN DisplayResult 900
		'WriteIndexPageFooter


END SELECT























' ----------------------------------------------------------------------------------------------------------------
' --- END OF MAIN PROGRAM ---
' ----------------------------------------------------------------------------------------------------------------









' -----------------------
  SUB ReadFormVariables
' -----------------------  
  
process=TRIM(LCASE(request("process")))

' --- Event and League ---
EventSelected=TRIM(Request("EventSelected"))
IF EventSelected="" THEN EventSelected="J"

sLeagueSelected=TRIM(Request("sLeagueSelected"))

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







' ------------------------------
  SUB DisplayResult (tabwidth)
' ------------------------------



	rs.movefirst



	%>
	<TABLE class="innertable" Align=center WIDTH=<%=tabwidth%>px >
		<%


		' ---------------  Displays table HEADINGS  ----------------------
		SELECT CASE process
				CASE "v_teamstatus"
						Display_TeamStatus_Heading

		END SELECT



		' --------------  Display table data here with paging --------------------------
		RowCount = 1 
		DO WHILE NOT rs.eof

				SELECT CASE process
						CASE "v_teamstatus"
								Display_TeamStatus_DataRow
				END SELECT 

	
				rowCount = rowCount + 1
				rs.movenext
	
		LOOP 
	
	
	
		%>
	</TABLE>
<br><br>
<%



' --- Displays Table Footer ---
SELECT CASE process
		CASE "v_teamstatus"
				Display_TeamStatus_Footer

		CASE "v_teammemberstatus"
				DisplayTeamMemberStatus_Footer
		
END SELECT



END SUB







' -------------------------------
  SUB Display_TeamStatus_DataRow
' -------------------------------

			%>
			<TR>
				<td ALIGN="center" style="<%=RowColor%>"><font SIZE="1"><%= Rank %></font></td>
				<td ALIGN="center" style="<%=RowColor%>"><font SIZE="1"><%= Team_ID %></font></td>
				<td ALIGN="center" style="<%=RowColor%>"><font SIZE="1"><%= Team_Name %></font></td>
				<td ALIGN="center" style="<%=RowColor%>"><font SIZE="1"><%= Team_Improvement %></font></td>
				<td ALIGN="center" style="<%=RowColor%>"><font SIZE="1"><%= Total_Score_Top2 %></font></td>
			</TR>			
			<% 

END SUB





' ---------------------
  SUB v_TeamStatus
' ---------------------

sSQL = "SELECT"
sSQL = sSQL + " RANK() OVER(Partition By Team_Type_ID ORDER BY ROUND(SUM(MemberDeltaApplied),2) DESC) AS Rank"
sSQL = sSQL + ", Team_ID, Team_Name"

sSQL = sSQL + ", ROUND(SUM(MemberDeltaApplied),2) AS Team_Improvement"
sSQL = sSQL + ", SUM(NumMembersApplied) AS No_Scoring_Members"
' sSQL = sSQL + ", ROUND(SUM(MemberDelta),2) AS [Team Improvement<br>All Scoring]"
sSQL = sSQL + ", COUNT(MemberID) AS [No_Team_Members]"

sSQL = sSQL + ", ROUND(SUM(MemberRankApplied),2) AS Total_Score_Top2"
sSQL = sSQL + ", ROUND(SUM(RankScore),2) AS Total_Score_AllScoring"

sSQL = sSQL + " FROM"
sSQL = sSQL + " ("

sSQL = sSQL + " SELECT Team_Type_Description, ttype.Team_Type_ID"
sSQL = sSQL + ", tmem.Team_ID, Team_Name, tmem.MemberID, FirstName, LastName, State"
sSQL = sSQL + ", tmem.Event"
sSQL = sSQL + ", RankScore"
sSQL = sSQL + ", Rank_LSY_BM"
sSQL = sSQL + ", Rank_2PYSY_BM"
sSQL = sSQL + ", RankScore - Rank_LSY_BM AS MemberDelta"
sSQL = sSQL + ", RANK() OVER(Partition By tmem.Team_ID ORDER BY RankScore - Rank_LSY_BM DESC) AS TeamRank"

sSQL = sSQL + ", CASE WHEN RANK() OVER(Partition By tmem.Team_ID ORDER BY RankScore)<=2 THEN RankScore ELSE 0 END AS MemberRankApplied"
sSQL = sSQL + ", CASE WHEN RANK() OVER(Partition By tmem.Team_ID ORDER BY RankScore - Rank_LSY_BM DESC)<=2 THEN RankScore - Rank_LSY_BM ELSE 0 END AS MemberDeltaApplied"
sSQL = sSQL + ", CASE WHEN RANK() OVER(Partition By tmem.Team_ID ORDER BY RankScore - Rank_LSY_BM DESC)<=2 THEN 1 ELSE 0 END AS NumMembersApplied"
	
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