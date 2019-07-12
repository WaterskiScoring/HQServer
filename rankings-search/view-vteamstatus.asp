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

ThisFileName="view-vteamstatus.asp"
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





process="v_teamstatus"





SELECT CASE process
	CASE "return"
		response.redirect("/rankings/defaultHQ.asp")

	CASE "v_teammemberstatus"
		'WriteIndexPageHeader
		PageTitle="V-Team Member Status Detail"
		PageSubTitle="Current Rank Score vs 2014 Ski EOY Rank Score<br>&nbsp;&nbsp;Order By Member Improvement"

		v_TeamMemberStatus

		IF NOT rs.eof THEN DisplayResult 900
		'WriteIndexPageFooter

	CASE "v_teamstatus"
		WriteIndexPageHeader
		
		tabwidth=725
		PageTitle="Virtual Team Ranking"
		PageSubTitle="Ranking Based on Scoring Members"

		Get_TeamStatus_RecordSet
		
		CreatePageHead (tabwidth)
		IF NOT rs.eof THEN Display_Teamstatus tabwidth

		WriteIndexPageFooter

	CASE "v_teamtypes"
		'WriteIndexPageHeader
		PageTitle="V-Team Type List"
		PageSubTitle="Defines Make-up of each Team Type"

		v_TeamType
		CreatePageHead 725

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
IF EventSelected="" THEN EventSelected="S"

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







' ---------------------------------
  SUB Display_Teamstatus (tabwidth)
' ---------------------------------

'response.write("<br>Line 181")

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
'		CASE "v_teamstatus"
'				Display_TeamStatus_Footer

'		CASE "v_teammemberstatus"
'				DisplayTeamMemberStatus_Footer
		
END SELECT



END SUB




' -------------------------------
  SUB Display_TeamStatus_DataRow
' -------------------------------

			Rank=rs("Rank")
			Team_ID=rs("Team_ID")
			Team_Name=rs("Team_Name")
			Team_Improvement=rs("Team_Improvement")
			Total_Score_Applied=rs("Total_Score_Applied")
			Team_LSY_BM=rs("Team_LSY_BM") 
			No_Team_Members=rs("No_Team_Members")
			BenchmarkApplied=rs("BenchmarkApplied")
			%>
			<TR>
				<td ALIGN="center" style="<%=RowColor%>"><font SIZE="1"><%= Rank %></font></td>
				<td ALIGN="center" style="<%=RowColor%>"><font SIZE="1"><%= Team_ID %></font></td>
				<td ALIGN="left" style="<%=RowColor%>"><font SIZE="1"><%= Team_Name %></font></td>
				<td ALIGN="center" style="<%=RowColor%>"><font SIZE="1"><%= No_Team_Members %></font></td>
				<td ALIGN="right" style="<%=RowColor%>"><font SIZE="1"><%= FormatNumber(Team_Improvement,2) %></font></td>
				<td ALIGN="right" style="<%=RowColor%>"><font SIZE="1"><%= FormatNumber(Total_Score_Applied,2) %></font></td>
				<td ALIGN="right" style="<%=RowColor%>"><font SIZE="1"><%= FormatNumber(BenchmarkApplied,2) %></font></td>
			</TR>			
			<% 

END SUB



' ----------------------------------
	SUB Display_TeamStatus_Heading
' ----------------------------------			

		Rank_Title="Based on Team Improvement for Scoring Members"
		Team_ID_TTitle="Team ID"
		Advance_Title="Calculated by summing the Member Improvement (for Scoring Members) "	
		Combined_RankScore_Title="Sum of Scoring Members Current Ranking Score"
		Combined_Benchmark_Title="Current formular uses 2014 Ski Year Ranking Score "
		%>
	  <TR>
	  	<th ALIGN="center" width=10% style="<%=RowColor%>"><font color="#FFFFFF" face="<%=font1%>" SIZE="<%=fontsize1%>">Rank</font></th>
	  	<th ALIGN="center" width=10% style="<%=RowColor%>"><font color="#FFFFFF" face="<%=font1%>" SIZE="<%=fontsize1%>">Team ID</font></th>
	  	<th ALIGN="left" width=25% style="<%=RowColor%>"><font color="#FFFFFF" face="<%=font1%>" SIZE="<%=fontsize1%>">Team Name</font></th>
	  	<th ALIGN="center" width=10% style="<%=RowColor%>"><font color="#FFFFFF" face="<%=font1%>" SIZE="<%=fontsize1%>" title="<%=Advance_Title%>"># Team<br>Members</font></th>
	  	<th ALIGN="right" width=15% style="<%=RowColor%>"><font color="#FFFFFF" face="<%=font1%>" SIZE="<%=fontsize1%>" title="<%=Advance_Title%>">Ranking Score<br>Advancement</font></th>

	  	<th ALIGN="right" width=15% style="<%=RowColor%>" title="<%=Combined_RankScore_Title%>"><font color="#FFFFFF" face="<%=font1%>" SIZE="<%=fontsize1%>">Combined<br>Rank Score</font></th>
	  	<th ALIGN="right" width=15% style="<%=RowColor%>" title="<%=Combined_Benchmark_Title%>"><font color="#FFFFFF" face="<%=font1%>" SIZE="<%=fontsize1%>">Combined<br>Benchmark</font></th>
	  </TR>
	  <%

END SUB


' -------------------------------
	SUB Display_TeamStatus_Footer
' -------------------------------	
		%>
		<TABLE class="innertable" Align=center style="width:<%=tabwidth%>px"; height:"200px; ">
			<tr>
				<td align="left" style="width:100%; height:250px; vertical-align:top;">
				<b>TERMS & DEFINITIONS:</b>
					<br><br>
					<i>Team Rank</i> - Based on Team Improvement for Scoring Members
					<br><br>
					<i>Team Improvement</i> - Calculated by summing the Member Improvement (for Scoring Members)
					<br><br>
					<i>Member Improvement</i> - Calculated as the difference between Current Ranking Score and Member Event Benchmark
					<br><br>
					<i>Member Event Benchmark</i> - Current formular uses 2014 Ski Year Ranking Score
					<br><br>
					<i># of Scoring Members</i> - As defined in Team Type Rules table (2 used in BOD & Sibling ranking)
				</td>
			</tr>
		</TABLE>		
		<%

END SUB



' --------------------------
  SUB Display_Team_Footer
' --------------------------

		%>
		<TABLE class="innertable" Align=center style="width:<%=tabwidth%>px"; height:"200px; ">
			<tr>
				<td align="left" style="width:100%; height:250px; vertical-align:top;">
				<b>TERMS & DEFINITIONS:</b>
					<br><br>
					<i>Rank Score</i> - Current 12 Month Rankings
					<br><br>
					<i>Rank Score Last SY Benchmark</i> - Ranking Score Last Ski Year [used as] Member Event Benchmark score
					<br><br>
					<i>Memb Improve</i> - Difference between Current Ranking Score and Member Event Benchmark Score
					<br><br>
					<i>In Team Member Rank</i> - Member ranking within Member's Team
					<br><br>
					<i>Rank Score 2SY Benchmark</i> - Ranking Score from 2nd Prior Ski Year (shown for comparison to Member Event Benchmark score)
				</td>
			</tr>
		</TABLE>		
		<%

END SUB




' ----------------------------------
  SUB CreatePageHead (PageHeadWidth)
' ----------------------------------

IF NOT(rs.EOF) THEN
		No_Scoring_Members=rs("No_Scoring_Members")
END IF

%>
<form action="/rankings/<%=ThisFileName%>" method="post">


<TABLE class="droptable" Align=center WIDTH=<%=PageHeadWidth%>px height=175 background="<%=MainImage%>">

  <% ' --- Total width 8 columns --- %>	
  <TR>
		<td colspan=6 align=left>
			<font color="<%=TextColor2%>" size="3">&nbsp;&nbsp;<b><%=PageTitle%></b></font>
			<br>
			<font color="<%=TextColor1%>" size="2">&nbsp;&nbsp;<b><%=PageSubTitle%></b></font>
		</td>
		<%

		IF AdminMenuLevel>=50 THEN  
				%>	
  			<td colspan=1 valign=top align="left">
					<font COlOR="<% =TitleColor %>" size="<% =fontsize2 %>"><b>Show SQL</b></font>
						<input type=checkbox name="sShowSQL" <% IF sShowSQL="on" THEN response.write "checked" %>>
				</td>
  			<td colspan=1 valign=top align="left">
					<font color="<% =TitleColor %>" size="<% =fontsize2 %>"><b>Stop</b></font>
						<input type=checkbox name="sStop" <% IF sStop="on" THEN response.write "checked" %>>
				</td>
				<%
		ELSE  
			%><td colspan=2 width="25%" valign=top align="left">&nbsp;</td><%
		
		END IF 


		%>
  	</TR>
  	<TR>
  	<%


	' --- Adds specific drop downs depending on PROCESS ---
	SELECT CASE process
	

		CASE "v_teammemberstatus", "v_teamstatus"
				%>
				<td align=right colspan=2 width="20%"><font size=1>&nbsp;&nbsp;Team Type: </font></td>
				<td align=left colspan=2 width="30%">
					<%
					' --- In Tools_
					BuildTeamType_DropDown 14,12,""  
					%>
				</td>
				<td colspan=4 width="50%">&nbsp;</td>
				<%

		CASE ELSE 
				%>
				<td colspan=2>&nbsp;</td>
				<%
		END SELECT 
		
		%>
  </TR>
  <TR>
		<td colspan=2 align=right><font size=1>&nbsp;&nbsp;# Scoring Members: </font></td> 
		<td colspan=2 align=left><font size=1>&nbsp;&nbsp;<%=No_Scoring_Members%></font></td> 
		</td>
		<td colspan=4>&nbsp;</td>
  </TR>

  <TR>
		<td colspan=2 width=25% align=center>
			<input type="submit" name="Action" style="width:9em" value="Update Display" title="Submit and reset this form">
		</td>

		<td colspan=2 width=25%  align=center>
			<input type="submit" name="Action" value="Return to Menu">
		</td>
		<td colspan=4>&nbsp;</td>
  </TR>	

</TABLE>

</form>

<%

END SUB







' ---------------------
  SUB Get_TeamStatus_RecordSet
' ---------------------

sSQL = "SELECT"
sSQL = sSQL + " RANK() OVER(Partition By Team_Type_ID ORDER BY SUM(MemberDeltaApplied) DESC) AS Rank"
sSQL = sSQL + ", Team_ID, Team_Name"

sSQL = sSQL + ", SUM(MemberDeltaApplied) AS Team_Improvement"
sSQL = sSQL + ", SUM(Rank_LSY_BM) AS Team_LSY_BM"
sSQL = sSQL + ", MAX(No_Scoring_Members) AS No_Scoring_Members"
sSQL = sSQL + ", SUM(BenchmarkApplied) AS BenchmarkApplied"
' sSQL = sSQL + ", ROUND(SUM(MemberDelta),2) AS [Team Improvement<br>All Scoring]"
sSQL = sSQL + ", COUNT(MemberID) AS No_Team_Members"


sSQL = sSQL + ", SUM(MemberRankApplied) AS Total_Score_Applied"
sSQL = sSQL + ", SUM(RankScore) AS Total_Score_AllMembers"

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
sSQL = sSQL + ", ttype.max_scoring AS No_Scoring_Members"

'sSQL = sSQL + ", CASE WHEN RANK() OVER(Partition By tmem.Team_ID ORDER BY RankScore)<=ttype.max_scoring THEN RankScore ELSE 0 END AS MemberRankApplied"
sSQL = sSQL + ", CASE WHEN RANK() OVER(Partition By tmem.Team_ID ORDER BY RankScore - Rank_LSY_BM DESC)<=ttype.max_scoring THEN RankScore ELSE 0 END AS MemberRankApplied"
sSQL = sSQL + ", CASE WHEN RANK() OVER(Partition By tmem.Team_ID ORDER BY RankScore - Rank_LSY_BM DESC)<=ttype.max_scoring THEN RankScore - Rank_LSY_BM ELSE 0 END AS MemberDeltaApplied"
'sSQL = sSQL + ", CASE WHEN RANK() OVER(Partition By tmem.Team_ID ORDER BY RankScore)<=ttype.max_scoring THEN Rank_LSY_BM ELSE 0 END AS BenchmarkApplied"
sSQL = sSQL + ", CASE WHEN RANK() OVER(Partition By tmem.Team_ID ORDER BY RankScore - Rank_LSY_BM DESC)<=ttype.max_scoring THEN Rank_LSY_BM ELSE 0 END AS BenchmarkApplied"
' sSQL = sSQL + ", CASE WHEN RANK() OVER(Partition By tmem.Team_ID ORDER BY RankScore - Rank_LSY_BM DESC)<=ttype.max_scoring THEN 1 ELSE 0 END AS NumMembersApplied"
	
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