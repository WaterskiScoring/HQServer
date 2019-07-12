<!--#include virtual="/rankings/settingsHQ.asp"-->
<%

Dim currentPage, rowCount, i
Dim TourDisplayWidth, ScorePageBorderDark, ScorePageBorderLight
Dim MainImage, tRCU, tFmtScore
Dim SkiYearSelected, PseudoSkiYear
Dim WorkStr, nTeams, TeamSQL 

Dim ThisFileName
ThisFileName="virtual-TeamStdgs.asp"


TourDisplayWidth=725
ScorePageBorderDark = HQSiteColor1
ScorePageBorderLight = HQSiteColor2


IF TRIM(Session("NewRankVis"))="" THEN
	KickTrafficCounter("NewRankVis")	
	Session("NewRankVis")="YES"
END IF

OpenCon
Set rs = Server.CreateObject("ADODB.recordset")

' --- Define Ski Year ---
SkiYearSelected = TRIM(Request("SkiYear"))
IF TRIM(SkiYearSelected) = "" AND TRIM(Session("SkiYear"))>"1" THEN SkiYearSelected=Session("SkiYear")

IF SkiYearSelected = "" THEN 
	sSQL = "SELECT top 1 * from " & SkiYearTableName & " WHERE DefaultYear = 1"
	rs.open sSQL, SConnectionToTRATable, 3, 3  
	IF NOT rs.EOF THEN
		SkiYearSelected = rs("SkiYearID")
		Session("SkiYear") = rs("SkiYearID")
	END IF
	rs.close
END IF	


WorkStr = Trim(request.form("Teams"))
IF WorkStr > "" THEN nTeams = 1 ELSE nTeams = 0
CommaLoc = instr(WorkStr,",")
if CommaLoc > 0 then
	TeamSQL = "'" & left(WorkStr, (CommaLoc - 1)) & "'"
else
	TeamSQL = "'" & WorkStr & "'"
end if

DO While instr(WorkStr,",") > 0 
	CommaLoc = instr(WorkStr,",")
	WorkStr = right(WorkStr, len(WorkStr) - (CommaLoc + 1))
	nTeams = nTeams + 1
	TeamSQL = TeamSQL & ",'" & left(WorkStr, (CommaLoc - 1)) & "'"
LOOP


' ------------------------------------------------------------
' If we have 2 or more teams, then create custom rankings,
' and then go to the display module citing this special entry.
' ------------------------------------------------------------

IF nTeams >= 2 THEN

'  WriteIndexPageHeader
  
'  		%><b><font color="red" size="2">
'		  <br>&nbsp;&nbsp;&nbsp;
'		  Teams = (<%=TeamSQL%>)  for SkiYear = <%=SkiYearSelected%></font></b><% 	

	' First Step is to come up with the Pseudo-SkiYearID for this particular
	' Visitor -- Pseudo-SkiYearID is 10000 * SkiYearSelected + increment.
	' Begin by getting highest existing increment for selected ski year.

	sSQL = "SELECT coalesce(Max(SkiYearID), 10000 * " & SkiYearSelected
	sSQL = sSQL & ") as MaxID from " & TmEvtScoTableName 
	sSQL = sSQL & " WHERE Cast(SkiYearID / 10000 as Smallint) = " & SkiYearSelected 

	PseudoSkiYear = 10000 * SkiYearSelected + 1  
	rs.open sSQL, SConnectionToTRATable, 3, 3
	IF NOT rs.EOF THEN
		PseudoSkiYear = rs("MaxID") + 1
	END IF
	rs.close

	' Next step is to insert into the Team/Event Scores table, to the
	' including the top 5 skiers for each Team, in each Division/Event.

	sSQL = "INSERT INTO " & TmEvtScoTableName & " (MemberID, Team,"
	sSQL = sSQL & " Event, Div, Score, TeamSeq, VirtTmStamp, SkiYearID)"
	sSQL = sSQL & " SELECT MemberID, Team, Event, Div, Score,"
	sSQL = sSQL & " TeamSeq, GetDate(), " & PseudoSkiYear
	sSQL = sSQL & " FROM " & TmEvtScoTableName 
	sSQL = sSQL & " WHERE SkiYearID = " & SkiYearSelected
	sSQL = sSQL & " AND Team in ( " & TeamSQL		
	sSQL = sSQL & " ) Order by Div, Event, Score"

	WriteDebugSQL(sSQL)

	Con.Execute(sSQL)


	'	Next we calculate NCWSA Placement Points for each Skier in each Div/Event 
	' placement set, according to latest NCWSA Rules.  This averages Min and Max 
	' PlacementSequence values, for each unique event (ranking) score value, 
	' to average across the tie groups, except zero where the raw score is zero.

	sSQL = "UPDATE EP SET PlcmtPts = Case when EP.Score <= 0 THEN 0" 
	sSQL = sSQL & " ELSE ((EMax.MaxSeq + EMin.MinSeq) / 2) - BMin.BaseSeq end"
	sSQL = sSQL & " FROM " & TmEvtScoTableName & " EP, (Select Div, Event,"
	sSQL = sSQL & " Score, Min(PlcmtSeq) as MinSeq FROM " & TmEvtScoTableName
	sSQL = sSQL & " WHERE SkiYearID = " & PseudoSkiYear
	sSQL = sSQL & " Group by Div, Event, Score) as EMin, (Select Div, Event,"
	sSQL = sSQL & " Score, Max(PlcmtSeq) as MaxSeq FROM " & TmEvtScoTableName
	sSQL = sSQL & " WHERE SkiYearID = " & PseudoSkiYear
	sSQL = sSQL & " Group by Div, Event, Score) as EMax, (Select Div, Event,"
	sSQL = sSQL & " Min(PlcmtSeq) - 10 as BaseSeq FROM " & TmEvtScoTableName
	sSQL = sSQL & " WHERE SkiYearID = " & PseudoSkiYear
	sSQL = sSQL & " Group by Div, Event) as BMin WHERE EP.Div = EMin.Div"
	sSQL = sSQL & " AND EP.Event = EMin.Event AND EP.Score = EMin.Score"
	sSQL = sSQL & " AND  EP.Div = EMax.Div AND EP.Event = EMax.Event"
	sSQL = sSQL & " AND EP.Score = EMax.Score AND  EP.Div = BMin.Div"
	sSQL = sSQL & " AND EP.Event = BMin.Event AND EP.SkiYearID = " & PseudoSkiYear

	' WriteDebugSQL(sSQL)

	Con.Execute(sSQL)


	' Next step is to Create Team Ranking Scores, by summing the Placement
	' Points of the best 4 skiers for each team, across each 2x3 Div/Event.
	' First we delete all the rows from the Team Ranking Table for this SkiYearID.

	sSQL = "Delete From " & TeamRankTableName & " where SkiYearID = " & PseudoSkiYear
	sSQL = sSQL & " ; INSERT INTO " & TeamRankTableName & " (Team, Div, Event, TeamScore,"
	sSQL = sSQL & " VirtTmStamp, SkiYearID) SELECT Team, Div, Event, Sum(PlcmtPts), "
	sSQL = sSQL & " max(VirtTmStamp), " & PseudoSkiYear
	sSQL = sSQL & " FROM " & TmEvtScoTableName & " WHERE TeamSeq <= 4"
	sSQL = sSQL & " AND SkiYearID = " & PseudoSkiYear
	sSQL = sSQL & " GROUP BY Team, Div, Event;"

	' WriteDebugSQL(sSQL)

	Con.Execute(sSQL)


	' Final Step is to roll up across Events and Divisions, creating both
	' Divisional Overall Totals, and Combined Team Event and Overall Totals.

	sSQL = "INSERT INTO " & TeamRankTableName & "(Team, Div, Event,"
	sSQL = sSQL & " TeamScore, VirtTmStamp, SkiYearID) SELECT Team, Div, 'O', Sum(TeamScore)," 
	sSQL = sSQL & " max(VirtTmStamp), " & PseudoSkiYear & " FROM " & TeamRankTableName 
	sSQL = sSQL & " WHERE SkiYearID = " & PseudoSkiYear & " GROUP BY Team, Div;"
	sSQL = sSQL & " INSERT INTO " & TeamRankTableName & "(Team, Div, Event,"
	sSQL = sSQL & " TeamScore, VirtTmStamp, SkiYearID) SELECT Team, 'CO', Event, Sum(TeamScore)," 
	sSQL = sSQL & " max(VirtTmStamp), " & PseudoSkiYear & " FROM " & TeamRankTableName
	sSQL = sSQL & " WHERE SkiYearID = " & PseudoSkiYear & " GROUP BY Team, Event;"

	' WriteDebugSQL(sSQL)

	Con.Execute(sSQL)

	response.redirect "view-TeamStdgsHQ.asp?SkiYear=" & PseudoSkiYear


ELSE

	' ---------------------------------------------
	' --- Writes header portion of HQ main page ---
	' ---------------------------------------------

	' --- Defines the image to be displayed in the drop downs box background ---
	' WhatDropdownImage EventSelected


	WriteIndexPageHeader


	' -------------------------------------------------------------------------------
	' -----  Check Recalculation Underway Flag for the Ski Year selected.  ----------
	' -----  If it's currently on, issue Come Back Later -- otherwise proceed.  -----
	' -------------------------------------------------------------------------------
	
	sSQL = "SELECT Case when RecalcUnderway=1 THEN 'Y' ELSE 'N' END as RCUFlag FROM " & SkiYearTableName & " WHERE SkiYearID = " & SkiYearSelected
	rs.open sSQL, SConnectionToTRATable, 3, 3  
	IF rs.EOF THEN tRCU = "N" ELSE tRCU = RS("RCUFlag")
	rs.close
	IF tRCU = "Y" and 1 = 2 THEN   ' --- Calc underway - Tell them to try again later
		%><b><font color="red" size="2">
		  <br>&nbsp;&nbsp;&nbsp;
		  Ranking Recalculations are currently underway For the Ski Year requested.&nbsp; Please try
		  <br>&nbsp;&nbsp;&nbsp;
		  your request again in a few minutes.&nbsp; We apologize for the temporary inconvenience.</font></b><% 	
	ELSE

		TitleColor=TextColor2 %>

		<TABLE width=90% align=center>

		<tr><td>&nbsp;</td></tr>

		<tr>
		  <td colspan=5 valign="top" align="left">
			<FONT size=4 face=<% =font1 %> Color=<% =textcolor2 %>><b>&nbsp;<I>NCWSA Custom Team Rankings</I></b></font>
		  </td>

		</tr>

		<tr><td>&nbsp;</td></tr>

		<% ' --------------------------------- Build SKI YEAR dropdown list  ------------------- %>

		<form method=post action="<%=ThisFileName%>">

		<tr>
		  <td width=9% align="center">
		    <font size="<%=fontsize2%>" face=<% =font1 %> color=<%=TitleColor%>><b>Period:</b></font>
		  </td>

		  <td>	
			<select OnChange=submit() name='SkiYear'><%

			sSQL = "SELECT DISTINCT RT.SkiYearID, SY.SkiYearName"
			sSQL = sSQL + " FROM " & TeamRankTableName & " AS RT"
			sSQL = sSQL + " JOIN " & SkiYearTableName & " AS SY ON RT.SkiYearID = SY.SkiYearID"
			sSQL = sSQL + " WHERE SY.SkiYearID <> 1"
		
			rs.open sSQL, SConnectionToTRATable

			' Loads dropdown and sets default to SkiYearSelected
			DO WHILE NOT rs.eof

				IF TRIM(rs("SkiYearID")) = TRIM(SkiYearSelected) THEN
					response.write("<option value =""" & rs("SkiYearID") &""" selected>")
					response.write(rs("SkiYearName"))
					response.write("</option><br>")
				ELSE
					response.write("<option value =""" & rs("SkiYearID") &""">")
					response.write(rs("SkiYearName"))
					response.write("</option><br>")
				END IF 

				rs.moveNEXT

			LOOP

			rs.close %>
			</select></td>

			<td width=3%>&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;</td> 

		</tr>

		<tr><td>&nbsp;</td></tr>

		</form>

		<form method=post action="<%=ThisFileName%>">
		<input type="hidden" name="SkiYear" value="<%=SkiYearSelected%>">
		
		<% ' ---- Build MULTIPLE Team Selection Dropdown ---- %>

		<tr>

			<td align="center"> 
				<font size=<% =fontsize2 %> face=<% =font1 %> color=<%=TitleColor%>><b><br>Teams:
				<br>&nbsp;<br>Listed<br>by<br>Region<br>and<br>Conf<br></b></font>
			</td>

			<td>	

				<select name="Teams" size="11" multiple id="Teams"><%
		
				sSQL = "SELECT TT.TeamID, TT.NCWRegion, TT.NCWConf, TT.TeamName"
				sSQL = sSQL + " FROM (SELECT Team FROM " & TeamRankTableName
				sSQL = sSQL + " WHERE SkiYearID = " & SkiYearSelected
				sSQL = sSQL + " GROUP BY Team) AS RT"
				sSQL = sSQL + " JOIN " & TeamTableName & " as TT on TT.TeamID = RT.Team"
				sSQL = sSQL + " ORDER BY TT.NCWRegion, TT.NCWConf, TT.TeamID"

				rs.open sSQL, SConnectionToTRATable

				DO WHILE NOT rs.eof

					response.write("<option value =""" & rs("TeamID") & """>")
					response.write(" " & rs("TeamName"))
					response.write("&nbsp; ( " & rs("NCWRegion"))
					IF rs("NCWConf") > " " THEN
						response.write(" / " & rs("NCWConf"))
					END IF	
					response.write(" )</option><br>")

					rs.moveNEXT

				LOOP

				rs.close %>
				</select></td>

			<td>&nbsp;</td>
	
			<td colspan=2><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
				<br>Indicate the set of two or more Teams which you wish to include
				in your Customized Virtual Team Competition.&nbsp; In order to 
				highlight multiple teams, <b><i>continuously hold down</i></b>
				the Ctrl Key, while you click on as many Teams as you desire.&nbsp;
				Then click the &#8220;Create Rankings&#8221; button, and I will 
				then create and score a Customized Virtual Tournament, based on the
				top five	ranked skiers in each event, for just those selected teams.
			</font></td>

		</tr>
	
		<tr><td>&nbsp;</td></tr>
	
		<tr>
	
			<td>&nbsp;</td>

		   <td align="center"><input type=submit style="width:11em" value="Create Rankings"
				title="Create and score a virtual competition &#13;for the specific Teams selected above"></td>

		</form>

		<form action="/rankings/news/FAQ_NCWRankings.htm" method="post" target="_blank">
	
			<td colspan=2 align=center><input type="submit" style="width:9em" value="FAQ/Tips"
				title="Tips and Explanations and Insights and &#13;Answers to Frequently Asked Questions">
			</td>
		</form>

		<form action="/rankings/view-TeamStdgsHQ.asp?SkiYear=<%=SkiYearSelected%>" method="post">
			<td align=center><input type="submit" style="width:11em" value="National Rankings"
				title="Take me back to the NCWSA &#13;National Team Rankings">
			</td>
		</form>

		</tr>

		</table>

		<%

	END IF

END IF

WriteIndexPageFooter

%>




