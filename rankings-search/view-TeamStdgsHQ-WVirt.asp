<!--#include virtual="/rankings/settingsHQ.asp"-->
<%

Dim currentPage, rowCount, i
Dim MemoryScore, MemoryPlc, MemoryRank, RecordNum, RankValueWithTies
Dim TourDisplayWidth, ScorePageBorderDark, ScorePageBorderLight
Dim MainImage, tRCU, tFmtScore
Dim SkiYearSelected, DivSelected, EventSelected, FilterSelected

Dim ThisFileName
ThisFileName="view-TeamStdgsHQ.asp"



'response.write("---")
TourDisplayWidth=725
ScorePageBorderDark = HQSiteColor1
ScorePageBorderLight = HQSiteColor2


IF TRIM(Session("NewRankVis"))="" THEN
	KickTrafficCounter("NewRankVis")	
	Session("NewRankVis")="YES"
END IF


' --- Define Ski Year ---
SkiYearSelected = TRIM(Request("SkiYear"))
IF TRIM(SkiYearSelected) = "" AND TRIM(Session("SkiYear"))>"1" THEN SkiYearSelected=Session("SkiYear")

' --- Define Event ---
EventSelected = TRIM(Request("event"))
IF EventSelected = "" THEN EventSelected = "O"

' --- Define Division ---
DivSelected = TRIM(Request("DivSelected"))
IF DivSelected = "" THEN DivSelected = "CO"

' --- Define Filtering ---
FilterSelected = TRIM(Request("filter"))
IF FilterSelected = "" THEN FilterSelected = "All"

' --- Defines the image to be displayed in the drop downs box background ---
WhatDropdownImage EventSelected



' ---------------------------------------------
' --- Writes header portion of HQ main page ---
' ---------------------------------------------

  WriteIndexPageHeader



    ' -----------------------------------------------------------------------------------------------------------
    ' ----------------   Sets Session("SkiYear") to request string from form   ------------------
    ' -----------------------------------------------------------------------------------------------------------

  IF SkiYearSelected = "" THEN 

    	OpenCon
      Set rs = Server.CreateObject("ADODB.recordset")
      sSQL = "SELECT top 1 * from " & SkiYearTableName & " WHERE DefaultYear = 1"
    	rs.open sSQL, SConnectionToTRATable, 3, 3  
      IF NOT rs.EOF THEN
      	 SkiYearSelected = rs("SkiYearID")
      	 Session("SkiYear") = rs("SkiYearID")
			END IF

  END IF	


  currentPage = TRIM(Request("currentPage"))
  IF currentPage = "" THEN currentPage = 1

  sID = TRIM(Request("id"))
  IF sID = "" THEN sID = 0
            
  ThisPage = Request.ServerVariables("SCRIPT_NAME")
            

	' ------------------------------------------------------------------------------------------------           
	' -------------------------------  BEGINS WRITING HEADERS AND RANKINGS  --------------------------
	' ------------------------------------------------------------------------------------------------


		' --- Displays picture box with drop downs ---

		DisplayDropDowns  

		' -------------------------------------------------------------------------------
		' -----  Check Recalculation Underway Flag for the Ski Year selected.  ----------
		' -----  If it's currently on, issue Come Back Later -- otherwise proceed.  -----
		' -------------------------------------------------------------------------------
		
		OpenCon
		Set rs = Server.CreateObject("ADODB.recordset")
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

		' -------------------------------------------------------------------------------
		' -----  Check for presence of a Selected Division Code.  If there has been 
		' -----  none specified yet, then ask user to Select one -- otherwise proceed.  -----
		' -------------------------------------------------------------------------------
	
		IF DivSelected = "" THEN   ' --- New Ranking Type -- Ask to select then hit Display
			%><b><font color="red" size="2">
			  <br>&nbsp;&nbsp;&nbsp;
			  Please Select a Division and Event using the drop-down boxes above,
			  <br>&nbsp;&nbsp;&nbsp;
			  then click the Display Rankings button to display that selection.</font></b><% 	
		ELSE

			 %>
			<TABLE width="<%=TourDisplayWidth%>px" style="border:1px solid <%=HQSiteColor2%>; white-space:nowrap" Align=center>
			<TR>
			  <TD style="white-space:nowrap">
				<div style="padding:5px; white-space:nowrap; overflow:auto; height:406px; "><%

				' --- Displays table header and ranking list ---

				DisplayRankList 

				KickTrafficCounter("NewRankPgs")

			    %></div>
		  	  </TD>
			</TR>

			</TABLE><br>&nbsp;<br><%

	  END IF

  END IF

' ---------------------------------------------
' --- Writes header portion of HQ main page ---
' ---------------------------------------------

  WriteIndexPageFooter


' ---------------------------------------------------------------------------------------------------------------
' ----------------------   END OF MAIN CODE HERE  ---------------------------------------------------------------
' ---------------------------------------------------------------------------------------------------------------




' -----------------------
   SUB DisplayRankList
' -----------------------

'	First Create and execute SQL query against Rankings Table for Selected Division/Event

OpenCon

SET rs=Server.CreateObject("ADODB.recordset")

sSQL = "Select RT.Team as TeamCode, TT.TeamName, RT.TeamScore,"
sSQL = sSQL & " TT.NCWRegion, TT.NCWConf FROM " & TeamRankTableName

IF FilterSelected = "All" THEN
	sSQL = sSQL & " as RT LEFT JOIN " & TeamTableName & " as TT on TT.TeamID = RT.Team"
ELSE
	sSQL = sSQL & " as RT JOIN " & TeamTableName & " as TT on TT.TeamID = RT.Team"
	IF Left(FilterSelected,1) = "1" THEN
		sSQL = sSQL & " AND TT.NCWRegion = '" & Mid(FilterSelected,2) & "'"
	ELSEIF Left(FilterSelected,1) = "2" THEN
		sSQL = sSQL & " AND TT.NCWConf = '" & Mid(FilterSelected,2) & "'"
	END IF
END IF

sSQL = sSQL & " WHERE RT.div = '" & DivSelected & "'"
sSQL = sSQL & " AND RT.event = '" & EventSelected & "'"
sSQL = sSQL & " AND RT.SkiYearID = " & SkiYearSelected


sSQL = sSQL & " ORDER BY TeamScore Desc"

' WriteDebugSQL(sSQL)

rs.CursorType = 3
rs.open sSQL, SConnectionToTRATable

rowCount = 0

' --------------- Now see if the record set is empty

IF rs.eof THEN

	%><TABLE class="innertable" width=95% align=center>
	  <TR><TD><br><br>
		<font color="red">No Rankings Found With These Filter Settings.</font>
    </TD></TR></TABLE><%

ELSE 


	' --- INITIALIZES the Ranking related memory fields for deal with ties.

	' --- RecordNum is essentially the record count
	' --- MemoryScore is the Score of the 
	' --- MemoryRank stores the highest value of placement - for which subsequent records may be tied 
	' --- rs("TeamScore") is the Score of the current record

	RecordNum = 1
	MemoryRank = 1
	MemoryScore = rs("TeamScore")
   
	' ---  After storing the values from the FIRST record then move to the second record to see if tied to know
	' ---     whether the FIRST record should have a T after it.  All others

	' --- Move to 2nd record ---
	rs.MoveNEXT

	' --- If the score from last tied record is same as current score 
	RankValueWithTies = "1"
	IF NOT rs.EOF THEN
		IF MemoryScore = rs("TeamScore") THEN
			RankValueWithTies = "1T"
		END IF
	END IF

	' --- Now move back to FIRST record and initialize First record in query ---

	rs.MoveFIRST

	' --- Set up table headings	

	%><TABLE class="innertable" width=95%><%

	' --- Displays the header on the top of the table ---

	HeadColor1="#FFFFFF"
	
	SELECT CASE DivSelected
	CASE "CO"
		tDivEvent = "Combined "
	CASE "CM"
		tDivEvent = "Men "
	CASE "CW"
		tDivEvent = "Women "
	CASE ELSE
		tDivEvent = "Unknown "
	END SELECT

	SELECT CASE EventSelected
	CASE "S"
		tDivEvent = tDivEvent & "Slalom"
	CASE "T"
		tDivEvent = tDivEvent & "Trick"
	CASE "J"
		tDivEvent = tDivEvent & "Jump"
	CASE "O"
		tDivEvent = tDivEvent & "Overall"
	CASE ELSE
		tDivEvent = tDivEvent & "Unknown"
	END SELECT


	' ---------------  Display Column Headings for TEAM Rankings  ------------------

		%>

		<TR><TH Colspan=6 ALIGN="Center"><font size=<%=fontsize1%>>&nbsp;<br>
		<font size=<%=fontsize4%> COLOR="#FFFFFF">
		<b>NCWSA Team Rankings for Division / Event:&nbsp;&nbsp; 
		<%=tDivEvent%></b></font><font size=<%=fontsize1%>><br>&nbsp;</font>
		</TH></TR>
		
		<TR>
		<Th ALIGN="Center" Width=9% ><font size=<%=fontsize2%> COLOR="#FFFFFF"><b>Rank</b></FONT></th>
		<Th ALIGN="Center" ><font size=<%=fontsize2%> COLOR="#FFFFFF"><b>Team Name</b></FONT></th>
		<Th ALIGN="Center" ><font size=<%=fontsize2%> COLOR="#FFFFFF"><b>Tm Code</b></FONT></th>
		<Th ALIGN="Center" ><font size=<%=fontsize2%> COLOR="#FFFFFF"><b>Regn</b></FONT></th>
		<Th ALIGN="Center" ><font size=<%=fontsize2%> COLOR="#FFFFFF"><b>Conf</b></FONT></th>
		<Th ALIGN="Center" ><font size=<%=fontsize2%> COLOR="#FFFFFF"><b>Team Score</b></FONT></th>
		</TR><%

	' --- Now loop over complete recordset and display all rows returned.

	DO WHILE NOT rs.eof

		' --- Displays one line of TEAM Ranking Detail ---

		%><TR>

			  <TD ALIGN="Center" vAlign="top">
				<font size=<% =fontsize2 %> COLOR="<%=TextColor1%>"><%=RankValueWithTies%></font>
			  </TD>

			  <TD ALIGN="Center" vAlign="top">
			  <font size=<% =fontsize2 %> COLOR="<%=TextColor1%>"><%=rs("TeamName")%></a>
			  </TD>

			  <TD ALIGN="Center" vAlign="top">
				<font size=<% =fontsize2 %> COLOR="<%=TextColor1%>"><%=rs("TeamCode")%></FONT>
			  </TD>

			  <TD ALIGN="Center" vAlign="top">
				<font size=<% =fontsize2 %> COLOR="<%=TextColor1%>"><%=rs("NCWRegion")%></FONT>
			  </TD>

			  <TD ALIGN="Center" vAlign="top">
				<font size=<% =fontsize2 %> COLOR="<%=TextColor1%>"><%=rs("NCWConf")%></FONT>
			  </TD>

			  <TD ALIGN="Center" vAlign="top">
			  <font size=<% =fontsize2 %> COLOR="<%=TextColor1%>"><%=rs("TeamScore")%></a>
			  </TD>

			</TR><%

		' --- Initializes NEXT record in query --- 
		rs.moveNEXT
		RecordNum = RecordNum + 1

		IF NOT rs.eof THEN
			

			' --- If the score from PREVIOUS record is same as current score 
			IF MemoryScore = rs("TeamScore") THEN
				RankValueWithTies = MemoryRank&"T"
			ELSE

				MemoryRank = RecordNum
				MemoryScore = rs("TeamScore")
				
				' --- Move to NEXT record to see if tied---
				rs.MoveNEXT
			    	IF NOT rs.eof THEN

					' --- If the score from last tied record is same as current score 
					IF MemoryScore = rs("TeamScore") THEN
						RankValueWithTies = RecordNum&"T"
					ELSE
						RankValueWithTies = RecordNum
					END IF

				ELSE
					' --- Can't be tied with EOF so set it to the current record ---
					RankValueWithTies = RecordNum
				END IF

				' --- Now move back to CURRENT record and initialize ---
				rs.MovePREVIOUS

			END IF

		ELSE

		END IF

	LOOP  

	rs.Close
	%>
    </TABLE><%

	

	' -------------- Now display underlying INDIVIDUAL skier detail, where they exist

	IF DivSelected <> "CO" AND EventSelected <> "O" THEN

	sSQL = "Select ET.MemberID, Coalesce(MT.LastName,'?') + ', ' + Coalesce(MT.FirstName,'?') as SkierName,"
	sSQL = sSQL & " ET.Score as RankScore, ET.Team as TeamCode, ET.PlcmtPts as TeamPts,"
	sSQL = sSQL & " RT.RnkScoBkup, TT.NCWRegion, TT.NCWConf"
	sSQL = sSQL & " FROM " & TmEvtScoTableName & " as ET LEFT JOIN " & MemberTableName 
	sSQL = sSQL & " as MT on ET.MemberID = MT.PersonIDWithCheckDigit"
	sSQL = sSQL & " LEFT JOIN " & RankTableName & " as RT on RT.MemberID = ET.MemberID"
	sSQL = sSQL & " AND RT.Div = ET.Div and RT.Event = ET.Event and RT.SkiYearID ="

	IF SkiYearSelected > 10000 THEN
		sSQL = sSQL & " cast(" & SkiYearSelected & "/10000 as smallint)"
	ELSE
		sSQL = sSQL & SkiYearSelected
	END IF
	
	IF FilterSelected = "All" THEN
		sSQL = sSQL & " LEFT JOIN " & TeamTableName & " as TT on TT.TeamID = ET.Team"
	ELSE
		sSQL = sSQL & " JOIN " & TeamTableName & " as TT on TT.TeamID = ET.Team"
		IF Left(FilterSelected,1) = "1" THEN
			sSQL = sSQL & " AND TT.NCWRegion = '" & Mid(FilterSelected,2) & "'"
		ELSEIF Left(FilterSelected,1) = "2" THEN
			sSQL = sSQL & " AND TT.NCWConf = '" & Mid(FilterSelected,2) & "'"
		END IF
	END IF

	sSQL = sSQL & " WHERE ET.Div = '" & DivSelected & "'"
	sSQL = sSQL & " AND ET.Event = '" & EventSelected & "'"
	sSQL = sSQL & " AND ET.SkiYearID = " & SkiYearSelected
	sSQL = sSQL & " ORDER BY ET.PlcmtPts Desc"

	' WriteDebugSQL(sSQL)

	rs.CursorType = 3
	rs.open sSQL, SConnectionToTRATable

	rowCount = 0


	' --- INITIALIZES the Ranking related memory fields for deal with ties.

	' --- RecordNum is essentially the record count
	' --- MemoryScore is the Score of the 
	' --- MemoryRank stores the highest value of placement - for which subsequent records may be tied 
	' --- rs("RankScore") is the Score of the current record

	RecordNum = 1
	MemoryRank = 1
	MemoryScore = rs("RankScore")
   
	' ---  After storing the values from the FIRST record then move to the second record to see if tied to know
	' ---     whether the FIRST record should have a T after it.  All others

	' --- Move to 2nd record ---
	rs.MoveNEXT

	' --- If the score from last tied record is same as current score 
	IF MemoryScore = rs("RankScore") THEN
		RankValueWithTies = "1T"
	ELSE
		RankValueWithTies = "1"
	END IF

	' --- Now move back to FIRST record and initialize First record in query ---

	rs.MoveFIRST

	
	' --- Set up Table with Headings and so forth

	%><br>&nbsp;<br>
	
	<TABLE class="innertable" width=95%>

	
	<%

	' --- Displays the header on the top of the table ---

	HeadColor1="#FFFFFF"

	' ---------------  Display Column Headings for INDIVIDUAL Rankings  ------------------

		%>
		
		<TR><TH Colspan=7 ALIGN="Center"><font size=<%=fontsize1%>>&nbsp;<br>
		<font size=<%=fontsize4%> COLOR="#FFFFFF">
		<b>Top 5 Contributors to Teams -- Ranking Scores and Team
		Pts</b></font><font size=<%=fontsize1%>><br>&nbsp;</font>
		</TH></TR>

		<TR>
		<Th ALIGN="Center" Width=9% ><font size=<%=fontsize2%> COLOR="#FFFFFF"><b>Rank</b></FONT></th>
		<Th ALIGN="Center" ><font size=<%=fontsize2%> COLOR="#FFFFFF"><b>Member Name</b></FONT></th>
		<Th ALIGN="Center" ><font size=<%=fontsize2%> COLOR="#FFFFFF"><b>Ranking Score</b></FONT></th>
		<Th ALIGN="Center" ><font size=<%=fontsize2%> COLOR="#FFFFFF"><b>Tm Code</b></FONT></th>
		<Th ALIGN="Center" ><font size=<%=fontsize2%> COLOR="#FFFFFF"><b>Regn</b></FONT></th>
		<Th ALIGN="Center" ><font size=<%=fontsize2%> COLOR="#FFFFFF"><b>Conf</b></FONT></th>
		<Th ALIGN="Center" ><font size=<%=fontsize2%> COLOR="#FFFFFF"><b>Team Pts</b></FONT></th>
		</TR><%

	' --- Now loop over complete recordset and display all rows returned.

	DO WHILE NOT rs.eof

		IF EventSelected = "S" THEN
			tFmtScore = FormatNumber(rs("RankScore"),2)
		ELSE
			IF EventSelected = "J" THEN
				tFmtScore = FormatNumber(rs("RankScore"),1)
			ELSE
				tFmtScore = FormatNumber(rs("RankScore"),0)
			END IF
		END IF
		
	' --- Displays one line of INDIVIDUAL Ranking Detail ---

		%><TR>

			  <TD ALIGN="Center" vAlign="top">
				<font size=<% =fontsize2 %> COLOR="<%=TextColor1%>"><%=RankValueWithTies%></FONT>
			  </TD>

			  <TD ALIGN="Center" vAlign="top">
			  <a href="/rankings/view-scoresHQ.asp?NSL=&sMemberID=<%=rs("MemberID")%>&EventSelected=<%=EventSelected%>&SkiYear=<%=SkiYearSelected%>&pvar=ByMember"
			  title="Click here to Display ALL of&#13;<%=mid(rs("SkierName"),instr(rs("SkierName"),", ")+2)%>'s scores in this Event"><font size=<%=fontsize2%> 
			  COLOR="<%=TextColor1%>"><%=rs("SkierName")%></a></FONT>
			  </TD>

			  <TD ALIGN="Center" vAlign="top">
			  <font size=<% =fontsize2 %> COLOR="<%=TextColor1%>"><a title="<%=rs("RnkScoBkup")%>"><%=tFmtScore%></a></FONT>
				   <% IF instr(rs("RnkScoBkup"),"Rule 1.13")<>0 THEN %>
						<font color="red"> #</font>
					<% ELSEIF instr(rs("RnkScoBkup"),"NO Penalty")=0 THEN %>
						<font color="red"> *</font>
					<% END IF %>
			  </TD>

			  <TD ALIGN="Center" vAlign="top">
				<font size=<% =fontsize2 %> COLOR="<%=TextColor1%>"><%=rs("TeamCode")%> (A)</FONT>
			  </TD>

			  <TD ALIGN="Center" vAlign="top">
				<font size=<% =fontsize2 %> COLOR="<%=TextColor1%>"><%=rs("NCWRegion")%></FONT>
			  </TD>

			  <TD ALIGN="Center" vAlign="top">
				<font size=<% =fontsize2 %> COLOR="<%=TextColor1%>"><%=rs("NCWConf")%></FONT>
			  </TD>

			  <TD ALIGN="Center" vAlign="top">
			  <font size=<% =fontsize2 %> COLOR="<%=TextColor1%>"><%=rs("TeamPts")%></FONT>
			  </TD>

			</TR><%

		' --- Initializes NEXT record in query --- 
		rs.moveNEXT
		RecordNum = RecordNum + 1

		IF NOT rs.eof THEN
			
			' --- If the score from PREVIOUS record is same as current score 
			IF MemoryScore = rs("RankScore") THEN
				RankValueWithTies = MemoryRank&"T"
			ELSE

				MemoryRank = RecordNum
				MemoryScore = rs("RankScore")
				
				' --- Move to NEXT record to see if tied---
				rs.MoveNEXT
			    	IF NOT rs.eof THEN

					' --- If the score from last tied record is same as current score 
					IF MemoryScore = rs("RankScore") THEN
						RankValueWithTies = RecordNum&"T"
					ELSE
						RankValueWithTies = RecordNum
					END IF

				ELSE
					' --- Can't be tied with EOF so set it to the current record ---
					RankValueWithTies = RecordNum
				END IF

				' --- Now move back to CURRENT record and initialize ---
				rs.MovePREVIOUS

			END IF

		ELSE
		
		END IF

		LOOP  

		rs.Close
		%>

		<% IF SkiYearSelected < 10000 THEN %>
		
			<form action="/rankings/view-StandingsHQ.asp?NSL=&pvar=NCWSA&Event=<%=EventSelected%>&DivSelected=<%=DivSelected%>&SkiYear=<%=SkiYearSelected%>" method="post">
			<td colspan=7 align=center>&nbsp;<br><input type="submit" style="width:16em" value="Complete Individual Rankings"
			title="Take me to the Complete NCWSA&#13;INDIVIDUAL <%=tDivEvent%> Rankings"><br>&nbsp;
			</td></form>		

		<% END IF %>
		
		</TABLE><%

	END IF


END IF

CloseCon

END SUB





' --------------------------------------------------------------------------------------------
   SUB DisplayDropdowns      ' -----   Begin form for selection /  filtering parameters ------
' --------------------------------------------------------------------------------------------

TitleColor=TextColor2

%>

<TABLE class="droptable" background="<%=MainImage%>" align=center width="<%=TourDisplayWidth%>" height=150><% '---Table to hold image --- %>
  <TR>
    <TD >
    <% ' ------ OUTER TABLE TO HOLD BACKGROUND IMAGE %>

<TABLE width=100% align=center>

<tr><td>&nbsp;</td></tr>

<tr>
  <td colspan=5 valign="top" align="left">
		<FONT size=4 face=<% =font1 %> Color=<% =textcolor2 %>><b><I>NCWSA
		<% IF IsNumeric(SkiYearSelected) AND SkiYearSelected > 10000 THEN Response.write("Custom") ELSE Response.write("National") %>
		Team Rankings</I></b></font>
  </td>

</tr>

<tr><td>&nbsp;</td></tr>

<form method=post action="<%=ThisFileName%>">

<% ' --------------------------------- Build SKI YEAR dropdown list  ------------------- %>

<tr>
  <td width=7% align="center">
    <font size="<%=fontsize2%>" face=<% =font1 %> color=<%=TitleColor%>><b>Period:</b></font>
  </td>

  <td colspan=2>	

	<% IF SkiYearSelected > 10000 THEN

		SET rsSelectFields=Server.CreateObject("ADODB.recordset")
		sSQL = "SELECT SkiYearName FROM " & SkiYearTableName
		sSQL = sSQL & " WHERE SkiYearID = " & FormatNumber(SkiYearSelected/10000,0)

		rsSelectFields.open sSQL, SConnectionToTRATable %>

		<font size=<%=fontsize2%> COLOR="<%=TextColor1%>"><b><%=rsSelectFields("SkiYearName")%></b></FONT>
		<input type="hidden" name="SkiYear" value="<%=SkiYearSelected%>">
		
	<% ELSE %>

		<select name='SkiYear'><%

		SET rsSelectFields=Server.CreateObject("ADODB.recordset")
		sSQL = "SELECT DISTINCT RT.SkiYearID, SY.SkiYearName"
		sSQL = sSQL & " FROM " & TeamRankTableName & " AS RT"
		sSQL = sSQL & " JOIN " & SkiYearTableName & " AS SY ON RT.SkiYearID = SY.SkiYearID"
		sSQL = sSQL & " WHERE SY.SkiYearID <> 1"

		rsSelectFields.open sSQL, SConnectionToTRATable

		' Loads dropdown and sets default to SkiYearSelected
		DO WHILE NOT rsSelectFields.eof	

			IF TRIM(rsSelectFields("SkiYearID")) = TRIM(SkiYearSelected) THEN
				response.write("<option value =""" & rsSelectFields("SkiYearID") &""" selected>")
				response.write(rsSelectFields("SkiYearName"))
				response.write("</option><br>")
			ELSE
				response.write("<option value =""" & rsSelectFields("SkiYearID") &""">")
				response.write(rsSelectFields("SkiYearName"))
				response.write("</option><br>")
			END IF 

			rsSelectFields.moveNEXT

		LOOP

		rsSelectFields.close %>
		</select>

	<% END IF %>

  </td>

		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>

	</tr>

<tr><td>&nbsp;</td></tr>



<% ' ---- Build Division Dropdown ---- %>

<tr>

  <td align="center"> 
     <font size=<% =fontsize2 %> face=<% =font1 %> color=<%=TitleColor%>><b>Division:</b></font>
  </td>

  <td colspan=2 width=25%>	
	<select name='DivSelected'>
	  <option value ='CO' <%IF DivSelected="CO" THEN response.write(" selected")%>>Combined Team</Option><br>
	  <option value ='CM' <%IF DivSelected="CM" THEN response.write(" selected")%>>Men Team</Option><br>
	  <option value ='CW' <%IF DivSelected="CW" THEN response.write(" selected")%>>Women Team</Option><br>
	</select>
  </td>

  <td align=left colspan=2>&nbsp;</td>
  <td>&nbsp;</td>
  
</tr>
  
<tr><td>&nbsp;</td></tr>



<% ' ---- Build Event Dropdown ---- %>

<tr>

  <td align="center"> 
     <font size=<% =fontsize2 %> face=<% =font1 %> color=<%=TitleColor%>><b>Event:</b></font>
  </td>

  <td colspan=2 width=25%>	
	<select name='event'>
		<option value ='S' <%IF EventSelected="S" THEN response.write(" selected")%>>Slalom</Option><br>
		<option value ='T' <%IF EventSelected="T" THEN response.write(" selected")%>>Trick</Option><br>
		<option value ='J' <%IF EventSelected="J" THEN response.write(" selected")%>>Jump</Option><br>
		<option value ='O' <%IF EventSelected="O" THEN Response.Write(" selected")%>>Overall</option><br>
	</select>
  </td>

  <td align=left colspan=2>&nbsp;</td>
  <td>&nbsp;</td>
  
</tr>
  
<tr><td>&nbsp;</td></tr>


<% ' ---- Build Filter Dropdown ---- %>

<tr>

	<% IF SkiYearSelected > 10000 THEN %>

		<td align="center"> 
			<font size=<%=fontsize2%> face=<%=font1%> color=<%=TitleColor%>><b>Teams:</b></font>
		</td>

		<td colspan=2 width=25%>	
			<font size=<%=fontsize2%> COLOR="<%=TextColor1%>"><b>Custom Team Selection</b></FONT>
		<input type="hidden" name="filter" value="All">
		</td>

	<% ELSE %>	

		<td align="center"> 
			<font size=<%=fontsize2%> face=<%=font1%> color=<%=TitleColor%>><b>Region or<br>Conference:</b></font>
		</td>

	  <td colspan=2 width=25%>	

		<select name='filter'>
		  <option value ='All'  <%IF FilterSelected="All"  THEN response.write(" selected")%>>All</Option><br>
		  <option value ='1E'  <%IF FilterSelected="1E"  THEN response.write(" selected")%>>Eastern Region</Option><br>
		  <option value ='2NE' <%IF FilterSelected="2NE" THEN response.write(" selected")%>>.. Northeast Conf</Option><br>
		  <option value ='2SA' <%IF FilterSelected="2SA" THEN response.write(" selected")%>>.. So Atlantic Conf</Option><br>
		  <option value ='2SO' <%IF FilterSelected="2SO" THEN response.write(" selected")%>>.. Southern Conf</Option><br>
		  <option value ='1M'  <%IF FilterSelected="1M"  THEN response.write(" selected")%>>Midwest Region</Option><br>
		  <option value ='2GL' <%IF FilterSelected="2GL" THEN response.write(" selected")%>>.. Great Lakes Conf</Option><br>
		  <option value ='2GP' <%IF FilterSelected="2GP" THEN response.write(" selected")%>>.. Great Plains Conf</Option><br>
		  <option value ='1SC' <%IF FilterSelected="1SC" THEN Response.Write(" selected")%>>South Central Region</option><br>
		  <option value ='1W'  <%IF FilterSelected="1W"  THEN response.write(" selected")%>>Western Region</Option><br>
		  <option value ='2NW' <%IF FilterSelected="2NW" THEN response.write(" selected")%>>.. Northwest Conf</Option><br>
		  <option value ='2PC' <%IF FilterSelected="2PC" THEN response.write(" selected")%>>.. Pacific Conf</Option><br>
		</select>

	  </td>

	<% END IF %>
	
<% IF DivSelected = "" THEN %>
   <td align="center"><input type=submit style="width:9em" value="Display Rankings"
   	  title="Display Rankings for the selection parameters above/left"></td>
<% ELSE %>
   <td align="center"><input type=submit style="width:9em" value="Update Display"
   	  title="Display a revised Rankings page after you have &#13;changed the selection parameters above/left"></td>
<% END IF %>

</form>


<form action="/rankings/news/FAQ_NCWRankings.htm" method="post" target="_blank">
	<td align=center><input type="submit" style="width:8em" value="FAQ/Tips"
	title="Tips and Explanations and Insights and &#13;Answers to Frequently Asked Questions">
	</td>
</form>

<% IF SkiYearSelected > 10000 THEN %>

	<form action="/rankings/view-TeamStdgsHQ.asp?SkiYear=<%=FormatNumber(SkiYearSelected/10000,0)%>" method="post">
		<td align=center><input type="submit" style="width:10em" value="National Rankings"
			title="Take me back to the NCWSA &#13;National Team Rankings">
		</td>
	</form>

<% ELSE %>

	<form action="/rankings/virtual-TeamStdgs.asp?SkiYear=<%=SkiYearSelected%>" method="post">
		<td align=center><input type="submit" style="width:9em" value="Custom Rankings"
			title="Build Custom Team Rankings for &#13;a selected set of NCWSA Teams">
		</td>
	</form>

<% END IF %>


</tr>

</table>

</TD>
</TR>
</TABLE><% ' --- Table to hold picture ---


END SUB


%>




