<!--#include file="settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->
<!--#include virtual="/rankings/tools_Definitions.asp"-->
<!--#include virtual="/rankings/tools_GRSite.asp"-->
<%


Dim EventSelected, BracketSelected, GenderSelected
Dim MainImage, sl, RecCnt, DefineLevelColor
Dim TourSelected, GROnly

Dim PageWidth
Dim SkiYearSelected, sSkiYear



ThisFileName="view-GRScores.asp"
PageWidth=725

DefineGRCSS

Set rs=Server.CreateObject("ADODB.recordset")

pvar="aws"

' --- IF GROnly=yes then tournament list includes only scores from ScoresGR table ---
GROnly="no"



SkiYearSelected=request("SkiYearSelected")
TourSelected=TRIM(request("TourSelected"))

EventSelected=TRIM(request("EventSelected"))
IF EventSelected="" THEN EventSelected="S"


'response.write("<br>EventSelected = "&EventSelected)
' --- Defines path to image and sets based on event ---
WhatDropdownImage (EventSelected)
'SetEventImage


'Session("SkiYearID")="14"

'response.write("<br>Ski Year Selected ="&SkiYearSelected)

IF SkiYearSelected<>"" THEN
	sSQL="SELECT SkiYearID, SkiYear FROM "&SkiYearTableName&" WHERE SkiYearID='"&CINT(SkiYearSelected)&"'" 
	rs.open sSQL, SConnectionToTRATable
	Session("SkiYearID")=rs("SkiYearID")
	sSkiYear=rs("SkiYear")
	rs.close
'response.write("<br> IN 1")

ELSEIF Session("SkiYearID")<>"" THEN
	sSQL="SELECT Top 1 SkiYearID, SkiYear FROM "&SkiYearTableName&" WHERE SkiYearID<>'1' ORDER BY SkiYearID DESC" 
	rs.open sSQL, SConnectionToTRATable
	Session("SkiYearID")=rs("SkiYearID")
	sSkiYear=rs("SkiYear")
	rs.close
'response.write("<br> IN 2")

END IF


'response.write("<br>SEES = "&Session("SkiYearID"))

WriteGRHeader

SELECT CASE pvar
	CASE "aws"
		PageTitle="Grassroots Competition Series"
		PageSubTitle= sSkiYear&" Ski Year"
		SQLGRScores
		IF rs.eof THEN 
			CreatePageHead PageWidth
			DisplayNoRecordsMessage
		ELSE	
			CreatePageHead PageWidth
			DisplayResult PageWidth
		END IF

END SELECT

WriteGRFooter


' ---------------------
  SUB SQLGRScores
' ---------------------


sSQL = " SELECT MT.FirstName AS [First<br>Name], MT.LastName AS [Last<br>Name], ST.MemberID, ST.TourID, ST.Event, ST.Div, Round, ROUND(ST.Score,2) AS Score"
sSQL = sSQL + " 	FROM "
sSQL = sSQL + " 		(SELECT MemberID, TourID, Event, Div, Round, Score, Class"
sSQL = sSQL + " 			FROM "&RawGRScoresTableName
sSQL = sSQL + " 				WHERE Class IN ('F', 'N') AND IsNull(score,1)<>1"

sSQL = sSQL + " 		UNION"

sSQL = sSQL + " 		SELECT MemberID, TourID, Event, Div, Round, Score, Class"
sSQL = sSQL + " 			FROM "&RawScoresTableName
sSQL = sSQL + " 				WHERE Class IN ('F', 'N') AND IsNull(score,1)<>1) AS ST"

sSQL = sSQL + " 		LEFT JOIN "
sSQL = sSQL + " 			( SELECT FirstName, LastName, PersonIDwithCheckDigit"
sSQL = sSQL + " 				FROM usawaterski.dbo.Members) AS MT"
sSQL = sSQL + " 		ON MT.PersonIDwithCheckDigit=ST.MemberID"			

sSQL = sSQL + "	LEFT JOIN"
sSQL = sSQL + "		(SELECT TournAppID, TDateE"
sSQL = sSQL + "			FROM "&SanctionTableName&") AS TS"
sSQL = sSQL + "	ON TS.TournAppID=LEFT(ST.TourID,6)"
	
sSQL = sSQL + "	LEFT JOIN "
sSQL = sSQL + "		(SELECT SkiYearID, EndDate, BeginDate FROM "&SkiYearTableName&") AS SY "
sSQL = sSQL + "	ON SY.SkiYearID='"&SkiYearSelected&"'"

sSQL = sSQL + "	WHERE SY.EndDate>=TS.TDateE AND SY.BeginDate<=TS.TDateE AND ST.TourID='"&TourSelected&"' AND ST.Event='"&EventSelected&"'"	

sSQL = sSQL + " 		ORDER BY ST.Event, ST.Div, ST.Round DESC"

'response.write(sSQL)
'response.end

rs.open sSQL, SConnectionToTRATable



END SUB


' ----------------------
  SUB CreatePageHead (PageHeadWidth)
' ----------------------

'class="grtable2" 

%>
<TABLE class="grtable2" Align=center cellpadding=0 WIDTH=<%=PageHeadWidth%>px height=200 background="<%=MainImage%>">
  <TR>
    <td colspan=6 >
	<font size=4 color=<%=textcolor3%> face=<%=font1%>><b><i><%=PageTitle%></i></b></font>
	<br>
	<font size=2 color=<%=textcolor2%> face=<%=font1%>><b><%=PageSubTitle%></b></font>
    </td>
  </TR>

 <form action="/rankings/<%=ThisFileName%>" method="post">
  <TR>
    <TD align=right width=70px><h5>Tour:</h5></td>
    <TD align=left width=225px><%
	' --- SUB in Tools_GRSite.asp ---
	LoadGRTournamentList GROnly
	%>
    </TD>
    <td align=left width=125>&nbsp;</td>	
    <td align=left width=125>&nbsp;</td>
    <td align=left width=60>&nbsp;</td>
    <td align=left width=60>&nbsp;</td>
  </TR>

  <TR>
    <td align=right valign=middle><p><h5>Event:</h5></td>
    <TD align=left><%
	LoadGREvents %>
    </td>
    <td align=left>&nbsp;</td>
    <td align=left>&nbsp;</td>
    <td align=left>&nbsp;</td>	
    <td align=left>&nbsp;</td>
  </TR>	

  <TR>
    <td align=right><h5>Year:</h5></td>
    <td align=left><%
		' --- SUB in Tools_Definitions.asp ---
		LoadSkiYearDropdown   %>
    </td>
    <td align=left>&nbsp;</td>
    <td align=left>&nbsp;</td>
    <td align=left>&nbsp;</td>
    <td align=left>&nbsp;</td>
  </TR>	

  <TR>
    <td align=left>&nbsp;</td>
    <td align=left>&nbsp;</td>
    <td align=center>
	<input type="submit" style="width:9em" value="Update Display" title="Submit and reset this form">
    </td>
   </form>
    <form action="/rankings/defaultHQ.asp" method="post">
    <td align=center>
	<input type="submit" style="width:9em" value="Main Menu">
    </td>
    <td align=left>&nbsp;</td>
    <td align=left>&nbsp;</td>

    </form>

	
  </TR>	


</TABLE>

<%



END SUB






' ----------------------------
  SUB DisplayNoRecordsMessage
' ----------------------------

%>
<TABLE class="grtable2" Align=center WIDTH=<%=PageWidth%>px height=100>
  <TR>
	<td align=center><h1>No Records Found</h1></td>
  </TR>
</TABLE><%


END SUB




' ---------------------
  SUB DisplayResult (tabwidth)
' ---------------------


	rs.movefirst

	RowCount=1
	LastLevelName=""

	' ---------------  Displays table HEADINGS  ----------------------

	%>
	<TABLE class="GRTable1" Align=center WIDTH=<%=tabwidth%>px>
	  <TR><%

		FOR i = 0 TO rs.fields.count - 1
			TempFN = rs.fields(i).name
			j = 0 %>

	   		<th ALIGN="Center" vAlign="top" nowrap>
			  <FONT COlOR="#FFFFFF" FACE="<%=font1%>" SIZE="<%=fontsize1%>"><b><%=Rs.Fields(i).name%></b></FONT>
			</th><%
		NEXT %>
	  </TR><%

	' --------------  Display table data here with paging --------------------------

	DO WHILE NOT rs.eof

		'IF rowCount = rs.PageSize THEN EXIT DO	%>

 		<TR><%


		FOR i = 0 TO rs.fields.count - 1
	
			RowColor=""	
			IF INT(RowCount/2)*2=RowCount THEN
				RowColor=GRTableColor4
			END IF

			%><TD ALIGN="center" style="background-color:<%=RowColor%>"><%

			    IF isnull(rs.Fields(i).value) THEN
				response.write ("&nbsp;")
			    ELSE %>
				<p><%=trim(Rs.Fields(i).Value)%></p><% 
			    END IF %> 
	
				</TD><%

		NEXT	%>

		</TR><% 
		rowCount = rowCount + 1
		rs.movenext
	LOOP %>

	</TABLE>
<br><br><%

END SUB



%>
