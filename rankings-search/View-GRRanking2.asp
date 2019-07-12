<%

IF TRIM(request("ws"))="" THEN 
'response.write("YES")
'response.end
	response.write("<!--#include file='settingsHQ.asp'-->")
	response.write("<!--#include virtual='/grassrootsseries/tools_include.asp'-->")
	response.write("<!--#include virtual='/grassrootsseries/tools_definitions.asp'-->")
	response.write("<!--#include virtual='/grassrootsseries/tools_GRSite.asp'-->")
ELSE
'response.write("NO")
'response.end

END IF



ThisFileName="view-grRanking2.asp"


Dim EventSelected, BracketSelected, GenderSelected, SkiYearSelected, sSkiYearID, AdminMenuLevel
Dim MainImage, sl, RecCnt, DefineLevelColor, sShowSQL

Dim BonMult_S, BonMult_T, BonMult_J, BonMult_WB, BonMult_WS, BonMult_WU, BonMult_KB, BonMult_KF, BonMult_KP
Dim Level1Name, Level2Name, Level3Name
Dim PageTitle, PageSubTitle
Dim WhatHeadFoot


BonMult_S=5
BonMult_T=100
BonMult_J=10

Level3Name="Outlaw"
Level2Name="Competitor"
Level1Name="Challenger"

Level3Color=sColor07
Level2Color=sColor08
Level1Color=sColor06


GRTableColor1="#000000"
GRTableColor2="#303030"
GRTableColor3="#707070"

TourTableWidth=675

sl="on"
SetEventImage



AdminMenuLevel=Session("AdminMenuLevel")

'PathtoGR = Server.mappath("/")
'response.write("<br>PathtoGR="&PathtoGR)



' --------------------------
' --- Get form variables ---
' --------------------------
'process=TRIM(LCASE(request("process")))
'process="grrank"

EventSelected=TRIM(Request("EventSelected"))
BracketSelected=TRIM(Request("BracketSelected"))
IF BracketSelected="" THEN BracketSelected="0"
GenderSelected=TRIM(Request("GenderSelected"))

'sDiv=TRIM(Request("sGrpDiv"))
sShowSQL=Request("sShowSQL")


' -----------------------------------
' --- Set SkiYear and SkiYearName ---
' -----------------------------------
sSkiYearID=TRIM(Request("sSkiYearID"))
'SkiYearSelected=TRIM(Request("SkiYearSelected"))

'response.write("<br>sSkiYearID="&sSkiYearID)

session("SkiYear")=sSkiYear

SET rsSY=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT SkiYear, SkiYearName, SkiYearID"
sSQL = sSQL + " FROM " &SkiYearTableName&" AS SY"
IF sSkiYearID<>"" THEN 
	sSQL = sSQL + "    WHERE SkiYearID='"&sSkiYearID&"'"
ELSE
	sSQL = sSQL + "    WHERE SkiYearID='1'"
END IF 
rsSY.open sSQL, SConnectionToTRATable

IF NOT rsSY.eof THEN 
	sSkiYearName=rsSY("SkiYearName")
	sSkiYear=rsSY("SkiYear")
	sSkiYearID=rsSY("SkiYearID")
END IF

PageTitle="Grassroots Rankings"
PageSubTitle=sSkiYearName&" Ranking List"


'response.write("<br>rsSY.eof="&rsSY.eof)
'response.end



Set rs=Server.CreateObject("ADODB.recordset")


' ---------------------------------------------------------
' --- Define whick background footer to use for display ---
' ---------------------------------------------------------
WhatHeadFoot=Request("WhatHeadFoot")
' --- rs (rankings site) allows display to be activated through rankings site ---
'WhatHeadFoot="rs"
'response.write("<br>WhatHeadFoot="&WhatHeadFoot)

IF TRIM(LCASE(WhatHeadFoot))="rs" THEN 
	' --- In Tools_Include.asp
	DefineTRAStyles
	WriteIndexPageHeader
ELSE
	' --- In Tools_GRSite.asp
	WriteGRHeader
END IF



' --------------------------------------------------
' --- Buildes box with drop downs ---
' --------------------------------------------------

DefineGRCSS
CreatePageHead

IF EventSelected<>"" THEN
	GRRanking
	DisplayResult
ELSE
	response.write("<br>In No Rec")
	NoRecordMessage
END IF



' ---------------------------------------------------------
' --- Define whick background footer to use for display ---
' ---------------------------------------------------------

IF TRIM(LCASE(WhatHeadFoot))="rs" THEN 
	WriteIndexPageFooter
ELSE
	WriteGRFooter
END IF



' *******************************************************************************************************
' --- End of MAIN PROGRAM 
' *******************************************************************************************************






' ---------------------
  SUB DisplayResult
' ---------------------

Dim RowCount, Rank

	SELECT CASE RIGHT(BracketSelected,1)
		CASE "3"
			StartRec=1
			Mid13Rec=INT(RecCnt/3)
			Mid23Rec=INT(RecCnt/3)
			EndRec=INT(RecCnt/3)
		CASE "2"
			StartRec=INT(RecCnt/3)+1
			Mid13Rec=INT(RecCnt/3)+1
			Mid23Rec=INT(2*RecCnt/3)
			EndRec=INT(2*RecCnt/3)
		CASE "1"  ' --- Level1Name ---
			StartRec=INT(2*RecCnt/3)+1
			Mid13Rec=INT(RecCnt/3)+1
			Mid23Rec=INT(2*RecCnt/3)+1
			EndRec=RecCnt
		CASE ELSE 
			StartRec=1
			Mid13Rec=INT(RecCnt/3)
			Mid23Rec=INT(2*RecCnt/3)
			EndRec=RecCnt
	END SELECT	


	rs.movefirst

	' ---------------  Displays table HEADINGS  ----------------------

	%>

	<TABLE class="GRtable1" Align=center WIDTH=725px >
	  <TR>
  		<th ALIGN="Center" >
		  <FONT COlOR="#FFFFFF" FACE="<%=font1%>" SIZE="<%=fontsize1%>">Rank</FONT>
		</th><%


		FOR i = 0 TO rs.fields.count - 1
			TempFN = rs.fields(i).name
			j = 0 %>

	   		<th ALIGN="Center" vAlign="top" nowrap>
			  <FONT COlOR="#FFFFFF" FACE="<%=font1%>" SIZE="<%=fontsize1%>"><%=Rs.Fields(i).name%></FONT>
			</th><%
		NEXT %>
	  </TR><%

	' --------------  Display table data here with paging --------------------------

	' --- Starts the display at the beginning of that bracket ---
	RowCount=1
	Rank=1
	LastLevelName=""

	TypeColor5="#000000" 	' --- Black
	TypeColor6="#FFFFFF"	' --- White 


	DO WHILE NOT rs.eof

		' --- If not in range of BRACKET then skip to next record ---
		IF RowCount>=StartRec THEN

			' --- Displays colored background in RANK cell ---
			DisplayBar="no"
			IF RowCount>=1 AND RowCount<Mid13Rec THEN
				ThisLevelColor=Level3Color
				ThisLevelName=Level3Name
				DisplayBar="yes"
			ELSEIF RowCount>=Mid13Rec AND RowCount<Mid23Rec THEN
				ThisLevelColor=Level2Color
				ThisLevelName=Level2Name
				DisplayBar="yes"
			ELSEIF RowCount>=Mid23Rec AND RowCount<EndRec THEN
				ThisLevelColor=Level1Color
				ThisLevelName=Level1Name
				DisplayBar="yes"
			END IF 

			IF ThisLevelName<>LastLevelName THEN %>
				<tr><td colspan=9 style="background-color:<%=ThisLevelColor%>"><FONT SIZE="3"  FACE="<%=font1%>" color="#000000"><b><%=ThisLevelName%></b></FONT></td></tr><%
				LastLevelName=ThisLevelName
			END IF  %>

			<tr>
			  <td ALIGN="Center" style="background-color:<%=ThisLevelColor%>">
				<FONT SIZE="<%=fontsize1%>"  FACE="<%=font1%>" color="<%=TypeColor5%>"><%=Rank%></FONT></td><%


			' --- Displays one line of data ---
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


			NEXT	

			Rank=Rank+1

		END IF 	%>

		</TR><% 
		rowCount = rowCount + 1
		IF RowCount>=EndRec THEN EXIT DO	' --- End of range of data selected by BRACKET ---
		rs.movenext

	LOOP %>

	</TABLE>

<br><br><%

END SUB





' -------------------------------------------------------------------------
  SUB LoadGRBrackets
' -------------------------------------------------------------------------


%>
<select name="BracketSelected" style="width:9em" >
  <option value ="0" <%IF BracketSelected = "0" THEN Response.Write(" selected ")%> >ALL</Option><br>
  <option value ="3" <%IF BracketSelected = "3" THEN Response.Write(" selected ")%> ><%=Level3Name%></Option><br>
  <option value ="2" <%IF BracketSelected = "2" THEN Response.Write(" selected ")%> ><%=Level2Name%></Option><br>
  <option value ="1" <%IF BracketSelected = "1" THEN Response.Write(" selected ")%> ><%=Level1Name%></Option><br>
</select><%

END SUB


' -------------------------------------------------------------------------
  SUB LoadGenderDrop
' -------------------------------------------------------------------------


%>
<select name="GenderSelected" style="width:9em" >
  <option value ="" <%IF GenderSelected = "" THEN Response.Write(" selected ")%> >None</Option><br>
  <option value ="M" <%IF GenderSelected = "M" THEN Response.Write(" selected ")%> >Men</Option><br>
  <option value ="W" <%IF GenderSelected = "W" THEN Response.Write(" selected ")%> >Women</Option><br>
</select><%

END SUB



' ----------------------
  SUB NoRecordMessage
' ----------------------


%>
<br>
<center><font color="<%=TextColor2%>" size="3"><b>Select Event and press Display Report button.</b></font></center><%

END SUB




' ----------------------
  SUB CreatePageHead
' ----------------------

IF sSkiYearID=1 THEN sSkiYear="Official"

'response.write("<Br>AdminMenuLevel="&AdminMenuLevel)
%>

<TABLE class="GRTable2" Align=center WIDTH=725px height=200 background="<%=MainImage%>">
<form action="<%=ThisFileName%>" method="post">
   <input type="hidden" name="WhatHeadFoot" value="<%=WhatHeadFoot%>">
  <TR>
	<td colspan=4>
		<font color="<%=TextColor2%>" size="3" face="<%=font1%>"><b>&nbsp;&nbsp;<%=sSkiYear%>&nbsp;<%=PageTitle%></b></font>
		<br>
		<font color="<%=TextColor1%>" size="2" face="<%=font1%>"><b>&nbsp;&nbsp;<%=PageSubTitle%></b></font>
	</td><%
	IF AdminMenuLevel>=50 THEN  %>	
  		<td colspan=1 valign=top align="left">
			<FONT COlOR="<% =TitleColor %>" size="<% =fontsize2 %>"><b>Show SQL</b></font>
			<input type=checkbox name="sShowSQL" <% IF sShowSQL="on" THEN response.write "checked" %>>

		</td>
  		<td colspan=1 valign=top align="left">
			<FONT COlOR="<% =TitleColor %>" size="<% =fontsize2 %>"><b>Stop</b></font>
			<input type=checkbox name="sStop" <% IF sStop="on" THEN response.write "checked" %>>

		</td><%
	ELSE  %>
  		<td colspan=2>&nbsp;</td><%
	END IF %>
	<td colspan=2>&nbsp;</td>

  </TR>

  <TR>
	<td align="right" width=60px>
		<font color="<%=TextColor2%>" size="1" face="<%=font1%>">Gender&nbsp;</font>
	</td>
	<td align="left" width=140px><%
		LoadGenderDrop %>
	</td>

	<td align="right" width=60px>
		<font color="<%=TextColor2%>" size="1" face="<%=font1%>">Year&nbsp;</font>
	</td>
	<td align="left" width=140px><%
		' --- SUB located in this program ---
		BuildSkiYearDrop %>
	</td>

	<td colspan=2>&nbsp;</td>
	<td colspan=2>&nbsp;</td>
  </TR>

  <TR>
	<td align="right">
		<font color="<%=TextColor2%>" size="1" face="<%=font1%>">Event&nbsp;</font>
	</td>
	<td align="left"><%
		'LoadAWSAEvents
		LoadGREvents %>
	</td>
	<td colspan=2>&nbsp;</td>
	<td colspan=2>&nbsp;</td>
	<td colspan=2>&nbsp;</td>
  </TR>
  <TR>
	<td align="right">
		<font color="<%=TextColor2%>" size="1" face="<%=font1%>">Bracket&nbsp;</font>
	</td>
	<td align="left"><%
		LoadGRBrackets %>
	</td>
	<td colspan=2>&nbsp;</td>
	<td colspan=2>&nbsp;</td>
	<td colspan=2>&nbsp;</td>
  </TR>

  <TR>
	<td>&nbsp;</td>
	<td >
	  	<input type="submit" name="action" value="Display Report">
	</td>

	<td colspan=2 align=center>
		<input type="submit" name="action" value="Return to Menu">
	</td>

	<td colspan=2>&nbsp;</td>
	<td colspan=2>&nbsp;</td>
  </TR>	
</form>
</TABLE>

<%



END SUB




' ---------------------
  SUB BuildSkiYearDrop
' ---------------------

%>
	<select name='sSkiYearID'><%

	SET rsSY=Server.CreateObject("ADODB.recordset")
	sSQL = "SELECT DISTINCT SY.SkiYearID, SY.SkiYearName"
	sSQL = sSQL + " FROM " &SkiYearTableName&" AS SY"
	rsSY.open sSQL, SConnectionToTRATable


	' --- Loads dropdown and sets default to Session("SkiYear")
	DO WHILE NOT rsSY.eof

		IF TRIM(rsSY("SkiYearID")) = sSkiYearID THEN
			response.write("<option value =""" & rsSY("SkiYearID") &""" selected>")
			response.write(rsSY("SkiYearName"))
			response.write("</option><br>")
		ELSE
			response.write("<option value =""" & rsSY("SkiYearID") &""">")
			response.write(rsSY("SkiYearName"))
			response.write("</option><br>")
		END IF 

		rsSY.moveNEXT

	LOOP

	rsSY.close %>
	</select>
<%

END SUB






' ---------------------
  SUB GRRanking
' ---------------------

' --- Defines which multiplier to use based on EventSelected
BonusMult=EVAL("BonMult_"&EventSelected)

sSQL = " SELECT MT.FirstName, MT.LastName, ST.MemberID"

OneCount="Y"
IF OneCount="Y" THEN
	sSQL = sSQL + ", Coalesce(TourCnt,0)+Coalesce(TourCntGR,0) AS [# of<br>Tours]"
ELSE
	' --- Use these two to show independent counts from Scores and ScoresGR ---
	sSQL = sSQL + ", TourCnt AS [# of<br>Tours]"
	sSQL = sSQL + ", TourCntGR AS [# of<br>GRTours]"
END IF


' --- Bonus Points ---

' --- Ranking Value with Bonus Points ---
SELECT CASE EventSelected 
  CASE "T"
	sSQL = sSQL + "	, CAST(CAST(SumScore/NumScores AS DECIMAL(7,0)) AS Char(10)) AS [Avg<br>Score]"
	sSQL = sSQL + "	, CASE WHEN Coalesce(TourCnt,0)+Coalesce(TourCntGR,0)>1 THEN CAST(CAST((Coalesce(TourCnt,0)+Coalesce(TourCntGR,0)-1)*5 AS DECIMAL(7,0)) AS Char(10)) ELSE '--' END AS [<a title='Bonus Points are awarded for each extra tournament'> Participation<br>Bonus Points</a>]"
	sSQL = sSQL + "	, CAST(CAST(SumScore/NumScores + (Coalesce(TourCnt,0)+Coalesce(TourCntGR,0)-1)*"&BonusMult&" AS DECIMAL(7,0)) AS Char(10)) AS [Ranking]"
  CASE ELSE
	sSQL = sSQL + "	, CAST(CAST(SumScore/NumScores AS DECIMAL(7,2)) AS Char(10)) AS [Avg<br>Score]"
	sSQL = sSQL + "	, CASE WHEN Coalesce(TourCnt,0)+Coalesce(TourCntGR,0)>1 THEN CAST(CAST((Coalesce(TourCnt,0)+Coalesce(TourCntGR,0)-1)*5 AS DECIMAL(7,0)) AS Char(10)) ELSE '--' END AS [<a title='Bonus Points are awarded for each extra tournament'> Participation<br>Bonus Points</a>]"
	sSQL = sSQL + "	, CAST(CAST(SumScore/NumScores + (Coalesce(TourCnt,0)+Coalesce(TourCntGR,0)-1)*"&BonusMult&" AS DECIMAL(7,2)) AS Char(10)) AS [Ranking]"
END SELECT


sSQL = sSQL + " 	FROM "


' --- Finds SUM and COUNT of scores for this member ---
sSQL = sSQL + " 	(SELECT MemberID, Event, COUNT(Score) AS NumScores, SUM(Score) AS SumScore"
sSQL = sSQL + "			FROM"

' --- All Scores from SCORES table that are in SkiYear ---
sSQL = sSQL + " 		(SELECT MemberID, Event, Score"
sSQL = sSQL + "			   FROM usawsrank.scores "
sSQL = sSQL + " 			WHERE Class IN ('F','N') AND Event='"&EventSelected&"' AND LEFT(Div,1) NOT IN ('C')"
IF LEFT(GenderSelected,1) = "M" THEN sSQL = sSQL + " AND LEFT(Div,1) IN ('M', 'B')"
IF LEFT(GenderSelected,1) = "W" THEN sSQL = sSQL + " AND LEFT(Div,1) IN ('W', 'G')"

sSQL = sSQL + " 			  AND TourID IN"
sSQL = sSQL + " 			     (SELECT Distinct TourID FROM usawsrank.Scores"
sSQL = sSQL + " 			     	JOIN "
sSQL = sSQL + " 			      	  (SELECT TournAppID, TDateE, TDateS FROM sanctions.dbo.TSchedul) AS TS"
sSQL = sSQL + " 			     	ON TournAppID=LEFT(TourID,6)" 
sSQL = sSQL + " 			     	WHERE TDateE<=(SELECT EndDate FROM usawsrank.SkiYear WHERE SkiYearID='"&sSkiYearID&"')"
sSQL = sSQL + " 			     	 AND  TDateE>=(SELECT BeginDate FROM usawsrank.SkiYear WHERE SkiYearID='"&sSkiYearID&"')"
sSQL = sSQL + " 			     	 )"

sSQL = sSQL + " 		UNION"

' --- All Scores from SCORESGR table that are in SkiYear ---
sSQL = sSQL + " 		SELECT MemberID, Event, Score"
sSQL = sSQL + "			   FROM usawsrank.scoresGR "
sSQL = sSQL + " 			WHERE Class IN ('F', 'N')"
sSQL = sSQL + " 			  AND IsNull(score,1)<>1"
sSQL = sSQL + " 			  AND Event='"&EventSelected&"'" 
sSQL = sSQL + " 			  AND LEFT(Div,1) NOT IN ('C')"
IF LEFT(GenderSelected,1) = "M" THEN sSQL = sSQL + " AND LEFT(Div,1) IN ('M', 'B')"
IF LEFT(GenderSelected,1) = "W" THEN sSQL = sSQL + " AND LEFT(Div,1) IN ('W', 'G')"

sSQL = sSQL + " 			  AND TourID IN"
sSQL = sSQL + " 			     (SELECT Distinct TourID FROM usawsrank.ScoresGR"
sSQL = sSQL + " 			     	JOIN "
sSQL = sSQL + " 			      	  (SELECT TournAppID, TDateE, TDateS FROM sanctions.dbo.TSchedul) AS TS"
sSQL = sSQL + " 			     	ON TournAppID=LEFT(TourID,6)" 
sSQL = sSQL + " 			     	   WHERE TDateE<=(SELECT EndDate FROM usawsrank.SkiYear WHERE SkiYearID='"&sSkiYearID&"')"
sSQL = sSQL + " 			     	 	AND  TDateE>=(SELECT BeginDate FROM usawsrank.SkiYear WHERE SkiYearID='"&sSkiYearID&"')"
sSQL = sSQL + " 			     )"


sSQL = sSQL + " 		) AS UN_SCR"
sSQL = sSQL + " 		GROUP BY MemberID, Event) AS ST"


' --- Member personal information ---
sSQL = sSQL + " 	LEFT JOIN "
sSQL = sSQL + " 		( SELECT FirstName, LastName, PersonIDwithCheckDigit"
sSQL = sSQL + " 			FROM usawaterski.dbo.Members) AS MT"
sSQL = sSQL + " 		ON MT.PersonIDwithCheckDigit=ST.MemberID"			

' --- Counts the number of tournaments from SCORES to calculate participation points ---
sSQL = sSQL + " 	LEFT JOIN "
sSQL = sSQL + " 		(SELECT MemberID, Event, COALESCE(Count(*),0) AS TourCnt"
sSQL = sSQL + " 			FROM"  
sSQL = sSQL + " 			(SELECT Distinct TourID, MemberID, Event"
sSQL = sSQL + " 				FROM usawsrank.Scores"
sSQL = sSQL + " 				WHERE Class IN ('F','N') AND LEFT(Div,1) NOT IN ('C')"

sSQL = sSQL + " 				  AND TourID IN"
sSQL = sSQL + " 				     (SELECT Distinct TourID FROM usawsrank.Scores"
sSQL = sSQL + " 			     		JOIN "
sSQL = sSQL + " 				      	  (SELECT TournAppID, TDateE, TDateS FROM sanctions.dbo.TSchedul) AS TS"
sSQL = sSQL + " 				     	ON TournAppID=LEFT(TourID,6)" 
sSQL = sSQL + " 			     		   WHERE TDateE<=(SELECT EndDate FROM usawsrank.SkiYear WHERE SkiYearID='"&sSkiYearID&"')"
sSQL = sSQL + " 			     	 		AND  TDateE>=(SELECT BeginDate FROM usawsrank.SkiYear WHERE SkiYearID='"&sSkiYearID&"')"
sSQL = sSQL + " 				     )"

sSQL = sSQL + " 				GROUP BY Event, MemberID, TourID) T"
sSQL = sSQL + " 		 	GROUP BY MemberID, Event) AS TCNT"
sSQL = sSQL + "			ON ST.MemberID=TCNT.MemberID AND ST.Event=TCNT.Event"


' --- Counts the number of tournaments from SCORESGR to calculate participation points ---
sSQL = sSQL + " 	LEFT JOIN "
sSQL = sSQL + " 		(SELECT MemberID, Event, COALESCE(Count(*),0) AS TourCntGR"
sSQL = sSQL + " 			FROM"  
sSQL = sSQL + " 			(SELECT Distinct TourID, MemberID, Event"
sSQL = sSQL + " 				FROM usawsrank.ScoresGR"
sSQL = sSQL + " 				WHERE Class IN ('F','N') AND IsNull(score,1)<>1"
sSQL = sSQL + " 				  AND TourID IN"
sSQL = sSQL + " 				     (SELECT Distinct TourID FROM usawsrank.ScoresGR"
sSQL = sSQL + " 			     		JOIN "
sSQL = sSQL + " 				      	  (SELECT TournAppID, TDateE, TDateS FROM sanctions.dbo.TSchedul) AS TS"
sSQL = sSQL + " 				     	ON TournAppID=LEFT(TourID,6)" 
sSQL = sSQL + " 			     		   WHERE TDateE<=(SELECT EndDate FROM usawsrank.SkiYear WHERE SkiYearID='"&sSkiYearID&"')"
sSQL = sSQL + " 			     	 		AND  TDateE>=(SELECT BeginDate FROM usawsrank.SkiYear WHERE SkiYearID='"&sSkiYearID&"')"
sSQL = sSQL + " 				     )"

sSQL = sSQL + " 				GROUP BY Event, MemberID, TourID) T"
sSQL = sSQL + " 		 	GROUP BY MemberID, Event) AS TCNTGR"
sSQL = sSQL + "			ON ST.MemberID=TCNTGR.MemberID AND ST.Event=TCNTGR.Event"


sSQL = sSQL + " ORDER BY ST.Event"
sSQL = sSQL + "		, SumScore/NumScores + (Coalesce(TourCnt,0)+Coalesce(TourCntGR,0)-1)*"&BonusMult&" DESC"



IF sShowSQL="on" THEN
response.write(sSQL)
'response.end
END IF
'response.write(sSQL)
'response.end


rs.open sSQL, SConnectionToTRATable


IF NOT rs.eof THEN
	rs.MoveFirst
	RecCnt=0
	DO WHILE NOT rs.eof
		RecCnt=RecCnt+1
		rs.MoveNext
	LOOP
END IF		


END SUB







%>