<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include file="Tools_Include.asp"-->
<%

DefineTRAStyles 

Dim SkiYearSelected









' --- Request from Dropdown and if a value is present reset session variable ---
SkiYearSelected=TRIM(Request("SkiYear"))
IF SkiYearSelected<>"" THEN
	Session("SkiYear")=SkiYearSelected
	sSkiYearID = SkiYearSelected
ELSE
	sSkiYearID = "12"
	Session("SkiYear")="12"
	SkiYearSelected="12"
END IF


' --- If there is a session variable then assign PageSubTitle to SkiYearName when match ---
set rsSelectFields=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT * FROM " & SkiYearTableName
rsSelectFields.open sSQL, SConnectionToTRATable
DO WHILE not rsSelectFields.eof
	IF TRIM(rsSelectFields("SkiYearID")) = session("SkiYear") THEN	
		PageSubTitle = rsSelectFields("SkiYearName")
	END IF
	rsSelectFields.MoveNext
LOOP
rsSelectFields.close





WriteIndexPageHeader


' -----------   Write Headers for DB Page  ---------------- 

PageTitle = "Ranking List Champions"


%>

<TABLE class="innertable" WIDTH="100%">
  <TR>
	<TH align="center" >
		<FONT color=#ffffff size=5><B><%=PageTitle%></B></FONT>
		<br>	
		<FONT color=#ffffff size=3><b><%=PageSubTitle%></b></B></FONT>
    	</TH>	

  </TR>
<tr>
   <form method=post action="RankChampions.asp">
  <td align="center"><font size=<% =fontsize3 %> COlOR="#FFFFFF">Range:</font>
  
    <select name='SkiYear' style="width: 150px"><%

	set rsSelectFields=Server.CreateObject("ADODB.recordset")
	sSQL = "SELECT * FROM " & SkiYearTableName& " ORDER BY SkiYearID"
	rsSelectFields.open sSQL, SConnectionToTRATable

	DO WHILE not rsSelectFields.eof
		IF TRIM(rsSelectFields("SkiYearID")) = "9" OR TRIM(rsSelectFields("SkiYearID")) = "1" OR TRIM(rsSelectFields("SkiYearID")) = "8" THEN

		ELSEIF Date <= rsSelectFields("EndDate") THEN

		ELSEIF TRIM(rsSelectFields("SkiYearID")) = session("SkiYear") THEN
			response.write("<option value =""" & rsSelectFields("SkiYearID") &""" selected>")
			response.write(rsSelectFields("SkiYearName"))
			response.write("</option><br>")
  		ELSE
			response.write("<option value =""" & rsSelectFields("SkiYearID") &""">")
			response.write(rsSelectFields("SkiYearName"))
			response.write("</option><br>")
		END IF
		rsSelectFields.movenext
  	LOOP
	rsSelectFields.close %>

    </select>

  <br>
    <input type="Submit" value="Submit" >
  </td>	
  </form> 	

 </tr>
</TABLE> 
<br><%




' -------------   Displays table heading  -------------------

%>
<TABLE class="innertable" width=100%>

  <TR>
    <TH ALIGN="Left"><font size=<% =fontsize3 %> COlOR="#FFFFFF"><b>Event</b></FONT></TH>
    <TH COLSPAN=2 ALIGN="Left"><font size=<% =fontsize3 %> COlOR="#FFFFFF"><b>1st</b></FONT></TH>
    <TH COLSPAN=2 ALIGN="Left"><font size=<% =fontsize3 %> COlOR="#FFFFFF"><b>2nd</b></FONT></TH>
    <TH COLSPAN=2 ALIGN="Left"><font size=<% =fontsize3 %> COlOR="#FFFFFF"><b>3rd</b></FONT></TH>
  </TR>
  <tr>	
    <td colspan=7>&nbsp;</td>
  </TR>
<%
' Loops through all rows of SELECT


Set rsDivList=Server.CreateObject("ADODB.recordset")

sSQL = "SELECT DISTINCT RT.Div, RT.event, DT.Div_Name"
sSQL = sSQL + " FROM "&RankTableName&" AS RT"
sSQL = sSQL + " JOIN " &DivisionsTableName&" AS DT ON DT.Div=RT.Div"
sSQL = sSQL + " WHERE RT.SkiYearID='"&SkiYearSelected&"' AND LOWER(LEFT(RT.Div,1)) IN ('w','m','b','g','o')"
sSQL = sSQL + " ORDER BY RT.Div, RT.Event"
rsDivList.open sSQL, sConnectionToTRATable, 3, 1

CurrentDiv=rsDivList("Div")
CurrentEvent=rsDivList("Event")
CurrentDivName=rsDivList("Div_Name")




'markdebug("sSkiYearID="&sSkiYearID)
'sSkiYearID="9"

CurrentEvent="XX"
CurrentDiv="XX"


Set rs=Server.CreateObject("ADODB.recordset")

DO WHILE Not rsDivList.EOF 

	' --- Display line for this Division ---- 

	IF rsDivList("Div")<>CurrentDiv THEN %>
	  <TD COLSPAN=3 ALIGN="Center" vAlign="top"><font size=<% =fontsize3 %> face=<% =font1 %> COlOR="<% =TextColor1 %>"><b><%=rsDivList("Div_Name")%></b></FONT></TD>
	  <TD COLSPAN=4>&nbsp;</TD> 	<%	          
	END IF 

	CurrentEvent=rsDivList("event")
	CurrentDiv=rsDivList("div")
	CurrentDivName=rsDivList("Div_Name") %>	
	<tr><%
	  SELECT CASE TRIM(CurrentEvent) 
		CASE "S" %>
			<TD ALIGN="Center" vAlign="top"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">Slalom</FONT></TD><%	          
		CASE "T" %>
			<TD ALIGN="Center" vAlign="top"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">Trick</FONT></TD><%	          
		CASE "J" %>
			<TD ALIGN="Center" vAlign="top"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">Jump</FONT></TD><%	          
		CASE "O" %>
			<TD ALIGN="Center" vAlign="top"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">Overall</FONT></TD><%	          
	END SELECT	

	' --- Query of Top 3 in this event/div ---
	sSQL = "SELECT TOP 4 RT.Div, RT.Event, RT.MemberID, MT.FirstName, MT.LastName, coalesce(RT.RankScore,0) AS RankScore"
	sSQL = sSQL + " FROM "&RankTableName&" AS RT"
	sSQL = sSQL + " JOIN "&MemberTableName&" AS MT ON RT.MemberID=MT.PersonIDWithCheckDigit"
	sSQL = sSQL + " WHERE RT.Event='"&CurrentEvent&"' AND RT.Div='"&CurrentDiv&"' AND RT.SkiYearID='"&SkiYearSelected&"'"
	sSQL = sSQL + " AND MT.FederationCode='USA'"
	sSQL = sSQL + " ORDER BY RankScore DESC"
	rs.open sSQL, sConnectionToTRATable, 3, 1
	rs.MoveFirst

	Name1=""
	Name2=""
	Name3=""
	Name4=""
	Score1=0
	Score2=0
	Score3=0
	Score3=0
	RCount=0

	DO WHILE NOT rs.eof
		RCount=RCount+1
	   	IF NOT rs.EOF THEN
			SELECT CASE RCount
				CASE 1
					Name1=rs("FirstName")&" "&rs("LastName")
					Score1=rs("RankScore")
				CASE 2
					Name2=rs("FirstName")&" "&rs("LastName")
					Score2=rs("RankScore")
					IF Score1=Score2 THEN
						Name1=Name1&" (tie)<br>&nbsp;"&rs("FirstName")&" "&rs("LastName")&" (tie)"
						Name2=""
						Score2=0
					END IF
				CASE 3
					Name3=rs("FirstName")&" "&rs("LastName")
					Score3=rs("RankScore")
					IF Score1=Score3 THEN
						Name1=Name1&"<br>&nbsp;"&rs("FirstName")&" "&rs("LastName")&" (tie)"
						Name3=""
						Score3=0
					ELSEIF Score2=Score3 THEN
						Name2=Name2&" (tie)<br>&nbsp;"&rs("FirstName")&" "&rs("LastName")&" (tie)"
						Name3=""
						Score3=0
					END IF

				CASE 4
					Name4=rs("FirstName")&" "&rs("LastName")
					Score4=rs("RankScore")
					IF Score1=Score4 THEN
						Name1=Name1&"<br>&nbsp;"&rs("FirstName")&" "&rs("LastName")&" (tie)"
						Name4=""
						Score4=0
					ELSEIF Score2=Score4 THEN
						Name2=Name2&"<br>&nbsp;"&rs("FirstName")&" "&rs("LastName")&" (tie)"
						Name4=""
						Score4=0
					ELSEIF Score3=Score4 THEN
						Name3=Name3&" (tie)<br>&nbsp;"&rs("FirstName")&" "&rs("LastName")&" (tie)"
						Name4=""
						Score4=0
					END IF

			END SELECT

			rs.MoveNext
		END IF
	LOOP %>	

	<TD ALIGN="Left" vAlign="top"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=Name1%></FONT></TD><%
	IF Score1=0 THEN %>
		<TD ALIGN="Right" vAlign="top">&nbsp;</FONT></TD> <%
	ELSEIF TRIM(rsDivList("Event"))="T" OR TRIM(rsDivList("Event"))="O" THEN %>
		<TD ALIGN="Right" vAlign="top"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=FormatNumber(Score1,0)%></FONT></TD> <%
	ELSEIF TRIM(rsDivList("Event"))="J" OR TRIM(rsDivList("Event"))="S" THEN %>
		<TD ALIGN="Right" vAlign="top"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=FormatNumber(Score1,2)%></FONT></TD> <%
	ELSE %>
		<TD ALIGN="Right" vAlign="top">&nbsp;</FONT></TD> <%
	END IF %>

	<TD ALIGN="Left" vAlign="top"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=Name2%></FONT></TD><%
	IF Score2=0 THEN %>
		<TD ALIGN="Right" vAlign="top">&nbsp;</FONT></TD> <%
	ELSEIF TRIM(rsDivList("Event"))="T" OR TRIM(rsDivList("Event"))="O" THEN %>
		<TD ALIGN="Right" vAlign="top"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=FormatNumber(Score2,0)%></FONT></TD> <%
	ELSEIF TRIM(rsDivList("Event"))="J" OR TRIM(rsDivList("Event"))="S" THEN  %>
		<TD ALIGN="Right" vAlign="top"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=FormatNumber(Score2,2)%></FONT></TD> <%
	ELSE %>
		<TD ALIGN="Right" vAlign="top">&nbsp;</FONT></TD> <%
	END IF %>

	<TD ALIGN="Left" vAlign="top"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=Name3%></FONT></TD><%
	IF Score3=0 THEN %>
		<TD ALIGN="Right" vAlign="top">&nbsp;</FONT></TD> <%
	ELSEIF TRIM(rsDivList("Event"))="T" OR TRIM(rsDivList("Event"))="O" THEN %>
		<TD ALIGN="Right" vAlign="top"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=FormatNumber(Score3,0)%></FONT></TD> <%
	ELSEIF TRIM(rsDivList("Event"))="J" OR TRIM(rsDivList("Event"))="S" THEN  %>
		<TD ALIGN="Right" vAlign="top"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;<%=FormatNumber(Score3,2)%></FONT></TD> <%
	ELSE %>
		<TD ALIGN="Right" vAlign="top">&nbsp;</FONT></TD> <%
	END IF  %>

	</tr><%

	rs.close
	rsDivList.MoveNext 
'EXIT DO

LOOP  %>

</TABLE><%


rsDivList.close
Set rsDivList = nothing


WriteIndexPageFooter




' --------------------
    SUB GetTop3Rank
' --------------------



END SUB





%>









