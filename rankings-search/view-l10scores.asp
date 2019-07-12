<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->

<%


' -----------------------------------------------------------------------------------------------------------
' ----------------------------------------- MAIN CODE -------------------------------------------------------
' -----------------------------------------------------------------------------------------------------------
'


Dim EventSelected, DivName, sMemberID
Dim sSQL
Dim TourDisplayWidth
Dim rowCount, i, Tempcolor
Dim  MainImage
Dim sConnection
dim sConnectionSantions

sConnection = "Provider=SQLOLEDB;SERVER=sql04.epolk.net;uid=tr45t4nd22;Password=Sk133r3t;Initial Catalog=cobra00025"
sConnectionSanctions = "Provider=SQLOLEDB;SERVER=sql04.epolk.net;uid=tr45t4nd22;Password=Sk133r3t;Initial Catalog=Sanctions"
' sConnection = "Provider=SQLOLEDB;SERVER=localhost;uid=sa;Password=5Hundred!;Initial Catalog=cobra00025"

NewsPageNum="FAQ_Scores"
TourDisplayWidth=700
ThisFileName="view-l10scoresHQ.asp"
DefineTRAStyles


sMemberID = RIGHT(TRIM(request("sMemberID")),9)
EventSelected = left(TRIM(Request("EventSelected")),1)
DivName = TRIM(Request("divname"))
IF TRIM(EventSelected) = "" THEN EventSelected="S" 

' Image def subroutine
WhatDropDownImage EventSelected

IF TRIM(sMemberID)="" THEN
		' --- Sends user to search-member routine to selected member
		Session("sSendingPage")="/rankings/"&ThisFileName&"?pvar=FoundMember"
		Response.Redirect("/rankings/search-memberHQ.asp?formstatus=search")
ELSE
		WriteIndexPageHeader
		ScoresByMember
		WriteIndexPageFooter
	
END IF

' ***********************************************************************
SUB ScoresByMember
' ***********************************************************************
' --- To keep spammers out test for numeric value ---
IF NOT IsNumeric(sMemberID) THEN Response.redirect("/rankings/defaultHQ.asp")

' --- Checks to see if there are any scores for this MemberID ---
SET rsMemb=Server.CreateObject("ADODB.recordset")
sSQL = "Select top 1 PersonIDwithCheckDigit, LastName, FirstName, City, State, BirthDate from "&MemberTableName&" WHERE PersonIDwithCheckDigit="&sqlclean(sMemberID)
rsMemb.open sSQL, sConnection, 3, 1

IF NOT rsMemb.eof THEN
	FullName=rsMemb("FirstName")&" "&rsMemb("LastName")
	sMembCity = rsMemb("City")
	StateSelected = rsMemb("state")
	sMembAge = AgeAtDate(Date, sMemberID)
ELSE
	FullName="Not Defined"
	sMembCity = "Not Defined"
	StateSelected = "Not Defined"
	sMembAge = 150
END IF

' ------ OUTER TABLE TO HOLD BACKGROUND IMAGE ---- %>
<TABLE border=1 class="droptable" align=center height=225px width="<%=TourDisplayWidth%>" background="<%=MainImage%>" >

<TR >
  <% ' --- Skier Name, MemberID, City/St and Age  --- %> 
   
  <td style="cell-padding:3px" colspan=1 align ="left">
	<font size=4 face="<%=font2%>" color="<%=Textcolor2%>"><b><I><a title="MemberID: <%=sMemberID%>">&nbsp;&nbsp;<%=FullName%></a></I></b></font>
  </td>
  <td colspan=2 align="left">	
	<font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor2%>"><b><%=sMembCity%>, <%=StateSelected %></b></font>
	&nbsp;&nbsp;&nbsp;<font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor2%>"><b>&nbsp;&nbsp;&nbsp;Age: <%=sMembAge%></b></font>
  </td>	
</TR>

<TR>
  <TD colspan=1 width=225px></TD>
  <TD align=left> </TD>
</TR>
</TABLE>

<style TYPE="text/css">
<!--  A:link {text-decoration: none; color:#375AE2}  A:visited {text-decoration: none; color:#375AE2}  A:active {text-decoration: none}   A:hover {text-decoration: ; color:#3E85BB; }-->
</style>
<%

SET rs=Server.CreateObject("ADODB.recordset")
sSQL = "EXEC spGetL10SlalomPerformances 'SM', 'Y', '300043260'"
rs.open sSQL, sConnection, 3, 1

DisplayScoresData

END SUB


' ------------------------
SUB DisplayScoresData
' ------------------------

IF rs.eof AND request("EventSelected") <> "O" THEN  %>
	<br><br>
	<center><font color="red">No Scores Found For These Search Criteria</font></center>
	<br><br><% 

ELSE 

ScoreBackground=Tablecolor1 



' ---------  TOP OF TABLE FOR DISPLAYING SCORES ------------------ %>

<br>
<TABLE class="innertable" align="center" WIDTH="<%=TourDisplayWidth%>px">

<TR>
  <Th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Tour ID</FONT></th><%



	IF EventSelected <> "O" THEN %>

		<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Score</FONT></th>
		<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Round</FONT></th><%

		IF DivName = "" THEN %>
			<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Div</FONT></th><%
		END IF

		IF left(EventSelected,1) = "S" THEN %>
			<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="<%=Textcolor1%>">Buoys</FONT></th>
			<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="<%=Textcolor1%>">Line</FONT></th><%
		END IF
		IF left(EventSelected,1) = "J" THEN %>
			<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Ramp</FONT></th><%
		END IF
		IF left(EventSelected,1) <> "T" THEN %>
			<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Speed</FONT></th><%
		END IF %>

		<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Class</FONT></th><%


	ELSE ' Overall Scores Stuff %>
      
		<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Round</FONT></th>
		<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Div</FONT></th>
		<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Slalom</FONT></th>
		<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Trick</FONT></th>
		<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Jump</FONT></th>
		<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>"><b>Total Score</b></FONT></th><%
    
	END IF %>

    
</TR><%


    

    ' ------------------  Loop to begin displaying SCORES for Tournament  ------------------	
    DO WHILE not rs.eof  %>

	<TR>
	  <TD ALIGN="Center" vAlign="top">
	    <font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>"><% 

		SET rsSelectFields=Server.CreateObject("ADODB.recordset")
	     	sSQL = "Select top 1 TSanction,TName,TCity,TState,TDateE from "& SanctionTableName &" where lower(TournAppID) = '" & sqlclean(lCASE(TRIM(left(rs("TourID"),6)))) & "'"
		rsSelectFields.open sSQL, sConnectionSanctions, 3, 1
  

		IF rsSelectFields.EOF THEN %>
			<a href="/rankings/<%=ThisFileName%>?tour_id=<% =TRIM(rs("TourID")) %>&pvar=ByTour&DivSelected=<% =TRIM(rs("Div")) %>"><% =rs("TourID") %></a></FONT></TD><% 
		ELSE %>
			<a href="/rankings/<%=ThisFileName%>?tour_id=<% =TRIM(rs("TourID")) %>&pvar=ByTour&DivSelected=<% =TRIM(rs("Div")) %>&EventSelected=<%=EventSelected%>&sTourSportsGroup=<%=sSptsGrpID%>"
			title="<% =rsSelectFields("tname") %>&#13;<% =rsSelectFields("tcity")%>, <% =rsSelectFields("tstate")%>&#13;<% =rsSelectFields("tdatee")%>"> <% =rs("TourID") %> </a></FONT></TD><%
		END IF 

		rsSelectFields.Close



		IF EventSelected <> "O" THEN

			IF rs("SCORE") <> "" THEN
				'---  You can not "formatnumber" IF the value is null ... so we throw this check in just to prevent errors.

		        	IF EventSelected = "Trick" THEN %>
	        			<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<% Response.Write formatnumber(rs("SCORE"),0) %></FONT></TD><%
				    ELSE %>
        				<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<% Response.Write formatnumber(rs("SCORE"),2) %></FONT></TD><%
				    END IF 

			ELSE %>
        			<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;</FONT></TD><%
			END IF %>

			<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<% =rs("ROUND") %></FONT></TD>
		    <TD Align="center" valign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<a href="/rankings/view-standingsHQ.asp?pvar=National&DivSelected=<%=rs("DIV")%>&EventSelected=<%=left(EventSelected,1)%>"><% =rs("DIV") %></a></font></td> 
            <%

			IF left(EventSelected,1) = "S" THEN %>
				<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<%=rs("ALTSCORE")%></FONT></TD><%
			END IF


			' --- Display Rope, Boat, Line, Class, etc -----
			IF rs("PERF_QUAL1") <> "" THEN
				' --- You can not "formatnumber" IF the value is null ... so we throw this check in just to prevent errors.
				IF left(EventSelected,1) = "S" THEN  %>
					<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<%= formatnumber(rs("PERF_QUAL1")/100, 2)%></FONT></TD><%
				ELSE %>
					<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;</FONT></TD><%
				END IF
			ELSE %>
				<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;</FONT></TD><%
		      	END IF

			IF left(EventSelected,1) = "J" THEN %>
				<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<%=rs("PERF_QUAL1")%></FONT></TD><%
			END IF
			IF left(EventSelected,1) <> "T" THEN %>
				<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<%=rs("PERF_QUAL2")%></FONT></TD><%
			END IF %>
			<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<% =rs("CLASS") %></FONT></TD><% 

	   ELSE  ' Overall Stuff
        Tempcolor = Textcolor1   
		%>

		<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=tempcolor%>">&nbsp;<% =rs("round") %></FONT></TD>
		<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=tempcolor%>">&nbsp;<% =rs("Div") %></FONT></TD>
		<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=tempcolor%>">&nbsp;<a title="<%=rs("S_OrigScore")%>"><% IF rs("SlalomOverall") <> "" THEN Response.Write formatnumber(rs("SlalomOverAll"),1) %></a></FONT></TD>
		<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=tempcolor%>">&nbsp;<a title="<%=rs("T_OrigScore")%>"><% IF rs("TrickOverall") <> "" THEN Response.Write formatnumber(rs("TrickOverAll"),1) %></a></FONT></TD>
		<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=tempcolor%>">&nbsp;<a title="<%=rs("J_OrigScore")%>"><% IF rs("JumpOverall") <> "" THEN Response.Write formatnumber(rs("JumpOverAll"),1) %></a></FONT></TD>
		<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=tempcolor%>">&nbsp;<b><% IF rs("TotalOverall") <> "" THEN Response.Write formatnumber(rs("TotalOverAll"),1) %></b></FONT></TD><%


	   END IF %>

    </TR><% 

    rs.movenext
    LOOP  %>



    </TABLE>
    <br>
    <br>
<%


END IF  ' End if for test of existence of scores

rs.close

END SUB

%>
