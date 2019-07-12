<!--#include virtual="/rankings/settingsHQ.asp"-->

<%

' -------------------------------------------------------------------------------------------------
' Displays Rankings in Regional Guidebook Format
'
'

Dim RegionSel, EventSel, SkiYrSel
Dim PrevDiv, PrevEvent
tNatl_Plc = ""
tRegl_Plc = ""
SeqDisp="2"


WriteIndexPageHeader


' Sets default values for Region and SkiYear
RegionSel = trim(Request("RegionSel"))
EventSel = trim(Request("EventSel"))
SkiYrSel = trim(Request("SkiYrSel"))

ThisPage = Request.ServerVariables("SCRIPT_NAME")

%>

<table class="droptable" width="<%=TourTableWidth%>" align=center>
<tr><td colspan=3 align="center">
<br><center><font size="5"><b>Rankings Formatted for Regional Guidebooks</font></b><br>
</td></tr>
<%
IF RegionSel="" or EventSel="" or SkiYrSel="" THEN
	%>
<tr><td colspan=3 align="center">
	<br>
	<font color="red">Select Desired Parameters for the Rankings you wish to extract, then click 
	<br> the <b>Display Rankings</b> button below the selection boxes.
	<BR></font><% 	
 	END IF
%>
</td></tr>
<tr><td><br></td></tr>

<form action="/rankings/guidebook.asp" method="post">

<tr>
	
<td>&nbsp;&nbsp;&nbsp;&nbsp;Region:&nbsp;
<select name="RegionSel" align=center">
<option value="1"<%IF RegionSel = "1" THEN Response.Write("Selected" )%>>South Central</option>
<option value="2"<%IF RegionSel = "2" THEN Response.Write("Selected" )%>>Midwest</option>
<option value="3"<%IF RegionSel = "3" THEN Response.Write("Selected" )%>>West</option>
<option value="4"<%IF RegionSel = "4" THEN Response.Write("Selected" )%>>South</option>
<option value="5"<%IF RegionSel = "5" THEN Response.Write("Selected" )%>>East</option>
</select>&nbsp;</td>

<td>Event:&nbsp;
<select name="EventSel" align="center">
<option value ="S" <%IF EventSel = "S" THEN response.write("Selected " )%>>S, J and T</Option><br>
<option value ="O" <%IF EventSel = "O" THEN response.write("Selected " )%>>Overall</Option><br>
</select>&nbsp;</td>

<td>Ski Year:&nbsp;
<SELECT name="SkiYrSel" align="center">

<%

' ------------   Builds Ski Year Drop Down list -----------------

Set rs=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT * FROM " & SkiYearTableName & " Order by SkiYearID Desc"
rs.open sSQL, sConnectionToTRATable, 3, 1

DO WHILE not rs.eof
  response.write("<option value=""" & rs("SkiYearID") & """")

  IF trim(rs("SkiYearID")) = SkiYrSel THEN
  	SkiYearName = rs("SkiYearName")
    response.write(" SELECTed")
  END IF

  response.write(">")
  response.write(rs("SkiYearName"))
  response.write("</option><br>")
  rs.movenext
LOOP
rs.close
Set rs = nothing

%>
</SELECT></td>

</tr>

<tr><td colspan=3 align="center">
<br>
<input type="submit" name="action" value="Display Rankings"></input>
<br><br>
</td></tr>

</form>

</table>


<%

' Entire Display section now conditional on having ALL parameters chosen

IF RegionSel<>"" and EventSel<>"" and SkiYrSel<>"" THEN

' Establishes Object Array
' Writes SQL string
' Reads in array

Set rs=Server.CreateObject("ADODB.recordset")

SELECT CASE EventSel
   ' Any Event other than Overall
   CASE "S", "J", "T"
	sSQL = "SELECT "&RankTableName&".MemberID, "&RankTableName&".div, "&RankTableName&".event, "&RankTableName&".RankScore as 'score', "&RankTableName&".SkiYearID, "&MemberTableName&".firstname, "&MemberTableName&".lastname,"
	sSQL = sSQL + " "&RegionTableName&".region, UPPER("&MemberTableName&".[state]) AS 'state', coalesce("&RankTableName&".natl_plc, '999') as natl_plc, coalesce("&RankTableName&".regl_plc, '999') as regl_plc"

	sSQL = sSQL + " FROM "&RankTableName&" JOIN "&MemberTableName&" ON "&RankTableName&".MemberID = "&MemberTableName&".personidwithcheckdigit" 
	sSQL = sSQL + " LEFT JOIN "&RegionTableName&" ON LOWER("&MemberTableName&".[state]) = LOWER("&RegionTableName&".[state])"

	sSQL = sSQL + " WHERE "&RankTableName&".SkiYearID="&SkiYrSel&" AND "&RankTableName&".Event<>'O' AND "&RankTableName&".RankScore IS NOT NULL AND SUBSTRING("&RankTableName&".div,1,1)<>'C' AND SUBSTRING("&RankTableName&".div,1,1)<>'I'AND "&MemberTableName&".FederationCode = 'USA' AND "&RegionTableName&".[region] = "&RegionSel&" "
	IF SeqDisp="1" THEN
		sSQL = sSQL + " ORDER BY "&RankTableName&".SkiYearID, "&RankTableName&".event, "&RankTableName&".div, "&RankTableName&".RankScore DESC"
	ELSE
		sSQL = sSQL + " ORDER BY "&RankTableName&".SkiYearID, "&RankTableName&".div, "&RankTableName&".event, "&RankTableName&".RankScore DESC"
	END IF

   ' User selected Overall
   CASE "O"
	sSQL = "SELECT "&RankTableName&".MemberID, "&RankTableName&".div, "&RankTableName&".event, "&RankTableName&".RankScore as 'score', "&RankTableName&".SkiYearID, "&MemberTableName&".firstname, "&MemberTableName&".lastname,"
	sSQL = sSQL + " "&RegionTableName&".region, UPPER("&MemberTableName&".[state]) AS 'state'"

	sSQL = sSQL + " FROM "&RankTableName&" JOIN "&MemberTableName&" ON "&RankTableName&".MemberID = "&MemberTableName&".personidwithcheckdigit" 
	sSQL = sSQL + " LEFT JOIN "&RegionTableName&" ON LOWER("&MemberTableName&".[state]) = LOWER("&RegionTableName&".[state])"

	sSQL = sSQL + " WHERE "&RankTableName&".SkiYearID="&SkiYrSel&" AND "&RankTableName&".Event='O' AND "&RankTableName&".RankScore IS NOT NULL AND SUBSTRING("&RankTableName&".div,1,1)<>'C' AND SUBSTRING("&RankTableName&".div,1,1)<>'I'AND "&MemberTableName&".FederationCode = 'USA' AND "&RegionTableName&".[region] = "&RegionSel&" "
	IF SeqDisp="1" THEN
		sSQL = sSQL + " ORDER BY "&RankTableName&".SkiYearID, "&RankTableName&".event, "&RankTableName&".div, "&RankTableName&".RankScore DESC"
	ELSE
		sSQL = sSQL + " ORDER BY "&RankTableName&".SkiYearID, "&RankTableName&".div, "&RankTableName&".event, "&RankTableName&".RankScore DESC"
	END IF
END SELECT

WriteDebugSql (sSQL)

rs.open sSQL, sConnectionToTRATable, 3, 1


' Display table heading then
' Loops through all rows of SELECT
%>



<TABLE class="innertable" ALIGN="Center" width="<%=TourTableWidth%>">
    <TR>	
    <TD ALIGN="Center" vAlign="top" colspan=6>
	<FONT COlOR="#000000" SIZE="2"><b><i><%=RegionSelected%>
	<BR><%=SkiYearName%></i></b>
	</FONT>
    </TD>
    </TR>
    <TR> 
    <TH ALIGN="Left"><Left><FONT COlOR="#FFFFFF" SIZE="1">Member</FONT></TH>
    <TH ALIGN="Center"><FONT COlOR="#FFFFFF" SIZE="1">Score</FONT></TH>
    <TH ALIGN="Center"><FONT COlOR="#FFFFFF" SIZE="1">State</FONT></TH>
    <TH ALIGN="Center"><FONT COlOR="#FFFFFF" SIZE="1">NatlPlc</FONT></TH>
    <TH ALIGN="Center"><FONT COlOR="#FFFFFF" SIZE="1">ReglPlc</FONT></TH>
    <TH ALIGN="Center"><FONT COlOR="#FFFFFF" SIZE="1">MemberID</FONT></TH>
  </TR>
  <TR><TD colspan=6>&nbsp;</TD></TR>

<%


PrevEvent="XX"
PrevDiv="XX"

DO WHILE Not rs.EOF 
        
	SELECT CASE TRIM(rs("event"))
		CASE "T"
			Fscore=FormatNumber(rs("score"),0)
			sEventName="Tricks"
		CASE "O"
			Fscore=FormatNumber(rs("score"),0)
			sEventName="Overall"
		CASE "J"
			Fscore=FormatNumber(rs("score"),2)
			sEventName="Jump"
		CASE "S"
			Fscore=FormatNumber(rs("score"),2)
			sEventName="Slalom"
		CASE ELSE
			sEventName="???"
	END SELECT

	IF rs("div")<>PrevDiv OR rs("event")<>PrevEvent THEN
		%>
		<tr>
		<TH ALIGN="Center"><FONT COlOR="#FFFFFF" SIZE="1"><%=rs("div")%></FONT></TH>
		<TH ALIGN="Center"><FONT COlOR="#FFFFFF" SIZE="1"><%=sEventName%></FONT></TH>
		<TH bgcolor="#C0C0C0">&nbsp;</TH>
		<TH bgcolor="#C0C0C0">&nbsp;</TH>
		<TH bgcolor="#C0C0C0">&nbsp;</TH>
		<TH bgcolor="#C0C0C0">&nbsp;</TH>
		</tr>
   		<%	
	END IF


	' Fills variables with a space so table will fill all cells 
	IF TRIM(rs("Event")) = "S" OR TRIM(rs("Event")) = "J" OR TRIM(rs("Event")) = "T" THEN
		tNatl_Plc = rs("Natl_Plc")
		tRegl_Plc = rs("Regl_Plc")
	ELSE
		tNatl_Plc = ""
		tRegl_Plc = ""
	END IF		

	IF tNatl_Plc = "" THEN tNatl_Plc="&nbsp;"
	IF tRegl_Plc = "" THEN tRegl_Plc="&nbsp;"
%>	

	
	<tr>
	<TD ALIGN="Left" vAlign="top"><FONT COlOR="#000000" SIZE="1"><%=TRIM(rs("lastname"))&", "&rs("firstname")%></FONT></TD>
        <TD ALIGN="Center" vAlign="top"><FONT COlOR="#000000" SIZE="1"><%=Fscore%></FONT></TD>
        <TD ALIGN="Center" vAlign="top"><FONT COlOR="#000000" SIZE="1"><%=rs("state")%></FONT></TD>
        <TD ALIGN="Center" vAlign="top"><FONT COlOR="#000000" SIZE="1"><%=tNatl_plc%></FONT></TD>
        <TD ALIGN="Center" vAlign="top"><FONT COlOR="#000000" SIZE="1"><%=tRegl_plc%></FONT></TD>
        <TD ALIGN="Center" vAlign="top"><FONT COlOR="#000000" SIZE="1"><%=rs("MemberID")%></FONT></TD>
        </tr>

<% 

	PrevEvent=rs("event")
	PrevDiv=rs("div")

	rs.MoveNext	
%> 

<%
LOOP 
rs.close
Set rs = nothing
%>
</TABLE>

<%

' End of Conditional IF on Display logic

END IF

WriteIndexPageFooter


%>




