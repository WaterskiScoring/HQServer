<% Option Explicit %>
<!--#include file="settingsHQ.asp"-->

<html><head><title>Post Tournament Status Processing</title></head><body>

<%

'	Mainline Code here -- set up variables, then act on "Process" Value
' Adding another comment line to see if we can save this revised file

Dim objRS, sSQL, ThisModule, Updatable, Updated, Missing, SetList
Dim ThisYear, Process, CalYear, SptsDiv, Classes, HideArchive
Dim TournAppID, TSanction, TSanType, TName, TdateE, MonthDay, TStatus
Dim TEventSlalom, TEventTrick, TEventJump, Tsite, TSiteID
Dim eMailCC, eMailTo, eMailBody
Dim PTF_SBK, PTF_WSP, PTF_TS, PTF_OD, PTF_BT, PTF_JT
Dim PTF_CS, PTF_CJ, PTF_SD, PTF_TU, PTF_HD, PTF_TNY
Dim IconName, How, TourStat


SET objRS=Server.CreateObject("ADODB.recordset")

ThisModule = "/rankings/PostTourStatus.asp"

'	Pick up all the Standard Form Variables into local Variables

ThisYear = DatePart("yyyy",Date)

Process = TRIM(Request("Process"))
IF len(Process) = 0 or Request("SptsDiv") = "Pls" then Process = "Start"

CalYear = TRIM(Request("CalYear"))
if len(Request("CalYear")) = 0 THEN CalYear = ThisYear: ELSE CalYear = Int(Trim(Request("CalYear")))

SptsDiv = TRIM(Request("SptsDiv"))
IF len(SptsDiv) = 0 then SptsDiv = "Pls"

Classes = TRIM(Request("Classes"))
HideArchive = TRIM(Request("HideArchive"))


'	Form Variable "Process" dictates what we do in this invocation.  All
'	processes done by Standard Subroutines, as outlined immediately below.

SELECT CASE Process 

CASE "Start"

	PresentChoices

CASE "Listem"

	PresentChoices
	ListTournaments

CASE "Editor"

	EditTournament
	
CASE "Update"

	UpdateTournament
	
CASE "Cancel"

	CancelTournament
	PresentChoices
	ListTournaments
	
CASE "Reinstate"

	UnCancelTournament
	PresentChoices
	ListTournaments
	
END SELECT	



'	---------------------
SUB	PresentChoices
'	---------------------	

	WriteIndexPageHeader

	%>

	<Table class="innertable" width=90% align=center><TR>

	<TH><center><b><font size="2" color="#FFFFFF">&nbsp; Specify/Revise Selection:&nbsp; </font></b></center></TH>
	<TH><center><b><font size="2" color="#FFFFFF">&nbsp; Legend to Status Indicators:&nbsp; </font></b></center></TH>
	
	</TR>
	
	<TR>    

	<TD><center><FONT size="2">
	
	<FONT Color="Red"><br>Lists only Approved Sanctions.</FONT><BR>
	
	<FORM method="post" action="<%=ThisModule%>">
	<INPUT type="hidden" name="Process" value="Listem">

	&nbsp;&nbsp; Calendar Year:&nbsp;&nbsp; 
		<SELECT name='CalYear'>
		  <option value="<%=ThisYear+1%>" <%IF CalYear=ThisYear+1 THEN response.write(" selected")%>> <%=ThisYear+1%> </Option><br>
		  <option value="<%=ThisYear%>"   <%IF CalYear=ThisYear   THEN response.write(" selected")%>> <%=ThisYear%>   </Option><br>
		  <option value="<%=ThisYear-1%>" <%IF CalYear=ThisYear-1 THEN response.write(" selected")%>> <%=ThisYear-1%> </Option><br>
		  <option value="<%=ThisYear-2%>" <%IF CalYear=ThisYear-2 THEN response.write(" selected")%>> <%=ThisYear-2%> </Option><br>
		  <option value="<%=ThisYear-3%>" <%IF CalYear=ThisYear-3 THEN response.write(" selected")%>> <%=ThisYear-3%> </Option><br>
		  <option value="<%=ThisYear-4%>" <%IF CalYear=ThisYear-3 THEN response.write(" selected")%>> <%=ThisYear-4%> </Option><br>
		  <option value="<%=ThisYear-5%>" <%IF CalYear=ThisYear-3 THEN response.write(" selected")%>> <%=ThisYear-5%> </Option><br>
		  <option value="<%=ThisYear-6%>" <%IF CalYear=ThisYear-3 THEN response.write(" selected")%>> <%=ThisYear-6%> </Option><br>
		</select><BR>&nbsp;<BR>
		
		<SELECT name='SptsDiv'>
		  <option value ='Pls' <%IF SptsDiv="Pls" THEN response.write(" selected")%>> [ Pls Select Jurisdiction ]</Option><br>
		  <option value ='AAR' <%IF SptsDiv="AAR" THEN response.write(" selected")%>>AWSA All Regions</Option><br>
		  <option value ='AEA' <%IF SptsDiv="AEA" THEN response.write(" selected")%>>AWSA Eastern Region</Option><br>
		  <option value ='AMW' <%IF SptsDiv="AMW" THEN response.write(" selected")%>>AWSA Midwest Region</Option><br>
		  <option value ='ASC' <%IF SptsDiv="ASC" THEN response.write(" selected")%>>AWSA S Central Region</Option><br>
		  <option value ='ASO' <%IF SptsDiv="ASO" THEN response.write(" selected")%>>AWSA Southern Region</Option><br>
		  <option value ='AWE' <%IF SptsDiv="AWE" THEN response.write(" selected")%>>AWSA Western Region</Option><br>
		  <option value ='NAR' <%IF SptsDiv="NAR" THEN response.write(" selected")%>>NCWSA All Regions</Option><br>
		  <option value ='NEA' <%IF SptsDiv="NEA" THEN response.write(" selected")%>>NCWSA Eastern Region</Option><br>
		  <option value ='NMW' <%IF SptsDiv="NMW" THEN response.write(" selected")%>>NCWSA Midwest Region</Option><br>
		  <option value ='NSC' <%IF SptsDiv="NSC" THEN response.write(" selected")%>>NCWSA S Central Region</Option><br>
		  <option value ='NWE' <%IF SptsDiv="NWE" THEN response.write(" selected")%>>NCWSA Western Region</Option><br>
		  <option value ='BAR' <%IF SptsDiv="BAR" THEN response.write(" selected")%>>ABC All Regions</Option><br>
		  <option value ='BEA' <%IF SptsDiv="BEA" THEN response.write(" selected")%>>ABC Eastern Region</Option><br>
		  <option value ='BMW' <%IF SptsDiv="BMW" THEN response.write(" selected")%>>ABC Midwest Region</Option><br>
		  <option value ='BSC' <%IF SptsDiv="BSC" THEN response.write(" selected")%>>ABC S Central Region</Option><br>
		  <option value ='BSO' <%IF SptsDiv="BSO" THEN response.write(" selected")%>>ABC Southern Region</Option><br>
		  <option value ='BWE' <%IF SptsDiv="BWE" THEN response.write(" selected")%>>ABC Western Region</Option><br>
		  <option value ='KAR' <%IF SptsDiv="KAR" THEN response.write(" selected")%>>AKA All Kneeboard</Option><br>
		  <option value ='WAR' <%IF SptsDiv="WAR" THEN response.write(" selected")%>>USW All Wakeboard</Option><br>

		  <option value ='GRA' <%IF SptsDiv="GRA" THEN response.write(" selected")%>>GrassRoots All Standalone</Option><br>

		  <option value ='GRE' <%IF SptsDiv="GRE" THEN response.write(" selected")%>>GrsRts StdAln AWSA East</Option><br>
		  <option value ='GRM' <%IF SptsDiv="GRM" THEN response.write(" selected")%>>GrsRts StdAln AWSA Midwst</Option><br>
		  <option value ='GRC' <%IF SptsDiv="GRC" THEN response.write(" selected")%>>GrsRts StdAln AWSA S Ctrl</Option><br>
		  <option value ='GRS' <%IF SptsDiv="GRS" THEN response.write(" selected")%>>GrsRts StdAln AWSA South</Option><br>
		  <option value ='GRW' <%IF SptsDiv="GRW" THEN response.write(" selected")%>>GrSRts StdAln AWSA West</Option><br>

		  <option value ='GRB' <%IF SptsDiv="GRB" THEN response.write(" selected")%>>GrsRts StdAln Barefoot</Option><br>
		  <option value ='GRY' <%IF SptsDiv="GRY" THEN response.write(" selected")%>>GrsRts StdAln Hydrofoil</Option><br>
		  <option value ='GRK' <%IF SptsDiv="GRK" THEN response.write(" selected")%>>GrsRts StdAln Kneeboard</Option><br>
		  <option value ='GRX' <%IF SptsDiv="GRX" THEN response.write(" selected")%>>GrsRts StdAln Wakeboard</Option><br>
		  
		  <option value ='All' <%IF SptsDiv="All" THEN response.write(" selected")%>>All Spts Divs All Regions</Option><br>
		</select><BR>&nbsp;<BR>
	
	Class(es):&nbsp;
		<SELECT name='Classes'>
		  <option value ='All' <%IF Classes="All" THEN response.write(" selected")%>>All Classes</Option><br>
		  <option value ='LR'  <%IF Classes="LR"  THEN response.write(" selected")%>>L or R</Option><br>
		  <option value ='LRP' <%IF Classes="LRP" THEN response.write(" selected")%>>L or R or P</Option><br>
		  <option value ='PP'  <%IF Classes="PP"  THEN response.write(" selected")%>>P Only</Option><br>
		  <option value ='ELR' <%IF Classes="ELR" THEN response.write(" selected")%>>E or L or R</Option><br>
		  <option value ='CR'  <%IF Classes="CR"  THEN response.write(" selected")%>>C or Higher</Option><br>
		  <option value ='C'   <%IF Classes="C"   THEN response.write(" selected")%>>C Only</Option><br>
		  <option value ='BC'  <%IF Classes="BC"  THEN response.write(" selected")%>>Below C Only</Option><br>
		</select><BR>&nbsp;<BR>

	&nbsp; Hide Archived Events:&nbsp;&nbsp;
		<input type="checkbox" name="HideArchive" value="Chk" <%IF HideArchive="Chk" THEN response.write(" checked")%>>
	<br>&nbsp;<BR>
			
		<INPUT type="Submit" value="List Tournaments"></FORM></FONT></center></TD>
	
	<TD><TABLE class="innertable" align=center>
		<tr><td>&nbsp;<img src="/rankings/images/buttons/sanctionok.gif"> </td><td><font size="2">&nbsp; Tournament Sanction OK</font></td></tr>
		<tr><td>&nbsp;<img src="/rankings/images/buttons/cancelled.gif"> </td><td><font size="2">&nbsp; Tournament Cancelled</font></td></tr>
		<tr><td>&nbsp;<img src="/rankings/images/buttons/scored.gif"> </td><td><font size="2">&nbsp; Some Documents Received <BR>
		<tr><td>&nbsp;<img src="/rankings/images/buttons/archived.gif"> </td><td><font size="2">&nbsp; Tournament Complete / Archived <BR>
		<tr><td>&nbsp;<img src="/rankings/images/buttons/questionred.gif"> </td><td><font size="2">&nbsp; Document not Received</font></td></tr>
		<tr><td>&nbsp;<img src="/rankings/images/buttons/lightning.gif"> </td><td><font size="2">&nbsp; Document Posted Electronically</font></td></tr>
		<tr><td>&nbsp;<img src="/rankings/images/buttons/smile19.gif"> </td><td><font size="2">&nbsp; Document Received Manually</font></td></tr>
		<tr><td>&nbsp;<img src="/rankings/images/buttons/notreqd.gif"> </td><td><font size="2">&nbsp; Document not Required</font></td></tr>
	</TABLE></TD>
	
	</TR>

	</TABLE>

<%

END SUB


'	------------------
SUB	ListTournaments
'	------------------

'	This subroutine builds a report page, with each qualifying tournament
'	presented on a single line, with status and action link(s).

sSQL = "Select ST.TournAppID, ST.TName, ST.TDateE, ST.TCity, ST.TState,"
sSQL = sSQL & " ST.TSite, ST.TSiteID, ST.TSponsor,"
sSQL = sSQL & " ST.TStatus, ST.TSanType, CASE WHEN left(ST.TSanction,6)<>ST.TournAppID"
sSQL = sSQL & " THEN ST.TournAppID+'?' ELSE ST.TSanction end as TSanction,"
sSQL = sSQL & " PT.PTF_SBK, PT.PTF_WSP, PT.PTF_TS, PT.PTF_OD, PT.PTF_BT, PT.PTF_JT,"
sSQL = sSQL & " PT.PTF_CS, PT.PTF_CJ, PT.PTF_SD, PT.PTF_TU, PT.PTF_HD, PT.PTF_TNY"
sSQL = sSQL & " FROM " & SanctionTableName & " ST LEFT JOIN "

sSQL = sSQL & PostTourTableName & " PT on PT.TournAppID = ST.TournAppID "

sSQL = sSQL & " WHERE ST.TStatus > 1 and ST.Deleted = 0 and ST.TYear = '" & CalYear & "'" 

IF HideArchive = "Chk" THEN
		sSQL = sSQL & " AND ST.TStatus NOT IN (3,5)"
END IF

SELECT CASE SptsDiv
	CASE "AAR"
		sSQL = sSQL & " AND ST.TSanType = 0"
	CASE "AEA"
		sSQL = sSQL & " AND ST.TSanType = 0 AND ST.TRegion = 'E'"
	CASE "AMW"
		sSQL = sSQL & " AND ST.TSanType = 0 AND ST.TRegion = 'M'"
	CASE "ASC"
		sSQL = sSQL & " AND ST.TSanType = 0 AND ST.TRegion = 'C'"
	CASE "ASO"
		sSQL = sSQL & " AND ST.TSanType = 0 AND ST.TRegion = 'S'"
	CASE "AWE"
		sSQL = sSQL & " AND ST.TSanType = 0 AND ST.TRegion = 'W'"
	CASE "NAR"
		sSQL = sSQL & " AND ST.TSanType = 1"
	CASE "NEA"
		sSQL = sSQL & " AND ST.TSanType = 1 AND ST.TRegion = 'E'"
	CASE "NMW"
		sSQL = sSQL & " AND ST.TSanType = 1 AND ST.TRegion = 'M'"
	CASE "NSC"
		sSQL = sSQL & " AND ST.TSanType = 1 AND ST.TRegion = 'C'"
	CASE "NWE"
		sSQL = sSQL & " AND ST.TSanType = 1 AND ST.TRegion = 'W'"
	CASE "BAR"
		sSQL = sSQL & " AND ST.TSanType = 3"
	CASE "BEA"
		sSQL = sSQL & " AND ST.TSanType = 3 AND ST.TRegion = 'E'"
	CASE "BMW"
		sSQL = sSQL & " AND ST.TSanType = 3 AND ST.TRegion = 'M'"
	CASE "BSC"
		sSQL = sSQL & " AND ST.TSanType = 3 AND ST.TRegion = 'C'"
	CASE "BSO"
		sSQL = sSQL & " AND ST.TSanType = 3 AND ST.TRegion = 'S'"
	CASE "BWE"
		sSQL = sSQL & " AND ST.TSanType = 3 AND ST.TRegion = 'W'"
	CASE "KAR"
		sSQL = sSQL & " AND ST.TSanType = 5"
	CASE "WAR"
		sSQL = sSQL & " AND ST.TSanType = 4"

	CASE "GRA"
		sSQL = sSQL & " AND ST.TSanType = 6"

	CASE "GRE"
		sSQL = sSQL & " AND ST.TSanType = 6 AND substring(ST.TSanction,3,1) = 'E'"
	CASE "GRM"
		sSQL = sSQL & " AND ST.TSanType = 6 AND substring(ST.TSanction,3,1) = 'M'"
	CASE "GRC"
		sSQL = sSQL & " AND ST.TSanType = 6 AND substring(ST.TSanction,3,1) = 'C'"
	CASE "GRS"
		sSQL = sSQL & " AND ST.TSanType = 6 AND substring(ST.TSanction,3,1) = 'S'"
	CASE "GRW"
		sSQL = sSQL & " AND ST.TSanType = 6 AND substring(ST.TSanction,3,1) = 'W'"


	CASE "GRB"
		sSQL = sSQL & " AND ST.TSanType = 6 AND substring(ST.TSanction,3,1) = 'B'"
	CASE "GRY"
		sSQL = sSQL & " AND ST.TSanType = 6 AND substring(ST.TSanction,3,1) = 'Y'"
	CASE "GRK"
		sSQL = sSQL & " AND ST.TSanType = 6 AND substring(ST.TSanction,3,1) = 'K'"
	CASE "GRX"
		sSQL = sSQL & " AND ST.TSanType = 6 AND substring(ST.TSanction,3,1) = 'X'"


END SELECT

SELECT CASE Classes
	CASE "LR"
		sSQL = sSQL & " AND substring(ST.TSanction,7,1) in ('L','R','A','B')"
	CASE "LRP"
		sSQL = sSQL & " AND substring(ST.TSanction,7,1) in ('L','R','A','B','P')"
	CASE "PP"
		sSQL = sSQL & " AND substring(ST.TSanction,7,1) = 'P'"
	CASE "ELR"
		sSQL = sSQL & " AND substring(ST.TSanction,7,1) in ('E','L','R','A','B','P')"
	CASE "CR"
		sSQL = sSQL & " AND substring(ST.TSanction,7,1) in ('C','E','L','R','A','B','P')"
	CASE "C"
		sSQL = sSQL & " AND substring(ST.TSanction,7,1) = 'C'"
	CASE "BC"
		sSQL = sSQL & " AND substring(ST.TSanction,7,1) not in ('C','E','L','R','A','B','P')"
END SELECT 

sSQL = sSQL & " order by ST.TDateE, ST.TournAppID"

'	WriteDebugSQL(sSQL)

objRS.open sSQL, sConnectionToTRATable, 3, 3

'	Finally we process the resulting record-set.

IF NOT objRS.eof THEN
	
	' First we format Column Headings for the Tournament Status Report
	
	%>

	<Table class="innertable" width=90% align=center>

	<TR>
		<TH Colspan=4><center><font size="2" color="#FFFFFF"><b>Tournament Identifiers</b></font></center></TH>
		<TH Colspan=12><center><font size="2" color="#FFFFFF"><b>Post-Tournament Document Status</b></font></center></TH>
	</TR>

	<tr>
		<th><center><font size="2" color="#FFFFFF"> <a title="Tournament Date Month-Day">MM-DD</a> </font></center></th>
		<th><center><font size="2" color="#FFFFFF"> <a title="Tournament Sanction ID">TourID</a> </font></center></th>
		<th><center><font size="2" color="#FFFFFF"> <a title="State in which held and Sponsor">State</a> </font></center></th>
		<th><center><font size="2" color="#FFFFFF"> <a title="Overall Tournament Status">Status</a> </font></center></th>
		<th><center><font size="2" color="#FFFFFF"> <a title="Full Scorebook Status">SBK</a> </font></center></th>
		<th><center><font size="2" color="#FFFFFF"> <a title="Officials Credits Status">OD</a> </font></center></th>
		<th><center><font size="2" color="#FFFFFF"> <a title="Safety Report Status">SD</a> </font></center></th>
		<th><center><font size="2" color="#FFFFFF"> <a title="Chief Judges Report Status">CJ</a> </font></center></th>
		<th><center><font size="2" color="#FFFFFF"> <a title="Towboat Utilization Report Status">TU</a> </font></center></th>
		<th><center><font size="2" color="#FFFFFF"> <a title="Homologation Dossier Status">HD</a> </font></center></th>
		<th><center><font size="2" color="#FFFFFF"> <a title="Condensed Scorebook Status">CS</a> </font></center></th>
		<th><center><font size="2" color="#FFFFFF"> <a title="Rankings Data File Status">WSP</a> </font></center></th>
		<th><center><font size="2" color="#FFFFFF"> <a title="Tournament Summary Report Status">TS</a> </font></center></th>
		<th><center><font size="2" color="#FFFFFF"> <a title="Tournament Settings Status">TNY</a> </font></center></th>
		<th><center><font size="2" color="#FFFFFF"> <a title="Boat Time Tracking Report Status">BT</a> </font></center></th>
		<th><center><font size="2" color="#FFFFFF"> <a title="Jump Time Report Status">JT</a> </font></center></th>
	</tr>
	
	<%

	DO Until objRS.eof

		TournAppID = objRS("TournAppID")
		TDateE = Replace(FormatDateTime(objRS("TDateE"),2),"/","-")
		IF Mid(TDateE,2,1) = "-" THEN TDateE = "0" & TDateE
		IF Mid(TDateE,5,1) = "-" THEN TDateE = Left(TDateE,3) & "0" & Right(TDateE,6)
		MonthDay = Left(TDateE,5)
		TStatus = objRS("TStatus")

		%>
		
		<tr>
			<td><center><font size="2"><%=MonthDay%></font></center></td>
			<td><center><font size="2"><a href="/rankings/View-TournamentsHQ.asp?pvar=TourInfo&TourID=<%=TournAppID%>"
				title="<%=objRS("TName")%> --&#13;Click here to view Announcement&#13;<%=objRS("TSiteID")%>&nbsp;&nbsp; <%=objRS("TSite")%>" Target="_blank" STYLE="text-decoration:none">
				<%=objRS("TSanction")%> </a></font></center></td>
			<td><center><font size="2"><a title="<%=objRS("TSponsor")%>&#13;Held in <%=objRS("TCity")%>, <%=objRS("TState")%>"><%=objRS("TState")%></a></font></center></td>

		<%
		
		SELECT CASE TStatus
			CASE 2
				IconName = "SanctionOK": TourStat = "Tournament Sanction OK"
			CASE 3
				IconName = "Cancelled": TourStat = "Tournament Cancelled"
			CASE 4
				IconName = "Scored": TourStat = "Some Documents Received"
			CASE 5
				IconName = "Archived": TourStat = "Tournament Complete / Archived"
		END SELECT
		
		IF Session("adminmenulevel") >= 10 THEN %>
		
			<FORM method="post" action="<%=ThisModule%>">
				<INPUT type="hidden" name="Process" value="Editor">
				<INPUT type="hidden" name="CalYear" value="<%=CalYear%>">
				<INPUT type="hidden" name="SptsDiv" value="<%=SptsDiv%>">
				<INPUT type="hidden" name="Classes" value="<%=Classes%>">
				<INPUT type="hidden" name="HideArchive" value="<%=HideArchive%>">
				<INPUT type="hidden" name="TourID" value="<%=TournAppID%>">
			<td><center><a title="<%=TourStat%> --&#13;Click here to update Tournament"><input type="image" 
				src="/rankings/images/buttons/<%=IconName%>.gif" Name="Editor"></a></center></td></FORM>

		<% ELSE %>

			<td><center><img src="/rankings/images/buttons/<%=IconName%>.gif" 
				Title="<%=TourStat%>"></center></td>

		<% END IF 
		
		FormatReportStatus objRS("PTF_SBK"), "Full Scorebook"
		FormatReportStatus objRS("PTF_OD"), "Officials Credits"
		FormatReportStatus objRS("PTF_SD"), "Safety Report"
		FormatReportStatus objRS("PTF_CJ"), "Chief Judges Report"
		FormatReportStatus objRS("PTF_TU"), "Towboat Utilization Report"
		FormatReportStatus objRS("PTF_HD"), "Homologation Dossier"
		FormatReportStatus objRS("PTF_CS"), "Condensed Scorebook"
		FormatReportStatus objRS("PTF_WSP"), "Rankings Data File"
		FormatReportStatus objRS("PTF_TS"), "Tournament Summary Report"
		FormatReportStatus objRS("PTF_TNY"), "Tournament Settings"
		FormatReportStatus objRS("PTF_BT"), "Boat Time Tracking Report"
		FormatReportStatus objRS("PTF_JT"), "Jump Time Report"

		%></tr><%
		
		objRS.MoveNext

	LOOP
		
ELSE

END IF

objRS.close
set objRS=nothing

%></table><br><%

WriteIndexPageFooter

END SUB


'	------------------
SUB	FormatReportStatus (Status, ReptDesc)
'	------------------

Dim IconName, How, TourStat

'	This subroutine formats the icon and hidden title for a single report flag

SELECT CASE Status
	CASE 0
		IconName = "questionred": How = " not received"
	CASE 1
		IconName = "lightning": How = " posted electronically"
	CASE 2
		IconName = "smile19": How = " received manually"
	CASE 3
		IconName = "notreqd": How = " not required"
	CASE ELSE
		IconName = "missing": How = " status unknown"
END SELECT

IF Status = 1 and ReptDesc = "Condensed Scorebook" THEN %>

	<td><center><a href="/rankings/Scorebks/<%=TournAppID%>CS.HTM" Target="_blank"
		title="<%=ReptDesc&How%>&#13;Click here to view Scorebook"><img src="/rankings/images/buttons/<%=IconName%>.gif" STYLE="border-style:none"></a></center></td>

<% ELSE %>

	<td><center><a title="<%=ReptDesc&How%>"><img src="/rankings/images/buttons/<%=IconName%>.gif"></center></td>

<% END IF 

END SUB


'	----------------
SUB	EditTournament
'	----------------

TournAppID = request("TourID")

sSQL = "Select ST.TSanction, ST.TName, ST.TDateE, ST.TCity, ST.TState,"
sSQL = sSQL & " ST.TEventSlalom, ST.TEventTrick, ST.TEventJump, ST.TStatus, ST.TSanType,"
sSQL = sSQL & " RT.SClassN, RT.SClassC, RT.SClassE, RT.SClassL, RT.SClassR, RT.SClassCash,"
sSQL = sSQL & " RT.TClassN, RT.TClassC, RT.TClassE, RT.TClassL, RT.TClassR, RT.TClassCash,"
sSQL = sSQL & " RT.JClassN, RT.JClassC, RT.JClassE, RT.JClassL, RT.JClassR, RT.JClassCash,"
sSQL = sSQL & " RT.USClassN, RT.USClassC, RT.UTClassN, RT.UTClassC, RT.UJClassN, RT.UJClassC,"
sSQL = sSQL & " Coalesce(PT.PTF_SBK,-1) as PTF_SBK, PT.PTF_WSP, PT.PTF_TS, PT.PTF_OD, PT.PTF_BT,"
sSQL = sSQL & " PT.PTF_JT, PT.PTF_CS, PT.PTF_CJ, PT.PTF_SD, PT.PTF_TU, PT.PTF_HD, PT.PTF_TNY"
sSQL = sSQL & " FROM " & SanctionTableName & " ST LEFT JOIN " & TRegSetupTableName
sSQL = sSQL & " RT on RT.TournAppID = ST.TournAppID LEFT JOIN " & PostTourTableName
sSQL = sSQL & " PT on PT.TournAppID = ST.TournAppID where upper(ST.TournAppID) = '"
sSQL = sSQL & TournAppID & "'"

objRS.open sSQL, sConnectionToSanctionTable, 3, 3

TStatus = objRS("TStatus")
TSanction = objRS("TSanction")
TSanType = objRS("TSanType")
TName = objRS("TName")
TDateE = Replace(FormatDateTime(objRS("TDateE"),2),"/","-")
IF Mid(TDateE,2,1) = "-" THEN TDateE = "0" & TDateE
IF Mid(TDateE,5,1) = "-" THEN TDateE = Left(TDateE,3) & "0" & Right(TDateE,6)

IF objRS("TEventSlalom") = True or objRS("SClassC") > 0 or objRS("SClassE") > 0 or objRS("SClassL") > 0 or objRS("SClassR") > 0 or objRS("USClassC") > 0 THEN TEventSlalom = True: else TEventSlalom = False
IF objRS("TEventTrick") = True or objRS("TClassC") > 0 or objRS("TClassE") > 0 or objRS("TClassL") > 0 or objRS("TClassR") > 0 or objRS("UTClassC") > 0 THEN TEventTrick = True: else TEventTrick = False
IF objRS("TEventJump") = True or objRS("JClassC") > 0 or objRS("JClassE") > 0 or objRS("JClassL") > 0 or objRS("JClassR") > 0 or objRS("UJClassC") > 0 THEN TEventJump = True: else TEventJump = False

IF objRS("PTF_SBK") > -1 THEN

	'	Pick up Existing PTF Flag Values from PostTourn Table Entry
		
	PTF_SBK = objRS("PTF_SBK")
	PTF_WSP = objRS("PTF_WSP")
	PTF_TS = objRS("PTF_TS")
	PTF_OD = objRS("PTF_OD")
	PTF_BT = objRS("PTF_BT")
	PTF_JT = objRS("PTF_JT")
	PTF_CS = objRS("PTF_CS")
	PTF_CJ = objRS("PTF_CJ")
	PTF_SD = objRS("PTF_SD")
	PTF_TU = objRS("PTF_TU")
	PTF_HD = objRS("PTF_HD")
	PTF_TNY = objRS("PTF_TNY")

ELSE

	'	Insert new row into PostTourn table if not already there, then
	'	Supply default zero values for all flags.  Any 3's will come below.
		
	sSQL = "Insert Into " & PostTourTableName & " Values ('" & TournAppID 
	sSQL = sSQL & "', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)"
	OpenConSanUpd
	ConSanUpd.Execute(sSQL)
	CloseConSanUpd
	
	PTF_SBK = 0: PTF_WSP = 0: PTF_TS = 0: PTF_OD = 0: PTF_BT = 0: PTF_JT = 0
	PTF_CS = 0: PTF_CJ = 0: PTF_SD = 0: PTF_TU = 0: PTF_HD = 0: PTF_TNY = 0
	TStatus = 2

END IF

'	Now Revise PTF_ flag status if incoming event TStatus = 2
'	Based on TSanType and Events Included.

IF TStatus = 2 THEN
	
	IF TSanType = 0 or TSanType = 1 THEN
		IF TEventSlalom = False and TEventJump = False Then PTF_BT = 3
		IF TEventJump = False Then PTF_JT = 3
		IF instr("ELRPAB",mid(TSanction,7,1)) = 0 THEN PTF_HD = 3
	ELSEIF TSanType = 3 THEN
		PTF_HD = 3: PTF_CS = 3: 
		PTF_WSP = 3: PTF_TS = 3: PTF_TNY = 3: PTF_BT = 3: PTF_JT = 3
	ELSE 
		PTF_CJ = 3: PTF_TU = 3: PTF_HD = 3: PTF_CS = 3
		PTF_WSP = 3: PTF_TS = 3: PTF_TNY = 3: PTF_BT = 3: PTF_JT = 3
	END IF

END IF

' Now list the tournament specifics along with applicable items.

WriteIndexPageHeader

%>
	
	<Table class="innertable" width=80% align=center>
		
	<TR><TD Colspan=2><center><b><%=objRS("TSanction")%>&nbsp;&nbsp;&nbsp; <%=objRS("TName")%></b>
		<br><%=TDateE%>&nbsp;&nbsp;&nbsp; <%=objRS("TCity")%>,&nbsp; <%=objRS("TState")%>
	</TD></TR>
	
	<TR>
	<TH><center><b><font size="2" color="#FFFFFF">&nbsp; Status or Action&nbsp; </font></b></center></TH>
	<TH><center><b><font size="2" color="#FFFFFF">&nbsp; Document Description or Explanation&nbsp; </font></b></center></TH>
	</TR>

	<FORM method="post" action="<%=ThisModule%>">
		<INPUT type="hidden" name="Process" value="Listem">
		<INPUT type="hidden" name="CalYear" value="<%=CalYear%>">
		<INPUT type="hidden" name="SptsDiv" value="<%=SptsDiv%>">
		<INPUT type="hidden" name="Classes" value="<%=Classes%>">
		<INPUT type="hidden" name="HideArchive" value="<%=HideArchive%>">
		<TR><TD><center><INPUT type="Submit" value="No Action"></center></TD>
		<TD><center><font size="2">Return to Tournament Status Listing</font></center></TD></TR>
	</FORM>

<% IF TStatus = 3 THEN %>

	<FORM method="post" action="<%=ThisModule%>">
		<INPUT type="hidden" name="Process" value="Reinstate">
		<INPUT type="hidden" name="CalYear" value="<%=CalYear%>">
		<INPUT type="hidden" name="SptsDiv" value="<%=SptsDiv%>">
		<INPUT type="hidden" name="Classes" value="<%=Classes%>">
		<INPUT type="hidden" name="HideArchive" value="<%=HideArchive%>">
		<INPUT type="hidden" name="TourID" value="<%=TournAppID%>">
		<TR><TD><center><INPUT type="Submit" value="Reinstate"></center></TD>
		<TD><center><font size="2">Reinstate (Un-Cancel) this Tournament</font></center></TD></TR>
	</FORM>
	
<% ELSE %>

	<%IF TStatus = 2 THEN %>

		<FORM method="post" action="<%=ThisModule%>">
			<INPUT type="hidden" name="Process" value="Cancel">
			<INPUT type="hidden" name="CalYear" value="<%=CalYear%>">
			<INPUT type="hidden" name="SptsDiv" value="<%=SptsDiv%>">
			<INPUT type="hidden" name="Classes" value="<%=Classes%>">
			<INPUT type="hidden" name="HideArchive" value="<%=HideArchive%>">
			<INPUT type="hidden" name="TourID" value="<%=TournAppID%>">
			<TR><TD><center><INPUT type="Submit" value="Cancel"></center></TD>
			<TD><center><font size="2">Cancel this Tournament</font></center></TD></TR>
		</FORM>

	<% END IF %>
	
	<FORM method="post" action="<%=ThisModule%>">
	<INPUT type="hidden" name="Process" value="Update">
	<INPUT type="hidden" name="CalYear" value="<%=CalYear%>">
	<INPUT type="hidden" name="SptsDiv" value="<%=SptsDiv%>">
	<INPUT type="hidden" name="Classes" value="<%=Classes%>">
	<INPUT type="hidden" name="HideArchive" value="<%=HideArchive%>">
	<INPUT type="hidden" name="TourID" value="<%=TournAppID%>">
	
	<%
	
	Updatable = "N"

	PresentEditOptions "PTF_SBK", PTF_SBK, "Full Scorebook"
	PresentEditOptions "PTF_OD", PTF_OD, "Officials Credits"
	PresentEditOptions "PTF_SD", PTF_SD, "Safety Report"
	PresentEditOptions "PTF_CJ", PTF_CJ, "Chief Judges Report"
	PresentEditOptions "PTF_TU", PTF_TU, "Towboat Utilization Report"
	PresentEditOptions "PTF_HD", PTF_HD, "Homologation Dossier"
	PresentEditOptions "PTF_CS", PTF_CS, "Condensed Scorebook"
	PresentEditOptions "PTF_WSP", PTF_WSP, "Rankings Data File"
	PresentEditOptions "PTF_TS", PTF_TS, "Tournament Summary Report"
	PresentEditOptions "PTF_TNY", PTF_TNY, "Tournament Settings"
	PresentEditOptions "PTF_BT", PTF_BT, "Boat Time Tracking Report"
	PresentEditOptions "PTF_JT", PTF_JT, "Jump Time Report"
	
	IF Updatable = "Y" THEN %>

		<TR><TD><center><INPUT type="Submit" value="Update"></center></TD>
		<TD><center><font size="2">Update above Manual Document Status</font></center>
		</TD></TR>
		
		<TR><TD><center><input type="checkbox" name="SendEmail" value="Chk" checked></center></TD>
		<TD><center><font size="2">eMail Confirmation to Tour Dir &amp; Chf Ofcls</font></center>
		</TD></TR>
		
		<TR><TD><center><font size="2">Comment:</font></center></TD>
		<TD><center><INPUT type="text" name="Comment" size="50" maxlength="72"></center>
		</TD></TR>
		
	<% END IF %>	
		
	</FORM>

<% END IF %>

	</Table>
	
<%
	
WriteIndexPageFooter

END SUB


'	--------------------
SUB	PresentEditOptions (FlagID, Status, ReptDesc)
'	--------------------

'	This subroutine formats a line presenting the icon and description for a single document,
'	including a checkbox for Manual Receipt if Status is not 1 or 3.
'	Items in Status 3 do not even get listed.

IF Status < 3 THEN

	%><tr><td><center><%

	SELECT CASE Status
	
	CASE 0 
		Updatable = "Y"
		%>
			<a title="Not Received"><img src="/rankings/images/buttons/questionred.gif"></a>&nbsp;&nbsp;
			<a title="Check this box to POST Manually Received Status"><input type="checkbox" name="<%=FlagID%>" value="Chk"></a>
		<%
	
	CASE 1
	
		IF Session("adminmenulevel") < 30 THEN 
			%>
				<a title="Received Electronically"><img src="/rankings/images/buttons/lightning.gif"></a>
				<input type="hidden" name="<%=FlagID%>" value="Chk">
			<%
		ELSE
			Updatable = "Y"
			%>
				<a title="Received Electronically"><img src="/rankings/images/buttons/lightning.gif"></a>&nbsp;&nbsp;
				<a title="Un-Check this box to REMOVE Electronically Received Status"><input type="checkbox" name="<%=FlagID%>" value="Chk" checked></a>
			<%
		END IF
	
	CASE 2
		Updatable = "Y"
		%>
			<a title="Received Manually"><img src="/rankings/images/buttons/smile19.gif"></a>&nbsp;&nbsp;
			<a title="Un-Check this box to REMOVE Manually Received Status"><input type="checkbox" name="<%=FlagID%>" value="Chk" checked></a>
		<%

	END SELECT

	%>
		</td><td><center><font size="2"><%=ReptDesc%> Status</center></font></td></tr>
	<%

END IF

%><input type="hidden" name="Old_<%=FlagID%>" value="<%=Status%>"><%

END SUB



'	-------------------
SUB	UpdateTournament
'	-------------------	

'	Now finally update the Document Status, as specified on the Edit Form.
'	First we review status of all Documents, building an SQL Update query,
'	Along with a list of Updates and also remaining Missing Documents
'	to be cited in an email below.

TournAppID = Request("TourID")

SetList = "": Updated = "": Missing = "": Updatable = "N"

UpdateDocStatus "PTF_SBK", "Full Scorebook"
UpdateDocStatus "PTF_OD", "Officials Credits"
UpdateDocStatus "PTF_SD", "Safety Report"
UpdateDocStatus "PTF_CJ", "Chief Judges Report"
UpdateDocStatus "PTF_TU", "Towboat Utilization Report"
UpdateDocStatus "PTF_HD", "Homologation Dossier"
UpdateDocStatus "PTF_CS", "Condensed Scorebook"
UpdateDocStatus "PTF_WSP", "Rankings Data File"
UpdateDocStatus "PTF_TS", "Tournament Summary Report"
UpdateDocStatus "PTF_TNY", "Tournament Settings"
UpdateDocStatus "PTF_BT", "Boat Time Tracking Report"
UpdateDocStatus "PTF_JT", "Jump Time Report"

'	Now Update the Post-Tournament Document Status table

sSQL = "Update " & PostTourTableName & Setlist & " Where TournAppID = '" & TournAppID & "'"
'	WriteDebugSQL sSQL
OpenConSanUpd
ConSanUpd.Execute(sSQL)
CloseConSanUpd

'	Now Update the Overall Tournament Status -- Depends on what we've done.

IF len(Missing) = 0 THEN
	TStatus = "5"
ELSEIF Updatable = "Y" THEN
	TStatus = "4"
ELSE
	TStatus = "2"
END IF

sSQL = "Update " & SanctionTableName & " Set TStatus = " & TStatus & " Where TournAppID = '" & TournAppID & "'"
'	WriteDebugSQL sSQL
OpenConSanUpd
ConSanUpd.Execute(sSQL)
CloseConSanUpd

'	Now get Tournament Details and Email addresses and construct a confirmation email.

sSQL = "Select ST.TSanction, ST.TName, ST.TDateE, ST.TCity, ST.TState,"
sSQL = sSQL & " ST.TSanType, ST.TStatus, ST.TDirName, ST.TDirEMail,"
sSQL = sSQL & " CJ.CJudgName, CJ.CJudgEMail, CC.CScorName, CC.CScorEMail,"
sSQL = sSQL & " PT.PTF_SBK, PT.PTF_WSP, PT.PTF_TS, PT.PTF_OD, PT.PTF_BT, PT.PTF_JT,"
sSQL = sSQL & " PT.PTF_CS, PT.PTF_CJ, PT.PTF_SD, PT.PTF_TU, PT.PTF_HD, PT.PTF_TNY"
sSQL = sSQL & " FROM " & SanctionTableName & " ST LEFT JOIN " & PostTourTableName
sSQL = sSQL & " PT on PT.TournAppID = ST.TournAppID"

sSQL = sSQL & " LEFT JOIN (Select SX.TournAppID, MT.FirstName + ' ' + MT.LastName"
sSQL = sSQL & " as CJudgName, MT.Email as CJudgEMail FROM " & MemberTableName
sSQL = sSQL & " MT JOIN (Select TournAppID, Cast(case when len(CJudgePID)<9 then"
sSQL = sSQL & " CJudgePID else right(CJudgePID,8) end as integer) as PID FROM "
sSQL = sSQL & TRegSetupTableName & " WHERE isnumeric(CJudgePID) = 1)"
sSQL = sSQL & " SX on SX.PID = MT.PersonID WHERE patindex('%@%',Email) > 0 )"
sSQL = sSQL & " CJ ON CJ.TournAppID = ST.TournAppID"

sSQL = sSQL & " LEFT JOIN (Select SX.TournAppID, MT.FirstName + ' ' + MT.LastName"
sSQL = sSQL & " as CScorName, MT.Email as CScorEMail FROM " & MemberTableName
sSQL = sSQL & " MT JOIN (Select TournAppID, Cast(case when len(CScorePID)<9 then"
sSQL = sSQL & " CScorePID else right(CScorePID,8) end as integer) as PID FROM "
sSQL = sSQL & TRegSetupTableName & " WHERE isnumeric(CScorePID) = 1)"
sSQL = sSQL & " SX on SX.PID = MT.PersonID WHERE patindex('%@%',Email) > 0 )"
sSQL = sSQL & " CC ON CC.TournAppID = ST.TournAppID"

sSQL = sSQL & " where upper(ST.TournAppID) = '" & TournAppID & "'"

objRS.open sSQL, sConnectionToSanctionTable, 3, 3

TStatus = objRS("TStatus")
TSanction = objRS("TSanction")
TName = objRS("TName")
TDateE = Replace(FormatDateTime(objRS("TDateE"),2),"/","-")
IF Mid(TDateE,2,1) = "-" THEN TDateE = "0" & TDateE
IF Mid(TDateE,5,1) = "-" THEN TDateE = Left(TDateE,3) & "0" & Right(TDateE,6)

PTF_SBK = objRS("PTF_SBK")
PTF_WSP = objRS("PTF_WSP")
PTF_TS = objRS("PTF_TS")
PTF_OD = objRS("PTF_OD")
PTF_BT = objRS("PTF_BT")
PTF_JT = objRS("PTF_JT")
PTF_CS = objRS("PTF_CS")
PTF_CJ = objRS("PTF_CJ")
PTF_SD = objRS("PTF_SD")
PTF_TU = objRS("PTF_TU")
PTF_HD = objRS("PTF_HD")
PTF_TNY = objRS("PTF_TNY")

'	Log the update activity and current updated status flags

WriteLog (date() & "  " & time() & "  Documents Posted to " & TSanction & " -- Status/Flags = " & TStatus & "/" & PTF_SBK & PTF_WSP & PTF_TS & PTF_OD & PTF_BT & PTF_JT & PTF_CS & PTF_CJ & PTF_SD & PTF_TU & PTF_HD & PTF_TNY)

WriteIndexPageHeader

'	Now we establish the primary email address string -- TD / CJ / CC

eMailTo = ""

IF len(objRS("TDirEMail")) > 0 THEN
	eMailTo = """" & objRS("TDirName") & """ <" & objRS("TDirEMail") & ">"
END IF

IF len(objRS("CJudgEmail")) > 0 and instr(eMailTo,objRS("CJudgName")) = 0 THEN
	IF len(eMailTo) > 0 THEN eMailTo = eMailTo & "; "
	eMailTo = eMailTo & """" & objRS("CJudgName") & """ <" & objRS("CJudgEmail") & ">"
END IF

IF len(objRS("CScorEmail")) > 0 and instr(eMailTo,objRS("CScorName")) = 0 THEN
	IF len(eMailTo) > 0 THEN eMailTo = eMailTo & "; "
	eMailTo = eMailTo & """" & objRS("CScorName") & """ <" & objRS("CScorEmail") & ">"
END IF

	%>
	
	<Table class="innertable" width=80% align=center>

	<TR><TD><center><b><%=objRS("TSanction")%>&nbsp;&nbsp;&nbsp; <%=objRS("TName")%></b>
		<br><%=TDateE%>&nbsp;&nbsp;&nbsp; <%=objRS("TCity")%>,&nbsp; <%=objRS("TState")%>
	</TD></TR>
	
	<TR><TD>Document Status Revised in this Update:
	<br><%=Replace(Replace(Updated,"     ","&nbsp;&nbsp;&nbsp;&nbsp; "),vbCrLf,"<br>")%></TD></TR>
	
	<TR><TD>Documents still Outstanding after this update:
	<br><%=Replace(Replace(Missing,"     ","&nbsp;&nbsp;&nbsp;&nbsp; "),vbCrLf,"<br>")%></TD></TR>
	
	<%
	
	IF Request("SendEmail") = "Chk" THEN

		IF len(eMailTo) = 0 THEN %>

			<TR><td>No eMail addresses found for Tour Dir or Chf Ofcls</td></TR>
	
		<% ELSEIF len(Updated) > 0 THEN
		
			' Create and Send Confirmation eMail to TD and Chiefs
			
			' First we Invoke "standard" Email Server Configuration -- defines objMessage object
			SetupEmailService
			objMessage.Subject = "Post-Tournament Reports from " & TSanction & " " & TName & " (" & TDateE & ")"
			objMessage.From = """USA Water Ski Competition"" <shardee@usawaterski.org>"

			'	Next we establish from and secondary addressing based on jurisdiction codes

			IF mid(TSanction,3,1) = "C" THEN
				eMailCC = """Danny LeBourgeois"" <dleboo@gmail.com>"
			ELSEIF mid(TSanction,3,1) = "M" THEN
		   	eMailCC = """Michael O'Conner"" <h2oskimo@gmail.com>"
			ELSEIF mid(TSanction,3,1) = "E" THEN
		   	eMailCC = """Jennifer Frederick-Kelly"" <fredmach@choiceonemail.com>"
			ELSEIF mid(TSanction,3,1) = "S" THEN
		   	eMailCC = """Kirby Whetsel"" <kwhetsel@charter.net>"
			ELSEIF mid(TSanction,3,1) = "W" THEN
		   	eMailCC = """Judy Stanford"" <judy-don@sbcglobal.net>"
			ELSEIF mid(TSanction,3,1) = "U" THEN
		   	eMailCC = """Robert Rhyne"" <rrriii@mindspring.com>"
			ELSEIF mid(TSanction,3,1) = "B" THEN
		   	eMailCC = ""
			ELSEIF mid(TSanction,3,1) = "K" THEN
		   	eMailCC = ""
			ELSEIF mid(TSanction,3,1) = "X" THEN
				eMailCC = ""
			ELSEIF mid(TSanction,3,1) = "Y" THEN
		   	eMailCC = ""
			END IF

			objMessage.To = eMailTo
			objMessage.CC = eMailCC

			IF instr(eMailCC,"Kirby Whetsel") = 0 THEN 
				objMessage.BCC = """Kirby Whetsel"" <kwhetsel@charter.net>"
			ELSE
				objMessage.BCC = ""
			END IF
			
			eMailBody = "Dear Tournament Organizer and/or Chief Official(s) --" & vbCrLf & vbCrLf

			eMailBody = eMailBody & "Post-tournament reports from " & TSanction & " " & TName & vbCrLf
			eMailBody = eMailBody & "ending " & TDateE
			eMailBody = eMailBody & " have been posted to the Sanction control system:" & vbCrLf & vbCrLf
			eMailBody = eMailBody & Updated & vbCrLf
			
			IF len(trim(Request("Comment"))) > 0 THEN
				eMailBody = eMailBody & trim(Request("Comment")) & vbCrLf & vbCrLf
			END IF
			
			IF len(Missing) = 0 THEN
				eMailBody = eMailBody & "All of the required post-tournament reports are now accounted for" & vbCrLf 
				eMailBody = eMailBody & "after storing these items.  These have been filed and checked off in" & vbCrLf
				eMailBody = eMailBody & "the Sanction Control System, and your Tournament marked complete." & vbCrLf & vbCrLf 
				eMailBody = eMailBody & "A big Thank You !!" & vbCrLf & vbCrLf
			ELSE
				eMailBody = eMailBody & "THE FOLLOWING ITEMS REMAIN OUTSTANDING FOR YOUR TOURNAMENT:" & vbCrLf & vbCrLf
				eMailBody = eMailBody & Missing & vbCrLf
				eMailBody = eMailBody & "Please ensure that the above-noted items are completed and submitted" & vbCrLf 
				eMailBody = eMailBody & "within 10 days.  Electronic submission through your regional seeding" & vbCrLf 
				eMailBody = eMailBody & "representative is preferred, although you may submit direct to USA" & vbCrLf 
				eMailBody = eMailBody & "Waterski HQ by postal mail instead." & vbCrLf & vbCrLf
				eMailBody = eMailBody & "If you have any questions about this subject, you can reply to this" & vbCrLf 
				eMailBody = eMailBody & "message, or call me at the number listed below." & vbCrLf & vbCrLf
			END IF
			

			eMailBody = eMailBody & "Sandy Hardee" & vbCrLf & "Competition Department" & vbCrLf
			eMailBody = eMailBody & "shardee@usawaterski.org" & vbCrLf & "1-863-324-4341 ext 126" & vbCrLf
			eMailBody = eMailBody & "Direct Line:  1-863-874-5681"

			' Now send the eMail message and clear that object.
	
			objMessage.TextBody = eMailBody
			objMessage.Send
			Set objMessage=nothing
		
			%>
	
			<td><p>Confirmation of these Status changes has been eMailed to:
				<br><%=Replace(Replace(eMailTo,"<","&lt;"),">","&gt;")%></p></td>

		<% END IF
	
	END IF %>

	<FORM method="post" action="<%=ThisModule%>">
		<INPUT type="hidden" name="Process" value="Listem">
		<INPUT type="hidden" name="CalYear" value="<%=CalYear%>">
		<INPUT type="hidden" name="SptsDiv" value="<%=SptsDiv%>">
		<INPUT type="hidden" name="Classes" value="<%=Classes%>">
		<INPUT type="hidden" name="HideArchive" value="<%=HideArchive%>">
		<TR><TD><center><INPUT type="Submit" value="Return to Tournament List"></center></TD></TR>
	</FORM>

	</Table>
	
	<%
	
WriteIndexPageFooter

END SUB


'	-----------------
SUB	UpdateDocStatus (FlagID, ReptDesc)
'	-----------------

'	This subroutine Digests the Updated Status of a particular Document
'	Resulting from the Tournament Document Status Update form.
'	Adds Set FlagID = Value to SQL Set Statement String, and
'	Adds Status Change Confirmations to Updates String, and
'	Adds Document Name to Missing List, if still missing.

IF len(SetList) = 0 THEN SetList = SetList & " SET ": ELSE SetList = SetList & ", "
SetList = SetList & FlagID & "="

IF Request("Old_"&FlagID) = "0" THEN
	IF Request(FlagID) = "Chk" THEN
		SetList = SetList & "2"
		Updated = Updated & "     " & ReptDesc & " -- Posted as Received Manually" & vbCrLf
		Updatable = "Y"
	ELSE
		SetList = Setlist & "0"
		Missing = Missing & "     " & ReptDesc & vbCrLf
	END IF

ELSEIF Request("Old_"&FlagID) = "1" THEN
	IF Request(FlagID) = "Chk" THEN
		SetList = SetList & "1"
		Updatable = "Y"
	ELSE
		SetList = Setlist & "0"
		Updated = Updated & "     " & ReptDesc & " -- Electronic Receipt Removed" & vbCrLf
		Missing = Missing & "     " & ReptDesc & vbCrLf
	END IF

ELSEIF Request("Old_"&FlagID) = "2" THEN
	IF Request(FlagID) = "Chk" THEN
		SetList = SetList & "2"
		Updatable = "Y"
	ELSE
		SetList = Setlist & "0"
		Updated = Updated & "     " & ReptDesc & " -- Manual Receipt Removed" & vbCrLf
		Missing = Missing & "     " & ReptDesc & vbCrLf
	END IF

ELSEIF Request("Old_"&FlagID) = "3" THEN
	SetList = SetList & "3"
	
END IF	

END SUB


'	-------------------
SUB	CancelTournament
'	-------------------

sSQL = "Update " & SanctionTableName & " Set TStatus = 3 Where TournAppID = '" & request("TourID") & "'"
'	WriteDebugSQL sSQL
OpenConSanUpd
ConSanUpd.Execute(sSQL)
CloseConSanUpd

END SUB


'	-------------------
SUB	UnCancelTournament
'	-------------------

sSQL = "Update " & SanctionTableName & " Set TStatus = 2 Where TournAppID = '" & request("TourID") & "'"
'	WriteDebugSQL sSQL
OpenConSanUpd
ConSanUpd.Execute(sSQL)
CloseConSanUpd

END SUB


%>