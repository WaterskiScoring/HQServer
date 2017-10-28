<!--#include virtual="epl/functions.asp" -->

<% 

Dim objConn, RS, HQConn, RSHQ
Set HQConn = Server.CreateObject("ADODB.Connection")
HQConn.Open Application("HQSQLConn")
Set RSHQ = Server.CreateObject("ADODB.RecordSet")
RSHQ.ActiveConnection = HQConn
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open Application("WaterSkiConn")
Set RS = Server.CreateObject("ADODB.RecordSet")
RS.ActiveConnection = objConn

Set objCDO = Server.CreateObject("CDO.Message")

Dim Usages(2,3,5), ErrStat, sSex
Dim SSkr, SWvr, SASlm, SATrk, SAJmp, SBSlm, SBTrk, SBJmp
Dim TSkr, TWvr, TASlm, TATrk, TAJmp, TBSlm, TBTrk, TBJmp

Function RemInvChr(strInput)
    dim workingstring
	On Error Resume Next
	For i = 1 to Len(strInput)
		If isNumeric(Mid(strInput, i, 1)) then
			workingstring = workingstring & Mid(strInput, i, 1)
		End If
		If (Mid(strInput, i, 1)) => "a" and (Mid(strInput, i, 1)) <=  "z" then
			workingstring = workingstring & Mid(strInput, i, 1)
		End If
		If (Mid(strInput, i, 1)) => "A" and (Mid(strInput, i, 1)) <=  "Z" then
			workingstring = workingstring & Mid(strInput, i, 1)
		End If
		If (Mid(strInput, i, 1)) = "@" Or (Mid(strInput, i, 1)) = "." Or (Mid(strInput, i, 1)) = " " Then
				workingstring = workingstring & Mid(strInput, i, 1)
		End If
	Next
	RemInvChr = workingstring
End Function


'	***** Bailout to Members Login if not auth or no Session("TeamId") value

IF not Session("auth") or Session("id") < 1 then response.redirect "https://www.usawaterski.org/members/login/index.asp"

'	*****	Store Selected Tournament to Session Variables, if re-entering at "GotTour" point

IF Request("FormStatus") = "GotTour" THEN 
	Session("TourID") = left(request("TourID"),6)
	Session("TourDate") = mid(request("TourID"),7,10)	
	Session("TourName") = mid(request("TourID"),17,len(request("TourID"))-16)
END IF

%>

<html>

<head>
<title>NCWSA Team Rotation Planning</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" >

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#42639F">
      <p>&nbsp;</p>
      <p align="center"><font face="Verdana" size="6" color="#FFFFFF">
      	<b>NCWSA Tournament Registration Report</b></font></p>
		<% IF Session("TourID") <> "" THEN %>
      <p align="center"><font face="Verdana" size="4" color="#FFFFFF">
      	<b><%=Session("TourID")%>&nbsp;&nbsp;&nbsp; <%=Session("TourName")%>&nbsp;&nbsp;&nbsp; <%=Session("TourDate")%></b></font></p>
      <% END IF %>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>

<table border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="180" valign="top" bgcolor="#42639F">

			<br>
	        &nbsp;<a href="https://www.usawaterski.org/members/"><font face="arial" size="2" COLOR="#FFFFFF">Member's Only Home</font></a><br>&nbsp;<br>
			<br>
            <font face="Verdana" size="1">&nbsp;<font COLOR="#FFFFFF">Powered by</font> <a href="http://www.epolk.com"><font COLOR="#FFFFFF">ePolk.com</font></a><br></font>
	
    </td>

<td width="720" >

<%

'	********** If Option not "GotTour" then present list of upcoming tournaments to pick from

IF Request("FormStatus") <> "GotTour" THEN 

	%>

  <table>
    <tr> 
      <td width="20">&nbsp;</td>
      <td width="560">&nbsp;</td>
      <td width="20">&nbsp;</td>
    </tr>
 
    <tr> 
      <td>&nbsp;</td>
      <td valign="top">
        <font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        Following is a list of upcoming NCWSA Tournaments.&nbsp; Please 
        select the tournament for which you want a Team Registration 
        report, from this drop-down list below.<br>&nbsp;
        </font></td>
      <td>&nbsp;</td>
    </tr>

    <form action="NCWSARegistrar.asp?FormStatus=GotTour" method="post">
 
    <tr> 
      <td>&nbsp;</td>
      <td>
				<select name="TourID" size="11" onclick=submit()><%
		
				'	Create SQL to pull upcoming tournament list.

				sSQL = "Select ST.TournAppID, ST.TName,"
				sSQL = sSQL & " convert(char(10),ST.TDateE, 111) as TDateE,"
				sSQL = sSQL & " ST.TCity+', '+ST.TState as TLocation,"
				sSQL = sSQL & " ST.TStatus, ST.TSanType, CASE WHEN left(ST.TSanction,6) <> ST.TournAppID"
				sSQL = sSQL & " THEN ST.TournAppID+'?' ELSE ST.TSanction end as TSanction"
				sSQL = sSQL & " FROM Sanctions.dbo.TSchedul ST WHERE ST.TStatus in (0,1,2,4,5)"
				sSQL = sSQL & " AND ST.TDateE >= GetDate() and ST.TSanType = 1" 
				sSQL = sSQL & " ORDER BY TDateE"
				
        RS.open sSQL

				DO WHILE NOT rs.eof

					response.write("<option value =""" & rs("TournAppID") & rs("TDateE") & RemInvChr(rs("TName")))
					response.write("""> " & rs("TDateE") & "&nbsp;&nbsp; " & rs("TournAppID"))
					response.write("&nbsp;&nbsp; " & rs("TName") & "&nbsp;&nbsp; ( " & rs("TLocation"))
					response.write(" ) </option>")

					rs.moveNEXT

				LOOP

				rs.close %>
				</select></form>
      </td>
      <td>&nbsp;</td>
    </tr>

	 <tr><td>&nbsp;</td></tr>
	   
  </table>
  
  </td></tr></table>

  <%

ELSEIF Request("FormStatus") = "GotTour" THEN 

	'	Create SQL to pull report for the specified tournament.

	sSQL = "Select TX.TeamName, RX.Sex, RX.Skiers, RX.Waivers,"
	sSQL = sSQL & " RX.ASlm, RX.ATrk, RX.AJmp, RX.BSlm, RX.BTrk, RX.BJmp"
	sSQL = sSQL & " FROM (Select Sex, Team, sum(case when SlalomEnt+TrickEnt+JumpEnt"
	sSQL = sSQL & " <> '      ' then 1 else 0 end) as Skiers,"
	sSQL = sSQL & " sum(case when SlalomEnt+TrickEnt+JumpEnt <> '      '"
	sSQL = sSQL & " and WaiverStat = 'X' then 1 else 0 end) as Waivers,"
	sSQL = sSQL & " sum(case when Left(SlalomEnt,1) = 'A' then 1 else 0 end) as ASlm,"
	sSQL = sSQL & " sum(case when Left(TrickEnt,1) = 'A' then 1 else 0 end) as ATrk,"
	sSQL = sSQL & " sum(case when Left(JumpEnt,1) = 'A' then 1 else 0 end) as AJmp,"
	sSQL = sSQL & " sum(case when Left(SlalomEnt,1) > 'A' then 1 else 0 end) as BSlm,"
	sSQL = sSQL & " sum(case when Left(TrickEnt,1) > 'A' then 1 else 0 end) as BTrk,"
	sSQL = sSQL & " sum(case when Left(JumpEnt,1) > 'A' then 1 else 0 end) as BJmp"
	sSQL = sSQL & " FROM Cobra00025.USAWSRank.TeamRotations where TournAppID = '"
	sSQL = sSQL & Session("TourID") & "' group by Sex, Team) as RX"
	sSQL = sSQL & " Join (Select TeamID, TeamName + ' (' + TeamID + ')' as TeamName"
	sSQL = sSQL & " from Cobra00025.USAWSRank.TeamsList Where SptsGrpID = 'NCW') as TX"
	sSQL = sSQL & " on TX.TeamID = RX.Team order by RX.Sex, RX.Team"
				
	RS.open sSQL

	sSex = "?":	SSkr=0: SWvr=0: SASlm=0: SATrk=0: SAJmp=0: SBSlm=0: SBTrk=0: SBJmp=0
	
	%>
		<br>
		<TABLE BORDER="1" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#F5F5F5" width=95% >
			<TR>
				<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center" colspan=3><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> &nbsp; </FONT></TD>
				<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center" colspan=3><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> A Team </FONT></TD>
				<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center" colspan=3><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> B Team </FONT></TD>
			</TR>

	<%

	DO WHILE NOT RS.EOF 
								
		IF sSex <> rs("Sex") THEN

			IF rs("Sex") = "F" THEN %>
				<TR><TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b> Women's Team Entries </b></FONT></TD>
			<% ELSE 
				IF sSex = "F" THEN RecapEntries " Women's Team Sub-" %>
				<TR><TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b> Men's Team Entries </b></FONT></TD>
			<% END IF %>

			<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b> Skrs </b></FONT></TD>
			<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b> Wvrs </b></FONT></TD>
			<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b> Slm </b></FONT></TD>
			<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b> Trk </b></FONT></TD>
			<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b> Jmp </b></FONT></TD>
			<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b> Slm </b></FONT></TD>
			<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b> Trk </b></FONT></TD>
			<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b> Jmp </b></FONT></TD>
			</TR>

			<%

			TSkr=SSkr: TWvr=SWvr: TASlm=SASlm: TATrk=SATrk: TAJmp=SAJmp: TABlm=SBSlm: TBTrk=SBTrk: TBJmp=SBJmp
			SSkr=0: SWvr=0: SASlm=0: SATrk=0: SAJmp=0: SBSlm=0: SBTrk=0: SBJmp=0
		
		END IF
		sSex = RS("Sex")

		%><TR>
			<TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> <%=rs("TeamName")%> </FONT></TD>
			<TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> <%=rs("Skiers")%> </FONT></TD>
			<TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> <%=rs("Waivers")%> </FONT></TD>
			<TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> <%=rs("ASlm")%> </FONT></TD>
			<TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> <%=rs("ATrk")%> </FONT></TD>
			<TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> <%=rs("AJmp")%> </FONT></TD>
			<TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> <%=rs("BSlm")%> </FONT></TD>
			<TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> <%=rs("BTrk")%> </FONT></TD>
			<TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> <%=rs("BJmp")%> </FONT></TD>
		</TR><%
		
		SSkr = SSkr + rs("Skiers"): SWvr = SWvr + rs("Waivers")
		SASlm = SASlm + rs("ASlm"): SATrk = SATrk + rs("ATrk"): SAJmp = SAJmp + rs("AJmp")
		SBSlm = SBSlm + rs("BSlm"): SBTrk = SBTrk + rs("BTrk"): SBJmp = SBJmp + rs("BJmp")

		rs.MoveNext

	Loop
	
	IF sSex = "M" THEN RecapEntries " Men's Team Sub-": ELSE RecapEntries " Women's Team Sub-"

	SSkr = TSkr + SSkr: SWvr = TWvr + SWvr
	SASlm = TASlm + SASlm: SATrk = TATrk + SATrk: SAJmp = TAJmp + SAJmp
	SBSlm = TBSlm + SBSlm: SBTrk = TBTrk + SBTrk: SBJmp = TBJmp + SBJmp

	%><TR>
		<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center" colspan=3><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> &nbsp; </FONT></TD>
		<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center" colspan=3><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> A Team </FONT></TD>
		<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center" colspan=3><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> B Team </FONT></TD>
	</TR><%

	RecapEntries " Overall Tournament "

	%></table><br>
		</td></tr></table><%
	
END IF


'	*************************
SUB RecapEntries (Title)
'	*************************

		%><TR>
			<TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> <b><%=Title%>Totals </b></FONT></TD>
			<TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> <%=SSkr%> </FONT></TD>
			<TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> <%=SWvr%> </FONT></TD>
			<TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> <%=SASlm%> </FONT></TD>
			<TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> <%=SATrk%> </FONT></TD>
			<TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> <%=SAJmp%> </FONT></TD>
			<TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> <%=SBSlm%> </FONT></TD>
			<TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> <%=SBTrk%> </FONT></TD>
			<TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> <%=SBJmp%> </FONT></TD>
		</TR><%
		
END SUB

%>
 