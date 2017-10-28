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
	Session("TourID") = request("TourID")
	Session("TourDate") = request("TourDate")
	Session("TourName") = request("TourName")
END IF

%>

<html>

<head>
<title>NCWSA Tournament Registration Reports</title>

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

	IF Request("TimeFrame") <> "" THEN 
		Session("TimeFrame") = Request("TimeFrame")
	ELSEIF Session("TimeFrame") = "" THEN
		Session("TimeFrame") = "Futr"
	END IF

	%>

	<TABLE BORDER="4" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#DDDDFF" width=95% >
		<tr><br>
			<td BGCOLOR="#42639F">
	        <center><font face="Verdana, Arial, Helvetica, sans-serif" size="4" COLOR="#FFFFFF">
	        		<b>NCWSA Tournament Registration Status Listing</b>
	        </font><br></td>
		</tr>  

		<tr>
			<td>
  			  <br>

				<TABLE BORDER="0" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#DDDDFF" width=95% >
				  <tr>
	
				    <TD width=20% align=center>
					<form action="https://www.usawaterski.org/members/" method="link">
					<input type="submit" style="width:9em" value="Member's Home"
						title="Return to the Member's Only Area Home Page"></form>
			    	</TD>

   			    <td width=55% align=center>
						<form action="ncwsaregistrar.asp" method="post">
						<FONT size="2" face="Verdana, Arial, Helvetica, sans-serif">TimeFrame:&nbsp;
							<input type=radio NAME=TimeFrame VALUE="Futr" <%IF Session("TimeFrame") = "Futr" then Response.Write("checked")%> onclick=submit()>Upcoming&nbsp;
							<input type=radio NAME=TimeFrame VALUE="Past" <%IF Session("TimeFrame") = "Past" then Response.Write("checked")%> onclick=submit()>Recent&nbsp;Past
						</font></form>
			   	</td>

   			    <td width=20% align=center>     				
					<form action="FAQ_NCWRosters.htm" method="post" target="_blank">
					<input type="submit" style="width:7em" value="Instructions"
						title="Instructions and Insights and Tips &#13;and Solutions to Common Problems"></form>
			  	   </td>
	
				</TR>
		     </table> 

				<TABLE BORDER="1" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#F5F5F5" width=95% >

			<TR>
				<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Date</b></FONT></TD>
				<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>TourID</b></FONT></TD>
				<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Tournament Name</b></FONT></TD>
				<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Location</b></FONT></TD>
				<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Sta</b></FONT></TD>
				<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Tms</b></FONT></TD>
				<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Skrs</b></FONT></TD>
				<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Wvrs</b></FONT></TD>
			</TR>

			<%

	'	Create SQL to pull upcoming tournament list.

	sSQL = "Select ST.TournAppID, ST.TName,"
	sSQL = sSQL & " convert(char(10),ST.TDateE, 111) as TDateE,"
	sSQL = sSQL & " ST.TCity+', '+ST.TState as TLocation, US.pwd as EditCode,"
	sSQL = sSQL & " ST.TStatus, ST.TSanType, CASE WHEN left(ST.TSanction,6) <> ST.TournAppID"
	sSQL = sSQL & " THEN ST.TournAppID+'?' ELSE ST.TSanction end as TSanction,"
	sSQL = sSQL & " Case when US.AllowAccess = 0 then 'L' else 'G' end as RegStatus,"
	sSQL = sSQL & " Coalesce(RS.NumTeams,0) as NumTeams, Coalesce(RS.NumSkier,0) as NumSkier,"
	sSQL = sSQL & " Coalesce(RS.NumWaiver,0) as NumWaiver FROM Sanctions.dbo.TSchedul ST"
	sSQL = sSQL & " LEFT JOIN (Select TournAppID, Count(*) as NumTeams,"
	sSQL = sSQL & " sum(Skiers) as NumSkier, Sum(Waivers) as NumWaiver"
	sSQL = sSQL & " FROM (select TournAppID, Team,"
	sSQL = sSQL & " sum(case when SlalomEnt+TrickEnt+JumpEnt"
	sSQL = sSQL & " <> '      ' then 1 else 0 end) as Skiers,"
	sSQL = sSQL & " sum(case when SlalomEnt+TrickEnt+JumpEnt <> '      '"
	sSQL = sSQL & " and WaiverStat = 'X' then 1 else 0 end) as Waivers"
	sSQL = sSQL & " from Cobra00025.USAWSRank.TeamRotations"	
	'	sSQL = sSQL & " where WaiverStat > 'C' group by"
	sSQL = sSQL & " group by"
	sSQL = sSQL & " TournAppID, Team) as TS group by TournAppID) as RS"
	sSQL = sSQL & " on RS.TournAppID = ST.TournAppID"
	sSQL = sSQL & " LEFT JOIN USAWaterski.dbo.Users999 as US"
	sSQL = sSQL & " on Left(US.Name,6) = ST.TournAppID WHERE ST.Deleted = 0"
	sSQL = sSQL & " AND ST.TSanType = 1 AND ST.TStatus in (1,2,4,5)"

	IF Session("TimeFrame") = "Past" THEN
		' sSQL = sSQL & " AND ST.TDateE between DateAdd(mm,-12,GetDate()) and GetDate() ORDER BY TDateE Desc"
		sSQL = sSQL & " AND ST.TDateE between '2011-01-01' and  GetDate() ORDER BY TDateE Desc"
	ELSE
		sSQL = sSQL & " AND ST.TDateE >= DateAdd(dd,-1,GetDate()) ORDER BY TDateE"
	END IF
						
	RS.open sSQL

	DO WHILE NOT rs.eof

		%><tr>
  			  <TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("TDateE")%></FONT></TD>
  			  <TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a
						href="/rankings/View-TournamentsHQ.asp?pvar=TourInfo&TourID=<%=rs("TournAppID")%>"
						<% IF Session("id") = 850 or Session("id") = 6433 or Session("id") = 15757 or Session("id") = 121749 or session("id") = 103995 or Session("id") = 9796 or Session("id") = 6921 or Session("id") = 106475 or Session("id") = 123143 or Session("id") = 138926 or Session("id") = 86262 THEN %>
							title="Click here to view Announcement&#13;Download Edit Code = <%=RS("EditCode")%>" Target="_blank" STYLE="text-decoration:none"><%=rs("TournAppID")%></FONT></TD>
						<% ELSE %>						
							title="Click here to view Announcement" Target="_blank" STYLE="text-decoration:none"><%=rs("TournAppID")%></FONT></TD>
						<% END IF %>
     		  <TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a
						href="ncwsaregistrar.asp?FormStatus=GotTour&TourID=<%=rs("TournAppID")%>&TourDate=<%=rs("TDateE")%>&TourName=<%=RemInvChr(rs("TName"))%>"
		  			  	title="Click here to display a Registration&#13;Detail recap for this Tournament" STYLE="text-decoration:none"><%=RemInvChr(rs("TName"))%></a></FONT></TD>
     		  <TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("TLocation")%></FONT></TD>

				<% IF rs("RegStatus") = "L" THEN %>
					<TD ALIGN="Center" vAlign="Center"><a title="Online Registration Closed for this Tournament"><img src="/admin/Locked.gif" STYLE="border-style:none"></a></td>
				<% ELSE %>
					<TD ALIGN="Center" vAlign="Center"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a 
					title="Online Registration Available for this Tournament"><b>OK</b></a></font></td>
				<% END IF %>

     		  <TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("NumTeams")%></FONT></TD>
     		  <TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("NumSkier")%></FONT></TD>
     		  <TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("NumWaiver")%></FONT></TD>

		</tr><% 

		RS.MoveNext 

	LOOP

	rs.Close
 			  
 	  %></table>

		<TABLE BORDER="0" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#DDDDFF" width=95% >
		  <tr><td>&nbsp;</td></tr></table>

 	  </td></tr>
 	</table><%    


ELSEIF Request("FormStatus") = "GotTour" THEN 

	'	Create SQL to pull report for the specified tournament.

	sSQL = "Select TX.TeamName, RX.Sex, RX.Pending, RX.Skiers, RX.Waivers,"
	sSQL = sSQL & " RX.ASlm, RX.ATrk, RX.AJmp, RX.BSlm, RX.BTrk, RX.BJmp"
	sSQL = sSQL & " FROM (Select Sex, Team,"
	sSQL = sSQL & " sum(case when SlalomEnt+TrickEnt+JumpEnt <> '      '"
	sSQL = sSQL & " then 1 else 0 end) as Skiers,"
	sSQL = sSQL & " sum(case when SlalomEnt+TrickEnt+JumpEnt <> '      '"
	sSQL = sSQL & " and WaiverStat = 'C' then 1 else 0 end) as Pending,"
	sSQL = sSQL & " sum(case when SlalomEnt+TrickEnt+JumpEnt <> '      '"
	sSQL = sSQL & " and WaiverStat = 'X' then 1 else 0 end) as Waivers,"
	sSQL = sSQL & " sum(case when Left(SlalomEnt,1) = 'A' then 1 else 0 end) as ASlm,"
	sSQL = sSQL & " sum(case when Left(TrickEnt,1) = 'A' then 1 else 0 end) as ATrk,"
	sSQL = sSQL & " sum(case when Left(JumpEnt,1) = 'A' then 1 else 0 end) as AJmp,"
	sSQL = sSQL & " sum(case when Left(SlalomEnt,1) > 'A' then 1 else 0 end) as BSlm,"
	sSQL = sSQL & " sum(case when Left(TrickEnt,1) > 'A' then 1 else 0 end) as BTrk,"
	sSQL = sSQL & " sum(case when Left(JumpEnt,1) > 'A' then 1 else 0 end) as BJmp"
	sSQL = sSQL & " FROM Cobra00025.USAWSRank.TeamRotations where TournAppID = '"
	'	sSQL = sSQL & Session("TourID") & "' and WaiverStat > 'C' group by Sex, Team) as RX"
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
				<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center" colspan=4><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> &nbsp; </FONT></TD>
				<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center" colspan=3><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b> A Team </b></FONT></TD>
				<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center" colspan=3><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b> B Team </b></FONT></TD>
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

			<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b> Status </b></FONT></TD>
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

			TSkr=SSkr: TWvr=SWvr: TASlm=SASlm: TATrk=SATrk: TAJmp=SAJmp: TBSlm=SBSlm: TBTrk=SBTrk: TBJmp=SBJmp
			SSkr=0: SWvr=0: SASlm=0: SATrk=0: SAJmp=0: SBSlm=0: SBTrk=0: SBJmp=0
		
		END IF
		sSex = RS("Sex")

		%><TR>

			<TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> <%=rs("TeamName")%> </FONT></TD>

			<% IF rs("Pending") > 0 THEN %>
				<td align="Center"><a title="Rotation Plan Pending / Not Finalized"><img src="/admin/questred17.gif"></a></td>
			<% ELSE %>
				<td align="Center"><a title="Rotation Plan Completed and Submitted"><img src="/admin/Smile17.gif"></a></td>
			<% END IF %>

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
		<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center" colspan=4><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> &nbsp; </FONT></TD>
		<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center" colspan=3><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b> A Team </b></FONT></TD>
		<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center" colspan=3><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b> B Team </b></FONT></TD>
	</TR><%

	RecapEntries " Overall Tournament "

	%></table> <br>
	
		<TABLE BORDER="0" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#FFFFFF" width=95% >
		  <tr>
	
		    <TD width=24% align=center>
				<form action="https://www.usawaterski.org/members/" method="link">
				<input type="submit" style="width:9em" value="Member's Home"
				title="Return to the Member's Only Area Home Page"></form>
	    	</TD>

		    <TD width=26% align=center>
				<form action="/admin/login.asp?UserName=<%=Session("TourID")%>" method="post">
				<input type="submit" style="width:10em" value="Download Entries"
					title="Download Entry details for this&#13;Tournament in an Excel Workbook"></form>
	    	</TD>

		    <td width=26% align=center>     				
				<form action="ncwsaregistrar.asp" method="link">
				<input type="submit" style="width:10em" value="Back to Tour List"
			   	title="Return to Tournament Listing"></form>
		   	</td>

		    <td width=20% align=center>     				
				<form action="FAQ_NCWRosters.htm" method="post" target="_blank">
				<input type="submit" style="width:7em" value="Instructions"
					title="Instructions and Insights and Tips &#13;and Solutions to Common Problems"></form>
				</td>
	
			</TR>
		</table> 
		     
	<br>
		</td></tr></table><%
	
END IF


'	*************************
SUB RecapEntries (Title)
'	*************************

		%><TR>
			<TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF" colspan=2><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> <b><%=Title%>Totals </b></FONT></TD>
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
 