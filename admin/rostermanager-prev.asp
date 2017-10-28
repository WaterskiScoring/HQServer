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

'	***** Define Session and Member and request form variables


IF request("FormStatus") = "Include" THEN
	IF Request("InclInactives") THEN
		Session("Inactive") = true
	ELSE
		Session("Inactive") = false
	END IF
END IF


'	***** Bailout to Members Login if not auth or no Session("id") value

IF not Session("auth") or Session("id") < 1 then response.redirect "https://www.usawaterski.org/members/login/index.asp"


'	********** Handle Member De-Activate and Re-Activate requests here

IF Right(request("FormStatus"),8) = "Activate" THEN

	sSQL = "Update Cobra00025.USAWSRank.TeamRoster Set DateInactive ="
	IF Left(request("FormStatus"),2) = "Re" THEN sSQL = sSQL & " NULL": ELSE sSQL = sSQL & " LastEvent"
	sSQL = sSQL & " Where MemberID='" & request("MemberID") & "' and Team = '" & Session("TeamID") & "'"
	objConn.Execute (sSQL)

END IF


'	********** Handle Member Removal requests here

IF request("FormStatus") = "Remove" THEN

	sSQL = "Delete from Cobra00025.USAWSRank.TeamRoster Where MemberID='"
	sSQL = sSQL & request("MemberID") & "' and Team = '" & Session("TeamID") & "'"
	objConn.Execute (sSQL)

END IF


'	********** Handle Member Addition requests here
'	********** Check for absence before adding, to avoid "Back" problem.

IF request("FormStatus") = "AddToTeam" THEN

	sSQL = "Select count(*) as Kount FROM Cobra00025.USAWSRank.TeamRoster"
	sSQL = sSQL & " Where Team = '" & Session("TeamID") & "' and MemberID = '"	
	sSQL = sSQL & Request("MemberID") & "'" 
  RS.open sSQL
  Kount = rs("Kount")
  RS.close

  IF Kount = 0 THEN
		sSQL = "Insert into Cobra00025.USAWSRank.TeamRoster (Team, MemberID,"
		sSQL = sSQL & " DateAdded, FirstEvent, LastEvent, NumEvents, DateInactive)"
		sSQL = sSQL & " VALUES ('" & Session("TeamID") & "', '" & request("MemberID")
		sSQL = sSQL & "', GetDate(), GetDate(), GetDate(), 0, NULL)"
		objConn.Execute (sSQL)
	END IF

END IF


'	***** Pull in the currently signed-in Person ID record including NCWSA Team via roster, if any

sSQL = "SELECT TOP 1 Mem.PersonIDWithCheckDigit AS MemberID," 
sSQL = sSQL & " Mem.FirstName+' '+Mem.LastName as MemberName, Mem.Email,"
sSQL = sSQL & " Case when TR.MemberID is not null then TR.TeamID else '???' end as TeamID," 
sSQL = sSQL & " Case when TT.TeamID is not null then TT.TeamName else '???' end as TeamName" 
sSQL = sSQL & " FROM USAWaterski.dbo.members as Mem"

'	Identify Latest Team affiliation for Member -- old version
' sSQL = sSQL & " Left Join (Select MemberID, Substring(Max(Convert(Char(10),LastEvent,111)+Team),"
' sSQL = sSQL & " 11, Len(Max(Convert(Char(10),LastEvent,111)+Team))-10) as TeamID"
' sSQL = sSQL & " from Cobra00025.USAWSRank.TeamRoster Group By MemberID) as TR"

'	Identify Latest Team affiliation for Member -- new version
sSQL = sSQL & " Left Join (Select RX.MemberID, RX.Team as TeamID"
sSQL = sSQL & " from Cobra00025.USAWSRank.TeamRoster as RX"
sSQL = sSQL & " join (select MemberID, Max(LastEvent) as MaxEvt"
sSQL = sSQL & " from Cobra00025.USAWSRank.TeamRoster group by MemberID)" 
sSQL = sSQL & " as ME on ME.MemberID = RX.MemberID"
sSQL = sSQL & " and ME.MaxEvt = RX.LastEvent) as TR"

sSQL = sSQL & " on TR.MemberID = Mem.PersonIDWithCheckDigit Left Join"
sSQL = sSQL & " (Select TeamID, TeamName from Cobra00025.USAWSRank.TeamsList"
sSQL = sSQL & " Where SptsGrpID = 'NCW') as TT on TT.TeamID = TR.TeamID"
sSQL = sSQL & " where Mem.PersonID = " & Session("id")

RS.open sSQL

IF RS.EOF THEN response.redirect "https://www.usawaterski.org/members/login/index.asp"

MemberName = RS("MemberName")
Session("MemName") = MemberName
Session("MemEmail") = RemInvChr(RS("Email"))
MemberID = RS("MemberID")
MemberTeam = RS("TeamID")
Session("TeamID") = MemberTeam
MemTeamName = RS("TeamName")

RS.close

'	********** If we have a Request value of NewTeam (from form), then
'	********** insert a Team Roster record for this member, and then
'	********** set that Team Code into the local and Session variables.

IF Request("FormStatus") = "NewTeam" THEN

	MemberTeam = left(Request("NewTeam"),3)
	MemTeamName = Right(Request("NewTeam"),len(Request("NewTeam"))-3)
	Session("TeamID") = MemberTeam
	
	sSQL = "Insert into Cobra00025.USAWSRank.TeamRoster (Team, MemberID,"
	sSQL = sSQL & " DateAdded, FirstEvent, LastEvent, NumEvents, DateInactive)"
	sSQL = sSQL & " VALUES ('" & MemberTeam & "', '" & MemberID & "', GetDate(),"
	sSQL = sSQL & " GetDate(), GetDate(), 0, NULL)"

	objConn.Execute (sSQL)

END IF


'	********** Now finally begin displaying -- this does the header

IF MemberTeam = "???" THEN
	MemTeamName = "( Team to be Determined )"
ELSE
	Session("TeamName") = MemTeamName
END IF

%>

<html>

<head>
<title>NCWSA Team Roster Management</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" >

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#42639F">
      <p>&nbsp;</p>
      <p align="center"><font face="Verdana" size="5" color="#FFFFFF">
      	<b>NCWSA Team Roster Management</b></font></p>
      <p align="center"><font face="Verdana" size="4" color="#FFFFFF">
      	<%=Session("MemName")%>&nbsp;&nbsp; as Administrator for:&nbsp;&nbsp;&nbsp;&nbsp;<%=Session("TeamName")%>&nbsp;&nbsp; ( <%=Session("TeamID")%> ) </font></p>
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

'	********** Primary alternatives here -- Depends on whether we have a team
'	********** defined for this member or not.  If not, then we present a list
'	********** of existing teams via drop-down for selection.

IF MemberTeam = "???" THEN 

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
        The system currently does not have you associated with a particular
        NCWSA Team.&nbsp; Please select the team for which you are going to
        serve as an administrator, from the drop-down list below.&nbsp;
        You will then be added to that team.<br>&nbsp;
        </font></td>
      <td>&nbsp;</td>
    </tr>

    <form action="rostermanager.asp?FormStatus=NewTeam" method="post">
 
    <tr> 
      <td>&nbsp;</td>
      <td>
				<select name="NewTeam" size="11" onclick=submit()><%
		
				sSQL = "SELECT TeamID, Max(TeamName) as TeamName"
				sSQL = sSQL & " FROM Cobra00025.USAWSRank.TeamsList GROUP BY TeamID"
				sSQL = sSQL & " Order by Max(TeamName)"
				
        RS.open sSQL

				DO WHILE NOT rs.eof

					response.write("<option value =""" & rs("TeamID") & rs("TeamName"))
					response.write("""> " & rs("TeamName") & " ( " & rs("TeamID"))
					response.write(" ) </option><br>")

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

ELSE

  ' Set up query to pull Team Roster Table for current TeamID

	sSQL = "SELECT Mem.PersonIDWithCheckDigit AS MemberID," 
	sSQL = sSQL & " Mem.LastName, Mem.FirstName, Mem.DivisionCode1 as SptsDiv,"
	sSQL = sSQL & " Mem.City+', '+Mem.State as HomeTown,"
	sSQL = sSQL & " Datepart(yyyy,Mem.BirthDate) as BirthYear, Left(Mem.Sex,1) as Sex,"
	sSQL = sSQL & " Convert(Char(10),TR.FirstEvent,111) as FirstEvent," 
	sSQL = sSQL & " Convert(Char(10),TR.LastEvent,111) as LastEvent, TR.NumEvents," 
	sSQL = sSQL & " CASE when TR.DateInactive is Null then 'A' else 'I' end as TeamStat," 
	sSQL = sSQL & " CASE when Mem.EffectiveTo < getdate() then 'X'"
	sSQL = sSQL & " when Typ.CanSkiInTournaments = 0 then 'U'" 
	sSQL = sSQL & " when Mem.WaiverStatusID = 0 then 'W'"
	sSQL = sSQL & " when DateAdd(dd,-21,Mem.EffectiveTo) < getdate() then 'P'"
	sSQL = sSQL & " else 'G' end as MemStat,"
	sSQL = sSQL & " CASE when DateAdd(dd,-21,Mem.EffectiveTo) < getdate() THEN"
	sSQL = sSQL & " 'Exp '+right(Convert(char(10),Mem.EffectiveTo,111),5) ELSE '' end as ExpMMDD,"
	sSQL = sSQL & " CASE when patindex('%@%',Mem.Email) > 0 then 'Y' else 'N' end as EMStat"
	sSQL = sSQL & " FROM Cobra00025.USAWSRank.TeamRoster as TR"
	sSQL = sSQL & " JOIN USAWaterski.dbo.members as Mem"
	sSQL = sSQL & " ON Mem.PersonIDWithCheckDigit = TR.MemberID"
	sSQL = sSQL & " JOIN USAWaterski.dbo.MembershipTypes as Typ"
	sSQL = sSQL & " ON Typ.MembershipTypeID = Mem.MemberShipTypeCode"
	sSQL = sSQL & " Where TR.Team = '" & MemberTeam & "'"
	
	IF not Session("Inactive") THEN
		sSQL = sSQL & " AND TR.DateInactive is NULL"
	END IF
	
	sSQL = sSQL & " ORDER BY Mem.Sex, Mem.LastName, Mem.FirstName"
	
 	RS.open sSQL

	%>

	<TABLE BORDER="4" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#DDDDFF" width=95% >
		<tr><br>
			<td BGCOLOR="#42639F">
	        <center><font face="Verdana, Arial, Helvetica, sans-serif" size="4" COLOR="#FFFFFF"><b>Click on Buttons below to Manage your Team Roster</b></font>
			<br></td>
		</tr>  

		<tr>
			<td>
  			  <br>

				<TABLE BORDER="0" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#DDDDFF" width=95% >
				  <tr>
	
				    <TD width=25% align=center>
					<form action="FindToAdd.asp?FormStatus=newsearch" method="post">
					<input type="submit" style="width:9em" value="Add a Member"
						title="Add a new Member to your Team Roster&#13;by Searching the Membership Database"></form>
			    	</TD>

   			    <td width=25% align=center>     				
					<form action="RotationPlanner.asp?FormStatus=PickTour" method="post">
				   <input type="submit" style="width:9em" value="Rotation Plans"
				   	title="Review or Edit or Build a Rotation Plan for&#13; your Team, for a specific Tournament"></form>
			   	</td>

   			    <td width=25% align=center>     				
					<form action="rostermanager.asp?FormStatus=Include" method="post">
				   <input type="CheckBox" name="InclInactives" onclick=submit() value="True"
					<% IF Session("Inactive") THEN response.write("checked" )%> title="Shows Inactive Members when Checked"
				    ><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> Show<br>Inactives</FONT></form>
			   	</td>

   			    <td width=25% align=center>     				
					<form action="FAQ_NCWRosters.htm" method="post" target="_blank">
					<input type="submit" style="width:7em" value="Tips/Help"
						title="Tips and Explanations and Insights and &#13;Answers to Frequently Asked Questions"></form>
			  	   </td>
	
				</TR>
		     </table> 

				<TABLE BORDER="1" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#F5F5F5" width=95% >

			<%

	sSex = "?"

	DO WHILE NOT RS.EOF 
								
		IF rs("Sex") <> sSex THEN %>

			<TR>
				<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Member ID</b></FONT></TD>

			<% IF rs("Sex") = "F" THEN %>
				<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><b> Women's Team </b></FONT></TD>
			<% ELSE %>
				<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><b> Men's Team </b></FONT></TD>
			<% END IF %>

				<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>City & State</b></FONT></TD>
				<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Age/Gender</b></FONT></TD>
				<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Status *</b></FONT></TD>
				<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Note(s)</b></FONT></TD>
				<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Action</b></FONT></TD>
			</TR>

		<% END IF

		sSex = rs("Sex")

		sMembAge = Datepart("yyyy",now) - rs("BirthYear") - 1

		sMemStat = rs("Memstat")
		sExpMMDD = rs("ExpMMDD")

		'	**********	Now if Membership Status is not "G" (good), 
		'	**********	then pull latest status from HQ server and re-classify

		IF sMemStat <> "G" THEN

			sSQL = "Select CASE when MH.EffectiveTo < GetDate() then 'X'"
			sSQL = sSQL & " when MT.CanCompete = 0 then 'U'"
			sSQL = sSQL & " when MH.WaiverStatusID = 0 then 'W'"
			sSQL = sSQL & " when DateAdd(dd,-21,MH.EffectiveTo) < GetDate() then 'P'"
			sSQL = sSQL & " else 'G' end as MemStat,"
			sSQL = sSQL & " CASE when DateAdd(dd,-21,MH.EffectiveTo) < GetDate() THEN"
			sSQL = sSQL & " 'Exp '+right(Convert(char(10),MH.EffectiveTo,111),5) ELSE '' end as ExpMMDD"
			sSQL = sSQL & " FROM Waterski.dbo.[Membership History] as MH JOIN"
			sSQL = sSQL & " (Select [Person ID] as PersonID, Max(EffectiveTo) as MaxEffTo"
			sSQL = sSQL & " From Waterski.dbo.[Membership History] where [Person id] = "
			sSQL = sSQL & right(rs("MemberID"),8) & " group by [Person ID]) as ME"
			sSQL = sSQL & " ON MH.[Person ID] = ME.PersonID AND MH.EffectiveTo = ME.MaxEffTo"			
			sSQL = sSQL & " JOIN	waterski.dbo.tblMembershipTypeCodes as MT"
			sSQL = sSQL & " on MT.[Membership Type Code] = MH.[Membership Type Code]"

			RSHQ.open sSQL
			IF NOT RSHQ.eof THEN sMemStat = RSHQ("MemStat"): sExpMMDD = RSHQ("ExpMMDD")
			RSHQ.close
		 	
		END IF

		SELECT CASE sMemStat&rs("EMStat")
			CASE "PN": sNote = sExpMMDD & "<br>No eMail"
			CASE "UN": sNote = "Nd Upgrd<br>No eMail"
			CASE "WN": sNote = "Nd Ann Wvr<br>No eMail"
			CASE "XN": sNote = "Nd Renew<br>No eMail"
			CASE "GN": sNote = "No eMail"
			CASE "PY": sNote = sExpMMDD
			CASE "UY": sNote = "Nd Upgrd"
			CASE "WY": sNote = "Nd Ann Wvr"
			CASE "XY": sNote = "Nd Renew"
			CASE ELSE: sNote = "OK"
		END SELECT
		
		IF rs("TeamStat") = "A" THEN sTeamStat = "Active": ELSE sTeamStat = "Inactive"
		IF rs("SptsDiv") <> "NCW" THEN
			sTeamStat = sTeamStat & "<font color=red><b> **</b></font>"
		else
			sTeamStat = sTeamStat & "&nbsp;&nbsp;&nbsp;&nbsp; "
		end if
			
		%><tr>
  			  <TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("MemberID")%></FONT></TD>
  			  <TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("LastName")%>, <%=rs("FirstName")%></FONT></TD>
     		  <TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("HomeTown")%></FONT></TD>
     		  <TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=sMembAge&" / "&rs("Sex")%></FONT></TD>


     		  <TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT Color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a
      		  title="First Event: <%=rs("FirstEvent")%>&#13;Latest Event: <%=rs("LastEvent")%>&#13;Number of Events: <%=rs("NumEvents")%>&#13;Primary SptsDiv: <%=rs("SptsDiv")%>">
      		<%=sTeamStat%></FONT></TD>

 
     		  <% IF sNote = "OK" THEN %>
	     		  <TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> OK </FONT></TD>
     		  <% ELSE %>
	     		  <TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#FF0000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=sNote%></FONT></TD>
     		  <% END IF %>

     		  <% IF rs("TeamStat") = "A" and rs("NumEvents") > 0 THEN %>
      		  <TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a
	      		  href="rostermanager.asp?FormStatus=DeActivate&MemberID=<%=rs("memberid")%>" title="Click to place this member into Inactive status"><img src="/admin/DeActivate.gif" STYLE="border-style:none"></a></FONT></TD>
	     	  <% ELSEIF rs("TeamStat") = "A" and rs("NumEvents") <= 0 THEN %>
      		  <TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a
	      		  href="rostermanager.asp?FormStatus=DeActivate&MemberID=<%=rs("memberid")%>" title="Click to place this member into Inactive status"><img src="/admin/DeActivate.gif" STYLE="border-style:none"></a>&nbsp;&nbsp; <a
	      		  href="rostermanager.asp?FormStatus=Remove&MemberID=<%=rs("memberid")%>" title="Click to Remove this member from your team"><img src="/admin/Remove.gif" STYLE="border-style:none"></a></FONT></TD>
     		  <% ELSE %>
      		  <TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a
	      		  href="rostermanager.asp?FormStatus=ReActivate&MemberID=<%=rs("memberid")%>" title="Click to restore this member to Active status"><img src="/admin/ReActivate.gif" STYLE="border-style:none"></a></FONT></TD>
     		  <% END IF %>

		</tr><% 

		RS.MoveNext 

	LOOP

	rs.Close
 			  
 	  %></table>

		<TABLE BORDER="0" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#DDDDFF" width=95% >
		  <tr><td>&nbsp;</td></tr></table>

 	  </td></tr>
 	</table><%    

END IF

set rs = nothing

%>
