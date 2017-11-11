<!--#include virtual="epl/functions.asp" -->

<% 

Dim objConn, RS, objfso, objMessage

Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open Application("WaterSkiConn")
Set RS = Server.CreateObject("ADODB.RecordSet")
RS.ActiveConnection = objConn

Set objfso = CreateObject("Scripting.FileSystemObject")

	Set objMessage = CreateObject("CDO.Message")

	'=====  This section provides the configuration information for the remote SMTP server.
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	'Name or IP of Remote SMTP Server
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.epolk.net"
	'Type of authentication, NONE, Basic (Base64 encoded), NTLM
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0
	'Your UserID on the SMTP server
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendusername") = "viper@epolk.org"
	'Your password on the SMTP server
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "V1p3rMAIL0090"
	'Server port (typically 25)
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	'Use SSL for the connection (False or True)
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
	'Connection Timeout in seconds (the maximum time CDO will try to establish a connection to the SMTP server)
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
	objMessage.Configuration.Fields.Update
	'=====  End remote SMTP server configuration section

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
		If (Mid(strInput, i, 1)) = "@" Or (Mid(strInput, i, 1)) = "." Or (Mid(strInput, i, 1)) = " " Or (Mid(strInput, i, 1)) = "-" Or (Mid(strInput, i, 1)) = "_" Then
				workingstring = workingstring & Mid(strInput, i, 1)
		End If
	Next
	RemInvChr = workingstring
End Function


Dim sPID, sMemberID, sLastName, sFirstName, sMemEmail, NumTour
Dim sTourID, sTourName, sTourDate, PathtoWaivers, ReleaseVersion
Dim sSQL


PathtoWaivers = Server.mappath("/")&"\rankings\release"
ReleaseVersion = "adlt2010"


'	*****	FormStatus on arrival should be "List" or "Tour" or "Accept" -- 

'	*****	FormStatus=List is invoked from Members-Only menu

'	*****	FormStatus=Tour is invoked from form in email, or from response to List

'	*****	FormStatus=Tour with Session("auth") false, we then 
'			set Session("auth") and Session("id") etc, as tho signed in.

'	*****	FormStatus=Accept is invoked upon acceptance of the waiver.

IF Request("FormStatus") = "Accept" THEN
	IF Request("PID") < 1 or Request("TourID") = "" then response.redirect "/members/login/index.asp"
	sPID = Request("PID")
ELSEIF Request("FormStatus") = "List" THEN
	IF not Session("auth") or Session("id") < 1 then response.redirect "/members/login/index.asp"
	sPID = Session("id")
ELSEIF Request("FormStatus") = "Tour" THEN
	IF Request("PID") < 1 or Request("TourID") = "" then response.redirect "/members/login/index.asp"
	sPID = Request("PID")
ELSE
	response.redirect "/members/login/index.asp"
END IF

'	***** Pull in the currently signed-in Person ID record, 
'	*****	left joined to any outstanding NCWSA Rotation Plan Records.
'	*****	Limit to requested TournAppID only, if FormStatus="Tour"

'	*****	So this query will pull at least one row for member, and may return
'	*****	multiple rows, if that member has more than one outstanding waiver.

sSQL = "SELECT TR.MemberID," 
sSQL = sSQL & " Mem.FirstName, Mem.LastName, Mem.Email,"
sSQL = sSQL & " Coalesce(TR.WaiverStat,' ') as WaiverStat,"
sSQL = sSQL & " ST.TournAppID as TourID, ST.TName,"
sSQL = sSQL & " convert(char(10),ST.TDateE, 101) as TDateE,"
sSQL = sSQL & " ST.TCity+', '+ST.TState as TLocation"
sSQL = sSQL & " From USAWaterski.dbo.memberslive as Mem"
sSQL = sSQL & " Left Join Cobra00025.USAWSRank.TeamRotations as TR"
sSQL = sSQL & " on Mem.PersonID = cast(right(TR.MemberID,8) as integer)"
sSQL = sSQL & " and TR.WaiverStat = 'E'"
IF Request("TourID") <> "" THEN
	sSQL = sSQL & " and TR.TournAppID = '" & Request("TourID") & "'"
END IF
sSQL = sSQL & " Left Join Sanctions.dbo.TSchedul as ST"	
sSQL = sSQL & " on ST.TournAppID = TR.TournAppID"
sSQL = sSQL & " Where Mem.PersonID = " & sPID
sSQL = sSQL & " order by ST.TDateE"

RS.open sSQL

IF RS.EOF THEN response.redirect "/members/login/index.asp"

NumTour = 0

DO WHILE NOT rs.EOF

	sMemberID = rs("MemberID")
	sFirstName = rs("FirstName")
	sLastName = rs("LastName")
	sMemEmail = RemInvChr(trim(rs("Email")))

	IF rs("WaiverStat") = "E" THEN
		sTourID = rs("TourID")
		sTourName = RemInvChr(trim(rs("TName")))
		sTourDate = rs("TDateE")
		NumTour = NumTour + 1
	END IF

	rs.movenext

LOOP


RS.close

Session("MemName") = sFirstName & " " & sLastName
Session("MemEmail") = sMemEmail

IF not Session("auth") THEN
	Session("auth") = True
	Session("id") = Request("PID")
	Session("name") = sFirstName
END IF


'	*****	Now display page headers

%>

<html>

<head>
<title>USA Water Ski Event Waiver Affirmation</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" >

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#42639F">
      <p>&nbsp;</p>
      <p align="center"><font face="Verdana" size="6" color="#FFFFFF">
      	<b>USA Water Ski Event Waiver Affirmation</b></font></p>
      <p align="center"><font face="Verdana" size="4" color="#FFFFFF">
      	Member:&nbsp;&nbsp; <%=Session("MemName")%>&nbsp;&nbsp; ( <%=sMemberID%> )
		<% IF Request("TourID") <> "" THEN Response.write("&nbsp;&nbsp;&nbsp;&nbsp; TourID:&nbsp;&nbsp; " & Request("TourID") & " ") %> 
      	 </font></p>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>

<table border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="180" valign="top" bgcolor="#42639F">

			<br>
	        &nbsp;<a href="/members/"><font face="arial" size="2" COLOR="#FFFFFF">Member's Only Home</font></a><br>&nbsp;<br>
			<br>
            <font face="Verdana" size="1">&nbsp;<font COLOR="#FFFFFF">Powered by</font> <a href="http://www.epolk.com"><font COLOR="#FFFFFF">ePolk.com</font></a><br></font>
	
    </td>

<td width="760" >

	<table>
    <tr> 
      <td width="20">&nbsp;</td>
      <td width="560">&nbsp;</td>
      <td width="20">&nbsp;</td>
    </tr>
    
<%

'	*****	If we're in accept mode, then email Waiver Details and post to table and confirm.

IF Request("FormStatus") = "Accept" and NumTour = 1 THEN

	ebody = "<html><head><title>Waiver and Release</title>"
	ebody = ebody & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
	ebody = ebody & "</head><body bgcolor=""#FFFFFF"" text=""#000000""><div align=""center"">"
	ebody = ebody & "<font face=""Verdana, Arial, Helvetica, sans-serif"">"

	ebody = ebody & "<TABLE BORDER=4 ALIGN=CENTER CELLPADDING=3 CELLSPACING=0 BGCOLOR=""#F5F5F5"" width=85% >"
	ebody = ebody & "<TR><TD BGCOLOR=red><center><font color=#FFFFFF size=4><b>Waiver and Release Form</b></font></TD></TR>"
 
	ebody = ebody & "<TR><TD VALIGN=top><TABLE border=""0"" cellspacing=""0"" cellpadding=""3"" width=""100%""><tr>"

	subTitle="Waiver for ADULT Participant - WaiverID: " & ReleaseVersion

	ebody = ebody & "<td Align=center>"	
	ebody = ebody & "<font size=4 ><b>PARTICIPANT WAIVER AND RELEASE OF LIABILITY,</b></font><br>"
	ebody = ebody & "<font size=4><b>ASSUMPTION OF RISK AND INDEMNITY AGREEMENT</b></font><br>"
	ebody = ebody & "<font size=2><b>" & subTitle & "</b></font><br><br>"
	ebody = ebody & "<font color=""blue"" size=3><b>" & sTourID & "&nbsp;&nbsp;&nbsp; " & sTourName & "&nbsp;&nbsp;&nbsp; " & sTourDate & "</b></font><br><br>"
	ebody = ebody & "<font size=2><b>MemberID = </font><font color=""blue"" size=2>" & sMemberID
	ebody = ebody & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font><font size=2>Participant:</font>"
	ebody = ebody & "<font color=""blue"" size=2>&nbsp;&nbsp;" & sFirstname & "&nbsp;" & sLastName & "</font></b><br>"

	ebody = ebody & "</center><br></td></tr>"

	ebody = ebody & "<td Align=left><P><font size=1>"
	
	IF objfso.FileExists(PathtoWaivers & "\waiver-"&ReleaseVersion&".txt") THEN
		SET objstream=objFSO.opentextfile(PathtoWaivers & "\waiver-"&ReleaseVersion&".txt")
		IF NOT objstream.atendofstream THEN
			DO WHILE not objstream.atendofstream
				ebody = ebody & objstream.readline & "<br>"
			LOOP
		END IF
		objstream.close 
	END IF

	ebody = ebody & "</font></P></td></tr><tr><td Align=center><br>"
	ebody = ebody & "<font color=""red"" size=3><b>By acccepting this waiver I have acknowledged that I am the 'PARTICIPANT' listed above.</b></font><br><br>"
	ebody = ebody & "<font size=2><b>Date Accepted:&nbsp;&nbsp</font><font color=""blue"" size=2>" & DATE & "</b></font><br>"
	ebody = ebody & "</td></tr></td><br></tr></TABLE></TD></TR></TABLE></font></div></body></html>"

	objMessage.Subject = "USA Water Ski WAIVER & RELEASE  TourID: " & sTourID & " - Member: "& sFirstName & " " & sLastName
	objMessage.To = """" & sFirstName & " " & sLastName & """ <" & sMemEmail & ">"

'	objMessage.CC = "competition@usawaterski.org"

	objMessage.From = "competition@usawaterski.org"
	objMessage.HTMLBody = ebody	

	objMessage.Send

	Set objMessage = Nothing


	'	**********	Next we post "Executed" Status to the Rotation Plan table.

	sSQL = "Update Cobra00025.USAWSRank.TeamRotations set WaiverStat = 'X'"
	sSQL = sSQL & " Where MemberID = '" & sMemberID
	sSQL = sSQL & "' and TournAppID = '" & sTourID & "'"
	objConn.Execute (sSQL)


	'	**********	Next we interrogate the Open/Closed status of the specified tournament.
	'	**********	and then present appropriate confirmation language, based on that status.
	
	sSQL = "Select AllowAccess from USAWaterski.dbo.Users999 where Name = '" & sTourID & "'"
 	RS.open sSQL


	%>

	<tr> 
		<td>&nbsp;</td>

		<TD><font face="Verdana, Arial, Helvetica, sans-serif" size="2">
				<p>Your acceptance of this Event Waiver has been noted in our system,
				and an e-mail copy has been sent to you for your records, to email
				address <<%=sMemEmail%>>.</p>

				<% IF not rs("AllowAccess") THEN %>
					<p><font color=red><b>Please be aware that the registration
					extract for this tournament has already been pulled, and
					that showed you as not yet having executed this Waiver.&nbsp;
					Therefore, we recommend that you print a hardcopy from the
					email we have sent you, and bring that with you to show to
					the Registrar at the tournament site.</b></font></p>
				<% END IF %>
				
				<p>Click the button below to return to the Member's Only area ...<br>
				&nbsp;</p></font>

				<form action="/members/" method="link">
			   <input type="submit" style="width:15em" value="Return to Member's Only Area"
			   	title="Return to Member's Only Area"></form>
		    
      </td>
      <td>&nbsp;</td>
    </tr>

	 <tr><td>&nbsp;</td></tr>
	   
	</table>

	</td></tr>
	</table>
 	
 	<%    

	rs.close


'	*****	If we find no tournaments, then display a "None Found" screen

ELSEIF NumTour = 0 THEN

		%>

	<tr> 
		<td>&nbsp;</td>

		<TD><font face="Verdana, Arial, Helvetica, sans-serif" size="2">

		<% IF Request("FormStatus") = "Tour" THEN %>

			<p><font color=red><b>The Event Waiver for your entry to NCWSA Tournament
			<%=Request("TourID")%> has already been executed and filed.&nbsp; 
			Thank you.</b></font></p>
			
			<p>Click the button below to see if there are any other Waivers outstanding 
			for you, for other tournaments ...<br>&nbsp; </p>
			</font>
		    
			<form action="NCWSAWaiver.asp?FormStatus=List" method="post">
		   <input type="submit" style="width:15em" value="Check for other Waivers"
		   	title="Check to see if I have any other Waivers outstanding"></form>
		    
		<% ELSE %>

			<p><font color=red><b>There are no NCWSA Tournament Entries pending for you 
			that require execution of an Event Waiver.</b></font></p>
		
			<p>Click the button below to return to the USA Water Ski
			Member's Only area ...<br>&nbsp; </p>
			</font>
		    
			<form action="/members/" method="link">
		   <input type="submit" style="width:15em" value="Return to Member's Only Area"
		   	title="Return to Member's Only Area"></form>

		<% END IF %>

      </td>
      <td>&nbsp;</td>
    </tr>

	 <tr><td>&nbsp;</td></tr>
	   
  </table>

	</td></tr>
	</table>
 	
 	<%    
	

'	*****	If we find more than one, then present a "Pick one" screen.
	
ELSEIF NumTour > 1 THEN

	%>

    <tr> 
      <td>&nbsp;</td>
      <td valign="top">
        <font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        Following is a list of upcoming NCWSA Tournaments for which you have
        been entered by your team captain, and for which you must execute an Event Waiver.&nbsp; 
        Please select the specific tournament for which you are going to execute an Event Waiver
        at this time, from this list ...<br>&nbsp;
        </font></td>
      <td>&nbsp;</td>
    </tr>

    <form action="NCWSAWaiver.asp?FormStatus=Tour" method="post">
    
	 <input type="hidden" name="PID" value="<%=Session("id")%>">
    
    <tr> 
      <td>&nbsp;</td>
      <td>
				<select name="TourID" size="4" onclick=submit()><%
		
        RS.open sSQL

				DO WHILE NOT rs.eof
				
					IF rs("WaiverStat") = "E" THEN
						response.write("<option value =""" & rs("TourID"))
						response.write("""> " & rs("TDateE") & "&nbsp;&nbsp; " & rs("TourID"))
						response.write("&nbsp;&nbsp; " & rs("TName") & "&nbsp;&nbsp; ( " & rs("TLocation"))
						response.write(" ) </option>")
					END IF

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


'	*****	Otherwise we have just one single tournament, so then let's show that waiver

ELSE

	%>

    <tr>
    <td>&nbsp;</td>
    <td>
 
		<font face="Verdana, Arial, Helvetica, sans-serif">
		
             <TABLE BORDER="4" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#F5F5F5" width=100%>
		     <TR>
			<TD BGCOLOR="red"><center><font  color="#FFFFFF" size="4"><b>Waiver and Release Form</b></font></TD>
		     </TR>  
 
		     <TR>
			<TD VALIGN="top">
  			   <TABLE BORDER="0" VALIGN="top" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" width=100%>
				<tr>
				   <% ' ----- BEGINNING OF CELL -------- %>	
				   <td>	

		<form action = "NCWSAWaiver.asp?FormStatus=Accept" method="post">

			<INPUT type="hidden" NAME="TourID" VALUE="<%=sTourID%>">
			<INPUT type="hidden" NAME="PID" VALUE="<%=sPID%>">
		
		   <center>	
	 	   <font size="4" ><b>PARTICIPANT WAIVER AND RELEASE OF LIABILITY,</b></font><br>
		   <font size="4"><b>ASSUMPTION OF RISK AND INDEMNITY AGREEMENT</b></font>
		   <br>
		   <font size="2"><b>Waiver for ADULT Participant - WaiverID: <%=ReleaseVersion%></b></font>
		   <br><br>

		   <font color="blue" size="3"><b><%=sTourID%>&nbsp;&nbsp;&nbsp; <%=sTourName%>&nbsp;&nbsp;&nbsp; <%=sTourDate%></b></font>
		   <br><br>
		   <font size="2"><b>MemberID = </font><font color="blue" size="2"><%=sMemberID%>
			</font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font size="2">Participant:</font>
			<font color="blue" size="2"><%=sFirstname%>&nbsp;<%=sLastName%></font></b><br>
		   </center><br>
		   <P><font size="1" ><left><%
	
		  IF objfso.FileExists(PathtoWaivers & "\waiver-"&ReleaseVersion&".txt") THEN
			SET objstream=objFSO.opentextfile(PathtoWaivers & "\waiver-"&ReleaseVersion&".txt")

			IF NOT objstream.atendofstream THEN
				DO WHILE not objstream.atendofstream
					response.write(objstream.readline)
				   	response.write("<br>")
				LOOP
			END IF

		  END IF
		  objstream.close  %>

		  </left></font></P>
			<center>
			<font size="4" color="red" ><b>The name listed above must be the person completing this form.</b></font>
			<br>
			<font size="4" color="red" ><b>Minors under 18 Years may NOT accept liability waiver.</b></font>
			</center>
		<center>
			<br>
			<font size="2"><b>By acccepting this waiver I acknowledge that I am the 'PARTICIPANT' listed above.</b></font><br>
			<br><br>
			<input type="submit" value="I Accept this Waiver" title="Click here to accept the above terms, and &#13;you will then be emailed a copy of this Waiver">
			&nbsp;&nbsp;&nbsp;&nbsp
			<font size="2"><b>Date: <%=DATE%></b></font>
		<br>
		</center>


		</form>

		</td> 

		<br>
		</tr>
		</TABLE>

		</TD></TR>
	     </TABLE>

		</font>

		</td>
		<td>&nbsp;</td>
		</tr>

	 <tr><td>&nbsp;</td></tr>
	   
  </table>
  
  </td></tr></table>

	<%

END IF

set rs = nothing

%>

