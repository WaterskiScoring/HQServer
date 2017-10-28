<!--#include virtual="/admin/includes/security.asp" -->

<% If not Session("aauth") then response.redirect "Login.asp"
	
IF Len(Request("TourID")) <> 6 then response.redirect "Login.asp" 
IF Request("Status") = "Close" THEN
	NewAllow = 0
ELSEIF Request("Status") = "Open" THEN
	NewAllow = 1
ELSE
	response.redirect "Login.asp" 
END IF

sSQL = "Update USAWaterski.dbo.Users999 set AllowAccess = " & NewAllow
sSQL = sSQL & " where Name = '" & Request("TourID") & "'"

Dim objRegist
Set objRegist = Server.CreateObject("ADODB.Connection")
objRegist.Open Application("WaterSkiConn")
objRegist.Execute (sSQL)
set objRegist = nothing

%>

<html>

<head>
<title>USA Water Ski Admin Index</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" >

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#42639F">
      <p>&nbsp;</p>
      <p align="center"><font face="Verdana" size="6" color="#FFFFFF">
      	USA Water Ski Admin Index</font></p>
      <p align="center"><font face="Verdana" size="4" color="#FFFFFF">
      	Registration Support for -- <%=session("TournamentName")%></font></p>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="185" bgcolor="#42639F" valign="top">

  <!--#include virtual="/admin/includes/menu.asp" -->

    </td>

    <td valign="top" >
    	
    	<table border="0" cellspacing="1" cellpadding="1">

        <tr>
          <td>&nbsp;&nbsp;&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;&nbsp;&nbsp;</td>
        </tr>

        <tr>
          <td>&nbsp;</td>
          <td valign="top"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 

					<% IF NewAllow = 1 THEN %>
					
            <p>Online Registration status for this tournament has now been 
            	<font color=red><b>Re-opened</b></font>.</p>

            <p>Team captains may now prepare new Online Entry and Rotation Plans for their team
            for your tournament, or make revisions to existing entries.&nbsp; Once you have
            eventually done your <b><i>final and official</i></b> download of entries, then
            you will want to close entries at that time.&nbsp; That will cause Captains
            to be referred to the Tournament Registrar at the tournament 
            site for any last-minute changes.</p>

					<% ELSE %>

            <p>Online Registration status for this tournament has now been  
            	<font color=red><b>Closed</b></font>.</p>
            
            <p>Team captains will no longer be allowed to prepare new Online Entry and 
            Rotation Plans for their team for your tournament, nor make revisions to 
            existing entries.&nbsp; Instead, they will be referred to the Tournament 
            Registrar at the tournament site for any last-minute changes.&nbsp; 
            Outstanding Event Waivers may still be executed by participants, but they 
            will now be told that in Registration extracts previously downloaded, their 
            Waiver Status will be shown as outstanding, and hence they should print a 
            copy of their executed waiver confirmation email, and bring it with them 
            to the tournament.</p>

					<% END IF %>
       
			  </font></td>
   		  <td>&nbsp;</td>
			</tr>

      </table></td>

  </tr>
</table>
</body>
</html>





