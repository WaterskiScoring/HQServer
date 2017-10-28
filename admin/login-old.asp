<!--#include virtual="/epl/functions.asp" -->
<html>

<head>
<title>Admin Index</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" >

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#42639F">
      <p>&nbsp;</p>
      <p align="center"><font face="Verdana" size="6" color="#FFFFFF">USA Water 
        Ski Admin Login</font></p>
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


<%

Dim FormUserName
FormUserName = Request("UserName")

If Request.Form <> "" then

Dim InvalidConn
Set InvalidConn = Server.CreateObject("ADODB.Connection")
InvalidConn.Open Application("ePolkConn")

Dim InvalidRS2
Set InvalidRS2 = Server.CreateObject("ADODB.RecordSet")
InvalidRS2.ActiveConnection = InvalidConn
InvalidRS2.LockType = 3	'adLockOptimistic
InvalidRS2.Open "select count(*) as county from InvalidLogins where IPAddress = '" & Request.ServerVariables("REMOTE_ADDR") & "' and TimeHappened >= DateAdd(n, -30,  { fn NOW() })"

if InvalidRS2("county") >= 5 then 
	response.write "<font face=""arial""><br><b>&nbsp;&nbsp;&nbsp;Login failed 5 times.&nbsp; You must wait 30 minutes before making another attempt.</b></font>"
	response.end
end if
InvalidRS2.close

	If Application("UseSave") then
		If Request.Form("SaveInfo") = "save" then
			Response.Cookies("ADirectory")("UserName") = Request.Form("UserName")
			Response.Cookies("ADirectory")("Password") = Request.Form("Password")
			Response.Cookies("ADirectory")("checked") = " checked"
			Response.Cookies("ADirectory").Expires = Date() + 365
			If Application("isSecure") then
				Response.Cookies("UserName").Secure = True
			End If
		Else
			Response.Cookies("ADirectory")("UserName") = ""
			Response.Cookies("ADirectory")("Password") = ""
			Response.Cookies("ADirectory")("checked") = ""
			Response.Cookies("ADirectory").Expires = Date() + 365
			If Application("isSecure") then
				Response.Cookies("UserName").Secure = True
			End If
		End If
	End If

	
	Dim objRS, ValidLogin
	ValidLogin = False
	Set objRS = Server.CreateObject("ADODB.RecordSet")
	objRS.ActiveConnection = objConn
	objRS.LockType = 3	'adLockOptimistic
	objRS.Open "SELECT * FROM Users999 WHERE lower(Name) = '" & epl_removeinvalidchars(Request.Form("UserName")) & "' AND Pwd = '" & epl_removeinvalidchars(Request.Form("Password")) & "'"
	If objRS.EOF then
		'IF they entered a 7 digit username, trim the 7th digit and try again
		if len(epl_removeinvalidchars(Request.Form("UserName"))) = 7 then
			objRS.close
			objRS.Open "SELECT * FROM Users999 WHERE lower(Name) = '" & left(epl_removeinvalidchars(Request.Form("UserName")),6) & "' AND Pwd = '" & epl_removeinvalidchars(Request.Form("Password")) & "'"	
			If not objRS.EOF then
				ValidLogin = True
			end if
		end if
	else
		ValidLogin = True
	end if
	if ValidLogin = False then
		strError = "<font face=""Verdana"" size=""2"" color=""#ff0000"">ERROR: Invalid UserName or Password!</font><br><br>"
		Dim InvalidRS
		Set InvalidRS = Server.CreateObject("ADODB.RecordSet")
		InvalidRS.ActiveConnection = InvalidConn
		InvalidRS.LockType = 3	'adLockOptimistic
		InvalidRS.Open "InvalidLogins"
		InvalidRS.AddNew
		InvalidRS("DomainName") = Request.ServerVariables("SERVER_NAME")
		InvalidRS("Location") = epl_removeinvalidchars(Request.Form("UserName"))
		InvalidRS("Username") = Request.Form("UserName")
		InvalidRS("Password") = Request.Form("Password")
		InvalidRS("IPAddress") = Request.ServerVariables("REMOTE_ADDR")
		InvalidRS.Update
		InvalidRS.Close
		Set InvalidRS = Nothing
		InvalidConn.Close
		Set InvalidConn = Nothing
	Else
		'	now validate that the Allow Access flag is set to true
		'		if objRS("AllowAccess") <> True then
		'		strError = "<font face=""Verdana"" size=""2"" color=""#ff0000"">ERROR: Access Denied!</font><br><br>"
		'		end if

		if objRS("FromUSAWS") = True then
			'response.write "xxx" & len("TournamentDate") & "yyyy"
			'response.end
			session("FromUSAWS") = objRS("FromUSAWS")
			session("TournamentDate") = objRS("TournamentDate")
			session("TournamentName") = objRS("TournamentName")
			Session("TournamentYear") = (2000 + left(objRS("Name"),2))
			
			
			if len(objRS("TournamentDate")) > 0 then
				Dim TournamentMonths 
				TournamentMonths = datediff("m",objRS("TournamentDate"),date())
				if TournamentMonths > 1 then
					strError = "<font face=""Verdana"" size=""2"" color=""#ff0000"">ERROR: Expired UserID!</font><br><br>"
				end if
			else
				strError = "<font face=""Verdana"" size=""2"" color=""#ff0000"">ERROR: No Tournament Date!</font><br><br>"
			end if
		end if
		if len(strError) = 0 then
			Session("aauth") = True
			Session("UserName") = objRS("Name")
			Session("TournamentID") = Left(Session("UserName"),6)
			Session("FullName") = objRS("FullName")
			Session("UserID") = objRS("UserID")
			IF left(Session("UserName"),1) > "9" OR len(session("TournamentDate")) = 0 THEN
				session("TournamentID") = ""
				session("TournamentDate") = date() + 2
				session("TournamentName") = "Administrator: " & objRS("FullName")
				IF datepart("y",date()) < 225 then yearadd = 0 else yearadd = 1
				session("TournamentYear") = datepart("yyyy",date()) + yearadd
			END IF
			Session.TimeOut = 60
					
			objRS("DateLastLogin") = Now()
			objRS.Update
			objRS.Close
			Response.Redirect "/admin/index.asp"
		end if		
	End If
	Set objRS = Nothing
	
End If
%>
<form action="login.asp" method="post">

            <center><b><i><br><H2><%= strError %></H2></i></b></center>
			  
        <table border="0" cellspacing="0" cellpadding="1" align="center">
        	<tr>

          	<td bgcolor="<%= thColor %>" width=250>

              <table border="0" cellspacing="0" cellpadding="6">
                  <tr bgcolor="<%= thColor %>"> 
                        
                    <td colspan="2" align="center"><font face="Verdana" size="3"><b><font color="#000000">Administrator's 
                    Login</font></b></font></td>
                 </tr>

                 <tr bgcolor="<%= tdCol1 %>"> 
                        
                   <td><font face="Verdana" size="2">Username: </font></td>
                   <td><input type="text" name="UserName" style="background-color: #ffffff; border: 1px solid <%= thColor %>; font-family: Verdana; font-size: 12px;" value="<%= FormUserName %>">
                       </td>
                 </tr>

                 <tr bgcolor="<%= tdCol2 %>"> 
                   <td><font face="Verdana" size="2">Password:</font></td>
                   <td><input type="password" name="Password" style="background-color: #ffffff; border: 1px solid <%= thColor %>; font-family: Verdana; font-size: 12px;" value="<%= Request.Cookies("ADirectory")("Password") %>">
                       </td>
                 </tr>

<% If Application("UseSave") then %>

                 <tr bgcolor="<%= tdCol1 %>"> 
                   <td colspan="2"> 
                     <input type="checkbox" name="SaveInfo" value="save"<%= Request.Cookies("ADirectory")("checked") %>>
                       <font face="Verdana" size="2">Save login information on this computer.</font>
                   </td>
                 </tr>
 
                 <tr bgcolor="<%= tdCol2 %>"> 
<% Else %>
                 <tr bgcolor="<%= tdCol1 %>"> 
<% End If %>
                   <td colspan="2" align="center"> 
                     <input type="image" border="0" name="enter" src="/templates/images/enter-button.gif" width="60" height="25">
                   </td>
                 </tr>
               </table>

             </td>

            <td width=100>
            	&nbsp;
            </td>

            <td width=220>
            	<div align="center">
            	<font color="#FF0000" size="3" face="Verdana, Arial, Helvetica, sans-serif">
            	<strong>IMPORTANT!</strong>
            	</font>

              <font face="Verdana, Arial, Helvetica, sans-serif">
              	
              <p><font size="2">To login to this system, you must have a valid User Name and 
                Password, provided by the Competition Department of USA Water Ski Headquarters.</font></p>
              </div>
            </td>
            
           </tr></table>

        <table width="100%" border="0" cellspacing="1" cellpadding="1">

          <tr>
            <td>&nbsp;&nbsp;&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;&nbsp;&nbsp;</td>
          </tr>

          <tr>

            <td>&nbsp;</td>

            <td><font face="Verdana, Arial, Helvetica, sans-serif">
              	
              <p><font color="#FF0000" size="2"><strong>SWIFT USERS (AWSA, NCWSA, 
                ABC, AKA and USA-WB):</strong></font><font size="2"> If you sanctioned an 
                event through our online sanctioning program (SWIFT), the following 
                information will be used as your user name and password:<br>
                &#8226; <font color="#FF0000">Tournament ID</font> (user name); 
                and<br>
                &#8226; <font color="#FF0000">Edit Code</font> (password)</font></p>
              <p><font size="2">The Tournament ID and Edit Code were provided through the 
              	online sanctioning program (SWIFT), at the time your tournament sanction
              	was first being set up.</font></p>
              </font><p><font size="2" face="Verdana, Arial, Helvetica, sans-serif">If you do not 
              	have the Tournament ID and Edit Code you need, please contact:<br>
                &#8226; Event Organizer;&nbsp; or<br>
                &#8226; Traci Baylon, USA Water Ski Headquarters<br>
                <a href="mailto:tbaylon@usawaterski.org?subject=Edit Code Needed">tbaylon@usawaterski.org</a><br>
                Phone 1-800-533-2972, Ext. 121</font></p>
              
              <p><font size="2" face="Verdana, Arial, Helvetica, sans-serif">NOTE: Access 
                to login to this system is only granted for future events that have been 
                approved for publication through SWIFT.</font><br>
              </p></td>
            <td>&nbsp;</td>

          </tr>

          <tr>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>

        </table>

      </form>

	
	
	</td>
  </tr>
</table>
</body>
</html>








