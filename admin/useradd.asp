<!--#include virtual="/admin/includes/security.asp" -->
<%

If not Session("aauth") then response.redirect "Login.asp"

If Request.Form("submit") = "Cancel" then Response.Redirect "useradmin.asp"

If Request.Form("submit") = "Add" then
	bOKToContinue = TRUE
	if Request.Form("UserName") = "" then
		response.write "<font color = ""red"">No user name entered.  You will need to correct this to continue.</font><br>"
		bOKToContinue = FALSE
	end if
	if Request.Form("Password") <> Request.Form("Password2") then
		response.write "<font color = ""red"">Passwords do not match.  You will need to correct this to continue.</font><br>"
		bOKToContinue = FALSE
	end if
	if Request.Form("Password") = "" then
		response.write "<font color = ""red"">No password entered.  You will need to correct this to continue.</font><br>"
		bOKToContinue = FALSE
	end if

	if bOKToContinue = TRUE	then 
		Dim objConn1
		Set objConn1 = Server.CreateObject("ADODB.Connection")
		objConn1.Open Application("WaterSkiConn")
		Dim objRS
		Set objRS = Server.CreateObject("ADODB.RecordSet")
		objRS.ActiveConnection = objConn1
		objRS.LockType = 3	'adLockOptimistic
		objRS.Open "Users999"
		objRS.AddNew
	
		objRS("FullName") = Request.Form("FullName")
		objRS("Name") = Request.Form("UserName")
		objRS("Pwd") = Request.Form("Password")
		objRS("EmailAddress") = Request.Form("EmailAddress")
		if Request.Form("DownloadMembers1") = "true" then objRS("DownloadMembers1") = True
		if Request.Form("downloadDBF") = "true" then objRS("downloadDBF") = True
		if Request.Form("CreateRegistrationTemplate") = "true" then objRS("CreateRegistrationTemplate") = True
		if Request.Form("adminUsers") = "true" then objRS("adminUsers") = True
		objRS("FromUSAWS") = False
		objRS.Update
		objRS.Close
		Set objRS = Nothing
		Response.Redirect "/admin/useradmin.asp"
	end if
End If
%>
<script language="JavaScript">
<!--
function checkPw(theform) {

if (theform.Password.value != theform.Password2.value ) {
	alert("Passwords entered do not match.  This must be corrected to continue.");
	theform.Password.focus();
	return false;
}else {
	return true;
}
}
// -->
</script>

<html>

<head>
<title>Admin Users</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" leftMargin=0 topMargin=0 marginwidth="0" marginheight="0" >
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="182" bgcolor="#42639F" valign="top"></td>
    <td bgcolor="#42639F">
      <p>&nbsp;</p>
      <p align="center"><font face="Verdana" size="6" color="#FFFFFF">USA Water 
        Ski Admin</font></p>
      <p>&nbsp;</p>
    </td>
  </tr>
  <tr> 
   <td bgcolor="#42639F" valign="top">
<!--#include virtual="/admin/includes/menu.asp" -->
  </td>
    <td valign="top"><table width="100%" border="0" cellspacing="1" cellpadding="1">
        <tr>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td valign="top"> 
		  <% 
		  thColor= "#999999" 
			tdCol1 = "#42639F"
			tdCol2 = "#FFFFFF"
		  %>
<form name="form1" method="post" action="/admin/useradd.asp" onSubmit="return checkPw(this)">
              <table border="0" cellspacing="0" cellpadding="6">
                <tr align="center"> 
                  <td colspan="2" bgcolor="<%= thColor %>"><font face="Verdana" size="2" color="#FFFFFF"><b>Add 
                    New User</b></font></td>
                </tr>
                <tr bgcolor="<%= tdCol1 %>"> 
                  <td><font face="Verdana" size="2">Full Name:</font></td>
                  <td> <input type="text" name="FullName" style="background-color: #ffffff; border: 1px solid <%= thColor %>; font-family: Verdana; font-size: 12px;" value="<%= Request.Form("FullName") %>"> 
                  </td>
                </tr>
                <tr bgcolor="<%= tdCol2 %>"> 
                  <td><font face="Verdana" size="2">User Name:</font></td>
                  <td> <input type="text" name="UserName" style="background-color: #ffffff; border: 1px solid <%= thColor %>; font-family: Verdana; font-size: 12px;" value="<%= Request.Form("UserName") %>"> 
                  </td>
                </tr>
                <tr bgcolor="<%= tdCol1 %>"> 
                  <td><font face="Verdana" size="2">Email Address:</font></td>
                  <td> <input type="text" name="EmailAddress" style="background-color: #ffffff; border: 1px solid <%= thColor %>; font-family: Verdana; font-size: 12px;" value="<%= Request.Form("EmailAddress") %>"> 
                  </td>
                </tr>
                <tr bgcolor="<%= tdCol2 %>"> 
                  <td><font face="Verdana" size="2">Password:</font></td>
                  <td> <input type="password" name="Password" style="background-color: #ffffff; border: 1px solid <%= thColor %>; font-family: Verdana; font-size: 12px;"> 
                  </td>
                </tr>
                <tr bgcolor="<%= tdCol1 %>"> 
                  <td><font face="Verdana" size="2">Password: (confirm)</font></td>
                  <td> <input type="password" name="Password2" style="background-color: #ffffff; border: 1px solid <%= thColor %>; font-family: Verdana; font-size: 12px;"> 
                  </td>
                </tr>
                <tr bgcolor="<%= tdCol1 %>"> 
                  <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Download 
                    Membership File</font></td>
                  <td><input name="DownloadMembers1" type="checkbox" id="DownloadMembers1" value="true"></td>
                </tr>
                <tr bgcolor="<%= tdCol1 %>"> 
                  <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Download 
                    Membership DBF</font></td>
                  <td><input name="downloadDBF" type="checkbox" id="downloadDBF" value="true"></td>
                </tr>
                <tr bgcolor="<%= tdCol1 %>"> 
                  <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Create 
                    Registration Template</font></td>
                  <td><input name="CreateRegistrationTemplate" type="checkbox" id="CreateRegistrationTemplate" value="true"></td>
                </tr>
                <tr bgcolor="<%= tdCol1 %>"> 
                  <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Admin 
                    Users</font></td>
                  <td><input name="adminUsers" type="checkbox" id="adminUsers" value="true"></td>
                </tr>
                <tr bgcolor="<%= thColor %>" align="center"> 
                  <td colspan="2"> <input type="submit" name="submit" value="Add"> 
                    <input name="submit" type="submit" id="submit" value="Cancel"> 
                  </td>
                </tr>
              </table>
			</form>            

          </td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>





