<!--#include virtual="/admin/includes/security.asp" -->

<%

If not Session("aauth") then response.redirect "Login.asp"

Dim objConn1
Set objConn1 = Server.CreateObject("ADODB.Connection")
objConn1.Open Application("WaterSkiConn")

If Request.Form("submit") = "Cancel" then Response.Redirect "useradmin.asp"

If Request.Form("submit") = "Delete" then
	


		Dim objUpdRS
		Set objUpdRS = Server.CreateObject("ADODB.RecordSet")
		objUpdRS.ActiveConnection = objConn1
		objUpdRS.LockType = 3	'adLockOptimistic
		objUpdRS.Open "Delete FROM Users999 WHERE UserID = '" & Request("UserID") & "'"
		Set objUpdRS = Nothing
		Response.Redirect "/admin/useradmin.asp"
	

End If

'Now get the initial values

	'user past value
	Dim objRS
	Set objRS = Server.CreateObject("ADODB.RecordSet")
	objRS.ActiveConnection = objConn1
	objRS.Open "SELECT * FROM Users999 WHERE UserID = '" & Request("UserID") & "'"
	UserName = objRS("Name")
	FullName = objRS("FullName") 
	Name = objRS("Name") 
	Password = objRS("Pwd") 
	EmailAddress = objRS("EmailAddress")
	DownloadMembers1 = objRS("DownloadMembers1")
	downloadDBF = objRS("downloadDBF") 
	CreateRegistrationTemplate = objRS("CreateRegistrationTemplate")
	adminUsers = objRS("adminUsers") 
	objRS.close	 

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
<title>Admin Users - Delete</title>

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
<form name="form1" method="post" action="/admin/userdelete.asp?UserID=<%= Request("UserID") %>" onSubmit="return checkPw(this)">
              <table border="0" cellspacing="0" cellpadding="6">
                <tr align="center"> 
                  <td colspan="2" bgcolor="<%= thColor %>"><font face="Verdana" size="2" color="#FFFFFF"><b>Delete 
                    User</b></font></td>
                </tr>
                <tr bgcolor="<%= tdCol1 %>"> 
                  <td><font face="Verdana" size="2">Full Name:</font></td>
                  <td> <input type="text" name="FullName" style="background-color: #ffffff; border: 1px solid <%= thColor %>; font-family: Verdana; font-size: 12px;" value="<%= FullName %>"> 
                  </td>
                </tr>
                <tr bgcolor="<%= tdCol2 %>"> 
                  <td><font face="Verdana" size="2">User Name:</font></td>
                  <td> <input type="text" name="UserName" style="background-color: #ffffff; border: 1px solid <%= thColor %>; font-family: Verdana; font-size: 12px;" value="<%= UserName %>"> 
                  </td>
                </tr>
                <tr bgcolor="<%= tdCol1 %>"> 
                  <td><font face="Verdana" size="2">Email Address:</font></td>
                  <td> <input type="text" name="EmailAddress" style="background-color: #ffffff; border: 1px solid <%= thColor %>; font-family: Verdana; font-size: 12px;" value="<%= EmailAddress %>"> 
                  </td>
                </tr>
                <tr bgcolor="<%= tdCol2 %>"> 
                  <td><font face="Verdana" size="2">Password:</font></td>
                  <td> <input name="Password" type="password" style="background-color: #ffffff; border: 1px solid <%= thColor %>; font-family: Verdana; font-size: 12px;" value="<%= Password %>"> 
                  </td>
                </tr>
                <tr bgcolor="<%= tdCol1 %>"> 
                  <td><font face="Verdana" size="2">Password: (confirm)</font></td>
                  <td> <input name="Password2" type="password" style="background-color: #ffffff; border: 1px solid <%= thColor %>; font-family: Verdana; font-size: 12px;" value="<%= Password %>"> 
                  </td>
                </tr>
                <tr bgcolor="<%= tdCol1 %>"> 
                  <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Download 
                    Membership File</font></td>
                  <td><input name="DownloadMembers1" type="checkbox" id="DownloadMembers1" value="true" <% if DownloadMembers1 = True then %>checked<% end if %>></td>
                </tr>
                <tr bgcolor="<%= tdCol1 %>"> 
                  <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Download 
                    Membership DBF</font></td>
                  <td><input name="downloadDBF" type="checkbox" id="downloadDBF" value="true" <% if downloadDBF = True then %>checked<% end if %>></td>
                </tr>
                <tr bgcolor="<%= tdCol1 %>"> 
                  <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Create 
                    Registration Template</font></td>
                  <td><input name="CreateRegistrationTemplate" type="checkbox" id="CreateRegistrationTemplate" value="true" <% if CreateRegistrationTemplate = True then %>checked<% end if %>></td>
                </tr>
                <tr bgcolor="<%= tdCol1 %>"> 
                  <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Admin 
                    Users</font></td>
                  <td><input name="adminUsers" type="checkbox" id="adminUsers" value="true" <% if adminUsers = True then %>checked<% end if %>></td>
                </tr>
                <tr bgcolor="<%= thColor %>" align="center"> 
                  <td colspan="2"> <input name="submit" type="submit" id="submit" value="Delete"> 
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





