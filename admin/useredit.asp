<!--#include virtual="/admin/includes/security.asp" -->
<%

If not Session("aauth") then response.redirect "Login.asp"

Dim objConn1
Set objConn1 = Server.CreateObject("ADODB.Connection")
objConn1.Open Application("WaterSkiConn")

If Request.Form("submit") = "Cancel" then Response.Redirect "useradmin.asp"

If Request.Form("submit") = "Save" then

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


		Dim objUpdRS
		Set objUpdRS = Server.CreateObject("ADODB.RecordSet")
		objUpdRS.ActiveConnection = objConn1
		objUpdRS.LockType = 3	'adLockOptimistic
		objUpdRS.Open "SELECT * FROM Users999 WHERE UserID = '" & Request("UserID") & "'"
		objUpdRS("FullName") = Request.Form("FullName")
		objUpdRS("Name") = Request.Form("UserName")
		objUpdRS("Pwd") = Request.Form("Password")
		objUpdRS("EmailAddress") = Request.Form("EmailAddress")


		if Request.Form("DownloadMembers1") = "true" then
			objUpdRS("DownloadMembers1") = True
		else
			objUpdRS("DownloadMembers1") = False
		end if

		if Request.Form("downloadDBF") = "true" then
			objUpdRS("downloadDBF") = True
		else
			objUpdRS("downloadDBF") = False
		end if

		if Request.Form("CreateRegistrationTemplate") = "true" then
			objUpdRS("CreateRegistrationTemplate") = True
		else
			objUpdRS("CreateRegistrationTemplate") = False
		end if

		if Request.Form("adminUsers") = "true" then
			objUpdRS("adminUsers") = True
		else
			objUpdRS("adminUsers") = False
		end if

		if Request.Form("adminDivisions") = "true" then
			objUpdRS("EditDivisions") = True
		else
			objUpdRS("EditDivisions") = False
		end if

  	objUpdRS.Update
		objUpdRS.Close
		Set objUpdRS = Nothing
		Response.Redirect "/admin/useradmin.asp"
	
	End If
End If

'Now get the initial values
If len(Request.Form) = 0 then
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
	adminDivisions = objRS("EditDivisions")
	objRS.close
else
	UserName =  Request.Form("UserName")
	Name = Request.Form("FullName")
	Password = Request.Form("Password")
	EmailAddress = Request.Form("EmailAddress")
	DownloadMembers1 = Request.Form("DownloadMembers1")
	downloadDBF = Request.Form("downloadDBF") 
	CreateRegistrationTemplate = Request.Form("CreateRegistrationTemplate")
	adminUsers = Request.Form("adminUsers")
  adminDivisions = Request.Form("adminDivisions") 
end if

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
<title>Admin Users - Add</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" background = "/templates/images/TopBackground.jpg" link="#000000" vlink="#000000" alink="#000000" leftMargin=0 topMargin=0 marginwidth="0" marginheight="0" >
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
<form name="form1" method="post" action="/admin/useredit.asp?UserID=<%= Request("UserID") %>" onSubmit="return checkPw(this)">
              <table border="0" cellspacing="0" cellpadding="6">
                <tr align="center"> 
                  <td colspan="2" bgcolor="<%= thColor %>"><font face="Verdana" size="2" color="#FFFFFF"><b>Edit 
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
                <tr bgcolor="<%= tdCol2 %>">
                  <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Download
                    Membership File</font></td>
                  <td><input name="DownloadMembers1" type="checkbox" id="DownloadMembers1" value="true" <% if DownloadMembers1 = True then %>checked<% end if %>></td>
                </tr>
                <tr bgcolor="<%= tdCol1 %>">
                  <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Download
                    Membership DBF</font></td>
                  <td><input name="downloadDBF" type="checkbox" id="downloadDBF" value="true" <% if downloadDBF = True then %>checked<% end if %>></td>
                </tr>
                <tr bgcolor="<%= tdCol2 %>">
                  <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Create
                    Registration Template</font></td>
                  <td><input name="CreateRegistrationTemplate" type="checkbox" id="CreateRegistrationTemplate" value="true" <% if CreateRegistrationTemplate = True then %>checked<% end if %>></td>
                </tr>
                <tr bgcolor="<%= tdCol1 %>">
                  <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Admin
                    Users</font></td>
                  <td><input name="adminUsers" type="checkbox" id="adminUsers" value="true" <% if adminUsers = True then %>checked<% end if %>></td>
                </tr>
                <tr bgcolor="<%= tdCol2 %>">
                  <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Admin
                    Divisions</font></td>
                  <td><input name="adminDivisions" type="checkbox" id="adminDivisions" value="true" <% if adminDivisions = True then %>checked<% end if %>></td>
                </tr>
                <tr bgcolor="<%= thColor %>" align="center">
                  <td colspan="2"> <input name="submit" type="submit" id="submit" value="Save">
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





