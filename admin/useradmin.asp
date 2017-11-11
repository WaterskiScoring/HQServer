<!--#include virtual="/admin/includes/security.asp" -->

<% If not Session("aauth") then response.redirect "Login.asp" %>

<html>

<head>
<title>Admin Index</title>

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
Dim col1
Dim col2
Dim currentcolor
col1 = "#42639F"
col2 = "#FFFFFF"
currentcolor = col1

Dim objRS
Set objRS = Server.CreateObject("ADODB.RecordSet")
objRS.ActiveConnection = objConn
objRS.Open "SELECT * FROM Users999 where FromUSAWS = 0"
%>
            <table border="0" cellspacing="0" cellpadding="6">
              <tr bgcolor="#999999"> 
                <td><b><font face="Verdana" size="2" color="#FFFFFF">Full Name</font></b></td>
                <td><b><font face="Verdana" size="2" color="#FFFFFF">User Name</font></b></td>
                <td><b><font face="Verdana" size="2" color="#FFFFFF">Email Address</font></b></td>
                <td><b><font face="Verdana" size="2" color="#FFFFFF">Action</font></b></td>
              </tr>
              <%


Do while not objRS.EOF
%>
              <tr bgcolor="<%= currentcolor %>"> 
                <td><font face="Verdana" size="2"><%= objRS("FullName") %></font></td>
                <td><font face="Verdana" size="2"><%= objRS("Name") %></font></td>
                <td><font face="Verdana" size="2"><a href="mailto:<%= objRS("EmailAddress") %>"><%= objRS("EmailAddress") %></a></font></td>
                <td><font face="Verdana" size="1"><a href="/admin/useredit.asp?UserID=<%= objRS("UserID") %>">edit</a> / <a href="/admin/userdelete.asp?UserID=<%= objRS("UserID") %>">delete</a></font></td>
              </tr>
              <%
	objRS.MoveNext
	If currentcolor = col1 then
		currentcolor = col2
	Else
		currentcolor = col1
	End If
Loop
%>
              <tr> 
                <td colspan="4" bgcolor="999999"><font face="Verdana" size="2"><a href="/admin/useradd.asp">Add 
                  a New User</a></font></td>
              </tr>
            </table>
			
			
		  </td>
          <td>&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>
</body>
</html>





