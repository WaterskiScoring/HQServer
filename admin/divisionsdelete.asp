<!--#include virtual="/admin/includes/security.asp" -->
<%
If Len(Request.QueryString("id")) = 0 Then
  Response.Redirect("/admin/divisionsadmin.asp")
End If
Dim SQLConnect, StrSql, RS
Set SQLConnect = CreateObject("ADODB.Connection")
SQLConnect.Open Application("WaterSkiConn")
StrSql = "SELECT * FROM Divisions WHERE DivisionID=" & Request.QueryString("id")
Set RS = SQLConnect.execute(StrSql)
If RS.Eof Then
  Response.Redirect("/admin/divisionsadmin.asp")
End If
If Request.Form("submit") = "Delete" Then
  StrSQL = "SELECT * FROM [Additional Divisions to Show on Registration Template] WHERE (PrimaryDivisionID=" & Request.QueryString("id") & ") OR (AdditionalDivisionID=" & Request.QueryString("id") & ")"
  Set CheckRS = SQLConnect.execute(StrSql)
  If CheckRS.eof Then
    StrSql = "DELETE FROM Divisions WHERE DivisionID=" & Request.QueryString("id")
    SQLConnect.execute(strSql)
    Response.Redirect("/admin/divisionsadmin.asp")
  Else
    StrMessage = "Please remove division dependencies before deletion"
  End If
End If
%>
<html>

<head>
<title>Admin Index</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" >
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
    <td valign="top" ><table width="100%" border="0" cellspacing="1" cellpadding="1">
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
If Len(StrMessage) > 0 Then
%>
          <p><font face="Verdana" size="2" color="#FF0000"><%= StrMessage %>!</font></p>
<% End If %>
          <form method="post" action="/admin/divisionsdelete.asp?id=<%= RS("DivisionID") %>">
              <table border="0" cellspacing="0" cellpadding="6">
                <tr align="center">
                  <td colspan="2" bgcolor="<%= thColor %>"><font face="Verdana" size="2" color="#FFFFFF"><b>Delete Division</b></font></td>
                </tr>
                <tr bgcolor="<%= tdCol1 %>">
                  <td><font face="Verdana" size="2" color="#000000">Division Name:</font></td>
                  <td><font face="Verdana" size="2" color="#000000"><%= RS("DIV") %></font></td>
                </tr>
                <tr bgcolor="<%= tdCol2 %>">
                  <td><font face="Verdana" size="2" color="#000000">Division Description:</font></td>
                  <td><font face="Verdana" size="2" color="#000000"><%= RS("DivisionDescription") %></font></td>
                </tr>
                <tr bgcolor="<%= thColor %>" align="center">
                  <td colspan="2"><input type="submit" name="submit" value="Delete">
                    <input type="button" name="cancel" value="Cancel" onClick="goToURL('/admin/divisionsadmin.asp')">
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





