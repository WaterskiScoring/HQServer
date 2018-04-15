<!--#include virtual="/admin/includes/security.asp" -->
<html>

<head>
<title>Admin Index</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" background = "/templates/images/TopBackground.jpg" link="#000000" vlink="#000000" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" >
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
Dim SQLConnect, StrSql, RS
Set SQLConnect = CreateObject("ADODB.Connection")
SQLConnect.Open Application("WaterSkiConn")
StrSql = "SELECT * FROM Divisions ORDER BY DIV"
Set RS = SQLConnect.execute(StrSql)
%>
            <table border="0" cellspacing="0" cellpadding="6">
              <tr bgcolor="#999999">
                <td><b><font face="Verdana" size="2" color="#FFFFFF">Division</font></b></td>
                <td><b><font face="Verdana" size="2" color="#FFFFFF">Description</font></b></td>
                <td><b><font face="Verdana" size="2" color="#FFFFFF">Action</font></b></td>
              </tr>
<%
If RS.Eof Then
%>
              <tr>
                <td colspan="3">No Divisions currently exist!</td>
              </tr>
<%
Else
  Dim col1
  Dim col2
  Dim currentcolor
  col1 = "#42639F"
  col2 = "#FFFFFF"
  currentcolor = col1
  While Not RS.Eof
%>
              <tr bgcolor="<%= currentcolor %>">
                <td><font face="Verdana" size="2"><%= RS("DIV") %></font></td>
                <td><font face="Verdana" size="2"><%= RS("DivisionDescription") %></font></td>
                <td><font face="Verdana" size="1"><a href="/admin/divisionsedit.asp?id=<%= RS("DivisionID") %>">edit</a>/ <a href="/admin/divisionsdelete.asp?id=<%= RS("DivisionID") %>">delete</a></font></td>
              </tr>
<%
  	If currentcolor = col1 then
  		currentcolor = col2
  	Else
  		currentcolor = col1
  	End If
    RS.MoveNext
  Wend
%>
              <tr>
                <td colspan="4" bgcolor="999999"><font face="Verdana" size="2"><a href="/admin/divisionsadd.asp">Add
                  a New Division</a></font></td>
              </tr>
<%
End If
%>
            </table>
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





