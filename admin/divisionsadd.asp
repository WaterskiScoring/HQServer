<!--#include virtual="/admin/includes/security.asp" -->
<%
If Request.Form("submit") = "Add" Then
  'validate entries
  Dim StrMessage
  If Len(Request.Form("DivisionName")) = 0 Then
    StrMessage = "Please enter a Division Name"
  ElseIf Len(Request.Form("DivisionDescription")) = 0 Then
    StrMessage = "Please enter a Division Description"
  End If
  If Len(StrMessage) = 0 Then
    'check for duplicate entry
    Dim SQLConnect, StrSql, RS
    Set SQLConnect = CreateObject("ADODB.Connection")
    SQLConnect.Open Application("WaterSkiConn")
    StrSql = "SELECT * FROM Divisions WHERE DIV='" & Left(Request.Form("DivisionName"),2) & "'"
    Set RS = SQLConnect.execute(StrSql)
    If Not RS.Eof Then
      'Division already exists with supplied name
      StrMessage = "A Division with that name already exists"
    Else
      'unique Division - add to DB
      Dim AddRS
      Set AddRS = Server.CreateObject("ADODB.RecordSet")
      AddRS.ActiveConnection = SQLConnect
      AddRS.LockType = 3
      AddRS.Open "Divisions"
      AddRS.AddNew
      AddRS("DIV") = Left(Request.Form("DivisionName"),2)
      AddRS("DivisionDescription") = Request.Form("DivisionDescription")
      AddRS.Update
      AddRS.Close
      Set AddRS = Nothing
      Response.Redirect("/admin/divisionsadmin.asp")
    End If
  End If
End If
%><html>

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
thColor= "#999999"
tdCol1 = "#42639F"
tdCol2 = "#FFFFFF"
If Len(StrMessage) > 0 Then
%>
          <p><font face="Verdana" size="2" color="#FF0000"><%= StrMessage %>!</font></p>
<% End If %>
          <form method="post" action="/admin/divisionsadd.asp">
              <table border="0" cellspacing="0" cellpadding="6">
                <tr align="center">
                  <td colspan="2" bgcolor="<%= thColor %>"><font face="Verdana" size="2" color="#FFFFFF"><b>Add
                    New Division</b></font></td>
                </tr>
                <tr bgcolor="<%= tdCol1 %>">
                  <td><font face="Verdana" size="2" color="#000000">Division Name:</font></td>
                  <td><input type="text" name="DivisionName" style="background-color: #ffffff; border: 1px solid <%= thColor %>; font-family: Verdana; font-size: 12px;" value="<%= Left(Request.Form("DivisionName"),2) %>"></td>
                </tr>
                <tr bgcolor="<%= tdCol2 %>">
                  <td><font face="Verdana" size="2" color="#000000">Division Description:</font></td>
                  <td><input type="text" name="DivisionDescription" style="background-color: #ffffff; border: 1px solid <%= thColor %>; font-family: Verdana; font-size: 12px;" value="<%= Request.Form("DivisionDescription") %>"></td>
                </tr>
                <tr bgcolor="<%= thColor %>" align="center">
                  <td colspan="2"><input type="submit" name="submit" value="Add">
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





