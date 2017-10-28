<!--#include virtual="/admin/includes/security.asp" -->
<%
Dim SQLConnect, StrSql, RS
Set SQLConnect = CreateObject("ADODB.Connection")
SQLConnect.Open Application("WaterSkiConn")
If Request.Form("submit") = "Update" Then
  'Update Division information (Name, Decsription)
  'validate entries
  If Len(Request.Form("DivisionName")) = 0 Then
    StrMessage = "Please enter a Division Name"
  ElseIf Len(Request.Form("DivisionDescription")) = 0 Then
    StrMessage = "Please enter a Division Description"
  End If
  If Len(StrMessage) = 0 Then
    'check for duplicate entry
    StrSql = "SELECT * FROM Divisions WHERE DIV='" & Left(Request.Form("DivisionName"),2) & "' AND DivisionID<>" & Request.Form("id")
    'Check for Dupicate -RS
    Set cRS = SQLConnect.execute(StrSql)
    If Not cRS.eof Then
      'A duplicate exists
      StrMessage = "A Division with that name already exists"
    Else
      'Division entry is unique - add to DB
      Dim UpdateRS
      Set UpdateRS = Server.CreateObject("ADODB.RecordSet")
      UpdateRS.ActiveConnection = SQLConnect
      UpdateRS.LockType = 3
      UpdateRS.Open "SELECT * FROM Divisions WHERE DivisionID=" & Request.Form("id")
      UpdateRS("DIV") = Left(Request.Form("DivisionName"),2)
      UpdateRS("DivisionDescription") = Request.Form("DivisionDescription")
      UpdateRS.Update
      UpdateRS.Close
      Set UpdateRS = Nothing
      Response.Redirect("/admin/divisionsadmin.asp")
    End If
  End If
ElseIf Request.Form("submit") = "Add" Then
  'Add AdditionalDivision for current Division
  Dim AddDivRS
  Set AddDivRS = Server.CreateObject("ADODB.RecordSet")
  AddDivRS.ActiveConnection = SQLConnect
  AddDivRS.LockType = 3
  AddDivRS.Open "[Additional Divisions to Show on Registration Template]"
  AddDivRS.AddNew
  AddDivRS("PrimaryDivisionID") = Request.QueryString("id")
  AddDivRS("AdditionalDivisionID") = Request.Form("Divisions")
  AddDivRS.Update
  AddDivRS.Close
  Set AddDivRS = Nothing
ElseIf Request.Form("submit") = "Remove" Then
  'check to see if a selection has been made
  If Len(Request.Form("AdditionalDivision")) = 0 Then
    StrMessage = "Please select Divisions to remove"
  End If
  If Len(StrMessage) = 0 Then
    'if no error has occurred remove AdditionalDivision from Division
    Dim DivisionsToRemove, NumberToRemove
    DivisionsToRemove = Split(Request.Form("AdditionalDivision"),", ")
    NumberToRemove = UBound(DivisionsToRemove)
    For x = 0 To NumberToRemove
      StrSql = "DELETE FROM [Additional Divisions to Show on Registration Template] WHERE PrimaryDivisionID=" & Request.QueryString("id") & " AND AdditionalDivisionID=" & DivisionsToRemove(x)
      SQLConnect.execute(StrSql)
    Next
  End If
End If
If Len(Request.QueryString("id")) = 0 Then
  'No Division to edit defined
  Response.Redirect("/admin/divisionsadmin.asp")
End If
'Get Division info for edit
StrSql = "SELECT * FROM Divisions WHERE DivisionID=" & Request.QueryString("id")
Set RS = SQLConnect.execute(StrSql)
If RS.eof Then
  Response.Redirect("/admin/divisionsadmin.asp")
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
          <form method="post" action="/admin/divisionsedit.asp?id=<%= Request.QueryString("id") %>">
              <table border="0" cellspacing="0" cellpadding="6">
                <tr align="center">
                  <td colspan="2" bgcolor="<%= thColor %>"><font face="Verdana" size="2" color="#FFFFFF"><b>Edit Division</b></font></td>
                </tr>
                <tr bgcolor="<%= tdCol1 %>">
                  <td><font face="Verdana" size="2" color="#000000">Division Name:</font></td>
                  <td><input type="text" name="DivisionName" style="background-color: #ffffff; border: 1px solid <%= thColor %>; font-family: Verdana; font-size: 12px;" value="<%= RS("DIV") %>"></td>
                </tr>
                <tr bgcolor="<%= tdCol2 %>">
                  <td><font face="Verdana" size="2" color="#000000">Division Description:</font></td>
                  <td><input type="text" name="DivisionDescription" style="background-color: #ffffff; border: 1px solid <%= thColor %>; font-family: Verdana; font-size: 12px;" value="<%= RS("DivisionDescription") %>"></td>
                </tr>
                <tr bgcolor="<%= thColor %>" align="center">
                  <td colspan="2"><input type="submit" name="submit" value="Update">
                    <input type="button" name="cancel" value="Cancel" onClick="goToURL('/admin/divisionsadmin.asp')">
                    <input type="hidden" name="id" value="<%= RS("DivisionID") %>">
                  </td>
                </tr>
                <tr>
                  <td colspan="2">&nbsp;</td>
                </tr>
                <tr align="center">
                  <td colspan="2" bgcolor="<%= thColor %>"><font face="Verdana" size="2" color="#FFFFFF"><b>Divisions to include with <%= RS("DIV") %></b></font></td>
                </tr>
                <tr>
                  <td colspan="2" bgcolor="<%= tdCol1 %>">
<%
'Get AdditionalDivisions for current Divisions
StrSql = "SELECT * FROM AdditionalDivisions WHERE PrimaryDivisionID=" & RS("DivisionID") & " ORDER BY DIV"
'Additional Divisions -RS
Set aRS = SQLConnect.execute(StrSql)
If aRS.eof Then
%>
                    <font face="Verdana" size="2" color="#000000">No Divisions to include!</font>
<%
Else
%>
                    <select name="AdditionalDivision" multiple="multiple" style="background-color: #ffffff; border: 1px solid <%= thColor %>; font-family: Verdana; font-size: 12px;"><%
  Dim CurDiv
  While Not aRS.eof
%>
                    <option value="<%= aRS("AdditionalDivisionID") %>"><%= aRS("DIV") %> - <%= aRS("DivisionDescription") %></option>
<%
    'Pass AdditionalDivisionID to variable to check later
    CurDiv = CurDiv & "," & aRS("AdditionalDivisionID")
    aRS.movenext
  Wend
%>
                    </select>
                    &nbsp;<input type="submit" name="submit" value="Remove" />
<%
End If
%>
                  </td>
                </tr>
                <tr>
                  <td colspan="2" bgcolor="<%= thColor %>">
<%
'Get all divisions to allow for addiing additional divisions under current division
StrSql = "SELECT * FROM Divisions ORDER BY DIV"
'Divisions for adding -RS
Set dRS = SQLConnect.execute(StrSql)
If Not dRS.eof Then
%>
                    <select name="Divisions" style="background-color: #ffffff; border: 1px solid <%= thColor %>; font-family: Verdana; font-size: 12px;">
<%
  While Not dRS.eof
    'If the division is already listed as an AdditionalDivision
    'and is the Division being edited
    'do not show it in the add select box
    If (InStr(CurDiv,dRS("DivisionID")) > 0) Or (dRS("DivisionID") = CInt(Request.QueryString("id"))) Then
      'do nothing - do not show Division
    Else
%>
                      <option value="<%= dRS("DivisionID") %>"><%= dRS("DIV") %> - <%= dRS("DivisionDescription") %></option>
<%
    End If
    dRS.movenext
  Wend
%>
                    </select>
<%
End If
%>
                    &nbsp;<input type="submit" name="submit" value="Add" />
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





