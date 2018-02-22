<!--#include virtual="/admin/includes/security.asp" -->
<html>
<head>
	<title>Admin Index</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" background = "/templates/images/TopBackground.jpg" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" >
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
   <td bgcolor="#42639F">
<!--#include virtual="/admin/includes/menu.asp" -->
  </td>
    <td valign="top" >

 <%


Dim objConn1
Dim objRS
Set objConn1 = Server.CreateObject("ADODB.Connection")
objConn1.Open Application("WaterSkiConn")
Set objRS = Server.CreateObject("ADODB.RecordSet")
objRS.ActiveConnection = objConn1

'objRS.Open "SELECT * FROM Members ORDER BY PersonID"
SQLString = "SELECT Members.PersonID, Members.PersonIDwithCheckDigit, Members.FirstName, Members.MiddleName, Members.LastName, Members.NameSuffix, "
SQLString = SQLString & " Members.SSN, Members.CompanyName, Members.BirthDate, Members.Sex, Members.DivisionCode1, Members.DivisionCode2, Members.State, "
SQLString = SQLString & " Members.MembershipTypeCode, Members.EffectiveFrom, Members.EffectiveTo, MembershipTypes.TypeCode, MembershipTypes.Description "
SQLString = SQLString & " FROM Members INNER JOIN "
SQLString = SQLString & " MembershipTypes ON Members.MembershipTypeCode = MembershipTypes.MemberShipTypeID"
objRS.Open SQLString


Dim objFSO
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Dim path
path = Server.MapPath("/admin/tmp/")
Randomize()
Dim num
Dim filename
num = Rnd()
num = split(Rnd(), ".")
filename = "\mem" & num(1) & ".csv"
path = path & filename
'Response.Write path

Dim objFile
Set objFile = objFSO.CreateTextFile(path)
'Write column headers
For each field in objRS.Fields
	objFile.Write field.Name & ","
Next
objFile.WriteLine

'Now cycle through the fields
Do until objRS.EOF
	For each field in objRS.Fields
		objFile.Write """" & field.value & """" & ","
	Next
	objFile.WriteLine
	objRS.MoveNext
Loop
objFile.Close
Set objFile = Nothing
Set objFSO = Nothing


objRS.Close
Set objRS = Nothing
objConn1.Close
Set objConn1 = Nothing

%>
      <p><a href="/admin/tmp<%= filename %>"><font face="Arial" size="2">Click
        here to download your file.</font></a> </p>
      <p><font face="Arial, Helvetica, sans-serif" size="-2">(to download, right
        click on the link and select save target as.)</font> </p>



	</td>
  </tr>
</table>
</body>
</html>






