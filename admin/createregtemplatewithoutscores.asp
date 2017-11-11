<!--#include virtual="/epl/functions.asp" -->
<!--#include virtual="/admin/includes/security.asp" -->

<html>

<head>
<title>Create Registration Template v1.3</title>

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
<%  	

Function MyFormatNumber(InNumber, NumberofPositions)	
	if len(InNumber) = 0 then
		MyFormatNumber = ""
	elseif isnull(InNumber) then
		MyFormatNumber = ""	
	elseif not isnumeric(InNumber) then
		MyFormatNumber = ""	
	else
		WorkingString = formatnumber(InNumber,6)
		'now remove any commas
		WorkingString2 = ""
		for x = 1 to len(WorkingString) 
			if mid(WorkingString, x ,1) = "," then
				'skip
			else
				WorkingString2 = WorkingString2 & mid(WorkingString,x,1)
			end if
		next
		MyFormatNumber = left(WorkingString2,6)
	end if
end Function


Dim objConn
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open Application("WaterSkiConn")
%>
	<% If Session("aauth") then %>
	<font face="Verdana" size="2" COLOR="#FFFFFF">Currently Logged in as: </font><br>
	<font face="Verdana" size="2" COLOR="#FFFFFF"><%= Session("UserName") %></font><br>
	<br>
	<% Else %>
	<font face="Verdana" size="2" COLOR="#FFFFFF">Not currently logged in.</font>
	<% End If %>
	
            <% If Session("aauth") then 
	
				Dim TopUser
				Set TopUser = Server.CreateObject("ADODB.RecordSet")
				TopUser.ActiveConnection = objConn
				TopUser.Open "SELECT * FROM Users999 where Name = '" & Session("UserName") & "'"
			%>
			<font face="Verdana" size="2"> 

			
            <a href="/admin/logout.asp"><font face="arial" COLOR="#FFFFFF"><br>Log Out</font></a>&nbsp;<br>
			</font>
            <% Else %>
			<br>
            <% End If %>
			<br>
			<br>
            <font face="Verdana" size="1">&nbsp;<font COLOR="#FFFFFF">Powered by</font> <a href="http://www.epolk.com"><font COLOR="#FFFFFF">ePolk.com</font></a><br></font>

  </td>
    <td valign="top" >
	
 <p>
        <%
Function CalculateDivision(SkiAge, Gender)
Dim AgeDivision
if len(SkiAge) = 0 then
	AgeDivision = "-"
elseif SkiAge >= 0 AND SkiAge < 10 THEN '1' 
	AgeDivision = "1"
elseif  SkiAge >= 10 AND SkiAge < 14 THEN '2' 
	AgeDivision = "2"
elseif  SkiAge >= 14 AND SkiAge < 18 THEN '3' 
	AgeDivision = "3"
elseif  SkiAge >= 18 AND SkiAge < 25 THEN '1' 
	AgeDivision = "1"
elseif  SkiAge >= 25 AND SkiAge < 35 THEN '2' 
	AgeDivision = "2"
elseif  SkiAge >= 35 AND SkiAge < 45 THEN '3' 
	AgeDivision = "3"
elseif  SkiAge >= 45 AND SkiAge < 53 THEN '4' 
	AgeDivision = "4"
elseif  SkiAge >= 53 AND SkiAge < 60 THEN '5' 
	AgeDivision = "5"
elseif  SkiAge >= 60 AND SkiAge < 65 THEN '6' 
	AgeDivision = "6"
elseif  SkiAge >= 65 AND SkiAge < 70 THEN '7' 
	AgeDivision = "7"
elseif  SkiAge >= 70 AND SkiAge < 75 THEN '8' 
	AgeDivision = "8"
elseif  SkiAge >= 75 AND SkiAge < 80 THEN '9' 
	AgeDivision = "9"
elseif  SkiAge >= 80 AND SkiAge < 85 THEN 'A' 
	AgeDivision = "A"
elseif  SkiAge >= 85 THEN 'B' 
	AgeDivision = "B"
else
	AgeDivision = "-"
end if
					  
if Gender = "M" AND SkiAge < 18 THEN 'B' 
	SkiGender = "B"
elseif Gender = "M" AND SkiAge >= 18 THEN 'M' 
	SkiGender = "M"
elseif Gender = "F" AND SkiAge < 18 THEN 'G' 
	SkiGender = "G"
elseif Gender = "F" AND SkiAge >= 18 THEN 'W' 
	SkiGender = "W"
else 
	SkiGender = "-"
end if					  

CalculateDivision = SkiGender & AgeDivision
				  
End Function

Dim objConn1
Dim objRS
Set objConn1 = Server.CreateObject("ADODB.Connection")
objConn1.Open Application("WaterSkiConn")
Set objRS = Server.CreateObject("ADODB.RecordSet")
objRS.ActiveConnection = objConn1

objRS.Open "SELECT * FROM [Export Members to Excel] Where " & Session("StateSQL") & " ;" 

Dim objFSO
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Dim path
path = Server.MapPath("/admin/Excel/")
'Randomize()
'Dim num

objFSO.CopyFile path & "/Templates/template_blank.xls", path & "/template.xls" , True

'Now open a connection to the new XLS file
        Set objExcelConn = Server.CreateObject("ADODB.Connection")
        objExcelConn.Open "ExcelDSN"

        Set objExcelSingleFields = Server.CreateObject("ADODB.Recordset")
        objExcelSingleFields.ActiveConnection = objExcelConn 
        objExcelSingleFields.CursorType = 3                    'Static cursor.
        objExcelSingleFields.LockType = 2                      'Pessimistic Lock.
        objExcelSingleFields.Source = "Select * from ActiveTournamentName"
        objExcelSingleFields.Open
		objExcelSingleFields.Fields(0).Value = session("TournamentName")
		objExcelSingleFields.update
		objExcelSingleFields.close
		
        objExcelSingleFields.Source = "Select * from ActiveTournamentID"
        objExcelSingleFields.Open
		objExcelSingleFields.Fields(0).Value = Session("UserName")	'this is the same as the tournament ID
		objExcelSingleFields.update
		objExcelSingleFields.close
		
        objExcelSingleFields.Source = "Select * from InActiveTournamentName"
        objExcelSingleFields.Open
		objExcelSingleFields.Fields(0).Value = session("TournamentName")
		objExcelSingleFields.update
		objExcelSingleFields.close
		
        objExcelSingleFields.Source = "Select * from InActiveTournamentID"
        objExcelSingleFields.Open
		objExcelSingleFields.Fields(0).Value = Session("UserName")
		objExcelSingleFields.update
		objExcelSingleFields.close
		

			'session("FromUSAWS") = objRS("FromUSAWS")
			'session("TournamentDate") = objRS("TournamentDate")
			'session("TournamentName") = objRS("TournamentName")

        Set objExcelRS = Server.CreateObject("ADODB.Recordset")
        objExcelRS.ActiveConnection = objExcelConn 
        objExcelRS.CursorType = 3                    'Static cursor.
        objExcelRS.LockType = 2                      'Pessimistic Lock.
        objExcelRS.Source = "Select * from ActiveRange"
        objExcelRS.Open

        Set objExcelInActive = Server.CreateObject("ADODB.Recordset")
        objExcelInActive.ActiveConnection = objExcelConn 
        objExcelInActive.CursorType = 3                    'Static cursor.
        objExcelInActive.LockType = 2                      'Pessimistic Lock.
        objExcelInActive.Source = "Select * from InActiveRange"
        objExcelInActive.Open




Do until objRS.EOF

	SkiAge = Session("TournamentYear") - DATEPART("yyyy", objRS("BirthDate")) - 1
	if objRS("EffectiveTo") >= cdate(session("tournamentdate")) and objRS("CanSkiInTournaments") = True then
		objExcelRS.addnew
		objExcelRS.Fields(0).Value = objRS("newmemid")
		objExcelRS.Fields(1).Value = objRS("lname")
		objExcelRS.Fields(2).Value = objRS("fname")
		'objExcelRS.Fields(4).Value = objRS("SkiDivision")
		objExcelRS.Fields(4).Value = CalculateDivision(SkiAge, objRS("Gender"))
		'MOK - 4-28-2004
		objExcelRS.Fields(5).Value = SkiAge
		'objExcelRS.Fields(5).Value = objRS("SkiAge")
		objExcelRS.Fields(6).Value = objRS("city")
		objExcelRS.Fields(7).Value = objRS("State")
	    objExcelRS.Fields(14).Value = "Yes"
		objExcelRS.Update
	else
		objExcelInActive.addnew
		objExcelInActive.Fields(0).Value = objRS("newmemid")
		objExcelInActive.Fields(1).Value = objRS("lname")
		objExcelInActive.Fields(2).Value = objRS("fname")
		objExcelInActive.Fields(4).Value = CalculateDivision(SkiAge, objRS("Gender"))
		'MOK - 4-28-2004
		objExcelInActive.Fields(5).Value = SkiAge
		'objExcelInActive.Fields(5).Value = objRS("SkiAge")
		objExcelInActive.Fields(6).Value = objRS("city")
		objExcelInActive.Fields(7).Value = objRS("State")
		if objRS("EffectiveTo") <= cdate(session("tournamentdate")) then
			objExcelInActive.Fields(14).Value = "    No"
			'objExcelInActive.Fields(15).Value = "Membership Expired " & objRS("EffectiveTo")
			objExcelInActive.Fields(15).Value = "Exp " & datepart("m",objRS("EffectiveTo")) & "/" & datepart("yyyy",objRS("EffectiveTo"))
		else
			objExcelInActive.Fields(14).Value = "    No"
			'objExcelInActive.Fields(15).Value = "Invalid Mem. Type " 
			objExcelInActive.Fields(15).Value = "Needs Upgrd" 
			objExcelInActive.Fields(16).Value = objRS("CosttoUpgrade")
		end if
		objExcelInActive.Update
	end if
	
	
	

	objRS.MoveNext
Loop

'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

objExcelRS.close
set objExcelRS = nothing
objExcelInActive.close
set objExcelInActive = nothing
objExcelConn.close
set objExcelConn = nothing
'
objRS.Close
Set objRS = Nothing
objConn1.Close
Set objConn1 = Nothing

'Now copy the file from Template to a file with the tournamentid
Dim filename

'Add the Tournament Name to the start of the file name
'session("TournamentName")
if len(session("TournamentName")) > 0 then
	filename = "TournamentRegistrationFile-" & session("UserName") & ".xls"
end if

'Append the username
if len(session("UserName")) > 0 then
	filename = "TournamentRegistrationFile-" & session("UserName") & ".xls"
else
	filename = "TournamentRegistrationFile.xls"

end if

objFSO.CopyFile path & "/template.xls", path & "/" & filename , True

'Clean up old files
Set f = objFSO.GetFolder("d:\webs\usawaterski.org\admin\excel\")  
Set fc = f.Files 
Response.Write "<br>"
For Each f1 in fc
	'Response.Write f1.name 
	Set myfile = objFSO.GetFile("d:\webs\usawaterski.org\admin\excel\" & f1.name)
	'Response.Write  "Date:"  & myfile.DateCreated 
	'Response.Write  "Age:"  & datediff("d",myfile.DateCreated,date()) & "<br>"
	if datediff("d",myfile.DateCreated,date()) > 2 and left(myfile.name,10) = "Tournament" then
		myfile.delete
	end if
	
Next  

Set f = nothing
Set fc = nothing

Set objFSO = Nothing


%>
      </p>
      <p><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Here is your Registration 
        Template</font>  </p>
      <p><a href="/admin/excel/<% response.write filename %>"><font face="Arial" size="2">RIGHT 
        Click Here</font></a> <font size="2" face="Verdana, Arial, Helvetica, sans-serif">to 
        download your Registration Template, then select the "Save As" option 
        from that menu, and choose a suitable location to store the download in 
        your PC. </font></p>
      <p>&nbsp;</p>
      <p>&nbsp;</p>

	
	
	</td>
  </tr>
</table>
</body>
</html>






