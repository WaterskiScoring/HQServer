<!--#include virtual="/epl/functions.asp" -->

<% 

If not Session("aauth") then response.redirect "Login.asp"

Server.ScriptTimeout = 300

Function RemoveInvalidChars(strInput)
    dim workingstring
	On Error Resume Next
	For i = 1 to Len(strInput)
		If isNumeric(Mid(strInput, i, 1)) then
			workingstring = workingstring & Mid(strInput, i, 1)
		End If
		If (Mid(strInput, i, 1)) => "a" and (Mid(strInput, i, 1)) <=  "z" then
			workingstring = workingstring & Mid(strInput, i, 1)
		End If
		If (Mid(strInput, i, 1)) => "A" and (Mid(strInput, i, 1)) <=  "Z" then
			workingstring = workingstring & Mid(strInput, i, 1)
		End If
		If (Mid(strInput, i, 1)) = "@" Or (Mid(strInput, i, 1)) = "." Then
				workingstring = workingstring & Mid(strInput, i, 1)
		End If
	Next
	RemoveInvalidChars = workingstring
	
End Function

'	-----------------------------------------------------------------------
'	Start by sucking Membership Pricing Info from HQ Table into local Array

Dim MT, MemPrice(200), MemUpgrd(200)
FOR MT = 1 to 200: MemPrice(MT) = 0: MemUpgrd(MT) = 0: NEXT

Set SQLConnect = CreateObject("ADODB.Connection")
SQLConnect.Open Application("HQSQLConn")

strSql = "SELECT * FROM [Membership Types with pricing]" 
strSql = strSql & " WHERE EffectiveFrom <= CONVERT(DATETIME, '" & session("tournamentdate") & " 00:00:00', 102)"
strSql = strSql & " AND EffectiveTo >= CONVERT(DATETIME, '" & session("tournamentdate") & " 00:00:00', 102)"
Set HQRS = SQLConnect.Execute(strSql)
DO UNTIL HQRS.EOF
	MT = HQRS("Membership Type Code")
	MemPrice(MT) = HQRS("MemberShipTypeRates")
	MemUpgrd(MT) = HQRS("CostToUpgrade")
	HQRS.MoveNext
LOOP

HQRS.Close
Set HQRS = Nothing


%> 

<html>

<head>
<title>Create Registration Template v1.4</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" background = "/images/TopBackground.jpg" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" >

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#42639F">
      <p>&nbsp;</p>
      <p align="center"><font face="Verdana" size="6" color="#FFFFFF">
      	USA Water Ski Registration Templates</font></p>
      <p align="center"><font face="Verdana" size="4" color="#FFFFFF">
      	Registration Support for -- <%=session("TournamentName")%></font></p>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>  
  
<table border="0" cellspacing="0" cellpadding="0">  
  <tr> 
    <td width="185" valign="top" bgcolor="#42639F">
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
	<font face="Verdana" size="2" COLOR="#FFFFFF"><br>&nbsp;Currently Logged in as: </font><br>
	<font face="Verdana" size="2" COLOR="#FFFFFF">&nbsp;<%=Session("UserName")%>&nbsp;&nbsp;
		<%=session("TournamentDate")%></font><br>
	<br>
	<% Else %>
	<font face="Verdana" size="2" COLOR="#FFFFFF"><br>&nbsp;Not currently logged in.</font>
	<% End If %>
	
            <% If Session("aauth") then 
	
				Dim TopUser
				Set TopUser = Server.CreateObject("ADODB.RecordSet")
				TopUser.ActiveConnection = objConn
				TopUser.Open "SELECT * FROM Users999 where Name = '" & Session("UserName") & "'"
			%>
			<font face="Verdana" size="2"> 
         <br>&nbsp;<a href="logout.asp"><font face="arial" COLOR="#FFFFFF">Log Out</font></a>&nbsp;<br>
			</font>
            <% Else %>
			<br>
            <% End If %>
			<br>
	        &nbsp;<a href="http://www.usawaterski.org"><font face="arial" size="2" COLOR="#FFFFFF">USA Water Ski Home</font></a><br>&nbsp;<br>
			<br>
            <font face="Verdana" size="1">&nbsp;<font COLOR="#FFFFFF">Powered by</font> <a href="http://www.epolk.com"><font COLOR="#FFFFFF">ePolk.com</font></a><br></font>

  </td>

	<td>

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



Dim objFSO
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Dim path
path = Server.MapPath("Excel/")
'Randomize()
'Dim num

Dim DateRaw, DateFmt, I1, I2
DateRaw = Date(): I1 = instr(DateRaw,"/"): I2 = instr(I1+1,DateRaw,"/")
DateFmt = Mid(DateRaw,I2+1): ' Start with Year value
IF I1=2 THEN DateFmt = DateFmt + "-0" + Left(DateRaw,1): ELSE DateFmt = DateFmt + "-" + Left(DateRaw,2)
IF I2-I1=2 THEN DateFmt = DateFmt + "-0" + Mid(DateRaw,I1+1,1): ELSE DateFmt = DateFmt + "-" + Mid(DateRaw,I1+1,2)

'""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'"""""""""""""" With Scores and Ratings """""""""""""""""""""""
'""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""


objFSO.CopyFile path & "/Templates/template_blank_with_scores.xls", path & "/template_with_scores.xls" , True

'Now open a connection to the new XLS file
        Set objExcelConn = Server.CreateObject("ADODB.Connection")
        objExcelConn.Open "ExcelDSNwithScores"

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
		
        objExcelSingleFields.Source = "Select * from ActiveAsOfRange"
        objExcelSingleFields.Open
		objExcelSingleFields.Fields(0).Value = " AS OF " & DateFmt
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
		
        objExcelSingleFields.Source = "Select * from InActiveAsOfDate"
        objExcelSingleFields.Open
		objExcelSingleFields.Fields(0).Value = " AS OF " & DateFmt
		objExcelSingleFields.update
		objExcelSingleFields.close
		
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





''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Now build a temp table with members and their primary division (step 1)  '
''' Then add other divisions (like Mens Open, etc) (step 2) ''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'clear out the temp table of any entries for this session
objConn1.execute "Delete FROM [Temp Registration Template Export Table] where sessionid = " & (Session.SessionID)

'clear out the temp table of any old entries
objConn1.execute "Delete FROM [Temp Registration Template Export Table] WHERE (DateandTimeRecordAdded < CONVERT(DATETIME, '" & date() & " 00:00:00', 102))"


Set MemberstoExport = Server.CreateObject("ADODB.RecordSet")
MemberstoExport.ActiveConnection = objConn1
MemberstoExport.Open "SELECT * FROM [Export Members to Excel] Where " & Session("StateSQL") & " ;" 



Dim TempTable
Set TempTable = Server.CreateObject("ADODB.RecordSet")
TempTable.ActiveConnection = objConn1
TempTable.LockType = 3	'adLockOptimistic
TempTable.Open "[Temp Registration Template Export Table]" 

Do until MemberstoExport.EOF

	SkiAge = Session("TournamentYear") - DATEPART("yyyy", MemberstoExport("BirthDate")) - 1

	TempTable.addnew
	TempTable("sessionid") = (Session.SessionID)
	TempTable("newmemid") = MemberstoExport("newmemid")
	TempTable("lname") = MemberstoExport("lname")
	TempTable("fname") = MemberstoExport("fname")
	TempTable("Div") = CalculateDivision(SkiAge, MemberstoExport("Gender"))
	TempTable("SkiAge") = SkiAge
	TempTable("city") = MemberstoExport("city")
	TempTable("State") = MemberstoExport("State")
	TempTable("PrimaryRecord") = True
	
	if MemberstoExport("EffectiveTo") >= cdate(session("tournamentdate")) and MemberstoExport("CanSkiInTournaments") = True then
		TempTable("Active") = True		
	else
		TempTable("Active") = False
		if MemberstoExport("EffectiveTo") <= cdate(session("tournamentdate")) then
			TempTable("UpgradeDescription") = "Exp " & datepart("m",MemberstoExport("EffectiveTo")) & "/" & datepart("yyyy",MemberstoExport("EffectiveTo"))
		else
			TempTable("UpgradeDescription") = "Needs Upgrd" 
			TempTable("CosttoUpgrade") = MemberstoExport("CosttoUpgrade")
		end if
	end if
	TempTable.Update	
	
	MemberstoExport.MoveNext
Loop



'''''''''''''  Step 2  '''''''''''''''''''
'Now add to the temp table everyone that has scores under the extra divisions
'Extra Divisions with Scores to Add to Registration Template Export Grouped
MemberstoExport.Close 
MemberstoExport.Open "SELECT * FROM [Extra Divisions with Scores to Add to Registration Template Export Grouped] Where SessionID = " & Session.SessionID & " ;" 


'Used for debugging
Dim ExtraCounter
ExtraCounter = 0

Do until MemberstoExport.EOF
	ExtraCounter = ExtraCounter + 1
	SkiAge = Session("TournamentYear") - DATEPART("yyyy", MemberstoExport("BirthDate")) - 1

	TempTable.addnew
	TempTable("sessionid") = (Session.SessionID)
	TempTable("newmemid") = MemberstoExport("PersonIDwithCheckDigit")
	TempTable("lname") = MemberstoExport("lname")
	TempTable("fname") = MemberstoExport("fname")
	TempTable("Div") = MemberstoExport("Div")
	TempTable("SkiAge") = SkiAge
	TempTable("city") = MemberstoExport("city")
	TempTable("State") = MemberstoExport("State")
	TempTable("PrimaryRecord") = False
	TempTable.Update	
	
	MemberstoExport.MoveNext
Loop

'Now add the 4 parts to the consolidated officals field
objConn1.execute "UPDATE [Officials Abbreviated Ratings Update View] SET Officials_Driver = LevelAbbreviationforTemplate WHERE     (RatingType_ID = 3)"
objConn1.execute "UPDATE [Officials Abbreviated Ratings Update View] SET Officials_Judge = LevelAbbreviationforTemplate WHERE     (RatingType_ID = 1)"
objConn1.execute "UPDATE [Officials Abbreviated Ratings Update View] SET Officials_Scorer = LevelAbbreviationforTemplate WHERE     (RatingType_ID = 2)"
objConn1.execute "UPDATE [Officials Abbreviated Ratings Update View] SET Officials_Safety = LevelAbbreviationforTemplate WHERE     (RatingType_ID = 9)"
objConn1.execute "UPDATE  [Officials Abbreviated Ratings Update View] SET OfficlasRatingsConsolidated = Officials_Driver + Officials_Judge + Officials_Scorer + Officials_Safety"

'This query used the temp table just built

objRS.Open "SELECT * FROM [Export Members to Excel New 2] Where (" & Session("StateSQL") & ") AND SessionID = " & (Session.SessionID) & ";" 

Dim Counter1
Dim Counter2

Do until objRS.EOF

	SkiAge = Session("TournamentYear") - DATEPART("yyyy", objRS("BirthDate")) - 1

	if objRS("EffectiveTo") >= cdate(session("tournamentdate")) and objRS("CanSkiInTournaments") = True then
		Counter1 = Counter1 + 1
		objExcelRS.addnew
		objExcelRS.Fields(0).Value = objRS("newmemid")
		objExcelRS.Fields(1).Value = objRS("lname")
		objExcelRS.Fields(2).Value = objRS("fname")
		'objExcelRS.Fields(4).Value = CalculateDivision(SkiAge, objRS("Gender"))
		objExcelRS.Fields(4).Value = objRS("Div")
		'MOK - 4-28-2004
		objExcelRS.Fields(5).Value = SkiAge
		objExcelRS.Fields(6).Value = objRS("city")
		objExcelRS.Fields(7).Value = objRS("State")
		
		'added 4-11-2007 MOK
		objExcelRS.Fields(11).Value = objRS("OfficlasRatingsConsolidated")
		objExcelRS.Fields(12).Value = MyFormatNumber(objRS("SlalomScore"),6)
		objExcelRS.Fields(13).Value = MyFormatNumber(objRS("TrickScore"),6)
		objExcelRS.Fields(14).Value = MyFormatNumber(objRS("JumpScore"),6)
		objExcelRS.Fields(15).Value = objRS("SlalomRating")
		objExcelRS.Fields(16).Value = objRS("TrickRating")
		objExcelRS.Fields(17).Value = objRS("JumpRating")
		
	    objExcelRS.Fields(21).Value = "Yes"
		objExcelRS.Update
	else
		Counter2 = Counter2 + 1
		objExcelInActive.addnew
		objExcelInActive.Fields(0).Value = objRS("newmemid")
		objExcelInActive.Fields(1).Value = objRS("lname")
		objExcelInActive.Fields(2).Value = objRS("fname")
		'objExcelInActive.Fields(4).Value = CalculateDivision(SkiAge, objRS("Gender"))
		objExcelInActive.Fields(4).Value = objRS("Div")
		'MOK - 4-28-2004
		objExcelInActive.Fields(5).Value = SkiAge
		'objExcelInActive.Fields(5).Value = objRS("SkiAge")
		objExcelInActive.Fields(6).Value = objRS("city")
		objExcelInActive.Fields(7).Value = objRS("State")
		
		'added 4-11-2007 MOK
		objExcelInActive.Fields(11).Value = objRS("OfficlasRatingsConsolidated")
		objExcelInActive.Fields(12).Value = MyFormatNumber(objRS("SlalomScore"),6)
		objExcelInActive.Fields(13).Value = MyFormatNumber(objRS("TrickScore"),6)
		objExcelInActive.Fields(14).Value = MyFormatNumber(objRS("JumpScore"),6)
		objExcelInActive.Fields(15).Value = objRS("SlalomRating")
		objExcelInActive.Fields(16).Value = objRS("TrickRating")
		objExcelInActive.Fields(17).Value = objRS("JumpRating")

		objExcelInActive.Fields(21).Value = "    No"

		' Figure applicable Renewal / Upgrade Amount based on MemType & Status

		MT = objRS("MembershipTypeCode")
		IF MT < 1 OR MT > 200 THEN MT = 1

		IF objRS("EffectiveTo") < cdate(session("tournamentdate")) THEN 
			IF objRS("CanSkiInTournaments") = False THEN
				objExcelInActive.Fields(22).Value = "Nds Rnw/Upg" 
				objExcelInActive.Fields(23).Value = FormatNumber(MemPrice(MT)+MemUpgrd(MT),2)
			ELSE
				objExcelInActive.Fields(22).Value = "Needs Renew" 
				objExcelInActive.Fields(23).Value = FormatNumber(MemPrice(MT),2)
			END IF
		ELSE 
			objExcelInActive.Fields(22).Value = "Needs Upgrd" 
			objExcelInActive.Fields(23).Value = FormatNumber(MemUpgrd(MT),2)
		END IF
		
		objExcelInActive.Update

	end if
	
	objRS.MoveNext
Loop


'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

'clear out the temp table of any entries for this session
objConn1.execute "Delete FROM [Temp Registration Template Export Table] where sessionid = " & (Session.SessionID)

'clear out the temp table of any old entries
objConn1.execute "Delete FROM [Temp Registration Template Export Table] WHERE (DateandTimeRecordAdded < CONVERT(DATETIME, '" & date() & " 00:00:00', 102))"


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
Dim filenamewithscores
'"06M123-Entries-SSSSSS-YYYYMMDD", 
filenamewithscores = "Entries-" & Session("StateList") & "-" & DateFmt

'Add the Tournament Name to the start of the file name
'session("TournamentName")
if len(session("TournamentName")) > 0 then
	'filename = "TournamentRegistrationFile-" & session("UserName") & ".xls"
	filenamewithscores = session("TournamentName") & "-" & filenamewithscores
end if

'5-18-2006 Remove any strange characters from the Tournamentname
filenamewithscores = RemoveInvalidChars(filenamewithscores)

'Append the username
if len(session("UserName")) > 0 then
	'filename = "TournamentRegistrationFile-" & session("UserName") & ".xls"
	filenamewithscores = filenamewithscores & "-" & session("UserName") & ".xls"
else
	'filename = "TournamentRegistrationFile.xls"
	filenamewithscores = filenamewithscores & ".xls"
end if

'objFSO.CopyFile path & "/template.xls", path & "/" & filename , True
objFSO.CopyFile path & "/template_with_scores.xls", path & "/" & filenamewithscores , True

'Clean up old files
Set f = objFSO.GetFolder("d:\webs\usawaterski.org\admin\excel\")  
Set fc = f.Files 
Response.Write "<br>"
For Each f1 in fc
	'Response.Write f1.name 
	Set myfile = objFSO.GetFile("d:\webs\usawaterski.org\admin\excel\" & f1.name)
	'Response.Write  "Date:"  & myfile.DateCreated 
	'Response.Write  "Age:"  & datediff("d",myfile.DateCreated,date()) & "<br>"
	if datediff("d",myfile.DateCreated,date()) > 2 and left(myfile.name,8) <> "Template" then
		myfile.delete
	end if
	
Next  

Set f = nothing
Set fc = nothing

Set objFSO = Nothing

'Clean up old records in temp table


%>

  <table>
      <tr> 
         <td width="14">&nbsp;</td>
         <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Here is your Registration 
         Template. </font></td>
      </tr>

      <tr> 
         <td>&nbsp;</td>
      </tr>

      <tr> 
         <td>&nbsp;</td>
         <td><a href="excel/<% response.write filenamewithscores %>"><font face="Arial" size="2">RIGHT 
         Click Here</font></a>&nbsp; <font size="2" face="Verdana, Arial, Helvetica, sans-serif">to 
         download your Registration Template, then select the "Save As" 
         option from that menu, and then choose a suitable location to 
         store the download in your PC. </font></td>
      </tr>
   
      <tr> 
         <td>&nbsp;</td>
      </tr>

      <tr> 
         <td>&nbsp;</td>
         <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
         After your Registration Template download has completed, then open the 
         Excel file from that location on your PC.&nbsp; It will open automatically 
         to an Instructions Tab.&nbsp; Please review that updated Instructions section 
         for the latest information on contents and usage. </font></td>
      </tr>

      <tr> 
         <td>&nbsp;</td>
      </tr>

      <tr> 
         <td>&nbsp;</td>
         <td><font color="#FF0000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>
         !! New Content for 2008 !!</strong>&nbsp;
         </font><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
         Rankings List levels in Slalom and Tricks and Jumping and Overall
         now occupy the columns which used to contain Rating codes.  Also,
         An Officials column precedes these, which shows each member's highest 
         rating as an official in Driving, Judging, Scoring and Safety.  
        </font></td>
      </tr>


      <tr> 
         <td>&nbsp;</td>
      </tr>

      <tr> 
         <td>&nbsp;</td>
         <td><font color="#FF0000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>
         !! New Function now Available in 2008 !!</strong>&nbsp;
         </font><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
         After you've downloaded this Registration Template, you can later 
         fold in additional selected members, one-by-one, using the lookup 
         feature noted on the earlier screen.&nbsp; With that feature, you
         can then just copy and paste the information for those additional 
         participants into your template using Excel.&nbsp; Detailed 
         instructions will appear on the lookup results window, when you 
         get to that point.
         </font></td>
      </tr>

      <tr> 
         <td>&nbsp;</td>
      </tr>
 	</table>

	<TABLE ALIGN="CENTER" WIDTH=70%>
		
		<TR><TD>&nbsp;</TD></TR>

		<TR>

	    <TD width=30% align=center>
		<form action="LookupMembers.asp?FormStatus=newsearch" method="post">
		<input type="submit" style="width:9em" value="Lookup Members"></form>
    	</TD>

	    <td width=30% align=center>     				
		<form action="CreateRegTemplateStep1.asp" method="post">
    <input type="submit" style="width:9em" value="Quit"></form>
 	    </td>
  	    
 	  </TR>

 	</TABLE>

  	  </td>
	  </tr>
</table>
</body>
</html>






