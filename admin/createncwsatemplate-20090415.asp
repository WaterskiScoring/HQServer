<!--#include virtual="/epl/functions.asp" -->

<% 

If not Session("aauth") then response.redirect "Login.asp"

Server.ScriptTimeout = 300

' The following lines of HTML display the "opening please wait" banner.

%>
    
<html><head><title>USA Water Ski NCWSA Registration Template</title>
    <SCRIPT LANGUAGE="JavaScript">
    // First we detect the browser type
    if(document.getElementById) { // IE 5 and up, NS 6 and up
    	var upLevel = true;
    	}
    else if(document.layers) { // Netscape 4
    	var ns4 = true;
    	}
    else if(document.all) { // IE 4
    	var ie4 = true;
    	}
    
    function showObject(obj) {
    if (ns4) {
    	obj.visibility = "show";
    	}
    else if (ie4 || upLevel) {
    	obj.style.visibility = "visible";
    	}
    }
    
    function hideObject(obj) {
    if (ns4) {
    	obj.visibility = "hide";
    	}
    if (ie4 || upLevel) {
    	obj.style.visibility = "hidden";
    	}
    }
    
    </SCRIPT>
    </head>
    <body>
    <DIV ID="splashScreen" STYLE="position:absolute;z-index:5;top:30%;left:35%;">
    <TABLE BGCOLOR="#000000" BORDER=1 BORDERCOLOR="#000000"	CELLPADDING=0 CELLSPACING=0 HEIGHT=150 WIDTH=300>
    <TR>
    <TD WIDTH="100%" HEIGHT="100%" BGCOLOR="#CCCCCC" ALIGN="CENTER" VALIGN="MIDDLE">
    <BR>
    <FONT FACE="Helvetica,Verdana,Arial" SIZE=2 COLOR="#000066">
    <B>Preparing your Registration Template.<br><br>
    This may take a minute or so ...<br><br><br>  
    </B></FONT>
    <IMG SRC="includes/wait.gif" BORDER=1 WIDTH=150 HEIGHT=15><BR><BR>
    </TD>
    </TR>
    </TABLE>
    </DIV>
    
<%

' Once the above "please wait" banner is written to HTML, we flush the response
' buffer to make the page appear to the users browser.  That sits on their display
' while the rest of the template preparation script processing takes place.
    
response.flush


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
'	-----------------------------------------------------------------------

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


Dim objConn
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open Application("WaterSkiConn")
    

Dim objFSO
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Dim path
path = Server.MapPath("Excel/")
'Randomize()
'Dim num

Dim DateRaw, DateFmt, I1, I2, RowNo
DateRaw = Date(): I1 = instr(DateRaw,"/"): I2 = instr(I1+1,DateRaw,"/")
DateFmt = Mid(DateRaw,I2+1): ' Start with Year value
IF I1=2 THEN DateFmt = DateFmt + "-0" + Left(DateRaw,1): ELSE DateFmt = DateFmt + "-" + Left(DateRaw,2)
IF I2-I1=2 THEN DateFmt = DateFmt + "-0" + Mid(DateRaw,I1+1,1): ELSE DateFmt = DateFmt + "-" + Mid(DateRaw,I1+1,2)

Dim objRS
Set objRS = Server.CreateObject("ADODB.RecordSet")
objRS.ActiveConnection = objConn

Dim objTL
Set objTL = Server.CreateObject("ADODB.RecordSet")
objTL.ActiveConnection = objConn



'""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'"""""""""""""" With Scores and Ratings """""""""""""""""""""""
'""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""


objFSO.CopyFile path & "/Templates/NCWSATemplateBlank.xls", path & "/template_with_scores.xls" , True

'Now open a connection to the new XLS file

Set objExcelConn = Server.CreateObject("ADODB.Connection")
objExcelConn.Open "ExcelDSNwithScores"

Set objExcelSingleFields = Server.CreateObject("ADODB.Recordset")
objExcelSingleFields.ActiveConnection = objExcelConn 
objExcelSingleFields.CursorType = 3                    'Static cursor.
objExcelSingleFields.LockType = 2                      'Pessimistic Lock.

objExcelSingleFields.Source = "Select * from AMenTourName"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = session("TournamentName")
objExcelSingleFields.update
objExcelSingleFields.close
		
objExcelSingleFields.Source = "Select * from AMenTourID"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = Session("TournamentID")	'this is the same as the tournament ID
objExcelSingleFields.update
objExcelSingleFields.close
		
objExcelSingleFields.Source = "Select * from AWomenTourName"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = session("TournamentName")
objExcelSingleFields.update
objExcelSingleFields.close
		
objExcelSingleFields.Source = "Select * from AWomenTourID"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = Session("TournamentID")	'this is the same as the tournament ID
objExcelSingleFields.update
objExcelSingleFields.close
		
objExcelSingleFields.Source = "Select * from UpgRnewTourName"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = session("TournamentName")
objExcelSingleFields.update
objExcelSingleFields.close
		
objExcelSingleFields.Source = "Select * from UpgRnewTourID"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = Session("TournamentID")
objExcelSingleFields.update
objExcelSingleFields.close
		
Set objExcelAMen = Server.CreateObject("ADODB.Recordset")
objExcelAMen.ActiveConnection = objExcelConn 
objExcelAMen.CursorType = 3                    'Static cursor.
objExcelAMen.LockType = 2                      'Pessimistic Lock.
objExcelAMen.Source = "Select * from AMenRange"
objExcelAMen.Open

Set objExcelAWomen = Server.CreateObject("ADODB.Recordset")
objExcelAWomen.ActiveConnection = objExcelConn 
objExcelAWomen.CursorType = 3                    'Static cursor.
objExcelAWomen.LockType = 2                      'Pessimistic Lock.
objExcelAWomen.Source = "Select * from AWomenRange"
objExcelAWomen.Open

Set objExcelUpgRnew = Server.CreateObject("ADODB.Recordset")
objExcelUpgRnew.ActiveConnection = objExcelConn 
objExcelUpgRnew.CursorType = 3                    'Static cursor.
objExcelUpgRnew.LockType = 2                      'Pessimistic Lock.
objExcelUpgRnew.Source = "Select * from UpgRnewRange"
objExcelUpgRnew.Open


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Next we open an extract of Team ID's and Names from the TeamList table.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim sSQL

sSQL = "Select TeamID, Max(TeamName) as TeamName from Cobra00025.USAWSRank.TeamsList"
sSQL = sSQL & " Where SptsGrpID = 'NCW' Group By TeamID Order by TeamID"

objTL.Open sSQL

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Now build a Query to Extract the Desired Members, joining in data 
''' pulled from the Rankings and Officials and Membership Type tables.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sSQL = "Select Substring(MX.MemberID,1,3) + '-' + Substring(MX.MemberID,4,2) + '-' +" 
sSQL = sSQL & " Substring(MX.MemberID,6,4) as MemID, MX.LastName, MX.FirstName,"
sSQL = sSQL & " Case when MX.Sex = 'F' Then 'CW' else 'CM' end as Div,"
sSQL = sSQL & " MX.Team, MX.Sorter, MX.Age, MX.City, MX.State,"

sSQL = sSQL & " Case when OD.PersonID is Null then '-' else Right(OD.RtgLvl,1) end +"
sSQL = sSQL & " Case when OJ.PersonID is Null then '-' else Right(OJ.RtgLvl,1) end +"
sSQL = sSQL & " Case when OC.PersonID is Null then '-' else Right(OC.RtgLvl,1) end +"
sSQL = sSQL & " Case when OS.PersonID is Null then '-' else Right(OS.RtgLvl,1) end as OffRat,"

sSQL = sSQL & " MX.EffTo, MX.Memtype, MX.MemCode, MX.CanSki, MX.SptsDiv, MX.Waiver"
		
sSQL = sSQL & " From (Select MT.PersonIDWithCheckDigit as MemberID, MT.PersonID,"
sSQL = sSQL & " Left(MT.LastName,12) as LastName, Left(MT.FirstName,10) as FirstName,"
sSQL = sSQL & " Case when MT.EffectiveTo < '" & session("TournamentDate") & "' OR"
sSQL = sSQL & " Typ.CanSkiInTournaments = 0 OR MT.WaiverStatusID = 0 then '???'"
sSQL = sSQL & " when RD.Team > ' ' then RD.Team else 'zzz' end as Sorter, RD.Team as Team,"
sSQL = sSQL & " (" & Session("TournamentYear") & "-Year(MT.BirthDate)-1) as Age,"
sSQL = sSQL & " MT.DivisionCode1 + '/' + MT.DivisionCode2 as SptsDiv,"
sSQL = sSQL & " Upper(Left(MT.Sex,1)) as Sex, MT.WaiverStatusID as Waiver,"
sSQL = sSQL & " Left(MT.City,12) as City, Left(MT.State,2) as State,"
sSQL = sSQL & " MT.EffectiveTo as EffTo, MT.MembershipTypeCode as MemType,"
sSQL = sSQL & " Typ.TypeCode as MemCode, Typ.CanSkiInTournaments as CanSki"
sSQL = sSQL & " from USAWaterski.dbo.Members as MT Inner Join"
sSQL = sSQL & " USAWaterski.dbo.MembershipTypes as Typ"
sSQL = sSQL & " ON MT.MembershipTypeCode = Typ.MemberShipTypeID"
sSQL = sSQL & " Left Join (Select MemberID, Max(Team) as Team"
sSQL = sSQL & " from Cobra00025.USAWSRank.Scores"

' **** The following restricts to 2008 or later ski years, for now
sSQL = sSQL & " where Left(Div,1) = 'C' and left(TourID,2) >= '08'"

sSQL = sSQL & " group by MemberID) as RD on RD.MemberID = MT.PersonIDWithCheckDigit"
sSQL = sSQL & " Where Typ.ExporttoTouramentRegistrationTemplate = 1"
sSQL = sSQL & " AND MT.Deceased = 0 AND ( (" & Session("TournamentYear")
sSQL = sSQL & " - Year(MT.BirthDate) - 1) between 16 and 29 OR"
sSQL = sSQL & " MT.DivisionCode1 = 'NCW' OR MT.DivisionCode2 = 'NCW' OR"

' sSQL = sSQL & Session("StateSQL") & " OR"

sSQL = sSQL & " PersonIDWithCheckDigit IN (Select Distinct"
sSQL = sSQL & " MemberID from Cobra00025.USAWSRank.Rankings" 
sSQL = sSQL & " Where left(Div,1) = 'C') ) ) as MX"

sSQL = sSQL & " Left Join (Select OT.PersonID,"
sSQL = sSQL & " Max(convert(char(1),LV.LevelOrderforTemplate)"
sSQL = sSQL & " + LV.LevelAbbreviationforTemplate) AS RtgLvl"
sSQL = sSQL & " FROM USAWaterski.dbo.Officials OT INNER JOIN"
sSQL = sSQL & " USAWaterski.dbo.Level LV ON OT.Level_ID = LV.Level_ID"
sSQL = sSQL & " WHERE OT.DivisionCode in ('AWS','USA')"
sSQL = sSQL & " AND LV.LevelOrderforTemplate IS NOT NULL"
sSQL = sSQL & " AND OT.RatingType_ID = 3 GROUP BY OT.PersonID) as OD"
sSQL = sSQL & " on OD.PersonID = MX.PersonID"

sSQL = sSQL & " Left Join (Select OT.PersonID,"
sSQL = sSQL & " Max(convert(char(1),LV.LevelOrderforTemplate)"
sSQL = sSQL & " + LV.LevelAbbreviationforTemplate) AS RtgLvl"
sSQL = sSQL & " FROM USAWaterski.dbo.Officials OT INNER JOIN"
sSQL = sSQL & " USAWaterski.dbo.Level LV ON OT.Level_ID = LV.Level_ID"
sSQL = sSQL & " WHERE OT.DivisionCode in ('AWS','USA')"
sSQL = sSQL & " AND LV.LevelOrderforTemplate IS NOT NULL"
sSQL = sSQL & " AND OT.RatingType_ID = 1 GROUP BY OT.PersonID) as OJ"
sSQL = sSQL & " on OJ.PersonID = MX.PersonID"

sSQL = sSQL & " Left Join (Select OT.PersonID,"
sSQL = sSQL & " Max(convert(char(1),LV.LevelOrderforTemplate)"
sSQL = sSQL & " + LV.LevelAbbreviationforTemplate) AS RtgLvl"
sSQL = sSQL & " FROM USAWaterski.dbo.Officials OT INNER JOIN"
sSQL = sSQL & " USAWaterski.dbo.Level LV ON OT.Level_ID = LV.Level_ID"
sSQL = sSQL & " WHERE OT.DivisionCode in ('AWS','USA')"
sSQL = sSQL & " AND LV.LevelOrderforTemplate IS NOT NULL"
sSQL = sSQL & " AND OT.RatingType_ID = 2 GROUP BY OT.PersonID) as OC"
sSQL = sSQL & " on OC.PersonID = MX.PersonID"

sSQL = sSQL & " Left Join (Select OT.PersonID,"
sSQL = sSQL & " Max(convert(char(1),LV.LevelOrderforTemplate)"
sSQL = sSQL & " + LV.LevelAbbreviationforTemplate) AS RtgLvl"
sSQL = sSQL & " FROM USAWaterski.dbo.Officials OT INNER JOIN"
sSQL = sSQL & " USAWaterski.dbo.Level LV ON OT.Level_ID = LV.Level_ID"
sSQL = sSQL & " WHERE OT.DivisionCode in ('AWS','USA')"
sSQL = sSQL & " AND LV.LevelOrderforTemplate IS NOT NULL"
sSQL = sSQL & " AND OT.RatingType_ID = 9 GROUP BY OT.PersonID) as OS"
sSQL = sSQL & " on OS.PersonID = MX.PersonID"

sSQL = sSQL & " Order By MX.Sorter, MX.LastName, MX.FirstName, MX.MemberID"

' Response.write sSQL

objRS.Open sSQL

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Now we finally loop through the Extract of Members, merging the
''' Team Headers into both Men's and Women's sections, by Team ID.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim LocalTeamID, LocalMenTeam, LocalWomTeam

LocalTeamID = Trim(objTL("TeamID"))

DO until objRS.EOF

	IF objRS("EffTo") < cdate(session("TournamentDate")) or objRS("CanSki") = False or objRS("Waiver") = 0 THEN

		objExcelUpgRnew.addnew
		objExcelUpgRnew.Fields(0).Value = objRS("MemID")
		objExcelUpgRnew.Fields(1).Value = objRS("LastName")
		objExcelUpgRnew.Fields(2).Value = objRS("FirstName")
		objExcelUpgRnew.Fields(3).Value = trim(objRS("Team"))
		objExcelUpgRnew.Fields(4).Value = objRS("Div")
		objExcelUpgRnew.Fields(5).Value = objRS("Age")
		objExcelUpgRnew.Fields(6).Value = objRS("City")
		objExcelUpgRnew.Fields(7).Value = objRS("State")

		objExcelUpgRnew.Fields(11).Value = objRS("OffRat")

		objExcelUpgRnew.Fields(16).Value = objRS("SptsDiv")

		' Figure applicable Renewal / Upgrade Amount based on MemType & Status

		MT = objRS("MemType")
		IF MT < 1 OR MT > 200 THEN MT = 1

		IF objRS("EffTo") < cdate(session("TournamentDate")) THEN 
			IF objRS("CanSki") = False THEN
				objExcelUpgRnew.Fields(17).Value = "Nds Rnw/Upg" 
				objExcelUpgRnew.Fields(18).Value = FormatNumber(MemPrice(MT)+MemUpgrd(MT),2)
			ELSE
				objExcelUpgRnew.Fields(17).Value = "Needs Renew" 
				objExcelUpgRnew.Fields(18).Value = FormatNumber(MemPrice(MT),2)
			END IF
		ELSE 
			IF objRS("CanSki") = False THEN
				objExcelUpgRnew.Fields(17).Value = "Needs Upgrd" 
				objExcelUpgRnew.Fields(18).Value = FormatNumber(MemUpgrd(MT),2)
			ELSE
				objExcelUpgRnew.Fields(17).Value = "Nds Ann Wvr" 
				objExcelUpgRnew.Fields(18).Value = FormatNumber(0,2)
			END IF				
		END IF
		
		objExcelUpgRnew.Update

	ELSE

		DO WHILE LocalTeamID < "zzz" AND LocalTeamID <= trim(objRS("Sorter"))

			objExcelAMen.addnew
			objExcelAMen.Fields(0).Value = " "
			objExcelAMen.Update
			objExcelAMen.addnew
			objExcelAMen.Fields(0).Value = "Team Header"
			objExcelAMen.Fields(1).Value = objTL("TeamName")
			LocalMenTeam = trim(objTL("TeamID"))
			objExcelAMen.Fields(3).Value = LocalMenTeam
			objExcelAMen.Update

			objExcelAWomen.addnew
			objExcelAWomen.Fields(0).Value = " "
			objExcelAWomen.Update	
			objExcelAWomen.addnew
			objExcelAWomen.Fields(0).Value = "Team Header"
			objExcelAWomen.Fields(1).Value = objTL("TeamName")
			LocalWomTeam = trim(objTL("TeamID"))
			objExcelAWomen.Fields(3).Value = LocalWomTeam
			objExcelAWomen.Update	

			objTL.MoveNext
			IF objTL.EOF THEN 
				LocalTeamID = "zzz"
			ELSE 
				LocalTeamID = Trim(objTL("TeamID"))
			END IF

		LOOP

		IF objRS("Div") = "CM" THEN

			IF (trim(objRS("Sorter")) <> "zzz" AND trim(objRS("Sorter")) <> LocalMenTeam) OR _
				 (trim(objRS("Sorter")) = "zzz" AND LocalMenTeam <> "zzz") THEN
				objExcelAMen.addnew
				objExcelAMen.Fields(0).Value = " "
				objExcelAMen.Update
				LocalMenTeam = trim(ObjRS("Sorter"))
			END IF

			objExcelAMen.addnew
			objExcelAMen.Fields(0).Value = objRS("MemID")
			objExcelAMen.Fields(1).Value = objRS("LastName")
			objExcelAMen.Fields(2).Value = objRS("FirstName")
			objExcelAMen.Fields(3).Value = trim(objRS("Team"))
			objExcelAMen.Fields(4).Value = objRS("Div")
			objExcelAMen.Fields(5).Value = objRS("Age")
			objExcelAMen.Fields(6).Value = objRS("City")
			objExcelAMen.Fields(7).Value = objRS("State")

			objExcelAMen.Fields(11).Value = objRS("OffRat")

			objExcelAMen.Fields(16).Value = objRS("SptsDiv")
			objExcelAMen.Fields(17).Value = " OK to Ski"
	
			objExcelAMen.Update

		ELSE

			IF (trim(objRS("Sorter")) <> "zzz" AND trim(objRS("Sorter")) <> LocalWomTeam) OR _
				 (trim(objRS("Sorter")) = "zzz" AND LocalWomTeam <> "zzz") THEN
				objExcelAWomen.addnew
				objExcelAWomen.Fields(0).Value = " "
				objExcelAWomen.Update
				LocalWomTeam = trim(ObjRS("Sorter"))
			END IF

			objExcelAWomen.addnew
			objExcelAWomen.Fields(0).Value = objRS("MemID")
			objExcelAWomen.Fields(1).Value = objRS("LastName")
			objExcelAWomen.Fields(2).Value = objRS("FirstName")
			objExcelAWomen.Fields(3).Value = trim(objRS("Team"))
			objExcelAWomen.Fields(4).Value = objRS("Div")
			objExcelAWomen.Fields(5).Value = objRS("Age")
			objExcelAWomen.Fields(6).Value = objRS("City")
			objExcelAWomen.Fields(7).Value = objRS("State")

			objExcelAWomen.Fields(11).Value = objRS("OffRat")
			
			objExcelAWomen.Fields(16).Value = objRS("SptsDiv")
			objExcelAWomen.Fields(17).Value = " OK to Ski"
			
			objExcelAWomen.Update	

		END IF
	
	END IF
	
	objRS.MoveNext

LOOP


'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""


objExcelAWomen.close
set objExcelAWomen = nothing
objExcelUpgRnew.close
set objExcelUpgRnew = nothing
objExcelConn.close
set objExcelConn = nothing
'
objRS.Close
Set objRS = Nothing

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

'5-18-2006 Remove any strange characters from the TournamentName
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


    
Response.Flush
      
' This final bit of HTML is written after processing is successfully completed
' to tell the user how to download their template, and where to go from here.
      
%>
    
    <SCRIPT LANGUAGE="JavaScript">
    if(upLevel) {
      var splash = document.getElementById("splashScreen");
    }
    else if(ns4) {
      var splash = document.splashScreen;
    }
    else if(ie4) {
      var splash = document.all.splashScreen;
    }
      
    hideObject(splash);
    </SCRIPT>  


<html>

<head>
<title>Create Pre-Registration Export v1.5</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" background = "/images/TopBackground.jpg" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" >

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#42639F">
      <p>&nbsp;</p>
      <p align="center"><font face="Verdana" size="6" color="#FFFFFF">
      	USA Water Ski NCWSA Registration Template</font></p>
      <p align="center"><font face="Verdana" size="4" color="#FFFFFF">
      	Registration Support for -- <%=session("TournamentName")%></font></p>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>  
  
<table border="0" cellspacing="0" cellpadding="0">  
  <tr> 
    <td width="185" valign="top" bgcolor="#42639F">

	<% If Session("aauth") then %>
	<font face="Verdana" size="2" COLOR="#FFFFFF"><br>&nbsp;Currently Logged in as: </font><br>
	<font face="Verdana" size="2" COLOR="#FFFFFF">&nbsp;<%=Session("UserName")%>&nbsp;&nbsp;
		<%=session("TournamentDate")%></font><br>
	<br>
	<% Else %>
	<font face="Verdana" size="2" COLOR="#FFFFFF"><br>&nbsp;Not currently logged in.</font>
	<% End If %>
	
			<font face="Verdana" size="2"> 
         <br>&nbsp;<a href="logout.asp"><font face="arial" COLOR="#FFFFFF">Log Out</font></a>&nbsp;<br>
			</font>
			<br>
	        &nbsp;<a href="/admin/index.asp"><font face="arial" size="2" COLOR="#FFFFFF">Back to Admin Index</font></a><br>&nbsp;<br>
	        &nbsp;<a href="http://www.usawaterski.org"><font face="arial" size="2" COLOR="#FFFFFF">USA Water Ski Home</font></a><br>&nbsp;<br>
			<br>
            <font face="Verdana" size="1">&nbsp;<font COLOR="#FFFFFF">Powered by</font> <a href="http://www.epolk.com"><font COLOR="#FFFFFF">ePolk.com</font></a><br></font>

  </td>

	<td>

  <table>
      <tr> 
         <td width="14">&nbsp;</td>
         <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><br>Your NCWSA
         Registration Template workbook is now complete and ready to download.&nbsp;</font>
         <br>&nbsp;<br>
         <font color="#FF0000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>
         !! New Registration Features for Collegiate Events, as of Fall 2008 !!
         </strong></font><br>&nbsp;<br>
         <font size="2" face="Verdana, Arial, Helvetica, sans-serif">This all-new NCWSA
         Registration Template that we have prepared for you here includes everything you
         should need to handle Registration and Scoring preparations for your Collegiate
         tournament.&nbsp; Since this will likely be your first exposure to these new
         features, <b><i>please</i></b> take a couple of minutes and review the Overview 
         and Checklist that you will find at the top of the Instructions section that you
         will find in the Template file, after you download and open it.
         </font></td>
      </tr>

      <tr> 
         <td>&nbsp;</td>
      </tr>

      <tr> 
         <td>&nbsp;</td>
         <td><a href="excel/<% response.write filenamewithscores %>"><font face="Arial" size="2">RIGHT 
         Click Here</font></a>&nbsp; <font size="2" face="Verdana, Arial, Helvetica, sans-serif">to 
         download your NCWSA Registration Template workbook, then select the "Save As" 
         option from that menu, and then choose a suitable location to store the download 
         file in your PC. </font></td>
      </tr>
   
      <tr> 
         <td>&nbsp;</td>
      </tr>

      <tr> 
         <td>&nbsp;</td>
         <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
         After your Registration Template download has completed, then open the 
         Excel file from that location on your PC.&nbsp; It will open automatically 
         to an Instructions Tab section.&nbsp; Please review the material in that
         section for the latest information on contents and usage.</font></td>
      </tr>

      <tr> 
         <td>&nbsp;</td>
      </tr>

      <tr> 
         <td>&nbsp;</td>
         <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
         After you've downloaded this Registration Template, you can later 
         fold in additional selected members, one-by-one, using the lookup 
         feature noted on the earlier screen.&nbsp; With that feature, you
         can then just copy and paste the information for any additional 
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
		
		<TR>

	    <TD width=30% align=center>
		<form action="LookupMembers.asp?FormStatus=newsearch" method="post">
		<input type="submit" style="width:9em" value="Lookup Members"></form>
    	</TD>

	    <td width=30% align=center>     				
		<form action="Index.asp" method="post">
    <input type="submit" style="width:9em" value="Quit"></form>
 	    </td>
  	    
 	  </TR>

 	</TABLE>

  	  </td>
	  </tr>
</table>
</body>
</html>






