<!--#include virtual="/epl/functions.asp" -->

<% 

If not Session("aauth") then response.redirect "Login.asp"

Server.ScriptTimeout = 600

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

' ---------------------------------------------------
   FUNCTION PersonIDwChkDgt (PersonID)
' ---------------------------------------------------

' This function is given an integer "PersonID" value, and returns the
' 9-Character "PersonIDWithCheckDigit" value for that particular member.

Dim PIDSum, PIDChar, PIDLen, PIDPtr
 
PIDSum = 0: PIDChar = trim(PersonID): PIDLen = Len(PIDChar)

FOR PIDPtr = 1 TO PIDLen STEP 2
	PIDSum = PIDSum + (3*MID(PIDChar,PIDPtr,1))
	IF PIDPtr+1 <= PIDLen THEN PIDSum = PIDSum + MID(PIDChar,PIDPtr+1,1)
	NEXT

PersonIDwChkDgt = right(100-PIDSum,1) & Right(100000000+PersonID,8)

END FUNCTION


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

' objConn.ConnectionTimeout = 0

objConn.CommandTimeout = 300 

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

'Get TStatus and TSanction from TSchedul table, and AllowAccess from Users999
Dim strTStatus, strTSanction, AllowAccess
sSQL = "Select Top 1 TS.TSanction, TS.TStatus, US.AllowAccess"
sSQL = sSQL & " from Sanctions.dbo.TSchedul as TS, USAWaterski.dbo.Users999 as US"
sSQL = sSQL & " where TS.TournAppID = '" & left(Session("TournamentID"),6)
sSQL = sSQL & "' and US.Name = '" & left(Session("TournamentID"),6) & "'"

objRS.Open sSQL
If objRS.EOF THEN
	strTStatus = -1: strTSanction = Session("TournamentID"): AllowAccess = false
ELSE 
	strTStatus = objRS("TStatus"): strTSanction = objRS("TSanction"): AllowAccess = objRS("AllowAccess")
	IF left(strTSanction,6) <> left(Session("TournamentID"),6) THEN
		strTSanction = Session("TournamentID")
	END IF
END IF
objRS.Close



'""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'"""""""""""""" With Scores and Ratings """""""""""""""""""""""
'""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""


'objFSO.CopyFile path & "/Templates/NCWSATemplateBlank.xls", path & "/template_with_scores.xls" , True
objFSO.CopyFile path & "/Templates/NCWSATemplate2012.xls", path & "/template_with_scores.xls" , True

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
objExcelSingleFields.Fields(0).Value = strTSanction	'this is the same as the tournament ID
objExcelSingleFields.update
objExcelSingleFields.close
		
objExcelSingleFields.Source = "Select * from AWomenTourName"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = session("TournamentName")
objExcelSingleFields.update
objExcelSingleFields.close
		
objExcelSingleFields.Source = "Select * from AWomenTourID"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = strTSanction	'this is the same as the tournament ID
objExcelSingleFields.update
objExcelSingleFields.close
		
objExcelSingleFields.Source = "Select * from BMenTourName"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = session("TournamentName")
objExcelSingleFields.update
objExcelSingleFields.close
		
objExcelSingleFields.Source = "Select * from BMenTourID"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = strTSanction	'this is the same as the tournament ID
objExcelSingleFields.update
objExcelSingleFields.close
		
objExcelSingleFields.Source = "Select * from BWomenTourName"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = session("TournamentName")
objExcelSingleFields.update
objExcelSingleFields.close
		
objExcelSingleFields.Source = "Select * from BWomenTourID"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = strTSanction	'this is the same as the tournament ID
objExcelSingleFields.update
objExcelSingleFields.close
		
objExcelSingleFields.Source = "Select * from RegistTourName"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = session("TournamentName")
objExcelSingleFields.update
objExcelSingleFields.close
		
objExcelSingleFields.Source = "Select * from RegistTourID"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = strTSanction
objExcelSingleFields.update
objExcelSingleFields.close
		
objExcelSingleFields.Source = "Select * from AllOthrTourName"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = session("TournamentName")
objExcelSingleFields.update
objExcelSingleFields.close
		
objExcelSingleFields.Source = "Select * from AllOthrTourID"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = strTSanction
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

Set objExcelBMen = Server.CreateObject("ADODB.Recordset")
objExcelBMen.ActiveConnection = objExcelConn 
objExcelBMen.CursorType = 3                    'Static cursor.
objExcelBMen.LockType = 2                      'Pessimistic Lock.
objExcelBMen.Source = "Select * from BMenRange"
objExcelBMen.Open

Set objExcelBWomen = Server.CreateObject("ADODB.Recordset")
objExcelBWomen.ActiveConnection = objExcelConn 
objExcelBWomen.CursorType = 3                    'Static cursor.
objExcelBWomen.LockType = 2                      'Pessimistic Lock.
objExcelBWomen.Source = "Select * from BWomenRange"
objExcelBWomen.Open

Set objExcelRegist = Server.CreateObject("ADODB.Recordset")
objExcelRegist.ActiveConnection = objExcelConn 
objExcelRegist.CursorType = 3                    'Static cursor.
objExcelRegist.LockType = 2                      'Pessimistic Lock.
objExcelRegist.Source = "Select * from RegistRange"
objExcelRegist.Open

Set objExcelAllOthr = Server.CreateObject("ADODB.Recordset")
objExcelAllOthr.ActiveConnection = objExcelConn 
objExcelAllOthr.CursorType = 3                    'Static cursor.
objExcelAllOthr.LockType = 2                      'Pessimistic Lock.
objExcelAllOthr.Source = "Select * from AllOthrRange"
objExcelAllOthr.Open



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Next we insert Chief and Appointed official Person ID's for the 
''' desired Tournament, from the Sanctions.Registration table into 
''' a work table, along with Applicable Chief Codes.  But first we
''' need to do a delete of any existing rows for that TournAppID.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim sSQL

sSQL = "Delete from USAWaterski.dbo.TempApptdOfcls where TournAppID = '" 
sSQL = sSQL & left(Session("TournamentID"),6) & "' OR DateAdd(Day,30,WhenAdded) < GetDate()"
objConn.Execute (sSQL)

sSQL = "Insert into USAWaterski.dbo.TempApptdOfcls (PersonID, TournAppID, OffCode, WhenAdded)"
sSQL = sSQL & " Select PersonID, '"& left(Session("TournamentID"),6)
sSQL = sSQL & "', Max(OffCode), GetDate() from ("

sSQL = sSQL & " Select Cast(case when len(CJudgePID)<9 then CJudgePID else"
sSQL = sSQL & " right(CJudgePID,8) end as integer) AS PersonID, 'CJ' AS OffCode"
sSQL = sSQL & " FROM sanctions.dbo.registration WHERE TournAppID = '"
sSQL = sSQL & left(Session("TournamentID"),6) & "' and isnumeric(CJudgePID) = 1 UNION"

sSQL = sSQL & " Select Cast(case when len(CDriverPID)<9 then CDriverPID else"
sSQL = sSQL & " right(CDriverPID,8) end as integer) AS PersonID, 'CD' AS OffCode"
sSQL = sSQL & " FROM sanctions.dbo.registration WHERE TournAppID = '"
sSQL = sSQL & left(Session("TournamentID"),6) & "' and isnumeric(CDriverPID) = 1 UNION"

sSQL = sSQL & " Select Cast(case when len(CScorePID)<9 then CScorePID else"
sSQL = sSQL & " right(CScorePID,8) end as integer) AS PersonID, 'CC' AS OffCode"
sSQL = sSQL & " FROM sanctions.dbo.registration WHERE TournAppID = '"
sSQL = sSQL & left(Session("TournamentID"),6) & "' and isnumeric(CScorePID) = 1 UNION"

sSQL = sSQL & " Select Cast(case when len(CSafPID)<9 then CSafPID else"
sSQL = sSQL & " right(CSafPID,8) end as integer) AS PersonID, 'CS' AS OffCode"
sSQL = sSQL & " FROM sanctions.dbo.registration WHERE TournAppID = '"
sSQL = sSQL & left(Session("TournamentID"),6) & "' and isnumeric(CSafPID) = 1 UNION"

sSQL = sSQL & " Select Cast(case when len(TechCPID)<9 then TechCPID else"
sSQL = sSQL & " right(TechCPID,8) end as integer) AS PersonID, 'CT' AS OffCode"
sSQL = sSQL & " FROM sanctions.dbo.registration WHERE TournAppID = '"
sSQL = sSQL & left(Session("TournamentID"),6) & "' and isnumeric(TechCPID) = 1 UNION"

sSQL = sSQL & " Select Cast(case when len(Ap1JPID)<9 then Ap1JPID else"
sSQL = sSQL & " right(Ap1JPID,8) end as integer) AS PersonID, 'APTJ' AS OffCode"
sSQL = sSQL & " FROM sanctions.dbo.registration WHERE TournAppID = '"
sSQL = sSQL & left(Session("TournamentID"),6) & "' and isnumeric(Ap1JPID) = 1 UNION"

sSQL = sSQL & " Select Cast(case when len(Ap2JPID)<9 then Ap2JPID else"
sSQL = sSQL & " right(Ap2JPID,8) end as integer) AS PersonID, 'APTJ' AS OffCode"
sSQL = sSQL & " FROM sanctions.dbo.registration WHERE TournAppID = '"
sSQL = sSQL & left(Session("TournamentID"),6) & "' and isnumeric(Ap2JPID) = 1 UNION"

sSQL = sSQL & " Select Cast(case when len(Ap3JPID)<9 then Ap3JPID else"
sSQL = sSQL & " right(Ap3JPID,8) end as integer) AS PersonID, 'APTJ' AS OffCode"
sSQL = sSQL & " FROM sanctions.dbo.registration WHERE TournAppID = '"
sSQL = sSQL & left(Session("TournamentID"),6) & "' and isnumeric(Ap3JPID) = 1 UNION"

sSQL = sSQL & " Select Cast(case when len(Ap4JPID)<9 then Ap4JPID else"
sSQL = sSQL & " right(Ap4JPID,8) end as integer) AS PersonID, 'APTJ' AS OffCode"
sSQL = sSQL & " FROM sanctions.dbo.registration WHERE TournAppID = '"
sSQL = sSQL & left(Session("TournamentID"),6) & "' and isnumeric(Ap4JPID) = 1 UNION"

sSQL = sSQL & " Select Cast(case when len(Ap5JPID)<9 then Ap5JPID else"
sSQL = sSQL & " right(Ap5JPID,8) end as integer) AS PersonID, 'APTJ' AS OffCode"
sSQL = sSQL & " FROM sanctions.dbo.registration WHERE TournAppID = '"
sSQL = sSQL & left(Session("TournamentID"),6) & "' and isnumeric(Ap5JPID) = 1 UNION"

sSQL = sSQL & " Select Cast(case when len(Ap1SPID)<9 then Ap1SPID else"
sSQL = sSQL & " right(Ap1SPID,8) end as integer) AS PersonID, 'APTS' AS OffCode"
sSQL = sSQL & " FROM sanctions.dbo.registration WHERE TournAppID = '"
sSQL = sSQL & left(Session("TournamentID"),6) & "' and isnumeric(Ap1SPID) = 1 UNION"

sSQL = sSQL & " Select Cast(case when len(Ap2SPID)<9 then Ap2SPID else"
sSQL = sSQL & " right(Ap2SPID,8) end as integer) AS PersonID, 'APTS' AS OffCode"
sSQL = sSQL & " FROM sanctions.dbo.registration WHERE TournAppID = '"
sSQL = sSQL & left(Session("TournamentID"),6) & "' and isnumeric(Ap2SPID) = 1 UNION"

sSQL = sSQL & " Select Cast(case when len(Ap3SPID)<9 then Ap3SPID else"
sSQL = sSQL & " right(Ap3SPID,8) end as integer) AS PersonID, 'APTS' AS OffCode"
sSQL = sSQL & " FROM sanctions.dbo.registration WHERE TournAppID = '"
sSQL = sSQL & left(Session("TournamentID"),6) & "' and isnumeric(Ap3SPID) = 1 UNION"

sSQL = sSQL & " Select Cast(case when len(Ap1DrPID)<9 then Ap1DrPID else"
sSQL = sSQL & " right(Ap1DrPID,8) end as integer) AS PersonID, 'APTD' AS OffCode"
sSQL = sSQL & " FROM sanctions.dbo.registration WHERE TournAppID = '"
sSQL = sSQL & left(Session("TournamentID"),6) & "' and isnumeric(Ap1DrPID) = 1 UNION"

sSQL = sSQL & " Select Cast(case when len(PanAmPID)<9 then PanAmPID else"
sSQL = sSQL & " right(PanAmPID,8) end as integer) AS PersonID, 'APTJ' AS OffCode"
sSQL = sSQL & " FROM sanctions.dbo.registration WHERE TournAppID = '"
sSQL = sSQL & left(Session("TournamentID"),6) & "' and isnumeric(PanAmPID) = 1)"

sSQL = sSQL & " SOX Group by PersonID"
objConn.Execute (sSQL)



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Next we open an extract of Team ID's and Names from the TeamList table.
' Note that we prefix each team name with "E" if the team has entries,
' or "Z" if no entries, so that all the entered teams list at the top.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sSQL = "Select Case when TE.Team is Null then 'Z'+TL.TeamID"
sSQL = sSQL & " else 'E'+TL.TeamID end as TeamID, TL.TeamName"
sSQL = sSQL & " from Cobra00025.USAWSRank.TeamsList as TL"
sSQL = sSQL & " left join (Select distinct team"
sSQL = sSQL & " from Cobra00025.USAWSRank.TeamRotations where"
sSQL = sSQL & " TournAppID = '" & left(strTSanction,6) 
sSQL = sSQL & "' and WaiverStat >= 'C') as TE"
sSQL = sSQL & " on TE.Team = TL.TeamID Where SptsGrpID = 'NCW'"
sSQL = sSQL & " Order by Case when TE.Team is Null then"
sSQL = sSQL & " 'Z'+TL.TeamID else 'E'+TL.TeamID end"

objTL.Open sSQL


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Now build a Query to Extract the Desired Members, joining in data 
''' pulled from the Rankings and Officials and Membership Type tables.
''' Note that we prefix each team ID with "E" if the team has entries,
''' or "Z" if no entries, so that all the entered teams list at the top,
''' then finally all those without any team affiliation last with Zzzz.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'	Begin with the overall select list for the Outer Query

sSQL = "Select MX.PersonID, MX.LastName, MX.FirstName,"
sSQL = sSQL & " Case when MX.Sex = 'F' Then 'CW' else 'CM' end as Div,"

sSQL = sSQL & " Case when MX.Age <= 17 and MX.Sex = 'F' Then 'G'"
sSQL = sSQL & " when MX.Age <= 17 then 'B' when MX.Sex = 'F' then 'W' else 'M' end + Case"
sSQL = sSQL & " when MX.Age <= 9 then '1' when MX.Age <= 13 then '2' when MX.Age <= 17 then '3'"
sSQL = sSQL & " when MX.Age <= 24 then '1' when MX.Age <= 34 then '2' when MX.Age <= 44 then '3'"
sSQL = sSQL & " when MX.Age <= 52 then '4' when MX.Age <= 59 then '5' when MX.Age <= 64 then '6'"
sSQL = sSQL & " when MX.Age <= 69 then '7' when MX.Age <= 74 then '8' when MX.Age <= 79 then '9'"
sSQL = sSQL & " when MX.Age <= 84 then 'A' else 'B' end as AgeDiv,"
		
sSQL = sSQL & " Case when SO.OffCode is not Null and MX.SlmEnt+MX.TrkEnt+MX.JmpEnt"
sSQL = sSQL & " = '      ' then 'E0FF' else MX.Sorter end as Sorter,"

sSQL = sSQL & " MX.Team, MX.TeamStat, MX.Sex, MX.Age, MX.City, MX.State,"

sSQL = sSQL & " Case when OD.PersonID is Null then '-' else Right(OD.RtgLvl,1) end +"
sSQL = sSQL & " Case when OJ.PersonID is Null then '-' else Right(OJ.RtgLvl,1) end +"
sSQL = sSQL & " Case when OC.PersonID is Null then '-' else Right(OC.RtgLvl,1) end +"
sSQL = sSQL & " Case when OS.PersonID is Null then '-' else Right(OS.RtgLvl,1) end as OffRat,"

sSQL = sSQL & " Coalesce(SO.OffCode,'') as OffCode,"

sSQL = sSQL & " MX.EffTo, MX.Memtype, MX.MemCode, MX.CanSki, MX.SptsDiv, MX.AnnWvr,"
sSQL = sSQL & " MX.EvtWvr, MX.SlmEnt, MX.TrkEnt, MX.JmpEnt"


'	This begins the major "MX" Sub-query, which pulls membership and team and entry information
		
sSQL = sSQL & " From (Select MT.PersonID,"
sSQL = sSQL & " Left(MT.LastName,12) as LastName, Left(MT.FirstName,10) as FirstName,"

sSQL = sSQL & " (" & Session("TournamentYear") & "-Year(MT.BirthDate)-1) as Age,"

sSQL = sSQL & " Left(MT.City,12) as City, Left(MT.State,2) as State,"
sSQL = sSQL & " MT.EffectiveTo as EffTo, MT.MembershipTypeCode as MemType,"
sSQL = sSQL & " Typ.TypeCode as MemCode, Typ.CanSkiInTournaments as CanSki,"
sSQL = sSQL & " MT.DivisionCode1 + '/' + MT.DivisionCode2 as SptsDiv,"
sSQL = sSQL & " Upper(Left(MT.Sex,1)) as Sex, MT.WaiverStatusID as AnnWvr,"

sSQL = sSQL & " Case when TE.Team is not null then 'E' else 'Z' end +"
sSQL = sSQL & " Case when Coalesce(RP.Team,TR.Team) is not null then"
sSQL = sSQL & " Coalesce(RP.Team,TR.Team) else 'zzz' end as Sorter,"

sSQL = sSQL & " Case when RP.MemberID is not null then 'A' when"
sSQL = sSQL & " TR.DateInactive is not null then 'I' else 'A' end as TeamStat,"

sSQL = sSQL & " Coalesce(RP.Team,TR.Team,'   ') as Team,"
sSQL = sSQL & " Coalesce(RP.SlalomEnt,'  ') as SlmEnt," 
sSQL = sSQL & " Coalesce(RP.TrickEnt,'  ') as TrkEnt," 
sSQL = sSQL & " Coalesce(RP.JumpEnt,'  ') as JmpEnt," 
sSQL = sSQL & " Coalesce(RP.WaiverStat,' ') as EvtWvr" 

'	Begin FROM and JOIN table list for "MX" Sub-Query

sSQL = sSQL & " FROM USAWaterski.dbo.MembersLive as MT Inner Join"
sSQL = sSQL & " USAWaterski.dbo.MembershipTypes as Typ"
sSQL = sSQL & " ON MT.MembershipTypeCode = Typ.MemberShipTypeID"


'	Here's the subquery which now pulls Team ID's from the Team Roster Extract.
'	Identify Latest Team affiliation for Member -- new version
sSQL = sSQL & " Left Join (Select RX.MemberID, RX.Team, RX.DateInactive"
sSQL = sSQL & " from Cobra00025.USAWSRank.TeamRoster as RX"
sSQL = sSQL & " join (select MemberID, Max(LastEvent) as MaxEvt"
sSQL = sSQL & " from Cobra00025.USAWSRank.TeamRoster group by MemberID) as ME" 
sSQL = sSQL & " on ME.MemberID = RX.MemberID and ME.MaxEvt = RX.LastEvent) as TR"
sSQL = sSQL & " on cast(right(TR.MemberID,8) as integer) = MT.PersonID"

'	This subquery pulls Rotation Plan information for this Person/TourID -- LEAVE TEAM OUT !! (All Stars)
sSQL = sSQL & " left join Cobra00025.USAWSRank.TeamRotations as RP"
sSQL = sSQL & " on RP.TournAppID = '" & left(strTSanction,6) & "'"
sSQL = sSQL & " and cast(right(RP.MemberID,8) as integer) = MT.PersonID"

'	This subquery identifies Teams that are Entered, used to preface Sorter extract column
sSQL = sSQL & " left join (Select distinct team" 
sSQL = sSQL & " from Cobra00025.USAWSRank.TeamRotations where"
sSQL = sSQL & " WaiverStat >= 'C' and TournAppID = '" & left(strTSanction,6)
sSQL = sSQL & "') as TE on TE.Team = Coalesce(RP.Team,TR.Team)"

' Now here's the "WHERE" condition clause for the Primary "MX" Sub-Query
sSQL = sSQL & " Where Typ.ExporttoTouramentRegistrationTemplate = 1"
sSQL = sSQL & " AND DateAdd(mm,18,MT.EffectiveTo) > GetDate()"
sSQL = sSQL & " AND MT.Deceased = 0 AND ( (" & Session("TournamentYear")
sSQL = sSQL & " - Year(MT.BirthDate) - 1) between 16 and 29 OR"
sSQL = sSQL & " MT.DivisionCode1 = 'NCW' OR MT.DivisionCode2 = 'NCW' OR"

sSQL = sSQL & " PersonID in (Select PersonID from USAWaterski.dbo.TempApptdOfcls"
sSQL = sSQL & " Where TournAppID = '" & left(Session("TournamentID"),6) & "') OR"

sSQL = sSQL & " PersonID IN (Select Distinct cast(right(MemberID,8) as integer) from"
sSQL = sSQL & " Cobra00025.USAWSRank.TeamRoster) ) ) as MX" 

'	End of MX Primary "MX" Select Subquery.  Appended Info Subqueries follow.

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

sSQL = sSQL & " Left Join	(Select PersonID, OffCode from USAWaterski.dbo.TempApptdOfcls"
sSQL = sSQL & " Where TournAppID = '" & left(Session("TournamentID"),6) & "')"
sSQL = sSQL & " as SO on SO.PersonID = MX.PersonID"

sSQL = sSQL & " Order By Case when SO.OffCode is not Null and MX.SlmEnt+MX.TrkEnt+MX.JmpEnt"
sSQL = sSQL & " = '      ' then 'E0FF' else MX.Sorter end,"
sSQL = sSQL & " MX.LastName, MX.FirstName, MX.PersonID"

' Response.write sSQL

Response.Write("Immediately before SQL Open " & Server.ScriptTimeout)

objRS.Open sSQL


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Start by plugging in an "Officials" Team Header into B Men's Section
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

objExcelBMen.addnew
objExcelBMen.Fields(0).Value = " "
objExcelBMen.Update
objExcelBMen.addnew
objExcelBMen.Fields(0).Value = "Team Header"
objExcelBMen.Fields(1).Value = "Officials & AWSA Skiers"
objExcelBMen.Fields(3).Value = "OFF"
objExcelBMen.Update


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Now we finally loop through the Extract of Members, merging the
''' Team Headers into All sections, by Team ID.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'	Now here's the scoop on the primary variables coming in and how we split things up
' Sorter is now four positions, with "E" at the front if member of Entered Team, otherwise "Z"
' Rest of Sorter indicates team, if any.  "zzz" for non-team skiers.  " OF" are apptd Officials

' as to section splits.  Sorter Prefix is "Z" then all go to OTHERS section
' Where sorter prefix is "E", then if TeamStat is "I" then goes to OTHERS,
' For xxxEnt stat "DD", we need to compute age div code.

Dim NextTeamID, LocalTeamID, LocalTeamCd
Dim NextRotSlmMen, NextRotTrkMen, NextRotJmpMen
Dim NextRotSlmWom, NextRotTrkWom, NextRotJmpWom

LocalTeamID = "E0FF"
LocalTeamCode = "OFF"
LastTeamName = "Appointed Officials"
NextTeamID = Trim(objTL("TeamID"))
TeamSlm = 0
TeamTrk = 0
TeamJmp = 0
TeamTot = 0
GrandSlm = 0
GrandTrk = 0
GrandJmp = 0
GrandTot = 0

DO until objRS.EOF
	
	' First we synthesize MemID value, from incoming PersonID, using internal function.

	MemberID = PersonIDwChkDgt(objRS("PersonID"))
	MemID = left(MemberID,3) & "-" & Mid(MemberID,4,2) & "-" & right(MemberID,4)


	'	Next step is to derive "Reason Not OK to Ski" and renew/upgrade amount strings
	
	MT = objRS("MemType")
	IF MT < 1 OR MT > 200 THEN MT = 1

	IF objRS("EffTo") < cdate(session("TournamentDate")) THEN 
		IF objRS("CanSki") = False THEN
			OKtoSki = "Nds Rnw/Upg" 
			UpgrdAmt = FormatNumber(MemPrice(MT)+MemUpgrd(MT),2)
		ELSE
			OKtoSki = "Needs Renew" 
			UpgrdAmt = FormatNumber(MemPrice(MT),2)
		END IF
	ELSE 
		IF objRS("CanSki") = False THEN
			OKtoSki = "Needs Upgrd" 
			UpgrdAmt = FormatNumber(MemUpgrd(MT),2)
		ELSEIF objRS("AnnWvr") = 0 THEN
			OKtoSki = "Nds Ann Wvr" 
			UpgrdAmt = ""
		ELSEIF objRS("EvtWvr") <> "X" THEN
			OKtoSki = "Nds Evt Wvr" 
			UpgrdAmt = ""
		ELSE
			OKtoSki = "" 
			UpgrdAmt = ""
		END IF				
	END IF
	

	'	First step is to see if we've got a new Team here.
	'	Put into all sections if Prefix is "E", otherwise only into All Other

	DO WHILE NextTeamID < "Zzzz" AND NextTeamID <= trim(objRS("Sorter"))

		'	Put out Registrar totals for last entered team we just finished
		
		'	IF left(LocalTeamID,1) = "E" and LocalTeamID <> "E0FF" THEN
		IF left(LocalTeamID,1) = "E" THEN
		
			objExcelRegist.addnew
			objExcelRegist.Fields(0).Value = "Team Totals"
			objExcelRegist.Fields(1).Value = LastTeamName
			objExcelRegist.Fields(3).Value = LocalTeamCd
			objExcelRegist.Fields(8).Value = TeamSlm
			objExcelRegist.Fields(9).Value = TeamTrk
			objExcelRegist.Fields(10).Value = TeamJmp
			objExcelRegist.Fields(11).Value = "<Rds"
			objExcelRegist.Fields(14).Value = "Skrs>"
			objExcelRegist.Fields(15).Value = TeamTot
			objExcelRegist.Update	
			objExcelRegist.addnew
			objExcelRegist.Fields(0).Value = " "
			objExcelRegist.Update	
	
			GrandSlm = GrandSlm + TeamSlm
			GrandTrk = GrandTrk + TeamTrk
			GrandJmp = GrandJmp + TeamJmp
			GrandTot = GrandTot + TeamTot
			TeamSlm = 0
			TeamTrk = 0
			TeamJmp = 0
			TeamTot = 0			

		END IF	

		LocalTeamID = trim(objTL("TeamID"))
		LocalTeamCd = mid(LocalTeamID,2,len(LocalTeamID)-1)

		IF left(LocalTeamID,1) = "E" THEN

			IF LocalTeamID <> "E0FF" THEN
				
				objExcelRegist.addnew
				objExcelRegist.Fields(0).Value = "Team Header"
				objExcelRegist.Fields(1).Value = objTL("TeamName")
				objExcelRegist.Fields(3).Value = LocalTeamCd
				objExcelRegist.Fields(4).Value = "CM"
				objExcelRegist.Fields(8).Value = "RD"
				objExcelRegist.Update

				objExcelRegist.addnew
				objExcelRegist.Fields(0).Value = "Team Header"
				objExcelRegist.Fields(1).Value = objTL("TeamName")
				objExcelRegist.Fields(3).Value = LocalTeamCd
				objExcelRegist.Fields(4).Value = "CW"
				objExcelRegist.Fields(8).Value = "RD"
				objExcelRegist.Update

			END IF

			objExcelAMen.addnew
			objExcelAMen.Fields(0).Value = " "
			objExcelAMen.Update
			objExcelAMen.addnew
			objExcelAMen.Fields(0).Value = "Team Header"
			objExcelAMen.Fields(1).Value = objTL("TeamName")
			objExcelAMen.Fields(3).Value = LocalTeamCd
			objExcelAMen.Update

			objExcelAWomen.addnew
			objExcelAWomen.Fields(0).Value = " "
			objExcelAWomen.Update	
			objExcelAWomen.addnew
			objExcelAWomen.Fields(0).Value = "Team Header"
			objExcelAWomen.Fields(1).Value = objTL("TeamName")
			objExcelAWomen.Fields(3).Value = LocalTeamCd
			objExcelAWomen.Update	

			objExcelBMen.addnew
			objExcelBMen.Fields(0).Value = " "
			objExcelBMen.Update
			objExcelBMen.addnew
			objExcelBMen.Fields(0).Value = "Team Header"
			objExcelBMen.Fields(1).Value = objTL("TeamName")
			objExcelBMen.Fields(3).Value = LocalTeamCd
			objExcelBMen.Update

			objExcelBWomen.addnew
			objExcelBWomen.Fields(0).Value = " "
			objExcelBWomen.Update	
			objExcelBWomen.addnew
			objExcelBWomen.Fields(0).Value = "Team Header"
			objExcelBWomen.Fields(1).Value = objTL("TeamName")
			objExcelBWomen.Fields(3).Value = LocalTeamCd
			objExcelBWomen.Update
			
			NextRotSlmMen = 6: NextRotTrkMen = 6: NextRotJmpMen = 6
			NextRotSlmWom = 6: NextRotTrkWom = 6: NextRotJmpWom = 6
			
			LastTeamName = objTL("TeamName")

		END IF
		
		objExcelAllOthr.addnew
		objExcelAllOthr.Fields(0).Value = " "
		objExcelAllOthr.Update	
		objExcelAllOthr.addnew
		objExcelAllOthr.Fields(0).Value = "Team Header"
		objExcelAllOthr.Fields(1).Value = objTL("TeamName")
		objExcelAllOthr.Fields(3).Value = LocalTeamCd
		objExcelAllOthr.Update	

		objTL.MoveNext
		IF objTL.EOF THEN 
			NextTeamID = "Zzzz"
		ELSE 
			NextTeamID = Trim(objTL("TeamID"))
		END IF

	LOOP


	'	Next we store this skier in the "Registrar" section, if an active member of an entered team.
	
'	IF left(objRS("Sorter"),1) = "E" and objRS("Sorter") <> "E0FF" and (objRS("SlmEnt") <> "  " or objRS("TrkEnt") <> "  " or objRS("JmpEnt") <> "  ") THEN
'	IF left(objRS("Sorter"),1) = "E" and objRS("Sorter") <> "E0FF" and objRS("TeamStat") = "A" THEN
	IF left(objRS("Sorter"),1) = "E" and (objRS("TeamStat") = "A" or objRS("OffCode") <> "") THEN

		NumEvts = 0
		objExcelRegist.addnew
		objExcelRegist.Fields(0).Value = MemID
		objExcelRegist.Fields(1).Value = objRS("LastName")
		objExcelRegist.Fields(2).Value = objRS("FirstName")

		IF objRS("Sorter") <> "E0FF" THEN
			objExcelRegist.Fields(3).Value = trim(objRS("Team"))
		ELSE
			objExcelRegist.Fields(3).Value = "OFF"
		END IF

		IF objRS("SlmEnt") = "DD" or objRS("TrkEnt") = "DD" or objRS("JmpEnt") = "DD" THEN
			objExcelRegist.Fields(4).Value = objRS("AgeDiv")
		ELSE
			objExcelRegist.Fields(4).Value = objRS("Div")
		END IF

		objExcelRegist.Fields(5).Value = objRS("Age")
		objExcelRegist.Fields(6).Value = objRS("City")
		objExcelRegist.Fields(7).Value = objRS("State")
		objExcelRegist.Fields(8).Value = objRS("SlmEnt")
		objExcelRegist.Fields(9).Value = objRS("TrkEnt")
		objExcelRegist.Fields(10).Value = objRS("JmpEnt")

		IF left(objRS("OffCode"),1) = "C" THEN
			objExcelRegist.Fields(11).Value = objRS("OffCode")
		ELSE
			objExcelRegist.Fields(11).Value = objRS("OffRat")
		END IF

		objExcelRegist.Fields(12).Value = objRS("SptsDiv")
		objExcelRegist.Fields(13).Value = OKtoSki
		objExcelRegist.Fields(14).Value = UpgrdAmt

		IF objRS("SlmEnt") > "  " THEN 
			NumEvts = NumEvts + 1
			TeamSlm = TeamSlm + 1
		END IF

		IF objRS("TrkEnt") > "  " THEN 
			NumEvts = NumEvts + 1
			TeamTrk = TeamTrk + 1
		END IF

		IF objRS("JmpEnt") > "  " THEN 
			NumEvts = NumEvts + 1
			TeamJmp = TeamJmp + 1
		END IF

		
		IF NumEvts > 0 THEN 
			objExcelRegist.Fields(15).Value = NumEvts
			TeamTot = TeamTot + 1
		END IF

		objExcelRegist.Update
			
	END IF


	'	Now we handle detail skier rows for the current LocalTeamID
	'	First primary split is whether this is row goes to actives A/B team or not
	'	Team must be Entered, and Skier Active AND Entered in at least one event.
	
	IF objRS("Sorter") = "E0FF" or (left(objRS("Sorter"),1) = "E" and objRS("TeamStat") = "A" and (objRS("SlmEnt") <> "  " or objRS("TrkEnt") <> "  " or objRS("JmpEnt") <> "  ")) THEN

		'	Active A/B Team section -- add to Both sections with Applicable Entry codes
		'	Unless "E0FF" without Entry Codes, which goes to B Men team only.

		IF objRS("Div") = "CW" and objRS("Sorter") <> "E0FF" THEN
		
			'	*******	A Women Entries done here ...

			objExcelAWomen.addnew
			objExcelAWomen.Fields(0).Value = MemID
			objExcelAWomen.Fields(1).Value = objRS("LastName")
			objExcelAWomen.Fields(2).Value = objRS("FirstName")
			objExcelAWomen.Fields(3).Value = trim(objRS("Team"))
			objExcelAWomen.Fields(4).Value = objRS("Div")
			objExcelAWomen.Fields(5).Value = objRS("Age")
			objExcelAWomen.Fields(6).Value = objRS("City")
			objExcelAWomen.Fields(7).Value = objRS("State")
			objExcelAWomen.Fields(16).Value = objRS("SptsDiv")
			objExcelAWomen.Fields(17).Value = OKtoSki
			objExcelAWomen.Fields(18).Value = UpgrdAmt
			
			IF left(objRS("SlmEnt"),1) = "A" THEN
				IF right(objRS("SlmEnt"),1) = "0" THEN							
					objExcelAWomen.Fields(8).Value = "10"
				ELSE
					objExcelAWomen.Fields(8).Value = Right(objRS("SlmEnt"),1)
				END IF 
			END IF
				
			IF left(objRS("TrkEnt"),1) = "A" THEN
				IF right(objRS("TrkEnt"),1) = "0" THEN							
					objExcelAWomen.Fields(9).Value = "10"
				ELSE
					objExcelAWomen.Fields(9).Value = Right(objRS("TrkEnt"),1)
				END IF
			END IF
				
			IF left(objRS("JmpEnt"),1) = "A" THEN
				IF right(objRS("JmpEnt"),1) = "0" THEN							
					objExcelAWomen.Fields(10).Value = "10"
				ELSE
					objExcelAWomen.Fields(10).Value = Right(objRS("JmpEnt"),1)
				END IF 
			END IF

			IF left(objRS("OffCode"),1) = "C" THEN
				objExcelAWomen.Fields(11).Value = objRS("OffCode")
			ELSE
				objExcelAWomen.Fields(11).Value = objRS("OffRat")
			END IF

			objExcelAWomen.Update	

			'	*******	B Women Entries done here ...

			objExcelBWomen.addnew
			objExcelBWomen.Fields(0).Value = MemID
			objExcelBWomen.Fields(1).Value = objRS("LastName")
			objExcelBWomen.Fields(2).Value = objRS("FirstName")
			objExcelBWomen.Fields(3).Value = trim(objRS("Team"))

			IF objRS("SlmEnt") = "DD" or objRS("TrkEnt") = "DD" or objRS("JmpEnt") = "DD" THEN
				objExcelBWomen.Fields(4).Value = objRS("AgeDiv")
			ELSE
				objExcelBWomen.Fields(4).Value = objRS("Div")
			END IF

			objExcelBWomen.Fields(5).Value = objRS("Age")
			objExcelBWomen.Fields(6).Value = objRS("City")
			objExcelBWomen.Fields(7).Value = objRS("State")
			objExcelBWomen.Fields(16).Value = objRS("SptsDiv")
			objExcelBWomen.Fields(17).Value = OKtoSki
			objExcelBWomen.Fields(18).Value = UpgrdAmt
			
			IF left(objRS("SlmEnt"),1) = "B" and right(objRS("SlmEnt"),1) <> "6" THEN
				objExcelBWomen.Fields(8).Value = Right(objRS("SlmEnt"),1)
			ELSEIF objRS("SlmEnt") = "DD" or objRS("SlmEnt") = "B6" THEN
				objExcelBWomen.Fields(8).Value = NextRotSlmWom
				NextRotSlmWom = NextRotSlmWom + 1
			END IF
				
			IF left(objRS("TrkEnt"),1) = "B" and right(objRS("TrkEnt"),1) <> "6" THEN
				objExcelBWomen.Fields(9).Value = Right(objRS("TrkEnt"),1)
			ELSEIF objRS("TrkEnt") = "DD" or objRS("TrkEnt") = "B6" THEN
				objExcelBWomen.Fields(9).Value = NextRotTrkWom
				NextRotTrkWom = NextRotTrkWom + 1
			END IF
				
			IF left(objRS("JmpEnt"),1) = "B" and right(objRS("JmpEnt"),1) <> "6" THEN
				objExcelBWomen.Fields(10).Value = Right(objRS("JmpEnt"),1)
			ELSEIF objRS("JmpEnt") = "DD" or objRS("JmpEnt") = "B6" THEN
				objExcelBWomen.Fields(10).Value = NextRotJmpWom
				NextRotJmpWom = NextRotJmpWom + 1
			END IF

			IF left(objRS("OffCode"),1) = "C" THEN
				objExcelBWomen.Fields(11).Value = objRS("OffCode")
			ELSE
				objExcelBWomen.Fields(11).Value = objRS("OffRat")
			END IF

			objExcelBWomen.Update	

		
		ELSEIF objRS("Div") = "CM" or objRS("Sorter") = "E0FF" THEN

			'	*******	A Men Entries done here ...

			IF objRS("Sorter") <> "E0FF" THEN

				objExcelAMen.addnew
				objExcelAMen.Fields(0).Value = MemID
				objExcelAMen.Fields(1).Value = objRS("LastName")
				objExcelAMen.Fields(2).Value = objRS("FirstName")
				objExcelAMen.Fields(3).Value = trim(objRS("Team"))
				objExcelAMen.Fields(4).Value = objRS("Div")
				objExcelAMen.Fields(5).Value = objRS("Age")
				objExcelAMen.Fields(6).Value = objRS("City")
				objExcelAMen.Fields(7).Value = objRS("State")
				objExcelAMen.Fields(16).Value = objRS("SptsDiv")
				objExcelAMen.Fields(17).Value = OKtoSki
				objExcelAMen.Fields(18).Value = UpgrdAmt

				IF left(objRS("SlmEnt"),1) = "A" THEN
					IF right(objRS("SlmEnt"),1) = "0" THEN							
						objExcelAMen.Fields(8).Value = "10"
					ELSE
						objExcelAMen.Fields(8).Value = Right(objRS("SlmEnt"),1)
					END IF 
				END IF
				
				IF left(objRS("TrkEnt"),1) = "A" THEN
					IF right(objRS("TrkEnt"),1) = "0" THEN							
						objExcelAMen.Fields(9).Value = "10"
					ELSE
						objExcelAMen.Fields(9).Value = Right(objRS("TrkEnt"),1)
					END IF 
				END IF
				
				IF left(objRS("JmpEnt"),1) = "A" THEN
					IF right(objRS("JmpEnt"),1) = "0" THEN							
						objExcelAMen.Fields(10).Value = "10"
					ELSE
						objExcelAMen.Fields(10).Value = Right(objRS("JmpEnt"),1)
					END IF 
				END IF

				IF left(objRS("OffCode"),1) = "C" THEN
					objExcelAMen.Fields(11).Value = objRS("OffCode")
				ELSE
					objExcelAMen.Fields(11).Value = objRS("OffRat")
				END IF

				objExcelAMen.Update	

			END IF
	

			'	*******	B Men Entries done here ...

			objExcelBMen.addnew
			objExcelBMen.Fields(0).Value = MemID
			objExcelBMen.Fields(1).Value = objRS("LastName")
			objExcelBMen.Fields(2).Value = objRS("FirstName")
			IF objRS("Sorter") <> "E0FF" THEN
				objExcelBMen.Fields(3).Value = trim(objRS("Team"))
			ELSE
				objExcelBMen.Fields(3).Value = "OFF"
			END IF

			IF objRS("SlmEnt") = "DD" or objRS("TrkEnt") = "DD" or objRS("JmpEnt") = "DD" THEN
				objExcelBMen.Fields(4).Value = objRS("AgeDiv")
			ELSE
				objExcelBMen.Fields(4).Value = objRS("Div")
			END IF

			objExcelBMen.Fields(5).Value = objRS("Age")
			objExcelBMen.Fields(6).Value = objRS("City")
			objExcelBMen.Fields(7).Value = objRS("State")
			objExcelBMen.Fields(16).Value = objRS("SptsDiv")
			objExcelBMen.Fields(17).Value = OKtoSki
			objExcelBMen.Fields(18).Value = UpgrdAmt
			
			IF left(objRS("SlmEnt"),1) = "B" and right(objRS("SlmEnt"),1) <> "6" THEN
				objExcelBMen.Fields(8).Value = Right(objRS("SlmEnt"),1)
			ELSEIF objRS("SlmEnt") = "DD" or objRS("SlmEnt") = "B6" THEN
				objExcelBMen.Fields(8).Value = NextRotSlmMen
				NextRotSlmMen = NextRotSlmMen + 1
			END IF
				
			IF left(objRS("TrkEnt"),1) = "B" and right(objRS("TrkEnt"),1) <> "6" THEN
				objExcelBMen.Fields(9).Value = Right(objRS("TrkEnt"),1)
			ELSEIF objRS("TrkEnt") = "DD" or objRS("TrkEnt") = "B6" THEN
				objExcelBMen.Fields(9).Value = NextRotTrkMen
				NextRotTrkMen = NextRotTrkMen + 1
			END IF
				
			IF left(objRS("JmpEnt"),1) = "B" and right(objRS("JmpEnt"),1) <> "6" THEN
				objExcelBMen.Fields(10).Value = Right(objRS("JmpEnt"),1)
			ELSEIF objRS("JmpEnt") = "DD" or objRS("JmpEnt") = "B6" THEN
				objExcelBMen.Fields(10).Value = NextRotJmpMen
				NextRotJmpMen = NextRotJmpMen + 1
			END IF

			IF left(objRS("OffCode"),1) = "C" THEN
				objExcelBMen.Fields(11).Value = objRS("OffCode")
			ELSE
				objExcelBMen.Fields(11).Value = objRS("OffRat")
			END IF

			objExcelBMen.Update	

		END IF


	ELSE

		'	*******	All Others go in this section here ...

		objExcelAllOthr.addnew
		objExcelAllOthr.Fields(0).Value = MemID
		objExcelAllOthr.Fields(1).Value = objRS("LastName")
		objExcelAllOthr.Fields(2).Value = objRS("FirstName")
		objExcelAllOthr.Fields(3).Value = trim(objRS("Team"))
		objExcelAllOthr.Fields(4).Value = objRS("Div")
		objExcelAllOthr.Fields(5).Value = objRS("Age")
		objExcelAllOthr.Fields(6).Value = objRS("City")
		objExcelAllOthr.Fields(7).Value = objRS("State")
		objExcelAllOthr.Fields(11).Value = objRS("OffRat")
		objExcelAllOthr.Fields(16).Value = objRS("SptsDiv")
		objExcelAllOthr.Fields(17).Value = OKtoSki
		objExcelAllOthr.Fields(18).Value = UpgrdAmt

		objExcelAllOthr.Update	

	END IF
		
	objRS.MoveNext

LOOP


'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

'	Finally tack on Overall Registration Total Rides and Skiers

objExcelRegist.addnew
objExcelRegist.Fields(0).Value = " "
objExcelRegist.Update	
objExcelRegist.addnew
objExcelRegist.Fields(0).Value = "Grand Tots"
objExcelRegist.Fields(1).Value = "Across All Teams"
objExcelRegist.Fields(3).Value = "Tot"
objExcelRegist.Fields(8).Value = GrandSlm
objExcelRegist.Fields(9).Value = GrandTrk
objExcelRegist.Fields(10).Value = GrandJmp
objExcelRegist.Fields(11).Value = "  <-- Rides"
objExcelRegist.Fields(13).Value = "       Skiers -->"
objExcelRegist.Fields(15).Value = GrandTot
objExcelRegist.Update	

objExcelAMen.close
set objExcelAMen = nothing
objExcelBMen.close
set objExcelBMen = nothing
objExcelAWomen.close
set objExcelAWomen = nothing
objExcelBWomen.close
set objExcelBWomen = nothing
objExcelRegist.close
set objExcelRegist = nothing
objExcelAllOthr.close
set objExcelAllOthr = nothing
objExcelConn.close
set objExcelConn = nothing
'
objRS.Close
Set objRS = Nothing
objTL.Close
Set objTL = Nothing

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
	filenamewithscores = filenamewithscores & "-" & strTSanction & ".xls"
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
         Registration Export Excel Workbook is now complete and ready to download.&nbsp;</font>
         <br>&nbsp;<br>

         <font color="#FF0000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>
         !! New for 2011 -- Online Team Entry and Rotation Plan details now included !!
         </strong></font><br>
         <font size="2" face="Verdana, Arial, Helvetica, sans-serif">Details of Team Entry
         and Rotation Plans that have been prepared and submitted by the respective team
         captains through the new online team entry system, are now incorporated into this 
         Excel workbook.&nbsp; See the revised instructions section of your Excel workbook 
         for details.&nbsp; <font color="#FF0000"><strong>New for Fall 2011 -- A Registrar 
         Recap is now included.</strong></font>&nbsp; This new section in the Excel workbook 
         makes it easier for Registrars to see each team's overall entry status, to see which 
         entered skiers still need to execute event waivers locally, and to assess each team's 
         total entry fees.</font>
         <br>&nbsp;<br>

         <a href="excel/<% response.write filenamewithscores %>"><font face="Arial" size="2"><b>RIGHT 
         Click Here</b></font></a>&nbsp; <font size="2" face="Verdana, Arial, Helvetica, 
         sans-serif">to download your NCWSA Registration Template workbook, then select the 
         "Save As" option from that menu, and then choose a suitable location to store the 
         download file in your PC.&nbsp; After your Registration Template download has 
         completed, then open the Excel file from that location on your PC.&nbsp; It will 
         open automatically to an Instructions Tab section.&nbsp; Please review the material 
         in that section for the latest information on contents and usage.</font>
         <br>&nbsp;<br>

         <% IF AllowAccess THEN %>

         <font size="2" face="Verdana, Arial, Helvetica, sans-serif">If you are now doing 
         your <b><i>final and official</i></b> download of entries for this tournament, then 
         <b><i>after</i></b> downloading your Excel workbook (see the <b>RIGHT Click Here</b> 
         link in paragraph above), then click the <b>Close Registration</b> button that you
         see below.&nbsp; That will block any further modifications to existing Team Entry and 
         Rotation Plans, and refer team captains to the Tournament Registrar at the tournament 
         site for any last-minute changes.</font>
         <br>&nbsp;<br>
         
         <% ELSE %>

         <font size="2" face="Verdana, Arial, Helvetica, sans-serif">Online Entry to this 
         tournament is currently set to <b>Closed</b>.&nbsp; If you have not yet done your 
         <b><i>final and official</i></b> download of entries for this tournament, then 
         you may want to re-open Online Entry status, by clicking the <b>Re-Open Registration</b> 
         button below.</font>
         <br>&nbsp;<br>
         
         <% END IF %>

         </td>
      </tr>

 	</table>

	<TABLE ALIGN="CENTER" WIDTH=80%>
		
		<TR>

    <% IF AllowAccess THEN %>

		    <TD width=35% align=center>
			<form action="NCWSAChgRegStat.asp?TourID=<%=left(strTSanction,6)%>&Status=Close" method="post">
			<input type="submit" style="width:12em" value="Close Registration"
			title="Close Online Registration -- No further Changes by Captains allowed"></form>
 		   	</TD>

    <% ELSE %>

		    <TD width=35% align=center>
			<form action="NCWSAChgRegStat.asp?TourID=<%=left(strTSanction,6)%>&Status=Open" method="post">
			<input type="submit" style="width:12em" value="Re-Open Registration"
			title="Close Online Registration -- No further Changes by Captains allowed"></form>
 		   	</TD>

    <% END IF %>

	    <TD width=30% align=center>
		<form action="LookupMembers.asp?FormStatus=newsearch" method="post">
		<input type="submit" style="width:10em" value="Lookup Members"></form>
    	</TD>

	    <td width=25% align=center>     				
		<form action="Index.asp" method="post">
    <input type="submit" style="width:7em" value="Quit"></form>
 	    </td>
  	    
 	  </TR>

 	</TABLE>

  	  </td>
	  </tr>
</table>
</body>
</html>






                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               e                                    p                                                             i                                                                                   r                                                                             