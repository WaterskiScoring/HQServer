<!--#include virtual="/epl/functions.asp" -->

<% 

If not Session("aauth") then response.redirect "Login.asp"

Server.ScriptTimeout = 300

' The following lines of HTML display the "opening please wait" banner.

%>
    
<html><head><title>USA Water Ski Registration Template</title>
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

Dim objRS
Set objRS = Server.CreateObject("ADODB.RecordSet")
objRS.ActiveConnection = objConn

'Get TStatus and TSanction from TSchedul table
Dim strTStatus, strTSanction
sSQL = "Select Top 1 TSanction, TStatus from Sanctions.dbo.TSchedul where TournAppID = '"
sSQL = sSQL & left(Session("TournamentID"),6) & "'"
objRS.Open sSQL
If objRS.EOF THEN
	strTStatus = -1: strTSanction = Session("TournamentID")
ELSE 
	strTStatus = objRS("TStatus"): strTSanction = objRS("TSanction")
	IF left(strTSanction,6) <> left(Session("TournamentID"),6) THEN
		strTSanction = Session("TournamentID")
	END IF
END IF
objRS.Close


'""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'"""""""""""""" With Scores and Ratings """""""""""""""""""""""
'""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""


objFSO.CopyFile path & "/Templates/NonPreRegTemplateBlank.xls", path & "/template_with_scores.xls" , True

'Now open a connection to the new XLS file

Set objExcelConn = Server.CreateObject("ADODB.Connection")
objExcelConn.Open "ExcelDSNwithScores"

Set objExcelSingleFields = Server.CreateObject("ADODB.Recordset")
objExcelSingleFields.ActiveConnection = objExcelConn 
objExcelSingleFields.CursorType = 3                    'Static cursor.
objExcelSingleFields.LockType = 2                      'Pessimistic Lock.

objExcelSingleFields.Source = "Select * from PreRegTournamentName"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = session("TournamentName")
objExcelSingleFields.update
objExcelSingleFields.close
		
objExcelSingleFields.Source = "Select * from PreRegTournamentID"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = strTSanction	'this is the same as the tournament ID
objExcelSingleFields.update
objExcelSingleFields.close
		
objExcelSingleFields.Source = "Select * from PreRegAsOfRange"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = "AS OF " & DateFmt
objExcelSingleFields.update
objExcelSingleFields.close
		
objExcelSingleFields.Source = "Select * from ActiveTournamentName"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = session("TournamentName")
objExcelSingleFields.update
objExcelSingleFields.close
		
objExcelSingleFields.Source = "Select * from ActiveTournamentID"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = strTSanction	'this is the same as the tournament ID
objExcelSingleFields.update
objExcelSingleFields.close
		
objExcelSingleFields.Source = "Select * from ActiveAsOfRange"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = "AS OF " & DateFmt
objExcelSingleFields.update
objExcelSingleFields.close
		
objExcelSingleFields.Source = "Select * from InActiveTournamentName"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = session("TournamentName")
objExcelSingleFields.update
objExcelSingleFields.close
		
objExcelSingleFields.Source = "Select * from InActiveTournamentID"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = strTSanction
objExcelSingleFields.update
objExcelSingleFields.close
		
objExcelSingleFields.Source = "Select * from InActiveAsOfDate"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = "AS OF " & DateFmt
objExcelSingleFields.update
objExcelSingleFields.close
		
Set objExcelPreReg = Server.CreateObject("ADODB.Recordset")
objExcelPreReg.ActiveConnection = objExcelConn 
objExcelPreReg.CursorType = 3                    'Static cursor.
objExcelPreReg.LockType = 2                      'Pessimistic Lock.
objExcelPreReg.Source = "Select * from PreRegRange"
objExcelPreReg.Open

Set objExcelActive = Server.CreateObject("ADODB.Recordset")
objExcelActive.ActiveConnection = objExcelConn 
objExcelActive.CursorType = 3                    'Static cursor.
objExcelActive.LockType = 2                      'Pessimistic Lock.
objExcelActive.Source = "Select * from ActiveRange"
objExcelActive.Open

Set objExcelInActive = Server.CreateObject("ADODB.Recordset")
objExcelInActive.ActiveConnection = objExcelConn 
objExcelInActive.CursorType = 3                    'Static cursor.
objExcelInActive.LockType = 2                      'Pessimistic Lock.
objExcelInActive.Source = "Select * from InActiveRange"
objExcelInActive.Open



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Next we insert Chief and Appointed official Person ID's for the 
''' desired Tournament, from the Sanctions.Registration table into 
''' a work table, along with Applicable Chief Codes.  But first we
''' need to do a delete of any existing rows for that TournAppID.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim sSQL, sSQL1, sSQL2, sSQL3


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
''' Now build a Query to Extract the Desired Members, joining in data 
''' from the Rankings and Officials and Membership Type tables.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

sSQL = "Select MX.PersonID, MX.LastName, MX.FirstName,"

sSQL = sSQL & " Coalesce(RD.Div, Case when MX.Age <= 17 and MX.Sex = 'F' Then 'G'"
sSQL = sSQL & " when MX.Age <= 17 then 'B' when MX.Sex = 'F' then 'W' else 'M' end + Case"
sSQL = sSQL & " when MX.Age <= 9 then '1' when MX.Age <= 13 then '2' when MX.Age <= 17 then '3'"
sSQL = sSQL & " when MX.Age <= 24 then '1' when MX.Age <= 34 then '2' when MX.Age <= 44 then '3'"
sSQL = sSQL & " when MX.Age <= 52 then '4' when MX.Age <= 59 then '5' when MX.Age <= 64 then '6'"
sSQL = sSQL & " when MX.Age <= 69 then '7' when MX.Age <= 74 then '8' when MX.Age <= 79 then '9'"
sSQL = sSQL & " when MX.Age <= 84 then 'A' else 'B' end) as Div,"
		
sSQL = sSQL & " MX.Age, MX.City, MX.State, MX.Waiver,"

sSQL = sSQL & " Case when OD.PersonID is Null then '-' else Right(OD.RtgLvl,1) end +"
sSQL = sSQL & " Case when OJ.PersonID is Null then '-' else Right(OJ.RtgLvl,1) end +"
sSQL = sSQL & " Case when OC.PersonID is Null then '-' else Right(OC.RtgLvl,1) end +"
sSQL = sSQL & " Case when OS.PersonID is Null then '-' else Right(OS.RtgLvl,1) end as OffRat,"

sSQL = sSQL & " Coalesce(SO.OffCode,'') as OffCode,"

sSQL = sSQL & " Coalesce(SX.SlmSco,'') as SlmSco,"
sSQL = sSQL & " Coalesce(TX.TrkSco,'') as TrkSco,"
sSQL = sSQL & " Coalesce(JX.JmpSco,'') as JmpSco,"

sSQL = sSQL & " Coalesce(SE.SlmEli,SX.SlmRat,'') as SlmRat,"
sSQL = sSQL & " Coalesce(TE.TrkEli,TX.TrkRat,'') as TrkRat,"
sSQL = sSQL & " Coalesce(JE.JmpEli,JX.JmpRat,'') as JmpRat,"
sSQL = sSQL & " Coalesce(OE.OvrEli,OX.OvrRat,'') as OvrRat,"

sSQL = sSQL & " MX.EffTo, MX.Memtype, MX.MemCode, MX.CanSki, MX.CanSkiGR"
		
sSQL = sSQL & " From (Select MT.PersonID,"
sSQL = sSQL & " Left(MT.LastName,12) as LastName, Left(MT.FirstName,10) as FirstName, "
sSQL = sSQL & Session("TournamentYear") & "-Year(MT.BirthDate)-1 as Age,"
sSQL = sSQL & " Upper(Left(MT.Sex,1)) as Sex, MT.WaiverStatusID as Waiver,"
sSQL = sSQL & " Left(MT.City,12) as City, Left(MT.State,2) as State,"
sSQL = sSQL & " MT.EffectiveTo as EffTo, MT.MembershipTypeCode as MemType,"
sSQL = sSQL & " Typ.TypeCode as MemCode, Typ.CanSkiInTournaments as CanSki,"
sSQL = sSQL & " Typ.CanSkiInGRTournaments as CanSkiGR"
sSQL = sSQL & " from USAWaterski.dbo.MembersLive as MT Inner Join"
sSQL = sSQL & " USAWaterski.dbo.MembershipTypes as Typ"
sSQL = sSQL & " ON MT.MembershipTypeCode = Typ.MemberShipTypeID"

sSQL = sSQL & " Where Typ.ExporttoTouramentRegistrationTemplate = 1"
sSQL = sSQL & " AND DateAdd(mm,18,MT.EffectiveTo) > GetDate()"
sSQL = sSQL & " AND MT.Deceased = 0 AND (" & Session("StateSQL") & " OR PersonID"
sSQL = sSQL & " in (Select PersonID from USAWaterski.dbo.TempApptdOfcls"
sSQL = sSQL & " Where TournAppID = '" & left(Session("TournamentID"),6)
sSQL = sSQL & "') ) ) as MX"


sSQL1 = " Left Join	(Select OT.PersonID,"
sSQL1 = sSQL1 & " Max(convert(char(1),LV.LevelOrderforTemplate)"
sSQL1 = sSQL1 & " + LV.LevelAbbreviationforTemplate) AS RtgLvl"
sSQL1 = sSQL1 & " FROM USAWaterski.dbo.Officials OT INNER JOIN"
sSQL1 = sSQL1 & " USAWaterski.dbo.Level LV ON OT.Level_ID = LV.Level_ID"
sSQL1 = sSQL1 & " WHERE OT.DivisionCode in ('AWS','USA')"
sSQL1 = sSQL1 & " AND LV.LevelOrderforTemplate IS NOT NULL"
sSQL1 = sSQL1 & " AND OT.RatingType_ID = 3 GROUP BY OT.PersonID) as OD"
sSQL1 = sSQL1 & " on OD.PersonID = MX.PersonID"

sSQL1 = sSQL1 & " Left Join	(Select OT.PersonID,"
sSQL1 = sSQL1 & " Max(convert(char(1),LV.LevelOrderforTemplate)"
sSQL1 = sSQL1 & " + LV.LevelAbbreviationforTemplate) AS RtgLvl"
sSQL1 = sSQL1 & " FROM USAWaterski.dbo.Officials OT INNER JOIN"
sSQL1 = sSQL1 & " USAWaterski.dbo.Level LV ON OT.Level_ID = LV.Level_ID"
sSQL1 = sSQL1 & " WHERE OT.DivisionCode in ('AWS','USA')"
sSQL1 = sSQL1 & " AND LV.LevelOrderforTemplate IS NOT NULL"
sSQL1 = sSQL1 & " AND OT.RatingType_ID = 1 GROUP BY OT.PersonID) as OJ"
sSQL1 = sSQL1 & " on OJ.PersonID = MX.PersonID"

sSQL1 = sSQL1 & " Left Join	(Select OT.PersonID,"
sSQL1 = sSQL1 & " Max(convert(char(1),LV.LevelOrderforTemplate)"
sSQL1 = sSQL1 & " + LV.LevelAbbreviationforTemplate) AS RtgLvl"
sSQL1 = sSQL1 & " FROM USAWaterski.dbo.Officials OT INNER JOIN"
sSQL1 = sSQL1 & " USAWaterski.dbo.Level LV ON OT.Level_ID = LV.Level_ID"
sSQL1 = sSQL1 & " WHERE OT.DivisionCode in ('AWS','USA')"
sSQL1 = sSQL1 & " AND LV.LevelOrderforTemplate IS NOT NULL"
sSQL1 = sSQL1 & " AND OT.RatingType_ID = 2 GROUP BY OT.PersonID) as OC"
sSQL1 = sSQL1 & " on OC.PersonID = MX.PersonID"

sSQL1 = sSQL1 & " Left Join	(Select OT.PersonID,"
sSQL1 = sSQL1 & " Max(convert(char(1),LV.LevelOrderforTemplate)"
sSQL1 = sSQL1 & " + LV.LevelAbbreviationforTemplate) AS RtgLvl"
sSQL1 = sSQL1 & " FROM USAWaterski.dbo.Officials OT INNER JOIN"
sSQL1 = sSQL1 & " USAWaterski.dbo.Level LV ON OT.Level_ID = LV.Level_ID"
sSQL1 = sSQL1 & " WHERE OT.DivisionCode in ('AWS','USA')"
sSQL1 = sSQL1 & " AND LV.LevelOrderforTemplate IS NOT NULL"
sSQL1 = sSQL1 & " AND OT.RatingType_ID = 9 GROUP BY OT.PersonID) as OS"
sSQL1 = sSQL1 & " on OS.PersonID = MX.PersonID"

sSQL1 = sSQL1 & " Left Join	(Select PersonID, OffCode from USAWaterski.dbo.TempApptdOfcls"
sSQL1 = sSQL1 & " Where TournAppID = '" & left(Session("TournamentID"),6) & "')"
sSQL1 = sSQL1 & " as SO on SO.PersonID = MX.PersonID"

sSQL1 = sSQL1 & " Left Join	(Select MemberID, Div from Cobra00025.USAWSRank.Rankings"
sSQL1 = sSQL1 & " where SkiYearID = 1 and RankScore is not Null"
sSQL1 = sSQL1 & " and Left(Div,1) in ('B','G','M','W','O')"
sSQL1 = sSQL1 & " group by MemberID, Div) as RD on cast(right(RD.MemberID,8) as integer) = MX.PersonID"


sSQL2 = " Left Join	(Select MemberID, Div, Reg_Ski, AWSA_Rat as SlmRat,"
sSQL2 = sSQL2 & " Left(Cast(Cast(RankScore as Decimal(7,2)) as Varchar(8)),6) as SlmSco"
sSQL2 = sSQL2 & " From Cobra00025.USAWSRank.Rankings Where SkiYearID = 1"
sSQL2 = sSQL2 & " and Left(Div,1) in ('B','G','M','W','O')"
sSQL2 = sSQL2 & " and Event = 'S' and RankScore is not null) as SX"
sSQL2 = sSQL2 & " on RD.MemberID = SX.MemberID and RD.Div = SX.Div"

sSQL2 = sSQL2 & " Left Join	(Select MemberID, Div, Reg_Ski, AWSA_Rat as TrkRat,"
sSQL2 = sSQL2 & " Left(Cast(Cast(RankScore as Decimal(7,1)) as Varchar(8)),6) as TrkSco"
sSQL2 = sSQL2 & " From Cobra00025.USAWSRank.Rankings Where SkiYearID = 1"
sSQL2 = sSQL2 & " and Left(Div,1) in ('B','G','M','W','O')"
sSQL2 = sSQL2 & " and Event = 'T' and RankScore is not null) as TX"
sSQL2 = sSQL2 & " on RD.MemberID = TX.MemberID and RD.Div = TX.Div"

sSQL2 = sSQL2 & " Left Join	(Select MemberID, Div, Reg_Ski, AWSA_Rat as JmpRat,"
sSQL2 = sSQL2 & " Left(Cast(Cast(RankScore as Decimal(6,2)) as Varchar(8)),6) as JmpSco"
sSQL2 = sSQL2 & " From Cobra00025.USAWSRank.Rankings Where SkiYearID = 1"
sSQL2 = sSQL2 & " and Left(Div,1) in ('B','G','M','W','O')"
sSQL2 = sSQL2 & " and Event = 'J' and RankScore is not null) as JX"
sSQL2 = sSQL2 & " on RD.MemberID = JX.MemberID and RD.Div = JX.Div"

sSQL2 = sSQL2 & " Left Join	(Select MemberID, Div,  AWSA_Rat as OvrRat,"
sSQL2 = sSQL2 & " Left(Cast(Cast(RankScore as Decimal(7,1)) as Varchar(8)),6) as OvrSco"
sSQL2 = sSQL2 & " From Cobra00025.USAWSRank.Rankings Where SkiYearID = 1"
sSQL2 = sSQL2 & " and Left(Div,1) in ('B','G','M','W','O')"
sSQL2 = sSQL2 & " and Event = 'O' and RankScore is not null) as OX"
sSQL2 = sSQL2 & " on RD.MemberID = OX.MemberID and RD.Div = OX.Div"

sSQL2 = sSQL2 & " Left Join	(Select MemberID, max(DivElite) as SlmEli"
sSQL2 = sSQL2 & " From Cobra00025.USAWSRank.EliteDates Where SkiYearID = 1"
sSQL2 = sSQL2 & " and Event = 'S' and QualThru >= '" & session("TournamentDate")
sSQL2 = sSQL2 & "' Group by MemberID) as SE on RD.MemberID = SE.MemberID"

sSQL2 = sSQL2 & " Left Join	(Select MemberID, max(DivElite) as TrkEli"
sSQL2 = sSQL2 & " From Cobra00025.USAWSRank.EliteDates Where SkiYearID = 1"
sSQL2 = sSQL2 & " and Event = 'T' and QualThru >= '" & session("TournamentDate")
sSQL2 = sSQL2 & "' Group by MemberID) as TE on RD.MemberID = TE.MemberID"

sSQL2 = sSQL2 & " Left Join	(Select MemberID, max(DivElite) as JmpEli"
sSQL2 = sSQL2 & " From Cobra00025.USAWSRank.EliteDates Where SkiYearID = 1"
sSQL2 = sSQL2 & " and Event = 'J' and QualThru >= '" & session("TournamentDate")
sSQL2 = sSQL2 & "' Group by MemberID) as JE on RD.MemberID = JE.MemberID"

sSQL2 = sSQL2 & " Left Join	(Select MemberID, max(DivElite) as OvrEli"
sSQL2 = sSQL2 & " From Cobra00025.USAWSRank.EliteDates Where SkiYearID = 1"
sSQL2 = sSQL2 & " and Event = 'O' and QualThru >= '" & session("TournamentDate")
sSQL2 = sSQL2 & "' Group by MemberID) as OE on RD.MemberID = OE.MemberID"

sSQL2 = sSQL2 & " Order By MX.LastName, MX.FirstName, RD.MemberID, RD.Div"


objRS.Open sSQL & sSQL1 & sSQL2

Dim Counter0, Counter1, Counter2
Dim SDiv, TDiv, JDiv, SPaid, TPaid, JPaid

Do until objRS.EOF

	' First we synthesize MemID value, from incoming PersonID, using internal function.

	MemberID = PersonIDwChkDgt(objRS("PersonID"))
	MemID = left(MemberID,3) & "-" & Mid(MemberID,4,2) & "-" & right(MemberID,4)

	IF len(objRS("OffCode")) > 0 THEN

		Counter0 = Counter0 + 1
		objExcelPreReg.addnew
		objExcelPreReg.Fields(0).Value = MemID
		objExcelPreReg.Fields(1).Value = objRS("LastName")
		objExcelPreReg.Fields(2).Value = objRS("FirstName")
		
		objExcelPreReg.Fields(4).Value = objRS("Div")
		objExcelPreReg.Fields(5).Value = objRS("Age")
		objExcelPreReg.Fields(6).Value = objRS("City")
		objExcelPreReg.Fields(7).Value = objRS("State")
	
		IF left(objRS("OffCode"),1) = "C" THEN
			objExcelPreReg.Fields(11).Value = objRS("OffCode")
		ELSE
			objExcelPreReg.Fields(11).Value = objRS("OffRat")
		END IF

		objExcelPreReg.Fields(12).Value = objRS("SlmSco")
		objExcelPreReg.Fields(13).Value = objRS("TrkSco")
		objExcelPreReg.Fields(14).Value = objRS("JmpSco")
		objExcelPreReg.Fields(15).Value = objRS("SlmRat")
		objExcelPreReg.Fields(16).Value = objRS("TrkRat")
		objExcelPreReg.Fields(17).Value = objRS("JmpRat")
		objExcelPreReg.Fields(18).Value = objRS("OvrRat")

		IF objRS("EffTo") >= cdate(session("TournamentDate")) and objRS("CanSki") = True and objRS("Waiver") > 0 THEN	
			objExcelPreReg.Fields(26).Value = "Yes"
			objExcelPreReg.Fields(27).Value = "Pre-Regist"
		ELSE
			objExcelPreReg.Fields(26).Value = "    No"
			' Figure applicable Renewal / Upgrade Amount based on MemType & Status
			MT = objRS("MemType")
			IF MT < 1 OR MT > 200 THEN MT = 1
			IF objRS("EffTo") < cdate(session("TournamentDate")) THEN 
				IF objRS("CanSki") = False THEN
					objExcelPreReg.Fields(27).Value = "Nds Rnw/Upg" 
					objExcelPreReg.Fields(28).Value = FormatNumber(MemPrice(MT)+MemUpgrd(MT),2)
				ELSE
					objExcelPreReg.Fields(27).Value = "Needs Renew" 
					objExcelPreReg.Fields(28).Value = FormatNumber(MemPrice(MT),2)
				END IF
			ELSE 
				IF objRS("CanSkiGR") = True THEN
					objExcelPreReg.Fields(27).Value = "** G/R Only" 
					objExcelPreReg.Fields(28).Value = FormatNumber(MemUpgrd(MT),2)
				ELSEIF objRS("CanSki") = False THEN
					objExcelPreReg.Fields(27).Value = "Needs Upgrd" 
					objExcelPreReg.Fields(28).Value = FormatNumber(MemUpgrd(MT),2)
				ELSE
					objExcelPreReg.Fields(27).Value = "Nds Ann Wvr" 
					objExcelPreReg.Fields(28).Value = FormatNumber(0,2)
				END IF				
			END IF
		END IF	

		objExcelPreReg.Update

	ELSEIF objRS("EffTo") >= cdate(session("TournamentDate")) and objRS("CanSki") = True and objRS("Waiver") > 0 THEN
		Counter1 = Counter1 + 1
		objExcelActive.addnew
		objExcelActive.Fields(0).Value = MemID
		objExcelActive.Fields(1).Value = objRS("LastName")
		objExcelActive.Fields(2).Value = objRS("FirstName")

		IF Mid(Session("TournamentID"),4,3) = "999" THEN
			objExcelActive.Fields(3).Value = objRS("Reg_Ski")
		END IF				

		objExcelActive.Fields(4).Value = objRS("Div")
		objExcelActive.Fields(5).Value = objRS("Age")
		objExcelActive.Fields(6).Value = objRS("City")
		objExcelActive.Fields(7).Value = objRS("State")
		
		objExcelActive.Fields(11).Value = objRS("OffRat")
		objExcelActive.Fields(12).Value = objRS("SlmSco")
		objExcelActive.Fields(13).Value = objRS("TrkSco")
		objExcelActive.Fields(14).Value = objRS("JmpSco")
		objExcelActive.Fields(15).Value = objRS("SlmRat")
		objExcelActive.Fields(16).Value = objRS("TrkRat")
		objExcelActive.Fields(17).Value = objRS("JmpRat")
		objExcelActive.Fields(18).Value = objRS("OvrRat")
		
	    objExcelActive.Fields(26).Value = "Yes"
		objExcelActive.Update

	ELSE
		Counter2 = Counter2 + 1
		objExcelInActive.addnew
		objExcelInActive.Fields(0).Value = MemID
		objExcelInActive.Fields(1).Value = objRS("LastName")
		objExcelInActive.Fields(2).Value = objRS("FirstName")

		IF Mid(Session("TournamentID"),4,3) = "999" THEN
			objExcelInActive.Fields(3).Value = objRS("Reg_Ski")
		END IF				

		objExcelInActive.Fields(4).Value = objRS("Div")
		objExcelInActive.Fields(5).Value = objRS("Age")
		objExcelInActive.Fields(6).Value = objRS("City")
		objExcelInActive.Fields(7).Value = objRS("State")
		
		'added 4-11-2007 MOK
		objExcelInActive.Fields(11).Value = objRS("OffRat")
		objExcelInActive.Fields(12).Value = objRS("SlmSco")
		objExcelInActive.Fields(13).Value = objRS("TrkSco")
		objExcelInActive.Fields(14).Value = objRS("JmpSco")
		objExcelInActive.Fields(15).Value = objRS("SlmRat")
		objExcelInActive.Fields(16).Value = objRS("TrkRat")
		objExcelInActive.Fields(17).Value = objRS("JmpRat")
		objExcelInActive.Fields(18).Value = objRS("OvrRat")

		objExcelInActive.Fields(26).Value = "    No"

		' Figure applicable Renewal / Upgrade Amount based on MemType & Status

		MT = objRS("MemType")
		IF MT < 1 OR MT > 200 THEN MT = 1

		IF objRS("EffTo") < cdate(session("TournamentDate")) THEN 
			IF objRS("CanSki") = False THEN
				objExcelInActive.Fields(27).Value = "Nds Rnw/Upg" 
				objExcelInActive.Fields(28).Value = FormatNumber(MemPrice(MT)+MemUpgrd(MT),2)
			ELSE
				objExcelInActive.Fields(27).Value = "Needs Renew" 
				objExcelInActive.Fields(28).Value = FormatNumber(MemPrice(MT),2)
			END IF
		ELSE 
			IF objRS("CanSkiGR") = True THEN
				objExcelInActive.Fields(27).Value = "** G/R Only" 
				objExcelInActive.Fields(28).Value = FormatNumber(MemUpgrd(MT),2)
			ELSEIF objRS("CanSki") = False THEN
				objExcelInActive.Fields(27).Value = "Needs Upgrd" 
				objExcelInActive.Fields(28).Value = FormatNumber(MemUpgrd(MT),2)
			ELSE
				objExcelInActive.Fields(27).Value = "Nds Ann Wvr" 
				objExcelInActive.Fields(28).Value = FormatNumber(0,2)
			END IF				
		END IF
		
		objExcelInActive.Update

	END IF
	
	objRS.MoveNext
Loop


'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""


objExcelActive.close
set objExcelActive = nothing
objExcelInActive.close
set objExcelInActive = nothing
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
      	USA Water Ski Pre-Registration Export</font></p>
      <p align="center"><font face="Verdana" size="4" color="#FFFFFF">
      	Registration Support for -- <%=session("TournamentName")%></font></p>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>  
  
<table border="0" cellspacing="0" cellpadding="0">  
  <tr> 
    <td width="185" valign="top" bgcolor="#42639F">

	<font face="Verdana" size="2" COLOR="#FFFFFF"><br>&nbsp;Currently Logged in as: </font><br>
	<font face="Verdana" size="2" COLOR="#FFFFFF">&nbsp;<%=Session("UserName")%>&nbsp;&nbsp;
		<%=session("TournamentDate")%></font><br>
	<br>
	
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
         <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><br>Your Pre-Registration 
         Export workbook is now complete and ready to download.</font></td>
      </tr>

      <tr> 
         <td>&nbsp;</td>
      </tr>

      <tr> 
         <td>&nbsp;</td>
         <td><font color="#FF0000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>
         !! New Structure to this Template in 2010 !!</strong>&nbsp;
         </font><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
         We have now created a separate "Participants" section in this Excel
         workbook for your to build your participant list in, and you will
         find that section will have initially been populated with the Chief
         and other Appointed Officials listed in the Sanction system for your
         Tournament.&nbsp; Further details on how to use this new framework
         can be found in the Instructions section of your download.
         </font></td>
      </tr>

      <tr> 
         <td>&nbsp;</td>
      </tr>

      <tr> 
         <td>&nbsp;</td>
         <td><a href="excel/<% response.write filenamewithscores %>"><font face="Arial" size="2"><b>RIGHT 
         Click Here</b></font></a>&nbsp; <font size="2" face="Verdana, Arial, Helvetica, sans-serif">to 
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
         <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
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






a                                                                                                                       