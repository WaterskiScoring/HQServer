<% 

If not Session("aauth") then response.redirect "Login.asp"

Server.ScriptTimeout = 300

Dim curTraceMsg, sTourID, sTourYear, sTourDate, sStateSQL, sTourName, sStateList, sUserName

'	-----------------------------------------------------------------------
' Validate TourID value for scores to be Exported.
'	-----------------------------------------------------------------------
sTourDate = ""
sStateSQL = "State IN ('')"
sStateList = ""
sTourName = ""
curTraceMsg = "<br />Process Begin"

sUserName = session("UserName")
sTourID = Session("TournamentID")
IF len(sTourID) > 0 THEN
    sTourID = Session("TournamentID")
    sTourDate = session("tournamentdate")
    sStateSQL = Session("StateSQL")
    sStateList = Session("StateList")
    sTourName = sTourName
ELSE 
    sTourID = Request.QueryString("TourID")
END IF	

sTourYear = 2000 + left(sTourID,2)
curTraceMsg = curTraceMsg & "<br />TourId=" & sTourID & ", sTourYear=" & sTourYear & ", sTourDate=" & sTourDate & ", sStateList=" & sStateList & ", sStateSQL=" & sStateSQL

'	-----------------------------------------------------------------------
'	Utility function defintion
'	-----------------------------------------------------------------------
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
' The following lines of HTML display the "opening please wait" banner.
'	-----------------------------------------------------------------------
%>
    
<html>
    <head>
        <title>USA Water Ski Registration Template Using OLR</title>
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

'	-----------------------------------------------------------------------
'Open connection to Sanction Database
'Get tournament attributes from TSchedul table
'	-----------------------------------------------------------------------
Set WaterskiConnect = Server.CreateObject("ADODB.Connection")
WaterskiConnect.Open Application("WaterSkiConn")
Dim rsWaterski
Set rsWaterski = Server.CreateObject("ADODB.RecordSet")
rsWaterski.ActiveConnection = WaterskiConnect

Dim curSqlStmt, strTStatus, strTSanction, strTourName, strTourDate
curSqlStmt = "Select Top 1 TSanction, TStatus, TournAppID, TDateE, TName, TCity, TState from Sanctions.dbo.TSchedul where TournAppID = '" & left(sTourID,6) & "'"
rsWaterski.Open curSqlStmt
If rsWaterski.EOF THEN
	strTStatus = -1
    strTSanction = sTourID
    strTourName = sTourId
ELSE 
	strTStatus = rsWaterski("TStatus")
    strTSanction = rsWaterski("TSanction")
    strTourDate = rsWaterski("TDateE")
    sTourDate = strTourDate
    strTourName = rsWaterski("TName")
    sTourName = strTourName

	IF left(strTSanction,6) <> left(sTourID,6) THEN
		strTSanction = sTourID
	END IF
END IF

rsWaterski.Close
Set rsWaterski = Nothing
WaterskiConnect.Close

'	-----------------------------------------------------------------------
'	Read applicable Membership Pricing Info from HQ Table into local Array
'	-----------------------------------------------------------------------
curTraceMsg = curTraceMsg & "<br /><br />Membership Types with pricing"
Dim MT, MemPrice(200), MemUpgrd(200)
FOR MT = 1 to 200: MemPrice(MT) = 0: MemUpgrd(MT) = 0: NEXT

'Open connection to HQ Database
Set HQConnect = CreateObject("ADODB.Connection")
HQConnect.Open Application("HQSQLConn")

curSqlStmt = "SELECT * FROM [Membership Types with pricing]" 
curSqlStmt = curSqlStmt & " WHERE EffectiveFrom <= CONVERT(DATETIME, '" & sTourDate & " 00:00:00', 102)"
curSqlStmt = curSqlStmt & " AND EffectiveTo >= CONVERT(DATETIME, '" & sTourDate & " 00:00:00', 102)"
Set rsMemType = HQConnect.Execute(curSqlStmt)
DO UNTIL rsMemType.EOF
	MT = rsMemType("Membership Type Code")
	MemPrice(MT) = rsMemType("MemberShipTypeRates")
	MemUpgrd(MT) = rsMemType("CostToUpgrade")

	rsMemType.MoveNext
LOOP

rsMemType.Close
Set rsMemType = Nothing
HQConnect.Close

'	-----------------------------------------------------------------------
'Open database connection 
'Check to determine if there are any qualification entries
'Then check to determine if there are qualifications for this tournamnet
'	-----------------------------------------------------------------------
Set WaterskiConnect = Server.CreateObject("ADODB.Connection")
WaterskiConnect.Open Application("WaterSkiConn")
Set rsWaterski = Server.CreateObject("ADODB.RecordSet")
rsWaterski.ActiveConnection = WaterskiConnect

Dim QfyNum, DateRaw, DateFmt, I1, I2, RowNo
DateRaw = Date(): I1 = instr(DateRaw,"/"): I2 = instr(I1+1,DateRaw,"/")
DateFmt = Mid(DateRaw,I2+1): ' Start with Year value
IF I1=2 THEN DateFmt = DateFmt + "-0" + Left(DateRaw,1): ELSE DateFmt = DateFmt + "-" + Left(DateRaw,2)
IF I2-I1=2 THEN DateFmt = DateFmt + "-0" + Mid(DateRaw,I1+1,1): ELSE DateFmt = DateFmt + "-" + Mid(DateRaw,I1+1,2)

curSqlStmt = "Select count(*) as QfyNum From Cobra00025.USAWSRank.RegisterQualify_TEST Where left(TourID,6) = '" & left(sTourID,6) & "';"
rsWaterski.Open curSqlStmt
QfyNum = rsWaterski("QfyNum")

rsWaterski.Close
Set rsWaterski = Nothing
WaterskiConnect.Close

'	-----------------------------------------------------------------------
'Now open a connection to the new XLS file
'Setup to reference blank registration template file
'	-----------------------------------------------------------------------
Dim fileRegXls
Set fileRegXls = Server.CreateObject("Scripting.FileSystemObject")
Dim pathExcelFiles
pathExcelFiles = Server.MapPath("Excel/")
dim copyFileSour, copyFileDest
curTraceMsg = curTraceMsg & "<br /><br />pathExcelFiles=" & pathExcelFiles

copyFileSour = pathExcelFiles & "/Templates/NonPreRegTemplateBlank.xls"
copyFileDest = pathExcelFiles & "/template.xls"
curTraceMsg = curTraceMsg & "<br />copyFileSour=" & copyFileSour & "<br />copyFileDest=" & copyFileDest

fileRegXls.CopyFile copyFileSour, copyFileDest , True

'	-----------------------------------------------------------------------
'Now open a connection to the new XLS file
'	-----------------------------------------------------------------------
Set objExcelConn = Server.CreateObject("ADODB.Connection")
objExcelConn.Provider = "Microsoft.ACE.OLEDB.12.0"
objExcelConn.ConnectionString = "Data Source=" & copyFileDest & ";Extended Properties=""Excel 8.0;"""
    On Error Resume Next
objExcelConn.Open
    If Err.Number <> 0 Then
        %>
            <DIV ID="debugMsg">
                <br />Error creating registration template file
                <br />Err.Number=<%=Err.Number %>
                <br />Err.Description=<%=Err.Description %>
                <br />
            </DIV>
        <%
        On Error Goto 0 ' But don't let other errors hide!
    End If
curTraceMsg = curTraceMsg & "<br />Open Excel file=" & copyFileDest

Set objExcelSingleFields = Server.CreateObject("ADODB.Recordset")
objExcelSingleFields.ActiveConnection = objExcelConn 
objExcelSingleFields.CursorType = 3                    'Static cursor.
objExcelSingleFields.LockType = 2                      'Pessimistic Lock.

objExcelSingleFields.Source = "Select * from PreRegTournamentName"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = sTourName
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
objExcelSingleFields.Fields(0).Value = sTourName
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
objExcelSingleFields.Fields(0).Value = sTourName
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

curTraceMsg = curTraceMsg & "<br />Create PreReg sheet"

Set objExcelActive = Server.CreateObject("ADODB.Recordset")
objExcelActive.ActiveConnection = objExcelConn 
objExcelActive.CursorType = 3                    'Static cursor.
objExcelActive.LockType = 2                      'Pessimistic Lock.
objExcelActive.Source = "Select * from ActiveRange"
objExcelActive.Open

curTraceMsg = curTraceMsg & "<br />Create Active sheet"

Set objExcelInActive = Server.CreateObject("ADODB.Recordset")
objExcelInActive.ActiveConnection = objExcelConn 
objExcelInActive.CursorType = 3                    'Static cursor.
objExcelInActive.LockType = 2                      'Pessimistic Lock.
objExcelInActive.Source = "Select * from InActiveRange"
objExcelInActive.Open

curTraceMsg = curTraceMsg & "<br />Create InActive sheet"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Next we insert Chief and Appointed official Person ID's for the 
''' desired Tournament, from the Sanctions.Registration table into 
''' a work table, along with Applicable Chief Codes.  But first we
''' need to do a delete of any existing rows for that TournAppID.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Set WaterskiConnect = Server.CreateObject("ADODB.Connection")
WaterskiConnect.Open Application("WaterSkiConn")
Set rsWaterski = Server.CreateObject("ADODB.RecordSet")
rsWaterski.ActiveConnection = WaterskiConnect

curSqlStmt = "Delete from USAWaterski.dbo.TempApptdOfcls where TournAppID = '" & left(sTourID,6) & "' OR DateAdd(Day,30,WhenAdded) < GetDate()"
WaterskiConnect.Execute (curSqlStmt)

curSqlStmt = "Insert into USAWaterski.dbo.TempApptdOfcls (PersonID, TournAppID, OffCode, WhenAdded) "
curSqlStmt = curSqlStmt & "SELECT PersonID, '" & left(sTourID,6) & "', Max(OffCode), GetDate() "
curSqlStmt = curSqlStmt & "FROM ( "
curSqlStmt = curSqlStmt & "    SELECT Cast(CASE When len(CJudgePID)<9 THEN CJudgePID ELSE right(CJudgePID,8) END as integer) AS PersonID, 'CJ' AS OffCode "
curSqlStmt = curSqlStmt & "    FROM sanctions.dbo.registration WHERE TournAppID = '" & left(sTourID,6) & "' and isnumeric(CJudgePID) = 1 "
curSqlStmt = curSqlStmt & "    UNION"
curSqlStmt = curSqlStmt & "    SELECT Cast(CASE When len(CDriverPID)<9 Then CDriverPID Else right(CDriverPID,8) END as integer) AS PersonID, 'CD' AS OffCode"
curSqlStmt = curSqlStmt & "    FROM sanctions.dbo.registration WHERE TournAppID = '" & left(sTourID,6) & "' and isnumeric(CDriverPID) = 1 "
curSqlStmt = curSqlStmt & "    UNION"
curSqlStmt = curSqlStmt & "    SELECT Cast(CASE When len(CScorePID)<9 Then CScorePID Else right(CScorePID,8) END as integer) AS PersonID, 'CC' AS OffCode "
curSqlStmt = curSqlStmt & "    FROM sanctions.dbo.registration WHERE TournAppID = '" & left(sTourID,6) & "' and isnumeric(CScorePID) = 1 "
curSqlStmt = curSqlStmt & "    UNION"
curSqlStmt = curSqlStmt & "    SELECT Cast(CASE When len(CSafPID)<9 Then CSafPID Else right(CSafPID,8) END as integer) AS PersonID, 'CS' AS OffCode"
curSqlStmt = curSqlStmt & "    FROM sanctions.dbo.registration WHERE TournAppID = '" & left(sTourID,6) & "' and isnumeric(CSafPID) = 1 "
curSqlStmt = curSqlStmt & "    UNION"
curSqlStmt = curSqlStmt & "    SELECT Cast(CASE When len(TechCPID)<9 Then TechCPID Else right(TechCPID,8) END as integer) AS PersonID, 'CT' AS OffCode"
curSqlStmt = curSqlStmt & "    FROM sanctions.dbo.registration WHERE TournAppID = '" & left(sTourID,6) & "' and isnumeric(TechCPID) = 1 "
curSqlStmt = curSqlStmt & "    UNION"
curSqlStmt = curSqlStmt & "    SELECT Cast(CASE When len(Ap1JPID)<9 Then Ap1JPID Else right(Ap1JPID,8) END as integer) AS PersonID, 'APTJ' AS OffCode"
curSqlStmt = curSqlStmt & "    FROM sanctions.dbo.registration WHERE TournAppID = '" & left(sTourID,6) & "' and isnumeric(Ap1JPID) = 1 "
curSqlStmt = curSqlStmt & "    UNION"
curSqlStmt = curSqlStmt & "    SELECT Cast(CASE When len(Ap2JPID)<9 Then Ap2JPID Else right(Ap2JPID,8) END as integer) AS PersonID, 'APTJ' AS OffCode"
curSqlStmt = curSqlStmt & "    FROM sanctions.dbo.registration WHERE TournAppID = '" & left(sTourID,6) & "' and isnumeric(Ap2JPID) = 1 "
curSqlStmt = curSqlStmt & "    UNION"
curSqlStmt = curSqlStmt & "    SELECT Cast(CASE when len(Ap3JPID)<9 Then Ap3JPID Else right(Ap3JPID,8) END as integer) AS PersonID, 'APTJ' AS OffCode"
curSqlStmt = curSqlStmt & "    FROM sanctions.dbo.registration WHERE TournAppID = '" & left(sTourID,6) & "' and isnumeric(Ap3JPID) = 1 "
curSqlStmt = curSqlStmt & "    UNION"
curSqlStmt = curSqlStmt & "    SELECT Cast(CASE when len(Ap4JPID)<9 Then Ap4JPID Else right(Ap4JPID,8) END as integer) AS PersonID, 'APTJ' AS OffCode"
curSqlStmt = curSqlStmt & "    FROM sanctions.dbo.registration WHERE TournAppID = '" & left(sTourID,6) & "' and isnumeric(Ap4JPID) = 1 "
curSqlStmt = curSqlStmt & "    UNION"
curSqlStmt = curSqlStmt & "    SELECT Cast(CASE when len(Ap5JPID)<9 Then Ap5JPID Else right(Ap5JPID,8) END as integer) AS PersonID, 'APTJ' AS OffCode"
curSqlStmt = curSqlStmt & "    FROM sanctions.dbo.registration WHERE TournAppID = '" & left(sTourID,6) & "' and isnumeric(Ap5JPID) = 1 "
curSqlStmt = curSqlStmt & "    UNION"
curSqlStmt = curSqlStmt & "    SELECT Cast(CASE when len(Ap1SPID)<9 Then Ap1SPID Else right(Ap1SPID,8) END as integer) AS PersonID, 'APTS' AS OffCode"
curSqlStmt = curSqlStmt & "    FROM sanctions.dbo.registration WHERE TournAppID = '" & left(sTourID,6) & "' and isnumeric(Ap1SPID) = 1 "
curSqlStmt = curSqlStmt & "    UNION"
curSqlStmt = curSqlStmt & "    SELECT Cast(CASE when len(Ap2SPID)<9 Then Ap2SPID Else right(Ap2SPID,8) END as integer) AS PersonID, 'APTS' AS OffCode"
curSqlStmt = curSqlStmt & "    FROM sanctions.dbo.registration WHERE TournAppID = '" & left(sTourID,6) & "' and isnumeric(Ap2SPID) = 1 "
curSqlStmt = curSqlStmt & "    UNION"
curSqlStmt = curSqlStmt & "    SELECT Cast(CASE when len(Ap3SPID)<9 Then Ap3SPID Else right(Ap3SPID,8) END as integer) AS PersonID, 'APTS' AS OffCode"
curSqlStmt = curSqlStmt & "    FROM sanctions.dbo.registration WHERE TournAppID = '" & left(sTourID,6) & "' and isnumeric(Ap3SPID) = 1 "
curSqlStmt = curSqlStmt & "    UNION"
curSqlStmt = curSqlStmt & "    SELECT Cast(CASE when len(Ap1DrPID)<9 Then Ap1DrPID Else right(Ap1DrPID,8) END as integer) AS PersonID, 'APTD' AS OffCode"
curSqlStmt = curSqlStmt & "    FROM sanctions.dbo.registration WHERE TournAppID = '" & left(sTourID,6) & "' and isnumeric(Ap1DrPID) = 1 "
curSqlStmt = curSqlStmt & "    UNION"
curSqlStmt = curSqlStmt & "    SELECT Cast(CASE when len(PanAmPID)<9 Then PanAmPID Else right(PanAmPID,8) END as integer) AS PersonID, 'APTJ' AS OffCode"
curSqlStmt = curSqlStmt & "    FROM sanctions.dbo.registration WHERE TournAppID = '" & left(sTourID,6) & "' and isnumeric(PanAmPID) = 1)"
curSqlStmt = curSqlStmt & " SOX Group by PersonID"
WaterskiConnect.Execute (curSqlStmt)

curTraceMsg = curTraceMsg & "<br />Updated chief official"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Now build a Query to Extract the Desired Members, joining in data 
' from the Rankings and Officials and Membership Type tables.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim curSqlStmt1, curSqlStmt2, curSqlStmt3, curSqlStmt4, curSqlStmt5

'Member Number and name
curSqlStmt1 = "SELECT Substring(MX.MemberID,1,3) + '-' + Substring(MX.MemberID,4,2) + '-' + Substring(MX.MemberID,6,4) as MemID"
curSqlStmt1 = curSqlStmt1 & ", MX.LastName, MX.FirstName"

'Skier division
curSqlStmt1 = curSqlStmt1 & ", Coalesce(RD.Div, Case when MX.Age <= 17 and MX.Sex = 'F' Then 'G'"
curSqlStmt1 = curSqlStmt1 & " when MX.Age <= 17 then 'B' when MX.Sex = 'F' then 'W' else 'M' end + Case"
curSqlStmt1 = curSqlStmt1 & " when MX.Age <= 9 then '1' when MX.Age <= 13 then '2' when MX.Age <= 17 then '3'"
curSqlStmt1 = curSqlStmt1 & " when MX.Age <= 24 then '1' when MX.Age <= 34 then '2' when MX.Age <= 44 then '3'"
curSqlStmt1 = curSqlStmt1 & " when MX.Age <= 52 then '4' when MX.Age <= 59 then '5' when MX.Age <= 64 then '6'"
curSqlStmt1 = curSqlStmt1 & " when MX.Age <= 69 then '7' when MX.Age <= 74 then '8' when MX.Age <= 79 then '9'"
curSqlStmt1 = curSqlStmt1 & " when MX.Age <= 84 then 'A' else 'B' end) as Div"
		
'Skier information
curSqlStmt1 = curSqlStmt1 & ", MX.Age, MX.City, MX.State, MX.Waiver"

'Skier official ratings
curSqlStmt1 = curSqlStmt1 & ", Case when OD.PersonID is Null then '-' else Right(OD.RtgLvl,1) end +"
curSqlStmt1 = curSqlStmt1 & " Case when OJ.PersonID is Null then '-' else Right(OJ.RtgLvl,1) end +"
curSqlStmt1 = curSqlStmt1 & " Case when OC.PersonID is Null then '-' else Right(OC.RtgLvl,1) end +"
curSqlStmt1 = curSqlStmt1 & " Case when OS.PersonID is Null then '-' else Right(OS.RtgLvl,1) end as OffRat"

curSqlStmt1 = curSqlStmt1 & ", Coalesce(SO.OffCode,'') as OffCode"

'Event Ratings
curSqlStmt1 = curSqlStmt1 & ", Coalesce(SX.SlmSco,'') as SlmSco"
curSqlStmt1 = curSqlStmt1 & ", Coalesce(TX.TrkSco,'') as TrkSco"
curSqlStmt1 = curSqlStmt1 & ", Coalesce(JX.JmpSco,'') as JmpSco"

curSqlStmt1 = curSqlStmt1 & ", Coalesce(SE.SlmEli,SX.SlmRat,'') as SlmRat"
curSqlStmt1 = curSqlStmt1 & ", Coalesce(TE.TrkEli,TX.TrkRat,'') as TrkRat"
curSqlStmt1 = curSqlStmt1 & ", Coalesce(JE.JmpEli,JX.JmpRat,'') as JmpRat"
curSqlStmt1 = curSqlStmt1 & ", Coalesce(OE.OvrEli,OX.OvrRat,'') as OvrRat"

'Other member stuff
curSqlStmt1 = curSqlStmt1 & " ,MX.EffTo, MX.Memtype, MX.MemCode, MX.CanSki, MX.CanSkiGR "
		
'	-----------------------------------------------------------------------
'FROM Statement
'	-----------------------------------------------------------------------
curSqlStmt2 = ""
curSqlStmt2 = curSqlStmt2 & "FROM ("

'	-----------------------------------------------------------------------
'Use select as a data source for member data
'	-----------------------------------------------------------------------
curSqlStmt2 = curSqlStmt2 & "    SELECT MT.PersonIDWithCheckDigit as MemberID, MT.PersonID"
curSqlStmt2 = curSqlStmt2 & "        , Left(MT.LastName,12) as LastName, Left(MT.FirstName,10) as FirstName"
curSqlStmt2 = curSqlStmt2 & "        , " & sTourYear & " - Year(MT.BirthDate) - 1 as Age"
curSqlStmt2 = curSqlStmt2 & "        , Upper(Left(MT.Sex,1)) as Sex, MT.WaiverStatusID as Waiver"
curSqlStmt2 = curSqlStmt2 & "        , Left(MT.City,12) as City, Left(MT.State,2) as State"
curSqlStmt2 = curSqlStmt2 & "        , MT.EffectiveTo as EffTo, MT.MembershipTypeCode as MemType"
curSqlStmt2 = curSqlStmt2 & "        , Typ.TypeCode as MemCode, Typ.CanSkiInTournaments as CanSki"
curSqlStmt2 = curSqlStmt2 & "        , Typ.CanSkiInGRTournaments as CanSkiGR"
curSqlStmt2 = curSqlStmt2 & "    FROM USAWaterski.dbo.Members as MT "
curSqlStmt2 = curSqlStmt2 & "      INNER JOIN USAWaterski.dbo.MembershipTypes as Typ ON MT.MembershipTypeCode = Typ.MemberShipTypeID "
curSqlStmt2 = curSqlStmt2 & "    WHERE Typ.ExporttoTouramentRegistrationTemplate = 1"
curSqlStmt2 = curSqlStmt2 & "      AND DateAdd(mm,18,MT.EffectiveTo) > GetDate()"
curSqlStmt2 = curSqlStmt2 & "      AND MT.Deceased = 0 "
curSqlStmt2 = curSqlStmt2 & "      AND (" & sStateSQL
curSqlStmt2 = curSqlStmt2 & "           OR PersonID in (Select PersonID from USAWaterski.dbo.TempApptdOfcls WHERE TournAppID = '" & left(sTourID,6) & "' )"
curSqlStmt2 = curSqlStmt2 & "           )"
curSqlStmt2 = curSqlStmt2 & "    ) as MX"

'	-----------------------------------------------------------------------
' Use select as a data source for officials information 
' (DLA: retrieving type 3 officials info but not sure what type 3 is but think it is for drivers ????)
'	-----------------------------------------------------------------------
curSqlStmt3 = ""
curSqlStmt3 = curSqlStmt3 & "  LEFT JOIN ("
curSqlStmt3 = curSqlStmt3 & "      SELECT OT.PersonID, Max(convert(char(1), LV.LevelOrderforTemplate) + LV.LevelAbbreviationforTemplate) AS RtgLvl"
curSqlStmt3 = curSqlStmt3 & "      FROM USAWaterski.dbo.Officials OT"
curSqlStmt3 = curSqlStmt3 & "        INNER JOIN USAWaterski.dbo.Level LV ON OT.Level_ID = LV.Level_ID"
curSqlStmt3 = curSqlStmt3 & "      WHERE OT.DivisionCode in ('AWS','USA')"
curSqlStmt3 = curSqlStmt3 & "        AND LV.LevelOrderforTemplate IS NOT NULL"
curSqlStmt3 = curSqlStmt3 & "        AND OT.RatingType_ID = 3"
curSqlStmt3 = curSqlStmt3 & "      GROUP BY OT.PersonID)"
curSqlStmt3 = curSqlStmt3 & "      AS OD"
curSqlStmt3 = curSqlStmt3 & "      ON OD.PersonID = MX.PersonID"

'	-----------------------------------------------------------------------
' Use select as a data source for officials information 
' (DLA: retrieving type 1 officials info but not sure what type 1 is but think it is for judges????)
'	-----------------------------------------------------------------------
curSqlStmt3 = curSqlStmt3 & "  LEFT JOIN ("
curSqlStmt3 = curSqlStmt3 & "      SELECT OT.PersonID, Max(convert(char(1), LV.LevelOrderforTemplate) + LV.LevelAbbreviationforTemplate) AS RtgLvl"
curSqlStmt3 = curSqlStmt3 & "      FROM USAWaterski.dbo.Officials OT"
curSqlStmt3 = curSqlStmt3 & "        INNER JOIN USAWaterski.dbo.Level LV ON OT.Level_ID = LV.Level_ID"
curSqlStmt3 = curSqlStmt3 & "      WHERE OT.DivisionCode in ('AWS','USA')"
curSqlStmt3 = curSqlStmt3 & "        AND LV.LevelOrderforTemplate IS NOT NULL"
curSqlStmt3 = curSqlStmt3 & "        AND OT.RatingType_ID = 1"
curSqlStmt3 = curSqlStmt3 & "      GROUP BY OT.PersonID)"
curSqlStmt3 = curSqlStmt3 & "      AS OJ"
curSqlStmt3 = curSqlStmt3 & "      ON OJ.PersonID = MX.PersonID"

'	-----------------------------------------------------------------------
' Use select as a data source for officials information 
' (DLA: retrieving type 2 officials info but not sure what type 2 is but think it is for scorers????)
'	-----------------------------------------------------------------------
curSqlStmt3 = curSqlStmt3 & "  LEFT JOIN ("
curSqlStmt3 = curSqlStmt3 & "      SELECT OT.PersonID, Max(convert(char(1), LV.LevelOrderforTemplate) + LV.LevelAbbreviationforTemplate) AS RtgLvl"
curSqlStmt3 = curSqlStmt3 & "      FROM USAWaterski.dbo.Officials OT"
curSqlStmt3 = curSqlStmt3 & "        INNER JOIN USAWaterski.dbo.Level LV ON OT.Level_ID = LV.Level_ID"
curSqlStmt3 = curSqlStmt3 & "      WHERE OT.DivisionCode in ('AWS','USA')"
curSqlStmt3 = curSqlStmt3 & "        AND LV.LevelOrderforTemplate IS NOT NULL"
curSqlStmt3 = curSqlStmt3 & "        AND OT.RatingType_ID = 2"
curSqlStmt3 = curSqlStmt3 & "      GROUP BY OT.PersonID)"
curSqlStmt3 = curSqlStmt3 & "      AS OC"
curSqlStmt3 = curSqlStmt3 & "      ON OC.PersonID = MX.PersonID"

'	-----------------------------------------------------------------------
' Use select as a data source for officials information 
' (DLA: retrieving type 9 officials info but not sure what type 9 is but think it is for safety????)
'	-----------------------------------------------------------------------
curSqlStmt3 = curSqlStmt3 & "  LEFT JOIN ("
curSqlStmt3 = curSqlStmt3 & "      SELECT OT.PersonID, Max(convert(char(1), LV.LevelOrderforTemplate) + LV.LevelAbbreviationforTemplate) AS RtgLvl"
curSqlStmt3 = curSqlStmt3 & "      FROM USAWaterski.dbo.Officials OT"
curSqlStmt3 = curSqlStmt3 & "        INNER JOIN USAWaterski.dbo.Level LV ON OT.Level_ID = LV.Level_ID"
curSqlStmt3 = curSqlStmt3 & "      WHERE OT.DivisionCode in ('AWS','USA')"
curSqlStmt3 = curSqlStmt3 & "        AND LV.LevelOrderforTemplate IS NOT NULL"
curSqlStmt3 = curSqlStmt3 & "        AND OT.RatingType_ID = 9"
curSqlStmt3 = curSqlStmt3 & "      GROUP BY OT.PersonID)"
curSqlStmt3 = curSqlStmt3 & "      AS OS"
curSqlStmt3 = curSqlStmt3 & "      ON OS.PersonID = MX.PersonID"

'	-----------------------------------------------------------------------
' Use select as a data source for chief officials
'	-----------------------------------------------------------------------
curSqlStmt3 = curSqlStmt3 & "  LEFT JOIN ("
curSqlStmt3 = curSqlStmt3 & "      SELECT PersonID, OffCode"
curSqlStmt3 = curSqlStmt3 & "      FROM USAWaterski.dbo.TempApptdOfcls"
curSqlStmt3 = curSqlStmt3 & "      WHERE TournAppID = '" & left(sTourID,6) & "')"
curSqlStmt3 = curSqlStmt3 & "      AS SO"
curSqlStmt3 = curSqlStmt3 & "      ON SO.PersonID = MX.PersonID"

'	-----------------------------------------------------------------------
' Use select as a data source for skier ranking data
' The RD subquery below UNIONS selects from Rankings PLUS RegisterEvents, to
' ensure that EVERY entered skier will show up SOMEWHERE in the template.
'	-----------------------------------------------------------------------
curSqlStmt4 = ""
curSqlStmt4 = curSqlStmt4 & "  LEFT JOIN ("
curSqlStmt4 = curSqlStmt4 & "      SELECT MemberID, Div"
curSqlStmt4 = curSqlStmt4 & "      FROM Cobra00025.USAWSRank.Rankings"
curSqlStmt4 = curSqlStmt4 & "      WHERE SkiYearID = 1 and RankScore is not Null"
curSqlStmt4 = curSqlStmt4 & "        AND Left(Div,1) in ('B','G','M','W','O')"
curSqlStmt4 = curSqlStmt4 & "      GROUP BY MemberID, Div)"
curSqlStmt4 = curSqlStmt4 & "      AS RD"
curSqlStmt4 = curSqlStmt4 & "      ON RD.MemberID = MX.MemberID"

'	-----------------------------------------------------------------------
' Slalom ratings
'	-----------------------------------------------------------------------
curSqlStmt4 = curSqlStmt4 & "  LEFT JOIN ("
curSqlStmt4 = curSqlStmt4 & "      SELECT MemberID, Div, Reg_Ski, AWSA_Rat as SlmRat"
curSqlStmt4 = curSqlStmt4 & "             , Left(Cast(Cast(RankScore as Decimal(7,2)) as Varchar(8)),6) as SlmSco"
curSqlStmt4 = curSqlStmt4 & "      FROM Cobra00025.USAWSRank.Rankings"
curSqlStmt4 = curSqlStmt4 & "      WHERE SkiYearID = 1"
curSqlStmt4 = curSqlStmt4 & "        AND Left(Div,1) in ('B','G','M','W','O')"
curSqlStmt4 = curSqlStmt4 & "        AND Event = 'S'"
curSqlStmt4 = curSqlStmt4 & "        AND RankScore is not null)"
curSqlStmt4 = curSqlStmt4 & "      AS SX"
curSqlStmt4 = curSqlStmt4 & "      ON RD.MemberID = SX.MemberID AND RD.Div = SX.Div"

'	-----------------------------------------------------------------------
' Trick ratings
'	-----------------------------------------------------------------------
curSqlStmt4 = curSqlStmt4 & "  LEFT JOIN ("
curSqlStmt4 = curSqlStmt4 & "      SELECT MemberID, Div, Reg_Ski, AWSA_Rat as TrkRat"
curSqlStmt4 = curSqlStmt4 & "             , Left(Cast(Cast(RankScore as Decimal(7,1)) as Varchar(8)),6) as TrkSco"
curSqlStmt4 = curSqlStmt4 & "      FROM Cobra00025.USAWSRank.Rankings"
curSqlStmt4 = curSqlStmt4 & "      WHERE SkiYearID = 1"
curSqlStmt4 = curSqlStmt4 & "        AND Left(Div,1) in ('B','G','M','W','O')"
curSqlStmt4 = curSqlStmt4 & "        AND Event = 'T'"
curSqlStmt4 = curSqlStmt4 & "        AND RankScore is not null)"
curSqlStmt4 = curSqlStmt4 & "      AS TX"
curSqlStmt4 = curSqlStmt4 & "      ON RD.MemberID = TX.MemberID AND RD.Div = TX.Div"

'	-----------------------------------------------------------------------
' Jump ratings
'	-----------------------------------------------------------------------
curSqlStmt4 = curSqlStmt4 & "  LEFT JOIN ("
curSqlStmt4 = curSqlStmt4 & "      SELECT MemberID, Div, Reg_Ski, AWSA_Rat as JmpRat"
curSqlStmt4 = curSqlStmt4 & "             , Left(Cast(Cast(RankScore as Decimal(6,2)) as Varchar(8)),6) as JmpSco"
curSqlStmt4 = curSqlStmt4 & "      FROM Cobra00025.USAWSRank.Rankings"
curSqlStmt4 = curSqlStmt4 & "      WHERE SkiYearID = 1"
curSqlStmt4 = curSqlStmt4 & "        AND Left(Div,1) in ('B','G','M','W','O')"
curSqlStmt4 = curSqlStmt4 & "        AND Event = 'J'"
curSqlStmt4 = curSqlStmt4 & "        AND RankScore is not null)"
curSqlStmt4 = curSqlStmt4 & "      AS JX"
curSqlStmt4 = curSqlStmt4 & "     ON RD.MemberID = JX.MemberID AND RD.Div = JX.Div"

'	-----------------------------------------------------------------------
' Overall ratings
'	-----------------------------------------------------------------------
curSqlStmt4 = curSqlStmt4 & "  LEFT JOIN ("
curSqlStmt4 = curSqlStmt4 & "      SELECT MemberID, Div,  AWSA_Rat as OvrRat"
curSqlStmt4 = curSqlStmt4 & "             , Left(Cast(Cast(RankScore as Decimal(7,1)) as Varchar(8)),6) as OvrSco"
curSqlStmt4 = curSqlStmt4 & "      FROM Cobra00025.USAWSRank.Rankings"
curSqlStmt4 = curSqlStmt4 & "      WHERE SkiYearID = 1"
curSqlStmt4 = curSqlStmt4 & "        AND Left(Div,1) in ('B','G','M','W','O')"
curSqlStmt4 = curSqlStmt4 & "        AND Event = 'O'"
curSqlStmt4 = curSqlStmt4 & "        AND RankScore is not null)"
curSqlStmt4 = curSqlStmt4 & "      AS OX"
curSqlStmt4 = curSqlStmt4 & "      ON RD.MemberID = OX.MemberID AND RD.Div = OX.Div"

'	-----------------------------------------------------------------------
' Slalom something to with elite dates 
'	-----------------------------------------------------------------------
curSqlStmt5 = ""
curSqlStmt5 = curSqlStmt5 & "  LEFT JOIN ("
curSqlStmt5 = curSqlStmt5 & "      SELECT MemberID, max(DivElite) as SlmEli"
curSqlStmt5 = curSqlStmt5 & "      FROM Cobra00025.USAWSRank.EliteDates"
curSqlStmt5 = curSqlStmt5 & "      WHERE SkiYearID = 1"
curSqlStmt5 = curSqlStmt5 & "        AND Event = 'S'"
curSqlStmt5 = curSqlStmt5 & "        AND QualThru >= '" & sTourDate & "'"
curSqlStmt5 = curSqlStmt5 & "      GROUP BY MemberID)"
curSqlStmt5 = curSqlStmt5 & "      AS SE"
curSqlStmt5 = curSqlStmt5 & "      ON RD.MemberID = SE.MemberID"

'	-----------------------------------------------------------------------
' Trick something to with elite dates 
'	-----------------------------------------------------------------------
curSqlStmt5 = curSqlStmt5 & "  LEFT JOIN ("
curSqlStmt5 = curSqlStmt5 & "      SELECT MemberID, max(DivElite) as TrkEli"
curSqlStmt5 = curSqlStmt5 & "      FROM Cobra00025.USAWSRank.EliteDates"
curSqlStmt5 = curSqlStmt5 & "      WHERE SkiYearID = 1"
curSqlStmt5 = curSqlStmt5 & "        AND Event = 'T'"
curSqlStmt5 = curSqlStmt5 & "        AND QualThru >= '" & sTourDate & "'"
curSqlStmt5 = curSqlStmt5 & "      GROUP BY MemberID)"
curSqlStmt5 = curSqlStmt5 & "      AS TE"
curSqlStmt5 = curSqlStmt5 & "      ON RD.MemberID = TE.MemberID"

'	-----------------------------------------------------------------------
' Jump something to with elite dates 
'	-----------------------------------------------------------------------
curSqlStmt5 = curSqlStmt5 & "  LEFT JOIN ("
curSqlStmt5 = curSqlStmt5 & "      SELECT MemberID, max(DivElite) as JmpEli"
curSqlStmt5 = curSqlStmt5 & "      FROM Cobra00025.USAWSRank.EliteDates"
curSqlStmt5 = curSqlStmt5 & "      WHERE SkiYearID = 1"
curSqlStmt5 = curSqlStmt5 & "        AND Event = 'J'"
curSqlStmt5 = curSqlStmt5 & "        AND QualThru >= '" & sTourDate & "'"
curSqlStmt5 = curSqlStmt5 & "      GROUP BY MemberID)"
curSqlStmt5 = curSqlStmt5 & "      AS JE"
curSqlStmt5 = curSqlStmt5 & "      ON RD.MemberID = JE.MemberID"

'	-----------------------------------------------------------------------
' Overall something to with elite dates 
'	-----------------------------------------------------------------------
curSqlStmt5 = curSqlStmt5 & "  LEFT JOIN ("
curSqlStmt5 = curSqlStmt5 & "      SELECT MemberID, max(DivElite) as OvrEli"
curSqlStmt5 = curSqlStmt5 & "      FROM Cobra00025.USAWSRank.EliteDates"
curSqlStmt5 = curSqlStmt5 & "      WHERE SkiYearID = 1"
curSqlStmt5 = curSqlStmt5 & "        AND Event = 'O'"
curSqlStmt5 = curSqlStmt5 & "        AND QualThru >= '" & sTourDate & "'"
curSqlStmt5 = curSqlStmt5 & "      GROUP BY MemberID)"
curSqlStmt5 = curSqlStmt5 & "      AS OE"
curSqlStmt5 = curSqlStmt5 & "      ON RD.MemberID = OE.MemberID"

'	-----------------------------------------------------------------------
' Order by statement
'	-----------------------------------------------------------------------
curSqlStmt5 = curSqlStmt5 & " Order By MX.LastName, MX.FirstName, RD.MemberID, RD.Div"

'	-----------------------------------------------------------------------
' Execute SQL statement to retrieve skier information and load to registration template
'	-----------------------------------------------------------------------
curSqlStmt = curSqlStmt1 & curSqlStmt2 & curSqlStmt3 & curSqlStmt4 & curSqlStmt5
curTraceMsg = curTraceMsg & "<br /><br />Skiers retrieved<br />"

Set WaterskiConnect = Server.CreateObject("ADODB.Connection")
WaterskiConnect.Open Application("WaterSkiConn")

Dim Counter0, Counter1, Counter2, Counter3
Dim rsMember
Set rsMember = Server.CreateObject("ADODB.RecordSet")
rsMember.ActiveConnection = WaterskiConnect

    On Error Resume Next
rsMember.Open curSqlStmt
    If Err.Number <> 0 Then
        %>
            <DIV ID="debugMsg">
                <br />Error opening SQL to retrieve skier list
                <br />Err.Number=<%=Err.Number %>
                <br />Err.Description=<%=Err.Description %>
                <br />
            </DIV>
        <%
        On Error Goto 0 ' But don't let other errors hide!
    End If
Do until rsMember.EOF
	Counter0 = Counter0 + 1

	IF len(rsMember("OffCode")) > 0 THEN

		Counter1 = Counter1 + 1
		objExcelPreReg.addnew
		objExcelPreReg.Fields(0).Value = rsMember("MemID")
		objExcelPreReg.Fields(1).Value = rsMember("LastName")
		objExcelPreReg.Fields(2).Value = rsMember("FirstName")
		
		objExcelPreReg.Fields(4).Value = rsMember("Div")
		objExcelPreReg.Fields(5).Value = rsMember("Age")
		objExcelPreReg.Fields(6).Value = rsMember("City")
		objExcelPreReg.Fields(7).Value = rsMember("State")
	
		IF left(rsMember("OffCode"),1) = "C" THEN
			objExcelPreReg.Fields(11).Value = rsMember("OffCode")
		ELSE
			objExcelPreReg.Fields(11).Value = rsMember("OffRat")
		END IF

		objExcelPreReg.Fields(12).Value = rsMember("SlmSco")
		objExcelPreReg.Fields(13).Value = rsMember("TrkSco")
		objExcelPreReg.Fields(14).Value = rsMember("JmpSco")
		objExcelPreReg.Fields(15).Value = rsMember("SlmRat")
		objExcelPreReg.Fields(16).Value = rsMember("TrkRat")
		objExcelPreReg.Fields(17).Value = rsMember("JmpRat")
		objExcelPreReg.Fields(18).Value = rsMember("OvrRat")

		IF rsMember("EffTo") >= cdate(sTourDate) and rsMember("CanSki") = True and rsMember("Waiver") > 0 THEN	
			objExcelPreReg.Fields(26).Value = "Yes"
			objExcelPreReg.Fields(27).Value = "Pre-Regist"
		ELSE
			objExcelPreReg.Fields(26).Value = "    No"
			' Figure applicable Renewal / Upgrade Amount based on MemType & Status
			MT = rsMember("MemType")
			IF MT < 1 OR MT > 200 THEN MT = 1
			IF rsMember("EffTo") < cdate(sTourDate) THEN 
				IF rsMember("CanSki") = False THEN
					objExcelPreReg.Fields(27).Value = "Nds Rnw/Upg" 
					objExcelPreReg.Fields(28).Value = FormatNumber(MemPrice(MT)+MemUpgrd(MT),2)
				ELSE
					objExcelPreReg.Fields(27).Value = "Needs Renew" 
					objExcelPreReg.Fields(28).Value = FormatNumber(MemPrice(MT),2)
				END IF
			ELSE 
				IF rsMember("CanSkiGR") = True THEN
					objExcelPreReg.Fields(27).Value = "** G/R Only" 
					objExcelPreReg.Fields(28).Value = FormatNumber(MemUpgrd(MT),2)
				ELSEIF rsMember("CanSki") = False THEN
					objExcelPreReg.Fields(27).Value = "Needs Upgrd" 
					objExcelPreReg.Fields(28).Value = FormatNumber(MemUpgrd(MT),2)
				ELSE
					objExcelPreReg.Fields(27).Value = "Nds Ann Wvr" 
					objExcelPreReg.Fields(28).Value = FormatNumber(0,2)
				END IF				
			END IF
		END IF	

		objExcelPreReg.Update

	ELSEIF rsMember("EffTo") >= cdate(sTourDate) and rsMember("CanSki") = True and rsMember("Waiver") > 0 THEN
		Counter2 = Counter2 + 1
		objExcelActive.addnew
		objExcelActive.Fields(0).Value = rsMember("MemID")
		objExcelActive.Fields(1).Value = rsMember("LastName")
		objExcelActive.Fields(2).Value = rsMember("FirstName")

		IF Mid(sTourID,4,3) = "999" THEN
			objExcelActive.Fields(3).Value = rsMember("Reg_Ski")
		END IF				

		objExcelActive.Fields(4).Value = rsMember("Div")
		objExcelActive.Fields(5).Value = rsMember("Age")
		objExcelActive.Fields(6).Value = rsMember("City")
		objExcelActive.Fields(7).Value = rsMember("State")
		
		objExcelActive.Fields(11).Value = rsMember("OffRat")
		objExcelActive.Fields(12).Value = rsMember("SlmSco")
		objExcelActive.Fields(13).Value = rsMember("TrkSco")
		objExcelActive.Fields(14).Value = rsMember("JmpSco")
		objExcelActive.Fields(15).Value = rsMember("SlmRat")
		objExcelActive.Fields(16).Value = rsMember("TrkRat")
		objExcelActive.Fields(17).Value = rsMember("JmpRat")
		objExcelActive.Fields(18).Value = rsMember("OvrRat")
		
	    objExcelActive.Fields(26).Value = "Yes"
		objExcelActive.Update

	ELSE
		Counter3 = Counter3 + 1
		objExcelInActive.addnew
		objExcelInActive.Fields(0).Value = rsMember("MemID")
		objExcelInActive.Fields(1).Value = rsMember("LastName")
		objExcelInActive.Fields(2).Value = rsMember("FirstName")

		IF Mid(sTourID,4,3) = "999" THEN
			objExcelInActive.Fields(3).Value = rsMember("Reg_Ski")
		END IF				

		objExcelInActive.Fields(4).Value = rsMember("Div")
		objExcelInActive.Fields(5).Value = rsMember("Age")
		objExcelInActive.Fields(6).Value = rsMember("City")
		objExcelInActive.Fields(7).Value = rsMember("State")
		
		'added 4-11-2007 MOK
		objExcelInActive.Fields(11).Value = rsMember("OffRat")
		objExcelInActive.Fields(12).Value = rsMember("SlmSco")
		objExcelInActive.Fields(13).Value = rsMember("TrkSco")
		objExcelInActive.Fields(14).Value = rsMember("JmpSco")
		objExcelInActive.Fields(15).Value = rsMember("SlmRat")
		objExcelInActive.Fields(16).Value = rsMember("TrkRat")
		objExcelInActive.Fields(17).Value = rsMember("JmpRat")
		objExcelInActive.Fields(18).Value = rsMember("OvrRat")

		objExcelInActive.Fields(26).Value = "    No"

		' Figure applicable Renewal / Upgrade Amount based on MemType & Status

		MT = rsMember("MemType")
		IF MT < 1 OR MT > 200 THEN MT = 1

		IF rsMember("EffTo") < cdate(sTourDate) THEN 
			IF rsMember("CanSki") = False THEN
				objExcelInActive.Fields(27).Value = "Nds Rnw/Upg" 
				objExcelInActive.Fields(28).Value = FormatNumber(MemPrice(MT)+MemUpgrd(MT),2)
			ELSE
				objExcelInActive.Fields(27).Value = "Needs Renew" 
				objExcelInActive.Fields(28).Value = FormatNumber(MemPrice(MT),2)
			END IF
		ELSE 
			IF rsMember("CanSkiGR") = True THEN
				objExcelInActive.Fields(27).Value = "** G/R Only" 
				objExcelInActive.Fields(28).Value = FormatNumber(MemUpgrd(MT),2)
			ELSEIF rsMember("CanSki") = False THEN
				objExcelInActive.Fields(27).Value = "Needs Upgrd" 
				objExcelInActive.Fields(28).Value = FormatNumber(MemUpgrd(MT),2)
			ELSE
				objExcelInActive.Fields(27).Value = "Nds Ann Wvr" 
				objExcelInActive.Fields(28).Value = FormatNumber(0,2)
			END IF				
		END IF
		
		objExcelInActive.Update

	END IF
	
	rsMember.MoveNext
Loop

curTraceMsg = curTraceMsg & "<br />Retrieve skier list complete: Counter0=" & Counter0 & ", Counter1=" & Counter1  & ", Counter2=" & Counter2

'	-----------------------------------------------------------------------
'	-----------------------------------------------------------------------
objExcelActive.close
set objExcelActive = nothing
objExcelInActive.close
set objExcelInActive = nothing
objExcelConn.close
set objExcelConn = nothing

rsMember.Close
Set rsMember = Nothing

'	-----------------------------------------------------------------------
'Now copy the file from Template to a file with the tournamentid
'	-----------------------------------------------------------------------
Dim regTemplateFilename
'"06M123-Entries-SSSSSS-YYYYMMDD", 
regTemplateFilename = "Entries-" & sStateList & "-" & DateFmt

'	-----------------------------------------------------------------------
'Add the Tournament Name to the start of the file name
'	-----------------------------------------------------------------------
if len(sTourName) > 0 then
	regTemplateFilename = RemoveInvalidChars(sTourName) & "-" & regTemplateFilename
end if

'	-----------------------------------------------------------------------
'Append the username
'	-----------------------------------------------------------------------
if len(sUserName) > 0 then
	regTemplateFilename = regTemplateFilename & "-" & strTSanction & ".xls"
else
	regTemplateFilename = regTemplateFilename & ".xls"
end if

fileRegXls.CopyFile copyFileDest, pathExcelFiles & "/" & regTemplateFilename , True

'	-----------------------------------------------------------------------
' Clean up old files
'	-----------------------------------------------------------------------
'Clean up old files
Set dataFolder = objFSO.GetFolder(pathExcelFiles)  
Set folderFileList = dataFolder.Files 
Response.Write "<br>"
For Each curFile in folderFileList
	Set myfile = objFSO.GetFile(pathExcelFiles & "/" & curFile.name)
	if datediff("d",myfile.DateCreated,date()) > 2 and left(myfile.name,8) <> "Template" then
		myfile.delete
	end if
Next  

Set dataFolder = nothing
Set folderFileList = nothing
Set objFSO = Nothing
Set fileRegXls = Nothing

'	-----------------------------------------------------------------------
'	-----------------------------------------------------------------------
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

    </body>

</html>

<%
' This final bit of HTML is written after processing is successfully completed
' to tell the user how to download their template, and where to go from here.
      
Response.Flush
%>

<html>

<head>
<title>Create Registration Export</title>

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
      	Registration Support for -- <%=sTourName%></font></p>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>  
  
<table border="0" cellspacing="0" cellpadding="0">  
  <tr> 
    <td width="185" valign="top" bgcolor="#42639F">

	<font face="Verdana" size="2" COLOR="#FFFFFF"><br>&nbsp;Currently Logged in as: </font><br>
	<font face="Verdana" size="2" COLOR="#FFFFFF">&nbsp;<%=sUserName%>&nbsp;&nbsp;
		<%=sTourDate%></font><br>
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
         <td><a href="excel/<% response.write regTemplateFilename %>"><font face="Arial" size="2"><b>RIGHT 
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

<%
curTraceMsg = curTraceMsg & "<br /><br />Process Complete"
''''% >
''''    <DIV style="width: 100%; Text-Align:Left; margin-left: 0; margin-right: auto; FONT-SIZE:1.0em; FONT-WEIGHT:normal;">
''''        <br /><br />< %=curTraceMsg % ><br />
''''    </DIV>
''''% >

curTraceMsg = ""

%>
</body>
</html>
                                                                                                                       