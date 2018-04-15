<!--#include virtual="/admin/MemberRegFunctions.asp"-->
<%
If not Session("aauth") then response.redirect "Login.asp"

Server.ScriptTimeout = 120000

Dim curTraceMsg, sTourID, sTourDate, sStateSQL, sTourName, sStateList, sUserName
Dim curSanctionId, curMemberId, curMemberFirstName, curMemberLastName

'	-----------------------------------------------------------------------
' Validate TourID value for scores to be Exported.
'	-----------------------------------------------------------------------
sTourDate = ""
sStateSQL = "State IN ('')"
sStateList = ""
sTourName = ""

sUserName = session("UserName")
sTourID = Session("TournamentID")
IF len(sTourID) > 0 THEN
    sTourID = Session("TournamentID")
    sTourDate = session("tournamentdate")
    sStateSQL = Session("StateSQL")
    sStateList = Session("StateList")
    sTourName = session("TournamentName")
ELSE
    sTourID = Request.QueryString("TourID")
END IF

curSanctionId = left(sTourID, 6)
curMemberId = Request.QueryString("MemberId")
curMemberFirstName = Request.QueryString("FirstName")
curMemberLastName = Request.QueryString("LastName")

curTraceMsg = curTraceMsg & "<br />TourId=" & sTourID & ", sTourName=" & sTourName & " (" & RemoveInvalidChars(sTourName) & "), sTourDate=" & sTourDate & ", sStateList=" & sStateList & ", sStateSQL=" & sStateSQL

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
curSqlStmt = "Select Distinct TSanction, TStatus, TournAppID, TDateE, TName, TCity, TState from " & SanctionTableName & " where TournAppID = '" & curSanctionId & "'"
rsWaterski.Open curSqlStmt
If rsWaterski.EOF THEN
	response.write "Invalid sanction number (" & curSanctionId & "), unable to complete request"
	response.status = "401 Unauthorized"
	response.flush
	response.end
ELSE
	strTStatus = rsWaterski("TStatus")
    strTSanction = rsWaterski("TSanction")
    strTourDate = rsWaterski("TDateE")
    sTourDate = strTourDate
    strTourName = rsWaterski("TName")
    sTourName = strTourName
END IF

curTraceMsg = curTraceMsg & "<br /><br />sTourID=" & sTourID & ", strTStatus=" & strTStatus & ", strTSanction=" & strTSanction & ", strTourDate=" & strTourDate & ", sTourDate=" & sTourDate

rsWaterski.Close
Set rsWaterski = Nothing
WaterskiConnect.Close

'	-----------------------------------------------------------------------
'Open database connection
'Check to determine if there are any qualification entries
'Then check to determine if there are qualifications for this tournamnet
'	-----------------------------------------------------------------------
Set WaterskiConnect = Server.CreateObject("ADODB.Connection")
WaterskiConnect.Open Application("WaterSkiConn")
Set rsWaterski = Server.CreateObject("ADODB.RecordSet")
rsWaterski.ActiveConnection = WaterskiConnect

Dim QfyNum, DateRaw, DateFmt, DateFmtForFile, I1, I2, RowNo
DateRaw = Date(): I1 = instr(DateRaw,"/"): I2 = instr(I1+1,DateRaw,"/")
DateFmt = Mid(DateRaw,I2+1): ' Start with Year value
IF I1=2 THEN DateFmt = DateFmt + "-0" + Left(DateRaw,1): ELSE DateFmt = DateFmt + "-" + Left(DateRaw,2)
IF I2-I1=2 THEN DateFmt = DateFmt + "-0" + Mid(DateRaw,I1+1,1): ELSE DateFmt = DateFmt + "-" + Mid(DateRaw,I1+1,2)
DateFmtForFile = Mid(DateFmt, 1, 4) + Mid(DateFmt, 6, 2) + Mid(DateFmt, 9, 2)

curSqlStmt = "Select count(*) as QfyNum From " & RegQualifyTableName & " Where left(TourID,6) = '" & curSanctionId & "';"
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

copyFileSour = pathExcelFiles & "/Templates/AWSATemplateBlank.xls"
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

'	-----------------------------------------------------------------------
' Refresh the list of chief and appointed officials for a tournament
' The data is stored in a temporary work table for use in build tournament registration entries
'	-----------------------------------------------------------------------
refreshApptOfficials(curSanctionId)

'	-----------------------------------------------------------------------
' Retrieve member entries for tournament registrations
' Include data from rankings, qualifications, membership status, and official ratings
'	-----------------------------------------------------------------------
Dim Counter0, Counter1, Counter2, Counter3
Dim rsMember

Set WaterskiConnect = Server.CreateObject("ADODB.Connection")
WaterskiConnect.Open Application("WaterSkiConn")
Set rsMember = Server.CreateObject("ADODB.RecordSet")
rsMember.ActiveConnection = WaterskiConnect

curSqlStmt = buildQueryMemberRegEntries(curSanctionId, sTourDate, sStateSQL, curMemberId, curMemberFirstName, curMemberLastName)

    On Error Resume Next
rsMember.Open curSqlStmt
    If Err.Number <> 0 Then
        %>
            <DIV ID="debugMsg">
                <br />Error opening SQL to retrieve skier list
                <br />Err.Number=<%=Err.Number %>
                <br />Err.Description=<%=Err.Description %>
                <br />SqlStmt <br /><%=curSqlStmt %>
                <br />
            </DIV>
        <%
        On Error Goto 0 ' But don't let other errors hide!
    End If

Counter0 = 0
Counter1 = 0
Counter2 = 0
Counter3 = 0

'	-----------------------------------------------------------------------
' Write retrieve member data to an Excel template spreadsheet file
'	-----------------------------------------------------------------------
Do until rsMember.EOF
    Counter0 = Counter0 + 1

	IF rsMember("PreReg") = "YES" OR len(rsMember("ApptdOfficial")) > 0 THEN

		IF rsMember("EventSlalom") = rsMember("Div") THEN
			EventSlalom = rsMember("EventSlalom"): SlalomPaid = rsMember("SlalomPaid")
		ELSE
			EventSlalom = "": SlalomPaid = ""
		END IF

		IF rsMember("EventTrick") = rsMember("Div") THEN
			EventTrick = rsMember("EventTrick"): TrickPaid = rsMember("TrickPaid")
		ELSE
			EventTrick = "": TrickPaid = ""
		END IF

		IF rsMember("EventJump") = rsMember("Div") THEN
			EventJump = rsMember("EventJump"): JumpPaid = rsMember("JumpPaid")
		ELSE
			EventJump = "": JumpPaid = ""
		END IF

		IF EventSlalom <> "" OR EventTrick <> "" OR EventJump <> "" OR len(rsMember("ApptdOfficial")) > 0 THEN
			Counter1 = Counter1 + 1: RowNo = FormatNumber(Counter1 + 5,0)

            objExcelPreReg.addnew
			objExcelPreReg.Fields(0).Value = rsMember("MemberID")
			objExcelPreReg.Fields(1).Value = rsMember("LastName")
			objExcelPreReg.Fields(2).Value = rsMember("FirstName")

			IF Mid(sTourID,4,3) = "999" THEN
				objExcelPreReg.Fields(3).Value = rsMember("Reg_Ski")
			END IF

			objExcelPreReg.Fields(4).Value = rsMember("Div")
			objExcelPreReg.Fields(5).Value = rsMember("Age")
			objExcelPreReg.Fields(6).Value = rsMember("City")
			objExcelPreReg.Fields(7).Value = rsMember("State")

			objExcelPreReg.Fields(8).Value = EventSlalom
			objExcelPreReg.Fields(9).Value = EventTrick
			objExcelPreReg.Fields(10).Value = EventJump

			objExcelPreReg.Fields(11).Value = rsMember("ApptdOfficial")

			objExcelPreReg.Fields(12).Value = rsMember("SlalomRank")
			objExcelPreReg.Fields(13).Value = rsMember("TrickRank")
			objExcelPreReg.Fields(14).Value = rsMember("JumpRank")

'			Insert Qualified Flags if Qualifications present, otherwise
'			Otherwise insert Ranking Levels by Events.

			IF QfyNum > 0 THEN
				objExcelPreReg.Fields(15).Value = rsMember("SlalomQfy")
				objExcelPreReg.Fields(16).Value = rsMember("TrickQfy")
				objExcelPreReg.Fields(17).Value = rsMember("JumpQfy")
			ELSE
				objExcelPreReg.Fields(15).Value = rsMember("SlalomRating")
				objExcelPreReg.Fields(16).Value = rsMember("TrickRating")
				objExcelPreReg.Fields(17).Value = rsMember("JumpRating")
			END IF
			objExcelPreReg.Fields(18).Value = rsMember("OverallRating")

			objExcelPreReg.Fields(19).Value = rsMember("TrickBoat")
			objExcelPreReg.Fields(20).Value = rsMember("JumpHeight")

			objExcelPreReg.Fields(21).Value = SlalomPaid
			objExcelPreReg.Fields(22).Value = TrickPaid
			objExcelPreReg.Fields(23).Value = JumpPaid

            objExcelPreReg.Fields(27).Value = rsMember("EffTo")

			IF rsMember("EffTo") >= cdate(sTourDate) and rsMember("CanSki") = True and rsMember("Waiver") > 0 THEN
		        objExcelPreReg.Fields(24).Value = "Yes"
				objExcelPreReg.Fields(25).Value = "Pre-Regist"
				objExcelPreReg.Fields(26).Value = FormatNumber(0,2)
			ELSE
				objExcelPreReg.Fields(24).Value = "No"
                objExcelPreReg.Fields(26).Value = rsMember("MembershipRate")
                objExcelPreReg.Fields(26).Value = rsMember("CostToUpgrade")

				' Figure applicable Renewal / Upgrade Amount based on MemType & Status
				IF rsMember("EffTo") < cdate(sTourDate) THEN
					IF rsMember("CanSki") = False THEN
						objExcelPreReg.Fields(25).Value = "Needs Renew/Upgrade"
						objExcelPreReg.Fields(26).Value = rsMember("MembershipRate")
					ELSE
						objExcelPreReg.Fields(25).Value = "Needs Renew"
						objExcelPreReg.Fields(26).Value = rsMember("MembershipRate")
					END IF
				ELSE
					IF rsMember("CanSkiGR") = True THEN
						objExcelPreReg.Fields(25).Value = "** Grass Roots Only"
                        objExcelPreReg.Fields(26).Value = rsMember("CostToUpgrade")
					ELSEIF rsMember("CanSki") = False THEN
						objExcelPreReg.Fields(25).Value = "Needs Upgrade"
						objExcelPreReg.Fields(26).Value = rsMember("MembershipRate")
					ELSE
						objExcelPreReg.Fields(25).Value = "Needs Annual Waiver"
						objExcelPreReg.Fields(26).Value = FormatNumber(0,2)
					END IF
				END IF
			END IF

			objExcelPreReg.Fields(32).Value = rsMember("JudgeSlalom")
			objExcelPreReg.Fields(33).Value = rsMember("JudgeTrick")
			objExcelPreReg.Fields(34).Value = rsMember("JudgeJump")
			objExcelPreReg.Fields(35).Value = rsMember("DriverSlalom")
			objExcelPreReg.Fields(36).Value = rsMember("DriverTrick")
			objExcelPreReg.Fields(37).Value = rsMember("DriverJump")
			objExcelPreReg.Fields(38).Value = rsMember("ScorerSlalom")
			objExcelPreReg.Fields(39).Value = rsMember("ScorerTrick")
			objExcelPreReg.Fields(40).Value = rsMember("ScorerJump")
			objExcelPreReg.Fields(41).Value = rsMember("Safety")
			objExcelPreReg.Fields(42).Value = rsMember("TechController")

			objExcelPreReg.Update

		END IF

	ELSEIF rsMember("EffTo") >= cdate(sTourDate) and rsMember("CanSki") = True and rsMember("Waiver") > 0 THEN

		Counter2 = Counter2 + 1
		objExcelActive.addnew
		objExcelActive.Fields(0).Value = rsMember("MemberID")
		objExcelActive.Fields(1).Value = rsMember("LastName")
		objExcelActive.Fields(2).Value = rsMember("FirstName")

		IF Mid(sTourID,4,3) = "999" THEN
			objExcelActive.Fields(3).Value = rsMember("Reg_Ski")
		END IF

		objExcelActive.Fields(4).Value = rsMember("Div")
		objExcelActive.Fields(5).Value = rsMember("Age")
		objExcelActive.Fields(6).Value = rsMember("City")
		objExcelActive.Fields(7).Value = rsMember("State")

		objExcelActive.Fields(12).Value = rsMember("SlalomRank")
		objExcelActive.Fields(13).Value = rsMember("TrickRank")
		objExcelActive.Fields(14).Value = rsMember("JumpRank")
		objExcelActive.Fields(15).Value = rsMember("SlalomRating")
		objExcelActive.Fields(16).Value = rsMember("TrickRating")
		objExcelActive.Fields(17).Value = rsMember("JumpRating")
		objExcelActive.Fields(18).Value = rsMember("OverallRating")

	    objExcelActive.Fields(24).Value = "Yes"
        objExcelActive.Fields(25).Value = rsMember("MemTypeDesc")
		objExcelActive.Fields(26).Value = FormatNumber(0,2)
        objExcelActive.Fields(27).Value = rsMember("EffTo")

		objExcelActive.Fields(32).Value = rsMember("JudgeSlalom")
		objExcelActive.Fields(33).Value = rsMember("JudgeTrick")
		objExcelActive.Fields(34).Value = rsMember("JudgeJump")
		objExcelActive.Fields(35).Value = rsMember("DriverSlalom")
		objExcelActive.Fields(36).Value = rsMember("DriverTrick")
		objExcelActive.Fields(37).Value = rsMember("DriverJump")
		objExcelActive.Fields(38).Value = rsMember("ScorerSlalom")
		objExcelActive.Fields(39).Value = rsMember("ScorerTrick")
		objExcelActive.Fields(40).Value = rsMember("ScorerJump")
		objExcelActive.Fields(41).Value = rsMember("Safety")
		objExcelActive.Fields(42).Value = rsMember("TechController")

		objExcelActive.Update

	ELSE
		Counter3 = Counter3 + 1
		objExcelInActive.addnew
		objExcelInActive.Fields(0).Value = rsMember("MemberID")
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
		objExcelInActive.Fields(12).Value = rsMember("SlalomRank")
		objExcelInActive.Fields(13).Value = rsMember("TrickRank")
		objExcelInActive.Fields(14).Value = rsMember("JumpRank")
		objExcelInActive.Fields(15).Value = rsMember("SlalomRating")
		objExcelInActive.Fields(16).Value = rsMember("TrickRating")
		objExcelInActive.Fields(17).Value = rsMember("JumpRating")
		objExcelInActive.Fields(18).Value = rsMember("OverallRating")

		objExcelInActive.Fields(24).Value = "    No"

		objExcelInActive.Fields(32).Value = rsMember("JudgeSlalom")
		objExcelInActive.Fields(33).Value = rsMember("JudgeTrick")
		objExcelInActive.Fields(34).Value = rsMember("JudgeJump")
		objExcelInActive.Fields(35).Value = rsMember("DriverSlalom")
		objExcelInActive.Fields(36).Value = rsMember("DriverTrick")
		objExcelInActive.Fields(37).Value = rsMember("DriverJump")
		objExcelInActive.Fields(38).Value = rsMember("ScorerSlalom")
		objExcelInActive.Fields(39).Value = rsMember("ScorerTrick")
		objExcelInActive.Fields(40).Value = rsMember("ScorerJump")
		objExcelInActive.Fields(41).Value = rsMember("Safety")
		objExcelInActive.Fields(42).Value = rsMember("TechController")

'	-----------------------------------------------------------------------
		' Figure applicable Renewal / Upgrade Amount based on MemType & Status
'	-----------------------------------------------------------------------
        objExcelInActive.Fields(27).Value = rsMember("EffTo")

		' Figure applicable Renewal / Upgrade Amount based on MemType & Status
		IF rsMember("EffTo") < cdate(sTourDate) THEN
			IF rsMember("CanSki") = False THEN
				objExcelInActive.Fields(25).Value = "Needs Renew/Upgrade"
				objExcelInActive.Fields(26).Value = rsMember("MembershipRate")
			ELSE
				objExcelInActive.Fields(25).Value = "Needs Renew"
				objExcelInActive.Fields(26).Value = rsMember("MembershipRate")
			END IF
		ELSE
			IF rsMember("CanSkiGR") = True THEN
				objExcelInActive.Fields(25).Value = "** Grass Roots Only"
                objExcelInActive.Fields(26).Value = rsMember("CostToUpgrade")
			ELSEIF rsMember("CanSki") = False THEN
				objExcelInActive.Fields(25).Value = "Needs Upgrade"
				objExcelInActive.Fields(26).Value = rsMember("MembershipRate")
			ELSE
				objExcelInActive.Fields(25).Value = "Needs Annual Waiver"
				objExcelInActive.Fields(26).Value = FormatNumber(0,2)
			END IF
		END IF

		objExcelInActive.Update

	END IF


	rsMember.MoveNext
Loop
curTraceMsg = curTraceMsg & "<br />Counter0=" & Counter0 & ", Counter1=" & Counter1 & ", Counter2=" & Counter2 & ", Counter3=" & Counter3

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
regTemplateFilename = "Entries-" & sStateList & "-" & DateFmtForFile

'	-----------------------------------------------------------------------
'Add the Tournament Name to the start of the file name
'	-----------------------------------------------------------------------
if len(sTourName) > 0 then
	regTemplateFilename = RemoveInvalidChars(sTourName) & "-" & regTemplateFilename
end if

'	-----------------------------------------------------------------------
'Append the username
'	-----------------------------------------------------------------------
if len(strTSanction) > 0 then
	regTemplateFilename = regTemplateFilename & "-" & strTSanction & ".xls"
else
	regTemplateFilename = regTemplateFilename & ".xls"
end if

curTraceMsg = curTraceMsg & "<br /><br />copyFileDest=" + pathExcelFiles & "/" & regTemplateFilename

fileRegXls.CopyFile copyFileDest, pathExcelFiles & "/" & regTemplateFilename , True

'	-----------------------------------------------------------------------
' Clean up old files
'	-----------------------------------------------------------------------
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
<title>Create Pre-Registration Export</title>

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
         <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><br>Your Pre-Registration
         Export workbook is now complete and ready to download.</font></td>
      </tr>

      <tr>
         <td>&nbsp;</td>
      </tr>

      <tr>
         <td>&nbsp;</td>
         <td><a href="excel/<% response.write regTemplateFilename %>"><font face="Arial" size="2"><b>RIGHT
         Click Here</b></font></a>&nbsp; <font size="2" face="Verdana, Arial, Helvetica, sans-serif">to
         download your Pre-Registration Export workbook, then select the "Save As"
         option from that menu, and then choose a suitable location to
         store the download in your PC. </font></td>
      </tr>

      <tr>
         <td>&nbsp;</td>
      </tr>

      <tr>
         <td>&nbsp;</td>
         <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
         After your Pre-Registration Export download has completed, then open the
         Excel file from that location on your PC.&nbsp; It will open automatically
         to an Instructions Tab.&nbsp; Please review that updated Instructions section
         for the latest information on contents and usage. </font></td>
      </tr>


<% IF QfyNum > 0 THEN %>

      <tr>
         <td>&nbsp;</td>
      </tr>

      <tr>
         <td>&nbsp;</td>
         <td><font color="#FF0000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>
         !! Qualification Indicators Included !!</strong>&nbsp;
         </font><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
         This tournament has qualification requirements.&nbsp; Members may
         enter in hopes of qualifying later.&nbsp; The three columns headed
         "Levels or Qualifications" in the Pre-Registration section indicate
         whether each pre-registered skier is actually qualified or not.&nbsp;
         Those with a <b>Y</b> in one of those three specific event columns
         are known to be qualified.&nbsp; Those without would need to submit
         proof of qualification to the Registrar.&nbsp; Upon seeing such proof
         of Qualification, the Registrar should place a <b>Y</b> in the applicable
         column.&nbsp; When processing your final entry list into WSTIMS, be sure
         that you respond "Yes" to the question about Qualifications being present
         in your entry list file(s).&nbsp; See the WSTIMS User Guide, and/or the
         instructions section in this Excel template, for more details on this
         important subject.
        </font></td>
      </tr>

<% END IF %>

      <tr>
         <td>&nbsp;</td>
      </tr>

      <tr>
         <td>&nbsp;</td>
         <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">After
         you've downloaded this Pre-Registration Export, you can later
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
''''< %

curTraceMsg = ""

%>
</body>
</html>
