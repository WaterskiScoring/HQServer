<!--#include virtual="/admin/JSON_2.0.4.asp"-->
<!--#include virtual="/admin/MemberRegFunctions.asp"-->

<%
'	-----------------------------------------------------------------------
' Validate TourID value for scores to be Exported.
' http://usawaterski.org/admin/GetMemberRegExportJson.asp?SanctionId=18E024&MemberId=700040630
' http://usawaterski.org/admin/GetMemberRegExportJson.asp?SanctionId=18E024&FirstName=Jeff&LastName=Clark
' http://usawaterski.org/admin/GetMemberRegExportJson.asp?SanctionId=18E024&FirstName=Jeff&LastName=Clark&user=18E024&password=10089
''	-----------------------------------------------------------------------

Dim curAuth, curAuthParts, curCredParts, curCount, curRqstAuth, curAuthResult
Dim curSanctionId, curMemberId, curStateSQL, curState, curTourYear, curTourDate

curRqstAuth = 0
curRqstAuth = CheckBasicAuth()
IF curRqstAuth = 0 THEN
	response.write "Invalid credentials, unable to complete request"
	response.status = "401 Unauthorized"
	response.flush
	response.end
END IF

curSanctionId = Request.QueryString("SanctionId")
curMemberId = Request.QueryString("MemberId")
curMemberFirstName = Request.QueryString("FirstName")
curMemberLastName = Request.QueryString("LastName")
curState = Request.QueryString("State")
curStateSQL = ""
IF len(curState) > 0 THEN
    curStateSQL = "State = '" & curState & "'"
END IF

''''response.write "<br />curSanctionId=" & curSanctionId & ", Mid(curSanctionId, 3, 1)=" & Mid(curSanctionId, 3, 1) & ", sTourName=" & ", curState=" & curState

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
curSqlStmt = "Select Top 1 TSanction, TStatus, TournAppID, TDateE, TName, TCity, TState from Sanctions.dbo.TSchedul where TournAppID = '" & curSanctionId & "'"
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
    curTourDate = strTourDate
    strTourName = rsWaterski("TName")
    sTourName = strTourName
    curTourYear = 2000 + left(curSanctionId,2)
END IF

rsWaterski.Close
Set rsWaterski = Nothing
WaterskiConnect.Close

'	-----------------------------------------------------------------------
' Retrieve member entries for tournament registrations
' Include data from rankings, qualifications, membership status, and official ratings
'	-----------------------------------------------------------------------
Dim Counter0, Counter1, Counter2, Counter3
Dim rsMember

Set WaterskiConnect = Server.CreateObject("ADODB.Connection")
WaterskiConnect.Open Application("WaterSkiConn")

curMemberLastName = Replace(curMemberLastName, "'", "''")
curMemberFirstName = Replace(curMemberFirstName, "'", "''")

curSqlStmt = ""
IF len(curMemberId) > 0 OR len(curMemberFirstName) > 0  OR len(curMemberLastName) > 0 OR len(curStateSQL) > 0 THEN
    curSqlStmt = buildQueryMemberRegEntries(curSanctionId, curTourDate, curStateSQL, curMemberId, curMemberFirstName, curMemberLastName)

ELSEIF Mid(curSanctionId, 3, 1) = "U" THEN
    curSqlStmt = buildQueryMemberRegNcwsaEntries(curSanctionId, curTourDate)

ELSE
    curSqlStmt = buildQueryMemberRegEntries(curSanctionId, curTourDate, curStateSQL, curMemberId, curMemberFirstName, curMemberLastName)
END IF

'	-----------------------------------------------------------------------
' Execute SQL statement to retrieve skier information and load to registration template
'	-----------------------------------------------------------------------
response.ContentType="application/json"
response.status = "200 Completed"
    On Error Resume Next
QueryToJSON(WaterskiConnect, curSqlStmt).flush
    If Err.Number <> 0 Then
        %>
            <DIV ID="debugMsg">
                <br />Error retrieving member registration information
                <br />Err.Number=<%=Err.Number %>
                <br />Err.Description=<%=Err.Description %>
                <br />
            </DIV>
        <%
        On Error Goto 0 ' But don't let other errors hide!
    End If

%>
