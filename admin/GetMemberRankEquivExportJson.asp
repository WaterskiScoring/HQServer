<!--#include virtual="/admin/JSON_2.0.4.asp"-->
<!--#include virtual="/admin/MemberRegFunctions.asp"-->

<%
'	-----------------------------------------------------------------------
' Validate TourID value for scores to be Exported.
' http://usawaterski.org/admin/GetMemberRankEquivExportJson.asp?SanctionId=19U038&MemberId=700040630
' http://usawaterski.org/admin/GetMemberRankEquivExportJson.asp?SanctionId=19U038&FirstName=Jeff&LastName=Clark
' http://usawaterski.org/admin/GetMemberRankEquivExportJson.asp?SanctionId=19U038&State=MA
'	-----------------------------------------------------------------------

Dim curAuth, curAuthParts, curCredParts, curCount, curRqstAuth, curAuthResult
Dim curSanctionId, curMemberId, curStateSQL, curState, curTourYear, curTourDate

curRqstAuth = 0
curRqstAuth = CheckBasicAuth()
IF curRqstAuth = 0 THEN
	response.status = "401 Unauthorized - Invalid credentials"
    response.Write(response.Status)
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

''''response.write "<br />curSanctionId=" & curSanctionId & ", Region=" & Mid(curSanctionId, 3, 1) & ", sTourName=" & ", curState=" & curState & ":"
''''response.write "<br />"
''''response.End

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
''''response.write "<br />curSqlStmt=" & curSqlStmt & ":<br/>"

rsWaterski.Open curSqlStmt
If rsWaterski.EOF THEN
	response.status = "200 401 Unauthorized - Invalid sanction number (" & curSanctionId & ")"
    response.Write(response.Status)
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

''''response.write "<br />strTSanction=" & strTSanction & ", strTourName=" & strTourName & "<br/>"

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
curSqlStmt = buildQueryMemberRankingEquivalents(curSanctionId, curTourDate, curStateSQL, curMemberId, curMemberFirstName, curMemberLastName)
''''response.write "<br />curSqlStmt=" & curSqlStmt & ":<br/>"

'	-----------------------------------------------------------------------
' Execute SQL statement to retrieve skier information and load to registration template
'	-----------------------------------------------------------------------
    On Error Resume Next
response.ContentType="application/json"
response.status = "200 Completed"
        If Err.Number <> 0 Then
            response.write "<br />Err.Number=" & Err.Number & ", Err.Description=" & Err.Description & " |"
            On Error Goto 0
        End If

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
