<!--#include virtual="/admin/JSON_2.0.4.asp"-->
<!--#include virtual="/admin/MemberRegFunctions.asp"-->

<%
'	-----------------------------------------------------------------------
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

'	-----------------------------------------------------------------------
'Open connection to Sanction Database
'Get tournament attributes from TSchedul table
'	-----------------------------------------------------------------------
Set WaterskiConnect = Server.CreateObject("ADODB.Connection")
WaterskiConnect.Open Application("WaterSkiConn")

Dim curSqlStmt
curSqlStmt = "Select TSanction, TournAppID, EditCode, TDateS, TDateE, TName"
curSqlStmt = curSqlStmt & ", TSiteID, TSite, TCity, TState, TSponsor"
curSqlStmt = curSqlStmt & ", TYear, TSkiYr, TSanApproved"
curSqlStmt = curSqlStmt & ", TEventSlalom, TEventJump, TEventTrick"
curSqlStmt = curSqlStmt & ", TDirName, TDirAddress, TDirCity, TDirState, TDirZip, TDirEmail, TDirPhoneAm"
curSqlStmt = curSqlStmt & ", TEntryFees, TEntryLimit, TRegistrarName, TRegistrarPhone, TRegistrarEmail, TRegistrarAddr, TRegistrarCity, TRegistrarState, TRegistrarZip"
curSqlStmt = curSqlStmt & ", TRoundsS, TRoundsT, TRoundsJ "
curSqlStmt = curSqlStmt & "FROM " & SanctionTableName & " "
curSqlStmt = curSqlStmt & "WHERE TournAppID = '" & curSanctionId & "' AND TsanApproved = 1"

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
