<!--#include virtual="/admin/JSON_2.0.4.asp"-->

<% 

If not Session("aauth") then response.redirect "Login.asp"

Server.ScriptTimeout = 300

Dim curTraceMsg, curSanctionId, curMemberId, curStateSQL, curStateList, curTourYear, curDate
Dim sUserName

'	-----------------------------------------------------------------------
' Validate TourID value for scores to be Exported.
' http://usawaterski.org/admin/GetOfficialExportJson.asp?MemberId=700040630
' http://usawaterski.org/admin/GetOfficialExportJson.asp?SanctionId=18E024
' http://usawaterski.org/admin/GetOfficialExportJson.asp?StateList=MA,CT
'	-----------------------------------------------------------------------
sUserName = session("UserName")
curSanctionId = Request.QueryString("SanctionId")
curMemberId = Request.QueryString("MemberId")
curStateList = Request.QueryString("StateList")

IF len(curSanctionId) > 0 THEN
    curTourYear = 2000 + left(curSanctionId,2)
ELSE 
    curDate = Date
    curTourYear = 2018
    'curTourYear = Right(curDate, 4)
    'curMonth = Left(curDate, 2)
    'if ( curMonth > 8 ) THEN curTourYear = curTourYear + 1
END IF	

IF len(curStateList) > 0 THEN
    curStateSQL = BuildStateSQL(curStateList)
ELSE
    curStateSQL = "State IN ('')"
END IF


curTraceMsg = curTraceMsg & "<br />curSanctionId=" & curSanctionId & ", curTourYear=" & curTourYear& ", curDate=" & curDate  & ", curMemberId=" & curMemberId & ", curStateList=" & curStateList & ", curStateSQL=" & curStateSQL

'	-----------------------------------------------------------------------
'	Utility function defintion
'	-----------------------------------------------------------------------
Function QueryToJSON(dbc, sql)
        Dim rs, jsa
        Set rs = dbc.Execute(sql)
        Set jsa = jsArray()
        While Not (rs.EOF Or rs.BOF)
                Set jsa(Null) = jsObject()
                For Each col In rs.Fields
                        jsa(Null)(col.Name) = col.Value
                Next
        rs.MoveNext
        Wend
        Set QueryToJSON = jsa
End Function

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

Function BuildStateSQL(ListValues)
    Dim WorkingString
    Dim StateCounter 
    Dim StateSQL 
    StateCounter = 1
    WorkingString = ListValues
    'Get the first state
    LocationofComma = instr(WorkingString,",")
    if LocationofComma > 0 then
	    StateSQL = "'" & left(WorkingString, (LocationofComma - 1)) & "'"
    else
	    StateSQL = "'" & WorkingString & "'"
    end if

    While instr(WorkingString,",") > 0 
	    LocationofComma = instr(WorkingString,",")
	    'Now trim the string
	    WorkingString = right(WorkingString, len(WorkingString) - (LocationofComma))
	    StateCounter = StateCounter + 1
	    StateSQL = StateSQL & ",'" & left(WorkingString, (LocationofComma - 1)) & "'"
    wend

    BuildStateSQL = "State IN (" & StateSQL & ")"
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Now build a Query to Extract the Desired Members, joining in data 
' from the Rankings and Officials and Membership Type tables.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim curSqlStmt
curSqlStmt = ""
curSqlStmt = curSqlStmt & "SELECT MT.PersonIDWithCheckDigit as MemberID, MT.PersonID , MT.LastName, MT.FirstName"
curSqlStmt = curSqlStmt & "	, Upper(Left(MT.Sex,1)) as Sex, Left(MT.City,12) as City, Left(MT.State,2) as State"
curSqlStmt = curSqlStmt & "	, 2018 - Year(MT.BirthDate) - 1 as Age , MT.EffectiveTo as EffTo, OT.RatingType_ID"
curSqlStmt = curSqlStmt & "	, Case When OT.RatingType_ID = 1 Then 'Judge'"
curSqlStmt = curSqlStmt & "			When OT.RatingType_ID = 2 Then 'Scorer'"
curSqlStmt = curSqlStmt & "			When OT.RatingType_ID = 3 Then 'Driver'" 
curSqlStmt = curSqlStmt & "			When OT.RatingType_ID = 9 Then 'Satety'" 
curSqlStmt = curSqlStmt & "			Else Convert(char(1), OT.RatingType_ID) END as RatingType"
curSqlStmt = curSqlStmt & "	, LV.Abbreviation as Rating, LV.LevelOrderforTemplate as RatingOrder, OT.EventsConsolidated as Events"
curSqlStmt = curSqlStmt & "	, Typ.CanSkiInTournaments as CanSki , Typ.CanSkiInGRTournaments as CanSkiGR, MT.WaiverStatusID as Waiver "
curSqlStmt = curSqlStmt & "FROM USAWaterski.dbo.Members as MT"
curSqlStmt = curSqlStmt & "    INNER JOIN USAWaterski.dbo.MembershipTypes as Typ ON MT.MembershipTypeCode = Typ.MemberShipTypeID"
curSqlStmt = curSqlStmt & "	   INNER JOIN USAWaterski.dbo.Officials as OT ON OT.PersonID = MT.PersonID"
curSqlStmt = curSqlStmt & "	   INNER JOIN USAWaterski.dbo.Level as LV ON OT.Level_ID = LV.Level_ID AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL "
curSqlStmt = curSqlStmt & "WHERE Typ.ExporttoTouramentRegistrationTemplate = 1"
curSqlStmt = curSqlStmt & "  AND DateAdd(mm,18,MT.EffectiveTo) > GetDate()"
curSqlStmt = curSqlStmt & "  AND MT.Deceased = 0"

curSqlStmt = curSqlStmt & "  AND (" & curStateSQL 
IF len(curSanctionId) > 0 THEN
    curSqlStmt = curSqlStmt & "       OR MT.PersonID in (Select PersonID from USAWaterski.dbo.TempApptdOfcls WHERE TournAppID = '" & curSanctionId & "' )"
    curSqlStmt = curSqlStmt & "       OR MT.PersonIDWithCheckDigit in (Select MemberID from Cobra00025.USAWSRank.RegisterEvents WHERE left(TourID,6) = '" & curSanctionId & "')"
END IF
IF len(curMemberId) > 0 THEN
    curSqlStmt = curSqlStmt & "       OR MT.PersonIDWithCheckDigit = '" & curMemberId & "' "
END IF

curSqlStmt = curSqlStmt & "       ) "
curSqlStmt = curSqlStmt & "Order by MT.LastName, MT.FirstName, OT.RatingType_ID, LV.LevelOrderforTemplate"

'	-----------------------------------------------------------------------
' Execute SQL statement to retrieve skier information and load to registration template
'	-----------------------------------------------------------------------
'	-----------------------------------------------------------------------
'	-----------------------------------------------------------------------
dim Counter
Set WaterskiConnect = Server.CreateObject("ADODB.Connection")
WaterskiConnect.Open Application("WaterSkiConn")

response.ContentType="application/json"

    On Error Resume Next
QueryToJSON(WaterskiConnect,curSqlStmt).Flush
    If Err.Number <> 0 Then
        %>
            <DIV ID="debugMsg">
                <br />Error QueryToJSON
                <br />Err.Number=<%=Err.Number %>
                <br />Err.Description=<%=Err.Description %>
                <br /><br />
            </DIV>
        <%
        On Error Goto 0 ' But don't let other errors hide!
    End If

%>
