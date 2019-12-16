<%
'	-----------------------------------------------------------------------
' Common functions and definitions used for member registration and related functionality
'	-----------------------------------------------------------------------
DIM TRADBName, MemberDBName, SanctionDBName
DIM MemberTableName, UsersTableName, OfficialsTableName, RatingLevelTableName, MembershipTypesTableName

TRADBName = "00025"
MemberDBName = "USAWaterski"
SanctionDBName = "Sanctions"

MemberTableName = "USAWaterski.dbo.members"
UsersTableName = "USAWaterski.dbo.Users999"
OfficialsTableName = "USAWaterski.dbo.Officials"
RatingLevelTableName = "USAWaterski.dbo.Level"
MembershipTypesTableName = "USAWaterski.dbo.MembershipTypes"
RegEventsTableName = "Cobra00025.USAWSRank.RegisterEvents"
ApptOfficialsTableName = "USAWaterski.dbo.TempApptdOfcls"
RegTableName = "sanctions.dbo.registration"
RankingsTableName = "Cobra00025.USAWSRank.Rankings"
EliteDatesTableName = "Cobra00025.USAWSRank.EliteDates"
RegGenTableName = "Cobra00025.USAWSRank.RegisterGen_05042014"
RegQualifyTableName = "Cobra00025.USAWSRank.RegisterQualify_TEST"
MembershipRatesTableName = "waterski.dbo.MembershipTypeRates"
TeamRotationsTableName = "Cobra00025.USAWSRank.TeamRotations"
TeamTableName = "Cobra00025.USAWSRank.TeamsList"
SanctionTableName = "Sanctions.dbo.TSchedul"
TeamRosterTableName = "Cobra00025.USAWSRank.TeamRoster"

Function CheckBasicAuth()
	Dim curAuth, curRqstAuth, curSessionUsernName

	'	-----------------------------------------------------------------------
	' Check for authorization for WSTIMS for Windows request for official ratings
	'	-----------------------------------------------------------------------
	curRqstAuth = 0
    curSessionUsernName = session("UserName")
    ''''curAuth = Request.ServerVariables("HTTP_AUTHORIZATION")
    curAuth = Request.ServerVariables("HTTP_WSTIMS")
	''''response.write "<br/>UserName=" & curSessionUsernName & "<br/>"
    ''''response.write "<br/>WSTIMS_AUTHORIZATION=" & curAuth & "<br/>"
	''''response.end

	IF len(curSessionUsernName) > 0 THEN
		curRqstAuth = 1

	ELSEIF len(curAuth) > 0 THEN
		curAuthParts = Split(curAuth, " ")
		IF IsArray(curAuthParts) THEN
			IF curAuthParts(0) = "Basic" THEN
				curCredParts = Split(curAuthParts(1), ":")
				curTraceMsg = curTraceMsg & "<br />curCredParts:Count=" & UBound(curCredParts) & ", IsArray=" & IsArray(curCredParts)
				IF IsArray(curCredParts) THEN
					IF curCredParts(0) = "wstims" AND curCredParts(1) = "Slalom38tTrick13Jump250" THEN
						curRqstAuth = 1
					ELSE
						curRqstAuth = ValidateSanctionAccess(curCredParts(0), curCredParts(1))
					END IF
				END IF
			END IF
		END IF
	
    ELSEIF len(curAuth) > 0 THEN
        curAuth= Request.QueryString("WSTIMS")
		curCredParts = Split(curAuthParts(1), ":")
		curTraceMsg = curTraceMsg & "<br />curCredParts:Count=" & UBound(curCredParts) & ", IsArray=" & IsArray(curCredParts)
		IF IsArray(curCredParts) THEN
			IF curCredParts(0) = "wstims" AND curCredParts(1) = "Slalom38tTrick13Jump250" THEN
				curRqstAuth = 1
			ELSE
				curRqstAuth = ValidateSanctionAccess(curCredParts(0), curCredParts(1))
			END IF
		END IF
    
    ELSE
        curUser = Request.QueryString("user")
        curpassword = Request.QueryString("password")
		curRqstAuth = ValidateSanctionAccess(curUser, curpassword)
	END IF
	CheckBasicAuth = curRqstAuth
End Function

Function QueryToJSON(dbc, sql)
        Dim rs, jsa
            On Error Resume Next
        Set rs = dbc.Execute(sql)
        If Err.Number <> 0 Then
            %>
                <DIV ID="debugMsg">
                    <br />Error running SQL statement and converting to JSON
                    <br />Err.Number=<%=Err.Number %>
                    <br />Err.Description=<%=Err.Description %>
                    <br />
                </DIV>
            <%
            On Error Goto 0 ' But don't let other errors hide!
        End If

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

Function BuildStateSQL(ListValues)
    Dim WorkingString, StateCounter, StateSQL
    StateCounter = 1
    WorkingString = ListValues
	'	-----------------------------------------------------------------------
    'Get the first state
	'	-----------------------------------------------------------------------
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

Function ValidateSanctionAccess(SanctionId, EditCode)
	'	-----------------------------------------------------------------------
	'Open connection to Sanction Database
	'Get tournament attributes from TSchedul table
	'	-----------------------------------------------------------------------
	Set WaterskiConnect = Server.CreateObject("ADODB.Connection")
	WaterskiConnect.Open Application("WaterSkiConn")
	Set rsWaterski = Server.CreateObject("ADODB.RecordSet")
	rsWaterski.ActiveConnection = WaterskiConnect

	Dim curSqlStmt
	curSqlStmt = ""
	curSqlStmt = curSqlStmt & "SELECT TournamentName, name FROM " & UsersTableName & " "
	curSqlStmt = curSqlStmt & "Where name = '" & SanctionId & "' "
	curSqlStmt = curSqlStmt & "And pwd = '" & EditCode & "' "

	Dim returnValue
    Set rsWaterski = WaterskiConnect.Execute(curSqlStmt)
	If rsWaterski.EOF THEN
		curTraceMsg = "<br />rsWaterski.EOF"
		response.write curTraceMsg
		returnValue = 0
	ELSE
		TourName = rsWaterski("TournamentName")
		IF len(TourName) > 0 THEN
			returnValue = 1
		ELSE
			returnValue = 0
		END IF
	END IF

	rsWaterski.Close
	Set rsWaterski = Nothing
	WaterskiConnect.Close
	ValidateSanctionAccess = returnValue

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

Function PersonIDwChkDgt (PersonID)
    ' ---------------------------------------------------
    ' This function is given an integer "PersonID" value, and returns the
    ' 9-Character "PersonIDWithCheckDigit" value for that particular member.
    ' ---------------------------------------------------
    Dim PIDSum, PIDChar, PIDLen, PIDPtr
    PIDSum = 0: PIDChar = trim(PersonID): PIDLen = Len(PIDChar)

    FOR PIDPtr = 1 TO PIDLen STEP 2
	    PIDSum = PIDSum + (3*MID(PIDChar,PIDPtr,1))
	    IF PIDPtr+1 <= PIDLen THEN PIDSum = PIDSum + MID(PIDChar,PIDPtr+1,1)
    NEXT

    PersonIDwChkDgt = right(100-PIDSum,1) & Right(100000000+PersonID,8)
End Function

'	-----------------------------------------------------------------------
' Build a query to extract member entries for tournament registrations
' Include data from rankings, qualifications, membership status, and official ratings
'	-----------------------------------------------------------------------
Function buildQueryMemberRegEntries(curSanctionId, curTourDate, curStateSQL, curMemberId, curMemberFirstName, curMemberLastName)

    Dim curTourYear, curSqlStmt1, curSqlStmt2, curSqlStmt3, curSqlStmt4, curSqlStmt5, curSqlStmt6, curSqlStmt7, curSqlStmt8
    curTourYear = 2000 + left(curSanctionId, 2)

    'Member Number and name
    curSqlStmt1 = ""
    curSqlStmt1 = curSqlStmt1 & "Select Substring(MX.MemberID,1,3) + '-' + Substring(MX.MemberID,4,2) + '-' + Substring(MX.MemberID,6,4) as MemberID"
    curSqlStmt1 = curSqlStmt1 & ", MX.LastName, MX.FirstName"

    'Skier division
    curSqlStmt1 = curSqlStmt1 & ", Coalesce(RD.Div"
    curSqlStmt1 = curSqlStmt1 & "   , Case when MX.Age <= 17 and MX.Sex = 'F' Then 'G'"
    curSqlStmt1 = curSqlStmt1 & "          when MX.Age <= 17 then 'B' when MX.Sex = 'F' then 'W' else 'M' end"
    curSqlStmt1 = curSqlStmt1 & "   + Case"
    curSqlStmt1 = curSqlStmt1 & "          when MX.Age <= 9 then '1'"
    curSqlStmt1 = curSqlStmt1 & "          when MX.Age <= 13 then '2'"
    curSqlStmt1 = curSqlStmt1 & "          when MX.Age <= 17 then '3'"
    curSqlStmt1 = curSqlStmt1 & "          when MX.Age <= 24 then '1'"
    curSqlStmt1 = curSqlStmt1 & "          when MX.Age <= 34 then '2'"
    curSqlStmt1 = curSqlStmt1 & "          when MX.Age <= 44 then '3'"
    curSqlStmt1 = curSqlStmt1 & "          when MX.Age <= 54 then '4'"
    curSqlStmt1 = curSqlStmt1 & "          when MX.Age <= 59 then '5'"
    curSqlStmt1 = curSqlStmt1 & "          when MX.Age <= 64 then '6'"
    curSqlStmt1 = curSqlStmt1 & "          when MX.Age <= 69 then '7'"
    curSqlStmt1 = curSqlStmt1 & "          when MX.Age <= 74 then '8'"
    curSqlStmt1 = curSqlStmt1 & "          when MX.Age <= 79 then '9'"
    curSqlStmt1 = curSqlStmt1 & "          when MX.Age <= 84 then 'A'"
    curSqlStmt1 = curSqlStmt1 & "          else 'B' end) as Div"

    'Skier information
    curSqlStmt1 = curSqlStmt1 & ", MX.Age, MX.Sex as Gender, MX.City, MX.State, Coalesce(MX.Federation, '') as Federation, MX.Waiver"

    'Skier official ratings
    curSqlStmt1 = curSqlStmt1 & ", Coalesce(SO.OffCode,'') as ApptdOfficial"

    curSqlStmt1 = curSqlStmt1 & ", Coalesce(SX.SlalomRank,'') as SlalomRank"
    curSqlStmt1 = curSqlStmt1 & ", Coalesce(TX.TrickRank,'') as TrickRank"
    curSqlStmt1 = curSqlStmt1 & ", Coalesce(JX.JumpRank,'') as JumpRank"

    'Event Ratings and qualifications
    curSqlStmt1 = curSqlStmt1 & ", Coalesce(SE.SlmEli,SX.SlalomRating,'') as SlalomRating"
    curSqlStmt1 = curSqlStmt1 & ", Coalesce(TE.TrkEli,TX.TrickRating,'') as TrickRating"
    curSqlStmt1 = curSqlStmt1 & ", Coalesce(JE.JmpEli,JX.JumpRating,'') as JumpRating"
    curSqlStmt1 = curSqlStmt1 & ", Coalesce(OE.OvrEli,OX.OverallRating,'') as OverallRating"

    curSqlStmt1 = curSqlStmt1 & ", Case when PS.SQfyOvr > '   ' then 'Y' else QS.SQfy end as SlalomQfy"
    curSqlStmt1 = curSqlStmt1 & ", Case when PT.TQfyOvr > '   ' then 'Y' else QT.TQfy end as TrickQfy"
    curSqlStmt1 = curSqlStmt1 & ", Case when PJ.JQfyOvr > '   ' then 'Y' else QJ.JQfy end as JumpQfy"

    'Skier event attributes and payments
    curSqlStmt1 = curSqlStmt1 & ", Coalesce(PT.TrickBoat,'') as TrickBoat"
    curSqlStmt1 = curSqlStmt1 & ", Coalesce(PR.JRamp,'') as JumpHeight"
    curSqlStmt1 = curSqlStmt1 & ", PR.Prereg"

	curSqlStmt1 = curSqlStmt1 & ", Case When PS.EventSlalom = null THEN '' WHEN RD.Div = null THEN PS.EventSlalom When PS.EventSlalom = RD.Div THEN PS.EventSlalom ELSE '' END as EventSlalom"
	curSqlStmt1 = curSqlStmt1 & ", Case When PT.EventTrick = null THEN '' WHEN RD.Div = null THEN PT.EventTrick When PT.EventTrick = RD.Div THEN PT.EventTrick ELSE '' END as EventTrick"
	curSqlStmt1 = curSqlStmt1 & ", Case When PJ.EventJump = null THEN '' WHEN RD.Div = null THEN PJ.EventJump When PJ.EventJump = RD.Div THEN PJ.EventJump ELSE '' END as EventJump"

    curSqlStmt1 = curSqlStmt1 & ", Coalesce(PS.SfeeCls,'') + Coalesce(PS.SFeeRds,'') as SlalomPaid"
    curSqlStmt1 = curSqlStmt1 & ", Coalesce(PT.TfeeCls,'') + Coalesce(PT.TFeeRds,'') as TrickPaid"
    curSqlStmt1 = curSqlStmt1 & ", Coalesce(PJ.JfeeCls,'') + Coalesce(PJ.JFeeRds,'') as JumpPaid"

    'Other member stuff
    curSqlStmt1 = curSqlStmt1 & ", MX.EffTo, MX.Memtype, MX.MemCode, MX.ActiveMember, MX.MemTypeDesc, MX.CanSki, MX.CanSkiGR, MX.SptsDiv, MembershipRate, CostToUpgrade"

    curSqlStmt1 = curSqlStmt1 & ", Case WHEN OPS.Rating = 'INT' THEN 'PanAm' ELSE Coalesce(OJS.Rating, '') END as JudgeSlalom"
    curSqlStmt1 = curSqlStmt1 & ", Case WHEN OPT.Rating = 'INT' THEN 'PanAm' ELSE Coalesce(OJT.Rating, '') END as JudgeTrick"
    curSqlStmt1 = curSqlStmt1 & ", Case WHEN OPJ.Rating = 'INT' THEN 'PanAm' ELSE Coalesce(OJJ.Rating, '') END as JudgeJump"
    curSqlStmt1 = curSqlStmt1 & ", Coalesce(ODS.Rating, '') as DriverSlalom, Coalesce(ODT.Rating, '') as DriverTrick, Coalesce(ODJ.Rating, '') as DriverJump"
    curSqlStmt1 = curSqlStmt1 & ", Coalesce(OCS.Rating, '') as ScorerSlalom, Coalesce(OCT.Rating, '') as ScorerTrick, Coalesce(OCJ.Rating, '') as ScorerJump"
    curSqlStmt1 = curSqlStmt1 & ", Coalesce(OS.Rating, '') as Safety, Coalesce(OTC.Rating, '') as TechController "

    '	-----------------------------------------------------------------------
    'FROM Statement
    '	-----------------------------------------------------------------------
    curSqlStmt2 = ""
    curSqlStmt2 = curSqlStmt2 & " FROM ("

    '	-----------------------------------------------------------------------
    'Use select as a data source for member data
    '	-----------------------------------------------------------------------
    curSqlStmt2 = curSqlStmt2 & "    SELECT MT.PersonIDWithCheckDigit as MemberID, MT.PersonID, MT.LastName, FirstName, MT.FederationCode as Federation"
    curSqlStmt2 = curSqlStmt2 & "        , (" & curTourYear & " - Year(MT.BirthDate) - 1) as Age, Upper(Left(MT.Sex,1)) as Sex, MT.WaiverStatusID as Waiver"
    curSqlStmt2 = curSqlStmt2 & "        , MT.City, Left(MT.State,2) as State"
    curSqlStmt2 = curSqlStmt2 & "        , MT.EffectiveTo as EffTo, MT.MembershipTypeCode as MemType"
    curSqlStmt2 = curSqlStmt2 & "        , MT.Deceased, MT.DivisionCode1 + '/' + MT.DivisionCode2 as SptsDiv"
    curSqlStmt2 = curSqlStmt2 & "        , Typ.ExporttoTouramentRegistrationTemplate as ExportToTemplate"
    curSqlStmt2 = curSqlStmt2 & "        , Typ.TypeCode as MemCode, Typ.ActiveMember, Typ.Description as MemTypeDesc, Typ.CanSkiInTournaments as CanSki, Typ.CanSkiInGRTournaments as CanSkiGR"
    curSqlStmt2 = curSqlStmt2 & "        , Coalesce(MR.MembershipTypeRates, 0) as MembershipRate, Coalesce(MR.CosttoUpgrade, 0) as CostToUpgrade"
    curSqlStmt2 = curSqlStmt2 & "    FROM " & MemberTableName & " as MT "
    curSqlStmt2 = curSqlStmt2 & "      INNER JOIN " & MembershipTypesTableName & " as Typ ON MT.MembershipTypeCode = Typ.MemberShipTypeID "
    curSqlStmt2 = curSqlStmt2 & "      LEFT JOIN " & MembershipRatesTableName & " as MR ON MR.[Membership Type Code] = MT.MembershipTypeCode "
    curSqlStmt2 = curSqlStmt2 & "           AND MR.EffectiveFrom <= CONVERT(DATETIME, '" & curTourDate & " 00:00:00', 102)"
    curSqlStmt2 = curSqlStmt2 & "           AND MR.EffectiveTo >= CONVERT(DATETIME, '" & curTourDate & " 00:00:00', 102)"
    curSqlStmt2 = curSqlStmt2 & "    ) as MX"

    '	-----------------------------------------------------------------------
    ' Use select as a data source for chief officials
    '	-----------------------------------------------------------------------
    curSqlStmt3 = curSqlStmt3 & "     LEFT JOIN " & ApptOfficialsTableName & " AS SO ON SO.PersonID = MX.PersonID AND TournAppID = '" & curSanctionId & "' "

    '	-----------------------------------------------------------------------
    ' Use select as a data source for skier ranking data
    ' The RD subquery below UNIONS selects from Rankings PLUS RegisterEvents, to
    ' ensure that EVERY entered skier will show up SOMEWHERE in the template.
    '	-----------------------------------------------------------------------
    curSqlStmt4 = ""
    curSqlStmt4 = curSqlStmt4 & "  LEFT JOIN ("
    curSqlStmt4 = curSqlStmt4 & "      SELECT MemberID, Div FROM " & RankingsTableName & ""
    curSqlStmt4 = curSqlStmt4 & "      WHERE SkiYearID = 1 and RankScore is not Null AND Left(Div,1) in ('B','G','M','W','O')"
    curSqlStmt4 = curSqlStmt4 & "      UNION"
    curSqlStmt4 = curSqlStmt4 & "      SELECT MemberID, Div FROM " & RegEventsTableName & ""
    curSqlStmt4 = curSqlStmt4 & "      WHERE left(TourID,6) = '" & curSanctionId & "' "
    curSqlStmt4 = curSqlStmt4 & "      GROUP BY MemberID, Div)"
    curSqlStmt4 = curSqlStmt4 & "      AS RD ON RD.MemberID = MX.MemberID"

    '	-----------------------------------------------------------------------
    ' Slalom ratings
    '	-----------------------------------------------------------------------
    curSqlStmt4 = curSqlStmt4 & "  LEFT JOIN ("
    curSqlStmt4 = curSqlStmt4 & "      SELECT MemberID, Div, Reg_Ski, AWSA_Rat as SlalomRating, Left(Cast(Cast(RankScore as Decimal(7,2)) as Varchar(8)),6) as SlalomRank"
    curSqlStmt4 = curSqlStmt4 & "      FROM " & RankingsTableName
    curSqlStmt4 = curSqlStmt4 & "      WHERE SkiYearID = 1"
    curSqlStmt4 = curSqlStmt4 & "        AND Left(Div,1) in ('B','G','M','W','O')"
    curSqlStmt4 = curSqlStmt4 & "        AND Event = 'S'"
    curSqlStmt4 = curSqlStmt4 & "        AND RankScore is not null)"
    curSqlStmt4 = curSqlStmt4 & "      AS SX ON MX.MemberID = SX.MemberID AND RD.Div = SX.Div"

    '	-----------------------------------------------------------------------
    ' Trick ratings
    '	-----------------------------------------------------------------------
    curSqlStmt4 = curSqlStmt4 & "  LEFT JOIN ("
    curSqlStmt4 = curSqlStmt4 & "      SELECT MemberID, Div, Reg_Ski, AWSA_Rat as TrickRating, Left(Cast(Cast(RankScore as Decimal(7,1)) as Varchar(8)),6) as TrickRank"
    curSqlStmt4 = curSqlStmt4 & "      FROM " & RankingsTableName
    curSqlStmt4 = curSqlStmt4 & "      WHERE SkiYearID = 1"
    curSqlStmt4 = curSqlStmt4 & "        AND Left(Div,1) in ('B','G','M','W','O')"
    curSqlStmt4 = curSqlStmt4 & "        AND Event = 'T'"
    curSqlStmt4 = curSqlStmt4 & "        AND RankScore is not null)"
    curSqlStmt4 = curSqlStmt4 & "      AS TX ON MX.MemberID = TX.MemberID AND RD.Div = TX.Div"

    '	-----------------------------------------------------------------------
    ' Jump ratings
    '	-----------------------------------------------------------------------
    curSqlStmt4 = curSqlStmt4 & "  LEFT JOIN ("
    curSqlStmt4 = curSqlStmt4 & "      SELECT MemberID, Div, Reg_Ski, AWSA_Rat as JumpRating, Left(Cast(Cast(RankScore as Decimal(6,2)) as Varchar(8)),6) as JumpRank"
    curSqlStmt4 = curSqlStmt4 & "      FROM " & RankingsTableName
    curSqlStmt4 = curSqlStmt4 & "      WHERE SkiYearID = 1"
    curSqlStmt4 = curSqlStmt4 & "        AND Left(Div,1) in ('B','G','M','W','O')"
    curSqlStmt4 = curSqlStmt4 & "        AND Event = 'J'"
    curSqlStmt4 = curSqlStmt4 & "        AND RankScore is not null)"
    curSqlStmt4 = curSqlStmt4 & "      AS JX ON MX.MemberID = JX.MemberID AND RD.Div = JX.Div"

    '	-----------------------------------------------------------------------
    ' Overall ratings
    '	-----------------------------------------------------------------------
    curSqlStmt4 = curSqlStmt4 & "  LEFT JOIN ("
    curSqlStmt4 = curSqlStmt4 & "      SELECT MemberID, Div,  AWSA_Rat as OverallRating, Left(Cast(Cast(RankScore as Decimal(7,1)) as Varchar(8)),6) as OvrSco"
    curSqlStmt4 = curSqlStmt4 & "      FROM " & RankingsTableName
    curSqlStmt4 = curSqlStmt4 & "      WHERE SkiYearID = 1"
    curSqlStmt4 = curSqlStmt4 & "        AND Left(Div,1) in ('B','G','M','W','O')"
    curSqlStmt4 = curSqlStmt4 & "        AND Event = 'O'"
    curSqlStmt4 = curSqlStmt4 & "        AND RankScore is not null)"
    curSqlStmt4 = curSqlStmt4 & "      AS OX ON MX.MemberID = OX.MemberID AND RD.Div = OX.Div"

    '	-----------------------------------------------------------------------
    ' Slalom something to with elite dates
    '	-----------------------------------------------------------------------
    curSqlStmt5 = ""
    curSqlStmt5 = curSqlStmt5 & "  LEFT JOIN ("
    curSqlStmt5 = curSqlStmt5 & "      SELECT MemberID, max(DivElite) as SlmEli"
    curSqlStmt5 = curSqlStmt5 & "      FROM " & EliteDatesTableName & ""
    curSqlStmt5 = curSqlStmt5 & "      WHERE SkiYearID = 1"
    curSqlStmt5 = curSqlStmt5 & "        AND Event = 'S'"
    curSqlStmt5 = curSqlStmt5 & "        AND QualThru >= '" & curTourDate & "'"
    curSqlStmt5 = curSqlStmt5 & "      GROUP BY MemberID)"
    curSqlStmt5 = curSqlStmt5 & "      AS SE ON MX.MemberID = SE.MemberID"

    '	-----------------------------------------------------------------------
    ' Trick something to with elite dates
    '	-----------------------------------------------------------------------
    curSqlStmt5 = curSqlStmt5 & "  LEFT JOIN ("
    curSqlStmt5 = curSqlStmt5 & "      SELECT MemberID, max(DivElite) as TrkEli"
    curSqlStmt5 = curSqlStmt5 & "      FROM " & EliteDatesTableName
    curSqlStmt5 = curSqlStmt5 & "      WHERE SkiYearID = 1"
    curSqlStmt5 = curSqlStmt5 & "        AND Event = 'T'"
    curSqlStmt5 = curSqlStmt5 & "        AND QualThru >= '" & curTourDate & "'"
    curSqlStmt5 = curSqlStmt5 & "      GROUP BY MemberID)"
    curSqlStmt5 = curSqlStmt5 & "      AS TE ON MX.MemberID = TE.MemberID"

    '	-----------------------------------------------------------------------
    ' Jump something to with elite dates
    '	-----------------------------------------------------------------------
    curSqlStmt5 = curSqlStmt5 & "  LEFT JOIN ("
    curSqlStmt5 = curSqlStmt5 & "      SELECT MemberID, max(DivElite) as JmpEli"
    curSqlStmt5 = curSqlStmt5 & "      FROM " & EliteDatesTableName
    curSqlStmt5 = curSqlStmt5 & "      WHERE SkiYearID = 1"
    curSqlStmt5 = curSqlStmt5 & "        AND Event = 'J'"
    curSqlStmt5 = curSqlStmt5 & "        AND QualThru >= '" & curTourDate & "'"
    curSqlStmt5 = curSqlStmt5 & "      GROUP BY MemberID)"
    curSqlStmt5 = curSqlStmt5 & "      AS JE ON MX.MemberID = JE.MemberID"

    '	-----------------------------------------------------------------------
    ' Overall something to with elite dates
    '	----------------------------------------------------------------------
    curSqlStmt5 = curSqlStmt5 & "  LEFT JOIN ("
    curSqlStmt5 = curSqlStmt5 & "      SELECT MemberID, max(DivElite) as OvrEli"
    curSqlStmt5 = curSqlStmt5 & "      FROM " & EliteDatesTableName
    curSqlStmt5 = curSqlStmt5 & "      WHERE SkiYearID = 1"
    curSqlStmt5 = curSqlStmt5 & "        AND Event = 'O'"
    curSqlStmt5 = curSqlStmt5 & "        AND QualThru >= '" & curTourDate & "'"
    curSqlStmt5 = curSqlStmt5 & "      GROUP BY MemberID)"
    curSqlStmt5 = curSqlStmt5 & "      AS OE ON MX.MemberID = OE.MemberID"

    '	-----------------------------------------------------------------------
    ' Retrieve entries from Online Registration
    '	-----------------------------------------------------------------------
    curSqlStmt6 = ""
    curSqlStmt6 = curSqlStmt6 & "  LEFT JOIN ("
    curSqlStmt6 = curSqlStmt6 & "      SELECT MemberID, BibNo, 'YES' as PreReg"
    curSqlStmt6 = curSqlStmt6 & "             , CASE When Len(RampHeight) < 3 Then RampHeight Else left(RampHeight,1) + right(RampHeight,1) END as JRamp"
    curSqlStmt6 = curSqlStmt6 & "      FROM " & RegGenTableName
    curSqlStmt6 = curSqlStmt6 & "      WHERE left(TourID,6) = '" & curSanctionId & "')"
    curSqlStmt6 = curSqlStmt6 & "      AS PR ON MX.MemberID = PR.MemberID"

    curSqlStmt6 = curSqlStmt6 & "  LEFT JOIN ("
    curSqlStmt6 = curSqlStmt6 & "      SELECT MemberID, Div as EventSlalom"
    curSqlStmt6 = curSqlStmt6 & "             , CASE when FeeClass='G' Then 'F' When FeeClass='S' Then 'C' Else FeeClass END as SFeeCls"
    curSqlStmt6 = curSqlStmt6 & "             , right(Cast(FeeRounds as Varchar(3)),1) as SFeeRds"
    curSqlStmt6 = curSqlStmt6 & "             , QfyOverride as SQfyOvr"
    curSqlStmt6 = curSqlStmt6 & "      FROM " & RegEventsTableName
    curSqlStmt6 = curSqlStmt6 & "      Where Left(Event,1) = 'S' AND left(TourID,6) = '" & curSanctionId & "')"
    curSqlStmt6 = curSqlStmt6 & "      AS PS ON MX.MemberID = PS.MemberID"

    curSqlStmt6 = curSqlStmt6 & "  LEFT JOIN ("
    curSqlStmt6 = curSqlStmt6 & "      SELECT MemberID, Div as EventTrick"
    curSqlStmt6 = curSqlStmt6 & "             , CASE when FeeClass='G' Then 'F' When FeeClass='S' Then 'C' Else FeeClass END as TFeeCls"
    curSqlStmt6 = curSqlStmt6 & "             , right(Cast(FeeRounds as Varchar(3)),1) as TFeeRds"
    curSqlStmt6 = curSqlStmt6 & "             , QfyOverride as TQfyOvr, Boat as TrickBoat"
    curSqlStmt6 = curSqlStmt6 & "      FROM " & RegEventsTableName
    curSqlStmt6 = curSqlStmt6 & "      WHERE Left(Event,1) = 'T'"
    curSqlStmt6 = curSqlStmt6 & "        AND left(TourID,6) = '" & curSanctionId & "')"
    curSqlStmt6 = curSqlStmt6 & "      AS PT ON MX.MemberID = PT.MemberID"

    curSqlStmt6 = curSqlStmt6 & "  LEFT JOIN ("
    curSqlStmt6 = curSqlStmt6 & "      SELECT MemberID, Div as EventJump"
    curSqlStmt6 = curSqlStmt6 & "             , CASE when FeeClass='G' Then 'F' When FeeClass='S' Then 'C' Else FeeClass END as JFeeCls"
    curSqlStmt6 = curSqlStmt6 & "             , right(Cast(FeeRounds as Varchar(3)),1) as JFeeRds"
    curSqlStmt6 = curSqlStmt6 & "             , QfyOverride as JQfyOvr"
    curSqlStmt6 = curSqlStmt6 & "      FROM " & RegEventsTableName
    curSqlStmt6 = curSqlStmt6 & "      WHERE Left(Event,1) = 'J'"
    curSqlStmt6 = curSqlStmt6 & "        AND left(TourID,6) = '" & curSanctionId & "')"
    curSqlStmt6 = curSqlStmt6 & "      AS PJ ON MX.MemberID = PJ.MemberID"

    '	-----------------------------------------------------------------------
    ' Retrieve qualifications
    '	-----------------------------------------------------------------------
    curSqlStmt7 = ""
    curSqlStmt7 = curSqlStmt7 & "  LEFT JOIN ("
    curSqlStmt7 = curSqlStmt7 & "      SELECT MemberID, Div as EventSlalom"
    curSqlStmt7 = curSqlStmt7 & "             , CASE When QfyStatus = 'Qualified' Then 'Y' Else ' ' END as SQfy"
    curSqlStmt7 = curSqlStmt7 & "      FROM " & RegQualifyTableName
    curSqlStmt7 = curSqlStmt7 & "      WHERE Left(Event,1) = 'S'"
    curSqlStmt7 = curSqlStmt7 & "        AND left(TourID,6) = '" & curSanctionId & "')"
    curSqlStmt7 = curSqlStmt7 & "      AS QS ON PS.MemberID = QS.MemberID AND PS.EventSlalom = QS.EventSlalom"

    curSqlStmt7 = curSqlStmt7 & "  LEFT JOIN ("
    curSqlStmt7 = curSqlStmt7 & "      SELECT MemberID, Div as EventTrick"
    curSqlStmt7 = curSqlStmt7 & "             , CASE When QfyStatus = 'Qualified' Then 'Y' Else ' ' END as TQfy"
    curSqlStmt7 = curSqlStmt7 & "      FROM " & RegQualifyTableName
    curSqlStmt7 = curSqlStmt7 & "      WHERE Left(Event,1) = 'T'"
    curSqlStmt7 = curSqlStmt7 & "        AND left(TourID,6) = '" & curSanctionId & "')"
    curSqlStmt7 = curSqlStmt7 & "      AS QT ON PT.MemberID = QT.MemberID AND PT.EventTrick = QT.EventTrick"

    curSqlStmt7 = curSqlStmt7 & "  LEFT JOIN ("
    curSqlStmt7 = curSqlStmt7 & "      SELECT MemberID, Div as EventJump"
    curSqlStmt7 = curSqlStmt7 & "             , CASE When QfyStatus = 'Qualified' Then 'Y' Else ' ' END as JQfy"
    curSqlStmt7 = curSqlStmt7 & "      FROM " & RegQualifyTableName
    curSqlStmt7 = curSqlStmt7 & "      WHERE Left(Event,1) = 'J'"
    curSqlStmt7 = curSqlStmt7 & "        AND left(TourID,6) = '" & curSanctionId & "')"
    curSqlStmt7 = curSqlStmt7 & "      AS QJ ON PJ.MemberID = QJ.MemberID AND PJ.EventJump = QJ.EventJump "

    '	-----------------------------------------------------------------------
    ' Retrieve officials ratings
    '	-----------------------------------------------------------------------
    curSqlStmt7 = curSqlStmt7 & "  LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
    curSqlStmt7 = curSqlStmt7 & "     		FROM " & OfficialsTableName & " as OT"
    curSqlStmt7 = curSqlStmt7 & "     		    INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
    curSqlStmt7 = curSqlStmt7 & "    		WHERE OT.RatingType_ID = 1 AND OT.EventsConsolidated like '%s%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL AND LV.LevelOrderforTemplate < 5"
    curSqlStmt7 = curSqlStmt7 & "         ) as OJS ON OJS.PersonID = MX.PersonID"
    curSqlStmt7 = curSqlStmt7 & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
    curSqlStmt7 = curSqlStmt7 & "			FROM " & OfficialsTableName & " as OT"
    curSqlStmt7 = curSqlStmt7 & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
    curSqlStmt7 = curSqlStmt7 & "			WHERE OT.RatingType_ID = 1 AND OT.EventsConsolidated like '%t%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL AND LV.LevelOrderforTemplate < 5"
    curSqlStmt7 = curSqlStmt7 & "			) as OJT ON OJT.PersonID = MX.PersonID"
    curSqlStmt7 = curSqlStmt7 & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
    curSqlStmt7 = curSqlStmt7 & "			FROM " & OfficialsTableName & " as OT"
    curSqlStmt7 = curSqlStmt7 & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
    curSqlStmt7 = curSqlStmt7 & "			WHERE OT.RatingType_ID = 1 AND OT.EventsConsolidated like '%j%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL AND LV.LevelOrderforTemplate < 5"
    curSqlStmt7 = curSqlStmt7 & "			) as OJJ ON OJJ.PersonID = MX.PersonID"

    curSqlStmt7 = curSqlStmt7 & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Abbreviation as Rating"
    curSqlStmt7 = curSqlStmt7 & "			FROM " & OfficialsTableName & " as OT"
    curSqlStmt7 = curSqlStmt7 & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
    curSqlStmt7 = curSqlStmt7 & "			WHERE OT.RatingType_ID = 1 AND OT.EventsConsolidated like '%s%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate = 5"
    curSqlStmt7 = curSqlStmt7 & "			) as OPS ON OPS.PersonID = MX.PersonID"
    curSqlStmt7 = curSqlStmt7 & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Abbreviation as Rating"
    curSqlStmt7 = curSqlStmt7 & "			FROM " & OfficialsTableName & " as OT"
    curSqlStmt7 = curSqlStmt7 & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
    curSqlStmt7 = curSqlStmt7 & "			WHERE OT.RatingType_ID = 1 AND OT.EventsConsolidated like '%t%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate = 5"
    curSqlStmt7 = curSqlStmt7 & "			) as OPT ON OPT.PersonID = MX.PersonID"
    curSqlStmt7 = curSqlStmt7 & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Abbreviation as Rating"
    curSqlStmt7 = curSqlStmt7 & "			FROM " & OfficialsTableName & " as OT"
    curSqlStmt7 = curSqlStmt7 & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
    curSqlStmt7 = curSqlStmt7 & "			WHERE OT.RatingType_ID = 1 AND OT.EventsConsolidated like '%j%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate = 5"
    curSqlStmt7 = curSqlStmt7 & "			) as OPJ ON OPJ.PersonID = MX.PersonID"

    curSqlStmt7 = curSqlStmt7 & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
    curSqlStmt7 = curSqlStmt7 & "			FROM " & OfficialsTableName & " as OT"
    curSqlStmt7 = curSqlStmt7 & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
    curSqlStmt7 = curSqlStmt7 & "			WHERE OT.RatingType_ID = 3 AND OT.EventsConsolidated like '%s%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL"
    curSqlStmt7 = curSqlStmt7 & "			) as ODS ON ODS.PersonID = MX.PersonID"
    curSqlStmt7 = curSqlStmt7 & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
    curSqlStmt7 = curSqlStmt7 & "			FROM " & OfficialsTableName & " as OT"
    curSqlStmt7 = curSqlStmt7 & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
    curSqlStmt7 = curSqlStmt7 & "			WHERE OT.RatingType_ID = 3 AND OT.EventsConsolidated like '%t%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL"
    curSqlStmt7 = curSqlStmt7 & "			) as ODT ON ODT.PersonID = MX.PersonID"
    curSqlStmt7 = curSqlStmt7 & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
    curSqlStmt7 = curSqlStmt7 & "			FROM " & OfficialsTableName & " as OT"
    curSqlStmt7 = curSqlStmt7 & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
    curSqlStmt7 = curSqlStmt7 & "			WHERE OT.RatingType_ID = 3 AND OT.EventsConsolidated like '%j%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL"
    curSqlStmt7 = curSqlStmt7 & "			) as ODJ ON ODJ.PersonID = MX.PersonID"

    curSqlStmt7 = curSqlStmt7 & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
    curSqlStmt7 = curSqlStmt7 & "			FROM " & OfficialsTableName & " as OT"
    curSqlStmt7 = curSqlStmt7 & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
    curSqlStmt7 = curSqlStmt7 & "			WHERE OT.RatingType_ID = 2 AND OT.EventsConsolidated like '%s%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL"
    curSqlStmt7 = curSqlStmt7 & "			) as OCS ON OCS.PersonID = MX.PersonID"
    curSqlStmt7 = curSqlStmt7 & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
    curSqlStmt7 = curSqlStmt7 & "			FROM " & OfficialsTableName & " as OT"
    curSqlStmt7 = curSqlStmt7 & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
    curSqlStmt7 = curSqlStmt7 & "			WHERE OT.RatingType_ID = 2 AND OT.EventsConsolidated like '%t%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL"
    curSqlStmt7 = curSqlStmt7 & "			) as OCT ON OCT.PersonID = MX.PersonID"
    curSqlStmt7 = curSqlStmt7 & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
    curSqlStmt7 = curSqlStmt7 & "			FROM " & OfficialsTableName & " as OT"
    curSqlStmt7 = curSqlStmt7 & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
    curSqlStmt7 = curSqlStmt7 & "			WHERE OT.RatingType_ID = 2 AND OT.EventsConsolidated like '%j%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL"
    curSqlStmt7 = curSqlStmt7 & "			) as OCJ ON OCJ.PersonID = MX.PersonID"

    curSqlStmt7 = curSqlStmt7 & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
    curSqlStmt7 = curSqlStmt7 & "			FROM " & OfficialsTableName & " as OT"
    curSqlStmt7 = curSqlStmt7 & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
    curSqlStmt7 = curSqlStmt7 & "			WHERE OT.RatingType_ID = 9 AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL"
    curSqlStmt7 = curSqlStmt7 & "			) as OS ON OS.PersonID = MX.PersonID"

    curSqlStmt7 = curSqlStmt7 & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
    curSqlStmt7 = curSqlStmt7 & "			FROM " & OfficialsTableName & " as OT"
    curSqlStmt7 = curSqlStmt7 & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
    curSqlStmt7 = curSqlStmt7 & "			WHERE OT.RatingType_ID = 4 AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL"
    curSqlStmt7 = curSqlStmt7 & "			) as OTC ON OTC.PersonID = MX.PersonID"

    curSqlStmt8 = " "
    curSqlStmt8 = curSqlStmt8 & "WHERE MX.ExportToTemplate = 1"
    curSqlStmt8 = curSqlStmt8 & "  AND MX.Deceased = 0 "

    IF len(curMemberId) = 0 AND len(curMemberFirstName) = 0  AND len(curMemberLastName) = 0 THEN
        curSqlStmt8 = curSqlStmt8 & "  AND DateAdd(mm,18,MX.EffTo) > GetDate()"
        curSqlStmt8 = curSqlStmt8 & "  AND ("
	    curSqlStmt8 = curSqlStmt8 & "   MX.PersonID in (Select PersonID from " & ApptOfficialsTableName & " WHERE TournAppID = '" & curSanctionId & "' ) "
	    curSqlStmt8 = curSqlStmt8 & "   OR MX.MemberID in (Select MemberID from " & RegEventsTableName & " WHERE left(TourID,6) = '" & curSanctionId & "') "
        IF len(curStateSQL) > 0 THEN
            curSqlStmt8 = curSqlStmt8 & "  OR " & curStateSQL
        END IF
        curSqlStmt8 = curSqlStmt8 & " ) "
    ELSE
        curSqlStmt8 = curSqlStmt8 & "  AND DateAdd(mm,30,MX.EffTo) > GetDate()"
        
        IF len(curMemberId) > 0 THEN
            curSqlStmt8 = curSqlStmt8 & "  AND MX.MemberID = '" & curMemberId & "' "
        ELSE
            IF len(curMemberFirstName) > 0 OR len(curMemberLastName) > 0 THEN
                curSqlStmt8 = curSqlStmt8 & "  AND  MX.FirstName like '" & curMemberFirstName & "%' "
                curSqlStmt8 = curSqlStmt8 & "  AND MX.LastName like '" & curMemberLastName & "%' "
                IF len(curStateSQL) > 0 THEN
                    curSqlStmt8 = curSqlStmt8 & "  AND " & curStateSQL
                END IF
            ELSEIF len(curStateSQL) > 0 THEN
                    curSqlStmt8 = curSqlStmt8 & "  AND " & curStateSQL
            END IF
        END IF

    END IF


    '	-----------------------------------------------------------------------
    ' Order by statement
    '	-----------------------------------------------------------------------
    curSqlStmt8 = curSqlStmt8 & " Order By MX.LastName, MX.FirstName, RD.MemberID, RD.Div"

	''''response.write curSqlStmt1 & curSqlStmt2 & curSqlStmt3 & curSqlStmt4 & curSqlStmt5 & curSqlStmt6 & curSqlStmt7 & curSqlStmt8
    ''''response.End

    '	-----------------------------------------------------------------------
    ' Execute SQL statement to retrieve skier information and load to registration template
    '	-----------------------------------------------------------------------
    buildQueryMemberRegEntries = curSqlStmt1 & curSqlStmt2 & curSqlStmt3 & curSqlStmt4 & curSqlStmt5 & curSqlStmt6 & curSqlStmt7 & curSqlStmt8

End Function

'	-----------------------------------------------------------------------
' Build a query to extract member entries for tournament registrations
' Include data from rankings, qualifications, membership status, and official ratings
'	-----------------------------------------------------------------------
Function buildQueryMemberRankingEquivalents(curSanctionId, curTourDate, curStateSQL, curMemberId, curMemberFirstName, curMemberLastName)
    Dim curTourYear, curSqlStmt1, curSqlStmt2, curSqlStmt3, curSqlStmt4
    curTourYear = 2000 + left(curSanctionId, 2)

    'Member Number and name
    curSqlStmt1 = ""
    curSqlStmt1 = curSqlStmt1 & "SELECT MX.PersonIDWithCheckDigit as MemberID, MX.PersonID, MX.LastName, MX.FirstName, MX.FederationCode as Federation" 
    curSqlStmt1 = curSqlStmt1 & ", (" & curTourYear & " - Year(MX.BirthDate) - 1) as Age"
    curSqlStmt1 = curSqlStmt1 & ", Upper(Left(MX.Sex,1)) as Sex"
    curSqlStmt1 = curSqlStmt1 & ", MX.City, MX.State"
    curSqlStmt1 = curSqlStmt1 & ", MX.MembershipTypeCode as MemType"
	
    curSqlStmt1 = curSqlStmt1 & ", MX.EffectiveTo as EffTo"
    curSqlStmt1 = curSqlStmt1 & ", Typ.ActiveMember"
    curSqlStmt1 = curSqlStmt1 & ", Typ.Description as MemTypeDesc"
    curSqlStmt1 = curSqlStmt1 & ", Typ.CanSkiInTournaments as CanSki"

    curSqlStmt1 = curSqlStmt1 & ", SX.Div, SX.AWSA_Rat as SlalomRating"
    curSqlStmt1 = curSqlStmt1 & ", Cast(SX.RankScore as Decimal(7,2)) as SlalomRank"
    curSqlStmt1 = curSqlStmt1 & ", SX.Top_Equiv_SC1 as SlalomRankEquiv1"
    curSqlStmt1 = curSqlStmt1 & ", SX.Top_Equiv_SC2 as SlalomRankEquiv2"
    curSqlStmt1 = curSqlStmt1 & ", SX.Top_Equiv_SC3 as SlalomRankEquiv3"

    curSqlStmt1 = curSqlStmt1 & ", TX.AWSA_Rat as TrickRating"
    curSqlStmt1 = curSqlStmt1 & ", Cast(TX.RankScore as Decimal(7,2)) as TrickRank"
    curSqlStmt1 = curSqlStmt1 & ", TX.Top_Equiv_SC1 as TrickRankEquiv1"
    curSqlStmt1 = curSqlStmt1 & ", TX.Top_Equiv_SC2 as TrickRankEquiv2"
    curSqlStmt1 = curSqlStmt1 & ", TX.Top_Equiv_SC3 as TrickRankEquiv3"

    curSqlStmt1 = curSqlStmt1 & ", JX.AWSA_Rat as JumpRating"
    curSqlStmt1 = curSqlStmt1 & ", Cast(JX.RankScore as Decimal(7,2)) as JumpRank"
    curSqlStmt1 = curSqlStmt1 & ", JX.Top_Equiv_SC1 as JumpRankEquiv1"
    curSqlStmt1 = curSqlStmt1 & ", JX.Top_Equiv_SC2 as JumpRankEquiv2"
    curSqlStmt1 = curSqlStmt1 & ", JX.Top_Equiv_SC3 as JumpRankEquiv3"

    curSqlStmt1 = curSqlStmt1 & ", OX.AWSA_Rat as OverallRating"
    curSqlStmt1 = curSqlStmt1 & ", Cast(OX.RankScore as Decimal(7,2)) as OverallRank"

    'From Tables
    curSqlStmt2 = " "
    curSqlStmt2 = curSqlStmt2 & "FROM " & MemberTableName & " AS MX "
    curSqlStmt2 = curSqlStmt2 & "INNER JOIN " & MembershipTypesTableName & " AS Typ ON Typ.MemberShipTypeID = MX.MembershipTypeCode "

    curSqlStmt2 = curSqlStmt2 & "LEFT JOIN " & RankingsTableName & " AS SX "
    curSqlStmt2 = curSqlStmt2 & "	ON SX.MemberID = MX.PersonIDWithCheckDigit AND SX.SkiYearID = 1 AND SX.Event = 'S' AND Left(SX.Div,1) in ('B','G','M','W','O') AND SX.RankScore is not null "
    curSqlStmt2 = curSqlStmt2 & "LEFT JOIN " & RankingsTableName & " AS TX "
    curSqlStmt2 = curSqlStmt2 & "	ON TX.MemberID = MX.PersonIDWithCheckDigit AND TX.SkiYearID = 1 AND TX.Event = 'T' AND Left(TX.Div,1) in ('B','G','M','W','O') AND TX.RankScore is not null "
    curSqlStmt2 = curSqlStmt2 & "LEFT JOIN " & RankingsTableName & " AS JX "
    curSqlStmt2 = curSqlStmt2 & "	ON JX.MemberID = MX.PersonIDWithCheckDigit AND JX.SkiYearID = 1 AND JX.Event = 'J' AND Left(JX.Div,1) in ('B','G','M','W','O') AND JX.RankScore is not null "
    curSqlStmt2 = curSqlStmt2 & "LEFT JOIN " & RankingsTableName & " AS OX "
    curSqlStmt2 = curSqlStmt2 & "	ON OX.MemberID = MX.PersonIDWithCheckDigit AND OX.SkiYearID = 1 AND OX.Event = 'O' AND Left(OX.Div,1) in ('B','G','M','W','O') AND OX.RankScore is not null "

    'Where clause
    curSqlStmt3 = " "
    curSqlStmt3 = curSqlStmt3 & "WHERE Typ.ExporttoTouramentRegistrationTemplate = 1"
    curSqlStmt3 = curSqlStmt3 & "  AND MX.Deceased = 0 "
    curSqlStmt3 = curSqlStmt3 & "  AND SX.Div is not null"

    IF len(curMemberId) = 0 AND len(curMemberFirstName) = 0  AND len(curMemberLastName) = 0 THEN
        curSqlStmt3 = curSqlStmt3 & "  AND DateAdd(mm, 18, MX.EffectiveTo) > GetDate()"
        IF len(curStateSQL) > 0 THEN
            curSqlStmt3 = curSqlStmt3 & "  AND " & curStateSQL
        END IF
    ELSE
        curSqlStmt3 = curSqlStmt3 & "  AND DateAdd(mm, 18, MX.EffectiveTo) > GetDate()"
        
        IF len(curMemberId) > 0 THEN
            curSqlStmt3 = curSqlStmt3 & "  AND MX.PersonIDWithCheckDigit = '" & curMemberId & "' "
        ELSE
            IF len(curMemberFirstName) > 0 OR len(curMemberLastName) > 0 THEN
                curSqlStmt3 = curSqlStmt3 & "  AND  MX.FirstName like '" & curMemberFirstName & "%' "
                curSqlStmt3 = curSqlStmt3 & "  AND MX.LastName like '" & curMemberLastName & "%' "
                IF len(curStateSQL) > 0 THEN
                    curSqlStmt3 = curSqlStmt3 & "  AND " & curStateSQL
                END IF
            ELSEIF len(curStateSQL) > 0 THEN
                    curSqlStmt3 = curSqlStmt3 & "  AND " & curStateSQL
            END IF
        END IF

    END IF

    '	-----------------------------------------------------------------------
    ' Order by statement
    '	-----------------------------------------------------------------------
    curSqlStmt4 = " Order By MX.LastName, MX.FirstName, MX.PersonIDWithCheckDigit, SX.Div"

    '	-----------------------------------------------------------------------
    ' Execute SQL statement to retrieve skier information and load to registration template
    '	-----------------------------------------------------------------------
    buildQueryMemberRankingEquivalents = curSqlStmt1 & curSqlStmt2 & curSqlStmt3 & curSqlStmt4 

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Now build a Query to Extract the Desired Members, joining in data 
''' pulled from the Rankings and Officials and Membership Type tables.
''' Note that we prefix each team ID with "E" if the team has entries,
''' or "Z" if no entries, so that all the entered teams list at the top,
''' then finally all those without any team affiliation last with Zzzz.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function buildQueryMemberRegNcwsaEntries(curSanctionId, curTourDate)
    Dim curTourYear, curSqlStmt
    curTourYear = 2000 + left(curSanctionId, 2)

    'Member Number and name
    curSqlStmt = ""
    curSqlStmt = curSqlStmt & "Select Substring(MX.MemberID,1,3) + '-' + Substring(MX.MemberID,4,2) + '-' + Substring(MX.MemberID,6,4) as MemberID"
    curSqlStmt = curSqlStmt & ", MX.LastName, MX.FirstName"

    'Skier division
    curSqlStmt = curSqlStmt & ", Case when MX.Sex = 'F' Then 'CW' else 'CM' END as Div"
    
    curSqlStmt = curSqlStmt & ", Case when MX.Age <= 17 and MX.Sex = 'F' Then 'G'"
    curSqlStmt = curSqlStmt & "       when MX.Age <= 17 then 'B'"
    curSqlStmt = curSqlStmt & "       when MX.Sex = 'F' then 'W'"
    curSqlStmt = curSqlStmt & "       ELSE 'M' END"
    curSqlStmt = curSqlStmt & "   + Case"
    curSqlStmt = curSqlStmt & "          when MX.Age <= 9 then '1'"
    curSqlStmt = curSqlStmt & "          when MX.Age <= 13 then '2'"
    curSqlStmt = curSqlStmt & "          when MX.Age <= 17 then '3'"
    curSqlStmt = curSqlStmt & "          when MX.Age <= 24 then '1'"
    curSqlStmt = curSqlStmt & "          when MX.Age <= 34 then '2'"
    curSqlStmt = curSqlStmt & "          when MX.Age <= 44 then '3'"
    curSqlStmt = curSqlStmt & "          when MX.Age <= 54 then '4'"
    curSqlStmt = curSqlStmt & "          when MX.Age <= 59 then '5'"
    curSqlStmt = curSqlStmt & "          when MX.Age <= 64 then '6'"
    curSqlStmt = curSqlStmt & "          when MX.Age <= 69 then '7'"
    curSqlStmt = curSqlStmt & "          when MX.Age <= 74 then '8'"
    curSqlStmt = curSqlStmt & "          when MX.Age <= 79 then '9'"
    curSqlStmt = curSqlStmt & "          when MX.Age <= 84 then 'A'"
    curSqlStmt = curSqlStmt & "          ELSE 'B' END as AgeDiv"

    'Skier information
    curSqlStmt = curSqlStmt & ", MX.Team, MX.TeamName, MX.TeamStat, MX.TeamTournAppID, MX.Age, MX.Sex as Gender, MX.City, MX.State, Coalesce(MX.Federation, '') as Federation, MX.Waiver"

    'Skier official ratings
    curSqlStmt = curSqlStmt & ", Coalesce(SO.OffCode,'') as ApptdOfficial"

    'Sort attribute
    curSqlStmt = curSqlStmt & ", Case when SO.OffCode is not Null AND MX.EventSlalom + MX.EventTrick + MX.EventJump = '      ' then 'E0FF' else MX.Sorter end as Sorter"

    'Skier event attributes and payments
    curSqlStmt = curSqlStmt & ", MX.EventWaiver, MX.EventSlalom, MX.EventTrick, MX.EventJump, MX.TrickBoat, MX.JumpHeight"

    'Other member stuff
    curSqlStmt = curSqlStmt & ", MX.EffTo, MX.Memtype, MX.MemCode, MX.ActiveMember, MX.MemTypeDesc, MX.CanSki, MX.CanSkiGR, MX.SptsDiv, MembershipRate, CostToUpgrade"

    curSqlStmt = curSqlStmt & ", Case WHEN OPS.Rating = 'INT' THEN 'PanAm' ELSE Coalesce(OJS.Rating, '') END as JudgeSlalom"
    curSqlStmt = curSqlStmt & ", Case WHEN OPT.Rating = 'INT' THEN 'PanAm' ELSE Coalesce(OJT.Rating, '') END as JudgeTrick"
    curSqlStmt = curSqlStmt & ", Case WHEN OPJ.Rating = 'INT' THEN 'PanAm' ELSE Coalesce(OJJ.Rating, '') END as JudgeJump"
    curSqlStmt = curSqlStmt & ", Coalesce(ODS.Rating, '') as DriverSlalom, Coalesce(ODT.Rating, '') as DriverTrick, Coalesce(ODJ.Rating, '') as DriverJump"
    curSqlStmt = curSqlStmt & ", Coalesce(OCS.Rating, '') as ScorerSlalom, Coalesce(OCT.Rating, '') as ScorerTrick, Coalesce(OCJ.Rating, '') as ScorerJump"
    curSqlStmt = curSqlStmt & ", Coalesce(OS.Rating, '') as Safety, Coalesce(OTC.Rating, '') as TechController "

    '	-----------------------------------------------------------------------
    'FROM Statement
    '	-----------------------------------------------------------------------
    curSqlStmt = curSqlStmt & " FROM ("

    '	-----------------------------------------------------------------------
    'Use select as a data source for member data
    '	-----------------------------------------------------------------------
    curSqlStmt = curSqlStmt & "    SELECT MT.PersonIDWithCheckDigit as MemberID, MT.PersonID, MT.LastName, FirstName, MT.FederationCode as Federation"
    curSqlStmt = curSqlStmt & "        , (" & curTourYear & " - Year(MT.BirthDate) - 1) as Age, Upper(Left(MT.Sex,1)) as Sex, MT.WaiverStatusID as Waiver"
    curSqlStmt = curSqlStmt & "        , MT.City, Left(MT.State,2) as State"
    curSqlStmt = curSqlStmt & "        , MT.EffectiveTo as EffTo, MT.MembershipTypeCode as MemType"
    curSqlStmt = curSqlStmt & "        , MT.Deceased, MT.DivisionCode1 + '/' + MT.DivisionCode2 as SptsDiv"
    curSqlStmt = curSqlStmt & "        , Typ.ExporttoTouramentRegistrationTemplate as ExportToTemplate"
    curSqlStmt = curSqlStmt & "        , Typ.TypeCode as MemCode, Typ.ActiveMember, Typ.Description as MemTypeDesc, Typ.CanSkiInTournaments as CanSki, Typ.CanSkiInGRTournaments as CanSkiGR"
    curSqlStmt = curSqlStmt & "        , Coalesce(MR.MembershipTypeRates, 0) as MembershipRate, Coalesce(MR.CosttoUpgrade, 0) as CostToUpgrade"

    curSqlStmt = curSqlStmt & "        , CASE WHEN TE.Team is not null THEN 'E' ELSE 'Z' END + CASE WHEN Coalesce(RP.Team, TR.Team) is not null THEN Coalesce(RP.Team, TR.Team) ELSE 'zzz' END as Sorter"
    curSqlStmt = curSqlStmt & "        , CASE WHEN RP.MemberID is not null THEN 'A' WHEN TR.DateInactive is not null THEN 'I' ELSE 'A' END as TeamStat"
    curSqlStmt = curSqlStmt & "        , Coalesce(RP.Team,TR.Team,'   ') as Team, Coalesce(TR.TeamName, '') as TeamName, Coalesce(RP.TournAppID, '') as TeamTournAppID"

    curSqlStmt = curSqlStmt & "        , Coalesce(CASE WHEN right(RP.SlalomEnt,1) <= '9' THEN RP.SlalomEnt ELSE left(RP.SlalomEnt,1) + cast(ascii(right(RP.SlalomEnt,1)) - 55 as varchar(2)) END, '  ') as EventSlalom" 
    curSqlStmt = curSqlStmt & "        , Coalesce(CASE WHEN right(RP.TrickEnt,1) <= '9' THEN RP.TRickEnt ELSE left(RP.TrickEnt,1) + cast(ascii(right(RP.TrickEnt,1)) - 55 as varchar(2)) END, '  ') as EventTrick" 
    curSqlStmt = curSqlStmt & "        , Coalesce(CASE WHEN right(RP.JumpEnt,1) <= '9' THEN RP.JumpEnt ELSE left(RP.JumpEnt,1) + cast(ascii(right(RP.JumpEnt,1)) - 55 as varchar(2)) END, '  ') as EventJump" 

    curSqlStmt = curSqlStmt & "        , Coalesce(RP.WaiverStat,' ') as EventWaiver" 
    curSqlStmt = curSqlStmt & "        , Coalesce(RP.TrickBoat,'  ') as TrickBoat" 
    curSqlStmt = curSqlStmt & "        , Coalesce(RP.RampHgt,'  ') as JumpHeight" 

    curSqlStmt = curSqlStmt & "    FROM " & MemberTableName & " as MT "
    curSqlStmt = curSqlStmt & "      INNER JOIN " & MembershipTypesTableName & " as Typ ON MT.MembershipTypeCode = Typ.MemberShipTypeID "
    curSqlStmt = curSqlStmt & "      LEFT JOIN " & MembershipRatesTableName & " as MR ON MR.[Membership Type Code] = MT.MembershipTypeCode "
    curSqlStmt = curSqlStmt & "           AND MR.EffectiveFrom <= CONVERT(DATETIME, '" & curTourDate & " 00:00:00', 102)"
    curSqlStmt = curSqlStmt & "           AND MR.EffectiveTo >= CONVERT(DATETIME, '" & curTourDate & " 00:00:00', 102)"

                            '	Subquery to retrieve Team ID's from the Team Roster Extract and identify Latest Team affiliation for Member
    curSqlStmt = curSqlStmt & "      LEFT JOIN ("
    curSqlStmt = curSqlStmt & "           SELECT RX.MemberID, RX.Team, TL.TeamName, RX.DateInactive "
    curSqlStmt = curSqlStmt & "           FROM " & TeamRosterTableName & " as RX"
    curSqlStmt = curSqlStmt & "             INNER JOIN " & TeamTableName & " as TL ON TL.TeamId = RX.Team AND SptsGrpId = 'NCW'"
    curSqlStmt = curSqlStmt & "             INNER JOIN (SELECT MemberID, Max(LastEvent) as MaxEvt FROM " & TeamRosterTableName & " Group By MemberID" 
    curSqlStmt = curSqlStmt & "                   ) as ME ON ME.MemberID = RX.MemberID and ME.MaxEvt = RX.LastEvent"
    curSqlStmt = curSqlStmt & "                ) as TR ON TR.MemberID = MT.PersonIDWithCheckDigit"

                            '	This subquery pulls Rotation Plan information for this Person/TourID -- LEAVE TEAM OUT !! (All Stars)
    curSqlStmt = curSqlStmt & "      LEFT JOIN "
    curSqlStmt = curSqlStmt & "          " & TeamRotationsTableName & " as RP ON RP.TournAppID = '" & curSanctionId & "' AND RP.MemberID = MT.PersonIDWithCheckDigit"

                            '	This subquery identifies Teams that are Entered, used to preface Sorter extract column
    curSqlStmt = curSqlStmt & "      LEFT JOIN ("
    curSqlStmt = curSqlStmt & "           Select distinct team FROM " & TeamRotationsTableName & " WHERE WaiverStat >= 'C' and TournAppID = '" & curSanctionId & "') as TE"
    curSqlStmt = curSqlStmt & "           ON TE.Team = Coalesce(RP.Team,TR.Team )"

    ' -----------------------------------------------
    '	End of MX Primary "MX" Select Subquery.  Appended Info Subqueries follow.
    ' -----------------------------------------------
    curSqlStmt = curSqlStmt & "    ) as MX"

    '	-----------------------------------------------------------------------
    ' Use select as a data source for chief officials
    '	-----------------------------------------------------------------------
    curSqlStmt = curSqlStmt & "     LEFT JOIN " & ApptOfficialsTableName & " AS SO ON SO.PersonID = MX.PersonID AND TournAppID = '" & curSanctionId & "' "

    '	-----------------------------------------------------------------------
    ' Retrieve officials ratings
    '	-----------------------------------------------------------------------
    curSqlStmt = curSqlStmt & " LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
    curSqlStmt = curSqlStmt & "     	FROM " & OfficialsTableName & " as OT"
    curSqlStmt = curSqlStmt & "     		    INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
    curSqlStmt = curSqlStmt & "    		WHERE OT.RatingType_ID = 1 AND OT.EventsConsolidated like '%s%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL AND LV.LevelOrderforTemplate < 5"
    curSqlStmt = curSqlStmt & "         ) as OJS ON OJS.PersonID = MX.PersonID"
    curSqlStmt = curSqlStmt & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
    curSqlStmt = curSqlStmt & "			FROM " & OfficialsTableName & " as OT"
    curSqlStmt = curSqlStmt & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
    curSqlStmt = curSqlStmt & "			WHERE OT.RatingType_ID = 1 AND OT.EventsConsolidated like '%t%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL AND LV.LevelOrderforTemplate < 5"
    curSqlStmt = curSqlStmt & "			) as OJT ON OJT.PersonID = MX.PersonID"
    curSqlStmt = curSqlStmt & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
    curSqlStmt = curSqlStmt & "			FROM " & OfficialsTableName & " as OT"
    curSqlStmt = curSqlStmt & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
    curSqlStmt = curSqlStmt & "			WHERE OT.RatingType_ID = 1 AND OT.EventsConsolidated like '%j%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL AND LV.LevelOrderforTemplate < 5"
    curSqlStmt = curSqlStmt & "			) as OJJ ON OJJ.PersonID = MX.PersonID"

    curSqlStmt = curSqlStmt & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Abbreviation as Rating"
    curSqlStmt = curSqlStmt & "			FROM " & OfficialsTableName & " as OT"
    curSqlStmt = curSqlStmt & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
    curSqlStmt = curSqlStmt & "			WHERE OT.RatingType_ID = 1 AND OT.EventsConsolidated like '%s%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate = 5"
    curSqlStmt = curSqlStmt & "			) as OPS ON OPS.PersonID = MX.PersonID"
    curSqlStmt = curSqlStmt & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Abbreviation as Rating"
    curSqlStmt = curSqlStmt & "			FROM " & OfficialsTableName & " as OT"
    curSqlStmt = curSqlStmt & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
    curSqlStmt = curSqlStmt & "			WHERE OT.RatingType_ID = 1 AND OT.EventsConsolidated like '%t%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate = 5"
    curSqlStmt = curSqlStmt & "			) as OPT ON OPT.PersonID = MX.PersonID"
    curSqlStmt = curSqlStmt & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Abbreviation as Rating"
    curSqlStmt = curSqlStmt & "			FROM " & OfficialsTableName & " as OT"
    curSqlStmt = curSqlStmt & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
    curSqlStmt = curSqlStmt & "			WHERE OT.RatingType_ID = 1 AND OT.EventsConsolidated like '%j%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate = 5"
    curSqlStmt = curSqlStmt & "			) as OPJ ON OPJ.PersonID = MX.PersonID"

    curSqlStmt = curSqlStmt & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
    curSqlStmt = curSqlStmt & "			FROM " & OfficialsTableName & " as OT"
    curSqlStmt = curSqlStmt & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
    curSqlStmt = curSqlStmt & "			WHERE OT.RatingType_ID = 3 AND OT.EventsConsolidated like '%s%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL"
    curSqlStmt = curSqlStmt & "			) as ODS ON ODS.PersonID = MX.PersonID"
    curSqlStmt = curSqlStmt & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
    curSqlStmt = curSqlStmt & "			FROM " & OfficialsTableName & " as OT"
    curSqlStmt = curSqlStmt & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
    curSqlStmt = curSqlStmt & "			WHERE OT.RatingType_ID = 3 AND OT.EventsConsolidated like '%t%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL"
    curSqlStmt = curSqlStmt & "			) as ODT ON ODT.PersonID = MX.PersonID"
    curSqlStmt = curSqlStmt & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
    curSqlStmt = curSqlStmt & "			FROM " & OfficialsTableName & " as OT"
    curSqlStmt = curSqlStmt & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
    curSqlStmt = curSqlStmt & "			WHERE OT.RatingType_ID = 3 AND OT.EventsConsolidated like '%j%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL"
    curSqlStmt = curSqlStmt & "			) as ODJ ON ODJ.PersonID = MX.PersonID"

    curSqlStmt = curSqlStmt & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
    curSqlStmt = curSqlStmt & "			FROM " & OfficialsTableName & " as OT"
    curSqlStmt = curSqlStmt & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
    curSqlStmt = curSqlStmt & "			WHERE OT.RatingType_ID = 2 AND OT.EventsConsolidated like '%s%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL"
    curSqlStmt = curSqlStmt & "			) as OCS ON OCS.PersonID = MX.PersonID"
    curSqlStmt = curSqlStmt & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
    curSqlStmt = curSqlStmt & "			FROM " & OfficialsTableName & " as OT"
    curSqlStmt = curSqlStmt & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
    curSqlStmt = curSqlStmt & "			WHERE OT.RatingType_ID = 2 AND OT.EventsConsolidated like '%t%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL"
    curSqlStmt = curSqlStmt & "			) as OCT ON OCT.PersonID = MX.PersonID"
    curSqlStmt = curSqlStmt & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
    curSqlStmt = curSqlStmt & "			FROM " & OfficialsTableName & " as OT"
    curSqlStmt = curSqlStmt & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
    curSqlStmt = curSqlStmt & "			WHERE OT.RatingType_ID = 2 AND OT.EventsConsolidated like '%j%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL"
    curSqlStmt = curSqlStmt & "			) as OCJ ON OCJ.PersonID = MX.PersonID"

    curSqlStmt = curSqlStmt & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
    curSqlStmt = curSqlStmt & "			FROM " & OfficialsTableName & " as OT"
    curSqlStmt = curSqlStmt & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
    curSqlStmt = curSqlStmt & "			WHERE OT.RatingType_ID = 9 AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL"
    curSqlStmt = curSqlStmt & "			) as OS ON OS.PersonID = MX.PersonID"

    curSqlStmt = curSqlStmt & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
    curSqlStmt = curSqlStmt & "			FROM " & OfficialsTableName & " as OT"
    curSqlStmt = curSqlStmt & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
    curSqlStmt = curSqlStmt & "			WHERE OT.RatingType_ID = 4 AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL"
    curSqlStmt = curSqlStmt & "			) as OTC ON OTC.PersonID = MX.PersonID"

    ' -----------------------------------------------
    ' Where clause
    ' -----------------------------------------------
    curSqlStmt = curSqlStmt & " WHERE MX.ExportToTemplate = 1"
    curSqlStmt = curSqlStmt & "  AND DateAdd(mm,18,MX.EffTo) > GetDate()"
    curSqlStmt = curSqlStmt & "  AND MX.Deceased = 0 "
    curSqlStmt = curSqlStmt & "  AND ( MX.TeamTournAppID = '" & curSanctionId & "'"
    curSqlStmt = curSqlStmt & "        OR MX.PersonID in (SELECT PersonID FROM " & ApptOfficialsTableName & " Where TournAppID = '" & curSanctionId & "')"
    curSqlStmt = curSqlStmt & "      )"

    ' -----------------------------------------------
    ' Order by clause
    ' -----------------------------------------------
    curSqlStmt = curSqlStmt & " ORDER BY CASE WHEN SO.OffCode is not Null and MX.EventSlalom + MX.EventTrick + MX.EventJump = '      ' THEN 'E0FF' ELSE MX.Sorter END"
    curSqlStmt = curSqlStmt & "         , MX.LastName, MX.FirstName, MX.MemberID"

    '	-----------------------------------------------------------------------
    ' Execute SQL statement to retrieve skier information and load to registration template
    '	-----------------------------------------------------------------------
    buildQueryMemberRegNcwsaEntries = curSqlStmt

End Function

'	-----------------------------------------------------------------------
' Refresh the list of chief and appointed officials for a tournament
' The data is stored in a temporary work table for use in build tournament registration entries
'	-----------------------------------------------------------------------
Function refreshApptOfficials(curSanctionId)
    Set WaterskiConnect = Server.CreateObject("ADODB.Connection")
    WaterskiConnect.Open Application("WaterSkiConn")

    curSqlStmt = "Delete from " & ApptOfficialsTableName & " where TournAppID = '" & curSanctionId & "' OR DateAdd(Day,30,WhenAdded) < GetDate()"
    WaterskiConnect.Execute (curSqlStmt)

    curSqlStmt = "Insert into " & ApptOfficialsTableName & " (PersonID, TournAppID, OffCode, WhenAdded) "
    curSqlStmt = curSqlStmt & "SELECT PersonID, '" & curSanctionId & "', Max(OffCode), GetDate() "
    curSqlStmt = curSqlStmt & "FROM ( "
    curSqlStmt = curSqlStmt & "    SELECT Cast(CASE When len(CJudgePID) < 9 THEN CJudgePID ELSE right(CJudgePID,8) END as integer) AS PersonID, 'CJ' AS OffCode "
    curSqlStmt = curSqlStmt & "    FROM " & RegTableName & " WHERE TournAppID = '" & curSanctionId & "' and isnumeric(CJudgePID) = 1 "
    curSqlStmt = curSqlStmt & "    UNION"
    curSqlStmt = curSqlStmt & "    SELECT Cast(CASE When len(CDriverPID) < 9 Then CDriverPID Else right(CDriverPID,8) END as integer) AS PersonID, 'CD' AS OffCode"
    curSqlStmt = curSqlStmt & "    FROM " & RegTableName & " WHERE TournAppID = '" & curSanctionId & "' and isnumeric(CDriverPID) = 1 "
    curSqlStmt = curSqlStmt & "    UNION"
    curSqlStmt = curSqlStmt & "    SELECT Cast(CASE When len(CScorePID) < 9 Then CScorePID Else right(CScorePID,8) END as integer) AS PersonID, 'CC' AS OffCode "
    curSqlStmt = curSqlStmt & "    FROM " & RegTableName & " WHERE TournAppID = '" & curSanctionId & "' and isnumeric(CScorePID) = 1 "
    curSqlStmt = curSqlStmt & "    UNION"
    curSqlStmt = curSqlStmt & "    SELECT Cast(CASE When len(CSafPID) < 9 Then CSafPID Else right(CSafPID,8) END as integer) AS PersonID, 'CS' AS OffCode"
    curSqlStmt = curSqlStmt & "    FROM " & RegTableName & " WHERE TournAppID = '" & curSanctionId & "' and isnumeric(CSafPID) = 1 "
    curSqlStmt = curSqlStmt & "    UNION"
    curSqlStmt = curSqlStmt & "    SELECT Cast(CASE When len(TechCPID) < 9 Then TechCPID Else right(TechCPID,8) END as integer) AS PersonID, 'CT' AS OffCode"
    curSqlStmt = curSqlStmt & "    FROM " & RegTableName & " WHERE TournAppID = '" & curSanctionId & "' and isnumeric(TechCPID) = 1 "
    curSqlStmt = curSqlStmt & "    UNION"
    curSqlStmt = curSqlStmt & "    SELECT Cast(CASE When len(Ap1JPID) < 9 Then Ap1JPID Else right(Ap1JPID,8) END as integer) AS PersonID, 'APTJ' AS OffCode"
    curSqlStmt = curSqlStmt & "    FROM " & RegTableName & " WHERE TournAppID = '" & curSanctionId & "' and isnumeric(Ap1JPID) = 1 "
    curSqlStmt = curSqlStmt & "    UNION"
    curSqlStmt = curSqlStmt & "    SELECT Cast(CASE When len(Ap2JPID) < 9 Then Ap2JPID Else right(Ap2JPID,8) END as integer) AS PersonID, 'APTJ' AS OffCode"
    curSqlStmt = curSqlStmt & "    FROM " & RegTableName & " WHERE TournAppID = '" & curSanctionId & "' and isnumeric(Ap2JPID) = 1 "
    curSqlStmt = curSqlStmt & "    UNION"
    curSqlStmt = curSqlStmt & "    SELECT Cast(CASE when len(Ap3JPID) < 9 Then Ap3JPID Else right(Ap3JPID,8) END as integer) AS PersonID, 'APTJ' AS OffCode"
    curSqlStmt = curSqlStmt & "    FROM " & RegTableName & " WHERE TournAppID = '" & curSanctionId & "' and isnumeric(Ap3JPID) = 1 "
    curSqlStmt = curSqlStmt & "    UNION"
    curSqlStmt = curSqlStmt & "    SELECT Cast(CASE when len(Ap4JPID) < 9 Then Ap4JPID Else right(Ap4JPID,8) END as integer) AS PersonID, 'APTJ' AS OffCode"
    curSqlStmt = curSqlStmt & "    FROM " & RegTableName & " WHERE TournAppID = '" & curSanctionId & "' and isnumeric(Ap4JPID) = 1 "
    curSqlStmt = curSqlStmt & "    UNION"
    curSqlStmt = curSqlStmt & "    SELECT Cast(CASE when len(Ap5JPID) < 9 Then Ap5JPID Else right(Ap5JPID,8) END as integer) AS PersonID, 'APTJ' AS OffCode"
    curSqlStmt = curSqlStmt & "    FROM " & RegTableName & " WHERE TournAppID = '" & curSanctionId & "' and isnumeric(Ap5JPID) = 1 "
    curSqlStmt = curSqlStmt & "    UNION"
    curSqlStmt = curSqlStmt & "    SELECT Cast(CASE when len(Ap1SPID) < 9 Then Ap1SPID Else right(Ap1SPID,8) END as integer) AS PersonID, 'APTS' AS OffCode"
    curSqlStmt = curSqlStmt & "    FROM " & RegTableName & " WHERE TournAppID = '" & curSanctionId & "' and isnumeric(Ap1SPID) = 1 "
    curSqlStmt = curSqlStmt & "    UNION"
    curSqlStmt = curSqlStmt & "    SELECT Cast(CASE when len(Ap2SPID) < 9 Then Ap2SPID Else right(Ap2SPID,8) END as integer) AS PersonID, 'APTS' AS OffCode"
    curSqlStmt = curSqlStmt & "    FROM " & RegTableName & " WHERE TournAppID = '" & curSanctionId & "' and isnumeric(Ap2SPID) = 1 "
    curSqlStmt = curSqlStmt & "    UNION"
    curSqlStmt = curSqlStmt & "    SELECT Cast(CASE when len(Ap3SPID) < 9 Then Ap3SPID Else right(Ap3SPID,8) END as integer) AS PersonID, 'APTS' AS OffCode"
    curSqlStmt = curSqlStmt & "    FROM " & RegTableName & " WHERE TournAppID = '" & curSanctionId & "' and isnumeric(Ap3SPID) = 1 "
    curSqlStmt = curSqlStmt & "    UNION"
    curSqlStmt = curSqlStmt & "    SELECT Cast(CASE when len(Ap1DrPID) < 9 Then Ap1DrPID Else right(Ap1DrPID,8) END as integer) AS PersonID, 'APTD' AS OffCode"
    curSqlStmt = curSqlStmt & "    FROM " & RegTableName & " WHERE TournAppID = '" & curSanctionId & "' and isnumeric(Ap1DrPID) = 1 "
    curSqlStmt = curSqlStmt & "    UNION"
    curSqlStmt = curSqlStmt & "    SELECT Cast(CASE when len(PanAmPID) < 9 Then PanAmPID Else right(PanAmPID,8) END as integer) AS PersonID, 'APTJ' AS OffCode"
    curSqlStmt = curSqlStmt & "    FROM " & RegTableName & " WHERE TournAppID = '" & curSanctionId & "' and isnumeric(PanAmPID) = 1"
    curSqlStmt = curSqlStmt & " ) SOX Group by PersonID"
    
    WaterskiConnect.Execute (curSqlStmt)
    WaterskiConnect.Close

End Function


'	-----------------------------------------------------------------------
' END COMMON FUNCTIONS AND DEFINITIONS
'	-----------------------------------------------------------------------
%>
