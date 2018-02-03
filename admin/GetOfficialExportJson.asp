<!--#include virtual="/admin/JSON_2.0.4.asp"-->
<!--#include virtual="/admin/MemberRegFunctions.asp"-->

<%
'	-----------------------------------------------------------------------
' Validate TourID value for scores to be Exported.
' http://usawaterski.org/admin/GetOfficialExportJson.asp?MemberId=700040630
' http://usawaterski.org/admin/GetOfficialExportJson.asp?SanctionId=18E024
' http://usawaterski.org/admin/GetOfficialExportJson.asp?StateList=MA,CT
' http://usawaterski.org/admin/GetOfficialExportJson.asp?MemberId=700040630&user=18E024&password=10089
''	-----------------------------------------------------------------------

Dim curAuth, curAuthParts, curCredParts, curCount, curRqstAuth, curAuthResult
Dim curSanctionId, curMemberId, curStateSQL, curStateList, curTourYear, curDate

curRqstAuth = 0
curRqstAuth = CheckBasicAuth()
IF curRqstAuth = 0 THEN
	response.write "Invalid credentials, unable to complete request"
	response.status = "401 Unauthorized"
	response.flush
	response.end
END IF

curStateList = Request.QueryString("StateList")
curSanctionId = Request.QueryString("SanctionId")
curMemberId = Request.QueryString("MemberId")
curMemberFirstName = Request.QueryString("FirstName")
curMemberLastName = Request.QueryString("LastName")

IF len(curSanctionId) > 0 THEN
    curTourYear = 2000 + left(curSanctionId,2)
ELSE
    curDate = Date
    curTourYear = Right(curDate, 4)
    curMonthStg = Left(curDate, 2)
    IF Right(curMonthStg, 1) = "/" THEN
        curMonth = Left(curDate, 1)
    ELSE
        curMonth = Left(curDate, 2)
    END IF
    if curMonth > 8 THEN curTourYear = curTourYear + 1
END IF

IF len(curStateList) > 0 THEN
    curStateSQL = BuildStateSQL(curStateList)
ELSE
    curStateSQL = "State IN ('')"
END IF

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Now build a Query to Extract the Desired Members, joining in data
' from the Rankings and Officials and Membership Type tables.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim curSqlStmt
curSqlStmt = ""
curSqlStmt = curSqlStmt & "SELECT MT.PersonIDWithCheckDigit as MemberID, MT.PersonID , MT.LastName, MT.FirstName, FederationCode as Federation"
curSqlStmt = curSqlStmt & ", MT.MembershipTypeCode, Typ.ActiveMember, Typ.Description as MemTypeDesc"
curSqlStmt = curSqlStmt & ", Upper(Left(MT.Sex,1)) as Gender, MT.City, MT.State"
curSqlStmt = curSqlStmt & ", 2018 - Year(MT.BirthDate) - 1 as Age , MT.EffectiveTo as EffTo"
curSqlStmt = curSqlStmt & ", Coalesce(Typ.CanSkiInTournaments, '') as CanSki, Coalesce(Typ.CanSkiInGRTournaments, '') as CanSkiGR, MT.WaiverStatusID as Waiver"

curSqlStmt = curSqlStmt & ", Coalesce(OJS.Rating, '') as JudgeSlalom, Coalesce(OJT.Rating, '') as JudgeTrick, Coalesce(OJJ.Rating, '') as JudgeJump"
curSqlStmt = curSqlStmt & ", Coalesce(ODS.Rating, '') as DriverSlalom, Coalesce(ODT.Rating, '') as DriverTrick, Coalesce(ODJ.Rating, '') as DriverJump"
curSqlStmt = curSqlStmt & ", Coalesce(OCS.Rating, '') as ScorerSlalom, Coalesce(OCT.Rating, '') as ScorerTrick, Coalesce(OCJ.Rating, '') as ScorerJump"
curSqlStmt = curSqlStmt & ", Coalesce(OS.Rating, '') as Safety, Coalesce(OTC.Rating, '') as TechController"
curSqlStmt = curSqlStmt & ", Coalesce(OPS.Rating, '') as JudgePanAmSlalom, Coalesce(OPT.Rating, '') as JudgePanAmTrick, Coalesce(OPJ.Rating, '') as JudgePanAmJump"

curSqlStmt = curSqlStmt & " FROM " & MemberTableName & " as MT"
curSqlStmt = curSqlStmt & "    INNER JOIN " & MembershipTypesTableName & " as Typ ON MT.MembershipTypeCode = Typ.MemberShipTypeID"

curSqlStmt = curSqlStmt & " LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
curSqlStmt = curSqlStmt & "     FROM " & OfficialsTableName & " as OT"
curSqlStmt = curSqlStmt & "         INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
curSqlStmt = curSqlStmt & "     WHERE OT.RatingType_ID = 1 AND OT.EventsConsolidated like '%s%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL AND LV.LevelOrderforTemplate < 5"
curSqlStmt = curSqlStmt & "         ) as OJS ON OJS.PersonID = MT.PersonID"
curSqlStmt = curSqlStmt & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
curSqlStmt = curSqlStmt & "			FROM " & OfficialsTableName & " as OT"
curSqlStmt = curSqlStmt & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
curSqlStmt = curSqlStmt & "			WHERE OT.RatingType_ID = 1 AND OT.EventsConsolidated like '%t%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL AND LV.LevelOrderforTemplate < 5"
curSqlStmt = curSqlStmt & "			) as OJT ON OJT.PersonID = MT.PersonID"
curSqlStmt = curSqlStmt & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
curSqlStmt = curSqlStmt & "			FROM " & OfficialsTableName & " as OT"
curSqlStmt = curSqlStmt & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
curSqlStmt = curSqlStmt & "			WHERE OT.RatingType_ID = 1 AND OT.EventsConsolidated like '%j%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL AND LV.LevelOrderforTemplate < 5"
curSqlStmt = curSqlStmt & "			) as OJJ ON OJJ.PersonID = MT.PersonID"

curSqlStmt = curSqlStmt & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Abbreviation as Rating"
curSqlStmt = curSqlStmt & "			FROM " & OfficialsTableName & " as OT"
curSqlStmt = curSqlStmt & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
curSqlStmt = curSqlStmt & "			WHERE OT.RatingType_ID = 1 AND OT.EventsConsolidated like '%s%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate = 5"
curSqlStmt = curSqlStmt & "			) as OPS ON OPS.PersonID = MT.PersonID"
curSqlStmt = curSqlStmt & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Abbreviation as Rating"
curSqlStmt = curSqlStmt & "			FROM " & OfficialsTableName & " as OT"
curSqlStmt = curSqlStmt & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
curSqlStmt = curSqlStmt & "			WHERE OT.RatingType_ID = 1 AND OT.EventsConsolidated like '%t%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate = 5"
curSqlStmt = curSqlStmt & "			) as OPT ON OPT.PersonID = MT.PersonID"
curSqlStmt = curSqlStmt & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Abbreviation as Rating"
curSqlStmt = curSqlStmt & "			FROM " & OfficialsTableName & " as OT"
curSqlStmt = curSqlStmt & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
curSqlStmt = curSqlStmt & "			WHERE OT.RatingType_ID = 1 AND OT.EventsConsolidated like '%j%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate = 5"
curSqlStmt = curSqlStmt & "			) as OPJ ON OPJ.PersonID = MT.PersonID"

curSqlStmt = curSqlStmt & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
curSqlStmt = curSqlStmt & "			FROM " & OfficialsTableName & " as OT"
curSqlStmt = curSqlStmt & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
curSqlStmt = curSqlStmt & "			WHERE OT.RatingType_ID = 3 AND OT.EventsConsolidated like '%s%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL"
curSqlStmt = curSqlStmt & "			) as ODS ON ODS.PersonID = MT.PersonID"
curSqlStmt = curSqlStmt & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
curSqlStmt = curSqlStmt & "			FROM " & OfficialsTableName & " as OT"
curSqlStmt = curSqlStmt & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
curSqlStmt = curSqlStmt & "			WHERE OT.RatingType_ID = 3 AND OT.EventsConsolidated like '%t%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL"
curSqlStmt = curSqlStmt & "			) as ODT ON ODT.PersonID = MT.PersonID"
curSqlStmt = curSqlStmt & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
curSqlStmt = curSqlStmt & "			FROM " & OfficialsTableName & " as OT"
curSqlStmt = curSqlStmt & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
curSqlStmt = curSqlStmt & "			WHERE OT.RatingType_ID = 3 AND OT.EventsConsolidated like '%j%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL"
curSqlStmt = curSqlStmt & "			) as ODJ ON ODJ.PersonID = MT.PersonID"

curSqlStmt = curSqlStmt & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
curSqlStmt = curSqlStmt & "			FROM " & OfficialsTableName & " as OT"
curSqlStmt = curSqlStmt & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
curSqlStmt = curSqlStmt & "			WHERE OT.RatingType_ID = 2 AND OT.EventsConsolidated like '%s%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL"
curSqlStmt = curSqlStmt & "			) as OCS ON OCS.PersonID = MT.PersonID"
curSqlStmt = curSqlStmt & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
curSqlStmt = curSqlStmt & "			FROM " & OfficialsTableName & " as OT"
curSqlStmt = curSqlStmt & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
curSqlStmt = curSqlStmt & "			WHERE OT.RatingType_ID = 2 AND OT.EventsConsolidated like '%t%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL"
curSqlStmt = curSqlStmt & "			) as OCT ON OCT.PersonID = MT.PersonID"
curSqlStmt = curSqlStmt & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
curSqlStmt = curSqlStmt & "			FROM " & OfficialsTableName & " as OT"
curSqlStmt = curSqlStmt & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
curSqlStmt = curSqlStmt & "			WHERE OT.RatingType_ID = 2 AND OT.EventsConsolidated like '%j%' AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL"
curSqlStmt = curSqlStmt & "			) as OCJ ON OCJ.PersonID = MT.PersonID"

curSqlStmt = curSqlStmt & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
curSqlStmt = curSqlStmt & "			FROM " & OfficialsTableName & " as OT"
curSqlStmt = curSqlStmt & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
curSqlStmt = curSqlStmt & "			WHERE OT.RatingType_ID = 9 AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL"
curSqlStmt = curSqlStmt & "			) as OS ON OS.PersonID = MT.PersonID"

curSqlStmt = curSqlStmt & "	LEFT OUTER JOIN (Select OT.PersonID, LV.Level as Rating"
curSqlStmt = curSqlStmt & "			FROM " & OfficialsTableName & " as OT"
curSqlStmt = curSqlStmt & "				INNER JOIN " & RatingLevelTableName & " as LV ON OT.Level_ID = LV.Level_ID"
curSqlStmt = curSqlStmt & "			WHERE OT.RatingType_ID = 4 AND OT.DivisionCode in ('AWS','USA') AND LV.LevelOrderforTemplate IS NOT NULL"
curSqlStmt = curSqlStmt & "			) as OTC ON OTC.PersonID = MT.PersonID"

curSqlStmt = curSqlStmt & " WHERE Typ.ExporttoTouramentRegistrationTemplate = 1"
curSqlStmt = curSqlStmt & "  AND DateAdd(mm,18,MT.EffectiveTo) > GetDate()"
curSqlStmt = curSqlStmt & "  AND MT.Deceased = 0"
curSqlStmt = curSqlStmt & "  AND ( OJS.PersonID is not null OR OJT.PersonID is not null OR OJJ.PersonID is not null"
curSqlStmt = curSqlStmt & "  		OR ODS.PersonID is not null OR ODT.PersonID is not null OR ODJ.PersonID is not null"
curSqlStmt = curSqlStmt & "  		OR OCS.PersonID is not null OR OCT.PersonID is not null OR OCJ.PersonID is not null"
curSqlStmt = curSqlStmt & "  		OR OS.PersonID is not null OR OTC.PersonID is not null "
curSqlStmt = curSqlStmt & "  		)"

IF len(curStateList) > 0 THEN
	curSqlStmt = curSqlStmt & "  AND (" & curStateSQL
	curSqlStmt = curSqlStmt & "       ) "
END IF
IF len(curSanctionId) > 0 THEN
    curSqlStmt = curSqlStmt & "  AND (MT.PersonID in (Select PersonID from " & ApptOfficialsTableName & " WHERE TournAppID = '" & curSanctionId & "' ) "
    curSqlStmt = curSqlStmt & "   OR MT.PersonIDWithCheckDigit in (Select MemberID from " & RegEventsTableName & " WHERE left(TourID,6) = '" & curSanctionId & "') "
	curSqlStmt = curSqlStmt & "       ) "
END IF
IF len(curMemberId) > 0 THEN
    curSqlStmt = curSqlStmt & "  And MT.PersonIDWithCheckDigit = '" & curMemberId & "' "
END IF
IF len(curMemberFirstName) > 0 OR len(curMemberLastName) > 0 THEN
    curSqlStmt = curSqlStmt & "  AND MT.FirstName like '" & curMemberFirstName & "%' AND MT.LastName like '" & curMemberLastName & "%' "
END IF

curSqlStmt = curSqlStmt & " Order by MT.LastName, MT.FirstName, MT.PersonIDWithCheckDigit"
''response.Write = curSqlStmt

'	-----------------------------------------------------------------------
' Execute SQL statement to retrieve skier information and load to registration template
'	-----------------------------------------------------------------------
Set WaterskiConnect = Server.CreateObject("ADODB.Connection")
WaterskiConnect.Open Application("WaterSkiConn")

response.ContentType="application/json"
QueryToJSON(WaterskiConnect, curSqlStmt).flush
%>
