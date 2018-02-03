<!--#include virtual="/epl/functions.asp" -->
<!--#include virtual="/admin/MemberRegFunctions.asp"-->

<% 

If not Session("aauth") then response.redirect "Login.asp"

Server.ScriptTimeout = 10

' The following lines of HTML display the "opening please wait" banner.
''''http://usawaterski.org/admin/CreateNCWSATemplate.asp

Dim curTraceMsg
Dim curSqlStmt

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



%>
    
<html>
    <head><title>USA Water Ski NCWSA Registration Template</title>
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
'	Start by sucking Membership Pricing Info from HQ Table into local Array
'	-----------------------------------------------------------------------
Dim MT, MemPrice(200), MemUpgrd(200)
FOR MT = 1 to 200: MemPrice(MT) = 0: MemUpgrd(MT) = 0: NEXT

Set SQLConnect = CreateObject("ADODB.Connection")
SQLConnect.Open Application("HQSQLConn")

curSqlStmt = "SELECT * FROM [Membership Types with pricing]" 
curSqlStmt = curSqlStmt & " WHERE EffectiveFrom <= CONVERT(DATETIME, '" & session("tournamentdate") & " 00:00:00', 102)"
curSqlStmt = curSqlStmt & " AND EffectiveTo >= CONVERT(DATETIME, '" & session("tournamentdate") & " 00:00:00', 102)"
    On Error Resume Next
Set HQRS = SQLConnect.Execute(curSqlStmt)
    If Err.Number <> 0 Then
        %>
            <DIV ID="debugMsg">
                <br />Error retrieving membership pricing 
                <br />Err.Number=<%=Err.Number %>
                <br />Err.Description=<%=Err.Description %>
                <br />
            </DIV>
        <%
        On Error Goto 0 ' But don't let other errors hide!
    End If
DO UNTIL HQRS.EOF
	MT = HQRS("Membership Type Code")
	MemPrice(MT) = HQRS("MemberShipTypeRates")
	MemUpgrd(MT) = HQRS("CostToUpgrade")
	HQRS.MoveNext
LOOP

HQRS.Close
Set HQRS = Nothing


'	-----------------------------------------------------------------------
'	???????
'	-----------------------------------------------------------------------
Dim objConn
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open Application("WaterSkiConn")

Dim objFSO
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

'	-----------------------------------------------------------------------
'	Format current date for using in file name
'	-----------------------------------------------------------------------
Dim DateRaw, DateFmt, I1, I2, RowNo
DateRaw = Date(): I1 = instr(DateRaw,"/"): I2 = instr(I1+1,DateRaw,"/")
DateFmt = Mid(DateRaw,I2+1): ' Start with Year value
IF I1=2 THEN DateFmt = DateFmt + "-0" + Left(DateRaw,1): ELSE DateFmt = DateFmt + "-" + Left(DateRaw,2)
IF I2-I1=2 THEN DateFmt = DateFmt + "-0" + Mid(DateRaw,I1+1,1): ELSE DateFmt = DateFmt + "-" + Mid(DateRaw,I1+1,2)

'	-----------------------------------------------------------------------
'	Retrieve Sanction information
'	-----------------------------------------------------------------------
Dim objRS
Set objRS = Server.CreateObject("ADODB.RecordSet")
objRS.ActiveConnection = objConn

Dim objTL
Set objTL = Server.CreateObject("ADODB.RecordSet")
objTL.ActiveConnection = objConn

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
curSqlStmt = "Select Distinct TSanction, TStatus, TournAppID, TDateE, TName, TCity, TState from Sanctions.dbo.TSchedul where TournAppID = '" & curSanctionId & "'"
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

'Get TStatus and TSanction from TSchedul table, and AllowAccess from Users999
Dim strTStatus, strTSanction, AllowAccess
curSqlStmt = "Select Top 1 TS.TSanction, TS.TStatus, US.AllowAccess"
curSqlStmt = curSqlStmt & " from Sanctions.dbo.TSchedul as TS, USAWaterski.dbo.Users999 as US"
curSqlStmt = curSqlStmt & " where TS.TournAppID = '" & curSanctionId
curSqlStmt = curSqlStmt & "' and US.Name = '" & curSanctionId & "'"

    On Error Resume Next
objRS.Open curSqlStmt
    If Err.Number <> 0 Then
        %>
            <DIV ID="debugMsg">
                <br />Error retrieving Sanction information to build registration template
                <br />Err.Number=<%=Err.Number %>
                <br />Err.Description=<%=Err.Description %>
                <br />
            </DIV>
        <%
        On Error Goto 0 ' But don't let other errors hide!
    End If
If objRS.EOF THEN
	strTStatus = -1: strTSanction = Session("TournamentID"): AllowAccess = false
ELSE 
	strTStatus = objRS("TStatus"): strTSanction = objRS("TSanction"): AllowAccess = objRS("AllowAccess")
	IF left(strTSanction,6) <> curSanctionId THEN
		strTSanction = Session("TournamentID")
	END IF
END IF
objRS.Close

'	-----------------------------------------------------------------------
'Now open a connection to the new XLS file
'Setup to reference blank registration template file
'	-----------------------------------------------------------------------
Dim path
path = Server.MapPath("Excel/")

Dim fileRegXls
Set fileRegXls = Server.CreateObject("Scripting.FileSystemObject")
Dim pathExcelFiles
pathExcelFiles = Server.MapPath("Excel/")
dim copyFileSour, copyFileDest
curTraceMsg = curTraceMsg & "<br /><br />pathExcelFiles=" & pathExcelFiles

copyFileSour = pathExcelFiles & "/Templates/NCWSATemplate2014.xls"
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

'	-----------------------------------------------------------------------
' Refresh the list of chief and appointed officials for a tournament
' The data is stored in a temporary work table for use in build tournament registration entries
'	-----------------------------------------------------------------------
refreshApptOfficials(curSanctionId)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Next we open an extract of Team ID's and Names from the TeamList table.
' Note that we prefix each team name with "E" if the team has entries,
' or "Z" if no entries, so that all the entered teams list at the top.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
curSqlStmt = "" 
curSqlStmt = curSqlStmt & "SELECT CASE WHEN TE.Team is Null THEN 'Z'+TL.TeamID ELSE 'E'+TL.TeamID END as TeamID, TL.TeamName "
curSqlStmt = curSqlStmt & "FROM Cobra00025.USAWSRank.TeamsList as TL "
curSqlStmt = curSqlStmt & "LEFT JOIN (Select DISTINCT team FROM " & TeamRotationsTableName & " "
curSqlStmt = curSqlStmt & "WHERE TournAppID = '" & left(strTSanction,6) 
curSqlStmt = curSqlStmt & "' AND WaiverStat >= 'C') as TE"
curSqlStmt = curSqlStmt & " ON TE.Team = TL.TeamID Where SptsGrpID = 'NCW'"
curSqlStmt = curSqlStmt & " ORDER BY CASE WHEN TE.Team is Null THEN"
curSqlStmt = curSqlStmt & " 'Z'+TL.TeamID ELSE 'E'+TL.TeamID END"

objTL.Open curSqlStmt


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Now build a Query to Extract the Desired Members, joining in data 
''' pulled from the Rankings and Officials and Membership Type tables.
''' Note that we prefix each team ID with "E" if the team has entries,
''' or "Z" if no entries, so that all the entered teams list at the top,
''' then finally all those without any team affiliation last with Zzzz.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'	Begin with the overall select list for the Outer Query

curSqlStmt = "Select Substring(MX.MemberID,1,3) + '-' + Substring(MX.MemberID,4,2) + '-' +" 
curSqlStmt = curSqlStmt & " Substring(MX.MemberID,6,4) as MemID, MX.LastName, MX.FirstName,"
curSqlStmt = curSqlStmt & " Case when MX.Sex = 'F' Then 'CW' else 'CM' end as Div,"

curSqlStmt = curSqlStmt & " Case when MX.Age <= 17 and MX.Sex = 'F' Then 'G'"
curSqlStmt = curSqlStmt & " when MX.Age <= 17 then 'B' when MX.Sex = 'F' then 'W' else 'M' end + Case"
curSqlStmt = curSqlStmt & " when MX.Age <= 9 then '1' when MX.Age <= 13 then '2' when MX.Age <= 17 then '3'"
curSqlStmt = curSqlStmt & " when MX.Age <= 24 then '1' when MX.Age <= 34 then '2' when MX.Age <= 44 then '3'"
curSqlStmt = curSqlStmt & " when MX.Age <= 52 then '4' when MX.Age <= 59 then '5' when MX.Age <= 64 then '6'"
curSqlStmt = curSqlStmt & " when MX.Age <= 69 then '7' when MX.Age <= 74 then '8' when MX.Age <= 79 then '9'"
curSqlStmt = curSqlStmt & " when MX.Age <= 84 then 'A' else 'B' end as AgeDiv,"
		
curSqlStmt = curSqlStmt & " Case when SO.OffCode is not Null and MX.SlmEnt+MX.TrkEnt+MX.JmpEnt"
curSqlStmt = curSqlStmt & " = '      ' then 'E0FF' else MX.Sorter end as Sorter,"

curSqlStmt = curSqlStmt & " MX.Team, MX.TeamStat, MX.Sex, MX.Age, MX.City, MX.State,"

curSqlStmt = curSqlStmt & " Case when OD.PersonID is Null then '-' else Right(OD.RtgLvl,1) end +"
curSqlStmt = curSqlStmt & " Case when OJ.PersonID is Null then '-' else Right(OJ.RtgLvl,1) end +"
curSqlStmt = curSqlStmt & " Case when OC.PersonID is Null then '-' else Right(OC.RtgLvl,1) end +"
curSqlStmt = curSqlStmt & " Case when OS.PersonID is Null then '-' else Right(OS.RtgLvl,1) end as OffRat,"

curSqlStmt = curSqlStmt & " Coalesce(SO.OffCode,'') as OffCode,"

curSqlStmt = curSqlStmt & " MX.EffTo, MX.Memtype, MX.MemCode, MX.CanSki, MX.SptsDiv, MX.AnnWvr,"
curSqlStmt = curSqlStmt & " MX.EvtWvr, MX.SlmEnt, MX.TrkEnt, MX.JmpEnt, MX.TrkBt, MX.JmpRH"


'	This begins the major "MX" Sub-query, which pulls membership and team and entry information
		
curSqlStmt = curSqlStmt & " From (Select MT.PersonIDWithCheckDigit as MemberID, MT.PersonID,"
curSqlStmt = curSqlStmt & " Left(MT.LastName,12) as LastName, Left(MT.FirstName,10) as FirstName,"

curSqlStmt = curSqlStmt & " (" & Session("TournamentYear") & "-Year(MT.BirthDate)-1) as Age,"

curSqlStmt = curSqlStmt & " Left(MT.City,12) as City, Left(MT.State,2) as State,"
curSqlStmt = curSqlStmt & " MT.EffectiveTo as EffTo, MT.MembershipTypeCode as MemType,"
curSqlStmt = curSqlStmt & " Typ.TypeCode as MemCode, Typ.CanSkiInTournaments as CanSki,"
curSqlStmt = curSqlStmt & " MT.DivisionCode1 + '/' + MT.DivisionCode2 as SptsDiv,"
curSqlStmt = curSqlStmt & " Upper(Left(MT.Sex,1)) as Sex, MT.WaiverStatusID as AnnWvr,"

curSqlStmt = curSqlStmt & " Case when TE.Team is not null then 'E' else 'Z' end +"
curSqlStmt = curSqlStmt & " Case when Coalesce(RP.Team,TR.Team) is not null then"
curSqlStmt = curSqlStmt & " Coalesce(RP.Team,TR.Team) else 'zzz' end as Sorter,"

curSqlStmt = curSqlStmt & " Case when RP.MemberID is not null then 'A' when"
curSqlStmt = curSqlStmt & " TR.DateInactive is not null then 'I' else 'A' end as TeamStat,"

curSqlStmt = curSqlStmt & " Coalesce(RP.Team,TR.Team,'   ') as Team,"

' curSqlStmt = curSqlStmt & " Coalesce(RP.SlalomEnt,'  ') as SlmEnt," 
curSqlStmt = curSqlStmt & " Coalesce(Case when right(RP.SlalomEnt,1) <= '9' then RP.SlalomEnt"
curSqlStmt = curSqlStmt & " else left(RP.SlalomEnt,1) + cast(ascii(right(RP.SlalomEnt,1)) - 55"
curSqlStmt = curSqlStmt & " as varchar(2)) end, '  ') as SlmEnt," 

' curSqlStmt = curSqlStmt & " Coalesce(RP.TrickEnt,'  ') as TrkEnt," 
curSqlStmt = curSqlStmt & " Coalesce(Case when right(RP.TrickEnt,1) <= '9' then RP.TRickEnt"
curSqlStmt = curSqlStmt & " else left(RP.TrickEnt,1) + cast(ascii(right(RP.TrickEnt,1)) - 55"
curSqlStmt = curSqlStmt & " as varchar(2)) end, '  ') as TrkEnt," 

' curSqlStmt = curSqlStmt & " Coalesce(RP.JumpEnt,'  ') as JmpEnt," 
curSqlStmt = curSqlStmt & " Coalesce(Case when right(RP.JumpEnt,1) <= '9' then RP.JumpEnt"
curSqlStmt = curSqlStmt & " else left(RP.JumpEnt,1) + cast(ascii(right(RP.JumpEnt,1)) - 55"
curSqlStmt = curSqlStmt & " as varchar(2)) end, '  ') as JmpEnt," 

curSqlStmt = curSqlStmt & " Coalesce(RP.WaiverStat,' ') as EvtWvr," 
curSqlStmt = curSqlStmt & " Coalesce(RP.TrickBoat,'  ') as TrkBt," 
curSqlStmt = curSqlStmt & " Coalesce(RP.RampHgt,'  ') as JmpRH" 

'	Begin FROM and JOIN table list for "MX" Sub-Query

curSqlStmt = curSqlStmt & " FROM USAWaterski.dbo.Members as MT Inner Join"
curSqlStmt = curSqlStmt & " USAWaterski.dbo.MembershipTypes as Typ"
curSqlStmt = curSqlStmt & " ON MT.MembershipTypeCode = Typ.MemberShipTypeID"


'	Here's the subquery which now pulls Team ID's from the Team Roster Extract.
'	Identify Latest Team affiliation for Member -- new version
curSqlStmt = curSqlStmt & " Left Join (Select RX.MemberID, RX.Team, RX.DateInactive"
curSqlStmt = curSqlStmt & " from Cobra00025.USAWSRank.TeamRoster as RX"
curSqlStmt = curSqlStmt & " join (select MemberID, Max(LastEvent) as MaxEvt"
curSqlStmt = curSqlStmt & " from Cobra00025.USAWSRank.TeamRoster group by MemberID) as ME" 
curSqlStmt = curSqlStmt & " on ME.MemberID = RX.MemberID and ME.MaxEvt = RX.LastEvent) as TR"
curSqlStmt = curSqlStmt & " on TR.MemberID = MT.PersonIDWithCheckDigit"

'	This subquery pulls Rotation Plan information for this Person/TourID -- LEAVE TEAM OUT !! (All Stars)
curSqlStmt = curSqlStmt & " left join Cobra00025.USAWSRank.TeamRotations as RP"
curSqlStmt = curSqlStmt & " on RP.TournAppID = '" & left(strTSanction,6) & "'"
curSqlStmt = curSqlStmt & " and RP.MemberID = MT.PersonIDWithCheckDigit"

'	This subquery identifies Teams that are Entered, used to preface Sorter extract column
curSqlStmt = curSqlStmt & " left join (Select distinct team" 
curSqlStmt = curSqlStmt & " from Cobra00025.USAWSRank.TeamRotations where"
curSqlStmt = curSqlStmt & " WaiverStat >= 'C' and TournAppID = '" & left(strTSanction,6)
curSqlStmt = curSqlStmt & "') as TE on TE.Team = Coalesce(RP.Team,TR.Team)"

' Now here's the "WHERE" condition clause for the Primary "MX" Sub-Query
curSqlStmt = curSqlStmt & " Where Typ.ExporttoTouramentRegistrationTemplate = 1"
curSqlStmt = curSqlStmt & " AND DateAdd(mm,18,MT.EffectiveTo) > GetDate()"
curSqlStmt = curSqlStmt & " AND MT.Deceased = 0 AND ( (" & Session("TournamentYear")
curSqlStmt = curSqlStmt & " - Year(MT.BirthDate) - 1) between 16 and 29 OR"
curSqlStmt = curSqlStmt & " MT.DivisionCode1 = 'NCW' OR MT.DivisionCode2 = 'NCW' OR"

curSqlStmt = curSqlStmt & " PersonID in (Select PersonID from USAWaterski.dbo.TempApptdOfcls"
curSqlStmt = curSqlStmt & " Where TournAppID = '" & curSanctionId & "') OR"

' Added "OR" condition to bring in all AWSA Rated Officials 2016-09-23
curSqlStmt = curSqlStmt & " PersonID in (Select distinct PersonID"
curSqlStmt = curSqlStmt & " FROM USAWaterski.dbo.Officials OT INNER JOIN"
curSqlStmt = curSqlStmt & " USAWaterski.dbo.Level LV ON OT.Level_ID = LV.Level_ID"
curSqlStmt = curSqlStmt & " WHERE OT.DivisionCode = 'AWS'"
curSqlStmt = curSqlStmt & " AND LV.LevelOrderforTemplate IS NOT NULL"
curSqlStmt = curSqlStmt & " AND OT.RatingType_ID in (1,2,3) ) OR"

' Final "OR" condition for ANYBODY appearing in ANY NCWSA Team Roster
curSqlStmt = curSqlStmt & " PersonIDWithCheckDigit IN (Select Distinct MemberID from"
curSqlStmt = curSqlStmt & " Cobra00025.USAWSRank.TeamRoster) ) ) as MX" 

'	End of MX Primary "MX" Select Subquery.  Appended Info Subqueries follow.

curSqlStmt = curSqlStmt & " Left Join (Select OT.PersonID,"
curSqlStmt = curSqlStmt & " Max(convert(char(1),LV.LevelOrderforTemplate)"
curSqlStmt = curSqlStmt & " + LV.LevelAbbreviationforTemplate) AS RtgLvl"
curSqlStmt = curSqlStmt & " FROM USAWaterski.dbo.Officials OT INNER JOIN"
curSqlStmt = curSqlStmt & " USAWaterski.dbo.Level LV ON OT.Level_ID = LV.Level_ID"
curSqlStmt = curSqlStmt & " WHERE OT.DivisionCode in ('AWS','USA')"
curSqlStmt = curSqlStmt & " AND LV.LevelOrderforTemplate IS NOT NULL"
curSqlStmt = curSqlStmt & " AND OT.RatingType_ID = 3 GROUP BY OT.PersonID) as OD"
curSqlStmt = curSqlStmt & " on OD.PersonID = MX.PersonID"

curSqlStmt = curSqlStmt & " Left Join (Select OT.PersonID,"
curSqlStmt = curSqlStmt & " Max(convert(char(1),LV.LevelOrderforTemplate)"
curSqlStmt = curSqlStmt & " + LV.LevelAbbreviationforTemplate) AS RtgLvl"
curSqlStmt = curSqlStmt & " FROM USAWaterski.dbo.Officials OT INNER JOIN"
curSqlStmt = curSqlStmt & " USAWaterski.dbo.Level LV ON OT.Level_ID = LV.Level_ID"
curSqlStmt = curSqlStmt & " WHERE OT.DivisionCode in ('AWS','USA')"
curSqlStmt = curSqlStmt & " AND LV.LevelOrderforTemplate IS NOT NULL"
curSqlStmt = curSqlStmt & " AND OT.RatingType_ID = 1 GROUP BY OT.PersonID) as OJ"
curSqlStmt = curSqlStmt & " on OJ.PersonID = MX.PersonID"

curSqlStmt = curSqlStmt & " Left Join (Select OT.PersonID,"
curSqlStmt = curSqlStmt & " Max(convert(char(1),LV.LevelOrderforTemplate)"
curSqlStmt = curSqlStmt & " + LV.LevelAbbreviationforTemplate) AS RtgLvl"
curSqlStmt = curSqlStmt & " FROM USAWaterski.dbo.Officials OT INNER JOIN"
curSqlStmt = curSqlStmt & " USAWaterski.dbo.Level LV ON OT.Level_ID = LV.Level_ID"
curSqlStmt = curSqlStmt & " WHERE OT.DivisionCode in ('AWS','USA')"
curSqlStmt = curSqlStmt & " AND LV.LevelOrderforTemplate IS NOT NULL"
curSqlStmt = curSqlStmt & " AND OT.RatingType_ID = 2 GROUP BY OT.PersonID) as OC"
curSqlStmt = curSqlStmt & " on OC.PersonID = MX.PersonID"

curSqlStmt = curSqlStmt & " Left Join (Select OT.PersonID,"
curSqlStmt = curSqlStmt & " Max(convert(char(1),LV.LevelOrderforTemplate)"
curSqlStmt = curSqlStmt & " + LV.LevelAbbreviationforTemplate) AS RtgLvl"
curSqlStmt = curSqlStmt & " FROM USAWaterski.dbo.Officials OT INNER JOIN"
curSqlStmt = curSqlStmt & " USAWaterski.dbo.Level LV ON OT.Level_ID = LV.Level_ID"
curSqlStmt = curSqlStmt & " WHERE OT.DivisionCode in ('AWS','USA')"
curSqlStmt = curSqlStmt & " AND LV.LevelOrderforTemplate IS NOT NULL"
curSqlStmt = curSqlStmt & " AND OT.RatingType_ID = 9 GROUP BY OT.PersonID) as OS"
curSqlStmt = curSqlStmt & " on OS.PersonID = MX.PersonID"

curSqlStmt = curSqlStmt & " Left Join	(Select PersonID, OffCode from USAWaterski.dbo.TempApptdOfcls"
curSqlStmt = curSqlStmt & " Where TournAppID = '" & curSanctionId & "')"
curSqlStmt = curSqlStmt & " as SO on SO.PersonID = MX.PersonID"

curSqlStmt = curSqlStmt & " Order By Case when SO.OffCode is not Null and MX.SlmEnt+MX.TrkEnt+MX.JmpEnt"
curSqlStmt = curSqlStmt & " = '      ' then 'E0FF' else MX.Sorter end,"
curSqlStmt = curSqlStmt & " MX.LastName, MX.FirstName, MX.MemberID"

' Response.write curSqlStmt

objRS.Open curSqlStmt



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Now we loop through the Extract of Members, merging the
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
	
	'	First step is to derive "Reason Not OK to Ski" and renew/upgrade amount strings
	
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
	

	'	Next step is to see if we've got a new Team here.
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
			objExcelRegist.Fields(11).Value = " <Rides  Skiers>"
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
		objExcelRegist.Fields(0).Value = objRS("MemID")
		objExcelRegist.Fields(1).Value = objRS("LastName")
		objExcelRegist.Fields(2).Value = objRS("FirstName")

		IF objRS("Sorter") <> "E0FF" THEN
			objExcelRegist.Fields(3).Value = trim(objRS("Team"))
		ELSE
			objExcelRegist.Fields(3).Value = "OFF"
		END IF

		IF objRS("SlmEnt") = "DD" or objRS("TrkEnt") = "DD" or objRS("JmpEnt") = "DD" or Instr(Ucase(Session("TournamentName")),"ALUMNI") > 0 THEN
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

		objExcelRegist.Fields(13).Value = objRS("TrkBt")
		objExcelRegist.Fields(14).Value = objRS("JmpRH")

		objExcelRegist.Fields(16).Value = objRS("SptsDiv")
		objExcelRegist.Fields(17).Value = OKtoSki
		objExcelRegist.Fields(18).Value = UpgrdAmt

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
			objExcelRegist.Fields(19).Value = NumEvts
			TeamTot = TeamTot + 1
		END IF

		objExcelRegist.Update
			
	END IF


	'	Now we handle detail skier rows for the current LocalTeamID
	'	First primary split is whether this is row goes to actives A/B team or not
	'	Team must be Entered, and Skier Active AND Entered in at least one event.
	
	IF objRS("Sorter") = "E0FF" or (left(objRS("Sorter"),1) = "E" and objRS("TeamStat") = "A" and (objRS("SlmEnt") <> "  " or objRS("TrkEnt") <> "  " or objRS("JmpEnt") <> "  ")) THEN

	ELSE

		'	*******	All Others go in this section here ...

		objExcelAllOthr.addnew
		objExcelAllOthr.Fields(0).Value = objRS("MemID")
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
objExcelRegist.Fields(11).Value = " <Rides  Skiers>"
' objExcelRegist.Fields(13).Value = "Skiers>"
objExcelRegist.Fields(15).Value = GrandTot
objExcelRegist.Update	

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
'objFSO.CopyFile path & "/template_with_scores.xls", path & "/" & filenamewithscores , True
objFSO.CopyFile path & "/template.xls", path & "/" & filenamewithscores , True

'Clean up old files
Set f = objFSO.GetFolder("D:\webs\usawaterski.org\admin\excel\")  
Set fc = f.Files 
Response.Write "<br>"
For Each f1 in fc
	'Response.Write f1.name 
	Set myfile = objFSO.GetFile("D:\webs\usawaterski.org\admin\excel\" & f1.name)
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