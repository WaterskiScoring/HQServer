<!--#include file="settingsHQ.asp"-->

<html><head><title>Sanction Official Exception Identification</title></head><body>

<%

'	Mainline Code here -- set up variables, then act on "Process" Value

Dim objRS, NumFix, NunLeft, Process
SET objRS=Server.CreateObject("ADODB.recordset")

'	Form Variable "Process" dictates what we do in this invocation

Process = Request("Process")
IF len(Process) = 0 then Process = "Start"

SELECT CASE Process 

CASE "Start"

	PopulateTempTable
	MatchToKnownOfficials
	ReplaceNicknames
	MatchToKnownOfficials
	UpdateSanctionTable
	ListExceptions
	
CASE "NoChgs"

	NumFix = 0
	ListExceptions

CASE "ChgName"

	sSQL1 = "Update USAWSRank.OfclsWithoutIDs set OffName = '" & request("NewName")
	sSQL1 = sSQL1 & "' where OffName = '" & request("OldName") & "'"
	OpenCon: con.execute ( sSQL1 ): CloseCon
	MatchToKnownOfficials
	UpdateSanctionTable
	ListExceptions

CASE "ApplyID"

	sSQL1 = "Update USAWSRank.OfclsWithoutIDs set PersonID = '" & request("PersonID")
	sSQL1 = sSQL1 & "' where OffName = '" & request("OldName") & "'"
	OpenCon: con.execute ( sSQL1 ): CloseCon
	UpdateSanctionTable
	ListExceptions

CASE "EditOff"

	WriteIndexPageHeader

	%>
	
	<Table class="innertable" width=80% align=center><TR><TD Colspan=3>

	<br><center><h2>Alter Official's Name or Supply their Person ID<br></h2></center>
	
	</TD></TR>
	
	<TR>    

	<TD><Center><br>Name from Sanction<br>&nbsp;<br> <font="2"><b><%=Request("OffName")%></b></font><br>
	<FORM method="post" action="/rankings/OffNameRepair.asp">
	<INPUT type="hidden" name="Process" value="NoChgs">
	<INPUT type="Submit" value="No Change">
	</FORM></center></TD>

	<TD><Center><br>Revised Name<br>
	<FORM method="post" action="/rankings/OffNameRepair.asp">
	<INPUT type="hidden" name="Process" value="ChgName">
	<INPUT type="hidden"  name="OldName" value="<%=Request("OffName")%>">
	&nbsp;&nbsp; <INPUT type="textbox" name="NewName" value="<%=Request("OffName")%>">&nbsp;&nbsp; <br>&nbsp;<br>
	<INPUT type="Submit" value="Revise Name">
	</FORM></center></TD>

	<TD><Center><br>Person ID<br>
	<FORM method="post" action="/rankings/OffNameRepair.asp">
	<INPUT type="hidden" name="Process" value="ApplyID">
	<INPUT type="hidden"  name="OldName" value="<%=Request("OffName")%>">
	&nbsp;&nbsp; <INPUT type="textbox" name="PersonID">&nbsp;&nbsp; <br>&nbsp;<br>
	<INPUT type="Submit" value="Apply Person ID">
	</FORM></center></TD>
	
	</TR></Table>
	
	<%
	
	WriteIndexPageFooter

END SELECT


'	---------------------
SUB	PopulateTempTable
'	---------------------	

sSQL1 = "Delete from USAWSRank.OfclsWithoutIDs;"

sSQL1 = sSQL1 & " Insert into USAWSRank.OfclsWithoutIDs Select"
sSQL1 = sSQL1 & " TournAppID, '', ltrim(rtrim(replace(CJudge,'  ',' '))),"
sSQL1 = sSQL1 & " 'CJdg' FROM sanctions.dbo.registration"
sSQL1 = sSQL1 & " WHERE substring(TournAppID,1,2) >= '10'"
sSQL1 = sSQL1 & " and substring(TournAppID,3,1) in"
sSQL1 = sSQL1 & " ('C','E','M','S','W','U','B') and"
sSQL1 = sSQL1 & " len(CJudge) > 5 and len(CJudgePID) = 0;"

sSQL1 = sSQL1 & " Insert into USAWSRank.OfclsWithoutIDs Select"
sSQL1 = sSQL1 & " TournAppID, '', ltrim(rtrim(replace(CDriver,'  ',' '))),"
sSQL1 = sSQL1 & " 'CDrv' FROM sanctions.dbo.registration"
sSQL1 = sSQL1 & " WHERE substring(TournAppID,1,2) >= '10'"
sSQL1 = sSQL1 & " and substring(TournAppID,3,1) in"
sSQL1 = sSQL1 & " ('C','E','M','S','W','U','B') and"
sSQL1 = sSQL1 & " len(CDriver) > 5 and len(CDriverPID) = 0;"

sSQL1 = sSQL1 & " Insert into USAWSRank.OfclsWithoutIDs Select"
sSQL1 = sSQL1 & " TournAppID, '', ltrim(rtrim(replace(CScorer,'  ',' '))),"
sSQL1 = sSQL1 & " 'CScr' FROM sanctions.dbo.registration"
sSQL1 = sSQL1 & " WHERE substring(TournAppID,1,2) >= '10'"
sSQL1 = sSQL1 & " and substring(TournAppID,3,1) in"
sSQL1 = sSQL1 & " ('C','E','M','S','W','U','B') and"
sSQL1 = sSQL1 & " len(CScorer) > 5 and len(CScorePID) = 0;"

sSQL1 = sSQL1 & " Insert into USAWSRank.OfclsWithoutIDs Select"
sSQL1 = sSQL1 & " TournAppID, '', ltrim(rtrim(replace(CSafety,'  ',' '))),"
sSQL1 = sSQL1 & " 'CSft' FROM sanctions.dbo.registration"
sSQL1 = sSQL1 & " WHERE substring(TournAppID,1,2) >= '10'"
sSQL1 = sSQL1 & " and substring(TournAppID,3,1) in"
sSQL1 = sSQL1 & " ('C','E','M','S','W','U','B') and"
sSQL1 = sSQL1 & " len(CSafety) > 5 and len(CSafPID) = 0;"

sSQL1 = sSQL1 & " Insert into USAWSRank.OfclsWithoutIDs Select"
sSQL1 = sSQL1 & " TournAppID, '', ltrim(rtrim(replace(TechCont,'  ',' '))),"
sSQL1 = sSQL1 & " 'CTch' FROM sanctions.dbo.registration"
sSQL1 = sSQL1 & " WHERE substring(TournAppID,1,2) >= '10'"
sSQL1 = sSQL1 & " and substring(TournAppID,3,1) in"
sSQL1 = sSQL1 & " ('C','E','M','S','W','U','B') and"
sSQL1 = sSQL1 & " len(TechCont) > 5 and len(TechCPID) = 0;"

sSQL2 = " Insert into USAWSRank.OfclsWithoutIDs Select"
sSQL2 = sSQL2 & " TournAppID, '', ltrim(rtrim(replace(Ap1Judge,'  ',' '))),"
sSQL2 = sSQL2 & " 'APJ1' FROM sanctions.dbo.registration"
sSQL2 = sSQL2 & " WHERE substring(TournAppID,1,2) >= '10'"
sSQL2 = sSQL2 & " and substring(TournAppID,3,1) in"
sSQL2 = sSQL2 & " ('C','E','M','S','W','U','B') and"
sSQL2 = sSQL2 & " len(Ap1Judge) > 5 and len(Ap1JPID) = 0;"

sSQL2 = sSQL2 & " Insert into USAWSRank.OfclsWithoutIDs Select"
sSQL2 = sSQL2 & " TournAppID, '', ltrim(rtrim(replace(Ap2Judge,'   ',' '))),"
sSQL2 = sSQL2 & " 'APJ2' FROM sanctions.dbo.registration"
sSQL2 = sSQL2 & " WHERE substring(TournAppID,1,2) >= '10'"
sSQL2 = sSQL2 & " and substring(TournAppID,3,1) in"
sSQL2 = sSQL2 & " ('C','E','M','S','W','U','B') and"
sSQL2 = sSQL2 & " len(Ap2Judge) > 5 and len(Ap2JPID) = 0;"

sSQL2 = sSQL2 & " Insert into USAWSRank.OfclsWithoutIDs Select"
sSQL2 = sSQL2 & " TournAppID, '', ltrim(rtrim(replace(Ap3Judge,'  ',' '))),"
sSQL2 = sSQL2 & " 'APJ3' FROM sanctions.dbo.registration"
sSQL2 = sSQL2 & " WHERE substring(TournAppID,1,2) >= '10'"
sSQL2 = sSQL2 & " and substring(TournAppID,3,1) in"
sSQL2 = sSQL2 & " ('C','E','M','S','W','U','B') and"
sSQL2 = sSQL2 & " len(Ap3Judge) > 5 and len(Ap3JPID) = 0;"

sSQL2 = sSQL2 & " Insert into USAWSRank.OfclsWithoutIDs Select"
sSQL2 = sSQL2 & " TournAppID, '', ltrim(rtrim(replace(Ap4Judge,'  ',' '))),"
sSQL2 = sSQL2 & " 'APJ4' FROM sanctions.dbo.registration"
sSQL2 = sSQL2 & " WHERE substring(TournAppID,1,2) >= '10'"
sSQL2 = sSQL2 & " and substring(TournAppID,3,1) in"
sSQL2 = sSQL2 & " ('C','E','M','S','W','U','B') and"
sSQL2 = sSQL2 & " len(Ap4Judge) > 5 and len(Ap4JPID) = 0;"

sSQL2 = sSQL2 & " Insert into USAWSRank.OfclsWithoutIDs Select"
sSQL2 = sSQL2 & " TournAppID, '', ltrim(rtrim(replace(Ap5Judge,'  ',' '))),"
sSQL2 = sSQL2 & " 'APJ5' FROM sanctions.dbo.registration"
sSQL2 = sSQL2 & " WHERE substring(TournAppID,1,2) >= '10'"
sSQL2 = sSQL2 & " and substring(TournAppID,3,1) in"
sSQL2 = sSQL2 & " ('C','E','M','S','W','U','B') and"
sSQL2 = sSQL2 & " len(Ap5Judge) > 5 and len(Ap5JPID) = 0;"

sSQL3 = " Insert into USAWSRank.OfclsWithoutIDs Select"
sSQL3 = sSQL3 & " TournAppID, '', ltrim(rtrim(replace(Ap1Scorer,'  ',' '))),"
sSQL3 = sSQL3 & " 'APS1' FROM sanctions.dbo.registration"
sSQL3 = sSQL3 & " WHERE substring(TournAppID,1,2) >= '10'"
sSQL3 = sSQL3 & " and substring(TournAppID,3,1) in"
sSQL3 = sSQL3 & " ('C','E','M','S','W','U','B') and"
sSQL3 = sSQL3 & " len(Ap1Scorer) > 5 and len(Ap1SPID) = 0;"

sSQL3 = sSQL3 & " Insert into USAWSRank.OfclsWithoutIDs Select"
sSQL3 = sSQL3 & " TournAppID, '', ltrim(rtrim(replace(Ap2Scorer,'  ',' '))),"
sSQL3 = sSQL3 & " 'APS2' FROM sanctions.dbo.registration"
sSQL3 = sSQL3 & " WHERE substring(TournAppID,1,2) >= '10'"
sSQL3 = sSQL3 & " and substring(TournAppID,3,1) in"
sSQL3 = sSQL3 & " ('C','E','M','S','W','U','B') and"
sSQL3 = sSQL3 & " len(Ap2Scorer) > 5 and len(Ap2SPID) = 0;"

sSQL3 = sSQL3 & " Insert into USAWSRank.OfclsWithoutIDs Select"
sSQL3 = sSQL3 & " TournAppID, '', ltrim(rtrim(replace(Ap3Scorer,'  ',' '))),"
sSQL3 = sSQL3 & " 'APS3' FROM sanctions.dbo.registration"
sSQL3 = sSQL3 & " WHERE substring(TournAppID,1,2) >= '10'"
sSQL3 = sSQL3 & " and substring(TournAppID,3,1) in"
sSQL3 = sSQL3 & " ('C','E','M','S','W','U','B') and"
sSQL3 = sSQL3 & " len(Ap3Scorer) > 5 and len(Ap3SPID) = 0;"

sSQL3 = sSQL3 & " Insert into USAWSRank.OfclsWithoutIDs Select"
sSQL3 = sSQL3 & " TournAppID, '', ltrim(rtrim(replace(Ap1Driver,'  ',' '))),"
sSQL3 = sSQL3 & " 'APD1' FROM sanctions.dbo.registration"
sSQL3 = sSQL3 & " WHERE substring(TournAppID,1,2) >= '10'"
sSQL3 = sSQL3 & " and substring(TournAppID,3,1) in"
sSQL3 = sSQL3 & " ('C','E','M','S','W','U','B') and"
sSQL3 = sSQL3 & " len(Ap1Driver) > 5 and len(Ap1DrPID) = 0;"

sSQL3 = sSQL3 & " Insert into USAWSRank.OfclsWithoutIDs Select"
sSQL3 = sSQL3 & " TournAppID, '', ltrim(rtrim(replace(PanAmJudge,'  ',' '))),"
sSQL3 = sSQL3 & " 'PAMJ' FROM sanctions.dbo.registration"
sSQL3 = sSQL3 & " WHERE substring(TournAppID,1,2) >= '10'"
sSQL3 = sSQL3 & " and substring(TournAppID,3,1) in"
sSQL3 = sSQL3 & " ('C','E','M','S','W','U','B') and"
sSQL3 = sSQL3 & " len(PanAmJudge) > 5 and len(PanAmPID) = 0;"

OpenCon: con.execute ( sSQL1 & sSQL2 & sSQL3 ): CloseCon

END SUB



'	------------------------
SUB	MatchToKnownOfficials
'	------------------------

'	This subroutine updates our temporary unknown officials table,
'	by matching against an extract of known officials, pulling 
'	from the USAWaterski Officials and Membership tables.

sSQL1 = "Update OT Set PersonID = MT.PersonID from USAWSRank.OfclsWithoutIDs OT,"
sSQL1 = sSQL1 & " (Select rtrim(FirstName) + ' ' + rtrim(LastName) as MemName," 
sSQL1 = sSQL1 & " convert(varchar(12),max(PersonID)) as PersonID"
sSQL1 = sSQL1 & " From usawaterski.dbo.members where PersonID in"
sSQL1 = sSQL1 & " (Select distinct personid from USAWaterski.dbo.Officials"
sSQL1 = sSQL1 & " where RatingType_ID in (1,2,3,4,5,9)"
sSQL1 = sSQL1 & " and Level_ID in (2,3,4,5,6,13,14,15))"
sSQL1 = sSQL1 & " group by rtrim(FirstName) + ' ' + rtrim(LastName)"
sSQL1 = sSQL1 & " Having count(*) = 1) MT where lower(MT.MemName) = lower(OT.OffName)"

OpenCon: con.execute ( sSQL1 ): CloseCon

END SUB



'	-------------------
SUB	ReplaceNicknames
'	-------------------

'	This subroutine runs a series of updates to the remaining unknown officials
'	table, replacing common nicknames with their more formal forms.

sSQL1 = "Update USAWSRank.OfclsWithoutIDs"
sSQL1 = sSQL1 & " Set OffName = 'Corrine ' + right(Offname,len(Offname)-7)"
sSQL1 = sSQL1 & " Where lower(left(Offname,7)) = 'corrie ' and len(PersonID) = 0;"

sSQL1 = sSQL1 & " Update USAWSRank.OfclsWithoutIDs"
sSQL1 = sSQL1 & " Set OffName = 'Catherine ' + right(Offname,len(Offname)-6)"
sSQL1 = sSQL1 & " Where lower(left(Offname,6)) = 'cathy ' and len(PersonID) = 0;"

sSQL1 = sSQL1 & " Update USAWSRank.OfclsWithoutIDs"
sSQL1 = sSQL1 & " Set OffName = 'Christopher' + right(Offname,len(Offname)-5)"
sSQL1 = sSQL1 & " Where lower(left(Offname,5)) = 'chris' and len(PersonID) = 0;"

sSQL1 = sSQL1 & " Update USAWSRank.OfclsWithoutIDs"
sSQL1 = sSQL1 & " Set OffName = 'Kenneth' + right(Offname,len(Offname)-5)"
sSQL1 = sSQL1 & " Where lower(left(Offname,5)) = 'kenny' and len(PersonID) = 0;"

sSQL1 = sSQL1 & " Update USAWSRank.OfclsWithoutIDs"
sSQL1 = sSQL1 & " Set OffName = 'Lawrence ' + right(Offname,len(Offname)-5)"
sSQL1 = sSQL1 & " Where lower(left(Offname,5)) = 'larry' and len(PersonID) = 0;"

sSQL1 = sSQL1 & " Update USAWSRank.OfclsWithoutIDs"
sSQL1 = sSQL1 & " Set OffName = 'Sandra ' + right(Offname,len(Offname)-5)"
sSQL1 = sSQL1 & " Where lower(left(Offname,5)) = 'sandy' and len(PersonID) = 0;"

sSQL1 = sSQL1 & " Update USAWSRank.OfclsWithoutIDs"
sSQL1 = sSQL1 & " Set OffName = 'Randall ' + right(Offname,len(Offname)-5)"
sSQL1 = sSQL1 & " Where lower(left(Offname,5)) = 'randy' and len(PersonID) = 0;"

sSQL1 = sSQL1 & " Update USAWSRank.OfclsWithoutIDs"
sSQL1 = sSQL1 & " Set OffName = 'Andrew ' + right(Offname,len(Offname)-5)"
sSQL1 = sSQL1 & " Where lower(left(Offname,5)) = 'andy ' and len(PersonID) = 0;"

sSQL1 = sSQL1 & " Update USAWSRank.OfclsWithoutIDs"
sSQL1 = sSQL1 & " Set OffName = 'William ' + right(Offname,len(Offname)-5)"
sSQL1 = sSQL1 & " Where lower(left(Offname,5)) = 'bill ' and len(PersonID) = 0;"

sSQL1 = sSQL1 & " Update USAWSRank.OfclsWithoutIDs"
sSQL1 = sSQL1 & " Set OffName = 'Bradley ' + right(Offname,len(Offname)-5)"
sSQL1 = sSQL1 & " Where lower(left(Offname,5)) = 'brad ' and len(PersonID) = 0;"

sSQL1 = sSQL1 & " Update USAWSRank.OfclsWithoutIDs"
sSQL1 = sSQL1 & " Set OffName = 'Charlene ' + right(Offname,len(Offname)-5)"
sSQL1 = sSQL1 & " Where lower(left(Offname,5)) = 'char ' and len(PersonID) = 0;"

sSQL1 = sSQL1 & " Update USAWSRank.OfclsWithoutIDs"
sSQL1 = sSQL1 & " Set OffName = 'David ' + right(Offname,len(Offname)-5)"
sSQL1 = sSQL1 & " Where lower(left(Offname,5)) = 'dave ' and len(PersonID) = 0;"

sSQL1 = sSQL1 & " Update USAWSRank.OfclsWithoutIDs"
sSQL1 = sSQL1 & " Set OffName = 'Douglas ' + right(Offname,len(Offname)-5)"
sSQL1 = sSQL1 & " Where lower(left(Offname,5)) = 'doug ' and len(PersonID) = 0;"

sSQL1 = sSQL1 & " Update USAWSRank.OfclsWithoutIDs"
sSQL1 = sSQL1 & " Set OffName = 'Jeffrey ' + right(Offname,len(Offname)-5)"
sSQL1 = sSQL1 & " Where lower(left(Offname,5)) = 'jeff ' and len(PersonID) = 0;"

sSQL2 = " Update USAWSRank.OfclsWithoutIDs"
sSQL2 = sSQL2 & " Set OffName = 'Matthew ' + right(Offname,len(Offname)-5)"
sSQL2 = sSQL2 & " Where lower(left(Offname,5)) = 'matt ' and len(PersonID) = 0;"

sSQL2 = sSQL2 & " Update USAWSRank.OfclsWithoutIDs"
sSQL2 = sSQL2 & " Set OffName = 'Michael ' + right(Offname,len(Offname)-5)"
sSQL2 = sSQL2 & " Where lower(left(Offname,5)) = 'mike ' and len(PersonID) = 0;"

sSQL2 = sSQL2 & " Update USAWSRank.OfclsWithoutIDs"
sSQL2 = sSQL2 & " Set OffName = 'Robert ' + right(Offname,len(Offname)-4)"
sSQL2 = sSQL2 & " Where lower(left(Offname,4)) = 'bob ' and len(PersonID) = 0;"

sSQL2 = sSQL2 & " Update USAWSRank.OfclsWithoutIDs"
sSQL2 = sSQL2 & " Set OffName = 'Daniel ' + right(Offname,len(Offname)-4)"
sSQL2 = sSQL2 & " Where lower(left(Offname,4)) = 'dan ' and len(PersonID) = 0;"

sSQL2 = sSQL2 & " Update USAWSRank.OfclsWithoutIDs"
sSQL2 = sSQL2 & " Set OffName = 'Janet ' + right(Offname,len(Offname)-4)"
sSQL2 = sSQL2 & " Where lower(left(Offname,4)) = 'jan ' and len(PersonID) = 0;"

sSQL2 = sSQL2 & " Update USAWSRank.OfclsWithoutIDs"
sSQL2 = sSQL2 & " Set OffName = 'James ' + right(Offname,len(Offname)-4)"
sSQL2 = sSQL2 & " Where lower(left(Offname,4)) = 'jim ' and len(PersonID) = 0;"

sSQL2 = sSQL2 & " Update USAWSRank.OfclsWithoutIDs"
sSQL2 = sSQL2 & " Set OffName = 'Joseph ' + right(Offname,len(Offname)-4)"
sSQL2 = sSQL2 & " Where lower(left(Offname,4)) = 'joe ' and len(PersonID) = 0;"

sSQL2 = sSQL2 & " Update USAWSRank.OfclsWithoutIDs"
sSQL2 = sSQL2 & " Set OffName = 'Jonathan ' + right(Offname,len(Offname)-4)"
sSQL2 = sSQL2 & " Where lower(left(Offname,4)) = 'jon ' and len(PersonID) = 0;"

sSQL2 = sSQL2 & " Update USAWSRank.OfclsWithoutIDs"
sSQL2 = sSQL2 & " Set OffName = 'Kenneth ' + right(Offname,len(Offname)-4)"
sSQL2 = sSQL2 & " Where lower(left(Offname,4)) = 'ken ' and len(PersonID) = 0;"

sSQL2 = sSQL2 & " Update USAWSRank.OfclsWithoutIDs"
sSQL2 = sSQL2 & " Set OffName = 'Lester ' + right(Offname,len(Offname)-4)"
sSQL2 = sSQL2 & " Where lower(left(Offname,4)) = 'les ' and len(PersonID) = 0;"

sSQL2 = sSQL2 & " Update USAWSRank.OfclsWithoutIDs"
sSQL2 = sSQL2 & " Set OffName = 'Robert ' + right(Offname,len(Offname)-4)"
sSQL2 = sSQL2 & " Where lower(left(Offname,4)) = 'rob ' and len(PersonID) = 0;"

sSQL2 = sSQL2 & " Update USAWSRank.OfclsWithoutIDs"
sSQL2 = sSQL2 & " Set OffName = 'Susan ' + right(Offname,len(Offname)-4)"
sSQL2 = sSQL2 & " Where lower(left(Offname,4)) = 'sue ' and len(PersonID) = 0;"

sSQL2 = sSQL2 & " Update USAWSRank.OfclsWithoutIDs"
sSQL2 = sSQL2 & " Set OffName = 'Thomas ' + right(Offname,len(Offname)-4)"
sSQL2 = sSQL2 & " Where lower(left(Offname,4)) = 'tom ' and len(PersonID) = 0;"

OpenCon: con.execute ( sSQL1 & sSQL2 ): CloseCon

END SUB



'	----------------------
SUB	UpdateSanctionTable
'	----------------------

'	This subroutine runs a series of updates to the "Registration" table,
'	replacing names and person ID's, keying on Tournament ID and Position Code.
'	Then at conclusion, counts number of updates applied, then deletes the rows
'	from the exceptions table, then finally counts the remaining exceptions.

sSQL1 = " Update RT Set CJudgePID = OT.PersonID, CJudge = left(OT.OffName,20)"
sSQL1 = sSQL1 & " from sanctions.dbo.registration RT, USAWSRank.OfclsWithoutIDs OT"
sSQL1 = sSQL1 & " where OT.OffCode = 'CJdg' and len(OT.PersonID) > 0"
sSQL1 = sSQL1 & " and OT.TournAppID = RT.TournAppID;"

sSQL1 = sSQL1 & " Update RT Set CScorePID = OT.PersonID, CScorer = left(OT.OffName,20)"
sSQL1 = sSQL1 & " from sanctions.dbo.registration RT, USAWSRank.OfclsWithoutIDs OT"
sSQL1 = sSQL1 & " where OT.OffCode = 'CScr' and len(OT.PersonID) > 0"
sSQL1 = sSQL1 & " and OT.TournAppID = RT.TournAppID;"

sSQL1 = sSQL1 & " Update RT Set CDriverPID = OT.PersonID, CDriver = left(OT.OffName,20)"
sSQL1 = sSQL1 & " from sanctions.dbo.registration RT, USAWSRank.OfclsWithoutIDs OT"
sSQL1 = sSQL1 & " where OT.OffCode = 'CDrv' and len(OT.PersonID) > 0"
sSQL1 = sSQL1 & " and OT.TournAppID = RT.TournAppID;"

sSQL1 = sSQL1 & " Update RT Set CSafPID = OT.PersonID, CSafety = left(OT.OffName,20)"
sSQL1 = sSQL1 & " from sanctions.dbo.registration RT, USAWSRank.OfclsWithoutIDs OT"
sSQL1 = sSQL1 & " where OT.OffCode = 'CSft' and len(OT.PersonID) > 0"
sSQL1 = sSQL1 & " and OT.TournAppID = RT.TournAppID;"

sSQL1 = sSQL1 & " Update RT Set TechCPID = OT.PersonID, TechCont = left(OT.OffName,20)"
sSQL1 = sSQL1 & " from sanctions.dbo.registration RT, USAWSRank.OfclsWithoutIDs OT"
sSQL1 = sSQL1 & " where OT.OffCode = 'CTch' and len(OT.PersonID) > 0"
sSQL1 = sSQL1 & " and OT.TournAppID = RT.TournAppID;"

sSQL1 = sSQL1 & " Update RT Set Ap1JPID = OT.PersonID, Ap1Judge = left(OT.OffName,20)"
sSQL1 = sSQL1 & " from sanctions.dbo.registration RT, USAWSRank.OfclsWithoutIDs OT"
sSQL1 = sSQL1 & " where OT.OffCode = 'APJ1' and len(OT.PersonID) > 0"
sSQL1 = sSQL1 & " and OT.TournAppID = RT.TournAppID;"

sSQL1 = sSQL1 & " Update RT Set Ap2JPID = OT.PersonID, Ap2Judge = left(OT.OffName,20)"
sSQL1 = sSQL1 & " from sanctions.dbo.registration RT, USAWSRank.OfclsWithoutIDs OT"
sSQL1 = sSQL1 & " where OT.OffCode = 'APJ2' and len(OT.PersonID) > 0"
sSQL1 = sSQL1 & " and OT.TournAppID = RT.TournAppID;"

sSQL2 = " Update RT Set Ap3JPID = OT.PersonID, Ap3Judge = left(OT.OffName,20)"
sSQL2 = sSQL2 & " from sanctions.dbo.registration RT, USAWSRank.OfclsWithoutIDs OT"
sSQL2 = sSQL2 & " where OT.OffCode = 'APJ3' and len(OT.PersonID) > 0"
sSQL2 = sSQL2 & " and OT.TournAppID = RT.TournAppID;"

sSQL2 = sSQL2 & " Update RT Set Ap4JPID = OT.PersonID, Ap4Judge = left(OT.OffName,20)"
sSQL2 = sSQL2 & " from sanctions.dbo.registration RT, USAWSRank.OfclsWithoutIDs OT"
sSQL2 = sSQL2 & " where OT.OffCode = 'APJ4' and len(OT.PersonID) > 0"
sSQL2 = sSQL2 & " and OT.TournAppID = RT.TournAppID;"

sSQL2 = sSQL2 & " Update RT Set Ap5JPID = OT.PersonID, Ap5Judge = left(OT.OffName,20)"
sSQL2 = sSQL2 & " from sanctions.dbo.registration RT, USAWSRank.OfclsWithoutIDs OT"
sSQL2 = sSQL2 & " where OT.OffCode = 'APJ5' and len(OT.PersonID) > 0"
sSQL2 = sSQL2 & " and OT.TournAppID = RT.TournAppID;"

sSQL2 = sSQL2 & " Update RT Set Ap1DrPID = OT.PersonID, Ap1Driver = left(OT.OffName,20)"
sSQL2 = sSQL2 & " from sanctions.dbo.registration RT, USAWSRank.OfclsWithoutIDs OT"
sSQL2 = sSQL2 & " where OT.OffCode = 'APD1' and len(OT.PersonID) > 0"
sSQL2 = sSQL2 & " and OT.TournAppID = RT.TournAppID;"

sSQL2 = sSQL2 & " Update RT Set Ap1SPID = OT.PersonID, Ap1Scorer = left(OT.OffName,20)"
sSQL2 = sSQL2 & " from sanctions.dbo.registration RT, USAWSRank.OfclsWithoutIDs OT"
sSQL2 = sSQL2 & " where OT.OffCode = 'APS1' and len(OT.PersonID) > 0"
sSQL2 = sSQL2 & " and OT.TournAppID = RT.TournAppID;"

sSQL2 = sSQL2 & " Update RT Set Ap2SPID = OT.PersonID, Ap2Scorer = left(OT.OffName,20)"
sSQL2 = sSQL2 & " from sanctions.dbo.registration RT, USAWSRank.OfclsWithoutIDs OT"
sSQL2 = sSQL2 & " where OT.OffCode = 'APS2' and len(OT.PersonID) > 0"
sSQL2 = sSQL2 & " and OT.TournAppID = RT.TournAppID;"

sSQL2 = sSQL2 & " Update RT Set Ap3SPID = OT.PersonID, Ap3Scorer = left(OT.OffName,20)"
sSQL2 = sSQL2 & " from sanctions.dbo.registration RT, USAWSRank.OfclsWithoutIDs OT"
sSQL2 = sSQL2 & " where OT.OffCode = 'APS3' and len(OT.PersonID) > 0"
sSQL2 = sSQL2 & " and OT.TournAppID = RT.TournAppID;"

sSQL2 = sSQL2 & " Update RT Set PanAmPID = OT.PersonID, PanAmJudge = left(OT.OffName,20)"
sSQL2 = sSQL2 & " from sanctions.dbo.registration RT, USAWSRank.OfclsWithoutIDs OT"
sSQL2 = sSQL2 & " where OT.OffCode = 'PAMJ' and len(OT.PersonID) > 0"
sSQL2 = sSQL2 & " and OT.TournAppID = RT.TournAppID;"

OpenCon: con.execute ( sSQL1 & sSQL2 ): CloseCon

sSQL1 = "Select count(*) as Kount from USAWSRank.OfclsWithoutIDs where len(PersonID) > 0"
objRS.open sSQL1, SConnectionToTRATable, 3, 3
NumFix = objRS("Kount"): objRS.close

sSQL1 = "Delete from USAWSRank.OfclsWithoutIDs where len(PersonID) > 0"
OpenCon: con.execute ( sSQL1 & sSQL2 ): CloseCon

END SUB



'	-----------------
SUB	ListExceptions
'	-----------------

'	This subroutine builds an exception list as an HTML table,
'	with each exception being offered with an "Edit" link.

WriteIndexPageHeader

sSQL1 = "Select count(*) as Kount from USAWSRank.OfclsWithoutIDs"
objRS.open sSQL1, SConnectionToTRATable, 3, 3
NumLeft = objRS("Kount"): objRS.close

sSQL1 = " Select * from USAWSRank.OfclsWithoutIDs where"
sSQL1 = sSQL1 & " len(PersonID) = 0 order by OffName"

objRS.open sSQL1, SConnectionToTRATable, 3, 3

IF NOT objRS.EOF THEN

	%><br>&nbsp;<br>
	
	<Table class="innertable" width=80% align=center>

	<TR><TD Colspan=4><center><font size="2"><b><br><%=NumFix%> Exceptions fixed, 
	<%=NumLeft%> Exceptions remaining<br> after scanning the table of unique known 
	officials. <br>&nbsp;<br>The following Officials names 
	appear as Appointed Official positions<br> in the Sanctions tables, without Person 
	ID's.&nbsp; Select an Exception<br> from the list below, then either revise the 
	spelling of that name<br> (after which I will then search again), or else you can 
	supply the<br> actual Person ID value that belongs to that official.&nbsp; Officials
	with<br> non-unique names can only be identified by the latter method.
	<br>&nbsp;</b></font></center></TD></TR>
	
	<tr><td Colspan=4>&nbsp;</td></tr>
	
	<tr>
		<th><font size="2" color="#FFFFFF"><center> Action </font></center></td>
		<th><center><font size="2" color="#FFFFFF"> Name from Sanction </font></center></td>
		<th><center><font size="2" color="#FFFFFF"> Tournament </font></center></td>
		<th><center><font size="2" color="#FFFFFF"> Position </font></center></td>
	</tr>
	
	<% DO UNTIL objRS.EOF %>

	<tr>
		<td><center><font size="2"> <A HREF="/rankings/OffNameRepair.asp?Process=EditOff&OffName=<%=objRS("OffName")%>">Select</A> </font></center></TD>
		<td><center><font size="2"> <%=objRS("OffName")%> </font></center></td>
		<td><center><font size="2"> <%=objRS("TournAppID")%> </font></center></td>
		<td><center><font size="2"> <%=objRS("OffCode")%> </font></center></td>
	</tr>	
	
	<%
	
	objRS.MoveNext
	
	LOOP
	
	%>
	
	</table> <br>&nbsp;<br>
	
	<%
	
ELSE

	%>
	
	<br> All Exceptions now Cleared -- select another function from the Nav Menu to the left.<br>
	
	<%

END IF

WriteIndexPageFooter

END SUB

%>