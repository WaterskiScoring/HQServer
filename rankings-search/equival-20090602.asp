<% Option Explicit %>
<!--#include virtual="/rankings/settingsHQ.asp"-->


<%

Response.Buffer = False
Server.ScriptTimeout = 1200 

%><html><head><title>Please Wait...</title></head>
<body>
<TABLE align=center BGCOLOR="#000000" BORDER=1 BORDERCOLOR="#000000"	CELLPADDING=0 CELLSPACING=0 HEIGHT=200 WIDTH=350>
<TR>
<TD WIDTH="100%" HEIGHT="100%" BGCOLOR="#CCCCCC" ALIGN="CENTER" VALIGN="MIDDLE">
<br>&nbsp;<br>
<FONT FACE="Helvetica,Verdana,Arial" SIZE=3 COLOR="#000066">
<B>Nightly Update / Recalculations.<br>&nbsp;<br>
This will take a few minutes.<br>&nbsp;<br>  
Please wait...</B></FONT> <br>&nbsp;<br>

<div id="ProgBar" style="font-family:Verdana; font-size=9pt;">
<TABLE style="color:red;" HEIGHT="16" Border=1><TR><TD BGCOLOR=RED ID=statuspic></TD></TR></TABLE><BR>
</div> 

<script language="Javascript">var progBarWidth=350;</script>

<SCRIPT LANGUAGE=JScript RUNAT=Server>
function y2k(number)   {
   return (number < 1000) ? number + 1900 : number;
                     }
function milliDif()   {
   var d = new Date();
      return d.getTime()
                  }
                  
function elapsedpretty(parm1)
{
  var elapsedsecs = 0
  var elapsedmins = 0
  
  elapsedsecs=Math.floor(parm1/1000)
  parm1=parm1%1000
  
  elapsedmins=Math.floor(elapsedsecs/60)
  elapsedsecs=elapsedsecs%60
  
  
elapsedpretty=elapsedmins + " minute"
if(elapsedmins!=1)
       elapsedpretty=elapsedpretty+"s"
  
elapsedpretty = elapsedpretty+" " + elapsedsecs+" second"
if(elapsedsecs!=1)
       elapsedpretty=elapsedpretty+"s"
  
elapsedpretty = elapsedpretty+ " "+parm1+" millisecond"
if(parm1!=1)
       elapsedpretty=elapsedpretty+"s"
  
  return elapsedpretty;
}  
</script>

<FONT FACE="Arial,Vendana,Helvetica" SIZE=1>The progress bar above will expand as members are processed.<br>
At 100%, it will stretch the entire width of this box.</font><br>
</td></tr></table><br><FONT FACE="Arial,Vendana,Helvetica" SIZE=2><%


' --------------------------------
   SUB ShowProgress(nPctComplete)
' --------------------------------

Response.Write "<SCR" & "IPT LANGUAGE=""JavaScript"">" & vbCrlf
Response.Write "statuspic.width = Math.ceil(" & nPctComplete & " * progBarWidth);" & vbCrlf
Response.Write "</SCR" & "IPT>"

END SUB


' --------------------------------
   SUB FinishProgress
' --------------------------------

Response.Write "<SCR" & "IPT LANGUAGE=""JavaScript"">" & vbCrlf
Response.Write "ProgBar.style.visibility ='hidden';" & vbCrLf
Response.Write "</SCR" & "IPT>"

END SUB



' -------------------------------   START OF MAIN PROGRAM ---------------------------------------------

Dim i, j
Dim tBeginDate, DupMemList
Dim nProcessedSoFar, nTotalMembers, tempCounter, tempvar, TempSum, TempLen, TempPtr
Dim strHTML, sSkiYearBegin, sSkiYearEnd, sProcessingYear, sPrevYear, sSkiYearName

Dim sSQL
Dim tUpAge, tBirthdate, tLatestBirthYear, tSkiAge, AgedOut
Dim R_Ski, R_PLC, N_PLC

Dim timeTHEN, timeNow
Dim EMailToWho, myMail, EmailErrors

' Overall Stuff
Dim TempMemberID, TempTeam, TempEvent, TempTourID, TempDiv, TempScore, TempAdd
Dim TempOverEvts, TempOverEvtsReq, TempOATot, TempDivOrig, InDivOrig
Dim Slalom1, Slalom2, Slalom3, Trick1, Trick2, Trick3
Dim Jump1, Jump2, Jump3, Class1, Class2, Class3
Dim S_Round1, J_Round1, T_Round1, S_Score1, J_Score1, T_Score1
Dim S_Round2, J_Round2, T_Round2, S_Score2, J_Score2, T_Score2
Dim S_Round3, J_Round3, T_Round3, S_Score3, J_Score3, T_Score3
Dim nScoC, nScoR, TotScore, TempRating, TempTName, RankExplain

' Ranking Penalty Table as Function of C vs ELR score count
Dim tPenalty(3,3)
FOR nScoC=0 to 3: FOR nScoR=0 to 3: tPenalty(nScoC,nScoR)=0: NEXT: NEXT
tPenalty(0,1)=5: tPenalty(1,0)=10: tPenalty(1,1)=2.5: tPenalty(2,0)=5

' Operational Controls
Dim RunOverride, RunEquivScore, RunOvrllScore, RunOvrllRanks, RunEventRanks, RunLevelLogic

Dim tSYEndDate, tBirthYear

' Membership Update/Merge Controls and variables
Dim TempHQPID, TempLclPID, PIDwCheckDigit, HQConnect, HQrs
Dim nHQExt, nLocal, nInserts, nUpdates, nDeletes, LastHQPID
Dim nScoUpdts, nOvrUpdts, nRnkUpdts, nOffUpdts
Dim nSlices, iSlice, nTotal, nSoFar, VBCrLf, StartTime

VBCrLf = Chr(13) & Chr(10): StartTime = Time(): DupMemList = ""
nHQExt = 0: nLocal = 0: nInserts = 0: nUpdates = 0: nDeletes = 0
LastHQPID = 0: nScoUpdts = 0: nOvrUpdts = 0: nRnkUpdts = 0: nOffUpdts = 0

' ---------------------------- TIME KEEPING FUNCTIONS -----------------------------

timeTHEN = milliDif()
WriteLog(date() &"  "& time() &"  Begin Nightly Member Update / Ranking Recalculation Process.")

OpenCon
Set rs = Server.CreateObject("ADODB.recordset")

' Response.write("Equival value = " & Request("Equival") & "<br>")

EMailToWho = "<clark.dave@comcast.net>; <letsgoski@embarqmail.com>; <mhanson@usawaterski.org>"
EmailErrors = 0


IF Request("Equival") <> "ReCalc" THEN

' ************** Beginning of Member Extract Update Conditional Section ***************

' Overall Update process runs in small slices, based on Mod function against PersonID values.

Set HQConnect = CreateObject("ADODB.Connection")
HQConnect.Open Application("HQSQLConn")

' Below is Usage to read local database connection as in Admin/DisplayOneMember.asp

' Dim objConn
' Set objConn = Server.CreateObject("ADODB.Connection")
' objConn.Open Application("WaterSkiConn")
' RS.ActiveConnection = objConn
' sSQL = "Statement ...."
' RS.open sSQL

nSlices = 10
FOR iSlice = 0 to nSlices - 1

' Begin by pulling a consolidated extract from HQ server, using a
' complex multi-table Join.  Order by Person ID for subsequent merge.

Response.write("Member Update: Slice " & iSlice & " -- Querying HQ Server ...")

sSQL = "SELECT PT.[Person ID] as PersonID, PT.[Name Prefix] as NamePrefix,"
sSQL = sSQL & " PT.[First Name] as FirstName, PT.[Middle] as MiddleName,"
sSQL = sSQL & " PT.[Last Name] as LastName, PT.[Name Suffix] as NameSuffix,"
sSQL = sSQL & " PT.SSN, PT.[Company Name] as CompanyName,"
sSQL = sSQL & " Substring(PT.Website,1,100) as Website, PT.Email, PT.MailPref,"
sSQL = sSQL & " PT.[Birth Date] as BirthDate, PT.Sex, D1.[Division Code] as DivisionCode1,"
sSQL = sSQL & " D2.[Division Code] as DivisionCode2, PT.[Federation Code] as FederationCode,"
sSQL = sSQL & " MT.MemberTypeID, PA.Phone, PA.Extension, PA.Fax,"
sSQL = sSQL & " PA.[Business Phone] as BusinessPhone,"
sSQL = sSQL & " PA.[Business Extension] as BusinessExtension,"
sSQL = sSQL & " PA.[Mobile Phone] as MobilePhone, PA.Address1, PA.Address2,"
sSQL = sSQL & " PA.City, PA.State, PA.Zip, PA.[Country ID] as CountryID,"
sSQL = sSQL & " MH.[Membership Type Code] as MembershipTypeCode,"
sSQL = sSQL & " MH.EffectiveFrom, MH.EffectiveTo,"
sSQL = sSQL & " Case when PT.DoNotEMail=1 then '1' else '0' end as DoNotEMail, PA.Region,"
sSQL = sSQL & " PT.[Member Since] as MemberSince, PT.[Date Updated] as DateUpdated,"
sSQL = sSQL & " Case when PT.DoNotCall=1 then '1' else '0' end as DoNotCall,"
sSQL = sSQL & " Left(MT.[Membership Type Description],10) as MembershipType,"
sSQL = sSQL & " Case when PT.Deceased=1 then '1' else '0' end as Deceased"
sSQL = sSQL & " FROM	Waterski.dbo.tblPeople PT, Waterski.dbo.tblPeopleAddresses PA,"
sSQL = sSQL & " Waterski.dbo.[Membership History] MH, (Select [Person ID] as PersonID,"
sSQL = sSQL & " Max(EffectiveTo) as MaxEffTo From Waterski.dbo.[Membership History]"
sSQL = sSQL & " group by [Person ID]) ME, Waterski.dbo.tblMembershipTypeCodes MT,"
sSQL = sSQL & " Waterski.dbo.tblDivisionCodes D1, Waterski.dbo.tblDivisionCodes D2"
sSQL = sSQL & " WHERE PA.[Person ID] = PT.[Person ID] AND PA.[Primary] = 1"
sSQL = sSQL & " AND PT.[Person ID] = MH.[Person ID] AND MH.[Person ID] = ME.PersonID"
sSQL = sSQL & " AND MH.EffectiveTo = ME.MaxEffTo"
sSQL = sSQL & " AND MH.[Membership Type Code] = MT.[Membership Type Code]"
sSQL = sSQL & " AND MH.PrimaryDivisionCodeID = D1.DivisionCodeID"
sSQL = sSQL & " AND MH.SecondaryDivisionCodeID = D2.DivisionCodeID"
' sSQL = sSQL & " AND (DateAdd(DD,-30,MH.EffectiveFrom) >= GetDate()": ' Old Daily Update Extract
' sSQL = sSQL & " OR DateAdd(DD,-335,MH.EffectiveTo) >= GetDate()":    ' Old Daily Update Extract
' sSQL = sSQL & " OR DateAdd(DD,+30,PT.[Date Updated]) >= GetDate())": ' Old Daily Update Extract
sSQL = sSQL & " AND DateAdd(MM,30,MH.EffectiveTo) >= GetDate()":           ' New Daily Slice Update
sSQL = sSQL & " AND (PT.[Person ID] % " & nSlices & " = " & iSlice & ")":  ' New Daily Slice Update
' sSQL = sSQL & " AND PT.Deceased = 0"
sSQL = sSQL & " ORDER BY PT.[Person ID]"

WriteDebugSQL(sSQL)

Set HQrs = HQConnect.Execute(sSql)


' Get first HQ Recordset Person ID, or 99999999 if empty.

IF HQrs.eof THEN
	TempHQPID = 99999999
	nTotal = 0
ELSE
	tempvar = HQrs.getrows()
	nTotal = ubound(tempvar,2)
	HQrs.MoveFirst
	TempHQPID = HQrs("PersonID")
	nHQExt = nHQExt + 1
END IF


' Now we pull some key columns from the local table, also 
' keyed by Person ID, so that we can merge with HQ Extract.

Response.write(" Querying Local Table ...")

sSQL = "SELECT PersonID, DateUpdated, EffectiveTo"
sSQL = sSQL & " FROM " & MemberTableName
sSQL = sSQL & " WHERE (PersonID % " & nSlices & " = " & iSlice & ")":  ' New Daily Slice Update
sSQL = sSQL & " ORDER BY PersonID"

WriteDebugSQL(sSQL)

rs.open sSQL, sConnectionToTRATable, 3, 3


' Below is Usage to read local database connection as in Admin/DisplayOneMember.asp

' Dim objConn
' Set objConn = Server.CreateObject("ADODB.Connection")
' objConn.Open Application("WaterSkiConn")
' RS.ActiveConnection = objConn
' sSQL = "Statement ...."
' RS.open sSQL

' Get first Local Recordset Person ID, or 99999999 if empty.


IF rs.eof THEN
	TempLclPID = 99999999
ELSE
	rs.MoveFirst
	TempLclPID = rs("PersonID")
	nLocal = nLocal + 1
END IF

nSoFar = 0

' Now we begin a merge loop between the HQ extract and the local table, for this slice.

Response.write(" Merging/Updating ...")

DO UNTIL HQrs.eof AND rs.eof

IF (nSoFar mod 17 = 9) THEN ShowProgress (nSoFar / nTotal)

IF TempHQPID < TempLclPID THEN
	
	' Where incoming Person ID is new (lower), then we would do an insert.
	' But only if this is NOT a duplicate in the incoming HQ extract ...
	
	IF TempHQPID > LastHQPID THEN

		' Yes this is a new Person ID -- so create the insert Query

		sSQL = "INSERT INTO " & MemberTableName & " VALUES ('"
		sSQL = sSQL & TempHQPID & "','" & PersonIDwChkDgt(TempHQPID) & "','"
		sSQL = sSQL & SQLClean(HQrs("NamePrefix")) & "','"
		sSQL = sSQL & SQLClean(HQrs("FirstName")) & "','"
		sSQL = sSQL & SQLClean(HQrs("MiddleName")) & "','"
		sSQL = sSQL & SQLClean(HQrs("LastName")) & "','"
		sSQL = sSQL & SQLClean(HQrs("NameSuffix")) & "','"
		sSQL = sSQL & SQLClean(HQrs("SSN")) & "','"
		sSQL = sSQL & SQLClean(HQrs("CompanyName")) & "','"
		sSQL = sSQL & SQLClean(HQrs("Website")) & "','"
		sSQL = sSQL & SQLClean(HQrs("Email")) & "','"
		sSQL = sSQL & SQLClean(HQrs("MailPref")) & "','"
		sSQL = sSQL & SQLClean(HQrs("BirthDate")) & "','"
		sSQL = sSQL & SQLClean(HQrs("Sex")) & "','"
		sSQL = sSQL & SQLClean(HQrs("DivisionCode1")) & "','"
		sSQL = sSQL & SQLClean(HQrs("DivisionCode2")) & "','"
		sSQL = sSQL & SQLClean(HQrs("FederationCode")) & "','"
		sSQL = sSQL & SQLClean(HQrs("MemberTypeID")) & "','"
		sSQL = sSQL & SQLClean(HQrs("Phone")) & "','"
		sSQL = sSQL & SQLClean(HQrs("Extension")) & "','"
		sSQL = sSQL & SQLClean(HQrs("Fax")) & "','"
		sSQL = sSQL & SQLClean(HQrs("BusinessPhone")) & "','"
		sSQL = sSQL & SQLClean(HQrs("BusinessExtension")) & "','"
		sSQL = sSQL & SQLClean(HQrs("MobilePhone")) & "','"
		sSQL = sSQL & SQLClean(HQrs("Address1")) & "','"
		sSQL = sSQL & SQLClean(HQrs("Address2")) & "','"
		sSQL = sSQL & SQLClean(HQrs("City")) & "','"
		sSQL = sSQL & SQLClean(HQrs("State")) & "','"
		sSQL = sSQL & SQLClean(HQrs("Zip")) & "','"
		sSQL = sSQL & SQLClean(HQrs("CountryID")) & "','"
		sSQL = sSQL & SQLClean(HQrs("MembershipTypeCode")) & "','"
		sSQL = sSQL & SQLClean(HQrs("EffectiveFrom")) & "','"
		sSQL = sSQL & SQLClean(HQrs("EffectiveTo")) & "','"
		sSQL = sSQL & SQLClean(HQrs("DoNotEMail")) & "','"
		sSQL = sSQL & SQLClean(HQrs("Region")) & "','"
		sSQL = sSQL & SQLClean(HQrs("MemberSince")) & "','"
		sSQL = sSQL & SQLClean(HQrs("DateUpdated")) & "','"
		sSQL = sSQL & SQLClean(HQrs("DoNotCall")) & "','"
		sSQL = sSQL & SQLClean(HQrs("MembershipType")) & "','"
		sSQL = sSQL & SQLClean(HQrs("Deceased")) & "')"

		' WriteDebugSQL(sSQL)

		' Invoke the Insert, and then tally number of Inserts,
		' and then save this inserted HQ PID to check later.

		Con.Execute(sSQL)
		nInserts = nInserts + 1
		LastHQPID = TempHQPID

	ELSE
		
		' Note incoming duplicates for Recap report.

		IF Len(DupMemList) < 100 THEN
			DupMemList = DupMemList & LastHQPID & " "
		END IF

'		WriteLog(date() &"  "& time() & " PersonID " & LastHQPID & " Duplicate in HQ Extract.")

	END IF
	
	' Now finally advance the incoming HQ Extract recordset.

	HQrs.moveNEXT
	IF HQrs.eof THEN TempHQPID = 99999999 ELSE TempHQPID = HQrs("PersonID"): nHQExt = nHQExt + 1
	nSoFar = nSoFar + 1

ELSEIF TempHQPID = TempLclPID THEN

	' Where Incoming Person ID matches local, then Update existing member details
	' with new data from HQ extract.  Then advance both record sets to next row.

	sSQL = "UPDATE " & MemberTableName & " SET"
	sSQL = sSQL & " NamePrefix = '" & SQLClean(HQrs("NamePrefix")) & "',"
	sSQL = sSQL & " FirstName = '" & SQLClean(HQrs("FirstName")) & "',"
	sSQL = sSQL & " MiddleName = '" & SQLClean(HQrs("MiddleName")) & "',"
	sSQL = sSQL & " LastName = '" & SQLClean(HQrs("LastName")) & "',"
	sSQL = sSQL & " NameSuffix = '" & SQLClean(HQrs("NameSuffix")) & "',"
	sSQL = sSQL & " SSN = '" & SQLClean(HQrs("SSN")) & "',"
	sSQL = sSQL & " CompanyName = '" & SQLClean(HQrs("CompanyName")) & "',"
	sSQL = sSQL & " Website = '" & SQLClean(HQrs("Website")) & "',"
	sSQL = sSQL & " Email = '" & SQLClean(HQrs("Email")) & "',"
	sSQL = sSQL & " MailPref = '" & SQLClean(HQrs("MailPref")) & "',"
	sSQL = sSQL & " BirthDate = '" & SQLClean(HQrs("BirthDate")) & "',"
	sSQL = sSQL & " Sex = '" & SQLClean(HQrs("Sex")) & "',"
	sSQL = sSQL & " DivisionCode1 = '" & SQLClean(HQrs("DivisionCode1")) & "',"
	sSQL = sSQL & " DivisionCode2 = '" & SQLClean(HQrs("DivisionCode2")) & "',"
	sSQL = sSQL & " FederationCode = '" & SQLClean(HQrs("FederationCode")) & "',"
	sSQL = sSQL & " MemberTypeID = '" & SQLClean(HQrs("MemberTypeID")) & "',"
	sSQL = sSQL & " Phone = '" & SQLClean(HQrs("Phone")) & "',"
	sSQL = sSQL & " Extension = '" & SQLClean(HQrs("Extension")) & "',"
	sSQL = sSQL & " Fax = '" & SQLClean(HQrs("Fax")) & "',"
	sSQL = sSQL & " BusinessPhone = '" & SQLClean(HQrs("BusinessPhone")) & "',"
	sSQL = sSQL & " BusinessExtension = '" & SQLClean(HQrs("BusinessExtension")) & "',"
	sSQL = sSQL & " MobilePhone = '" & SQLClean(HQrs("MobilePhone")) & "',"
	sSQL = sSQL & " Address1 = '" & SQLClean(HQrs("Address1")) & "',"
	sSQL = sSQL & " Address2 = '" & SQLClean(HQrs("Address2")) & "',"
	sSQL = sSQL & " City = '" & SQLClean(HQrs("City")) & "',"
	sSQL = sSQL & " State = '" & SQLClean(HQrs("State")) & "',"
	sSQL = sSQL & " Zip = '" & SQLClean(HQrs("Zip")) & "',"
	sSQL = sSQL & " CountryID = '" & SQLClean(HQrs("CountryID")) & "',"
	sSQL = sSQL & " MembershipTypeCode = '" & SQLClean(HQrs("MembershipTypeCode")) & "',"
	sSQL = sSQL & " EffectiveFrom = '" & SQLClean(HQrs("EffectiveFrom")) & "',"
	sSQL = sSQL & " EffectiveTo = '" & SQLClean(HQrs("EffectiveTo")) & "',"
	sSQL = sSQL & " DoNotEMail = '" & SQLClean(HQrs("DoNotEMail")) & "',"
	sSQL = sSQL & " Region = '" & SQLClean(HQrs("Region")) & "',"
	sSQL = sSQL & " MemberSince = '" & SQLClean(HQrs("MemberSince")) & "',"
	sSQL = sSQL & " DateUpdated = '" & SQLClean(HQrs("DateUpdated")) & "',"
	sSQL = sSQL & " DoNotCall = '" & SQLClean(HQrs("DoNotCall")) & "',"
	sSQL = sSQL & " MembershipType = '" & SQLClean(HQrs("MembershipType")) & "',"
	sSQL = sSQL & " Deceased = '" & SQLClean(HQrs("Deceased")) & "'"
	sSQL = sSQL & " WHERE PersonID = " & TempLclPID
		
	' WriteDebugSQL(sSQL)

	' Invoke the update and then tally number of Member rows updated

	Con.Execute(sSQL)
	nUpdates = nUpdates + 1
	LastHQPID = TempHQPID
		
	' Now advance both Extract and Local recordsets

	HQrs.moveNEXT
	IF HQrs.eof THEN TempHQPID = 99999999 ELSE TempHQPID = HQrs("PersonID"): nHQExt = nHQExt + 1
	nSoFar = nSoFar + 1
	rs.moveNEXT
	IF rs.eof THEN TempLclPID = 99999999 ELSE TempLclPID = rs("PersonID"): nLocal = nLocal + 1

ELSE

	' Where Local Person ID is Lower, then that Local Person ID is no longer on the HQ server,
	' so we delete this Person ID from the local table, and then tally the Delete counter.
	' Then finally we advance the Local recordset to the next row.

	sSQL = "DELETE FROM " & MemberTableName & " WHERE PersonID = " & TempLclPID
	Con.Execute(sSQL)
	nDeletes = nDeletes + 1

	rs.moveNEXT
	IF rs.eof THEN TempLclPID = 99999999 ELSE TempLclPID = rs("PersonID"): nLocal = nLocal + 1
		
END IF

LOOP

' End of Membership Update Loop for current slice.  Close record sets and report time.

rs.Close
HQrs.Close

Response.write(" DONE at " & Time() & "<br>")

NEXT

' End of Member Update Loop over Slices.  Now spit out an update recap report.

Response.write("&nbsp;<br>" & Formatnumber(nHQExt,0) & " Member rows supplied from HQ Server<br>")
Response.write(Formatnumber(nLocal,0) & " Member rows found in Local Server Table<br>")
Response.write(Formatnumber(nUpdates,0) & " Member rows updated with new Data<br>")
Response.write(Formatnumber(nInserts,0) & " New Member rows added<br>")
Response.write(Formatnumber(nDeletes,0) & " Old Member rows deleted<br>")
Response.write(Formatnumber(nLocal+nInserts-nDeletes,0) & " Member rows now in Local Server Table<br>&nbsp;<br>")
IF Len(DupMemList) > 0 THEN Response.write("Duplicate Person ID's encountered: " & DupMemList & "<br>&nbsp;<br>")

WriteLog(date() &"  "& time() &"  Membership Extract & Update Completed Successfully.")

' ************** Bottom of Member Extract Update Conditional Section ***************


' ********* This next section pulls the Membership Consolidation "Was-to-Is"
' ********* cross-reference table, and Updates Member IDs in the various
' ********* Rankings Database table -- Raw Scores, Overall Scores, Rankings.

Response.write("Member Consolidation -- Querying HQ Server ...")

sSQL = "SELECT PersonIDDeleted as OldMemID, PersonIDConsolidatedTo as NewMemID"
sSQL = sSQL & " FROM	Waterski.dbo.[Consolidated Members]"

' WriteDebugSQL(sSQL)

Set HQrs = HQConnect.Execute(sSql)
tempvar = HQrs.getrows()
nTotal = ubound(tempvar,2)
nSoFar = 0

HQrs.MoveFirst

Response.write(" Consolidating Scores ...")

DO UNTIL HQrs.eof

	' Loop over Consolidated Membership ID rows.
	' First step is to see if "NewMemID" (ConsolidateTo) exists in the member table ...

	nSoFar = nSoFar + 1
	IF (nSoFar mod 5 = 2) THEN ShowProgress (nSoFar / nTotal)

	sSQL = "Select count(*) as Kount from "  & MemberTableName
	sSQL = sSQL & " Where PersonIDWithCheckDigit = '" & PersonIDwChkDgt(HQrs("NewMemID")) & "'" 
	rs.open sSQL, sConnectionToTRATable, 3, 3
	rs.MoveFirst
	j = rs("Kount")
	rs.close

	IF j > 0 THEN

		' "NewMemID" record found -- so now update any score table records 
		
		sSQL = "Select count(*) as Kount from " & RawScoresTableName
		sSQL = sSQL & " Where MemberID = '" & PersonIDwChkDgt(HQrs("OldMemID")) & "'" 
		rs.open sSQL, sConnectionToTRATable, 3, 3
		rs.MoveFirst
		IF rs("Kount") > 0 THEN
			sSQL = "UPDATE " & RawScoresTableName & " SET MemberID = '"
			sSQL = sSQL & PersonIDwChkDgt(HQrs("NewMemID")) & "' WHERE MemberID = '"
			sSQL = sSQL & PersonIDwChkDgt(HQrs("OldMemID")) & "'"
			Con.Execute(sSQL)
			nScoUpdts = nScoUpdts + rs("Kount")
		END IF
		rs.close
	
		' Next update any Overall Score table records

		sSQL = "Select count(*) as Kount from " & OverallScoresTableName
		sSQL = sSQL & " Where MemberID = '" & PersonIDwChkDgt(HQrs("OldMemID")) & "'" 
		rs.open sSQL, sConnectionToTRATable, 3, 3
		rs.MoveFirst
		IF rs("Kount") > 0 THEN
			sSQL = "UPDATE " & OverallScoresTableName & " SET MemberID = '"
			sSQL = sSQL & PersonIDwChkDgt(HQrs("NewMemID")) & "' WHERE MemberID = '"
			sSQL = sSQL & PersonIDwChkDgt(HQrs("OldMemID")) & "'"
			Con.Execute(sSQL)
			nOvrUpdts = nOvrUpdts + rs("Kount")
		END IF
		rs.close
	
		' Next update any Ranking table records

		sSQL = "Select count(*) as Kount from " & RankTableName
		sSQL = sSQL & " Where MemberID = '" & PersonIDwChkDgt(HQrs("OldMemID")) & "'" 
		rs.open sSQL, sConnectionToTRATable, 3, 3
		rs.MoveFirst
		IF rs("Kount") > 0 THEN
			sSQL = "UPDATE " & RankTableName & " SET MemberID = '"
			sSQL = sSQL & PersonIDwChkDgt(HQrs("NewMemID")) & "' WHERE MemberID = '"
			sSQL = sSQL & PersonIDwChkDgt(HQrs("OldMemID")) & "'"
			Con.Execute(sSQL)
			nRnkUpdts = nRnkUpdts + rs("Kount")
		END IF
		rs.close
		
		' Finally update any Officials table records

		sSQL = "Select count(*) as Kount from USAWaterski.dbo.Officials"
		sSQL = sSQL & " Where PersonID = '" & HQrs("OldMemID") & "'" 
		rs.open sSQL, sConnectionToTRATable, 3, 3
		rs.MoveFirst
		IF rs("Kount") > 0 THEN
			sSQL = "UPDATE USAWaterski.dbo.Officials SET PersonID = '"
			sSQL = sSQL & HQrs("NewMemID") & "' WHERE PersonID = '"
			sSQL = sSQL & HQrs("OldMemID") & "'"
			Con.Execute(sSQL)
			nOffUpdts = nOffUpdts + rs("Kount")
		END IF
		rs.close
		
	ELSE
	
		Response.write("<br>Cons " & HQrs("OldMemID") & " to " & HQrs("NewMemID") & " not in MXT")

	END IF
	
	HQrs.moveNEXT

LOOP

HQrs.Close

Response.write(" DONE at " & Time() & "<br>")
Response.write(Formatnumber(nScoUpdts,0) & " Scores updated to consolidated Member IDs<br>")
Response.write(Formatnumber(nOvrUpdts,0) & " Overalls updated to consolidated Member IDs<br>")
Response.write(Formatnumber(nRnkUpdts,0) & " Rankings updated to consolidated Member IDs<br>")
Response.write(Formatnumber(nOffUpdts,0) & " Officials updated to consolidated Member IDs<br>&nbsp;<br>")


' Finally release the HQ record set Object -- all remaining steps local to ePolk tables.

Set HQrs = Nothing



END IF



IF Request("Equival") <> "MemUpd" THEN

' ************** Beginning of Rankings Recalculation Conditional Section ***************

' ========  Begin Rankings Calculation Logic.  First we set some ==========
' ========  flags to control which particular sections are run  ===========

RunOverride="NO"

RunEquivScore = "YES"
RunOvrllScore = "YES"
RunOvrllRanks = "YES"
RunEventRanks = "YES"
RunLevelLogic = "YES"


'------------------ First step in this process is to UPDATE the date range on the 12 month period ------------

tBeginDate = FormatDateTime(dateadd("yyyy",-1,now()),2)
sSQL = "SELECT top 1 enddate from " & RawScoresTableName & " WHERE UPPER(right(rtrim(tourid),1)) = 'A' ORDER BY enddate desc"
  rs.open sSQL, SConnectionToTRATable, 3, 3  
   IF rs.EOF THEN
      session("message") = "No National Scores Found! This is a Strange Error -- Please Report To Admin.  Line #135, Equival.asp"
      WriteLog ("********** ERROR *********** Looking for National Score to set ski year dates and can not find any nationals.")
      Response.Redirect("/?process=logout")
    ELSE
      IF cdate(tBeginDate) > cdate(rs("EndDate")) THEN tBeginDate = rs("EndDate")
    END IF
  rs.Close

sSQL = "UPDATE " & SkiYearTableName & " set BeginDate = '" & tBeginDate & "' , EndDate  = '" & date() & "' WHERE SkiYearID = 1"
Con.Execute(sSQL)


' IF the request included a ski year, we recalculate that one alone.
IF trim(Request("skiyear")) <> "" THEN

    sSQL = "SELECT top 1 * from " & SkiYearTableName & " WHERE skiyearid = " & sqlclean(Request("skiyear"))
    rs.open sSQL, SConnectionToTRATable, 3, 3  
    ' IF the specified year doesn't exist, THEN someone messed up! :)
      IF rs.EOF THEN
        session("message") = "Ski Year ID " & request("skiyear") & " was not found."
        WriteLog ("Ski Year ID " & request("skiyear") & " was not found.")
        Response.Redirect("/?process=logout")
      ELSE
      '  IF there is, THEN we save all the variables we will need.
        sProcessingYear = rs("SkiYearID")
        sPrevYear = rs("PrevYearID")
        sSkiYearName = rs("SkiYearName")
        sSkiYearBegin = rs("BeginDate")
        sSkiYearEnd = rs("EndDate")    
      END IF
    rs.Close

ELSE

' Otherwise we have to figure out what the current ski year is
' so we look it up in the table.
    sSQL = "SELECT top 1 * from " & SkiYearTableName & " WHERE DefaultYear = 1"
    rs.open sSQL, SConnectionToTRATable, 3, 3  
    ' IF there is no current year, THEN we just do the 12 month calc.
      IF rs.EOF THEN
        session("message") = "There is no current ski year SELECTed.  Ranking calculation can not be performed."
        WriteLog ("**** There is no current ski year SELECTed.  Ranking calculation can not be performed. *****")
        Response.Redirect("/?process=logout")
      ELSE
      '  IF there is, THEN we save all the variables we will need.
        sProcessingYear = rs("SkiYearID")
        sPrevYear = rs("PrevYearID")
        sSkiYearName = rs("SkiYearName")
        sSkiYearBegin = rs("BeginDate")
        sSkiYearEnd = rs("EndDate")    
      END IF
    rs.Close
END IF



' This do loop keeps repeating the recalc until we tell it to stop.
' The value in sProcessingYear is the Ski Year ID that we are processing for.
' sPrevYear is the Previous Ski Year ID that we use for testing CutOffScores.

DO WHILE sProcessingYear <> "STOP"

' First we set the Recalculation Underway flag for the Ski Year being processsed.

sSQL = "UPDATE " & SkiYearTableName & " set RecalcUnderway = 1 WHERE SkiYearID = " & sProcessingYear
Con.Execute(sSQL)

Response.write(sSkiYearName & " Rankings ...")


' Section # 1

IF RunEquivScore = "YES" OR RunOverride = "YES" THEN

Response.write(" Equiv Scores ...")

' ------------------------  Prepare Equivalent Scores Table --------------------------

' This section creates the EQUIVALENT SCORES Table.  This is used for both OverAll and
' Event Ranking purposes, and reflects a number of common pre-processing features --

'   1.  Equivalencies are applied, according to applicable Division Table entries, and
'       other considerations.  This can (and does) lead to the same underlying performance
'       appearing in the table, for EACH equivalent division, for which that score COULD
'       contribute to an overall or event ranking.  In Slalom, adjustments may be made
'       to conform to cross-divisional scoring differences.  Each Equivalent Score row
'       includes both the Original Division code, as well as the Effective Division
'       to which the performance is being Equivalenced.

'   2.  For the 12 Month rolling window, the logic includes special considerations
'       to deal with GRADUATING SKIERS, who change divisions during the span of the
'       12 month window.  See note below about EQUIVALENT DIVISIONS.

'   3.  For each surviving Equivalent Score, a Rating is assigned, according to the 
'       applicable Ski Year Division Control Table entry for that Effective Division. 
'       To present a meaningful collating sequence, 2-characters codes of 4O for Open,
'       3E for EP, 2M for Masters, and 1X for Expert are used. 

'   4.  For each surviving Equivalent Score, the Overall score value is computed according
'       to the parameters for the applicable Effective Division control table entry.  
'       Updated Logic (April 2007) Calculates 2007+ Smooth NOPS if OverExp_? present in 
'       DivTab, otherwise continues to calculate older-format Connect-the-Dots Ratings-Based 
'       NOPS formulations.  DC 4/24/2007

'   5.  A Formatted Score string is derived, used later in MouseOver detail presentations.


' The First step is to fill in a table of Allowable Equivalent Divisions, for each Skier and 
' Event, for the current Processing Year.  Where the Processing Year is a complete ski year, 
' that list of divisions consists solely of divisions for which scores by each skier actually
' appear, which are coded as EStatus=2.

' For the 12 month Rolling window, this list will also include New Graduate Equivalencies, 
' to pick up scores for recently graduated skiers who have not yet produced a score in their 
' new age division, coded as EStatus=1 -- which are to apply to scores from all eligible 
' equivalent divisions EXCEPT Ox or Ix.  Also for the 12 month rolling window, this list
' will include EStatus=3 entries, for the Aged Out division for graduating skiers, to block
' those equivalencies which might already be present as EStatus=2.  This table is grouped 
' by MemberID, Event & Division, with MAX(EStatus) -- so EStatus=2 overrides EStatus=1, 
' and EStatus=3 overrides EStatus=2.  This EStatus=3 replaces the old "Age Out" logic.

' NOTE -- Unlike the Equivalent Scores and Overall Scores and Rankings tables, this EquivDivs 
' table is local just to the current sProcessingYear, and the content is not retained across 
' sProcessingYears.  It is used only locally within the processing for that sProcessingYear.

' First clear the table

sSQL = "DELETE FROM " & EquivDivsTableName

' WriteDebugSQL(sSQL)

Con.Execute(sSQL)

' Now generate the Query which will populate the Equivalent Divisions table,
' for the current sProcessingYear.

sSQL = "INSERT INTO " & EquivDivsTableName & " (MemberID, Event, Div, EStatus)"

IF sProcessingYear = 1 THEN
sSQL = sSQL & " SELECT MemberID, Event, Div, Max(EStatus) as EStatus FROM ("
END IF

' EStatus = 2 Entries derived from Actual performance rows present for this period.
' But only for "Ranking Divisions", where Left(RS.Div,1) in ('B','G','M','W','O','C')
' Hence omits Novice and International

sSQL = sSQL & " SELECT RS.MemberID, RS.Event, RS.Div, '2' as EStatus FROM " & RawScoresTableName & " as RS,"
sSQL = sSQL & " (Select begindate as BDate, enddate as EDate from " & SkiYearTableName & " where SkiYearID = " & sProcessingYear & ") as SY"
sSQL = sSQL & " WHERE RS.EndDate between SY.BDate and SY.EDate and RS.Score is not null"
sSQL = sSQL & " and Left(RS.Div,1) in ('B','G','M','W','O','C')"
sSQL = sSQL & " GROUP BY RS.MemberID, RS.Event, RS.Div"

' When sProcessingYear = 1, we need special Estatus codes to deal with Graduates.
' First Union folds in EStatus = "3" (Aged Out) for the "Graduating From" division,
'    this will later suppress these skiers in their "From" Division's Rankings,
' then the Second Union folds in an EStatus for the "Graduating To" division ...
'    That EStatus value for the "To" (incoming) Division will be "2", if there are 
'    scores for the event in the "From" Division, otherwise it will be "1"
'    Hence a "Presence" (EStatus=2) in the new division can be established
'    either by actual scores in that new "To" age division (from above), or by 
'    actual scores in the previous "From" age division, as derived here below.

IF sProcessingYear = 1 THEN

sSQL = sSQL & " UNION SELECT EL.MemberID, EL.Event, DT.Div, '3' as EStatus" 
sSQL = sSQL & " FROM	" & MemberTableName & "	as	MT," & DivisionsTableName & " as DT,"
sSQL = sSQL & " (SELECT Year(BeginDate)-1 as BYear FROM " & SkiYearTableName & " where DefaultYear = 1) as EY,"
sSQL = sSQL & " (SELECT MemberID, Event FROM " & RawScoresTableName & ","
sSQL = sSQL & " (Select begindate as BDate, enddate as EDate from " & SkiYearTableName & " where SkiYearID = 1) as SY"
sSQL = sSQL & " WHERE Score is not null and EndDate between SY.BDate and SY.EDate GROUP BY MemberID, Event) as EL"
sSQL = sSQL & " WHERE	MT.PersonIDWithCheckDigit = EL.MemberID"
sSQL = sSQL & " and Left(MT.Sex,1) = DT.Sex"
sSQL = sSQL & " and EY.BYear - Year(MT.BirthDate) = DT.UP_Age"
sSQL = sSQL & " and DT.Next_Div > 'AA'"
sSQL = sSQL & " and	DT.SkiYearID = 1"

sSQL = sSQL & " UNION SELECT GL.MemberID, GL.Event, GL.Next_Div,"
sSQL = sSQL & " Case when EG.MemberID is not null then '2' else '1' end as EStatus FROM"
sSQL = sSQL & " (SELECT EL.MemberID, EL.Event, DT.Div, DT.Next_Div"
sSQL = sSQL & " FROM " & MemberTableName & "	as	MT," & DivisionsTableName & " as DT,"
sSQL = sSQL & " (SELECT Year(BeginDate)-1 as BYear FROM " & SkiYearTableName & " where DefaultYear = 1) as EY,"
sSQL = sSQL & " (SELECT MemberID, Event FROM " & RawScoresTableName & ","
sSQL = sSQL & "  (Select begindate as BDate, enddate as EDate from " & SkiYearTableName & " where SkiYearID = 1) as SY"
sSQL = sSQL & "  WHERE Score is not null and EndDate between SY.BDate and SY.EDate GROUP BY MemberID, Event) as EL"
sSQL = sSQL & " WHERE	MT.PersonIDWithCheckDigit = EL.MemberID"
sSQL = sSQL & " and Left(MT.Sex,1) = DT.Sex"
sSQL = sSQL & " and EY.BYear - Year(MT.BirthDate) = DT.UP_Age"
sSQL = sSQL & " and DT.Next_Div > 'AA'"
sSQL = sSQL & " and	DT.SkiYearID = 1) as GL"
sSQL = sSQL & " LEFT JOIN (SELECT MemberID, Div, Event FROM " & RawScoresTableName & ","
sSQL = sSQL & "  (Select begindate as BDate, enddate as EDate from " & SkiYearTableName & " where SkiYearID = 1) as SY"
sSQL = sSQL & "  WHERE Score is not null and EndDate between SY.BDate and SY.EDate GROUP BY MemberID, Div, Event) as EG"
sSQL = sSQL & "  on EG.MemberID = GL.MemberID and EG.Div = GL.Div and EG.Event = GL.Event"

sSQL = sSQL & " ) as ED GROUP BY MemberID, Event, Div"
END IF

' Finally execute the constructed query

WriteDebugSQL(sSQL)

Con.Execute(sSQL)

' Now ready to populate the Equivalent Scores table for this SkiYear ... BUT
' first we delete all EquivScores rows for the SkiYearID that we are processing.

sSQL = "DELETE FROM " & EquivScoresTableName & " WHERE SkiYearID = " & sProcessingYear

'	WriteDebugSQL(sSQL)

Con.Execute(sSQL)

' Second step is to extract all "Equivalent" scores into the EquivScores table.
' This is done separately by Event, in each case spreading all scores out to 
' whatever eligible divisions we find In the EQUIVALENT DIVISION table built above.
' Listed Equivalent divisions are checked in the DCT for Slalom and Jumping, and 
' a max speed adjustment may be made in slalom, if applicable.  At the same time, 
' an overall score value is derived for each such equivalent performance, 
' according to the parameters in the applicable Division table entry, along with
' a prioritized classification of the rating.  A prioritized Event Class is also 
' derived for each performance:  5R/4L/3E/2C/1?.  This is subsequently used in
' deriving the composite class for each overall score, and then later in the event 
' and overall ranking calculations.  Note that scores from all classes are pulled
' at this stage -- then after all 3 events are added, a final Update query is run
' to "Cap" any class F/I/N scores at the prior ski year's Level 5 COA.
' This logic last updated May 2008 DJC.

' Candidate Tournaments that are extracted depends on the Ski Year being processed:
'
'    Full Ski Year candidates consist merely of all EndDates that fall
'    within the date range specified in the applicable ski year table entry.
'	
'    Last 12 Months candidates consist of two sets -- the first set consists
'    of the most recent Regional and National tournament for each Region, and
'    the second set of all other tournaments (those with Suffix codes of C or 
'    higher), that fall within the moving 365 day range ending today.


' First we do Slalom, using a single complex query, of two levels.  The innermost 
' level spreads each actual score out to other divisions in the EQUIVALENT DIVISIONS
' Table built above, where that combination is explicitly listed in the Division
' Control Table as an allowed equivalent, adjusting if necessary for any Max Speed 
' differences where applicable, and extracting the "Formatted Score" string for later 
' display.  The outer level then matches in the parameters for the effective division on 
' each such equivalenced score, derives the prioritized rating and class, and calculates
' the overall score component.

sSQL = "INSERT INTO " & EquivScoresTableName
sSQL = sSQL & " (MemberID, Team, TourID, Div, Event, Round, Class, PrioClass, Place, ScoreOrig, DivOrig, Score, Rating, FmtScore, OAScore, SkiYearID)"
sSQL = sSQL & " SELECT ES.MemberID, ES.Team, ES.TourID, DE.Div, ES.Event, ES.Round, ES.Class, Case when ES.Class='R' then '5R'"
sSQL = sSQL & " when ES.Class='L' then '4L' when ES.Class='E' then '3E' when ES.Class='C' then '2C'"
sSQL = sSQL & " else '1' + ES.Class end as PrioClass, ES.Place, ES.ScoreOrig, ES.DivOrig,"
sSQL = sSQL & " Case when ES.Score < 0 then 0 else ES.Score end,"
sSQL = sSQL & " Case when (ES.Score >= DE.OP_S and ES.Class in ('E','L','R')) then '4O'"
sSQL = sSQL & " when (ES.Score >= DE.EP_S and LEFT(DE.Div,1) <> 'O') then '3E'"
sSQL = sSQL & " when (ES.Score >= DE.MS_S and LEFT(DE.Div,1) <> 'O') then '2M'"
sSQL = sSQL & " when (ES.Score >= DE.XP_S and LEFT(DE.Div,1) <> 'O') then '1X' else '  ' end as Rating,"
sSQL = sSQL & " 'Rd ' + ES.Round + 'as ' + ES.DivOrig + '&#13;Score: ' + Cast (Cast(ES.Score as Decimal(5,2)) as Varchar(6)) + '&#13;' + ES.FmtScore,"
sSQL = sSQL & " Case when (DE.OverExp_S > 0) and (ES.Score < 6) then ES.Score * DE.OverPtsBy_S"
sSQL = sSQL & " when DE.OverExp_S > 0 then (6 * DE.OverPtsBy_S) + ((1500 - (6 * DE.OverPtsBy_S)) * Power ((ES.Score - 6) / (DE.Over_S - 6), DE.OverExp_S ))"
sSQL = sSQL & " when ES.Score <= DE.FirstClass_S  and  DE.FirstClass_S > 0 then  200 * ES.Score / DE.FirstClass_S"
sSQL = sSQL & " when ES.Score <= DE.XP_S and DE.XP_S > DE.FirstClass_S then 200 + (200 * (ES.Score - DE.FirstClass_S) / (DE.XP_S - DE.FirstClass_S))"
sSQL = sSQL & " when ES.Score <= DE.MS_S and DE.MS_S > DE.XP_S then 400 + (200 * (ES.Score - DE.XP_S) / (DE.MS_S - DE.XP_S))"
sSQL = sSQL & " when ES.Score <= DE.EP_S and DE.EP_S > DE.MS_S then 600 + (200 * (ES.Score - DE.MS_S) / (DE.EP_S - DE.MS_S))"
sSQL = sSQL & " when DE.NationalRec_S > DE.EP_S then 800 + (700 * (ES.Score - DE.EP_S) / (DE.NationalRec_S - DE.EP_S))"
sSQL = sSQL & " else 0 end as OAScore, " & sProcessingYear & " as SkiYearID"
sSQL = sSQL & " FROM (SELECT RS.MemberID, RS.Team, RS.TourID, DE.Div, RS.Event, RS.Round, RS.Class, RS.Place, RS.Score as ScoreOrig, RS.Div as DivOrig,"
sSQL = sSQL & " CASE when RS.Div = DE.Div then RS.Score"
sSQL = sSQL & " when (DE.Max_S1 < DO.Max_S1 and RS.Perf_Qual2 > DE.Max_S1) then RS.Score - (2 * (DO.Max_S1 - DE.Max_S1))"
sSQL = sSQL & " else RS.Score end as Score,"
sSQL = sSQL & " Cast (Cast(RS.AltScore as Decimal(5,2)) as Varchar(5)) + '@' + Cast (Cast(RS.Perf_Qual2 as Decimal(3)) as Varchar(3)) + 'k ' + Cast (Cast(RS.Perf_Qual1/100 as Decimal(5,2)) as Varchar(5)) + 'm&#13;Class: ' + RS.Class as FmtScore"
sSQL = sSQL & " FROM " & RawScoresTableName & " as RS, " & DivisionsTableName & " as DO, " & DivisionsTableName & " as DE, " & EquivDivsTableName & " as ED"

sSQL = sSQL & " WHERE RS.Event = 'S' and RS.Score is not null and (UPPER(RS.Class) in ('F','N','I','C','E','L','R')) and RS.TourID in"

IF sProcessingYear = 1 THEN
	sSQL = sSQL & " (Select TourID From (Select Rgn, Right(Max(DateTour),7) as TourID From (Select Distinct"
	sSQL = sSQL & " Case when Substring(TourID,7,1)='A' then 'N' else Substring(TourID,3,1) end as Rgn,"
	sSQL = sSQL & " Convert(char,EndDate,112) + Left(TourID,7) as DateTour From " & RawScoresTableName
	sSQL = sSQL & " Where  Substring(TourID,7,1) in ('A','B') ) as RDT Group by Rgn) as RNT"
	sSQL = sSQL & " UNION Select Distinct TourID From " & RawScoresTableName & " as ST,"
	sSQL = sSQL & " (Select begindate as BDate, enddate as EDate from " & SkiYearTableName & " where SkiYearID = " & sProcessingYear & ") as SY"
	sSQL = sSQL & " Where ST.EndDate between SY.BDate and SY.EDate and Substring(TourID,7,1)>'B')"
ELSE
	sSQL = sSQL & " (Select Distinct TourID From " & RawScoresTableName & " as ST,"
	sSQL = sSQL & " (Select begindate as BDate, enddate as EDate from " & SkiYearTableName & " where SkiYearID = " & sProcessingYear & ") as SY"
	sSQL = sSQL & " Where ST.EndDate between SY.BDate and SY.EDate)"
END IF

sSQL = sSQL & " and (RS.Div = DE.Div or (RS.Div = DE.SL_ED1 or RS.Div = DE.SL_ED2 or RS.Div = DE.SL_ED3 or RS.Div = DE.SL_ED4 or RS.Div = DE.SL_ED5 or RS.Div = DE.SL_ED6 or RS.Div = DE.SL_ED7 or RS.Div = DE.SL_ED8))"
sSQL = sSQL & " and (LEFT(DE.Div,1) <> 'O' or RS.Class in ('E','L','R'))"
sSQL = sSQL & " and DO.Div = RS.Div and DO.SkiYearID = " & sProcessingYear
sSQL = sSQL & " and DE.Div = ED.Div and DE.SkiYearID = " & sProcessingYear
sSQL = sSQL & " and RS.MemberID = ED.MemberID and ED.Event = 'S'"
sSQL = sSQL & " and ED.EStatus <> '3' and (ED.EStatus = '2' OR Left(RS.Div,1) not in ('O','I'))"
sSQL = sSQL & " ) as ES, " & DivisionsTableName & " as DE"
sSQL = sSQL & " WHERE ES.Div = DE.Div and DE.SkiYearID = " & sProcessingYear

WriteDebugSQL(sSQL)

Con.Execute(sSQL)


' Next we do Trick, using a single complex query which spreads each actual 
' score out to other divisions listed for this skier in the EQUIVALENT DIVISION
' Table.  We extract the "Formatted Score" string for later display, matching in
' the parameters for the effective division on each such equivalenced score, and 
' then deriving a prioritized rating and calculating the overall value therewith.
' A prioritized class code is also derived.  Since the Division Control table 
' doesn't specifically list allowed equivalent divisions for Tricks, all 
' combinations would be allowed, except we screen for CM / CW to keep AWSA 
' scores out of those NCWSA divisions.

sSQL = "INSERT INTO " & EquivScoresTableName
sSQL = sSQL & " (MemberID, Team, TourID, Div, Event, Round, Class, PrioClass, Place, ScoreOrig, DivOrig, Score, Rating, FmtScore, OAScore, SkiYearID)"
sSQL = sSQL & " SELECT RS.MemberID, RS.Team, RS.TourID, DE.Div, RS.Event, RS.Round, RS.Class, Case when RS.Class='R' then '5R'"
sSQL = sSQL & " when RS.Class='L' then '4L' when RS.Class='E' then '3E' when RS.Class='C' then '2C'"
sSQL = sSQL & " else '1' + RS.Class end as PrioClass, RS.Place, RS.Score, RS.Div, RS.Score,"
sSQL = sSQL & " Case when (RS.Score >= DE.OP_T and RS.Class in ('E','L','R')) then '4O'"
sSQL = sSQL & " when (RS.Score >= DE.EP_T and LEFT(DE.Div,1) <> 'O') then '3E'"
sSQL = sSQL & " when (RS.Score >= DE.MS_T and LEFT(DE.Div,1) <> 'O') then '2M'"
sSQL = sSQL & " when (RS.Score >= DE.XP_T and LEFT(DE.Div,1) <> 'O') then '1X' else '  ' end as Rating,"
sSQL = sSQL & " 'Rd ' + RS.Round + 'as ' + RS.Div + '&#13;Score: ' + Cast (Cast(RS.Score as Decimal(6,0)) as Varchar(6)) + '&#13;Class: ' + RS.Class as FmtScore,"
sSQL = sSQL & " Case when  DE.OverExp_T > 0 then 900 * Power(RS.Score/DE.Over_T, DE.OverExp_T)"
sSQL = sSQL & " when RS.Score <= DE.FirstClass_T and DE.FirstClass_T > 0 then 200 * RS.Score / DE.FirstClass_T"
sSQL = sSQL & " when RS.Score <= DE.XP_T and DE.XP_T > DE.FirstClass_T then 200 + (200 * (RS.Score - DE.FirstClass_T) / (DE.XP_T - DE.FirstClass_T))"
sSQL = sSQL & " when RS.Score <= DE.MS_T and DE.MS_T > DE.XP_T then 400 + (200 * (RS.Score - DE.XP_T) / (DE.MS_T - DE.XP_T))"
sSQL = sSQL & " when RS.Score <= DE.EP_T and DE.EP_T > DE.MS_T then 600 + (200 * (RS.Score - DE.MS_T) / (DE.EP_T - DE.MS_T))"
sSQL = sSQL & " when DE.NationalRec_T > DE.EP_T then 800 + (700 * (RS.Score - DE.EP_T) / (DE.NationalRec_T - DE.EP_T))"
sSQL = sSQL & " else 0 end as OAScore, " & sProcessingYear & " as SkiYearID"
sSQL = sSQL & " FROM " & RawScoresTableName & " as RS, " & DivisionsTableName & " as DE, " & EquivDivsTableName & " as ED"

sSQL = sSQL & " WHERE RS.Event = 'T' and RS.Score is not null and (UPPER(RS.Class) in ('F','N','I','C','E','L','R')) and RS.TourID in"

IF sProcessingYear = 1 THEN
	sSQL = sSQL & " (Select TourID From (Select Rgn, Right(Max(DateTour),7) as TourID From (Select Distinct"
	sSQL = sSQL & " Case when Substring(TourID,7,1)='A' then 'N' else Substring(TourID,3,1) end as Rgn,"
	sSQL = sSQL & " Convert(char,EndDate,112) + Left(TourID,7) as DateTour From " & RawScoresTableName
	sSQL = sSQL & " Where  Substring(TourID,7,1) in ('A','B') ) as RDT Group by Rgn) as RNT"
	sSQL = sSQL & " UNION Select Distinct TourID From " & RawScoresTableName & " as ST,"
	sSQL = sSQL & " (Select begindate as BDate, enddate as EDate from " & SkiYearTableName & " where SkiYearID = " & sProcessingYear & ") as SY"
	sSQL = sSQL & " Where ST.EndDate between SY.BDate and SY.EDate and Substring(TourID,7,1)>'B')"
ELSE
	sSQL = sSQL & " (Select Distinct TourID From " & RawScoresTableName & " as ST,"
	sSQL = sSQL & " (Select begindate as BDate, enddate as EDate from " & SkiYearTableName & " where SkiYearID = " & sProcessingYear & ") as SY"
	sSQL = sSQL & " Where ST.EndDate between SY.BDate and SY.EDate)"
END IF

sSQL = sSQL & " and (LEFT(DE.Div,1) <> 'O' or RS.Class in ('E','L','R'))"
sSQL = sSQL & " and RS.MemberID = ED.MemberID and ED.Event = 'T'"
sSQL = sSQL & " and ED.EStatus <> '3' and (ED.EStatus = '2' OR Left(RS.Div,1) not in ('O','I'))"
sSQL = sSQL & " and (left(DE.Div,1)<>'C' or left(RS.Div,1)='C')"
sSQL = sSQL & " and DE.Div = ED.Div and DE.SkiYearID = " & sProcessingYear

WriteDebugSQL(sSQL)

Con.Execute(sSQL)


' Finally we do Jump, using a single complex query which spreads each actual score 
' out to other divisions listed for this skier in the EQUIVALENT DIVISION Table, 
' where the combination is explicitly cited in the Division Control Table, and
' where the actual conditions do not exceed the allowed limits for that division,
' extracting the "Formatted Score" string for later display, matching in the 
' parameters for the effective division on each such equivalenced score, and 
' deriving the prioritized rating and class, and calculating the overall 
' score component.

sSQL = "INSERT INTO " & EquivScoresTableName
sSQL = sSQL & " (MemberID, Team, TourID, Div, Event, Round, Class, PrioClass, Place, ScoreOrig, DivOrig, Score, Rating, FmtScore, OAScore, SkiYearID)"
sSQL = sSQL & " SELECT RS.MemberID, RS.Team, RS.TourID, DE.Div, RS.Event, RS.Round, RS.Class, Case when RS.Class='R' then '5R'"
sSQL = sSQL & " when RS.Class='L' then '4L' when RS.Class='E' then '3E' when RS.Class='C' then '2C'"
sSQL = sSQL & " else '1' + RS.Class end as PrioClass, RS.Place, RS.Score, RS.Div, RS.Score,"
sSQL = sSQL & " Case when (((RS.Score >= DE.OP_J and RS.Perf_Qual1 <= DE.Ramp1 and RS.Perf_Qual2 <= DE.Max_J1) or (DE.OP_J2 > 0 and RS.Score >= DE.OP_J2 and RS.Perf_Qual1 <= DE.Ramp2 and RS.Perf_Qual2 <= DE.Max_J2)) and RS.Class in ('E','L','R')) then '4O'"
sSQL = sSQL & " when (RS.Score >= DE.EP_J and LEFT(DE.Div,1) <> 'O') then '3E'"
sSQL = sSQL & " when (RS.Score >= DE.MS_J and LEFT(DE.Div,1) <> 'O') then '2M'"
sSQL = sSQL & " when (RS.Score >= DE.XP_J and LEFT(DE.Div,1) <> 'O') then '1X' else '  ' end as Rating,"
sSQL = sSQL & " 'Rd ' + RS.Round + 'as ' + RS.Div + '&#13;Score: ' + Cast (Cast(RS.Score as Decimal(3)) as Varchar(4)) + '&#13;' + Cast (Cast(RS.Perf_Qual1 as Decimal(5,3)) as Varchar(6)) + ' @ ' + Cast (Cast(RS.Perf_Qual2 as Decimal(3)) as Varchar(3)) + 'k&#13;Class: ' + RS.Class as FmtScore,"
sSQL = sSQL & " Case when  (DE.OverExp_J > 0) and (RS.Score < (0.15*DE.NationalRec_J)) then 0"
sSQL = sSQL & " when DE.OverExp_J > 0 then 700 * Power ((RS.Score - (0.15*DE.NationalRec_J)) / (DE.Over_J - (0.15*DE.NationalRec_J)), DE.OverExp_J)"
sSQL = sSQL & " when RS.Score <= DE.FirstClass_J and DE.FirstClass_J > 0 then  200 * RS.Score / DE.FirstClass_J"
sSQL = sSQL & " when RS.Score <= DE.XP_J  and  DE.XP_J > DE.FirstClass_J then 200 + (200 * (RS.Score - DE.FirstClass_J) / (DE.XP_J - DE.FirstClass_J))"
sSQL = sSQL & " when RS.Score <= DE.MS_J and DE.MS_J > DE.XP_J then 400 + (200 * (RS.Score - DE.XP_J) / (DE.MS_J - DE.XP_J))"
sSQL = sSQL & " when RS.Score <= DE.EP_J  and  DE.EP_J > DE.MS_J then 600 + (200 * (RS.Score - DE.MS_J) / (DE.EP_J - DE.MS_J))"
sSQL = sSQL & " when DE.NationalRec_J > DE.EP_J then 800 + (700 * (RS.Score - DE.EP_J) / (DE.NationalRec_J - DE.EP_J))"
sSQL = sSQL & " else 0 end as OAScore, " & sProcessingYear & " as SkiYearID"
sSQL = sSQL & " FROM " & RawScoresTableName & " as RS, " & DivisionsTableName & " as DE, " & EquivDivsTableName & " as ED"

sSQL = sSQL & " WHERE RS.Event = 'J' and RS.Score is not null and (UPPER(RS.Class) in ('F','N','I','C','E','L','R')) and RS.TourID in"

IF sProcessingYear = 1 THEN
	sSQL = sSQL & " (Select TourID From (Select Rgn, Right(Max(DateTour),7) as TourID From (Select Distinct"
	sSQL = sSQL & " Case when Substring(TourID,7,1)='A' then 'N' else Substring(TourID,3,1) end as Rgn,"
	sSQL = sSQL & " Convert(char,EndDate,112) + Left(TourID,7) as DateTour From " & RawScoresTableName
	sSQL = sSQL & " Where  Substring(TourID,7,1) in ('A','B') ) as RDT Group by Rgn) as RNT"
	sSQL = sSQL & " UNION Select Distinct TourID From " & RawScoresTableName & " as ST,"
	sSQL = sSQL & " (Select begindate as BDate, enddate as EDate from " & SkiYearTableName & " where SkiYearID = " & sProcessingYear & ") as SY"
	sSQL = sSQL & " Where ST.EndDate between SY.BDate and SY.EDate and Substring(TourID,7,1)>'B')"
ELSE
	sSQL = sSQL & " (Select Distinct TourID From " & RawScoresTableName & " as ST,"
	sSQL = sSQL & " (Select begindate as BDate, enddate as EDate from " & SkiYearTableName & " where SkiYearID = " & sProcessingYear & ") as SY"
	sSQL = sSQL & " Where ST.EndDate between SY.BDate and SY.EDate)"
END IF

sSQL = sSQL & " and (RS.Div = DE.Div or ((RS.Div = DE.JU_ED1 or RS.Div = DE.JU_ED2 or RS.Div = DE.JU_ED3 or RS.Div = DE.JU_ED4 or RS.Div = DE.JU_ED5 or RS.Div = DE.JU_ED6 or RS.Div = DE.JU_ED7 or RS.Div = DE.JU_ED8) and RS.Perf_Qual1 <= DE.Ramp1 and RS.Perf_Qual2 <= DE.Max_J1))"
sSQL = sSQL & " and (LEFT(DE.Div,1) <> 'O' or RS.Class in ('E','L','R'))"
sSQL = sSQL & " and RS.MemberID = ED.MemberID and ED.Event = 'J'"
sSQL = sSQL & " and ED.EStatus <> '3' and (ED.EStatus = '2' OR Left(RS.Div,1) not in ('O','I'))"
sSQL = sSQL & " and DE.Div = ED.Div and DE.SkiYearID = " & sProcessingYear

WriteDebugSQL(sSQL)

Con.Execute(sSQL)


' --- Last step is an Update query to "Cap" any Equivalent Scores in Class F/I/N, where the
' score exceeds the preceding ski year's level 5 COA score for the applicable Div/Event.

sSQL = "UPDATE ES Set Score = CO.COA5"
sSQL = sSQL & " FROM " & EquivScoresTableName & " AS ES, " & CutOffTableName & " as CO"
sSQL = sSQL & " WHERE	CO.Div = ES.Div and CO.Event = ES.Event"	
sSQL = sSQL & " AND	ES.SkiYearID = " & sProcessingYear & " AND CO.SkiYearID = " & sPrevYear
sSQL = sSQL & " AND	UPPER(ES.Class) in ('F','I','N') AND ES.Score > CO.COA5"

WriteDebugSQL(sSQL)

Con.Execute(sSQL)


' -----------------------------   End of EQUIVALENT SCORES Processing  --------------------------------

' ----   Finish by posting a new UPDATE Date in SkiYear Table, to show Last Tim/Date Recalculated  ----

sSQL = "UPDATE " & SkiYearTableName & " SET LastRecalc = '" & time() & " on " & date() & "' WHERE SkiYearID = " & sProcessingYear 

WriteDebugSQL(sSQL)

Con.Execute(sSQL)

' Bottom of IF to exclude the Equivalent Scores section above from running
END IF




' Section # 2

IF RunOvrllScore = "YES" OR RunOverride = "YES" THEN

Response.write(" Overall Scores ...")

' --------------  Calculate OVERALL SCORES  ------------------------

' Now we're ready to start assembling overall scores.  But first we must Delete 
' all of the old OverAllScores rows for the SkiYearID that we are processing.

sSQL = "DELETE FROM " & OverAllScoresTableName & " WHERE SkiYearID = " & sProcessingYear
Con.Execute(sSQL)

' Now we are ready to assemble the equivalenced scores by MemberID, TourID, Div and Round,
' Inserting each such qualified collection as a row in the OverAllScores table.
' Adding "Aged Out" logic to include Member Table Birthdate value 20070504 DJC
' First step thereto is to get the Year value for the End Date for processing Ski Year

sSQL = "SELECT YEAR(EndDate) AS SYEndDate FROM " & SkiYearTableName & " where SkiYearID = " & sProcessingYear
RS.open ssql, sConnectionToTRATable, 3, 3
tSYEndDate = rs("SYEndDate")
RS.close

' Following Query pulls all the applicable Overall Scoring elements together, 
' for each unique combination of MemberID / TourID / Div -- ordered by "round".
' But only for Divisions that start with B/G/M/W/O -- leaves others out.

sSQL = "SELECT ES.*, DE.OverNumEvts FROM " & EquivScoresTableName & " ES, " & DivisionsTableName & " as DE"
sSQL = sSQL & " WHERE ES.Div = DE.Div AND ES.SkiYearID = DE.SkiYearID AND ES.SkiYearID = " & sProcessingYear
sSQL = sSQL & " AND LEFT(ES.Div,1) in ('B','G','M','W','O')"
sSQL = sSQL & " ORDER BY ES.memberid, ES.tourid, ES.div, ES.round"

WriteDebugSQL(sSQL)

rs.open sSQL, sConnectionToTRATable, 3, 3

tempvar = rs.getrows()
nTotalMembers = ubound(tempvar,2)

rs.MoveFirst
nProcessedSoFar = 0


' ---------  Outer loop of Overall Score Assembly for all MemberID and TourID's   -----------
DO UNTIL rs.eof

  	nProcessedSoFar = nProcessedSoFar + 1
    IF (nProcessedSoFar mod 10 = 5) THEN ShowProgress (nProcessedSoFar / nTotalMembers)

	TempMemberID = rs("MemberID"): TempTeam = trim(rs("Team"))
	TempTourID = rs("TourID"): TempDiv = rs("Div")
	TempDivOrig = rs("DivOrig"): TempOverEvtsReq = rs("OverNumEvts") 
	Class1 = "9?": Class2 = "9?": Class3 = "9?"
	Slalom1 = "": Slalom2 = "": Slalom3 = ""
	S_Round1 = "": S_Round2 = "": S_Round3 = ""
	S_Score1 = "": S_Score2 = "": S_Score3 = ""
	Trick1 = "": Trick2 = "": Trick3 = ""
	T_Round1 = "": T_Round2 = "": T_Round3 = ""
	T_Score1 = "": T_Score2 = "": T_Score3 = ""
	Jump1 = "": Jump2 = "": Jump3 = ""
	J_Round1 = "": J_Round2 = "": J_Round3 = ""
	J_Score1 = "": J_Score2 = "": J_Score3 = ""
    

	' Loop through Record Set Rows, while THIS MemberID and TourID and Div are the same,
	' Collecting event Overall scores components by "ROUND" -- as defined in AWSA 4.02(b)
	' Only write scores if NOT Aged Out, for the effective division in this ski year.
	' Derive a composite event class for each such round score, based on the lowest
	' class of any contributing event (added Feb 2008 DJC)
  
    DO WHILE TempMemberID = rs("MemberID") AND TempTourID = rs("TourID") AND TempDiv = rs("Div")
     
		TempEvent = rs("Event"): InDivOrig = rs("DivOrig")
		IF TempTeam="" THEN TempTeam = Trim(rs("Team"))

		' Only include event scores where the Original Division Codes agree.
		' Added Jan 2008 DJC following rules discussion & vote.  Also derive
		' the lowest class for each round from incoming score classes. 

  	IF InDivOrig = TempDivOrig THEN

	  	SELECT CASE TempEvent
  		
			CASE "S"
        IF Slalom1 <> "" THEN
           IF Slalom2 <> "" THEN
              Slalom3 = rs("OAScore"): S_Round3 = rs("Round"): S_Score3 = rs("FmtScore")
              IF Class3 > rs("PrioClass") THEN Class3 = rs("PrioClass")
           ELSE    
              Slalom2 = rs("OAScore"): S_Round2 = rs("Round"): S_Score2 = rs("FmtScore")
              IF Class2 > rs("PrioClass") THEN Class2 = rs("PrioClass")
           END IF
        ELSE
           Slalom1 = rs("OAScore"): S_Round1 = rs("Round"): S_Score1 = rs("FmtScore")
           IF Class1 > rs("PrioClass") THEN Class1 = rs("PrioClass")
        END IF
         
			CASE "T"
        IF Trick1 <> "" THEN
           IF Trick2 <> "" THEN
              Trick3 = rs("OAScore"): T_Round3 = rs("Round"): T_Score3 = rs("FmtScore")
              IF Class3 > rs("PrioClass") THEN Class3 = rs("PrioClass")
           ELSE    
              Trick2 = rs("OAScore"): T_Round2 = rs("Round"): T_Score2 = rs("FmtScore")
              IF Class2 > rs("PrioClass") THEN Class2 = rs("PrioClass")
           END IF
        ELSE
           Trick1 = rs("OAScore"): T_Round1 = rs("Round"): T_Score1 = rs("FmtScore")
           IF Class1 > rs("PrioClass") THEN Class1 = rs("PrioClass")
        END IF

			CASE "J"
       	IF Jump1 <> "" THEN
           IF Jump2 <> "" THEN
              Jump3 = rs("OAScore"): J_Round3 = rs("Round"): J_Score3 = rs("FmtScore")
              IF Class3 > rs("PrioClass") THEN Class3 = rs("PrioClass")
           ELSE    
              Jump2 = rs("OAScore"): J_Round2 = rs("Round"): J_Score2 = rs("FmtScore")
              IF Class2 > rs("PrioClass") THEN Class2 = rs("PrioClass")
           END IF
        ELSE
           Jump1 = rs("OAScore"): J_Round1 = rs("Round"): J_Score1 = rs("FmtScore")
           IF Class1 > rs("PrioClass") THEN Class1 = rs("PrioClass")
        END IF

			END SELECT
			
		END IF

		rs.moveNEXT
		IF rs.eof THEN EXIT DO

		' --------   Bottom of loop where overall scores are assembled for this MemberID/TourID/Div combination  ----------

		LOOP


	' ------  Writes values for this MemberID/TourID/Div to OverallScoresTableName (i.e. raw OVERALL scores tables)  
  
  ' Round 1 Overall Score -- Determine Overall Eligibility for this Round Then Insert if so
  
  TempOverEvts = 0: TempOATot = 0
  IF Slalom1 <> "" THEN TempOverEvts = TempOverEvts + 1: TempOATot = TempOATot + Slalom1
  IF Trick1 <> "" THEN TempOverEvts = TempOverEvts + 1: TempOATot = TempOATot + Trick1
  IF Jump1 <> "" THEN TempOverEvts = TempOverEvts + 1: TempOATot = TempOATot + Jump1

  IF TempOverEvts >= TempOverEvtsReq THEN
  
		sSQL = "INSERT INTO " & OverallScoresTableName
		sSQL = sSQL & " (MemberID, Team, TourID, Round, Div, Class, PrioClass, SlalomOverAll, S_Round, S_OrigScore, JumpOverAll, J_Round, J_OrigScore, TrickOverAll, T_Round, T_OrigScore, TotalOverAll, SkiYearID, DivOrig)"
		sSQL = sSQL & " VALUES ('" & TempMemberID & "','" & TempTeam & "','" & TempTourID & "','1','" & TempDiv & "','" & Right(Class1,1) & "','" & Class1 & "',"
		IF Slalom1 = "" THEN sSQL = sSQL & "NULL,NULL,NULL," ELSE sSQL = sSQL & "'" & Slalom1 & "','" & S_Round1 & "','" & S_Score1 & "',"
 		IF Jump1 = "" THEN sSQL = sSQL & "NULL,NULL,NULL," ELSE sSQL = sSQL & "'" & Jump1 & "','" & J_Round1 & "','" & J_Score1 & "',"
		IF Trick1 = "" THEN sSQL = sSQL & "NULL,NULL,NULL," ELSE sSQL = sSQL & "'" & Trick1 & "','" & T_Round1 & "','" & T_Score1 & "',"
		sSQL = sSQL & "'" & TempOATot & "','" & sProcessingYear & "','" & TempDivOrig & "')"
		
		Con.Execute(sSQL)

	END IF

  ' Round 2 Overall Score -- Determine Overall Eligibility for this Round Then Insert if so
  
  TempOverEvts = 0: TempOATot = 0
  IF Slalom2 <> "" THEN TempOverEvts = TempOverEvts + 1: TempOATot = TempOATot + Slalom2
  IF Trick2 <> "" THEN TempOverEvts = TempOverEvts + 1: TempOATot = TempOATot + Trick2
  IF Jump2 <> "" THEN TempOverEvts = TempOverEvts + 1: TempOATot = TempOATot + Jump2

  IF TempOverEvts >= TempOverEvtsReq THEN
  
		sSQL = "INSERT INTO " & OverallScoresTableName
		sSQL = sSQL & " (MemberID, Team, TourID, Round, Div, Class, PrioClass, SlalomOverAll, S_Round, S_OrigScore, JumpOverAll, J_Round, J_OrigScore, TrickOverAll, T_Round, T_OrigScore, TotalOverAll, SkiYearID, DivOrig)"
		sSQL = sSQL & " VALUES ('" & TempMemberID & "','" & TempTeam & "','" & TempTourID & "','2','" & TempDiv & "','" & Right(Class2,1) & "','" & Class2 & "',"
		IF Slalom2 = "" THEN sSQL = sSQL & "NULL,NULL,NULL," ELSE sSQL = sSQL & "'" & Slalom2 & "','" & S_Round2 & "','" & S_Score2 & "',"
 		IF Jump2 = "" THEN sSQL = sSQL & "NULL,NULL,NULL," ELSE sSQL = sSQL & "'" & Jump2 & "','" & J_Round2 & "','" & J_Score2 & "',"
		IF Trick2 = "" THEN sSQL = sSQL & "NULL,NULL,NULL," ELSE sSQL = sSQL & "'" & Trick2 & "','" & T_Round2 & "','" & T_Score2 & "',"
		sSQL = sSQL & "'" & TempOATot & "','" & sProcessingYear & "','" & TempDivOrig & "')"

		Con.Execute(sSQL)

	END IF

  ' Round 3 Overall Score -- Determine Overall Eligibility for this Round Then Insert if so
  
  TempOverEvts = 0: TempOATot = 0
  IF Slalom3 <> "" THEN TempOverEvts = TempOverEvts + 1: TempOATot = TempOATot + Slalom3
  IF Trick3 <> "" THEN TempOverEvts = TempOverEvts + 1: TempOATot = TempOATot + Trick3
  IF Jump3 <> "" THEN TempOverEvts = TempOverEvts + 1: TempOATot = TempOATot + Jump3

  IF TempOverEvts >= TempOverEvtsReq THEN
  
		sSQL = "INSERT INTO " & OverallScoresTableName
		sSQL = sSQL & " (MemberID, Team, TourID, Round, Div, Class, PrioClass, SlalomOverAll, S_Round, S_OrigScore, JumpOverAll, J_Round, J_OrigScore, TrickOverAll, T_Round, T_OrigScore, TotalOverAll, SkiYearID, DivOrig)"
		sSQL = sSQL & " VALUES ('" & TempMemberID & "','" & TempTeam & "','" & TempTourID & "','3','" & TempDiv & "','" & Right(Class3,1) & "','" & Class3 & "',"
		IF Slalom3 = "" THEN sSQL = sSQL & "NULL,NULL,NULL," ELSE sSQL = sSQL & "'" & Slalom3 & "','" & S_Round3 & "','" & S_Score3 & "',"
		IF Jump3 = "" THEN sSQL = sSQL & "NULL,NULL,NULL," ELSE sSQL = sSQL & "'" & Jump3 & "','" & J_Round3 & "','" & J_Score3 & "',"
		IF Trick3 = "" THEN sSQL = sSQL & "NULL,NULL,NULL," ELSE sSQL = sSQL & "'" & Trick3 & "','" & T_Round3 & "','" & T_Score3 & "',"
		sSQL = sSQL & "'" & TempOATot & "','" & sProcessingYear & "','" & TempDivOrig & "')"

		Con.Execute(sSQL)

	END IF
 
' -----------------------  Bottom of Outer LOOP for OVERALL SCORES  ----------------------
LOOP

rs.close


' Bottom of IF to exclude the Overall Scores section above from running
END IF




' Section # 3

IF RunOvrllRanks = "YES" OR RunOverride = "YES" THEN

Response.write(" Overall Ranks ...")

' ----------------------  Now Begin Calculations of OVERALL RANKINGS  ---------------------

' Delete all of the old records which match the Ski Year ID that we are processing.

sSQL = "DELETE FROM " & RankTableName & " WHERE Event = 'O' and SkiYearID = " & sProcessingYear

' WriteDebugSQL(sSQL)

Con.Execute(sSQL)

' Query with Max function and Group By clause, to roll to best single round from each TourID
' Revised Feb 2008 such that Score string is now char(10), with leftmost 8 being the
' actual Total Overall Score (formatted as 8,2), followed by 2-char prioritized class.

sSQL = "SELECT OA.MemberID, OA.TourID, OA.Div, Max("
sSQL = sSQL & " Substring(Cast(Cast(OA.TotalOverAll+400000 as Decimal(8,2)) as Char(9)),2,8)"
sSQL = sSQL & " + OA.PrioClass) as MaxOAScore,"
sSQL = sSQL & " Max(OA.DivOrig) as DivOrig, Max(OA.SkiYearID) as SkiYearID,"
sSQL = sSQL & " Max(ST.TName) as TName, Max(OA.Team) as Team"
sSQL = sSQL & " FROM " & OverAllScoresTableName & " as OA, " & SanctionTableName & " as ST"
sSQL = sSQL & " WHERE OA.SkiYearID = " & sProcessingYear & " and OA.TotalOverall > 0"
sSQL = sSQL & " and Left(OA.TourID,6) = Left(ST.TSanction,6)"
sSQL = sSQL & " GROUP BY MemberID, TourID, Div"
sSQL = sSQL & " ORDER BY MemberID, Div, MaxOAScore Desc"

WriteDebugSQL(sSQL)

rs.open sSQL, sConnectionToTRATable, 3, 3

tempvar = rs.getrows()
nTotalMembers = ubound(tempvar,2)

rs.MoveFirst
nProcessedSoFar = 0

' Outer loop of all overall scores
DO UNTIL rs.eof

  nProcessedSoFar = nProcessedSoFar + 1
  IF (nProcessedSoFar mod 10 = 5) THEN ShowProgress (nProcessedSoFar / nTotalMembers)

  TempMemberID = rs("MemberID"): TempTeam = trim(rs("Team")): TempDiv = rs("Div")
  nScoC = 0: nScoR = 0: TotScore = 0: RankExplain = ""
  
  ' Inner Loop of overall scores for this MemberID/Division
  ' Feb 2008 -- now implements the "Do No Harm" philosophy, by evaluating if
  ' the Ranking score including this new score, would be lower or not, 
  ' factoring in the penalty levels both before and after -- using separate
  ' counters for C and ELR scores and the new penalty matrix (Feb 2008).
 
  DO WHILE TempMemberID = rs("MemberID") AND TempDiv = rs("Div")

     IF TempTeam = "" THEN TempTeam = trim(rs("Team"))
     TempScore = left(rs("MaxOAScore"),8)

     IF nScoC+nScoR < 3 THEN     

        IF nScoC+nScoR = 0 THEN TempAdd = "Y" ELSE TempAdd = "N"

        IF Mid(rs("MaxOAScore"),9,1) >= "3" THEN

           IF TempAdd = "N" THEN IF (1-(tPenalty(nScoC,nScoR)/100))*(TotScore)/(nScoC+nScoR) < (1-(tPenalty(nScoC,nScoR+1)/100))*(TotScore+TempScore)/(nScoC+nScoR+1) THEN TempAdd = "Y"

           IF TempAdd = "Y" THEN
              nScoR = nScoR + 1
              TotScore = TotScore + TempScore
              RankExplain = RankExplain + FormatNumber(TempScore,1) + " (" + Right(rs("MaxOAScore"),1) + ")"
              IF rs("Div")<>rs("DivOrig") THEN RankExplain = RankExplain + " as " & rs("DivOrig")
              RankExplain = RankExplain + " from " + SQLClean(rs("TName")) + "&#13;"
           END IF

        ELSE

           IF TempAdd = "N" THEN IF (1-(tPenalty(nScoC,nScoR)/100))*(TotScore)/(nScoC+nScoR) < (1-(tPenalty(nScoC+1,nScoR)/100))*(TotScore+TempScore)/(nScoC+nScoR+1) THEN TempAdd = "Y"

           IF TempAdd = "Y" THEN
              nScoC = nScoC + 1
              TotScore = TotScore + TempScore
              RankExplain = RankExplain + FormatNumber(TempScore,1) + " (" + Right(rs("MaxOAScore"),1) + ")"
              IF rs("Div")<>rs("DivOrig") THEN RankExplain = RankExplain + " as " & rs("DivOrig")
              RankExplain = RankExplain + " from " + SQLClean(rs("TName")) + "&#13;"
           END IF

        END IF 
        
     END IF

   	rs.moveNEXT
   	IF rs.eof THEN exit do
  LOOP
  ' Bottom of inner loop	

	' ************** Calculate and Explain Penalty for Overall Ranking Score HERE ************

  TotScore = (1-(tPenalty(nScoC,nScoR)/100))*(TotScore)/(nScoC+nScoR)
  IF tPenalty(nScoC,nScoR) > 0 THEN
     RankExplain = RankExplain + "with " & tPenalty(nScoC,nScoR) & "% Penalty"
  ELSE
     RankExplain = RankExplain + "with NO Penalty"
  END IF	

  ' Finally insert this computed Overall Ranking Score into the RankScore table.

  sSQL = "INSERT INTO " & RankTableName & " (MemberID, Team, Event, Div, RankScore, RnkScoBkup, SkiYearID)"
  sSQL = sSQL & " VALUES ('" & TempMemberID & "','" & TempTeam & "', 'O', '" & TempDiv & "', '" & TotScore
  sSQL = sSQL & "', '" & RankExplain & "', " & sProcessingYear & ")"

  Con.Execute(sSQL)

LOOP
' Bottom of Outer Loop

rs.close

' ----------------------   End of OVERALL RANKING Calculations  ---------------------


' Bottom of IF to exclude the Overall Rankings section above from running
END IF





' Section # 4

IF RunEventRanks = "YES" OR RunOverride = "YES" THEN

Response.write(" Event Ranks ...")

' --------  Begin New Streamlined Event Ranking Logic, Using EquivScores Table  --------

' Delete all of the old Event Rank records for the Ski Year ID that we are processing.

sSQL = "DELETE FROM " & RankTableName & " WHERE Event <> 'O' and SkiYearID = " & sProcessingYear

' WriteDebugSQL(sSQL)

Con.Execute(sSQL)

' Query now with Max function and Group By clause, to roll to best single round from each TourID
' Revised Feb 2008 such that Score string is now char(10), with leftmost 8 being the
' actual Total Overall Score (formatted as 8,2), followed by 2-char prioritized class.

sSQL = "SELECT ES.MemberID, ES.TourID, ES.Div, ES.Event, Max("
sSQL = sSQL & " Substring(Cast(Cast(ES.Score+400000 as Decimal(8,2)) as Char(9)),2,8)"
sSQL = sSQL & " + ES.PrioClass) as MaxScore, Max(ES.DivOrig) as DivOrig,Max(ES.Place) as Place, Max(ES.Rating) as Rating,"
sSQL = sSQL & "  Max(ES.SkiYearID) as SkiYearID, Max(ST.TName) as TName, Max(ES.Team) as Team"
sSQL = sSQL & " FROM " & EquivScoresTableName & " as ES, " & SanctionTableName & " as ST"
sSQL = sSQL & " WHERE ES.SkiYearID = " & sProcessingYear 
sSQL = sSQL & " and Left(ES.TourID,6) = Left(ST.TSanction,6)"
sSQL = sSQL & " GROUP BY ES.MemberID, ES.TourID, ES.Div, ES.Event"
sSQL = sSQL & " ORDER BY ES.MemberID, ES.Div, ES.Event, MaxScore Desc"

WriteDebugSQL(sSQL)

rs.open sSQL, sConnectionToTRATable, 3, 3

tempvar = rs.getrows()
nTotalMembers = ubound(tempvar,2)

rs.MoveFirst
nProcessedSoFar = 0

' Outer loop of all Equivalent scores
DO UNTIL rs.eof

  nProcessedSoFar = nProcessedSoFar + 1
  IF (nProcessedSoFar mod 10 = 5) THEN ShowProgress (nProcessedSoFar / nTotalMembers)

  TempMemberID = rs("MemberID"): TempDiv = rs("Div")
  TempEvent = rs("Event"): TempTeam = trim(rs("Team"))
  TempRating = "   ": nScoR = 0: nScoC = 0: TotScore = 0: 
  RankExplain = "":   R_Ski = "": R_PLC = "": N_PLC = ""
  
  ' Inner Loop of Scores for this MemberID/Division/Event
  ' First Phase is to Add top 3 (or 2 for NCWSA) to Ranking score total
  
  DO WHILE TempMemberID = rs("MemberID") AND TempDiv = rs("Div") AND TempEvent = rs("Event")

     IF TempTeam = "" THEN TempTeam = trim(rs("Team"))
     TempScore = left(rs("MaxScore"),8)

     IF left(TempDiv,1)<>"C" THEN
     	
        ' AWSA Logic -- now Differential Penalty function of nScoC and nScoR

        IF nScoC+nScoR < 3 THEN     

           IF nScoC+nScoR = 0 THEN TempAdd = "Y" ELSE TempAdd = "N"

           IF Mid(rs("MaxScore"),9,1) >= "3" THEN

              IF TempAdd = "N" THEN IF (1-(tPenalty(nScoC,nScoR)/100))*(TotScore)/(nScoC+nScoR) < (1-(tPenalty(nScoC,nScoR+1)/100))*(TotScore+TempScore)/(nScoC+nScoR+1) THEN TempAdd = "Y"

              IF TempAdd = "Y" THEN
                 nScoR = nScoR + 1
                 TotScore = TotScore + TempScore
                 IF TempEvent = "S" THEN
                    RankExplain = RankExplain + FormatNumber(TempScore,2) + " (" + Right(rs("MaxScore"),1) + ")"
                 ELSE   
                    RankExplain = RankExplain + FormatNumber(TempScore,0) + " (" + Right(rs("MaxScore"),1) + ")"
                 END IF
                 IF rs("Div")<>rs("DivOrig") THEN RankExplain = RankExplain + " as " & rs("DivOrig")
                 RankExplain = RankExplain + " from " + SQLClean(rs("TName")) + "&#13;"
              END IF

           ELSE

              IF TempAdd = "N" THEN IF (1-(tPenalty(nScoC,nScoR)/100))*(TotScore)/(nScoC+nScoR) < (1-(tPenalty(nScoC+1,nScoR)/100))*(TotScore+TempScore)/(nScoC+nScoR+1) THEN TempAdd = "Y"

              IF TempAdd = "Y" THEN
                 nScoC = nScoC + 1
                 TotScore = TotScore + TempScore
                 IF TempEvent = "S" THEN
                    RankExplain = RankExplain + FormatNumber(TempScore,2) + " (" + Right(rs("MaxScore"),1) + ")"
                 ELSE   
                    RankExplain = RankExplain + FormatNumber(TempScore,0) + " (" + Right(rs("MaxScore"),1) + ")"
                 END IF
                 IF rs("Div")<>rs("DivOrig") THEN RankExplain = RankExplain + " as " & rs("DivOrig")
                 RankExplain = RankExplain + " from " + SQLClean(rs("TName")) + "&#13;"
              END IF

           END IF 

        END IF 

     ELSE

        ' NCWSA Logic -- function of nScoC alone

        IF nScoC < 2 THEN     
           nScoC = nScoC + 1
           TotScore = TotScore + TempScore
           IF TempEvent = "S" THEN
              RankExplain = RankExplain + FormatNumber(TempScore,2)
           ELSE   
              RankExplain = RankExplain + FormatNumber(TempScore,0)
           END IF
           RankExplain = RankExplain + " from " + SQLClean(rs("TName")) + "&#13;"
        END IF

     END IF 



   ' Second Phase is to pick up Nationals and Regionals Placements
   
     IF UCASE(RIGHT(TRIM(rs("TourID")),1)) = "A" THEN N_Plc = rs("Place")

     IF UCASE(RIGHT(TRIM(rs("TourID")),1)) = "B" THEN R_Ski = mid(rs("TourID"),3,1): R_Plc = rs("Place")
        
   	rs.moveNEXT
   	IF rs.eof THEN exit do
  LOOP
  
  ' Bottom of inner loop -- Now Finalize Ranking For this Member/Div/Event

  ' NCWSA Formulation First for "C" Divs
 	
  ' ************** Calculate and Explain Penalty for NCWSA Event Ranking Scores HERE ************

  IF left(TempDiv,1) = "C" THEN
  	 TotScore = TotScore / 2
     IF nScoC = 2 THEN 
  	    RankExplain = RankExplain + "with NO Penalty": 
     ELSE
        RankExplain = RankExplain + "with 50% Penalty"
     END IF
  
  ' AWSA Formulation Otherwise

  ' ************** Calculate and Explain Penalty for AWSA Event Ranking Scores HERE ************

  ELSE
 
     TotScore = (1-(tPenalty(nScoC,nScoR)/100))*(TotScore)/(nScoC+nScoR)
     IF tPenalty(nScoC,nScoR) > 0 THEN
        RankExplain = RankExplain + "with " & tPenalty(nScoC,nScoR) & "% Penalty"
     ELSE
        RankExplain = RankExplain + "with NO Penalty"
     END IF	

  END IF

  sSQL = "INSERT INTO " & RankTableName & " (MemberID, Team, Event, Div, SC_1, RankScore, RnkScoBkup, AWSA_Rat, Reg_SKI, Regl_Plc, Natl_Plc, SkiYearID)"
  sSQL = sSQL & " VALUES ('" & TempMemberID & "', '" & TempTeam & "','" & TempEvent & "', '" & TempDiv & "', '" & TotScore & "', '" & TotScore & "', '" & RankExplain
  sSQL = sSQL & "', '  ', '" & R_Ski & "', '" & R_Plc & "', '" & N_Plc & "', " & sProcessingYear & ")"

  Con.Execute(sSQL)

LOOP
' Bottom of Outer Loop

rs.close

' -------- End of New Streamlined Event Ranking Logic, Using EquivScores Table  --------


' Now that we are done doing the AWSA Skiers, 
' we can do the rankings for the NSL folks.
' Create NSL Rankings Rows by Insert Into Select from with Sum and Condition

sSQL = "INSERT INTO " & RankTableName
sSQL = sSQL & " (MemberID, Event, Div, Team, SC_3, SkiYearID)"
sSQL = sSQL & " SELECT MemberID, Event, Div, Max(Team), round(sum(NSL_Placement_Points),0) as SC_3, " & sProcessingYear & " as SkiYearID"
sSQL = sSQL & " from " & RawScoresTableName
sSQL = sSQL & " WHERE Enddate <= '" & FormatDateTime(sSkiYearEnd,2) & "' AND EndDate >= '" & FormatDateTime(sSkiYearBegin,2) & "' AND NSL_Placement_Points > 0"
sSQL = sSQL & " group by MemberID, Div, Event ORDER BY MemberID, Event"

' WriteDebugSQL(sSQL)

Con.Execute(sSQL)

' DONE SAVING NSL SCORES


' Bottom of IF to exclude the Event Ranking section above from running
END IF




' ---------   This is the NEW SQL-Based Ranking LEVEL logic -- Does S/T/J & Overall Levels   ---------

' Section # 5

IF RunLevelLogic = "YES" OR RunOverride = "YES" THEN

Response.write(" Levels ...")

' Overall process is based on a "Rank Sequence Numbers" table.  This is a temp 
' table, which establishes the number and range of scores for each Division 
' and Event, including Overall, for the Ski Year currently being processed.
' That table is then used to calculate the Rank Number and Percentile for each
' Score value for the Division and Event -- along with the (10-0) Level for now.
' These determinations are done with two sequential steps, and then the results
' in that temp table are posted to the main Rankings table.  As part of this
' process we also populate a CutOff table by Division and Event.  This contains
' the Level percentiles from the Divisions table, along with the corresponding
' CutOffAverage (COA) Scores actually found in the Rankings table, for each
' level, for the Division and Event.  All told this is done in xx steps below.

' First Empty the Rank Numbers Table and Reset the Identity Column base -- Drop 
' and re-Add the ID column.  Then also Delete any existing rows from the CutOff
' Table for the current Ski Year ID.

sSQL = "DELETE FROM " & RankNumsTableName & ";"
sSQL = sSQL & " Alter Table " & RankNumsTableName & " Drop Column RankSeq;"
sSQL = sSQL & " Alter Table " & RankNumsTableName & " Add RankSeq Int Identity(1,1);"
sSQL = sSQL & " DELETE FROM " & CutOffTableName & " WHERE SkiYearID = " & sProcessingYear & ";"

WriteDebugSQL(sSQL)

Con.Execute(sSQL)

' Next we populate the Rank Numbers Table for the current sProcessingYear.
' Note that this assigns the RankSeq Identity values automatically, and
' in order by RankScore within each Division and Event.  Only include
' Members where their Membership record Federation Code is "USA" ****.

sSQL = "INSERT INTO " & RankNumsTableName
sSQL = sSQL & " (MemberID, Event, Div, RankScore)"
sSQL = sSQL & " SELECT RT.MemberID, RT.Event, RT.Div, RT.RankScore"
sSQL = sSQL & " from " & RankTableName & " as RT, "& MemberTableName & " as MT" 
sSQL = sSQL & " WHERE RT.MemberID = MT.PersonIDWithCheckDigit"
sSQL = sSQL & " AND UPPER(MT.FederationCode) = 'USA'"
sSQL = sSQL & " AND RankScore is not null AND SkiYearID = " &sProcessingYear
sSQL = sSQL & " ORDER BY Div, Event, RankScore"

WriteDebugSQL(sSQL)

Con.Execute(sSQL)

' Next step is to update the Rank Numbers table, and calculate the Percentile 
' for each score, as a function of the RankSeq Number, within the scope of the
' Min and Max RankSeq values for the corresponding Division and Event.  We also 
' Store the actual RankNum (1 to Max) for each score, and for the moment also the
' Rank_Level for the old "Base 10 level logic" -- delete this later DJC 1/19/2008.

sSQL = "UPDATE RN SET RankPct = 100 * Cast(RM.RankSeq - RS.MinRank + 1 as Real) / Cast(RS.MaxRank - RS.MinRank + 1 as Real),"
sSQL = sSQL & " Rank_Level = 10 * Cast(RM.RankSeq - RS.MinRank + 1 as Real) / Cast(RS.MaxRank - RS.MinRank + 1 as Real),"
sSQL = sSQL & " RankNum = RS.MaxRank - RM.RankSeq + 1"
sSQL = sSQL & " FROM " & RankNumsTableName & " as RN,"
sSQL = sSQL & " (Select Div, Event, RankScore, Max(RankSeq) as RankSeq from " & RankNumsTableName & " GROUP BY Div, Event, RankScore) as RM,"
sSQL = sSQL & " (SELECT Div, Event, Min(RankSeq) as MinRank, Max(RankSeq) as MaxRank from " & RankNumsTableName & " GROUP BY Div, Event) as RS"
sSQL = sSQL & " WHERE RN.RankScore = RM.RankScore and RN.Div = RM.Div and RN.Event = RM.Event"
sSQL = sSQL & " AND RN.Div = RS.Div and RN.Event = RS.Event"

WriteDebugSQL(sSQL)

Con.Execute(sSQL)

' Next step is to prepare the CutOff Table entries for the current Ski Year ID,
' initially populating these rows by extracting the current Level Percentiles 
' from the Division Control Table, by Division and Event.  This recasts
' these parameters into a Division / Event keying framework, which is more
' compatible with the way the scores data is organized.

' sSQL = "Drop Table " & CutOffTableName & "; Create Table " & CutOffTableName & "(Div Char(2), Event Char(1), SkiYearID Int, Pct9 Int, COA9 real, Pct8 Int, COA8 real, Pct7 Int, COA7 real, Pct6 Int, COA6 real, Pct5 Int, COA5 real, Pct4 Int, COA4 real, Pct3 Int, COA3 real, Pct2 Int, COA2 real, Pct1 Int, COA1 real); "

sSQL = "INSERT INTO " & CutOffTableName & " (Div, Event, SkiYearID, Pct1, Pct2, Pct3,  Pct4, Pct5, Pct6, Pct7, Pct8, Pct9)"
sSQL = sSQL & " Select Div, 'S', " & sProcessingYear & ", Percent_01_S, Percent_02_S, Percent_03_S, Percent_04_S, Percent_05_S, Percent_06_S, Percent_07_S, Percent_08_S, Percent_09_S From " & DivisionsTableName & " WHERE left(Div,1) in ('B','G','M','W','O') and SkiYearID = " & sProcessingYear & " UNION"
sSQL = sSQL & " Select Div, 'T', " & sProcessingYear & ", Percent_01_T, Percent_02_T, Percent_03_T, Percent_04_T, Percent_05_T, Percent_06_T, Percent_07_T, Percent_08_T, Percent_09_T From " & DivisionsTableName & " WHERE left(Div,1) in ('B','G','M','W','O') and SkiYearID = " & sProcessingYear & " UNION"
sSQL = sSQL & " Select Div, 'J', " & sProcessingYear & ", Percent_01_J, Percent_02_J, Percent_03_J, Percent_04_J, Percent_05_J, Percent_06_J, Percent_07_J, Percent_08_J, Percent_09_J From " & DivisionsTableName & " WHERE left(Div,1) in ('B','G','M','W','O') and SkiYearID = " & sProcessingYear & " UNION"
sSQL = sSQL & " Select Div, 'O', " & sProcessingYear & ", Percent_01_O, Percent_02_O, Percent_03_O, Percent_04_O, Percent_05_O, Percent_06_O, Percent_07_O, Percent_08_O, Percent_09_O From " & DivisionsTableName & " WHERE left(Div,1) in ('B','G','M','W','O') and SkiYearID = " & sProcessingYear

WriteDebugSQL(sSQL)

Con.Execute(sSQL)

' Next we update the Rankings Table -- Match in the Pre-Calculated RankNum
' Value (True Rank from 1 to N), RankPct (True Percentile fm 100 down), and 
' for now also the old Rank_Level (fm 10 down).  Ties at the same RankScore 
' value, will get equal values for these three measures, since we Group 
' the Rank Numbers Table by RankScore.  Lastly we Classify the Ranking 
' Level, according to the Level Percentiles joined in from the CutOff Table,
' by Division and event, that we just created in the step immediately above.

sSQL = "UPDATE RT SET RankNum = RN.RankNum, RankPct = RN.RankPct, Rank_Level = RN.Rank_Level, AWSA_Rat = CASE "
sSQL = sSQL & " when RN.RankPct > CT.Pct9 then RN.Event + '9'"
sSQL = sSQL & " when RN.RankPct > CT.Pct8 then RN.Event + '8'"
sSQL = sSQL & " when RN.RankPct > CT.Pct7 then RN.Event + '7'"
sSQL = sSQL & " when RN.RankPct > CT.Pct6 then RN.Event + '6'"
sSQL = sSQL & " when RN.RankPct > CT.Pct5 then RN.Event + '5'"
sSQL = sSQL & " when RN.RankPct > CT.Pct4 then RN.Event + '4'"
sSQL = sSQL & " when RN.RankPct > CT.Pct3 then RN.Event + '3'"
sSQL = sSQL & " when RN.RankPct > CT.Pct2 then RN.Event + '2'"
sSQL = sSQL & " when RN.RankPct > CT.Pct1 then RN.Event + '1'"
sSQL = sSQL & " else RN.Event + '0' end"
sSQL = sSQL & " FROM " & RankTableName & " as RT, " & CutOffTableName & " as CT, " & MemberTableName & "	as	MT,"
sSQL = sSQL & " (Select Div, Event, RankScore, Min(RankNum) as RankNum, Max(RankPct) as RankPct, Max(Rank_Level) as Rank_Level"
sSQL = sSQL & " FROM " & RankNumsTableName & " GROUP BY Div, Event, RankScore) as RN"
sSQL = sSQL & " WHERE RT.Div = CT.Div and RT.Event = CT.Event and CT.SkiYearID = " & sProcessingYear
sSQL = sSQL & " AND RT.RankScore = RN.RankScore and RT.Div = RN.Div and RT.Event = RN.Event"
sSQL = sSQL & " AND RT.MemberID = MT.PersonIDWithCheckDigit"
sSQL = sSQL & " AND UPPER(MT.FederationCode) = 'USA'"
sSQL = sSQL & " AND RT.RankScore is not Null and RT.SkiYearID = " & sProcessingYear

WriteDebugSQL(sSQL)

Con.Execute(sSQL)

' Final step is to recap the Ranking Level Cut Off Scores, into the Cut Off Table, 
' using the Min(RankScore) values found for each of the 9 levels, for each division 
' and event, for the current Ski Year ID.  For each level,  we use COALESCE to find
' the first existing equal or lower level cutoff score, or zero if none exists.

sSQL = "UPDATE CT Set COA1 = Coalesce(LT1.Cutoff,0),"
sSQL = sSQL & " COA2 = Coalesce(LT2.Cutoff,LT1.Cutoff,0),"
sSQL = sSQL & " COA3 = Coalesce(LT3.Cutoff,LT2.Cutoff,LT1.Cutoff,0),"
sSQL = sSQL & " COA4 = Coalesce(LT4.Cutoff,LT3.Cutoff,LT2.Cutoff,LT1.Cutoff,0),"
sSQL = sSQL & " COA5 = Coalesce(LT5.Cutoff,LT4.Cutoff,LT3.Cutoff,LT2.Cutoff,LT1.Cutoff,0),"
sSQL = sSQL & " COA6 = Coalesce(LT6.Cutoff,LT5.Cutoff,LT4.Cutoff,LT3.Cutoff,LT2.Cutoff,LT1.Cutoff,0),"
sSQL = sSQL & " COA7 = Coalesce(LT7.Cutoff,LT6.Cutoff,LT5.Cutoff,LT4.Cutoff,LT3.Cutoff,LT2.Cutoff,LT1.Cutoff,0),"
sSQL = sSQL & " COA8 = Coalesce(LT8.Cutoff,LT7.Cutoff,LT6.Cutoff,LT5.Cutoff,LT4.Cutoff,LT3.Cutoff,LT2.Cutoff,LT1.Cutoff,0),"
sSQL = sSQL & " COA9 = Coalesce(LT9.Cutoff,LT8.Cutoff,LT7.Cutoff,LT6.Cutoff,LT5.Cutoff,LT4.Cutoff,LT3.Cutoff,LT2.Cutoff,LT1.Cutoff,0)"
sSQL = sSQL & " FROM " & CutOffTableName & " as CT "
sSQL = sSQL & " LEFT JOIN (Select Div, Event, Min(RankScore) as Cutoff FROM " & RankTableName & " Where SkiYearID = " & sProcessingYear & " and RankNum is not Null and Right(AWSA_Rat,1)='9' Group by Div, Event) as LT9 on LT9.Div=CT.Div and LT9.Event=CT.Event"
sSQL = sSQL & " LEFT JOIN (Select Div, Event, Min(RankScore) as Cutoff FROM " & RankTableName & " Where SkiYearID = " & sProcessingYear & " and RankNum is not Null and Right(AWSA_Rat,1)='8' Group by Div, Event) as LT8 on LT8.Div=CT.Div and LT8.Event=CT.Event"
sSQL = sSQL & " LEFT JOIN (Select Div, Event, Min(RankScore) as Cutoff FROM " & RankTableName & " Where SkiYearID = " & sProcessingYear & " and RankNum is not Null and Right(AWSA_Rat,1)='7' Group by Div, Event) as LT7 on LT7.Div=CT.Div and LT7.Event=CT.Event"
sSQL = sSQL & " LEFT JOIN (Select Div, Event, Min(RankScore) as Cutoff FROM " & RankTableName & " Where SkiYearID = " & sProcessingYear & " and RankNum is not Null and Right(AWSA_Rat,1)='6' Group by Div, Event) as LT6 on LT6.Div=CT.Div and LT6.Event=CT.Event"
sSQL = sSQL & " LEFT JOIN (Select Div, Event, Min(RankScore) as Cutoff FROM " & RankTableName & " Where SkiYearID = " & sProcessingYear & " and RankNum is not Null and Right(AWSA_Rat,1)='5' Group by Div, Event) as LT5 on LT5.Div=CT.Div and LT5.Event=CT.Event"
sSQL = sSQL & " LEFT JOIN (Select Div, Event, Min(RankScore) as Cutoff FROM " & RankTableName & " Where SkiYearID = " & sProcessingYear & " and RankNum is not Null and Right(AWSA_Rat,1)='4' Group by Div, Event) as LT4 on LT4.Div=CT.Div and LT4.Event=CT.Event"
sSQL = sSQL & " LEFT JOIN (Select Div, Event, Min(RankScore) as Cutoff FROM " & RankTableName & " Where SkiYearID = " & sProcessingYear & " and RankNum is not Null and Right(AWSA_Rat,1)='3' Group by Div, Event) as LT3 on LT3.Div=CT.Div and LT3.Event=CT.Event"
sSQL = sSQL & " LEFT JOIN (Select Div, Event, Min(RankScore) as Cutoff FROM " & RankTableName & " Where SkiYearID = " & sProcessingYear & " and RankNum is not Null and Right(AWSA_Rat,1)='2' Group by Div, Event) as LT2 on LT2.Div=CT.Div and LT2.Event=CT.Event"
sSQL = sSQL & " LEFT JOIN (Select Div, Event, Min(RankScore) as Cutoff FROM " & RankTableName & " Where SkiYearID = " & sProcessingYear & " and RankNum is not Null and Right(AWSA_Rat,1)='1' Group by Div, Event) as LT1 on LT1.Div=CT.Div and LT1.Event=CT.Event"
sSQL = sSQL & " WHERE CT.SkiYearID = " & sProcessingYear

WriteDebugSQL(sSQL)

Con.Execute(sSQL)


' Bottom of IF to exclude above section for debugging purposes
END IF

' -------------------  END of NEW SQL-Based Ranking LEVEL logic  -----------------------------------





' MUST run this section as this is the cue to STOP processsing.



' Finally we reset (turn off) the Recalculation Underway flag for the Ski Year just processsed.

sSQL = "UPDATE " & SkiYearTableName & " set RecalcUnderway = 0 WHERE SkiYearID = " & sProcessingYear
Con.Execute(sSQL)

Response.write(" Done !!<br>")



' ---------------  Determines if all desired records have been processed   --------------------
' IF the year we just did was 1 (which is the 12 month recalc)
' or IF the year was specified in a request, THEN we can go ahead
' and stop processing.  

IF (sProcessingYear = 1) or (trim(request("skiyear")) <> "") THEN
  sProcessingYear = "STOP"
ELSE
' Otherwise we set the year to 1 which does the 12 month recalc
' and we loop again.

  sSQL = "SELECT top 1 * from " & SkiYearTableName & " WHERE skiyearid = 1"

	SET rs=Server.CreateObject("ADODB.recordset")
  rs.open sSQL, SConnectionToTRATable, 3, 3  

    ' IF the specified year doesn't exist, THEN someone messed up! :)
    IF rs.EOF THEN
      session("message") = "Ski Year ID " & request("skiyear") & " was not found."
      WriteLog ("Ski Year ID " & request("skiyear") & " was not found.")
      Response.Redirect("/?process=logout")
    ELSE
    '  IF there is, THEN we save all the variables we will need.
      sProcessingYear = rs("SkiYearID")
      sPrevYear = rs("PrevYearID")
      sSkiYearName = rs("SkiYearName")
      sSkiYearBegin = rs("BeginDate")
      sSkiYearEnd = rs("EndDate")    
    END IF

  rs.Close
END IF


' ---------------------------------------------------------------------
' This is the big loop which sends us back to recalc the NEXT ski year.
' We basically do the current default year followed by the 12 month.
' After the 12 month we just do a stop to end the loop. 

LOOP   



set rs = nothing
CloseCon

FinishProgress

'----- Email Any Errors ------
IF EmailErrors <> 0 THEN
  myMail.Send
  set myMail=nothing
END IF

WriteLog(date() &"  "& time() &"  Ranking Recalculations Completed Successfully.")


' ************** Bottom of Rankings Recalculation Conditional ***************

END IF

timeNow = milliDif()

' ----- Email Results recap, only if no Equival request value -----

' ----- Hence this emails ONLY for the nightly scheduled unattended run.
' ----- Otherwise if run manually, then recap will appear on user's screen
' ----- as process runs, and no need to e-mail -- copy/paste/save if desired.

IF Request("Equival") = "" THEN

Set myMail=CreateObject("CDO.Message")
myMail.Subject="TRAWEB - Nightly Run Recap"
myMail.From="TRAWeb@USAWaterSki.ORG"
myMail.To=EMailToWho
IF Err.Number = 0 THEN

'   sSQL = "This is an automated message to report that the Nightly ePolk Run" & VBCrLf
'   sSQL = sSQL & "has been successfully completed.  Member Extract & Update Recap:" & VBCrLf & VBCrLf
'   sSQL = sSQL & Formatnumber(nHQExt,0) & " Member data rows supplied from HQ Server" & VBCrLf
'   sSQL = sSQL & Formatnumber(nLocal,0) & " Member rows found in ePolk Server Table" & VBCrLf
'   sSQL = sSQL & Formatnumber(nUpdates,0) & " Member rows updated with new Data" & VBCrLf
'   sSQL = sSQL & Formatnumber(nInserts,0) & " New Member rows added" & VBCrLf
'   sSQL = sSQL & Formatnumber(nDeletes,0) & " Old Member rows deleted" & VBCrLf
'   sSQL = sSQL & Formatnumber(nLocal+nInserts-nDeletes,0) & " Member rows now in ePolk Server Table" & VBCrLf
'   IF Len(DupMemList) > 0 THEN sSQL = sSQL & VCCrLf & "!! Duplicate Person ID's encountered: " & DupMemList & VBCrLf
'   sSQL = sSQL & VBCrLf & Formatnumber(nScoUpdts,0) & " Scores updated to consolidated Member IDs"
'   sSQL = sSQL & VBCrLf & Formatnumber(nOvrUpdts,0) & " Overalls updated to consolidated Member IDs"
'   sSQL = sSQL & VBCrLf & Formatnumber(nRnkUpdts,0) & " Rankings updated to consolidated Member IDs"
'   sSQL = sSQL & VBCrLf & VBCrLf & "Default Ski Year plus 12 Month Rankings have been Recalculated"
'   sSQL = sSQL & VBCrLf & VBCrLf & "Overall Process ... Start: " & StartTime & ", Finish: " & Time()
'   myMail.TextBody = sSQL

   IF Len(DupMemList) > 0 THEN DupMemList = "!! Duplicate Person ID's encountered: " & DupMemList & VBCrLf

   myMail.TextBody="This is an automated message to report that the Nightly ePolk Run" & VBCrLf _
      & "has been successfully completed.  Member Extract & Update Recap:" & VBCrLf & VBCrLf _
      & Formatnumber(nHQExt,0) & " Member data rows supplied from HQ Server" & VBCrLf & DupMemList _
      & Formatnumber(nLocal,0) & " Member rows found in ePolk Server Table" & VBCrLf _
      & Formatnumber(nUpdates,0) & " Member rows updated with new Data" & VBCrLf _
      & Formatnumber(nInserts,0) & " New Member rows added" & VBCrLf _
      & Formatnumber(nDeletes,0) & " Old Member rows deleted" & VBCrLf _
      & Formatnumber(nLocal+nInserts-nDeletes,0) & " Member rows now in ePolk Server Table" & VBCrLf _
  		& VBCrLf & Formatnumber(nScoUpdts,0) & " Scores updated to consolidated Member IDs" _
  		& VBCrLf & Formatnumber(nOvrUpdts,0) & " Overalls updated to consolidated Member IDs" _
  		& VBCrLf & Formatnumber(nRnkUpdts,0) & " Rankings updated to consolidated Member IDs" _
  		& VBCrLf & Formatnumber(nOffUpdts,0) & " Officials updated to consolidated Member IDs" _
      & VBCrLf & VBCrLf & "Default Ski Year plus 12 Month Rankings have been Recalculated" _
      & VBCrLf & VBCrLf & "Overall Process ... Start: " & StartTime & ", Finish: " & Time()

ELSE
   myMail.TextBody="This is an automated message to report that the Nightly Run has been " _
      & "completed but there were errors detected during the processing. " _
      & "<br>&nbsp;<br><B>Page Error Object</B><BR>" _
      & "Error Number: " & response.write(Err.Number) & " <BR> " _
      & "Description: " & response.write(Err.Description) & " <BR>	" _
      & "Source: " & response.write(Err.Source) & " <BR> " _
      & "Line Number: " & response.write(Err.Line) & " <BR>"
END IF

myMail.Send
set myMail=nothing

END IF

%>

<br>Nightly Update / Ranking Recalc Completed.<br>&nbsp;<br>
Total Time to Process: <%=elapsedpretty(timeNow - timeTHEN)%><br>&nbsp;<br>

<a href="/rankings/defaultHQ.asp">Return to Main Menu</a></center>
</BODY>
</HTML>
              




