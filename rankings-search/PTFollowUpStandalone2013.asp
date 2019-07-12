<% Option Explicit %>
<% IF Session("adminmenulevel")<10 THEN Response.Redirect "DefaultHQ.asp?process=login" %>

<!--#include file="settingsHQ.asp"-->

<% WriteIndexPageHeader %>

		<table border="0" cellspacing="1" cellpadding="1">

	<tr>
		<td>&nbsp;&nbsp;&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;&nbsp;&nbsp;</td>
	</tr>

	<tr>
		<td>&nbsp;&nbsp;&nbsp;</td>
		<td valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" Size="3">
    	
<%

Dim objRS, NumTours, NumEmails, sSQL
Dim TName, TDateE, TournAppID, TSanction, TStatus, TSanType, SptsGrpID, TRegion, WksOld
Dim PTF_SBK, PTF_WSP, PTF_TS, PTF_OD, PTF_BT, PTF_JT
Dim PTF_CS, PTF_CJ, PTF_SD, PTF_TU, PTF_HD, PTF_TNY
Dim eMailTo, eMailFrom, eMailCC, eMailBCC, eMailBody, strMissing, SeedRep, SeedFName

NumTours = 0: NumEmails = 0


%><p>Begin Post Tournament Followup Simulation for <%=FormatDateTime(date(),1)%>.</p><%

Set objRS = Server.CreateObject("ADODB.recordset")

SetupEmailService

'	First we generate the select query, to pull tournaments for followup.

'	Please note the DateDiff(wk,dateadd(d,-1,TDateE),GetDate()) logic --
'	what this accomplishes is to shift the "week boundary" to Sun/Mon,
'	where DateDiff normally uses Sat/Sun as the boundary.  Hence all events
'	with end dates thru Sunday will all be in the "Same Week".

sSQL = "Select ST.TournAppID, ST.TSanction, ST.SptsGrpID, ST.TRegion, ST.TName, ST.TDateE,"
sSQL = sSQL & " DateDiff(wk,DateAdd(d,-1,ST.TDateE),GetDate()) as WksOld,"
sSQL = sSQL & " ST.TStatus, ST.TDirName, ST.TDirEMail, Coalesce(PT.PTF_SBK,-1)"
sSQL = sSQL & " AS PTF_SBK, PT.PTF_WSP, PT.PTF_TS, PT.PTF_OD, PT.PTF_BT, PT.PTF_JT,"
sSQL = sSQL & " PT.PTF_CS, PT.PTF_CJ, PT.PTF_SD, PT.PTF_TU, PT.PTF_HD, PT.PTF_TNY,"

sSQL = sSQL & " CJ.CJudgName, CJ.CJudgEMail, CC.CScorName, CC.CScorEMail,"
sSQL = sSQL & " ST.TStatus, ST.TSanType FROM " & SanctionTableName & " ST LEFT JOIN "

sSQL = sSQL & PostTourTableName & " PT on PT.TournAppID = ST.TournAppID "

sSQL = sSQL & " LEFT JOIN (Select SX.TournAppID, MT.FirstName + ' ' + MT.LastName"
sSQL = sSQL & " as CJudgName, MT.Email as CJudgEMail FROM " & MemberTableName
sSQL = sSQL & " MT JOIN (Select TournAppID, Cast(case when len(CJudgePID)<9 then"
sSQL = sSQL & " CJudgePID else right(CJudgePID,8) end as integer) as PID FROM "
sSQL = sSQL & TRegSetupTableName & " WHERE isnumeric(CJudgePID) = 1)"
sSQL = sSQL & " SX on SX.PID = MT.PersonID WHERE patindex('%@%',Email) > 0 )"
sSQL = sSQL & " CJ ON CJ.TournAppID = ST.TournAppID"

sSQL = sSQL & " LEFT JOIN (Select SX.TournAppID, MT.FirstName + ' ' + MT.LastName"
sSQL = sSQL & " as CScorName, MT.Email as CScorEMail FROM " & MemberTableName
sSQL = sSQL & " MT JOIN (Select TournAppID, Cast(case when len(CScorePID)<9 then"
sSQL = sSQL & " CScorePID else right(CScorePID,8) end as integer) as PID FROM "
sSQL = sSQL & TRegSetupTableName & " WHERE isnumeric(CScorePID) = 1)"
sSQL = sSQL & " SX on SX.PID = MT.PersonID WHERE patindex('%@%',Email) > 0 )"
sSQL = sSQL & " CC ON CC.TournAppID = ST.TournAppID"

sSQL = sSQL & " WHERE ST.TStatus in (2,4) and ST.Deleted = 0"
sSQL = sSQL & " AND substring(ST.TournAppID,3,1) in ('C','E','M','S','W','U')"
sSQL = sSQL & " AND substring(ST.TSanction,7,1) in ('F','C','E','L','R','A','B')"
sSQL = sSQL & " AND (DateDiff(wk,DateAdd(d,-1,ST.TDateE),GetDate()) in (2,3,4,6,8,12,16,20)"
sSQL = sSQL & " OR (substring(ST.TSanction,7,1) = 'F' and ST.Tyear = 2012)"
sSQL = sSQL & " OR (DateDiff(wk,DateAdd(d,-1,ST.TDateE),GetDate())=1 and ST.TStatus=2))"
sSQL = sSQL & " order by ST.TDateE, ST.TournAppID"

'	WriteDebugSQL(sSQL)

objRS.open sSQL, sConnectionToTRATable, 3, 3

'	Finally we process the resulting record-set, generating and
'	sending a customized email note to each such tournament, but
'	only if we have at least one email address present.

IF NOT objRS.eof THEN
	
	objRS.MoveFirst
	DO Until objRS.eof

		NumTours = NumTours + 1

		TournAppID = objRS("TournAppID")
		TSanction = objRS("TSanction")
		TName = objRS("TName")
		TDateE = Replace(FormatDateTime(objRS("TDateE"),2),"/","-")
		IF Mid(TDateE,2,1) = "-" THEN TDateE = "0" & TDateE
		IF Mid(TDateE,5,1) = "-" THEN TDateE = Left(TDateE,3) & "0" & Right(TDateE,6)
		TSanType = objRS("TSanType")
		TRegion = objRS("TRegion")
		SptsGrpID = objRS("SptsGrpID")
		TStatus = objRS("TStatus")
		WksOld = objRS("WksOld")
		PTF_SBK = objRS("PTF_SBK")

		%><p><%=TSanction%>&nbsp;&nbsp; (<%=SptsGrpID%>/<%=TRegion%>/<%=TSanType%>)&nbsp;&nbsp; <%=TName%>&nbsp;&nbsp; (<%=TDateE%>, <%=WksOld%> wks old)<br><%

		'	First we establish the primary email address string -- TD / CJ / CC

		eMailTo = ""

		IF len(objRS("TDirEMail")) > 0 THEN
			eMailTo = """" & objRS("TDirName") & """ <" & objRS("TDirEMail") & ">"
		END IF

		IF len(objRS("CJudgEmail")) > 0 and instr(eMailTo,objRS("CJudgName")) = 0 THEN
			IF len(eMailTo) > 0 THEN eMailTo = eMailTo & "; "
			eMailTo = eMailTo & """" & objRS("CJudgName") & """ <" & objRS("CJudgEmail") & ">"
		END IF

		IF len(objRS("CScorEmail")) > 0 and instr(eMailTo,objRS("CScorName")) = 0 THEN
			IF len(eMailTo) > 0 THEN eMailTo = eMailTo & "; "
			eMailTo = eMailTo & """" & objRS("CScorName") & """ <" & objRS("CScorEmail") & ">"
		END IF

			'	next we establish secondary addressing based on jurisdiction code

		IF mid(TSanction,3,1) = "C" THEN
			eMailCC = """Bob Mayhew"" <skident@gmail.com>"
			eMailFrom = """Melissa Huitt"" <melissaskier@gmail.com>": eMailBCC = eMailFrom
			SeedRep = "Melissa Huitt" & vbCrLf & "AWSA South Central Seeding" & vbCrLf & "melissaskier@gmail.com" & vbCrLf & "(832) 746-0626"
		ELSEIF mid(TSanction,3,1) = "M" THEN
	   	eMailCC = """Kate Knafla"" <dr.kate@hotmail.com>"
	   	eMailFrom = """Dave Clark"" <awsatechdude@comcast.net>": eMailBCC = eMailFrom
			SeedRep = "Dave Clark" & vbCrLf & "AWSA Midwest Seeding" & vbCrLf & "awsatechdude@comcast.net" & vbCrLf & "(847) 269-7041"
		ELSEIF mid(TSanction,3,1) = "E" THEN
	   	eMailCC = """Pat Byrne"" <pjbyrn@aol.com>"
	   	eMailFrom = """Jennifer Frederick-Kelly"" <jennifer@frederickmachine.com>": eMailBCC = eMailFrom
			SeedRep = "Jennifer Frederick-Kelly" & vbCrLf & "AWSA East Seeding" & vbCrLf & "jennifer@frederickmachine.com" & vbCrLf & "(716) 892-1425"
		ELSEIF mid(TSanction,3,1) = "S" THEN
	   	eMailCC = """Bob Archambeau"" <evp@awsasouth.org>"
	   	eMailFrom = """Kirby Whetsel"" <kwhetsel@charter.net>": eMailBCC = eMailFrom
			SeedRep = "Kirby Whetsel" & vbCrLf & "AWSA South Seeding" & vbCrLf & "kwhetsel@charter.net" & vbCrLf & "(931) 409-0389"
		ELSEIF mid(TSanction,3,1) = "W" THEN
	   	eMailCC = """Elaine Bush"" <elainebush@att.net>"
	   	eMailFrom = """Judy Stanford"" <judy-don@sbcglobal.net>": eMailBCC = eMailFrom
			SeedRep = "Judy Stanford" & vbCrLf & "AWSA West Seeding" & vbCrLf & "judy-don@sbcglobal.net" & vbCrLf & "(925) 932-7781"
		ELSEIF mid(TSanction,3,1) = "U" THEN
	   	eMailCC = """Jeff Surdej"" <j_surdej@yahoo.com>"
	   	eMailFrom = """Dave Clark"" <awsatechdude@comcast.net>": eMailBCC = eMailFrom
			SeedRep = "Dave Clark" & vbCrLf & "NCWSA Seeding" & vbCrLf & "awsatechdude@comcast.net" & vbCrLf & "(847) 269-7041"
		ELSE
			eMailCC = """Dave Clark"" <awsatechdude@comcast.net>"
	   	eMailFrom = """USA Water Ski Competition"" <bhoyland@usawaterski.org>"
			SeedRep = "Britt Hoyland" & vbCRLF & "Competition & Events Coordinator" & vbCRLF & "bhoyland@usawaterski.org" & vbCRLF & "1-863-324-4341 ext 121"
		END IF

		IF len(eMailTo) > 0 THEN
		
			'	Then if we actually have primary address(es), we go on to prepare the email body

			'	First we compile a list of the missing items for the eMail body
			
			strMissing = ""
			
			IF TStatus = 2 THEN
				strMissing = vbCrLf & "   Entire collection of required Post-Tournament Reports"  
			ELSEIF PTF_SBK > -1 THEN

				IF PTF_SBK = 0 THEN 
					strMissing = strMissing & vbCrLf & "   Full Scorebook Report           "
					strMissing = strMissing & "( " & TSanction & ".SBK )"
				END IF

				IF objRS("PTF_WSP") = 0 THEN 
					strMissing = strMissing & vbCrLf & "   Rankings Data File (Scores)     "
					strMissing = strMissing & "( " & TSanction & ".WSP )"
				END IF

				IF objRS("PTF_TS") = 0 THEN 
					strMissing = strMissing & vbCrLf & "   Tournament Summary Report       "
					strMissing = strMissing & "( " & TournAppID & "TS.PRN )"
				END IF

				IF objRS("PTF_OD") = 0 THEN 
					strMissing = strMissing & vbCrLf & "   Officials Data File (Credits)   "
					strMissing = strMissing & "( " & TournAppID & "OD.TXT )"
				END IF

				IF objRS("PTF_BT") = 0 THEN 
					strMissing = strMissing & vbCrLf & "   Boat Time Tracking Report       "
					strMissing = strMissing & "( " & TournAppID & "BT.PRN )"
				END IF

				IF objRS("PTF_JT") = 0 THEN 
					strMissing = strMissing & vbCrLf & "   Jump Time Data File             "
					strMissing = strMissing & "( " & TournAppID & "JT.CSV )"
				END IF

				IF objRS("PTF_CS") = 0 THEN 
					strMissing = strMissing & vbCrLf & "   Condensed Scorebook Report      "
					strMissing = strMissing & "( " & TournAppID & "CS.HTM )"
				END IF

				IF objRS("PTF_CJ") = 0 THEN 
					strMissing = strMissing & vbCrLf & "   Chief Judges Tournament Report  "
					strMissing = strMissing & "( " & TournAppID & "CJ.PDF )"
				END IF

				IF objRS("PTF_SD") = 0 THEN 
					strMissing = strMissing & vbCrLf & "   Safety Directors Report         "
					strMissing = strMissing & "( " & TournAppID & "SD.PDF )"
				END IF

				IF objRS("PTF_TU") = 0 THEN 
					strMissing = strMissing & vbCrLf & "   Towboat Utilization Report      "
					strMissing = strMissing & "( " & TournAppID & "TU.PDF )"
				END IF

				IF objRS("PTF_HD") = 0 THEN 
					strMissing = strMissing & vbCrLf & "   Homologation Dossier            "
					strMissing = strMissing & "( " & TournAppID & "HD.TXT )"
				END IF

				IF objRS("PTF_TNY") = 0 THEN 
					strMissing = strMissing & vbCrLf & "   Tournament Settings "
					strMissing = strMissing & "( WSPARM.TNY or WWPARM.TXT )"
				END IF

			ELSE
				strMissing = vbCrLf & "   Entire collection of required Post-Tournament Reports"  
			END IF

			'	Now we construct the actual body of the eMail note, depending on vintage and status.

			objMessage.To = eMailTo
			objMessage.CC = eMailCC
			
			%>Email to:&nbsp;&nbsp; <%=replace(replace(eMailTo,"<","&lt;"),">","&gt;")%> <%

			objMessage.Subject = "Post-Tournament Reports from " & TSanction & " " & TName & " (" & TDateE & ")"

			eMailBody = "Dear Tournament Organizer and/or Chief Official(s) --" & vbCRLF & vbCRLF

			IF WksOld = 1 AND TStatus = 2 THEN

				eMailBody = eMailBody & "The post-tournament reports from " & TSanction & " " & TName & vbCRLF
				eMailBody = eMailBody & TDateE & " have not yet appeared in the Sanction control system." & vbCRLF & vbCRLF

				eMailBody = eMailBody & "USA Water Ski's real-time ranking database has matured into a" & vbCRLF
				eMailBody = eMailBody & "qualifications and ranking platform that now serves many time-" & vbCRLF
				eMailBody = eMailBody & "critical purposes.  We hope you can help us meet those objectives," & vbCRLF
				eMailBody = eMailBody & "by submitting the WSTIMS Zip file for this event at your earliest" & vbCRLF
				eMailBody = eMailBody & "convenience." & vbCRLF & vbCRLF

				eMailBody = eMailBody & "The Zip file should be emailed to me at the email address shown" & vbCrLf 
				eMailBody = eMailBody & "below.  If you are having difficulty producing any of those reports," & vbCrLf 
				eMailBody = eMailBody & "or the Zip file itself, please contact me for assistance." & vbCrLf & vbCrLf

				eMailBody = eMailBody & "If the above-referenced competition did not actually take place as" & vbCRLF
				eMailBody = eMailBody & "planned, then please advise us so that we can revise the status" & vbCRLF
				eMailBody = eMailBody & "of your tournament in the Sanction control system." & vbCRLF & vbCRLF

			ELSEIF TStatus = 2 THEN

				eMailBody = eMailBody & "The post-tournament reports from " & TSanction & " " & TName & vbCRLF
				eMailBody = eMailBody & TDateE & " have not yet been posted in the Sanction control system." & vbCRLF & vbCRLF

				eMailBody = eMailBody & "It has now been " & WksOld & " weeks since the above-referenced event was" & vbCRLF 
				eMailBody = eMailBody & "scheduled to take place, according to our records.  If the" & vbCRLF 
				eMailBody = eMailBody & "competition did NOT actually take place as planned, then please" & vbCRLF 
				eMailBody = eMailBody & "advise us so that we can revise the status of your tournament in" & vbCRLF 
				eMailBody = eMailBody & "the Sanction control system." & vbCRLF & vbCRLF

				eMailBody = eMailBody & "However, if the competition DID run as planned, then you need to" & vbCrLf
				eMailBody = eMailBody & "get the post-tournament reports in soon.  The WSTIMS Zip file" & vbCrLf
				eMailBody = eMailBody & "should be emailed to me at the email address shown below.  If you" & vbCrLf
				eMailBody = eMailBody & "are having difficulty producing any of those reports, or the Zip" & vbCrLf
				eMailBody = eMailBody & "file itself, please contact me for assistance." & vbCrLf & vbCrLf

				IF WksOld > 8 THEN
					eMailBody = eMailBody & "You need to be aware that any subsequent sanction applications for" & vbCRLF 
					eMailBody = eMailBody & "your organization cannot be approved, until these missing reports" & vbCRLF
					eMailBody = eMailBody & "have been received and checked off." & vbCRLF & vbCRLF
				END IF

			ELSE

				eMailBody = eMailBody & "The post-tournament reports from " & TSanction & " " & TName & vbCRLF
				eMailBody = eMailBody & TDateE & " are not completely posted in the Sanction control system." & vbCRLF & vbCRLF

				eMailBody = eMailBody & "While some of the required reports have been received and checked" & vbCRLF
				eMailBody = eMailBody & "off, the following items are still outstanding as of today:" & vbCRLF

				eMailBody = eMailBody & strMissing & vbCRLF & vbCRLF

				eMailBody = eMailBody & "If these missing items have been mailed in the past few days and" & vbCRLF
				eMailBody = eMailBody & "this eMail is crossing your package in the mail, please disregard" & vbCRLF 
				eMailBody = eMailBody & "this notice.  Otherwise, please note that it has now been " & WksOld & " weeks" & vbCRLF
				eMailBody = eMailBody & "since the above-referenced event took place, and these missing items" & vbCRLF 
				eMailBody = eMailBody & "need to be submitted before we can mark your event as complete." & vbCRLF & vbCRLF

				IF WksOld > 8 THEN
					eMailBody = eMailBody & "You need to be aware that any subsequent sanction applications for" & vbCRLF 
					eMailBody = eMailBody & "your organization cannot be approved, until these missing reports" & vbCRLF
					eMailBody = eMailBody & "have been received and checked off." & vbCRLF & vbCRLF
				END IF

				eMailBody = eMailBody & "Emailing these items to me in electronic form is preferred.  However," & vbCrLf 
				eMailBody = eMailBody & "if you have paper documents instead, then you should send those to" & vbCrLf 
				eMailBody = eMailBody & "Britt Hoyland at USA Waterski HQ, by postal mail.  If you are having" & vbCrLf 
				eMailBody = eMailBody & "difficulty producing any of those reports, or the WSTIMS Zip file" & vbCrLf 
				eMailBody = eMailBody & "itself, please contact me for assistance." & vbCrLf & vbCrLf
	
			END IF

			eMailBody = eMailBody & "Thank you for your hard work and continued support," & vbCRLF & vbCRLF

			eMailBody = eMailBody & SeedRep

			'	Now finally we send the constructed eMail message
	
			objMessage.TextBody = eMailBody

			IF TournAppID <> "10X123" THEN
				' objMessage.Send
				NumEmails = NumEmails + 1
			END IF

			%>Email to:&nbsp;&nbsp; <%=replace(replace(eMailTo,"<","&lt;"),">","&gt;")%> <%
			
		ELSE

			%>No eMail addresses found for this tournament.<%

		END IF

			%><br>From:&nbsp;&nbsp; <%=replace(replace(eMailFrom,"<","&lt;"),">","&gt;")%>
			<br>CC To:&nbsp;&nbsp; <%=replace(replace(eMailCC,"<","&lt;"),">","&gt;")%>&nbsp;&nbsp; Status:&nbsp;(<%=TStatus%>):&nbsp;
			<%=PTF_SBK&objRS("PTF_WSP")&objRS("PTF_TS")&objRS("PTF_OD")&objRS("PTF_BT")&objRS("PTF_JT")%>
			<%=objRS("PTF_CS")&objRS("PTF_CJ")&objRS("PTF_SD")&objRS("PTF_TU")&objRS("PTF_HD")&objRS("PTF_TNY")%></p><%		

		objRS.MoveNext

	LOOP
		
END IF

objRS.close
set objRS = Nothing
set objMessage = Nothing




%>
		<p><%=NumTours%> Tournaments with one or more items missing.</p>
		<p>E-Mail Followups could be sent for <%=NumEmails%> Tournaments.</p>

		</td>
		<td>&nbsp;&nbsp;&nbsp;</td>
	</tr>

	<tr>
		<td>&nbsp;&nbsp;&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;&nbsp;&nbsp;</td>
	</tr>

	</table>
<%


WriteIndexPageFooter

%>

