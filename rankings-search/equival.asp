<% Option Explicit %>
<% Response.Buffer = True %>

<!--#include virtual="/rankings/settingsHQ.asp"-->


<%


Server.ScriptTimeout = 1500 


'	Declare variables used in Weekly Followup processes.

Dim objRS, nPTFeMails, sSQL, vbCrLf
Dim TName, TDateE, TournAppID, TSanction, TStatus, WksOld
Dim PTF_SBK, PTF_WSP, PTF_TS, PTF_OD, PTF_BT, PTF_JT
Dim PTF_CS, PTF_CJ, PTF_SD, PTF_TU, PTF_HD, PTF_TNY
Dim eMailTo, eMailFrom, eMailCC, eMailBCC, eMailBody, eMailReplyTo, strMissing, SeedRep
nPTFeMails = 0: vbCrLf = Chr(13) & Chr(10): 

'	=====================================================================================
'	NOTE -- All HTML code, headers and etc, ONLY occurs when run manually, where there IS
'	a value for Request("Equival") -- otherwise is executed as automatic nightly run,
'	in which case at end will do a Response.Redirect to Qualifications Recalculation, 
'	after sending the recap eMail report.
'	=====================================================================================

IF Request("Equival") <> "" THEN

	IF trim(Request("skiyear")) <> "" THEN
		IF trim(request("skiyear")) > 1 and trim(request("skiyear")) <= 13 THEN
			Response.Redirect "/rankings/equival-2009.asp?Equival=" & Request("Equival") & "&skiyear=" & request("skiyear")
		ELSEIF trim(request("skiyear")) > 1 and trim(request("skiyear")) <= 16 THEN
			Response.Redirect "/rankings/equival-2012.asp?Equival=" & Request("Equival") & "&skiyear=" & request("skiyear")
		END IF
	END IF

	%>
	<html><head><title>Member Update / Rankings Recalculations</title></head>
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
	</td></tr></table><br><FONT FACE="Arial,Vendana,Helvetica" SIZE=2>
	
	<%
	
'	======================================================================
'	If NO value present for Request("Equival"), then IF today is Wednesday,
'	then we'll generate the once-a-week Post-Tournament Follow-up emails.
'	======================================================================
	
ELSEIF DatePart("w",Date()) = 4 THEN

	Set objRS = Server.CreateObject("ADODB.recordset")

	SetupEmailService


	'	-- First we generate the select query, to pull tournaments for followup. --
	' ---------------------------------------------------------------------------

	'	Please note the DateDiff(wk,dateadd(d,-1,TDateE),GetDate()) logic --
	'	what this accomplishes is to shift the "week boundary" to Sun/Mon,
	'	where DateDiff normally uses Sat/Sun as the boundary.  Hence all events
	'	with end dates thru Sunday will all be in the "Same Week".

	sSQL = "Select ST.TournAppID, ST.TSanction, ST.TName, ST.TDateE,"
	sSQL = sSQL & " DateDiff(wk,DateAdd(d,-1,ST.TDateE),GetDate()) as WksOld,"
	sSQL = sSQL & " ST.TStatus, ST.TDirName, ST.TDirEMail, Coalesce(PT.PTF_SBK,-1)"
	sSQL = sSQL & " AS PTF_SBK, PT.PTF_WSP, PT.PTF_TS, PT.PTF_OD, PT.PTF_BT, PT.PTF_JT,"
	sSQL = sSQL & " PT.PTF_CS, PT.PTF_CJ, PT.PTF_SD, PT.PTF_TU, PT.PTF_HD, PT.PTF_TNY,"

	sSQL = sSQL & " CJ.CJudgName, CJ.CJudgEMail, CC.CScorName, CC.CScorEMail,"
	sSQL = sSQL & " ST.TStatus, TSanType FROM " & SanctionTableName & " ST LEFT JOIN "

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
	sSQL = sSQL & " AND substring(ST.TSanction,7,1) in ('C','E','L','R','A','B','P')"
	sSQL = sSQL & " AND (DateDiff(wk,DateAdd(d,-1,ST.TDateE),GetDate()) in (2,3,4,6,8,12,16,20)"
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

		TournAppID = objRS("TournAppID")
		TSanction = objRS("TSanction")
		TName = objRS("TName")
		TDateE = Replace(FormatDateTime(objRS("TDateE"),2),"/","-")
		IF Mid(TDateE,2,1) = "-" THEN TDateE = "0" & TDateE
		IF Mid(TDateE,5,1) = "-" THEN TDateE = Left(TDateE,3) & "0" & Right(TDateE,6)
		TStatus = objRS("TStatus")
		WksOld = objRS("WksOld")
		PTF_SBK = objRS("PTF_SBK")

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

		IF len(eMailTo) > 0 THEN
		

			'	-- Next we establish from and secondary addressing based on jurisdiction codes --
      ' ---------------------------------------------------------------------------------
	
			IF mid(TSanction,3,1) = "C" THEN
				eMailCC = """Gordon Hall"" <gordon@lsfdev.com>"
				eMailFrom = """Danny LeBourgeois"" <competition@usawaterski.org>": eMailBCC = "dleboo@gmail.com"
				SeedRep = "Danny LeBourgeois" & vbCrLf & "AWSA South Central Seeding" & vbCrLf & "dleboo@gmail.com" & vbCrLf & "(713) 213-1779"
				eMailReplyTo = "dleboo@gmail.com"
			ELSEIF mid(TSanction,3,1) = "M" THEN
				' eMailCC = """Dean Chappell"" <dean.chappell@comcast.net>"
				' eMailFrom = """Dave Clark"" <awsatechdude@comcast.net>": eMailBCC = eMailFrom
				' SeedRep = "Dave Clark" & vbCrLf & "AWSA Midwest Seeding" & vbCrLf & "awsatechdude@comcast.net" & vbCrLf & "(847) 269-7041"
				eMailCC = """Dean Chappell"" <dean.chappell@comcast.net>"
				eMailFrom = """Dave Clark"" <competition@usawaterski.org>": eMailBCC = "h2oskimo@gmail.com"
				SeedRep = "Mike O'Connor" & vbCrLf & "AWSA Midwest Seeding" & vbCrLf & "h2oskimo@gmail.com" & vbCrLf & "(573) 864-2138"
				eMailReplyTo = "h2oskimo@gmail.com"
			ELSEIF mid(TSanction,3,1) = "E" THEN
				eMailCC = """James Powell"" <easternevp@gmail.com>"
				eMailFrom = """Jennifer Frederick-Kelly"" <competition@usawaterski.org>": eMailBCC = "jennifer@frederickmachine.com"
				SeedRep = "Jennifer Frederick-Kelly" & vbCrLf & "AWSA East Seeding" & vbCrLf & "jennifer@frederickmachine.com" & vbCrLf & "(716) 892-1425"
				eMailReplyTo = "jennifer@frederickmachine.com"
			ELSEIF mid(TSanction,3,1) = "S" THEN
				eMailCC = """Bob Archambeau"" <evp@awsasouth.org>"
				eMailFrom = """Kirby Whetsel"" <competition@usawaterski.org>": eMailBCC = "kwhetsel@charter.net"
				SeedRep = "Kirby Whetsel" & vbCrLf & "AWSA South Seeding" & vbCrLf & "kwhetsel@charter.net" & vbCrLf & "(931) 409-0389"
				eMailReplyTo = "kwhetsel@charter.net"
			ELSEIF mid(TSanction,3,1) = "W" THEN
				eMailCC = """Owen Letcher"" <owen.letcher@comcast.net>"
				eMailFrom = """Judy Stanford"" <competition@usawaterski.org>": eMailBCC = "judy-don@sbcglobal.net"
				SeedRep = "Judy Stanford" & vbCrLf & "AWSA West Seeding" & vbCrLf & "judy-don@sbcglobal.net" & vbCrLf & "(925) 932-7781"
				eMailReplyTo = "judy-don@sbcglobal.net"
			ELSEIF mid(TSanction,3,1) = "U" THEN
				eMailCC = """Jeff Surdej"" <j_surdej@yahoo.com>"
				eMailFrom = """Robert Rhyne"" <competition@usawaterski.org>": eMailBCC = "rrriii@mindspring.com"
				SeedRep = "Robert Rhyne" & vbCrLf & "NCWSA Seeding" & vbCrLf & "rrriii@mindspring.com" & vbCrLf & "(704) 906-7779"
				eMailReplyTo = "rrriii@mindspring.com"
			ELSE
				eMailCC = """Kirby Whetsel"" <kwhetsel@charter.net>"
				eMailFrom = """USA Water Ski Competition"" <shardee@usawaterski.org>"
				SeedRep = "Sandy Hardee" & vbCRLF & "Sanctioning & Competition" & vbCRLF & "shardee@usawaterski.org" & vbCRLF & "1-863-874-5681"
				eMailReplyTo = "shardee@usawaterski.org"
			END IF


'response.write("<br>Line 237: ")
'response.end

			'	-- Next we compile a list of the missing items for the eMail body --
			' --------------------------------------------------------------------
			
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
					strMissing = strMissing & "( " & TournAppID & "CJ.PRN )"
				END IF

				IF objRS("PTF_SD") = 0 THEN 
					strMissing = strMissing & vbCrLf & "   Safety Directors Report         "
					strMissing = strMissing & "( " & TournAppID & "SD.PRN )"
				END IF

				IF objRS("PTF_TU") = 0 THEN 
					strMissing = strMissing & vbCrLf & "   Towboat Utilization Report      "
					strMissing = strMissing & "( " & TournAppID & "TU.PRN )"
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

			
			
			'	-- Now we construct the actual body of the eMail note, depending on vintage and status. --
			' ------------------------------------------------------------------------------------------
			
			objMessage.To = eMailTo

			IF instr(eMailCC,eMailFrom) = 0 then eMailCC = eMailCC & "; " & eMailFrom
			objMessage.CC = eMailCC
			' objMessage.From = """Competition Support"" <Post_Tour@usawaterski.org>"
			objMessage.From = """Competition Support"" <competition@usawaterski.org>"

			' IF instr(eMailBCC,"Dave Clark") = 0 THEN eMailBCC = eMailBCC & "; ""Dave Clark"" <awsatechdude@comcast.net>"
			IF instr(eMailBCC,"Kirby Whetsel") = 0 THEN eMailBCC = eMailBCC & "; ""Kirby Whetsel"" <kwhetsel@charter.net>"
			' eMailBCC = eMailBCC & "; ""Steve Hansen"" <shansen@dakotatechgroup.com>"
			objMessage.BCC = eMailBCC	
			objMessage.ReplyTo = eMailReplyTo					
			objMessage.Subject = "Post-Tournament Reports from " & TSanction & " " & TName & " (" & TDateE & ")"

			eMailBody = "Dear Tournament Organizer and/or Chief Official(s) --" & vbCrLf & vbCrLf

			IF WksOld = 1 AND TStatus = 2 THEN

				eMailBody = eMailBody & "The post-tournament reports from " & TSanction & " " & TName & vbCrLf
				eMailBody = eMailBody & TDateE & " have not yet appeared in the Sanction control system." & vbCrLf & vbCrLf

				eMailBody = eMailBody & "USA Water Ski's real-time ranking database has matured into a" & vbCrLf
				eMailBody = eMailBody & "qualifications and ranking platform that now serves many time-" & vbCrLf
				eMailBody = eMailBody & "critical purposes.  We hope you can help us meet those objectives," & vbCrLf
				eMailBody = eMailBody & "by submitting the WSTIMS Zip file for this event at your earliest" & vbCrLf
				eMailBody = eMailBody & "convenience." & vbCrLf & vbCrLf

				eMailBody = eMailBody & "The Zip file should be emailed to me at the email address shown" & vbCrLf 
				eMailBody = eMailBody & "below.  If you are having difficulty producing any of those reports," & vbCrLf 
				eMailBody = eMailBody & "or the Zip file itself, please contact me for assistance." & vbCrLf & vbCrLf

				eMailBody = eMailBody & "If the above-referenced competition did not actually take place as" & vbCrLf
				eMailBody = eMailBody & "planned, then please advise USA Water Ski HQ so they can revise the" & vbCrLf
				eMailBody = eMailBody & "status of the tournament in the Sanction Control System." & vbCrLf & vbCrLf

			ELSEIF TStatus = 2 THEN

				eMailBody = eMailBody & "The post-tournament reports from " & TSanction & " " & TName & vbCrLf
				eMailBody = eMailBody & TDateE & " have not yet been posted in the Sanction control system." & vbCrLf & vbCrLf

				eMailBody = eMailBody & "It has now been " & WksOld & " weeks since the above-referenced event was" & vbCrLf 
				eMailBody = eMailBody & "scheduled to take place, according to our records.  If the" & vbCrLf 
				eMailBody = eMailBody & "competition did NOT actually take place as planned, then please" & vbCrLf 
				eMailBody = eMailBody & "advise USA Water Ski HQ so they can revise the status of your" & vbCrLf 
				eMailBody = eMailBody & "tournament in the Sanction Control System." & vbCrLf & vbCrLf

				eMailBody = eMailBody & "However, if the competition DID run as planned, then you need to" & vbCrLf
				eMailBody = eMailBody & "get the post-tournament reports in soon.  The WSTIMS Zip file" & vbCrLf
				eMailBody = eMailBody & "should be emailed to me at the email address shown below.  If you" & vbCrLf
				eMailBody = eMailBody & "are having difficulty producing any of those reports, or the Zip" & vbCrLf
				eMailBody = eMailBody & "file itself, please contact me for assistance." & vbCrLf & vbCrLf

				IF WksOld > 8 THEN
					eMailBody = eMailBody & "You need to be aware that any subsequent sanction applications for" & vbCrLf 
					eMailBody = eMailBody & "your organization cannot be approved, until these missing reports" & vbCrLf
					eMailBody = eMailBody & "have been received and checked off." & vbCrLf & vbCrLf
				END IF

			ELSE

				eMailBody = eMailBody & "The post-tournament reports from " & TSanction & " " & TName & vbCrLf
				eMailBody = eMailBody & TDateE & " are not completely posted in the Sanction control system." & vbCrLf & vbCrLf

				eMailBody = eMailBody & "While some of the required reports have been received and checked" & vbCrLf
				eMailBody = eMailBody & "off, the following items are still outstanding as of today:" & vbCrLf

				eMailBody = eMailBody & strMissing & vbCrLf & vbCrLf

				eMailBody = eMailBody & "If these missing items have been mailed in the past few days and" & vbCrLf
				eMailBody = eMailBody & "this eMail is crossing your package in the mail, please disregard" & vbCrLf 
				eMailBody = eMailBody & "this notice.  Otherwise, please note that it has now been " & WksOld & " weeks" & vbCrLf
				eMailBody = eMailBody & "since the above-referenced event took place, and these missing items" & vbCrLf 
				eMailBody = eMailBody & "need to be submitted before we can mark your event as complete." & vbCrLf & vbCrLf

				IF WksOld > 8 THEN
					eMailBody = eMailBody & "You need to be aware that any subsequent sanction applications for" & vbCrLf 
					eMailBody = eMailBody & "your organization cannot be approved, until these missing reports" & vbCrLf
					eMailBody = eMailBody & "have been received and checked off." & vbCrLf & vbCrLf
				END IF

				eMailBody = eMailBody & "Emailing these items to me in electronic form is preferred.  However," & vbCrLf 
				eMailBody = eMailBody & "if you have paper documents instead, then you should send those to" & vbCrLf 
				eMailBody = eMailBody & "Matthew Stone at USA Waterski HQ, by postal mail.  If you are having" & vbCrLf 
				eMailBody = eMailBody & "difficulty producing any of those reports, or the WSTIMS Zip file" & vbCrLf 
				eMailBody = eMailBody & "itself, please contact me for assistance." & vbCrLf & vbCrLf
	
			END IF

			eMailBody = eMailBody & "Thank you for your hard work and continued support," & vbCrLf & vbCrLf

			eMailBody = eMailBody & SeedRep

			'	Now finally we send the constructed eMail message
	
			objMessage.TextBody = eMailBody

			objMessage.Send
			nPTFeMails = nPTFeMails + 1

		END IF
		
		objRS.MoveNext

	LOOP
		
	WriteLog(date() & "  " & time() & "  Weekly Post-Tournament Follow-ups Concluded, " & nPTFeMails & " eMail Notices Sent.")

	END IF

	objRS.close
	set objRS = nothing
	set objMessage = nothing

END IF 






' --------------------------------
   SUB ShowProgress(nPctComplete)
' --------------------------------

Response.Write "<SCR" & "IPT LANGUAGE=""JavaScript"">" & vbCrLf
Response.Write "statuspic.width = Math.ceil(" & nPctComplete & " * progBarWidth);" & vbCrLf
Response.Write "</SCR" & "IPT>"
Response.Flush

END SUB


' --------------------------------
   SUB FinishProgress
' --------------------------------

Response.Write "<SCR" & "IPT LANGUAGE=""JavaScript"">" & vbCrLf
Response.Write "ProgBar.style.visibility ='hidden';" & vbCrLf
Response.Write "</SCR" & "IPT>"
Response.Flush

END SUB








' -----------------------------------------------------------------------------------------------------
' -------------------------------   START OF MAIN PROGRAM ---------------------------------------------
' -----------------------------------------------------------------------------------------------------

Dim i, j
Dim tBeginDate, DupMemList
Dim nProcessedSoFar, nTotalMembers, tempCounter, tempvar, TempSum, TempLen, TempPtr
Dim strHTML, sSkiYearBegin, sSkiYearEnd, sProcessingYear, sPrevYear, sSkiYearName

Dim tUpAge, tBirthdate, tLatestBirthYear, tSkiAge, AgedOut
Dim R_Ski, R_PLC, N_PLC

Dim timeTHEN, timeNow
Dim EMailToWho

' Overall and Event Ranking Stuff
Dim TempMemberID, TempTeam, TempTeamStat, TempEvent, TempTourID, TempDiv, TempScore, TempOASc, TempAdd
Dim TempOverEvts, TempOverEvtsReq, TempOATot, TempDivOrig, InDivOrig, TempDivType
Dim Slalom1, Slalom2, Slalom3, Trick1, Trick2, Trick3
Dim Jump1, Jump2, Jump3, Class1, Class2, Class3
Dim S_Round1, J_Round1, T_Round1, S_Score1, J_Score1, T_Score1
Dim S_Round2, J_Round2, T_Round2, S_Score2, J_Score2, T_Score2
Dim S_Round3, J_Round3, T_Round3, S_Score3, J_Score3, T_Score3
Dim nScoC, nScoR, TotScore, RSco1, RSco2, RSco3, FmtSco
Dim TempSlmNR, TempSlmOP, TempSlmExp, TempTrkNR, TempTrkExp, TempJmpNR, TempJmpExp, OASc, FmtOASc
Dim RExp1, RExp2, RExp3, RPen1, RPen2, RPen3, RMaxRat

' Ranking Penalty Table as Function of C vs ELR score count
Dim tPenalty(3,3)
FOR nScoC=0 to 3: FOR nScoR=0 to 3: tPenalty(nScoC,nScoR)=0: NEXT: NEXT
tPenalty(0,1)=5: tPenalty(1,0)=10: tPenalty(1,1)=2.5: tPenalty(2,0)=5

' Operational Controls
Dim RunOverride, RunEquivScore, RunOvrllScore, RunOvrllRanks, RunEventRanks, RunLevelLogic, ReCalc12

Dim tSYEndDate, tBirthYear

' Membership Update/Merge Controls and variables
Dim TempHQPID, TempLclPID, PIDwCheckDigit, HQConnect, HQrs
Dim nHQExt, nLocal, nInserts, nUpdates, nDeletes, LastHQPID
Dim nConsMems, nConsHits, nScoUpdts, nOvrUpdts, nRnkUpdts
Dim nOffUpdts, nColUpdts, nRegUpdts, nEvtUpdts, nNewElites
Dim nSlices, iSlice, nTotal, nSoFar, StartTime

StartTime = Time(): DupMemList = "": ReCalc12 = "N"
nHQExt = 0: nLocal = 0: nInserts = 0: nUpdates = 0: nDeletes = 0
LastHQPID = 0: nConsMems = 0: nConsHits = 0
nRnkUpdts = 0: nOffUpdts = 0: nColUpdts = 0: nRegUpdts = 0: nNewElites = 0




' ---------------------------------------------------------------------------------
' ---------------------------- TIME KEEPING FUNCTIONS -----------------------------
' ---------------------------------------------------------------------------------

timeTHEN = milliDif()
WriteLog(date() &"  "& time() &"  Begin Nightly Member Update / Ranking Recalculation Process.")

OpenCon
Set rs = Server.CreateObject("ADODB.recordset")

' Response.write("Equival value = " & Request("Equival") & "<br>")

EMailToWho = "<awsatechdude@comcast.net>; <cronemarka@gmail.com>; <shardee@usawaterski.org>"


IF Request("Equival") <> "ReCalc" THEN




' --------------------------------------------------------------------------------------
' ************** Beginning of Member Extract Update Conditional Section ***************
' --------------------------------------------------------------------------------------


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

' Begin by pulling a consolidated extract from HQ server, using the
' new MembersLive view.  Order by Person ID for subsequent merge.

IF Request("Equival") <> "" THEN 
	Response.write("Member Update: Slice " & iSlice & " -- Querying HQ Server ...")
	Response.Flush
END IF

sSQL = "SELECT * FROM " & MemberLiveTableName
sSQL = sSQL & " WHERE EffectiveTo >= '2004-08-15'"
sSQL = sSQL & " AND PersonID % " & nSlices & " = " & iSlice:  ' New Daily Slice Update
sSQL = sSQL & " ORDER BY PersonID"

' WriteDebugSQL(sSQL)

Set HQrs = Server.CreateObject("ADODB.recordset")
HQrs.open sSQL, sConnectionToTRATable, 3, 3

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








' -- Now we pull some key columns from the local table, also keyed by Person ID, so that we can merge with HQ Extract. --
' -----------------------------------------------------------------------------------------------------------------------

IF Request("Equival") <> "" THEN 
	Response.write(" Querying Local Table ...")
	Response.Flush
END IF

sSQL = "SELECT PersonID, DateUpdated, EffectiveTo"
sSQL = sSQL & " FROM " & MemberTableName
sSQL = sSQL & " WHERE (PersonID % " & nSlices & " = " & iSlice & ")":  ' New Daily Slice Update
sSQL = sSQL & " ORDER BY PersonID"

' WriteDebugSQL(sSQL)

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


' -- Now we begin a merge loop between the HQ extract and the local table, for this slice.
' -------------------------------------------------------------------------------------------

IF Request("Equival") <> "" THEN 
	Response.write(" Merging/Updating ...")
	Response.Flush
END IF

DO UNTIL HQrs.eof AND rs.eof

IF Request("Equival") <> "" THEN 
	IF (nSoFar mod 17 = 9) THEN ShowProgress (nSoFar / nTotal)
END IF

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
		sSQL = sSQL & SQLClean(HQrs("Deceased")) &  "','"
		sSQL = sSQL & SQLClean(HQrs("WaiverStatusID")) & "','"
		sSQL = sSQL & SQLClean(HQrs("WaiverGoodTo")) & "')"

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
	


	' -- Now finally advance the incoming HQ Extract recordset.
  ' -----------------------------------------------------------
    
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
	sSQL = sSQL & " Deceased = '" & SQLClean(HQrs("Deceased")) & "',"
	sSQL = sSQL & " WaiverStatusID = '" & SQLClean(HQrs("WaiverStatusID")) & "',"
	sSQL = sSQL & " WaiverGoodTo = '" & SQLClean(HQrs("WaiverGoodTo")) & "'"
	sSQL = sSQL & " WHERE PersonID = " & TempLclPID
		
	' WriteDebugSQL(sSQL)


	' -- Invoke the update and then tally number of Member rows updated --
	' --------------------------------------------------------------------

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

IF Request("Equival") <> "" THEN 
	Response.write(" DONE at " & Time() & "<br>")
	Response.Flush
END IF

NEXT

' End of Member Update Loop over Slices.  Now spit out an update recap report.

IF Request("Equival") <> "" THEN 
	Response.write("&nbsp;<br>" & Formatnumber(nHQExt,0) & " Member rows supplied from HQ Server<br>")
	Response.write(Formatnumber(nLocal,0) & " Member rows found in Local Server Table<br>")
	Response.write(Formatnumber(nUpdates,0) & " Member rows updated with new Data<br>")
	Response.write(Formatnumber(nInserts,0) & " New Member rows added<br>")
	Response.write(Formatnumber(nDeletes,0) & " Old Member rows deleted<br>")
	Response.write(Formatnumber(nLocal+nInserts-nDeletes,0) & " Member rows now in Local Server Table<br>&nbsp;<br>")
	IF Len(DupMemList) > 0 THEN Response.write("Duplicate Person ID's encountered: " & DupMemList & "<br>&nbsp;<br>")
	Response.Flush
END IF

WriteLog(date() &"  "& time() &"  Membership Extract & Update Completed Successfully.")


' ********* This next section pulls the Membership Consolidation "Was-to-Is"
' ********* cross-reference table, and Updates Member IDs in the various
' ********* Rankings Database table -- Raw Scores, Overall Scores, Rankings.

' ********* First step is to empty the local ConsolidatedMembers table, then
' ********* issue Query to HQ Database to return that table as a RecordSet.
' ********* Note that this Query is structured to eliminate any possible
' ********* duplicate "from" PersonID references.

IF Request("Equival") <> "" THEN 
	Response.write("Member Consolidation -- Querying HQ Server ...")
	Response.Flush
END IF



' -- Delete ?? 

sSQL = "DELETE FROM " & ConsMemTableName
Con.Execute(sSQL)

' sSQL = "SELECT PersonIDDeleted as OldMemID, PersonIDConsolidatedTo as NewMemID"
' sSQL = sSQL & " FROM waterski.dbo.[Consolidated Members]"

sSQL = "Select OldMemID, cast(substring(max(MaxDate),9,7) as integer) as NewMemID"
sSQL = sSQL & " FROM (Select cm.PersonIDDeleted as OldMemID, mhx.MaxDate"
sSQL = sSQL & " FROM waterski.dbo.[Consolidated Members] cm JOIN (Select"
sSQL = sSQL & " [Person ID] as PersonID, convert(char(8),max(EffectiveTo),112)"
sSQL = sSQL & " + right(convert(char(8),10000000+[Person ID]),7) as MaxDate"
sSQL = sSQL & " FROM waterski.dbo.[Membership History]"
sSQL = sSQL & " Where [Person ID] in (Select distinct PersonIDConsolidatedTo"
sSQL = sSQL & " FROM waterski.dbo.[Consolidated Members]) Group by [Person ID]"
sSQL = sSQL & " ) mhx on mhx.PersonID=cm.PersonIDConsolidatedTo) cmx"
sSQL = sSQL & " group by OldMemID order by OldMemID;"

' WriteDebugSQL(sSQL)

Set HQrs = HQConnect.Execute(sSql)
tempvar = HQrs.getrows()
nTotal = ubound(tempvar,2)
nSoFar = 0

HQrs.MoveFirst



' -- Loop over Consolidated Membership ID rows returned from HQ Table. --
' -----------------------------------------------------------------------

DO UNTIL HQrs.eof

	nSoFar = nSoFar + 1

	IF Request("Equival") <> "" THEN 
		IF (nSoFar mod 5 = 2) THEN ShowProgress (nSoFar / nTotal)
	END IF

	' Add this entry to the local ConsolidatedMembers table

	sSQL = "INSERT INTO " & ConsMemTableName & " VALUES ("
	sSQL = sSQL & HQrs("OldMemID") & "," & HQrs("NewMemID") & ",'" 
	sSQL = sSQL & PersonIDwChkDgt(HQrs("OldMemID")) & "','"
	sSQL = sSQL & PersonIDwChkDgt(HQrs("NewMemID")) & "')"
	Con.Execute(sSQL)
	nConsMems = nConsMems + 1

	HQrs.moveNEXT

LOOP

HQrs.Close




' -- Next step is to purge any consolidation records where the "To" (new)
' -- Member ID is NOT present in our newly-updated Membership Table.
' ---------------------------------------------------------------------------

sSQL = "DELETE FROM " & ConsMemTableName & " where ToPersonID not in"
sSQL = sSQL & " (select PersonID from " & MemberTableName & ")"
Con.Execute(sSQL)

' Now count the net remaining consolidation records

sSQL = "Select count(*) as Kount from " & ConsMemTableName
rs.open sSQL, sConnectionToTRATable, 3, 3
rs.MoveFirst
nConsHits = rs("Kount")
rs.close



' ---------------------------------------------------------------------------------------------------------------
' -- Now apply the net valid consolidations to various tables in Scores/Rankings, Officials, and Registration
' ---------------------------------------------------------------------------------------------------------------

IF Request("Equival") <> "" THEN 
	Response.write(" Translating Consolidated Member IDs ...")
	Response.Flush
END IF

' First step updates Raw Score table records
		
sSQL = "Select count(*) as Kount from " & RawScoresTableName & " Where"
sSQL = sSQL & " MemberID in (select FromMemberID from " & ConsMemTableName & ")"
rs.open sSQL, sConnectionToTRATable, 3, 3
rs.MoveFirst
IF rs("Kount") > 0 THEN
	nRnkUpdts = nRnkUpdts + rs("Kount")
	sSQL = "UPDATE ST Set MemberID = CM.ToMemberID FROM "
	sSQL = sSQL & RawScoresTableName & " AS ST, " & ConsMemTableName
	sSQL = sSQL & " AS CM WHERE ST.MemberID = CM.FromMemberID"	
	Con.Execute(sSQL)
END IF
rs.close
	
' Next update any Overall Score table records

sSQL = "Select count(*) as Kount from " & OverallScoresTableName & " Where"
sSQL = sSQL & " MemberID in (select FromMemberID from " & ConsMemTableName & ")" 
rs.open sSQL, sConnectionToTRATable, 3, 3
rs.MoveFirst
IF rs("Kount") > 0 THEN
	nRnkUpdts = nRnkUpdts + rs("Kount")
	sSQL = "UPDATE OT Set MemberID = CM.ToMemberID FROM "
	sSQL = sSQL & OverallScoresTableName & " AS OT, " & ConsMemTableName
	sSQL = sSQL & " AS CM WHERE OT.MemberID = CM.FromMemberID"	
	Con.Execute(sSQL)
END IF
rs.close
	
' Next update any Ranking table records

sSQL = "Select count(*) as Kount from " & RankTableName & " Where"
sSQL = sSQL & " MemberID in (select FromMemberID from " & ConsMemTableName & ")"
rs.open sSQL, sConnectionToTRATable, 3, 3
rs.MoveFirst
IF rs("Kount") > 0 THEN
	nRnkUpdts = nRnkUpdts + rs("Kount")
	sSQL = "UPDATE RT Set MemberID = CM.ToMemberID FROM "
	sSQL = sSQL & RankTableName & " AS RT, " & ConsMemTableName
	sSQL = sSQL & " AS CM WHERE RT.MemberID = CM.FromMemberID"	
	Con.Execute(sSQL)
END IF
rs.close
		
' Next update any Officials table records

sSQL = "Select count(*) as Kount from USAWaterski.dbo.Officials Where"
sSQL = sSQL & " PersonID in (select FromPersonID from " & ConsMemTableName & ")"
rs.open sSQL, sConnectionToTRATable, 3, 3
rs.MoveFirst
IF rs("Kount") > 0 THEN
	nOffUpdts = nOffUpdts + rs("Kount")
	sSQL = "UPDATE OT Set PersonID = CM.ToPersonID FROM "
	sSQL = sSQL & " USAWaterski.dbo.Officials AS OT, " & ConsMemTableName
	sSQL = sSQL & " AS CM WHERE OT.PersonID = CM.FromPersonID"	
	Con.Execute(sSQL)
END IF
rs.close

' Next update any Collegiate Roster table records

sSQL = "Select count(*) as Kount from " & TeamRosterTableName & " Where"
sSQL = sSQL & " MemberID in (select FromMemberID from " & ConsMemTableName & ")"
rs.open sSQL, sConnectionToTRATable, 3, 3
rs.MoveFirst
IF rs("Kount") > 0 THEN
	nColUpdts = nColUpdts + rs("Kount")
	sSQL = "UPDATE RT Set MemberID = CM.ToMemberID FROM "
	sSQL = sSQL & TeamRosterTableName & " AS RT, " & ConsMemTableName
	sSQL = sSQL & " AS CM WHERE RT.MemberID = CM.FromMemberID"	
	Con.Execute(sSQL)
END IF
rs.close
		
' Next update any Collegiate Rotations table records

sSQL = "Select count(*) as Kount from " & TeamRotationsTableName & " Where"
sSQL = sSQL & " MemberID in (select FromMemberID from " & ConsMemTableName & ")"
rs.open sSQL, sConnectionToTRATable, 3, 3
rs.MoveFirst
IF rs("Kount") > 0 THEN
	nColUpdts = nColUpdts + rs("Kount")
	sSQL = "UPDATE RT Set MemberID = CM.ToMemberID FROM "
	sSQL = sSQL & TeamRotationsTableName & " AS RT, " & ConsMemTableName
	sSQL = sSQL & " AS CM WHERE RT.MemberID = CM.FromMemberID"	
	Con.Execute(sSQL)
END IF
rs.close
		
' Next we update any Registration Gen table records

sSQL = "Select count(*) as Kount from " & RegGenTableName & " Where"
sSQL = sSQL & " MemberID in (select FromMemberID from " & ConsMemTableName & ")"
rs.open sSQL, sConnectionToTRATable, 3, 3
rs.MoveFirst
IF rs("Kount") > 0 THEN
	nRegUpdts = nRegUpdts + rs("Kount")
	sSQL = "UPDATE RT Set MemberID = CM.ToMemberID FROM "
	sSQL = sSQL & RegGenTableName & " AS RT, " & ConsMemTableName
	sSQL = sSQL & " AS CM WHERE RT.MemberID = CM.FromMemberID"	
	Con.Execute(sSQL)
END IF
rs.close

' Next we update any Registration Temp table records

sSQL = "Select count(*) as Kount from " & RegTempTableName & " Where"
sSQL = sSQL & " MemberID in (select FromMemberID from " & ConsMemTableName & ")"
rs.open sSQL, sConnectionToTRATable, 3, 3
rs.MoveFirst
IF rs("Kount") > 0 THEN
	nRegUpdts = nRegUpdts + rs("Kount")
	sSQL = "UPDATE RT Set MemberID = CM.ToMemberID FROM "
	sSQL = sSQL & RegTempTableName & " AS RT, " & ConsMemTableName
	sSQL = sSQL & " AS CM WHERE RT.MemberID = CM.FromMemberID"	
	Con.Execute(sSQL)
END IF
rs.close

' Next we update any Registration Detail table records

sSQL = "Select count(*) as Kount from " & RegDetailTableName & " Where"
sSQL = sSQL & " MemberID in (select FromMemberID from " & ConsMemTableName & ")"
rs.open sSQL, sConnectionToTRATable, 3, 3
rs.MoveFirst
IF rs("Kount") > 0 THEN
	nRegUpdts = nRegUpdts + rs("Kount")
	sSQL = "UPDATE RT Set MemberID = CM.ToMemberID FROM "
	sSQL = sSQL & RegDetailTableName & " AS RT, " & ConsMemTableName
	sSQL = sSQL & " AS CM WHERE RT.MemberID = CM.FromMemberID"	
	Con.Execute(sSQL)
END IF
rs.close

' Next we update any Registration Detail Temp table records

sSQL = "Select count(*) as Kount from " & RegDetailTempTableName & " Where"
sSQL = sSQL & " MemberID in (select FromMemberID from " & ConsMemTableName & ")"
rs.open sSQL, sConnectionToTRATable, 3, 3
rs.MoveFirst
IF rs("Kount") > 0 THEN
	nRegUpdts = nRegUpdts + rs("Kount")
	sSQL = "UPDATE RT Set MemberID = CM.ToMemberID FROM "
	sSQL = sSQL & RegDetailTempTableName & " AS RT, " & ConsMemTableName
	sSQL = sSQL & " AS CM WHERE RT.MemberID = CM.FromMemberID"	
	Con.Execute(sSQL)
END IF
rs.close

' Next we update any Registration Transactions table records

sSQL = "Select count(*) as Kount from " & RegTransTableName & " Where"
sSQL = sSQL & " MemberID in (select FromMemberID from " & ConsMemTableName & ")"
rs.open sSQL, sConnectionToTRATable, 3, 3
rs.MoveFirst
IF rs("Kount") > 0 THEN
	nRegUpdts = nRegUpdts + rs("Kount")
	sSQL = "UPDATE RT Set MemberID = CM.ToMemberID FROM "
	sSQL = sSQL & RegTransTableName & " AS RT, " & ConsMemTableName
	sSQL = sSQL & " AS CM WHERE RT.MemberID = CM.FromMemberID"	
	Con.Execute(sSQL)
END IF
rs.close

' Next we update any Registration Payments table records

sSQL = "Select count(*) as Kount from " & RegPaymentTableName & " Where"
sSQL = sSQL & " MemberID in (select FromMemberID from " & ConsMemTableName & ")"
rs.open sSQL, sConnectionToTRATable, 3, 3
rs.MoveFirst
IF rs("Kount") > 0 THEN
	nRegUpdts = nRegUpdts + rs("Kount")
	sSQL = "UPDATE RT Set MemberID = CM.ToMemberID FROM "
	sSQL = sSQL & RegPaymentTableName & " AS RT, " & ConsMemTableName
	sSQL = sSQL & " AS CM WHERE RT.MemberID = CM.FromMemberID"	
	Con.Execute(sSQL)
END IF
rs.close

' Next we update any Registration Qualify table records

sSQL = "Select count(*) as Kount from " & RegQualifyTableName & " Where"
sSQL = sSQL & " MemberID in (select FromMemberID from " & ConsMemTableName & ")"
rs.open sSQL, sConnectionToTRATable, 3, 3
rs.MoveFirst
IF rs("Kount") > 0 THEN
	nRegUpdts = nRegUpdts + rs("Kount")
	sSQL = "UPDATE RT Set MemberID = CM.ToMemberID FROM "
	sSQL = sSQL & RegQualifyTableName & " AS RT, " & ConsMemTableName
	sSQL = sSQL & " AS CM WHERE RT.MemberID = CM.FromMemberID"	
	Con.Execute(sSQL)
END IF
rs.close

' Next we update any Registration Passwords table records

sSQL = "Select count(*) as Kount from " & RegPWTableName & " Where"
sSQL = sSQL & " MemberID in (select FromMemberID from " & ConsMemTableName & ")"
rs.open sSQL, sConnectionToTRATable, 3, 3
rs.MoveFirst
IF rs("Kount") > 0 THEN
	nRegUpdts = nRegUpdts + rs("Kount")
	sSQL = "UPDATE RT Set MemberID = CM.ToMemberID FROM "
	sSQL = sSQL & RegPWTableName & " AS RT, " & ConsMemTableName
	sSQL = sSQL & " AS CM WHERE RT.MemberID = CM.FromMemberID"	
	Con.Execute(sSQL)
END IF
rs.close

' Next we update any Registration Skier Bios table records

sSQL = "Select count(*) as Kount from " & BioTableName & " Where"
sSQL = sSQL & " MemberID in (select FromMemberID from " & ConsMemTableName & ")"
rs.open sSQL, sConnectionToTRATable, 3, 3
rs.MoveFirst
IF rs("Kount") > 0 THEN
	nRegUpdts = nRegUpdts + rs("Kount")
	sSQL = "UPDATE RT Set MemberID = CM.ToMemberID FROM "
	sSQL = sSQL & BioTableName & " AS RT, " & ConsMemTableName
	sSQL = sSQL & " AS CM WHERE RT.MemberID = CM.FromMemberID"	
	Con.Execute(sSQL)
END IF
rs.close

' Next we update any Registration Temporary table records

sSQL = "Select count(*) as Kount from " & RegTemporary & " Where"
sSQL = sSQL & " MemberID in (select FromMemberID from " & ConsMemTableName & ")"
rs.open sSQL, sConnectionToTRATable, 3, 3
rs.MoveFirst
IF rs("Kount") > 0 THEN
	nRegUpdts = nRegUpdts + rs("Kount")
	sSQL = "UPDATE RT Set MemberID = CM.ToMemberID FROM "
	sSQL = sSQL & RegTemporary & " AS RT, " & ConsMemTableName
	sSQL = sSQL & " AS CM WHERE RT.MemberID = CM.FromMemberID"	
	Con.Execute(sSQL)
END IF
rs.close



' --- MARK:  What other tables should be added to this update? --
' ---------------------------------------------------------------


		
IF Request("Equival") <> "" THEN 
	Response.write(" DONE at " & Time() & "<br>")
	Response.write(Formatnumber(nConsMems,0) & " Consolidated Member Rows Read from HQ Table<br>")
	Response.write(Formatnumber(nConsHits,0) & " Consolidated Member Rows hit Updated Member Table<br>")
	Response.write(Formatnumber(nRnkUpdts,0) & " Ranking Tables Entries updated to Cons Mbr IDs<br>")
	Response.write(Formatnumber(nOffUpdts,0) & " Officials Table Entries updated to Cons Mbr IDs<br>")
	Response.write(Formatnumber(nColUpdts,0) & " Collegiate Table Entries updated to Cons Mbr IDs<br>")
	Response.write(Formatnumber(nRegUpdts,0) & " Registration Table Entries updated to Cons Mbr IDs<br>&nbsp;<br>")
	Response.Flush
END IF


' Finally release the HQ record set Object -- all remaining steps are local to ePolk tables.

Set HQrs = Nothing

WriteLog(date() &"  "& time() &"  Membership Consolidation Update Completed Successfully.")



END IF

' ------------------------------------------------------------------------------------------------
' ************** Bottom of Member Extract / Update / Translate Conditional Section ***************
' ------------------------------------------------------------------------------------------------










' ---------------------------------------------------------------------------------------
' ************** Beginning of Rankings Recalculation Conditional Section ***************
' ---------------------------------------------------------------------------------------

IF Request("Equival") <> "MemUpd" THEN



' ========  Begin Rankings Calculation Logic.  First we set some ==========
' ========  flags to control which particular sections are run  ===========

RunOverride="NO"

RunEquivScore = "YES"
RunOvrllScore = "YES"
RunOvrllRanks = "YES"
RunEventRanks = "YES"
RunLevelLogic = "YES"


' ======== Special Score Table maintenance to be run nightly, to clean up specific problems
' ======== that have been identified in the scores table.  First one is 4.5' ramp heights,
' ======== that come in as .255 where they ought to be .215 -- for older Age Divs.

sSQL = "UPDATE " & RawScoresTableName & " SET Perf_Qual1 = .215 WHERE Perf_Qual1 = .255 AND Div IN "
sSQL = sSQL & "('M3','M4','M5','M6','M7','M8','M9','MA','MB','W3','W4','W5','W6','W7','W8','W9','WA','WB')"
Con.Execute(sSQL)


'------------     First step in this process is to UPDATE the date range on the 12 month period      ------------

tBeginDate = FormatDateTime(dateadd("yyyy",-1,now()),2)
sSQL = "SELECT top 1 enddate from " & RawScoresTableName & " WHERE UPPER(right(rtrim(tourid),1)) = 'A' ORDER BY enddate desc"
  rs.open sSQL, SConnectionToTRATable, 3, 3  
   IF rs.EOF THEN
      session("message") = "No National Scores Found! This is a Strange Error -- Please Report To Admin.  Line #681, Equival.asp"
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
' so we look it up in the table.  -- OVERRIDE NOW TO Last 12 MOS INSTEAD -- DJC 2014-09-11
' Override reversed 2015-09-07

    sSQL = "SELECT top 1 * from " & SkiYearTableName & " WHERE DefaultYear = 1"
'    sSQL = "SELECT top 1 * from " & SkiYearTableName & " WHERE SkiYearID = 1"
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



' -- This do loop keeps repeating the recalc until we tell it to stop.
' -- The value in sProcessingYear is the Ski Year ID that we are processing for.
' -- sPrevYear is the Previous Ski Year ID that we use for testing CutOffScores.
' ---------------------------------------------------------------------------------



DO WHILE sProcessingYear <> "STOP"

' First we set the Recalculation Underway flag for the Ski Year being processsed.

sSQL = "UPDATE " & SkiYearTableName & " set RecalcUnderway = 1 WHERE SkiYearID = " & sProcessingYear
Con.Execute(sSQL)

IF Request("Equival") <> "" THEN 
	Response.write(sSkiYearName & " Rankings ...")
	Response.Flush
END IF


' Section # 1

IF RunEquivScore = "YES" OR RunOverride = "YES" THEN

IF Request("Equival") <> "" THEN 
	Response.write(" Eq Scrs ...")
	Response.Flush
END IF







' ------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------
' ------    SECTION 1:    Prepare Equivalent Scores Table        ---------------------
' ------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------

' ----  GENERAL OVERVIEW  ----

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



' STEP 1: Generate the Query which will populate the Equivalent Divisions table,
'     for the current sProcessingYear.
' ----------------------------------------------------------------------------------


' -- Clear the table
sSQL = "DELETE FROM " & EquivDivsTableName
' WriteDebugSQL(sSQL)
Con.Execute(sSQL)



sSQL = "INSERT INTO " & EquivDivsTableName & " (MemberID, Event, Div, EStatus)"

' When sProcessingYear = 1, we need special Estatus codes to deal with Graduates.
' First Select folds in EStatus = "3" (Aged Out) for the "Graduating From" division,
'    this will later suppress these skiers in their "From" Division's Rankings,
' then the Second Union Select folds in an EStatus for the "Graduating To" division ...
'    That EStatus value for the "To" (incoming) Division will be "2", if there are 
'    scores for the event in the "From" Division, otherwise it will be "1"
'    Hence a "Presence" (EStatus=2) in the new division can be established
'    either by actual scores in that new "To" age division (from above), or by 
'    actual scores in the previous "From" age division, as derived here below.
' Then finally tack on a last UNION for the standard EStatus = 2 scores select.

IF sProcessingYear = 1 THEN

sSQL = sSQL & " SELECT MemberID, Event, Div, Max(EStatus) as EStatus FROM ("

sSQL = sSQL & " SELECT EL.MemberID, EL.Event, DT.Div, '3' as EStatus" 
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

sSQL = sSQL & " UNION SELECT GL.MemberID, GL.Event, GL.Next_Div as Div,"
sSQL = sSQL & " Case when EG.MemberID is not null then '2' else '1' end as EStatus FROM"
sSQL = sSQL & "  (SELECT EL.MemberID, EL.Event, DT.Div, DT.Next_Div"
sSQL = sSQL & "  FROM " & MemberTableName & "	as	MT," & DivisionsTableName & " as DT,"
sSQL = sSQL & "  (SELECT Year(BeginDate)-1 as BYear FROM " & SkiYearTableName & " where DefaultYear = 1) as EY,"
sSQL = sSQL & "  (SELECT MemberID, Event FROM " & RawScoresTableName & ","
sSQL = sSQL & "  (Select begindate as BDate, enddate as EDate from " & SkiYearTableName & " where SkiYearID = 1) as SY"
sSQL = sSQL & "  WHERE Score is not null and EndDate between SY.BDate and SY.EDate GROUP BY MemberID, Event) as EL"
sSQL = sSQL & "  WHERE	MT.PersonIDWithCheckDigit = EL.MemberID"
sSQL = sSQL & "  and Left(MT.Sex,1) = DT.Sex"
sSQL = sSQL & "  and EY.BYear - Year(MT.BirthDate) = DT.UP_Age"
sSQL = sSQL & "  and DT.Next_Div > 'AA'"
sSQL = sSQL & "  and	DT.SkiYearID = 1) as GL"
sSQL = sSQL & " LEFT JOIN (SELECT MemberID, Div, Event FROM " & RawScoresTableName & ","
sSQL = sSQL & "  (Select begindate as BDate, enddate as EDate from " & SkiYearTableName & " where SkiYearID = 1) as SY"
sSQL = sSQL & "  WHERE Score is not null and EndDate between SY.BDate and SY.EDate GROUP BY MemberID, Div, Event) as EG"
sSQL = sSQL & "  on EG.MemberID = GL.MemberID and EG.Div = GL.Div and EG.Event = GL.Event"

sSQL = sSQL & " UNION "

END IF

' Next EStatus = 2 Entries, derived from Actual performance rows present for this period.
' But only for "Ranking Divisions", where Left(RS.Div,1) in ('B','G','M','W','O','C'),
'	as well as translations from "International" divisions, as defined in detail below.  
' Hence omits Novice.

sSQL = sSQL & " SELECT RS.MemberID, RS.Event, CASE"
sSQL = sSQL & " when RS.Div = 'IM' then 'OM' when RS.Div = 'IW' then 'OW'"

' -- Test line for JM and JW --
sSQL = sSQL & " when RS.Div = 'JM' then 'OM' when RS.Div = 'JW' then 'OW'"
sSQL = sSQL & " when RS.Div = 'JB' then 'B2' when RS.Div = 'JG' then 'G2'"
sSQL = sSQL & " when RS.Div = 'IB' then 'B3' when RS.Div = 'IG' then 'G3'"
' sSQL = sSQL & " when RS.Div = 'S1' then 'M3' when RS.Div = 'L1' then 'W3'"
' sSQL = sSQL & " when RS.Div = 'S2' then 'M4' when RS.Div = 'L2' then 'W4'"
' sSQL = sSQL & " when RS.Div = 'S3' then 'M5' when RS.Div = 'L3' then 'W5'" 
sSQL = sSQL & " else RS.Div end as Div, '2' as EStatus FROM " & RawScoresTableName
sSQL = sSQL & " as RS, (Select begindate as BDate, enddate as EDate from " & SkiYearTableName & " where SkiYearID = "
sSQL = sSQL & sProcessingYear & " ) as SY WHERE RS.EndDate between SY.BDate and SY.EDate and RS.Score is not null"
sSQL = sSQL & " and Left(RS.Div,1) in ('B','G','I','J','M','S','W','L','C','O')"
sSQL = sSQL & " GROUP BY RS.MemberID, RS.Event, RS.Div"

' Finally 

IF sProcessingYear = 1 THEN
sSQL = sSQL & " ) as ED GROUP BY MemberID, Event, Div"
END IF

' Finally execute the constructed query

WriteDebugSQL(sSQL)

Con.Execute(sSQL)

IF Request("Equival") <> "" THEN 
	Response.write(" EqDiv.")
	Response.Flush
END IF




' -- STEP 2: Populate the Equivalent Scores table for this SkiYear --
' -------------------------------------------------------------------

'    First delete all EquivScores rows for the SkiYearID that we are processing.
sSQL = "DELETE FROM " & EquivScoresTableName & " WHERE SkiYearID = " & sProcessingYear

WriteDebugSQL(sSQL)

Con.Execute(sSQL)




' -- Extract all "Equivalent" scores into the EquivScores table.
' --      This is done separately by Event
' --         1) Slalom
' --         2) Trick
' --         3) Jump


' In each case spreading all scores out to 
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


' SLALOM: Using a single complex query, of two levels.  The innermost 
' level spreads each actual score out to other divisions in the EQUIVALENT DIVISIONS
' Table built above, where that combination is explicitly listed in the Division
' Control Table as an allowed equivalent, adjusting if necessary for any Max Speed 
' differences where applicable, and extracting the "Formatted Score" string for later 
' display.  The outer level then matches in the parameters for the effective division on 
' each such equivalenced score, derives the prioritized rating and class, and calculates
' the overall score component.

sSQL = "INSERT INTO " & EquivScoresTableName
sSQL = sSQL & " (MemberID, Team, TeamStat, TourID, Div, Event, Round, Class, PrioClass, Place, ScoreOrig, DivOrig, DivType, Score, Rating, FmtScore, OAScore, SkiYearID)"

sSQL = sSQL & " SELECT ES.MemberID, ES.Team, ES.TeamStat, ES.TourID, DE.Div, ES.Event, ES.Round, ES.Class, Case when ES.Class='R' then '5R'"
sSQL = sSQL & " when ES.Class='L' then '4L' when ES.Class='E' then '3E' when ES.Class='C' then '2C'"
sSQL = sSQL & " else '1' + ES.Class end as PrioClass, ES.Place, ES.ScoreOrig, ES.DivOrig,"
sSQL = sSQL & " Case when ES.Div = ES.DivOrig then 'A' else 'B' end, Case when ES.Score < 0 then 0 else ES.Score end,"
sSQL = sSQL & " Case when ES.Score >= DE.OP_S then '4O' when ES.Score >= DE.EP_S then '3E'"
sSQL = sSQL & " when ES.Score >= DE.MS_S then '2M' when ES.Score >= DE.XP_S then '1X' else '  ' end as Rating,"
sSQL = sSQL & " 'Rd ' + ES.Round + 'as ' + ES.DivOrig + '&#13;Score: ' + Cast (Cast(ES.Score as Decimal(5,2)) as Varchar(6)) + '&#13;' + ES.FmtScore,"
sSQL = sSQL & " Case when (DE.OverExp_S > 0) and (ES.Score < 6) then ES.Score * DE.OverPtsBy_S"
sSQL = sSQL & " when DE.OverExp_S > 0 then (6 * DE.OverPtsBy_S) + ((1500 - (6 * DE.OverPtsBy_S)) * Power ((ES.Score - 6) / (DE.NationalRec_S - 6), DE.OverExp_S ))"
sSQL = sSQL & " when ES.Score <= DE.FirstClass_S  and  DE.FirstClass_S > 0 then  200 * ES.Score / DE.FirstClass_S"
sSQL = sSQL & " when ES.Score <= DE.XP_S and DE.XP_S > DE.FirstClass_S then 200 + (200 * (ES.Score - DE.FirstClass_S) / (DE.XP_S - DE.FirstClass_S))"
sSQL = sSQL & " when ES.Score <= DE.MS_S and DE.MS_S > DE.XP_S then 400 + (200 * (ES.Score - DE.XP_S) / (DE.MS_S - DE.XP_S))"
sSQL = sSQL & " when ES.Score <= DE.EP_S and DE.EP_S > DE.MS_S then 600 + (200 * (ES.Score - DE.MS_S) / (DE.EP_S - DE.MS_S))"
sSQL = sSQL & " when DE.NationalRec_S > DE.EP_S then 800 + (700 * (ES.Score - DE.EP_S) / (DE.NationalRec_S - DE.EP_S))"
sSQL = sSQL & " else 0 end as OAScore, " & sProcessingYear & " as SkiYearID"
sSQL = sSQL & " FROM (SELECT RS.MemberID, RS.Team, RS.TeamStat, RS.TourID, DE.Div, RS.Event, RS.Round, RS.Class, RS.Place, RS.Score as ScoreOrig, RS.Div as DivOrig,"
sSQL = sSQL & " CASE when RS.Div = DE.Div then RS.Score"
' UPDATE 9/19/2017
' See email IMPORTANT -- Adjustment in Slalom scores for Graduating Skiers -- change for ZBS 2017+ logic dated 9/16/2017
'sSQL = sSQL & " when (DE.Max_S1 < DO.Max_S1 and RS.Perf_Qual2 > DE.Max_S1) then RS.Score - (2 * (DO.Max_S1 - DE.Max_S1))"
sSQL = sSQL & " when ( DE.Max_S1 < DO.Max_S1 and RS.Perf_Qual2 > DE.Max_S1 and UPPER(RS.Class) in ('E','L','R') ) then RS.Score - (2 * (DO.Max_S1 - DE.Max_S1))"
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
	sSQL = sSQL & " Where ST.EndDate between SY.BDate and SY.EDate and Substring(ST.TourID,7,1)>'B' and ST.Event = 'S')"
ELSE
	sSQL = sSQL & " (Select Distinct TourID From " & RawScoresTableName & " as ST,"
	sSQL = sSQL & " (Select begindate as BDate, enddate as EDate from " & SkiYearTableName & " where SkiYearID = " & sProcessingYear & ") as SY"
	sSQL = sSQL & " Where ST.EndDate between SY.BDate and SY.EDate and ST.Event = 'S')"
END IF

sSQL = sSQL & " and (RS.Div = DE.Div or (RS.Div = DE.SL_ED1 or RS.Div = DE.SL_ED2 or RS.Div = DE.SL_ED3 or RS.Div = DE.SL_ED4 or RS.Div = DE.SL_ED5 or RS.Div = DE.SL_ED6 or RS.Div = DE.SL_ED7 or RS.Div = DE.SL_ED8))"
'	sSQL = sSQL & " and (LEFT(DE.Div,1) <> 'O' or RS.Class in ('E','L','R'))" -- Provision removed 4/4/2010 in accord with 2010 rules change.
sSQL = sSQL & " and DO.Div = RS.Div and DO.SkiYearID = " & sProcessingYear
sSQL = sSQL & " and DE.Div = ED.Div and DE.SkiYearID = " & sProcessingYear
sSQL = sSQL & " and RS.MemberID = ED.MemberID and ED.Event = 'S'"
'	sSQL = sSQL & " and ED.EStatus <> '3' and (ED.EStatus = '2' OR RS.Div not in ('OM','OW','IM','IW','MM','MW','IB','IG'))"
sSQL = sSQL & " and ED.EStatus = '2'": ' Simplified from above 28 Apr 2010 to eliminate any EStatus = 1 bleeds
sSQL = sSQL & " ) as ES, " & DivisionsTableName & " as DE"
sSQL = sSQL & " WHERE ES.Div = DE.Div and DE.SkiYearID = " & sProcessingYear

 WriteDebugSQL(sSQL)

Con.Execute(sSQL)

IF Request("Equival") <> "" THEN 
	Response.write(" Slm.")
	Response.Flush
END IF


' TRICK - using a single complex query which spreads each actual 
' score out to other divisions listed for this skier in the EQUIVALENT DIVISION
' Table.  We extract the "Formatted Score" string for later display, matching in
' the parameters for the effective division on each such equivalenced score, and 
' then deriving a prioritized rating and calculating the overall value therewith.
' A prioritized class code is also derived.  Since the Division Control table 
' doesn't specifically list allowed equivalent divisions for Tricks, all 
' combinations would be allowed, except we screen for CM / CW to keep AWSA 
' scores out of those NCWSA divisions.

sSQL = "INSERT INTO " & EquivScoresTableName
sSQL = sSQL & " (MemberID, Team, TeamStat, TourID, Div, Event, Round, Class, PrioClass, Place, ScoreOrig, DivOrig, DivType, Score, Rating, FmtScore, OAScore, SkiYearID)"

sSQL = sSQL & " SELECT RS.MemberID, RS.Team, RS.TeamStat, RS.TourID, DE.Div, RS.Event, RS.Round, RS.Class, Case when RS.Class='R' then '5R'"
sSQL = sSQL & " when RS.Class='L' then '4L' when RS.Class='E' then '3E' when RS.Class='C' then '2C'"
sSQL = sSQL & " else '1' + RS.Class end as PrioClass, RS.Place, RS.Score, RS.Div,"
sSQL = sSQL & " Case when DE.Div = RS.Div then 'A' else 'B' end, RS.Score,"
sSQL = sSQL & " Case when RS.Score >= DE.OP_T then '4O' when RS.Score >= DE.EP_T then '3E'"
sSQL = sSQL & " when RS.Score >= DE.MS_T then '2M' when RS.Score >= DE.XP_T then '1X' else '  ' end as Rating,"
sSQL = sSQL & " 'Rd ' + RS.Round + 'as ' + RS.Div + '&#13;Score: ' + Cast (Cast(RS.Score as Decimal(6,0)) as Varchar(6)) + '&#13;Class: ' + RS.Class as FmtScore,"
sSQL = sSQL & " Case when  DE.OverExp_T > 0 then 1500 * Power(RS.Score/DE.NationalRec_T, DE.OverExp_T)"
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
	sSQL = sSQL & " Where ST.EndDate between SY.BDate and SY.EDate and Substring(ST.TourID,7,1)>'B' and ST.Event = 'T')"
ELSE
	sSQL = sSQL & " (Select Distinct TourID From " & RawScoresTableName & " as ST,"
	sSQL = sSQL & " (Select begindate as BDate, enddate as EDate from " & SkiYearTableName & " where SkiYearID = " & sProcessingYear & ") as SY"
	sSQL = sSQL & " Where ST.EndDate between SY.BDate and SY.EDate and ST.Event = 'T')"
END IF

'	sSQL = sSQL & " and (LEFT(DE.Div,1) <> 'O' or RS.Class in ('E','L','R'))" -- Provision removed 4/4/2010 in accord with 2010 rules change.
sSQL = sSQL & " and RS.MemberID = ED.MemberID and ED.Event = 'T'"
'	sSQL = sSQL & " and ED.EStatus <> '3' and (ED.EStatus = '2' OR RS.Div not in ('OM','OW','IM','IW','MM','MW','IB','IG'))"
sSQL = sSQL & " and ED.EStatus = '2'": ' Simplified from above 28 Apr 2010 to eliminate any EStatus = 1 bleeds
sSQL = sSQL & " and (left(DE.Div,1)<>'C' or left(RS.Div,1)='C')"
sSQL = sSQL & " and DE.Div = ED.Div and DE.SkiYearID = " & sProcessingYear

WriteDebugSQL(sSQL)

Con.Execute(sSQL)

IF Request("Equival") <> "" THEN 
	Response.write(" Trk.")
	Response.Flush
END IF


' JUMP - using a single complex query which spreads each actual score 
' out to other divisions listed for this skier in the EQUIVALENT DIVISION Table, 
' where the combination is explicitly cited in the Division Control Table, and
' where the actual conditions do not exceed the allowed limits for that division,
' extracting the "Formatted Score" string for later display, matching in the 
' parameters for the effective division on each such equivalenced score, and 
' deriving the prioritized rating and class, and calculating the overall 
' score component.

sSQL = "INSERT INTO " & EquivScoresTableName
sSQL = sSQL & " (MemberID, Team, TeamStat, TourID, Div, Event, Round, Class, PrioClass, Place, ScoreOrig, DivOrig, DivType, Score, Rating, FmtScore, OAScore, SkiYearID)"

sSQL = sSQL & " SELECT RS.MemberID, RS.Team, RS.TeamStat, RS.TourID, DE.Div, RS.Event, RS.Round, RS.Class"
sSQL = sSQL & " , CASE WHEN RS.Class='R' THEN '5R'"
sSQL = sSQL & "    WHEN RS.Class='L' THEN '4L'"
sSQL = sSQL & "    WHEN RS.Class='E' THEN '3E'"
sSQL = sSQL & "    WHEN RS.Class='C' THEN '2C'"
sSQL = sSQL & "    ELSE '1' + RS.Class END AS PrioClass"
sSQL = sSQL & " , RS.Place, RS.Score, RS.Div,"
sSQL = sSQL & " CASE WHEN DE.Div = RS.Div THEN 'A' ELSE 'B' END, RS.Score,"

sSQL = sSQL & " CASE WHEN ( (RS.Score >= DE.OP_J AND RS.Perf_Qual1 <= DE.Ramp1 AND RS.Perf_Qual2 <= DE.Max_J1)" 
sSQL = sSQL & "          OR (DE.OP_J2 > 0 AND RS.Score >= DE.OP_J2 AND RS.Perf_Qual1 <= DE.Ramp2 AND RS.Perf_Qual2 <= DE.Max_J2) ) THEN '4O'"
sSQL = sSQL & "    WHEN RS.Score >= DE.EP_J then '3E' when RS.Score >= DE.MS_J then '2M'"
sSQL = sSQL & "    WHEN RS.Score >= DE.XP_J THEN '1X' ELSE '  ' END as Rating,"

sSQL = sSQL & " 'Rd ' + RS.Round + 'as ' + RS.Div + '&#13;Score: ' + Cast (Cast(RS.Score as Decimal(3)) as Varchar(4)) + '&#13;' + Cast (Cast(RS.Perf_Qual1 as Decimal(5,3)) as Varchar(6)) + ' @ ' + Cast (Cast(RS.Perf_Qual2 as Decimal(3)) as Varchar(3)) + 'k&#13;Class: ' + RS.Class as FmtScore,"

sSQL = sSQL & " CASE WHEN  (DE.OverExp_J > 0) and (RS.Score < (0.15*DE.NationalRec_J)) then 0"
sSQL = sSQL & "   WHEN DE.OverExp_J > 0 THEN 1500 * Power ((RS.Score - (0.15*DE.NationalRec_J)) / (DE.NationalRec_J - (0.15*DE.NationalRec_J)), DE.OverExp_J)"
sSQL = sSQL & "   WHEN RS.Score <= DE.FirstClass_J AND DE.FirstClass_J > 0 THEN  200 * RS.Score / DE.FirstClass_J"
sSQL = sSQL & "   WHEN RS.Score <= DE.XP_J AND  DE.XP_J > DE.FirstClass_J THEN 200 + (200 * (RS.Score - DE.FirstClass_J) / (DE.XP_J - DE.FirstClass_J))"
sSQL = sSQL & "   WHEN RS.Score <= DE.MS_J AND DE.MS_J > DE.XP_J THEN 400 + (200 * (RS.Score - DE.XP_J) / (DE.MS_J - DE.XP_J))"
sSQL = sSQL & "   WHEN RS.Score <= DE.EP_J AND  DE.EP_J > DE.MS_J THEN 600 + (200 * (RS.Score - DE.MS_J) / (DE.EP_J - DE.MS_J))"
sSQL = sSQL & "   WHEN DE.NationalRec_J > DE.EP_J THEN 800 + (700 * (RS.Score - DE.EP_J) / (DE.NationalRec_J - DE.EP_J))"
sSQL = sSQL & "   ELSE 0 END AS OAScore, " & sProcessingYear & " AS SkiYearID"

sSQL = sSQL & " FROM " & RawScoresTableName & " AS RS, " & DivisionsTableName & " AS DE, " & EquivDivsTableName & " AS ED"

sSQL = sSQL & " WHERE RS.Event = 'J'" 
sSQL = sSQL & "   AND RS.Score IS NOT NULL" 
sSQL = sSQL & "   AND (UPPER(RS.Class) in ('F','N','I','C','E','L','R'))" 
sSQL = sSQL & "   AND RS.TourID IN"

IF sProcessingYear = 1 THEN
	sSQL = sSQL & " (SELECT TourID FROM (Select Rgn, Right(Max(DateTour),7) as TourID FROM (Select Distinct"
	sSQL = sSQL & "   CASE WHEN Substring(TourID,7,1)='A' THEN 'N' ELSE Substring(TourID,3,1) END AS Rgn,"
	sSQL = sSQL & "   CONVERT(char,EndDate,112) + Left(TourID,7) AS DateTour FROM " & RawScoresTableName
	sSQL = sSQL & "    WHERE  Substring(TourID,7,1) in ('A','B') ) as RDT Group by Rgn) AS RNT"
	sSQL = sSQL & " UNION" 
	sSQL = sSQL & " SELECT Distinct TourID FROM " & RawScoresTableName & " AS ST,"
	sSQL = sSQL & "  (SELECT begindate as BDate, enddate AS EDate FROM " & SkiYearTableName & " WHERE SkiYearID = " & sProcessingYear & ") as SY"
	sSQL = sSQL & "    WHERE ST.EndDate BETWEEN SY.BDate AND SY.EDate AND Substring(ST.TourID,7,1)>'B' AND ST.Event = 'J')"
ELSE
	sSQL = sSQL & " (SELECT Distinct TourID FROM " & RawScoresTableName & " AS ST,"
	sSQL = sSQL & " (SELECT begindate as BDate, enddate as EDate from " & SkiYearTableName & " where SkiYearID = " & sProcessingYear & ") as SY"
	sSQL = sSQL & "    WHERE ST.EndDate between SY.BDate and SY.EDate and ST.Event = 'J')"
END IF

sSQL = sSQL & " AND  (RS.Div = DE.Div"
sSQL = sSQL & "   OR ("
sSQL = sSQL & "     (  RS.Div = DE.JU_ED1" 
sSQL = sSQL & "     OR RS.Div = DE.JU_ED2" 
sSQL = sSQL & "     OR RS.Div = DE.JU_ED3"
sSQL = sSQL & "     OR RS.Div = DE.JU_ED4" 
sSQL = sSQL & "     OR RS.Div = DE.JU_ED5"
sSQL = sSQL & "     OR RS.Div = DE.JU_ED6" 
sSQL = sSQL & "     OR RS.Div = DE.JU_ED7" 
sSQL = sSQL & "     OR RS.Div = DE.JU_ED8 )" 
sSQL = sSQL & "      AND RS.Perf_Qual1 <= DE.Ramp1"
sSQL = sSQL & "     AND RS.Perf_Qual2 <= DE.Max_J1)"
sSQL = sSQL & "  )"

sSQL = sSQL & " AND RS.MemberID = ED.MemberID"
sSQL = sSQL & " AND ED.Event = 'J'"
sSQL = sSQL & " AND ED.EStatus = '2'": ' Simplified from above 28 Apr 2010 to eliminate any EStatus = 1 bleeds
sSQL = sSQL & " AND DE.Div = ED.Div"
sSQL = sSQL & " AND DE.SkiYearID = " & sProcessingYear

WriteDebugSQL(sSQL)

Con.Execute(sSQL)

IF Request("Equival") <> "" THEN 
	Response.write(" Jmp.")
	Response.Flush
END IF




' --- STEP 3:  Update query to "Cap" any Equivalent Scores in Class F/I/N, where the
' score exceeds the preceding ski year's level 5 COA score for the applicable Div/Event.

sSQL = "UPDATE ES Set Score = CO.COA5"
sSQL = sSQL & " FROM " & EquivScoresTableName & " AS ES, " & CutOffTableName & " as CO"
sSQL = sSQL & " WHERE	CO.Div = ES.Div and CO.Event = ES.Event"	
sSQL = sSQL & " AND	ES.SkiYearID = " & sProcessingYear & " AND CO.SkiYearID = " & sPrevYear
sSQL = sSQL & " AND	UPPER(ES.Class) in ('F','I','N') AND ES.Score > CO.COA5"

WriteDebugSQL(sSQL)

Con.Execute(sSQL)

IF Request("Equival") <> "" THEN 
	Response.write(" Cap.")
	Response.Flush
END IF




'	--- STEP 4:  Insert to copy in all eligible "Elite Pool"
'	candidate performances, from selected divisions, adjusting Jump scores where 
'	the original division conditions are less than those of the applicable Elite 
'	Division.  Overall component values are also calculated, using the parameters 
'	for the applicable Elite Division.  This process stages creation of the four
'	Consolidated "Elite Pool" rankings, in all 3 events plus overall for each Pool,
'	which will then naturally occur during the subsequent portions of this nightly 
'	process, for these additional Elite Division which we're loading here.
'
' This insert exists as a three-level nested process.  The innermost subquery 
'	reduces any duplicate occurrences of the same original performance -- which 
'	may have already been equivalenced between an age division and an Elite 
'	division -- giving precedence to the original age division code, to support 
'	the Jump adjustments.  The second level subquery then applies adjustments to 
'	Jump scores, for differences (if any) between the Native age division and 
'	the applicable Elite division.  Then finally the outermost select computes
'	the overall score for that performance, using the applicable Elite division 
'	overall parameters.

'	Run as two overall steps, because of overlapping selection and DivType coding
'	to C or D, and with M3/M4/M5/M6/MM appearing in BOTH candidate sets, in order to
'	support separate overall qualification for MM versus OM, for these guys.  Messy.





'	STEP 4a) First step here inserts EM/EW from possible candidates
'            DivType Decode: Depending on whether DivBase is in Primary 3.03(c)3 set or not 
'                C =  M1,M2,OM,W1,W2,OW
'                D = all other divisions

sSQL = "INSERT INTO " & EquivScoresTableName
sSQL = sSQL & " (MemberID, Team, TeamStat, TourID, Div, Event, Round, Class, PrioClass, Place, ScoreOrig, DivOrig, DivType, Score, Rating, FmtScore, SkiYearID, OAScore)"

sSQL = sSQL & " SELECT ES.MemberID, ES.Team, ES.TeamStat, ES.TourID, DE.Div, ES.Event, ES.Round, ES.Class,"
sSQL = sSQL & " ES.PrioClass, ES.Place, ES.ScoreOrig, ES.DivOrig, ES.DivType, ES.Score, ES.Rating, ES.FmtScore, ES.SkiYearID,"
sSQL = sSQL & " CASE when ES.Event = 'S' then Case when ES.Score < 6 then ES.Score * DE.OverPtsBy_S"
sSQL = sSQL & " else (6 * DE.OverPtsBy_S) + ((1500 - (6 * DE.OverPtsBy_S)) * Power"
sSQL = sSQL & " ((ES.Score - 6) / (DE.NationalRec_S - 6), DE.OverExp_S )) end"
sSQL = sSQL & " when ES.Event = 'T' then 1500 * Power(ES.Score/DE.NationalRec_T, DE.OverExp_T)"
sSQL = sSQL & " when ES.Event = 'J' then 1500 * Power ((ES.Score - (0.15*DE.NationalRec_J))" 
sSQL = sSQL & " / (DE.NationalRec_J - (0.15*DE.NationalRec_J)), DE.OverExp_J)"
sSQL = sSQL & " end as OAScore FROM (SELECT XS.MemberID, XS.TourID, XS.Event, XS.Round, XS.Place, XS.DivOrig, XS.DivElite as Div,"
sSQL = sSQL & " XS.Rating, XS.Class, XS.PrioClass, XS.Team, XS.TeamStat, XS.Score as ScoreOrig, XS.FmtScore, XS.SkiYearID,"

' -- Changed 3/4/2019 by Mark Crone to elimate JUMP equivalencies -- 
sSQL = sSQL & " CASE WHEN DivBase IN ('M1','M2','OM','W1','W2','OW') THEN 'C' ELSE 'D' END AS DivType,"
sSQL = sSQL & " CASE WHEN XS.Event IN ('S','T') OR DE.Ramp1 < DO.Ramp1 OR DE.Max_J1 < DO.Max_J1 THEN XS.Score"
' sSQL = sSQL & "     ELSE XS.Score + (12 * ((DE.Ramp1-DO.Ramp1)/.020)) + (8 * ((DE.Max_J1-DO.Max_J1)/3)) END AS Score"
sSQL = sSQL & "     ELSE XS.Score END AS Score"

sSQL = sSQL & " FROM (SELECT MemberID, TourID, Event, Round, Class, DivOrig, Place, SkiYearID, Team, PrioClass, TeamStat,"
sSQL = sSQL & "          CASE WHEN Max(Div) IN ('OW','W1','W2','W3','W4','W5') THEN Min(Div) ELSE Max(Div) END AS DivBase," 
sSQL = sSQL & "             Max(Score) as Score, Max(FmtScore) AS FmtScore, Max(Rating) as Rating,"
sSQL = sSQL & "             Max(CASE WHEN Div in ('B2','B3','M1','M2','M3','M4','M5','OM','MM') THEN 'EM' ELSE 'EW' END) as DivElite"
sSQL = sSQL & "                FROM " & EquivScoresTableName 
sSQL = sSQL & "                   WHERE SkiYearID = " & sProcessingYear
sSQL = sSQL & "                      AND Div in ('B2','B3','M1','M2','M3','M4','M5','OM','MM','G2','G3','W1','W2','W3','W4','W5','OW','MW')"
sSQL = sSQL & "             GROUP BY 	MemberID, TourID, Event, Round, Class, DivOrig, Place, SkiYearID, Team, PrioClass, TeamStat) as XS, "
sSQL = sSQL & DivisionsTableName & " as DO, " & DivisionsTableName & " as DE Where DO.Div = XS.DivOrig"
' sSQL = sSQL & DivisionsTableName & " as DO, " & DivisionsTableName & " as DE Where DO.Div = XS.DivBase"
sSQL = sSQL & " and DO.SkiYearID = XS.SkiYearID and	DE.Div = XS.DivElite and DE.SkiYearID = XS.SkiYearID) as ES, "
sSQL = sSQL & DivisionsTableName & " as DE WHERE DE.Div = ES.Div and DE.SkiYearID = ES.SkiYearID"

Con.Execute(sSQL)




'	STEP 4b) Second step here inserts SM/SW from possible candidates, Coding as DivType C or D, 
'            DivType Decode: Depending on whether DivBase is in Primary 3.03(c)4 set or not
'                C = M3,M4,MM,W3,W4,MW
'                D = All other divisions

sSQL = "INSERT INTO " & EquivScoresTableName
sSQL = sSQL & " (MemberID, Team, TeamStat, TourID, Div, Event, Round, Class, PrioClass, Place, ScoreOrig, DivOrig, DivType, Score, Rating, FmtScore, SkiYearID, OAScore)"

sSQL = sSQL & " SELECT ES.MemberID, ES.Team, ES.TeamStat, ES.TourID, DE.Div, ES.Event, ES.Round, ES.Class,"
sSQL = sSQL & " ES.PrioClass, ES.Place, ES.ScoreOrig, ES.DivOrig, ES.DivType, ES.Score, ES.Rating, ES.FmtScore, ES.SkiYearID,"
sSQL = sSQL & " CASE when ES.Event = 'S' then Case when ES.Score < 6 then ES.Score * DE.OverPtsBy_S"
sSQL = sSQL & " else (6 * DE.OverPtsBy_S) + ((1500 - (6 * DE.OverPtsBy_S)) * Power"
sSQL = sSQL & " ((ES.Score - 6) / (DE.NationalRec_S - 6), DE.OverExp_S )) end"
sSQL = sSQL & " when ES.Event = 'T' then 1500 * Power(ES.Score/DE.NationalRec_T, DE.OverExp_T)"
sSQL = sSQL & " when ES.Event = 'J' then 1500 * Power ((ES.Score - (0.15*DE.NationalRec_J))" 
sSQL = sSQL & " / (DE.NationalRec_J - (0.15*DE.NationalRec_J)), DE.OverExp_J)"
sSQL = sSQL & " end as OAScore FROM (SELECT XS.MemberID, XS.TourID, XS.Event, XS.Round, XS.Place, XS.DivOrig, XS.DivElite as Div,"
sSQL = sSQL & " XS.Rating, XS.Class, XS.PrioClass, XS.Team, XS.TeamStat, XS.Score as ScoreOrig, XS.FmtScore, XS.SkiYearID,"

sSQL = sSQL & " CASE WHEN DivBase in ('M3','M4','MM','W3','W4','MW') THEN 'C' ELSE 'D' END AS DivType,"

sSQL = sSQL & " CASE WHEN XS.Event IN ('S','T')"
sSQL = sSQL & "     OR DE.Ramp1 < DO.Ramp1"
sSQL = sSQL & "     OR DE.Max_J1 < DO.Max_J1 THEN XS.Score"
sSQL = sSQL & "    ELSE XS.Score + (12 * ((DE.Ramp1-DO.Ramp1)/.020)) + (8 * ((DE.Max_J1-DO.Max_J1)/3)) END AS Score"

sSQL = sSQL & " FROM (SELECT MemberID, TourID, Event, Round, Class, DivOrig, Place, SkiYearID, Team, PrioClass, TeamStat,"
sSQL = sSQL & "       CASE WHEN Max(Div) IN ('MW','W3','W4','W5','W6') THEN Min(Div) ELSE Max(Div) END AS DivBase," 
sSQL = sSQL & "        Max(Score) as Score, Max(FmtScore) as FmtScore, Max(Rating) as Rating,"
sSQL = sSQL & "        Max(CASE WHEN Div IN ('M3','M4','M5','M6','M7','MM') THEN 'SM' ELSE 'SW' END) AS DivElite"
sSQL = sSQL & "            FROM " & EquivScoresTableName
sSQL = sSQL & "                WHERE SkiYearID = " & sProcessingYear 
sSQL = sSQL & "                  AND Div IN ('M3','M4','M5','M6','M7','MM','W3','W4','W5','W6','MW')"
sSQL = sSQL & "            GROUP BY MemberID, TourID, Event, Round, Class, DivOrig, Place, SkiYearID, Team, PrioClass, TeamStat) as XS, "
sSQL = sSQL & DivisionsTableName & " as DO, " & DivisionsTableName & " as DE Where DO.Div = XS.DivOrig"
' sSQL = sSQL & DivisionsTableName & " as DO, " & DivisionsTableName & " as DE Where DO.Div = XS.DivBase"
sSQL = sSQL & " and DO.SkiYearID = XS.SkiYearID and	DE.Div = XS.DivElite and DE.SkiYearID = XS.SkiYearID) as ES, "
sSQL = sSQL & DivisionsTableName & " as DE WHERE DE.Div = ES.Div and DE.SkiYearID = ES.SkiYearID"

Con.Execute(sSQL)

IF Request("Equival") <> "" THEN 
	Response.write(" Elite.")
	Response.Flush
END IF





' -----------------------------------------------------------------------------------------------------
' -- Adds EVENT skiers to OM/OW or MM/MW division when they have never participated in that division --
' -----------------------------------------------------------------------------------------------------

' -- CONCEPT:
' -- Added 3-25-2018 by Mark Crone 
' -- Needed because a member may achieve a Level 10 score but has never skied in an Open/Masters division so no ranking would be calculated
' --   This process adds member which overrides the normal approach
' -- Selects scores for all members from EM/EW/SM/SW divs to insert just one set of scores into the Open/Masters divisions
' -- Second section makes query return only those who have a Level 10 score
' -- Last part of query limits to Member/Events/Divs that do not have Div in OM/OW/MM/MW



sSQL = "		INSERT INTO " & EquivScoresTableName			
sSQL = sSQL + " (MemberID, TourID, Div, Event, Round, Class, ScoreOrig, DivOrig, Score, Place, Rating"
sSQL = sSQL + " , FmtScore, OAScore, SkiYearID, Team, PrioClass, TeamStat, DivType)"

sSQL = sSQL + " SELECT es1.MemberID, TourID"
sSQL = sSQL + "		, CASE WHEN es1.div='EM' THEN 'OM'"
sSQL = sSQL + "					WHEN es1.div='EW' THEN 'OW'"
sSQL = sSQL + "					WHEN es1.div='SM' THEN 'MM'"
sSQL = sSQL + "					WHEN es1.div='SW' THEN 'MW' END AS Div"	 
sSQL = sSQL + "		, es1.Event, Round, Class, ScoreOrig, DivOrig, Score, Place, Rating"
sSQL = sSQL + "		, FmtScore, OAScore, es1.SkiYearID, Team, PrioClass, TeamStat, DivType"
		
sSQL = sSQL + "		FROM " & EquivScoresTableName & " es1" 
sSQL = sSQL + "		JOIN " & SkiYearTableName & " sy ON sy.skiyearid = es1.skiyearid" 
sSQL = sSQL + "		JOIN " & DivisionsTableName & " dt ON dt.SkiYearid = sy.SkiYearid AND dt.div = es1.Div" 
sSQL = sSQL + "		JOIN " & MemberTableName & " mt ON mt.PersonID=RIGHT(es1.MemberID,8)"
		
'   							-- Members who have a Level 10 score in any division in the default Ski Year
sSQL = sSQL + "		LEFT JOIN" 
sSQL = sSQL + "		( SELECT st.MemberID, event, dt.div AS EliteDiv" 
sSQL = sSQL + "				FROM " & EquivScoresTableName & " st" 
sSQL = sSQL + "				JOIN " & SkiYearTableName & " sy ON sy.skiyearid = st.skiyearid" 
sSQL = sSQL + "				JOIN " & DivisionsTableName & " dt ON dt.SkiYearid = sy.SkiYearid"
sSQL = sSQL + "					AND dt.div = CASE WHEN st.Div in ('M3','M4','M5','M6','M7','MM') THEN 'SM'" 
sSQL = sSQL + "                              WHEN st.Div in ('W3','W4','W5','W6','W7','MW') THEN 'SW'"
sSQL = sSQL + "                              WHEN st.Div in ('B2','B3','M1','M2','MM','JM') THEN 'EM'" 
sSQL = sSQL + "                              WHEN st.Div in ('G2','G3','W1','W2','MW','JW') THEN 'EW'" 
sSQL = sSQL + "                              ELSE 'PP' END"	   
sSQL = sSQL + "						WHERE sy.defaultyear = 1" 
' sSQL = sSQL + "							AND st.div in ('EM','EW','SM','SW')"
sSQL = sSQL + "							AND ( (st.event = 'S' AND st.score >= dt.lvl_10_s) OR"
sSQL = sSQL + "										(st.event = 'T' AND st.score >= dt.lvl_10_t) OR"
sSQL = sSQL + "										(st.event = 'J' AND st.score >= dt.lvl_10_j) )"

' sSQL = sSQL & " and RS.Score >= DE.OP_J2 and RS.Perf_Qual1 <= DE.Ramp2 and RS.Perf_Qual2 <= DE.Max_J2))


sSQL = sSQL + "				GROUP BY st.MemberID, st.event, dt.div"
sSQL = sSQL + "		) es2"
sSQL = sSQL + "		ON es2.MemberID=es1.MemberID AND es2.Event=es1.Event AND es2.EliteDiv=es1.Div"

'									-- Suppress those who have and OM/OW or MM/MW scores in processing year
sSQL = sSQL + "		LEFT JOIN"
sSQL = sSQL + "		( SELECT st.MemberID, st.Div, st.Event" 
sSQL = sSQL + "				FROM " & EquivScoresTableName & " st" 
sSQL = sSQL + "				JOIN " & SkiYearTableName & " sy ON sy.skiyearid = st.skiyearid" 
sSQL = sSQL + "				JOIN " & DivisionsTableName & " dt ON dt.SkiYearid = sy.SkiYearid AND dt.div = st.div"

sSQL = sSQL + "						WHERE st.SkiYearID="& sProcessingYear
sSQL = sSQL + "						AND st.div IN ('OM','OW','MM','MW')"			  
sSQL = sSQL + "				GROUP BY st.memberid, st.Event, st.Div )	es3"
sSQL = sSQL + "		ON es3.MemberID=es1.MemberID AND es2.Event=es1.Event" 
sSQL = sSQL + "					AND es3.Div=CASE WHEN es1.Div='SM' THEN 'MM'" 
sSQL = sSQL + "                              WHEN es1.Div='SW' THEN 'MW'"
sSQL = sSQL + "                              WHEN es1.Div='EM' THEN 'OM'" 
sSQL = sSQL + "                              WHEN es1.Div='EW' THEN 'OW'" 
sSQL = sSQL + "                              ELSE 'PP' END"	
		
sSQL = sSQL + "				WHERE sy.SkiYearID="& sProcessingYear
sSQL = sSQL + "					AND es1.div IN ('EM','EW','SM','SW')"
sSQL = sSQL + "					AND mt.federationcode = 'USA'"
sSQL = sSQL + "					AND sy.skiyear-datepart(yyyy,mt.birthdate)-1 >= 18"
sSQL = sSQL + "					AND es2.MemberID IS NOT NULL" 
sSQL = sSQL + "					AND es3.MemberID IS NULL"	
sSQL = sSQL + "	ORDER BY es1.MemberID, es1.Event, es1.Div"


WriteDebugSQL(sSQL)

Con.Execute(sSQL)





' -- STEP 5: Post a new UPDATE Date in SkiYear Table to show Last Time/Date Recalculated  --

sSQL = "UPDATE " & SkiYearTableName & " SET LastRecalc = '" & time() & " on " & date() & "' WHERE SkiYearID = " & sProcessingYear 

' WriteDebugSQL(sSQL)

Con.Execute(sSQL)



' Bottom of IF to exclude the Equivalent Scores section above from running
END IF


' -----------------------------------------------------------------------------------------------------
' -----------------------------   End of EQUIVALENT SCORES Processing  --------------------------------
' -----------------------------------------------------------------------------------------------------









' -------------------------------------------------------------------------------
' --------------  SECTION 2:   Calculate OVERALL SCORES  ------------------------
' -------------------------------------------------------------------------------


IF RunOvrllScore = "YES" OR RunOverride = "YES" THEN

IF Request("Equival") <> "" THEN 
	Response.write(" OA Scrs ...")
	Response.Flush
END IF




' STEP 1: Start assembling overall scores.  But first we must Delete 
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
' But only for Divisions that start with B/G/M/W/O/E/W -- leaves others out.

sSQL = "SELECT ES.*, DE.OverNumEvts FROM " & EquivScoresTableName & " ES, " & DivisionsTableName & " as DE"
sSQL = sSQL & " WHERE ES.Div = DE.Div AND ES.SkiYearID = DE.SkiYearID AND ES.SkiYearID = " & sProcessingYear
sSQL = sSQL & " AND LEFT(ES.Div,1) in ('B','G','M','W','O','E','S')"
sSQL = sSQL & " ORDER BY ES.memberid, ES.tourid, ES.div, ES.round"

WriteDebugSQL(sSQL)

rs.open sSQL, sConnectionToTRATable, 3, 3

IF not rs.eof THEN

tempvar = rs.getrows()
nTotalMembers = ubound(tempvar,2)

rs.MoveFirst
nProcessedSoFar = 0




' ---------  Outer loop of Overall Score Assembly for all MemberID and TourID's   -----------

DO UNTIL rs.eof

	nProcessedSoFar = nProcessedSoFar + 1
		
	IF Request("Equival") <> "" THEN 
		IF (nProcessedSoFar mod 10 = 5) THEN ShowProgress (nProcessedSoFar / nTotalMembers)
	END IF

	TempMemberID = rs("MemberID"): TempTeam = trim(rs("Team"))
	TempTourID = rs("TourID"): TempDiv = rs("Div"): TempDivType = rs("DivType")
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
		IF TempTeam < trim(rs("Team")) THEN TempTeam = trim(rs("Team"))
		IF TempDivType < rs("DivType") THEN TempDivType = rs("DivType")


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
  ' -----------------------------------------------------------------------------------------------------------------
  
  ' ROUND 1 Overall Score -- Determine Overall Eligibility for this Round Then Insert if so
  
  TempOverEvts = 0: TempOATot = 0
  IF Slalom1 <> "" THEN TempOverEvts = TempOverEvts + 1: TempOATot = TempOATot + Slalom1
  IF Trick1 <> "" THEN TempOverEvts = TempOverEvts + 1: TempOATot = TempOATot + Trick1
  IF Jump1 <> "" THEN TempOverEvts = TempOverEvts + 1: TempOATot = TempOATot + Jump1

  IF TempOverEvts >= TempOverEvtsReq THEN
  
		sSQL = "INSERT INTO " & OverallScoresTableName
		sSQL = sSQL & " (MemberID, Team, TourID, Round, Div, Class, PrioClass, SlalomOverAll, S_Round, S_OrigScore, JumpOverAll, J_Round, J_OrigScore, TrickOverAll, T_Round, T_OrigScore, TotalOverAll, SkiYearID, DivOrig, DivType)"
		sSQL = sSQL & " VALUES ('" & TempMemberID & "','" & TempTeam & "','" & TempTourID & "','1','" & TempDiv & "','" & Right(Class1,1) & "','" & Class1 & "',"
		IF Slalom1 = "" THEN sSQL = sSQL & "NULL,NULL,NULL," ELSE sSQL = sSQL & "'" & Slalom1 & "','" & S_Round1 & "','" & S_Score1 & "',"
 		IF Jump1 = "" THEN sSQL = sSQL & "NULL,NULL,NULL," ELSE sSQL = sSQL & "'" & Jump1 & "','" & J_Round1 & "','" & J_Score1 & "',"
		IF Trick1 = "" THEN sSQL = sSQL & "NULL,NULL,NULL," ELSE sSQL = sSQL & "'" & Trick1 & "','" & T_Round1 & "','" & T_Score1 & "',"
		sSQL = sSQL & "'" & TempOATot & "','" & sProcessingYear & "','" & TempDivOrig & "','" & TempDivType & "')"
		
		Con.Execute(sSQL)

	END IF


WriteDebugSQL("Round 1 Overall (Temp) - " & sSQL)


  ' ROUND 2 Overall Score -- Determine Overall Eligibility for this Round Then Insert if so
  
  TempOverEvts = 0: TempOATot = 0
  IF Slalom2 <> "" THEN TempOverEvts = TempOverEvts + 1: TempOATot = TempOATot + Slalom2
  IF Trick2 <> "" THEN TempOverEvts = TempOverEvts + 1: TempOATot = TempOATot + Trick2
  IF Jump2 <> "" THEN TempOverEvts = TempOverEvts + 1: TempOATot = TempOATot + Jump2

  IF TempOverEvts >= TempOverEvtsReq THEN
  
		sSQL = "INSERT INTO " & OverallScoresTableName
		sSQL = sSQL & " (MemberID, Team, TourID, Round, Div, Class, PrioClass, SlalomOverAll, S_Round, S_OrigScore, JumpOverAll, J_Round, J_OrigScore, TrickOverAll, T_Round, T_OrigScore, TotalOverAll, SkiYearID, DivOrig, DivType)"
		sSQL = sSQL & " VALUES ('" & TempMemberID & "','" & TempTeam & "','" & TempTourID & "','2','" & TempDiv & "','" & Right(Class2,1) & "','" & Class2 & "',"
		IF Slalom2 = "" THEN sSQL = sSQL & "NULL,NULL,NULL," ELSE sSQL = sSQL & "'" & Slalom2 & "','" & S_Round2 & "','" & S_Score2 & "',"
 		IF Jump2 = "" THEN sSQL = sSQL & "NULL,NULL,NULL," ELSE sSQL = sSQL & "'" & Jump2 & "','" & J_Round2 & "','" & J_Score2 & "',"
		IF Trick2 = "" THEN sSQL = sSQL & "NULL,NULL,NULL," ELSE sSQL = sSQL & "'" & Trick2 & "','" & T_Round2 & "','" & T_Score2 & "',"
		sSQL = sSQL & "'" & TempOATot & "','" & sProcessingYear & "','" & TempDivOrig & "','" & TempDivType & "')"

		Con.Execute(sSQL)

	END IF


  ' ROUND 3 Overall Score -- Determine Overall Eligibility for this Round Then Insert if so
  
  TempOverEvts = 0: TempOATot = 0
  IF Slalom3 <> "" THEN TempOverEvts = TempOverEvts + 1: TempOATot = TempOATot + Slalom3
  IF Trick3 <> "" THEN TempOverEvts = TempOverEvts + 1: TempOATot = TempOATot + Trick3
  IF Jump3 <> "" THEN TempOverEvts = TempOverEvts + 1: TempOATot = TempOATot + Jump3

  IF TempOverEvts >= TempOverEvtsReq THEN
  
		sSQL = "INSERT INTO " & OverallScoresTableName
		sSQL = sSQL & " (MemberID, Team, TourID, Round, Div, Class, PrioClass, SlalomOverAll, S_Round, S_OrigScore, JumpOverAll, J_Round, J_OrigScore, TrickOverAll, T_Round, T_OrigScore, TotalOverAll, SkiYearID, DivOrig, DivType)"
		sSQL = sSQL & " VALUES ('" & TempMemberID & "','" & TempTeam & "','" & TempTourID & "','3','" & TempDiv & "','" & Right(Class3,1) & "','" & Class3 & "',"
		IF Slalom3 = "" THEN sSQL = sSQL & "NULL,NULL,NULL," ELSE sSQL = sSQL & "'" & Slalom3 & "','" & S_Round3 & "','" & S_Score3 & "',"
		IF Jump3 = "" THEN sSQL = sSQL & "NULL,NULL,NULL," ELSE sSQL = sSQL & "'" & Jump3 & "','" & J_Round3 & "','" & J_Score3 & "',"
		IF Trick3 = "" THEN sSQL = sSQL & "NULL,NULL,NULL," ELSE sSQL = sSQL & "'" & Trick3 & "','" & T_Round3 & "','" & T_Score3 & "',"
		sSQL = sSQL & "'" & TempOATot & "','" & sProcessingYear & "','" & TempDivOrig & "','" & TempDivType & "')"

		Con.Execute(sSQL)

	END IF
 
' -----------------------  Bottom of Outer LOOP for OVERALL SCORES  ----------------------
LOOP

END IF

rs.close








' -------------------------------------------------------------------------------------------------------
' -- Adds OVERALL skiers to OM/OW or MM/MW division when they have never participated in that division --
' -------------------------------------------------------------------------------------------------------

' -- Added 3-25-2018 by Mark Crone 

' -- CONCEPT:
' -- Needed because a member may achieve a Level 10 score but has never skied in an Open/Masters division so no ranking would be calculated
' --   This process adds member which overrides the normal approach
' -- Selects Overall scores for all members from EM/EW/SM/SW divs to insert just one set of scores into the Open/Masters divisions
' -- Second section makes query return only those who have a Level 10 score
' -- Last part of query limits to Member/Divs that do not have Div in OM/OW/MM/MW
' -- Although at the time of development of this query, Overall exists only in OM, query is set up to handle all four 


sSQL = "INSERT INTO " & OverAllScoresTableName
sSQL = sSQL + "	SELECT es1.MemberID, TourID, Round"
sSQL = sSQL + "		, CASE WHEN es1.div='EM' THEN 'OM'"
sSQL = sSQL + "					WHEN es1.div='EW' THEN 'OW'"
sSQL = sSQL + "					WHEN es1.div='SM' THEN 'MM'"
sSQL = sSQL + "					WHEN es1.div='SW' THEN 'MW' END AS Div"	 
sSQL = sSQL + "		, SlalomOverall, S_Round, S_OrigScore, JumpOverall, J_Round, J_OrigScore, TrickOverall, T_Round, T_OrigScore"
sSQL = sSQL + "		, TotalOverall, es1.SkiYearID, DivOrig, Team, Class, PrioClass, DivType"  
sSQL = sSQL + "		FROM " & OverAllScoresTableName & " es1" 
sSQL = sSQL + "		JOIN " & SkiYearTableName & " sy ON sy.skiyearid = es1.skiyearid" 
sSQL = sSQL + "		JOIN " & DivisionsTableName & " dt ON dt.SkiYearid = sy.SkiYearid AND dt.div = es1.Div" 
sSQL = sSQL + "		JOIN " & MemberTableName & " mt ON mt.PersonID=RIGHT(es1.MemberID,8)"
		
' 		-- Members who have a Level 10 score in any division in the default Ski Year
sSQL = sSQL + "		LEFT JOIN" 
sSQL = sSQL + "		( SELECT st.MemberID, dt.div AS EliteDiv" 
sSQL = sSQL + "				FROM " & OverAllScoresTableName & " st" 
sSQL = sSQL + "				JOIN " & SkiYearTableName & " sy ON sy.skiyearid = st.skiyearid" 
sSQL = sSQL + "				JOIN " & DivisionsTableName & " dt ON dt.SkiYearid = sy.SkiYearid"
sSQL = sSQL + "					AND dt.div = CASE WHEN st.Div in ('M3','M4','M5','M6','M7','MM') THEN 'SM'" 
sSQL = sSQL + "                              WHEN st.Div in ('W3','W4','W5','W6','W7','MW') THEN 'SW'"
sSQL = sSQL + "                              WHEN st.Div in ('B2','B3','M1','M2','MM') THEN 'EM'" 
sSQL = sSQL + "                              WHEN st.Div in ('G2','G3','W1','W2','MW') THEN 'EW'" 
sSQL = sSQL + "                              ELSE 'PP' END"	   
sSQL = sSQL + "						WHERE sy.defaultyear = 1" 
sSQL = sSQL + "							AND st.TotalOverall>=dt.lvl_10_o"
sSQL = sSQL + "				GROUP BY st.MemberID, dt.div"
sSQL = sSQL + "		) es2"
sSQL = sSQL + "		ON es2.MemberID=es1.MemberID AND es2.EliteDiv=es1.Div"

' 		-- Suppress those who have and OM/OW or MM/MW scores in processing year
sSQL = sSQL + "		LEFT JOIN"
sSQL = sSQL + "		( SELECT st.MemberID, st.Div" 
sSQL = sSQL + "				FROM " & OverAllScoresTableName & " st" 
sSQL = sSQL + "				JOIN " & SkiYearTableName & " sy ON sy.skiyearid = st.skiyearid" 
sSQL = sSQL + "				JOIN " & DivisionsTableName & " dt ON dt.SkiYearid = sy.SkiYearid AND dt.div = st.div"

sSQL = sSQL + "						WHERE st.SkiYearID="& sProcessingYear	
sSQL = sSQL + "						AND st.div IN ('OM','OW','MM','MW')"  
sSQL = sSQL + "				GROUP BY st.memberid, st.Div ) es3"
sSQL = sSQL + "		ON es3.MemberID=es1.MemberID" 
sSQL = sSQL + "					AND es3.Div=CASE WHEN es1.Div='SM' THEN 'MM'" 
sSQL = sSQL + "													WHEN es1.Div='SW' THEN 'MW'"
sSQL = sSQL + "													WHEN es1.Div='EM' THEN 'OM'" 
sSQL = sSQL + "													WHEN es1.Div='EW' THEN 'OW'" 
sSQL = sSQL + "													ELSE 'PP' END"	
		
sSQL = sSQL + "				WHERE sy.SkiYearID="& sProcessingYear
sSQL = sSQL + "					AND es1.div IN ('EM','EW','SM','SW')"
sSQL = sSQL + "					AND mt.federationcode = 'USA'"
sSQL = sSQL + "					AND sy.skiyear-datepart(yyyy,mt.birthdate)-1 >= 18"
sSQL = sSQL + "					AND es2.MemberID IS NOT NULL" 
sSQL = sSQL + "					AND es3.MemberID IS NULL"	
sSQL = sSQL + "	ORDER BY es1.MemberID"



WriteDebugSQL("Adds OVERALL skiers (Temp) - " & sSQL)

Con.Execute(sSQL)






' -- Bottom of IF to exclude the Overall Scores section above from running
END IF









' -------------------------------------------------------------------------------------------
' -------------------------------------------------------------------------------------------
' ----------------------  SECTION 3:  Calculations of OVERALL RANKINGS  ---------------------
' -------------------------------------------------------------------------------------------
' -------------------------------------------------------------------------------------------

IF RunOvrllRanks = "YES" OR RunOverride = "YES" THEN

IF Request("Equival") <> "" THEN 
	Response.write(" OA Ranks ...")
	Response.Flush
END IF





' -- Step 1) OVERALL - Update the Equiv_Level10_Dates table which retains the First Instance date for each SkiYear/Member/Event/EliteDiv
' ------------------------------------------------------------------------------------------------------------------------------------------

' 		Added 3/13/2018 by Mark Crone: 
'     All qualifiers get added to the Equiv_Level10_Dates table.  But, the [new] suppression logic (next step) may include certain members  
'          in Age rankings (next steo) if they are under 18 or First Instance was after July 1st.
'     Used in suppression logic in next step


sSQL = " INSERT INTO " & EquivLevel10TableName
sSQL = sSQL & " (skiyearid, MemberID, Event, Div, First_Instance, Sent_Notice)" 

sSQL = sSQL & " SELECT L10.*"
sSQL = sSQL & " 		FROM"

sSQL = sSQL & "   ( SELECT MAX(st.skiyearid) AS skiyearid, st.memberid AS L10_MemberID, 'O' AS Event, st.div, MIN(TDateS) AS First_Instance, 'N' AS Sent_Notice"
sSQL = sSQL & " 				FROM " &OverAllScoresTableName& " st" 
sSQL = sSQL & " 				JOIN " & SkiYearTableName & " sy ON sy.skiyearid = st.skiyearid"
sSQL = sSQL & " 				JOIN  " & DivisionsTableName & "  dt ON dt.div = st.div AND dt.SkiYearid = sy.SkiYearid" 
sSQL = sSQL & " 				JOIN  " & SanctionTableName & "  ts ON ts.TournAppID=LEFT(TourID,6)"
sSQL = sSQL & " 				WHERE sy.defaultyear = 1" 
sSQL = sSQL & " 					AND st.div in ('EM','EW','SM','SW')" 
sSQL = sSQL & " 					AND st.TotalOverall >= dt.lvl_10_O"
sSQL = sSQL & " 				GROUP BY st.memberid, st.div ) L10"

sSQL = sSQL & " 		LEFT JOIN"
sSQL = sSQL & " 			( SELECT ldt.skiyearid, MemberID, Event, Div, First_Instance"
sSQL = sSQL & " 					FROM " & EquivLevel10TableName & " ldt"
sSQL = sSQL & " 					JOIN " & SkiYearTableName & " sy ON sy.skiyearid = ldt.skiyearid" 
sSQL = sSQL & " 						WHERE sy.defaultyear = 1 AND Event='O' ) pr"
sSQL = sSQL & " 		ON pr.SkiYearID=L10.SkiYearID AND pr.MemberID=L10.L10_MemberID AND pr.Event=L10.Event AND pr.Div=L10.Div"									

sSQL = sSQL & " 		WHERE pr.MemberID IS NULL"	

WriteDebugSQL("Update the Equiv_Level10_Dates (Temp) - " & sSQL)

Con.Execute(sSQL)





' -- Delete all of the old records which match the Ski Year ID that we are processing.

sSQL = "DELETE FROM " & RankTableName & " WHERE Event = 'O' and SkiYearID = " & sProcessingYear

WriteDebugSQL(sSQL)

Con.Execute(sSQL)







' Query with Max function and Group By clause, to roll to best single round from each TourID
' Revised Feb 2008 such that Score string is now char(10), with leftmost 8 being the
' actual Total Overall Score (formatted as 8,2), followed by 2-char prioritized class.

sSQL = "SELECT OA.MemberID, OA.TourID, OA.Div, Max("
sSQL = sSQL & " Substring(Cast(Cast(OA.TotalOverAll+400000 as Decimal(8,2)) as Char(9)),2,8)"
sSQL = sSQL & " + OA.PrioClass) as MaxOAScore, Max(DivType) as DivType,"
sSQL = sSQL & " Max(OA.DivOrig) as DivOrig, Max(OA.SkiYearID) as SkiYearID,"
sSQL = sSQL & " Coalesce(Max(ST.TName),'(Tournament Unknown)') as TName, Max(OA.Team) as Team"

sSQL = sSQL & " FROM " & OverAllScoresTableName & " OA" 
sSQL = sSQL & " LEFT JOIN " & SanctionTableName & " ST on Left(ST.TournAppID,6) =  Left(OA.TourID,6)"


' -- First of two Level 10 suppression code added 3/11/2018 by Mark Crone --

Dim AddOverallSuppression
AddOverallSuppression="Y"
IF AddOverallSuppression="Y" THEN 

sSQL = sSQL & " LEFT JOIN "	
sSQL = sSQL & " ( SELECT es1.MemberID, es1.Div AS Div_Orig"
	
sSQL = sSQL & "        FROM " & OverAllScoresTableName & " es1"

sSQL = sSQL & " 	      LEFT JOIN"	
sSQL = sSQL & "          ( SELECT st.memberid AS L10_MemberID, st.Div, MAX(TotalOverAll) AS Max_Score"
sSQL = sSQL & "              FROM " & OverAllScoresTableName & " st"
sSQL = sSQL & "              JOIN " & SkiYearTableName & " sy ON sy.skiyearid = st.skiyearid"
sSQL = sSQL & "              JOIN " & DivisionsTableName & " dt ON dt.div = st.div AND dt.SkiYearid = sy.SkiYearid"
sSQL = sSQL & "                 WHERE sy.defaultyear = 1" 
sSQL = sSQL & "                    AND st.div in ('EM')"
sSQL = sSQL & "                    AND st.TotalOverAll >= dt.lvl_10_o"
sSQL = sSQL & "              GROUP BY st.memberid, st.div"
sSQL = sSQL & "        ) L10"   
sSQL = sSQL & " 		    ON L10.L10_MemberID=es1.MemberID"
sSQL = sSQL & "              AND L10.Div=CASE WHEN es1.Div in ('M3','M4','M5','M6','M7','MM') THEN 'SM'" 
sSQL = sSQL & "                               WHEN es1.Div in ('W3','W4','W5','W6','W7','MW') THEN 'SW'"
sSQL = sSQL & "                               WHEN es1.Div in ('B2','B3','M1','M2','MM') THEN 'EM'" 
sSQL = sSQL & "                               WHEN es1.Div in ('G2','G3','W1','W2','MW') THEN 'EW'" 
sSQL = sSQL & "                               ELSE 'PP' END"

sSQL = sSQL & "        JOIN " & MemberTableName & " mt ON RIGHT(es1.memberid,8) = mt.personid"
sSQL = sSQL & "        JOIN " & SkiYearTableName & " sy ON sy.skiyearid = es1.skiyearid"
sSQL = sSQL & "        JOIN " & DivisionsTableName & " dt ON dt.div = es1.div AND dt.skiyearid = es1.skiyearid"

Dim mtemp
mtemp=1
IF mtemp=1 THEN 
'                     -- Eliminates from Suppression any member whose First_Instance is after Level10_EndDate of SkiYear 
sSQL = sSQL + " 			LEFT JOIN"
sSQL = sSQL + " 					( SELECT MemberID, Div, First_Instance"
sSQL = sSQL + " 								FROM " & EquivLevel10TableName & " l"
sSQL = sSQL + " 								JOIN " & SkiYearTableName & " sy ON sy.skiyearid = l.skiyearid"
sSQL = sSQL + " 									WHERE sy.defaultyear = 1 AND Event='O'" 
sSQL = sSQL + " 										AND First_Instance >= Level10_EndDate  ) ltd"

sSQL = sSQL + " 			ON es1.MemberID=ltd.MemberID"	
sSQL = sSQL + " 						AND ltd.Div=CASE WHEN es1.Div in ('M3','M4','M5','M6','M7','MM') THEN 'SM'" 
sSQL = sSQL + " 														WHEN es1.Div in ('W3','W4','W5','W6','W7','MW') THEN 'SW'"
sSQL = sSQL + " 														WHEN es1.Div in ('B2','B3','M1','M2','MM') THEN 'EM'" 
sSQL = sSQL + " 														WHEN es1.Div in ('G2','G3','W1','W2','MW') THEN 'EW'" 
sSQL = sSQL + " 														ELSE 'PP' END"
		
END IF 	' -- Bottom of mtemp		

		
sSQL = sSQL & "           WHERE mt.federationcode = 'USA'"
sSQL = sSQL & "              AND sy.defaultyear = 1"
sSQL = sSQL & "              AND sy.skiyear-datepart(yyyy,mt.birthdate)-1 >= 18"	
sSQL = sSQL & "              AND es1.div NOT IN ('EM','EW','SM','SW')"
sSQL = sSQL & "              AND es1.div NOT IN ('OM','OW','MM','MW')"
sSQL = sSQL & "              AND L10.L10_MemberID IS NOT NULL"


IF mtemp=1 THEN sSQL = sSQL + " 						AND ltd.MemberID IS NULL"

sSQL = sSQL & "  	  GROUP BY es1.memberid, es1.div" 

sSQL = sSQL + " ) L10Supp"
sSQL = sSQL + " ON L10Supp.MemberID=OA.MemberID AND L10Supp.Div_Orig=OA.Div"

END IF				' -- Bottom of conditional suppression


sSQL = sSQL & " WHERE OA.SkiYearID = " & sProcessingYear & " AND OA.TotalOverall > 0"

' -- Logic - If member is in suppression then do not include - Mark Crone 3/11/2018 --
IF AddOverallSuppression="Y" THEN sSQL = sSQL + " AND L10Supp.MemberID IS NULL"


sSQL = sSQL & " GROUP BY OA.MemberID, OA.TourID, OA.Div"
sSQL = sSQL & " ORDER BY OA.MemberID, OA.Div, MaxOAScore Desc"



WriteDebugSQL("Line 2480 " &sSQL)
rs.open sSQL, sConnectionToTRATable, 3, 3

IF NOT rs.eof THEN
	tempvar = rs.getrows()
	nTotalMembers = ubound(tempvar,2)
	rs.MoveFirst
ELSE 
	nTotalMembers = 1
END IF

nProcessedSoFar = 0



' ------------      Outer loop of all OVERALL SCORES    --------------

DO UNTIL rs.eof

  nProcessedSoFar = nProcessedSoFar + 1

  IF Request("Equival") <> "" THEN 
		IF (nProcessedSoFar mod 10 = 5) THEN ShowProgress (nProcessedSoFar / nTotalMembers)
	END IF

  TempMemberID = rs("MemberID"): TempTeam = trim(rs("Team")): TempDiv = rs("Div"): TempDivType = rs("DivType")
  nScoC = 0: nScoR = 0: RSco1 = 0: RSco2 = 0: RSco3 = 0: TotScore = 0
  
  
  ' Inner Loop of overall scores for this MemberID/Division
  ' Jun 2008 -- now implements the "Do No Harm" philosophy, by evaluating 
  ' All possible Ranking scores using 1, 2 or 3 best scores, factoring in
  ' the applicable penalty levels for each possibility.  Uses separate 
  ' counters for C and ELR scores and the new penalty matrix (Feb 2008).
 
  DO WHILE TempMemberID = rs("MemberID") AND TempDiv = rs("Div")

		IF TempTeam < trim(rs("Team")) THEN TempTeam = trim(rs("Team"))
		IF TempDivType > rs("DivType") THEN TempDivType = rs("DivType")
		TempScore = left(rs("MaxOAScore"),8)

     IF nScoC+nScoR < 3 THEN     

        TotScore = TotScore + TempScore
        IF Mid(rs("MaxOAScore"),9,1) >= "3" THEN nScoR = nScoR + 1: ELSE nScoC = nScoC + 1
        
        IF nScoC+nScoR = 1 THEN
           RExp1 = FormatNumber(TempScore,1) & " (" & Right(rs("MaxOAScore"),1) & ")"
           IF rs("Div")<>rs("DivOrig") THEN RExp1 = RExp1 & " as " & rs("DivOrig")
           RExp1 = RExp1 & " from " & SQLClean(rs("TName")) & "&#13;"
           RSco1 = (1-(tPenalty(nScoC,nScoR)/100))*(TotScore)/(nScoC+nScoR)
           IF tPenalty(nScoC,nScoR) > 0 THEN RPen1 = tPenalty(nScoC,nScoR) & "% Penalty": ELSE RPen1 = "NO Penalty"

        ELSEIF nScoC+nScoR = 2 THEN
           RExp2 = RExp1 & FormatNumber(TempScore,1) & " (" & Right(rs("MaxOAScore"),1) & ")"
           IF rs("Div")<>rs("DivOrig") THEN RExp2 = RExp2 & " as " & rs("DivOrig")
           RExp2 = RExp2 & " from " & SQLClean(rs("TName")) & "&#13;"
           RSco2 = (1-(tPenalty(nScoC,nScoR)/100))*(TotScore)/(nScoC+nScoR)
           IF tPenalty(nScoC,nScoR) > 0 THEN RPen2 = tPenalty(nScoC,nScoR) & "% Penalty": ELSE RPen2 = "NO Penalty"

        ELSE
           RExp3 = RExp2 & FormatNumber(TempScore,1) & " (" & Right(rs("MaxOAScore"),1) & ")"
           IF rs("Div")<>rs("DivOrig") THEN RExp3 = RExp3 & " as " & rs("DivOrig")
           RExp3 = RExp3 & " from " & SQLClean(rs("TName")) & "&#13;"
           RSco3 = (1-(tPenalty(nScoC,nScoR)/100))*(TotScore)/(nScoC+nScoR)
           IF tPenalty(nScoC,nScoR) > 0 THEN RPen3 = tPenalty(nScoC,nScoR) & "% Penalty": ELSE RPen3 = "NO Penalty"

        END IF

     END IF

     rs.moveNEXT
     IF rs.eof THEN exit do
  LOOP
  ' -- Bottom of INNER loop	


  ' ************** Rule 1.13 Arbitration of Overall Ranking Score and Penalty HERE ************

  IF RSco2 >= RSco1 THEN
     RSco1 = RSco2: RExp1 = RExp2: RPen1 = RPen2
  ELSE
     IF nScoC+nScoR = 2 THEN RPen1 = RPen1 & "&#13;Rule 1.13 Applied; see FAQ/Tips"
  END IF

  IF RSco3 >= RSco1 THEN
     RSco1 = RSco3: RExp1 = RExp3: RPen1 = RPen3
  ELSE
     IF nScoC+nScoR > 2 THEN RPen1 = RPen1 & "&#13;Rule 1.13 Applied; see FAQ/Tips"
  END IF

  ' Finally insert this computed Overall Ranking Score into the RankScore table.

  sSQL = "INSERT INTO " & RankTableName & " (MemberID, Team, Event, Div, DivType, RankScore, RnkScoBkup, SkiYearID)"
  sSQL = sSQL & " VALUES ('" & TempMemberID & "','" & TempTeam & "', 'O', '" & TempDiv & "', '" & TempDivType
  sSQL = sSQL & "', '" & RSco1 & "', '" & RExp1 & "With " & RPen1 & "', " & sProcessingYear & ")"

  Con.Execute(sSQL)

LOOP
' -- Bottom of OUTER Loop

rs.close



' -----------------------------------------------------------------------------------
' ----------------------   End of OVERALL RANKING Calculations  ---------------------
' -----------------------------------------------------------------------------------

' Bottom of IF to exclude the Overall Rankings section above from running
END IF





' -----------------------------------------------------------------------------------------------
' -- SECTION: Temporary for Change in Age Division M4-W4 mas age from 52 to 54 --
' -- First step adds M5 Scores not already in M4 to the M4 Equival scores --
' -- Second step deletes the M5 scores from Equival Scores thereby omitting them from rankings
' -----------------------------------------------------------------------------------------------

IF sProcessingYear=1 THEN 
 
sSQL = "INSERT INTO usawsrank.EquivScores (MemberID, TourID, Div, Event, Round, Class, ScoreOrig, DivOrig, Score, Place, Rating, FmtScore, OAScore, SkiYearID, Team, PrioClass, TeamStat, DivType)"


sSQL = sSQL + " SELECT es.MemberID, es.TourID, CASE WHEN es.Div='M5' THEN 'M4'" 
sSQL = sSQL + "					WHEN es.Div='W5' THEN 'W4' END AS Div" 
sSQL = sSQL + "			, es.Event, es.Round, es.Class"
sSQL = sSQL + "			, es.ScoreOrig, es.DivOrig, es.Score, es.Place, es.Rating, es.FmtScore, es.OAScore"  
sSQL = sSQL + "			, es.SkiYearID, es.Team, es.PrioClass, es.TeamStat, es.DivType"
sSQL = sSQL + "	FROM usawsrank.EquivScores es"
	
sSQL = sSQL + "	JOIN" 
sSQL = sSQL + "		( SELECT MemberID, TourID, s.Div, Event, Round, Class"
sSQL = sSQL + "			, ScoreOrig, DivOrig, Score, Place, Rating, FmtScore, OAScore"  
sSQL = sSQL + "			, sy.SkiYear, Team, PrioClass, TeamStat, DivType"
sSQL = sSQL + "			, FirstName, LastName, sy.skiyear-datepart(yyyy,mt.birthdate)-1 AS  Age_RuleAtTime"
sSQL = sSQL + " 					, up_age, low_age, s.SkiYearID" 
 					
sSQL = sSQL + "			FROM usawsrank.EquivScores s"
sSQL = sSQL + "			LEFT JOIN usawsrank.skiyear sy ON sy.SkiYearID=s.SkiYearID"
sSQL = sSQL + "			LEFT JOIN usawsrank.Division d ON d.SkiYearID=sy.SkiYearID AND d.div=s.div"
sSQL = sSQL + "			LEFT JOIN usawaterski.dbo.membershort mt ON personid=CAST(RIGHT(memberID,8) AS INT)"
sSQL = sSQL + "				WHERE s.div IN ('M5','W5')" 
sSQL = sSQL + "					AND s.SkiYearID IN (1)"
sSQL = sSQL + "					AND sy.skiyear-datepart(yyyy,mt.birthdate)-1 BETWEEN 53 AND 54"
sSQL = sSQL + "	) A"
sSQL = sSQL + "	ON es.MemberID=A.MemberID AND es.TourID=A.TourID AND es.Event=A.Event AND es.Div=A.Div"
sSQL = sSQL + "		AND es.SkiYearID=A.SkiYearID AND es.Round=A.Round"
		 
sSQL = sSQL + "	LEFT JOIN"
sSQL = sSQL + "		( SELECT MemberID, Event, Round, TourID, SkiYearID"
sSQL = sSQL + "			, CASE WHEN Div='M4' THEN 'M5'" 
sSQL = sSQL + "					WHEN Div='W4' THEN 'W5' END AS NewDiv"
sSQL = sSQL + "					FROM usawsrank.EquivScores s"
sSQL = sSQL + "						WHERE Div IN ('M4','W4')" 
sSQL = sSQL + "							AND SkiYearID IN (1)"
sSQL = sSQL + "		) om"
sSQL = sSQL + "	ON om.MemberID=es.MemberID AND om.TourID=es.TourID AND om.Event=es.Event AND om.NewDiv=es.Div"
sSQL = sSQL + "		AND om.SkiYearID=es.SkiYearID AND om.Round=es.Round"
	
sSQL = sSQL + "	WHERE om.MemberID IS NULL"	

WriteDebugSQL("Line 2649")
WriteDebugSQL(sSQL)
Con.Execute(sSQL)


sSQL = " DELETE s"
sSQL = sSQL + "	FROM usawsrank.EquivScores s"
sSQL = sSQL + "				LEFT JOIN usawsrank.skiyear sy ON sy.SkiYearID=s.SkiYearID"
sSQL = sSQL + "				LEFT JOIN usawsrank.Division d ON d.SkiYearID=sy.SkiYearID AND d.div=s.div"
sSQL = sSQL + "				LEFT JOIN usawaterski.dbo.membershort mt ON personid=CAST(RIGHT(memberID,8) AS INT)"
sSQL = sSQL + "					WHERE s.div IN ('M5','W5')" 
sSQL = sSQL + "						AND s.SkiYearID IN (1)"
sSQL = sSQL + "						AND sy.skiyear-datepart(yyyy,mt.birthdate)-1 BETWEEN 53 AND 54"

WriteDebugSQL("Line 2661")
WriteDebugSQL(sSQL)
Con.Execute(sSQL)


END IF













' --------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------
' --------  SECTION 4:  Begin New Streamlined Event Ranking Logic, Using EquivScores Table  --------
' --------------------------------------------------------------------------------------------------
' --------------------------------------------------------------------------------------------------

' NOTES:
' Modified May 2016 to add NOPS score value to bottom of "RankScoBkup" Popup Explanation DJC


IF RunEventRanks = "YES" OR RunOverride = "YES" THEN

IF Request("Equival") <> "" THEN 
	Response.write(" Ev Ranks ...")
	Response.Flush
END IF






' -- Step 1) Update the Equiv_Level10_Dates table which retains the First Instance date for each SkiYear/Member/Event/EliteDiv
' ---------------------------------------------------------------------------------------------------------------------------------

' 		Added 3/13/2018 by Mark Crone: 
'     All qualifiers get added to the Equiv_Level10_Dates table.  But, the [new] suppression logic (next step) may include certain members  
'          in Age rankings (next steo) if they are under 18 or First Instance was after July 1st.
'     Used in suppression logic in next step
 
' -- First for the Events -- 
sSQL = "INSERT INTO " & EquivLevel10TableName
sSQL = sSQL & " (skiyearid, MemberID, Event, Div, First_Instance, Sent_Notice)" 

sSQL = sSQL & " SELECT L10.*"
sSQL = sSQL & " 		FROM"

sSQL = sSQL & "   ( SELECT MAX(st.skiyearid) AS skiyearid, st.memberid AS L10_MemberID, st.event, st.div, MIN(TDateS) AS First_Instance, 'N' AS Sent_Notice" 
sSQL = sSQL & " 				FROM " & EquivScoresTableName & " st "
sSQL = sSQL & " 				JOIN " & SkiYearTableName & " sy ON sy.skiyearid = st.skiyearid"
sSQL = sSQL & " 				JOIN " & DivisionsTableName & " dt ON dt.div = st.div AND dt.SkiYearid = sy.SkiYearid" 
sSQL = sSQL & " 				JOIN " & SanctionTableName & " ts ON ts.TournAppID=LEFT(TourID,6)"
sSQL = sSQL & " 				WHERE sy.defaultyear = 1" 
sSQL = sSQL & " 					AND st.div in ('EM','EW','SM','SW')" 
sSQL = sSQL & " 					AND ( (st.event = 'S' AND st.score >= dt.lvl_10_s)" 
sSQL = sSQL & " 						OR (st.event = 'T' AND st.score >= dt.lvl_10_t)" 
sSQL = sSQL & " 						OR (st.event = 'J' AND st.score >= dt.lvl_10_j) )" 
sSQL = sSQL & " 				GROUP BY st.memberid, st.event, st.div ) L10"

sSQL = sSQL & " 		LEFT JOIN"
sSQL = sSQL & " 			( SELECT ldt.skiyearid, MemberID, Event, Div, First_Instance"
sSQL = sSQL & " 					FROM " & EquivLevel10TableName & " ldt"
sSQL = sSQL & " 					JOIN " & SkiYearTableName & " sy ON sy.skiyearid = ldt.skiyearid" 
sSQL = sSQL & " 						WHERE sy.defaultyear = 1 ) pr"
sSQL = sSQL & " 		ON pr.SkiYearID=L10.SkiYearID AND pr.MemberID=L10.L10_MemberID AND pr.Event=L10.Event AND pr.Div=L10.Div"									

sSQL = sSQL & " 		WHERE pr.MemberID IS NULL"		
	
' WriteDebugSQL("LINE 2589: " & sSQL)

Con.Execute(sSQL)





' -- Delete all of the old Event Rank records for the Ski Year ID that we are processing.

sSQL = "DELETE FROM " & RankTableName & " WHERE Event <> 'O' and SkiYearID = " & sProcessingYear

WriteDebugSQL(sSQL)

Con.Execute(sSQL)



' -- Step 2) Roll to best single round from each TourID
' Note - Query with Max function and Group By clause
'      - Revised Feb 2008 such that Score string is now char(10), with leftmost 8 being the
'            actual Total Overall Score (formatted as 8,2), followed by 2-char prioritized class.
' --------------------------------------------------------------------------------------------------

sSQL = "SELECT ES.MemberID, ES.TourID, ES.Div, ES.Event, Max("
sSQL = sSQL & " Substring(Cast(Cast(ES.Score+400000 as Decimal(8,2)) as Char(9)),2,8)"
sSQL = sSQL & " + ES.PrioClass) as MaxScore, Max(ES.ScoreOrig) as ScoreOrig,"
sSQL = sSQL & " Max(coalesce(ES.OAScore,0.0)) as OAScore,"
sSQL = sSQL & " Max(ES.DivOrig) as DivOrig, Max(ES.Place) as Place, Max(ES.Rating) as Rating,"
sSQL = sSQL & " Max(ES.SkiYearID) as SkiYearID, Coalesce(Max(ST.TName),'(Tournament Unknown)') as TName,"
sSQL = sSQL & " Max(ES.Team) as Team, Max(ES.TeamStat) as TeamStat, Max(DivType) as DivType,"
sSQL = sSQL & " Max(DT.NationalRec_S) as SlmNR, Max(DT.OverPtsBy_S) as SlmOP, Max(DT.OverExp_S) as SlmExp,"
sSQL = sSQL & " Max(DT.NationalRec_T) as TrkNR, Max(DT.OverExp_T) as TrkExp,"
sSQL = sSQL & " Max(DT.NationalRec_J) as JmpNR, Max(DT.OverExp_J) as JmpExp"

sSQL = sSQL & " FROM " & EquivScoresTableName & " ES "
sSQL = sSQL & " LEFT JOIN " & SanctionTableName & " ST ON Left(ES.TourID,6) = Left(ST.TournAppID,6) "
sSQL = sSQL & " JOIN " & DivisionsTableName & " DT ON ES.Div = DT.Div and DT.SkiYearID = " & sProcessingYear



' -- New Level 10 suppression code added 3/11/2018 by Mark Crone --
' --------------------------------------------------------------------------------------------------


Dim AddEventSuppression
AddEventSuppression="Y"
IF AddEventSuppression="Y" THEN 
	
	
' -- Assumption: Previous sections have synthesized EquivScores into an Elite division if the skier has competed 1x or more in an Elite div --
' --   1) Does not apply to members Under 18 
' --   2) Suppression is not applied to OM, OW, MM and Elite Divisions because we want those EquivScore scores sets to remain in the final rankings 

sSQL = sSQL + " LEFT JOIN " 
sSQL = sSQL + "   ( SELECT es1.MemberID, es1.Event, es1.Div AS Div_Orig" 
sSQL = sSQL + "       FROM " & EquivScoresTableName & " es1"

sSQL = sSQL + "       LEFT JOIN"	
sSQL = sSQL + "         ( SELECT st.memberid AS L10_MemberID, st.event, st.div, MAX(Score) AS Max_Score"
sSQL = sSQL + " 						FROM " & EquivScoresTableName & " st"
sSQL = sSQL + " 						JOIN " & SkiYearTableName & " sy ON sy.skiyearid = st.skiyearid"
sSQL = sSQL + " 						JOIN " & DivisionsTableName & " dt ON dt.div = st.div AND dt.SkiYearid = sy.SkiYearid"
sSQL = sSQL + " 								WHERE sy.defaultyear = 1" 
sSQL = sSQL + " 									AND st.div in ('EM','EW','SM','SW')"
sSQL = sSQL + " 									AND ( (st.event = 'S' AND st.score >= dt.lvl_10_s) OR"
sSQL = sSQL + " 												(st.event = 'T' AND st.score >= dt.lvl_10_t) OR"
sSQL = sSQL + " 												(st.event = 'J' AND st.score >= dt.lvl_10_j) )"
sSQL = sSQL + " 						GROUP BY st.memberid, st.event, st.div"
sSQL = sSQL + " 			) L10"   
sSQL = sSQL + " 			ON L10.L10_MemberID=es1.MemberID AND L10.event=es1.event"
sSQL = sSQL + " 					AND L10.Div=CASE WHEN es1.Div in ('M3','M4','M5','M6','M7','MM') THEN 'SM'" 
sSQL = sSQL + " 														WHEN es1.Div in ('W3','W4','W5','W6','W7','MW') THEN 'SW'"
sSQL = sSQL + " 														WHEN es1.Div in ('B2','B3','M1','M2','MM') THEN 'EM'" 
sSQL = sSQL + " 														WHEN es1.Div in ('G2','G3','W1','W2','MW') THEN 'EW'" 
sSQL = sSQL + " 														ELSE 'PP' END"

sSQL = sSQL + " 			JOIN " & MemberTableName & " mt ON es1.memberid = mt.personidwithcheckdigit"
sSQL = sSQL + " 			JOIN " & SkiYearTableName & " sy ON sy.skiyearid = es1.skiyearid"

'                     -- Eliminates from Suppression any member whose First_Instance is after Level10_EndDate of SkiYear 
sSQL = sSQL + " 			LEFT JOIN"
sSQL = sSQL + " 					( SELECT MemberID, Event, Div, First_Instance"
sSQL = sSQL + " 								FROM " & EquivLevel10TableName & " l"
sSQL = sSQL + " 								JOIN " & SkiYearTableName & " sy ON sy.skiyearid = l.skiyearid"
sSQL = sSQL + " 									WHERE sy.defaultyear = 1" 


sSQL = sSQL + " 										AND First_Instance >= Level10_EndDate  ) ltd"
' sSQL = sSQL + " 										AND First_Instance >=  CAST('7/1/' + RIGHT(SkiYear,2) AS Date)  ) ltd"
sSQL = sSQL + " 			ON es1.MemberID=ltd.MemberID AND es1.Event=ltd.Event"	
sSQL = sSQL + " 						AND ltd.Div=CASE WHEN es1.Div in ('M3','M4','M5','M6','M7','MM') THEN 'SM'" 
sSQL = sSQL + " 														WHEN es1.Div in ('W3','W4','W5','W6','W7','MW') THEN 'SW'"
sSQL = sSQL + " 														WHEN es1.Div in ('B2','B3','M1','M2','MM') THEN 'EM'" 
sSQL = sSQL + " 														WHEN es1.Div in ('G2','G3','W1','W2','MW') THEN 'EW'" 
sSQL = sSQL + " 														ELSE 'PP' END"

		
sSQL = sSQL + " 				WHERE mt.federationcode = 'USA'"
sSQL = sSQL + " 						AND sy.skiyear-datepart(yyyy,mt.birthdate)-1 >= 18"	
sSQL = sSQL + " 						AND es1.SkiYearID = "& sProcessingYear 
sSQL = sSQL + " 						AND es1.div NOT IN ('EM','EW','SM','SW')"
sSQL = sSQL + " 						AND es1.div NOT IN ('OM','OW','MM','MW')"
sSQL = sSQL + " 						AND L10.L10_MemberID IS NOT NULL"
sSQL = sSQL + " 						AND ltd.MemberID IS NULL"


sSQL = sSQL + " GROUP BY es1.memberid, es1.event, es1.div" 
sSQL = sSQL + " ) L10Supp"
sSQL = sSQL + " ON L10Supp.MemberID=ES.MemberID AND L10Supp.Div_Orig=ES.Div AND L10Supp.Event=ES.Event"

END IF				' -- Bottom of conditional suppression


sSQL = sSQL & " WHERE ES.SkiYearID = " & sProcessingYear & " and DT.OverPtsBy_S > 0"

' -- Logic - If member is in suppression then do not include - Mark Crone 3/11/2018 --
IF AddEventSuppression="Y" THEN sSQL = sSQL + " AND COALESCE(L10Supp.MemberID,0)=0 "


sSQL = sSQL & " GROUP BY ES.MemberID, ES.TourID, ES.Div, ES.Event, L10Supp.MemberID"
sSQL = sSQL & " ORDER BY ES.MemberID, ES.Div, ES.Event, MaxScore Desc"


WriteDebugSQL(sSQL)

' response.write("<br>Line 2272: "&sSQL)
' response.end


rs.open sSQL, sConnectionToTRATable, 3, 3

IF not rs.eof THEN

tempvar = rs.getrows()
nTotalMembers = ubound(tempvar,2)

rs.MoveFirst
nProcessedSoFar = 0


Dim EqScr1, EqScr2, EqScr3


' ----------------      Outer loop of all Equivalent scores     ------------------
' --------------------------------------------------------------------------------

DO UNTIL rs.eof

  nProcessedSoFar = nProcessedSoFar + 1

	IF Request("Equival") <> "" THEN 
	  IF (nProcessedSoFar mod 10 = 5) THEN ShowProgress (nProcessedSoFar / nTotalMembers)
	END IF

  TempMemberID = rs("MemberID"): TempDiv = rs("Div"): TempDivType = rs("DivType")
  TempEvent = rs("Event"): TempTeam = trim(rs("Team")): TempTeamStat = rs("TeamStat")
  
  nScoC = 0: nScoR = 0: RSco1 = 0: RSco2 = 0: RSco3 = 0: TotScore = 0
  R_Ski = "": R_PLC = "": N_PLC = "": RExp1 = "": RMaxRat = "  "
  EqScr1 = 0
  EqScr2 = 0
  EqScr3 = 0
  
  ' INNER Loop of Scores for this MemberID/Division/Event

  ' First Phase is to Add top 3 (or 2 for NCWSA) to Ranking score total

  ' WriteDebugSQL("Begin Ranking Processing " & TempMemberID & " / " & TempDiv & " / " & TempEvent)
  
  DO WHILE TempMemberID = rs("MemberID") AND TempDiv = rs("Div") AND TempEvent = rs("Event")

     IF rs("Rating") > RMaxRat THEN RMaxRat = rs("Rating")
     IF TempTeam < trim(rs("Team")) THEN TempTeam = trim(rs("Team"))
     IF TempDivType > rs("DivType") THEN TempDivType = rs("DivType")
     IF rs("TeamStat") = "A" or TempTeamStat <> "A" then TempTeamStat = rs("TeamStat")
     TempScore = left(rs("MaxScore"),8)
     IF TempEvent = "S" THEN FmtSco = FormatNumber(rs("ScoreOrig"),2): ELSE FmtSco = FormatNumber(rs("ScoreOrig"),0)
     TempSlmNR = rs("SlmNR"): TempSlmOP = rs("SlmOP"): TempSlmExp = rs("SlmExp")
     TempTrkNR = rs("TrkNR"): TempTrkExp = rs("TrkExp")
     TempJmpNR = rs("JmpNR"): TempJmpExp = rs("JmpExp")
     IF left(TempDiv,1) = "C" THEN
     	
        ' NCWSA Logic beginning 2013 ski year -- Best 1 only.

        IF nScoR < 1 THEN     
           nScoR = nScoR + 1
           RSco1 = RSco1 + TempScore
           RExp1 = RExp1 & FmtSco & " from " & SQLClean(rs("TName")) & "&#13;"
        END IF

     ELSE

        ' AWSA Logic -- now Differential Penalty function of nScoC and nScoR

        IF nScoC+nScoR < 3 THEN     

           TotScore = TotScore + TempScore
           IF Mid(rs("MaxScore"),9,1) >= "3" THEN nScoR = nScoR + 1: ELSE nScoC = nScoC + 1
        
           IF nScoC+nScoR = 1 THEN
              RExp1 = FmtSco & " (" & Right(rs("MaxScore"),1) & ")"
              EqScr1 = rs("ScoreOrig")
              IF rs("Div")<>rs("DivOrig") THEN RExp1 = RExp1 & " as " & rs("DivOrig")
              RExp1 = RExp1 & " from " & SQLClean(rs("TName")) & "&#13;"
              RSco1 = (1-(tPenalty(nScoC,nScoR)/100))*(TotScore)/(nScoC+nScoR)
              IF tPenalty(nScoC,nScoR) > 0 THEN RPen1 = tPenalty(nScoC,nScoR) & "% Penalty": ELSE RPen1 = "NO Penalty"

           ELSEIF nScoC+nScoR = 2 THEN
              RExp2 = RExp1 & FmtSco & " (" & Right(rs("MaxScore"),1) & ")"
              EqScr2 = rs("ScoreOrig")
              IF rs("Div")<>rs("DivOrig") THEN RExp2 = RExp2 & " as " & rs("DivOrig")
              RExp2 = RExp2 & " from " & SQLClean(rs("TName")) & "&#13;"
              RSco2 = (1-(tPenalty(nScoC,nScoR)/100))*(TotScore)/(nScoC+nScoR)
              IF tPenalty(nScoC,nScoR) > 0 THEN RPen2 = tPenalty(nScoC,nScoR) & "% Penalty": ELSE RPen2 = "NO Penalty"

           ELSE
              RExp3 = RExp2 & FmtSco & " (" & Right(rs("MaxScore"),1) & ")"
              EqScr3 = rs("ScoreOrig")
              IF rs("Div")<>rs("DivOrig") THEN RExp3 = RExp3 & " as " & rs("DivOrig")
              RExp3 = RExp3 & " from " & SQLClean(rs("TName")) & "&#13;"
              RSco3 = (1-(tPenalty(nScoC,nScoR)/100))*(TotScore)/(nScoC+nScoR)
              IF tPenalty(nScoC,nScoR) > 0 THEN RPen3 = tPenalty(nScoC,nScoR) & "% Penalty": ELSE RPen3 = "NO Penalty"

           END IF

        END IF

     END IF 

     ' Second Phase is to pick up Nationals and Regionals Placements ... BUT
     ' ONLY if the 0riginal Division and Equivalent Division are appropriate
     ' Meaning they are both the same, or that neither is OM/OW/MM/IM/IW etc
     
     IF (rs("DivOrig")=TempDiv) OR (right(rs("DivOrig"),1)<"C" AND right(TempDiv,1)<"C") THEN
	     IF UCASE(RIGHT(TRIM(rs("TourID")),1)) = "A" THEN N_Plc = rs("Place")
	     IF UCASE(RIGHT(TRIM(rs("TourID")),1)) = "B" THEN R_Ski = mid(rs("TourID"),3,1): R_Plc = rs("Place")
     END IF


   	rs.moveNEXT
   	IF rs.eof THEN exit do
  LOOP
  
  ' -------   Bottom of INNER loop -- Now Finalize Ranking For this Member/Div/Event



  ' WriteDebugSQL("Begin Ranking Finalization " & TempMemberID & " / " & TempDiv & " / " & TempEvent)

  ' NCWSA Formulation First for "C" Divs
 	
  ' ************** Calculate and Explain Penalty for NCWSA Event Ranking Scores HERE ************

  IF left(TempDiv,1) = "C" THEN

 	  RPen1 = "NO Penalty"
  
  ' AWSA Formulation Otherwise

  ' ************** Finalize AWSA Ranking Score and Penalty and compute associated NOPS value ************

  ELSE

     ' ********** First portion applies Rule 1.13 for best Ranking Score with Penalty **************

     IF RSco2 >= RSco1 THEN
        RSco1 = RSco2: RExp1 = RExp2: RPen1 = RPen2
     ELSE
        IF nScoC+nScoR = 2 THEN RPen1 = RPen1 & "&#13;Rule 1.13 Applied; see FAQ/Tips"
     END IF

     IF RSco3 >= RSco1 THEN
        RSco1 = RSco3: RExp1 = RExp3: RPen1 = RPen3
     ELSE
        IF nScoC+nScoR > 2 THEN RPen1 = RPen1 & "&#13;Rule 1.13 Applied; see FAQ/Tips"
     END IF

     ' *********** Second Portion computes and formats Associated NOPS value *************

     ' WriteDebugSQL("Begin Ranking NOPS Computation " & TempMemberID & " / " & TempDiv & " / " & TempEvent & " / " & FormatNumber(RSco1,2))
     
     IF TempEvent = "S" THEN
     
        IF RSco1 < 6 THEN 
        	 OASc = RSco1 * TempSlmOP
        ELSE
        	 OASc = (6 * TempSlmOP) + ( (1500 - (6 * TempSlmOP)) * (((RSco1 - 6) / (TempSlmNR - 6)) ^ TempSlmExp) )
        END IF

     ELSEIF TempEvent = "T" THEN OASc = 1500 * ((RSco1 / TempTrkNR) ^ TempTrkExp)
     
     ELSEIF TempEvent = "J" THEN
     
        IF RSco1 < (0.15 * TempJmpNR) THEN 
        	 OASc = 0
        ELSE
     	     OASc = 1500 *  (((RSco1 - (0.15 * TempJmpNR)) / (TempJmpNR - (0.15 * TempJmpNR))) ^ TempJmpExp)
        END IF
     
     END IF

     FmtOASc = FormatNumber(OASc,1): RPen1 = RPen1 & "&#13;( NOPS = " & FmtOASc & " )"
  
     ' WriteDebugSQL("Done With Ranking Overall Computation " & TempMemberID & " / " & TempDiv & " / " & TempEvent & " / " & FormatNumber(RSco1,2)) & " / " & FmtOASc
     
  END IF
  
  ' WriteDebugSQL("Ready to Write Final Ranking Record " & TempMemberID & " / " & TempDiv & " / " & TempEvent & " / " & RPen1)

  ' -- Now Insert Constructed Event Ranking Row into the Rankings Table
  ' -- Added

  sSQL = "INSERT INTO " & RankTableName & " (MemberID, Team, TeamStat, Event, Div, DivType, SC_1, RankScore, RnkScoBkup, AWSA_Rat, OpenRating, Reg_SKI, Regl_Plc, Natl_Plc, SkiYearID, Top_Equiv_SC1, Top_Equiv_SC2, Top_Equiv_SC3)"
  sSQL = sSQL & " VALUES ('" & TempMemberID & "', '" & TempTeam & "','" & TempTeamStat & "','" & TempEvent & "', '" & TempDiv & "', '" & TempDivType & "', '" & RSco1 & "', '" & RSco1 & "', '"
  sSQL = sSQL & RExp1 & "With " & RPen1 & "', '  ', '" & RMaxRat & "', '" & R_Ski & "', '" & R_Plc & "', '" & N_Plc & "', " & sProcessingYear & ","
  sSQL = sSQL & "'" & EqScr1 & "','" & EqScr2 & "','" & EqScr3 & "')" 

  Con.Execute(sSQL)

LOOP
' Bottom of Outer Loop
END IF

rs.close


' --------------------------------------------------------------------------------------
' -------- End of New Streamlined Event Ranking Logic, Using EquivScores Table  --------
' --------------------------------------------------------------------------------------











' --------------------------------------------------------------------------------------------------
' --------  SECTION 5:    Creates the new NCWSA Team Rankings.  Added Sep 2008 - Dave Clark --------
' --------------------------------------------------------------------------------------------------

' This is a multi-stage process --
'   1.  Populate RankNums table with all CM/CW scores with Team codes, to Rank by Team
'       within each Div/Event.  Stages so we can pull top 5 skiers for Pseudo-Events.
'   2.  Then use the above table to insert just the top 5 skiers for each Team by 
'       Div/Event into the Team Event Scores Table, w/Order by and by-10's Identity
'       which stages things to compute the Placement Points value for each skier.
'   3.  Then run an update using Min & Max subqueries to compute final NCWSA placement
'       points per skier across each division/event, averaging across ties groups.
'   4.  Next we roll up Team Scores into the final TeamRankings table, for all 2x3 
'       Divisions/Events, picking up only the best 4 placement points for each team.
'   5.  Then finally we roll up across that table twice, once to do the Overall
'       across events for each division, then across Divisions for all 4 events
'       to get the combined team total scores.   That's it !!
'       

' ALTER TABLE USAWSRank.TeamRankings ADD VirtTmStamp DateTime;
' ALTER TABLE USAWSRank.TmEvtScores ADD VirtTmStamp DateTime;


IF Request("Equival") <> "" THEN 
	Response.write(" NCWSA Tms ...")
	Response.Flush
END IF



' -- First Empty the Rank Numbers Table and Reset the Identity Column base 
' -- Drop and then Add the ID column back in again.

sSQL = "DELETE FROM " & RankNumsTableName & ";"
sSQL = sSQL & " Alter Table " & RankNumsTableName & " Drop Column RankSeq;"
sSQL = sSQL & " Alter Table " & RankNumsTableName & " Add RankSeq Int Identity(1,1);"

WriteDebugSQL(sSQL)

Con.Execute(sSQL)




' Now we populate the RankNums table with all Collegiate scores for the Ski Year
' being processed, ordering by Score within team for each division/event, so that
' we will later be able to isolate only the best 5 scoring skiers for each team
' within each division and event -- new Fall 2009 use only TeamStat = 'A' skiers.

sSQL = "INSERT INTO " & RankNumsTableName & " (MemberID, Team, Event, Div, RankScore)"
sSQL = sSQL & " SELECT MemberID, Team, Event, Div, RankScore FROM " & RankTableName
sSQL = sSQL & " WHERE RankScore is not null AND SkiYearID = " & sProcessingYear
sSQL = sSQL & " AND LEFT(Div,1) = 'C' AND Team >= 'AAA' AND Event in ('S','T','J')"
sSQL = sSQL & " AND TeamStat = 'A' Order by Div, Event, Team, RankScore"

WriteDebugSQL(sSQL)

Con.Execute(sSQL)




' Next we Empty the Team/Event Scores Table for the Current sProcessingYear, and 
' also delete any "Custom Team" data rows from both the Team/Event and Team Rankings
' tables, whose Timestamps are over 24 hours old.  Then also Reset the Placement 
' Sequence base in the Team/Event -- Drop and then Add the ID column back in again.

sSQL = "DELETE FROM " & TmEvtScoTableName & " WHERE SkiYearID = " & sProcessingYear & ";"
sSQL = sSQL & " DELETE FROM " & TmEvtScoTableName & " WHERE VirtTmStamp is"
sSQL = sSQL & " NOT Null AND DateDiff(hour,VirtTmStamp,GetDate()) > 24;"
sSQL = sSQL & " DELETE FROM " & TeamRankTableName & " WHERE VirtTmStamp is"
sSQL = sSQL & " NOT Null AND DateDiff(hour,VirtTmStamp,GetDate()) > 24;"
sSQL = sSQL & " Alter Table " & TmEvtScoTableName & " Drop Column PlcmtSeq;"
sSQL = sSQL & " Alter Table " & TmEvtScoTableName & " Add PlcmtSeq Int Identity(10,10);"

WriteDebugSQL(sSQL)

Con.Execute(sSQL)




' Next step is to insert into the Team/Event Scores table, but only
' including the top 5 skiers for each Team, in each Division/Event.

sSQL = "INSERT INTO " & TmEvtScoTableName & " (MemberID, Team,"
sSQL = sSQL & " Event, Div, Score, TeamSeq, SkiYearID)"
sSQL = sSQL & " SELECT RN.MemberID, RN.Team, RN.Event, RN.Div, RN.RankScore,"
sSQL = sSQL & " RM.MaxSeq-RN.RankSeq+1 as TeamSeq, " & sProcessingYear
sSQL = sSQL & " FROM " & RankNumsTableName & " RN, (Select Div, Event,"
sSQL = sSQL & " Team, Max(RankSeq) as MaxSeq FROM  " & RankNumsTableName
sSQL = sSQL & " Group By Div, Event, Team) as RM WHERE RN.Div = RM.Div" 
sSQL = sSQL & " AND RN.Event = RM.Event AND RN.Team = RM.Team"
sSQL = sSQL & " AND RM.MaxSeq - RN.RankSeq <= 4"
sSQL = sSQL & " Order by RN.Div, RN.Event, RN.RankScore"

WriteDebugSQL(sSQL)

Con.Execute(sSQL)




'	Next we calculate NCWSA Placement Points for each Skier in each Div/Event 
' placement set, according to latest NCWSA Rules.  This averages Min and Max 
' PlacementSequence values, for each unique event (ranking) score value, 
' to average across the tie groups, except zero where the raw score is zero.

sSQL = "UPDATE EP SET PlcmtPts = Case when EP.Score <= 0 THEN 0" 
sSQL = sSQL & " ELSE ((EMax.MaxSeq + EMin.MinSeq) / 2) - BMin.BaseSeq end"
sSQL = sSQL & " FROM " & TmEvtScoTableName & " EP, (Select Div, Event,"
sSQL = sSQL & " Score, Min(PlcmtSeq) as MinSeq FROM " & TmEvtScoTableName
sSQL = sSQL & " WHERE SkiYearID = " & sProcessingYear
sSQL = sSQL & " Group by Div, Event, Score) as EMin, (Select Div, Event,"
sSQL = sSQL & " Score, Max(PlcmtSeq) as MaxSeq FROM " & TmEvtScoTableName
sSQL = sSQL & " WHERE SkiYearID = " & sProcessingYear
sSQL = sSQL & " Group by Div, Event, Score) as EMax, (Select Div, Event,"
sSQL = sSQL & " Min(PlcmtSeq) - 10 as BaseSeq FROM " & TmEvtScoTableName
sSQL = sSQL & " WHERE SkiYearID = " & sProcessingYear
sSQL = sSQL & " Group by Div, Event) as BMin WHERE EP.Div = EMin.Div"
sSQL = sSQL & " AND EP.Event = EMin.Event AND EP.Score = EMin.Score"
sSQL = sSQL & " AND  EP.Div = EMax.Div AND EP.Event = EMax.Event"
sSQL = sSQL & " AND EP.Score = EMax.Score AND  EP.Div = BMin.Div"
sSQL = sSQL & " AND EP.Event = BMin.Event AND EP.SkiYearID = " & sProcessingYear

WriteDebugSQL(sSQL)

Con.Execute(sSQL)




' Next step is to Create Team Ranking Scores, by summing the Placement
' Points of the best 4 skiers for each team, across each 2x3 Div/Event.
' First we delete all the rows from the Team Ranking Table for this SkiYearID.

sSQL = "Delete From " & TeamRankTableName & " where SkiYearID = " & sProcessingYear
sSQL = sSQL & " ; INSERT INTO " & TeamRankTableName & " (Team, Div, Event, TeamScore,"
sSQL = sSQL & " SkiYearID) SELECT Team, Div, Event, Sum(PlcmtPts), " & sProcessingYear
sSQL = sSQL & " FROM " & TmEvtScoTableName & " WHERE TeamSeq <= 4"
sSQL = sSQL & " AND SkiYearID = " & sProcessingYear
sSQL = sSQL & " GROUP BY Team, Div, Event;"

WriteDebugSQL(sSQL)

Con.Execute(sSQL)




' Last Step is to roll up across Events and Divisions, creating both
' Divisional Overall Totals, and Combined Team Event and Overall Totals.

sSQL = "INSERT INTO " & TeamRankTableName & "(Team, Div, Event,"
sSQL = sSQL & " TeamScore, SkiYearID) SELECT Team, Div, 'O', Sum(TeamScore)," 
sSQL = sSQL & sProcessingYear & " FROM " & TeamRankTableName 
sSQL = sSQL & " WHERE SkiYearID = " & sProcessingYear & " GROUP BY Team, Div;"
sSQL = sSQL & " INSERT INTO " & TeamRankTableName & "(Team, Div, Event,"
sSQL = sSQL & " TeamScore, SkiYearID) SELECT Team, 'CO', Event, Sum(TeamScore)," 
sSQL = sSQL & sProcessingYear & " FROM " & TeamRankTableName
sSQL = sSQL & " WHERE SkiYearID = " & sProcessingYear & " GROUP BY Team, Event;"

WriteDebugSQL(sSQL)

Con.Execute(sSQL)




'	Final NCWSA Steps are to update the Activity information in the TeamRoster Table ...
'	This is done in two sub-steps, both based on essentially the same "New Activity" sub-query.
'	Note special handling of Team ID's -- excludes special All Star Team ID's from Updates

if sProcessingYear = 1 THEN

'	First sub-step updates existing TeamRoster table rows for new activity.

	sSQL = "Update TR Set	FirstEvent = Case When TR.NumEvents = 0 then"
	sSQL = sSQL & " NA.MinDate else TR.FirstEvent end, LastEvent = CASE "
	sSQL = sSQL & " when left(TR.Team,1) < 'A' then TR.LastEvent else NA.MaxDate"
	sSQL = sSQL & "	end, NumEvents = CASE when left(TR.Team,1) < 'A' then" 
	sSQL = sSQL & "	NA.NumEvts else	TR.NumEvents + NA.NumEvts end FROM "
	sSQL = sSQL & TeamRosterTableName & " TR, (select SX.MemberID, SX.Team,"
	sSQL = sSQL & "	max(SX.EndDate) as MaxDate, min(SX.EndDate) as MinDate,"
	sSQL = sSQL & "	count(*) as NumEvts from (Select MemberID, Team, TourID,"
	sSQL = sSQL & "	max(EndDate) as EndDate from " & RawScoresTableName
	sSQL = sSQL & "	where left(TourID,2) >= '09' and left(Div,1) = 'C' and"
	sSQL = sSQL & "	substring(TourID,3,1) = 'U' and Team in (select TeamID FROM "
	sSQL = sSQL & TeamTableName & ") group by MemberID, Team, TourID) SX JOIN "
	sSQL = sSQL & TeamRosterTableName & "	TR on TR.MemberID = SX.MemberID"
	sSQL = sSQL & "	and TR.Team = SX.Team where (SX.EndDate > TR.LastEvent"
	sSQL = sSQL & "	or TR.NumEvents = 0) group by SX.MemberID, SX.Team ) NA"
	sSQL = sSQL & "	Where	NA.MemberID = TR.MemberID and	NA.Team = TR.Team;"

	WriteDebugSQL(sSQL)

	Con.Execute(sSQL)


'	Second sub-step inserts new TeamRoster table rows for new skiers.

	sSQL = "Insert into " & TeamRosterTableName & " Select NA.Team,"
	sSQL = sSQL & " NA.MemberID, GetDate(), NA.MinDate, CASE when "
	sSQL = sSQL & " left(NA.Team,1) < 'A' then '2000-01-01' else NA.MaxDate"
	sSQL = sSQL & " end, NA.NumEvts, NULL FROM (select SX.MemberID, SX.Team,"
	sSQL = sSQL & "	max(SX.EndDate) as MaxDate, min(SX.EndDate) as MinDate,"
	sSQL = sSQL & "	count(*) as NumEvts from (Select MemberID, Team, TourID,"
	sSQL = sSQL & "	max(EndDate) as EndDate from " & RawScoresTableName
	sSQL = sSQL & "	where left(TourID,2) >= '09' and left(Div,1) = 'C' and"
	sSQL = sSQL & "	substring(TourID,3,1) = 'U' and Team in (select TeamID FROM "
	sSQL = sSQL & TeamTableName & ") group by MemberID, Team, TourID) SX LEFT JOIN "
	sSQL = sSQL & TeamRosterTableName & "	TR on TR.MemberID = SX.MemberID"
	sSQL = sSQL & "	and TR.Team = SX.Team Where TR.MemberID is NULL"
	sSQL = sSQL & "	group by SX.MemberID, SX.Team) NA;"

	WriteDebugSQL(sSQL)

	Con.Execute(sSQL)


END IF


' ------------------------------------------------------------------
' -------  End of NCWSA Team Rankings Calculation Logic ------------
' ------------------------------------------------------------------




' -------------------------------------------------------------------------------------
' -- NSL ranking logic - now that we are done doing the AWSA and NCWSA Skiers, 
'        Create NSL Rankings Rows by Insert Into Select from with Sum and Condition
' -------------------------------------------------------------------------------------

sSQL = "INSERT INTO " & RankTableName
sSQL = sSQL & " (MemberID, Event, Div, Team, SC_3, SkiYearID)"
sSQL = sSQL & " SELECT MemberID, Event, Div, Max(Team), round(sum(NSL_Placement_Points),0) as SC_3, " & sProcessingYear & " as SkiYearID"
sSQL = sSQL & " from " & RawScoresTableName
sSQL = sSQL & " WHERE Enddate <= '" & FormatDateTime(sSkiYearEnd,2) & "' AND EndDate >= '" & FormatDateTime(sSkiYearBegin,2) & "' AND NSL_Placement_Points > 0"
sSQL = sSQL & " group by MemberID, Div, Event ORDER BY MemberID, Event"

WriteDebugSQL(sSQL)

Con.Execute(sSQL)

' DONE SAVING NSL SCORES




' ----------     Bottom of IF to exclude the Event Ranking section above from running   ----------
END IF











' ------------    Outer ID condition to Run Levels Logic   -------------
IF RunLevelLogic = "YES" OR RunOverride = "YES" THEN

IF Request("Equival") <> "" THEN 
	Response.write(" Levels ...")
	Response.Flush
END IF



' ---------------------------------------------------------------------------------------
' ---------   SECTION 6:   Ranking LEVEL logic -- Does S/T/J & Overall Levels   ---------
' ---------------------------------------------------------------------------------------


' If this is sProcessingYear = 1 --> 12 Month rolling, then set flag to enable
' end-of-run response.redirect to QualifyRecalc.asp.

ReCalc12 = "Y"

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





' Step 1a) Empty the Rank Numbers Table and Reset the Identity Column base -- Drop 
' and re-Add the ID column.  Then also Delete any existing rows from the CutOff
' Table for the current Ski Year ID.

sSQL = "DELETE FROM " & RankNumsTableName & ";"
sSQL = sSQL & " Alter Table " & RankNumsTableName & " Drop Column RankSeq;"
sSQL = sSQL & " Alter Table " & RankNumsTableName & " Add RankSeq Int Identity(1,1)"

WriteDebugSQL(sSQL)

Con.Execute(sSQL)

sSQL = "DELETE FROM " & CutOffTableName & " WHERE SkiYearID = " & sProcessingYear

WriteDebugSQL(sSQL)

Con.Execute(sSQL)




' ----  Step 1b) Populate the Rank Numbers Table for the current sProcessingYear  ----
'    NOTES: 
'      a) Asigns the RankSeq Identity values automatically, and
'          in order by RankScore within each Division and Event.  
'      b) Only include Members where their Membership record Federation Code is "USA" ****
'      c) Exclude DivType D which are only for Elite Overall Qualification (later)
'	           and hence don't contribute to ranking percentiles.

sSQL = "INSERT INTO " & RankNumsTableName
sSQL = sSQL & " (MemberID, Event, Div, RankScore)"
sSQL = sSQL & " SELECT RT.MemberID, RT.Event, RT.Div, RT.RankScore"
sSQL = sSQL & " from " & RankTableName & " as RT, "& MemberTableName & " as MT" 
sSQL = sSQL & " WHERE RT.MemberID = MT.PersonIDWithCheckDigit"
sSQL = sSQL & " AND UPPER(MT.FederationCode) = 'USA' and RT.DivType <> 'D'"
sSQL = sSQL & " AND RankScore is not null AND SkiYearID = " &sProcessingYear
sSQL = sSQL & " ORDER BY Div, Event, RankScore"

WriteDebugSQL(sSQL)

Con.Execute(sSQL)



' ----  Step 2) Update the Rank Numbers table and calculate the Percentile for each score  ----
'   NOTES:      
'      a) As a function of the RankSeq Number, within the scope of the
'           Min and Max RankSeq values for the corresponding Division and Event.  
'      b) Also store the actual RankNum (1 to Max) for each score
'            for now, also store the Rank_Level for the old "Base 10 level logic" 
'           DJC 1/19/2008 - DELETE this later 

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




' ----  Step 3) Prepare the CutOff Table entries for the current Ski Year ID  ----
'   NOTES:
'     a) Initial populate these rows by extracting the current Level Percentiles 
'           from the Division Control Table, by Division and Event.  
'     b) This recasts these parameters into a Division / Event keying framework, which is more
'           compatible with the way the scores data is organized.

sSQL = "INSERT INTO " & CutOffTableName & " (Div, Event, SkiYearID, Pct1, Pct2, Pct3,  Pct4, Pct5, Pct6, Pct7, Pct8, Pct9)"

sSQL = sSQL & " Select Div, 'S', SkiYearID, Percent_01_S, Percent_02_S, Percent_03_S, Percent_04_S,"
sSQL = sSQL & " Percent_05_S, Percent_06_S, Percent_07_S, Percent_08_S, Percent_09_S From " & DivisionsTableName 
sSQL = sSQL & " WHERE (left(Div,1) in ('B','G','M','W','O','E') or Div = 'SM' or Div = 'SW') and SkiYearID = " & sProcessingYear

sSQL = sSQL & " UNION Select Div, 'T', SkiYearID, Percent_01_T, Percent_02_T, Percent_03_T, Percent_04_T,"
sSQL = sSQL & " Percent_05_T, Percent_06_T, Percent_07_T, Percent_08_T, Percent_09_T From " & DivisionsTableName 
sSQL = sSQL & " WHERE (left(Div,1) in ('B','G','M','W','O','E') or Div = 'SM' or Div = 'SW') and SkiYearID = " & sProcessingYear

sSQL = sSQL & " UNION Select Div, 'J', SkiYearID, Percent_01_J, Percent_02_J, Percent_03_J, Percent_04_J,"
sSQL = sSQL & " Percent_05_J, Percent_06_J, Percent_07_J, Percent_08_J, Percent_09_J From " & DivisionsTableName 
sSQL = sSQL & " WHERE (left(Div,1) in ('B','G','M','W','O','E') or Div = 'SM' or Div = 'SW') and SkiYearID = " & sProcessingYear

sSQL = sSQL & " UNION Select Div, 'O', SkiYearID, Percent_01_O, Percent_02_O, Percent_03_O, Percent_04_O,"
sSQL = sSQL & " Percent_05_O, Percent_06_O, Percent_07_O, Percent_08_O, Percent_09_O From " & DivisionsTableName 
sSQL = sSQL & " WHERE (left(Div,1) in ('B','G','M','W','O','E') or Div = 'SM' or Div = 'SW') and SkiYearID = " & sProcessingYear

WriteDebugSQL(sSQL)

Con.Execute(sSQL)




' ----  Step 4) Update the RankNum and RankPct on the Rankings Table  ----
'    NOTES:
'       a) Match in the Pre-Calculated RankNum Value (True Rank from 1 to N)
'       b) RankPct (True Percentile fm 100 down)
'         For now also the old Rank_Level (fm 10 down).  
'       c) Ties at the same RankScore value - will get equal values for these three measures, since we Group 
'            the Rank Numbers Table by RankScore.  
'       d) Classify the Ranking Level, according to the Level Percentiles joined in from the CutOff Table,
'            by Division and event, that we just created in the step immediately above.

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

' IF sProcessingYear <> 21 then 
Con.Execute(sSQL)
' End IF





' ----  Step 5) Recap the Ranking Level Cut Off Score, into the Cut Off Table  ---- 
'   NOTES:
'     a) Using the Min(RankScore) values found for each of the 9 levels for each Div and event - for the current Ski Year ID
'     b) For each level, uses COALESCE to find the first existing equal or lower level cutoff score, or zero if none exists

sSQL = "UPDATE CT Set COA1 = Coalesce(LT1.Cutoff,0),"
sSQL = sSQL & " COA2 = Coalesce(LT2.Cutoff,LT1.Cutoff,0),"
sSQL = sSQL & " COA3 = Coalesce(LT3.Cutoff,LT2.Cutoff,LT1.Cutoff,0),"
sSQL = sSQL & " COA4 = Coalesce(LT4.Cutoff,LT3.Cutoff,LT2.Cutoff,LT1.Cutoff,0),"
sSQL = sSQL & " COA5 = Coalesce(LT5.Cutoff,LT4.Cutoff,LT3.Cutoff,LT2.Cutoff,LT1.Cutoff,0),"
sSQL = sSQL & " COA6 = Coalesce(LT6.Cutoff,LT5.Cutoff,LT4.Cutoff,LT3.Cutoff,LT2.Cutoff,LT1.Cutoff,0),"
sSQL = sSQL & " COA7 = Coalesce(LT7.Cutoff,LT6.Cutoff,LT5.Cutoff,LT4.Cutoff,LT3.Cutoff,LT2.Cutoff,LT1.Cutoff,0),"
sSQL = sSQL & " COA8 = Coalesce(LT8.Cutoff,LT7.Cutoff,LT6.Cutoff,LT5.Cutoff,LT4.Cutoff,LT3.Cutoff,LT2.Cutoff,LT1.Cutoff,0),"
sSQL = sSQL & " COA9 = Coalesce(LT9.Cutoff,LT8.Cutoff,LT7.Cutoff,LT6.Cutoff,LT5.Cutoff,LT4.Cutoff,LT3.Cutoff,LT2.Cutoff,LT1.Cutoff,0)"
sSQL = sSQL & " FROM " & CutOffTableName & " AS CT "
sSQL = sSQL & " LEFT JOIN (Select Div, Event, Min(RankScore) as Cutoff FROM " & RankTableName & " Where SkiYearID = " & sProcessingYear
sSQL = sSQL & " and RankNum is not Null and Right(AWSA_Rat,1)='9' Group by Div, Event) as LT9 on LT9.Div=CT.Div and LT9.Event=CT.Event"
sSQL = sSQL & " LEFT JOIN (Select Div, Event, Min(RankScore) as Cutoff FROM " & RankTableName & " Where SkiYearID = " & sProcessingYear
sSQL = sSQL & " and RankNum is not Null and Right(AWSA_Rat,1)='8' Group by Div, Event) as LT8 on LT8.Div=CT.Div and LT8.Event=CT.Event"
sSQL = sSQL & " LEFT JOIN (Select Div, Event, Min(RankScore) as Cutoff FROM " & RankTableName & " Where SkiYearID = " & sProcessingYear
sSQL = sSQL & " and RankNum is not Null and Right(AWSA_Rat,1)='7' Group by Div, Event) as LT7 on LT7.Div=CT.Div and LT7.Event=CT.Event"
sSQL = sSQL & " LEFT JOIN (Select Div, Event, Min(RankScore) as Cutoff FROM " & RankTableName & " Where SkiYearID = " & sProcessingYear
sSQL = sSQL & " and RankNum is not Null and Right(AWSA_Rat,1)='6' Group by Div, Event) as LT6 on LT6.Div=CT.Div and LT6.Event=CT.Event"
sSQL = sSQL & " LEFT JOIN (Select Div, Event, Min(RankScore) as Cutoff FROM " & RankTableName & " Where SkiYearID = " & sProcessingYear
sSQL = sSQL & " and RankNum is not Null and Right(AWSA_Rat,1)='5' Group by Div, Event) as LT5 on LT5.Div=CT.Div and LT5.Event=CT.Event"
sSQL = sSQL & " LEFT JOIN (Select Div, Event, Min(RankScore) as Cutoff FROM " & RankTableName & " Where SkiYearID = " & sProcessingYear
sSQL = sSQL & " and RankNum is not Null and Right(AWSA_Rat,1)='4' Group by Div, Event) as LT4 on LT4.Div=CT.Div and LT4.Event=CT.Event"
sSQL = sSQL & " LEFT JOIN (Select Div, Event, Min(RankScore) as Cutoff FROM " & RankTableName & " Where SkiYearID = " & sProcessingYear
sSQL = sSQL & " and RankNum is not Null and Right(AWSA_Rat,1)='3' Group by Div, Event) as LT3 on LT3.Div=CT.Div and LT3.Event=CT.Event"
sSQL = sSQL & " LEFT JOIN (Select Div, Event, Min(RankScore) as Cutoff FROM " & RankTableName & " Where SkiYearID = " & sProcessingYear
sSQL = sSQL & " and RankNum is not Null and Right(AWSA_Rat,1)='2' Group by Div, Event) as LT2 on LT2.Div=CT.Div and LT2.Event=CT.Event"
sSQL = sSQL & " LEFT JOIN (Select Div, Event, Min(RankScore) as Cutoff FROM " & RankTableName & " Where SkiYearID = " & sProcessingYear
sSQL = sSQL & " and RankNum is not Null and Right(AWSA_Rat,1)='1' Group by Div, Event) as LT1 on LT1.Div=CT.Div and LT1.Event=CT.Event"
sSQL = sSQL & " WHERE CT.SkiYearID = " & sProcessingYear

WriteDebugSQL(sSQL)

' IF sProcessingYear <> 21 then 
Con.Execute(sSQL)
' End IF





' ----  Step 5) Propagate the Level 9 Cutoff Score values from the Elite pools across all other age division events 
'  NOTES:
'    a) For the applicable gender - setting the COA9_Opn cutoffs.  
'    b) Adjusts the Jump cutoffs for condition differences between the Elite and Native divisions
'    c) Done as two sub-steps


'	    Step 5a) -- First sub-step does OPEN for all divisions all ages, except NOT overriding EM/EW/SM/SW of course.

sSQL = "UPDATE CA" 
sSQL = sSQL & " SET COA9 = CASE WHEN CA.Event in ('O','S','T') OR DA.Ramp1 = 0" 
sSQL = sSQL & "  OR DE.Ramp1 < DA.Ramp1 OR DE.Max_J1 < DA.Max_J1 THEN CE.COA9"
sSQL = sSQL & "  ELSE CE.COA9 - (12 * ((DE.Ramp1-DA.Ramp1)/.020)) - (8 * ((DE.Max_J1-DA.Max_J1)/3)) END"

sSQL = sSQL & " , COA9_Opn = CASE WHEN CA.Event IN ('O','S','T') OR DA.Ramp1 = 0" 
sSQL = sSQL & "  OR DE.Ramp1 < DA.Ramp1 or DE.Max_J1 < DA.Max_J1 THEN CE.COA9"
sSQL = sSQL & "  ELSE CE.COA9 - (12 * ((DE.Ramp1-DA.Ramp1)/.020)) - (8 * ((DE.Max_J1-DA.Max_J1)/3)) END"

sSQL = sSQL & " FROM " & CutOffTableName & " AS CA, " & CutOffTableName & " AS CE"
sSQL = sSQL & "  , " & DivisionsTableName & " as DA, " & DivisionsTableName & " AS DE"

sSQL = sSQL & " WHERE CA.Div = DA.Div" 
sSQL = sSQL & "  AND CA.SkiYearID = DA.SkiYearID"
sSQL = sSQL & "  AND CE.Div = DE.Div"
sSQL = sSQL & "  AND CE.SkiYearID = DE.SkiYearID" 
sSQL = sSQL & "  AND CA.Event = CE.Event" 
sSQL = sSQL & "  AND DA.Sex = DE.Sex"
sSQL = sSQL & "  AND DA.SkiYearID = DE.SkiYearID"
sSQL = sSQL & "  AND Upper(left(DE.Div,1)) = 'E'"
sSQL = sSQL & "  AND DA.Div not in ('EM','EW','SM','SW')" 
sSQL = sSQL & "  AND DE.SkiYearID = " & sProcessingYear

WriteDebugSQL(sSQL)

Con.Execute(sSQL)



'	    Step 5b) -- Second substep, does MASTERS for just those divisions for Age 35 and up.

sSQL = "UPDATE CA" 
sSQL = sSQL & " SET COA9 = CASE WHEN CA.Event IN ('O','S','T') OR DA.Ramp1 = 0" 
sSQL = sSQL & "    OR DE.Ramp1 < DA.Ramp1 OR DE.Max_J1 < DA.Max_J1 THEN CE.COA9"
sSQL = sSQL & "   ELSE CE.COA9 - (12 * ((DE.Ramp1-DA.Ramp1)/.020)) - (8 * ((DE.Max_J1-DA.Max_J1)/3)) END"
sSQL = sSQL & ", COA9_Mst = CASE WHEN CA.Event IN ('O','S','T') OR DA.Ramp1 = 0" 
sSQL = sSQL & "    OR DE.Ramp1 < DA.Ramp1 OR DE.Max_J1 < DA.Max_J1 THEN CE.COA9"
sSQL = sSQL & "   ELSE CE.COA9 - (12 * ((DE.Ramp1-DA.Ramp1)/.020)) - (8 * ((DE.Max_J1-DA.Max_J1)/3)) END"

sSQL = sSQL & " FROM " & CutOffTableName & " AS CA"
sSQL = sSQL & " , " & CutOffTableName & " AS CE"
sSQL = sSQL & " , " & DivisionsTableName & " AS DA"
sSQL = sSQL & " , " & DivisionsTableName & " AS DE"

sSQL = sSQL & " WHERE CA.Div = DA.Div" 
sSQL = sSQL & "  AND CA.SkiYearID = DA.SkiYearID" 
sSQL = sSQL & "  AND CE.Div = DE.Div"
sSQL = sSQL & "  AND CE.SkiYearID = DE.SkiYearID" 
sSQL = sSQL & "  AND CA.Event = CE.Event" 
sSQL = sSQL & "  AND DA.Sex = DE.Sex"
sSQL = sSQL & "  AND DA.SkiYearID = DE.SkiYearID" 
sSQL = sSQL & "  AND Upper(left(DE.Div,1)) = 'S'"
sSQL = sSQL & "  AND	DA.Low_Age >= 35" 
sSQL = sSQL & "  AND DE.SkiYearID = " & sProcessingYear

WriteDebugSQL(sSQL)

Con.Execute(sSQL)





' ----  Step 6) Update Rankings table - Set AWSA_Rat to Level 9  ----

'   NOTES:
'     a) Apply skiers with ranking scores that equal or exceed the (possibly adjusted) Level 9 
'	         cutoff score for their applicable division and event.  


'     Two steps:
'        Step 6a) -- First we do all events for all Age divisions, except do Overall ONLY for the Elite Pseudo-Divs EM/EW/SM/SW.

sSQL = "UPDATE RT"
sSQL = sSQL & " SET AWSA_Rat = CT.Event + '9'"
sSQL = sSQL & "  FROM " & RankTableName & " AS RT," & CutOffTableName & " AS CT"
sSQL = sSQL & "  WHERE CT.Div = RT.Div"
sSQL = sSQL & "    AND CT.Event = RT.Event"
sSQL = sSQL & "    AND CT.SkiYearID = RT.SkiYearID" 
sSQL = sSQL & "    AND RT.SkiYearID = " & sProcessingYear
sSQL = sSQL & "    AND ( (RT.Div in ('EM','EW','SM','SW') and RT.Event = 'O') OR RT.Event <> 'O' )"
sSQL = sSQL & "    AND CT.COA9 <= RT.RankScore"

WriteDebugSQL(sSQL)

' IF sProcessingYear <> 21 then 
Con.Execute(sSQL)
' End IF



'        Step 6b) --Then Post overall Level 9 for Age Divs that fall into the Elite
'	          pools (both DivTypes C and D), against the already-coded rankings rows for
'	          those applicable Elite pool pseudo-divisions as set in the previous query.

sSQL = "UPDATE RT SET AWSA_Rat = 'O9' FROM "
sSQL = sSQL & RankTableName & " as RT, " & RankTableName & " as ET" 
sSQL = sSQL & " WHERE ET.MemberID = RT.MemberID"
sSQL = sSQL & " and ET.Event = RT.Event and RT.Event = 'O'"
sSQL = sSQL & " and ET.SkiYearID = RT.SkiYearID"
sSQL = sSQL & " and RT.SkiYearID = " & sProcessingYear & " and ET.Div = CASE"
sSQL = sSQL & " When RT.Div in ('B2','B3','M1','M2','OM') then 'EM'"
sSQL = sSQL & " When RT.Div in ('G2','G3','W1','W2','OW') then 'EW'"
sSQL = sSQL & " When RT.Div in ('M3','M4','M5','M6','M7','MM') then 'SM'"
sSQL = sSQL & " When RT.Div in ('W3','W4','W5','W6','MW') then 'SW' else '??' end"
sSQL = sSQL & " and right(ET.AWSA_Rat,1) = '9'"

WriteDebugSQL(sSQL)

' IF sProcessingYear <> 21 then 
Con.Execute(sSQL)
' End IF




' ----  Step 7)  Inserts to the Elite Qualified Through Date table  ----

'	  NOTES:
'     a) Apply Updates for those MemberID/Div/Events in Level 9 today using the EM/EW/SM pool pseudo-divisions 
'         in the Rankings table as the source.
'	    b) Done in three substeps
'        - 7a) First or Updates
'        - 7b) Second one creates eMails for the new Elite skiers
'        - 7c) final step does Inserts for those new Elite skiers


'	Before updating the Elite Qualified Through Date table, first we need to delete
'	any "Expired" Elite Qualifications, those with QualThru dates less than the
'	End Date for the current sProcessingYear 

sSQL = "DELETE FROM " & EliteDateTableName & " Where SkiYearID = " & sProcessingYear & " and QualThru <"
sSQL = sSQL & " (Select EndDate FROM " & SkiYearTableName & " WHERE SkiYearID = " & sProcessingYear & ")"

WriteDebugSQL(sSQL)

Con.Execute(sSQL)







'	-------------------------------------------------------------------------------------
'	---  SECTION 7:  Elite Qualified Through Dates only done for sProcessingYear = 1  ---
'	-------------------------------------------------------------------------------------

if sProcessingYear = 1 THEN

	'	7a) -- First sub-step does Updates for all existing MemberID / Div / Event rows.

	sSQL = "UPDATE EQD SET QualThru = CASE WHEN GetDate() > SY.EndDate THEN DateAdd(Day,365,SY.EndDate) ELSE DateAdd(Day,365,GetDate()) END" 
	sSQL = sSQL & " , DivOrig = CASE WHEN patindex('%) AS %', RnkScoBkup) between 3 and 15"
	sSQL = sSQL & "     THEN substring(RnkScoBkup,patindex('%) AS %',RnkScoBkup)+5,2) ELSE RT.Div END"

	sSQL = sSQL & " FROM " & EliteDateTableName & " AS EQD, "
	sSQL = sSQL & RankTableName & " as RT, " & SkiYearTableName & " as SY" 

	sSQL = sSQL & " WHERE EQD.MemberID = RT.MemberID" 
	sSQL = sSQL & "   AND EQD.SkiYearID = RT.SkiYearID"
	sSQL = sSQL & "   AND RT.SkiYearID = SY.SkiYearID" 
	sSQL = sSQL & "   AND SY.SkiYearID = " & sProcessingYear 
	sSQL = sSQL & "   AND RT.Div in ('EM','EW','SM','SW')" 
	sSQL = sSQL & "   AND EQD.DivElite = CASE WHEN RT.Div = 'EM' THEN 'OM'" 
	sSQL = sSQL & "     WHEN RT.Div = 'EW' THEN 'OW'"
	sSQL = sSQL & "     WHEN RT.Div = 'SM' THEN 'MM'"
	sSQL = sSQL & "     WHEN RT.Div = 'SW' THEN 'MW' END"
	sSQL = sSQL & "   AND EQD.Event = left(RT.Event,1)" 
	sSQL = sSQL & "   AND right(RT.AWSA_Rat,1) = '9'"
	
	'	WriteDebugSQL(sSQL)

	Con.Execute(sSQL)




	'	7b) -- Second sub-step creates "Welcome to Elite" emails, for all Level 9 
	'	MemberID / Div / Events not already present in the table -- these are
	'	Newly Qualified Elite Skiers, so let's give 'em a rousing welcome.

	' Generate the select query,

	sSQL = " Select RT.MemberID, MT.FirstName, MT.LastName, MT.Email,"
	sSQL = sSQL & " CASE when RT.Div = 'EM' then 'Open Men' when RT.Div = 'EW' then 'Open Women'"
	sSQL = sSQL & " when RT.Div = 'SM' then 'Masters Men' when RT.Div = 'SW' then 'Masters Women' end"
	sSQL = sSQL & " as EliteDiv, CASE when RT.Event = 'S' then 'Slalom' when RT.Event = 'T' then 'Tricks'"
	sSQL = sSQL & " when RT.Event = 'J' then 'Jumping' else 'Overall' end as EliteEvent,"
	sSQL = sSQL & " CASE when patindex('%) as %', RnkScoBkup) between 3 and 15"
	sSQL = sSQL & " then substring(RnkScoBkup,patindex('%) as %',RnkScoBkup)+5,2)"
	sSQL = sSQL & " else RT.Div end as OrigDiv" 

	sSQL = sSQL & " FROM " & RankTableName & " as RT"
	 
	sSQL = sSQL & " JOIN " & MemberTableName & " AS MT ON MT.PersonIDWithCheckDigit = RT.MemberID"
	
	sSQL = sSQL & " LEFT JOIN " & EliteDateTableName & " AS EQD ON EQD.MemberID = RT.MemberID"
	sSQL = sSQL & "   AND EQD.DivElite = CASE WHEN RT.Div = 'EM' THEN 'OM'" 
	sSQL = sSQL & "     WHEN RT.Div = 'EW' THEN 'OW'"
	sSQL = sSQL & " 		WHEN RT.Div = 'SM' THEN 'MM'" 
	sSQL = sSQL & "     WHEN RT.Div = 'SW' THEN 'MW' END"
	sSQL = sSQL & "	  AND EQD.Event = left(RT.Event,1)" 
	sSQL = sSQL & "   AND EQD.SkiYearID = RT.SkiYearID" 

	sSQL = sSQL & "	WHERE RT.SkiYearID = " & sProcessingYear 
	sSQL = sSQL & "   AND RT.Div IN ('EM','EW','SM','SW')" 
	sSQL = sSQL & "   AND RT.AWSA_Rat = LEFT(RT.Event,1) + '9'"
	sSQL = sSQL & "   AND patindex('%@%',MT.Email) > 0"
	sSQL = sSQL & "   AND MT.FederationCode = 'USA'"
	sSQL = sSQL & "   AND EQD.MemberID is Null" 

	sSQL = sSQL & " ORDER BY RT.MemberID desc"
	
	WriteDebugSQL(sSQL)


  '	Set up to produce the emails.

	SetupEmailService

	objMessage.Subject="Welcome to AWSA Elite Skier Status"

	' objMessage.From = """AWSA President"" <awsa_prez@usawaterski.org>"
	objMessage.From = """AWSA President"" <competition@usawaterski.org>"
	
'	objMessage.BCC = """Dave Clark"" <awsatechdude@comcast.net>; ""Bob Mayhew"" <SkiDent@gmail.com>; ""Kelvin Kelm"" <kelvin@sportpixusa.net>"
' objMessage.BCC = """Jeff Surdej"" <j_surdej@yahoo.com>; ""Kelvin Kelm"" <kelvin@sportpixusa.net>; ""Dave Clark"" <awsatechdude@comcast.net>"
objMessage.BCC = """Jeff Surdej"" <j_surdej@yahoo.com>; ""Kelvin Kelm"" <kelvin@sportpixusa.net>; ""Kirby Whetsel"" <kwhetsel@charter.net>"
' 	objMessage.BCC = """Jeff Surdej"" <j_surdej@yahoo.com>; ""Kelvin Kelm"" <kelvin@sportpixusa.net>; ""Dave Clark"" <awsatechdude@comcast.net>; ""Steve Hansen"" <shansen@dakotatechgroup.com>"

	rs.open sSQL, sConnectionToTRATable, 0, 1

'	Finally we process the resulting record-set, generating and
'	sending a customized email note to each such qualifier.

	IF NOT rs.eof THEN
		
'		rs.MoveFirst
		DO Until rs.eof

'
'		Here below is a sample of what the plain text version looks like ...
'
'		To:    [FirstNm] [LastNm]
'		Re:    AWSA Elite Status in [EliteDiv] [Event]
'		Date:  [long format date]
'
'		Dear [Skier],
'
'		Your prowess as a competitive water skier has elevated you to a 
'		new pinnacle -- status as an AWSA Elite Skier, in the [Masters Men] 
'		Division in [Jumping], based on your Ranking today in the [B3] 
'		Division.  This is a significant achievement for which you 
'		should feel quite proud.  Feel free to pat yourself on the back.  
'	
'		This reflects not only your athleticism, but your dedication to 
'		disciplined practice sessions and your participation at numerous 
'		tournaments.  Please note that this credential is limited to the 
'		top 7% of the ranked [Jumping] skiers eligible for consideration 
'		as [Masters Men].  Simply stated, this puts you in the company of 
'		AWSA's most outstanding water skier athletes.
'
'		Your next step:  a number one ranking among these elite skiers.
'		Again, my heartiest congratulations !
'	
'		Jeff Surdej
'		President
'		American Water Ski Association
'

			objMessage.To = """" & rs("FirstName") & " " & rs("LastName") & """ <" & rs("Email") & ">"
'			' objMessage.To = """Dave Clark"" <awsatechdude@comcast.net>; ""Jeff Surdej"" <j_surdej@yahoo.com>; ""Kelvin Kelm"" <kelvin@sportpixusa.net>"
			objMessage.To = """Kirby Whetsel"" <kwhetsel@charter.net>; ""Jeff Surdej"" <j_surdej@yahoo.com>; ""Kelvin Kelm"" <kelvin@sportpixusa.net>"
			
			WriteDebugSQL ( "Elite Email being Generated TO: " & objMessage.To & " -- CC: " & objMessage.cc & " -- BCC: " & objMessage.bcc & " -- From: " & objMessage.from )

			sSQL = "<html><head><title>Welcome to AWSA Elite Skier Status</title></head>"
			sSQL = sSQL & "<body><basefont face=""arial,sans-serif,helvetica,verdana,tahoma"" color=""#000000"" size=""2"">"

			sSQL = sSQL & "<div style=""border: double 20px #ff0505;"
			sSQL = sSQL & " padding: 25px;"
			sSQL = sSQL & " margin: 10;"
'			sSQL = sSQL & " text-align: justify;"
			sSQL = sSQL & " line-height: 23px;"
			sSQL = sSQL & " color: #070707;"
			sSQL = sSQL & " font-size: 18px"">"
			
			sSQL = sSQL & "<p>To:&nbsp;&nbsp;&nbsp;&nbsp; " & rs("FirstName") & " " & rs("LastName")
			sSQL = sSQL & "<br>Re:&nbsp;&nbsp;&nbsp;&nbsp; AWSA Elite Status in " & rs("EliteDiv") & " " & rs("EliteEvent")
			sSQL = sSQL & "<br>Date:&nbsp; " & FormatDateTime(date(),1) & "</p>"

'			Line below is header for when running in debug mode going to developers -- this documents who would go to.
'			sSQL = sSQL & "<p>HTML eMail to:  """ & rs("FirstName") & " " & rs("LastName") & """ &lt;" & rs("Email") & "&gt;</p>"

			sSQL = sSQL & "<p>Dear " & rs("FirstName") & ",</p>"
'			sSQL = sSQL & "<p>Dear " & rs("FirstName") & " " & rs("LastName") & ",</p>"

			sSQL = sSQL & "<p>Your prowess as a competitive water skier has elevated you to a new"
			sSQL = sSQL & " pinnacle -- status as an AWSA Elite Skier, in the " & rs("EliteDiv")
			sSQL = sSQL & " Division in " & rs("EliteEvent") & ", based on your new Level 9 Ranking today in the " & rs("OrigDiv")
			sSQL = sSQL & " Division.&nbsp; This is a significant achievement for which you"
			sSQL = sSQL & " should feel quite proud.&nbsp; Feel free to pat yourself on the back.</p>"
		 
			sSQL = sSQL & "<p>This reflects not only your athleticism, but your dedication to"
			sSQL = sSQL & " disciplined practice sessions and your participation at numerous"
			sSQL = sSQL & " tournaments.&nbsp; Please note that this credential is limited to the"
			sSQL = sSQL & " top 7% of the ranked " & rs("EliteEvent") & " skiers eligible for consideration"
			sSQL = sSQL & " as " & rs("EliteDiv") & ".&nbsp; Simply stated, this puts you in the company of"
			sSQL = sSQL & " AWSA's most outstanding water ski athletes.</p>"

			sSQL = sSQL & "<p>What does this mean?  It means you are now eligible compete in the elite division listed above for" 
			sSQL = sSQL & "the next 12 months.  Your next step:&nbsp; Level 10 and a number one ranking among these elite skiers.&nbsp; "
			sSQL = sSQL & "Again, my heartiest congratulations !</p>"

			sSQL = sSQL & "<p>Jeff Surdej<br>President,<br>American Water Ski Association</p>"
			sSQL = sSQL & "</div></body></html>"

			objMessage.HTMLBody = sSQL

			objMessage.Send
			nNewElites = nNewElites + 1

			rs.MoveNext

		LOOP
		
	END IF

	rs.close
	set objMessage = Nothing





	'	----   7b) -- Final sub-step does Inserts for all Level 9 MemberID / Div / Events not already present.

	sSQL = "Insert Into USAWSRank.EliteDates (MemberID, DivElite, DivOrig, Event, SkiYearID, QualThru)"
	sSQL = sSQL & " Select RT.MemberID, CASE when RT.Div = 'EM' then 'OM'"
	sSQL = sSQL & " when RT.Div = 'EW' then 'OW' when RT.Div = 'SM' then 'MM' when RT.Div = 'SW' then 'MW' end,"
	sSQL = sSQL & " CASE when patindex('%) as %', RnkScoBkup) between 3 and 15 then"
	sSQL = sSQL & " substring(RnkScoBkup,patindex('%) as %',RnkScoBkup)+5,2) else RT.Div end,"
	sSQL = sSQL & " RT.Event, RT.SkiYearID, Case When GetDate() > SY.EndDate"
	sSQL = sSQL & " then DateAdd(Day,365,SY.EndDate) else DateAdd(Day,365,GetDate()) end FROM "
	sSQL = sSQL & SkiYearTableName & " as SY JOIN " & RankTableName & " as RT"
	sSQL = sSQL & " on RT.SkiYearID = SY.SkiYearID LEFT JOIN " & EliteDateTableName
	sSQL = sSQL & " as EQD on EQD.MemberID = RT.MemberID and EQD.DivElite = CASE"
	sSQL = sSQL & " when RT.Div = 'EM' then 'OM' when RT.Div = 'EW' then 'OW'"
	sSQL = sSQL & " when RT.Div = 'SM' then 'MM' when RT.Div = 'SW' then 'MW' end and EQD.Event = left(RT.Event,1)"
	sSQL = sSQL & " and EQD.SkiYearID = RT.SkiYearID WHERE EQD.MemberID is Null"
	sSQL = sSQL & " and RT.Div in ('EM','EW','SM','SW') and RT.AWSA_Rat = left(RT.Event,1) + '9'"
	sSQL = sSQL & " and SY.SkiYearID = " & sProcessingYear 

	WriteDebugSQL(sSQL)

	Con.Execute(sSQL)

'	Bottom of IF for sProcessingYear = 1 for Elite Date updates.
END IF

' Bottom of IF to exclude above LEVELS section for debugging purposes
END IF








IF sProcessingYear = 1 THEN

Dim RunLevel10
RunLevel10 = "Y"
IF RunLevel10 = "Y" THEN 



' ----------------------------------------------------------------------
' #########    LEVEL 10 EMAIL NOTICE - 3/12/2018 Mark Crone  ###########
' ----------------------------------------------------------------------

	'	7d) --  "Level 10 Notice" emails sent 1x to all new Level 10 instances in a Ski Year  
	'	   MemberID / Div / Events not already present in the table Newly Qualified Level 10 Skiers
	'    Added: 3/12/2018 by Mark Crone

	' -- Generate the select query to get all NEW Level 10 members for email list.
	' -- Creates one row per member with all Level 10 possibilities appended.
	' -- Omits Members + Event + Div who have first instance before July 1 of Ski Year.   

	sSQL = " SELECT row.MemberID, row.SkiYearID, Email, FirstName, LastName"
	sSQL = sSQL & " 	, S_First_Instance_E, T_First_Instance_E, J_First_Instance_E, O_First_Instance_E"
	sSQL = sSQL & " 	, S_First_Instance_S, T_First_Instance_S"
	sSQL = sSQL & " 	, MemberAge"
	sSQL = sSQL & " FROM"

	sSQL = sSQL & " ( SELECT DISTINCT MemberID, l.SkiYearID, skiyear, sy.skiyear-datepart(yyyy,mt.birthdate)-1 AS MemberAge"
	sSQL = sSQL & "	, Email, FirstName, LastName" 
	sSQL = sSQL & " 			FROM " & EquivLevel10TableName & " l"
	sSQL = sSQL & " 			JOIN " & SkiYearTableName & " sy ON sy.skiyearid = l.skiyearid"
 	sSQL = sSQL & "       JOIN " & MemberShortTableName & " mt ON mt.PersonID = RIGHT(l.MemberID,8)"	
	sSQL = sSQL & " 					WHERE sy.defaultyear = 1"
	sSQL = sSQL & " 						AND ( (Event IN ('S','T','J','O') AND Div IN ('EM','EW') AND (sy.skiyear-datepart(yyyy,mt.birthdate)-1 BETWEEN 18 AND 35)) "
	sSQL = sSQL & " 						   OR " 	
	sSQL = sSQL & " 						      (Event IN ('S','T') AND Div IN ('SM','SW')) ) "
	sSQL = sSQL & " 						AND ( Sent_Notice IS NULL OR Sent_Notice='N')"
	sSQL = sSQL & "							AND First_Instance < Level10_EndDate "
	sSQL = sSQL & " ) row"
 	
	sSQL = sSQL & " LEFT JOIN" 
	sSQL = sSQL & " 		( SELECT MemberID, Event, Div, l.SkiYearID, First_Instance AS S_First_Instance_E" 	
	sSQL = sSQL & " 				FROM " & EquivLevel10TableName & " l"
	sSQL = sSQL & " 					WHERE Event='S' AND Div IN ('EM','EW') AND ( Sent_Notice IS NULL OR Sent_Notice='N') ) se"
	sSQL = sSQL & " ON se.SkiYearID=row.SkiYearID AND se.MemberID=row.MemberID"	
	sSQL = sSQL & " LEFT JOIN" 
	sSQL = sSQL & " 		( SELECT MemberID, Event, Div, l.SkiYearID, First_Instance AS S_First_Instance_S" 	
	sSQL = sSQL & " 				FROM " & EquivLevel10TableName & " l"
	sSQL = sSQL & " 					WHERE Event='S' AND Div IN ('SM','SW') AND ( Sent_Notice IS NULL OR Sent_Notice='N') ) ss"
	sSQL = sSQL & " ON ss.SkiYearID=row.SkiYearID AND ss.MemberID=row.MemberID"	
					
	sSQL = sSQL & " LEFT JOIN" 
	sSQL = sSQL & " 		( SELECT MemberID, Event, Div, l.SkiYearID, First_Instance AS T_First_Instance_E" 	
	sSQL = sSQL & " 				FROM " & EquivLevel10TableName & " l"
	sSQL = sSQL & " 					WHERE Event='T' AND Div IN ('EM','EW') AND ( Sent_Notice IS NULL OR Sent_Notice='N') ) te"
	sSQL = sSQL & " ON te.SkiYearID=row.SkiYearID AND te.MemberID=row.MemberID"
	sSQL = sSQL & " LEFT JOIN" 
	sSQL = sSQL & " 		( SELECT MemberID, Event, Div, l.SkiYearID, First_Instance AS T_First_Instance_S" 	
	sSQL = sSQL & " 				FROM " & EquivLevel10TableName & " l"
	sSQL = sSQL & " 					WHERE Event='T' AND Div IN ('SM','SW') AND ( Sent_Notice IS NULL OR Sent_Notice='N') ) ts"
	sSQL = sSQL & " ON ts.SkiYearID=row.SkiYearID AND ts.MemberID=row.MemberID"

	sSQL = sSQL & " LEFT JOIN" 
	sSQL = sSQL & " 		( SELECT MemberID, Event, Div, l.SkiYearID, First_Instance AS J_First_Instance_E" 	
	sSQL = sSQL & " 				FROM " & EquivLevel10TableName & " l"
	sSQL = sSQL & " 					WHERE Event='J' AND Div IN ('EM','EW') AND ( Sent_Notice IS NULL OR Sent_Notice='N') ) je"
	sSQL = sSQL & " ON je.SkiYearID=row.SkiYearID AND je.MemberID=row.MemberID"
	sSQL = sSQL & " LEFT JOIN" 
	sSQL = sSQL & " 		( SELECT MemberID, Event, Div, l.SkiYearID, First_Instance AS O_First_Instance_E" 	
	sSQL = sSQL & " 				FROM " & EquivLevel10TableName & " l"
	sSQL = sSQL & " 					WHERE Event='O' AND Div IN ('EM','EW') AND ( Sent_Notice IS NULL OR Sent_Notice='N') ) oe"
	sSQL = sSQL & " ON oe.SkiYearID=row.SkiYearID AND oe.MemberID=row.MemberID"

	sSQL = sSQL & " WHERE Email IS NOT NULL AND LEN(Email)>2 AND Email<>'unsubscribe'"
' 	sSQL = sSQL & "              AND row.skiyear-datepart(yyyy,mt.birthdate)-1 >= 18" 
	
WriteDebugSQL(sSQL)
	
	' response.write(sSQL)
	' response.end





  '	-- Set up to produce the emails.
	SetupEmailService

	objMessage.Subject="Welcome to AWSA Level 10 Status"
	' objMessage.From = """AWSA President"" <awsa_prez@usawaterski.org>"
	objMessage.From = """AWSA President"" <competition@usawaterski.org>"
	objMessage.BCC = """Jeff Surdej"" <j_surdej@yahoo.com>; ""Kelvin Kelm"" <kelvin@sportpixusa.net>; ""Mark Crone"" <cronemarka@gmail.com>"
	' objMessage.BCC = ""

	rs.open sSQL, sConnectionToTRATable, 0, 1


'	---   Process the resulting record-set sending a customized email note to each such qualifier ---

	Dim nNewLevel10
	nNewLevel10 = 0
	
	IF NOT rs.eof THEN
		
'		rs.MoveFirst
		DO Until rs.eof

'
'		****  NOTE EDITED - SEE ACTUAL HTML - Here below is a sample of what the plain text version looks like ...
'
'		To:    [FirstNm] [LastNm]
'		Re:    AWSA Elite Status in [EliteDiv] [Event]
'		Date:  [long format date]
'
'		Dear [Skier],
'
'		Congratulations on your recent performance.  You are now qualified for Level 10,
'		the highest rating in 3 Event water skiing.  As a result of having met this threshold,
'		for the remainder of the current ski year you must compete only in Elite Divisions
'		(OM/OW or MM/MW whichever is applicable)."  
'	
'		You have qualified for Level 10 in the events listed below.

'		Open Slalom.........9/1/2017
'		Masters Tricks......9/10/2017

'		Simply stated, you are truly among the elite in the sport and one of
'		AWSA's most outstanding water ski athletes.</p>

'		If you have any questions, please check out Level 10 FAQ

'		A very heartfelt congratulations 
'	
'		Jeff Surdej
'		President
'		American Water Ski Association
'

			Dim Email, FirstName, LastName, MemberID, SkiYearID
			Dim S_First_Instance_E, S_First_Instance_S, T_First_Instance_E, T_First_Instance_S, J_First_Instance_E, O_First_Instance_E
			Dim eBody

			' Email = "cronemarka@gmail.com"		 		
			Email = rs("Email")
			FirstName = rs("FirstName")
			LastName = rs("LastName")
			MemberID = rs("MemberID")
			SkiYearID = rs("SkiYearID")

			' -- As of 3/2018 there were only 6 Elite division groupings Slalom & Trick (E and S), Jump & Overall (E only)  --
			S_First_Instance_E = rs("S_First_Instance_E")
			S_First_Instance_S = rs("S_First_Instance_S")
			T_First_Instance_E = rs("T_First_Instance_E")
			' T_First_Instance_S = rs("S_First_Instance_S")
			T_First_Instance_S = rs("T_First_Instance_S")
			J_First_Instance_E = rs("J_First_Instance_E")		
			O_First_Instance_E = rs("O_First_Instance_E")			


			objMessage.To = """" & FirstName & " " & LastName & """ <" & Email & ">"	
			
			WriteDebugSQL ( "Level 10 Email being Generated TO: " & objMessage.To & " -- CC: " & objMessage.cc & " -- BCC: " & objMessage.bcc & " -- From: " & objMessage.from )

			eBody = "<html><head><title>Welcome to AWSA Level 10 Status</title></head>"
			eBody = eBody & "<body><basefont face=""arial,sans-serif,helvetica,verdana,tahoma"" color=""#000000"" size=""2"">"

			eBody = eBody & "<div style=""border: 2px solid blue;"
			eBody = eBody & " padding: 25px;"
			eBody = eBody & " margin: 10;"
'			eBody = eBody & " text-align: justify;"
			eBody = eBody & " line-height: 23px;"
			eBody = eBody & " color: #070707;"
			eBody = eBody & " font-size: 18px"">"
			
			eBody = eBody & "<br>"
			eBody = eBody & "<br><p>Date:&nbsp; " & FormatDateTime(date(),1) 
			eBody = eBody & "<br>"
			eBody = eBody & "To:&nbsp;&nbsp;&nbsp;&nbsp; " & FirstName & " " & LastName
			eBody = eBody & "<br>Re:&nbsp;&nbsp;&nbsp;&nbsp; AWSA Level 10 Status</p>"


'			Line below is header for when running in debug mode going to developers -- this documents who would go to.
'			eBody = eBody & "<p>HTML eMail to:  """ & rs("FirstName") & " " & rs("LastName") & """ &lt;" & rs("Email") & "&gt;</p>"

			eBody = eBody & "<p>Dear " & FirstName & ",</p>"

			eBody = eBody & "<p>Congratulations on your recent performance.  You are now qualified for Level 10," 
			eBody = eBody & " the highest rating in 3 Event water skiing.  As a result of having met this threshold,"
			eBody = eBody & " for the remainder of the current ski year you must compete only in Elite Divisions"
			eBody = eBody & " (OM/OW or MM/MW whichever is applicable).</p>" 
			
			eBody = eBody & "<p>You have qualified for Level 10 in the events listed below.</p>"

		 	IF NOT(IsNull(S_First_Instance_E)) THEN eBody = eBody & "<br>Open Slalom........" & S_First_Instance_E
		 	IF NOT(IsNull(S_First_Instance_S)) THEN eBody = eBody & "<br>Masters Slalom....." & S_First_Instance_S
		 	IF NOT(IsNull(T_First_Instance_E)) THEN eBody = eBody & "<br>Open Trick........." & T_First_Instance_E
		 	IF NOT(IsNull(T_First_Instance_S)) THEN eBody = eBody & "<br>Masters Trick......" & T_First_Instance_S
		 	IF NOT(IsNull(J_First_Instance_E)) THEN eBody = eBody & "<br>Open Jump.........." & J_First_Instance_E
		 	IF NOT(IsNull(O_First_Instance_E)) THEN eBody = eBody & "<br>Open Overall......." & O_First_Instance_E
					
			eBody = eBody & " <br><p>You are truly among the elite in the sport and one of AWSA's most outstanding"
			eBody = eBody & " water ski athletes.</p>"

			eBody = eBody & "<p>If you have any questions, please check out <a href=http://www.usawaterski.org/pages/divisions/3event/Level10Cheatsheet.htm>Level 10 FAQ</a></p>"

			eBody = eBody & "<p><br>A very heartfelt congratulations !</p>"

			eBody = eBody & "<p>Jeff Surdej<br>President,<br>American Water Ski Association</p>"
			eBody = eBody & "<br>"

			dim USAWS_Logo
			USAWS_Logo ="http://www.usawaterski.org/rankings/images/logos/usawslogo_no_sub.jpg"			
    	eBody = eBody & "<div style='width:100%; text-align:center; font-size:8pt; margin:0px 0px 0px 0px;'>A Service of</div>" 
    	eBody = eBody & "<div style='width:100%; text-align:center; margin:10px 0px 0px 0px;'><img src=" & USAWS_Logo & " style='width:100px;'></div>" 
    	eBody = eBody & "<div style='width:100%; text-align:center; font-size:8pt; font-style:bold; margin:10px 0px 0px 0px;'>180 Holy Cow Rd<br>Polk City, FL 33883</div>" 

			eBody = eBody & "</div></body></html>"

			objMessage.HTMLBody = eBody
			objMessage.Send
			nNewLevel10 = nNewLevel10 + 1



			' ---  Now update Level 10 table with a "Y" showing that an email notice has been sent.
			' sSQL = "UPDATE " & EquivLevel10TableName & " SET Sent_Notice='Y' WHERE MemberID = '" & MemberID & "' AND SkiYearID = '" & SkiYearID	& "'"
			' Con.Execute(sSQL)
			
			rs.MoveNext

		LOOP
		
	END IF

	rs.close
	set objMessage = Nothing





		' -- Updates Level 10 table to Sent_Notice=Y unless after L10 COD - If N then L10 is not shown on rankings --  
		sSQL = "UPDATE l" 
		sSQL = sSQL + " SET Sent_Notice='Y'"
		sSQL = sSQL + " FROM " & EquivLevel10TableName & " l"
		sSQL = sSQL + " JOIN " & SkiYearTableName & " sy ON sy.skiyearid = l.skiyearid" 
		sSQL = sSQL + " WHERE sy.defaultyear = 1"
		sSQL = sSQL + " 	AND First_Instance<=Level10_EndDate"
	
		Con.Execute(sSQL)



END IF


' -- Bottom of IF condition for sProcessingYear = 1 THEN
END IF


' --------------------------------------------------------------------------------------------------
' -------------------  END of NEW SQL-Based Ranking LEVEL logic  -----------------------------------
' --------------------------------------------------------------------------------------------------







' MUST run this section as this is the cue to STOP processsing. - WHAT DOES THIS MEAN?






' ----   Finally we reset (turn off) the Recalculation Underway flag for the Ski Year just processsed.  ----

sSQL = "UPDATE " & SkiYearTableName & " set RecalcUnderway = 0 WHERE SkiYearID = " & sProcessingYear






' IF sProcessingYear <> 21 then 
Con.Execute(sSQL)
' End IF

IF Request("Equival") <> "" THEN 
	Response.write(" Done !!<br>")
	Response.Flush
END IF





' ---------------------------------------------------------------------------------------------
' ---------------  Determines if all desired records have been processed   --------------------
' ---------------------------------------------------------------------------------------------

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

IF Request("Equival") <> "" THEN 
	FinishProgress
END IF


WriteLog(date() &"  "& time() &"  Ranking Recalculations Completed Successfully.")



' ---------------------------------------------------------------------------------
' **************    Bottom of Rankings Recalculation Conditional    ***************
' ---------------------------------------------------------------------------------



END IF

timeNow = milliDif()







' ----------------------------------------------------------------------------------
' ---------      Email Results recap, ONLY if no Equival request value     ---------
' ----------------------------------------------------------------------------------

' ----- Hence this emails ONLY for the nightly scheduled unattended run.
' ----- Otherwise if run manually, then recap will appear on user's screen
' ----- as process runs, and no need to e-mail -- copy/paste/save if desired.

IF Request("Equival") = "" THEN 

	SetupEmailService
	objMessage.Subject = "USAWS Rankings - Nightly Run Recap"
	' objMessage.From = "USAWS.Rankings@USAWaterSki.ORG"
	' objMessage.From = """AWSA Rankings"" <competition@usawaterski.org>"
	objMessage.From = "competition@usawaterski.org"
	objMessage.To = EMailToWho

	IF Err.Number = 0 THEN

		IF Len(DupMemList) > 0 THEN DupMemList = "!! Duplicate Person ID's encountered: " & DupMemList & vbCrLf

		objMessage.TextBody="This is an automated message to report that the Nightly ePolk Run" & vbCrLf _
      & "has been successfully completed.  Member Extract & Update Recap:" & vbCrLf & vbCrLf _
      & Formatnumber(nHQExt,0) & " Member data rows supplied from HQ Server" & vbCrLf & DupMemList _
      & Formatnumber(nLocal,0) & " Member rows found in ePolk Server Table" & vbCrLf _
      & Formatnumber(nUpdates,0) & " Member rows updated with new Data" & vbCrLf _
      & Formatnumber(nInserts,0) & " New Member rows added" & vbCrLf _
      & Formatnumber(nDeletes,0) & " Old Member rows deleted" & vbCrLf _
      & Formatnumber(nLocal+nInserts-nDeletes,0) & " Member rows now in ePolk Server Table" & vbCrLf _
  		& vbCrLf & Formatnumber(nConsMems,0) & " Consolidated Member Rows Read from HQ Table" _
  		& vbCrLf & Formatnumber(nConsHits,0) & " ConsMems Rows hit to Updated Member Table" _
  		& vbCrLf & Formatnumber(nRnkUpdts,0) & " Ranking Table Entries updated to Cons Mbr IDs" _
  		& vbCrLf & Formatnumber(nOffUpdts,0) & " Officials Table Entries updated to Cons Mbr IDs" _
  		& vbCrLf & Formatnumber(nColUpdts,0) & " Collegiate Table Entries updated to Cons Mbr IDs" _
  		& vbCrLf & Formatnumber(nRegUpdts,0) & " Registration Table Entries updated to Cons Mbr IDs" _
  		& vbCrLf & Formatnumber(nNewElites,0) & " New Elite Status Welcome eMails generated" _
  		& vbCrLf & Formatnumber(nNewLevel10,0) & " New Level 10 Status Welcome eMails generated" _
  		& vbCrLf & Formatnumber(nPTFeMails,0) & " Post-Tournament Followup eMails generated" _
      & vbCrLf & vbCrLf & "Default Ski Year plus 12 Month Rankings have been Recalculated" _
      & vbCrLf & vbCrLf & "Overall Process ... Start: " & StartTime & ", Finish: " & Time()



	ELSE
		objMessage.TextBody="This is an automated message to report that the Nightly Run has been " _
      & "completed but there were errors detected during the processing. " _
      & "<br>&nbsp;<br><B>Page Error Object</B><BR>" _
      & "Error Number: " & response.write(Err.Number) & " <BR> " _
      & "Description: " & response.write(Err.Description) & " <BR>	" _
      & "Source: " & response.write(Err.Source) & " <BR> " _
      & "Line Number: " & response.write(Err.Line) & " <BR>"
	END IF

WriteDebugSQL ( "Line 4320 objMessage.TextBody - "&objMessage.TextBody)

	objMessage.Send
	set objMessage = nothing

	IF ReCalc12 = "Y" THEN Response.Redirect("/rankings/QualifyRecalc.asp")

END IF

WriteDebugSQL ( "Line 4329 End of Equival.asp - Right before display of browser results")

%>

<br>Nightly Update / Ranking Recalc Completed.<br>&nbsp;<br>
Total Time to Process: <%=elapsedpretty(timeNow - timeTHEN)%><br>&nbsp;<br>

</BODY>
</HTML>


<% IF ReCalc12 = "Y" THEN %>

<a href="/rankings/QualifyRecalc.asp">Proceed to Qualification Recalculation</a>

<% ELSE %>

<a href="/rankings/defaultHQ.asp">Return to Main Menu</a>

<% END IF %>

              




 