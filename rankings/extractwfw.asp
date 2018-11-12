<%
IF Session("adminmenulevel")<10 THEN Response.Redirect "DefaultHQ.asp?process=login" 
IF Session("UploadMode") <> "Zip" AND Session("UploadMode") <> "Rpt" THEN Response.Redirect "DefaultHQ.asp?process=login"
%>

<!--#include file="settingsHQ.asp"-->

<%

WriteIndexPageHeader

%>
		<table border="2" cellspacing="1" cellpadding="1">

		<tr>
			<td>&nbsp;&nbsp;&nbsp;DLA</td>

			<td valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" Size="2">
<%

Dim strTourID, strTourZip, strTourFldr
Dim sSQL, strFile, strStatus, strAction
Dim TSanction, TName, TDateE, TStatus
Dim TEventSlalom, TEventTrick, TEventJump
Dim objFSO, objZip, objRS, eMailSubj, eMailFrom, eMailTo, eMailCC, eMailReplyTo, SeedRep, Owner
Dim FoundNewWSP, WSPFileName

Dim PTF_SBK, PTF_WSP, PTF_TS, PTF_OD, PTF_BT, PTF_JT
Dim PTF_CS, PTF_CJ, PTF_SD, PTF_TU, PTF_HD, PTF_TNY, nFilSto

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Set objZip = Server.CreateObject("SoftComplex.Zip")
Set objRS = Server.CreateObject("ADODB.recordset")
OpenConSanUpd
			
'	Startup -- create some folder and file name variables

strTourID = UCase(left(Session("strTourID"),6))
strTourZip = Session("strTourZip")
strTourFldr = PathtoScratch & "\" & Session("strTourID")
WriteDebugSQL ("ExtractWfw.asp: TourID: " & strTourID & ", Tourzip=" & strTourZip & ", TourFldr=" & strTourFldr)

'	============= Get SWIFT entry into objRS answerset & pull TD name & eMail adrs

eMailTo = ""
sSQL = "Select top 1 ST.TSanction, ST.TName, ST.TDateE, ST.TDirName, ST.TDirEMail," 
sSQL = sSQL & " CJ.CJudgName, CJ.CJudgEMail, CC.CScorName, CC.CScorEMail,"
sSQL = sSQL & " ST.TEventSlalom, ST.TEventTrick, ST.TEventJump, ST.TStatus, TSanType,"
sSQL = sSQL & " RT.SClassN, RT.SClassC, RT.SClassE, RT.SClassL, RT.SClassR, RT.SClassCash,"
sSQL = sSQL & " RT.TClassN, RT.TClassC, RT.TClassE, RT.TClassL, RT.TClassR, RT.TClassCash,"
sSQL = sSQL & " RT.JClassN, RT.JClassC, RT.JClassE, RT.JClassL, RT.JClassR, RT.JClassCash,"
sSQL = sSQL & " RT.USClassN, RT.USClassC, RT.UTClassN, RT.UTClassC, RT.UJClassN, RT.UJClassC,"
sSQL = sSQL & " Coalesce(PT.PTF_SBK,-1) as PTF_SBK, PT.PTF_WSP, PT.PTF_TS, PT.PTF_OD, PT.PTF_BT,"
sSQL = sSQL & " PT.PTF_JT, PT.PTF_CS, PT.PTF_CJ, PT.PTF_SD, PT.PTF_TU, PT.PTF_HD, PT.PTF_TNY"
sSQL = sSQL & " FROM " & SanctionTableName & " ST LEFT JOIN " & TRegSetupTableName
sSQL = sSQL & " RT on RT.TournAppID = ST.TournAppID LEFT JOIN " & PostTourTableName
sSQL = sSQL & " PT on PT.TournAppID = ST.TournAppID LEFT JOIN (Select '" & strTourID
sSQL = sSQL & "' as TournAppID, FirstName + ' ' + LastName as CJudgName, Email as CJudgEMail"
sSQL = sSQL & " FROM " & MemberTablename & " Where patindex('%@%',Email) > 0 and"
sSQL = sSQL & " PersonID in (Select Cast(case when len(CJudgePID)<9 then CJudgePID"
sSQL = sSQL & " else right(CJudgePID,8) end as integer) FROM " & TRegSetupTableName 
sSQL = sSQL & " WHERE TournAppID = '" & strTourID & "' and isnumeric(CJudgePID) = 1))"
sSQL = sSQL & " CJ on CJ.TournAppID = ST.TournAppID LEFT JOIN (Select '" & strTourID
sSQL = sSQL & "' as TournAppID, FirstName + ' ' + LastName as CScorName, Email as CScorEMail"
sSQL = sSQL & " FROM " & MemberTablename & " Where patindex('%@%',Email) > 0 and"
sSQL = sSQL & " PersonID in (Select Cast(case when len(CScorePID)<9 then CScorePID"
sSQL = sSQL & " else right(CScorePID,8) end as integer) FROM " & TRegSetupTableName 
sSQL = sSQL & " WHERE TournAppID = '" & strTourID & "' and isnumeric(CScorePID) = 1))"
sSQL = sSQL & " CC on CC.TournAppID = ST.TournAppID where upper(ST.TournAppID) = '"
sSQL = sSQL & strTourID & "'"

WriteDebugSQL ("ExtractWfw.asp: Get Swift Entry: " & SanctionTableName)

objRS.open sSQL, sConnectionToSanctionTable, 3, 3
If objRS.EOF THEN
	strTStatus = -1
    WriteDebugSQL ("ExtractWfw.asp: Failed to retrieve Swift Entry")
ELSE 
	strTStatus = objRS("TStatus")
	TSanction = objRS("TSanction")
	TName = objRS("TName")
	TDateE = Replace(FormatDateTime(objRS("TDateE"),2),"/","-")
	IF Mid(TDateE,2,1) = "-" THEN TDateE = "0" & TDateE
	IF Mid(TDateE,5,1) = "-" THEN TDateE = Left(TDateE,3) & "0" & Right(TDateE,6)
	IF Session("strTourDate") <> TDateE THEN 
		TDateE = "<font color=red>" & TDateE & "</font>"
	END IF

	IF len(objRS("TDirEMail")) > 0 THEN
		eMailTo = """" & objRS("TDirName") & """ <" & objRS("TDirEMail") & ">"
	END IF

	IF len(objRS("CJudgEmail")) > 0 and instr(eMailTo,objRS("CJudgName")) = 0 THEN
		IF len(eMailTo) > 0 THEN eMailTo = eMailTo & "; "
		eMailTo = eMailTo & """" & objRS("CJudgName") & """ <" & objRS("CJudgEmail") & ">"
	END IF

	IF objRS("TEventSlalom") = True or objRS("SClassC") > 0 or objRS("SClassE") > 0 or objRS("SClassL") > 0 or objRS("SClassR") > 0 or objRS("USClassC") > 0 THEN TEventSlalom = True: else TEventSlalom = False
	IF objRS("TEventTrick") = True or objRS("TClassC") > 0 or objRS("TClassE") > 0 or objRS("TClassL") > 0 or objRS("TClassR") > 0 or objRS("UTClassC") > 0 THEN TEventTrick = True: else TEventTrick = False
	IF objRS("TEventJump") = True or objRS("JClassC") > 0 or objRS("JClassE") > 0 or objRS("JClassL") > 0 or objRS("JClassR") > 0 or objRS("UJClassC") > 0 THEN TEventJump = True: else TEventJump = False

	IF objRS("PTF_SBK") > -1 THEN

		'	Pick up Existing PTF Flag Values from PostTourn Table Entry
		
		PTF_SBK = objRS("PTF_SBK")
		PTF_WSP = objRS("PTF_WSP")
		PTF_TS = objRS("PTF_TS")
		PTF_OD = objRS("PTF_OD")
		PTF_BT = objRS("PTF_BT")
		PTF_JT = objRS("PTF_JT")
		PTF_CS = objRS("PTF_CS")
		PTF_CJ = objRS("PTF_CJ")
		PTF_SD = objRS("PTF_SD")
		PTF_TU = objRS("PTF_TU")
		PTF_HD = objRS("PTF_HD")
		PTF_TNY = objRS("PTF_TNY")

	ELSE

		'	Insert new row into PostTourn table if not already there, then
		'	Supply default zero values for all flags.  Any 3's will come below.
		
		sSQL = "Insert Into " & PostTourTableName & " Values ('" & strTourID 
		sSQL = sSQL & "', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)"
		ConSanUpd.Execute(sSQL)

		PTF_SBK = 0: PTF_WSP = 0: PTF_TS = 0: PTF_OD = 0: PTF_BT = 0: PTF_JT = 0
		PTF_CS = 0: PTF_CJ = 0: PTF_SD = 0: PTF_TU = 0: PTF_HD = 0: PTF_TNY = 0

	END IF

END IF


' Begin by listing the tournament and incoming file specifics.
' Note -- we have WWPARM data only if we're handling a ZIP file.

	%>
		<p><%=Session("strFileInfo")%></p>
	
		<% IF Session("") = "Zip" THEN %>
			<p><b>Upload:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<%=Session("strTourID")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<%=Session("strTourName")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<%=Session("strTourDate")%><br>
		<% ELSE %>
			<p><b>
		<% END IF %>
			&nbsp;SWIFT:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<%=TSanction%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<%=TName%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<%=TDateE%></b></p>	
	<%


'	First create temporary working folder for this tournament (in the scratch area).

IF objFSO.FolderExists(strTourFldr) = false THEN
	objFSO.CreateFolder(strTourFldr)
END IF


'	Then Check to see if Post Tournament Archive Zip file already 
'	exists.  If not, then create it.  Otherwise extract from the 
'	existing Zip Archive into the working folder, then report as 
'	updating, and cite number of files found in existing zip.

IF objFSO.FileExists(strTourZip) = false THEN

	objZip.New(strTourZip)
	tmpFiles = objZip.Save
	%><p>Tournament Archive for <%=Session("strTourID")%> Created.</p><%
ELSE
	objZip.Open(strTourZip)
	objZip.Read
	objZip.DestDirectory = strTourFldr
	tmpFiles = objZip.UnZip
	%><p>Tournament Archive for <%=Session("strTourID")%> 
		being <b>Updated</b>	(<%=tmpFiles%>).</p><%
END IF


'	Now produce headings for the Recap Table that we're going to build.


	%>
			</td>
			<td>&nbsp;&nbsp;&nbsp;</td>
		</tr>
	
		<tr>
			<td>&nbsp;&nbsp;&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;&nbsp;&nbsp;</td>
		</tr>

		<tr><td>&nbsp;&nbsp;&nbsp;</td>
			<td><TABLE class="innertable" width=95% align=center>
				<tr>
		    	<th ALIGN="Left"><font color="#FFFFFF" size=2 face="arial"><b>&nbsp;File Name&nbsp;</b></font></th>
		    	<th ALIGN="Left"><font color="#FFFFFF" size=2 face="arial"><b>&nbsp;Explanation&nbsp;</b></font></th>
		    	<th ALIGN="Center"><font color="#FFFFFF" size=2 face="arial"><b>&nbsp;Incoming Status&nbsp;</b></font></th>
		    	<th ALIGN="Center"><font color="#FFFFFF" size=2 face="arial"><b>&nbsp;Action&nbsp;</b></font></th>
		    </tr>
	<%


'	Final setup.  Initialize strMissing string as empty, 

objZip.Open(Session("strZipPath"))
objZip.Read
objZip.DestDirectory = strTourFldr
strMissing = ""


' ===================================================================
'	Now we begin itemizing the individual post-tournament report files.
'	We do these individually so that they can be handled uniquely
'	where required, and then reported on separately.
'
'	Each file is examined/extracted/stored using a standard subroutine.  
'	For each expected standard file, if we're processing a ZIP file and 
'	that file is actually stored as new -- or if we're processing a single 
'	report file and that is the file being uploaded -- then we store it 
'	in the final zip, and depending on the file, also possibly store the 
'	uncompressed version in a specific folder elsewhere too.
'
'	If present, note status.  If not incoming, note presence in existing
'	Archive ZIP file, plus flags previously posted by HQ in S_PostTourn.
' ===================================================================


'	============ Full Scorebook file

strFileName = Session("strTourID") & ".SBK"
strDescription = "Full Scorebook Report"
ExtractFile strFileName, strTourFldr, 600
IF objFSO.FileExists (strTourFldr & "\" & strFileName) THEN
	IF Instr(strAction, "Stored") > 0 THEN
		PTF_SBK = 1
		tmpFiles = objZip.ZipFilesTo(strTourFldr&"\"&strFileName, strTourZip)
	ELSEIF strAction = "- -" AND PTF_SBK = 1 THEN
		strAction = "Prev Rcvd"
	END IF
	objFSO.DeleteFile (strTourFldr & "\" & strFileName)
ELSEIF PTF_SBK = 2 THEN
	strAction = "<font color=green><b>Posted@HQ</b></font>"
ELSE
	strMissing = strMissing & vbCRLF & "   " & strDescription & space(32-len(strDescription))
	strMissing = strMissing & "( " & strFileName & " )"
END IF
DisplayFile


'	============ WSP file

strFileName = Session("strTourID") & ".WSP"
strDescription = "Rankings Data File (Scores)"
ExtractFile strFileName, strTourFldr, 350
IF instr(strAction,"Stored") > 0 THEN 

	IF mid(strTourID,3,1) = "U" THEN

		' If NCWSA file, New section -- temporary -- fall 2009.  
		' Copy and classify Team Status on each skier record.  Begin with 
		' TeamStat set as "A", until we find a Male Record following a Female 
		' record, then for that record and thereafter we set TeamStat as "B"
		' File is processed from xxxxxx.WSP into xxxxxx.NEW

		strTeamStat = "A": strLastGender = "M"
		Set objWSP = objFSO.OpenTextFile (strTourFldr & "\" & strFileName, 1, 0)
		Set objNewWSP = objFSO.CreateTextFile (strTourFldr & "\" & Session("strTourID") & ".NEW", true) 
		strWSPLine = objWSP.ReadLine
		objNewWSP.WriteLine (strWSPLine)
		
		DO WHILE NOT objWSP.AtEndOfStream
			strWSPLine = objWSP.ReadLine
			I1 = instr(strWSPLine,",")
			IF I1>0 THEN I2 = instr(I1+1,strWSPLine,","): ELSE I2 = 0
			IF I2>0 THEN I3 = instr(I2+1,strWSPLine,","): ELSE I3 = 0
			IF I3>0 THEN I4 = instr(I3+1,strWSPLine,","): ELSE I4 = 0
			IF I4>0 THEN I5 = instr(I4+1,strWSPLine,","): ELSE I5 = 0
			IF I5>0 THEN I6 = instr(I5+1,strWSPLine,","): ELSE I6 = 0
			IF I6>0 THEN I7 = instr(I6+1,strWSPLine,","): ELSE I7 = 0
			IF I7>0 THEN I8 = instr(I7+1,strWSPLine,","): ELSE I8 = 0
			IF I8>0 THEN I9 = instr(I8+1,strWSPLine,","): ELSE I9 = 0
			IF I9>0 THEN
				IF strLastGender = "F" AND Mid(strWSPLine,I4+2,1) = "M" THEN strTeamStat = "B"
				strLastGender = Mid(strWSPLine,I4+2,1)
				IF I9-I8 < 7 THEN
          strWSPLine = Left(strWSPLine,I9-2) & "/" & strTeamStat & Mid(strWSPLine,I9-1)
				END IF
			END IF
			objNewWSP.WriteLine (strWSPLine)
		LOOP
		
		objWSP.Close
		Set objWSP = Nothing
		objNewWSP.Close
		Set objNewWSP = Nothing

		' NCWSA Team Status Modification processing is completed -- now copy revised xxxxxx.NEW 
		' file back on top of xxxxxx.WSP, then delete the xxxxxx.NEW file -- leaving revised file
		' under original name xxxxxx.WSP, for subsequently processing through VerifyNewWSP.asp
		
		objFSO.CopyFile strTourFldr & "\" & Session("strTourID") & ".NEW", strTourFldr & "\" & strFileName, TRUE
		objFSO.DeleteFile (strTourFldr & "\" & Session("strTourID") & ".NEW")

	END IF

	objFSO.CopyFile strTourFldr & "\" & strFileName, PathtoRawWSPs & "\", TRUE
	WSPFileName = strFilename
	FoundNewWSP = True
ELSE
	FoundNewWSP = False
END IF


' Get submitter Name and eMail ID from WSP Header for eMail confirmation,
'	Otherwise use Chief Scorer name and eMail address obtained from Sanction.
IF objFSO.FileExists (strTourFldr & "\" & strFileName) THEN
    WriteDebugSQL ("ExtractWfw.asp: Sending emails")
	
    IF Instr(strAction, "Stored") > 0 THEN
		tmpFiles = objZip.ZipFilesTo(strTourFldr&"\"&strFileName, strTourZip)
		PTF_WSP = 1
	ELSEIF strAction = "- -" AND PTF_WSP = 1 THEN
		strAction = "Prev Rcvd"
	END IF
	Set objWSP = objFSO.OpenTextFile(strTourFldr & "\" & strFileName, 1, 0)
	strWSPHdr = objWSP.ReadLine
	objWSP.Close
	Set objWSP = Nothing
	I1 = instr(strWSPHdr,",")
	IF I1>0 THEN I2 = instr(I1+1,strWSPHdr,","): ELSE I2 = 0
	IF I2>0 THEN I3 = instr(I2+1,strWSPHdr,","): ELSE I3 = 0
	IF I3>0 THEN I4 = instr(I3+1,strWSPHdr,","): ELSE I4 = 0
	IF I4>0 THEN I5 = instr(I4+1,strWSPHdr,","): ELSE I5 = 0
	IF I5>0 THEN I6 = instr(I5+1,strWSPHdr,","): ELSE I6 = 0
	IF I6>0 THEN I7 = instr(I6+1,strWSPHdr,","): ELSE I7 = 0
	IF I7>0 THEN I8 = instr(I7+1,strWSPHdr,","): ELSE I8 = 0
	IF I8>0 THEN I9 = instr(I8+1,strWSPHdr,","): ELSE I9 = 0
	IF I9>0 THEN I10 = instr(I9+1,strWSPHdr,","): ELSE I10 = 0
	IF I10>0 THEN I11 = instr(I10+1,strWSPHdr,","): ELSE I11 = 0
	IF I11>0 THEN I12 = instr(I11+1,strWSPHdr,","): ELSE I12 = 0
	IF I12>0 THEN I13 = instr(I12+1,strWSPHdr,","): ELSE I13 = 0
	IF I13 > 0 and Mid(strWSPHdr, I10+1, 2) <> """""" THEN
		IF len(eMailTo) > 0 THEN eMailTo = eMailTo & "; "
		eMailTo = eMailTo & """" & Mid(strWSPHdr, I11+2, 1) & Lcase(Mid(strWSPHdr, I11+3, I12-I11-4))
		eMailTo = eMailTo & " " & Mid(strWSPHdr, I10+2, 1) & Lcase(Mid(strWSPHdr, I10+3, I11-I10-3))
		eMailTo = eMailTo & """ <" & Lcase(Mid(strWSPHdr, I12+2, I13-I12-3)) & ">"
	ELSE
		IF len(eMailTo) > 0 THEN eMailTo = eMailTo & "; "
		eMailTo = eMailTo & """" & objRS("CScorName") & """ <" & objRS("CScorEmail") & ">"
	END IF	
	objFSO.DeleteFile (strTourFldr & "\" & strFileName)
ELSE
	strMissing = strMissing & vbCRLF & "   " & strDescription & space(32-len(strDescription))
	strMissing = strMissing & "( " & strFileName & " )"
	IF len(objRS("CScorEmail")) > 0 and instr(eMailTo,objRS("CScorName")) = 0 THEN
		IF len(eMailTo) > 0 THEN eMailTo = eMailTo & "; "
		eMailTo = eMailTo & """" & objRS("CScorName") & """ <" & objRS("CScorEmail") & ">"
	END IF
END IF
WriteDebugSQL ("ExtractWfw.asp: eMailTo: " & eMailTo)

WriteDebugSQL ("ExtractWfw.asp: DisplayFile ")

DisplayFile


'	============ Tournament Summary --

nFilSt0 = 0
strFileName = strTourID & "TS.PRN"
strDescription = "Tournament Summary Report"
WriteDebugSQL ("ExtractWfw.asp: " & strFileName & " " & strDescription)

On Error Resume Next
    ExtractFile strFileName, strTourFldr, 550
    If Err.Number <> 0 Then
        WriteDebugSQL ("ExtractWfw.asp: Error sending email: Err.Number=" & Err.Number & " Message="  & Err.Description )
        On Error Goto 0 ' But don't let other errors hide!
    End If
IF objFSO.FileExists (strTourFldr & "\" & strFileName) THEN
	nFilSto = nFilSto + 1
	IF Instr(strAction, "Stored") > 0 THEN
		tmpFiles = objZip.ZipFilesTo(strTourFldr&"\"&strFileName, strTourZip)
	ELSEIF strAction = "- -" AND PTF_TS = 1 THEN
		strAction = "Prev Rcvd"
	END IF
	objFSO.DeleteFile (strTourFldr & "\" & strFileName)
ELSEIF PTF_TS = 2 THEN
	strAction = "<font color=green><b>Posted@HQ</b></font>"
ELSE
	strMissing = strMissing & vbCRLF & "   " & strDescription & space(32-len(strDescription))
	strMissing = strMissing & "( " & strFileName & " )"
END IF
DisplayFile

'	---	Also bring in companion WfW TS.TXT data file

strFileName = strTourID & "TS.TXT"
strDescription = "WfW Tournament Export Data File"
WriteDebugSQL ("ExtractWfw.asp: " & strFileName & " " & strDescription)

On Error Resume Next
    ExtractFile strFileName, strTourFldr, 100
    If Err.Number <> 0 Then
        WriteDebugSQL ("ExtractWfw.asp: Error sending email: Err.Number=" & Err.Number & " Message="  & Err.Description )
        On Error Goto 0 ' But don't let other errors hide!
    End If
IF objFSO.FileExists (strTourFldr & "\" & strFileName) THEN
	nFilSto = nFilSto + 1
	IF Instr(strAction, "Stored") > 0 THEN
		tmpFiles = objZip.ZipFilesTo(strTourFldr&"\"&strFileName, strTourZip)
	ELSEIF strAction = "- -" AND PTF_TS = 1 THEN
		strAction = "Prev Rcvd"
	END IF
	objFSO.DeleteFile (strTourFldr & "\" & strFileName)
ELSEIF PTF_TS = 2 THEN
	strAction = "<font color=green><b>Posted@HQ</b></font>"
ELSE
	strMissing = strMissing & vbCRLF & "   " & strDescription & space(32-len(strDescription))
	strMissing = strMissing & "( " & strFileName & " )"
END IF
DisplayFile

'	Now set PTF-TS only both present and not already set

If PTF_TS = 0 and nFilSto = 2 then PTF_TS = 1



'	============ Officials Data file

strFileName = strTourID & "OD.TXT"
strDescription = "Officials Data File (Credits)"
WriteDebugSQL ("ExtractWfw.asp: " & strFileName & " " & strDescription)

On Error Resume Next
    ExtractFile strFileName, strTourFldr, 750
    If Err.Number <> 0 Then
        WriteDebugSQL ("ExtractWfw.asp: Error sending email: Err.Number=" & Err.Number & " Message="  & Err.Description )
        On Error Goto 0 ' But don't let other errors hide!
    End If

'	If file stored, Check for chiefs and reset status if missing

IF instr(strAction,"Stored") > 0 THEN 
	Set objOff = objFSO.OpenTextFile(strTourFldr & "\" & strFileName, 1, 0)
	LinesRead = 0: ChfCd1 = 0: ChfCd2 = 0: ChfCd3 = 0: ChfCd4 = 0: ChfCd5 = 0: ChfCd6 = 0
	DO WHILE NOT objOff.AtEndOfStream
		strOffLine = objOff.ReadLine
		LinesRead = LinesRead + 1
		IF LinesRead > 4 AND len(strOffLine) > 35 THEN 
			IF left(strOffLine,5) <> "*****" THEN
				IF Len(strOffLine) >= 41 THEN 
					IF Instr(mid(ucase(strOffLine),39,3),"C") > 0 THEN ChfCd1 = 1
				END IF
				IF Len(strOffLine) >= 48 THEN 
					IF Instr(mid(ucase(strOffLine),46,3),"C") > 0 THEN ChfCd2 = 1
				END IF
				IF Len(strOffLine) >= 55 THEN 
					IF Instr(mid(ucase(strOffLine),53,3),"C") > 0 THEN ChfCd3 = 1
				END IF
				IF Len(strOffLine) >= 62 THEN 
					IF Instr(mid(ucase(strOffLine),60,3),"C") > 0 THEN ChfCd4 = 1
				END IF
				IF Len(strOffLine) >= 69 THEN 
					IF Instr(mid(ucase(strOffLine),67,3),"C") > 0 THEN ChfCd5 = 1
				END IF
				IF Len(strOffLine) >= 76 THEN 
					IF Instr(mid(ucase(strOffLine),74,3),"C") > 0 THEN ChfCd6 = 1
				END IF
			END IF
		END IF
		LOOP
	objOff.Close
	Set objOff = Nothing
	IF ChfCd1+ChfCd2+ChfCd3+ChfCd4 < 4 or (instr("ELRPAB",mid(Session("strTourID"),7,1)) > 0 and ChfCd5 = 0) THEN
		strStatus = "<font color=red>Chief Code(s) Missing</font>"
		strAction = "- -"
		objFSO.DeleteFile (strTourFldr & "\" & strFileName)
	END IF
END IF

IF instr(strAction,"Stored") > 0 THEN 
	objFSO.CopyFile strTourFldr & "\" & strFileName, PathtoHQInBox & "\", TRUE
END IF

IF objFSO.FileExists (strTourFldr & "\" & strFileName) THEN
	IF Instr(strAction, "Stored") > 0 THEN
		tmpFiles = objZip.ZipFilesTo(strTourFldr&"\"&strFileName, strTourZip)
		PTF_OD = 1
	ELSEIF strAction = "- -" AND PTF_OD = 1 THEN
		strAction = "Prev Rcvd"
	END IF
	objFSO.DeleteFile (strTourFldr & "\" & strFileName)
ELSEIF PTF_OD = 2 THEN
	strAction = "<font color=green><b>Posted@HQ</b></font>"
ELSE
	strMissing = strMissing & vbCRLF & "   " & strDescription & space(32-len(strDescription))
	strMissing = strMissing & "( " & strFileName & " )"
	IF Instr(strStatus,"Chief Code") > 0 THEN strMissing = strMissing & "  Chief Cd(s) Missing"
END IF
DisplayFile


'	============ Boat Time Tracking Report ... only if Slalom or Jump events.

IF TEventSlalom = True or TEventJump = True THEN
	strFileName = strTourID & "BT.PRN"
	strDescription = "Boat Time Tracking Report"
    WriteDebugSQL ("ExtractWfw.asp: " & strFileName & " " & strDescription)

    On Error Resume Next
        ExtractFile strFileName, strTourFldr, 500
        If Err.Number <> 0 Then
            WriteDebugSQL ("ExtractWfw.asp: Error sending email: Err.Number=" & Err.Number & " Message="  & Err.Description )
            On Error Goto 0 ' But don't let other errors hide!
        End If
	IF instr(strAction,"Stored") > 0 THEN 
		objFSO.CopyFile strTourFldr & "\" & strFileName, PathtoTiming & "\", TRUE
	END IF
	IF objFSO.FileExists (strTourFldr & "\" & strFileName) THEN
		IF Instr(strAction, "Stored") > 0 THEN
			tmpFiles = objZip.ZipFilesTo(strTourFldr&"\"&strFileName, strTourZip)
			PTF_BT = 1
		ELSEIF strAction = "- -" AND PTF_BT = 1 THEN
			strAction = "Prev Rcvd"
		END IF
		objFSO.DeleteFile (strTourFldr & "\" & strFileName)
	ELSEIF PTF_BT = 2 THEN
		strAction = "<font color=green><b>Posted@HQ</b></font>"
	ELSE
		strMissing = strMissing & vbCRLF & "   " & strDescription & space(32-len(strDescription))
		strMissing = strMissing & "( " & strFileName & " )"
	END IF
	DisplayFile
ELSE
	PTF_BT = 3
END IF


'	============ Jump Timing Data File ... only if Jump events
'	Special handling here -- Consider missing ONLY if R/C tournament

IF TEventJump = True THEN
	strFileName = strTourID & "JT.CSV"
	strDescription = "Jump Time Data File"
    WriteDebugSQL ("ExtractWfw.asp: " & strFileName & " " & strDescription)

    On Error Resume Next
        ExtractFile strFileName, strTourFldr, 500
        If Err.Number <> 0 Then
            WriteDebugSQL ("ExtractWfw.asp: Error sending email: Err.Number=" & Err.Number & " Message="  & Err.Description )
            On Error Goto 0 ' But don't let other errors hide!
        End If
	IF instr(strAction,"Stored") > 0 THEN 
		objFSO.CopyFile strTourFldr & "\" & strFileName, PathtoTiming & "\", TRUE
		IF instr("LRPAB",mid(Session("strTourID"),7,1)) > 0 THEN
			objFSO.CopyFile strTourFldr & "\" & strFileName, PathtoIWWF & "\", TRUE
			EmailToIWWF strFileName
		END IF
	END IF
	IF objFSO.FileExists (strTourFldr & "\" & strFileName) THEN
		IF Instr(strAction, "Stored") > 0 THEN
			tmpFiles = objZip.ZipFilesTo(strTourFldr&"\"&strFileName, strTourZip)
			PTF_JT = 1
		ELSEIF strAction = "- -" AND PTF_JT = 1 THEN
			strAction = "Prev Rcvd"
		END IF
		objFSO.DeleteFile (strTourFldr & "\" & strFileName)
	ELSEIF PTF_JT = 2 THEN
		strAction = "<font color=green><b>Posted@HQ</b></font>"
	ELSE
		strMissing = strMissing & vbCRLF & "   " & strDescription & space(32-len(strDescription))
		strMissing = strMissing & "( " & strFileName & " )"
	END IF
	DisplayFile
ELSE
	PTF_JT = 3
END IF


'	============ Condensed Scorebook file

strFileName = strTourID & "CS.HTM"
strDescription = "Condensed Scorebook Report"
WriteDebugSQL ("ExtractWfw.asp: " & strFileName & " " & strDescription)

On Error Resume Next
    ExtractFile strFileName, strTourFldr, 400
    If Err.Number <> 0 Then
        WriteDebugSQL ("ExtractWfw.asp: Error sending email: Err.Number=" & Err.Number & " Message="  & Err.Description )
        On Error Goto 0 ' But don't let other errors hide!
    End If
IF instr(strAction,"Stored") > 0 THEN 
	objFSO.CopyFile strTourFldr & "\" & strFileName, PathtoScoreBks & "\", TRUE
	IF instr("LRPAB",mid(Session("strTourID"),7,1)) > 0 THEN
		objFSO.CopyFile strTourFldr & "\" & strFileName, PathtoIWWF & "\", TRUE
		EmailToIWWF strFileName
	END IF
END IF
IF objFSO.FileExists (strTourFldr & "\" & strFileName) THEN
	IF Instr(strAction, "Stored") > 0 THEN
		tmpFiles = objZip.ZipFilesTo(strTourFldr&"\"&strFileName, strTourZip)
		PTF_CS = 1
	ELSEIF strAction = "- -" AND PTF_CS = 1 THEN
		strAction = "Prev Rcvd"
	END IF
	objFSO.DeleteFile (strTourFldr & "\" & strFileName)
ELSEIF PTF_CS = 2 THEN
	strAction = "<font color=green><b>Posted@HQ</b></font>"
ELSE
	strMissing = strMissing & vbCRLF & "   " & strDescription & space(32-len(strDescription))
	strMissing = strMissing & "( " & strFileName & " )"
END IF
DisplayFile


'	============ Chief Judges Report

nFilSto = 0
strFileName = strTourID & "CJ.PRN"
strDescription = "Chief Judges Tournament Report"
WriteDebugSQL ("ExtractWfw.asp: " & strFileName & " " & strDescription)

On Error Resume Next
    ExtractFile strFileName, strTourFldr, 800
    If Err.Number <> 0 Then
        WriteDebugSQL ("ExtractWfw.asp: Error sending email: Err.Number=" & Err.Number & " Message="  & Err.Description )
        On Error Goto 0 ' But don't let other errors hide!
    End If

IF objFSO.FileExists (strTourFldr & "\" & strFileName) THEN
	nFilSto = nFilSto + 1
	IF Instr(strAction, "Stored") > 0 THEN
		tmpFiles = objZip.ZipFilesTo(strTourFldr&"\"&strFileName, strTourZip)
	ELSEIF strAction = "- -" AND PTF_CJ = 1 THEN
		strAction = "Prev Rcvd"
	END IF
	objFSO.DeleteFile (strTourFldr & "\" & strFileName)
ELSEIF PTF_CJ = 2 THEN
	strAction = "<font color=green><b>Posted@HQ</b></font>"
ELSE
	strMissing = strMissing & vbCRLF & "   " & strDescription & space(32-len(strDescription))
	strMissing = strMissing & "( " & strFileName & " )"
END IF
DisplayFile

'	---	Also bring in companion WfW CJ.TXT data file -- No flags

strFileName = strTourID & "CJ.TXT"
strDescription = "Chief Judges Rept Data File"
WriteDebugSQL ("ExtractWfw.asp: " & strFileName & " " & strDescription)

On Error Resume Next
    ExtractFile strFileName, strTourFldr, 100
    If Err.Number <> 0 Then
        WriteDebugSQL ("ExtractWfw.asp: Error sending email: Err.Number=" & Err.Number & " Message="  & Err.Description )
        On Error Goto 0 ' But don't let other errors hide!
    End If

IF instr(strAction,"Stored") > 0 THEN 
	objFSO.CopyFile strTourFldr & "\" & strFileName, PathtoHQInBox & "\", TRUE
END IF

IF objFSO.FileExists (strTourFldr & "\" & strFileName) THEN
	nFilSto = nFilSto + 1
	IF Instr(strAction, "Stored") > 0 THEN
		tmpFiles = objZip.ZipFilesTo(strTourFldr&"\"&strFileName, strTourZip)
	ELSEIF strAction = "- -" AND PTF_CJ = 1 THEN
		strAction = "Prev Rcvd"
	END IF
	objFSO.DeleteFile (strTourFldr & "\" & strFileName)
ELSEIF PTF_CJ = 2 THEN
	strAction = "<font color=green><b>Posted@HQ</b></font>"
ELSE
	strMissing = strMissing & vbCRLF & "   " & strDescription & space(32-len(strDescription))
	strMissing = strMissing & "( " & strFileName & " )"
END IF
DisplayFile

'	Now set PTF-CJ only both present and not already set

If PTF_CJ = 0 and nFilSto = 2 then PTF_CJ = 1

'	============ Safety Report
nFilSto = 0
strFileName = strTourID & "SD.PRN"
strDescription = "Safety Directors Report"
WriteDebugSQL ("ExtractWfw.asp: " & strFileName & " " & strDescription)

On Error Resume Next
    ExtractFile strFileName, strTourFldr, 800
    If Err.Number <> 0 Then
        WriteDebugSQL ("ExtractWfw.asp: Error sending email: Err.Number=" & Err.Number & " Message="  & Err.Description )
        On Error Goto 0 ' But don't let other errors hide!
    End If
IF objFSO.FileExists (strTourFldr & "\" & strFileName) THEN
	nFilSto = nFilSto + 1
	IF Instr(strAction, "Stored") > 0 THEN
		tmpFiles = objZip.ZipFilesTo(strTourFldr&"\"&strFileName, strTourZip)
	ELSEIF strAction = "- -" AND PTF_SD = 1 THEN
		strAction = "Prev Rcvd"
	END IF
	objFSO.DeleteFile (strTourFldr & "\" & strFileName)
ELSEIF PTF_SD = 2 THEN
	strAction = "<font color=green><b>Posted@HQ</b></font>"
ELSE
	strMissing = strMissing & vbCRLF & "   " & strDescription & space(32-len(strDescription))
	strMissing = strMissing & "( " & strFileName & " )"
END IF
DisplayFile

'	---	Also bring in companion WfW SD.TXT data file

strFileName = strTourID & "SD.TXT"
strDescription = "Safety Directors Data File"
WriteDebugSQL ("ExtractWfw.asp: " & strFileName & " " & strDescription)

On Error Resume Next
    ExtractFile strFileName, strTourFldr, 100
    If Err.Number <> 0 Then
        WriteDebugSQL ("ExtractWfw.asp: Error sending email: Err.Number=" & Err.Number & " Message="  & Err.Description )
        On Error Goto 0 ' But don't let other errors hide!
    End If

IF instr(strAction,"Stored") > 0 THEN 
	objFSO.CopyFile strTourFldr & "\" & strFileName, PathtoHQInBox & "\", TRUE
END IF

IF objFSO.FileExists (strTourFldr & "\" & strFileName) THEN
	nFilSto = nFilSto + 1
	IF Instr(strAction, "Stored") > 0 THEN
		tmpFiles = objZip.ZipFilesTo(strTourFldr&"\"&strFileName, strTourZip)
	ELSEIF strAction = "- -" AND PTF_SD = 1 THEN
		strAction = "Prev Rcvd"
	END IF
	objFSO.DeleteFile (strTourFldr & "\" & strFileName)
ELSEIF PTF_SD = 2 THEN
	strAction = "<font color=green><b>Posted@HQ</b></font>"
ELSE
	strMissing = strMissing & vbCRLF & "   " & strDescription & space(32-len(strDescription))
	strMissing = strMissing & "( " & strFileName & " )"
END IF
DisplayFile

'	Now set PTF-SD only both present and not already set

If PTF_SD = 0 and nFilSto = 2 then PTF_SD = 1


'	============ Towboat Utilization Report

nFilSto = 0
strFileName = strTourID & "TU.PRN"
strDescription = "Towboat Utilization Report"
WriteDebugSQL ("ExtractWfw.asp: " & strFileName & " " & strDescription)

On Error Resume Next
    ExtractFile strFileName, strTourFldr, 800
    If Err.Number <> 0 Then
        WriteDebugSQL ("ExtractWfw.asp: Error sending email: Err.Number=" & Err.Number & " Message="  & Err.Description )
        On Error Goto 0 ' But don't let other errors hide!
    End If
IF instr(strAction,"Stored") > 0 THEN 
	objFSO.CopyFile strTourFldr & "\" & strFileName, PathtoHQInBox & "\", TRUE
END IF
IF objFSO.FileExists (strTourFldr & "\" & strFileName) THEN
	nFilSto = nFilSto + 1
	IF Instr(strAction, "Stored") > 0 THEN
		tmpFiles = objZip.ZipFilesTo(strTourFldr&"\"&strFileName, strTourZip)
	ELSEIF strAction = "- -" AND PTF_TU = 1 THEN
		strAction = "Prev Rcvd"
	END IF
	objFSO.DeleteFile (strTourFldr & "\" & strFileName)
ELSEIF PTF_TU = 2 THEN
	strAction = "<font color=green><b>Posted@HQ</b></font>"
ELSE
	strMissing = strMissing & vbCRLF & "   " & strDescription & space(32-len(strDescription))
	strMissing = strMissing & "( " & strFileName & " )"
END IF
DisplayFile

'	---	Also bring in companion WfW TU.TXT data file -- No flags

strFileName = strTourID & "TU.TXT"
strDescription = "Towboat Utilization Data File"
WriteDebugSQL ("ExtractWfw.asp: " & strFileName & " " & strDescription)

On Error Resume Next
    ExtractFile strFileName, strTourFldr, 100
    If Err.Number <> 0 Then
        WriteDebugSQL ("ExtractWfw.asp: Error sending email: Err.Number=" & Err.Number & " Message="  & Err.Description )
        On Error Goto 0 ' But don't let other errors hide!
    End If

IF instr(strAction,"Stored") > 0 THEN 
	objFSO.CopyFile strTourFldr & "\" & strFileName, PathtoHQInBox & "\", TRUE
END IF

IF objFSO.FileExists (strTourFldr & "\" & strFileName) THEN
	nFilSto = nFilSto + 1
	IF Instr(strAction, "Stored") > 0 THEN
		tmpFiles = objZip.ZipFilesTo(strTourFldr&"\"&strFileName, strTourZip)
	ELSEIF strAction = "- -" AND PTF_TU = 1 THEN
		strAction = "Prev Rcvd"
	END IF
	objFSO.DeleteFile (strTourFldr & "\" & strFileName)
ELSEIF PTF_TU = 2 THEN
	strAction = "<font color=green><b>Posted@HQ</b></font>"
ELSE
	strMissing = strMissing & vbCRLF & "   " & strDescription & space(32-len(strDescription))
	strMissing = strMissing & "( " & strFileName & " )"
END IF
DisplayFile

'	Now set PTF-TU only both present and not already set

If PTF_TU = 0 and nFilSto = 2 then PTF_TU = 1


'	============ Homologation Dossier -- Tech Report
'	Only expected / handled if Record Capability Tournament

IF instr("ELRPAB",mid(Session("strTourID"),7,1)) > 0 THEN
	strFileName = strTourID & "HD.TXT"
	strDescription = "Homologation Dossier"
    On Error Resume Next
        ExtractFile strFileName, strTourFldr, 6000
        If Err.Number <> 0 Then
            WriteDebugSQL ("ExtractWfw.asp: Error sending email: Err.Number=" & Err.Number & " Message="  & Err.Description )
            On Error Goto 0 ' But don't let other errors hide!
        End If

	IF instr(strAction,"Stored") > 0 and mid(Session("strTourID"),7,1) <> "E" THEN 
		objFSO.CopyFile strTourFldr & "\" & strFileName, PathtoTiming & "\", TRUE
		objFSO.CopyFile strTourFldr & "\" & strFileName, PathtoIWWF & "\", TRUE
		EmailToIWWF strFileName
	END IF

	IF objFSO.FileExists (strTourFldr & "\" & strFileName) THEN
		IF Instr(strAction, "Stored") > 0 THEN
			tmpFiles = objZip.ZipFilesTo(strTourFldr&"\"&strFileName, strTourZip)
			PTF_HD = 1
		ELSEIF strAction = "- -" AND PTF_HD = 1 THEN
			strAction = "Prev Rcvd"
		END IF
		objFSO.DeleteFile (strTourFldr & "\" & strFileName)
	ELSEIF PTF_HD = 2 THEN
		strAction = "<font color=green><b>Posted@HQ</b></font>"
	ELSE
		strMissing = strMissing & vbCRLF & "   " & strDescription & space(32-len(strDescription))
		strMissing = strMissing & "( " & strFileName & " )"
	END IF
	DisplayFile
ELSE
	PTF_HD = 3
END IF


'	============ WWPARM.TNY File

strFileName = "WWPARM.TNY"
strDescription = "Tournament Control File"
WriteDebugSQL ("ExtractWfw.asp: " & strFileName & " " & strDescription)

On Error Resume Next
    ExtractFile strFileName, strTourFldr, 60
    If Err.Number <> 0 Then
        WriteDebugSQL ("ExtractWfw.asp: Error sending email: Err.Number=" & Err.Number & " Message="  & Err.Description )
        On Error Goto 0 ' But don't let other errors hide!
    End If
IF objFSO.FileExists (strTourFldr & "\" & strFileName) THEN
	IF Instr(strAction, "Stored") > 0 THEN
		tmpFiles = objZip.ZipFilesTo(strTourFldr&"\"&strFileName, strTourZip)
		PTF_TNY = 1
	END IF
	objFSO.DeleteFile (strTourFldr & "\" & strFileName)
END IF
DisplayFile


'	======================================================
'	Done with handling incoming files.  Now we Post 
'	updated PTF flags to S_PostTourn table.  Build an
'	update "Set " query string and summarize Status Flags.
'	======================================================

strUpdt =        " Set PTF_SBK=" & PTF_SBK
strUpdt = strUpdt & ", PTF_WSP=" & PTF_WSP
strUpdt = strUpdt & ", PTF_TS=" & PTF_TS
strUpdt = strUpdt & ", PTF_OD=" & PTF_OD
strUpdt = strUpdt & ", PTF_BT=" & PTF_BT
strUpdt = strUpdt & ", PTF_JT=" & PTF_JT
strUpdt = strUpdt & ", PTF_CS=" & PTF_CS
strUpdt = strUpdt & ", PTF_CJ=" & PTF_CJ
strUpdt = strUpdt & ", PTF_SD=" & PTF_SD
strUpdt = strUpdt & ", PTF_TU=" & PTF_TU
strUpdt = strUpdt & ", PTF_HD=" & PTF_HD
strUpdt = strUpdt & ", PTF_TNY=" & PTF_TNY

sSQL = "Update " & PostTourTableName & strUpdt & " WHERE TournAppID='" 
sSQL = sSQL & ucase(left(Session("strTourID"),6)) & "'"

'	WriteDebugSQL (sSQL)

ConSanUpd.Execute(sSQL)

'	Now that we've posted the updated flags, then we determine 
'	the new TStatus value, and update that to TSchedul table

IF len(strMissing) = 0 THEN TStatus = 5: ELSE TStatus = 4

sSQL = "Update " & SanctionTableName & " Set TStatus = " & TStatus
sSQL = sSQL & " where upper(TournAppID) = '"
sSQL = sSQL & ucase(left(Session("strTourID"),6)) & "'"

ConSanUpd.Execute(sSQL)

sSQL = "Delete from Sanctions.dbo.TSchedulW "
sSQL = sSQL & " where upper(TournAppID) = '"
sSQL = sSQL & ucase(left(Session("strTourID"),6)) & "'"
ConSanUpd.Execute(sSQL)



' ================ Now generate confirmation email, but only if we have addresses

IF len(eMailTo) > 0 THEN
    WriteDebugSQL ("ExtractWfw.asp: send emailto " & eMailTo)

	eMailSubj = "Post-Tournament Reports from " & Session("strTourID") & " " & Session("strTourName") & " (" & Session("strTourDate") & ")"

	IF mid(strTourID,3,1) = "C" THEN
		Owner = """Danny LeBourgeois"" <dleboo@gmail.com>"
	ELSEIF mid(strTourID,3,1) = "M" THEN
		' Owner = """Dave Clark"" <awsatechdude@comcast.net>"
		Owner = """Michael O'Conner"" <h2oskimo@gmail.com>"
	ELSEIF mid(strTourID,3,1) = "U" THEN
		Owner = """Robert Rhyne"" <rrriii@mindspring.com>"
	ELSEIF mid(strTourID,3,1) = "E" THEN
		Owner = """Jennifer Frederick-Kelly"" <jennifer@frederickmachine.com>"
	ELSEIF mid(strTourID,3,1) = "S" THEN
		Owner = """Kirby Whetsel"" <kwhetsel@charter.net>"
	ELSEIF mid(strTourID,3,1) = "W" THEN
		Owner = """Judy Stanford"" <judy-don@sbcglobal.net>"
	ELSE
		Owner = ""
	END IF
    WriteDebugSQL ("ExtractWfw.asp: emailto Owner " & Owner)

	' eMailCC = """Dave Clark"" <awsatechdude@comcast.net>; ""Kirby Whetsel"" <kwhetsel@charter.net>"
	eMailCC = """Kirby Whetsel"" <kwhetsel@charter.net>"

	IF Session("Firstname") & Session("LastName") = "DannyLeBourgeois" THEN
		IF instr("CSM",mid(strTourID,3,1)) = 0 and len(Owner) > 0 THEN eMailCC = eMailCC & "; " & Owner
		''''eMailFrom = """Danny LeBourgeois"" <dleboo@gmail.com>"
		eMailFrom = """USA Water Ski Competition"" <dleboo@gmail.com>"
		eMailReplyTo = "dleboo@gmail.com"
		SeedRep = "Danny LeBourgeois" & vbCrLf & "AWSA South Central Seeding" & vbCrLf & "dleboo@gmail.com" & vbCrLf & "(713) 213-1779"
	ELSEIF Session("Firstname") & Session("LastName") = "RobertRhyne" THEN
		IF instr("USM",mid(strTourID,3,1)) = 0 and len(Owner) > 0 THEN eMailCC = eMailCC & "; " & Owner
		''''eMailFrom = """Robert Rhyne"" <rrriii@mindspring.com>"
		eMailFrom = """USA Water Ski Competition"" <rrriii@mindspring.com>"
		eMailReplyTo = "rrriii@mindspring.com"
		SeedRep = "Robert Rhyne" & vbCrLf & "NCWSA Seeding" & vbCrLf & "rrriii@mindspring.com" & vbCrLf & "(704) 906-7779"
	ELSEIF Session("Firstname") & Session("LastName") = "JenniferFrederick-Kelley" THEN
		IF instr("ESM",mid(strTourID,3,1)) = 0 and len(Owner) > 0 THEN eMailCC = eMailCC & "; " & Owner
		''''eMailFrom = """Jennifer Frederick-Kelly"" <jennifer@frederickmachine.com>"
		eMailFrom = """USA Water Ski Competition"" <jennifer@frederickmachine.com>"
		eMailReplyTo = "jennifer@frederickmachine.com"
		SeedRep = "Jennifer Frederick-Kelly" & vbCrLf & "AWSA East Seeding" & vbCrLf & "jennifer@frederickmachine.com" & vbCrLf & "(716) 892-1425"
	ELSEIF Session("Firstname") & Session("LastName") = "KirbyWhetsel" THEN
		IF instr("SM",mid(strTourID,3,1)) = 0 and len(Owner) > 0 THEN eMailCC = eMailCC & "; " & Owner
		''''eMailFrom = """Kirby Whetsel"" <kwhetsel@charter.net>"
		eMailFrom = """USA Water Ski Competition"" <kwhetsel@charter.net>"
		eMailReplyTo = "kwhetsel@charter.net"
		SeedRep = "Kirby Whetsel" & vbCrLf & "AWSA South Seeding" & vbCrLf & "kwhetsel@charter.net" & vbCrLf & "(931) 409-0389"
	ELSEIF Session("Firstname") & Session("LastName") = "JudyStanford" THEN
		IF instr("WSM",mid(strTourID,3,1)) = 0 and len(Owner) > 0 THEN eMailCC = eMailCC & "; " & Owner
		''''eMailFrom = """Judy Stanford"" <judy-don@sbcglobal.net>"
		eMailFrom = """USA Water Ski Competition"" <judy-don@sbcglobal.net>"
		eMailReplyTo = "judy-don@sbcglobal.net"
		SeedRep = "Judy Stanford" & vbCrLf & "AWSA West Seeding" & vbCrLf & "judy-don@sbcglobal.net" & vbCrLf & "(925) 932-7781"
	ELSEIF Session("Firstname")="Mike" AND INSTR(LCASE(Session("LastName")),"connor")>0 THEN
		IF instr("SM",mid(strTourID,3,1)) = 0 and len(Owner) > 0 THEN eMailCC = eMailCC & "; " & Owner
		''''eMailFrom = """Mike O'Connor"" <h2oskimo@gmail.com>"
		eMailFrom = """USA Water Ski Competition"" <h2oskimo@gmail.com>"
		eMailReplyTo = "h2oskimo@gmail.com"	
		SeedRep = "Mike O'Connor" & vbCrLf & "AWSA Midwest Seeding" & vbCrLf & "h2oskimo@gmail.com" & vbCrLf & "(573) 864-2138"
	ELSE
		IF instr("SMU",mid(strTourID,3,1)) = 0 and len(Owner) > 0 THEN eMailCC = eMailCC & "; " & Owner
		''''eMailFrom = """USA Water Ski Competition"" <shardee@usawaterski.org>"
		''''eMailReplyTo = "shardee@usawaterski.org"
		eMailFrom = """USA Water Ski Competition"" <mawsa@comcast.net>"
		eMailReplyTo = "mawsa@comcast.net"
		SeedRep = Session("Firstname") & " " & Session("LastName") & " on behalf of Sandy Hardee" & vbCRLF & "Competition Department HQ" & vbCRLF & "shardee@usawaterski.org" & vbCRLF & "1-863-324-4341 ext 126" & vbCRLF & "Direct Line:1-863-874-5681"
	END IF

    ''''WriteDebugSQL ("ExtractWfw.asp: emailto SeedRep " & SeedRep)
	IF mid(strTourID,3,1) = "U" THEN
		eMailCC = eMailCC & "; ""Jeff Surdej"" <j_surdej@yahoo.com>; ""Adam Koehler"" <adam.t.koehler@gmail.com>; ""Joey McNamara"" <ncwsa@joeymcnamara.com>"
	END IF

	eMailBody = "Dear Tournament Organizer and/or Chief Official(s) --" & vbCRLF & vbCRLF

	IF Session("UploadMode") = "Zip" THEN
		eMailBody = eMailBody & "The post-tournament reports from " & Session("strTourID")
		eMailBody = eMailBody & " " & Session("strTourName") & vbCRLF & "ending " & Session("strTourDate")
		eMailBody = eMailBody & " have been uploaded to the Sanction control system." & vbCRLF & vbCRLF
	ELSE
		eMailBody = eMailBody & "Post-tournament report file:  " & Session("strFileName") & "  has been uploaded to " & vbCRLF
		eMailBody = eMailBody & "to the Sanction control system and posted to Tournament " & Session("strTourID") & vbCRLF 
		eMailBody = eMailBody & "-- " & Session("strTourName") & " -- ending " & Session("strTourDate") & vbCRLF & vbCRLF
	END IF		

    WriteDebugSQL ("ExtractWfw.asp: emailto UploadMode " & Session("UploadMode"))
	IF len(strMissing) = 0 THEN

		eMailBody = eMailBody & "All of the required post-tournament reports are now accounted for" & vbCRLF 
		eMailBody = eMailBody & "after storing this upload.  These have been filed and checked off in" & vbCRLF
		eMailBody = eMailBody & "the Sanction Control System, and your Tournament marked complete." & vbCRLF & vbCRLF 
		eMailBody = eMailBody & "A big Thank You !!" & vbCRLF & vbCRLF

	ELSE

		IF Session("UploadMode") = "Zip" THEN

			eMailBody = eMailBody & "THE FOLLOWING ITEMS WERE MISSING OR INCOMPLETE IN THIS UPLOAD:" & vbCRLF
			eMailBody = eMailBody & strMissing & vbCRLF & vbCRLF
			eMailBody = eMailBody & "Please ensure that the above-noted items are completed and submitted" & vbCRLF 
			eMailBody = eMailBody & "soon.  Submission of these to me in a revised ZIP file via eMail is" & vbCRLF 
			eMailBody = eMailBody & "preferred, although you may submit paper documents direct to USA" & vbCRLF 
			eMailBody = eMailBody & "Waterski HQ by postal mail instead.  If the missing reports have" & vbCRLF 
			eMailBody = eMailBody & "already been sent, then please disregard this notice." & vbCRLF & vbCRLF
			eMailBody = eMailBody & "If you are having difficulty producing any of these reports, please" & vbCrLf
			eMailBody = eMailBody & "contact me for assistance." & vbCRLF & vbCRLF

		ELSE 

			eMailBody = eMailBody & "THE FOLLOWING ITEMS REMAIN OUTSTANDING AFTER THIS UPLOAD:" & vbCRLF
			eMailBody = eMailBody & strMissing & vbCRLF & vbCRLF
			eMailBody = eMailBody & "If you submitted some of these along with the Item being uploaded" & vbCRLF 
			eMailBody = eMailBody & "here, then you should receive additional advisory messages about" & vbCRLF 
			eMailBody = eMailBody & "those additional items shortly.  If the remaining missing items have" & vbCRLF 
			eMailBody = eMailBody & "already been mailed, please disregard this notice." & vbCRLF & vbCRLF
			eMailBody = eMailBody & "If you are having difficulty producing any of these reports, please" & vbCrLf
			eMailBody = eMailBody & "contact me for assistance." & vbCRLF & vbCRLF

		END IF

	END IF

    WriteDebugSQL ("ExtractWfw.asp: emailto strMissing " & strMissing)

	eMailBody = eMailBody & SeedRep
	
	IF Instr(strMissing, "CS.HTM") = 0 THEN
		eMailBody = eMailBody & vbCRLF & vbCRLF & "WSTIMS Windows Scorebook can be viewed at: " & vbCRLF
		eMailBody = eMailBody & "http://www.usawaterski.org/rankings/scorebks/"
		eMailBody = eMailBody & strTourID & "CS.HTM"
	END IF

	' Now send the eMail message -- three steps below ...
	
	' First we Invoke "standard" Email Server Configuration -- defines objMessage object
	SetupEmailService

	' Now supply email message details.
	objMessage.Subject = eMailSubj
	objMessage.To = eMailTo

	IF instr(eMailCC,eMailFrom) = 0 THEN eMailCC = eMailCC & "; " & eMailFrom

	objMessage.cc = eMailCC
	''''objMessage.From = """Competition Support"" <Post_Tour@usawaterski.org>"
    objMessage.From = eMailFrom
	objMessage.ReplyTo = eMailReplyTo

	objMessage.bcc = "mawsa@comcast.net"
	objMessage.TextBody = eMailBody

	WriteDebugSQL ("ExtractWfw.asp: Upload Time:  " & Date() & " " & Time())
    WriteDebugSQL ("ExtractWfw.asp: MailTo: " & eMailTo & " eMailCC: " & eMailCC)

	' Finally send the message, and then clear that object
    On Error Resume Next
    	objMessage.Send
        If Err.Number <> 0 Then
            WriteDebugSQL ("ExtractWfw.asp: Error sending email: Err.Number=" & Err.Number & " Message="  & Err.Description )
            On Error Goto 0 ' But don't let other errors hide!
        End If
    WriteDebugSQL ("ExtractWfw.asp: email sent")
	set objMessage = Nothing
	
	' Now append this message details to the eMails.txt file, creating if not already present.
	
	IF objFSO.FileExists(strTourFldr & "\eMails.txt") THEN
		Set objEMailTxt = objFSO.OpenTextFile(strTourFldr & "\eMails.txt", 8 ,true)
	ELSE
		Set objEMailTxt = objFSO.CreateTextFile(strTourFldr & "\eMails.txt", true)
	END IF

	objeMailTxt.WriteLine ("  To:  " & eMailTo)
	objeMailTxt.WriteLine ("Date:  " & Date() & " " & Time() & vbCRLF)
	objeMailTxt.WriteLine ("  Re:  Reports from " & Session("strTourID") & " " & Session("strTourName") & " (" & Session("strTourDate") & ")")
	objeMailTxt.WriteLine (vbCRLF & eMailBody)
	objeMailTxt.WriteBlankLines (4)
	objeMailTxt.Close
	Set objeMailTxt = Nothing
    WriteDebugSQL ("ExtractWfw.asp: emailto closed")

	' Now Zip this newly-modified eMails.txt file into the archive, then delete it.

	tmpFiles = objZip.ZipFilesTo(strTourFldr & "\eMails.txt", strTourZip)
	objFSO.DeleteFile (strTourFldr & "\eMails.txt")

END IF


'	==========================================
'	Done extracting and re-archiving.  Finally we kill the incoming ZIP
'	File, and also kill the	temporary working folder in the Scratch area.
'	==========================================

objFSO.DeleteFile (Session("strZipPath"))
objFSO.DeleteFolder strTourFldr

' Close and Release objects

Set objFSO = Nothing
Set objZip = Nothing
objRS.Close
CloseConSanUpd


'	Close out recap table and one empty line, ready for final recap stuff

%>

		</table></td>
		<td>&nbsp;&nbsp;&nbsp;</td>
	</tr>

	<tr>
		<td>&nbsp;&nbsp;&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;&nbsp;&nbsp;</td>
	</tr>

<%

'	List Tournament Status on Uploader's screen and also note in Log File

%>
	<tr>
		<td>&nbsp;&nbsp;&nbsp;</td>
		<td>
<%

	IF TStatus = 5 THEN
	WriteLog (date() & "  " & time() & "  Upload into " & strTourZip & " -- all Status flags set and Tournament posted as Complete.")
	%>
		<p><b>All required reports &amp; files present, Tournament has been Archived.</b></p>
	<%
ELSE
	WriteLog (date() & "  " & time() & "  Upload into " & strTourZip & " -- some items missing or incomplete, Flags = " & PTF_SBK & PTF_WSP & PTF_TS & PTF_OD & PTF_BT & PTF_JT & PTF_CS & PTF_CJ & PTF_SD & PTF_TU & PTF_HD & PTF_TNY)

	%>
		<p><b>Some requirements not filled, Tournament Status <i><font color=red>Not</font></i> set Complete</b>.</p>
	<%
END IF

%>
		</td>
	</tr>

	<tr>
		<td>&nbsp;&nbsp;&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;&nbsp;&nbsp;</td>
	</tr>
<%
					

'	=============================
'	Processing all finished -- Note email (if any) on Recap screen we've built
'	=============================

%>
		<tr>
			<td>&nbsp;&nbsp;&nbsp;</td>

			<% IF len(eMailTo) > 0 THEN %>
				<td><p>Receipt confirmation of this upload has been eMailed to:
					<br><%=Replace(Replace(eMailTo,"<","&lt;"),">","&gt;")%></p></td>
			<% ELSE %>
				<td><p>No eMail addresses found for Tournament Director or Submitter</p></td>
			<% END IF %>
			
			<td>&nbsp;&nbsp;&nbsp;</td>
		</tr>

		<tr>
			<td>&nbsp;&nbsp;&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;&nbsp;&nbsp;</td>
		</tr>

		<% IF instr("LRPAB",mid(Session("strTourID"),7,1)) > 0 THEN %>

		<tr>
			<td>&nbsp;&nbsp;&nbsp;</td>
			<td>&nbsp;IWWF Sees:&nbsp; <%=session("strIWWFSubj")%>&nbsp;</td>
			<td>&nbsp;&nbsp;&nbsp;</td>
		</tr>

		<tr>
			<td>&nbsp;&nbsp;&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;&nbsp;&nbsp;</td>
		</tr>

		<% END IF %>

<%


'	======================================
'	If FoundNewWSP flag, then check for existing 
'	scores and issue warning note if any found.
'	======================================

IF FoundNewWSP = True THEN

	sSQL = "SELECT count(*) as ScoreCount from (Select distinct MemberID from "
	sSQL = sSQL & RawScoresTableName & " WHERE upper(TourID) = '" 
	sSQL = sSQL & Ucase(Session("strTourID")) & "') xx;"
	
	'	WriteDebugSQL sSQL

	objRS.open sSQL, sConnectionToTRATable, 3, 3
   IF objRS.eof THEN ScoreCount = 0 ELSE ScoreCount = objRS("ScoreCount")
   objRS.Close
   
   IF ScoreCount > 0 THEN

		%>
		<tr>
			<td>&nbsp;&nbsp;&nbsp;</td>
			<td><p><font color=red><b>Warning !!</b></font>&nbsp; 
				A new or updated WSP file was found in this upload, yet scores <br>are 
				already present in the scores table for <font color=red><b><%=ScoreCount%> 
				skiers</b></font> for this tournament.  <br>If the Import option below 
				is selected, then those existing scores will be <br>deleted before that 
				import.&nbsp; <b><i>Do Not</i></b> proceed with the import, unless you 
				<br>are certain this <b><i>is</i></b> an intentionally revised WSP 
				file.</p></td>
			<td>&nbsp;&nbsp;&nbsp;</td>
		</tr>
	
		<tr>
			<td>&nbsp;&nbsp;&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;&nbsp;&nbsp;</td>
		</tr>
		<%

	END IF
END IF


'	==============================
'	Now set up for buttons at bottom
'	==============================

%>
		<tr>
			<td>&nbsp;&nbsp;&nbsp;</td>
			<td><TABLE align=center><tr>
<%

' Offer an Import Scores option button, only if the 
' Extract process found a new or updated WSP file here.

IF FoundNewWSP = True THEN %>
		
				<td><form method=post action="VerifyNewWSP.asp?WSPFile=<%=WSPFileName%>">

				<% IF ScoreCount > 0 THEN %>
 					<input type=submit style="width:13em" value="Re-Import WSP Scores"
 						title="Delete existing scores, then process the new/updated WSP file through the Ranking Scores Import Module">						

				<% ELSE %>
 					<input type=submit style="width:13em" value="Import Ranking Scores"
 						title="Process the included WSP file through the Ranking Scores Import Module">						
				<% END IF %>

	 	 		</form></td>
				<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>

<% END IF %>

				<td><form method=post action="DefaultHQ.asp?process=uploadany" method="post">
					<input type="submit" style="width:13em" value="Done with Upload"
					title="Return to the Upload Control Page">
					</form></td>	

			</tr></table></td>
			<td>&nbsp;&nbsp;&nbsp;</td>
		</tr>


	</table>
		
<%

Set objRS = Nothing

WriteIndexPageFooter
		

'	Subroutine to Extract a Named File from the incoming archive.
'	Tests for first presence, then date, then for size compared
'	to a specified minimum.  If present and not unused / incomplete,
'	then we attempt to extract/store, and then tests for not newer.

Sub ExtractFile (File2Ext, ToPath, MinSize)

IF Session("UploadMode") = "Zip" THEN

	'	Check for specified file -- check against incoming Zip File

	objZip.Open(Session("strZipPath"))
	objZip.Read
	objZip.DestDirectory = strTourFldr

	k = -1
	For i = 0 to objZip.Count - 1
		IF UCase(objZip.FileName(i)) = UCase(File2Ext) THEN k = i
	next

	IF k < 0 THEN
		IF objFSO.FileExists(ToPath & "\" & File2Ext) = true THEN
			Set objFile = objFSO.GetFile(ToPath & "\" & File2Ext)
			strStatus = trim(objFile.DateLastModified)
			Set objFile = Nothing
			strAction = "Prev Rcvd"
		ELSE
			strStatus = "<font color=red><b>Not Present</b></font>": strAction = "- -"
		END IF
	ELSE
		FileDate = objZip.FileDateTime(k)

		IF FileDate <= #01/02/2000# OR objZip.FileSize(k) < MinSize THEN
	'	IF FileDate <= #01/02/2000# THEN
			strStatus = "<font color=red><b>Empty/Incomplete</b></font>": strAction = "- -"
		ELSE
			strStatus = trim(FileDate)
			IF objZip.UnzipFileTo (File2Ext, ToPath, "IfNewer") = true THEN
				strAction = "<font color=blue><b>Stored</b></font>"
			ELSE 
				strAction = "Not Newer"
			END IF
		END IF
	END IF

ELSE

	'	Check for specified file against incoming single report file
	
	IF File2Ext <> Session("strFileName") THEN
		IF objFSO.FileExists(ToPath & "\" & File2Ext) = true THEN
			Set objFile = objFSO.GetFile(ToPath & "\" & File2Ext)
			strStatus = trim(objFile.DateLastModified)
			Set objFile = Nothing
			strAction = "Prev Rcvd"
		ELSE
			strStatus = "<font color=red><b>Not Present</b></font>": strAction = "- -"
		END IF
	ELSE
		Set objFile = objFSO.GetFile(Session("strZipPath"))
		strStatus = trim(objFile.DateLastModified)
		IF objFile.Size < MinSize THEN
			strStatus = "<font color=red><b>Empty/Incomplete</b></font>": strAction = "- -"
		ELSE
			strStatus = trim(ObjFile.DateLastModified)
			objFSO.CopyFile Session("strZipPath"), ToPath & "\", TRUE
			strAction = "<font color=blue><b>Stored</b></font>"
		END IF
	END IF

END IF

END SUB


' 
'	Subroutine to eMail a File to IWWF for Posting there.
' File Name dictates the particular eMail target Address.
' Specified file then moved to InBoxIWWF\Archived\ on completion.
'

Sub EmailToIWWF (File2Send)

' First we Invoke "standard" Email Server Configuration -- defines objMessage object
SetupEmailService

' Now supply email message details.
objMessage.Subject = session("strIWWFSubj")
objMessage.TextBody = session("strIWWFSubj")
objMessage.From = """USA Water Ski"" <dclark@usawaterski.org>"

' Choose CC addressee(s) depending on file being sent
IF mid(File2Send,7,2) = "HD" THEN
   objMessage.CC = """Kirby Whetsel"" <kwhetsel@charter.net>; ""Peter Dahl"" <awsatechcontroller@gmail.com>; ""Jerry Jackson"" <slalomjj@bellsouth.net>; ""Chip Shand"" <slalom@cox.net>; ""Jim Thompson"" <jim@lsfdev.com>; ""Will Bush"" <willbush@att.net>; ""Rodger Logan"" <rodgerlogan@gmail.com>"
ELSE
   objMessage.CC = """Kirby Whetsel"" <kwhetsel@charter.net>"
END IF	

' Choose Appropriate Send To Email Address depending on what file is being sent
IF Mid(File2Send,7,2) = "HD" THEN
	objMessage.To = """IWWF Dossier"" <dossier@iwsftournament.com>"
ELSEIF Mid(File2Send,7,2) = "CS" THEN
	objMessage.To = """IWWF Scorebook"" <scorebook@iwsftournament.com>"
ELSEIF Mid(File2Send,7,2) = "JT" THEN
	objMessage.To = """IWWF Jumptimes"" <jumptimes@iwsftournament.com>"
ELSE
	objMessage.To = """Kirby Whetsel"" <kwhetsel@charter.net>"
END IF

objMessage.AddAttachment PathtoIWWF & "\" & File2Send
objMessage.Send
Set objMessage=nothing

IF objFSO.FileExists (PathtoIWWF & "\Archived\" & File2Send) THEN
   objFSO.DeleteFile (PathtoIWWF & "\Archived\" & File2Send)
END IF

objFSO.CopyFile PathtoIWWF & "\" & File2Send, PathtoIWWF & "\Archived\", TRUE
objFSO.DeleteFile (PathtoIWWF & "\" & File2Send)
END Sub


'	Subroutine to display the description and file name and 
'	status and action taken on an expected file from archive.
'	These items displayed as elements in a row of an HTML Table.

Sub DisplayFile

	%>
		<tr>
			<td ALIGN="Left"><font size=2 face="arial">&nbsp;<%=strFileName%>&nbsp;</font></td>
			<td ALIGN="Left"><font size=2 face="arial">&nbsp;<%=strDescription%>&nbsp;</font></td>
			<td ALIGN="Center"><font size=2 face="arial">&nbsp;<%=strStatus%>&nbsp;</font></td>
			<td ALIGN="Center"><font size=2 face="arial">&nbsp;<%=strAction%>&nbsp;</font></td>
		</tr>
	<%

END Sub

%> 



                                                                    