<% IF Session("adminmenulevel")<10 THEN Response.Redirect "DefaultHQ.asp?process=login" %>

<!--#include file="settingsHQ.asp"-->

Server.ScriptTimeout = 600 

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

Dim objFSO, objZip, objFile, objFolder, objFilesInFolder, sSQL, strUpdt
Dim strTourID, strTourFldr, ToursStaged, NewPTFRows, strTourApp
Dim Scored0, Scored1, Scored2, Scored3, Scored4
Dim PTF_SBK, PTF_WSP, PTF_TS, PTF_OD, PTF_BT, PTF_JT
Dim PTF_CS, PTF_CJ, PTF_SD, PTF_TU, PTF_HT, PTF_TNY
Dim TName, TDateE, TSanction, strTStatus

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Set objZip = Server.CreateObject("SoftComplex.Zip")
Set objRS = Server.CreateObject("ADODB.recordset")
OpenConSanUpd


ToursStaged = 0: NewPTFRows = 0

IF objFSO.FileExists (PathtoZips & "\StagePostTournFlags.txt") THEN
	Set objLogFile = objFSO.OpenTextFile (PathtoZips & "\StagePostTournFlags.txt", 8)
ELSE
	Set objLogFile = objFSO.CreateTextFile (PathtoZips & "\StagePostTournFlags.txt", True)
END IF

Set objFolder = objFSO.GetFolder(PathToZips)
Set objFilesInFolder = objFolder.Files

objLogFile.WriteLine ("Begin PostTourFlag Staging Run on " & Date() & " at " & Time())

%><p>Begin PTF Staging process for Ski Year 2010 Tournaments.</p><%

IF objFilesInFolder.Count <> 0 THEN

	'	Main Loop across Stored Zip files -- 2009 + only

	For Each objFile In objFolder.Files

		IF ucase(right(objfile.name,4)) = ".ZIP" and ucase(left(objfile.name,2)) = "10" THEN
			
			strTourID = ucase(left(objfile.name,instr(objfile.name,".ZIP")-1))
			strTourApp = left(strTourID,6)
			Session("strTourID") = strTourID

			'
			'	Read Sanction Table entry and get key items plus Scored0-4 flags
			'

			sSQL = "Select top 1 * from " & SanctionTableName & " where upper(TournAppID) = '"
			sSQL = sSQL & strTourApp & "'"
			objRS.open sSQL, sConnectionToSanctionTable, 3, 3
			IF objRS.EOF THEN
				strTStatus = -1
				objLogFile.WriteLine (strTourID & " No Sanction Table Entry Found")
			ELSE 
				strTStatus = objRS("TStatus")
				TSanction = objRS("TSanction")
				TName = objRS("TName")
				TDateE = Replace(FormatDateTime(objRS("TDateE"),2),"/","-")
				IF Mid(TDateE,2,1) = "-" THEN TDateE = "0" & TDateE
				IF Mid(TDateE,5,1) = "-" THEN TDateE = Left(TDateE,3) & "0" & Right(TDateE,6)
				IF objRS("Scored0") = True THEN Scored0 = 1 else Scored0 = 0
				IF objRS("Scored1") = True THEN Scored1 = 1 else Scored1 = 0
				IF objRS("Scored2") = True THEN Scored2 = 1 else Scored2 = 0
				IF objRS("Scored3") = True THEN Scored3 = 1 else Scored3 = 0
				IF objRS("Scored4") = True THEN Scored4 = 1 else Scored4 = 0
				IF objRS("TEventSlalom") = True THEN TEventSlalom = 1 else TEventSlalom = 0
				IF objRS("TEventJump") = True THEN TEventJump = 1 else TEventJump = 0
			END IF
			objRS.Close
			
			IF strTStatus <> -1 THEN

			'
			'	Unpack the Zip file into temporary folder in prep for analysis of content files
			'

			strTourFldr = PathtoScratch & "\" & strTourID
			strTourZip = PathtoZips & "\" & objfile.name
			objZip.Open(strTourZip)
			objZip.Read
			objZip.DestDirectory = strTourFldr
			tmpFiles = objZip.UnZip

			'
			'	Now Deal with all 12 individual files and set flags accordingly, based 
			'	first on whether the file is present, then secondarily based on whether
			'	the associated Scoredx flag is set, indicating manual receipt.
			'
			
			' [Sanction].SBK Full Scorebook Report Flag
			IF objFSO.FileExists(strTourFldr & "\" & strTourID & ".SBK") THEN
				PTF_SBK = 1
			ELSEIF Scored0 = 1 THEN
				PTF_SBK = 2
			ELSE 
				PTF_SBK = 0
			END IF

			' [Sanction].WSP Seeding Data File Flag
			IF objFSO.FileExists(strTourFldr & "\" & strTourID & ".WSP") THEN
				PTF_WSP = 1
			ELSEIF Scored0 = 1 THEN
				PTF_WSP = 2
			ELSE 
				PTF_WSP = 0
			END IF

			' [Sanction]TS.PRN Tournament Summary Report Flag
			IF objFSO.FileExists(strTourFldr & "\" & strTourApp & "TS.PRN") THEN
				PTF_TS = 1
			ELSEIF Scored1 = 1 THEN
				PTF_TS = 2
			ELSE 
				PTF_TS = 0
			END IF

			' [Sanction]OD.TXT Officials Credits Data File Flag
			IF objFSO.FileExists(strTourFldr & "\" & strTourApp & "OD.TXT") THEN
				PTF_OD = 1
			ELSEIF Scored4 = 1 THEN
				PTF_OD = 2
			ELSE 
				PTF_OD = 0
			END IF

			' [Sanction]BT.PRN Boat Time Tracking Report Flag
			IF TEventSlalom = 0 and TEventJump = 0 THEN
				PTF_BT = 3
			ELSEIF objFSO.FileExists(strTourFldr & "\" & strTourApp & "BT.PRN") THEN
				PTF_BT = 1
			ELSEIF Scored3 = 1 THEN
				PTF_BT = 2
			ELSE 
				PTF_BT = 0
			END IF

			' [Sanction]JT.CSV Jump Time Data File Flag
			IF TEventJump = 0 THEN
				PTF_JT = 3
			ELSEIF objFSO.FileExists(strTourFldr & "\" & strTourApp & "JT.CSV") THEN
				PTF_JT = 1
			ELSEIF Scored3 = 1 THEN
				PTF_JT = 2
			ELSE 
				PTF_JT = 0
			END IF

			' [Sanction]CS.HTM Condensed Scorebook Report Flag
			IF objFSO.FileExists(strTourFldr & "\" & strTourApp & "CS.HTM") THEN
				PTF_CS = 1
			ELSEIF Scored0 = 1 THEN
				PTF_CS = 2
			ELSE 
				PTF_CS = 0
			END IF

			' [Sanction]CJ.PDF Chief Judge Tournament Report Flag
			IF objFSO.FileExists(strTourFldr & "\" & strTourApp & "CJ.PDF") THEN
				PTF_CJ = 1
			ELSEIF Scored1 = 1 THEN
				PTF_CJ = 2
			ELSE 
				PTF_CJ = 0
			END IF

			' [Sanction]SD.PDF Safety Director Report Flag
			IF objFSO.FileExists(strTourFldr & "\" & strTourApp & "SD.PDF") THEN
				PTF_SD = 1
			ELSEIF Scored2 = 1 THEN
				PTF_SD = 2
			ELSE 
				PTF_SD = 0
			END IF

			' [Sanction]TU.PDF Towboat Utilization Report Flag
			IF objFSO.FileExists(strTourFldr & "\" & strTourApp & "TU.PDF") THEN
				PTF_TU = 1
			ELSEIF Scored3 = 1 THEN
				PTF_TU = 2
			ELSE 
				PTF_TU = 0
			END IF

			' [Sanction]HD.TXT Homologation Dossier Report Flag
			IF instr("ELRPAB",mid(strTourID,7,1)) = 0 THEN
				PTF_HD = 3
			ELSEIF objFSO.FileExists(strTourFldr & "\" & strTourApp & "HD.TXT") THEN
				PTF_HD = 1
			ELSEIF Scored1 = 1 THEN
				PTF_HD = 2
			ELSE 
				PTF_HD = 0
			END IF

			' WSPARM.TNY or WWPARM.TXT Tournament Settings Control File Flag
			IF objFSO.FileExists(strTourFldr & "\WSPARM.TNY") THEN
				PTF_TNY = 1
			ELSEIF objFSO.FileExists(strTourFldr & "\WWPARM.TXT") THEN
				PTF_TNY = 1
			ELSEIF Scored0 = 1 THEN
				PTF_TNY = 2
			ELSE 
				PTF_TNY = 0
			END IF

			'
			'	Now delete extracted zip content along with temp folder we created.
			'

			objFSO.DeleteFile (strTourFldr & "\*.*")
			objFSO.DeleteFolder strTourFldr

			'
			'	Next we update S_PostTourn table flags -- but check and insert row first, if none there
			'

			sSQL = "Select top 1 * from " & PostTourTableName & " where upper(TournAppID) = '"
			sSQL = sSQL & strTourApp & "'"
			objRS.open sSQL, sConnectionToSanctionTable, 3, 3
			If objRS.EOF THEN
				sSQL = "Insert Into " & PostTourTableName & " Values ('" & strTourApp 
				sSQL = sSQL & "', 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)"
				ConSanUpd.Execute(sSQL)
				objLogFile.WriteLine (strTourID & " " & TDateE & " S_PostTourn Row Added for " & TName)
				NewPTFRows = NewPTFRows + 1
			END IF 
			objRS.Close

			strUpdt = " Set PTF_SBK=" & trim(PTF_SBK)
			strUpdt = strUpdt & ", PTF_WSP=" & trim(PTF_WSP)
			strUpdt = strUpdt & ", PTF_TS=" & trim(PTF_TS)
			strUpdt = strUpdt & ", PTF_OD=" & trim(PTF_OD)
			strUpdt = strUpdt & ", PTF_BT=" & trim(PTF_BT)
			strUpdt = strUpdt & ", PTF_JT=" & trim(PTF_JT)
			strUpdt = strUpdt & ", PTF_CS=" & trim(PTF_CS)
			strUpdt = strUpdt & ", PTF_CJ=" & trim(PTF_CJ)
			strUpdt = strUpdt & ", PTF_SD=" & trim(PTF_SD)
			strUpdt = strUpdt & ", PTF_TU=" & trim(PTF_TU)
			strUpdt = strUpdt & ", PTF_HD=" & trim(PTF_HD)
			strUpdt = strUpdt & ", PTF_TNY=" & trim(PTF_TNY)

			sSQL = "Update " & PostTourTableName & strUpdt & " WHERE TournAppID='" & strTourApp & "'"
			
			ConSanUpd.Execute(sSQL)
			
			objLogFile.WriteLine (strTourID & " " & TDateE & strUpdt & " " & TName)
			
			ToursStaged = ToursStaged + 1
			
			END IF

		END IF

	NEXT

END IF

objLogFile.Close
Set objLogFile = Nothing

%>
		<p>New PTF Rows Created for <%=NewPTFRows%> Tournaments.</p>
		<p>Updated PostTourFlags for <%=ToursStaged%> Tournaments.</p>

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

Set objFilesInFolder = Nothing
Set objFolder = Nothing
Set objFSO = Nothing
Set objZip = Nothing
CloseConSanUpd
	

WriteIndexPageFooter

%>

