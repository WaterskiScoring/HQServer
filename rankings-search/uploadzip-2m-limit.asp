<% IF Session("adminmenulevel")<10 THEN Response.Redirect "DefaultHQ.asp?process=login" %>

<!--#include file="includes/clsUpload.asp"-->

<!--#include file="settingsHQ.asp"-->


<%

'	**********	
'	**********	Modified March 2011 to now support both DOS and WfW ZIP files
'	**********	


Dim strError
strError = ""

WriteIndexPageHeader

%>

		<table border="0" cellspacing="1" cellpadding="1">

		<tr>
			<td>&nbsp;&nbsp;&nbsp;</td>

			<td valign="top"><font face="Verdana, Arial, Helvetica, sans-serif" Size="2">
    	
<%

Dim objUpload, objfso, objZip
Dim strFileName, strFileExt, strFileLength, strWSPARM

' Instantiate Upload Class and FSO object
Set objUpload = New clsUpload
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Set objZip = Server.CreateObject("SoftComplex.Zip")
Set objRS = Server.CreateObject("ADODB.recordset")


' Pick up the uploaded file and diagnose it.
' Grab the file name and extension and size.
'	Modified Mar 2010 for IE 8 -- try FileName
'	first, if null then use FilePath instead.

'	strFileName = ucase(objUpload.Fields("ZipFile").FileName)

strFileName = trim(ucase(objUpload.Fields("ZipFile").FileName))
IF strFileName = "" then strFileName = ucase(objUpload.Fields("ZipFile").FilePath)

strFileExt = ucase(objUpload.Fields("ZipFile").FileExt)
strFileLength = objUpload.Fields("ZipFile").Length
Session("strFileName") = strFileName
Session("strFileInfo") = "Upload File: " & strFileName & _
	"&nbsp;&nbsp;&nbsp; Type: " & strFileExt & _
	"&nbsp;&nbsp;&nbsp;	Size: " & Formatnumber(strFileLength,0)

'	Validate Uploaded File Type and Size

IF strFileLength > 2000000 THEN
	strError = "Upload File too Large -- Skipped"
ELSEIF strFileName = "" OR strFileExt = "" THEN
	strError = "Invalid Upload File: " & strFileName
ELSEIF strFileExt = "LZH" THEN
	strError = "LZH archive files no longer allowed for uploading.&nbsp; Use the "
	strError = strError & "latest WSTIMS to create an archive for this tournament "
	strError = strError & "in ZIP format, then upload that ZIP file instead."
ELSEIF Instr(".CSV.HTM.PDF.PRN.SBK.TXT.WSP.ZIP", "." & ucase(strFileExt)) = 0 THEN
	strError = "Unrecognized Report File Type:&nbsp; " & strFileName
ELSEIF strFileExt = "ZIP" THEN
	
	' Uploaded File is ZIP Type.
	' Create a Session variable with pathname in the Scratch
	' sub-folder, and store the incoming upload file there.
	
	Session("UploadMode") = "Zip"
	Session("strZipPath") = PathtoScratch & "\" & strFileName
	objUpload("ZipFile").SaveAs Session("strZipPath")

	' Open the ZIP File and extract either WSPARM.TNY or WWParm.TNY
	' into a Text String, setting type and Extract link appropriately,
	' and then Parse out the key parameter values.

	objZip.Open(Session("strZipPath"))
	objZip.Read
	strCount = objZip.Count
	strWSPARM = objZip.UnzipFileToText("WSPARM.TNY")

	if len(strWSPARM) > 50 THEN
		strZipType = "DOS"
		strExtractModule = "ExtractZip.asp"
	ELSE
		strWSPARM = objZip.UnzipFileToText("WWPARM.TNY")
		strZipType = "WfW"
		strExtractModule = "ExtractWfW.asp"
	END IF	

	Session("strFileInfo") = Session("strFileInfo") & "&nbsp;&nbsp;&nbsp; (" & strZipType & " " & strCount & ")"

	I1 = instr(strWSPARM,vbCRLF)
	IF I1>0 THEN I2 = instr(I1+2,strWSPARM,vbCRLF): ELSE I2 = 0
	IF I2>0 THEN I3 = instr(I2+2,strWSPARM,vbCRLF): ELSE I3 = 0
	IF I3>0 THEN C1 = instr(I1+2,strWSPARM,","): ELSE C1 = 0
	IF C1>0 THEN C2 = instr(C1+2,strWSPARM,","): ELSE C2 = 0

	IF I1=0 or I2=0 or I3=0 or C1=0 or C2=0 THEN
		strError = "WxPARM.TNY Data not Found (FilesInZip=" & strCount & ")"
	ELSE

		' Extract key Parameters in Session variables for display and subsequent
		' usage in downstream processes that will follow.

		Session("strTourID") = ucase(Mid(strWSPARM,C1+1,C2-C1-1))
		Session("strTourDate") = Mid(strWSPARM,C2+1,I2-C2-1)
		Session("strTourName") = Mid(strWSPARM,I2+2,I3-I2-2)
		Session("strTourZip") = PathtoZips & "\" & Session("strTourID") & ".ZIP"

		' Now that we have the official Sanction ID, let's look to SWIFT 
		' Tschedul table and get access to the particulars of that affair.
		
		' ====================================================
		'	SWIFT Lookup & Validation logic here
		' ====================================================
             
		sSQL = "Select top 1 * from " & SanctionTableName & " where upper(TournAppID) = '"
		sSQL = sSQL & ucase(left(Session("strTourID"),6)) & "'"
		objRS.open sSQL, sConnectionToSanctionTable
		If objRS.EOF Then
			TSanction = "<font color=red>Missing"
			TName = "No Such TourID Found</font>"
			TDateE = ""
			TStatus = -1
		ELSE
			TSanction = objRS("TSanction")
			IF Session("strTourID") <> TSanction THEN 
				TSanction = "<font color=red>" & TSanction & "</font>"
			END IF
			TName = objRS("TName")
			TDateE = Replace(FormatDateTime(objRS("TDateE"),2),"/","-")
			IF Mid(TDateE,2,1) = "-" THEN TDateE = "0" & TDateE
			IF Mid(TDateE,5,1) = "-" THEN TDateE = Left(TDateE,3) & "0" & Right(TDateE,6)
			IF Session("strTourDate") <> TDateE THEN 
				TDateE = "<font color=red>" & TDateE & "</font>"
			END IF
			Tstatus = objRS("Tstatus")
		END IF
		objRS.Close

		'	Now recap what we've found so far and then check for various exceptions.
		'	First show Input file particulars, then SWIFT's
		
		%>

		<p><%=Session("strFileInfo")%></p>
	
		<p><b>Upload:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<%=Session("strTourID")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<%=Session("strTourName")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<%=Session("strTourDate")%><br>
			&nbsp;SWIFT:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<%=TSanction%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<%=TName%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
			<%=TDateE%></b></p>	

		<%
			
		'	If Can't find SWIFT entry for supplied Sanction App ID, then display error message.

		IF TStatus < 0 THEN 
				
			%>
				<p><font color="red"><b>Exception:&nbsp; TourID used to score this Tournament 
					not found in SWIFT.</b></font></p>
			<%	
		
		'	If Sanction Status incomplete or Cancelled, then display error message.

		ELSEIF TStatus < 2 or TStatus = 3 THEN 
				
			%>
				<p><font color="red"><b>Exception:&nbsp; Sanction for this Event is either 
					incomplete or has been cancelled.&nbsp; Post-tournament upload will not be
					allowed.&nbsp; Please report these particulars to your Regional EVP and  
					copy the competition department at USA Waterski HQ.</b></font></p>
			<%	
		
		'	If Sanction Suffix code doesn't match, then display error message.  Further,
		'	If SWIFT suffix is A or B or P, then disallow Upload -- otherwise just warn.

		ELSEIF TSanction <> Session("strTourID") THEN 
			
			IF	Instr("ABP",Mid(TSanction,7,1)) > 0 THEN 
			
				TStatus = -1 %>

				<p><font color="red"><b>Exception:&nbsp; Highest Class suffix on TourID used to
					score this tournament does not match SWIFT.&nbsp; Post-tournament upload will not be
					allowed.&nbsp; Please advise the submitter of this discrepancy and ask them to either
					revise their scoring details and resubmit -- or else have SWIFT corrected, if the 
					error lies on that side.</b></font></p>
			
			<% ELSE %>
			
				<p><font color="red"><b>Warning:&nbsp; Highest Class scored as
					(<%=Mid(Session("strTourID"),7,1)%>) does not match SWIFT.</b></font></p>

			<%	END IF
		
		END IF
		
		'	If End Dates don't match, then display warning.

		IF TStatus > 0 AND Session("strTourDate") <> TDateE THEN 
				
			%>
				<p><font color="red"><b>Warning:&nbsp; Reported End Date does
					not match SWIFT.</b></font></p>
			<%	
		
		END IF
		
		'	If previously uploaded then warn and show last date updated.

		IF objFSO.FileExists(Session("strTourZip")) = true THEN 
			Set objFile = objFSO.GetFile(Session("strTourZip"))
			strZipDate = objFile.DateLastModified
			Set objFile = Nothing	
				
			%>
				<p><font color="red"><b>Reports Archive for Tournament 
					<%=Session("strTourID")%> already exists</b></font><br>
					&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
					( last posted to on <%=strZipDate%> ).</p>
			<%	
		
		END IF
		
		' ===========================================
		' possibly add other advisory conditions here
		' ===========================================

		%>

				</td></tr>

		<tr>
			<td>&nbsp;&nbsp;&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;&nbsp;&nbsp;</td>
		</tr>

		<tr>
			<td>&nbsp;&nbsp;&nbsp;</td>
			<td><TABLE align=center><tr>

		<%
		' Now finally offer an Upload option button (with flavors), but only
		' if the SWIFT Entry Found and acceptable conditions to allow Upload.
		%>

		<% IF TStatus = 2 or TStatus > 3 THEN %>
		
					<td><form method=post action="<%=strExtractModule%>">

			<% IF objFSO.FileExists(Session("strTourZip")) = true THEN %> 

		 				<input type=submit style="width:13em" value="Re-Upload Tournament"
 	 					title="Re-Process this Tournament, adding any new files, and replacing any previously uploaded files that now have newer datestamps">						

			<% ELSE %>

 						<input type=submit style="width:13em" value="Upload Tournament"
							title="Process this Tournament, extracting key data files from the incoming Zip file">					

			<% END IF %>

 		  		</form></td>
				<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>

		<% END IF 

	END IF

ELSE

	' Uploaded File is something other than ZIP Type.
	' Create a Session variable with pathname in the Scratch
	' sub-folder, and store the incoming upload file there.
	
	Session("UploadMode") = "Rpt"

	' First we validate the file name 

	IF strFileExt <> "WSP" AND strFileExt <> "SBK" AND Instr("BT.PRN/JT.CSV/OD.TXT/CJ.PDF/CJ.TXT/HD.TXT/SD.PDF/SD.TXT/TU.PDF/TU.TXT/TS.PRN/TS.TXT/CS.HTM", Mid(strFileName,7)) = 0 THEN
		strError = "Unrecognized Report File Type:&nbsp; " & strFileName
	ELSE

		' File Name is of a valid format/type -- look for the Tournament Archive Zip file

		Set objFolder = objFSO.GetFolder(PathtoZips)
		set objFilesInFolder = objFolder.Files
		strZipFile = ""
		For Each objFile In objFolder.Files
			IF Left(strFileName,6) = Left(objFile.Name,6) THEN strZipFile = objFile.Name
		Next
		Set objFilesInFolder = Nothing
		Set objFolder = Nothing
		
		IF strZipFile = "" THEN
			strError = "Tournament Archive not found for " & Left(strFileName,6) & "<br>"
			strError = strError & "&nbsp;&nbsp; (Main .ZIP file must be uploaded first)"
		ELSE
		
			' Upload File Validated, and associated Zip Archive file was found. 
			' We're good -- finish setting up to process this single report file.



			Session("strTourID") = Left(strZipFile,7)			
			Session("strZipPath") = PathtoScratch & "\" & strFileName
			objUpload("ZipFile").SaveAs Session("strZipPath")
			Session("strTourZip") = PathtoZips & "\" & strZipFile

			objZip.Open(Session("strTourZip"))
			objZip.Read
			strCount = objZip.Count
			strWSPARM = objZip.UnzipFileToText("WSPARM.TNY")

			if len(strWSPARM) > 50 THEN
				strZipType = "DOS"
				strExtractModule = "ExtractZip.asp"
			ELSE
				strWSPARM = objZip.UnzipFileToText("WWPARM.TNY")
				strZipType = "WfW"
				strExtractModule = "ExtractWfW.asp"
			END IF	

			Session("strFileInfo") = Session("strFileInfo") & "&nbsp;&nbsp;&nbsp; (" & strZipType & ")"


			' ====================================================
			'	SWIFT Lookup & Validation logic here
			' ====================================================
             
			sSQL = "Select top 1 * from " & SanctionTableName & " where upper(TournAppID) = '"
			sSQL = sSQL & ucase(left(Session("strTourID"),6)) & "'"
			objRS.open sSQL, sConnectionToSanctionTable
			If objRS.EOF Then
				TSanction = "<font color=red>Missing"
				TName = "No Such TourID Found</font>"
				TDateE = ""
				TStatus = -1
			ELSE
				TSanction = objRS("TSanction")
				TName = objRS("TName")
				TDateE = Replace(FormatDateTime(objRS("TDateE"),2),"/","-")
				IF Mid(TDateE,2,1) = "-" THEN TDateE = "0" & TDateE
				IF Mid(TDateE,5,1) = "-" THEN TDateE = Left(TDateE,3) & "0" & Right(TDateE,6)
				TStatus = objRS("Tstatus")
				Session("strTourDate") = TDateE
				Session("strTourName") = TName

			END IF
			objRS.Close

			'	Now recap status and then offer upload and abort options
			'	Show Input file particulars, then SWIFT info on this tournament
		
			%>

			<p><%=Session("strFileInfo")%></p>
	
			<p><b>SWIFT:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<%=TSanction%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<%=TName%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
				<%=TDateE%></b></p>	

				</td></tr>

			<tr>
				<td>&nbsp;&nbsp;&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;&nbsp;&nbsp;</td>
			</tr>

			<tr>
				<td>&nbsp;&nbsp;&nbsp;</td>
				<td><TABLE align=center><tr>
		
				<td><form method=post action="<%=strExtractModule%>">

 						<input type=submit style="width:13em" value="Upload This File"
							title="Process this Report File and post to the indicated Tournament">

 		  		</form></td>
				<td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>

			<% 
			
		END IF
		
	END IF
	
END IF

' If we have an error message string, then spit out only incoming 
' file info along with that specified Error message string.

IF strError <> "" THEN


	%>
		<p><%=Session("strFileInfo")%></p>
		<p><font color="red"><b><%=strError%></b></font></p>
		</td></tr>

		<tr>
			<td>&nbsp;&nbsp;&nbsp;</td>
			<td>&nbsp;</td>
			<td>&nbsp;&nbsp;&nbsp;</td>
		</tr>

		<tr>
			<td>&nbsp;&nbsp;&nbsp;</td>
			<td><TABLE align=center><tr>

	<%

END IF

%>

			<td><form method=post action="DefaultHQ.asp?process=uploadany" method="post">
				<input type="submit" style="width:13em" value="Abort this Upload"
				title="Abort this Upload and return to the Upload Control Page">
				</form></td>

		</tr></table></td>
		<td>&nbsp;&nbsp;&nbsp;</td>
	</tr>

	<tr>
		<td>&nbsp;&nbsp;&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;&nbsp;&nbsp;</td>
	</tr>

	</table>

<%

' Release objects from memory

Set objUpload = Nothing
Set objFSO = Nothing
Set objZip = Nothing
Set objRS = Nothing

WriteIndexPageFooter

%>




