<% IF Session("adminmenulevel")<10 THEN Response.Redirect "DefaultHQ.asp?process=login" %>

<!--#include file="includes/clsUpload.asp"-->

<!--#include file="settingsHQ.asp"-->


<%

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
Dim strFileName ,strFileExt, strFileLength, strWSPARM

' Instantiate Upload Class and FSO object
Set objUpload = New clsUpload
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Set objZip = Server.CreateObject("SoftComplex.Zip")
Set objRS = Server.CreateObject("ADODB.recordset")


' Pick up the uploaded file and diagnose it.
' Grab the file name and extension and size

strFileName = objUpload.Fields("ZipFile").FileName
strFileExt = objUpload.Fields("ZipFile").FileExt
strFileLength = objUpload.Fields("ZipFile").Length
Session("strFileInfo") = "Upload File: " & strFileName & _
	"&nbsp;&nbsp;&nbsp; Type: " & strFileExt & _
	"&nbsp;&nbsp;&nbsp;	Size: " & Formatnumber(strFileLength,0)

'	Validate Uploaded File Type and Size

IF strFileExt = "LZH" THEN
	strError = "LZH archive files no longer allowed for uploading.&nbsp; Use the "
	strError = strError & "latest WSTIMS to create an archive for this tournament "
	strError = strError & "in ZIP format, then upload that ZIP file instead."
ELSEIF strFileExt <> "ZIP" THEN
	strError = "Upload File not ZIP Type"
ELSEIF strFileLength > 1500000 THEN
	strError = "Upload File too Large -- Skipped"
ELSE
	
	' Uploaded File is valid ZIP Type and not too large, so
	' Create a Session variable with pathname in the Scratch
	' sub-folder, and store the incoming upload file there.
	
	Session("strZipPath") = PathtoScratch & "\" & strFileName
	objUpload("ZipFile").SaveAs Session("strZipPath")

	' Open the ZIP File and extract WSPARM.TNY into a Text String, 
	' and then Parse out the key parameter values.

	objZip.Open(Session("strZipPath"))
	objZip.Read
	strCount = objZip.Count
	strWSPARM = objZip.UnzipFileToText("WSPARM.TNY")
	I1 = instr(strWSPARM,vbCRLF)
	IF I1>0 THEN I2 = instr(I1+2,strWSPARM,vbCRLF): ELSE I2 = 0
	IF I2>0 THEN I3 = instr(I2+2,strWSPARM,vbCRLF): ELSE I3 = 0
	IF I3>0 THEN C1 = instr(I1+2,strWSPARM,","): ELSE C1 = 0
	IF C1>0 THEN C2 = instr(C1+2,strWSPARM,","): ELSE C2 = 0

	IF I1=0 or I2=0 or I3=0 or C1=0 or C2=0 THEN
		strError = "WSPARM.TNY Data not Found (FilesInZip=" & strCount & ")"
	ELSE

		' Extract key Parameters in Session variables for display and subsequent
		' usage in downstream processes that will follow.

		Session("strTourID") = Mid(strWSPARM,C1+1,C2-C1-1)
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
			TStatus = 0
		ELSE
			TSanction = objRS("TSanction")
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

		'	Now recap what we've found so far and present options
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
			
		'	If End Dates don't match, then display warning.

		IF (TStatus < 2 and len(TSanction) <= 7) or TStatus = 3 THEN 
				
			%>
				<p><font color="red"><b>Exception:&nbsp; Sanction for this Event is either 
					incomplete or has been cancelled.&nbsp; Post-tournament upload will not be
					allowed.&nbsp; Please report these particulars to your Regional EVP and  
					copy the competition department at USA Waterski HQ.</b></font></p>
			<%	
		
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

		<% IF TStatus = 2 or TStatus >= 4 THEN %>
		
					<td><form method=post action="ExtractZip.asp">

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
	
END IF

' Otherwise if we had an earlier error then spit out only incoming file info
' and the specified Error message string.

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

			<td><form method=post action="DefaultHQ.asp?process=uploadzip" method="post">
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




