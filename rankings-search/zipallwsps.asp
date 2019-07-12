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

Dim objFSO, objZip, objFile, objFolder, objFilesInFolder

Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Set objZip = Server.CreateObject("SoftComplex.Zip")

NewZipCount = 0
WSPsStored = 0

Set objFolder = objFSO.GetFolder(PathtoUploads)
Set objFilesInFolder = objFolder.Files

IF objFSO.FileExists (PathtoZips & "\ZipUpWSPLog.txt") THEN
	Set objLogFile = objFSO.OpenTextFile (PathtoZips & "\ZipUpWSPLog.txt", 8)
ELSE
	Set objLogFile = objFSO.CreateTextFile (PathtoZips & "\ZipUpWSPLog.txt", True)
END IF

IF objFilesInFolder.Count <> 0 THEN

	For Each objFile In objFolder.Files
		IF ucase(right(objfile.name,4)) = ".WSP" THEN
'		IF ucase(objfile.name) = "08S999A.WSP" THEN
			strTourZip = PathtoZips & "\" & ucase(left(objfile.name,7)) & ".ZIP"
			IF objFSO.FileExists(strTourZip) = false THEN
				objZip.New(strTourZip)
				tmpFiles = objZip.Save
				NewZipCount = NewZipCount + 1
			END IF
			tmpFiles = objZip.ZipFilesTo(PathtoUploads & "\" & objFile.Name, strTourZip)
			IF tmpFiles > 0 THEN
				objLogFile.WriteLine (objFile.name & "  " & objFile.datecreated & "  " & objFile.size)
				WSPsStored = WSPsStored + tmpFiles
			END IF
		END IF
	NEXT

END IF

objLogFile.Close
Set objLogFile = Nothing
%>
		<p>Created <%=NewZipCount%> new Zip Files</p>
		<p>Stored <%=WSPsStored%> WSP files.</p>

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

WriteIndexPageFooter

%>


