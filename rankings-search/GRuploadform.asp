<%@ Language=VBScript %>

<%Option explicit

%>
<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_registration.asp"-->
<!--#include virtual="/rankings/tools_Include.asp"-->
<%

Dim strMessage, strFolder, GRUploadedFolderFullPath, GRImportedFolder, GRUploadedFolder, GRImportedFolderFullPath
Dim httpref, lngFileSize
Dim strExcludes, strIncludes
Dim ThisFileName, UploadFormFile
Dim ScoresGRTableName



' --- Names of folders (note there is no / at end)
GRImportedFolder=Server.mappath("/rankings/GR_Results/Imported")
GRImportedFolderFullPath="http://usawaterski.org/rankings/GR_Results/Imported"
GRUploadedFolder=Server.mappath("/rankings/GR_Results")
GRUploadedFolderFullPath="http://usawaterski.org/rankings/GR_Results"


' --- Defines the location of the uplader program and table in database ---
UploadFormFile="GRUploadform.asp"



' --- Creates css and draws standard window ---
DefineTRAStyles 
WriteIndexPageHeader






	'-----------------------------------------------
	'This is the complete upload file program.
	'This is intended to upload graphics onto the web and
	'to delete them if required.
	'Set up the configurations below to define which
	'directory to use etc, then set the permissions on
	'the directory to 'Change' i.e. Read/Write
	'-----------------------------------------------



	' --- This code is used in place of a CONFIG file ---
	'----------------------------------------------------
	' --- name of folder (note there is no / at end)
	strFolder=Server.mappath("/rankings/GR_Results/")



	' --- name of folder in http format (note there is no / at end)
	httpRef="http://usawaterski.org/rankings/GR_Results"

	' --- the max size of file which can be uploaded, 0 will give unlimited file size
	lngFileSize = 1000000


	' --- the files to be excluded (must be in format ".aaa;.bbb") and must be set to blank ("") if none are to be excluded
	strExcludes = ""
	' --- the files to be included (must be in format ".aaa;.bbb") and must be set to blank ("") if none are to be excluded
	strIncludes = ".xls"
	
	strMessage = Request.QueryString ("msg")


	
' ---------------
  SUB main()
' ---------------


	IF Request.Form ("AskDelete") = "Delete" THEN		' --- ask if to delete
		'response.write("<br>RFFile = "&Request.Form("fileId")) 	
		'response.end
		call askDelete(Request.Form("fileId"))

	ELSEIF Request.Form("main") = "Upload Excel Results" THEN	' --- Upload file routine
		call displayform()

	ELSEIF Request.Form("main") = "Import Uploaded File" THEN	' --- Upload file routine
		response.redirect("/rankings/GRExcelImport.asp")
		'response.end

	ELSEIF Request.Form("main") = "Delete Imported File" THEN	' --- Delete file routine
		call BuildFileList(GRUploadedFolder)


	ELSEIF Request.Form ("delete") = "Yes" THEN		 ' --- make deletion
		call delete(Request.form("fileId"))
		'call displayForm()
		call BuildFileList(GRUploadedFolder)
	ELSEIF Request.Form ("delete") = "No" THEN		' --- No pressed from DELETE routine - don't make deletion
		'call displayForm()
		call BuildFileList(GRUploadedFolder)

	ELSEIF Request.Form("delete") = "" THEN			' --- display at start up
		call displaymainform()

	END IF


END SUB


' ---------------------
  SUB DisplayMainForm()
' ---------------------

' --- Displays main header for selection of Action --- %>

	<html>
	<head>
	<title>File Upload</title>

	</head>
	<body>

	<form action="/rankings/GRuploadform.asp" method="post" >

	<TABLE class="innertable" width=500px align="center" style="padding:2px; border-collapse:collapse; border:1px solid <%=HQSiteColor2%>;" >

		<tr>
			<th colspan="2" class="text" align="center">
				<font size="3" color="white"><b>Grassroots Upload Home Page</b></font>
			</th>		
		</tr>
		<tr>
		  <td align="center" style="border-style:none;">
			<br>
			<font size="2">Select the function you wish to perform</font>
			<br>
		  </td>
		</tr>		
		<tr>
		  <td align="center" style="border-style:none;">
			<input type="submit" style="width:14em;" value="Upload Excel Results" name="main">
			<br>
		  </td>
		</tr>		
		<tr>
		  <td align="center" style="border-style:none;">
			<input type="submit" style="width:14em;" value="Import Uploaded File" name="main">
			<br>
		  </td>
		</tr>		
		<tr>
		  <td align="center" style="border-style:none;">
			<input type="submit" style="width:14em;" value="Delete Imported File" name="main">
			<br><br>
		  </td>
		</tr>		


	</TABLE>
	</form>

	</body>
	</html><%


END SUB




' --------------------
  SUB displayForm()
' --------------------

' --- Displays the form to allow uploading

Dim i, tempArray

	'Results box
	IF strMessage <> "" THEN %>
		<html>
		<head>
		  <title>File Upload</title>
		</head>

		<body>

		<TABLE class="droptable" width=500px align="center">
		  <tr>
		    <td class="text"><%=strMessage%></td>
	 	  </tr>
		</TABLE><%
	END IF  



	' -------- Displays form with upload box --------- 
	%>
	<TABLE class="innertable" width=500px align="center" style="padding:2px; border-collapse:collapse; border:1px solid <%=HQSiteColor2%>;" >

		<tr>
		  <th colspan="2" class="text" align="center">
			<font size="3" color="white"><b>Select Excel Spreadsheet to Upload</b></font>
		  </th>		
		</tr>

		<form action="/rankings/GRuploadform.asp" method="post" >
		  <tr>
		    <td style="border-style:none;" align="center">
			<font size="1">Use Browse Button To Select File then Press Upload
			<br><br>
			<b>IMPORTANT:</b> File name must be contiguous <i><u>NO SPACES!</u></i></font>
		    </td>
		    <td align="center" style="border-style:none;">
			<br>
			<input type="submit" style="width:6em;" value="Done" name="submit">
			<br><br>
		    </td>
		  </tr>
		</form>

		<form action="/rankings/GRuploadfile.asp" method="post" enctype="multipart/form-data">
		<tr>
		  <td align="center" style="border-style:none;">
				<br>
				<b>File: </b><input type="file" name="file1" /><br/>	
		  </td>
		  <td align="center" style="border-style:none;">
			<input type="submit" style="width:6em;" value="Upload" name="submit">
			<br>
		  </td>
		</tr>		

		<tr>
		  <td class="text" colspan=2 align="center" style="border-style:none;"><%
	
			IF strExcludes <> "" THEN %>
				<font size=1>File types which cannot be uploaded = <br><%
				tempArray = Split(strExcludes,";")
				FOR i = 0 TO UBOUND(tempArray)
					Response.Write (tempArray(i)) & " "
				NEXT %>
				<font><%
			END IF

			IF strIncludes <> "" THEN %>
				<font size=1>Only <%
				tempArray = Split(strIncludes,";")
				FOR i = 0 TO UBOUND(tempArray)
					Response.Write (tempArray(i)) & " "
				NEXT %>
				<font><%
			END IF 		

			IF lngFileSize > 0 THEN %>
				<font size=1> file types are allowed with Max file size = <%=lngFileSize%> bytes</font>
				<br><%
			END IF  %>
		
		  </td>
		</tr>
		</form>

	</TABLE>
		

	</body>
	</html><%


END SUB




' ---------------------------------
  SUB BuildFileList(GRUploadedFolder)
' ---------------------------------


' --------------------------------------------
' --- Builds a list of files on the directory
' --- INPUT : the folder to be used
' ---------------------------------------------

    Dim oFS, oFolder, intNoOfFiles, FileName

    Set oFS = Server.CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFS.getFolder(GRImportedFolder)

	%>
	<html>
	<head>
	  <title>File Upload</title>
	</head>
	<body>

	<form action="/rankings/GRuploadform.asp" method="post" >

	<TABLE class="innertable" width=500px align="center">
	  <tr>
	    <th class="text" align="center" colspan="3">
		<font color="white" size="3">Files Previously Uploaded</font>
	    </th>
	  </tr>

	  <tr>
	    <td align="left" style="border-style:none;" width=50px>&nbsp;</td>

	    <td align="left" style="border-style:none;">
		<font size="1"><b> Options:</b><br>&nbsp;&nbsp;- Click Link to Download Spreadsheet<br>&nbsp;&nbsp;- Delete file<br>&nbsp;&nbsp;- Press Done to Return to Menu</font>
	    </td>
	    <td align="center" style="border-style:none;">
		<br>
		<input type="submit" style="width:6em;" value="Done" name="submit">
		<br><br>
	    </td>
	  </tr>
	</TABLE>
	<TABLE class="innertable" width=500px align="center">
	  <tr>
	    <td class="text" align="left" colspan="2"><font size=2><b>Filename</b></font></td>
	    <td class="text" align="center"><font size=2><b>Action</b></font></td>
	  </tr><%


	intNoOfFiles = 0


	Dim CurrentFile
    FOR Each FileName IN oFolder.Files	
		CurrentFile=mid(FileName.Path,instrrev(FileName.Path,"\")+1)
		%>
		<tr>		
			<form Name="frmDelete" method="post" action="<%=Request.ServerVariables("PATH_INFO")%>">
				<td class="text" align="left" colspan="2">
					<font size="1">
					<a href="<%=GRImportedFolderFullPath%>/<%=CurrentFile%>" title="Download the Excel document <%=CurrentFile%>" target="_blank"><%=CurrentFile%></a>
				</td>
				<td class="text" align="center">
					<input type="hidden" name="fileId" value="<%=mid(FileName.Path,instrrev(FileName.Path,"\")+1)%>">
					<input type="submit" style="width:6em; height:1.75em" name="AskDelete" value="Delete">
				</td>
			</form>			
		</tr>
		<%
		intNoOfFiles = intNoOfFiles + 1
    NEXT
    
    SET oFolder = nothing
   
	IF intNoOfFiles = 0 THEN %>
		<tr align="center">
			<td colspan="2" class="text">No files available</td>		
			<td colspan="1" class="text">&nbsp;</td>
		</tr><%
	END IF
  
	%>
    </TABLE>

	</body>
	</html>
	<%

   
END SUB



' ---------------------------
   SUB askDelete(strFileName)
' ---------------------------

' ------------------------------------------------------
' --- Ask if to delete this file
' --- INPUT : the file name to be deleted, less the path
' -------------------------------------------------------
	%>
	<html>
	<head>
	<title>Delete file y/n?</title>
	</head>
	<body>
	
	<form name="frmConfirmDelete" method="post" action="<%=Request.ServerVariables("PATH_INFO")%>">
	<table border="0" align="center">
		<tr>
			<td class="text">
				Are you sure you wish to delete <b><%=strFileName%></b> ?
			</td>
		</tr>
		<tr>
			<td class="text" align="center">
				<br>
				<input type="hidden" name="fileId" value="<%=strFileName%>">
				<INPUT type="submit" value="Yes" name="Delete" style="width:6em; height:1.75em">
				&nbsp;&nbsp;
				<INPUT type="submit" value="No" name="Delete" style="width:6em; height:1.75em">
			</td>
		</tr>
	</table>
	</form>

	</body>
	</html>
	<%

END SUB




' -------------------------
  SUB delete(strFileName)
' -------------------------

' ------------------------------------------------------
' --- Deletes the file given the full file name strFileName
' --- INPUT : the file name to be deleted, less the path
' ------------------------------------------------------


	'Response.write("<br> Delete - ="&GRImportedFolder & "\" & strFileName) 
	'Response.End 

	Dim oFS, a

    	SET oFS = Server.CreateObject("Scripting.FileSystemObject")
	a = oFS.DeleteFile(GRImportedFolder & "\" & strFileName)

	SET oFs = nothing
	SET a = nothing	
	'Response.End 


END SUB



'--------------------------------------------
call main()

%>

