<!--#include file="settingsHQ.asp"-->
<!--#include file="tools_include.asp"-->
<!--#include file="clsUpload.asp"-->
<%

' --- Upload was ToFileSystem.asp


Dim ThisFileName, pvar
Dim objUpload, objfso, strFileName, strPath
Dim RawIWSFScoresTableName
Dim IWSF_File

ThisFileName="TeamTrials.asp"

PathtoIWSFUploads=Server.mappath("/rankings/teams/")
PathtoExceptions = Server.mappath("/rankings/teams/")
PathtoReasons = Server.mappath("/rankings/teams/")
RawIWSFScoresTableName="usawsrank.TeamsIWSFRaw"

sRunByWhat=Request("pvar")


SELECT CASE sRunByWhat
	CASE "selectfile"
		SelectFile
	CASE "upload"
		UploadTheFile
	CASE "verify"
		VerifyUploadedFile
	CASE "processfile"
		Process_IWSF_File

END SELECT



'-----------------
 SUB SelectFile
'-----------------


    WriteIndexPageHeader
    
    %>
    <center>
    <H2>Upload IWSF File</H2>

    
    <H4>Step 1: Browse to locate an IWSF file on your hard drive.<br><br>
    Step 2: Upload the selected file to add your IWSF file to the World Scores table.<br>
    </h4>
    <br><br>    
    <FORM method="post" encType="multipart/form-data" action="\rankings\<%=ThisFileName%>?pvar=UpLoadTheFile">
    	<INPUT type="File" name="File1">
    	<INPUT type="Submit" value="Upload">
    </FORM>
    
    <%
    
    WriteIndexPageFooter
    
    
END SUB



' ------------------
   SUB UploadTheFile
' ------------------

set objFSO=server.createobject("scripting.filesystemObject")

' Instantiate Upload Class
Set objUpload = New clsUpload

' Grab the file name
strFileName = objUpload.Fields("File1").FileName

' Compile path to save file to
'strPath = Server.MapPath("/rankings/uploads/") & "\" & strFileName
strPath=PathtoIWSFUploads & "\" & strFileName



IF objfso.FileExists(strPath) = true Then
  	' Reject upload if the file already exists
	'WriteLog(date() &"  "& time() &"  "& strpath & " duplicate file upload attempted.  File rejected.")
	SET objfso=nothing
	SET objupload=nothing

	WriteIndexPageHeader %>

	<html><head><title>File Upload Failed</title></head><body>
	<br><br>
	<center><h2><font color="red">The file <%=strFileName%> already exists.</font></h2><br><br><br>

	<h4>Please check the file name or contact your Mark Crone or Headquarters for further assistance.
	<br><br><br>
	</h4></center>
	<br><br>
	</body></html><%

	WriteIndexPageFooter
ELSE
	' Save the binary data to the file system if it doesn't exist
	objUpload("File1").SaveAs strPath

	' Release upload object from memory
	SET objfso = Nothing
	SET objUpload = Nothing


	' --- Runs verififcation of upload ---
	VerifyUpLoadedFile strfilename

END IF

END SUB



' --------------------------
  SUB VerifyUploadedFile (file)
' --------------------------

' --- Verifies format of the uploaded file ---

IF file = "" THEN 
  	WriteIndexPageHeader %>

	<center>
	<br><br>
	<h3><font color="red">No file specified for upload.</font></h3>
	<br><br>
	<font color="red">Please try again.</font>
	<br><br>
	</center><%

	WriteIndexPageFooter
ELSE

	Response.Buffer = True
    
	' Ran into some problems with large files (particularly the Nationals File)
	' Where the time out expired before the server finished processing.
	' 300 seconds (5 minutes) seems to be plenty of time.
    
	Server.ScriptTimeout = 300 
    
	' The following lines of HTML display the "please wait" banner. %>
    
    
	<html><head><title>Please Wait...</title>

    <SCRIPT LANGUAGE="JavaScript">
    // First we detect the browser type
    if(document.getElementById) { // IE 5 and up, NS 6 and up
    	var upLevel = true;
    	}
    else if(document.layers) { // Netscape 4
    	var ns4 = true;
    	}
    else if(document.all) { // IE 4
    	var ie4 = true;
    	}
    
    function showObject(obj) {
    if (ns4) {
    	obj.visibility = "show";
    	}
    else if (ie4 || upLevel) {
    	obj.style.visibility = "visible";
    	}
    }
    
    function hideObject(obj) {
    if (ns4) {
    	obj.visibility = "hide";
    	}
    if (ie4 || upLevel) {
    	obj.style.visibility = "hidden";
    	}
    }
    
    </SCRIPT>
    </head>
    <body>
    <DIV ID="splashScreen" STYLE="position:absolute;z-index:5;top:30%;left:35%;">
    <TABLE BGCOLOR="#000000" BORDER=1 BORDERCOLOR="#000000"	CELLPADDING=0 CELLSPACING=0 HEIGHT=200 WIDTH=300>
    <TR>
    <TD WIDTH="100%" HEIGHT="100%" BGCOLOR="#CCCCCC" ALIGN="CENTER" VALIGN="MIDDLE">
    <BR><BR>
    <FONT FACE="Helvetica,Verdana,Arial" SIZE=3 COLOR="#000066">
    <B>Processing File Upload.<br><br>
    Please wait a moment ...<br><br>  
    </B></FONT>
    <IMG SRC="images/buttons/wait.gif" BORDER=1 WIDTH=75 HEIGHT=15><BR><BR>
    </TD>
    </TR>
    </TABLE>
    </DIV><%


	response.flush
    
	' Once the "please wait" banner is written to HTML, we flush the response
	' buffer to make the page appear to the users browser while the rest of the
	' script processing takes place.
    
    	Dim RoundNum
	Dim filein, IWSF_File
	Dim objfso, objstreamin, fileoutgood, objstreamgood, fileoutbad, fileoutexplainations
	Dim objstreambad, objstreamexplainations, errorcheck, goodrec, badrec
	Dim SkiYearID
	Dim tempFed, tempFName, tempLName, tempBirthdate, tempGender, tempSkiYear
	Dim PDF_Div, AgeInYears
	Dim ValidDivs, DivArray, i
	Dim ZBSFactor
    
	Opencon
	SET rs=Server.CreateObject("ADODB.recordset")
    
    	errorcheck = 0
    	goodrec = 0
    	badrec = 0
   

	filein=PathtoIWSFUploads&"\"& request("file")  
	IWSF_File=PathtoIWSFUploads&"\"& left(Request("file"),7) & ".csv"  

	' --- Change these to badfiles (failed) and imported (good)
	fileoutbad = Server.MapPath("/rankings/badfiles/") & "\" & Request("file") & "--" & month(Date) & "-" & day(date) & "-" & year(date) & "--" & left(FormatDateTime(Now, 4),2) &"-"& right(FormatDateTime(Now,4),2) & right(Request("file"),4)
	fileoutgood = Server.MapPath("/rankings/imported/") & "\" & request("file") & "--" & month(Date) & "-" & day(date) & "-" & year(date) & "--" & left(FormatDateTime(Now, 4),2) &"-"& right(FormatDateTime(Now,4),2) & right(request("file"),4)




	SET objFSO=server.createobject("scripting.filesystemObject")
    

         ' ************************************************
         ' Check if the file name matches a valid tour ID - NOT APPLICABLE IN TEAM PROGRAM
         ' ************************************************
    
	'     If not, set errorcheck bit and copy the file to the bad-upload section.
    
	'sSQL = "Select top 1 * from "&SanctionTableName&" where upper(TournAppID) = '" & ucase(left(Request("file"),6)) & "'"
	'rs.open sSQL, sConnectionToSanctionTable
	'IF rs.EOF THEN
	'	objfso.CopyFile filein, fileoutbad, 1
	'	objfso.DeleteFile filein
	'	errorcheck = 1
	'END IF
	'rs.Close
    


	IF UCASE(RIGHT(file,4)) <> ".TXT" THEN
       		' Only move the file if it hasn't already been moved.
		IF ErrorCheck = 0 THEN
			objfso.copyfile filein, fileoutbad, 1
			objfso.deletefile filein
		END IF
		errorcheck = 2
	END IF


' **********  WHERE ARE THESE ROUTINES ???? - Verify_Upload.asp ???  ************
    
	If ErrorCheck = 1 Then BadSanctionCode
	If ErrorCheck = 2 Then BadFileExtension
    	If ErrorCheck <> 0 Then BadFile
    


' **********  CHANGE PATH  ************

	fileoutbad=PathtoExceptions & "\exceptions-" & file & "--" & month(Date) & "-" & day(date) & "-" & year(date) & "--" & left(FormatDateTime(Now, 4),2) &"-"& right(FormatDateTime(Now,4),2) & right(IWSF_File,4)
	fileoutexplainations=PathtoReasons & "\exceptions-" & file & "--" & month(Date) & "-" & day(date) & "-" & year(date) & "--" & left(FormatDateTime(Now, 4),2) &"-"& right(FormatDateTime(Now,4),2) & right(IWSF_File,4)
    
    
	' --- If no errors were found, then OK to process the file ---
	If ErrorCheck = 0 and UCASE(RIGHT(IWSF_File,4)) = ".TXT"  THEN Process_IWSF_File
    
	' If good records were found, close the "good file" object stream.
	If GoodRec > 0 Then 
        	objstreamgood.close
	End If


	' --- If bad records were found, close the "bad file" object stream.
	' --- Also close the "explanations file" object stream.
	IF BadRec > 0 Then
		objstreambad.close
		objstreamexplainations.close
	END IF

	' --- Finally, close the "in file" object stream.
	objstreamin.close
    
	'WriteLog(date() &"  "& time() &"  "& filein & " has been processed through verification. " & goodrec & " good recs and " & badrec & " bad recs.")
    
    
	Response.Flush
      
	' This final bit of HTML is written after processing is successfully completed
	' to show the user that processing was successful and also how many
	' good and bad records were discovered inside the IWSF file. %>
    
       <SCRIPT LANGUAGE="JavaScript">
       if(upLevel) {
         var splash = document.getElementById("splashScreen");
       }
       else if(ns4) {
         var splash = document.splashScreen;
       }
       else if(ie4) {
         var splash = document.all.splashScreen;
       }
      
       hideObject(splash);
       </SCRIPT><%

	WriteIndexPageHeader%>
        <br><br>
        <center><h2>The file <%=IWSF_File%> has been uploaded successfully.</h2><br><br><br>
    
        <h4>IWSF File is Ready for Additional Processing
        <br><br><br>
        <font color="red"><%=goodrec%></font> out of <font color="red"><%=(goodrec + badrec)%></font> score records were successfully imported.
        <br><br><br>
        The score file had <font color="red"><%=badrec%></font> records which failed verification.
        <br><br><br>
        </h4></center>
        <br><br>
        </body></html><%

	WriteIndexPageFooter 
       
	' KickTrafficCounter("ScoreUpLds")   

END IF ' This big loop checks if there was a file uploaded or not.  We don't process if there is no file uploaded.



END SUB



' ---------------------
  SUB Process_IWSF_File
' ---------------------


  ' --- Using ADO and the OLEDB recordset method.

  filein=PathtoIWSFUploads&"\" & IWSF_File  
'  IWSF_File=PathtoIWSFUploads&"\" & left(IWSF_File,7) & ".csv"

'  objfso.CopyFile filein, IWSF_File, 1

  	
  Set objstreamin=objFSO.OpenTextFile(filein)

  ' --- Reads in FIRST line of IWSF txt file ---	
  IF NOT objstreamin.atendofStream THEN lineText=objstreamin.ReadLine



' ********* NEEDS WORK ??????????  **********
Dim sDSN
sDSN = "FileDSN=" & PathToTRA & "IWSFDelim.DSN;DefaultDir=" & PathtoIWSFUploads & "\;DBQ=" & PathtoIWSFUploads & "\;Extensions=txt;"

Dim ConTest, rsIWSF
Set ConTest = Server.CreateObject("ADODB.Connection")
Set rsIWSF=Server.CreateObject("ADODB.recordset")
ConTest.Open sDSN

Dim sSQL
sSQL = "Select * from " & IWSF_File
rsIWSF.open sSQL, sDSN

DO WHILE NOT rsIWSF.EOF


	' --- Parces records from line of IWSF txt file ---
	tempLast=SQLClean(ucase(rsIWSF.fields(0)))
	tempFirst=SQLClean(ucase(rsIWSF.fields(1)))
	tempFed=SQLClean(ucase(rsIWSF.fields(2)))
	tempGender=SQLClean(ucase(rsIWSF.fields(3)))
	tempTourID=SQLClean(ucase(rsIWSF.fields(4)))
	tempSlScore=SQLClean(ucase(rsIWSF.fields(5)))
	tempTrScore=SQLClean(ucase(rsIWSF.fields(6)))
	tempJuScore=SQLClean(ucase(rsIWSF.fields(7)))
	tempYrBirth=SQLClean(ucase(rsIWSF.fields(8)))
	tempClass=SQLClean(ucase(rsIWSF.fields(9)))
	tempRound=SQLClean(ucase(rsIWSF.fields(10)))
	tempDiv=SQLClean(ucase(rsIWSF.fields(11)))
	tempPerfQual1=SQLClean(ucase(rsIWSF.fields(12)))
	tempPerfQual2=SQLClean(ucase(rsIWSF.fields(13)))
	tempEndDate=SQLClean(ucase(rsIWSF.fields(14)))
	tempJunior=SQLClean(ucase(rsIWSF.fields(15)))
	tempUnknown=SQLClean(ucase(rsIWSF.fields(16)))
	tempIWSFMemberID=SQLClean(ucase(rsIWSF.fields(17)))

   

	' --- Inserts parced records into table ---
	sSQL = "INSERT INTO " & RawIWSFScoresTableName
	sSQL = sSQL + " (Last, First, Fed, Sex, TourID, SlScore, TrScore, JuScore, YrBirth, Class"
	sSQL = sSQL + " , Round, Div, PerfQual1, PerfQual2, EndDate, Junior, Unknown, IWSFMemberID)"
	sSQL = sSQL + " VALUES ("
	sSQL = sSQL + "'" & tempLast & "',"
	sSQL = sSQL + "'" & tempFirst & "',"
	sSQL = sSQL + "'" & tempFed & "',"
	sSQL = sSQL + "'" & tempSex & "',"
	sSQL = sSQL + "'" & tempTourID & "',"
	sSQL = sSQL + "'" & tempSlScore & "',"
	sSQL = sSQL + "'" & tempTrScore & "',"
	sSQL = sSQL + "'" & tempJuScore & "',"
	sSQL = sSQL + "'" & tempYrBirth & "',"
	sSQL = sSQL + "'" & tempClass & "',"
	sSQL = sSQL + "'" & tempRound & "',"
	sSQL = sSQL + "'" & tempDiv & "',"
	sSQL = sSQL + "'" & tempPerfQual1 & "',"
	sSQL = sSQL + "'" & tempPerfQual2 & "',"
	sSQL = sSQL + "'" & tempEndDate & "',"
	sSQL = sSQL + "'" & tempJunior & "',"
	sSQL = sSQL + "'" & tempUnknown & "',"
	sSQL = sSQL + "'" & tempIWSFMemberID & "')"

	Con.Execute(sSQL)

       ' --- Saved our record ... so now we read the next line (assuming there is one)
	GoodRec = GoodRec + 1
	IF NOT objstreamin.atendofStream THEN lineText=objstreamin.readline

	rsIWSF.MoveNext   '--- Move to the next record
Loop


'Close our recordset and connection
rsIWSF.close
Set rsIWSF = Nothing
conTest.close
set conTest = nothing

objfso.DeleteFile IWSF_File



END SUB

%>




