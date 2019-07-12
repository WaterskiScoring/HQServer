<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_registration.asp"-->
<!--#include file="includes/clsUpload.asp"-->
<%





' --- Author       : Mark Crone
' --- Macro purpose: To add record to SQL database using ADO and SQL
' --- NOTE:  Reference to Microsoft ActiveX Data Objects Libary required


' --- Dimension variables ---
Dim ExcelPath, ExcelFile
Dim dbPath, tblName 
Dim rngColHeads, rngTblRcds, colHead, rcdDetail, ch, cl, notNull
DIM sTourID, GRScoresTableName

DIM sFirstCol, sLastCol, sMemberIDCol, sDivisionTempCol, sEventCol, sScoreCol, sPlaceCol
DIM sClassCol, sRoundCol, sProAmPointsCol, sSptsGrpIDCol, sSkiYearIDCol, sTeamCol, sEndDateCol

DIM sMemberID, sMembSex, sMembAge, sFirst, sLast, sEvent, sEventTemp, sDivisionTemp, sDivision, sScore, sPlace, sDivGroup
DIM sClass, sRound, sProAmPoints, sSkiYearID, sTeam, sEndDate

DIM ValidMember, ValidEvent, ValidDiv, ValidScore, ValidPlace, ValidRound, LayoutOK
Dim ResponseString

Dim ImportCount, InvalidCount, FoundCount

Dim rsMem, rsScr, sAction, sFileID
Dim objExcelRS, objExcelConn
Dim MainImage

' --- Define variables
ImportCount=0
InvalidCount=0
FoundCount = 0



ThisFileName="GRExcelImport.asp"

' --- Names of folders (note there is no / at end)
GRImportedFolder=Server.mappath("/rankings/GR_Results/Imported")
GRImportedFolderFullPath="http://usawaterski.org/rankings/GR_Results/Imported"
GRUploadedFolder=Server.mappath("/rankings/GR_Results")
GRUploadedFolderFullPath="http://usawaterski.org/rankings/GR_Results"



ScoresGRTableName = "usawsrank.ScoresGR"

sAction=LCASE(Request("Action"))



ExcelFile=request("fileId")
Excelpath = Server.MapPath("gr_results/") 




WriteIndexPageHeader

ReportTitle="Grassroots Import Program"








SELECT CASE sAction

    CASE "import"
	DisplaySelectBox
	ImportTheUploadedFile
	

    CASE "confirm delete"
        'response.write("<br>READY TO DELETE")
	DeleteTheImportedFile

	DisplaySelectBox

    CASE "delete"
	ConfirmToDeleteFile
 
    CASE "import complete"
	DisplaySelectBox

	CopyOverTheImportedFile
	response.redirect("/rankings/GRUploadform.asp")

    CASE "done"

	'objExcelRS.Close
	'objExcelConn.Close

	response.redirect("/rankings/GRUploadform.asp")

    CASE ELSE
	DisplaySelectBox

	' --- Do nothing						

END SELECT



' -------------------------------------------------------------------------------------------------------------------------
' --- Bottom of main program
' -------------------------------------------------------------------------------------------------------------------------



' --------------------------
  SUB ImportTheUploadedFile
' --------------------------


	' --- Set the connections to the Excel sheet
	SetExcelConnection

	' --- Gets the SkiYearID based on the default year settings ---
	SET rs = Server.CreateObject("ADODB.RecordSET")
	rs.open "SELECT SkiYearID FROM "&SkiYearTableName&" WHERE DefaultYear='1'", sConnectionToTRATable, 3, 1
	IF NOT rs.eof THEN sSkiYearID = rs("SkiYearID")


	%><br>
	  <font size="3"><b>Import Results Summary</b></font><%

	' --- Validate TourID ---
	ValidateTour

	IF sValidTour THEN


	    ' --- Validates that columns are where they are supposed to be
	    ValidateColumns
	    IF LayoutOK THEN

		objExcelRS.MoveNext
		' --- Validate and read in the data ---
		ReadData	

	    END IF

	END IF	

	' --- Displays the Total Found, Total Imported and Total Invalid Records ---
	DisplayImportSummary

	%>
	<form action="/rankings/<%=ThisFileName%>" method="post" >
		<input type="hidden" name="fileID" value="<%=ExcelFile%>">

		<br>
		<input type="submit" style="width:10em;" value="Import Complete" name="Action">
		&nbsp;&nbsp;
		<input type="submit" style="width:10em;" value="Changes Needed" name="Action">
		<br><br>
	</form>
	<%


ListImportedTours


END SUB




' ---------------------------
  SUB CopyOverTheImportedFile
' ---------------------------

	Dim oFS, a

	' --- Copies it from UPLOADED folder into the IMPORTED folder ---
    	SET oFS = Server.CreateObject("Scripting.FileSystemObject")
	a = ofs.CopyFile (GRUploadedFolder& "\" &ExcelFile, GRImportedFolder& "\" &ExcelFile)

	' --- Deletes it from the UPLOADED ---
    	SET oFS = Server.CreateObject("Scripting.FileSystemObject")
	a = oFS.DeleteFile(GRUploadedFolder & "\" &ExcelFile)

	Set oFs = nothing
	Set a = nothing	


END SUB



' ---------------------------
  SUB DeleteTheImportedFile
' ---------------------------

	Dim oFS, a

	' --- Deletes it from the UPLOADED ---
    	SET oFS = Server.CreateObject("Scripting.FileSystemObject")
	a = oFS.DeleteFile(GRUploadedFolder & "\" &ExcelFile)

	Set oFs = nothing
	Set a = nothing	


END SUB


' ---------------------------
  SUB ConfirmToDeleteFile
' ---------------------------


   %>
	<html>
	<head>
	<title>File Upload</title>

	</head>
	<body>

<form action="/rankings/<%=ThisFileName%>" method="post" >
  <input type="hidden" name="fileID" value="<%=ExcelFile%>">

   <TABLE class="innertable" width=70% align=center>
     <TR>
       <th align=center colspan=2>
          <font size="3" color="white"><b>Confirm To Delete This Uploaded File</b></font>
       </th> 
     </TR>

     <TR>
       <TD colspan=2 align=center>
          <br><br>
          <font size="2" ><b>Excel File Name: <%=ExcelFile%></b></font>
	  <br><br>
          <font size="1" >Pressing 'Confirm Delete' will remove this Excel Sheet from the uploaded list<br><br> Press 'Cancel' to return to uploaded list </font>
	  <br><br>
       </TD>
     </TR>

     <TR>
       <TD align=center>
	<br>
	<input type="submit" style="width:10em;" value="Confirm Delete" name="Action">
	<br><br>
       </TD>

       <TD align=center>
	<br>
	<input type="submit" style="width:10em;" value="Cancel" name="Action">
	<br><br>
       </TD>
     </TR>

   </TABLE>
</form>

<%

END SUB


' ------------------------
  SUB SetExcelConnection
' ------------------------

	' --- Open connection and recordset for spreadsheet ---
	SET objExcelConn = Server.CreateObject("ADODB.Connection")
	SET objExcelRS = Server.CreateObject("ADODB.RecordSET")


	objExcelConn.Provider = "Microsoft.Jet.OLEDB.4.0"
	objExcelConn.Properties("Extended Properties").Value = "Excel 8.0;HDR=No; IMEX=1"
	objExcelConn.Open ExcelPath &"\"& ExcelFile

	objExcelRS.Open "SELECT * FROM [Results$A1:J50];", objExcelConn
	objExcelRS.MoveFirst


END SUB





' -------------------------
  SUB DisplayImportSummary
' -------------------------

	response.write("<br><br><b>SUMMARY:</b>")
	response.write("<br> Total Found - "&FoundCount)
	response.write("<br> Total Imported - "&ImportCount)
	response.write("<br> Total Invalid Records - "&InvalidCount)



END SUB




' ---------------------
  SUB DisplaySelectBox 
' ---------------------

    Dim oFS, oFolder, intNoOfFiles, FileName


    ' --- Defines file object and then loads the list of files in the specified folders into an array ---	

    Set oFs = nothing
    Set oFS = Server.CreateObject("Scripting.FileSystemObject")
    Set oFolder = nothing
    Set oFolder = oFS.getFolder(GRUploadedFolder)


	%>
	<html>
	<head>
	<title>File Upload</title>

	</head>
	<body>

	<form action="/rankings/<%=ThisFileName%>" method="post" >

	<TABLE class="innertable" width=500px align="center" style="padding:2px; border-collapse:collapse; border:1px solid <%=HQSiteColor2%>;" >
	  <tr>
		<th colspan="3" class="text" align="center" colspan="3">
			<font size="3" color="white"><b>Import Grassroots Excel Document</b></font>
		</th>		
	  </tr>

	  <tr>
	    <td align="left" style="border-style:none;" width=50px>&nbsp;</td>

	    <td align="left" style="border-style:none;">
		<font size="1"><b> Options:</b><br>&nbsp;&nbsp;- Import a File From the List Below<br>&nbsp;&nbsp;- Click Link to Download Spreadsheet<br>&nbsp;&nbsp;- Press Done to Return to Menu</font>
	    </td>
	    <td align="center" style="border-style:none;">
		<br>
		<input type="submit" style="width:6em;" value="Done" name="Action">
		<br><br>
	    </td>
	  </tr><%

	IF sAction="select tournament" THEN %>
	  <tr>
	   <td colspan=3 align="center">
		<INPUT type="File" name="ZipFile" size="40">
	   </td>
	 </tr><%
	END IF %>

	</TABLE>


	<TABLE class="innertable" width=500px align="center">
	  <tr>
	    <td class="text" align="left" colspan="2"><font size=2><b>Filename</b></font></td>
	    <td class="text" colspan=2 align="center"><font size=2><b>Action</b></font></td>
	  </tr><%


	intNoOfFiles = 0
	Dim CurrFile

        FOR Each FileName IN oFolder.Files	
	   CurrentFile=mid(FileName.Path,instrrev(FileName.Path,"\")+1)

		%>
		<form Name="ReadyToImportFiles" method="post" action="/rankings/<%=ThisFileName%>">
		<input type="hidden" name="fileId" value="<%=CurrentFile%>">

		<tr>		
			<td class="text" align="left" colspan="2">
				<font size="1">
				<a href="<%=GRUploadedFolderFullPath%>/<%=CurrentFile%>" title="Download the Excel document <%=CurrentFile%>" target="_blank"><%=CurrentFile%></a>
			</td>
			<td class="text" align="center">
				<input type="submit" style="width:8em; height:1.75em" name="Action" value="Import" title="Import <%=CurrentFile%> into the Scores database">
			</td>
			<td class="text" align="center">
				<input type="submit" style="width:8em; height:1.75em" name="Action" value="Delete" title="Delete <%=CurrentFile%> from the uploaded files">
			</td>
		</tr>
		</form><%

		intNoOfFiles = intNoOfFiles + 1
        NEXT
    
        SET oFolder = nothing
   
	IF intNoOfFiles = 0 THEN %>
		<tr align="center">
			<td colspan="2" class="text">No files available</td>
		</tr><%
	END IF 	%>  

	</TABLE><%






END SUB


' --------------
  SUB ReadData
' --------------


response.write("<br><br>STEP 3 - Reading Score Details<br>")


' --- Insert records into database from worksheet table ---
DO WHILE NOT objExcelRS.EOF


	sFirst=""
	sLast=""
	sMemberID=""
	sDivisionTemp=""
	sPlace=""
	sEvent=""
	sScore=""

	
	' --- Evaluate each field in the record/row
	FOR col = 0 TO 8
	

		notNull=false
		' --- There is a value in the LAST name column ---
		IF TRIM(objExcelRS.Fields(1).Value)<>"" THEN
			' --- If not empty, set notNull to true, and append value to string
			notNull = True

			IF col = sFirstCol THEN sFirst = objExcelRS.Fields(col).Value
			IF col = sLastCol THEN sLast = objExcelRS.Fields(col).Value
			IF col = sMemberIDCol THEN sMemberID = objExcelRS.Fields(col).Value
			IF col = sDivisionTempCol THEN sDivisionTemp = objExcelRS.Fields(col).Value
			IF col = sPlaceCol THEN sPlace = objExcelRS.Fields(col).Value
			IF col = sEventCol THEN sEventTemp = objExcelRS.Fields(col).Value
			IF col = sScoreCol THEN sScore = objExcelRS.Fields(col).Value
			IF col = sProAmPointsCol THEN sProAmPoints = objExcelRS.Fields(col).Value
			IF col = sTeamCol THEN sTeam = objExcelRS.Fields(col).Value
			IF col = sClassCol AND sClass<>"" THEN sClass = objExcelRS.Fields(col).Value
			IF col = sRoundCol AND sClass<>"" THEN sRound = objExcelRS.Fields(col).Value

		END IF

	NEXT





        ' --- If record consists of only Null values do not insert it to table otherwise insert the record ---
	sSQL = ""
	IF notNull THEN

		ResponseString=""
		FoundCount =  FoundCount + 1

		' --- Validates Member against Membership table
		ValidateMember

		IF ValidMember = true THEN

			' --- Validate as an allowed event
			ValidateEvent

			' --- Validates and converts to an allowed division
			ValidateDivision

			' --- Checks for fractional scores and converts ---
			ValidateScore

			' --- Checks for a valid PLACE value ---
			ValidatePlace

			' --- Checks for a valid ROUND value ---
			ValidateRound

		END IF

		IF ValidMember AND ValidDiv AND ValidEvent AND ValidScore AND ValidPlace AND ValidRound THEN

			' --- Checks GR Score file and then imports ---
			InsertTheScore
		ELSE
			response.write(ResponseString)
		END IF


        END IF


	objExcelRS.MoveNext

LOOP



END SUB



' ------------------
  SUB ValidateMember
' ------------------


' --- Displays MemberID ---
sMemberID = replace(sMemberID,"-","")

SET rsMem = Server.CreateObject("ADODB.RecordSET") 

ValidMember = false

ResponseString="<br>"&FoundCount&") MemberID: "&sMemberID

' --- Tests for Numeric value in field --- 
IF IsNumeric(sMemberID) THEN

  rsMem.open "SELECT PersonIDWithCheckDigit, Sex FROM "&MemberTableName&" WHERE PersonIDWithCheckDigit="&sMemberID&"", sConnectionToTRATable, 3, 1

  IF NOT rsMem.eof THEN 
	sMembSex=rsMem("Sex")
	ValidMember = true
	ResponseString=ResponseString & " - MemberID Found"
  ELSE
	ResponseString=ResponseString & " - Invalid MemberID"	
	InvalidCount = InvalidCount + 1
  END IF

ELSE
	ResponseString=ResponseString & " - ERROR: Text found in MemberID column"	
	InvalidCount = InvalidCount + 1

END IF


END SUB




' -------------------
  SUB ValidateScore
' -------------------

'response.write("<br>sScore="&sScore)


'sScore="61 1/2"

	ValidScore=false
	SELECT CASE RIGHT(sScore,3)
		CASE "1/2" 
			sScore=CDbl(TRIM(LEFT(sScore,LEN(sScore)-3))+0.50)
			ValidScore=true

		CASE "1/4"
			sScore=CDbl(TRIM(LEFT(sScore,LEN(sScore)-3))+0.25)
			ValidScore=true
		CASE ELSE
			IF isNumeric(sScore) THEN ValidScore=true

	END SELECT


	IF ValidScore = false THEN
		InvalidCount = InvalidCount + 1
		ResponseString=ResponseString & " --- Invalid Score "
	END IF

END SUB




' -------------------
  SUB ValidateEvent
' -------------------

	ValidEvent=false
	SELECT CASE TRIM(LCASE(sEventTemp))
		CASE "slalom", "s"
			sEvent = "S"
			ValidEvent=true
		CASE "trick", "tricks", "t"
			sEvent = "T"
			ValidEvent=true
		CASE "wakeboard", "wb"
			sEvent = "WB"
			ValidEvent=true
		CASE "kneeboard", "kb"
			sEvent = "KB"
			ValidEvent=true
	END SELECT

	IF ValidEvent = false THEN
		InvalidCount = InvalidCount + 1
		ResponseString=ResponseString & " --- Invalid Event "
	END IF

END SUB




' ------------------
  SUB ValidatePlace
' ------------------

	ValidPlace=true
	IF IsNumeric(sPlace) OR TRIM(sPlace)="" THEN 
		' --- OK
	ELSE
		ValidPlace=false
	END IF


	IF ValidPlace = false THEN
		InvalidCount = InvalidCount + 1
		ResponseString=ResponseString & " --- Invalid Place "
	END IF


END SUB



' ------------------
  SUB ValidateRound
' ------------------

'response.write("<br> Test of ValidRound - "&LCASE(RIGHT(sRound,2)))
'response.write("<br> Test True ? = ")
'response.write(LCASE(RIGHT(sRound,2))="st" OR LCASE(RIGHT(sRound,2))="nd" OR LCASE(RIGHT(sRound,2))="rd"  OR LCASE(RIGHT(sRound,2))="th")

	ValidRound=true
	IF IsNumeric(sRound) THEN
		' --- OK
        ELSEIF LCASE(RIGHT(sRound,2))="st" OR LCASE(RIGHT(sRound,2))="nd" OR LCASE(RIGHT(sRound,2))="rd"  OR LCASE(RIGHT(sRound,2))="th" THEN
		sRound = Cint(LEFT(sRound,LEN(sRound)-2))
	ELSE
		ValidRound=false
	END IF


	IF ValidRound = false THEN
		InvalidCount = InvalidCount + 1
		ResponseString=ResponseString & " --- Invalid Round "
	END IF


END SUB




' ------------------
  SUB ValidateDivision
' ------------------

	sMembAge = AgeAtDate(sTDateS, sMemberID)
	sDivGroup=sDivisionTemp
	sMembSex=LCASE(sMembSex)



	ValidDiv=true
	IF LCASE(sMembSex)="male" AND Cdbl(sMembAge)<cdbl(18) THEN 
		sDivision="B"
	ELSEIF LCASE(sMembSex)="male" AND Cdbl(sMembAge)>cdbl(18) THEN
		sDivision="M"
	ELSEIF LCASE(sMembSex)="female" AND Cdbl(sMembAge)<cdbl(18) THEN 
		sDivision="G"
	ELSEIF LCASE(sMembSex)="female" AND Cdbl(sMembAge)>cdbl(18) THEN
		sDivision="W"
	ELSE
		ValidDiv=false
		sDivision="U"
	END IF

	IF ValidDiv = false THEN
		InvalidCount = InvalidCount + 1
		ResponseString=ResponseString & " --- Invalid Division "
	END IF




END SUB





' -------------------
  SUB InsertTheScore
' -------------------

	sSQL = "SELECT Score FROM "&ScoresGRTableName
	sSQL = sSQL + " WHERE MemberID='"&sMemberID&"' AND TourID='"&sTourID&"' AND Round='"&sRound&"' AND Event='"&sEvent&"' AND Div='"&sDivision&"'"
	SET rsScr = Server.CreateObject("ADODB.RecordSET")
	rsScr.open sSQL, sConnectionToTRATable, 3, 1

	ScoreNotImported = false
	IF rsScr.eof THEN 

		OpenCon
		ScoreNotImported = true

		sSQL = "INSERT INTO " &ScoresGRTableName&" (TourID, MemberID, FName, LName, Event, Div, Score, Place"
		sSQL = sSQL + ", Class, Round, ProAmPoints, SptsGrpID, SkiYearID, Team, EndDate, DivGroup)"
		sSQL = sSQL + " VALUES ('"&sTourID&"', '"&sMemberID&"', '"&sFirst&"', '"&sLast&"', '"&sEvent&"', '"&sDivision&"'"
		sSQL = sSQL + ", '"&sScore&"', '"&sPlace&"', '"&sClass&"', '"&sRound&"', '"&sProAmPoints&"', '"&sSptsGrpID&"'"
		sSQL = sSQL + ", '"&sSkiYearID&"', '"&sTeam&"', '"&sTDateE&"', '"&sDivGroup&"')"

		response.write("<br><br>sSQL = "&sSQL)
		response.write(" --- Score Imported<br>")
		ImportCount = ImportCount + 1

		con.execute(sSQL)
		response.write(" --- Score ADDED")
		
	ELSE
		response.write(" --- Score previously imported!")

	END IF
	



END SUB





' -----------------
  SUB ValidateTour
' -----------------

	' --- Gets TourID from first 5 rows of spreadsheet ---

	DO WHILE NOT objExcelRS.eof
	   FOR col = 0 TO 8
		IF TRIM(LCASE(objExcelRS.Fields(col).Value))="sanction #:" THEN
			sTourID= objExcelRS.Fields(col+1).Value
			EXIT DO
		END IF
	   NEXT
	   objExcelRS.movenext	
	LOOP

	'objExcelRS.Close


	' --- Defines tournament information ---
	DefineTourVariables_New

	IF sValidTour=true THEN
		response.write("<br><br>STEP 1 - TourID "&sTourID&" has been validated")
	ELSE
		response.write("STEP 1 - TourID "&sTourID&" is not valid")
	END IF



	END SUB



' --------------------
  SUB ValidateColumns
' --------------------

	sFirstCol = -1
	sLastCol = -1
	sMemberIDCol = -1
	sTempAgeCol = -1
	sDivisionTempCol = -1
	sPlaceCol = -1
	sEventCol = -1
	sScoreCol = -1
	sTeamCol = -1
	sProAmPointsCol = -1
	sClassCol =-1 
	sRoundCol =-1 


	DO WHILE NOT objExcelRS.EOF
	    FOR col = 0 TO 8

		sColName = TRIM(LCASE(objExcelRS.Fields(col).Value))
		IF Instr(objExcelRS.Fields(col).Value,"Score")>0 AND Instr(objExcelRS.Fields(col).Value,"@")=0 THEN sColName="score"
		IF Instr(objExcelRS.Fields(col).Value,"Membership")>0 OR (Instr(objExcelRS.Fields(col).Value,"USA")>0 AND Instr(objExcelRS.Fields(col).Value,"Mem")>0) THEN sColName="memberid"

'response.write("<br>sColName= "&sColName)
		SELECT CASE sColName
			CASE "first" 
				sFirstCol=col
			CASE "last" 
				sLastCol=col
			CASE "memberid" 
				sMemberIDCol=col
			CASE "age"
				sTempAgeCol=col
			CASE "division"
				sDivisionTempCol=col
			CASE "place"
				sPlaceCol=col
			CASE "event"
				sEventCol=col
			CASE "score"
				sScoreCol=col
			CASE "class"
				sClassCol=col
			CASE "round"
				sRoundCol=col
			CASE "team"
				sTeamCol=col
			CASE "pro-am points"			
				sProAmPointsCol=col
	'		CASE ELSE
	'			response.write("<br>Other - "&TRIM(LCASE(objExcelRS.Fields(col).Value)))
			
		END SELECT
	    NEXT

	    IF sFirstCol<>-1 OR sLastCol<>-1 OR sMemberIDCol<>-1 OR sDivisionTempCol<>-1 OR sEventCol<>-1 OR sScoreCol<>-1 THEN
		EXIT DO
	    END IF	

	    objExcelRS.movenext


	LOOP

	IF sClassCol=-1 THEN sClass="F"
	IF sRoundCol=-1 THEN sRound=1
	'IF sTeamCol=-1 THEN sTeam="ABC"
	

	LayoutOK=true
	IF sFirstCol=-1 AND sLastCol=-1 AND sMemberIDCol=-1 AND sTempDivCol=-1 AND sEventCol=-1 AND sScoreCol=-1 THEN
		LayoutOK=false
		IF sFirstCol=-1 THEN response.write("<br>No accurate columns could be found")

	ELSEIF sFirstCol=-1 OR sLastCol=-1 OR sMemberIDCol=-1 OR sTempDivCol=-1 OR sEventCol=-1 OR sScoreCol=-1 THEN
		IF sFirstCol=-1 THEN response.write("<br>First Name column missing")
		IF sLastCol=-1 THEN response.write("<br>Last Name column missing")
		IF sMemberIDCol=-1 THEN response.write("<br>MemberID column missing")
		IF sDivisionTempCol=-1 THEN response.write("<br>Division column missing")
		IF sEventCol=-1 THEN response.write("<br>Event column missing")
		IF sScoreCol=-1 THEN response.write("<br>Score column missing")

		LayoutOK=false
	END IF

	IF LayoutOK = true THEN
		response.write("<br><br>STEP 2 - Column layout has been validated")
	ELSE
		response.write("<br><br>STEP 2 - *** ERROR *** Column layout could not be validated")
	END IF


END SUB










' ----------------
   SUB EndUpdate
' ----------------

    ' --- Check if error was encounted
    If Err.Number <> 0 Then
        'Error encountered.  Rollback transaction and inform user
        On Error Resume Next
        cnt.RollbackTrans
        MsgBox "There was an error.  Update was not succesful!", vbCritical, "Error!"
    Else
        On Error Resume Next
        cnt.CommitTrans
    End If

    'Close the ADO objects
    cnt.Close
    Set rst = Nothing
    Set cnt = Nothing
    On Error GoTo 0
End Sub



' -----------------------
   SUB ListImportedTours
' -----------------------

	sSQL = "SELECT DISTINCT TourID, TName, TDateS FROM "&ScoresGRTableName&" AS GRS, "&SanctionTableName&" AS ST"


	sSQL = sSQL + " WHERE LEFT(TourID,2) IN ('09','10') AND LEFT(ST.TournAppID,6)=LEFT(GRS.TourID,6)"
	SET rs = Server.CreateObject("ADODB.RecordSET")
	rs.open sSQL, sConnectionToTRATable, 3, 1

	%>
	<br>
	<h1><center><% response.write("Tournaments Previously Imported")%></center></h1>

	<TABLE class="innertable" width=75% align="center">
	  <th><font color="white" size=1>TourID</font></th>
	  <th><font color="white" size=1>Tournament Name</font></th>
	  <th><font color="white" size=1>Date</font></th><%

	DO WHILE NOT rs.EOF
		%>	
		<TR>
		<td><font size=1><%=rs("TourID")%></font></td>
		<td><font size=1><%=rs("TName")%></font></td>
		<td><font size=1><%=rs("TDateS")%></font></td>

		</tr><%
		rs.movenext
	LOOP  %>

	</TABLE><%

END SUB

%>
