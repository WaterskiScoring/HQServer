<!--#include file="settingsHQ.asp"-->
<%
If Request("file") = "" Then 
  WriteIndexPageHeader
  %>
  <center>
  <br><br>
  <h3><font color="red">No file specified for upload.</font></h3>
  <br><br>
  <font color="red">Please try again.</font>
  <br><br>
  </center>
  <%
    WriteIndexPageFooter
  Else

    Response.Buffer = True
    
    ' Ran into some problems with large files (particularly the Nationals File)
    ' Where the time out expired before the server finished processing.
    ' 300 seconds (5 minutes) seems to be plenty of time.
    
    Server.ScriptTimeout = 300 
    
    ' The following lines of HTML display the "please wait" banner.
    %>
    
    
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
    </DIV>
    
    
    
    
    <%
    response.flush
    
    ' Once the "please wait" banner is written to HTML, we flush the response
    ' buffer to make the page appear to the users browser while the rest of the
    ' script processing takes place.
    
    
    
    Dim RoundNum
    Dim filein, CSVfile
    dim objfso
    dim objstreamin
    dim fileoutgood
    dim objstreamgood
    dim fileoutbad
    dim fileoutexplainations
    dim objstreambad
    dim objstreamexplainations
    dim errorcheck
    dim goodrec
    Dim badrec
    Dim SkiYearID
    Dim tempFed, tempFName, tempLName, tempBirthdate, tempGender, tempSkiYear
    Dim PDF_Div, AgeInYears
    Dim ValidDivs, DivArray, i
    Dim ZBSFactor
    
    Opencon
    Set rs=Server.CreateObject("ADODB.recordset")
    
    errorcheck = 0
    goodrec = 0
    badrec = 0
   

'    filein=PathToTRA & "uploads\" & request("file")  
    filein=PathtoUploads&"\"& request("file")  

'    CSVfile=PathToTRA & "uploads\" & left(Request("file"),7) & ".csv"
    CSVfile=PathtoUploads&"\"& left(Request("file"),7) & ".csv"  

    'fileoutbad=PathToTRA & "badfiles\" & Request("file") & "--" & month(Date) & "-" & day(date) & "-" & year(date) & "--" & left(FormatDateTime(Now, 4),2) &"-"& right(FormatDateTime(Now,4),2) & right(Request("file"),4)
    'fileoutgood=PathToTRA & "imported\" & request("file") & "--" & month(Date) & "-" & day(date) & "-" & year(date) & "--" & left(FormatDateTime(Now, 4),2) &"-"& right(FormatDateTime(Now,4),2) & right(request("file"),4)


	' --- Change these to badfiles (failed) and imported (good)
	fileoutbad = Server.MapPath("/rankings/badfiles/") & "\" & Request("file") & "--" & month(Date) & "-" & day(date) & "-" & year(date) & "--" & left(FormatDateTime(Now, 4),2) &"-"& right(FormatDateTime(Now,4),2) & right(Request("file"),4)
	fileoutgood = Server.MapPath("/rankings/imported/") & "\" & request("file") & "--" & month(Date) & "-" & day(date) & "-" & year(date) & "--" & left(FormatDateTime(Now, 4),2) &"-"& right(FormatDateTime(Now,4),2) & right(request("file"),4)

'Markdebug("v_U fileoutgood="&fileoutgood)    


    set objFSO=server.createobject("scripting.filesystemObject")
    
         ' ************************************************
         ' Check if the file name matches a valid tour ID
         ' ************************************************
    
    '     If not, set errorcheck bit and copy the file to the bad-upload section.
    
    sSQL = "Select top 1 * from "&SanctionTableName&" where upper(TournAppID) = '" & ucase(left(Request("file"),6)) & "'"
    rs.open sSQL, sConnectionToSanctionTable
    If rs.EOF Then
       objfso.CopyFile filein, fileoutbad, 1
       objfso.DeleteFile filein
       errorcheck = 1
		ELSE
       If rs("Tstatus") <> 2 AND rs("Tstatus") <> 4 AND rs("Tstatus") <> 5 THEN
          objfso.CopyFile filein, fileoutbad, 1
          objfso.DeleteFile filein
          errorcheck = 2
       END IF
    END IF
    rs.Close
    
    If (ucase(right(request("file"),4)) <> ".PDF") and (ucase(right(request("file"),4)) <> ".WSP") and (ucase(right(request("file"),4)) <> ".CSV") then
       ' Only move the file if it hasn't already been moved.
       If ErrorCheck = 0 Then
         objfso.copyfile filein, fileoutbad, 1
         objfso.deletefile filein
       End If
       errorcheck = 3
    End If
    
    If ErrorCheck = 1 Then BadSanctionCode
    If ErrorCheck = 2 Then BadSanctionStatus
    If ErrorCheck = 3 Then BadFileExtension
    If ErrorCheck <> 0 Then BadFile
    
    fileoutbad=PathtoExceptions & "\exceptions-" & request("file") & "--" & month(Date) & "-" & day(date) & "-" & year(date) & "--" & left(FormatDateTime(Now, 4),2) &"-"& right(FormatDateTime(Now,4),2) & right(request("file"),4)
    fileoutexplainations=PathtoReasons & "\exceptions-" & request("file") & "--" & month(Date) & "-" & day(date) & "-" & year(date) & "--" & left(FormatDateTime(Now, 4),2) &"-"& right(FormatDateTime(Now,4),2) & right(request("file"),4)
    
    
    If ErrorCheck = 0 and (ucase(right(request("file"),4)) = ".PDF") Then ProcessPDF
    If ErrorCheck = 0 and (ucase(right(request("file"),4)) = ".WSP" or ucase(right(request("file"),4)) = ".CSV") Then ProcessWSP
    
      ' If good records were found, close the "good file" object stream.
      If GoodRec > 0 Then 
        objstreamgood.close
      End If
      ' If bad records were found, close the "bad file" object stream.
      ' Also close the "explanations file" object stream.
      If BadRec > 0 Then
        objstreambad.close
        objstreamexplainations.close
      End If
      ' Finally, close the "in file" object stream.
      objstreamin.close
    
      WriteLog(date() &"  "& time() &"  "& filein & " has been processed through verification. " & goodrec & " good recs and " & badrec & " bad recs.")
    
    
      Response.Flush
      
      ' This final bit of HTML is written after processing is successfully completed
      ' to show the user that processing was successful and also how many
      ' good and bad records were discovered inside the PDF.
      
       %>
    
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
       </SCRIPT>  
       <%WriteIndexPageHeader%>
        <br><br>
        <center><h2>The file <%=Request("file")%> has been uploaded successfully.</h2><br><br><br>
    
        <h4>These scores will be included in the rankings<br>following the next overnight Re-Calculation.
        <br><br><br>
        <font color="red"><%=goodrec%></font> out of <font color="red"><%=(goodrec + badrec)%></font> score records were successfully imported.
        <br><br><br>
        The score file had <font color="red"><%=badrec%></font> records which<br>failed verification and will need to be manually corrected.
        <br><br><br>
        </h4></center>
        <br><br>
        </body></html>
       <%
       WriteIndexPageFooter 
       
       KickTrafficCounter("ScoreUpLds")   

End If ' This big loop checks if there was a file uploaded or not.  We don't process if there is no file uploaded.





' ----------------------
   Sub ProcessWSP
' ----------------------

  Dim LineOne

  ' We need two files ... the original WSP which we will move around line by line
  ' using the Object Stream ... and the CSV version which we will open and read
  ' using ADO and the OLEDB recordset method.


' TESTING ONLY - REMOVE WHEN DONE ---


'filein="usawaterski.org\rankings\uploads\" & request("file")  
'CSVfile="usawaterski.org\rankings\uploads\" & left(Request("file"),7) & ".csv"

'filein=PathtoUploads&"\"& "07s160e.wsp"
'CSVfile=PathtoUploads&"\"& "07s160e.csv"


filein=PathtoUploads&"\" & request("file")  
CSVfile=PathtoUploads&"\" & left(Request("file"),7) & ".csv"

'markdebug("From ProcessWSP filein= "& filein)
'markdebug("From ProcessWSP CSVfile= "& CSVFile)


  ' We will delete the CSV when we are done processing.
  objfso.CopyFile filein, CSVfile, 1

  Set objstreamin=objFSO.OpenTextFile(filein)

  '  This looks a little silly, but we actually skip the first line because we
  '  don't need that data for TRAWEB.
  If Not objstreamin.atendofStream Then LineOne=objstreamin.ReadLine
  If Not objstreamin.atendofStream Then lineText=objstreamin.ReadLine



Dim sDSN

sDSN = "FileDSN=" & PathToTRA & "WSPDelim.DSN;DefaultDir=" & PathToUploads & "\;DBQ=" & PathToUploads & "\;Extensions=csv,wsp;"

Dim ConTest, rsCSV
Set ConTest = Server.CreateObject("ADODB.Connection")
Set rsCSV=Server.CreateObject("ADODB.recordset")
ConTest.Open sDSN

Dim sSQL
sSQL = "Select * from " & CSVfile
rsCSV.open sSQL, sDSN

'Print out the contents of our recordset
Do While Not rsCSV.EOF
     errorcheck = ""
     ValidDivs = ""

     ' ****************************
     ' Check MemberID exists in Master DB
     ' ****************************

     sSQL = "Select top 1 * from "&MemberTableName&" where PersonIDwithCheckDigit = '" & SQLClean(rsCSV.fields(1)) & "'"
     rs.open sSQL, sConnectionToTRATable
     if rs.eof then
       tempFed = SQLClean(ucase(rsCSV.fields(0)))
       tempBirthdate = SQLClean(rsCSV.fields(5))
       tempGender = SQLClean(ucase(rsCSV.fields(4)))
       errorcheck = "*" & errorcheck & "Member ID Not Found -- "
     Else
       tempFed = trim(rs("FederationCode"))
       tempBirthdate = rs("BirthDate")
       tempGender = ucase(left(rs("Sex"),1))
     end if

     ' ****************************
     ' Check MemberID matches Member Name
     ' ****************************

     if not rs.eof then
       if ucase(rsCSV.fields(2)) <> ucase(trim(rs("LastName"))) then
         tempLName = SQLClean(ucase(rsCSV.fields(2)))
         errorcheck = "*" & errorcheck & "Member Last Name Incorrect -- "
       Else
         tempLName = SQLClean(trim(rs("LastName")))
       End If
'       if ucase(rsCSV.fields(3)) <> ucase(trim(rs("FirstName"))) then
'         errorcheck = errorcheck & "Member First Name Incorrect -- "
'       end if
       If ucase(left(rsCSV.fields(3),1)) <> left(ucase(trim(rs("FirstName"))),1) Then
         tempFName = SQLClean(ucase(rsCSV.fields(3)))
         errorcheck = "*" & errorcheck & "Member First Name Incorrect -- "
       Else
         tempFName = SQLClean(trim(rs("FirstName")))
       End If
     end if
     rs.Close

' Build a list of valid divisions for this member.

'    First we figure out what ski year they are in
     sSQL = "Select top 1 * from " & SkiYearTableName & " where '"&CDate("01/01/"&(2000+(left(Request("file"),2))))&"' BETWEEN BeginDate and EndDate and SkiYearID <> 1"
     rs.open sSQL, SConnectionToTRATable, 3, 1
     If rs.EOF Then
        sSQL = "Select top 1 * from " & SkiYearTableName & " where DefaultYear = 1"
        rs.open sSQL, SConnectionToTRATable, 3, 1
     End If
     If rs.EOF Then
        tempSkiYear = "SkiYearID" ' We can't find any matching Ski Years so we just trick the SQL into ignoring this condition. (skiyearid = skiyearid)
     Else
        tempSkiYear = rs("SkiYearID")
     End If
     rs.Close ' Close the Ski Year Table
 
'    Next we figure out their age.
     ' NOTE: Division is based on age relative to ski year.
     If tempBirthDate <> "" Then
       ' get absolute number of years 
       AgeInYears = cint(datediff("YYYY", tempBirthDate, CDate("01/01/"&(2000+(left(rsCSV.fields(1).name,2)))))) - 1
       If tempGender = "M" Then
         sSQL = "Select distinct div from " & DivisionsTableName & " where upper(sex) = 'M' and "&AgeInYears&" <= Up_Age and "&AgeInYears&" >= Low_Age and SkiYearID = "& tempSkiYear &" order by Div"
       Else
         sSQL = "Select distinct div from " & DivisionsTableName & " where upper(sex) = 'F' and "&AgeInYears&" <= Up_Age and "&AgeInYears&" >= Low_Age and SkiYearID = "& tempSkiYear &" order by Div"
       End If
     Else
       If tempGender = "M" Then
         sSQL = "Select distinct div from " & DivisionsTableName & " where upper(sex) = 'M' and SkiYearID = "& tempSkiYear &" order by Div"
       Else
         sSQL = "Select distinct div from " & DivisionsTableName & " where upper(sex) = 'F' and SkiYearID = "& tempSkiYear &" order by Div"
       End If
     End If

    rs.open sSQL, SConnectionToTRATable
    If Not rs.EOF Then 
      DivArray = rs.GetRows()
      for i = 0 to ubound(DivArray,2)
        ValidDivs = ValidDivs + DivArray(0,i)
        if i < ubound(DivArray,2) then ValidDivs = ValidDivs + ","
      Next
    End If
    rs.Close

     ' ****************************
     ' Check Tour ID
     ' ****************************
     ' Tour ID in the score record should match the file name.
     if ucase(left(rsCSV.fields(1).name,6)) <> ucase(left(request("file"),6)) then
       errorcheck = errorcheck & "Field 'Tour ID' Doesn't Match Filename -- "
     end if

     ' ****************************
     ' Check For Valid Date
     ' ****************************
     ' Date must be valid or crash will occur.
     If Not isDate(right(left(rsCSV.fields(9).name,6),2) &"/"& right(rsCSV.fields(9).name,2) &"/"& left(rsCSV.fields(9).name,4)) Then
       errorcheck = errorcheck & "Invalid Tour Date -- "
     end if

     ' ****************************
     ' Check Rank Placement Value
     ' ****************************
     
     ' First check if a score exists for the event that we are checking
     ' Also check if the value of the placement field is numerical.
     ' If there is a score (ie: not = "   ") and the value is NOT nemerical
     ' Then we have a problem.
     if (rsCSV.fields(29) <> "") and (not isnumeric(rsCSV.fields(10))) then
       ' Next we check if the placement field is blank. 
       ' Except for the "A" and "B" Tour's, the placement field is never
       ' used.  So we don't mind if it is blank excpet for "A" and "B" tours.
       if (rsCSV.fields(10) = "") and (ucase(right(rsCSV.fields(1).name,1)) = "A" or ucase(right(rsCSV.fields(1).name,1)) = "B") then
         errorcheck = errorcheck & "Slalom Placement is not a number -- "
       end if
     end if
     if (rsCSV.fields(33) <> "") and (not isnumeric(rsCSV.fields(13))) then 
       if (rsCSV.fields(13) = "") and (ucase(right(rsCSV.fields(1).name,1)) = "A" or ucase(right(rsCSV.fields(1).name,1)) = "B") then
         errorcheck = errorcheck & "Trick Placement is not a number -- "
       end if
     end if
     if (rsCSV.fields(40) <> "") and (not isnumeric(rsCSV.fields(16))) then
       if (rsCSV.fields(16) = "") and (ucase(right(rsCSV.fields(1).name,1)) = "A" or ucase(right(rsCSV.fields(1).name,1)) = "B") then
         errorcheck = errorcheck & "Jump Placement is not a number -- "
       end if
     end if

     If rsCSV.fields(9) > 0 AND rsCSV.fields(9) < 9 Then
       For RoundNum = 1 to rsCSV.fields(9)
        
             ' ****************************
             ' Set SkiYear Value
             ' ****************************
         '      Set rs=Server.CreateObject("ADODB.recordset")
         '      sSQL = "Select top 1 * from " & SkiYearTableName & " where BeginDate < '" & cdate(right(left(rsCSV.fields(12),6),2) &"/"& right(rsCSV.fields(12),2) &"/"& left(rsCSV.fields(12),4)) & "' and EndDate > '" & cdate(right(left(rsCSV.fields(12),6),2) &"/"& right(rsCSV.fields(12),2) &"/"& left(rsCSV.fields(12),4)) & "' and SkiYearID <> 1"
         '      rs.open sSQL, sConnectionToMemberTable
         '        If rs.EOF Then
         '          errorcheck = errorcheck & "Tour Date is not within a defined Ski Year -- "
         '        Else
         '          SkiYearID = rs("SkiYearID")
         '        End If
         '      rs.Close
        
             ' ****************************
             '  Check Slalom Scores (if they exist)
             ' ****************************       
               If rsCSV.fields(RoundNum * 22 + 7) <> "" Then
                 If NOT isalpha(rsCSV.fields(RoundNum * 22 + 1)) Then 
                   ErrorCheck = ErrorCheck & "Slalom " & RoundNum & " class is not a valid character. -- "
                 End If
         
                 ' Check Slalom Division Based on Age and Gender
                 PDF_Div = ucase(rsCSV.fields(RoundNum * 22 + 2))
        
                 If instr(ValidDivs,PDF_Div) = 0 Then 
                    If left(tempGender,1) = "M" And (instr(PDF_Div,"M") = 0 and instr(PDF_Div,"B") = 0) Or left(tempGender,1) = "F" And (instr(PDF_Div,"W") = 0 and instr(PDF_Div,"G") = 0)  Then
                       errorcheck = "*" & errorcheck & "Slalom "&RoundNum&" Div doesn't match gender. ("&PDF_Div&") -- "
                    Else
                       errorcheck = "*" & errorcheck & "Slalom " & RoundNum & " Div doesn't match age. ("&AgeInYears&") -- "
                    End If
                 End If
                 
                 If NOT isnumeric(rsCSV.fields(RoundNum * 22 + 6)) Then
                   ErrorCheck = ErrorCheck & "Slalom " & RoundNum & " EndPass Line is not a number (" & rsCSV.fields(RoundNum * 22 + 6) & ") -- "
                 End If
                 If NOT isnumeric(rsCSV.fields(RoundNum * 22 + 5)) Then
                   ErrorCheck = ErrorCheck & "Slalom " & RoundNum & " EndPass Speed is not a number (" & rsCSV.fields(RoundNum * 22 + 5) & ")-- "
                 End If
                 If NOT isnumeric(rsCSV.fields(RoundNum * 22 + 4)) Then
                   ErrorCheck = ErrorCheck & "Slalom " & RoundNum & " EndPass Score is not a number (" & rsCSV.fields(RoundNum * 22 + 4) & ") -- "
                 End If
                 If NOT isnumeric(rsCSV.fields(RoundNum * 22 + 7)) Then
                   ErrorCheck = ErrorCheck & "Slalom " & RoundNum & " Score is not a number (" & rsCSV.fields(RoundNum * 22 + 7) & ") -- "
                 End If
               End If
        
        
             ' ****************************
             '  Check Trick Scores (if they exist)
             ' ****************************     
               If rsCSV.fields(RoundNum * 22 + 11) <> "" Then
                 If NOT isalpha(rsCSV.fields(RoundNum * 22 + 8)) Then 
                   ErrorCheck = ErrorCheck & "Trick " & RoundNum & " class is not a valid character. -- "
                 End If
        
                 ' Check Trick Division Based on Age and Gender
                 PDF_Div = ucase(rsCSV.fields(RoundNum * 22 + 9))
        
                 If instr(ValidDivs,PDF_Div) = 0 Then 
                    If left(tempGender,1) = "M" And (instr(PDF_Div,"M") = 0 and instr(PDF_Div,"B") = 0) Or left(tempGender,1) = "F" And (instr(PDF_Div,"W") = 0 and instr(PDF_Div,"G") = 0)  Then
                       ErrorCheck = "*" & ErrorCheck & "Trick "&RoundNum&" Div doesn't match gender. ("&PDF_Div&") -- "
                    Else
                       ErrorCheck = "*" & ErrorCheck & "Trick " & RoundNum & " Div doesn't match age. ("&AgeInYears&") -- "
                    End If
                 End If
        
                 If NOT isnumeric(rsCSV.fields(RoundNum * 22 + 11)) Then
                   ErrorCheck = ErrorCheck & "Trick " & RoundNum & " Score is not a number (" & rsCSV.fields(RoundNum * 22 + 11) & ") -- "
                 End If
               End If
        
             ' ****************************
             '  Check Jump Scores (if they exist)
             ' ****************************     
               If rsCSV.fields(RoundNum * 22 + 18) <> "" Then
                 If NOT isalpha(rsCSV.fields(RoundNum * 22 + 12)) Then 
                   ErrorCheck = ErrorCheck & "Jump " & RoundNum & " class is not a valid character. -- "
                 End If
        
                 ' Check Jump Division Based on Age and Gender
                 PDF_Div = ucase(rsCSV.fields(RoundNum * 22 + 13))
        
                 If instr(ValidDivs,PDF_Div) = 0 Then 
                    If left(tempGender,1) = "M" And (instr(PDF_Div,"M") = 0 and instr(PDF_Div,"B") = 0) Or left(tempGender,1) = "F" And (instr(PDF_Div,"W") = 0 and instr(PDF_Div,"G") = 0)  Then
                       ErrorCheck = "*" & ErrorCheck & "Jump "&RoundNum&" Div doesn't match gender. ("&PDF_Div&") -- "
                    Else
                       ErrorCheck = "*" & ErrorCheck & "Jump " & RoundNum & " Div doesn't match age. ("&AgeInYears&") -- "
                    End If
                 End If
        
                 If NOT isnumeric(rsCSV.fields(RoundNum * 22 + 15)) Then
                   ErrorCheck = ErrorCheck & "Jump " & RoundNum & " Ramp Height is not a number (" & rsCSV.fields(RoundNum * 22 + 15) & ") -- "
                 End If
                 If NOT isnumeric(rsCSV.fields(RoundNum * 22 + 16)) Then
                   ErrorCheck = ErrorCheck & "Jump " & RoundNum & " Jump Speed is not a number (" & rsCSV.fields(RoundNum * 22 + 16) & ") -- "
                 End If
                 If NOT isnumeric(rsCSV.fields(RoundNum * 22 + 18)) Then
                   ErrorCheck = ErrorCheck & "Jump " & RoundNum & " Distance (Meters) is not a number (" & rsCSV.fields(RoundNum * 22 + 18) & ") -- "
                 End If
                 If NOT isnumeric(rsCSV.fields(RoundNum * 22 + 17)) Then
                   ErrorCheck = ErrorCheck & "Jump " & RoundNum & " Distance (Feet) is not a number (" & rsCSV.fields(RoundNum * 22 + 17) & ") -- "
                 End If
               End If
        

       Next
     End If

     If ErrorCheck = "" Then
       If GoodRec = 0 Then
         If NOT (objFSO.FileExists(fileoutgood)) = TRUE Then
           Set objstreamgood=objFSO.CreateTextFile(fileoutgood)
           objstreamgood.WriteLine (LineOne)
         Else
           Set objstreamgood=objFSO.opentextfile(fileoutgood,8,TRUE)
         End If
       End If
       objstreamgood.writeline (linetext)

       If rsCSV.fields(9) > 0 AND rsCSV.fields(9) < 9 Then
         For RoundNum = 1 to rsCSV.fields(9)

         ' ****************************
         '  Add in Slalom Scores (if they exist)
         ' ****************************
            If rsCSV.fields(RoundNum * 22 + 7) <> "" Then
              ' ****************************
              '  Before adding any score, check for duplicates
              ' ****************************
                 sSQL = "Select * from " & RawScoresTableName
                 sSQL = sSQL + " where MemberID = '" & SQLClean(rsCSV.fields(1)) & "' and"
                 sSQL = sSQL + " TourID = '" & SQLClean(rsCSV.fields(1).name) & "' and"
                 sSQL = sSQL + " Event = 'S' and"
                 sSQL = sSQL + " [Round] = '" & RoundNum & "' and"
                 sSQL = sSQL + " Div = '" & SQLClean(rsCSV.fields(RoundNum * 22 + 2)) & "'"
                 rs.open sSQL, sConnectionToTRATable, 3,3
    
                 ' If there are no duplicate records and there have been no errors
                 If rs.eof and (Err.Number = 0) Then 
                 
                    rs.close
                    sSQL = "select zbsconversion from " & DivisionsTableName & " join " & SkiYearTableName & " on " & DivisionsTableName & ".skiyearid = " & SkiYearTableName & ".skiyearid where upper(div) = '" & SQLClean(ucase(rsCSV.fields(RoundNum * 22 + 2))) & "' and '" & cdate(right(left(rsCSV.fields(9).name,6),2) &"/"& right(rsCSV.fields(9).name,2) &"/"& left(rsCSV.fields(9).name,4)) & "' between begindate and enddate AND USEZBS IS NULL"
                    rs.open sSQL, sConnectionToTRATable, 3,3

                   If rs.eof Then
                     ZBSFactor = 0
                   Else
                     ZBSFactor = rs("ZBSConversion")
                   End If
                   sSQL = "insert into " & RawScoresTableName
                   sSQL = sSQL + " (MemberFED, MemberID, LName, FName, TourFED, TourID, [H-Class], EndDate, Event, Place, [Round], Class, Div, Perf_Qual1, Perf_Qual2, AltScore, Score, Team, PreZBSConvScore)"
                   sSQL = sSQL + " VALUES ("
                   sSQL = sSQL + "'" & tempFed & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(1)) & "',"
                   sSQL = sSQL + "'" & tempLName & "',"
                   sSQL = sSQL + "'" & tempFName & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(0).name) & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(1).name) & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(4).name) & "',"
                   sSQL = sSQL + "'" & cdate(right(left(rsCSV.fields(9).name,6),2) &"/"& right(rsCSV.fields(9).name,2) &"/"& left(rsCSV.fields(9).name,4)) & "',"
                   sSQL = sSQL + "'S',"
                   sSQL = sSQL + "'" & SQLClean(trim(rsCSV.fields(10))) & "',"
                   sSQL = sSQL + "'" & RoundNum & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(RoundNum * 22 + 1)) & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(RoundNum * 22 + 2)) & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(RoundNum * 22 + 6)) & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(RoundNum * 22 + 5)) & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(RoundNum * 22 + 4)) & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(RoundNum * 22 + 7)) + ZBSFactor & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(8)) & "',"
                   If ZBSFactor = 0 Then
                     sSQL = sSQL + "null)"
                   Else
                     sSQL = sSQL + "'" & SQLClean(rsCSV.fields(RoundNum * 22 + 7)) & "')"                   
                   End If
         '         sSQL = sSQL + "'" & SkiYearID & "')"
         '         When you enable this, don't forget to add it to the field list above.
                   Con.Execute(sSQL)
               End If
               rs.close
             End If
         ' ****************************
         '  Add in Trick Scores (if they exist)
         ' ****************************
            If rsCSV.fields(RoundNum * 22 + 11) <> "" Then
              ' ****************************
              '  Before adding any score, check for duplicates
              ' ****************************
                 sSQL = "Select * from " & RawScoresTableName
                 sSQL = sSQL + " where MemberID = '" & SQLClean(rsCSV.fields(1)) & "' and"
                 sSQL = sSQL + " TourID = '" & SQLClean(rsCSV.fields(1).name) & "' and"
                 sSQL = sSQL + " Event = 'T' and"
                 sSQL = sSQL + " [Round] = '" & RoundNum & "' and"
                 sSQL = sSQL + " Div = '" & SQLClean(rsCSV.fields(RoundNum * 22 + 9)) & "'"
                 rs.open sSQL, sConnectionToTRATable, 3,3
    
                 ' If there are no duplicate records and there have been no errors
                 If rs.eof and (Err.Number = 0) Then 
                   sSQL = "insert into " & RawScoresTableName
                   sSQL = sSQL + " (MemberFED, MemberID, LName, FName, TourFED, TourID, [H-Class], EndDate, Event, Place, [Round], Class, Div, Team, Score)"
                   sSQL = sSQL + " VALUES ("
                   sSQL = sSQL + "'" & tempFed & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(1)) & "',"
                   sSQL = sSQL + "'" & tempLName & "',"
                   sSQL = sSQL + "'" & tempFName & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(0).name) & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(1).name) & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(4).name) & "',"
                   sSQL = sSQL + "'" & cdate(right(left(rsCSV.fields(9).name,6),2) &"/"& right(rsCSV.fields(9).name,2) &"/"& left(rsCSV.fields(9).name,4)) & "',"
                   sSQL = sSQL + "'T',"
                   sSQL = sSQL + "'" & SQLClean(trim(rsCSV.fields(13))) & "',"
                   sSQL = sSQL + "'" & RoundNum & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(RoundNum * 22 + 8)) & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(RoundNum * 22 + 9)) & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(8)) & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(RoundNum * 22 + 11)) & "')"
         '         sSQL = sSQL + "'" & SkiYearID & "')"
         '         When you enable this, don't forget to add it to the field list above.
                   Con.Execute(sSQL)
                 End If
                 rs.close
            End If
         ' ****************************
         '  Add in Jump Scores (if they exist)
         ' ****************************     
           If rsCSV.fields(RoundNum * 22 + 18) <> "" Then
              ' ****************************
              '  Before adding any score, check for duplicates
              ' ****************************
                 sSQL = "Select * from " & RawScoresTableName
                 sSQL = sSQL + " where MemberID = '" & SQLClean(rsCSV.fields(1)) & "' and"
                 sSQL = sSQL + " TourID = '" & SQLClean(rsCSV.fields(1).name) & "' and"
                 sSQL = sSQL + " Event = 'J' and"
                 sSQL = sSQL + " [Round] = '" & RoundNum & "' and"
                 sSQL = sSQL + " Div = '" & SQLClean(rsCSV.fields(RoundNum * 22 + 13)) & "'"
                 rs.open sSQL, sConnectionToTRATable, 3,3
    
                 ' If there are no duplicate records and there have been no errors
                 If rs.eof and (Err.Number = 0) Then 
                   sSQL = "insert into " & RawScoresTableName
                   sSQL = sSQL + " (MemberFED, MemberID, LName, FName, TourFED, TourID, [H-Class], EndDate, Event, Place, [Round], Class, Div, Perf_Qual1, Perf_Qual2, AltScore, Team, Score)"
                   sSQL = sSQL + " VALUES ("
                   sSQL = sSQL + "'" & tempFed & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(1)) & "',"
                   sSQL = sSQL + "'" & tempLName & "',"
                   sSQL = sSQL + "'" & tempFName & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(0).name) & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(1).name) & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(4).name) & "',"
                   sSQL = sSQL + "'" & cdate(right(left(rsCSV.fields(9).name,6),2) &"/"& right(rsCSV.fields(9).name,2) &"/"& left(rsCSV.fields(9).name,4)) & "',"
                   sSQL = sSQL + "'J',"
                   sSQL = sSQL + "'" & SQLClean(trim(rsCSV.fields(16))) & "',"
                   sSQL = sSQL + "'" & RoundNum & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(RoundNum * 22 + 12)) & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(RoundNum * 22 + 13)) & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(RoundNum * 22 + 15)) & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(RoundNum * 22 + 16)) & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(RoundNum * 22 + 18)) & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(8)) & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(RoundNum * 22 + 17)) & "')"
         '         sSQL = sSQL + "'" & SkiYearID & "')"
         '         When you enable this, don't forget to add it to the field list above.
                   Con.Execute(sSQL)
                 End If
                 rs.close
             End If

         Next
       End If
       ' We have saved our good record ... so now we read the next line (assuming there is one)
       GoodRec = GoodRec + 1
       If not objstreamin.atendofStream Then lineText=objstreamin.readline
     Else
       If BadRec = 0 Then
         If NOT (objFSO.FileExists(fileoutbad)) = TRUE Then
           Set objstreambad=objFSO.createtextfile(fileoutbad)
           objstreambad.WriteLine (LineOne)
         Else
           Set objstreambad=objFSO.opentextfile(fileoutbad,8,TRUE)
         End If
         If NOT (objFSO.FileExists(fileoutexplainations)) = TRUE Then
           Set objstreamexplainations=objFSO.createtextfile(fileoutexplainations)
           objstreamexplainations.WriteLine ("FILE HEADER -- NO DATA HERE")
         Else
           Set objstreamexplainations=objFSO.opentextfile(fileoutexplainations,8,TRUE)
         End If
       End If
       objstreambad.writeline (linetext)
       If len(ErrorCheck) > 200 Then 
         If left(ErrorCheck,1) = "*" Then
           objstreamexplainations.WriteLine ("*Multiple Errors Found.")
         Else
           objstreamexplainations.WriteLine ("Multiple Errors Found.")
         End If
       Else
        objstreamexplainations.writeline (ErrorCheck)
       End If
       ' We have saved our bad record ... so now we read the next line (assuming there is one)
       BadRec = BadRec + 1
       If not objstreamin.atendofStream Then lineText=objstreamin.readline
     End If



	rsCSV.MoveNext   'Move to the next record
Loop


'Close our recordset and connection
rsCSV.close
Set rsCSV = Nothing
conTest.close
set conTest = nothing

objfso.DeleteFile CSVfile


End Sub














Sub ProcessPDF

  set objstreamin=objFSO.opentextfile(filein)

  do while not objstreamin.atendofStream
     lineText=objstreamin.readline
     ErrorCheck = ""
     ValidDivs = ""

     ' ****************************
     ' Check MemberID exists in Master DB
     ' ****************************

     sSQL = "Select top 1 * from "&MemberTableName&" where PersonIDwithCheckDigit = '" & right(left(linetext,12),9) & "'"
     rs.open sSQL, sConnectionToTRATable
     if rs.eof then
       tempFed = left(linetext,3)
       tempBirthdate = ""
       tempGender = ""
       errorcheck = "*" & errorcheck & "Member ID Not Found -- "
     Else
       tempFed = trim(rs("FederationCode"))
       tempBirthdate = rs("BirthDate")
       tempGender = ucase(left(rs("Sex"),1))
     end if
  
     ' ****************************
     ' Check MemberID matches Member Name
     ' ****************************

     if not rs.eof then
       if ucase(trim(right(left(linetext,31),17))) <> left(ucase(trim(rs("LastName"))),17) then
         tempLName = replace(trim(right(left(linetext,31),17)), "'", "''")
         errorcheck = "*" & errorcheck & "Member Last Name Incorrect -- "
       Else
         tempLName = replace(trim(rs("LastName")), "'", "''")
       End If
'       if ucase(trim(right(left(linetext,44),13))) <> ucase(trim(rs("FirstName"))) then
'         errorcheck = errorcheck & "Member First Name Incorrect -- "
'       end if
       If ucase(trim(right(left(linetext,32),1))) <> left(ucase(trim(rs("FirstName"))),1) Then
         tempFName = replace(trim(right(left(linetext,32),1)), "'", "''")
         errorcheck = "*" & errorcheck & "Member First Name Incorrect -- "
       Else
         tempFName = replace(trim(rs("FirstName")), "'", "''")
       End If
     end if
     rs.Close

' Build a list of valid divisions for this member.

'    First we figure out what ski year they are in
     sSQL = "Select top 1 * from " & SkiYearTableName & " where '"&CDate("01/01/"&(2000+(left(Request("file"),2))))&"' BETWEEN BeginDate and EndDate and SkiYearID <> 1"
     rs.open sSQL, SConnectionToTRATable, 3, 1
     If rs.EOF Then
        sSQL = "Select top 1 * from " & SkiYearTableName & " where DefaultYear = 1"
        rs.open sSQL, SConnectionToTRATable, 3, 1
     End If
     If rs.EOF Then
        tempSkiYear = "SkiYearID" ' We can't find any matching Ski Years so we just trick the SQL into ignoring this condition. (skiyearid = skiyearid)
     Else
        tempSkiYear = rs("SkiYearID")
     End If
     rs.Close ' Close the Ski Year Table
 
'    Next we figure out their age.
     ' NOTE: Division is based on age relative to ski year.
     If tempBirthDate <> "" Then
       ' get absolute number of years 
       AgeInYears = cint(datediff("YYYY", tempBirthDate, CDate("01/01/"&(2000+(right(left(linetext,65),2)))))) - 1
       If tempGender = "M" Then
         sSQL = "Select distinct div from " & DivisionsTableName & " where upper(sex) = 'M' and "&AgeInYears&" <= Up_Age and "&AgeInYears&" >= Low_Age and SkiYearID = "& tempSkiYear &" order by Div"
       Else
         sSQL = "Select distinct div from " & DivisionsTableName & " where upper(sex) = 'F' and "&AgeInYears&" <= Up_Age and "&AgeInYears&" >= Low_Age and SkiYearID = "& tempSkiYear &" order by Div"
       End If
     Else
       sSQL = "Select top 1 div from " & DivisionsTableName & " where 0=1"
     End If

    rs.open sSQL, SConnectionToTRATable
    If Not rs.EOF Then 
      DivArray = rs.GetRows()
      for i = 0 to ubound(DivArray,2)
        ValidDivs = ValidDivs + DivArray(0,i)
        if i < ubound(DivArray,2) then ValidDivs = ValidDivs + ","
      Next
    End If
    rs.Close

 

     ' ****************************
     ' Check TourID with Master DB
     ' ****************************

' The original idea was to verify each score against the headquarters SWIFT
' tables.

' We really don't have to verify each score because we have already proven
' that the file name is valid.

' Therefore all we have to do is prove that the score matches the file name.

'     Set rs=Server.CreateObject("ADODB.recordset")
'     sSQL = "Select top 1 * from TSchedul where TSanction = '" & right(left(linetext,6),3) & "'"
'     rs.open sSQL, sConnectionToSanctionTable
'     if rs.eof then
'       errorcheck = errorcheck & "Tour ID Not Found -- "
'     End If
'     rs.Close


     ' ****************************
     ' Check blanks between memberid and lastname
     ' ****************************
     if right(left(linetext,14),2) <> "  " then
       errorcheck = errorcheck & "No Blanks Between MemberID and Lastname (14) -- "
     end if 

     ' ****************************
     ' Check Tour ID
     ' ****************************
     ' Tour ID in the score record should match the file name.
     if ucase(right(left(linetext,69),6)) <> ucase(left(request("file"),6)) then
       errorcheck = errorcheck & "Field 'Tour ID' Doesn't Match Filename (70) -- "
     end if

     ' ****************************
     ' Check For Valid Date
     ' ****************************
     ' Date must be valid or crash will occur.
     if not isDate(right(left(linetext,78),2) &"/"& right(left(linetext,80),2) &"/"& right(left(linetext,76),4)) then
       errorcheck = errorcheck & "Invalid Tour Date (80) -- "
     end if

     ' ****************************
     ' Check Rank Placement Value
     ' ****************************
     
     ' First check if a score exists for the event that we are checking
     ' Also check if the value of the placement field is numerical.
     ' If there is a score (ie: not = "   ") and the value is NOT nemerical
     ' Then we have a problem.
     if (right(left(linetext,131),5) <> "     ") and (not isnumeric(trim(right(left(linetext,93),3)))) then
       ' Next we check if the placement field is blank. 
       ' Except for the "A" and "B" Tour's, the placement field is never
       ' used.  So we don't mind if it is blank excpet for "A" and "B" tours.
       if (trim(right(left(linetext,93),3)) = "") and (ucase(right(left(linetext,70),1)) = "A" or ucase(right(left(linetext,70),1)) = "B") then
         errorcheck = errorcheck & "Slalom Placement is not a number (93) -- "
       end if
     end if
     if (right(left(linetext,145),5) <> "     ") and (not isnumeric(trim(right(left(linetext,97),3)))) then 
       if (trim(right(left(linetext,97),3)) = "") and (ucase(right(left(linetext,70),1)) = "A" or ucase(right(left(linetext,70),1)) = "B") then
         errorcheck = errorcheck & "Trick Placement is not a number (97) -- "
       end if
     end if
     if (right(left(linetext,166),4) <> "    ") and (not isnumeric(trim(right(left(linetext,101),3)))) then
       if (trim(right(left(linetext,101),3)) = "") and (ucase(right(left(linetext,70),1)) = "A" or ucase(right(left(linetext,70),1)) = "B") then
         errorcheck = errorcheck & "Jump Placement is not a number (101) -- "
       end if
     end if


     RoundNum = 1
     Do While len(trim(linetext)) > (RoundNum * 60 + 52)

     ' ****************************
     ' Set SkiYear Value
     ' ****************************
 '      Set rs=Server.CreateObject("ADODB.recordset")
 '      sSQL = "Select top 1 * from " & SkiYearTableName & " where BeginDate < '" & cdate(right(left(linetext,78),2) &"/"& right(left(linetext,80),2) &"/"& right(left(linetext,76),4)) & "' and EndDate > '" & cdate(right(left(linetext,78),2) &"/"& right(left(linetext,80),2) &"/"& right(left(linetext,76),4)) & "' and SkiYearID <> 1"
 '      rs.open sSQL, sConnectionToMemberTable
 '        If rs.EOF Then
 '          errorcheck = errorcheck & "Tour Date is not within a defined Ski Year -- "
 '        Else
 '          SkiYearID = rs("SkiYearID")
 '        End If
 '      rs.Close


     ' ****************************
     ' Check start of round
     ' ****************************
       if (right(left(linetext,(RoundNum * 60 + 51)),1) <> cstr(RoundNum) and trim(right(left(linetext,(RoundNum * 60 + 107)),40)) <> "") then
         errorcheck = errorcheck & "Round " & RoundNum & " scores present but no Round " & RoundNum & " Marker (" & RoundNum * 60 + 51 & ")  -- "
       end if

     ' ****************************
     '  Check Slalom Scores (if they exist)
     ' ****************************       
       if right(left(linetext,(RoundNum * 60 + 71)),5) <> "     " then
         if not isalpha( right(left(linetext,(RoundNum * 60 + 52)),1) ) then 
           errorcheck = errorcheck & "Slalom " & RoundNum & " class is not a valid character. (" & RoundNum * 60 + 52 & ") -- "
         end if
 
         ' Check Slalom Division Based on Age and Gender
           PDF_Div = ucase(right(left(linetext,(RoundNum * 60 + 54)),2))

           if instr(ValidDivs,PDF_Div) = 0 then 
              If left(tempGender,1) = "M" And (instr(PDF_Div,"M") = 0 and instr(PDF_Div,"B") = 0) Or left(tempGender,1) = "F" And (instr(PDF_Div,"W") = 0 and instr(PDF_Div,"G") = 0)  Then
                 errorcheck = errorcheck & "Slalom "&RoundNum&" Div doesn't match gender. -- "
              Else
                 errorcheck = errorcheck & "Slalom " & RoundNum & " Div doesn't match age "&AgeInYears&". (" & RoundNum * 60 + 54 & ") -- "
              End If
           end if
         
         if not isnumeric(right(left(linetext,(RoundNum * 60 + 66)),4)) then
           errorcheck = errorcheck & "Slalom " & RoundNum & " EndPass Line is not a number (" & RoundNum * 60 + 66 & ") -- "
         end if
         if not isnumeric(right(left(linetext,(RoundNum * 60 + 62)),2)) then
           errorcheck = errorcheck & "Slalom " & RoundNum & " EndPass Speed is not a number (" & RoundNum * 60 + 62 & ") -- "
         end if
         if not isnumeric(right(left(linetext,(RoundNum * 60 + 60)),4)) then
           errorcheck = errorcheck & "Slalom " & RoundNum & " EndPass Score is not a number (" & RoundNum * 60 + 60 & ") -- "
         end if
         if not isnumeric(right(left(linetext,(RoundNum * 60 + 71)),5)) then
           errorcheck = errorcheck & "Slalom " & RoundNum & " Score is not a number (" & RoundNum * 60 + 71 & ") -- "
         end if
       end if

     ' ****************************
     ' Check blanks between slalom1 and trick1
     ' ****************************
       if right(left(linetext,(RoundNum * 60 + 75)),4) <> "    " then
         errorcheck = errorcheck & "Blanks missing between slalom " & RoundNum & " and trick " & RoundNum & " scores (" & RoundNum * 60 + 75 & ") -- "
       end if


     ' ****************************
     '  Check Trick Scores (if they exist)
     ' ****************************     
       if right(left(linetext,(RoundNum * 60 + 85)),5) <> "     " then
         if not isalpha(right(left(linetext,(RoundNum * 60 + 76)),1)) then 
           errorcheck = errorcheck & "Trick " & RoundNum & " class is not a valid character. (" & RoundNum * 60 + 76 & ") -- "
         end if

         ' Check Trick Division Based on Age and Gender
           PDF_Div = ucase(right(left(linetext,(RoundNum * 60 + 78)),2))

           if instr(ValidDivs,PDF_Div) = 0 then 
              If left(tempGender,1) = "M" And (instr(PDF_Div,"M") = 0 and instr(PDF_Div,"B") = 0) Or left(tempGender,1) = "F" And (instr(PDF_Div,"W") = 0 and instr(PDF_Div,"G") = 0)  Then
                 errorcheck = errorcheck & "Trick "&RoundNum&" Div doesn't match gender. -- "
              Else
                 errorcheck = errorcheck & "Trick " & RoundNum & " Div doesn't match age "&AgeInYears&". (" & RoundNum * 60 + 78 & ") -- "
              End If
           end if

         if not isnumeric(right(left(linetext,(RoundNum * 60 + 85)),5)) then
           errorcheck = errorcheck & "Trick " & RoundNum & " Score is not a number (" & RoundNum * 60 + 85 & ") -- "
         end if
       end if

     ' ****************************
     ' Check blanks between trick and jump
     ' ****************************     
       if right(left(linetext,(RoundNum * 60 + 88)),3) <> "   " then
         errorcheck = errorcheck & "Blanks missing between trick " & RoundNum & " and jump " & RoundNum & " scores (" & RoundNum * 60 + 88 & ") -- "
       end if

     ' ****************************
     '  Check Jump Scores (if they exist)
     ' ****************************     
       if right(left(linetext,(RoundNum * 60 + 106)),4) <> "    " then
         if not isalpha(right(left(linetext,(RoundNum * 60 + 89)),1)) then 
           errorcheck = errorcheck & "Jump " & RoundNum & " class is not a valid character. (" & RoundNum * 60 + 89 & ") -- "
         end if

         ' Check Jump Division Based on Age and Gender
           PDF_Div = ucase(right(left(linetext,(RoundNum * 60 + 91)),2))

           if instr(ValidDivs,PDF_Div) = 0 then 
              If left(tempGender,1) = "M" And (instr(PDF_Div,"M") = 0 and instr(PDF_Div,"B") = 0) Or left(tempGender,1) = "F" And (instr(PDF_Div,"W") = 0 and instr(PDF_Div,"G") = 0)  Then
                 errorcheck = errorcheck & "Jump "&RoundNum&" Div doesn't match gender. -- "
              Else
                 errorcheck = errorcheck & "Jump " & RoundNum & " Div doesn't match age "&AgeInYears&". (" & RoundNum * 60 + 91 & ") -- "
              End If
           end if

         if not isnumeric(right(left(linetext,(RoundNum * 60 + 97)),4)) then
           errorcheck = errorcheck & "Jump " & RoundNum & " Ramp Height is not a number (" & RoundNum * 60 + 97 & ") -- "
         end if
         if not isnumeric(right(left(linetext,(RoundNum * 60 + 99)),2)) then
           errorcheck = errorcheck & "Jump " & RoundNum & " Jump Speed is not a number (" & RoundNum * 60 + 99 & ") -- "
         end if
         if not isnumeric(right(left(linetext,(RoundNum * 60 + 106)),4)) then
           If trim(right(left(linetext,(RoundNum * 60 + 106)),4)) <> ".F" and trim(right(left(linetext,(RoundNum * 60 + 106)),4)) <> ".P" Then
             errorcheck = errorcheck & "Jump " & RoundNum & " Distance (Meters) is not a number (" & RoundNum * 60 + 106 & ") -- "
           End If
         End If
         if not isnumeric(right(left(linetext,(RoundNum * 60 + 102)),3)) then
           If trim(right(left(linetext,(RoundNum * 60 + 102)),3)) <> "F" And trim(right(left(linetext,(RoundNum * 60 + 102)),3)) <> "P" Then
             errorcheck = errorcheck & "Jump " & RoundNum & " Distance (Feet) is not a number (" & RoundNum * 60 + 102 & ") -- "
           End If
         end if
         if ((right(left(linetext,(RoundNum * 60 + 94)),1) <> "." or (right(left(linetext,(RoundNum * 60 + 105)),1) <> "." and right(left(datain,(RoundNum * 60 + 104)),1) <> ".")) and right(left(datain,(RoundNum * 60 + 99)),1) <> " ") then
           If trim(right(left(linetext,(RoundNum * 60 + 106)),4)) <> ".F" and trim(right(left(linetext,(RoundNum * 60 + 106)),4)) <> ".P" Then
             errorcheck = errorcheck & "Decimal not found for jump score (" & RoundNum * 60 + 94 & "/" & RoundNum * 60 + 105 & ") -- "
           End If 
         end if
       end if

     ' ****************************
     ' Check blanks between jump and next round
     ' ****************************     
       if right(left(linetext,RoundNum * 60 + 110),4) <> "    " then
         errorcheck = errorcheck & "Blanks missing between round " & RoundNum & " and round " & RoundNum + 1 & " (" & RoundNum * 60 + 110 & ") -- "
       end if

       RoundNum = RoundNum + 1
     Loop                                 




     if errorcheck = "" then
       if goodrec = 0 then
         if Not (objFSO.FileExists(fileoutgood)) = true then
           set objstreamgood=objFSO.createtextfile(fileoutgood)
         else
           set objstreamgood=objFSO.opentextfile(fileoutgood,8,true)
         end if
       end if
       objstreamgood.writeline (linetext)

       RoundNum = 1
       Do While len(trim(linetext)) > (RoundNum * 60 + 52) 

     ' ****************************
     '  Add in Slalom Scores (if they exist)
     ' ****************************
         if right(left(linetext,(RoundNum * 60 + 71)),5) <> "     " then
          ' ****************************
          '  Before adding any score, check for duplicates
          ' ****************************
             sSQL = "Select * from " & RawScoresTableName
             sSQL = sSQL + " where MemberID = '" & right(left(linetext,12),9) & "' and"
             sSQL = sSQL + " TourID = '" & right(left(linetext,71),8) & "' and"
             sSQL = sSQL + " Event = 'S' and"
             sSQL = sSQL + " [Round] = '" & RoundNum & "' and"
             sSQL = sSQL + " Div = '" & right(left(linetext,(RoundNum * 60 + 54)),2) & "'"
             rs.open sSQL, sConnectionToTRATable, 3,3

           ' If there are no duplicate records and there have been no errors
           If rs.eof and (Err.Number = 0) Then 
             
             rs.close
             sSQL = "select zbsconversion from " & DivisionsTableName & " join " & SkiYearTableName & " on " & DivisionsTableName & ".skiyearid = " & SkiYearTableName & ".skiyearid where upper(div) = '" & ucase(right(left(linetext,(RoundNum * 60 + 54)),2)) & "' and '" & cdate(right(left(linetext,78),2) &"/"& right(left(linetext,80),2) &"/"& right(left(linetext,76),4)) & "' between begindate and enddate"
             rs.open sSQL, sConnectionToTRATable, 3,3

             If rs.eof Then
               ZBSFactor = 0
               WriteLog ("ERROR - ZBS could not be adjusted with Division '" & ucase(right(left(linetext,(RoundNum * 60 + 54)),2)) & "' and '" & cdate(right(left(linetext,78),2) &"/"& right(left(linetext,80),2) &"/"& right(left(linetext,76),4)) & "'.")
             Else
               ZBSFactor = rs("ZBSConversion")
             End If

             sSQL = "insert into " & RawScoresTableName
             sSQL = sSQL + " (MemberFED, MemberID, LName, FName, TourFED, TourID, [H-Class], EndDate, Event, Place, [Round], Class, Div, Perf_Qual1, Perf_Qual2, AltScore, Score, Team, PreZBSConvScore)"
             sSQL = sSQL + " VALUES ("
             sSQL = sSQL + "'" & tempFed & "',"
             sSQL = sSQL + "'" & right(left(linetext,12),9) & "',"
             sSQL = sSQL + "'" & tempLName & "',"
             sSQL = sSQL + "'" & tempFName & "',"
             sSQL = sSQL + "'" & right(left(linetext,63),3) & "',"
             sSQL = sSQL + "'" & right(left(linetext,71),8) & "',"
             sSQL = sSQL + "'" & right(left(linetext,72),1) & "',"
             sSQL = sSQL + "'" & cdate(right(left(linetext,78),2) &"/"& right(left(linetext,80),2) &"/"& right(left(linetext,76),4)) & "',"
             sSQL = sSQL + "'S',"
             if isnumeric(trim(right(left(linetext,93),3))) then
               sSQL = sSQL + "'" & right(left(linetext,93),3) & "',"
             else
               sSQL = sSQL + "NULL,"
             end if
             sSQL = sSQL + "'" & RoundNum & "',"
             sSQL = sSQL + "'" & right(left(linetext,(RoundNum * 60 + 52)),1) & "',"
             sSQL = sSQL + "'" & right(left(linetext,(RoundNum * 60 + 54)),2) & "',"
             sSQL = sSQL + "'" & right(left(linetext,(RoundNum * 60 + 66)),4) & "',"
             sSQL = sSQL + "'" & right(left(linetext,(RoundNum * 60 + 62)),2) & "',"
             sSQL = sSQL + "'" & right(left(linetext,(RoundNum * 60 + 60)),4) & "',"
             sSQL = sSQL + "'" & (right(left(linetext,(RoundNum * 60 + 71)),5) + ZBSFactor) & "',"
             sSQL = sSQL + "'" & SQLClean(rsCSV.fields(8)) & "',"
             sSQL = sSQL + "'" & right(left(linetext,(RoundNum * 60 + 71)),5) & "')"
   '         sSQL = sSQL + "'" & SkiYearID & "')"
   '         When you enable this, don't forget to add it to the field list above.
             Con.Execute(sSQL)
           End If
           rs.close
         end if
     ' ****************************
     '  Add in Trick Scores (if they exist)
     ' ****************************
         if right(left(linetext,(RoundNum * 60 + 85)),5) <> "     " then
          ' ****************************
          '  Before adding any score, check for duplicates
          ' ****************************
             sSQL = "Select * from " & RawScoresTableName
             sSQL = sSQL + " where MemberID = '" & right(left(linetext,12),9) & "' and"
             sSQL = sSQL + " TourID = '" & right(left(linetext,71),8) & "' and"
             sSQL = sSQL + " Event = 'T' and"
             sSQL = sSQL + " [Round] = '" & RoundNum & "' and"
             sSQL = sSQL + " Div = '" & right(left(linetext,(RoundNum * 60 + 78)),2) & "'"
             rs.open sSQL, sConnectionToTRATable, 3,3

           ' If there are no duplicate records and there have been no errors
           If rs.eof and (Err.Number = 0) Then 
             sSQL = "insert into " & RawScoresTableName
             sSQL = sSQL + " (MemberFED, MemberID, LName, FName, TourFED, TourID, [H-Class], EndDate, Event, Place, [Round], Class, Div, Team, Score)"
             sSQL = sSQL + " VALUES ("
             sSQL = sSQL + "'" & tempFed & "',"
             sSQL = sSQL + "'" & right(left(linetext,12),9) & "',"
             sSQL = sSQL + "'" & tempLName & "',"
             sSQL = sSQL + "'" & tempFName & "',"
             sSQL = sSQL + "'" & right(left(linetext,63),3) & "',"
             sSQL = sSQL + "'" & right(left(linetext,71),8) & "',"
             sSQL = sSQL + "'" & right(left(linetext,72),1) & "',"
             sSQL = sSQL + "'" & cdate(right(left(linetext,78),2) &"/"& right(left(linetext,80),2) &"/"& right(left(linetext,76),4)) & "',"
             sSQL = sSQL + "'T',"
             if isnumeric(trim(right(left(linetext,97),3))) then
               sSQL = sSQL + "'" & right(left(linetext,97),3) & "',"
             else
               sSQL = sSQL + "NULL,"
             end if
             sSQL = sSQL + "'" & RoundNum & "',"
             sSQL = sSQL + "'" & right(left(linetext,(RoundNum * 60 + 76)),1) & "',"
             sSQL = sSQL + "'" & right(left(linetext,(RoundNum * 60 + 78)),2) & "',"
             sSQL = sSQL + "'" & SQLClean(rsCSV.fields(8)) & "',"
             sSQL = sSQL + "'" & right(left(linetext,(RoundNum * 60 + 85)),5) & "')"
   '         sSQL = sSQL + "'" & SkiYearID & "')"
   '         When you enable this, don't forget to add it to the field list above.
             Con.Execute(sSQL)
           End If
           rs.close
         end if
     ' ****************************
     '  Add in Jump Scores (if they exist)
     ' ****************************     
         if right(left(linetext,(RoundNum * 60 + 106)),4) <> "    " then
          ' ****************************
          '  Before adding any score, check for duplicates
          ' ****************************
             sSQL = "Select * from " & RawScoresTableName
             sSQL = sSQL + " where MemberID = '" & right(left(linetext,12),9) & "' and"
             sSQL = sSQL + " TourID = '" & right(left(linetext,71),8) & "' and"
             sSQL = sSQL + " Event = 'J' and"
             sSQL = sSQL + " [Round] = '" & RoundNum & "' and"
             sSQL = sSQL + " Div = '" & right(left(linetext,(RoundNum * 60 + 91)),2) & "'"
             rs.open sSQL, sConnectionToTRATable, 3,3

           ' If there are no duplicate records and there have been no errors
           If rs.eof and (Err.Number = 0) Then 
             sSQL = "insert into " & RawScoresTableName
             sSQL = sSQL + " (MemberFED, MemberID, LName, FName, TourFED, TourID, [H-Class], EndDate, Event, Place, [Round], Class, Div, Perf_Qual1, Perf_Qual2, AltScore, Team, Score)"
             sSQL = sSQL + " VALUES ("
             sSQL = sSQL + "'" & tempFed & "',"
             sSQL = sSQL + "'" & right(left(linetext,12),9) & "',"
             sSQL = sSQL + "'" & tempLName & "',"
             sSQL = sSQL + "'" & tempFName & "',"
             sSQL = sSQL + "'" & right(left(linetext,63),3) & "',"
             sSQL = sSQL + "'" & right(left(linetext,71),8) & "',"
             sSQL = sSQL + "'" & right(left(linetext,72),1) & "',"
             sSQL = sSQL + "'" & cdate(right(left(linetext,78),2) &"/"& right(left(linetext,80),2) &"/"& right(left(linetext,76),4)) & "',"
             sSQL = sSQL + "'J',"
             if isnumeric(trim(right(left(linetext,101),3))) then
               sSQL = sSQL + "'" & right(left(linetext,101),3) & "',"
             else
               sSQL = sSQL + "NULL,"
             end if
             sSQL = sSQL + "'" & RoundNum & "',"
             sSQL = sSQL + "'" & right(left(linetext,(RoundNum * 60 + 89)),1) & "',"
             sSQL = sSQL + "'" & right(left(linetext,(RoundNum * 60 + 91)),2) & "',"
             sSQL = sSQL + "'" & right(left(linetext,(RoundNum * 60 + 97)),4) & "',"
             sSQL = sSQL + "'" & right(left(linetext,(RoundNum * 60 + 99)),2) & "',"
             If trim(right(left(linetext,(RoundNum * 60 + 106)),4)) = ".F" Or trim(right(left(linetext,(RoundNum * 60 + 106)),4)) = ".P" Then
               sSQL = sSQL + "'0.0',"
             Else
               sSQL = sSQL + "'" & right(left(linetext,(RoundNum * 60 + 106)),4) & "',"
             End If
             sSQL = sSQL + "'" & SQLClean(rsCSV.fields(8)) & "',"
             If trim(right(left(linetext,(RoundNum * 60 + 102)),3)) = "F" Or trim(right(left(linetext,(RoundNum * 60 + 102)),3)) = "P" Then
               sSQL = sSQL + "'0')"
             Else
               sSQL = sSQL + "'" & right(left(linetext,(RoundNum * 60 + 102)),3) & "')"
             End If
   '         sSQL = sSQL + "'" & SkiYearID & "')"
   '         When you enable this, don't forget to add it to the field list above.
             Con.Execute(sSQL)
           End If
           rs.close
         end if
         RoundNum = RoundNum + 1
       Loop          
       goodrec=goodrec+1
     else
       if badrec = 0 then
         if Not (objFSO.FileExists(fileoutbad)) = true then
           set objstreambad=objFSO.createtextfile(fileoutbad)
         else
           set objstreambad=objFSO.opentextfile(fileoutbad,8,true)
         end if
         if Not (objFSO.FileExists(fileoutexplainations)) = true then
           set objstreamexplainations=objFSO.createtextfile(fileoutexplainations)
         else
           set objstreamexplainations=objFSO.opentextfile(fileoutexplainations,8,true)
         end if
       end if
       objstreambad.writeline (linetext)
       If len(errorcheck) > 200 Then 
         If left(errorcheck,1) = "*" Then
           objstreamexplainations.WriteLine ("*Multiple Errors Found.")
         Else
           objstreamexplainations.WriteLine ("Multiple Errors Found.")
         End If
       else
        objstreamexplainations.writeline (errorcheck)
       end if
       badrec=badrec+1
     end if

  Loop
  CloseCon

   
End Sub








Sub BadSanctionCode

' This HTML code is used when the Tour ID could not be found in the
' SWIFT tables indicating that the Tour is probably not valid.

' The scores file gets saved to a temporary location but the individual
' scores do not get processed.

  WriteLog(date() &"  "& time() &"  "& filein & " TourID not found in SWIFT.  File saved to BADFILES folder.")

  Response.Flush
  %>

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
  </SCRIPT>  
<%WriteIndexPageHeader%>
    <br><br>
    <center><h3>The file <%=Request("file")%> has been uploaded.</h3><br><br><br>

    <h4><font color="red">The sanction code designated in the file name was NOT found in SWIFT.
    <br><br><br>
    The scores in this file have NOT been processed.</font>
    <br><br><br>
    </h4></center>
    <br><br>
    </body></html>
   <%
   WriteIndexPageFooter    

End Sub


Sub BadSanctionStatus

' This HTML code is used when the Tour ID Status in SWIFT is NOT
' 2 or 4 or 5 -- indicating that the Sanction Status is NOT COMPLETE.

' The scores file gets saved to a temporary location but the individual
' scores do not get processed.

  WriteLog(date() &"  "& time() &"  "& filein & " SWIFT Sanction Status is NOT COMPLETE.  File saved to BADFILES folder.")

  Response.Flush
  %>

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
  </SCRIPT>  
<%WriteIndexPageHeader%>
    <br><br>
    <center><h3>The file <%=Request("file")%> has been uploaded.</h3><br><br><br>

    <h4><font color="red">The Sanction Status for the Tournament ID indicated<br>in this file name is NOT COMPLETE at this time.
    <br><br>
    Please Contact your Regional EVP, as well as the Competition<br> Department at USA Waterski, with the particulars.
    <br><br>
    The scores in this file have NOT been processed.
    </font>
    <br><br><br>
    </h4></center>
    <br><br>
    </body></html>
   <%
   WriteIndexPageFooter    

End Sub

Sub BadFileExtension

' This HTML code is used when the file name indicates an invalid file type.

' The file gets saved to a temporary location, but the individual
' scores do not get processed.

  WriteLog(date() &"  "& time() &"  "& filein & " not a valid scores file type.  File saved to BADFILES folder.")

  Response.Flush
  %>

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
  </SCRIPT>  
<%WriteIndexPageHeader%>
    <br><br>
    <center><h3>The file <%=Request("file")%> has been uploaded.</h3><br><br><br>

    <h4><font color="red">The file prefix must be either PDF or WSP.  Other file types can not be processed by this system.
    <br><br><br>
    The scores in this file have NOT been processed.</font>
    <br><br><br>
    </h4></center>
    <br><br>
    </body></html>
   <%
   WriteIndexPageFooter    

End Sub


Sub BadFile

' This HTML code is used when there are other file problems.

' The scores file gets saved to a temporary location but the individual
' scores do not get processed.

  WriteLog(date() &"  "& time() &"  "& filein & " has failed file verification.  File saved to BADFILES folder.")

  Response.Flush
  %>

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
  </SCRIPT>  
<%WriteIndexPageHeader%>
    <br><br>
    <center><h3>The file <%=Request("file")%> has been uploaded.</h3><br><br><br>

    <h4><font color="red">The file can not be processed.  Please contact an administrator for help.
    <br><br><br>
    The scores in this file have NOT been processed.</font>
    <br><br><br>
    </h4></center>
    <br><br>
    </body></html>
   <%
   WriteIndexPageFooter    

End Sub

%>



