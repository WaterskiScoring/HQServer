<% IF Session("adminmenulevel")<10 THEN Response.Redirect "DefaultHQ.asp?process=login" %>

<!--#include file="settingsHQ.asp"-->

<%
If Request("WSPFile") = "" Then 
  WriteIndexPageHeader
  %>
  <center>
  <br><br>
  <h3><font color="red">No file specified for upload.</font></h3>
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
    <B>Importing Ranking File Scores Data.<br><br>
    Please stand by ...<br><br>  
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
    dim fileoutreasons
    dim objstreambad
    dim objstreamexplainations
    dim errorcheck
    dim goodrec
    Dim badrec
    Dim SkiYearID
    Dim tempFed, tempFName, tempLName, tempBirthdate, tempGender, tempSkiYear
    Dim tempTeam, tempTeamStat
    Dim PDF_Div, AgeInYears
    Dim ValidDivs, DivArray, i
    Dim ZBSFactor
    Dim sTourID
    
    Opencon
    Set rs=Server.CreateObject("ADODB.recordset")
    
    errorcheck = 0
    goodrec = 0
    badrec = 0
   
    ' Start by deleting any pre-existing scores for this tournament
    ' Note the count for later recap.

    sSQL = "SELECT count(*) as ScoreCount from (Select distinct MemberID from "
    sSQL = sSQL & RawScoresTableName & " WHERE upper(TourID) = '" 
    sSQL = sSQL & Ucase(Left(Request("WSPFile"),7)) & "') xx;"
    rs.open sSQL, sConnectionToTRATable, 3, 3
    IF rs.eof THEN OldScores = 0 ELSE OldScores = rs("ScoreCount")
    rs.Close
    IF OldScores > 0 THEN
       sSQL = "DELETE from " & RawScoresTableName & " WHERE upper(TourID) = '"
       sSQL = sSQL & Ucase(Left(Request("WSPFile"),7)) & "';"
       Con.Execute(sSQL)
    END IF


    filein=PathtoRawWSPs & "\" & request("WSPFile")  

    CSVfile=PathtoRawWSPs & "\" & left(Request("WSPFile"),7) & ".csv"  

    set objFSO=server.createobject("scripting.filesystemObject")
    
    fileoutgood = Server.MapPath("/rankings/imported/") & "\" & request("WSPFile") & "--" & month(Date) & "-" & day(date) & "-" & year(date) & "--" & left(FormatDateTime(Now, 4),2) &"-"& right(FormatDateTime(Now,4),2) & right(request("WSPFile"),4)
    fileoutbad=PathtoExceptions & "\exceptions-" & request("WSPFile") & "--" & month(Date) & "-" & day(date) & "-" & year(date) & "--" & left(FormatDateTime(Now, 4),2) &"-"& right(FormatDateTime(Now,4),2) & right(request("WSPFile"),4)
    fileoutreasons=PathtoReasons & "\exceptions-" & request("WSPFile") & "--" & month(Date) & "-" & day(date) & "-" & year(date) & "--" & left(FormatDateTime(Now, 4),2) &"-"& right(FormatDateTime(Now,4),2) & right(request("WSPFile"),4)
    
    ProcessWSP
    
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
      ' good and bad records were discovered inside the WSP File.
      
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
        <br><br><center>
 
        <% IF OldScores > 0 THEN %>
          <h4><Font color=red>Existing scores for <%=OldScores%> skiers were deleted for TourID <%=left(request("WSPFile"),7)%>.</font></h4>
        <% END IF %>
 
        <h4>New scores from <%=Request("WSPFile")%> have been imported successfully.
        <br><br>
        <font color="red"><%=goodrec%></font> of <font color="red"><%=(goodrec + badrec)%></font> skier records were successfully imported.
        <br><br>
        These scores will be included in the rankings<br>following the next overnight Re-Calculation.
        <br><br>

        <% IF badrec > 0 THEN %>

           The file had <font color="red"><%=badrec%></font> records which failed<br>verification and need to be manually resolved.
           <br><br>

           <form method=post action="exceptionmgmt-wsp.asp?file=<%=mid(fileoutbad,instr(fileoutbad,"\exceptions-")+1)&"&line=2"%>">
              <input type=submit style="width:13em" value="Process Exceptions"	
              		title="Process the exceptions found from this tournament">						
           </form>

        <% ELSE 
        	
          sTourID = ucase(left(request("WSPFile"),7))
          IF instr("LRPAB",right(sTourID,1)) > 0 THEN
            Session("TourID") = sTourID
       		
       		%>
        	    
        	 Click the button below to produce the IWWF Rankings<br>Export of all Class L/R scores, which will be sent to IWWF ...
        	 <br><br><br>

           <form method=post action="IWWF-Export.asp" method="post">
              <input type="submit" style="width:13em" value="Export to IWWF"
                  title="Produce the IWWF Rankings Export File of L/R Scores">
           </form>

        <% ELSE %>

           <form method=post action="DefaultHQ.asp?process=uploadany" method="post">
              <input type="submit" style="width:13em" value="Done with Import"
                  title="Return to the Upload Control Page">
           </form>

        <% END IF 
        END IF %>

        <br><br><br>
        </h4></center>
        <br>
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


'filein="usawaterski.org\rankings\uploads\" & request("WSPFile")  
'CSVfile="usawaterski.org\rankings\uploads\" & left(Request("WSPFile"),7) & ".csv"

'filein=PathtoRawWSPs&"\"& "07s160e.wsp"
'CSVfile=PathtoRawWSPs&"\"& "07s160e.csv"


filein=PathtoRawWSPs&"\" & request("WSPFile")  
CSVfile=PathtoRawWSPs&"\" & left(Request("WSPFile"),7) & ".csv"

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

sDSN = "FileDSN=" & PathToTRA & "WSPDelim.DSN;DefaultDir=" & PathtoRawWSPs & "\;DBQ=" & PathtoRawWSPs & "\;Extensions=csv,wsp;"

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

'    sSQL = "Select top 1 * from "&MemberTableName&" where PersonIDwithCheckDigit = '" & SQLClean(rsCSV.fields(1)) & "'"
     sSQL = "Select top 1 * from "&MemberShortTableName&" where PersonID = '" & right(SQLClean(rsCSV.fields(1)),8) & "'"
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

' Extract Team Name and Team Status code (if present following "/") into hold variables

     tempTeam = SQLClean(rsCSV.fields(8))
     IF LEN(tempTeam) > 3 THEN
        IF Mid(tempTeam,len(tempTeam)-1,1) = "/" THEN
           tempTeamStat = Right(tempTeam,1)
           tempTeam = Left(TempTeam,len(tempTeam)-2)
        ELSE 
           tempTeamStat = " "
        END IF
     ELSE
        tempTeamStat = " "
     END IF

' Build a list of valid divisions for this member.

'    First we figure out what ski year they are in
     sSQL = "Select top 1 * from " & SkiYearTableName & " where '"&CDate("01/01/"&(2000+(left(Request("WSPFile"),2))))&"' BETWEEN BeginDate and EndDate and SkiYearID <> 1"
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
     if ucase(left(rsCSV.fields(1).name,6)) <> ucase(left(request("WSPFile"),6)) then
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
                   sSQL = sSQL + " (MemberFED, MemberID, LName, FName, TourFED, TourID, [H-Class], EndDate, Event, Place, [Round], Class, Div, Perf_Qual1, Perf_Qual2, AltScore, Score, Team, TeamStat, PreZBSConvScore)"
                   sSQL = sSQL + " VALUES ("
                   sSQL = sSQL + "'" & tempFed & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(1)) & "',"
                   sSQL = sSQL + "'" & left(tempLName,17) & "',"
                   sSQL = sSQL + "'" & left(tempFName,13) & "',"
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
                   sSQL = sSQL + "'" & tempTeam & "',"
                   sSQL = sSQL + "'" & tempTeamStat & "',"
                   If ZBSFactor = 0 Then
                     sSQL = sSQL + "null)"
                   Else
                     sSQL = sSQL + "'" & SQLClean(rsCSV.fields(RoundNum * 22 + 7)) & "')"                   
                   End If
         '         sSQL = sSQL + "'" & SkiYearID & "')"
         '         When you enable this, don't forget to add it to the field list above.
         '         IF SQLClean(rsCSV.fields(1).name) = "10S127R" THEN WriteDebugSQL(sSQL)
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
                   sSQL = sSQL + " (MemberFED, MemberID, LName, FName, TourFED, TourID, [H-Class], EndDate, Event, Place, [Round], Class, Div, Team, TeamStat, Score)"
                   sSQL = sSQL + " VALUES ("
                   sSQL = sSQL + "'" & tempFed & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(1)) & "',"
                   sSQL = sSQL + "'" & left(tempLName,17) & "',"
                   sSQL = sSQL + "'" & left(tempFName,13) & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(0).name) & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(1).name) & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(4).name) & "',"
                   sSQL = sSQL + "'" & cdate(right(left(rsCSV.fields(9).name,6),2) &"/"& right(rsCSV.fields(9).name,2) &"/"& left(rsCSV.fields(9).name,4)) & "',"
                   sSQL = sSQL + "'T',"
                   sSQL = sSQL + "'" & SQLClean(trim(rsCSV.fields(13))) & "',"
                   sSQL = sSQL + "'" & RoundNum & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(RoundNum * 22 + 8)) & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(RoundNum * 22 + 9)) & "',"
                   sSQL = sSQL + "'" & tempTeam & "',"
                   sSQL = sSQL + "'" & tempTeamStat & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(RoundNum * 22 + 11)) & "')"
         '         sSQL = sSQL + "'" & SkiYearID & "')"
         '         When you enable this, don't forget to add it to the field list above.
         '         IF SQLClean(rsCSV.fields(1).name) = "10S127R" THEN WriteDebugSQL(sSQL)
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
                   sSQL = sSQL + " (MemberFED, MemberID, LName, FName, TourFED, TourID, [H-Class], EndDate, Event, Place, [Round], Class, Div, Perf_Qual1, Perf_Qual2, AltScore, Team, TeamStat, Score)"
                   sSQL = sSQL + " VALUES ("
                   sSQL = sSQL + "'" & tempFed & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(1)) & "',"
                   sSQL = sSQL + "'" & left(tempLName,17) & "',"
                   sSQL = sSQL + "'" & left(tempFName,13) & "',"
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
                   sSQL = sSQL + "'" & tempTeam & "',"
                   sSQL = sSQL + "'" & tempTeamStat & "',"
                   sSQL = sSQL + "'" & SQLClean(rsCSV.fields(RoundNum * 22 + 17)) & "')"
         '         sSQL = sSQL + "'" & SkiYearID & "')"
         '         When you enable this, don't forget to add it to the field list above.
         '         IF SQLClean(rsCSV.fields(1).name) = "10S127R" THEN WriteDebugSQL(sSQL)
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
         If NOT (objFSO.FileExists(fileoutreasons)) = TRUE Then
           Set objstreamexplainations=objFSO.createtextfile(fileoutreasons)
           objstreamexplainations.WriteLine ("FILE HEADER -- NO DATA HERE")
         Else
           Set objstreamexplainations=objFSO.opentextfile(fileoutreasons,8,TRUE)
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







%>



