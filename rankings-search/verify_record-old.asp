<!--#include virtual="/rankings/secure-settings.asp"-->

<%

  dim datain
  dim RoundNum
  dim objfso
  dim objstream
  dim fileoutgood
  dim fileoutbad
  dim fileoutexplainations
  dim errorcheck
  dim linecount
  Dim tempFed, tempFName, tempLName, tempBirthdate, tempGender, tempSkiYear, tempSkiYear_AgeCheck
  Dim PDF_Div, AgeInYears
  Dim ValidDivs, DivArray, i
  Dim ZBSFactor

  Opencon
  Set rs=Server.CreateObject("ADODB.recordset")

  ErrorCheck = ""
  ValidDivs = ""


  CSVfile=PathtoExceptions & "\" & right(left(Request("file"),18),7) & ".csv"
  fileoutgood=PathToTRA & "imported\" & right(Request("file"),len(Request("file"))-11)
  fileoutbad=PathtoExceptions & "\" & request("file")
  fileoutexplainations=PathtoReasons & "\" & request("file")

  set objFSO=server.createobject("scripting.filesystemObject")

  set objstream=objFSO.opentextfile(fileoutbad)
  do while not objstream.atendofStream
    objstream.skipline
  loop
  linecount = objstream.Line
  objstream.close

  set objstream=objFSO.opentextfile(fileoutbad)
  Do While (Not objstream.atendofstream) And objstream.line - Request("line") < 1
    datain = objstream.readline
  loop
  objstream.close


  If ucase(right(request("file"),4)) = ".PDF" Then ProcessPDF
  If ucase(right(request("file"),4)) = ".WSP" Then ProcessWSP

  CloseCon




Sub ProcessWSP

  '  We need to save the header line just in case this is the first good
  '  record to come out of this file.  In that case, we need to copy the header
  '  to the "success" file log.
  
  '  So we'll save the first line here in case we need it later.
  Dim LineOne

  set objstream=objFSO.opentextfile(fileoutbad)
  LineOne = objstream.readline
  objstream.close

  ' We need two files ... the original WSP which we will move around line by line
  ' using the Object Stream ... and the CSV version which we will open and read
  ' using ADO and the OLEDB recordset method.

  ' We will delete the CSV when we are done processing.
  objfso.CopyFile fileoutbad, CSVfile, 1


  Dim sDSN, ConCSV, rsCSV
  Dim sSQL
  Set ConCSV = Server.CreateObject("ADODB.Connection")
  Set rsCSV=Server.CreateObject("ADODB.recordset")

  sDSN = "FileDSN=" & PathToTRA & "WSPDelim.DSN;DefaultDir=" & PathToExceptions & "\;DBQ=" & PathToExceptions & "\;Extensions=csv,wsp;"
  ConCSV.Open sDSN

  sSQL = "Select * from " & CSVfile
  rsCSV.open sSQL, sDSN

  ' The line number is always +1 in WSP files because the header row doesn't count
  ' in the ADO recordset.
  ' We have to do -2 here because we compensate for the header and we also 
  ' don't have to move anywhere for the first record -- which is where we start
  ' by default.  So -1 for header and -1 so we don't move on the first record = -2.
  For i = 1 to (request("line")-2)
    rsCSV.MoveNext
  Next
  
  
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
       If tempGender <> "M" and tempGender <> "F" Then
         errorcheck = "*" & errorcheck & "Invalid Gender in MemberTrak system.  -- "
       End If      
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
       IF len(tempLName)>17 THEN tempLName = Left(tempLName,17)
'       if ucase(rsCSV.fields(3)) <> ucase(trim(rs("FirstName"))) then
'         errorcheck = errorcheck & "Member First Name Incorrect -- "
'       end if
       If ucase(left(rsCSV.fields(3),1)) <> left(ucase(trim(rs("FirstName"))),1) Then
         tempFName = SQLClean(ucase(rsCSV.fields(3)))
         errorcheck = "*" & errorcheck & "Member First Name Incorrect -- "
       Else
         tempFName = SQLClean(trim(rs("FirstName")))
       End If
       IF len(tempFName)>13 THEN tempFName = Left(tempFName,13)
     end if
     rs.Close

' Build a list of valid divisions for this member.

'    First we figure out what ski year they are in
     sSQL = "Select top 1 * from " & SkiYearTableName & " where '"&CDate("01/01/"&(2000+(right(left(Request("file"),13),2))))&"' BETWEEN BeginDate and EndDate and SkiYearID <> 1"     
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

'    Next we figure out what Ski Year to use for their age verification
'    Usually this is the same as the first ski year, but for Nationals events, it is the prior year.
     If right(ucase(left(rsCSV.fields(1).name,7)),1) = "A" Then
       sSQL = "Select top 1 * from " & SkiYearTableName & " where '"&CDate("01/01/"&(1999+(right(left(Request("file"),13),2))))&"' BETWEEN BeginDate and EndDate and SkiYearID <> 1"     
       rs.open sSQL, SConnectionToTRATable, 3, 1
       If rs.EOF Then
          sSQL = "Select top 1 * from " & SkiYearTableName & " where DefaultYear = 1"
          rs.open sSQL, SConnectionToTRATable, 3, 1
       End If
       If rs.EOF Then
          tempSkiYear_AgeCheck = "SkiYearID" ' We can't find any matching Ski Years so we just trick the SQL into ignoring this condition. (skiyearid = skiyearid)
       Else
          tempSkiYear_AgeCheck = rs("SkiYearID")
       End If
       rs.Close ' Close the Ski Year Table
     Else  ' This little twist is as a result of the new nationals scores still using the old division information.  So if nationals, we want minus 1 on the year for divisions purposes.
       tempSkiYear_AgeCheck = tempSkiYear
     End If
 
'    Next we figure out their age.
     ' NOTE: Division is based on age relative to ski year.
     If tempBirthDate <> "" Then
       ' get absolute number of years 
       AgeInYears = cint(datediff("YYYY", tempBirthDate, CDate("01/01/"&(2000+(left(rsCSV.fields(1).name,2)))))) - 1
       If tempGender = "M" Then
         sSQL = "Select distinct div from " & DivisionsTableName & " where upper(sex) = 'M' and "&AgeInYears&" <= Up_Age and "&AgeInYears&" >= Low_Age and SkiYearID = "& tempSkiYear_AgeCheck &" order by Div"
       End If 
       If tempGender = "F" Then
         sSQL = "Select distinct div from " & DivisionsTableName & " where upper(sex) = 'F' and "&AgeInYears&" <= Up_Age and "&AgeInYears&" >= Low_Age and SkiYearID = "& tempSkiYear_AgeCheck &" order by Div"
       End If
       If tempGender <> "M" and tempGender <> "F" Then ' Poison the request if the gender is not correct.
         sSQL = "Select distinct div from " & DivisionsTableName & " where 0=1"
       End If 
     Else
       If tempGender = "M" Then
         sSQL = "Select distinct div from " & DivisionsTableName & " where upper(sex) = 'M' and SkiYearID = "& tempSkiYear_AgeCheck &" order by Div"
       End If
       If tempGender = "F" Then
         sSQL = "Select distinct div from " & DivisionsTableName & " where upper(sex) = 'F' and SkiYearID = "& tempSkiYear_AgeCheck &" order by Div"
       End If
       If tempGender <> "M" and tempGender <> "F" Then ' Poison the request if the gender is not correct.
         sSQL = "Select distinct div from " & DivisionsTableName & " where 0=1"
       End If 
     End If

    rs.open sSQL, SConnectionToTRATable
    If Not rs.EOF Then 
      DivArray = rs.GetRows()
      for i = 0 to ubound(DivArray,2)
        ValidDivs = ValidDivs + DivArray(0,i)
        if i < ubound(DivArray,2) then ValidDivs = ValidDivs + ","
      Next
    Else
      ValidDivs = "***"
    End If
    rs.Close

     ' ****************************
     ' Check Tour ID
     ' ****************************
     ' Tour ID in the score record should match the file name.
     if ucase(left(rsCSV.fields(1).name,6)) <> ucase(right(left(request("file"),17),6)) then
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
                       If ValidDivs <> "***" Then
                         errorcheck = "*" & errorcheck & "Slalom " & RoundNum & " Div doesn't match age. ("&AgeInYears&") -- "
                       End If
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
                       If ValidDivs <> "***" Then
                         ErrorCheck = "*" & ErrorCheck & "Trick " & RoundNum & " Div doesn't match age. ("&AgeInYears&") -- "
                       End If
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
                       If ValidDivs <> "***" Then
                         ErrorCheck = "*" & ErrorCheck & "Jump " & RoundNum & " Div doesn't match age. ("&AgeInYears&") -- "
                       End If
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
' Remove the old bad record from the exception file.

      set objstream=objFSO.opentextfile(fileoutbad)
      textFile = "" ' this will hold the contents of the text file

      Do While not objStream.AtEndOfStream
        strFileLine = objStream.Readline
        if objstream.line - request("line") <> 1 then
          textFile = textFile & strFileLine & vbCrLf
        end if
      Loop

      objstream.close
      set objstream=objfso.opentextfile(fileoutbad,2,true)
      objstream.write(textfile)
      objstream.close

' Remove the old bad reason from the reasons file.

      set objstream=objFSO.opentextfile(fileoutexplainations)
      textFile = "" ' this will hold the contents of the text file

      Do While not objStream.AtEndOfStream
        strFileLine = objStream.Readline
        if objstream.line - request("line") <> 1 then
          textFile = textFile & strFileLine & vbCrLf
        end if
      Loop

      objstream.close
      set objstream=objfso.opentextfile(fileoutexplainations,2,true)
      objstream.write(textfile)
      objstream.close


' Write the good record to the success log.

      if Not (objFSO.FileExists(fileoutgood)) = true then
        set objstream=objFSO.createtextfile(fileoutgood)
        objstream.writeline (LineOne)
      else
        set objstream=objFSO.opentextfile(fileoutgood,8,true)
      end if
      objstream.writeline (datain)
      objstream.close

' Write the good record to the scores database.

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
                 
                   rs.Close
                   sSQL = "select zbsconversion from " & DivisionsTableName & " join " & SkiYearTableName & " on " & DivisionsTableName & ".skiyearid = " & SkiYearTableName & ".skiyearid where upper(div) = '" & ucase(SQLClean(rsCSV.fields(RoundNum * 22 + 2))) & "' and '" & cdate(right(left(rsCSV.fields(9).name,6),2) &"/"& right(rsCSV.fields(9).name,2) &"/"& left(rsCSV.fields(9).name,4)) & "' between begindate and enddate and usezbs is null"
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
                   if ZBSFactor = 0 Then
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
    
   WriteIndexPageHeader
%>

   <html><head><title>Record Verified</title></head><body>
   <center><h2>The record is now good.  The data has been <br>
saved to the database and the exception has been removed.</h2>
   <p>
<%

' The line count is always 1 more then the actual number of lines because of the way
' the end of stream loop works.  That's why we check for linecount - 1.


  If linecount-1 > 2 Then
    If request("line") > 2 Then
      If linecount-1 = request("line")+0 Then
        Response.Write "<a href=""/rankings/exceptionmgmt-wsp.asp?file=" & Request("file") & "&line=" & Request("line")-1 & """>Return to Exception Management</a>"
      Else
        Response.Write "<a href=""/rankings/exceptionmgmt-wsp.asp?file=" & Request("file") & "&line=" & Request("line") & """>Return to Exception Management</a>"
      End If          
    Else
      Response.Write "<a href=""/rankings/exceptionmgmt-wsp.asp?file=" & Request("file") & "&line=2"">Return to Exception Management</a>"
    End If 
  Else
    WriteLog(date() &"  "& time() &"  "& fileoutbad & " is now corrected and has been automatically deleted.")
    objfso.DeleteFile(fileoutbad)
    objfso.DeleteFile(fileoutexplainations)
    
		KickTrafficCounter("FixBadScores")    

%>

<h2> All of the exceptions in <%=Request("file")%> <br> have now been
corrected.&nbsp; You will now return to the main menu.</h2>

  <p>
<%

    Response.Write "<a href=""/rankings/defaultHQ.asp"">Return to Main Menu.</a>"

  End If

	WriteIndexPageFooter

     ELSE     '  If ErrorCheck <> ""


' We've already saved the record ... save the errors to the reasons file now.

   set objstream=objFSO.opentextfile(fileoutexplainations)

   textFile = "" ' this will hold the contents of the text file

   Do While not objStream.AtEndOfStream
     strFileLine = objStream.Readline
     if objstream.line - request("line") = 1 then
       if len(errorcheck) > 150 then 
         If left(errorcheck,1) = "*" Then
           textfile = textfile & "*Multiple Errors Found." & vbCrLf
         Else
           textfile = textfile & "Multiple Errors Found." & vbCrLf
         End If
       else
         textfile = textfile & errorcheck & vbCrLf
       end if
     else
       textFile = textFile & strFileLine & vbCrLf
     end if
   Loop
   objstream.close
   set objstream=objfso.opentextfile(fileoutexplainations,2,true)
   objstream.write(textfile)
   objstream.close
   WriteIndexPageHeader

%> 
   <html><head><title>Record Still Bad</title></head><body>
   <center><h2>The record is still invalid.</h2>
   <textarea rows=10 cols=80><%Response.Write(errorcheck)%></textarea>
   <p>
<%
   Response.Write "<a href=""/rankings/exceptionmgmt-wsp.asp?file=" & Request("file") & "&line=" & Request("line") & """>Return to Exception Management</a>"

WriteIndexPageFooter



    END IF

'Close our recordset and connection
rsCSV.close
Set rsCSV = Nothing
conCSV.close
set conCSV = nothing

objfso.DeleteFile CSVfile

End Sub

















Sub ProcessPDF

  set objstream=objFSO.opentextfile(fileoutbad)

  Do While (Not objstream.atendofstream) And objstream.line - Request("line") < 1
    datain = objstream.readline
  loop
  objstream.close

     ' Check MemberID exists in Master DB

     sSQL = "Select top 1 * from "&MemberTableName&" where PersonIDwithCheckDigit = '" & right(left(datain,12),9) & "'"
     rs.open sSQL, sConnectionToTRATable
     if rs.eof then
       tempFed = left(datain,3)
       tempBirthdate = ""
       tempGender = ""
       errorcheck = "*" & errorcheck & "Member ID Not Found -- "
     Else
       tempFed = trim(rs("FederationCode"))
       tempBirthdate = rs("BirthDate")
       tempGender = ucase(left(rs("Sex"),1))
       If tempGender <> "M" and tempGender <> "F" Then
         errorcheck = "*" & errorcheck & "Invalid Gender in MemberTrak system.  -- "
       End If      
     end if

     ' Check MemberID matches Member Name

     if not rs.eof then
       if ucase(trim(right(left(datain,31),17))) <> left(ucase(trim(rs("LastName"))),17) then
         tempLName = replace(trim(right(left(datain,31),17)), "'", "''")
         errorcheck = "*" & errorcheck & "Member Last Name Incorrect -- "
       Else
         tempLName = replace(trim(rs("LastName")), "'", "''")
       end if
       IF len(tempLName)>17 THEN tempLName = Left(tempLName,17)
'       if ucase(trim(right(left(datain,44),13))) <> ucase(trim(rs("FirstName"))) then
'         errorcheck = errorcheck & "Member First Name Incorrect -- "
'       end if
       If ucase(trim(right(left(datain,32),1))) <> left(ucase(trim(rs("FirstName"))),1) Then
         tempFName = replace(trim(right(left(linetext,32),1)), "'", "''")
         errorcheck = "*" & errorcheck & "Member First Name Incorrect -- "
       Else
         tempFName = replace(trim(rs("FirstName")), "'", "''")
       End If
       IF len(tempFName)>13 THEN tempFName = Left(tempFName,13)
     End If
     rs.close

' Build a list of valid divisions for this member.

'    First we figure out what ski year they are in
     sSQL = "Select top 1 * from " & SkiYearTableName & " where '"&CDate("01/01/"&(2000+(right(left(datain,65),2))))&"' BETWEEN BeginDate and EndDate and SkiYearID <> 1"     
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

'    Next we figure out what Ski Year to use for their age verification
'    Usually this is the same as the first ski year, but for Nationals events, it is the prior year.
     If ucase(right(left(datain,70),1)) = "A" Then
       sSQL = "Select top 1 * from " & SkiYearTableName & " where '"&CDate("01/01/"&(1999+(right(left(datain,65),2))))&"' BETWEEN BeginDate and EndDate and SkiYearID <> 1"     
       rs.open sSQL, SConnectionToTRATable, 3, 1
       If rs.EOF Then
          sSQL = "Select top 1 * from " & SkiYearTableName & " where DefaultYear = 1"
          rs.open sSQL, SConnectionToTRATable, 3, 1
       End If
       If rs.EOF Then
          tempSkiYear_AgeCheck = "SkiYearID" ' We can't find any matching Ski Years so we just trick the SQL into ignoring this condition. (skiyearid = skiyearid)
       Else
          tempSkiYear_AgeCheck = rs("SkiYearID")
       End If
       rs.Close ' Close the Ski Year Table
     Else  ' This little twist is as a result of the new nationals scores still using the old division information.  So if nationals, we want minus 1 on the year for divisions purposes.
       tempSkiYear_AgeCheck = tempSkiYear
     End If

 
'    Next we figure out their age.
     ' NOTE: Division is based on age relative to ski year.
     If tempBirthDate <> "" Then
       ' get absolute number of years 
       AgeInYears = cint(datediff("YYYY", tempBirthDate, CDate("01/01/"&(2000+(right(left(linetext,65),2)))))) - 1
          If tempGender = "M" Then
            sSQL = "Select distinct div from " & DivisionsTableName & " where ((left(Div,1) = 'M' or right(Div,1) = 'M') or (left(Div,1) = 'B' or right(Div,1) = 'B')) and "&AgeInYears&" <= Up_Age and "&AgeInYears&" >= Low_Age and SkiYearID = "& tempSkiYear_AgeCheck &" order by Div"
          End If
          If tempGender = "F" Then
            sSQL = "Select distinct div from " & DivisionsTableName & " where ((left(Div,1) = 'W' or right(Div,1) = 'W') or (left(Div,1) = 'G' or right(Div,1) = 'G')) and "&AgeInYears&" <= Up_Age and "&AgeInYears&" >= Low_Age and SkiYearID = "& tempSkiYear_AgeCheck &" order by Div"
          End If
          If tempGender <> "M" and tempGender <> "F" Then ' Poison the request if the gender is not correct.
            sSQL = "Select distinct div from " & DivisionsTableName & " where 0=1"
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
    Else
      ValidDivs = "***"
    End If
    rs.Close
 
     ' Check blanks between memberid and lastname
     if right(left(datain,14),2) <> "  " then
       errorcheck = errorcheck & "No Blanks Between MemberID and Lastname (14) -- "
     end if 

     ' Check Tour ID
     if ucase(right(left(datain,69),6)) <> ucase(right(left(request("file"),17),6)) then
       errorcheck = errorcheck & "Field 'Tour ID' Doesn't Match Filename (70) -- "
     end if

     if (right(left(datain,131),5) <> "     ") and (not isnumeric(trim(right(left(datain,93),3)))) and (trim(right(left(datain,93),3)) <> "") then
       errorcheck = errorcheck & "Slalom Placement is not a number (93) -- "
     end if
     if (right(left(datain,145),5) <> "     ") and (not isnumeric(trim(right(left(datain,97),3)))) and (trim(right(left(datain,97),3)) <> "") then
       errorcheck = errorcheck & "Trick Placement is not a number (97) -- "
     end if
     if (right(left(datain,166),4) <> "     ") and (not isnumeric(trim(right(left(datain,101),3)))) and (trim(right(left(datain,101),3)) <> "") then
       errorcheck = errorcheck & "Jump Placement is not a number (101) -- "
     end if

     ' Check Valid Tour Date
     if not isDate(trim(right(left(datain,78),2)) & "/" & trim(right(left(datain,80),2)) & "/" & trim(right(left(datain,76),4))) then
       errorcheck = errorcheck & "Tour End Date not a valid date.  -- "
     end if


     RoundNum = 1
     Do While len(trim(datain)) > (RoundNum * 60 + 52)

' Check start of round
       if (right(left(datain,(RoundNum * 60 + 51)),1) <> cstr(RoundNum) and trim(right(left(datain,(RoundNum * 60 + 107)),40)) <> "") then
         errorcheck = errorcheck & "Round " & RoundNum & " scores present but no Round " & RoundNum & " Marker (" & RoundNum * 60 + 51 & ")  -- "
       end if

'  Check Slalom Scores (if they exist)
       if right(left(datain,(RoundNum * 60 + 71)),5) <> "     " then
         if not isalpha( right(left(datain,(RoundNum * 60 + 52)),1) ) then 
           errorcheck = errorcheck & "Slalom " & RoundNum & " class is not a valid character. (" & RoundNum * 60 + 52 & ") -- "
         end if

         ' Check Slalom Division Based on Age and Gender
           PDF_Div = ucase(right(left(datain,(RoundNum * 60 + 54)),2))

           if instr(ValidDivs,PDF_Div) = 0 then 
              If left(tempGender,1) = "M" And (instr(PDF_Div,"M") = 0 and instr(PDF_Div,"B") = 0) Or left(tempGender,1) = "F" And (instr(PDF_Div,"F") = 0 and instr(PDF_Div,"G") = 0)  Then
                 errorcheck = errorcheck & "Slalom "&RoundNum&" Div doesn't match gender. -- "
              Else
                 If ValidDivs <> "***" Then
                   errorcheck = errorcheck & "Slalom " & RoundNum & " Div doesn't match age "&AgeInYears&". (" & RoundNum * 60 + 54 & ") -- "
                 End If
              End If
           end if
         
         if not isnumeric(right(left(datain,(RoundNum * 60 + 66)),4)) then
           errorcheck = errorcheck & "Slalom " & RoundNum & " EndPass Line is not a number (" & RoundNum * 60 + 66 & ") -- "
         end if
         if not isnumeric(right(left(datain,(RoundNum * 60 + 62)),2)) then
           errorcheck = errorcheck & "Slalom " & RoundNum & " EndPass Speed is not a number (" & RoundNum * 60 + 62 & ") -- "
         end if
         if not isnumeric(right(left(datain,(RoundNum * 60 + 60)),4)) then
           errorcheck = errorcheck & "Slalom " & RoundNum & " EndPass Score is not a number (" & RoundNum * 60 + 60 & ") -- "
         end if
         if not isnumeric(right(left(datain,(RoundNum * 60 + 71)),5)) then
           errorcheck = errorcheck & "Slalom " & RoundNum & " Score is not a number (" & RoundNum * 60 + 71 & ") -- "
         end if
       end if

' Check blanks between slalom1 and trick1
       if right(left(datain,(RoundNum * 60 + 75)),4) <> "    " then
         errorcheck = errorcheck & "Blanks missing between slalom " & RoundNum & " and trick " & RoundNum & " scores (" & RoundNum * 60 + 75 & ") -- "
       end if


'  Check Trick Scores (if they exist)
       if right(left(datain,(RoundNum * 60 + 85)),5) <> "     " then
         if not isalpha(right(left(datain,(RoundNum * 60 + 76)),1)) then 
           errorcheck = errorcheck & "Trick " & RoundNum & " class is not a valid character. (" & RoundNum * 60 + 76 & ") -- "
         end if

         ' Check Trick Division Based on Age and Gender
           PDF_Div = ucase(right(left(datain,(RoundNum * 60 + 78)),2))

           if instr(ValidDivs,PDF_Div) = 0 then 
              If left(tempGender,1) = "M" And (instr(PDF_Div,"M") = 0 and instr(PDF_Div,"B") = 0) Or left(tempGender,1) = "F" And (instr(PDF_Div,"F") = 0 and instr(PDF_Div,"G") = 0)  Then
                 errorcheck = errorcheck & "Trick "&RoundNum&" Div doesn't match gender. -- "
              Else
                 If ValidDivs <> "***" Then
                   errorcheck = errorcheck & "Trick " & RoundNum & " Div doesn't match age "&AgeInYears&". (" & RoundNum * 60 + 78 & ") -- "
                 End If
              End If
           End If

         if not isnumeric(right(left(datain,(RoundNum * 60 + 85)),5)) then
           errorcheck = errorcheck & "Trick " & RoundNum & " Score is not a number (" & RoundNum * 60 + 85 & ") -- "
         end if
       end if

' Check blanks between trick and jump
       if right(left(datain,(RoundNum * 60 + 88)),3) <> "   " then
         errorcheck = errorcheck & "Blanks missing between trick " & RoundNum & " and jump " & RoundNum & " scores (" & RoundNum * 60 + 88 & ") -- "
       end if

'  Check Jump Scores (if they exist)
       if right(left(datain,(RoundNum * 60 + 106)),4) <> "    " then
         if not isalpha(right(left(datain,(RoundNum * 60 + 89)),1)) then 
           errorcheck = errorcheck & "Jump " & RoundNum & " class is not a valid character. (" & RoundNum * 60 + 89 & ") -- "
         end if

         ' Check Jump Division Based on Age and Gender
           PDF_Div = ucase(right(left(datain,(RoundNum * 60 + 91)),2))

           if instr(ValidDivs,PDF_Div) = 0 then 
              If left(tempGender,1) = "M" And (instr(PDF_Div,"M") = 0 and instr(PDF_Div,"B") = 0) Or left(tempGender,1) = "F" And (instr(PDF_Div,"F") = 0 and instr(PDF_Div,"G") = 0)  Then
                 errorcheck = errorcheck & "Jump "&RoundNum&" Div doesn't match gender. -- "
              Else
                 If ValidDivs <> "***" Then
                   errorcheck = errorcheck & "Jump " & RoundNum & " Div doesn't match age "&AgeInYears&". (" & RoundNum * 60 + 91 & ") -- "
                 End If
              End If
           end if

         if not isnumeric(right(left(datain,(RoundNum * 60 + 97)),4)) then
           errorcheck = errorcheck & "Jump " & RoundNum & " Ramp Height is not a number (" & RoundNum * 60 + 97 & ") -- "
         end if
         if not isnumeric(right(left(datain,(RoundNum * 60 + 99)),2)) then
           errorcheck = errorcheck & "Jump " & RoundNum & " Jump Speed is not a number (" & RoundNum * 60 + 99 & ") -- "
         end if
         if not isnumeric(right(left(datain,(RoundNum * 60 + 106)),4)) then
           If trim(right(left(datain,(RoundNum * 60 + 106)),4)) <> ".F" and trim(right(left(datain,(RoundNum * 60 + 106)),4)) <> ".P" Then
             errorcheck = errorcheck & "Jump " & RoundNum & " Distance (Meters) is not a number (" & RoundNum * 60 + 106 & ") -- "
           End If
         End If
         if not isnumeric(right(left(datain,(RoundNum * 60 + 102)),3)) then
           If trim(right(left(datain,(RoundNum * 60 + 102)),3)) <> "F" And trim(right(left(datain,(RoundNum * 60 + 102)),3)) <> "P" Then
             errorcheck = errorcheck & "Jump " & RoundNum & " Distance (Feet) is not a number (" & RoundNum * 60 + 102 & ") -- "
           End If
         end if

         if ((right(left(datain,(RoundNum * 60 + 94)),1) <> "." or (right(left(datain,(RoundNum * 60 + 105)),1) <> "." and right(left(datain,(RoundNum * 60 + 104)),1) <> ".")) and right(left(datain,(RoundNum * 60 + 99)),1) <> " ") then
           If trim(right(left(datain,(RoundNum * 60 + 106)),4)) <> ".F" and trim(right(left(datain,(RoundNum * 60 + 106)),4)) <> ".P" Then
             errorcheck = errorcheck & "Decimal not found for jump score (" & RoundNum * 60 + 94 & "/" & RoundNum * 60 + 105 & ") -- "
           End If 
         end if
       end if

       RoundNum = RoundNum + 1
     Loop

   if errorcheck = "" then

' Remove the old bad record from the exception file.

      set objstream=objFSO.opentextfile(fileoutbad)
      textFile = "" ' this will hold the contents of the text file

      Do While not objStream.AtEndOfStream
        strFileLine = objStream.Readline
        if objstream.line - request("line") <> 1 then
          textFile = textFile & strFileLine & vbCrLf
        end if
      Loop

      objstream.close
      set objstream=objfso.opentextfile(fileoutbad,2,true)
      objstream.write(textfile)
      objstream.close

' Remove the old bad reason from the reasons file.

      set objstream=objFSO.opentextfile(fileoutexplainations)
      textFile = "" ' this will hold the contents of the text file

      Do While not objStream.AtEndOfStream
        strFileLine = objStream.Readline
        if objstream.line - request("line") <> 1 then
          textFile = textFile & strFileLine & vbCrLf
        end if
      Loop

      objstream.close
      set objstream=objfso.opentextfile(fileoutexplainations,2,true)
      objstream.write(textfile)
      objstream.close


' Write the good record to the success log.

      if Not (objFSO.FileExists(fileoutgood)) = true then
        set objstream=objFSO.createtextfile(fileoutgood)
      else
        set objstream=objFSO.opentextfile(fileoutgood,8,true)
      end if
      objstream.writeline (datain)
      objstream.close

' Write the good record to the scores database.

       OpenCon
       RoundNum = 1
       Do While len(trim(datain)) > (RoundNum * 60 + 52)
     ' ****************************
     '  Add in Slalom Scores (if they exist)
     ' ****************************
         if right(left(datain,(RoundNum * 60 + 71)),5) <> "     " then
          ' ****************************
          '  Before adding any score, check for duplicates
          ' ****************************
             set rs=Server.CreateObject("ADODB.recordset")
             sSQL = "Select * from " & RawScoresTableName
             sSQL = sSQL + " where MemberID = '" & right(left(datain,12),9) & "' and"
             sSQL = sSQL + " TourID = '" & right(left(datain,71),8) & "' and"
             sSQL = sSQL + " Event = 'S' and"
             sSQL = sSQL + " [Round] = '" & RoundNum & "' and"
             sSQL = sSQL + " Div = '" & right(left(datain,(RoundNum * 60 + 54)),2) & "'"
             rs.open sSQL, sConnectionToTRATable, 3,3

           ' If there are no duplicate records and there have been no errors
           If rs.eof and (Err.Number = 0) Then 

             rs.Close
             sSQL = "select zbsconversion from " & DivisionsTableName & " join " & SkiYearTableName & " on " & DivisionsTableName & ".skiyearid = " & SkiYearTableName & ".skiyearid where upper(div) = '" & ucase(right(left(datain,(RoundNum * 60 + 54)),2)) & "' and '" & cdate(right(left(datain,78),2) &"/"& right(left(datain,80),2) &"/"& right(left(datain,76),4)) & "' between begindate and enddate"
             rs.open sSQL, sConnectionToTRATable, 3,3

             If rs.eof Then
               ZBSFactor = 0
               WriteLog ("ERROR - ZBS could not be adjusted with Division '" & ucase(right(left(datain,(RoundNum * 60 + 54)),2)) & "' and '" & cdate(right(left(datain,78),2) &"/"& right(left(datain,80),2) &"/"& right(left(datain,76),4)) & "'.")
             Else
               ZBSFactor = rs("ZBSConversion")
             End If


             sSQL = "insert into " & RawScoresTableName
             sSQL = sSQL + " (MemberFED, MemberID, LName, FName, TourFED, TourID, [H-Class], EndDate, Event, Place, [Round], Class, Div, Perf_Qual1, Perf_Qual2, AltScore, Score, Team, PreZBSConvScore)"
             sSQL = sSQL + " VALUES ("
             sSQL = sSQL + "'" & tempFed & "',"
             sSQL = sSQL + "'" & right(left(datain,12),9) & "',"
             sSQL = sSQL + "'" & tempLName & "',"
             sSQL = sSQL + "'" & tempFName & "',"
             sSQL = sSQL + "'" & right(left(datain,63),3) & "',"
             sSQL = sSQL + "'" & right(left(datain,71),8) & "',"
             sSQL = sSQL + "'" & right(left(datain,72),1) & "',"
             sSQL = sSQL + "'" & cdate(right(left(datain,78),2) &"/"& right(left(datain,80),2) &"/"& right(left(datain,76),4)) & "',"
             sSQL = sSQL + "'S',"
             if isnumeric(trim(right(left(datain,93),3))) then
               sSQL = sSQL + "'" & right(left(datain,93),3) & "',"
             else
               sSQL = sSQL + "NULL,"
             end if
             sSQL = sSQL + "'" & RoundNum & "',"
             sSQL = sSQL + "'" & right(left(datain,(RoundNum * 60 + 52)),1) & "',"
             sSQL = sSQL + "'" & right(left(datain,(RoundNum * 60 + 54)),2) & "',"
             sSQL = sSQL + "'" & right(left(datain,(RoundNum * 60 + 66)),4) & "',"
             sSQL = sSQL + "'" & right(left(datain,(RoundNum * 60 + 62)),2) & "',"
             sSQL = sSQL + "'" & right(left(datain,(RoundNum * 60 + 60)),4) & "',"
             sSQL = sSQL + "'" & (right(left(datain,(RoundNum * 60 + 71)),5) + ZBSFactor) & "',"
             sSQL = sSQL + "'" & SQLClean(rsCSV.fields(8)) & "',"
             sSQL = sSQL + "'" & right(left(datain,(RoundNum * 60 + 71)),5) & "')"
   '         sSQL = sSQL + "'" & SkiYearID & "')"
   '         When you enable this, don't forget to add it to the field list above.
             Con.Execute(sSQL)
           End If
           rs.close
         end if
     ' ****************************
     '  Add in Trick Scores (if they exist)
     ' ****************************
         if right(left(datain,(RoundNum * 60 + 85)),5) <> "     " then
          ' ****************************
          '  Before adding any score, check for duplicates
          ' ****************************
             set rs=Server.CreateObject("ADODB.recordset")
             sSQL = "Select * from " & RawScoresTableName
             sSQL = sSQL + " where MemberID = '" & right(left(datain,12),9) & "' and"
             sSQL = sSQL + " TourID = '" & right(left(datain,71),8) & "' and"
             sSQL = sSQL + " Event = 'T' and"
             sSQL = sSQL + " [Round] = '" & RoundNum & "' and"
             sSQL = sSQL + " Div = '" & right(left(datain,(RoundNum * 60 + 78)),2) & "'"
             rs.open sSQL, sConnectionToTRATable, 3,3

           ' If there are no duplicate records and there have been no errors
           If rs.eof and (Err.Number = 0) Then 
             sSQL = "insert into " & RawScoresTableName
             sSQL = sSQL + " (MemberFED, MemberID, LName, FName, TourFED, TourID, [H-Class], EndDate, Event, Place, [Round], Class, Div, Team, Score)"
             sSQL = sSQL + " VALUES ("
             sSQL = sSQL + "'" & tempFed & "',"
             sSQL = sSQL + "'" & right(left(datain,12),9) & "',"
             sSQL = sSQL + "'" & tempLName & "',"
             sSQL = sSQL + "'" & tempFName & "',"
             sSQL = sSQL + "'" & right(left(datain,63),3) & "',"
             sSQL = sSQL + "'" & right(left(datain,71),8) & "',"
             sSQL = sSQL + "'" & right(left(datain,72),1) & "',"
             sSQL = sSQL + "'" & cdate(right(left(datain,78),2) &"/"& right(left(datain,80),2) &"/"& right(left(datain,76),4)) & "',"
             sSQL = sSQL + "'T',"
             if isnumeric(trim(right(left(datain,97),3))) then
               sSQL = sSQL + "'" & right(left(datain,97),3) & "',"
             else
               sSQL = sSQL + "NULL,"
             end if
             sSQL = sSQL + "'" & RoundNum & "',"
             sSQL = sSQL + "'" & right(left(datain,(RoundNum * 60 + 76)),1) & "',"
             sSQL = sSQL + "'" & right(left(datain,(RoundNum * 60 + 78)),2) & "',"
             sSQL = sSQL + "'" & SQLClean(rsCSV.fields(8)) & "',"
             sSQL = sSQL + "'" & right(left(datain,(RoundNum * 60 + 85)),5) & "')"
   '         sSQL = sSQL + "'" & SkiYearID & "')"
   '         When you enable this, don't forget to add it to the field list above.
             Con.Execute(sSQL)
           End If
           rs.close
         end if
     ' ****************************
     '  Add in Jump Scores (if they exist)
     ' ****************************     
         if right(left(datain,(RoundNum * 60 + 106)),4) <> "    " then
          ' ****************************
          '  Before adding any score, check for duplicates
          ' ****************************
             set rs=Server.CreateObject("ADODB.recordset")
             sSQL = "Select * from " & RawScoresTableName
             sSQL = sSQL + " where MemberID = '" & right(left(datain,12),9) & "' and"
             sSQL = sSQL + " TourID = '" & right(left(datain,71),8) & "' and"
             sSQL = sSQL + " Event = 'J' and"
             sSQL = sSQL + " [Round] = '" & RoundNum & "' and"
             sSQL = sSQL + " Div = '" & right(left(datain,(RoundNum * 60 + 91)),2) & "'"
             rs.open sSQL, sConnectionToTRATable, 3,3

           ' If there are no duplicate records and there have been no errors
           If rs.eof and (Err.Number = 0) Then 
             sSQL = "insert into " & RawScoresTableName
             sSQL = sSQL + " (MemberFED, MemberID, LName, FName, TourFED, TourID, [H-Class], EndDate, Event, Place, [Round], Class, Div, Perf_Qual1, Perf_Qual2, AltScore, Team, Score)"
             sSQL = sSQL + " VALUES ("
             sSQL = sSQL + "'" & tempFed & "',"
             sSQL = sSQL + "'" & right(left(datain,12),9) & "',"
             sSQL = sSQL + "'" & tempLName & "',"
             sSQL = sSQL + "'" & tempFName & "',"
             sSQL = sSQL + "'" & right(left(datain,63),3) & "',"
             sSQL = sSQL + "'" & right(left(datain,71),8) & "',"
             sSQL = sSQL + "'" & right(left(datain,72),1) & "',"
             sSQL = sSQL + "'" & cdate(right(left(datain,78),2) &"/"& right(left(datain,80),2) &"/"& right(left(datain,76),4)) & "',"
             sSQL = sSQL + "'J',"
             if isnumeric(trim(right(left(datain,101),3))) then
               sSQL = sSQL + "'" & right(left(datain,101),3) & "',"
             else
               sSQL = sSQL + "NULL,"
             end if
             sSQL = sSQL + "'" & RoundNum & "',"
             sSQL = sSQL + "'" & right(left(datain,(RoundNum * 60 + 89)),1) & "',"
             sSQL = sSQL + "'" & right(left(datain,(RoundNum * 60 + 91)),2) & "',"
             sSQL = sSQL + "'" & right(left(datain,(RoundNum * 60 + 97)),4) & "',"
             sSQL = sSQL + "'" & right(left(datain,(RoundNum * 60 + 99)),2) & "',"
             If trim(right(left(datain,(RoundNum * 60 + 106)),4)) = ".F" Or trim(right(left(datain,(RoundNum * 60 + 106)),4)) = ".P" Then
               sSQL = sSQL + "'0.0',"
             Else
               sSQL = sSQL + "'" & right(left(datain,(RoundNum * 60 + 106)),4) & "',"
             End If
             sSQL = sSQL + "'" & SQLClean(rsCSV.fields(8)) & "',"
             If trim(right(left(datain,(RoundNum * 60 + 102)),3)) = "F" Or trim(right(left(datain,(RoundNum * 60 + 102)),3)) = "P" Then
               sSQL = sSQL + "'0')"
             Else
               sSQL = sSQL + "'" & right(left(datain,(RoundNum * 60 + 102)),3) & "')"
             End If
   '         sSQL = sSQL + "'" & SkiYearID & "')"
   '         When you enable this, don't forget to add it to the field list above.
             Con.Execute(sSQL)
           End If
           rs.close
         end if
         RoundNum = RoundNum + 1
       Loop          
       



%>
   <html><head><title>Record Verified</title></head><body>
   <center><h2>The record is now good.  The data has been <br>
saved to the database and the exception has been removed.</h2>
   <p>
<%


' The line count is always 1 more then the actual number of lines because of the way
' the end of stream loop works.  That's why we check for linecount - 1.


  If linecount-1 > 1 Then
    If request("line") > 1 Then
      If linecount-1 = request("line") Then
        Response.Write "<a href=""/rankings/exceptionmgmt-pdf.asp?file=" & Request("file") & "&line=" & Request("line")-1 & """>Return to Exception Management</a>"
      Else
        Response.Write "<a href=""/rankings/exceptionmgmt-pdf.asp?file=" & Request("file") & "&line=" & Request("line") & """>Return to Exception Management</a>"
      End If        
    Else
      Response.Write "<a href=""/rankings/exceptionmgmt-pdf.asp?file=" & Request("file") & "&line=1"">Return to Exception Management</a>"
    End If 
  Else
    WriteLog(date() &"  "& time() &"  "& fileoutbad & " is now corrected and has been automatically deleted.")
    objfso.DeleteFile(fileoutbad)
    objfso.DeleteFile(fileoutexplainations)

		KickTrafficCounter("FixBadScores")    

%>
<h2> All of the exceptions in <%=Request("file")%> <br> have now been
corrected.&nbsp; You will now return to the main menu.</h2>
  <p>
<%

    Response.Write "<a href=""/rankings/defaultHQ.asp"">Return to Main Menu.</a>"
  end if

   else

' We've already saved the record ... save the errors to the reasons file now.

   set objstream=objFSO.opentextfile(fileoutexplainations)

   textFile = "" ' this will hold the contents of the text file

   Do While not objStream.AtEndOfStream
     strFileLine = objStream.Readline
     if objstream.line - request("line") = 1 then
       if len(errorcheck) > 150 then 
         If left(errorcheck,1) = "*" Then
           textfile = textfile & "*Multiple Errors Found." & vbCrLf
         Else
           textfile = textfile & "Multiple Errors Found." & vbCrLf
         End If
       else
         textfile = textfile & errorcheck & vbCrLf
       end if
     else
       textFile = textFile & strFileLine & vbCrLf
     end if
   Loop
   objstream.close
   set objstream=objfso.opentextfile(fileoutexplainations,2,true)
   objstream.write(textfile)
   objstream.close
%> 
   <html><head><title>Record Still Bad</title></head><body>
   <center><h2>The record is still invalid.</h2>
   <textarea rows=10 cols=80><%Response.Write(errorcheck)%></textarea>
   <p>
<%
   Response.Write "<a href=""/rankings/exceptionmgmt-pdf.asp?file=" & Request("file") & "&line=" & Request("line") & """>Return to Exception Management</a>"


   end if

End Sub

%>



