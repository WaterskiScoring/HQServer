<!--#include virtual="/rankings/secure-settings.asp"-->
<%
Response.Buffer = True
Server.ScriptTimeout = 36000 


If (trim(request("process")) = "clearflags") Then

  OpenCon
  sSQL = "Update " & RawScoresTableName & " set CorsonDL = NULL where 1=1"
  Con.Execute(sSQL)
  CloseCon

  WriteIndexPageHeader%>
    <br><br>
    <center><h2>Previously Downloaded Record<br>Flags cleared.</h2><br><br><br>
    <h4>The next time you generate an export, all records will be downloaded.</h4>
   <%
   WriteIndexPageFooter    

End If

Dim tempLast, tempFirst, tempAltScore, tempPlace, tempPQ1, tempPQ2, tempSex, tempYOB, tempConv, tempMiss, tempSpecial, Off_Score
Dim RecordsSaved
Dim InsertCmd
Dim PreZBSScore
' Scores are now saved in ZBS format but this export wants to use the old format.
' So we get the score from the DB but then immediately convert it in pre-zbs format.


If Request("SelectTours") = "on" Then
  If Not isDate(Request("begindate")) Then Response.Redirect("/rankings/defaultHQ.asp?process=corson&failure=1")
  If Not isDate(Request("enddate")) Then Response.Redirect("/rankings/defaultHQ.asp?process=corson&failure=1")

  WriteIndexPageHeader
  NewsTitle="IWSF Export"
  News="The following tournaments are between the date range you specified.  <br><br>Highlight the tournaments you wish to download scores for. <br><br>To select more then one tournament, hold down the CTRL key on your keyboard while selecting the individual tournaments with your mouse."
    %>
    <br><br>
    <center><h2>Select Tournaments to Download.</h2><br>
    <%
    OpenCon
      set rs = Server.CreateObject("ADODB.recordset")
      sSQL = "Select distinct TourID from " & RawScoresTableName
      sSQL = sSQL + " where [class] in ('R','L','A','B') and EndDate >= '" & SQLClean(Request("begindate")) & "' and EndDate <= '" & SQLClean(Request("enddate")) & "'"  
      sSQL = sSQL + " order by TourID"

      WriteDebugSQL(sSQL)

      rs.open sSQL, SConnectionToTRATable, 3, 3

      If rs.eof Then 
        Response.Redirect("/rankings/defaultHQ.asp?process=corson&failure=2")
      Else
        %>
        You have chosen to download <% 
          If request("PreviouslySelected") = "No" Then 
            response.write("only qualifying scores which were not previously downloaded.")
          Else
            response.write("all qualifying scores from the tournaments selected below.")
          End If %>
        <br><br>
        <b>To select multiple entries, hold down CTRL while selecting tournaments.</b><br><br>
        <FORM method="post" action="/rankings/corson.asp">
        <select multiple name='ToursToDownload' size="10">
        <%

        do while not rs.eof
          %>
          <option value="++<%=trim(rs("TourID"))%>++"><%=rs("TourID")%><br>
          <%      
          rs.movenext
        loop
        
        %>
        </select><br><br>
        <%

      End If

      rs.close
      set rs = nothing
    CloseCon
    %>
    <INPUT type="hidden" name="process" value="corson"><br><br>
    <INPUT type="hidden" name="begindate" value="<%=Request("begindate")%>">
    <INPUT type="hidden" name="enddate" value="<%=Request("enddate")%>">
    <INPUT type="hidden" name="PreviouslySelected" value="<%=Request("PreviouslySelected")%>">
    <INPUT type="Submit" value="Create Export File">
    </FORM>
   <%
   WriteIndexPageFooter    
End If

If (trim(request("process")) = "corson") and Request("SelectTours") <> "on" Then

  RecordsSaved = 0
  
  If Not isDate(Request("begindate")) Then Response.Redirect("/rankings/defaultHQ.asp?process=corson&failure=1")
  If Not isDate(Request("enddate")) Then Response.Redirect("/rankings/defaultHQ.asp?process=corson&failure=1")


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
<B>Processing Records for Export.<br><br>
Please wait a moment ...<br><br>  
</B></FONT>
<IMG SRC="/rankings/images/buttons/wait.gif" BORDER=1 WIDTH=75 HEIGHT=15><BR><BR>
</TD>
</TR>
</TABLE>
</DIV>



  <%
  response.flush

  'Copy the table structure from our original template.
  Set tempFSO = Server.CreateObject("Scripting.FileSystemObject")

  Set tempObjStream = tempFSO.GetFile(PathtoTRA & "news\CORSON-original.DBF")
  tempObjStream.Copy(PathtoTRA & "news\CORSON.DBF")

  set tempObjStream = nothing
  set tempFSO = nothing
  
  
  'Open connection For DBF files In C:\ folder
  Dim DBFCon
  Set DBFCon = CreateObject("ADODB.Connection")
  DBFCon.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                   "Data Source=" & PathtoTRA & "news\;" & _
                   "Extended Properties=""DBASE IV;"";" 

  'Open Raw Scores Table
  OpenCon
  set rs = Server.CreateObject("ADODB.recordset")
  sSQL = "Select RS.FName, RS.LName, RS.MemberID, MT.FederationCode as MemberFed, RS.TourID, RS.EndDate, RS.Event, RS.Div, RS.Class, RS.Round, RS.Place, RS.Perf_Qual1, RS.Perf_Qual2, RS.AltScore, RS.Score, MT.Sex, MT.BirthDate, DT.ZBSConversion, ST.TSiteID"
  sSQL = sSQL + " from " & RawScoresTableName & " as RS left join "&MemberTableName&" as MT on RS.MemberID = MT.PersonIDwithCheckDigit"
  sSQL = sSQL + " left join " & SanctionTableName & " as ST on lower(left(RS.TourID,6)) = lower(left(ST.TSanction,6))"
  sSQL = sSQL + " left join " & SkiYearTableName & " as SY on RS.EndDate between SY.BeginDate and SY.EndDate and SY.SkiYearID <> 1"
  sSQL = sSQL + " left join " & DivisionsTableName & " as DT on RS.Div = DT.Div and SY.skiyearid = DT.skiyearid"
  sSQL = sSQL + " where RS.Class in ('R','L','A','B') and RS.EndDate >= '" & SQLClean(Request("begindate")) & "' and RS.EndDate <= '" & SQLClean(Request("enddate")) & "'"  
  If Request("PreviouslySelected") = "No" Then sSQL = sSQL + " and RS.CorsonDL IS NULL"
  If trim(Request("ToursToDownload")) <> "" Then sSQL = sSQL + " and RS.TourID in (" & SQLClean(request("ToursToDownload")) & ")"
  sSQL = sSQL + " order by case RS.[class] when 'R' then 1 when 'L' then 2 when 'A' then 3 when 'B' then 4 end, RS.memberid, RS.EndDate, RS.Round, RS.Event"

  WriteDebugSQL(sSQL)

  rs.open sSQL, SConnectionToTRATable, 3, 3

  do while not rs.eof
    If rs("Event") = "S" Then
      PreZBSScore = rs("Score") - rs("ZBSConversion")
    Else
      PreZBSScore = rs("Score")
    End If

    InsertCmd = ""
    tempSex = left(rs("sex"),1)
    tempYOB = right(rs("birthdate"),4)  
    '
    ' Some names have a special character (apostrophe) which
    ' SQL is fine with but DBase chokes.  Replace it with a blank.
    '
    tempFirst = left(ucase(rs("FName")),1)
    tempFirst = tempFirst + mid(lcase(rs("FName")),2)
    tempFirst = replace(tempFirst,"'","")
    tempLast = replace(ucase(rs("LName")),"'","")

    tempPlace = rs("Place")
    if ucase(right(tempPlace,1)) = "T" Then tempPlace = left(trim(tempPlace),len(tempPlace)-1)

    if rs("AltScore") <> "" then tempAltScore = rs("AltScore") else tempAltScore = NULL
    if rs("Perf_Qual1") <> "" then tempPQ1 = rs("Perf_Qual1") else tempPQ1 = NULL
    if rs("Perf_Qual2") <> "" then tempPQ2 = rs("Perf_Qual2") else tempPQ2 = NULL
    
    select case rs("Event")
    case "S"
      if tempPQ1 > 0 then tempPQ1 = tempPQ1 / 100
      If (tempSex = "M" and (tempPQ2 = 36 or tempPQ2 = 58)) or (tempSex = "F" and (tempPQ2 = 34 or tempPQ2 = 55)) then
        tempConv = 0
        If left(ucase(rs("Div")),1) = "W" or ucase(rs("Div")) = "OW" or left(ucase(rs("Div")),1) = "G" then
          Off_Score = PreZBSScore - 6
          tempConv = PreZBSScore
        Else
          Off_Score = PreZBSScore
        End If
        if PreZBSScore < 6 then tempMiss = "Y" else tempMiss = "N"
       
        'Insert row To the table
        ' CAUTION *** *** ***
        '   This line is a real pain because an incorrect assingment causes a
        '   very vague error message.  You end up having to back trace each value
        '   to determine which value caused the error.
        ' BE CAREFUL WHEN CHANGING THE INSERT
        InsertCmd = "Insert into CORSON Values('" & tempLast & "','" & tempFirst & "','" & rs("MemberID") & "','','" & rs("MemberFed") & "','" & tempSex & "','" & rs("TourID") & "'," & Off_Score & ",NULL,NULL," & tempAltScore & ",'" & tempYOB & "','" & rs("Class") & "','" & rs("Round") & "','" & rs("Div") & "'," & tempPQ1 & "," & tempPQ2 & ",'" & rs("EndDate") & "','','Y','" & tempMiss & "'," & tempPlace & "," & tempConv & ",NULL,'" & rs("TSiteID") & "')"
      End If
    case "J"
      If tempPQ1 = 0.275 then tempPQ1 = 0.271
      If (tempSex = "M" and PreZBSScore >= 125.666) or (tempSex = "F" and PreZBSScore >= 91.106) then
        If tempAltScore > 0 Then
          Off_Score = rs("AltScore")
        Else
          Off_Score = PreZBSScore * .3048
        End If
        tempSpecial = ""
        If tempPQ1 = 0.239 then tempSpecial = "S"
        tempAltScore = PreZBSScore
       
        'Insert row To the table
        ' CAUTION *** *** ***
        '   This line is a real pain because an incorrect assingment causes a
        '   very vague error message.  You end up having to back trace each value
        '   to determine which value caused the error.
        ' BE CAREFUL WHEN CHANGING THE INSERT
        InsertCmd = "Insert into CORSON Values('" & tempLast & "','" & tempFirst & "','" & rs("MemberID") & "','','" & rs("MemberFed") & "','" & tempSex & "','" & rs("TourID") & "',NULL,NULL," & Off_Score & "," & tempAltScore & ",'" & tempYOB & "','" & rs("Class") & "','" & rs("Round") & "','" & rs("Div") & "'," & tempPQ1 & "," & tempPQ2 & ",'" & rs("EndDate") & "','" & tempSpecial & "','Y',''," & tempPlace & ",NULL,NULL,'" & rs("TSiteID") & "')"
      End If
    case "T"
      If (tempSex = "M" and PreZBSScore >= 4000) or (tempSex = "F" and PreZBSScore >= 3200) then
        'Insert row To the table
        ' CAUTION *** *** ***
        '   This line is a real pain because an incorrect assingment causes a
        '   very vague error message.  You end up having to back trace each value
        '   to determine which value caused the error.
        ' BE CAREFUL WHEN CHANGING THE INSERT
        InsertCmd = "Insert into CORSON Values('" & tempLast & "','" & tempFirst & "','" & rs("MemberID") & "','','" & rs("MemberFed") & "','" & tempSex & "','" & rs("TourID") & "',NULL," & PreZBSScore & ",NULL,NULL,'" & tempYOB & "','" & rs("Class") & "','" & rs("Round") & "','" & rs("Div") & "',NULL,NULL,'" & rs("EndDate") & "','','Y',''," & tempPlace & ",NULL,NULL,'" & rs("TSiteID") & "')"
      End If
    End Select
    ' The build above results in some null values. 
    ' We have to replace the blanks with the word null for DBase
    ' We do it three times in case there are multiple nulls next
    ' to one another.  Dumb, but it works.
    '
    If InsertCmd <> "" Then
      InsertCmd = replace(InsertCmd,",,",",NULL,")
      InsertCmd = replace(InsertCmd,",,",",NULL,")
      InsertCmd = replace(InsertCmd,",,",",NULL,")
'markdebug(insertcmd)  
      DBFCon.Execute InsertCmd
      RecordsSaved = RecordsSaved + 1
    End If

    rs.movenext
  loop

  rs.close
  set rs = nothing
  DBFCon.Close
  set DBFCon = nothing
  
  sSQL = "Update " & RawScoresTableName & " set CorsonDL = 1"
  sSQL = sSQL + " where [class] in ('R','L','A','B') and EndDate >= '" & SQLClean(Request("begindate")) & "' and EndDate <= '" & SQLClean(Request("enddate")) & "'"  
  If trim(Request("ToursToDownload")) <> "" Then sSQL = sSQL + " and TourID in (" & SQLClean(request("ToursToDownload")) & ")"
  Con.Execute(sSQL)

  CloseCon
  


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
  <%WriteIndexPageHeader
  
  If RecordsSaved > 0 Then
  %>
    <br><br>
    <center><h2>File export complete.</h2><br>
    <%=RecordsSaved%> records were exported.<br><br>
    <h4><a href="/rankings/defaultHQ.asp?process=corson2&rid=<%=rid%>">To download your file, click here.</a></h4>
   <%
   Else
   %>
    <br><br>
    <center><h2>File export complete.</h2><br><br><br>
    <h4>No new records were exported.</h4>
   <%
   End If
   WriteIndexPageFooter    
     
End If


%>




