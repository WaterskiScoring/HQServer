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

Dim tempLast, tempFirst, tempPlace, tempDiv, tempSL, tempTR, tempJU, tempScore, tempAlt, tempPQ1, tempPQ2, tempSex, tempYOB, TempSlmMiss, tempSpecial, tempIWSF, tempExport
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
      sSQL = sSQL & " where [class] in ('R','L','A','B') and EndDate >= '" & SQLClean(Request("begindate")) & "' and EndDate <= '" & SQLClean(Request("enddate")) & "'"  
      sSQL = sSQL & " order by TourID"

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

 
  ' Set up output text file.

  Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
  ExportFile=PathToTRA & "news\IWWF-Export.txt"
  Set objTextOut = objFSO.opentextfile(ExportFile,2,true)


  'Open Raw Scores Table
  OpenCon
  set rs = Server.CreateObject("ADODB.recordset")
  sSQL = "Select RS.FName, RS.LName, RS.MemberID, MT.FederationCode as MemberFed, RS.TourID, Convert(char(8),RS.EndDate,112) as EndDate,"
  sSQL = sSQL & " RS.Event, RS.Div, RS.Class, RS.Round, RS.Place, RS.Perf_Qual1, RS.Perf_Qual2, RS.AltScore,"
  sSQL = sSQL & " Case when RS.Event = 'S' then RS.Score-DT.ZBSConversion else RS.Score end as Score, MT.Sex, MT.BirthDate, ST.TSiteID"
  sSQL = sSQL & " from " & RawScoresTableName & " as RS left join " & MemberTableName & " as MT on RS.MemberID = MT.PersonIDwithCheckDigit"
  sSQL = sSQL & " left join " & SanctionTableName & " as ST on upper(left(RS.TourID,6)) = upper(left(ST.TSanction,6))"
  sSQL = sSQL & " left join " & SkiYearTableName & " as SY on RS.EndDate between SY.BeginDate and SY.EndDate and SY.SkiYearID <> 1"
  sSQL = sSQL & " left join " & DivisionsTableName & " as DT on RS.Div = DT.Div and SY.skiyearid = DT.skiyearid"
  sSQL = sSQL & " where RS.Class in ('R','L') and RS.EndDate >= '" & SQLClean(Request("begindate")) 
'  sSQL = sSQL & " where RS.Class in ('R','L','E') and RS.EndDate >= '" & SQLClean(Request("begindate")) 
  sSQL = sSQL & "' and RS.EndDate <= '" & SQLClean(Request("enddate")) & "'"  

'	************* Special for this run -- W5 & W6 Slalom scores only.
'	sSQL = sSQL & " and RS.Div in ('W5','W6') and RS.Event = 'S'"
'	************* Special for this run -- W5 & W6 Slalom scores only.

  If Request("PreviouslySelected") = "No" Then sSQL = sSQL & " and RS.CorsonDL IS NULL"
  If trim(Request("ToursToDownload")) <> "" Then sSQL = sSQL & " and RS.TourID in (" & SQLClean(request("ToursToDownload")) & ")"
	'  sSQL = sSQL & " order by case RS.[class] when 'R' then 1 when 'L' then 2 when 'A' then 3 when 'B' then 4 end, RS.memberid, RS.EndDate, RS.Round, RS.Event"
  sSQL = sSQL & " order by RS.memberid, RS.EndDate, RS.Round, RS.Event"
  WriteDebugSQL (sSQL)

  rs.open sSQL, SConnectionToTRATable, 3, 3

  do while not rs.eof

    tempSex = ucase(left(rs("sex"),1))
    tempYOB = right(rs("birthdate"),4)
    tempAge = right(trim(request("enddate")),4) - TempYOB
    If rs("MemberFed") = "USA" THEN tempIWSF = "USA" & rs("MemberID"): ELSE tempIWSF = ""

    '
    ' Formatting of Name fields.  Some names have a special character (apostrophe).  Replace it with a blank.
    '

    tempFirst = left(ucase(rs("FName")),1)
    tempFirst = tempFirst + mid(lcase(rs("FName")),2)
    tempFirst = replace(tempFirst,"'","")
    tempLast = replace(ucase(rs("LName")),"'","")

    tempPlace = ucase(trim(rs("Place")))
    IF right(tempPlace,1) = "T" THEN tempPlace = left(tempPlace,len(tempPlace)-1)
    IF tempPlace = "" THEN tempPlace = "999"

    tempDiv = ucase(rs("Div"))

    If rs("Perf_Qual1") = 0.239 THEN tempSpecial = "S": ELSE tempSpecial = ""
    If Instr(tempDiv,"B") > 0 or Instr(tempDiv,"G") > 0 then tempSpecial = "J"

    tempSlmMiss = ""
    tempScore = rs("Score")

    SELECT CASE rs("Event")

    CASE "S"

      tempAlt = FormatNumber(rs("AltScore"),2)
      tempPQ1 = FormatNumber(rs("Perf_Qual1")/100,2)
      tempPQ2 = rs("Perf_Qual2")
      If left(tempDiv,1) = "W" or tempDiv = "OW" or left(tempDiv,1) = "G" then tempScore = tempScore - 6
      tempSL = FormatNumber(tempScore,2)
      tempTR = ""
      tempJU = ""
      IF tempScore < 6 THEN tempSlmMiss = "Y": ELSE tempSlmMiss = "N"
      IF tempScore > 0 THEN tempExport = "Y": ELSE tempExport = "N"
      IF tempDiv = "B1" or tempDiv = "G1" or tempDiv = "B2" or tempDiv = "G2" THEN tempExport = "N"
      IF tempDiv = "M7" or tempDiv = "M8" or tempDiv = "M9" or tempDiv = "MA" or tempDiv = "MB" THEN tempExport = "N"
      IF tempDiv = "W7" or tempDiv = "W8" or tempDiv = "W9" or tempDiv = "WA" or tempDiv = "WB" THEN tempExport = "N"

    CASE "J"

      tempAlt = tempScore
      If rs("Perf_Qual1") = 0.275 THEN tempPQ1 = "0.271": ELSE tempPQ1 = FormatNumber(rs("Perf_Qual1"),3)
      tempPQ2 = rs("Perf_Qual2")
      tempSL = ""
      tempTR = ""
      tempJU = FormatNumber(rs("AltScore"),1)
      
      IF (tempSex = "M" and tempScore >= 60) or (tempSex = "F" and tempScore >= 45)  THEN tempExport = "Y": ELSE tempExport = "N"
      IF tempPQ1 < "0.235" THEN tempExport = "N"
      
    CASE "T"

      tempAlt = ""
      tempPQ1 = ""
      tempPQ2 = ""
      tempSL = ""
      tempTR = tempScore
      tempJU = ""

      IF (tempSex = "M" and tempScore >= 800) or (tempSex = "F" and tempScore >= 600)  THEN tempExport = "Y": ELSE tempExport = "N"

    END SELECT

    IF tempAge < 35 and rs("Class") = "E" THEN tempExport = "N"

    IF tempExport = "Y" THEN

      objTextOut.write ( tempLast & ";" )
      objTextOut.write ( tempFirst & ";" )
      objTextOut.write ( rs("MemberID") & ";;" )
      objTextOut.write ( rs("MemberFed") & ";" )
      objTextOut.write ( tempSex & ";" )
      objTextOut.write ( rs("TourID") & ";" )
      objTextOut.write ( tempSL & ";" )
      objTextOut.write ( tempTR & ";" )
      objTextOut.write ( tempJU & ";" )
      objTextOut.write ( tempAlt & ";" )
      objTextOut.write ( tempYOB & ";" )
      objTextOut.write ( rs("Class") & ";" )
      objTextOut.write ( trim(rs("Round")) & ";" )
      objTextOut.write ( tempDiv & ";" )
      objTextOut.write ( tempPQ1 & ";" )
      objTextOut.write ( tempPQ2 & ";" )
      objTextOut.write ( rs("EndDate") & ";" )
      objTextOut.write ( tempSpecial & ";" )
      objTextOut.write ( "Y;" )
      objTextOut.write ( tempSlmMiss & ";" )
      objTextOut.write ( tempPlace & ";" )
      objTextOut.write ( tempIWSF & ";" )
      objTextOut.writeline ( rs("TSiteID") )
      
      RecordsSaved = RecordsSaved + 1

    END IF

    rs.movenext

  loop

  rs.close
  set rs = nothing
  objTextOut.close
  set objTextOut = nothing
  set objFSO = nothing

  sSQL = "Update " & RawScoresTableName & " set CorsonDL = 1"
  sSQL = sSQL + " where [class] in ('R','L') and EndDate >= '" & SQLClean(Request("begindate")) & "' and EndDate <= '" & SQLClean(Request("enddate")) & "'"  
'  sSQL = sSQL + " where [class] in ('R','L','E') and EndDate >= '" & SQLClean(Request("begindate")) & "' and EndDate <= '" & SQLClean(Request("enddate")) & "'"  
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




