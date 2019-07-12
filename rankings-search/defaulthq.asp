<!--#include file="settingsHQ.asp"-->
<!--#include file="tools_include16.asp"-->
<title>Rankings Administration</title>



<html><head><title>TRA Web Administration</title></head><body>
<%

'Path for Mapped Files
Dim sMapPath
'File System Object
Dim objFSO	
'File Object
Dim objFile
'Folder Object
Dim objFolder
'SubFiles Object
Dim objFilesInFolder
' Authorities Stuff
Dim tsgi, usgi, amlvl


Session.Timeout = 30



' ---------------------------- LOG OUT --------------------------------



IF (trim(request("process")) = "logout") THEN
  IF (trim(Session("message")) <> "") THEN
    WriteIndexPageHeader
    %>
    <br><br><br><br>
    <font color="red"><center>
    <%Response.Write Session("message")%>
    </center></font>
    <br><br><br><br>
    <br><br><br><br>
    <%
    WriteIndexPageFooter
    Session.Abandon
  ELSE
    Session.Abandon
    Response.Redirect "/rankings/DefaultHQ.asp?rid=" & rid
  END IF
END IF


' --------- UPLOAD ANY -- ZIP OR OTHER SINGLE POST-TOURNAMENT REPORT FILE --------



IF (trim(request("process")) = "uploadany") and (session("UserLevel") > 9) THEN

' --- Tests the authority of this person to be in this module ---
IF Session("UserSptsGrpID")<>"AWS" THEN
	response.redirect("/rankings/tools.asp?svar=reject")
END IF
 
    WriteIndexPageHeader
    
    NewsPageNum = "1"

    %>

			<center>
			<font face="Verdana, Arial, Helvetica, sans-serif" Size="2">
			<p><b>Post Tournament Reporting -- Zip or Report Upload.</b></p>
			<p>Step 1:&nbsp; Click the <b>Browse</b> button below, then locate the<br>
				desired WSTIMS Zip or Report file you wish to upload.</p>
			<p>Step 2:&nbsp; Click the <b>Upload</b> button below to begin the <br> 
				Post-Tournament Reporting Upload process.</p>
	    <FORM method="post" encType="multipart/form-data" action="UploadZip.asp">
				<INPUT type="File" name="ZipFile" size="40">
				&nbsp;&nbsp;&nbsp;
				<INPUT type="Submit" value="Upload">
			</FORM>
			<p><font color="red"><b>Please Note:</b></font>&nbsp; There will be a 
				noticeable delay after you click <br>the <b>Upload</b> button, while
				the file is actually transferred across <br>the internet.&nbsp; The
				larger the size of the file, or the slower your <br>connection, the
				longer this delay will be.&nbsp; Please be patient :-)</p>
			</center>
			
    <%
    
    WriteIndexPageFooter
    
END IF




' ---------------------------- UNEXPECTED END OF FILE --------------------------------



IF (trim(request("process")) = "endoffile") THEN

    WriteIndexPageHeader
    
    %>
    <center>
    <br><br>
    <H2><font color="red">ERROR</font></H2>
    <br><br>
    An error has occured.  You have reached <br>
    the unexpected end of a file and you <br>
    were still looking for additional lines.<br><br>
    <br><br>

    Please report this error to your Regional <br>
    Seeding Committee member or Headquarters.
    <br><br><br><br><br><br>
    <%
    
    WriteIndexPageFooter
    
END IF


' ---------------------------- SET DEFAULT SKI YEAR --------------------------------



IF (trim(request("process")) = "defaultyear") THEN
  OpenCon
  set rs = Server.CreateObject("ADODB.recordset")
  sSQL = "SELECT top 8 * from " & SkiYearTableName & " WHERE skiyearid <> 1 order by EndDate desc"
  rs.open sSQL, SConnectionToTRATable, 3, 3  
  
  WriteIndexPageHeader
  
  NewsTitle="Set Default Ski Year"
  News="SELECT the Ski Year that you wish to set as default.  <br><br> This ski year will be used by the system when calculating standings, generating reports, and for all functions accessed by members."
  
  %>
  <br><br>
  <center><h2>Set Default Ski Year<br></h2>
  <br><br>    
  <FORM method="post" action="/rankings/DefaultHQ.asp">
  SELECT the default ski year:<br>
  <SELECT name="SkiYearID">
  
  <%
  DO WHILE NOT rs.eof %>
    <option 
    <%IF rs("DefaultYear") THEN Response.Write ("SELECTed ") %>
    value="<%=rs("SkiYearID")%>"><%=rs("SkiYearName") & " - " & rs("BeginDate") & " to " & rs("EndDate")%></option>
  <% rs.MoveNext
  LOOP
  %></SELECT>
  <br><br>
  <INPUT type="hidden" name="process" value="defaultyear2">
  <INPUT type="Submit" value="Go!">
  </FORM></center>
  <%
  rs.close
  set rs = Nothing
  CloseCon
  WriteIndexPageFooter
END IF  


IF (trim(request("process")) = "defaultyear2") THEN
  OpenCon

  '
  ' First we set the new default year
  '
  sSQL = "Update " & SkiYearTableName & " set DefaultYear = 1 WHERE SkiYearID = '" & SQLClean(request("SkiYearID")) & "'"
  con.execute(sSQL)

  '
  ' THEN we clear all the other default years
  '
  sSQL = "Update " & SkiYearTableName & " set DefaultYear = 0 WHERE SkiYearID <> '" & SQLClean(request("SkiYearID")) & "'"
  con.execute(sSQL)

  '
  ' THEN we set the SkiYear and PrevYearID in the Last 12 Mos Ski Year record to match
  '
  sSQL = "update SY1 set PrevYearID = SY2.PrevYearID, SkiYear = SY2.SkiYear from "
  sSQL = sSQL & SkiYearTableName & " SY1, " & SkiYearTableName & " SY2 "
  sSQL = sSQL & "where SY1.SkiYearID = 1 and SY2.DefaultYear = 1;"
  con.execute(sSQL)

  CloseCon
  WriteIndexPageHeader
  NewsTitle="Set Default Ski Year"
  News="SELECT the Ski Year that you wish to set as default.  <br><br> This ski year will be used by the system when calculating standings, generating reports, and for all functions accessed by members."
  %>
  <br><br>
  <center><h2>Default Ski Year Set<br></h2>
  <br><br>    
  The new default ski year has been saved.
  <%
  WriteIndexPageFooter
END IF  


' ---------------------------- Manual Recalc Process --------------------------------



IF (trim(request("process")) = "recalc") THEN
  OpenCon
  set rs = Server.CreateObject("ADODB.recordset")
'  sSQL = "SELECT top 8 * from " & SkiYearTableName & " WHERE DefaultYear <> 1 order by EndDate desc"
  sSQL = "SELECT top 8 * from " & SkiYearTableName & " order by EndDate desc"
  rs.open sSQL, SConnectionToTRATable, 3, 3  
  WriteIndexPageHeader
  NewsTitle="ReCalc Rankings"
  News="SELECT the Functions to be Performed, and the Ski Year to Recalculate rankings for.<br>All rankings will be immediately recalculated for the tournaments within this ski year."
  IF Session("adminmenulevel") >= 50 THEN 
    Response.write ("<br>sConnectionToTRATable = " + Application("sConnectionToTRATable") + "<br>&nbsp;")
    Response.write ("<br>HQSQLConn = " + Application("HQSQLConn") + "<br>&nbsp;")
    Response.write ("<br>HQOfficialsConn = " + Application("HQOfficialsConn") + "<br>&nbsp;")
    Response.write ("<br>SanUpdtCnn = " + Application("sConnectionToSanctionTable") + "<br>&nbsp;")
  END IF
  %>
  <center><font color=red size=4><b>SELECT Functions and Ski Year desired --</b></font><br>
  <FORM method="post" action="equival.asp">
  (1)&nbsp;&nbsp;SELECT the Specific Functions to be Performed:<br>&nbsp;<br>
    <input type="radio" name="Equival" value="MemUpd"> Update Member Extract Only&nbsp<br>
    <input type="radio" name="Equival" value="ReCalc"> Recalculate Rankings Only&nbsp;&nbsp;&nbsp;&nbsp;<br>
    <input checked type="radio" name="Equival" value="BothMR"> Both Upd Mbr Ext AND Recalc<br>&nbsp;<br>
  (2)&nbsp;&nbsp;SELECT the ski year to recalculate rankings for:<br>
  <br>
  <SELECT name="SkiYear">
  <%
  DO WHILE NOT rs.eof
    IF rs("SkiYearID") = "1" THEN 
      %> <option value="1">Last 12 Months ONLY</option> <%
      %> <option SELECTed value="">Curr Yr AND Last 12 Mo BOTH</option> <%
    ELSE
      %> <option value="<%=rs("SkiYearID")%>"><%=rs("SkiYearName") & " (" & rs("SkiYearID") & ") - " & rs("BeginDate") & " to " & rs("EndDate")%></option> <%
    END IF
    rs.MoveNext
  LOOP
  %></SELECT>
  <br>&nbsp;<br>&nbsp;<br>
  <INPUT type="Submit" value="Perform Selected Functions Now!">
  </FORM></center>
  <%
  rs.close
  set rs = Nothing
  CloseCon
  WriteIndexPageFooter
END IF  


' ---------------------------- Reset Recalculation Underway Flags --------------------------------



IF (trim(request("process")) = "resetrcuf") THEN
  OpenCon
  set rs = Server.CreateObject("ADODB.recordset")
  sSQL = "Update " & SkiYearTableName & " Set RecalcUnderway = 0"
  Con.Execute(sSQL)
  CloseCon
  set rs = nothing
  WriteIndexPageHeader

  %>
  <br><br>
  <center><h2>The Rankings Recalculation Underway<br>Indicator Flags have now been Reset.<br></h2>
  <br><br>    
  <%

  NewsTitle="ReCalc Rankings"
  News=""
  WriteIndexPageFooter
END IF  


' ---------------------------- ADD SKI YEAR --------------------------------





IF (trim(request("process")) = "addyear") THEN
  WriteIndexPageHeader
  NewsTitle="Add Ski Year"
  News="Enter the new Ski Year information. <br><br> No events in this ski year can begin before the beginning date. <br><br> No events in this ski year can end after the ending date."
  %>
  <br><br>
  <center><h2>Add Ski Year<br></h2>
  <br><br>    
  <FORM method="post" action="/rankings/DefaultHQ.asp">
  Enter a Beginning Date for the Ski Year:<br>
  <INPUT type="textbox" name="begindate"><br>
      <small><font color=gray>( mm/dd/yyyy )</font></small><br><br>
  Enter an Ending Date for the Ski Year:<br>
  <INPUT type="textbox" name="enddate"><br>
      <small><font color=gray>( mm/dd/yyyy )</font></small><br><br>
  <br>
  <INPUT type="hidden" name="process" value="addyear2">
  <INPUT type="Submit" value="Go!">
  </FORM></center>
  <%
  WriteIndexPageFooter
END IF  


IF (trim(request("process")) = "addyear2") THEN
  IF isDate(request("begindate")) and isDate(request("enddate")) THEN
    OpenCon
    set rs = Server.CreateObject("ADODB.recordset")
    sSQL = "SELECT top 1 * from " & SkiYearTableName & " order by EndDate desc"
    rs.open sSQL, SConnectionToTRATable, 3, 1
    
    IF cdate(request("begindate")) <= rs("EndDate") THEN
      WriteIndexPageHeader
      NewsTitle="Add Ski Year"
      News="Enter the new Ski Year information. <br><br> No tournaments in this ski year can begin before the beginning date. <br><br> No events in this ski year can end after the ending date."
      %>
      <br><br>
      <center><h2>Add Ski Year</h2>
      <br>    
      <font color=red>The new ski year must begin AFTER <%=rs("enddate")%>.</font><br><br><br>
      <FORM method="post" action="/rankings/DefaultHQ.asp">
      Enter a Beginning Date for the Ski Year:<br>
      <INPUT type="textbox" name="begindate"><br>
          <small><font color=gray>( mm/dd/yyyy )</font></small><br><br>
      Enter an Ending Date for the Ski Year:<br>
      <INPUT type="textbox" name="enddate"><br>
          <small><font color=gray>( mm/dd/yyyy )</font></small><br><br>
      <br>
      <INPUT type="hidden" name="process" value="addyear2">
      <INPUT type="Submit" value="Go!">
      </FORM></center>
      <%
      WriteIndexPageFooter
    END IF
    
    IF cdate(request("enddate")) <= cdate(request("begindate")) THEN
      WriteIndexPageHeader
      NewsTitle="Add Ski Year"
      News="Enter the new Ski Year information. <br><br> No tournaments in this ski year can begin before the beginning date. <br><br> No events in this ski year can end after the ending date."
      %>
      <br><br>
      <center><h2>Add Ski Year</h2>
      <br>    
      <font color=red>The end date must be GREATER THAN the begin date.</font><br><br><br>
      <FORM method="post" action="/rankings/DefaultHQ.asp">
      Enter a Beginning Date for the Ski Year:<br>
      <INPUT type="textbox" name="begindate"><br>
          <small><font color=gray>( mm/dd/yyyy )</font></small><br><br>
      Enter an Ending Date for the Ski Year:<br>
      <INPUT type="textbox" name="enddate"><br>
          <small><font color=gray>( mm/dd/yyyy )</font></small><br><br>
      <br>
      <INPUT type="hidden" name="process" value="addyear2">
      <INPUT type="Submit" value="Go!">
      </FORM></center>
      <%
      WriteIndexPageFooter
    END IF
     
    IF cdate(request("begindate")) > rs("EndDate") and cdate(request("enddate")) > cdate(request("begindate")) THEN
	      rs.close
			
				' --- Changed 2-20-2013 - By Mark Crone ---
      
    	  ' --- Determine the values from the largest record based on th value of SkiYear ---
      	sSQL = "SELECT TOP 1 SkiYearID, SkiYear FROM " & SkiYearTableName
      	sSQL = sSQL + " ORDER BY SkiYear DESC"  
      	set rs = Server.CreateObject("ADODB.recordset")
      	rs.open sSQL, SConnectionToTRATable
      
      	Dim NewPrevYearID, NewSkiYear
				IF NOT rs.eof THEN
						NewPrevYearID=rs("SkiYearID")
						NewSkiYear=rs("SkiYear")+1
				END IF
			   
      
      	sSQL = "INSERT into " & SkiYearTableName
      	sSQL = sSQL + " ([SkiYearName], [BeginDate], [EndDate], [DefaultYear], [LastRecalc], [RecalcUnderway], [SkiYear], [PrevYearID])"
      	sSQL = sSQL + " VALUES ("
      	sSQL = sSQL + "'Ski Year: " & year(SQLClean(request("EndDate"))) & "'"
      	sSQL = sSQL + ", '" & SQLClean(request("begindate")) & "'"
      	sSQL = sSQL + ", '" & SQLClean(request("enddate")) & "'"
      	sSQL = sSQL + ", '0','0','0'"
      	sSQL = sSQL + ", '"&NewSkiYear&"'"
      	sSQL = sSQL + ", '"&NewPrevYearID&"'"
      	sSQL = sSQL + ")"

				'response.write(sSQL)
    		'response.end
      
      	Con.Execute(sSQL)

      	rs.close
      	CloseCon
      	set rs = nothing

      	WriteIndexPageHeader
      	NewsTitle="Add Ski Year"
      	News="Enter the new Ski Year information. <br><br> No Tournaments in this ski year can begin before the beginning date. <br><br> No events in this ski year can end after the ending date."
      	%>
      	<br><br>
      	<center><h2>Ski Year Added</h2>
      	<br>    
      	<% Response.Write("Ski Year: " & year(Request("EndDate"))) %> has been added.<br>
      	<%
      	WriteIndexPageFooter
    ELSE
      	rs.close
      	CloseCon
      	set rs = nothing
    END IF

  ELSE
    WriteIndexPageHeader
    NewsTitle="Add Ski Year"
    News="Enter the new Ski Year information. <br><br> No tournaments in this ski year can begin before the beginning date. <br><br> No events in this ski year can end after the ending date."
    %>
    <br><br>
    <center><h2>Add Ski Year</h2>
    <br>    
    <font color=red>Invalid Date, Please Try Again</font><br><br><br>
    <FORM method="post" action="/rankings/DefaultHQ.asp">
    Enter a Beginning Date Range:<br>
    <INPUT type="textbox" name="begindate"><br>
    <small><font color=red>( mm/dd/yyyy )</font></small><br><br>
    Enter a Ending Date Range:<br>
    <INPUT type="textbox" name="enddate"><br>
    <small><font color=red>( mm/dd/yyyy )</font></small><br><br>
    <INPUT type="hidden" name="process" value="addyear2">
    <INPUT type="Submit" value="Go!">
    </FORM></center>
    <%
    WriteIndexPageFooter
  END IF
END IF




' ---------------------------- ORPHANS REPORT --------------------------------

IF (trim(request("process")) = "orphanreport") THEN
  WriteIndexPageHeader
  NewsTitle="Orphan Record Report"
  News="This report will display scores which have orphaned Member ID's or Tournament ID's. These scores will not be accurately reported by the system until they are correctly linked to their appropriate Member or Tour."
  ErrCount=0
  %>
  <TABLE BORDER="1" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#FFFFFF" width=60%>
  <TR>
    <TD align="Center" colspan=7>Orphan Report</td>
  </TR>
  <TR>
    <TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" FACE="courier" SIZE="1">&nbsp;</FONT></Center></TD>
    <TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" FACE="courier" SIZE="1">MemberID</FONT></Center></TD>
    <TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" FACE="courier" SIZE="1">Lastname</FONT></Center></TD>
    <TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" FACE="courier" SIZE="1">Firstname</FONT></Center></TD>
    <TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" FACE="courier" SIZE="1">Div</FONT></Center></TD>
    <TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" FACE="courier" SIZE="1">TourID</FONT></Center></TD>
    <TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" FACE="courier" SIZE="1">Date</FONT></Center></TD>
    <TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" FACE="courier" SIZE="1">Error</FONT></Center></TD>
  </TR>
  <%
  Set rs=Server.CreateObject("ADODB.recordset")
  OpenCon
  ' FIRST DISPLAY SCORES WITH BAD MEMBER IDs
  sSQL = "SELECT * from " & RawScoresTableName 
  sSQL = sSQL + " WHERE MemberID NOT IN (SELECT PersonIDwithCheckDigit from " & MemberTableName & ")"
  sSQL = sSQL + " and DateAdd(year,3,EndDate) >= GetDate() order by EndDate Desc, TourID Asc"
  rs.open sSQL, sConnectionToTRATable, 3, 1

  DO WHILE NOT rs.EOF
    ErrCount = ErrCount + 1
    %>
    <tr>
      <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><%=ErrCount%></FONT></TD>
      <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><%=rs("MemberID")%></FONT></TD>
      <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><%=rs("LName")%></FONT></TD>
      <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><%=rs("FName")%></FONT></TD>
      <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><%=rs("Div")%></FONT></TD>
      <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><%=rs("TourID")%></FONT></TD>
      <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><%=rs("EndDate")%></FONT></TD>
      <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0" nowrap><FONT COlOR="#000000" FACE="courier" SIZE="1">MemberID Not Found in MemberTrak.</FONT></TD>
    </tr>
    
    <% 
    rs.MoveNext
  LOOP 
  rs.close

  ' THEN DISPLAY SCORES WITH BAD TOUR IDs
  sSQL = "SELECT * from " & RawScoresTableName 
  sSQL = sSQL + " WHERE left(TourID,6) NOT IN (SELECT TournAppID from "&SanctionTableName&")"
  sSQL = sSQL + " order by MemberID"
  rs.open sSQL, sConnectionToTRATable, 3, 1

  DO WHILE NOT rs.EOF
    ErrCount = ErrCount + 1
    %>
    <tr>
      <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><%=ErrCount%></FONT></TD>
      <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><%=rs("memberid")%></FONT></TD>
      <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><%=rs("LName")%></FONT></TD>
      <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><%=rs("FName")%></FONT></TD>
      <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><%=rs("TourID")%></FONT></TD>
      <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><FONT COlOR="#000000" FACE="courier" SIZE="1"><%=rs("EndDate")%></FONT></TD>
      <TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0" nowrap><FONT COlOR="#000000" FACE="courier" SIZE="1">Tour ID Not Found in SWIFT.</FONT></TD>
    </tr>
    
    <% 
    rs.MoveNext
  LOOP 
  rs.close

  IF ErrCount = 0 THEN
    %>
    <TR>
      <TD align="Center" colspan=7>&nbsp;</td>
    </TR>
    <TR>
      <TD align="Center" colspan=7><font color=red>No Orphaned Records Found</font></td>
    </TR>
    <%
  END IF  
  %>
  </TABLE>
  <%
  Set rs = Nothing
  CloseCon
  WriteIndexPageFooter
END IF



' ---------------------------- CORSON EXPORT --------------------------------


IF (trim(request("process")) = "corson") and (session("UserLevel") > 49) THEN
  SELECT Case request("failure") 
  Case "1"
    WriteIndexPageHeader
    NewsTitle="IWSF Export"
    News="To download the IWSF Export, begin by entering a date range.  <br><br>IF you wish to download specific tournaments, check the check box and you will be able to SELECT which specific tournaments within your date range to download.  <br><br>The system remembers which records have been previously downloaded.  You can also choose to download only new scores or to download all scores.  <br><br>Scores matching the criteria SELECTed will be exported into the DBF file."
    %>
    <center><br>
    <font color=red><b>IWSF Rankings Data Export<br></b></font><br>    
    <br><font color=red>Invalid Date(s), Please Try Again</font><br>
    <FORM method="post" action="/rankings/corson.asp">
    <input type="radio" name="PreviouslySELECTed" value="No"> Export only score records <b>not</b> previously exported.&nbsp; <br><br>
    <input type="radio" name="PreviouslySELECTed" value="Yes"> Export <b>all</b> score records in a specified date range ...<br>
    Specify a Date Range: ( enter as mm/dd/yyyy )<br>    
    <INPUT type="textbox" name="begindate"> through <INPUT type="textbox" name="enddate"><br><br>
    <input type="checkbox" name="SELECTTours"> Check here to export only selected<br>TourIDs within the above date range.<br><br>
    <input type="radio" name="PreviouslySELECTed" value="ThisOne" checked> Export just one single TourID:&nbsp;&nbsp; <INPUT type="textbox" name="TourID">.<br>
    <INPUT type="hidden" name="process" value="corson"><br><br>
    <INPUT type="Submit" value="Create Export File">
    </FORM>
    <br>
    <FORM method="post" action="/rankings/corson.asp">
    <INPUT type="hidden" name="process" value="clearflags">
    <INPUT type="Submit" value="Clear Previously Downloaded Score Flags">
    </FORM>
    </center>
    <%
    WriteIndexPageFooter
  Case "2"
    WriteIndexPageHeader
    NewsTitle="IWSF Export"
    News="To download the IWSF Export, begin by entering a date range.  <br><br>IF you wish to download specific tournaments, check the check box and you will be able to SELECT which specific tournaments within your date range to download.  <br><br>The system remembers which records have been previously downloaded.  You can also choose to download only new scores or to download all scores.  <br><br>Scores matching the criteria SELECTed will be exported into the DBF file."
    %>
    <center><br>
    <font color=red><b>IWSF Rankings Data Export<br></b></font><br>    
    <br><font color=red>No Scores/Tournaments Found within specified Date Range</font><br>
    <FORM method="post" action="/rankings/corson.asp">
    <input type="radio" name="PreviouslySELECTed" value="No"> Export only score records <b>not</b> previously exported.&nbsp; <br><br>
    <input type="radio" name="PreviouslySELECTed" value="Yes"> Export <b>all</b> score records in a specified date range ...<br>
    Specify a Date Range: ( enter as mm/dd/yyyy )<br>    
    <INPUT type="textbox" name="begindate"> through <INPUT type="textbox" name="enddate"><br><br>
    <input type="checkbox" name="SELECTTours"> Check here to export only selected<br>TourIDs within the above date range.<br><br>
    <input type="radio" name="PreviouslySELECTed" value="ThisOne" checked> Export just one single TourID:&nbsp;&nbsp; <INPUT type="textbox" name="TourID">.<br>
    <INPUT type="hidden" name="process" value="corson"><br><br>
    <INPUT type="Submit" value="Create Export File">
    </FORM>
    <br>
    <FORM method="post" action="/rankings/corson.asp">
    <INPUT type="hidden" name="process" value="clearflags">
    <INPUT type="Submit" value="Clear Previously Downloaded Score Flags">
    </FORM>
    </center>
    <%
    WriteIndexPageFooter
  Case ELSE
    WriteIndexPageHeader
    NewsTitle="IWSF Export"
    News="To download the IWSF Export, begin by entering a date range.  <br><br>IF you wish to download specific tournaments, check the check box and you will be able to SELECT which specific tournaments within your date range to download.  <br><br>The system remembers which records have been previously downloaded.  You can also choose to download only new scores or to download all scores.  <br><br>Scores matching the criteria SELECTed will be exported into the DBF file."
    %>
    <center><br>
    <font color=red><b>IWSF Rankings Data Export<br></b></font><br>    
    <FORM method="post" action="/rankings/corson.asp">
    <input type="radio" name="PreviouslySELECTed" value="No"> Export only score records <b>not</b> previously exported.&nbsp; <br><br>
    <input type="radio" name="PreviouslySELECTed" value="Yes"> Export <b>all</b> score records in a specified date range ...<br>
    Specify a Date Range: ( enter as mm/dd/yyyy )<br>    
    <INPUT type="textbox" name="begindate"> through <INPUT type="textbox" name="enddate"><br><br>
    <input type="checkbox" name="SELECTTours"> Check here to export only selected<br>TourIDs within the above date range.<br><br>
    <input type="radio" name="PreviouslySELECTed" value="ThisOne" checked> Export just one single TourID:&nbsp;&nbsp; <INPUT type="textbox" name="TourID">.<br>
    <INPUT type="hidden" name="process" value="corson"><br><br>
    <INPUT type="Submit" value="Create Export File">
    </FORM>
    <br>
    <FORM method="post" action="/rankings/corson.asp">
    <INPUT type="hidden" name="process" value="clearflags">
    <INPUT type="Submit" value="Clear Previously Downloaded Score Flags">
    </FORM>
    </center>
    <%
    WriteIndexPageFooter
  End SELECT
END IF  


IF (trim(request("process")) = "corson2") and (session("UserLevel") > 49) THEN
  ' Now the file has been built and saved to c:\corson.dbf
  ' We should now send the file to the user.

  WriteLog(date() &"  "& time() &"  Corson.dbf downloaded.")

  WriteIndexPageHeader

  %>
  <br><br><br><br>
  <br><br><br><br>
  
  <a href="news/IWWF-Export.txt"><font face="Arial" size="2"><b>RIGHT 
         Click Here</b></font></a>&nbsp; <font size="2" face="Verdana, 
         Arial, Helvetica, sans-serif">to download the 
         IWWF Rankings Export Data File, then select the "Save As" 
         option from that menu, and then choose a suitable location 
         to store the download in your PC. </font>
  
  <br><br><br><br>
  <br><br><br><br>
  <%

  WriteIndexPageFooter
  
END IF


' ---------------------------- TRA DOS EXPORT --------------------------------


IF (trim(request("process")) = "trados") and (session("UserLevel") > 49) THEN

  sTourRegion = trim(Request("TRegion"))
  sMemberRegion = trim(Request("MRegion"))
  sTourClass = trim(Request("Class"))


  ' Gets latest 8 SKIYEAR records from SkiYearTableName
  OpenCon
  set rs = Server.CreateObject("ADODB.recordset")
  sSQL = "SELECT top 8 * from " & SkiYearTableName & " order by EndDate desc"
  rs.open sSQL, SConnectionToTRATable, 3, 3  


  SELECT Case request("failure") 
  Case "1"
    WriteIndexPageHeader
    NewsTitle="TRA Dos Export"
    News="To download the TRAdos Export, enter a date range.  All scores between the date range will be exported into the DBF file."
    %>

    <br><br>
    <center><h2>TRA-DOS Export<br></h2>
    <br>
    <FORM method="post" action="/rankings/trados.asp">
    Enter a Date Range:<br>
    <INPUT type="textbox" name="begindate"> through <INPUT type="textbox" name="enddate"><br>
    <font color=red> ( mm/dd/yyyy )</font><br><br>

    Tournament Region: <SELECT name="TRegion">
      <option value=""<%IF sTourRegion = "" THEN Response.Write(" SELECTed ")%>>All Regions</option>
      <option value="C"<%IF sTourRegion = "C" THEN Response.Write(" SELECTed ")%>>S. Central</option>
      <option value="M"<%IF sTourRegion = "M" THEN Response.Write(" SELECTed ")%>>Midwest</option>
      <option value="W"<%IF sTourRegion = "W" THEN Response.Write(" SELECTed ")%>>West</option>
      <option value="S"<%IF sTourRegion = "S" THEN Response.Write(" SELECTed ")%>>South</option>
      <option value="E"<%IF sTourRegion = "E" THEN Response.Write(" SELECTed ")%>>East</option>
    </SELECT><br>
    Member's Region: <SELECT name="MRegion">
      <option value=""<%IF sMemberRegion = "" THEN Response.Write(" SELECTed ")%>>All Regions</option>
      <option value="1"<%IF sMemberRegion = "1" THEN Response.Write(" SELECTed ")%>>S. Central</option>
      <option value="2"<%IF sMemberRegion = "2" THEN Response.Write(" SELECTed ")%>>Midwest</option>
      <option value="3"<%IF sMemberRegion = "3" THEN Response.Write(" SELECTed ")%>>West</option>
      <option value="4"<%IF sMemberRegion = "4" THEN Response.Write(" SELECTed ")%>>South</option>
      <option value="5"<%IF sMemberRegion = "5" THEN Response.Write(" SELECTed ")%>>East</option>
    </SELECT><br>
    Class: <SELECT name="Class">
      <option value=""<%IF sTourClass = "" THEN Response.Write(" SELECTed ")%>>All Classes</option>
      <option value="LR"<%IF sTourClass = "LR" THEN Response.Write(" SELECTed ")%>>L or R</option>
      <option value="ELR"<%IF sTourClass = "ELR" THEN Response.Write(" SELECTed ")%>>E or L or R</option>
      <option value="CELR"<%IF sTourClass = "CELR" THEN Response.Write(" SELECTed ")%>>C or E or L or R</option>
      <option value="F"<%IF sTourClass = "F" THEN Response.Write(" SELECTed ")%>>F or N or I</option>
    </SELECT><br><br>
    Ski Year:<br>
    <SELECT name="SkiYear">
      <option SELECTed value=""> &nbsp; </option>
    <%
    DO WHILE NOT rs.eof
      %> <option value="<%=rs("SkiYearID")%>"><%=rs("SkiYearName") & " - " & rs("BeginDate") & " to " & rs("EndDate")%></option> <%
      rs.MoveNext
    LOOP
    %></SELECT>
    <br><br>
    <input type="checkbox" name="EP"> EP and Open Ratings only.
    <font color=gray>(must SELECT a ski year)</font><br>
    <INPUT type="hidden" name="process" value="trados"><br><br>
    <INPUT type="Submit" value="Go!">
    </FORM></center>
    <%


  Case "2"
    WriteIndexPageHeader
    NewsTitle="TRA Dos Export"
    News="To download the TRAdos Export, enter a date range.  All scores between the date range will be exported into the DBF file."
    %>
    <br><br>
    <center><h2>TRA-DOS Export<br></h2>
    <br>
    <FORM method="post" action="/rankings/trados.asp">
    Enter a Date Range:<br>
    <INPUT type="textbox" name="begindate"> through <INPUT type="textbox" name="enddate"><br>
    <font color=gray> ( mm/dd/yyyy )</font><br><br>
    Tournament Region: <SELECT name="TRegion">
      <option value=""<%IF sTourRegion = "" THEN Response.Write(" SELECTed ")%>>All Regions</option>
      <option value="C"<%IF sTourRegion = "C" THEN Response.Write(" SELECTed ")%>>S. Central</option>
      <option value="M"<%IF sTourRegion = "M" THEN Response.Write(" SELECTed ")%>>Midwest</option>
      <option value="W"<%IF sTourRegion = "W" THEN Response.Write(" SELECTed ")%>>West</option>
      <option value="S"<%IF sTourRegion = "S" THEN Response.Write(" SELECTed ")%>>South</option>
      <option value="E"<%IF sTourRegion = "E" THEN Response.Write(" SELECTed ")%>>East</option>
    </SELECT><br>
    Member's Region: <SELECT name="MRegion">
      <option value=""<%IF sMemberRegion = "" THEN Response.Write(" SELECTed ")%>>All Regions</option>
      <option value="1"<%IF sMemberRegion = "1" THEN Response.Write(" SELECTed ")%>>S. Central</option>
      <option value="2"<%IF sMemberRegion = "2" THEN Response.Write(" SELECTed ")%>>Midwest</option>
      <option value="3"<%IF sMemberRegion = "3" THEN Response.Write(" SELECTed ")%>>West</option>
      <option value="4"<%IF sMemberRegion = "4" THEN Response.Write(" SELECTed ")%>>South</option>
      <option value="5"<%IF sMemberRegion = "5" THEN Response.Write(" SELECTed ")%>>East</option>
    </SELECT><br>
    Class: <SELECT name="Class">
      <option value=""<%IF sTourClass = "" THEN Response.Write(" SELECTed ")%>>All Classes</option>
      <option value="LR"<%IF sTourClass = "LR" THEN Response.Write(" SELECTed ")%>>L or R</option>
      <option value="ELR"<%IF sTourClass = "ELR" THEN Response.Write(" SELECTed ")%>>E or L or R</option>
      <option value="CELR"<%IF sTourClass = "CELR" THEN Response.Write(" SELECTed ")%>>C or E or L or R</option>
      <option value="F"<%IF sTourClass = "F" THEN Response.Write(" SELECTed ")%>>F or N or I</option>
    </SELECT><br><br>
    Ski Year:<br>
    <SELECT name="SkiYear">
      <option SELECTed value=""> &nbsp; </option>
    <%
    DO WHILE NOT rs.eof
      %> <option value="<%=rs("SkiYearID")%>"><%=rs("SkiYearName") & " - " & rs("BeginDate") & " to " & rs("EndDate")%></option> <%
      rs.MoveNext
    LOOP
    %></SELECT>
    <br><br>
    <input type="checkbox" name="EP"> EP and Open Ratings only. 
    <font color=red>(must SELECT a ski year)</font><br>
    <INPUT type="hidden" name="process" value="trados"><br><br>
    <INPUT type="Submit" value="Go!">
    </FORM></center>
    <%


  Case ELSE
    WriteIndexPageHeader
    NewsTitle="TRA Dos Export"
    News="To download the TRAdos Export, enter a date range.  All scores between the date range will be exported into the DBF file."
    %>
    <br><br>
    <center><h2>TRA-DOS Export<br></h2>
    <br>
    <FORM method="post" action="/rankings/trados.asp">
    Enter a Date Range:<br>
    <INPUT type="textbox" name="begindate"> through <INPUT type="textbox" name="enddate"><br>
    <font color=gray> ( mm/dd/yyyy )</font><br><br>
    Tournament Region: <SELECT name="TRegion">
      <option value=""<%IF sTourRegion = "" THEN Response.Write(" SELECTed ")%>>All Regions</option>
      <option value="C"<%IF sTourRegion = "C" THEN Response.Write(" SELECTed ")%>>S. Central</option>
      <option value="M"<%IF sTourRegion = "M" THEN Response.Write(" SELECTed ")%>>Midwest</option>
      <option value="W"<%IF sTourRegion = "W" THEN Response.Write(" SELECTed ")%>>West</option>
      <option value="S"<%IF sTourRegion = "S" THEN Response.Write(" SELECTed ")%>>South</option>
      <option value="E"<%IF sTourRegion = "E" THEN Response.Write(" SELECTed ")%>>East</option>
    </SELECT><br>
    Member's Region: <SELECT name="MRegion">
      <option value=""<%IF sMemberRegion = "" THEN Response.Write(" SELECTed ")%>>All Regions</option>
      <option value="1"<%IF sMemberRegion = "1" THEN Response.Write(" SELECTed ")%>>S. Central</option>
      <option value="2"<%IF sMemberRegion = "2" THEN Response.Write(" SELECTed ")%>>Midwest</option>
      <option value="3"<%IF sMemberRegion = "3" THEN Response.Write(" SELECTed ")%>>West</option>
      <option value="4"<%IF sMemberRegion = "4" THEN Response.Write(" SELECTed ")%>>South</option>
      <option value="5"<%IF sMemberRegion = "5" THEN Response.Write(" SELECTed ")%>>East</option>
    </SELECT><br>
    Class: <SELECT name="Class">
      <option value=""<%IF sTourClass = "" THEN Response.Write(" SELECTed ")%>>All Classes</option>
      <option value="LR"<%IF sTourClass = "LR" THEN Response.Write(" SELECTed ")%>>L or R</option>
      <option value="ELR"<%IF sTourClass = "ELR" THEN Response.Write(" SELECTed ")%>>E or L or R</option>
      <option value="CELR"<%IF sTourClass = "CELR" THEN Response.Write(" SELECTed ")%>>C or E or L or R</option>
      <option value="F"<%IF sTourClass = "F" THEN Response.Write(" SELECTed ")%>>F or N or I</option>
    </SELECT><br><br>
    Ski Year:<br>
    <SELECT name="SkiYear">
      <option SELECTed value=""> &nbsp; </option>
    <%
    DO WHILE NOT rs.eof
      %> <option value="<%=rs("SkiYearID")%>"><%=rs("SkiYearName") & " - " & rs("BeginDate") & " to " & rs("EndDate")%></option> <%
      rs.MoveNext
    LOOP
    %></SELECT>
    <br><br>
    <input type="checkbox" name="EP"> EP and Open Ratings only.<br>
    <font color=gray>(must SELECT a ski year)</font><br>
    <INPUT type="hidden" name="process" value="trados"><br><br>
    <INPUT type="Submit" value="Go!">
    </FORM></center>
    <%
  End SELECT


  WriteIndexPageFooter
  rs.close
  set rs = Nothing
  CloseCon
END IF

IF (trim(request("process")) = "trados2") and (session("UserLevel") > 49) THEN
  ' Now the three files have been built and saved to c:\
  ' Filenames are skiscore.dbf, member.dbf, and tourname.dbf
  ' We should now send the files to the user.
  Response.Buffer = True

  WriteLog(date() &"  "& time() &"  TRADOS Export downloaded.")

  Response.Clear
	'PWS for win98 has MDAC 1.5 - will not work
	'use MDAC 2.5 and above - will probably work okay on your server.
	'IIS on Win NT/2000 will normally have the proper MDAC
	'NOTE:IF you forget to set response.buffer=true, download will not work.

' SEND SKISCORE.DBF

  Set objStream = Server.CreateObject("ADODB.Stream")
    objStream.Open
    objStream.Type = 1
    objStream.LoadFromFile (PathtoTRA & "news\SKISCORE.DBF")
      ContentType = "application/octet-stream"
      Response.AddHeader "Content-Disposition", "attachment; filename=SKISCORE.DBF"
    ' Response.AddHeader "Content-Length", sFileSize
    ' In a Perfect World, Your Client would also have UTF-8 as the default 
    ' In Their Browser
      Response.Charset = "UTF-8"
      Response.ContentType = ContentType
      Response.BinaryWrite objStream.Read
      Response.Flush
    objStream.Close
  set objStream = Nothing
END IF

IF (trim(request("process")) = "trados3") and (session("UserLevel") > 49) THEN
  ' Now the three files have been built and saved to c:\
  ' Filenames are skiscore.dbf, member.dbf, and tourname.dbf
  ' We should now send the files to the user.
  Response.Buffer = True

  WriteLog(date() &"  "& time() &"  TRADOS Export downloaded.")

  Response.Clear
	'PWS for win98 has MDAC 1.5 - will not work
	'use MDAC 2.5 and above - will probably work okay on your server.
	'IIS on Win NT/2000 will normally have the proper MDAC
	'NOTE:IF you forget to set response.buffer=true, download will not work.

' SEND MEMBER.DBF
  
  Set objStream = Server.CreateObject("ADODB.Stream")
    objStream.Open
    objStream.Type = 1
    objStream.LoadFromFile (PathtoTRA & "news\MEMBER.DBF")
      ContentType = "application/octet-stream"
      Response.AddHeader "Content-Disposition", "attachment; filename=MEMBER.DBF"
    ' Response.AddHeader "Content-Length", sFileSize
    ' In a Perfect World, Your Client would also have UTF-8 as the default 
    ' In Their Browser
      Response.Charset = "UTF-8"
      Response.ContentType = ContentType
      Response.BinaryWrite objStream.Read
      Response.Flush
    objStream.Close
  set objStream = Nothing
END IF

IF (trim(request("process")) = "trados4") and (session("UserLevel") > 49) THEN
  ' Now the three files have been built and saved to c:\
  ' Filenames are skiscore.dbf, member.dbf, and tourname.dbf
  ' We should now send the files to the user.
  Response.Buffer = True

  WriteLog(date() &"  "& time() &"  TRADOS Export downloaded.")

  Response.Clear
	'PWS for win98 has MDAC 1.5 - will not work
	'use MDAC 2.5 and above - will probably work okay on your server.
	'IIS on Win NT/2000 will normally have the proper MDAC
	'NOTE:IF you forget to set response.buffer=true, download will not work.

' SEND TOURNAME.DBF

  Set objStream = Server.CreateObject("ADODB.Stream")
    objStream.Open
    objStream.Type = 1
    objStream.LoadFromFile (PathtoTRA & "news\TOURNAME.DBF")
      ContentType = "application/octet-stream"
      Response.AddHeader "Content-Disposition", "attachment; filename=TOURNAME.DBF"
    ' Response.AddHeader "Content-Length", sFileSize
    ' In a Perfect World, Your Client would also have UTF-8 as the default 
    ' In Their Browser
      Response.Charset = "UTF-8"
      Response.ContentType = ContentType
      Response.BinaryWrite objStream.Read
      Response.Flush
    objStream.Close
  set objStream = Nothing
END IF


IF (trim(request("process")) = "trados5") and (session("UserLevel") > 49) THEN

'markdebug(RankPath & "\news\SKISCORE.DBF")
  set objStream = CreateObject("Scripting.FileSystemObject")
    objStream.DeleteFile(PathtoTRA & "news\SKISCORE.DBF")  
    objStream.DeleteFile(PathtoTRA & "news\MEMBER.DBF")  
    objStream.DeleteFile(PathtoTRA & "news\TOURNAME.DBF")  
  Set objStream = Nothing

  WriteIndexPageHeader
  %>
  <br><br><br><br>
  <br><br><br><br>
  <br><br><br><br>
  <br><br><br><br>
  <%
  WriteIndexPageFooter
  
END IF

' ---------------------------- NEWS EDITOR --------------------------------



IF (trim(request("process")) = "news") and (session("UserLevel") > 49) THEN

    WriteIndexPageHeader
    
    %>
    <center>
    <br><br>
    <H2>NEWS</H2>
    
    Enter the information that you would like<br>
    displayed in the news sidebar.<br>
    <br><br>

    <form method=post action="/rankings/DefaultHQ.asp">
    <input type="hidden" name="process" value="savenews">
    <input type="hidden" name="page" value="<%=Request("page")%>">
    <textarea name=newsdata rows=10 cols=75 style="overflow:hidden">
<%
    Set objfso = CreateObject("Scripting.FileSystemObject")
    IF objFSO.FileExists(PathToNews & "\news-" & request("page") & ".txt") THEN
      set objstream=objFSO.opentextfile(PathToNews & "\news-" & request("page") & ".txt")
    ELSE
      set objstream=objFSO.opentextfile(PathToNews & "\news.txt")
    END IF
    IF objstream.atendofstream THEN
      response.write("No News Today.")
    ELSE
      DO WHILE not objstream.atendofstream
        response.write(objstream.readline)
        response.write(chr(10))
      LOOP
    END IF
    objstream.close  
%>
    </textarea><br>
    <br><center><input type="submit" value="Save News"></center>
    </form>

    <br><br><br><br><br><br>
    <%
    
    WriteIndexPageFooter
    
END IF



IF (trim(request("process")) = "savenews") and (session("UserLevel") > 49) THEN
    

  Set objfso = CreateObject("Scripting.FileSystemObject")
  IF objFSO.FileExists(PathToNews & "\news-" & request("page") & ".txt") THEN
    set objstream=objFSO.opentextfile(PathToNews & "\news-" & request("page") & ".txt",2,true)
  ELSE
    set objstream=objFSO.opentextfile(PathToNews & "\news.txt",2,true)
  END IF
  IF trim(request.form("newsdata")) = "" THEN
    objstream.write("No News Today")
  ELSE
    objstream.write(request.form("newsdata"))
  END IF
  objstream.close

  Response.Redirect "/rankings/DefaultHQ.asp?rid=" & rid

END IF


' --------------- DISPLAY WEBSITE TRAFFIC STATS REPORT --------------------


IF (trim(request("process")) = "traffic") and (session("UserLevel") > 9) THEN

WriteIndexPageHeader_NoMenu

%> <center><br><h2>Rankings Website Traffic Activity</h2> <%

Dim DTStart, DTEnd, DateRaw, DateFmt, I1, I2

IF len(request("DTEnd")) = 0 THEN
   DateRaw = Date(): I1 = instr(DateRaw,"/"): I2 = instr(I1+1,DateRaw,"/")
   DateFmt = Mid(DateRaw,I2+1): ' Start with Year value
   IF I1=2 THEN DateFmt = DateFmt + "-0" + Left(DateRaw,1): ELSE DateFmt = DateFmt + "-" + Left(DateRaw,2)
   IF I2-I1=2 THEN DateFmt = DateFmt + "-0" + Mid(DateRaw,I1+1,1): ELSE DateFmt = DateFmt + "-" + Mid(DateRaw,I1+1,2)
   DTEnd = DateFmt
ELSE
   DTEnd = SQLClean(trim(request("DTEnd")))
END IF

IF len(request("DTStart")) = 0 THEN
   IF Mid(DTEnd,6,2) > "01" THEN
      DTStart = Left(DTEnd,5) + Right(CStr(Cint(Mid(DTEnd,6,2))+99),2) + "-" + Right(CStr(Cint(Mid(DTEnd,9,2))+101),2)
   ELSE
      DTStart = Right(CStr(CInt(Mid(DTEnd,1,4))-1),4) + "-12-" + Right(CStr(Cint(Mid(DTEnd,9,2))+101),2)
   END IF
ELSE
   DTStart = SQLClean(trim(request("DTStart")))
END IF

%> <table align="Center" Cellpadding="0" Cellspacing="5">
	<tr align="Center">   
		<td><FONT size="1" face="Verdana, Arial, Helvetica, sans-serif"> &nbsp;&nbsp;&nbsp;Date 
			Range Start&nbsp;&nbsp;&nbsp;&nbsp; <br> (as YYYY-MM-DD) </font></td> 
		<td><FONT size="1" face="Verdana, Arial, Helvetica, sans-serif"> &nbsp;&nbsp;&nbsp;Date Range 
			End&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <br> &nbsp;&nbsp;&nbsp;(as YYYY-MM-DD)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font></td> 
		<td>&nbsp;</td>
	</tr>
		
	<tr align="Center"><form action="/rankings/DefaultHQ.asp?process=traffic" method="post">  
		<td><input name="DTStart" type="text" size=10 id="DTStart" value="<%= Response.Write(DTStart) %>"></td>
		<td><input name="DTEnd" type="text" size=10 id="DTEnd" value="<%= Response.Write(DTEnd) %>"></td>
		<td><input type="submit" style="width:9em" value="Display Activity"
			 title="Create a new report based on the range of Dates specified in the boxes to the left"></td></form>
		
   <form action="/rankings/DefaultHQ.asp" method="post">
     <td align=center>&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" style="width:9em" value="Back to Menu"
	  	title="Take me back to the Rankings Administration Index Menu">
	    </td></form>

	</tr>
	
	<tr><td>&nbsp;</td></tr>
	
</table> <%

SET rs=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT * FROM "&TrafficTableName&" Where ActivityDate Between '"&DTStart&"' and '"&DTEnd&"' Order by ActivityDate Desc"
rs.open sSQL, sConnectionToTRATable, 3, 1

IF NOT rs.eof THEN

	rs.movefirst  %>
	<table align="Center" BORDER="1" >
	<tr align="Center"><%
	  FOR i = 0 TO rs.fields.count - 1
		%><TD Align="Center" vAlign="top" nowrap><FONT COlOR="#000000" size="1">&nbsp;<%
		Response.Write(trim(Rs.Fields(i).Name)) 
	  %>&nbsp;</FONT></TD><%
	  NEXT%>
	</tr><%

	  DO WHILE NOT rs.eof
		rowCount = rowCount + 1
		%><tr align="Center"><%
		FOR i = 0 TO rs.fields.count - 1
			%><TD Align="Center" vAlign="top" nowrap><FONT COlOR="#000000" size="2">&nbsp;<%
			Response.Write(trim(Rs.Fields(i).Value))
			%>&nbsp;</FONT></TD><%
		NEXT%>
		</TR><%
		rs.movenext
	  LOOP  %>
	</tr>
</table><% 

	rs.close
	Set rs = nothing

	ELSE
		response.write("No data found in table for specified dates.")
	END IF

WriteIndexPageFooter

END IF





' ---------------------------- DELETE FILES --------------------------


IF (trim(request("process")) = "delete") and (session("UserLevel") > 9) THEN





WriteIndexPageHeader

  IF lcase(Request("confirm")) = "yes" THEN
    Set tempFSO = CreateObject("Scripting.FileSystemObject")
    'error handling here to make sure that things go smoothly
    'IF the file exists 
    IF tempFSO.FileExists(PathToTRA & request("file")) THEN


	' --- First check SWIFT to determine the SptsGrpID ---
	sSQL = "SELECT TOP 1 * FROM "&SanctionsTableName
	sSQL = sSQL + " WHERE LEFT(TSanction,6) = '"&LEFT(Request("file"),6)&"'" 
	Set rsTour=Server.CreateObject("ADODB.recordset")
	rsTour.open sSQL, SConnectionToSanctionTable, 3, 1

	' --- Tests the SD authority of this person to this tournament ---
	' --- Note that revised logic allows AWS users to act for NCW ---
	tsgi = rsTour("SptsGrpID"): usgi = Session("UserSptsGrpID"): amlvl = Session("adminmenulevel")
	IF (tsgi<>usgi AND (NOT tsgi="NCW" AND usgi="AWS")) AND amlvl<50 THEN
		Session("SptsGrpID") = tsgi
		rsTour=nothing
		response.redirect("/rankings/tools.asp?svar=reject")
	END IF


      'THEN we delete it
      tempFSO.DeleteFile(PathToTRA & request("file"))
      'IF this was an exceptions file, also wipe out the corresponding reasons file.
      IF lcase(left(request("file"),10)) = "exceptions" THEN
        tempFSO.DeleteFile(PathToTRA & "reasons\" & right(request("file"),len(request("file")) - instr(request("file"), "\")))
      END IF
      WriteLog(date() &"  "& time() &"  "& request("file") & " deleted.")
    END IF
    'this line destroys the instance of the File Scripting Object
    SET tempFSO = NOTHING
    %>   
    
    <h2> File <% = Request("File") %> has been deleted.</h2>
    <br><br>
    <a href="<%
    IF request("returnurl") <> "" THEN 
      response.write(request("returnurl") & "&rid=" & rid) 
    ELSE 
      response.write("/rankings/DefaultHQ.asp?rid=" & rid) 
    END IF
    %>">Click Here To Return</a>

<%END IF
  IF lcase(Request("confirm")) = "" THEN
%>  
    <br><br><h3>
    Type the word "YES" IF you are sure you wish to delete the file<br>
    <%= Request("file")%>.</h3>
    <br><br>
    <form action="/rankings/DefaultHQ.asp" method="post"> 
    <input type="hidden" name="process" value="delete">
    <input type="hidden" name="file" value="<%=Request("file")%>">
    <input type="hidden" name="returnurl" value="<%=Request("returnurl")%>">
    <input type="text" name="confirm" size="5">
    <input type="submit" value="Confirm Deletion?">
    </form>
<%END IF
  IF lcase(Request("confirm")) <> "yes" and lcase(Request("confirm")) <> "" THEN
     %>  <br><br>
         The file was NOT deleted.
         <br><br>
    <a href="<%
    IF request("returnurl") <> "" THEN 
      response.write(request("returnurl") & "&rid=" & rid) 
    ELSE 
      Response.Write("/rankings/DefaultHQ.asp?rid=" & rid) 
    END IF
    %>">Click Here To Return</a> 
<%END IF

WriteIndexPageFooter

END IF

' ---------------------------- DOWNLOAD FILES --------------------------

IF (trim(request("process")) = "download") and (session("UserLevel") > 9) THEN
	
    Response.Buffer = true
    dim filespec
    dim sFileName
    
    filespec = PathToTRA & request("file")
    sfilename = right(request("file"),len(request("file")) - instr(request("file"), "\"))

   WriteLog(date() &"  "& time() &"  "& filespec & " downloaded.")

	Response.Clear
	'PWS for win98 has MDAC 1.5 - will not work
	'use MDAC 2.5 and above - will probably work okay on your server.
	'IIS on Win NT/2000 will normally have the proper MDAC
	'NOTE:IF you forget to set response.buffer=true, download will not work.
	Set objStream = Server.CreateObject("ADODB.Stream")
	objStream.Open
	objStream.Type = 1
	objStream.LoadFromFile filespec

	sFileType = lcase(Right(request("file"), 4))
	    
	    SELECT Case sFileType
	        Case ".asf"
	            ContentType = "video/x-ms-asf"
	        Case ".avi"
	            ContentType = "video/avi"
	        Case ".doc"
	            ContentType = "application/msword"
	        Case ".zip"
	            ContentType = "application/zip"
	        Case ".xls"
	            ContentType = "application/vnd.ms-excel"
	        Case ".gif"
	            ContentType = "image/gif"
	        Case ".jpg", "jpeg"
	            ContentType = "image/jpeg"
	        Case ".wav"
	            ContentType = "audio/wav"
	        Case ".mp3"
	            ContentType = "audio/mpeg3"
	        Case ".mpg", "mpeg"
	            ContentType = "video/mpeg"
	        Case ".rtf"
	            ContentType = "application/rtf"
	        Case ".htm", "html"
	            ContentType = "text/html"
	        'Case ".asp"
	            'ContentType = "text/asp"
	        Case ELSE
	            'Handle All Other Files
	            ContentType = "application/octet-stream"
	    End SELECT
		
		
		Response.AddHeader "Content-Disposition", "attachment; filename=" & sFileName
		' Response.AddHeader "Content-Length", sFileSize
		' In a Perfect World, Your Client would also have UTF-8 as the default 
		' In Their Browser
		Response.Charset = "UTF-8"
		Response.ContentType = ContentType
		
		Response.BinaryWrite objStream.Read
		Response.Flush

	objStream.Close
	Set objStream = Nothing

END IF

' ---------------------------- EXCEPTION MANAGEMENT --------------------------------

Dim BadToursTotal, BadRecordsTotal

BadRecordsTotal = 0


IF (trim(request("process")) = "badscores") and (session("userlevel") > 9) THEN
    RegionSELECTed = request("region") 
    YearSELECTed = request("year")

' IF it's blank, THEN we use the Session value.
    IF RegionSELECTed = "" THEN RegionSELECTed = trim(session("RegionSELECTed"))
' IF it's still blank, THEN we use "??"
    IF RegionSELECTed = "" THEN RegionSELECTed = "??"

' IF it's blank, THEN we use the Session value.
    IF YearSELECTed = "" THEN YearSELECTed = trim(session("YearSELECTed"))
' IF it's still blank, THEN we use current calendar year
    IF YearSELECTed = "" THEN YearSELECTed = right(FormatDateTime(Date,2),2)

    session("RegionSELECTed") = RegionSELECTed
    session("YearSELECTed") = YearSELECTed

    sMapPath = PathToTRA & "exceptions\"
    WriteIndexPageHeader
    NewsPageNum = "2"
    %>

    <center>
    <br><br>
    You are currently viewing uploaded score files which have exceptions remaining.
    <br>
    </center>

    <form method=post action="/rankings/DefaultHQ.asp">
    <input type="hidden" name="process" value="badscores">
    <table width=90% align=center>
    <tr><td>
 
    Region: 

    <SELECT name="region">
      <option value ="??"<%IF RegionSELECTed = "??" THEN Response.Write(" SELECTed ")%>>Please Select</Option><br>
      <option value ="All"<%IF RegionSELECTed = "All" THEN Response.Write(" SELECTed ")%>>All Regions</Option><br>
      <option value ="C"<%IF RegionSELECTed = "C" THEN Response.Write(" SELECTed ")%>>S. Central</Option><br>
      <option value ="M"<%IF RegionSELECTed = "M" THEN Response.Write(" SELECTed ")%>>MidWest</Option><br>
      <option value ="W"<%IF RegionSELECTed = "W" THEN Response.Write(" SELECTed ")%>>West</Option><br>
      <option value ="S"<%IF RegionSELECTed = "S" THEN Response.Write(" SELECTed ")%>>South</Option><br>
      <option value ="E"<%IF RegionSELECTed = "E" THEN Response.Write(" SELECTed ")%>>East</Option><br>
      <option value ="U"<%IF RegionSELECTed = "U" THEN Response.Write(" SELECTed ")%>>All NCWSA</Option><br>
    </SELECT>
    </td><td align=center>
 
    Ski Year:

    <%
	BuildYearDrop
    %>


 <td ALIGN="center" vAlign="top">
 </td>

    </td></tr></table>
    <center><input type=submit value="Filter Results"></center>
    </form>

    <hr>

    <TABLE class="innertable" width=90% align=center>
    <TR align="left" valign="top" bgcolor="#000000"> 
  
    <TH width="15%"><FONT color="#FFFFFF"><center><B><FONT size="2" face="Verdana, Arial, Helvetica, sans-serif">File Name</FONT></B></center></FONT></TH>
    <TH width="9%"><FONT color="#FFFFFF"><center><B><FONT size="2" face="Verdana, Arial, Helvetica, sans-serif">Records</FONT></B></center></FONT></TH>
    <TH width="12%"><FONT color="#FFFFFF"><center><B><FONT size="2" face="Verdana, Arial, Helvetica, sans-serif">Actions</FONT></B></center></FONT></TH>
    <TH width="10%"><FONT color="#FFFFFF"><center><B><FONT size="2" face="Verdana, Arial, Helvetica, sans-serif">Size</FONT></B></center></FONT></TH>
    <TH width="20%"><FONT color="#FFFFFF"><center><B><FONT size="2" face="Verdana, Arial, Helvetica, sans-serif">Date</FONT></B></center></FONT></TH>
    </TR>
    <br><br>
    <%

    	'Create File System Object To Get list of files
    	Set objFSO = CreateObject("Scripting.FileSystemObject")
    	'Get The path For the web page and its dir.
    	'Set the object folder To the mapped path
    	Set objFolder = objFSO.GetFolder(sMapPath)
    	set objFilesInFolder = objFolder.Files
        ListingsCount = 0
    	'For Each file In the folder

        IF objFilesInFolder.Count <> 0 AND RegionSelected <> "??" THEN
          For Each objFile In objFolder.Files
    	    IF (RegionSELECTed = "All" or RegionSELECTed = ucase(right(left(objFile.Name,14),1))) and (YearSELECTed = "All" or YearSELECTed = right(left(objFile.Name,13),2)) THEN
              ListingsCount = ListingsCount + 1
              %> 
    	      <TR align="left" valign="top" bordercolor="#999999" bgcolor="#FFFFFF"> 
    	        <TD><center> 
    	          <FONT size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">
                  <%
                  IF ucase(right(objfile.name,3)) = "WSP" THEN
                  %> <a href="/rankings/exceptionmgmt-wsp.asp?file=<%Response.Write(objfile.name)%>&line=2&rid=<%
                  ELSE
                  %> <a href="/rankings/exceptionmgmt-pdf.asp?file=<%Response.Write(objfile.name)%>&line=1&rid=<%
                  END IF
                    
                    Response.Write (rid)

                    OpenCon
                    Set rs=Server.CreateObject("ADODB.recordset")
    			    'Search for the sanction to get the name, city, and other info
                    sSQL = "SELECT top 1 TSanction,TName,TCity,TState,TDateE from "& SanctionTableName &" WHERE lower(TournAppID) = '" & sqlclean(lcase(right(left(objFile.Name,17),6))) & "'"
                    rs.open sSQL, sConnectionToTRATable, 3, 1
                    IF NOT rs.EOF THEN
     			      'write the title popup box info
                      Response.Write (""" title=""")
                      Response.Write (rs("tname") & "&#13;" & rs("tcity") & ", " & rs("tstate") & "&#13;" & rs("tdatee")) 
                    END IF
    			    'write the files name
                    Response.Write (""">" & right(left(objFile.Name,22),11))
                    Response.Write ("</a></strong>")
                    rs.Close
                    set rs = nothing
                    CloseCon
    	            %>
    	          </A>
    	        </FONT></center>
    	        </TD>
                <%
                dim objFSO2
                dim objstream
                dim linecount
                set objFSO2=server.createobject("scripting.filesystemObject")
                  set objstream=objFSO2.opentextfile(smappath & objfile.name, 1, 0)
                    DO WHILE NOT objstream.atendofStream
                      objstream.skipline
                    LOOP
                    IF ucase(right(objfile.name,3)) = "WSP" THEN
                      '  WSP files have a header line which is not really a bad record.
                      linecount = objstream.Line - 1
                    ELSE
                      linecount = objstream.Line
                    END IF
                  objstream.close
                set objFSO2 = nothing
                response.write "<td><center><font size=""2"">"
                response.write (linecount-1)
                response.write "</font></center></td>"
                BadRecordsTotal = BadRecordsTotal + linecount - 1
                %>
                <td>
                  <center>
                    <%IF Session("UserLevel") > 9 THEN %>
                      <a href="/rankings/DefaultHQ.asp?process=download&file=exceptions\<%Response.Write(objFile.Name)%>&rid=<%=rid%>"><img SRC="/rankings/images/buttons/download.gif" VALUE="Download SELECTed File" TITLE="Download the SELECTed file" border=0></a>
                    <% END IF %>
                  <%
                  IF ucase(right(objfile.name,3)) = "WSP" THEN
                  %>
                    <a href="/rankings/exceptionmgmt-wsp.asp?file=<%Response.Write(objfile.name)%>&line=2&rid=<%=rid%>"><img SRC="/rankings/images/buttons/toolbut.gif" VALUE="Fix Exceptions" TITLE="Fix Exceptions" border=0></a>
                  <%
                  ELSE
                  %>
                    <a href="/rankings/exceptionmgmt-pdf.asp?file=<%Response.Write(objfile.name)%>&line=1&rid=<%=rid%>"><img SRC="/rankings/images/buttons/toolbut.gif" VALUE="Fix Exceptions" TITLE="Fix Exceptions" border=0></a>
                  <%
                  END IF
                  %>
                    <%
                    IF Session("UserLevel") > 9 THEN 
                      %>
                      <a href="/rankings/DefaultHQ.asp?process=delete&file=exceptions\<%Response.Write(objFile.Name)%>&rid=<%=rid%>&returnurl=/rankings/DefaultHQ.asp?process=badscores"><img SRC="/rankings/images/buttons/delete.gif" VALUE="Delete SELECTed File" TITLE="Delete the SELECTed file" border=0></a>
                    <%END IF%>
                  </center>
                </td>

                <TD><center>
    	          <FONT size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">
    	          <% 
    		 	  'We will format the file size so it looks pretty
    	 		   IF objFile.Size <1024 THEN
    				  Response.Write objFile.Size & " Bytes"
    			   ELSEIF objFile.Size < 1048576 THEN
    				  Response.Write Round(objFile.Size / 1024.1) & " KB"
    			   ELSE
    				  Response.Write Round((objFile.Size/1024)/1024.1) & " MB"
    			   END IF
    	           %>
	   	         </FONT></center>
  	  	       </TD>

               <TD><center>
    		       <FONT size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">
    			   <%	'the files Date 
    			  	  Response.Write objFile.DateLastModified
    			   %>
    		      </FONT></center>
    	        </TD>
     	       </TR>
    	       <%
    		 END IF ' This one determines IF it is filtered or not
    	  Next ' This is the FOR LOOP of every file in the directory
        END IF ' This determines IF there are any files at all or not
      IF ListingsCount = 0 THEN
    	%>
          <TR> 
          <TD colspan=5 align="center"><font color="red">Update Search Settings 
          	and Press 'Filter Results' - Or There Are No Files to Display</font></TD> 
</TR>
        <%  
      ELSE
        %>
               <tr>
                 <td colspan=5 align="center">
                   <%=BadRecordsTotal%> records require correction in <%=ListingsCount%>
                    unique tournaments. 
                 </td>
               </tr>
        <%
      END IF
      

     %>
     </TABLE>
     <%      
    
    WriteIndexPageFooter
    
END IF


' ---------------------------- UPLOADED WSP FILE MANAGEMENT --------------------------------


IF (trim(request("process")) = "uploadedwsps") and (session("userlevel") > 1) THEN

    RegionSELECTed = request("region") 
    YearSELECTed = request("year")
' IF it's blank, THEN we use the Session value.
    IF RegionSELECTed = "" THEN RegionSELECTed = trim(session("RegionSELECTed"))
' IF it's still blank, THEN we use "ALL"
    IF RegionSELECTed = "" THEN RegionSELECTed = "All"

' IF it's blank, THEN we use the Session value.
    IF YearSELECTed = "" THEN YearSELECTed = trim(session("YearSELECTed"))
' IF it's still blank, THEN we use current calendar year
    IF YearSELECTed = "" THEN YearSELECTed = right(FormatDateTime(Date,2),2)

    session("RegionSELECTed") = RegionSELECTed
    session("YearSELECTed") = YearSELECTed
    
    
    WriteIndexPageHeader
    sMapPath = PathToTRA & "uploads\"

    %>
    <center>
    <br><br>
    You are currently viewing WSP files which have been uploaded and processed.
    <br>
    </center>
        <form method=post action="/rankings/DefaultHQ.asp">
    <input type="hidden" name="process" value="uploadedwsps">
    <table width=90% align=center>
    <tr><td>

    Region: 

    <SELECT name="region">
      <option value ="All"<%IF RegionSELECTed = "All" THEN Response.Write(" SELECTed ")%>>All Regions</Option><br>
      <option value ="C"<%IF RegionSELECTed = "C" THEN Response.Write(" SELECTed ")%>>S. Central</Option><br>
      <option value ="M"<%IF RegionSELECTed = "M" THEN Response.Write(" SELECTed ")%>>MidWest</Option><br>
      <option value ="W"<%IF RegionSELECTed = "W" THEN Response.Write(" SELECTed ")%>>West</Option><br>
      <option value ="S"<%IF RegionSELECTed = "S" THEN Response.Write(" SELECTed ")%>>South</Option><br>
      <option value ="E"<%IF RegionSELECTed = "E" THEN Response.Write(" SELECTed ")%>>East</Option><br>
      <option value ="U"<%IF RegionSELECTed = "U" THEN Response.Write(" SELECTed ")%>>All NCWSA</Option><br>
    </SELECT>

    </td><td align=center>

    Ski Year:

    <%
	BuildYearDrop
    %>

    </td></tr>
    </table>
    <center><input type=submit value="Filter Results"></center>
    </form>
    <hr>
    <TABLE class="innertable" width=90% align=center>
    <TR align="left" valign="top" bgcolor="#000000"> 
    <TH width="12%"><FONT color="#FFFFFF"><center><B><FONT size="2" face="Verdana, Arial, Helvetica, sans-serif">File Name</FONT></B></center></FONT></TH>
    <TH width="10%"><FONT color="#FFFFFF"><center><B><FONT size="2" face="Verdana, Arial, Helvetica, sans-serif">Actions</FONT></B></center></FONT></TH>
    <TH width="8%"><FONT color="#FFFFFF"><center><B><FONT size="2" face="Verdana, Arial, Helvetica, sans-serif">Size</FONT></B></center></FONT></TH>
    <TH width="20%"><FONT color="#FFFFFF"><center><B><FONT size="2" face="Verdana, Arial, Helvetica, sans-serif">Date</FONT></B></center></FONT></TH>
    </TR>
    <%

	'Create File System Object To Get list of files
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	'Get The path For the web page and its dir.
	'Set the object folder To the mapped path
	Set objFolder = objFSO.GetFolder(sMapPath)
	set objFilesInFolder = objFolder.Files

        ListingsCount = 0
    	'For Each file In the folder

        IF objFilesInFolder.Count <> 0 THEN
          For Each objFile In objFolder.Files

    	    IF (RegionSELECTed = "All" or RegionSELECTed = ucase(right(left(objFile.Name,3),1))) AND (YearSELECTed = "All" or RIGHT(TRIM(YearSELECTed),2) = left(objFile.Name,2)) THEN


              ListingsCount = ListingsCount + 1
%>
          	<TR align="left" valign="top" bordercolor="#999999" bgcolor="#FFFFFF"> 
        	<TD><center><FONT size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">

            <a 

        	<%
                    OpenCon
                    Set rs=Server.CreateObject("ADODB.recordset")
    			    'Search for the sanction to get the name, city, and other info
                    sSQL = "SELECT top 1 TSanction,TName,TCity,TState,TDateE from "& SanctionTableName &" WHERE lower(TournAppID) = '" & sqlclean(lcase(left(objFile.Name,6))) & "'"
                    rs.open sSQL, sConnectionToTRATable, 3, 1
                    IF NOT rs.EOF THEN
     			      'write the title popup box info
                      Response.Write ("title=""" & rs("tname") & "&#13;" & rs("tcity") & ", " & rs("tstate") & "&#13;" & rs("tdatee")) 
                    END IF
    			    'write the files name
			        			Response.Write """>" & objFile.Name
                    rs.Close
                    set rs = nothing
                    CloseCon


    

        	%>
        </A></FONT></center>
        	</TD>
            <td><center>
    
            <%IF Session("UserLevel") > 9 THEN %>
                <a href="/rankings/DefaultHQ.asp?process=download&file=uploads\<%Response.Write(objFile.Name)%>&rid=<%=rid%>"><img SRC="/rankings/images/buttons/download.gif" VALUE="Open or Download this Score File" TITLE="Open or Download this Score File" border=0></a>
            <%END IF%>
            <%IF Session("UserLevel") > 9 AND Session("AdminMenuLevel") > 9 THEN %>
                <a href="/rankings/DefaultHQ.asp?process=delete&file=uploads\<%Response.Write(objFile.Name)%>&rid=<%=rid%>&returnurl=/rankings/DefaultHQ.asp?process=uploadedwsps"><img SRC="/rankings/images/buttons/delete.gif" VALUE="Delete this Score File" TITLE="Delete this Score File" border=0></a>
            <%END IF%>
            </center></td>

            <TD><center>
	        	<FONT size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">
        	
    	    	<%
        			'We will format the file size so it looks pretty
        			IF objFile.Size <1024 THEN
        				Response.Write objFile.Size & " Bytes"
        			ELSEIF objFile.Size < 1048576 THEN
        				Response.Write Round(objFile.Size / 1024.1) & " KB"
        			ELSE
        				Response.Write Round((objFile.Size/1024)/1024.1) & " MB"
        			END IF
      	  	%>
  	  	    </FONT></center></TD>

            <TD><center>
            <FONT size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">
        			<%	'the files Date 
        				Response.Write objFile.DateLastModified
        			%>
	        	</FONT></center></TD>

        	</TR>
        	<%
    		 END IF ' This one determines IF it is filtered or not
    	  Next ' This is the FOR LOOP of every file in the directory
        END IF ' This determines IF there are any files at all or not
      IF ListingsCount = 0 THEN
    	%>
    	<TR align="center" valign="top"> 
    	<TD colspan=4><font size="1" color="red">Update Search Settings and Press 
    		'Filter Results' - Or There Are No Files to Display</font></TD>
    	</TR> 
        <%  
      END IF
    %> 
    </TABLE>

    <%
    
    WriteIndexPageFooter
    
END IF


' ---------------------------- UPLOADED ZIP FILE MANAGEMENT --------------------------------


IF (trim(request("process")) = "uploadedzips") and (session("userlevel") > 1) THEN

    RegionSELECTed = request("region") 
    YearSELECTed = request("year")
' IF it's blank, THEN we use the Session value.
    IF RegionSELECTed = "" THEN RegionSELECTed = trim(session("RegionSELECTed"))
' IF it's still blank, THEN we use "??"
    IF RegionSELECTed = "" THEN RegionSELECTed = "??"

' IF it's blank, THEN we use the Session value.
    IF YearSELECTed = "" THEN YearSELECTed = trim(session("YearSELECTed"))
' IF it's still blank, THEN we use current calendar year
    IF YearSELECTed = "" THEN YearSELECTed = right(FormatDateTime(Date,2),2)

    session("RegionSELECTed") = RegionSELECTed
    session("YearSELECTed") = YearSELECTed
    
    
    WriteIndexPageHeader
    sMapPath = PathToTRA & "PostTourZips\"

    %>
    <center>
    <br><br>
    You are currently viewing ZIP archive files which have been uploaded.
    <br>
    </center>
        <form method=post action="/rankings/DefaultHQ.asp">
    <input type="hidden" name="process" value="uploadedzips">
    <table width=90% align=center>
    <tr><td>

    Region: 

    <SELECT name="region">
      <option value ="??"<%IF RegionSELECTed = "??" THEN Response.Write(" SELECTed ")%>>Please Select</Option><br>
      <option value ="All"<%IF RegionSELECTed = "All" THEN Response.Write(" SELECTed ")%>>All Regions</Option><br>
      <option value ="C"<%IF RegionSELECTed = "C" THEN Response.Write(" SELECTed ")%>>South Central</Option><br>
      <option value ="M"<%IF RegionSELECTed = "M" THEN Response.Write(" SELECTed ")%>>MidWestern</Option><br>
      <option value ="W"<%IF RegionSELECTed = "W" THEN Response.Write(" SELECTed ")%>>Western</Option><br>
      <option value ="S"<%IF RegionSELECTed = "S" THEN Response.Write(" SELECTed ")%>>Southern</Option><br>
      <option value ="E"<%IF RegionSELECTed = "E" THEN Response.Write(" SELECTed ")%>>Eastern</Option><br>
      <option value ="U"<%IF RegionSELECTed = "U" THEN Response.Write(" SELECTed ")%>>All NCWSA</Option><br>
    </SELECT>

    </td><td align=center>

    Ski Year:

    <%
	BuildYearDrop
    %>

    </td></tr>
    </table>
    <center><input type=submit value="Filter Results"></center>
    </form>
    <hr>
    <TABLE class="innertable" width=90% align=center>
    <TR align="left" valign="top" bgcolor="#000000"> 
    <TH width="12%"><FONT color="#FFFFFF"><center><B><FONT size="2" face="Verdana, Arial, Helvetica, sans-serif">File Name</FONT></B></center></FONT></TH>
    <TH width="10%"><FONT color="#FFFFFF"><center><B><FONT size="2" face="Verdana, Arial, Helvetica, sans-serif">Actions</FONT></B></center></FONT></TH>
    <TH width="8%"><FONT color="#FFFFFF"><center><B><FONT size="2" face="Verdana, Arial, Helvetica, sans-serif">Size</FONT></B></center></FONT></TH>
    <TH width="20%"><FONT color="#FFFFFF"><center><B><FONT size="2" face="Verdana, Arial, Helvetica, sans-serif">Date</FONT></B></center></FONT></TH>
    </TR>
    <%

	'Create File System Object To Get list of files
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	'Get The path For the web page and its dir.
	'Set the object folder To the mapped path
	Set objFolder = objFSO.GetFolder(sMapPath)
	set objFilesInFolder = objFolder.Files

        ListingsCount = 0
    	'For Each file In the folder

        IF objFilesInFolder.Count <> 0 and RegionSelected <> "??" THEN
          For Each objFile In objFolder.Files

    	    IF (RegionSELECTed = "All" or RegionSELECTed = ucase(right(left(objFile.Name,3),1))) AND (YearSELECTed = "All" or RIGHT(TRIM(YearSELECTed),2) = left(objFile.Name,2)) THEN


              ListingsCount = ListingsCount + 1
%>
          	<TR align="left" valign="top" bordercolor="#999999" bgcolor="#FFFFFF"> 
        	<TD><center><FONT size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">

            <a 

        	<%
                    OpenCon
                    Set rs=Server.CreateObject("ADODB.recordset")
                    'Search for the sanction to get the name, city, and other info
                    sSQL = "SELECT top 1 TSanction,TName,TCity,TState,TDateE from "& SanctionTableName &" WHERE lower(TournAppID) = '" & sqlclean(lcase(left(objFile.Name,6))) & "'"
                    rs.open sSQL, sConnectionToTRATable, 3, 1

                    'write the title popup box info
                    IF NOT rs.EOF THEN
                      Response.Write ("title=""" & rs("tname") & "&#13;" & rs("tcity") & ", " & rs("tstate") & "&#13;" & rs("tdatee")) 
                    ELSE
                      Response.write ("title=""Sanction Details Missing")
                    END IF

                   'write the files name
                    Response.Write """>" & objFile.Name

                    rs.Close
                    set rs = nothing
                    CloseCon


    

        	%>
        </A></FONT></center>
        	</TD>
            <td><center>
    
            <%IF Session("UserLevel") > 9 THEN %>
                <a href="/rankings/DefaultHQ.asp?process=download&file=PostTourZips\<%Response.Write(objFile.Name)%>&rid=<%=rid%>"><img SRC="/rankings/images/buttons/download.gif" VALUE="Open or Download this Zip Archive File" TITLE="Open or Download this Zip Archive File" border=0></a>
            <%END IF%>

            </center></td>

            <TD><center>
	        	<FONT size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">
        	
    	    	<%
        			'We will format the file size so it looks pretty
        			IF objFile.Size <1024 THEN
        				Response.Write objFile.Size & " Bytes"
        			ELSEIF objFile.Size < 1048576 THEN
        				Response.Write Round(objFile.Size / 1024.1) & " KB"
        			ELSE
        				Response.Write Round((objFile.Size/1024)/1024.1) & " MB"
        			END IF
      	  	%>
  	  	    </FONT></center></TD>

            <TD><center>
            <FONT size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">
        			<%	'the files Date 
        				Response.Write objFile.DateLastModified
        			%>
	        	</FONT></center></TD>

        	</TR>
        	<%
    		 END IF ' This one determines IF it is filtered or not
    	  Next ' This is the FOR LOOP of every file in the directory
        END IF ' This determines IF there are any files at all or not
      IF ListingsCount = 0 THEN
    	%>
    	<TR align="center" valign="top"> 
    	<TD colspan=4><font size="1" color="red">Update Search Settings and Press 
    		'Filter Results' - Or There Are No Files to Display</font></TD>
    	</TR> 
        <%  
      END IF
    %> 
    </TABLE>

    <%
    
    WriteIndexPageFooter
    
END IF


' ---------------------------- UPLOADED TIMING FILE MANAGEMENT --------------------------------


IF (trim(request("process")) = "uploadedjmptms") and (session("userlevel") > 1) THEN

    RegionSELECTed = request("region") 
    YearSELECTed = request("year")
' IF it's blank, THEN we use the Session value.
    IF RegionSELECTed = "" THEN RegionSELECTed = trim(session("RegionSELECTed"))
' IF it's still blank, THEN we use "ALL"
    IF RegionSELECTed = "" THEN RegionSELECTed = "All"

' IF it's blank, THEN we use the Session value.
    IF YearSELECTed = "" THEN YearSELECTed = trim(session("YearSELECTed"))
' IF it's still blank, THEN we use current calendar year
    IF YearSELECTed = "" THEN YearSELECTed = right(FormatDateTime(Date,2),2)

    session("RegionSELECTed") = RegionSELECTed
    session("YearSELECTed") = YearSELECTed
    
    
    WriteIndexPageHeader
    sMapPath = PathToTRA & "TimingRpts\"

    %>
    <center>
    <br><br>
    Jump Timing CSV Data files which have been uploaded.
    <br>
    </center>
        <form method=post action="/rankings/DefaultHQ.asp">
    <input type="hidden" name="process" value="uploadedjmptms">
    <table width=90% align=center>
    <tr><td>

    Region: 

    <SELECT name="region">
      <option value ="All"<%IF RegionSELECTed = "All" THEN Response.Write(" SELECTed ")%>>All Regions</Option><br>
      <option value ="C"<%IF RegionSELECTed = "C" THEN Response.Write(" SELECTed ")%>>S. Central</Option><br>
      <option value ="M"<%IF RegionSELECTed = "M" THEN Response.Write(" SELECTed ")%>>MidWest</Option><br>
      <option value ="W"<%IF RegionSELECTed = "W" THEN Response.Write(" SELECTed ")%>>West</Option><br>
      <option value ="S"<%IF RegionSELECTed = "S" THEN Response.Write(" SELECTed ")%>>South</Option><br>
      <option value ="E"<%IF RegionSELECTed = "E" THEN Response.Write(" SELECTed ")%>>East</Option><br>
      <option value ="U"<%IF RegionSELECTed = "U" THEN Response.Write(" SELECTed ")%>>All NCWSA</Option><br>
    </SELECT>

    </td><td align=center>

    Ski Year:

    <%
	BuildYearDrop
    %>

    </td></tr>
    </table>
    <center><input type=submit value="Filter Results"></center>
    </form>
    <hr>
    <TABLE class="innertable" width=90% align=center>
    <TR align="left" valign="top" bgcolor="#000000"> 
    <TH width="12%"><FONT color="#FFFFFF"><center><B><FONT size="2" face="Verdana, Arial, Helvetica, sans-serif">File Name</FONT></B></center></FONT></TH>
    <TH width="10%"><FONT color="#FFFFFF"><center><B><FONT size="2" face="Verdana, Arial, Helvetica, sans-serif">Actions</FONT></B></center></FONT></TH>
    <TH width="8%"><FONT color="#FFFFFF"><center><B><FONT size="2" face="Verdana, Arial, Helvetica, sans-serif">Size</FONT></B></center></FONT></TH>
    <TH width="20%"><FONT color="#FFFFFF"><center><B><FONT size="2" face="Verdana, Arial, Helvetica, sans-serif">Date</FONT></B></center></FONT></TH>
    </TR>
    <%

	'Create File System Object To Get list of files
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	'Get The path For the web page and its dir.
	'Set the object folder To the mapped path
	Set objFolder = objFSO.GetFolder(sMapPath)
	set objFilesInFolder = objFolder.Files

        ListingsCount = 0
    	'For Each file In the folder

        IF objFilesInFolder.Count <> 0 THEN
          For Each objFile In objFolder.Files

    	    IF ucase(right(objfile.name,4)) = ".CSV" and (RegionSELECTed = "All" or RegionSELECTed = ucase(right(left(objFile.Name,3),1))) AND (YearSELECTed = "All" or RIGHT(TRIM(YearSELECTed),2) = left(objFile.Name,2)) THEN


              ListingsCount = ListingsCount + 1
%>
          	<TR align="left" valign="top" bordercolor="#999999" bgcolor="#FFFFFF"> 
        	<TD><center><FONT size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">

            <a 

        	<%
                    OpenCon
                    Set rs=Server.CreateObject("ADODB.recordset")
    			    'Search for the sanction to get the name, city, and other info
                    sSQL = "SELECT top 1 TSanction,TName,TCity,TState,TDateE from "& SanctionTableName &" WHERE lower(TournAppID) = '" & sqlclean(lcase(left(objFile.Name,6))) & "'"
                    rs.open sSQL, sConnectionToTRATable, 3, 1
                    IF NOT rs.EOF THEN
     			      'write the title popup box info
                      Response.Write ("title=""" & rs("tname") & "&#13;" & rs("tcity") & ", " & rs("tstate") & "&#13;" & rs("tdatee")) 
                    END IF
    			    'write the files name
			        			Response.Write """>" & objFile.Name
                    rs.Close
                    set rs = nothing
                    CloseCon


    

        	%>
        </A></FONT></center>
        	</TD>
            <td><center>
    
            <%IF Session("UserLevel") > 9 THEN %>
                <a href="/rankings/DefaultHQ.asp?process=download&file=TimingRpts\<%Response.Write(objFile.Name)%>&rid=<%=rid%>"><img SRC="/rankings/images/buttons/download.gif" VALUE="Open or Download this Zip Archive File" TITLE="Open or Download this Jump Times CSV File" border=0></a>
            <%END IF%>

            </center></td>

            <TD><center>
	        	<FONT size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">
        	
    	    	<%
        			'We will format the file size so it looks pretty
        			IF objFile.Size <1024 THEN
        				Response.Write objFile.Size & " Bytes"
        			ELSEIF objFile.Size < 1048576 THEN
        				Response.Write Round(objFile.Size / 1024.1) & " KB"
        			ELSE
        				Response.Write Round((objFile.Size/1024)/1024.1) & " MB"
        			END IF
      	  	%>
  	  	    </FONT></center></TD>

            <TD><center>
            <FONT size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#000000">
        			<%	'the files Date 
        				Response.Write objFile.DateLastModified
        			%>
	        	</FONT></center></TD>

        	</TR>
        	<%
    		 END IF ' This one determines IF it is filtered or not
    	  Next ' This is the FOR LOOP of every file in the directory
        END IF ' This determines IF there are any files at all or not
      IF ListingsCount = 0 THEN
    	%>
    	<TR align="center" valign="top"> 
    	<TD colspan=4><font size="1" color="red">Update Search Settings and Press 
    		'Filter Results' - Or There Are No Files to Display</font></TD>
    	</TR> 
        <%  
      END IF
    %> 
    </TABLE>

    <%
    
    WriteIndexPageFooter
    
END IF


' ---------------------------- FIX (REPLACE) BAD SANCTION ID --------------------------------

IF (trim(request("process")) = "fixsanction") and (session("userlevel") > 9) THEN

    WriteIndexPageHeader
    NewsPageNum = "3"
    %>
    
    
    <br><br>
    <center><h2>Fix Bad Sanction Code<br></h2>
    <br><br>
    
    <form action="/rankings/DefaultHQ.asp" method="post">
    <input type="hidden" name="process" value="fixsanction2">
    Enter the Sanction ID That you Would Like to Change:<br> 
    <input type="text" name="OrigTourID" size="10"><br><br>
    <input type="submit" value="Submit"><br><br><br>
    </form>
    </center>
    
    <%
    
    WriteIndexPageFooter

END IF

IF (trim(request("process")) = "fixsanction2") and (session("userlevel") > 9) THEN

    WriteIndexPageHeader
    News="Now enter the correct Sanction ID that you would like to replace the other one with."
    %>
    
    
    <br><br>
    <center><h2>Fix Bad Sanction Code<br></h2>
    <br><br>
    
    <form action="/rankings/DefaultHQ.asp" method="post">
    <input type="hidden" name="process" value="fixsanction3">
    Enter the new Sanction ID which will replace the old Sanction ID:<br> 
    <input type="text" name="NewTourID" size="10"><br><br>
    <input type="hidden" name="OrigTourID" value="<%=Request("OrigTourID")%>"><br><br>
    <input type="submit" value="Submit"><br><br><br>
    </form>
    </center>
    <%
    
    WriteIndexPageFooter

END IF

IF (trim(request("process")) = "fixsanction3") and (session("userlevel") > 9) THEN

    WriteIndexPageHeader

    OpenConSanction
    Set rs=Server.CreateObject("ADODB.recordset")
    sSQL = "SELECT top 1 * from TSchedul WHERE upper(TSanction) = '" & SQLClean(ucase(Request("NewTourID"))) & "'"
    rs.open sSQL, sConnectionToSanctionTable
    IF rs.EOF THEN
      ErrorCheck = 1
    END IF
    rs.Close
    Set rs= Nothing
    CloseConSanction
    IF Len(trim(Request("OrigTourID"))) <> Len(trim(Request("NewTourID"))) THEN
      ErrorCheck = 1
    END IF

    %>
    
    <br><br>
    <center><h2>Fix Bad Sanction Code<br></h2>
    <br><br>
        
<% IF ErrorCheck = 0 THEN
    News="Confirm Replacement of " & request("OrigTourID") & " with " & request("NewTourID")
    %>
    <h4>Are you sure you want to replace the<br>
    original sanction code, <font color="red"><%=Request("OrigTourID")%> </font><br>
    with the new sanction code, <font color="red"><%=Request("NewTourID")%> </font>??
    </h4>
    
    <form action="/rankings/DefaultHQ.asp" method="post">
    <input type="hidden" name="process" value="fixsanction4">
    <input type="submit" value="NO"><br>
    </form>    
    <form action="/rankings/DefaultHQ.asp" method="post">
    <input type="hidden" name="OrigTourID" value="<%=ucase(Request("OrigTourID"))%>">
    <input type="hidden" name="NewTourID" value="<%=ucase(Request("NewTourID"))%>">
    <input type="hidden" name="process" value="fixsanction5">
    <input type="submit" value="YES"><br><br><br>
    </form>
<% ELSE
    News="Invalid Update Request"
    %>
    
    
    <br><br>
    <center>
    <br><br>
    The value you entered is not valid.  Update aborted.
    <br><br><br>
<% END IF
    response.write ("</center>")
    
    WriteIndexPageFooter

END IF


IF (trim(request("process")) = "fixsanction4") and (session("userlevel") > 9) THEN

    WriteIndexPageHeader
    %>
    <br><br><br><br>
    <center>By request, the update has <b><u>not</u></b> been saved.<br>
    <br><br>
    </center>
    
    <%
    
    WriteIndexPageFooter

END IF

IF (trim(request("process")) = "fixsanction5") and (session("userlevel") > 9) THEN

    Dim FindException
    
    FindException = 0 ' We will use this to determine IF we found an excpt file.

    ' This code replaces the old TourID in the scores database.
    '    
    OpenCon
    set rs = Server.CreateObject("ADODB.recordset")
    sSQL = "Update " & RawScoresTableName & " set tourid = '" & SQLClean(ucase(request("NewTourID"))) & "' WHERE upper(TourID) = '" & SQLClean(ucase(request("OrigTourID"))) & "'"
    con.execute(sSQL)
    WriteLog(date() &"  "& time() &"  Replaced TourID " & Request("OrigTourID") & " with " & Request("NewTourID") & ".")
    CloseCon
    
    set tempFSO=Server.CreateObject("Scripting.FileSystemObject")
    set tempObjStream=tempFSO.GetFolder(PathToExceptions)
    for each FileName in tempObjStream.files
      'Search the name of all files in the folder for the keyword
      '
      IF inStr(ucase(FileName),Request("OrigTourID")) > 0 THEN
         FindException = 1
         ' We found a file with the bad tour id ... now scan through the file
         ' and replace all the tour ids within the file
         '
         set tempFSO2=Server.CreateObject("Scripting.FileSystemObject")
         set tempObjStream2=tempFSO2.opentextfile(FileName)
         TextFile = "" ' this will hold the contents of the new text
         DO WHILE NOT tempObjStream2.AtEndOfStream
           TextFile = TextFile & replace(ucase(tempObjStream2.Readline), Request("OrigTourID"), Request("NewTourID")) & vbCrLf
         LOOP
         tempObjStream2.close
         set tempObjStream2 = Nothing
         ' Now we have the new text to rewrite into the file.
         '
         set tempObjStream2=tempFSO2.opentextfile(FileName,2,true)
         tempObjStream2.write(textfile)
         tempObjStream2.close
         set tempObjStream2 = Nothing
         ' Next we want to change the file name
         '
         tempFSO2.MoveFile FileName, replace(FileName,Request("OrigTourID"),Request("NewTourID"))
         set tempFSO2 = NOTHING
       END IF
    next
    ' Now we can destroy the File Scripting Objects
    set tempObjStream = Nothing
    set tempFSO = Nothing
    IF FindException > 0 THEN
      ' Now we do the same search for the Reasons folder
      set tempFSO=Server.CreateObject("Scripting.FileSystemObject")
      set tempObjStream=tempFSO.GetFolder(PathToReasons)
      for each FileName in tempObjStream.files
        'Search the name of all files in the folder for the keyword
        '
        IF inStr(ucase(FileName),Request("OrigTourID")) > 0 THEN
           ' We found a file with the bad tour id ... now change the file name.
           '
           set tempFSO2=Server.CreateObject("Scripting.FileSystemObject")
           tempFSO2.MoveFile FileName, replace(FileName,Request("OrigTourID"),Request("NewTourID"))
           set tempFSO2 = NOTHING
         END IF
      next
    END IF
    
    ' Now we can destroy the File Scripting Objects
    set tempObjStream = nothing
    set tempFSO = NOTHING
  
    ' Now we do the same replace on the files in the uploads director.
    '
    set tempFSO=Server.CreateObject("Scripting.FileSystemObject")
    set tempObjStream=tempFSO.GetFolder(PathToTRA & "\uploads")
    for each FileName in tempObjStream.files
      'Search the name of all files in the folder for the keyword
      '
      IF inStr(ucase(FileName),Request("OrigTourID")) > 0 THEN
         FindException = 1
         ' We found a file with the bad tour id ... now scan through the file
         ' and replace all the tour ids within the file
         '
         set tempFSO2=Server.CreateObject("Scripting.FileSystemObject")
         set tempObjStream2=tempFSO2.opentextfile(FileName)
         TextFile = "" ' this will hold the contents of the new text
         DO WHILE NOT tempObjStream2.AtEndOfStream
           TextFile = TextFile & replace(ucase(tempObjStream2.Readline), Request("OrigTourID"), Request("NewTourID")) & vbCrLf
         LOOP
         tempObjStream2.close
         set tempObjStream2 = Nothing
         ' Now we have the new text to rewrite into the file.
         '
         set tempObjStream2=tempFSO2.opentextfile(FileName,2,true)
         tempObjStream2.write(textfile)
         tempObjStream2.close
         set tempObjStream2 = Nothing
         ' Next we want to change the file name
         '
         tempFSO2.MoveFile FileName, replace(FileName,Request("OrigTourID"),Request("NewTourID"))
         set tempFSO2 = NOTHING
       END IF
    next
    ' Now we can destroy the File Scripting Objects
    set tempObjStream = Nothing
    set tempFSO = Nothing

    WriteIndexPageHeader

    %>
    <br><br><br><br>
    <center>The update has been made.  All instances of <%=Request("OrigTourID")%> have been
    replaced with <%=Request("NewTourID")%>.<br>
    <br>
<% IF FindExceptions > 0 THEN %>
    Some exceptions were still waiting to be processed for the old Sanction ID, <%=Request("OrigTourID")%>.<br>
    Those records have been updated as well and now display the new Sandction ID, <%=Request("NewTourID")%>.<br>
<% END IF %>
    <br><br>
    </center>
    
    <%
    
    WriteIndexPageFooter

END IF

' ---------------------------- DELETE ENTIRE TOURNAMENT --------------------------------

sfname=request("sfname")

IF (trim(request("process")) = "deletetour") and (session("userlevel") > 9) THEN


    WriteIndexPageHeader
    NewsPageNum = "54a"

    %>
	
  
    <br><br>
    <center><h2>Delete Tournament Scores<br></h2>
    <br>
    <form action="/rankings/DefaultHQ.asp" method="post"><%

    IF Session("adminmenulevel")>=30 THEN  %>	
	
	Sports Discipline &nbsp;&nbsp;&nbsp;&nbsp;	
    	<input type="radio" name="sfname" value=1>AWSA or NCWSA &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    	<input type="radio" name="sfname" value=2>All Other SD's<br><br><%
    ELSE %>
    	<input type="hidden" name="sfname" value=0><%	
    END IF  %>		 

    <input type="hidden" name="process" value="deletetour2">
    Enter the Sanction ID That you Would Like to Delete:<br><br> 

    <input type="text" name="TourID2Del" size="10"><br><br>
    <input type="submit" style="width:9em" value="Submit"><br><br><br>
    </form>
    </center>
    
    <%
    
    WriteIndexPageFooter

END IF

IF (trim(request("process")) = "deletetour2") and (session("userlevel") > 9) THEN

    OpenCon
      set rs = Server.CreateObject("ADODB.recordset")
      IF Session("UserSptsGrpID")="AWS" OR Session("UserSptsGrpID")="NCW" OR sfname=1 THEN 
	      sSQL = "SELECT count(*) as ScoreCount from (Select distinct MemberID from "
	      sSQL = sSQL & RawScoresTableName & " WHERE upper(TourID) = '" 
	      sSQL = sSQL & SQLClean(ucase(trim(request("TourID2Del")))) & "') xx;"
      ELSEIF (Session("UserSptsGrpID")<>"AWS" AND Session("UserSptsGrpID")<>"NCW") OR sfname=2 THEN 
	      sSQL = "SELECT count(*) as ScoreCount from (Select distinct MemberID from "
	      sSQL = sSQL & RawScoresOtherTableName & " WHERE upper(TourID) = '" 
	      sSQL = sSQL & SQLClean(ucase(trim(request("TourID2Del")))) & "') xx;"
      END IF	

      rs.open sSQL, SConnectionToTRATable, 3, 3        
      IF rs.eof THEN

        WriteIndexPageHeader
        News="No Matching Scores Found."
        %>
        <br><br>
        <center><h2>Delete Tournament Scores<br></h2>
        <br><br>
        
        No scores were found which matched Tournament ID: <%=Request("TourID2Del")%><br>
        You may try another Tournament.<br><br>
        <form action="/rankings/DefaultHQ.asp" method="post">
        <input type="hidden" name="process" value="deletetour2">
	<input type="hidden" name="sfname" value="<%=sfname%>">
        Enter the Sanction ID That you Would Like to Delete:<br> 
        <input type="text" name="TourID2Del" size="10"><br><br>
        <input type="submit" value="Submit"><br><br><br>
        </form>
        </center>
        <%
        WriteIndexPageFooter

      ELSE

	' --- First check SWIFT to determine the SptsGrpID ---
	TempTour=LEFT(SQLClean(ucase(request("TourID2Del"))),6)

	sSQL = "SELECT TOP 1 * FROM "&SanctionTableName
	sSQL = sSQL + " WHERE LEFT(TSanction,6) = '"&TempTour&"'" 
	Set rsTour=Server.CreateObject("ADODB.recordset")
	rsTour.open sSQL, SConnectionToSanctionTable, 3, 1



	' --- Tests the SD authority of this person to this tournament ---
	' --- Note that revised logic allows AWS users to act for NCW ---
	tsgi = rsTour("SptsGrpID"): usgi = Session("UserSptsGrpID"): amlvl = Session("adminmenulevel")
	IF (tsgi<>usgi AND (NOT tsgi="NCW" AND usgi="AWS")) AND amlvl<50 THEN
		Session("SptsGrpID") = tsgi
		rsTour.close
		response.redirect("/rankings/tools.asp?svar=reject")
	END IF


        WriteIndexPageHeader
        News="Confirm Deletion of " & request("TourID2Del") & "."
        %>
        <br>
        <center><h2>Delete Tournament Scores<br></h2>
        <br><br>
        
        There are scores found for <b><%=rs("ScoreCount")%> skiers</b> for this Tournament ID.<br><br>
        Are you certain you wish to delete ALL Scores and Files that<br>
        are associated with Tournament Sanction ID: <%=Request("TourID2Del")%></center><br> 

	<table align=center width=65%>
	<tr>
        <form action="/rankings/DefaultHQ.asp" method="post">
	<td align=center>
        <input type="hidden" name="process" value="deletetour3">
        <input type="submit" style="width:9em" value="NO"><br>
	</td>
        </form>    
	<td align=center>
        <form action="/rankings/DefaultHQ.asp" method="post">
	<input type="hidden" name="sfname" value="<%=sfname%>">
        <input type="hidden" name="TourID2Del" value="<%=Request("TourID2Del")%>">
        <input type="hidden" name="process" value="deletetour4">
        <input type="submit" style="width:9em" value="YES"><br><br><br>
	</td>
        </form>
	</tr>
	</table>

        <%
        WriteIndexPageFooter

      END IF
      rs.close
      set rs = Nothing
    CloseCon

END IF

IF (trim(request("process")) = "deletetour3") and (session("userlevel") > 9) THEN

    WriteIndexPageHeader
    %>
    <br><br><br><br>
    <center>By request, the Tournament has <b><u>not</u></b> been deleted.<br>
    <br><br>
    </center>
    
    <%
    
    WriteIndexPageFooter

END IF

IF (trim(request("process")) = "deletetour4") and (session("userlevel") > 9) THEN

    ' This code deletes score rows for a TourID in the scores database.
    '    
    OpenCon

    IF Session("UserSptsGrpID")="AWS" OR Session("UserSptsGrpID")="NCW" OR sfname=1 THEN 
       sSQL = "DELETE from " & RawScoresTableName & " WHERE upper(TourID) = '" & SQLClean(ucase(trim(request("TourID2Del")))) & "'"
    ELSEIF (Session("UserSptsGrpID")<>"AWS" AND Session("UserSptsGrpID")<>"NCW") OR sfname=2 THEN 
       sSQL = "DELETE from " & RawScoresOtherTableName & " WHERE upper(TourID) = '" & SQLClean(ucase(trim(request("TourID2Del")))) & "'"
    END IF

    con.execute(sSQL)
    WriteLog(date() &"  "& time() &"  Deleted TourID " & Request("TourID2Del") & ".")
    CloseCon
   

    ' Delete any files in the Exceptions folder with this TourID    
    set tempFSO=Server.CreateObject("Scripting.FileSystemObject")
    set tempObjStream=tempFSO.GetFolder(PathToExceptions)
    for each FileName in tempObjStream.files
      IF inStr(ucase(FileName),ucase(Request("TourID2Del"))) > 0 THEN
         set tempFSO2=Server.CreateObject("Scripting.FileSystemObject")
           tempFSO2.DeleteFile(FileName)
         set tempFSO2 = NOTHING
       END IF
    next
    set tempObjStream = Nothing
    set tempFSO = Nothing


    ' Delete any files in the Reasons folder with this TourID    
    set tempFSO=Server.CreateObject("Scripting.FileSystemObject")
    set tempObjStream=tempFSO.GetFolder(PathToReasons)
    for each FileName in tempObjStream.files
      IF inStr(ucase(FileName),ucase(Request("TourID2Del"))) > 0 THEN
         set tempFSO2=Server.CreateObject("Scripting.FileSystemObject")
           tempFSO2.DeleteFile(FileName)
         set tempFSO2 = NOTHING
       END IF
    next
    set tempObjStream = nothing
    set tempFSO = NOTHING

  
    ' Delete any files in the Uploads folder with this TourID    
    set tempFSO=Server.CreateObject("Scripting.FileSystemObject")
    set tempObjStream=tempFSO.GetFolder(PathtoUploads)
    for each FileName in tempObjStream.files
      IF inStr(ucase(FileName),ucase(Request("TourID2Del"))) > 0 THEN
         set tempFSO2=Server.CreateObject("Scripting.FileSystemObject")
           tempFSO2.DeleteFile(FileName)
         set tempFSO2 = NOTHING
       END IF
    next
    set tempObjStream = Nothing
    set tempFSO = Nothing

    ' Delete any files in the PostTourZips folder with this TourID    
    set tempFSO=Server.CreateObject("Scripting.FileSystemObject")
    set tempObjStream=tempFSO.GetFolder(PathtoZips)
    for each FileName in tempObjStream.files
      IF inStr(ucase(FileName),ucase(Request("TourID2Del"))) > 0 THEN
         set tempFSO2=Server.CreateObject("Scripting.FileSystemObject")
           tempFSO2.DeleteFile(FileName)
         set tempFSO2 = NOTHING
       END IF
    next
    set tempObjStream = Nothing
    set tempFSO = Nothing

    ' Delete any files in the RawWSPs folder with this TourID    
    set tempFSO=Server.CreateObject("Scripting.FileSystemObject")
    set tempObjStream=tempFSO.GetFolder(PathtoRawWSPs)
    for each FileName in tempObjStream.files
      IF inStr(ucase(FileName),ucase(Request("TourID2Del"))) > 0 THEN
         set tempFSO2=Server.CreateObject("Scripting.FileSystemObject")
           tempFSO2.DeleteFile(FileName)
         set tempFSO2 = NOTHING
       END IF
    next
    set tempObjStream = Nothing
    set tempFSO = Nothing

    ' Delete any files in the HQInBox folder with this TourID    
    set tempFSO=Server.CreateObject("Scripting.FileSystemObject")
    set tempObjStream=tempFSO.GetFolder(PathtoHQInBox)
    for each FileName in tempObjStream.files
      IF inStr(ucase(FileName),left(ucase(Request("TourID2Del")),6)) > 0 THEN
         set tempFSO2=Server.CreateObject("Scripting.FileSystemObject")
           tempFSO2.DeleteFile(FileName)
         set tempFSO2 = NOTHING
       END IF
    next
    set tempObjStream = Nothing
    set tempFSO = Nothing

    ' Delete any files in the Scorebks folder with this TourID    
    set tempFSO=Server.CreateObject("Scripting.FileSystemObject")
    set tempObjStream=tempFSO.GetFolder(PathtoScorebks)
    for each FileName in tempObjStream.files
      IF inStr(ucase(FileName),left(ucase(Request("TourID2Del")),6)) > 0 THEN
         set tempFSO2=Server.CreateObject("Scripting.FileSystemObject")
           tempFSO2.DeleteFile(FileName)
         set tempFSO2 = NOTHING
       END IF
    next
    set tempObjStream = Nothing
    set tempFSO = Nothing

    ' Delete any files in the TimingRpts folder with this TourID    
    set tempFSO=Server.CreateObject("Scripting.FileSystemObject")
    set tempObjStream=tempFSO.GetFolder(PathtoTiming)
    for each FileName in tempObjStream.files
      IF inStr(ucase(FileName),left(ucase(Request("TourID2Del")),6)) > 0 THEN
         set tempFSO2=Server.CreateObject("Scripting.FileSystemObject")
           tempFSO2.DeleteFile(FileName)
         set tempFSO2 = NOTHING
       END IF
    next
    set tempObjStream = Nothing
    set tempFSO = Nothing

    ' Delete any files in the InBoxIWWF folder with this TourID    
    set tempFSO=Server.CreateObject("Scripting.FileSystemObject")
    set tempObjStream=tempFSO.GetFolder(PathtoIWWF)
    for each FileName in tempObjStream.files
      IF inStr(ucase(FileName),left(ucase(Request("TourID2Del")),6)) > 0 THEN
         set tempFSO2=Server.CreateObject("Scripting.FileSystemObject")
           tempFSO2.DeleteFile(FileName)
         set tempFSO2 = NOTHING
       END IF
    next
    set tempObjStream = Nothing
    set tempFSO = Nothing

    ' Reset any Post Tournament Flags From "1" to "0" in S_PostTourn table
    ' Then also reset TStatus to 2 or 4, depending on whether any manual reports posted.
    sSQL = "Select top 1 * from " & PostTourTableName & " where upper(TournAppID) = '"
    sSQL = sSQL & left(ucase(Request("TourID2Del")),6) & "'"
    SET objRS=Server.CreateObject("ADODB.recordset")
    objRS.open sSQL, sConnectionToSanctionTable, 3, 3
    If NOT objRS.EOF THEN
      TStatus = 2
      sSQL = "Update " & PostTourTableName & " Set PTF_TNY = 0"
      if objRS("PTF_SBK") = 1 then sSQL = sSQL & ", PTF_SBK = 0": else if objRS("PTF_SBK") = 2 then TStatus = 4
      if objRS("PTF_WSP") = 1 then sSQL = sSQL & ", PTF_WSP = 0": else if objRS("PTF_WSP") = 2 then TStatus = 4
      if objRS("PTF_TS") = 1 then sSQL = sSQL & ", PTF_TS = 0": else if objRS("PTF_TS") = 2 then TStatus = 4
      if objRS("PTF_OD") = 1 then sSQL = sSQL & ", PTF_OD = 0": else if objRS("PTF_OD") = 2 then TStatus = 4
      if objRS("PTF_BT") = 1 then sSQL = sSQL & ", PTF_BT = 0": else if objRS("PTF_BT") = 2 then TStatus = 4
      if objRS("PTF_JT") = 1 then sSQL = sSQL & ", PTF_JT = 0": else if objRS("PTF_JT") = 2 then TStatus = 4
      if objRS("PTF_CS") = 1 then sSQL = sSQL & ", PTF_CS = 0": else if objRS("PTF_CS") = 2 then TStatus = 4
      if objRS("PTF_CJ") = 1 then sSQL = sSQL & ", PTF_CJ = 0": else if objRS("PTF_CJ") = 2 then TStatus = 4
      if objRS("PTF_SD") = 1 then sSQL = sSQL & ", PTF_SD = 0": else if objRS("PTF_SD") = 2 then TStatus = 4
      if objRS("PTF_TU") = 1 then sSQL = sSQL & ", PTF_TU = 0": else if objRS("PTF_TU") = 2 then TStatus = 4
      if objRS("PTF_HD") = 1 then sSQL = sSQL & ", PTF_HD = 0": else if objRS("PTF_HD") = 2 then TStatus = 4
      sSQL = sSQL & " Where upper(TournAppID) = '" & left(ucase(Request("TourID2Del")),6) & "'"
      OpenConSanUpd
      ConSanUpd.Execute(sSQL)
      sSQL = "Update " & SanctionTableName & " Set TStatus = " & TStatus
      sSQL = sSQL & " Where upper(TournAppID) = '" & left(ucase(Request("TourID2Del")),6) & "'"
      ConSanUpd.Execute(sSQL)
      CloseConSanUpd
    END IF
    objRS.close
    SET objRS = Nothing

    WriteIndexPageHeader

    %>
    <br><br><br><br>
    <center>The deletion is complete.  All instances of <%=Request("TourID2Del")%> have been
    deleted.<br>
    <br>
    </center>
    
    <%
    
    WriteIndexPageFooter

END IF

' ---------------------------- LOGIN --------------------------------

IF (trim(request("process")) = "login") THEN

    session("reallogin") = "Valid"
    session("membermenulevel") = ""
    WriteIndexPageHeader
    %>
    
    
    <br><br>
    <center><h2>Welcome to TRA<br>Score Management Center</h2><br>
    <h4>Please log in below.</h4>
    
    <br><br>
    <form action="/rankings/LoginHQ.asp" method="post">
    Username: <input type="text" name="username" size="10"><br><br>
    Password: <input type="password" name="password" size="10"><br><br>
    <br><br>
    <input type="submit" value="Login"><br><br><br>
    </form>
    </center>
    <%
    
    WriteIndexPageFooter

END IF


' ---------------------------------------------------------------------------
' -----------  Defines the stuff in the center frame of the window ----------
' ---------------------------------------------------------------------------

IF trim(request("process")) = "" THEN

Dim MainImage
WriteIndexPageHeader
TourDisplayWidth=675

EventSelected="S"
WhatDropdownImage EventSelected


%>
<TABLE class="droptable" background="<%=MainImage%>" align=center width="<%=TourDisplayWidth%>px" height=250><% '---Table to hold image --- %>
  <TR>
    <TD >&nbsp;

    </TD>
  </TR>
</TABLE><%




WriteIndexPageFooter

END IF

' --- Variable definitions are in Settings.asp ---
M=2
IF M=1 AND trim(request("process")) = "" THEN

'response.write("INSIDE")

    WriteIndexPageHeader
    %>
<TABLE CELLSPACING="0" Align=CENTER CELLPADDING="0" BORDER="0" WIDTH=100%>
  <TR>  
    <TD valign="top" colspan=8 Align=middle nowrap>

	<table WIDTH=100% ALIGN=CENTER CELLSPACING="0" CELLPADDING="8" BORDER="0">
	  <tr>
	    <TD align="center" vAlign=bottom noWrap background="/rankings/images/buttons/Vertical_Shade_564x152_New.jpg">
		    <FONT face="Verdana, Arial, Helvetica, sans-serif" color=#ffffff size=5>
		    <B><% = MainHead_01 %></B></FONT>&nbsp</TD>	
	  </tr>	
	</table>

    </TD>	

  </TR>


  <% ' --- News Box 1 Heading and Text  ---- %>
  <TR><TD colspan=8>&nbsp</TD></TR>
  <TR>
    <TD colspan=8 width=100%>
	<table bordercolor="#00008B" width=100% ALIGN=CENTER CELLSPACING="0" CELLPADDING="0" BORDER="1">
		<tr>
		  <td align=left vAlign="top"><a title="<% =NewsBoxBalloon01 %>"><img src="<% =NewsBoxImage01 %>"></a></td>
		  <td width=20% align=left valign="top">

		  <TABLE CELLPADDING=4><tr><td>

		    <font size="3" face=<% =font2 %> COlOR="<% =textcolor3 %>"><b><%

			' Reads and displays text from news folder
			Set objfso = CreateObject("Scripting.FileSystemObject")
			IF objFSO.FileExists(PathToNews & "\NewsBoxHead_01.txt") THEN
				set objstream=objFSO.opentextfile(PathToNews & "\NewsBoxHead_01.txt")
			   	IF NOT objstream.atendofstream THEN
					DO WHILE not objstream.atendofstream
						response.write(objstream.readline)
						response.write("<br>")
					LOOP
			   	END IF
				objstream.close
			END IF

		    %></b></font>	
		    <br>
		    <font size=1 face=<% =font2 %> COlOR="#000000"><%

			' Reads and displays text from news folder
			Set objfso = CreateObject("Scripting.FileSystemObject")
			IF objFSO.FileExists(PathToNews & "\NewsBoxText_01.txt") THEN
				set objstream=objFSO.opentextfile(PathToNews & "\NewsBoxText_01.txt")
			   	IF NOT objstream.atendofstream THEN
					DO WHILE not objstream.atendofstream
						response.write(objstream.readline)
						response.write("<br>")
					LOOP
			   	END IF
				objstream.close
			END IF

			%></font>
		    </td></tr></TABLE>  
		  </td>



		  <td align=left vAlign="top"><a title="<% =NewsBoxBalloon02 %>"><img src="<% =NewsBoxImage02 %>"></a></td>
		  <td width=20% align=left valign="top">

		  <TABLE CELLPADDING=4><tr><td>
		    <font size="3" face=<% =font2 %> COlOR="<% =textcolor3 %>"><b><%

			' Reads and displays text from news folder
			Set objfso = CreateObject("Scripting.FileSystemObject")
			IF objFSO.FileExists(PathToNews & "\NewsBoxHead_02.txt") THEN
				set objstream=objFSO.opentextfile(PathToNews & "\NewsBoxHead_02.txt")
			   	IF NOT objstream.atendofstream THEN
					DO WHILE not objstream.atendofstream
						response.write(objstream.readline)
						response.write("<br>")
					LOOP
			   	END IF
				objstream.close
			END IF

		    %></b></font>	
		    <br>
		    <font size=1 face=<% =font2 %> COlOR="#000000"><%

			' Reads and displays text from news folder
			Set objfso = CreateObject("Scripting.FileSystemObject")
			IF objFSO.FileExists(PathToNews & "\NewsBoxText_02.txt") THEN
				set objstream=objFSO.opentextfile(PathToNews & "\NewsBoxText_02.txt")
			   	IF NOT objstream.atendofstream THEN
					DO WHILE not objstream.atendofstream
						response.write(objstream.readline)
						response.write("<br>")
					LOOP
			   	END IF
				objstream.close
			END IF

			%></font>
		    </td></tr></TABLE>  

		  </td>


		  <td align=left vAlign="top"><a title="<% =NewsBoxBalloon03 %>"><img src="<% =NewsBoxImage03 %>"></a></td>
		  <td width=20% align=left valign="top">

		  <TABLE CELLPADDING=4><tr><td>
		    <font size="3" face=<% =font2 %> COlOR="<% =textcolor3 %>"><b><%

			' Reads and displays text from news folder
			Set objfso = CreateObject("Scripting.FileSystemObject")
			IF objFSO.FileExists(PathToNews & "\NewsBoxHead_03.txt") THEN
				set objstream=objFSO.opentextfile(PathToNews & "\NewsBoxHead_03.txt")
			   	IF NOT objstream.atendofstream THEN
					DO WHILE not objstream.atendofstream
						response.write(objstream.readline)
						response.write("<br>")
					LOOP
			   	END IF
				objstream.close
			END IF

		    %></b></font>	
		    <br>
		    <font size=1 face=<% =font2 %> COlOR="#000000"><%

			' Reads and displays text from news folder
			Set objfso = CreateObject("Scripting.FileSystemObject")
			IF objFSO.FileExists(PathToNews & "\NewsBoxText_03.txt") THEN
				set objstream=objFSO.opentextfile(PathToNews & "\NewsBoxText_03.txt")
			   	IF NOT objstream.atendofstream THEN
					DO WHILE not objstream.atendofstream
						response.write(objstream.readline)
						response.write("<br>")
					LOOP
			   	END IF
				objstream.close
			END IF

			%></font>
		    </td></tr></TABLE>  

		  </td>



		</tr>
	</table>
     </TD> 	

    <td colspan=3 width=50%>&nbsp</td>
  </TR>



  <TR><TD colspan=8>&nbsp</TD></TR>


  <TR>
    <TD width=33% valign="top" colspan=2>
	<table bordercolor="#00008B" VALIGN=top ALIGN=CENTER CELLSPACING="0" CELLPADDING="0" BORDER="1">
		<tr>
		    <TD  align=center vAlign="top" noWrap background="/rankings/images/buttons/Vertical_Shade_216x20.jpg">
		    <FONT face="Verdana, Arial, Helvetica, sans-serif" color=#ffffff size=1>
		    <B>&nbsp;&nbsp;&nbsp;<% =TVHead01 %></B></FONT>&nbsp</TD>
		    	
		</tr>
		<tr>
		    <td valign="top"><a title="<% =TVBalloon01 %>"><img src="<% =TVImage01 %>"></a></td>
		</tr>
	</table>
    </TD>

    <TD width=33% valign="top" colspan=2>
	<table bordercolor="#00008B" ALIGN=CENTER CELLSPACING="0" CELLPADDING="0" BORDER="1">
		<tr>
		    <TD  align=center vAlign=bottom noWrap background="/rankings/images/buttons/Vertical_Shade_216x20.jpg">
		    <FONT face="Verdana, Arial, Helvetica, sans-serif" color=#ffffff size=1>
		    <B> &nbsp;&nbsp;&nbsp;<% =TVHead02 %></B></FONT>&nbsp</TD>

		</tr>
		<tr>
		    <td valign="top"><a title="<% =TVBalloon02 %>"><img src="<% =TVImage02 %>"></a></td>
		</tr>
	</table>

    </TD>


    <TD width=33% valign="top" colspan=2>
	<table bordercolor="#00008B" ALIGN=CENTER CELLSPACING="0" CELLPADDING="0" BORDER="1">
		<tr VALIGN="top">
		    <TD  align=center vAlign=bottom noWrap background="/rankings/images/buttons/Vertical_Shade_216x20.jpg">
		    <FONT face="Verdana, Arial, Helvetica, sans-serif" color=#ffffff size=1>
		    <B> &nbsp;&nbsp;&nbsp;<% =TVHead04 %> </B></FONT>&nbsp;</TD>

		</tr>
		<tr>
		    <td align=middle vAlign="top"><a title="<% =TVBalloon04 %>"><img src="<% =TVImage04 %>"></a></td>
		</tr>
	</table>
    </TD> 

  </TR>

</TABLE>


    <%
    WriteIndexPageFooter

END IF




' ------------------------
   SUB BuildSkiYearDrop
' ------------------------

%>
  <select name='SkiYear'><%

	' --- Query Ski Year Table for all instances that also exist in Raw Scores table ---
	SET rsSelectFields=Server.CreateObject("ADODB.recordset")
        sSQL = "SELECT * FROM " &SkiYearTableName&" AS SY" 
	sSQL = sSQL + " ORDER BY SY.SkiYearID DESC"
        rsSelectFields.open sSQL, SConnectionToTRATable


          IF LCASE(sSkiYear) = "all" THEN
	    response.write("<option value =""All"" selected>All Years</option>")
	  ELSE
	    response.write("<option value =""All"">All Years</option>")
	  END IF

        DO WHILE Not rsSelectFields.EOF

	  IF TRIM(rsSelectFields("SkiYearID")) = TRIM(sSkiYear) THEN
            Response.Write("<option value =""" & rsSelectFields("SkiYearID") &""" selected>")
            Response.Write(rsSelectFields("SkiYearName"))
	    sSkiYearName=rsSelectFields("SkiYearName")
            Response.Write("</option><br>")
          ELSE
            Response.Write("<option value =""" & rsSelectFields("SkiYearID") &""">")
            Response.Write(rsSelectFields("SkiYearName"))
            Response.Write("</option><br>")
          END IF
          rsSelectFields.MoveNext
        LOOP
        rsSelectFields.Close
        %>
  </select>
<%
END SUB


' ------------------------
   SUB BuildYearDrop
' ------------------------


%>
  <select name='Year'><%

	' --- Query Ski Year Table for all instances that also exist in Raw Scores table ---
	SET rsSelectFields=Server.CreateObject("ADODB.recordset")
        sSQL = "SELECT * FROM " &SkiYearTableName&" AS SY" 
	sSQL = sSQL + " WHERE SY.SkiYearID<>'1'"
	sSQL = sSQL + " ORDER BY SY.SkiYearID DESC"
        rsSelectFields.open sSQL, SConnectionToTRATable



        DO WHILE Not rsSelectFields.EOF

	  IF RIGHT(TRIM(rsSelectFields("SkiYearName")),2) = TRIM(YearSELECTed) THEN
            Response.Write("<option value =""" & RIGHT(rsSelectFields("SkiYearName"),2) &""" selected>")
            Response.Write(RIGHT(rsSelectFields("SkiYearName"),4))
      	    sSkiYearName=rsSelectFields("SkiYearName")
            Response.Write("</option><br>")
          ELSE
            Response.Write("<option value =""" & RIGHT(rsSelectFields("SkiYearName"),2) &""">")
            Response.Write(RIGHT(rsSelectFields("SkiYearName"),4))
            Response.Write("</option><br>")
          END IF
          rsSelectFields.MoveNext
        LOOP
          IF LCASE(YearSELECTed) = "all" THEN
	    response.write("<option value =""All"" selected>All Years</option>")
	  ELSE
	    response.write("<option value =""All"">All Years</option>")
	  END IF

        rsSelectFields.Close
        %>
  </select>
<%
END SUB





%>




