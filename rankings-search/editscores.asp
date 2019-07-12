<!--#include file="secure-settings.asp"-->
<!--#include file="tools_include.asp"-->
<%

' --- Defines table and cell formatting ---
DefineTRAStyles

Dim sAction, pvar, usgi, amlvl
Dim sRecordSet
Dim UsingSQL
Dim ThisPage
Dim sID
Dim sSkiYearName, sSkiYear
Dim ThisScoresTableName

' --- pvar parameter passed from menu to initiate the process ---
pvar=TRIM(Request("pvar"))
IF pvar<>"" THEN
	Session("SptsGrpID")=pvar
ELSE
	pvar=Session("SptsGrpID")
	IF pvar="" THEN
		Response.redirect("/rankings/DefaultHQ.asp")
	END IF
END IF


' --- Tests the authority of this person to be in this module ---
' --- Note that revised logic allows AWS users to act for NCW ---
usgi = Session("UserSptsGrpID"): amlvl = Session("adminmenulevel")
IF (pvar<>usgi AND (NOT pvar="NCW" AND usgi="AWS")) AND amlvl<50 THEN
	response.redirect("/rankings/tools.asp?svar=reject")
END IF


IF pvar="AWS" OR pvar="NCW" THEN
	ThisScoresTableName=RawScoresTableName
ELSE
	ThisScoresTableName=RawScoresOtherTableName
END IF

if request("search") = "0" then
  session("SearchFilter") = "0"
END IF

if request("search") = "1" then
  session("SearchFilter") = "1"
  session("TourIDFilter") = request("Tour_ID")
  session("MemberIDFilter") = request("Member_ID")
  session("LastNameFilter") = request("Last_Name")
  session("FirstNameFilter") = request("First_Name")
  session("DivisionFilter") = request("Division")
  session("EventFilter") = request("Event")
  session("ClassFilter") = request("Class")
end if
if request("search") = "clear" then
  session.contents.remove("SearchFilter")
  session.contents.remove("TourIDFilter")
  session.contents.remove("MemberIDFilter")
  session.contents.remove("LastNameFilter")
  session.contents.remove("FirstNameFilter")
  session.contents.remove("DivisionFilter")
  session.contents.remove("EventFilter")
  session.contents.remove("ClassFilter")
end if

sAction = trim(Request("action"))

Dim currentPage, rowCount, i
currentPage = TRIM(Request("currentPage"))
if currentPage = "" then currentPage = 1

sID = trim(Request("id"))
If sID = "" then sID = 0



' --- New added by Mark Crone 12/2/2007 ---
sSkiYear=TRIM(Request("SkiYear"))
IF sSkiYear="" THEN 
   ' --- If the Session variable is NOT set then get from list ---
   IF TRIM(Session("SkiYear"))="" THEN	
	' --- Get the first one from the SkiYearTable ---
        SET rsSelectFields=Server.CreateObject("ADODB.recordset")
	rsSelectFields.open "Select Top 1 * FROM "&SkiYearTableName&" WHERE SkiYearID<>'1' ORDER BY BeginDate DESC", SConnectionToTRATable
	IF NOT rsSelectFields.eof THEN 
		sSkiYear=rsSelectFields("SkiYearID")
		'markdebug("INSIDE sSkiYear="&sSkiYear)
	ELSE
		' --- Do nothing ---
	END IF
   ELSE  '  --- Make session the current setting ---
	sSkiYear=Session("SkiYear")
   END IF	
ELSE  ' --- Make the current the session variable ---
	Session("SkiYear")=sSkiYear
END IF 


'markdebug("sSkiYear="&sSkiYear)




ThisPage = Request.ServerVariables("SCRIPT_NAME")


'---- DataTypeEnum Values ----'
Const adEmpty = 0
Const adTinyInt = 16
Const adSmallInt = 2
Const adInteger = 3
Const adBigInt = 20
Const adUnsignedTinyInt = 17
Const adUnsignedSmallInt = 18
Const adUnsignedInt = 19
Const adUnsignedBigInt = 21
Const adSingle = 4
Const adDouble = 5
Const adCurrency = 6
Const adDecimal = 14
Const adNumeric = 131
Const adBoolean = 11
Const adError = 10
Const adUserDefined = 132
Const adVariant = 12
Const adIDispatch = 9
Const adIUnknown = 13
Const adGUID = 72
Const adDate = 7
Const adDBDate = 133
Const adDBTime = 134
Const adDBTimeStamp = 135
Const adBSTR = 8
Const adChar = 129
Const adVarChar = 200
Const adLongVarChar = 201
Const adWChar = 130
Const adVarWChar = 202
Const adLongVarWChar = 203
Const adBinary = 128
Const adVarBinary = 204
Const adLongVarBinary = 205


'-Main Page Code-----------------------------------'

WriteIndexPageHeader
%>
<TABLE width="<%=TourTableWidth%>" class="droptable" style="background-color:white">  <% '--- Outer table for background --- %>
<TR>
 <TD><%


IF sAction = "" then sAction = "listrec"
WriteHeader

SELECT CASE lcase(sAction)
  CASE "listrec"
    WriteHeaders "List Records "
    ListRecords
  CASE "editrec"
    WriteHeaders "Edit Record: " & sID & " in Scores Table"
    ShowEditor
  CASE "saverec"
    WriteHeaders "Record Saved "
    SaveRec
  CASE "delrec"
    WriteHeaders "Delete Record: " & sID & " from Scores Table"
    DeleteRec
END SELECT

WriteFooter

%>
 </TD>
</TR>
</TABLE><% '--- Outer table for background --- 

WriteIndexPageFooter


' --------------------------
   SUB WriteHeaders(sTitle)
' --------------------------
' Write Headers for DB Page

%>
<TABLE align=center>
<TR>
<TD ALIGN="Left"><Font SIZE="4"><B><% Response.Write(sTitle) %></B></FONT></TD>
</TR>
</TABLE>
<BR>

<%
End Sub


' ----------------------
  SUB ListRecords
' ----------------------
'  List DB Records

'Response.Write("<BR>")

%> 
<form action="/rankings/editscores.asp" method="post">
<input type="hidden" name="search" value="1">
<TABLE align=center class="innertable" width="<%=TourTableWidth%>">
<TR>
 <TH ALIGN="Left" vAlign="top"><Center><FONT SIZE="2">Tour ID</FONT></Center></TH>
 <TH ALIGN="Left" vAlign="top"><Center><FONT SIZE="2">Member ID</FONT></Center></TH>
 <TH ALIGN="Left" vAlign="top"><Center><FONT SIZE="2">Last Name</FONT></Center></TH>
 <TH ALIGN="Left" vAlign="top"><Center><FONT SIZE="2">First Name</FONT></Center></TH>
 <TH ALIGN="Left" vAlign="top"><Center><FONT SIZE="2">Division</FONT></Center></TH>
 <TH ALIGN="Left" vAlign="top"><Center><FONT SIZE="2">Event</FONT></Center></TH>
 <TH ALIGN="Left" vAlign="top"><Center><FONT SIZE="2">Class</FONT></Center></TH>
 <TH ALIGN="Left" vAlign="top"><Center><FONT SIZE="2">Ski Year</FONT></Center></TH>
</TR>

<TR>
 <TD ALIGN="Left" vAlign="top"><Center><FONT SIZE="2"><input type="text" name="Tour_ID" value="<%=Session("TourIDFilter")%>" size=9></input></FONT></Center></TD>
 <TD ALIGN="Left" vAlign="top"><Center><FONT SIZE="2"><input type="text" name="Member_ID" value="<%=Session("MemberIDFilter")%>" size=9></input></FONT></Center></TD>
 <TD ALIGN="Left" vAlign="top"><Center><FONT SIZE="2"><input type="text" name="Last_Name" value="<%=Session("LastNameFilter")%>" size=15></input></FONT></Center></TD>
 <TD ALIGN="Left" vAlign="top"><Center><FONT SIZE="2"><input type="text" name="First_Name" value="<%=Session("FirstNameFilter")%>" size=15></input></FONT></Center></TD>
 <TD ALIGN="Left" vAlign="top"><Center><FONT SIZE="2"><input type="text" name="Division" value="<%=Session("DivisionFilter")%>" size=5></input></FONT></Center></TD>
 <TD ALIGN="Left" vAlign="top"><Center><FONT SIZE="2"><input type="text" name="Event" value="<%=Session("EventFilter")%>" size=5></input></FONT></Center></TD>
 <TD ALIGN="Left" vAlign="top"><Center><FONT SIZE="2"><input type="text" name="Class" value="<%=Session("ClassFilter")%>" size=5></input></FONT></Center></TD>

 <td ALIGN="center" vAlign="top">
  <select name='SkiYear'><%

	' --- Query Ski Year Table for all instances that also exist in Raw Scores table ---
	SET rsSelectFields=Server.CreateObject("ADODB.recordset")
        sSQL = "SELECT * FROM " &SkiYearTableName&" AS SY" 
	sSQL = sSQL & " ORDER BY SY.SkiYearID DESC"
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
 </td>
</TR>
</table>

<br>
<table width=60% align=center>
<tr>
 <td align=center>
  <input type="submit" style="width:9em" value="Begin Search">
 </td>
   </form>
   <form action="/rankings/defaultHQ.asp" method="post">
 <td align=center>
   <input type="submit" style="width:9em" value="Main Menu">
 </td>
   </form>
   <form action="/rankings/editscores.asp" method="post">
 <td align=center>
  <input type="hidden" name="search" value="clear">
  <input type="submit" style="width:9em" value="Reset Filters">
 </td>
   </form>
</tr>
</table>
<%
OpenCon

' ---------------------------------------------------------------------
' -----------------  Begins Building Query  ---------------------------
' ---------------------------------------------------------------------


IF session("SearchFilter") = "1" THEN
  sSQL = "Select * from " & ThisScoresTableName
  sSQL = sSQL & " WHERE 1=1"

  IF session("TourIDFilter") <> "" then
    sSQL = sSQL & " AND tourid LIKE '%" & SQLClean(Session("TourIDFilter")) & "%'"
  END IF

  IF session("MemberIDFilter") <> "" then
    sSQL = sSQL & " AND memberid LIKE '%" & SQLClean(Session("MemberIDFilter")) & "%'"
  END IF

  IF session("LastNameFilter") <> "" then
    sSQL = sSQL & " AND lower(lname) LIKE '%" & SQLClean(lcase(Session("LastNameFilter"))) & "%'"
  END IF

  IF session("FirstNameFilter") <> "" then
    sSQL = sSQL & " AND lower(fname) LIKE '%" & SQLClean(lcase(Session("FirstNameFilter"))) & "%'"
  END IF

  IF session("DivisionFilter") <> "" then
    sSQL = sSQL & " AND lower(div) LIKE '%" & SQLClean(lcase(Session("DivisionFilter"))) & "%'"    
  END IF

  IF session("EventFilter") <> "" then
    sSQL = sSQL & " AND lower(event) LIKE '%" & SQLClean(lcase(Session("EventFilter"))) & "%'"
  END IF

  IF session("ClassFilter") <> "" then
    sSQL = sSQL & " AND lower(class) LIKE '%" & SQLClean(lcase(Session("ClassFilter"))) & "%'"
  END IF

  '--- New 12/2/2007 added by Mark Crone ---
  IF LCASE(sSkiYear) <> "all" AND LCASE(sSkiYear) <> "1" THEN 
	sSQL = sSQL & " AND LEFT(TourID,2) = '"&RIGHT(sSkiYearName,2)&"'" 
  ELSEIF LCASE(sSkiYear) = "1" THEN 
		Dim FirstDate, LastDate
		FirstDate = CDate(DateAdd("d", -365, DATE))
		LastDate = CDate(DateAdd("d", 1, DATE))
	sSQL = sSQL & " AND EndDate <='"&Date&"' AND EndDate>='"&FirstDate&"'"

  END IF

	sSQL = sSQL & " Order by EndDate, Event, Round"

	ChoosePagesSQL sSQL,currentPage, 40
	ShowDataHead
	ShowAllData
ELSE
	%>
	<center><font color=red>Input Search Criteria and Press 'Begin Search'</font></center>
	<br>
	<%
	
END IF

rowCount = 0

END SUB


' ---------------------
  SUB ShowDataHead
' ---------------------
%>
<br>
<font size="1"><center>NOTE:  Alt Score value is “Distance in Meters” for Jumping, or “Last Pass Buoy Count” for Slalom<br>&nbsp;</center></font>
<TABLE class="innertable" Align=center>
<TR>
<% If Session("UserLevel") >= Administration Then %>
<TH ALIGN="Center" vAlign="top"><FONT SIZE="1">Delete</FONT></TH>
<% End If %>
<TH ALIGN="Center" vAlign="top"><FONT SIZE="1">Edit</FONT></TH>
<TH ALIGN="Center" vAlign="top"><FONT SIZE="1">Member ID</FONT></TH>
<TH ALIGN="Center" vAlign="top"><FONT SIZE="1">Last Name</FONT></TH>
<TH ALIGN="Center" vAlign="top"><FONT SIZE="1">First Name</FONT></TH>
<TH ALIGN="Center" vAlign="top"><FONT SIZE="1">Team</FONT></TH>
<TH ALIGN="Center" vAlign="top"><FONT SIZE="1">Tour ID</FONT></TH>
<TH ALIGN="Center" vAlign="top"><FONT SIZE="1">End Date</FONT></TH>
<TH ALIGN="Center" vAlign="top"><FONT SIZE="1">Score</FONT></TH>
<TH ALIGN="Center" vAlign="top"><FONT SIZE="1">EV</FONT></TH>
<TH ALIGN="Center" vAlign="top"><FONT SIZE="1">Div</FONT></TH>
<TH ALIGN="Center" vAlign="top"><FONT SIZE="1">Rd</FONT></TH>
<TH ALIGN="Center" vAlign="top"><FONT SIZE="1">Cls</FONT></TH>
<TH ALIGN="Center" vAlign="top"><FONT SIZE="1">Alt Score</FONT></TH>
<TH ALIGN="Center" vAlign="top"><FONT SIZE="1">Ramp<br>/Line</FONT></TH>
<TH ALIGN="Center" vAlign="top"><FONT SIZE="1">Spd</FONT></TH>
<TH ALIGN="Center" vAlign="top"><FONT SIZE="1">Plc</FONT></TH>
<TH ALIGN="Center" vAlign="top"><FONT SIZE="1">Plc Pts</FONT></TH>
</TR>
<%
'add table data here with paging'

END SUB


' --------------------
  SUB ShowAllData
' --------------------

DO WHILE NOT rs.eof
	IF rowCount = rs.PageSize THEN exit DO
	%>
	<TR>
	<% If Session("UserLevel") >= Administration Then %>
		<TD ALIGN="Left" vAlign="top"><FONT SIZE="1">&nbsp;<% WriteLink "?action=delrec&id=" & rs.fields(0).Value,"Delete","" %>&nbsp;</FONT></TD>
	<% End If %>

	<TD ALIGN="Left" vAlign="top"><FONT SIZE="1">&nbsp;<% WriteLink "?action=editrec&id=" & rs.fields(0).Value,"Edit","" %>&nbsp;</FONT></TD>
	<TD ALIGN="Left" vAlign="top"><FONT SIZE="1">&nbsp;<% Response.Write(trim(rs("MemberID"))) %>&nbsp;</FONT></TD>
	<TD ALIGN="Left" vAlign="top"><FONT SIZE="1">&nbsp;<% Response.Write(trim(rs("LName"))) %>&nbsp;</FONT></TD>
	<TD ALIGN="Left" vAlign="top"><FONT SIZE="1">&nbsp;<% Response.Write(trim(rs("FName"))) %>&nbsp;</FONT></TD>
	<TD ALIGN="Left" vAlign="top"><FONT SIZE="1">&nbsp;<% Response.Write(trim(rs("Team"))&"/"&trim(rs("TeamStat"))) %>&nbsp;</FONT></TD>
	<TD ALIGN="Left" vAlign="top"><FONT SIZE="1">&nbsp;<% Response.Write(trim(rs("TourID"))) %>&nbsp;</FONT></TD>
	<TD ALIGN="Left" vAlign="top"><FONT SIZE="1">&nbsp;<% Response.Write(trim(rs("EndDate"))) %>&nbsp;</FONT></TD>
	<TD ALIGN="Left" vAlign="top"><FONT SIZE="1">&nbsp;<% Response.Write(trim(rs("Score"))) %>&nbsp;</FONT></TD>
	<TD ALIGN="Left" vAlign="top"><FONT SIZE="1">&nbsp;<% Response.Write(trim(rs("Event"))) %>&nbsp;</FONT></TD>
	<TD ALIGN="Left" vAlign="top"><FONT SIZE="1">&nbsp;<% Response.Write(trim(rs("Div"))) %>&nbsp;</FONT></TD>
	<TD ALIGN="Left" vAlign="top"><FONT SIZE="1">&nbsp;<% Response.Write(trim(rs("Round"))) %>&nbsp;</FONT></TD>
	<TD ALIGN="Left" vAlign="top"><FONT SIZE="1">&nbsp;<% Response.Write(trim(rs("Class"))) %>&nbsp;</FONT></TD>
	<TD ALIGN="Left" vAlign="top"><FONT SIZE="1">&nbsp;<% Response.Write(trim(rs("AltScore"))) %>&nbsp;</FONT></TD>
	<TD ALIGN="Left" vAlign="top"><FONT SIZE="1">&nbsp;<% Response.Write(trim(rs("Perf_Qual1"))) %>&nbsp;</FONT></TD>
	<TD ALIGN="Left" vAlign="top"><FONT SIZE="1">&nbsp;<% Response.Write(trim(rs("Perf_Qual2"))) %>&nbsp;</FONT></TD>
	<TD ALIGN="Left" vAlign="top"><FONT SIZE="1">&nbsp;<% Response.Write(trim(rs("Place"))) %>&nbsp;</FONT></TD>
	<TD ALIGN="Left" vAlign="top"><FONT SIZE="1">&nbsp;<% Response.Write(trim(rs("NSL_Placement_Points"))) %>&nbsp;</FONT></TD>
	</TR>

	<% 
	rowCount = rowCount + 1
	rs.movenext
LOOP

%>
</TABLE>
<br><br>
<%
DoCount currentPage
CloseCon


END SUB




' ----------------------
  SUB ShowEditor
' ----------------------

%>
<form action="/rankings/editscores.asp?action=listrec" method="post"> 
    <center><input type="submit" value="Return to Score Listing"></center>
</form><%

OpenCon
ChooseSQL "Select * from " & ThisScoresTableName & " where(id=" & sID & ")"
%>


<FORM METHOD="POST" ACTION="<% Response.Write(ThisPage) %>?action=saverec">
<TABLE class="innertable" align="center">
<%

For i = 0 to rs.fields.count - 1
   %>
   <TR>
	<TD ALIGN="Left" vAlign="Top"><Font SIZE="1"><B><% Response.Write(Rs.Fields(i).name) %></B></FONT><% 
	  IF rs.fields(i).name = "ALTSCORE" THEN %>
		<font size="1">&nbsp;&nbsp;NOTE:  AltScore is “Distance in Meters” <br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;or “Last Pass Buoy Count”</font><% 
	  END IF %>
	</TD>
	<TD ALIGN="Left" vAlign="Top"><Font SIZE="1"><% WriteType i %></FONT></TD>
   </TR>
   <%
NEXT

%>
</TABLE>
<BR>
<TABLE BORDER="0" width=100%>
<TR>
<TD ALIGN="right"><input type="submit" style="width:9em" value="Save"></TD>
<TD ALIGN="left"><input type="reset" style="width:9em" value="Reset"></TD>
</TR>
</TABLE>

</FORM>
<%
CloseCon
END SUB


' -------------------
   SUB SaveRec
' -------------------

'Save the record to the table'
OpenCon
ChooseSQL "Select * from " & ThisScoresTableName & " where(ID=" & sID & ")"
if sID = 0 then
    rs.addnew
else
    rs.movefirst
end if

For i = 1 to rs.fields.count - 1
    'set the field value'
      select case rs.fields(i).type
        case adBigInt
          rs.fields(i).value = csng(Request.Form(rs.fields(i).name))
        case adBoolean 
          if trim(Request.Form(rs.fields(i).name) = "") then
            rs.fields(i).value = null
          else
            rs.fields(i).value = True
          end if
        case adCurrency
          if Request.Form(rs.fields(i).name) <> "" then
            rs.fields(i).value = ccur(Request.Form(rs.fields(i).name))
          else 
            rs.fields(i).value = null
          end if
        case adDate,adDBDate,adDBTime,adDBTimeStamp
          if Request.Form(rs.fields(i).name) <> "" then
            rs.fields(i).value = cdate(Request.Form(rs.fields(i).name))
          else 
            rs.fields(i).value = null
          end if
        case adDecimal
          if Request.Form(rs.fields(i).name) <> "" then
             rs.fields(i).value = cdec(Request.Form(rs.fields(i).name))
          else 
            rs.fields(i).value = null
          end if
        case adDouble
          if Request.Form(rs.fields(i).name) <> "" then
            rs.fields(i).value = cdbl(Request.Form(rs.fields(i).name))
          else 
            rs.fields(i).value = null
          end if
        case adInteger
          if Request.Form(rs.fields(i).name) <> "" then
            rs.fields(i).value = cint(Request.Form(rs.fields(i).name))
          else 
            rs.fields(i).value = null
          end if
        case adSingle
          if Request.Form(rs.fields(i).name) <> "" then
            rs.fields(i).value = csng(Request.Form(rs.fields(i).name))
          else 
            rs.fields(i).value = null
          end if
        case else
          rs.fields(i).value = Request.Form(rs.fields(i).name)
     end select
next
rs.UpdateBatch 3
CloseCon
%>
<center><font color="red">Your record has been saved.</font></center><BR>
<%
WriteButton "?action=listrec","Click here to continue.","<BR>"

END SUB



' -------------------
   SUB DeleteRec
' -------------------

%>
<form action="/rankings/editscores.asp?action=listrec" method="post"> 
    <center><input type="submit" value="Return to Listing"></center>
</form><%

OpenCon
ChooseSQL "Select * from " & ThisScoresTableName & " where(ID=" & sID & ")"
if lcase(Request("confirm")) = "yes" then
    'delete the record'
    WriteLog(date() &"  "& time() &"  "& rs.fields(0).value & " " & rs.fields(1).value & " " & rs.fields(2).value & " has been deleted.")
    if isrecordsetempty = false then
        rs.movefirst
        rs.delete
        rs.UpdateBatch 3
    end if
    CloseCon
    %>
    The record was deleted.<BR>
    <%
end if
if lcase(Request("confirm")) = "" then
%>  <br>
    <center>Type the word "YES" if you are sure you wish to delete this record.</center>
    <br>
    <form action="/rankings/editscores.asp" method="post"> 
    <input type="hidden" name="action" value="delrec">
    <input type="hidden" name="id" value="<%=sid%>">
    <center><input type="text" name="confirm" size="5">
    <input type="submit" value="Confirm Deletion?"></center>
    </form>

<%


end if
if lcase(Request("confirm")) <> "yes" and lcase(Request("confirm")) <> "" then
     %>  <br><br>
         The record was NOT deleted.
         <br><br>
     <%
end if

End Sub

' ---------------------
  Sub GetCheckValue(i)
' ---------------------
    if lcase(sAction) = "editrec" then
        if rs.fields(i).value = 0 then
            Response.Write("")
        else
            Response.Write("checked")
        end if
    else
        Response.Write("")
    end if
End Sub


' ---------------------
  Sub GetFieldValue(i)
' ---------------------
    if lcase(sAction) = "editrec" then
        Response.Write(rs.fields(i).value)
    else
        Response.Write("")
    end if
End Sub

' ---------------------
  Sub WriteType(I)
' ---------------------
Select Case Rs.Fields(i).type
case 3 'primary key / auto number ?'
    if lcase(Rs.Fields(i).name) = "id" then
       if sid = 0 then
          %>
          <input type="hidden" name="id" value="<% Response.Write(sID) %>"> Auto Number (new)
          <%
       else
          %>
          <input type="hidden" name="id" value="<% Response.Write(sID) %>"> Auto Number (<% Response.Write(sID) %>)
          <%
       end if
    else
        %>
        <input type="text" name="<% Response.Write(Rs.Fields(i).name) %>" SIZE="25" value="<% GetFieldValue i %>">
        <%
    end if
case 20 'primary key / auto number ?'
    if lcase(Rs.Fields(i).name) = "id" then
       if sid = 0 then
          %>
          <input type="hidden" name="id" value="<% Response.Write(sID) %>"> Auto Number (new)
          <%
       else
          %>
          <input type="hidden" name="id" value="<% Response.Write(sID) %>"> Auto Number (<% Response.Write(sID) %>)
          <%
       end if
    else
        %>
        <input type="text" name="<% Response.Write(Rs.Fields(i).name) %>" SIZE="25" value="<% GetFieldValue i %>">
        <%
    end if
case 11 'boolean'
    %>
    <INPUT TYPE="checkbox" NAME="<% Response.Write(Rs.Fields(i).name) %>" VALUE="1" <% GetCheckValue i %>>
    <%
case 203 'memo'
    %>
    <TEXTAREA NAME="<% Response.Write(Rs.Fields(i).name) %>" ROWS="20" COLS="56"><% GetFieldValue i %></TEXTAREA>
    <%
case else 'not handled by this function'
    IF Rs.Fields(i).name<>"SptsGrpID" THEN
	%><input type="text" name="<% Response.Write(Rs.Fields(i).name) %>" SIZE="25" value="<% GetFieldValue i %>"><%
    ELSE  %>
          <input type="hidden" name="SptsGrpID" value="<%=pvar%>">
	  <font size="2"><%=pvar%></font><%
    END IF 	

End Select

END SUB



' --------------------------------
  SUB ChooseSQL(sSQL)
' --------------------------------
set rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, sConnectionToTRATable, 3, 3

END SUB



' -------------------
  SUB WriteHeader
' -------------------

%>
<HTML>
<HEAD><TITLE>TRA Database Tool</TITLE>
</HEAD>

<BODY BGCOLOR="#FFFFFE" Text="#0A0D0A" LINK="#375AE2" VLINK="#36566D" ALINK="#3E85BB">
<style TYPE="text/css">
<!--  A:link {text-decoration: none; color:#375AE2}  A:visited {text-decoration: none; color:#375AE2}  A:active {text-decoration: none}   A:hover {text-decoration: ; color:#3E85BB; }-->
</style>
<%
End Sub

' -------------------
  Sub WriteFooter
' -------------------
%>
<hr>
<form action="/rankings/defaultHQ.asp" method="post"> 
    <center><input type="submit" value="Return to Main Menu"></center>
</form>
<hr>
</BODY>
</HTML>
<%
End Sub

' ---------------------------
  Function IsRecordSetEmpty
' ---------------------------

if rs.bof = true and rs.eof = true then
    IsRecordSetEmpty = true
else
    IsRecordSetEmpty = false
end if
end Function

' ----------------------------------------
  Sub ChoosePagesSQL(sSQL,sStart, sSize)
' ----------------------------------------
  set rs=Server.CreateObject("ADODB.recordset")
  sqlstmt = sSQL
  rs.CursorType = 3
  rs.PageSize = cint(sSize)
  rs.open sqlstmt, SConnectionToTRATable
  if isrecordsetempty = false then
    rs.AbsolutePage = cINT(sStart)
  end if
End Sub

' ----------------------------------------
  Sub WriteLink(sParms,sDisplay,sBreak)
' ----------------------------------------
%><A HREF="<% Response.Write(ThisPage & sParms) %>"><% Response.Write(sDisplay) %></A><%
End Sub


' ---------------------------------------
  SUB WriteButton(sParms,sDisplay,sBreak)
' ---------------------------------------

%>
<form action="<%=ThisPage%><%=sParms%>">
  <center><input type=submit value="<%=sDisplay%>" method="post"></center>
</form>
<%

END SUB



' ----------------------------------------
  Sub DoCount(currentPage) 
' ----------------------------------------
h = rs.PageCount

if h > 21 then
  if currentpage - 10 > 1 then
    Response.Write("... ")
  end if
  for i = 10 to 1 step -1
    if currentpage - i > 0 then
      Response.Write(" <a href=" & chr(34) & ThisPage & "?currentpage=" &  currentpage - i  & "&SkiYear="&sSkiYear&"&action=" & sAction & chr(34) & ">" & currentpage - i & "</a>")
    end if
  next
  Response.Write(" " & currentpage & " ")
  for i = 1 to 10
   if currentpage + i <= h then
      Response.Write(" <a href=" & chr(34) & ThisPage & "?currentpage=" &  currentpage + i  & "&SkiYear="&sSkiYear&"&action=" & sAction & chr(34) & ">" & currentpage + i & "</a>")
    end if
  next
  if currentpage + 10 <= h then
    Response.Write(" ...")
  end if
else
  for i = 1 to h
    Response.Write(" <a href=" & chr(34) & ThisPage & "?currentpage=" &  i  & "&SkiYear="&sSkiYear&"&action=" & sAction & chr(34) & ">" & i & "</a>")
  next
end if

if h = 0 then h = 1
Response.Write("<BR><Small>Page " & currentPage & " of  " & h & "</SMALL></center><BR><BR>")
end sub

' ----------------------------------------
Function GetFieldTypeName(I)
' ----------------------------------------
select case i
case 0
GetFieldTypeName = "Empty"
case 16
GetFieldTypeName = "TinyInt"
case 2
GetFieldTypeName = "SmallInt"
case 3
GetFieldTypeName = "Integer"
case 20
GetFieldTypeName = "BigInt"
case 17
GetFieldTypeName = "UnsignedTinyInt"
case 18
GetFieldTypeName = "UnsignedSmallInt"
case 19
GetFieldTypeName = "UnsignedInt"
case 21
GetFieldTypeName = "UnsignedBigInt"
case 4
GetFieldTypeName = "Single"
case 5
GetFieldTypeName = "Double"
case 6
GetFieldTypeName = "Currency"
case 14
GetFieldTypeName = "Decimal"
case 131
GetFieldTypeName = "Numeric"
case 11
GetFieldTypeName = "Boolean"
case 10
GetFieldTypeName = "Error"
case 132
GetFieldTypeName = "UserDefined"
case 12
GetFieldTypeName = "Variant"
case 9
GetFieldTypeName = "IDispatch"
case 13
GetFieldTypeName = "IUnknown"
case 72
GetFieldTypeName = "GUID"
case 7
GetFieldTypeName = "Date"
case 133
GetFieldTypeName = "DBDate"
case 134
GetFieldTypeName = "DBTime"
case 135
GetFieldTypeName = "DBTimeStamp"
case 8
GetFieldTypeName = "BSTR"
case 129
GetFieldTypeName = "Char"
case 200
GetFieldTypeName = "VarChar"
case 201
GetFieldTypeName = "LongVarChar"
case 130
GetFieldTypeName = "WChar"
case 202
GetFieldTypeName = "VarWChar"
case 203
GetFieldTypeName = "LongVarWChar"
case 128
GetFieldTypeName = "Binary"
case 204
GetFieldTypeName = "VarBinary"
case 205
GetFieldTypeName = "LongVarBinary"
End Select
End Function

' ----------------------------------------
Function GetFieldTypeCode(sTXT,sLen)
' ----------------------------------------
'I am not overly familar with this stuff'
'you may have to edit these values'
select case sTXT
case "Empty"
GetFieldTypeCode = "Empty"
case "TinyInt"
GetFieldTypeCode = "TinyInt"
case "SmallInt"
GetFieldTypeCode = "SmallInt"
case "Integer"
GetFieldTypeCode = "Integer"
case "BigInt"
GetFieldTypeCode = "BigInt"
case "UnsignedTinyInt"
GetFieldTypeCode = "UnsignedTinyInt"
case "UnsignedSmallInt"
GetFieldTypeCode = "UnsignedSmallInt"
case "UnsignedInt"
GetFieldTypeCode = "UnsignedInt"
case "UnsignedBigInt"
GetFieldTypeCode = "UnsignedBigInt"
case "Single"
GetFieldTypeCode = "Single"
case "Double"
GetFieldTypeCode = "Double"
case "Currency"
GetFieldTypeCode = "Currency"
case "Decimal"
GetFieldTypeCode = "Decimal"
case "Numeric"
GetFieldTypeCode = "Numeric"
case "Boolean"
GetFieldTypeCode = "Boolean"
case "Error"
GetFieldTypeCode = "Error"
case "UserDefined"
GetFieldTypeCode = "UserDefined"
case "Variant"
GetFieldTypeCode = "Variant"
case "IDispatch"
GetFieldTypeCode = "IDispatch"
case "IUnknown"
GetFieldTypeCode = "IUnknown"
case "GUID"
GetFieldTypeCode = "GUID"
case "Date"
GetFieldTypeCode = "Date"
case "DBDate"
GetFieldTypeCode = "DBDate"
case "DBTime"
GetFieldTypeCode = "DBTime"
case "DBTimeStamp"
GetFieldTypeCode = "DBTimeStamp"
case "BSTR"
GetFieldTypeCode = "BSTR(" & sLen & ")"
case "Char"
GetFieldTypeCode = "Char(" & sLen & ")"
case "VarChar"
GetFieldTypeCode = "VarChar(" & sLen & ")"
case "LongVarChar"
GetFieldTypeCode = "LongVarChar(" & sLen & ")"
case "WChar"
GetFieldTypeCode = "WChar(" & sLen & ")"
case "VarWChar"
GetFieldTypeCode = "VarWChar(" & sLen & ")"
case "LongVarWChar"
GetFieldTypeCode = "LongVarWChar(" & sLen & ")"
case "Binary"
GetFieldTypeCode = "Binary(" & sLen & ")"
case "VarBinary"
GetFieldTypeCode = "VarBinary(" & sLen & ")"
case "LongVarBinary"
GetFieldTypeCode = "LongVarBinary"
case else
GetFieldTypeCode = "IUnknown"
End Select
End Function

%>




