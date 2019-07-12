<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->
<%


Dim sAction
Dim sDiv
Dim sSYID
Dim sEditFld
Dim TempEOF
Dim sRecordSet
Dim UsingSQL
Dim ThisPage
Dim sValues
Dim TempFN
Dim sSptsGrpID
Dim TempDivTableName

IF request("SptsGrpID") <>"" THEN session("SptsGrpID") = UCASE(request("SptsGrpID"))

IF session("SptsGrpID") = "TEST" THEN
	TempDivTableName = "USAWSRank.DivTest"
ELSE
	IF session("SptsGrpID") = "AWS" OR session("SptsGrpID") = "NCW" THEN
		TempDivTableName = "USAWSRank.Division"
	ELSEIF session("SptsGrpID")="USW" OR session("SptsGrpID")="HYD" OR session("SptsGrpID")="AKA" THEN
		TempDivTableName = "USAWSRank.DivisionOther"
	ELSE
		Session.contents.remove("SptsGrpID")
		response.redirect("/rankings/default.asp")
	END IF
END IF


' --- Tests the authority of this person to be in this module ---
IF Session("SptsGrpID")<>Session("UserSptsGrpID") AND Session("adminmenulevel")<50 THEN
	response.redirect("/rankings/tools.asp?svar=reject")
END IF


IF request("search") = "1" THEN
  session("SearchFilter") = "1"
  session("SkiYearFilter") = request("SkiYear")
  session("DivCodeFilter") = request("DivCode")
  session("EditFldFilter") = request("EditFld")
END IF
IF request("search") = "clear" THEN
  session.contents.remove("SearchFilter")
  session.contents.remove("SkiYearFilter")
  session.contents.remove("DivCodeFilter")
  session("EditFldFilter") = "ALL"
END IF

sAction = trim(Request("action"))
IF left(sAction,7) = "Add New" then sAction = "addrec"
IF sAction = "Search" then sAction = "listrec"
	
Dim currentPage, rowCount, i
currentPage = TRIM(Request("currentPage"))
IF currentPage = "" THEN currentPage = 1

sDiv = trim(Request("div"))

sSYID = trim(Request("syid"))
IF sSYID = "" then sSYID = request("SkiYear")

sEditFld = session("EditFldFilter")
IF sEditFld = "" then sEditFld = "ALL"

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
IF sAction = "" THEN sAction = "listrec"

WriteHeader

SELECT CASE LCASE(sAction)

  CASE "listrec"
    WriteHeaders "List Divisions for Table: " & TempDivTableName
    ListRecords

  CASE "addrec"
		IF sDiv = "" THEN sDiv = sqlclean(ucase(trim(request("DivCode"))))
    WriteHeaders "Add Record:  Div Code = "&sDiv&" for SYID = "&sSYID&" and SptsGrpID = '"&Session("SptsGrpID")&"' IN "&TempDivTableName
    IF len(sDiv) <> 2 THEN
			%><br><H2><center><font color="red">
			You must Specify a two-character Division Code for the<BR>
			Division that you wish to have added, in the box below.
			<BR></font></center></H2><% 	
    	Listrecords
	  ELSE    	
			ChooseSQL "SELECT * FROM " & TempDivTableName & " WHERE (SptsGrpID='"&Session("SptsGrpID")&"' AND Div='"&sDiv&"' AND SkiYearID="&sSYID&")"
			IF rs.EOF THEN TempEOF = "Y" ELSE TempEOF = "N"
			rs.close: set rs = nothing
			IF TempEOF = "N" THEN
				%><br><H2><center><font color="red">
				The Division Code that you requested to be<br>
				added already exists for this Ski Year.<br>
				<BR></font></center></H2><%
	    	Listrecords
			ELSE
				AddRecord
				ShowEditor
			END IF
  	END IF

  CASE "editrec"
    WriteHeaders "Edit Division Record:  Div Code = "&sDiv&" for SYID = "&sSYID&" and Sports Discipline = "&session("SptsGrpID")&" in " & TempDivTableName
    ShowEditor

  CASE "saverec"
    WriteHeaders "Record Saved :  Div Code = " & sDiv & " for SYID = " & sSYID & " and Sports Discipline = "&session("SptsGrpID")&" in " & TempDivTableName
    SaveRec

  CASE "delrec"
    WriteHeaders "Delete Division Record:  Div Code = "&sDiv&" for SYID = "&sSYID&" and Sports Discipline = "&session("SptsGrpID")&" from " & TempDivTableName
    DeleteRec

END SELECT

WriteFooter



' ------------------------
  SUB WriteHeaders(sTitle)
' ------------------------

' Write Headers for Page
%>
<TABLE style="scores" BORDER="0" CELLPADDING="6" CELLSPACING="0" WIDTH="100%" BGCOLOR="<%=HQSiteColor2%>" BORDERCOLOR="#C0C0C0" BORDERCOLORDARK="#C0C0C0" BORDERCOLORLIGHT="#C0C0C0" >
<TR>
<TD ALIGN="CENTER"><Font COLOR="#FFFFFF" FACE="<%=font1%>" SIZE="2"><B><% Response.Write(sTitle) %></font></B></FONT></TD>
</TR>
</TABLE>
<BR>

<form action="/rankings/defaultHQ.asp" >
  <center><input type=submit value="Return to Main Menu" method="post"></center>
</form><%

END SUB



' ----------------
  SUB ListRecords
' ----------------

'  Lists the table Records

%> 
<form action="/rankings/editdivisions.asp" method="post">
<input type="hidden" name="search" value="1">

<TABLE border=1 align="center" bgcolor="<%=TableColor1%>" class="innertable" width=60%>
<TR>
  <th ALIGN="center" bgcolor="<%=HQSiteColor2%>"><FONT COlOR="#FFFFFF" FACE="<%=font1%>" SIZE="<%=fontsize1%>">Ski Year</FONT></th>
  <th ALIGN="center" bgcolor="<%=HQSiteColor2%>"><FONT COlOR="#FFFFFF" FACE="<%=font1%>" SIZE="<%=fontsize1%>">Division Code</FONT></th>
  <th ALIGN="center" bgcolor="<%=HQSiteColor2%>"><FONT COlOR="#FFFFFF" FACE="<%=font1%>" SIZE="<%=fontsize1%>">Display & Edit</FONT></th>
  <th ALIGN="center" bgcolor="<%=HQSiteColor2%>"><FONT COlOR="#FFFFFF" FACE="<%=font1%>" SIZE="<%=fontsize1%>">Sports Discipline</FONT></th>
</TR>

<TR>
<TD ALIGN="center">
<%


SkiYearDropBuild





%>
</TD>
<TD ALIGN="center">
  <input type="text" name="DivCode" value="<%=Session("DivCodeFilter")%>" size=2></input>
</TD>
<%

' ------------   Builds Edit/Display Fields Drop Down list -----------------

'	WriteDebugSQL("EditFld = " & Session("EditFldFilter"))


%>
  <TD ALIGN="center"><Center>
	<select name="EditFld">
		<option value="ALL"<%IF sEditFld = "ALL" THEN Response.Write(" selected ")%>>All Fields</option>
		<option value="PCT"<%IF sEditFld = "PCT" THEN Response.Write(" selected ")%>>Rank Level Percentiles</option>
		<option value="EQD"<%IF sEditFld = "EQD" THEN Response.Write(" selected ")%>>Equivalent Divisions</option>
		<option value="NOT"<%IF sEditFld = "NOT" THEN Response.Write(" selected ")%>>All but EquivDiv & Pcts</option>
    	</select>
  </TD>

<TD ALIGN="center"><FONT COlOR="#000000" FACE="<%=font1%>" SIZE="<%=fontsize1%>"><%=Session("SptsGrpID")%></FONT></TD>


</TR>

</table>

<br>
<TABLE width="<%=TourTableWidth%>" align=center>
<TR>
  <TD align=center>
	<input type="submit" style="width:15em" name="action" value="Search">
  </TD>
  <TD align=center>	
	<input type="submit" style="width:15em" name="action" value="Add New Div Record">
  </TD>
</form>

<form action="/rankings/editdivisions.asp" method="post">
  <TD align=center>	
	<input type="hidden" name="search" value="clear">
	<input type="submit" style="width:15em" value="Reset Search Filters">
  </TD>
</form>
</TR>
</TABLE>
<%






IF session("SearchFilter") = "1" THEN
  sSQL = "SELECT * FROM " & TempDivTableName & " WHERE 1=1"

  IF session("SkiYearFilter") <> "" THEN
    sSQL = sSQL + " AND SkiYearID = " & SQLClean(Session("SkiYearFilter")) & ""
  END IF

  IF session("DivCodeFilter") <> "" THEN
    	sSQL = sSQL + " AND Div = '" & SQLClean(Session("DivCodeFilter")) & "'"
  END IF

  sSQL = sSQL + " AND SptsGrpID='"&Session("SptsGrpID")&"'"

  sSQL = sSQL + " order by SkiYearID, Div"


'markdebug(sSQL)

  ChoosePagesSQL sSQL,currentPage, 60

ELSE

  sSQL = "SELECT " & TempDivTableName & ".* FROM " & TempDivTableName & " JOIN " & SkiYearTableName & " on "&TempDivTableName&".SkiYearID = "&SkiYearTableName&".SkiYearID WHERE "&SkiYearTableName&".DefaultYear = '1' AND SptsGrpID='"&session("SptsGrpID")&"' ORDER BY Div"

  ChoosePagesSQL sSQL,currentPage, 60

END IF


rowCount = 0

%>
&nbsp;<BR>&nbsp;<BR>
<TABLE BORDER="1" CELLPADDING="3" bgcolor="<%=TableColor1%>" CELLSPACING="0" WIDTH="100%" >
<TR>
<TD ALIGN="Center" bgcolor="<%=HQSiteColor2%>"><FONT COlOR="#FFFFFF" FACE="<%=font1%>" SIZE="<%=fontsize1%>">Delete</FONT></TD>
<TD ALIGN="Center" bgcolor="<%=HQSiteColor2%>"><FONT COlOR="#FFFFFF" FACE="<%=font1%>" SIZE="<%=fontsize1%>">Edit</FONT></TD>
<%


FOR i = 0 TO rs.fields.count - 1

	TempFN = rs.fields(i).name
	j = 0
	IF i < 6 THEN j = 1
	IF sEditFld = "PCT" AND lcase(left(TempFN,8)) = "percent_" THEN j = 1
	IF sEditFld = "EQD" AND (lcase(left(TempFN,5)) = "sl_ed" or lcase(left(TempFN,5)) = "ju_ed") THEN j = 1
	IF sEditFld = "NOT" AND lcase(left(TempFN,8)) <> "percent_" AND lcase(left(TempFN,5)) <> "sl_ed" AND lcase(left(TempFN,5)) <> "ju_ed" THEN j = 1
	IF sEditFld = "ALL" THEN j = 1

'	IF TempFN = "ZBSConversion" THEN j = 0
	
	IF j = 1 THEN  %>
    		<TD ALIGN="Center" vAlign="top" bgcolor="<%=HQSiteColor2%>" nowrap>
		  <FONT COlOR="#FFFFFF" FACE="<%=font1%>" SIZE="<%=fontsize1%>"><% Response.Write(Rs.Fields(i).name) %></FONT>
		</TD><%
	END IF

NEXT

%>
</TR>
<%

' --------------  Display table data here with paging --------------------------

DO WHILE NOT rs.eof

	IF rowCount = rs.PageSize THEN EXIT DO

	%>
 	<TR>
	<TD ALIGN="Left" vAlign="top"><FONT COlOR="#000000" FACE="<%=font1%>" SIZE="<%=fontsize1%>"><% WriteLink "?action=delrec&div=" & rs.fields(0).Value & "&syid=" & rs.fields(1).Value,"Delete","" %></FONT></TD>
	<TD ALIGN="Left" vAlign="top"><FONT COlOR="#000000" FACE="<%=font1%>" SIZE="<%=fontsize1%>"><% WriteLink "?action=editrec&div=" & rs.fields(0).Value & "&syid=" & rs.fields(1).Value,"Edit","" %></FONT></TD>
	<%

	FOR i = 0 TO rs.fields.count - 1
	
		TempFN = rs.fields(i).name
		j = 0
		IF i < 6 THEN j = 1
		IF sEditFld = "PCT" AND lcase(left(TempFN,8)) = "percent_" THEN j = 1
		IF sEditFld = "EQD" AND (lcase(left(TempFN,5)) = "sl_ed" or lcase(left(TempFN,5)) = "ju_ed") THEN j = 1
		IF sEditFld = "NOT" AND lcase(left(TempFN,8)) <> "percent_" AND lcase(left(TempFN,5)) <> "sl_ed" AND lcase(left(TempFN,5)) <> "ju_ed" THEN j = 1
		IF sEditFld = "ALL" THEN j = 1

'		IF TempFN = "ZBSConversion" THEN j = 0
	
		IF j = 1 THEN
		
		%><TD ALIGN="Left" vAlign="top" nowrap><FONT COlOR="#000000" FACE="<%=font1%>" SIZE="<%=fontsize1%>">&nbsp;<%

		IF UCASE(TRIM(rs.Fields(i).name)) = "SKIYEARID" THEN
		    Set rsSELECTFields=Server.CreateObject("ADODB.recordset")
    
		    sSQL = "SELECT SkiYearName FROM " & SkiYearTableName & " WHERE SkiYearID = " & rs.Fields(i).Value
		    rsSELECTFields.open sSQL, SConnectionToTRATable
  
		    IF NOT rsSELECTFields.EOF THEN
		        Response.Write(rsSELECTFields("SkiYearName"))
    		    ELSE
		        Response.Write(rs.Fields(i).Value)
		    END IF

		    rsSELECTFields.Close
		    set rsSELECTFields = nothing

		ELSE
		    IF isnull(rs.Fields(i).value) THEN
			response.write ("&nbsp;")
    		    ELSE
			Response.Write(trim(Rs.Fields(i).Value)) 
		    END IF
		END IF

		%>&nbsp;</FONT></TD><%

	END IF

	NEXT

		%>
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

rs.close
set rs = nothing

END SUB


' -----------------------
   SUB SkiYearDropBuild
' -----------------------

' ------------   Builds Ski Year Drop Down list ----------------- %>

<SELECT name='SkiYear'>
<%

ChooseSQL("SELECT * FROM " & SkiYearTableName)

DO WHILE not rs.eof
  response.write("<option value =""" & rs("SkiYearID") & """")

  IF trim(rs("SkiYearID")) = session("SkiYearFilter") THEN
    response.write(" SELECTed")
  END IF

  IF session("SkiYearFilter") = "" and rs("DefaultYear") THEN
    response.write(" SELECTed")
  END IF

  response.write(">")
  response.write(rs("SkiYearName"))
  response.write("</option><br>")
  rs.movenext
LOOP

%>
</SELECT><%

END SUB



' --------------
  SUB ShowEditor
' --------------

WriteButton "?action=listrec","No Change - Go To Div List","<BR>"

ChooseSQL "SELECT * FROM " & TempDivTableName & " WHERE (Div='"&sDiv&"' AND SkiYearID="&sSYID&" AND SptsGrpID='"&session("SptsGrpID")&"')"



%>
<FORM METHOD="POST" ACTION="<% Response.Write(ThisPage) %>?action=saverec">
<TABLE class="innertable" BORDER="1" ALIGN=center >

<%

' *** Important -- first two fields are Division Code and Ski Year ID code.
' *** These two fields serve as the record key, and hence are NOT editable.

FOR i = 2 TO rs.fields.count - 1

	TempFN = rs.fields(i).name
	j = 0
	IF i < 6 THEN j = 1
	IF sEditFld = "PCT" AND lcase(left(TempFN,8)) = "percent_" THEN j = 1
	IF sEditFld = "EQD" AND (lcase(left(TempFN,5)) = "sl_ed" or lcase(left(TempFN,5)) = "ju_ed") THEN j = 1
	IF sEditFld = "NOT" AND lcase(left(TempFN,8)) <> "percent_" AND lcase(left(TempFN,5)) <> "sl_ed" AND lcase(left(TempFN,5)) <> "ju_ed" THEN j = 1
	IF sEditFld = "ALL" THEN j = 1

'	IF TempFN = "ZBSConversion" THEN j = 0
	
	IF j = 1 THEN
		%>
		<TR>
		  <TD ALIGN="Left" bgcolor="<%=HQSiteColor2%>" width="150px"><Font COLOR="#FFFFFF" FACE="<%=font1%>" SIZE="<%=fontsize1%>"><B><% Response.Write(Rs.Fields(i).name) %></B></FONT></TD><%


		IF TempFN="SptsGrpID" AND Session("SptsGrpID")<>"AWS" THEN 	' --- If NOT AWSA --- %>
		  <TD ALIGN="Left">
			<input type="hidden" name="SptsGrpID" value="<%=Session("SptsGrpID")%>">
			<font size="2"><%=Session("SptsGrpID") %></font>
		  </TD><%
		ELSE %>
			<TD ALIGN="Left" width="100px"><Font COLOR="#000000" FACE="<%=font1%>" SIZE="<%=fontsize1%>"><% WriteType i %></FONT></TD><%
		END IF %>
		</TR>
		<%
	END IF

NEXT

%>
</TABLE><BR>
<TABLE BORDER="0" ALIGN=center WIDTH=30% CELLPADDING="3" CELLSPACING="0">
<TR>
<TD ALIGN="center"><input type="submit" style="width:9em" value="Save">
    <input type="hidden" name="Div" value="<%=sDiv%>">
    <input type="hidden" name="syid" value="<%=sSYID%>">
</TD>
<TD ALIGN="center"><input type="reset" style="width:9em" value="Reset"></TD>
</TR>
</TABLE>

</FORM>
<%

rs.close
set rs = nothing

END SUB



' -------------------------------
  SUB SaveRec
' -------------------------------


'Save the record to the table'

ChooseSQL "SELECT * FROM " & TempDivTableName & " WHERE (Div='"&sDiv&"' AND SkiYearID="&sSYID&" AND SptsGrpID='"&session("SptsGrpID")&"')"

	sSQL = "UPDATE " & TempDivTableName & " SET "

  	FOR i = 2 TO rs.fields.count - 1

		IF Request.Form(rs.fields(i).name) <> "" THEN

			IF RIGHT(sSQL,1) <> "," and RIGHT(sSQL,1) <> " " THEN sSQL = sSQL + ","

			sSQL = sSQL + rs.fields(i).name
			sSQL = sSQL + "='" + sqlclean(Request.Form(rs.fields(i).name)) + "'"

		END IF

	NEXT      

rs.close
set rs = nothing

sSQL = sSQL + " WHERE Div='" & sDiv & "' AND SkiYearID=" &sSYID&" AND SptsGrpID='"&session("SptsGrpID")&"'"

OpenCon
' WriteDebugSQL(sSQL)
con.execute(sSQL)
WriteLog(date() &"  "& time() &"   Division Record Updated - "& sSQL)
CloseCon

%>
<center><FONT COlOR="red" FACE="<%=font1%>" SIZE="<%=fontsize3%>"><b><i>Your updated record has been saved.</I></b></font></center>
<BR>
<%
WriteButton "?action=listrec","Click here to continue.","<BR>"

END SUB



'------------------
  SUB AddRecord
'------------------

' First we check for existence of a new Specified Division Code
' If not found then go ahead and add it, otherwise don't.
' Then the Editing window will be presented -- by mainline CASE.

 		sSQL = "INSERT INTO "&TempDivTableName&" (DIV, SkiYearID, SptsGrpID) VALUES ('"&sDiv&"', "&sSYID&", '"&Session("SptsGrpID")&"')"
		OpenCon
'  	WriteDebugSQL(sSQL)
		con.execute(sSQL)
		Closecon
		WriteLog(date() &"  "& time() &"   New Division Has Been Added - "& sSQL)

		%><br>
		<center><Font COLOR="#FFFFFF" FACE="<%=font1%>" SIZE="2"><b>
		A new record for the requested Division Code has been added.&nbsp; The specifications for <br>
		that new Division need to be supplied into the Editing form presented below.
		</b></font><BR>&nbsp;<BR><% 	

session.contents.remove("DivCodeFilter")	

END SUB



'-------------------
  SUB DeleteRec
'-------------------

ChooseSQL "SELECT * FROM "&TempDivTableName&" WHERE Div='" & sDiv & "' AND SkiYearID="&sSYID&" AND SptsGrpID='"&session("SptsGrpID")&"'" 

IF LCASE(Request("confirm")) = "yes" THEN
    'delete the record'
    'WriteLog(date() &"  "& time() &"  Division Table Record " & rs("Div") & " for SY=" & rs("SkiYearID") & " (" & rs("Div_Name") & ") has been deleted.")


    IF isrecordsetempty = false THEN
        rs.movefirst
        rs.delete
        rs.UPDATEBatch 3
    END IF

		rs.close
		set rs = nothing
    
    %>
    <center><Font COLOR="#FFFFFF" FACE="<%=font1%>" SIZE="2"><I><b>The record has been deleted.</b></I></font></center>
    <BR>&nbsp;<BR>
    <%

END IF

WriteButton "?action=listrec","Return To Division List","<BR><BR>"

IF LCASE(Request("confirm")) = "" THEN
%>  <br><br>
    <center>
     <Font COLOR="#FFFFFF" FACE="<%=font1%>" SIZE="2">
    Type the word "YES" IF you are sure you wish to delete this record. </font>
    <br>
    <Font COLOR="red" FACE="<%=font1%>" SIZE="2">
    Note: Scores which rely on this division may be affected.
    </font>
    <br><br>
    <form action="/rankings/editdivisions.asp" method="post"> 
    <input type="hidden" name="action" value="delrec">
    <input type="hidden" name="Div" value="<%=sDiv%>">
    <input type="hidden" name="syid" value="<%=sSYID%>">
    <input type="text" name="confirm" size="5">
    <input type="submit" value="Confirm Deletion?">
    </form>

<%
    WriteButton "?action=listrec","No - do not delete the record","<BR><BR>"

END IF

IF LCASE(Request("confirm")) <> "yes" and LCASE(Request("confirm")) <> "" THEN
     %>  <br><br>
         The record was NOT deleted.
         <br><br>
     <%
END IF

END SUB



'---------------------
  Function GetCheckValue(i)
'---------------------

    IF LCASE(sAction) = "editrec" THEN
        IF rs.fields(i).value = "1" THEN
            GetCheckValue = "checked"
        ELSE
            GetCheckValue = ""
        END IF
    ELSE
            GetCheckValue = ""
    END IF


'    IF LCASE(sAction) = "editrec" THEN
'       IF rs.fields(i).value = 0 or rs.fields(i).value = "" THEN
'            Response.Write("")
'        ELSE
'            Response.Write("1")
'        END IF
'    ELSE
'        Response.Write("")
'    END IF

END Function



'---------------------
  Function GetValue(i)
'---------------------

    IF LCASE(sAction) = "editrec" THEN
        GetValue = rs.fields(i).value
    ELSE
        GetValue = ""
    END IF

End Function



'---------------------
  SUB GetFieldValue(i)
'---------------------
    IF LCASE(sAction) = "editrec" THEN
        Response.Write(rs.fields(i).value)
    ELSE
        Response.Write("")
    END IF

END SUB



'---------------------
SUB WriteType(I)
'---------------------

SELECT CASE ucase(Rs.Fields(i).name)

	CASE "ID" %>
	   	<input type="hidden" name="id" value="<% Response.Write(sID) %>"> Auto Number<% 
		IF sid = 0 THEN 
     			response.write("(new)")
	   	ELSE
     			response.write(sID)
   	END IF

	CASE "SEX" %>
		<SELECT name="Sex">
		<option value="M" <%IF GetValue(i) = "M" THEN Response.Write("SELECTed")%>>Male</option>
		<option value="F" <%IF GetValue(i) = "F" THEN Response.Write("SELECTed")%>>Female</option>
		</SELECT>
		<%

	CASE "SKIYEARID" 
		response.write("  <SELECT name=""SkiYearID"">   ")

		set rsSELECTFields=Server.CreateObject("ADODB.recordset")
    
    		sSQL = "SELECT * FROM " & SkiYearTableName
		rsSELECTFields.open sSQL, SConnectionToTRATable
  
    		DO WHILE not rsSELECTFields.eof
      			response.write("<option value =""" & rsSELECTFields("SkiYearID") & """")

			IF trim(rsSELECTFields("SkiYearID")) = trim(GetValue(i)) THEN
				response.write(" SELECTed")
			END IF

			IF GetValue(i) = "" and rsSELECTFields("DefaultYear") THEN
				response.write(" SELECTed")
			END IF

			response.write(">")
			response.write(rsSELECTFields("SkiYearName"))
			response.write("</option><br>")

			rsSELECTFields.movenext
		LOOP

		rsSELECTFields.close
		set rsSELECTFields = nothing
        
		response.write("  </SELECT>  ")

	CASE ELSE

		SELECT CASE Rs.Fields(i).type
			CASE 3 'primary key / auto number ?'
				%><input type="text" name="<% Response.Write(Rs.Fields(i).name) %>" SIZE="25" value="<% GetFieldValue i %>"><%
			CASE 20 'primary key / auto number ?'
				%><input type="text" name="<% Response.Write(Rs.Fields(i).name) %>" SIZE="25" value="<% GetFieldValue i %>"><%
			CASE 11 'boolean'
        			%><INPUT TYPE="checkbox" NAME="<% Response.Write(Rs.Fields(i).name) %>" VALUE="0"<% GetcheckValue i %>><%
			CASE 203 'memo'
        			%><TEXTAREA NAME="<% Response.Write(Rs.Fields(i).name) %>" ROWS="20" COLS="56"><% GetFieldValue i %></TEXTAREA><%
			CASE ELSE 'not handled by this function'
			        %><input type="text" name="<% Response.Write(Rs.Fields(i).name) %>" SIZE="25" value="<% GetFieldValue i %>"><%
		END SELECT

END SELECT 

END SUB



' --------------------
   SUB ChooseSQL(sSQL)
' --------------------

set rs=Server.CreateObject("ADODB.recordset")
' WriteDebugSQL(sSQL)
rs.open sSQL, sConnectionToTRATable, 3, 3

END SUB



' --------------------
   SUB WriteHeader
' --------------------

%><HTML><HEAD><TITLE>TRA Database Tool</TITLE></HEAD>

<BODY BGCOLOR="#FFFFFE" Text="#0A0D0A" LINK="#375AE2" VLINK="#36566D" ALINK="#3E85BB">
<style TYPE="text/css">
<!--  A:link {text-decoration: none; color:#375AE2}  A:visited {text-decoration: none; color:#375AE2}  A:active {text-decoration: none}   A:hover {text-decoration: ; color:#3E85BB; }-->
</style><%

END SUB


' --------------------
  SUB WriteFooter
' --------------------


%><hr>
<form action="/rankings/defaultHQ.asp" >
  <center><input type=submit value="Return to Main Menu" method="post"></center>
</form>
<hr>
</BODY></HTML><%

END SUB



' --------------------------
  Function IsRecordSetEmpty
' --------------------------

IF rs.bof = true and rs.eof = true THEN
    IsRecordSetEmpty = true
ELSE
    IsRecordSetEmpty = false
END IF

END FUNCTION



' ---------------------------------------
  SUB ChoosePagesSQL(sSQL,sStart, sSize)
' ---------------------------------------

  set rs=Server.CreateObject("ADODB.recordset")
  sqlstmt = sSQL
  rs.CursorType = 3
  rs.PageSize = cint(sSize)
'	WriteDebugSQL(sSQL)
  rs.open sqlstmt, SConnectionToTRATable
  IF isrecordsetempty = false THEN
    rs.AbsolutePage = cINT(sStart)
  END IF

END SUB



' ---------------------------------------
  SUB WriteLink(sParms,sDisplay,sBreak)
' ---------------------------------------

%>
<A HREF="<% Response.Write(ThisPage & sParms) %>"><% Response.Write(sDisplay) %></A><%

END SUB


' ---------------------------------------
  SUB WriteButton(sParms,sDisplay,sBreak)
' ---------------------------------------

%>
<form action="<%=ThisPage%><%=sParms%>">
  <center><input type=submit value="<%=sDisplay%>" method="post"></center>
</form>
<%

END SUB



' ---------------------------------------
    SUB DoCount(currentPage) 
' ---------------------------------------

h = rs.PageCount

IF h > 21 THEN
  IF currentpage - 10 > 1 THEN
    	Response.Write("... ")
  END IF

  FOR i = 10 TO 1 step -1
    IF currentpage - i > 0 THEN
      	Response.Write(" <a href=" & chr(34) & ThisPage & "?currentpage=" &  currentpage - i  & "&action=" & sAction & chr(34) & ">" & currentpage - i & "</a>")
    END IF
  NEXT

  Response.Write(" " & currentpage & " ")

  FOR i = 1 TO 10
   	IF currentpage + i <= h THEN
      		Response.Write(" <a href=" & chr(34) & ThisPage & "?currentpage=" &  currentpage + i  & "&action=" & sAction & chr(34) & ">" & currentpage + i & "</a>")
	END IF
  NEXT

  IF currentpage + 10 <= h THEN
    Response.Write(" ...")
  END IF

ELSE
  FOR i = 1 TO h
    Response.Write(" <a href=" & chr(34) & ThisPage & "?currentpage=" &  i  & "&action=" & sAction & chr(34) & ">" & i & "</a>")
  next

END IF

IF h = 0 THEN h = 1
	Response.Write("<BR><Small>Page " & currentPage & " of  " & h & "</SMALL></center><BR><BR>")
END SUB



' ---------------------------------------
   Function GetFieldTypeName(I)
' ---------------------------------------

SELECT CASE i
CASE 0
GetFieldTypeName = "Empty"
CASE 16
GetFieldTypeName = "TinyInt"
CASE 2
GetFieldTypeName = "SmallInt"
CASE 3
GetFieldTypeName = "Integer"
CASE 20
GetFieldTypeName = "BigInt"
CASE 17
GetFieldTypeName = "UnsignedTinyInt"
CASE 18
GetFieldTypeName = "UnsignedSmallInt"
CASE 19
GetFieldTypeName = "UnsignedInt"
CASE 21
GetFieldTypeName = "UnsignedBigInt"
CASE 4
GetFieldTypeName = "Single"
CASE 5
GetFieldTypeName = "Double"
CASE 6
GetFieldTypeName = "Currency"
CASE 14
GetFieldTypeName = "Decimal"
CASE 131
GetFieldTypeName = "Numeric"
CASE 11
GetFieldTypeName = "Boolean"
CASE 10
GetFieldTypeName = "Error"
CASE 132
GetFieldTypeName = "UserDefined"
CASE 12
GetFieldTypeName = "Variant"
CASE 9
GetFieldTypeName = "IDispatch"
CASE 13
GetFieldTypeName = "IUnknown"
CASE 72
GetFieldTypeName = "GUID"
CASE 7
GetFieldTypeName = "Date"
CASE 133
GetFieldTypeName = "DBDate"
CASE 134
GetFieldTypeName = "DBTime"
CASE 135
GetFieldTypeName = "DBTimeStamp"
CASE 8
GetFieldTypeName = "BSTR"
CASE 129
GetFieldTypeName = "Char"
CASE 200
GetFieldTypeName = "VarChar"
CASE 201
GetFieldTypeName = "LongVarChar"
CASE 130
GetFieldTypeName = "WChar"
CASE 202
GetFieldTypeName = "VarWChar"
CASE 203
GetFieldTypeName = "LongVarWChar"
CASE 128
GetFieldTypeName = "Binary"
CASE 204
GetFieldTypeName = "VarBinary"
CASE 205
GetFieldTypeName = "LongVarBinary"
END SELECT
End Function

Function GetFieldTypeCode(sTXT,sLen)
'I am not overly familar with this stuff'
'you may have to edit these values'
SELECT CASE sTXT
CASE "Empty"
GetFieldTypeCode = "Empty"
CASE "TinyInt"
GetFieldTypeCode = "TinyInt"
CASE "SmallInt"
GetFieldTypeCode = "SmallInt"
CASE "Integer"
GetFieldTypeCode = "Integer"
CASE "BigInt"
GetFieldTypeCode = "BigInt"
CASE "UnsignedTinyInt"
GetFieldTypeCode = "UnsignedTinyInt"
CASE "UnsignedSmallInt"
GetFieldTypeCode = "UnsignedSmallInt"
CASE "UnsignedInt"
GetFieldTypeCode = "UnsignedInt"
CASE "UnsignedBigInt"
GetFieldTypeCode = "UnsignedBigInt"
CASE "Single"
GetFieldTypeCode = "Single"
CASE "Double"
GetFieldTypeCode = "Double"
CASE "Currency"
GetFieldTypeCode = "Currency"
CASE "Decimal"
GetFieldTypeCode = "Decimal"
CASE "Numeric"
GetFieldTypeCode = "Numeric"
CASE "Boolean"
GetFieldTypeCode = "Boolean"
CASE "Error"
GetFieldTypeCode = "Error"
CASE "UserDefined"
GetFieldTypeCode = "UserDefined"
CASE "Variant"
GetFieldTypeCode = "Variant"
CASE "IDispatch"
GetFieldTypeCode = "IDispatch"
CASE "IUnknown"
GetFieldTypeCode = "IUnknown"
CASE "GUID"
GetFieldTypeCode = "GUID"
CASE "Date"
GetFieldTypeCode = "Date"
CASE "DBDate"
GetFieldTypeCode = "DBDate"
CASE "DBTime"
GetFieldTypeCode = "DBTime"
CASE "DBTimeStamp"
GetFieldTypeCode = "DBTimeStamp"
CASE "BSTR"
GetFieldTypeCode = "BSTR(" & sLen & ")"
CASE "Char"
GetFieldTypeCode = "Char(" & sLen & ")"
CASE "VarChar"
GetFieldTypeCode = "VarChar(" & sLen & ")"
CASE "LongVarChar"
GetFieldTypeCode = "LongVarChar(" & sLen & ")"
CASE "WChar"
GetFieldTypeCode = "WChar(" & sLen & ")"
CASE "VarWChar"
GetFieldTypeCode = "VarWChar(" & sLen & ")"
CASE "LongVarWChar"
GetFieldTypeCode = "LongVarWChar(" & sLen & ")"
CASE "Binary"
GetFieldTypeCode = "Binary(" & sLen & ")"
CASE "VarBinary"
GetFieldTypeCode = "VarBinary(" & sLen & ")"
CASE "LongVarBinary"
GetFieldTypeCode = "LongVarBinary"
CASE ELSE
GetFieldTypeCode = "IUnknown"
END SELECT
End Function

%>




