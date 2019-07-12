<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->
<%

DefineTRAStyles




Dim currentPage, rowCount, i
Dim sAction
Dim sTeam
Dim sSYID
Dim sEditFld
Dim TempEOF
Dim sRecordSet
Dim UsingSQL
Dim ThisPage
Dim sValues
Dim TempFN
Dim sSptsGrpID
Dim ThisTempTableName
Dim ThisFieldID, sThisFieldID, KeyFieldName, KeyFieldShortName, KeyFieldName2, KeyCode, KeyCodeFilter
Dim LengthofID, sSortField
Dim DropFilterName1, DropFilterID1
Dim sLeagueSelected, sTourSelected, sTourName, sRegionSelected, sTourType, sUseForLCQScr
Dim ThisFileName

ThisFileName="EditLeagueTours.asp"

ThisFieldID="leagueid"
KeyFieldName="LeagueD"
KeyFieldShortName="LeagueID"
LengthofID=7

' --- KeyFieldName2 is the one the list gets sorted by on the listing ---
KeyFieldName2="LeagueName"
KeyCode="LeagueCode"
KeyCodeFilter="LeagueCodeFilter"

' --- If filter is defined it will show up in Filtering Option drop down
'DropFilterName1 = "NCWSA Region"
'DropFilterName2 = ""
'DropFilterName3 = ""
'DropFilterID1="REG"
'DropFilterID2=""
'DropFilterID3=""


'response.write(" <br>TOP request(sSptsGrpID)="& request("sSptsGrpID"))
'response.write(" <br>   Session(UserSptsGrpID)="& Session("UserSptsGrpID"))



IF request("sSptsGrpID") <>"" THEN session("sSptsGrpID") = UCASE(request("sSptsGrpID"))


IF session("sSptsGrpID") = "TEST" THEN
	ThisTempTableName = "USAWSRank.LeagueTours"
ELSE
	IF session("sSptsGrpID") = "AWS" OR session("sSptsGrpID") = "NCW" THEN
		ThisTempTableName = "usawsrank.LeagueTours"
	ELSEIF session("sSptsGrpID")="USW" OR session("sSptsGrpID")="HYD" OR session("sSptsGrpID")="AKA" THEN
		ThisTempTableName = "usawsrank.LeagueTours"
	ELSE
		Session.contents.remove("sSptsGrpID")
		response.redirect("/rankings/defaulthq.asp")
	END IF
END IF


' --- Tests the authority of this person to be in this module ---
IF Session("sSptsGrpID")<>Session("UserSptsGrpID") AND Session("adminmenulevel")<50 THEN
	response.redirect("/rankings/tools.asp?svar=reject")
END IF



IF request("search") = "1" THEN
  session("SearchFilter") = "1"
'  session("sSYIDFilter") = request("SkiYear")
  session("NewKeyCode") = request("KeyCode")
  session("EditFldFilter") = request("EditFld")
END IF

IF request("search") = "clear" THEN
  session.contents.remove("SearchFilter")
'  session.contents.remove("SkiYearFilter")
  session.contents.remove("NewKeyCode")
  session("EditFldFilter") = "ALL"
END IF

sAction = trim(Request("action"))
SELECT CASE left(sAction,7)
	CASE "Add New"
		sAction = "addrec"
	CASE "Update " 
		sAction = "listrec"
END SELECT
	

currentPage = TRIM(Request("currentPage"))
IF currentPage = "" THEN currentPage = 1

sThisFieldID = trim(Request("LeagueID"))

sSYID = trim(Request("sSYID"))
IF sSYID = "" THEN 
	sSQL = "SELECT TOP 1 SkiYearID FROM "&SkiYearTableName&" WHERE SkiYearID<>'1' ORDER BY SkiYearID DESC"
  	ChooseSQL(sSQL)
	sSYID=TRIM(rs("SkiYearID"))
END IF

'response.write("sSYID="&sSYID)

sListFilter= request("ListFilter")
IF sListFilter<>"" THEN Session("ListFilter")=sListFilter


sEditFld = session("EditFldFilter")
IF sEditFld = "" then sEditFld = "ALL"

sSortField=TRIM(Request("sSortField"))
IF sSortField = "" then sSortField = "TourID"

sRegionSelected=TRIM(Request("RegionSelected"))
IF sRegionSelected = "" then sRegionSelected = "ALL"

sTourSelected=TRIM(Request("sTourSelected"))
IF sTourSelected = "" THEN sTourSelected = "select"

sLeagueSelected=TRIM(Request("sLeagueSelected"))
'IF sLeagueSelected = "" then sLeagueSelected = "ALL"

sTourType=TRIM(Request("TourType"))
IF sTourType = "" then sTourType = "0"

sUseForLCQScr=TRIM(Request("UseForLCQScr"))
SELECT CASE sUseForLCQScr
	CASE "True", ""
		sUseForLCQScr=1
'	response.write("<br>Inside true")
	CASE "False"
		sUseForLCQScr=0
	response.write("<br>Inside False")
	CASE ELSE
	response.write("<br>sUseForLCQScr="&sUseForLCQScr)
	response.write("<br>Inside ELSE")

END SELECT

'response.end

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


WriteIndexPageHeader


' ------------------------------
' --- MAIN BRANCHING SECTION ---
' ------------------------------


SELECT CASE LCASE(sAction)

  CASE "listrec"

'	response.write("SELECT LeagueName FROM "&LeagueTableName&" WHERE LeagueID='"&sLeagueSelected&"'")
'	response.end

	sSQL="SELECT LeagueName FROM "&LeagueTableName
	IF sLeagueSelected<>"ALL" THEN sSQL = sSQL + " WHERE LeagueID='"&sLeagueSelected&"'"

	IF sLeagueSelected<>"" THEN 
		ChooseSQL(sSQL)
		sLeagueName=rs("LeagueName")
		WriteHeaders "Tournaments in LeagueID '"&sLeagueSelected&"' - "&sLeagueName&" "
	END IF

	ListRecords


  CASE "addrec"
	IF sThisFieldID = "" THEN sThisFieldID = Session("NewKeyCode")

    	WriteHeaders "Add Record:  TourID = "&sTourSelected&" "&sTourName&"'<br> into "&ThisTempTableName


    	IF len(sTourSelected) < 6 THEN %>
		<br><H2><center><font color="red">Invalid TourID Length.<BR></font></center></H2><% 	
    		Listrecords

    	ELSEIF sLeagueSelected = "ALL" THEN %>
		<br><H2><center><font color="red">You must select a League into which this TourID will be added.<BR></font></center></H2><% 	
    		Listrecords

	ELSE    	
		ChooseSQL("SELECT * FROM "&LeagueToursTableName&" WHERE LeagueID='"&sLeagueSelected&"' AND LEFT(TourID,6)='"&LEFT(sTourSelected,6)&"'")
		IF rs.EOF THEN TempEOF = "Y" ELSE TempEOF = "N"
		rs.close: set rs = nothing

		IF TempEOF = "N" THEN %>
			<br><H2><center><font color="red">This TourID already exists in this League.<br><BR></font></center></H2><%
	    		Listrecords
		ELSE

			AddRecord
			'Listrecords
			ShowEditor
		END IF
  	END IF

  CASE "editrec"
    WriteHeaders "Edit LeaqueID Record:  TourID = "&sTourID&" for SYID = "&sSYID&" and Sports Discipline = "&session("sSptsGrpID")&" in " & ThisTempTableName
    ShowEditor

  CASE "saverec"
    WriteHeaders "Record Saved :  TourID = "&sTourID&" for SYID = "&sSYID&" and Sports Discipline = "&session("sSptsGrpID")&" in " & ThisTempTableName
    SaveRec

  CASE "delrec"
    	IF sLeagueSelected <> "ALL" THEN 
		WriteHeaders "Delete TourID '"&sTourSelected&"' from LeagueID '"&sLeagueSelected&"' <br>in table " & ThisTempTableName
	ELSE %>
		<br><H2><center><font color="red">You must select a League into which this TourID will be added.<BR></font></center></H2><% 	
    		Listrecords
	END IF

	DeleteRec



END SELECT


WriteFooter
WriteIndexPageFooter






' ------------------------
  SUB WriteHeaders(sTitle)
' ------------------------

' Write Headers for Page
%>
<br>
<TABLE class="innertable" ALIGN="center" WIDTH="<%=TourTableWidth%>">
  <TR>
    <th ALIGN="CENTER">
	<Font COLOR="#FFFFFF" SIZE="2"><B><%= sTitle %></B></FONT>
    </th>
  </TR>
</TABLE>
<BR><%


END SUB



' ----------------
  SUB ListRecords
' ----------------

'  Lists the table Records

%> 
<form action="/rankings/editleaguetours.asp" method="post">
<input type="hidden" name="search" value="1">

<TABLE align="center" class="innertable" WIDTH="<%=TourTableWidth%>">
<TR>
  <th ALIGN="center"><FONT COlOR="#FFFFFF" SIZE="<%=fontsize1%>">League ID</FONT></th>
  <th ALIGN="center"><FONT COlOR="#FFFFFF" SIZE="<%=fontsize1%>">TourID</FONT></th>
  <th ALIGN="center"><FONT COlOR="#FFFFFF" SIZE="<%=fontsize1%>">Tour Ski Year</FONT></th>
  <th ALIGN="center"><FONT COlOR="#FFFFFF" SIZE="<%=fontsize1%>">Sort By</FONT></th>
  <th ALIGN="center"><FONT COlOR="#FFFFFF" SIZE="<%=fontsize1%>">SD</FONT></th>
</TR>

<TR>
<TD ALIGN="center">
<%

' --- Builds list of League from Master table LEAGUES ---
LeagueDropBuild




' ------------   Gets NEW CODE  -----------------
%>
</TD>
<TD ALIGN="center"><%
	TourDropBuild %>
</TD><%

' ------------   Builds NCWRegion Drop Down list ----------------- %>
  <TD ALIGN="center"><%
	' --- Builds Ski Year Drop down based on Ski Year table ---
	SkiYearDropBuild %>
  </TD>

  <TD ALIGN="center">
	<select name="sSortField">
		<option value="tourID" <%IF sSortField = "tourID" THEN Response.Write("selected")%>>TourID</option>
		<option value="LG.leagueID" <%IF sSortField = "LG.leagueID" THEN Response.Write(" selected ")%>>LeagueID</option>
    	</select>
  </TD>

  <TD ALIGN="center"><FONT COlOR="#000000" FACE="<%=font1%>" SIZE="<%=fontsize1%>"><%=Session("sSptsGrpID")%></FONT></TD>
</TR>

</table>

<br>
<TABLE width="<%=TourTableWidth%>" align=center>
<TR>
  <TD align=center style="border-style:none;">
	<input type="submit" style="width:12em" name="action" value="Update Listing">
  </TD>
  <TD align=center style="border-style:none;">	
	<input type="submit" style="width:12em" name="action" value="Add New TourID">
  </TD>
</form>

<form action="/rankings/defaultHQ.asp" >
  <TD align=center style="border-style:none;">
	<input type=submit style="width:12em" value=" Main Menu" method="post">
  </td>
</form><%


Mark=1
IF Mark=2 THEN %>
	<form action="/rankings/EditLeagueTours.asp" method="post">

	  <TD align=center style="border-style:none;">	
		<input type="hidden" name="search" value="clear">
		<input type="submit" style="width:12em" value="Reset Search Filters">
	  </TD>
	</form><%
END IF %>
</TR>
</TABLE>
<%



' --- First find the two-degit year for SkiYearSelected ---
sSQL="SELECT RIGHT(SkiYear,2) AS ThisYear FROM "&SkiYearTableName&" AS ST WHERE ST.SkiYearID = '"&sSYID&"'"
ChoosePagesSQL sSQL,currentPage, 40
ThisYear=rs("ThisYear")


' -------------  Run Query to League Tour List Recordset  ----------------------
  sSQL = "SELECT LG.LeagueID, LT.TourID, ST.TName AS [Tournament Name], LT.TourType AS [Tour Type], LT.UseForLCQScr, LG.Status FROM "&LeagueToursTableName&" AS LT, "&LeagueTableName&" AS LG"
  sSQL = sSQL + ", "&SanctionTableName&" AS ST "
  sSQL = sSQL + " WHERE LEFT(ST.TournAppID,6)=LEFT(LT.TourID,6) "
  sSQL = sSQL + " AND LT.LeagueID=LG.LeagueID"
'  sSQL = sSQL + " AND SkiYearID='"&sSYID&"'"

  IF sLeagueSelected<>"ALL" THEN sSQL = sSQL + " AND LG.LeagueID = '"&sLeagueSelected&"'" 
  sSQL = sSQL + " AND LG.SptsGrpID = '"&Session("sSptsGrpID")&"'" 
  sSQL = sSQL + " ORDER BY "&sSortField

  ChoosePagesSQL sSQL,currentPage, 40


IF rs.eof THEN  %>
	<br><br>
	<center><font color="red" size="3"><i><b>Please select a League ID</b></i></font></center><%
ELSE
	rowCount = 0

	' ---------------  Displays table HEADINGS  ----------------------

	%>
	<BR>
	<TABLE class="innertable" Align=center WIDTH="<%=TourTableWidth%>" >
	  <TR>
	    <th ALIGN="Center"><FONT COlOR="#FFFFFF" SIZE="<%=fontsize1%>">Delete</FONT></th>
	    <th ALIGN="Center"><FONT COlOR="#FFFFFF" SIZE="<%=fontsize1%>">Edit</FONT></th><%

		FOR i = 0 TO rs.fields.count - 1
			TempFN = rs.fields(i).name
			j = 0 %>

	   		<th ALIGN="Center" vAlign="top" nowrap>
			  <FONT COlOR="#FFFFFF" FACE="<%=font1%>" SIZE="<%=fontsize1%>"><%=Rs.Fields(i).name%></FONT>
			</th><%
		NEXT %>
	  </TR><%

	' --------------  Display table data here with paging --------------------------

	DO WHILE NOT rs.eof

		IF rowCount = rs.PageSize THEN EXIT DO	%>

 		<TR>
		<TD ALIGN="center" vAlign="top"><FONT COlOR="#FFFFFF" SIZE="<%=fontsize1%>"><% WriteLink "?action=delrec&sTourSelected="&rs("TourID")&"&sLeagueSelected="&rs("LeagueID")&"&sSYID="&sSYID&"&TourType="&rs("Tour Type")&"&UseForLCQScr="&rs("UseForLCQScr"),"Delete","" %></FONT></TD><%
		AllowEdit=true
		IF AllowEdit=true THEN %>
			<TD ALIGN="center"><FONT COlOR="#000000" SIZE="<%=fontsize1%>"><% WriteLink "?action=editrec&sTourSelected="&rs("TourID")&"&sLeagueSelected="&rs("LeagueID")&"&sSYID="&sSYID&"&TourType="&rs("Tour Type")&"&UseForLCQScr="&rs("UseForLCQScr"),"Edit","" %></FONT></TD><%
		ELSE %>
			<TD ALIGN="center"><FONT COlOR="#000000" SIZE="<%=fontsize1%>">Edit</FONT></TD><%
		END IF 

		FOR i = 0 TO rs.fields.count - 1
	
			RowColor=""
			TempFN = rs.fields(i).name
			IF TempFN="TourID" THEN
				IF RIGHT(LEFT(rs.Fields(i).value,6),3)="001" OR (RIGHT(LEFT(rs.Fields(i).value,6),3)="999" AND ThisYear<>LEFT(rs.Fields(i).value,2)) THEN
					RowColor="background-color:"&scolor08
				ELSEIF ThisYear<>LEFT(rs.Fields(i).value,2) THEN
					RowColor="background-color:"&scolor04
				END IF
			END IF
		
			%><TD ALIGN="center" style="<%=RowColor%>">
				<FONT COlOR="#000000" SIZE="<%=fontsize1%>">&nbsp;<%

			    IF isnull(rs.Fields(i).value) THEN
				response.write ("&nbsp;")
    			    ELSE
				Response.Write(trim(Rs.Fields(i).Value)) 
			    END IF  
	
			%>&nbsp;
			  </FONT>
			</TD><%

		NEXT	%>

		</TR><% 
		rowCount = rowCount + 1
		rs.movenext
	LOOP %>

	</TABLE>
	<br>
	<TABLE class="blanktable" align=center width="<%=TourTableWidth%>">
	  <tr>
	    <td width=25px bgcolor="<%=scolor08%>">&nbsp;</td>
	    <td width=150px><FONT COlOR="#000000" SIZE="<%=fontsize1%>">&nbsp;Prior Nationals</font></td>
	    <td width=25px bgcolor="<%=scolor07%>">&nbsp;</td>
	    <td width=150px><FONT COlOR="#000000" SIZE="<%=fontsize1%>">&nbsp;Prior Regionals</font></td>
	    <td width=25px bgcolor="<%=scolor04%>">&nbsp;</td>
	    <td width=150px><FONT COlOR="#000000" SIZE="<%=fontsize1%>">&nbsp;Not Current Year</font></td>
	  </tr>
	  <tr><td colspan=6>&nbsp;</td></tr>
	  <tr><td align=left colspan=6><FONT COlOR="#000000" SIZE="<%=fontsize1%>"><b>TOUR TYPE</b></font></td></tr>
	  <tr><td align=left colspan=6><FONT COlOR="#000000" SIZE="<%=fontsize1%>">A - Nationals</font></tr>
	  <tr><td align=left colspan=6><FONT COlOR="#000000" SIZE="<%=fontsize1%>">B - Regionals </font></td></tr>
	  <tr><td align=left colspan=6><FONT COlOR="#000000" SIZE="<%=fontsize1%>">C - State Qualifier or equivalent</font></tr>
	  <tr><td align=left colspan=6><FONT COlOR="#000000" SIZE="<%=fontsize1%>">D - Qualifier Other </font></td></tr>

	</table>
	<br><br>
	<%
	DoCount currentPage
END IF

rs.close
set rs = nothing

END SUB





' -----------------------
   SUB LeagueDropBuild
' -----------------------

' ------------   Builds Ski Year Drop Down list ----------------- 

ChooseSQL("SELECT DISTINCT LeagueID, LeagueName FROM "&LeagueTableName&" WHERE SptsGrpID='"&Session("sSptsGrpID")&"' ORDER BY LeagueID") %>


<SELECT name='sLeagueSelected' style="width:8em"><%

  response.write("<option value ='ALL'")
  IF sLeagueSelected = "ALL" THEN response.write(" SELECTed")
  response.write(">ALL</option><br>")

  IF NOT rs.eof THEN
	rs.movefirst
	DO WHILE not rs.eof
	  response.write(" <option value ="""&rs("LeagueID")&""" ")
	  response.write(" <a title="""&rs("LeagueName")&"""")

	  IF trim(rs("LeagueID")) = sLeagueSelected THEN
	    response.write(" selected")
	  END IF

	  response.write(">")
	  response.write(rs("LeagueID"))
	  response.write("</a></option><br>")
	  rs.movenext
	LOOP
  END IF %>

</SELECT><%

END SUB


' -----------------------
   SUB TourDropBuild
' -----------------------

' ------------   Builds Tournament Drop Down list ----------------- 

' --- First find the two-degit year for SkiYearSelected ---
sSQL="SELECT RIGHT(SkiYear,2) AS ThisYear FROM "&SkiYearTableName&" AS ST WHERE ST.SkiYearID = '"&sSYID&"'"

ChooseSQL("SELECT TournAppID, TName FROM "&SanctionTableName&" WHERE LEFT(TournAppID,2) IN ("&sSQL&") AND SptsGrpID='"&Session("sSptsGrpID")&"' ORDER BY TournAppID") %>


<SELECT name='sTourSelected' style="width:8em"><%

  response.write("<option value ='select'")
  IF sTourSelected = "select" THEN response.write(" Selected")
  response.write(">Select Tour</option><br>")

  IF NOT rs.eof THEN
	rs.movefirst
	DO WHILE not rs.eof
	  response.write("<option value =""" & rs("TournAppID") & """")
	  response.write(" <a title="""&rs("TName")&"""")
	  IF trim(rs("TournAppID")) = sTourSelected THEN
	    response.write(" SELECTed")
	  END IF

	  response.write(">")
	  response.write(rs("TournAppID"))
	  response.write("</a></option><br>")
	  rs.movenext
	LOOP
  END IF %>

</SELECT><%

END SUB






' -----------------------
  SUB RegionDropBuild
' -----------------------


' ------------   Builds Ski Year Drop Down list ----------------- 

ChooseSQL("SELECT DISTINCT NCWRegion FROM "&ThisTempTableName&" WHERE SptsGrpID='"&Session("sSptsGrpID")&"'") %>

<SELECT name='RegionSelected' style="width:6em"><%

  response.write("<option value ='ALL'")
  IF sRegionSelected = "ALL" THEN response.write(" SELECTed")
  response.write(">ALL</option><br>")

  IF NOT rs.eof THEN
	rs.movefirst
	DO WHILE not rs.eof
	  response.write("<option value =""" & rs("NCWRegion") & """")

	  IF trim(rs("NCWRegion")) = sRegionSelected THEN
	    response.write(" SELECTed")
	  END IF

	  response.write(">")
	  response.write(rs("NCWRegion"))
	  response.write("</option><br>")
	  rs.movenext
	LOOP
  END IF %>
</SELECT><%

END SUB






' -----------------------
   SUB SkiYearDropBuild
' -----------------------

' ------------   Builds Ski Year Drop Down list ----------------- %>

<SELECT name='sSYID'>
<%

ChooseSQL("SELECT * FROM "&SkiYearTableName&" WHERE SkiYearID<>'1'")

DO WHILE not rs.eof
  response.write("<option value =""" & rs("SkiYearID") & """")

  IF trim(rs("SkiYearID")) = sSYID THEN
    response.write(" SELECTed")
  END IF

  IF sSYID = "" and rs("DefaultYear") THEN
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



WriteButton "?action=listrec","No Change - Go To League List","<BR>"


ChooseSQL("SELECT * FROM "&LeagueToursTableName&" WHERE (LEFT(TourID,6)='"&LEFT(sTourSelected,6)&"' AND LeagueID='"&sLeagueSelected&"')")


%>
<FORM METHOD="POST" ACTION="<% Response.Write(ThisPage) %>?action=saverec">
<TABLE class="innertable" BORDER="1" ALIGN=center >
  <TR>
    <th ALIGN="Left"><Font SIZE="<%=fontsize1%>" color="#FFFFFF"><B>Field</B></FONT></th>
    <th ALIGN="Left"><Font SIZE="<%=fontsize1%>" color="#FFFFFF"><B>Value</B></FONT></th>
    <th ALIGN="Left"><Font SIZE="<%=fontsize1%>" color="#FFFFFF"><B>Field Type</B></FONT></th>
  </TR>

<%

' *** Important -- first two fields are Code (ID) and Ski Year ID code.
' *** These two fields serve as the record key, and hence are NOT editable.

'response.write(" <br>sUseForLCQScr="& sUseForLCQScr)
'response.write(" <br>UseForLCQScr="& UseForLCQScr)

FOR i = 0 TO rs.fields.count - 1

	TempFN = rs.fields(i).name

'response.write(" <br>Show Editor NAME="& Rs.Fields(i).name)
	%>
	<TR>
	  <TD ALIGN="Left" width="100px">
		<Font SIZE="<%=fontsize1%>"><B><% Response.Write(Rs.Fields(i).name) %></B></FONT>
	  </TD>
	   <TD ALIGN="Left" width="300px"><%
		IF i = 0 THEN  %>
			<input type="hidden" name="LeagueID" value="<%=rs.fields(i).value%>">
			<font size="2"><%=rs.fields(i).value %></font><%
		ELSEIF i = 1 THEN  %>
			<input type="hidden" name="SptsGrpID" value="<%=rs.fields(i).value%>">
			<font size="2"><%=rs.fields(i).value %></font><%

		ELSEIF Rs.Fields(i).name="TourType" THEN %>
			<SELECT name="TourType" style="width:9em">
			  <option value ="A" <%IF sTourType = "A" THEN Response.Write(" selected ")%>>A - Nationals</Option><br>
			  <option value ="B" <%IF sTourType = "B" THEN Response.Write(" selected ")%>>B - Regionals</Option><br>
			  <option value ="C" <%IF sTourType = "C" THEN Response.Write(" selected ")%>>C - States or Other</Option><br>
  			  <option value ="D" <%IF sTourType = "D" THEN Response.Write(" selected ")%>>D - Qualifier Other</Option><br><%

		ELSEIF Rs.Fields(i).name="UseForLCQScr" THEN %>
			<SELECT name="UseForLCQScr" style="width:9em">
			  <option value = 1 <%IF sUseForLCQScr = 1 THEN Response.Write(" selected ")%>>Yes</Option><br>
			  <option value = 0 <%IF sUseForLCQScr = 0 THEN Response.Write(" selected ")%>>No</Option><br>
			</SELECT><%			
		ELSE %>
			<Font COLOR="#000000" FACE="<%=font1%>" SIZE="<%=fontsize1%>"><% WriteType i %></FONT></TD><%
		END IF %>
	    <TD><%= len(rs.fields(i).value) %></TD>
	  </TD>
	</TR>
	<%


NEXT




%>
</TABLE><BR>
<TABLE BORDER="0" ALIGN=center WIDTH=30% CELLPADDING="3" CELLSPACING="0">
<TR>
<TD ALIGN="center">


 	<% ' --- NEW TEST ----- %>

    <input type="submit" style="width:9em" value="Save">
    <input type="hidden" name="sSYID" value="<%=sSYID%>">
    <input type="hidden" name="sTourSelected" value="<%=sTourSelected%>">
    <input type="hidden" name="sLeagueSelected" value="<%=sLeagueSelected%>">
    <input type="hidden" name="sSortField" value="<%=sSortField%>">
</TD>
<TD ALIGN="center"><input type="reset" style="width:9em" value="Reset"></TD>
</TR>
</TABLE>

</FORM>
<%

rs.close
set rs = nothing


%>
<TABLE>
  <TR>
  	<TD align="left">
  		<Font COLOR="#000000" FACE="<%=font1%>" SIZE="<%=fontsize1%>">
  		    <b>This information must be set for each tournament being applied to this LeagueID</b>
  		</font>
  	</TD>
  </TR> 

  <TR>
  	<TD align="left">
  		<Font COLOR="#000000" FACE="<%=font1%>" SIZE="<%=fontsize1%>">
  		    <b>LeagueID</b> - The ID for which the qualifications are being established (format NATL2013)
  		</font>
  	</TD>
  </TR> 
  <TR>
  	<TD align="left">
  		<Font COLOR="#000000" FACE="<%=font1%>" SIZE="<%=fontsize1%>">
  		    <b>TourID</b> - A qualification tournament belonging to this league (e.g. Regionals belong to Nationals)
  		</font>
  	</TD>
  </TR> 
  <TR>
  	<TD align="left">
  		<Font COLOR="#000000" FACE="<%=font1%>" SIZE="<%=fontsize1%>">
  		    <b>TourType</b> - Description of the type for this tournament being added to this league.
  		</font>
  	</TD>
  </TR> 
  <TR>
  	<TD align="left">
  		<Font COLOR="#000000" FACE="<%=font1%>" SIZE="<%=fontsize1%>">
  		    <b>UseForLCQScore</b> - Check if this tournament is unique where scores can be used to meet qualification requirements and no other LCQ qualifications are being used.
  		</font>
  	</TD>
  </TR> 
</TABLE>
<%


END SUB



' -------------------------------
  SUB SaveRec
' -------------------------------


'Save the record to the table'

ChooseSQL("SELECT * FROM "&LeagueToursTableName&" WHERE LEFT(TourID,6)='"&LEFT(sTourSelected,6)&"' AND LeagueID='"&sLeagueSelected&"'")

	sSQL = "UPDATE "&LeagueToursTableName&" SET "
	' --- Ignores 1st field since it assumes this is the KEY ---
  	FOR i = 1 TO rs.fields.count - 1

		'response.write("<br>Name = "&Request.Form(rs.fields(i).name))

		IF Request.Form(rs.fields(i).name) <> ""  THEN

			IF RIGHT(sSQL,1) <> "," and RIGHT(sSQL,1) <> " " THEN sSQL = sSQL + ", "

			sSQL = sSQL + rs.fields(i).name
			sSQL = sSQL + "='" + sqlclean(Request.Form(rs.fields(i).name)) + "'"
			
		END IF


	NEXT      

rs.close
set rs = nothing

sSQL = sSQL + " WHERE LEFT(TourID,6)='"&LEFT(sTourSelected,6)&"' AND LeagueID='"&sLeagueSelected&"'"

'response.write("<br>IN SAVEREC - "&sSQL)
'response.end

OpenCon
con.execute(sSQL)
'WriteLog(date() &"  "& time() &"   "&KeyFieldName&" Record Updated - "& sSQL)
CloseCon

%>
<center><FONT COlOR="red" FACE="<%=font1%>" SIZE="<%=fontsize3%>"><b><i>Your updated record has been saved.</I></b></font></center>
<BR>

<form action="/rankings/EditLeagueTours.asp?action=listrec">
  <center><input type=submit value="Click here to Continue" method="post"></center>
      <input type="hidden" name="sSYID" value="<%=sSYID%>">
      <input type="hidden" name="sTourSelected" value="<%=sTourSelected%>">
      <input type="hidden" name="sLeagueSelected" value="<%=sLeagueSelected%>">
    <input type="hidden" name="sSortField" value="<%=sSortField%>">
</form>
<%
END SUB



'------------------
  SUB AddRecord
'------------------



' First we check for existence of a new Specified Code (ID)
' If not found then go ahead and add it, otherwise don't.
' Then the Editing window will be presented -- by mainline CASE.

	sSQL = "INSERT INTO "&ThisTempTableName&" (LeagueID, TourID) VALUES ('"&sLeagueSelected&"', '"&sTourSelected&"')"
	OpenCon

	con.execute(sSQL)
	Closecon
	'WriteLog(date() &"  "& time() &"   New TourID Has Been Added to LeagueTour - "& sSQL)

END SUB



'-------------------
  SUB DeleteRec
'-------------------

ChooseSQL("SELECT * FROM "&LeagueToursTableName&" WHERE LEFT(TourID,6)='"&LEFT(sTourSelected,6)&"' AND LeagueID='"&sLeagueSelected&"'") 


IF LCASE(Request("confirm")) = "yes" THEN

    'delete the record'
    'WriteLog(date() &"  "& time() &"  "&KeyFieldName&" Table Record " & rs("&ThisFieldID&") & " for SY=" & rs("SkiYearID") & " (" & rs("&MyKeyFieldName2&") & ") has been deleted.")

    	IF isrecordsetempty = false THEN
        	rs.movefirst
        	rs.delete
        	rs.UPDATEBatch 3
	END IF

	rs.close
	set rs = nothing
    
    %>
    <center><font COLOR="#FFFFFF" FACE="<%=font1%>" SIZE="2"><I><b>The record has been deleted.</b></I></font></center>
    <BR>&nbsp;<BR>
    <%

END IF


WriteButton "?action=listrec","Return To League List","<BR><BR>"

IF LCASE(Request("confirm")) = "" THEN
%>  <br><br>
    <center>
     <Font COLOR="#FFFFFF" FACE="<%=font1%>" SIZE="2">
    Type the word "YES" IF you are sure you wish to delete this record. </font>
    <br>
    <Font COLOR="red" FACE="<%=font1%>" SIZE="2">
    Note: Scores from this Tournament which are to be grouped by League may be affected.
    </font>
    <br><br>
    <form action="/rankings/EditLeagueTours.asp" method="post"> 
      <input type="hidden" name="action" value="delrec">
      <input type="hidden" name="sTourSelected" value="<%=sTourSelected%>">
      <input type="hidden" name="sLeagueSelected" value="<%=sLeagueSelected%>">
      <input type="hidden" name="sSYID" value="<%=sSYID%>">
      <input type="text" name="confirm" Size="4" MAXLENGTH="3">
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
	    response.write(sSYID)
	    %><input type="hidden" name="SkiYearID" value="<%=sSYID%>"><%

	CASE "OLDSKIYEARID" 
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
			CASE 129 'char'
				%><input type="text" name="<% Response.Write(Rs.Fields(i).name) %>" SIZE="8" value="<% GetFieldValue i %>"><%
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

'response.write("In ChooseSQL "&sSQL)

set rs=Server.CreateObject("ADODB.recordset")
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

%>
<form action="/rankings/defaultHQ.asp" >
  <center><input type=submit value="Return to Main Menu" method="post"></center>
</form><%

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
'markdebug(sSQL)
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
  <input type="hidden" name="sTourSelected" value="<%=sTourSelected%>">
  <input type="hidden" name="sLeagueSelected" value="<%=sLeagueSelected%>">
  <input type="hidden" name="sSYID" value="<%=sSYID%>">
  <input type="hidden" name="sSortField" value="<%=sSortField%>">
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
      	Response.Write(" <a href=" & chr(34) & ThisPage & "?currentpage=" &  currentpage - i  & "&action1=" & sAction & chr(34) & "&sSortField="&sSortField&"&sLeagueSelected="&sLeagueSelected&">" & currentpage - i & "</a>")
    END IF
  NEXT

  Response.Write(" " & currentpage & " ")

  FOR i = 1 TO 10
   	IF currentpage + i <= h THEN
      		Response.Write(" <a href=" & chr(34) & ThisPage & "?currentpage=" &  currentpage + i  & "&action2=" & sAction & "&sSortField="&sSortField&"&sLeagueSelected="&sLeagueSelected& chr(34) &">" & currentpage + i & "</a>")
	END IF
  NEXT

  IF currentpage + 10 <= h THEN
    Response.Write(" ...")
  END IF

ELSE
  FOR i = 1 TO h
    Response.Write(" <a href=" & chr(34) & ThisPage & "?currentpage=" &i& "&action=" &sAction& "&sSortField=" &sSortField& "&sLeagueSelected="&sLeagueSelected& chr(34) & ">" &i& "</a>")
  next

END IF

IF h = 0 THEN h = 1
	Response.Write("<BR><Font COLOR=#FFFFFF FACE=font1 SIZE=0>Page " & currentPage & " of  "&h&"</font></center><BR><BR>")
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