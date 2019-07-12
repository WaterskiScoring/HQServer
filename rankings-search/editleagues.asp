<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->
<%

DefineTRAStyles




Dim currentPage, rowCount, i
Dim sAction, addstatus
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
Dim sLeagueSelected, sTourSelected, sTourName, sRegionSelected, sHomoType, sUseForCOA, State
Dim ThisFileName
Dim OldLeagueID
Dim UpperTableWidth
Dim LowerTableWidth

ThisFileName="EditLeagues.asp"
ThisFieldID="leagueid"
KeyFieldName="LeagueD"
KeyFieldShortName="LeagueID"
LengthofID=7

' --- KeyFieldName2 is the one the list gets sorted by on the listing ---
KeyFieldName2="LeagueName"
KeyCode="LeagueCode"
KeyCodeFilter="LeagueCodeFilter"

UpperTableWidth="80%"
LowerTableWidth="80%"





IF request("sSptsGrpID") <>"" THEN session("sSptsGrpID") = UCASE(request("sSptsGrpID"))

IF session("sSptsGrpID") = "TEST" THEN
	ThisTempTableName = "USAWSRank.Leagues"
ELSE
	IF session("sSptsGrpID") = "AWS" OR session("sSptsGrpID") = "NCW" THEN
		ThisTempTableName = "usawsrank.Leagues"
	ELSEIF session("sSptsGrpID")="USW" OR session("sSptsGrpID")="HYD" OR session("sSptsGrpID")="AKA" THEN
		ThisTempTableName = "usawsrank.Leagues"
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
  session("NewKeyCode") = request("KeyCode")
  session("EditFldFilter") = request("EditFld")
END IF

IF request("search") = "clear" THEN
  session.contents.remove("SearchFilter")
  session.contents.remove("NewKeyCode")
  session("EditFldFilter") = "ALL"
END IF



'response.write("<br>trim(Request(action)="&trim(Request("action")))

' --- Define ACTION ---
sAction = trim(Request("action"))
SELECT CASE left(sAction,7)
	CASE "Add New", "addnew" 
		sAction = "addnewleague"

	CASE "Update " 
		sAction = "listrec"
END SELECT
IF sAction = "" THEN sAction = "listrec"



currentPage = TRIM(Request("currentPage"))
IF currentPage = "" THEN currentPage = 1

sThisFieldID = trim(Request("LeagueID"))


' --- SkiYearID is an Integer field in table ---
sSYID = Request("sSYID")
IF sSYID = "" THEN 
	sSQL = "SELECT TOP 1 SkiYearID FROM "&SkiYearTableName&" WHERE SkiYearID<>'1' ORDER BY SkiYearID DESC"
	set rs=Server.CreateObject("ADODB.recordset")
	rs.open sSQL, sConnectionToTRATable, 3, 3	
	sSYID=rs("SkiYearID")
END IF


sListFilter= request("ListFilter")
IF sListFilter<>"" THEN Session("ListFilter")=sListFilter

sEditFld = session("EditFldFilter")
IF sEditFld = "" then sEditFld = "ALL"

sSortField=TRIM(Request("sSortField"))
IF sSortField = "" then sSortField = "LeagueID"

sRegionSelected=TRIM(Request("RegionSelected"))
IF sRegionSelected = "" then sRegionSelected = "ALL"

sTourSelected=TRIM(Request("sTourSelected"))
IF sTourSelected = "" THEN sTourSelected = "select"

sLeagueSelected=TRIM(Request("sLeagueSelected"))
'IF sLeagueSelected = "" then sLeagueSelected = "ALL"

sHomoType=TRIM(Request("HomoType"))
IF sHomoType = "" then sHomoType = "0"

sUseForCOA=TRIM(Request("UseForCOA"))
IF sUseForCOA = "" then sUseForCOA = "Y"

State=TRIM(Request("State"))

'response.write("State="&State)

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









WriteIndexPageHeader_NoMenu



'response.write("<br>AT BOTTOM OF TOP - sAction="&sAction)

'response.end



' ------------------------------
' --- MAIN BRANCHING SECTION ---
' ------------------------------


SELECT CASE LCASE(sAction)

  CASE "listrec"


			sSQL="SELECT LeagueName FROM "&LeagueTableName
			IF sLeagueSelected<>"ALL" THEN sSQL = sSQL + " WHERE LeagueID='"&sLeagueSelected&"'"

			IF sLeagueSelected<>"" THEN 
					set rs=Server.CreateObject("ADODB.recordset")
					rs.open sSQL, sConnectionToTRATable, 3, 3	

					'sLeagueName=rs("LeagueName")
					WriteHeaders "Leagues currently in LeagueTable"
			END IF

			ListRecords

  CASE "addnewleague"

    	WriteHeaders "Add a New League into "&ThisTempTableName

			AddRecord_New

  CASE "addrec"
			'IF sThisFieldID = "" THEN sThisFieldID = Session("NewKeyCode")

    	WriteHeaders "Add Record:  LeagueID = "&sLeagueSelected&"<br> into "&ThisTempTableName

    	IF len(sLeagueSelected) <> 8 THEN 
    			%>
					<br><H2><center><font color="red">Invalid LeagueID Length.<BR></font></center></H2>
					<% 	
    			Listrecords

    	ELSEIF sLeagueSelected = "ALL" THEN 
    			%>
					<br><H2><center><font color="red">You must select a League .<BR></font></center></H2>
					<% 	
    			Listrecords

			ELSE    	
					sSQL = "SELECT * FROM "&LeagueTableName&" WHERE LeagueID='"&sLeagueSelected&"'"
					set rs=Server.CreateObject("ADODB.recordset")
					rs.open sSQL, sConnectionToTRATable, 3, 3	



				IF rs.EOF THEN TempEOF = "Y" ELSE TempEOF = "N"
						rs.close: set rs = nothing

						IF TempEOF = "N" THEN 
								%>
								<br><H2><center><font color="red">This LeagueID already exists in this League.<br><BR></font></center></H2>
								<%
	    					Listrecords
						ELSE
								AddRecord
								'Listrecords
								ShowEditor
						END IF
  			END IF

  CASE "editrec"
    	WriteHeaders "Edit LeaqueID Record:  LeagueID = "&sLeagueSelected&" for SkiYearID = "&sSYID&" and Sports Discipline = "&session("sSptsGrpID")&" in " & ThisTempTableName
    	ShowEditor

  CASE "saverec"
    	WriteHeaders "Record Saved :  LeagueID = "&sLeagueSelected&" for SkiYearID = "&sSYID&" and Sports Discipline = "&session("sSptsGrpID")&" in " & ThisTempTableName
    	SaveRec

  CASE "delrec"
    	IF sLeagueSelected <> "ALL" THEN 
					WriteHeaders "Delete LeagueID '"&sLeagueSelected&"' <br>in table " & ThisTempTableName
			ELSE 
					%>
					<br><H2><center><font color="red">You must select a League .<BR></font></center></H2>
					<% 	
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
<form action="/rankings/<%=ThisFileName%>" method="post">
<input type="hidden" name="search" value="1">

<TABLE align="center" class="innertable" WIDTH="<%=UpperTableWidth%>">
<TR>
  <th ALIGN="center"><FONT COlOR="#FFFFFF" SIZE="<%=fontsize1%>">League ID</FONT></th>
  <th ALIGN="center"><FONT COlOR="#FFFFFF" SIZE="<%=fontsize1%>">League Ski Year</FONT></th>
  <th ALIGN="center"><FONT COlOR="#FFFFFF" SIZE="<%=fontsize1%>">Sort By</FONT></th>
  <th ALIGN="center"><FONT COlOR="#FFFFFF" SIZE="<%=fontsize1%>">SD</FONT></th>
</TR>

<TR>
<TD ALIGN="center"><%
	' --- Builds list of League from Master table LEAGUES ---
	LeagueDropBuild %>
</TD><%

' ------------   Builds NCWRegion Drop Down list ----------------- %>
  <TD ALIGN="center"><%
	' --- Builds Ski Year Drop down based on Ski Year table ---
	SkiYearDropBuild %>
  </TD>

  <TD ALIGN="center">
	<select name="sSortField">
		<option value="LG.leagueID" <%IF sSortField = "LG.leagueID" THEN Response.Write(" selected ")%>>LeagueID</option>
    	</select>
  </TD>

  <TD ALIGN="center"><FONT COlOR="#000000" FACE="<%=font1%>" SIZE="<%=fontsize1%>"><%=Session("sSptsGrpID")%></FONT></TD>
</TR>
</TABLE>

<br>

<TABLE width="<%=TourTableWidth%>" align=center>
<TR>
  <TD align=center style="border-style:none;">
	<input type="submit" style="width:12em" name="action" value="Update Listing">
  </TD>
  <TD align=center style="border-style:none;">	
	<input type="submit" style="width:12em" name="action" value="Add New LeagueID">
  </TD>
</form>

<form action="/rankings/defaultHQ.asp" >
  <TD align=center style="border-style:none;">
		<input type=submit style="width:12em" name="action" value=" Main Menu" method="post">
  </td>
</form>
<%


Mark=1
IF Mark=2 THEN 
	%>
	<form action="/rankings/<%=ThisFileName%>" method="post">
	  <TD align=center style="border-style:none;">	
			<input type="hidden" name="search" value="clear">
			<input type="submit" style="width:12em" value="Reset Search Filters">
	  </TD>
	</form>
	<%
END IF 
%>
</TR>
</TABLE>
<%



sSQL = "SELECT LG.LeagueID, LG.LeagueName AS [League Name], LG.Status, LG.SkiYearID AS [Ski<br>Year], LG.COD, LG.QualifyTour AS [Qualify<br>Tour], LG.HomoType AS Type, LG.COAMinClass AS [Min Class<br>For COA], LG.UseLCQRank AS [Use LCQ<br>By Rank], LG.UseLCQScore AS [Use LCQ<br>By Score], LG.RequirePart AS [Class of<br>Required<br>Participation], Qfy_By_AnyOverall_InStates AS [Qual by<br>Any State<br>Overall], State FROM "&LeagueTableName&" AS LG"
sSQL = sSQL + " WHERE 1=1"

IF sSYID<>0 THEN sSQL = sSQL + " AND SkiYearID="&sSYID
IF sLeagueSelected<>"ALL" THEN sSQL = sSQL + " AND LG.LeagueID = '"&sLeagueSelected&"'" 

sSQL = sSQL + " AND LG.SptsGrpID = '"&Session("sSptsGrpID")&"'" 
sSQL = sSQL + " ORDER BY "&sSortField


'response.write("sSQL="&sSQL)
'response.end

ChoosePagesSQL sSQL,currentPage, 40


'response.end


IF rs.eof THEN  
		%>
		<br><br>
		<center>
			<font color="red" size="3"><i><b>Please select a League ID and League Ski Year</b></i></font>
			<br>
			<font color="#000000" size="2"></font><b>Note: League ID and League Ski Year must correspond</b></font>
		</center>
		<br>
		<%
ELSE
		rowCount = 0

		' ----------------------------------------------------------------
		' ---------------  Displays table HEADINGS  ----------------------
		' ----------------------------------------------------------------

		%>
		<BR>
		<TABLE class="innertable" Align=center WIDTH="<%=LowerTableWidth%>" >
	  	<TR>
	    	<th ALIGN="Center"><FONT COlOR="#FFFFFF" SIZE="<%=fontsize1%>">Delete</FONT></th>
	    	<th ALIGN="Center"><FONT COlOR="#FFFFFF" SIZE="<%=fontsize1%>">Edit</FONT></th>
	    	<%

				FOR i = 0 TO rs.fields.count - 1
						TempFN = rs.fields(i).name
						SELECT CASE TempFN
			    			CASE "Min Class For COA"
										ThisTile="Minimum Score Class for LCQ Qualification By Score"
						    CASE ELSE
										ThisTitle=""	
						END SELECT

						j = 0 
						
						%>
			   		<th ALIGN="Center" vAlign="top" nowrap>
						  <FONT COlOR="#FFFFFF" FACE="<%=font1%>" SIZE="<%=fontsize1%>">
 								<a title="<%=ThisTitle%>"><%=Rs.Fields(i).name%></a>
			  			</FONT>
						</th>
						<%
				NEXT 
				
				%>
	  		</TR>
	  		<%

				' ----------------------------------------------------------------
				' --------------  Display ???? --------------------------
				' ----------------------------------------------------------------

				DO WHILE NOT rs.eof

						IF rowCount = rs.PageSize THEN EXIT DO	%>

				 		<TR>
							<TD ALIGN="center" vAlign="top"><FONT COlOR="#FFFFFF" SIZE="<%=fontsize1%>"><% WriteLink "?action=delrec&sLeagueSelected="&rs("LeagueID")&"&sSYID="&sSYID,"Delete","" %></FONT></TD><%
								AllowEdit=true
								IF AllowEdit=true THEN %>
										<TD ALIGN="center"><FONT COlOR="#000000" SIZE="<%=fontsize1%>"><% WriteLink "?action=editrec&sLeagueSelected="&rs("LeagueID")&"&sSYID="&sSYID,"Edit","" %></FONT></TD><%
								ELSE %>
										<TD ALIGN="center"><FONT COlOR="#000000" SIZE="<%=fontsize1%>">Edit</FONT></TD><%
								END IF 

								FOR i = 0 TO rs.fields.count - 1
										RowColor=""
										TempFN = rs.fields(i).name

										' --- Colors the LeagueID depending on whether the league is in the current Ski Year Selection
										IF TempFN="TourID" THEN
										IF RIGHT(LEFT(rs.Fields(i).value,6),3)="001" OR (RIGHT(LEFT(rs.Fields(i).value,6),3)="999" AND ThisYear<>LEFT(rs.Fields(i).value,2)) THEN
												RowColor="background-color:"&scolor08
										ELSEIF ThisYear<>LEFT(rs.Fields(i).value,2) THEN
												RowColor="background-color:"&scolor04
										END IF
								END IF
		
							%>
							<TD ALIGN="center" style="<%=RowColor%>">
								<FONT COlOR="#000000" SIZE="<%=fontsize1%>">&nbsp;<%

			    			IF isnull(rs.Fields(i).value) THEN
										response.write ("&nbsp;")
    			    	ELSE
										Response.Write(trim(Rs.Fields(i).Value)) 
			    			END IF  
	
								%>&nbsp;
			  				</FONT>
							</TD>
							<%

		NEXT	

		%>
		</TR>
		<% 

		rowCount = rowCount + 1
		rs.movenext
		LOOP 
		%>

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
	  	<tr><td align=left colspan=6><FONT COlOR="#000000" SIZE="<%=fontsize1%>">C - States or equivalent</font></tr>
	  	<tr><td align=left colspan=6><FONT COlOR="#000000" SIZE="<%=fontsize1%>">D - Qualifier Other </font></td></tr>
	  	<tr><td colspan=6>&nbsp;</td></tr>
	  	<tr><td align=left colspan=6><FONT COlOR="#000000" SIZE="<%=fontsize1%>"><b>OTHER INFORMATION</b></font></td></tr>
	  	<tr><td align=left colspan=6><FONT COlOR="#000000" SIZE="<%=fontsize1%>"><b>Min Class For COA</b> - Sets the minimum SCORE class for LCQ qualification By Score.</font></td></tr>
	  	<tr><td align=left colspan=6><FONT COlOR="#000000" SIZE="<%=fontsize1%>"><b>Class Required For Participation</b> - Specifies the class of the tournament in qhich there is REQUIRED particpation prior to competing in this tournament</font></td></tr>

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

	sSQL = "SELECT DISTINCT LeagueID, LeagueName FROM "&LeagueTableName&" WHERE SptsGrpID='"&Session("sSptsGrpID")&"' ORDER BY LeagueID"
	set rs=Server.CreateObject("ADODB.recordset")
	rs.open sSQL, sConnectionToTRATable, 3, 3	

%>
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
	sInnerSQL="SELECT RIGHT(SkiYear,2) AS ThisYear FROM "&SkiYearTableName&" AS ST WHERE ST.SkiYearID = "&sSYID&""

	sSQL = "SELECT TournAppID, TName FROM "&SanctionTableName&" WHERE LEFT(TournAppID,2) IN ("&sInnerSQL&") AND SptsGrpID='"&Session("sSptsGrpID")&"' ORDER BY TournAppID" 
	set rs=Server.CreateObject("ADODB.recordset")
	rs.open sSQL, sConnectionToTRATable, 3, 3	

%>
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

	sSQL = "SELECT DISTINCT NCWRegion FROM "&ThisTempTableName&" WHERE SptsGrpID='"&Session("sSptsGrpID")&"'"
	set rs=Server.CreateObject("ADODB.recordset")
	rs.open sSQL, sConnectionToTRATable, 3, 3	

%>

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


IF sSYID<>"" THEN sSYID=INT(sSYID)

' ------------   Builds Ski Year Drop Down list ----------------- 
	sSQL = "SELECT * FROM "&SkiYearTableName&" WHERE SkiYearID<>1 ORDER BY SkiYearID DESC" 
	set rs=Server.CreateObject("ADODB.recordset")
	rs.open sSQL, sConnectionToTRATable, 3, 3	

%>
<SELECT name='sSYID' style="width:10em">
	<option value = "0" <% IF sSYID=0 THEN response.write(" selected ")%>>All Years</option>><%

  DO WHILE not rs.eof %>

	<option value = "<%=rs("SkiYearID")%>" <%IF rs("SkiYearID") = sSYID THEN response.write(" selected ")%>><%=rs("SkiYearName")%></option><%

	rs.movenext
  LOOP %>
</SELECT><%

END SUB



' --------------
  SUB ShowEditor
' --------------


' --- Gets recordset for possible Ski Years ---
SET rsSY = Server.CreateObject("ADODB.recordset")
sSQL = "SELECT * from " & SkiYearTableName
sSQL = sSQL + " WHERE SkiYearID<>1"
sSQL = sSQL + " ORDER BY EndDate DESC"
rsSY.open sSQL, SConnectionToTRATable, 3, 3  


sSQL = "SELECT * FROM "&LeagueTableName&" WHERE LeagueID='"&sLeagueSelected&"'"
set rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, sConnectionToTRATable, 3, 3	

'response.write("SELECT * FROM "&LeagueTableName&" WHERE LeagueID='"&sLeagueSelected&"'")
'response.end

%>
<FORM METHOD="POST" ACTION="<% Response.Write(ThisPage) %>?action=saverec">
<TABLE class="innertable" BORDER="1" ALIGN=center >
  <TR>
    <th ALIGN="Left"><Font SIZE="<%=fontsize1%>" color="#FFFFFF"><B>Field</B></FONT></th>
    <th ALIGN="Left"><Font SIZE="<%=fontsize1%>" color="#FFFFFF"><B>Value</B></FONT></th>
    <th ALIGN="center"><Font SIZE="<%=fontsize1%>" color="#FFFFFF"><B>Length</B></FONT></th>
    <th ALIGN="center"><Font SIZE="<%=fontsize1%>" color="#FFFFFF"><B>Type</B></FONT></th>
  </TR>
<%


FOR i = 0 TO rs.fields.count - 1

	TempFN = rs.fields(i).name

'response.write(" <br>Show Editor NAME="& Rs.Fields(i).name)
	%>
	<TR>
	  <TD ALIGN="Left" width="220px">
		<Font SIZE="<%=fontsize1%>"><B><% Response.Write(Rs.Fields(i).name) %></B></FONT>
	  </TD>
	   <TD ALIGN="Left" width="300px"><%

		IF i = 0 THEN  %>
			<input type="hidden" name="LeagueID" value="<%=rs.fields(i).value%>">
			<font size="2"><%=rs.fields(i).value %></font><%
		ELSEIF i = 1 THEN  %>
			<input type="hidden" name="SptsGrpID" value="<%=rs.fields(i).value%>">
			<font size="2"><%=rs.fields(i).value %></font><%

		ELSEIF Rs.Fields(i).name="SkiYearID" THEN %>
			<SELECT name="SkiYearID" style="width:12em"><%
				DO WHILE NOT rsSY.eof %>
					<option value ="<%=INT(rsSY("SkiYearID"))%>" <%IF rs.Fields(i).value = rsSY("SkiYearID") THEN Response.Write(" selected ")%>><%=rsSY("SkiYearID")%> - <%=rsSY("SkiYearName")%></Option><br><%
					rsSY.movenext
				LOOP %>		
			</SELECT><%

		ELSEIF Rs.Fields(i).name="Status" THEN %>
			<SELECT name="Status" style="width:9em">
			  <option value ="A" <%IF Rs.Fields(i).value = "A" THEN Response.Write(" selected ")%>>Active</Option><br>
			  <option value ="X" <%IF Rs.Fields(i).value = "X" THEN Response.Write(" selected ")%>>Inactive</Option><br>
			</SELECT><%
			
		ELSEIF Rs.Fields(i).name="RequirePart" THEN %>
			<SELECT name="RequirePart" style="width:9em">
			  <option value ="B" <%IF Rs.Fields(i).value = "B" THEN Response.Write(" selected ")%>>Regionals</Option><br>
			  <option value ="C" <%IF Rs.Fields(i).value = "C" THEN Response.Write(" selected ")%>>States</Option><br>
			  <option value ="-" <%IF Rs.Fields(i).value = "-" THEN Response.Write(" selected ")%>>None</Option><br>
			</SELECT><%

		ELSEIF Rs.Fields(i).name="HomoType" THEN %>
			<SELECT name="HomoType" style="width:9em">
			  <option value ="A" <%IF Rs.Fields(i).value = "A" THEN Response.Write(" selected ")%>>A - Nationals</Option><br>
			  <option value ="B" <%IF Rs.Fields(i).value = "B" THEN Response.Write(" selected ")%>>B - Regionals</Option><br>
			  <option value ="C" <%IF Rs.Fields(i).value = "C" THEN Response.Write(" selected ")%>>C - States</Option><br>
  			  <option value ="D" <%IF Rs.Fields(i).value = "D" THEN Response.Write(" selected ")%>>D - Qualifier Other</Option><br>
			</SELECT><%

		ELSEIF Rs.Fields(i).name="COAMinClass" THEN %>
			<SELECT name="COAMinClass" style="width:9em">
			  <option value ="R" <%IF Rs.Fields(i).value = "R" THEN Response.Write(" selected ")%>>E/L/R - Record</Option><br>
			  <option value ="C" <%IF Rs.Fields(i).value = "C" THEN Response.Write(" selected ")%>>C - Standard</Option><br>
			</SELECT><%

		ELSEIF Rs.Fields(i).name="UseForCOA" THEN %>
			<SELECT name="UseForCOA" style="width:9em">
			  <option value ="Y" <%IF sUseForCOA = "Y" THEN Response.Write(" selected ")%>>Yes</Option><br>
			  <option value ="N" <%IF sUseForCOA = "N" THEN Response.Write(" selected ")%>>No</Option><br>
			</SELECT><%			

		ELSEIF Rs.Fields(i).name="State" THEN

				StateArray = Split(USStatesList2,",")  %>
				<select name="State" style="width:4em"><%
				  FOR kvar = 0 TO UBOUND(StateArray)
				    'IF TRIM(State) = TRIM(StateArray(kvar)) THEN
						IF TRIM(Rs.Fields(i).value) = TRIM(StateArray(kvar)) THEN
								response.write("<option value = """&State&""" SELECTED>"&Rs.Fields(i).value&"</option>")
				    ELSE
								response.write("<option value = """&StateArray(kvar)&""">"&StateArray(kvar)&"</option>")
				    END IF
				  NEXT  %>
				</select><%
		ELSE %>
			<Font COLOR="#000000" FACE="<%=font1%>" SIZE="<%=fontsize1%>"><% WriteType i, len(rs.fields(i).value) %></FONT></TD><%
		END IF %>

	    <TD align="center"><Font COLOR="#000000" FACE="<%=font1%>" SIZE="<%=fontsize1%>"><%= len(rs.fields(i).value) %></font></TD>
 	    <TD align="center"><Font COLOR="#000000" FACE="<%=font1%>" SIZE="<%=fontsize1%>"><%=Rs.Fields(i).type%></font></TD>

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

WriteButton "?action=listrec","No Change - Go To League List",""


rs.close
set rs = nothing

%>
<TABLE>
  <TR>
  	<TD align="left">
  		<Font COLOR="#000000" FACE="<%=font1%>" SIZE="<%=fontsize1%>">
  		    <b>The following settings apply to the tournament for which the Qualification/League settings are being input</b>
  		</font>
  	</TD>
  </TR> 

  <TR>
  	<TD align="left">
  		<Font COLOR="#000000" FACE="<%=font1%>" SIZE="<%=fontsize1%>">
  		    <b>QualifyTour</b> - Sanction # of the tournament
  		</font>
  	</TD>
  </TR> 
  <TR>
  	<TD align="left">
  		<Font COLOR="#000000" FACE="<%=font1%>" SIZE="<%=fontsize1%>">
  		    <b>HomoType</b> - Level of tournament (Nationals, Regionals, States, Other)
  		</font>
  	</TD>
  </TR> 
  <TR>
  	<TD align="left">
  		<Font COLOR="#000000" FACE="<%=font1%>" SIZE="<%=fontsize1%>">
  		    <b>COAMinClass</b> - Minimum class required for qualification by Score AFTER COD.
  		</font>
  	</TD>
  </TR> 
  <TR>
  	<TD align="left">
  		<Font COLOR="#000000" FACE="<%=font1%>" SIZE="<%=fontsize1%>">
  		    <b>UseLCQRank</b> - If checked, the system allows qualification by LCQ Ranking method.
  		</font>
  	</TD>
  </TR> 
  <TR>
  	<TD align="left">
  		<Font COLOR="#000000" FACE="<%=font1%>" SIZE="<%=fontsize1%>">
  		    <b>UseLCQScore</b> - If checked, the system allows qualification by LCQ Score method.
  		</font>
  	</TD>
  </TR> 
  <TR>
  	<TD align="left">
  		<Font COLOR="#000000" FACE="<%=font1%>" SIZE="<%=fontsize1%>">
  		    <b>Qfy_By_AnyOverall_InStates</b> - If checked, the system allows qualification for Regionals by achieving an Overall Score in States.
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

	sSQL = "SELECT * FROM "&LeagueTableName&" WHERE LeagueID='"&sLeagueSelected&"'"
	set rs=Server.CreateObject("ADODB.recordset")
	rs.open sSQL, sConnectionToTRATable, 3, 3	

	sSQL = "UPDATE "&LeagueTableName&" SET "

	' --- Ignores 1st field since it assumes this is the KEY ---
  	FOR i = 1 TO rs.fields.count - 1

		'response.write("<br>"&rs.fields(i).name&" = "&Request.Form(rs.fields(i).name))

		'IF Request.Form(rs.fields(i).name) <> "" AND Request.Form(rs.fields(i).type) <> 11 THEN

			IF RIGHT(sSQL,1) <> "," and RIGHT(sSQL,1) <> " " THEN sSQL = sSQL + ", "

			SELECT CASE rs.fields(i).type
				CASE 11 'boolean'
					sSQL = sSQL + rs.fields(i).name
					IF Request.Form(rs.fields(i).name)="on" THEN
						sSQL = sSQL + "=1"
					ELSE
						sSQL = sSQL + "=0"
					END IF
				CASE 3
					sSQL = sSQL + rs.fields(i).name
					sSQL = sSQL + "=" + sqlclean(Request.Form(rs.fields(i).name))
			
				CASE ELSE
					sSQL = sSQL + rs.fields(i).name
					sSQL = sSQL + "='" + sqlclean(Request.Form(rs.fields(i).name)) + "'"
			END SELECT	
		'END IF


	NEXT      

rs.close
set rs = nothing

sSQL = sSQL + " WHERE LeagueID='"&sLeagueSelected&"'"

'response.write("<br>IN SAVEREC - "&sSQL)
'response.end

OpenCon
con.execute(sSQL)
'WriteLog(date() &"  "& time() &"   "&KeyFieldName&" Record Updated - "& sSQL)
CloseCon

%>
<center><FONT COlOR="red" FACE="<%=font1%>" SIZE="<%=fontsize3%>"><b><i>Your updated record has been saved.</I></b></font></center>
<BR>

<form action="/rankings/<%=ThisFileName%>?action=listrec">
  <center><input type=submit value="Click here to Continue" method="post"></center>
      <input type="hidden" name="sSYID" value="<%=sSYID%>">
      <input type="hidden" name="sTourSelected" value="<%=sTourSelected%>">
      <input type="hidden" name="sLeagueSelected" value="<%=sLeagueSelected%>">
    <input type="hidden" name="sSortField" value="<%=sSortField%>">
</form>
<%
END SUB



'---------------------------
  SUB AddRecord_New 
'---------------------------

' --- First ask if user wants to make the NEW League by copying previous league settings
' --- If YES and PREVIOUS LeagueID was entered and FOUND
' ---	  Ask user to confirm this is the correct PREVIOS LeagueID
' ---		IF YES, then copy all records from previous league
' ---		IF eof, then inform user not found
' --- IF YES and no previous LeagueID was entered by user then RECYCLE message
' --- IF NO and then 

' --- First ask for new LeagueID ---
' --- Then check for existence of that NEW League ID

addstatus=LCASE(LEFT(Request("addstatus"),2))

NewLeagueID=UCASE(TRIM(Request("NewLeagueID")))
NewLeagueName=TRIM(Request("NewLeagueName"))
OldLeagueID=UCASE(TRIM(Request("OldLeagueID")))
CopyQfy=Request("CopyQfy")

'response.write("addstatus="&addstatus)
'response.write("<br>OldLeagueID="&OldLeagueID)
'response.write("<br>NewLeagueID="&NewLeagueID)
'response.write("<br>NewLeagueName="&NewLeagueName)
'response.write("<br>CopyQfy="&CopyQfy)


%>
<FORM action="/rankings/<%=ThisFileName%>">
<input type="hidden" name="action" value="addnew"><%

' --- This section uses the first 2 letters from the button VALUE to determine branching
' --- If button text (Value) is changed you must update branching codes

SELECT CASE addstatus
	CASE ""

		%><center>
		<font size="2"><b>Do you want to create the new league by copying records from another League?</font>
		<br><br>
		<center><INPUT type=submit name="addstatus" value="YES - Copy from previous LeagueID" method="post"></center>
		<br>
		<center><INPUT type=submit name="addstatus" value="NO - Start From Scratch" method="post"></center>
		<br><br><%

	CASE "ye", "ba"
		%><center>
		<font size="2"><b>This function will copy an existing League Setup into a New LeagueID</b></font>
		<br><br>
		<center>
		<font size="2" >Enter the existing LeaqueID from which to copy</font>
		<INPUT type=text name="OldLeagueID" value="<%=OldLeagueID%>" Size=8 MAXLENGTH=8></center>
		<br>

		<center><INPUT type=submit style="width:14em" name="addstatus" value="Continue" method="post"></center><%


	CASE "co", "re"
		

			sSQL = "SELECT * FROM "&LeagueTableName&" WHERE LeagueID='"&OldLeagueID&"'"
			set rs=Server.CreateObject("ADODB.recordset")
			rs.open sSQL, sConnectionToTRATable, 3, 3	

			IF NOT rs.eof THEN
			
					IF UCASE(TRIM(Request("NewLeagueID")))="" THEN 
							NewLeagueID=LEFT(OldLeagueID,7)+CStr(Cdbl(RIGHT(OldLeagueID,1))+1)
					END IF  

					%>
					<center>
					<font size="2"><b>League to be used as template for the new League</b></font>
					<br><br>
					<font size="3" color="red"><b><%=rs("LeagueName")%></b></font>
					<br>
					<font size="2">LeagueID: <%=rs("LeagueID")%></font>
					<br><br>
					<font size="2" >Enter the NEW LeaqueID [format 'AAAA9999']</font>
					<INPUT type=text name="NewLeagueID" value="<%=NewLeagueID%>" Size=8 MAXLENGTH=8>
					<br>
					<font size="2" >Enter the NEW Leaque Name</font>
					<INPUT type=text name="NewLeagueName" value="<%=NewLeagueName%>" Size=35 MAXLENGTH=35>
					<br>
					<font size="2">Check here to include a copy qualifications settings?</font>
					<INPUT TYPE="checkbox" name="CopyQfy" <%IF CopyQfy="on" THEN Response.write( "checked" )%>>
					<br><br><br>
					<INPUT type=submit name="addstatus" value="Proceed to Copy to New League" method="post">
					<input type="hidden" name="OldLeagueID" value="<%=OldLeagueID%>">
					</center>
					<%
		

		ELSE
				%>
				<center>
				<font size="2" ><b>The 'Existing' LeagueID you entered was not found</font>
				<br><br>
				<center><INPUT type=submit style="width:14em" name="addstatus" value="Back" method="post"></center>
				<input type="hidden" name="CopyQfy" value="<%=CopyQfy%>">
				<input type="hidden" name="NewLeagueID" value="<%=NewLeagueID%>">
				<input type="hidden" name="NewLeagueName" value="<%=NewLeagueName%>">
				<input type="hidden" name="OldLeagueID" value="<%=OldLeagueID%>">
				<%
		END IF


	CASE "pr"

			sSQL = "SELECT * FROM "&LeagueTableName&" WHERE LeagueID='"&NewLeagueID&"'"
			set rs=Server.CreateObject("ADODB.recordset")
			rs.open sSQL, sConnectionToTRATable, 3, 3	

			IF NOT rs.eof THEN

					%>
					<center>
					<font size="2" ><b>The 'NEW' LeagueID you entered already exists</font>
					<br><br>
					<center><INPUT type=submit style="width:14em" name="addstatus" value="Re-enter LeagueID" method="post"></center>
					<input type="hidden" name="CopyQfy" value="<%=CopyQfy%>">
					<input type="hidden" name="NewLeagueID" value="<%=NewLeagueID%>">
					<input type="hidden" name="NewLeagueName" value="<%=NewLeagueName%>">
					<input type="hidden" name="OldLeagueID" value="<%=OldLeagueID%>"><%			
			ELSE

					' --- Final Confirmation ---
					sSQL = "SELECT * FROM "&LeagueTableName&" WHERE LeagueID='"&OldLeagueID&"'"
					set rs=Server.CreateObject("ADODB.recordset")
					rs.open sSQL, sConnectionToTRATable, 3, 3	

					%>
					<center>
					<font size="2"><b>This is your FINAL confirmation</b></font>
					<br><br>
					<font size="2"><b>FROM</b></font>
					<br>
					<font size="3" color="red"><b><%=rs("LeagueName")%></b></font>
					<br>
					<font size="2">LeagueID: <%=rs("LeagueID")%></font>

					<br><br>
					<center>
					<font size="2"><b>TO</b></font>
					<br>
					<font size="3" color="red"><b><%=NewLeagueName%></b></font>
					<br>
					<font size="2">LeagueID: <%=NewLeagueID%></font>

					<br><br><br>
					<center><INPUT type=submit name="addstatus" value="Final Confirmation - Execute Copy" method="post"></center>
					<input type="hidden" name="NewLeagueID" value="<%=NewLeagueID%>">
					<input type="hidden" name="NewLeagueName" value="<%=NewLeagueName%>">
					<input type="hidden" name="OldLeagueID" value="<%=OldLeagueID%>">
					<input type="hidden" name="CopyQfy" value="<%=CopyQfy%>"><%


			END IF

	CASE "fi"

			' --- Copies record from LeagueTable ---
			sSQL = "SELECT * FROM "&LeagueQfyTableName&" WHERE LeagueID='"&NewLeagueID&"'"
			set rs=Server.CreateObject("ADODB.recordset")
			rs.open sSQL, sConnectionToTRATable, 3, 3	
		
			IF rs.eof  THEN
					sSQL = "INSERT INTO "&LeagueTableName
					sSQL = sSQL + " (LeagueID, SptsGrpID, LeagueName, Status, SkiYearID, COD, QualifyTour, HomoType, COAMinClass, UseLCQRank, UseLCQScore, RequirePart, Qfy_By_AnyOverall_InStates, State)"
		 			sSQL = sSQL + " SELECT '"&NewLeagueID&"', SptsGrpID, '"&NewLeagueName&"', Status, SkiYearID, COD, QualifyTour, HomoType, COAMinClass, UseLCQRank, UseLCQScore, RequirePart, Qfy_By_AnyOverall_InStates, State"
					sSQL = sSQL + " FROM usawsrank.Leagues WHERE LeagueID='"&OldLeagueID&"'"
					'response.write(sSQL)
					'response.end
					set rs=Server.CreateObject("ADODB.recordset")
					rs.open sSQL, sConnectionToTRATable, 3, 3	
			END IF

			' --- If the NEWLeague table data doesn't already exist, then copy the League Qualification records ---
			sSQL = "SELECT * FROM "&LeagueQfyTableName&" WHERE LeagueID='"&NewLeagueID&"'"
			set rs=Server.CreateObject("ADODB.recordset")
			rs.open sSQL, sConnectionToTRATable, 3, 3	
			
			IF rs.eof AND CopyQfy="on" THEN

				sSQL="INSERT INTO "&LeagueQfyTableName 
		 		sSQL=sSQL+" SELECT '"&NewLeagueID&"', Event, Div, COA, Level_A, Level_B, Place_TourA, Place_TourB, Place_TourC, LevelBy3rdEvt, Place_TourD"
	 			sSQL=sSQL+" FROM "&LeagueQfyTableName&" WHERE LeagueID='"&OldLeagueID&"'"
				'response.write(sSQL)
				'response.end
				set rs=Server.CreateObject("ADODB.recordset")
				rs.open sSQL, sConnectionToTRATable, 3, 3	


				' --- Zeros COA for all Div/Events
				sSQL="UPDATE LQ SET COA=0 FROM "&LeagueQfyTableName&" AS LQ" 
	 			sSQL=sSQL+" WHERE LeagueID='"&NewLeagueID&"'"
				set rs=Server.CreateObject("ADODB.recordset")
				rs.open sSQL, sConnectionToTRATable, 3, 3	


			END IF

			%>
			<center>
			<font size="3" ><b>Copy Complete</b></font>
			<br><br>
		
			<font size="2" >Proceed to Editing this League Set-up</font>
			<br><br><br>
			<INPUT type=submit style="width:14em" name="addstatus" value="Edit LeagueID" method="post">
			</center>
			<input type="hidden" name="NewLeagueID" value="<%=NewLeagueID%>">
			<input type="hidden" name="OldLeagueID" value="<%=OldLeagueID%>">
			<%


	CASE "no"

		%><center>
		<font size="2"><b>This operation will create a new LeagueID record in the League table</b></font>

		<br><br>
		<font size="2" >Enter the NEW LeaqueID [format 'AAAA9999']</font>
		<INPUT type=text name="NewLeagueID" value="<%=NewLeagueID%>" Size=8 MAXLENGTH=8></center>
		<br><br><br>
		<center><INPUT type=submit style="width:14em" name="addstatus" value="Create New LeagueID" method="post"></center><%

	CASE "cr"

		sSQL = "SELECT * FROM "&LeagueTableName&" WHERE LeagueID='"&NewLeagueID&"'"
		set rs=Server.CreateObject("ADODB.recordset")
		rs.open sSQL, sConnectionToTRATable, 3, 3	

		IF NOT rs.eof THEN

			%><center>
			<font size="2" ><b>The 'NEW' LeagueID you entered already exists</font>
			<br><br>
			<center><INPUT type=submit style="width:14em" name="addstatus" value="Re-enter LeagueID" method="post"></center>
			<input type="hidden" name="NewLeagueID" value="<%=NewLeagueID%>"><%			
		ELSE

			' --- Copies record from LeagueQfyTable ---
			sSQL = "SELECT * FROM "&LeagueQfyTableName&" WHERE LeagueID='"&NewLeagueID&"'"
			set rs=Server.CreateObject("ADODB.recordset")
			rs.open sSQL, sConnectionToTRATable, 3, 3	

		
			IF rs.eof  THEN
				sSQL="INSERT INTO "&LeagueTableName
		 		sSQL=sSQL+" (LeagueID, SptsGrpID)"
				sSQL=sSQL+" VALUES ('"&NewLeagueID&"', 'AWS')"
				set rs=Server.CreateObject("ADODB.recordset")
				rs.open sSQL, sConnectionToTRATable, 3, 3	
			END IF  %>

			<center>
			<font size="3" ><b>New LeagueID Record Added</b></font>
			<br><br>

			<font size="2" >Proceed to Editing this League Set-up</font>
			<br><br><br>
			<INPUT type=submit style="width:14em" name="addstatus" value="Edit LeagueID" method="post">
			</center>
			<input type="hidden" name="NewLeagueID" value="<%=NewLeagueID%>">


			<%

		END IF

	CASE "ed"

		response.redirect("/rankings/"&ThisFileName&"?action=editrec&sLeagueSelected="&NewLeagueID)

	CASE ELSE
		WriteButton "?action=addrec&addstatus=ELSE","Else","<BR><BR>"
		
END SELECT %>


</form><%



END SUB







'-------------------
  SUB DeleteRec
'-------------------

	sSQL = "SELECT * FROM "&LeagueTableName&" WHERE LeagueID='"&sLeagueSelected&"'"
	set rs=Server.CreateObject("ADODB.recordset")
	rs.open sSQL, sConnectionToTRATable, 3, 3	



IF NOT rs.eof AND LCASE(Request("confirm")) = "yes" THEN
	
    	'--- Delete the record'
	sSQL = "DELETE FROM "&LeagueTableName&" WHERE LeagueID='"&sLeagueSelected&"'"
	set rs=Server.CreateObject("ADODB.recordset")
	rs.open sSQL, sConnectionToTRATable, 3, 3	
	
	sSQL = "DELETE FROM "&LeagueQfyTableName&" WHERE LeagueID='"&sLeagueSelected&"'"
	set rs=Server.CreateObject("ADODB.recordset")
	rs.open sSQL, sConnectionToTRATable, 3, 3	
	%>

	<center>
	<font SIZE="3"><b>The record has been deleted.</b></font>
	</center>
	<BR><BR><%

END IF


WriteButton "?action=listrec","Return To League List","<BR><BR>"

IF LCASE(Request("confirm")) = "" THEN  %>  
	<br><br>
	<center>
	<font size="2">Type the word "YES" if you are sure you wish to delete this record. </font>
	<br>
	<font solor="red" size="2">Qualifications for this League will be affected.</font>
	<br><br>
	<form action="/rankings/<%=ThisFileName%>" method="post"> 
	  <input type="hidden" name="action" value="delrec">
	  <input type="hidden" name="sTourSelected" value="<%=sTourSelected%>">
	  <input type="hidden" name="sLeagueSelected" value="<%=sLeagueSelected%>">
	  <input type="hidden" name="sSYID" value="<%=sSYID%>">
	  <input type="text" name="confirm" Size="4" MAXLENGTH="3">
	  <input type="submit" value="Confirm Deletion?">
	</form><%

	WriteButton "?action=listrec","NO - Do not delete the record","<BR><BR>"

END IF

IF (LCASE(Request("confirm")) <> "yes" AND LCASE(Request("confirm")) <> "") THEN %>  
	<center>
	<br><br>
	<font size="3">The record was NOT deleted.</font>
        <br><br>
	</center><%
END IF


END SUB



'---------------------
  Function GetCheckValue(i)
'---------------------

'response.write("INSIDE")
'response.write(rs.fields(i).value)



    IF LCASE(sAction) = "editrec" THEN
        IF rs.fields(i).value = True THEN
            GetCheckValue = "checked"
        ELSE
            GetCheckValue = ""
        END IF
    ELSE
            GetCheckValue = ""
    END IF

'response.write(GetCheckValue)
'response.end

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
SUB WriteType (I, FieldLen)
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


	CASE ELSE

		SELECT CASE Rs.Fields(i).type
			CASE 3 'primary key / auto number ?'
				%><input type="text" name="<% Response.Write(Rs.Fields(i).name) %>" SIZE="25" value="<% GetFieldValue i %>"><%
			CASE 20 'primary key / auto number ?'
				%><input type="text" name="<% Response.Write(Rs.Fields(i).name) %>" SIZE="25" value="<% GetFieldValue i %>"><%
			CASE 11 'boolean'
				%><INPUT TYPE="checkbox" NAME="<% Response.Write(Rs.Fields(i).name) %>"<%IF Rs.Fields(i).value = True THEN Response.Write("Checked")%>><%
			CASE 129 'char'
				%><input type="text" name="<% Response.Write(Rs.Fields(i).name) %>" value="<% GetFieldValue i %>" SIZE="<%=FieldLen%>" MAXLENGTH="<%=FieldLen%>"><%
			CASE 203 'memo'
        			%><TEXTAREA NAME="<% Response.Write(Rs.Fields(i).name) %>" ROWS="20" COLS="56"><% GetFieldValue i %></TEXTAREA><%
			CASE ELSE 'not handled by this function'
			        %><input type="text" name="<% Response.Write(Rs.Fields(i).name) %>" SIZE="<%=FieldLen%>" MAXLENGTH="<%=FieldLen%>" value="<% GetFieldValue i %>"><%
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