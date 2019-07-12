<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->
<!--#include virtual="/rankings/tools_Definitions.asp"--><%

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
Dim sLeagueSelected, sTourSelected, sTourName, sRegionSelected, sEventSelected, sDivSelected
Dim sTSptsGrpID
Dim sLevel_A, sLevel_B, sLevelBy3rdEvt, sPlace_TourA, sPlace_TourB, sPlace_TourC, sPlace_TourD

Dim ThisFileName
ThisFileName="/rankings/EditLeagueQualifications.asp"


sTSptsGrpID="AWS"

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


'session("sSptsGrpID") = "TEST"
Session("sSptsGrpID")="AWS"
Session("UserSptsGrpID")="AWS"
'Session("adminmenulevel")=50

IF request("sSptsGrpID") <>"" THEN session("sSptsGrpID") = UCASE(request("sSptsGrpID"))


IF session("sSptsGrpID") = "TEST" THEN
	ThisTempTableName = "USAWSRank.LeagueQualifications"
ELSE
	IF session("sSptsGrpID") = "AWS" OR session("sSptsGrpID") = "NCW" THEN
		ThisTempTableName = "USAWSRank.LeagueQualifications"
	ELSEIF session("sSptsGrpID")="USW" OR session("sSptsGrpID")="HYD" OR session("sSptsGrpID")="AKA" THEN
		ThisTempTableName = "USAWSRank.LeagueQualifications"
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
SELECT CASE TRIM(sAction)
	CASE "Add Event/Div"
		sAction = "dispselector"
	CASE "Add Record(s)"
		sAction = "addrec"
	CASE "Update Listing", "No, List Records"
		sAction = "listrec"
	CASE "Delete All"
		sAction = "godelete"
	CASE "Delete Recs"
		sAction = "delallrecs"

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



sListFilter= request("ListFilter")
IF sListFilter<>"" THEN Session("ListFilter")=sListFilter


sEditFld = session("EditFldFilter")
IF sEditFld = "" then sEditFld = "ALL"

sSortField=TRIM(Request("sSortField"))
'response.write("sSortField="&sSortField)

IF sSortField = "" then sSortField = "LQ.Event"

sRegionSelected=TRIM(Request("RegionSelected"))
IF sRegionSelected = "" then sRegionSelected = "ALL"

sEventSelected=TRIM(Request("sEventSelected"))
IF sEventSelected = "" THEN sEventSelected = "ALL"

sDivSelected=TRIM(Request("sDivSelected"))
IF sDivSelected = "" THEN sDivSelected = "ALL"

sLeagueSelected=TRIM(Request("sLeagueSelected"))
IF TRIM(sLeagueSelected) = "" THEN sLeagueSelected = "ALL"

sLevel_A=TRIM(Request("sLevel_A"))
sLevel_B=TRIM(Request("sLevel_B"))
sLevelBy3rdEvt=TRIM(Request("sLevelBy3rdEvt"))
sPlace_TourA=TRIM(Request("sPlace_TourA"))
sPlace_TourB=TRIM(Request("sPlace_TourB"))
sPlace_TourC=TRIM(Request("sPlace_TourC"))
sPlace_TourD=TRIM(Request("sPlace_TourD"))

IF sLevel_A="" THEN sLevel_A=0
IF sLevel_B="" THEN sLevel_B=0
IF sLevelBy3rdEvt="" THEN sLevelBy3rdEvt=0
IF sPlace_TourA="" THEN sPlace_TourA=0
IF sPlace_TourB="" THEN sPlace_TourB=0
IF sPlace_TourC="" THEN sPlace_TourC=0
IF sPlace_TourD="" THEN sPlace_TourD=0

IF sAction = "" THEN sAction = "listrec"


ThisPage = Request.ServerVariables("SCRIPT_NAME")



'response.write("<br>sDivSelected="&sDivSelected)
'response.write("<br>sEventSelected="&sEventSelected)





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



'response.write("sAction="&sAction)
'response.end


'-Main Page Code-----------------------------------'



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
	


	ChooseSQL(sSQL)

	sLeagueName=rs("LeagueName")
	WriteHeaders "Event/Divs in LeagueID '"&sLeagueSelected&"' - "&sLeagueName&" "
	ListRecords


  CASE "dispselector"
	IF sThisFieldID = "" THEN sThisFieldID = Session("NewKeyCode")

    	WriteHeaders "Add Records:  LeagueID = "&sLeagueSelected&" - Event(s)="&sEventSelected&" - Div(s)="&sDivSelected&"<br> into "&ThisTempTableName

    	IF sLeagueSelected = "ALL" THEN %>
		<br><H2><center><font color="red">You must select a League into which this Event/Div will be added.<BR></font></center></H2><% 	
    		Listrecords

	ELSE
		' --- Displays the various fields for selecting default values.
		ShowValueSelector
  	END IF

  CASE "addrec"
    	WriteHeaders "Add the following records:  for SYID = "&sSYID&" and Sports Discipline = "&session("sSptsGrpID")&" in " & ThisTempTableName
    	AddAllRecords
	Listrecords

  CASE "editrec"
	WriteHeaders "Edit LeaqueID Record:  Event = "&sEventSelected&" - Div = "&sDivSelected&" for SYID = "&sSYID&" and Sports Discipline = "&session("sSptsGrpID")&" in " & ThisTempTableName
	ShowEditor

  CASE "saverec"
	WriteHeaders "Record Saved :  Event = "&sEventSelected&" - Div = "&sDivSelected&" for SYID = "&sSYID&" and Sports Discipline = "&session("sSptsGrpID")&" in " & ThisTempTableName
	SaveRec

  CASE "delrec"
    	IF sLeagueSelected <> "ALL" THEN 
		WriteHeaders "Delete Event = "&sEventSelected&" - Div = "&sDivSelected&" from LeagueID '"&sLeagueSelected&"' <br>in table " & ThisTempTableName
	ELSE %>
		<br><H2><center><font color="red">You must select a League into which this Event/Div will be added.<BR></font></center></H2><% 	
    		Listrecords
	END IF

	DeleteRec

  CASE "delallrecs"

    	IF sLeagueSelected = "ALL" THEN %>
		<br><H2><center><font color="red">You must select a League to Delete Event/Div records.<BR></font></center></H2><% 	
    		Listrecords

	ELSE
		' --- Displays the various fields for selecting default values.
		ShowDeleteConfirm
  	END IF

  CASE "godelete"
    	WriteHeaders "Deleted ALL Event/Div records for LeagueID '"&sLeagueSelected&"' in " & ThisTempTableName
    	DeleteAllRecords
	Listrecords
	


END SELECT


WriteFooter
WriteIndexPageFooter






' ------------------------
  SUB WriteHeaders(sTitle)
' ------------------------

' Write Headers for Page
%>
<br>
<TABLE class="innertable" align="center" WIDTH="<%=TourTableWidth%>">
  <tr>
    <th align="center">
	<font color="#FFFFFF" size="2"><B><%= sTitle %></B></font>
    </th>
  </tr>
</TABLE>
<BR><%


END SUB


' -----------------------
  SUB DeleteAllRecords
' -----------------------

OpenCon
sSQL = "DELETE FROM "&LeagueQfyTableName&" WHERE LeagueID='"&sLeagueSelected&"'"

'response.write(sSQL)
'response.end

con.execute(sSQL)
Closecon

END SUB




' -----------------------
  SUB ShowDeleteConfirm
' -----------------------

WriteHeaders "You are about to DELETE ALL Event/Div records from LeagueID '"&sLeagueSelected&"' <br>in table " & ThisTempTableName

%>
<TABLE class="innertable" align="center" width="450px">
  <form action="<%=ThisFileName%>" method="post">

  <tr>
    <td align="center">
	<input type="submit" style="width:9em" name="action" value="Delete All">
	<input type="hidden" name="sLeagueSelected" value="<%=sLeagueSelected%>">
    </td>
    <td align="center">
	<input type="submit" style="width:12em" name="action" value="No, List Records">
    </td>
  </tr>
  </form>
</TABLE>
<br><%			



END SUB



' -----------------------
  SUB ShowValueSelector
' -----------------------

%>

<TABLE class="innertable" align="center" width="450px">
  <form action="<%=ThisFileName%>" method="post">
  <tr>
    <th align="right" width 50%><font size="<%=fontsize2%>" color="#FFFFFF"><b>Description</b>&nbsp;&nbsp;</font></th>
    <th align="left" width 50%><font size="<%=fontsize2%>" color="#FFFFFF"><b>&nbsp;&nbsp;Default Value</b></font></th>
  </tr>
  <tr>
    <td align="right" width 50%><font size="<%=fontsize2%>">Division(s)</font></td>
    <td align="left" width=50%><%
	' --- SUB located in tools_include.asp  ---

'response.write("<br> 1 sDivSelected="&sDivSelected)

	LoadDivDropNoAgeGender sDivSelected, sEventSelected, "sDivSelected", "enabled"  %>
    </td>
  </tr>
  <tr>	
    <td align="right"><font size="<%=fontsize2%>">Event(s)</font></td>
    <td align="left">
    	<% 
			BuildAWSAEvents 
			%>			
    </td>
  </tr>
  <tr>	
    <td align="right"><font size="<%=fontsize2%>">Nationals Level (Level_A)</font></td>
    <td align="left">
    	<%
			LoadValuePulldown "sLevel_A", sLevel_A, 0, 8, 1, "enabled", false  
			%>
    </td>
  </tr>
  <tr>	
    <td align="right"><font size="<%=fontsize2%>">Regionals Level (Level_B)</font></td>
    <td align="left">
    	<%
			LoadValuePulldown "sLevel_B", sLevel_B, 0, 8, 1, "enabled", false  
			%>
    </td>
  </tr>
  <tr>	
    <td align="right"><font size="<%=fontsize2%>">3rd Event Level (LevelBy3rdEvt)</font></td>
    <td align="left">
    	<%
			LoadValuePulldown "sLevelBy3rdEvt", sLevelBy3rdEvt, 0, 8, 1, "enabled", false  
			%>
    </td>
  </tr>
  <tr>	
    <td align="right"><font size="<%=fontsize2%>">Nationals Placement (Place_TourA)</font></td>
    <td align="left">
    	<%
			LoadValuePulldown "sPlace_TourA", sPlace_TourA, 0, 5, 1, "enabled", false  
			%>
    </td>
  </tr>
  <tr>	
    <td align="right"><font size="<%=fontsize2%>">Regionals Placement (Place_TourB)</font></td>
    <td align="left">
    	<%
			LoadValuePulldown "sPlace_TourB", sPlace_TourB, 0, 5, 1, "enabled", false  
			%>
    </td>
  </tr>
  <tr>	
    <td align="right"><font size="<%=fontsize2%>">State Placement (Place_TourC)</font></td>
    <td align="left">
    	<%
			LoadValuePulldown "sPlace_TourC", sPlace_TourC, 0, 5, 1, "enabled", false  
			%>
    </td>
  </tr>
  <tr>	
    <td align="right"><font size="<%=fontsize2%>">Other Placement (Place_TourD)</font></td>
    <td align="left">
    	<%
			LoadValuePulldown "sPlace_TourD", sPlace_TourD, 0, 5, 1, "enabled", false  
			%>
    </td>
  </tr>
</table>
<br>
<table width="70%" align="center" border="0px"> 
  <tr>
    <td align="center">
			<input type="submit" style="width:12em" name="action" value="Add Record(s)">
			<input type="hidden" name="sLeagueSelected" value="<%=sLeagueSelected%>">
    </td>
    <td align="center">
			<input type="submit" style="width:12em" name="action" value="No, List Records">
    </td>
  </tr>
  </form>
</TABLE>
<br>


<TABLE BORDER="0" align=center WIDTH=60% CELLPADDING="3" CELLSPACING="0">
	<tr>
		<td align="left" style="word-wrap:break-word">
			<font face=<%=font1%> size=<%=fontsize1%> color=<%=textcolor1%>>
			<b>Procedure when some Divisions do NOT have qualification requirements</b>
			<br><br>STEP 1 - Add the divisions one at a time that HAVE a qualification requirement. 
			<br><br>STEP 2 - Add the remainder of the divs/events offered (those which DO NOT HAVE qualifications) by setting Level_A to zero (0) 
				and using All Events and All Divisions setting.  It will populate all divs/events that do not already have a qualification 
				assigned.  <br><br>NOTE: Divisions where no qualification requirement has been set will not display in the qualifications system.
		</td>
	</tr>
</TABLE>
<br>
<%			


END SUB



' ------------------
  SUB AddAllRecords
' ------------------

sDefCOD="07/14/2008"

'response.write("sLeagueSelected="&sLeagueSelected)

' --- REMOVE ---
'sLeagueSelected="NATL2008"


' --- Creates ALL records in LeagueQualify for this TourID from Divisions
OpenCon

sSQL = "INSERT INTO "&LeagueQfyTableName&" (LeagueID, Event, Div, Level_A, Level_B, LevelBy3rdEvt, Place_TourA, Place_TourB, Place_TourC, Place_TourD)"
IF TRIM(sEventSelected)="S" OR TRIM(sEventSelected)="ALL" THEN
  sSQL = sSQL + " (  SELECT '"&sLeagueSelected&"' AS LeagueID, 'S' AS Event, DT.Div AS Div, '"&sLevel_A&"' AS Level_A, '"&sLevel_B&"' AS Level_B"
  sSQL = sSQL + " , '"&sLevelBy3rdEvt&"' AS LevelBy3rdEvt, '"&sPlace_TourA&"' AS Place_TourA, '"&sPlace_TourB&"' AS Place_TourB, '"&sPlace_TourC&"' AS Place_TourC, '"&sPlace_TourD&"' AS Place_TourD" 
  sSQL = sSQL + "	FROM "&DivisionsTableName&" AS DT"
  sSQL = sSQL + "	WHERE DT.Div NOT IN (SELECT Div FROM "&LeagueQFyTableName&" WHERE LeagueID='"&sLeagueSelected&"' AND Event='S')"
  sSQL = sSQL + "	AND DT.SkiYearID=1 AND UPPER(LEFT(DT.Div,1)) NOT IN ('Y','X','C','I','N','E','L','J','S')" 
  IF TRIM(sDivSelected)<>"ALL" THEN sSQL = sSQL + " AND DT.Div='"&sDivSelected&"'"  
  sSQL = sSQL + ")"
END IF

IF TRIM(sEventSelected)="ALL" THEN
	sSQL = sSQL + " UNION"
END IF

IF TRIM(sEventSelected)="T" OR TRIM(sEventSelected)="ALL" THEN
  sSQL = sSQL + " (  SELECT '"&sLeagueSelected&"' AS LeagueID, 'T' AS Event, DT.Div AS Div, '"&sLevel_A&"' AS Level_A, '"&sLevel_B&"' AS Level_B"
  sSQL = sSQL + " , '"&sLevelBy3rdEvt&"' AS LevelBy3rdEvt, '"&sPlace_TourA&"' AS Place_TourA, '"&sPlace_TourB&"' AS Place_TourB, '"&sPlace_TourC&"' AS Place_TourC, '"&sPlace_TourD&"' AS Place_TourD" 
  sSQL = sSQL + "	FROM "&DivisionsTableName&" AS DT"
  sSQL = sSQL + "	WHERE DT.Div NOT IN (SELECT Div FROM "&LeagueQFyTableName&" WHERE LeagueID='"&sLeagueSelected&"' AND Event='T')"
  sSQL = sSQL + "	AND DT.SkiYearID=1 AND UPPER(LEFT(DT.Div,1)) NOT IN ('Y','X','C','I','N','E','L','J','S')" 
  IF TRIM(sDivSelected)<>"ALL" THEN sSQL = sSQL + " AND DT.Div='"&sDivSelected&"'"  
  sSQL = sSQL + ")"
END IF

IF TRIM(sEventSelected)="ALL" THEN
	sSQL = sSQL + " UNION"
END IF


IF TRIM(sEventSelected)="J" OR TRIM(sEventSelected)="ALL" THEN
  sSQL = sSQL + " (  SELECT '"&sLeagueSelected&"' AS LeagueID, 'J' AS Event, DT.Div AS Div, '"&sLevel_A&"' AS Level_A, '"&sLevel_B&"' AS Level_B"
  sSQL = sSQL + " , '"&sLevelBy3rdEvt&"' AS LevelBy3rdEvt, '"&sPlace_TourA&"' AS Place_TourA, '"&sPlace_TourB&"' AS Place_TourB, '"&sPlace_TourC&"' AS Place_TourC, '"&sPlace_TourD&"' AS Place_TourD" 
  sSQL = sSQL + "	FROM "&DivisionsTableName&" AS DT"
  sSQL = sSQL + "	WHERE DT.Div NOT IN (SELECT Div FROM "&LeagueQFyTableName&" WHERE LeagueID='"&sLeagueSelected&"' AND Event='J')"
  sSQL = sSQL + "	AND DT.SkiYearID=1 AND UPPER(LEFT(DT.Div,1)) NOT IN ('Y','X','C','I','N','E','L','J','S')" 
  IF TRIM(sDivSelected)<>"ALL" THEN sSQL = sSQL + " AND DT.Div='"&sDivSelected&"'"  
  sSQL = sSQL + ")"
END IF

IF TRIM(sEventSelected)="ALL" THEN
	sSQL = sSQL + " UNION"
END IF

IF TRIM(sEventSelected)="O" OR TRIM(sEventSelected)="ALL" THEN
  sSQL = sSQL + " (  SELECT '"&sLeagueSelected&"' AS LeagueID, 'O' AS Event, DT.Div AS Div, '"&sLevel_A&"' AS Level_A, '"&sLevel_B&"' AS Level_B"
  sSQL = sSQL + " , '"&sLevelBy3rdEvt&"' AS LevelBy3rdEvt, '"&sPlace_TourA&"' AS Place_TourA, '"&sPlace_TourB&"' AS Place_TourB, '"&sPlace_TourC&"' AS Place_TourC, '"&sPlace_TourD&"' AS Place_TourD" 
  sSQL = sSQL + "	FROM "&DivisionsTableName&" AS DT"
  sSQL = sSQL + "	WHERE DT.Div NOT IN (SELECT Div FROM "&LeagueQFyTableName&" WHERE LeagueID='"&sLeagueSelected&"' AND Event='O')"
  sSQL = sSQL + "	AND DT.SkiYearID=1 AND UPPER(LEFT(DT.Div,1)) NOT IN ('Y','X','C','I','N','E','L','J','S')" 
  IF TRIM(sDivSelected)<>"ALL" THEN sSQL = sSQL + " AND DT.Div='"&sDivSelected&"'"  
  sSQL = sSQL + ")"
END IF



'response.write("<br>"&sSQL)
'response.end

con.execute(sSQL)
Closecon


WriteLog(date() &"  "& time() &"   New Event/Div Has Been Added to LeagueQualify - "& sSQL)


END SUB



' ----------------
  SUB ListRecords
' ----------------

'  Lists the table Records

%> 
<form action="<%=ThisFileName%>" method="post">
<input type="hidden" name="search" value="1">

<TABLE align="center" class="innertable" WIDTH="<%=TourTableWidth%>">
<tr>
  <th align="center"><font color="#FFFFFF" size="<%=fontsize1%>">League ID</font></th>
  <th align="center"><font color="#FFFFFF" size="<%=fontsize1%>">League Ski Year</font></th>
  <th align="center"><font color="#FFFFFF" size="<%=fontsize1%>">Sort By</font></th>
  <th align="center"><font color="#FFFFFF" size="<%=fontsize1%>">SD</font></th>
</tr>

<tr>
<td align="center">
<%

' --- Builds list of League from Master table LEAGUES ---
LeagueDropBuild




' ------------   Gets NEW CODE  -----------------
%>
</td>

  <td align="center"><%
	' --- Builds Ski Year Drop down based on Ski Year table ---
	SkiYearDropBuild %>
  </td>

  <td align="center">
	<select name="sSortField">
		<option value="LQ.Event" <%IF sSortField = "LQ.Event" THEN Response.Write("selected")%>>Event</option>
		<option value="LQ.Div" <%IF sSortField = "LQ.Div" THEN Response.Write(" selected ")%>>Division</option>
    	</select>
  </td>

  <td align="center"><font color="#000000" FACE="<%=font1%>" size="<%=fontsize1%>"><%=Session("sSptsGrpID")%></font></td>
</tr>

</table>

<br>
<TABLE width="<%=TourTableWidth%>" align=center>
<tr>
  <td align=center style="border-style:none;">
	<input type="submit" style="width:12em" name="action" value="Update Listing">
  </td>
  <td align=center style="border-style:none;">	
	<input type="submit" style="width:12em" name="action" value="Add Event/Div">
  </td>
  <td align=center style="border-style:none;">	
	<input type="submit" style="width:12em" name="action" value="Delete Recs">
  </td>
</form>

<form action="/rankings/defaultHQ.asp" >
  <td align=center style="border-style:none;">
	<input type=submit style="width:12em" value=" Main Menu" method="post">
  </td>
</form>

</tr>
</TABLE>
<%


'response.write("<br>sSYID="&sSYID)

' --- First find the two-degit year for SkiYearSelected ---
sSQL="SELECT RIGHT(SkiYear,2) AS ThisYear FROM "&SkiYearTableName&" AS ST WHERE ST.SkiYearID = '"&sSYID&"'"
ChoosePagesSQL sSQL,currentPage, 20

rs.movefirst


ThisYear=rs("ThisYear")


' -------------  Run Query to League Event/Div List Recordset  ----------------------
  sSQL = "SELECT LQ.LeagueID, LQ.Event, LQ.Div, LQ.COA, LQ.Level_A AS Nationals, LQ.Level_B AS Regionals, LQ.LevelBy3rdEvt AS [3rd Event]"
  sSQL = sSQL + ", LQ.Place_TourA AS Natls, LQ.Place_TourB AS Regls, LQ.Place_TourC AS State, LQ.Place_TourD AS Other" 
  sSQL = sSQL + " FROM "&LeagueQfyTableName&" AS LQ"
  IF sLeagueSelected<>"ALL" THEN sSQL = sSQL + " WHERE LQ.LeagueID = '"&sLeagueSelected&"'" 
  IF sSortField="LQ.Event" THEN
			sSQL = sSQL + " ORDER BY LQ.Event, LQ.Div"
  ELSEIF sSortField="LQ.Div" THEN
			sSQL = sSQL + " ORDER BY LQ.Div, LQ.Event"
  END IF

  ChoosePagesSQL sSQL,currentPage, 20




rowCount = 0

' ---------------  Displays table HEADINGS  ----------------------

%>
<BR>
<TABLE class="innertable" align=center WIDTH="<%=TourTableWidth%>" >
  <tr>
    <th align="center" colspan=6 style="background-color:#ffffff; border:0px solid">&nbsp;</th>
    <th align="center" colspan=3><font color="#FFFFFF" size="<%=fontsize1%>">Qualify Level</font></th>
    <th align="center" colspan=4><font color="#FFFFFF" size="<%=fontsize1%>">Placement</font></th>    
  </tr>
  <tr>
    <th align="center"><font color="#FFFFFF" size="<%=fontsize1%>">Delete</font></th>
    <th align="center"><font color="#FFFFFF" size="<%=fontsize1%>">Edit</font></th>
<%

FOR i = 0 TO rs.fields.count - 1
		TempFN = rs.fields(i).name
		j = 0 
		' Rs.Fields(i).name
		%>
   	<th align="center" valign="top" nowrap>
	  	<font color="#FFFFFF" FACE="<%=font1%>" size="<%=fontsize1%>"><%=Rs.Fields(i).name%></font>
		</th><%
NEXT

%>
</tr>
<%


IF sLeagueSelected="ALL" THEN  %>
	</TABLE>
	<TABLE align="center">
	<tr><td align="center">
		<br><br>
		<font size="3" color="red">Please select a LeagueID</font>
	</td></tr><%	
ELSE 


' --------------  Display table data here with paging --------------------------

    DO WHILE NOT rs.eof

	IF rowCount = rs.PageSize THEN EXIT DO

	%>
 	<tr>
	<td align="center" valign="top"><font color="#FFFFFF" size="<%=fontsize1%>"><% WriteLink "?action=delrec&sEventSelected="&TRIM(rs("Event"))&"&sDivSelected="&rs("Div")&"&sLeagueSelected="&rs("LeagueID")&"&sSYID="&sSYID,"Delete","" %></font></td><%
	AllowEdit=true
	IF AllowEdit=true THEN %>
		<td align="center"><font color="#000000" size="<%=fontsize1%>"><% WriteLink "?action=editrec&sEventSelected="&TRIM(rs("Event"))&"&sDivSelected="&rs("Div")&"&sLeagueSelected="&rs("LeagueID")&"&sSYID="&sSYID,"Edit","" %></font></td><%
	ELSE %>
		<td align="center"><font color="#000000" size="<%=fontsize1%>">Edit</font></td><%
	END IF 

	FOR i = 0 TO rs.fields.count - 1
	
		Rowcolor=""
		TempFN = rs.fields(i).name
		
		%><td align="center" style="<%=RowColor%>">
			<font color="#000000" size="<%=fontsize1%>">&nbsp;<%

		    IF isnull(rs.Fields(i).value) THEN
			response.write ("&nbsp;")
		    ELSEIF rs.fields(i).name="COA" THEN
			Response.Write(formatnumber(trim(Rs.Fields(i).Value),2)) 
    		    ELSE
			Response.Write(trim(Rs.Fields(i).Value)) 
		    END IF  

		%>&nbsp;
		  </font>
		</td><%
	NEXT %>

	</tr><% 
	rowCount = rowCount + 1
	rs.movenext

    LOOP

END IF
%>
</TABLE>
<br>
<%


DoCount currentPage

rs.close
set rs = nothing




ExcludeThis= "Y"
IF ExcludeThis= "Y" THEN 
%>
<TABLE BORDER="0" align=left WIDTH=100% CELLPADDING="3" CELLSPACING="0" style="padding-bottom:15px;">
	<tr>
		<td align="left" style="word-wrap:break-word">
			<font face=<%=font1%> size=<%=fontsize1%> color=<%=textcolor1%>>
			<b>GENERAL INFORMATION</b>
			<br><br>1) COA will fluctuate constantly until the qualify recalculation is run on the COD, after which it will be locked. 
			<br><br>2) The Placement Other field may be used for tournaments like Junior Development and similar
		</td>
	</tr>
</TABLE>
<% 

END IF



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

<SELECT name="sSYID">
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



WriteButton "?action=listrec","No Change - Go To League List","<br><br>"

sSQL = "SELECT * FROM "&LeagueQfyTableName&" WHERE LeagueID='"&sLeagueSelected&"' AND Event='"&sEventSelected&"' AND Div='"&sDivSelected&"'"
'response.write("<br> sSQL 929 = "&sSQL)

ChooseSQL(sSQL)



%>
<FORM METHOD="POST" ACTION="<% Response.Write(ThisPage) %>?action=saverec">
<TABLE class="innertable" align="center" style="border:1px solid; margin-top:10px;" >
  <tr>
    <th align="Left"><font size="<%=fontsize1%>" color="#FFFFFF"><B>Field</B></font></th>
    <th align="Left"><font size="<%=fontsize1%>" color="#FFFFFF"><B>Value</B></font></th>
    <th align="center"><font size="<%=fontsize1%>" color="#FFFFFF"><B>Field Len</B></font></th>
  </tr>

<%

' *** Important -- first two fields are Code (ID) and Ski Year ID code.
' *** These two fields serve as the record key, and hence are NOT editable.

FOR i = 0 TO rs.fields.count - 1

		TempFN = rs.fields(i).name
		ThisDesc=""
		SELECT CASE trim(lcase(Rs.Fields(i).name))
				CASE "level_a"
						ThisDesc = "Nationals Level"
				CASE "level_b"
						ThisDesc = "Regionals Level"
				CASE "place_toura"
						ThisDesc = "Nationals Placement"
				CASE "place_tourb"
						ThisDesc = "Regionals Placement"
				CASE "place_tourc"
						ThisDesc = "State Placement"
				CASE "place_tourd"
						ThisDesc = "Other Tournament Placement"
				CASE "levelby3rdevt"
						ThisDesc = "Level to Drag 3rd Event"
				CASE ELSE
						ThisDesc = Rs.Fields(i).name
			END SELECT				
						
		' Rs.Fields(i).name
		%>
		<tr>
	  	<td align="Left" width="180px">
				<font size="<%=fontsize1%>"><B><% Response.Write(ThisDesc) %></B></font>
	  	</td>
	   	<td align="Left" width="230px">
	   		<%
				' IF i = 0 OR i = 1 THEN 
				IF i = 1 THEN 
						%>
						<input type="hidden" name="LeagueID" value="<%=rs.fields(i).value%>">
						<font size="2"><%=rs.fields(i).value %></font>
						<%
				ELSEIF i = 0 THEN
						%><font size="2"><%=rs.fields(i).value %></font><%
				ELSE 
						%><font color="#000000" FACE="<%=font1%>" size="<%=fontsize1%>"><% WriteType i %></font></td><%
				END IF 
				%>
	    <td align=center width="80px"><%= len(rs.fields(i).value) %></td>
		</tr>
		<%

NEXT




%>
</TABLE><BR>
<TABLE BORDER="0" align="center" style="width:50%; padding-bottom:10px;" CELLPADDING="3" CELLSPACING="0">
<tr>
<td align="center">
    <input type="hidden" name="sSYID" value="<%=sSYID%>">
    <input type="hidden" name="sEventSelected" value="<%=sEventSelected%>">
    <input type="hidden" name="sDivSelected" value="<%=sDivSelected%>">
    <input type="hidden" name="sLeagueSelected" value="<%=sLeagueSelected%>">
    <input type="hidden" name="sSortField" value="<%=sSortField%>">

    <input type="submit" style="width:9em" value="Save">
</td>
<td align="center">
	<input type="reset" style="width:9em" value="Reset">
</td>
</tr>
</TABLE>

</FORM>


<%


rs.close
set rs = nothing

END SUB



' -------------------------------
  SUB SaveRec
' -------------------------------

'response.write("<br>sLeagueSelected="&sLeagueSelected)
'response.write("<br>sDivSelected="&sDivSelected)
'response.write("<br>sEventSelected="&sEventSelected)
'response.write("<br>HERE")
'response.end


'Save the record to the table'


ChooseSQL("SELECT * FROM "&LeagueQfyTableName&" WHERE LeagueID='"&sLeagueSelected&"' AND Event='"&sDivSelected&"' AND Event='"&sDivSelected&"'")

	sSQL = "UPDATE "&LeagueQfyTableName&" SET "

	' --- Ignores 1st field since it assumes this is the KEY ---
  	FOR i = 1 TO rs.fields.count - 1

		response.write("<br>Name ("&i&") = "&Request.Form(rs.fields(i).name))

		IF Request.Form(rs.fields(i).name) <> "" THEN

			IF RIGHT(sSQL,1) <> "," and RIGHT(sSQL,1) <> " " THEN sSQL = sSQL + ","

			sSQL = sSQL + rs.fields(i).name
			sSQL = sSQL + "='" + sqlclean(Request.Form(rs.fields(i).name)) + "'"

		END IF

	NEXT      

rs.close
set rs = nothing

sSQL = sSQL + " WHERE LeagueID='"&sLeagueSelected&"' AND Event='"&sEventSelected&"' AND Div='"&sDivSelected&"'"

'response.write("<br>"&sSQL)
'response.end

OpenCon
con.execute(sSQL)
'WriteLog(date() &"  "& time() &"   "&KeyFieldName&" Record Updated - "& sSQL)
CloseCon

%>
<center><font  color="red" FACE="<%=font1%>" size="<%=fontsize3%>"><b><i>Your updated record has been saved.</I></b></font></center>
<BR>

<form action="<%=ThisFileName%>?action=listrec">
  <center><input type=submit value="Click here to Continue" method="post"></center>
      <input type="hidden" name="sSYID" value="<%=sSYID%>">
      <input type="hidden" name="sEventSelected" value="<%=sEventSelected%>">
      <input type="hidden" name="sDivSelected" value="<%=sDivSelected%>">
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

	sSQL = "INSERT INTO "&ThisTempTableName&" (LeagueID, Event, Div) VALUES ('"&sLeagueSelected&"', '"&sEventSelected&"', '"&sDivSelected&"')"
	OpenCon

	con.execute(sSQL)
	Closecon
	'WriteLog(date() &"  "& time() &"   New Event/Div Has Been Added to LeagueQualify - "& sSQL)

END SUB



'-------------------
  SUB DeleteRec
'-------------------

ChooseSQL("SELECT * FROM "&LeagueQfyTableName&" WHERE LeagueID='"&sLeagueSelected&"' AND Event='"&sDivSelected&"' AND Event='"&sDivSelected&"'")


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
    <center><font  color="#FFFFFF" FACE="<%=font1%>" size="2"><I><b>The record has been deleted.</b></I></font></center>
    <BR>&nbsp;<BR>
    <%

END IF


WriteButton "?action=listrec","Return To League List","<BR><BR>"

IF LCASE(Request("confirm")) = "" THEN
%>  <br><br>
    <center>
     <font color="#FFFFFF" FACE="<%=font1%>" size="2">
    Type the word "YES" IF you are sure you wish to delete this record. </font>
    <br>
    <font color="red" FACE="<%=font1%>" size="2">
    Note: Qualifications from this League may be affected.
    </font>
    <br><br>
    <form action="<%=ThisFileName%>" method="post"> 
      <input type="hidden" name="action" value="delrec">
      <input type="hidden" name="sEventSelected" value="<%=sEventSelected%>">
      <input type="hidden" name="sDivSelected" value="<%=sDivSelected%>">
      <input type="hidden" name="sLeagueSelected" value="<%=sLeagueSelected%>">
      <input type="hidden" name="sSYID" value="<%=sSYID%>">
      <input type="text" name="confirm" size="4" MAXLENGTH="3">
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
				%><input type="text" name="<% Response.Write(Rs.Fields(i).name) %>" size="25" value="<% GetFieldValue i %>"><%
			CASE 20 'primary key / auto number ?'
				%><input type="text" name="<% Response.Write(Rs.Fields(i).name) %>" size="25" value="<% GetFieldValue i %>"><%
			CASE 11 'boolean'
        			%><INPUT TYPE="checkbox" NAME="<% Response.Write(Rs.Fields(i).name) %>" VALUE="0"<% GetcheckValue i %>><%
			CASE 129 'char'
				%><input type="text" name="<% Response.Write(Rs.Fields(i).name) %>" size="8" value="<% GetFieldValue i %>"><%
			CASE 203 'memo'
        			%><TEXTAREA NAME="<% Response.Write(Rs.Fields(i).name) %>" ROWS="20" COLS="56"><% GetFieldValue i %></TEXTAREA><%
			CASE ELSE 'not handled by this function'
			        %><input type="text" name="<% Response.Write(Rs.Fields(i).name) %>" size="25" value="<% GetFieldValue i %>"><%
		END SELECT

END SELECT 

END SUB



' ---------------------
  SUB BuildAWSAEvents
' ---------------------

%>
<select name="sEventSelected" style="width:12em">
  <option value ="ALL" <%IF sEventSelected = "ALL" THEN Response.Write(" selected ")%> >All Events</Option><br>
  <option value ="S" <%IF sEventSelected = "S" THEN Response.Write(" selected ")%> >Slalom</Option><br>
  <option value ="T" <%IF sEventSelected = "T" THEN Response.Write(" selected ")%> >Tricks</Option><br>
  <option value ="J" <%IF sEventSelected = "J" THEN Response.Write(" selected ")%> >Jump</Option><br>
  <option value ="O" <%IF sEventSelected = "O" THEN Response.Write(" selected ")%> >Overall</Option><br>
</select><%



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

<BODY BGcolor="#FFFFFE" Text="#0A0D0A" LINK="#375AE2" VLINK="#36566D" ALINK="#3E85BB">
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
  <input type="hidden" name="sEventSelected" value="<%=sEventSelected%>">
  <input type="hidden" name="sDivSelected" value="<%=sDivSelected%>">
  <input type="hidden" name="sLeagueSelected" value="<%=sLeagueSelected%>">
  <input type="hidden" name="sSYID" value="<%=sSYID%>">
  <input type="hidden" name="sSortField" value="<%=sSortField%>">
</form>
<%

END SUB



' ---------------------------------------
    SUB DoCount(currentPage) 
' ---------------------------------------

'response.write("<br>sLeagueSelected="&sLeagueSelected)
'response.write("<br>sSortField="&sSortField)
'response.write("<br>action="&sAction)
'response.write("<br>sEventSelected="&sEventSelected)
'response.write("<br>TRIM(rs(Event)="&rs("Event"))
'response.write("<br>sDivSelected="&sDivSelected)
'response.write("<br>rs(Div)="&rs("Div"))
'response.write("<br>sSYID="&sSYID)
'response.end


h = rs.PageCount

IF h > 21 THEN
  IF currentpage - 10 > 1 THEN
    	Response.Write("... ")
  END IF


  FOR i = 10 TO 1 step -1
    IF currentpage - i > 0 THEN
      	Response.Write(" <a href=" & chr(34) & ThisPage & "?currentpage=" &  currentpage - i  & "&action1=" & sAction & chr(34) & "&sSortField="&sSortField&"&sLeagueSelected="&sLeagueSelected&"&sEventSelected="&TRIM(rs("Event"))&"&sDivSelected="&rs("Div")&"&sSYID="&sSYID&">" & currentpage - i & "</a>")
    END IF
  NEXT

  Response.Write(" " & currentpage & " ")

  FOR i = 1 TO 10
   	IF currentpage + i <= h THEN
      		Response.Write(" <a href=" & chr(34) & ThisPage & "?currentpage=" &  currentpage + i  & "&action2=" & sAction & "&sSortField="&sSortField&"&sLeagueSelected="&sLeagueSelected&"&sEventSelected="&TRIM(rs("Event"))&"&sDivSelected="&rs("Div")&"&sSYID="&sSYID& chr(34) &">" & currentpage + i & "</a>")
	END IF
  NEXT

  IF currentpage + 10 <= h THEN
    Response.Write(" ...")
  END IF

ELSE



  FOR i = 1 TO h
    Response.Write(" <a href=" & chr(34) & ThisPage & "?currentpage=" &i& "&action=" &sAction& "&sSortField="&sSortField&"&sLeagueSelected="&sLeagueSelected&"&sEventSelected="&sEventSelected&"&sDivSelected="&sDivSelected&"&sSYID="&sSYID& chr(34) & ">" &i& "</a>")
  next

END IF

IF h = 0 THEN h = 1 

Response.Write("<BR><font color=#FFFFFF FACE=font1 size=0>Page " & currentPage & " of  "&h&"</font></center><BR><BR>")


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