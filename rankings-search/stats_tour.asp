<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/Tools_Definitions.asp"-->
<!--#include virtual="/rankings/Tools_include.asp"-->
<!--#include virtual="/rankings/tools_registration.asp"-->

<%

ThisFileName = "Stats_Tour.asp"

Dim sSkiYearID, sSortBy, sIncludePast, sIncludeFuture
Dim sTourID, sShowAdminCode, Action
Dim ThisTableWidth, AdminMenuLevel


AdminMenuLevel=Session("AdminMenuLevel")

sShowAdminCode=TRIM(Request("sShowAdminCode"))

sSkiYearID = TRIM(Request("sSkiYearID"))
IF TRIM(sSkiYearID) = "" THEN 
	SET rs=Server.CreateObject("ADODB.recordset")
	sSQL = "SELECT TOP 1 SkiYearID"
	sSQL = sSQL + " FROM "&SkiYearTableName&" AS SY"
	sSQL = sSQL + " WHERE DefaultYear=1"
	sSQL = sSQL + " ORDER BY SkiYearID DESC"
	rs.open sSQL, SConnectionToTRATable
	sSkiYearID = rs("SkiYearID")
END IF
'IF sSkiYearID="" THEN sSkiYearID=14

sSortBy = TRIM(Request("SortBy"))
IF sSortBy = "" THEN sSortBy="1"

sIncludePast = Request("sIncludePast")
'response.write(sIncludePast)

sIncludeFuture = Request("sIncludeFuture")

pvar = TRIM(Request("pvar"))
Action=Request("Action")
IF pvar<>"" THEN Action=pvar

IF sIncludeFuture <> "on" AND sIncludePast <> "on" THEN sIncludeFuture= "on"


ThisTableWidth=900


'WriteIndexPageHeader
WriteIndexPageHeader_NoMenu

'response.write("<br><br>action="&action)


SELECT CASE LCASE(Action)
	CASE "done"
			'response.write("<br>HERE")
			'response.end
			response.redirect("/rankings/DefaultHQ.asp")
	
	CASE "entrystat", "submit"
		' --- Runs OLR Entries By Tournament report
		DisplayTourReg_Header
		RunTourTotalSQL

		IF NOT rs.eof THEN 

			DisplayTourRegistrations
		END IF

	CASE "fulltest"
		RunFullTest

END SELECT

WriteIndexPageFooter


' --------------------------------------------------------------------------------------------------------------







' ------------------------------
  SUB DisplayTourReg_Header
' ------------------------------

Dim ReportTitle
ReportTitle = "OLR Entries By Tournament"

' 

' --- Draws Page Header ---
%>
  <form action="/rankings/<%=ThisFileName%>" method="post">
<TABLE class="droptable" height="140px" align=center width="<%=ThisTableWidth%>" >
  <input type="hidden" name="WhatReport" value="<%=WhatReport%>">
  <tr>
	  <td colspan=2 align="left" width="50%">
		  <FONT color="<%=TextColor2%>" size=3><B><% Response.Write(ReportTitle) %></B></FONT>
			<br>
	  </td>
    <td colspan=2>&nbsp;</td>
  </tr>
  <tr>
		<td align="right" width="20%">
				<font color="<%=textcolor2%>"> Ski Year:</font>
		</td>	
		<td align="left">
		<%  

		' --- SUB in Tools_Include.asp ---
		CreateSkiYearDropDown sSkiYearID  

		%>
		</td>		
    <td align=right width="20%"><font color="<%=textcolor2%>"> Include Completed</font></td>	
    <td align=left>
			<input type="checkbox" name="sIncludePast" <% IF sIncludePast <> "" THEN Response.Write("Checked") %>>
    </td>	
  </tr>	
  <tr>
    <td align="right"><font color="<%=textcolor2%>">Sort By</td>	
    <td >
			<select name='SortBy'>
				<option value ='1'<%IF sSortBy = "1" THEN Response.Write(" selected ")%>>Sanction/Region</Option><br>
				<option value ='2'<%IF sSortBy = "2" THEN Response.Write(" selected ")%>>Date</Option><br>
				<option value ='3'<%IF sSortBy = "3" THEN Response.Write(" selected ")%>>Tour Name</Option><br>
			</select>
    </td><%	

    IF AdminMenuLevel>=50 THEN  %>
        <td align=right><font color="<%=textcolor2%>"> Include AdminCode</font></td>
	    	<td align=left>
					<input type="checkbox" name="sShowAdminCode" <% IF sShowAdminCode <> "" THEN Response.Write("Checked") %>>
	    	</td><%	
		ELSE %>
	    	<td colspan=2>&nbsp;</td><%	
		END IF 
		%>
  </tr>	
  <tr>
    <td align=center colspan=2>
      <input type="submit" name="Action" value="Submit" style="width:9em">
    </td>
    <td align=center colspan=2>
			<input type="submit" name="Action" value="Done" style="width:9em">
		</td>
  </tr>		
</TABLE>
<br>
</form> 

<%



END SUB



' ------------------------------
  SUB DisplayTourRegistrations
' ------------------------------

'Response.write("sIncludePast = "&sIncludePast)
'Response.write("sIncludeFuture = "&sIncludeFuture)

%>

<TABLE Align="center" class="innertable" width="<%=ThisTableWidth%>">
	  <TR>
		<th align=center><font size="<%=fontsize2%>" color="#FFFFFF">Tour ID</font></th>
		<th align=left><font size="<%=fontsize2%>" color="#FFFFFF">Tournament Name</font></th>
		<th align=center><font size="<%=fontsize2%>" color="#FFFFFF">Date</font></th><%
		IF sShowAdminCode="on" THEN %>
			<th align=center><font size="<%=fontsize2%>" color="#FFFFFF">Admin Code</font></th><%
		END IF %>
		<th align=center><font size="<%=fontsize2%>" color="#FFFFFF">Entries</font></th>
		<th align=center><font size="<%=fontsize2%>" color="#FFFFFF">Tot<br>Pulls</font></th>
		<th align=center><font size="<%=fontsize2%>" color="#FFFFFF">Max<br>Pulls</font></th>
		<th align=center><font size="<%=fontsize2%>" color="#FFFFFF">Res<br>Pulls</font></th>
		<th align=center><font size="<%=fontsize2%>" color="#FFFFFF">PayPal<br>Account</font></th>
		<th align=center><font size="<%=fontsize2%>" color="#FFFFFF">Tour<br>Status</font></th>
		<th align=center><font size="<%=fontsize2%>" color="#FFFFFF">OLR_Pd<br>Status</font></th>
		<th align=center><font size="<%=fontsize2%>" color="#FFFFFF">Grass</font></th>
		<th align=center><font size="<%=fontsize2%>" color="#FFFFFF">OLR<br>Display<br>Status</font></th>

	  </TR><%

	' --- Loops through query ---
	DO WHILE Not rs.EOF 
		TourCount=TourCount+1

		sTourID=rs("TourID")
		sTSEmail=rs("TSEmail")
		sTName=rs("TName")
		sAdminCode=""
		IF sShowAdminCode="on" THEN
			sAdminCode=rs("AdminCode")
		END IF
		sEntries=rs("Entries") 
		sTFeeRounds=rs("FeeRounds")
		sTMaxPulls=rs("MaxPulls")
		IF sTMaxPulls<5 AND sTMaxPulls<>0 THEN sTMaxPullColor="Red" ELSE sTMaxPullColor=""
		sTReservedPulls=rs("ReservedPulls")
		IF rs("PayPalOK")=true THEN sPayPalOK="Yes" ELSE sPayPalOK="No"

		IF rs("OLR_pd") = true THEN 
			sOLR_Pd="Yes" 
			OLR_PdColor=""
		ELSE 
			sOLR_Pd = "No"
			OLR_PdColor="red"
		END IF
	
		IF rs("Grassroots") THEN sGrassroots="Yes" ELSE sGrassroots=""


		IF rs("TStatus")>1 THEN AppvdCount=AppvdCount+1  
		SELECT CASE rs("TStatus")
			CASE 0 
				sTStatus="Pend"
				sTStatusColor="red"
			CASE 1 
				sTStatus="RegApv"
				sTStatusColor="yellow"
			CASE 2
				sTStatus="Appvd"
				sTStatusColor=""
		END SELECT

		LineColor=TableColor1
		IF sTFeeRounds > sTMaxPulls AND sTMaxPulls <> 0 THEN 
			LineColor = "red"
		ELSEIF sTFeeRounds >= 0.75 * sTMaxPulls AND sTMaxPulls <> 0 THEN 
			LineColor = "yellow"
		END IF
	  %>	
	  <TR >

		<td align=center ><font size="<%=fontsize2%>"><%=sTourID%></font></td>
		<td align=left>
			<font size="<%=fontsize2%>" >
			<a title="Send Email to <% =sTSEmail%>" href="mailto:<%=sTSEmail%>?subject=Online Registration - <%=LEFT(sTName,25)%>"><%=rs("TName")%>&nbsp;</a>
			</font>
		</td>
		<td align=center><font size="<%=fontsize2%>"><%=rs("Date")%></font></td><%
		IF sShowAdminCode="on" THEN %>
			<td align=center><font size="<%=fontsize2%>"><%=sAdminCode%></font></td><%
		END IF %>
		<td align=center style="background:<%=LineColor%>"><font size="<%=fontsize2%>"><%=sEntries%></font></td>
		<td align=center style="background:<%=LineColor%>"><font size="<%=fontsize2%>"><%=sTFeeRounds%></font></td>
		<td align=center style="background:<%=sTMaxPullColor%>"><font size="<%=fontsize2%>"><%=sTMaxPulls%></font></td>

		<td align=center style="background:<%=LineColor%>"><font size="<%=fontsize2%>"><%=sTReservedPulls%></font></td>
		<td align=center><font size="<%=fontsize2%>"><%=sPayPalOK%></font></td>
		<td align=center style="background:<%=sTStatusColor%>"><font size="<%=fontsize2%>"><%=sTStatus%></font></td>
		<td align=center style="background:<%=OLR_PdColor%>"><font size="<%=fontsize2%>"><%=sOLR_Pd%></font></td>
		<td align=center><font size="<%=fontsize2%>"><%=sGrassroots%></font></td><% 

		OLRColor=TableColor1
		OLRStat="ON"
		IF rs("OLRDisplayStatus")=False THEN 
			OLRColor="red"
			OLRStat="OFF" 
		END IF%>

		<td align=center style="background:<%=OLRColor%>"><font size="<%=fontsize2%>"><%=OLRStat%></font></td>
	  </TR><%
   	  rs.movenext
	LOOP %>

</TABLE>
<br>
<font size=<% =fontsize1 %>>Total Count = <% =TourCount %></FONT>
<br>
<font size=<% =fontsize1 %>>Total Approved = <% =AppvdCount %></FONT> 
<%


END SUB 





' --------------------------
   SUB RunTourTotalSQL
' --------------------------


' --- Determines the LATEST SkiYearID if one has not been selected ---
IF TRIM(sSkiYearID) = "" THEN 
	SET rs=Server.CreateObject("ADODB.recordset")
	sSQL = "SELECT TOP 1 SkiYearID"
	sSQL = sSQL + " FROM "&SkiYearTableName&" AS SY"
	sSQL = sSQL + " WHERE DefaultYear=1"
	sSQL = sSQL + " ORDER BY SkiYearID DESC"
	rs.open sSQL, SConnectionToTRATable
	sSkiYearID = rs("SkiYearID")
END IF

' --- Sets query for OLR Tournament display ---
sSQL =  " SELECT ST.TournAppID AS TourID, ST.TName, CONVERT(Varchar, TDateE,110) AS Date, EmailAddress as RegistrarEmail"
sSQL = sSQL + "		, RT.MaxPulls, RT.ReservedPulls, PayPalOK, TStatus, Grassroots, OLRDisplayStatus, TSEmail, OLR_Pd, ST.UseOLReg AS OLR" 
sSQL = sSQL + "		, Entries, FeeRounds, ST.UseOLReg"
IF sShowAdminCode="on" THEN
	sSQL = sSQL + ", RT.AdminCode"
END IF

sSQL = sSQL + "		FROM Sanctions.dbo.TSchedul AS ST"
sSQL = sSQL + "				LEFT JOIN usawsrank.SkiYear AS SY" 
sSQL = sSQL + "				ON LEFT(ST.TournAppID,2)=RIGHT(SY.SkiYear,2)" 

sSQL = sSQL + "				JOIN Sanctions.dbo.registration AS RT" 
sSQL = sSQL + "				ON LEFT(RT.TournAppID,6)=LEFT(ST.TournAppID,6)" 

sSQL = sSQL + "				LEFT JOIN" 
sSQL = sSQL + "				( SELECT TourID, Coalesce(COUNT(DISTINCT MemberID),0) AS Entries"
sSQL = sSQL + "					FROM usawsrank.RegisterGenNew"
sSQL = sSQL + "					GROUP BY TourID) AS RG" 
sSQL = sSQL + "				ON LEFT(ST.TournAppID,6)=LEFT(RG.TourID,6)" 

sSQL = sSQL + "				LEFT JOIN" 
sSQL = sSQL + "				( SELECT TourID, Coalesce(SUM(FeeRounds),0) AS FeeRounds"
sSQL = sSQL + "						FROM usawsrank.RegisterEvents"
sSQL = sSQL + "						GROUP BY TourID) RE" 
sSQL = sSQL + "				ON LEFT(RE.TourID,6)=LEFT(ST.TournAppID,6)"

sSQL = sSQL + " WHERE 1=1 "
sSQL = sSQL + " AND ST.UseOLReg=1" 

IF sIncludePast="on" THEN
		sSQL = sSQL + " AND SY.SkiYearID='"&sSkiYearID&"'"
ELSE
		sSQL = sSQL + " AND SY.SkiYearID='"&sSkiYearID&"' AND ST.TDateS>'"&Date&"'"
END IF
	
	
' --- Required to make "Sort By" drop-down work properly --- 
SELECT CASE sSortBy
	CASE "1"
		sSQL = sSQL + " ORDER BY RIGHT(LEFT(ST.TournAppID,3),1), ST.TDateS"
	CASE "2"
		sSQL = sSQL + " ORDER BY ST.TDateS, RIGHT(LEFT(ST.TournAppID,3),1)"
	CASE "3"
		sSQL = sSQL + " ORDER BY ST.TName, ST.TDateS"
END SELECT


'response.write("<br>"&sSQL)
'response.end


SET rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable, 3, 3




END SUB


' ----------------
  SUB RunFullTest
' ----------------

	sTourID="10M021"
	'response.write("<br>IN RUNFULLTEST")
	response.write("<br>Answer to function = "&EntriesExceedLimit(sTourID))


END SUB

 












%>