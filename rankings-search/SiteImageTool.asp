<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->
<!--#include virtual="/rankings/tools_registration.asp"-->
<%


Dim sTSiteID, sSiteID, sNewHeaderImage, sOldHeaderImage
Dim ThisFileName, sTSiteName, sTSiteCity, sTSiteState
Dim AllButtonStatus

DefineTRAStyles

WriteIndexPageHeader  


ThisFileName="SiteImageTool.asp"

sTourID=LEFT(TRIM(Request("sTourID")),6)

AllButtonStatus="enabled"
IF sTourID="" THEN 
	AllButtonStatus="disabled"
	%><center><font size=3 color=red>Select a Tournament to Define the SiteID</font></center><%
END IF



sRunByWhat = TRIM(LCASE(Request("pvar")))

SELECT CASE sRunByWhat

   CASE "save new"
	ReadFormVariables
	AddTheImage
	ReadTableValues
	DisplaySiteIDPage

   CASE "update"
	ReadFormVariables
	UpdateTheImage
	ReadTableValues
	DisplaySiteIDPage

   CASE "siteid"
	FindsiteID
	ReadTableValues
	DisplaySiteIDPage
	
   CASE ELSE
	ReadTableValues
	DisplaySiteIDPage
END SELECT




' -------------------------
  SUB DisplaySiteIDPage
' -------------------------

'response.write("<br>Pos 4")

'response.write("<br> sTourName = "&sTourName)
'response.end



	%>
	<br><br>
    <TABLE class="innertable" Align="center" width=70%>
        <form name="ImageForm" method="post" action="/rankings/<%=ThisFileName%>" >
	<input type="hidden" name="sTSiteID" value="<%=sTSiteID%>">
	<input type="hidden" name="sOldHeaderImage" value="<%=sOldHeaderImage%>">
	<TR>
	  <TH align=center colspan=4 bgcolor="#2F4F4F"><font size="4" color="#FFFFFF"><b>Site Image Update</b></font></TD> 
	</TR>

	<TR>
	  <TD align=right width=33%><font size="<%=fontsize2%>" color="<%=TextColor1%>">TourID&nbsp;</font></TD> 
	  <TD align=left colspan=2 width=67%><%
		LoadTourList %>
	  </TD> 
	</TR>

	<TR>
	  <TD align=right><font size="<%=fontsize2%>" color="<%=TextColor1%>">Tournament&nbsp;</font></TD> 
	  <TD align=left colspan=2 width=67%><font size="<%=fontsize2%>" color="<%=TextColor2%>">&nbsp;<%=sTourName%></font></TD> 
	</TR>

	<TR>
	  <TD align=right><font size="<%=fontsize2%>" color="<%=TextColor1%>">&nbsp;</font></TD> 
	  <TD align=left colspan=2 width=67%><font size="<%=fontsize2%>" color="<%=TextColor2%>">&nbsp;</font></TD> 
	</TR>

	<TR>
	  <TD align=right><font size="<%=fontsize2%>" color="<%=TextColor1%>">IWSF SiteID&nbsp;</font></TD> 
	  <TD align=left colspan=2 width=67%><font size="<%=fontsize2%>" color="<%=TextColor2%>">&nbsp;<%=sTSiteID%></font></TD> 
	</TR>

	<TR>
	  <TD align=right><font size="<%=fontsize2%>" color="<%=TextColor1%>">Site Name&nbsp;</font></TD> 
	  <TD align=left colspan=2 width=67%><font size="<%=fontsize2%>" color="<%=TextColor2%>">&nbsp;<%=sTSiteName%></font></TD> 
	</TR>

	<TR>
	  <TD align=right><font size="<%=fontsize2%>" color="<%=TextColor1%>">Site Location&nbsp;</font></TD> 
	  <TD align=left colspan=2 width=67%><font size="<%=fontsize2%>" color="<%=TextColor2%>">&nbsp;<%=sTSiteCity%>, <%=sTSiteState%></font></TD> 
	</TR>

	<TR>
	  <TD align=right><font size="<%=fontsize2%>" color="<%=TextColor1%>">&nbsp;</font></TD> 
	  <TD align=left colspan=2 width=67%><font size="<%=fontsize2%>" color="<%=TextColor2%>">&nbsp;</font></TD> 
	</TR>

	<TR>
	  <TD align=right><font size="<%=fontsize2%>" color="<%=TextColor1%>"> Current Image&nbsp;</font></TD> 
	  <TD align=left colspan=2><font size="<%=fontsize2%>" color="<%=TextColor2%>">&nbsp;<%=sOldHeaderImage%></font></TD> 
	</TR>
	<TR>
	  <TD align=right><font size="<%=fontsize2%>" color="<%=TextColor1%>"> New Image&nbsp;</font></TD> 
	  <TD align=left colspan=2><input type="text" name="sNewHeaderImage" maxlength=35 size=38></TD> 
	</TR>

	<TR>
	  <td align="center">
		<br><%
		IF sOldHeaderImage="NF" THEN %>
			<input type="submit" name="pvar" value="Save New" style="width:9em" title="Save the current settings to the table" <%=AllButtonStatus%>><%
		ELSE %>
			<input type="submit" name="pvar" value="Update" style="width:9em" title="Update the current settings to the table" <%=AllButtonStatus%>><%
		END IF %>
		<br>
	  </td>
	</form>
        <form name="ContDisp2" method="post" action="/rankings/DefaultHQ.asp" id="ContDisp2">
	  <td align="center">
		<br>
		<input type="submit" name="pvar" value="Main Menu" style="width:9em" title="Return to Main Menu">
		<br>
	  </td>
	</form>

        <form name="ContDisp2" method="post" action="/rankings/<%=ThisFileName%>" id="ContDisp2">
	  <td align="center">
		<br>
		<input type="submit" name="pvar" value="Upload" style="width:9em" title="Upload a New Site Image" disabled>
		<br>
	  </td>
	</form>

	</TR>

   </TABLE>
<br>
<center><font size="<%=fontsize3%>" color="<%=TextColor2%>"><br>If 'New Image' field is blank, nothing will be Saved.  <br> The 'Current Image' will be retained if being Updated.</b></font></center>
<br><br>

<%

'response.write("<br>Pos 5")
'response.end


WriteIndexPageFooter


END SUB



' ------------------
  SUB FindSiteID
' ------------------

	Session("sSendingPage")="/rankings/SiteImageTool.asp"
	Response.Redirect("/rankings/search-memberHQ.asp?rid="&rid&"&formstatus=search")

END SUB



' -----------------------
  SUB ReadFormVariables
' -----------------------

	sTSiteID=TRIM(Request("sTSiteID"))
	sOldHeaderImage=TRIM(Request("sOldHeaderImage"))
	sNewHeaderImage=TRIM(Request("sNewHeaderImage"))

'response.write("<br>IN READ")

'response.write("<br>sTSiteID = "&sTSiteID)
'response.write("<br>sOldHeaderImage = "&sOldHeaderImage)
'response.write("<br>sNewHeaderImage = "&sNewHeaderImage)


END SUB



' ----------------------------------
  SUB ReadTableValues
' ----------------------------------

sSiteID=TRIM(Session("sSiteID"))


' --- Read transactions from Credit Card Table to determine Total Fees actually completed ----
SET rs=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT ST.TournAppID, TName, TE.SiteID, WSS.TCity AS TSiteCity, WSS.TState AS TSiteState,"
sSQL = sSQL + " WSS.TSite, ST.TSiteID, coalesce(HeaderImage,'None Loaded') AS OldHeaderImage"

sSQL = sSQL + " FROM "&sanctionTableName&" AS ST"
sSQL = sSQL + " LEFT JOIN usawsrank.TourExtras AS TE"
sSQL = sSQL + " 	ON SiteID=ST.TSiteID" 

sSQL = sSQL + " LEFT JOIN sanctions.dbo.WSSites AS WSS"
sSQL = sSQL + " 	ON WSS.TSiteID=ST.TSiteID"

sSQL = sSQL + " WHERE ST.TournAppID='"&sTourID&"'" 


'response.write(sSQL)
'response.end

rs.open sSQL, SConnectionToTRATable, 3, 3

IF NOT rs.eof THEN
	sOldSiteID= TRIM(rs("SiteID"))
	sTourName = TRIM(rs("TName"))
	sTSiteName = TRIM(rs("TSite"))
	sTSiteID = TRIM(rs("TSiteID"))
	sTSiteCity = TRIM(rs("TSiteCity"))
	sTSiteState = TRIM(rs("TSiteState"))

	sSiteID = TRIM(rs("SiteID"))
	sOldHeaderImage = TRIM(rs("OldHeaderImage"))
END IF

'response.write("<br> sTName = "&sTourName)

END SUB



' -----------------
  SUB LoadTourList
' -----------------

' --- Read transactions from Credit Card Table to determine Total Fees actually completed ----
SET rs=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT ST.TournAppID FROM "&sanctionTableName&" AS ST"

sSQL = sSQL + " LEFT JOIN "&TRegSetupTableName&" AS TRS"
sSQL = sSQL + " ON TRS.TournAppID=ST.TournAppID"

sSQL = sSQL + " WHERE TSTATUS > 1 AND PayPalOK=1"
sSQL = sSQL + " AND LEFT(ST.TournAppID,2) IN (SELECT RIGHT(SkiYear,2) FROM usawsrank.SkiYear WHERE EndDate>=GetDate())"

sSQL = sSQL + " ORDER BY ST.TournAppID"


rs.open sSQL, SConnectionToTRATable, 3, 3


%><select name='sTourID' style="width:12em" onchange=submit()><%

IF NOT rs.eof THEN 
  	rs.movefirst

  	DO WHILE NOT rs.eof %>
		<option value = "<%=rs("TournAppID")%>" <%IF rs("TournAppID") = sTourID THEN response.write(" selected ")%>><%=rs("TournAppID")%></option><%
		rs.movenext
	LOOP
ELSE
	response.write("<option value =""None"" selected>None Available</option>")
END IF  %>
</select><%


END SUB




' -------------------
  SUB UpdateTheImage
' -------------------

'response.write("<br>IN UPDATE")

'response.write("<br>sTSiteID = "&sTSiteID)
'response.write("<br>sOldHeaderImage = "&sOldHeaderImage)
'response.write("<br>sNewHeaderImage = "&sNewHeaderImage)

'response.end

IF sNewHeaderImage<> sOldHeaderImage AND sNewHeaderImage<>"" THEN
	OpenCon
	sSQL = "UPDATE usawsrank.TourExtras"
	sSQL = sSQL + " SET HeaderImage ='"&sNewHeaderImage&"'"
	sSQL = sSQL + " WHERE SiteID='"&sTSiteID&"'"

'response.write("<br>"&sSQL)
'response.end

	con.execute(sSQL)
	closecon
END IF

END SUB



' -------------------
  SUB AddTheImage
' -------------------


'response.write("<br>IN ADD")
'response.write("<br>sTSiteID = "&sTSiteID)
'response.write("<br>sOldHeaderImage = "&sOldHeaderImage)
'response.write("<br>sNewHeaderImage = "&sNewHeaderImage)

IF sNewHeaderImage<>"" THEN
	OpenCon
	sSQL = "INSERT INTO usawsrank.TourExtras"
	sSQL = sSQL + " (HeaderImage, SiteID) VALUES ('"&sNewHeaderImage&"', '"&sTSiteID&"')"

'response.write("<br>"&sSQL)
'response.end

	con.execute(sSQL)
	closecon
END IF

END SUB




%>