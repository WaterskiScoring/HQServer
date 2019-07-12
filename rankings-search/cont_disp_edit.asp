<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"--><%



Dim sEntryEmail, sWaiverEmail, sPasswordEmail, sSkipWaiver, sForceWaiver
Dim sEntryEmailAdm, sWaiverEmailAdm, sPasswordEmailAdm, sSkipWaiverAdm, sForceWaiverAdm
Dim sEntryEmailHQ, sWaiverEmailHQ, sPasswordEmailHQ, sSkipWaiverHQ, sForceWaiverHQ
Dim sDispDebugButtons, sDispDebugButtonsAdm, sDispDebugButtonsHQ



sRunByWhat = TRIM(Request("pvar"))

SELECT CASE sRunByWhat

   CASE "Save"
	ReadFormVariables
	StoreTheValues
	ReadControlDisplayTableValues
	DisplaySettingsPage

   CASE ELSE
	ReadControlDisplayTableValues
	DisplaySettingsPage
END SELECT




' -------------------------
  SUB DisplaySettingsPage
' -------------------------


	DefineTRAStyles

	WriteIndexPageHeader  %>

	<br><br><br><br>
    <TABLE class="innertable"  Align="center">
        <form name="ContDisp1" method="post" action="/rankings/Cont_Disp_Edit.asp" id="ContDisp1">
	<TR>
	  <TH align=center colspan=4 bgcolor="#2F4F4F"><font size="4" color="#FFFFFF"><b>DISPLAY CONTROL SETTINGS</b></font></TD> 
	</TR>

	<TR>
	  <TD align=center width=150>&nbsp;</td>
	  <TD align=center width=100><font size="<%=fontsize2%>" color="<%=TextColor1%>">Member</font></TD> 
	  <TD align=center width=100><font size="<%=fontsize2%>" color="<%=TextColor1%>">Admin</font></TD> 
	  <TD align=center width=100><font size="<%=fontsize2%>" color="<%=TextColor1%>">HQ</font></TD> 
	</TR>

	<TR>
	  <TD align=right><font size="<%=fontsize2%>" color="<%=TextColor1%>"> Send Entry Email&nbsp;</font></TD> 
	  <TD align=center><input type=checkbox name="sEntryEmail" <%IF sEntryEmail = True THEN Response.Write("checked")%>></TD>
	  <TD align=center><input type=checkbox name="sEntryEmailAdm" <%IF sEntryEmailAdm = True THEN Response.Write("checked")%>></TD>
	  <TD align=center><input type=checkbox name="sEntryEmailHQ" <%IF sEntryEmailHQ = True THEN Response.Write("checked")%>></TD>
	</TR>
	<TR>
	  <TD align=right><font size="<%=fontsize2%>" color="<%=TextColor1%>"> Send Waiver Email&nbsp;</font></TD> 
	  <TD align=center><input type=checkbox name="sWaiverEmail" <%IF sWaiverEmail = True THEN Response.Write("checked")%>></TD></TD>
	  <TD align=center><input type=checkbox name="sWaiverEmailAdm" disabled <%IF sWaiverEmailAdm = True THEN Response.Write("checked")%>></TD></TD>
	  <TD align=center><input type=checkbox name="sWaiverEmailHQ" <%IF sWaiverEmailHQ = True THEN Response.Write("checked")%>></TD></TD>
	</TR>
	<TR>
	  <TD align=right><font size="<%=fontsize2%>" color="<%=TextColor1%>">Send Password Email&nbsp;</font></TD> 
	  <TD align=center><input type=checkbox name="sPasswordEmail" <%IF sPasswordEmail = True THEN Response.Write("checked")%>></TD>
	  <TD align=center><input type=checkbox name="sPasswordEmailAdm" <%IF sPasswordEmailAdm = True THEN Response.Write("checked")%>></TD>
	  <TD align=center><input type=checkbox name="sPasswordEmailHQ" <%IF sPasswordEmailHQ = True THEN Response.Write("checked")%>></TD>
	</TR>
	<TR>
	  <TD align=right><font size="<%=fontsize2%>" color="<%=TextColor1%>">Skip Waiver&nbsp;</font></TD> 
	  <TD align=center><input type=checkbox name="sSkipWaiver" <%IF sSkipWaiver = True THEN Response.Write("checked")%>></TD>
	  <TD align=center><input type=checkbox name="sSkipWaiverAdm" <%IF sSkipWaiverAdm = True THEN Response.Write("checked")%>></TD>
	  <TD align=center><input type=checkbox name="sSkipWaiverHQ" <%IF sSkipWaiverHQ = True THEN Response.Write("checked")%>></TD>
	</TR>
	<TR>
	  <TD align=right><font size="<%=fontsize2%>" color="<%=TextColor1%>">Force Waiver&nbsp;</font></TD> 
	  <TD align=center><input type=checkbox name="sForceWaiver" <%IF sForceWaiver = True THEN Response.Write("checked")%>></TD>
	  <TD align=center><input type=checkbox name="sForceWaiverAdm" <%IF sForceWaiverAdm = True THEN Response.Write("checked")%>></TD>
	  <TD align=center><input type=checkbox name="sForceWaiverHQ" <%IF sForceWaiverHQ = True THEN Response.Write("checked")%>></TD>
	</TR>
	<TR>
	  <TD>&nbsp;</td>
	  <TD>&nbsp;</td>
	  <TD>&nbsp;</td>
	  <TD>&nbsp;</td>
	</TR>
	<TR>
	  <TD>&nbsp;</td>
	  <TD>&nbsp;</td>
	  <TD>&nbsp;</td>
	  <TD>&nbsp;</td>
	</TR>
	<TR>
	  <TD align=right><font size="<%=fontsize2%>" color="<%=TextColor1%>">Display Debug Buttons</font></TD> 
	  <TD align=center><input type=checkbox name="sDispDebugButtons" <%IF sDispDebugButtons = True THEN Response.Write("checked")%>></TD>
	  <TD align=center><input type=checkbox name="sDispDebugButtonsAdm" <%IF sDispDebugButtonsAdm = True THEN Response.Write("checked")%>></TD>
	  <TD align=center><input type=checkbox name="sDispDebugButtonsHQ" <%IF sDispDebugButtonsHQ = True THEN Response.Write("checked")%>></TD>
	</TR>


	<TR>
	  <td align="center" colspan=2>
		<br>
		<input type="submit" name="pvar" value="Save" style="width:9em" title="Save the current settings to the table">
		<br>
	  </td>
	</form>
        <form name="ContDisp2" method="post" action="/rankings/DefaultHQ.asp" id="ContDisp2">
	  <td align="center" colspan=2>
		<br>
		<input type="submit" name="pvar" value="Main Menu" style="width:9em" title="Return to Main Menu">
		<br>
	  </td>
	</form>
	</TR>

   </TABLE>


<%

WriteIndexPageFooter


END SUB


' -----------------------
  SUB ReadFormVariables
' -----------------------


	' --- Values read are convert to store in bit fields of table ---
	IF TRIM(Request("sEntryEmail")) = "on" THEN sEntryEmail=1 ELSE sEntryEmail=0
	IF TRIM(Request("sEntryEmailAdm")) = "on" THEN sEntryEmailAdm=1 ELSE sEntryEmailAdm=0
	IF TRIM(Request("sEntryEmailHQ")) = "on" THEN sEntryEmailHQ=1 ELSE sEntryEmailHQ=0
	IF TRIM(Request("sWaiverEmail")) = "on" THEN sWaiverEmail=1 ELSE sWaiverEmail=0
	IF TRIM(Request("sWaiverEmailAdm")) = "on" THEN sWaiverEmailAdm=1 ELSE sWaiverEmailAdm=0
	IF TRIM(Request("sWaiverEmailHQ")) = "on" THEN sWaiverEmailHQ=1 ELSE sWaiverEmailHQ=0
	IF TRIM(Request("sPasswordEmail")) = "on" THEN sPasswordEmail=1 ELSE sPasswordEmail=0
	IF TRIM(Request("sPasswordEmailAdm")) = "on" THEN sPasswordEmailAdm=1 ELSE sPasswordEmailAdm=0
	IF TRIM(Request("sPasswordEmailHQ")) = "on" THEN sPasswordEmailHQ=1 ELSE sPasswordEmailHQ=0
	IF TRIM(Request("sSkipWaiver")) = "on" THEN sSkipWaiver=1 ELSE sSkipWaiver=0
	IF TRIM(Request("sSkipWaiverAdm")) = "on" THEN sSkipWaiverAdm=1 ELSE sSkipWaiverAdm=0
	IF TRIM(Request("sSkipWaiverHQ")) = "on" THEN sSkipWaiverHQ=1 ELSE sSkipWaiverHQ=0
	IF TRIM(Request("sForceWaiver")) = "on" THEN sForceWaiver=1 ELSE sForceWaiver=0
	IF TRIM(Request("sForceWaiverAdm")) = "on" THEN sForceWaiverAdm=1 ELSE sForceWaiverAdm=0
	IF TRIM(Request("sForceWaiverHQ")) = "on" THEN sForceWaiverHQ=1 ELSE sForceWaiverHQ=0

	IF TRIM(Request("sDispDebugButtons")) = "on" THEN sDispDebugButtons=1 ELSE sDispDebugButtons=0
	IF TRIM(Request("sDispDebugButtonsAdm")) = "on" THEN sDispDebugButtonsAdm=1 ELSE sDispDebugButtonsAdm=0
	IF TRIM(Request("sDispDebugButtonsHQ")) = "on" THEN sDispDebugButtonsHQ=1 ELSE sDispDebugButtonsHQ=0

END SUB



' ----------------------------------
  SUB ReadControlDisplayTableValues
' ----------------------------------

' --- Read transactions from Credit Card Table to determine Total Fees actually completed ----
SET rsContDisp=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT * FROM "&ControlDisplayTableName
rsContDisp.open sSQL, SConnectionToTRATable, 3, 3

IF NOT rsContDisp.eof THEN

	' --- Values read from table are convert back to control the form checkboxes ---
	sEntryEmail= rsContDisp("EntryEmail")
	sEntryEmailAdm=rsContDisp("EntryEmailAdm")
	sEntryEmailHQ=rsContDisp("EntryEmailHQ")
	sWaiverEmail=rsContDisp("WaiverEmail")
	sWaiverEmailAdm=rsContDisp("WaiverEmailAdm")
	sWaiverEmailHQ=rsContDisp("WaiverEmailHQ")
	sPasswordEmail=rsContDisp("PasswordEmail")
	sPasswordEmailAdm=rsContDisp("PasswordEmailAdm")
	sPasswordEmailHQ=rsContDisp("PasswordEmailHQ")
	sSkipWaiver=rsContDisp("SkipWaiver")
	sSkipWaiverAdm=rsContDisp("SkipWaiverAdm")
	sSkipWaiverHQ=rsContDisp("SkipWaiverHQ")
	sForceWaiver=rsContDisp("ForceWaiver")
	sForceWaiverAdm=rsContDisp("ForceWaiverAdm")
	sForceWaiverHQ=rsContDisp("ForceWaiverHQ")

	sDispDebugButtons=rsContDisp("DispDebugButtons")
	sDispDebugButtonsAdm=rsContDisp("DispDebugButtonsAdm")
	sDispDebugButtonsHQ=rsContDisp("DispDebugButtonsHQ")

END IF


END SUB



' -------------------
  SUB StoreTheValues
' -------------------


	OpenCon
	sSQL = "UPDATE "&ControlDisplayTableName
	sSQL = sSQL + " SET EntryEmail = "&sEntryEmail&", WaiverEmail="&sWaiverEmail&", PasswordEmail = "&sPasswordEmail
	sSQL = sSQL + " , SkipWaiver= "&sSkipWaiver&", ForceWaiver= "&sForceWaiver
	sSQL = sSQL + " , EntryEmailAdm = "&sEntryEmailAdm&", WaiverEmailAdm="&sWaiverEmailAdm&", PasswordEmailAdm = "&sPasswordEmailAdm
	sSQL = sSQL + " , SkipWaiverAdm= "&sSkipWaiverAdm&", ForceWaiverAdm= "&sForceWaiverAdm
	sSQL = sSQL + " , EntryEmailHQ = "&sEntryEmailHQ&", WaiverEmailHQ="&sWaiverEmailHQ&", PasswordEmailHQ = "&sPasswordEmail
	sSQL = sSQL + " , SkipWaiverHQ= "&sSkipWaiverHQ&", ForceWaiverHQ= "&sForceWaiverHQ
	sSQL = sSQL + " , DispDebugButtons="&sDispDebugButtons&", DispDebugButtonsAdm="&sDispDebugButtonsAdm&", DispDebugButtonsHQ="&sDispDebugButtonsHQ

	con.execute(sSQL)

	closecon


END SUB

%>





