<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"--><%


Dim sOldEmail, sOldPassword, sOldStatus, sCreateDate, sStatus, sFirstName, sLastName
Dim sNewEmail, sNewPassword





sRunByWhat = TRIM(LCASE(Request("pvar")))

SELECT CASE sRunByWhat

   CASE "save"
	ReadFormVariables
	StoreTheValues
	ReadTableValues
	DisplayPasswordPage
   CASE "member"
	FindMember
	ReadTableValues
	DisplayPasswordPage
	
   CASE ELSE

	ReadTableValues
	DisplayPasswordPage
END SELECT




' -------------------------
  SUB DisplayPasswordPage
' -------------------------


	DefineTRAStyles

	WriteIndexPageHeader  %>

	<br><br><br><br>
    <TABLE class="innertable" Align="center" width=70%>
        <form name="ContDisp1" method="post" action="/rankings/PW_Update_Tool.asp" id="ContDisp1">
	<TR>
	  <TH align=center colspan=4 bgcolor="#2F4F4F"><font size="4" color="#FFFFFF"><b>OLR Email & Password Update</b></font></TD> 
	</TR>

	<TR>
	  <TD align=right width=33%><font size="<%=fontsize2%>" color="<%=TextColor1%>">Name&nbsp;</font></TD> 
	  <TD align=left colspan=2 width=67%><font size="<%=fontsize2%>" color="<%=TextColor2%>">&nbsp;<%=sFirstName%>&nbsp;<%=sLastName%></font></TD> 
	</TR>

	<TR>
	  <TD align=right><font size="<%=fontsize2%>" color="<%=TextColor1%>"> Old Email&nbsp;</font></TD> 
	  <TD align=left colspan=2><font size="<%=fontsize2%>" color="<%=TextColor2%>">&nbsp;<%=sOldEmail%></font></TD> 
	</TR>
	<TR>
	  <TD align=right><font size="<%=fontsize2%>" color="<%=TextColor1%>"> New Email&nbsp;</font></TD> 
	  <TD align=left colspan=2><input type="text" name="sNewEmail" maxlength=35 size=38></TD> 
	</TR>
	<TR>
	  <TD align=right><font size="<%=fontsize2%>" color="<%=TextColor1%>"> Old Password&nbsp;</font></TD> 
	  <TD align=left colspan=2><font size="<%=fontsize2%>" color="<%=TextColor2%>">&nbsp;<%=sOldPassword%></font></TD> 
	</TR>
	<TR>
	  <TD align=right><font size="<%=fontsize2%>" color="<%=TextColor1%>"> New Password&nbsp;</font></TD> 
	  <TD align=left colspan=2><input type="text" name="sNewPassword" maxlength=35 size=38></TD> 
	</TR>

	<TR>
	  <TD align=right><font size="<%=fontsize2%>" color="<%=TextColor1%>"> Status&nbsp;</font></TD> 
	  <TD align=left colspan=2>
		<select name="sStatus" style="width:10em">
		  <option value ="T" <%IF sOldStatus = "T" THEN Response.Write(" selected ")%> >Temporary</Option><br>
		  <option value ="A" <%IF sOldStatus = "A" THEN Response.Write(" selected ")%> >Active</Option><br>
		</select>
	  </TD> 
	</TR>

	<TR>
	  <td align="center">
		<br>
		<input type="submit" name="pvar" value="Save" style="width:9em" title="Save the current settings to the table">
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

        <form name="ContDisp2" method="post" action="/rankings/PW_Update_Tool.asp?" id="ContDisp2">
	  <td align="center">
		<br>
		<input type="submit" name="pvar" value="Member" style="width:9em" title="Select a Member to Update Password">
		<br>
	  </td>
	</form>

	</TR>

   </TABLE>
<br><br>
<center><font size="<%=fontsize3%>"><b>If 'New Email' or Password field are left blank, the system will retain 'Old Email' and 'Old Password' when Saved.</b></font></center>

<%

WriteIndexPageFooter


END SUB



' ------------------
  SUB FindMember
' ------------------

	Session("sSendingPage")="/rankings/PW_Update_Tool.asp"
	Response.Redirect("/rankings/search-memberHQ.asp?rid="&rid&"&formstatus=search")

END SUB



' -----------------------
  SUB ReadFormVariables
' -----------------------

	sNewEmail=TRIM(Request("sNewEmail"))
	sStatus=TRIM(Request("sStatus"))
	sNewPassword=TRIM(Request("sNewPassword"))

END SUB



' ----------------------------------
  SUB ReadTableValues
' ----------------------------------

sMemberID=TRIM(Session("sMemberID"))


' --- Read transactions from Credit Card Table to determine Total Fees actually completed ----
SET rsPW=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT PW.Email, PW.Password, PW.CreateDate, PW.Status, MT.FirstName, MT.LastName FROM "&RegPWTableName&" AS PW"
sSQL = sSQL + " JOIN "&MemberTableName&" AS MT ON PW.MemberID=MT.PersonIDWithCheckDigit"
sSQL = sSQL + " WHERE PW.MemberID = '"&sMemberID&"'"


rsPW.open sSQL, SConnectionToTRATable, 3, 3

IF NOT rsPW.eof THEN
	sOldEmail= TRIM(rsPW("Email"))
	sOldPassword= TRIM(rsPW("Password"))
	sOldCreateDate= rsPW("CreateDate")
	sOldStatus= TRIM(rsPW("Status"))
	sFirstName = TRIM(rsPW("FirstName")) 	
	sLastName = TRIM(rsPW("LastName"))
END IF



END SUB



' -------------------
  SUB StoreTheValues
' -------------------


	OpenCon
	sSQL = "UPDATE "&RegPWTableName
	sSQL = sSQL + " SET MemberID='"&Session("sMemberID")&"'"

	IF TRIM(sNewEmail)<>"" THEN sSQL = sSQL + " , Email = '"&sNewEmail&"'"
	IF TRIM(sNewPassword)<>"" THEN sSQL = sSQL + " , Password = '"&sNewPassword&"'"

	IF sOldStatus<>TRIM(sStatus) THEN sSQL = sSQL + " , Status = '"&sStatus&"'"

	sSQL = sSQL + " WHERE MemberID='"&Session("sMemberID")&"'"

	con.execute(sSQL)
'response.write("<br>"&sSQL)
'response.end
	closecon


END SUB

%>



