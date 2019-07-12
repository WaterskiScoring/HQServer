<%
' ------------------------------------------------------------------------------------------------------------------
' ----- Eago_Tools_Drops.asp      Tools where an include statement is put in the calling program
' ------------------------------------------------------------------------------------------------------------------






'*********************************************************************************
 SUB LoadMemberListDropDown (MemberSelected, MemberDropName, MemberDropStatus)
'*********************************************************************************

sSQL = "SELECT M.MemberID, M.Company, M.First, M.Last, M.Address1, M.Address2, M.City, M.State, M.Zip"
sSQL = sSQL + " , M.WorkPhone, M.CellPhone, M.Email, M.Category, M.Status "
sSQL = sSQL + " FROM "&MembersTableName&" as M"
sSQL = sSQL + " WHERE M.Status = 'A'"
sSQL = sSQL + " ORDER BY Company"

'response.write("sSQL= "&sSQL)
'response.end


SET rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable


' ---------------------------------------------------------------------
' --------------- Builds VALID Members DROP DOWN list   ---------------
' ---------------------------------------------------------------------

%><select name='<%=MemberDropName%>' <%=MemberDropStatus%> style="width:20em" onchange=submit()>
<option value="None">Select Member</option><%
IF MemberDropName="sToMemberID" AND sToMemberID="All" THEN %>
	<option value="All" selected>All Members</option><%
ELSEIF MemberDropName="sToMemberID" THEN %>
	<option value="All">All Members</option><%
END IF

IF NOT rs.eof THEN 
  	rs.movefirst

  	DO WHILE NOT rs.eof
		IF TRIM(rs("MemberID")) = MemberSelected THEN %>
			<option value="<%=rs("MemberID")%>" selected><%=rs("Company")%></option><br><%
 		ELSE %>
			<option value="<%=rs("MemberID")%>"><%=rs("Company")%></option><br><%
		END IF	

		rs.moveNEXT
	LOOP
END IF  %>
</select><%

rs.close



f=1
IF f=2 THEN %>
<link rel="stylesheet" type="text/css" href="/css/styles.css" /><%
END IF


END SUB



' ---------------------------------------------
   SUB BuildStateDropDown  (StateListStatus)
' ---------------------------------------------
  StateArray = Split(USStatesList,",")  
  %>
	<select name="sLeadState" <%= StateListStatus %>><%
		  FOR kvar = 0 TO UBOUND(StateArray)
		    IF TRIM(sLeadState) = TRIM(StateArray(kvar)) THEN
					response.write("<option value = """&sLeadState&""" SELECTED>"&sLeadState&"</option>")
			  ELSE
					response.write("<option value = """&StateArray(kvar)&""">"&StateArray(kvar)&"</option>")
		    END IF
			NEXT  
	%>
	</select><%

END SUB


' --------------------------------------------------------------
   SUB LoadContactUrgencyDropDown (ContactUrgencyDropStatus)
' --------------------------------------------------------------

%><select name='sContactUrgency' <%=UrgencyDropStatus%> style="width:12em">
	<option value =''<%IF sContactUrgency = "" THEN Response.Write(" selected ")%>>Select</Option><br>
	<option value ='Urgent'<%IF sContactUrgency = "Urgent" THEN Response.Write(" selected ")%>>Urgent</Option><br>
	<option value ='ASAP'<%IF sContactUrgency = "ASAP" THEN Response.Write(" selected ")%>>ASAP</Option><br>
	<option value ='Next 7 Days'<%IF sContactUrgency = "Next 7 Days" THEN Response.Write(" selected ")%>>Next 7 Days</Option><br>
	<option value ='Next 30 Days'<%IF sContactUrgency = "Next 30 Days" THEN Response.Write(" selected ")%>>Next 30 Days</Option><br>
	<option value ='Undetermined'<%IF sContactUrgency = "Undetermined" THEN Response.Write(" selected ")%>>Undetermined</Option><br>
</select><%

END SUB


' --------------------------------------------------------------
   SUB LoadPurchaseUrgencyDropDown (PurchaseUrgencyDropStatus)
' --------------------------------------------------------------

%><select name='sPurchaseUrgency' <%=UrgencyDropStatus%> style="width:12em">
	<option value =''<%IF sPurchaseUrgency = "" THEN Response.Write(" selected ")%>>Select</Option><br>
	<option value ='Urgent'<%IF sPurchaseUrgency = "Urgent" THEN Response.Write(" selected ")%>>Urgent</Option><br>
	<option value ='ASAP'<%IF sPurchaseUrgency = "ASAP" THEN Response.Write(" selected ")%>>ASAP</Option><br>
	<option value ='Next 7 Days'<%IF sPurchaseUrgency = "Next 7 Days" THEN Response.Write(" selected ")%>>Next 7 Days</Option><br>
	<option value ='Next 30 Days'<%IF sPurchaseUrgency = "Next 30 Days" THEN Response.Write(" selected ")%>>Next 30 Days</Option><br>
	<option value ='Next 6 Months'<%IF sPurchaseUrgency = "Next 6 Months" THEN Response.Write(" selected ")%>>Next 6 Months</Option><br>
	<option value ='Next Year'<%IF sPurchaseUrgency = "Next Year" THEN Response.Write(" selected ")%>>Next Year</Option><br>
	<option value ='Undetermined'<%IF sPurchaseUrgency = "Undetermined" THEN Response.Write(" selected ")%>>Undetermined</Option><br>
</select><%

END SUB


' -------------------------------------------------------
   SUB LoadLeadTypeDropDown (LeadTypeDropStatus)
' -------------------------------------------------------

%><select name='sLeadType' <%=LeadTypeDropStatus%> style="width:12em">
	<option value =''<%IF sLeadType = "" THEN Response.Write(" selected ")%>>Select</Option><br>
	<option value ='Direct'<%IF sLeadType = "Direct" THEN Response.Write(" selected ")%>>Direct</Option><br>
	<option value ='Confidential'<%IF sLeadType = "Confidential" THEN Response.Write(" selected ")%>>Confidential</Option><br>
	<option value ='First Business'<%IF sLeadType = "First Business" THEN Response.Write(" selected ")%>>First Business</Option><br>
	<option value ='Repeat Business'<%IF sLeadType = "Repeat Business" THEN Response.Write(" selected ")%>>Repeat Business</Option><br>
</select><%

END SUB



' -------------------------------------------------------
   SUB LoadLeadCategoryDropDown (LeadCategoryDropStatus)
' -------------------------------------------------------

%><select name='sLeadCategory' <%=LeadCategoryDropStatus%> style="width:15em">
	<option value =''<%IF sLeadCategory = "" THEN Response.Write(" selected ")%>>Select</Option><br>
	<option value ='Relocation - Business'<%IF sLeadCategory = "Relocation - Business" THEN Response.Write(" selected ")%>>Relocation - Business</Option><br>
	<option value ='Relocation - Individuals'<%IF sLeadCategory = "Relocation - Individuals" THEN Response.Write(" selected ")%>>Relocation - Individuals</Option><br>
	<option value ='Construction - New'<%IF sLeadCategory = "Construction - New" THEN Response.Write(" selected ")%>>Construction - New</Option><br>
	<option value ='Construction - Remodel'<%IF sLeadCategory = "Construction - Remodel" THEN Response.Write(" selected ")%>>Construction - Remodel</Option><br>
	<option value ='Personnel Change'<%IF sLeadCategory = "Personnel Change" THEN Response.Write(" selected ")%>>Personnel Change</Option><br>
  <option value ='Ownership Change'<%IF sLeadCategory = "Ownership Change" THEN Response.Write(" selected ")%>>Ownership/Management Change</Option><br>
</select><%

END SUB



' ---------------------
   SUB DefineEAGOStyles 
' ----------------------

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Welcome to EAGO</title>
<script language="javascript" type="text/JavaScript" src="/jscripts/scripts.js"></script>
<script language="javascript" type="text/javascript" src="/jscripts/swfobject.js"></script>
<style type="text/css">


/* this style applies to the DropTable table */
table.droptable {padding:2px; background-position: center right; border:3px solid <%=EAGOColor2%>}
/* this style applies to all th (cells) within the 'scores' table */ 
table.droptable th {padding:1px; border:0px solid <%=EAGOColor3%>;} 
/* this style applies to all td (cells) within the 'scores' table */ 
table.droptable td {padding:1px; vertical-align:middle; border:0px solid <%=EAGOColor2%>;} 
/* this style applies to the scores table */

/* this style applies to the BlankTable table */
table.blanktable {padding:2px; background-position: center right;}
/* this style applies to all th (cells) within the 'blank' table */ 
table.blanktable th {padding:1px; border:0px solid <%=EAGOColor3%>;} 
/* this style applies to all td (cells) within the 'blank' table */ 
table.blanktable td {padding:1px; vertical-align:middle; border:0px solid <%=EAGOColor2%>; word-wrap:break-word;} 
/* this style applies to the blank table */

/*
/* this style applies to the Scores table */
table.scores {padding:0px; border:3px solid <%=EAGOColor2%>; border-collapse:collapse;}
/* this style applies to all th (cells) within the 'scores' table */ 
table.scores th {padding:0px; background-color:<%=EAGOColor2%>; border:1px solid black; vertical-align:bottom;} 
/* this style applies to all td (cells) within the 'scores' table */ 
table.scores td {padding:0px; border:1px solid <%=EAGOColor2%>; border-style:solid; background-color:<%=TableColor1%>; vertical-align:middle;} 

/*
/* this style applies to the SpaceTable table */
table.spacetable {padding:2px; border:1px solid <%=EAGOColor2%>}
/* this style applies to all th (cells) within the 'spacetable' table */ 
table.spacetable th {padding:3px; border:1px solid black; background-color:<%=TableColor1%>; vertical-align:bottom;} 
/* this style applies to all td (cells) within the 'spacetable' table */ 
table.spacetable td {padding:6px; border:1px solid black; background-color:<%=TableColor1%>; vertical-align:middle;} 

/*
/* this style applies to the innertable table */
table.innertable {padding:0px; border:1px solid <%=EAGOColor2%>; border-collapse:collapse;}
/* this style applies to all th (cells) within the 'innertable' table */ 
table.innertable th {padding:1px; border:1px solid <%=EAGOColor1%>; border-style:solid; background-color:<%=EAGOColor2%>; vertical-align:bottom;} 
/* this style applies to all td (cells) within the 'innertable' table */ 
table.innertable td {padding:3px; border:1px solid <%=EAGOColor2%>; border-style:solid; background-color:<%=TableColor1%>; vertical-align:middle;  word-wrap:break-word;} 

/*
/* this style applies to the messagetable table */
table.messagetable {padding:0px; border:1px solid <%=EAGOColor2%>; border-collapse:collapse;}
/* this style applies to all th (cells) within the 'messagetable' table */ 
table.messagetable th {padding:1px; border:1px solid <%=EAGOColor1%>; border-style:solid; background-color:<%=EAGOColor2%>; vertical-align:bottom;} 
/* this style applies to all td (cells) within the 'messagetable' table */ 
table.messagetable td {padding-left:15px; padding-right:15px; padding-top:8px; padding-bottom:8px; border:0px solid <%=EAGOColor2%>; border-style:solid; background-color:<%=TableColor1%>; vertical-align:middle; white-space:nowrap;} 

/*
/* this style applies to the noborder table */
table.noborder {padding:0px; border:0px solid <%=EAGOColor2%>; border-collapse:collapse;}
/* this style applies to all th (cells) within the 'noborder' table */ 
table.noborder th {padding:1px; border:0px solid <%=EAGOColor1%>; border-style:solid; background-color:white; vertical-align:bottom;} 
/* this style applies to all td (cells) within the 'noborder' table */ 
table.noborder td {padding:1px; text-align: left; border:0px solid <%=EAGOColor2%>; border-style:solid; background-color:white; vertical-align:middle; word-wrap:break-word;} 

/*
/* this style applies to the tourlist table */
table.tourlist {padding:0px; border:1px solid <%=EAGOColor2%>; border-collapse:collapse; overflow-scroll}
/* this style applies to all th (cells) within the 'tourlist' table */ 
table.tourlist th {padding:1px; border:1px solid <%=EAGOColor1%>; border-style:solid; background-color:<%=EAGOColor2%>; vertical-align:bottom;} 
/* this style applies to all td (cells) within the 'tourlist' table */ 
table.tourlist td {padding:3px; border:1px solid <%=EAGOColor2%>; border-style:solid; background-color:<%=TableColor1%>; vertical-align:middle;  word-wrap:break-word;} 

</style>
</head><%

END SUB




' ---------------------------
    SUB SendLeadSubmitEMail
' ---------------------------


Dim DateNow, TimeNow

DateNow = Date
TimeNow = Time

ebody = "<html>"
ebody = ebody & "<head>"
ebody = ebody & "<title>Important Lead Notice</title>"
ebody = ebody & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
ebody = ebody & "</head>"
ebody = ebody & "<body bgcolor=""#FFFFFF"" text=""#000000"">"
ebody = ebody & "<div align=""center"">"


ebody = ebody & "<TABLE BORDER=4 ALIGN=CENTER CELLPADDING=3 CELLSPACING=0 BGCOLOR="&TableColor1&" width=350px >"
ebody = ebody & "<TR>"
ebody = ebody & "<TD BGCOLOR=red><center><font face="&font1&" color=#FFFFFF size=4><b>Important Lead Notice</b></font></TD>"
ebody = ebody & "</TR>"
 
ebody = ebody & "<TR>"
ebody = ebody & "<TD VALIGN=top>"


ebody = ebody & "<table border=""0"" cellspacing=""0"" cellpadding=""3"" width=""100%"">"
ebody = ebody & "<tr>"

ebody = ebody & "<td Align=center>"	
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>Lead submitted by EAGO Member</b></font>"
ebody = ebody & "<br>"

ebody = ebody & "<font color="&TextColor2&" face="&font1&" size=2>"&sFromFirst&" "&sFromLast&"</font>"
ebody = ebody & "<br>"
ebody = ebody & "<font color="&TextColor2&" face="&font1&" size=3><b>"&sFromCompany&"</b></font>"
ebody = ebody & "<br><br><br>"

ebody = ebody & "<font face="&font1&" size=3><b>Lead Information </b></font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=1><b>Name </b></font><font color="&TextColor2&" face="&font1&" size=2>"&sLeadFirst&" "&sLeadLast&"</font>"
ebody = ebody & "<br>"

ebody = ebody & "<font face="&font1&" size=1><b>Company </b></font><font color="&TextColor2&" face="&font1&" size=2>"&sLeadCompany&"</font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=1><b>Address </b></font><font color="&TextColor2&" face="&font1&" size=2>"&sLeadAddress1&"</font>"
ebody = ebody & "<br>"
IF TRIM(sLeadAddress2)<>"" THEN
	ebody = ebody & "<font face="&font1&" size=1></font><font color="&TextColor2&" face="&font1&" size=2>"&sLeadAddress2&"</font>"
	ebody = ebody & "<br>"
END IF
ebody = ebody & "<font face="&font1&" size=1><b>City </b></font><font color="&TextColor2&" face="&font1&" size=2>"&sLeadCity&",</font>"
ebody = ebody & "<font face="&font1&" size=1></font><font color="&TextColor2&" face="&font1&" size=2>"&sLeadState&"</font>"
ebody = ebody & "<font face="&font1&" size=1></font><font color="&TextColor2&" face="&font1&" size=2>"&sLeadZip&"</font>"
ebody = ebody & "<br><br>"

ebody = ebody & "<font face="&font1&" size=1><b>Work Phone </b></font><font color="&TextColor2&" face="&font1&" size=2>"&sLeadWorkPhone&"</font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=1><b>Cell Phone </b></font><font color="&TextColor2&" face="&font1&" size=2>"&sLeadCellPhone&"</font>"
ebody = ebody & "<br><br>"
ebody = ebody & "<font face="&font1&" size=1><b>Email </b></font><font color="&TextColor2&" face="&font1&" size=2>"&sLeadEmail&"</font>"
ebody = ebody & "<br><br>"


ebody = ebody & "<font face="&font1&" size=2><b>Special Information</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=1><b>Contact Urgency </b></font><font color="&TextColor2&" face="&font1&" size=2>"&sContactUrgency&"</font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=1><b>Purchase Urgency </b></font><font color="&TextColor2&" face="&font1&" size=2>"&sPurchaseUrgency&"</font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=1><b>Lead Type </b></font><font color="&TextColor2&" face="&font1&" size=2>"&sLeadType&"</font>"
ebody = ebody & "<br><br>"


ebody = ebody & "<br>"

ebody = ebody & "</center>"
ebody = ebody & "<br>"
ebody = ebody & "</td></tr>"

ebody = ebody & "</TABLE>"

ebody = ebody & "</TD></TR>"
ebody = ebody & "</TABLE>"




Dim objCDO
Dim MembEntryEmail, HQEntryEmail, SendAddress


marksemail="mark@productdesign-biz.com"
FromEmail="eago@eago.com"
ExecDirEmail="lora@productdesign-biz.com"


Set objCDO = Server.CreateObject("CDO.Message")
objCDO.To = sToEMail
objCDO.CC = ExecDirEmail

sEntryEmailMC=true
IF sEntryEmailMC=true THEN 
	objCDO.BCC = " "&marksemail
END IF

objCDO.From = ""&FromEmail
objCDO.Subject = " Lead Confirmation for "&sLeadCompany&" "&sLeadFirst&" "&sLeadLast 
objCDO.HTMLBody = ebody	

response.write(ebody)
'response.end

objCDO.Send

Set objCDO = Nothing



END SUB


%>
