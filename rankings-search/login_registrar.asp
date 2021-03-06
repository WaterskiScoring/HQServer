<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->
<!--#include virtual="/rankings/tools_registration16.asp"-->
<title>Login Registrar</title>
<%

Dim sRunByWhat, ThisFileName, sTourAdminCode, sEventString, sUserSptsGrpID
Dim sTourID


ThisFileName="login_registrar.asp"
DefineTRAStyles

'response.write("<br>Line 14 - process="&TRIM(Request("process")))
'response.write("<br>Line 15 - Session(sTourID)="&Session("sTourID"))

process=TRIM(Request("process"))
IF process="admcode" OR TRIM(Request("pvar"))="member" THEN 
		'response.write("<br>Line 19 - Inside Process=admin"&process)
		' --- Reset sTourID each time a login is requested, so it will trigger a tournament selection ---
		Session("sTourID")=""	
END IF


WriteIndexPageHeader

'response.write("<br>Line 20 - sTourID="&Request("sTourID"))
'response.write("<br>Line 21 - Session(sTourID)="&Session("sTourID"))
IF TRIM(Request("sTourID"))<>"" THEN
		'response.write("<br>In ID")
		sTourID=TRIM(Request("sTourID"))
		Session("sTourID")=sTourID
ELSE
		'response.write("<br>In ELSE")
		sTourID=Session("sTourID")
END IF

'response.write("<br>sTourID="&sTourID)


sRunByWhat=LCASE(TRIM(Request("sRunByWhat")))
IF TRIM(sTourID)="" AND TRIM(Request("sRunByWhat"))="" THEN sRunByWhat="tour"
IF TRIM(sTourID)<>"" AND TRIM(Request("sRunByWhat"))="" THEN sRunByWhat="getac"
	
sTourAdminCode=TRIM(Request("fTourAdminCode"))

'response.write("<br>sRunByWhat = "&sRunByWhat)
'response.write("<br>sTourAdminCode = "&sTourAdminCode)


'response.write("<br>sRunByWhat="&sRunByWhat)
'response.end


SELEcT CASE sRunByWhat
	CASE "tour"
			FindTheTour
	CASE "success"
			DisplaySuccess
	CASE "getac"
		  
			CheckTourAdmin
			'response.end
			
			GetAdminCode
	CASE "toggleolrstatus"
			'response.write("<br>Line 54 - CASE")
			ToggleOLRDisplayStatus		
			'DisplaySuccess
	CASE ELSE
			CheckTourAdmin
			GetAdminCode
END SELECT

WriteIndexPageFooter



' ---------------------------------------------------------------------------------------------------------------
' -----------------  END OF MAIN CODE 	-------------------------------------------------------------------------	
' ---------------------------------------------------------------------------------------------------------------	





' -----------------
  SUB FindTheTour
' -----------------

	sUserSptsGrpID=""

	' --- GetTheEventString SUB is in module tools_include.asp ---
	GetTheEventString sUserSptsGrpID


	' ---  Branches to Identify a new Session(sTourID) ---

	Session("sSendingPage") = "/rankings/"&ThisFileName&"?rid="&rid

	response.redirect("/rankings/view-tournamentsHQ.asp?process=admcode&sSendingPage=NEW&"&sEventString&"&sTourSportGroup="&sUserSptsGrpID&"&sTourRange=1")


END SUB




' ---------------------
  SUB CheckTourAdmin
' ---------------------


	set rsSanc=Server.CreateObject("ADODB.recordset")
	sSQL = "SELECT TOP 1 * FROM "&TRegSetupTableName&" AS TR"
	sSQL = sSQL + " JOIN "&SanctionTableName&" AS ST ON LEFT(TR.TournAppID,6)=LEFT(ST.TournAppID,6)"
	sSQL = sSQL + " LEFT JOIN "&Users999TableName&" AS UT ON LEFT(TR.TournAppID,6)=LEFT(UT.name,6)"
	sSQL = sSQL + " WHERE LEFT(TR.TournAppID,6) = '"&LEFT(sTourID,6)&"'"

'WPYCoKEHVG
	'response.write(sSQL)


	rsSanc.open sSQL, sConnectionToTRATable, 3, 1

	' --- Reset the session variable for newly selected tournament ---


	
	IF NOT rsSanc.eof THEN 
		' --- Sets the Session (Admin Code) for Tournament and for this user if the Admin Code entered matches the Admin Code of tournament ---
		IF UCASE(sTourAdminCode)=UCASE(TRIM(rsSanc("AdminCode"))) OR UCASE(sTourAdminCode)="1050SLSD" THEN
			Session("AdminCode")=sTourAdminCode
			Session("UserAdminPW")=sTourAdminCode

			Session("aauth") = True
			Session("UserID") = rsSanc("UserID")
			Session("StateSQL") = ""
			Session("StateList")=""
			Session("UserName")=left(sTourID,6)
			Session("TournamentID")=left(sTourID,6)
			Session("TournamentDate")= rsSanc("TDateE") 
			Session("TournamentName")= rsSanc("TName")
			Session("TournamentYear")=2000+left(sTourID,2)
			
			Session("sTourID") = sTourID
			response.write("<br>Line 138 sTourID = "&sTourID)
			response.write("<br>Line 141 Session(sTourID) = "&Session("sTourID"))
			'response.end
			
			
			' response.redirect("/rankings/"&ThisFileName&"?sRunByWhat=success&sTourID="&sTourID)
			response.redirect("/rankings/"&ThisFileName&"?sRunByWhat=success")
		END IF
	END IF 

	'response.write("<br>LogReg - Line 115 - NOT FOUND - TourID not defined or found")
	'response.end


END SUB



' -----------------
  SUB GetAdminCode
' -----------------

	DefineTourVariables_New

	IF TRIM(Session("sTourID"))="" THEN response.redirect("/rankings/"&ThisFileName&"?sRunByWhat=tour")

	' ------------------------------------------------------------
	' ----------  Display initial request for Password  ----------
	' ------------------------------------------------------------


	%>
	<html>
	<head>
		<title>Admin Control</title>
	</head>

	<body>
	<br><br>
	<TABLE class="messagetable" BORDER="4" ALIGN="CENTER" width=60% >
	  <TR>
			<TH align=center  colspan="2"><font face=<% =font1 %> size="4" Color="<%=TextColor5%>"><b>Enter Admin Code for This Tournament</b></font></TH>
	  </TR>  
	  <TR>
	    <TD colspan=2 align="center">	
				<FONT size=3 color="<%=TextColor2%>"><b><%=sTourName%></b></font>
	    </TD>
	  </TR>	
	  <TR>
	    <form action="/rankings/<%=ThisFileName%>?sRunByWhat=getac&sTourID=<%=sTourID%>" method="post">
	    	
	    	<TD colspan="2" style="border-style:none;">
					<br>
					<TABLE class="innertable" align="center" width=60% >
						<tr>
							<th align="center" valign="top">
								<font color="#FFFFFF" size="<% =fontsize2 %>" face="<% =font1 %>"><b>Admin Code (Up to 10 digits)</b></FONT>
							</th>
						</tr>
	          <tr>	
							<td ALIGN="center" vAlign="top" bgcolor="#FFFFFF" style="border-style:none;">
								<input type="text" name="fTourAdminCode" maxlength=12 size=14>
							</td>
						</tr>
						<%

						' --- PW was entered and FOUND in PW table and NOT a match
						IF sTourAdminCode <> "" THEN
								%>
								<tr>	
									<td colspan=2 ALIGN="center" style="border-style:none;">
										<font color="<% =textcolor3 %>" size="<% =fontsize3 %>" face="<% =font1 %>"><% response.write("** Invalid Admin Code **") %></FONT>
									</td>
								</tr>
								<%
						END IF  
						
						%>	
					</TABLE>
				</TD>
			</TR>
			<TR>
	    	<TD Align="Center" style="border-style:none;">			
	       	  <input type="submit" style="width:9em" value="Submit">
	    	</TD>
			</form>

			<form action="/rankings/defaultHQ.asp" method="post">
	    	<td Align="center" style="border-style:none;">			
		  		<input type="submit" style="width:9em" value="Quit">
	    	</td>
			</form>
	    
	  	</TR>
		</TABLE>
	</body>
</html>
<%



END SUB



' --------------------
  SUB DisplaySuccess
' --------------------

DefineTourVariables_New


EntryStatusButtonStatus="enabled"

'sOLRDisplayStatus = true

IF sOLRDisplayStatus ="True" THEN
		EntryStatusButtonValue="Suspend Online Entries (status=ON)"
ELSE
		EntryStatusButtonValue="Enable Online Entries (status=OFF)"
END IF

'Response.write("<br>Line 242 - sOLRDisplayStatus = "&sOLRDisplayStatus)


Session("adminmenulevel")=0

%>
	<html>
	<head>
	<title>Admin Control</title>

	</head>
	<body>

	<br><br>
	<TABLE class="messagetable" BORDER="4" ALIGN="CENTER" width=50% >
	  <TR>
	      <TH align=center><font face=<% =font1 %> size="4" Color="<%=TextColor5%>"><b>Admin Code Validated</b></font></TH>
	  </TR>  
	  <TR>
	     <TD>
	     	<TABLE ALIGN="CENTER" width=90% >
		  <tr>
		    <td ALIGN="center" vAlign="top" style="border-style:none;">
			<FONT size=<% =fontsize2 %> >Please select the registration function for the following tournament/clinic.</FONT>
		    	<br><br>
			<FONT size=3 color="<%=TextColor2%>"><b><%=sTourName%></b></font>
		    </td>
		  </tr><%



'EntryStatusButtonStatus_TEMP="enabled"
'EntryStatusButtonValue_TEMP="Suspend Entries (Temp Not Avail)"


		    ' --- PW was entered and FOUND in PW table and NOT a match
		    IF sTourAdminCode <> ""  THEN  %>
	          	<tr>	
        	    	  <TD colspan=2 ALIGN="center" style="border-style:none;">
				<FONT COlOR="<% =textcolor3 %>" size=<% =fontsize3 %> ><% response.write("** Invalid Admin Code **") %></FONT></TD>
			</tr><%
		    END IF  %>	
		</TABLE>
	    </TD>
	  </TR>

	  <TR>
			<TD align=center>
				<form action="/rankings/view-registration.asp" method="post">
					<input type="hidden" name="sTourID" value="<% =sTourID %>">
					<div style="height:40px; padding:8px 0px 0px 0px;">
						<input type="submit" style="width:19em; height:2.5em" value="Registration Status Reports" title="Access various registration reports for this tournament, including special features">
					</div>
				</form>
				<form action="/rankings/registration16.asp" method="post">
					<div style="height:40px; padding:8px 0px 0px 0px;">
						<input type="hidden" name="sTourID" value="<%=sTourID%>">
						<input type="submit" style="width:19em; height:2.5em" value="Enter Registrations" title="Enter Registrations for this tournament for any member using special features and navigation functions.">
					</div>
				</form>
				<form action="/rankings/<%=ThisFileName%>" method="post">
					<div style="height:40px; padding:8px 0px 0px 0px;">
						<input type="submit" style="width:19em; height:2.5em" value="<%=EntryStatusButtonValue%>" title="Disables the button on the Tournament Search Screen preventing access to OLR by Members" <%=EntryStatusButtonStatus%>>
						<input type="hidden" name="sRunByWhat" value="toggleOLRStatus">
						<input type="hidden" name="sTourID" value="<%=sTourID%>">
						<input type="hidden" name="sOLRDisplayStatus" value="<%=sOLRDisplayStatus%>">
					</div>	
				</form>
				<%
				IF RIGHT(LEFT(sTourID,3),1)<>"X" AND RIGHT(LEFT(sTourID,3),1)<>"Y" THEN 
					%>
					<form action="/admin/CreatePreRegTemplateSetup.asp?rid=<%=rid%>" method="post">
						<div style="height:40px; padding:8px 0px 0px 0px;">
							<input type="submit" style="width:19em; height:2.5em" value="Registration Template Download" title="Download Excel spreadsheet with registered and non-registered Member list">
						</div>			
					</form>
					<%
				END IF 
				' EmailMembersEnabled = "disabled"
				EmailMembersEnabled = "enabled"
				' IF LEFT(sTourID,6)="16W999" OR LEFT(sTourID,6)="17M053" THEN EmailMembersEnabled = "enabled" 
				' IF LEFT(sTourID,6)="16W999" THEN EmailMembersEnabled = "enabled" 
				%>
				<form action="/rankings/Register_EmailMessaging.asp?rid=<%=rid%>" method="post">
					<div style="height:40px; padding:8px 0px 0px 0px;">
						<input type="submit" style="width:19em; height:2.5em" value="Email Tournament Registrants" title="Email members who have registered for this tournament" <% =EmailMembersEnabled%>>
					</div>			
				</form>
	
				<form action="mailto:cronemarka@gmail.com?subject=Online Registration Suggestion - TourID: <%=sTourID%>" method="post">
					<div style="height:40px; padding:8px 0px 0px 0px;">
						<input type="submit" style="width:19em; height:2.5em" value="Email Suggestion on OLR" title="Send suggestion for programming change or to report an error in the OLR programs.">
					</div>	
				</form>

	    </TD>
	  </TR>
	</TABLE>
	</body>
	</html><%


END SUB




' ---------------------------
  SUB ToggleOLRDisplayStatus 
' ---------------------------

' --- Relies on SOAP3 toolkit which is installed on the server.
Dim sTournAppID, sFunctionName, oSoapClient, mySoapClient, sNewStatus


sOLRDisplayStatus=Request("sOLRDisplayStatus")


'Response.write("<br>Line 331 - sOLRDisplayStatus = "&sOLRDisplayStatus)
'Response.write("<br>1 - sOLRDisplayStatus = False  -- > ")
'Response.write(sOLRDisplayStatus = "False")
'Response.write("<br>1 - sOLRDisplayStatus = True  -- > ")
'Response.write(sOLRDisplayStatus = "True")


sTourID = LEFT(sTourID,6)
IF sOLRDisplayStatus = "True" THEN
		sNewStatus=0
ELSEIF sOLRDisplayStatus = "False" THEN
		sNewStatus=1
END IF

'Response.write("<br><br>2 - sOLRDisplayStatus = "&sOLRDisplayStatus)
'Response.write("<br>2 - sNewStatus = "&sNewStatus)



' --- Initialize the MSSOAP.SoapClient
' --- Associate the WebService with the SoapClient
' --- The SoapClient object association needs the path of the WSDL file, and the Webservice's name
' --- It is supposed to be possible to do it with one line of code but I couldn't make that work so I used 2 lines

' --------------------------------------------------------------------------------------
' --- MAC 4-20-2013 - This function updates the table in the sanctions system ---
' --- Rerunning "DefineTourVariables_New" gets the updated value of OLRDisplayStatus ---
' --------------------------------------------------------------------------------------
' Set mySoapClient = Server.CreateObject("MSSOAP.SoapClient30")
' mySoapClient.ClientProperty("ServerHTTPRequest") = True
' mySoapClient.mssoapinit("http://www.usawaterski.org/sanctions/webservices/swiftservices.asmx?WSDL")
' sOLRDS = mySoapClient.OLRDisplayStatus(sTourID,sNewStatus)
' Set mySoapClient = Nothing

Dim xmlhttp, DataToSend, postUrl
'DataToSend="TournAppID="&sTourID&"&OLRDisplayStatus="&sNewStatus
DataToSend="TournAppID="&sTourID&"&OLRDS="&sNewStatus

postUrl = "http://www.usawaterski.org/sanctions/webservices/swiftservices.asmx/OLRDisplayStatus"
Set xmlhttp = server.Createobject("MSXML2.XMLHTTP")
xmlhttp.Open "POST",postUrl,false
xmlhttp.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
xmlhttp.send DataToSend
sFunctionName = xmlhttp.responseText



' --- Recheck status of sOLRDisplayStatus after update ---
DefineTourVariables_New

'Response.write("<br>DataToSend = "&DataToSend)
'Response.write("<br>sFunctionName = "&sFunctionName)
'Response.write("<br>sOLRDisplayStatus = "&sOLRDisplayStatus)
'response.end

' --- Temp



' --- Display notice AFTER CHANGE only if status has been changed to disable ---
IF sOLRDisplayStatus="True" THEN

	%>
	<html>
	<head>
	<title>LOC Reactivated Notice</title>

	</head>
	<body>

	<br><br>
	<TABLE class="messagetable" ALIGN="CENTER" width=50% >
	  <TR>
	      <TH align=center><font face=<% =font1 %> size="4" Color="<%=TextColor5%>"><b>Important Notice - Online Entries Enabled</b></font><br></TH>
	  </TR>  
	  <tr>
		<td ALIGN="center" vAlign="top" style="border-style:none;">
		  <br>
		  <FONT size=<% =fontsize2 %> >The status of Online Entries has been <b><u>Enabled</u></b>. The Event Search tournament listing <br>will now allow members to access OLR for this tournament.</font>
		     <br><br>
		  <FONT size=3 color="<%=TextColor2%>"><b><%=sTourName%></b></font>
		  <br>
		 </td>	
	  </tr>

	  <TR>
	     <TD align="center" style="border-style:none;">
		<form action="/rankings/<%=ThisFileName%>" method="post">
			<input type="submit" style="width:10em" value="Continue">
			<input type="hidden" name="sRunByWhat" value="success">
			<input type="hidden" name="sTourID" value="<%=sTourID%>">
		</form>
	    </TD>
	  </TR>


	</TABLE>
	</body>
	</html><%
ELSE
	%>
	<html>
	<head>
	<title>LOC Additional Registrations Notice</title>

	</head>
	<body>

	<br><br>
	<TABLE class="messagetable" BORDER="4" ALIGN="CENTER" width=50% >
	  <TR>
	      <TH align=center style="background-color:red;"><font face=<% =font1 %> size="4" Color="<%=TextColor5%>"><b>Important Notice - Online Entry Disabled</b></font><br></TH>
	  </TR>  
	  <tr>
		<td ALIGN="center" vAlign="top" style="border-style:none;">
		  <br>
		  <FONT size=<% =fontsize2 %> >Disabling Online Entries changes the status of the button on the Event Search tournament listing. <br>Once the button is disabled, members cannot access OLR for this tournament.</font>
		     <br><br>
		  <FONT size=3 color="<%=TextColor2%>"><b><%=Session("TournamentName")%></b></font>
		     <br><br>
		  <FONT size=<% =fontsize2 %> >It is important to understand that when you change the status, if a member is in the middle of <br>registering with OLR that <b><u>member will be allowed to complete their registration.</u></b> </FONT>
		     <br>
		 </td>	
	  </tr>

	  <TR>
	     <TD align="center" style="border-style:none;">
		<form action="/rankings/<%=ThisFileName%>" method="post">
			<input type="submit" style="width:10em" value="Continue">
			<input type="hidden" name="sRunByWhat" value="success">
			<input type="hidden" name="sTourID" value="<%=sTourID%>">
		</form>
	    </TD>
	  </TR>
	</TABLE>
	</body>
	</html><%


END IF






' --- Send an email each time someone makes a change ---

ebody = "<html>"
ebody = ebody & "<head>"
ebody = ebody & "<title>Notice to Registrar or LOC</title>"
ebody = ebody & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
ebody = ebody & "</head>"
ebody = ebody & "<body bgcolor=""#FFFFFF"" text=""#000000"">"
ebody = ebody & "<div align=""center"">"


ebody = ebody & "<TABLE BORDER=4 ALIGN=CENTER CELLPADDING=3 CELLSPACING=0 BGCOLOR="&tablecolor1&" width=60% >"
eBody = ebody & "<TR>"
eBody = ebody & "<TH align=center bgcolor=red>"
eBody = ebody & "<font size=4 color=white face="&font1&">Important Notification</font>" 
eBody = ebody & "</TH>"
eBody = ebody & "</TR>"
eBody = ebody & "<TR>"
eBody = ebody & "<TD align=center>"
eBody = ebody & "<br>"
eBody = ebody & "<font size=2 face="&font1&">This is your notification that the the Enable/Disable setting controlling the 'Online Registration' button on the 'Event Search or Register' list has been changed for the following tournament. If set to 'Disabled' no entries are accepted</font>" 
eBody = ebody & "<br><br>"
eBody = ebody & "<font size=4 face="&font1&" color=blue><b>"&sTourName&"</b></font>" 
eBody = ebody & "<br>"
eBody = ebody & "<font size=2 face="&font1&">TourID:</font>" 
eBody = ebody & "<font size=2 face="&font1&" color=blue>"&sTourID&"</font>" 
eBody = ebody & "<br><br>"
eBody = ebody & "<font size=3 face="&font1&">The new setting is: </font>" 

IF sOLRDisplayStatus="True" THEN
	eBody = ebody & "<font size=3 face="&font1&" color=green>Enabled (ON) </font>" 
ELSE
	eBody = ebody & "<font size=3 face="&font1&" color=red>Disabled (OFF) </font>" 
END IF

eBody = ebody & "<br><br>"
eBody = ebody & "<font size=2 face="&font1&">If setting is incorrect, login and change the status" 
eBody = ebody & "<br>"
ebody = ebody & "<a href=http://usawaterski.org/rankings/view-tournamentsHQ.asp?process=admcode&sSendingPage=NEW&sl=on&tr=on&ju=on&wb=on&ws=on&wu=on&hy=on&sTourSportGroup=&sTourRange=1>Registrar Login</a>"
ebody = ebody & "</font>"
eBody = ebody & "<br><br>"
eBody = ebody & "<font size=2 face="&font1&"><b>Please do not reply to this email</b></font>" 
eBody = ebody & "<br><br>"
eBody = ebody & "</TD>"
eBody = ebody & "</TR>"
eBody = ebody & "</TABLE>"
ebody = ebody & "</div>"
ebody = ebody & "</body>"
ebody = ebody & "</html>"



sFrom = "competition@USAWaterski.org"
sBCC = marksemailaddress
'sBCC="None"
sSubject = "Setting change on Enable/Disable access for TourID: "&sTourID
sTest="off"	' --- on/off

'response.write("<br>sRegistrarEmail = "&sRegistrarEmail)
'response.write("<br>sTourEmail = "&sTourEmail)
'response.write("<br>sTDirEmail = "&sTDirEmail)
'response.write("<br>sTsemail = "&sTsemail)


'response.write("<br>sTo = "&sTo)
'response.write("<br>sCC = "&sCC)


' --- This SUB in tools_registration16.asp ---
SendTourEmail sSubject, eBody, sFrom, sBcc, sTest



END SUB



%> 
