<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_mobile_version.asp"-->
<!--#include virtual="/rankings/tools_registration16.asp"-->

<%



Dim ThisFileName, action
Dim LinkEmail, SMS_PhoneNumber

ThisFileName = "Mobile_SendAppLink.asp"

' -- Use regular unless originating from TEST version --
MenuFileName = "MainMenu_m.asp"
IF Session("TESTING")="mainmenu_test" THEN MenuFileName = "MainMenu_m_TEST.asp"


' --- Displays the html, head and opening body tag ---
OpenState="sendlink"
DisplayHeadOpenBodyAndBannerTags OpenState

uyt=2
IF uyt=1 THEN 
%><div class="errorbox" style="width:95%; border:1px solid white; margin:0px 0px 0px 0px; padding:0px 0px 0px 0px; height:20px;"></div><%
END IF



DisplayNoUserSetScreen			' --- Initially hidden ---


action = TRIM(LCASE(Request("action")))

SELECT CASE action
	
	CASE "send"
			SendLinkToAppViaEmail			
	CASE ELSE
			DisplayEnterAppLinkScreen
END SELECT






' --- Writes the Closing tags for HTML ---
DisplayCloseBodyAndHTMLTags




' ******************************************************************************
' --- End of MAIN PROGRAM ---
' ******************************************************************************



' -----------------------------
  SUB DisplayNoUserSetScreen
' -----------------------------

%>
<div id="NoAuthorizedUserSet" class="tabrankings" style="width:99%; margin:3px 0px 0px 0px; padding:0px 0px 0px 0px; height:465px; display:none;">
	<div style="margin:20px 10px 0px 10px; text-align:center;">
		<span class="span90" style="color:yellow; text-align:center;"><b>This function may not be accessed until an Authorized User has been set up on this device.</b></span>
	</div>
	<div class="span95" style="margin-top:50px; text-align:center;">
		<form action="/rankings/<%= MenuFileName %>" method="post">
			<input type="submit" value="Return to Main Menu" title="Go to Main Menu" style="font-size:12pt; size:12em;">
		</form>
	</div>
</div>	
<%

END SUB



' -----------------------------
  SUB DisplayEnterAppLinkScreen
' -----------------------------


'LinkEmail = "mark@productdesign-biz.com"



%>
<div class="errorbox" id="SendAppLinkEntryScreen" style="width:99%; margin:3px 0px 0px 0px; padding:0px 0px 0px 0px; height:465px;">
	<form action="<%=ThisFileName%>?action=send" method="post">
		<input type="hidden" id="sMemberID" name="sMemberID" value="">
		<input type="hidden" id="sMembEmail" name="sMembEmail" value="">	
		<div style="margin-top:10px; border:0px solid white;">
  		<span class="span95" style="text-align:center; color:yellow; font-size:14pt">Send Application Link to Friend</span>
		</div>
		<div style="width:96%; margin:20px 10px 0px 0px; padding:0px 0px 0px 10px; color:white;">Enter Email address to forward a link to access the AWSA application.</div>	
		<div style="width:96%; margin-top:20px; padding-left:10px;">		
			<span class="span95" style="text-align:center; color:#FFFFFF; font-size:12pt">Email Address</span> 
			<span class="span95" style="text-align:center;">
				<input type="text" name="LinkEmail" id="LinkEmail" value="<%=LinkEmail%>" style="font-size:12pt;" size="29" maxlength="50">
			</span>
		</div>
		<div style="width:96%; margin-top:20px; padding-left:10px; font-size:14pt;">OR</div>	
		<div style="width:96%; margin-top:20px; padding-left:10px;">		
			<span class="span95" style="text-align:center; color:#FFFFFF; font-size:12pt">Mobile Phone Number(N/A)</span> 
			<span class="span95" style="text-align:center;">
				<input type="Tel" name="SMS_PhoneNumber" id="SMS_PhoneNumber" value="" size="10" style="font-size:12pt;" maxlength="10" disabled> 
			</span>
		</div>
		
		<div style="height:50px; margin:50px 0px 0px 0px; padding:0px 0px 0px 0px; border:0px solid white; border:0px solid red;">
			<span class="span45" style="margin:0px; padding:0px; text-align:center; border:0px solid white;">
				<input type="submit" name="SendEmail" value="Send Email"  style="width:9em; height:2em; font-size:12pt;">
			</span>
		</div>
	</form>	
</div>  
<%	


END SUB




' -----------------------------
  SUB DisplaySentAppLinkScreen
' -----------------------------

SMS_PhoneNumber="407-383-0921" 
' LinkEmail="mark@productdesign-biz.com"
SMS_Message="Here is the link to the AWSA mobile App http://www.usawaterski.org/rankings/mainmenu_m.asp"

%>
<div class="errorbox" id="SendAppLinkEntryScreen" style="width:99%; margin:3px 0px 0px 0px; padding:0px 0px 0px 0px; height:465px">
		<div style="margin-top:10px; border:0px solid white;">
  		<span class="span95" style="text-align:center; color:yellow; font-size:14pt">Send Application Link via Text Message to the number below.  To complete click on link</span>
		</div>
		<%
		
IF TRIM(SMS_PhoneNumber)<>"" AND LEN(SMS_PhoneNumber)=12 THEN 

'response.write("</div><div style=color:red;>HERE</div>
'response.end

	%>
	<div style="color:white; margin:20px 0px 0px 0px;">
		<a href="sms:<%=SMS_PhoneNumber%>;body=<%=SMS_Message%>" style="font-size:12pt; color:white; text-decoration:none;">Click to Send to <%=SMS_PhoneNumber%></a>
	</div>
	<%
END IF
%>
</div>	
<%

END SUB





' ---------------------------------
  SUB SendLinkToAppViaEmail
' ---------------------------------


sMemberID = Request("sMemberID")

' --- Test if a record already exists --
sSQL = "SELECT FirstName, LastName FROM "&MemberShortTableName
sSQL = sSQL + " WHERE PersonID = '"&RIGHT(sMemberID,8)&"'"

SET rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable


' --- Update or Add dpending on existence --
IF NOT rs.eof THEN
		FirstName = rs("FirstName")
		LastName = rs("LastName")
		FullName = FirstName&" "&LastName
END IF

rs.close





' --- Test if a record already exists --
sSQL = "SELECT COALESCE(Num_Forwards,0) AS Num_Forwards FROM "&MobileAppUserTable
sSQL = sSQL + " WHERE MemberID = '"&sMemberID&"'"

SET rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable


' --- Update or Add dpending on existence --
IF NOT rs.eof THEN
		Num_Forwards = rs("Num_Forwards")
		New_Num_Forwards = Num_Forwards + 1

		sSQL = "UPDATE "&MobileAppUserTable
		sSQL = sSQL + " SET Num_Forwards="&New_Num_Forwards
		sSQL = sSQL + " WHERE MemberID = '"&sMemberID&"'"		
END IF

rs.close

'response.write("</div><div style=color:red>"&sSQL&"</div>")
'response.end
OpenCon
con.execute(sSQL)
closecon





' --- CSS for styling Email --
SQT = "'"
ecss = "<style type=text/css>"
ecss = ecss & " body { font-family: Arial, Helvetica, sans-serif; text-align:center;}"
ecss = ecss & " .outer {color:white; font-size:14pt; background-color:#FFFFFF; text-align:center; min-width:320px; max-width:500px; height:500px; border:1px solid;}"
ecss = ecss & " p {color:black; font-size:12pt; text-align:left; font-style:normal; position:relative;}"
ecss = ecss & " .pblue {color:blue; font-size:12pt; text-align:left;}"
ecss = ecss & " .pblack {color:#000000; font-size:12pt; text-align:left;}"
ecss = ecss & " .actionbutton {background-color:#006400; color:white; -moz-border-radius:15px; -webkit-border-radius:15px; border:5px solid; padding:5px;}"
ecss = ecss & " .psuedobuttoncellgreen {width:175px; text-align:center; -webkit-border-radius: 3px; -moz-border-radius: 3px; border-radius: 3px; background-color:#006400;}"
ecss = ecss & " .psuedobuttoncellred {width:175px; text-align:center; -webkit-border-radius: 3px; -moz-border-radius: 3px; border-radius: 3px; background-color:#DC143C;}"
ecss = ecss & " .psuedobuttongreen {width:100%; font-size:16pt; font-family:Helvetica, Arial, sans-serif; color:#ffffff; text-decoration:none; color:#ffffff; text-decoration:none; -webkit-border-radius:3px; -moz-border-radius:3px; border-radius:3px; padding:12px 0px; border: 1px solid #7FFF00; display: inline-block;}"
ecss = ecss & " .psuedobuttonred {width:100%; font-size:16pt; font-family:Helvetica, Arial, sans-serif; color:#ffffff; text-decoration:none; color:#ffffff; text-decoration:none; -webkit-border-radius:3px; -moz-border-radius:3px; border-radius:3px; padding:12px 0px; border: 1px solid #FFA500; display: inline-block;}"
ecss = ecss & " </style>"


' --- Data for Invite Email ---

sInviteBannerText = "Link to AWSA Mobile App"

eMailFrom = TRIM(Request("sMembEmail"))
LinkEmail = TRIM(Request("LinkEmail"))



'eMailTo = "mark@productdesign-biz.com"
eMailTo = LinkEmail
eMailCC = ""
eMailBCC = ""

' -- Changed to usawaterski.org on 12/15/2016 when the email problem started occuring ---
' -- IF eMailFrom = "" THEN eMailFrom = "competition@usawaterski.org"
eMailFrom = "competition@usawaterski.org"


eMailSubj = "New AWSA Mobile App"

Dim AWSA_Logo
AWSA_Logo = "AWSA_Oval_BlueSquare_197x83.png"



' --- Create Email message ---
ebody = ecss & "<html>"
ebody = ebody & "<head>"
ebody = ebody & "<title>Link to AWSA Mobile App</title>"
ebody = ebody & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
ebody = ebody & "</head>"
ebody = ebody & "<body bgcolor=""#FFFFFF"" text=""#000000"">"

ebody = ebody & "<div class=outer style=""margin:0px 0px 0px 0px; padding:0px 0px 0px 0px;"">"

ebody = ebody & "<div style="&SQT&"text-align:center;"&SQT&">"
ebody = ebody & "<img style='width:180px;' name=BannerLogo src='http://usawaterski.org/rankings/images/logos/"&AWSA_Logo&"' alt=AWSA Logo>"
ebody = ebody & "</div>"
ebody = ebody & "<div style="&SQT&"margin:20px 0px 10px 0px; text-align:center; font-size:24px; color:red;"&SQT&"><i>"&FullName&" Wants You To Try This App</i></div>"

ebody = ebody & "<div class=pblack style=""margin:20px 10px 0px 0px; text-align:center;"">"
ebody = ebody & "  Click on the mobile phone image below to download the new mobile app from the <b>American Water Ski Association</b>. You can access tournament listings, scores, rankings, NOPS calculator, rule book and much more."
ebody = ebody & "</div>"

ebody = ebody & "<div class=pblack style=""margin:20px 10px 0px 0px; text-align:center;"">"
ebody = ebody & "  It's easy to put the App icon onto the home screen of your mobile device to make access just one click away."
ebody = ebody & "</div>"

ebody = ebody & "<div class=pblack style=""text-align:center; margin:15px 10px 0px 0px; padding-bottom;30px"">Click the image below from your phone to try out the new AWSA mobile App."
ebody = ebody & "<br><br>" 
ebody = ebody & " <a href='http://usawaterski.org/rankings/"&MenuFileName&"' style=""text-decoration:none;"">"
ebody = ebody & "  <img style=""width:200px;"" name=""MobileAppIcon"" src=""http://www.usawaterski.org/rankings/images/Mobile/iPhone_MyStats.png"" alt=""Mobile App"">"
ebody = ebody & " </a>"
ebody = ebody & "</div>"

ebody = ebody & "</div>"
ebody = ebody & "<br><br><br>"
ebody = ebody & "</body></html>"

eMailBody = ebody



'http://www.usawaterski.org/rankings/images/icons/AWSA_HomeScreen_57.PNG



'response.write("</div><br>eMailTo = "&eMailTo&"<br>eMailCC = "&eMailCC&"<br>eMailBCC = "&eMailBCC&"<br>eMailFrom = "&eMailFrom&"")
'response.write("<br>eMailSubj = "&eMailSubj&"<br><br>"&eMailBody)
'response.end

Dim EmailValidTest, EmailValidErrorMessage, EmailNotice

'response.write("</div><div style:color:red; font-size:16pt;>LinkEmail = "&LinkEmail&"</div>")
'response.write("<div style:color:red; font-size:16pt;>EmailValidTest = "&EmailValidTest&"</div>")
'response.end


' --- Tests for simple validity of email contruction ---
EmailValidTest = "N"
IF LEN(LinkEmail)>=10 AND Instr(LinkEmail,"@")>0 AND Instr(LinkEmail,".")>0 THEN EmailValidTest = "Y"

'response.write("</div><div style:color:red; font-size:16pt;>LinkEmail = "&LinkEmail&"</div>")
'response.write("<div style:color:red; font-size:16pt;>EmailValidTest = "&EmailValidTest&"</div>")
'response.write("<div style:color:red;>Instr(LinkEmail,@)=1 : "&Instr(LinkEmail,"@")>0&"</div>")
'response.write("<div style:color:red;>Instr(LinkEmail,.)>0 : "&Instr(LinkEmail,".")>0&"</div>")


EmailValidErrorMessage = "Email Address Not Valid"
EmailNotice = "The Email Address You Provided Was:"

IF EmailValidTest="Y" THEN 
		SendEmailFromGenericMethod eMailTo,eMailCC,eMailBCC,eMailFrom,eMailSubj,eMailBody
		EmailValidErrorMessage = "Link to App Has Been Sent"
		EmailNotice = "The Link to the AWSA Mobile App has been sent to the following email address:"
END IF


'  -- style="inline-block; margin-top:5px; height:440px;"

tiy=1
IF tiy=2 THEN
%>
<div class="errorbox" id="EmailSentConfirmationScreen" style="border:0px solid white; width:99%; margin:3px 0px 0px 0px; padding:0px 0px 0px 0px; height:465px;">
	<form action="<%=MenuFileName%>" method="post">
		<input type="hidden" id="sMemberID" name="sMemberID" value="">	
		<input type="hidden" id="sMembEmail" name="sMembEmail" value="">
		<div style="margin-top:10px; border:0px solid white;">
  		<span class="span95" style="text-align:center; color:#FFFFFF; font-size:14pt; padding-left:10px;"><%= EmailValidErrorMessage %></span>
		</div>
		
		<div style="width:96%; margin:20px 10px 0px 0px; color:yellow; padding-left:10px;"><%= EmailNotice %></div>	
		<div style="width:96%; margin-top:20px; color:white; padding-left:10px;"><%= eMailTo %></div>	
		<div id="ConfirmEmailSentButtons" class="menucell" style="padding:0px; margin:80px 0px 0px 0px;">
			<input type="submit" name="MainMenu" value="Main Menu" style="width:8em; height:2em; font-size:12pt;">
		</div>
	</form>	
</div>  
<%	
END IF ' -- Skip for testing 

%>
<div class="errorbox" id="EmailSentConfirmationScreen" style="width:99%; margin:3px 0px 0px 0px; padding:0px 0px 0px 0px; height:465px;">
	<form action="<%=MenuFileName%>" method="post">
		<input type="hidden" id="sMemberID" name="sMemberID" value="">	
		<input type="hidden" id="sMembEmail" name="sMembEmail" value="">
		<div style="margin-top:10px; border:0px solid red;">
  		<span class="span95" style="text-align:center; color:#FFFFFF; font-size:14pt; padding-left:10px;"><%= EmailValidErrorMessage %></span>
		</div>
		
		<div style="width:96%; margin:20px 10px 0px 0px; color:yellow; padding-left:10px;"><%= EmailNotice %></div>	
		<div style="width:96%; margin-top:20px; color:white; padding-left:10px;"><%= eMailTo %></div>	
		<div id="ConfirmEmailSentButtons" class="menucell" style="padding:0px; margin:80px 0px 0px 0px;">
			<input type="submit" name="MainMenu" value="Main Menu" style="width:8em; height:2em; font-size:12pt;">
		</div>
	</form>	
</div>  
<%


END SUB







' -----------------------------------------
  SUB HeresJavascriptCodeForDeviceTesting
' -----------------------------------------


'var ua = navigator.userAgent.toLowerCase();
'var url;
        
'if (ua.indexOf("iphone") > -1 || ua.indexOf("ipad") > -1)
'    url = "sms:;body=" + encodeURIComponent("I'm at " + mapUrl + " @ " + pos.Address);
'else
'    url = "sms:?body=" + encodeURIComponent("I'm at " + mapUrl + " @ " + pos.Address);

'location.href = url;


END SUB

%>
