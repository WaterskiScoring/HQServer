<%@ Language=VBScript %>
<HTML>
<%
'Make sure the necessary variables are set.  Will vary by sports group.
'Response.Write("SptsGrpID = " & Session("SptsGrpID") & "<br>RegnID = " & Session("RegnID"))
'Response.End
if Session("SptsGrpID") = ""  or Session("RegnID") = "" then
	Response.Write ("The page you requested will not display correctly unless you first make some selections on the start page. <br> Please use the link below to return to the start page and make your selections again.<P> <a href=" & "'" & "../../default.html" & "'" & "> Return to Start Page </a>")
	session.Abandon
	Response.End
end if
dim RequestMethod
'Capture the request method for this page.  If post then proceed.  If Get then redirect to default.html (beginning of system)
'Input Form always uses Post.  All other methods of accessing the script use Get.  
'Prevents access to this page by any method except form input.
sMode = "firstime" 'sets up form for first use or reuse.
RequestMethod = Request.ServerVariables("REQUEST_METHOD")
IF RequestMethod = "GET" THEN
	if request.QueryString("OK") <> "OK" then
		session.Abandon
		Response.Redirect "../../default.html"
		Response.End
	else
		sMode = "again"
	end if
end if

%>
<%	'Make sure user is logged in
	If Session("LoggedIn")= false then
		session.Abandon
		Response.Redirect ("../../default.html")
		Response.End
	end if
	'Make sure Region Level ID is correct for this page.
	if not Session("HQUser") = true then
		session.Abandon
		Response.Redirect ("../../default.html")
		Response.End
	end if
%>

<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<%
dim Conn3, sConn, SQL, rsS, sYear, sRegnID, REmail, HQEmail, LogoFileName, RHUserName, TestMonth, cMonth, cMonthColor, sGBPolicy, CalcIconAlt, CalcIconName, CalcMissing, CalcRegnStatus, SafeStart, SafeEnd, SafeTSite, SafeTName, Calc5Star, CalcFormStatus, CalcRFee, CalcAWSAFee, CalcSafety, s0, s1, s2, s3, s4, tempCalcIconName, ScoresPath, CF, CalcClubMemb, CalcHQStatus, sLogo, sHeader, sBGColor, sSptsGrpID

sSptsGrpID = Session("SptsGrpID")
Select case sSptsGrpID
	Case "AWS"
		sHeader = " for AWSA 3 Event"
		sLogo = "../../images/" & Session("RegnLogo") '"../../images/logo_awsa_sm.jpg"
		sBGColor = "Maroon"
	Case "ABC"
		sHeader = " for ABC Barefoot"
		sLogo = "../../images/logo_abc_" & lcase(session("RegnID")) & ".gif"
		sBGColor="Blue"
	Case "NCW"
		sHeader = " for Collegiate"
		sLogo = "../../images/logo_ncw_" & lcase(session("RegnID")) & ".jpg"
		sBGColor="Blue"
end select

sRegnID = Session("RegnID")
Session("TAID") = "" 'session variable holds user's choice of tournament to revise.  Start with none selected.
CF = "HQ"  'Passed when Tournament Advertisement is requested.  Limits size of recordset.
'set the existing variables for this form
REmail = session("RBossEmail")
LogoFileName = sLogo
sGBPolicy = Session("GBPolicy")
HQEmail = Session("HQBossEmail")
ScoresPath = Session("ScoresPath")

if sMode = "firstime" then
	sYear = request("Year")
	Session("Year") = sYear 
	'get the SQL parameters
	sckWhere1 = false
	if request("ckWhere1") = "on" then 
		sckWhere1 = true
		sWhere1 = " AND TDateE > '" & cdate(request("Where1")) & "' "
	else
		sWhere1 = ""
	end if
	sckWhere2 = false
	if request("ckWhere2") = "on" then 
		sckWhere2 = true
		sWhere2 = " AND TDateE < '" & cdate(request("Where2")) & "' "
	else
		sWhere2 = ""
	end if
	sTStatus = request("TStatus")
	
	Select case sTStatus
		case 0, 1, 2, 3, 4, 5
			sqlTStatus = " And TSTATUS = " & sTStatus & " "
		case 7
			sqlTStatus = " AND TSTATUS IN (2, 3, 4) "
		case else  '6 specifies all but use catchall instead
			sqlTStatus = ""
	end select

	SQLWhere = " Where TRegion = '" & sRegnID & "'" & " AND TYear = '" & sYear & "' " & sqlTStatus & sWhere1 & sWhere2
	SQL = "SELECT * from Tschedul " & SQLWhere & " Order by TDateE"
	session("SQL") = SQL
'	Response.Write("firsttime:  <br>" & SQL & "<br>Tstatus = " & request("TStatus"))
'	Response.End
else
	sYear = Session("Year")
	SQL = Session("SQL")
'	Response.Write("else:  <br>" & SQL & "<br>Tstatus = " & request("TStatus"))
'	Response.End
end if
%>

<!--Set up the page header.-->
<body>
<h2><img src="../../images/usawski.gif" width="150" height="39" alt="USA Water Ski logo" align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;USA Water Ski <br>&nbsp;&nbsp;&nbsp;&nbsp;Post Tournament Functions</h2>
<%if sSptsGrpID = "AWS" then
	RHUserName = "AWSA"%>
	<h3><table><tr><td><img src="../../images/logo_awsa_sm.jpg" WIDTH="104" HEIGHT="66"></td><td><img src="<%=sLogo%>"></td></tr>
<%end if
  if sSptsGrpID = "ABC" then
	RHUserName = "ABC"%>
	<h3><table><tr><td><img src="../../images/logo_abc_sm.jpg" WIDTH="141" HEIGHT="79"></td><td><img src="<%=sLogo%>"></td></tr>
	    
<% end if
if sSptsGrpID = "USW" then
	RHUserName = "USA Wakeboard"%>
	<h3><table><tr><td><img src="../../images/logo_usw_150.gif" WIDTH="75" HEIGHT="30"></td><td></td></tr>
	    
<% end if
if sSptsGrpID = "AKA" then
	RHUserName = "AKA"%>
	<h3><table><tr><td><img src="../../images/logo_aka_150.jpg" WIDTH="75" HEIGHT="75"></td><td></td></tr>
	    
<% end if%>
</table>
<table width="100%"><tr><td></td><td>
<font size="5" color="#0000ff"><center><%= sYear %> &nbsp;&nbsp;<%= RHUserName %> &nbsp;&nbsp;   Tournaments</center></font>
<center><P> <a href="Mailto:<%= REmail %>"> Regions Tournament Co-ordinator </a>
<br> To <b>view the advertisement</b> click on the Tournament Name.
<br></font><font size="3">To <b>Update Paperwork Status</b> - Click on the Tournament Status Icon. 
<br><a href="awsh.asp"><b>Select a different function</b></a></font></center></td></tr></table></h3>
<center>
<table BORDER="1" CELLPADDING="3" WIDTH="100%"><th><img src="../../images/status.jpg" alt="Status of Application or Tournament" WIDTH="21" HEIGHT="21"></th>
	<th>DATE(s)</th><th> ID# &nbsp;&nbsp;&nbsp;  TOURNAMENT NAME / <u>Advert</u></th><th>SITE</th><th>EVENTS OFFERED</th></tr>

<%
Set Conn3 = Server.CreateObject("ADODB.Connection")
Set rsS = server.CreateObject("ADODB.Recordset")
sConn = Application("PSAConnStr")
Conn3.Open sConn  
Set rsS = Conn3.Execute (SQL)
if rsS.EOF and rsS.BOF then
	Response.Write("<h1><font color = red>No Applications Found in Calendar Year " & sYear & "</font></h1>")
	Response.End
end if
TestMonth =  0 'Force Month header for first record
	do while not rsS.EOF
	
	Select Case month(rsS("TDateE"))
		Case 1
			cMonth = "January"
			cMonthColor = "#fafad2"
		case 2
			cMonth = "February"
			cMonthColor = "#f5deb3"
		case 3
			cMonth = "March"
			cMonthColor = "#F5f5f5"
		case 4
			cMonth = "April"
			cMonthColor = "#fff5ee"
		case 5
			cMonth = "May"
			cMonthColor = "#f5fffa"
		case 6
			cMonth = "June"
			cMonthColor = "#ffe4e1"
		case 7
			cMonth = "July"
			cMonthColor = "#f0ffff"
		case 8
			cMonth = "August"
			cMonthColor = "#f5f5dc"
		case 9
			cMonth = "September"
			cMonthColor = "#F0F8FF"
		case 10
			cMonth = "October"
			cMonthColor = "#faebd7"
		case 11
			cMonth = "November"
			cMonthColor = "#fff5ee"
		case 12
			cMonth = "December"
			cMonthColor = "#faf0e6"
	end select	

'Calculate Icon and alt values for Returning Tournament Paperwork
'Calculate Form Status (for returning paperwork from scored tournaments)
	if rsS("Scored0") = true then   
		s0 = "<img src = " & chr(34) & "../../images/score0g.jpg" & chr(34) & ">"  
	else   
		s0 = "<img src = " & "'" & "../../images/score0r.jpg" & "'" & "alt=" & chr(34) & "Missing Scorebook" & chr(34) & ">"
	end if
	if rsS("Scored1") = true then   
		s1 = "<img src = " & "'" & "../../images/score1g.jpg" & "'" & ">"  
	else   
		s1 = "<img src = " & "'" & "../../images/score1r.jpg" & "'" & "alt=" & chr(34) & "Missing Chief Judges Report" & chr(34) & ">"
	end if
	if rsS("Scored2") = true then   
		s2 = "<img src = " & "'" & "../../images/score2g.jpg" & "'" & ">"  
	else   
		s2 = "<img src = " & "'" & "../../images/score2r.jpg" & "'" & "alt=" & chr(34) & "Missing Safety Report" & chr(34) & ">"
	end if
	if rsS("Scored3") = true then   
		s3 = "<img src = " & "'" & "../../images/score3g.jpg" & "'" & ">"  
	else   
		s3 = "<img src = " & "'" & "../../images/score3r.jpg" & "'" & "alt=" & chr(34) & "Missing Chief Drivers Report" & chr(34) & ">"
	end if
	if rsS("Scored4") = true then   
		s4 = "<img src = " & "'" & "../../images/score4g.jpg" & "'" & ">"  
	else   
		s4 = "<img src = " & "'" & "../../images/score4r.jpg" & "'" & "alt=" & chr(34) & "Missing Officials Work Record" & chr(34) & ">"
	end if
	CalcFormStatus = s0 & s1 & s2 & s3 & s4
'
	'Build string of what is missing if anything
	if Len(CalcFormStatus) < 7 then 
		CalcMissing = "All application materials received" 
	else 
		CalcMissing = CalcFormStatus
	end if
'
'Calculate the Icon and Alt Values for the STATUS column
'See what checkoffs are missing.

	'See if Region Fee has been paid
	if rsS("TKitOKRegnFeePd") = true then 
		CalcRFee = "" 
	else 
		CalcRFee = " Region Fee "
	end if
'
	'See if AWSA fee has been paid
	CalcAWSAFee = ""
	if rsS("chkHQOnly4") = false then 
		CalcAWSAFee = " AWSA Fee "
	end if
'
	'See if Guidebook is approved
	if rsS("TKitOKGuidebookAd") = true then 
		CalcGB = "" 
	else 
		CalcGB = " Guidebook Ad "
	end if
'
'See if club membership is checked off by HQ
	CalcClubMemb = ""
	if rsS("chkHQOnly2") = false then
		CalcClubMemb = " Club Membership "
	end if
'See if Safety Form has been checked in.
	CalcSafety = ""
	if rsS("chkHQOnly3") = false then
		CalcSafety = "Safety Form"	
	end if
'See if Officials are approved
	CalcOffl = ""
	CalcROffl = ""
	CalcRegionOK = ""
	if rsS("chkRegionOK") = false then 'not yet region approved
		CalcRegionOK = " Region Approval "
	end if
	if rsS("TKitOKRegnOfficials") = false then
		CalcROffl = " Region Officials Approval "
	end if

	if rsS("ChkHQOnly1") = false then 
		CalcOffl = " HQ Officials Approval "
	end if

'See if Pan Am Sanction is needed and has been approved
	CalcPanAmStatus = ""
	if rsS("THSClassL") = true or rsS("THSClassR") = true or rsS("THTClassL") = true or rsS("THTClassR") = true or rsS("THJClassL") = true or rsS("THJClassR") = true then
		if rsS("PASanApproved") = false then
			CalcPanAmStatus = " PanAm Approval "
		end if
	end if	


	'build the string of missing checkoffs
	CalcRegnStatus = ""
	if len(CalcRFee & CalcGB & CalcROff & CalcRegionOK) > 6 then
		CalcRegnStatus = ", Missing " & CalcRFee & CalcGB & CalcROffl & CalcRegionOK
	end if		
'Calculate icon name for Status 0 based on region preference.
	if  sGBPolicy =  true  then 'Require only guidebook OK  for guidebook ad.
		tempCalcIconName = "status" & cStr(rsS("TSTATUS")) 
		if rsS("TSTATUS") = 0 and rsS("TKitOKGuidebookAd") = true then 
			tempCalcIconName = tempCalcIconName & "a"
		end if
			tempCalcIconName = tempCalcIconName & ".jpg"
	else 'Require Region OK for guidebook ad
		tempCalcIconName = "status" & cStr(rsS("TSTATUS")) & ".jpg"
	end if
	CalcIconName = tempCalcIconName
'	
	if len(CalcOffl & CalcClubMemb & CalcAWSAFee & CalcSafety & CalcPanAmStatus) > 6 then
		if len(CalcRegnStatus) > 6 then
			CalcHQStatus = ", " & CalcOffl & CalcClubMemb & CalcAWSAFee & CalcSafety & CalcPanAmStatus
		else
			CalcHQStatus = ", Missing " & CalcOffl & CalcClubMemb & CalcAWSAFee & CalcSafety & CalcPanAmStatus
		end if
	end if
 ' alt text to display if graphic is not available
	if CalcIconName = "status0a.jpg" then 
		CalcIconAlt = "Application Received, Guidebook Ad Approved" & CalcRegnStatus & CalcSafety
	else 
		Select Case rsS("TSTATUS")
			Case 0 'TStatus = 0
				CalcIconAlt = "Application Received" & CalcRegnStatus & CalcHQStatus
			Case 1 'TStatus = 1
				CalcIconAlt = "Approved by Region" & CalcRegnStatus & CalcHQStatus
			Case 2 'TStatus = 2
				CalcIconAlt = "Sanctioned by USAWaterski"
			Case 3 'TStatus = 3
				CalcIconAlt = "Tournament Canceled"
			Case 4 'TStatus = 4
				CalcIconAlt = "Competition Complete, Paperwork Pending"
			Case 5 'TStatus = 5
				CalcIconAlt = "Archived - All Processing Complete"
		end select
	end if

'Test Start Date
	if IsNull (rsS("TDateS")) then 
		SafeStart = "TBA"
	else 
		SafeStart = cstr(rsS("TDateS"))
	end if

' Test End date
	if IsNull (rsS("TDateE")) then 
		SafeEnd = "TBA"
	else 
		SafeEnd = cstr(rsS("TDateE"))
	end if
	if IsNull (rsS("TName")) then
		SafeTName = "TBA"
	else 
		SafeTName = rsS("TName")
	end if

	if IsNull (rsS("TSite")) then
		SafeTSite = "TBA"
	else 
		SafeTSite = rsS("TSite")
	end if

		sIcos = ""
	if rsS("T5Star") = true then sIcos = sIcos & "<img src=""../../images/5star.jpg"" WIDTH=""31"" HEIGHT=""34"" align=""center"">"
	if rsS("TEventNSL") = true then sIcos = sIcos & "<img src=""../../images/ico_nsl.jpg"" WIDTH=""31"" HEIGHT=""34"" align=""center"">"
	if rsS("TEventNBL") = true then sIcos = sIcos & "<img src=""../../images/ico_nbl.gif"" WIDTH=""31"" HEIGHT=""34"" align=""center"">"
	if rsS("TEventNWL") = true then sIcos = sIcos & "<img src=""../../images/ico_nwl.gif"" WIDTH=""31"" HEIGHT=""34"" align=""center"">"
	
if rsS("TEventSlalom") = False and rsS("TEventTrick") = False and rsS("TEventJump") = False and rsS("TEventFun") = True then
	strDescription = rsS("FDescription")
elseif rsS("TEventSlalom") = true or rsS("TEventTrick") = true or rsS("TEventJump") = true then
	strDescription = rsS("TDescription")
	if rsS("TEventFun") = true then
		strDescription = strDescription & "<br>" & rsS("FDescription")
	end if
End if
%>

		 
		<%	if month(rsS("TDateE")) <> TestMonth then	
		
				TestMonth = month(rsS("TDateE")) %>
				<TR><TD HALIGN="LEFT" COLSPAN="5" BGCOLOR="<%= cMonthColor %>"><%= cMonth %></TD></TR>					
			<%end if%>
			
		<tr><td><a href="awsh_posttourn.asp?mTAID=<%= rsS("TournAppID")%>&TS=<%=rsS("TSTATUS")%>"><img src="../../images/<%= CalcIconName %> " alt = "<%= CalcIconAlt %>"></a></td>

		<TD bgcolor="<%= cMonthColor %>" ><font size="2"> <%= SafeStart %>, <%= rsS("TDateE") %> </font></TD>
			
		<TD><font size="2"><%= rsS("TournAppID")%>  <a href="../advert.asp?TAID=<%= rsS("TournAppID")%>&R2=<%= CF%>"> <%= SafeTName %> </a></font></TD>
			
		<td><font size="2"> <%= SafeTSite %>,  <%= rsS("TState") %> </font></td>
			
		<td><font size="2"><center> 
			<% if rsS("TSTATUS") = 4 then %>
				<%= CalcMissing %>
			<% else %>
				<%= strDescription%>  <%= sIcos%></center></font></td></tr>
			<%end if%>



<% rsS.movenext

loop 

rsS.close
Conn3.close
%>
<tr><td colspan = 5><a href="../../default.html"> Log Out </a></td></tr>
</table>
</BODY>
</HTML>