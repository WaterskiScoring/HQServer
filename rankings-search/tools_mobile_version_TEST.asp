<%
' ***********************************************************************************************
' ***********************************************************************************************
' --- This file contains common modules that can be used by multiple segments of the programs ---
' --- Created: Mark Crone  5/24/2015 

' --- TESTING ---
' ---  Purpose is to allow change to include statements common to all modules -- 

' --- Modification dates:
' ---   ver: 4-18-2016 - Test new dedection of Mobile iOS before triggering install popup --
' ---   ver: 

' ***********************************************************************************************
' ***********************************************************************************************



Dim SearchFileName, RankingsMobileFilename, TournamentsMobileFilename, TeamsMobileFilename
Dim SendLinkFileName
Dim MenuFileName, LocalVarFileName, MyStatsFilename, NOPSCalcFilename


Dim ThisSitePath

ThisSitePath = "/rankings"






' --------------------------------------------------
  SUB DisplayHeadOpenBodyAndBannerTags (OpenState)
' --------------------------------------------------  
  
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta charset="utf-8">
<title>AWSA Mobile Rankings Page Test</title>
<link rel="stylesheet" href="css/stylesheet_mob_tours.css" media="screen">
<meta charset="utf-8"> 		
<meta name="apple-touch-fullscreen" content="yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="viewport" content="width=device-width, height=device-height, minimum-scale=1, maximum-scale=1, user-scalable=no, minimal-ui">
<meta name="apple-mobile-web-app-status-bar-style" content="black">
<meta name="apple-mobile-web-app-title" content="AWSA Mob">
<meta name="format-detection" content="telephone=no">
<link rel="apple-touch-icon" href="http://www.usawaterski.org/rankings/images/icons/AWSA_HomeScreen_57.png">
<! '--- For iPad --- ->
<link rel="apple-touch-icon" sizes="72x72" href="http://www.usawaterski.org/rankings/images/icons/AWSA_HomeScreen_57.png">
<! --- For pre-retina iPhone, iPod Touch, and Android 2.1+ devices --- ->
<link rel="apple-touch-icon" href="http://www.usawaterski.org/rankings/images/icons/AWSA_HomeScreen_57.png">
<script language="javascript" type="text/javascript" src="js/view-tours-mobile_TEST.js"></script>
<script language="javascript" type="text/JavaScript" src="/jscripts/scripts.js"></script>
<script language="javascript" type="text/javascript" src="/jscripts/swfobject.js"></script>
</head>
<%



' -- Determines menu based on a session variable that can only be set in the TEST mainmenu_m_test.asp --
MenuFileName="mainmenu_m.asp" 
IF TRIM(Session("TESTING"))<>"" THEN
		MenuFileName="mainmenu_m_TEST.asp"
END IF



' -- Sets logo displayed on each page --
' -- Default logo --
SponsorHeader = "PolkCounty_135.png"		

' response.write("</div></div> OpenState = "&OpenState)

SELECT CASE OpenState
		CASE "rankings"
				SponsorHeader = "Nautiques_135.png"
				%><body onload="javascript:UpdateMemberID_IntoHidden_FromStart(); javascript:InitialRankingsSettingsRecall();"><%
		CASE "tournaments"
				SponsorHeader = "Mastercraft_135.png"
				%><body onload="javascript:UpdateMemberID_IntoHidden_FromStart(); javascript:InitialTournamentSettingsRecall();"><%
		CASE "myteamlisting"
				SponsorHeader = "Indmar_135.png"
				%><body onload="javascript:UpdateMemberID_IntoHidden_FromStart();"><%
		CASE "setuser_enter"
				SponsorHeader = "Centurion_135.png"
				%><body onload="javascript:UpdateMemberID_IntoHidden_FromStart(); Javascript:CheckWatchersButtonStatus('managewatcherbutton'); javascript:SetUserNav('load');"><%
		CASE "setuser_find"
				SponsorHeader = "Centurion_135.png"
				%><body onload="javascript:UpdateMemberID_IntoHidden_FromStart()"><%
		CASE "manageteams"
				%><body onload="javascript:UpdateMemberID_IntoHidden_FromStart();"><%
		CASE "vteamrankings"
				%><body onload=""><%
				SponsorHeader = "Indmar_135.png"				
		CASE "mystats"
				SponsorHeader = "Malibu_135.png"
		CASE "sendlink"
				%><body onload="javascript:TestAuthorizedUserSet('sendlink')"><%
		CASE "mainmenu"
				SponsorHeader = "VisitCentralFlorida.png"
				%><body onload="javascript:UpdateMemberID_IntoHidden_FromStart(); javascript:UpdateWatchMemberID(); javascript:IsAppInstalled();"><%
		CASE "mainmenu_test"
				SponsorHeader = "VisitCentralFlorida.png"
				%><body onload="javascript:UpdateMemberID_IntoHidden_FromStart(); javascript:UpdateWatchMemberID(); javascript:IsAppInstalled();"><%

END SELECT		


' -- images/logos/usa-water-ski-logo-109x63.png
' -- images/logos/awsa.jpg


'-- Main Menu org logo --
Dim AWSA_Logo
AWSA_Logo = "AWSA_Oval_BlueSquare_197x83.png"
' -----------------------------
' --- Displays Banner Line --- 
' -----------------------------

%>
<div class="container" style="height:100%; border:0px solid red;">
	<a name="TopTop" title="Page Navigation"></a>
	<div id="bannerheader" style="width:100%; background-color:<%=HQSiteColor2%>; height:70px; margin:0px; padding:0px; border:0px solid white;">
		<a href='<%=ThisSitePath%>/<%=MenuFilename%>' title="Rankings" style="text-decoration:none;" >
			<span class="span45" style="width:45%; height:100%; border:0px solid white;">
				<img src="images/logos/<%=AWSA_Logo%>" style="height:57px; margin:7px 0px 0px 3px; padding:0px 0px 0px 0px; border:0px solid green;" alt="AWSA New Logo" />
			</span>
			<span class="span55" style="width:50%; height:100%; vertical-align:top; text-align:center; padding:0px 0px 0px 0px; margin:0px 0px 0px 0px; border:0px solid white;">
					<img src="images/logos/<%=SponsorHeader%>" style="width:135px; margin:13px 0px 0px 13px; border:1px solid green;" alt="Banner Ad" />
					<span class="span95" style="text-align:center; padding:5px 0px 0px 0px; margin:0px 0px 0px 8px; color:#FFFFFF; border:0px solid red;">Share Life On The Water</span>
			</span>
		</a>	
	</div>
<%


END SUB 






' -----------------------------------
  SUB DisplayCloseBodyAndHTMLTags
' -----------------------------------

%>
</div><! -- Container -- ->
</body>
</html>
<%

END SUB





' -----------------------------------------
	SUB DisplayMenuButtons_ViewTournaments_WithDivs_NOT
' -----------------------------------------	

' --- NOT USED ---
%>
<div id="tourmenubuttons2">
	<div style="padding-top:7px; padding-left:3px;" >
		<span class="span100" style="padding-bottom:2px; padding-top:2px; width:100%;">
			<a href="<%=TournamentsMobileFilename%>" style="text-decoration:none;">
				<img src="images/buttons/TournamentsButton.png" style="padding:0px; width:75px; text-align:right;" alt="Tournament Button" />
			</a>
			<a href="javascript:SelectFromTournamentMenu('2');" style="text-decoration:none;">
				<img src="images/buttons/ScoresButton.png" style="padding:0px; width:75px; text-align:right;" alt="Scores Button" />
			</a>
			<a href="<%=RankingsMobileFilename%>" style="text-decoration:none;">
				<img src="images/buttons/RankingsButton.png" style="padding:0px; width:75px; text-align:right;" alt="Rankings Button" />
			</a>
			<a href="javascript:SelectFromTournamentMenu('4');" style="text-decoration:none;">
				<img src="images/buttons/FAQButton.png" style="padding:0px; width:75px; text-align:right;" alt="FAQ Button" />
			</a>
		</span>	
	</div>
</div>

	<td width="5px" style="border:none; background-color:<%=HQSiteColor2%>; padding:0px; margin:0px;">&nbsp;</td>
			
			
<%

END SUB




' -----------------------------------------
	SUB DisplayMenuButtons_ViewTournaments
' -----------------------------------------	

'sMemberID = TRIM(Request.Cookies("sMemberID"))

fg=2
IF fg=1 THEN
		%><div style="color:red">HERE <%= sMemberID %></div><%
		response.end
END IF

%>
<div id="tourmenubuttons" class="menucell" style="padding:0px; margin:0px; background-color:<%=HQSiteColor2%>">
	<TABLE align=center width=100% style="padding:0px; margin:0px; border:0px solid; border-color:yellow;"">
		<tr>
			<td width="23%" height="30px" background="images/buttons/ButtonBlank.png" style="background-position:center center; background-repeat:no-repeat; border:0px solid; border-color:#FFFFFF; background-size:75px; margin:0px; padding:0px; text-align:center;">
				<font size="2" color="blue"><a href="<%=TournamentsMobileFilename%>" style="text-decoration:none;">EVENTS</a></font>
			</td>
			<td width="23%" background="images/buttons/ButtonBlank.png" style="background-position:center center; background-repeat:no-repeat; border:0px solid; border-color:#FFFFFF; background-size:75px; margin:0px; padding:0px; text-align:center;">
				<font size="2" color="blue"><a href="<%=TeamsMobileFilename%>" style="text-decoration:none;">TEAMS</a></font>
			</td>				
			<td width="23%" background="images/buttons/ButtonBlank.png" style="background-position:center center; background-repeat:no-repeat; border:0px solid; border-color:#FFFFFF; background-size:75px; margin:0px; padding:0px; text-align:center;">
				<font size="2" color="blue"><a href="<%=RankingsMobileFilename%>" style="text-decoration:none;">RANKINGS</a></font>
			</td>				
			<%
			' IF TRIM(Request.Cookies("sMemberID"))<>"" THEN
			IF TRIM(sMemberID)<>"" THEN
					%>
					<td width="23%" background="images/buttons/ButtonBlank.png" style="background-position:center center; background-repeat:no-repeat; border:0px solid; border-color:#FFFFFF; background-size:75px; margin:0px; padding:0px; text-align:center;">
						<font size="2" color="blue"><a href="<%=LocalVarFileName%>" style="text-decoration:none;">PROFILE</a></font>
					</td>				
					<%
			ELSE		
					%>
					<td width="23%" background="images/buttons/ButtonBlank.png" style="background-position:center center; background-repeat:no-repeat; border:0px solid; border-color:#FFFFFF; background-size:75px; margin:0px; padding:0px; text-align:center;">
						<font size="2" color="blue"><a href="<%=LocalVarFileName%>" style="text-decoration:none;">FAQ</a></font>
					</td>				
					<%
			END IF
			%>
		</tr>
	</TABLE>
</div>	
<div style="background-color:<%=HQSiteColor2%>; height:7px; margin:0px 0px 0px 0px; padding:0px 0px 0px 0px">&nbsp;</div>
<%

END SUB



' -------------------------------------
  SUB Get_localStorage_fromTools_m_ver
' -------------------------------------

' -- When function is called this reads localStorage value and saves to sMemberID --

' --- OBSOLETE ---

	%>
	<script type="text/javascript"> 

		function get_localStorage_fromTools_m_ver() {

				alert("BEFORE");
				// alert("BEFORE getElementById.value = " + document.getElementById("sMemberID").value);
				// document.getElementById("sMemberID").value=localStorage.getItem("sMemberID");	
				document.getElementById("sMemberID_InRankingsSettings").value=localStorage.getItem("sMemberID");	
				document.getElementById("sFirstName_InRankingsSettings").value=localStorage.getItem("sFirstName");	
				// alert('AFTER Statement - localStorage.getItem(sMemberID) = ' + localStorage.getItem("sMemberID"));
				//alert("AFTER");
				//alert('AFTER getElementById.value = ' + localStorage.getItem("sMemberID") + ' - ' + localStorage.getItem("sFirstName");
				alert('AFTER getElementById.value = ' + localStorage.getItem("sMemberID");
		}
		</script>
		<%

END SUB







%>
