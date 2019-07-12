<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_leagues.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->
<!--#include virtual="/rankings/tools_registration.asp"-->
<!--#include virtual="/rankings/Tools_TournamentListQuery.asp"-->
<!--#include virtual="/rankings/tools_mobile_version.asp"-->
<%



' ------------------------------------------------
' --- Dimensions variables used in this module ---
' ------------------------------------------------

Dim ThisFileName



Dim TeamMemberStatusText, TeamMemberStatusTextColor, TeamStatusText, TeamStatusTextColor
Dim TabColor
Dim MenuItemPath, FAQImagePath
Dim MyStatsFilename, LaunchFilename, AddIconNowInstruction
Dim action

AddIconNowInstruction=""
IF Request("action") = "applehomeicon" THEN AddIconNowInstruction="iPhone users press the Bookmark icon in the center of the tray below to begin Add to Home Screen process"
		

MenuItemPath = "images\icons\"
FAQImagePath = "images\mobile_faq\"
ThisFileName = "mainmenu_m.asp"


' --- Names related programs for linking ---
SearchFileName = "search-memberHQ.asp"
RankingsMobileFilename="view-standings_m.asp"
TournamentsMobileFilename="view-tournaments_m.asp"
TeamsMobileRankingFilename="View-vteamstatus_m.asp"
LocalVarFileName="User_Set.asp"
MenuFileName = "mainmenu_m.asp"
MyStatsFilename = "view-mystats_m.asp"
LaunchFilename = "awsa_launch.asp"




' --- Displays the html, head and opening body tag ---
OpenState="launch"
DisplayHeadOpenBodyAndBannerTags OpenState




DisplayMenu

DisplayiPhoneAddIcon


'' --- Writes the Closing tags for HTML ---
DisplayCloseBodyAndHTMLTags






' ---------------------------------------------------
' --- BOTTOM OF MAIN CODE ---
' ---------------------------------------------------







' ------------------
  SUB DisplayMenu
' ------------------ 

AddICon = MenuItemPath&"Down_Green_Arrow.png"			' --- Home Screen down arrow


%>
<input type="hidden" id="sMemberID_Hidden_InRankingsSettings" name="sMemberID" value="">
<div id="mainlaunch" style="width:100%; display:inline;">
	<div style="text-align:center; margin-top:10px;">
  	<img style="width:300px;" name="BannerLogo" src="http://usawaterski.com/rankings/images/General/JoinTheFun.JPG" alt="Accept_Join">
  </div>
	<div style="height:40px; text-align:center; display:inline-block; margin-top:30px; padding-left:100px;">
		<a href="<%=ThisSitePath%>/<%=MenuFilename%>" title="MainMenu" style="text-decoration:none;">
			<input type=button align="center" class="buttonblue" style="font-size:12pt;" value="Continue" name="Continue">
  	</a>
	</div>	
	<div style="width:96%; margin-top:10px; padding-left:10px; text-align:center; border:0px solid red; position:inline-block;">		
		<span id="" class="span15" style="border:0px solid; border-color:white; vertical-align:top;">
			<img src="images/buttons/Button-Info-icon.png" style="padding:0px; width:30px; text-align:right;" alt="Tip" onclick="javascript:LaunchPageOptions('iphone');">
		</span>
		<span id="" class="span80" style="border:0px solid; border-color:white; color:white; text-align:left; margin-top:5px;">
			Setting Icon On Your Home Screen
		</span> 
	</div>
	<div style="width:96%; text-align:center; margin-top:10px; padding-left:10px; border:0px solid red; position:inline-block;">		
  	<span class="span95" style="text-align:center; padding:0px; margin:0px; width:100%;">
			<textarea id="AddIconNowInstruction" name="AddIconNowInstruction" style="width:100%; font-size:12pt; text-align:center; color:yellow; background-color:black; border:0px solid;" cols=29 rows=3 wrap=physical><%=AddIconNowInstruction%></textarea>
  	</span> 	
	</div>

</div> <! -- Overall Mainmenu div -- ->
<%

END SUB





' --------------------------
  SUB DisplayiPhoneAddIcon
' --------------------------


FAQImage_A1 = FAQImagePath&"IPhone_Step1.PNG"
FAQImage_B1 = FAQImagePath&"IPhone_Step2.PNG"
FAQImage_A2 = FAQImagePath&"IPhone_Step3.PNG"
FAQImage_B2 = FAQImagePath&"IPhone_Step4.PNG"

%>
<div id="iPhoneAddIcon" style="display:none; margin-top:5px; text-align:center;">
	<div style="width:96%; margin-top:10px; padding-left:10px; text-align:left; border:0px solid red;">		
			<span class="span95" style="margin-left:0px; padding-left:0px; text-align:center; font-size:14px; color:yellow; border:0px solid white;">Adding App Icon to iPhone Home Screen</span> 
	</div>	
	<div class="scroll" style="margin-top:5px; padding:0px; margin-left:0px; height:400px; border:0px solid white;">
	<div style="width:96%; margin-top:10px; padding-left:10px; text-align:left; border:0px solid red;">		
			<span class="span95" style="margin-top:10px; margin-left:0px; padding-left:0px; text-align:center; font-size:12px; color:white; border:0px solid white;">To create the AWSA App icon on your Home Screen you <b>must begin at the Launch Page</b> which may be accessed from the button at the bottom of this instruction. Performing these simple steps will permit you to access the App with a single click.<br><br>Scroll on the images below to continue.</span> 
	</div>
		<div style="margin-top:10px;">
			<span class="span95"><img src="<%=FAQImage_A1%>" style="width:250px;" title='iPhone-Step1'></span>
			<span class="span95" style="margin-top:10px; margin-left:0px; padding-left:0px; text-align:center; font-size:12px; color:white; border:0px solid white;">STEP 1 - From the Launch Page (not now), first press the Bookmark icon at the bottom center of the screen. If the bookmark icon is not visible, touch near the top of the screen and slide it downward slightly.</span> 			
			<span class="span95"><img src="<%=FAQImage_B1%>" style="width:250px; margin-top:15px;" title='iPhone-Step2'></span>
			<span class="span95" style="margin-top:10px; margin-left:0px; padding-left:0px; text-align:center; font-size:12px; color:white; border:0px solid white;">STEP 2 - Next, locate and press the Add to Home Screen icon in the lower carousel.</span> 			
			<span class="span95"><img src="<%=FAQImage_A2%>" style="width:250px; margin-top:15px;" title='iPhone-Step3'></span>
			<span class="span95" style="margin-top:10px; margin-left:0px; padding-left:0px; text-align:center; font-size:12px; color:white; border:0px solid white;">STEP 3 - The last step is to press ADD in the upper right corner of the screen.</span> 			
			<span class="span95"><img src="<%=FAQImage_B2%>" style="width:250px; margin-top:15px;" title='iPhone-Step4'></span>
		</div>
		<div class="span95" style="height:50px; margin-top:20px; padding-bottom:50px;">
			<input type="button" class="buttonblue" name="Return To Launch Screen" value="Return To Launch Screen" style="width:16em; height:2em; font-size:12pt;" onclick="javascript:LaunchPageOptions('mainlaunch');">
		</div>
	</div>	
</div>
<%

END SUB  


%>

