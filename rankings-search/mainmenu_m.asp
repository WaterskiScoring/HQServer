<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_leagues.asp"-->
<!--#include virtual="/rankings/tools_include16.asp"-->
<!--#include virtual="/rankings/tools_registration16.asp"-->
<!--#include virtual="/rankings/Tools_TournamentListQuery.asp"-->
<!--#include virtual="/rankings/tools_mobile_version.asp"-->
<%



' ------------------------------------------------
' --- Dimensions variables used in this module ---
' ------------------------------------------------

Dim ThisFileName
Dim FAQ_RankingsFileName


Dim TeamMemberStatusText, TeamMemberStatusTextColor, TeamStatusText, TeamStatusTextColor
Dim TabColor
Dim MenuItemPath, FAQImagePath






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
NOPSCalcFilename = "NOPS_m.asp"
SendLinkFileName = "Mobile_SendAppLink.asp"
ZBSCheatFilename = "view-ZBS_CheatSheet_m.asp"
TimingChartFilename = "view-TimingCharts_m.asp"

' -- Rankings FAQ - Does not update with changes to FAQ for main site - Input from Dave Clark -- 
FAQ_RankingsFileName = "Mobile_faq_rankings.asp"







' --- Displays the html, head and opening body tag ---
OpenState="mainmenu"
DisplayHeadOpenBodyAndBannerTags OpenState



' --- Displays the menu for view tournaments --- 
' DisplayMenuButtons_ViewTournaments




DisplayTwitterFeed				' --- Initially hidden

DisplayFAQOptions					' --- Initially hidden

DisplayiPhoneAddIcon			' --- Initially hidden

DisplayResourcesOptions		' --- Initially hidden

DisplaySavingSearchSettings		' -- An FAQ Answer - Appropriate Location in MainMenu_m ?--

DisplayRealTime


' --- Primary Page to Display ---
DisplayMenu


' --- Writes the Closing tags for HTML ---
DisplayCloseBodyAndHTMLTags



' ---------------------------------------------------
' --- BOTTOM OF MAIN CODE ---
' ---------------------------------------------------










' ------------------
  SUB DisplayMenu
' ------------------ 

' --- Defines what Icons are showing ---
MenuIcon_A1 = MenuItemPath&"Events_Blk_57.jpg"							' --- World with Magnifying glass
MenuIcon_B1 = MenuItemPath&"NationalRank_Blk_57.jpg"				' --- Graph
MenuIcon_C1 = MenuItemPath&"CollegeRank_Blk_57.jpg"					' --- Graduation cap
MenuIcon_D1 = MenuItemPath&"VTeamRank_Blk_57.jpg"						' --- Family

MenuIcon_A2 = MenuItemPath&"MyTeams_Blk_57.jpg"							' --- 4 People in circle
MenuIcon_B2 = MenuItemPath&"MyStats_57.jpg"									' --- Green Graph with Yellow Arrow
MenuIcon_C2 = MenuItemPath&"Twitter_Blk_57.jpg"							' --- Twitter Feed --
MenuIcon_D2 = MenuItemPath&"FullSite_57.jpg"								' --- USA Waterski Logo to link to full site

MenuIcon_A3 = MenuItemPath&"SetUser_Blk_57.jpg"							' --- Settings wheel
MenuIcon_B3 = MenuItemPath&"FAQ_Blk_57.jpg"									' --- Question Mark
MenuIcon_C3 = MenuItemPath&"Resource_Blk_57.jpg"						' --- Blue circle with Help
MenuIcon_D3 = MenuItemPath&"SendAppToFriend_Blk_57.jpg"			' --- Blue circle with Help

MenuIcon_A4 = MenuItemPath&"RealTime_Blk_57.jpg"						' --- Real Time scores - Digital Timer
MenuIcon_B4 = MenuItemPath&"Down_Green_Arrow.png"						' --- Home Screen down arrow
MenuIcon_C4 = ""
MenuIcon_D4 = ""

 

%>
<div id="mainmenuscreen" style="width:100%; height:100%; display:inline-block; border:0px solid white;">
<input type="hidden" id="sName_InRankingsSettings" name="sName_InRankingsSettings" value="">
<! -- ROW 1 -- ->
<div style="width:100%; margin-top:10px; border:0px solid white;">
	<span class="menuicon" style="border:0px solid white; position:inline">
		<a href='<%=ThisSitePath%>/<%=TournamentsMobileFilename%>?df=yes' title='Tournaments' style="text-decoration:none;" >
  		<img class="menuimage" src="<%=MenuIcon_A1%>">
  	</a>
	</span>
	<span class="menuicon" style="border:0px solid white; position:inline">
		<a href='<%=ThisSitePath%>/<%=RankingsMobileFilename%>?RankingListType=National&df=yes' title='Rankings' style="text-decoration:none;" >
  		<img class="menuimage" src="<%=MenuIcon_B1%>">
  	</a>
	</span>
	<span class="menuicon" style="border:0px solid white; position:inline">
		<a href='<%=ThisSitePath%>/<%=RankingsMobileFilename%>?RankingListType=NCWSA&df=yes' title='Rankings' style="text-decoration:none;" >
  		<img class="menuimage" src="<%=MenuIcon_C1%>">
  	</a>
	</span>
	<span class="menuicon" style="border:0px solid white;">
		<a href='<%=ThisSitePath%>/<%=TeamsMobileRankingFilename%>' title='Virtual Team Rankings' style="text-decoration:none;" >
  		<img class="menuimage" src="<%=MenuIcon_D1%>">

  	</a>	
	</span>
</div>	
<! -- ROW 2 -- ->
<div style="width:100%">
	<span class="menuicon" style="border:0px solid white;">
	 		<img class="menuimage" src="<%=MenuIcon_A2%>" onclick="javascript:NoService('Virtual Team Builder');">
  </span>
	<span class="menuicon" style="border:0px solid white;">
		<form action='<%=ThisSitePath%>/<%=MyStatsFilename%>' title='My Waterski Stats' style="text-decoration:none;" method="post">
			<input type="hidden" id="sMemberID_Hidden_InRankingsSettings" name="sMemberID" value="">
			<input type="hidden" id="sWatchMemberIDs_Local" name="sWatchMemberIDs" value="">
	 		<img class="menuimage" src="<%=MenuIcon_B2%>" onclick="submit()">
		</form>
  </span>
	<span class="menuicon" style="border:0px solid white;">
  		<img class="menuimage" src="<%=MenuIcon_C2%>" title='Twitter' onclick="javascript:MainMenuOptions('twitter');">
	</span>
	<span class="menuicon" style="border:0px solid white;">
		<a href='http://www.usawaterski.org' title='Link to Main USA Waterski Site' style="text-decoration:none;" >
  		<img class="menuimage" src="<%=MenuIcon_D2%>">
  	</a>	
	</span>
</div>	
<! -- ROW 3 -- ->
<div style="width:100%">
	<span class="menuicon" style="border:0px solid white;">
		<a href='<%=ThisSitePath%>/<%=LocalVarFileName%>' title='Set User' style="text-decoration:none;" >
  		<img class="menuimage" src="<%=MenuIcon_A3%>">
  	</a>	
	</span>
	<span class="menuicon" style="border:0px solid white;">
  	<img class="menuimage" src="<%=MenuIcon_B3%>" title='FAQ' onclick="javascript:MainMenuOptions('faq');">
	</span>
	<span class="menuicon" style="border:0px solid white;">
 		<img class="menuimage" src="<%=MenuIcon_C3%>" title='AWSA Resources' onclick="javascript:MainMenuOptions('ResourcesFromHome');">
	</span>
	<span class="menuicon" style="border:0px solid white;">
		<a href='<%=ThisSitePath%>/<%=SendLinkFileName%>' title='Send Link To Friend' style="text-decoration:none;" >		
 			<img class="menuimage" src="<%=MenuIcon_D3%>" title='Send Link to Friend'>
 		</a>
	</span>
</div>	
<! -- ROW 4 -- ->
<div style="width:100%">
	<span class="menuicon" style="border:0px solid white;">
	 	<img class="menuimage" src="<%=MenuIcon_A4%>" title="Real Time Scores" onclick="javascript:MainMenuOptions('realtimefromhome');">
  </span>
	<span class="menuicon" style="border:0px solid white;">
		&nbsp;
	</span>
	<span class="menuicon" style="border:0px solid white;">
		&nbsp;
	</span>
	<span class="menuicon" style="border:0px solid white;">
		&nbsp;
	</span>
</div>
<div style="width:100%; color:#FFFFFF; text-align:center; font-size:11pt; padding-top:25px;">To return to main menu tap the AWSA logo</div>	
<div style="width:96%; text-align:center; margin-top:10px; padding:0px; border:0px solid red; position:none; display:none;">		
  	<span class="span95" style="text-align:center; padding:0px; margin:0px; width:100%;">
			<textarea id="AddIconNowInstruction" name="AddIconNowInstruction" style="width:100%; font-size:12pt; padding-left:10px; text-align:center; color:yellow; background-color:black; border:0px solid;" cols=29 rows=3 wrap=physical><%=AddIconNowInstruction%></textarea>
  	</span> 	
</div>
</div> <! -- Overall Mainmenu div -- ->
<%

END SUB




' -----------------------------
  SUB DisplayTwitterFeed
' -----------------------------

%>
<div id="twitterfeedscreen" style="height=100%; width:96%; margin:0px 0px 0px 0px; padding:0px 0px 0px 0px; display:none;">
	<a class="twitter-timeline" style="width:95%; margin-top:0px;" href="https://twitter.com/USAWaterSki" data-widget-id="309767811616604160">Tweets by @USAWaterSki</a>
	<script>!function(d,s,id){var js,fjs=d.getElementsByTagName(s)[0];if(!d.getElementById(id)){js=d.createElement(s);js.id=id;js.src="//platform.twitter.com/widgets.js";fjs.parentNode.insertBefore(js,fjs);}}(document,"script","twitter-wjs");</script>	
</div>
<%	

END SUB



' ----------------------------------
  SUB DisplayRealTime
' ----------------------------------

%>
<div id="RealTimeScores" style="display:none; margin-top:5px; text-align:center;">
	<iframe src="http://www.waterskiresults.com/WfwWeb/wfwShowTourScores.php" style="width:310px; float:left; height:auto; color:white; margin-top:5px; padding:0px; margin-left:0px; border:0px solid white;"></iframe>	
</div>
<%

END SUB  



' ------------------------
  SUB DisplayFAQOptions
' ------------------------

%>
<div id="FAQMenuScreen" style="display:none;  margin:0px 0px 0px 0px; padding:0px 0px 0px 0px; text-align:center;">
	<div style="width:96%; margin-top:10px;  margin:0px 0px 0px 0px; padding:0px 0px 0px 10px; text-align:left; border:0px solid red;">		
			<span class="span95" style="margin:10px 0px 0px 0px; padding:0px 0px 0px 0px; text-align:center; font-size:14px; color:yellow; border:0px solid white;">Frequently Asked Questions</span> 
	</div>
	<div class="scroll" style="margin:10px 0px 0px 0px; padding:0px 0px 0px 0px; height:440px; border:0px solid white;">
		<div class="span95" style="height:50px;">
			<input type="button" class="buttonblue" name="Putting App on Home Screen" value="Put App iCon on iPhone" style="width:16em; height:2em; font-size:12pt;" onclick="javascript:FAQOptions('iphone');">
		</div>
		<div class="span95" style="height:50px;">
				<input type="button" class="buttonblue" name="Saving Search Settings" value="Saving Search Settings" style="width:16em; height:2em; font-size:12pt;" onclick="javascript:FAQOptions('savesearch');">
		</div>
		<div class="span95" style="height:50px;">
				<input type="button" class="buttonblue" name="Tournament Search" value="Tournament Search" style="width:16em; height:2em; font-size:12pt;" onclick="javascript:NoService('Tournament Search FAQ');">
		</div>
		<div class="span95" style="height:50px;">
			<a href="<%=ThisSitePath%>/<%=FAQ_RankingsFileName%>" title="FAQ for Rankings">
				<input type="button" class="buttonblue" name="National Rankings" value="National Rankings" style="width:16em; height:2em; font-size:12pt;">
			</a>
		</div>
		<div class="span95" style="height:50px;">
				<input type="button" class="buttonblue" name="Collegiate Rankings" value="Collegiate Rankings" style="width:16em; height:2em; font-size:12pt;" onclick="javascript:NoService('Collegiate Rankings FAQ');">
		</div>
		<div class="span95" style="height:50px;">
				<input type="button" class="buttonblue" name="Virtual Team Rankings" value="Virtual Team Rankings" style="width:16em; height:2em; font-size:12pt;" onclick="javascript:NoService('Virtual Team Rankings FAQ');">
		</div>
		<div class="span95" style="height:50px;">
				<input type="button" class="buttonblue" name="Virtual Team - Creating" value="Creating a Virtual Team" style="width:16em; height:2em; font-size:12pt;" onclick="javascript:NoService('Creating a Virtual Team FAQ');">
		</div>
		<div class="span95" style="height:50px;">
				<input type="button" class="buttonblue" name="My Stats" value="My Statistics" style="width:16em; height:2em; font-size:12pt;" onclick="javascript:NoService('My Statistics');">
		</div>
		<div class="span95" style="height:50px;">
				<input type="button" class="buttonblue" name="Set User" value="Set User" style="width:16em; height:2em; font-size:12pt;" onclick="javascript:NoService('Set User FAQ');">
		</div>
		<div class="span95" style="height:50px;">
				<input type="button" class="buttonblue" name="Online Registration" value="Online Registration" style="width:16em; height:2em; font-size:12pt;" onclick="javascript:NoService('FAQ');">
		</div>
	</div>	
</div>
<%

END SUB  




' -----------------------------
  SUB DisplayResourcesOptions
' -----------------------------

' -- http://www.usawaterski.org/pages/divisions/3event/2016AWSARuleBook.pdf
' -- http://www.usawaterski.org/pages/divisions/3event/AWSARuleBook.pdf

%>
<div id="ResoucesMenuScreen" style="display:none; width:96%; margin:0px 0px 0px 0px; padding:0px 0px 0px 0px; text-align:center; border:0px solid white;">
	<div style="width:96%; margin:0px 0px 0px 0px; padding:10px 0px 0px 10px; text-align:left; border:0px solid red;">		
		<span class="span95" style="margin:0px 0px 0px 0px; padding:0px 0px 0px 0px; text-align:center; font-size:14px; color:yellow; border:0px solid white;">AWSA Competition Resources</span> 
	</div>
	<div class="scroll" style="margin:10px 0px 0px 0px; padding:0px 0px 0px 0px; height:440px; border:0px solid white;">
		<div class="span95" style="height:50px;">
			<a href="http://www.usawaterski.org/pages/divisions/3event/AWSARuleBook.pdf" title="AWSA Rule Book" style="text-decoration:none;" target="_blank">
				<input type="button" class="buttonblue" name="AWSA Rule Book" value="AWSA Rule Book" style="width:16em; height:2em; font-size:12pt;">
			</a>
		</div>
		<div class="span95" style="height:50px;">
			<a href="<%=ThisSitePath%>/<%=NOPSCalcFilename%>" title="NOPS Overall Calculator" style="text-decoration:none;">
				<input type="button" class="buttonblue" name="NOPS Overall Calculator" value="NOPS Overall Calculator" style="width:16em; height:2em; font-size:12pt;">
  		</a>	
		</div>
		<div class="span95" style="height:50px;">
			<a href="<%=ThisSitePath%>/<%=ZBSCheatFilename%>" title="ZBS Cheat Sheet" style="text-decoration:none;">
				<input type="button" class="buttonblue" name="ZBS_Cheat_Sheet" value="ZBS Cheat Sheet" style="width:16em; height:2em; font-size:12pt;">
  		</a>	
		</div>
		<div class="span95" style="height:50px;">
			<a href="<%=ThisSitePath%>/<%=TimingChartFilename%>" title="Timing Charts" style="text-decoration:none;">
				<input type="button" class="buttonblue" name="Timing_Charts" value="Timing Charts" style="width:16em; height:2em; font-size:12pt;">
  		</a>	
		</div>					
	</div>	
</div>
<%

END SUB  





' --------------------------
  SUB DisplayiPhoneAddIcon
' --------------------------


FAQImage_A1 = FAQImagePath&"IPhone_Step1_Home.PNG"
FAQImage_B1 = FAQImagePath&"IPhone_Step2.PNG"
FAQImage_A2 = FAQImagePath&"IPhone_Step3.PNG"
FAQImage_B2 = FAQImagePath&"IPhone_Step4.PNG"

%>
<div id="iPhoneAddIcon" style="display:none; margin-top:5px; text-align:center;">
	<div style="width:96%; margin-top:10px; padding-left:10px; text-align:left; border:0px solid red;">		
			<span class="span95" style="margin-left:0px; padding-left:0px; text-align:center; font-size:14px; color:yellow; border:0px solid white;">Adding App Icon to iPhone Home Screen</span> 
	</div>
	<div class="scroll" style="margin-top:00px; padding:0px; margin-left:0px; height:400px; border:0px solid white;">
		<div style="width:96%; margin-top:10px; padding-left:10px; text-align:left; border:0px solid red;">		
			<span class="span95" style="margin-top:10px; margin-left:0px; padding-left:0px; text-align:center; font-size:12px; color:white; border:0px solid white;">To create the AWSA App icon on your phone's Home Screen you will begin at the App's Main Menu which may be accessed from the button at the bottom of this instruction. Performing these simple steps will permit you to access the App with a single click.<br><br>Scroll on the images below to continue.</span> 
		</div>
		<div style="margin-top:10px;">
			<span class="span95" style="margin-top:10px; margin-left:0px; padding-left:0px; text-align:center; font-size:12px; color:white; border:0px solid white;">STEP 1 - From the App's Home Screen (not now), press the Bookmark icon at the bottom center of the screen. If the bookmark icon is not visible, touch near the top of the screen and slide it downward slightly.</span> 			
			<span class="span95" style="margin-top:15px;"><img src="<%=FAQImage_A1%>" style="width:250px;" title='iPhone-Step1'></span>
			<span class="span95" style="margin-top:20px; margin-left:0px; padding-left:0px; text-align:center; font-size:12px; color:white; border:0px solid white;">STEP 2 - Next, locate and press the Add to Home Screen icon in the lower carousel.</span> 			
			<span class="span95"><img src="<%=FAQImage_B1%>" style="width:250px; margin-top:10px;" title='iPhone-Step2'></span>
			<span class="span95" style="margin-top:20px; margin-left:0px; padding-left:0px; text-align:center; font-size:12px; color:white; border:0px solid white;">STEP 3 - The last step is to press ADD in the upper right corner of the screen.</span> 			
			<span class="span95"><img src="<%=FAQImage_A2%>" style="width:250px; margin-top:10px;" title='iPhone-Step3'></span>
			<span class="span95" style="margin-top:20px; margin-left:0px; padding-left:0px; text-align:center; font-size:12px; color:white; border:0px solid white;">The icon should appear on your phone's Home Screen.  Launch the App from the icon.</span> 			
			<span class="span95"><img src="<%=FAQImage_B2%>" style="width:250px; margin-top:10px;" title='iPhone-Step4'></span>
		</div>
		<div class="span95" style="height:50px; margin-top:20px;">
			<input type="button" class="buttonblue" name="SetHomeScreen" value="Set Icon From Home Screen" style="width:16em; height:2em; font-size:12pt;" onclick="javascript:MainMenuOptions('HomeFromiPhone');">
		</div>
		<div class="span95" style="height:200px; margin-top:20px; padding-bottom:50px; vertical-align:top;">
			<input type="button" class="buttonblue" name="Return to FAQ" value="Return to FAQ" style="width:16em; height:2em; font-size:12pt;" onclick="javascript:FAQOptions('faqfromiPhone');">
		</div>
	</div>	
</div>
<%

END SUB  


  


' ----------------------------------
  SUB DisplaySavingSearchSettings
' ----------------------------------

%>
<div id="SavingSearchSettings" class="errorbox" style="height:480px; display:none; padding:0px 5px 0px 5px; margin:5px 0px 0px 0px; text-align:center;">
	<div style="width:96%; margin-top:10px; padding-left:10px; text-align:left; border:0px solid red;">		
			<span class="span95" style="margin-left:0px; padding-left:0px; text-align:center; font-size:14px; color:yellow; border:0px solid white;">Saving Your Search Settings</span> 
	</div>
	<div class="scroll" style="color:white; margin-top:5px; padding:0px; margin-left:0px; height:380px; border:0px solid yellow;">
		<div style="margin-top:10px;">
			Saving search settings lets you create a search view that can be recalled at any time. 
			<br><br>Saving search settings or setting the User is <b>NOT possible when your iPhone is set to Private Browsing.</b> 
			<br><br>Separate settings may be saved for the National and Collegiate Rankings.  
			<br><br>Once the settings have been saved, you can change the drop downs to perform other searches without affecting the saved settings.  If you change the dropdowns after pressing Save, returning to the search settings page will not automatically recall the saved settings.
			<br><br>To make your saved settings active again, press the Recall button. 	
		</div>
	</div>	
	<div class="span95" style="margin-top:10px; border:0px solid white;">
		<input type="button" class="buttonblue" name="Return to FAQ" value="Return to FAQ" style="width:16em; height:2em; font-size:12pt;" onclick="javascript:FAQOptions('faqfromSaveSearch');">
	</div>
</div>
<%

END SUB






' ----------------------------------
  SUB DisplayRuleBook
' ----------------------------------

%>
<div id="DisplayRulebook_TEMP" style="display:none; margin-top:5px; text-align:center;">
	<div class="scroll" style="width:310px; float:left; height:auto; color:white; margin-top:5px; padding:0px; margin-left:0px; border:0px solid white;">
		<embed type="application/pdf" src="http://www.usawaterski.org/pages/divisions/3event/2015AWSARuleBook.pdf" width="100px">
	</div>	
</div>
<%

END SUB  



' ----------------------------------
  SUB DisplayRankingsFAQ
' ----------------------------------

%>
<div id="DisplayRulebook" style="display:none; margin-top:5px; text-align:center;">
	<div class="scroll" style="width:310px; float:left; height:auto; color:white; margin-top:5px; padding:0px; margin-left:0px; border:0px solid white;">
		<embed type="application/pdf" src="http://www.usawaterski.org/pages/divisions/3event/2015AWSARuleBook.pdf" width="100px">
	</div>	
</div>
<%

END SUB  
%>

