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
Dim MyStatsFilename, LaunchFilename



MenuItemPath = "images\icons\"
FAQImagePath = "images\mobile_faq\"
ThisFileName = "mainmenu_m.asp"


' --- Names related programs for linking ---
SearchFileName = "search-memberHQ.asp"
RankingsMobileFilename="view-standings_m.asp"
TournamentsMobileFilename="view-tournaments_m.asp"
TeamsMobileRankingFilename="View-vteamstatus_m.asp"
LocalVarFileName="User_Set.asp"
MenuFileName = "mainmenu_m2.asp"
MyStatsFilename = "view-mystats_m.asp"
LaunchFilename = "awsa_launch.asp"






' --- Displays the html, head and opening body tag ---
OpenState="mainmenu"
DisplayHeadOpenBodyAndBannerTags OpenState



' --- Displays the menu for view tournaments --- 
'DisplayMenuButtons_ViewTournaments



DisplayMenu


DisplayTwitterFeed			' --- Initially hidden

DisplayFAQOptions				' --- Initially hidden

DisplayiPhoneAddIcon		' --- Initially hidden

'DisplayRuleBook					' --- Initially hidden

DisplaySavingSearchSettings

' --- Writes the Closing tags for HTML ---
DisplayCloseBodyAndHTMLTags






' ---------------------------------------------------
' --- BOTTOM OF MAIN CODE ---
' ---------------------------------------------------







' ------------------
  SUB DisplayMenu
' ------------------ 

' --- Defines what Icons are showing ---
MenuIcon_A1 = MenuItemPath&"Events_Blk_57.jpg"					' --- World with Magnifying glass
MenuIcon_B1 = MenuItemPath&"NationalRank_Blk_57.jpg"		' --- Graph
MenuIcon_C1 = MenuItemPath&"CollegeRank_Blk_57.jpg"			' --- Graduation cap
MenuIcon_D1 = MenuItemPath&"VTeamRank_Blk_57.jpg"						' --- Family

MenuIcon_A2 = MenuItemPath&"MyTeams_Blk_57.jpg"					' --- 4 People in circle
MenuIcon_B2 = MenuItemPath&"MyStats_57.jpg"							' --- Green Graph with Yellow Arrow
MenuIcon_C2 = MenuItemPath&""														' --- Green Graph with Yellow Arrow
MenuIcon_D2 = MenuItemPath&"FullSite_57.jpg"						' --- USA Waterski Logo to link to full site

MenuIcon_A3 = MenuItemPath&"Twitter_Blk_57.jpg"					' --- Twitter Feed --MenuIcon_C3 = MenuItemPath&"Twitter_Blk_57.jpg"					' --- Twitter Feed --
MenuIcon_B3 = MenuItemPath&"SetUser_Blk_57.jpg"					' --- Settings wheel
MenuIcon_C3 = MenuItemPath&"FAQ_Blk_57.jpg"							' --- Question Mark
MenuIcon_D3 = MenuItemPath&""														' --- 

MenuIcon_A4 = MenuItemPath&"Rulebook_Blk_57.jpg"								' --- Home Screen down arrow
MenuIcon_B4 = MenuItemPath&"Down_Green_Arrow.png"				' --- Home Screen down arrow
MenuIcon_C4 = ""
MenuIcon_D4 = ""


'class="span20" 
' class="span75" 

%>
<input type="hidden" id="sName_InRankingsSettings" name="sName_InRankingsSettings" value=""> 
<div id="mainmenuscreen" style="width:100%; display:inline;">
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
		<a href='<%=ThisSitePath%>/Vteams_Manage.asp' title='Manage My Teams' style="text-decoration:none;" >
	 		<img class="menuimage" src="<%=MenuIcon_A2%>">
		</a>
  </span>
	<span class="menuicon" style="border:0px solid white;">
		<form action='<%=ThisSitePath%>/<%=MyStatsFilename%>' title='My Waterski Stats' style="text-decoration:none;" >
			<input type="hidden" id="sMemberID_Hidden_InRankingsSettings" name="sMemberID" value="">
	 		<img class="menuimage" src="<%=MenuIcon_B2%>" onclick="submit()">
		</form>
  </span>
	<span class="menuicon" style="border:0px solid white;">
		&nbsp;
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
  		<img class="menuimage" src="<%=MenuIcon_A3%>" title='Twitter' onclick="javascript:MainMenuOptions('twitter');">
	</span>
	<span class="menuicon" style="border:0px solid white;">
		<a href='<%=ThisSitePath%>/<%=LocalVarFileName%>' title='Set User' style="text-decoration:none;" >
  		<img class="menuimage" src="<%=MenuIcon_B3%>">
  	</a>	
	</span>
	<span class="menuicon" style="border:0px solid white;">
  		<img class="menuimage" src="<%=MenuIcon_C3%>" title='FAQ' onclick="javascript:MainMenuOptions('faq');">
	</span>
	<span class="menuicon" style="border:0px solid white;">
  		&nbsp;
	</span>

</div>	
<! -- ROW 4 -- ->
<div style="width:100%">
	<span class="menuicon" style="border:0px solid white;">
  		<img class="menuimage" src="<%=MenuIcon_A4%>" title='Rulebook'>
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
</div> <! -- Overall Mainmenu div -- ->
	<div style="width:96%; text-align:center; margin-top:10px; padding-left:10px; border:0px solid red; position:none;">		
  	<span class="span95" style="text-align:center; padding:0px; margin:0px; width:100%;">
			<textarea id="AddIconNowInstruction" name="AddIconNowInstruction" style="width:100%; font-size:12pt; text-align:center; color:yellow; background-color:black; border:0px solid;" cols=29 rows=3 wrap=physical><%=AddIconNowInstruction%></textarea>
  	</span> 	
	</div>
<%

END SUB




' -----------------------------
  SUB DisplayTwitterFeed
' -----------------------------

%>
<div id="twitterfeedscreen" style="width:96%; margin:0px 0px 0px 0px; padding:0px 0px 0px 0px; display:none;">
	<a class="twitter-timeline" style="width:95%; margin-top:0px;" href="https://twitter.com/USAWaterSki" data-widget-id="309767811616604160">Tweets by @USAWaterSki</a>
	<script>!function(d,s,id){var js,fjs=d.getElementsByTagName(s)[0];if(!d.getElementById(id)){js=d.createElement(s);js.id=id;js.src="//platform.twitter.com/widgets.js";fjs.parentNode.insertBefore(js,fjs);}}(document,"script","twitter-wjs");</script>	
</div>
<%	

END SUB




' ------------------------
  SUB DisplayFAQOptions
' ------------------------

%>
<div id="FAQMenuScreen" style="display:none; margin-top:5px; text-align:center;">
	<div style="width:96%; margin-top:10px; padding-left:10px; text-align:left; border:0px solid red;">		
			<span class="span95" style="margin-left:0px; padding-left:0px; text-align:center; font-size:14px; color:yellow; border:0px solid white;">Frequently Asked Questions</span> 
	</div>
	<div class="scroll" style="margin-top:5px; padding:0px; margin-left:0px; height:400px; border:0px solid white;">
		<div class="span95" style="height:50px;">
			<input type="button" class="buttonblue" name="Putting App on Home Screen" value="Put App iCon on iPhone" style="width:16em; height:2em; font-size:12pt;" onclick="javascript:FAQOptions('iphone');">
		</div>
		<div class="span95" style="height:50px;">
				<input type="button" class="buttonblue" name="Saving Search Settings" value="Saving Search Settings" style="width:16em; height:2em; font-size:12pt;" onclick="javascript:FAQOptions('savesearch');">
		</div>
		<div class="span95" style="height:50px;">
				<input type="button" class="buttonblue" name="Tournament Search" value="Tournament Search" style="width:16em; height:2em; font-size:12pt;">
		</div>
		<div class="span95" style="height:50px;">
				<input type="button" class="buttonblue" name="National Rankings" value="National Rankings" style="width:16em; height:2em; font-size:12pt;">
		</div>
		<div class="span95" style="height:50px;">
				<input type="button" class="buttonblue" name="Collegiate Rankings" value="Collegiate Rankings" style="width:16em; height:2em; font-size:12pt;">
		</div>
		<div class="span95" style="height:50px;">
				<input type="button" class="buttonblue" name="Virtual Team Rankings" value="Virtual Team Rankings" style="width:16em; height:2em; font-size:12pt;">
		</div>
		<div class="span95" style="height:50px;">
				<input type="button" class="buttonblue" name="Virtual Team - Creating" value="Creating a Virtual Team" style="width:16em; height:2em; font-size:12pt;">
		</div>
		<div class="span95" style="height:50px;">
				<input type="button" class="buttonblue" name="My Stats" value="My Statistics" style="width:16em; height:2em; font-size:12pt;">
		</div>
		<div class="span95" style="height:50px;">
				<input type="button" class="buttonblue" name="Set User" value="Set User" style="width:16em; height:2em; font-size:12pt;">
		</div>
		<div class="span95" style="height:50px;">
				<input type="button" class="buttonblue" name="Online Registration" value="Online Registration" style="width:16em; height:2em; font-size:12pt;">
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



  SUB TT
 %>
 			<a href='http://usawaterski.org/rankings/awsa_launch.asp?action=applehomeicon' title='Go To Launch Screen' style="text-decoration:none;" >
				<input type="button" class="buttonblue" name="Set Home Screen Icon" value="Set Icon From Launch Screen" style="width:16em; height:2em; font-size:12pt;">
			</a>
<%
END SUB 
  


' ----------------------------------
  SUB DisplaySavingSearchSettings
' ----------------------------------


%>
<div id="SavingSearchSettings" style="display:none; margin-top:5px; text-align:center;">
	<div style="width:96%; margin-top:10px; padding-left:10px; text-align:left; border:0px solid red;">		
			<span class="span95" style="margin-left:0px; padding-left:0px; text-align:center; font-size:14px; color:yellow; border:0px solid white;">Saving Your Search Settings</span> 
	</div>
	<div class="scroll" style="color:white; margin-top:5px; padding:0px; margin-left:0px; height:400px; border:0px solid white;">
		<div style="margin-top:10px;">
			Saving your search settings allows you to create a search view that can be recalled at any time. 
			<br><br>Saving search settings or setting the User is NOT possible when your iPhone is set to Private Browsing.  This does not apply once you have the App installed as an icon on your home screen. 
			<br><br>Separate settings may be saved for the National and Collegiate Rankings.  
			<br><br>Once the settings have been saved, you may then change the drop downs to perform other searches without affecting the saved settings.  If you change the dropdowns after pressing Save, simply returning to the search settings page without will not automatically recall the saved settings.
			<br><br>To make your saved settings active, simply press the Recall button. 	
		</div>
		<div class="span95" style="height:50px; margin-top:20px; padding-bottom:50px;">
			<input type="button" class="buttonblue" name="Return to FAQ" value="Return to FAQ" style="width:16em; height:2em; font-size:12pt;" onclick="javascript:FAQOptions('faqfromSaveSearch');">
		</div>
	</div>	
</div>
<%

END SUB



SUB TT
%>
<embed src="http://www.usawaterski.org/pages/divisions/3event/2015AWSARuleBook.pdf">
<object data="http://www.usawaterski.org/pages/divisions/3event/2015AWSARuleBook.pdf" type="application/pdf" width="300px">2015AWSARuleBook.pdf</object>
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

