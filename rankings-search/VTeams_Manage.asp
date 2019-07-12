<!--#include file="settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->
<!--#include virtual="/rankings/tools_Definitions.asp"-->
<!--#include virtual="/rankings/tools_mobile_version.asp"-->
<!--#include virtual="/rankings/tools_leagues.asp"-->
<!--#include virtual="/rankings/tools_registration16.asp"-->
<%


Dim ThisFileName

Dim cMemberID, cFirstName, cLastName, cPassword
Dim sMemberID, sLastName, sFirstName, sFullName, sMembSex, sMembSexText, sMembCity, sMembState, sMembAge, sPassword, sMembPhone, sMembTypeID, sCanSkiTour, sMembTypeCode
Dim sMembEmail, sEffectiveTo, sMembBirth, sEmail, sCostToUpgrade, sTypeDesc

Dim sTeam_Type_Description, sMax_Members, sMin_Members, sMax_Male, sMin_Male, sMax_Female, sMin_Female, sMax_Age, sMin_Age
Dim sMax_Scoring, sMin_Scoring, sMax_Scoring_Male, sMin_Scoring_Male, sMax_Scoring_Female, sMin_Scoring_Female

Dim sTeam_ID, sTeam_Name, sTeam_Level, sManager_MemberID, sManager_FirstName, sManager_LastName, sTeamStatus
Dim sTeam_Type_ID, sCreated_Date, sNo_Team_Members
Dim MemberCount
Dim sTeamMemberStatus, sTeamMemberStatusText, sTeamMemberStatusTextColor, sTeamStatusText, sTeamStatusTextColor
Dim sThisTeamTypeDescription, sThisTeamName
Dim TabColor
Dim ebody, ecss

Dim AddTeamMemberButtonStatus, DeleteTeamMemberButtonStatus
Dim sThis_TeamMember_Count, sThis_Female_Count, sThis_Male_Count, sThis_Accepted, sThis_Invited, sThis_Needs_Invite

Dim action, sMemberFound
Dim SetLocalButtonStatus

Dim TeamTypeIDSelected, TeamIDSelected

ThisFileName="VTeams_Manage.asp"


' --- Names related programs for linking ---
SearchFileName = "search-memberHQ.asp"
RankingsMobileFilename=ThisFileName 
TournamentsMobileFilename="view-tournaments_m.asp"
LocalVarFileName="Test_localstorage_SET.asp"
TeamsMobileFilename="virtualteam_m.asp"
MenuFileName = "mainmenu_m2.asp"




action = LCASE(Request("action"))

TeamIDSelected = 0
IF TRIM(Request("TeamIDSelected")) <> "" THEN TeamIDSelected = Request("TeamIDSelected")

'response.write("TeamIDSelected = "&TeamIDSelected)

TeamTypeIDSelected = 0
IF TRIM(Request("TeamTypeIDSelected")) <> "" THEN TeamTypeIDSelected = Request("TeamTypeIDSelected")

'response.write("<br>action = " & action)



' --- Displays the html, head and opening body tag ---
OpenState="manageteams"
DisplayHeadOpenBodyAndBannerTags OpenState






SELECT CASE action
		CASE "createteam"
				'response.write("</div>Create</div>")
		
		' -- Temporary 		
		'CASE "teamlist"	
				RunMyTeamListingQuery
				
				DisplayManageMyTeam
				
				DisplayConfirmInviteTeamMemberScreen

	
		CASE "sendinvite"
				
				SendEmailinviteToTeamMember

		
		CASE "acceptinvite", "declineinvite"		
				
				UpdateFromInvitation


		CASE "teamtypeselect"
				RunMyTeamListingQuery
				
				DisplayTeamWhatAction "none"
				

				RunTeamTypeParametersQuery				
				
				DisplayTeamTypeSelectionScreen "inline-block"
				
				DisplayTeamNameEntryScreen "none"			

		
		CASE ELSE
							
				RunMyTeamListingQuery
				
				DisplayTeamWhatAction "inline-block"
				

				RunTeamTypeParametersQuery				
				
				DisplayTeamTypeSelectionScreen "none"
				
				DisplayTeamNameEntryScreen "none"
				
				'response.write("</div><div>CASE ELSE</div>")
END SELECT







' --- Writes the Closing tags for HTML ---
DisplayCloseBodyAndHTMLTags



' ------------------------------------------------------------------------------------------
' --- BOTTOM OF MAIN CODE ---
' ------------------------------------------------------------------------------------------






' ---------------------------------------
  SUB DisplayTeamWhatAction (whatdisplay)
' ---------------------------------------

%>
<form action="<%=ThisFileName%>" method="post">
	<div id="TeamWhatActionScreen" style="width:96%; display:<%= whatdisplay %>; margin-top:5px; height:470px; padding:0px; margin:0px;">
		<input type="hidden" id="sMemberID_Hidden_InRankingsSettings" name="sMemberID_Hidden_InRankingsSettings" value="">
		<input type="hidden" id="sName_InRankingsSettings" name="sName_InRankingsSettings" value="">		
		<div style="width:100%; margin-top:5px; padding-left:10px; font-size:10pt; text-align:left; border:0px solid red;">
			<span class="span95" style="text-align:center;">
			<%
			
			BuildTeamListDropDown 20,15, "submit()"
			
			%>
			</span>			
		</div>
		<div class="scroll" style="width:100%; height:300px; margin:10px 0px 0px 0px; padding:0px 0px 0px 2px;">
			<%   
		
			' --- Displays the filter dropdowns inside ---
			LoopThruMyTeam_SimpleList
		
			%>
		</div> <! -- Bottom of scroll box -- ->
		<div id="teambuildselectionbuttons" class="menucell" style="width:100%; text-align:center; padding:0px; margin-top:35px; border:0px solid white; display:inline-block;">	
				<span style="width:48%; height:30px; border:0px solid green; margin:0px; padding:0px; text-align:left;">
					<input type="button" name="AddTeam" value="Add Team" style="width:8em; height:2em; font-size:12pt; text-align:center;" <%=AddTeamButtonStatus%> onclick="javascript:TeamCreateNav('totypeselectfromteamwhataction_add');">
				</span>
				<span style="width:48%; border:0px solid yellow; margin:0px; padding:1px; text-align:right;">
					<input type="button" name="EditTeam" formaction="<%=TheFileName%>?action=CreateTeam&sTeam_Type_ID=<%=TeamTypeIDSelected%>" value="Edit Team" style="width:8em; height:2em; font-size:12pt; text-align:center;" <%=EditTeamButtonStatus%> onclick="javascript:TeamCreateNav('totypeselectfromteamwhataction_edit');">
				</span>
		</div>
	</div>
</form>
<%	


END SUB








' --------------------------------------------------
  SUB DisplayTeamTypeSelectionScreen (whatdisplay)
' --------------------------------------------------

%>
<div class="errorbox" id="teamtypeselectionscreen" style="display:<%= whatdisplay %>; margin-top:5px; height:440px;">
	<form method="post">
		<input type="hidden" id="action" name="action" value="TeamTypeSelect">
		<div style="margin-top:10px; padding-left:10px; border:0px solid white;">
  		<span class="span100" style="text-align:center; color:yellow; font-size:12pt">Select a 'League' and make sure your team meets all of the criteria</span>
		</div>
		<div style="width:100%; margin-top:10px; padding-left:10px; text-align:center; border:0px solid red;">		
			<span class="span95" style="text-align:left;">
			<%
				' -- In Tools_Leagues --
				BuildTeamType_DropDown 15,14, "submit()"
			
			%>
		</span>
	</div>

		<div style="width:100%; margin-top:10px; padding-left:10px; font-size:12pt; text-align:left;">		
			<span class="span45" style="text-align:right; color:#FFFFFF;">&nbsp;</span> 
			<span class="span15" style="text-align:center; color:#FFFFFF;">Max</span>
			<span class="span15" style="text-align:center; color:#FFFFFF;">Min</span>
		</div>
		<div style="width:100%; margin-top:5px; padding-left:10px; font-size:10pt; text-align:left;">		
			<span class="span45" style="text-align:right; color:#FFFFFF;">Total Members:</span> 
			<span class="span15" style="text-align:center; font-size:12pt; color:yellow;"><%=sMax_Members%></span>
			<span class="span15" style="text-align:center; font-size:12pt; color:yellow;"><%=sMin_Members%></span>
		</div>
		<div style="width:100%; margin-top:4px; padding-left:10px; font-size:10pt; text-align:left;">		
			<span class="span45" style="text-align:right; color:#FFFFFF; font-size:10pt">Total Male:</span> 
			<span class="span15" style="text-align:center; font-size:12pt; color:yellow;"><%=sMax_Male%></span>
			<span class="span15" style="text-align:center; font-size:12pt; color:yellow;"><%=sMin_Male%></span>
		</div>
		<div style="width:100%; margin-top:4px; padding-left:10px; font-size:10pt; text-align:left;">		
			<span class="span45" style="text-align:right; color:#FFFFFF; font-size:10pt">Total Female:</span> 
			<span class="span15" style="text-align:center; font-size:12pt; color:yellow;"><%=sMax_Female%></span>
			<span class="span15" style="text-align:center; font-size:12pt; color:yellow;"><%=sMin_Female%></span>
		</div>
		<div style="width:100%; margin-top:4px; padding-left:10px; font-size:10pt; text-align:left;">		
			<span class="span45" style="text-align:right; color:#FFFFFF; font-size:10pt">Age Range:</span> 
			<span class="span15" style="text-align:center; font-size:12pt; color:yellow;"><%=sMax_Age%></span>
			<span class="span15" style="text-align:center; font-size:12pt; color:yellow;"><%=sMin_Age%></span>
		</div>
		<hr style="width:85%;">
		
		<div style="width:100%; margin-top:15px; padding-left:10px; font-size:10pt; text-align:left;">		
			<span class="span95" style="text-align:center; color:#FFFFFF;"># of Members Counted in Team Scoring</span>
		</div>
		<div style="width:100%; margin-top:8px; padding-left:10px; font-size:10pt; text-align:left;">		
			<span class="span45" style="text-align:right; color:#FFFFFF;">Total Members:</span> 
			<span class="span15" style="text-align:center; font-size:12pt; color:yellow;"><%=sMax_Scoring%></span>
			<span class="span15" style="text-align:center; font-size:12pt; color:yellow;"><%=sMin_Scoring%></span>
		</div>
		<div style="width:100%; margin-top:4px; padding-left:10px; font-size:10pt; text-align:left;">		
			<span class="span45" style="text-align:right; color:#FFFFFF;">Male:</span> 
			<span class="span15" style="text-align:center; font-size:12pt; color:yellow;"><%=sMax_Scoring_Male%></span>
			<span class="span15" style="text-align:center; font-size:12pt; color:yellow;"><%=sMin_Scoring_Male%></span>
		</div>					
		<div style="width:100%; margin-top:4px; padding-left:10px; font-size:10pt; text-align:left;">		
			<span class="span45" style="text-align:right; color:#FFFFFF; font-size:10pt">Female:</span> 
			<span class="span15" style="text-align:center; font-size:12pt; color:yellow;"><%=sMax_Scoring_Female%></span>
			<span class="span15" style="text-align:center; font-size:12pt; color:yellow;"><%=sMin_Scoring_Female%></span>
		</div>

		<div style="width:100%; margin-top:28px; padding-left:10px; font-size:10pt; text-align:left; height:25px; border: 0px solid white;">		
			<span id="TeamSelect" class="span95" style="text-align:center; color:red; font-size:12pt; display:none;">Please SELECT a League!</span>
		</div>								
		<div id="teamtypeselectionbuttons" class="menucell" style="padding:0px; margin-top:15px;">
			<TABLE align=center width=100% style="padding:0px; margin:0px; border:0px solid; border-color:yellow;">
				<tr>
				<td width="46%" height="30px" style="border:0px solid; border-color:#FFFFFF; margin:0px 0px 0px 0px; padding:0px 0px 0px 0px; text-align:center;">
					<input type="button" name="Create_Team" value="Create Team" style="width:8em; height:2em; font-size:12pt; text-align:center;" onclick="javascript:TeamCreateNav('toteamnameentryfromtypeselect');">
				</td>
				<td width="46%" align="center">
					<input type="submit" formaction="<%=MenuFileName%>" name="Cancel" value="Cancel" style="width:8em; height:2em; font-size:12pt; text-align:center;">
				</td>				
			</tr>
			</TABLE>
		</div>
	</form>	
</div>  
<%	

 

END SUB





' ----------------------------------------------
  SUB DisplayTeamNameEntryScreen (whatdisplay)
' ----------------------------------------------

%>
<div class="errorbox" id="teamnameentryscreen" style="display:<%= whatdisplay %>; margin-top:5px; height:440px;">
	<form method="post" name="CreateTeamForm" id="CreateTeamForm" action="<%=ThisFileName%>?action=CreateTeam&sTeam_Type_ID=<%=TeamTypeIDSelected%>">
		<input type="hidden" id="sMemberID_Hidden_InRankingsSettings" name="sMemberID_Hidden_InRankingsSettings" value="">
		
		<div style="margin-top:10px; padding-left:10px; border:0px solid white;">
  		<span class="span90" style="text-align:center; color:yellow; font-size:12pt">Enter Team Name and a Location</span>
		</div>
		<div style="width:100%; margin-top:10px; padding-left:10px; text-align:left; border:0px solid red;">		
			<span class="span25" style="text-align:right; font-size:12pt; color:white; border:0px solid white;">League:</span> 
			<span class="span70" style="text-align:left; font-size:12pt; color:yellow;"><%=sTeam_Type_Description%></span>
		</div>
		<div style="width:100%; margin-top:5px; padding-left:10px; text-align:left; border:0px solid red;">		
			<span class="span25" style="text-align:right; font-size:12pt; color:#FFFFFF; border:0px solid white; margin-top:3px;">Manager:</span> 
			<span class="span70" style="text-align:left;">
				<input type="text" class="textbox_hidden_banner" id="sName_InRankingsSettings" name="sName_InRankingsSettings"  value="" style="margin:0px; padding:0px; width:200px; text-align:left; font-weight:bold; color:yellow; font-size:12pt;" MaxLength="25">
			</span>	
		</div>

		<div style="width:100%; margin-top:10px; padding-left:10px; font-size:10pt; text-align:left;">		
			<span class="span25" style="text-align:right; color:#FFFFFF;">Team Name:</span> 
			<span class="span70" style="text-align:left; color:yellow;">
				<input type="text" name="cTeamName" id="sTeamName" placeholder="Team Name" value="" size="20" maxlength="20" style="font-size:12pt;">
			</span>
		</div>
		<div style="width:100%; margin:15px 0px 10px 0px; padding-left:10px; font-size:10pt; text-align:left;">		
			<span class="span25" style="text-align:right; color:#FFFFFF;">Location:</span> 
			<span class="span70" style="text-align:left; color:yellow;">
				<input type="text" name="cTeamLocation" id="sTeamLocation" placeholder="Location Reference" value="" maxlength="20" size="20" style="font-size:12pt;">
			</span>
		</div>

		<div style="width:100%; margin-top:20px; padding-left:10px; font-size:10pt; text-align:left;">		
			<span class="span25" style="text-align:right; color:white;">Status:</span>
			<span class="span70" style="text-align:left;">
				<SELECT id="sDisplayStatus" name="sDisplayStatus" style="width:11em; font-size:14pt;">
					<option value="A">Active</option>
					<option value="I">Inactive</option>
				</SELECT>		
			</span>
		</div>
		<div style="margin-top:60px; padding-left:10px; border:0px solid yellow;">
  		<span class="span90" style="text-align:center; color:white; font-size:10pt">When you press Save Team you can begin adding members.</span>
		</div>

		<div style="width:100%; margin-top:20px; padding-left:10px; font-size:10pt; text-align:left; height:25px; border: 0px solid white;">	
			<span id="TeamNameError" class="span95" style="text-align:center; color:red; font-size:12pt; display:none;">Team Name and Location must be at least 5 character (A-Z;a-z;0-9 allowed)</span>
		</div>
				
		<div id="teamtypeselectionbuttons" class="menucell" style="padding:0px; margin-top:25px;">
			<TABLE align=center width=100% style="padding:0px; margin:0px; border:0px solid yellow;">
				<tr>
					<td width="46%" height="30px" style="border:0px solid; border-color:#FFFFFF; margin:0px; padding:0px; text-align:center;">
						<input type="button"  name="ConfirmTeam" id="ConfirmTeam" value="Save Team" id="TeamNameSave" onclick="javascript:TeamCreateNav('toteammemberlistfromnameinput');" style="width:8em; height:2em; font-size:12pt; text-align:center;">
					</td>
					<td width="46%" align="center">
						<input type="button" name="Back" value="Back" style="width:8em; height:2em; font-size:12pt; text-align:center;" onclick="javascript:TeamCreateNav('totypeselectfromteamnameentry');">
					</td>				
				</tr>
			</TABLE>
		</div>
	</form>	
</div>  
<%	


END SUB






' ------------------------------------
  SUB DisplayManageMyTeam
' ------------------------------------


' --- Determines if the current team make-up meets the min and max requirements of this Team Type ---
DetermineTeamStatus


%>
<div class="errorbox" id="TeamManageOptionsScreen" style="width:100%; display:inline-block; margin-top:5px; height:440px;">
		<input type="hidden" id="sMemberID_Hidden_InRankingsSettings" name="sMemberID_Hidden_InRankingsSettings" value="">
		<input type="hidden" id="sName_InRankingsSettings" name="sName_InRankingsSettings" value="">		
		<div style="margin-top:10px; padding-left:10px; border:0px solid white;">
  		<span class="span90" style="text-align:center; color:yellow; font-size:12pt">Add or Delete members to meet League requirements</span>
		</div>
		<div style="width:100%; margin-top:5px; padding-left:10px; font-size:10pt; text-align:left; border:0px solid red;">		
			<span class="span25" style="text-align:right; color:#FFFFFF; border:0px solid white;">Status:</span> 
			<span class="span70" style="text-align:left; color:orange;"><%=sTeamStatus%></span>	
		</div>
		<div class="scroll" style="height:225px; margin:10px 0px 0px 0px; padding:0px 0px 0px 0px;">
		<%   
		
		' --- Displays the filter dropdowns inside ---
		LoopThruMyTeam_Manage
		
		%>
	</div> <! -- Bottom of scroll box -- ->

		
		<div id="teamtypeselectionbuttons" class="menucell" style="padding:0px; margin-top:12px;">
			<TABLE align=center width=100% style="padding:0px; margin:0px; border:0px solid; border-color:yellow;">
				<tr>
					<td width="32%" height="30px" style="border:0px solid; border-color:#FFFFFF; margin:0px; padding:0px; text-align:center;">
						<input type="submit" name="AddMember" value="Add" style="width:6em; height:2em; font-size:12pt; text-align:center;" <%=AddTeamMemberButtonStatus%>>
					</td>
					<td width="32%" height="30px" style="border:0px solid; border-color:#FFFFFF; margin:0px; padding:0px; text-align:center;">
						<input type="submit" name="DeleteMember" value="Drop" style="width:6em; height:2em; font-size:12pt; text-align:center;" <%=DeleteTeamMemberButtonStatus%>>
					</td>
					<td width="32%" align="center">
						<input type="button" name="Done" value="Done" style="width:6em; height:2em; font-size:12pt; text-align:center;">
					</td>				
				</tr>
			</TABLE>
		</div>
</div>  
<%	


END SUB






' ------------------------------------------
  SUB DisplayConfirmInviteTeamMemberScreen
' ------------------------------------------

%>
<div class="errorbox" id="ConfirmInviteNewMemberScreen" style="display:none; margin-top:5px; height:440px;">
	<form action="<%=ThisFileName%>?action=sendinvite&returnaction=teamlist" method="post">
		<input type="hidden" id="sTeam_ID" name="sTeam_ID" value="<%=sTeam_ID%>">
		<div style="margin-top:10px; border:0px solid white;">
  		<span class="span100" style="text-align:center; color:#FFFFFF; font-size:14pt; padding-left:10px;">Confirm invitation to this member</span>
		</div>
		
		<div style="width:100%; margin-top:20px; color:yellow; padding-left:10px;">When you press 'Send Invite' this member will receive an email containing a link.  When the member clicks 'Confirm Invitation' from the email their record for this team automatically update</div>	
		<div style="width:100%; margin-top:20px; padding-left:10px;">		
			<span class="span20" style="text-align:right; color:#FFFFFF; font-size:10pt">Name:</span> 
			<span class="span70" style="text-align:left;">
				<input type="text" class="textbox_hidden_banner" id="sInviteName" name="sInviteName" value="" style="text-align:left; color:yellow;" MaxLength="25">
			</span>
		</div>
		<div style="width:100%; margin-top:5px; padding-left:10px;">		
			<span class="span20" style="text-align:right; color:#FFFFFF; font-size:10pt">City/ST</span> 
			<span class="span70" style="text-align:left;">
				<input type="text" class="textbox_hidden_banner" id="sInviteCityState" name="sInviteCityState" value="" style="text-align:left; color:yellow;" MaxLength="25">
			</span>
		</div>
		<div style="width:100%; margin-top:5px; padding-left:10px;">		
			<span class="span20" style="text-align:right; color:#FFFFFF; font-size:10pt">MemberID:</span> 
			<span class="span70" style="text-align:left;">
				<input type="text" class="textbox_hidden_banner" id="sInviteMemberID" name="sInviteMemberID" value="" style="text-align:left; color:yellow;" MaxLength="10">
			</span>
		</div>
		<div style="width:100%; margin-top:20px; padding-left:10px;">		
			<span class="span20" style="text-align:right; color:#FFFFFF; font-size:10pt">Email:</span> 
			<span class="span70" style="text-align:left;">
				<input type="text" class="textbox_hidden_banner" id="sInviteEmail" name="sInviteEmail" value="" style="text-align:left; color:yellow;" size="25" MaxLength="50">
			</span>
		</div>

		<div id="ConfirmInviteButtons" class="menucell" style="padding:0px; margin-top:37px;">
			<TABLE align=center width=100% style="padding:0px; margin:0px; border:0px solid; border-color:yellow;">
				<tr>
				<td width="46%" height="30px" style="border:0px solid; border-color:#FFFFFF; margin:0px; padding:0px; text-align:center;">
						<input type="submit" name="SendInvite" value="Send Invite" style="width:8em; height:2em; font-size:12pt;">
				</td>
				<td width="46%" align="center">
					<input type="button" name="Cancel" value="Cancel" style="width:8em; height:2em; font-size:12pt;" onclick="javscript:TeamConfirmInviteNav('ToMyTeamListFromConfirmSendInvite','','','','','')">
				</td>				
			</tr>
			</TABLE>
		</div>
	</form>	
</div>  
<%	


END SUB






' --------------------------------------
  SUB DisplaySearchNewTeamMemberScreen
' --------------------------------------

%>
<div class="errorbox" id="SearchNewMemberScreen" style="display:none; margin-top:5px; height:440px;">
	<form action="<%=ThisFileName%>?action=findmember" method="post">
		
		<div style="margin-top:10px; border:0px solid white;">
  		<span class="span100" style="text-align:center; color:#FFFFFF; font-size:14pt">Set User For This Device</span>
		</div>
		
		<div style="width:100%; margin-top:20px; color:yellow;">To validate a member you must provide 
			<br>Password and Name (or MemberID)
		</div>	
		<div style="width:100%; margin-top:20px; padding-left:10px;">		
			<span class="span45" style="text-align:left; color:#FFFFFF; font-size:10pt">Last</span> 
			<span class="span45" style="text-align:left; color:#FFFFFF; font-size:10pt">First</span> 
			<span class="span45" style="text-align:left;">
				<input type="text" name="cLastName" id="LastName" value="" size="18">
			</span>
			<span class="span45" style="text-align:left;">
				<input type="text" name="cFirstName" id="FirstName" value="" size="18">
			</span>
		</div>
		<div style="width:100%; margin-top:15px; padding-left:10px;">		
			<span class="span45" style="text-align:left; color:#FFFFFF; font-size:10pt">MemberID</span> 
			<span class="span45" style="text-align:left;">
				<input type="text" name="cMemberID" id="MemberID" value="" size="10"> 
			</span>
		</div>
		
		<div id="userentrybuttons" class="menucell" style="padding:0px; margin-top:37px;">
			<TABLE align=center width=100% style="padding:0px; margin:0px; border:0px solid; border-color:yellow;">
				<tr>
				<td width="46%" height="30px" style="border:0px solid; border-color:#FFFFFF; margin:0px; padding:0px; text-align:center;">
						<input type="submit" name="FindMember" value="Find Member" style="width:8em; height:2em; font-size:12pt;">
				</td>
				<td width="46%" align="center">
					<input type="button" name="Cancel" value="Cancel" style="width:8em; height:2em; font-size:12pt;" onclick="javascript:SetUserNav('returntocurrentuser');">
				</td>				
			</tr>
			</TABLE>
		</div>
	</form>	
</div>  
<%	


END SUB



' -----------------------------------
  SUB DisplayConfirmNewMemberScreen
' -----------------------------------

sMembName = sFirstName&" "&sLastName
sMembCityState = sMembCity&" "&sMembState

IF sMemberFound = false THEN SetLocalButtonStatus="disabled"


%>
<div id="ConfirmNewMemberScreen" class="errorbox" style="display:block-inline; margin-top:5px; height:440px;">
	<input type="hidden" id="sMemberID_Hidden_InRankingsSettings" name="sMemberID_Hidden_InRankingsSettings" value="">
	<input type="hidden" id="sName_InRankingsSettings" name="sName_InRankingsSettings" value="">
	<div style="margin-top:10px;">
  		<span class="span100" style="text-align:center; color:#FFFFFF; font-size:14pt">Set to This Member</span>
	</div>
	<div style="width:100%; margin-top:20px; padding-left:10px; text-align:left; border:0px solid red;">		
		<span class="span20" style="margin-left:0px; padding-left:0px; text-align:right; color:#FFFFFF; font-size:10pt; border:0px solid white;">Name:</span> 
		<span class="span75" style="text-align:left;">
			<input type="text" class="textbox_hidden_banner" id="sMembName" name="sMembName"  value="<%=sMembName%>" style="text-align:left; color:yellow;" MaxLength="25">
		</span>
	</div>
	<div style="width:100%; margin-top:10px; padding-left:7px; text-align:left;">		
			<span class="span20" style="text-align:right; color:#FFFFFF; font-size:10pt; border:0px solid white;">City/ST:</span>
			<span class="span75" style="text-align:left;">
				<input type="text" name="sMembCityState" id="sMembCityState" value="<%=sMembCityState%>" size="10" style="text-align:left; color:yellow; background-color:#000000; border:0px"> 
			</span>
	</div>
	<div style="width:100%; margin-top:10px; padding-left:10px; text-align:left;">		
			<span class="span20" style="text-align:right; color:#FFFFFF; font-size:10pt; border:0px solid white;">Age:</span>
			<span class="span25" style="text-align:left;">
				<input type="text" name="sMembAge" id="sMembAge" value="<%=sMembAge%>" size="10" style="text-align:left; color:yellow; background-color:#000000; border:0px"> 
			</span>
			<span class="span20" style="text-align:right; color:#FFFFFF; font-size:10pt; border:0px solid white;">Gender:</span>
			<span class="span25" style="text-align:left;">
				<input type="text" name="sMembSexText" id="sMembSexText" value="<%=sMembSexText%>" size="10" style="text-align:left; color:yellow; background-color:#000000; border:0px"> 
			</span>
	</div>
	<div style="width:100%; margin-top:10px; padding-left:7px; text-align:left;">		
			<span class="span20" style="text-align:right; color:#FFFFFF; font-size:10pt; border:0px solid white;">MemberID:</span>
			<span class="span25" style="text-align:left;">
				<input type="text" name="sMemberID" id="sMemberID" value="<%=sMemberID%>" size="10" style="text-align:left; color:yellow; background-color:#000000; border:0px"> 
			</span>
			<span class="span20" style="text-align:right; color:#FFFFFF; font-size:10pt; border:0px solid white;">Expires:</span>
			<span class="span25" style="text-align:left;">
				<input type="text" name="sEffectiveto" id="sEffectiveto" value="<%=sEffectiveto%>" size="10" style="text-align:left; color:yellow; background-color:#000000; border:0px"> 
			</span>
	</div>

	<div style="width:100%; margin-top:10px; padding-left:15px; text-align:left;">		
			<span class="span85" style="text-align:center; color:yellow; font-size:12pt; border:0px solid white;">WARNING !!</span>
			<span class="span85" style="text-align:center; color:#FFFFFF; font-size:8pt; border:0px solid white;">Accessing this function for another member is strictly prohibited without their expressed permission.  Continuing without said permission may be subject to civil liability or criminal prosecution under <b> state and federal laws.</span>
	</div>

	<div id="storeuserbuttons" class="menucell" style="padding:0px; margin-top:15px;">
		<TABLE align=center width=100% style="padding:0px; margin:0px; border:0px solid; border-color:yellow;">
			<tr>
				<td width="46%" height="30px" style="border:0px solid; border-color:#FFFFFF; margin:0px; padding:0px; text-align:center;">
					<form action="<%=ThisFileName%>?action=done" method="post">
						<input type="submit" name="submit" id="submit5" value="Add Now" style="width:8em; height:2em; font-size:12pt;">
					</form>
				</td>
				<td width="46%" align="center">
					<form action="<%=ThisFileName%>?action=cancel" method="post">
						<input type="submit" name="Cancel" value="Cancel" style="width:8em; height:2em; font-size:12pt;"">
					</form>	
				</td>				
			</tr>
		</TABLE>
	</div>
</div>  <! -- for errorbox -- ->
<%	


END SUB






' ---------------------------
  SUB LoopThruMyTeam_Manage
' ---------------------------
 
MemberCount = 1

ThisTeam_ID = rs("Team_ID")

DO WHILE NOT rs.eof

		GetCurrentMemberLine
		
		IF MemberCount=1 THEN DisplayTeamTab

		DisplayTeamMemberLine
		
		MemberCount = MemberCount + 1
		
		rs.movenext

LOOP

DisplayTeamBottomLine

rs.close

END SUB




' ---------------------------
  SUB LoopThruMyTeam_SimpleList
' ---------------------------
 
MemberCount = 0
ThisTeam_ID = 0

DO WHILE NOT rs.eof

		IF rs("Team_ID")<>ThisTeamID THEN 
				ThisTeamID = rs("Team_ID")
				IF MemberCount <> 0 THEN DisplayTeamBottomLine
				MemberCount = 1
		END IF

		GetCurrentMemberLine
		
		IF MemberCount=1 THEN DisplayTeamTab_Simple
				
		DisplayTeamMemberLine_Simple
		
		MemberCount = MemberCount + 1
		
		rs.movenext
		

LOOP

DisplayTeamBottomLine

rs.close

END SUB





' -------------------
  SUB DisplayTeamTab
' -------------------

%> 
  <div class="tabrankings" style="width:96%; background-color:<%=TabColor%>; height:35px; margin-top:0px; padding-top:0px;" >
		<span class="span80" style="font-size:12pt; color:#000000; font-weight:bold;"><%= sTeam_Name %></span>
	  <br>
  	<span style="width=15%; color:black; font-weight:normal; border:0px solid white;"># Members:</span>
  	<span class="span10" style="color:blue; font-weight:normal; text-align:left;"><%= sNo_Team_Members %></span>
  	<span class="span10" style="color:black; text-align:right; font-weight:normal;">ID:</span>
	  <span class="span20" style="color:blue; font-weight:normal; text-align:left;"><%= sTeam_ID %></span>
	  <span class="span10" style="font-weight:normal; color:black; text-align:right;">Status: </span>	  
	  <span class="span20" style="color:<%=sTeamStatusTextColor%>; font-weight:normal; text-align:left;"> <%= sTeamStatusText %></span>	  
	</div>
<%	

END SUB


' --------------------------
  SUB DisplayTeamTab_Simple
' --------------------------

%> 
  <div class="tabrankings" style="width:96%; background-color:<%=TabColor%>; height:35px; margin-top:2px; padding-top:0px;" >
		<span class="span80" style="font-size:12pt; color:#000000; font-weight:bold;"><%= sTeam_Name %></span>
	  <br>
  	<span class="span20" style="color:black; font-weight:normal; font-size:10pt; border:0px solid red;">Members: </span>
  	<span class="span10" style="color:blue; font-weight:normal; font-size:10pt; text-align:left; border:0px solid red;"><%= sNo_Team_Members %></span>
  	<span class="span10" style="color:black; text-align:right; font-size:10pt; font-weight:normal; border:0px solid red;">ID: </span>
	  <span class="span15" style="color:blue; font-weight:normal; font-size:10pt; text-align:left; border:0px solid red;"><%= sTeam_ID %></span>
	  <span class="span15" style="font-weight:normal; color:black; font-size:10pt; text-align:right; border:0px solid red;">Display: </span>	  
	  <span class="span20" style="color:<%=sTeamStatusTextColor%>; font-size:10pt; font-weight:normal; text-align:left; border:0px solid red;"> <%= sTeamStatusText %></span>	  
	</div>
<%	

END SUB


' ---------------------------
  SUB DisplayTeamMemberLine
' ---------------------------

sInviteName = TRIM(sFirstName)&" "&TRIM(sLastName)
sInviteMemberID = sMemberID
sInviteCityState = sMembCity&" "&sMembState
sInviteEmail = sEmail



sThisJavaScriptLine = "'"&sInviteMemberID&"','"&sInviteName&"','"&sInviteCityState&"','"&sInviteEmail&"','"&sTeam_ID&"'"
'response.write("<br></div><div> sThisJavaScriptLine = "&sThisJavaScriptLine)
'response.end
%>
  <div class="rankingsbody" style="width:96%; height:18px; padding-top:15px; font-size:10pt;">
		<span class="span50" style="color:black; font-weight:normal;"><%= sFirstName %>&nbsp;<%= sLastName %></span>
		<span class="span15" style="color:black; font-weight:normal;"><%= sMembSexText %></span>
		<%
			IF TRIM(sEmail)<>"" AND ( TRIM(sTeamMemberStatus)="I" OR TRIM(sTeamMemberStatus)="" OR IsNull(sTeamMemberStatus) ) THEN
					InviteButtonText = "Send Invite"
					InviteButtonBackgroundColor = "orange"
					IF TRIM(sTeamMemberStatus)="I" THEN 
							InviteButtonText = "Resend Invite"
							InviteButtonBackgroundColor = "yellow"
					END IF		
					%>
					<span class="span25" style="color:<%=sTeamMemberStatusTextColor%>; font-weight:normal; text-align:center;">
						<input type="button" style="height:1.75em; width:8em; font-size:9pt; background-color:<%=InviteButtonBackgroundColor%>; padding:0px 0px 0px 0px; margin:0px 0px 0px 0px;" value="<%=InviteButtonText%>" onclick="javscript:TeamConfirmInviteNav('ToConfirmSendInviteFromMyTeamList',<%=sThisJavaScriptLine%>)">
					</span>
					<%
			ELSE		
					%><span class="span20" style="color:<%=sTeamMemberStatusTextColor%>; font-weight:normal;"><%= sTeamMemberStatus %></span><%
			END IF		
		%>
	</div>
<%
	

END SUB  



' ---------------------------------
  SUB DisplayTeamMemberLine_Simple
' ---------------------------------

sInviteName = TRIM(sFirstName)&" "&TRIM(sLastName)
sInviteMemberID = sMemberID
sInviteCityState = sMembCity&" "&sMembState
sInviteEmail = sEmail



'InviteButtonText = "Send Invite"
'InviteButtonBackgroundColor = "orange"
'IF TRIM(sTeamMemberStatus)="I" THEN 
'		InviteButtonText = "Resend Invite"
'		InviteButtonBackgroundColor = "yellow"
'END IF		


sThisJavaScriptLine = "'"&sInviteMemberID&"','"&sInviteName&"','"&sInviteCityState&"','"&sInviteEmail&"','"&sTeam_ID&"'"
'response.write("<br></div><div> sThisJavaScriptLine = "&sThisJavaScriptLine)
'response.end
' =sTeamMemberStatusTextColor

%>
  <div class="rankingsbody" style="width:96%; height:18px; padding-top:5px; font-size:12pt;">
		<span class="span50" style="color:black; height:16px; font-weight:normal; border:0px solid red; padding:0px; margin:0px;"><%= sFirstName %>&nbsp;<%= sLastName %></span>
		<span class="span10" style="color:black; height:16px; font-weight:normal; border:0px solid yellow; padding:0px; margin:0px;"><%= sMembSexText %></span>
		<span class="span30" style="color:<%=sTeamMemberStatusTextColor%>; height:16px; font-weight:normal; border:0px solid green; padding:0px; margin:0px;"><%= sTeamMemberStatusText %></span>
	</div>
<%
	

END SUB  





' ---------------------------
  SUB DisplayTeamBottomLine
' ---------------------------  

%>
<div class="tourbottom"  style="background-color:#FFFFFF; height:7px;">
		<span class="span95">&nbsp;</span>
</div>
<%

END SUB





' --------------------------
  SUB GetCurrentMemberLine
' --------------------------

    sTeam_ID=rs("Team_ID")
		sTeam_Name=rs("Team_Name")
		sTeam_Level=rs("Team_Level")
		sCreated_Date=rs("Created_Date")
		sTeamStatus=rs("TeamStatus")
		sManager_MemberID=rs("Manager_MemberID")
		sManager_FirstName=rs("Manager_FirstName")
		sManager_LastName=rs("Manager_LastName")

		sMemberID=rs("MemberID")
		sFirstName=rs("FirstName")
		sLastName=rs("LastName")
		sMembCity=rs("City")
		sMembState=rs("State")		
		
		' -- Remap Sex --
		sMembSex=rs("Sex")
		sMembSexText="M"
		IF sMembSex="Female" THEN sMembSexText="F"
		
		sEffectiveTo=rs("EffectiveTo")

		sTeamMemberStatus=rs("TeamMemberStatus")
		sTeam_Type_ID=rs("Team_Type_ID")
		sTeam_Type_Description=rs("Team_Type_Description")
		sNo_Team_Members = rs("No_Team_Members")
		sEmail = rs("Email")
		sEmail = "mark@productdesign-biz.com"
		
		' --- Elements from Team_Type table
		sMax_Members = rs("Max_Members")
		sMin_Members = rs("Min_Members")
		
		
		
		SELECT CASE TRIM(sTeamMemberStatus)
				CASE "A" 
						sTeamMemberStatusText="Accepted"
						sTeamMemberStatusTextColor="green"
				CASE "I"
						sTeamMemberStatusText="Pending"			
						sTeamMemberStatusTextColor="orange"
				CASE "D"
						sTeamMemberStatusText="Declined"			
						sTeamMemberStatusTextColor="red"
				CASE ELSE
						sTeamMemberStatusText="Need Invite"			
						sTeamMemberStatusTextColor="red"
		END SELECT			

		SELECT CASE TRIM(sTeamStatus)
				CASE "A" 
						sTeamStatusText="Active"
						sTeamStatusTextColor="green"
				CASE "P"
						sTeamStatusText="Pending"			
						sTeamStatusTextColor="orange"
				CASE "H"
						sTeamStatusText="Inactive"			
						sTeamStatusTextColor="red"
				CASE ELSE
						sTeamStatus="U"
						sTeamStatusText="Unknown"			
						sTeamStatusTextColor="red"
		END SELECT			

		SELECT CASE TRIM(sTeam_Type_ID)
				CASE 1 
						TabColor=scolor01
				CASE 2 
						TabColor=scolor02
				CASE 3 
						TabColor=scolor03
				CASE 4 
						TabColor=scolor04
				CASE 5 
						TabColor=scolor05
				CASE 6 
						TabColor=scolor06
				CASE 7 
						TabColor=scolor07
				CASE 8 
						TabColor=scolor08
				CASE 9 
						TabColor=scolor09
				CASE 10 
						TabColor=scolor10
				CASE ELSE
						TabColor=scolor05							
		END SELECT
END SUB





' -----------------------
  SUB DetermineTeamStatus
' -----------------------

sTeamStatus = "League Requires 3 Members"
AddTeamMemberButtonStatus="enabled"
' IF sNo_Team_Members = Team_Type_Max_Members THEN AddTeamMemberButtonStatus="disabled"

AddTeamMemberButtonStatus="disabled"
DeleteTeamMemberButtonStatus="disabled"

This_Team_ID=1002

sSQL = sSQL + " SELECT tsum.Team_ID, This_TeamMember_Count, This_Female_Count, This_Male_Count, This_Accepted, This_Invited, This_Needs_Invite"
sSQL = sSQL + " , t.*, tt.*"
sSQL = sSQL + " FROM"
sSQL = sSQL + " 	( SELECT Team_ID, COUNT(*) AS This_TeamMember_Count"
sSQL = sSQL + " 		, SUM(CASE WHEN LOWER(Sex)='female' THEN 1 ELSE 0 END) AS This_Female_Count"
sSQL = sSQL + " 		, SUM(CASE WHEN LOWER(Sex)='male' THEN 1 ELSE 0 END) AS This_Male_Count"
sSQL = sSQL + " 		, SUM(CASE WHEN Status='A' THEN 1 ELSE 0 END) AS This_Accepted"
sSQL = sSQL + " 		, SUM(CASE WHEN Status='I' THEN 1 ELSE 0 END) AS This_Invited"
sSQL = sSQL + " 		, SUM(CASE WHEN Status NOT IN ('A','I') THEN 1 ELSE 0 END) AS This_Needs_Invite"
sSQL = sSQL + " 		FROM "&V_TeamMembersTableName&" tmem"
sSQL = sSQL + " 		LEFT JOIN "&MemberTableName&" m ON RIGHT(tmem.MemberID,8)=m.PersonID"
sSQL = sSQL + "				WHERE tmem.Team_ID = "&This_Team_ID
sSQL = sSQL + "				GROUP BY Team_ID"
sSQL = sSQL + "		)	tsum"

sSQL = sSQL + " JOIN "&V_TeamTableName&" t ON t.Team_ID=tsum.Team_ID"
sSQL = sSQL + " JOIN "&V_TeamTypeTableName&" tt ON tt.Team_Type_ID=t.Team_Type_ID"

mt=2
IF mt=1 THEN
		%></div><div style="color:red;"><%
		response.write(sSQL)
		response.end
END IF

SET rsTT=Server.CreateObject("ADODB.recordset")
rsTT.open sSQL, SConnectionToTRATable



' --- Need DIM ***
sThis_TeamMember_Count = rsTT("This_TeamMember_Count")
sThis_Female_Count = rsTT("This_Female_Count")
sThis_Male_Count = rsTT("This_Male_Count")
sThis_Accepted = rsTT("This_Accepted")
sThis_Invited = rsTT("This_Invited")
sThis_Needs_Invite = rsTT("This_Needs_Invite")


sTeam_Type_Description = rsTT("Team_Type_Description")
sMax_Members = rsTT("Max_Members")
sMin_Members = rsTT("Min_Members")
sMax_Male = rsTT("Max_Male")
sMin_Male = rsTT("Min_Male")
sMax_Female = rsTT("Max_Female")
sMin_Female = rsTT("Min_Female")
sMax_Age = rsTT("Max_Age")
sMin_Age = rsTT("Min_Age")


sTeamStatus = ""
AddTeamMemberButtonStatus="enabled"
DeleteTeamMemberButtonStatus="enabled"

sThisTeam_Remaining_Total = sMax_Members - sThis_TeamMember_Count
sThisTeam_Remaining_Male = sMax_Male - sThis_Male_Count
sThisTeam_Remaining_Female = sMax_Female - sThis_Female_Count


IF sThis_TeamMember_Count = Max_Members THEN
		AddTeamMemberButtonStatus="disabled"				' --- Max team members reached so disable Add button
END IF
IF sThis_TeamMember_Count = 0 THEN
		DeleteTeamMemberButtonStatus="disabled"			' --- No team members so disable Delete button
END IF
IF sThis_Invited>0 THEN sTeamStatus = sTeamStatus + "<br>Some team members must accept invite" 
IF sThis_Needs_Invite>0 THEN sTeamStatus = sTeamStatus + "<br>Invitation needed for members" 


IF TRIM(LEN(sTeamStatus))=0 THEN sTeamStatus="<br>Team meets minimum requirements"


' -- Converting to 9 position MemberID ---
' PersonIDwChkDgt(rs("PersonID"))



rsTT.close


END SUB  






' ---------------------------
  SUB RunMyTeamListingQuery
' ---------------------------

'response.write("</div><div style=color:yellow;>TeamIDSelected = "&TeamIDSelected)
'sThis_Team_ID=1002

sSQL = "SELECT t.Team_ID, Team_Name, t.Team_Type_ID, Team_Level, Manager_MemberID, t.Created_Date"
sSQL = sSQL + ", m.FirstName, m.LastName, tmem.MemberID"
sSQL = sSQL + ", m2.FirstName AS Manager_FirstName, m2.LastName AS Manager_LastName, m.Email"
sSQL = sSQL + ", m.City, m.State, m.Sex, m.Birthdate, m.EffectiveTo"
sSQL = sSQL + ", tmem.Status AS TeamMemberStatus"
sSQL = sSQL + ", t.Status AS TeamStatus, Team_Type_Description, tcnt.No_Team_Members"
sSQL = sSQL + ", tt.*"

sSQL = sSQL + " FROM "&V_TeamMembersTableName&" tmem"
sSQL = sSQL + " JOIN "&V_TeamTableName&" t ON t.Team_ID=tmem.Team_ID"
sSQL = sSQL + " JOIN "&V_TeamTypeTableName&" tt ON tt.Team_Type_ID=t.Team_Type_ID"
sSQL = sSQL + " LEFT JOIN "&MemberTableName&" m ON RIGHT(tmem.MemberID,8)=m.PersonID"
sSQL = sSQL + " LEFT JOIN "&MemberTableName&" m2 ON RIGHT(t.Manager_MemberID,8)=m2.PersonID"

sSQL = sSQL + " LEFT JOIN "
sSQL = sSQL + "   ( SELECT Team_ID, COUNT(*) AS No_Team_Members FROM "&V_TeamMembersTableName
sSQL = sSQL + "      GROUP BY Team_ID ) tcnt"  
sSQL = sSQL + " ON tcnt.Team_ID=tmem.Team_ID"

IF TeamIDSelected <> 0 THEN
		sSQL = sSQL + " WHERE tmem.Team_ID='"&TeamIDSelected&"'"
END IF

sSQL = sSQL + " ORDER BY t.Team_Type_ID, tmem.Team_id, m.LastName, m.FirstName"

mt=2
IF mt=1 THEN
		%></div><div style="color:red;"><%
		response.write(sSQL)
		'response.end
END IF

SET rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable


END SUB  



' ----------------------------------------------------------------
  SUB BuildTeamListDropDown (thiswidth,thisfont,onchangeaction)
' ----------------------------------------------------------------

sManager_MemberID = "000001151"

sSQL = "SELECT Team_ID, Team_Name"
sSQL = sSQL + " FROM "&V_TeamTableName
sSQL = sSQL + " WHERE Manager_MemberID = '"&sManager_MemberID&"'"

SET rsTT=Server.CreateObject("ADODB.recordset")
rsTT.open sSQL, SConnectionToTRATable

'response.write("</span></div><div style=""color:white;"">onchangeaction = "&onchangeaction&"</div>")

%>
<SELECT id="TeamIDSelected" name="TeamIDSelected" style="width:<%=thiswidth%>em; font-size:<%=thisfont%>px;" onchange=<%=onchangeaction%>;>
<option value=0>Select Your Team</option><% 

		DO WHILE NOT rsTT.eof 
						IF CStr(rsTT("Team_ID")) = CStr(TeamIDSelected) THEN 
								%><option value = "<%=rsTT("Team_ID")%>" selected><%= rsTT("Team_ID") %> - <%= rsTT("Team_Name") %></option><%
						ELSE
								%><option value = "<%=rsTT("Team_ID")%>"><%= rsTT("Team_ID") %> - <%= rsTT("Team_Name") %></option><%
						END IF	
				rsTT.movenext
  	LOOP 

%></SELECT><%

rsTT.close
		
	


END SUB



' ------------------------------
  SUB RunTeamTypeParametersQuery
' ------------------------------  

sMax_Members="--"
sMin_Members="--"
sMax_Male="--"
sMin_Male="--"
sMax_Female="--"
sMin_Female="--"
sMax_Age="--"
sMin_Age="--"

sMax_Scoring="--"
sMin_Scoring="--"
sMax_Scoring_Male="--"
sMin_Scoring_Male="--"
sMax_Scoring_Female="--"
sMin_Scoring_Female="--"
sTeam_Type_Description="--"


sSQL = "SELECT *" 
sSQL = sSQL + " FROM "&V_TeamTypeTableName 
sSQL = sSQL + " WHERE Team_Type_ID = "&TeamTypeIDSelected
sSQL = sSQL + " ORDER BY Team_Type_ID_Seq"

'response.write("</div><div>"&sSQL&"</div>")
'response.end

set rsTT=Server.CreateObject("ADODB.recordset")
rsTT.open sSQL, SConnectionToTRATable, 3, 3 

IF NOT rsTT.eof THEN 
		sTeam_Type_Description = rsTT("Team_Type_Description")
		sMax_Members = rsTT("Max_Members")
		sMin_Members = rsTT("Min_Members")
		sMax_Male = rsTT("Max_Male")
		sMin_Male = rsTT("Min_Male")
		sMax_Female = rsTT("Max_Female")
		sMin_Female = rsTT("Min_Female")
		sMax_Age = rsTT("Max_Age")
		sMin_Age = rsTT("Min_Age")

		sMax_Scoring = rsTT("Max_Scoring")
		' sMin_Scoring = rsTT("Min_Scoring")
		' sMax_Scoring_Male = rsTT("Max_Scoring_Male")
		sMin_Scoring_Male = rsTT("Min_Scoring_Male")
		' sMax_Scoring_Female = rsTT("Max_Scoring_Female")
		sMin_Scoring_Female = rsTT("Min_Scoring_Female")
END IF		

END SUB




' ---------------------------------
  SUB SendEmailinviteToTeamMember
' ---------------------------------



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

sInviteBannerText = "Please Join My Team"
sInviteName = Request("sInviteName")
sInviteMemberID = Request("sInviteMemberID")
sInviteEmail = Request("sInviteEmail")
sTeam_ID = Request("sTeam_ID")

sSQL = "SELECT t.Team_ID, t.Team_Name"
sSQL = sSQL + ", Manager_MemberID"
sSQL = sSQL + ", m2.FirstName AS Manager_FirstName, m2.LastName AS Manager_LastName, m2.Email"
sSQL = sSQL + ", tt.Team_Type_Description"
sSQL = sSQL + " FROM "&V_TeamTableName&" t"
sSQL = sSQL + " JOIN "&V_TeamTypeTableName&" tt ON tt.Team_Type_ID=t.Team_Type_ID"
sSQL = sSQL + " LEFT JOIN "&MemberTableName&" m2 ON RIGHT(t.Manager_MemberID,8)=m2.PersonID"
sSQL = sSQL + " WHERE Team_ID = "&sTeam_ID 

set rsTT=Server.CreateObject("ADODB.recordset")
rsTT.open sSQL, SConnectionToTRATable, 3, 3 

IF NOT rsTT.eof THEN
		sThisTeamTypeDescription = rsTT("Team_Type_Description")
		sThisTeamName = rsTT("Team_Name")
		sThisTeamManagerName = TRIM(rsTT("Manager_FirstName"))&" "& TRIM(rsTT("Manager_LastName"))
		sThisTeamTypeDescription = rsTT("Team_Type_Description")
END IF

rsTT.close

eMailTo = sInviteEmail
eMailCC = ""
eMailBCC = ""
eMailFrom = "competition@usawaterski.org"
eMailSubj = "Your Invitation to Join My Team"





' --- Create Email message ---
ebody = ecss & "<html>"
ebody = ebody & "<head>"
ebody = ebody & "<title>Invite to Join My Team</title>"
ebody = ebody & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
ebody = ebody & "</head>"
ebody = ebody & "<body bgcolor=""#FFFFFF"" text=""#000000"">"

ebody = ebody & "<div class=outer>"

ebody = ebody & "<div style="&SQT&"text-align:center;"&SQT&">"
ebody = ebody & "<img style='width:300px;' name=BannerLogo src='http://usawaterski.com/rankings/images/General/JoinTheFun.JPG' alt=Accept_Join>"
ebody = ebody & "</div>"
ebody = ebody & "<div style="&SQT&"margin-top:20px; text-align:center; font-size:32px; color:red;"&SQT&"><i>Please Join My Team</i></div>"

ebody = ebody & "<div class=pblack style="&SQT&"margin-top:20px;"&SQT&">Hi "&sInviteName&":</div>"
ebody = ebody & "<div class=pblack style=""margin-top:20px;"">"
ebody = ebody & "  I am building a waterski team using the new mobile app from the <b>American Water Ski Association</b>. This system uses your real scores together with scores of other team members to establish a Team Ranking.  The ranking is based on each member's improvement throughout the year."
ebody = ebody & "</div>"

ebody = ebody & "<div class=pblack style=""margin-top:20px;"">"
ebody = ebody & "  To join my team you have to accept my invitation so my team can become active."
ebody = ebody & "  Once everyone I have invited has accepted my invitation, the team will appear under the Team Rankings in a League called <b>"&sThisTeamTypeDescription&"</b>. We will be competing against other teams in the same league throughout the year."  
ebody = ebody & "</div>"
ebody = ebody & "<div style=margin-top:15px; text-align:center;><span class=pblack style=font-size:14pt;>Team Name:&nbsp;</span><br><span class=pblue style=font-size:16pt;>"&sThisTeamName&"</span></div>"
ebody = ebody & "<div style='margin-top:5px; text-align:center;' ><span class=pblack>Team ID: </span><span class=pblue>"&sTeam_ID&"</span></div>"

ebody = ebody & "<div class=pblack style='margin-top:15px; text-align:center;'>To <b>Accept</b> and be part of my team, click below</div>"
ebody = ebody & "<div class=pblack style=""margin-top:10px; text-align:center;"">"
ebody = ebody & "      <table align=""center"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
ebody = ebody & "        <tr>"
ebody = ebody & "          <td class=psuedobuttoncellgreen><a href='http://usawaterski.org/rankings/vteams_manage.asp?action=acceptinvite&team_id="&sTeam_ID&"&sMemberID="&sInviteMemberID&"' target='_blank' class=psuedobuttongreen>Join My Team</a></td>"
ebody = ebody & "        </tr>"
ebody = ebody & "      </table>"
ebody = ebody & "</div>"

ebody = ebody & "<div class=pblack style='margin-top:15px; text-align:center;'>To <b>Decline</b> participation with this team, click below.</div>"
ebody = ebody & "<div class=pblack style='margin-top:10px; text-align:center;'>"
ebody = ebody & "      <table align=""center"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
ebody = ebody & "        <tr>"
ebody = ebody & "          <td class=psuedobuttoncellred><a href='http://usawaterski.org/rankings/vteams_manage.asp?action=declineinvite&team_id="&sTeam_ID&"&sMemberID="&sInviteMemberID&"' target='_blank' class=psuedobuttonred>Decline Invitation</a></td>"
ebody = ebody & "        </tr>"
ebody = ebody & "      </table>"
ebody = ebody & "</div>"

ebody = ebody & "<div class=pblack style='margin-top:20px; text-align:center;'>"
ebody = ebody & "I am looking forward to having you on my team."
ebody = ebody & "<br><b>"&sThisTeamManagerName&"</b>"
ebody = ebody & "</div>"

ebody = ebody & "<div class=pblack style='text-align:center; margin-top:15px; padding-bottom;30px'>Click the image below from your phone to try out the new AWSA mobile App."
ebody = ebody & "<br>" 
ebody = ebody & " <a href='http://usawaterski.org/rankings/mainmenu_m.asp' style=""text-decoration:none;"">"
ebody = ebody & "  <img style=""width:57px;"" name=""MobileAppIcon"" src=""http://www.usawaterski.com/rankings/images/icons/AWSA_HomeScreen_57.PNG"" alt=""Mobile App"">"
ebody = ebody & " </a>"
ebody = ebody & "</div>"

ebody = ebody & "</div>"
ebody = ebody & "<br><br><br>"
ebody = ebody & "</body></html>"

eMailBody = ebody

'response.write("</div><br>eMailTo = "&eMailTo&"<br>eMailCC = "&eMailCC&"<br>eMailBCC = "&eMailBCC&"<br>eMailFrom = "&eMailFrom&"")
'response.write("<br>eMailSubj = "&eMailSubj&"<br><br>"&eMailBody)
'response.end

SendEmailFromGenericMethod eMailTo,eMailCC,eMailBCC,eMailFrom,eMailSubj,eMailBody

%>
<div class="errorbox" id="EmailSentConfirmationScreen" style="inline-block; margin-top:5px;">
	<form action="<%=ThisFileName%>?action=teamlist" method="post">
		<input type="hidden" name="Team_ID" value="<%=Team_ID%>">
		<div style="margin-top:10px; border:0px solid white;">
  		<span class="span100" style="text-align:center; color:#FFFFFF; font-size:14pt; padding-left:10px;">Email Invitation Sent</span>
		</div>
		
		<div style="width:100%; margin-top:20px; color:yellow; padding-left:10px;">Your invitation to join has been sent to </div>	
		<div style="width:100%; margin-top:20px; padding-left:10px;">		
			<span class="span95" style="text-align:center;">
				<input type="text" class="textbox_hidden_banner" id="sInviteName" name="sInviteName" value="<%=sInviteName%>" style="text-align:center; color:yellow;" MaxLength="25">
			</span>
		</div>
		<div style="margin-top:35px; border:0px solid white;">
  		<span class="span100" style="text-align:center; color:#FFFFFF; font-size:10pt; padding-left:10px;">Click below to continue managing this team.</span>
		</div>
		<div id="ConfirmEmailSentButtons" class="menucell" style="padding:0px; margin-top:55px;">
			<TABLE align=center width=100% style="padding:0px; margin:0px; border:0px solid; border-color:yellow;">
				<tr>
				<td width="46%" align="center">
					<input type="button" name="ReturnToTeamList" value="Continue" style="width:8em; height:2em; font-size:10pt;">
				</td>				
			</tr>
			</TABLE>
		</div>
	</form>	
</div>  
<%	


END SUB



  




' -------------------------
  SUB UpdateFromInvitation 
' -------------------------  

sTeam_ID = Request("team_id")
sMemberID = Request("sMemberID")



' --- Get Team Manager information ---
sSQL = "SELECT t.Team_ID, t.Team_Name"
sSQL = sSQL + ", Manager_MemberID"
sSQL = sSQL + ", m2.FirstName AS Manager_FirstName, m2.LastName AS Manager_LastName, m2.Email AS Manager_Email"
sSQL = sSQL + ", tt.Team_Type_Description"

sSQL = sSQL + " FROM "&V_TeamTableName&" t"
sSQL = sSQL + " LEFT JOIN "&MemberTableName&" m2 ON RIGHT(t.Manager_MemberID,8)=m2.PersonID"
sSQL = sSQL + " JOIN "&V_TeamTypeTableName&" tt ON tt.Team_Type_ID=t.Team_Type_ID"
sSQL = sSQL + " WHERE Team_ID = "&sTeam_ID

'response.write("</div><div>sQL = "&sSQL&"</div>")


set rsTT=Server.CreateObject("ADODB.recordset")
rsTT.open sSQL, SConnectionToTRATable, 3, 3 

IF NOT rsTT.eof THEN 
		sThisTeamName = rsTT("Team_Name")
		sThisTeamManagerName = TRIM(rsTT("Manager_FirstName"))&" "& TRIM(rsTT("Manager_LastName"))
		sThisTeamTypeDescription = rsTT("Team_Type_Description")
		sManager_Email = rsTT("Manager_Email")
END IF	

rsTT.close




' --- Update the record based on the response --
SELECT CASE action
		CASE "acceptinvite"
				InviteValue="A"
				InviteConfirmationText = "You have Accepted the invitation from"
		CASE "declineinvite"
				InviteValue="D"
				InviteConfirmationText = "You have Decline the invitation from"
END SELECT	

sSQL = "UPDATE "&V_TeamMembersTableName
sSQL = sSQL + " SET Status='"&InviteValue&"'" 
sSQL = sSQL + " WHERE MemberID='"&sMemberID&"' AND Team_ID='"&sTeam_ID&"'"

'response.write("<div>sQL = "&sSQL&"</div>")

OpenCon
con.execute(sSQL)
CloseCon




' --- Now get Member information for response to Manager ---
sSQL = "SELECT FirstName, LastName"
sSQL = sSQL + " , Team_Name, Team_Type_Description"
sSQL = sSQL + " FROM "&V_TeamMembersTableName&" tmem"
sSQL = sSQL + " JOIN "&V_TeamTableName&" t ON t.Team_ID=tmem.Team_ID"
sSQL = sSQL + " LEFT JOIN "&MemberTableName&" m ON RIGHT(tmem.MemberID,8)=m.PersonID"
sSQL = sSQL + " JOIN "&V_TeamTypeTableName&" tt ON tt.Team_Type_ID=t.Team_Type_ID"
sSQL = sSQL + " WHERE tmem.MemberID='"&sMemberID&"' AND tmem.Team_ID='"&sTeam_ID&"'"

'response.write("</div><div>sSQL = "&sSQL&"</div>")
'response.end

set rsTT=Server.CreateObject("ADODB.recordset")
rsTT.open sSQL, SConnectionToTRATable, 3, 3 

IF NOT rsTT.eof THEN 
		sFullName = TRIM(rsTT("FirstName"))&" "&rsTT("LastName")
		sThisTeamTypeDescription = rsTT("Team_Type_Description")
		sThisTeamName = rsTT("Team_Name") 
END IF	

rsTT.close

' --- Now send confirmation email to Manager about that the Team Member did
'OrgFromEmailAddress&" "&OrgFriendlyFrom

sInvitationResponse="Declined"
IF InviteValue="A" THEN sInvitationResponse="Accepted"

eMailTo = sManager_Email
eMailCC = ""
eMailBCC = ""
'eMailFrom = "competition@usawaterski.org AWSA Team Building App"
eMailFrom = "competition@usawaterski.org"
eMailSubj = sFullName& " Has "&sInvitationResponse&" Your Invitation"

ConfirmAcceptOrRejectionByMemberToManagerEmail eMailTo,eMailCC,eMailBCC,eMailFrom,eMailSubj,sInvitationResponse


%>
<div class="errorbox" id="UpdateInviteScreen" style="inline-block; margin-top:5px;">
	<form action="<%=MenuFileName%>" method="post">
		<div style="margin-top:10px; border:0px solid white;">
  		<span class="span100" style="text-align:center; color:#FFFFFF; font-size:14pt; padding-left:10px;">Your Record Has Been Updated</span>
		</div>
		
		<div style="width:100%; margin-top:20px; color:yellow; padding-left:10px;"><%=InviteConfirmationText%></div>	
		<div style="width:100%; margin-top:10px; padding-left:10px;">		
			<span class="span95" style="text-align:center;"><%=sThisTeamManagerName%></span>
		</div>
		

		<div style="width:100%; margin-top:5px; padding-left:10px;">		
			<span class="span95" style="text-align:center;"><%=sThisTeamName%></span>
			<br>
			<span class="span95" style="text-align:center;"><%=sThisTeamTypeDescription%></span>
		</div>
		<div style="margin-top:20px; border:0px solid white;">
  		<span class="span100" style="text-align:center; color:#FFFFFF; font-size:10pt; padding-left:10px;">Click below to link to the mobile app.</span>
		</div>
		<div id="LinkToAppButton" class="menucell" style="padding:0px; margin-top:55px;">
				<input type="submit" style='width:12em; height:3em; background-color:#DC143C;' value='Visit Mobile App'>
		</div>
	</form>	
</div>  
<%	

END SUB





'--------------------------------------------------------------------------------------------------------------------------  
  SUB ConfirmAcceptOrRejectionByMemberToManagerEmail ( eMailTo,eMailCC,eMailBCC,eMailFrom,eMailSubj,sInvitationResponse )
'--------------------------------------------------------------------------------------------------------------------------  

' --- Create Email message ---

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

sResponseAdjective="Great"
sResponseColor="green"
sResponseIcon="http://usawaterski.com/rankings/images/icons/CheckYes.PNG"
IF sInvitationResponse="Declined" THEN 
		sResponseAdjective="Sorry"
		sResponseColor="red"
		sResponseIcon="http://usawaterski.com/rankings/images/icons/XNo.PNG"
END IF

ebody = ecss & "<html>"
ebody = ebody & "<head>"
ebody = ebody & "<title>Invitation Response</title>"
ebody = ebody & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
ebody = ebody & "</head>"
ebody = ebody & "<body bgcolor=""#FFFFFF"" text=""#000000"">"

ebody = ebody & "<div class=outer>"

ebody = ebody & "<div style="&SQT&"text-align:center;"&SQT&">"
ebody = ebody & "<img name=USAWaterskiLogo src=""http://usawaterski.com/rankings/images/logos/usawslogo_no_sub.jpg"" alt=""USA Waterski"" >"
ebody = ebody & "</div>"
ebody = ebody & "<div style=""margin-top:20px; text-align:center; font-size:24pt; color:"&sResponseColor&";""><i>Invitation "&sInvitationResponse&"</i></div>"
ebody = ebody & "<div style="&SQT&"text-align:center;"&SQT&">"
ebody = ebody & "<img name=ResponseIcon src="""&sResponseIcon&""" alt=""ResponseIcon"" >"
ebody = ebody & "</div>"


ebody = ebody & "<div class=pblack style=""margin-top:20px; text-align:center;"">"
ebody = ebody & ""&sResponseAdjective&", <b>"&sFullName&"</b> has <u><i>"&sInvitationResponse&"</i></u> your invitation to become a member of your team in the <b>"&sThisTeamTypeDescription&"</b> league." 
IF sInvitationResponse="Declined" THEN ebody = ebody & " Please select another member to fill your team."
ebody = ebody & "</div>"
ebody = ebody & "<div style=""margin-top:15px; text-align:center;""><span class=pblack style=""font-size:14pt;"">Team Name:&nbsp;</span><br><span class=pblue style=""font-size:16pt;"">"&sThisTeamName&"</span></div>"
ebody = ebody & "<div style='margin-top:5px; text-align:center;' ><span class=pblack>Team ID: </span><span class=pblue>"&sTeam_ID&"</span></div>"

ebody = ebody & "<div class=pblack style='margin-top:15px; text-align:center;'>To manage your teams from your Mobile Device, click below.</div>"
ebody = ebody & "<div class=pblack style='margin-top:10px; text-align:center;'>"
ebody = ebody & " <a href='http://usawaterski.org/rankings/mainmenu_m.asp' style=""text-decoration:none;"">"
ebody = ebody & "  <img style=""width:57px;"" name=""MobileAppIcon"" src=""http://www.usawaterski.com/rankings/images/icons/AWSA_HomeScreen_57.PNG"" alt=""Mobile App"">"
ebody = ebody & " </a>"
ebody = ebody & "</div>"

ebody = ebody & "</div>"
ebody = ebody & "<br><br><br>"
ebody = ebody & "</body></html>"

eMailBody = ebody

'response.write("</div><div>"&eMailBody&"</div>")

'response.write("</div><br>eMailTo = "&eMailTo&"<br>eMailCC = "&eMailCC&"<br>eMailBCC = "&eMailBCC&"<br>eMailFrom = "&eMailFrom&"")
'response.write("<br>eMailSubj = "&eMailSubj&"<br><br>"&eMailBody)
'response.end


SendEmailFromGenericMethod eMailTo,eMailCC,eMailBCC,eMailFrom,eMailSubj,eMailBody




END SUB


%>
