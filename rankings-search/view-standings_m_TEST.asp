<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_leagues.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->
<!--#include virtual="/rankings/tools_registration16.asp"-->
<!--#include virtual="/rankings/Tools_TournamentListQuery.asp"-->
<!--#include virtual="/rankings/tools_mobile_version_TEST.asp"-->
<%



' ------------------------------------------------
' --- Dimensions variables used in this module ---
' ------------------------------------------------

Dim rowCount, i
Dim MemoryScore, MemoryPlc, MemoryRank, RecordNum, RankValueWithTies
Dim tRankScore, tRankPct, tFmtScore, tRnkScoBkup, tMemberID
DIM tOpenDiv, tMastDiv, tEliteStat, tEliteBkup, tRating
Dim tName, tState, tRegion, tNatPlace, tRegSki, tRegPlace, tMemberFed
Dim tTeam, tTeamStat, tNCWRegn, tNCWConf, tDefaultYear, LastRegion
Dim DefineRowcolor, DefineLevelcolor, LastLevelcolor, tPerc3, tPerc4, tPerc5, tPerc6, tPerc7, tPerc8, tPerc9
Dim tLevelNo, LastLevelNo, LastDiv
Dim tBirth, tAge, tBirthday, RankingsRecalcUnderway
Dim sMemberID, sFullName, sEventName
Dim TourDisplaywidth, ScorePageBorderDark, ScorePageBorderLight
Dim MainImage
Dim COALevel9, COALevel8, COALevel7, COALevel6, COALevel5, COALevel4, COALevel3
Dim sDefaultNationals
Dim sShowSQL, sRunByWhat, RankingListType
Dim DisplayFilters
Dim ThisFileName



' --- Added 5/25/2015 ---
Dim EventSelected, RankingsDivSelected, RankingsStateRegionSelected, FederationSelected, SkiYearSelected, RankingsSkiYearIDSelected
' Dim sLeagueSelected







ThisFileName="view-standings_m.asp"

' --- Names related programs for linking ---
SearchFileName = "search-memberHQ.asp"
RankingsMobileFilename=ThisFileName 
TournamentsMobileFilename="view-tournaments_m.asp"
LocalVarFileName="Test_localstorage_SET.asp"
TeamsMobileFilename="virtualteam_m.asp"
MenuFileName = "mainmenu_m.asp"


tLevelNo = 999999
LastLevelNo = 999998
LastDiv = "ABC"

TourDisplaywidth=725
ScorePageBorderDark = HQSitecolor1
ScorePageBorderLight = HQSitecolor2


IF TRIM(Session("NewRankVis"))="" THEN
		KickTrafficCounter("NewRankVis")	
		Session("NewRankVis")="YES"
END IF




' --- This is a TEMPORARY fix.  Actual cut-off logic needs to be dynamic.
CutOffDate = "07/23/2008"


' ------------------------------------
' --- Reads NVP's from querystring ---
' ------------------------------------
ReadRankingsFormVariables




' --- Finds the name of the current sMemberID ---
DetermineNameOfCurrentMember




RecordNum = TRIM(Request("RecordNum"))    
IF RecordNum = "" THEN RecordNum = 1



' --------------------------------------------------------------------------
' --- Defines the image to be displayed in the drop downs box background ---
' --------------------------------------------------------------------------
WhatDropdownImage EventSelected



' --- Displays the html, head and opening body tag ---
OpenState="rankings"
DisplayHeadOpenBodyAndBannerTags OpenState




' ------------------------------------------------------------------------------------------------           
' -------------------------------  BEGINS WRITING HEADERS AND RANKINGS  --------------------------
' ------------------------------------------------------------------------------------------------

	

' --- Displays the menu for view tournaments --- 
'DisplayMenuButtons_ViewTournaments


' --- Displays the search filter for settings - NVP df=yes causes the initial page to hide filters ---
DisplayFilters="none"
IF TRIM(Request("df"))="yes" THEN DisplayFilters="inline"
DisplayRankingsSearchFilters




	' --- If User pressed Find My Rankings button and MemberID was not set OR user pressed get a New Member button ---
'	IF (TRIM(Request("SingleRanking"))<>"" AND TRIM(sMemberID)="") OR TRIM(Request("NewMember"))<>"" THEN
				
			' --- Sends user to search-member routine to selected member
'			Session("SkiYear")="1"
'			Session("sSendingPage")="/rankings/"&ThisFileName&"?SingleRanking=Find&RankingListType="&sRunByWhat
'			Response.Redirect("/rankings/"&SearchFileName&"?rid="&rid&"&formstatus=search")

'	ELSEIF Trim(Request("SingleRanking"))<>"" THEN
'			SingleRanking="Find My Rankings"
'			FindRankingInstances


'	ELSE 		

			' -----  Check Recalculation Underway Flag for the Ski Year selected - if it's currently on Come Back Later otherwise proceed.  -----
			CheckForRankingsRecalcUnderway	

			
			IF RankingsRecalcUnderway= "Y" AND Session("AdminMenuLevel") = 0 and sRunByWhat <> "NCWSA" THEN   ' --- Calc underway - Tell them to try again later
					DisplayRankingsUnderwayMessage


   		ELSE

					' -----  Check for presence of a Selected Division Code and if null then ask user to Select one -- otherwise proceed.  ---
					
					IF RankingsDivSelected = "" THEN   
							DisplayNoDivisionSelected_Message	
						

					ELSE

							' ----------------------------------------
							' --- Begin RANKINGS QUERY and DISPLAY ---
							' ----------------------------------------

							KickTrafficCounter("NewRankPgs")
									
							' --- Runs the query to select tournaments ---
							RunStandingsQueryNew

							' --- Creates listing ---
							DisplayRankingsListing
							
							' --- Gets the Level Percentiles and COA Scores for the selected Division/Event ---
							FindCOAScoreAll

							' --- Displays grid at bottom with Key to Percentiles and ranges ---
							' DisplayPercentilesandPageFooter

		
					END IF		' --- Division selected 

			END IF		' --- Rankings being recalced 

	'END IF	' --- General condition 




' --- Writes the Closing tags for HTML ---
DisplayCloseBodyAndHTMLTags



' ---------------------------------------------------------------------------------------------------------------
' ----------------------   END OF MAIN CODE HERE  ---------------------------------------------------------------
' ---------------------------------------------------------------------------------------------------------------













' ------------------------------------
  SUB CheckForRankingsRecalcUnderway
' ------------------------------------  

			OpenCon
			sSQL = "SELECT Case when RecalcUnderway=1 THEN 'Y' ELSE 'N' END as RCUFlag FROM " & SkiYearTableName & " WHERE SkiYearID = " & SkiYearSelected
			Set rs = Server.CreateObject("ADODB.recordset")
			rs.open sSQL, SConnectionToTRATable, 3, 3  
			
			IF rs.EOF THEN RankingsRecalcUnderway = "N" ELSE RankingsRecalcUnderway = RS("RCUFlag")

			'response.write("<br>" & sSQL)
			'response.write("<br> RankingsRecalcUnderway = " & RankingsRecalcUnderway)
			'response.write("<br>")
			'response.write("<br> FOUND = " & NOT(rs.EOF))
			

END SUB



' -------------------------------------
  SUB DisplayRankingsUnderwayMessage
' -------------------------------------

%>
<div id="displayrankingunderwaymessage" style="display:inline;">
	<a href="javascript:RankingsNavigation('searchfiltersfromrecalcerror');" style="text-decoration:none;">
		<div class="searchimagediv" style="background: url(images/buttons/LongButtonBlank.png); background-repeat: no-repeat;">
			<p class="searchbannerline">CHANGE SETTINGS</p>
		</div>
	</a>

	<div class="errorbox" style="width:90%; padding-top:50px; padding-right:10px; padding-left:10px;">
		<span id="" class="span100">Ranking Recalculations are currently underway For the Ski Year requested. <br>Please try your request again in a few minutes. </span>
	</div>	

</div>
<% 	

END SUB 




' ---------------------------------------
  SUB DisplayNoDivisionSelected_Message
' ---------------------------------------  

' errorbox  width:95%;
%>
<div id="nodivisionselected" style="display:none">
	<div class="errorbox" style="padding-top:50px; min-height:300px; padding-left:0px;">
		<span id="" class="span10" style="border:0px solid; border-color:white; vertical-align:top;">
			<img src="images/buttons/Button-Info-icon.png" style="padding:0px; width:30px; text-align:right; " alt="Tip" />
		</span>
		<span id="" class="span80" style="border:0px solid; border-color:white;">
			To update Rankings search use <br>GO TO SEARCH SETTINGS<br>button above
		</span> 
		<br><br>
		<span id="" class="span30" style="padding-left:12px">
			<img src="images/mobile/TaylorWoolsey_Vert.jpg" style="padding:0px; width:90px; height:112px; text-align:center;" alt="slalom" />
		</span>
		<span id="" class="span30">
			<img src="images/awsa/trick_rtb.jpg" style="padding:0px; width:90px; text-align:center; height:112px" alt="trick" />
		</span>
		<span id="" class="span30">
			<img src="images/awsa/jump_001.jpg" style="padding:0px; width:90px; text-align:center; height:112px;" alt="jump" />								
		</span>
		<br><br>
			<div class="">
				<input type="button" class="buttonblue" name="Continue" value="Continue" style="width:7em; height:2em; font-size:14pt" onclick="javascript:OnOpen('nodivisionselected');">
			</div>
	</div> 		<% ' -- Error box  %>	
</div>			<% ' -- nodivisionselected  %>
<% 	


END SUB





' -------------------------------------------------------
  SUB DisplayNoRankingsFoundForFilter_Message_OBSOLETE 
' -------------------------------------------------------  

' --- NOT USED --
%>
<div id="norankingsfound" style="display:inline-block">
	<a href="javascript:ReturnToListing('norankingsfound'); javascript:get_localStorage();" style="text-decoration:none;">
				<div class="searchimagediv" style="background: url(images/buttons/LongButtonBlank.png); background-repeat: no-repeat;">
			<p class="searchbannerline">GO TO SEARCH SETTINGS</p>
		</div>
	</a>
	<div class="error">
		<span id="" class="span100">No Rankings Found With These Filter Settings.</span>
	</div>	

</div>
<% 	


END SUB




' -------------------------------
	SUB ReadRankingsFormVariables 
' -------------------------------	 	

' --- Determines what functionality was selected ---
sShowSQL=Request("sShowSQL")
sRunByWhat = TRIM(Request("RankingListType"))
IF sRunByWhat = "" THEN sRunByWhat="National"
RankingListType = sRunByWhat


' --- Temporary ---
adminmenulevel = TRIM(Request("adminmenulevel"))
IF Session("adminmenulevel")>=50 THEN
		response.write("Line 342 - RankingListType = "&RankingListType)
 END IF

'response.write("<br>Line 345 BEFORE Request  sMemberID= "&sMemberID)

' --- Define Member information if available ---
sMemberID=TRIM(Request("sMemberID"))
IF Len(sMemberID) > 9 then sMemberID = Left(sMemberID,9)

'response.write("<br>Line 355 AFTER Request sMemberID = "&sMemberID)


' --- What filter? was selected ---
RankingsStateRegionSelected = TRIM(Request("RankingsStateRegionSelected"))
IF RankingsStateRegionSelected = "" THEN RankingsStateRegionSelected = "All"
IF Len(RankingsStateRegionSelected) > 3 then RankingsStateRegionSelected = Left(RankingsStateRegionSelected,3)

' --- Determines what LEAGUE was selected ---
sLeagueSelected = TRIM(Request("sLeagueSelected"))
IF sLeagueSelected = "" THEN sLeagueSelected = "None"

' --- Note - The following request line must be above the definition within the CURRENT member section below. ---
FederationSelected = TRIM(Request("Include_International"))

' -------- If NCWSA then select ALL Federations, otherwise only USA --------
IF FederationSelected = "" AND sRunByWhat="NCWSA" THEN 
		FederationSelected = "ALL"	                
ELSEIF FederationSelected = "" THEN 
		FederationSelected = "USA"
END IF	


' --- Define SkiYear ---
SkiYearSelected = TRIM(Request("RankingsSkiYearIDSelected"))
IF TRIM(SkiYearSelected) = "" AND TRIM(Session("SkiYear"))<>"" THEN SkiYearSelected=Session("SkiYear")



' --- Sets Session("SkiYear") to request string from form  - NCWSA test is done first 
IF (SkiYearSelected = "1" OR SkiYearSelected = "") AND sRunByWhat="NCWSA" THEN
 		OpenCon
		Set rs = Server.CreateObject("ADODB.recordset")
		sSQL = "SELECT top 1 * from " & SkiYearTableName & " WHERE DefaultYear = 1"
 		rs.open sSQL, SConnectionToTRATable, 3, 3  

		IF NOT rs.EOF THEN
				Session("SkiYear")=rs("SkiYearID")
		END IF		

ELSEIF SkiYearSelected <> "" THEN 	' --- Assigns SkiYear to whatever current setting is if there is a variable set on form
		Session("SkiYear") = SkiYearSelected
    
ELSE 																' --- If nothing is assigned, then set it to 12 month rankings
		Session("SkiYear")="1"	

END IF	




' --- Define Event ---
EventSelected = TRIM(Request("RankingsEventSelected"))
IF TRIM(Request("MyEvent"))<>"" THEN EventSelected=TRIM(Request("MyEvent"))
IF EventSelected = "" THEN EventSelected = "S"
IF Len(EventSelected) > 1 then EventSelected = Left(EventSelected,1)

' --- Define Division ---
RankingsDivSelected = TRIM(Request("RankingsDivSelected"))
IF TRIM(Request("MyDiv"))<>"" THEN RankingsDivSelected=TRIM(Request("MyDiv"))
IF Len(RankingsDivSelected) > 2 then RankingsDivSelected = Left(RankingsDivSelected,2)



END SUB



' ----------------------------------
	SUB DetermineNameOfCurrentMember
' ----------------------------------

'response.write("<br>Line 455 sMemberID= "&sMemberID)

' --  Determine Name of CURRENT user ---
SET rs=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT * FROM "&MemberLiveTableName&" AS MT"
'sSQL = sSQL + " WHERE PersonID = '" &RIGHT(sMemberID,8)& "'"
sSQL = sSQL + " WHERE PersonID = '" &RIGHT(sMemberID,8)& "'"
rs.open sSQL, SConnectionToTRATable  


IF NOT rs.EOF THEN
		sFullName=rs("FirstName")&" "&rs("LastName")
		' --- Only reset FED DropDown if the person has just selected a NEW member as determined by MyDiv or MyEvent NOT Null. ---
		IF UCASE(TRIM(rs("FederationCode")))<>"USA" AND (TRIM(Request("MyDiv"))<>"" OR TRIM(Request("MyEvent"))<>"") THEN
				FederationSelected=rs("FederationCode")
		END IF
ELSE
		sFullName="None Selected"
END IF



END SUB



  


' --------------------------------
  SUB DisplayRankingsListing 
' --------------------------------  

' formerly in <p> element position:absolute;
' div formerly class="navbox1" 
%>
<div id="rankingslisting">
	<a href="javascript:DisplaySearchFilters('searchfilters');" style="text-decoration:none;">
				<div class="searchimagediv" style="background: url(images/buttons/LongButtonBlank.png); background-repeat: no-repeat;">
			<p class="searchbannerline">GO TO SEARCH SETTINGS</p>
		</div>
	</a>

	<div class="scroll">
	<%
			IF NOT rs.eof THEN 
					' --- Creates the rankings list ---
					CreateRankingListDisplay
			ELSE
					'response.write("</div><div style=""color:red;"">HERE lin 499</div>")
					DisplayNoListingFound
			END IF	
	%>
	<br><br><br><br><br><br>
	</div> 		<! -- Bottom of scroll box -- ->
</div> 			<! -- Bottom of div for hidding and displaying - TourListing -- ->
<%


END SUB



' --------------------------
  SUB DisplayNoListingFound
' --------------------------  

'response.write("</div><div style=""color:red;"">HERE lin 516</div>")
%>
<div class="tabrankings" style="height:20px; margin-top:30px; padding-top:4px; background-color:<%=DefineLevelcolor%>; text-color:<%=Textcolor1%>; font-size:16px;">
	<span class="span90" style="color:white; text-align:center;"><b>No Rankings Found for these Settings</b></span>
</div>
<%

END SUB




' --------------------------------
   SUB CreateRankingListDisplay
' --------------------------------

		' --- INITIALIZES the Ranking related memory fields for deal with ties.

		' --- RecordNum is essentially the record count
		' --- MemoryScore is the Score of the 
		' --- MemoryRank stores the highest value of placement - for which subsequent records may be tied 
		' --- tRankScore is the Score of the current record

		' ---  After storing the values from the FIRST record then move to the second record to see if tied to know
		' ---     whether the FIRST record should have a T after it.  All others


		RecordNum = 1
		MemoryRank = 1
		IF sRunByWhat = "NSL" THEN MemoryScore = rs("sc_3") ELSE MemoryScore = rs("RankScore")
   


		' --- Displays the header on the top of the table ---
		'DisplayRankHeader


		' -----------------------------------------------------------------------------
		' --- Move to 2nd record to determine if 1st record is tied with the second ---
		' -----------------------------------------------------------------------------
		rs.MoveNEXT
		IF NOT rs.EOF THEN
				' --- Initializes 2nd record in query --- 
				DefineRankingDataLine
		END IF


		' --- If the score from last tied record is same as current score 
		IF MemoryScore = tRankScore THEN
				RankValueWithTies = "1T"
		ELSE
				RankValueWithTies = "1"
		END IF


		' --------------------------------------------------------------------------
		' --- Now move back to FIRST record and initialize First record in query ---
		' --------------------------------------------------------------------------
		rs.MoveFIRST
		DefineRankingDataLine
  


		' -----------------------------------------------------------------------------------------------
		' -----  BEGINNING OF LOOPING FOR DISPLAYING ALL RECORDS MATCHING QUERY  ------------------------
		' -----   Loops thru the remaining 2,3...nth records
		' -----------------------------------------------------------------------------------------------

		DO WHILE NOT rs.eof

				' *****************************************
				' --- Displays one line of ranking list ---
				' *****************************************
				DisplayRankingLine

				
				' --- Initializes NEXT record in query --- 
				rs.moveNEXT
				RecordNum = RecordNum + 1

				IF NOT rs.eof THEN

						' --- Defines the CURRENT record ---
						DefineRankingDataLine
			

						' --- If the score from PREVIOUS record is same as current score ---
						IF cdbl(MemoryScore) = cdbl(tRankScore) THEN
								RankValueWithTies = MemoryRank&"T"
						ELSE
								MemoryRank = RecordNum
								MemoryScore = tRankScore
				
								' --- Move to NEXT record to determine if next record is tied with previous ---
								rs.MoveNEXT
			    			IF NOT rs.eof THEN
					
										' --- Initializes the record beyond the current record in query --- 
										DefineRankingDataLine

										' --- If the score from last tied record is same as current score 
										IF MemoryScore = tRankScore THEN
												RankValueWithTies = RecordNum&"T"
										ELSE
												RankValueWithTies = RecordNum
										END IF

								ELSE
										' --- Can't be tied with EOF so set it to the current record ---
										RankValueWithTies = RecordNum

								END IF


								' --- Now move back to CURRENT record and initialize ---
								IF NOT(rs.bof)  THEN
									
										' --- Move back to the previous record so that MoveNext at top of loop works properly ---
										rs.MovePREVIOUS
									
										' --- Initializes the record beyond the current record in query ---
										DefineRankingDataLine

								END IF
								
						END IF

				ELSE

				END IF

		LOOP  



CloseCon


END SUB





' ----------------------------------------------------------------------------------
    SUB DisplayRankingLine	' --- Displays a single line of the ranking list ---
' ----------------------------------------------------------------------------------

Dim sbgcolor, MembHighlightColor

    ' --- Changes background to red if current member is set and found in this ranking list ---	
    IF rs("MemberID")=sMemberID THEN
				sbgcolor=DefineLevelcolor
				MembHighlightColor=Textcolor3
    ELSE
				sbgcolor=DefineLevelcolor	
				MembHighlightColor=DefineLevelcolor	
    END IF 

		' --- Determines if # or * should be a suffix of score ---
    IF (sRunByWhat <>"NSL" AND instr(tRnkScoBkup,"Rule 1.13")<>0 AND instr(tRnkScoBkup,"Click Skier")=0) THEN
				displaysuffix="#"
    ELSEIF (sRunByWhat <>"NSL" AND instr(tRnkScoBkup,"NO Penalty")=0 AND instr(tRnkScoBkup,"Click Skier")=0) THEN
				displaysuffix="*"
    ELSE
				displaysuffix=""
    END IF 

		' --- Elite status ---
	  IF tEliteStat = "None" THEN
				elitecolor="Red"
		ELSE 
		  	elitecolor=Textcolor1
		END IF   

		IF len(tEliteStat) > 0 THEN	
			  elitetitle=tEliteBkup
			  elitestatus=tEliteStat
		ELSE 
				elitestatus=tEliteStat
		END IF	  


		NationalsPlacement=""
		RegionalsPlacement=""
		RegionalsSkied=""
		IF tRegPlace <> "" THEN 
				RegionalsPlacement=tRegPlace
				IF ucase(tRegSki) <> tRegion THEN
         		RegionalsSkied = ucase(tRegSki)
	  		END IF                  
		END IF 

		IF tNatPlace <> "" THEN
				NationalsPlacement=tNatPlace
	  END IF 


		
		IF TRIM(rs("TourStatus"))<>"X" AND ( TRIM(rs("HomoType"))="A" OR TRIM(tRegion)=RIGHT(LEFT(rs("RQTourID"),3),1) ) THEN  
				RankingQfyTitle="Check details of Qualifications for this Member"
			
				SELECT CASE TRIM(rs("QfyStatus"))
					CASE "QFY-RPR" 
							RankingQfyTitle="QUALIFIED PENDING REGIONAL PARTICIPATION - Check details of Qualifications for this Member" 
					CASE "Qualified" 
							RankingQfyTitle="QUALIFIED - Check details of Qualifications for this Member" 
				END SELECT
				QualifyStatusDisplay=rs("QfyStatus")
						
	 	ELSEIF 	TRIM(rs("TourStatus"))<>"X" AND TRIM(tRegion)<>RIGHT(LEFT(rs("RQTourID"),3),1) THEN 
 				RankingQfyTitle="This Member's HOME REGION is"&tRegion
	    	QualifyStatusDisplay=rs("QfyStatus")="OOR"
		ELSE
				RankingQfyTitle="This Member's HOME REGION is"&tRegion
				QualifyStatusDisplay="---"
		END IF 



		IF tLevelNo > 0 AND tLevelNo <> LastLevelNo AND sRunByWhat <> "NCWSA" THEN 	
				' --- Put a blank row in to separate from heading ---
				%>
				<div class="tabrankings" style="height:20px; padding-top:4px; background-color:<%=DefineLevelcolor%>; text-color:<%=Textcolor1%>; font-size:16px;">
	  			<span class="span45" style="text-align:left;"><b>Level: <%=tLevelNo%></b></span>
	  			<span class="span25" style="text-align:left;"><b>Div: <%=RankingsDivSelected%></b></span>
	  			<span class="span25" style="text-align:left;"><b>Event: <%=EventSelected%></b></span>
				</div>
				<%
		END IF 
		
		IF LastDiv<>RankingsDivSelected AND sRunByWhat = "NCWSA" THEN 	
				' --- Put a blank row in to separate from heading ---
				LastDiv=RankingsDivSelected
				%>
				<div class="tabrankings" style="height:20px; padding-top:4px; background-color:<%=DefineLevelcolor%>; text-color:<%=Textcolor1%>; font-size:16px;">
	  			<span class="span45" style="text-align:left;"><b>Level:</b></span>
	  			<span class="span25" style="text-align:left;"><b>Div: <%=RankingsDivSelected%></b></span>
	  			<span class="span25" style="text-align:left;"><b>Event: <%=EventSelected%></b></span>
				</div>
				<%
		END IF 
		
		

	' --------------------------------------------
	' --- DISPLAY SINGLE MEMBER IN LISTING ---
	' --------------------------------------------

  %>
  <div class="tabrankings" style="height:17px; background-color:<%=sbgcolor%>; font-size:12pt; margin-top:0px; padding-top:0px;" >
		<span class="span15" style="background-color:<%=MembHighlightColor%>"><b><%=RankValueWithTies%></b></span>
  	<span class="span60" style="width:57%;"><b><%=LEFT(tName,20)%></b></span>
	  <span class="span20" style="width:20%; text-align:right;"><b><%=tFmtScore%></b></span>
 		<span style="width=5px; color:red; text-align:left;"><%=displaysuffix%></span>
	</div>
  
  <div class="rankingsbody" style="background-color:#FFFFFF; font-size:10pt; font-weight:normal; margin-top:0px; padding-top:2px;">
  	<%
  	IF sRunByWhat = "NCWSA" THEN 
  			%>
				<span class="span30">Team:&nbsp;<% =tTeam %> (<% =tTeamStat %>)</span>
				<span class="span30">Reg:&nbsp;<% =tNCWRegn %></span>
				<span class="span30">Conf:&nbsp;<% =tNCWConf %></span>
  			<% 
  	ELSE 
  			%>
				<span class="span25">Ntls Plc:&nbsp;<% =NationalsPlacement %></span>
				<span class="span25">&nbsp;Rgl Plc:&nbsp;<% =RegionalsPlacement %></span>
				<span class="span25" style="width:22%;">&nbsp;</span>
				<span class="span20" style="width:16% text-align:right;">Pctl:<% =tRankPct%></span>
	 			<% 
  	END IF
		%>
	</div>
  
  <div class="rankingsbottom"  style="background-color:#FFFFFF; margin-left:2px;">
   	<span class="span10" style="font-size:10pt;"><b><% =tState %></b></span>
		<span class="span15"><% =tMemberFed %></span>
		<span class="span25" style="text-align:left;">&nbsp;Region:&nbsp;<% =tRegion %></span>
		<%
    
    IF sRunByWhat <> "NSL" AND sRunByWhat <> "NCWSA" THEN
				%>
				<span class="span45" id="qualificationstatus">Qfy Stat: <%=QualifyStatusDisplay%></span>
				<%
    END IF 
		
		%>
	</div>
	<%


' --- Saves last color for drawing bar across screen at level break, Only if Not Zero ---
IF tLevelNo > 0 THEN LastLevelNo = tLevelNo


END SUB




' ----------------------
  SUB DisplayRankHeader
' ----------------------

%>
<div class="rankingsheader" id="rankingsheader">
<%


Headcolor1="blue"


' --- Top of large condition of branching to most of rest of code   ---


  %>
   <div class="headerrankings" style="background-color:<%=Headcolor1%>;" align="Center" valign="top">
		<span class="span10">Rank</span>
  	<span class="span50">Member</span>
	  <span class="span15">Ranking</span>
 		<span style="width=10px; color="red"">&nbsp;</span>
		<span class="span15"><b>Pctl</b></span>
	</div>
	
	<div class="headerrankings" style="background-color:<%=Headcolor1%>;" align="Center" valign="top">
  	<%	
  	IF sRunByWhat = "NCWSA" THEN 	
  			%>
  			<span class="span20">Team</span>
	  		<span class="span20">Region</span>
 				<span class="span20">Conf</span>
 				<%
		ELSE
				%>
				<span class="span20"><b>Reg-Plc</b></span>
				<span class="span20"><b>Nat-Plc</b></span>
				<span class="span20"><b>Home-Rg</b></span>
				<%
		END IF


		%>	
	</div>
   
  <div class="headerrankings" style="background-color:<%=Headcolor1%>;" align="Center" valign="top">
		<%
    IF sRunByWhat <> "NSL" AND sRunByWhat <> "NCWSA" THEN
				%>
   			<span class="span10">State</span>
				<span class="span25">Fed</span>
				<%
    END IF 
		
		%>
		<span class="span25"><%=QualifyStatusDisplay%></span>

	</div>   	

		
</div>	
<%



END SUB








' -----------------------------
   SUB DefineRankingDataLine
' -----------------------------

' Extract key items from the Query Answerset for this Member's Ranking

tName = TRIM(rs("LastName")) & ", " & rs("FirstName")
tMemberID = rs("MemberID")
tState = rs("State")
tRegion = rs("Region")
tMemberFed = rs("FederationCode")
tDefaultYear = rs("DefaultYear")
tTeam = rs("Team")
tTeamStat = rs("TeamStat")
tNCWRegn = rs("NCWRegion")
tNCWConf = rs("NCWConf")

If len(trim(rs("OpenDate"))) > 0 THEN
   if len(trim(rs("MastDate"))) > 0 THEN
      tEliteStat = tOpenDiv & "/" & tMastDiv
      tEliteBkup = tOpenDiv & " Qual fm " & rs("OpenDivOrig") & " thru " & rs("OpenDate") & Chr(13) & Chr(10) & tMastDiv & " Qual fm " & rs("MastDivOrig") & " thru " & rs("MastDate")
   ELSE
      tEliteStat = tOpenDiv
      tEliteBkup = tOpenDiv & " Qual fm " & rs("OpenDivOrig") & " thru " & rs("OpenDate")
   END IF
ELSE   
   if len(trim(rs("MastDate"))) > 0 THEN
      tEliteStat = tMastDiv
      tEliteBkup = tMastDiv & " Qual fm " & rs("MastDivOrig") & " thru " & rs("MastDate")
   ELSEIF tDefaultYear = 0 and (RankingsDivSelected = "OM" or RankingsDivSelected = "OW" or RankingsDivSelected = "MM") THEN
      tEliteStat = "None"
      tEliteBkup = "Not Elite Qualified in this Event"
   ELSE
      tEliteStat = ""
      tEliteBkup = ""
   END IF
END IF

IF sRunByWhat = "NSL" THEN tRankScore = rs("sc_3") ELSE tRankScore = rs("RankScore")

tRnkScoBkup = rs("RnkScoBkup")
tRating = rs("AWSA_Rat")

IF rs("RankPct")>0 THEN
		tRankPct=FormatNumber(rs("RankPct"),2) 
	ELSE
		tRankPct=""
	END IF		

IF left(tRating,1)= EventSelected THEN
		tLevelNo=RIGHT(tRating,1)
	ELSE
		tLevelNo=0
	END IF		

tNatPlace = TRIM(rs("natl_plc"))
tRegPlace = TRIM(rs("regl_plc"))
tRegSki = rs("reg_ski")

SELECT CASE tRegion
	CASE 1
		tRegion = "C"
	CASE 2
		tRegion = "M"
	CASE 3
		tRegion = "W"
	CASE 4
		tRegion = "S"
	CASE 5
		tRegion = "E"
END SELECT


IF EventSelected = "S" THEN
	tFmtScore = FormatNumber(tRankScore,2)
ELSE
	IF EventSelected = "O" or EventSelected = "J" THEN
		tFmtScore = FormatNumber(tRankScore,1)
	ELSE
		tFmtScore = FormatNumber(tRankScore,0)
	END IF
END IF
       



' --- Establishes background color for the current record
' --- Also determines the current Level Number, as a function
' --- of the "Rank_Level" value and the DT Percents

IF RankingListType="NCWSA" THEN 
		DefineLevelcolor=scolor06
	ELSE 
		SELECT CASE tLevelNo
			CASE 0
				DefineLevelcolor=scolor01
			CASE 1
				DefineLevelcolor=scolor01
			CASE 2
				DefineLevelcolor=scolor02
			CASE 3
				DefineLevelcolor=scolor03
			CASE 4
				DefineLevelcolor=scolor04
			CASE 5
				DefineLevelcolor=scolor05
			CASE 6
				DefineLevelcolor=scolor06
			CASE 7
				DefineLevelcolor=scolor07
			CASE 8
				DefineLevelcolor=scolor08
			CASE 9
				DefineLevelcolor=scolor09
		END SELECT
	END IF

END SUB




' ------------------------
  SUB CheckValidDivision
' ------------------------

' ----------------------------------------------------------------------------------------------------------------
' -------- If division is not in query list for sRunByWhat then set DivisionSelected to the FIRST division
' --------   found in Rankings Table meeting the filtering parameters for the current sRunByWhat 
' ----------------------------------------------------------------------------------------------------------------

SET rsSelectFields=Server.CreateObject("ADODB.recordset")
RunDivQuery
sSQL = sSQL + " AND DIV='"&RankingsDivSelected&"'"
rsSelectFields.open sSQL, SConnectionToTRATable

' --- Not found so reset DivisionSelected to first one found in Rankings table---
IF rsSelectFields.eof THEN 
	SET rsSelectFields=Server.CreateObject("ADODB.recordset")
	RunDivQuery
	sSQL = sSQL + " order by div"
	rsSelectFields.open sSQL, SConnectionToTRATable

	IF NOT rsSelectFields.eof THEN RankingsDivSelected=rsSelectFields("Div")
END IF

rsSelectFields.close


END SUB


' -----------------
  SUB RunDivQuery
' -----------------

sSQL = "Select top 1 div from " & RankTableName

SELECT CASE sRunByWhat
	CASE "Junior"
			sSQL = sSQL + " WHERE LOWER(left(div,1)) in ('b','g')"
	CASE "NSL"
			sSQL = sSQL + " WHERE LOWER(left(div,1)) in ('x','y')"
	CASE "NCWSA"
			sSQL = sSQL + " WHERE LOWER(left(div,1)) = 'c'"
	CASE ELSE
			IF Session("AdminMenuLevel")>0 THEN
					sSQL = sSQL + " WHERE (lower(left(RT.div,1)) in ('b','g','m','w','o','e') or lower(RT.Div) = 'sm')"
			ELSE
					sSQL = sSQL + " WHERE lower(left(RT.div,1)) in ('b','g','m','w','o')"
			END IF
END SELECT

END SUB




'--------------------------- 
   SUB RunStandingsQueryNew
'--------------------------- 

'	Creates and execute SQL query against Rankings Table for Selected Division/Event
'	Newly restructured Feb 2008, one single universal query for all ranking types


    IF instr("MW",Right(RankingsDivSelected,1)) > 0 THEN
       tOpenDiv = "O" & Right(RankingsDivSelected,1)
    ELSEIF instr("BM",left(RankingsDivSelected,1)) > 0 THEN
       tOpenDiv = "OM"
    ELSE
       tOpenDiv = "OW"
    END IF
    tMastDiv = "M" & right(tOpenDiv,1)

    sSQL = "Select distinct MT.lastname, MT.firstname, MT.federationcode, MT.state, MT.Region," 

	' Explore removing reference to DT entry -- Up_Age appears unused. -- Ditto for SY table mebbe ????

    sSQL = sSQL & " RT.*, SY.DefaultYear, TT.NCWRegion, TT.NCWConf,"
    sSQL = sSQL & " Coalesce(Convert(char(10),OQ.QualThru,111),'') as OpenDate,"
    sSQL = sSQL & " Coalesce(Convert(char(10),MQ.QualThru,111),'') as MastDate,"
		sSQL = sSQL & " OQ.DivOrig as OpenDivOrig, MQ.DivOrig as MastDivOrig,"
'    sSQL = sSQL & " RQ.QfyStatus, RQ.QfyStatusTEMP, RQ.TourID AS RQTourID, Coalesce(LT.HomoType,'-') AS HomoType, LT.Status AS TourStatus"
    sSQL = sSQL & " RQ.QfyStatus, RQ.TourID AS RQTourID, Coalesce(LT.HomoType,'-') AS HomoType, LT.Status AS TourStatus"

    sSQL = sSQL & " FROM "&RankTableName&" as RT"
    ' sSQL = sSQL & " JOIN "&MemberTableName&" as MT on RT.memberid = MT.personidwithcheckdigit "
		
		sSQL = sSQL + "	JOIN "&MemberShortTableName&" AS MT ON CAST(RIGHT(RT.MemberID,8) AS INT)=MT.PersonID"
		

    sSQL = sSQL & " JOIN " & SkiYearTableName & " as SY on SY.SkiYearID = RT.SkiYearID"

    sSQL = sSQL & " LEFT JOIN " & TeamTableName & " as TT on TT.TeamID = RT.Team"

    sSQL = sSQL & " LEFT JOIN " & EliteDateTableName & " as OQ on OQ.MemberID = RT.MemberID"
    sSQL = sSQL & " and OQ.Event = RT.Event and OQ.DivElite = '" & tOpenDiv & "'"
    sSQL = sSQL & " and OQ.SkiYearID = " & Session("SkiYear")
    
    sSQL = sSQL & " LEFT JOIN " & EliteDateTableName & " as MQ on MQ.MemberID = RT.MemberID"
    sSQL = sSQL & " and MQ.Event = RT.Event and MQ.DivElite = '" & tMastDiv & "'"
    sSQL = sSQL & " and MQ.SkiYearID = " & Session("SkiYear")
        
'    sSQL = sSQL & " LEFT JOIN "&RegGenTableName&" AS RGEN ON RGEN.MemberID=RT.MemberID AND LEFT(RGEN.TourID,6)='07W999'" 	
'    sSQL = sSQL & " LEFT JOIN "&RegionTableName&" as RG ON lower(MT.state) = lower(RG.state) "

    sSQL = sSQL & " LEFT JOIN "&LeagueTableName&" AS LT ON LT.LeagueID='"&sLeagueSelected&"'"
    sSQL = sSQL & " LEFT JOIN "&RegQualifyTableName&" AS RQ ON LEFT(RQ.TourID,6)=LEFT(LT.QualifyTour,6) AND RQ.MemberID=RT.MemberID AND RQ.Event=RT.Event AND RQ.Div=RT.Div"
	
    sSQL = sSQL & " WHERE RT.div = '" & RankingsDivSelected & "'"
'    sSQL = sSQL & " AND RT.DivType <> 'D'"
    sSQL = sSQL & " AND RT.[event] = '" & EventSelected & "'"
    sSQL = sSQL & " AND RT.SkiYearID = " & Session("SkiYear")

		IF Left(RankingsStateRegionSelected,1) = "1" THEN
				sSQL = sSQL & " AND TT.NCWRegion = '" & Mid(RankingsStateRegionSelected,2) & "'"
		ELSEIF Left(RankingsStateRegionSelected,1) = "2" THEN
				sSQL = sSQL & " AND TT.NCWConf = '" & Mid(RankingsStateRegionSelected,2) & "'"
		ELSEIF Left(RankingsStateRegionSelected,1) = "3" THEN
				sSQL = sSQL & " AND MT.Region = '" & Mid(RankingsStateRegionSelected,2) & "'"
		ELSEIF Left(RankingsStateRegionSelected,1) = "4" THEN
				sSQL = sSQL & " AND MT.State = '" & Mid(RankingsStateRegionSelected,2) & "'"
		END IF

    IF FederationSelected = "USA" THEN sSQL = sSQL & " AND MT.federationcode = 'USA'"

    ' --- NSL (still needed ?) --- 
    IF sRunByWhat = "NSL" THEN 
    		sSQL = sSQL & " AND RT.sc_3 is not NULL ORDER by RT.sc_3 DESC" 
    ELSE
    		sSQL = sSQL & " AND RT.RankScore is not NULL ORDER by RT.RankScore DESC" 
    END IF


		IF Session("Adminmenulevel")=50 AND sShowSQL="on" THEN 
				response.write(sSQL)
				'response.end
		END IF


		' -------------------------
		' --- Execute SQL query ---
		' -------------------------
		OpenCon
    SET rs=Server.CreateObject("ADODB.recordset")
		rs.CursorType = 3
		rs.open sSQL, SConnectionToTRATable
		' adOpenDynamic  



rowCount = 0
'Response.Write("<BR>")



END SUB




' -------------------------------------------
  SUB ChoosePagesSQL(sSQL, sStart, sSize)
' -------------------------------------------
	SET rs=Server.CreateObject("ADODB.recordset")

	'WriteDebugSQL(sSQL)

	sqlstmt = sSQL
	rs.CursorType = 3
'  rs.PageSize = cint(sSize)
	rs.open sqlstmt, SConnectionToTRATable
'  IF isrecordsetempty = false THEN
'    rs.AbsolutePage = cINT(sStart)
'  END IF
END SUB















' --------------------
  SUB FindCOAScoreAll
' --------------------

' --------------------------------------------------------------------------------
' --- Look up the Level Percents and COA Scores for the selected Division / Event
' --------------------------------------------------------------------------------


IF sRunByWhat <>"NCWSA" AND sRunByWhat<>"NSL" THEN

' Here is the new logic for this section -- DC 20080131

'	Get Values from new CutOffTable

	IF EventSelected="S" THEN i=2 ELSE IF EventSelected="J" THEN i=1 ELSE i=0

	Set rs = Server.CreateObject("ADODB.recordset")

	sSQL = " SELECT TOP 1 * FROM " & CutOffTableName  
	sSQL = sSQL + " WHERE Div = '" & RankingsDivSelected & "' AND Event = '" & EventSelected & "' AND SkiYearID = " & Session("SkiYear")

	'WriteDebugSQL(sSQL)

	rs.open sSQL, SConnectionToTRATable

	IF NOT rs.EOF THEN 
		tPerc3 = rs("Pct3"): COALevel3 = FormatNumber(rs("COA3"),i)
		tPerc4 = rs("Pct4"): COALevel4 = FormatNumber(rs("COA4"),i)
		tPerc5 = rs("Pct5"): COALevel5 = FormatNumber(rs("COA5"),i)
		tPerc6 = rs("Pct6"): COALevel6 = FormatNumber(rs("COA6"),i)
		tPerc7 = rs("Pct7"): COALevel7 = FormatNumber(rs("COA7"),i)
		tPerc8 = rs("Pct8"): COALevel8 = FormatNumber(rs("COA8"),i)
		tPerc9 = rs("Pct9"): COALevel9 = FormatNumber(rs("COA9"),i)
	ELSE
		tPerc3 = 0: COALevel3 = "0"
		tPerc4 = 0: COALevel4 = "0"
		tPerc5 = 0: COALevel5 = "0"
		tPerc6 = 0: COALevel6 = "0"
		tPerc7 = 0: COALevel7 = "0"
		tPerc8 = 0: COALevel8 = "0"
		tPerc9 = 0: COALevel9 = "0"
	END IF

	rs.close

END IF

END SUB







' --------------------------------------------------------------------------------------------
   SUB DisplayRankingsSearchFilters      ' -----   Begin form for selection /  filtering parameters ------
' --------------------------------------------------------------------------------------------

Titlecolor=Textcolor2


%>
<div id="rankingssearchsettings" style="display:<%=DisplayFilters%>; margin-top:10px;">
	<%		
		
		' --- Displays the filter dropdowns inside ---
		CreateRankingsFilters
		
	%>
</div> <! -- Bottom of div for hidding and displaying -- ->
<%


END SUB



' --------------------------------------------------------------------------------------------
   SUB DisplayRankingsSearchFilters_OBSOLETE      ' -----   Begin form for selection /  filtering parameters ------
' --------------------------------------------------------------------------------------------

Titlecolor=Textcolor2


%>
<div id="rankingssearchsettings" style="display:<%=DisplayFilters%>; margin-top:10px;">

	<div class="scroll">
		<%   
		
		' --- Displays the filter dropdowns inside ---
		CreateRankingsFilters
		
		%>
	</div> <! -- Bottom of scroll box -- ->
	
	
</div> <! -- Bottom of div for hidding and displaying -- ->
<%


END SUB




' -------------------------
  SUB CreateRankingsFilters
' -------------------------  


%>
<div id="createrankingsfilters" class="errorbox" style="width:99%; margin:3px 0px 0px 0px; padding:0px 0px 0px 0px; height:465px">
<form method=post action="<%=ThisFileName%>">
  <input type="hidden" name="adminmenulevel" value="<%=adminmenulevel%>">
  <input type="hidden" id="sMemberID_Hidden_InRankingsSettings" name="sMemberID" value="">
 	<div style="width:96%; margin-top:10px; padding-left:10px; text-align:left; border:0px solid red;">		
			<span class="span95" style="margin-left:0px; padding-left:0px; text-align:center; font-size:16px; color:yellow; border:0px solid white;">Set Rankings Filters For Search</span> 
	</div>

	<div class="rankingfilterdropdownline" style="margin-top:8px">
		<span class="span20" style="text-align:right;">Div</span>
		<span class="span75" style="text-align:left; border:0px solid white;">
			<%

			' ---  Build DIVISION dropdown list --- 
			BuildDivisionDropDown

			%>
  	</span>
	</div>

	<div class="rankingfilterdropdownline">
		<span class="span20" style="text-align:right;">Event</span>
		<span class="span75" style="text-align:left;">
		<%

		BuildEventDropDown

		%>
  	</span>
	</div>

	<div class="rankingfilterdropdownline">
		<span class="span20" style="text-align:right;">
			<% 
			IF sRunByWhat = "NCWSA" THEN 
					%><b>Reg/Conf</b><%
			ELSE 
					%><b>Reg/ST</b><% 
			END IF 
			%>    
		</span>
		<span class="span75" style="text-align:left;">
		<%

		' ----  Build RankingsStateRegionSelected  REGION or STATE (AWSA)  or  REGION OR CONFERENCE (NCWSA)  dropdown list  ---- 
		BuildRankingsStateRegionSelectedDropDown

		%>
  	</span>
	</div>



	<% ' --------------------------------- Build SKI YEAR dropdown list  ------------------- %>
	<div class="rankingfilterdropdownline">
		<span class="span20" style="text-align:right;">Ski Year</span>
		<span class="span75" style="text-align:left;">
		<%

		BuildSkiYearDropDown

		%>	
  	</span>
	</div>



	<div class="rankingfilterdropdownline">
		<span class="span20" style="text-align:right;">Ranking</span>
		<span class="span75" style="text-align:left;">
			<select id='RankingListType' name='RankingListType' style="width:12em; font-size:12pt">
	  		<option value ='National' <% IF RankingListType="National" THEN response.write(" selected")%> >National</Option><br>
		  	<option value ='Junior' <% IF RankingListType="Junior" THEN response.write(" selected")%> >Junior</Option><br>
		  	<option value ='NCWSA' <% IF RankingListType="NCWSA" THEN response.write(" selected")%> >NCWSA</Option><br>
			</select>
  	</span>
	</div>
	<%
	
	IF sRunByWhat <> "NSL" AND sRunByWhat <> "NCWSA" AND TRIM(session("SkiYear"))=1 THEN 
			%>
			<div class="rankingfilterdropdownline">
				<span class="span20" style="text-align:right;">League</span>
				<span class="span75" style="text-align:left;">
  					<%
						' --- Procedure found in Tools_Leagues.asp ---
						BuildLeagueDrop_Mobile true, "None" 

						%>
  			</span>
			</div>
			<%
	END IF  
				
	%>
	<div class="rankingfilterdropdownline" style="margin-left:0px; padding-bottom:0px;"">
		<span class="span20" style="text-align:right;">Fed</span>
	  <span class="span75" style="text-align:left;">
			<%
			
			' --- Build the FEDERATION drop down ---
			BuildFederationDropDown
			
			%>
  	</span>
	</div>
	<%
	
	DispMemberName="N"
	IF DispMemberName="Y" THEN
			%>	
	<div class="rankingfilterdropdownline" style="margin:0px; padding:0px; border:0px solid white; height:25px;">
		<span class="span20" style="text-align:right; padding-bottom:0px;">Name:</span>
  	<span class="span75" style="text-align:left; padding-bottom:0px;">
			&nbsp;<input type="text" class="textbox_hidden_banner" style="padding:0px; margin:0px width:95%;" name="sName_InRankingsSettings" id="sName_InRankingsSettings" value="" MaxLength="15">
  	</span> 	
	</div>
	<%
	END IF
	
	DispMemberID="N"
	IF DispMemberID="Y" THEN
			%>	
			<div class="rankingfilterdropdownline">
				<span class="span20" style="text-align:right;">ID:</span>
  			<span class="span75" style="text-align:left;">
  				&nbsp;<input type="text" class="textbox_hidden_banner" name="sMemberID_InRankingsSettings" id="sMemberID_InRankingsSettings" value="" MaxLength=15>
  			</span>  	
			</div>
			<%
	END IF
	
	%>
	<div style="border:0px solid white; margin-left:0px; padding-left;0px; margin-top:30px; border:0px solid white;">
		<span class="span45" style="margin-left:0px; padding-left:0px; text-align:left;">
			<input type=button class="buttonblue" style="width:8.5em;" value="Save Settings" onclick="javascript:StoreRankingsSettingsToLocalVar();" title="Look up my Rankings in all Events">
		</span>
		<span class="span45" style="text-align:center;">
			<input type=button class="buttonblue" style="width:8.5em;" value="Recall Settings" onclick="javascript:UpdateRankingsSettingsFromLocal();" title="Look up my Rankings in all Events">
		</span>
	</div>
	<div style="border:0px solid white; margin-top:40px; margin-left:0px; padding-left;5px;">
		<span class="span95" style="text-align:center;">
  		<input type=submit class="buttonblue" name="DisplayRankings" style="width:180px; font-size:12pt;" value="Display Rankings" title="Display Rankings for these parameters" onclick="javascript:DisplaySearchFilters('displayrankings')">
  	</span>
	</div>
</form>
</div>
<%


END SUB




' ---------------------------
  SUB BuildSkiYearDropDown
' ---------------------------

	%>
	<SELECT id='RankingsSkiYearIDSelected' name='RankingsSkiYearIDSelected' style="width:12em; font-size:12pt;"><%

		
		sSQL = "SELECT DISTINCT RT.SkiYearID, SY.SkiYearName"
		sSQL = sSQL + " FROM " &RankTableName&" AS RT"
		sSQL = sSQL + " JOIN " &SkiYearTableName&" AS SY ON RT.SkiYearID = SY.SkiYearID"

		' --- NCWSA does not display 12 Month Rankings
		IF sRunByWhat="NCWSA" THEN
				sSQL = sSQL + " WHERE SY.SkiYearID <> 1"
		END IF

		SET rsSelectFields=Server.CreateObject("ADODB.recordset")
		rsSelectFields.open sSQL, SConnectionToTRATable

		' -- Loads dropdown and sets default to Session("SkiYear")
		DO WHILE NOT rsSelectFields.eof

			IF TRIM(rsSelectFields("SkiYearID")) = TRIM(session("SkiYear")) THEN
					response.write("<option value =""" & rsSelectFields("SkiYearID") &""" selected>")
					response.write(rsSelectFields("SkiYearName"))
					response.write("</option><br>")
			ELSE
			response.write("<option value =""" & rsSelectFields("SkiYearID") &""">")
			response.write(rsSelectFields("SkiYearName"))
			response.write("</option><br>")
		END IF 

		rsSelectFields.moveNEXT

	LOOP

	rsSelectFields.close 
	%>
	</select>
	<%
  

END SUB




' ------------------------
  SUB BuildEventDropDown
' ------------------------
  
	%>
	<select id='RankingsEventSelected' name='RankingsEventSelected' style="width:12em; font-size:12pt">
	  <option value ='S' <%IF EventSelected="S" THEN response.write(" selected")%>>Slalom</Option><br>
	  <option value ='J' <%IF EventSelected="J" THEN response.write(" selected")%>>Jump</Option><br>
	  <option value ='T' <%IF EventSelected="T" THEN response.write(" selected")%>>Trick</Option><br>

	  <% IF sRunByWhat <> "NSL" AND sRunByWhat <> "NCWSA" THEN %>
	  	<option value ='O' <%IF EventSelected = "O" THEN Response.Write(" selected")%>>Overall</option><br>
	  <% END IF %>

	</select>
	<%

END SUB  




' ---------------------------
  SUB BuildDivisionDropDown
' ---------------------------

	%>
	<select id='RankingsDivSelected' name='RankingsDivSelected' style="width:12em; font-size:12pt"><%

	SET rsSelectFields=Server.CreateObject("ADODB.recordset")
	sSQL = "Select distinct RT.div, DT.div_name from "&RankTableName&" as RT JOIN "&DivisionsTableName&" as DT ON RT.div = DT.div"

	SELECT CASE sRunByWhat
  		CASE "National"
				IF Session("AdminMenuLevel")>0 THEN
					sSQL = sSQL + " WHERE (lower(left(RT.div,1)) in ('b','g','m','w','o','e') or lower(RT.Div) = 'sm')"
				ELSE
					sSQL = sSQL + " WHERE lower(left(RT.div,1)) in ('b','g','m','w','o')"
				END IF
			CASE "NSL"
    			sSQL = sSQL + " WHERE lower(left(RT.div,1)) in ('y','x')"
	  	CASE "Junior"
    			sSQL = sSQL + " WHERE lower(left(RT.div,1)) in ('b','g')"
	  	CASE "NCWSA"
					sSQL = sSQL + " WHERE lower(left(RT.div,1)) = 'c'"
	END SELECT

	sSQL = sSQL + " order by RT.div"
	rsSelectFields.open sSQL, SConnectionToTRATable


	' ---  This section deals with case WHERE no scores exist for any of the divisions  ---
	IF NOT rsSelectFields.eof THEN 
	  	rsSelectFields.movefirst
  		DO WHILE NOT rsSelectFields.eof
	    		IF TRIM(rsSelectFields("Div")) = RankingsDivSelected THEN
      				response.write("<option value ="""&rsSelectFields("Div")&""" selected>"&rsSelectFields("Div")&" - "&rsSelectFields("Div_Name")&"</option><br>")
    			ELSE
      				response.write("<option value ="""&rsSelectFields("Div")&""">"&rsSelectFields("Div")&" - "&rsSelectFields("Div_Name")&"</option><br>")
	    		END IF	
			rsSelectFields.moveNEXT
  		LOOP
	ELSE
  		response.write("<option value =""None"" selected>None</option>")
	END IF

	rsSelectFields.close %>
	</select>
	<%


END SUB





' ---------------------------------
  SUB BuildRankingsStateRegionSelectedDropDown
' ---------------------------------

	%>
	<select id='RankingsStateRegionSelected' name='RankingsStateRegionSelected' style="width:12em; font-size:12pt">

	<% IF sRunByWhat = "NCWSA" THEN %>

	  <option value ='All'  <%IF RankingsStateRegionSelected="All"  THEN response.write(" selected")%>>All</Option><br>
	  <option value ='1E'  <%IF RankingsStateRegionSelected="1E"  THEN response.write(" selected")%>>Eastern Region</Option><br>
	  <option value ='2NE' <%IF RankingsStateRegionSelected="2NE" THEN response.write(" selected")%>>.. Northeast Conf</Option><br>
	  <option value ='2SA' <%IF RankingsStateRegionSelected="2SA" THEN response.write(" selected")%>>.. So Atlantic Conf</Option><br>
	  <option value ='2SO' <%IF RankingsStateRegionSelected="2SO" THEN response.write(" selected")%>>.. Southern Conf</Option><br>
	  <option value ='1M'  <%IF RankingsStateRegionSelected="1M"  THEN response.write(" selected")%>>Midwest Region</Option><br>
	  <option value ='2GL' <%IF RankingsStateRegionSelected="2GL" THEN response.write(" selected")%>>.. Great Lakes Conf</Option><br>
	  <option value ='2GP' <%IF RankingsStateRegionSelected="2GP" THEN response.write(" selected")%>>.. Great Plains Conf</Option><br>
	  <option value ='1SC' <%IF RankingsStateRegionSelected="1SC" THEN Response.Write(" selected")%>>South Central Region</option><br>
	  <option value ='1W'  <%IF RankingsStateRegionSelected="1W"  THEN response.write(" selected")%>>Western Region</Option><br>
	  <option value ='2NW' <%IF RankingsStateRegionSelected="2NW" THEN response.write(" selected")%>>.. Northwest Conf</Option><br>
	  <option value ='2PC' <%IF RankingsStateRegionSelected="2PC" THEN response.write(" selected")%>>.. Pacific Conf</Option><br>

	<% ELSE %>

	  <option value ='All'  <%IF RankingsStateRegionSelected="All"  THEN response.write(" selected")%>>All</Option><br>

		<% sSQL = "SELECT CASE When Region = '5' then 'Eastern' When Region = '2' then 'Midwest'"
		sSQL = sSQL + " When Region = '1' then 'South Central' When Region = '4' then 'Southern'"
		sSQL = sSQL + " When Region = '3' then 'Western' else 'Unknown' end as RegionName,"
		sSQL = sSQL + " Region, State, StateName FROM " & RegionTableName 
		sSQL = sSQL + " Order by Case When Region = '5' then 'E' When Region = '2' then 'M'"
		sSQL = sSQL + " When Region = '1' then 'P' when Region = '4' then 'S'"
		sSQL = sSQL + " When Region = '3' then 'W' else 'Z' end, StateName;"

		
		SET rsSelectFields=Server.CreateObject("ADODB.recordset")
		rsSelectFields.open sSQL, SConnectionToTRATable

		LastRegion = "0"
		DO WHILE NOT rsSelectFields.eof

			IF Trim(rsSelectFields("Region")) <> LastRegion THEN
				LastRegion = Trim(rsSelectFields("Region")) %>

				<option value ='3<%=LastRegion%>'<%IF RankingsStateRegionSelected = "3"&LastRegion THEN Response.Write(" selected ")%>><%=rsSelectFields("RegionName")&" Region"%></Option><br>

			<% END IF %>

			<option value ='4<%=Trim(rsSelectFields("State"))%>'<%IF RankingsStateRegionSelected = "4"&Trim(rsSelectFields("State")) THEN Response.Write(" selected ")%>>... <%=rsSelectFields("StateName")%></Option><br>

			<% rsSelectFields.moveNEXT

		LOOP

		rsSelectFields.close

	END IF %>    

	</select>
	<%


END SUB



' -------------------------------
  SUB BuildFederationDropDown
' -------------------------------

	' --------------------------------  Build FEDERATION dropdown list  -------------------- 
	%>
	<select id="Include_International" name="Include_International" style="width:12em; font-size:12pt">
	<option value="ALL"<%IF FederationSelected = "ALL" THEN Response.Write(" selected")%>>All Federations</option>
	<option value="USA"<%IF FederationSelected = "USA" THEN Response.Write(" selected")%>>USA Only</option>
	</select>
	<%


END SUB  



' --------------------------
  SUB FindRankingInstances
' -------------------------- 

'response.write("<br>MADE IT - Line 1340 Top of FindRankingInstances")
'response.end


sSkiYearID=TRIM(Session("SkiYear"))

sSQL = "SELECT * FROM "&RankTableName&" AS RT"
'sSQL = sSQL + " JOIN "&MemberTableName&" AS MT ON MT.PersonIDwithCheckDigit=RT.MemberID"
sSQL = sSQL + "	JOIN "&MemberShortTableName&" AS MT ON CAST(RIGHT(RT.MemberID,8) AS INT)=MT.PersonID"
sSQL = sSQL + " WHERE MemberID='"&sMemberID&"' AND SkiYearID='"&sSkiYearID&"'"

SELECT CASE sRunByWhat
	CASE "National"
			IF Session("AdminMenuLevel")>50 THEN
				sSQL = sSQL + " AND (lower(left(RT.div,1)) in ('b','g','m','w','o','e') or lower(RT.Div) = 'sm')"
			ELSE
				sSQL = sSQL + " AND lower(left(RT.div,1)) in ('b','g','m','w','o')"
			END IF
	CASE "NSL"
    	sSQL = sSQL + " AND lower(left(RT.div,1)) in ('y','x')"
	CASE "Junior"
  			sSQL = sSQL + " AND lower(left(RT.div,1)) in ('b','g')"
	CASE "NCWSA"
			sSQL = sSQL + " AND lower(left(RT.div,1)) = 'c'"
END SELECT

sSQL = sSQL + " ORDER BY RT.Div, CASE when RT.Event='S' then 1 when RT.event='T' then 2 when RT.Event='J' then 3 else 4 end"


'response.write("<br>Line 1369 After SQL build")
'response.write("<br><br>"& sSQL)
'response.end

' --- Error trap for missing Tour and Member ---
Dim QuerySuccess
QuerySuccess=false
SET rsRankList=Server.CreateObject("ADODB.recordset")
IF TRIM(sMemberID)<>"" AND TRIM(sSkiYearID)<>"" THEN
		rsRankList.open sSQL, SConnectionToTRATable
		IF NOT(rsRankList.eof) THEN QuerySuccess=true
ELSE
		ErrorSubject="ERROR -  Line 1374 - View-StandingsHQ - TRIM(sMemberID) IS NULL OR TRIM(sSkiYearID) IS NULL"
		ErrorSQL = sSQL
		SendErrorEmailToMark ErrorSubject, ErrorSQL
END IF

'response.write("<br><br>Line 1392 After Query - QuerySuccess="&QuerySuccess)
'response.end




IF QuerySuccess=true THEN 

		'response.write("<br><br>Line 1393 Inside QuerySuccess=true IF <br>")
'		response.end
	
		%>
		<TABLE class="innertable" width=500px align="center"><% '--- Table for border --- %>
		<TR>
  		<th>
    		<center><font size=<% =fontsize3 %> color="#FFFFFF"><b> National Rankings For</font>
      		<font size=4 color="white"><br><%=rsRankList("FirstName")%>&nbsp;<%=rsRankList("LastName")%></b></FONT>
      		<font size=2 color="#FFFFFF"><br><b><%=rsRankList("City")%>, <%=rsRankList("State")%></b></FONT>
    		</center>
  		</TH>
		</TR>
		<TR>
  		<TD>  
    		<TABLE class="innertable" align="Center" width="90%">
	    		<tr>
    				<th align="Center" width=20% valign="center"><font size=<% =fontsize2 %> color="#FFFFFF"><b>National<br>Rank</b></FONT></th>
    				<th align="Center" width=30% valign="center"><font size=<% =fontsize2 %>  color="#FFFFFF"><b>Division<br>and Event</b></FONT></th> 	
    				<th align="Center" width=30% valign="center"><font size=<% =fontsize2 %>  color="#FFFFFF"><b>Ranking<br>Score</b></FONT></th> 	
    			</tr>
    			<%
 	
					DO WHILE NOT rsRankList.eof 

							sEvent=TRIM(rsRankList("Event"))

							IF TRIM(rsRankList("RankScore"))<>"" THEN
				  				SELECT CASE TRIM(rsRankList("Event"))
										CASE "J"
											sEventName="Jump"
											sRankScore=formatnumber(rsRankList("RankScore"),1)		
										CASE "S"
											sEventName="Slalom"
											sRankScore=formatnumber(rsRankList("RankScore"),2)		
										CASE "T"
											sEventName="Trick"
											sRankScore=formatnumber(rsRankList("RankScore"),0)
										CASE "O"
											sEventName="Overall"
											sRankScore=formatnumber(rsRankList("RankScore"),0)
	  								END SELECT
        			ELSE
									sEventName="No Ranking"
									sRankScore=formatnumber(0,0)
 							END IF

							%>
    					<tr>
	  						<td align="Center" width=9% valign="top" bgcolor="<%=Tablecolor1%>">
	  							<font size=<% =fontsize2 %> color="#000000"><b><%=rsRankList("RankNum")%></b></font>
	  						</td>
	  						<td align="Center" width=9% valign="top" bgcolor="<%=Tablecolor1%>">
									<font size=<% =fontsize2 %> color="#000000">
		  							<b><a href="/rankings/<%=ThisFileName%>?MyEvent=<%=sEvent%>&MyDiv=<%=rsRankList("Div")%>&RankingListType=<%=sRunByWhat%>&sMemberID=<%=sMemberID%>&SkiYear=<%=sSkiYear%>"><%=rsRankList("Div")%>&nbsp;&nbsp;<%=sEventName%></a></b>
									</font>
								</td>
	  						<td align="Center" valign="top" bgcolor="<%=Tablecolor1%>">
	  							<font size=<% =fontsize2 %>  color="#000000"><b><%=sRankScore%></b></font>
	  						</td> 	
							</tr>
							<%
        
        
        			rsRankList.MoveNext	
    			LOOP 
    			
    			
    			%>
		   <br>
  		</TD>
		</TR>
		</TABLE>
		<%


		'response.write("<br><br>Line 1476 Below primary table ")
		'response.end


		%>
		<TABLE class="innertable" align="center" style="border-width:0px;" width="70%">
			<tr>
  			<td colspan=3 align="center" style="border-style:none;">
					<font size=<% =fontsize2 %> face=<% =font3 %> color="#000000"><b>&nbsp;<br>Rankings include USA skiers in Div/Event.<br> Regional or Filtered rankings may be different.</b></font>
					<br><br>
					<font size=<% =fontsize2 %> face=<% =font3 %> color="#000000"><b>Click on Links above to display Rankings for that Div/Event</b></font>
					<br><br>
  			</td>
			</tr>
			<tr>
  			<td width=20% align="right" style="border-style:none;">
    			<font size="<%=fontsize2%>" face=<% =font1 %> color=<%=Titlecolor%>><b>Range:&nbsp;&nbsp;</b></font>
  			</td>

  			<form method=post action="<%=ThisFileName%>?RankingListType=<%=sRunByWhat%>">
  				<td width=30% align="left" style="border-style:none;">
						<select OnChange=submit() name='RankingsSkiYearIDSelected'>
						<%

							SET rsSelectFields=Server.CreateObject("ADODB.recordset")
							sSQL = "SELECT DISTINCT RT.SkiYearID, SY.SkiYearName"
							sSQL = sSQL + " FROM " &RankTableName&" AS RT"
							sSQL = sSQL + " JOIN " &SkiYearTableName&" AS SY ON RT.SkiYearID = SY.SkiYearID"


							' --- NCWSA does not display 12 Month Rankings
							IF sRunByWhat="NCWSA" THEN
									sSQL = sSQL + " WHERE SY.SkiYearID <> 1"
							END IF
	
							'response.write("<br>Line 1511 - Before Open Statement of Query")
							'response.write("<br>"&sSQL)
							'response.end
							rsSelectFields.open sSQL, SConnectionToTRATable

							' Loads dropdown and sets default to Session("SkiYear")
							DO WHILE NOT rsSelectFields.eof

									IF TRIM(rsSelectFields("SkiYearID")) = TRIM(session("SkiYear")) THEN
											response.write("<option value =""" & rsSelectFields("SkiYearID") &""" selected>")
											response.write(rsSelectFields("SkiYearName"))
											response.write("</option><br>")
									ELSE
											response.write("<option value =""" & rsSelectFields("SkiYearID") &""">")
											response.write(rsSelectFields("SkiYearName"))
											response.write("</option><br>")
									END IF 

									rsSelectFields.moveNEXT
							LOOP

							rsSelectFields.close 
							
						%>
						</select>
   					<input type="hidden" name="SingleRanking" value="Find My Rankings">
   					<input type="hidden" name="sMemberID" value="<%=sMemberID%>">
  				</td>
			</form>

  		<form method=post action="<%=ThisFileName%>?RankingListType=<%=sRunByWhat%>">
  			<td align="center" style="border-style:none;">
					<input type=submit name="NewMember" value="New Member" title="Select another member and Display their Personal Rankings">
			  </td>
  		</form>
		</tr>
		</TABLE>

  </TD>
	</TR>
</TABLE> 
<% 
'--- Table for border --- 

		'response.write("<br><br>Line 1553 Below border table ")
		'response.end




ELSE  
	
	'response.write("<br><br>Line 1561 ELSE Statement of QuerySuccess=true IF ")
	'response.end

	%>
	<br>
	<TABLE class="innertable" align=center width=60% border=3>
	<TR>
	  <th Colspan=2 align=center>
		<font size=4 color="<%=Textcolor5%>"><b> Notice</b></font>
	  </th>
	</TR>

	<TR>
	<TD>
	<TABLE align=center width=90% border=0>

	<TR>
	  <TD Colspan=2 align=center>
		<font size=<% =fontsize3 %> color="<%=Textcolor3%>"><b> No Records Found For This Member In This Range</b></font>
	  <br><br>
	  </TD>
	</TR>

	<TR>

	  <form method=post action="<%=ThisFileName%>?RankingListType=<%=sRunByWhat%>&sMemberID=<%=sMemberID%>">
	    <TD align=center>
				<input type=submit name="Continue" style="width:9em" value="Continue"></center>
	    </TD>
	  </form>

	  <form method=post action="<%=ThisFileName%>?RankingListType=<%=sRunByWhat%>">
	    <TD align=center>
  		<input type=submit name="NewMember" style="width:9em" value="New Member"></center>
	    </TD>
	  </form>
	</TR>
	</TABLE>
	</TD>
	</TR>
	</TABLE>



	<%
END IF

'response.write("<br>Line 1576 Bottom of SUBROUTINE ")
'response.end

END SUB










' -------------------------------------
  SUB DisplayPercentilesandPageFooter
' -------------------------------------


' Writes percentages in text at bottom of list	

SELECT CASE EventSelected
  CASE "S"
    sEventName="Slalom"	
  CASE "T"
    sEventName="Tricks"	
  CASE "J"
    sEventName="Jumping"
  CASE "O"
    sEventName="Overall"
  CASE "WB"
    sEventName="Wakeboard"	
END SELECT


		' Displays the last re-calculation date/time at bottom of screen	
		SET rsSelectFields=Server.CreateObject("ADODB.recordset")
	    	sSQL = "SELECT * FROM " & SkiYearTableName & " WHERE "

		IF session("SkiYear") = "0" THEN
        		sSQL = sSQL + "DefaultYear = 1"
		ELSE
		sSQL = sSQL + "SkiYearID = " + SQLClean(session("skiyear"))
		END IF

		rsSelectFields.open sSQL, SConnectionToTRATable

		IF not rsSelectFields.eof THEN 
				LastRecalcText="Last Recalc: "&rsSelectFields("LastRecalc")
		END IF 


sSQL="SELECT " 
sSQL=sSQL+" (Select COD FROM "&LeagueTableName&" WHERE RIGHT(LEFT(QualifyTour,3),1)='E' AND HomoType='B' AND RIGHT(LeagueID,4)='2010') AS EADate"
sSQL=sSQL+", (Select COD FROM "&LeagueTableName&" WHERE RIGHT(LEFT(QualifyTour,3),1)='M' AND HomoType='B' AND RIGHT(LeagueID,4)='2010') AS MWDate"
sSQL=sSQL+", (Select COD FROM "&LeagueTableName&" WHERE RIGHT(LEFT(QualifyTour,3),1)='C' AND HomoType='B' AND RIGHT(LeagueID,4)='2010') AS SCDate"
sSQL=sSQL+", (Select COD FROM "&LeagueTableName&" WHERE RIGHT(LEFT(QualifyTour,3),1)='S' AND HomoType='B' AND RIGHT(LeagueID,4)='2010') AS SODate"
sSQL=sSQL+", (Select COD FROM "&LeagueTableName&" WHERE RIGHT(LEFT(QualifyTour,3),1)='W' AND HomoType='B' AND RIGHT(LeagueID,4)='2010') AS WEDate"
sSQL=sSQL+" FROM "&LeagueTableName
sSQL=sSQL+" WHERE RIGHT(LeagueID,4)='2010'"
SET rsReg=Server.CreateObject("ADODB.recordset")
rsReg.open sSQL, SConnectionToTRATable

EADate=(rsReg("EADate"))
MWDate=(rsReg("MWDate"))
SCDate=(rsReg("SCDate"))
SODate=(rsReg("SODate"))
WEDate=(rsReg("WEDate"))




IF sRunByWhat = "National" OR sRunByWhat = "Junior" THEN 

	%>
	<div class="coallevels" style="color:white; font-size:10px; text-align:center;">
		<span class="span100">Percentiles and COA Scores For&nbsp; <%=RankingsDivSelected%>&nbsp; <%=sEventName%></span>
	</div>
	<div class="coallevels" style="color:black; font-size:10px; border:1px solid; border-color:white;">		
		<br>
		<span class="span40" style="background-color:<%=scolor09%>">Level 9 - <%=COALevel9%></span>
		<span class="span40" style="background-color:<%=scolor06%>">Level 6 - <%=COALevel6%></span>		
		<br>
		<span class="span40" style="background-color:<%=scolor08%>">Level 8 - <%=COALevel8%></span>		
		<span class="span40" style="background-color:<%=scolor05%>">Level 5 - <%=COALevel5%></span>
		<br>
		<span class="span40" style="background-color:<%=scolor07%>">Level 7 - <%=COALevel7%></span>		
		<span class="span40" style="background-color:<%=scolor04%>">Level 4 - <%=COALevel4%></span>		
		<br>
		<span class="span100" style="color:white; font-size:10px; text-align:center;">Last Recalc<%=LastRecalcText%></span>		
	</div>
	<%
	


END IF

END SUB






%>
