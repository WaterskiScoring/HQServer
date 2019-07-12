<!--#include file="settingsHQ.asp"-->
<!--#include virtual="/rankings/Tools_Include16.asp"-->
<!--#include virtual="/rankings/Tools_Registration16.asp"-->
<!--#include virtual="/rankings/Tools_TournamentListQuery.asp"-->
<!--#include virtual="/rankings/tools_mobile_version.asp"-->
<%




' ------------------------------------------------
' --- Dimensions variables used in this module ---
' ------------------------------------------------

Dim pvar, sTourID
Dim CurrentPage, rsUsers
Dim InRange

Dim MobileTableWidth1
Dim sl, ju, tr, kb, wb, ws, wu, hy, bf, ad, jd, da 

Dim sThisTourDate, ThisDescription, sTCity, sTState


Dim sTourLevel, sTourRange, sSportsGroup, sMonth, StartMonth, EndMonth
Dim MonthColor, s_greenflag, s_yellowflag, s_redflag, ThisFlag, FlagMessage
Dim ThisFileName

Dim sTourIDArray(400), sTSanctionArray(400), sTourNameArray(400), sTCityArray(400), sTStateArray(400)
Dim sTStatusArray(400), ThisDescriptionArray(400)
Dim sThisTourDateArray(400), sTDateSArray(400), sTDateEArray(400)
Dim ListingLinkArray(400), tabcolorArray(400), bottomcolorArray(400)




ThisSitePath="/rankings"
ThisFileName="View-Tournaments_m.asp"

SearchFileName = "search-memberHQ.asp"
RankingsMobileFilename="view-standings_m.asp"
TournamentsMobileFilename="view-tournaments_m.asp"
TeamsMobileFilename="virtualteam_m.asp"
LocalVarFileName="Test_localstorage_SET.asp"
MenuFileName = "mainmenu_m.asp"



' --- Flag images for status of tournament ---
s_greenflag = "images\buttons\Flag-green16.png"
s_yellowflag = "images\buttons\Flag-yellow16.png"
s_redflag = "images\buttons\Flag-red16.png"



' ------------------------------------
' --- Reads NVP's from querystring ---
' ------------------------------------

Read_Form_Variables



CheckForMobileDevice()
' response.write("<div style=color:red>OperatingSystem = "&OperatingSystem&"</div>")



' ------------------------------------------------------------
' --- MAIN BRANCHING between listing and single tournament ---
' ------------------------------------------------------------

' --- Displays the html, head and opening body tag ---


'DisplayHeadOpenBodyAndBannerTags
OpenState="tournaments"
DisplayHeadOpenBodyAndBannerTags OpenState


				
' --- Displays the menu for view tournaments --- 
'DisplayMenuButtons_ViewTournaments




SELECT CASE pvar
		CASE "TourInfo"
		
				' --- Displays the details for s single tournament ---
				DisplaySingleTournamentDetails
				
		CASE ELSE
				
				' --- Displays the search filter for settings - initially hidden ---
				DisplayFilters="none"
				IF TRIM(Request("df"))="yes" THEN DisplayFilters="inline"
				
				DisplayTournamentSearchFilters

				' --- Runs the query to select tournaments ---
				Load_Tournament_Query	

				
				IF rs.eof AND TRIM(Request("df"))<>"yes" THEN
						DisplayNoRankingsFoundForFilter_Message

				ELSEIF TRIM(Request("df"))="yes" THEN 
						' --- Do nothing --
				ELSE	
						' --- Displays the tournament listing in the scroll box ---
						DisplayTournamentListing
				END IF


				
END SELECT		




' --- Writes the Closing tags for HTML ---
DisplayCloseBodyAndHTMLTags






' --------------------------------------------------------------------------------------------------------------
' --- BOTTOM OF MAIN CODE ---
' --------------------------------------------------------------------------------------------------------------










' --------------------------------
  SUB DisplayTournamentListing 
' --------------------------------  

%>
<div id="tournamentlisting" >
	<a href="javascript:DisplaySearchFilters('tournamentsearchfilters');" style="text-decoration:inline; text-decoration:none;">
		<div class="searchimagediv" style="margin:0px 0px 0px 0px; background: url(images/buttons/LongButtonBlank.png); background-position:center; background-repeat: no-repeat; text-decoration:none; text-align:center; border:0px solid yellow;">
			<p class="searchbannerline" style="margin:0px 0px 0px 0px;" >GO TO SEARCH SETTINGS</p>
		</div>
	</a>
	<div class="scroll">
	<%

	' --- Loops thru the tournaments selected filtering out some based on other criteria ---
	LoopThru_Tournament_Listing

	%>
	<br><br><br>
	</div> 		<! -- Bottom of scroll box -- >
</div> 			<! -- Bottom of div for hidding and displaying - TournamentListing --  >
<%

END SUB




' -------------------------------------
  SUB DisplayTournamentSearchFilters
' -------------------------------------

%>
<div id="toursearchsettings" style="display:<%=DisplayFilters%>">
	<input type="hidden" id="sMemberID_Hidden_InRankingsSettings" name="sMemberID" value="">
	<input type="hidden" id="sName_InRankingsSettings" name="sName_InRankingsSettings">	
		<%
		
		' --- Displays the filter dropdowns inside ---
		Display_Filters_New   
		
		%>
</div> <! -- Bottom of div for hidding and displaying -- >
<%



END SUB



' ---------------------------------------------
  SUB DisplayNoRankingsFoundForFilter_Message
' ---------------------------------------------  

' --- NOT USED ???
%>
<div id="tournamentlisting" >
	<a href="javascript:DisplaySearchFilters('tournamentsearchfilters');" style="text-decoration:none;">
		<div class="searchimagediv" style="background: url(images/buttons/LongButtonBlank.png); background-repeat: no-repeat;">
			<p class="searchbannerline">RETURN TO SEARCH SETTINGS</p>
		</div>
	</a>
	<div class="error" style="height:310px; padding-top:40px; color:#FFFFFF; text-align:center; font-size:12pt;">
		<span id="" class="span100">No Rankings Found For These Settings.</span>
	</div>	

</div>
<% 	


END SUB




' ---------------------------------------
  SUB Display_NoRecords_ErrorMessage
' ---------------------------------------  

%>
<div id="notournamentselectederror" style="display:inline">
	<a href="javascript:DisplaySearchFilters('norecordserror');" style="text-decoration:inline;">
		<div class="searchimagediv" style="background: url(images/buttons/LongButtonBlank.png); background-repeat: no-repeat; text-decoration:none;">
			<p class="searchbannerline">RETURN TO SEARCH SETTINGS</p>
		</div>
	</a>

	<div class="errorbox" style="padding-top:30px; min-height:280px; padding-left:0px;">
		<div style="margin-top:10px;">
			<span id="" class="span10" style="border:0px solid; border-color:white; vertical-align:top;">
				<img src="images/buttons/Button-Info-icon.png" style="padding:0px; width:30px; text-align:right; " alt="Tip" />
			</span>
			<span id="" class="span80" style="border:0px solid; border-color:white;">
				To update Tournament search use <br>GO TO SEARCH SETTINGS<br>button above
			</span> 
			<span id="" class="span100" style="text-align; margin-top:20px;">
				<img src="images/logos/GEICO.jpg" style="padding:0px; width:200px; text-align:right;" alt="Geico_Ad" />
			</span>
		</div>
		<div class="" style="margin-top:40px;">
				<input type="button" name="Start" value="START" style="width:7em; height:2em; font-size:14pt" onclick="javascript:DisplaySearchFilters('norecordserror');">
		</div>

	</div>	


</div>
<% 	


END SUB















' ------------------------------------
  SUB DisplaySingleTournamentDetails
' ------------------------------------

ReturnLinkURL=""&ThisFileName&"?sSportsGroup="&sSportsGroup&"&sTourRange="&sTourRange&"&sTourLevel="&sTourLevel&"&State="&sTourState&"&Region="&sTourRegion&"&sClass="&sClass&"&StartMonth="&StartMonth&"&EndMonth="&EndMonth
' ListingLinkArray(TourCount)=""&ThisFileName&"?pvar=TourInfo&TourID="&sTourID&"&sSportsGroup="&sSportsGroup&"&sTourRange="&sTourRange&"&sTourLevel="&sTourLevel&"&State="&sTourState&"&Region="&sTourRegion&"&sClass="&sClass&"&StartMonth="&StartMonth&"&EndMon

%>
<div id="tourdetails">
	<input type="hidden" id="sMemberID_Hidden_InRankingsSettings" name="sMemberID" value="">
	<input type="hidden" id="sName_InRankingsSettings" name="sName_InRankingsSettings">	
	<a href="<%=ReturnLinkURL%>" style="text-decoration:none;">
		<div class="searchimagediv" style="background: url(images/buttons/LongButtonBlank.png); background-position:center; background-repeat: no-repeat; text-align:center">
			<p class="searchbannerline" style="margin:0px 0px 0px 0px;">GO TO TOURNAMENT LIST</p>
		</div>
	</a>
	
	<div class="scroll">
	<%

	' --- Displays the details for a single tournament ---
	DisplaySingleListingDetails

	%>
	<br><br><br>
	</div> <! -- Bottom of scroll box -- >
</div> <! -- Bottom of div for hidding and displaying - TourListing --  >
<%


END SUB




' ---------------------------------
	SUB LoopThru_Tournament_Listing
' ---------------------------------	
	

	counter = 0
	rs.movefirst
	DO WHILE NOT rs.eof

			' --------------------------------------------------------------------------
			' --- Determines all the conditions where a data line could be displayed ---
			' --------------------------------------------------------------------------
			DispDataLineYorN = "N"
			IF adminmenulevel > 19 THEN
					DispDataLineYorN = "Y"
			ELSEIF (rs("SptsGrpID")="AWS" AND rs("STRegion")="D") THEN
					DispDataLineYorN = "Y"				
			ELSEIF (rs("SptsGrpID")="ABC" AND rs("STRegion")="B") THEN
					DispDataLineYorN = "Y"			
			ELSEIF (rs("SptsGrpID")="USH" AND rs("Pending")=0) THEN
					DispDataLineYorN = "Y"			
			ELSEIF (rs("Pending") = 0 AND rs("ShowPSched") <> 0 AND rs("ShowRegistrar") <> 0 AND ( (rs("TSTATUS") > 0 AND rs("TSTATUS") <> 3) OR ( rs("GBPolicy") <> 0 AND (rs("TKitOKGuideBookAd") OR rs("OK2Publish")) <> 0) ) ) THEN
					DispDataLineYorN = "Y"	
			ELSEIF rs("TSantype") = "6" AND rs("OK2Publish") = true THEN
					DispDataLineYorN = "Y"
			END IF
			
					
			' ---------------------------------------------------
			' --- Displays as single line of data for listing ---
			' ---------------------------------------------------

			IF DispDataLineYorN = "Y" THEN
					TourCount=TourCount+1

					' --- Listing tournaments go here ---
    			Define_TournamentValues_ForListing
			END IF

			rs.movenext
	LOOP

		

END SUB








' -----------------------------------------
  SUB Define_TournamentValues_ForListing
' -----------------------------------------

	' ---------------------------------------------------------------
	' --- CONTROLS FOR BUTTONS FOR ONLINE ENTRY, ADMIN LOGIN, ETC ---
	' ---------------------------------------------------------------

	' --- TStatus > 1 means the tournament has been sanctioned by HQ
	' --- sAllowRegistrationsCheck="on" is set at the top of this file as an override
 	' --- adminmenulevel is defined by user profile in SWIFT LOGIN
	' --- TestValidAdminCode is a function in Tools_Registration that verifies AdminCode for this tournament

	sTourID = rs("TournAppID")

	sTourIDArray(TourCount) = rs("TournAppID")
	sTSanctionArray(TourCount) = rs("TSanction")
	IF left(sTSanctionArray(TourCount),6) <> sTourIDArray(TourCount) THEN sTSanctionArray(TourCount) = sTourID & "-"
	sTourNameArray(TourCount)=LEFT(rs("TName"),29)
	sTCityArray(TourCount)=rs("TCity")
	sTStateArray(TourCount)=rs("TState")



	' --- Adds all descriptions into one string with breaks between SptsGrpID's ---
	IF TRIM(rs("TDescription"))<>"" THEN ThisDescriptionArray(TourCount) = TRIM(rs("TDescription")) + "<br>" 
	IF TRIM(rs("FDescription"))<>"" THEN ThisDescriptionArray(TourCount) = ThisDescriptionArray(TourCount) + TRIM(rs("FDescription")) + "<br>" 
	IF TRIM(rs("WDescription"))<>"" THEN ThisDescriptionArray(TourCount) = ThisDescriptionArray(TourCount) + TRIM(rs("WDescription")) + "<br>" 
	IF TRIM(rs("KDescription"))<>"" THEN ThisDescriptionArray(TourCount) = ThisDescriptionArray(TourCount) + TRIM(rs("KDescription")) + "<br>" 
	IF TRIM(rs("CDescription"))<>"" THEN ThisDescriptionArray(TourCount) = ThisDescriptionArray(TourCount) + TRIM(rs("CDescription")) + "<br>" 

	sTStatusArray(TourCount)=rs("TStatus") 
	
	' --- Sets the color of the tab and outline --- 
	SetTabAndBottomColor TourCount


	' --- PayPal and OLR status ---	  
	sPayPalOK=rs("PayPalOK") 
	sPayPalAct=rs("PayPalAct")
	sUseOLReg=rs("UseOLReg")
	sOLR_PD=rs("OLR_PD")

	' --- Gets Start and End Date and formats into date-range ---	
	sTDateEArray(TourCount)=rs("TDateE")
	sTDateSArray(TourCount)=rs("TDateS")
	IF sTDateSArray(TourCount)=sTDateEArray(TourCount) THEN 
			sThisTourDateArray(TourCount)=Month(sTDateSArray(TourCount))&"/"&Day(sTDateSArray(TourCount))&"/"&RIGHT(Year(sTDateEArray(TourCount)),2)
	ELSE 
			sThisTourDateArray(TourCount)=Month(sTDateSArray(TourCount))&"/"&Day(sTDateSArray(TourCount))&"-"&Day(sTDateEArray(TourCount))&"/"&RIGHT(Year(sTDateEArray(TourCount)),2)
	END IF 



	' --- Define what color to use based on the month so it alternates ---
	SELECT CASE Month(sTDateSArray(TourCount))
		CASE 1,5,9
				MonthColor="#EEDDDD"
		CASE 2,6,10
				MonthColor="#CCCCFF"
		CASE 3,7,11
				MonthColor="#FFFF66"
		CASE 4,8,12
				MonthColor="#CCFFCC"
	END SELECT 

	IF sTDateSArray(TourCount) = sTDateEArray(TourCount) THEN 
			DisplayDate = Month(sTDateS) & "/" & Day(sTDateS) & "/" & RIGHT(cStr(Year(sTDateS)),2)
	ELSE
			DisplayDate = Month(sTDateS) & "/" & Day(sTDateS) & "-" & Day(sTDateE) & "/" & RIGHT(cStr(Year(sTDateS)),2)
	END IF


	sSL_Offered="N"
	sTR_Offered="N"
	sJU_Offered="N"

	sWB_Offered="N"
	sWS_Offered="N"
	sWU_Offered="N"

	sBF_Offered="N"
	sKB_Offered="N"
	sHY_Offered="N"

	sDA_Offered="N"
	sJD_Offered="N"
	sAD_Offered="N"

	' --- Begins in 2010
	IF sTourRange = "0" OR sTourRange = "1" OR sTourRange = "2" OR sTourRange >= "5" THEN
			IF sl="on" AND ( ( rs("sClassC") + rs("sClassE") + rs("sClassL") + rs("sClassR") + rs("sClassCash") + rs("sClassX") > 0 )  OR rs("Gr2AWS_SPulls")<>0 OR rs("Gr1AWSPulls") ) THEN sSL_Offered="Y"
			IF tr="on" AND (  (rs("tClassC") + rs("tClassE") + rs("tClassL") + rs("tClassR") + rs("tClassCash") + rs("tClassX") > 0 )  OR rs("Gr2AWS_TPulls")<>0  ) THEN sTR_Offered ="Y"
			IF ju="on" AND (  (rs("jClassC") + rs("jClassE") + rs("jClassL") + rs("jClassR") + rs("jClassCash") + rs("jClassX") > 0 )  ) THEN sJU_Offered ="Y"

			IF wb="on" AND (  rs("WWakeW")>0 OR rs("Gr2USW_WPulls")<>0 OR rs("Gr2USW_RailJamPulls")<>0 OR rs("Gr1USWPulls") <> 0 ) THEN sWB_Offered="Y"
			IF ws="on" AND rs("WSkateW")>0 OR rs("Gr2USW_SkatePulls") THEN sWS_Offered="Y"
			IF wu="on" AND rs("WSurfW")>0 OR rs("Gr2USW_SurfPulls") THEN sWU_Offered="Y"

			IF bf="on" AND (  rs("SptsGrpID")="ABC" OR rs("Gr1ABCPulls")<>0  ) THEN sBF_Offered="Y" 
			IF kb="on" AND (  rs("SptsGrpID")="AKA" OR rs("Gr2AKA_SPulls")<>0 OR rs("Gr2AKA_TPulls")<>0 OR rs("Gr2AKA_FreePulls")<>0 OR rs("Gr2AKA_FlipPulls")<>0 OR rs("Gr1AKAPulls")<>0  ) THEN sKB_Offered="Y"	
			IF hy="on" AND (  rs("TEventFHF")<>0 OR rs("Gr2USH_FreeRidePulls")<>0 OR rs("Gr2USH_JumpOutPulls")<>0 OR rs("Gr2USH_BigAirPulls")<>0 OR rs("Gr2USH_3TrickPulls")<>0 OR rs("Gr1USHPulls")<>0  ) THEN sHY_Offered="Y"

			IF da="on" AND (rs("Gr1WSDPulls")) THEN sDA_Offered="Y"
			IF jd="on" AND (rs("JDClin")<>0) THEN sJD_Offered="Y"
			IF ad="on" AND (rs("ADClin")<>0) THEN sAD_Offered="Y"

	ELSE
			IF sl="on" AND (rs("TEventSlalom")<>0 OR rs("TEventF3ev")<>0 OR rs("Gr2AWS_SPulls")<>0 OR rs("Gr1AWSPulls")<>0) THEN sSL_Offered ="Y"
			IF tr="on" AND (rs("TEventTrick")<>0 OR rs("Gr2AWS_TPulls")<>0) THEN sTR_Offered ="Y"
			IF ju="on" AND (rs("TEventJump")<>0) THEN sJU_Offered ="Y"

			IF wb="on" AND (rs("TEventWake")<>0 OR rs("TEventFW")<>0 OR rs("Gr2USW_WPulls")<>0 OR rs("Gr2USW_RailJamPulls")<>0 OR rs("Gr1USWPulls")<>0) THEN sWB_Offered="Y"
			IF ws="on" AND (rs("TEventWSkate")<>0 OR rs("Gr2USW_SkatePulls")<>0) THEN sWS_Offered="Y"
			IF wu="on" AND (rs("TEventWSurf")<>0 OR rs("Gr2USW_SurfPulls")<>0) THEN sWU_Offered="Y"

			IF bf="on" AND (rs("SptsGrpID")="ABC" OR rs("TEventNBL")<>0 OR rs("Gr1ABCPulls")<>0) THEN sBF_Offered="Y" 
			IF kb="on" AND (rs("SptsGrpID")="AKA" OR rs("TEventFKB")<>0 OR rs("Gr2AKA_SPulls")<>0 OR rs("Gr2AKA_TPulls")<>0 OR rs("Gr2AKA_FreePulls")<>0 OR rs("Gr2AKA_FlipPulls")<>0 OR rs("Gr1AKAPulls")<>0) THEN sKB_Offered="Y"	
			IF hy="on" AND (rs("TEventFHF")<>0 OR rs("Gr2USH_FreeRidePulls")<>0 OR rs("Gr2USH_JumpOutPulls")<>0 OR rs("Gr2USH_BigAirPulls")<>0 OR rs("Gr2USH_3TrickPulls")<>0 OR rs("Gr1USHPulls")<>0) THEN sHY_Offered="Y"

			IF da="on" AND (rs("Gr1WSDPulls")) THEN sDA_Offered="Y"
			IF jd="on" AND (rs("JDClin")<>0) THEN sJD_Offered="Y"
			IF ad="on" AND (rs("ADClin")<>0) THEN sAD_Offered="Y"
	END IF

	'----------------------------------------------------------------
	' --- Determines whether or not to GREY out OLR button ---
	'----------------------------------------------------------------

	OLRButtonEntryStatus="enabled"
	EntryButtonTitle="Enter this tournament with our online entry form"

	
	' --- Determines if corresponding record exists in Ski Year table - avoids errors until record is entered --
	set rsSkiYear=Server.CreateObject("ADODB.recordset")
	rsSkiYearSQL = "SELECT * FROM "&SkiYearTableName&" WHERE EndDate>='"&sTDateE&"' AND BeginDate<='"&sTDateS&"'" 
	rsSkiYear.open rsSkiYearSQL, SConnectionToTRATable
	
	whiletesting=0	
	
	IF rsSkiYear.eof THEN 
			OLRButtonEntryStatus="disabled"
			EntryButtonTitle="Ski Year Administrative table must be updated for this Ski Year before available for OLR "
	ELSEIF whiletesting=1 AND adminmenulevel < 50 AND RIGHT(cStr(Year(sTDateS)),2)>="16" THEN 
  		OLRButtonEntryStatus="disabled"
			EntryButtonTitle="Online Registration has not been activated for 2016 - Thanks for your patience"
	ELSEIF NOT(rs("OLRDisplayStatus")) THEN
			sExclude="no"
			OLRButtonEntryStatus="disabled"
			EntryButtonTitle="This tournament is no longer open to Online Registration"
	ELSEIF (rs("UseOLReg")=true AND (rs("PayPalOK")=0 OR rs("OLR_PD")=0 OR TRIM(rs("PayPalAct"))="" ) ) THEN
			sExclude="no"
			OLRButtonEntryStatus="disabled"
			EntryButtonTitle="This tournament is not yet available for Online Registration"
	END IF


	' --- OLR Button status ---
	IF DisableOLRButtons=true THEN
			OLRButtonEntryStatus="disabled"
			EntryButtonTitle="Online Registration is Temporarily Disabled - Thanks for your patience"
	END IF


	' --- Determines which OLR program is used ---  
  IF RIGHT(cStr(Year(sTDateS)),2)>="16" THEN 
  		RegFileForLink = "registration16.asp"
  ELSE
  		RegFileForLink = "registration.asp"
  END IF		




	' --- Defines whether there is a link in the tournament name
	IF sTStatusArray(TourCount)<>"3" THEN 
			ListingLinkArray(TourCount)=""&ThisFileName&"?pvar=TourInfo&TourID="&sTourID&"&sSportsGroup="&sSportsGroup&"&sTourRange="&sTourRange&"&sTourLevel="&sTourLevel&"&State="&sTourState&"&Region="&sTourRegion&"&sClass="&sClass&"&StartMonth="&StartMonth&"&EndMonth="&EndMonth
	ELSE
			ListingLinkArray(TourCount)=""
	END IF 


	' --------------------------------------------
	' --- Displays a tournament in the listing ---
	' --------------------------------------------
	DisplayTournamentInListing



END SUB




' --------------------------------
  SUB DisplayTournamentInListing
' --------------------------------

		' --- Puts a colored tab before the first record of the month ---
		IF sMonth <> Month(sTDateSArray(TourCount)) THEN 
				sMonth = Month(sTDateSArray(TourCount)) 
				ThisMonthName = MonthName(MONTH(sTDateSArray(TourCount))) & " - " & YEAR(sTDateSArray(TourCount)) 
				' --- Put a blank row in to separate from heading ---
				%>
				<div class="tabtour" style="font-size:14pt; height:20px; margin-top:10px; background-color:<%=MonthColor%>; text-color:<%=Textcolor1%>;">
	  			<span class="span100" style="text-align:left;"><b><%= ThisMonthName %></b></span>
				</div>
				<%
		END IF 
		
		%>
		<div class="tabtour" style="height:18px; background:<%=MonthColor%>;">
			<a href="<%=ListingLinkArray(TourCount)%>" title="Check details for <%= sTourNameArray(TourCount)%>" style="text-decoration:none;">
				<span class="span70" style="width:70%; margin:0px 0px 0px 0px; padding:0px 0px 0px 0px; border:0px solid black;">
					<%= LEFT(sTourNameArray(TourCount),27)%>
				</span>	
			</a>
			<span class="span25" style="width:28%; margin:0px 0px 0px 0px; padding:0px 0px 0px 0px; text-align:right; border:0px solid red;"><%=sThisTourDateArray(TourCount)%></span>
		</div>

		<div class="tourbody" style="vertical-align:top; font-size:9pt;">
			<span class="span85" style="border:0px solid black;"><%=ThisDescriptionArray(TourCount)%></span>
			<span class="span10" style="vertical-align:top; text-align:right; border:0px solid black;"><img src="<%=ThisFlag%>" title="<%=FlagMessage%>" height="15px" width="15px"></span>		
		</div>
		<div class="tourbottom">
			<span class="span60" style="font-size:10pt"><%=sTCityArray(TourCount)%>, <%=sTStateArray(TourCount)%></span>	
			<span class="span40" style="width:32%; text-align:right;"><%=sTourIDArray(TourCount)%></span>	
		</div>
	<%

END SUB















' -----------------------------------
    SUB DisplaySingleListingDetails
' -----------------------------------

		TourCount=1

		sTourID = TRIM(Request("TourID"))
		sExclude = Request("sExclude")


		' --- Uses variable definition from Tools_Registration.asp ---
		DefineTourVariables_New


		' --- Reassign to array variable for future design -- 
		sTourNameArray(TourCount)=LEFT(sTourName,30)
		sTSanctionArray(TourCount)=sTournAppID
		sTStatusArray(TourCount)=sTStatus

		OLRButtonEntryStatus="enabled"
		EntryButtonTitle="Enter this tournament with our online entry form"


		' --- Tournament reached it's entry limit ---
		IF request("olrds") ="disabled" and sTStatusArray(TourCount)>=0 THEN 
				OLRButtonEntryStatus="disabled"
				EntryButtonTitle="This tournament is no longer open to Online Registration"
		ELSEIF sUseOLReg=true AND sPayPalOK=0 OR sOLR_PD=0 OR TRIM(sPayPalAct)="" THEN
				sExclude="no"
				OLRButtonEntryStatus="disabled"
				EntryButtonTitle="This tournament is not yet available for Online Registration"
		END IF

		' --- Control of OLR button enabled/disabled --- 
		IF DisableOLRButtons=true THEN
				OLRButtonEntryStatus="disabled"
				EntryButtonTitle="Online Registration is Temporarily Disabled - Thanks for your patience"
		END IF


		

		' --- Sets the color of the tab and outline --- 
		SetTabAndBottomColor TourCount

		' --- Gets Start and End Date and formats into date-range ---	
		sTDateEArray(TourCount)=sTDateE
		sTDateSArray(TourCount)=sTDateS
		IF sTDateSArray(TourCount)=sTDateEArray(TourCount) THEN 
				sThisTourDateArray(TourCount)=Month(sTDateS)&"/"&Day(sTDateS)&"/"&RIGHT(Year(sTDateE),2)
		ELSE 
				sThisTourDateArray(TourCount)=Month(sTDateS)&"/"&Day(sTDateS)&"-"&Day(sTDateE)&"/"&RIGHT(Year(sTDateE),2)
		END IF 

		IF TRIM(sTDescription)<>"" THEN
				ThisDescriptionArray(TourCount)=sTDescription
		ELSEIF TRIM(sFDescription)<>"" THEN 
				ThisDescriptionArray(TourCount)=sFDescription
		ELSEIF TRIM(sWDescription)<>"" THEN 
				ThisDescriptionArray(TourCount)=sWDescription
		ELSEIF TRIM(KTDescription)<>"" THEN 
				ThisDescriptionArray(TourCount)=sKDescription
		ELSEIF TRIM(sTDescription)<>"" THEN 
				ThisDescriptionArray(TourCount)=sCDescription
		END IF 


		' --- Remps true/false value of sTOpenClosed for display ---
		IF sTOpenClosed THEN 
				sTOpenClosed="Closed" 
		ELSE 
				sTOpenClosed="Open"
		END IF

		' --- Remaps sMaxPulls value to clarify for display ---
		IF sMaxPulls>0 THEN 
				MaxPullsText = CStr(sMaxPulls)+ " Pulls"
		ELSEIF sTEntryLimit <> "None" THEN
				MaxPullsText = sTEntryLimit
		END IF

		' --- Remaps Late fee to clarify for display ---
		IF sTLFPerDay=true THEN 
	 			sTLateFeeText=sTLateFee&" Per Day"
	 	ELSE  
	 			sTLateFeeText=sTLateFee&" Total"
	 	END IF 





' -----------------------------
' --- Google Map Parameters ---
' -----------------------------
Dim ToMapLocation, FromMapLocation





ToMapLocation = ""
'sSiteAddress = "14351 Isleview Dr"
'sSiteZip = "34787"
'sSiteLatitude = "28.449679"  
'sSiteLongitude = "-81.603401"

' -- Okeeheelee FL -- 26.655669, -80.16763
' -- Isles of Lake Hancock (USAS0378) -- 28.449679, -81.603401
' http://www.usawaterski.org/sanctions/sites/SiteEdit.aspx?Site=USAS0378



' -------------------------------------------
' --- Defines locations for Map function URL ---
' -------------------------------------------

FromMapLocation = "saddr=My%20Location"
MapToSiteStatus="disabled"

IF TRIM(sSiteLatitude)<>"" AND TRIM(sSiteLongitude)<>"" THEN 
		ToMapLocation = "daddr="&sSiteLatitude&","&sSiteLongitude
		MapToSiteStatus="enabled"
ELSEIF TRIM(sSiteAddress)<>"" AND TRIM(sSiteZip)<>"" THEN
		ToMapLocation = "daddr="&Replace(sSiteAddress, " ", "+")&"+"&Replace(sSiteZip, " ", "+")
		MapToSiteStatus="enabled"
END IF		





MapToSiteButtonTitle="Click to Display Map to Site"


' --- Defines which map to use and builds URL for link --
WhichMapApplication = "http://www.maps.google.com/maps"		

SELECT CASE Session("OperatingSystem")
		CASE "Android"
				WhichMapApplication = "http://www.maps.google.com/maps"
		CASE "iOS" 
				WhichMapApplication = "http://maps.apple.com/"
		CASE "Windows" 
				WhichMapApplication = "http://www.maps.google.com/maps"
END SELECT

FullMappingURL = WhichMapApplication& "?"&FromMapLocation&"&"&ToMapLocation&"&dirflg=d"




		' ----------------------------------------------------
		' --- Begins display of single tournament destails ---
		' ----------------------------------------------------

		%>
		<div class="tabtour" style="font-size:12pt; height:40px; background:<%=tabcolorArray(TourCount)%>;">
			<span class="span70"><%=sTourNameArray(TourCount)%></span>	
			<span class="span25" style="text-align:right;"><%=sThisTourDateArray(TourCount)%></span>	
		</div>
		
		<div class="tourdetails" style="background-color:#FFFFFF;">
			<div class="tourdetailline">
				<span class="span25">Sanction:</span>
				<span class="span70"><%=sTSanctionArray(TourCount)%></span>
			</div>
			<div class="tourdetailline">
				<span class="span25">SiteID:</span>
				<span class="span70"><%=sTSiteID%></span>
			</div>
			<div class="tourdetailline">
				<span class="span25" style="font-weight:bold;">City/ST:</span>
				<span class="span70"><%=sTourCity%>, <%=sTourState%></span>
			</div>
			<div class="tourdetailline">
				<span class="span25" style="font-weight:bold;">Description:</span>
				<span class="span70"><%=ThisDescriptionArray(TourCount)%></span>
			</div>
			<div class="tourdetailline">
				<span class="span25">Sponsor:</span>
				<span class="span70"><%=sTSponsor%></span>
			</div>


			<div class="tourdetailline">
				<span class="span25">Site:</span>
				<span class="span70"><%=sTSite%></span>
			</div>

			<div class="tourdetailline">
				<span class="span25" style="vertical-align:top; border:0px solid; border-color:blue; padding:0px; margin:0px;">Directions:</span>
				<span class="span65" style="border:0px solid; border-color:blue; padding:0px; margin:0px; text-align:top;"><%=sGTSDirections%></b></span>
			</div>

			<div class="tourdetailline">	 			
				<%


				' --------------------------
				' --- Link to Google Map ---
				' --------------------------		
				
'response.write("<br>sTStatusArray(TourCount)=0 - "&sTStatusArray(TourCount))
'response.write("<br>TRIM(ToMapLocation)<>null =")
'response.write(TRIM(ToMapLocation)<>"")


	 			'IF sTStatusArray(TourCount) >= 0 AND TRIM(ToMapLocation)<>"" THEN 
	 			IF sTStatusArray(TourCount) >= 0 THEN
						%>
						<span class="span50" style="border:0px solid; border-color:red; padding:0px; margin:0px; text-align:center;">
							<form action="<%=FullMappingURL%>" method="post" target="_blank">
								<input type="submit" style="width:7em; height:2em; font-size:9pt;" value="Map To Site" title="<%=MapToSiteTitle%>" <%=MapToSiteStatus%>>
							</form>
						</span>
						<%
				END IF		


				
				' -----------------------------------
				' --- Online REGISTRATION button ---
				' -----------------------------------


		process="olr" 
		tempadminlevel=adminmenulevel
		adminmenulevel=50
		sUseOLReg=True
		sPayPalAct="mark.bogusemail.com"

			
	 			IF process<>"viewreg" AND sTStatusArray(TourCount) >= 0 AND (sAllowRegistrationsCheck="on" OR adminmenulevel>=50) AND sUseOLReg=True AND sPayPalAct<>"" AND ( sTDateE>=Date OR adminmenulevel>=20 ) THEN 
					%>
					<span class="span45" style="text-align:center; border:0px solid; border-color:red; padding:0px; margin:0px; text-align:top;">
						<form action="<%=RegFileForLink%>" method="post" target="_blank">
							<input type="submit" style="width:7em; height:2em; font-size:9pt;" value="Enter Now" title="<%=EntryButtonTitle%>" <%=OLRButtonEntryStatus%>>
							<input type="hidden" name="sTourID" value="<%=sTourID%>">
						</form>
					</span>
					<%
				END IF		
				adminmenulevel=tempadminlevel



				%>
				</div>

			<div class="tourdetailline">
				<span class="span100"><hr width="90%"></span>
			</div>
			<div class="tourdetailline">
				<span class="span25">Entry:</span>
				<span class="span25"><%=sTOpenClosed%></span>
				<%



					
			' --- Entry Fees and deadlines ---						
			%>
			</div>
			<div class="tourdetailline">
				<span class="span25">Entry Limit:</span>
				<span class="span70"><%=MaxPullsText%></span>
			</div>

			<div class="tourdetailline" style="vertical-align:top;">
				<span class="span25" style="vertical-align:top;">Entry Fees:</span>
				<span class="span70"><%=sTEntryFees%></span>
			</div>
			<div class="tourdetailline">
				<span class="span25" style="vertical-align:top;">Deadline:</span>
				<span class="span70"><%=sTLateDate%></span>
			</div>
			<div class="tourdetailline">
				<span class="span25">Late Fee:</span>
				<span class="span70"><%=sTLateFeeText%></span>
			</div>
			<%











			' --- Tournament Registrar ---
			%>
			<div class="tourdetailline">
				<span class="span100"><hr width="90%"></span>
			</div>


			<div class="tourdetailline">
				<span class="span25" style="vertical-align:top;">Entries To:</span>
				<span class="span70"><%=sTRegistrarName%>
						<br><%=sTRegistrarAddr%>
						<br><%=sTRegistrarCity%>, <%=sTRegistrarState%>&nbsp;<%=sTRegistrarZip%>
						<br>
						<br><a href="tel:<%=sTRegistrarPhone%>" style=text-decoration:none;><%=sTRegistrarPhone%></a>
						<br>
						<br>
						<a href="mailto:<%=sTRegistrarEmail%>?subject=<%=sTourNameArray(TourCount)%>" style="text-decoration:none;"><%=sTRegistrarEmail%></a>
				</span>
			</div>


			<div class="tourdetailline">
				<span class="span100"><hr width="100%"></span>
			</div>
			<div class="tourdetailline">
				<span class="span25" style="vertical-align:top;">Lodging:</span>
				<span class="span70"><%=sGTAccommodation%></span>
			</div>
			<div class="tourdetailline">
				<span class="span25" style="vertical-align:top;">Awards:</span>
				<span class="span70"><%=sGTAwards%></span>
			</div>

			<div class="tourdetailline">
				<span class="span25"style="vertical-align:top;">Practice:</span>
				<span class="span70"><%=sGTPractice%></span>
			</div>
			<div class="tourdetailline">
				<span class="span25">Start Time:</span>
				<span class="span70"><%=sGTStartTime%></span>
			</div>
			<div class="tourdetailline">
				<span class="span25" style="vertical-align:top;">Schedule:</span>
				<span class="span70"><%=sGTSofE%></span>
			</div>
			<div class="tourdetailline">
				<span class="span25">Entry Reqts:</span>
				<span class="span70"><%=sG_IWWF_req%></span>
			</div>
			<div class="tourdetailline">
				<span class="span25" style="vertical-align:top;">Comments:</span>
				<span class="span70"><%=sGTComments%></span>
			</div>
	
			<% ' ------------ BEGIN OFFICIALS SECTION ---------------- %>

			<div class="tourdetailline">
				<span class="span100"><hr width="100%"></span>
			</div>

			<div class="tourdetailline">
				<span class="span25">Tour Dir:</span>
				<span class="span70"><%=sTDirName%></span>
			</div>
			<div class="tourdetailline">
				<span class="span25">Chief Judge:</span>
				<span class="span70"><%=sCJudge%></span>
			</div>
			<div class="tourdetailline">
				<span class="span25">Chief Scorer:</span>
				<span class="span70"><%=sCScorer%></span>
			</div>
			<div class="tourdetailline">
				<span class="span25">Chief Driver:</span>
				<span class="span70"><%=sCDriver%></span>
			</div>
			<div class="tourdetailline">
				<span class="span25">Chief Safety:</span>
				<span class="span70"><%=sCSafety%></span>
			</div>
			<div class="tourdetailline">
				<span class="span25">Announcer:</span>
				<span class="span70"><%=sAnnouncer%></span>
			</div>

			<div class="tourdetailline">
				<span class="span25">Tech Cntr:</span>
				<span class="span70"><%=sTechCont%></span>
			</div>
			<div class="tourdetailline">
				<span class="span25">PanAm Judg:</span>
				<span class="span70"><%=sPanAmJudge%></span>
			</div>
			<div class="tourdetailline">
				<span class="span25">Apt Judges:</span>
				<span class="span70"><%=sAp1Judge%></span>
			</div>
			<div class="tourdetailline">
				<span class="span25">&nbsp;</span>
				<span class="span70"><%=sAp2Judge%></span>
			</div>
			<div class="tourdetailline">
				<span class="span25">&nbsp;</span>
				<span class="span70"><%=sAp3Judge%></span>
			</div>
			<div class="tourdetailline">
				<span class="span25">&nbsp;</span>
				<span class="span70"><%=sAp4Judge%></span>
			</div>
			<div class="tourdetailline">
				<span class="span25">&nbsp;</span>
				<span class="span70"><%=sAp5Judge%></span>
			</div>
			<%

END SUB






' --------------------------------------
  SUB SetTabAndBottomColor (TourCount) 
' --------------------------------------
  
  	bottomcolorArray(TourCount)="#FFFFFF"
		ThisFlag=""
		FlagMessage=""  
 		IF sTStatusArray(TourCount) > 1 THEN
				ThisFlag=s_greenflag
				FlagMessage="All approvals received"
				'tabcolorArray(TourCount)="#80FF80"
				tabcolorArray(TourCount)=scolor08
				'bottomcolorArray(TourCount)="#CCFFCC"		' --- Light shade of lime green
		ELSEIF sTStatusArray(TourCount) = 1 THEN
				ThisFlag=s_yellowflag
				FlagMessage="Final sanction approvals pending"			
				'tabcolorArray(TourCount)="#FFFF4D"
				tabcolorArray(TourCount)=scolor07
				'bottomcolorArray(TourCount)="#FFFFB2"
		ELSEIF sTStatusArray(TourCount) = 3 THEN
				ThisFlag=s_redflag			 		
				FlagMessage="Tournament cancelled"			
				'tabcolorArray(TourCount)="#FF4D4D"
				tabcolorArray(TourCount)=scolor09
				'bottomcolorArray(TourCount)="#FFE6E6"	
		ELSEIF sTStatusArray(TourCount) = 0 THEN
				'tabcolorArray(TourCount)="#FF4D4D"
				'bottomcolorArray(TourCount)="#FFE6E6"
				ThisFlag=s_redflag
				tabcolorArray(TourCount)=scolor09
				FlagMessage="No sanction approvals received"			
		END IF	


END SUB









' **************************
  SUB Load_Tournament_Query
' **************************



	' -------------------------------------------------------------------------------------
	' --- Determines program branching - either listing or details on single tournament ---
	' -------------------------------------------------------------------------------------
	
	IF CInt(sTourRange)>=0 AND CInt(sTourRange)<=7 THEN InRange=true 


	' -----------------------------------------------------------------------------------------
	' --- Runs the appropriate SQL query depending on the range selected from the drop-down ---
	' -----------------------------------------------------------------------------------------




	IF sTourRange = "0" OR sTourRange = "1" OR sTourRange = "2" OR sTourRange = "3" OR sTourRange = "5" OR sTourRange = "6" THEN
			' --- Executes query for displaying tournaments ---
			PerformSQLQuery_2010
	ELSEIF sTourRange = "4" OR sTourRange = "7" THEN
			' --- Executes query for displaying tournaments ---
			PerformSQLQuery_Pre2009

	ELSE
			InRange=false
	END IF

END SUB






' -------------------------
  SUB Display_Filters_New
' -------------------------
  
 	' --- Displays if Filters div is on ---
	%>
	<div id="Filters" class="errorbox" style="width:99%; height:460px; margin:2px 0px 0px 0px; padding:0px 0px 0px 0px;">
		<form name="Filters" action="<%=ThisSitePath%>/<%=ThisFileName%>?action=Listing" method="post" style="margin:0px; padding:0px;">	
 			<input type="hidden" name="TourID" value="<%=TourID%>">
			<input type="hidden" name="pvar" value="<%=pvar%>">
			<input type="hidden" name="process" value="<%=process%>">
	 	
	 		<div style="width:96%; margin-top:10px; padding-left:10px; text-align:left; border:0px solid red;">		
				<span class="span95" style="margin-left:0px; padding-left:0px; text-align:center; font-size:16px; color:yellow; border:0px solid white;">Set Tournament Filters For Search</span> 
			</div>
	
			<div class="tourfilterdropdownline" style="text-align:left; color:white;">
				<span class="span65" style="margin-top:10px">Sport Discipline</span>
  			<span class="span30" style="padding-top:1px">Region</span>
				
  			<span class="span65" >
  				<select id="sSportsGroup" name="sSportsGroup" style="width:9em; font-size:12pt">
						<option value="aws" <%IF sSportsGroup = "aws" THEN Response.Write(" selected ") %>>AWSA</option>
						<option value="usw" <%IF sSportsGroup = "usw" THEN Response.Write(" selected ") %>>Wakeboard</option>
						<option value="aka" <%IF sSportsGroup = "aka" THEN Response.Write(" selected ") %>>Kneeboard</option>
						<option value="hyd" <%IF sSportsGroup = "hyd" THEN Response.Write(" selected ") %>>Hydrofoil</option>
    			</select>
    		</span>
    		
				<span class="span30">
    			<select id="Region" name="Region" style="width:5em; font-size:12pt">
  					<option value="" <% IF sTourRegion = "" THEN Response.Write(" selected ") %>>All</option>
  					<%
						IF sTourLevel<>"collegiate" THEN 
								%>
								<option value="C" <% IF sTourRegion = "C" THEN Response.Write(" selected ") %>>SC</option>
								<option value="M" <% IF sTourRegion = "M" THEN Response.Write(" selected ") %>>MW</option>
								<option value="W" <% IF sTourRegion = "W" THEN Response.Write(" selected ") %>>WE</option>
								<option value="S" <% IF sTourRegion = "S" THEN Response.Write(" selected ") %>>SO</option>
								<option value="E" <% IF sTourRegion = "E" THEN Response.Write(" selected ") %>>EA</option>
								<%
						END IF  
							%>
    			</select>
				</span>
		</div>		

	
		<div class="tourfilterdropdownline"  style="text-align:left; color:white;">
				<span class="span65" style="margin-top:10px">Type</span>
  			<span class="span30" style="margin-top:10px">State</span>

  			<span class="span65" >
					<select id="sTourLevel" name="sTourLevel" style="width:9em; font-size:12pt">
						<option value="premier" <%IF sTourLevel = "premier" THEN Response.Write(" selected ")%>>AWSA Premier</option>
						<option value="grass" <%IF sTourLevel = "grass" THEN Response.Write(" selected ")%>>GrassRoots</option>
						<option value="collegiate" <%IF sTourLevel = "collegiate" THEN Response.Write(" selected ")%>>Collegiate</option>
						<option value="cash" <%IF sTourLevel = "cash" THEN Response.Write(" selected ")%>>Cash Prize</option>
						<option value="clinic" <%IF sTourLevel = "clinic" THEN Response.Write(" selected ")%>>Clinics</option>
    			</select>
				</span>

				<span class="span30">
					<%
					StateArray = Split(USStatesList3,",")  %>
    			<select id="State" name="State" style="width:5em; font-size:12pt" ><%
      		FOR kvar = 0 TO UBOUND(StateArray)
        			IF TRIM(sTourState) = TRIM(StateArray(kvar)) THEN
	  							response.write("<option value = """&sTourState&""" SELECTED>"&sTourState&"</option>")
        			ELSE
	  							response.write("<option value = """&StateArray(kvar)&""">"&StateArray(kvar)&"</option>")
        			END IF
      		NEXT  
      		
      		%>
    			</select>
				</span>	
		</div>			
		
		<div class="tourfilterdropdownline"  style="text-align:left; color:white;">
  			<span class="span65" style="margin-top:10px">Range</span>
  			<span class="span30" style="margin-top:10px">Class</span>
				<br>				
  			<span class="span65">
					<select id='sTourRange' name='sTourRange' style="width:9em; font-size:12pt">
  					<option value="0"<%IF sTourRange = "0" THEN Response.Write(" selected ")%>>Custom</option>
  					<option value="1"<%IF sTourRange = "1" THEN Response.Write(" selected ")%>>Future</option>
  					<%

        		set rsSelectFields=Server.CreateObject("ADODB.recordset")
						rsSelectFields.open "Select Top 3 * FROM "&SkiYearTableName&" WHERE SkiYearID<>'1' ORDER BY BeginDate DESC", SConnectionToTRATable

						IF NOT rsSelectFields.eof THEN %>
	  						<option value="2"<%IF sTourRange = "2" THEN Response.Write(" selected ")%>>Ski Year <%=right(right(TRIM(rsSelectFields("SkiYearName")),4),4)%></option><%
	  						rsSelectFields.movenext 
	  						IF NOT rsSelectFields.eof THEN %>
										<option value="3"<%IF sTourRange = "3" THEN Response.Write(" selected ")%>>Ski Year <%=right(right(TRIM(rsSelectFields("SkiYearName")),4),4)%></option><%
										rsSelectFields.movenext 
										IF NOT rsSelectFields.eof THEN %>
		   									<option value="4"<%IF sTourRange = "4" THEN Response.Write(" selected ")%>>Ski Year <%=right(right(TRIM(rsSelectFields("SkiYearName")),4),4)%></option>
	  	   								<option value="5"<%IF sTourRange = "5" THEN Response.Write(" selected ")%>>Calendar <%=Year(Date())%></option><%
										END IF 
	  						END IF	
		
						ELSE  ' --- Applies only if no SkiYears are found in Ski Year table  %>
								<option value="2"<%IF sTourRange = "2" THEN Response.Write(" selected ")%>>Ski Year <%=Year(Date())%></option>
								<option value="3"<%IF sTourRange = "3" THEN Response.Write(" selected ")%>>Ski Year <%=Year(Date())-1%></option>
								<option value="4"<%IF sTourRange = "4" THEN Response.Write(" selected ")%>>Ski Year <%=Year(Date())-2%></option>
	  						<option value="5"<%IF sTourRange = "5" THEN Response.Write(" selected ")%>>Calendar <%=Year(Date())%></option><%
						END IF
						rsSelectFields.close

						IF adminmenulevel>19 THEN %>
	  						<option value="6"<%IF sTourRange = "6" THEN Response.Write(" selected ")%>>Calendar <%=Year(Date())-1%></option>
	  						<option value="7"<%IF sTourRange = "7" THEN Response.Write(" selected ")%>>Calendar <%=Year(Date())-2%></option><%
						END IF 
						%>	
    			</select>
				</span>

				<span class="span30">
      		<select id="sClass" name="sClass" style="width:5em; font-size:12pt">
            <option value="All" <%IF sClass = "All" THEN Response.Write(" selected ")%>>All</option>
            <option value="R" <%IF sClass = "R" THEN Response.Write(" selected ")%>>R</option>
            <option value="L" <%IF sClass = "L" THEN Response.Write(" selected ")%>>L</option>
            <option value="E" <%IF sClass = "E" THEN Response.Write(" selected ")%>>E</option>
            <option value="C" <%IF sClass = "C" THEN Response.Write(" selected ")%>>C</option>
            <option value="N" <%IF sClass = "N" THEN Response.Write(" selected ")%>>N</option>
            <option value="F" <%IF sClass = "F" THEN Response.Write(" selected ")%>>F</option>
            <option value="F_O" <%IF sClass = "F_O" THEN Response.Write(" selected ")%>>F W/O</option>
    			</select>
  			</span>
			
		</div>			


		<div class="tourfilterdropdownline"  style="text-align:left; color:white;">
  			<span class="span65" style="margin-top:10px">Start Month</span>
  			<span class="span30" style="margin-top:10px">End Month</span>
  			
  			<span class="span65" >	
  				<%
    			LoadMonthsPulldown_New "StartMonth", StartMonth, "width:5em; font-size:12pt" 
    			%>
  			</span>
  			<span class="span30" >	
  				<%
    			LoadMonthsPulldown_New "EndMonth", EndMonth, "width:5em; font-size:12pt" 
    			%>
  			</span>
		</div>
					
		<div style="border:0px solid white; margin:40px 0px 0px 0px; padding:0px 0px 0px 0px; border:0px solid white;">
				<span class="span45" style="text-align:left; margin:0px 0px 0px 0px; padding:0px 0px 0px 0px; border:0px solid white;">
					<input type=button class="buttonblue" style="width:8.5em;" value="Save Settings" onclick="javascript:StoreTournamentSettingsToLocalVar();" title="Store These Settings">
				</span>
				<span class="span45" style="text-align:center;">
					<input type=button class="buttonblue" style="width:8.5em;" value="Recall Settings" onclick="javascript:UpdateTournamentSettingsFromLocal();" title="Restore My Previously Saved Tournament Filters">
				</span>
		</div>
		<div style="text-align:center; color:white;">
				<span class="span100" style="margin-top:65px;">
					<input class="buttonblue" type="submit" style="width:180px; font-size:12pt;" name="thisaction" value="Display Listing">
  			</span>
		</div>
	</form>

</div> <!- Filters id /-->
<% 
		
END SUB







' *************************
  SUB Read_Form_Variables
' *************************  

Action = TRIM(LCASE(Request("Action")))

TourID=TRIM(Request("TourID"))

sTourSportGroup=Request("sTourSportGroup")
sTourRange = TRIM(Request("sTourRange"))
IF sTourRange="" THEN sTourRange="1"
sTourLevel = TRIM(Request("sTourLevel"))

' --- If resulting from link on tournament listing --
IF Request("rg")<>"" THEN sTourRange=Request("rg")

sTourState = TRIM(Request("State"))
sTourDate1 = TRIM(Request("Tour_Date1"))
sTourDate2 = TRIM(Request("Tour_Date2"))
sTourRegion = TRIM(Request("Region"))

sClass=TRIM(Request("sClass"))
StartMonth = TRIM(Request("StartMonth"))	
EndMonth = TRIM(Request("EndMonth"))
IF TRIM(StartMonth)="" THEN StartMonth=0
IF TRIM(EndMonth)="" THEN EndMonth=0
	
process=TRIM(Request("process"))

pvar=Request("pvar")
thisaction=TRIM(Request("thisaction"))
IF thisaction="Update Search" THEN pvar=""
OpenNewForm=TRIM(Request("OpenNewForm"))

sShowSQL = Request("sShowSQL")
adminmenulevel=Session("adminmenulevel")

IF TRIM(Request("SkiYear")) <> "" THEN Session("SkiYear") = TRIM(Request("SkiYear"))

sSportsGroup=LCASE(Request("sSportsGroup"))
SELECT CASE sSportsGroup
		CASE "aws"
				sl="on"
				tr="on"
				ju="on"
		CASE "aka"
				kb="on"
		CASE "usw"
				wb="on"
				ws="on"
				wu="on"
		CASE "hyd"
				hy="on"
		CASE "abc"
				bf="on"		
		
		'da="on"
		'jd="on"
		'ad="on"

END SELECT			

END SUB  






%>





