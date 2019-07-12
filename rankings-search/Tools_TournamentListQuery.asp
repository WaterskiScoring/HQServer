<%



' ---------------------------------------------------------------------------------------------------------
    SUB PerformSQLQuery_2010  ' ----------------  BUILD SQL statement   -----------------------------------
' ---------------------------------------------------------------------------------------------------------

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' 			--- IMPORTANT ---

' --- This modules is only applicable to tournaments in 2010 and after

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++





' --- From Jim Meis 10-29-2007
' Class I (and N) ARE "traditional" AWSA/NCWSA classes.  NCWSA is supposed to use Class I, and Class I is expected by WSTIMS for collegiate events.
' Class I should only be used by NCWSA so I should probably remove it as an option on the AWSA sanction form.
' Classes I and N predate Grassroots, have different officials requirements, allow different officials work credits, and have different sanction fees.
' Classes I and N also require AWSA or NCWSA Region Sanction Approval which Grassroots technically does not.   
' If it has a traditional event, the Sports Division admin and HQ give approval as part of the traditional approval. 
'     If all the traditional requirements are in place any Grassroots program is automatically OK without much thought

' TEventSlalom, TEventTrick, and TEventJump -   Barefoot traditional tournaments (ABC) as users of those fields.  
' Users are:  AWS, NCW, ABC use all 3 and AKA uses TEventSlalom and TEventTrick

' THSClassF is the original Fun field.  It is and always was distinct from THSClassN.   Officials required are different etc.
' Current designation for fun 3 event is TEventF3ev=1.  More specifically this indicates a Grassroots event that the sponsor is characterizing
'    as 3 event type.  Could be sanctioned under any of the sports divisions.
' All the fields beginning with TEventF, except TEventFun, are the most recent Grassroots fields.
' Sponsor can offer multiple classes and skier can pick what level he wants to ski at - so THSClassR is 1 if R is offered and 0 if it is not.  


' --- From Jim Meis 10/28/2007
' In response to the question "why so many fields?"
' They came about because of the umpteen revisions to the "Fun", NWL, NBL, NSL, and Grassroots programs in the past 3 or 4 years.
' Started out with THSClassF, then added TEventFun to allow stand alone fun, then dropped THSClassF to separate FUN from 3 event 
'   and to allow fun to be sanctioned by other sports divisions, then added NSL, NWL, NBL, then added Grassroots, 
'   then dropped NSL, NWL, and NBL from sanction form but left it on the adverts whenever grassroots was selected 
'   for 3ev, Barefoot or wakeboard).  Latest directive says drop NSL, NWL, and NBL entirely, change the fun "events" 
'   already offered and add new ones


' --- From Jim Meis 3/1/2007 ---
' There are separate description fields in swift for traditional, fun, clinics, Wakeboard and Kneeboard events.  
' They have zero length strings if there is no description  (no matching events).
' When a sanction includes more than one of these categories you need to concatenate the description fields to get all the information.


' Tschedul.TDescription - AWS, ABC, or NCW standard events
' Tschedul.WDescription - Wakeboard standard events
' Tschedul.KDescription - Kneeboard standard events
' Tschedul.FDescription - Fun Events including NSL 
' Tschedul.CDescription - Clinics

' Tschedul.TStatus = 0     Application received
' Tschedul.TStatus = 1     Region approved
' Tschedul.TStatus = 2    USAWS Approved
 
' Tschedul.TPending   True until first save by an administrator - To publish must be TPending must be false and the other conditions below 
'     must be met.
' RegnSetup.ShowPSchedule - On Off switch for the entire schedule for a Region.
' RegnSetup.ShowGBLink  -  controls display of tournament schedule as pick list for sponsors. Generally set same as ShowPSchedule
' RegnSetup.GBPolicy = true if ad is allowed to be displayed before full Region approval of the sanction.  Necessary but not sufficient.
' Tschedul.TKitOKGuidebookAd  -   Set by Region Admin on each sanction application - Gives Region approval of content of the Ad.  Allows 
'     publication if GBPolicy is true and ShowPSchedule is true. True = Bit 1 False = Bit 0
 
' Tschedule.TPending is true by default - it is changed to false after the first review and save by an administrator.  Nothing should be 
'    displayed unless TPending is false (has received its first review and save).

' Some regions particularly Western Region do not want tournament information posted at all until the Guidebook is published.  Regions can 
'    toggle ShowPSchedule on and off in their Region Preferences. True = OK to show as long as the rest of the conditions are met.  
'    False means do not show under any conditions.
 
' ShowGBLink is related but only important for SWIFT - it determines if the tournament schedule is used as a pick list for sponsors revising 
'    tournaments or if they have to supply their tournAppID and Edit Code blind.
 
' GBPolicy -    Guidebook Policy determines if advertisement is allowed to display before the Region has given their sanction approval.
' Some regions require that the region part of the sanction process be complete (fees paid) and approved before displaying the advertisement.  
'    If guidebook = false then don't display unless TStatus >= 1
 
' Other Regions Others don't care and allow publication prior to region approval.   They only require that the ad itself be approved.
' In this case TStatus could be 0 or higher, Guidebook must be true, and TKitOKGuidebookAd must be true (ad itself has region approval)

' The  ShowReg, ShowAppointed, etc control display of specific parts of an ad - also set in region preferences. Some regions don't want 
'    registrar information published online until the guidebook is published on the theory that it levels the playing field for entries.
' ---------------------------------------------------------------------------------------------------------------------------------------



sSQL = "SELECT TOP 800 "
sSQL = sSQL + "ST.TournAppID, TName, ST.SptsGrpID, TDescription, WDescription, ST.FDescription, KDescription, CDescription" 
sSQL = sSQL + ", TSanction, TSanType, TDateE, TDateS, TCity, Tstate, Pending, Deleted"
sSQL = sSQL + ", ShowPSched, TKitOKGuideBookAd, GBPolicy, TStatus, ShowRegistrar"
sSQL = sSQL + ", OK2Publish"

sSQL = sSQL + ", Gr1AWSPulls, Gr1ABCPulls, Gr1USWPulls, Gr1AKAPulls, Gr1USHPulls, Gr1WSDPulls"
sSQL = sSQL + ", Gr2USH_FreeRidePulls, Gr2USH_JumpOutPulls, Gr2USH_BigAirPulls, Gr2USH_3TrickPulls"
sSQL = sSQL + ", Gr2AWS_SPulls, Gr2AWS_TPulls, Gr2ABC_SPulls, Gr2ABC_TPulls, Gr2USW_WPulls, Gr2USW_SkatePulls, Gr2USW_SurfPulls, Gr2USW_RailJamPulls" 
sSQL = sSQL + ", Gr2AKA_SPulls, Gr2AKA_TPulls, Gr2AKA_FreePulls, Gr2AKA_FlipPulls"
sSQL = sSQL + ", OLRDisplayStatus, UseOLReg, OLR_PD"

sSQL = sSQL + ", TRS.PayPalAct, TRS.PayPalOK"

sSQL = sSQL + ", ST.SptsGrpID AS sSptsGrpID, ST.TRegion AS STRegion"
sSQL = sSQL + ", ST.TEventNWL, ST.TEventNBL, ST.TEventNSL, ST.THSClassN, ST.THTClassN, ST.THJClassN"
sSQL = sSQL + ", TEventF3ev"
sSQL = sSQL + ", ST.TEventWake, ST.TEventWSkate, ST.TEventWSurf, ST.TEventFW"
sSQL = sSQL + ", WWakeW, WSkateW, WSurfW"
sSQL = sSQL + ", TRS.sClassC, TRS.sClassE, TRS.sClassL, TRS.sClassR, TRS.sClassCash, TRS.sClassX"
sSQL = sSQL + ", TRS.tClassC, TRS.tClassE, TRS.tClassL, TRS.tClassR, TRS.tClassCash, TRS.tClassX"
sSQL = sSQL + ", TRS.jClassC, TRS.jClassE, TRS.jClassL, TRS.jClassR, TRS.jClassCash, TRS.jClassX"
sSQL = sSQL + ", USClassC, UTClassC, UJClassC"

' --- Fields obsolete beginning in 2010
sSQL = sSQL + ", ST.TEventSlalom, ST.TEventTrick, ST.TEventJump"
'sSQL = sSQL + ", ST.TEventFun"



sSQL = sSQL + ", ST.THSClassI, ST.THJClassI, ST.THTClassI"
sSQL = sSQL + ", ST.JDClin, ST.ADClin, ST.TEventFHF, ST.TEventFKB"

sSQL = sSQL + " FROM " &SanctionTableName&" AS ST"

sSQL = sSQL + " LEFT JOIN "&RegnSetupTableName&" AS RT ON ST.SptsGrpID = RT.SptsGrpID AND ST.TRegion = RT.TRegion"
sSQL = sSQL + " LEFT JOIN "&TRegSetupTableName&" AS TRS ON TRS.TournAppID = ST.TournAppID"


'response.write("<br>sTourLevel="&sTourLevel)
'response.write("<br>sTourRange="&sTourRange)
'response.write("<br>sl="&sl)

	sSQL = sSQL + " WHERE (11=12 "   ' --- This is the top of the bracket of all event inclusions ---

	IF sTourLevel="cash" THEN
		sSQL = sSQL +" OR (ST.THSClassCash<>0 OR ST.THTClassCash<>0 OR ST.THJClassCash<>0)"
	END IF

	IF sTourLevel="premier" OR sTourLevel="all" THEN

			' --- 3 Event Premier ---		
			IF sl="on" OR tr="on" OR ju="on" THEN 		

					' --- Top of AWS bracket "OR" ---
					' -----------------------------
					sSQL = sSQL + " OR (ST.SptsGrpID='AWS' AND (3=4" 	' --- Top of AWS stuff

					' ---  ST.TEventSlalom etc maintained to allow fall 2009 tournaments to display in 2010 criteria
					IF sl="on" THEN 
							' --- Changed 1-23-2010
							sSQL = sSQL + " OR (TRS.sClassC + TRS.sClassE + TRS.sClassL + TRS.sClassR + TRS.sClassCash + TRS.sClassX)>0  OR ST.TEventSlalom<>0"
					END IF
					IF tr="on" THEN 
							' --- Changed 1-23-2010
							sSQL = sSQL + " OR (TRS.tClassC + TRS.tClassE + TRS.tClassL + TRS.tClassR + TRS.tClassCash + TRS.tClassX)>0  OR ST.TEventTrick<>0"
					END IF
					IF ju="on" THEN
							' --- Changed 1-23-2010
							sSQL = sSQL + " OR (TRS.jClassC + TRS.jClassE + TRS.jClassL + TRS.jClassR + TRS.jClassCash + TRS.jClassX)>0  OR ST.TEventJump<>0"
					END IF
					sSQL = sSQL + "))"				' --- Bottom of AWS stuff ---
			END IF		

		' --- Wakeboard Premier ---
			IF wb="on" OR ws="on" OR wu="on" THEN 
					sSQL = sSQL + " OR (1=2"
					' --- Changed 1-23-2010 ---
					IF wb="on" THEN sSQL = sSQL + " OR ST.TEventWake<>0 OR WWakeW<>0"
					IF ws="on" THEN sSQL = sSQL + " OR ST.TEventWSkate<>0 OR WSkateW<>0"
					IF wu="on" THEN sSQL = sSQL + " OR ST.TEventWSurf<>0 OR WSurfW<>0"

					sSQL = sSQL + ")" 	
			END IF	
	END IF

	IF sTourLevel="grass" OR sTourLevel="all" THEN
		
			sSQL = sSQL +" OR (5=6"  	'---- Open bracket Grassroots

			' --- Grassroots 3 Event ---
			IF sl="on" THEN 
					sSQL = sSQL + " OR ST.Gr2AWS_SPulls<>0 OR Gr1AWSPulls<>0" 
					sSQL = sSQL + " OR ST.THSClassN<>0"
					' --- TEventFun and THSClassF included for legacy Pre-2009 system ---
					sSQL = sSQL + " OR ST.THSClassF<>0 OR ST.TEventF3ev<>0"
			END IF
			IF tr="on" THEN 		
					sSQL = sSQL + " OR ST.Gr2AWS_TPulls<>0" 
					sSQL = sSQL + " OR ST.THTClassN<>0"
					' --- TEventFun and THTClassF included for legacy Pre-2009 system ---
					sSQL = sSQL + " OR ST.THTClassF<>0"
			END IF
			IF ju="on" THEN 		
					sSQL = sSQL + " OR ST.THJClassN<>0"
			END IF


			' --- Grassroots Wakeboard ---		
			IF wb="on" OR ws="on" OR wu="on" THEN 
					' --- Legacy from Pre-2009 system ---
					sSQL = sSQL + " OR (ST.TEventFW<>0 OR ST.TEventNWL<>0"

					IF wb="on" THEN 
							' --- Changed 1-23-2010 ---
							sSQL = sSQL + " OR Gr2USW_WPulls<>0 OR Gr2USW_RailJamPulls<>0 OR Gr1USWPulls<>0 OR WWakeW>0" 
					END IF
					IF ws="on" THEN 
							' --- Changed 1-23-2010 ---
							sSQL = sSQL + " OR Gr2USW_SkatePulls<>0 OR WSkateW>0"
					END IF 
					IF wu="on" THEN 
							' --- Changed 1-23-2010 ---
							sSQL = sSQL + " OR Gr2USW_SurfPulls<>0 OR WSurfW>0"
					END IF 
					sSQL = sSQL + ")" 	
		END IF

		sSQL = sSQL + ")" 		'---- Close bracket Grass

	END IF



	' --- Collegiate ---
	IF sTourLevel="collegiate" THEN 
			' --- 3 Event ---
			IF sl="on" OR tr="on" OR ju="on" THEN 		
					sSQL = sSQL + " OR (ST.SptsGrpID='NCW' AND (1=2 "
					' ---  ST.TEventSlalom etc maintained to allow fall 2009 tournaments to display in 2010 criteria
					IF sl="on" THEN sSQL = sSQL + " OR (TRS.USClassC>0 OR ST.TEventSlalom<>0)"
					IF tr="on" THEN sSQL = sSQL + " OR (TRS.UTClassC>0 OR ST.TEventTrick<>0)"
					IF ju="on" THEN sSQL = sSQL + " OR (TRS.UJClassC>0 OR ST.TEventTrick<>0)"
					sSQL = sSQL + "))" 
			END IF

			' --- Wakeboard ---		
			IF wb="on" OR ws="on" OR wu="on" THEN 
					sSQL = sSQL + " OR (ST.SptsGrpID='NCW' AND ST.TEventFW<>0 OR ST.TEventNWL<>0 OR WWakeW>0  OR WSurfW>0 OR WSkateW>0)" 
			END IF
	END IF
	
	' --- Barefoot ---
	IF bf="on" THEN sSQL = sSQL + " OR (ST.SptsGrpID='ABC' OR ST.TEventNBL<>0 OR Gr1ABCPulls<>0 OR Gr2ABC_SPulls<>0 OR Gr2ABC_TPulls<>0)"

	' --- Kneeboard ---
	IF kb="on" THEN sSQL = sSQL + " OR (ST.SptsGrpID='AKA' OR ST.TEventFKB<>0) OR Gr1AKAPulls<>0 OR Gr2AKA_SPulls<>0 OR Gr2AKA_TPulls<>0 OR Gr2AKA_FreePulls<>0 OR Gr2AKA_FlipPulls<>0"

	' --- Hydrofoil ---
	IF hy="on" THEN 
		' --- Legacy from Pre-2009 ---
		sSQL = sSQL + " OR (ST.TEventFHF<>0"
		sSQL = sSQL + " OR ST.Gr2USH_FreeRidePulls<>0 OR ST.Gr2USH_JumpOutPulls<>0 OR ST.Gr2USH_BigAirPulls<>0 OR ST.Gr2USH_3TrickPulls<>0 OR ST.Gr1USHPulls<>0)"
	END IF

	' --- Clinic ---
	IF ad="on" THEN sSQL = sSQL + " OR ST.ADClin<>0"
	IF jd="on" THEN sSQL = sSQL + " OR ST.JDClin<>0"

	sSQL = sSQL + ")"    ' --- This is the bottom of the bracket of all event inclusions ---


	


	' --- Filters for highest homologation class ---
	HighClass=99
	IF sClass="R" AND (sl="on" OR tr="on" OR ju="on") THEN 
			sSQL = sSQL + " AND (THSClassR<>0 OR THTClassR<>0 OR THJClassR<>0 OR SClassR>0 OR TClassR>0 OR JClassR>0)" 
			HighClass=1

	ELSEIF sClass="L" AND HighClass>1 AND (sl="on" OR tr="on" OR ju="on") THEN
			sSQL = sSQL + " AND (THSClassL<>0 OR THTClassL<>0 OR THJClassL<>0 OR SClassL>0 OR TClassL>0 OR JClassL>0)" 
			HighClass=2

	ELSEIF sClass="E" AND HighClass>2 AND (sl="on" OR tr="on" OR ju="on") THEN
			sSQL = sSQL + " AND (THSClassE<>0 OR THTClassE<>0 OR THJClassE<>0 OR SClassE>0 OR TClassE>0 OR JClassE>0)" 
			HighClass=3

	ELSEIF sClass="C" AND HighClass>3 AND (sl="on" OR tr="on" OR ju="on") THEN
			sSQL = sSQL + " AND (THSClassC<>0 OR THTClassC<>0 OR THJClassC<>0 OR SClassC>0 OR TClassC>0 OR JClassC>0)" 
			HighClass=4

	ELSEIF sClass="N" AND HighClass>4 AND (sl="on" OR tr="on" OR ju="on") THEN
			sSQL = sSQL + " AND (THSClassN<>0 OR THTClassN<>0 OR THJClassN<>0)" 
			HighClass=5

	ELSEIF sClass="F" AND HighClass>5 AND (sl="on" OR tr="on" OR ju="on") THEN
			sSQL = sSQL + " AND (THSClassF<>0 OR THTClassF<>0 OR THJClassF<>0)" 
			HighClass=6

	ELSEIF sClass="F_O" AND (sl="on" OR tr="on" OR ju="on") THEN
			sSQL = sSQL + " AND TEventF3ev<>0" 
			HighClass=7

	END IF



		IF sTourRange <> "" AND sTourRange <> "0" THEN
				IF sTourRange = "1" THEN
						sSQL = sSQL + " AND (ST.TDateE >= '" & Date() & "')"
				END IF

		' --- Ski Year defined as Latest in DivisionTable ---
		IF sTourRange = "2" THEN
				set rsSelectFields=Server.CreateObject("ADODB.recordset")
				rsSelectFields.open "Select Top 1 * FROM "&SkiYearTableName&" WHERE SkiYearID<>'1' ORDER BY BeginDate DESC", SConnectionToTRATable
				IF NOT rsSelectFields.eof THEN
						sSQL = sSQL + " AND (left(ST.TournAppID,2) = '" & right(right(TRIM(rsSelectFields("SkiYearName")),4),2) & "')"
				END IF
				rsSelectFields.close

		' --- Ski Year defined as SECOND latest in DivisionTable ---
		ELSEIF sTourRange = "3" THEN 
				set rsSelectFields=Server.CreateObject("ADODB.recordset")
				rsSelectFields.open "Select Top 2 * FROM "&SkiYearTableName&" WHERE SkiYearID<>'1' ORDER BY BeginDate DESC", SConnectionToTRATable
				IF NOT rsSelectFields.eof THEN
						rsSelectFields.movenext
						sSQL = sSQL + " AND (left(ST.TournAppID,2) = '" & right(right(TRIM(rsSelectFields("SkiYearName")),4),2) & "')"
				END IF
				rsSelectFields.close

		' --- Ski Year defined as THIRD latest in DivisionTable ---
		ELSEIF sTourRange = "4" THEN 
				set rsSelectFields=Server.CreateObject("ADODB.recordset")
				rsSelectFields.open "Select Top 3 * FROM "&SkiYearTableName&" WHERE SkiYearID<>'1' ORDER BY BeginDate DESC", SConnectionToTRATable
				IF NOT rsSelectFields.eof THEN
					  rsSelectFields.movenext
				  	IF NOT rsSelectFields.eof THEN
								rsSelectFields.movenext
								IF NOT rsSelectFields.eof THEN
				  					sSQL = sSQL + " AND (left(ST.TournAppID,2) = '" & right(right(TRIM(rsSelectFields("SkiYearName")),4),2) & "')"
			  				END IF
			  		END IF
				END IF
				rsSelectFields.close

		' --- Current Calendar year if the year is nearly over otherwise last calendar year ---
		ELSEIF sTourRange = "5" THEN
'				sSQL = sSQL + " AND Year(TDateS) = '"&Year(Date())&"'"
	
		' --- Last Calendar year if this year is nearly over otherwise two calendar years ago ---
		ELSEIF sTourRange = "6" THEN
				sSQL = sSQL + " AND Year(TDateS) = '"&Year(Date())-1&"'"
	
		' --- Two calendar years ago if this year is nearly over otherwise three calendar years ago ---
		ELSEIF sTourRange = "7" THEN
				sSQL = sSQL + " AND Year(TDateS) = '"&Year(Date())-2&"'"
	
		END IF
	END IF

	IF StartMonth<>0 THEN
		sSQL = sSQL + " AND Month(TDateS) >= '"&StartMonth&"'"
	END IF

	IF EndMonth<>0 THEN
		sSQL = sSQL + " AND Month(TDateE) <= '"&EndMonth&"'"
	END IF

	IF sTourState <> "" AND LCASE(sTourState) <> "all" THEN sSQL = sSQL + " AND lower(TState) = '" & sqlclean(lcase(sTourState)) & "'"

	IF sTourRegion <> "" THEN sSQL = sSQL + " AND lower(right(left(ST.TournAppID,3),1)) = '" & sqlclean(lcase(sTourRegion)) & "'"

	IF sTourDate1 <> "" THEN sSQL = sSQL + " AND (TDateE >= '" & sTourDate1 & "' or TDateS >= '" & sTourDate1 & "')"

	IF sTourDate2 <> "" THEN sSQL = sSQL + " AND (TDateE <= '" & sTourDate2 & "' or TDateS <= '" & sTourDate2 & "')"


	IF process="register" OR process="viewreg" OR process="admcode" THEN sSQL = sSQL + " AND PayPalOK=1 AND PayPalAct<>''"
	
  sSQL = sSQL + " ORDER BY TDateS"

	
	'response.write("<br> 2010"&sSQL)	
	IF Session("adminmenulevel")>=50 AND TRIM(sShowSQL)<>"" THEN
			response.write("<br>"&sSQL)
			'response.end
	END IF

	set rs=Server.CreateObject("ADODB.recordset")
	rs.open sSQL, SConnectionToTRATable

	IF Session("adminmenulevel")>=50 AND TRIM(sShowSQL)<>"" THEN
			IF NOT rs.eof THEN response.write("<br><br>FOUND")
	END IF

END SUB







' ---------------------------------------------------------------------------------------------------------
    SUB PerformSQLQuery_Pre2009  ' ---------------------  BUILD SQL statement   -----------------------------------
' ---------------------------------------------------------------------------------------------------------

'IF Session("AdminMenuLevel")>=50 THEN
'	response.write("<br> Mark's TEMP stop")
'	response.write("<br> sTourRange="&sTourRange)

'	response.end

'END IF

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' 			*** IMPORTANT ***

' --- This module is only applicable for tournaments prior to 2010 

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++



' --- From Jim Meis 10-29-2007
' Class I (and N) ARE "traditional" AWSA/NCWSA classes.  NCWSA is supposed to use Class I, and Class I is expected by WSTIMS for collegiate events.
' Class I should only be used by NCWSA so I should probably remove it as an option on the AWSA sanction form.
' Classes I and N predate Grassroots, have different officials requirements, allow different officials work credits, and have different sanction fees.
' Classes I and N also require AWSA or NCWSA Region Sanction Approval which Grassroots technically does not.   
' If it has a traditional event, the Sports Division admin and HQ give approval as part of the traditional approval. 
'     If all the traditional requirements are in place any Grassroots program is automatically OK without much thought

' TEventSlalom, TEventTrick, and TEventJump -   Barefoot traditional tournaments (ABC) as users of those fields.  
' Users are:  AWS, NCW, ABC use all 3 and AKA uses TEventSlalom and TEventTrick

' THSClassF is the original Fun field.  It is and always was distinct from THSClassN.   Officials required are different etc.
' Current designation for fun 3 event is TEventF3ev=1.  More specifically this indicates a Grassroots event that the sponsor is characterizing
'    as 3 event type.  Could be sanctioned under any of the sports divisions.
' All the fields beginning with TEventF, except TEventFun, are the most recent Grassroots fields.
' Sponsor can offer multiple classes and skier can pick what level he wants to ski at - so THSClassR is 1 if R is offered and 0 if it is not.  


' --- From Jim Meis 10/28/2007
' In response to the question "why so many fields?"
' They came about because of the umpteen revisions to the "Fun", NWL, NBL, NSL, and Grassroots programs in the past 3 or 4 years.
' Started out with THSClassF, then added TEventFun to allow stand alone fun, then dropped THSClassF to separate FUN from 3 event 
'   and to allow fun to be sanctioned by other sports divisions, then added NSL, NWL, NBL, then added Grassroots, 
'   then dropped NSL, NWL, and NBL from sanction form but left it on the adverts whenever grassroots was selected 
'   for 3ev, Barefoot or wakeboard).  Latest directive says drop NSL, NWL, and NBL entirely, change the fun "events" 
'   already offered and add new ones


' --- From Jim Meis 3/1/2007 ---
' There are separate description fields in swift for traditional, fun, clinics, Wakeboard and Kneeboard events.  
' They have zero length strings if there is no description  (no matching events).
' When a sanction includes more than one of these categories you need to concatenate the description fields to get all the information.


' Tschedul.TDescription - AWS, ABC, or NCW standard events
' Tschedul.WDescription - Wakeboard standard events
' Tschedul.KDescription - Kneeboard standard events
' Tschedul.FDescription - Fun Events including NSL 
' Tschedul.CDescription - Clinics

' Tschedul.TStatus = 0     Application received
' Tschedul.TStatus = 1     Region approved
' Tschedul.TStatus = 2    USAWS Approved
 
' Tschedul.TPending   True until first save by an administrator - To publish must be TPending must be false and the other conditions below 
'     must be met.
' RegnSetup.ShowPSchedule - On Off switch for the entire schedule for a Region.
' RegnSetup.ShowGBLink  -  controls display of tournament schedule as pick list for sponsors. Generally set same as ShowPSchedule
' RegnSetup.GBPolicy = true if ad is allowed to be displayed before full Region approval of the sanction.  Necessary but not sufficient.
' Tschedul.TKitOKGuidebookAd  -   Set by Region Admin on each sanction application - Gives Region approval of content of the Ad.  Allows 
'     publication if GBPolicy is true and ShowPSchedule is true. True = Bit 1 False = Bit 0
 
' Tschedule.TPending is true by default - it is changed to false after the first review and save by an administrator.  Nothing should be 
'    displayed unless TPending is false (has received its first review and save).

' Some regions particularly Western Region do not want tournament information posted at all until the Guidebook is published.  Regions can 
'    toggle ShowPSchedule on and off in their Region Preferences. True = OK to show as long as the rest of the conditions are met.  
'    False means do not show under any conditions.
 
' ShowGBLink is related but only important for SWIFT - it determines if the tournament schedule is used as a pick list for sponsors revising 
'    tournaments or if they have to supply their tournAppID and Edit Code blind.
 
' GBPolicy -    Guidebook Policy determines if advertisement is allowed to display before the Region has given their sanction approval.
' Some regions require that the region part of the sanction process be complete (fees paid) and approved before displaying the advertisement.  
'    If guidebook = false then don't display unless TStatus >= 1
 
' Other Regions Others don't care and allow publication prior to region approval.   They only require that the ad itself be approved.
' In this case TStatus could be 0 or higher, Guidebook must be true, and TKitOKGuidebookAd must be true (ad itself has region approval)

' The  ShowReg, ShowAppointed, etc control display of specific parts of an ad - also set in region preferences. Some regions don't want 
'    registrar information published online until the guidebook is published on the theory that it levels the playing field for entries.
' ---------------------------------------------------------------------------------------------------------------------------------------



sSQL = "SELECT TOP 800 "
sSQL = sSQL + "ST.TournAppID, TName, ST.SptsGrpID, TDescription, WDescription, ST.FDescription, KDescription, CDescription" 
sSQL = sSQL + ", TSanction, TSanType, TDateE, TDateS, TCity, Tstate, Pending"
sSQL = sSQL + ", ShowPSched, TKitOKGuideBookAd, GBPolicy, TStatus, ShowRegistrar"
sSQL = sSQL + ", OK2Publish"
sSQL = sSQL + ", sClassC, sClassE, sClassL, sClassR, sClassX, sClassCash"
sSQL = sSQL + ", tClassC, tClassE, tClassL, tClassR, tClassX, tClassCash"
sSQL = sSQL + ", jClassC, jClassE, jClassL, jClassR, jClassX, jClassCash"
sSQL = sSQL + ", WWakeW, WSkateW, WSurfW"
sSQL = sSQL + ", Gr1AWSPulls, Gr1ABCPulls, Gr1USWPulls, Gr1AKAPulls, Gr1USHPulls, Gr1WSDPulls"
sSQL = sSQL + ", Gr2USH_FreeRidePulls, Gr2USH_JumpOutPulls, Gr2USH_BigAirPulls, Gr2USH_3TrickPulls"
sSQL = sSQL + ", Gr2AWS_SPulls, Gr2AWS_TPulls, Gr2ABC_SPulls, Gr2ABC_TPulls, Gr2USW_WPulls, Gr2USW_SkatePulls, Gr2USW_SurfPulls, Gr2USW_RailJamPulls" 
sSQL = sSQL + ", Gr2AKA_SPulls, Gr2AKA_TPulls, Gr2AKA_FreePulls, Gr2AKA_FlipPulls"
sSQL = sSQL + ", OLRDisplayStatus, UseOLReg, OLR_PD"


'		IF wb="on" AND (  rs("WWakeW")>0 OR rs("Gr2USW_WPulls")<>0 OR rs("Gr2USW_RailJamPulls")<>0 OR rs("Gr1USWPulls") <> 0 ) THEN sWB_Offered="Y"
'		IF ws="on" AND rs("WSkateW")>0 OR rs("Gr2USW_SkatePulls") THEN sWS_Offered="Y"
'		IF wu="on" AND rs("WSurfW")>0 OR rs("Gr2USW_SurfPulls") THEN sWU_Offered="Y"

'		IF bf="on" AND (  rs("SptsGrpID")="ABC" OR rs("Gr1ABCPulls")<>0  ) THEN sBF_Offered="Y" 
'		IF kb="on" AND (  rs("SptsGrpID")="AKA" OR rs("Gr2AKA_SPulls")<>0 OR rs("Gr2AKA_TPulls")<>0 OR rs("Gr2AKA_FreePulls")<>0 OR rs("Gr2AKA_FlipPulls")<>0 OR rs("Gr1AKAPulls")<>0  ) THEN sKB_Offered="Y"	
'		IF hy="on" AND (  rs("TEventFHF")<>0 OR rs("Gr2USH_FreeRidePulls")<>0 OR rs("Gr2USH_JumpOutPulls")<>0 OR rs("Gr2USH_BigAirPulls")<>0 OR rs("Gr2USH_3TrickPulls")<>0 OR rs("Gr1USHPulls")<>0  ) THEN sHY_Offered="Y"

'		IF da="on" AND (rs("Gr1WSDPulls")) THEN sDA_Offered="Y"
'		IF jd="on" AND (rs("JDClin")<>0) THEN sJD_Offered="Y"
'		IF ad="on" AND (rs("ADClin")<>0) THEN sAD_Offered="Y"




sSQL = sSQL + ", TRS.PayPalAct, TRS.PayPalOK"

sSQL = sSQL + ", ST.SptsGrpID AS sSptsGrpID, ST.TRegion AS STRegion"
sSQL = sSQL + ", ST.TEventNWL, ST.TEventNBL, ST.TEventNSL, ST.THSClassN, ST.THTClassN, ST.THJClassN"
sSQL = sSQL + ", TEventF3ev"
sSQL = sSQL + ", ST.TEventWake, ST.TEventWSkate, ST.TEventWSurf, ST.TEventFW"
sSQL = sSQL + ", ST.TEventSlalom, ST.TEventTrick, ST.TEventJump, ST.TEventFun"
sSQL = sSQL + ", ST.THSClassI, ST.THJClassI, ST.THTClassI"
sSQL = sSQL + ", ST.JDClin, ST.ADClin, ST.TEventFHF, ST.TEventFKB"

sSQL = sSQL + " FROM " &SanctionTableName&" AS ST"

sSQL = sSQL + " LEFT JOIN "&RegnSetupTableName&" AS RT ON ST.SptsGrpID = RT.SptsGrpID AND ST.TRegion = RT.TRegion"
sSQL = sSQL + " LEFT JOIN "&TRegSetupTableName&" AS TRS ON TRS.TournAppID = ST.TournAppID"






	sSQL = sSQL + " WHERE (1=2 "   ' --- This is the top of the bracket of all event inclusions ---

	IF sTourLevel="cash" THEN
		sSQL = sSQL +" OR (ST.THSClassCash<>0 OR ST.THTClassCash<>0 OR ST.THJClassCash<>0)"
	END IF

	IF sTourLevel="premier" OR sTourLevel="all" THEN

		' --- 3 Event Premier ---		
		IF sl="on" OR tr="on" OR ju="on" THEN 		

			' --- Top of AWS bracket "OR" ---
			' -----------------------------
			sSQL = sSQL + " OR (ST.SptsGrpID='AWS' AND (3=4" 	' --- Top of AWS stuff

			IF sl="on" THEN 
				sSQL = sSQL + " OR ST.TEventSlalom<>0"
			END IF
			IF tr="on" THEN 
				sSQL = sSQL + " OR ST.TEventTrick<>0"
			END IF
			IF ju="on" THEN
				sSQL = sSQL + " OR ST.TEventJump<>0"
			END IF
			sSQL = sSQL + "))"				' --- Bottom of AWS stuff ---
		END IF		

		' --- Wakeboard Premier ---
		IF wb="on" OR ws="on" OR wu="on" THEN 
			sSQL = sSQL + " OR (1=2"
			IF wb="on" THEN sSQL = sSQL + " OR ST.TEventWake<>0"
			IF ws="on" THEN sSQL = sSQL + " OR ST.TEventWSkate<>0"
			IF wu="on" THEN sSQL = sSQL + " OR ST.TEventWSurf<>0"
			sSQL = sSQL + ")" 	
		END IF	
	END IF

	IF sTourLevel="grass" OR sTourLevel="all" THEN
		
		sSQL = sSQL +" OR (5=6"  	'---- Open bracket Grass

		' --- Grassroots 3 Event ---
		IF sl="on" THEN 
			sSQL = sSQL + " OR ST.Gr2AWS_SPulls<>0 OR Gr1AWSPulls<>0" 
			sSQL = sSQL + " OR ST.THSClassN<>0"
			' --- TEventFun and THSClassF included for legacy Pre-2009 system ---
			sSQL = sSQL + " OR ST.TEventFun<>0 OR ST.THSClassF<>0 OR ST.TEventF3ev<>0"
		END IF
		IF tr="on" THEN 		
			sSQL = sSQL + " OR ST.Gr2AWS_TPulls<>0" 
			sSQL = sSQL + " OR ST.THTClassN<>0"
			' --- TEventFun and THTClassF included for legacy Pre-2009 system ---
			sSQL = sSQL + " OR ST.THTClassF<>0"
		END IF
		IF ju="on" THEN 		
			sSQL = sSQL + " OR ST.THJClassN<>0"
		END IF


		' --- Grassroots Wakeboard ---		
		IF wb="on" OR ws="on" OR wu="on" THEN 
			' --- Legacy from Pre-2009 system ---
			sSQL = sSQL + " OR (ST.TEventFW<>0 OR ST.TEventNWL<>0"
			IF wb="on" THEN 
				sSQL = sSQL + " OR Gr2USW_WPulls<>0 OR Gr2USW_RailJamPulls<>0 OR Gr1USWPulls<>0" 
			END IF
			IF ws="on" THEN 
				sSQL = sSQL + " OR Gr2USW_SkatePulls<>0"
			END IF 
			IF wu="on" THEN 
				sSQL = sSQL + " OR Gr2USW_SurfPulls<>0"
			END IF 
			sSQL = sSQL + ")" 	
		END IF

		sSQL = sSQL + ")" 		'---- Close bracket Grass


	END IF

	' --- Collegiate ---
	IF sTourLevel="collegiate" THEN 
		' --- 3 Event ---
		IF sl="on" OR tr="on" OR ju="on" THEN 		
			sSQL = sSQL + " OR (ST.SptsGrpID='NCW' AND (1=2 "
			IF sl="on" THEN sSQL = sSQL + " OR ST.TEventSlalom<>0 OR ST.TEventFun<>0"
			IF tr="on" THEN sSQL = sSQL + " OR ST.TEventTrick<>0"
			IF ju="on" THEN sSQL = sSQL + " OR ST.TEventJump<>0"
			sSQL = sSQL + "))" 
		END IF

		' --- Wakeboard ---		
		IF wb="on" OR ws="on" OR wu="on" THEN 
			sSQL = sSQL + " OR (ST.SptsGrpID='NCW' AND ST.TEventFW<>0 OR ST.TEventNWL<>0)" 
		END IF

	END IF


IF Session("AdminMenuLevel")>=50 THEN
	' response.write("SptsGrpID=")
END IF
	' --- Barefoot ---
	IF bf="on" THEN sSQL = sSQL + " OR (ST.SptsGrpID='ABC' OR ST.TEventNBL<>0 OR Gr1ABCPulls<>0 OR Gr2ABC_SPulls<>0 OR Gr2ABC_TPulls<>0)"

	' --- Kneeboard ---
	IF kb="on" THEN sSQL = sSQL + " OR (ST.SptsGrpID='AKA' OR ST.TEventFKB<>0) OR Gr1AKAPulls<>0 OR Gr2AKA_SPulls<>0 OR Gr2AKA_TPulls<>0 OR Gr2AKA_FreePulls<>0 OR Gr2AKA_FlipPulls<>0"


	' --- Hydrofoil ---
	IF hy="on" THEN 
		' --- Legacy from Pre-2009 ---
		sSQL = sSQL + " OR (ST.TEventFHF<>0"
		sSQL = sSQL + " OR ST.Gr2USH_FreeRidePulls<>0 OR ST.Gr2USH_JumpOutPulls<>0 OR ST.Gr2USH_BigAirPulls<>0 OR ST.Gr2USH_3TrickPulls<>0 OR ST.Gr1USHPulls<>0)"
	END IF


	' --- Clinic ---
	IF ad="on" THEN sSQL = sSQL + " OR ST.ADClin<>0"
	IF jd="on" THEN sSQL = sSQL + " OR ST.JDClin<>0"


	sSQL = sSQL + ")"    ' --- This is the bottom of the bracket of all event inclusions ---
	

	' --- Filters for highest homologation class ---


	HighClass=99
	IF sClass="R" AND (sl="on" OR tr="on" OR ju="on") THEN 
		sSQL = sSQL + " AND (THSClassR<>0 OR THTClassR<>0 OR THJClassR<>0)" 
		HighClass=1

	ELSEIF sClass="L" AND HighClass>1 AND (sl="on" OR tr="on" OR ju="on") THEN
		sSQL = sSQL + " AND (THSClassL<>0 OR THTClassL<>0 OR THJClassL<>0)" 
		HighClass=2

	ELSEIF sClass="E" AND HighClass>2 AND (sl="on" OR tr="on" OR ju="on") THEN
		sSQL = sSQL + " AND (THSClassE<>0 OR THTClassE<>0 OR THJClassE<>0)" 
		HighClass=3

	ELSEIF sClass="C" AND HighClass>3 AND (sl="on" OR tr="on" OR ju="on") THEN
		sSQL = sSQL + " AND (THSClassC<>0 OR THTClassC<>0 OR THJClassC<>0)" 
		HighClass=4

	ELSEIF sClass="N" AND HighClass>4 AND (sl="on" OR tr="on" OR ju="on") THEN
		sSQL = sSQL + " AND (THSClassN<>0 OR THTClassN<>0 OR THJClassN<>0)" 
		HighClass=5

	ELSEIF sClass="F" AND HighClass>5 AND (sl="on" OR tr="on" OR ju="on") THEN
		sSQL = sSQL + " AND (THSClassF<>0 OR THTClassF<>0 OR THJClassF<>0)" 
		HighClass=6

	ELSEIF sClass="F_O" AND (sl="on" OR tr="on" OR ju="on") THEN
		sSQL = sSQL + " AND TEventF3ev<>0" 
		HighClass=7

	END IF



	IF sTourRange <> "" AND sTourRange <> "0" THEN
        	IF sTourRange = "1" THEN
			sSQL = sSQL + " AND (ST.TDateE >= '" & Date() & "')"
		END IF

		' --- Ski Year defined as Latest in DivisionTable ---
		IF sTourRange = "2" THEN
		        set rsSelectFields=Server.CreateObject("ADODB.recordset")
            		rsSelectFields.open "Select Top 1 * FROM "&SkiYearTableName&" WHERE SkiYearID<>'1' ORDER BY BeginDate DESC", SConnectionToTRATable
			IF NOT rsSelectFields.eof THEN
				sSQL = sSQL + " AND (left(ST.TournAppID,2) = '" & right(right(TRIM(rsSelectFields("SkiYearName")),4),2) & "')"
			END IF
			rsSelectFields.close

		' --- Ski Year defined as SECOND latest in DivisionTable ---
		ELSEIF sTourRange = "3" THEN 
		        set rsSelectFields=Server.CreateObject("ADODB.recordset")
            		rsSelectFields.open "Select Top 2 * FROM "&SkiYearTableName&" WHERE SkiYearID<>'1' ORDER BY BeginDate DESC", SConnectionToTRATable
			IF NOT rsSelectFields.eof THEN
				rsSelectFields.movenext
				sSQL = sSQL + " AND (left(ST.TournAppID,2) = '" & right(right(TRIM(rsSelectFields("SkiYearName")),4),2) & "')"
			END IF
			rsSelectFields.close

		' --- Ski Year defined as THIRD latest in DivisionTable ---
		ELSEIF sTourRange = "4" THEN 
		        set rsSelectFields=Server.CreateObject("ADODB.recordset")
            		rsSelectFields.open "Select Top 3 * FROM "&SkiYearTableName&" WHERE SkiYearID<>'1' ORDER BY BeginDate DESC", SConnectionToTRATable
			IF NOT rsSelectFields.eof THEN
			  rsSelectFields.movenext
			  IF NOT rsSelectFields.eof THEN
				rsSelectFields.movenext
				IF NOT rsSelectFields.eof THEN
				  sSQL = sSQL + " AND (left(ST.TournAppID,2) = '" & right(right(TRIM(rsSelectFields("SkiYearName")),4),2) & "')"
			  	END IF
			  END IF
			END IF
			rsSelectFields.close

		' --- Current Calendar year if the year is nearly over otherwise last calendar year ---
		ELSEIF sTourRange = "5" THEN
'			IF month(date())>10 THEN 
				sSQL = sSQL + " AND Year(TDateS) = '"&Year(Date())&"'"
'			ELSE
'response.write("<br>Year(Date)="&Year(Date()))

'response.end
'				sSQL = sSQL + " AND Year(TDateS) = '"&Year(Date())-1&"'"
'			END IF

		' --- Last Calendar year if this year is nearly over otherwise two calendar years ago ---
		ELSEIF sTourRange = "6" THEN
			'IF month(date())>10 THEN 
				sSQL = sSQL + " AND Year(TDateS) = '"&Year(Date())-1&"'"
			'ELSE
			'	sSQL = sSQL + " AND Year(TDateS) = '"&Year(Date())-2&"'"
			'END IF

		' --- Two calendar years ago if this year is nearly over otherwise three calendar years ago ---
		ELSEIF sTourRange = "7" THEN
			'IF month(date())>10 THEN 
				sSQL = sSQL + " AND Year(TDateS) = '"&Year(Date())-2&"'"
			'ELSE
			'	sSQL = sSQL + " AND Year(TDateS) = '"&Year(Date())-3&"'"
			'END IF

		END IF
	END IF

	IF StartMonth<>0 THEN
		sSQL = sSQL + " AND Month(TDateS) >= '"&StartMonth&"'"
	END IF

	IF EndMonth<>0 THEN
		sSQL = sSQL + " AND Month(TDateE) <= '"&EndMonth&"'"
	END IF

	IF sTourState <> "" AND LCASE(sTourState) <> "all" THEN sSQL = sSQL + " AND lower(TState) = '" & sqlclean(lcase(sTourState)) & "'"

	IF sTourRegion <> "" THEN sSQL = sSQL + " AND lower(right(left(ST.TournAppID,3),1)) = '" & sqlclean(lcase(sTourRegion)) & "'"

	IF sTourDate1 <> "" THEN sSQL = sSQL + " AND (TDateE >= '" & sTourDate1 & "' or TDateS >= '" & sTourDate1 & "')"

	IF sTourDate2 <> "" THEN sSQL = sSQL + " AND (TDateE <= '" & sTourDate2 & "' or TDateS <= '" & sTourDate2 & "')"


	IF process="register" OR process="viewreg" OR process="admcode" THEN sSQL = sSQL + " AND PayPalOK=1"

	ShowCancelled="no"
	IF ShowCancelled = "no" THEN  sSQL = sSQL + " AND TStatus<>'3'" 
	
	sSQL = sSQL + " ORDER BY TDateS"


'	response.write("<br> 2009"&sSQL)	
	IF Session("adminmenulevel")>=50 AND TRIM(sShowSQL)<>"" THEN
			response.write("<br>2 - "&Session("adminmenulevel"))
			response.write("<br>"&sSQL)
			'	response.end
	END IF


	set rs=Server.CreateObject("ADODB.recordset")
	rs.open sSQL, SConnectionToTRATable

END SUB



%>
