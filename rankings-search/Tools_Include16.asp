<%
' ------------------------------------------------------------------------------------------------------------------
' ----- Tools_Include.asp      Tools where an include statement is put in the calling program
' ------------------------------------------------------------------------------------------------------------------



' ------------------------------------------
   SUB WhatDropdownImage (EventSelected)
' ------------------------------------------



' ---- Define background images ---
IF EventSelected="S" THEN
	DropImagePath=PathToTRA&"images\DropDown\Slalom"
	DropPartial="images\DropDown\Slalom"
ELSEIF EventSelected="T" THEN
	DropImagePath=PathToTRA&"images\DropDown\Tricks"
	DropPartial="images\DropDown\Tricks"
ELSEIF EventSelected="J" THEN
	DropImagePath=PathToTRA&"images\DropDown\Jump"
	DropPartial="images\DropDown\Jump"
ELSEIF EventSelected="WB" THEN
	DropImagePath=PathToTRA&"images\DropDown\Wakeboard"
	DropPartial="images\DropDown\Wakeboard"
ELSEIF EventSelected="WS" THEN
	DropImagePath=PathToTRA&"images\DropDown\Wakeskate"
	DropPartial="images\DropDown\Wakeskate"
ELSEIF EventSelected="WU" THEN
	DropImagePath=PathToTRA&"images\DropDown\Wakesurf"
	DropPartial="images\DropDown\Wakesurf"
ELSEIF EventSelected="KB" THEN
	DropImagePath=PathToTRA&"images\DropDown\Kneeboard"
	DropPartial="images\DropDown\Kneeboard"
ELSEIF EventSelected="BF" THEN
	DropImagePath=PathToTRA&"images\DropDown\Barefoot"
	DropPartial="images\DropDown\Barefoot"
ELSEIF EventSelected="HY" THEN
	DropImagePath=PathToTRA&"images\DropDown\Hydrofoil"
	DropPartial="images\DropDown\Hydrofoil"
ELSE
	DropImagePath=PathToTRA&"images\DropDown\Jump"
	DropPartial="images\DropDown\Jump"
END IF

'IF Session("adminmenulevel")>20 THEN
'	markdebug("DropImagePath="&DropImagePath)
'	markdebug("DropPartial="&DropPartial)
'END IF





' ------------------------------------------------------
' --- Create File System Object To Get list of files ---
' --- Get The path For the web page and its dir. -------
' --- Set the object folder To the mapped path ---------
' ------------------------------------------------------
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(DropImagePath)
set objFilesInFolder = objFolder.Files

Dim ListingCount, X
ListingCount = 0

' --- Create an integer based on time (sec) divided to segment to 60/n where 
WhichPhoto=INT(Cdbl(DatePart("s", Now, 1))/6)

' --- Assumes 10 second intervals and 3 CASE statements assume 3 images per folder ---
SELECT CASE WhichPhoto
  CASE 1,4,7
	WhichPhoto=1	
  CASE 2,5,8,0
	WhichPhoto=2
  CASE 3,6,9 
	WhichPhoto=3
END SELECT

' --- Reads all of the images from the folder defined above ---

' ----------------
' --- DAVE CLARK - Would like to have one definition to TEST existance and read into list and define for MainImage ---
' ----------------
IF objFilesInFolder.Count <> 0 THEN
  For Each objFile In objFolder.Files
	ListingCount=ListingCount+1
	IF ListingCount=WhichPhoto THEN
		MainImage=DropPartial&"\"&objFile.Name
	END IF
  Next
END IF


'IF Session("adminmenulevel")>20 THEN
'	markdebug("MainImage="&MainImage)
'END IF



END SUB



' --------------------
  SUB SetEventImage
' --------------------

' --- Runs subroutine to define image from a common location in Tools_include.asp  ---
IF wb="on" THEN
	WhatDropDownImage "WB"
ELSEIF ws="on" THEN
	WhatDropDownImage "WS"
ELSEIF WU="on" THEN
	WhatDropDownImage "WU"
ELSEIF sl="on" THEN
	WhatDropDownImage "S"
ELSEIF tr="on" THEN
	WhatDropDownImage "T"
ELSEIF ju="on" THEN
	WhatDropDownImage "J"
ELSEIF sSptsGrpID="NCW" THEN
	WhatDropDownImage "S"
ELSEIF kb="on" THEN
	WhatDropDownImage "KB"
ELSEIF bf="on" THEN
	WhatDropDownImage "BF"
ELSEIF hy="on" OR hf="on" THEN
	WhatDropDownImage "HY"
ELSE
	WhatDropDownImage "XX"
END IF

END SUB 



' -------------------------
   SUB DisplaySponsorImages_ORIG
' -------------------------

Dim wp, SponsorImagePath, SponsorPartial, WhiteImage, SponsorLogo
Dim objFSO, objFolder, objFilesInFolder, objFile
Dim ListingCount, X


wp="rank"
SponsorImagePath=""

SELECT CASE wp
	CASE "rank"
		SponsorImagePath=PathToTRA&"images\Sponsors\Rankings"
		SponsorPartial="images\Sponsors\Rankings"
END SELECT

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(SponsorImagePath)
set objFilesInFolder = objFolder.Files

ListingCount = 0

IF objFilesInFolder.Count <> 0 THEN
  For Each objFile In objFolder.Files

	SponsorLogo=SponsorPartial&"\"&objFile.Name
	WhiteImage=SponsorPartial&"\WhiteSpace.bmp"
 	%>	
	<br>&nbsp;<br><img src="<%=SponsorLogo%>" width="180"><br><img src="<%=WhiteImage%>" width="180"><%	
  Next
END IF


END SUB




' -------------------------
   SUB DisplaySponsorImages_TEMP
' -------------------------

%>
<div class="leftsidebar">
<p><a target="_blank" href="http://shop.usawaterski.org/?utm_source=usa-ws&utm_medium=site-sponsor"><img src="http://www.usawaterski.org/Logos/USA-WSShopLogo.jpg" alt="" width="220" height="186" border="0" /></a></p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p><a target="new" href="http://www.geico.com/disc/usawaterski/?utm_source=usa-ws&utm_medium=site-sponsor"><img src="http://www.usawaterski.org/Logos/GEICO.jpg" alt="" width="227" height="112" border="0" /></a></p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p><a target="new" href="http://www.visitcentralflorida.org/?utm_source=usa-ws&utm_medium=site-sponsor"><img src="http://www.usawaterski.org/Logos/VisitCentralFlorida.gif" alt="" width="192" height="49" border="0" /></a></p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p><a target="_blank" href="http://www.D3Skis.com/?utm_source=usa-ws&utm_medium=site-sponsor"><img src="http://www.usawaterski.org/logos/D3SkisBlack.jpg" alt="" width="234" height="57" /></a></p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p><a target="new" href="http://www.hyndsightvision.com/?utm_source=usa-ws&utm_medium=site-sponsor"><img src="http://www.usawaterski.org/Logos/HyndsightLogo.jpg" alt="" width="220" height="221" border="0" /></a></p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p><a target="new" href="https://www.usawaterski.org/members/login/index.asp"><img src="http://www.usawaterski.org/Logos/UnitedAirlines.gif" alt="" width="230" height="230" border="0" /></a></p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p><a target="new" href="https://www.hertz.com/?utm_source=usa-ws&utm_medium=site-sponsor"><img src="http://www.usawaterski.org/Logos/Hertz.gif" alt="" width="230" height="230" border="0" /></a></p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p><a target="_blank" href="http://www.usawaterskifoundation.org"><img src="/logos/USAWSFoundationLogo.jpg" alt="" width="220" height="139" border="0" /></a></p>
<p>&nbsp;</p>

</div>

<%
END SUB


' -------------------------
   SUB DisplaySponsorImages
' -------------------------

%>
<div class="leftsidebar">
<p><a target="_blank" href="http://www.nautique.com/?utm_source=usa-ws&utm_medium=site-sponsor"><img src="http://www.usawaterski.org/rankings/images/Sponsors/Rankings/2008 Nautiques.jpg" alt="Nautiques" width="220" border="0" /></a></p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p><a target="_blank" href="http://www.goode.com/?utm_source=usa-ws&utm_medium=site-sponsor"><img src="http://www.usawaterski.org/rankings/images/Sponsors/Rankings/8-Goode.jpg" alt="Goode Skis" width="220" border="0" /></a></p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p><a target="_blank" href="http://www.ojprops.com/?utm_source=usa-ws&utm_medium=site-sponsor"><img src="http://www.usawaterski.org/rankings/images/Sponsors/Rankings/9-OJ Props.jpg" alt="OJ Propos" width="220" border="0" /></a></p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p><a target="_blank" href="https://www.masterlineusa.com/?utm_source=usa-ws&utm_medium=site-sponsor"><img src="http://www.usawaterski.org/rankings/images/Sponsors/Rankings/15-Masterline.jpg" alt="Masterline" width="220" border="0" /></a></p>
<p>&nbsp;</p>
<p>&nbsp;</p>
</div>
<%



END SUB





' ----------------------
  SUB GetTheEventString (sUserSptsGrpID)
' ----------------------

	' --- Used in passing paramters to different modules such as view-tournamentsHQ and view-registration ---

	SELECT CASE sUserSptsGrpID
	   CASE ""
		sEventString = "sl=on&tr=on&ju=on&wb=on&ws=on&wu=on"
	   CASE "AWS"
		sEventString = "sl=on&tr=on&ju=on"
	   CASE "USW"
		sEventString = "wb=on&ws=on&wu=on"
	   CASE "AKA"
		sEventString = "kb=on"
	   CASE "ABC"
		sEventString = "bf=on"
	   CASE "HYD"
		sEventString = "hy=on"
	   CASE "JDC"
		sEventString = "jd=on"
	   CASE "ADC"
		sEventString = "ad=on"
	END SELECT



END SUB



' -----------------------------------
  SUB SetLogoParameters (sSptsGrpID)
' -----------------------------------

  SELECT CASE sSptsGrpID
    CASE "USW"
				MainLogo="/rankings/images/logos/logo_usw_150.gif"
				MainLogoWidth=145
				MainLogoHeight=60
				SD_Heading="Wakeboard"
    CASE "AWS"
				MainLogo="/rankings/images/logos/logo_awsa_sm.jpg"
				MainLogoWidth=100
				MainLogoHeight=60
				SD_Heading="Water Ski"
    CASE "AKA"
				MainLogo="/rankings/images/logos/logo_aka_75.jpg"
				MainLogoWidth=75
				MainLogoHeight=75
				SD_Heading="Kneeboard"
    CASE "NCW"
				MainLogo="/rankings/images/logos/logo_ncw_sm.jpg"
				MainLogoWidth=122
				MainLogoHeight=92
				SD_Heading="Collegiate"
  END SELECT

END SUB




' ----------------------------------
  SUB SetEventsOffered (sTSptsGrpID)
' ----------------------------------


' ++++++++++++++++++++++  DO NOT CHANGE THIS ONE  ++++++++++++++++++++

Session("sTEvent1")=""
Session("sTEvent2")=""
Session("sTEvent3")=""
Session("sTEvent4")=""
Session("sTEventRounds")=0

 ' --- Replace with some query from SWIFT ---
  SELECT CASE sTSptsGrpID
    CASE "AWS", "NCW"
				IF TRIM(Session("sTRoundsS"))<>"" THEN 
						Session("sTEvent1")="S"
						Session("sTEvent1Name") = "Slalom"
						IF TRIM(sEvent)="" THEN SetDefaultEvent "S", "Slalom", Session("sTRoundsS") ELSE Session("sTEventRounds")=Session("sTRoundsS") END IF
				END IF
				IF TRIM(Session("sTRoundsT"))<>"" THEN
						Session("sTEvent2")="T"
						Session("sTEvent2Name") = "Tricks"
						IF TRIM(sEvent)="" THEN SetDefaultEvent "T", "Tricks", Session("sTRoundsT") ELSE Session("sTEventRounds")=Session("sTRoundsT") END IF
				END IF
				IF TRIM(Session("sTRoundsJ"))<>"" THEN 
						Session("sTEvent3")="J"
						Session("sTEvent3Name") = "Jump"
						IF TRIM(sEvent)="" THEN SetDefaultEvent "J", "Jump", Session("sTRoundsJ") ELSE Session("sTEventRounds")=Session("sTRoundsJ") END IF
				END IF
	
				Session("sTEvent4")=""
				Session("sTEvent4Name") = ""


    CASE "USW"
				IF TRIM(Session("sTRoundsWakeBd"))<>"" THEN 
						Session("sTEvent1")="WB"
						Session("sTEvent1Name") = "Wakeboard"
						IF TRIM(sEvent)="" THEN SetDefaultEvent "WB", "Wakeboard", Session("sTRoundsWakeBd") ELSE Session("sTEventRounds")=Session("sTRoundsWakeBd") END IF
				END IF
				IF TRIM(Session("sTRoundsWSkate"))<>"" THEN 
						Session("sTEvent2")="WS"
						Session("sTEvent2Name") = "Wake Skate"
						IF TRIM(sEvent)="" THEN SetDefaultEvent "WS", "Wake Skate", Session("sTRoundsWSkate") ELSE Session("sTEventRounds")=Session("sTRoundsWSkate") END IF
				END IF
				IF TRIM(Session("sTRoundsWUrf"))<>"" THEN 
						Session("sTEvent3")="WU"
						Session("sTEvent3Name") = "Wake Surf"
						IF TRIM(sEvent)="" THEN SetDefaultEvent "WU", "Wake Surf", Session("sTRoundsWUrf") ELSE Session("sTEventRounds")=Session("sTRoundsWUrf") END IF
				END IF

				Session("sTEvent4")=""
				Session("sTEvent4Name") = ""


    CASE "AKA"
				IF TRIM(Session("sTRoundsS"))<>"" THEN 
						Session("sTEvent1")="S"
						Session("sTEvent1Name") = "Slalom"
						IF TRIM(sEvent)="" THEN SetDefaultEvent "S", "Slalom", Session("sTRoundsS") ELSE Session("sTEventRounds")=Session("sTRoundsS") END IF
				END IF
				IF TRIM(Session("sTRoundsT"))<>"" THEN 
						Session("sTEvent2")="T"
						Session("sTEvent2Name") = "Tricks"
						IF TRIM(sEvent)="" THEN SetDefaultEvent "T", "Tricks", Session("sTRoundsT") ELSE Session("sTEventRounds")=Session("sTRoundsT") END IF
				END IF
				IF TRIM(Session("sKRoundsFlip"))<>"" THEN 
						Session("sTEvent3")="KP"
						Session("sTEvent3Name") = "Flip"
						IF TRIM(sEvent)="" THEN SetDefaultEvent "KP", "Flip", Session("sKRoundsFlip") ELSE Session("sTEventRounds")=Session("sKRoundsFlip") END IF
				END IF
				IF TRIM(Session("sKRoundsFree"))<>"" THEN 
						Session("sTEvent4")="KR"
						Session("sTEvent4Name") = "Freestyle"
						IF TRIM(sEvent)="" THEN SetDefaultEvent "KR", "Freestyle", Session("sKRoundsFree") ELSE Session("sTEventRounds")=Session("sKRoundsFree") END IF
				END IF

  END SELECT


END SUB



' ------------------------------------------------------------
   SUB SetDefaultEvent (TempEvent, TempEventName, TempRounds)
' ------------------------------------------------------------

'response.write("Inside SetDefaultEvent TempRounds =  "&TempRounds)
'response.write("TempEvent =  "&TempEvent)
'response.write("sTEvent1Name =  "&TempEventName)

Session("sTEventRounds") = TempRounds
sEvent=TempEvent
'Session("sTEvent1Name")=TempEventName


END SUB





' --------------------------------------------
  SUB RegistrationEventsOffered (sTSptsGrpID)
' --------------------------------------------


' --- Updated 7-10-2008 per Brandon to cause WB to display slalom ---

' ---------------------------------------------------------------------------------------------------------------------
' ---  IMPORTANT:  The SEQUENCE in which the events are listed below is important and must match the ORDER BY clause
' ---	in the Registration program wherein the EventsDetail records are read in from the table. 
' ---------------------------------------------------------------------------------------------------------------------


' ***********************************************************************************************************
' **** MUST rework the manner in which system determines whether or not to display each header for column
' ***********************************************************************************************************


	' --- For populating SKILL Dropdown
	IF sTEventSlalom=true OR sTEventTrick=true THEN
		sSkillName1=""
		sSkillName2=""
		sSkillName3="Challenger"
		sSkillName4="Competitor"
		sSkillName5="Outlaw"
	ELSEIF sTEventWake=true OR Gr1USWPulls<>0 OR Gr2USW_WPulls<>0 THEN
		sSkillName1="Novice"
		sSkillName2="Intermediate"
		sSkillName3="Advanced"
		sSkillName4="Expert"
		sSkillName5="Outlaw"
	END IF	



' ******************
' --- IMPORTANT:  
' ******************
' --- Sequence events are defined must match the way the events are ORDERED in ReadFromRegisterEvents SUB in Registraton.asp  ---




	' --- Does not deal with PandC events ---
	EvtNo=0


	' ************************************
	' ************************************
	' *********   3 EVENT  ***************
	' ************************************
	' ************************************

	' *****  SLALOM  *****
	IF sTEventSlalom=true AND (sTSptsGrpID="AWS" OR sTSptsGrpID="NCW") THEN 
			EvtNo=EvtNo+1
			sShowGR(EvtNo)=false
			sShowStd(EvtNo)=false
			sShowRec(EvtNo)=false

			sTEvent(EvtNo)="S"		
			sTEventName(EvtNo) = "Slalom"
			
			
			' --- Defines which radio buttons to show for this event ---
			IF SClassC>0 AND SClassE+SClassL+SClassR=0 THEN 			' --- Only premier (C) option ---
					sShowStd(EvtNo)=true
					sShowRec(EvtNo)=false
			ELSEIF SClassC>0 AND SClassE+SClassL+SClassR>0 THEN		' --- C and a Record so BOTH Heading options ---
					sShowStd(EvtNo)=true
					sShowRec(EvtNo)=true
			ELSEIF SClassC=0 AND SClassE+SClassL+SClassR>0 AND TRIM(LCASE(sTBaseClassText))="record" THEN 	' --- Show only premier (Record) option ---	
					sShowStd(EvtNo)=false
					sShowRec(EvtNo)=true
					ResetRecHeadingToRecord = "Y"
			ELSEIF TRIM(sTBaseClassText)<>"" AND TRIM(sTUpgradeClassText)<>"" AND TRIM(LCASE(sTUpgradeClassText))<>"none" AND ((SClassE>0 AND SClassL>0) OR (SClassE>0 AND SClassR>0) OR (SClassL>0 AND SClassR>0)) THEN
					sShowStd(EvtNo)=true															' --- Separate options for two record Classes ---
					sShowRec(EvtNo)=true
			ELSEIF SClassC=0 AND SClassE+SClassL+SClassR>0 THEN 	' --- Show only premier (Record) option ---	
					sShowStd(EvtNo)=false
					sShowRec(EvtNo)=true
			END IF
		
			' --- Grassroots Beginning in 2010 ---
			IF Gr2AWS_SPulls<>0 THEN 
					sShowGR(EvtNo)=true
			END IF




			IF sTPandC=true THEN
					sTRounds(EvtNo) = 3
					IF sTPandCPulls <= 3 THEN sTRounds(EvtNo)=sTPandCPulls
			ELSE
					sTRounds(EvtNo)=sTRoundsS
			END IF

	END IF


	' *****  TRICKS  *****
	IF sTEventTrick=true AND (sTSptsGrpID="AWS" OR sTSptsGrpID="NCW") THEN
			EvtNo=EvtNo+1
			sShowGR(EvtNo)=false
			sShowStd(EvtNo)=false
			sShowRec(EvtNo)=false

			sTEvent(EvtNo)="T"
			sTEventName(EvtNo) = "Tricks"


			' --- Defines which radio buttons to show for this event ---
			IF TClassC>0 AND TClassE+TClassL+TClassR=0 THEN 			' --- Only premier (C) option ---
					sShowStd(EvtNo)=true
					sShowRec(EvtNo)=false
			ELSEIF TClassC>0 AND TClassE+TClassL+TClassR>0 THEN		' --- C and a Record so BOTH Heading options ---
					sShowStd(EvtNo)=true
					sShowRec(EvtNo)=true
			ELSEIF TClassC=0 AND TClassE+TClassL+TClassR>0 AND TRIM(LCASE(sTBaseClassText))="record" THEN 	' --- Show only premier (Record) option ---	
					sShowStd(EvtNo)=false
					sShowRec(EvtNo)=true
			ELSEIF TRIM(sTBaseClassText)<>"" AND TRIM(sTUpgradeClassText)<>"" AND TRIM(LCASE(sTUpgradeClassText))<>"none" AND ((TClassE>0 AND TClassL>0) OR (TClassE>0 AND TClassR>0) OR (TClassL>0 AND TClassR>0)) THEN
					sShowStd(EvtNo)=true															' --- Separate options for two record Classes ---
					sShowRec(EvtNo)=true
			ELSEIF TClassC=0 AND TClassE+TClassL+TClassR>0 THEN 	' --- Show only premier (Record) option ---	
					sShowStd(EvtNo)=false
					sShowRec(EvtNo)=true
			END IF

			' --- Beginning in 2010 ---
			IF Gr2AWS_TPulls<>0 THEN 
					sShowGR(EvtNo)=true
			END IF

			IF sTPandC=true THEN
					sTRounds(EvtNo) = 3
					IF sTPandCPulls <= 3 THEN sTRounds(EvtNo)=sTPandCPulls
			ELSE
					sTRounds(EvtNo)=sTRoundsT
			END IF

	END IF


	
	' *****  JUMP  *****
	IF sTEventJump=true AND (sTSptsGrpID="AWS" OR sTSptsGrpID="NCW") THEN 
			EvtNo=EvtNo+1
			sShowGR(EvtNo)=false
			sShowStd(EvtNo)=false
			sShowRec(EvtNo)=false

			sTEvent(EvtNo)="J"
			sTEventName(EvtNo) = "Jump"


			' --- Defines which radio buttons to show for this event ---
			IF JClassC>0 AND JClassE+JClassL+JClassR=0 THEN 			' --- Only premier (C) option ---
					sShowStd(EvtNo)=true
					sShowRec(EvtNo)=false
			ELSEIF JClassC>0 AND JClassE+JClassL+JClassR>0 THEN		' --- C and a Record so BOTH Heading options ---
					sShowStd(EvtNo)=true
					sShowRec(EvtNo)=true
			ELSEIF JClassC=0 AND JClassE+JClassL+JClassR>0 AND TRIM(LCASE(sTBaseClassText))="record" THEN 	' --- Show only premier (Record) option ---	
					sShowStd(EvtNo)=false
					sShowRec(EvtNo)=true
			ELSEIF TRIM(sTBaseClassText)<>"" AND TRIM(sTUpgradeClassText)<>"" AND TRIM(LCASE(sTUpgradeClassText))<>"none" AND ((JClassE>0 AND JClassL>0) OR (JClassE>0 AND JClassR>0) OR (JClassL>0 AND JClassR>0)) THEN
					sShowStd(EvtNo)=true															' --- Separate options for two record Classes ---
					sShowRec(EvtNo)=true
			ELSEIF JClassC=0 AND JClassE+JClassL+JClassR>0 THEN 	' --- Show only premier (Record) option ---	
					sShowStd(EvtNo)=false
					sShowRec(EvtNo)=true
			END IF


			' --- Beginning in 2010 ---
			IF JClassN<>0 THEN 
					sShowGR(EvtNo)=true
			END IF

			IF sTPandC=true THEN
					sTRounds(EvtNo) = 3
					IF sTPandCPulls <= 3 THEN sTRounds(EvtNo)=sTPandCPulls
			ELSE
					sTRounds(EvtNo)=sTRoundsJ
			END IF

	END IF




 	' ----------------------------------------------------------------------------------------
	' --- Display separate GrassRoots AWS event only when there is no slalom event offered ---
 	' ----------------------------------------------------------------------------------------


y=2
IF y=1 AND Session("AdminMenuLevel")>49 THEN
	IF LEFT(sTourID,6)="13M081" THEN
			'response.write("<br>Line 547 Tools_Include - sShowGR("&EvtNo&") = "&sShowGR(EvtNo))
			response.write("<br><br><br>Line 547 Tools_Include")

			response.write("<br>Gr1AWSPulls = "&Gr1AWSPulls)
			response.write("<br>sTEventSlalom = "&sTEventSlalom)
			response.write("<br>sTEventTrick = "&sTEventTrick)
			response.write("<br>Gr2AWS_SPulls = "&Gr2AWS_SPulls)
			response.write("<br>Gr2AWS_TPulls = "&Gr2AWS_TPulls)


			response.write("<br>sShowGR("&EvtNo&") - "&sShowGR(EvtNo))
			response.write("<br>sTEventName("&EvtNo&") - "&sTEventName(EvtNo))
	END IF
END IF	



	IF Gr2AWS_SPulls=0 AND Gr1AWSPulls<>0 THEN
			EvtNo=EvtNo+1
			sShowGR(EvtNo)=true
			sShowStd(EvtNo)=false
			sShowRec(EvtNo)=false

			sTEvent(EvtNo)="3G"
			sTEventName(EvtNo) = "AWSA Grassroots"
			sTRounds(EvtNo)=sTRoundsS

  ' --- Added 7-14-2011 because of TourID=11W190 --
'	ELSEIF Gr2AWS_SPulls<>0 AND Gr1AWSPulls=0 THEN
	' --- Changed 7-5-2013 
	'ELSEIF sShowStd(EvtNo)=false AND sShowRec(EvtNo)=false AND Gr2AWS_SPulls<>0 AND Gr1AWSPulls=0 THEN
'			EvtNo=EvtNo+1
'			sShowGR(EvtNo)=true
'			sShowStd(EvtNo)=false
'			sShowRec(EvtNo)=false
'			sTEvent(EvtNo)="3G"
'			sTEventName(EvtNo) = "AWSA Grassroots"
'			sTRounds(EvtNo)=Gr2AWS_SPulls

	ELSEIF sTEventSlalom=false AND Gr2AWS_SPulls=0 AND (sTEventF3ev=true OR Gr1AWSPulls<>0) THEN	' --- Legacy for 2009 ---
			EvtNo=EvtNo+1
			sShowGR(EvtNo)=true
			sShowStd(EvtNo)=false
			sShowRec(EvtNo)=false

			sTEvent(EvtNo)="3G"
			sTEventName(EvtNo) = "AWSA Grassroots"
			sTRounds(EvtNo)=sTRoundsS
	END IF


' --- Temporary fix to insure record is always in the right column --
IF ResetRecHeadingToRecord = "Y" THEN sTUpgradeClassText="Record"








	' ****************************
	' ********  WAKEBOARD ********
	' ****************************


	' --- Wakeboard for 2010 ---
	IF WWakeW>0 OR Gr1USWPulls<>0 OR Gr2USW_WPulls<>0 THEN
		EvtNo=EvtNo+1
		sShowGR(EvtNo)=false
		sShowStd(EvtNo)=false
		sShowRec(EvtNo)=false

		sTEvent(EvtNo)="WB"
		sTEventName(EvtNo) = "Wakeboard"

		IF WWakeW>0 THEN 
			sTRounds(EvtNo)=WWakeW
			sShowStd(EvtNo)=true
		END IF

		IF Gr1USWPulls<>0 OR Gr2USW_WPulls<>0 THEN
			IF Gr1USWPulls>sTRounds(EvtNo) THEN sTRounds(EvtNo)=Gr1USWPulls
			IF Gr2USW_WPulls>sTRounds(EvtNo) THEN sTRounds(EvtNo)=Gr2USW_WPulls
			sShowGR(EvtNo)=true
		END IF

	' --- Legacy 2009 structure
	ELSEIF sTEventFW=true OR sTEventWake=true OR Gr1USWPulls<>0 OR Gr2USW_WPulls<>0 THEN
		EvtNo=EvtNo+1
		sShowGR(EvtNo)=false
		sShowStd(EvtNo)=false
		sShowRec(EvtNo)=false

		sTEvent(EvtNo)="WB"
		sTEventName(EvtNo) = "Wakeboard"

		IF sTEventWake=true THEN
			sTRounds(EvtNo)=sTRoundsWakeBd
			sShowStd(EvtNo)=true
		END IF

		IF Gr1USWPulls<>0 OR Gr2USW_WPulls<>0 THEN
			IF Gr1USWPulls>sTRounds(EvtNo) THEN sTRounds(EvtNo)=Gr1USWPulls
			IF Gr2USW_WPulls>sTRounds(EvtNo) THEN sTRounds(EvtNo)=Gr2USW_WPulls
			sShowGR(EvtNo)=true
		END IF
	END IF


	' --- WakeSkate ---
	IF WSkateW>0 OR Gr2USW_SkatePulls<>0 THEN 
		EvtNo=EvtNo+1
		sShowGR(EvtNo)=false
		sShowStd(EvtNo)=false
		sShowRec(EvtNo)=false

		sTEvent(EvtNo)="WS"		
		sTEventName(EvtNo) = "Wake Skate"

		IF WSkateW>0 THEN 
			sTRounds(EvtNo)=WSkateW
			sShowStd(EvtNo)=true
		END IF 

		IF Gr2USW_SkatePulls<>0 THEN
			IF Gr2USW_SkatePulls>sTRounds(EvtNo) THEN sTRounds(EvtNo)=Gr2USW_SkatePulls
			sShowGR(EvtNo)=true
		END IF

	' --- Legacy 2009 structure
	ELSEIF sTEventWSkate=true OR Gr2USW_SkatePulls<>0 THEN 
		EvtNo=EvtNo+1
		sShowGR(EvtNo)=false
		sShowStd(EvtNo)=false
		sShowRec(EvtNo)=false

		sTEvent(EvtNo)="WS"		
		sTEventName(EvtNo) = "Wake Skate"

		IF sTEventWSkate=true THEN 
			sTRounds(EvtNo)=sTRoundsWSkate
			sShowStd(EvtNo)=true
		END IF 

		IF Gr2USW_SkatePulls<>0 THEN
			IF Gr2USW_SkatePulls>sTRounds(EvtNo) THEN sTRounds(EvtNo)=Gr2USW_SkatePulls
			sShowGR(EvtNo)=true
		END IF
	END IF

'IF Session("AdminMenuLevel")>=50 THEN
'Y=13
'IF MarkTester AND Y=12 THEN
	'response.write("<br>MarkTester="&MarkTester)
'	response.write("<br>WSkateW = "&WSkateW)
'	response.write("<br>sTEvent(EvtNo) = "&sTEvent(EvtNo))
	'response.write("<br>sTEventWake = "&sTEventWake)
	'response.write("<br>Gr1USWPulls = "&Gr1USWPulls)
	'response.write("<br>Gr2USW_WPulls = "&Gr2USW_WPulls)
'END IF



	' --- WakeSurf ---
	IF WSurfW>0 OR Gr2USW_SurfPulls<>0 THEN 
		EvtNo=EvtNo+1
		sShowGR(EvtNo)=false
		sShowStd(EvtNo)=false
		sShowRec(EvtNo)=false

		sTEvent(EvtNo)="WU"		
		sTEventName(EvtNo) = "Wake Surf"

		IF WSurfW>0 THEN 
			sTRounds(EvtNo)=WSurfW
			sShowStd(EvtNo)=true
		END IF 

		IF Gr2USW_SurfPulls<>0 THEN
			IF Gr2USW_SurfPulls>sTRounds(EvtNo) THEN sTRounds(EvtNo)=Gr2USW_SurfPulls
			sShowGR(EvtNo)=true
		END IF

	' --- Legacy 2009 structure
	ELSEIF sTEventWSurf=true OR Gr2USW_SurfPulls<>0 THEN 
		EvtNo=EvtNo+1
		sShowGR(EvtNo)=false
		sShowStd(EvtNo)=false
		sShowRec(EvtNo)=false

		sTEvent(EvtNo)="WU"		
		sTEventName(EvtNo) = "Wake Surf"

		IF sTEventWSurf=true THEN 
			sTRounds(EvtNo)=sTRoundsWSurf
			sShowStd(EvtNo)=true
		END IF 

		IF Gr2USW_SurfPulls<>0 THEN
			IF Gr2USW_SurfPulls>sTRounds(EvtNo) THEN sTRounds(EvtNo)=Gr2USW_SurfPulls
			sShowGR(EvtNo)=true
		END IF
	END IF

	' --- RailJam Grassroots ---
	IF Gr2USW_RailJamPulls<>0 THEN 
		EvtNo=EvtNo+1
		sShowGR(EvtNo)=false
		sShowStd(EvtNo)=false
		sShowRec(EvtNo)=false

		sTEvent(EvtNo)="WJ"		
		sTEventName(EvtNo) = "Rail Jam"

		IF Gr2USW_RailJamPulls<>0 THEN
			IF Gr2USW_RailJamPulls>sTRounds(EvtNo) THEN sTRounds(EvtNo)=Gr2USW_RailJamPulls
			sShowGR(EvtNo)=true
		END IF
	END IF


	



	' ****************************
	' ********  KNEEBOARD ********
	' ****************************

	'IF (sTSptsGrpID="AKA" AND sTEventSlalom=true) OR Gr1AKAPulls<>0 OR Gr2AKA_SPulls<>0 OR sTEventFKB=true THEN 
	IF KSClassQ>0 OR KSClassT>0 OR Gr1AKAPulls<>0 OR Gr2AKA_SPulls<>0 THEN 
		EvtNo=EvtNo+1
		sShowGR(EvtNo)=false
		sShowStd(EvtNo)=false
		sShowRec(EvtNo)=false

		sTEvent(EvtNo)="KS"		
		sTEventName(EvtNo) = "Kneeboard Slalom"

		IF KSClassT>0 THEN 
			sTRounds(EvtNo)=KSClassT
			sShowRec(EvtNo)=true
		END IF

		IF KSClassQ>0 THEN
			sTRounds(EvtNo)=KSClassQ
			sShowStd(EvtNo)=true
		END IF
		
		IF Gr2AKA_SPulls<>0 THEN
			IF Gr2AKA_SPulls>sTRounds(EvtNo) THEN sTRounds(EvtNo)=Gr2AKA_SPulls
			sShowGR(EvtNo)=true
		END IF

		IF Gr1AKAPulls<>0 THEN
			sTRounds(EvtNo)=Gr1AKAPulls
			sShowGR(EvtNo)=true
			sTEventName(EvtNo) = "Kneeboard Fun"
		END IF
	END IF

	IF KTClassQ>0 OR KTClassT>0 OR Gr2AKA_TPulls<>0 THEN 
	'IF (sTSptsGrpID="AKA" AND sTEventTrick=true) OR Gr2AKA_TPulls<>0 THEN 
		EvtNo=EvtNo+1
		sShowGR(EvtNo)=false
		sShowStd(EvtNo)=false
		sShowRec(EvtNo)=false

		sTEvent(EvtNo)="KT"
		sTEventName(EvtNo) = "Kneeboard Tricks"

		IF KTClassT>0 THEN 
			sTRounds(EvtNo)=KTClassT
			sShowRec(EvtNo)=true
		END IF

		IF KTClassQ>0 THEN
			sTRounds(EvtNo)=KTClassQ
			sShowStd(EvtNo)=true
		END IF
		
		IF Gr2AKA_TPulls<>0 THEN
			IF Gr2AKA_TPulls>sTRounds(EvtNo) THEN sTRounds(EvtNo)=Gr2AKA_TPulls
				sShowGR(EvtNo)=true
		END IF
	END IF

	IF KFLClassQ>0 OR KFLClassT>0 OR Gr2AKA_FlipPulls<>0 THEN 
	'IF sTEventFlip=true OR Gr2AKA_FlipPulls<>0 THEN 
		EvtNo=EvtNo+1
		sShowGR(EvtNo)=false
		sShowStd(EvtNo)=false
		sShowRec(EvtNo)=false

		sTEvent(EvtNo)="KF"
		sTEventName(EvtNo) = "Kneeboard Flip"

		IF KFLClassT>0 THEN 
			sTRounds(EvtNo)=KFLClassT
			sShowRec(EvtNo)=true
		END IF

		IF KFLClassQ>0 THEN
			sTRounds(EvtNo)=KFLClassQ
			sShowStd(EvtNo)=true
		END IF
		
		IF Gr2AKA_FlipPulls<>0 THEN
			IF Gr2AKA_FlipPulls>sTRounds(EvtNo) THEN sTRounds(EvtNo)=Gr2AKA_FlipPulls
			sShowGR(EvtNo)=true
		END IF
	END IF

	IF KFrClassQ>0 OR KFrClassT>0 OR Gr2AKA_FreePulls<>0 THEN 
	'IF sTEventFree=true OR Gr2AKA_FreePulls<>0 THEN 
		EvtNo=EvtNo+1
		sShowGR(EvtNo)=false
		sShowStd(EvtNo)=false
		sShowRec(EvtNo)=false

		sTEvent(EvtNo)="KR"
		sTEventName(EvtNo) = "Kneeboard Freestyle"

		IF KFrClassT>0 THEN 
			sTRounds(EvtNo)=KFrClassT
			sShowRec(EvtNo)=true
		END IF

		IF KFrClassQ>0 THEN
			sTRounds(EvtNo)=KFrClassQ
			sShowStd(EvtNo)=true
		END IF
		
		IF Gr2AKA_FreePulls<>0 THEN
			IF Gr2AKA_FreePulls>sTRounds(EvtNo) THEN sTRounds(EvtNo)=Gr2AKA_FreePulls
			sShowGR(EvtNo)=true
		END IF
	END IF




	' ****************************
	' *******   HYDROFOIL  *******
	' ****************************



	IF Gr2USH_FreeRidePulls<>0 OR Gr1USHPulls<>0 THEN 
		EvtNo=EvtNo+1
		sShowGR(EvtNo)=true
		sShowStd(EvtNo)=false
		sShowRec(EvtNo)=false

		sTEvent(EvtNo)="HF"
		sTEventName(EvtNo) = "Hydrofoil Free Ride"
		sTRounds(EvtNo)=Gr2USH_FreeRidePulls

		IF Gr1USHPulls>Gr2USH_FreeRidePulls THEN
			sTRounds(EvtNo)=Gr1USHPulls
			IF Gr2USH_FreeRidePulls=0 THEN sTEventName(EvtNo) = "Hydrofoil Fun Day"
		END IF
	END IF

	IF Gr2USH_JumpOutPulls<>0 THEN 
		EvtNo=EvtNo+1
		sShowGR(EvtNo)=true
		sShowStd(EvtNo)=false
		sShowRec(EvtNo)=false

		sTEvent(EvtNo)="HJ"
		sTEventName(EvtNo) = "Hydrofoil Jump Out"
		sTRounds(EvtNo)=Gr2USH_JumpOutPulls
	END IF

	IF Gr2USH_BigAirPulls<>0 THEN 
		EvtNo=EvtNo+1
		sShowGR(EvtNo)=true
		sShowStd(EvtNo)=false
		sShowRec(EvtNo)=false

		sTEvent(EvtNo)="HB"
		sTEventName(EvtNo) = "Hydrofoil Big Air"
		sTRounds(EvtNo)=Gr2USH_BigAirPulls
	END IF

	IF Gr2USH_3TricksPulls<>0 THEN 
		EvtNo=EvtNo+1
		sShowGR(EvtNo)=true
		sShowStd(EvtNo)=false
		sShowRec(EvtNo)=false

		sTEvent(EvtNo)="H3"
		sTEventName(EvtNo) = "Hydrofoil 3-Tricks"
		sTRounds(EvtNo)=Gr2USH_3TricksPulls
	END IF


	' ******************************
	' ********   DISABLED  *********
	' ******************************
	IF Gr1WSDPulls<>0  THEN
		EvtNo=EvtNo+1
		sShowGR(EvtNo)=true
		sShowStd(EvtNo)=false
		sShowRec(EvtNo)=false

		sTEvent(EvtNo)="DA"
		sTEventName(EvtNo) = "Disabled"
		sTRounds(EvtNo)=1
	END IF



	' ****************************
	' ********  BAREFOOT ********
	' ****************************


	IF sTEventFB=true THEN		
			' ---------------------------------------------------------------------
			' --- Used for inclusion of barefoot as an event in an AWS sanction ---
			' ---------------------------------------------------------------------
			EvtNo=EvtNo+1
			sTEvent(EvtNo)="BF"
			sTEventName(EvtNo) = "Barefoot"
			sShowGR(EvtNo)=true
			sShowStd(EvtNo)=false
			sShowRec(EvtNo)=false
			sTRounds(EvtNo)=1
	END IF


  ' -------------------------
  ' --- BAREFOOT - SLALOM ---
  ' -------------------------


	IF BSClassC>0 OR BSClassL>0 OR BSClassR>0 OR Gr2ABC_SPulls>0 THEN 
			EvtNo=EvtNo+1
			sShowGR(EvtNo)=false
			sShowStd(EvtNo)=false
			sShowRec(EvtNo)=false

			sTEvent(EvtNo)="BS"		
			sTEventName(EvtNo) = "Barefoot Slalom"

			' --- Show only premier (C) option ---
			IF BSClassC>0 THEN 
					sShowStd(EvtNo)=true
			END IF
			' --- Show only premier (Record) option ---
			IF BSClassR>0 THEN 
					sShowStd(EvtNo)=false
					sShowRec(EvtNo)=true
			END IF
			' --- Show both Heading options ---
			IF BSClassC>0 AND (BSClassL>0 OR BSClassR>0) THEN
					sShowStd(EvtNo)=true
					sShowRec(EvtNo)=true
			END IF
		
			IF Gr2ABC_SPulls<>0 THEN
				IF Gr2ABC_SPulls>sTRounds(EvtNo) THEN sTRounds(EvtNo)=Gr2ABC_SPulls
				sShowGR(EvtNo)=true
			END IF

	END IF

  ' -----------------------
  ' --- Barefoot TRICKS ---
  ' -----------------------
	IF BTClassC>0 OR BTClassL>0 OR BTClassR>0 OR Gr2ABC_TPulls>0  THEN 
			EvtNo=EvtNo+1
			sShowGR(EvtNo)=false
			sShowStd(EvtNo)=false
			sShowRec(EvtNo)=false

			sTEvent(EvtNo)="BT"		
			sTEventName(EvtNo) = "Barefoot Tricks"

			' --- Show only premier (C) option ---
			IF BTClassC>0 THEN 
					sShowStd(EvtNo)=true
			END IF
			' --- Show only premier (Record) option ---
			IF BTClassL>0 OR BTClassR>0 THEN 
					sShowStd(EvtNo)=false
					sShowRec(EvtNo)=true
			END IF
			' --- Show both Heading options ---
			IF BTClassC>0 AND (BTClassL>0 OR BTClassR>0) THEN
					sShowStd(EvtNo)=true
					sShowRec(EvtNo)=true
			END IF

			IF Gr2ABC_TPulls<>0 THEN
				IF Gr2ABC_TPulls>sTRounds(EvtNo) THEN sTRounds(EvtNo)=Gr2ABC_TPulls
				sShowGR(EvtNo)=true
			END IF

	END IF

  ' -----------------------
  ' --- Barefoot JUMP ---
  ' -----------------------
	IF BJClassC>0 OR BJClassL>0 OR BJClassR>0 THEN 
			EvtNo=EvtNo+1
			sShowGR(EvtNo)=false
			sShowStd(EvtNo)=false
			sShowRec(EvtNo)=false

			sTEvent(EvtNo)="BJ"		
			sTEventName(EvtNo) = "Barefoot Jump"

			' --- Show only premier (C) option ---
			IF BJClassC>0 THEN 
					sShowStd(EvtNo)=true
			END IF
			' --- Show only premier (Record) option ---
			IF BJClassL>0 OR BJClassR>0 THEN 
					sShowStd(EvtNo)=false
					sShowRec(EvtNo)=true
			END IF
			' --- Show both Heading options ---
			IF BJClassC>0 AND (BJClassL>0 OR BJClassR>0) THEN
					sShowStd(EvtNo)=true
					sShowRec(EvtNo)=true
			END IF
	END IF
  
  ' ---------------------------------------
  ' --- Barefoot GRASSROOTS Stand Alone ---
  ' ---------------------------------------

	IF Gr1ABCPulls<>0 THEN	
			EvtNo=EvtNo+1
			sShowGR(EvtNo)=true

			sTEvent(EvtNo)="BG"
			sTEventName(EvtNo) = "Barefoot Grassroots"
			sTRounds(EvtNo)=Gr1ABCPulls
	END IF







Session("TotEv")=EvtNo


		' --- Uses selections to set overall heading controls ---
	FOR EvtNo=1 TO TotEv
			IF sShowGR(EvtNo)=true THEN ShowGRhead=true
			IF sShowStd(EvtNo)=true THEN ShowStdHead=true
			IF sShowRec(EvtNo)=true THEN ShowRecHead=true
	NEXT



END SUB



' ------------------------------------------------------------
   SUB SetDefaultEvent (TempEvent, TempEventName, TempRounds)
' ------------------------------------------------------------

'response.write("Inside SetDefaultEvent TempRounds =  "&TempRounds)
'response.write("TempEvent =  "&TempEvent)
'response.write("sTEvent1Name =  "&TempEventName)

Session("sTEventRounds") = TempRounds
sEvent=TempEvent
'Session("sTEvent1Name")=TempEventName


END SUB



' -------------------------------------
  SUB CreateSkiYearDropDown (sSkiYearID)
' -------------------------------------

' ------------   Builds Ski Year Drop Down list of all Ski Years EXCEPT 12 Month  ----------------- 

sSQL = "SELECT * FROM "&SkiYearTableName&" WHERE SkiYearID<>1 ORDER BY SkiYearID DESC"
SET rsSY=Server.CreateObject("ADODB.recordset")
rsSY.open sSQL, SConnectionToTRATable

'response.write("<br>EOF")
'response.write(rsSY.eof)
'response.write("<br>sSkiYearID="&sSkiYearID)

'response.end

%>
<SELECT name=sSkiYearID style="width:10em"><%
  DO WHILE NOT rsSY.eof %>
	<option value = "<%=rsSY("SkiYearID")%>" <%IF rsSY("SkiYearID") = CInt(sSkiYearID) THEN response.write(" selected ")%>><%=rsSY("SkiYearName")%></option><%
	rsSY.movenext
  LOOP %>
</SELECT><%

rsSY.close

END SUB






'----------------------------------------------------------------------------------------------
 SUB LoadDivDropWithAgeGender (DivSelected, EventSelected, DivDropName, DivDropStatus)
'----------------------------------------------------------------------------------------------


' Loads applicable divisions into a division pulldown for each event selected

'  ---  DAVE CLARK "NOTE:" Need to add filter to filter to current SkiYear  ----

IF sTSptsGrpID="AWS" OR sTSptsGrpID="NCW" THEN
	sSQL = "SELECT DT.div, DT.div_name, DT.Up_Age FROM "&DivisionsTableName&" as DT"
ELSE
	sSQL = "SELECT DT.div, DT.div_name, DT.Up_Age FROM "&DivisionsOtherTableName&" as DT"
END IF
sSQL = sSQL + " WHERE DT.SkiYearID = (Select SkiYearID from "&SkiYearTableName&" where DefaultYear=1)"


'response.write("sTSptsGrpID="&sTSptsGrpID)

SELECT CASE sTSptsGrpID
  	CASE "AWS"
		sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'y' AND lower(left(DT.div,1)) <> 'x'  AND lower(left(DT.div,1)) <> 'l'  AND lower(left(DT.div,1)) <> 's'  AND lower(left(DT.div,1)) <> 'e'" 
		sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'c'"
		

		SELECT CASE TRIM(EventSelected) 
		   CASE "S"
					IF NOT (sTHSClassR OR sTHSClassL OR SClassR>0 OR SClassL>0) THEN sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'i' AND lower(left(DT.div,1)) <> 'j'"
					IF ((NOT sTHSClassN) AND (NOT sTHSClassF) AND (NOT sTEventF3Ev)) THEN sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'n'"

		   CASE "T"
					IF NOT (sTHTClassR OR sTHTClassL OR TClassR>0 OR TClassL>0) THEN sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'i' AND lower(left(DT.div,1)) <> 'j'"
					IF ((NOT sTHTClassN) AND (NOT sTHTClassF)) THEN sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'n'"

		   CASE "J"
					IF NOT (sTHJClassR OR sTHJClassL OR JClassR>0 OR JClassL>0) THEN sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'i' AND lower(left(DT.div,1)) <> 'j' AND lower(DT.div)<>'b1' AND lower(DT.div)<>'g1'"
					IF ((NOT sTHJClassN) AND (NOT sTHJClassN)) THEN sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'n' AND lower(DT.div)<>'b1' AND lower(DT.div)<>'g1'"

		   CASE "3G"	
					IF NOT (sTHSClassR OR sTHSClassL) THEN sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'i' AND lower(left(DT.div,1)) <> 'j'"
					sSQL = sSQL + " AND lower(left(DT.div,1)) = 'n'"

		   CASE ELSE	
					sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'i' AND lower(left(DT.div,1)) <> 'j'"			

		END SELECT

		SELECT CASE sMembSex
			CASE "Male" 
					sSQL = sSQL + " AND LOWER(left(DT.sex,1)) = 'm'" 
			CASE "Female" 
					sSQL = sSQL + " AND LOWER(left(DT.sex,1)) = 'f'" 
		END SELECT

		IF sMembAge > 1 AND NOT IsNull(sMembAge) THEN
				sSQL = sSQL + " AND DT.Up_Age >= "&sMembAge&" AND DT.Low_Age <= "&sMembAge   
		END IF


	CASE "NCW"

		SELECT CASE sMembSex
			CASE "Male" 
				sSQL = sSQL + " AND LOWER(left(DT.sex,1)) = 'm'" 
			CASE "Female" 
				sSQL = sSQL + " AND LOWER(left(DT.sex,1)) = 'f'" 
		END SELECT

		sSQL = sSQL + " AND (lower(DT.div) = 'cm' or lower(DT.div) = 'cw')"

	CASE "USW"
		sSQL = sSQL + " AND SptsGrpID='USW'" 

		SELECT CASE sMembSex
			CASE "Male" 
				sSQL = sSQL + " AND LOWER(left(DT.sex,1)) = 'm'" 
			CASE "Female" 
				sSQL = sSQL + " AND LOWER(left(DT.sex,1)) = 'f'" 
		END SELECT

		IF sMembAge > 1 AND NOT IsNull(sMembAge) THEN
			sSQL = sSQL + " AND DT.Up_Age >= "&sMembAge&" AND DT.Low_Age <= "&sMembAge   
		END IF

	CASE "AKA"
		sSQL = sSQL + " AND SptsGrpID='AKA'" 

		SELECT CASE sMembSex
			CASE "Male" 
				sSQL = sSQL + " AND LOWER(left(DT.sex,1)) = 'm'" 
			CASE "Female" 
				sSQL = sSQL + " AND LOWER(left(DT.sex,1)) = 'f'" 
		END SELECT

		IF sMembAge > 1 AND NOT IsNull(sMembAge) THEN
			sSQL = sSQL + " AND DT.Up_Age >= "&sMembAge&" AND DT.Low_Age <= "&sMembAge   
		END IF

	CASE "USH"
		sSQL = sSQL + " AND SptsGrpID='USH'" 

		SELECT CASE sMembSex
			CASE "Male" 
				sSQL = sSQL + " AND LOWER(left(DT.sex,1)) = 'm'" 
			CASE "Female" 
				sSQL = sSQL + " AND LOWER(left(DT.sex,1)) = 'f'" 
		END SELECT

		IF sMembAge > 1 AND NOT IsNull(sMembAge) THEN
			sSQL = sSQL + " AND DT.Up_Age >= "&sMembAge&" AND DT.Low_Age <= "&sMembAge   
		END IF

	CASE "ABC"
		sSQL = sSQL + " AND SptsGrpID='ABC'" 

		SELECT CASE sMembSex
			CASE "Male" 
				sSQL = sSQL + " AND LOWER(left(DT.sex,1)) = 'm'" 
			CASE "Female" 
				sSQL = sSQL + " AND LOWER(left(DT.sex,1)) = 'f'" 
		END SELECT

		IF sMembAge > 1 AND NOT IsNull(sMembAge) THEN
			sSQL = sSQL + " AND DT.Up_Age >= "&sMembAge&" AND DT.Low_Age <= "&sMembAge   
		END IF


END SELECT


SET rsDivisions=Server.CreateObject("ADODB.recordset")
sSQL = sSQL + " ORDER BY CASE WHEN LOWER(LEFT(DT.div,1)) IN ('w','m') THEN 1 "
sSQL = sSQL + "   WHEN LOWER(LEFT(DT.div,1)) IN ('b','g') THEN 2"
sSQL = sSQL + "   WHEN LOWER(LEFT(DT.div,1)) IN ('o') THEN 3"
sSQL = sSQL + "   WHEN LOWER(LEFT(DT.div,1)) IN ('i') THEN 4"
sSQL = sSQL + "   ELSE 5"
sSQL = sSQL + "  END " 

rsDivisions.open sSQL, SConnectionToTRATable


DispSQL="N"
IF Session("AdminMenuLevel")>=50 AND DispSQL="Y" THEN
	response.write("<br>")
	response.write(sSQL)
END IF


' ---------------------------------------------------------------------------------------------
' --------------- Builds VALID Division DROP DOWN list based on criteria above  ---------------
' ---------------------------------------------------------------------------------------------



%><select name='<%=DivDropName%>' <%=DivDropStatus%> style="width:12em"><%

IF NOT rsDivisions.eof THEN 
  	rsDivisions.movefirst

  	DO WHILE NOT rsDivisions.eof
		IF TRIM(rsDivisions("Div")) = DivSelected THEN %>
			<option value="<%=rsDivisions("Div")%>" selected><%=rsDivisions("Div")%> - <%=rsDivisions("Div_Name")%></option><br><%
		ELSEIF DivSelected = "" AND rsDivisions("Up_Age") < 999 THEN %>
			<option value="<%=rsDivisions("Div")%>" selected><%=rsDivisions("Div")%> - <%=rsDivisions("Div_Name")%></option><br><%
    		ELSE %>
			<option value="<%=rsDivisions("Div")%>"><%=rsDivisions("Div")%> - <%=rsDivisions("Div_Name")%></option><br><%
		END IF	

		rsDivisions.moveNEXT
	LOOP
ELSE
	response.write("<option value =""NA"" selected>None Available</option>")
END IF  %>
</select><%

rsDivisions.close

END SUB



'----------------------------------------------------------------------------------------------
 SUB LoadDivDropWithAgeGender_IntlIncluded (DivSelected, EventSelected, DivDropName, DivDropStatus)
'----------------------------------------------------------------------------------------------


' Loads applicable divisions into a division pulldown for each event selected

'  ---  DAVE CLARK "NOTE:" Need to add filter to filter to current SkiYear  ----

IF sTSptsGrpID="AWS" OR sTSptsGrpID="NCW" THEN
	sSQL = "SELECT DT.div, DT.div_name, DT.Up_Age FROM "&DivisionsTableName&" as DT"
ELSE
	sSQL = "SELECT DT.div, DT.div_name, DT.Up_Age FROM "&DivisionsOtherTableName&" as DT"
END IF
sSQL = sSQL + " WHERE DT.SkiYearID = (Select SkiYearID from "&SkiYearTableName&" where DefaultYear=1)"


'response.write("sTSptsGrpID="&sTSptsGrpID)

SELECT CASE sTSptsGrpID
  	CASE "AWS"
		'sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'y' AND lower(left(DT.div,1)) <> 'x'  AND lower(left(DT.div,1)) <> 'l'  AND lower(left(DT.div,1)) <> 's'  AND lower(left(DT.div,1)) <> 'e'" 
		' --- 
		sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'y'" 
		sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'x'"
		sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'e'"
		sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'c'"
		sSQL = sSQL + " AND lower(left(DT.div,1)) <> 's'"		

		SELECT CASE TRIM(EventSelected) 
		   CASE "S"
			IF NOT (sTHSClassR OR sTHSClassL OR SClassR>0 OR SClassL>0 OR SClassE>0) THEN sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'i'"
			IF ((NOT sTHSClassN) AND (NOT sTHSClassF) AND (NOT sTEventF3Ev)) THEN sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'n'"

		   CASE "T"
			IF NOT (sTHTClassR OR sTHTClassL OR TClassR>0 OR TClassL>0 OR TClassE>0) THEN sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'i'"
			IF ((NOT sTHTClassN) AND (NOT sTHTClassF)) THEN sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'n'"

		   CASE "J"
			IF NOT (sTHJClassR OR sTHJClassL OR JClassR>0 OR JClassL>0 OR JClassE>0) THEN sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'i'"
			IF ((NOT sTHJClassN) AND (NOT sTHJClassN)) THEN sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'n'"

		   CASE "3G"	
			IF NOT (sTHSClassR OR sTHSClassL) THEN sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'i'"
			sSQL = sSQL + " AND lower(left(DT.div,1)) = 'n'"

		   CASE ELSE	
			sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'i'"			

		END SELECT

		SELECT CASE sMembSex
			CASE "Male" 
				sSQL = sSQL + " AND LOWER(left(DT.sex,1)) = 'm'" 
			CASE "Female" 
				sSQL = sSQL + " AND LOWER(left(DT.sex,1)) = 'f'" 
		END SELECT

		IF sMembAge > 1 AND NOT IsNull(sMembAge) THEN
			sSQL = sSQL + " AND DT.Up_Age >= "&sMembAge&" AND DT.Low_Age <= "&sMembAge   
		END IF


	CASE "NCW"

		SELECT CASE sMembSex
			CASE "Male" 
				sSQL = sSQL + " AND LOWER(left(DT.sex,1)) = 'm'" 
			CASE "Female" 
				sSQL = sSQL + " AND LOWER(left(DT.sex,1)) = 'f'" 
		END SELECT

		sSQL = sSQL + " AND (lower(DT.div) = 'cm' or lower(DT.div) = 'cw')"

	CASE "USW"
		sSQL = sSQL + " AND SptsGrpID='USW'" 

		SELECT CASE sMembSex
			CASE "Male" 
				sSQL = sSQL + " AND LOWER(left(DT.sex,1)) = 'm'" 
			CASE "Female" 
				sSQL = sSQL + " AND LOWER(left(DT.sex,1)) = 'f'" 
		END SELECT

		IF sMembAge > 1 AND NOT IsNull(sMembAge) THEN
			sSQL = sSQL + " AND DT.Up_Age >= "&sMembAge&" AND DT.Low_Age <= "&sMembAge   
		END IF

	CASE "AKA"
		sSQL = sSQL + " AND SptsGrpID='AKA'" 

		SELECT CASE sMembSex
			CASE "Male" 
				sSQL = sSQL + " AND LOWER(left(DT.sex,1)) = 'm'" 
			CASE "Female" 
				sSQL = sSQL + " AND LOWER(left(DT.sex,1)) = 'f'" 
		END SELECT

		IF sMembAge > 1 AND NOT IsNull(sMembAge) THEN
			sSQL = sSQL + " AND DT.Up_Age >= "&sMembAge&" AND DT.Low_Age <= "&sMembAge   
		END IF

	CASE "USH"
		sSQL = sSQL + " AND SptsGrpID='USH'" 

		SELECT CASE sMembSex
			CASE "Male" 
				sSQL = sSQL + " AND LOWER(left(DT.sex,1)) = 'm'" 
			CASE "Female" 
				sSQL = sSQL + " AND LOWER(left(DT.sex,1)) = 'f'" 
		END SELECT

		IF sMembAge > 1 AND NOT IsNull(sMembAge) THEN
			sSQL = sSQL + " AND DT.Up_Age >= "&sMembAge&" AND DT.Low_Age <= "&sMembAge   
		END IF

END SELECT


SET rsDivisions=Server.CreateObject("ADODB.recordset")
sSQL = sSQL + " ORDER BY DT.div DESC"
rsDivisions.open sSQL, SConnectionToTRATable


DispSQL="N"
IF Session("AdminMenuLevel")>=50 AND DispSQL="Y" THEN
	response.write("<br>")
	response.write(sSQL)
END IF


' ---------------------------------------------------------------------------------------------
' --------------- Builds VALID Division DROP DOWN list based on criteria above  ---------------
' ---------------------------------------------------------------------------------------------



%><select name='<%=DivDropName%>' <%=DivDropStatus%> style="width:12em"><%

IF NOT rsDivisions.eof THEN 
  	rsDivisions.movefirst

  	DO WHILE NOT rsDivisions.eof
				IF TRIM(rsDivisions("Div")) = DivSelected THEN 
						%><option value="<%=rsDivisions("Div")%>" selected><%=rsDivisions("Div")%> - <%=rsDivisions("Div_Name")%></option><br><%
				ELSEIF DivSelected = "" AND rsDivisions("Up_Age") < 999 THEN 
						%><option value="<%=rsDivisions("Div")%>" selected><%=rsDivisions("Div")%> - <%=rsDivisions("Div_Name")%></option><br><%
    		ELSE 
    				%><option value="<%=rsDivisions("Div")%>"><%=rsDivisions("Div")%> - <%=rsDivisions("Div_Name")%></option><br><%
				END IF	

				rsDivisions.moveNEXT
		LOOP
ELSE
	response.write("<option value =""NA"" selected>None Available</option>")
END IF  %>
</select><%

rsDivisions.close

END SUB





'----------------------------------------------------------------------------------------------
 SUB LoadDivDropNoAgeGender (DivSelected, EventSelected, DivDropName, DivDropStatus)
'----------------------------------------------------------------------------------------------

' Loads applicable divisions into a division pulldown for each event selected

'  ---  DAVE CLARK "NOTE:" Need to add filter to filter to current SkiYear  ----

IF sTSptsGrpID="AWS" OR sTSptsGrpID="NCW" THEN
		sSQL = "SELECT DT.div, DT.div_name, DT.Up_Age FROM "&DivisionsTableName&" as DT"
ELSE
		sSQL = "SELECT DT.div, DT.div_name, DT.Up_Age FROM "&DivisionsOtherTableName&" as DT"
END IF

sSQL = sSQL + " WHERE DT.SkiYearID = (Select SkiYearID from "&SkiYearTableName&" where DefaultYear=1)"

SELECT CASE sTSptsGrpID
  	CASE "AWS"
				sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'y' AND lower(left(DT.div,1)) <> 'x'"
				sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'c'"
				sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'n'"
				sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'j'"
		CASE "NCW"
				sSQL = sSQL + " AND (lower(DT.div) = 'cm' or lower(DT.div) = 'cw')"
		CASE "USW"
				sSQL = sSQL + " AND SptsGrpID='USW'" 
		CASE "AKA"
				sSQL = sSQL + " AND SptsGrpID='AKA'" 
END SELECT

sSQL = sSQL + " ORDER BY DT.div"

SET rsDivisions=Server.CreateObject("ADODB.recordset")
rsDivisions.open sSQL, SConnectionToTRATable




' ---------------------------------------------------------------------------------------------
' --------------- Builds VALID Division DROP DOWN list based on criteria above  ---------------
' ---------------------------------------------------------------------------------------------


%><select name='<%=DivDropName%>' <%=DivDropStatus%> style="width:12em">
<option value="ALL">All Divisions</option><%

IF NOT rsDivisions.eof THEN 
  	rsDivisions.movefirst

  	DO WHILE NOT rsDivisions.eof
		IF TRIM(rsDivisions("Div")) = DivSelected THEN %>
			<option value="<%=rsDivisions("Div")%>" selected><%=rsDivisions("Div")%> - <%=rsDivisions("Div_Name")%></option><br><%
		ELSEIF DivSelected = "" AND rsDivisions("Up_Age") < 999 THEN %>
			<option value="<%=rsDivisions("Div")%>" selected><%=rsDivisions("Div")%> - <%=rsDivisions("Div_Name")%></option><br><%
    		ELSE %>
			<option value="<%=rsDivisions("Div")%>"><%=rsDivisions("Div")%> - <%=rsDivisions("Div_Name")%></option><br><%
		END IF	

		rsDivisions.moveNEXT
	LOOP
ELSE
	response.write("<option value =""None"" selected>None Available</option>")
END IF  %>
</select><%

rsDivisions.close

END SUB




' ----------------------------------------------------------------------------------------	
	SUB LoadClassEnteredDropDownByEvent (ClassDropName, tEvent, fFeeClass, ClassDropStatus)
' ----------------------------------------------------------------------------------------	

'response.write("<br> In Tools_Include16.asp - 1583")
'response.write("<br>ClassDropName ="&ClassDropName)
'response.write("<br>fFeeClass ="&fFeeClass)
'response.write("<br>")
'response.write(TRIM(fFeeClass)="")

'response.write("<br>SClassE = "&SClassE)
'response.write("<br>SClassL = "&SClassL)
'response.write("<br>SClassR = "&SClassR)

Dim TClassG_Only
Dim SClassC_Only, SClassE_Only, SClassL_Only, SClassR_Only
Dim TClassC_Only, TClassE_Only, TClassL_Only, TClassR_Only
Dim JClassC_Only, JClassE_Only, JClassL_Only, JClassR_Only

SClassG_Only="N"
SClassC_Only="N"
SClassE_Only="N"
SClassL_Only="N"
SClassR_Only="N"

TClassG_Only="N"
TClassC_Only="N"
TClassE_Only="N"
TClassL_Only="N"
TClassR_Only="N"

JClassC_Only="N"
JClassE_Only="N"
JClassL_Only="N"
JClassR_Only="N"



' --- Defaults to the class if only ONE class offered so user doesn't have to select ---
IF (Gr1AWSPulls>0 OR Gr2AWS_SPulls=0) AND SClassC=0 AND SClassE=0 AND SClassL=0 AND SClassR=0 AND SClassX=0 AND SClassCash=0 THEN SClassG_Only="Y"
IF SClassC>0 AND Gr1AWSPulls>0 AND Gr2AWS_SPulls=0 AND SClassE=0 AND SClassL=0 AND SClassR=0 AND SClassX=0 AND SClassCash=0 THEN SClassC_Only="Y"
IF SClassE>0 AND Gr1AWSPulls>0 AND Gr2AWS_SPulls=0 AND SClassC=0 AND SClassL=0 AND SClassR=0 AND SClassX=0 AND SClassCash=0 THEN SClassE_Only="Y"
IF SClassL>0 AND Gr1AWSPulls>0 AND Gr2AWS_SPulls=0 AND SClassC=0 AND SClassE=0 AND SClassR=0 AND SClassX=0 AND SClassCash=0 THEN SClassL_Only="Y"
IF SClassR>0 AND Gr1AWSPulls>0 AND Gr2AWS_SPulls=0 AND SClassC=0 AND SClassE=0 AND SClassL=0 AND SClassX=0 AND SClassCash=0 THEN SClassR_Only="Y"


IF Gr2AWTPulls>0 AND TClassC=0 AND TClassE=0 AND TClassL=0 AND TClassR=0 AND TClassX=0 AND TClassCash=0 THEN TClassG_Only="Y"	
IF TClassC>0 AND Gr1AWTPulls>0 AND TClassE=0 AND TClassL=0 AND TClassR=0 AND TClassX=0 AND TClassCash=0 THEN TClassC_Only="Y"
IF TClassE>0 AND Gr1AWTPulls>0 AND TClassC=0 AND TClassL=0 AND TClassR=0 AND TClassX=0 AND TClassCash=0 THEN TClassE_Only="Y"
IF TClassL>0 AND Gr1AWTPulls>0 AND TClassC=0 AND TClassE=0 AND TClassR=0 AND TClassX=0 AND TClassCash=0 THEN TClassL_Only="Y"
IF TClassR>0 AND Gr1AWTPulls>0 AND TClassC=0 AND TClassE=0 AND TClassL=0 AND TClassX=0 AND TClassCash=0 THEN TClassR_Only="Y"

IF JClassC>0 AND JClassE=0 AND JClassL=0 AND JClassR=0 AND JClassX=0 AND JClassCash=0 THEN JClassC_Only="Y"
IF JClassE>0 AND JClassC=0 AND JClassL=0 AND JClassR=0 AND JClassX=0 AND JClassCash=0 THEN JClassE_Only="Y"
IF JClassL>0 AND JClassC=0 AND JClassE=0 AND JClassR=0 AND JClassX=0 AND JClassCash=0 THEN JClassL_Only="Y"
IF JClassR>0 AND JClassC=0 AND JClassE=0 AND JClassL=0 AND JClassX=0 AND JClassCash=0 THEN JClassR_Only="Y"

'response.write("<br>TClassC_Only = "&TClassC_Only)

'IF sMemberID = "000001151" AND tEvent="S" THEN
IF sMemberID = "000001151"  THEN
		'response.write("<br>TI fCl = "&TRIM(fFeeClass))
		'response.write("<br>tEvent = "&tEvent)
		'response.write(" TEST - ")
		'TestValue = "R"
		'response.write(LEN(ClassDropName))
		'response.write(fFeeClass)
		' response.write(eval(""&ClassDropName&""))
		
		' response.write(" - FC = "&fFeeClass)
END IF
	
%>
<select name='<%=ClassDropName%>' style="width:8em" <%=ClassDropStatus%>>
<%


		' --- Slalom for AWS, ABC and NCW ---
		IF tEvent="S" OR tEvent="3G" THEN
				' --- Provide select option only if multiple classes are offered ---
				IF SClassC_Only="N" AND SClassE_Only="N" AND SClassL_Only="N" AND SClassR_Only="N" THEN
							%><option value="">Select</option><%
				END IF		

				IF (Gr1AWSPulls>0 OR Gr2AWS_SPulls>0) AND (TRIM(fFeeClass)="G" OR SClassG_Only="Y") THEN 
						%><option value="G" selected>Grassroots</option><br><%
				ELSEIF Gr1AWSPulls>0 OR Gr2AWS_SPulls>0 THEN 
						%><option value="G">Grassroots</option><br><%
				END IF
				IF (SClassC>0 OR BSClassC>0 OR USClassC>0) AND (TRIM(fFeeClass)="C" OR SClassC_Only="Y") THEN 
						%><option value="C" selected>Class C</option><br><%
				ELSEIF (SClassC>0 OR BSClassC>0 OR USClassC) THEN 
						%><option value="C">Class C</option><br><%
				END IF
				IF (SClassE>0) AND (TRIM(fFeeClass)="E" OR SClassE_Only="Y") THEN 
						%><option value="E" selected>Class E</option><br><%
				ELSEIF (SClassE>0) THEN 
						%><option value="E">Class E</option><br><%
				END IF
				IF (SClassL>0 OR BSClassL>0) AND (TRIM(fFeeClass)="L" OR SClassL_Only="Y") THEN 
						%><option value="L" selected>Class L</option><br><%
				ELSEIF (SClassL>0 OR BSClassL>0 OR SClassL_Only="Y") THEN 
						%><option value="L">Class L</option><br><%
				END IF

IF sMemberID = "000001151" AND 2=1 THEN
				IF (SClassR>0 OR BSClassR>0) AND (ClassDropName="R" OR SClassR_Only="Y") THEN 
						%><option value="R" selected>Class R</option><br><%
				ELSEIF (SClassR>0 OR BSClassR>0 OR SClassR_Only="Y") THEN 
						%><option value="R">Class R</option><br><%
				END IF	
ELSE
				IF (SClassR>0 OR BSClassR>0) AND (TRIM(fFeeClass)="R" OR SClassR_Only="Y") THEN 
						%><option value="R" selected>Class R</option><br><%
				ELSEIF (SClassR>0 OR BSClassR>0 OR SClassR_Only="Y") THEN 
						%><option value="R">Class R</option><br><%
				END IF	
END IF
				IF SClassX>0 AND TRIM(fFeeClass)="X" THEN 
						%><option value="X" selected>Class X</option><br><%
				ELSEIF SClassX>0 THEN 
						%><option value="X">Class X</option><br><%
				END IF	
				IF SClassCash>0 AND TRIM(fFeeClass)="$" THEN 
						%><option value="$" selected>Class $</option><br><%
				ELSEIF SClassCash>0 THEN 
						%><option value="$">Class $</option><br><%
				END IF	
				
		END IF


		' --- Trick for AWS, ABC and NCW ---
		IF tEvent="T" THEN
				' --- Provide select option only if multiple classes are offered ---
				IF TClassC_Only="N" AND TClassE_Only="N" AND TClassL_Only="N" AND TClassR_Only="N" THEN
							%><option value="">Select</option><%
				END IF		
				IF (Gr1AWSPulls>0 OR Gr2AWS_TPulls>0) AND (TRIM(fFeeClass)="G" OR TClassG_Only="Y") THEN 
						%><option value="G" selected>Grassroots</option><br><%
				ELSEIF Gr1AWSPulls>0 OR Gr2AWS_TPulls>0 THEN 
						%><option value="G">Grassroots</option><br><%
				END IF

				IF (TClassC>0 OR BTClassC>0 OR UTClassC>0) AND (TRIM(fFeeClass)="C" OR TClassC_Only="Y") THEN 
								%><option value="C" selected>Class C</option><br><%
				ELSEIF (TClassC>0 OR BTClassC>0 OR UTClassC) THEN 
						%><option value="C">Class C</option><br><%
				END IF
				IF (TClassE>0) AND (TRIM(fFeeClass)="E" OR TClassE_Only="Y") THEN 
						%><option value="E" selected>Class E</option><br><%
				ELSEIF TClassE>0 THEN 
						%><option value="E">Class E</option><br><%
				END IF
				IF (TClassL>0 OR BTClassL>0) AND (TRIM(fFeeClass)="L" OR TClassL_Only="Y") THEN 
						%><option value="L" selected>Class L</option><br><%
				ELSEIF (TClassL>0 OR BTClassL>0) THEN 
						%><option value="L">Class L</option><br><%
				END IF
				IF (TClassR>0 OR BTClassR>0) AND (TRIM(fFeeClass)="R" OR TClassR_Only="Y") THEN 
						%><option value="R" selected>Class R</option><br><%
				ELSEIF (TClassR>0 OR BTClassR>0) THEN 
						%><option value="R">Class R</option><br><%
				END IF	
				IF TClassX>0 AND TRIM(fFeeClass)="X" THEN 
						%><option value="X" selected>Class X</option><br><%
				ELSEIF TClassX>0 THEN 
						%><option value="X">Class X</option><br><%
				END IF	
				IF TClassCash>0 AND TRIM(fFeeClass)="$" THEN 
						%><option value="$" selected>Class $</option><br><%
				ELSEIF TClassCash>0 THEN 
						%><option value="$">Class $</option><br><%
				END IF	

		END IF

		' --- Jump for AWS, ABC and NCW ---
		IF tEvent="J" THEN
				' --- Provide select option only if multiple classes are offered ---
				IF JClassC_Only="N" AND JClassE_Only="N" AND JClassL_Only="N" AND JClassR_Only="N" THEN
							%><option value="">Select</option><%
				END IF		

				IF (JClassC>0 OR BJClassC>0 OR UJClassC>0) AND (TRIM(fFeeClass)="C" OR JClassC_Only="Y") THEN 
						%><option value="C" selected>Class C</option><br><%
				ELSEIF (JClassC>0 OR BJClassC>0 OR UJClassC) THEN 
						%><option value="C">Class C</option><br><%
				END IF
				IF (JClassE>0) AND (TRIM(fFeeClass)="E" OR JClassE_Only="Y") THEN 
						%><option value="E" selected>Class E</option><br><%
				ELSEIF JClassE>0 THEN 
						%><option value="E">Class E</option><br><%
				END IF
				IF (JClassL>0 OR BJClassL>0) AND (TRIM(fFeeClass)="L" OR JClassL_Only="Y") THEN 
						%><option value="L" selected>Class L</option><br><%
				ELSEIF (JClassL>0 OR BJClassL>0) THEN 
						%><option value="L">Class L</option><br><%
				END IF
				IF (JClassR>0 OR BJClassR>0) AND (TRIM(fFeeClass)="R" OR JClassR_Only="Y") THEN 
						%><option value="R" selected>Class R</option><br><%
				ELSEIF (JClassR>0 OR BJClassR>0) THEN 
						%><option value="R">Class R</option><br><%
				END IF	
				IF JClassN>0 AND TRIM(fFeeClass)="N" THEN 
						%><option value="N" selected>Class N</option><br><%
				ELSEIF JClassN>0 THEN 
						%><option value="N">Class N</option><br><%
				END IF
				IF JClassX>0 AND TRIM(fFeeClass)="X" THEN 
						%><option value="X" selected>Class X</option><br><%
				ELSEIF JClassX>0 THEN 
						%><option value="X">Class X</option><br><%
				END IF	
				IF JClassCash>0 AND TRIM(fFeeClass)="$" THEN 
						%><option value="$" selected>Class $</option><br><%
				ELSEIF JClassCash>0 THEN 
						%><option value="$">Class $</option><br><%
				END IF	

		END IF


		' --- BAREFOOT ---
		IF tEvent="KS" THEN
				IF KSClassQ>0 AND TRIM(fFeeClass)="Q" THEN 
						%><option value="Q" selected>Class Q</option><br><%
				ELSEIF KSClassQ>0 THEN 
						%><option value="Q">Class Q</option><br><%
				END IF
				IF KSClassT>0 AND TRIM(fFeeClass)="T" THEN 
						%><option value="T" selected>Class T</option><br><%
				ELSEIF KSClassT>0 THEN 
						%><option value="T">Class T</option><br><%
				END IF
		END IF

		IF tEvent="KT" THEN
				IF KSClassQ>0 AND TRIM(fFeeClass)="Q" THEN 
						%><option value="Q" selected>Class Q</option><br><%
				ELSEIF KSClassQ>0 THEN 
						%><option value="Q">Class Q</option><br><%
				END IF
				IF KSClassT>0 AND TRIM(fFeeClass)="T" THEN 
						%><option value="T" selected>Class T</option><br><%
				ELSEIF KSClassT>0 THEN 
						%><option value="T">Class T</option><br><%
				END IF
		END IF

		IF tEvent="KF" THEN
				IF KFlClassQ>0 AND TRIM(fFeeClass)="Q" THEN 
						%><option value="Q" selected>Class Q</option><br><%
				ELSEIF KFlClassQ>0 THEN 
						%><option value="Q">Class Q</option><br><%
				END IF
				IF KFlClassT>0 AND TRIM(fFeeClass)="T" THEN 
						%><option value="T" selected>Class T</option><br><%
				ELSEIF KFlClassT>0 THEN 
						%><option value="T">Class T</option><br><%
				END IF
		END IF
		IF tEvent="KR" THEN
				IF KFrClassQ>0 AND TRIM(fFeeClass)="Q" THEN 
						%><option value="Q" selected>Class Q</option><br><%
				ELSEIF KFrClassQ>0 THEN 
						%><option value="Q">Class Q</option><br><%
				END IF
				IF KFrClassT>0 AND TRIM(fFeeClass)="T" THEN 
						%><option value="T" selected>Class T</option><br><%
				ELSEIF KFrClassT>0 THEN 
						%><option value="T">Class T</option><br><%
				END IF
		END IF


		' --- WAKEBOARD ---
		IF tEvent="WB" THEN
				IF WWakeW>0 AND TRIM(fFeeClass)="C" THEN 
						%><option value="C" selected>Class C</option><br><%
				ELSEIF WWakeW>0 THEN 
						%><option value="C">Class C</option><br><%
				END IF
		END IF
		IF tEvent="WS" THEN
				IF WSkateW>0 AND TRIM(fFeeClass)="C" THEN 
						%><option value="C" selected>Class C</option><br><%
				ELSEIF WSkateW>0 THEN 
						%><option value="C">Class C</option><br><%
				END IF
		END IF
		IF tEvent="WU" THEN
				IF WSurfW>0 AND TRIM(fFeeClass)="C" THEN 
						%><option value="C" selected>Class C</option><br><%
				ELSEIF WSurfW>0 THEN 
						%><option value="C">Class C</option><br><%
				END IF
		END IF


		' --- HYDROFOIL ---
		IF tEvent="HF" THEN
				IF (Gr2USH_FreeRidePulls>0 OR Gr1USHPulls>0) AND TRIM(fFeeClass)="G" THEN 
						%><option value="G" selected>Class G</option><br><%
				ELSEIF (Gr2USH_FreeRidePulls>0 OR Gr1USHPulls>0) THEN 
						%><option value="G">Class G</option><br><%
				END IF
		END IF
		IF tEvent="HJ" THEN
				IF Gr2USH_JumpOutPulls>0 AND TRIM(fFeeClass)="G" THEN 
						%><option value="G" selected>Class G</option><br><%
				ELSEIF Gr2USH_JumpOutPulls>0 THEN 
						%><option value="G">Class G</option><br><%
				END IF
		END IF
		IF tEvent="HB" THEN
				IF Gr2USH_BigAirPulls>0 AND TRIM(fFeeClass)="G" THEN 
						%><option value="G" selected>Class G</option><br><%
				ELSEIF Gr2USH_BigAirPulls>0 THEN 
						%><option value="G">Class G</option><br><%
				END IF
		END IF
		IF tEvent="H3" THEN
				IF Gr2USH_3TricksPulls>0 AND TRIM(fFeeClass)="G" THEN 
						%><option value="G" selected>Class G</option><br><%
				ELSEIF Gr2USH_3TricksPulls>0 THEN 
						%><option value="G">Class G</option><br><%
				END IF
		END IF



%></select><%


END SUB










' ---------------------
   SUB DefineTRAStyles 
' ---------------------

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Welcome to USA Water Ski</title>
<link rel="stylesheet" type="text/css" href="/css/styles.css" />
<script language="javascript" type="text/JavaScript" src="/jscripts/scripts.js"></script>
<script language="javascript" type="text/javascript" src="/jscripts/swfobject.js"></script>
<style type="text/css">

/* this style applies to the DropTable table */
table.droptable {padding:2px; background-position: center right; border:3px solid <%=HQSiteColor2%>}
/* this style applies to all th (cells) within the 'scores' table */ 
table.droptable th {padding:1px; border:0px solid <%=HQSiteColor3%>;} 
/* this style applies to all td (cells) within the 'scores' table */ 
table.droptable td {padding:1px; vertical-align:middle; border:0px solid <%=HQSiteColor2%>;} 
/* this style applies to the scores table */

/* this style applies to the BlankTable table */
table.blanktable {padding:2px; background-position: center right;}
/* this style applies to all th (cells) within the 'blank' table */ 
table.blanktable th {padding:1px; border:0px solid <%=HQSiteColor3%>;} 
/* this style applies to all td (cells) within the 'blank' table */ 
table.blanktable td {padding:1px; vertical-align:middle; border:0px solid <%=HQSiteColor2%>; word-wrap:break-word;} 
/* this style applies to the blank table */

/*
/* this style applies to the Scores table */
table.scores {padding:0px; border:3px solid <%=HQSiteColor2%>; border-collapse:collapse;}
/* this style applies to all th (cells) within the 'scores' table */ 
table.scores th {padding:0px; background-color:<%=HQSiteColor2%>; border:1px solid black; vertical-align:bottom;} 
/* this style applies to all td (cells) within the 'scores' table */ 
table.scores td {padding:0px; border:1px solid <%=HQSiteColor2%>; border-style:solid; background-color:<%=TableColor1%>; vertical-align:middle;} 

/*
/* this style applies to the SpaceTable table */
table.spacetable {padding:2px; border:1px solid <%=HQSiteColor2%>}
/* this style applies to all th (cells) within the 'spacetable' table */ 
table.spacetable th {padding:3px; border:1px solid black; background-color:<%=TableColor1%>; vertical-align:bottom;} 
/* this style applies to all td (cells) within the 'spacetable' table */ 
table.spacetable td {padding:6px; border:1px solid black; background-color:<%=TableColor1%>; vertical-align:middle;} 

/*
/* this style applies to the innertable table */
table.innertable {padding:0px; border:1px solid <%=HQSiteColor2%>; border-collapse:collapse;}
/* this style applies to all th (cells) within the 'innertable' table */ 
table.innertable th {padding:1px; border:1px solid <%=HQSiteColor1%>; border-style:solid; background-color:<%=HQSiteColor2%>; vertical-align:bottom;} 
/* this style applies to all td (cells) within the 'innertable' table */ 
table.innertable td {padding:3px; border:1px solid <%=HQSiteColor2%>; border-style:solid; background-color:<%=TableColor1%>; vertical-align:middle;  word-wrap:break-word;} 

/*
/* this style applies to the messagetable table */
table.messagetable {padding:0px; border:1px solid <%=HQSiteColor2%>; border-collapse:collapse;}
/* this style applies to all th (cells) within the 'messagetable' table */ 
table.messagetable th {padding:1px; border:1px solid <%=HQSiteColor1%>; border-style:solid; background-color:<%=HQSiteColor2%>; vertical-align:bottom;} 
/* this style applies to all td (cells) within the 'messagetable' table */ 
table.messagetable td {padding-left:15px; padding-right:15px; padding-top:8px; padding-bottom:8px; border:0px solid <%=HQSiteColor2%>; border-style:solid; background-color:<%=TableColor1%>; vertical-align:middle; white-space:nowrap;} 

/*
/* this style applies to the noborder table */
table.noborder {padding:0px; border:0px solid <%=HQSiteColor2%>; border-collapse:collapse;}
/* this style applies to all th (cells) within the 'noborder' table */ 
table.noborder th {padding:1px; border:0px solid <%=HQSiteColor1%>; border-style:solid; background-color:white; vertical-align:bottom;} 
/* this style applies to all td (cells) within the 'noborder' table */ 
table.noborder td {padding:1px; text-align: left; border:0px solid <%=HQSiteColor2%>; border-style:solid; background-color:white; vertical-align:middle; word-wrap:break-word;} 

/*
/* this style applies to the tourlist table */
table.tourlist {padding:0px; border:1px solid <%=HQSiteColor2%>; border-collapse:collapse; overflow-scroll}
/* this style applies to all th (cells) within the 'tourlist' table */ 
table.tourlist th {padding:1px; border:1px solid <%=HQSiteColor1%>; border-style:solid; background-color:<%=HQSiteColor2%>; vertical-align:bottom;} 
/* this style applies to all td (cells) within the 'tourlist' table */ 
table.tourlist td {padding:3px; border:1px solid <%=HQSiteColor2%>; border-style:solid; background-color:<%=TableColor1%>; vertical-align:middle;  word-wrap:break-word;} 

/*
/* this style applies to the mobiletable1 table */
table.mobiletable1 {padding:0px; border:0px solid <%=TableColor2%>; border-collapse:collapse;}
/* this style applies to all th (cells) within the 'innertable' table */ 
table.mobiletable1 th {padding:0px; border:0px solid <%=TableColor1%>; border-style:solid; background-color:<%=TableColor2%>; vertical-align:bottom;} 
/* this style applies to all td (cells) within the 'innertable' table */ 
table.mobiletable1 td {padding:0px; border:0px solid <%=TableColor2%>; border-style:solid; background-color:<%=TableColor1%>; vertical-align:middle;  word-wrap:break-word;} 

</style>
</head><%

END SUB


SUB Test
%>
<link rel="stylesheet" type="text/css" href="/css/styles.css" />
<%
END SUB




' ----------------------------
   SUB DefineTRAStyles_mobile 
' ----------------------------

'response.write("<br>TableColor2="&TableColor1)
'response.write("<br>TableColor2="&TableColor2)
%>

<script language="javascript" type="text/JavaScript" src="/jscripts/scripts.js"></script>
<script language="javascript" type="text/javascript" src="/jscripts/swfobject.js"></script>
<style type="text/css">
/*
/* this style applies to the mobiletable1 table */
table.mobiletable1 {padding:0px; border:0px solid <%=HQSiteColor2%>; border-collapse:collapse;}
/* this style applies to all th (cells) within the 'innertable' table */ 
table.mobiletable1 th {padding:0px; border:0px solid <%=HQSiteColor1%>; border-style:solid; background-color:<%=HQSiteColor2%>; vertical-align:bottom;} 
/* this style applies to all td (cells) within the 'innertable' table */ 
table.mobiletable1 td {padding:0px; border:0px solid <%=HQSiteColor2%>; border-style:solid; background-color:<%=TableColor3%>; vertical-align:middle;  word-wrap:break-word;} 

body {background-color: #FFFFFF;}


input.textbox_hidden_banner
{
	padding-bottom:0px; 
	vertical-align:bottom; 
	text-align:center; 
	background-color:<%=TableColor2%>; 
	border-width:0px; 
	color:<%=HeaderTextColor1%>; 
	width:18em;
	font-size:16px; 
	font-weight:bold;
}

</style>
<%

END SUB


' --------------
   SUB HQHead1 
' --------------

DefineTRAStyles


%>
<body onload="MM_preloadImages('/images/interior/img_06_f2.jpg','/images/interior/img_08_f2.jpg','/images/interior/img_10_f2.jpg','/images/interior/img_12_f2.jpg','/images/interior/img_14_f2.jpg','/images/interior/img_16_f2.jpg','/images/interior/img_18_f2.jpg','/images/interior/img_20_f2.jpg','/images/interior/img_22_f2.jpg')">
<table cellspacing="0" Class="layout">
  <tr>
    <td><img src="/images/img_01.jpg" alt="" width="1014" height="35" /></td>
  </tr>
  <tr>
    <td>
    	<table width="1014" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td Class="logo" width="241"><img src="/images/img_02.jpg" alt="USA Water Ski" width="241" height="108" /></td>
          <td Class="top">
          	<table width="773" height="108" cellpadding="0" cellspacing="0">
        			<tr>
          			<td width="62"></td>
        					<td align="center" width="468" Class="top_ad_space">
        						<div id="flashcontent" style="width:100%;text-align:center;"></div>
          					<!--#include virtual="/inc/banners.asp" -->
						<layer id="placeholderlayer"></layer><div id="placeholderdiv"></div></td>
									</td>
          				<td Class="top_search"><!--#include virtual="/inc/search_form.asp" --></td>
        				</tr>
						</table>
					</td>
        </tr>
    </table>
    </td>
  </tr>
  <!--#include virtual="/inc/icon_nav.asp" -->
  <tr>
    <td><img src="/images/interior/img_24.jpg" alt="Having FUN Today...Building CHAMPIONS For Tomorrow" width="1014" height="18" /></td>
  </tr>
  <%
  Dim tyu
  tyu=1
  IF tyu=2 THEN
  %>
  <tr>
    <td><img src="/images/img_26.jpg" alt="" width="1014" height="24" /></td>
  </tr>
  <%
	END IF
  %>
  <tr>
    <td style="border-style:solid; border-width:0px;">
    	<table width="100%" border="0" cellspacing="0" cellpadding="0">
      	<tr>
        	<td width="245px">
	  				<table width=100% border="0" cellspacing="0" cellpadding="0">
            	<tr>
              	<td>&nbsp;</td>
              	<%
		
								IF Session("adminmenulevel")>1 THEN 
										%>
										<td Class="sidebar"><!--#include virtual="/rankings/menu_admin.asp" --></td>
										<%
								ELSE 
										%>
										<td width=98% style="border-style:solid; border-width:0px;" Class="sidebar"><!--#include virtual="/inc/nav.asp" --></td>
										<%
								END IF 
									
								%>
            	</tr>
            	<tr>
              	<td>&nbsp;</td>
              	<td><img src="/images/img_31.jpg" alt="" width="234" height="71" /></td>
            	</tr>
	    				
	    				<tr>
              	<td>&nbsp;</td>
	      				<td>
	      					<%
									' --- Displays all the logos in the appropriate folder ---
									DisplaySponsorImages  
									%>
	      				</td> 	
	    				</tr>			
          	</table>
					</td>
        	<td>
	  			
	  			<table width="100%" border="0" cellpadding="0" cellspacing="0" Class="contentcontainer" style="padding:10px; margin:2px;">
            <tr>
              <td Class="content_4" style="padding:0px">
							<%

				' --- Done with Header and Menus - now display CONTENT ---		

END SUB


' ------------------
   SUB HQHeadNoMenu 
' ------------------

DefineTRAStyles

%>
<body onload="MM_preloadImages('/images/interior/img_06_f2.jpg','/images/interior/img_08_f2.jpg','/images/interior/img_10_f2.jpg','/images/interior/img_12_f2.jpg','/images/interior/img_14_f2.jpg','/images/interior/img_16_f2.jpg','/images/interior/img_18_f2.jpg','/images/interior/img_20_f2.jpg','/images/interior/img_22_f2.jpg')">
<table cellspacing="0" Class="layout">
  <tr>
    <td><img src="/images/img_01.jpg" alt="" width="1014" height="35" /></td>
  </tr>
  <tr>
    <td><table width="1014" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td Class="logo" width="241"><img src="/images/img_02.jpg" alt="USA Water Ski" width="241" height="108" /></td>
          <td Class="top"><table width="773" height="108" cellpadding="0" cellspacing="0">
              <tr>
                <td width="62"></td>
                <td align="center" width="468" Class="top_ad_space"><div id="flashcontent" style="width:100%;text-align:center;"></div>
                    <script type="text/javascript">
				// <![CDATA[
				var fo = new FlashObject("/banner/USAWaterski_logoBanner.swf", "", "468", "60", 6, "#003");
				fo.addParam("salign", "t");
				fo.write("flashcontent");
				// ]]>
				</script></td>
                <td Class="top_search"><!--#include virtual="/inc/search_form.asp" --></td>
              </tr>
          </table></td>
        </tr>
    </table>
    </td>
  </tr>

  <!--#include virtual="/inc/icon_nav.asp" -->

  <tr>
    <td><img src="/images/interior/img_24.jpg" alt="Having FUN Today...Building CHAMPIONS For Tomorrow" width="1014" height="18" /></td>
  </tr>
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>
	  <table width="100%" border="0" cellpadding="0" cellspacing="0" Class="contentcontainer">
          <tr>
            <td Class="content_4" style="padding:0px">
		
<%

END SUB




' ------------------
   SUB HQHeadNoButtons 
' ------------------

DefineTRAStyles

%>
<body onload="MM_preloadImages('/images/interior/img_06_f2.jpg','/images/interior/img_08_f2.jpg','/images/interior/img_10_f2.jpg','/images/interior/img_12_f2.jpg','/images/interior/img_14_f2.jpg','/images/interior/img_16_f2.jpg','/images/interior/img_18_f2.jpg','/images/interior/img_20_f2.jpg','/images/interior/img_22_f2.jpg')">
<table cellspacing="0" Class="layout">
  <tr>
    <td><img src="/images/img_01.jpg" alt="" width="1014" height="35" /></td>
  </tr>
  <tr>
    <td><table width="1014" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td Class="logo" width="241"><img src="/images/img_02.jpg" alt="USA Water Ski" width="241" height="108" /></td>
          <td Class="top"><table width="773" height="108" cellpadding="0" cellspacing="0">
              <tr>
                <td width="62"></td>
                <td align="center" width="468" Class="top_ad_space"><div id="flashcontent" style="width:100%;text-align:center;"></div>
                    <script type="text/javascript">
				// <![CDATA[
				var fo = new FlashObject("/banner/USAWaterski_logoBanner.swf", "", "468", "60", 6, "#003");
				fo.addParam("salign", "t");
				fo.write("flashcontent");
				// ]]>
				</script></td>
                <td Class="top_search"><!--#include virtual="/inc/search_form.asp" --></td>
              </tr>
          </table></td>
        </tr>
    </table>
    </td>
  </tr>
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td>
	  <table width="100%" border="0" cellpadding="0" cellspacing="0" Class="contentcontainer" style="padding:2px;margin:2px;">
          <tr>
            <td Class="content_4" style="padding:0px">
		
<%



END SUB







' ------------------
   SUB HQFooter1
' ------------------

%>
            </td>
	  </tr>
        </table>
       </td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><!--#include virtual="/inc/footer.asp" --></td>
  </tr>
</table>
</body>
</html>
<%

END SUB





' -------------------------------------------------------------------
  SUB AutoLoadPulldown (CurrentValue, MinValue, MaxValue, StepValue)
' -------------------------------------------------------------------

Dim iCounter


response.write("<option value = 0 >All</option>")

FOR iCounter = MinValue TO MaxValue STEP StepValue
	IF iCounter = CurrentValue THEN
		response.write("<option value = """&iCounter&""" SELECTED>"&iCounter&"</option>")
	ELSE
		response.write("<option value = """&iCounter&""">"&iCounter&"</option>")
	END IF
NEXT

END SUB


' -------------------------------------------------------
  SUB LoadMonthsPulldown (WhatEndofRange, CurrentValue)
' -------------------------------------------------------

' --- Modified 7/14/2016 to test for NOT Numeric and changed to select CurrentValue=1 --
IF NOT(IsNumeric(CurrentValue)) OR TRIM(CurrentValue)="" OR TRIM(CurrentValue)="=" OR TRIM(CurrentValue)="==" THEN CurrentValue=0

MonthArray = Split(MonthList,",")  %>
<select name="<%=WhatEndofRange%>"><%

  response.write("<option value = 0 >N/A</option>")
  FOR kvar = 1 TO UBOUND(MonthArray)

    ' --- Values are 1 thru 12, but display is JAN, FEB, etc. ---
    IF cdbl(kvar)=cdbl(CurrentValue) THEN
				response.write("<option value = """&kvar&""" SELECTED>"&MonthArray(kvar)&"</option>")
    ELSE
				response.write("<option value = """&kvar&""">"&MonthArray(kvar)&"</option>")
    END IF
  NEXT  %>
</select><%

END SUB


' ---------------------------------------------------------------------
  SUB LoadMonthsPulldown_New (WhatEndofRange, CurrentValue, ThisStyle)
' ---------------------------------------------------------------------

IF TRIM(CurrentValue)="" OR TRIM(CurrentValue)="=" OR TRIM(CurrentValue)="==" THEN CurrentValue=0

MonthArray = Split(MonthList,",")  %>
<select name="<%=WhatEndofRange%>" style="<%=ThisStyle%>;"><%

  response.write("<option value = 0 >N/A</option>")
  FOR kvar = 1 TO UBOUND(MonthArray)

    ' --- Values are 1 thru 12, but display is JAN, FEB, etc. ---
    IF cdbl(kvar)=cdbl(CurrentValue) THEN
	response.write("<option value = """&kvar&""" SELECTED>"&MonthArray(kvar)&"</option>")
    ELSE
	response.write("<option value = """&kvar&""">"&MonthArray(kvar)&"</option>")
    END IF
  NEXT  %>
</select><%

END SUB


%>





