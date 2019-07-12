<%


Dim sTournAppID, WhatCalendarYear, sTSanction

' --- Read from SanctionTableName and TRegSetupTableName --- 
Dim sTSptsGrpID, sTRegSetUpStatus
Dim sTourName, sTourCity, sTourState, sTDateS, sTDateE, sSptsGrpID
Dim sTEntryFee1, sTEntryFee2, sTEntryFee3, sTEntryFeeFamily, sOtherFee, sTLateFee, sTLateDate, sTLateDate_Adjusted, sTLFPerDay
Dim sTourEmail, sTDirEmail, sTsemail, GTComments, GTSofE

' --- Added 5/23/2015 for filtering parameters --- 
Dim sTourRegion, sClass	


Dim sTRegistrarPhone, sTRegistrarEmail
Dim sTRegistrarName, sTRegistrarCity, sTRegistrarState, sTRegistrarAddr, sTRegistrarZip  


Dim sGTAccommodation, sGTAwards, sGTPractice, sGTStartTime, sGTSofE, sGTComments, sG_IWWF_req
Dim sTDirName, sTSponsor, sTSite, sTSiteID, sTOpenClosed, sTEntryLimit, sTStatus, sTDeleted
Dim sCJudge, sCScorer, sCDriver, sCSafety, sAnnouncer, sTechCont, sPanAmJudge, sOOJ
Dim sAp1Judge, sAp2Judge, sAp3Judge, sAp4Judge, sAp5Judge


' --- Question about using these ???
'Dim sTO1ApScore, sTOCScore, sTOCDriver, sTOAnnounce, sTOTechCont, sTOPanAmJudge


'--- This is a legacy field which JMeis continues to populate for those not wanting to use OLR settings ---
Dim sTEntryFees


Dim sOffDiscPerc, sJrDiscPerc, sSrDiscPerc, sClubDiscPerc, sTourClubCode, sDiscMeth, sDiscNote
Dim sBTickCost, sBTickWithE, sAWSEFDon_OK


' **** Event offered information ***

' --- AWS legacy 2009 fields ---
Dim sTEventJump, sTEventSlalom, sTEventTrick
Dim sTHJClassR, sTHJClassL, sTHJClassE, sTHJClassC, sTHJClassN, sTHJClassF
Dim sTHSClassR, sTHSClassL, sTHSClassE, sTHSClassC, sTHSClassN, sTHSClassF
Dim sTHTClassR, sTHTClassL, sTHTClassE, sTHTClassC, sTHTClassN, sTHTClassF
Dim sTRoundsS, sTRoundsT, sTRoundsJ


' --- New AWS fields for 2010
Dim SClassC, SClassE, SClassL, SClassR, SClassCash, SClassX
Dim TClassC, TClassE, TClassL, TClassR, TClassCash, TClassX
Dim JClassC, JClassE, JClassL, JClassR, JClassCash, JClassX
Dim JClassN 	' --- This is used for Grassroots jump since F cannot hold Jump ---

' --- BAREFOOT ---
Dim BSClassC, BSClassL, BSClassR 
Dim BTClassC, BTClassL, BTClassR
Dim BJClassC, BJClassL, BJClassR

' --- Legacy 2009 class F fields for all SD's
Dim sTEventFKB, sTEventFDA, sTEventFW, sTEventFB, sTEventFHF, sTEventF3ev

Dim SLPremierCnt, TRPremierCnt, JUPremierCnt


' ********  WAKEBOARD  *********
Dim sTEventWake, sTEventWSkate, sTEventWSurf		' --- Legacy 2009 fields
Dim WWakeW, WSkateW, WSurfW

' ********  KNEEBOARD  *********
' --- Legacy 2009 fields
Dim sKEventFlip, sKEventFree
Dim sKFlipClassT, sKFlipClassQ, sKFreeClassT, sKFreeClassQ, sKSlalomClassT, sKSlalomClassQ, sKTrickClassT, sKTrickClassQ
' --- New fields for 2010
Dim KSClassQ, KSClassT, KTClassQ, KTClassT, KFlClassQ, KFlClassT, KFrClassQ, KFrClassT


' ********  GRASSROOTS  *********
Dim Gr1AWSPulls, Gr1ABCPulls, Gr1USWPulls, Gr1AKAPulls, Gr1USHPulls, Gr1WSDPulls
Dim Gr2ABC_SPulls, Gr2ABC_TPulls
Dim Gr2AWS_SPulls, Gr2AWS_TPulls
Dim Gr2USH_FreeRidePulls, Gr2USH_JumpOutPulls, Gr2USH_BigAirPulls, Gr2USH_3TrickPulls
Dim Gr2USW_WPulls, Gr2USW_SkatePulls, Gr2USW_SurfPulls, Gr2USW_RailJamPulls
Dim Gr2AKA_SPulls, Gr2AKA_TPulls, Gr2AKA_FreePulls, Gr2AKA_FlipPulls

Dim sGREntryFeeIncluded, sGRDiscount
Dim sGREntryFee1, sGREntryFee2, sGREntryFee3, sGREntryFee4, sGREntryFee5
'Dim GRFee_1, GRFee_2, GRFee_3 
Dim sGrassroots, sGRFunDay, sGRTournament, sGRBoat, sGRCable



Dim sTDescription, sFDescription, sWDescription, sKDescription, sCDescription
Dim sTDvOffered, sGTSDirections


' --- From TSchedule and Registration ---
Dim sCMulti, sEMulti, sLMulti, sRMulti, sCSurchg, sESurchg, sLSurchg, sRSurchg
Dim sClinFeeJD, sClinFeeAD, sMixedOptions, sMaxPulls, sReservedPulls, sReservedPullsCode, sTPandC, sTPandCPulls, sPullsReceived
Dim sForm1Name, sForm2Name, sForm3Name, sForm4Name, sForm5Name, sForm6Name, sQualLevel 
Dim sTotalPreviousPayments, MaxDisc, sReceiveEmail
Dim sHQAccount

Dim SpecialWaiver_IncludeHQ





' --- Fields describe optional charges ---
' --- New for 2009
Dim OFItem, TotNumOptItems 
Dim sOF1Desc, sOF2Desc, sOF3Desc, sOF4Desc, sOF5Desc, sOF6Desc, sOF7Desc, sOF8Desc, sOF9Desc, sOF10Desc
Dim sOF1Amt, sOF2Amt, sOF3Amt, sOF4Amt, sOF5Amt, sOF6Amt, sOF7Amt, sOF8Amt, sOF9Amt, sOF10Amt
Dim sOF1MaxQty, sOF2MaxQty, sOF3MaxQty, sOF4MaxQty, sOF5MaxQty, sOF6MaxQty, sOF7MaxQty, sOF8MaxQty, sOF9MaxQty, sOF10MaxQty
Dim sOF1Required, sOF2Required, sOF3Required, sOF4Required, sOF5Required, sOF6Required, sOF7Required, sOF8Required, sOF9Required, sOF10Required
Dim sOF1Qty, sOF2Qty, sOF3Qty, sOF4Qty, sOF5Qty, sOF6Qty, sOF7Qty, sOF8Qty, sOF9Qty, sOF10Qty  ' --- Defines USER Qty ---
Dim sOF1Fee, sOF2Fee, sOF3Fee, sOF4Fee, sOF5Fee, sOF6Fee, sOF7Fee, sOF8Fee, sOF9Fee, sOF10Fee  ' --- Defines USER Total $$$ for this item ---







' --- Headings on display ---
Dim sTGRClassText, sTBaseClassText, sTUpgradeClassText
Dim ShowGRHead, ShowStdHead, ShowRecHead
Dim sReachedReceiptPage

Dim sTEntryFeeFamExtra, sMaxFamMembers, MembList(10), MembListName(10), TotQualifyingFamMemb 
Dim sMaxSlPulls, sMaxTrPulls, sMaxJuPulls
Dim sOLRDisplayStatus, sOLR_Pd, sUseOLReg

' --- Defined for current state of the FORM ---
Dim sTEvent(10), sTEventName(10), sDiv(10), sFeeClass(10), sFeeRounds(10), sQfyOverride(10), QualEvent(10)
Dim sSelectEvent(10), sShowGR(10), sShowStd(10), sShowRec(10), sShowCash(10)
'Dim sEnableGR, sEnableStd, sEnableRec

Dim QfyStatusTitle(10), QfyStatusText(10), QfyStatusLongText(10), QfyStatusColor(10), QualStatusEvent(10), sBoat(10)
Dim sTRounds(10)


' --- Added 8-29-2009 initially for Wakeboard ---
Dim sSkillName1, sSkillName2, sSkillName3, sSkillName4, sSkillName5
Dim sSkill(10), sShowSkills

Dim ScratchCount
Dim sValidTour

' --- Used in qualifications
Dim sRegl_Plc, sNatl_Plc


' --- NEED TO ADD TO LIST AND DELETE FROM Registration.asp ---


Dim sPayPalAct, sPayPalOK, sPayPalActionURL, sAllowOfflinePmt
Dim marksemail


Dim EvtNo, TotEv
TotEv=10


marksemail="mark@productdesign-biz.com"






' --- Added 7-1-2010 --- to avoid error at line 1776 in RegFormDisplay ---
sOF1Fee=0
sOF2Fee=0
sOF3Fee=0
sOF4Fee=0
sOF5Fee=0
sOF6Fee=0
sOF7Fee=0
sOF8Fee=0
sOF9Fee=0
sOF10Fee=0

sOF1Qty=0
sOF2Qty=0
sOF3Qty=0
sOF4Qty=0
sOF5Qty=0
sOF6Qty=0
sOF7Qty=0
sOF8Qty=0
sOF9Qty=0
sOF10Qty=0


' -----------------------------
    SUB DefineTourVariables_New
' -----------------------------


' -------------------------------------------------------------------------
' -----------------------  LOAD TOURNAMENT INFORMATION   ------------------
' -------------------------------------------------------------------------



response.write("IN tools_Reg2 - sTourID = "&sTourID)


' --- Dim the input and output variables
' --- This sections calls a web service buily by Jim Meis and returns a functionname used later in a query
Dim sFunctionName 

' --- Define sTournAppID based on the sTourID variable used throughout the OLR programming
sTournAppID = LEFT(sTourID,6)




' --- MOK 4-15-2013  Since we moved to a Server 2008 box, the Soap Client is not isntalled on the server
' --- so I replaced it with the XMLHTTP control
' --- Dim oSoapClient, mySoapClient
' --- Set mySoapClient = Server.CreateObject("MSSOAP.SoapClient30")
' --- mySoapClient.ClientProperty("ServerHTTPRequest") = True
' --- mySoapClient.mssoapinit("http://www.usawaterski.org/sanctions/webservices/swiftservices.asmx?WSDL")
' --- sFunctionName = "dbo."&mySoapClient.GetLJSearchFunction(sTournAppID)
' --- Set mySoapClient = Nothing



' -----------------------------------------
' --- Post the XML and request response ---
' -----------------------------------------

Dim xmlhttp, DataToSend, postUrl
DataToSend="TournAppID="&sTournAppID

postUrl = "http://www.usawaterski.org/sanctions/webservices/swiftservices.asmx/GetLJSearchFunction"
Set xmlhttp = server.Createobject("MSXML2.XMLHTTP")
xmlhttp.Open "POST",postUrl,false
xmlhttp.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
xmlhttp.send DataToSend
sFunctionName = xmlhttp.responseText

session("debug-XMLReturned") = sFunctionName
session("debug-sTournAppID") = sTournAppID

' --- This is what the response looks like ---
'=<?xml version="1.0" encoding="utf-8"?><string xmlns="http://usawaterski.org/sanctions/webservices">fn_LJsearch2010XTournAppID('13S154')</string>



' ---------------------------------------------------------------------------------------------------------
' --- Test for the word error or invalid and if not present strip the XML to find sFunction used in SQL ---
' ---------------------------------------------------------------------------------------------------------

'response.write "Starting XML " & sXML & "</br>"
Dim ValidFunction
ValidFunction = False

IF InStr(LCASE(sFunctionName),"error") > 0 or InStr(LCASE(sFunctionName),"invalid") > 0 THEN
		' --- Do nothing because we want ValidFunction to be false ---
ELSE
		' --- No error so find the part of the response that describes the WebServices ---
		LocationofWordWebservies = InStr(sFunctionName,"webservices")
		IF LocationofWordWebservies > 0 THEN
				' --- Remove the XML tag from the front of the response ---
				sFunctionName = MID(sFunctionName, LocationofWordWebservies + 13,99)
				
				' --- Now remove the XML tag from the end of the response ---
				LocationofwordString = InStr(sFunctionName,"</string>")
				IF LocationofwordString > 0 THEN
						sFunctionName = LEFT(sFunctionName, LocationofwordString -1)
						ValidFunction = True
				END IF
		END IF
END IF

' response.write "HHH" & sFunctionName & "HHH </br>"
session("debug-sFunctionName") = sFunctionName
session("debug-ValidFunction") = ValidFunction

WhatCalendarYear=0


' --- Remnant from some other approach ?? --- 
'SELECT CASE LEFT(sFunctionName,36)
'	CASE "dbo.fn_LJsearch2010XTournApp"
'		WhatCalendarYear=2010
'	CASE "dbo.fn_search2010XTournAppID"
'		WhatCalendarYear=2010
'	CASE "dbo.fn_SearchxTournAppID_NoL"
'		WhatCalendarYear=2009
'	CASE "dbo.fn_SearchxTournAppID_Lin"
'		WhatCalendarYear=2009
'	CASE "dbo.fn_search2008XTournAppID"
'		WhatCalendarYear=2008
'END SELECT

' --- Temporary display to sort out the functions ---
'response.write("<br><br>Position 1 ")
'response.write("<br>sTournAppID = "&sTournAppID)
'response.write("<br>sFunctionName = "&sFunctionName)
'response.write("<br>sfunctionName36="&LEFT(sFunctionName,36))
'response.write("<br>WhatCalendarYear="&WhatCalendarYear)
'response.end



' ----------------------------------------------------------------------
' --- FUNCTIONS RETURNED WHEN THE GetSearchFunction IS USED          ---
' ----------------------------------------------------------------------

'   sFunctionName = "dbo."&mySoapClient.GetSearchFunction(sTournAppID)

' --- Calendar year 2010
' --- sFunctionName = dbo.fn_search2010XTournAppID('10S075')

' --- Use for calendar year 2009
' --- sFunctionName = dbo.fn_SearchxTournAppID_NoLink('09S078')
' --- sFunctionName = dbo.fn_SearchxTournAppID_Linked('09W082')

' --- For Calendar Year 2008
' --- sFunctionName = dbo.fn_search2008XTournAppID('08S052')
' ----------------------------------------------------------------------



' ----------------------------------------------------------------------
' --- FUNCTIONS RESULTS WHEN THE GetOLRFunction IS USED           ---
' ----------------------------------------------------------------------

' *** NOTE:  GetOLRFunction is NOT used - For reference only      ***

' --- sFunctionName = mySoapClient.GetOLRFunction(sTournAppID)

' --- For Calendar Year 2010
' --- sFunctionName = dbo.fn_SwiftFields2010XTournAppID('10S075')
' --- sFunctionName = dbo.Invalid TournAppID - no records returned

' --- For Calendar Year 2009
' --- sFunctionName = dbo.fn_OLRxTournAppIDLinked('09C069')
' --- sFunctionName = dbo.fn_OLRxTournAppID_NoLink('10W049')

' --- For Calendar Year 2008
' --- sFunctionName = dbo.fn_OLRegFieldsXTournAppID('08S017')
' ----------------------------------------------------------------------



' ---------------------------------
' --- How to tell if OLR is in use
' ---------------------------------    
' --- UseOLReg=1 (true) 

' --- Qualifying fields
' --- 	OLR_PD=1 if $15 fee is paid
' --- 	PayPalOK=1 if PayPal Account has been validated
' --- 	PayPalAct has PayPal account 




' -------------------------------------
' --- Top of TEST for ValidFunction ---
' -------------------------------------

IF ValidFunction = False THEN
		sValidTour=false

ELSE	' --- Found some valid tour and function
		sValidTour = true


		' ------------------------------------------------------------------------------------	
		' --- Opens connection and runs query for tournament based on the correct function ---
		' ------------------------------------------------------------------------------------
		OpenConOLReg
		set rsTSetUp=Server.CreateObject("ADODB.recordset")

		sSQL = "SELECT * FROM "&sFunctionName
		rsTSetUp.open sSQL, sConnectionToOLRegFunction



    ' --- Same test regardless of whether or not OLR is set up ---
    IF NOT rsTSetup.eof THEN

				'response.write("<br><br>In Tools_Reg - After eof test")


				' --- Moved inside IF on 9-8-2010
				sTYear=rsTSetUp("TYear")

				' ----------------------------------------------------------------------
				' ---- 			INITIALIZE VARIABLES			     ---	 					
				' ----------------------------------------------------------------------

				sHQAccount=false

				sTSptsGrpID = rsTSetUp("SptsGrpID")

				' --- Settings for OLR
				sUseOLReg=0
				sOLR_Pd=0

				sTSanction=""
				sTourName = ""
				sTourCity = ""
				sTourState = ""
				sTDateS = ""
				sTDateE = ""
				sSptsGrpID = ""
				sTLateDate = ""
				sTLFPerDay = ""
				sTLateFee = ""
				sTDirEmail = ""

				sTPandC = 0
				sTPandCPulls = 0

				sTRegistrarName = ""
				sTRegistrarAddr = ""
				sTRegistrarCity = ""
				sTRegistrarState = ""
				sTRegistrarZip = ""
				sTRegistrarPhone = ""
				sTRegistrarEmail = ""

				sTsemail = ""


				' --- SITE AND SPONSOR INFORMATION ---
				sTStatus=""
				sTDeleted=""
				sTDirName=""
				sTSponsor=""
				sTSite=""
				sTOpenClosed=""
				sTEntryLimit=""

				' --- Description of tournament ---
				sTDescription = ""
				sFDescription = ""
				sWDescription = ""
				sKDescription = ""
				sCDescription = ""

				' --- Guidebook information ---
				sTDvOffered=""
				sGTSDirections=""
				GTSofE=""
				GTComments=""

				sGTAccommodation=""
				sGTAwards=""
				sGTPractice=""
				sGTStartTime=""
				sGTSofE=""
				sGTComments=""
				sG_IWWF_req=""

' --- NOT FORMATTED FROM HERE DOWN ---

	' --- OFFICIALS - Names used begining in 2010 ---
	sCJudge=""
	sCScorer=""
	sCDriver=""
	sCSafety=""
	sAnnouncer=""
	sTechCont=""
	sPanAmJudge=""
	sOOJ=""
	sAp1Judge=""
	sAp2Judge=""
	sAp3Judge=""
	sAp4Judge=""
	sAp5Judge=""


	' --- Fun Event information - Used prior to 2010---
	'sTEventF3ev=""
	sTEventFKB=""
	sTEventFDA=""

	sTEventFW=""
	sTEventFB=""
	sTEventFHF=""



	' --- Primarily AWS
	SLPremierCnt=0
	TRPremierCnt=0
	JUPremierCnt=0

	sTEventSlalom=false
	sTEventTrick=false
	sTEventJump=false

	' --- Legacy fields from pre-2010
	sTHSClassR=0	' --- Bit field
	sTHSClassL=0
	sTHSClassE=0
	sTHSClassC=0
	sTHSClassN=0
	sTHSClassF=0 

	sTHTClassR=0
	sTHTClassL=0
	sTHTClassE=0
	sTHTClassC=0
	sTHTClassN=0
	sTHTClassF=0 

	sTHJClassR=0
	sTHJClassL=0
	sTHJClassE=0
	sTHJClassC=0
	sTHJClassN=0
	sTHJClassF=0



	' --- AWS Legacy fields from pre-2010 ---
	sTRoundsS=0
	sTRoundsT=0
	sTRoundsJ=0

	' --- Fields for defining # of rounds of each AWS class - used beginning 2010 --
	SClassC=0
	SClassE=0
	SClassL=0
	SClassR=0
	SClassCash=0
	SClassX=0

	TClassC=0
	TClassE=0
	TClassL=0
	TClassR=0
	TClassCash=0
	TClassX=0

	JClassN=0
	JClassC=0
	JClassE=0
	JClassL=0
	JClassR=0
	JClassCash=0
	JClassX=0

	' --- NCWSA Collegiate ---

	sUSClassC=0
	sUTClassC=0
	sUTClassC=0

	sUSClassL=0
	sUTClassL=0
	sUTClassL=0



	' --- USW Events and Classes from 2009 ---
	sTEventWake=""
	sTEventWSkate=""
	sTEventWSurf=""

	' --- USW new field names for 2010 - Each are integers indicating the number of rounds
	WakeW=0
	WSkateW=0
	WSurfW=0



	' --- AKA legacy fields from 2009 - OLR not used by any kneeboard tournaments in 2009 --- 
	sTEventFlip = 0		' --- bit field
	sTEventFree = 0

	sKFlipClassT=0		' --- bit field
	sKFlipClassQ=0
	sKFreeClassT=0
	sKFreeClassQ=0
	sKSlalomClassT=0
	sKSlalomClassQ=0
	sKTrickClassT=0
	sKTrickClassQ=0


	' --- AKA Kneeboard ---
	KSClassQ=0
	KSClassT=0
	KTClassQ=0
	KTClassT=0
	KFlClassQ=0
	KFlClassT=0
	KFrClassQ=0
	KFrClassT=0



	BSClassC=0
	BSClassL=0
	BSClassR=0
	BTClassC=0
	BTClassL=0
	BTClassR=0
	BJClassC=0
	BJClassL=0
	BJClassR=0



	' --- GRASSROOTS fields 3-28-2009 ---
	sGRFunDay = 0		' --- bit field
	sGRTournament = 0	' --- bit field

	Gr1AWSPulls = 0
	Gr1ABCPulls = 0
	Gr1USWPulls = 0
	Gr1AKAPulls = 0
	Gr1USHPulls = 0
	Gr1WSDPulls = 0

	Gr2USH_FreeRidePulls = 0
	Gr2USH_JumpOutPulls = 0
	Gr2USH_BigAirPulls = 0
	Gr2USH_3TrickPulls = 0
	Gr2AWS_SPulls = 0
	Gr2AWS_TPulls = 0
	Gr2ABC_SPulls = 0
	Gr2ABC_TPulls = 0
	Gr2USW_WPulls = 0
	Gr2USW_SkatePulls = 0
	Gr2USW_SurfPulls = 0
	Gr2USW_RailJamPulls = 0
	Gr2AKA_SPulls = 0
	Gr2AKA_TPulls = 0
	Gr2AKA_FreePulls = 0
	Gr2AKA_FlipPulls = 0

	IF Gr1USWPulls>0 OR Gr2USW_WPulls>0 THEN
		sShowSkills=true
	END IF

	sGRBoat = false
	sGRCable = false

	' --- OLR fields ------


	' --- Controls whether button is visible or not ---
	' TRegSetUpStatus

	sOLRDisplayStatus = 0  	' --- bit field

	sPayPalAct = ""
	sPayPalOK = 0		' --- bit field
	sPayPalAction = ""   '--- Live or Testing
	sAWSEFDon_OK = 0	' --- bit field  
	sAllowOfflinePmt=0


	sTEntryFee1=0
	sTEntryFee2=0
	sTEntryFee3=0
	sTEntryFeeFamily=0
	sOtherFee = 0

	sMixedOptions=0

	sCSurchg = 0
	sESurchg = 0
	sLSurchg = 0
	sRSurchg = 0

	' ----------------   GRASSROOTS FIELDS --------------------
	sGREntryFeeIncluded = 0		' --- bit field
	sGRDiscount = 0			' --- bit field

	' --- Changed or added 3-22-2009 -----------
	sGREntryFee1 = 0
	sGREntryFee2 = 0
	sGREntryFee3 = 0

	sClinFeeJD = 0
	sClinFeeAD = 0


	' --- MaxPulls is for traditional tournaments to set the limit on # of people regsitered - FUTURE ---
	sMaxPulls = 0
	sReservedPulls = 0
	sReservedPullsCode = 0

	sOffDiscPerc = 0
	sJrDiscPerc = 0
	sSrDiscPerc = 0
	sClubDiscPerc = 0
	sTourClubCode = ""
	sDiscMeth = 0
	sBTickCost = 0
	sBTickWithE = 0		' --- bit field

	' -- Now TLateDate from TSanction
	sQualLevel = 0
	sTourEmail = ""
	sReceiveEmail = 0 	' --- bit field



	' --- OPTIONAL ITEM Fields added March 2009 ---------
	sOF1Desc = ""
	sOF2Desc = ""
	sOF3Desc = ""
	sOF4Desc = ""
	sOF5Desc = ""
	sOF6Desc = ""
	sOF7Desc = ""
	sOF8Desc = ""
	sOF9Desc = ""
	sOF10Desc = ""

	sOF1Amt = 0
	sOF2Amt = 0
	sOF3Amt = 0
	sOF4Amt = 0
	sOF5Amt = 0
	sOF6Amt = 0
	sOF7Amt = 0
	sOF8Amt = 0
	sOF9Amt = 0
	sOF10Amt = 0

	sOF1MaxQty = 0
	sOF2MaxQty = 0
	sOF3MaxQty = 0
	sOF4MaxQty = 0
	sOF5MaxQty = 0
	sOF6MaxQty = 0
	sOF7MaxQty = 0
	sOF8MaxQty = 0
	sOF9MaxQty = 0
	sOF10MaxQty = 0
	
	sOF1Required = 0	' --- bit fields
	sOF2Required = 0
	sOF3Required = 0
	sOF4Required = 0
	sOF5Required = 0
	sOF6Required = 0	
	sOF7Required = 0
	sOF8Required = 0
	sOF9Required = 0
	sOF10Required = 0


	' --- Sets Column Heading Names on RegFormDisplay.asp  ----
	sTGRClassText = ""
	sTBaseClassText = ""
	sTUpgradeClassText = ""


	' --- Determines how many family members included in Family Entry and Cost for Extra Members above limit
	sMaxFamMembers = 0
	sTEntryFeeFamExtra = 0

	sMaxSLPulls = 0
	sMaxTRPulls = 0
	sMaxJUPulls = 0







	' -------------------------------------------------------------------------------------------------
	' ---  			BEGIN READING RECORDS INTO LOCAL VARIABLES 				---
	' -------------------------------------------------------------------------------------------------

' --------------------------------------
	' --- General tournament information ---
	' --------------------------------------

	sTSanction = rsTSetup("TSanction")
	sTourName = rsTSetUp("TName")
	sTourCity = rsTSetUp("TCity")
	sTourState = rsTSetUp("TState")
	sTDateS = rsTSetUp("TDateS")
	sTDateE = rsTSetUp("TDateE")
	sSptsGrpID = rsTSetUp("SptsGrpID")


	sTLateDate = rsTSetUp("TLateDate")
	sTLFPerDay = rsTSetUp("TLFPerDay")
	
	' --- Changed 4-30-2012 per JMeis request 
	'sTLateFee = rsTSetUp("TLateFee")
	sTLateFee = rsTSetUp("LateFee")


  
  ' -----------------------------------------------------------------------------------  
  ' --- Adjustment for time difference between Tournament and USAWS server location ---
  ' -----------------------------------------------------------------------------------  
  sSQL = "SELECT * FROM "&RegionTableName 
  sSQL = sSQL + " WHERE State = '"&sTourState&"'" 
  SET rsTour=Server.CreateObject("ADODB.recordset")
  rsTour.open sSQL, SConnectionToTRATable, 3, 3
 
  IF NOT(rsTour.eof) THEN
  		TimeAdjust = rsTour("TimeAdjust")
  		sTLateDate_WithTime = sTLateDate&" 11:59:59 PM"
  		sTLateDate_Adjusted = DateAdd("h",-TimeAdjust,sTLateDate_WithTime)
  ELSE
  		sTLateDate_Adjusted = sTLateDate 
  END IF 



'  IF sMemberID="000001151" THEN
'  	response.write("<br>"&sSQL) 
'  	response.write("<br><br>Test WEST COAST CUT-OFF")
'		response.write("<br>Saction sTLateDate="&sTLateDate)
'  	response.write("<br>TimeAdjust="&TimeAdjust)
'  	response.write("<br>sTLateDate_Adjusted="&sTLateDate_Adjusted)
		
'  	sTLateDate_WithTime=FormatDateTime(sTLateDate,0)
'		response.write("<br>sTLateDate_WithTime="&sTLateDate_WithTime)
'	END IF
	
	

	' -----------------------
	' --- Pick and Choose ---
	' -----------------------
	sTPandC = rsTSetup("TPandC")			' bit 
	sTPandCPulls = rsTSetup("TPandCPulls")
	IF sTPandCPulls="" THEN sTPandCPulls=0

	' -------------------------------------
	' --- LOC and Registrar information ---
	' -------------------------------------
	sTDirEmail = rsTSetUp("TDirEmail")

	sTRegistrarName = rsTSetUp("TRegistrarName")
	sTRegistrarAddr = rsTSetUp("TRegistrarAddr")
	sTRegistrarCity = rsTSetUp("TRegistrarCity")
	sTRegistrarState = rsTSetUp("TRegistrarState")
	sTRegistrarZip = rsTSetUp("TRegistrarZip")
	sTRegistrarPhone = rsTSetUp("TRegistrarPhone")
	sTRegistrarEmail = rsTSetUp("TRegistrarEmail")

	' --- Latest rep to make a change
	sTsemail = rsTSetUp("Tsemail")


	' -----------------------------------------
	' --- SITE AND SPONSOR INFORMATION ---
	' -----------------------------------------
	sTStatus=rsTSetUp("TStatus")
	'sTDeleted=rsTSetUp("Deleted")
	sTDirName=rsTSetUp("TDirName")
	sTSponsor=rsTSetUp("TSponsor")
	sTSite=rsTSetUp("TSite")
	sTSiteID=rsTSetUp("TSiteID")
	sTOpenClosed=rsTSetUp("TOpenClosed")
	sTEntryLimit=rsTSetUp("TEntryLimit")
	IF IsNull(sTEntryLimit) THEN sTEntryLimit="None"


	' -----------------------------------------
	' --- Text field for Tournament listing ---
	' -----------------------------------------	
	sTDescription=rsTSetUp("TDescription")
	sFDescription=rsTSetUp("FDescription")
	sWDescription=rsTSetUp("WDescription")
	sKDescription=rsTSetUp("KDescription")
	sCDescription=rsTSetUp("CDescription")


	IF sTYear<2010 THEN
		' -------------------------------------------
		' --- Fun Event information pre 2010 ---
		' -------------------------------------------
		sTEventFKB=rsTSetUp("TEventFKB")
		sTEventFDA=rsTSetUp("TEventFDA")

		sTEventFW=rsTSetUp("TEventFW")
		sTEventFB=rsTSetUp("TEventFB")
		sTEventFHF=rsTSetUp("TEventFHF")
	END IF

	' ----------------------------------------------------------------------------------
	' --- Added 1-9-2010 on conversion of view-tournaments.asp to these definitions  ---
	' ----------------------------------------------------------------------------------
	sGTAccommodation=rsTSetUp("GTAccommodation")
	sGTAwards=rsTSetUp("GTAwards")
	sGTPractice=rsTSetUp("GTPractice")
	sGTStartTime=rsTSetUp("GTStartTime")
	sGTSofE=rsTSetUp("GTSofE")
	sGTComments=rsTSetUp("GTComments")
	sG_IWWF_req=rsTSetUp("G_IWWF_req")
	

	' -----------------------------
	' --- Guidebook information ---
	' -----------------------------
	sTDvOffered=rsTSetUp("TDvOffered") 
	sGTSDirections=rsTSetUp("GTSDirections") 
	GTSofE=rsTSetUp("GTSofE")
	GTComments=rsTSetUp("GTComments")


	'--- CAUTION: Program continues to populate this even though not stored - for those not wanting to use OLR settings ---
	sTEntryFees=rsTSetUp("TEntryFees")


	' --------------------------
	' --- OFFICIALS section ---
	' --------------------------

	IF sTYear>=2010 THEN
		sCJudge=rsTSetUp("CJudge")
		sCScorer=rsTSetUp("CScorer")
		sCDriver=rsTSetUp("CDriver")
		sCSafety=rsTSetUp("CSafety")
		sAnnouncer=rsTSetUp("Announcer")
		sTechCont=rsTSetUp("TechCont")

		sPanAmJudge=rsTSetUp("PanAmJudge")
		sAp1Judge=rsTSetUp("Ap1Judge")
		sAp2Judge=rsTSetUp("Ap2Judge")
		sAp3Judge=rsTSetUp("Ap3Judge")
		sAp4Judge=rsTSetUp("Ap4Judge")
		sAp5Judge=rsTSetUp("Ap5Judge")
	ELSE
		sCJudge=rsTSetUp("TOCJudge")
		sCScorer=rsTSetUp("TOCScore")
		sCDriver=rsTSetUp("TOCDriver")
		sCSafety=rsTSetUp("TOCSafety") 
		sAnnouncer=rsTSetUp("TOAnnounce")
		sTechCont=rsTSetUp("TOTechCont")

		sPanAmJudge=rsTSetUp("TOPanAmJudge")
		sOOJ=rsTSetUp("TOOoAJudge")
		sAp1Judge=rsTSetUp("TO1ApJudge")
		sAp2Judge=rsTSetUp("TO2ApJudge")
		sAp3Judge=rsTSetUp("TO3ApJudge")
		sAp4Judge=rsTSetUp("TO4ApJudge")
		sAp5Judge=rsTSetUp("TO5ApJudge")
	END IF




	' --------------------------------------------------	
	' --- EVENTS OFFERED AND CLASSES OF COMPETITION ---
	' --------------------------------------------------
	IF sTYear>=2010 THEN

		' --- The following fields are INTEGER format
		SClassC=rsTSetUp("SClassC")
		SClassE=rsTSetUp("SClassE")
		SClassL=rsTSetUp("SClassL")
		SClassR=rsTSetUp("SClassR")
		SClassCash=rsTSetUp("SClassCash")
		SClassX=rsTSetUp("SClassX")

		TClassC=rsTSetUp("TClassC")
		TClassE=rsTSetUp("TClassE")
		TClassL=rsTSetUp("TClassL")
		TClassR=rsTSetUp("TClassR")
		TClassCash=rsTSetUp("TClassCash")
		TClassX=rsTSetUp("TClassX")

		' --- Class N used in the case of grassroots Jump
		JClassN=rsTSetUp("JClassN")
		JClassC=rsTSetUp("JClassC")
		JClassE=rsTSetUp("JClassE")
		JClassL=rsTSetUp("JClassL")
		JClassR=rsTSetUp("JClassR")
		JClassCash=rsTSetUp("JClassCash")
		JClassX=rsTSetUp("JClassX")


		' --------------------
		' --- BAREFOOT ABC ---
		' --------------------
		BSClassC=rsTSetUp("BSClassC")
		BSClassL=rsTSetUp("BSClassL")
		BSClassR=rsTSetUp("BSClassR")

		BTClassC=rsTSetUp("BTClassC")
		BTClassL=rsTSetUp("BTClassL")
		BTClassR=rsTSetUp("BTClassR")

		BJClassC=rsTSetUp("BJClassC")
		BJClassL=rsTSetUp("BJClassL")
		BJClassR=rsTSetUp("BJClassR")





	sPayPalAct = rsTSetup("PayPalAct")
	sPayPalOK = rsTSetup("PayPalOK")

	sAllowOfflinePmt = rsTSetup("AllowOfflinePmt")
	'IF adminmenulevel>=50 THEN
	'		sAllowOfflinePmt=1
	'		 response.write("sAllowOfflinePmt="&sAllowOfflinePmt)
	'END IF



	IF sTYear>=2009 THEN

		IF sPayPalOK THEN sOLR_PD=1
		IF sPayPalOK THEN sUseOLReg=1

		' --- New GRASSROOTS fields 3-28-2009 ---
		sGRFunDay = rsTSetUp("GRFunDay")
		sGRTournament = rsTSetUp("GRTournament")

		Gr1AWSPulls = rsTSetUp("Gr1AWSPulls")
		Gr1ABCPulls = rsTSetUp("Gr1ABCPulls")
		Gr1USWPulls = rsTSetUp("Gr1USWPulls")
		Gr1AKAPulls = rsTSetUp("Gr1AKAPulls")
		Gr1USHPulls = rsTSetUp("Gr1USHPulls")
		Gr1WSDPulls = rsTSetUp("Gr1WSDPulls")

		Gr2USH_FreeRidePulls = rsTSetUp("Gr2USH_FreeRidePulls")
		Gr2USH_JumpOutPulls = rsTSetUp("Gr2USH_JumpOutPulls")
		Gr2USH_BigAirPulls = rsTSetUp("Gr2USH_BigAirPulls")
		Gr2USH_3TrickPulls = rsTSetUp("Gr2USH_3TrickPulls")
		Gr2AWS_SPulls = rsTSetUp("Gr2AWS_SPulls")
		Gr2AWS_TPulls = rsTSetUp("Gr2AWS_TPulls")
		Gr2ABC_SPulls = rsTSetUp("Gr2ABC_SPulls")
		Gr2ABC_TPulls = rsTSetUp("Gr2ABC_TPulls")
		Gr2USW_WPulls = rsTSetUp("Gr2USW_WPulls")
		Gr2USW_SkatePulls = rsTSetUp("Gr2USW_SkatePulls")
		Gr2USW_SurfPulls = rsTSetUp("Gr2USW_SurfPulls")
		Gr2USW_RailJamPulls = rsTSetUp("Gr2USW_RailJamPulls")
		Gr2AKA_SPulls = rsTSetUp("Gr2AKA_SPulls")
		Gr2AKA_TPulls = rsTSetUp("Gr2AKA_TPulls")
		Gr2AKA_FreePulls = rsTSetUp("Gr2AKA_FreePulls")
		Gr2AKA_FlipPulls = rsTSetUp("Gr2AKA_FlipPulls")

		IF Gr1USWPulls>0 OR Gr2USW_WPulls>0 THEN
			sShowSkills=true
		END IF

		sGRBoat = rsTSetup("GRBoat")
		sGRCable = rsTSetup("GRCable")
	END IF



		' ----------------------------------------------------------------------
		' --- Finds MAX rounds for each event regardless of class ---
		' --- Plan would be to modify entry to specify rounds for each class ---
		' ----------------------------------------------------------------------
		sTRoundsS=SClassC
		IF SClassE>sTRoundsS then sTRoundsS=SClassE
		IF SClassL>sTRoundsS then sTRoundsS=SClassL
		IF SClassR>sTRoundsS then sTRoundsS=SClassR
		IF SClassCash>sTRoundsS then sTRoundsS=SClassCash
		IF SClassX>sTRoundsS then sTRoundsS=SClassX
		IF Gr2AWS_SPulls>sTRoundsS then sTRoundsS=Gr2AWS_SPulls
		IF sTRoundsS>0 THEN sTEventSlalom=true

		sTRoundsT=TClassC
		IF TClassE>sTRoundsT then sTRoundsT=TClassE
		IF TClassL>sTRoundsT then sTRoundsT=TClassL
		IF TClassR>sTRoundsT then sTRoundsT=TClassR
		IF TClassCash>sTRoundsT then sTRoundsT=TClassCash
		IF TClassX>sTRoundsT then sTRoundsT=TClassX
		IF Gr2AWS_TPulls>sTRoundsT then sTRoundsT=Gr2AWS_TPulls
		IF sTRoundsT>0 THEN sTEventTrick=true

		sTRoundsJ=JClassC
		IF JClassN>sTRoundsJ then sTRoundsJ=JClassN
		IF JClassE>sTRoundsJ then sTRoundsJ=JClassE
		IF JClassL>sTRoundsJ then sTRoundsJ=JClassL
		IF JClassR>sTRoundsJ then sTRoundsJ=JClassR
		IF JClassCash>sTRoundsJ then sTRoundsJ=JClassCash
		IF JClassX>sTRoundsJ then sTRoundsJ=JClassX
		IF Gr2AWS_JPulls>sTRoundsJ then sTRoundsJ=Gr2AWS_JPulls
		IF sTRoundsJ>0 THEN sTEventJump=true


		' ----------------------------------------------------------------------------------------------------------
		' --- This is a somewhat cumbersome process now that SClassC, etc show # of rounds for each, but this is 
		' ---   done this way to prevent having to change downstream code.
		' ----------------------------------------------------------------------------------------------------------
		IF SClassC>0 OR SClassE>0 OR SClassL>0 OR SClassR>0 OR SClassCash>0 OR SClassX>0 THEN SLPremierCnt=SLPremierCnt+1
		IF TClassC>0 OR TClassE>0 OR TClassL>0 OR TClassR>0 OR TClassCash>0 OR TClassX>0 THEN TRPremierCnt=TRPremierCnt+1
		IF JClassC>0 OR JClassE>0 OR JClassL>0 OR JClassR>0 OR JClassCash>0 OR JClassX>0 THEN JUPremierCnt=JUPremierCnt+1


		' --------------------
		' --- BAREFOOT ABC ---
		' --------------------

'IF LEFT(sTourID,6)="11B046" THEN BSClassC=1


		IF BSClassC>0 OR BSClassL>0 OR BSClassR>0 THEN SLPremierCnt=SLPremierCnt+1
		IF BTClassC>0 OR BTClassL>0 OR BTClassL>0 THEN TRPremierCnt=TRPremierCnt+1
		IF BJClassC>0 OR BJClassL>0 OR BJClassL>0 THEN TRPremierCnt=TRPremierCnt+1

		IF BSClassC + BSClassL + BSClassR > 0 then sEventSlalom = true	
		IF BTClassC + BTClassL + BTClassR > 0 then sEventTrick = true	
		IF BJClassC + BJClassL + BJClassR > 0 then sEventJump = true			



		'-------------------------
		' --- Fields for NCWSA ---
		'-------------------------

		sUSClassC=rsTSetUp("USClassC")
		sUTClassC=rsTSetUp("UTClassC")
		sUTClassC=rsTSetUp("UJClassC")
		IF USClassC <> 0 THEN
				sTEventSlalom = true		
				SLPremierCnt = 1
		END IF
		IF UTClassC <> 0 THEN 
				sTEventTrick = true
				TRPremierCnt = 1
		END IF
		IF UJClassC <> 0 THEN 
				sTEventJump = true
				JUPremierCnt = 1
		END IF




		' ------------------------------
		' --- AKA Kneeboard ---
		' ------------------------------

		KSClassQ=rsTSetUp("KSClassQ")
		KSClassT=rsTSetUp("KSClassT")
		KTClassQ=rsTSetUp("KTClassQ")
		KTClassT=rsTSetUp("KTClassT")
		KFlClassQ=rsTSetUp("KFlClassQ")
		KFlClassT=rsTSetUp("KFlClassT")
		KFrClassQ=rsTSetUp("KFrClassQ")
		KFrClassT=rsTSetUp("KFrClassT")

		IF KSClassQ+KSClassT>0 then sTEventSlalom = true
		IF KTClassQ+KTClassT>0 then sTEventTrick = true
		IF KFlClassQ+KFlClassT>0 then sTEventFlip = true
		IF KFrClassQ+KFrClassT>0 then sTEventFree = true
		

		' ------------------------------
		' --- USW Events and Classes ---
		' ------------------------------
		WWakeW=rsTSetUp("WWakeW")
		WSkateW=rsTSetUp("WSkateW")
		WSurfW=rsTSetUp("WSurfW")
		IF WWakeW>0 THEN sTEventWake=true
		IF WSkateW>0 THEN sTEventWSkate=true
		IF WSurfW>0 THEN sTEventWSurf=true

		

	ELSE
		
		' ***********************************
		' ***********************************
		' --- Pre 2010 definition section ---
		' ***********************************
		' ***********************************		

		sTRoundsS=rsTSetUp("TRoundsS")
		sTRoundsT=rsTSetUp("TRoundsT")
		sTRoundsJ=rsTSetUp("TRoundsJ")

		sTEventSlalom=rsTSetUp("TEventSlalom")
		sTEventJump=rsTSetUp("TEventJump")
		sTEventTrick=rsTSetUp("TEventTrick")

		' --------------------- AWS Classes ----------------------
		' --- SLALOM ---
		sTHSClassR=rsTSetup("THSClassR") 
		sTHSClassL=rsTSetup("THSClassL")
		sTHSClassE=rsTSetup("THSClassE")
		sTHSClassC=rsTSetup("THSClassC")
		sTHSClassN=rsTSetup("THSClassN")
		sTHSClassF=rsTSetup("THSClassF") 

		' --- Determine # of classes offered to set whether both BASE and UPGRADE radio buttons are displayed
		SLPremierCnt=0
		IF sTHSClassC=true THEN SLPremierCnt=SLPremierCnt+1
		IF sTHSClassE=true THEN SLPremierCnt=SLPremierCnt+1
		IF sTHSClassL=true THEN SLPremierCnt=SLPremierCnt+1
		IF sTHSClassR=true THEN SLPremierCnt=SLPremierCnt+1

		' --- TRICKS ---
		sTHTClassR=rsTSetup("THTClassR") 
		sTHTClassL=rsTSetup("THTClassL")
		sTHTClassE=rsTSetup("THTClassE")
		sTHTClassC=rsTSetup("THTClassC")
		sTHTClassN=rsTSetup("THTClassN")
		sTHTClassF=rsTSetup("THTClassF") 

		' --- Determine # of classes offered to set whether both BASE and UPGRADE radio buttons are displayed
		TRPremierCnt=0
		IF sTHSClassC=true THEN TRPremierCnt=TRPremierCnt+1
		IF sTHSClassE=true THEN TRPremierCnt=TRPremierCnt+1
		IF sTHSClassL=true THEN TRPremierCnt=TRPremierCnt+1
		IF sTHSClassR=true THEN TRPremierCnt=TRPremierCnt+1


		' --- JUMP ---
		sTHJClassR=rsTSetup("THJClassR") 
		sTHJClassL=rsTSetup("THJClassL")
		sTHJClassE=rsTSetup("THJClassE")
		sTHJClassC=rsTSetup("THJClassC")
		sTHJClassN=rsTSetup("THJClassN")
		sTHJClassF=rsTSetup("THJClassF") 

		' --- Determine # of classes offered to set whether both BASE and UPGRADE radio buttons are displayed
		JUPremierCnt=0
		IF sTHSClassC=true THEN JUPremierCnt=JUPremierCnt+1
		IF sTHSClassE=true THEN JUPremierCnt=JUPremierCnt+1
		IF sTHSClassL=true THEN JUPremierCnt=JUPremierCnt+1
		IF sTHSClassR=true THEN JUPremierCnt=JUPremierCnt+1


		
		' ------------------------------
		' --- USW Events and Classes ---
		' ------------------------------
		sTEventWake=rsTSetUp("TEventWake")
		sTEventWSkate=rsTSetup("TEventWSkate")
		sTEventWSurf=rsTSetup("TEventWSurf")

		' ------------------------------
		' --- AKA Events and Classes ---
		' ------------------------------
		sKEventFlip=rsTSetup("KEventFlip") 
		sKEventFree=rsTSetup("KEventFree")
		sTEventSlalom=rsTSetUp("TEventSlalom")
		sTEventTrick=rsTSetUp("TEventTrick")

		sKFlipClassT=rsTSetup("KFlipClassT")
		sKFlipClassQ=rsTSetup("KFlipClassQ")
		sKFreeClassT=rsTSetup("KFreeClassT")
		sKFreeClassQ=rsTSetup("KFreeClassQ")
		sKSlalomClassT=rsTSetup("KSlalomClassT")
		sKSlalomClassQ=rsTSetup("KSlalomClassQ")
		sKTrickClassT=rsTSetup("KTrickClassT")
		sKTrickClassQ=rsTSetup("KTrickClassQ")

	END IF  ' --- End of selected definitions based on YEAR of tournament 







'IF Session("AdminMenuLevel")>49 AND LEFT(sTourID,6)="13M081" THEN
'			response.write("<br>Line 1209 Tools_Registration - TournAppID = 13M081")
'			response.write("<br> Gr2AWS_SPulls = "&Gr2AWS_SPulls)
'			response.write("<br> Gr2AWS_TPulls = "&Gr2AWS_TPulls)
'			response.write("<br> Gr2AWS_JPulls = "&Gr2AWS_JPulls)			
'END IF	




		' -------------------------------------------------
		' --- Controls whether button is visible or not ---
		' -------------------------------------------------
		sOLRDisplayStatus = rsTSetup("OLRDisplayStatus")
		sUseOLReg=rsTSetup("UseOLReg")

	' --- Change to 2009 once Jim adds the OLR fields into GetSearch for 2009 		
		IF sTYear>=2010 THEN	' --- Stuff from OLR setup ---
			sOLR_Pd=rsTSetup("OLR_Pd")
		END IF



' *****************************************
' *****************************************
' *****************************************

		Session("AdminCode")=rsTSetup("AdminCode")
		'IF LEFT(sTourID,6)="15W091" THEN response.write("<br>Tools 1286 Session(AdminCode) = "&rsTSetup("AdminCode"))
		'response.write("<br>Session(AdminCode) = "&Session("AdminCode"))

		' ------------------
		' --- Entry Fees ---
		' ------------------
		sAWSEFDon_OK = rsTSetUp("AWSEFDon_OK")  
		sTEntryFee1=rsTSetUp("EntryFee1")
		sTEntryFee2=rsTSetUp("EntryFee2")
		sTEntryFee3=rsTSetUp("EntryFee3")
		sTEntryFeeFamily=rsTSetUp("EntryFeeFamily")
		sOtherFee = rsTSetUp("OtherFee")

		sMixedOptions = rsTSetUp("MixedOptions")

		sCSurchg = rsTSetUp("CSurchg")
		sESurchg = rsTSetUp("ESurchg")
		sLSurchg = rsTSetUp("LSurchg")
		sRSurchg = rsTSetUp("RSurchg")

'IF LEFT(sTourID,6)="14S100" THEN
		'sTEntryFee1=105
		'sTEntryFee2=20
		'sTEntryFee3=20
		'		response.write("<br>sTEntryFee1 = "&sTEntryFee1)
		'		response.write("<br>sTEntryFee2 = "&sTEntryFee2)
		'		response.write("<br>sTEntryFee3 = "&sTEntryFee3)		
		
		'		response.write("<br>sRSurchg = "&sRSurchg)		
'END IF

		' ---------------------------
		' ---   GRASSROOTS FIELDS ---
		' ---------------------------
		sGREntryFeeIncluded = rsTSetUp("GREntryFeeIncluded")
		sGRDiscount = rsTSetUp("GRDiscount")
		' --- Changed or added 3-22-2009 ---
		sGREntryFee1 = rsTSetUp("GRFee_1")
		sGREntryFee2 = rsTSetUp("GRFee_2")
		sGREntryFee3 = rsTSetUp("GRFee_3")

		' -------------------
		' --- Clinic Fees ---
		' -------------------
		sClinFeeJD = rsTSetUp("ClinFeeJD")
		sClinFeeAD = rsTSetUp("ClinFeeAD")


		' ---------------------------------------------------------------------------------------------------
		' --- MaxPulls is for traditional tournaments to set the limit on # of people regsitered - FUTURE ---
		' ---------------------------------------------------------------------------------------------------
		sMaxPulls = rsTSetUp("MaxPulls")
		IF IsNull(sMaxPulls) THEN sMaxPulls=0
		sReservedPulls = rsTSetUp("ReservedPulls")
		sReservedPullsCode = rsTSetUp("ReservedPullsCode")

		sOffDiscPerc = rsTSetUp("Disc1")
		sJrDiscPerc = rsTSetUp("Disc2")
		sSrDiscPerc = rsTSetUp("Disc3")
		sClubDiscPerc = rsTSetUp("Disc4")
		sTourClubCode = rsTSetup("ClubCode")
		sDiscMeth = rsTSetUp("DiscMeth")
		sBTickCost = rsTSetUp("BTickCost")
		sBTickWithE = rsTSetUp("BTickWithE")

		' -- Now TLateDate from TSanction
		sQualLevel = rsTSetUp("QualLevel")
		sTourEmail = TRIM(rsTSetUp("EmailAddress"))
		sReceiveEmail = rsTSetup("ReceiveEmail")

		IF LEFT(sTourID,6)="12S999" THEN sQualLevel=8

		' --------------------------------------------------------------
		' --- Determines if payment processor goes to PayPal or Card ---
		' --------------------------------------------------------------
		IF LCASE(TRIM(sPayPalAct))="usawaterski@usawaterski.org" OR LCASE(TRIM(sPayPalAct))="hqmerchant@usawaterski.org" THEN sHQAccount=true

		

		' -----------------------------------------------------------------------------------		
		' --- Determine whether to use the test sandbox or the live action URL for PayPal ---	
		' -----------------------------------------------------------------------------------
		TestPayPal="N"
		IF sMemberID="000001151" AND TestPayPal="Y" THEN
				sHQAccount=false
				sPayPalActionURL="https://www.sandbox.paypal.com/cgi-bin/webscr"
				sPayPalAct="mark@kingsbridgehomes.com"		
		ELSE
				sPayPalActionURL= "https://www.paypal.com/cgi-bin/webscr"
		END IF





		' ---------------------------------------------------
		' --- OPTIONAL ITEM Fields added March 2009 ---------
		' ---------------------------------------------------

		sOF1Desc = rsTSetup("OF1Desc")
		sOF2Desc = rsTSetup("OF2Desc")
		sOF3Desc = rsTSetup("OF3Desc")
		sOF4Desc = rsTSetup("OF4Desc")
		sOF5Desc = rsTSetup("OF5Desc")
		sOF6Desc = rsTSetup("OF6Desc")
		sOF7Desc = rsTSetup("OF7Desc")
		sOF8Desc = rsTSetup("OF8Desc")
		sOF9Desc = rsTSetup("OF9Desc")
		sOF10Desc = rsTSetup("OF10Desc")

		sOF1Amt = rsTSetup("OF1Amt")
		sOF2Amt = rsTSetup("OF2Amt")
		sOF3Amt = rsTSetup("OF3Amt")
		sOF4Amt = rsTSetup("OF4Amt")
		sOF5Amt = rsTSetup("OF5Amt")
		sOF6Amt = rsTSetup("OF6Amt")
		sOF7Amt = rsTSetup("OF7Amt")
		sOF8Amt = rsTSetup("OF8Amt")
		sOF9Amt = rsTSetup("OF9Amt")
		sOF10Amt = rsTSetup("OF10Amt")

		sOF1MaxQty = rsTSetup("OF1MaxQty")
		sOF2MaxQty = rsTSetup("OF2MaxQty")
		sOF3MaxQty = rsTSetup("OF3MaxQty")
		sOF4MaxQty = rsTSetup("OF4MaxQty")
		sOF5MaxQty = rsTSetup("OF5MaxQty")
		sOF6MaxQty = rsTSetup("OF6MaxQty")
		sOF7MaxQty = rsTSetup("OF7MaxQty")
		sOF8MaxQty = rsTSetup("OF8MaxQty")
		sOF9MaxQty = rsTSetup("OF9MaxQty")
		sOF10MaxQty = rsTSetup("OF10MaxQty")

		sOF1Required = rsTSetup("OF1Required")
		sOF2Required = rsTSetup("OF2Required")
		sOF3Required = rsTSetup("OF3Required")
		sOF4Required = rsTSetup("OF4Required")
		sOF5Required = rsTSetup("OF5Required")
		sOF6Required = rsTSetup("OF6Required")
		sOF7Required = rsTSetup("OF7Required")
		sOF8Required = rsTSetup("OF8Required")
		sOF9Required = rsTSetup("OF9Required")
		sOF10Required = rsTSetup("OF10Required")

IF sMemberID="000001151" AND LEFT(sTourID,6)="13S999" THEN
		sOF1Desc="Mark Test Item"
		sOF1MaxQty=10
		sOF1Amt=1
END IF		

TotNumOptItems=0
IF TRIM(sOF1Desc)<>"" THEN TotNumOptItems=TotNumOptItems+1
IF TRIM(sOF2Desc)<>"" THEN TotNumOptItems=TotNumOptItems+1
IF TRIM(sOF3Desc)<>"" THEN TotNumOptItems=TotNumOptItems+1
IF TRIM(sOF4Desc)<>"" THEN TotNumOptItems=TotNumOptItems+1
IF TRIM(sOF5Desc)<>"" THEN TotNumOptItems=TotNumOptItems+1
IF TRIM(sOF6Desc)<>"" THEN TotNumOptItems=TotNumOptItems+1
IF TRIM(sOF7Desc)<>"" THEN TotNumOptItems=TotNumOptItems+1
IF TRIM(sOF8Desc)<>"" THEN TotNumOptItems=TotNumOptItems+1
IF TRIM(sOF9Desc)<>"" THEN TotNumOptItems=TotNumOptItems+1
IF TRIM(sOF10Desc)<>"" THEN TotNumOptItems=TotNumOptItems+1




'IF Session("AdminMenuLevel")>=50 THEN
'		sOF1Desc="Select '1' to enter in <b>BOTH</b> your Age and Pro divisions"
'		sOF1Amt=0
'		sOF1MaxQty=1
'		sOF1Required=0

		'response.write("<br>MarkTester="&MarkTester)
		'	response.write("<br>WSkateW = "&WSkateW)
		'	response.write("<br>sTEvent(EvtNo) = "&sTEvent(EvtNo))
		'response.write("<br>sTEventWake = "&sTEventWake)
		'response.write("<br>Gr1USWPulls = "&Gr1USWPulls)
		'response.write("<br>Gr2USW_WPulls = "&Gr2USW_WPulls)
'END IF


		' ---------------------------------------------------------
		' --- Sets Column Heading Names on RegFormDisplay.asp  ----
		' ---------------------------------------------------------
		sTGRClassText = TRIM(rsTSetup("GRClassText"))
		sTBaseClassText = TRIM(rsTSetup("BaseClassText"))
		sTUpgradeClassText = TRIM(rsTSetup("UpgradeClassText"))

		IF sTGRClassText = "" THEN sTGRClassText="Grassroots"
		IF sTBaseClassText = "" THEN sTBaseClassText = "Standard"
		IF sTUpgradeClassText = "" THEN sTUpgradeClassText = "Record"

		'IF sTBaseClassText = "Premier" AND sTUpgradeClassText = "None" THEN sTUpgradeClassText = "Premier"

		' --- TEMPORARY ---
		IF LEFT(sTourID,6)="09S999" THEN
				sTBaseClassText="Record Only"
		END IF


		' ----------------------------------------------------------------------------------------------------------
		' --- Determines how many family members included in Family Entry and Cost for Extra Members above limit
		' ----------------------------------------------------------------------------------------------------------
		sMaxFamMembers = rsTSetup("MaxFamMember")
		sTEntryFeeFamExtra = rsTSetup("TEntryFeeFamExtra")


		' --------------------------------------
		' --- Max slalom pulls used in P & C ---
		' --------------------------------------
		sMaxSLPulls = rsTSetup("MaxSLPulls")
		sMaxTRPulls = rsTSetup("MaxTRPulls")
		sMaxJUPulls = rsTSetup("MaxJPPulls")	' ---Note SWIFT calls it JP not JU

   END IF



END IF			' --- IF related to whether function was INVALID and tour was not found ---

END SUB





' -------------------------------
  Function TestValidAdminCode
' -------------------------------


' --- Tests whether the person is logged in with SWIFT AdminCode ---
TestValidAdminCode=false
IF Session("AdminMenuLevel")>=30 OR ( TRIM(Session("UserAdminPW"))<>"" AND UCASE(TRIM(Session("UserAdminPW")))=UCASE(TRIM(Session("AdminCode"))) ) THEN
	TestValidAdminCode=true
END IF

Dim t
t=2
IF t=1 AND Session("AdminMenuLevel")>=50 THEN
   response.write("<br>in Tools - TestVAdmin=")
   response.write(TestValidAdminCode)
   response.write("<br>UCASE(TRIM(Session(UserAdminPW)))=UCASE(TRIM(Session(AdminCode)))")
   response.write(UCASE(TRIM(Session("UserAdminPW")))=UCASE(TRIM(Session("AdminCode"))))
END IF

END Function





' -------------------------------------
   Function EntriesExceedLimit(sTourID)
' -------------------------------------

'response.write("<br>sTourID="&sTourID)


' --- Sets query for OLR Tournament display ---

sSQL = "SELECT Count(MemberID) AS Registered, SUM(MembRounds) AS PullsReceived" 	
sSQL = sSQL + "	FROM"
sSQL = sSQL + "	( SELECT RG.MemberID, MembRounds, Result"
sSQL = sSQL + "		FROM "&RegGenTableName&" AS RG"
sSQL = sSQL + "			LEFT JOIN"
sSQL = sSQL + "				(SELECT MemberID, Result" 
sSQL = sSQL + "					FROM "&RegPaymentTableName
sSQL = sSQL + "						WHERE LEFT(TourID,6)='"&LEFT(sTourID,6)&"' AND Result='0'"
sSQL = sSQL + "					GROUP BY MemberID, Result) AS PT"	
sSQL = sSQL + "			ON PT.MemberID=RG.MemberID"
sSQL = sSQL + "			LEFT JOIN"
sSQL = sSQL + "				(SELECT MemberID, Coalesce(SUM(FeeRounds),0) AS MembRounds"
sSQL = sSQL + "					FROM "&RegDetailTableName
sSQL = sSQL + "						WHERE LEFT(TourID,6)='"&sTourID&"'"
sSQL = sSQL + "					GROUP BY TourID, MemberID) AS RE"	
sSQL = sSQL + "			ON RE.MemberID=RG.MemberID"

sSQL = sSQL + "	WHERE RG.TourID='"&LEFT(sTourID,6)&"' AND Result='0') AS T"

SET rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable, 3, 3


EntriesExceedLimit=false
IF NOT rs.EOF THEN
	sPullsReceived=rs("PullsReceived")
	IF rs("PullsReceived")>=sMaxPulls AND sMaxPulls<> 0 THEN 
	'IF rs("PullsReceived")>=sMaxPulls AND sMaxPulls<> 1 THEN 
		EntriesExceedLimit=true
	END IF
ELSE
	eMailSubj = "ERROR - Tournament Data Not Found"
	sBody = "This error was generated from tournament "&sTourID&" in the EntriesExceedLimit subroutine in tools_registration.asp.  An End-of-File error was generated."
	MailMarkEmail eMailSubj, eMailBody

END IF


'response.write("<br>"&sSQL)
'response.write("<br>IN Tools_Registration.asp")
'response.write("<br>sPullsReceived = "&sPullsReceived)
'response.write("<br>sMaxPulls = "&sMaxPulls)
'response.write("<br>EntriesExceedLimit = "&EntriesExceedLimit)


'response.end




END FUNCTION




' -------------------------------
    Function ExistingEntry_OLD(sMemberID)
' -------------------------------

' --- Changed to OLD on 2/12/2013 in preparation to DELETE ---

Dim sHOHMemberID


' --- STEP - 1
' --- Find out MemberTypeID from People Table for this member ---
' --- If the next query is 3, then we've got a head of household and can jump to step #3 below.

Set SQLConnect = CreateObject("ADODB.Connection")
SQLConnect.Open Application("HQSQLConn")
Dim SQLString
SQLString = "SELECT PT.PersonID, PT.FirstName, PT.LastName, PT.[MemberTypeID] FROM Members AS PT"
SQLString = SQLString & " WHERE PT.PersonID=Cast(right('"&sMemberID&"',8) as Int)" 
Set RS = SQLConnect.Execute(SQLString)

'response.write("<br><br>")
'response.write(SQLString)

IF rs("MemberTypeID")=3 THEN sHOHMemberID=Int(right(sMemberID,8))



' --- STEP - 2
' --- If this returns no rows, then this is an individual, and you're done.  
' --- But if one row is returned, then the returned HOHMemberID will be the Person ID of the HOH of which this registrant is a sub-member.   
' --- If there were to be more than one row, that would represent an error, IMHO -- But probably worth checking for.

Set SQLConnect = CreateObject("ADODB.Connection")
SQLConnect.Open Application("HQSQLConn")

sSQL = "SELECT PT.PersonID AS HOHMemberID, PT.FirstName, PT.LastName, PT.[MemberTypeID]"
sSQL = sSQL + " FROM [Sub Members] AS SM, Members AS PT"
sSQL = sSQL + " WHERE PT.PersonID = SM.PrimaryPersonID and PT.MemberTypeID = 3 and SM.SubMemberPersonID = Cast(right('"&sMemberID&"',8) as Int)"
Set RS = SQLConnect.Execute(sSQL)

response.write("<br><br>")
response.write(sSQL)


DO WHILE NOT RS.eof 
	sHOHMemberID=RS("HOHMemberID")
	RS.movenext
LOOP

rs.close



' --- STEP 3
' --- Then taking that HOHMemberID from the above -- from (1) if that comes back with MemberTypeID of 3 
' ---   or the ID returned from the query in (2), if there is one 
' --- You'd go to the [Sub Members] table looking for the SubMemberPersonID's of all rows where the PrimaryMemberID = HOHMemberID.  
' --- The returned answer set will be the ADDITIONAL members of that household -- one of which may be the Registrant you started with
' --- if you went through step #2 above.


Set SQLConnect = CreateObject("ADODB.Connection")
SQLConnect.Open Application("HQSQLConn")

sSQL = "SELECT SM.SubMemberPersonID AS ThisID, PT.[First Name] AS FirstName, PT.[Last Name] AS LastName, PT.[MemberTypeID]" 
sSQL = sSQL + " FROM [Sub Members] AS SM, tblPeople AS PT"
sSQL = sSQL + " WHERE SM.PrimaryPersonID='"&sHOHMemberID&"' AND SM.SubMemberPersonID=PT.[Person ID]"
sSQL = sSQL + " UNION "
sSQL = sSQL + " SELECT PT.[Person ID] AS ThisID, PT.[First Name] AS FirstName, PT.[Last Name] AS LastName, PT.[MemberTypeID]" 
sSQL = sSQL + " FROM tblPeople AS PT"
sSQL = sSQL + " WHERE PT.[Person ID]='"&sHOHMemberID&"'"

response.write("<br><br>")
response.write(sSQL)
response.end

Set RS = SQLConnect.Execute(sSQL)


Dim MembNo
MembNo=0
DO WHILE NOT RS.eof 
	MembNo=MembNo+1
'	MembList(MembNo)=TRIM(RS("SubMemberPersonID"))
	MembList(MembNo)=TRIM(RS("ThisID"))
	MembListName(MembNo)=TRIM(Rs("FirstName"))&" "&Rs("LastName")
	RS.movenext
LOOP

TotQualifyingFamMemb=MembNo

' --- Only produces a TRUE result if there is "Family member" <> "current member" that has an EntryFee=FamilyEntryFee ---

Set rsRegGen=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT RG.MemberID, RG.EntryFee, MT.FirstName, MT.LastName" 
sSQL = sSQL + " FROM "&RegGenTableName&" AS RG, "&MemberTableName&" AS MT"
sSQL = sSQL + " WHERE Cast(right(RG.MemberID,8) AS INT) IN ('"&MembList(1)&"', '"&MembList(2)&"'"
IF TRIM(MembList(3))<>"" THEN sSQL = sSQL + " , '"&MembList(3)&"'"
IF TRIM(MembList(4))<>"" THEN sSQL = sSQL + " , '"&MembList(4)&"'"
IF TRIM(MembList(5))<>"" THEN sSQL = sSQL + " , '"&MembList(5)&"'"
IF TRIM(MembList(6))<>"" THEN sSQL = sSQL + " , '"&MembList(6)&"'"
IF TRIM(MembList(7))<>"" THEN sSQL = sSQL + " , '"&MembList(7)&"'"
IF TRIM(MembList(8))<>"" THEN sSQL = sSQL + " , '"&MembList(8)&"'"
sSQL = sSQL + ")"
sSQL = sSQL + " AND LEFT(TourID,6)='"&LEFT(sTourID,6)&"'"
sSQL = sSQL + " AND RG.MemberID=MT.PersonIDWithCheckDigit AND RG.MemberID <>'"&sMemberID&"'"

rsRegGen.open sSQL, sConnectionToTRATable, 3, 1

ExistingEntry=false
Session("sWhichFamilyMemberPaid")=""

Dim TotRegisteredFamMembers
TotRegisteredFamMembers=0

DO WHILE NOT RSRegGen.eof 
	TotRegisteredFamMembers=TotRegisteredFamMembers+1
	IF cdbl(RSRegGen("EntryFee"))=sTEntryFeeFamily THEN 
		ExistingEntry=true
		Session("sWhichFamilyMemberPaid")=rsRegGen("FirstName")&" "&rsRegGen("LastName")&" ("&rsRegGen("MemberID")&")"
	END IF
	RSRegGen.movenext
LOOP

Session("TotRegisteredFamMembers")=TotRegisteredFamMembers

rsRegGen.close


END Function



' -------------------------------
    Function ExistingEntry(sMemberID)
' -------------------------------


TestThis=false

Dim sHOHMemberID

'sTourID="12S999"

' ------------------------------------------------------------------------------------------------------------------------------------
' --- STEP - 1
' --- Find out MemberTypeID from People Table for this member ---
' --- If the query in STEP 1 is 3, then we've got a head of household and can jump to STEP #3 below.
' ------------------------------------------------------------------------------------------------------------------------------------

Set SQLConnect = CreateObject("ADODB.Connection")
SQLConnect.Open Application("HQSQLConn")
Dim SQLString
	
SQLString = "SELECT PT.[Person ID] AS PersonID, PT.[First Name] as FirstName, PT.[Last Name] as LastName, PT.[MemberTypeID] FROM Waterski.dbo.tblPeople AS PT"
SQLString = SQLString & " WHERE PT.[Person ID]=Cast(right('"&sMemberID&"',8) as Int)" 
Set RS = SQLConnect.Execute(SQLString)

IF rs("MemberTypeID")=3 THEN sHOHMemberID=Int(right(sMemberID,8))

IF TestThis=true THEN 
	response.write("<br><br>")
	response.write(SQLString)
	response.write("<br>EOF= ")
	response.write(rs.EOF)
	
	response.write("<br>rs(MemberTypeID)= "&rs("MemberTypeID"))
	response.write("<br>sHOHMemberID= "&sHOHMemberID)
END IF



' ------------------------------------------------------------------------------------------------------------------------------------
' --- STEP - 2
' --- If STEP 2 returns no rows, then this is an individual and not a member of a family membership...and you're done.  
' --- But if one row is returned, then the returned HOHMemberID will be the Person ID of the HOH of which this registrant is a sub-member.   
' --- If there were to be more than one row, that would represent an error, IMHO -- But probably worth checking for.
' ------------------------------------------------------------------------------------------------------------------------------------

Set SQLConnect = CreateObject("ADODB.Connection")
SQLConnect.Open Application("HQSQLConn")


sSQL = "SELECT PT.[Person ID] AS HOHMemberID, PT.[First Name] AS FirstName, PT.[Last Name] AS LastName, PT.[MemberTypeID]"
sSQL = sSQL + " FROM Waterski.dbo.[Sub Members] AS SM, Waterski.dbo.tblPeople AS PT"
sSQL = sSQL + " WHERE PT.[Person ID] = SM.PrimaryPersonID and PT.MemberTypeID = 3 and SM.SubMemberPersonID = Cast(right('"&sMemberID&"',8) as Int)"
Set RS = SQLConnect.Execute(sSQL)

IF TestThis=true THEN 
	response.write("<br><br>")
	response.write(sSQL)
	response.write("<br>EOF= ")
	response.write(rs.EOF)
END IF

DO WHILE NOT RS.eof 
	sHOHMemberID=RS("HOHMemberID")
	RS.movenext
LOOP

IF TestThis=true THEN 
	response.write("<br>sHOHMemberID= "&sHOHMemberID)
END IF

rs.close



' ------------------------------------------------------------------------------------------------------------------------------------
' --- STEP 3
' --- Then taking that HOHMemberID from the above -- from (1) if that comes back with MemberTypeID of 3 
' ---   or the ID returned from the query in (2), if there is one 
' --- You'd go to the [Sub Members] table looking for the SubMemberPersonID's of all rows where the PrimaryMemberID = HOHMemberID.  
' --- The returned answer set will be the ADDITIONAL members of that household -- one of which may be the Registrant you started with
' --- if you went through step #2 above.
' ------------------------------------------------------------------------------------------------------------------------------------

Set SQLConnect = CreateObject("ADODB.Connection")
SQLConnect.Open Application("HQSQLConn")

sSQL = "SELECT SM.SubMemberPersonID AS ThisID, PT.[First Name] AS FirstName, PT.[Last Name] AS LastName, PT.[MemberTypeID]" 
sSQL = sSQL + " FROM Waterski.dbo.[Sub Members] AS SM, Waterski.dbo.tblPeople AS PT"

IF sHOHMemberID="" THEN
  	sSQL = sSQL + " WHERE SM.PrimaryPersonID='987654321'"
ELSE
	sSQL = sSQL + " WHERE SM.PrimaryPersonID='"&sHOHMemberID&"'"
END IF

sSQL = sSQL + " AND SM.SubMemberPersonID=PT.[Person ID]"
sSQL = sSQL + " UNION "
sSQL = sSQL + " SELECT PT.[Person ID] AS ThisID, PT.[First Name] AS FirstName, PT.[Last Name] AS LastName, PT.[MemberTypeID]" 
sSQL = sSQL + " FROM Waterski.dbo.tblPeople AS PT"

IF sHOHMemberID="" THEN
  	sSQL = sSQL + " WHERE PT.[Person ID]='456456456'"
ELSE
  	sSQL = sSQL + " WHERE PT.[Person ID]='"&sHOHMemberID&"'"
END IF


Set RS = SQLConnect.Execute(sSQL)

IF TestThis=true THEN 
	response.write("<br><br>")
	response.write(sSQL)
	response.write("<br>EOF= ")
	response.write(rs.EOF)

	response.write("<br><br>Test sHOHMemberID= "&sHOHMemberID)
	response.write("<br>4-")
	response.write(sHOHMemberID="")
	response.write("<br><br>")
END IF


Dim MembNo
MembNo=0
DO WHILE NOT RS.eof AND MembNo<8
	MembNo=MembNo+1
'	MembList(MembNo)=TRIM(RS("SubMemberPersonID"))
	MembList(MembNo)=TRIM(RS("ThisID"))
	MembListName(MembNo)=TRIM(Rs("FirstName"))&" "&Rs("LastName")

	IF TestThis=true THEN 
		response.write("<br>MembNo= "&MembNo)
		response.write("<br>MembList(MembNo)= "&MembList(MembNo))
		response.write("<br>MembList(MembNo)= "&MembList(MembNo))
	END IF
	'IF MembNo>8 THEN response.end

	RS.movenext
LOOP

TotQualifyingFamMemb=MembNo



IF TestThis=true THEN 
	response.write("<br>")
	response.write("TotQualifyingFamMemb="&TotQualifyingFamMemb)
END IF

' --- Only produces a TRUE result if there is "Family member" <> "current member" that has an EntryFee=FamilyEntryFee ---

Set rsRegGen=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT RG.MemberID, RG.EntryFee, MT.FirstName, MT.LastName" 

' --- Changed 2-12-2013 ---
sSQL = sSQL + " FROM "&RegGenTableName&" AS RG, "&MemberShortTableName&" AS MT"
sSQL = sSQL + " WHERE Cast(right(RG.MemberID,8) AS INT) IN ('"&MembList(1)&"', '"&MembList(2)&"'"
IF TRIM(MembList(3))<>"" THEN sSQL = sSQL + " , '"&MembList(3)&"'"
IF TRIM(MembList(4))<>"" THEN sSQL = sSQL + " , '"&MembList(4)&"'"
IF TRIM(MembList(5))<>"" THEN sSQL = sSQL + " , '"&MembList(5)&"'"
IF TRIM(MembList(6))<>"" THEN sSQL = sSQL + " , '"&MembList(6)&"'"
IF TRIM(MembList(7))<>"" THEN sSQL = sSQL + " , '"&MembList(7)&"'"
IF TRIM(MembList(8))<>"" THEN sSQL = sSQL + " , '"&MembList(8)&"'"
sSQL = sSQL + ")"
sSQL = sSQL + " AND LEFT(TourID,6)='"&LEFT(sTourID,6)&"'"

' --- Changed 2-12-2013 ---
sSQL = sSQL + " AND CAST(RIGHT(RG.MemberID,8) AS INT)=MT.PersonID AND RG.MemberID <>'"&sMemberID&"'"
'sSQL = sSQL + " AND RG.MemberID=MT.PersonIDWithCheckDigit AND RG.MemberID <>'"&sMemberID&"'"
'response.write("<br><br>HERE "&sSQL)
'response.end

rsRegGen.open sSQL, sConnectionToTRATable, 3, 1

ExistingEntry=false
Session("sWhichFamilyMemberPaid")=""

Dim TotRegisteredFamMembers
TotRegisteredFamMembers=0

'response.write("<br><br>sTEntryFeeFamily="&sTEntryFeeFamily)

DO WHILE NOT RSRegGen.eof 
	TotRegisteredFamMembers=TotRegisteredFamMembers+1
	'response.write("<br>TotRegisteredFamMembers="&TotRegisteredFamMembers&" - "&rsRegGen("FirstName")&" "&rsRegGen("LastName"))
	IF cdbl(RSRegGen("EntryFee"))=sTEntryFeeFamily THEN 
			ExistingEntry=true
			Session("sWhichFamilyMemberPaid")=rsRegGen("FirstName")&" "&rsRegGen("LastName")&" ("&rsRegGen("MemberID")&")"
			'response.write("<br>YES - "&rsRegGen("FirstName")&" "&rsRegGen("LastName")&" - "&cdbl(RSRegGen("EntryFee")))
	END IF
	RSRegGen.movenext
LOOP

Session("TotRegisteredFamMembers")=TotRegisteredFamMembers

rsRegGen.close


END Function







' ---------------------------
    SUB SendTourFullEMail
' ---------------------------


Dim DateNow, TimeNow
Dim eMailSubj, eMailFrom, eMailTo, eMailCC, eMailBCC, eMailBody

DateNow = Date
TimeNow = Time

ebody = "<html>"
ebody = ebody & "<head>"
ebody = ebody & "<title>Notice to Registrar or LOC</title>"
ebody = ebody & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
ebody = ebody & "</head>"
ebody = ebody & "<body bgcolor=""#FFFFFF"" text=""#000000"">"
ebody = ebody & "<div align=""center"">"


ebody = ebody & "<TABLE BORDER=4 ALIGN=CENTER CELLPADDING=3 CELLSPACING=0 BGCOLOR="&TableColor1&" width=60% >"
ebody = ebody & "<TR>"
ebody = ebody & "<TD BGCOLOR=red><center><font face="&font1&" color=#FFFFFF size=4><b>Notice to Registrar or LOC</b></font></TD>"
ebody = ebody & "</TR>"
 
ebody = ebody & "<TR>"
ebody = ebody & "<TD VALIGN=top>"


ebody = ebody & "<table border=""0"" cellspacing=""0"" cellpadding=""3"" width=""100%"">"
ebody = ebody & "<tr>"


ebody = ebody & "<td Align=center>"	
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>This message was generated because the number of registered pulls/rides has exceeded <br>the limit set for the tournament/clinic listed below.</b></font>"

ebody = ebody & "<br><br>"


ebody = ebody & "<font color="&TextColor2&" face="&font1&" size=4><b>"&sTourName&"</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>SanctionID = </b></font><font color="&TextColor2&" face="&font1&" size=2>"&sTourID&"</font>"
ebody = ebody & "<br>"

ebody = ebody & "<font face="&font1&" size=2><b>Date = </b></font><font color="&TextColor2&" face="&font1&" size=2>"&sTDateS&" to "&sTDateE&"</font></b>"

ebody = ebody & "<br><br>"
ebody = ebody & "<font face="&font1&" size=2><b>The most recent entry received was from:</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "<font color="&TextColor2&" face="&font1&" size=4><b>"&sFirstName&" "&sLastName&"</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>MemberID = </b></font><font color="&TextColor2&" face="&font1&" size=2>"&sMemberID&"</font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>Date/Time: </b></font><font color="&TextColor2&" face="&font1&" size=2>"&DateNow&" -- "&TimeNow&"</font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>Maximum Pulls/Rides Setting: </b></font><font color="&TextColor2&" face="&font1&" size=2>"&sMaxPulls&"</font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>PAID Rides/Pulls Received: </b></font><font color="&TextColor2&" face="&font1&" size=2>"&sPullsReceived&"</font>"
ebody = ebody & "</td></tr>"

ebody = ebody & "<tr><td Align=left>"	
ebody = ebody & "<br><br>"
ebody = ebody & "<font face="&font1&" size=2><b>To <u>Stop Accepting Online Entries</u> and disable the 'Register Online' button on the tournament list:</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;1) Login from the 'Events & Registration: Event Search or Register: Registrar Login' on main menu.</font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Or click on the link: "
ebody = ebody & "<a href=http://usawaterski.org/rankings/view-tournamentsHQ.asp?process=admcode&sSendingPage=NEW&sl=on&tr=on&ju=on&wb=on&ws=on&wu=on&hy=on&sTourSportGroup=&sTourRange=1>Registrar Login</a></font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2) Select your tournament and enter the correct AdminCode when prompted.</font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;3) Press the Disable OnLine Registration Button. </font>"
ebody = ebody & "</td></tr>"

ebody = ebody & "<tr><td Align=left>"	
ebody = ebody & "<br><br>"
ebody = ebody & "<font face="&font1&" size=2><b>To <u>Modify the Maximum Pulls/Rides</u> setting in your OLR setup profile to stop these emails:</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;1) Go to the 'Sanctions Management System'</font>"
'ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2> by clicking on the link: "
ebody = ebody & "<a href=http://www.usawaterski.org/sanctions>Sanctions Management</a></font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;2) Select Edit Existing Sanction and enter your TournAppID (sanction #) when prompted.</font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;3) Click the Event Fees header and change the Max Pulls. One (1) pull = One (1) skier performance.<br>  For more than 500 pulls select unlimited.</font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;4) Save the form.</font>"
ebody = ebody & "<br><br>"

ebody = ebody & "</td></tr>"

ebody = ebody & "<tr><td Align=Center>"	
ebody = ebody & "<font face="&font1&" size=2>&nbsp;&nbsp;&nbsp;&nbsp;Your changes will be reflected immediately and you may re-enable the 'Register Online' button <br>on the tournament listing at any time.</font>"
ebody = ebody & "<br>"

ebody = ebody & "</center>"
ebody = ebody & "<br>"
ebody = ebody & "</td></tr>"

ebody = ebody & "</TABLE>"

ebody = ebody & "</TD></TR>"
ebody = ebody & "</TABLE>"

ebody = ebody & "</div>"
ebody = ebody & "</body>"
ebody = ebody & "</html>"


eMailFrom="competition@usawaterski.org"
eMailBCC=MarksEmail
eMailSubj=" "&sTourID&" - "&sTourName&" - Registration Limit Has Been Reached" 
sTest="off"	' --- on/off

SendTourEmail eMailSubj, eBody, eMailFrom, eMailBCC, sTest


END SUB





' ----------------------------------------------------------------------
  SUB SendTourEmail (eMailSubj, eMailBody, eMailFrom, eMailBCC, sTest)
' ----------------------------------------------------------------------


' --- Dimension and define all the email related variables ---

Dim SendAddress, HQWaiverEmail, MembWaiverEmail

eMailTo = sTsEmail
eEmailCC=""

' --- First on COPY TO --- 
IF TRIM(sTRegistrarEmail)<>"" AND TRIM(sTRegistrarEmail)<>TRIM(sTsEmail) THEN eMailCC = sTRegistrarEmail

' --- Second on COPY TO ---
IF TRIM(sTourEmail)<>"" AND TRIM(sTourEmail)<>TRIM(sTsEmail) AND TRIM(sTourEmail)<>TRIM(sRegistrarEmail) THEN
		IF TRIM(sCC)<>"" THEN 
				eMailCC=eMailCC&", "&sTourEmail
		ELSE
				eMailCC=sTourEmail
		END IF
END IF

' --- Third on COPY TO ---
IF TRIM(sTDirEmail)<>"" AND TRIM(sTDirEmail)<>sTsEmail AND TRIM(sTDirEmail)<>TRIM(sTRegistrarEmail) AND TRIM(sTDirEmail)<>TRIM(sTourEmail) THEN
		IF TRIM(sCC)<>"" THEN 
				eMailCC=eMailCC&", "&sTDirEmail
		ELSE
				eMailCC=sTDirEmail
		END IF
END IF

' --- TEST ---
'sTest="off"
IF sTest="on" THEN
		eMailTo=marksemail
		eMailCC=""
		'eMailBCC=""
END IF





' ---------------------------------------------------------------
' --- Now assign the components to the standard email objects ---
' ---------------------------------------------------------------

SetupEmailService

objMessage.Subject = eMailSubj
objMessage.From = eMailFrom
objMessage.To = eMailTo
objMessage.cc = eMailCC
IF LCASE(eMailBCC)<>"none" THEN objMessage.bcc = eMailBCC
objMessage.HTMLBody = eMailBody
 
 ' --- Finally send the message, and then clear that object
IF TRIM(SendAddress)<>"" THEN
		objMessage.Send
END IF
set objMessage = Nothing


END SUB



' -------------------------------------------------------------
  SUB MailMarkEmail (eMailSubj, eMailBody)
' -------------------------------------------------------------

eMailFrom"competition@usawaterski.org"
eMailTo = marksemail
eMailCC=""
eMailBCC=""

' ---------------------------------------------------------------
' --- Now assign the components to the standard email objects ---
' ---------------------------------------------------------------

SetupEmailService

objMessage.Subject = eMailSubj
objMessage.From = eMailFrom
objMessage.To = eMailTo
objMessage.cc = eMailCC
IF LCASE(eMailBCC)<>"none" THEN objMessage.bcc = eMailBCC
objMessage.HTMLBody = eMailBody
 
 ' --- Finally send the message, and then clear that object
IF TRIM(SendAddress)<>"" THEN
		objMessage.Send
END IF
set objMessage = Nothing


END SUB




' -------------------------------------------------------------------------------------------------------
    SUB SendSPECIALWaiverEmail_Tools (sSpecialWaiverCode, sSpecialWaiverHeadline, sSpecialReleaseBannerText)
' -------------------------------------------------------------------------------------------------------

' --- Gets Waiver Info from RegGenTable --- 
SET rsRegTemp=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT * FROM "&RegTempTableName
sSQL = sSQL + " WHERE Left(TourID,6) = '"&SQLClean(left(sTourID,6))&"' AND MemberID = '"&sMemberID&"'"
rsRegTemp.open sSQL, SConnectionToTRATable, 3, 3

sWaiverCode = rsRegTemp("WaiverCode")
sSignWaiver = SQLClean(rsRegTemp("SignWaiver"))

rsRegTemp.close


DefineTourVariables_New
DefineMemberVariables_Tools

'response.write("<br> Line 2060 Tools_Registration.asp - sMemberID = "&sMemberID)
'response.write("<br>sFirstName = "&sFirstName)


' --- New 4-28-2013 - Gets SPECIAL WAIVER info from table based on SiteID rather than hard coding specific tournaments ---
'Dim swaiverSQL, sSpecialWaiverHeadline, sSpecialReleaseBannerText
swaiverSQL = "SELECT SpecialWaiverCode, SpecialWaiverHeadline, SpecialReleaseBannerText, SpecialWaiver_LOCEmail, SpecialWaiver_IncludeHQ"
swaiverSQL = swaiverSQL + " FROM usawsrank.TourExtras TE"
swaiverSQL = swaiverSQL + " JOIN sanctions.dbo.TSchedul AS TS"
swaiverSQL = swaiverSQL + "   ON SiteID=TS.TSiteID"
swaiverSQL = swaiverSQL + " WHERE LEFT(TS.TournAppID,6)='"&LEFT(sTourID,6)&"'"

Set rswaiver=Server.CreateObject("ADODB.recordset")
rswaiver.open swaiverSQL, sConnectionToTRATable, 3, 1

testwaiver=false
IF testwaiver=true THEN
		Response.write("<br><br><br>Found = ")
		response.write(NOT(rswaiver.eof))
		response.write("<br>rswaiver(SpecialWaiverHeadline) = "&rswaiver("SpecialWaiverHeadline"))
		response.write("<br>rswaiver(SpecialWaiver_LOCEmail) = "&rswaiver("SpecialWaiver_LOCEmail"))
		response.write("<br>sMemberID = "&sMemberID)
		response.write("<br>sTourID = "&sTourID)
END IF

IF NOT(rswaiver.EOF) THEN
		sSpecialWaiverCode=rswaiver("SpecialWaiverCode")
		sSpecialWaiverHeadline=rswaiver("SpecialWaiverHeadline")
		sSpecialReleaseBannerText=rswaiver("SpecialReleaseBannerText")
		LOCSpecialWaiverEmail=rswaiver("SpecialWaiver_LOCEmail")
		SpecialWaiver_IncludeHQ=rswaiver("SpecialWaiver_IncludeHQ")
END IF




ebody = "<html>"
ebody = ebody & "<head>"
ebody = ebody & "<title>Waiver and Release</title>"
ebody = ebody & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
ebody = ebody & "</head>"
ebody = ebody & "<body bgcolor=""#FFFFFF"" text=""#000000"">"
ebody = ebody & "<div align=""center"">"


ebody = ebody & "<TABLE BORDER=4 ALIGN=CENTER CELLPADDING=3 CELLSPACING=0 BGCOLOR="&TableColor1&" width=85% >"
ebody = ebody & "<TR>"
ebody = ebody & "<TD BGCOLOR=orange><center><font face="&font1&" color=#FFFFFF size=4><b>"&sSpecialReleaseBannerText&"</b></font></TD>"
ebody = ebody & "</TR>"
 
ebody = ebody & "<TR>"
ebody = ebody & "<TD VALIGN=top>"


ebody = ebody & "<table border=""0"" cellspacing=""0"" cellpadding=""3"" width=""100%"">"
ebody = ebody & "<tr>"


ebody = ebody & "<td Align=center>"	
ebody = ebody & "<font face="&font1&" size=4 ><b>"&sSpecialWaiverHeadline&"</b></font><br>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>"&subTitle&"</b></font>"
ebody = ebody & "<br><br>"
ebody = ebody & "<font face="&font1&" color="&TextColor2&" size=3><b>"&sTourName&"</font></b>"
ebody = ebody & "<br><br>"
ebody = ebody & "<font face="&font1&" size=2><b>MemberID = </font><font color="&TextColor2&" face="&font1&" size=2>"&sMemberID&""
ebody = ebody & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="&TextColor1&" face="&font1&" size=2>Participant:</font>"
ebody = ebody & "<font color="&TextColor2&" face="&font1&" size=2>&nbsp;&nbsp;"&sFirstname&"&nbsp;"&sLastName&"</font></b><br>"

ebody = ebody & "</center>"
ebody = ebody & "<br>"
ebody = ebody & "</td></tr>"


ebody = ebody & "<td Align=left>"	
ebody = ebody & "<P><font color="&TextColor1&" size=1 face="&font1&">"
	
Set objfso = CreateObject("Scripting.FileSystemObject")

' --- Formerly ReleaseVersion
IF objfso.FileExists(PathtoWaivers & "\waiver-"&sSpecialWaiverCode&".txt") THEN
	SET objstream=objFSO.opentextfile(PathtoWaivers & "\waiver-"&sSpecialWaiverCode&".txt")

	IF NOT objstream.atendofstream THEN
		DO WHILE not objstream.atendofstream
			'response.write(objstream.readline)
			ebody = ebody & objstream.readline
			ebody = ebody & "<br>"
		LOOP
	END IF

END IF

objstream.close 

ebody = ebody & "</font></P>"
ebody = ebody & "</td></tr>"

ebody = ebody & "<tr>"	
ebody = ebody & "<td Align=center>"	
		
ebody = ebody & "<br>"
ebody = ebody & "<font color="&TextColor3&" face="&font1&" size=3><b>I agree to be fully responsible for my conduct at the tournament and/or for the conduct of the minor on whose behalf I sign.</b></font>"

ebody = ebody & "<br><br>"
ebody = ebody & "<font color="&TextColor1&" face="&font1&" size=2><b>Date Accepted:&nbsp;&nbsp</font><font color="&TextColor2&" face="&font1&" size=2>"&DATE&"</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "<font color="&TextColor1&" face="&font1&" size=2><b>Accepted By:&nbsp;&nbsp</font><font color="&TextColor2&" face="&font1&" size=2>"&sSignWaiver&"</b></font>"

ebody = ebody & "</td></tr>"

ebody = ebody & "</form>"

ebody = ebody & "</td>"
ebody = ebody & "<br>"
ebody = ebody & "</tr>"
ebody = ebody & "</TABLE>"



ebody = ebody & "</TD></TR>"
ebody = ebody & "</TABLE>"


' ------------------------------------------------------------
' --- Dimension and define all the email related variables ---
' ------------------------------------------------------------

Dim eMailSubj, eMailFrom, eMailTo, eMailCC, eMailBCC, eMailBody
Dim SendAddress, HQWaiverEmail, MembWaiverEmail

HQWaiverEmail="competition@usawaterski.org"

sWaiverEmail=true




' ---------------------------------------------------
' --- Get all the information from Password table ---
' ---------------------------------------------------
set rsPW=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT TOP 1 * FROM "&RegPWTableName&" WHERE MemberID = "&sqlclean(sMemberID)
rsPW.open sSQL, sConnectionToTRATable, 3, 1

IF NOT rsPW.eof THEN 
		MembWaiverEmail=rsPW("email")
		IF sWaiverEmail=true THEN SendAddress=MembWaiverEmail
		eMailSubj = "SPECIAL WAIVER & RELEASE  TourID: "&Session("sTourID")&" - Member: "&sFirstName&" "&sLastName
ELSE
		SendAddress = HQWaiverEmail
		eMailSubj = "SPECIAL WAIVER & RELEASE  TourID: "&Session("sTourID")&" - Member: "&sFirstName&" "&sLastName&" - Admin Override"
END IF

eMailTo = SendAddress


eMailCC=""
IF SpecialWaiver_IncludeHQ="Y" THEN eMailCC = HQWaiverEmail
IF TRIM(eMailCC)="" THEN
		eMailCC=LOCSpecialWaiverEmail
ELSE
		eMailCC=eMailCC+"; "&LOCSpecialWaiverEmail
END IF		

IF sSpecialWaiverEmailMC=true THEN 
		eMailBCC = " "&marksemailaddress
END IF

eMailFrom = ""&HQWaiverEmail
eMailBody = ebody	

testwaiver=false
IF testwaiver=true THEN
		response.write("<br><br>eMailTo = "&eMailTo)
		response.write("<br>MembWaiverEmail = "&MembWaiverEmail)
		response.write("<br>sWaiverEmail = "&sWaiverEmail)
		response.write("<br>LOCSpecialWaiverEmail = "&LOCSpecialWaiverEmail)
		response.write("<br>eMailSubj = "&eMailSubj)
END IF

response.write("<br><br>"&eMailBody)




' ---------------------------------------------------------------
' --- Now assign the components to the standard email objects ---
' ---------------------------------------------------------------

SetupEmailService

objMessage.Subject = eMailSubj
objMessage.From = eMailFrom
objMessage.To = eMailTo
objMessage.cc = eMailCC
objMessage.bcc = eMailBCC
objMessage.HTMLBody = eMailBody
 
 ' --- Finally send the message, and then clear that object
IF TRIM(SendAddress)<>"" THEN
		objMessage.Send
END IF
set objMessage = Nothing



END SUB




' ---------------------------------------------------
  SUB SendErrorEmailToMark (ErrorSubject, ErrorSQL)
' ---------------------------------------------------

' ------------------------------------------------------------
' --- Dimension and define all the email related variables ---
' ------------------------------------------------------------

Dim eMailSubj, eMailFrom, eMailTo, eMailCC, eMailBCC, eMailBody
Dim SendAddress, HQErrorEmail


sBannerText = "Information for Mark Crone"
HQErrorEmail="competition@usawaterski.org"

sErrorEmail=true
SendAddress = marksemailaddress
eMailSubj = ErrorSubject
eMailTo = SendAddress
eMailCC=""
eMailBCC=""
eMailFrom = ""&HQErrorEmail


ebody = "<html>"
ebody = ebody & "<head>"
ebody = ebody & "<title>Waiver and Release</title>"
ebody = ebody & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
ebody = ebody & "</head>"
ebody = ebody & "<body bgcolor=""#FFFFFF"" text=""#000000"">"
ebody = ebody & "<div align=""center"">"


ebody = ebody & "<TABLE BORDER=4 ALIGN=CENTER CELLPADDING=3 CELLSPACING=0 BGCOLOR="&TableColor1&" width=85% >"
ebody = ebody & "<TR>"
ebody = ebody & "<TD BGCOLOR=orange><center><font face="&font1&" color=#FFFFFF size=4><b>"&sBannerText&"</b></font></TD>"
ebody = ebody & "</TR>"
 
ebody = ebody & "<TR>"
ebody = ebody & "<TD VALIGN=top>"
ebody = ebody & ErrorSQL
ebody = ebody & "</TD>"
ebody = ebody & "</TR>"
ebody = ebody & "</TABLE>"
ebody = ebody & "</div>"
ebody = ebody & "</body>"
ebody = ebody & "</html>"

eMailBody = ebody	


testerror=false
IF testerror=true THEN
		response.write("<br><br>eMailTo = "&eMailTo)
		response.write("<br>eMailSubj = "&eMailSubj)
END IF

'response.write("<br><br>"&eMailBody)



' ---------------------------------------------------------------
' --- Now assign the components to the standard email objects ---
' ---------------------------------------------------------------

SetupEmailService

objMessage.Subject = eMailSubj
objMessage.From = eMailFrom
objMessage.To = eMailTo
objMessage.cc = eMailCC
objMessage.bcc = eMailBCC
objMessage.HTMLBody = eMailBody
 
 ' --- Finally send the message, and then clear that object
IF TRIM(SendAddress)<>"" THEN
		objMessage.Send
END IF
set objMessage = Nothing



END SUB


  





' -----------------------------
   SUB DefineMemberVariables_Tools
' -----------------------------

	
'response.end


' --- Changed 2-12-2013 --- 
NewDbSchema="Y"
IF NewDbSchema="Y" THEN

	sSQL = "SELECT TOP 1 FirstName, LastName, City, State, Sex, Phone, MembershipTypeCode,"
	sSQL = sSQL + " Birthdate, Email, EffectiveTo,"  
	sSQL = sSQL + " Description,"
	sSQL = sSQL + " coalesce(MembershipTypeID,0) AS MembershipTypeID, "
	sSQL = sSQL + " coalesce(CanSkiInTournaments,0) AS CanSkiInTournaments, "
	sSQL = sSQL + " coalesce(CanSkiInGRTournaments,0) AS CanSkiInGRTournaments, "
	sSQL = sSQL + " coalesce(TypeCode,'XXX') AS TypeCode"

	sSQL = sSQL + " FROM "&MemberLiveTableName&" MT"
	sSQL = sSQL + " LEFT JOIN "&MemberTypeTableName&" MTT ON MTT.MembershipTypeID = MT.MembershipTypeCode"
	sSQL = sSQL + " WHERE PersonID = cast(right("&sqlclean(sMemberID)&",8) AS INTEGER)"
'  sSQL = sSQL + " WHERE PersonID = "& CAST(right(sMemberID,8) AS integer)&""

'response.write("<br>sMemberID= "&sMemberID)
tu=2
IF tu=1 AND sMemberID="000001151" THEN
		response.write("Line 1975 - <br>"&sSQL)
END IF
'response.end

	'ELSE
		
	'	sSQL = "SELECT TOP 1 FirstName, LastName, City, State, Sex, Phone, MembershipTypeCode,"
	'	sSQL = sSQL + " Birthdate, Email, EffectiveTo,"  
	'	sSQL = sSQL + " Description,"
	'	sSQL = sSQL + " coalesce(MembershipTypeID,0) AS MembershipTypeID, "
	'	sSQL = sSQL + " coalesce(CanSkiInTournaments,0) AS CanSkiInTournaments, "
	'	sSQL = sSQL + " coalesce(CanSkiInGRTournaments,0) AS CanSkiInGRTournaments, "
	'	sSQL = sSQL + " coalesce(TypeCode,'XXX') AS TypeCode"

	'	sSQL = sSQL + " FROM "&MemberTableName
	'	sSQL = sSQL + " LEFT JOIN "&MemberTypeOLRTableName&" ON "&MemberTypeOLRTableName&".MembershipTypeID = "&MemberTableName&".MembershipTypeCode"
	'	sSQL = sSQL + " WHERE PersonIDwithCheckDigit = '"&sqlclean(sMemberID)&"'"

END IF


	set rsMemb=Server.CreateObject("ADODB.recordset")
	rsMemb.open sSQL, sConnectionToTRATable, 3, 1

	sFirstName = SQLClean(rsMemb("FirstName"))
	sLastName = SQLClean(rsMemb("LastName"))
	sFullName = SQLClean(rsMemb("FirstName")&" "&rsMemb("LastName"))
	sMembCity = SQLClean(rsMemb("City"))
	sMembState = rsMemb("State")
	sMembSex = rsMemb("Sex")
	sMembPhone = rsMemb("Phone")
	sMembBirth = rsMemb("Birthdate")
	sMembEmail = rsMemb("Email")
	sEffectiveto = rsMemb("Effectiveto")

	sMembTypeID = rsMemb("MembershipTypeID")
	sCanSkiTour = rsMemb("CanSkiInTournaments")
	sCanSkiGRTour = rsMemb("CanSkiInGRTournaments")
	sMembTypeCode = rsMemb("TypeCode")
	sTypeDesc = rsMemb("Description")

'response.write("<br> Line 2313 Tools_Registration.asp - sMemberID = "&sMemberID)
'response.write("<br>sFirstName = "&sFirstName)

TestExpireDate=2
IF TestExpireDate=1 AND sMemberID="000001151" THEN
    ' --- Test Effective Date
		sEffectiveto="4/1/2013"


		response.write("<br>sEffectiveto= "&sEffectiveto)
		response.write("<br>TRIM(Session(sMembCanSkiText))=null - ")
		response.write(TRIM(Session("sMembCanSkiText"))="")
		response.write("<br>sCanSkiTour = 0 - ")
		response.write(sCanSkiTour = 0)
		response.write("<br>sCanSkiTour= "&sCanSkiTour)
		response.write("<br><br>Session(sMembCanSkiText) = "&TRIM(Session("sMembCanSkiText")))
		response.write("<br><br>")
		response.write(DateDiff("d", sEffectiveto, sTDateE))
		response.write("<br>DateDiff(d, sEffectiveto, sTDateE) > 0 - ")
		response.write(DateDiff("d", sEffectiveto, sTDateE) > 0)
		
		response.write("<br>TestExpireDate<>1 - ")
		response.write(TestExpireDate<>1)
END IF



 
' -------------------------------------------------------------
' ------- Checks competition status from Member file(s) -------
' -------------------------------------------------------------



' *****************************************************************************************************************************
' **************  IMPORTANT - Renewal or Upgrade code has NOT been changed to include sEnable variables for GR etc ***********
' *****************************************************************************************************************************


'Session("sCanSkiTour") = rsMemb("CanSkiInTournaments")

' --- Has not previously been set and can ski then set to OK in first two positions ---
IF TRIM(Session("sMembCanSkiText"))="" AND sCanSkiTour = 1 THEN  
		Session("sMembCanSkiText")="OK - "&sMembTypeCode&" - "&sTypeDesc
		Session("sMembCanSkiColor")=TextColor2
		Session("sEnableGR")="Y"
		Session("sEnableStd")="Y"
		Session("sEnableRec")="Y"

' --- TEMPORARY FIX for GR Membership ---
ELSEIF TRIM(Session("sMembCanSkiText"))="" AND sCanSkiGRTour = 1  AND (sGRTournament=true OR sGRFunDay=true) THEN
		Session("sMembCanSkiText")="OK - "&sMembTypeCode&" - "&sTypeDesc
		Session("sMembCanSkiColor")=TextColor2
		Session("sEnableGR")="Y"

' --- Has not previously been set and CANNOT ski then set to upgrade condition ---
ELSEIF TRIM(Session("sMembCanSkiText"))="" AND sCanSkiTour = 0 THEN
		Session("sMembCanSkiText")=sMembTypeCode&" - Competition Upgrade Required"
		Session("sMembCanSkiColor")="red"
END IF




'response.end




' ---- Needs both Member and Tournament information to define sMembAge  ----
sMembAge = AgeAtDate_New(sTDateS, sMemberID)		' Function finds Member Age
'IF sMemberID = "000001151" THEN sMembAge = 12


' --------------------------------------------------------
' ------- Sets the appropriate waiver based on age -------
' --------------------------------------------------------

Session("sMembAge") = sMembAge


' ------------------------------------------------------------------------------------
' ---  Checks End Date of tournament against Expiration Date of membership record  ---
' ------------------------------------------------------------------------------------
' --- Variable that displays in competition status area of OLR form ---
'Session("sExpirationStatusText")

yo=1
IF yo=2 AND sMemberID="000001151" THEN
	   Response.write("<br>Line 2116 - Session(sExpirationStatusText) = ")
	   Response.write(Session("sExpirationStatusText"))
END IF


IF Session("sExpirationStatusText")="" AND DateDiff("d", sEffectiveto, sTDateE) <= 0  THEN
		Session("sExpirationStatusText")="OK - "&sEffectiveto
		Session("sExpirationStatusColor")="blue"
ELSEIF Session("sExpirationStatusText")= "" AND DateDiff("d", sEffectiveto, sTDateE) > 0 THEN 
		Session("sExpirationStatusText")="Renewal - Expire: "&sEffectiveto
		Session("sExpirationStatusColor")="red"
' --- Added 4-2-2013 ---
ELSEIF DateDiff("d", sEffectiveto, sTDateE) <= 0  THEN
		Session("sExpirationStatusText")="OK - "&sEffectiveto
		Session("sExpirationStatusColor")="blue"
END IF

rsMemb.Close


IF TestExpireDate=1 AND sMemberID="000001151" AND DateDiff("d", sEffectiveto, sTDateE) < 0 THEN
		response.write("<br> - Line 2121 IF")
		Session("sExpirationStatusText")="OK - "&sEffectiveto	
		Session("sExpirationStatusColor")="blue"
ELSEIF TestExpireDate=1 AND sMemberID="000001151" AND DateDiff("d", sEffectiveto, sTDateE) > 0 THEN	
		response.write("<br> - Line 2124 ELSEIF")	
		Session("sExpirationStatusText")="Renewal - Expire: "&sEffectiveto		
		Session("sExpirationStatusColor")="red"
END IF	



' ------------------------------------------------------------------------------
' ------   Determines if Bio has been completed to indicate on display   -------
' ------------------------------------------------------------------------------

SET rsBio=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT MemberID, LastUpdate FROM "&BioTableName
sSQL = sSQL + " WHERE MemberID = '"&sMemberID&"'"
rsBio.open sSQL, SConnectionToTRATable, 3, 3

IF rsBio.eof THEN
		sBioDone = "N"
		Session("sBioDoneText")="InComplete"
		Session("sBioDoneTextColor")="red"
ELSEIF Year(rsBio("LastUpdate"))<Year(Date) THEN
		sBioDone = "N"
		Session("sBioDoneText")="Out of Date"
		Session("sBioDoneTextColor")="red"
ELSE
		sBioDone = "Y"
		Session("sBioDoneText")="Complete"
		Session("sBioDoneTextColor")="blue"
END IF

rsBio.close


END SUB




' -----------------------
   SUB ValidatePayPal
' -----------------------

' --- Moved to Tools_Registration.asp on 1-5-2014 ---

' --- Checks for latest order for this memberID and tourID ---
SET rsPayLog=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT Count(*) AS RCount, MAX(OrderNo) AS MaxOrder FROM "&RegPaymentTableName
sSQL = sSQL + " WHERE MemberID='"&sMemberID&"' AND TourID='"&sTourID&"'"
rsPayLog.open sSQL, sConnectionToTRATable, 3, 1

sPaymentResult=""
IF NOT rsPayLog.eof THEN
	IF TRIM(rsPayLog("MaxOrder"))<>"" THEN
			MaxOrder=rsPayLog("MaxOrder")
	ELSE
			MaxOrder=999999999
	END IF

	' --- If the MaxOrder from RegPaymentLog is the same as sOrderNo returned from the PayPal process them the sPaymentResult="0"
	IF cdbl(MaxOrder)=cdbl(sOrderNo) THEN sPaymentResult="0"
	sPayAmount = sTotalFormFees-sTotalPreviousPayments

END IF

END SUB


' -----------------------------
 SUB UpdatePaymentTransaction
' -----------------------------

' --- Moved to Tools_Registration.asp on 1-5-2014 ---

IF sMemberID="000001151" AND TestMode="yes"  THEN
	response.write("<br><br>In UpdatePaymentTransaction")
	response.write("<br>sPayType="&sPayType)
	response.write("<br>sOrderNo="&sOrderNo)
	response.write("<br>sPaymentResult="&sPaymentResult)
	response.write("<br>resp_message="&resp_message)
END IF

'sPaymentResult = Request("sPaymentResult")


IF sPayType="PayPal" THEN

	' ---- Read RegPaymentTableName for the updated record ----
	SET rsRegPay=Server.CreateObject("ADODB.recordset")
	sSQL = "SELECT * FROM "&RegPaymentTableName
	sSQL = sSQL + " WHERE Left(TourID,6) = '"&left(sTourID,6)&"' AND MemberID = '"&sMemberID&"' AND OrderNo='"&sOrderNo&"'"
	rsRegPay.open sSQL, SConnectionToTRATable, 3, 3


	' --- Verify that this record is only in the table once ---
	IF NOT rsRegPay.eof THEN
		DateNow = Date
	
		' --- Look into setting messages on paypal failure --- 	

		OpenCon
		sSQL = "UPDATE "&RegPaymentTableName
		sSQL = sSQL + " SET MemberID='"&sMemberID&"', TourID='"&sTourID&"', FirstName='"&sFirstName&"', LastName='"&sLastName&"'"
		sSQL = sSQL + ", City='"&sMembCity&"', State='"&sMembState&"', Amount='"&sPayAmount&"', OrderNo='"&sOrderNo&"'"
		sSQL = sSQL + ", Result='"&sPaymentResult&"', Message='"&resp_message&"', TransDate='"&DateNow&"', PayType='"&sPayType&"'"
		sSQL = sSQL + " WHERE MemberID='"&sMemberID&"' AND TourID='"&sTourID&"' AND OrderNo='"&sOrderNo&"'"
	END IF


	con.execute(sSQL)

ELSEIF sPayType="Check" OR sPayType="Cash" OR sPayType="Refund" THEN  ' --- Check Cash ---

	sPayAmount=Request("sPayAmount")

	DateNow = Date

	' --- Makes the value a negative number no matter whether Admin person entered a negative or positive ---
	IF sPayType="Refund" AND cdbl(sPayAmount)>cdbl(0) THEN sPayAmount = -sPayAmount
	

	OpenCon
	sSQL = "INSERT INTO "&RegPaymentTableName
	sSQL = sSQL + " (MemberID, TourID, CheckNo, OrderNO, TransDate, PayType, Amount, Result)"
	sSQL = sSQL + " VALUES ('"&sMemberID&"', '"&sTourID&"', '"&sCheckNo&"', '"&sOrderNo&"', '"&DateNow&"', '"&sPayType&"', '"&sPayAmount&"', '0')"

	con.execute(sSQL)
END IF

END SUB











' ========================================================================================================
' ---		BOTTOM OF PROGRAM CODE 		---  
' ========================================================================================================


' --- 2010 Functions
' --- dbo.fn_Search2010xTournAppID (@TAID char(6))  
' --- dbo.fn_SwiftFields2010xTournAppID (@TAID char(6))  

' --- 2009 Functions
' --- dbo.fn_OLRxTournAppID_NoLink (@TAID char(6))  
' --- dbo.fn_OLRxTournAppIDLInked (@TAID char(6))  
' --- dbo.fn_SearchxTournAppID_Linked (@TAID char(6)) 
' --- dbo.fn_SearchxTournAppID_NoLink (@TAID char(6))  

' --- 2008 Functions
' --- dbo.fn_OLRegFieldsXTournAppID (@TAID char(6))  






' **********************************************************************
' **************           2010 functions		****************
' **********************************************************************


' --------------------------------------------
' --- Returned in 2010 from GetSearchFunction
' --------------------------------------------

'CREATE FUNCTION dbo.fn_Search2010xTournAppID (@TAID char(6))  
'RETURNS TABLE
'AS  

'	RETURN

'	SELECT        P.ID, P.USAWSID, P.TSanction, P.TournAppID, P.SptsGrpID, P.EditCode, P.TDateS, P.TDateE, 
'                      P.TLateDate, P.TLateFee, P.TLFPerDay, P.TName, P.TSite, P.TSiteID, P.TCity, 
'                      P.TState, P.TSponsor, P.TSponsorID, P.TYear, P.TSkiYr, P.TDescription, 
'                      P.TSTATUS, P.TEventSlalom, P.TEventJump, P.TEventTrick, P.TEventFun, 
'                      P.FDescription, P.TOpenClosed, 
'                      P.TPandC, P.TTowBoatClosed, P.TDvOffered, P.TPandCPulls, P.TDirName, 
'                      P.TDirAddress, P.TDirCity, P.TDirState, P.TDirZip, P.TDirEmail, P.TDirFAX, 
'                      P.TDirPhoneAm, P.TDirPhonePm, P.TEntryFees, P.TEntryLimit, P.Canceled0, 
'                      P.TRegistrarName, P.TRegistrarPhone, P.TRegistrarEmail, P.TRegistrarAddr, 
'                      P.TRegistrarCity, P.TRegistrarState, P.TRegistrarZip, P.TRegistrarFax, P.TSigner, 
'                      P.TsEmail, P.TRegion, P.T5Star, P.TEventClin, P.JDClin, P.ADClin, 
'                      P.CDescription, P.GrassRoots, P.ClinNumParticipants, P.TEventWake, P.TRoundsWakeBd, 
'                      P.BoatWake, P.CableWake, P.TEventWSkate, P.TRoundsWSkate, P.BoatSkate, 
'                      P.CableSkate, P.TEventWSurf, P.TRoundsWSurf, P.WDescription,
'                      P.KEventFlip, P.KEventFree, P.KDescription, P.UseOLReg, P.AdminCode,
                      
'	         P.Gr1AWSPulls, P.Gr1ABCPulls, P.Gr1USWPulls, P.Gr1AKAPulls, P.Gr1USHPulls, P.Gr1WSDPulls, 
'                     P.Gr2AWS_SPulls, P.Gr2AWS_TPulls, P.Gr2ABC_SPulls, P.Gr2ABC_TPulls, P.Gr2USW_WPulls, 
'	         P.Gr2USW_SkatePulls, P.Gr2USW_SurfPulls, P.Gr2USW_RailJamPulls, P.Gr2AKA_SPulls, 
'                      P.Gr2AKA_TPulls, P.Gr2AKA_FreePulls, P.Gr2AKA_FlipPulls, P.Gr2USH_FreeRidePulls, P.Gr2USH_JumpOutPulls, 
'	         P.Gr2USH_FlipOutPulls, P.Gr2USH_BigAirPulls, P.Gr2USH_3TrickPulls, 
'                      P.TOCSafetyID, P.TOCDriverID, P.GrBoat, P.GrCable, P.AcceptTerms, P.GrFunDay, 
'	         P.GrTournament, P.GrFee_1, P.GrFee_2, P.GrFee_3, P.GrFee_4, P.GrFee_5,
                      

'	        R.EntryFee1, R.EntryFee2, R.EntryFee3, R.EntryFee4, R.EntryFee5, R.EntryFeeFamily, 
'                      R.FamEntType, R.LateFee, R.OtherFee, R.EmailAddress, R.ReceiveEmail, 
'                      R.Disc1, R.Disc2, R.Disc3, R.Disc4, R.ClubCode, R.DiscMeth, 
'                      R.GrEntryFee1, R.GrEntryFee2, R.GrEntryFee3, R.GrEntryFee4, R.AllowEdit,
'                      R.GrEntryFee5, R.ClinFeeJD, R.ClinFeeAD, R.MaxPulls, R.ReservedPulls, R.ReservedPullsCode, R.CMulti, 
'                      R.EMulti, R.LMulti, R.RMulti, R.MixedOptions, R.CSurchg, R.ESurchg, R.LSurchg, R.RSurchg, R.GREntryFeeIncluded, 
'                      R.GRDiscount, R.ClinicDiscount, R.Bio_Reqd, R.TrkLst_Reqd,  R.Form1Name,  R.Form2Name, R.Form3Name, 
'                      R.Form4Name, R.Form5Name, R.Form6Name, R.QualLevel1, R.RegisterDate,
'                      R.AthContr_Reqd, R.USAWSWaiver_Reqd, R.QualLevel, R.RestrictDv, R.BTickCost, R.BTickWithE, R.BTickFamily, 
'                      R.PayPalAct, R.PayPalOK, R.AWSEFDon_OK, R.MaxSlPulls, R.MaxTrPulls, R.MaxJpPulls,
'                      R.OF1Desc,R.OF2Desc,R.OF3Desc,R.OF4Desc,R.OF5Desc,R.OF6Desc,
'                      R.OF1Amt,R.OF2Amt,R.OF3Amt,R.OF4Amt,R.OF5Amt,R.OF6Amt,
'                      R.OF1MaxQty,R.OF2MaxQty,R.OF3MaxQty,R.OF4MaxQty,R.OF5MaxQty,R.OF6MaxQty,
'                      R.OF1Required, R.OF2Required, R.OF3Required, R.OF4Required, R.OF5Required, R.OF6Required,
'	         R.GRClassText, R.BaseClassText, R.UpgradeClassText, R.TEntryFeeFamExtra, R.MaxFamMember, R.OLRDisplayStatus,

'	         R.SClassC, R.SClassE, R.SClassL, R.SClassR, R.SClassCASH, R.SClassX, 
'	         R.JClassN, R.JClassC, R.JClassE, R.JClassL, R.JClassR, R.JClassCASH, R.JClassX, 
'                      R.TClassC, R.TClassE, R.TClassL, R.TClassR, R.TClassCASH, R.TClassX,  

'	         R.CJudgePID,	R.CDriverPID,  R.CScorePID, R.CSafPID,  R.TechCPID,
'	         R.Ap1JPID, R.Ap2JPID, R.Ap3JPID, R.Ap4JPID, R.Ap5JPID, 
'	         R.CJudge,	R.CDriver,  R.CScorer, R.CSafety,  R.TechCont,
'	         R.Ap1Judge, R.Ap2Judge, R.Ap3Judge, R.Ap4Judge, R.Ap5Judge, 
'                       R.PanAmJudge, R.PanAmPID, R.Announcer, R.AnnouncerPID, R.UseSkillLevels,

'                      R.KFlClassQ, R.KFlClassT,  R.KFrClassQ, R.KFrClassT, R.KSClassQ, R.KSClassT, R.KTClassT, R.KTClassQ,

'		R.USClassC, R.UTClassC, R.UJClassC,
		
'		R.BSClassC, R.BSClassL, R.BSClassR,
'		R.BTClassC, R.BTClassL, R.BTClassR,
'		R.BJClassC, R.BJClassL, R.BJClassR,

'		R.WWakeW, R.WSkateW, R.WSurfW, 

'	         G.*
'	FROM         dbo.Registration AS R INNER JOIN  dbo.Tschedul as P ON R.TournAppID = P.TournAppID
'						INNER JOIN dbo.GuideBk as G ON P.TournAppID = G.GTournAppID
'	WHERE     P.TournAppID = @TAID AND P.TYear > '2009'



' ------------------------------------------
' --- Returned in 2010 from GetOLRFunction
' ------------------------------------------

'CREATE FUNCTION dbo.fn_SwiftFields2010xTournAppID (@TAID char(6))  
'RETURNS TABLE
'AS  

'	RETURN

'	SELECT        P.ID, P.USAWSID, P.TSanction, P.TournAppID, P.SptsGrpID, P.EditCode, P.TDateS, P.TDateE, 
'                      P.TLateDate, P.TLateFee, P.TLFPerDay, P.TName, P.TSite, P.TSiteID, P.TCity, 
'                      P.TState, P.TSponsor, P.TSponsorID, P.TYear, P.TSkiYr, P.TDescription, 
'                      P.TSTATUS, P.TEventSlalom, P.TEventJump, P.TEventTrick, P.TEventFun, 
'                      P.FDescription, P.TOpenClosed, 
'                      P.TPandC, P.TTowBoatClosed, P.TDvOffered, P.TPandCPulls, P.TDirName, 
'                      P.TDirAddress, P.TDirCity, P.TDirState, P.TDirZip, P.TDirEmail, P.TDirFAX, 
'                      P.TDirPhoneAm, P.TDirPhonePm, P.TEntryFees, P.TEntryLimit, P.Canceled0, 
'                      P.TRegistrarName, P.TRegistrarPhone, P.TRegistrarEmail, P.TRegistrarAddr, 
'                      P.TRegistrarCity, P.TRegistrarState, P.TRegistrarZip, P.TRegistrarFax, P.TSigner, 
'                      P.TsEmail, P.TRegion, P.T5Star, P.TEventClin, P.JDClin, P.ADClin, 
'                      P.CDescription, P.GrassRoots, P.ClinNumParticipants, P.TEventWake, P.TRoundsWakeBd, 
'                      P.BoatWake, P.CableWake, P.TEventWSkate, P.TRoundsWSkate, P.BoatSkate, 
'                      P.CableSkate, P.TEventWSurf, P.TRoundsWSurf, P.WDescription,
'                      P.KEventFlip, P.KEventFree, P.KDescription, P.UseOLReg, P.AdminCode,
                      
'	         P.Gr1AWSPulls, P.Gr1ABCPulls, P.Gr1USWPulls, P.Gr1AKAPulls, P.Gr1USHPulls, P.Gr1WSDPulls, 
'                      P.Gr2AWS_SPulls, P.Gr2AWS_TPulls, P.Gr2ABC_SPulls, P.Gr2ABC_TPulls, P.Gr2USW_WPulls, 
'	         P.Gr2USW_SkatePulls, P.Gr2USW_SurfPulls, P.Gr2USW_RailJamPulls, P.Gr2AKA_SPulls, 
'                      P.Gr2AKA_TPulls, P.Gr2AKA_FreePulls, P.Gr2AKA_FlipPulls, P.Gr2USH_FreeRidePulls, P.Gr2USH_JumpOutPulls, 
'	         P.Gr2USH_FlipOutPulls, P.Gr2USH_BigAirPulls, P.Gr2USH_3TrickPulls, 
'                      P.TOCSafetyID, P.TOCDriverID, P.GrBoat, P.GrCable, P.AcceptTerms, P.GrFunDay, 
'	         P.GrTournament, P.GrFee_1, P.GrFee_2, P.GrFee_3, P.GrFee_4, P.GrFee_5,
                      

'	        R.EntryFee1, R.EntryFee2, R.EntryFee3, R.EntryFee4, R.EntryFee5, R.EntryFeeFamily, 
'                      R.FamEntType, R.LateFee, R.OtherFee, R.EmailAddress, R.ReceiveEmail, 
'                      R.Disc1, R.Disc2, R.Disc3, R.Disc4, R.ClubCode, R.DiscMeth, 
'                      R.GrEntryFee1, R.GrEntryFee2, R.GrEntryFee3, R.GrEntryFee4, R.AllowEdit,
'                      R.GrEntryFee5, R.ClinFeeJD, R.ClinFeeAD, R.MaxPulls, R.ReservedPulls, R.ReservedPullsCode, R.CMulti, 
'                      R.EMulti, R.LMulti, R.RMulti, R.MixedOptions, R.CSurchg, R.ESurchg, R.LSurchg, R.RSurchg, R.GREntryFeeIncluded, 
'                      R.GRDiscount, R.ClinicDiscount, R.Bio_Reqd, R.TrkLst_Reqd,  R.Form1Name,  R.Form2Name, R.Form3Name, 
'                      R.Form4Name, R.Form5Name, R.Form6Name, R.QualLevel1, R.RegisterDate,
'                      R.AthContr_Reqd, R.USAWSWaiver_Reqd, R.QualLevel, R.RestrictDv, R.BTickCost, R.BTickWithE, R.BTickFamily, 
'                      R.PayPalAct, R.PayPalOK, R.AWSEFDon_OK, R.MaxSlPulls, R.MaxTrPulls, R.MaxJpPulls,
'                      R.OF1Desc,R.OF2Desc,R.OF3Desc,R.OF4Desc,R.OF5Desc,R.OF6Desc,
'                      R.OF1Amt,R.OF2Amt,R.OF3Amt,R.OF4Amt,R.OF5Amt,R.OF6Amt,
'                      R.OF1MaxQty,R.OF2MaxQty,R.OF3MaxQty,R.OF4MaxQty,R.OF5MaxQty,R.OF6MaxQty,
'                      R.OF1Required, R.OF2Required, R.OF3Required, R.OF4Required, R.OF5Required, R.OF6Required,
'	         R.GRClassText, R.BaseClassText, R.UpgradeClassText, R.TEntryFeeFamExtra, R.MaxFamMember, R.OLRDisplayStatus,

'	         R.SClassC, R.SClassE, R.SClassL, R.SClassR, R.SClassCASH, R.SClassX, 
'	         R.JClassN, R.JClassC, R.JClassE, R.JClassL, R.JClassR, R.JClassCASH, R.JClassX, 
'                      R.TClassC, R.TClassE, R.TClassL, R.TClassR, R.TClassCASH, R.TClassX,  

'	         R.CJudgePID,	R.CDriverPID,  R.CScorePID, R.CSafPID,  R.TechCPID,
'	         R.Ap1JPID, R.Ap2JPID, R.Ap3JPID, R.Ap4JPID, R.Ap5JPID, 
'	         R.CJudge,	R.CDriver,  R.CScorer, R.CSafety,  R.TechCont,
'	         R.Ap1Judge, R.Ap2Judge, R.Ap3Judge, R.Ap4Judge, R.Ap5Judge, 
'                       R.PanAmJudge, R.PanAmPID, R.Announcer, R.AnnouncerPID, R.UseSkillLevels,

'                      R.KFlClassQ, R.KFlClassT,  R.KFrClassQ, R.KFrClassT, R.KSClassQ, R.KSClassT, R.KTClassT, R.KTClassQ, 

'		R.USClassC, R.UTClassC, R.UJClassC,
		
'		R.BSClassC, R.BSClassL, R.BSClassR,
'		R.BTClassC, R.BTClassL, R.BTClassR,
'		R.BJClassC, R.BJClassL, R.BJClassR,

'		R.WWakeW, R.WSkateW, R.WSurfW,

'	         G.*
'	FROM         dbo.Registration AS R INNER JOIN  dbo.Tschedul as P ON R.TournAppID = P.TournAppID
'						INNER JOIN dbo.GuideBk as G ON P.TournAppID = G.GTournAppID
'	WHERE     P.TournAppID = @TAID AND P.TSanType <> 2 AND P.TSTATUS > 1






' **********************************************************************
' **************           2009 function 		****************
' **********************************************************************


' ---------------------------------------------------------------------------
' --- Returned in 2009 from GetSearchFunction for a tournament WITH Link
' ---------------------------------------------------------------------------

'CREATE FUNCTION dbo.fn_SearchxTournAppID_Linked (@TAID char(6)) 
'RETURNS TABLE
'AS  

' RETURN

' SELECT        P.ID, P.USAWSID, P.TSanction, P.TournAppID, P.SptsGrpID, P.EditCode, P.TDateS, P.TDateE, 
'                      P.TLateDate, P.TLateFee, P.TLFPerDay, P.TName, P.TSite, P.TSiteID, P.TCity, 
'                      P.TState, P.TSponsor, P.TSponsorID, P.TYear, P.TSkiYr, P.TDescription, 
'                      P.TSTATUS, P.TEventSlalom, P.TEventJump, P.TEventTrick, P.TEventFun, 
'                      P.TEventNSL, P.TEventNBL, P.TEventNWL, P.TOpenClosed, 
'                      P.TPandC, P.TTowBoatClosed, P.TDvOffered, P.THSClassF, P.THSClassN, P.THSClassI, 
'                      P.THSClassC, P.THSClassE, P.THSClassL, P.THSClassR, P.THSClassCASH, 
'                      P.THSClassX, P.THJClassF, P.THJClassN, P.THJClassI, P.THJClassC, 
'                      P.THJClassE, P.THJClassL, P.THJClassR, P.THJClassCASH, P.THJClassX, 
'                      P.THTClassF, P.THTClassN, P.THTClassI, P.THTClassC, P.THTClassE, 
'                      P.THTClassL, P.THTClassR, P.THTClassCASH, P.THTClassX, P.TRoundsS, 
'                      P.TRoundsT, P.TRoundsJ, P.TRoundsF, P.TPandCPulls, P.TDirName, 
'                      P.TDirAddress, P.TDirCity, P.TDirState, P.TDirZip, P.TDirEmail, P.TDirFAX, 
'                      P.TDirPhoneAm, P.TDirPhonePm, P.TEntryFees, P.TEntryLimit, P.Canceled0, 
'                      P.TRegistrarName, P.TRegistrarPhone, P.TRegistrarEmail, P.TRegistrarAddr, 
'                      P.TRegistrarCity, P.TRegistrarState, P.TRegistrarZip, P.TRegistrarFax, P.TSigner, 
'                      P.TsEmail, P.TRegion, P.T5Star, P.TEventClin, P.JDClin, P.ADClin, 
'                      P.CDescription, P.ClinNumParticipants, P.TEventWake, P.TRoundsWakeBd, 
'                      P.BoatWake, P.CableWake, P.TEventWSkate, P.TRoundsWSkate, P.BoatSkate, 
'                      P.CableSkate, P.TEventWSurf, P.TRoundsWSurf, P.WDescription, P.TEventFHF, 
'                      P.TEventFKB, P.TEventFDA, P.TEventF3ev, P.TEventFB, P.TEventFW, 
'                      P.KEventFlip, P.KRoundsFlip, P.KFlipClassQ, P.KFlipClassT, P.KEventFree, 
'                      P.KRoundsFree, P.KFreeClassQ, P.KFreeClassT, P.KDescription, P.KSlalomClassQ, 
'                      P.KSlalomClassT, P.KTrickClassT, P.KTrickClassQ, P.UseOLReg, P.AdminCode,
'                      P.TO1ApDrive, P.TO1ApJudge, P.TO1ApScore, P.TO2ApDrive, P.TO2ApJudge, P.TO3ApJudge, 
'                      P.TO4ApJudge, P.TO5ApJudge, P.TOCDriver, P.TOCJudge, P.TOCSafety, 
'                      P.TOCScore, P.TOOoAJudge, P.TOPanAmJudge, P.TOTechCont, P.TOAnnounce,
          
 '                      Gr.FDescription,  P.GrassRoots, Gr.Gr1AWSPulls, Gr.Gr1ABCPulls, Gr.Gr1USWPulls, Gr.Gr1AKAPulls, Gr.Gr1USHPulls, Gr.Gr1WSDPulls, 
'                      Gr.Gr2AWS_SPulls, Gr.Gr2AWS_TPulls, Gr.Gr2ABC_SPulls, Gr.Gr2ABC_TPulls, Gr.Gr2USW_WPulls, 
'          Gr.Gr2USW_SkatePulls, Gr.Gr2USW_SurfPulls, Gr.Gr2USW_RailJamPulls, Gr.Gr2AKA_SPulls, 
'                      Gr.Gr2AKA_TPulls, Gr.Gr2AKA_FreePulls, Gr.Gr2AKA_FlipPulls, Gr.Gr2USH_FreeRidePulls, Gr.Gr2USH_JumpOutPulls, 
'          Gr.Gr2USH_FlipOutPulls, Gr.Gr2USH_BigAirPulls, Gr.Gr2USH_3TrickPulls, 
'                      Gr.TOCSafetyID, Gr.TOCDriverID, Gr.GrBoat, Gr.GrCable, Gr.AcceptTerms, Gr.GrFunDay, 
'          Gr.GrTournament, Gr.GrFee_1, Gr.GrFee_2, Gr.GrFee_3, Gr.GrFee_4, Gr.GrFee_5,
      
'          G.*
 
' FROM    Tschedul as P INNER JOIN Tschedul as Gr ON P.USAWSID = Gr.USAWSID 
'    INNER JOIN GuideBk as G on P.TournAppID = G.GTournAppID 
'  WHERE Gr.USAWSID <> Gr.ID and P.TournAppID =  @TAID 




' ---------------------------------------------------------------------------
' --- Returned in 2009 from GetSearchFunction for a tournament with NO LINK
' ---------------------------------------------------------------------------

'CREATE FUNCTION dbo.fn_SearchxTournAppID_NoLink (@TAID char(6))  
'RETURNS TABLE
'AS  

' RETURN

' SELECT        P.ID, P.USAWSID, P.TSanction, P.TournAppID, P.SptsGrpID, P.EditCode, P.TDateS, P.TDateE, 
'                      P.TLateDate, P.TLateFee, P.TLFPerDay, P.TName, P.TSite, P.TSiteID, P.TCity, 
'                      P.TState, P.TSponsor, P.TSponsorID, P.TYear, P.TSkiYr, P.TDescription, 
'                      P.TSTATUS, P.TEventSlalom, P.TEventJump, P.TEventTrick, P.TEventFun, 
'                      P.TEventNSL, P.TEventNBL, P.TEventNWL, P.FDescription, P.TOpenClosed, 
'                      P.TPandC, P.TTowBoatClosed, P.TDvOffered, P.THSClassF, P.THSClassN, P.THSClassI, 
'                      P.THSClassC, P.THSClassE, P.THSClassL, P.THSClassR, P.THSClassCASH, 
'                      P.THSClassX, P.THJClassF, P.THJClassN, P.THJClassI, P.THJClassC, 
'                      P.THJClassE, P.THJClassL, P.THJClassR, P.THJClassCASH, P.THJClassX, 
'                      P.THTClassF, P.THTClassN, P.THTClassI, P.THTClassC, P.THTClassE, 
'                      P.THTClassL, P.THTClassR, P.THTClassCASH, P.THTClassX, P.TRoundsS, 
'                      P.TRoundsT, P.TRoundsJ, P.TRoundsF, P.TPandCPulls, P.TDirName, 
'                      P.TDirAddress, P.TDirCity, P.TDirState, P.TDirZip, P.TDirEmail, P.TDirFAX, 
'                      P.TDirPhoneAm, P.TDirPhonePm, P.TEntryFees, P.TEntryLimit, P.Canceled0, 
'                      P.TRegistrarName, P.TRegistrarPhone, P.TRegistrarEmail, P.TRegistrarAddr, 
'                      P.TRegistrarCity, P.TRegistrarState, P.TRegistrarZip, P.TRegistrarFax, P.TSigner, 
'                      P.TsEmail, P.TRegion, P.T5Star, P.TEventClin, P.JDClin, P.ADClin, 
'                      P.CDescription, P.GrassRoots, P.ClinNumParticipants, P.TEventWake, P.TRoundsWakeBd, 
'                      P.BoatWake, P.CableWake, P.TEventWSkate, P.TRoundsWSkate, P.BoatSkate, 
'                      P.CableSkate, P.TEventWSurf, P.TRoundsWSurf, P.WDescription, P.TEventFHF, 
'                      P.TEventFKB, P.TEventFDA, P.TEventF3ev, P.TEventFB, P.TEventFW, 
'                      P.KEventFlip, P.KRoundsFlip, P.KFlipClassQ, P.KFlipClassT, P.KEventFree, 
'                      P.KRoundsFree, P.KFreeClassQ, P.KFreeClassT, P.KDescription, P.KSlalomClassQ, 
'                      P.KSlalomClassT, P.KTrickClassT, P.KTrickClassQ, P.UseOLReg, P.AdminCode,
'                      P.TO1ApDrive, P.TO1ApJudge, P.TO1ApScore, P.TO2ApDrive, P.TO2ApJudge, P.TO3ApJudge, 
'                      P.TO4ApJudge, P.TO5ApJudge, P.TOCDriver, P.TOCJudge, P.TOCSafety, 
'                      P.TOCScore, P.TOOoAJudge, P.TOPanAmJudge, P.TOTechCont, P.TOAnnounce,'

'                      P.Gr1AWSPulls, P.Gr1ABCPulls, P.Gr1USWPulls, P.Gr1AKAPulls, P.Gr1USHPulls, P.Gr1WSDPulls, 
'                      P.Gr2AWS_SPulls, P.Gr2AWS_TPulls, P.Gr2ABC_SPulls, P.Gr2ABC_TPulls, P.Gr2USW_WPulls, 
'          P.Gr2USW_SkatePulls, P.Gr2USW_SurfPulls, P.Gr2USW_RailJamPulls, P.Gr2AKA_SPulls, 
'                      P.Gr2AKA_TPulls, P.Gr2AKA_FreePulls, P.Gr2AKA_FlipPulls, P.Gr2USH_FreeRidePulls, P.Gr2USH_JumpOutPulls, 
'          P.Gr2USH_FlipOutPulls, P.Gr2USH_BigAirPulls, P.Gr2USH_3TrickPulls, 
'                      P.TOCSafetyID, P.TOCDriverID, P.GrBoat, P.GrCable, P.AcceptTerms, P.GrFunDay, 
'          P.GrTournament, P.GrFee_1, P.GrFee_2, P.GrFee_3, P.GrFee_4, P.GrFee_5,
'                      
'          G.*
' FROM         dbo.Tschedul as P INNER JOIN dbo.GuideBk as G ON P.TournAppID = G.GTournAppID
' WHERE     P.TournAppID = @TAID 





' ---------------------------------------------------------------------------
' --- Returned in 2009 from GetOLRFunction for a tournament with NO LINK
' ---------------------------------------------------------------------------

'CREATE FUNCTION dbo.fn_OLRxTournAppID_NoLink (@TAID char(6))  
'RETURNS TABLE
'AS  

' RETURN

' SELECT        P.ID, P.USAWSID, P.TSanction, P.TournAppID, P.SptsGrpID, P.EditCode, P.TDateS, P.TDateE, 
'                      P.TLateDate, P.TLateFee, P.TLFPerDay, P.TName, P.TSite, P.TSiteID, P.TCity, 
'                      P.TState, P.TSponsor, P.TSponsorID, P.TYear, P.TSkiYr, P.TDescription, 
'                      P.TSTATUS, P.TEventSlalom, P.TEventJump, P.TEventTrick, P.TEventFun, 
'                      P.TEventNSL, P.TEventNBL, P.TEventNWL, P.FDescription, P.TOpenClosed, 
'                      P.TPandC, P.TTowBoatClosed, P.TDvOffered, P.THSClassF, P.THSClassN, P.THSClassI, 
'                      P.THSClassC, P.THSClassE, P.THSClassL, P.THSClassR, P.THSClassCASH, 
'                      P.THSClassX, P.THJClassF, P.THJClassN, P.THJClassI, P.THJClassC, 
'                      P.THJClassE, P.THJClassL, P.THJClassR, P.THJClassCASH, P.THJClassX, 
'                      P.THTClassF, P.THTClassN, P.THTClassI, P.THTClassC, P.THTClassE, 
'                      P.THTClassL, P.THTClassR, P.THTClassCASH, P.THTClassX, P.TRoundsS, 
'                      P.TRoundsT, P.TRoundsJ, P.TRoundsF, P.TPandCPulls, P.TDirName, 
'                      P.TDirAddress, P.TDirCity, P.TDirState, P.TDirZip, P.TDirEmail, P.TDirFAX, 
'                      P.TDirPhoneAm, P.TDirPhonePm, P.TEntryFees, P.TEntryLimit, P.Canceled0, 
'                      P.TRegistrarName, P.TRegistrarPhone, P.TRegistrarEmail, P.TRegistrarAddr, 
'                      P.TRegistrarCity, P.TRegistrarState, P.TRegistrarZip, P.TRegistrarFax, P.TSigner, 
'                      P.TsEmail, P.TRegion, P.T5Star, P.TEventClin, P.JDClin, P.ADClin, 
'                      P.CDescription, P.GrassRoots, P.ClinNumParticipants, P.TEventWake, P.TRoundsWakeBd, 
'                      P.BoatWake, P.CableWake, P.TEventWSkate, P.TRoundsWSkate, P.BoatSkate, 
'                      P.CableSkate, P.TEventWSurf, P.TRoundsWSurf, P.WDescription, P.TEventFHF, 
'                      P.TEventFKB, P.TEventFDA, P.TEventF3ev, P.TEventFB, P.TEventFW, 
'                      P.KEventFlip, P.KRoundsFlip, P.KFlipClassQ, P.KFlipClassT, P.KEventFree, 
'                      P.KRoundsFree, P.KFreeClassQ, P.KFreeClassT, P.KDescription, P.KSlalomClassQ, 
'                      P.KSlalomClassT, P.KTrickClassT, P.KTrickClassQ, P.UseOLReg, P.AdminCode,
'                      P.TO1ApDrive, P.TO1ApJudge, P.TO1ApScore, P.TO2ApDrive, P.TO2ApJudge, P.TO3ApJudge, 
'                      P.TO4ApJudge, P.TO5ApJudge, P.TOCDriver, P.TOCJudge, P.TOCSafety, 
'                      P.TOCScore, P.TOOoAJudge, P.TOPanAmJudge, P.TOTechCont, P.TOAnnounce,

'                      P.Gr1AWSPulls, P.Gr1ABCPulls, P.Gr1USWPulls, P.Gr1AKAPulls, P.Gr1USHPulls, P.Gr1WSDPulls, 
'                      P.Gr2AWS_SPulls, P.Gr2AWS_TPulls, P.Gr2ABC_SPulls, P.Gr2ABC_TPulls, P.Gr2USW_WPulls, 
'          P.Gr2USW_SkatePulls, P.Gr2USW_SurfPulls, P.Gr2USW_RailJamPulls, P.Gr2AKA_SPulls, 
'                      P.Gr2AKA_TPulls, P.Gr2AKA_FreePulls, P.Gr2AKA_FlipPulls, P.Gr2USH_FreeRidePulls, P.Gr2USH_JumpOutPulls, 
'          P.Gr2USH_FlipOutPulls, P.Gr2USH_BigAirPulls, P.Gr2USH_3TrickPulls, 
'                      P.TOCSafetyID, P.TOCDriverID, P.GrBoat, P.GrCable, P.AcceptTerms, P.GrFunDay, 
'          P.GrTournament, P.GrFee_1, P.GrFee_2, P.GrFee_3, P.GrFee_4, P.GrFee_5,
                      

'         R.EntryFee1, R.EntryFee2, R.EntryFee3, R.EntryFee4, R.EntryFee5, R.EntryFeeFamily, 
'                      R.FamEntType, R.LateFee, R.OtherFee, R.EmailAddress, R.ReceiveEmail, 
'                      R.Disc1, R.Disc2, R.Disc3, R.Disc4, R.ClubCode, R.DiscMeth, 
'                      R.GrEntryFee1, R.GrEntryFee2, R.GrEntryFee3, R.GrEntryFee4, R.AllowEdit,
'                      R.GrEntryFee5, R.ClinFeeJD, R.ClinFeeAD, R.MaxPulls, R.ReservedPulls, R.ReservedPullsCode, R.CMulti, 
'                      R.EMulti, R.LMulti, R.RMulti, R.MixedOptions, R.CSurchg, R.ESurchg, R.LSurchg, R.RSurchg, R.GREntryFeeIncluded, 
'                      R.GRDiscount, R.ClinicDiscount, R.Bio_Reqd, R.TrkLst_Reqd,  R.Form1Name,  R.Form2Name, R.Form3Name, 
'                      R.Form4Name, R.Form5Name, R.Form6Name, R.QualLevel1, R.RegisterDate,
'                      R.AthContr_Reqd, R.USAWSWaiver_Reqd, R.QualLevel, R.RestrictDv, R.BTickCost, R.BTickWithE, R.BTickFamily, 
'                      R.PayPalAct, R.PayPalOK, R.AWSEFDon_OK, R.MaxSlPulls, R.MaxTrPulls, R.MaxJpPulls,
'                      R.OF1Desc,R.OF2Desc,R.OF3Desc,R.OF4Desc,R.OF5Desc,R.OF6Desc,
'                      R.OF1Amt,R.OF2Amt,R.OF3Amt,R.OF4Amt,R.OF5Amt,R.OF6Amt,
'                      R.OF1MaxQty,R.OF2MaxQty,R.OF3MaxQty,R.OF4MaxQty,R.OF5MaxQty,R.OF6MaxQty,
'                      R.OF1Required, R.OF2Required, R.OF3Required, R.OF4Required, R.OF5Required, R.OF6Required,
'          R.GRClassText, R.BaseClassText, R.UpgradeClassText, R.TEntryFeeFamExtra, R.MaxFamMember, R.OLRDisplayStatus,

'          G.*
' FROM         dbo.Registration AS R INNER JOIN  dbo.Tschedul as P ON R.TournAppID = P.TournAppID
'      INNER JOIN dbo.GuideBk as G ON P.TournAppID = G.GTournAppID
' WHERE     P.TournAppID = @TAID AND P.TSanType <> 2 AND P.TSTATUS > 1


 

' ----------------------------------------------------------------------
' --- Returned in 2009 from GetOLRFunction for a tournament WITH LINK
' ----------------------------------------------------------------------

'CREATE FUNCTION dbo.fn_OLRxTournAppIDLInked (@TAID char(6))  
'RETURNS TABLE
'AS  

' RETURN

' SELECT        P.ID, P.USAWSID, P.TSanction, P.TournAppID, P.SptsGrpID, P.EditCode, P.TDateS, P.TDateE, 
'                      P.TLateDate, P.TLateFee, P.TLFPerDay, P.TName, P.TSite, P.TSiteID, P.TCity, 
'                      P.TState, P.TSponsor, P.TSponsorID, P.TYear, P.TSkiYr, P.TDescription, 
'                      P.TSTATUS, P.TEventSlalom, P.TEventJump, P.TEventTrick, P.TEventFun, 
'                      P.TEventNSL, P.TEventNBL, P.TEventNWL, P.TOpenClosed, 
'                      P.TPandC, P.TTowBoatClosed, P.TDvOffered, P.THSClassF, P.THSClassN, P.THSClassI, 
'                      P.THSClassC, P.THSClassE, P.THSClassL, P.THSClassR, P.THSClassCASH, 
'                      P.THSClassX, P.THJClassF, P.THJClassN, P.THJClassI, P.THJClassC, 
'                      P.THJClassE, P.THJClassL, P.THJClassR, P.THJClassCASH, P.THJClassX, 
'                      P.THTClassF, P.THTClassN, P.THTClassI, P.THTClassC, P.THTClassE, 
'                      P.THTClassL, P.THTClassR, P.THTClassCASH, P.THTClassX, P.TRoundsS, 
'                      P.TRoundsT, P.TRoundsJ, P.TRoundsF, P.TPandCPulls, P.TDirName, 
'                      P.TDirAddress, P.TDirCity, P.TDirState, P.TDirZip, P.TDirEmail, P.TDirFAX, 
'                      P.TDirPhoneAm, P.TDirPhonePm, P.TEntryFees, P.TEntryLimit, P.Canceled0, 
'                      P.TRegistrarName, P.TRegistrarPhone, P.TRegistrarEmail, P.TRegistrarAddr, 
'                      P.TRegistrarCity, P.TRegistrarState, P.TRegistrarZip, P.TRegistrarFax, P.TSigner, 
'                      P.TsEmail, P.TRegion, P.T5Star, P.TEventClin, P.JDClin, P.ADClin, 
'                      P.CDescription, P.ClinNumParticipants, P.TEventWake, P.TRoundsWakeBd, 
'                      P.BoatWake, P.CableWake, P.TEventWSkate, P.TRoundsWSkate, P.BoatSkate, 
'                      P.CableSkate, P.TEventWSurf, P.TRoundsWSurf, P.WDescription, P.TEventFHF, 
'                      P.TEventFKB, P.TEventFDA, P.TEventF3ev, P.TEventFB, P.TEventFW, 
'                      P.KEventFlip, P.KRoundsFlip, P.KFlipClassQ, P.KFlipClassT, P.KEventFree, 
'                      P.KRoundsFree, P.KFreeClassQ, P.KFreeClassT, P.KDescription, P.KSlalomClassQ, 
'                      P.KSlalomClassT, P.KTrickClassT, P.KTrickClassQ, P.UseOLReg, P.AdminCode,
'                      P.TO1ApDrive, P.TO1ApJudge, P.TO1ApScore, P.TO2ApDrive, P.TO2ApJudge, P.TO3ApJudge, 
'                      P.TO4ApJudge, P.TO5ApJudge, P.TOCDriver, P.TOCJudge, P.TOCSafety, 
'                      P.TOCScore, P.TOOoAJudge, P.TOPanAmJudge, P.TOTechCont, P.TOAnnounce,
          
'                       Gr.FDescription,  P.GrassRoots, Gr.Gr1AWSPulls, Gr.Gr1ABCPulls, Gr.Gr1USWPulls, Gr.Gr1AKAPulls, Gr.Gr1USHPulls, Gr.Gr1WSDPulls, 
'                      Gr.Gr2AWS_SPulls, Gr.Gr2AWS_TPulls, Gr.Gr2ABC_SPulls, Gr.Gr2ABC_TPulls, Gr.Gr2USW_WPulls, 
'          Gr.Gr2USW_SkatePulls, Gr.Gr2USW_SurfPulls, Gr.Gr2USW_RailJamPulls, Gr.Gr2AKA_SPulls, 
'                      Gr.Gr2AKA_TPulls, Gr.Gr2AKA_FreePulls, Gr.Gr2AKA_FlipPulls, Gr.Gr2USH_FreeRidePulls, Gr.Gr2USH_JumpOutPulls, 
'          Gr.Gr2USH_FlipOutPulls, Gr.Gr2USH_BigAirPulls, Gr.Gr2USH_3TrickPulls, 
'                      Gr.TOCSafetyID, Gr.TOCDriverID, Gr.GrBoat, Gr.GrCable, Gr.AcceptTerms, Gr.GrFunDay, 
'          Gr.GrTournament, Gr.GrFee_1, Gr.GrFee_2, Gr.GrFee_3, Gr.GrFee_4, Gr.GrFee_5,
      

'         R.EntryFee1, R.EntryFee2, R.EntryFee3, R.EntryFee4, R.EntryFee5, R.EntryFeeFamily, 
'                      R.FamEntType, R.LateFee, R.OtherFee, R.EmailAddress, R.ReceiveEmail, 
'                      R.Disc1, R.Disc2, R.Disc3, R.Disc4, R.ClubCode, R.DiscMeth, 
'                      R.GrEntryFee1, R.GrEntryFee2, R.GrEntryFee3, R.GrEntryFee4, R.AllowEdit,
'                      R.GrEntryFee5, R.ClinFeeJD, R.ClinFeeAD, R.MaxPulls, R.ReservedPulls, R.ReservedPullsCode, R.CMulti, 
'                      R.EMulti, R.LMulti, R.RMulti, R.MixedOptions, R.CSurchg, R.ESurchg, R.LSurchg, R.RSurchg, R.GREntryFeeIncluded, 
'                      R.GRDiscount, R.ClinicDiscount, R.Bio_Reqd, R.TrkLst_Reqd,  R.Form1Name,  R.Form2Name, R.Form3Name, 
'                      R.Form4Name, R.Form5Name, R.Form6Name, R.QualLevel1, R.RegisterDate,
'                      R.AthContr_Reqd, R.USAWSWaiver_Reqd, R.QualLevel, R.RestrictDv, R.BTickCost, R.BTickWithE, R.BTickFamily, 
'                      R.PayPalAct, R.PayPalOK, R.AWSEFDon_OK, R.MaxSlPulls, R.MaxTrPulls, R.MaxJpPulls,
'                      R.OF1Desc,R.OF2Desc,R.OF3Desc,R.OF4Desc,R.OF5Desc,R.OF6Desc,
'                      R.OF1Amt,R.OF2Amt,R.OF3Amt,R.OF4Amt,R.OF5Amt,R.OF6Amt,
'                      R.OF1MaxQty,R.OF2MaxQty,R.OF3MaxQty,R.OF4MaxQty,R.OF5MaxQty,R.OF6MaxQty,
'                      R.OF1Required, R.OF2Required, R.OF3Required, R.OF4Required, R.OF5Required, R.OF6Required,
'          R.GRClassText, R.BaseClassText, R.UpgradeClassText, R.TEntryFeeFamExtra, R.MaxFamMember, R.OLRDisplayStatus,

'          G.*
 
' FROM    Tschedul as P INNER JOIN Tschedul as Gr ON P.USAWSID = Gr.USAWSID 
'    INNER JOIN Registration as R ON R.TournAppID = P.TournAppID
'         INNER JOIN GuideBk as G on P.TournAppID = G.GTournAppID 
'  WHERE Gr.USAWSID <> Gr.ID and P.TournAppID =  @TAID AND P.TSanType <> 2 AND P.TStatus > 1  

'******************************************************





' **********************************************************************
' **************           2008 function 		****************
' **********************************************************************



' -------------------------------------------
' --- Returned in 2008 from GetSearchFunction
' -------------------------------------------


'CREATE FUNCTION dbo.fn_OLRegFieldsXTournAppID (@TAID char(6))  
'RETURNS TABLE
'AS  
'RETURN
'SELECT      dbo.Tschedul.ID, dbo.Tschedul.USAWSID, dbo.Tschedul.TSanction, dbo.Tschedul.TournAppID, dbo.Tschedul.SptsGrpID, dbo.Tschedul.TSanType, dbo.Tschedul.EditCode, 
'                         dbo.Tschedul.TDateS, dbo.Tschedul.TDateE, dbo.Tschedul.TLateDate, dbo.Tschedul.TLateFee, dbo.Tschedul.TLFPerDay, dbo.Tschedul.TName, dbo.Tschedul.TSiteID, 
'                         dbo.Tschedul.TSite, dbo.Tschedul.TCity, dbo.Tschedul.TState, dbo.Tschedul.TSponsor, dbo.Tschedul.TSponsorID, dbo.Tschedul.PersonID, dbo.Tschedul.TSponsorIDOld, 
'                         dbo.Tschedul.TYear, dbo.Tschedul.TSkiYr, dbo.Tschedul.TDescription, dbo.Tschedul.MaxPulls, dbo.Tschedul.RestrictDv, dbo.Tschedul.Pending, dbo.Tschedul.chkHQOnly1, 
'                         dbo.Tschedul.chkHQOnly2, dbo.Tschedul.chkHQOnly3, dbo.Tschedul.chkHQOnly4, dbo.Tschedul.chkSanOK, dbo.Tschedul.TSanApproved, dbo.Tschedul.txtHQOnly0, 
'                         dbo.Tschedul.txtHQOnly1, dbo.Tschedul.txtHQOnly2, dbo.Tschedul.txtHQOnly4, dbo.Tschedul.HQOtherDetail, dbo.Tschedul.HQCreditDetail, 
'                         dbo.Tschedul.TKitOKRegnOfficials, dbo.Tschedul.TKitOKSafetyForm, dbo.Tschedul.TKitOKRegnFeePd, dbo.Tschedul.TKitOKAWSAFeePd, dbo.Tschedul.TKitOKGuidebookAd, 
'                         dbo.Tschedul.chkRegionOK, dbo.Tschedul.TSTATUS, dbo.Tschedul.TKitPaidAWSA, dbo.Tschedul.TKitPaidRegn, dbo.Tschedul.TKitCkNoAWSA, dbo.Tschedul.TKitCkNoRegn, 
'                         dbo.Tschedul.TEventSlalom, dbo.Tschedul.TEventJump, dbo.Tschedul.TEventTrick, dbo.Tschedul.TEventFun, dbo.Tschedul.TEventNSL, dbo.Tschedul.TEventNBL, 
'                         dbo.Tschedul.TEventNWL, dbo.Tschedul.TEventCustom, dbo.Tschedul.FDescription, dbo.Tschedul.TOpenClosed, dbo.Tschedul.TPandC, dbo.Tschedul.TTowBoatClosed, 
'                         dbo.Tschedul.TDvOffered, dbo.Tschedul.THSClassF, dbo.Tschedul.THSClassN, dbo.Tschedul.THSClassI, dbo.Tschedul.THSClassC, dbo.Tschedul.THSClassE, 
'                         dbo.Tschedul.THSClassL, dbo.Tschedul.THSClassR, dbo.Tschedul.THSClassCASH, dbo.Tschedul.THSClassX, dbo.Tschedul.THJClassF, dbo.Tschedul.THJClassN, 
'                         dbo.Tschedul.THJClassI, dbo.Tschedul.THJClassC, dbo.Tschedul.THJClassE, dbo.Tschedul.THJClassL, dbo.Tschedul.THJClassR, dbo.Tschedul.THJClassCASH, 
'                         dbo.Tschedul.THJClassX, dbo.Tschedul.THTClassF, dbo.Tschedul.THTClassN, dbo.Tschedul.THTClassI, dbo.Tschedul.THTClassC, dbo.Tschedul.THTClassE, 
'                         dbo.Tschedul.THTClassL, dbo.Tschedul.THTClassR, dbo.Tschedul.THTClassCASH, dbo.Tschedul.THTClassX, dbo.Tschedul.TRoundsS, dbo.Tschedul.TRoundsT, 
'                         dbo.Tschedul.TRoundsJ, dbo.Tschedul.TRoundsF, dbo.Tschedul.TPandCPulls, dbo.Tschedul.TSlSkiersPerRnd, dbo.Tschedul.TTrSkiersPerRnd, 
'                         dbo.Tschedul.TJpSkiersPerRnd, dbo.Tschedul.TFunPerRnd, dbo.Tschedul.TDirName, dbo.Tschedul.TDirAddress, dbo.Tschedul.TDirCity, dbo.Tschedul.TDirState, 
'                         dbo.Tschedul.TDirZip, dbo.Tschedul.TDirEmail, dbo.Tschedul.TDirFAX, dbo.Tschedul.TDirPhoneAm, dbo.Tschedul.TDirPhonePm, dbo.Tschedul.TEntryFees, 
'                         dbo.Tschedul.TEntryLimit, dbo.Tschedul.TKitExtraRides, dbo.Tschedul.TKitLateFeeAWSA, dbo.Tschedul.Canceled0, dbo.Tschedul.Canceled1, dbo.Tschedul.Canceled2, 
'                         dbo.Tschedul.Canceled3, dbo.Tschedul.TKitRefund0, dbo.Tschedul.TKitRefund1, dbo.Tschedul.Scored0, dbo.Tschedul.Scored1, dbo.Tschedul.Scored2, dbo.Tschedul.Scored3, 
'                         dbo.Tschedul.Scored4, dbo.Tschedul.Scored5, dbo.Tschedul.TKitType, dbo.Tschedul.TO1ApDrive, dbo.Tschedul.TO1ApDrRate, dbo.Tschedul.TO1ApJudge, 
'                         dbo.Tschedul.TO1ApJRate, dbo.Tschedul.TO1ApScore, dbo.Tschedul.TO1ApSRate, dbo.Tschedul.TO2ApDrive, dbo.Tschedul.TO2ApDrRate, dbo.Tschedul.TO2ApJudge, 
'                         dbo.Tschedul.TO2ApJRate, dbo.Tschedul.TO3ApJudge, dbo.Tschedul.TO3ApJRate, dbo.Tschedul.TO4ApJudge, dbo.Tschedul.TO4ApJRate, dbo.Tschedul.TO5ApJudge, 
'                         dbo.Tschedul.TO5ApJRate, dbo.Tschedul.TOCDRate, dbo.Tschedul.TOCDriver, dbo.Tschedul.TOCJRate, dbo.Tschedul.TOCJudge, dbo.Tschedul.TOCSafety, 
'                         dbo.Tschedul.TOCSafRate, dbo.Tschedul.TOCScore, dbo.Tschedul.TOCSRate, dbo.Tschedul.TOOoAJudge, dbo.Tschedul.TOOoAJRate, dbo.Tschedul.TOPanAmJRate, 
'                         dbo.Tschedul.TOPanAmJudge, dbo.Tschedul.TOTechCont, dbo.Tschedul.TOTechCRate, dbo.Tschedul.TRegistrarName, dbo.Tschedul.TRegistrarPhone, 
'                         dbo.Tschedul.TRegistrarEmail, dbo.Tschedul.TRegistrarAddr, dbo.Tschedul.TRegistrarCity, dbo.Tschedul.TRegistrarState, dbo.Tschedul.TRegistrarZip, 
'                         dbo.Tschedul.TRegistrarFax, dbo.Tschedul.TFirstTourn, dbo.Tschedul.TSEvJNo, dbo.Tschedul.TTEvJNo, dbo.Tschedul.TSigner, dbo.Tschedul.TsTitle, dbo.Tschedul.TsDate, 
'                         dbo.Tschedul.TsHPhone, dbo.Tschedul.TsBPhone, dbo.Tschedul.TsEmail, dbo.Tschedul.TRegion, dbo.Tschedul.LastSaved, dbo.Tschedul.THowShipped, 
'                         dbo.Tschedul.TShipped0, dbo.Tschedul.TShipped1, dbo.Tschedul.KitSent, dbo.Tschedul.T5Star, dbo.Tschedul.WaterSkier, dbo.Tschedul.Deleted, 
'                         dbo.Tschedul.PASanApproved, dbo.Tschedul.PAFee, dbo.Tschedul.PAFeeDetail, dbo.Tschedul.PACanceled, dbo.Tschedul.PAFeeRefund, dbo.Tschedul.PARefundDetail, 
'                         dbo.Tschedul.PAPDPanAm, dbo.Tschedul.PAPDPanAmAmt, dbo.Tschedul.PAPDDetail, dbo.Tschedul.TOCJudgeC, dbo.Tschedul.TOCDriverC, dbo.Tschedul.TOCScoreC, 
'                         dbo.Tschedul.TOTechContC, dbo.Tschedul.TEventClin, dbo.Tschedul.JDClin, dbo.Tschedul.ADClin, dbo.Tschedul.CDescription, dbo.Tschedul.ClinNumParticipants, 
'                         dbo.Tschedul.ClinLevel, dbo.Tschedul.HQClinFee, dbo.Tschedul.AllowAccess, dbo.Tschedul.RegnFee, dbo.Tschedul.IWSFFee, dbo.Tschedul.TOAnnounce, 
'                         dbo.Tschedul.GrassRoots, dbo.Tschedul.TEventWake, dbo.Tschedul.TRoundsWakeBd, dbo.Tschedul.TWakeBdPerRnd, dbo.Tschedul.BoatWake, dbo.Tschedul.CableWake, 
'                         dbo.Tschedul.TEventWSkate, dbo.Tschedul.TRoundsWSkate, dbo.Tschedul.TWSkatePerRnd, dbo.Tschedul.BoatSkate, dbo.Tschedul.CableSkate, dbo.Tschedul.TEventWSurf, 
'                         dbo.Tschedul.TRoundsWSurf, dbo.Tschedul.TWSurfPerRnd, dbo.Tschedul.WDescription, dbo.Tschedul.TEventFHF, dbo.Tschedul.TEventFKB, dbo.Tschedul.TEventFDA, 
'                         dbo.Tschedul.TEventF3ev, dbo.Tschedul.TEventFB, dbo.Tschedul.TEventFW, dbo.Tschedul.KEventFlip, dbo.Tschedul.KRoundsFlip, dbo.Tschedul.KFlipPerRnd, 
'                         dbo.Tschedul.KFlipClassQ, dbo.Tschedul.KFlipClassT, dbo.Tschedul.KEventFree, dbo.Tschedul.KRoundsFree, dbo.Tschedul.KFreePerRnd, dbo.Tschedul.KFreeClassQ, 
'                         dbo.Tschedul.KFreeClassT, dbo.Tschedul.KDescription, dbo.Tschedul.KSlalomClassQ, dbo.Tschedul.KSlalomClassT, dbo.Tschedul.KTrickClassT, dbo.Tschedul.KTrickClassQ, 
'                         dbo.Tschedul.SISposted, dbo.Tschedul.Qfy_Placement, dbo.Tschedul.Qfy_Rank_By_CutOffDate, dbo.Tschedul.Qfy_Rank_After_CutOffDate, 
'                         dbo.Tschedul.Qfy_ScoreRankLevel_At_Qualifier, dbo.Tschedul.Qfy_Natls_Place, dbo.Tschedul.Qfy_Regls_Place, dbo.Tschedul.Qfy_Rank_Level, 
'                         dbo.Tschedul.Qfy_CutOffDate, dbo.Tschedul.UseOLReg, dbo.Tschedul.AdminCode, dbo.Tschedul.OLR_pd, dbo.Tschedul.Gr1AWSPulls, dbo.Tschedul.Gr1ABCPulls, 
'                         dbo.Tschedul.Gr1USWPulls, dbo.Tschedul.Gr1AKAPulls, dbo.Tschedul.Gr1USHPulls, dbo.Tschedul.Gr1WSDPulls, dbo.Tschedul.Gr2AWS_SPulls, dbo.Tschedul.Gr2AWS_TPulls, 
'                         dbo.Tschedul.Gr2ABC_SPulls, dbo.Tschedul.Gr2ABC_TPulls, dbo.Tschedul.Gr2USW_WPulls, dbo.Tschedul.Gr2USW_SkatePulls, dbo.Tschedul.Gr2USW_SurfPulls, 
'                         dbo.Tschedul.Gr2USW_RailJamPulls, dbo.Tschedul.Gr2AKA_SPulls, dbo.Tschedul.Gr2AKA_TPulls, dbo.Tschedul.Gr2AKA_FreePulls, dbo.Tschedul.Gr2AKA_FlipPulls, 
'                         dbo.Tschedul.Gr2USH_FreeRidePulls, dbo.Tschedul.Gr2USH_JumpOutPulls, dbo.Tschedul.Gr2USH_FlipOutPulls, dbo.Tschedul.Gr2USH_BigAirPulls, 
'                         dbo.Tschedul.Gr2USH_3TrickPulls, dbo.Tschedul.TOCSafetyID, dbo.Tschedul.TOCDriverID, dbo.Tschedul.GrBoat, dbo.Tschedul.GrCable, dbo.Tschedul.AcceptTerms, 
'                         dbo.Tschedul.GrFunDay, dbo.Tschedul.GrTournament, dbo.Tschedul.GrFee_1, dbo.Tschedul.GrFee_2, dbo.Tschedul.GrFee_3, dbo.Tschedul.GrFee_4, dbo.Tschedul.GrFee_5, 
'                         dbo.Tschedul.OK2Publish, dbo.Tschedul.PreSafetyOK, dbo.Tschedul.OfficialsOK, dbo.Tschedul.ClubOK, dbo.Tschedul.HQSanFeesPd, dbo.Tschedul.HQMessage, 
'                         dbo.Tschedul.HQItem1Desc, dbo.Tschedul.HQItem2Desc, dbo.Tschedul.HQItem3Desc, dbo.Tschedul.HQItem4Desc, dbo.Tschedul.HQ1Qty, dbo.Tschedul.HQ2Qty, 
'                         dbo.Tschedul.HQ3Qty, dbo.Tschedul.HQ4Qty, dbo.Tschedul.HQ1MaxQty, dbo.Tschedul.HQ2MaxQty, dbo.Tschedul.HQ3MaxQty, dbo.Tschedul.HQ4MaxQty, 
'                          dbo.Registration.EntryFee1, 
'                         dbo.Registration.EntryFee2, dbo.Registration.EntryFee3, dbo.Registration.EntryFee4, dbo.Registration.EntryFee5, dbo.Registration.EntryFeeFamily, 
'                         dbo.Registration.FamEntType, dbo.Registration.LateFee, dbo.Registration.OtherFee, dbo.Registration.EmailAddress, dbo.Registration.ReceiveEmail, dbo.Registration.Disc1, 
'                         dbo.Registration.Disc2, dbo.Registration.Disc3, dbo.Registration.Disc4, dbo.Registration.ClubCode, dbo.Registration.DiscMeth, dbo.Registration.GrEntryFee1, 
'                         dbo.Registration.GrEntryFee2, dbo.Registration.GrEntryFee3, dbo.Registration.GrEntryFee4, dbo.Registration.GrEntryFee5, dbo.Registration.ClinFeeJD, 
'                         dbo.Registration.ClinFeeAD, dbo.Registration.MaxPulls AS Expr4, dbo.Registration.ReservedPulls, dbo.Registration.ReservedPullsCode, dbo.Registration.CMulti, 
'                         dbo.Registration.EMulti, dbo.Registration.LMulti, dbo.Registration.RMulti, dbo.Registration.MixedOptions, dbo.Registration.CSurchg, dbo.Registration.ESurchg, 
'                         dbo.Registration.LSurchg, dbo.Registration.RSurchg, dbo.Registration.GREntryFeeIncluded, dbo.Registration.GRDiscount, dbo.Registration.ClinicDiscount, 
'                         dbo.Registration.Bio_Reqd, dbo.Registration.TrkLst_Reqd, dbo.Registration.AthContr_Reqd, dbo.Registration.USAWSWaiver_Reqd, dbo.Registration.QualLevel, 
'                         dbo.Registration.RestrictDv AS Expr5, dbo.Registration.BTickCost, dbo.Registration.BTickWithE, dbo.Registration.BTickFamily, dbo.Registration.Form1Name, 
'                         dbo.Registration.Form2Name, dbo.Registration.Form3Name, dbo.Registration.Form4Name, dbo.Registration.Form5Name, dbo.Registration.Form6Name, 
'                         dbo.Registration.EVENT1, dbo.Registration.EVENT2, dbo.Registration.EVENT3, dbo.Registration.EVENT4, dbo.Registration.QualLevel1, dbo.Registration.RegisterDate, 
'                         dbo.Registration.PayPalAct, dbo.Registration.PayPalOK, dbo.Registration.AllowEdit, dbo.Registration.AWSEFDon_OK, dbo.Registration.MaxSlPulls, 
'                         dbo.Registration.MaxTrPulls, dbo.Registration.MaxJpPulls, dbo.Registration.OF1Desc, dbo.Registration.OF2Desc, dbo.Registration.OF3Desc, dbo.Registration.OF4Desc, 
'                         dbo.Registration.OF5Desc, dbo.Registration.OF6Desc, dbo.Registration.OF1Amt, dbo.Registration.OF2Amt, dbo.Registration.OF3Amt, dbo.Registration.OF4Amt, 
'                         dbo.Registration.OF5Amt, dbo.Registration.OF6Amt, dbo.Registration.OF1MaxQty, dbo.Registration.OF2MaxQty, dbo.Registration.OF3MaxQty, dbo.Registration.OF4MaxQty, 
'                         dbo.Registration.OF5MaxQty, dbo.Registration.OF6MaxQty, dbo.Registration.OF1Required, dbo.Registration.OF2Required, dbo.Registration.OF3Required, 
'                         dbo.Registration.OF4Required, dbo.Registration.OF5Required, dbo.Registration.OF6Required, dbo.Registration.GRClassText, dbo.Registration.BaseClassText, 
'                         dbo.Registration.UpgradeClassText, dbo.Registration.TEntryFeeFamExtra, dbo.Registration.MaxFamMember,  dbo.Registration.OLRDisplayStatus, dbo.GuideBk.GTournAppId, dbo.GuideBk.GTSofE, 
'                         dbo.GuideBk.GTAwards, dbo.GuideBk.GTStartTime, dbo.GuideBk.GTPractice, dbo.GuideBk.GTSDirections, dbo.GuideBk.GTAccommodation, dbo.GuideBk.GTComments, 
'                         dbo.GuideBk.GTNotes, dbo.GuideBk.GTText1

'FROM         dbo.Registration INNER JOIN
'                      dbo.Tschedul ON dbo.Registration.TournAppID = dbo.Tschedul.TournAppID
'  INNER JOIN dbo.GuideBk on dbo.Tschedul.TournAppID = dbo.GuideBk.GTournAppID
'WHERE Tschedul.TournAppID = @TAID AND Tschedul.TSanType <> 2 AND Tschedul.TSTATUS > 1













%>
