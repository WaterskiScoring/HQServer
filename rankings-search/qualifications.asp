<%





' ------------------------------------------------------------------------------------------------------------------------------
   SUB MembershipStatus (sMemberOverride, sCanSkiInTournaments, sCanSkiInGRTournaments, sFeeClassEvt, sTDateE, sEffectiveTo, sMemberType, sMembTypeDesc)
' ------------------------------------------------------------------------------------------------------------------------------

' Uses variables - sMemberOverride sCanSkiInTournaments sTDateE sEffectiveTo, sMemberType, sMembTypeDesc

'response.write("<br>sFeeClassEvt = "&sFeeClassEvt)
'response.write("<br>sCanSkiInGRTournaments = "&sCanSkiInGRTournaments)

' --- Test Membership Types and Expiration Date of Membership ---
MembStatusTitle=""
MembStatusText = ""
MembStatusColor = ""

IF sMembOverride <> "" THEN
       	MembStatusTitle="Administrative Override of Membership Status &#13; Override Code: "&sMembOverride
	MembStatusText = "(OK)"
	MembStatusColor = TextColor2
	MembStatusOK=true
ELSEIF sCanSkiInTournaments = "True" AND DateDiff("d", sTDateE, sEffectiveTo) > 0 THEN 
	MembStatusTitle = "Membership Good Thru - "&sEffectiveTo&" &#13; Member Type - "&sMemberType&" - "&sMembTypeDesc  
	MembStatusText = "OK"
	MembStatusColor = TextColor2

ELSEIF sCanSkiInGRTournaments = "True" AND sFeeClassEvt = "G" AND DateDiff("d", sTDateE, sEffectiveTo) > 0 THEN 
	MembStatusTitle = "Membership Good Thru - "&sEffectiveTo&" &#13; Member Type - "&sMemberType&" - "&sMembTypeDesc  
	MembStatusText = "OK(G)"
	MembStatusColor = TextColor2
	' --- From Registration test - sGRTournament=true OR sGRFunDay=true
ELSE
	MembStatusColor = TextColor3
	ScratchCount = ScratchCount + 1
	IF sCanSkiInTournaments <> "True" THEN 
		MembStatusTitle = "Not eligible to Compete - Membership Type = "&sMemberType
		MembStatusText = "MT"
	END IF
	IF DateDiff("d", sTDateE, sEffectiveTo) < 0 THEN	
		MembStatusTitle = MembStatusTitle + " Membership Expired on "&sEffectiveTo 
		MembStatusText = MembStatusText = "EXP"
	END IF
END IF

END SUB



' -----------------------
   SUB MembershipStatus2
' -----------------------

' +++++++++++++++++++++++++++++++++++++++++
' +++ NOT USED BY VIEW-REGISTRATION.asp +++
' +++++++++++++++++++++++++++++++++++++++++

' --- Test Membership Types and Expiration Date of Membership ---
MembStatusTitle=""
MembStatusText = ""
MembStatusColor = ""

IF TRIM(rs("MembOverride")) <> "" THEN
       	MembStatusTitle="Administrative Override of Membership Status &#13; Override Code: "&rs("MembOverride")
	MembStatusText = "(OK)"
	MembStatusColor = TextColor2
ELSEIF rs("CanSkiInTournaments") = "True" AND TS <> "NO" AND DateDiff("d", sTDateE, rs("EffectiveTo")) > 0 THEN 
	MembStatusTitle = "Membership Good Thru - "&rs("EffectiveTo")&" &#13; Member Type - "&rs("MemberType")&" - "&rs("MembTypeDesc")  
	MembStatusText = "OK"
	MembStatusColor = TextColor2
ELSE
	MembStatusColor = TextColor3
	ScratchCount = ScratchCount + 1
      		IF rs("CanSkiInTournaments") <> "True" OR TS = "NO" THEN 
		MembStatusTitle = "Not eligible to Compete - Membership Type = "&rs("MemberType")
		MembStatusText = "MT"
	END IF
	IF DateDiff("d", sTDateE, rs("EffectiveTo")) < 0 THEN	
		MembStatusTitle = MembStatusTitle + "Membership Expired on "&rs("EffectiveTo") 
		MembStatusText = MembStatusText = "EXP"
	END IF
END IF

END SUB




' ----------------------------------------------------------------------------
   SUB PaymentStatus2 (sPayments, sMoneyOverride, sTotalEntry, sEntryType, sMemberID)
' ----------------------------------------------------------------------------

' --- Tests if all FEES are paid ---
FeesText = ""
FeesColor = ""
FeesTitle = ""


' --- These session variables are set in tools_registration.asp ---
' 	Session("sWhichFamilyMemberPaid") - specifies which person paid the family entry fee.
' 	Session("TotRegisteredFamMembers") - Determines how many family members there are registered now.


' --- Here for reference from SWIFT settings ---
' sTEntryFeeFamily
' sTEntryFeeFamExtra
' sMaxFamMembers



'IF Session("AdminMenuLevel")>=50 AND sMemberID="100041664" THEN
'	response.write("<br>sEntryType = "&sEntryType)
'	response.write("<br>sTotalEntry = "&sTotalEntry)
'	response.write("<br>sPayments = "&sPayments)
'	response.write("<br>sTEntryFeeFamily = "&sTEntryFeeFamily)
'	response.write("<br>WFMP = "&Session("sWhichFamilyMemberPaid"))
'END IF




' ---- Family Entry and this person is the first to enter ---
IF sEntryType="FAM" AND cdbl(sTotalEntry) = cdbl(sPayments) AND cdbl(sTotalEntry) >= sTEntryFeeFamily THEN
	FeesText = "OK-1"
	FeesColor = TextColor2
	'FeesTitle = "Actual Payments = "&sPayments&"&#13; Family Entry - First Entry Paid = $"&sTEntryFeeFamily
	FeesTitle = "Actual Payments = "&sPayments&"&#13; Family Entry - First Entry Paid = $"&sTotalEntry

' ---- Family Entry and NOT the first to enter AND (less than MaxFamMembers or no MaxFamMembers specified) ---
ELSEIF sEntryType="FAM" AND cdbl(sTotalEntry) = cdbl(sPayments) AND cdbl(sTotalEntry) = 0 THEN
	FeesText = "OK-0"
	FeesColor = TextColor2
	FeesTitle = "Actual Payments = "&sPayments&"&#13; Family Entry - This Member is FREE = $0.00"

' ---- Family Entry and NOT the first to enter AND MORE than MaxFamMembers ---
ELSEIF sEntryType="FAM" AND cdbl(sTotalEntry) = cdbl(sPayments) AND cdbl(sTotalEntry) =sTEntryFeeFamExtra THEN
	FeesText = "OK-A"
	FeesColor = TextColor2
	FeesTitle = "Actual Payments = "&sPayments&"&#13; Family Entry - This Member Pays Additional Fee = $"&sTEntryFeeFamExtra

' ---- Individual entry and this person has paid more than required ---
ELSEIF cdbl(sTotalEntry) < cdbl(sPayments) THEN
	IF adminmenulevel >= 20 OR TestValidAdminCode=true THEN
		FeesText = "("&FormatCurrency( cdbl(sPayments) - cdbl(sTotalEntry) )&")"
	ELSE
		FeesText = "(OK)"
	END IF

	FeesColor = TextColor3
	FeesTitle = "Actual Payments = "&sPayments&"&#13; Possible Overpayment - Fees Due = $"&sTotalEntry

' ---- Payments match fees
ELSEIF cdbl(sTotalEntry) = cdbl(sPayments) THEN		
	FeesText = "OK"
	FeesColor = TextColor2
	FeesTitle = "Actual Payments = "&FormatCurrency(sPayments)&"&#13; Balance Due = $0.00"

' --- Fees are greater than payments
ELSEIF cdbl(sTotalEntry) > cdbl(sPayments) THEN		 
	IF adminmenulevel >= 20 OR TestValidAdminCode=true THEN
		FeesText = FormatCurrency( cdbl(sTotalEntry) - cdbl(sPayments) )  
	ELSE
		FeesText = "$$"
	END IF
	FeesColor = TextColor3
	FeesTitle = "Total Fees & Charges = "&FormatCurrency(sTotalEntry)&" &#13; Payments Made = "&FormatCurrency(sPayments)&" &#13; Balance Due = "&FormatCurrency(cdbl(sTotalEntry) - cdbl(sPayments))
	ScratchCount = ScratchCount + 1

' ---- ADMIN OVERRIDE ---
ELSEIF TRIM(rs("MoneyOverride"))<>"" THEN
	FeesText = "(OK)"
	FeesColor = TextColor2
	FeesTitle = "Actual Payments = "&sPayments&"&#13; Entry Fee - Administrative Override &#13; Actual Payments = "&FormatCurrency(sPayments)

END IF



END SUB





' ----------------------------------------------------------------------------
   SUB PaymentStatus (sPayments, sMoneyOverride, sTotalEntry, sEntryType)
' ----------------------------------------------------------------------------


' +++++++++++++++++  OBSOLETE   ++++++++++++++++++++++++++


' --- Tests if all FEES are paid ---
FeesText = ""
FeesColor = ""
FeesTitle = ""



' ---- Family Entry and HOH so fees should be $345 or greater - So check if entry was made with individual payments - REMOVE LATER ?
IF TRIM(rs("MoneyOverride"))<>"" AND cdbl(sPayments)<345 AND sEntryType="HOH" THEN
	FeesText = "(OK)"
	FeesColor = TextColor3
	FeesTitle = "Actual Payments = "&sPayments&"&#13; Family Entry - HOH = $345.00"

' ---- Family Entry and MEM so fees should be $0 - So check if entry was made with individual payments - REMOVE LATER ?
ELSEIF TRIM(rs("MoneyOverride"))<>"" AND cdbl(sPayments)<>0 AND sEntryType="MEM" THEN
	FeesText = "(OK)"
	FeesColor = TextColor3
	FeesTitle = "Actual Payments = "&sPayments&"&#13; Family Entry - MEM = $0.00"

' ---- Family Entry and HOH so fees should be $345 or greater
ELSEIF TRIM(rs("MoneyOverride"))<>"" AND sEntryType="HOH" THEN
	FeesText = "(OK)"
	FeesColor = TextColor2
	FeesTitle = "Actual Payments = "&sPayments&"&#13; Family Entry - HOH = $345.00"

' ---- Family Entry and MEM so fees should be $0
ELSEIF TRIM(rs("MoneyOverride"))<>"" AND sEntryType="MEM" THEN
	FeesText = "(OK)"
	FeesColor = TextColor2
	FeesTitle = "Actual Payments = "&sPayments&"&#13; Family Entry - Member = $0.00"


' ---- Individual entry so fees should be $120 and this person has paid more
ELSEIF cdbl(sTotalEntry) < cdbl(sPayments) THEN
	FeesText = "(OK)"
	FeesColor = TextColor3
	FeesTitle = "Actual Payments = "&sPayments&"&#13; Possible Overpayment - Fees = $"&sTotalEntry

' ---- Payments match fees
ELSEIF cdbl(sTotalEntry) = cdbl(sPayments) THEN		
	FeesText = "OK"
	FeesColor = TextColor2
	FeesTitle = "Actual Payments = "&FormatCurrency(sPayments)

' --- Fees are greater than payments
ELSEIF cdbl(sTotalEntry) > cdbl(sPayments) THEN		 
	FeesText = "$$"
	FeesColor = TextColor3
	FeesTitle = "Total Fees & Charges = "&FormatCurrency(sTotalEntry)&" &#13; Payments Made = "&FormatCurrency(sPayments)&" &#13; Balance Due = "&FormatCurrency((cdbl(sTotalEntry) - cdbl(sPayments)))
	ScratchCount = ScratchCount + 1
END IF

END SUB



' ----------------------------
  SUB RequiredParticipation (sReglPartStat)
' ----------------------------

' --- Possible values of sReglPartStat are
' ---  O = Open
' ---  X = Override
' ---  - = Not Required for this League
' ---  W,E,S,C or M for specific Regionals
' ---  blank

'IF Session("AdminMenuLevel")>=50 AND sMemberID="100040827" THEN
'	response.write("sRequirePart="&sRequirePart)
'	response.write("sQualLevel="&sQualLevel)
'	response.write("sReglPartStat="&sReglPartStat)
'END IF

'		IF sMemberID="400124481" THEN
'				response.write("<br>Line 270 qualificaitons.asp sRequirePart = "& sRequirePart )
'				response.write("<br>sReglPartStat = "& sReglPartStat )		
'				response.write("<br>sQualLevel= "& sQualLevel )		
'		END IF
		



' --- Represents OTHER possible conditions W,E,S,M,C ---
IF sReglPartStat="S" OR sReglPartStat="M" OR sReglPartStat="C" OR sReglPartStat="E" OR sReglPartStat="W" THEN
		RegStatusColor = TextColor2
		RegStatusText = "OK"
		RegStatusTitle="Skied "&sReglPartStat&" Regionals"

ELSEIF sReglPartStat="X" THEN
		RegStatusColor = TextColor2
		RegStatusText = "(OK)"
    RegStatusTitle="Administrative Override of participation requirement"

ELSEIF sReglPartStat="O" THEN
		RegStatusColor = TextColor2
		RegStatusText = "(OK)"
    RegStatusTitle="Participation Not Required &#13; in Open Division"

ELSEIF sReglPartStat="P" THEN
		RegStatusColor = TextColor2
		RegStatusText = "Pending"
		RegStatusTitle="Must ski in Regionals in this Event"	

ELSEIF  sReglPartStat="" THEN
		RegStatusColor = TextColor3
		RegStatusText = "DNS"
		RegStatusTitle="Did not ski Regionals in this Event"	
		ScratchCount = ScratchCount + 1

ELSEIF sQualLevel=0 OR sRequirePart="-" THEN
		RegStatusColor = TextColor2
		RegStatusText = "---"
    RegStatusTitle="No other tournament participation required"


END IF 




END SUB



' -------------------------------------------------------------------------
   SUB RegionalParticipation (sReg_Ski, sRegionalOverride, div, sTDateE, sHomeRegion)
' -------------------------------------------------------------------------

' --- Checks for Participation in Current SKI YEAR REGIONALS  ---
' ---- Must make this look for the CURRENT regionals  ----
' ---- Current code is a band-aid

RegStatusColor = ""
RegStatusText = ""
RegStatusTitle=""
		
ThisDate = DATE
ThisRegion = TRIM(sReg_Ski)

' --- Tournament does not have qualifications ---



IF RIGHT(sTSanction,1)="B" THEN
	RegStatusColor = TextColor2
	RegStatusText = "---"
       	RegStatusTitle="Participation in State tournament not required &#13;"


ELSEIF sQualLevel=0 THEN
	RegStatusColor = TextColor2
	RegStatusText = "---"
       	RegStatusTitle="Regionals participation not required &#13;"
	IF ThisRegion<> "" THEN
		RegStatusTitle=RegStatusTitle + " Skied in Regionals Code: "&ThisRegion
	ELSE
		RegStatusTitle=RegStatusTitle + " Did not ski Regionals"	
	END IF

' --- Administrative Override ---
ELSEIF TRIM(sRegionalOverride) <> "" THEN
	RegStatusColor = TextColor2
	RegStatusText = "(OK)"
       	RegStatusTitle="Administrative Override of Regional Participation &#13; Override Code: "&sRegionalOverride

ELSEIF LEFT(div,1)="O" THEN
	RegStatusColor = TextColor2
	RegStatusText = "(OK)"
       	RegStatusTitle="Regional Participation Not Required &#13; in Open Division"

' -- Temporary fix
ELSEIF sReg_Ski <> "" AND (ThisRegion = "W" OR ThisRegion = "S" OR ThisRegion = "M" OR ThisRegion = "E" OR ThisRegion = "C") THEN
	RegStatusColor = TextColor2
	RegStatusText = "OK"
	RegStatusTitle="Skied "&sReg_Ski&" Regionals"

' --- Based on Nationals End Date of 8/18/07 and Tue after Regionals of 7/30/2007 ---
'ELSEIF ThisRegion = "C" THEN
'	RegStatusColor = TextColor2
'	RegStatusText = "Pending"
'      	RegStatusTitle="Must Compete in Regionals"
'	ScratchCount = ScratchCount + 1

ELSEIF ThisRegion = "" AND (ThisRegion = "S" AND ThisRegion = "W") THEN
	' ---- THIS WOULD BE AN EASIER TEST IF LOGIC WAS CORRECT   ----  rs("Reg_Ski") = NULL OR ThisRegion = "" THEN
	RegStatusColor = TextColor3
	RegStatusText = "***"
        RegStatusTitle="Did not ski in a Regionals"
	ScratchCount = ScratchCount + 1


ELSE
	RegStatusColor = TextColor3
	'RegStatusText = "OK"
	'RegStatusTitle="Skied "&sReg_Ski&" Regionals"	

ScratchCount = ScratchCount + 1
	'RegStatusText = "Pending"
RegStatusText = "DNS"
	'RegStatusTitle="Must ski in Regionals in this Event"	
RegStatusTitle="Did not ski Regionals in this Event"	

END IF 




END SUB



' -------------------------------------------------------
  SUB ChkQualALL2 (sTourID, sMemberID, sEvent, sDiv, subQfyOverride)
' -------------------------------------------------------



Set rsQfy=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT RQ.QfyStatus"
sSQL = sSQL + " FROM "&RegQualifyTableName&" AS RQ"
sSQL = sSQL + " WHERE RQ.MemberID = '"&sMemberID&"' AND LEFT(RQ.TourID,6)='"&LEFT(sTourID,6)&"' AND RQ.Div='"&sDiv&"' AND RQ.Event='"&sEvent&"'"
rsQfy.open sSQL, sConnectionToTRATable, 3, 1

IF sMemberID="700041969" THEN 
'	response.write(sSQL)

'	response.write("rs="&rsQFy("QfyStatus"))
END IF





QfyStatusText(1)=""

IF sQualLevel=0 THEN 
		QfyStatusText(1)="---"
		QfyStatusTitle(1)="No qualifications required for this tournament"
ELSEIF TRIM(subQfyOverride) = "DNS" THEN
		QfyStatusColor(EvtNo) = TextColor3
		QfyStatusText(EvtNo) = "Scratch"
		QfyStatusLongText(EvtNo) = "&nbsp; DNS - Scratch: "&subQfyOverride
   	QfyStatusTitle(EvtNo)="Administrative Override of Qualifications &#13; Override Code: "&subQfyOverride
		ScratchCount = ScratchCount + 1
ELSEIF TRIM(subQfyOverride) <> "" THEN
		QfyStatusColor(EvtNo) = TextColor2
		QfyStatusText(EvtNo) = "(OV)"
		QfyStatusLongText(EvtNo) = "&nbsp; OK - Override: "&subQfyOverride
    QfyStatusTitle(EvtNo)="Administrative Override of Qualifications &#13; Override Code: "&subQfyOverride

ELSEIF NOT rsQfy.eof THEN 

		QfyStatusText(1)=TRIM(rsQFy("QfyStatus"))

		SELECT CASE TRIM(QfyStatusText(1))
			CASE "Qualified", "Qfy-RO"
					QfyStatusColor(1)=TextColor2
					QfyStatusTitle(EvtNo)=TRIM(QfyStatusText(1))
			CASE "Pending", "QFY-RPR"
					QfyStatusColor(1)="Orange"
					QfyStatusTitle(EvtNo)="Qualification for this event is Pending and will LOCK at the Cut-off-Date for this tournament."
			CASE "NCQ"
					QfyStatusColor(1)="red"
					QfyStatusTitle(EvtNo)="Not Currently Qualified by any qualification method."
					ScratchCount = ScratchCount + 1			
			CASE ELSE
					QfyStatusColor(EvtNo)=TextColor2
					QfyStatusTitle(EvtNo)="Unknown ELSE"
					ScratchCount = ScratchCount + 1
		END SELECT 

ELSEIF rsQfy.eof THEN 
		QfyStatusText(1)="---"
		QfyStatusTitle(1)="No qualifications found for this event."

END IF

END SUB



' -------------------------------------------------------
  SUB ChkQualALL (sTourID, sMemberID, sEvent, sDiv)
' -------------------------------------------------------



Set rsQfy=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT RQ.QfyStatus"
sSQL = sSQL + " FROM "&RegQualifyTableName&" AS RQ"
sSQL = sSQL + " WHERE RQ.MemberID = '"&sMemberID&"' AND LEFT(RQ.TourID,6)='"&LEFT(sTourID,6)&"' AND RQ.Div='"&sDiv&"' AND RQ.Event='"&sEvent&"'"

rsQfy.open sSQL, sConnectionToTRATable, 3, 1


QfyStatusText(1)=""

IF sQualLevel=0 THEN 
	QfyStatusText(1)="---"
	QfyStatusTitle(1)="No qualifications required for this tournament"


ELSEIF NOT rsQfy.eof THEN 
	QfyStatusText(1)=rsQFy("QfyStatus")

	SELECT CASE TRIM(QfyStatusText(1))
		CASE "Qualified"
			QfyStatusColor(1)=TextColor2
			QfyStatusTitle(EvtNo)=TRIM(QfyStatusText(1))
		CASE "Pending", "QFY-RPR"
			QfyStatusColor(1)="Orange"
			QfyStatusTitle(EvtNo)="Qualification for this event is Pending and will LOCK at the Cut-off-Date for this tournament."
		CASE "NCQ"
			QfyStatusColor(1)="red"
			QfyStatusTitle(EvtNo)="Not Currently Qualified by any qualification method."
		CASE ELSE
			QfyStatusColor(EvtNo)=TextColor2
			QfyStatusTitle(EvtNo)="Unknown ELSE"
	END SELECT 

ELSEIF rsQfy.eof THEN 
	QfyStatusText(1)="---"
	QfyStatusTitle(1)="No qualifications found for this event."

END IF

END SUB




' -----------------------------------------------------------------------------------------------------------------------
   SUB CheckNationalQualificationsNew (sMemberID, sEvent, sDiv, subQfyOverride, EvtNo)
' -----------------------------------------------------------------------------------------------------------------------


sEvent=TRIM(sEvent)


' --- IF the current event is a Masters check to see if this event qualifies by 3rd event
Set rsQfy=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT RT.Natl_Plc, RT.Regl_Plc, RT.AWSA_Rat"
sSQL = sSQL + " FROM "&RankTableName&" AS RT"
sSQL = sSQL + " WHERE RT.MemberID = '"&sMemberID&"' AND RT.Div='"&sDiv&"' AND RT.Event='"&sEvent&"' AND RT.sc_1>0"
sSQL = sSQL + " AND SkiYearID=(SELECT TOP 1 SkiYearID FROM usawsrank.SkiYear WHERE BeginDate<=CAST('"&sTDateE&"' AS DateTime) AND EndDate>=CAST('"&sTDateE&"' AS DateTime) ORDER BY SkiYearID DESC)" 

rsQfy.open sSQL, sConnectionToTRATable, 3, 1


IF NOT rsQfy.eof THEN
	sNatl_Plc=rsQfy("Natl_Plc")
	sRegl_Plc=rsQfy("Regl_Plc")
	sRating=rsQfy("AWSA_Rat")
ELSE

	sNatl_Plc=""
	sRegl_Plc=""
	sRating=""
END IF




' +++++++  REMOVE AFTER DEBUGGING  +++++++++

sRank=10


Dim EventLoop
FOR EventLoop=1 TO TotEv
	QualStatusEvent(EventLoop)=""
	QfyStatusColor(EventLoop) = TextColor1
	QfyStatusText(EventLoop) = ""
	QfyStatusTitle(EventLoop)=""
NEXT



' --- Removes TIES from National placement to allow testing
TPosNatl = InStr(lcase(sNatl_Plc), "t") - 1

IF TRIM(sNatl_Plc)<>"" THEN
	IF TPosNatl > 0 THEN
		sNatl_Plc = (Left(sNatl_Plc, TPosNatl))
	END IF
ELSE
	sNatl_Plc=999
END IF


' --- Removes TIES from Regional placement to allow testing
TPosRegl = InStr(lcase(sRegl_Plc), "t") - 1
IF TRIM(sRegl_Plc)<>"" THEN
	IF TPosRegl > 0 THEN
		sRegl_Plc = Left(sRegl_Plc, TPosRegl)
	END IF
ELSE 
	sRegl_Plc=999
END IF


'response.write("<br>SE("&EvtNo&")="&sSelectEvent(EvtNo))
'response.write("<br>sDiv="&sDiv)


IF LEFT(sTourID,6)="08W167" THEN

  
  IF NOT (TRIM(sRating)="") THEN
    ' --- M3, M4 and M5 require Level 7 all others level 6 --- 	
    IF ( (TRIM(sDiv)="M3" OR TRIM(sDiv)="M4" OR TRIM(sDiv)="M5") AND cdbl(RIGHT(sRating,1))>=cdbl(7) ) OR ( TRIM(sDiv)<>"M3" AND TRIM(sDiv)<>"M4" AND TRIM(sDiv)<>"M5" AND cdbl(RIGHT(sRating,1))>=cdbl(6) ) THEN

				IF DateDiff("d", #06/24/2008#, Date) < 0 THEN 
						QfyStatusColor(EvtNo) = "#FFA500"		
						QfyStatusTitle(EvtNo)="Qualifications in this event are OK at this time, however&#13;final qualifications not locked-in until Cut-Off-Date&#13;Current Level - ("&RIGHT(srating,1)&")"
				ELSE	
						QfyStatusColor(EvtNo) = TextColor2
						QfyStatusTitle(EvtNo)="Qualifications = Level - ("&RIGHT(srating,1)&")"
				END IF
				QfyStatusText(EvtNo) = "OK"
				QfyStatusLongText(EvtNo) = "&nbsp; OK - Level ("&RIGHT(srating,1)&")"

    ELSE
				QfyStatusColor(EvtNo) = TextColor3
				QfyStatusText(EvtNo) = "* "&RIGHT(sRating,1)&" *"
				QfyStatusLongText(EvtNo) = "&nbsp; Not Currently Qualified = LEVEL - "&RIGHT(sRating,1)
				QfyStatusTitle(EvtNo)="Not Currently Qualified - LEVEL - "&RIGHT(srating,1)
				ScratchCount = ScratchCount + 1
    END IF
  END IF


ELSEIF sQualLevel=0 THEN
		QfyStatusColor(EvtNo) = TextColor2
		QfyStatusText(EvtNo) = "--"
		QfyStatusLongText(EvtNo) = "&nbsp; No Qualifications Required"
		QfyStatusTitle(EvtNo)="No Qualifications Required - Member Ranking Level ("&RIGHT(srating,1)&")"

ELSEIF LEFT(sRating,1)="O" THEN
		QfyStatusColor(EvtNo) = TextColor2
		QfyStatusText(EvtNo) = "OK"
		QfyStatusLongText(EvtNo) = "&nbsp; OK - Open Rating"
		QfyStatusTitle(EvtNo)="Current Rating = Open ("&RIGHT(srating,1)&")"

ELSEIF RIGHT(sRating,1)=TRIM(sQualLevel) THEN
		QfyStatusColor(EvtNo) = TextColor2
		QfyStatusText(EvtNo) = "OK"
		QfyStatusLongText(EvtNo) = "&nbsp; OK - Level ("&RIGHT(srating,1)&")"
		QfyStatusTitle(EvtNo)="Qualifications = Level - ("&RIGHT(srating,1)&")"

ELSEIF TRIM(subQfyOverride) <> "" THEN
		QfyStatusColor(EvtNo) = TextColor2
		QfyStatusText(EvtNo) = "(OV)"
		QfyStatusLongText(EvtNo) = "&nbsp; OK - Override: "&subQfyOverride
    QfyStatusTitle(EvtNo)="Administrative Override of Qualifications &#13; Override Code: "&subQfyOverride

ELSEIF cdbl(sRank) = 0 OR TRIM(sRating) = "0" OR TRIM(sRating) = "" THEN
		QfyStatusColor(EvtNo) = TextColor3
		QfyStatusText(EvtNo) = "***"
		QfyStatusLongText(EvtNo) = "&nbsp; No Current Rating"
    QfyStatusTitle(EvtNo)="No Current Rating"
		ScratchCount = ScratchCount + 1

' --- Drag the 3rd event providing 1st two are Level sQualLevel (8 for Nationals) and 3rd is Level sQualLevel-1 (7 for Nationals) ---
ELSEIF INT(RIGHT(sRating,1))=INT(TRIM(sQualLevel))-1 THEN

		Set rsQualify=Server.CreateObject("ADODB.recordset")
		sSQL = "SELECT RT.sc_1, RT.Event, RT.Div, RT.Natl_Plc, RT.Regl_Plc, RT.AWSA_Rat"
		sSQL = sSQL + " FROM "&RankTableName&" AS RT"
		sSQL = sSQL + " WHERE RT.MemberID = '"&sMemberID&"' AND RT.Div='"&sDiv&"' AND RT.sc_1>0"
		sSQL = sSQL + " AND SkiYearID=(SELECT TOP 1 SkiYearID FROM usawsrank.SkiYear WHERE BeginDate<=CAST('"&sTDateE&"' AS DateTime) AND EndDate>=CAST('"&sTDateE&"' AS DateTime) ORDER BY SkiYearID DESC)" 

		rsQualify.open sSQL, sConnectionToTRATable, 3, 1


		DO WHILE NOT rsQualify.EOF 	

				' --- Checks for EP or Open Rating ---
				IF UCASE(RIGHT(rsQualify("AWSA_Rat"),1)) = TRIM(sQualLevel) OR UCASE(LEFT(rsQualify("AWSA_Rat"),1)) = "O" THEN
						ThisEventStat = "OK"
				ELSE
						ThisEventStat = "Rating"
				END IF

				SELECT CASE TRIM(rsQualify("Event"))
			  	CASE "S"
							QualEvent(1) = rsQualify("AWSA_Rat")
							QualStatusEvent(1) = ThisEventStat
			  	CASE "T"
							QualEvent(2) = rsQualify("AWSA_Rat")
							QualStatusEvent(2) = ThisEventStat
			  	CASE "J"
							QualEvent(3) = rsQualify("AWSA_Rat")
							QualStatusEvent(3) = ThisEventStat
				END SELECT
		
				rsQualify.movenext

		LOOP


	' --- Examines for Open or LEVEL 8 in two events and Level 7 in the 3rd event ---
	IF (sEvent="S" AND QualStatusEvent(2) = "OK" AND QualStatusEvent(3) = "OK") OR (sEvent="T" AND QualStatusEvent(1) = "OK" AND QualStatusEvent(3) = "OK") OR (sEvent="J" AND QualStatusEvent(1) = "OK" AND QualStatusEvent(2) = "OK") THEN 
			QfyStatusColor(EvtNo) = TextColor2		
			QfyStatusText(EvtNo) = "(OK)"
			QfyStatusLongText(EvtNo) = "&nbsp; (OK) - By 3rd Event"
			QfyStatusTitle(EvtNo)="Current Level  = LEVEL - ("&RIGHT(srating,1)&") &#13; Qualified by 3rd Event"

	' --- Checks for Masters rating plus PREVIOUS Nationals Placement  ---
	ELSEIF sNatl_Plc <= 5 THEN
			QfyStatusColor(EvtNo) = TextColor2		
			QfyStatusText(EvtNo) = "(OK)"
			QfyStatusLongText(EvtNo) = "&nbsp; (OK) - Nationals Placement"
			QfyStatusTitle(EvtNo)="Current LEVEL = LEVEL - ("&RIGHT(srating,1)&") &#13; Qualified by Nationals Placement"

	' --- Checks for Masters rating plus CURRENT Regionals Placement  ---
	ELSEIF sRegl_Plc <= 5 AND RegStatusText <> "PEND" THEN
			QfyStatusColor(EvtNo) = TextColor2		
			QfyStatusText(EvtNo) = "(OK)"
			QfyStatusLongText(EvtNo) = "&nbsp; (OK) - Regionals Placement"
			QfyStatusTitle(EvtNo)="Current Level = LEVEL ("&RIGHT(srating,1)&") &#13; Qualified by Regionals Placement"

	ELSE
			QfyStatusColor(EvtNo) = TextColor3
			QfyStatusText(EvtNo) = RIGHT(sRating,1)
			QfyStatusLongText(EvtNo) = "&nbsp; Not Qualified = LEVEL - "&RIGHT(srating,1)
			QfyStatusTitle(EvtNo)="Current Level = LEVEL - "&RIGHT(srating,1)
			ScratchCount = ScratchCount + 1
	END IF

ELSEIF LEFT(sRating,1)="X" THEN
	QfyStatusColor(EvtNo) = TextColor3
	QfyStatusText(EvtNo) = RIGHT(sRating,1)
	QfyStatusLongText(EvtNo) = "&nbsp; Not Qualified = LEVEL - "&RIGHT(sRating,1)
	QfyStatusTitle(EvtNo)="Not Currently Qualified - LEVEL - "&RIGHT(srating,1)
	ScratchCount = ScratchCount + 1
ELSE
	QfyStatusColor(EvtNo) = TextColor3
	QfyStatusText(EvtNo) = RIGHT(sRating,1)
	QfyStatusLongText(EvtNo) = "&nbsp; Not Qualified = LEVEL - "&RIGHT(srating,1)
	QfyStatusTitle(EvtNo)="Not Currently Qualified - LEVEL - "&RIGHT(srating,1)
	ScratchCount = ScratchCount + 1
END IF




END SUB




' -----------------------------------------------------------------------------------------------------------------------
   SUB CheckNationalQualifications (sNatl_Plc, sRegl_Plc, sRating, sRank, sQfyOverride, sTourID, sDiv, sEvent, sMemberID)
' -----------------------------------------------------------------------------------------------------------------------

' ------------------------------------------------------------------------------------
' --- Rating, Masters as 3rd, Regional & National Placement Tested 	--------------
' --- Requires Regional Participation test	 			--------------
' ------------------------------------------------------------------------------------

QualStatusEvent1=""
QualStatusEvent2=""
QualStatusEvent3=""
QualStatusEvent4=""

QfyStatusColor = TextColor1
QfyStatusText = ""
QfyStatusTitle=""

' --- Removes TIES from National placement to allow testing
TPosNatl = InStr(lcase(sNatl_Plc), "t") - 1
Natl_Plc = sNatl_Plc
IF TRIM(Natl_Plc)<>"" THEN
	IF TPosNatl > 0 THEN
		Natl_Plc = (Left(sNatl_Plc, TPosNatl))
	END IF
ELSE
	Natl_Plc=999
END IF

' --- Removes TIES from Regional placement to allow testing
TPosRegl = InStr(lcase(sRegl_Plc), "t") - 1
Regl_Plc = sRegl_Plc
IF TRIM(Regl_Plc)<>"" THEN
	IF TPosRegl > 0 THEN
		Regl_Plc = Left(sRegl_Plc, TPosRegl)
	END IF
ELSE 
	Regl_Plc=999
END IF


IF LEFT(rs("Rating"),1)="O" THEN
	QfyStatusColor = TextColor2
	QfyStatusText = "OK"
	QfyStatusTitle="Current Rating = Open ("&srating&")"

ELSEIF LEFT(sRating,1)="E" AND RIGHT(sRating,1)<>"1" THEN
	QfyStatusColor = TextColor2
	QfyStatusText = "OK"
	QfyStatusTitle="Current Rating = EP ("&srating&")"

ELSEIF TRIM(sQfyOverride) <> "" THEN
	QfyStatusColor = TextColor2
	QfyStatusText = "(OV)"
       	QfyStatusTitle="Administrative Override of Qualifications &#13; Override Code: "&sQfyOverride

ELSEIF cdbl(sRank) = 0 OR TRIM(sRating) = "0" OR TRIM(sRating) = "" THEN
	QfyStatusColor = TextColor3
	QfyStatusText = "***"
        QfyStatusTitle="No Current Rating"
	ScratchCount = ScratchCount + 1

ELSEIF LEFT(sRating,1)="M" OR (LEFT(sRating,1)="E" AND RIGHT(sRating,1)="1") THEN

	' --- IF the current event is a Masters check to see if this event qualifies by 3rd event
	Set rsQualify=Server.CreateObject("ADODB.recordset")
	sSQL = "SELECT RGT.MemberID, RGT.Event, RGT.Div, RGT.AWSA_Rat"
	sSQL = sSQL + " FROM "&RegTemporary&" AS RGT"
	sSQL = sSQL + " WHERE RGT.MemberID = '"&sMemberID&"' AND RGT.Div='"&sDiv&"' AND LEFT(RGT.TourID,6)='"&LEFT(sTourID,6)&"'"
	rsQualify.open sSQL, sConnectionToTRATable, 3, 1

	DO WHILE NOT rsQualify.EOF 	

		' --- Checks for EP or Open Rating ---
		IF (UCASE(LEFT(rsQualify("AWSA_Rat"),1)) = "E" AND RIGHT(rsQualify("AWSA_Rat"),1)<>"1") OR UCASE(LEFT(rsQualify("AWSA_Rat"),1)) = "O" THEN
			ThisEventStat = "OK"
		ELSE
			ThisEventStat = "Rating"
		END IF

		SELECT CASE rsQualify("Event")
			  CASE "S"
				QualEvent1 = rsQualify("AWSA_Rat")
				QualStatEvent1 = ThisEventStat
			  CASE "T"
				QualEvent2 = rsQualify("AWSA_Rat")
				QualStatEvent2 = ThisEventStat
			  CASE "J"
				QualEvent3 = rsQualify("AWSA_Rat")
				QualStatEvent3 = ThisEventStat
		END SELECT
		
		rsQualify.movenext

	LOOP

	' --- Examines for Open/EP in two events and Masters in the 3rd event ---
	IF (sEvent="S" AND QualStatEvent2 = "OK" AND QualStatEvent3 = "OK") OR (sEvent="T" AND QualStatEvent1 = "OK" AND QualStatEvent3 = "OK") OR (sEvent="J" AND QualStatEvent1 = "OK" AND QualStatEvent2 = "OK") THEN 
		QfyStatusColor = TextColor2		
		QfyStatusText = "(OK)"
		QfyStatusTitle="Current Rating = Masters or Single EP - ("&srating&") &#13; Qualified by 3rd Event Masters or Single EP in Class C"

	' --- Checks for Masters rating plus PREVIOUS Nationals Placement  ---
	ELSEIF Natl_Plc <= 5 THEN
		QfyStatusColor = TextColor2		
		QfyStatusText = "(OK)"
		QfyStatusTitle="Current Rating = Masters or Single EP - ("&srating&") &#13; Qualified by Nationals Placement & Masters Rating"

	' --- Checks for Masters rating plus CURRENT Regionals Placement  ---
	ELSEIF Regl_Plc <= 5 AND RegStatusText <> "PEND" THEN
		QfyStatusColor = TextColor2		
		QfyStatusText = "(OK)"
		QfyStatusTitle="Current Rating = Masters ("&srating&") &#13; Qualified by Regionals Placement & Masters Rating"

	ELSE
		QfyStatusColor = TextColor3
		QfyStatusText = rs("rating")
		QfyStatusTitle="Current Rating = Masters ("&srating&")"
		ScratchCount = ScratchCount + 1
	END IF

ELSEIF LEFT(Rating,1)="X" THEN
	QfyStatusColor = TextColor3
	QfyStatusText = srating
	QfyStatusTitle="Current Rating = Expert ("&srating&")"
	ScratchCount = ScratchCount + 1
ELSE
	QfyStatusColor = TextColor3
	QfyStatusText = srating
	QfyStatusTitle="Current Rating = ??? ("&srating&")"
	ScratchCount = ScratchCount + 1
END IF

IF sMemberID="700000577" THEN
	'markdebug("scratchcount in qualifications = "&ScratchCount)
END IF

END SUB





' --------------------------------
   SUB VerifyWaiver (sWaiverCode)
' --------------------------------

WaiverTitle = ""
WaiverColor = "" 
WaiverText = ""

IF TRIM(sWaiverCode) = "" OR TRIM(sWaiverCode) = "None" THEN
	WaiverTitle = "No Waiver On File"
	WaiverText = "***"
	WaiverColor = TextColor3
	ScratchCount = ScratchCount + 1
ELSE		
	WaiverTitle = "Signed Waiver &#13; Waiver Code/Type = "&sWaiverCode
	WaiverColor = TextColor2 
	WaiverText = "OK"
END IF	


END SUB





' ----------------------------------------------------------
   SUB VerifyTrickList (sForm2Name, sEvent, EventSelected)
' ----------------------------------------------------------

TrickTitle = ""
TrickColor = "" 
TrickText = ""

' ------------------------------------------------------------------------------------------
' Trick Form Requirement - NEED TO ADD logic for looking at TouRGTTable to see if required.
' ------------------------------------------------------------------------------------------
IF TRIM(lCase(sForm2Name))="list" AND TRIM(UCase(sEvent)) = "T" THEN
	TrickTitle = "Trick List required before event."
	TrickColor = TextColor3 
	TrickText = "***"
	ScratchCount = ScratchCount + 1			
ELSEIF TRIM(lCase(sForm2Name))="list" AND TRIM(sEvent) <> "T"  AND EventSelected = "ALL" THEN		
	TrickTitle = "Not Trick Event"
	TrickText = " "
	TrickColor = TextColor2
ELSEIF TRIM(lCase(sForm2Name))="list" THEN 
	TrickTitle = "No Trick List Needed"
	TrickText = "OK"
	TrickColor = TextColor2
END IF


END SUB




' -----------------------------------------------------------------------
   SUB VerifyPersonalBio (sForm1Name, sBioMemberID, sMemberID, sTourID)
' -----------------------------------------------------------------------

BioText = ""
BioColor = ""
BioTitle = ""
BioLink = ""

IF Trim(sBioMemberID) <> "" THEN 
	BioText = "OK"
	BioColor = TextColor2
	BioTitle = ""
	BioLink = "/rankings/view-registration.asp?WhatReport=viewbio&sMemberID="&sMemberID&"&sTourID="&sTourID

ELSEIF Trim(sBioMemberID) = "" AND sBio_Reqd=false THEN 
	BioText = "--"
	BioColor = TextColor2
	BioTitle = "Bio Not Required"
	BioLink = ""
ELSE
	BioText = "***"
	BioColor = TextColor3
	BioTitle = "No Personal Bio On File"
	BioLink = ""
END IF 



END SUB








%>




