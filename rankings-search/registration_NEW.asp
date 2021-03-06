<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->
<!--#include virtual="/rankings/tools_registration.asp"-->
<!--#include virtual="/rankings/tools_TourDefine.asp"-->
<!--#include virtual="/rankings/tools_Definitions.asp"-->
<!--#include virtual="/rankings/qualifications.asp"-->
<!--#include virtual="/rankings/RegFormDisplay_New.asp"-->
<!--#include virtual="/rankings/Register_Survey.asp"-->
<%

' -------------------------------------------
' --- GENERAL FRAMEWORK: 
' --- ???
' --- VER 9-5-2009 ---
' -------------------------------------------


Dim RegFileName, CardFileName, DisplayFileName
RegFileName="Registration_New.asp"
CardFileName="CCReg2012.asp"



' --- Allows use across Subroutines ---
Dim currentPage, sSendingPage
'Dim TestValidAdminCode

Dim TestMode
TestMode="no"
'TestMode="yes"


' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' --- CONCEPTS ---
' --- Variables not defined herein are typically defined globally in SettingsHQ.asp or in Tools_Registration.asp 
' --- Variables that are event specific are dimensioned as arrays() and defined based on EvtNo
' --- 
' --- DEFINING VARIABLES
' --- DefineMemberVariables - Defines the member related variables.  Note requires Tournament to determine age
' --- ZeroOutVariables - Initilizes all variables
' --- InitializeFromTable - Sets variables based on the value from TEMP or PERMANENT table
' --- ReadFromRegisterEvents - Sets event related variables based on values from TEMP or PERMANENT table
' --- 	   NOTE:  Sequence events are loaded must match order events are set up in RegistrationEventsOffered in tools_include.asp 
' --- ReadEntryFormValues - Reads FORM variables for General and Event level varaibles
' --- CalculateEntryFees - Determines appropriate entry fees depending on settings in SWIFT 
' --- RecalcFormValues - Resets the values for all form variables depending on selections or file data  
' --- GetFinancialTotals - Finds the totals for the entire form 
' ---	   NOTE:  Some values are being calculated in RegFormDisplay.asp and should be moved to GetFinancialTotals or other module	
' ---
' ---
' ---

' --- NEW or Upgrading Members
' --- The web Member table is updating nightly.  Changes to the Membership status are tested against the MemberHistory table
' --- 	on the HQ server.
' ---
' --- I/O ROUTINES  
' --- CopyDataToTables - Copies the defined GENERAL variables to either the TEMP or permanent tables
' --- CopyToEventDetail - Copies the EVENT level variables to either the TEMP or permanent tables
' --- UpdateTransTable - Updates the table which captures the line item detail of the fees
' --- UpdatePaymentTransaction - Updates the General Registration table to indicate payment was made





' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' ++++ Check why RunByWhat is used (vs sRunByWhat) - Maybe in Search-MemberHQ.asp or View-TournamentsHQ.asp +++
Dim RunByWhat

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' --- Read from MemberTableName ---
Dim sMemberID, sLastName, sFirstName, sFullName, sMembSex, sMembCity, sMembState, sMembAge, sMembPhone, sMembTypeID, sCanSkiTour, sMembTypeCode
Dim sMembEmail, sEffectiveTo, sMembBirth, sCostToUpgrade, sTypeDesc


' --- Variables for establishing the base or other codes for classes of the tournament ---
Dim sTotalFormFees, sTotalEntry, sEntryFee, sLateFeeTot, sAWSEFDonation, sOffDiscAmt, sJrDiscAmt, sSrDiscAmt, sClubDiscAmt, TotEvents

' --- Variables used to determin values from RegTransactions table ---
Dim sEntryFeeTrans, sLateFeeTotTrans, sBanquetTotTrans, sAWSEFDonationTrans, sOffDiscAmtTrans, sSrDiscAmtTrans, sJrDiscAmtTrans, sClubDiscAmtTrans


' --- Defining current state of form ---
Dim sBanquetQty, sBanquetTot, sAWSEFCheck, sOfficial, sClubMemb, sClubCode 
Dim sLateDays, sMembRegDate, BoatStatus, sRampHeight, RampStatus, sEntryType, EntryTypeStatus
Dim sMembOverride, sRegionalOverride, sMoneyOverride, sBioDone, TeamSelected


' --- Headings on RegFormDisplay.asp based on the specifics of the tournament ---
Dim sClassCols, sClassWidth, sGrassOffered
Dim sFormError


' --- Internal program control variables ---
Dim WhichTable, DetailTable
Dim nav

' --- Settings in Registration Form
Dim DisplayVars


' --- Button status controls
Dim MainStatusValue, BioButtonStatus, FormStatus, MainStatus, AllObjectStatus, EditButtonStatus, PreviousButtonStatus

' --- Used in control and processing of payments ---
Dim sPaymentResult, sPayType, sPPResult, sPayStatus
Dim sOrderNo, sPayAmount, MaxOrder
Dim sErrorNo


' --- Variables relating to the Waiver/Release ---
Dim ReleaseVersion, sRelease, sReleaseType, sWaiverCode, sSignWaiver, sTableEmail, sWaiverSubtitle
Dim sSpecialWaiverCode, sSpecialWaiverHeadline, sSpecialReleaseBannerText

' --- Notes specific to display and receipts ---
Dim ReceiptNote1, ReceiptNote2, ReceiptNote3, ReceiptNote4, ReceiptNote5


' --- Read these from table in Cont_Disp.asp
Dim sEntryEmail, sWaiverEmail, sPasswordEmail, sSkipWaiver, sForceWaiver
Dim sEntryEmailAdm, sWaiverEmailAdm, sPasswordEmailAdm, sSkipWaiverAdm, sForceWaiverAdm
Dim sEntryEmailHQ, sWaiverEmailHQ, sPasswordEmailHQ, sSkipWaiverHQ, sForceWaiverHQ
Dim sDispDebugButtons, sDispDebugButtonsAdm, sDispDebugButtonsHQ

Dim sSpecialWaiverEmailMC, sWaiverEmailMC, LOCSpecialWaiverEmail


' --- Controls whether Mark Crone is copied on waiver emails ---
sSpecialWaiverEmailMC=false
sWaiverEmailMC=false
LOCSpecialWaiverEmail="2011natent@harvat.com"





adminmenulevel = Session("adminmenulevel")
IF TRIM(adminmenulevel) = "" THEN adminmenulevel = "1"


' ----------------------------------------------------------------------------------------------------------------------
' ---- These display in various text boxes on Entry Form and Receipt - Some may need to be moved to the Tour Setup table
' ----------------------------------------------------------------------------------------------------------------------

ReceiptNote1 = "Registration check-in and familiarization with the ski site is recommended."
ReceiptNote2 = "Registration typically closes 20 minutes prior to each event."
ReceiptNote3 = "A paper copy of a signed Waiver and Release is required for payment on site or by mail." 
ReceiptNote4 = "Speed Control in approved towboats. Distance by Video Jump. Trick lists NOT required." 
ReceiptNote5 = "Password protected Bio saves you time and reduces confusion for announcers." 


' -----------------------------------
' ------- Current waiver codes  -----
' -----------------------------------

' --- Changed t0 2010 on 3-5-2011 ---
adult_waiver = "adlt2010"
minor_waiver = "min_2010"

'adult_waiver = "adlt2007"
'minor_waiver = "min_2007"




' -------------------------------------------------------------------------------------------
' ---------------  Values associated with branching stuff  ----------------------------------
' -------------------------------------------------------------------------------------------

' --- Initializes values when sent from menu ---
sProcess=TRIM(Request("process"))
IF sProcess="register" THEN 
		Session("sTourID")=""
		Session("sMemberID")=""
END IF


' -------------------------------------------------------------------------------------------------
' ---- If Member, Tournament, and main branch variables are all null the send user to Welcome screen
' -------------------------------------------------------------------------------------------------

' --- sRunByWhat is the MAIN branching variable  ----
sRunByWhat=TRIM(Request("sRunByWhat"))


IF TRIM(Request("sTourID"))<>"" THEN
		sTourID=Request("sTourID")
		Session("sTourID")=sTourID
ELSE
		sTourID = Session("sTourID")
END IF

IF TRIM(Request("sMemberID"))<>"" THEN
		sMemberID=Request("sMemberID")
		Session("sMemberID")=sMemberID
ELSE
		sMemberID = Session("sMemberID")
END IF



' ----------------------------------
' --- TEST DATA --------------------
' ----------------------------------


'Session("sTourID")="11E061"
'sTourID="11E061"

' --- Mark Crone
'Session("sMemberID")="000001151"
'sMemberID="000001151"


' ---------------------------------------------------------------------------------------
' --- Test to determine whether to display the welcome screen based on Session values ---
' ---------------------------------------------------------------------------------------
IF TRIM(Session("SeenWelcome"))="" THEN
		Session("SeenWelcome")=true
		IF Session("sMemberID")="" OR Session("sTourID")="" OR sprocess="tlink" THEN 
				IF adminmenulevel >= 30 THEN
						sRunByWhat="Edit"
				ELSE
						sRunByWhat="Welcome"
				END IF
		END IF
END IF




' -------------------------------------------------------------------------------------------------
' ----  FormStatus defines whether the registration form is in modify or confirm mode  ----
' -------------------------------------------------------------------------------------------------

nav=TRIM(Request("nav"))
IF nav="" THEN nav=1

FormStatus = Request("FormStatus")
MainStatus = Request("MainStatus")
Previous=Request("Previous")
Edit=Request("Edit")



' -------------------------------------------------------------------------------------------------
' ----  ChargeStatus is returned from CC_Process2.asp to specify successful transaction  ----
' -------------------------------------------------------------------------------------------------
ChargeStatus = Request("ChargeStatus")    ' Value is null until returned from CC_Process2.asp
IF Chargestatus = "success" THEN sRunByWhat = "Done"

sOrderNo=Request("sOrderNo")
sPayType=Request("sPayType")


' ---------------------------------------------------------
' --- Reads the settings for email and display controls ---
' ---------------------------------------------------------
ReadContDispTableValues


' ---------------------------------------------
' ------    Redisplay the PAGE Footer  --------
' ---------------------------------------------

IF sRunByWhat <> "Print" THEN 
	WriteIndexPageHeader
END IF


' --------------------------------
' --- Displays off line notice ---
' --------------------------------
'IF Session("AdminMenuLevel")<50 THEN sRunByWhat="NotActive"
tp=1
IF tp=2 AND sMemberID="000001151" THEN
			Response.write("<br>Line 300 - This Is Mark")
			Response.write("<br>sRunByWhat = "&sRunByWhat)
END IF


' -------------------------------------------------------------
' --- Top of SELECT statement that controls program routing ---
' -------------------------------------------------------------

SELECT CASE sRunByWhat 



 ' ---------------------
 ' ---------------------
  CASE "NotActive"
 ' ---------------------
 ' --------------------- 

	%>
	<TABLE class="droptable" ALIGN="center" width=75% >
		<tr>
	  	<td align=center>
		     <font size="4" ><b>USA Water Ski On-Line Registration</b></font>
	  	   <br>
	  	</td>
  	</tr>
		<tr>
	  	<td align=center>
	    	<font size="2" ><br>The member entry module of Online Registration is being upgraded to accommodate recent upgrades to the Online Sanctioning system.  <br><br> The upgraded version of OLR is expected to be available by February 1, 2010.<br><br> Pardon the inconvience and thanks for your understanding.<br></font>
	     	<br>
	  	</td>
     </tr>
	</TABLE>
	<%


  

 ' ---------------------
 ' ---------------------
  CASE "VerifyUpgrade"
 ' ---------------------
 ' ---------------------
 
	' --- Case established from RegFormDisplay.asp action on Upgrade button ---

	'--- In SUB tools_registration.asp ---
	DefineTourVariables_New

	' --- Determines EVENTS and EVENT NAMES for this tourament - in tools_include.asp ---
	RegistrationEventsOffered (sTSptsGrpID)

	DefineMemberVariables
	
	IF sMemberID="000001151" THEN
			Response.write("<br>This Is Mark")
	END IF		
  ' --- Changed 2-12-2013 - History table no longer necessary - a simple refresh will work --- 
	' VerifyMemberHistoryUpdate
	response.redirect("/rankings/"&RegFileName&"?nav=2")



 ' ----------------
   CASE "Welcome"
 ' ----------------

	WelcomeButton="Select Tournament"
	IF TRIM(Session("sTourID"))<>"" THEN
			WelcomeButton="Select Member to Register"
	ELSEIF TRIM(Session("sMemberID"))<>"" THEN
			WelcomeButton="Select Tournament"
	END IF
		
	
	%><br>
	<TABLE class="innertable" ALIGN="center" width=95% >
	<tr>
	  <th align=center>
	     <font size="4" color="#FFFFFF"><b>Welcome to</b></font><br>
	     <font size="4" color="#FFFFFF"><b>USA Water Ski On-Line Registration</b></font>
	     <br>
	  </th>
        </tr>

	<tr>
	<td align="center" style=" white-space:wrap;">
	<br>
	     	<font size="<%=fontsize2%>">
		<p>Once you have selected a tournament, you'll locate your name and complete the online registration Entry Form.  
		<br>
		 Online registration is available only for those tournaments which have activated the online registration option.</font>   
		<br><br>
		<font size="<%=fontsize2%>">
		<b>Members must have their password</b> to access Online Registration for tounament entry.
		<br>
		 If you do not remember your password, you can go to <b>Member Section</b> of USA Waterski website to recover password information. 
		<br>
		 If you are not currently a member, require renewal, or if a competitor-status upgrade is required,
		<br>
		 please complete the <font size=<% =fontsize2 %> color="blue"><A href="https://www.usawaterski.org/renew/"><font size=<% =fontsize3 %> face=<% =font1 %> color="blue">USA Water Ski Membership Application</font></A></font> before continuing.  
		<br><br>
		 <b>New Members</b> should retain Member # to access Find Member tool for Online Registration.</p></font> 
		<font size="<%=fontsize2%>" color="red"><b>You must have an email address to use Online Registration.</b></p></font>
		<font size="<%=fontsize2%>"><b>Payments for entry fees</b> are processed through each tournament organizer's PayPal account.  
		<br>
		  In addition to the USA Water Ski registration acknowledgement that is emailed to you at the end of this transaction,
		<br>
		 you will also receive a separate PayPal receipt. <b>Please retain the PayPal receipt as this is the only proof of payment
		<br>
		 you will receive.</b>  Refunds, credits or other issues relating to entry fees and payments should be directed to the
		<br> tournament organizer, or to the contact information on your PayPal receipt.</font> 
		<br><br>
		<font size="<%=fontsize2%>"><b>IMPORTANT</b> <br> 1) You must be 18 to use the on line registration system.
		<br>
		<font size="<%=fontsize2%>" color="red"><b> 2) Please use only the program's navigation buttons 
		<br>Do not use your browser's back button!</b></font>

		<br><br>
		<font size="<%=fontsize2%>">For general questions, contact: USA Water Ski Competition Dept at 800-533-2972.</font>
	     
	    <form align="center" action="/rankings/<%=RegFileName%>?sRunByWhat=Edit" method="post">
	    	<input type="submit" value="<%=WelcomeButton%>">
	    </form>
	  </td>
	</tr>
	</TABLE>

	<% 


 ' ----------------
 ' ----------------
   CASE "Tour"
 ' ----------------
 ' ----------------

	sUserSptsGrpID="AWS"

	SELECT CASE sUserSptsGrpID
	   CASE "AWS"
		sEventString = "sl=on&tr=on&ju=on"
	   CASE "USW"
		sEventString = "wb=on&ws=on&wsu=on"
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


	' --- Resets the important session variables for the member that have also to do with the TourID
	ResetMembTourSessionVar


	' ---  Branches to Identify a new Session(sTourID) ---
	Session("sSendingPage") = "/rankings/"&RegFileName&"?rid="&rid
	Session("sTourID") = ""
	Session("sMemberID")=""

	response.redirect("/rankings/view-tournamentsHQ.asp?process=register&sSendingPage=NEW&"&sEventString&"&sTourSportGroup="&sUserSptsGrpID&"&sTourRange=1&sTest=on")



 ' ----------------
 ' ----------------
   CASE "NewMember"
 ' ----------------
 ' ----------------

	Session("sMemberID")=""
	Session("sExpirationStatusText")=""
	Session("sExpirationStatusColor")=""
	' Session("sOrderNo")

	' --- Resets the important session variables for the member
	ResetMembTourSessionVar



	Session("sSendingPage")="/rankings/"&RegFileName
	Response.Redirect("/rankings/search-memberHQ.asp?rid="&rid&"&formstatus=search")


 ' ----------------
   CASE "Member"
 ' ----------------

	'DefineTourVariables_New
	'IF LEFT(Session("sTourID"),6)="15W091" THEN response.write("<br>Session(AdminCode) ="&Session("AdminCode"))

	sMemberRadio=TRIM(Request("MemberRadio"))

	' --------- If Session(sMemberID) is NOT null then the MemberID has been selected  ---------

	IF adminmenulevel >=19 THEN
			Session("sSendingPage")="/rankings/"&RegFileName
			Response.Redirect("/rankings/search-memberHQ.asp?rid="&rid&"&formstatus=search")
	END IF


	IF Session("sMemberID") = "" THEN 

		%>
		<br><br>

		<TABLE class="innertable" ALIGN="CENTER" width=50%>
		<tr>
		  <th align=center>
		    <font face=<% =font1 %> size="4" color="#FFFFFF"><b>Membership Status</b></font>
		  </th>
		</tr>
		<tr>
		  <td align=center>
		    <br>
		    <form action = "/rankings/<%=RegFileName%>?rid=<%=rid%>&sRunByWhat=Member" method="post">
		      <center><FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR="#0000CD">Member:</font><input type="radio" name="MemberRadio" <%IF sMemberRadio="Member" THEN Response.write("checked")%> value="Member">
		      <FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR="#0000CD">Non_Member:</font><input type="radio" name="MemberRadio" <%IF sMemberRadio="NonMember" THEN Response.write("checked")%> value="NonMember">
		      <br>
		      <br><% 

		IF sMemberRadio="Member" THEN
				Session("sSendingPage")="/rankings/"&RegFileName
				Response.Redirect("/rankings/search-memberHQ.asp?rid="&rid&"&formstatus=search")
		END IF

		IF sMemberRadio="NonMember" THEN
			%>
			    <br>
			    <font size=<% =fontsize3 %> ><b>Notice - Non Members Must Join USA Water Ski Before Proceeding</b></font>
			    <br><br>
			    <font size=<% =fontsize3 %>><a href="https://www.usawaterski.org/renew/" title="USA Water Ski online membership application">Join or Renew</a>   
	     		    <br><br>
				Membership data is immediately available once you have completed your membership application.
				<br>Age and gender data is used to validate division and other parameters.
			    <br><br>
			    </font>
			<%
	  		
		END IF %>

		<input type="submit" value="Select Member Status ">
		</form>
  		</td> 	
		</tr>
		</table><%

	ELSE		
		' ----------------------------------------------------------------------------------------		
		' ------ Session(sMemberID) has been established  ----------------------------------------
		' ----------------------------------------------------------------------------------------
		' ------ Resets Session variables to restart transaction tracking ------------------------
		' ----------------------------------------------------------------------------------------


		Session("Know_Orig_Trans") = ""

		' ----------------------------------------------------------------------------------------
		' ------ Deletes any temporary registration data for from the RegTempTable  --------------
		' ----------------------------------------------------------------------------------------
		OpenCon
		sSQL = "DELETE FROM "&RegTempTableName
		sSQL = sSQL + " WHERE Left(TourID,6) = '" & SQLClean(left(sTourID,6)) & "' AND MemberID = '"&sMemberID&"'"
		con.execute(sSQL)
		closecon

		
		' -----   Branches to default of main SELECT = ELSE  -----
  		response.redirect("/rankings/"&RegFileName&"?sRunByWhat=Edit")   	

	END IF


' ----------------	
' ----------------
    CASE "Print"
' ----------------	
' ----------------	

	PrintReceipt



' ---------------------------
' ---------------------------
    CASE "ReturnToMainMenu"
' ---------------------------
' ---------------------------

	Session("sMemberID") = ""
	Session("sTourID") = ""
	Session("Know_Orig_Trans") = ""
	Session.Abandon

	response.redirect("/rankings/"&RegFileName&"?process=register")	


 ' -------------------------
 ' --------------------------
   CASE "DeletePayments"
 ' -------------------------
 ' -------------------------
 
	OpenCon
	sSQL = "DELETE FROM "&RegPaymentTableName&" WHERE MemberID='000001151'"
	con.execute(sSQL)

	sSQL = "DELETE FROM "&RegTransTableName&" WHERE MemberID='000001151' AND OrderNo >=2000"
	con.execute(sSQL)
	CloseCon

	response.redirect("/rankings/"&RegFileName&"?nav=1")


 ' -------------------------
 ' --------------------------
   CASE "missingageorgender"
 ' -------------------------
 ' -------------------------
 
  ' --- Is there really any such thing?  ----

		%>
		<TABLE BORDER="4" ALIGN="CENTER" CELLPADDING="0" CELLSPACING="0" BGCOLOR="<%=TableColor1%>" width=60% >
		  <TR>
		      <TD BGCOLOR="red"><center><font face=<% =font1 %> size="4"><b>Important Message</b></font><br></TD>
		  </TR>  
		  <TR>
		     <TD>

			<TABLE ALIGN="Center" BORDER="0" BGCOLOR="<% =tablecolor1 %>" CELLPADDING="6" CELLSPACING="3" width=100% BGCOLOR="#FFFFFF">
			<tr>
			   <td colspan="2" align="center">
			    <br> 
				<font face="<% =font1 %>" size="3"><b>Your membership record is missing critical information</b></font>
			    <br><br> 
				<font face="<% =font1 %>" size="2">Until this information is corrected, you may not proceed with on line registration.</font><br><% 

				IF IsNull(sMembAge)=true THEN  %>
					<br>
					<font face="<% =font1 %>" size="3" color="<% =textcolor2 %>"><b>Member Age is Missing.</b></font><%
				END IF

				IF TRIM(sMembSex) = "" THEN  %>
					<br> 
					<font face="<% =font1 %>" size="3" color="<% =textcolor2 %>"><b>Member Gender is Missing.</b></font><%
				END IF %>
				<br>
				<font face="<% =font1 %>" size="1"><p>Please contact USA Water Ski - Membership Department at 800-533-2972.</font>
			  </td>
			</tr>
			<tr>
			   <td align="center">
				<form action="/rankings/<%=RegFileName%>?sRunByWhat=missingageorgender&FormStatus=OK" method="post">
				  <input type="submit" value=" Continue ">
				</form>
			   </td>	
			</tr>
			</TABLE>

		    </TD>
		  </TR>
			
		</TABLE> <%



' --------------------
' --------------------
  CASE "tournotsetup"
' --------------------
' --------------------


	DisplayNotSetupNotice


 ' -------------------------------------------------------------------------------------------------------------------
 ' -------------------------------------------------------------------------------------------------------------------
   CASE ELSE				' This is a catch all - since there is no CASE EDIT this is where it goes 
 ' -------------------------------------------------------------------------------------------------------------------
 ' -------------------------------------------------------------------------------------------------------------------

	'--- Determines if the TourID Session variable has been set
	IF Session("sTourID")="" THEN response.redirect("/rankings/"&RegFileName&"?sRunByWhat=Tour")



	'--- Sets all tournament and registration related variables in SUB in tools_registration.asp ---
	DefineTourVariables_New




	' --- Verifies that the tournament is set up in SWIFT ---
	IF sPayPalOK=0 OR TRIM(sPayPalAct)="" OR sUseOLReg=0 OR sOLR_PD=0 THEN response.redirect("/rankings/"&RegFileName&"?sRunByWhat=tournotsetup")



	' --- Determines EVENTS and EVENT NAMES for this tourament - in tools_include.asp ---
	RegistrationEventsOffered (sTSptsGrpID)


	' --- If Sesssion variable has not previously been set then send to get the Member info ---
	IF Session("sMemberID")="" THEN response.redirect("/rankings/"&RegFileName&"?sRunByWhat=Member")
	DefineMemberVariables


	' --- If system has not read variables for 1st time then test whether MemberID is previously entered 
	' ---    Checks for AdminCode or AdminLevel
	' ---    Then, if NOT entered do not allow entry if sOLRDisplayStatus is false 

'response.write("<br>Session(Know_Orig_Trans) is null = ")
'response.write(Session("Know_Orig_Trans") = "")
'response.write("<br>NOT(IsMemberEntered) = "&NOT(IsMemberEntered))
'response.write("<br>NOT(sOLRDisplayStatus) = "&NOT(sOLRDisplayStatus))
'response.write("<br>Session(AdminMenuLevel)<50 = ")
'response.write(Session("AdminMenuLevel")<50)
'response.write("<br>NOT(TestValidAdminCode) = ")
'response.write(NOT(TestValidAdminCode))


	IF Session("Know_Orig_Trans")="" AND NOT(IsMemberPaid) AND NOT(sOLRDisplayStatus) AND NOT(TestValidAdminCode) AND Session("AdminMenuLevel")<50 THEN 
				DisplayOLRDisabledNotice
	END IF



	' --- Determines the Fees recorded and Sets value of sTotalPreviousPayments ---
	DetermineTotalFeesActuallyPaid


	IF nav="" THEN nav=1
	SetNavigationVariables


	' ++++++++++++++++++++++
	' --- TESTING SCRIPT ---
	' ++++++++++++++++++++++
	IF sMemberID="000001151" AND TestMode="yes" THEN
			response.write("<br>Before all NAV Conditions")
			response.write("<br>MainStatus = "&MainStatus)
			response.write("<br>nav = "&nav&"<br>")
	END IF

	
	' ----------------------
	' --- TOURNAMENT TAB ---
	' ----------------------
		
	IF nav=1 THEN
			' --- Checks Age or Gender data and posts dialog box if either is missing ---
			IF IsNull(sMembAge)=true OR TRIM(sMembSex) = "" THEN
					response.redirect("/rankings/"&RegFileName&"?sRunByWhat=missingageorgender")
			END IF

			' --- Know_Orig_Trans indicates values previously read from RegGenTableName ---
			Session("Know_Orig_Trans") = "REGGEN"
			MainStatus="Verify"

			WhichTable=RegGenTableName
			InitializeFromTable
		

			IF MainStatus= "Verify" THEN
					WhichTable=RegTempTableName
					CopyDataToTables
			END IF
	END IF


	' --------------------------------------
	' --- ENTRY FORM and ENTRY FEES TABS ---
	' --------------------------------------
	IF (nav=3 OR nav=4) AND MainStatus<>"Verify" THEN
			' --- When note equal Verify then load all the variables from the Temporary table
			WhichTable=RegTempTableName
			InitializeFromTable
	END IF


	IF (nav=3 OR nav=4) AND MainStatus="Verify" THEN
			' --- When in VERIFY mode read the variables from the form ---
			ReadEntryFormValues

			IF nav=3 THEN
					ValidateFormValues
			ELSEIF nav=4 THEN
					ValidateFormDateEntered		
			END IF

			p=2
			IF p=1 AND sMemberID="000001151" THEN
					Response.write("<br>nav= "&nav)
					Response.write("<br>MainStatus= "&MainStatus)
					Response.write("<br>sFormError= "&sFormError)
					'response.end
			END IF

			IF MainStatus= "Verify" AND TRIM(sFormError)="" THEN
					WhichTable=RegTempTableName
					CopyDataToTables
			END IF
	END IF



	' ----------------------------------------------------------------------------------------	
	' --- If the release has already been signed and Force is not on - OR SkipWaiver is on ---
	' ----------------------------------------------------------------------------------------
	IF nav=5 AND (  (TRIM(Session("sRelease"))<>"" AND TRIM(Session("sRelease"))<>"None" AND sForceWaiver<>true AND TestValidAdminCode<>true) OR sSkipWaiver=true ) THEN
			nav=6
	END IF
		

	' -------------------
	' --- PAYMENT TAB ---
	' -------------------
	IF nav=6 THEN

			' ------  Read from tables not form  --------
			WhichTable=RegTempTableName
			InitializeFromTable

				' --- TESTING SCRIPT ---
				Y=98
				IF MarkTester=true AND Y=99 THEN
						response.write("<br>Pos 1")
						response.write("<br>sTotalPreviousPayments = "&sTotalPreviousPayments)
						response.write("<br>sTotalFormFees = "&sTotalFormFees)
						response.write("<br>Who Paid = "&Session("sWhichFamilyMemberPaid"))
						'response.end
				END IF

			' --- Check FormTotal against previous payments ---
			IF cdbl(sTotalFormFees) > sTotalPreviousPayments THEN
					sPayStatus="O"
			ELSEIF cdbl(sTotalFormFees) < sTotalPreviousPayments THEN
					sPayStatus="R"
			ELSEIF cdbl(sTotalFormFees) = sTotalPreviousPayments THEN
					sPayStatus="C"
			END IF

			' --- Copies the data to the RegisterGenNew and RegisterEvents tables
			WhichTable=RegGenTableName
			CopyDataToTables


			' ++++ MOVE THIS TO A SUB THAT GETS INTITIATED WITH DOUBLE FORM POST ??? ++++

			' -------------------------------------------------------------------------------------------
			' --- Establish the setting for sPayType to determine whether to use PayPal or HQ Account ---
			' -------------------------------------------------------------------------------------------
			IF sTotalPreviousPayments < cdbl(sTotalFormFees) AND sHQAccount<>true THEN
					sPayType="PayPal"
			ELSEIF sTotalPreviousPayments < cdbl(sTotalFormFees) AND sHQAccount=true THEN
					sPayType="Card"
			ELSEIF sTotalPreviousPayments > cdbl(sTotalFormFees) THEN
					sPayType="Refund"
			END IF

			' --- Establishes the total amount to be recorded in Payment Transaction table ---		
			sPayAmount = sTotalFormFees-sTotalPreviousPayments



			' -----------------------------------------------------
			' --- Creates a record of this order before sending ---
			' -----------------------------------------------------
			IF TRIM(Session("sOrderNo"))="" THEN 
					InitializePaymentRecord
			ELSE
					sOrderNo=Session("sOrderNo")
			END IF


			' --- If nothing owed AND this is NOT an AdminUser then skip to Tab 7 ---
			IF nav=6 AND sPayAmount=0 AND TestValidAdminCode<>true AND Session("AdminMenuLevel")<50 THEN
					nav=7
			END IF

			' ++++++++++++++++++++++
			' --- TESTING SCRIPT ---
			' ++++++++++++++++++++++
			IF sMemberID="000001151" AND TestMode="yes" THEN
					response.write("<br><br>Above nav=7 Branch")
					response.write("<br>Session(sOrderNo)="&Session("sOrderNo"))
					response.write("<br>sPayAmount = "&sPayAmount)
					'response.end
			END IF

	END IF



	' -------------------
	' --- RECEIPT TAB ---
	' -------------------
	IF nav=7 THEN

			' --- Update PaymentLog Record even if failed ---
			WhichTable=RegGenTableName
			InitializeFromTable


			' ++++++++++++++++++++++
			' --- TESTING SCRIPT ---
			' ++++++++++++++++++++++
			IF sMemberID="000001151" AND TestMode="yes" THEN
					response.write("<br>Above IF in nav=7 Branch")
					response.write("<br>sPayType="&sPayType)
			END IF

		
			' -----------------------------------------------
			' --- Perform validation on Result of Payment --- 
			' -----------------------------------------------
			IF sPayType="PayPal" THEN
					' --- Tests for matching OrderNo on return from PayPal and sets Result=0 ---
					ValidatePayPal
			ELSEIF sPayType="Check" OR sPayType="Cash" OR sPayType="Refund" OR sPayType="NoSale" THEN
					sPaymentResult="0"			
			ELSEIF sPayType="Card" THEN
					' --- Tests for Result=0 when it is a CC transaction ---
					ValidateCreditCard
			ELSEIF sPayType="PPErr" THEN
					sPaymentResult=""
					response.write("Paypal Returned the correct info for payment failure")
			ELSEIF sPayType="ByPass" THEN
					sPaymentResult=""			
			END IF

			' ---------------------------------------
			' --- Updates Payment Transaction Log ---
			' ---------------------------------------
			UpdatePaymentTransaction


			' ++++++++++++++++++++++
			' --- TESTING SCRIPT ---
			' ++++++++++++++++++++++
			IF sMemberID="000001151" AND TestMode="yes"  THEN
					response.write("<br><br>After UpdatePaymemtTransaction")
					response.write("<br>sPaymentResult="&sPaymentResult)
			END IF
			IF sPaymentResult="0" THEN

					' ++++++++++++++++++++++
					' --- TESTING SCRIPT ---
					' ++++++++++++++++++++++
					IF sMemberID="000001151" AND TestMode="yes"  THEN
							response.write("<br><br>INSIDE IF sPaymentResult=0")
							response.write("<br>sPaymentResult="&sPaymentResult)
					END IF


					' --- Change PayStatus in RegGen to Complete [C] ---
					UpdateRegGenPaymentStatus

					' --- Update transaction table with detail ---
					UpdateTransTable
	
					' --- Sets value of sTotalPreviousPayments ---
					DetermineTotalFeesActuallyPaid

					' --- Sets Session Variables for Payment Status Text ---
					SetSessionStatusText

					' --- Send email acknowledgement ---
					SendEntryConfirm

					' --- Reset Session variable sOrderNo so any additional changes will be under a new OrderNo ---			
					Session("sOrderNo")=""
			ELSE
					' --- Cause form to display PAYMENT ERROR message
			END IF
	
	END IF


	' ----------------------------------------
	' --------  BEGIN DISPLAYING DATA  -------
	' ----------------------------------------
	' --- DisplayAccordion is a procedure in RegFormDisplay.asp ---
	sErrorNo = 0
	DisplayAccordion


	' ------------------------------------------------------------------------
	' --- Done with this transaction so release to allow the next sOrderNo ---
	' ------------------------------------------------------------------------
	IF nav=7 THEN
			zx=2
			' --- Function EntriesExceedLimit in tools_registration.asp ---
			IF EntriesExceedLimit(sTourID) OR zx=1 THEN
					' --- Sub in tools_registration.asp ---
					SendTourFullEMail
			END IF
	END IF


END SELECT



IF sRunByWhat <> "Print" THEN
	WriteIndexPageFooter
END IF




'END IF




' =======================================================================================================================================
' =====================================  END OF MAIN PROGRAM ============================================================================
' =======================================================================================================================================



' --------------------------
  FUNCTION IsMemberPaid
' --------------------------
	SET rs=Server.CreateObject("ADODB.recordset")
	sSQL = "SELECT TOP 1 MemberID FROM "&RegPaymentTableName
	sSQL = sSQL + " WHERE MemberID='"&sMemberID&"' AND TourID='"&sTourID&"' AND Result='0'"
	sSQL = sSQL + " ORDER BY TransDate"  	
	rs.open sSQL, sConnectionToTRATable, 3, 1

	IsMemberPaid=false
	IF NOT rs.eof THEN IsMemberPaid=true
' --- For Test
'IsMemberPaid=false

  END FUNCTION



' ------------------------------
   SUB ResetMembTourSessionVar
' ------------------------------

' --- Resets the important session variables for the member ---
' --- These have to be reset because the are dependent on the tournament
Session("sEnableGR")=""
Session("sEnableStd")=""
Session("sEnableRec")=""
Session("sMembCanSkiText")=""
Session("sMembCanSkiColor")=""
Session("sOrderNo")


END SUB




' ---------------------------
  SUB DisplayOLRDisabledNotice
' ---------------------------

' --- Resets some Member variables and the sTourID Session ---
ResetMembTourSessionVar
Session("sTourID")=""


	%>
	<html>
	<head>
	<title>Online Entry Disabled Notice</title>
	</head>
	<body>

	<br>
	<TABLE class="messagetable" ALIGN="center" width=70% >
	  <tr>
	    <th align=center><font size="4" color="#FFFFFF"><b>Important Notice</b></font></th>
          </tr>

	  <tr>
	    <td align="center">
		<br>
		<font size="<%=fontsize3%>" color=red><b>Online Registration has been de-activated by the registrar for</b></font>   
		<br><br>
		<font size="<%=fontsize4%>" color=blue><b><%=sTourName%></b></font>   
		<br><br>
		<font size="<%=fontsize2%>">Contact tournament registrar for more information</font>
	    </td>
	  </tr>

	  <tr>
	    <td align="center">
		<form action="/rankings/<%=RegFileName%>" method="post">
			<input type="submit" name="Action" style="width:15em;" value="Select Another Tournament">
		</form>
	    </td>
	  </tr>


	</TABLE>
	</body>
	</html><%

	' --- Do not remove Response.end --- user must take action with the button
	response.end

END SUB




' ---------------------------
  SUB DisplayNotSetupNotice
' ---------------------------


	%><br>
	<TABLE class="innertable" ALIGN="center" width=70% >
	<tr>
	  <th align=center><font size="4" color="#FFFFFF"><b>Online Registration Not Activated</b></font></th>
        </tr>
	  <td align="center">
	  <br>
	     	<font size="<%=fontsize2%>">You have reached this page in error. </font>
		<br><br>
		<font size="<%=fontsize2%>" color=red><b>Online Registration has not been activated for this tournament.</b></font>   
		<br><br>
		<font size="<%=fontsize2%>"><p>For more information about USA Water Ski, contact:
		<font size=<% =fontsize2 %> color="blue"><A href="http://www.usawaterski.org/index1.html">USA Water Ski</a></font>
		<br><br>
		<font size="<%=fontsize2%>">For general questions, contact: USA Water Ski Competition Dept at 800-533-2972.</font>
	     
	    <form align="center" action="http://www.usawaterski.org/index1.html" method="post">
	    	<input type="submit" value="Main Menu">
	    </form>
	  </td>
	</tr>
	</TABLE>
	<% 

END SUB





' -------------------------------
    SUB VerifyMemberHistoryUpdate
' -------------------------------

'sMemberID="000001151"


Dim tMembershipTypeCode

' --- REMOVE after DEBUGGING ---
' --- Used this to test my expiration date (12/31/2008) against a �bogus� sTDateS of 1/1/2009 to trick the system into
' ---  initially thinking membership has expired.  The IF statement below, then reset the date to insure screens behave as intended.

'IF sMemberID="000001151" THEN sTDateS="08/20/2008"

' --- Checks the Member History Table on HQ server to if Effectiveto is now current ---
Set SQLConnect = CreateObject("ADODB.Connection")
SQLConnect.Open Application("HQSQLConn")


'sSQL = "SELECT TOP 1 [Membership Type Code] AS tblMembTypeCode, EffectiveTo FROM [Membership History] WHERE [Person ID]='"&sMemberID&"' ORDER BY EffectiveTo DESC"

Set RS = SQLConnect.Execute("SELECT TOP 1 [Membership Type Code] AS tblMembTypeCode, EffectiveTo FROM [Membership History] WHERE [Person ID]='"&sMemberID&"' ORDER BY EffectiveTo DESC")
IF NOT rs.eof THEN
	' ---  Checks End Date of tournament against Expiration Date of membership record  ---
	IF DateDiff("d", rs("EffectiveTo"), sTDateS) <= 0  THEN
		Session("sExpirationStatusText")="OK - "&rs("EffectiveTo")
		Session("sExpirationStatusColor")="blue"
	END IF
	tMembershipTypeCode=rs("tblMembTypeCode")
END IF


'response.write("<br>VMHU - tMembershipTypeCode = "&tMembershipTypeCode)
'response.write("<br>VMHU "&sSQL)

'response.end


rs.close




' --- Uses value from check of membership to check the Jaguar MembershipType Table for competition status ---
SET rs=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT CanSkiInTournaments, CanSkiInGRTournaments FROM "&MemberTypeOLRTableName
sSQL = sSQL + " WHERE MembershipTypeID='"&tMembershipTypeCode&"'"
rs.open sSQL, sConnectionToTRATable, 3, 1

IF NOT rs.eof THEN
	IF rs("CanSkiInTournaments") = true THEN
		Session("sMembCanSkiText")="OK - "&sMembTypeCode&"(MH) - "&sTypeDesc
		Session("sMembCanSkiColor")=TextColor2

		Session("sEnableGR")="Y"
		Session("sEnableStd")="Y"
		Session("sEnableRec")="Y"

	' --- NEW CODE related to GR membership type
	ELSEIF rs("CanSkiInGRTournaments")= true AND (sGRTournament=true OR sGRFunDay=true) THEN
		Session("sMembCanSkiText")="OK - "&sMembTypeCode&" - "&sTypeDesc
		Session("sMembCanSkiColor")=TextColor2
		Session("sEnableGR")="Y"

	END IF
END IF



' ********   CODE FROM DefineMemberVariables section  **********

' --- Has not previously been set and can ski then set to OK in first two positions ---
'IF TRIM(Session("sMembCanSkiText"))="" AND sCanSkiTour = 1 THEN  
'	Session("sMembCanSkiText")="OK - "&sMembTypeCode&" - "&sTypeDesc
'	Session("sMembCanSkiColor")=TextColor2
'	Session("sEnableGR")="Y"
'	Session("sEnableStd")="Y"
'	Session("sEnableRec")="Y"

' --- TEMPORARY FIX for GR Membership ---
'ELSEIF TRIM(Session("sMembCanSkiText"))="" AND sCanSkiGRTour = 1  AND (sGRTournament=true OR sGRFunDay=true) THEN
'	Session("sMembCanSkiText")="OK - "&sMembTypeCode&" - "&sTypeDesc
'	Session("sMembCanSkiColor")=TextColor2
'	Session("sEnableGR")="Y"

' --- Has not previously been set and CANNOT ski then set to upgrade condition ---
'ELSEIF TRIM(Session("sMembCanSkiText"))="" AND sCanSkiTour = 0 THEN
'	Session("sMembCanSkiText")=sMembTypeCode&" - Competition Upgrade Required"
'	Session("sMembCanSkiColor")="red"
'END IF






rs.close


END SUB



' -----------------------
  SUB ValidateCreditCard
' -----------------------


' **************************************************************
' ---  **** NEED ORDER STATEMENT OR SOMETHING HERE  ************
' **************************************************************



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

	IF cdbl(MaxOrder)=cdbl(sOrderNo) THEN sPaymentResult="0"
	sPayAmount = sTotalFormFees-sTotalPreviousPayments

END IF

'response.end




IF sMemberID="000001151" AND TestMode="yes"  THEN
	response.write("<br><br>In ValidateCreditCard")
	'response.write("<br>rsPayLog.eof=")
	'response.write(rsPayLog.eof)
	response.write("<br>rsPayLog(MaxOrder)="&rsPayLog("MaxOrder"))
	response.write("<br>RCount="&RCount)
	response.write("<br>sPaymentResult="&sPaymentResult)
END IF


END SUB




' -----------------------
   SUB ValidatePayPal
' -----------------------

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

	IF cdbl(MaxOrder)=cdbl(sOrderNo) THEN 
			sPaymentResult="0"
			
	END IF		
	sPayAmount = sTotalFormFees-sTotalPreviousPayments

END IF

END SUB




' ---------------------------
  SUB ReleaseSesVars
' ---------------------------

Session.Abandon

END SUB



' ---------------------------
  SUB ValidateFormValues
' ---------------------------


sFormError=""

		
IF sTPandC=true AND ((sSelectEvent(1)="on" AND cdbl(sFeeRounds(1))<cdbl(1)) OR (sSelectEvent(2)="on" AND cdbl(sFeeRounds(2))<cdbl(1)) OR (sSelectEvent(3)="on" AND cdbl(sFeeRounds(3))<cdbl(1)) OR (sSelectEvent(4)="on" AND cdbl(sFeeRounds(4))<cdbl(1))) THEN
		sFormError="The number of <b>Rounds</b> for each event you are entering must be at least '1'." 
END IF

IF sTPandC=true AND cdbl(sTPandCPulls)<(cdbl(sFeeRounds(1))+cdbl(sFeeRounds(2))+cdbl(sFeeRounds(3))+cdbl(sFeeRounds(4))) THEN
		sFormError="The <b>Maximum</b> total Rounds/Pulls for all events combined cannot exceed "&sTPandCPulls&"." 
END IF

Dim AtLeastOneEvent
AtLeastOneEvent=false

' --- Loops thru all events and tests if at least one button is checked ---
FOR EvtNo=1 TO TotEv
		IF sSelectEvent(EvtNo)="on" AND sShowSkills=true AND TRIM(sSkill(EvtNo))="" AND sTEvent(EvtNo)="WB" THEN
				sFormError="You must select a Skill Level for each event entered." 
		END IF

		IF sSelectEvent(EvtNo)="on" AND TRIM(sFeeClass(EvtNo))="" THEN
				sFormError="You must select an Entry Classification for each event entered." 
		END IF
		IF sSelectEvent(EvtNo)="on" THEN AtLeastOneEvent=true

		IF sSelectEvent(EvtNo)="on" AND RIGHT(TRIM(sTEvent(EvtNo)),1)="T" AND TRIM(sBoat(EvtNo))="" THEN  
				IF TRIM(sFormError)="" THEN sFormError="You must select a TRICK boat."
		END IF
NEXT


'Response.write("<br>TestValidAdminCode=")
'Response.write(TestValidAdminCode)

IF TestValidAdminCode THEN 
		' --- Do nothing because it is admin ---	
ELSE	
		IF AtLeastOneEvent=false THEN
				sFormError="To continue, you must select at least one event."
		END IF
END IF



IF sFormError<>"" THEN
		'response.write("<br> IN FORM ERROR "&sFormError)
		PreviousButtonStatus="enabled"
		EditButtonStatus="disabled"
		MainButtonStatus="enabled"
		MainStatusValue="Verify"
		AllObjectStatus="enabled"
END IF


END SUB



' ---------------------------
  SUB ValidateFormDateEntered
' ---------------------------

' --- Test for valid Registration date - could only be invalid if entered incorrectly by Admin level ---
sFormError=""

SlashCount = UBound(split(sMembRegDate, "/"))
IF NOT(IsDate(sMembRegDate)) OR SlashCount<>2 OR LEN(Year(sMembRegDate))<>4 THEN
		sFormError="Please enter a valid date format mm/dd/yyyy"
END IF

p=2
IF p=1 AND sMemberID="000001151" THEN

		SlashCount = UBound(split(sMembRegDate, "/"))
		'IF NOT(IsDate(sMembRegDate)) OR SlashCount<>2 THEN
		'		sFormError="Please enter a valid date format mm/dd/yyyy"
		'END IF

		response.Write("<br> Line 1318 Registration.asp --- Test for Mark Crone Only = ")
		Response.Write("<br>IsDate(sMembRegDate)= "&sMembRegDate)
		Response.Write("<br>Valid Date = ") 
		Response.Write(IsDate(sMembRegDate)) 
		Response.Write("<br>Instr(sMembRegDate,/) = "&Instr(sMembRegDate,"/"))
		Response.Write("<br>SlashCount = "&SlashCount)
		Response.Write("<br>sFormError = "&sFormError) 
		'response.end
END IF

'response.write(LEN(Year(sMembRegDate)))
'response.write("<br>")
'response.write(Year(sMembRegDate))

'IF sMemberID="000001151" THEN
	IF sFormError<>"" THEN
		'response.write("<br> IN FORM ERROR "&sFormError)
		PreviousButtonStatus="disabled"
		EditButtonStatus="enabled"
		MainButtonStatus="disabled"
		'MainStatusValue="Continue"
		MainStatusValue="Continue"
		AllObjectStatus="disabled"
	ELSE
		'response.write("<br> IN FORM ERROR "&sFormError)
		PreviousButtonStatus="enabled"
		EditButtonStatus="enabled"
		MainButtonStatus="enabled"
		MainStatusValue="Continue"
		AllObjectStatus="enabled"
	END IF
'END IF


END SUB


' --------------------------------
   SUB SetNavigationVariables
' --------------------------------


'response.write("<br>SET MainStatus = "& TRIM(request("MainStatus")))

SELECT CASE nav

  CASE 1, 2, 3, 4, 5, 6

	IF TRIM(request("MainStatus"))="Verify" THEN
		PreviousButtonStatus="enabled"
		EditButtonStatus="enabled"
		MainButtonStatus="enabled"
		MainStatusValue="Continue"
		AllObjectStatus="disabled"

'response.write("<br>Updated - MainStatusValue="&MainStatusValue)

	ELSEIF TRIM(request("MainStatus"))="Continue" THEN
		nav=nav+1
		PreviousButtonStatus="enabled"
		EditButtonStatus="disabled"
		MainButtonStatus="enabled"
		MainStatusValue="Verify"
		AllObjectStatus="enabled"
'response.write("<br>Updated - MainStatusValue="&MainStatusValue)

	ELSEIF TRIM(request("Edit"))="Edit" THEN
		PreviousButtonStatus="enabled"
		EditButtonStatus="disabled"
		MainButtonStatus="enabled"
		MainStatusValue="Verify"
		AllObjectStatus="enabled"

	ELSEIF TRIM(request("Previous"))="Previous" THEN
		nav=nav-1
		PreviousButtonStatus="enabled"
		EditButtonStatus="enabled"
		MainButtonStatus="enabled"
		MainStatusValue="Verify"
		NextStatus="enabled"
		AllObjectStatus="disabled"

	ELSEIF TRIM(request("Next"))="Next" THEN
		nav=nav+1
		PreviousButtonStatus="enabled"
		EditButtonStatus="disabled"
		MainButtonStatus="enabled"
		MainStatusValue="Verify"
		NextStatus="enabled"
		AllObjectStatus="enabled"
	
	ELSE
		PreviousButtonStatus="enabled"
		EditButtonStatus="disabled"
		MainButtonStatus="enabled"
		MainStatusValue="Verify"
		AllObjectStatus="enabled"
		'Session("FinancialComplete")=true

	END IF  

END SELECT


END SUB





' ----------------------------
  SUB InitializePaymentRecord
' ----------------------------

' --- Required for PayPal since PayPal button posts directly to PayPal site --
' --- Finds last OrderNo and increments by one then saves Pending Record to Payment Log File  ----

SET rsPayLog=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT Count(*) AS RCount, MAX(OrderNo) AS MaxOrder FROM "&RegPaymentTableName
rsPayLog.open sSQL, sConnectionToTRATable, 3, 1

IF rsPayLog("RCount") = 0 THEN
		sOrderNo = 2000
ELSE	
		sOrderNo = rsPayLog("MaxOrder") + 1
END IF

'IF sMemberID="000001151" THEN Response.write("Line 1535 Registration - sPayType = " & sPayType)
	
IF sPayType<>"Card" THEN	' --- Initialization not needed when sPayType=Card.  This is done in card processor. ---

		DateNow = Now
		sAmount=0.00

		Dim tLastName, tFirstName, tAddress1, tCity, tState, tZipCode, tEmail
		'tLastName=""
		'tFirstName=""
		tAddress1=""
		tCity=""
		tState=""
		tZipCode=""
		tEmail=""

		Dim sLast4Card, sExpMonth, sExpYear, sApvl_Code, sCvv2_Resp, sAVS_Resp, sCheckNo, sPayStatus		
		sCheckNo=""
		sLast4Card=""
		sExpMonth=""
		sExpYear=""
		sApvl_Code=""
		sCvv2_Resp=""
		sAVS_Resp=""		
		sResult=""
		sMessage="Initialized"
		sPayStatus="I"

		' ------------------------		
		' --- PayStatus Decode ---
		' ------------------------
		' --- I = Initialized in OLR
		' --- C = PayPal IPN fired
		' --- O = Made it to Receipt Page of OLR
		
		sSQL = "INSERT INTO "&RegPaymentTableName
		sSQL = sSQL + " (MemberID, TourID" 
		sSQL = sSQL + ", FirstName, LastName, Address1, City, State, ZipCode, Email" 
		sSQL = sSQL + ", Amount, OrderNo, Txn_ID, TransDate, PayType, Result, Message"
		sSQL = sSQL + ", CheckNo, Last4Card, Apvl_Code, Cvv2_Resp, AVS_Resp, ExpYear, ExpMonth, PayStatus)" 		

		
		sSQL = sSQL + " VALUES ('"&sMemberID&"', '"&sTourID&"'"
		sSQL = sSQL + ", '"&sFirstName&"', '"&sLastName&"', '"&tAddress1&"', '"&tCity&"', '"&tState&"', '"&tZipCode&"', '"&tEmail&"'" 
		sSQL = sSQL + ", '"&sPayAmount&"', "&sOrderNo&", '"&sTxn_ID&"', '"&DateNow&"', '"&sPayType&"', '"&sResult&"', '"&sMessage&"'"
		sSQL = sSQL + ", '"&sCheckNo&"', '"&sLast4Card&"', '"&sApvl_Code&"', '"&sCvv2_Resp&"', '"&sAVS_Resp&"'"
		sSQL = sSQL + ", '"&sLast4Card&"', '"&sExpMonth&"', '"&sPayStatus&"')"

'IF sMemberID="000001151" THEN 
'	response.write(sSQL)
'	response.end
'END IF
		OpenCon
		con.execute(sSQL)
		CloseCon
END IF

Session("sOrderNo")=sOrderNo


END SUB



' ------------------------------
   SUB UpdateRegGenPaymentStatus
' ------------------------------

' ---- Finds last OrderNo and updates Payment Log File  ----

' ------------------------		
' --- PayStatus Decode ---
' ------------------------
' --- I = Initialized in OLR
' --- N = PayPal IPN fired
' --- C = Complete


sSQL = "UPDATE "&RegGenTableName
sSQL = sSQL + " Set PayStatus='C'" 
sSQL = sSQL + " WHERE MemberID='"&sMemberID&"' AND TourID='"&sTourID&"'"

OpenCon
con.execute(sSQL)
CloseCon


END SUB




' ---------------------
  SUB CheckWaiverStatus
' ---------------------

'response.write("sWaiverCode = "&sWaiverCode)
IF TRIM(sWaiverCode) = "" THEN 
		Session("sReleaseText")="Not Signed"
		Session("sReleaseTextColor")="red"
ELSEIF sWaiverCode = adult_waiver OR sWaiverCode = minor_waiver THEN 
		Session("sReleaseText")="Complete"
		Session("sReleaseTextColor")="yellow"
ELSEIF sWaiverCode = "Paper" THEN 
		Session("sReleaseText")="Paper Waiver"
		Session("sReleaseTextColor")="yellow"
ELSE 
		Session("sReleaseText")="No Waiver on File"
		Session("sReleaseTextColor")="red"
END IF 

END SUB



' ----------------------------------
  SUB DetermineTotalFeesActuallyPaid
' ----------------------------------

' --- Read transactions from Credit Card Table to determine Total Fees actually completed ----
SET rsPayLog=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT MemberID, SUM(Amount) AS TotalPreviousFees FROM "&RegPaymentTableName
sSQL = sSQL + " WHERE Left(TourID,6) = '" &LEFT(sTourID,6)& "' AND MemberID = '"&sMemberID&"' AND Result='0'"
sSQL = sSQL + " GROUP BY MemberID "
rsPayLog.open sSQL, SConnectionToTRATable, 3, 3

sTotalPreviousPayments=cdbl(0)

IF NOT rsPayLog.eof THEN
		sTotalPreviousPayments = cdbl(rsPayLog("TotalPreviousFees"))
END IF

rsPayLog.close


END SUB




' -------------------------- 
   SUB CopyDataToTables
' --------------------------

	' -------------------------  UPDATE OR ADD to General Table  ----------------------

	' ---- Read WhichTable for existing record ----
	SET rs=Server.CreateObject("ADODB.recordset")
	sSQL = "SELECT * FROM "&WhichTable
	sSQL = sSQL + " WHERE Left(TourID,6) = '" & SQLClean(left(sTourID,6)) & "' AND MemberID = '"&sMemberID&"'"
	rs.open sSQL, SConnectionToTRATable, 3, 3

	IF sLateFeeTot>999 THEN		' --- Protect against entering a NEW record in an old tournament until RegDate can be changed ---
			sTotalFormFees=sTotalFormFees-sLateFeeTot
			sLateFeeTot=11
	END IF

	OpenCon
	IF NOT rs.eof	THEN	' --- Found EXISTING record
			sSQL = "UPDATE "&WhichTable
			sSQL = sSQL + " SET MemberID = '"&sMemberID&"', TourID = '"&sTourID&"'"
			sSQL = sSQL + " , EntryFee = '"&sEntryFee&"', TotalEntry = '"&sTotalFormFees&"', LateFee = '"&sLateFeeTot&"'"
			sSQL = sSQL + " , AWSEFDonation = '"&sAWSEFDonation&"', OffDisc = '"&sOffDiscAmt&"', JrDisc = '"&sJrDiscAmt&"', ClubDisc = '"&sClubDiscAmt&"'"
			sSQL = sSQL + " , RampHeight = '"&sRampHeight&"', RegisterDate = '"&sMembRegDate&"', EntryType = '"&sEntryType&"'"
			sSQL = sSQL + " , MembOverride = '"&sMembOverride&"', RegionalOverride = '"&sRegionalOverride&"'"
			'IF sMemberID="700133787" THEN 
					sSQL = sSQL + " , WaiverCode = '"&sWaiverCode&"', SignWaiver='"&SQLClean(sSignWaiver)&"'"
			'ELSE
			'		sSQL = sSQL + " , WaiverCode = '"&sWaiverCode&"', SignWaiver='"&sSignWaiver&"'"
			' END IF		
			sSQL = sSQL + " , MoneyOverride = '"&sMoneyOverride&"', BanquetQty='"&sBanquetQty&"', BanquetFee = '"&sBanquetTot&"', PayStatus='"&sPayStatus&"'" 	
			sSQL = sSQL + " , OF1Qty = '"&sOF1Qty&"', OF2Qty = '"&sOF2Qty&"', OF3Qty = '"&sOF3Qty&"', OF4Qty = '"&sOF4Qty&"', OF5Qty = '"&sOF5Qty&"'"
			sSQL = sSQL + " , OF6Qty = '"&sOF6Qty&"', OF7Qty = '"&sOF7Qty&"', OF8Qty = '"&sOF8Qty&"', OF9Qty = '"&sOF9Qty&"', OF10Qty = '"&sOF10Qty&"'"
			sSQL = sSQL + " , OF1Fee = '"&sOF1Fee&"', OF2Fee = '"&sOF2Fee&"', OF3Fee = '"&sOF3Fee&"', OF4Fee = '"&sOF4Fee&"', OF5Fee = '"&sOF5Fee&"'"
			sSQL = sSQL + " , OF6Fee = '"&sOF6Fee&"', OF7Fee = '"&sOF7Fee&"', OF8Fee = '"&sOF8Fee&"', OF9Fee = '"&sOF9Fee&"', OF10Fee = '"&sOF10Fee&"'"

			sSQL = sSQL + " WHERE Left(TourID,6) = '" & SQLClean(left(sTourID,6)) & "' AND MemberID = '"&sMemberID&"'"
			IF sMemberID="700133787" THEN 
					' response.write("<br>IN UPDATE -sMembRegDate = "&sMembRegDate)
					'response.write("<br>"&sSQL)
					'response.end
			END IF
			con.execute(sSQL)


	ELSE			' --- No existing so ADD new record  ---

			' --- Creates current date in BOTH tables, even if the Temp and Gen tables are not created on same date ---
			'sMembRegDate = DATE

			sSQL = "INSERT INTO "&WhichTable
			sSQL = sSQL + " (MemberID, TourID"
			sSQL = sSQL + ", TotalEntry, EntryFee, LateFee, OtherFee, AWSEFDonation, JrDisc, SrDisc, OffDisc"
			sSQL = sSQL + ", ClubDisc, RampHeight, RegisterDate, EntryType"
			sSQL = sSQL + ", MembOverride, RegionalOverride, MoneyOverride, BanquetQty, BanquetFee, PayStatus"
			sSQL = sSQL + ", WaiverCode, SignWaiver"
			sSQL = sSQL + ", OF1Qty, OF2Qty, OF3Qty, OF4Qty, OF5Qty"
			sSQL = sSQL + ", OF6Qty, OF7Qty, OF8Qty, OF9Qty, OF10Qty"
			sSQL = sSQL + ", OF1Fee, OF2Fee, OF3Fee, OF4Fee, OF5Fee"
			sSQL = sSQL + ", OF6Fee, OF7Fee, OF8Fee, OF9Fee, OF10Fee)"

			sSQL = sSQL + " VALUES ('"&sMemberID&"', '"&sTourID&"'"
			sSQL = sSQL + ", '"&sTotalFormFees&"', '"&sEntryFee&"', '"&sLateFeeTot&"', '"&sOtherFee&"', '"&sAWSEFDonation&"', '"&sJrDiscAmt&"', '"&sSrDiscAmt&"', '"&sOffDiscAmt&"'"
			sSQL = sSQL + ", '"&sClubDiscAmt&"', '"&sRampHeight&"', '"&sMembRegDate&"', '"&sEntryType&"'"
			sSQL = sSQL + ", '"&sMembOverride&"', '"&sRegionalOverride&"', '"&sMoneyOverride&"', '"&sBanquetQty&"', '"&sBanquetTot&"', 'O'"
			sSQL = sSQL + ", '"&sWaiverCode&"', '"&SQLClean(sSignWaiver)&"'"
			sSQL = sSQL + ", '"&sOF1Qty&"', '"&sOF2Qty&"', '"&sOF3Qty&"', '"&sOF4Qty&"', '"&sOF5Qty&"'"
			sSQL = sSQL + ", '"&sOF6Qty&"', '"&sOF7Qty&"', '"&sOF8Qty&"', '"&sOF9Qty&"', '"&sOF10Qty&"'"
			sSQL = sSQL + ", '"&sOF1Fee&"', '"&sOF2Fee&"', '"&sOF3Fee&"', '"&sOF4Fee&"', '"&sOF5Fee&"'"
			sSQL = sSQL + ", '"&sOF6Fee&"', '"&sOF7Fee&"', '"&sOF8Fee&"', '"&sOF9Fee&"', '"&sOF10Fee&"')"

			con.execute(sSQL)

	END IF

	IF WhichTable=RegGenTableName THEN
			DetailTable=RegDetailTableName
	ELSEIF WhichTable=RegTempTableName THEN
			DetailTable=RegDetailTempTableName		
	END IF

  CopyToEventDetail

	rs.close


END SUB






'---------------------------
  SUB CopyToEventDetail
'---------------------------

'response.write("<br>CopyToEvent")
'response.write("<br>sDiv(1)="&sDiv(1)&" - sDiv(2)="&sDiv(2))


' ---- DELETE the EVENT detail from WHICHTABLE for this MemberID/TourID ----
sSQL = "DELETE FROM "&DetailTable
sSQL = sSQL + " WHERE Left(TourID,6) = '"&left(sTourID,6)&"' AND MemberID = '"&sMemberID&"'"
con.execute(sSQL)


FOR EvtNo = 1 TO TotEv
	' ---- INSERT the new EVENT records into DETAILTABLE for this MemberID/TourID ----
	IF TRIM(sDiv(EvtNo))<>"" THEN

		'response.write("<br>CopyToEventDetail - sSkill(EvtNo)="&sSkill(EvtNo))
		sSQL = "INSERT INTO "&DetailTable
		sSQL = sSQL + " (MemberID, TourID, Event, Div, QfyOverride, FeeClass, FeeRounds, Boat, Skill)" 
		sSQL = sSQL + " VALUES ('"&sMemberID&"', '"&sTourID&"', '"&sTEvent(EvtNo)&"', '"&sDiv(EvtNo)&"',"
		sSQL = sSQL + " '"&sQfyOverride(EvtNo)&"', '"&sFeeClass(EvtNo)&"', '"&sFeeRounds(EvtNo)&"',"
		sSQL = sSQL + " '"&sBoat(EvtNo)&"', '"&sSkill(EvtNo)&"')"
		con.execute(sSQL)

	'response.write("<br>"&sSQL)

	END IF

NEXT  



'DisplayCurrentValues "Bottom of CopyToEventDetail "&DetailTable

END SUB



' ----------------------- 
   SUB UpdateTransTable
' -----------------------

	' -------  Finds total of transactions in RegPaymentTableName  -------
	DetermineTotalFeesActuallyPaid


	' ---------------------------------------------------------------------------------------
	' --- NOTE:  Transaction table is updated only when payment is deemed to be a SUCCESS ---
	' ---------------------------------------------------------------------------------------

	' ---- Read RegTempTableName for the updated record ----
	SET rsRegTemp=Server.CreateObject("ADODB.recordset")
	sSQL = "SELECT * FROM "&RegTempTableName
	sSQL = sSQL + " WHERE Left(TourID,6) = '"&left(sTourID,6)&"' AND MemberID = '"&sMemberID&"'"
	rsRegTemp.open sSQL, SConnectionToTRATable, 3, 3




	' ---------------------------------------------
	' ---- Gets latest transaction date/time  -----
	' ---------------------------------------------
	Dim mdate

	SET rsRegTrans=Server.CreateObject("ADODB.recordset")
	sSQL = "(SELECT MAX(TransDate) AS maxdate FROM "&RegTransTableName
	sSQL = sSQL + " WHERE Left(TourID,6) = '" & SQLClean(left(sTourID,6)) & "' AND MemberID = '"&sMemberID&"') AS d"
	rsRegTrans.open sSQL, SConnectionToTRATable, 3, 3
	
	mdate = "01/01/2099"
	IF NOT rsRegTrans.eof THEN
		mDate = rsRegTrans("maxdate")     ' --- Latest date		

		' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
		' +++ Define All fields from current list of transactions +++
		' +++ Compare to Current Values and ONLY if different add a new Transaction Set +++ 
		' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

	END IF

	' ----  Reads all transactions with matching date/time  ----
	SET rsRegTrans=Server.CreateObject("ADODB.recordset")
	sSQL = "SELECT MemberID, TourID, TransCode, TransDate, -Amount AS Amount, TransNo, OrderNo FROM "&RegTransTableName
	sSQL = sSQL + " WHERE TransDate = '"&mDate&"' AND Left(TourID,6) = '" & SQLClean(left(sTourID,6)) & "' AND MemberID = '"&sMemberID&"'"
	rsRegTrans.open sSQL, SConnectionToTRATable, 3, 3

	OpenCon
	NowDate = NOW	


	' ----------------------------------------------------------------------------------------------------------------------
	' At least one record DOES exists so CREATE A CREDIT (i.e. reverse previous transaction) for EACH line 
	' ----------------------------------------------------------------------------------------------------------------------

	IF NOT rsRegTrans.eof THEN  	

		rsRegTrans.movefirst
		i=0
		
		Dim Cred_Code, Cred_Amt
		DO WHILE NOT rsRegTrans.eof
			Cred_Code = "C" + RIGHT(rsRegTrans("TransCode"),2)
			Cred_Amt = rsRegTrans("Amount")
			i = i + 1
		
			sSQL = "INSERT INTO " & RegTransTableName
			sSQL = sSQL + " (MemberID, TourID, TransCode, Amount, TransDate, TransNo, OrderNo)"
			sSQL = sSQL + " VALUES ('"&sMemberID&"', '"&sTourID&"', '"&Cred_Code&"', '"&Cred_Amt&"', '"&NowDate&"', '"&i&"', '"&sOrderNo&"')"
			con.execute(sSQL)

			rsRegTrans.movenext
			
			IF i=100 THEN EXIT DO
			
		LOOP
		
	END IF



	' ------------------------------------------------------------------------------------------------------------------------
	' ---------   Store each type of transaction from RegisterGeneral (current form values) and save to RegTransTable  -------
	' ------------------------------------------------------------------------------------------------------------------------

	' --- Initialize the TransNo counter and make new transaction date 1 sec later to differentiate from credits ---
	i = 0
	NowDate = DateAdd("s", 1, NowDate)

	IF cdbl(rsRegTemp("EntryFee")) > cdbl(0) THEN
		i = i +1 
		sSQL = "INSERT INTO " & RegTransTableName
		sSQL = sSQL + " (MemberID, TourID, TransCode, Amount, TransDate, TransNo, OrderNo)"
		sSQL = sSQL + " VALUES ('"&sMemberID&"', '"&sTourID&"', 'FEF', '"&rsRegTemp("EntryFee")&"', '"&NowDate&"', '"&i&"', '"&sOrderNo&"')"
		con.execute(sSQL)
	END IF


	IF cdbl(rsRegTemp("LateFee")) > 0 THEN 
		i = i +1
		sSQL = "INSERT INTO " & RegTransTableName
		sSQL = sSQL + " (MemberID, TourID, TransCode, Amount, TransDate, TransNo, OrderNo)"
		sSQL = sSQL + " VALUES ('"&sMemberID&"', '"&sTourID&"', 'FLF', '"&rsRegTemp("LateFee")&"', '"&NowDate&"', '"&i&"', '"&sOrderNo&"')"
		con.execute(sSQL)
	END IF		

	IF cdbl(rsRegTemp("AWSEFDonation")) > 0 THEN 
		i = i +1
		sSQL = "INSERT INTO " & RegTransTableName
		sSQL = sSQL + " (MemberID, TourID, TransCode, Amount, TransDate, TransNo, OrderNo)"
		sSQL = sSQL + " VALUES ('"&sMemberID&"', '"&sTourID&"', 'OBF', '"&rsRegTemp("AWSEFDonation")&"', '"&NowDate&"', '"&i&"', '"&sOrderNo&"')"
		con.execute(sSQL)
	END IF

	IF cdbl(rsRegTemp("OffDisc")) <> 0 THEN 
		i = i +1
		sSQL = "INSERT INTO " & RegTransTableName
		sSQL = sSQL + " (MemberID, TourID, TransCode, Amount, TransDate, TransNo, OrderNo)"
		sSQL = sSQL + " VALUES ('"&sMemberID&"', '"&sTourID&"', 'DOF', '"&rsRegTemp("OffDisc")&"', '"&NowDate&"', '"&i&"', '"&sOrderNo&"')"
		con.execute(sSQL)
	END IF

	IF cdbl(rsRegTemp("JrDisc")) <> 0 THEN 
		i = i +1
		sSQL = "INSERT INTO " & RegTransTableName
		sSQL = sSQL + " (MemberID, TourID, TransCode, Amount, TransDate, TransNo, OrderNo)"
		sSQL = sSQL + " VALUES ('"&sMemberID&"', '"&sTourID&"', 'DJR', '"&rsRegTemp("JrDisc")&"', '"&NowDate&"', '"&i&"', '"&sOrderNo&"')"
		con.execute(sSQL)
	END IF

	IF cdbl(rsRegTemp("SrDisc")) <> 0 THEN 
		i = i +1
		sSQL = "INSERT INTO " & RegTransTableName
		sSQL = sSQL + " (MemberID, TourID, TransCode, Amount, TransDate, TransNo, OrderNo)"
		sSQL = sSQL + " VALUES ('"&sMemberID&"', '"&sTourID&"', 'DSR', '"&rsRegTemp("SrDisc")&"', '"&NowDate&"', '"&i&"', '"&sOrderNo&"')"
		con.execute(sSQL)
	END IF

	IF cdbl(rsRegTemp("ClubDisc")) <> 0 THEN 
		i = i +1
		sSQL = "INSERT INTO " & RegTransTableName
		sSQL = sSQL + " (MemberID, TourID, TransCode, Amount, TransDate, TransNo, OrderNo)"
		sSQL = sSQL + " VALUES ('"&sMemberID&"', '"&sTourID&"', 'DCL', '"&rsRegTemp("ClubDisc")&"', '"&NowDate&"', '"&i&"', '"&sOrderNo&"')"
		con.execute(sSQL)
	END IF

	IF cdbl(rsRegTemp("BanquetFee")) <> 0 THEN 
		i = i +1
		sSQL = "INSERT INTO " & RegTransTableName
		sSQL = sSQL + " (MemberID, TourID, TransCode, Amount, TransDate, TransNo, OrderNo)"
		sSQL = sSQL + " VALUES ('"&sMemberID&"', '"&sTourID&"', 'BAN', '"&rsRegTemp("BanquetFee")&"', '"&NowDate&"', '"&i&"', '"&sOrderNo&"')"
		con.execute(sSQL)
	END IF

	IF cdbl(rsRegTemp("OF1Fee")) <> 0 THEN 
		i = i +1
		sSQL = "INSERT INTO " & RegTransTableName
		sSQL = sSQL + " (MemberID, TourID, TransCode, Amount, TransDate, TransNo, OrderNo)"
		sSQL = sSQL + " VALUES ('"&sMemberID&"', '"&sTourID&"', 'OF1', '"&rsRegTemp("OF1Fee")&"', '"&NowDate&"', '"&i&"', '"&sOrderNo&"')"
		con.execute(sSQL)
	END IF

	IF cdbl(rsRegTemp("OF2Fee")) <> 0 THEN 
		i = i +1
		sSQL = "INSERT INTO " & RegTransTableName
		sSQL = sSQL + " (MemberID, TourID, TransCode, Amount, TransDate, TransNo, OrderNo)"
		sSQL = sSQL + " VALUES ('"&sMemberID&"', '"&sTourID&"', 'OF2', '"&rsRegTemp("OF2Fee")&"', '"&NowDate&"', '"&i&"', '"&sOrderNo&"')"
		con.execute(sSQL)
	END IF

	IF cdbl(rsRegTemp("OF3Fee")) <> 0 THEN 
		i = i +1
		sSQL = "INSERT INTO " & RegTransTableName
		sSQL = sSQL + " (MemberID, TourID, TransCode, Amount, TransDate, TransNo, OrderNo)"
		sSQL = sSQL + " VALUES ('"&sMemberID&"', '"&sTourID&"', 'OF3', '"&rsRegTemp("OF3Fee")&"', '"&NowDate&"', '"&i&"', '"&sOrderNo&"')"
		con.execute(sSQL)
	END IF

	IF cdbl(rsRegTemp("OF4Fee")) <> 0 THEN 
		i = i +1
		sSQL = "INSERT INTO " & RegTransTableName
		sSQL = sSQL + " (MemberID, TourID, TransCode, Amount, TransDate, TransNo, OrderNo)"
		sSQL = sSQL + " VALUES ('"&sMemberID&"', '"&sTourID&"', 'OF4', '"&rsRegTemp("OF4Fee")&"', '"&NowDate&"', '"&i&"', '"&sOrderNo&"')"
		con.execute(sSQL)
	END IF

	IF cdbl(rsRegTemp("OF5Fee")) <> 0 THEN 
		i = i +1
		sSQL = "INSERT INTO " & RegTransTableName
		sSQL = sSQL + " (MemberID, TourID, TransCode, Amount, TransDate, TransNo, OrderNo)"
		sSQL = sSQL + " VALUES ('"&sMemberID&"', '"&sTourID&"', 'OF5', '"&rsRegTemp("OF5Fee")&"', '"&NowDate&"', '"&i&"', '"&sOrderNo&"')"
		con.execute(sSQL)
	END IF

	IF cdbl(rsRegTemp("OF6Fee")) <> 0 THEN 
		i = i +1
		sSQL = "INSERT INTO " & RegTransTableName
		sSQL = sSQL + " (MemberID, TourID, TransCode, Amount, TransDate, TransNo, OrderNo)"
		sSQL = sSQL + " VALUES ('"&sMemberID&"', '"&sTourID&"', 'OF6', '"&rsRegTemp("OF6Fee")&"', '"&NowDate&"', '"&i&"', '"&sOrderNo&"')"
		con.execute(sSQL)
	END IF

	IF cdbl(rsRegTemp("OF7Fee")) <> 0 THEN 
		i = i +1
		sSQL = "INSERT INTO " & RegTransTableName
		sSQL = sSQL + " (MemberID, TourID, TransCode, Amount, TransDate, TransNo, OrderNo)"
		sSQL = sSQL + " VALUES ('"&sMemberID&"', '"&sTourID&"', 'OF7', '"&rsRegTemp("OF7Fee")&"', '"&NowDate&"', '"&i&"', '"&sOrderNo&"')"
		con.execute(sSQL)
	END IF

	IF cdbl(rsRegTemp("OF8Fee")) <> 0 THEN 
		i = i +1
		sSQL = "INSERT INTO " & RegTransTableName
		sSQL = sSQL + " (MemberID, TourID, TransCode, Amount, TransDate, TransNo, OrderNo)"
		sSQL = sSQL + " VALUES ('"&sMemberID&"', '"&sTourID&"', 'OF8', '"&rsRegTemp("OF8Fee")&"', '"&NowDate&"', '"&i&"', '"&sOrderNo&"')"
		con.execute(sSQL)
	END IF

	IF cdbl(rsRegTemp("OF9Fee")) <> 0 THEN 
		i = i +1
		sSQL = "INSERT INTO " & RegTransTableName
		sSQL = sSQL + " (MemberID, TourID, TransCode, Amount, TransDate, TransNo, OrderNo)"
		sSQL = sSQL + " VALUES ('"&sMemberID&"', '"&sTourID&"', 'OF9', '"&rsRegTemp("OF9Fee")&"', '"&NowDate&"', '"&i&"', '"&sOrderNo&"')"
		con.execute(sSQL)
	END IF

	IF cdbl(rsRegTemp("OF10Fee")) <> 0 THEN 
		i = i +1
		sSQL = "INSERT INTO " & RegTransTableName
		sSQL = sSQL + " (MemberID, TourID, TransCode, Amount, TransDate, TransNo, OrderNo)"
		sSQL = sSQL + " VALUES ('"&sMemberID&"', '"&sTourID&"', 'OF10', '"&rsRegTemp("OF10Fee")&"', '"&NowDate&"', '"&i&"', '"&sOrderNo&"')"
		con.execute(sSQL)
	END IF

	closecon



END SUB





' -----------------------------
 SUB UpdatePaymentTransaction
' -----------------------------

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
				DateNow = Now
	
				' --- Look into setting messages on paypal failure --- 	
				Dim resp_message
				Dim sCheckNo, sLast4Card, sExpMonth, sExpYear, sApvl_Code, sCvv2_Resp, sAVS_Resp
				sCheckNo=""
				sLast4Card=""
				sExpMonth=""
				sExpYear=""
				sApvl_Code=""
				sCvv2_Resp=""
				sAVS_Resp=""			

				OpenCon
				sSQL = "UPDATE "&RegPaymentTableName
				sSQL = sSQL + " SET"
				sSQL = sSQL + " Result='"&sPaymentResult&"', PayStatus='"&sPayStatus&"'"
				IF Request("SpecialAction")="Y" THEN
						resp_message="Manual PayPal Acknowledgement"
						sSQL = sSQL + ", Message='"&resp_message&"'"
				END IF
					
				sSQL = sSQL + ", CheckNo='"&sCheckNo&"'"		
				sSQL = sSQL + ", Last4Card='"&sLast4Card&"', ExpYear='"&sLast4Card&"', ExpMonth='"&sExpMonth&"'"
				sSQL = sSQL + ", Apvl_Code='"&sApvl_Code&"', Cvv2_Resp='"&sCvv2_Resp&"', AVS_Resp='"&sAVS_Resp&"'"
			
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




' -----------------------------
   SUB DefineMemberVariables
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



'		response.write("<br>sTDateS = "&TRIM(sTDateS))
'		response.write("<br>sMemberID = "&TRIM(sMemberID))
		
' ---- Needs both Member and Tournament information to define sMembAge  ----
IF TRIM(sMemberID)<>"" AND TRIM(sTDateS)<>"" THEN
		sMembAge = AgeAtDate_New(sTDateS, sMemberID)		' Function finds Member Age
END IF
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





' --------------------------
  SUB SetSessionStatusText
' --------------------------


'Markdebug("SessStatText = nav="&nav&" - sMemberID = "&sMemberID&" - sTourID = "&sTourID&" ")


' --- Text used in various places of the FORM ---

' ----  Override of Fees has been set  -----
IF TRIM(sMoneyOverride) <> "" THEN
	Session("FeeStatusText")="Fee Override - "&sMoneyOverride
	Session("FeeStatusTextColor")="yellow"

' ----  Fees from RegPaymentTablName (Previous Charges) are less than current form values  -----
ELSEIF cdbl(sTotalFormFees) <> 0 AND sTotalPreviousPayments < cdbl(sTotalFormFees) THEN
	Session("FeeStatusText")="Balance Due"
	Session("FeeStatusTextColor")="red"

' ----  Fees from RegGenTable (Previous) are greater than current form values  -----
ELSEIF sTotalFormFees <> 0 AND sTotalPreviousPayments > cdbl(sTotalFormFees) THEN
	Session("FeeStatusText")="Refund Due"
	Session("FeeStatusTextColor")="red"

ELSEIF cdbl(sTotalFormFees) = sTotalPreviousPayments AND sPayStatus="C" THEN
	Session("FeeStatusText")="Paid In Full"
	Session("FeeStatusTextColor")="yellow"

' ---- Confirm has not been pressed and not at the end of process ---
ELSEIF (ExistingEntry(sMemberID)<>true AND sTotalFormFees = 0 AND sTotalPreviousPayments = cdbl(0)) AND nav<>7 THEN
	Session("FeeStatusText")="Not Entered"
	Session("FeeStatusTextColor")="red"

' ----------------------------------------------------------------------------------------------
' ---- *****  MARK - DO WE NEED A NEW CONDITION?  when FORM has never been confirmed and displaying original information 
ELSE
	Session("FeeStatusText")="Paid In Full"
	Session("FeeStatusTextColor")="yellow"
END IF 



END SUB







' ----------------------------
   SUB ZeroOutVariables
' ----------------------------

' --- Make values null to make sure any residual values are overridden


FOR EvtNo = 1 TO TotEv
	sSelectEvent(EvtNo) = ""
	sDiv(EvtNo)=""
	sQfyOverride(EvtNo) = ""
	sFeeClass(EvtNo) = ""
	IF sShowStd(EvtNo)=true THEN
		IF TRIM(sTEvent(EvtNo))<>"" THEN sFeeClass(EvtNo) = sTBaseClass  ' --- If event is offered assign BASE FeeClass ---
	ELSEIF sShowGR(EvtNo)=true THEN
		IF TRIM(sTEvent(EvtNo))<>"" THEN sFeeClass(EvtNo) = "G"  	' --- If event is offered assign GRASSROOTS FeeClass ---	
	ELSEIF sShowRec(EvtNo)=true THEN
		IF TRIM(sTEvent(EvtNo))<>"" THEN sFeeClass(EvtNo) = "R"  	' --- If event is offered assign RECORD FeeClass ---	
	END IF
NEXT



sAWSEFCheck = ""
sOfficial = ""
sClubMemb = ""
sClubCode = ""

sEntryFee = cdbl(0)
sLateFeeTot = cdbl(0)
sBanquetTot = cdbl(0)
sBanquetQty = cint(0)
sAWSEFDonation = cdbl(0)
sOffDiscAmt = cdbl(0)
sSrDiscAmt = cdbl(0)
sJrDiscAmt = cdbl(0)
sClubDiscAmt = cdbl(0)

sOF1Qty = cdbl(0)
sOF2Qty = cdbl(0)
sOF3Qty = cdbl(0)
sOF4Qty = cdbl(0)
sOF5Qty = cdbl(0)
sOF6Qty = cdbl(0)
sOF7Qty = cdbl(0)
sOF8Qty = cdbl(0)
sOF9Qty = cdbl(0)
sOF10Qty = cdbl(0)

sOF1Fee = cdbl(0)
sOF2Fee = cdbl(0)
sOF3Fee = cdbl(0)
sOF4Fee = cdbl(0)
sOF5Fee = cdbl(0)
sOF6Fee = cdbl(0)
sOF7ee = cdbl(0)
sOF8Fee = cdbl(0)
sOF9Fee = cdbl(0)
sOF10Fee = cdbl(0)

sTotalFormFees = cdbl(0)

TotEvents = 0
sBoat2 = ""
sRampHeight = ""


sMembOverride = ""
sRegionalOverride = ""
sMoneyOverride = ""

sFormError=""


'response.write("<br>IN Initialize = sLateFeeTot = "&sLateFeeTot)



END SUB





' ------------------------------
   SUB InitializeFromTable
' ------------------------------

' --- INITIALIZES EVENT ENTRY VALUES & CHECKBOXES from RegGenTableName or Zeros if EOF ---

' --- Make values null or zero to make sure any residual values are overridden ---
ZeroOutVariables


' --- Gets current event settings ---
set rsGen=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT TOP 1 * FROM " &WhichTable
sSQL = sSQL + " WHERE Left(TourID,6) = '" & SQLClean(left(sTourID,6)) & "' AND MemberID = '"&sMemberID&"'"

rsGen.open sSQL, SConnectionToTRATable


		
IF NOT rsGen.eof THEN 	
	
	sEntryType = rsGen("EntryType")

	sWaiverCode = rsGen("WaiverCode")
	sSignWaiver = rsGen("SignWaiver")
	Session("sRelease") = rsGen("WaiverCode")


	' --- Sets Session TEXT variables for display ---
	CheckWaiverStatus

	sMembRegDate = rsGen("RegisterDate")		
	sMembOverride = rsGen("MembOverride")
	sRegionalOverride = rsGen("RegionalOverride")
	sMoneyOverride = rsGen("MoneyOverride")
	sRampHeight = rsGen("RampHeight")

	sBanquetQty = rsGen("BanquetQty")
	sBanquetTot = Cdbl(rsGen("BanquetFee"))	
	sEntryFee = cdbl(rsGen("EntryFee"))
	sLateFeeTot = cdbl(rsGen("LateFee"))
	'sOtherFee = cdbl(rsGen("OtherFee"))
	sOtherFee = cdbl(0)

	sAWSEFDonation = cdbl(rsGen("AWSEFDonation"))
	sOffDiscAmt = cdbl(rsGen("OffDisc"))
	sSrDiscAmt = cdbl(rsGen("SrDisc"))
	sJrDiscAmt = cdbl(rsGen("JrDisc"))
	sClubDiscAmt = cdbl(rsGen("ClubDisc"))
	sPayStatus = rsGen("PayStatus")

	IF IsNull(rsGen("OF1Qty")) THEN sOF1Qty = cdbl(0) ELSE sOF1Qty = cdbl(rsGen("OF1Qty"))
	IF IsNull(rsGen("OF2Qty")) THEN sOF2Qty = cdbl(0) ELSE sOF2Qty = cdbl(rsGen("OF2Qty"))
	IF IsNull(rsGen("OF3Qty")) THEN sOF3Qty = cdbl(0) ELSE sOF3Qty = cdbl(rsGen("OF3Qty"))
	IF IsNull(rsGen("OF4Qty")) THEN sOF4Qty = cdbl(0) ELSE sOF4Qty = cdbl(rsGen("OF4Qty"))
	IF IsNull(rsGen("OF5Qty")) THEN sOF5Qty = cdbl(0) ELSE sOF5Qty = cdbl(rsGen("OF5Qty"))
	IF IsNull(rsGen("OF6Qty")) THEN sOF6Qty = cdbl(0) ELSE sOF6Qty = cdbl(rsGen("OF6Qty"))
	IF IsNull(rsGen("OF7Qty")) THEN sOF7Qty = cdbl(0) ELSE sOF7Qty = cdbl(rsGen("OF7Qty"))
	IF IsNull(rsGen("OF8Qty")) THEN sOF8Qty = cdbl(0) ELSE sOF8Qty = cdbl(rsGen("OF8Qty"))
	IF IsNull(rsGen("OF9Qty")) THEN sOF9Qty = cdbl(0) ELSE sOF9Qty = cdbl(rsGen("OF9Qty"))
	IF IsNull(rsGen("OF10Qty")) THEN sOF10Qty = cdbl(0) ELSE sOF10Qty = cdbl(rsGen("OF10Qty"))

	IF IsNull(rsGen("OF1Fee")) THEN sOF1Fee = cdbl(0) ELSE sOF1Fee = cdbl(rsGen("OF1Fee"))
	IF IsNull(rsGen("OF2Fee")) THEN sOF2Fee = cdbl(0) ELSE sOF2Fee = cdbl(rsGen("OF2Fee"))
	IF IsNull(rsGen("OF3Fee")) THEN sOF3Fee = cdbl(0) ELSE sOF3Fee = cdbl(rsGen("OF3Fee"))
	IF IsNull(rsGen("OF4Fee")) THEN sOF4Fee = cdbl(0) ELSE sOF4Fee = cdbl(rsGen("OF4Fee"))
	IF IsNull(rsGen("OF5Fee")) THEN sOF5Fee = cdbl(0) ELSE sOF5Fee = cdbl(rsGen("OF5Fee"))
	IF IsNull(rsGen("OF6Fee")) THEN sOF6Fee = cdbl(0) ELSE sOF6Fee = cdbl(rsGen("OF6Fee"))
	IF IsNull(rsGen("OF7Fee")) THEN sOF7Fee = cdbl(0) ELSE sOF7Fee = cdbl(rsGen("OF7Fee"))
	IF IsNull(rsGen("OF8Fee")) THEN sOF8Fee = cdbl(0) ELSE sOF8Fee = cdbl(rsGen("OF8Fee"))
	IF IsNull(rsGen("OF9Fee")) THEN sOF9Fee = cdbl(0) ELSE sOF9Fee = cdbl(rsGen("OF9Fee"))
	IF IsNull(rsGen("OF10Fee")) THEN sOF10Fee = cdbl(0) ELSE sOF10Fee = cdbl(rsGen("OF10Fee"))

'IF sMemberID="000001151" THEN 
'	sMembRegDate="03/01/2008"
'	'sMembRegDate="05/07/2008"
'END IF


	' -----------------------------------------------------------
	' --- Sets sLateDays ---
	' -----------------------------------------------------------



	IF sLateFeeTot > cdbl(0) AND sTLFPerDay=true THEN 	' --- Daily Late Fee ---	
			sLateDays = cint(sLateFeeTot/cdbl(sTLateFee))
	ELSEIF sLateFeeTot > cdbl(0) AND sTLFPerDay<>true THEN 	' --- One time late fee ---
			IF DateDiff("d", sTLateDate, sMembRegDate) < 21 THEN 
					sLateDays=DateDiff("d", sTLateDate_Adjusted, sMembRegDate)	' --- Provision for no sMembRegDate yet ---
			ELSE
					sLateDays = 1
			END IF
	END IF  


	' ---------------------------------------------------------------
	' --- AWSEF Officials and Club Discount to turn on checkboxes --- 
	' ---------------------------------------------------------------
	IF sAWSEFDonation > cdbl(0) THEN sAWSEFCheck = "on"

	IF sOffDiscAmt < 0 THEN sOfficial = "on"							

	IF sClubDiscAmt < cdbl(0) THEN 
			sClubMemb = "on"
			sClubCode = TRIM(sTourClubCode)
	END IF 	

	' ----------------------------------------------
	' --- Reads detail from event file - FIX ??? ---
	' ----------------------------------------------
	IF WhichTable=RegGenTableName THEN
			DetailTable=RegDetailTableName
	ELSEIF WhichTable=RegTempTableName THEN
			DetailTable=RegDetailTempTableName		
	END IF

	' -------------------------------	
	' --- Read event detail table ---
	' -------------------------------
	ReadFromRegisterEvents

	' -----------------------------------
	' --- Calculates Financial total  ---
	' -----------------------------------
	GetFinancialTotals		


ELSE 	' --- Initialize form if NOT found

	sWaiverCode = ""
	sEntryType = "IND"		
	sLateDays=0
	sMembRegDate=Date

	sFeeClass(1)="S"
	sFeeClass(2)="S"
	sFeeClass(3)="S"
	sFeeClass(4)="S"
	sFeeClass(5)="S"
	sFeeClass(6)="S"

	GetFinancialTotals
	sPayStatus="P"		

END IF 


END SUB



' ---------------------------
  SUB ReadFromRegisterEvents
' ---------------------------

' --- Reads in the EVENT DETAIL data from temporary or permanent table based on value of DetailTable ---

' --- Gets current event settings ---
SET rsDet=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT * FROM " &DetailTable
sSQL = sSQL + " WHERE Left(TourID,6) = '"&left(sTourID,6)&"' AND MemberID = '"&sMemberID&"'"
sSQL = sSQL + " ORDER BY"



' --- OLD CODE for reference ---
' --- Sequence must match the way the events are set up in RegistrationEventsOffered tools_include.asp  ---
'SELECT CASE sTSptsGrpID
'   CASE "AWS"
'	sSQL = sSQL + " CASE WHEN Event='S' THEN '1' WHEN Event='T' THEN '2' WHEN Event='J' THEN '3'"
'	sSQL = sSQL + " WHEN Event='KB' THEN '4' WHEN Event='DA' THEN '5' WHEN Event='WB' THEN '6'"
'	sSQL = sSQL + " WHEN Event='BF' THEN '7' WHEN Event='HF' THEN '8' ELSE '9' END" 
'  CASE "USW"
'	sSQL = sSQL + " CASE WHEN Event='WB' THEN '1' WHEN Event='WS' THEN '2' WHEN Event='WU' THEN '3'"
'	sSQL = sSQL + " ELSE '4' END" 
'
'   CASE "AKA"
'	sSQL = sSQL + " CASE WHEN Event='KS' THEN '1' WHEN Event='KT' THEN '2' WHEN Event='KF' THEN '3'"
'	sSQL = sSQL + " WHEN Event='KR' THEN '4' ELSE '5' END" 
'END SELECT


' --- Sequence must match the way the events are set up in RegistrationEventsOffered tools_include.asp  ---
' --- Revised code for 2009 which incorporates GR and is independent of SptsGrpID ---
sSQL = sSQL + " CASE WHEN Event='S' THEN '1' WHEN Event='T' THEN '2' WHEN Event='J' THEN '3' WHEN Event='3G' THEN '4'" 
sSQL = sSQL + " WHEN Event='WB' THEN '5' WHEN Event='WS' THEN '6' WHEN Event='WU' THEN '7' WHEN Event='WJ' THEN '8'"
sSQL = sSQL + " WHEN Event='KS' THEN '9' WHEN Event='KT' THEN '10' WHEN Event='KF' THEN '11' WHEN Event='KR' THEN '12'"
sSQL = sSQL + " WHEN Event='HF' THEN '13' WHEN Event='HJ' THEN '14' WHEN Event='HB' THEN '15' WHEN Event='H3' THEN '16'"
sSQL = sSQL + " WHEN Event='DA' THEN '17' WHEN Event='BF' THEN '18' ELSE '19' END" 





rsDet.open sSQL, SConnectionToTRATable

IF NOT rsDet.eof THEN
   rsDet.movefirst

   DO WHILE NOT rsDet.eof		' --- Loop through all records in table ---

	EvtNo=1
	FOR EvtNo=1 TO TotEv 		' --- Compare record value to sTEvent(EvtNo) for this tournament ---

		IF TRIM(rsDet("Event"))=TRIM(sTEvent(EvtNo)) THEN		

			sDiv(EvtNo)=TRIM(rsDet("Div"))
			sQfyOverride(EvtNo)=TRIM(rsDet("QfyOverride"))
			sFeeClass(EvtNo)=TRIM(rsDet("FeeClass"))
			sFeeRounds(EvtNo)=rsDet("FeeRounds")
			sBoat(EvtNo)=TRIM(rsDet("Boat"))
			sSkill(EvtNo)=TRIM(rsDet("Skill"))
			sSelectEvent(EvtNo) = "on"		
			TotEvents = TotEvents + 1 
		END IF
	NEXT

	rsDet.movenext
  LOOP

END IF



END SUB




' ----------------------
  SUB ReadFromTransTable
' ----------------------

	' --- Reads data from the (accounting) transaction table ---

	Dim mdate

	' ---  Query the latest group of TRANSACTION records from RegTransTableName based on time/date ---
	SET rsRegTrans=Server.CreateObject("ADODB.recordset")
	sSQL = "(SELECT MAX(TransDate) AS maxdate FROM "&RegTransTableName
	sSQL = sSQL + " WHERE Left(TourID,6) = '" & SQLClean(left(sTourID,6)) & "' AND MemberID = '"&sMemberID&"' AND OrderNo>'1999') AS d"
	rsRegTrans.open sSQL, SConnectionToTRATable, 3, 3

	IF NOT rsRegTrans.eof THEN
		mDate = rsRegTrans("maxdate")     ' Defines the LAST date/time variable mDate

		' ---   Reads all transactions matching the LATEST date/time  ---
		SET rsRegTrans=Server.CreateObject("ADODB.recordset")
		sSQL = "SELECT MemberID, TourID, TransCode, TransDate, Amount AS Amount, TransNo FROM "&RegTransTableName
		sSQL = sSQL + " WHERE TransDate = '"&mDate&"' AND Left(TourID,6) = '"&left(sTourID,6)&"' AND MemberID = '"&sMemberID&"' AND OrderNo>'1999'"
		rsRegTrans.open sSQL, SConnectionToTRATable, 3, 3
	END IF


	' -------------------------------------------------------------------------------------------------------
	' Sets checkboxes and Previous Charges and DISCOUNTS based on RegTransTableName - TRANSACTION table  ----  
	' -------------------------------------------------------------------------------------------------------

	IF NOT rsRegTrans.eof THEN

		rsRegTrans.movefirst
		DO WHILE NOT rsRegTrans.eof
				
			SELECT CASE TRIM(rsRegTrans("Transcode"))
				CASE "FEF"
					sEntryFeeTrans = cdbl(rsRegTrans("Amount"))
				CASE "FLF"
					sLateFeeTotTrans = cdbl(rsRegTrans("Amount"))
				CASE "BAN"
					sBanquetTotTrans = cdbl(rsRegTrans("Amount"))
				CASE "OBF"
					sAWSEFCheckTrans = "on"
					sAWSEFDonationTrans = cdbl(rsRegTrans("Amount"))
				CASE "DOF"
					sOfficial = "on"							
					sOffDiscAmtTrans = cdbl(rsRegTrans("Amount"))
				CASE "DSR"
					sSrDiscAmtTrans = cdbl(rsRegTrans("Amount"))
				CASE "DJR"
					sJrDiscAmtTrans = cdbl(rsRegTrans("Amount"))
				CASE "DCL"
					sClubMemb = "on"
					sClubCode = TRIM(sTourClubCode) 	
					sClubDiscAmtTrans = cdbl(rsRegTrans("Amount"))
				CASE "OF1"
					sOF1Trans = cdbl(rsRegTrans("Amount"))
				CASE "OF2"
					sOF2Trans = cdbl(rsRegTrans("Amount"))
				CASE "OF3"
					sOF3Trans = cdbl(rsRegTrans("Amount"))
				CASE "OF4"
					sOF4Trans = cdbl(rsRegTrans("Amount"))
				CASE "OF5"
					sOF5Trans = cdbl(rsRegTrans("Amount"))
				CASE "OF6"
					sOF6Trans = cdbl(rsRegTrans("Amount"))
				CASE "OF7"
					sOF7Trans = cdbl(rsRegTrans("Amount"))
				CASE "OF8"
					sOF8Trans = cdbl(rsRegTrans("Amount"))
				CASE "OF9"
					sOF9Trans = cdbl(rsRegTrans("Amount"))
				CASE "OF10"
					sOF10Trans = cdbl(rsRegTrans("Amount"))
			END SELECT
			rsRegTrans.movenext
		LOOP
	END IF




END SUB






' --------------------------
   SUB ReadEntryFormValues
' --------------------------

' --- Reads the values once the form has been submitted ---

ZeroOutVariables


sRampHeight = TRIM(Request("sRampHeight"))
sEntryType = TRIM(Request("sEntryType"))

sOfficial = TRIM(Request("fOfficial"))
sAWSEFCheck = TRIM(Request("fAWSEFCheck"))
sClubMemb = TRIM(Request("fClubMemb"))
sClubCode = TRIM(Request("fClubCode"))

sMembOverride = TRIM(Request("sMembOverride"))
sRegionalOverride = TRIM(Request("sRegionalOverride"))
sMoneyOverride = TRIM(Request("sMoneyOverride"))
IF Request("sBanquetQty")="" THEN sBanquetQty = CInt(0) ELSE sBanquetQty = cint(Request("sBanquetQty"))

IF Request("sOF1Qty")="" THEN sOF1Qty = CInt(0) ELSE sOF1Qty = cint(Request("sOF1Qty"))
sOF2Qty = cint(Request("sOF2Qty"))
sOF3Qty = cint(Request("sOF3Qty"))
sOF4Qty = cint(Request("sOF4Qty"))
sOF5Qty = cint(Request("sOF5Qty"))
sOF6Qty = cint(Request("sOF6Qty"))
sOF7Qty = cint(Request("sOF7Qty"))
sOF8Qty = cint(Request("sOF8Qty"))
sOF9Qty = cint(Request("sOF9Qty"))
sOF10Qty = cint(Request("sOF10Qty"))


sWaiverCode = Request("sWaiverCode")
sSignWaiver = Request("sSignWaiver")

' Logic for validating date numbers
'IF (isnumeric(left(sMembRegDate,2)) And isnumeric(right(left(sMembRegDate,5),2)) And isnumeric(right(sMembRegDate,4)) And right(left(sMembRegDate,3),1) = "/" And right(left(sMembRegDate,6),1) = "/" And isDate(sMembRegDate)) THEN

IF Request("sMembRegDate") <> "" THEN


		sMembRegDate = sqlclean(Request("sMembRegDate"))
		'IF sMemberID="000001151" THEN
		'		response.write("<br><br>In ReadFormVar")
		'		response.write("<br>sMembRegDate = "&sMembRegDate)
		'		response.write("<br>")
		'END IF
ELSE
		sMembRegDate = DATE
END IF	 


TotEvents = 0

FOR EvtNo = 1 TO TotEv
	fSelEvt="fSelectEvent"&EvtNo

	IF TRIM(Request("fSelectEvent"&EvtNo)) = "on" THEN 
			TotEvents = TotEvents + 1
			sSelectEvent(EvtNo) = TRIM(Request("fSelectEvent"&EvtNo))
			sDiv(EvtNo) = TRIM(Request("fDiv"&EvtNo))
			sQfyOverride(EvtNo) = TRIM(Request("fQfyOverride"&EvtNo))
			sFeeClass(EvtNo) = TRIM(Request("fFeeClass"&EvtNo))
			sFeeRounds(EvtNo) = TRIM(Request("fFeeRounds"&EvtNo))
			sSkill(EvtNo) = TRIM(Request("fSkill"&EvtNo))	

			IF sTPandC<>true THEN sFeeRounds(EvtNo)=1
		
			' --- Added 7-13-2013 to deal with Grassroots not being offered as PandC ---
			IF sTPandC=true AND sFeeClass(EvtNo)="G" THEN
					sFeeRounds(EvtNo)=1
			END IF
	END IF
	sBoat(EvtNo) = TRIM(Request("fBoat"&EvtNo))
NEXT


CalculateEntryFees


END SUB









' -----------------------
  SUB CalculateEntryFees
' -----------------------

' --- Calculates the Entry Fees ---

sEntryFee=cdbl(0)

' --- Establish total number of Pulls for each FeeClass ---
' --- For Traditional Format tournaments (NOT PandC) the Pulls are actually Events in each FeeClass --- 
Dim sTotGPulls, sTotSPulls, sTotRPulls, sTotAllPulls
sTotGPulls=cdbl(0)
sTotSPulls=cdbl(0)
sTotRPulls=cdbl(0)
sTotAllPulls=cdbl(0)

FOR EvtNo=1 TO TotEv
	IF sFeeClass(EvtNo)="G" THEN sTotGPulls = sTotGPulls + sFeeRounds(EvtNo)
	IF sFeeClass(EvtNo)="S" THEN sTotSPulls = sTotSPulls + sFeeRounds(EvtNo)
	IF sFeeClass(EvtNo)="R" THEN sTotRPulls = sTotRPulls + sFeeRounds(EvtNo)
NEXT


IF sMemberID="000001151"  AND TestMode="yes" THEN
	response.write("<br>sEntryType = "&sEntryType)
END IF



IF sEntryType="FAM" THEN


	sEntryFee=sTEntryFeeFamily

	' --- Run the test for a previous Family Entry for this tournament in tools_registration.asp ---
	IF ExistingEntry(sMemberID) THEN
		sEntryFee=0

		' --- Tests to see if the total registered is greater than the limit under family type ---
		IF cdbl(Session("TotRegisteredFamMembers"))>=sMaxFamMembers AND cdbl(Session("TotRegisteredFamMembers"))<>cdbl(0) THEN
			sEntryFee=sTEntryFeeFamExtra
		END IF

	END IF


  IF sMemberID="000001151"  AND TestMode="yes" THEN
	response.write("<br>MemberID = 000001151")
	response.write("<br>TourID = "&sTourID)

	response.write("<br>Total Family Members = "&cdbl(Session("TotRegisteredFamMembers"))&"<br>")
	response.write("<br>Existing sMemberID = ")
	response.write(ExistingEntry(sMemberID))
	response.write("<br>Session Test = ")
	response.write(cdbl(Session("TotRegisteredFamMembers"))>=sMaxFamMembers AND cdbl(Session("TotRegisteredFamMembers"))<>cdbl(0))
	response.write("<br>sEntryFee = "&sEntryFee)
	'response.end
  END IF



ELSEIF sTPandC=true THEN

		Dim ApplySurchargeToAllRounds
		IF LEFT(sTourID,6)="16C017" THEN		' --- Apply RSurcharge to every round
				ApplySurchargeToAllRounds="Y"
				' response.write("<br>Line2977")
		ELSE	
				ApplySurchargeToAllRounds="N"
		END IF			


		' --- Sets BASE Entry to be charged for first PULL based on HIGHEST FeeClass entered ---
		' --- Modified 3-23-2009
		IF sTotGPulls>0 AND sGREntryFee1>0 THEN 
				sEntryFee = sGREntryFee1
		ELSEIF sTotGPulls>0 AND sGREntryFee1=0 THEN
				sEntryFee = sTEntryFee1
		END IF


		' --- Find the Combined number of Standard and R-Surcharge pulls ---
		' --- If there are 2 pulls, then the first two fees apply (NOTE: R surcharge already added above) ---   
		' --- If there are 3 or more pulls, then the 3rd Fee is used for all additional pulls ---
		sTotSandR = sTotSPulls+sTotRPulls


		' --- Changed 8-4-2015 ---
		SELECT CASE sTotSandR
				CASE 1
								' sEntryFee = (sTEntryFee1+sRSurchg)
								IF sTotRPulls>0 THEN 
										sEntryFee = sTEntryFee1+sRSurchg
										'response.write("<br>Pos AA = "&sEntryFee)
								ELSE
										sEntryFee = sTEntryFee1		
										'response.write("<br>Pos AB = "&sEntryFee)
								END IF
				CASE 2																		' --- BASE entry plus one Addl Pull --- 
								IF ApplySurchargeToAllRounds="Y" THEN
										IF sTotRPulls=2 THEN  		' Both rounds are R
												sEntryFee = (sTEntryFee1+sRSurchg) + (sTEntryFee2+sRSurchg)
												'response.write("<br>Pos A = "&sEntryFee)
										ELSEIF sTotRPulls=1 THEN  ' One round R and other S 		
												sEntryFee = (sTEntryFee1+sRSurchg) + sTEntryFee2
												'response.write("<br>Pos B = "&sEntryFee)
										ELSE 											' Both rounds S
												sEntryFee = sTEntryFee1 + sTEntryFee2
												'response.write("<br>Pos C = "&sEntryFee)
										END IF

								ELSE 
										IF sTotRPulls>=1 THEN  		' At least one round is R
												sEntryFee = sTEntryFee1 + sTEntryFee2 + sRSurchg
										ELSE 											' Both rounds S
												sEntryFee = sTEntryFee1 + sTEntryFee2
										END IF
								END IF

								' sEntryFee = sEntryFee + sTEntryFee2

				CASE 3, 4, 5, 6, 7, 8, 9, 10, 11, 12			' --- BASE entry plus remaining Addl Pulls --- 
								'response.write("<br>IN CASE 3 - sTotRPulls = "&sTotRPulls)
								'response.write("<br>sTotSPulls = "&sTotSPulls)
								'response.write("<br>sTotSandR = "&sTotSandR)
								IF ApplySurchargeToAllRounds="Y" THEN
										IF sTotRPulls>=3 THEN  		' All rounds are R
												sEntryFee = (sTEntryFee1+sRSurchg) + (sTEntryFee2+sRSurchg) + ( (sTotRPulls-2) * (sTEntryFee3+sRSurchg) ) + ( sTotSPulls * sTEntryFee3 )
												'response.write("<br>Pos 1 = "&sEntryFee)
										ELSEIF sTotRPulls=2 THEN  ' Two rounds R and other S 		
												sEntryFee = (sTEntryFee1+sRSurchg) + (sTEntryFee2+sRSurchg) + ( (sTotSandR-2) * sTEntryFee3 )
												'response.write("<br>Pos 2 = "&sEntryFee)
										ELSEIF sTotRPulls=1 THEN  ' One round R and other S 		
												sEntryFee = (sTEntryFee1+sRSurchg) + sTEntryFee2 + ( (sTotSandR-2) * sTEntryFee3 )
												'response.write("<br>Pos 3 = "&sEntryFee)
										ELSE 											' All rounds S
												sEntryFee = sTEntryFee1 + sTEntryFee2 + ( (sTotSandR-2) * sTEntryFee3 )
												'response.write("<br>Pos 4 = "&sEntryFee)
										END IF

										' sEntryFee = (sTEntryFee1+sRSurchg) + (sTEntryFee2+sRSurchg) + ( (sTotSandR-2) * (sTEntryFee3+sRSurchg) )
								ELSE
										IF sTotRPulls>=1 THEN  		' All rounds are R
												sEntryFee = (sTEntryFee1+sRSurchg) + sTEntryFee2 + ( (sTotSandR-2) * sTEntryFee3 )
										ELSE 											' All rounds are S
												sEntryFee = sTEntryFee1 + sTEntryFee2 + ( (sTotSandR-2) * sTEntryFee3 )
										END IF

										' sEntryFee = sEntryFee1 + sTEntryFee2 + ( (sTotSandR-2) * sTEntryFee3 ) + sRSurchg
								END IF

								' sEntryFee = sEntryFee + sTEntryFee2 + ( (sTotSandR-2) * sTEntryFee3 )
			END SELECT


			' --- If the total SandR are less than 3, then check for Grassroots pulls ---
			' --- If GR pulls are greater than zero add as appropriate ---
			' --- Deals wit up to 12 pulls in tournament between all events ---
			IF sTotSandR=2 AND sTotGPulls>0 THEN
					sEntryFee = sEntryFee + (sGREntryFee3 * sTotGPulls)
					
			ELSEIF sTotSandR=1 AND sTotGPulls>=1 THEN		' --- Count Addl PULLS from Grassroots pricing if not at 3 Pulls or more ---
					sEntryFee = sEntryFee + sGREntryFee2 + (sGREntryFee3 * (sTotGPulls-1))
					
			ELSEIF sTotSandR=0 AND sTotGPulls>=2 THEN
					sEntryFee = sGREntryFee1 + sGREntryFee2 + (sGREntryFee3 * (sTotGPulls-2))
					
			END IF

	


ELSE	' --- Traditional tournaments Individual Entry Type

	' --- This deals with the fact that SWIFT does not require the 2nd and 3rd fields to be set and they default to zero ---
	IF sTEntryFee2=0 THEN sTEntrYFee2=sTEntryFee1
	IF sTEntryFee3=0 THEN sTEntrYFee3=sTEntryFee1
	' --- Added 3-22-2009 to account for when the GR fee and regular fee are the same ----
	IF sGREntryFee1=0 THEN sGREntrYFee1=sTEntryFee1

	IF sGREntryFee2=0 THEN sGREntrYFee2=sGREntryFee1
	IF sGREntryFee3=0 THEN sGREntrYFee3=sGREntryFee1


	' --- Sets BASE Entry to be charged for first EVENT based on HIGHEST FeeClass entered ---
	IF sTotGPulls>0 THEN sEntryFee = sGREntryFee1
	IF sTotSPulls>0 THEN sEntryFee = sTEntryFee1
	IF sTotRPulls>0 THEN sEntryFee = sTEntryFee1+sRSurchg		


	' --- Find the Combined number of Standard and R-Surcharge EVENTS (Note if not PandC, then each FeeRound is 1) ---
	' --- If there are 2 EVENTS, then the first two fees apply (NOTE: R surcharge already added above) ---   
	' --- If there are 3 or more EVENTS, then the 3rd Fee is used for all additional pulls ---
	sTotSandR = sTotSPulls+sTotRPulls
	SELECT CASE sTotSandR
	   CASE 2						' --- Change BASE to be 2 EVENT Fee --- 
				sEntryFee = sEntryFee-sTEntryFee1+sTEntryFee2
	   CASE 3, 4, 5, 6, 7, 8, 9, 10				' --- Change BASE to be 3 EVENT Fee --- 
				sEntryFee = sEntryFee - sTEntryFee1 + sTEntryFee3
	END SELECT


	' --- If the total SandR are less than 3, then check for Grassroots pulls ---
	' --- Change BASE to be 3 EVENT Fee - NOTE: Use EntryFee3 for 3 Events (vs GR fields) because there is at least one SandR event ---
	IF sTotSandR=2 AND sTotGPulls>0 AND sTEntryFee3 >= sTEntryFee2 THEN			 
			sEntryFee = sEntryFee - sTEntryFee2 + sTEntryFee3
	ELSEIF sTotSandR=1 AND sTotGPulls>=1 THEN					' --- Count Addl PULLS from Grassroots pricing if not at 3 Pulls or more ---
			SELECT CASE sTotGPulls
		 	  CASE 1
						sEntryFee = sEntryFee - sTEntryFee1 + sTEntryFee2
				CASE 2, 3, 4, 5, 6, 7, 8, 9, 10					' --- sTotSandR=1, plus TWO Addl GR Pulls --- 
						sEntryFee = sEntryFee - sTEntryFee1 + sTEntryFee3
			END SELECT

			' --- Since there are no SandR the entry is just the appropriate GR fee ---
			' --- GR Event #1 is set as BASE ---
			ELSEIF sTotSandR=0 AND sTotGPulls=2 THEN
					sEntryFee = sGREntryFee2
			ELSEIF sTotSandR=0 AND sTotGPulls>2 THEN
					sEntryFee = sGREntryFee3
			END IF



	' ++++ NOT SURE THE FOLLOWING IS NEEDED +++++

	TotEvents=0
	FOR EvtNo=1 TO TotEv
		IF TRIM(sDiv(EvtNo))<>"" THEN TotEvents=TotEvents+1
	NEXT 
END IF



' --- sEntryFee is the TOTAL entry fees for these settings ---
sEntryFee = Cdbl(sEntryFee)

' --- Reads RegPaymentTableName to see what fees were actually paid ---
DetermineTotalFeesActuallyPaid

' --- Determines discounts and total form fees ---
RecalcFormValues


END SUB





' -------------------------
   SUB RecalcFormValues
' -------------------------


'response.write("<br>Top of RecalcForm = sLateFeeTot = "&sLateFeeTot)
'response.write("<br>sTLateDate = "&sTLateDate)

' ------------------
' ---  Late Fee ----
' ------------------

' --- Test for Mark ---
'IF sMemberID="000001151" THEN
'		sMembRegDate="7/18/2012 03:01:00 AM"
'END IF

sLateHours_Adjusted = DateDiff("h", sTLateDate_Adjusted, sMembRegDate)

' --- sTLateDate from Sanctions system - sMemRegDate from ???
IF DateDiff("h", sTLateDate_Adjusted, sMembRegDate)>0 THEN

		sLateDays = DateDiff("d", sTLateDate, sMembRegDate)
		' --- If the hours are still not late and it is the first LateDate then adjust for timezone ---
		IF sLateHours_Adjusted<0 THEN 
				sLateDays_Adjusted = 0
		ELSE
				sLateDays_Adjusted=sLateDays
		END IF
ELSE
		sLateDays = 0
		sLateDays_Adjusted=0
END IF

'IF sMemberID="000001151" THEN 
'		response.write("<br><br>REG 2863")
'		response.write("<br>sMembRegDate = "&sMembRegDate)
'		response.write("<br>sTLateDate = "&sTLateDate)
'		response.write("<br>sTLateDate_Adjusted = "&sTLateDate_Adjusted)
'		response.write("<br>sLateDays = "&sLateDays)
'		response.write("<br>sLateDays_Adjusted = "&sLateDays_Adjusted)
'END IF




'IF cint(sLateDays) > 0 AND sEntryFee > 0 THEN
IF cint(sLateDays_Adjusted) > 0 AND sEntryFee > 0 THEN
	
		IF sTLFPerDay=true AND sTLateFee>cdbl(0.00) THEN 	' --- Daily Late Fee ---
				' sLateFeeTot = sLateDays * Cdbl(sTLateFee)
				sLateFeeTot = sLateDays_Adjusted * Cdbl(sTLateFee)
		ELSEIF sTLFPerDay<>true AND sTLateFee>0.00 THEN  	' --- One time Late Fee ---
				sLateFeeTot = Cdbl(sTLateFee)
		ELSE  
				sLateFeeTot = Cdbl(0.00)
		END IF  
ELSE
		sLateFeeTot = Cdbl(0.00)
END IF


' ---- Banquet Tickets and Fees         -----
' -------------------------------------------

IF sBanquetQty > cint(0) AND sBTickWithE=false THEN
	sBanquetTot = sBanquetQty * sBTickCost
ELSEIF sBanquetQty > cint(1) AND sBTickWithE=true THEN
	sBanquetTot = (sBanquetQty-cdbl(1)) * sBTickCost
ELSE
	sBanquetTot = cdbl(0)
END IF


' ---- Donation to AWSEF Building Fund  -----
' -------------------------------------------

ShowAWSEFDonation=true
IF sAWSEFCheck = "on" THEN
	sAWSEFDonation = Cdbl(10.00)
ELSE
	sAWSEFDonation = Cdbl(0.00)
END IF


' ---- Discount to Junior B/G 1-3 per Tour_Manager.asp   -----
' ------------------------------------------------------------

sJrDiscAmt = Cdbl(0)

'IF sMemberID="000001151" THEN 
'	sMembAge=16
'	sJrDiscPerc=40
'END IF

' --- A positive value in sJrDiscPerc indicates the discount is a percentage
IF sJrDiscPerc > 0 THEN
	IF sMembAge < 18 AND sEntryFee > 0 THEN  
		sJrDiscAmt = - (Cdbl(sEntryFee) * Cdbl(sJrDiscPerc))/100
	END IF

' --- A negative value in sJrDiscPerc indicates the discount is in $$
ELSEIF sJrDiscPerc < 0 THEN
	IF sMembAge < 18 AND sEntryFee > 0 THEN  
		sJrDiscAmt = CDbl(sJrDiscPerc)
	END IF
END IF 	


' ---- Discount to divisions M/W-6  ----
' --------------------------------------

sSrDiscAmt = Cdbl(0)

'IF sMemberID="000001151" THEN 
'	sMembAge=80
'	sSrDiscPerc=40
'END IF

' --- A positive value in sSrDiscPerc indicates the discount is a percentage
IF sSrDiscPerc > 0 THEN
	IF sMembAge > 59 AND sEntryFee > 0 THEN
		sSrDiscAmt = - (Cdbl(sEntryFee) * Cdbl(sSrDiscPerc))/100
	END IF
' --- A negative value in sJrDiscPerc indicates the discount is in $$
ELSEIF sSrDiscPerc < 0 THEN
	IF sMembAge > 59 AND sEntryFee > 0 THEN  
		sSrDiscAmt = CDbl(sSrDiscPerc)
	END IF

END IF


' ---- Calculate Officials Discount ----
' --------------------------------------

sOffDiscAmt = Cdbl(0)

'IF sMemberID="000001151" THEN 
'	sOffDiscPerc=10
'END IF

' --- A positive value in sOffDiscPerc indicates the discount is a percentage
IF sOffDiscPerc > 0 THEN
	IF sOfficial = "on" AND sEntryFee > 0 THEN
		sOffDiscAmt = - (Cdbl(sEntryFee) * CDbl(sOffDiscPerc))/100

		' --- If discount method is CUMM and the discounts total more than the entry fee, then make discount equal to what is left after Jr/Sr
		IF sDiscMeth <> 1 AND (cdbl(sEntryFee) + cdbl(sOffDiscAmt) + cdbl(sSrDiscAmt) + cdbl(sJrDiscAmt)) <= 0 THEN
			sOffDiscAmt = - (cdbl(sEntryFee) + cdbl(sSrDiscAmt) + cdbl(sJrDiscAmt))
		END IF
	END IF
' --- A negative value in sOffDiscPerc indicates the discount is in $$
ELSEIF sOffDiscPerc < 0 THEN
	IF sOfficial = "on" AND sEntryFee > 0 THEN
		sOffDiscAmt = CDbl(sOffDiscPerc)

		' --- If discount method is CUMM and the discounts total more than the entry fee, then make the discount equal to what is left after Jr/Sr
		IF sDiscMeth <> 1 AND (cdbl(sEntryFee) + cdbl(sOffDiscAmt) + cdbl(sSrDiscAmt) + cdbl(sJrDiscAmt)) <= 0 THEN
				sOffDiscAmt = - (cdbl(sEntryFee) + cdbl(sSrDiscAmt) + cdbl(sJrDiscAmt))
		END IF
	END IF
END IF

'response.write("<br>Off2="&sOffDiscAmt)
'response.write("<br>OffPerc="&sOffDiscPerc)
'response.write("<br>sClubDiscPerc="&sClubDiscPerc)



' ---------- Discount to CLUB MEMBERS if match to ClubCode ----
' -------------------------------------------------------------

sClubDiscAmt = Cdbl(0)

'IF sMemberID="000001151" THEN 
'	sClubDiscPerc=10
'	sClubCode="X"
'	sTourClubCode="X"
'END IF

' --- A positive value in sClubDiscPerc indicates the discount is a percentage
IF sClubDiscPerc > 0 AND sClubMemb = "on" AND sEntryFee > 0 AND TRIM(sClubCode) = TRIM(sTourClubCode) THEN
	IF TRIM(sClubCode) <> "" AND TRIM(sClubCode)=TRIM(sTourClubCode) THEN
		sClubDiscAmt = - (Cdbl(sEntryFee) * CDbl(sClubDiscPerc))/100

		' --- If disc method is CUMM and Entry Fee less discounts is less than zero --- 
		'IF (cdbl(sEntryFee) + cdbl(sClubDiscAmt) + cdbl(sOffDiscAmt) + cdbl(sSrDiscAmt) + cdbl(sJrDiscAmt)) <= 0 THEN
		IF sDiscMeth <> 1 AND (cdbl(sEntryFee) + cdbl(sClubDiscAmt) + cdbl(sOffDiscAmt) + cdbl(sSrDiscAmt) + cdbl(sJrDiscAmt)) <= 0 THEN
			' --- Make the discount equal what is left ---
			sClubDiscAmt = - (cdbl(sEntryFee) + cdbl(sOffDiscAmt) + cdbl(sSrDiscAmt) + cdbl(sJrDiscAmt))
		END IF
	END IF

' --- A negative value in sClubDiscPerc indicates the discount is in $$
ELSEIF sClubDiscPerc < 0 AND sClubMemb = "on" AND sEntryFee > 0 AND TRIM(sClubCode) = TRIM(sTourClubCode) THEN
	IF TRIM(sClubCode) <> "" AND TRIM(sClubCode)=TRIM(sTourClubCode) THEN
		sClubDiscAmt = CDbl(sClubDiscPerc)

		' --- If disc method is CUMM and the Entry Fee less discounts is less than zero --- 
		'IF (cdbl(sEntryFee) + cdbl(sClubDiscAmt) + cdbl(sOffDiscAmt) + cdbl(sSrDiscAmt) + cdbl(sJrDiscAmt)) <= 0 THEN
		IF sDiscMetho<>1 AND (cdbl(sEntryFee) + cdbl(sClubDiscAmt) + cdbl(sOffDiscAmt) + cdbl(sSrDiscAmt) + cdbl(sJrDiscAmt)) <= 0 THEN
			' --- Make the discount equal what is left ---
			sClubDiscAmt = - (cdbl(sEntryFee) + cdbl(sOffDiscAmt) + cdbl(sSrDiscAmt) + cdbl(sJrDiscAmt))
		END IF
	END IF


END IF	



' --- OPTIONAL ADDITIONAL FEES ---
' --------------------------------

IF sOF1Qty > 0 THEN
	sOF1Fee = sOF1Qty * sOF1Amt
ELSE
	sOF1Fee = cdbl(0)
END IF
IF sOF2Qty > 0 THEN
	sOF2Fee = sOF2Qty * sOF2Amt
ELSE
	sOF2Fee = cdbl(0)
END IF
IF sOF3Qty > 0 THEN
	sOF3Fee = sOF3Qty * sOF3Amt
ELSE
	sOF3Fee = cdbl(0)
END IF
IF sOF4Qty > 0 THEN
	sOF4Fee = sOF4Qty * sOF4Amt
ELSE
	sOF4Fee = cdbl(0)
END IF
IF sOF5Qty > 0 THEN
	sOF5Fee = sOF5Qty * sOF5Amt
ELSE
	sOF5Fee = cdbl(0)
END IF
IF sOF6Qty > 0 THEN
	sOF6Fee = sOF6Qty * sOF6Amt
ELSE
	sOF6Fee = cdbl(0)
END IF
IF sOF7Qty > 0 THEN
	sOF7Fee = sOF7Qty * sOF7Amt
ELSE
	sOF7Fee = cdbl(0)
END IF
IF sOF8Qty > 0 THEN
	sOF8Fee = sOF8Qty * sOF8Amt
ELSE
	sOF8Fee = cdbl(0)
END IF
IF sOF9Qty > 0 THEN
	sOF9Fee = sOF9Qty * sOF9Amt
ELSE
	sOF9Fee = cdbl(0)
END IF
IF sOF10Qty > 0 THEN
	sOF10Fee = sOF10Qty * sOF10Amt
ELSE
	sOF10Fee = cdbl(0)
END IF






' --- Sets total form discount and form total ---
GetFinancialTotals


END SUB






' --------------------------
  SUB GetFinancialTotals
' --------------------------


' Response.write("<br>sDiscMeth="&sDiscMeth)

' ----  Sets total form discount based on Discount Method (Cum=0 or Max=1)  ----
IF sDiscMeth = 0 THEN
	' --- Discounts have been adjusted to make sure total (CUMM) does not exceed entry fee
	ActualDisc = sJrDiscAmt + sSrDiscAmt + sOffDiscAmt + sClubDiscAmt
ELSE		


	' --- Find the maximum discount ---
	IF sJrDiscAmt < 0 THEN
		ActualDisc = sJrDiscAmt
	ELSE
		ActualDisc = sSrDiscAmt
	END IF 	

	IF sOffDiscAmt <= ActualDisc THEN
		ActualDisc = sOffDiscAmt
	END IF

	IF sClubDiscAmt <= ActualDisc THEN
		ActualDisc = sClubDiscAmt
	END IF

	'Response.write("<br><br>Pos 1")	
	'Response.write("<br>sJrDiscAmt = "&sJrDiscAmt)
	'Response.write("<br>sSrDiscAmt = "&sSrDiscAmt)
	'Response.write("<br>sOffDiscAmt = "&sOffDiscAmt)
	'Response.write("<br>sClubDiscAmt = "&sClubDiscAmt)	
	'Response.write("<br>ActualDisc = "&ActualDisc)	


	' --- Now reset all other discounts not used to zero ---
	IF ActualDisc=sJrDiscAmt THEN
		sSrDiscAmt=0
		sOffDiscAmt=0
		sClubDiscAmt=0
	ELSEIF ActualDisc=sSrDiscAmt THEN
		sJrDiscAmt=0
		sOffDiscAmt=0
		sClubDiscAmt=0
	ELSEIF ActualDisc=sOffDiscAmt THEN
		sJrDiscAmt=0
		sSrDiscAmt=0
		sClubDiscAmt=0
	ELSEIF ActualDisc=sClubDiscAmt THEN
		sJrDiscAmt=0
		sSrDiscAmt=0
		sOffDiscAmt=0
	END IF	

	'Response.write("<br><br>Pos 2")	
	'Response.write("<br>sJrDiscAmt = "&sJrDiscAmt)
	'Response.write("<br>sStDiscAmt = "&sSrDiscAmt)
	'Response.write("<br>sOffDiscAmt = "&sOffDiscAmt)
	'Response.write("<br>sClubDiscAmt = "&sClubDiscAmt)	
	'Response.write("<br>ActualDisc = "&ActualDisc)	

	
	IF ActualDisc = 0 THEN sDiscNote =""
END IF



IF TRIM(sMoneyOverride)<>"" THEN
	sTotalFormFees = cdbl(0)
ELSE
	sTotalFormFees = sEntryFee + sLateFeeTot + sAWSEFDonation + sBanquetTot + ActualDisc + sOF1Fee + sOF2Fee + sOF3Fee + sOF4Fee + sOF5Fee + sOF6Fee + sOF7Fee + sOF8Fee + sOF9Fee + sOF10Fee
END IF

'IF sMemberID="000001151" THEN
'		response.write("<br>Line 3460 Registration - sTotalFormFees = "&sTotalFormFees)
'END IF


IF TestMode="yes" AND Session("AdminMenuLevel")>=50 THEN
		Response.write("<br><br>ActualDisc = "&ActualDisc)
		Response.write("<br>sTotalFormFees = "&sTotalFormFees)
		' response.end
END IF
 

' --- Sets the discount note that goes at the bottom of the page ---
SetDiscountNote




SetSessionStatusText




END SUB



' --------------------
  SUB SetDiscountNote
' --------------------

sDiscNote=""
IF ActualDisc < 0 THEN 
	IF sDiscMeth = 0 THEN
		sDiscNote = "NOTE: Cummulative discount does NOT apply to Late Fees !"
	ELSE
		sDiscNote = "NOTE: Discount based on largest single discount (N/A to Late Fees)"		
	END IF
END IF



END SUB


' --------------------------------------
  SUB DisplayCurrentValues (FromWhere)
' --------------------------------------

response.write("THIS Request Sent From "&FromWhere)%>
<br>
<TABLE class="innertable" Align=center WIDTH=100%>
	<TR>
	  <TD>sTEvent(<%=EvtNo%>)</TD>
	  <TD>sDiv(<%=EvtNo%>)</TD>
	  <TD>sFeeClass(<%=EvtNo%>)</TD>
	  <TD>sFeeRounds(<%=EvtNo%>)</TD>
	  <TD>sQfyOverride(<%=EvtNo%>)</TD>
	</TR><%

FOR EvtNo=1 TO TotEv  %>
	<TR>
	  <TD><%=sTEvent(EvtNo)%></TD>
	  <TD><%=sDiv(EvtNo)%></TD>
	  <TD><%=sFeeClass(EvtNo)%></TD>
	  <TD><%=sFeeRounds(EvtNo)%></TD>
	  <TD><%=sQfyOverride(EvtNo)%></TD>
	</TR><%

NEXT  %>

</TABLE>
<br><%	
	response.write("sTotalFormFees = "&sTotalFormFees)%><br><%
	response.write("sEntryFee = "&sEntryFee)%><br><%
	response.write("sBanquetTot = "&sBanquetTot)%><br><%
	response.write("sOffDiscAmt = "&sOffDiscAmt)%><br><%
	response.write("sClubDiscAmt = "&sClubDiscAmt)%><br><%
	response.write("sAWSEFDonation = "&sAWSEFDonation)%><br><%
	response.write("sSrDiscAmt = "&sSrDiscAmt)%><br><%
	response.write("sJrDiscAmt = "&sJrDiscAmt)%><br><%
	response.write("ActualDisc = "&ActualDisc)%><br><%


END SUB





' ------------------
  SUB PRINTRECEIPT
' ------------------



%>
  <HTML><HEAD>

  <STYLE TYPE="text/css"><!--
      #bgimg  {
        background-color: #FFFFFF;
        background-image: url(/images/logos/USAWatermark80.jpg);
        background-position: center;
        background-repeat: no-repeat;
        background-attachment: fixed;
        width:100%;
        height:100%;
        margin:0px;
      }
    --></STYLE>
    <TITLE>Registration</TITLE>
    </HEAD>
    <body>
    <div id="bgimg">

<%


' --- SUB In tools_registration.asp ---
DefineTourVariables_New

' --- Determines EVENTS and EVENT NAMES for this tourament - in tools_include.asp ---
RegistrationEventsOffered (sTSptsGrpID)

DefineMemberVariables

WhichTable=RegGenTableName
InitializeFromTable
DetermineTotalFeesActuallyPaid

' --- Read transactions from Credit Card Table ----
SET rsPayLog=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT MemberID, OrderNo, TransDate, Amount, PayType FROM "&RegPaymentTableName
sSQL = sSQL + " WHERE Left(TourID,6) = '" & SQLClean(left(sTourID,6)) & "' AND MemberID = '"&sMemberID&"' AND Result='0'"
sSQL = sSQL + " ORDER BY TransDate DESC, OrderNo DESC"
rsPayLog.open sSQL, SConnectionToTRATable, 3, 3

'response.write(rsPayLog.eof)


%>
<TABLE ALIGN="Center" BORDER="0" WIDTH=100%>
	<tr>
	    <td Colspan="5" ALIGN="center" valign="top"><img src="/rankings/images/logos/usawslogo_no_sub.jpg"></td>
	</tr>
	<tr>

	     <td align=center><a href='#' onclick='window.print()' title="Click here to Print"><input type=submit value="Print Receipt"></a></td>
	    <TD  WIDTH = 70% COLSPAN=3 ALIGN="center" vAlign="top"><FONT size="5" face="<% =font1%>" COlOR="#0000CD"><b>Registration Receipt<b></FONT></TD>

	   <td><input type=button value="Close Window" title="Close this window to return to registration." onclick="javascript:window.close();"></td>


	</tr>
</TABLE>
<br>
<TABLE ALIGN="Center" BORDER="1" CELLPADDING=5 WIDTH=100%>
	<tr>
	  <TD ALIGN="left" width=20% vAlign="top"><font size=<% =fontsize2 %> face="<% =font1%>" COlOR="#000000">&nbsp; <b>Member Name</b></FONT>
	  <br><FONT size=<% =fontsize2 %> face="<% =font1%>" COlOR="#0000CD">&nbsp; <% =sFirstName&" "&sLastName %></FONT></TD>

	  <TD ALIGN="left" width=15% vAlign="top"><font size=<% =fontsize2 %> face="<% =font1%>" COlOR="#000000">&nbsp; <b>Member ID</b></FONT>
	  <br><FONT size=<% =fontsize2 %> face="<% =font1%>" COlOR="#0000CD">&nbsp; <% =sMemberID %></FONT></TD>
  
	  <TD ALIGN="left" width=15% vAlign="top"><font size=<% =fontsize2 %> face="<% =font1%>" COlOR="#000000">&nbsp; <b>City/ST</b></FONT>
	  <br><FONT size=<% =fontsize2 %> face="<% =font1%>" COlOR="#0000CD">&nbsp; <% =sMembCity&", "&sMembState %></FONT></TD>

	  <TD ALIGN="left" vAlign="top"><font size=<% =fontsize2 %> face="<% =font1%>" COlOR="#000000">&nbsp; <b>Birth Date</b></FONT>
	  <br><FONT size=<% =fontsize2 %> face="<% =font1%>" COlOR="#0000CD">&nbsp; <% =sMembBirth %></FONT></TD>

	  <TD ALIGN="left" vAlign="top"><font size=<% =fontsize2 %> face="<% =font1%>" COlOR="#000000">&nbsp; <b>Gender</b></FONT>
	  <br><FONT size=<% =fontsize2 %> face="<% =font1%>" COlOR="#0000CD">&nbsp; <% =sMembSex %></FONT></TD>
	</tr>  

	<% 

	' --------------------------------------------------------------------------------
	' ----------------------  MEMBERSHIP AND ENTRY STATUS  ------------------------------ 
	' ----------------------------------------------------------------------------------- %> 		

	<tr>
	  <TD ALIGN="left" vAlign="top"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="#000000">&nbsp; <b>Competition Status</b></FONT><%



	  	IF sCanSkiTour = 1 THEN  
	    		%><br><FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR="#0000CD">&nbsp; OK - <%=sTypeDesc%></FONT></td><%
	  	ELSE
	   	 	%><br><FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR=red>&nbsp; <% =sMembTypeCode %> - Upgrade Required</FONT></td><%
	  	END IF 

		%><TD ALIGN="left" vAlign="top"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="#000000">&nbsp; <b>Expiration</b></FONT><%


		' ---  Checks End Date of tournament against Expiration Date of membership record  ---
		IF DateDiff("d", sEffectiveto, sTDateE) <= 0  THEN
	    		%><br><FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR=blue>&nbsp; OK - <% =sEffectiveto %></FONT></td><%
	  	ELSE
	    		%><br><FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR=red>&nbsp; Renew - <% =sEffectiveto %></FONT></td><%
	  	END IF 

		' -------------------------------
		' ------  Payment Status  -------
		' -------------------------------

		%><TD ALIGN="left" vAlign="top"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="#000000"><b>Payment Status</b></FONT>&nbsp;&nbsp;<%

		' ----  Fees from RegGenTable (Previous) are less than current form values  -----
		IF sTotalFormFees <> 0 AND sTotalPreviousPayments < sTotalFormFees THEN
		 	%><br><FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR="red"&nbsp; >Balance Due</FONT></TD><%		 

		' ----  Fees from RegGenTable (Previous) are greater than current form values  -----
		  ELSEIF sTotalFormFees <> 0 AND sTotalPreviousPayments > cdbl(sTotalFormFees) THEN
		 	%><br><FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR="red">&nbsp; Refund Due</FONT></TD><%		 

		' ----------------------------------------------------------------------------------------------
		' ---- *****  MARK - DO WE NEED A NEW CONDITION?  when FORM has never been confirmed and displaying original information 
		  ELSE
			%><br><FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR="#0000CD">&nbsp; Paid in Full</FONT></TD><%
		END IF 

		' ----------------------------------------
		' -------- Liability Waiver --------------
		' ----------------------------------------  
		%>
		<TD ALIGN="left" vAlign="top"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="#000000"><b>&nbsp; Release</b></FONT>&nbsp;&nbsp;<%

		IF TRIM(Session("sRelease")) = "" THEN
		 	%><br><FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR="red">&nbsp; Not Signed</FONT></TD><%
		ELSE
			%><br><FONT size=<% =fontsize2 %> face="<% =font1 %>" COlOR="#0000CD">&nbsp; Complete</FONT></TD><%
		END IF 
	
		' ----------------------------------------
		' ----------- Personal Bio  --------------
		' ----------------------------------------
		%>
		<TD ALIGN="left" vAlign="top"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="#000000"><b>&nbsp; Pers Bio</b></FONT>&nbsp;&nbsp;
		<%

		IF sBioDone = "Y" THEN 
			%><br><FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor2 %>">&nbsp; Complete</FONT></TD><%
		ELSE  
			%><br><FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR="red">&nbsp; Incomplete</FONT></TD><% 
			sErrorNo = sErrorNo + 1
		END IF %>

	  </td>
	</tr>
	    

	<tr>
	  <TD ALIGN="left" vAlign="top"><FONT COlOR="#000000" size=<% =fontsize2 %> face="<% =font1%>" >&nbsp; <b>Tour ID</b></FONT>
	  <br><FONT COlOR="#0000CD" size=<% =fontsize2 %> face="<% =font1%>">&nbsp; <% =sTourID %></FONT></TD>

	  <TD colspan=2 ALIGN="left" vAlign="top"><FONT COlOR="#000000" size=<% =fontsize2 %> face="<% =font1%>">&nbsp; <b>Tournament or Clinic Name</b></FONT>
	  <br><FONT COlOR="#0000CD" size=<% =fontsize2 %> face="<% =font1%>">&nbsp;<% =sTourName %></FONT></TD>
  
	  <TD ALIGN="left" vAlign="top"><FONT COlOR="#000000" size=<% =fontsize2 %> face="<% =font1%>">&nbsp; <b>City</b>/ST</FONT>
	  <br><FONT COlOR="#0000CD" size=<% =fontsize2 %> face="<% =font1%>">&nbsp; <% =sTourCity&", "&sTourState %></FONT></TD>

	  <TD ALIGN="left" vAlign="top"><FONT COlOR="#000000" size=<% =fontsize2 %> face="<% =font1%>">&nbsp; <b>Dates</b></FONT>
	  <br><FONT COlOR="#0000CD" size=<% =fontsize2 %> face="<% =font1%>">&nbsp; <% =sTDateS&" to "&sTDateE %></FONT></TD>
	</tr>

</TABLE>
<br>
<TABLE ALIGN="Center" CELLPADDING="2" CELLSPACING="0" BORDER="1" WIDTH=100%>


	<tr>
	  <TD WIDTH = 20% ALIGN="left" vAlign="top"><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>><br>&nbsp; <b>Event</b></FONT>
	  <TD WIDTH = 30% ALIGN="left" vAlign="top"><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>><br>&nbsp; <b>Divisions Entered</b></FONT>
	  <TD WIDTH = 40% ALIGN="left" vAlign="top"><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>><br>&nbsp; <b>Boat/Ramp/Weight</b></FONT>
	</tr><%  
	

	FOR EvtNo=1 TO TotEv 
	  IF sDiv(EvtNo) <> "" THEN %>
		<tr>
		<td><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>>&nbsp; <%= sTEventName(EvtNo) %></td>
		<td><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>>&nbsp; <%= sDiv(EvtNo) %></td><%

		IF TRIM(sBoat(EvtNo)) <> "" THEN 
			BoatList = ",Correct Craft, Malibu, Mastercraft"
			BoatCodeList =",CC, MA, MC"
			BoatArray = Split(BoatList,",")  
			BoatCodeArray = Split(BoatCodeList,",")  

			FOR kvar = 1 to UBOUND(BoatArray)
				IF TRIM(BoatCodeArray(kvar)) = TRIM(sBoat(EvtNo)) THEN BoatName=BoatArray(kvar)
			NEXT %>
			<td><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>> &nbsp;Trick Boat Selected - &nbsp;<% =BoatName %></td><%		
		ELSEIF TRIM(sRampHeight) <> "" THEN %>
			<td><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>>&nbsp;<% =sRampHeight %> &nbsp;Ft &nbsp;Ramp Height</td><%		
		ELSE %>
			<td><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>>&nbsp; </td><%
		END IF %>
		</tr><%  
	  END IF
	NEXT	 %>


	</TABLE>

	<br>

	<TABLE ALIGN="center" WIDTH=50% BORDER="1" CELLPADDING="5" CELLSPACING="0">
	  <tr>
 	    <td align="center"><font color="<% =TextColor2 %>" size=<% =fontsize2 %> face=<% =font1 %>><b>NOTES:</b>
		<br><% =ReceiptNote1 %>
		<br><% =ReceiptNote2 %>
		<br><% =ReceiptNote3 %>
		<br><% =ReceiptNote4 %>
		<br><% =ReceiptNote5 %>
		<br></font></td>
	  </tr>

	</TABLE>
	<br>

<%
' --------------------------------------------------------------------------------------------------------------
' ---------------  Beginning of Financial Section --------------------------------------------------------------
' --------------------------------------------------------------------------------------------------------------

%>

<TABLE ALIGN="center" WIDTH=100% BORDER="0" CELLPADDING="3" CELLSPACING="0" width=100%>
  <TR>
    <TD VALIGN="top">

      <TABLE VALIGN="top" ALIGN="left" WIDTH=45% BORDER="1" CELLPADDING="3" CELLSPACING="0" width=100%>
	<tr>
		<td colspan=2 align="center"><font color="<% =textcolor1 %>" size=<% =fontsize2 %> face=<% =font1 %>><b>Current Payment Status</b></font></td>
	</tr>

	<tr>
		<td align="right"><font color="<% =textcolor1 %>" size=<% =fontsize2 %> face=<% =font1 %>><b>Description</b></font></td>
		<td align="right"><font color="<% =textcolor1 %>" size=<% =fontsize2 %> face=<% =font1 %>><b>Amount</b></font></td>
	</tr>

	<tr> 
	<td width=20% align="right"><font color="<% =textcolor1 %>" size=<% =fontsize2 %> face=<% =font1 %>>Sub-Total Entry Fees</font></td><% 
		

		 ' ---------------------   NEED TO DEAL WITH FAMILY MEMBERSHIP   ----------------------------


		  %><td width=15% align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>><%=FormatCurrency(sEntryFee,2)%></font></td>
	</tr><%

		' -------------------------------------------	
		' ---- Donation to AWSEF Building Fund  -----
		' -------------------------------------------
	  
		IF cdbl(sAWSEFDonation) > 0 THEN  %>
			<tr>
			<td align="right"><font size=<% =fontsize2 %> face=<% =font1 %>>AWSEF Donation</font></td>
			<td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>><%=FormatCurrency(sAWSEFDonation,2)%></font></td>
			</tr><%
		END IF 

		' ------------------------------------------------------------	
		' ---- Discount to Junior B/G 1-3 per Tour_Manager.asp   -----
		' ------------------------------------------------------------

		  IF cdbl(sJrDiscAmt) <> 0 THEN %>
			<tr>
			  <td align="right"><font size=<% =fontsize2 %> face=<% =font1 %>>Junior Discount</font></td>
			  <td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>><%= FormatCurrency(sJrDiscAmt,2) %></font></td>
			</tr><%
		  END IF 	




' --------------------------------------------------------------------------------------------------------------------
' ----  "NOTE:"  FUTURE - Make AGE for Senior Discount established by division setting in Tour and DivisionTable ----- 
' --------------------------------------------------------------------------------------------------------------------





		' -------------------------------------------------------------------------	
		' ---- Discount to divisions M/W-6 if specified in Tour_Manager.asp   -----
		' -------------------------------------------------------------------------

		IF cdbl(sSrDiscAmt) <> 0 THEN  %>
			<tr>
			  <td align="right"><font size=<% =fontsize2 %> face=<% =font1 %>>Senior Discount</font></td>
			  <td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>><%= FormatCurrency(sSrDiscAmt,2) %></font></td>
			</tr><%
		END IF

		' -------------------------------------------------------------------------	
		' ---------- Discount to OFFICIALS if specified in Tour_Manager.asp   -----
		' -------------------------------------------------------------------------  

		IF cdbl(sOffDiscAmt) <> 0 THEN  %>
		  	<tr>
			  <td align="right"><font color="<% =textcolor1 %>" size=<% =fontsize2 %> face=<% =font1 %>>Officials Discount</font></td>
			  <td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>><%= FormatCurrency(sOffDiscAmt,2) %></font></td>	
			</tr><%
		END IF  

		' -------------------------------------------------------------------------------------------------	
		' ---------- Discount to CLUB MEMBERS if match to ClubCode as specified in Tour_Manager.asp   -----
		' -------------------------------------------------------------------------------------------------  

		IF cdbl(sClubDiscAmt) <> 0 THEN  %>
			  <tr>
			    <td align="right"><font size=<% =fontsize2 %> face=<% =font1 %>>Club Member Discount</font></td>
			    <td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>><%= FormatCurrency(sClubDiscAmt,2) %></font></td>	
			  </tr><%
		END IF  


		' ---------------------------------------------
		' --------  LATE FEE --------------------------
		' ---------------------------------------------  %>
		<tr>
		  <td align="right">
			    <font size=<% =fontsize2 %> face=<% =font1 %>>Registration Date:</font>
			    <font color=<% =TextColor2 %> size=<% =fontsize2 %> face=<% =font1 %>><%=sMembRegDate%></font>
		  </td>

		  <td align="right">&nbsp;  <%
			    IF Cdbl(sLateFeeTot) <> 0 THEN  %>
				  <font size=<% =fontsize2 %> face=<% =font1 %>>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Late Fee - <%=sLateDays%> Days&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>
				   <font color="<% = textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>><%= FormatCurrency(sLateFeeTot,2) %></font><%
			    END IF %>
		  </td> 
		</tr>  <%


		' ----------------------------------	
		' ---------- Banquet Tickets   -----
		' ----------------------------------  

		IF cdbl(sBanquetQty) > 0 THEN  %>
			  <tr>
			    <td align="right"><font size=<% =fontsize2 %> face=<% =font1 %>>Banquet Tickets</font></td>
			    <td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>><%= FormatCurrency(sBanquetTot,2) %></font></td>	
			  </tr><%
		END IF  

		IF cdbl(sOF1Qty) > 0 THEN  %>
			  <tr>
			    <td align="right"><font size=<% =fontsize2 %> face=<% =font1 %>><%=sOF1Desc%></font></td>
			    <td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>><%= FormatCurrency(sOF1Fee,2) %></font></td>	
			  </tr><%
		END IF  
		IF cdbl(sOF2Qty) > 0 THEN  %>
			  <tr>
			    <td align="right"><font size=<% =fontsize2 %> face=<% =font1 %>><%=sOF2Desc%></font></td>
			    <td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>><%= FormatCurrency(sOF2Fee,2) %></font></td>	
			  </tr><%
		END IF  
		IF cdbl(sOF3Qty) > 0 THEN  %>
			  <tr>
			    <td align="right"><font size=<% =fontsize2 %> face=<% =font1 %>><%=sOF3Desc%></font></td>
			    <td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>><%= FormatCurrency(sOF3Fee,2) %></font></td>	
			  </tr><%
		END IF  
		IF cdbl(sOF4Qty) > 0 THEN  %>
			  <tr>
			    <td align="right"><font size=<% =fontsize2 %> face=<% =font1 %>><%=sOF4Desc%></font></td>
			    <td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>><%= FormatCurrency(sOF4Fee,2) %></font></td>	
			  </tr><%
		END IF  
		IF cdbl(sOF5Qty) > 0 THEN  %>
			  <tr>
			    <td align="right"><font size=<% =fontsize2 %> face=<% =font1 %>><%=sOF5Desc%></font></td>
			    <td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>><%= FormatCurrency(sOF5Fee,2) %></font></td>	
			  </tr><%
		END IF  
		IF cdbl(sOF6Qty) > 0 THEN  %>
			  <tr>
			    <td align="right"><font size=<% =fontsize2 %> face=<% =font1 %>><%=sOF6Desc%></font></td>
			    <td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>><%= FormatCurrency(sOF6Fee,2) %></font></td>	
			  </tr><%
		END IF  
		IF cdbl(sOF7Qty) > 0 THEN  %>
			  <tr>
			    <td align="right"><font size=<% =fontsize2 %> face=<% =font1 %>><%=sOF7Desc%></font></td>
			    <td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>><%= FormatCurrency(sOF7Fee,2) %></font></td>	
			  </tr><%
		END IF  
		IF cdbl(sOF8Qty) > 0 THEN  %>
			  <tr>
			    <td align="right"><font size=<% =fontsize2 %> face=<% =font1 %>><%=sOF8Desc%></font></td>
			    <td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>><%= FormatCurrency(sOF8Fee,2) %></font></td>	
			  </tr><%
		END IF  
		IF cdbl(sOF9Qty) > 0 THEN  %>
			  <tr>
			    <td align="right"><font size=<% =fontsize2 %> face=<% =font1 %>><%=sOF9Desc%></font></td>
			    <td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>><%= FormatCurrency(sOF9Fee,2) %></font></td>	
			  </tr><%
		END IF  
		IF cdbl(sOF10Qty) > 0 THEN  %>
			  <tr>
			    <td align="right"><font size=<% =fontsize2 %> face=<% =font1 %>><%=sOF10Desc%></font></td>
			    <td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>><%= FormatCurrency(sOF10Fee,2) %></font></td>	
			  </tr><%
		END IF  



		' -------------------------------------------------------------------------------------------------
		' -----  Calculate Applied Discount depending on which discount method was selected  --------------
		' -------------------------------------------------------------------------------------------------  %>

		<tr>		
		    <td align="right"><font color="#000000" size=<% =fontsize2 %> face=<% =font1 %>><b>TOTAL ALL</b></font></td>
		    <td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>><%= FormatCurrency(sTotalFormFees,2) %></font></td>
		</tr>

		<tr>  <%
			Dim SomeDisc
			SomeDisc = "NO"
			IF cdbl(sClubDiscAmt) <> 0 OR cdbl(sOffDiscAmt) <> 0 OR cdbl(sJrDiscAmt) <> 0 OR cdbl(sSrDiscAmt) <> 0 THEN
				SomeDisc = "YES"
			END IF

			' --- Sets the discount note that goes at the bottom of the page ---
			SetDiscountNote %>

			<td colspan="2" align="center"><font color="#000000" size=<% =fontsize2 %> face=<% =font1 %>><%=sDiscNote%></font></td>
		</tr> 
		<tr>
			<td colspan="2" align="center"><%
			  IF TRIM(Session("sWhichFamilyMemberPaid"))<>"" AND sMaxFamMembers>1 THEN %>
				<font color="<% =textcolor1 %>" size=<% =fontsize2 %> face=<% =font1 %>>&nbsp;<%=Session("sWhichFamilyMemberPaid")%> was charged for the 'Family Entry Fee'.<br>&nbsp;Late entry fees and other charges are not included in Family Entry Fee.</font><%
			  ELSEIF TRIM(Session("sWhichFamilyMemberPaid"))<>"" AND sMaxFamMembers=1 THEN %>
				<font color="<%=textcolor1%>" size=<% =fontsize2 %> face=<% =font1 %>>&nbsp;<%=Session("sWhichFamilyMemberPaid")%> was charged for the 'Family Entry Fee'. All other entries for family members will be charged the 'Additional Family Member' fee.&nbsp;Late entry fees and other charges are not included in Family Entry Fee.</font><%
			  END IF %>
			</td> 
		</tr><%

		Dim PreviousPaid
		PreviousPaid = 0

		IF NOT rsPayLog.eof THEN
			rsPayLog.movefirst		


			DO WHILE NOT rsPayLog.eof  

				IF IsNull(rsPayLog("Amount")) <> False THEN 

				ELSE 
					PreviousPaid = cdbl(PreviousPaid) + cdbl(rsPayLog("Amount"))
				END IF
				rsPayLog.movenext
			LOOP %>
			  <tr>
			    <td align="right"><font color="<% =textcolor1 %>" size=<% =fontsize2 %> face=<% =font1 %>><b>Total All Payments</b></font></td>
			    <td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>>&nbsp;<% = FormatCurrency(PreviousPaid,2) %></font></td>
			  </tr>
			  <tr> <%
			    IF PreviousPaid > sTotalFormFees THEN  %>
				    <td align="right"><font color="<% =textcolor1 %>" size=<% =fontsize2 %> face=<% =font1 %>><b>Credit Due</b></font></td>
				    <td align="right"><font color="red" size=<% =fontsize2 %> face=<% =font1 %>>&nbsp;<% = FormatCurrency(sTotalFormFees - PreviousPaid,2) %></font></td><%
			    ELSE  %>
				    <td align="right"><font color="<% =textcolor1 %>" size=<% =fontsize2 %> face=<% =font1 %>><b>Balance Due</b></font></td>
				    <td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>>&nbsp;<% = FormatCurrency(sTotalFormFees - PreviousPaid,2) %></font></td><%
			    END IF  %>		
			  </tr>	<%

			' --- Resets for next section ---
			rsPayLog.movefirst
		  END IF %>

		</TABLE><%



		' -------------------------------------------------------------------------------------------------
		' -----------------------------  BEGIN Transaction Table  -----------------------------------------
		' -------------------------------------------------------------------------------------------------  %>

		<TABLE VALIGN="top" ALIGN="right" WIDTH=45% BORDER="1" CELLPADDING="3" CELLSPACING="0" width=100%>
		  <tr>
			<td colspan=4 align="center"><font color="<% =textcolor1 %>" size=<% =fontsize2 %> face=<% =font1 %>><b>Transaction History For This Tournament</b></font></td>
		  </tr>

		  <tr>
			<td align="right"><font color="<% =textcolor1 %>" size=<% =fontsize2 %> face=<% =font1 %>><b>Trans No</b></font></td>
			<td align="right"><font color="<% =textcolor1 %>" size=<% =fontsize2 %> face=<% =font1 %>><b>Method</b></font></td>
			<td align="right"><font color="<% =textcolor1 %>" size=<% =fontsize2 %> face=<% =font1 %>><b>Date</b></font></td>
			<td align="right"><font color="<% =textcolor1 %>" size=<% =fontsize2 %> face=<% =font1 %>><b>Amount</b></font></td>  <%


		  PreviousPaid = 0		  
		  IF NOT rsPayLog.eof THEN


			DO WHILE NOT rsPayLog.eof  

				IF rsPayLog.eof=true THEN EXIT DO
				IF cdbl(rsPayLog("Amount"))>=cdbl(0) THEN %>
					<tr>
					<td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>>&nbsp;<% =rsPayLog("OrderNo") %></font></td>	
					<td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>>&nbsp;<% =rsPayLog("PayType") %></font></td>	
					<td align="right"><font color="<% =textcolor1 %>" size=<% =fontsize2 %> face=<% =font1 %>>&nbsp;<% =rsPayLog("TransDate") %></font></td>
					<td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>>&nbsp;<% =formatCurrency(rsPayLog("Amount"),2) %></font></td>	
					</tr>  <%

				   	PreviousPaid = cdbl(PreviousPaid) + cdbl(rsPayLog("Amount"))
				ELSE
					'response.write("Amount = "&rsPayLog("Amount"))
					'PreviousPaid = cdbl(PreviousPaid) + cdbl(rsPayLog("Amount"))
				END IF
				rsPayLog.movenext
	
			LOOP %>
			  <tr>
			    <td colspan=3 align="right"><font color="<% =textcolor1 %>" size=<% =fontsize2 %> face=<% =font1 %>><b>Total All Payments</b></font></td>
			    <td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>>&nbsp;<% = FormatCurrency(PreviousPaid,2) %></font></td>
			  </tr>	<%

		  END IF %>


		</TABLE>

	    </TD>
	  </TR>
	</TABLE><% 


END SUB  ' Bottom of CASE Else of sRunByWhat 



' ----------------------------------
  SUB ReadContDispTableValues
' ----------------------------------


' --- Read values from Cont_Display table and set parameters to determine functions ----
SET rsContDisp=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT * FROM "&ControlDisplayTableName
rsContDisp.open sSQL, SConnectionToTRATable, 3, 3

IF NOT rsContDisp.eof THEN

	sEntryEmail= rsContDisp("EntryEmail")
	sEntryEmailAdm=rsContDisp("EntryEmailAdm")
	sEntryEmailHQ=rsContDisp("EntryEmailHQ")
	sWaiverEmail=rsContDisp("WaiverEmail")
	sWaiverEmailAdm=rsContDisp("WaiverEmailAdm")
	sWaiverEmailHQ=rsContDisp("WaiverEmailHQ")
	sPasswordEmail=rsContDisp("PasswordEmail")
	sPasswordEmailAdm=rsContDisp("PasswordEmailAdm")
	sPasswordEmailHQ=rsContDisp("PasswordEmailHQ")
	sSkipWaiver=rsContDisp("SkipWaiver")
	sSkipWaiverAdm=rsContDisp("SkipWaiverAdm")
	sSkipWaiverHQ=rsContDisp("SkipWaiverHQ")
	sForceWaiver=rsContDisp("ForceWaiver")
	sForceWaiverAdm=rsContDisp("ForceWaiverAdm")
	sForceWaiverHQ=rsContDisp("ForceWaiverHQ")

	sDispDebugButtons=rsContDisp("DispDebugButtons")
	sDispDebugButtonsAdm=rsContDisp("DispDebugButtonsAdm")
	sDispDebugButtonsHQ=rsContDisp("DispDebugButtonsHQ")

END IF


END SUB



' -----------------------
    SUB SendWaiverEmail 
' -----------------------


' --- Gets Waiver Info from RegGenTable --- 
SET rsRegTemp=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT * FROM "&RegTempTableName
sSQL = sSQL + " WHERE Left(TourID,6) = '"&SQLClean(left(sTourID,6))&"' AND MemberID = '"&sMemberID&"'"
rsRegTemp.open sSQL, SConnectionToTRATable, 3, 3

sWaiverCode = rsRegTemp("WaiverCode")
' sSignWaiver = SQLClean(rsRegTemp("SignWaiver"))
sSignWaiver = rsRegTemp("SignWaiver")

rsRegTemp.close

'response.write("</div></div><br>"&sSQL)
'response.end

DefineTourVariables_New
DefineMemberVariables



ebody = "<html>"
ebody = ebody & "<head>"
ebody = ebody & "<title>Waiver and Release</title>"
ebody = ebody & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
ebody = ebody & "</head>"
ebody = ebody & "<body bgcolor=""#FFFFFF"" text=""#000000"">"
ebody = ebody & "<div align=""center"">"


ebody = ebody & "<TABLE BORDER=4 ALIGN=CENTER CELLPADDING=3 CELLSPACING=0 BGCOLOR="&TableColor1&" width=85% >"
ebody = ebody & "<TR>"
ebody = ebody & "<TD BGCOLOR=red><center><font face="&font1&" color=#FFFFFF size=4><b>Waiver and Release Form</b></font></TD>"
ebody = ebody & "</TR>"
 
ebody = ebody & "<TR>"
ebody = ebody & "<TD VALIGN=top>"


ebody = ebody & "<table border=""0"" cellspacing=""0"" cellpadding=""3"" width=""100%"">"
ebody = ebody & "<tr>"


IF Session("sMembAge") < 18 THEN
	subTitle="Waiver for MINOR Participant - WaiverID: "&sWaiverCode
ELSE
	subTitle="Waiver for ADULT Participant - WaiverID: "&sWaiverCode
END IF  

ebody = ebody & "<td Align=center>"	
ebody = ebody & "<font face="&font1&" size=4 ><b>PARTICIPANT WAIVER AND RELEASE OF LIABILITY,</b></font><br>"
ebody = ebody & "<font face="&font1&" size=4><b>ASSUMPTION OF RISK AND INDEMNITY AGREEMENT</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>"&subTitle&"</b></font>"
ebody = ebody & "<br><br>"
ebody = ebody & "<font face="&font1&" color="&TextColor2&" size=3><b>"&sTourName&"</font></b>"
ebody = ebody & "<br><br>"
ebody = ebody & "<font face="&font1&" size=2><b>MemberID = </font><font color="&TextColor2&" face="&font1&" size=2>"&Session("sMemberID")
ebody = ebody & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="&TextColor1&" face="&font1&" size=2>Participant:</font>"
ebody = ebody & "<font color="&TextColor2&" face="&font1&" size=2>&nbsp;&nbsp;"&sFirstname&"&nbsp;"&sLastName&"</font></b><br>"

ebody = ebody & "</center>"
ebody = ebody & "<br>"
ebody = ebody & "</td></tr>"


ebody = ebody & "<td Align=left>"	
ebody = ebody & "<P><font color="&TextColor1&" size=1 face="&font1&">"

'response.write("</div></div><font color=red>sWaiverCode = "&sWaiverCode&"</font>")
'response.end
	
Set objfso = CreateObject("Scripting.FileSystemObject")

' --- Formerly ReleaseVersion
IF objfso.FileExists(PathtoWaivers & "\waiver-"&sWaiverCode&".txt") THEN
	SET objstream=objFSO.opentextfile(PathtoWaivers & "\waiver-"&sWaiverCode&".txt")

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
		
IF Session("sMembAge") < 18 THEN
		ebody = ebody & "<br>"
		ebody = ebody & "<font color="&textcolor3&" face="&font1&" size=3><b>Minors under 18 Years may NOT accept liability waiver.</b></font>"
		ebody = ebody & "<br><br>"
		ebody = ebody & "<font color="&textcolor3&" face="&font1&" size=3><b>Name of Parent or Guardian acccepting this waiver on behalf of this minor:</b></font>&nbsp;&nbsp"
		ebody = ebody & "<font color="&textcolor2&" face="&font1&" size=3><b>"&sSignWaiver&"</b></font>"
ELSE  
		ebody = ebody & "<br>"
		ebody = ebody & "<font color="&TextColor3&" face="&font1&" size=3><b>By acccepting this waiver I have acknowledged that I am the 'PARTICIPANT' listed above.</b></font>"
END IF 

ebody = ebody & "<br><br>"
ebody = ebody & "<font color="&TextColor1&" face="&font1&" size=2><b>Date Accepted:&nbsp;&nbsp</font><font color="&TextColor2&" face="&font1&" size=2>"&DATE&"</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "</td></tr>"

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

set rsPW=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT TOP 1 * FROM "&RegPWTableName&" WHERE MemberID = "&sqlclean(sMemberID)
rsPW.open sSQL, sConnectionToTRATable, 3, 1


' ------------------------------------------------
' --- Build the components of the email object ---
' ------------------------------------------------
IF NOT rsPW.eof THEN 
		MembWaiverEmail=rsPW("email")
		IF sWaiverEmail=true THEN SendAddress=MembWaiverEmail
		eMailSubj = "USA Water Ski WAIVER & RELEASE  TourID: "&Session("sTourID")&" - Member: "&sFirstName&" "&sLastName
ELSE
		SendAddress = HQWaiverEmail
		eMailSubj = "USA Water Ski WAIVER & RELEASE  TourID: "&Session("sTourID")&" - Member: "&sFirstName&" "&sLastName&" - Admin Override"
END IF

eMailTo = SendAddress
eMailFrom = ""&HQWaiverEmail
eMailBody = ebody	

IF sWaiverEmailHQ=true THEN eMailBCC = HQWaiverEmail
IF sWaiverEmailMC=true THEN eMailBCC = " "&marksemailaddress


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





' -------------------------------------------------------------------------------------------------------
    SUB SendSPECIALWaiverEmail (sSpecialWaiverCode, sSpecialWaiverHeadline, sSpecialReleaseBannerText)
' -------------------------------------------------------------------------------------------------------

' --- Gets Waiver Info from RegGenTable --- 
SET rsRegTemp=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT * FROM "&RegTempTableName
sSQL = sSQL + " WHERE Left(TourID,6) = '"&SQLClean(left(sTourID,6))&"' AND MemberID = '"&sMemberID&"'"
rsRegTemp.open sSQL, SConnectionToTRATable, 3, 3

sWaiverCode = rsRegTemp("WaiverCode")
' sSignWaiver = SQLClean(rsRegTemp("SignWaiver"))
sSignWaiver = rsRegTemp("SignWaiver")

rsRegTemp.close


DefineTourVariables_New
DefineMemberVariables



' --- New 4-28-2013 - Gets SPECIAL WAIVER info from table based on SiteID rather than hard coding specific tournaments ---
'Dim swaiverSQL, sSpecialWaiverHeadline, sSpecialReleaseBannerText
swaiverSQL = "SELECT SpecialWaiverCode, SpecialWaiverHeadline, SpecialReleaseBannerText FROM usawsrank.TourExtras TE"
swaiverSQL = swaiverSQL + " JOIN sanctions.dbo.TSchedul AS TS"
swaiverSQL = swaiverSQL + "   ON SiteID=TS.TSiteID"
swaiverSQL = swaiverSQL + " WHERE LEFT(TS.TournAppID,6)='"&LEFT(sTourID,6)&"'"

Set rswaiver=Server.CreateObject("ADODB.recordset")
rswaiver.open swaiverSQL, sConnectionToTRATable, 3, 1

testwaiver=false
IF testwaiver=true AND sMemberID="000001151" THEN
		Response.write("<br>Found = ")
		response.write(NOT(rswaiver.eof))
		response.write("<br>rswaiver(SpecialWaiverHeadline) = "&rswaiver("SpecialWaiverHeadline"))
END IF

IF NOT(rswaiver.EOF) THEN
		sSpecialWaiverCode=rswaiver("SpecialWaiverCode")
		sSpecialWaiverHeadline=rswaiver("SpecialWaiverHeadline")
		sSpecialReleaseBannerText=rswaiver("SpecialReleaseBannerText")
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
ebody = ebody & "<font face="&font1&" size=2><b>MemberID = </font><font color="&TextColor2&" face="&font1&" size=2>"&Session("sMemberID")
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
IF sWaiverEmailHQ=true THEN eMailCC = LOCSpecialWaiverEmail


IF sSpecialWaiverEmailMC=true THEN 
		eMailBCC = " "&marksemailaddress
END IF

eMailFrom = ""&HQWaiverEmail
eMailBody = ebody	




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






' ---------------------------
    SUB SendEntryConfirm
' ---------------------------


ebody = "<html>"
ebody = ebody & "<head>"
ebody = ebody & "<title>Notice of On Line Entry</title>"
ebody = ebody & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
ebody = ebody & "</head>"
ebody = ebody & "<body bgcolor=""#FFFFFF"" text=""#000000"">"
ebody = ebody & "<div align=""center"">"


ebody = ebody & "<TABLE BORDER=4 ALIGN=CENTER CELLPADDING=3 CELLSPACING=0 BGCOLOR="&TableColor1&" width=60% >"
ebody = ebody & "<TR>"
ebody = ebody & "<TD BGCOLOR=red><center><font face="&font1&" color=#FFFFFF size=4><b>Notice of USA Water Ski Event Registration</b></font></TD>"
ebody = ebody & "</TR>"
 
ebody = ebody & "<TR>"
ebody = ebody & "<TD VALIGN=top>"


ebody = ebody & "<table border=""0"" cellspacing=""0"" cellpadding=""3"" width=""100%"">"
ebody = ebody & "<tr>"


ebody = ebody & "<td Align=center>"	
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>This message is your notification that an on-line entry has been received for the tournament/clinic listed below.</b></font>"

ebody = ebody & "<br><br>"


ebody = ebody & "<font color="&TextColor2&" face="&font1&" size=4><b>"&sTourName&"</b></font>"
ebody = ebody & "<br>"

ebody = ebody & "<font face="&font1&" size=2><b>SanctionID = </b></font><font color="&TextColor2&" face="&font1&" size=2>"&sTourID&"</font>"
ebody = ebody & "<br>"

ebody = ebody & "<font face="&font1&" size=2><b>Date = </b></font><font color="&TextColor2&" face="&font1&" size=2>"&sTDateS&" to "&sTDateE&"</font></b>"
ebody = ebody & "<br><br>"

ebody = ebody & "<font color="&TextColor2&" face="&font1&" size=4><b>"&sFirstName&" "&sLastName&"</b></font>"
ebody = ebody & "<br>"

ebody = ebody & "<font face="&font1&" size=2><b>MemberID = </b></font><font color="&TextColor2&" face="&font1&" size=2>"&sMemberID&"</font>"
ebody = ebody & "<br><br>"

ebody = ebody & "<font face="&font1&" size=2><b>Events Entered</b></font>"
ebody = ebody & "<br>"

FOR EvtNo=1 TO TotEv
	IF TRIM(sDiv(EvtNo)) <> "" THEN
		ebody = ebody & "<font color="&TextColor2&" face="&font1&" size=2>"&sDiv(EvtNo)&" - "&sTEventName(EvtNo)&"</font>"
		ebody = ebody & "<br>"
	END IF
NEXT

ebody = ebody & "<br>"

'old
ebody = ebody & "<font face="&font1&" size=2>See 'Check Registration Status' report in the 'Events & Register' link for updated qualifications and registration detail.</font><br>"

IF LEFT(sTourID,6)="08S999" THEN 
	ebody = ebody & "<font face="&font1&" color=red size=2><br><b>Please check-in at site to receive your goodie bag.</font></b><br>"
ELSEIF LEFT(sTourID,6)="08M103" THEN 
	ebody = ebody & "<font face="&font1&" color=red size=2><br><b>Please contact Delaina Downes @ 316-213-4949 or jumpskier@cox.net for info on the Family Discount or any other registration questions.</font></b><br>"
END IF

ebody = ebody & "</center>"
ebody = ebody & "<br>"
ebody = ebody & "</td></tr>"

ebody = ebody & "</TABLE>"

ebody = ebody & "</TD></TR>"
ebody = ebody & "</TABLE>"






' ------------------------------------------------------------
' --- Dimension and define all the email related variables ---
' ------------------------------------------------------------

Dim eMailSubj, eMailFrom, eMailTo, eMailCC, eMailBCC, eMailBody
Dim SendAddress, HQEntryEmail, MembWaiverEmail


' --- Testing purposes only ---
'HQEntryEmail="kingsbridge.homes@embarqmail.com"
'sTourEmail="kingsbridge.homes@embarqmail.com"
'MembEntryEmail="mcronehsd@earthlink.net"


HQEntryEmail="competition@usawaterski.org"
SendAddress=""

set rsPW=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT TOP 1 * FROM "&RegPWTableName&" WHERE MemberID = "&sMemberID
rsPW.open sSQL, sConnectionToTRATable, 3, 1




IF NOT rsPW.eof THEN
		MembEntryEmail=rsPW("email")
		IF sEntryEmail=true THEN SendAddress=TRIM(MembEntryEmail)
ELSEIF TRIM(sTourEmail)<>"" AND sReceiveEmail=true THEN
		IF sEntryEmail=true THEN SendAddress=TRIM(sTourEmail)
END IF

eMailTo = SendAddress

IF sEntryEmailAdm=true AND sReceiveEmail=true AND TRIM(sTourEmail)<>"" THEN 
	eMailCC = sTourEmail
END IF

sEntryEmailMC=false
IF sEntryEmailMC=true THEN 
		eMailBCC = " "&marksemailaddress
END IF

eMailFrom = ""&sTourEmail
eMailSubj = " "&sTourID&" - "&sTourName&" - Registration Confirmation for "&sFirstName&" "&sLastName 
eMailBody = ebody	






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
IF TRIM(SendAddress)<>"" AND TRIM(sTourEmail) <> "" THEN
		objMessage.Send
ELSEIF TRIM(SendAddress)<>"" AND TRIM(sTourEmail) = "" THEN
		objMessage.From = HQEntryEmail
		objMessage.Send
END IF
set objMessage = Nothing



END SUB




%>











