<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/includes/func.asp"-->

<%

' ----------------------------------------------------------
' --- The upper section of code provided by Mike Kingham ---
' --- When set to Y it redraws the page as secure page ---
' ----------------------------------------------------------
Draw_Page_As_Secure="Y"
IF Draw_Page_As_Secure="Y" THEN
		IF Request.ServerVariables("HTTPS") = "off" THEN
				URL = "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("SCRIPT_NAME") 

				IF Request.QueryString <> "" THEN
						URL = URL & "?" & Request.QueryString
				END IF

				Response.Redirect URL
				Response.End
		END IF
END IF


' --------------------------------------------------------------------------------
' --- Defines file names throughout this program ---
' --- NOTES FOR JIM MEIS ---
' ---     This must be redefined to the location where you have the file ---
' --------------------------------------------------------------------------------

Dim ThisFileName, RegFileName
ThisFileName="/rankings/CCSJ_Sanctions.asp"



' --------------------------------------------------------------------------------------
' --- Used to get CC on certain automated emails associated with this program ---
' --------------------------------------------------------------------------------------

Dim marksemailaddress2, meisemail, SendMarkSanctionPaymentEmail, SendMeisSanctionPaymentEmail
' --- setting moved to SettingsHQ.asp
marksemailaddress2="cronemarka@gmail.com"
meisemail="jim.meis@meisconsulting.com"
SendMarkSanctionPaymentEmail = "N"
SendMeisSanctionPaymentEmail = "Y"

Dim sOrderNo, CCAmount

Dim FieldErr, SelectedMonth, SelectedYear, StateArray, kvar, sAddressList
Dim CCFirstName, CCLastName, CCAddress1, CCCity, CCState, CCZipCode, CCCardNumber, CCPhone, CCEmail
Dim CCExpMonth, CCExpYear, CCExp2DigitYear
Dim CCV
Dim Resp_result, Resp_Message, Resp_TXN_ID, Resp_Apvl_Code, Resp_CVV2, Resp_AVS, DateNow, sLast4Card
Dim Action, cstat, TestMode, sPayType, MOrder
Dim sPayStatus, sLP
Dim CodeVer

Dim FirstNameError, LastNameError, Address1Error, CityError, StateError, ZipCodeError, PhoneError, EmailError
Dim ExpMonthError, ExpYearError, CardNumberError, CCVError

Dim SanctionPaymentTableName
Dim VisaImage, MasterCardIMage, AmericanExpressImage

Dim ResponseCode, ResponseSubCode, ResponseReasonCode, ResponseReasonText, ApprovalCode, AVSResultCode
Dim TransactionID

Dim ClubID, TID, HQF, OLR, LF, PA, IW, RF, PF

' ===============================================================================================================================
' ===============================================================================================================================

' --- NOTES FOR JIM MEIS ---

' --- This table name is currenting in "usawsrank" but could be moved and redefined in "Sanctions.dbo" with proper connection ---
SanctionPaymentTableName="usawsrank.SanctionPaymentLog"

' --- Session variables must be set in sending program for ClubID and TID to prevent TIMEOUT notice ---

' --- Defines names and locations of card images displayed on form
VisaImage="/rankings/images/logos/visa_card.gif"
MasterCardIMage="/rankings/images/logos/master_card.gif"
AmericanExpressImage="/rankings/images/logos/amex_card.gif"


ReturnOnCancel="http://usawaterski.org/"				' --- Defines redirection when the member abandons the payment processor ---
ReturnOnSuccess="Test_PaymentLaunch.asp"	' --- Defines redirection when transaction is a success ---
ReturnOnTimeOutError = RankPath&"/defaultHQ.asp" 


' ===============================================================================================================================
' ===============================================================================================================================




' --------------------------------------------------------
' --------------------------------------------------------

' --- sPayType is always "Card" in this module ---
sPayType="Card"
CodeVer="12112010"


' --- Form control variables ---
Action = Request("Action")
sLP = Request("sLP")
IF Action="Change Card Info" THEN Action="Modify Information"
IF Action="Verify & Continue" AND FieldErr>0 THEN Action="Modify Information"


' --- Session variables MUST be established in the Sending page as NVP's cannot be passed between a non-secure and secure page ---

ClubID = Request("ClubID")
TID = Request("TID")

'response.write("<br>ClubID = "&ClubID)
'response.write("<br>TID = "&TID)
'response.end

Session("ClubID") = ClubID
Session("TID") = TID


' --- AdminMenuLevel is used in OLR - typically values > 50 have full rights ---
adminmenulevel = Session("adminmenulevel")
IF TRIM(adminmenulevel) = "" THEN adminmenulevel = "1"




sLP = Request("sLP")






' --- This function is used in "rankings" screen themes but not applicable to Sanctions format ---
WriteIndexPageHeader



' ------------------------------------------------------------------------------------------------------
' - Tests whether the system remembers the person's name.  If not, then display the timeout screen.  ---
' ------------------------------------------------------------------------------------------------------

IF TRIM(ClubID) = "" THEN
		Action= "displaytimeout"

ELSE

	IF Action="new" THEN	' --- First time through CCSJ_Sanctions.asp --- 
			Action = "Modify Information"
	
			set rsMemb=Server.CreateObject("ADODB.recordset")
			sSQL = "SELECT TOP 1 * FROM "&MemberTableName
			sSQL = sSQL + " WHERE PersonIDwithCheckDigit = "&sqlclean(ClubID)
			rsMemb.open sSQL, sConnectionToTRATable, 3, 1

			' --- Took this out because Jim Meis didn't want to prepopulate ---
			'CCFirstName = rsMemb("FirstName")
			'CCLastName = rsMemb("LastName")

			IF CCAmount="" THEN CCAmount=0.00
			rsMemb.close

	ELSE
			CCFirstName=SQLClean(Request("CCFirstName"))
			CCLastName=SQLClean(Request("CCLastName"))
	END IF

	' --- Reads remaining variables from the form ---
	ReadAndValidateFormVariables

END IF


'SendPaymentConfirm
'response.end



'sLP = "Test"
IF LCASE(sLP)="test" THEN
		response.write("<br>Line 185 TOP OF PROGRAM")
		response.write("<br>Action="&Action)
		response.write("<br>ClubID="&ClubID)
		response.write("<br>TID="&TID)
		response.write("<br>sOrderNo="&sOrderNo)
		response.write("<br>CCAmount="&CCAmount)
		response.write("<br>sLP="&sLP)
		response.write("<br>HQF="&HQF)
		response.write("<br>OLR="&OLR)
		response.write("<br>LF="&LF)
		response.write("<br>PA="&PA)
		response.write("<br>IW="&IW)
		response.write("<br>RF="&RF)				
END IF


' --- For testing only ---
'Action = "DisplaySuccussForm"

' --- Primary branching performed here ---
SELECT CASE Action

   	CASE "displaytimeout"
				DisplayTimeOutNotice

  	CASE "Return To Main Menu"		
				Response.Redirect(ReturnOnCancel)

   	CASE "Verify & Continue"
				'SendPaymentConfirm
				'response.end
				
				' --- Displays the form that captures name, address, credit cards etc ---
				CCFormStatus="disabled"
				DisplayDataForm

   CASE "Modify Information"
				' --- Displays the form that captures name, address, credit cards etc ---
				CCFormStatus="enabled"
				DisplayDataForm

   CASE "Make Payment Now"
				IF Session("ClubID")<>"" AND Session("TID")<>"" THEN
							' --- Runs the card info through the Wauchovia Charge System ---
							ProcessCardTransaction
				ELSE
							DisplayTimeOutNotice
				END IF

		CASE "DisplayFailureForm"		' --- For testing only ---
				DisplayFailureForm

		CASE "DisplaySuccussForm" 		' --- For testing only ---
				DisplaySuccussForm
		
		CASE ELSE
				response.write("ELSE - not supposed to happen")		
END SELECT


WriteIndexPageFooter



' ====================  END OF MAIN CODE HERE  ===================================




' ----------------------
   SUB LogPaymentInTable
' ----------------------

	' ------------------------------------------
	' ---- Log all PAYMENTS into the table  ----
	' ------------------------------------------

	DateNow = Date

	OpenCon

	sSQL = "INSERT INTO "&SanctionPaymentTableName
	sSQL = sSQL + " (MemberID, TourID, FirstName, LastName, Address1, City, State, ZipCode, Email"

	' -- Removed 12-15-2018 -- MAC
	' sSQL = sSQL + ", Last4Card, ExpMonth, ExpYear"
	sSQL = sSQL + ", Amount, OrderNo, Result, Message, TXN_ID"
	sSQL = sSQL + ", Apvl_Code, CVV2_Resp, AVS_Resp, TransDate, CheckNo, PayType, PayStatus)"

	sSQL = sSQL + " VALUES ('"&ClubID&"', '"&TID&"', '"&CCFirstName&"', '"&CCLastName&"', '"&CCAddress1&"', '"&CCCity&"', '"&CCState&"'"
	sSQL = sSQL + ", '"&CCZipCode&"', '"&CCEmail&"'"

	' -- Removed 12-15-2018 -- MAC
	' sSQL = sSQL + ", '"&sLast4Card&"', '"&CCExpMonth&"', '"&CCExpYear&"'"
	sSQL = sSQL + ", '"&CCAmount&"', '"&sOrderNo&"', '"&ResponseSubCode&"', '"&ResponseReasonText&"', '"&TransactionID&"'"
	sSQL = sSQL + ", '"&ApprovalCode&"', '"&ResponseSubCode&"', '"&AVSResultCode&"', '"&DateNow&"', '"&sCheckNo&"', '"&sPayType&"', '"&sPayStatus&"')"


' --- How variables are defined based on response string ---
'	ResponseCode = aryRString(0)
'	ResponseSubCode = aryRString(1)
'	ResponseReasonCode = aryRString(2)
'	ResponseReasonText = aryRString(3)
'	ApprovalCode = aryRString(4)
'	AVSResultCode = aryRString(5)
'	TransactionID = aryRString(6)

' --- Example of a successful result --- 
' aryRString(1) = 1
' aryRString(2) = 1
' aryRString(3) = This transaction has been approved.
' aryRString(4) = 12009Z
' aryRString(5) = Y
' aryRString(6) = 4013786608 


	IF TestMode = true THEN
		'response.write("<br><br>"&sSQL)
		'response.end
	END IF

	con.execute(sSQL)

	closecon

END SUB


' ---------------------------------
  SUB DisplayCurrentValues_ForTest
' ---------------------------------

Response.write("<font size=1>")
	response.write("<br>Action="&Action)

	response.write("<br>CCAmount="&CCAmount)
	response.write("<br>sOrderNo="&sOrderNo)

	response.write("<br>CCFirstName="&CCFirstName)
	response.write("<br>CCLastName="&CCLastName)
	response.write("<br>CCAddress1="&CCAddress1)
	response.write("<br>CCCity="&CCCity)
	response.write("<br>CCState="&CCState)
	response.write("<br>CCZipCode="&CCZipCode)
	response.write("<br>CCEmail="&CCEmail)
	response.write("<br>CCCardNumber="&CCCardNumber)
	response.write("<br>CCExpMonth="&CCExpMonth)
	response.write("<br>CCExpYear="&CCExpYear)
	response.write("<br>CCV="&CCV)
	response.write("<br>sPayType="&sPayType)
Response.write("</font>")


END SUB


' ---------------------------------
  SUB ReadAndValidateFormVariables
' ---------------------------------


	FieldErr=0

	FirstNameError=false
	LastNameError=false
	Address1Error=false
	CityError=false
	StateError=false
	ZipCodeError=false
	PhoneError=false

	EmailError=false
	ExpMonthError=false
	ExpYearError=false
	CardNumberError=false
	CCVError=false

	sOrderNo=TRIM(SQLClean(Request("sOrderNo")))
	CCAmount = TRIM(SQLClean(Request("CCAmount")))
	sLP = Request("sLP")

	CCAddress1=SQLClean(TRIM(Request("CCAddress1")))
	CCCity=SQLClean(TRIM(Request("CCCity")))
	CCState=SQLClean(TRIM(Request("CCState")))
	CCZipCode=SQLClean(TRIM(Request("CCZipCode")))
	CCCardNumber=SQLClean(Request("CCCardNumber"))

	IF sLP="Test" THEN
		' --- Mastercard Test Number ---
		CCCardNumber="542400000000015"
	END IF


	sLast4Card=RIGHT(CCCardNumber,4)
	CCPhone=RemoveSpace(SQLClean(TRIM(Request("CCPhone"))))

	CCEmail=SQLClean(TRIM(Request("CCEmail")))
	CCExpMonth=SQLClean(Request("CCExpMonth"))
	CCExpYear=SQLClean(Request("CCExpYear"))
	CCExp2DigitYear=RIGHT(CCExpYear,2)
	CCV=SQLClean(Request("CCV"))
	sCheckNo=""

	IF CCExpMonth="" THEN CCExpMonth=0 
	IF CCExpYear="" THEN CCExpYear=0

  HQF = SQLClean(TRIM(Request("HQF")))	
	OLR = SQLClean(TRIM(Request("OLR")))
	LF = SQLClean(TRIM(Request("LF")))
	PA = SQLClean(TRIM(Request("PA")))
	IW = SQLClean(TRIM(Request("IW")))
	RF = SQLClean(TRIM(Request("RF")))
	PF = SQLClean(TRIM(Request("PF")))
	
	IF Action="Verify & Continue" THEN
		IF TRIM(CCFirstName)="" THEN 
			FirstNameError=true
			FieldErr = FieldErr +1
		END IF   
		IF TRIM(CCLastName)="" THEN 
			LastNameError=true
			FieldErr = FieldErr +1
		END IF   
		IF TRIM(CCAddress1)="" THEN 
			Address1Error=true
			FieldErr = FieldErr +1
		END IF   
		IF TRIM(CCCity)="" THEN 
			CityError=true
			FieldErr = FieldErr +1
		END IF   
		IF TRIM(CCState)="" THEN 
			StateError=true
			FieldErr = FieldErr +1
		END IF   
		IF TRIM(CCZipCode)="" THEN 
			ZipCodeError=true
			FieldErr = FieldErr +1
		END IF   
		IF TRIM(CCEmail)="" THEN 
			EmailError=true
			FieldErr = FieldErr +1
		END IF   
		IF TRIM(CCCardNumber)="" OR NOT(IsNumeric(CCCardNumber)) THEN 
			CardNumberError=true
			FieldErr = FieldErr +1
		END IF   
		IF TRIM(CCV)="" OR NOT(IsNumeric(CCV)) THEN 
			CCVError=true
			FieldErr = FieldErr +1
		END IF   
		IF TRIM(CCExpMonth)=0 THEN 
			ExpMonthError=true
			FieldErr = FieldErr +1
		END IF   

		IF TRIM(CCExpYear)=0 THEN 
			ExpYearError=true
			FieldErr = FieldErr +1
		END IF   

	END IF


'response.write("<br> In Verify - FieldErr="&FieldErr)

IF TestMode=true THEN
	'DisplayCurrentValues_ForTest
END IF


END SUB



' -------------------
  SUB DisplayDataForm
' -------------------

	IF FieldErr>0 THEN CCFormStatus="enabled"

	%>

	<FORM action="<%=ThisFileName%>" AUTOCOMPLETE = "off" method="post">   

	  <input type="hidden" name="sPayType" value="<%=sPayType%>">
	  <input type="hidden" name="sOrderNo" value="<%=sOrderNo%>">
	  <input type="hidden" name="CCAmount" value="<%=CCAmount%>">  
	  <input type="hidden" name="sLP" value="<%=sLP%>">  
	  <input type="hidden" name="ClubID" value="<%=ClubID%>">  
	  <input type="hidden" name="TID" value="<%=TID%>">  

	  <input type="hidden" name="HQF" value="<%=HQF%>">  
	  <input type="hidden" name="OLR" value="<%=OLR%>">  
	  <input type="hidden" name="LF" value="<%=LF%>">  
	  <input type="hidden" name="PA" value="<%=PA%>">  
	  <input type="hidden" name="IW" value="<%=IW%>">  
	  <input type="hidden" name="RF" value="<%=RF%>">  

	  <%

	     IF Action="Verify & Continue" AND FieldErr=0 THEN %>
			    	
		<input type="hidden" name="CCFirstName" value="<%=CCFirstName%>">
		<input type="hidden" name="CCLastName" value="<%=CCLastName%>">
		<input type="hidden" name="CCAddress1" value="<%=CCAddress1%>">
		<input type="hidden" name="CCCity" value="<%=CCCity%>">
		<input type="hidden" name="CCState" value="<%=CCState%>">
		<input type="hidden" name="CCZipCode" value="<%=CCZipCode%>">
		<input type="hidden" name="CCPhone" value="<%=CCPhone%>">
		<input type="hidden" name="CCEmail" value="<%=CCEmail%>">
		<input type="hidden" name="CCCardNumber" value="<%=CCCardNumber%>">
		<input type="hidden" name="CCV" value="<%=CCV%>">
		<input type="hidden" name="CCExpMonth" value="<%=CCExpMonth%>">
		<input type="hidden" name="CCExpYear" value="<%=CCExpYear%>"><%

	     END IF

	   %>





      <% ' -------  UPPER TABLE  -------- %>	
      <TABLE ALIGN="center" width=90%>

      <tr> 
	<td ALIGN="center" >
		<font size="4"><b>Payment Processing</b></font>
	</td>
      </tr>	


      <tr> 
	<td ALIGN="left" >
		<font size="1">
			<br>
			Complete the information below to make a payment for your event sanctioning.  Payment processing is managed on a secure server.  Your credit card is not stored locally on any USA Water Ski server.  If you are using Internet Explorer, you may obtain a ‘Security Report’ and review information on secure servers by clicking on the Padlock icon located at the top of this page to the right of the URL field. 
		</font>
	</td>
      </tr>	

    </TABLE><% 






    ' -------  CARDHOLDER NAME AND BILLING ADDRESS TABLE  -------- %>
    <br>
    <TABLE ALIGN="center" class="innertable" width=90%>   	

      <tr>
				<th colspan="4" align="center">
					<font size=<% =fontsize3 %> COlOR="#FFFFFF"><strong>Cardholder & Card Billing Address</strong></FONT>
				</th>
      </tr>
      <tr> 
				<td Width=80px ALIGN="right">
					<font size=<% =fontsize2 %>  COlOR="<% =textcolor1 %>"><b>First:</b>&nbsp;&nbsp</FONT>
				</td>
				<td align="left">
					<a title="First Name must match name on the credit card">
						<input type="text" <% =CCFormStatus %> name="CCFirstName" value= "<% =CCFirstName %>" MaxLength=20 size="24">
					</a><%
					IF FirstNameError THEN response.write("<font size=3  COlOR=red>*</font>") %>
				</td>
        <td align=right>
        	<font size=<% =fontsize2 %>  COlOR="<% =textcolor1 %>"><b>OrderNo:</b>&nbsp</FONT>
        </td>
				<td align=left Width=100px>
					<font size=<% =fontsize3 %>  COlOR="<% =textcolor2 %>">&nbsp;<%=sOrderNo%></FONT>
				</td>
      </tr> 	
      <tr>
				<td ALIGN="right"><FONT size=<% =fontsize2 %>  COlOR=<% =textcolor1 %>><b>Last:</b>&nbsp;&nbsp</FONT></td>
				<td align=left colspan=3>
					<a title="Last Name must match name on the credit card">
						<input type="text" <% =CCFormStatus %> name="CCLastName" value= "<% =CCLastName %>" MaxLength=25 size="29">
					</a><%
					IF LastNameError THEN response.write("<font size=3  COlOR=red>*</font>") %>
				</td>
      </tr>
      <tr>
				<td ALIGN="right">
					<font size=<% =fontsize2 %>  COlOR=<% =textcolor4 %>><b>Address:</b></FONT>
				</td>
				<td colspan=3>
					<input type="text" <% =CCFormStatus %> name="CCAddress1" value= "<% =CCAddress1 %>" MaxLength=30 size="30"><%
					IF Address1Error THEN response.write("<font size=3 COlOR=red>*</font>") %>
				</td>	
      </tr>
      <tr>
				<td ALIGN="right">
					<FONT size=<% =fontsize2 %>  COlOR=<% =textcolor4 %>>&nbsp;&nbsp;&nbsp;<b>City:</b></FONT>
				</td>
				<td align=left>
					<input type="text" <% =CCFormStatus %> name="CCCity" value= "<% =CCCity %>" MaxLength=20 size="24"><%
					IF CityError THEN response.write("<font size=3  COlOR=red>*</font>") %>
				</td>
				<td>
					<FONT size=<% =fontsize2 %>  COlOR=<% =textcolor4 %>><b>&nbsp;&nbsp;&nbsp;&nbsp;State:</b></FONT><%
					IF Action="Verify & Continue" THEN %>
							<input type="text" <% =CCFormStatus %> name="CCState" value= "<% =CCState %>" size="2"><%
							IF StateError THEN Response.write("<font size=3 COlOR=red>*</font>")
					ELSE 
							StateArray = Split(USStatesList2,",")  %>
							<select name="CCState" <% =CCFormStatus %>><%
			  			FOR kvar = 0 TO UBOUND(StateArray)
			    				IF TRIM(CCState) = TRIM(StateArray(kvar)) THEN
											response.write("<option value = """&CCState&""" SELECTED>"&CCState&"</option>")
			    				ELSE
											response.write("<option value = """&StateArray(kvar)&""">"&StateArray(kvar)&"</option>")
			    				END IF
			  			NEXT  %>
							</select><%
					END IF  %>
				</td>
				<td>	
					<FONT size=<% =fontsize2 %>  COlOR="<% =textcolor1 %>">&nbsp;&nbsp;&nbsp;<b>Zip:</b></FONT>
					<input type="text" <% =CCFormStatus %> name="CCZipCode" value= "<% =CCZipCode %>" MaxLength=6 size="8"><%
					IF ZipCodeError THEN Response.write("<font size=3 COlOR=red>*</font>") %>
				</td>	
      </tr>   
      <tr> 
				<td ALIGN="right"><FONT size=<% =fontsize2 %>  COlOR="<% =textcolor1 %>"><b>Phone:</b></FONT></td>
				<td colspan="1">
					<input type="text" <% =CCFormStatus %> name="CCPhone" value= "<% =CCPhone %>" MaxLength=10 size="10"><%
					IF PhoneError THEN response.write("<font size=3 COlOR=red>*</font>")  %>
				</td>	
				<td ALIGN="right">
					<FONT size=<% =fontsize2 %>  COlOR="<% =textcolor1 %>"><b>Email:</b></FONT>
				</td>
				<td colspan="1">
					<input type="text" <% =CCFormStatus %> name="CCEmail" value= "<% =CCEmail %>" MaxLength=35 size="35"><%
					IF EmailError THEN response.write("<font size=3 COlOR=red>*</font>") %>
				</td>	
      </tr>
   </table>       
   <% '---  Cell Divider Table ---- 
   
   
   
   
   '---  CREDIT CARD TABLE ---- 
   %>
   <br>
   <TABLE ALIGN="center" class="innertable" width=90% >   	
		<tr>
			<th colspan="8" align="center"><FONT size=<% =fontsize3 %> COlOR="#FFFFFF"><strong>Credit Card Information</strong></FONT></th>
   	</tr>
		<tr>	
			<td ALIGN="right">
				<FONT size=<% =fontsize2 %>  COlOR="<% =textcolor1 %>"><b>Card No:</b>
			</td>
			<td colspan=4>
				<a title="Enter your Visa, Mastercard or American Express credit card number.">
				<input type="text" MAXLENGTH=16 <% =CCFormStatus %> name="CCCardNumber" value= "<% =CCCardNumber %>" size="16">
				</a><%
				' --- Added check for number only to credit card number 12-11-2010 ---
				IF CardNumberError THEN response.write("<font size=3  COlOR=red>*</font>")  %>
				<FONT size=<%=fontsize2%>  COlOR="<% =textcolor1 %>"> (No hypens/spaces)</font>
			</td> 
			<td ALIGN="right">
				<FONT size=<% =fontsize2 %>  COlOR="<% =textcolor1 %>"><b>Exp (mo/yr):</b></FONT>
			</td>
			<td colspan=2><%
					SelectedMonth = CCExpMonth %>
					<select name="CCExpMonth" <% =CCFormStatus %>><%
		  		response.write("<option value = 0 >NA</option>")
		  		FOR iCounter = 1 TO 12 STEP 1
							CurVal = TRIM(iCounter)
							IF iCounter < 10 THEN
									CurVal = TRIM("0"&iCounter)
							END IF
							IF CCExpMonth = CurVal THEN
									response.write("<option value = """&CCExpMonth&""" SELECTED>"&CCExpMonth&"</option>")
							ELSE
									response.write("<option value = """&CurVal&""">"&CurVal&"</option>")
							END IF
		  		NEXT  %>
					</select><%   

					SelectedYear = CCExpYear %>
					<select name="CCExpYear" <% =CCFormStatus %>><%
		  		response.write("<option value = 0 >NA</option>")
		  		FOR iCounter = DatePart("yyyy",DATE) TO DatePart("yyyy",DATE) + 10 STEP 1
							CurVal = TRIM(iCounter)
							IF CCExpYear = CurVal THEN
									response.write("<option value = """&CCExpYear&""" SELECTED>"&CCExpYear&"</option>")
							ELSE
									response.write("<option value = """&CurVal&""">"&CurVal&"</option>")
							END IF
		  		NEXT  %>
					</select><%
					IF ExpMonthError THEN Response.write("<font size=3 COlOR=red>*</font>") %>
			</td>	
		</tr>
    <tr> 
			<td ALIGN="right">
				<FONT size=<% =fontsize2 %>  COlOR="<% =textcolor1 %>"><b>CCV</b></FONT>
			</td>
			<td ALIGN="left" colspan=4>
				<a title="Security Code is a 3 or 4 digit number located on the rear of your credit card">
					<input type="text" <% =CCFormStatus %> name="CCV" value= "<% =CCV %>" MaxLength=4 size="5">
				</a><%
				IF CCVError THEN response.write("<font size=3 COlOR=red>*</font>")  %>
        <FONT size=<% =fontsize2 %>  COlOR="<% =textcolor1 %>">(3 or 4-digit security code)</FONT>
			</td>
			<td ALIGN="right">
				<FONT size=<% =fontsize2 %>  COlOR="<% =textcolor1 %>"><b>Payment Amount:</b></FONT>
			</td>
			<td ALIGN="left" colspan=2>
				<font size=<% =fontsize3 %>  COlOR="<% =textcolor2 %>"><% =formatCurrency(CCAmount,2) %></font>
			</td>	
		</tr><%  


	IF FieldErr > 0 THEN %>
	  	<tr><td colspan="8" ALIGN="center"><font size=<% =fontsize2 %>  COlOR="red">* Indicates Required Missing or Incorrect Information</font></td></tr><%
	END IF %>


    <TR>	<% '--- Row from main table holding button table --- %>
			<TD colspan=8>
				<br>
				<% ' --- Table to hold buttons and images --- %>
				<table align=center width=95%>  
		  		<tr><%

						IF FieldErr = 0 AND Action<>"Modify Information" THEN %>
								<td WIDTH=50% ALIGN="center" colspan=2 style="border-style:none;">
									<input type="submit" Name="Action" style="width:13em" value="Make Payment Now">
								</td> 
								<td WIDTH=50% ALIGN="center" colspan=2 style="border-style:none;">
									<input type="submit" Name="Action" value="Modify Information" style="width:13em">   
								</td><%	
						ELSE %>
								<td colspan=2 style="border-style:none;">&nbsp</td>  
								<td ALIGN="center" colspan=2 style="border-style:none;">
										<input type="submit" style="width:13em" name="Action" value="Verify & Continue">
								</td>
								<td colspan=2 style="border-style:none;">&nbsp</td><%
						END IF 
						
						' --- Images of credit card types ---
						%>
						<td align=right style="border-style:none;"><a title="Visa"><img src="/rankings/images/logos/visa_card.gif"></a></td>
						<td align=right style="border-style:none;"><a title="MasterCard"><img src="/rankings/images/logos/master_card.gif"></a></td>
						<td align=right style="border-style:none;"><a title="American Express"><img src="/rankings/images/logos/amex_card.gif"></a></td>
		  		</tr>
				</table> <% ' --- Table for holding buttons and images --- %>
			</TD>
     </TR>
   </TABLE>   <% ' ---------  END OF TABLE FOR BUTTON POSITIONING ----------- %>

   </form> <%	'--- Bottom of form ---


END SUB






' ------------------------
  SUB DisplaySuccussForm
' ------------------------

' --- Resets these Session variables to prevent the card from being charged twice ---

Session("ClubID")=""
Session("TID")=""


%>
<form action="<% =ReturnOnSuccess %>" method="post">   
	  <input type="hidden" name="ClubID" value="<%=ClubID%>">
	  <input type="hidden" name="TID" value="<%=TID%>">
	  <input type="hidden" name="sOrderNo" value="<%=sOrderNo%>">
	  <input type="hidden" name="CCAmount" value="<%=CCAmount%>">  
	  <input type="hidden" name="sLP" value="<%=sLP%>">  
	<TABLE class="innertable" ALIGN="center" WIDTH=50% class="innertable">
	  <tr>
			<th ALIGN="center">
				<font size="4"  COlOR="#FFFFFF"><b>NOTICE</b></FONT>
			</th>
	  </tr>
	  <tr>
	    <td Align="center" style="border-style:none;">
				<br>
				<font size="3" ><b>Payment Was Successful</b></font>	
	 			<br><br> 
				<font size=<% =fontsize2 %>  COlOR=<% =textcolor4 %>>Your Order Number is <% =sOrderNo %></FONT>
	    </td>
	  </tr>
	  <tr>
			<td ALIGN="center" style="border-style:none;">
				<br>
				<input type="button" value="Close Window" onclick="window.close();">
				<br><br>
	     </td>			
	  </tr>
	</TABLE>
 </form>
<%


END SUB



' ------------------------
  SUB DisplayFailureForm
' ------------------------

%>

<form action="<%=ThisFileName%>" method="post">   
	<input type="hidden" name="CCAmount" value="<%=CCAmount%>">
	<input type="hidden" name="sOrderNo" value="<%=sOrderNo%>">
	<input type="hidden" name="sPayType" value="<%=sPayType%>">
	<input type="hidden" name="sLast4Card" value="<%=sLast4Card%>">
	<input type="hidden" name="sLP" value="<%=sLP%>">  
  <input type="hidden" name="ClubID" value="<%=ClubID%>">  
  <input type="hidden" name="TID" value="<%=TID%>">  

	  <input type="hidden" name="HQF" value="<%=HQF%>">  
	  <input type="hidden" name="OLR" value="<%=OLR%>">  
	  <input type="hidden" name="LF" value="<%=LF%>">  
	  <input type="hidden" name="PA" value="<%=PA%>">  
	  <input type="hidden" name="IW" value="<%=IW%>">  
	  <input type="hidden" name="RF" value="<%=RF%>">  


	<input type="hidden" name="CCFirstName" value="<%=CCFirstName%>">
	<input type="hidden" name="CCLastName" value="<%=CCLastName%>">
	<input type="hidden" name="CCAddress1" value="<%=CCAddress1%>">
	<input type="hidden" name="CCCity" value="<%=CCCity%>">
	<input type="hidden" name="CCState" value="<%=CCState%>">
	<input type="hidden" name="CCZipCode" value="<%=CCZipCode%>">
	<input type="hidden" name="CCPhone" value="<%=CCPhone%>">
	<input type="hidden" name="CCEmail" value="<%=CCEmail%>">
	<input type="hidden" name="CCCardNumber" value="<%=CCCardNumber%>">
	<input type="hidden" name="CCV" value="<%=CCV%>">
	<input type="hidden" name="CCExpMonth" value="<%=CCExpMonth%>">
	<input type="hidden" name="CCExpYear" value="<%=CCExpYear%>">

	<TABLE class="innertable" ALIGN="center" WIDTH=60% BORDER="4" BGCOLOR=<% =TableColor1 %> >
		<tr>
			<th ALIGN="center" colspan=2>
				<font size="4"  COlOR="#FFFFFF"><b>IMPORTANT NOTICE</b></FONT>
			</th>
	  </tr>
	  <tr>
	    <td colspan=2 Align="center" style="border-style:none;">
	    	<br>
				<font size="4" ><b>Transaction Failed</b></font>	
	 			<br>
				<font size=<% =fontsize2 %>  COlOR=<% =textcolor4 %>><% = ResponseReasonText %></FONT>
				<font size=<% =fontsize2 %>  COlOR=<% =textcolor4 %>>Code: <% = AVSResultCode %></FONT>
	    </td>
	  </tr>
	  <tr>
	    <td width=50% ALIGN="center" style="border-style:none;">
			<br>
			<input type="submit" Name="Action" value="Change Card Info" style="width:11em">
			<br><br>
		</td>			
		<td width=50% ALIGN="center"  style="border-style:none;">
			<br>
			<input type="submit" name="Action" value="Return to Main" style="width:11em">
			<br><br>
	  </td>			
	 </tr>
	</TABLE>
</form>
<%

END SUB




' -----------------------------
   SUB DisplayTimeToCloseForm
' -----------------------------


%>
        <form action="<%=ThisFileName%>" method="post">   


	<TABLE ALIGN="center" WIDTH=60% BORDER="4" BGCOLOR=<% =TableColor1 %> >
	  <TR>
		<TD ALIGN="center" BGCOLOR="<% =HeadColor1 %>"><font size="4"  COlOR="#FFFFFF"><b>IMPORTANT NOTICE</b></FONT></TD>
	  </TR>

	  <TR>
	  <TD>


		
	<TABLE ALIGN="center" WIDTH=100% BORDER="0" BGCOLOR=<% =TableColor1 %> >
	  <tr>
	    <td colspan=2 Align="center" ><br>
		<font size="4" ><b>Transaction Already Completed</b></font>	
	    </td>

	  </tr>
	  <tr>
	    <td width=50% ALIGN="center" style="border-style:none;">
		<br>
		<input type="submit" Name="Action" value="Change Card Info">
		<input type="hidden" name="cstat" value="enter">
		<br><br>
	    </td>			
		  <input type="hidden" name="ClubID" value="<%=ClubID%>">  
		  <input type="hidden" name="TID" value="<%=TID%>">  

	  <input type="hidden" name="HQF" value="<%=HQF%>">  
	  <input type="hidden" name="OLR" value="<%=OLR%>">  
	  <input type="hidden" name="LF" value="<%=LF%>">  
	  <input type="hidden" name="PA" value="<%=PA%>">  
	  <input type="hidden" name="IW" value="<%=IW%>">  
	  <input type="hidden" name="RF" value="<%=RF%>">  


	    <input type="hidden" name="CCFirstName" value="<%=CCFirstName%>">
	    <input type="hidden" name="CCLastName" value="<%=CCLastName%>">
	    <input type="hidden" name="CCAddress1" value="<%=CCAddress1%>">
	    <input type="hidden" name="CCCity" value="<%=CCCity%>">
	    <input type="hidden" name="CCState" value="<%=CCState%>">
	    <input type="hidden" name="CCZipCode" value="<%=CCZipCode%>">
	    <input type="hidden" name="CCEmail" value="<%=CCEmail%>">
	    <input type="hidden" name="CCPhone" value="<%=CCPhone%>">
	    <input type="hidden" name="CCCardNumber" value="<%=CCCardNumber%>">
	    <input type="hidden" name="CCV" value="<%=CCV%>">
	    <input type="hidden" name="CCExpMonth" value="<%=CCExpMonth%>">
	    <input type="hidden" name="CCExpYear" value="<%=CCExpYear%>">
	    <input type="hidden" name="CCAmount" value="<%=CCAmount%>">
	    <input type="hidden" name="sOrderNo" value="<%=sOrderNo%>">
	    <input type="hidden" name="sPayType" value="<%=sPayType%>">
	    <input type="hidden" name="sLast4Card" value="<%=sLast4Card%>">

 	    <input type="hidden" name="sLP" value="<%=sLP%>">  

	    <td width=50% ALIGN="center"  style="border-style:none;">
	    	<br>
		<input type="submit" name="Action" value="Return to Main">
	    	<br><br>
	    </td>			
	  </TR>
	</TABLE>


	</TABLE>

    </form>
<%




END SUB





' ---------------------------
  SUB ProcessCardTransaction
' ---------------------------


	' --------------------------------------------------------------------------------------
	' ----------------  Prepare to Charge Credit Card  -------------------------------------
	' --------------------------------------------------------------------------------------


	Dim SendString

	SendString = "x_description=USA Waterski Sanction Fees&"
	SendString = SendString & "x_delim_data=TRUE&"
	SendString = SendString & "x_delim_char=|&"
	SendString = SendString & "x_relay_response=FALSE&"
	SendString = SendString & "x_login=4G998wGB9vy&"
	SendString = SendString & "x_tran_key=4A2edy4x9Y88TWEc&" 
	SendString = SendString & "x_type=AUTH_CAPTURE&"
	SendString = SendString & "x_method=CC&"
	'IF sLP="Test" THEN SendString = SendString & "x_test_request=TRUE & _"
	SendString = SendString & "x_amount=" & CCAmount &"&"
	SendString = SendString & "x_card_num=" & CCCardNumber &"&"
	SendString = SendString & "x_exp_date=" & CCExpMonth & "/" & CCExp2DigitYear &"&"
	SendString = SendString & "x_card_code=" & CCV  & "&"
	SendString = SendString & "x_first_name=" & CCFirstName & "&"
	SendString = SendString & "x_last_name=" & CCLastName & "&  "
	SendString = SendString & "x_address=" & CCAddress1 & "&"
	SendString = SendString & "x_city=" & CCCity & "&"
	SendString = SendString & "x_state=" & CCState & "&"
	SendString = SendString & "x_zip=" & CCZipCode


	'IF TestMode=true THEN 
	IF sLP="Test" THEN
		' --- Test values 
		ResponseCode = 1
		ResponseSubCode = "0"
		ResponseReasonCode = "X"
		ResponseReasonText = "This is the Text for ReasonText"
		ApprovalCode = "ABC123"
		AVSResultCode = "Y"
		TransactionID = "ABCDEF_12345_XYZ"

	ELSE
		' -- SUB added 3/5/2018 to send Mark and email to help determine why error is occuring --		
		' SendMarkDebugger SendString

		' -----------------------------------------
		' --- Post to the credit card processor ---
		' -----------------------------------------
		' -- Original changed to new one on 3/5/2018 --
		' -- Set objSendString = Server.CreateObject("MSXML2.ServerXMLHTTP")
		Set objSendString = Server.CreateObject ("MSXML2.XMLHTTP.6.0")
		objSendString.Open "POST", "https://secure.authorize.net/gateway/transact.dll", False
		objSendString.Send(SendString)
		ResponseString = objSendString.ResponseText


		
		
		Dim aryRString
		aryRString = Split(ResponseString, "|")

		IF INSTR(aryRString(0),"The merchant login ID or password is invalid or the account is inactive") THEN 
				Response.write("<center><red><br>The merchant LOGIN is INVALID<br>Contact Mark Crone</red></center>")
			
				Display_SendString="Y"
				IF Display_SendString="Y" THEN
						Response.write("<br><br>"&ResponseString)
						Response.write("<br><br>SendString = "&SendString)
				END IF	
		END IF
		'response.end

		

 		' =================================================================================================
		' --- GATEWAY RESPONSE API 
 
		' --- POSITION	ARRAY	FIELD_NAME		DESCRIPTION
		' ---    1	  0	ResponseCode 		1 = Approved, 2 = Declined, 3 = Error
		' ---    2	  1	ResponseSubCode 	Code used by the system for internal transaction tracking - Values ???
		' ---    3 	  2	ResponseReasonCode 	Code representing more details about the the result of the transaction
		' ---						  1=Approved
		' ---						  2=Declined
		' ---						  6=Card Number Invalid   
		' ---						  7=Expiration Date Invalid
		' ---						  8=Card Has Expired
		' ---						  11=Duplicate transaction within last two minutes
		' ---						  17=Card Type not allowed								   
		' ---    4  	  3	ResponseReasonText 	Brief description of the result which corresponds to the ResponseReasonCode - OK to Echo to customer
		' --- 	 5	  4	ApprovalCode		The six (6) digit alphanumeric authorization or approval code 
		' --- 	 6	  5	AVSResultCode	 	Result of Address Verification System checks 
		' ---						  A = Address matches, Zip does NOT
		' ---						  B = Address information not provided
		' ---						  E = AVS error
		' ---						  N = No Match on Address or Zip
		' ---						  X = Address and 9-digit Zip match
		' ---						  Y = Address and 5-digit Zip match
		' ---						  Z = 5-digit Zip matches, Address does not
		' --- 	 7	  6	TransactionID	 	Number identifies transaction in the system. Can be used for voiding, crediting or capturing 
		' --- 	 39	  38	CardCode	 	Results of CardCode verification M=Match, N=No Match, P= Not Processed, S=Should have been present, U=Issuer unable to process request
		' --- 	 40	  39	CCV	 		Results of CCV verification 0=Erroneous Data; 1=Failed; 2or8orAorB=Passed; 3=Incomplete; 4=System Error; 7=Failed Other


 		' =================================================================================================


		ResponseCode = aryRString(0)
		ResponseSubCode = aryRString(1)
		ResponseReasonCode = aryRString(2)
		ResponseReasonText = aryRString(3)
		ApprovalCode = aryRString(4)
		AVSResultCode = aryRString(5)
		TransactionID = aryRString(6)

'		ss_result = ""
'		ssl_result_message = "System in Test Mode"
'		ssl_txn_id = ""
'		ssl_approval_code = "TEST"
'		ssl_cvv2_response = ""
'		ssl_avs_response = ""


		set objSendString = Nothing

		

	END IF

'response.write("<br>aryRString(1) = "&aryRString(1))
'response.write("<br>aryRString(2) = "&aryRString(2))
'response.write("<br>aryRString(3) = "&aryRString(3))
'response.write("<br>aryRString(4) = "&aryRString(4))
'response.write("<br>aryRString(5) = "&aryRString(5))
'response.write("<br>aryRString(6) = "&aryRString(6))


	IF ResponseCode = 1 THEN   	' --- SUCCESS ---
		
		' Send Email to USA Waterski and copy Mark Crone if SendMarkSanctionPaymentEmail="Y"
		SendPaymentConfirm
		
		' --- Write payment info to SQL table ---
		LogPaymentInTable

		' --- String used to pass NVP's back to sending page --	
		sSendingPage = ReturnOnSuccess & "?TID="&TID&"&ClubID="&ClubID&"&sOrderNo="&sOrderNo&"&sPayType=Card"
		DisplaySuccussForm

	ELSE 	' --- FAILURE ---

		' --- Write payment info to SQL table ---
		LogPaymentInTable

		' --- Diaplays Failure form ---
		DisplayFailureForm

	END IF


     	

'
'		set objWaCC = Nothing
'		cstat="success"


END SUB




' ------------------------------
   SUB DisplayTimeOutNotice
' ------------------------------


%>
<br><br>

<form action="<% =ReturnOnTimeOutError%>" method="post">

<TABLE class="innertable" ALIGN="CENTER" width=65%>
  <tr>
      <th><center><font color="#FFFFFF" size="4"><b>Important Notice !!</b></font></th>
  </tr>  

  <tr>
     <td VALIGN="top" align="center">
	<br>
	<font color="<% =TextColor1 %>" face="<% =font1 %>" size="3"><b><i><center>Your Payment Session Has Ended</center></i></b></font>
	<br><br>
	<font face="<% =font1 %>" size="<%=fontsize3%>">This message results from either: <br><br>1) Payment was made successfully<br> OR <br>2) Maximum time limit was reached.  <br><br>Check sanction system for payment status.  
	<br><br>
	<center>
	If you have any questions, please contact:
	<br>
	USA Water Ski - Competition Dept at 800-533-5972</b></font>
	<br><br>
	<input type="submit" value=" Continue "></center>
    </td>	
  </tr>
</TABLE>   

</form>
<% 


END SUB



' ---------------------------
    SUB SendPaymentConfirm
' ---------------------------

Dim sTName, sTCity, sTState

sSQL = "SELECT TournAppID, TName, TCity, TState FROM "&SanctionTableName
sSQL = sSQL + " WHERE TournAppID = '"&TID&"'"
SET rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, sConnectionToTRATable, 3, 1
IF NOT rs.eof THEN
		sTName = rs("TName")
		sTCity = rs("TCity")
		sTState = rs("TState")
ELSE
		sTName = "Not in Sanction System"
		sTCity = ""
		sTState = ""		
END IF


ecss = "<style type=text/css>"
ecss = ecss & " body { font-family: Arial, Helvetica, sans-serif; text-align:center;}"
ecss = ecss & " .outer {color:white; font-size:14px; background-color:#FFFFFF; text-align:center; min-width:320px; max-width:500px; height:500px; border:1px solid;}"
ecss = "</style>"

ebody = "<html>"
ebody = ebody & "<head>"
ebody = ebody & "<title>Notice of Payment</title>"
ebody = ebody & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
ebody = ebody & "</head>"
ebody = ebody & "<body bgcolor=""#FFFFFF"" text=""#000000"">"
ebody = ebody & "<div align=""center"">"


' ebody = ebody & "<TABLE BORDER=1 ALIGN=CENTER CELLPADDING=3 CELLSPACING=0 BGCOLOR="&TableColor1&" width=50% >"
ebody = ebody & "<TABLE align=center style=""min-width=320px; max-width:450px; background-color:"&TableColor1&"; padding:3px; border:1px solid;"">"
ebody = ebody & "<TR>"
ebody = ebody & "<TD BGCOLOR=red><center><font face="&font1&" color=#FFFFFF size=4><b>Notice To Card Holder & Accounting Dept</b></font></TD>"
ebody = ebody & "</TR>"
 
ebody = ebody & "<TR>"
ebody = ebody & "<TD VALIGN=top>"

ebody = ebody & "<table border=""0"" cellspacing=""0"" cellpadding=""3"" width=""100%"">"
ebody = ebody & "<tr>"

ebody = ebody & "<td align=center>"	
ebody = ebody & "<font face="&font1&" size=4><b>Credit Card Payment Made</b></font>"
ebody = ebody & "<br><br>"
ebody = ebody & "<img style=""width:109px;"" name=""USAWaterskiLogo"" src=""http://www.usawaterski.org/rankings/images/logos/usawslogo_no_sub.jpg"" alt=""USA Waterski"" >"
ebody = ebody & "<br><br>"
ebody = ebody & "<font face="&font1&" size=2><b>This message is a notification that a credit card payment has been made in conjunction with the USA WaterSki Event Sanctioning system.</b></font>"

ebody = ebody & "<br><br>"

ebody = ebody & "<font face="&font1&" size=2><b>TOURNAMENT INFORMATION</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "<font color="&TextColor2&" face="&font1&" size=3>"&sTName&"</font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>City/ST: </b></font><font color="&TextColor2&" face="&font1&" size=2>"&sTCity&", "&sTState&"</font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>Sanction #: </b></font><font color="&TextColor2&" face="&font1&" size=2>"&TID&"</font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>ClubID/MemberID: </b></font><font color="&TextColor2&" face="&font1&" size=2>"&ClubID&"</font>"

ebody = ebody & "<br><br>"
ebody = ebody & "<font face="&font1&" size=2><b>PAYMENT INFORMATION</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>Cardholder: </b></font><font color="&TextColor2&" face="&font1&" size=2>"&CCFirstName&" "&CCLastName&"</font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>Order No: </b></font><font color="&TextColor2&" face="&font1&" size=2>"&sOrderNo&"</font>"
ebody = ebody & "<br>"
' ebody = ebody & "<font face="&font1&" size=2><b>Credit Card (Last 4): </b></font><font color="&TextColor2&" face="&font1&" size=2>"&sLast4Card&"</font>"
' ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>Amount: </b></font><font color="&TextColor2&" face="&font1&" size=2>"&formatCurrency(CCAmount,2)&"</font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>Date: </b></font><font color="&TextColor2&" face="&font1&" size=2>"&DATE&"</font></b>"

ebody = ebody & "<br><br>"
ebody = ebody & "<font face="&font1&" size=2><b>FEE SCHEDULE</b></font>"

IF CDbl(HQF)>CDbl(0) THEN
		ebody = ebody & "<br>"
		ebody = ebody & "<font face="&font1&" size=2><b>Sanction Fee: </b></font><font color="&TextColor2&" face="&font1&" size=2>"&formatCurrency(HQF,2)&"</font></b>"
END IF

IF CDbl(OLR)>CDbl(0) THEN
		ebody = ebody & "<br>"
		ebody = ebody & "<font face="&font1&" size=2><b>Online Registration Fee: </b></font><font color="&TextColor2&" face="&font1&" size=2>"&formatCurrency(OLR,2)&"</font></b>"
END IF

IF CDbl(LF)>CDbl(0) THEN
		ebody = ebody & "<br>"
		ebody = ebody & "<font face="&font1&" size=2><b>Late Fee: </b></font><font color="&TextColor2&" face="&font1&" size=2>"&formatCurrency(LF,2)&"</font></b>"
END IF

IF CDbl(PA)>CDbl(0) THEN
		ebody = ebody & "<br>"
		ebody = ebody & "<font face="&font1&" size=2><b>Pan Am Fee: </b></font><font color="&TextColor2&" face="&font1&" size=2>"&formatCurrency(PA,2)&"</font></b>"
END IF

IF CDbl(IW)>CDbl(0) THEN
		ebody = ebody & "<br>"
		ebody = ebody & "<font face="&font1&" size=2><b>IWSF Fee: </b></font><font color="&TextColor2&" face="&font1&" size=2>"&formatCurrency(IW,2)&"</font></b>"
END IF

IF CDbl(RF)>CDbl(0) THEN
		ebody = ebody & "<br>"
		ebody = ebody & "<font face="&font1&" size=2><b>Region Fee: </b></font><font color="&TextColor2&" face="&font1&" size=2>"&formatCurrency(RF,2)&"</font></b>"
END IF

' IF CDbl(PF)>CDbl(0) THEN
' 		ebody = ebody & "<br>"
' 		ebody = ebody & "<font face="&font1&" size=2><b>Practice Fee: </b></font><font color="&TextColor2&" face="&font1&" size=2>"&formatCurrency(PF,2)&"</font></b>"
' END IF


ebody = ebody & "<br><br>"
ebody = ebody & "<font face="&font1&" size=2>For questions about this sanction, contact your region EVP,"
ebody = ebody & "<br>or email <a title='Send Email to USA WaterSki competition dept' href='mailto:competition@usawaterski.org?subject=Sanction "&TID&" - "&sTName&"'>USA WaterSki Competition Dept</a></font>"
ebody = ebody & "<br><br>"
ebody = ebody & "<font face="&font1&" size=2>Transactions except for card data are stored in the USA WaterSki site.</font></b><br>"

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

sAddressList = "onlinepayments@usawaterski.org"
IF TRIM(CCEmail) <> "" THEN
		sAddressList = sAddressList & "; "& CCEmail
END IF
eMailTo = sAddressList

SendMarkSanctionPaymentEmail = "Y"
IF SendMarkSanctionPaymentEmail = "Y" THEN
		eMailBCC = marksemailaddress
END IF
IF SendMeisSanctionPaymentEmail = "Y" THEN
		eMailBCC = eMailBCC& "; "& meisemail
END IF

eMailFrom = "competition@usawaterski.org"
eMailSubj = "USA Water Ski - Sanction System Payment Notification - "&CCFirstName&" "&CCLastName&" - TourID: "&Session("TID")
eMailBody = ebody	

DisplayValues="N"
IF DisplayValues="Y" THEN
		response.write("<br>marksemailaddress= "&marksemailaddress)
		response.write("<br>eMailTo= "&eMailTo)
		response.write("<br>eMailFrom= "&eMailFrom)
		response.write("<br>eMailCC= "&eMailCC)
		response.write("<br>eMailBCC= "&eMailBCC)				

		response.write("<br>eMailSubj= "&eMailSubj)						
		response.write("<br><br>eMailBody= "&eMailBody)				
		'response.end
END IF


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
objMessage.Send 
set objMessage = Nothing


MarkDebug("USAWS SkipJack Sanctions Payment = "&CCAmount&" - TID = "&TID&" - Member = "&ClubID&" ")

END SUB




' ---------------------------
    SUB SendRefundConfirm
' ---------------------------


ebody = "<html>"
ebody = ebody & "<head>"
ebody = ebody & "<title>Notice of Refund</title>"
ebody = ebody & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
ebody = ebody & "</head>"
ebody = ebody & "<body bgcolor=""#FFFFFF"" text=""#000000"">"
ebody = ebody & "<div align=""center"">"


ebody = ebody & "<TABLE BORDER=4 ALIGN=CENTER CELLPADDING=3 CELLSPACING=0 BGCOLOR="&TableColor1&" width=50% >"
ebody = ebody & "<TR>"
ebody = ebody & "<TD BGCOLOR=red><center><font face="&font1&" color=#FFFFFF size=4><b>Notice To Accounting</b></font></TD>"
ebody = ebody & "</TR>"
 
ebody = ebody & "<TR>"
ebody = ebody & "<TD VALIGN=top>"


ebody = ebody & "<table border=""0"" cellspacing=""0"" cellpadding=""3"" width=""100%"">"
ebody = ebody & "<tr>"


ebody = ebody & "<td Align=center>"	

ebody = ebody & "<font face="&font1&" size=4><b>A Refund Has Been Made</b></font>"
ebody = ebody & "<br><br>"
ebody = ebody & "<font face="&font1&" size=2><b>This message is a notification that a refund has been credited to this on-line event registration account.</b></font>"

ebody = ebody & "<br><br>"

ebody = ebody & "<font face="&font1&" size=2><b>MemberID = </b></font><font color="&TextColor2&" face="&font1&" size=2>"&ClubID&"</font>"
ebody = ebody & "<br>"

ebody = ebody & "<font face="&font1&" size=2><b>Cardholder = </b></font><font color="&TextColor2&" face="&font1&" size=2>"&CCFirstName&" "&CCLastName&"</font>"
ebody = ebody & "<br>"

ebody = ebody & "<font face="&font1&" size=2><b>Order No = </b></font><font color="&TextColor2&" face="&font1&" size=2>"&sOrderNo&"</font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>Amount = </b></font><font color="&TextColor2&" face="&font1&" size=2>"&formatCurrency(CCAmount,2)&"</font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>Date = </b></font><font color="&TextColor2&" face="&font1&" size=2>"&DATE&"</font></b><br>"

ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2>All transaction data is stored in tables on HQ rankings site.</font></b><br>"

ebody = ebody & "</center>"
ebody = ebody & "<br>"
ebody = ebody & "</td></tr>"

ebody = ebody & "</TABLE>"

ebody = ebody & "</TD></TR>"
ebody = ebody & "</TABLE>"


CCAmount=formatCurrency(CCAmount,2)


sAddressList = "onlinepayments@usawaterski.org"

eMailTo = sAddressList
eMailBCC = marksemailaddress
eMailFrom = "competition@usawaterski.org"
eMailSubj = "USA Water Ski - Sanction System REFUND Notification - "&CCFirstName&" "&CCLastName&" - TourID: "&Session("TID")
eMailBody = ebody	

'response.write("<br>marksemailaddress= "&marksemailaddress)
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
objMessage.Send 
set objMessage = Nothing



END SUB





' ------------------------------------
  SUB SendMarkDebugger (Body)
' ------------------------------------


Dim eMailSubj, eMailFrom, eMailTo, eMailCC, eMailBCC, eMailBody


sAddressList = "cronemarka@gmail.com"
IF TRIM(CCEmail) <> "" THEN
		sAddressList = sAddressList & "; "& CCEmail
END IF
eMailTo = sAddressList
eMailCC = ""
eMailBCC = ""
' SendMarkSanctionPaymentEmail = "Y"
' IF SendMarkSanctionPaymentEmail = "Y" THEN
' 		eMailBCC = marksemailaddress
' END IF
' IF SendMeisSanctionPaymentEmail = "Y" THEN
' 		eMailBCC = eMailBCC& "; "& meisemail
' END IF

eMailFrom = "competition@usawaterski.org"
eMailSubj = "USA Water Ski - Sanction System Payment DEBUGGER - "&CCFirstName&" "&CCLastName&" - TourID: "&Session("TID")
eMailBody = "<br><br><h2>The following information is for debugging only"	
eMailBody = "<br>SendString = " & Body

DisplayValues="N"
IF DisplayValues="Y" THEN
		response.write("<br>marksemailaddress= "&marksemailaddress)
		response.write("<br>eMailTo= "&eMailTo)
		response.write("<br>eMailFrom= "&eMailFrom)
		response.write("<br>eMailCC= "&eMailCC)
		response.write("<br>eMailBCC= "&eMailBCC)				

		response.write("<br>eMailSubj= "&eMailSubj)						
		response.write("<br><br>eMailBody= "&eMailBody)				
		'response.end
END IF


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
objMessage.Send 
set objMessage = Nothing


MarkDebug("USAWS SkipJack Sanctions Payment = "&CCAmount&" - TID = "&TID&" - Member = "&ClubID&" ")

END SUB




 
' -------------------------------------------------
   FUNCTION GetValue(strSearch, valName, delimeter)
' -------------------------------------------------

            'This function takes the value of strString and seraches for name=value based on the delimeter.

            Dim arySearch

            'split the items into an array
            arySearch = Split(strSearch, delimeter)

            For i = 0 to UBound(arySearch)

                        Dim aryNameVal
                        aryNameVal = Split(arySearch(i), "=")                  
                        if UBound(aryNameVal) = -1 then
                                    GetValue = ""
                        else
                                    If lcase(aryNameVal(0)) = lcase(valName) then
                                                If ubound(aryNameVal) < 1 then
                                                            GetValue = ""
                                                Else
                                                            GetValue = aryNameVal(1)
                                                            Exit For 
                                                End If
                                    End If
                        end if
            Next

End Function




' ------------------------------------------------------------
  SUB LoadDropDown (DefaultNum, MinNum, MaxNum, StepNum)
' ------------------------------------------------------------

Dim iCounter

DefaultNum = Cint(DefaultNum)

response.write("<option value = 0 >NA</option>")

FOR iCounter = MinNum TO MaxNum STEP StepNum
	IF iCounter = DefaultNum THEN
		response.write("<option value = """&iCounter&""" SELECTED>"&iCounter&"</option>")
	ELSE
		response.write("<option value = """&iCounter&""">"&iCounter&"</option>")
	END IF
NEXT



END SUB






 
%>

  
 
