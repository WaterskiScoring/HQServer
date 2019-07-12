<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/includes/func.asp"-->
<!--#include virtual="/rankings/tools_registration16.asp"-->
<%


If Request.ServerVariables("HTTPS") = "off" then
	URL = "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("SCRIPT_NAME") 
	If Request.QueryString <> "" then
		URL = URL & "?" & Request.QueryString
	End If
	Response.Redirect URL
	Response.End
End If



' --- Defines file names throughout this program ---
Dim ThisFileName, RegFileName
ThisFileName="/rankings/CCReg2012.asp"
RegFileName="Registration16.asp"

marksemail="cronemarka@gmail.com"

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






TestMode = false	' --- Testing mode set to true ---
sPayType="Card"		' --- Always card from this module ---
ReturnOnCancel="http://usawaterski.org/"				' --- Defines redirection when the member abandons the payment processor ---
ReturnOnSuccess="Test_PaymentLaunch.asp"	' --- Defines redirection when transaction is a success ---
ReturnOnTimeOutError = RankPath&"/defaultHQ.asp" 



' --- Read primary variables ---
Action = Request("Action")

IF TRIM(Request("sMemberID"))<>"" THEN
		sMemberID=TRIM(Request("sMemberID"))
		sTourID=TRIM(Request("sTourID"))
		Session("sMemberID")=TRIM(Request("sMemberID"))
		Session("sTourID")=TRIM(Request("sTourID"))
ELSE
		sMemberID = Session("sMemberID")
		sTourID = Session("sTourID")
END IF

'sMemberID = "000001151"
'sTourID = "13S999"


' --- Determine the level of the admin user if applicable ---
adminmenulevel = Session("adminmenulevel")
IF TRIM(adminmenulevel) = "" THEN adminmenulevel = "1"




WriteIndexPageHeader


' ------------------------------------------------------------------------------------------------------
' - Tests whether the system remembers the person's name.  If not, then display the timeout screen.  ---
' ------------------------------------------------------------------------------------------------------

'response.write("<br>Line81 in CC sMemberID="&sMemberID)		
'response.write("<br>Sesssion sMemberID="&Session("sMemberID"))	

'response.write("sTourID = "&sTourID)
'response.end



IF TRIM(sMemberID) = "" THEN
		Action= "displaytimeout"

ELSE

	IF Action="new" THEN	' --- First time through CCSJ_Sanctions.asp --- 
			sOrderNo=TRIM(SQLClean(Request("sOrderNo")))
			CCAmount = TRIM(SQLClean(Request("CCAmount")))
			sLP = Request("sLP")

			Action = "Modify Information"
	
			set rsMemb=Server.CreateObject("ADODB.recordset")
			sSQL = "SELECT TOP 1 * FROM "&MemberTableName
			sSQL = sSQL + " WHERE PersonIDwithCheckDigit = "&sqlclean(sMemberID)
			rsMemb.open sSQL, sConnectionToTRATable, 3, 1

			IF TRIM(Request("ppf"))="1" THEN
					CCFirstName=rsMemb("FirstName")
					CCLastName=rsMemb("LastName")
					CCAddress1=rsMemb("Address1")
					CCCity=rsMemb("City")
					CCState=rsMemb("State")
					CCZipCode=rsMemb("Zip")
					CCEmail=rsMemb("Email")
					CCPhone=rsMemb("Phone")
			END IF
			IF CCAmount="" THEN CCAmount=0.00
			rsMemb.close

	ELSE
			sOrderNo=TRIM(SQLClean(Request("sOrderNo")))
			CCAmount = TRIM(SQLClean(Request("CCAmount")))
			sLP = Request("sLP")
			CCFirstName=SQLClean(Request("CCFirstName"))
			CCLastName=SQLClean(Request("CCLastName"))

			' --- Reads remaining variables from the form ---
			ReadAndValidateFormVariables
	END IF


END IF



' --- This section for testing and debugging ---
'sLP = "Test"
IF sLP="Test" THEN
		response.write("<br>TOP OF PROGRAM")
		response.write("<br>Action="&Action)
		response.write("<br>sMemberID="&sMemberID)
		response.write("<br>TID="&TID)
		response.write("<br>sOrderNo="&sOrderNo)
		response.write("<br>CCAmount="&CCAmount)
		response.write("<br>sLP="&sLP)
END IF

'response.write("<br>Action="&Action)

' --- Primary branching performed here ---
SELECT CASE Action

   CASE "displaytimeout"
				DisplayTimeOutNotice

  	CASE "Return To Main Menu"		
				Response.Redirect(ReturnOnCancel)

   	CASE "Verify & Continue"			' --- Displays the form that captures name, address, credit cards etc ---
				CCFormStatus="disabled"
				DisplayDataForm

   CASE "Modify Information"			' --- Displays the form that captures name, address, credit cards etc ---
				CCFormStatus="enabled"
				DisplayDataForm

   CASE "Make Payment Now"
				IF Session("sMemberID")<>"" AND Session("sTourID")<>"" THEN		
							' --- Runs the card info through the Wauchovia Charge System ---
							ProcessCardTransaction
							' --- Write payment info to SQL table ---
							'			LogPaymentInTable

				ELSE
							DisplayTimeOutNotice
				END IF

		CASE "Finalize & Record Entry"

	  		sSendingPage = "http://usawaterski.org/rankings/"&RegFileName&"?sTourID="&sTourID&"&sMemberID="&sMemberID&"&sOrderNo="&sOrderNo&"&nav=7&sPayType=Card"
				Response.Redirect(sSendingPage)

		CASE "DisplayFailureForm"		' --- For testing only ---
				DisplayFailureForm

		CASE "DisplaySuccussForm" 		' --- For testing only ---
				DisplaySuccussForm

END SELECT


WriteIndexPageFooter


' ================================================================================
' ================================================================================
' ====================  END OF MAIN CODE HERE  ===================================
' ================================================================================
' ================================================================================




' ----------------------
   SUB LogPaymentInTable
' ----------------------

	' ------------------------------------------
	' ---- Log all PAYMENTS into the table  ----
	' ------------------------------------------

' --- Codes from Wauchovia ---
'	ResponseCode = aryRString(0)
' ResponseSubCode = aryRString(1)
' ResponseReasonCode = aryRString(2)
' ResponseReasonText = aryRString(3)
' ApprovalCode = aryRString(4)
' AVSResultCode = aryRString(5)
' TransactionID = aryRString(6)

sConvertedResponseCode = ResponseCode
IF ResponseCode=1 THEN sConvertedResponseCode=0		' --- Makes compatible with former conventions for successful transaction ---
sResponseReasonText=LEFT(ResponseReasonText,99)
sTransactionID=LEFT(TransactionID,39)

	DateNow = Date
	OpenCon

	sSQL = "INSERT INTO "&RegPaymentTableName
	sSQL = sSQL + " (MemberID, TourID, FirstName, LastName, Address1, City, State, ZipCode, Email"

	' -- Removed 12-15-2018 -- MAC
	' sSQL = sSQL + ", Last4Card, ExpMonth, ExpYear"
	sSQL = sSQL + ", Amount, OrderNo, Result, Message, TXN_ID"
	sSQL = sSQL + ", Apvl_Code, CVV2_Resp, AVS_Resp, TransDate, PayType, PayStatus)"

	sSQL = sSQL + " VALUES ('"&sMemberID&"', '"&sTourID&"', '"&CCFirstName&"', '"&CCLastName&"', '"&CCAddress1&"', '"&CCCity&"', '"&CCState&"'"
	sSQL = sSQL + ", '"&CCZipCode&"', '"&CCEmail&"'"
	
		' -- Removed 12-15-2018 -- MAC
	' sSQL = sSQL + ", '"&sLast4Card&"', '"&CCExpMonth&"', '"&CCExpYear&"'"
	sSQL = sSQL + ", '"&CCAmount&"', '"&sOrderNo&"', '"&sConvertedResponseCode&"', '"&sResponseReasonText&"', '"&sTransactionID&"'"
	sSQL = sSQL + ", '"&ApprovalCode&"', '"&resp_CVV2&"', '"&AVSResultCode&"', '"&DateNow&"', '"&sPayType&"', '"&sPayStatus&"')"

'response.write("<br>"&sSQL)
'response.end
	con.execute(sSQL)
'response.end

	closecon

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

	' --- Form field validation ---	
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
	  <input type="hidden" name="sMemberID" value="<%=sMemberID%>">  
	  <input type="hidden" name="sTourID" value="<%=sTourID%>">  
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



		' -------  UPPER TABLE  -------- 
		%>	
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
					<input type="text" <% =CCFormStatus %> name="CCZipCode" value= "<% =CCZipCode %>" MaxLength=6 size="10"><%
					IF ZipCodeError THEN Response.write("<font size=3 COlOR=red>*</font>") %>
				</td>	
      </tr>   
      <tr> 
				<td ALIGN="right"><FONT size=<% =fontsize2 %>  COlOR="<% =textcolor1 %>"><b>Phone:</b></FONT></td>
				<td colspan="1">
					<input type="text" <% =CCFormStatus %> name="CCPhone" value= "<% =CCPhone %>" MaxLength=10 size="14"><%
					IF PhoneError THEN response.write("<font size=3 COlOR=red>*</font>")  %>
				</td>	
				<td ALIGN="right">
					<FONT size=<% =fontsize2 %>  COlOR="<% =textcolor1 %>"><b>Email:</b></FONT>
				</td>
				<td colspan="1">
					<input type="text" <% =CCFormStatus %> name="CCEmail" value= "<% =CCEmail %>" MaxLength=35 size="38"><%
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

sSendingPage = "http://usawaterski.org/rankings/"&RegFileName&"?sTourID="&sTourID&"&sMemberID="&sMemberID&"&sOrderNo="&sOrderNo&"&nav=7&sPayType=Card"


%>
<form action=""<%=ThisFileName%>"" method="post">
	<input type="hidden" name="sTourID" value="<%=sTourID%>">  
	<input type="hidden" name="sOrderNo" value="<%=sOrderNo%>">
	<input type="hidden" name="sMemberID" value="<%=sMemberID%>">
	<input type="hidden" name="nav" value="7">
	<input type="hidden" name="sPayType" value="Card">
	<input type="hidden" name="sPayType" value="Card">	

	<TABLE ALIGN="center" WIDTH=50% class="innertable">
	  <tr>
			<th ALIGN="center" BGCOLOR="<% =HeadColor1 %>"><font size="4"  COlOR="#FFFFFF"><b>NOTICE</b></FONT></td>
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
				<input type="submit" name="Action" value="Finalize & Record Entry">
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
	    <input type="hidden" name="CCAmount" value="<%=CCAmount%>">
	    <input type="hidden" name="sOrderNo" value="<%=sOrderNo%>">
	    <input type="hidden" name="sPayType" value="<%=sPayType%>">
	    <input type="hidden" name="sLast4Card" value="<%=sLast4Card%>">

		
	<TABLE align="center" WIDTH=60% class="innertable" >
	  <tr>
			<th align="center" colspan=2>
				<font size="4"  COlOR="#FFFFFF"><b>IMPORTANT NOTICE</b></FONT>
			</th>
	  </tr>
	  <tr>
	    <td colspan=2 Align="center" style="border-style:none;"><br>
				<font size="4" ><b>Transaction Failed</b></font>	
	 			<br>
				<font size=<% =fontsize2 %>  COlOR=<% =textcolor4 %>><% = Resp_Message %></FONT>
				<font size=<% =fontsize2 %>  COlOR=<% =textcolor4 %>><% = ResponseReasonText %></FONT>
	    </td>

	  </tr>
	  <tr>
	    <td width=50% ALIGN="center" style="border-style:none;">
				<br>
				<input type="submit" name="Action" value="Change Card Info" style="width:11em">
				<br><br>
			</td>			
      <td width=50% ALIGN="center"  style="border-style:none;">
				<br>
				<input type="submit" name="Action" value="Return to Main Menu" style="width:11em">
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

	  sSendingPage = "http://usawaterski.org/rankings/"&RegFileName&"?sTourID="&sTourID&"&sMemberID="&sMemberID&"&sOrderNo="&sOrderNo&"&nav=7&sPayType=Card"

%>
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
	    <form action="<%=ThisFileName%>?pvar=displayform&cstat=modify" method="post">   
	    <td width=50% ALIGN="center" style="border-style:none;">
		<br>
		<input type="submit" value="Change Card Info">
		<br><br>
		</td>			

	    <input type="hidden" name="pvar" value="displayform">
	    <input type="hidden" name="cstat" value="enter">
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


	    </form><%

		sSendingPage = RankPath&"/defaultHQ.asp" %>

	    <form action="<% =sSendingPage %>" method="post">   
	      <td width=50% ALIGN="center"  style="border-style:none;">
		<br>
		<input type="submit" value="Return to Main">
		<br><br>
	      </td>			
	    </form>
	  </TR>
	</TABLE>

	</TABLE><%




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

		'response.write("<br>SendString = "&SendString)
		'response.end

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

		set objSendString = Nothing

		

	END IF

'response.write("<br>aryRString(0) = "&aryRString(0))
'response.write("<br>aryRString(1) = "&aryRString(1))
'response.write("<br>aryRString(2) = "&aryRString(2))
'response.write("<br>aryRString(3) = "&aryRString(3))
'response.write("<br>aryRString(4) = "&aryRString(4))
'response.write("<br>aryRString(5) = "&aryRString(5))
'response.write("<br>aryRString(6) = "&aryRString(6))


	IF ResponseCode = 1 THEN   	' --- SUCCESS ---
		
			' Send Email to USA Waterski and copy Mark Crone if SendMarkEmail="Y"
			SendPaymentConfirm
		
			' --- Write payment info to SQL table ---
			LogPaymentInTable

			' --- String used to pass NVP's back to sending page --	
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

sSendingPage = RankPath&"/default.asp" 

%>
<br><br>

<TABLE BORDER="4" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="<% =TableColor1 %>" width=65%>
  <TR>
      <TD BGCOLOR="red"><center><font  color="#FFFFFF" size="4"><b>Important Notice !!</b></font></TD>
  </TR>  

  <TR>
     <TD VALIGN="top">
	<TABLE BORDER="0" VALIGN="top" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="<%=TableColor1%>" width=100%>
	   <tr>
	      <td VALIGN="top" ALIGN="center">
		<br>
		<font color="<% =TextColor1 %>" face="<% =font1 %>" size="5"><b><i>Your Session Timed Out</i></b></font>
		<br><br>
		<font face="<% =font1 %>" size="2">We are sorry for the inconvenience, but you may not register without starting over.</font>
		<br><br>	
		<font face="<% =font1 %>" size="2">The inactivity caused our server to reach the maximum time limit for maintaining your member and tournament selections.  Therefore, your registration was not recorded and the record you were working on is no longer active. Please try again.  
		<br><br>
		If you have any questions, please contact:
		<br>
		USA Water Ski - Competition Dept at 800-533-5972</b></font>
	    </td>
	  </tr>
	<tr>
	   <td align="center">
		<br>
		<form action="<% =sSendingPage %>" method="post">
		  <center><input type="submit" value=" Continue "></center>
		</form>
		</TABLE>
		   </td>	
	</tr>
    </TD>
  </TR>
</TABLE>   <% 


END SUB



' ---------------------------
    SUB SendPaymentConfirm
' ---------------------------



DefineTourVariables_New


sSQL = "SELECT TOP 1 FirstName, LastName"
sSQL = sSQL + " FROM "&MemberLiveTableName&" MT"
sSQL = sSQL + " WHERE PersonID = cast(right("&sqlclean(sMemberID)&",8) AS INTEGER)"
	
set rsMemb=Server.CreateObject("ADODB.recordset")
rsMemb.open sSQL, sConnectionToTRATable, 3, 1

sFullName = ""
IF NOT rsMemb.eof THEN
		sFullName = SQLClean(rsMemb("FirstName")&" "&rsMemb("LastName"))
END IF




USAWS_Logo ="http://www.usawaterski.org/rankings/images/logos/usawslogo_no_sub.jpg"

SQT = "'"
ecss = "<style type=text/css>"
ecss = ecss & " body { font-family: Arial, Helvetica, sans-serif; text-align:center;}"
ecss = ecss & " .outer {color:white; font-size:14pt; background-color:#FFFFFF; text-align:center; min-width:300px; max-width:500px; height:500px; border:1px solid;}"
ecss = ecss & " p {color:black; font-size:12pt; text-align:left; font-style:normal; position:relative;}"
ecss = ecss & " .pblue {color:blue; font-size:12pt; text-align:left;}"
ecss = ecss & " .pblack {color:#000000; font-size:12pt; text-align:left;}"
ecss = ecss & " .actionbutton {background-color:#006400; color:white; -moz-border-radius:15px; -webkit-border-radius:15px; border:5px solid; padding:5px;}"
ecss = ecss & " .psuedobuttoncellgreen {width:175px; text-align:center; -webkit-border-radius: 3px; -moz-border-radius: 3px; border-radius: 3px; background-color:#006400;}"
ecss = ecss & " .psuedobuttoncellred {width:175px; text-align:center; -webkit-border-radius: 3px; -moz-border-radius: 3px; border-radius: 3px; background-color:#DC143C;}"
ecss = ecss & " .psuedobuttongreen {width:100%; font-size:16pt; font-family:Helvetica, Arial, sans-serif; color:#ffffff; text-decoration:none; color:#ffffff; text-decoration:none; -webkit-border-radius:3px; -moz-border-radius:3px; border-radius:3px; padding:12px 0px; border: 1px solid #7FFF00; display: inline-block;}"
ecss = ecss & " .psuedobuttonred {width:100%; font-size:16pt; font-family:Helvetica, Arial, sans-serif; color:#ffffff; text-decoration:none; color:#ffffff; text-decoration:none; -webkit-border-radius:3px; -moz-border-radius:3px; border-radius:3px; padding:12px 0px; border: 1px solid #FFA500; display: inline-block;}"
ecss = ecss & " </style>"
ecss = ecss & " <meta name=format-detection content=telephone=no>"



ebody = "<html>"
ebody = ebody & "<head>"
ebody = ebody & "<title>Notice of Payment</title>"
ebody = ebody & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
ebody = ebody & "</head>"
ebody = ebody & "<body bgcolor=""#FFFFFF"" text=""#000000"">"

ebody = ebody & "<div class=outer style=""margin:0px 0px 0px 0px; padding:0px 0px 0px 0px;"">"


ebody = ebody & "<TABLE BORDER=4 ALIGN=CENTER CELLPADDING=3 CELLSPACING=0 BGCOLOR="&TableColor1&" width=95% >"
ebody = ebody & "<TR>"
ebody = ebody & "<TD BGCOLOR=red><center><font face="&font1&" color=#FFFFFF size=4><b>Notice To Accounting</b></font></TD>"
ebody = ebody & "</TR>"
 
ebody = ebody & "<TR>"
ebody = ebody & "<TD VALIGN=top>"


ebody = ebody & "<table border=""0"" cellspacing=""0"" cellpadding=""3"" width=""100%"">"
ebody = ebody & "<tr>"


ebody = ebody & "<td Align=center>"	

ebody = ebody & "<font face="&font1&" size=4><b>Credit Card Payment Has Been Made</b></font>"
ebody = ebody & "<br><br>"
ebody = ebody & "<font face="&font1&" size=2><b>This message is a notification that a payment has been made by credit card for on-line event registration.</b></font>"
ebody = ebody & "<br><br>"
ebody = ebody & "<font face="&font1&" size=2><b>TOURNAMENT INFORMATION</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "<font color="&TextColor2&" face="&font1&" size=3>"&sTourName&"</font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>Sanction: </b></font><font color="&TextColor2&" face="&font1&" size=2>"&sTourID&"</font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>City/ST: </b></font><font color="&TextColor2&" face="&font1&" size=2>"&sTourCity&", "&sTourState&"</font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>Event Date(s): </b></font><font color="&TextColor2&" face="&font1&" size=2>"&sTDateS&" - "&sTDateE&"</font>"


ebody = ebody & "<br><br>"
ebody = ebody & "<font face="&font1&" size=2><b>MEMBER INFORMATION</b></font>"

ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>Member: </b></font><font color="&TextColor2&" face="&font1&" size=2>"&sFullName&"</font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>MemberID: </b></font><font color="&TextColor2&" face="&font1&" size=2>"&sMemberID&"</font>"

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
ebody = ebody & "<font face="&font1&" size=2><b>Date = </b></font><font color="&TextColor2&" face="&font1&" size=2>"&DATE&"</font></b><br>"

ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2>All transaction data is stored in tables on HQ rankings site.</font></b><br>"



ebody = ebody & "</center>"
ebody = ebody & "<br>"
ebody = ebody & "</td></tr>"

ebody = ebody & "</table>"
' ebody = ebody & "</TD></TR>"

' ebody = ebody & "<TR><TD>"
ebody = ebody + "<div style='width:100%; text-align:center; font-size:8pt; margin:0px 0px 0px 0px;'>A Service of</div>" 
ebody = ebody + "<div style='width:100%; text-align:center; margin:10px 0px 0px 0px;'><img src='"&USAWS_Logo&"' style='width:100px;'></div>" 
ebody = ebody + "<div style='width:100%; text-align:center; font-size:8pt; font-style:bold; margin:10px 0px 0px 0px;'>180 Holy Cow Rd<br>Polk City, FL 33883</div>" 

ebody = ebody & "</TD></TR>"
ebody = ebody & "</TABLE>"
ebody = ebody & "</div>"
ebody = ebody & "</body>"
ebody = ebody & "</html>"


' ebody = ebody + "<div style='width:100%; text-align:center; font-size:8pt; margin:0px 0px 0px 0px;'>A Service of</div>" 
' ebody = ebody + "<div style='width:100%; text-align:center; margin:10px 0px 0px 0px;'><img src='"&USAWS_Logo&"' style='width:100px;'></div>" 
' ebody = ebody + "<div style='width:100%; text-align:center; font-size:8pt; font-style:bold; margin:10px 0px 0px 0px;'>180 Holy Cow Rd<br>Polk City, FL 33883</div>" 



' ------------------------------------------------------------
' --- Dimension and define all the email related variables ---
' ------------------------------------------------------------

Dim eMailSubj, eMailFrom, eMailTo, eMailCC, eMailBCC, eMailBody


eMailTo=USAWaterski_AccountingEmail
'response.write("<br>eMailTo = "&eMailTo)
'response.end

eMailFrom = "competition@usawaterski.org"
eMailBCC = marksemailaddress
eMailSubj = "USA Water Ski - On Line Registration Payment Notification - "&CCFirstName&" "&CCLastName&" - TourID: "&Session("sTourID")
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
objMessage.Send 
set objMessage = Nothing


MarkDebug("USAWS SkipJack Payment = "&CCAmount&" - sTourID = "&sTourID&" - Member = "&sMemberID&" ")


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

ebody = ebody & "<font face="&font1&" size=2><b>MemberID = </b></font><font color="&TextColor2&" face="&font1&" size=2>"&sMemberID&"</font>"
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


' ------------------------------------------------------------
' --- Dimension and define all the email related variables ---
' ------------------------------------------------------------

Dim eMailSubj, eMailFrom, eMailTo, eMailCC, eMailBCC, eMailBody


eMailTo=USAWaterski_AccountingEmail
eMailFrom = "competition@usawaterski.org"
'eMailBCC = marksemail
eMailSubj = "USA Water Ski - On Line Registration REFUND Notification - "&CCFirstName&" "&CCLastName&" - TourID: "&Session("sTourID")
eMailBody = ebody	


' ---------------------------------------------------------------
' --- Now assign the components to the standard email objects ---
' ---------------------------------------------------------------

objMessage.Subject = eMailSubj
objMessage.From = eMailFrom
objMessage.To = eMailTo
objMessage.cc = eMailCC
objMessage.bcc = eMailBCC
objMessage.HTMLBody = eMailBody
objMessage.Send 
set objMessage = Nothing


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

  
 
