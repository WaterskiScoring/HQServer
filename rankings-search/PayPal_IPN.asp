<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->
<!--#include virtual="/rankings/tools_registration.asp"-->
<%

' =========================================================================================
' =========================================================================================
' --- PayPal *** INSTANT PAYMENT NOTIFICATION - aka IPN ***

' --- Reads post from PayPal upon completion of transaction and updates Reg_PaymentLog

' =========================================================================================
' =========================================================================================


Dim PayPalMessage, PayPalResponse, PayPalStr, urlstr, ErrorNo, xmlDoc 
Dim txn_id, sTourID, sMemberID
Dim ResponseValid


' --- Variables returned from PayPal ---
Dim invoice
Dim payment_status, mc_gross, payer_email, first_name, last_name, address_city, address_state, address_street, address_zip, ipn_track_id
Dim sPaymentResult, DateNow



' --- Defines all the styles used by the site ---
DefineTRAStyles



' -----------------------------------------------------------------------------------------------------
' --- This HTML only displays if the page is explicitly placed in URL.  If post then nothing happens 
' -----------------------------------------------------------------------------------------------------
%>
<br><br><br><br><br>
<table class="innertable" align="center" width=50%>
	<tr>
		<th align="center" height="30px">
			<font size="4" color="white"><b>Pay Pal IPN RETURN</b></font>
		</th>	
	</tr>
	<tr>
		<td align="center" height="100px">
			<font size="4" color="black">This pages demonstrates return for IPN process</font>
		</td>
	</tr>
</table>
<%



' --- IPN Message from PayPal
' --- mc_gross=2.00&invoice=105&protection_eligibility=Eligible&address_status=confirmed&item_number1=&tax=0.00&item_number2=&ipn_track_id=GHX9VQGLRFC5W&address_street=1+Main+St&payment_date=12%3A35%3A11+Jan+04%2C+2014+PST&payment_status=Completed&charset=windows-1252&address_zip=95131&mc_shipping=0.00&mc_handling=0.00&first_name=Mark&mc_fee=0.36&address_country_code=US&address_name=Mark+Crone¬ify_version=3.7&custom=&payer_status=verified&business=mark%40kingsbridgehomes.com&address_country=United+States&num_cart_items=2&mc_handling1=0.00&mc_handling2=0.00&address_city=San+Jose&verify_sign=AFcWxV21C7fd0v3bYYYRCpSSRl31ADjFfGdm0VSCxQtqrOUgBpXL3zbu&payer_email=mark%40paypal-developer-test.com&mc_shipping1=0.00&mc_shipping2=0.00&tax1=0.00&tax2=0.00&txn_id=2MS598739L339801R&payment_type=instant&last_name=Crone&address_state=CA&item_name1=My+Test+Item+1&receiver_email=mark%40kingsbridgehomes.com&item_name2=My+Test+Item+2&payment_fee=0.36&quantity1=1&quantity2=2&receiver_id=75N5U43C592VC&txn_type=cart&mc_gross_1=1.00&mc_currency=USD&mc_gross_2=1.00&residence_country=US&test_ipn=1&transaction_subject=&payment_gross=2.00&ipn_track_id=a48f5ecfe16cb


' --- Before you can trust the contents of the message, you must first verify that the message came from PayPal. To verify the message, you must send back the contents in the exact order they were received and precede it with the command _notify-validate, as follows:
' https://www.sandbox.paypal.com/cgi-bin/webscr?cmd=_notify-validate&mc_gross=19.95&protection_eligibility=Eligible&address_status=confirmed&ipn_track_id=LPLWNMTBWMFAY&tax=0.00&...&payment_gross=19.95&shipping=0.00

' mc_gross=2.00&invoice=107&protection_eligibility=Eligible&address_status=confirmed&item_number1=&tax=0.00&item_number2=&ipn_track_id=GHX9VQGLRFC5W&address_street=1+Main+St&payment_date=13%3A09%3A50+Jan+04%2C+2014+PST&payment_status=Completed&charset=windows-1252&address_zip=95131&mc_shipping=0.00&mc_handling=0.00&first_name=Mark&mc_fee=0.36&address_country_code=US&address_name=Mark+Crone¬ify_version=3.7&custom=&payer_status=verified&business=mark%40kingsbridgehomes.com&address_country=United+States&num_cart_items=2&mc_handling1=0.00&mc_handling2=0.00&address_city=San+Jose&verify_sign=AFcWxV21C7fd0v3bYYYRCpSSRl31A79Z5nzSnpQEAPBj4xrRh97SEYYc&payer_email=mark%40paypal-developer-test.com&mc_shipping1=0.00&mc_shipping2=0.00&tax1=0.00&tax2=0.00&txn_id=6R4456897F673333S&payment_type=instant&last_name=Crone&address_state=CA&item_name1=My+Test+Item+1&receiver_email=mark%40kingsbridgehomes.com&item_name2=My+Test+Item+2&payment_fee=0.36&quantity1=1&quantity2=2&receiver_id=75N5U43C592VC&txn_type=cart&mc_gross_1=1.00&mc_currency=USD&mc_gross_2=1.00&residence_country=US&test_ipn=1&transaction_subject=&payment_gross=2.00&ipn_track_id=bd96dab536f7e
 

 


'cmd=_notify-validate&mc_gross=19.95&protection_eligibility=Eligible&address_status=confirmed&ipn_track_id=LPLWNMTBWMFAY&tax=0.00&...&payment_gross=19.95&shipping=0.00



' --- Listens and reads the variables posted by PayPal when payment is made as a result of inclusion of the "notify_URL" parameter 
' ---   in original data from OLR payment post 
ReadPostVariables

' --- Validates the response by reading the full string and reposting to PayPal for proper response ---
'PayPal_Validate_New


PayPal_Validate_PaymentLog



' =============================================================================================================================================
' ---                                                         ***** END OF PROGRAM *****
' =============================================================================================================================================




' ----------------------
  SUB ReadPostVariables
' ----------------------  

txn_id=Request("txn_id")
mc_gross=Request("mc_gross")
sTourID=Request("sTourID")
sMemberID=Request("sMemberID")
payment_status=Request("payment_status")
mc_gross=Request("mc_gross")

first_name=LEFT(TRIM(Request("first_name")),20)
' --- Changed 5-13-2014 ---
last_name=SQLClean(LEFT(TRIM(Request("last_name")),25))
address_street=LEFT(TRIM(REPLACE(Request("address_street"),vbCrLf," ")),30)
address_city=LEFT(TRIM(Request("address_city")),25)
address_state=LEFT(TRIM(Request("address_state")),25)
address_zip=LEFT(TRIM(Request("address_zip")),10)
payer_email=LEFT(TRIM(Request("payer_email")),50)

invoice=Request("invoice")
ipn_track_id=Request("ipn_track_id")
PayPalMessage = request.form


END SUB




' ----------------------------------
  SUB PayPal_Validate_PaymentLog
' ----------------------------------

SET rsPayLog=Server.CreateObject("ADODB.recordset")

sSQL = "SELECT TOP 1 PayType AS LogPayType, Amount AS LogAmount, OrderNo AS LogOrderNo FROM "&RegPaymentTableName
sSQL = sSQL + " WHERE MemberID='"&sMemberID&"' AND TourID='"&sTourID&"'"
sSQL = sSQL + " ORDER BY OrderNo DESC"
rsPayLog.open sSQL, sConnectionToTRATable, 3, 1

sPaymentResult=""
LogOrder=999999999
LogAmount=-99999
IF NOT rsPayLog.eof THEN
		IF TRIM(rsPayLog("LogPayType"))<>"" THEN
				LogPayType=rsPayLog("LogPayType")
				LogAmount=rsPayLog("LogAmount")
				LogOrderNo=rsPayLog("LogOrderNo")
		END IF
END IF

' --- Validated log amount against IPN and status is completed and log last transaction is PayPal for this tournament ---
'IF TRIM(LogPayType)="PayPal" AND LCASE(payment_status)="completed" AND cdbl(mc_gross)=cdbl(LogAmount) AND cdbl(LogOrderNo)=cdbl(invoice) THEN 
IF TRIM(LogPayType)="PayPal" AND LCASE(payment_status)="completed" AND cdbl(mc_gross)=cdbl(LogAmount) THEN 

			' --- Filler values ---
			Dim sLast4Card, sExpMonth, sExpYear, sApvl_Code, sCvv2_Resp, sAVS_Resp, sCheckNo, sPayStatus

			' --- Primary filtering criteria for confirming payment - sPaymentResult="0"
			sPaymentResult="0"

			' --- Updates the date originally inserted when record was created and prepopulated before PayPal button was pressed ---
			DateNow = Now

			sCheckNo=""
			sLast4Card=""
			sExpMonth=""
			sExpYear=""
			sApvl_Code=""
			sCvv2_Resp=""
			sAVS_Resp=""						


			sPayStatus="N"
			
			' --- Versioning ?? ---
			OLR_PayPal_IPN_Message="PayPal_IPN.asp - Effective 01-27-2014"

			' ------------------------		
			' --- PayStatus Decode ---
			' ------------------------
			' --- I = Initialized in OLR
			' --- N = PayPal IPN fired
			' --- C = Made it to Receipt Page of OLR



						
			' --- Update the RegPaymentLog table ---
			OpenCon
			sSQL = "UPDATE "&RegPaymentTableName
			sSQL = sSQL + " SET"
			sSQL = sSQL + " Result='"&sPaymentResult&"'"
			sSQL = sSQL + ", FirstName='"&SQLClean(first_name)&"', LastName='"&SQLClean(last_name)&"'"
			sSQL = sSQL + ", Address1='"&SQLClean(address_street)&"', City='"&SQLClean(address_city)&"', State='"&SQLClean(address_state)&"', ZipCode='"&address_zip&"'"
			sSQL = sSQL + ", Email='"&payer_email&"'"
			
			sSQL = sSQL + ", Txn_ID='"&ipn_track_id&"'"
			sSQL = sSQL + ", Message='"&OLR_PayPal_IPN_Message&"'"

			sSQL = sSQL + ", CheckNo='"&sCheckNo&"'"		
			sSQL = sSQL + ", Last4Card='"&sLast4Card&"', Apvl_Code='"&sApvl_Code&"', Cvv2_Resp='"&sCvv2_Resp&"', AVS_Resp='"&sAVS_Resp&"'"
			sSQL = sSQL + ", ExpYear='"&sLast4Card&"', ExpMonth='"&sExpMonth&"'"
			sSQL = sSQL + ", PayStatus='"&sPayStatus&"'"
			sSQL = sSQL + ", TransDate='"&DateNow&"'"

			sSQL = sSQL + " WHERE MemberID='"&sMemberID&"' AND TourID='"&sTourID&"' AND OrderNo='"&LogOrderNo&"'"
			con.execute(sSQL)
	END IF


	


END SUB

' ------------------------
  SUB PayPal_Validate_New 
' ------------------------

  response.write("<br>PayPalMessage: "&PayPalMessage&"<br><br><br>")
	'response.end

	PayPalStr = "cmd=_notify-validate&"&PayPalMessage
	
	dim xmlhttp 
	set xmlhttp = server.Createobject("MSXML2.ServerXMLHTTP")
	xmlhttp.Open "POST","http://localhost/Receiver.asp",false
	xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	xmlhttp.send PayPalStr
	Response.ContentType = "text/html"
	Response.Write xmlhttp.responsexml.xml
	
	ResponseValid=xmlhttp.responsexml.xml
  
  Set xmlhttp = nothing

	SendMyEmail

END SUB


' ---------------------
  SUB PayPal_Validate
' ---------------------

PayPalStr = "cmd=_notify-validate&"&PayPalMessage

urlstr="https://www.sandbox.paypal.com/cgi-bin/webscr"


'set xmlhttp = CreateObject("Microsoft.XMLHTTP")
'xmlhttp.Open "POST", urlstr&"?"&PayPalStr, False
'xmlhttp.Send

'Set obj_post=Server.CreateObject("Msxml2.SERVERXMLHTTP ")

Dim obj_post
Set obj_post=CreateObject("Microsoft.XMLHTTP")
obj_post.Open "POST", urlstr, False
obj_post.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
obj_post.Send PayPalStr


ErrorNo=Err.Number
'IF Err.Number <> 0 Then
'    Response.Write("Your web site does not appear to be up right now. Please try again later.")
'    Err.Clear 
'ELSE
'    Response.write("data successfully posted<br />")
'END IF


'Response.ContentType = "application/xml"
'response.write("<br><HTML>"& PayPalStr &"</HTML><br>")
'response.end

' --- Read response
Set xmlDoc=CreateObject("Microsoft.XMLDOM")
xmlDoc.async="false"
xmlDoc.loadXML(obj_post.responseText)


' --- Determines if a valid post was made - If so, then parse the XML the data ---
NumNodes=xmlDoc.getElementsByTagName("RESULT").length
Set xmlList = xmlDoc.getElementsByTagName("RESULT")

FOR Each xmlItem In xmlList
		IF LCASE(xmlItem.childNodes(0).text)="true" THEN

				'GetValuesForCustomer
				' --- Displays the form ---
				'ShowLandingPageForm
		ELSE
				response.write("<br><br>False")
				'DisplayErrorMessage
				'ShowLandingPageForm
		END IF
NEXT


Set xmlDoc = Nothing

' --- During testing sends an email with the data that was recevied ---
SendMyEmail

END SUB




' ----------------------
   SUB SendMyEmail
' ----------------------

' http://usawaterski.org/rankings/Test_PayPal_IPN_Notify.asp

Dim eMailSubj, eBody, eMailFrom, eMailBCC, sTest
Dim sTsEmail




ebody = "<html>"
ebody = ebody & "<head>"
ebody = ebody & "<title>Notice to Mark Crone</title>"
ebody = ebody & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
ebody = ebody & "</head>"
ebody = ebody & "<body bgcolor=""#FFFFFF"" text=""#000000"">"
ebody = ebody & "<div align=""center"">"


ebody = ebody & "<TABLE BORDER=4 ALIGN=CENTER CELLPADDING=3 CELLSPACING=0 BGCOLOR="&TableColor1&" width=60% >"
ebody = ebody & "<TR>"
ebody = ebody & "<TD BGCOLOR=red><center><font face="&font1&" color=#FFFFFF size=4><b>Notice to Mark Crone</b></font></TD>"
ebody = ebody & "</TR>"
 
ebody = ebody & "<TR>"
ebody = ebody & "<td Align=center>"	
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>This message was generated because the IPN Posted.</b></font>"
ebody = ebody & "</td>"
ebody = ebody & "</TR>"

ebody = ebody & "<TR>"
ebody = ebody & "<td Align=center>"	
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>TourID: "&sTourID&"</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>MemberID: "&sMemberID&"</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>Name: "&first_name&" "&last_name&"</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>Trans ID: "&txn_id&"</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>Amount: "&mc_gross&"</b></font>"
ebody = ebody & "</td>"
ebody = ebody & "</TR>"

ebody = ebody & "<TR>"
ebody = ebody & "<td Align=center>"	
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>PayPalStr: "&PayPalStr&"</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "</td>"
ebody = ebody & "</TR>"

ebody = ebody & "<TR>"
ebody = ebody & "<td Align=center>"	
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>urlstr: "&urlstr&"</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "</td>"
ebody = ebody & "</TR>"

ebody = ebody & "<TR>"
ebody = ebody & "<td Align=center>"	
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>ErrorNo: "&ErrorNo&"</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "</td>"
ebody = ebody & "</TR>"

ebody = ebody & "<TR>"
ebody = ebody & "<td Align=center>"	
ebody = ebody & "<br>"
ebody = ebody & "<font face="&font1&" size=2><b>ResponseValid: "&ResponseValid&"</b></font>"
ebody = ebody & "<br>"
ebody = ebody & "</td>"
ebody = ebody & "</TR>"

ebody = ebody & "</TABLE>"




eMailSubj="Test Listener"
eMailTo="mark@productdesign-biz.com"
eMailFrom="mark@productdesign-biz.com"
eMailCC=""
eMailBCC=""


SetupEmailService

objMessage.To = eMailTo
objMessage.From = eMailFrom
IF LCASE(eMailBCC)<>"" THEN objMessage.bcc = eMailBCC
IF LCASE(eMailCC)<>"" THEN objMessage.cc = eMailCC

objMessage.Subject = eMailSubj
objMessage.HTMLBody = ebody
 
 ' --- Finally send the message, and then clear that object
objMessage.Send
set objMessage = Nothing


END SUB






%>
