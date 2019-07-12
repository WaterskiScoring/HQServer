<%

' --- Test form for PayPal ---


runningtotal=1.00
sTourID="07W999A"
OrderID="MarkTest1"
simage_url="http://www.usawaterski.org/rankings/images/logos/usawslogo_no_sub.jpg"
sTourName="2007 Goode Water Ski National Championships"


%>        

<form action="https://www.paypal.com/cgi-bin/webscr" method="post" name="PPForm">
	<INPUT type=hidden value=_xclick name=cmd>
	<INPUT type=hidden value=letsgoski@embarqmail.com name=business>
	<INPUT type=hidden value="<%=sTourName%> Registration" name=item_name> 	
	<INPUT type=hidden value="<%=OrderID%>" name="invoice">
	<INPUT type=hidden value="http://www.usawaterski.org/rankings/registration_bywizard.asp?nav=6" name=return>
	<INPUT type=hidden value="http://www.usawaterski.org/rankings/registration_bywizard.asp?nav=6" name=cancel_return>
	<INPUT type=hidden value="<%=simage_url%>" name="cpp_header_image">
	<INPUT type=hidden value="1.00" name=amount>
	<INPUT type=submit title="Make payments with PayPal - it's fast and secure!" value="Continue Payment" border=0 name=dosubmit>
</form>

<%


SUB TestHold  

%>



<form action="https://www.sandbox.paypal.com/us/cgi-bin/webscr" method="post" name="PPForm">
	<INPUT type=hidden value=kingsb_1201116227_biz@embarqmail.com name=business>


<%

END SUB

%>









