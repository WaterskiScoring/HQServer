<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->

<%


Dim ThisFileName, ThisStatus, ProcessorFileName
Dim sPaymentMethod, sLP
Dim sMemberID, sTournAppID, sOrderNo, CCAmount

ThisFileName="TEST_PaymentLaunch.asp"
ProcessorFileName="CCSJ_Sanctions.asp"


' ------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------
' --- NOTES





' ------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------




Action=TRIM(Request("Action"))

sMemberID=TRIM(Request("sMemberID"))
sTournAppID=TRIM(Request("sTournAppID"))

sPaymentMethod=TRIM(Request("sPaymentMethod"))
sOrderNo=TRIM(Request("sOrderNo"))
CCAmount=TRIM(Request("CCAmount"))
sLP=TRIM(Request("sLP"))

' --- Change session to current value --- 
IF sMemberID<>"" THEN Session("sMemberID")=sMemberID
IF sTournAppID<>"" THEN Session("sTournAppID")=sTournAppID



response.write("<br>TOP OF FORM - sTournAppID="&sTournAppID)
response.write("<br>Session(sTournAppID)="&Session("sTournAppID"))
response.write("<br>sMemberID="&sMemberID)


response.write("<br>Action="&Action)
response.write("<br>CCAmount="&CCAmount)
response.write("<br>OrderNo="&OrderNo)
response.write("<br>sPaymentMethod="&sPaymentMethod)
response.write("<br>sLP="&sLP)



IF sPaymentMethod="Check By Mail" THEN Action="MailPayment"

'response.end




' --- Defines CSS table and font styles ---
DefineTRAStyles 


SELECT CASE Action

   CASE ""
	BuildForm

   CASE "Confirm Values"
	ValidateValues	

   CASE "MailPayment"
	MailPayment	

   CASE "Check Out"
	IF sPaymentMethod="Check By Mail" THEN
		DoneWithAll
	ELSE
		Session("sMemberID")=sMemberID
		Session("sTournAppID")=sTournAppID
		Response.redirect(ProcessorFileName&"?action=new&sOrderNo="&sOrderNo&"&CCAmount="&CCAmount&"&sLP="&sLP	)	
	END IF

   CASE "Return To Display Receipt"
	ThisStatus="NOW DISPLAY RECEIPT"	
	ValidateValues

   CASE ELSE
       
	Response.write("<br><br><br><br>THE END")
 	
END SELECT







' --------------
  SUB BuildForm
' --------------
	%>
	<form action="<%=ThisFileName%>" method="post">   

	<br><br><br>
	<table ALIGN="center" class="innertable" width=40%>   	
	  <tr>
	    <th colspan="2" align="center"><FONT size=<% =fontsize3 %> COlOR="#FFFFFF"><strong>Cardholder Processor Launch Page</strong></FONT></th>
	  </tr>

	  <tr> 
	    <td Width=200px ALIGN="right">
		<font size=<% =fontsize2 %>  COlOR="<% =textcolor1 %>"><b>TournAppID:</b>&nbsp;&nbsp</FONT>
	    </td>
            <td width="350" align="left">
		<a title="Enter a valid TournAppID for Test">
		<input type="text" name="sTournAppID" value= "<% =sTournAppID %>" MaxLength=7 size="7"></a>
	    </td>
	  </tr>

	  <tr> 
	    <td ALIGN="right">
		<font size=<% =fontsize2 %>  COlOR="<% =textcolor1 %>"><b>ClubID (MemberID):</b>&nbsp;&nbsp</FONT>
	    </td>
            <td align="left">
		<a title="Enter a valid CLUB MemberID for Test">
		<input type="text" name="sMemberID" value= "<% =sMemberID %>" MaxLength=9 size="9"></a>
	    </td>
	  </tr>

	  <tr> 
	    <td ALIGN="right">
		<font size=<% =fontsize2 %>  COlOR="<% =textcolor1 %>"><b>Order No:</b>&nbsp;&nbsp</FONT>
	    </td>
            <td align="left">
		<a title="Enter a valid Order Number (6 digit) for Test">
		<input type="text" name="sOrderNo" value= "<% =sOrderNO %>" MaxLength=7 size="7"></a>
	    </td>
	  </tr>

	  <tr> 
	    <td ALIGN="right">
		<font size=<% =fontsize2 %>  COlOR="<% =textcolor1 %>"><b>Amount:</b>&nbsp;&nbsp</FONT>
	    </td>
            <td align="left">
		<a title="Enter a valid Amount in format 9999.99">
		<input type="text" name="CCAmount" value= "<% =CCAmount %>" MaxLength=7 size="7"></a>
	    </td>
	  </tr>

	  <tr><td colspan=2>&nbsp;</td></tr>

	  <tr> 
	    <td ALIGN="right">
		<font size=<% =fontsize2 %>  COlOR="<% =textcolor1 %>"><b>Payment Method</FONT>
	    </td>
	    <td align="left">
		   <FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR=<% =textcolor1 %> ><b>Check By Mail</b></font>
		   <input type=radio NAME="sPaymentMethod" VALUE="Check By Mail" <% IF sPaymentMethod="Check By Mail" THEN response.write "checked"%> >

		   <FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR=<% =textcolor1 %> ><b>Credit Card</b></font>
		   <input type=radio NAME="sPaymentMethod" VALUE="Credit Card" <% IF sPaymentMethod="Credit Card" THEN response.write "checked"%> >
	    </td>
	  </tr>

	  <tr> 
	    <td ALIGN="right">
		<font size=<% =fontsize2 %>  COlOR="<% =textcolor1 %>"><b>Live or Test Payment</FONT>
	    </td>
	    <td align="left">
		   <FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR=<% =textcolor1 %> ><b>Live</b></font>
		   <input type=radio NAME="sLP" VALUE="Live" <% IF sLP="Live" THEN response.write "checked"%> >

		   <FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR=<% =textcolor1 %> ><b>Test</b></font>
		   <input type=radio NAME="sLP" VALUE="Test" <% IF sLP="Test" THEN response.write "checked"%> >
	    </td>
	  </tr>

	  <tr> 
	    <td Colspan=2 ALIGN="center">
		   <input type="submit" NAME="Action" VALUE="Confirm Values">
	    </td>
	  <tr>

	</table>

	</form><%
 
END SUB





' -------------------
   SUB ValidateValues
' -------------------

Dim Headline
IF ThisStatus="NOW DISPLAY RECEIPT" THEN
	Headline="Returned From Processor to Sending Page"
ELSE
	Headline="Cardholder Processor Launch Page"
END IF

%>
	<form action="<%=ThisFileName%>" method="post">   

	<input type=hidden name="sMemberID" value="<%=sMemberID%>">
	<input type=hidden name="sTournAppID" value="<%=sTournAppID%>">
	<input type=hidden name="sOrderNo" value="<%=sOrderNo%>">
	<input type=hidden name="CCAmount" value="<%=CCAmount%>">
	<input type=hidden name="sLP" value="<%=sLP%>">

	<br><br><br>
	<table ALIGN="center" class="innertable" width=40%>   	
	  <tr>
	    <th colspan="2" align="center"><FONT size=<% =fontsize3 %> COlOR="#FFFFFF"><strong><%=Headline%></strong></FONT></th>
	  </tr>

	  <tr> 
	    <td Width=200px ALIGN="right">
		<font size=<% =fontsize2 %>  COlOR="<% =textcolor1 %>"><b>Member-Club ID:</b>&nbsp;&nbsp</FONT>
	    </td>
	    <td Width=200px ALIGN="left">
		<font size=<% =fontsize2 %>  COlOR="<% =textcolor1 %>"><%=sMemberID%></b>&nbsp;&nbsp</FONT>
	    </td>
	  </tr>
	  <tr> 
	    <td Width=200px ALIGN="right">
		<font size=<% =fontsize2 %>  COlOR="<% =textcolor1 %>"><b>TournAppID:</b>&nbsp;&nbsp</FONT>
	    </td>
	    <td Width=200px ALIGN="left">
		<font size=<% =fontsize2 %>  COlOR="<% =textcolor1 %>"><%=sTournAppID%></b>&nbsp;&nbsp</FONT>
	    </td>
	  </tr>
	  <tr> 
	    <td Width=200px ALIGN="right">
		<font size=<% =fontsize2 %>  COlOR="<% =textcolor1 %>"><b>Order No:</b>&nbsp;&nbsp</FONT>
	    </td>
	    <td Width=200px ALIGN="left">
		<font size=<% =fontsize2 %>  COlOR="<% =textcolor1 %>"><%=sOrderNo%></b>&nbsp;&nbsp</FONT>
	    </td>
	  </tr>
	  <tr> 
	    <td Width=200px ALIGN="right">
		<font size=<% =fontsize2 %>  COlOR="<% =textcolor1 %>"><b>Amount:</b>&nbsp;&nbsp</FONT>
	    </td>
	    <td Width=200px ALIGN="left">
		<font size=<% =fontsize2 %>  COlOR="<% =textcolor1 %>"><%=CCAmount%></b>&nbsp;&nbsp</FONT>
	    </td>
	  </tr>

	  <tr> 
	    <td Width=200px ALIGN="right">
		<font size=<% =fontsize2 %>  COlOR="<% =textcolor1 %>"><b>Transcation (Live or Test):</b>&nbsp;&nbsp</FONT>
	    </td>
	    <td Width=200px ALIGN="left">
		<font size=<% =fontsize2 %>  COlOR="<% =textcolor1 %>"><%=sLP%></b>&nbsp;&nbsp</FONT>
	    </td>
	  </tr>


	  <tr> 
	    <td Colspan=2 ALIGN="center">
		<br><%
		IF ThisStatus="NOW DISPLAY RECEIPT" THEN %>
		   <input type="submit" VALUE="Done All"><%
		ELSE %>
		   <input type="submit" NAME="Action" VALUE="Check Out"><%
		END IF %>
		<br>
	    </td>
	  <tr>

	</table>

	</form><%



END SUB



' ---------------
  SUB DoneWithAll
' ---------------

Headline="Start Over"

%> 

	<form action="<%=ThisFileName%>" method="post">   

	<br><br><br>
	<table ALIGN="center" class="innertable" width=40%>   	
	  <tr>
	    <th colspan="2" align="center"><FONT size=<% =fontsize3 %> COlOR="#FFFFFF"><strong><%=Headline%></strong></FONT></th>
	  </tr>

	  <tr> 
	    <td Colspan=2 ALIGN="center">
		   <input type="submit" NAME="Action" VALUE="Start Over">
	    </td>
	  <tr>

	</table>

	</form><%




END SUB




' ---------------
SUB MailPayment
' ---------------

' --- Just needed for demonstration ---

Headline="END PAGE"

%>
	<form action="<%=ThisFileName%>" method="post">   

	<br><br><br>
	<table ALIGN="center" class="innertable" width=40%>   	
	  <tr>
	    <th colspan="2" align="center"><FONT size=<% =fontsize3 %> COlOR="#FFFFFF"><strong><%=Headline%></strong></FONT></th>
	  </tr>

	  <tr> 
	    <td Colspan=2 ALIGN="center">
		<br><br>
		   <input type="submit" NAME="Action" VALUE="Done Go To Your Confirmation Page">
		<br><br>
	    </td>
	  <tr>

	</table>

	</form><%


END SUB
%>