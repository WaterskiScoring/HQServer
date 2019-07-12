<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->

<%


Dim ThisFileName, ThisStatus, ProcessorFileName

Dim sMemberID, sTournAppID, sOrderNo, CCAmount

ThisFileName="PaymentLaunch.asp"
ProcessorFileName="CCSJ_Temp.asp"


' ------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------
' --- NOTES





' ------------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------------




Action=TRIM(Request("Action"))

sMemberID=TRIM(Request("sMemberID"))
IF sMemberID<>"" THEN Session("sMemberID")=sMemberID

sTournAppID=TRIM(Request("sTournAppID"))
'response.write("<br>HERE - sTournAppID="&sTournAppID)

IF sTournAppID<>"" THEN Session("sTournAppID")=sTournAppID
'response.write("<br>Session(sTournAppID)="&Session("sTournAppID"))



sOrderNo=TRIM(Request("sOrderNo"))

CCAmount=TRIM(Request("CCAmount"))

'response.write("<br>Action="&Action)








' --- Defines CSS table and font styles ---
DefineTRAStyles 


SELECT CASE Action

   CASE ""
	BuildForm

   CASE "Confirm Values"
	ValidateValues	

   CASE "Check Out"
	IF sPaymentMethod="Check By Mail" THEN
		DoneWithAll
	ELSE
		Session("sMemberID")=sMemberID
		Session("sTournAppID")=sTournAppID
		Response.redirect(ProcessorFileName&"?action=new&sOrderNo="&sOrderNo&"&CCAmount="&CCAmount)	
	END IF

   CASE "Return To Display Receipt"
	ThisStatus="NOW DISPLAY RECEIPT"	
	ValidateValues

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
	    <td Colspan=2 ALIGN="center"><%
		IF ThisStatus="NOW DISPLAY RECEIPT" THEN %>
		   <input type="submit" VALUE="Done All"><%
		ELSE %>
		   <input type="submit" NAME="Action" VALUE="Check Out"><%
		END IF %>
	    </td>
	  <tr>

	</table>

	</form><%



END SUB




  SUB DoneWithAll

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

%>