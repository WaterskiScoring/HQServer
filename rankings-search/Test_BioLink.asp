<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include16.asp"-->
<!--#include virtual="/rankings/tools_registration16.asp"-->

<%

Dim sMemberID

WriteIndexPageHeader


' response.end
'adminmenulevel = Session("adminmenulevel")

'response.write("<br>adminmenulevel = "&adminmenulevel)
'response.end

'IF TRIM(Session("sMemberID")) = "" THEN adminmenulevel = "1"
'Session("sMemberID")=""

' response.write("<br>TRIM(Session(sMemberID))  = "&TRIM(Session("sMemberID")) )


IF TRIM(Session("sMemberID")) = "" THEN
		Session("sSendingPage")="/rankings/Test_BioLink.asp"
		Response.Redirect("/rankings/search-memberHQ.asp?rid="&rid&"&formstatus=search")
ELSE
		Session("sTourID")="999999"
		DisplayContinueWindow
END IF

WriteIndexPageFooter



IF 2=1 THEN
%>

<br><br><br><br>
<a href="http://www.usawaterski.org/rankings/bio-form.asp?FormStatus=new target="_blank">OLR Bio Update</a>
<%
END IF




' -------------------------
   SUB DisplayContinueWindow
' -------------------------

sMemberID = Session("sMemberID")

'response.write("<br>sMemberID = "&sMemberID)
'response.end
Dim ActionLink
ActionLink = "http://www.usawaterski.org/rankings/bio-form.asp?FormStatus=new&sMemberID="&sMemberID&"&sTourID=999999"

%>
<br><br><br><br><br><br>
<form action="<%=ActionLink%>" method="post" target="_blank">
<TABLE class="innertable" border="4" align="center" width="60%">
<TR>
  <TH align="center">
	<FONT size="3" COlOR="#FFFFFF">Press Continue to Update Personal Bio</FONT>
  </TH>
</TR>  
 
<TR>
  <TD align=center>
	<br><br>
	<input type="submit" name="Continue" value="Continue">
	<br><br><br>
  </TD>
</TR>  
</TABLE>
</form><%


END SUB
%>


