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
		Session("sSendingPage")="/rankings/bio-form_launcher.asp"
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
ActionLinkUpdate = "http://www.usawaterski.org/rankings/bio-form.asp?FormStatus=new&sMemberID="&sMemberID&"&sTourID=999999"
ActionLinkDone = "/rankings/defaultHQ.asp"
%>
<br><br><br><br><br><br>

<TABLE class="innertable" border="4" align="center" width="40%">
<TR>
  <TH colspan=2 align="center">
	<FONT size="3" COlOR="#FFFFFF">Update Personal Bio</FONT>
  </TH>
</TR>  
 
<TR>
	  
  <TD colspan=1 align=center style="height:65px; width:50%; border:0px solid;">
		<form method="post" action="<%= ActionLinkUpdate %>" target="_blank">
			<input type="submit" name="Continue" value="Update" style="width:7em;">
		</form>
	</td>
	<td colspan=1 align=center style="width:50%; border:0px solid;">
		<form method="post" action="<%= ActionLinkDone %>">
			<input type="submit" name="Done" value="Done" style="width:7em;" formction="<%= ActionLinkDone %>">
  	</form>
  </td>
</TR>  
</TABLE>
</form><%


END SUB
%>


