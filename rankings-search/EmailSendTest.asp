<!--#include file="settingsHQ.asp"-->


<html><head><title>Email Send Testing</title></head><body>

<%

ThisModule = "/rankings/EmailSendTest.asp"

WriteIndexPageHeader

eMailToName = TRIM(Request("eMailToName"))
eMailToAdrs = TRIM(Request("eMailToAdrs"))
eMailFmName = TRIM(Request("eMailFmName"))
eMailFmAdrs = TRIM(Request("eMailFmAdrs"))
eMailRpName = TRIM(Request("eMailRpName"))
eMailRpAdrs = TRIM(Request("eMailRpAdrs"))

IF len(eMailFmAdrs) > 0 and len(eMailFmName) > 0 THEN
	
		SetupEmailService
		objMessage.Subject = "Metisentry Email Send Test"
		objMessage.To = """" & eMailToName & """ <" & eMailToAdrs & ">"
		objMessage.From = """" & eMailFmName & """ <" & eMailFmAdrs & ">"
		objMessage.ReplyTo = """" & eMailRpName & """ <" & eMailRpAdrs & ">"
		objMessage.TextBody = "Metisentry eMail Send Test"& Chr(13) & Chr(10) & Chr(13) & Chr(10) & "To Name / Adrs: " & objMessage.To & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "From Name/Adrs: " & objMessage.From & Chr(13) & Chr(10) & Chr(13) & Chr(10) & "ReplyTo Name/Adrs: " & objMessage.ReplyTo 
		objMessage.Send
		%> 	<table border="0" cellspacing="1" cellpadding="1">
				<tr><td colspan=2><font size="2">eMail appears to have been successfully sent ...<br><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
							To: "<%=eMailToName%>"&nbsp;&lt;<%=eMailToAdrs%>&gt;<br>&nbsp;&nbsp;
							From: "<%=eMailFmName%>"&nbsp;&lt;<%=eMailFmAdrs%>&gt;<br>&nbsp;
							Reply: "<%=eMailRpName%>"&nbsp;&lt;<%=eMailRpAdrs%>&gt;</td></tr>
				<TR><TD>&nbsp;</TD></TR></table>	<%
		Set objMessage=nothing

END IF

%>

	<table border="0" cellspacing="1" cellpadding="1">

		<TR><TD colspan=2> &nbsp; </TD></TR>
		<TR><TD colspan=2><font size=2 color="red"><b>Supply To and From Names and eMail Addresses below ...</b></font></TD></TR>
		<TR><TD colspan=2> &nbsp; </TD></TR>
		
	<FORM method="post" action="<%=ThisModule%>">

		<TR><TD><center><font size="2">To Name: </font></center></TD>
		<TD><center><INPUT type="text" name="eMailToName" size="40" maxlength="50" value="<%=eMailToName%>"></center>
		</TD></TR>
		
		<TR><TD><center><font size="2">To Adrs: </font></center></TD>
		<TD><center><INPUT type="text" name="eMailToAdrs" size="40" maxlength="50" value=<%=eMailToAdrs%>></center>
		</TD></TR>	

		<TR><TD>&nbsp;</TD></TR>
		
		<TR><TD><center><font size="2">From Name: </font></center></TD>
		<TD><center><INPUT type="text" name="eMailFmName" size="40" maxlength="50" value="<%=eMailFmName%>"></center>
		</TD></TR>
		
		<TR><TD><center><font size="2">From Adrs: </font></center></TD>
		<TD><center><INPUT type="text" name="eMailFmAdrs" size="40" maxlength="50" value=<%=eMailFmAdrs%>></center>
		</TD></TR>	

		<TR><TD>&nbsp;</TD></TR>
		
		<TR><TD><center><font size="2">Reply Name: </font></center></TD>
		<TD><center><INPUT type="text" name="eMailRpName" size="40" maxlength="50" value="<%=eMailRpName%>"></center>
		</TD></TR>
		
		<TR><TD><center><font size="2">Reply Adrs: </font></center></TD>
		<TD><center><INPUT type="text" name="eMailRpAdrs" size="40" maxlength="50" value=<%=eMailRpAdrs%>></center>
		</TD></TR>	

		<TR><TD>&nbsp;</TD></TR>
		
		<TR><TD colspan=2><center><INPUT type="Submit" value="Send Test Email To/From/Reply Above Name & Adrs Details" 
			    title="Test eMail Send from The Above Name and Address"></center></TD></TR>
		</FORM>
		
	</table></body></</html>

                                                                    