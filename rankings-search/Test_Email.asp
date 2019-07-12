<%

response.write("TEST EMAILING")


 ' response.write("<br><br>Application = ")
 ' response.write(Application("emailUNranking"))
 ' response.write("<br>Application = ")
 ' response.write(Application("emailPWranking"))
' response.end

' Recipient = "cronemarka@gmail.com"
' Recipient = "mark.crone@bonniercorp.com"
' Recipient = "mawsa@comcast.net"
Recipient = "mawsa#comcast.net"
MailBody = "Testing Program Only"


				'MOK for migration to office 365
				Set objMessage = CreateObject("CDO.Message")
				objMessage.Subject = "USA-WSWS Testing Email"
				' objMessage.From = "noreply@usawaterski.org"
				objMessage.From = "competition@usawaterski.org"
				' objMessage.From = "memberservices@usawaterski.org"
				objMessage.To = Recipient
				' objMessage.bcc = "archive@epolk.com"
				objMessage.HTMLBody = MailBody
				'==This section provides the configuration information for the remote SMTP server.
				objMessage.Configuration.Fields.Item _
				("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
				'Name or IP of Remote SMTP Server
				objMessage.Configuration.Fields.Item _
				("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.office365.com"
				'Type of authentication, NONE, Basic (Base64 encoded), NTLM
				objMessage.Configuration.Fields.Item _
				("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'cdoBasic
				'Your UserID on the SMTP server
				objMessage.Configuration.Fields.Item _
				("http://schemas.microsoft.com/cdo/configuration/sendusername") = Application("emailUNranking")
				' objMessage.Configuration.Fields.Item _
				' ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "competition@usawaterski.org"
				'Your password on the SMTP server
				objMessage.Configuration.Fields.Item _
				("http://schemas.microsoft.com/cdo/configuration/sendpassword") = Application("emailPWranking")

				'Server port (typically 25)
				objMessage.Configuration.Fields.Item _
				("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
				'Use SSL for the connection (False or True)
				objMessage.Configuration.Fields.Item _
				("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
				'Connection Timeout in seconds (the maximum time CDO will try to establish a connection to the SMTP server)
				objMessage.Configuration.Fields.Item _
				("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
				objMessage.Configuration.Fields.Update
				'==End remote SMTP server configuration section==
				''''objMessage.Send
				
				On Error Resume Next
					objMessage.Send
					If Err.Number <> 0 Then
						%><p>Error sending email: Err.Number=<%=Err.Number%> Message=<%=Err.Description%></p><%
						On Error Goto 0 ' But don't let other errors hide!
					ELSE
						%><p>email sent</p><%
						%><p>Error sending email: Err.Number=<%=Err.Number%> Message=<%=Err.Description%></p><%
					End If
				
				Set objMessage = Nothing

response.write("<br><br>COMPLETE")				
%>				