<!--#include virtual="/rankings/settingsHQ.asp"-->
<%





	' -- Must dimension the object from the calling program --
	' Dim objMessage
	SetupEmailService_membership


	Recipient = "cronemarka@gmail.com"
	MailBody = "Testing Program Only - This is where the body of the email goes "

				
	objMessage.To = Recipient
	objMessage.Subject = "USA-WSWS Testing Email"
	objMessage.From = "competition@usawaterski.org"
				
	objMessage.HTMLBody = MailBody


	objMessage.Send
	Set objMessage = Nothing


response.write("<br><br>COMPLETE")				




' -----------------------------------
  SUB SetupEmailService_membership
' -----------------------------------

	Dim EmailAccountTableName
	Dim t_senderusername, t_sendpassword, t_smtpserver, t_smtpport, t_enableSsl, t_displayname, t_userdefaultcredentials
	
	
	EmailAccountTableName = "usawaterski.dbo.EmailAccount"

	
	SET rs=Server.CreateObject("ADODB.recordset")
	sSQL = "SELECT * FROM "&EmailAccountTableName
	sSQL = sSQL + " WHERE ID = 1"
	rs.open sSQL, sConnectionToTRATable, 3, 1

	IF NOT rs.eof THEN
				t_emailFrom = rs("Email")
				t_senderusername = rs("Username")
				t_sendpassword = rs("Password")
				t_smtpserver = rs("Host")
				t_smtpserverport = rs("Port")
				t_enableSsl = rs("enablessl")
				t_displayname = rs("DisplayName")
				t_userdefaultcredentials = rs("UseDefaultCredentials")				 
 
   			t_sendusing = 2				 
 				t_smtpauthenticate = 1
    		t_smtpconnectiontimeout = 60
    		
 				' t_sendusing = rs("sendusing")				 
 				' t_smtpauthenticate = rs("smtpauthenticate")
    		' t_smtpconnectiontimeout = rs("smtpconnectiontimeout")
	ELSE 
				t_emailFrom = Application("emailUNranking")
				t_senderusername = Application("emailUNranking")
				t_sendpassword = Application("emailPWranking")
				t_smtpserver = "smtp.office365.com"
				t_smtpserverport = 25
				t_enableSsl = "true"
				t_displayname = "Competition Services"
				t_userdefaultcredentials = "unknown what this means and is for"			
 
  			t_sendusing = 2				 
 				t_smtpauthenticate = 1
    		t_smtpconnectiontimeout = 60
    		   		
	END IF

	Dim emailFrom, FriendlyName
	emailFrom = t_emailFrom
	FriendlyName = t_displayname


	Set objMessage = CreateObject("CDO.Message")
	'=====  This section provides the configuration information for the remote SMTP server.
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendusing") = t_sendusing
	'Name or IP of Remote SMTP Server
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserver") = t_smtpserver

	'Type of authentication, NONE, Basic (Base64 encoded), NTLM
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = t_smtpauthenticate

	'Your UserID on the SMTP server
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendusername") = t_senderusername

	'Your password on the SMTP server
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendpassword") = t_sendpassword

	'Server port (typically 25)
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = t_smtpserverport
	'Use SSL for the connection (False or True)
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = t_enableSsl
	'Connection Timeout in seconds (the maximum time CDO will try to establish a connection to the SMTP server)
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = t_smtpconnectiontimeout
	objMessage.Configuration.Fields.Update
	'=====  End remote SMTP server configuration section


END SUB



%>				