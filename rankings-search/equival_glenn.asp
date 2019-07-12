<% Response.Buffer = True %>

<!--#include virtual="/rankings/settingsHQ.asp"-->


<%
Response.Write("Starting! (v 1.2) " & Now() & "<br>")

Server.ScriptTimeout = 2400 

sSql = "Select RT.MemberID, MT.FirstName, MT.LastName, MT.Email, "
sSQL = sSQL & "CASE when RT.Div = 'EM' then 'Open Men' "
sSQL = sSQL & "when RT.Div = 'EW' then 'Open Women' "
sSQL = sSQL & "when RT.Div = 'SM' then 'Masters Men' end as EliteDiv, "
sSQL = sSQL & "CASE when RT.Event = 'S' then 'Slalom' "
sSQL = sSQL & "when RT.Event = 'T' then 'Tricks' "
sSQL = sSQL & "when RT.Event = 'J' then 'Jumping' "
sSQL = sSQL & "else 'Overall' end as EliteEvent, "
sSQL = sSQL & "CASE when patindex('%) as %', RnkScoBkup) between 3 and 15 "
sSQL = sSQL & "then substring(RnkScoBkup,patindex('%) as %',RnkScoBkup)+5,2) "
sSQL = sSQL & "else RT.Div end as OrigDiv "
sSQL = sSQL & "FROM usawsrank.Rankings as RT "
sSQL = sSQL & "JOIN USAWaterski.dbo.members as MT "
sSQL = sSQL & "on MT.PersonIDWithCheckDigit = RT.MemberID "
sSQL = sSQL & "LEFT JOIN usawsrank.EliteDates as EQD "
sSQL = sSQL & "on EQD.MemberID = RT.MemberID "
sSQL = sSQL & "and EQD.DivElite = CASE "
sSQL = sSQL & "when RT.Div = 'EM' then 'OM' "
sSQL = sSQL & "when RT.Div = 'EW' then 'OW' "
sSQL = sSQL & "when RT.Div = 'SM' then 'MM' end "
sSQL = sSQL & "and EQD.Event = left(RT.Event,1) "
sSQL = sSQL & "and EQD.SkiYearID = RT.SkiYearID "
sSQL = sSQL & "WHERE RT.SkiYearID = 1 "
sSQL = sSQL & "and RT.Div in ('EM','EW','SM') "
sSQL = sSQL & "and RT.AWSA_Rat = left(RT.Event,1) + '9' "
sSQL = sSQL & "and patindex('%@%',MT.Email) > 0 "
sSQL = sSQL & "and MT.FederationCode = 'USA' "
sSQL = sSQL & "and EQD.MemberID is Null"
sSQL = sSQL & ";"

	Set rs = Server.CreateObject("ADODB.recordset")

	rs.open sSQL, sConnectionToTRATable, 3, 3
	
	'set CursorType to adOpenForwardOnly and LockType to adLockReadOnly
	'rs.open sSQL, sConnectionToTRATable, 0, 3

	Response.Write("RecordSet opened! " & Now() & "<br>")

	Response.Flush()
	
	If Not rs.eof Then
		
		Response.Write("EOF checked and is False! " & Now() & "<br>")
		
		'rs.MoveFirst 

		Response.Write("Starting to open mail object! " & Now() & "<br>")
			
		Set myMail=CreateObject("CDO.Message")
		myMail.Subject="Welcome to AWSA Elite Skier Status"
		myMail.From = """AWSA President"" <skijump@att.net>"
		myMail.BCC = """Dave Clark"" <awsatechdude@comcast.net>; ""Gene Davis"" <skijump@att.net>"
		
		Response.Write("Finished opening mail object! " & Now() & "<br>")
		Response.Write("Starting Loop! " & Now() & "<br>")		
		
		Response.Write("<pre>")
		While Not rs.eof
			
			Response.Write("MemberID: " & rs("MemberID") & "     " & Now() & "" & vbnewline)
			Response.Flush()

			sSQL = "<html><head><title>Welcome to AWSA Elite Skier Status</title></head>"
			sSQL = sSQL & "<body><basefont face=""arial,sans-serif,helvetica,verdana,tahoma"" color=""#000000"" size=""2"">"

			sSQL = sSQL & "<div style=""border: double 20px #ff0505;"
			sSQL = sSQL & " padding: 25px;"
			sSQL = sSQL & " margin: 10;"
'			sSQL = sSQL & " text-align: justify;"
			sSQL = sSQL & " line-height: 23px;"
			sSQL = sSQL & " color: #070707;"
			sSQL = sSQL & " font-size: 18px"">"
			
			sSQL = sSQL & "<p>To:&nbsp;&nbsp;&nbsp;&nbsp; " & rs("FirstName") & " " & rs("LastName")
			sSQL = sSQL & "<br>Re:&nbsp;&nbsp;&nbsp;&nbsp; AWSA Elite Status in " & rs("EliteDiv") & " " & rs("EliteEvent")
			sSQL = sSQL & "<br>Date:&nbsp; " & FormatDateTime(date(),1) & "</p>"
			
			sSQL = sSQL & "<p>Dear " & rs("FirstName") & ",</p>"

			sSQL = sSQL & "<p>Your prowess as a competitive water skier has elevated you to a new"
			sSQL = sSQL & " pinnacle -- status as an AWSA Elite Skier, in the " & rs("EliteDiv")
			sSQL = sSQL & " Division in " & rs("EliteEvent") & ", based on your Ranking today in the " & rs("OrigDiv")
			sSQL = sSQL & " Division.&nbsp; This is a significant achievement for which you"
			sSQL = sSQL & " should feel quite proud.&nbsp; Feel free to pat yourself on the back.</p>"
		 
			'sSQL = sSQL & "<p>This reflects not only your athleticism, but your dedication to"
			sSQL = sSQL & " disciplined practice sessions and your participation at numerous"
			sSQL = sSQL & " tournaments.&nbsp; Please note that this credential is limited to the"
			sSQL = sSQL & " top 7% of the ranked " & rs("EliteEvent") & " skiers eligible for consideration"
			sSQL = sSQL & " as " & rs("EliteDiv") & ".&nbsp; Simply stated, this puts you in the company of"
			sSQL = sSQL & " AWSA's most outstanding water ski athletes.</p>"

			sSQL = sSQL & "<p>Your next step:&nbsp; a number one ranking among these elite skiers.&nbsp; "
			sSQL = sSQL & "Again, my heartiest congratulations !</p>"

			sSQL = sSQL & "<p>Gene Davis<br>President,<br>American Water Ski Association</p>"
			sSQL = sSQL & "</div></body></html>"

			myMail.To = """" & rs("FirstName") & " " & rs("LastName") & """ <" & rs("Email") & ">"
			myMail.HTMLBody = sSQL

			'myMail.Send
			
			rs.MoveNext

		Wend
		Response.Write("</pre>")
		
		rs.close
		Set myMail = Nothing
		
	Else
		Response.Write("EOF checked and is True! No Records! " & Now() & "<br>")
	End If

	Set rs = Nothing

	Response.Write("DONE! " & Now() & "<br>")
%>

              




 