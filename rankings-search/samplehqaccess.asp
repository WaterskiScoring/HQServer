
<%

Set SQLConnect = CreateObject("ADODB.Connection")
SQLConnect.Open Application("HQSQLConn")
Set RS = SQLConnect.Execute("SELECT * FROM [Membership History]")
rowcounter = 1
While (Not RS.eof and rowcounter < 10)
	response.write  RS("Person ID") & " | "
	response.write  RS("Membership Type Code") & " | "
	response.write  RS("EffectiveFrom") & " | "
	response.write  RS("EffectiveTo") & "</br>"
	
	RS.movenext
	rowcounter = rowcounter + 1
wend

rs.close

%>





