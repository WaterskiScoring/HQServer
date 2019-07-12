
<%

Set SQLConnect = CreateObject("ADODB.Connection")
SQLConnect.Open Application("sConnectionToMemberTable")
Set RS = SQLConnect.Execute("SELECT Top 10 FROM [Members]")
rowcounter = 1
While (Not RS.eof and rowcounter < 10)
	response.write  RS(0) & " | "
	response.write  RS(1) & " | "
	response.write  RS(2) & " | "
	response.write  RS(3) & "</br>"
	
	RS.movenext
	rowcounter = rowcounter + 1
wend

rs.close

%>





