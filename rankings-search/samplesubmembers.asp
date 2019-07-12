
<%

Set SQLConnect = CreateObject("ADODB.Connection")
SQLConnect.Open Application("HQSQLConn")
Dim SQLString
SQLString = "SELECT tblPeople.[Person ID], tblPeople.[First Name], tblPeople.[Last Name] FROM tblPeople "
SQLString = SQLString & "INNER JOIN [Sub Members] ON tblPeople.[Person ID] = [Sub Members].SubMemberPersonID "
SQLString = SQLString & "WHERE  ([Sub Members].PrimaryPersonID = 121105)"

Set RS = SQLConnect.Execute(SQLString)

While Not RS.eof 
	response.write  RS("First Name") & " | "
	response.write  RS("Last Name") & " | "
	response.write  RS("Person ID") &  "</br>"
	
	RS.movenext

wend

rs.close

%>





