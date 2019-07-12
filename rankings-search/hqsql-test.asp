<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<body>

<form method="post">
	Person ID : <input name="id" value="<%=Request.Form("id")%>" /> <input type="submit" name="submit" value="submit" />
</form>

<%

If Request.Form("submit") = "submit" Then
	
	Set SQLConnect = CreateObject("ADODB.Connection")
	SQLConnect.Open Application("HQSQLConn")
	
	Response.Write("<div>Member</div>")
	
	sql = "SELECT [Person ID] AS [PersonID], [First Name] AS [FirstName], [Last Name] AS [LastName], [MemberTypeID]"
	sql = sql + " FROM [tblPeople] WHERE ([Person ID] = " & Request.Form("id") & ");"
	
	Set rs = SQLConnect.Execute(sql)
	If rs.eof Then
		Response.Write("<div>EOF</div>")
	Else
		Response.Write("<table>")
		Response.Write("<tr>")
		Response.Write("<th>Person ID</th>")
		Response.Write("<th>First Name</th>")
		Response.Write("<th>Last Name</th>")
		Response.Write("<th>MemberTypeID</th>")
		Response.Write("</tr>")
		While Not rs.eof
			Response.Write("<tr>")
			Response.Write("<td>" & rs("PersonID") & "</td>")
			Response.Write("<td>" & rs("FirstName") & "</td>")
			Response.Write("<td>" & rs("LastName") & "</td>")
			Response.Write("<td>" & rs("MemberTypeID") & "</td>")
			Response.Write("</tr>")
			rs.MoveNext
		Wend
		Response.Write("</table>")
		rs.Close
	End If
	Set rs = Nothing
	
	Response.Write("<div>Head of Household</div>")

	sql = "SELECT [Sub Members].PrimaryPersonID AS HOHMemberID, tblPeople.[First Name] AS FirstName,"
	sql = sql + " tblPeople.[Last Name] AS LastName, tblPeople.MemberTypeID"
	sql = sql + " FROM [Sub Members]"
	sql = sql + " INNER JOIN tblPeople ON [Sub Members].PrimaryPersonID = tblPeople.[Person ID]"
	sql = sql + " WHERE (tblPeople.MemberTypeID = 3) AND ([Sub Members].SubMemberPersonID = " & Request.Form("id") & ");"
	
	Set rs = SQLConnect.Execute(sql)
	If rs.eof Then
		Response.Write("<div>EOF</div>")
	Else
		Response.Write("<table>")
		Response.Write("<tr>")
		Response.Write("<th>HOHMemberID</th>")
		Response.Write("<th>First Name</th>")
		Response.Write("<th>Last Name</th>")
		Response.Write("<th>MemberTypeID</th>")
		Response.Write("</tr>")
		While Not rs.eof
			Response.Write("<tr>")
			Response.Write("<td>" & rs("HOHMemberID") & "</td>")
			Response.Write("<td>" & rs("FirstName") & "</td>")
			Response.Write("<td>" & rs("LastName") & "</td>")
			Response.Write("<td>" & rs("MemberTypeID") & "</td>")
			Response.Write("</tr>")
			
			sHOHMemberID = rs("HOHMemberID")
			
			rs.MoveNext
		Wend
		Response.Write("</table>")
		rs.Close
	End If
	Set rs = Nothing
	
	If HOHMemberID > 0 Then
	
		Response.Write("<div>Sub Members of Head of Household</div>")
	
		sql = "SELECT [Sub Members].SubMemberPersonID AS [PersonID], tblPeople.[First Name] AS [FirstName],"
		sql = sql +" tblPeople.[Last Name] AS [LastName], tblPeople.MemberTypeID"
 		sql = sql + "FROM [Sub Members]"
		sql = sql + "INNER JOIN tblPeople" 
		sql = sql + "ON [Sub Members].SubMemberPersonID = tblPeople.[Person ID]" 
		sql = sql + "WHERE ([Sub Members].PrimaryPersonID = 119092);"
		'sql = sql + "WHERE ([Sub Members].PrimaryPersonID = '"&sHOHMemberID&"');"

		Set rs = SQLConnect.Execute(sql)
		If rs.eof Then
			Response.Write("<div>EOF</div>")
		Else
			Response.Write("<table>")
			Response.Write("<tr>")
			Response.Write("<th>Person ID</th>")
			Response.Write("<th>First Name</th>")
			Response.Write("<th>Last Name</th>")
			Response.Write("<th>MemberTypeID</th>")
			Response.Write("</tr>")
			While Not rs.eof
				Response.Write("<tr>")
				Response.Write("<td>" & rs("PersonID") & "</td>")
				Response.Write("<td>" & rs("FirstName") & "</td>")
				Response.Write("<td>" & rs("LastName") & "</td>")
				Response.Write("<td>" & rs("MemberTypeID") & "</td>")
				Response.Write("</tr>")
				rs.MoveNext
			Wend
			Response.Write("</table>")
			rs.Close
		End If
		Set rs = Nothing
		
	End If
	
	SQLConnect.Close
	Set SQLConnect = Nothing
	
End If

%>


</body>
</html>
