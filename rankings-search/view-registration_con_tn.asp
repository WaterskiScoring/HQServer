<%

Server.ScriptTimeout = 3000 

Dim Con
Dim comm
Dim rs
Dim sSQL

SUB OpenCon_special
	 Response.Write("Opening connection.<br><br>")
     Set Con = Server.CreateObject("ADODB.Connection")
     Con.ConnectionTimeout = 3000
     Con.Open "Provider=SQLOLEDB;SERVER=10.208.224.87;uid=U54W5r4NK;Password=4QU4r4nkSAA;Initial Catalog=cobra00025"
     Con.CommandTimeout = 30 'This is in seconds. 3000 seconds would be 50 minutes
END SUB

SUB CloseCon_special
  Response.Write("Closing connection.<br><br>")
  Con.close
  Set Con = Nothing
END SUB

' Wait for 60 seconds and then bring back some rows.
sSQL = "WAITFOR DELAY '00:01'; SELECT TOP 10 * FROM [usawsrank].[IAC_TEMP];"

' Open the database connection
OpenCon_special

' Create the Command object and set some key properties
Set comm = Server.CreateObject("ADODB.Command")
comm.activeConnection = Con
comm.commandText = sSQL
comm.commandTimeout = 30

Response.Write("State = " & Con.State & "<br>")
Response.Write("SQL = " & Comm.commandText & "<br>")
Response.Write("Timeout = " & Comm.commandTimeout & "<br><br>")

On Error Resume Next

' Timestamps before and after
Response.Write("Current Time= " & Now & "<br>")
Set rs = comm.Execute(, , 1) 'adCmdText
Response.Write("Current Time= " & Now & "<br>")

If Err <> 0 Then
	Response.Write(Err.Source & ": " & Err.Description & "<br><br>")
Else
	Response.Write("Recordset retrieved." & "<br><br>")
End If

' Close the database connection
CloseCon_special 

%>

<body>
<form id="form1" name="form1" method="post" action="">
  <label>End of test.</label>
  </p>
</form>
</body>
</html>

