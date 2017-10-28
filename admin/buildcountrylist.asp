<%
Dim objConn
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open Application("WaterSkiConn")
Dim objRS
Set objRS = Server.CreateObject("ADODB.RecordSet")
objRS.ActiveConnection = objConn
objRS.Open "SELECT * FROM Countries ORDER BY Country_Name"

Dim Stuff, myFSO, WriteStuff

'this line creates an instance of the File Scripting Object named myFSO
Set myFSO = CreateObject("Scripting.FileSystemObject")
'this line opens the file, notice the 1, it will cause the script to write to the file 
'(overwriting existing text)
Set WriteStuff = myFSO.OpenTextFile(Server.MapPath("/admin/includes/countries.asp"), 2, True)

'Writ eout USA First
'<option value="USA" selected>USA 
Stuff = "<option value=""US"" selected>USA"
WriteStuff.WriteLine(Stuff)

Do while not objRS.EOF
	if objRS("Country_Name") = "USA" then
		'skip writing out since we already did this by hand at the top
	else
		Stuff = "<option value=""" & objRS("Country_Name") & """>" & objRS("Country_Name")
		WriteStuff.WriteLine(Stuff)
	end if



'Parameter Description 
'fname  Required. The name of the file to open 
'mode  Optional. How to open the file 
'1=ForReading - Open a file for reading. You cannot write to this file.
'2=ForWriting - Open a file for writing.
'8=ForAppending - Open a file and write to the end of the file.
 
'create  Optional. Sets whether a new file can be created if the filename does not exist. True indicates that a new file can be created, and False indicates that a new file will not be created. False is default 
'format  Optional. The format of the file 
'0=TristateFalse - Open the file as ASCII. This is default.
'-1=TristateTrue - Open the file as Unicode.
'-2=TristateUseDefault - Open the file using the system default.
 

'this line actually writes STUFF from above to the file


	objRS.MoveNext
Loop

''this line closes the file
WriteStuff.Close

'this line destroys the instance of the File Scripting Object named WriteStuff
SET WriteStuff = NOTHING

'this line destroys the instance of the File Scripting Object named myFSO
SET myFSO = NOTHING

%>
             





