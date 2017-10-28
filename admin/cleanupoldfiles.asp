<%	
'now clean up old files
Dim objFSO
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

	Set f = objFSO.GetFolder("d:\webs\usawaterski.org\admin\excel\")  
	Set fc = f.Files 
	Response.Write "<br>"
	For Each f1 in fc
		Response.Write f1.name 
		Set myfile = objFSO.GetFile("d:\webs\usawaterski.org\admin\excel\" & f1.name)
		Response.Write  "Date:"  & myfile.DateCreated 
		Response.Write  "Age:"  & datediff("d",myfile.DateCreated,date()) & "<br>"
		if datediff("d",myfile.DateCreated,date()) > 2 and left(myfile.name,10) = "Tournament" then
			myfile.delete
		end if
		
	Next  
	
	Set objFSO = nothing
	Set f = nothing
	Set fc = nothing

Set objFSO = Nothing

%>





