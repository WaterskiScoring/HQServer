<%



' -- Tests writing a Waiver file to waivers folder
sMemberID = "000001151"
sTourID = "16W999A"
html_var = "<table width=100 height=100><tr><td>TEST</td></tr></table>"

Write_PDFWaiver_ToFolder "WaiverReleaseOLR", html_var, sTourID, sMemberID, "waivers"






' ---------------------------------------------------------------------------------------
   SUB Write_PDFWaiver_ToFolder (file_prefix, html_var, sTourID, sMemberID, folder_name)
' ---------------------------------------------------------------------------------------


Dim PathtoOLRWaivers

' --- Define the path to the folder in the form required by file system object ---
PathtoOLRWaivers = Server.mappath("/")&"\rankings\"&folder_name
response.write("<br><br>PathtoOLRWaivers = ")
response.write(PathtoOLRWaivers)


' --- Create yyyymmdd format for date ---
WaiverYear = CStr(DATEPART("yyyy", NOW))
IF LEN(DATEPART("m", NOW))=1 THEN WaiverMonth = "0" + CStr(DATEPART("m", NOW)) ELSE WaiverMonth = CStr(DATEPART("m", NOW))
IF LEN(DATEPART("d", NOW))=1 THEN WaiverDay = "0" + CStr(DATEPART("d", NOW)) ELSE WaiverDay = CStr(DATEPART("d", NOW)) 	
WaiverDate = WaiverYear + WaiverMonth + WaiverDay

WaiverFilename = PathtoOLRWaivers & "\"&file_prefix&"_"&sTourID&"_MID_"&sMemberID&"_"&WaiverDate&".html"

response.write("<br> WaiverFilename = "&WaiverFilename)
response.write("<br> html_var = "&html_var)
'response.end


' --- Write the file to the web server ---
ty=2
IF ty=2 THEN
dim fs,f
set fs=Server.CreateObject("Scripting.FileSystemObject") 
set f=fs.CreateTextFile(WaiverFilename,true)
'f.write("Hello World!")
'f.write("How are you today?")
f.write(html_var)
f.close
set f=nothing
set fs=nothing

' response.write("<br><br>HERE")

END IF


END SUB


%>