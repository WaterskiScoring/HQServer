<!--#include file="clsUpload.asp"-->
<!--#include file="settingsHQ.asp"-->

<%
Dim objUpload
Dim objfso
Dim strFileName
Dim strPath

set objFSO=server.createobject("scripting.filesystemObject")

' Instantiate Upload Class
Set objUpload = New clsUpload

' Grab the file name
strFileName = objUpload.Fields("File1").FileName

' Compile path to save file to
'strPath = Server.MapPath("/rankings/uploads/") & "\" & strFileName
strPath=PathtoUploads & "\" & strFileName



if objfso.FileExists(strPath) = true Then
  ' Reject upload if the file already exists
   WriteLog(date() &"  "& time() &"  "& strpath & " duplicate file upload attempted.  File rejected.")
   set objfso=nothing
   set objupload=nothing
   WriteIndexPageHeader
   %>
    <html><head><title>File Upload Failed</title></head><body>
    <br><br>
    <center><h2><font color="red">The file <%=strFileName%> already exists.</font></h2><br><br><br>

    <h4>Please check the file name or contact your Regional Seeding Committee member or Headquarters for further assistance.
    <br><br><br>
    </h4></center>
    <br><br>
    </body></html>
   <%
   WriteIndexPageFooter
Else
  ' Save the binary data to the file system if it doesn't exist
  objUpload("File1").SaveAs strPath
  ' Release upload object from memory
  Set objfso = Nothing
  Set objUpload = Nothing
  Response.Redirect "/rankings/verify_upload.asp?file="&strfilename
End if


%>





