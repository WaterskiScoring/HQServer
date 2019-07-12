//Ignore all the VBSTART..VBEND stuff and edit lines 125 to 127 for your URL,file,fieldname

VBSTART
'Cribbed from:
'http://www.motobit.com/tips/detpg_uploadvbsie/
'VBScript Code is Copyright (C) 2001 Antonin Foller, PSTRUH Software

'******************* upload - begin
'Upload file using input type=file
Function UploadFile(DestURL, FileName, FieldName)
  'Boundary of fields.
  'Be sure this string is Not In the source file
  Const Boundary = "---------------------------0123456789012"
  
  Dim FileContents, FormData
  'Get source file As a binary data.
  FileContents = GetFile(FileName)
  
  'Build multipart/form-data document
  FormData = BuildFormData(FileContents, Boundary, FileName, FieldName)
  
  'Post the data To the destination URL
  UploadFile = IEPostBinaryRequest(DestURL, FormData, Boundary)
End Function

'Build multipart/form-data document with file contents And header info
Function BuildFormData(FileContents, Boundary, FileName, FieldName)
  Dim FormData, Pre, Po
  Const ContentType = "application/upload"
  
  'The two parts around file contents In the multipart-form data.
  Pre = "--" + Boundary + vbCrLf + mpFields(FieldName, FileName, ContentType)
  Po = vbCrLf + "--" + Boundary + "--" + vbCrLf
  
  'Build form data using recordset binary field
  Const adLongVarBinary = 205
  Dim RS: Set RS = CreateObject("ADODB.Recordset")
  RS.Fields.Append "b", adLongVarBinary, Len(Pre) + LenB(FileContents) + Len(Po)
  RS.Open
  RS.AddNew
    Dim LenData
    'Convert Pre string value To a binary data
    LenData = Len(Pre)
    RS("b").AppendChunk (StringToMB(Pre) & ChrB(0))
    Pre = RS("b").GetChunk(LenData)
    RS("b") = ""
    
    'Convert Po string value To a binary data
    LenData = Len(Po)
    RS("b").AppendChunk (StringToMB(Po) & ChrB(0))
    Po = RS("b").GetChunk(LenData)
    RS("b") = ""
    
    'Join Pre + FileContents + Po binary data
    RS("b").AppendChunk (Pre)
    RS("b").AppendChunk (FileContents)
    RS("b").AppendChunk (Po)
  RS.Update
  FormData = RS("b")
  RS.Close
  BuildFormData = FormData
End Function

'sends multipart/form-data To the URL using IE
Function IEPostBinaryRequest(URL, FormData, Boundary)
  'Create InternetExplorer
  Dim IE: Set IE = CreateObject("InternetExplorer.Application")
  
  'You can uncoment Next line To see form results
  'IE.Visible = True
   
  'Send the form data To URL As POST multipart/form-data request
  IE.Navigate URL, , , FormData, _
    "Content-Type: multipart/form-data; boundary=" + Boundary + vbCrLf

  Do While IE.Busy
    Wait 1, "Upload To " & URL
  Loop
  
  'Get a result of the script which has received upload
  On Error Resume Next
  IEPostBinaryRequest = IE.Document.body.innerHTML
  IE.Quit
End Function

'Infrormations In form field header.
Function mpFields(FieldName, FileName, ContentType)
  Dim MPTemplate 'template For multipart header
  MPTemplate = "Content-Disposition: form-data; name=""{field}"";" + _
   " filename=""{file}""" + vbCrLf + _
   "Content-Type: {ct}" + vbCrLf + vbCrLf
  Dim Out
  Out = Replace(MPTemplate, "{field}", FieldName)
  Out = Replace(Out, "{file}", FileName)
  mpFields = Replace(Out, "{ct}", ContentType)
End Function


Sub Wait(Seconds, Message)
  On Error Resume Next
  'CreateObject("wscript.shell").Popup Message, Seconds, "", 64
End Sub


'Returns file contents As a binary data
Function GetFile(FileName)
  Dim Stream: Set Stream = CreateObject("ADODB.Stream")
  Stream.Type = 1 'Binary
  Stream.Open
  Stream.LoadFromFile FileName
  GetFile = Stream.Read
  Stream.Close
End Function

'Converts OLE string To multibyte string
Function StringToMB(S)
  Dim I, B
  For I = 1 To Len(S)
    B = B & ChrB(Asc(Mid(S, I, 1)))
  Next
  StringToMB = B
End Function
'******************* upload - end
VBEND

//Modify these lines
Let>URL=http://www.mjtnet.com/demos/uploadfile.php
Let>file=C:\Users\User\Documents\myfile.txt
Let>fieldname=uploaded
VBEval>UploadFile("%URL%","%file%","%fieldname%"),result
MessageModal>result