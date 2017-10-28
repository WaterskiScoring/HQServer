<!--#include virtual="/admin/includes/security.asp" -->
<html>

<head>
<title>Admin Index</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" background = "/templates/images/TopBackground.jpg" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" >
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="182" bgcolor="#42639F" valign="top"></td>
    <td bgcolor="#42639F">
      <p>&nbsp;</p>
      <p align="center"><font face="Verdana" size="6" color="#FFFFFF">USA Water 
        Ski Admin</font></p>
      <p>&nbsp;</p>
    </td>
  </tr>
  <tr> 
   <td bgcolor="#42639F">
<!--#include virtual="/admin/includes/menu.asp" -->
  </td>
    <td valign="top" >
	
<% 
Dim msFolder
  Dim msBaseFolder
  Dim msAccessEmpty
  Dim msAccessFull
  Dim msProject
  Dim msFTPFile
  Dim msFTPFileLog
  Dim msToday
  Dim msDTSServer
  Dim msDTSUser
  Dim msDTSPwd
  Dim msDTSLog
  Dim msDTSPkg




  ' SQL Server connection settings and DTS Package name

  msDTSServer = "(local)"
  msDTSUser = "waterski"
  msDTSPwd = "usa456"
  msDTSPkg = "Copy USA Member data to DBF"


'remove the DTS package log
RemoveFile "D:\webs\usawaterski.org\admin\dbf\log.txt"
RemoveFile "D:\webs\usawaterski.org\admin\dbf\member.dbf"
'copy the result file to a random filename for the user

  Set oFS = CreateObject("Scripting.FileSystemObject") 
 
     oFS.CopyFile "D:\webs\usawaterski.org\admin\dbf\MEMBER_empty.DBF","D:\webs\usawaterski.org\admin\dbf\MEMBER.DBF"
  
  Set oFS = Nothing 


ProcessJob 



 Sub ProcessJob()

  'on error resume next

    Const DTSSQLStgFlag_Default = 0
    Const DTSStepExecResult_Failure = 1
    Const ForWriting = 2 

    Dim oPkg
    Dim oStep
    Dim bStatus
    Dim oFS
    Dim oFile

  ' Remove files from the previous instances of this job
    RemoveFile msFolder & "\files\"  & msAccessEmpty

 ' Copy the empty.mdb shell to our working folder

   ' Execute our DTS Package stored in SQL Server

    Set oPkg = CreateObject("DTS.Package")
    oPkg.LoadFromSQLServer msDTSServer,msDTSUser,msDTSPwd,DTSSQLStgFlag_Default,"","","",msDTSPkg

    oPkg.Execute()
	
    bStatus = True
	
   ' Write a log file displaying success or failure for each package step.



    response.write "<html><body>" 
    'response.write "Execution time: " & Now() 

    For Each oStep In oPkg.Steps

        'response.write "   Pkg Step " & oStep.Name & " "

        If oStep.ExecutionResult = DTSStepExecResult_Failure Then
           response.write " failed " & oStep.errorcode & " <br>"
           bStatus = False
        Else
           'response.write " succeeded<br>"
        End If
 
        'response.write "Task """ & oPkg.Tasks(oStep.TaskName).Description & """<br>"

   Next
	
   If bStatus Then
      'response.write "Package " & oPkg.Name & " succeeded<br>"
      response.write "<br>"
      response.write  "<a href=""/admin/dbf/member.dbf""><font face=""Arial"" size=""2"">Right click here to download the current member.dbf file.</font></a>" 
      response.write "<br>"
   Else
      response.write "Package " & oPkg.Name & " failed<br>"
   End If
	
   Set oPkg = nothing	 

   'response.write  "Done: " & Now() & " " &  "<br></body></html>"


 End Sub
 

 
 

 Sub RemoveFile(sFilePathAndName) 

  Set oFS = CreateObject("Scripting.FileSystemObject") 
   
  If oFS.FileExists(sFilePathAndName) = True Then 
     oFS.DeleteFile sFilePathAndName, True 
  end if 

  Set oFS = Nothing 
   
 End Sub 

 Sub CopyFile(sFileFromFolder,sFileFrom,sFileToFolder,sFileTo) 

  Set oFS = CreateObject("Scripting.FileSystemObject") 
 
     oFS.CopyFile sFileFromFolder & "\"  & sFileFrom,sFileToFolder & "\" & sFileTo
  
  Set oFS = Nothing 
   
 End Sub 

 Sub RenameFile(sFolder,sFileFrom,sFileTo) 

  Set oFS = CreateObject("Scripting.FileSystemObject") 
 
     oFS.MoveFile sFolder & "\"  & sFileFrom,sFolder & "\"  & sFileTo 
  
  Set oFS = Nothing 
   
 End Sub 


Sub CopyFile(sFileFromFolder,sFileFrom,sFileToFolder,sFileTo) 
  Set oFS = CreateObject("Scripting.FileSystemObject") 
  oFS.CopyFile sFileFromFolder & "\"  & sFileFrom,sFileToFolder & "\" 
  Set oFS = Nothing 
  
End Sub 

%>	
	</td>
  </tr>
</table>
</body>
</html>





