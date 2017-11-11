<!--#include virtual="/epl/functions.asp" -->


    
<html><head><title>Test Page</title>

    </head>
    <body>
    <DIV ID="splashScreen" STYLE="position:absolute;z-index:5;top:30%;left:35%;">
    <TABLE BGCOLOR="#000000" BORDER=1 BORDERCOLOR="#000000"	CELLPADDING=0 CELLSPACING=0 HEIGHT=150 WIDTH=300>
    <TR>
    <TD WIDTH="100%" HEIGHT="100%" BGCOLOR="#CCCCCC" ALIGN="CENTER" VALIGN="MIDDLE">
    <BR>
    <FONT FACE="Helvetica,Verdana,Arial" SIZE=2 COLOR="#000066">
    <B>Preparing your Registration Template.<br><br>
    This may take a minute or so ...<br><br><br>  
    </B></FONT>
    <IMG SRC="includes/wait.gif" BORDER=1 WIDTH=150 HEIGHT=15><BR><BR>
    </TD>
    </TR>
    </TABLE>
    </DIV>
    
<%

' Once the above "please wait" banner is written to HTML, we flush the response
' buffer to make the page appear to the users browser.  That sits on their display
' while the rest of the template preparation script processing takes place.
    

Function RemoveInvalidChars(strInput)
    dim workingstring
	'On Error Resume Next
	For i = 1 to Len(strInput)
		If isNumeric(Mid(strInput, i, 1)) then
			workingstring = workingstring & Mid(strInput, i, 1)
		End If
		If (Mid(strInput, i, 1)) => "a" and (Mid(strInput, i, 1)) <=  "z" then
			workingstring = workingstring & Mid(strInput, i, 1)
		End If
		If (Mid(strInput, i, 1)) => "A" and (Mid(strInput, i, 1)) <=  "Z" then
			workingstring = workingstring & Mid(strInput, i, 1)
		End If
		If (Mid(strInput, i, 1)) = "@" Or (Mid(strInput, i, 1)) = "." Then
				workingstring = workingstring & Mid(strInput, i, 1)
		End If
	Next
	RemoveInvalidChars = workingstring
	
End Function

'	-----------------------------------------------------------------------
'	Start by sucking Membership Pricing Info from HQ Table into local Array
'	-----------------------------------------------------------------------

Dim MT, MemPrice(200), MemUpgrd(200)


    

Dim objFSO
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
Dim path
path = Server.MapPath("Excel/")
'Randomize()
'Dim num



'""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
'"""""""""""""" With Scores and Ratings """""""""""""""""""""""
'""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""


'objFSO.CopyFile path & "/Templates/NCWSATemplateBlank.xls", path & "/template_with_scores.xls" , True
'objFSO.CopyFile path & "/Templates/NCWSATemplate2012.xls", path & "/template_with_scores.xls" , True

'Now open a connection to the new XLS file

Set objExcelConn = Server.CreateObject("ADODB.Connection")
'objExcelConn.Open "ExcelDSNwithScores"
'objExcelConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\webs\usawaterski.org\admin\excel\template_with_scores.xls; Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1;ReadOnly=0"";"

'this worked!!
objExcelConn.Open "Driver={Microsoft Excel Driver (*.xls)};DBQ=C:\junk\test.xls;ReadOnly=0;"
'objExcelConn.Open "Driver={Microsoft Excel Driver (*.xls)};DBQ=C:\webs\usawaterski.org\admin\excel\template_with_scores.xls;ReadOnly=0;"

Set objExcelSingleFields = Server.CreateObject("ADODB.Recordset")
objExcelSingleFields.ActiveConnection = objExcelConn 
objExcelSingleFields.CursorType = 3                    'Static cursor.
objExcelSingleFields.LockType = 2                      'Pessimistic Lock.

objExcelSingleFields.Source = "Select * from AMenTourName"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = "0"
objExcelSingleFields.update
objExcelSingleFields.close
		
		
'Now copy the file from Template to a file with the tournamentid
Dim filename
Dim filenamewithscores
'"06M123-Entries-SSSSSS-YYYYMMDD", 
filenamewithscores = "Entries-" & Session("StateList") & "-" & DateFmt

'Add the Tournament Name to the start of the file name
'session("TournamentName")
if len(session("TournamentName")) > 0 then
	'filename = "TournamentRegistrationFile-" & session("UserName") & ".xls"
	filenamewithscores = session("TournamentName") & "-" & filenamewithscores
end if

'5-18-2006 Remove any strange characters from the TournamentName
filenamewithscores = RemoveInvalidChars(filenamewithscores)

'Append the username
if len(session("UserName")) > 0 then
	'filename = "TournamentRegistrationFile-" & session("UserName") & ".xls"
	filenamewithscores = filenamewithscores & "-" & strTSanction & ".xls"
else
	'filename = "TournamentRegistrationFile.xls"
	filenamewithscores = filenamewithscores & ".xls"
end if

'objFSO.CopyFile path & "/template.xls", path & "/" & filename , True
'objFSO.CopyFile path & "/template_with_scores.xls", path & "/" & filenamewithscores , True
objFSO.CopyFile "c:/junk/test.xls", path & "/" & filenamewithscores , True

'Clean up old files
Set f = objFSO.GetFolder("c:\webs\usawaterski.org\admin\excel\")  
Set fc = f.Files 
Response.Write "<br>"
For Each f1 in fc
	'Response.Write f1.name 
	Set myfile = objFSO.GetFile("c:\webs\usawaterski.org\admin\excel\" & f1.name)
	'Response.Write  "Date:"  & myfile.DateCreated 
	'Response.Write  "Age:"  & datediff("d",myfile.DateCreated,date()) & "<br>"
	if datediff("d",myfile.DateCreated,date()) > 2 and left(myfile.name,8) <> "Template" then
		myfile.delete
	end if
	
Next  

Set f = nothing
Set fc = nothing

Set objFSO = Nothing

'Clean up old records in temp table

    
Response.Flush
      
' This final bit of HTML is written after processing is successfully completed
' to tell the user how to download their template, and where to go from here.
      
%>
    



<html>

<head>
<title>Create Pre-Registration Export v1.5</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" background = "/images/TopBackground.jpg" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" >

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#42639F">
      <p>&nbsp;</p>
      <p align="center"><font face="Verdana" size="6" color="#FFFFFF">
      	USA Water Ski NCWSA Registration Template</font></p>
      <p align="center"><font face="Verdana" size="4" color="#FFFFFF">
      	Registration Support for -- <%=session("TournamentName")%></font></p>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>  
  
<table border="0" cellspacing="0" cellpadding="0">  
  <tr> 
    <td width="185" valign="top" bgcolor="#42639F">

	<% If Session("aauth") then %>
	<font face="Verdana" size="2" COLOR="#FFFFFF"><br>&nbsp;Currently Logged in as: </font><br>
	<font face="Verdana" size="2" COLOR="#FFFFFF">&nbsp;<%=Session("UserName")%>&nbsp;&nbsp;
		<%=session("TournamentDate")%></font><br>
	<br>
	<% Else %>
	<font face="Verdana" size="2" COLOR="#FFFFFF"><br>&nbsp;Not currently logged in.</font>
	<% End If %>
	
			<font face="Verdana" size="2"> 
         <br>&nbsp;<a href="logout.asp"><font face="arial" COLOR="#FFFFFF">Log Out</font></a>&nbsp;<br>
			</font>
			<br>
	        &nbsp;<a href="/admin/index.asp"><font face="arial" size="2" COLOR="#FFFFFF">Back to Admin Index</font></a><br>&nbsp;<br>
	        &nbsp;<a href="http://www.usawaterski.org"><font face="arial" size="2" COLOR="#FFFFFF">USA Water Ski Home</font></a><br>&nbsp;<br>
			<br>
            <font face="Verdana" size="1">&nbsp;<font COLOR="#FFFFFF">Powered by</font> <a href="http://www.epolk.com"><font COLOR="#FFFFFF">ePolk.com</font></a><br></font>

  </td>

	<td>

  <table>
      <tr> 
         <td width="14">&nbsp;</td>
         <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><br>Your NCWSA
         Registration Export Excel Workbook is now complete and ready to download.&nbsp;</font>
         <br>&nbsp;<br>

         <font color="#FF0000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>
         !! New for 2011 -- Online Team Entry and Rotation Plan details now included !!
         </strong></font><br>
         <font size="2" face="Verdana, Arial, Helvetica, sans-serif">Details of Team Entry
         and Rotation Plans that have been prepared and submitted by the respective team
         captains through the new online team entry system, are now incorporated into this 
         Excel workbook.&nbsp; See the revised instructions section of your Excel workbook 
         for details.&nbsp; <font color="#FF0000"><strong>New for Fall 2011 -- A Registrar 
         Recap is now included.</strong></font>&nbsp; This new section in the Excel workbook 
         makes it easier for Registrars to see each team's overall entry status, to see which 
         entered skiers still need to execute event waivers locally, and to assess each team's 
         total entry fees.</font>
         <br>&nbsp;<br>

         <a href="excel/<% response.write filenamewithscores %>"><font face="Arial" size="2"><b>RIGHT 
         Click Here</b></font></a>&nbsp; <font size="2" face="Verdana, Arial, Helvetica, 
         sans-serif">to download your NCWSA Registration Template workbook, then select the 
         "Save As" option from that menu, and then choose a suitable location to store the 
         download file in your PC.&nbsp; After your Registration Template download has 
         completed, then open the Excel file from that location on your PC.&nbsp; It will 
         open automatically to an Instructions Tab section.&nbsp; Please review the material 
         in that section for the latest information on contents and usage.</font>
         <br>&nbsp;<br>

         <% IF AllowAccess THEN %>

         <font size="2" face="Verdana, Arial, Helvetica, sans-serif">If you are now doing 
         your <b><i>final and official</i></b> download of entries for this tournament, then 
         <b><i>after</i></b> downloading your Excel workbook (see the <b>RIGHT Click Here</b> 
         link in paragraph above), then click the <b>Close Registration</b> button that you
         see below.&nbsp; That will block any further modifications to existing Team Entry and 
         Rotation Plans, and refer team captains to the Tournament Registrar at the tournament 
         site for any last-minute changes.</font>
         <br>&nbsp;<br>
         
         <% ELSE %>

         <font size="2" face="Verdana, Arial, Helvetica, sans-serif">Online Entry to this 
         tournament is currently set to <b>Closed</b>.&nbsp; If you have not yet done your 
         <b><i>final and official</i></b> download of entries for this tournament, then 
         you may want to re-open Online Entry status, by clicking the <b>Re-Open Registration</b> 
         button below.</font>
         <br>&nbsp;<br>
         
         <% END IF %>

         </td>
      </tr>

 	</table>

	<TABLE ALIGN="CENTER" WIDTH=80%>
		
		<TR>

    <% IF AllowAccess THEN %>

		    <TD width=35% align=center>
			<form action="NCWSAChgRegStat.asp?TourID=<%=left(strTSanction,6)%>&Status=Close" method="post">
			<input type="submit" style="width:12em" value="Close Registration"
			title="Close Online Registration -- No further Changes by Captains allowed"></form>
 		   	</TD>

    <% ELSE %>

		    <TD width=35% align=center>
			<form action="NCWSAChgRegStat.asp?TourID=<%=left(strTSanction,6)%>&Status=Open" method="post">
			<input type="submit" style="width:12em" value="Re-Open Registration"
			title="Close Online Registration -- No further Changes by Captains allowed"></form>
 		   	</TD>

    <% END IF %>

	    <TD width=30% align=center>
		<form action="LookupMembers.asp?FormStatus=newsearch" method="post">
		<input type="submit" style="width:10em" value="Lookup Members"></form>
    	</TD>

	    <td width=25% align=center>     				
		<form action="Index.asp" method="post">
    <input type="submit" style="width:7em" value="Quit"></form>
 	    </td>
  	    
 	  </TR>

 	</TABLE>

  	  </td>
	  </tr>
</table>
</body>
</html>
