<!--#include virtual="epl/functions.asp" -->

<html>

<head>
<title>Membership Lookup</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" >

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#42639F">
      <p>&nbsp;</p>
      <p align="center"><font face="Verdana" size="6" color="#FFFFFF">
      	USA Water Ski Member Lookup</font></p>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>

<table border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="180" valign="top" bgcolor="#42639F">

			<br>
	        &nbsp;<a href="http://www.usawaterski.org"><font face="arial" size="2" COLOR="#FFFFFF">USA Water Ski Home</font></a><br>&nbsp;<br>
			<br>
            <font face="Verdana" size="1">&nbsp;<font COLOR="#FFFFFF">Powered by</font> <a href="http://www.epolk.com"><font COLOR="#FFFFFF">ePolk.com</font></a><br></font>
	
    </td>

<td width="600" >

<%

Dim objConn
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open Application("WaterSkiConn")

Function RemInvChr(strInput)
    dim workingstring
	On Error Resume Next
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
	RemInvChr = workingstring
End Function



Dim currentPage, sMemberID, sLastName, sFirstName, sState, sGender, FormStatus, LastMemID, NumMems

FormStatus = TRIM(Request("FormStatus"))
sMemberID = TRIM(Request("MemberID"))
sLastName = TRIM(Request("LastName"))
sFirstName = TRIM(Request("FirstName"))
sState = TRIM(Request("State"))
sGender = TRIM(Request("Gender"))

Session("ValidationAttempts") = 1

IF FormStatus = "newsearch" then sMemberID = "": sLastName = "": sFirstName = "": sState = "": sGender = ""

Dim RS
Set RS = Server.CreateObject("ADODB.RecordSet")
RS.ActiveConnection = objConn

'	set rs=Server.CreateObject("ADODB.recordset")

IF (sMemberID = "" AND sLastName = "" AND sFirstName = "" AND sState = "" and sGender = "") or FormStatus = "revise" THEN 

' *****************************************************
'   Nothing was put in any of the 5 primary fields yet.
' *****************************************************

	DisplayMemberSearchFilters

' ************************************************
'   User entered something in at least one field
' ************************************************

ELSE

  ' Set up search criteria in MemberTrak for members matching criteria

	sSQL = "SELECT TOP 11 Mem.PersonIDWithCheckDigit AS MemberID, Mem.LastName, Mem.FirstName,"
	sSQL = sSQL + " Mem.City, Mem.State, Datepart(yyyy,Mem.BirthDate) as BirthYear, Left(Mem.Sex,1) as Sex" 
	sSQL = sSQL + " FROM USAWaterski.dbo.members as Mem, USAWaterski.dbo.MembershipTypes as Typ"
	sSQL = sSQL + " WHERE Mem.MembershipTypeCode = Typ.MemberShipTypeID"
	sSQL = sSQL + " AND Typ.ExporttoTouramentRegistrationTemplate = 1"

	IF sMemberID <> "" THEN
		sSQL = sSQL + " AND Mem.PersonIDWithCheckDigit LIKE '%" & RemInvChr(sMemberID) & "%'"
	END IF
	
	IF sLastName <> "" THEN
		sSQL = sSQL + " AND lower(left(Mem.lastname," & len(sLastName) & ")) = '" & RemInvChr(lCASE(sLastName)) & "'"
	END IF
	
	IF sFirstName <> "" THEN
		sSQL = sSQL + " AND lower(left(Mem.firstname," & len(sFirstName) & ")) = '" & RemInvChr(lCASE(sFirstName)) & "'"
	END IF

	IF sState <> "" THEN
		sSQL = sSQL + " AND lower(Mem.state) = '" & RemInvChr(lCASE(sState)) & "'"
	END IF

	IF sGender <> "" THEN
		sSQL = sSQL + " AND lower(left(Mem.sex,1)) = '" & RemInvChr(lCASE(sGender)) & "'"
	END IF

	' Initial search based on user input in boxes
	
	IF sMemberID <> "" THEN
		sSQL = sSQL + " ORDER BY PersonIDWithCheckDigit"
	ELSE
		sSQL = sSQL + " ORDER BY LastName, FirstName"
	END IF


' *******************************************************
'	Display constructed query in debug log, then execute it
' *******************************************************

'	Set tempFSO=Server.CreateObject("Scripting.FileSystemObject")
'	IF Not (tempFSO.FileExists(Server.mappath("/")&"\..\" & "sql-debug-log.txt")) = true THEN
'		Set logobject=tempFSO.CreateTextFile(Server.mappath("/")&"\..\" & "sql-debug-log.txt",true)
'	ELSE
'		Set logobject=tempFSO.OpenTextFile(Server.mappath("/")&"\..\" & "sql-debug-log.txt",8,true)
'	END IF
'		logobject.WriteLine("SQL = " & sSQL & " -+- " & date() & " " & time() & " " & session("UserName"))
'		logobject.Close
'		Set logobject=nothing
'		Set tempFSO=nothing


 	RS.open sSQL


	' ******************************************
 	' No records found matching search criteria
	' ******************************************
	
	IF RS.EOF THEN 
		%><table align="Center"><TD><font face="Verdana, Arial, Helvetica, sans-serif" size="3" Color="Red">
		    <b>&nbsp;<br>No Members found for these Specs -- pls try again.</b></font></TD></table><%

		DisplayMemberSearchFilters

	' **************************************************
	' Found at least ONE or MANY members in Member Table
	' **************************************************

	ELSE

		NumMems = 0

			%>
			<TABLE BORDER="4" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#DDDDFF" width=90% >
			  <tr><br>
		  	         <td BGCOLOR="#42639F">
			        <center><font face="Verdana, Arial, Helvetica, sans-serif" size="4" COLOR="#FFFFFF"><b>Click on Member Name in List Below</b></font>
			        <br>
			      </td>
			  </tr>  
			<br>
			  <tr>
			     <td>
				  <br>
				<TABLE BORDER="1" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#F5F5F5" width=95% >
			          <TR>
			       	    <TD ALIGN="Center" vAlign="top"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Member ID</b></FONT></TD>
			            <TD ALIGN="Center" vAlign="top"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>LastName, First</b></FONT></TD>
        			    <TD ALIGN="Center" vAlign="top"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>City & State</b></FONT></TD>
        			    <TD ALIGN="Center" vAlign="top"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Age / Gender</b></FONT></TD>
			          </TR>

		          <TR><%

		    DO WHILE NOT RS.EOF 
								
				sMembAge = Datepart("yyyy",now) - rs("BirthYear") - 1

				%><tr>
       			  <TD ALIGN="Center" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("memberid")%></a></FONT></TD>
           		<TD ALIGN="Center" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="ValidateMember.asp?MemberID=<%=rs("memberid")%>"><%=rs("LastName")%>,&nbsp;<%=rs("FirstName")%></a></FONT></TD>
         		  <TD ALIGN="Center" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;<%=rs("City")&", "&rs("State")%></a></FONT></TD>
         		  <TD ALIGN="Center" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;<% response.write(sMembAge&" / "&rs("Sex"))%></a></FONT></TD>
        </tr><% 

					LastMemID = rs("memberid")
	        NumMems = NumMems + 1
					RS.MoveNext 

  			  LOOP

		IF NumMems = 1 then
			  rs.close
			  set rs=nothing
				response.redirect "ValidateMember.asp?MemberID=" & LastMemID
		END IF
  			  
  			  %></table><br>
					<TABLE BORDER="0" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#DDDDFF" width=80% >
				  <tr><%  

				' Found more than 10 records in file - ie encourage to modify search  
				IF NumMems > 10 THEN
					
				%><td colspan=3 align=center>
				  <FONT COlOR="red" size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>More 
				  	than ten members found.&nbsp; Only the first ten are displayed above.&nbsp; 
				  	Please refine your search parameters if you are not one of those displayed.</b></font>
				  <br>&nbsp;<br>
				</td><%
				
			END IF
			
			 %></td>
			  <tr>
	
				    <TD width=32% align=center>
					<form action="FindMember.asp?FormStatus=newsearch" method="post">
					<input type="submit" style="width:8em" value="New Search"></form>
			    	</TD>

   			    <td width=32% align=center>     				
					<form action="FindMember.asp?FormStatus=revise&MemberID=<%=sMemberID%>&LastName=<%=sLastName%>&FirstName=<%=sFirstName%>&State=<%=sState%>&Gender=<%=sGender%>" method="post">
				   <input type="submit" style="width:8em" value="Revise Search"></form>
			   	</td>

   			    <td width=32% align=center>     				
					<form action="Index.asp" method="post">
					<input type="submit" style="width:8em" value="Quit"></form>
			  	   </td>
	
	       </TR>

		        </table> 

			</table><%    

			rs.Close
			
	 END IF

END IF

set rs = nothing




' ---------------------------------
   SUB DisplayMemberSearchFilters
' ---------------------------------

%>
<br>
<form action="FindMember.asp?FormStatus=search" method="post">


<TABLE BORDER="4" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#DDDDFF" width=60% >
	
	<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><p>Use the following
	search function to help you locate your existing member credentials.&nbsp; 
	<font color=red><b>You do <i>not</i> have to fill in <i>every</i>	field 
	below</b></font>.&nbsp; If you have your member number readily at hand, then
	certainly enter that alone.&nbsp; Otherwise, you may enter only enough of your
	first and/or last name to enable me to find you in the Membership table.</p>

	<p>If you have a common name, or if the list returned by an earlier search selected 
	more members than	could be listed (and you were not among those), then you can add 
	state or gender to the specifications below to tighten the selection, and then 
	search again.</p>
	
	<p>Once you have located your membership, you will then be taken to a screen 
	where you will be asked to confirm certain details in your membership record, 
	to ensure security.</p></font>
	
  <tr><br>
      <td BGCOLOR="#42639F">
        <center><font face="Verdana, Arial, Helvetica, sans-serif" size="4" COLOR="#FFFFFF"><b>Enter Member Search Specifications</b></font>
        <br>
      </td>
  </tr>  

  <tr>
     <td>
	<br>
	<TABLE BORDER="1" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#DDDDFF" width=95% >
          <TR>
            <TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Member ID</FONT></Center></TD>
            <TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Last Name</FONT></Center></TD>
            <TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">First Name</FONT></Center></TD>
            <TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">State</FONT></Center></TD>
            <TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Gender (M/F)</FONT></Center></TD>
          </TR>

          <TR>
            <TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><input type="text" name="MemberID" value="<%=sMemberID%>" maxlength=11 size=13></input></FONT></Center></TD>
            <TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><input type="text" name="LastName" value="<%=sLastName%>" maxlength=15 size=18></input></FONT></Center></TD>
            <TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><input type="text" name="FirstName" value="<%=sFirstName%>" maxlength=15 size=18></input></FONT></Center></TD>
            <TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><input type="text" name="State" value="<%=sState%>" maxlength=2 size=4></input></FONT></Center></TD>
            <TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><input type="text" name="Gender" value="<%=sGender%>" maxlength=1 size=3></input></FONT></Center></TD>
          </TR>
        </table> 

	<TABLE BORDER="0" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="<%=TableColor1%>" width=60% >  
	  <tr>	
	    <td>
	      <br>
	          <input type="hidden" name="pvar" value="SelectMember">
	          <center><input type="submit" value="Begin Search">

	    </td>
	   </tr>
	</TABLE>

     </td>
  </tr>
</TABLE>
</form>

<%

END SUB

%>









