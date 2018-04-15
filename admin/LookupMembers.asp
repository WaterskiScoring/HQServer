<!--#include virtual="epl/functions.asp" -->

<% If not Session("aauth") then response.redirect "Login.asp" %>

<html>

<head>
<title>Lookup Members</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" >

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#42639F">
      <p>&nbsp;</p>
      <p align="center"><font face="Verdana" size="6" color="#FFFFFF">
      	USA Water Ski Member Lookup</font></p>
      <p align="center"><font face="Verdana" size="4" color="#FFFFFF">
      	Registration Support for -- <%=session("TournamentName")%></font></p>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>

<table border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="177" valign="top" bgcolor="#42639F">

<% 
Dim objConn
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open Application("WaterSkiConn")
%>
	<font face="Verdana" size="2" COLOR="#FFFFFF"><br>&nbsp;Currently Logged in as: </font><br>
	<font face="Verdana" size="2" COLOR="#FFFFFF">&nbsp;<%=Session("UserName")%>&nbsp;&nbsp;
		<%=session("TournamentDate")%></font><br>
	<br>
	
			<font face="Verdana" size="2"> 
			<br>&nbsp;<a href="logout.asp"><font face="arial" COLOR="#FFFFFF">Log Out</font></a>&nbsp;<br>
			</font>
			<br>
	        &nbsp;<a href="/admin/index.asp"><font face="arial" size="2" COLOR="#FFFFFF">Back to Admin Index</font></a><br>&nbsp;<br>
	        &nbsp;<a href="http://www.usawaterski.org"><font face="arial" size="2" COLOR="#FFFFFF">USA Water Ski Home</font></a><br>&nbsp;<br>
			<br>
            <font face="Verdana" size="1">&nbsp;<font COLOR="#FFFFFF">Powered by</font> <a href="http://www.epolk.com"><font COLOR="#FFFFFF">ePolk.com</font></a><br></font>
	
    </td>

<td width="600" >

<%


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


' ---------------------------------------------------
   FUNCTION PersonIDwChkDgt (PersonID)
' ---------------------------------------------------

' This function is given an integer "PersonID" value, and returns the
' 9-Character "PersonIDWithCheckDigit" value for that particular member.

Dim PIDSum, PIDChar, PIDLen, PIDPtr
 
PIDSum = 0: PIDChar = trim(PersonID): PIDLen = Len(PIDChar)

FOR PIDPtr = 1 TO PIDLen STEP 2
	PIDSum = PIDSum + (3*MID(PIDChar,PIDPtr,1))
	IF PIDPtr+1 <= PIDLen THEN PIDSum = PIDSum + MID(PIDChar,PIDPtr+1,1)
	NEXT

PersonIDwChkDgt = right(100-PIDSum,1) & Right(100000000+PersonID,8)

END FUNCTION


Function CalculateDivision(SkiAge, Gender)
Dim AgeDivision
if len(SkiAge) = 0 then
	AgeDivision = "-"
elseif SkiAge >= 0 AND SkiAge < 10 THEN '1' 
	AgeDivision = "1"
elseif  SkiAge >= 10 AND SkiAge < 14 THEN '2' 
	AgeDivision = "2"
elseif  SkiAge >= 14 AND SkiAge < 18 THEN '3' 
	AgeDivision = "3"
elseif  SkiAge >= 18 AND SkiAge < 25 THEN '1' 
	AgeDivision = "1"
elseif  SkiAge >= 25 AND SkiAge < 35 THEN '2' 
	AgeDivision = "2"
elseif  SkiAge >= 35 AND SkiAge < 45 THEN '3' 
	AgeDivision = "3"
elseif  SkiAge >= 45 AND SkiAge < 53 THEN '4' 
	AgeDivision = "4"
elseif  SkiAge >= 53 AND SkiAge < 60 THEN '5' 
	AgeDivision = "5"
elseif  SkiAge >= 60 AND SkiAge < 65 THEN '6' 
	AgeDivision = "6"
elseif  SkiAge >= 65 AND SkiAge < 70 THEN '7' 
	AgeDivision = "7"
elseif  SkiAge >= 70 AND SkiAge < 75 THEN '8' 
	AgeDivision = "8"
elseif  SkiAge >= 75 AND SkiAge < 80 THEN '9' 
	AgeDivision = "9"
elseif  SkiAge >= 80 AND SkiAge < 85 THEN 'A' 
	AgeDivision = "A"
elseif  SkiAge >= 85 THEN 'B' 
	AgeDivision = "B"
else
	AgeDivision = "-"
end if
					  
if Gender = "M" AND SkiAge < 18 THEN 'B' 
	SkiGender = "B"
elseif Gender = "M" AND SkiAge >= 18 THEN 'M' 
	SkiGender = "M"
elseif Gender = "F" AND SkiAge < 18 THEN 'G' 
	SkiGender = "G"
elseif Gender = "F" AND SkiAge >= 18 THEN 'W' 
	SkiGender = "W"
else 
	SkiGender = "-"
end if					  

CalculateDivision = SkiGender & AgeDivision
				  
End Function



Dim currentPage, sMemberID, sLastName, sFirstName, sState, sGender, FormStatus, LastPersID, NumMems

FormStatus = TRIM(Request("FormStatus"))
sMemberID = TRIM(Request("MemberID"))
sLastName = TRIM(Request("LastName"))
sFirstName = TRIM(Request("FirstName"))
sState = TRIM(Request("State"))
sGender = TRIM(Request("Gender"))

IF FormStatus = "newsearch" then sMemberID = "": sLastName = "": sFirstName = "": sState = "": sGender = ""

Dim RS
Set RS = Server.CreateObject("ADODB.RecordSet")
RS.ActiveConnection = objConn

'	set rs=Server.CreateObject("ADODB.recordset")

IF ((sMemberID = "" or NOT IsNumeric(sMemberID)) AND sLastName = "" AND sFirstName = "" AND sState = "" and sGender = "") or FormStatus = "revise" THEN 

' *****************************************************
'   Nothing was put in any of the 5 primary fields yet.
' *****************************************************

	DisplayMemberSearchFilters

' ************************************************
'   User entered something in at least one field
' ************************************************

ELSE

  ' Set up search criteria in MembersLive for members matching criteria

	sSQL = "SELECT TOP 11 Mem.PersonID, Mem.LastName, Mem.FirstName, Mem.City, Mem.State,"
	sSQL = sSQL + " Datepart(yyyy,Mem.BirthDate) as BirthYear, Left(Mem.Sex,1) as Sex" 
	sSQL = sSQL + " FROM USAWaterski.dbo.memberslive as Mem, USAWaterski.dbo.MembershipTypes as Typ"
	sSQL = sSQL + " WHERE Mem.MembershipTypeCode = Typ.MemberShipTypeID"
	sSQL = sSQL + " AND Typ.ExporttoTouramentRegistrationTemplate = 1"

	IF sMemberID <> "" and IsNumeric(sMemberID) THEN
		sSQL = sSQL + " AND Mem.PersonID = " & RemInvChr(right(sMemberID,8))
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
		sSQL = sSQL + " ORDER BY mem.PersonID"
	ELSE
		sSQL = sSQL + " ORDER BY mem.LastName, mem.FirstName"
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
			        <center><font face="Verdana, Arial, Helvetica, sans-serif" size="4" COLOR="#FFFFFF"><b>Select Desired Member from list Below</b></font>
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
								
				sMembAge = Session("TournamentYear") - rs("BirthYear") - 1

				%><tr>
       			  <TD ALIGN="Center" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=PersonIDwChkDgt(rs("PersonID"))%></a></FONT></TD>
           		  <TD ALIGN="Center" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="DisplayOneMember.asp?PersonID=<%=rs("PersonID")%>"><%=rs("LastName")%>,&nbsp;<%=rs("FirstName")%></a></FONT></TD>
         		  <TD ALIGN="Center" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;<%=rs("City")&", "&rs("State")%></a></FONT></TD>
         		  <TD ALIGN="Center" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;<% response.write(sMembAge&" / "&rs("Sex"))%></a></FONT></TD>
            </tr><% 

					LastPersID = rs("personid")
	        NumMems = NumMems + 1
					RS.MoveNext 

  			  LOOP

		IF NumMems = 1 then
			  rs.close
			  set rs=nothing
			  response.redirect "DisplayOneMember.asp?PersonID=" & LastPersID
		END IF
  			  
  			  %></table><br>
					<TABLE BORDER="0" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#DDDDFF" width=80% >
				  <tr><%  

				' Found more than 10 records in file - ie encourage to modify search  
				IF NumMems > 10 THEN
					
				%><td colspan=3 align=center>
				  <FONT COlOR="red" size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>More than ten members found.&nbsp;
				  	Only the first ten are displayed above.&nbsp; Please refine your search 
				  	parameters if the desired member is not one of those displayed.</b></font>
				  <br>&nbsp;<br>
				</td><%
				
			END IF
			
			 %></td>
			  <tr>
	
				    <TD width=30% align=center>
					<form action="LookupMembers.asp?FormStatus=newsearch" method="post">
					<input type="submit" style="width:9em" value="New Search"></form>
			    	</TD>

   			    <td width=30% align=center>     				
					<form action="LookupMembers.asp?FormStatus=revise&MemberID=<%=sMemberID%>&LastName=<%=sLastName%>&FirstName=<%=sFirstName%>&State=<%=sState%>&Gender=<%=sGender%>" method="post">
				  <input type="submit" style="width:9em" value="Revise Search"></form>
			   	 </td>

   			    <td width=30% align=center>     				
					<form action="Index.asp" method="post">
				  <input type="submit" style="width:9em" value="Quit"></form>
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
<form action="LookupMembers.asp?FormStatus=search" method="post">
<TABLE BORDER="4" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#DDDDFF" width=60% >
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








