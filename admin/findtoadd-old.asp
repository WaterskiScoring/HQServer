<!--#include virtual="epl/functions.asp" -->

<html>

<head>
<title>USA Water Ski Member Lookup</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" >

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#42639F">
      <p>&nbsp;</p>
      <p align="center"><font face="Verdana" size="6" color="#FFFFFF">
      	<b>USA Water Ski Membership Lookup</b></font></p>
      <p align="center"><font face="Verdana" size="4" color="#FFFFFF">
      	<%=Session("MemName")%>&nbsp;&nbsp; as Administrator for:&nbsp;&nbsp;&nbsp;&nbsp;<%=Session("TeamName")%>&nbsp;&nbsp; ( <%=Session("TeamID")%> ) </font></p>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>

<table border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="180" valign="top" bgcolor="#42639F">

			<br>
	        &nbsp;<a href="https://www.usawaterski.org/members/"><font face="arial" size="2" COLOR="#FFFFFF">USA Water Ski Home</font></a><br>&nbsp;<br>
			<br>
            <font face="Verdana" size="1">&nbsp;<font COLOR="#FFFFFF">Powered by</font> <a href="http://www.epolk.com"><font COLOR="#FFFFFF">ePolk.com</font></a><br></font>
	
    </td>

<td width="760" >

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


'	***** Bailout to Members Login if not auth or no Session("id") value

IF not Session("auth") or Session("id") < 1 or Session("TeamID") = "" then response.redirect "https://www.usawaterski.org/members/login/index.asp"

Dim currentPage, sMemberID, sLastName, sFirstName, sState, sGender, FormStatus, LastMemID, NumMems

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

IF (sMemberID = "" AND sLastName = "" AND sFirstName = "" AND sState = "" and sGender = "") or FormStatus = "revise" THEN 

' *****************************************************
'   Nothing was put in any of the 5 primary fields yet.
' *****************************************************

	DisplayMemberSearchFilters

' ************************************************
'   User entered something in at least one field
' ************************************************

ELSE

  ' Set up search criteria in Member Table for members matching criteria
  ' Note that this NCWSA-Specific logic pulls existing NCWSA team codes,
  ' and also explicitly excludes all existing members of the active team.

	sSQL = "SELECT TOP 11 Mem.PersonIDWithCheckDigit AS MemberID, Mem.LastName, Mem.FirstName,"
	sSQL = sSQL & " Mem.City, Mem.State, Datepart(yyyy,Mem.BirthDate) as BirthYear, Left(Mem.Sex,1)" 
	sSQL = sSQL & " as Sex, Case when TR.TeamID is NULL then '---' else TR.TeamID end as TeamID" 
	sSQL = sSQL & " FROM USAWaterski.dbo.members as Mem JOIN USAWaterski.dbo.MembershipTypes"
	sSQL = sSQL & " as Typ ON Mem.MembershipTypeCode = Typ.MemberShipTypeID"
	sSQL = sSQL & " Left Join (Select MemberID, Substring(Max(Convert(Char(10),LastEvent,111)+Team),"
	sSQL = sSQL & " 11, Len(Max(Convert(Char(10),LastEvent,111)+Team))-10) as TeamID"
	sSQL = sSQL & " from Cobra00025.USAWSRank.TeamRoster Group By MemberID) as TR"
	sSQL = sSQL & " on TR.MemberID = Mem.PersonIDWithCheckDigit"
	sSQL = sSQL & " WHERE Typ.ExporttoTouramentRegistrationTemplate = 1"
	sSQL = sSQL & " AND (TR.TeamID is Null or TR.TeamID <> '" & Session("TeamID") & "')"	

	IF sMemberID <> "" THEN
		sSQL = sSQL & " AND Mem.PersonIDWithCheckDigit LIKE '%" & RemInvChr(sMemberID) & "%'"
	END IF
	
	IF sLastName <> "" THEN
		sSQL = sSQL & " AND lower(left(Mem.lastname," & len(sLastName) & ")) = '" & RemInvChr(lCASE(sLastName)) & "'"
	END IF
	
	IF sFirstName <> "" THEN
		sSQL = sSQL & " AND lower(left(Mem.firstname," & len(sFirstName) & ")) = '" & RemInvChr(lCASE(sFirstName)) & "'"
	END IF

	IF sState <> "" THEN
		sSQL = sSQL & " AND lower(Mem.state) = '" & RemInvChr(lCASE(sState)) & "'"
	END IF

	IF sGender <> "" and left(FormStatus,5) = "Indiv" THEN
		sSQL = sSQL & " AND lower(left(Mem.sex,1)) = '" & RemInvChr(lCASE(sGender)) & "'"
	END IF

	' Initial search based on user input in boxes
	
	IF sMemberID <> "" THEN
		sSQL = sSQL & " ORDER BY PersonIDWithCheckDigit"
	ELSE
		sSQL = sSQL & " ORDER BY LastName, FirstName"
	END IF


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
			      <br></td>
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
	        			    <TD ALIGN="Center" vAlign="top"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Age/Gender</b></FONT></TD>
	        			    <TD ALIGN="Center" vAlign="top"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Team</b></FONT></TD>
			          </TR>

		          <TR><%

		    DO WHILE NOT RS.EOF 
								
				sMembAge = Datepart("yyyy",now) - rs("BirthYear")

				%><tr>

       			  <TD ALIGN="Center" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("memberid")%></FONT></TD>
	           	  <TD ALIGN="Center" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="rostermanager.asp?FormStatus=AddToTeam&MemberID=<%=rs("memberid")%>"
		           	  title="Click here to add <%=rs("FirstName")%> to your Team Roster"><%=rs("LastName")%>,&nbsp;<%=rs("FirstName")%></a></FONT></TD>
         		  <TD ALIGN="Center" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;<%=rs("City")&", "&rs("State")%></FONT></TD>
 	        		  <TD ALIGN="Center" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;<% response.write(sMembAge&" / "&rs("Sex"))%></FONT></TD>
       			  <TD ALIGN="Center" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("TeamID")%></FONT></TD>

        </tr><% 

					LastMemID = rs("memberid")
	        NumMems = NumMems + 1
					RS.MoveNext 

  			  LOOP

  			  
  			  %></table><br>
					<TABLE BORDER="0" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#DDDDFF" width=80% >
				  <tr><%  

				' Found more than 10 records in file - ie encourage to modify search  
				IF NumMems > 10 THEN
					
				%><td colspan=3 align=center>
				  <FONT COlOR="red" size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>More 
				  	than ten members found.&nbsp; Only the first ten are shown.&nbsp; 
				  	Please refine search parameters if the desired member is not listed.</b></font>
				  <br>&nbsp;<br>
				</td><%
				
			END IF
			
			 %></td>
			  <tr>
	
				    <TD width=32% align=center>
					<form action="FindToAdd.asp?FormStatus=newsearch" method="link">
					<input type="submit" style="width:8em" value="New Search"></form>
			    	</TD>

   			    <td width=32% align=center>     				
					<form action="FindToAdd.asp?FormStatus=revise&MemberID=<%=sMemberID%>&LastName=<%=sLastName%>&FirstName=<%=sFirstName%>&State=<%=sState%>&Gender=<%=sGender%>" method="post">
				   <input type="submit" style="width:8em" value="Revise Search"></form>
			   	</td>

   			    <td width=32% align=center>     				
					<form action="rostermanager.asp" method="link">
					<input type="submit" style="width:9em" value="Return to Roster"></form>
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
<form action="FindToAdd.asp" method="post">


<TABLE BORDER="4" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#DDDDFF" width=75% >
	
	<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><p>Use the following
	search function to help you locate existing USA Waterski members to add to your
	NCWSA Team Roster.&nbsp; <font color=red>You do <b><i>not</i></b> have to fill 
	in <b><i>every</i></b> field below</b></font>.&nbsp; If you have a member number 
	for the person you wish to add, then certainly enter that alone.&nbsp; Otherwise, 
	you may enter just enough of the first and/or last name, to enable me to find
	that person in the Membership table.</p>

	<p>If you are looking for a common name, or if the list returned by an earlier 
	search selected more members than could be listed (and the one you want was not 
	among those), then you can add state or gender to the specifications below, to 
	tighten the selection, and then search again.</p>
	
	<p>Once you have located your desired member, then you can click that person's 
	name to add them to your NCWSA Team Roster.&nbsp; <font color=red><b>Caution !!
	&nbsp; </b></font> an NCWSA Team Code appearing with a member indicates that
	person has recorded scores for another NCWSA Team -- if so, be sure this is
	the correct individual who is transferring to your school.&nbsp; Otherwise you
	may not have the person you're looking for.</p></font>
	
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

	    <td>&nbsp;<br>
	          <center><input type="submit" name="FormStatus" value="Search for Member">
	    </td>

		</form>	    

	    <td>&nbsp;<br>
	          <center>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	    </td>
	    
	    <td>&nbsp;<br>
	          <center><form action="rostermanager.asp" method="link">
					<input type="submit" style="width:9em" value="Return to Roster"></form>
	    </td>
	    
	   </tr>
	   
	</TABLE>

     </td>
  </tr>
</TABLE>


<%

END SUB

%>









