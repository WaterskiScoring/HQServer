<!--#include virtual="epl/functions.asp" -->

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

Set SQLConnect = CreateObject("ADODB.Connection")
SQLConnect.Open Application("HQSQLConn")

'	Pull Member Data direct from HQ table for the selected MemberID

Dim RS, sSQL, ZipHint, eMailHint, NumInput, ErrFlds, iat, jat
Dim sMemberID, sBirthDate, sZipcode, sEMailID
Dim HQBirthDate, HQZipcode, HQEMailID

sMemberID = Request("MemberID")

if sMemberID = 0 or len(sMemberID) = 0 THEN response.redirect "FindMember.asp?FormStatus=newsearch"

sBirthDate = Request("BirthDate")
sZipcode = Request("Zipcode")
sEMailID = Request("EMailID")

sSQL = "SELECT PT.[Last Name] as LastName, PT.[First Name] as FirstName," 
sSQL = sSQL & " PT.EMail, PT.Password, Left(PT.Sex,1) as Sex,"
sSQL = sSQL & " Convert(char(10),PT.[Birth Date],101) as BirthDate,"
sSQL = sSQL & " PA.Address1, PA.Address2, PA.City, PA.State, PA.Zip as Zipcode"
sSQL = sSQL & " FROM Waterski.dbo.tblPeople as PT, Waterski.dbo.tblPeopleAddresses as PA"
sSQL = sSQL & " WHERE PT.[Person ID] = " & right(sMemberID,8) 
sSQL = sSQL & " AND PA.[Person ID] = " & right(sMemberID,8)
sSQL = sSQL & " AND PA.[Primary] = 1"

Set RS = SQLConnect.Execute(sSql)

HQBirthDate = RS("BirthDate")
HQZipcode = RS("Zipcode")
HQEMailID = RS("Email")

sMembAge = Datepart("yyyy",now) - right(HQBirthDate,4)

IF sBirthDate <> "" THEN
	IF mid(sBirthDate,2,1) = "/" THEN sBirthDate = "0" & sBirthDate
	IF mid(sBirthDate,5,1) = "/" THEN
		sBirthDate = left(sBirthDate,3) & "0" & Mid(sBirthDate,4)
	END IF
	IF Len(sBirthDate) = 8 THEN
		sBirthDate = left(sBirthDate,6) & "19" & right(sBirthDate,2)
	END IF		
END IF

IF mid(HQZipcode,6,1) = "-" AND (HQBirthDate <> "01/01/1900" OR instr(HQEMailID,"@")) THEN HQZipcode = left(HQZipcode,5)

%>


<html>

<head>
<title>Validate Member</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" >

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#42639F">
      <p>&nbsp;</p>
      <p align="center"><font face="Verdana" size="6" color="#FFFFFF">
      	USA Water Ski Member Validation</font></p>
      <p align="center"><font face="Verdana" size="3" color="#FFFFFF"><b>
      	Validation for:&nbsp;&nbsp; <%=RS("FirstName")%>&nbsp;<%=RS("LastName")%>,&nbsp;&nbsp;
      	<%=RS("City")%>, <%=RS("State")%>&nbsp;&nbsp; ( <%=sMembAge%> / <%=RS("Sex")%> )</font></b></p>

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

<td width="640" >

<br>

<% 

' We're now set up for the rest of the page / table

' Now Examine Supplied values (if any), and for each supplied value, 
' see if that values matches what's in the Membership Record.

NumInput = 0: ErrFlds = "": 
iat = instr(trim(HQEMailID),"@")
jat = instr(trim(sEMailID),"@")

IF len(sBirthDate) > 0 THEN NumInput = NumInput + 1
IF len(sZipcode) > 0 THEN NumInput = NumInput + 1
IF len(sEMailID) > 0 THEN NumInput = NumInput + 1
	
IF NumInput > 0 THEN

	IF len(HQBirthDate) > 0 THEN
		IF trim(sBirthDate) <> trim(HQBirthDate) THEN
			ErrFlds = "BirthDate"
		END IF
	END IF

	IF len(HQZipcode) > 0 THEN
		IF lcase(trim(sZipcode)) <> lcase(trim(HQZipcode)) THEN
			IF left(lcase(trim(sZipcode)),5) <> left(lcase(trim(HQZipcode)),5) or NumInput = 1 THEN
				IF len(ErrFlds) > 0 THEN ErrFlds = ErrFlds & " and "
				ErrFlds = ErrFlds & "Zipcode"
			END IF
		END IF
	END IF

	IF jat = 0 THEN jat = len(trim(sEMailID)) ELSE jat = jat - 1
	IF iat > 0 THEN 
		IF left(lcase(trim(sEMailID)),jat) <> left(lcase(trim(HQEMailID)),iat-1) THEN
			IF len(ErrFlds) > 0 THEN ErrFlds = ErrFlds & " and "
			ErrFlds = ErrFlds & "E-Mail ID"
		END IF
	END IF

END IF

' Now Examine analyzed inputs and validation values and construct an appropriate display


' Start -- if we have inputs and no errors, then we've validated 
' in which case we're good to go -- paint confirmation screen.

	IF len(ErrFlds) = 0 and NumInput > 0 THEN %>

		<form action="http://www.usawaterski.org/members/login/index.asp?m=<%=sMemberID%>&p=<%=RS("Password")%>&c=987" method="post">

		<TABLE BORDER="4" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#DDDDFF" width=80% >

		<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><p><b>Validation Successful.</b></p>
		
		<p>We are now prepared to take you back to the Members Sign-In screen, where I will
		automatically supply your Member Number and Password.&nbsp; To help you sign in
		more directly on future visits, please make a note of these two key items:</p>

		<p><b>Member#:&nbsp; <font color="red"><%=sMemberID%></font></p>
		<p>Password:&nbsp; <font color="red"><%=RS("Password")%></font></p></font>

	          <center><input type="submit" value="Take me to Members Sign In"></form>

<% ELSE %>

	<form action="ValidateMember.asp" method="post">

	<TABLE BORDER="4" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#DDDDFF" width=90% >

	<% IF len(ErrFlds) > 0 THEN %>

		<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><p><b>Validation failed for:&nbsp;
		<font color="red"><%=ErrFlds%></font></b></p>
		
		<% Session("ValidationAttempts") = 1 + Session("ValidationAttempts")
		IF Session("ValidationAttempts") > 3 THEN %>
			
			<p>Validation attempt limit exceeded.&nbsp; Sorry.</p></font></form>

		<% ELSE 
			
			IF len(sZipCode) = 5 and len(HQZipcode) = 10 and instr(ErrFlds,"Zipcode") > 0 THEN %>

				<p>Please note that you need all 10 digits of your Zipcode, including the hypen
				(ie like 12345-6789) -- if you don't have this on the tip of your tongue, you
				should be able to find it on any piece of incoming mail.</p>

			<% END IF %>
		
			<p>This will be validation attempt # <%=Session("ValidationAttempts")%>.</p></font>
	
		<% END IF

	ELSE %>
	
	<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><p>To validate that you
	indeed <b><i>are</i></b> the member you selected, you will need to supply the following
	details.&nbsp; What you enter here will be checked against your membership record, 
	to ensure security.&nbsp; For each of the items below, please enter the requested 
	information, then click the "Validate" button at the bottom of the form.</p>
		
	<% IF iat > 0 THEN %>

		<p>I am showing you the @server.typ part of the E-Mail address that we have in 
		your membership record, as a reminder.&nbsp; Please supply only the ID part of that
		E-Mail address -- just the characters to the left of the @ sign -- as we have it on
		file currently.&nbsp; Don't worry if your E-Mail address has changed	since, you 
		will be given an opportunity to update your E-Mail address and other details, 
		once we've got you signed into the Members-only area.</p></font>
		
	<% END IF
	
	END IF

IF Session("ValidationAttempts") <= 3 THEN %>

  <tr><br>
      <td BGCOLOR="#42639F">
        <center><font face="Verdana, Arial, Helvetica, sans-serif" size="4" COLOR="#FFFFFF"><b>Enter the Following Information:</b></font>
        <br>
      </td>
  </tr>  

  <tr>
    <td>&nbsp; <br>
			<TABLE BORDER="1" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#DDDDFF" width=95% >
          <TR>
            <TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Item Description</FONT></Center></TD>
            <TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Hint or Format</FONT></Center></TD>
            <TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Enter Value Here</FONT></Center></TD>
          </TR>

				<% IF HQBirthDate <> "01/01/1900" THEN %>

          <TR>
            <TD ALIGN="Left" vAlign="center"><Center><FONT COlOR="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Birth Date</FONT></Center></TD>
            <TD ALIGN="Left" vAlign="center"><Center><FONT COlOR="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">mm/dd/yyyy</FONT></Center></TD>
            <TD ALIGN="Left" vAlign="center"><Center><FONT COlOR="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><input type="text" name="BirthDate" value="<%=sBirthDate%>" maxlength=10 size=12></input></FONT></Center></TD>
          </TR>

				<% END IF %>

				<% IF trim(HQZipcode) <> "" THEN 

							ZipHint = "( " & len(trim(HQZipcode)) & " characters )" %>

          <TR>
            <TD ALIGN="Left" vAlign="center"><Center><FONT COlOR="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Zip Code</FONT></Center></TD>
            <TD ALIGN="Left" vAlign="center"><Center><FONT COlOR="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=ZipHint%></FONT></Center></TD>
            <TD ALIGN="Left" vAlign="center"><Center><FONT COlOR="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><input type="text" name="Zipcode" value="<%=sZipcode%>" maxlength=20 size=12></input></FONT></Center></TD>
          </TR>

				<% END IF %>


				<% IF iat > 0 THEN 

							eMailHint = "( ???" & mid(trim(HQEMailID),iat) & " )" %>

          <TR>
            <TD ALIGN="Left" vAlign="center"><Center><FONT COlOR="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">E-Mail ID</FONT></Center></TD>
            <TD ALIGN="Left" vAlign="center"><Center><FONT COlOR="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=eMailHint%></FONT></Center></TD>
            <TD ALIGN="Left" vAlign="center"><Center><FONT COlOR="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><input type="text" name="EMailID" value="<%=left(trim(sEMailID),jat)%>" maxlength=50 size=25></input></FONT></Center></TD>
          </TR>

				<% END IF %>


      </table> 

	<TABLE BORDER="0" VALIGN="CENTER" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="<%=TableColor1%>" width=80% >  
	  <tr>	

	    <td width=35% align=center valign=center>
	    	<input type="hidden" name="MemberID" value="<%=sMemberID%>">
	      <center><input type="submit" style="width:7em" value="Validate">
	    </td>
	    
	  </form>

	    <TD width=55% align=center valign=center><br>
				<form action="FindMember.asp?FormStatus=newsearch" method="post">
				<center><input type="submit" style="width:13em" value="Not me, New Search"></form>
		 </TD>
 
	  </tr>
	</TABLE>


	<% END IF

	END IF %>

     </td>
  </tr>
</TABLE>


	</td>
  </tr>
</table>
</body>
</html>





