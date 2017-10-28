<!--#include virtual="epl/functions.asp" -->

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
      <p>&nbsp;</p>
    </td>
  </tr>
</table>  
  
<table border="0" cellspacing="0" cellpadding="0">  
  <tr> 
    <td width="185" valign="top" bgcolor="#42639F">

<%  	
Dim objConn
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open Application("WaterSkiConn")
%>
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

	<td width="640" >

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


'	Pull Member Data for selected MemberID

Dim RS, sSQL, ZipHint, eMailHint
Set RS = Server.CreateObject("ADODB.RecordSet")
RS.ActiveConnection = objConn

sSQL = "SELECT Mem.PersonIDWithCheckDigit as MemberID, Mem.LastName, Mem.FirstName, Mem.City," 
sSQL = sSQL + " Mem.State, Mem.Zip as Zipcode, Mem.Address1, Mem.Address2, Mem.Email,"
sSQL = sSQL + " Convert(char(10),Mem.BirthDate,101) as BirthDate,"
sSQL = sSQL + " Mem.MembershipTypeCode as MemType, Typ.TypeCode"
sSQL = sSQL + " FROM USAWaterski.dbo.members as Mem, USAWaterski.dbo.MembershipTypes as Typ"
sSQL = sSQL + " WHERE Mem.MembershipTypeCode = Typ.MemberShipTypeID"
sSQL = sSQL + " AND Mem.PersonIDWithCheckDigit='" & Request("MemberID") & "';" 

RS.open sSQL

	%>

<br>
<form action="ValidateMember.asp" method="post">


<TABLE BORDER="4" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#DDDDFF" width=90% >
	
	<font size="2" face="Verdana, Arial, Helvetica, sans-serif"><p>To validate that you
	indeed <b><i>are</i></b> the member you selected, you will need to supply the following
	details.&nbsp; What you enter here will be checked against your membership record, 
	to ensure security.&nbsp; For each of the items below, please enter the requested 
	information, then click the "Validate" button at the bottom of the form.</p>

	<% iat = instr(trim(RS("email")),"@")
	IF iat > 0 THEN %>

		<p>I am showing you the @server.typ part of the E-Mail address that we have in 
		your membership record, as a reminder.&nbsp; Please supply the ID part of that
		E-Mail address as we have it.&nbsp; Don't worry if your E-Mail address has changed
		since, you will be given an opportunity to update your E-Mail address and other 
		details, once we've got you signed into the Members-only area.</p>
		
	<% END IF %>
				 
	</font>
	
  <tr><br>
      <td BGCOLOR="#42639F">
        <center><font face="Verdana, Arial, Helvetica, sans-serif" size="4" COLOR="#FFFFFF"><b>Enter the Following Information:</b></font>
        <br>
      </td>
  </tr>  

  <tr>
     <td>
	<br>
			<TABLE BORDER="1" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#DDDDFF" width=95% >
          <TR>
            <TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Item Description</FONT></Center></TD>
            <TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Hint or Format</FONT></Center></TD>
            <TD ALIGN="Left" vAlign="top"><Center><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Enter Value Here</FONT></Center></TD>
          </TR>

				<% IF RS("BirthDate") <> "01/01/1900" THEN %>

          <TR>
            <TD ALIGN="Left" vAlign="center"><Center><FONT COlOR="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Birth Date</FONT></Center></TD>
            <TD ALIGN="Left" vAlign="center"><Center><FONT COlOR="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">mm/dd/yyyy</FONT></Center></TD>
            <TD ALIGN="Left" vAlign="center"><Center><FONT COlOR="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><input type="text" name="BirthDate" value="<%=RS("Birthdate")%>" maxlength=10 size=13></input></FONT></Center></TD>
          </TR>

				<% END IF %>

				<% IF trim(RS("Zipcode")) <> "" THEN 

							ZipHint = "( " & FormatNumber(len(trim(RS("Zipcode"))),0) & " characters )" %>

          <TR>
            <TD ALIGN="Left" vAlign="center"><Center><FONT COlOR="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Zip Code</FONT></Center></TD>
            <TD ALIGN="Left" vAlign="center"><Center><FONT COlOR="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=ZipHint%></FONT></Center></TD>
            <TD ALIGN="Left" vAlign="center"><Center><FONT COlOR="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><input type="text" name="Zipcode" value="<%=RS("Zipcode")%>" maxlength=10 size=13></input></FONT></Center></TD>
          </TR>

				<% END IF %>


				<% IF iat > 0 THEN 

							eMailHint = "( " & mid(trim(RS("Email")),iat) & " )" %>

          <TR>
            <TD ALIGN="Left" vAlign="center"><Center><FONT COlOR="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">E-Mail ID</FONT></Center></TD>
            <TD ALIGN="Left" vAlign="center"><Center><FONT COlOR="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=eMailHint%></FONT></Center></TD>
            <TD ALIGN="Left" vAlign="center"><Center><FONT COlOR="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><input type="text" name="EMailID" value="<%=left(trim(RS("EMail")),iat-1)%>" maxlength=30 size=33></input></FONT></Center></TD>
          </TR>

				<% END IF %>


        </table> 

	<TABLE BORDER="0" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="<%=TableColor1%>" width=60% >  
	  <tr>	
	    <td>
	      <br>
	      
	          <input type="hidden" name="MemberID" value="<%=Request("MemberID")%>">
	      
	          <center><input type="submit" value="Validate">

	    </td>
	   </tr>
	</TABLE>

     </td>
  </tr>
</TABLE>
</form>


	</td>
  </tr>
</table>
</body>
</html>





