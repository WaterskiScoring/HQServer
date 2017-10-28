<!--#include virtual="/epl/functions.asp" -->

<% If not Session("aauth") then response.redirect "Login.asp"
	
	Server.ScriptTimeout = 300

	' The following lines of HTML display the "opening please wait" banner.

	 %>

<html>

<head>
<title>USA Water Ski Duplicate Member Problems v1.0</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" >

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#42639F">
      <p>&nbsp;</p>
      <p align="center"><font face="Verdana" size="6" color="#FFFFFF">
      	USA Water Ski Possible Duplicate Members</font></p>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>
  
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="185" bgcolor="#42639F" valign="top">

  <!--#include virtual="/admin/includes/menu.asp" -->

    </td>

		<td>

<%

Dim DTStart, DTEnd, DateRaw, DateFmt, I1, I2

IF len(request("DTEnd")) = 0 THEN
   DateRaw = DateAdd("d",-7,Date())
   I1 = instr(DateRaw,"/"): I2 = instr(I1+1,DateRaw,"/")
   DateFmt = Mid(DateRaw,I2+1): ' Start with Year value
   IF I1=2 THEN DateFmt = DateFmt + "-0" + Left(DateRaw,1): ELSE DateFmt = DateFmt + "-" + Left(DateRaw,2)
   IF I2-I1=2 THEN DateFmt = DateFmt + "-0" + Mid(DateRaw,I1+1,1): ELSE DateFmt = DateFmt + "-" + Mid(DateRaw,I1+1,2)
   DTStart = DateFmt
ELSE
   DTEnd = trim(request("DTEnd"))
END IF

IF len(request("DTStart")) = 0 THEN
   DateRaw = DateAdd("d",-1,Date())
   I1 = instr(DateRaw,"/"): I2 = instr(I1+1,DateRaw,"/")
   DateFmt = Mid(DateRaw,I2+1): ' Start with Year value
   IF I1=2 THEN DateFmt = DateFmt + "-0" + Left(DateRaw,1): ELSE DateFmt = DateFmt + "-" + Left(DateRaw,2)
   IF I2-I1=2 THEN DateFmt = DateFmt + "-0" + Mid(DateRaw,I1+1,1): ELSE DateFmt = DateFmt + "-" + Mid(DateRaw,I1+1,2)
   DTEnd = DateFmt
ELSE
   DTStart = trim(request("DTStart"))
END IF

%> <table width="65%" align="Center" Cellpadding="0" Cellspacing="5">

	<tr><td>&nbsp;</td></tr>
	
	<tr align="Center">   
		<td><FONT size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
			&nbsp;&nbsp;&nbsp; Effective Date &nbsp;&nbsp;&nbsp;<br>
			&nbsp;&nbsp;&nbsp; Range Start &nbsp;&nbsp;&nbsp; </font></td> 
		<td><FONT size="1" face="Verdana, Arial, Helvetica, sans-serif">
			&nbsp;&nbsp;&nbsp; Effective Date &nbsp;&nbsp;&nbsp; <br> 
			&nbsp;&nbsp;&nbsp; Range End &nbsp;&nbsp;&nbsp; </font></td> 
		<td><FONT size="1" face="Verdana, Arial, Helvetica, sans-serif">
			&nbsp;&nbsp;&nbsp; Specify Dates&nbsp;&nbsp;&nbsp; <br>
			&nbsp;&nbsp;&nbsp; (as YYYY-MM-DD) &nbsp;&nbsp;&nbsp; </font></td>
	</tr>
		
	<tr align="Center"><form action="/admin/FindPossDupMembers.asp" method="post">  
		<td><input name="DTStart" type="text" size=10 id="DTStart" value="<%= Response.Write(DTStart) %>"></td>
		<td><input name="DTEnd" type="text" size=10 id="DTEnd" value="<%= Response.Write(DTEnd) %>"></td>
		<td><input type="submit" value="Find Poss Dups"></td></form>
	</tr>
	
	<tr><td>&nbsp;</td></tr>
	
</table>

<%

Dim rs
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.ActiveConnection = objConn

sSQL = "Select * From ("

sSQL = sSQL & " Select OldID, OldEffFrom, OldEffTo, OldMemType, OldDiv1, OldDiv2, OldBDate, OldLast, OldFirst, OldAdrs, OldCity, OldState, OldZip,"
sSQL = sSQL & " 'LB' as MatType, NewID, NewEffFrom, NewEffTo, NewMemType, NewDiv1, NewDiv2, NewBDate, NewLast, NewFirst, NewAdrs, NewCity, NewState, NewZip"
sSQL = sSQL & " FROM (Select PersonID as OldID, LastName as OldLast, FirstName as OldFirst,"
sSQL = sSQL & " DivisionCode1 as OldDiv1, DivisionCode2 as OldDiv2, MembershipType as OldMemType,"
sSQL = sSQL & " Address1 as OldAdrs, Left(City,13) as OldCity, State as OldState, left(Zip,5) as OldZip,"
sSQL = sSQL & " Convert(char(10),BirthDate,101) as OldBDate,"
sSQL = sSQL & " Convert(char(10),EffectiveFrom,101) as OldEffFrom,"
sSQL = sSQL & " Convert(char(10),EffectiveTo,101) as OldEffTo"
sSQL = sSQL & " From USAWaterski.dbo.Members"
sSQL = sSQL & " Where EffectiveTo >= '" & DtEnd & "'"
sSQL = sSQL & " and left(LastName,1) > ' ' and left(FirstName,1) > ' '"
sSQL = sSQL & " and MembershipType not in ('CA','CS','SO','IB','IS','IG','ISP','AAL','WSC','WSM')) as Old,"
sSQL = sSQL & " (Select PersonID as NewID, LastName as NewLast, FirstName as NewFirst,"
sSQL = sSQL & " DivisionCode1 as NewDiv1, DivisionCode2 as NewDiv2, MembershipType as NewMemType,"
sSQL = sSQL & " Address1 as NewAdrs, Left(City,13) as NewCity, State as NewState, left(Zip,5) as NewZip,"
sSQL = sSQL & " Convert(char(10),BirthDate,101) as NewBDate,"
sSQL = sSQL & " Convert(char(10),EffectiveFrom,101) as NewEffFrom,"
sSQL = sSQL & " Convert(char(10),EffectiveTo,101) as NewEffTo"
sSQL = sSQL & " From USAWaterski.dbo.Members"
sSQL = sSQL & " Where EffectiveFrom between '" & DtStart & "' and '" & DtEnd & "'"
sSQL = sSQL & " and left(LastName,1) > ' ' and left(FirstName,1) > ' '"
sSQL = sSQL & " and MembershipType not in ('CA','CS','SO','IB','IS','IG','ISP','AAL','WSC','WSM')) as New"
sSQL = sSQL & " WHERE OldID < NewID and OldLast = NewLast"
sSQL = sSQL & " and OldBDate = NewBDate and NewBDate <> '01/01/1900'"

sSQL = sSQL & " UNION Select OldID, OldEffFrom, OldEffTo, OldMemType, OldDiv1, OldDiv2, OldBDate, OldLast, OldFirst, OldAdrs, OldCity, OldState, OldZip,"
sSQL = sSQL & " 'FBS' as MatType, NewID, NewEffFrom, NewEffTo, NewMemType, NewDiv1, NewDiv2, NewBDate, NewLast, NewFirst, NewAdrs, NewCity, NewState, NewZip"
sSQL = sSQL & " FROM (Select PersonID as OldID, LastName as OldLast, FirstName as OldFirst,"
sSQL = sSQL & " DivisionCode1 as OldDiv1, DivisionCode2 as OldDiv2, MembershipType as OldMemType,"
sSQL = sSQL & " Address1 as OldAdrs, Left(City,13) as OldCity, State as OldState, left(Zip,5) as OldZip,"
sSQL = sSQL & " Convert(char(10),BirthDate,101) as OldBDate,"
sSQL = sSQL & " Convert(char(10),EffectiveFrom,101) as OldEffFrom,"
sSQL = sSQL & " Convert(char(10),EffectiveTo,101) as OldEffTo"
sSQL = sSQL & " From USAWaterski.dbo.Members"
sSQL = sSQL & " Where EffectiveTo >= '" & DtEnd & "'"
sSQL = sSQL & " and left(LastName,1) > ' ' and left(FirstName,1) > ' '"
sSQL = sSQL & " and MembershipType not in ('CA','CS','SO','IB','IS','IG','ISP','AAL','WSC','WSM')) as Old,"
sSQL = sSQL & " (Select PersonID as NewID, LastName as NewLast, FirstName as NewFirst,"
sSQL = sSQL & " DivisionCode1 as NewDiv1, DivisionCode2 as NewDiv2, MembershipType as NewMemType,"
sSQL = sSQL & " Address1 as NewAdrs, Left(City,13) as NewCity, State as NewState, left(Zip,5) as NewZip,"
sSQL = sSQL & " Convert(char(10),BirthDate,101) as NewBDate,"
sSQL = sSQL & " Convert(char(10),EffectiveFrom,101) as NewEffFrom,"
sSQL = sSQL & " Convert(char(10),EffectiveTo,101) as NewEffTo"
sSQL = sSQL & " From USAWaterski.dbo.Members"
sSQL = sSQL & " Where EffectiveFrom between '" & DtStart & "' and '" & DtEnd & "'"
sSQL = sSQL & " and left(LastName,1) > ' ' and left(FirstName,1) > ' '"
sSQL = sSQL & " and MembershipType not in ('CA','CS','SO','IB','IS','IG','ISP','AAL','WSC','WSM')) as New"
sSQL = sSQL & " WHERE OldID < NewID and OldLast <> NewLast"
sSQL = sSQL & " and OldBDate = NewBDate and NewBDate <> '01/01/1900'"
sSQL = sSQL & " and OldFirst = NewFirst and OldState = NewState"

sSQL = sSQL & " UNION Select OldID, OldEffFrom, OldEffTo, OldMemType, OldDiv1, OldDiv2, OldBDate, OldLast, OldFirst, OldAdrs, OldCity, OldState, OldZip,"
sSQL = sSQL & " 'FLZ' as MatType, NewID, NewEffFrom, NewEffTo, NewMemType, NewDiv1, NewDiv2, NewBDate, NewLast, NewFirst, NewAdrs, NewCity, NewState, NewZip"
sSQL = sSQL & " FROM (Select PersonID as OldID, LastName as OldLast, FirstName as OldFirst,"
sSQL = sSQL & " DivisionCode1 as OldDiv1, DivisionCode2 as OldDiv2, MembershipType as OldMemType,"
sSQL = sSQL & " Address1 as OldAdrs, Left(City,13) as OldCity, State as OldState, left(Zip,5) as OldZip,"
sSQL = sSQL & " Convert(char(10),BirthDate,101) as OldBDate,"
sSQL = sSQL & " Convert(char(10),EffectiveFrom,101) as OldEffFrom,"
sSQL = sSQL & " Convert(char(10),EffectiveTo,101) as OldEffTo"
sSQL = sSQL & " From USAWaterski.dbo.Members"
sSQL = sSQL & " Where EffectiveTo >= '" & DtEnd & "'"
sSQL = sSQL & " and left(LastName,1) > ' ' and left(FirstName,1) > ' '"
sSQL = sSQL & " and MembershipType not in ('CA','CS','SO','IB','IS','IG','ISP','AAL','WSC','WSM')) as Old,"
sSQL = sSQL & " (Select PersonID as NewID, LastName as NewLast, FirstName as NewFirst,"
sSQL = sSQL & " DivisionCode1 as NewDiv1, DivisionCode2 as NewDiv2, MembershipType as NewMemType,"
sSQL = sSQL & " Address1 as NewAdrs, Left(City,13) as NewCity, State as NewState, left(Zip,5) as NewZip,"
sSQL = sSQL & " Convert(char(10),BirthDate,101) as NewBDate,"
sSQL = sSQL & " Convert(char(10),EffectiveFrom,101) as NewEffFrom,"
sSQL = sSQL & " Convert(char(10),EffectiveTo,101) as NewEffTo"
sSQL = sSQL & " From USAWaterski.dbo.Members"
sSQL = sSQL & " Where EffectiveFrom between '" & DtStart & "' and '" & DtEnd & "'"
sSQL = sSQL & " and left(LastName,1) > ' ' and left(FirstName,1) > ' '"
sSQL = sSQL & " and MembershipType not in ('CA','CS','SO','IB','IS','IG','ISP','AAL','WSC','WSM')) as New"
sSQL = sSQL & " WHERE OldID < NewID and OldLast = NewLast and OldZip = NewZip"
sSQL = sSQL & " and OldFirst = NewFirst and OldBDate <> NewBDate"
sSQL = sSQL & " and (right(OldBDate,4) = '1900' or right(NewBDate,4) = '1900' or"
sSQL = sSQL & " abs(cast(right(OldBDate,4) as smallint)-cast(right(NewBDate,4) as smallint)) < 15)"

sSQL = sSQL & " ) US Order by	NewEffFrom, OldLast, OldFirst"

rs.open sSQL

IF NOT rs.eof THEN

	rs.movefirst
	
	%>

	<table align="Center" BORDER="0">

		<tr>
		  <th align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="black"><u>##</u><br>Typ</font></th>
		  <th align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="black">Person<br>ID No.</font></th>
		  <th align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="black">Eff Fm<br>Date</font></th>
		  <th align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="black">Eff To<br>Date</font></th>
		  <th align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="black">Mem<br>Type</font></th>
		  <th align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="black">Div<br>#1</font></th>
		  <th align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="black">Div<br>#2</font></th>
		  <th align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="black">Birth<br>Date</font></th>
		  <th align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="black">Last<br>Name</font></th>
		  <th align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="black">First<br>Name</font></th>
		  <th align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="black">Street<br>Address</font></th>
		  <th align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="black">City<br>Name</font></th>
		  <th align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="black">State<br>Code</font></th>
		  <th align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="black">Zip<br>Code</font></th>
		</tr>

		<%
	  DO WHILE NOT rs.eof
		rowCount = rowCount + 1
		%>

		<tr>			
			<td colspan="14"><hr></td>
		</tr>
		
		<tr>			
		  <td align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="blue"><%=rowcount%></font></td>
		  <td align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="blue"><%=rs("OldID")%></font></td>
		  <td align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="blue"><%=rs("OldEffFrom")%></font></td>
		  <td align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="blue"><%=rs("OldEffTo")%></font></td>
		  <td align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="blue"><%=rs("OldMemType")%></font></td>
		  <td align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="blue"><%=rs("OldDiv1")%></font></td>
		  <td align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="blue"><%=rs("OldDiv2")%></font></td>
		  <td align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="blue"><%=rs("OldBDate")%></font></td>
		  <td align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="blue"><%=rs("OldLast")%></font></td>
		  <td align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="blue"><%=rs("OldFirst")%></font></td>
		  <td align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="blue"><%=rs("OldAdrs")%></font></td>
		  <td align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="blue"><%=rs("OldCity")%></font></td>
		  <td align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="blue"><%=rs("OldState")%></font></td>
		  <td align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="blue"><%=rs("OldZip")%></font></td>
		</tr>
		
		<tr>			
		  <td align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="red"><%=rs("MatType")%></font></td>
		  <td align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="red"><%=rs("NewID")%></font></td>
		  <td align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="red"><%=rs("NewEffFrom")%></font></td>
		  <td align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="red"><%=rs("NewEffTo")%></font></td>
		  <td align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="red"><%=rs("NewMemType")%></font></td>
		  <td align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="red"><%=rs("NewDiv1")%></font></td>
		  <td align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="red"><%=rs("NewDiv2")%></font></td>
		  <td align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="red"><%=rs("NewBDate")%></font></td>
		  <td align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="red"><%=rs("NewLast")%></font></td>
		  <td align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="red"><%=rs("NewFirst")%></font></td>
		  <td align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="red"><%=rs("NewAdrs")%></font></td>
		  <td align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="red"><%=rs("NewCity")%></font></td>
		  <td align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="red"><%=rs("NewState")%></font></td>
		  <td align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="red"><%=rs("NewZip")%></font></td>
		</tr>
		
		<%

		rs.movenext
	  LOOP
	  
	  %>

	</tr>
	</table>

	<% 

	rs.close
	Set rs = nothing

	ELSE
		response.write("No duplicate possibles found for specified dates.")
	END IF

%>	
	
	</td>
  </tr>
</table>
</body>
</html>





