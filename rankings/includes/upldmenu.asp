	
<html>

<head>
<title>USA Water Ski Post-Tournament Uploads</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" >

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#42639F">
      <p>&nbsp;</p>
      <p align="center"><font face="Verdana" size="6" color="#FFFFFF">
      	<%=TopHead%>
      	</font></p>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  
  <tr> 
    <td width="185" bgcolor="#42639F" valign="top">

	
	<% If Session("aauth") then %>
	<font face="Verdana" size="2" COLOR="#FFFFFF"><br>&nbsp;Currently Logged in as: </font><br>&nbsp;<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
	<font face="Verdana" size="2" COLOR="#FFFFFF">&nbsp;<%=Session("UserName")%>&nbsp;&nbsp;
		<br>&nbsp;<br>&nbsp;<br>
	
	    &nbsp;<a href="UpldLogout.asp"><font face="arial" COLOR="#FFFFFF">Log Out</font></a><br>&nbsp;<br>

		</font>
	
	<% Else %>
	<font face="Verdana" size="2" COLOR="#FFFFFF"><br>&nbsp;Not currently logged in.<br>&nbsp;<br></font>
	<% End If %>
	
	        &nbsp;<a href=UpldIndx.asp><font face="arial" size="2" COLOR="#FFFFFF">Back to Admin Index</font></a><br>&nbsp;<br>

	        &nbsp;<a href="http://usaws.usawaterski.org"><font face="arial" size="2" COLOR="#FFFFFF">Back to Online Sanctioning</font></a><br>&nbsp;<br>

	        &nbsp;<a href="http://www.usawaterski.org"><font face="arial" size="2" COLOR="#FFFFFF">USA Water Ski Home</font></a><br>&nbsp;<br>

			<br>
            <font face="Verdana" size="1">&nbsp;<font COLOR="#FFFFFF">Powered by</font> <a href="http://www.AWSATech.com"><font COLOR="#FFFFFF">AWSATech</font></font></a>
            <br>&nbsp;<br>

    </td>
    <td valign="top" >




