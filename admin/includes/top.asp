<html>
<head>
<title><%= Title %></title>
<!--#include virtual="/admin/includes/js.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" background = "/templates/images/TopBackground.jpg" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" background="../adminbkgnd.jpg">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="182" bgcolor="#42639F" valign="top"></td>
    <td style="background-image: url('/templates/images/TopSide.jpg')" ><img src="/admin/masthead.gif"></td>
  </tr>
  <tr> 
    <td width="182" bgcolor="#42639F" valign="top"> 
<!--#include virtual="/includes/constants.asp" -->
<%
Function confirmUserName(UserName, DocID, method)
	Dim objConn
	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.Open Application("WaterSkiConn")
	
	If method = "add" then
		SQL = "SELECT UserName FROM Doctors WHERE lower(UserName) = '" & lcase(UserName) & "'"
	Else
		SQL = "SELECT UserName FROM Doctors WHERE lower(UserName) = '" & lcase(UserName) & "' AND DocID <> '" & DocID & "'"
	End If
	
	Dim objRS
	Set objRS = Server.CreateObject("ADODB.RecordSet")
	objRS.ActiveConnection = objConn
	objRS.Open SQL
	
	If objRS.EOF then
		confirmUserName = True
	Else
		confirmUserName = False
	End If
	
	Set objRS = Nothing
	objConn.Close
	Set objConn = Nothing
End Function

Dim objConn
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open Application("WaterSkiConn")
%>
	<% If Session("aauth") then %>
	<font face="Verdana" size="2" COLOR="#FFFFFF">Currently Logged in as: </font><br>
	<font face="Verdana" size="2" COLOR="#FFFFFF"><%= Session("FullName") %></font><br>
	<br>
	<% Else %>
	<font face="Verdana" size="2" COLOR="#FFFFFF">Not currently logged in.</font>
	<% End If %>
	
            <% If Session("aauth") then %>
			<% 
				Dim TopUser
				Set TopUser = Server.CreateObject("ADODB.RecordSet")
				TopUser.ActiveConnection = objConn
				TopUser.Open "SELECT * FROM Users where Name = '" & Session("UserName") & "'"
			%>
			<font face="Verdana" size="2"> 

			<%if TopUser("DownloadMembers1") then%>
            <a href="/admin/createmembershipfile.asp"><FONT face="arial" COLOR="#FFFFFF">Download Membership File</font></a>&nbsp;<br>
			<%end if %>
			
			<%if TopUser("DownloadDBF") then%>
            <a href="/admin/createdbf.asp"><FONT face="arial" COLOR="#FFFFFF">Download member DBF</font></a>&nbsp;<br>
			<%end if %>
			
            <a href="/admin/logout.asp"><font face="arial" COLOR="#FFFFFF"><br>Log Out</font></a>&nbsp;<br>
			</font>
            <% Else %>
			<br>
            <% End If %>
			<br>
			<br>
            <font face="Verdana" size="1">&nbsp;<font COLOR="#FFFFFF">Powered by</font> <a href="http://www.epolk.com"><font COLOR="#FFFFFF">ePolk.com</font></a><br></font>
    </td>
    <td valign="top">



