<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" background="../adminbkgnd.jpg">
<table width="778" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td colspan="2"><img src="/admin/masthead.gif" width="778" height="104"></td>
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
<html>
<head>
<title><%= Title %></title>
<!--#include virtual="/admin/includes/js.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta name="description" content="Fireworks Splice HTML">

</head>
<body bgcolor="#ffffff" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" text="#000033" link="#0000ff" vlink="#0000ff" alink="#0000ff">

	<% If Session("aauth") then %>
	Currently Logged in as: <br>
	<%= Session("FullName") %><br>
	<% Else %>
	Not currently logged in.
	<% End If %>
	
            <% If Session("aauth") then %>
			<% 
				Dim TopUser
				Set TopUser = Server.CreateObject("ADODB.RecordSet")
				TopUser.ActiveConnection = objConn
				TopUser.Open "SELECT * FROM Users where Name = '" & Session("UserName") & "'"
			%>
			<font face="Verdana" size="2"> 
			<%if TopUser("adminJobs") then%>
            <a href="/admin/jobsadmin.asp"><FONT COLOR="#FFFFFF"> Job Listings Administration</font></a>&nbsp;<br>
			<%end if 
			if TopUser("adminSpecialties") then%>
            <a href="/admin/specsadmin.asp"><font COLOR="#FFFFFF">Specialities Administration</font></a>&nbsp;<br>
			<%end if 
			if TopUser("adminDoctors") then%>
            <a href="/admin/doctorsadmin.asp"><font COLOR="#FFFFFF">Doctor List Administration</font></a>&nbsp;<br>
			<%end if 
			if TopUser("adminArticles") then%>
            <a href="/admin/articlesadmin.asp"><font COLOR="#FFFFFF">Articles Administration</font></a>&nbsp;<br>
			<%end if 
			if TopUser("adminNewsletters") then%>
            <a href="/admin/newslettersadmin.asp"><font COLOR="#FFFFFF">Newsletters Administration</font></a>&nbsp;<br>
			<%end if 
			if TopUser("adminUsers") then%>
            <a href="/admin/useradmin.asp"><font COLOR="#FFFFFF">Users Administration</font></a>&nbsp;<br>
			<%end if
			if TopUser("adminDepartments") then%>
            <a href="/admin/departmentsadmin.asp"><font COLOR="#FFFFFF">Departments Administration</font></a>&nbsp;<br>
			<%end if
			if TopUser("adminMedKeyApplications") then%>
            <a href="/admin/adminMedKeyApplication.asp"><font COLOR="#FFFFFF">Med-Key Application Administration</font></a>&nbsp;<br>
			<%end if%>
            <a href="/admin/logout.asp"><font COLOR="#FFFFFF">Log Out</font></a>&nbsp;<br>
			</font>
            <% Else %>
			<br>
            <% End If %>
            <font face="Verdana" size="1">&nbsp;Powered by <a href="http://www.epolk.com"><font COLOR="#FFFFFF">ePolk.com</font></a><br></font>
    </td>
    <td width="596" valign="top">



