<!--#include virtual="/admin/includes/security.asp" -->
<!--#include virtual="/epl/functions.asp" -->

<% 
Function GetStateCount(ListValues)
Dim WorkingString
Dim StateCounter 
StateCounter = 1
WorkingString = ListValues
While instr(WorkingString,",") > 0 
	LocationofComma = instr(WorkingString,",")
	'Now trim the string
	WorkingString = right(WorkingString, len(WorkingString) - (LocationofComma + 1))
	StateCounter = StateCounter + 1
wend
GetStateCount = StateCounter
End Function

Function BuildStateSQL(ListValues)
Dim WorkingString
Dim StateCounter 
Dim StateSQL 
StateCounter = 1
WorkingString = ListValues
'Get the first state
LocationofComma = instr(WorkingString,",")
if LocationofComma > 0 then
	StateSQL = "State = '" & left(WorkingString, (LocationofComma - 1)) & "'"
else
	StateSQL = "State = '" & WorkingString & "'"
end if

While instr(WorkingString,",") > 0 
	LocationofComma = instr(WorkingString,",")
	'Now trim the string
	WorkingString = right(WorkingString, len(WorkingString) - (LocationofComma + 1))
	StateCounter = StateCounter + 1
	StateSQL = StateSQL & " OR State = '" & left(WorkingString, (LocationofComma - 1)) & "'"
wend
BuildStateSQL = StateSQL
End Function

Function BuildStateList(ListValues)
Dim WorkingString
Dim StateCounter 
Dim StateSQL 
StateCounter = 1
WorkingString = ListValues
'Get the first state
LocationofComma = instr(WorkingString,",")
if LocationofComma > 0 then
	StateSQL = left(WorkingString, (LocationofComma - 1))
else
	StateSQL = WorkingString
end if

While instr(WorkingString,",") > 0 
	LocationofComma = instr(WorkingString,",")
	'Now trim the string
	WorkingString = right(WorkingString, len(WorkingString) - (LocationofComma + 1))
	StateCounter = StateCounter + 1
	StateSQL = StateSQL & left(WorkingString, (LocationofComma - 1))
wend
BuildStateList = StateSQL
End Function

If Request.Form <> "" then 	'this is a postback

'Count the number of states
	Dim StateCount
	if len(request.form("States")) > 0 then
		'Get a state count
		StateCount = GetStateCount(request.form("States"))
		if StateCount > 5 then
			ErroronPage = True
			ErrorMessage = "You can only select up to 5 states."
		else 
			'assign this to a session var to be used on the next page
			Session("StateSQL") = BuildStateSQL(request.form("States"))
			Session("StateList") = BuildStateList(request.form("States"))
			if Request.Form("IncludeForeign") = "True" then
				Session("StateSQL") = "(" & Session("StateSQL") & ") or StateCode is Null"
				Session("StateList") = Session("StateList") & "FN"
			end if
		end if
	else
		ErroronPage = True
		ErrorMessage = "Please select at least one state"
	end if
	
	if len(request.form("tournamentdate")) > 0 then
		'Validate the date and insert into the temp table
		if not epl_isValidDate(request.form("tournamentdate")) then
			ErroronPage = True
			ErrorMessage = "Invalid Tournament Date"			
		else
			Session("tournamentdate") = request.form("tournamentdate")
		end if
	else
		ErroronPage = True
		ErrorMessage = "Please enter a tournament date."
	end if
	

	if len(request.form("TournamentYear")) > 0 then
		Session("TournamentYear") = request.form("TournamentYear")
	else
		ErroronPage = True
		ErrorMessage = "Please enter a tournament year."
	end if
		
	if ErroronPage = True then
		'show page with the error
	else
		if request.form("FileFormat") = "without_scores" then
			response.redirect "createRegTemplatewithoutscores.asp"
		else
			response.redirect "createRegTemplate.asp"
		end if
	end if
	
	
end if
%>
<html>

<head>
<title>Admin Index</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" background = "/templates/images/TopBackground.jpg" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" >
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="27%" valign="top"  bgcolor="#42639F"></td>
    <td width="73%"  bgcolor="#42639F">
      <p>&nbsp;</p>
      <p align="center"><font face="Verdana" size="6" color="#FFFFFF">USA Water 
        Ski Admin</font></p>
      <p>&nbsp;</p>
    </td>
  </tr>
  
</table>
<table width="800" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="177" valign="top"  bgcolor="#42639F">
<%  	Dim objConn
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open Application("WaterSkiConn")
%>
	<% If Session("aauth") then %>
	<font face="Verdana" size="2" COLOR="#FFFFFF">Currently Logged in as: </font><br>
	<font face="Verdana" size="2" COLOR="#FFFFFF"><%= Session("UserName") %></font><br>
	<br>
	<% Else %>
	<font face="Verdana" size="2" COLOR="#FFFFFF">Not currently logged in.</font>
	<% End If %>
	
            <% If Session("aauth") then 
	
				Dim TopUser
				Set TopUser = Server.CreateObject("ADODB.RecordSet")
				TopUser.ActiveConnection = objConn
				TopUser.Open "SELECT * FROM Users999 where Name = '" & Session("UserName") & "'"
			%>
			<font face="Verdana" size="2"> 

			
            <a href="/admin/logout.asp"><font face="arial" COLOR="#FFFFFF"><br>Log Out</font></a>&nbsp;<br>
			</font>
            <% Else %>
			<br>
            <% End If %>
			<br>
			<br>
            <font face="Verdana" size="1">&nbsp;<font COLOR="#FFFFFF">Powered by</font> <a href="http://www.epolk.com"><font COLOR="#FFFFFF">ePolk.com</font></a><br></font>

	
	</td>
    <td width="600" >
<form action="CreateRegTemplateStep1.asp" method="post">
        <table width="600" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="14">&nbsp;</td>
            <td width="233">
			<% if len(ErrorMessage) > 0 then %>
			<font color="#FF0000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><%= ErrorMessage %></strong></font>
			<%  end if %>
			</td>
            <td width="323">&nbsp;</td>
            <td width="14">&nbsp;</td>
            <td width="16">&nbsp;</td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Select 
              the participant's State(s)</font></td>
            <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Tournament 
              Date<font size="1"> (mm/dd/yyyy)</font></font></td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td> <select name="States" size="10" multiple id="States">
                <%   
Dim objRS
Set objConn1 = Server.CreateObject("ADODB.Connection")
objConn1.Open Application("WaterSkiConn")
Set objRS = Server.CreateObject("ADODB.RecordSet")
objRS.ActiveConnection = objConn1

objRS.Open "SELECT * FROM [USStates] Order by StateName;" 

Do until objRS.EOF %>
                <option value="<% Response.write objrs("StateCode") %>" > 
                <% Response.write objrs("StateName") %>
                </option>
                <%	objRS.MoveNext
Loop



objRS.Close
Set objRS = Nothing
objConn1.Close
Set objConn1 = Nothing

%>
              </select></td>
            <td valign="top"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td><input name="tournamentdate" type="text" id="tournamentdate" value="<%= Session("TournamentDate") %>"></td>
                </tr>
                <tr>
                  <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Tournament 
                    Year<font size="1"> </font></font></td>
                </tr>
                <tr>
                  <td><select name="TournamentYear" id="TournamentYear">
				  <% 
				  Dim CurrentYear
				  CurrentYear = datepart("yyyy",date())
				  if len(session("TournamentName")) > 0 and epl_IsNumeric(left(Session("UserName"),2)) = true then
						TournamentYear = (2000 + left(Session("UserName"),2))
				  else
						TournamentYear = datepart("yyyy",date())
				  end if
				  %>
					  <option value="<%= CurrentYear - 1 %>"<% if TournamentYear = (CurrentYear - 1) then response.write "Selected" %>><%= CurrentYear - 1 %></option>
                      <option value="<%= CurrentYear %>"<% if TournamentYear = (CurrentYear) then response.write "Selected" %>><%= CurrentYear %></option>
                      <option value="<%= CurrentYear + 1 %>"<% if TournamentYear = (CurrentYear + 1) then response.write "Selected" %>><%= CurrentYear + 1 %></option>
					  <option value="<%= CurrentYear + 2 %>"<% if TournamentYear = (CurrentYear + 2) then response.write "Selected" %>><%= CurrentYear + 2 %></option>
                    </select></td>
                </tr>
				<tr>
                  <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;
                    </font></td>
                </tr>
				<tr>
                  <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
                    <input type="checkbox" name="IncludeForeign" value="True">
Include
                  Foreign Skiers also<font size="1"> </font> </font></td>
                </tr>

				<tr>

                  <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><br>
                  File format version</font> </td>
                </tr>
				<tr>
                  <td>
				  <input name="FileFormat" type="radio" value="with_scores" checked>
				  <font size="2" face="Arial, Helvetica, sans-serif">with scores</font><font size="1" face="Arial, Helvetica, sans-serif"> (needed
				  for importing into WSTIMS)</font><br>
				  <input name="FileFormat" type="radio" value="without_scores">
				  <font size="2" face="Arial, Helvetica, sans-serif">without scores</font>				  </td>
                </tr>

              </table></td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td><font color="#FF0000" size="2" face="Verdana, Arial, Helvetica, sans-serif">Instructions:</font><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
              Hold down the Control Key and click on each state (up to five) that 
              you would like to retrieve information for and then click the &#8220;Next&#8221; 
              button to create the template.</font></td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td colspan="2"><hr></td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td><input type="submit" name="Submit" value="Next"></td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
        </table>

</form>
    </td>
  </tr>
  
</table>
</body>
</html>






