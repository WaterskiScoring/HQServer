
<!--#include virtual="epl/functions.asp" -->

<% 

If not Session("aauth") then response.redirect "Login.asp"
	
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
	StateSQL = "'" & left(WorkingString, (LocationofComma - 1)) & "'"
else
	StateSQL = "'" & WorkingString & "'"
end if

While instr(WorkingString,",") > 0 
	LocationofComma = instr(WorkingString,",")
	'Now trim the string
	WorkingString = right(WorkingString, len(WorkingString) - (LocationofComma + 1))
	StateCounter = StateCounter + 1
	StateSQL = StateSQL & ",'" & left(WorkingString, (LocationofComma - 1)) & "'"
wend

BuildStateSQL = "State IN (" & StateSQL & ")"
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

		end if
	
	end if

	if Request.Form("IncludeForeign") = "True" then
		if len(Session("StateSQL")) > 0 then Session("StateSQL") = Session("StateSQL") & " OR"
		Session("StateSQL") = Session("StateSQL") & " State is Null OR State in ('AA','AB','AE','AP','AS',"
		Session("StateSQL") = Session("StateSQL") & "'BC','FM','GU','MB','MH','MP','NB','NF','NS','  ',"
		Session("StateSQL") = Session("StateSQL") & "'NU','ON','PE','PR','PW','QC','SK','VI','YT')"
		Session("StateList") = Session("StateList") & "FN"
	end if

	if Request.Form("IncludeElite") = "True" then
		if len(Session("StateSQL")) > 0 then Session("StateSQL") = Session("StateSQL") & " OR"
		Session("StateSQL") = Session("StateSQL") & " PersonIDWithCheckDigit in"
		Session("StateSQL") = Session("StateSQL") & " (Select MemberID FROM Cobra00025.usawsrank.Rankings" 
		Session("StateSQL") = Session("StateSQL") & " where left(Div,1) in ('I','O') group by MemberID)"
		Session("StateList") = Session("StateList") & "OP"
	end if


	if len(request.form("tournamentdate")) > 0 then
		'Validate the date and insert into the temp table
		if not epl_isValidDate(request.form("tournamentdate")) then
			ErroronPage = True
			ErrorMessage = "Invalid Tournament Date -- pls revise."			
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
		ErrorMessage = "Please Select a tournament year."
	end if
		

	IF left(request.form("NowWhat"),4) = "Look" and ErroronPage = False then
			response.redirect "LookupMembers.asp"
	END IF


	if ErroronPage = False and len(request.form("States")) = 0 and Request.Form("IncludeForeign") <> "True" and Request.Form("IncludeElite") <> "True" then
		ErroronPage = True
		ErrorMessage = "Please specify at least one State or other Selection Option below."
	end if
	

	if ErroronPage = True then
		'show page with the error
	else

		'Display constructed SQL WHERE clause in debug log
		
		'	Set tempFSO=Server.CreateObject("Scripting.FileSystemObject")
		'	IF Not (tempFSO.FileExists(Server.mappath("/")&"\..\" & "sql-debug-log.txt")) = true THEN
   	'		Set logobject=tempFSO.CreateTextFile(Server.mappath("/")&"\..\" & "sql-debug-log.txt",true)
		'	ELSE
   	'		Set logobject=tempFSO.OpenTextFile(Server.mappath("/")&"\..\" & "sql-debug-log.txt",8,true)
		'	END IF
		'		logobject.WriteLine("SQL = " & Session("StateSQL") & " -+- " & date() & " " & time() & " " & session("UserName"))
		'		logobject.Close
		'	Set logobject=nothing
		'	Set tempFSO=nothing
				
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
<title>Registration Template Controls</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" background = "/images/TopBackground.jpg" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" >

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#42639F">
      <p>&nbsp;</p>
      <p align="center"><font face="Verdana" size="6" color="#FFFFFF">
      	USA Water Ski Registration Templates</font></p>
      <p align="center"><font face="Verdana" size="4" color="#FFFFFF">
      	Registration Support for -- <%=session("TournamentName")%></font></p>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>


<table width="800" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="185" valign="top"  bgcolor="#42639F">
<%  	Dim objConn
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

        <% if len(ErrorMessage) > 0 then %>
    			<font color="#FF0000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><br>&nbsp;
    			&nbsp;&nbsp;&nbsp;&nbsp;<i><%= ErrorMessage %></i></strong></font>
    			<%  end if %>

  <table width="600" border="0" cellspacing="0" cellpadding="0">
      <tr> 
        <td width="20">&nbsp;</td>
        <td width="150">&nbsp;</td>
        <td width="20">&nbsp;</td>
        <td width="390">&nbsp;</td>
        <td width="20">&nbsp;</td>
      </tr>
 
      <tr> 
         <td>&nbsp;</td>
         <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Select 
              up to 5 States</font></td>
      </tr>

      <form action="CreateRegTemplateStep1.asp" method="post">
 
      <tr> 
         <td>&nbsp;</td>
         <td> <select name="States" size="11" multiple id="States">

                <%   
Dim objRS
Set objConn1 = Server.CreateObject("ADODB.Connection")
objConn1.Open Application("WaterSkiConn")
Set objRS = Server.CreateObject("ADODB.RecordSet")
objRS.ActiveConnection = objConn1

objRS.Open "SELECT * FROM [USStates] Where StateCode in ('AL','AK','AZ','AR','CA','CO','CT','DE','DC','FL','GA','HI','ID','IL','IN','IA','KS','KY','LA','ME','MD','MA','MI','MN','NS','MO','MT','NE','NV','NH','NJ','NM','NY','NC','ND','OH','OK','OR','PA','RI','SC','SD','TN','TX','UT','VT','VA','WA','WV','WI','WY') Order by StateName;" 

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

         <td>&nbsp;</td>

         <td valign="top">
         	 <font color="#FF0000" size="2" face="Verdana, Arial, Helvetica, 
         	 	sans-serif"><strong>Instructions:&nbsp; </strong></font>
         	 <font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
           Indicate the set of members you wish to retrieve into a 
           Registration Template, using the options which appear to 
           the left.&nbsp; You may select up to five states, and/or 
           check one or both of the special selection option boxes which 
           appear below that state selection window.&nbsp; To specify two
           or more states, hold down the Ctrl Key and click on each 
           state (up to 5 total) that you desire.&nbsp; Then click the 
           &#8220;Create Template&#8221; button, and I will then build an 
           Excel template containing those members for you to download.
           </font></td>
         
      </tr>

      <tr>
         <td>&nbsp;</td>
         <td>&nbsp;</td>
         <td>&nbsp;</td>
         <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Tournament 
              Date<font size="1"> (mm/dd/yyyy):</font></font></td>
      </tr>

   		<tr>
         <td>&nbsp;</td>
         <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
             <input type="checkbox" name="IncludeForeign" value="True">
   					 Select Foreign<font size="1"> </font></td>
         <td>&nbsp;</td>
         <td><input name="tournamentdate" type="text" id="tournamentdate" value="<%= Session("TournamentDate") %>">
      </tr>

			<tr>
         <td>&nbsp;</td>
         <td>&nbsp;</td>
         <td>&nbsp;</td>
         <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Tournament 
                    Year:<font size="1"> </font></font></td>
      </tr>

      <tr>
         <td>&nbsp;</td>
         <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
             <input type="checkbox" name="IncludeElite" value="True">
		         Select Open/Elite<font size="1"> </font></td>
         <td>&nbsp;</td>
         <td><select name="TournamentYear" id="TournamentYear"> 
         	<% Dim CurrentYear
				     CurrentYear = datepart("yyyy",date())
				     if len(session("TournamentName")) > 0 and epl_IsNumeric(left(Session("UserName"),2)) = true then
				    		TournamentYear = (2000 + left(Session("UserName"),2))
				     else
			    			TournamentYear = datepart("yyyy",date())
				     end if %>
			 		   <option value="<%= CurrentYear - 1 %>"<% if TournamentYear = (CurrentYear - 1) then response.write "Selected" %>><%= CurrentYear - 1 %></option>
             <option value="<%= CurrentYear %>"<% if TournamentYear = (CurrentYear) then response.write "Selected" %>><%= CurrentYear %></option>
             <option value="<%= CurrentYear + 1 %>"<% if TournamentYear = (CurrentYear + 1) then response.write "Selected" %>><%= CurrentYear + 1 %></option>
					   <option value="<%= CurrentYear + 2 %>"<% if TournamentYear = (CurrentYear + 2) then response.write "Selected" %>><%= CurrentYear + 2 %></option>
          </select></td>
      </tr>

			<tr>
         <td>&nbsp;</td>
         <td>&nbsp;</td>
         <td>&nbsp;</td>
         <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
         File format version:</font> </td>
   	  </tr>

			<tr>
         <td>&nbsp;</td>
         <td><input type="submit" name="NowWhat" value="Create Template"
         	title="Click here to build your template.&#13;This will take a minute or two;&#13;    So please stand by ... "></td>
         <td>&nbsp;</td>
		     <td>
					  <input name="FileFormat" type="radio" value="with_scores" checked>
					  <font size="2" face="Arial, Helvetica, sans-serif">with scores</font><font size="1" face="Arial, Helvetica, sans-serif"> (needed
					  for importing into WSTIMS)</font><br>
					  <input name="FileFormat" type="radio" value="without_scores">
					  <font size="2" face="Arial, Helvetica, sans-serif">without scores</font>
      	 </td>
      </tr>

     <tr> 
        <td>&nbsp;</td>
        <td colspan="4"><hr></td>
     </tr>

     <tr> 
        <td>&nbsp;</td>
        <td align="center"><input type="submit" name="NowWhat" value="Look Up&#13;Individual&#13;Members"></td>
        <td>&nbsp;</td>
        <td><font color="#FF0000" size="2" face="Verdana, Arial, Helvetica, 
        	 sans-serif"><strong>Alternatively:&nbsp; </strong></font><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
           If you need to get information for a few additional scattered members, we now 
           provide the means for you to look up individuals.&nbsp; You can then copy and 
           paste their information into your existing template.&nbsp; Click on the &#8220;Look 
           Up Individual Members&#8221; button to the left, to begin that process.</font></td>
     </tr>

   </form>

   </table>

    </td>
  </tr>
  
</table>
</body>
</html>






