<%
' ------------------------------------------------------------------------------------------------------------------
' ----- Tools_Definitions.asp      Tools where an include statement is put in the calling program
' ------------------------------------------------------------------------------------------------------------------



' ---------------------------------------------------------------------------
   SUB BuildClassDrop_NEW (sClassFieldName, sClassFieldValue, sEvent, DropStatus)
' ---------------------------------------------------------------------------

'IF sMemberID="000050050" AND LEFT(sTourID,6)="08M111" THEN 
'	response.write("<br>TEST")
'	response.write("<br>sEvent="&sEvent&"<br>")
'	response.write(TRIM(sEvent)="S")
'	response.write("<br>SGID"&sTSptsGrpID)
'END IF

'	response.write("<br>sEvent="&sEvent)
'	response.write("<br>sTSptsGrpID="&sTSptsGrpID&"<br>")
'	response.write(sTSptsGrpID="AWS")

'	response.write("<br>tClassE="&tClassE)
'	response.write("<br>")
'	response.write(tClassE>0 OR tClassL>0 OR tClassR>0)

%>
<form action="/rankings/Tools_Definitions.asp" method="post">
<Select name = "<%=sClassFieldName%>" <%=DropStatus%>><%


IF sEvent="S" AND sTSptsGrpID="AWS" THEN
	response.write("AWS S")	
	IF sClassR>0 THEN %><option value="R" <%IF sClassFieldName = "R" THEN Response.Write(" Selected")%>>R</option><% END IF
	IF sClassL>0 OR sClassR>0 THEN %><option value="L" <%IF sClassFieldValue = "L" THEN Response.Write(" Selected")%>>L</option><% END IF
	IF sClassE>0 OR sClassL>0 OR sClassR>0 THEN %><option value="E" <%IF sClassFieldValue = "E" THEN Response.Write(" Selected")%>>E</option><% END IF
	IF sClassC>0 OR sClassE>0 OR sClassL>0 OR sClassR>0 THEN %><option value="C" <%IF sClassFieldValue = "C" THEN Response.Write(" Selected")%>>C</option><% END IF
	IF Gr2AWS_SPulls THEN %><option value="F" <%IF sClassFieldValue = "F" THEN Response.Write(" Selected")%>>F</option><% END IF

	IF sClassFieldValue="R" AND NOT (sClassR>0) THEN %><option value="R" <%IF sClassFieldValue = "R" THEN Response.Write(" Selected")%>>Class R Invalid</option><% END IF  
	IF sClassFieldValue="L" AND NOT (sClassR>0 OR sClassL>0) THEN %><option value="L" <%IF sClassFieldValue = "L" THEN Response.Write(" Selected")%>>Class L Invalid</option><% END IF  
	IF sClassFieldValue="E" AND NOT (sClassR>0 OR sClassL>0 OR sClassE>0) THEN %><option value="E" <%IF sClassFieldValue = "E" THEN Response.Write(" Selected")%>>Class E Invalid</option><% END IF  
	IF sClassFieldValue="C" AND NOT (sClassR>0 OR sClassL>0 OR sClassE>0 OR sClassC>0) THEN %><option value="C" <%IF sClassFieldValue = "C" THEN Response.Write(" Selected")%>>Class C Invalid</option><% END IF  


ELSEIF sEvent="T" AND sTSptsGrpID="AWS" THEN

	IF tClassR>0 THEN %><option value="R" <%IF sClassFieldValue = "R" THEN Response.Write(" Selected")%>>R</option><% END IF
	IF tClassL>0 OR tClassR>0 THEN %><option value="L" <%IF sClassFieldValue = "L" THEN Response.Write(" Selected")%>>L</option><% END IF
	IF tClassE>0 OR tClassL>0 OR tClassR>0 THEN %><option value="E" <%IF sClassFieldValue = "E" THEN Response.Write(" Selected")%>>E</option><% END IF
	IF tClassC>0 OR tClassE>0 OR tClassL>0 OR tClassR>0 THEN %><option value="C" <%IF sClassFieldValue = "C" THEN Response.Write(" Selected")%>>C</option><% END IF
	IF Gr2AWS_TPulls>0 THEN %><option value="F" <%IF sClassFieldValue = "F" THEN Response.Write(" Selected")%>>F</option><% END IF

	IF sClassFieldValue="R" AND (NOT tClassR>0) THEN %><option value="R" <%IF sClassFieldValue = "R" THEN Response.Write(" Selected")%>>Class R Invalid</option><% END IF  
	IF sClassFieldValue="L" AND NOT (tClassR>0 OR tClassL>0) THEN %><option value="L" <%IF sClassFieldValue = "L" THEN Response.Write(" Selected")%>>Class L Invalid</option><% END IF  
	IF sClassFieldValue="E" AND NOT (tClassR>0 OR tClassL>0 OR tClassE>0) THEN %><option value="E" <%IF sClassFieldValue = "E" THEN Response.Write(" Selected")%>>Class E Invalid</option><% END IF  
	IF sClassFieldValue="C" AND NOT (tClassR>0 OR tClassL>0 OR tClassE>0 OR tClassC>0) THEN %><option value="C" <%IF sClassFieldValue = "C" THEN Response.Write(" Selected")%>>Class C Invalid</option><% END IF  

ELSEIF sEvent="J" AND sTSptsGrpID="AWS" THEN

	IF jClassR>0 THEN %><option value="R" <%IF sClassFieldValue = "R" THEN Response.Write(" Selected")%>>R</option><% END IF
	IF (jClassL>0 OR jClassR>0) THEN %><option value="L" <%IF sClassFieldValue = "L" THEN Response.Write(" Selected")%>>L</option><% END IF
	IF (jClassE>0 OR jClassL>0 OR jClassR>0) THEN %><option value="E" <%IF sClassFieldValue = "E" THEN Response.Write(" Selected")%>>E</option><% END IF
	IF (jClassC>0 OR jClassE>0 OR jClassL>0 OR jClassR>0) THEN %><option value="C" <%IF sClassFieldValue = "C" THEN Response.Write(" Selected")%>>C</option><% END IF
	IF jClassN>0 THEN %><option value="N" <%IF sClassFieldValue = "N" THEN Response.Write(" Selected")%>>N</option><% END IF


	IF sClassFieldValue="R" AND (NOT jClassR>0) THEN %><option value="R" <%IF sClassFieldValue = "R" THEN Response.Write(" Selected")%>>Class R Invalid</option><% END IF  
	IF sClassFieldValue="L" AND NOT (jClassR>0 OR jClassL>0) THEN %><option value="L" <%IF sClassFieldValue = "L" THEN Response.Write(" Selected")%>>Class L Invalid</option><% END IF  
	IF sClassFieldValue="E" AND NOT (jClassR>0 OR jClassL>0 OR jClassE>0) THEN %><option value="E" <%IF sClassFieldValue = "E" THEN Response.Write(" Selected")%>>Class E Invalid</option><% END IF  
	IF sClassFieldValue="C" AND NOT (jClassR>0 OR jClassL>0 OR jClassE>0 OR jClassC>0) THEN %><option value="C" <%IF sClassFieldValue = "C" THEN Response.Write(" Selected")%>>Class C Invalid</option><% END IF  

ELSEIF sEvent="S" AND sTSptsGrpID="AKA" THEN

	IF sKSlalomClassT THEN %><option value="T" <%IF sClassFieldValue = "T" THEN Response.Write(" Selected")%>>T</option><% END IF
	IF sKSlalomClassQ THEN %><option value="Q" <%IF sClassFieldValue = "Q" THEN Response.Write(" Selected")%>>Q</option><% END IF
ELSEIF sEvent="T" AND sTSptsGrpID="AKA" THEN

	IF sKTrickClassT THEN %><option value="T" <%IF sClassFieldValue = "T" THEN Response.Write(" Selected")%>>T</option><% END IF
	IF sKTrickClassQ THEN %><option value="Q" <%IF sClassFieldValue = "Q" THEN Response.Write(" Selected")%>>Q</option><% END IF
ELSEIF sEvent="KP" AND sTSptsGrpID="AKA" THEN
	IF sKFreeClassT THEN %><option value="T" <%IF sClassFieldValue = "T" THEN Response.Write(" Selected")%>>T</option><% END IF
	IF sKFreeClassQ THEN %><option value="Q" <%IF sClassFieldValue = "Q" THEN Response.Write(" Selected")%>>Q</option><% END IF
ELSEIF sEvent="KR" AND sTSptsGrpID="AKA" THEN
	IF sKFlipClassT THEN %><option value="T" <%IF sClassFieldValue = "T" THEN Response.Write(" Selected")%>>T</option><% END IF
	IF sKFlipClassQ THEN %><option value="Q" <%IF sClassFieldValue = "Q" THEN Response.Write(" Selected")%>>Q</option><% END IF
ELSEIF (sEvent="WB" OR sEvent="WS" OR sEvent="WU") AND sTSptsGrpID="USW" THEN %>
        <option value="W" <%If sClassFieldValue = "W" Then Response.Write(" Selected")%>>W</option>
        <option value="F" <%If sClassFieldValue = "F" Then Response.Write(" Selected")%>>F</option><%
ELSEIF sEvent="S" AND sTSptsGrpID="NCW" THEN	
	IF sClassC>0 THEN %><option value="C" <%IF sClassFieldValue = "C" THEN Response.Write(" Selected")%>>C</option><% END IF
	IF Gr2AWS_SPulls>0 THEN %><option value="F" <%IF sClassFieldValue = "F" THEN Response.Write(" Selected")%>>F</option><% END IF
	IF sClassL>0 THEN %><option value="L" <%IF sClassFieldValue = "L" THEN Response.Write(" Selected")%>>L</option><% END IF
	IF sClassR>0 THEN %><option value="R" <%IF sClassFieldValue = "R" THEN Response.Write(" Selected")%>>R</option><% END IF
ELSEIF sEvent="T" AND sTSptsGrpID="NCW" THEN	
	IF tClassC>0 THEN %><option value="C" <%IF sClassFieldValue = "C" THEN Response.Write(" Selected")%>>C</option><% END IF
	IF Gr2AWS_TPulls>0 THEN %><option value="F" <%IF sClassFieldValue = "F" THEN Response.Write(" Selected")%>>F</option><% END IF
	IF tClassL>0 THEN %><option value="L" <%IF sClassFieldValue = "L" THEN Response.Write(" Selected")%>>L</option><% END IF
	IF tClassR>0 THEN %><option value="R" <%IF sClassFieldValue = "R" THEN Response.Write(" Selected")%>>R</option><% END IF
ELSEIF sEvent="J" AND sTSptsGrpID="NCW" THEN
	IF jClassC>0 THEN %><option value="C" <%IF sClassFieldValue = "C" THEN Response.Write(" Selected")%>>C</option><% END IF
	IF jClassN>0 THEN %><option value="N" <%IF sClassFieldValue = "N" THEN Response.Write(" Selected")%>>N</option><% END IF
	IF jClassL>0 THEN %><option value="L" <%IF sClassFieldValue = "L" THEN Response.Write(" Selected")%>>L</option><% END IF
	IF jClassR>0 THEN %><option value="R" <%IF sClassFieldValue = "R" THEN Response.Write(" Selected")%>>R</option><% END IF
END IF

%></select><%

END SUB




' ---------------------------------------------------------------------------
   SUB BuildClassDrop (sClassFieldName, sClassFieldValue, sEvent, DropStatus)
' ---------------------------------------------------------------------------

'IF sMemberID="000050050" AND LEFT(sTourID,6)="08M111" THEN 
'	response.write("<br>TEST")
'	response.write("<br>sEvent="&sEvent&"<br>")
'	response.write(TRIM(sEvent)="S")
'	response.write("<br>SGID"&sTSptsGrpID)
'END IF

%>
<form action="/rankings/Tools_Definitions.asp" method="post">
<Select name = "<%=sClassFieldName%>" <%=DropStatus%>><%


IF sEvent="S" AND sTSptsGrpID="AWS" THEN

	IF sTHSClassR THEN %><option value="R" <%IF sClassFieldName = "R" THEN Response.Write(" Selected")%>>R</option><% END IF
	IF sTHSClassL OR sTHSClassR THEN %><option value="L" <%IF sClassFieldValue = "L" THEN Response.Write(" Selected")%>>L</option><% END IF
	IF sTHSClassE OR sTHSClassL OR sTHSClassR THEN %><option value="E" <%IF sClassFieldValue = "E" THEN Response.Write(" Selected")%>>E</option><% END IF
	IF sTHSClassC OR sTHSClassE OR sTHSClassL OR sTHSClassR THEN %><option value="C" <%IF sClassFieldValue = "C" THEN Response.Write(" Selected")%>>C</option><% END IF
	IF sTHSClassN OR sTHSClassC THEN %><option value="N" <%IF sClassFieldValue = "N" THEN Response.Write(" Selected")%>>N</option><% END IF
	IF sTHSClassF THEN %><option value="F" <%IF sClassFieldValue = "F" THEN Response.Write(" Selected")%>>F</option><% END IF

	IF sClassFieldValue="R" AND NOT (sTHSClassR) THEN %><option value="R" <%IF sClassFieldValue = "R" THEN Response.Write(" Selected")%>>Class R Invalid</option><% END IF  
	IF sClassFieldValue="L" AND NOT (sTHSClassR OR sTHSClassL) THEN %><option value="L" <%IF sClassFieldValue = "L" THEN Response.Write(" Selected")%>>Class L Invalid</option><% END IF  
	IF sClassFieldValue="E" AND NOT (sTHSClassR OR sTHSClassL OR sTHSClassE) THEN %><option value="E" <%IF sClassFieldValue = "E" THEN Response.Write(" Selected")%>>Class E Invalid</option><% END IF  
	IF sClassFieldValue="C" AND NOT (sTHSClassR OR sTHSClassL OR sTHSClassE OR sTHSClassC) THEN %><option value="C" <%IF sClassFieldValue = "C" THEN Response.Write(" Selected")%>>Class C Invalid</option><% END IF  


ELSEIF sEvent="T" AND sTSptsGrpID="AWS" THEN	
	IF sTHTClassR THEN %><option value="R" <%IF sClassFieldValue = "R" THEN Response.Write(" Selected")%>>R</option><% END IF
	IF (sTHTClassL OR sTHTClassR) THEN %><option value="L" <%IF sClassFieldValue = "L" THEN Response.Write(" Selected")%>>L</option><% END IF
	IF (sTHTClassE OR sTHTClassL OR sTHTClassR) THEN %><option value="E" <%IF sClassFieldValue = "E" THEN Response.Write(" Selected")%>>E</option><% END IF
	IF (sTHTClassC OR sTHTClassE OR sTHTClassL OR sTHTClassR) THEN %><option value="C" <%IF sClassFieldValue = "C" THEN Response.Write(" Selected")%>>C</option><% END IF
	IF (sTHTClassN OR sTHTClassC) THEN %><option value="N" <%IF sClassFieldValue = "N" THEN Response.Write(" Selected")%>>N</option><% END IF
	IF sTHTClassF THEN %><option value="F" <%IF sClassFieldValue = "F" THEN Response.Write(" Selected")%>>F</option><% END IF

	IF sClassFieldValue="R" AND (NOT sTHTClassR) THEN %><option value="R" <%IF sClassFieldValue = "R" THEN Response.Write(" Selected")%>>Class R Invalid</option><% END IF  
	IF sClassFieldValue="L" AND NOT (sTHTClassR OR sTHTClassL) THEN %><option value="L" <%IF sClassFieldValue = "L" THEN Response.Write(" Selected")%>>Class L Invalid</option><% END IF  
	IF sClassFieldValue="E" AND NOT (sTHTClassR OR sTHTClassL OR sTHTClassE) THEN %><option value="E" <%IF sClassFieldValue = "E" THEN Response.Write(" Selected")%>>Class E Invalid</option><% END IF  
	IF sClassFieldValue="C" AND NOT (sTHTClassR OR sTHTClassL OR sTHTClassE OR sTHTClassC) THEN %><option value="C" <%IF sClassFieldValue = "C" THEN Response.Write(" Selected")%>>Class C Invalid</option><% END IF  

ELSEIF sEvent="J" AND sTSptsGrpID="AWS" THEN
	IF sTHJClassR THEN %><option value="R" <%IF sClassFieldValue = "R" THEN Response.Write(" Selected")%>>R</option><% END IF
	IF (sTHJClassL OR sTHJClassR) THEN %><option value="L" <%IF sClassFieldValue = "L" THEN Response.Write(" Selected")%>>L</option><% END IF
	IF (sTHJClassE OR sTHJClassL OR sTHJClassR) THEN %><option value="E" <%IF sClassFieldValue = "E" THEN Response.Write(" Selected")%>>E</option><% END IF
	IF (sTHJClassC OR sTHJClassE OR sTHJClassL OR sTHJClassR) THEN %><option value="C" <%IF sClassFieldValue = "C" THEN Response.Write(" Selected")%>>C</option><% END IF

	IF sClassFieldValue="R" AND (NOT sTHJClassR) THEN %><option value="R" <%IF sClassFieldValue = "R" THEN Response.Write(" Selected")%>>Class R Invalid</option><% END IF  
	IF sClassFieldValue="L" AND NOT (sTHJClassR OR sTHJClassL) THEN %><option value="L" <%IF sClassFieldValue = "L" THEN Response.Write(" Selected")%>>Class L Invalid</option><% END IF  
	IF sClassFieldValue="E" AND NOT (sTHJClassR OR sTHJClassL OR sTHJClassE) THEN %><option value="E" <%IF sClassFieldValue = "E" THEN Response.Write(" Selected")%>>Class E Invalid</option><% END IF  
	IF sClassFieldValue="C" AND NOT (sTHJClassR OR sTHJClassL OR sTHJClassE OR sTHJClassC) THEN %><option value="C" <%IF sClassFieldValue = "C" THEN Response.Write(" Selected")%>>Class C Invalid</option><% END IF  

ELSEIF sEvent="S" AND sTSptsGrpID="AKA" THEN
	IF sKSlalomClassT THEN %><option value="T" <%IF sClassFieldValue = "T" THEN Response.Write(" Selected")%>>T</option><% END IF
	IF sKSlalomClassQ THEN %><option value="Q" <%IF sClassFieldValue = "Q" THEN Response.Write(" Selected")%>>Q</option><% END IF
ELSEIF sEvent="T" AND sTSptsGrpID="AKA" THEN
	IF sKTrickClassT THEN %><option value="T" <%IF sClassFieldValue = "T" THEN Response.Write(" Selected")%>>T</option><% END IF
	IF sKTrickClassQ THEN %><option value="Q" <%IF sClassFieldValue = "Q" THEN Response.Write(" Selected")%>>Q</option><% END IF
ELSEIF sEvent="KP" AND sTSptsGrpID="AKA" THEN
	IF sKFreeClassT THEN %><option value="T" <%IF sClassFieldValue = "T" THEN Response.Write(" Selected")%>>T</option><% END IF
	IF sKFreeClassQ THEN %><option value="Q" <%IF sClassFieldValue = "Q" THEN Response.Write(" Selected")%>>Q</option><% END IF
ELSEIF sEvent="KR" AND sTSptsGrpID="AKA" THEN
	IF sKFlipClassT THEN %><option value="T" <%IF sClassFieldValue = "T" THEN Response.Write(" Selected")%>>T</option><% END IF
	IF sKFlipClassQ THEN %><option value="Q" <%IF sClassFieldValue = "Q" THEN Response.Write(" Selected")%>>Q</option><% END IF
ELSEIF (sEvent="WB" OR sEvent="WS" OR sEvent="WU") AND sTSptsGrpID="USW" THEN %>
        <option value="W" <%If sClassFieldValue = "W" Then Response.Write(" Selected")%>>W</option>
        <option value="F" <%If sClassFieldValue = "F" Then Response.Write(" Selected")%>>F</option><%
ELSEIF sEvent="S" AND sTSptsGrpID="NCW" THEN	
	IF sTHSClassC THEN %><option value="C" <%IF sClassFieldValue = "C" THEN Response.Write(" Selected")%>>C</option><% END IF
	IF sTHSClassN THEN %><option value="N" <%IF sClassFieldValue = "N" THEN Response.Write(" Selected")%>>N</option><% END IF
	IF sTHSClassL THEN %><option value="L" <%IF sClassFieldValue = "L" THEN Response.Write(" Selected")%>>L</option><% END IF
	IF sTHSClassR THEN %><option value="R" <%IF sClassFieldValue = "R" THEN Response.Write(" Selected")%>>R</option><% END IF
ELSEIF sEvent="T" AND sTSptsGrpID="NCW" THEN	
	IF sTHTClassC THEN %><option value="C" <%IF sClassFieldValue = "C" THEN Response.Write(" Selected")%>>C</option><% END IF
	IF sTHTClassN THEN %><option value="N" <%IF sClassFieldValue = "N" THEN Response.Write(" Selected")%>>N</option><% END IF
	IF sTHTClassL THEN %><option value="L" <%IF sClassFieldValue = "L" THEN Response.Write(" Selected")%>>L</option><% END IF
	IF sTHTClassR THEN %><option value="R" <%IF sClassFieldValue = "R" THEN Response.Write(" Selected")%>>R</option><% END IF
ELSEIF sEvent="J" AND sTSptsGrpID="NCW" THEN
	IF sTHJClassC THEN %><option value="C" <%IF sClassFieldValue = "C" THEN Response.Write(" Selected")%>>C</option><% END IF
	IF sTHJClassN THEN %><option value="N" <%IF sClassFieldValue = "N" THEN Response.Write(" Selected")%>>N</option><% END IF
	IF sTHJClassL THEN %><option value="L" <%IF sClassFieldValue = "L" THEN Response.Write(" Selected")%>>L</option><% END IF
	IF sTHJClassR THEN %><option value="R" <%IF sClassFieldValue = "R" THEN Response.Write(" Selected")%>>R</option><% END IF

END IF

%></select><%

END SUB






' ------------------------------------------
   SUB LoadTeam (TeamSelected, TeamStatus)
' -------------------------------------------


' ----  Define TEAM drop down from TeamTableName ----
set rsTeam=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT * FROM "&TeamTableName
sSQL = sSQL + " WHERE SptsGrpID = '"&Session("sTSptsGrpID")&"'"
sSQL = sSQL + " ORDER BY TeamName"
rsTeam.open sSQL, SConnectionToTRATable

IF rsTeam.eof THEN TeamStatus="disabled"

%>
<select name="TeamDrop" <%=TeamStatus%>><%

IF TRIM(LeagueSelected) = "" THEN
	response.write("<option value =""None"" selected>None</option><br>")
ELSE
	response.write("<option value =""None"">None</option><br>")
END IF

IF NOT rsTeam.eof THEN 
  	rsTeam.movefirst
  	DO WHILE NOT rsTeam.eof
		IF TRIM(rsTeam("TeamID")) = TRIM(TeamSelected) THEN
			response.write("<option value =""" & rsTeam("TeamID") &""" selected>"&rsTeam("TeamName")&"</option><br>")
			Session("TeamName")=rsTeam("TeamName")
    		ELSE
			response.write("<option value =""" & rsTeam("TeamID") &""">"&rsTeam("TeamName")&"</option><br>")
		END IF	
		rsTeam.moveNEXT
	LOOP
ELSE
	response.write("<option value =""None"" selected>None Available</option>")
END IF  %>

</select><%

rsTeam.close

END SUB



' -----------------------------------------------
   SUB LoadLeague (LeagueSelected, LeagueStatus)
' -----------------------------------------------


' ----  Define LEAGUE drop down from LeagueTableName ----
set rsLeague=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT * FROM "&LeagueTableName
'sSQL = sSQL + " WHERE SptsGrpID = '"&Session("sTSptsGrpID")&"'"
sSQL = sSQL + " ORDER BY LeagueName"
rsLeague.open sSQL, SConnectionToTRATable

IF rsLeague.eof THEN LeagueStatus="disabled"

%>
<select name="LeagueDrop" <%=LeagueStatus%>><%

IF TRIM(LeagueSelected) = "" THEN
	response.write("<option value =""None"" selected>None</option><br>")
ELSE
	response.write("<option value =""None"">None</option><br>")
END IF

IF NOT rsLeague.eof THEN 
  	rsLeague.movefirst
  	DO WHILE NOT rsLeague.eof
		IF TRIM(rsLeague("LeagueID")) = TRIM(LeagueSelected) THEN
			response.write("<option value =""" & rsLeague("LeagueID") &""" selected>"&rsLeague("LeagueName")&"</option><br>")
			Session("LeagueName")=rsLeague("LeagueName")
    		ELSE
			response.write("<option value =""" & rsLeague("LeagueID") &""">"&rsLeague("LeagueName")&"</option><br>")
		END IF	
		rsLeague.moveNEXT
	LOOP
ELSE
	response.write("<option value =""NA"" selected>None Available</option>")
END IF  %>

</select><%

rsLeague.close

END SUB




'-----------------------------------------------------------------------------
 SUB LoadRampPulldown_Nov1_2010 (JumpDiv, RampFieldName, sRampHeight, sRampStatus)
'-----------------------------------------------------------------------------


' Builds RAMP Dropdown for event selected

sSQL = "SELECT DT.div, DT.div_name, DT.Ramp1, DT.Ramp2, DT.Ramp3, DT.Up_Age"
sSQL = sSQL + "  FROM "&DivisionsTableName&" as DT"
sSQL = sSQL + " WHERE DT.SkiYearID = (Select SkiYearID from "&SkiYearTableName&" where DefaultYear=1)"
sSQL = sSQL + " AND DT.div = '"&JumpDiv&"'"

SET rsDiv=Server.CreateObject("ADODB.recordset")
rsDiv.open sSQL, SConnectionToTRATable



IF NOT rsDiv.eof THEN

    ' --- Condition tests for existence of any valid ramp height in RAMP2 --- 	
   IF rsDiv("Ramp2") = 0.235 OR rsDiv("Ramp2") = 0.255 OR rsDiv("Ramp2") = 0.275 THEN  %>

	<select name="<%=RampFieldName%>" <%=sRampStatus%>><%

	IF rsDiv("Ramp1") = 0.235 OR rsDiv("Ramp2") = 0.235 OR rsDiv("Ramp3") = 0.235 THEN
	  %><option value = "0.235" <% IF sRampHeight = "0.235" OR sRampHeight="5.0" THEN Response.Write(" selected ") %>>5.0</Option><br><%
	END IF 

	IF rsDiv("Ramp1") = 0.255 OR rsDiv("Ramp2") = 0.255 OR rsDiv("Ramp3") = 0.255 THEN	
	  %><option value = "0.255" <% IF sRampHeight = "0.255" OR sRampHeight="5.5" THEN Response.Write(" selected ") %>>5.5</Option><br><%
	END IF

	IF rsDiv("Ramp1") = 0.275 OR rsDiv("Ramp2") = 0.275 OR rsDiv("Ramp3") = 0.275 THEN	
	  %><option value = "0.275" <% IF sRampHeight = "0.275" OR sRampHeight = "0.271" OR sRampHeight="6.0" THEN Response.Write(" selected ") %>>6.0</Option><br><%
	END IF %>

	</select><%
 	IF sRampStatus="disabled" THEN %>	
		<input type="hidden" name="<%=RampFieldName%>" value="<% =sRampHeight %>" ><%
	END IF
   ELSE
	'SELECT CASE rsDiv("Ramp1")
	'	CASE "0.235"
	'		sRampHeight = "5.0"
	'		'RampFieldName = "5.0"
	'	CASE "0.255"
	'		sRampHeight = "5.5"
	'		'RampFieldName = "5.5"
	'	CASE "0.275", "0.271"
	'		sRampHeight = "6.0"
	'		'RampFieldName = "6.0"
	'END SELECT
	sRampHeight=rsDiv("Ramp1")	

	response.write(sRampHeight) %>
	<input type="hidden" name="<%=RampFieldName%>" value="<% =sRampHeight %>" > <%
		
   END IF

END IF

rsDiv.close


END SUB




'-----------------------------------------------------------------------------
 SUB LoadRampPulldownNEW (JumpDiv, RampFieldName, sRampHeight, sRampStatus)
'-----------------------------------------------------------------------------


' Builds RAMP Dropdown for event selected

sSQL = "SELECT DT.div, DT.div_name, DT.Ramp1, DT.Ramp2, DT.Ramp3, DT.Up_Age"
sSQL = sSQL + "  FROM "&DivisionsTableName&" as DT"
sSQL = sSQL + " WHERE DT.SkiYearID = (Select SkiYearID from "&SkiYearTableName&" where DefaultYear=1)"
sSQL = sSQL + " AND DT.div = '"&JumpDiv&"'"

SET rsDiv=Server.CreateObject("ADODB.recordset")
rsDiv.open sSQL, SConnectionToTRATable



IF NOT rsDiv.eof THEN

    ' --- Condition tests for existence of any valid ramp height in RAMP2 --- 	
   IF rsDiv("Ramp2") = 0.235 OR rsDiv("Ramp2") = 0.255 OR rsDiv("Ramp2") = 0.275 THEN  %>

	<select name="<%=RampFieldName%>" <%=sRampStatus%>><%

	IF rsDiv("Ramp1") = 0.235 OR rsDiv("Ramp2") = 0.235 OR rsDiv("Ramp3") = 0.235 THEN
	  %><option value = "0.235" <% IF sRampHeight = "0.235" OR sRampHeight="5.0" THEN Response.Write(" selected ") %>>5.0</Option><br><%
	END IF 

	IF rsDiv("Ramp1") = 0.255 OR rsDiv("Ramp2") = 0.255 OR rsDiv("Ramp3") = 0.255 THEN	
	  %><option value = "0.255" <% IF sRampHeight = "0.255" OR sRampHeight="5.5" THEN Response.Write(" selected ") %>>5.5</Option><br><%
	END IF

	IF rsDiv("Ramp1") = 0.275 OR rsDiv("Ramp2") = 0.275 OR rsDiv("Ramp3") = 0.275 THEN	
	  %><option value = "0.275" <% IF sRampHeight = "0.275" OR sRampHeight = "0.271" OR sRampHeight="6.0" THEN Response.Write(" selected ") %>>6.0</Option><br><%
	END IF %>

	</select><%
 	IF sRampStatus="disabled" THEN %>	
		<input type="hidden" name="<%=RampFieldName%>" value="<% =sRampHeight %>" ><%
	END IF
   ELSE
	SELECT CASE rsDiv("Ramp1")
		CASE "0.235"
			sRampHeight = "5.0"
			'RampFieldName = "5.0"
		CASE "0.255"
			sRampHeight = "5.5"
			'RampFieldName = "5.5"
		CASE "0.275", "0.271"
			sRampHeight = "6.0"
			'RampFieldName = "6.0"
	END SELECT

	response.write(sRampHeight) %>
	<input type="hidden" name="<%=RampFieldName%>" value="<% =sRampHeight %>" > <%
		
   END IF

END IF

rsDiv.close


END SUB






'-----------------------------------------------------------------------------
 SUB LoadRampPulldownRegister_11072015 (JumpDiv, RampFieldName, sRampHeight, sRampStatus)
'-----------------------------------------------------------------------------

' Builds RAMP Dropdown for event selected

sSQL = "SELECT DT.div, DT.div_name, DT.Ramp1, DT.Ramp2, DT.Ramp3, DT.Up_Age FROM "&DivisionsTableName&" as DT"
sSQL = sSQL + " WHERE DT.SkiYearID = (Select SkiYearID from "&SkiYearTableName&" where DefaultYear=1)"
sSQL = sSQL + " AND DT.div = '"&JumpDiv&"'"

SET rsDiv=Server.CreateObject("ADODB.recordset")
rsDiv.open sSQL, SConnectionToTRATable



IF NOT rsDiv.eof THEN

    ' --- Condition tests for existence of any valid ramp height in RAMP2 --- 	
	IF JumpDiv="MB" OR JumpDiv="MA" OR JumpDiv="M9" OR JumpDiv="M8" OR JumpDiv="M7" OR JumpDiv="M6" OR JumpDiv="WB" OR JumpDiv="WA" OR JumpDiv="W9" OR JumpDiv="W8" OR JumpDiv="W7" OR JumpDiv="W6" THEN  
				%>
				<select name="<% =RampFieldName %>" style="width:8em;" <%=sRampStatus%>>
					<option value = "5.0" <% IF sRampHeight = "0.235" OR sRampHeight="5.0" THEN Response.Write(" selected ") %>>5.0 ft (0.235)</Option><br>
				  <option value = "4.5" <% IF sRampHeight = "0.215" OR sRampHeight="4.5" THEN Response.Write(" selected ") %>>4.5 ft (0.215)</Option><br>
				</select>
				<%

 	ELSE

				%><select id="<% =RampFieldName %>" name="<% =RampFieldName %>" style="width:8em;" <% =sRampStatus %>><%

				IF rsDiv("Ramp1") = 0.235 OR rsDiv("Ramp2") = 0.235 OR rsDiv("Ramp3") = 0.235 THEN
	  				%><option value = "5.0" <% IF sRampHeight = "0.235" OR sRampHeight="5.0" THEN Response.Write(" selected ") %>>5.0 ft (0.235)</Option><br><%
				END IF 

				IF rsDiv("Ramp1") = 0.255 OR rsDiv("Ramp2") = 0.255 OR rsDiv("Ramp3") = 0.255 THEN	
	  				%><option value = "5.5" <% IF sRampHeight = "0.255" OR sRampHeight="5.5" THEN Response.Write(" selected ") %>>5.5 ft (0.255)</Option><br><%
				END IF

				IF rsDiv("Ramp1") = 0.275 OR rsDiv("Ramp2") = 0.275 OR rsDiv("Ramp3") = 0.275 THEN	
	  				%><option value = "6.0" <% IF sRampHeight = "0.275" OR sRampHeight = "0.271" OR sRampHeight="6.0" THEN Response.Write(" selected ") %>>6.0 ft (0.275)</Option><br><%
				END IF 
				%></select><%

		END IF

ELSE		' --- Division not found or established ---
				%><select id="<% =RampFieldName %>" name="<% =RampFieldName %>" style="width:8em;" <% =sRampStatus %>>
	  			<option value = "4.5" <% IF sRampHeight = "0.215" OR sRampHeight="4.5" THEN Response.Write(" selected ") %>>4.5 ft (0.215)</Option><br>
	  			<option value = "5.0" <% IF sRampHeight = "0.235" OR sRampHeight="5.0" THEN Response.Write(" selected ") %>>5.0 ft (0.235)</Option><br>
	  			<option value = "5.5" <% IF sRampHeight = "0.255" OR sRampHeight="5.5" THEN Response.Write(" selected ") %>>5.5 ft (0.255)</Option><br>
					<option value = "6.0" <% IF sRampHeight = "0.275" OR sRampHeight = "0.271" OR sRampHeight="6.0" THEN Response.Write(" selected ") %>>6.0 ft (0.275)</Option><br>
				</select><%

	

END IF

rsDiv.close


END SUB





'-----------------------------------------------------------------------------
 SUB LoadRampPulldownRegister (JumpDiv, RampFieldName, sRampHeight, sRampStatus)
'-----------------------------------------------------------------------------


' Builds RAMP Dropdown for event selected

sSQL = "SELECT DT.div, DT.div_name, DT.Ramp1, DT.Ramp2, DT.Ramp3, DT.Up_Age FROM "&DivisionsTableName&" as DT"
sSQL = sSQL + " WHERE DT.SkiYearID = (Select SkiYearID from "&SkiYearTableName&" where DefaultYear=1)"
sSQL = sSQL + " AND DT.div = '"&JumpDiv&"'"

SET rsDiv=Server.CreateObject("ADODB.recordset")
rsDiv.open sSQL, SConnectionToTRATable





IF NOT rsDiv.eof THEN


    ' --- Condition tests for existence of any valid ramp height in RAMP2 --- 	
   IF rsDiv("Ramp2") = "0.235" OR rsDiv("Ramp2") = "0.255" OR rsDiv("Ramp2") = "0.275" OR rsDiv("Ramp2") = "0.215" THEN  

				%>
				<select id="<%=RampFieldName%>" name="<%=RampFieldName%>" <%=sRampStatus%>><%

				IF rsDiv("Ramp1") = 0.235 OR rsDiv("Ramp2") = 0.235 OR rsDiv("Ramp3") = 0.235 THEN
	  				%><option value = "5.0" <% IF sRampHeight = "0.235" OR sRampHeight="5.0" THEN Response.Write(" selected ") %>>5.0</Option><br><%
				END IF 

				IF rsDiv("Ramp1") = 0.255 OR rsDiv("Ramp2") = 0.255 OR rsDiv("Ramp3") = 0.255 THEN	
	  				%><option value = "5.5" <% IF sRampHeight = "0.255" OR sRampHeight="5.5" THEN Response.Write(" selected ") %>>5.5</Option><br><%
				END IF

				IF rsDiv("Ramp1") = 0.275 OR rsDiv("Ramp2") = 0.275 OR rsDiv("Ramp3") = 0.275 THEN	
	  				%><option value = "6.0" <% IF sRampHeight = "0.275" OR sRampHeight = "0.271" OR sRampHeight="6.0" THEN Response.Write(" selected ") %>>6.0</Option><br><%
				END IF %>
				</select><%

			 	IF sRampStatus="disabled" THEN %>	
						<input type="hidden" name="<%=RampFieldName%>" value="<% =sRampHeight %>" ><%
				END IF

	ELSEIF JumpDiv="MB" OR JumpDiv="MA" OR JumpDiv="M9" OR JumpDiv="M8" OR JumpDiv="M7" OR JumpDiv="WB" OR JumpDiv="WA" OR JumpDiv="W9" OR JumpDiv="W8" OR JumpDiv="W7" THEN  %>
				<select name="<%=RampFieldName%>" <%=sRampStatus%>>
					<option value = "5.0" <% IF sRampHeight = "0.235" OR sRampHeight="5.0" THEN Response.Write(" selected ") %>>5.0</Option><br>
				  <option value = "4.5" <% IF sRampHeight = "0.215" OR sRampHeight="4.5" THEN Response.Write(" selected ") %>>4.5</Option><br>
				</select><%
			 	IF sRampStatus="disabled" THEN %>	
						<input type="hidden" name="<%=RampFieldName%>" value="<% =sRampHeight %>" ><%
				END IF

  ELSE
			SELECT CASE rsDiv("Ramp1")
				CASE "0.215"
						sRampHeight = "4.5"
				CASE "0.235"
						sRampHeight = "5.0"
				CASE "0.255"
						sRampHeight = "5.5"
				CASE "0.275", "0.271"
						sRampHeight = "6.0"
			END SELECT

			response.write(sRampHeight) %>
			<input type="hidden" name="<%=RampFieldName%>" value="<% =sRampHeight %>" > <%
		
   END IF

END IF

rsDiv.close


END SUB



'--------------------------------------------
 SUB LoadRampPulldown (JumpDiv)
'--------------------------------------------

' Builds RAMP Dropdown for event selected


sSQL = "SELECT DT.div, DT.div_name, DT.Ramp1, DT.Ramp2, DT.Ramp3, DT.Up_Age FROM "&DivisionsTableName&" as DT"
sSQL = sSQL + " WHERE DT.SkiYearID = (Select SkiYearID from "&SkiYearTableName&" where DefaultYear=1)"
sSQL = sSQL + " AND DT.div = '"&JumpDiv&"'"

SET rsDiv=Server.CreateObject("ADODB.recordset")
rsDiv.open sSQL, SConnectionToTRATable



IF NOT rsDiv.eof THEN

    ' --- Condition tests for existence of any valid ramp height in RAMP2 --- 	
   IF rsDiv("Ramp2") = 0.235 OR rsDiv("Ramp2") = 0.255 OR rsDiv("Ramp2") = 0.275 THEN  %>

	<select name="sRampHeight" value="<% =sRampHeight %>" <%=RampStatus %>><%

	IF rsDiv("Ramp1") = 0.235 OR rsDiv("Ramp2") = 0.235 OR rsDiv("Ramp3") = 0.235 THEN
	  %><option value = "5.0" <% IF sRampHeight = "5.0" THEN Response.Write(" selected ") %>>5.0</Option><br><%
	END IF 

	IF rsDiv("Ramp1") = 0.255 OR rsDiv("Ramp2") = 0.255 OR rsDiv("Ramp3") = 0.255 THEN	
	  %><option value = "5.5" <% IF sRampHeight = "5.5" THEN Response.Write(" selected ") %>>5.5</Option><br><%
	END IF

	IF rsDiv("Ramp1") = 0.275 OR rsDiv("Ramp2") = 0.275 OR rsDiv("Ramp3") = 0.275 THEN	
	  %><option value = "6.0" <% IF sRampHeight = "6.0" THEN Response.Write(" selected ") %>>6.0</Option><br><%
	END IF %>

	</select><%
   ELSE
	SELECT CASE rsDiv("Ramp1")
		CASE "0.235"
			sRampHeight = "5.0"
		CASE "0.255"
			sRampHeight = "5.5"
		CASE "0.275"
			sRampHeight = "6.0"
	END SELECT

	response.write(sRampHeight) %>
	<input type="hidden" name="sRampHeight" value="<% =sRampHeight %>" > <%
		
   END IF

END IF

rsDiv.close


END SUB


' -------------------------------------------------------------------------
  SUB LoadQualificationsOverrideDropDown (QfyOverrideName, QfyOverride, OverrideStatus)
' -------------------------------------------------------------------------

'IF sMemberID="700002960" THEN 
'	response.write("<br>In Tools="&QfyOverride)
'END IF

%>
<select name="<%=QfyOverrideName%>" <% =OverrideStatus %> style="width:12em">
  <option value ="" <%IF QfyOverride = "" THEN Response.Write(" selected ")%> >None</Option><br>
  <option value ="LEV" <%IF QfyOverride = "LEV" THEN Response.Write(" selected ")%> >Qfy By Level</Option><br>
  <option value ="3EV" <%IF QfyOverride = "3EV" THEN Response.Write(" selected ")%> >3rd Event</Option><br>
  <option value ="OVR" <%IF QfyOverride = "OVR" THEN Response.Write(" selected ")%> >By Overall</Option><br>
  <option value ="ALT" <%IF QfyOverride = "ALT" THEN Response.Write(" selected ")%> >Alt Div Qfy</Option><br>
  <option value ="PLC" <%IF QfyOverride = "PLC" THEN Response.Write(" selected ")%> >By Placement</Option><br>
  <option value ="OTH" <%IF QfyOverride = "OTH" THEN Response.Write(" selected ")%> >Other Qfy</Option><br>
  <option value ="DNS" <%IF QfyOverride = "DNS" THEN Response.Write(" selected ")%> >Scratch</Option><br>
</select><%

END SUB



' -------------------------------------------------------------------------
  SUB LoadAWSARegionDrop
' -------------------------------------------------------------------------

%>
<select name='region'>
	<option value ='All'<%IF RegionSelected = "All" THEN Response.Write(" selected ")%>>All </Option><br>
	<option value ='1'<%IF RegionSelected = "1" THEN Response.Write(" selected ")%>>S. Central</Option><br>
	<option value ='2'<%IF RegionSelected = "2" THEN Response.Write(" selected ")%>>MidWest</Option><br>
	<option value ='3'<%IF RegionSelected = "3" THEN Response.Write(" selected ")%>>West</Option><br>
	<option value ='4'<%IF RegionSelected = "4" THEN Response.Write(" selected ")%>>South</Option><br>
	<option value ='5'<%IF RegionSelected = "5" THEN Response.Write(" selected ")%>>East</Option><br>
</select><%

END SUB






' -------------------------------------------------------------------------
  SUB RatingLevelDropBuild
' -------------------------------------------------------------------------

%>
<select name='RatingLevel'>
	<option value =''<%IF RatingLevel = "" THEN Response.Write(" selected ")%>>Select</Option><br>
	<option value ='9'<%IF RatingLevel = "9" THEN Response.Write(" selected ")%>>Level 9</Option><br>
	<option value ='8'<%IF RatingLevel = "8" THEN Response.Write(" selected ")%>>Level 8</Option><br>
	<option value ='7'<%IF RatingLevel = "7" THEN Response.Write(" selected ")%>>Level 7</Option><br>
	<option value ='6'<%IF RatingLevel = "6" THEN Response.Write(" selected ")%>>Level 6</Option><br>
	<option value ='5'<%IF RatingLevel = "5" THEN Response.Write(" selected ")%>>Level 5</Option><br>
	<option value ='4'<%IF RatingLevel = "4" THEN Response.Write(" selected ")%>>Level 4</Option><br>
</select><%

END SUB




' -------------------------------------------------------------------------
  SUB LoadNCWSARegionDrop
' -------------------------------------------------------------------------


%>
<select name="Region" <% =RegionSelected %> >
  <option value ="" <%IF RegionSelected = "" THEN Response.Write(" selected ")%> >None</Option><br>
  <option value ="E" <%IF RegionSelected = "E" THEN Response.Write(" selected ")%> >E</Option><br>
  <option value ="C" <%IF RegionSelected = "C" THEN Response.Write(" selected ")%> >S</Option><br>
  <option value ="M" <%IF RegionSelected = "M" THEN Response.Write(" selected ")%> >M</Option><br>
  <option value ="W" <%IF RegionSelected = "W" THEN Response.Write(" selected ")%> >W</Option><br>
</select><%

END SUB




' -------------------------------------------------------------------------
  SUB LoadAWSAEvents
' -------------------------------------------------------------------------

%>
<select name="EventSelected" style="width:9em">
  <option value ="" <%IF EventSelected = "" THEN Response.Write(" selected ")%> >None</Option><br>
  <option value ="S" <%IF EventSelected = "S" THEN Response.Write(" selected ")%> >Slalom</Option><br>
  <option value ="T" <%IF EventSelected = "T" THEN Response.Write(" selected ")%> >Tricks</Option><br>
  <option value ="J" <%IF EventSelected = "J" THEN Response.Write(" selected ")%> >Jump</Option><br>
</select><%

END SUB



' -------------------------------------------------------------------------
  SUB LoadAWSAEvents_AndAll
' -------------------------------------------------------------------------

%>
<select name="EventSelected" style="width:9em">
  <option value ="ALL" <%IF EventSelected = "ALL" THEN Response.Write(" selected ")%> >All</Option><br>
  <option value ="S" <%IF EventSelected = "S" THEN Response.Write(" selected ")%> >Slalom</Option><br>
  <option value ="T" <%IF EventSelected = "T" THEN Response.Write(" selected ")%> >Tricks</Option><br>
  <option value ="J" <%IF EventSelected = "J" THEN Response.Write(" selected ")%> >Jump</Option><br>
</select><%

END SUB






' -------------------------------------------------------------------------
  SUB LoadGREvents
' -------------------------------------------------------------------------


%>
<select name="EventSelected" style="width:9em">
  <option value ="" <%IF EventSelected = "" THEN Response.Write(" selected ")%> >None</Option><br>
  <option value ="S" <%IF EventSelected = "S" THEN Response.Write(" selected ")%> >Slalom</Option><br>
  <option value ="T" <%IF EventSelected = "T" THEN Response.Write(" selected ")%> >Tricks</Option><br>
  <% 
  t=1
  IF t=2 THEN %>
  <option value ="J" <%IF EventSelected = "J" THEN Response.Write(" selected ")%> >Jump</Option><br>
  <option value ="WB" <%IF EventSelected = "WB" THEN Response.Write(" selected ")%> >Wakeboard</Option><br>
  <option value ="WS" <%IF EventSelected = "WS" THEN Response.Write(" selected ")%> >Wake Skate</Option><br>
  <option value ="WU" <%IF EventSelected = "WU" THEN Response.Write(" selected ")%> >Wake Surf</Option><br>
  <option value ="KB" <%IF EventSelected = "KB" THEN Response.Write(" selected ")%> >Kneeboard Slalom</Option><br>
  <option value ="KT" <%IF EventSelected = "KT" THEN Response.Write(" selected ")%> >Kneeboard Tricks</Option><br>
  <option value ="HY" <%IF EventSelected = "HY" THEN Response.Write(" selected ")%> >Hydrofoil</Option><br>
  <%
  END IF
  %>  

</select><%

END SUB


' -------------------------------------------------------------------------
  SUB LoadGRSkillPulldown (fSkillName, SkillSelected, SkillStatus)
' -------------------------------------------------------------------------


%>
<select name="<%=fSkillName%>" <% =SkillStatus %> style="width:7em">
  <option value ="" <%IF fSkillName = "" THEN Response.Write(" selected ")%> >Select</Option><br><%
  IF sTEvent(EvtNo)="WB" OR sTEvent(EvtNo)="WS"  OR sTEvent(EvtNo)="WU" THEN %>	
	  <option value ="1" <%IF SkillSelected = "1" THEN Response.Write(" selected ")%> >Novice</Option><br>
	  <option value ="2" <%IF SkillSelected = "3" THEN Response.Write(" selected ")%> >Intermed</Option><br>
	  <option value ="3" <%IF SkillSelected = "3" THEN Response.Write(" selected ")%> >Advanced</Option><br>
	  <option value ="4" <%IF SkillSelected = "4" THEN Response.Write(" selected ")%> >Expert</Option><br>
	  <option value ="5" <%IF SkillSelected = "5" THEN Response.Write(" selected ")%> >Outlaw</Option><br><%
  ELSEIF sTEvent(EvtNo)="S" OR sTEvent(EvtNo)="T" THEN %>	
	  <option value ="3" <%IF SkillSelected = "3" THEN Response.Write(" selected ")%> >Challenger</Option><br>
	  <option value ="4" <%IF SkillSelected = "4" THEN Response.Write(" selected ")%> >Competitor</Option><br>
	  <option value ="5" <%IF SkillSelected = "5" THEN Response.Write(" selected ")%> >Outlaw</Option><br><%
  END IF %>	

</select><%

END SUB







' -----------------------
   SUB LoadSkiYearDropdown
' -----------------------

sSQL = "SELECT * FROM "&SkiYearTableName&" ORDER BY SkiYearID DESC"
SET rsSY=Server.CreateObject("ADODB.recordset")
rsSY.open sSQL, SConnectionToTRATable

'response.write("<br>IN Tools -Ski Year Selected ="&SkiYearSelected)

%><SELECT name='SkiYearSelected' style="width:10em"><%
  DO WHILE not rsSY.eof 
	%><option value = "<%=rsSY("SkiYearID")%>" <%IF rsSY("SkiYearID") = CINT(SkiYearSelected) THEN response.write(" selected ")%>><%=rsSY("SkiYearName")%></option><%
	rsSY.movenext
  LOOP %>
</SELECT><%

rsSY.close

END SUB






' -------------------------
   Function SkillDecode(str)
' -------------------------

' --- This function decodes the Skill depending on the Event
SkillDecode=""
IF sTEvent(EvtNo)="WB" OR sTEvent(EvtNo)="WS" OR sTEvent(EvtNo)="WU" THEN	
  SELECT CASE str
    CASE "1"	
	SkillDecode="Novice"
    CASE "2"	
	SkillDecode="Inter"
    CASE "3"	
	SkillDecode="Advcd"
    CASE "4"	
	SkillDecode="Expert"
    CASE "5"	
	SkillDecode="Outlaw"
  END SELECT
END IF

' --- AWSA Events ---
IF sTEvent(EvtNo)="S" OR sTEvent(EvtNo)="T" THEN	
  SELECT CASE str
    CASE "3"	
	SkillDecode="Chal"
    CASE "4"	
	SkillDecode="Comp"
    CASE "5"	
	SkillDecode="Outlaw"
  END SELECT
END IF



End Function





' -----------------------
  SUB NCWRegionDropBuild
' -----------------------


' ------------   Builds Ski Year Drop Down list ----------------- %>

<SELECT name='Region'>
<%

ChooseSQL("SELECT DISTINCT NCWRegion FROM "&TeamTableName)

  response.write("<option value ='ALL'")
  IF sRegionSelected = "ALL" THEN response.write(" SELECTed")
  response.write(">ALL</option><br>")

rs.movefirst
DO WHILE not rs.eof
  response.write("<option value =""" & rs("NCWRegion") & """")

  IF trim(rs("NCWRegion")) = sRegionSelected THEN
    response.write(" SELECTed")
  END IF

  response.write(">")
  response.write(rs("NCWRegion"))
  response.write("</option><br>")
  rs.movenext
LOOP

%>
</SELECT><%

END SUB





'-----------------------------------------------------
 SUB LoadBoatPulldown (BoatField, BoatCode, BoatStatus)
'------------------------------------------------------



BoatList = "Ski Nautique, Malibu, Mastercraft, Centurion, Undetermined"
BoatCodeList ="CC, MA, MC, CN, UK"

BoatArray = Split(BoatList,",")  
BoatCodeArray = Split(BoatCodeList,",")  

%>
<select id="<%=BoatField%>" name="<%=BoatField%>" <%=BoatStatus %> style="width:8em;"><%
  response.write("<option value = """" SELECTED>Select Boat</option>")
  FOR kvar = 0 TO UBOUND(BoatArray)

    IF TRIM(BoatCode) = TRIM(BoatCodeArray(kvar)) THEN
				response.write("<option value = """&BoatCodeArray(kvar)&""" SELECTED>"&BoatArray(kvar)&"</option>")
    ELSE
				response.write("<option value = """&BoatCodeArray(kvar)&""">"&BoatArray(kvar)&"</option>")
    END IF
  NEXT  %>
</select>
<%


END SUB



' ------------------------------------------------------------------------------------------------------------
  SUB LoadValuePulldown (PulldownName, CurrentValue, MinValue, MaxValue, StepValue, PulldownStatus, IncludeNA)
' ------------------------------------------------------------------------------------------------------------

Dim iCounter
CurrentValue = Cint(CurrentValue)

%>
<select name="<%=PulldownName%>" style="width:4em;" <%=PulldownStatus%>>
<%

IF IncludeNA="true" THEN response.write("<option value = 0 >NA</option>")

FOR iCounter = MinValue TO MaxValue STEP StepValue
		IF iCounter = CurrentValue THEN
				response.write("<option value = """&iCounter&""" SELECTED>"&iCounter&"</option>")
		ELSE
				response.write("<option value = """&iCounter&""">"&iCounter&"</option>")
		END IF
NEXT 

%>
</select><%

IF PulldownStatus="disabled" THEN %>
		<input type="hidden" name="<%=PulldownName%>" value="<%=CurrentValue%>"><%
END IF


END SUB




' --------------------------------------------------------------------------------------------------
  SUB LoadRoundSkiedPulldown (PulldownName, CurrentValue, MinValue, MaxValue, StepValue, PulldownStatus, IncludeNA)
' --------------------------------------------------------------------------------------------------




Dim iCounter
CurrentValue = Cint(CurrentValue)

%>
<a title"Select the number of rounds you wish to ski in this event"> 
<select name="<%=PulldownName%>" id="<%=PulldownName%>" style="width:3em;" title="Select the number of rounds you wish to ski in this event" <%=PulldownStatus%>><%
IF IncludeNA="true" THEN response.write("<option value = 0 >NA</option>")

FOR iCounter = MinValue TO MaxValue STEP StepValue
	IF iCounter = CurrentValue THEN
		response.write("<option value = """&iCounter&""" SELECTED>"&iCounter&"</option>")
	ELSE
		response.write("<option value = """&iCounter&""">"&iCounter&"</option>")
	END IF
NEXT %>
</select>
</a><%

IF PulldownStatus="disabled" THEN %>
	<input type="hidden" name="<%=PulldownName%>" value="<%=CurrentValue%>"><%
END IF


jk=1
IF sMemberID="000001151" AND jk=2 THEN
		' response.write("<\div><div>")
		response.write("<br><br><br><br><br><br><br><br><br><br><br>")
		response.write("<br><br>N="&PulldownName)
		response.write("<br><br>CurVal = "&CurrentValue)
		'response.write("<br><br>MinValue = "&MinValue)
		'response.write("<br><br>MaxValue = "&MaxValue)
		'response.write("<br><br>StepValue = "&StepValue)
	 ' response.end 
END IF

END SUB




' ---------------------------------------
   Function GetFieldTypeName(I)
' ---------------------------------------

SELECT CASE i
CASE 0
GetFieldTypeName = "Empty"
CASE 16
GetFieldTypeName = "TinyInt"
CASE 2
GetFieldTypeName = "SmallInt"
CASE 3
GetFieldTypeName = "Integer"
CASE 20
GetFieldTypeName = "BigInt"
CASE 17
GetFieldTypeName = "UnsignedTinyInt"
CASE 18
GetFieldTypeName = "UnsignedSmallInt"
CASE 19
GetFieldTypeName = "UnsignedInt"
CASE 21
GetFieldTypeName = "UnsignedBigInt"
CASE 4
GetFieldTypeName = "Single"
CASE 5
GetFieldTypeName = "Double"
CASE 6
GetFieldTypeName = "Currency"
CASE 14
GetFieldTypeName = "Decimal"
CASE 131
GetFieldTypeName = "Numeric"
CASE 11
GetFieldTypeName = "Boolean"
CASE 10
GetFieldTypeName = "Error"
CASE 132
GetFieldTypeName = "UserDefined"
CASE 12
GetFieldTypeName = "Variant"
CASE 9
GetFieldTypeName = "IDispatch"
CASE 13
GetFieldTypeName = "IUnknown"
CASE 72
GetFieldTypeName = "GUID"
CASE 7
GetFieldTypeName = "Date"
CASE 133
GetFieldTypeName = "DBDate"
CASE 134
GetFieldTypeName = "DBTime"
CASE 135
GetFieldTypeName = "DBTimeStamp"
CASE 8
GetFieldTypeName = "BSTR"
CASE 129
GetFieldTypeName = "Char"
CASE 200
GetFieldTypeName = "VarChar"
CASE 201
GetFieldTypeName = "LongVarChar"
CASE 130
GetFieldTypeName = "WChar"
CASE 202
GetFieldTypeName = "VarWChar"
CASE 203
GetFieldTypeName = "LongVarWChar"
CASE 128
GetFieldTypeName = "Binary"
CASE 204
GetFieldTypeName = "VarBinary"
CASE 205
GetFieldTypeName = "LongVarBinary"
END SELECT

End Function


Function GetFieldTypeCode(sTXT,sLen)

'I am not overly familar with this stuff'
'you may have to edit these values'
SELECT CASE sTXT
CASE "Empty"
GetFieldTypeCode = "Empty"
CASE "TinyInt"
GetFieldTypeCode = "TinyInt"
CASE "SmallInt"
GetFieldTypeCode = "SmallInt"
CASE "Integer"
GetFieldTypeCode = "Integer"
CASE "BigInt"
GetFieldTypeCode = "BigInt"
CASE "UnsignedTinyInt"
GetFieldTypeCode = "UnsignedTinyInt"
CASE "UnsignedSmallInt"
GetFieldTypeCode = "UnsignedSmallInt"
CASE "UnsignedInt"
GetFieldTypeCode = "UnsignedInt"
CASE "UnsignedBigInt"
GetFieldTypeCode = "UnsignedBigInt"
CASE "Single"
GetFieldTypeCode = "Single"
CASE "Double"
GetFieldTypeCode = "Double"
CASE "Currency"
GetFieldTypeCode = "Currency"
CASE "Decimal"
GetFieldTypeCode = "Decimal"
CASE "Numeric"
GetFieldTypeCode = "Numeric"
CASE "Boolean"
GetFieldTypeCode = "Boolean"
CASE "Error"
GetFieldTypeCode = "Error"
CASE "UserDefined"
GetFieldTypeCode = "UserDefined"
CASE "Variant"
GetFieldTypeCode = "Variant"
CASE "IDispatch"
GetFieldTypeCode = "IDispatch"
CASE "IUnknown"
GetFieldTypeCode = "IUnknown"
CASE "GUID"
GetFieldTypeCode = "GUID"
CASE "Date"
GetFieldTypeCode = "Date"
CASE "DBDate"
GetFieldTypeCode = "DBDate"
CASE "DBTime"
GetFieldTypeCode = "DBTime"
CASE "DBTimeStamp"
GetFieldTypeCode = "DBTimeStamp"
CASE "BSTR"
GetFieldTypeCode = "BSTR(" & sLen & ")"
CASE "Char"
GetFieldTypeCode = "Char(" & sLen & ")"
CASE "VarChar"
GetFieldTypeCode = "VarChar(" & sLen & ")"
CASE "LongVarChar"
GetFieldTypeCode = "LongVarChar(" & sLen & ")"
CASE "WChar"
GetFieldTypeCode = "WChar(" & sLen & ")"
CASE "VarWChar"
GetFieldTypeCode = "VarWChar(" & sLen & ")"
CASE "LongVarWChar"
GetFieldTypeCode = "LongVarWChar(" & sLen & ")"
CASE "Binary"
GetFieldTypeCode = "Binary(" & sLen & ")"
CASE "VarBinary"
GetFieldTypeCode = "VarBinary(" & sLen & ")"
CASE "LongVarBinary"
GetFieldTypeCode = "LongVarBinary"
CASE ELSE
GetFieldTypeCode = "IUnknown"
END SELECT

End Function

%>









