<%


Dim sLeagueSelected


' ---------------------------------------------
   SUB BuildLeagueDrop (ActiveOnly, AllorNone)
' ---------------------------------------------



' ------------   Builds Ski Year Drop Down list ----------------- 

sSQL = "SELECT DISTINCT LeagueID, LeagueName"
sSQL = sSQL + " FROM "&LeagueTableName

sSQL = sSQL + " WHERE QualifyTour<>''"
IF Session("SkiYear")=1 THEN
		sSQL = sSQL + " AND SkiYearID=(SELECT SkiYearID FROM "&SkiYearTableName&" WHERE DefaultYear=1)"
END IF
sSQL = sSQL + " AND Status<>'X'"

sSQL = sSQL + " ORDER BY LeagueName DESC" 
set rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable, 3, 3 


' --- Added 10-15-2017 - Sets to current Nationals if nothing selected ---
IF NOT rs.eof AND (TRIM(sLeagueSelected)="" OR LCASE(TRIM(sLeagueSelected))="none") THEN
		DO WHILE not rs.eof
				IF InStr(rs("LeagueID"),"NATL")>0 THEN sLeagueSelected=rs("LeagueID")
				rs.movenext
		LOOP
		rs.movefirst
END IF

' response.write("<br> sLeagueSelected"&sLeagueSelected)

%>
<SELECT name='sLeagueSelected' style="width:16em"><%

  response.write("<option value ='"&AllorNone&"'")
  IF sLeagueSelected = "&AllorNone&" THEN response.write(" SELECTed")
  response.write(">"&AllorNone&"</option><br>")

  IF NOT rs.eof THEN
	rs.movefirst
	DO WHILE not rs.eof
	  response.write(" <option value ="""&rs("LeagueID")&""" ")
	  response.write(" <a title="""&rs("LeagueName")&"""")

	  IF trim(rs("LeagueID")) = sLeagueSelected THEN
	    response.write(" selected")
	  END IF

	  response.write(">")
	  response.write(rs("LeagueName"))
	  response.write("</a></option><br>")
	  rs.movenext
	LOOP
  END IF %>

</SELECT><%

rs.close


END SUB





' ---------------------------------------------
   SUB BuildLeagueDrop_Mobile (ActiveOnly, AllorNone)
' ---------------------------------------------



' ------------   Builds Ski Year Drop Down list ----------------- 

sSQL = "SELECT DISTINCT LeagueID, LeagueName"
sSQL = sSQL + " FROM "&LeagueTableName

'Response.write("SkierYearID="&Session("SkiYear"))

sSQL = sSQL + " WHERE QualifyTour<>''"
IF Session("SkiYear")=1 THEN
	sSQL = sSQL + " AND SkiYearID=(SELECT SkiYearID FROM "&SkiYearTableName&" WHERE DefaultYear=1)"
END IF
sSQL = sSQL + " AND Status<>'X'"


'IF ActiveOnly=true THEN
'	sSQL = sSQL + " AND LEFT(QualifyTour,2)=(SELECT RIGHT(SkiYearName,2) FROM usawsrank.SkiYear WHERE DefaultYear='1') AND QualifyTour<>''"
'END IF

'sSQL = sSQL + " SptsGrpID='"&Session("sSptsGrpID")&"'"
sSQL = sSQL + " ORDER BY LeagueName DESC" 
set rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable, 3, 3 

%>

<SELECT name='sLeagueSelected' style="width:14em; font-size:12pt;"><%

  response.write("<option value ='"&AllorNone&"'")
  IF sLeagueSelected = "&AllorNone&" THEN response.write(" SELECTed")
  response.write(">"&AllorNone&"</option><br>")

  IF NOT rs.eof THEN
	rs.movefirst
	DO WHILE not rs.eof
	  response.write(" <option value ="""&rs("LeagueID")&""" ")
	  response.write(" <a title="""&rs("LeagueName")&"""")

	  IF trim(rs("LeagueID")) = sLeagueSelected THEN
	    response.write(" selected")
	  END IF

	  response.write(">")
	  response.write(rs("LeagueName"))
	  response.write("</a></option><br>")
	  rs.movenext
	LOOP
  END IF %>

</SELECT><%

rs.close


END SUB






' ----------------------------------------------------------------
   SUB BuildTeamType_DropDown (thiswidth,thisfont,onchangeaction)
' ----------------------------------------------------------------

' ------------   Builds TEAM TYPE Drop Down list ----------------- 

sSQL = "SELECT DISTINCT Team_Type_ID, Team_Type_ID_Seq, Team_Type_Description"
sSQL = sSQL + " FROM "&V_TeamTypeTableName
sSQL = sSQL + " ORDER BY Team_Type_ID_Seq"

set rsTT=Server.CreateObject("ADODB.recordset")
rsTT.open sSQL, SConnectionToTRATable, 3, 3 

'response.write("</span></div><div style=""color:white;"">onchangeaction = "&onchangeaction&"</div>")

%>
<SELECT id="TeamTypeIDSelected" name="TeamTypeIDSelected" style="width:<%=thiswidth%>em; font-size:<%=thisfont%>pt;" onchange=<%=onchangeaction%>;>
<option value=0>Select</option><% 

		DO WHILE NOT rsTT.eof 
						IF CStr(rsTT("Team_Type_ID")) = CStr(TeamTypeIDSelected) THEN 
								%><option value = "<%=rsTT("Team_Type_ID")%>" selected><%= rsTT("Team_Type_ID") %> - <%= rsTT("Team_Type_Description") %></option><%
						ELSE
								%><option value = "<%=rsTT("Team_Type_ID")%>"><%= rsTT("Team_Type_ID") %> - <%= rsTT("Team_Type_Description") %></option><%
						END IF	
				rsTT.movenext
  	LOOP 

%></SELECT><%

rsTT.close




END SUB




%>