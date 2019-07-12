<!--#include file="settingsHQ.asp"-->
<!--#include file="Tools_Include.asp"-->
<!--#include file="Tools_Definitions.asp"-->
<!--#include file="Tools_TourDefine.asp"-->
<!--#include file="Tools_Registration16.asp"-->
<%

Dim currentPage, rowCount, i, usgi, amlvl
Dim sScore(5)

Dim sAltScore(5)
Dim sPQ1(5)
Dim sPQ2(5)
Dim sClassArray(5)
'Dim sDiv(5)
'Dim sTRounds(5)
Dim RoundLoop
Dim sMemberID, sTourID, sEvent, sMembAge, sMembSex
Dim TempMemFed, TempMemLName, TempMemFName, TempBirthDate, TempMemGender
Dim DateGood
Dim ErrorCheck
Dim sProceedToAdd 	' 0 = no, 1 = member good, 2 = tour good, 3 = member and tour good.
Dim sPlace, Action

Dim SD_Desc
Dim MainLogo, MainLogoWidth, MainLogoHeight, SD_Heading
Dim ScoreFieldStatus, ScoresExist, ScoresInForm_YorN, SomeScoresExistForThisEvent
Dim LeagueSelected, TeamSelected

' ----------------------------------------------------------------
' --- Reads variables from form that are common to all records ---
' ----------------------------------------------------------------

sMemberID = trim(Request("sMemberID"))
sLastName = trim(request("Last_Name"))
sFirstName = trim(request("First_Name"))
sTourID = trim(Request("Tour_ID"))
sTourName = trim(Request("Tour_Name"))
sTourDate = trim(Request("Tour_Date"))
sPlace = SQLClean(trim(Request("Place")))
ScoreFieldStatus=TRIM(Request("ScoreFieldStatus"))
ScoresExist=TRIM(Request("ScoresExist"))
Action=Request("Action")


sEvent = trim(Request("Event"))
IF sEvent="" OR sEvent<>Session("LastEvent") THEN
	Action="Edit Existing Scores"

	IF sEvent="" THEN sEvent="S"
	Session("LastEvent")=sEvent
END IF





LeagueSelected = TRIM(Request("LeagueDrop"))
IF TRIM(LeagueSelected)<>"" THEN Session("LastLeague")=LeagueSelected

TeamSelected = TRIM(Request("TeamDrop"))
IF TRIM(TeamSelected)<>"" THEN Session("TeamLeague")=TeamSelected




' --- Test settings ---
' sTourID="11S076E"
'sMemberID="600049564"
'sMemberID="100000047"




' ---------------------------------------------------------------------------------
' --- Sets Session
' ---------------------------------------------------------------------------------

sSptsGrpID = TRIM(Request("SD_Desc"))
IF sSptsGrpID="AWS" OR sSptsGrpID="USW" OR sSptsGrpID="AKA" OR sSptsGrpID="NCW" OR sSptsGrpID="HYD" THEN
	Session("sSptsGrpID")=sSptsGrpID

ELSEIF sSptsGrpID="R" OR Session("sSptsGrpID")="" THEN
	Session("sSptsGrpID")=""
	response.redirect("/rankings/defaultHQ.asp")
ELSE
	sSptsGrpID = Session("sSptsGrpID")
END IF




' --- Tests the authority of this person to be in this module ---
' --- Note that revised logic allows AWS users to act for NCW ---
usgi = Session("UserSptsGrpID"): amlvl = Session("adminmenulevel")
IF (sSptsGrpID<>usgi AND (NOT sSptsGrpID="NCW" AND usgi="AWS")) AND amlvl<50 THEN
	response.redirect("/rankings/tools.asp?svar=reject")
END IF




sProceedToAdd = 0

OpenCon
Set rs=Server.CreateObject("ADODB.recordset")
set rsSelectFields = Server.CreateObject("ADODB.recordset")
set rsCheckDups = Server.CreateObject("ADODB.recordset")



WriteIndexPageHeader


SELECT CASE trim(Request("pvar")) 

  CASE "SearchMember" 
    	SearchMember

  CASE "SearchTour"
    	SearchTour

  CASE "" 
    
    If sMemberID <> "" then
      sSQL = "Select top 10 PersonIDwithCheckDigit,FederationCode,LastName,FirstName,City,State,BirthDate,Sex from " & MemberTableName & " where "
      sSQL = sSQL + "PersonIDwithCheckDigit LIKE '%" & SQLClean(sMemberID) & "%'"
      sSQL = sSQL + " and membertypeid <> 2 order by PersonIDWithCheckDigit"
      rs.open sSQL, sConnectionToTRATable, 3, 1
      If Not rs.EOF Then
        If rs.recordcount = 1 Then
          TempMemFed = rs("FederationCode")
          TempMemLName = rs("LastName")
          TempMemFName = rs("FirstName")
          TempBirthDate = rs("BirthDate")
          sMembSex = rs("Sex")
          sMemberID = rs("PersonIDwithCheckDigit")
	  sMembCity = rs("City")
	  sMembState = rs("State")
          sProceedToAdd = 1
        End If
      End If
      rs.Close
    End If


    
	IF sTourID <> "" THEN 

		'--- In SUB tools_registration.asp ---
		DefineTourVariables_New

		' --- Determines EVENTS and EVENT NAMES for this tourament - in tools_include.asp ---
		RegistrationEventsOffered (sTSptsGrpID)

		sProceedToAdd = sProceedToAdd + 2
	END IF


	TourSptsGrpID = Session("sTSptsGrpID")


	SELECT CASE TRIM(sEvent)
		CASE sTEvent(1)
			EvtNo=1
		CASE sTEvent(2)
			EvtNo=2
		CASE sTEvent(3)
			EvtNo=3
		CASE sTEvent(4)
			EvtNo=4
	END SELECT



	' --- Reads all score sets ----
	GetScoresInForm


    IF TRIM(sTDateS)<>"" AND TRIM(sMemberID)<>"" THEN 	
	    sMembAge = AgeAtDate(sTDateS, sMemberID)		' --- Function finds Member Age
    END IF


 
	     
    ' ------------------------------------------------------------------------------	
    ' ---- Display Member and Tournament upper boxes
    ' ------------------------------------------------------------------------------

    BannerTitle = "Add Scores"	
    DisplayPageBanner (BannerTitle)
    %>
    <br>
    <TABLE align=center class="innertable" width="<%=TourTableWidth%>"px>
    <tr>
      <th width=250px align=center><font size="3" color="#FFFFFF"><b>Member</b></font></th>
      <th align=center><font size="3" color="#FFFFFF"><b>Tournament</b></font></th>
    </tr>
    <tr>
      <td width=50%>	<% ' --- TOURNAMENT cell of outer Table --- %>

    <% If (sProceedToAdd <> 1) And (sProceedToAdd <> 3) Then %>
          <br>
	  <form action="/rankings/addscores.asp" method="post">
          <input type="hidden" name="pvar" value="SearchMember">
          <input type="hidden" name="Tour_ID" value="<%=sTourID%>">
          <center>
            <input type="submit" style="width:9em" value="Search">
          </center>
          </form>
    <% Else  %>
    	<TABLE class="noborder" align=center width="100%">
	<TR>
  	  <TD align=right><font size="<%=fontsize2%>">&nbsp;<b>Name:</b>&nbsp;</font></TD>
	  <TD colspan=2 align=left><font size="<%=fontsize2%>" color="<%=TextColor2%>"><%=TempMemFName%>&nbsp;<%=TempMemLName%></font></TD>
	</TR>
	<TR>
	  <TD align=right><font size="<%=fontsize2%>">&nbsp;<b>ID:</b>&nbsp;</font></TD>
	  <TD colspan=2 align=left><font size="<%=fontsize2%>" color="<%=TextColor2%>"><%=sMemberID%></font></TD>

	</TR>
	<TR>
	  <TD align=right><font size="<%=fontsize2%>">&nbsp;<b>Age/Gend:</b>&nbsp;</font></TD>
	  <TD align=left><font size="<%=fontsize2%>" color="<%=TextColor2%>"><%=sMembAge%>/<%=sMembSex%></font></TD>
          <Form action="/rankings/addscores.asp" method="post">
	    <TD width=170px rowspan=3 align=center>
            <input Type="hidden" name="pvar" Value="SearchMember">
            <input type="hidden" name="Tour_ID" value="<%=sTourID%>">
            <input Type="submit" style="width:9em" Value="Change">
          </TD>
          </Form>
	</TR>
	<TR>
	  <TD align=right><font size="<%=fontsize2%>">&nbsp;<b>City/ST:</b>&nbsp;</font></TD>
	  <TD align=left><font size="<%=fontsize2%>" color="<%=TextColor2%>"><%=sMembCity%>,&nbsp;<%=sMembState%></font></TD>
	</TR>
	</TABLE>
	<%
      
    END IF %>
    </td>

    <td width=50%><% ' --- TOURNAMENT cell of outer Table --- 

    IF (sProceedToAdd <> 2) And (sProceedToAdd <> 3) Then %>
          <br>
	  <form action="/rankings/addscores.asp" method="post">
          <input type="hidden" name="pvar" value="SearchTour">
          <input type="hidden" name="sMemberID" value="<%=sMemberID%>">
          <center>
            <input type="submit" style="width:9em" value="Search">
          </center>
          </form><% 
    ELSE  %>
    	<TABLE  class="noborder" align=center width="100%">	
	  <TR>
	    <TD width=70px align=right><font size="<%=fontsize2%>">&nbsp;<b>Name:</b>&nbsp;</font></TD>
	    <TD colspan=2 width=220px align=left><font size="<%=fontsize2%>" color="<%=TextColor2%>"><%=sTourName%></font></TD>
	  </TR>
	  <TR>
	    <TD align=right><font size="<%=fontsize2%>">&nbsp;<b>TourID:</b>&nbsp;</font></TD>
	    <TD colspan=2 align=left><font size="<%=fontsize2%>" color="<%=TextColor2%>"><%=sTourID%></font></TD>
	  </TR>
          <TR>
	    <TD align=right><font size="<%=fontsize2%>">&nbsp;<b>City/ST:</b>&nbsp;</font></TD>
	    <TD align=left><font size="<%=fontsize2%>" color="<%=TextColor2%>"><%=sTourCity%>,&nbsp;<%=sTState%></font></TD>
        <Form action="/rankings/addscores.asp" method="post">
	    <TD width=170px rowspan=2 align=center>
          	<input Type="hidden" name="pvar" Value="SearchTour">
          	<input type="hidden" name="sMemberID" value="<%=sMemberID%>">
                <input Type="submit" style="width:9em" Value="Change">
     	    </TD>
        </Form>

	  </TR>
          <TR>
	    <TD align=right><font size="<%=fontsize2%>">&nbsp;<b>Dates:</b>&nbsp;</font></TD><%
	    IF sTDateS=sTDateE THEN %>	
	    	<TD align=left><font size="<%=fontsize2%>" color="<%=TextColor2%>"><%=sTDateS%></font></TD><%
	    ELSE  %>	
	    	<TD align=left><font size="<%=fontsize2%>" color="<%=TextColor2%>"><%=sTDateS%>&nbsp;- &nbsp;<%=sTDateE%></font></TD><%
	    END IF %>
	  </TR>
	</TABLE><%

    END IF  %>
    </td>
   </tr>
 </table>
    <br>
    <%


'Response.write("<br>Line 303 - Action="&Action)

    ' -----------------------------------------------    
    ' --- If Member and Tournament are both valid ---	
    ' -----------------------------------------------
    IF sProceedToAdd = 3 THEN

	' --- Reads any scores from form ---
	GetScoresInForm

	' --- Checks if any scores exist in the table for this event ---
	CheckForAnyScoresInThisEvent

'response.write("<br>Action="&Action)
'response.write("<br>rsCheckDups.eof=")
'response.write("<br>ScoresInForm_YorN="&ScoresInForm_YorN)



	IF Action="Confirm Scores" THEN 
		'response.write("Pos 1B - Confirm or Update")
		ConfirmScores

	ELSEIF Action="Save Scores" THEN
		UpdateScoreTable
		' --- Redisplays the add scores function ---

		ScoreFieldStatus="disabled"
		IF ScoresInForm_YorN=false THEN ScoreFieldStatus="enabled"

		AddEditScores

	ELSEIF Action="Edit Existing Scores" THEN

		ScoreFieldStatus="enabled"
		IF SomeScoresExistForThisEvent="Y" THEN ScoreFieldStatus="disabled"

        	AddEditScores

	ELSEIF Action="Cancel" OR Action="" THEN
		ScoreFieldStatus="enabled"
        	AddEditScores

	ELSE
		ScoreFieldStatus="disabled"
		ScoresExist="overwrite"

        	AddEditScores

		'response.write("Pos 1A - AddEditScores only")
		'response.end
	END IF		' --- Check for Scores ---

    END IF	' --- sProceedToAdd ---

END SELECT	' --- Select based on pvar ---



'response.write("Session(sTEventRounds)="&Session("sTEventRounds"))





CloseCon
set rsSelectFields = Nothing
set rsCheckDups = Nothing
Set rs = Nothing

WriteIndexPageFooter


' ---------------------------------------------------------------------------------------------
' -----------------   BOTTOM OF MAIN CODE 	-----------------------------------------------
' ---------------------------------------------------------------------------------------------




' ----------------------
  SUB GetScoresInForm
' ----------------------

ScoresInForm_YorN=false

' --- Loops thru all rounds to collect the form settings ---
FOR RoundLoop = 1 TO sTRounds(Session("TotEv"))
	sScore(RoundLoop) = trim(Request("Score_"&RoundLoop))
	sDiv(RoundLoop) = trim(Request("Div_"&RoundLoop))
	sPQ1(RoundLoop) = trim(Request("PQ1_"&RoundLoop))
	sPQ2(RoundLoop) = trim(Request("PQ2_"&RoundLoop))
	sAltScore(RoundLoop) = trim(Request("AltScore_"&RoundLoop))
	sClassArray(RoundLoop) = trim(Request("Class_"&RoundLoop))

	IF TRIM(sClassArray(RoundLoop))<>"" THEN Session("LastClass")=sClassArray(RoundLoop)

	IF Roundloop>9 THEN response.end
	IF sScore(RoundLoop)<>"" THEN ScoresInForm_YorN=true
NEXT



END SUB


' ---------------------------------
  SUB CheckForAnyScoresInThisEvent
' ---------------------------------

	SomeScoresExistForThisEvent="N"

	' -----------------------------------------------------------------------------
	' --- Checks for existence of scores for this Member/Tournament/Event/Round ---
	' -----------------------------------------------------------------------------
	SET rsCheckDups = Server.CreateObject("ADODB.recordset")
        IF sSPtsGrpID = "AWS" OR sSPtsGrpID = "NCW" THEN
		sSQL = "SELECT * from " & RawScoresTableName
	ELSE
		sSQL = "SELECT * from " & RawScoresOtherTableName
	END IF
        sSQL = sSQL + " WHERE MemberID = '" & sMemberID & "' AND"
        sSQL = sSQL + " TourID = '" & sTourID & "' AND"
        sSQL = sSQL + " Event = '" & sEvent & "' AND"
        sSQL = sSQL + " [Round] = '" & RoundLoop & "'"
        rsCheckDups.open sSQL, sConnectionToTRATable, 3,3

	IF NOT rsCheckDups.eof THEN
		SomeScoresExistForThisEvent="Y"
	END IF 

	rsCheckDups.close
END SUB




' -------------------
   Sub AddEditScores
' -------------------

' -----------------------------------------------------------------------
' Once we have a member and a tour, this procedure will 
' actually take care of receiving the new score data
' -----------------------------------------------------------------------

' -----------------------------------------------------------------------
' First we look to see if scores already exist. 
' If they do, we prompt the user for a quick-link to the edit function.
' -----------------------------------------------------------------------


' -------------------------------------------------------------------------------------------------
' --- Checks for existence of ANY scores for this Member/Tournament to set TRIM(sEvent) if none selected
' -------------------------------------------------------------------------------------------------


IF sSPtsGrpID = "AWS" OR sSPtsGrpID = "NCW" THEN
	sSQL = "SELECT DISTINCT MemberID, TourID, Event FROM " & RawScoresTableName
ELSE
	sSQL = "SELECT DISTINCT MemberID, TourID, Event FROM " & RawScoresOtherTableName
END IF
sSQL = sSQL + " WHERE MemberID = '" & sMemberID & "' AND"
sSQL = sSQL + " TourID = '" & sTourID & "'"

SET rsCheckDups = Server.CreateObject("ADODB.recordset")
rsCheckDups.open sSQL, sConnectionToTRATable, 3,3

IF sEvent="" AND (NOT rsCheckDups.EOF) THEN sEvent=rsCheckDups("Event")

rsCheckDups.close

'response.write("Below Checkdupes")



' ------------------------------------------------------------------------------------
' --- Finds the top 10 scores for the current Member in sEvent at this Tournament ---
' ------------------------------------------------------------------------------------

IF sSPtsGrpID = "AWS" OR sSPtsGrpID = "NCW" THEN
	sSQL = "SELECT TOP 10 * FROM " & RawScoresTableName
ELSE
	sSQL = "SELECT TOP 10 * FROM " & RawScoresOtherTableName
END IF

sSQL = sSQL + " WHERE MemberID = '"&sMemberID&"' and left(TourID,6) = '"&left(sTourID,6)&"'"
sSQL = sSQL + " AND Event='"&sEvent&"'"
rs.open sSQL, SConnectionToTRATable, 3, 1

'response.write("<br>In AddEdit - sSQL="&sSQL)
'response.end






' --- Found existing scores so initially disable the Submit Scores button and display Edit Scores button ---
' --- Set Place to the first place encountered ---
IF NOT rs.EOF THEN
  
    IF ScoreFieldStatus="" THEN ScoreFieldStatus="disabled"

    rs.movefirst
    DO WHILE NOT rs.EOF   

	' --- If no place has been assigned and the current record has a place value, then set sPlace ---
	IF sPlace="" AND TRIM(rs("Place"))<>"" THEN
		sPlace=rs("Place")
	END IF
  	rs.movenext
    LOOP

END IF
rs.close

'response.write("<br>Top of form for SUBMIT SCORES")


' -------------------------------------------
' --- Top of form for SUBMIT SCORES ---
' --------------------------------------------
%>
<form action="/rankings/addscores.asp" method="post">

  <input type="hidden" name="sMemberID" value="<%=sMemberID%>">
  <input type="hidden" name="Tour_ID" value="<%=sTourID%>">
  <input type="hidden" name="TeamDrop" value="<%=TeamSelected%>">
  <input type="hidden" name="LeagueDrop" value="<%=LeagueSelected%>">

<TABLE class="innertable" align=center width="<%=TourTableWidth%>"px >
  <tr>
    <TH width=250px align=center><b><font size="<%=fontsize1%>" color="#FFFFFF">Event</font></b></TH>
    <TH align=center><b><font size="<%=fontsize1%>" color="#FFFFFF">Place</font></b></TH><%
    IF sSptsGrpID="NCW" THEN %>
	    <TH width=200px align=center><b><font size="<%=fontsize1%>" color="#FFFFFF">Team</font></b></TH><%
    END IF %>
    <TH align=center><b><font size="<%=fontsize1%>" color="#FFFFFF">League/Series</font></b></TH>
  </tr>
  <tr>
    <td align=center><%
	 IF ScoresExist="overwrite" THEN 
		EventStatus="disabled"
		%><input type="hidden" name="Event" value="<%=sEvent%>"><%
	 END IF
	 BuildEventDrop EventStatus %>

    </td>
    <td align=center><input type=text name="Place" size=5 value="<%=sPlace%>" <%=ScoreFieldStatus%>></td><%
    IF sSptsGrpID="NCW" THEN %>
	    <td align=center><%
		' --- SUB located in tools_definitions.asp ---
		IF Session("SptsGrpID")<>"NCW" THEN TeamStatus="disabled"
		LoadTeam TeamSelected, TeamStatus  %>
	    </Td><%
    END IF %>
    <td align=center><%
	' --- SUB located in tools_definitions.asp ---
	IF Session("SptsGrpID")="NCW" THEN LeagueStatus="disabled"
	LoadLeague LeagueSelected, LeagueStatus %>
    </td>	
  </tr>
</table>


<br>
<TABLE align=center class="innertable" width="<%=TourTableWidth%>"px>
  <tr><%
	DisplayScoreLineHeader %>
  </tr>
<%


    sMembAge = AgeAtDate(sTDateS, sMemberID)		' Function finds Member Age


     FOR RoundLoop = 1 TO sTRounds(Session("TotEv"))
	' -----------------------------------------------------------------------------
	' --- Checks for existence of scores for this Member/Tournament/Event/Round ---
	' -----------------------------------------------------------------------------
	SET rsCheckDups = Server.CreateObject("ADODB.recordset")
        IF sSPtsGrpID = "AWS" OR sSPtsGrpID = "NCW" THEN
		sSQL = "SELECT * from " & RawScoresTableName
	ELSE
		sSQL = "SELECT * from " & RawScoresOtherTableName
	END IF
        sSQL = sSQL + " WHERE MemberID = '" & sMemberID & "' AND"
        sSQL = sSQL + " TourID = '" & sTourID & "' AND"
        sSQL = sSQL + " Event = '" & sEvent & "' AND"
        sSQL = sSQL + " [Round] = '" & RoundLoop & "'"
        rsCheckDups.open sSQL, sConnectionToTRATable, 3,3




	IF NOT rsCheckDups.EOF THEN	' --- Found scores for this round ---

		DivSelected=rsCheckDups("Div")		
		EventSelected=sEvent
		DivDropName="Div_"&RoundLoop
		DivDropStatus=ScoreFieldStatus
		
		%>
		<tr>
		  <td align=center width=130><b><font size="<%=fontsize1%>">EDIT - Round <%=RoundLoop%>&nbsp;</font></b></td>
		  <td align=center width=150><%

		      ' --- SUB located in tools_include.asp
		      LoadDivDropWithAgeGender_IntlIncluded DivSelected, EventSelected, DivDropName, DivDropStatus %>
		  </td>          
		  <td align=center>
		    <input type=text size=8 MaxLength=7 <%=ScoreFieldStatus%> name="Score_<%=RoundLoop%>" value="<%=rsCheckDups("Score")%>"></td><%

		    IF sEvent="S" OR sEvent="J" THEN %>
			  <td align=center><font size="<%=fontsize1%>"><%=rsCheckDups("AltScore")%></font></td><%
		    ELSE %>	
			  <td align=center>&nbsp;</td><%
		    END IF  

		 
		    IF sEvent="J" THEN  %>
		  	<td align=center>
			  <font size="<%=fontsize1%>"><%
				RampFieldName="PQ1_"&RoundLoop
				sRampStatus="disabled"
				' --- SUB in Tools_Definitions.asp --
				LoadRampPulldown_Nov1_2010 rsCheckDups("Div"), RampFieldName, rsCheckDups("Perf_Qual1"), sRampStatus  %>
			  </font>
			</td><%
		    ELSEIF sEvent="S" THEN %>
			<td align=center><font size="<%=fontsize1%>"><%=rsCheckDups("Perf_Qual1")%></font></td><%
		    ELSE %>
			  <td align=center>&nbsp;</td><%
		    END IF 
		  
		  IF sEvent<>"T" THEN %>	
			  <td align=center><font size="<%=fontsize1%>"><%=rsCheckDups("Perf_Qual2")%></font></td><%
		  ELSE %>
			  <td align=center>&nbsp;</td><%
		  END IF %>

		  <td align=center><%
			' --- Define parameters and run SUB in tools_definitions.asp ---
			sClassArray(RoundLoop)=rsCheckDups("Class")
			sClassFieldName="Class_"&RoundLoop

'response.write("<br>Pos1 - sClassFieldName="&sClassFieldName)
'response.write("<br>sClassArray(RoundLoop)="&sClassArray(RoundLoop))
'response.write("<br>sClassArray(RoundLoop)="&sClassArray(RoundLoop))


			BuildClassDrop_NEW sClassFieldName, sClassArray(RoundLoop), TRIM(sEvent), ScoreFieldStatus 
			       
			' --- Use to capture existing hidden values to fill inputs in confirm rather than calculating new ones --- %>
			<input type="hidden" name="AltScore_<%=RoundLoop%>" value="<%=rsCheckDups("AltScore")%>"><%
			IF sEvent<>"J" THEN %>
				<input type="hidden" name="PQ1_<%=RoundLoop%>" value="<%=rsCheckDups("Perf_Qual1")%>"><%
			END IF %>
			<input type="hidden" name="PQ2_<%=RoundLoop%>" value="<%=rsCheckDups("Perf_Qual2")%>">
		</td>
		</tr><%
	ELSE 

		DivSelected=Session("sDiv")		

		'DivSelected=""		
		EventSelected=sEvent
		DivDropName="Div_"&RoundLoop
		DivDropStatus=ScoreFieldStatus

		%>
		<tr>
		<td align=center width=130><b><font size="<%=fontsize1%>" color=red>NEW</font><font size="<%=fontsize1%>" color="black"> Round <%=RoundLoop%>&nbsp;</font></b></td>
		  <td align=center width=150><%
		      LoadDivDropWithAgeGender_IntlIncluded DivSelected, EventSelected, DivDropName, DivDropStatus %>
		  </td>          
		  <td align=center width=130>
		    <input type=text size=8 MaxLength=7 name="Score_<%=RoundLoop%>"></td>

		    <td align=center>--</td>
		    <td align=center>--</td>
		    <td align=center>--</td>

		  <td align=center><%
		     	' --- Define parameters and run sub in tools_define.asp ---
			sClassArray(RoundLoop)=""
			sClassFieldName="Class_"&RoundLoop

'response.write("<br>Pos2 - sClassFieldName="&sClassFieldName)
'response.write("<br>sClassArray(RoundLoop)="&sClassArray(RoundLoop))

			BuildClassDrop_NEW sClassFieldName, sClassArray(RoundLoop), TRIM(sEvent), ScoreFieldStatus %>
         	  </td>
		</tr><%
	END IF	

	' --- Sets the Session class to the LAST class
	Session("Class")=sClassArray(RoundLoop)

    NEXT  
%>
</table>
<center><small><small>* Leave score blank if skier did not participate in that round.</small></small></center>
<br>

<TABLE align=center width=60%><%
  IF ScoreFieldStatus="disabled" THEN 

    tWarning="N"
    IF tWarning="Y" THEN	
    %>
     <tr>
       <td colspan=2 align=center>
	   <font color="red"><%response.write("Event Scores already exist for this Member")%></font>
	   <br><br>
       </td>
     </tr><%
     END IF

  END IF %>
  <tr>
   <td align=center>
	<input type=submit <%=ScoreFieldStatus%> style="width:12em" name="Action" value="Confirm Scores">
   </td><%	

' --- Found existing scores so initially disable the Submit Scores button and display Edit Scores button ---
IF ScoreFieldStatus="disabled" THEN
	%>
     <td align=center> 
	<input type=submit style="width:12em" name="Action" value="Edit Existing Scores">
     </td><%
ELSE  %>
    <td align=center>
	<input type=submit style="width:12em" name="Action" value="Cancel">
    </td>
<%
END IF  
    %>
  </tr>
</TABLE>
</form>

<%

END SUB


' ----------------------------
  SUB DisplayScoreLineHeader
' ----------------------------

%>
    <th align=center><font color="#FFFFFF" size="<%=fontsize1%>">Round</font></th>
    <th align=center><font color="#FFFFFF" size="<%=fontsize1%>">Division</font></th>
    <th align=center><font color="#FFFFFF" size="<%=fontsize1%>">Score</font></th><%

     SELECT CASE TRIM(sEvent)
	CASE "S"  %>	
	    <th align=center><font color="#FFFFFF" size="<%=fontsize1%>">Buoys</font></th>
	    <th align=center><font color="#FFFFFF" size="<%=fontsize1%>">End Line</font></th>
	    <th><font color="#FFFFFF" size="<%=fontsize1%>">Speed</font></th><%
	CASE "J" %>	
	    <th align=center><font color="#FFFFFF" size="<%=fontsize1%>">Metric</font></th>
	    <th align=center><font color="#FFFFFF" size="<%=fontsize1%>">Ramp</font></th>
	    <th align=center><font color="#FFFFFF" size="<%=fontsize1%>">Boat Speed</font></th><%
	CASE "T" %>
	    <th align=center>&nbsp;</th>
	    <th align=center>&nbsp;</th>
	    <th align=center>&nbsp;</th><%
	CASE "WB" %>	
	    <th align=center>&nbsp;</th>
	    <th align=center>&nbsp;</th>
	    <th align=center>&nbsp;</th><%


     END SELECT	 %>
    <th align=center><font color="#FFFFFF"  size="<%=fontsize1%>">Class</font></th>
<%


END SUB



' ------------------
   SUB ConfirmScores
' ------------------



' ------------------------------------------------------------------------
' Display scores
' Determine and display boxes to edit PQ1, PQ2, and AltScore
' Display confirmation for save
' ------------------------------------------------------------------------

FOR RoundLoop = 1 TO sTRounds(Session("TotEv"))
  	sScore(RoundLoop) = trim(Request("Score_"&RoundLoop))
	'sDiv(RoundLoop)="Div_"&RoundLoop
  	IF sScore(RoundLoop) = "" THEN sScore(RoundLoop) = null

NEXT

'response.write("<br><br>ConfirmScoreDetails - sDiv(1)="&sDiv(1))


FOR RoundLoop = 1 TO sTRounds(Session("TotEv"))


   SET rsSelectFields = Server.CreateObject("ADODB.recordset")
   sSQL = "Select top 1 * from " & DivisionsTableName & " WHERE upper(div)='" & sqlclean(ucase(sDiv(RoundLoop))) &"'"
   rsSelectFields.open sSQL, SConnectionToTRATable


   IF rsSelectFields.eof THEN
	response.write("<font color=red>REPORT ERROR:  DIVISION "&ucase(sDiv(RoundLoop))&" not in the Division Table</font>")
   ELSE



	' --- The event is SLALOM and nothing is in the AltScore field ---		 
	IF sScore(RoundLoop)<>"" AND sEvent = "S" AND sSptsGrpID="AWS" AND TRIM(sAltScore(RoundLoop))="" THEN
        	sAltScore(RoundLoop)=sScore(RoundLoop)-INT(sScore(RoundLoop)/6)*6
        	IF cdbl(sScore(RoundLoop))>0 and sAltScore(RoundLoop)=0 THEN
 			sAltScore(RoundLoop)=6.00
		END IF

        
        	' --- Score is < or = number of buoys at completion of max boat speed ---
        	IF cdbl(sScore(RoundLoop)) <= cdbl(rsSelectFields("BOUY_MAX")) THEN
			' --- Set the line length to LINE_MAX for this division ---
        		sPQ1(RoundLoop) = rsSelectFields("LINE_MAX")


			' --- If the Min Starting Speed is > 38 then it is metric speed --- 
        		'IF cdbl(rsSelectFields("MIN_S1"))>38 then
        			' --- OLD pre-ZBS 
				' sPQ2(RoundLoop) = rsSelectFields("MIN_S1") + INT(cdbl(sScore(RoundLoop))/6)*3
			' --- Only the following line applies with ZBS ---
			sPQ2(RoundLoop) = 25+INT(cdbl(sScore(RoundLoop))/6)*3
        		'ELSE
				' --- Kept the same because this could only happen in OLD English speeds ---
        			'sPQ2(RoundLoop) = rsSelectFields("MIN_S1") + INT(sScore(RoundLoop)/6)*2
        		'END IF            
        		
        		' --- If extra buoys are 0, make 6 to show 6 at previous speed ---
        		sAltScore(RoundLoop) = cdbl(sScore(RoundLoop)) - INT(cdbl(sScore(RoundLoop))/6)*6
        		IF cdbl(sScore(RoundLoop)) > 0 AND cdbl(sAltScore(RoundLoop)) = 0 THEN
        			sAltScore(RoundLoop) = 6.00


        			' --- If Metric less 3 if USA less 2
        			'If cdbl(sPQ2(RoundLoop)) > cdbl(38) then
        				sPQ2(RoundLoop) = cdbl(sPQ2(RoundLoop)) - 3                  
        			'ELSE                        
        			'	sPQ2(RoundLoop) = cdbl(sPQ2(RoundLoop)) - 2                  
        			'END IF
        		END IF
        	END IF		' --- Score is < or = number of buoys at completion of max boat speed ---

        	' --- Score is greater than # of buoys at completion of max boat speed ---
        	IF cdbl(sScore(RoundLoop)) > cdbl(rsSelectFields("BOUY_MAX")) THEN
        		sPQ2(RoundLoop) = rsSelectFields("MAX_S1")
        
        		LSHORT = INT((cdbl(sScore(RoundLoop)) - rsSelectFields("BOUY_MAX"))/6)
        		IF cdbl(rsSelectFields("LINE_MAX")) = 23.00 THEN STLINE=2
        		IF cdbl(rsSelectFields("LINE_MAX")) = 18.25 THEN STLINE=4
        		
        		FINDLINE=STLINE+LSHORT
        		
			' --- If extra buoys are 0, make 6 to show 6 at previous speed/line
        		sAltScore(RoundLoop) = sScore(RoundLoop) - INT(cdbl(sScore(RoundLoop))/6)*6
        		IF cdbl(sScore(RoundLoop)) > 0 AND cdbl(sAltScore(RoundLoop)) = 0 then
        			sAltScore(RoundLoop) = 6.00
        			FINDLINE = FINDLINE - 1
        		END IF
        		
        		if FINDLINE < 2 then sPQ1(RoundLoop) = rsSelectFields("Line_Max")
        		if FINDLINE = 2 then sPQ1(RoundLoop) = 18.25
        		if FINDLINE = 3 then sPQ1(RoundLoop) = 16.00
        		if FINDLINE = 4 then sPQ1(RoundLoop) = 14.25
        		if FINDLINE = 5 then sPQ1(RoundLoop) = 13.00
        		if FINDLINE = 6 then sPQ1(RoundLoop) = 12.00
        		if FINDLINE = 7 then sPQ1(RoundLoop) = 11.25
        		if FINDLINE = 8 then sPQ1(RoundLoop) = 10.75
        		if FINDLINE = 9 then sPQ1(RoundLoop) = 10.25
        		if FINDLINE = 10 then sPQ1(RoundLoop) = 9.75
		END IF

      	END IF		' --- Event is SLALOM


      	IF sEvent = "J" THEN
        	sAltScore(RoundLoop) = sScore(RoundLoop)*0.3048
	        sPQ2(RoundLoop) = rsSelectFields("MAX_J1")
      	END IF

     END IF	' --- Testing if DIV is in Division Table ---

NEXT 


rsSelectFields.close




%>
<table class="innertable" align=center width="<%=TourTableWidth%>"px border=0>
  <tr>
	<th align=center width=250><font color="#FFFFFF" size="<%=fontsize1%>">Event</font></th>
	<th align=center width=100><font color="#FFFFFF" size="<%=fontsize1%>">Place</font></th><%
	IF Session("sTSptsGrpID")="NCW" THEN %>
		<th align=center><font color="#FFFFFF" size="<%=fontsize1%>">Team</font></th><%
	END IF %>
	<th align=center><font color="#FFFFFF" size="<%=fontsize1%>">League</font></td>
  </tr>
  <tr><%

	
	SELECT CASE sEvent
	  CASE sTEvent(1)
		TempEventName=sTEventName(1)
	  CASE sTEvent(2)
		TempEventName=sTEventName(2)
	  CASE sTEvent(3)
		TempEventName=sTEventName(3)
	  CASE sTEvent(4)
		TempEventName=sTEventName(4)
   	END SELECT %>

	<td align=center><%=TempEventName%></td>
	<td align=center><%=Request("Place")%></td><%
	IF Session("sTSptsGrpID")="NCW" THEN %>
		<td align=center><%
			LoadTeam TeamSelected, "disabled"  %>
		</td><%
	END IF %>
	<td align=center>&nbsp;<%=Session("LeagueName")%></td>
  </tr>
</table>
<br>

<form action="/rankings/addscores.asp" method="post">

<input type="hidden" name="sMemberID" value="<%=sMemberID%>">
<input type="hidden" name="Tour_ID" value="<%=sTourID%>">
<input type="hidden" name="Event" value="<%=sEvent%>">
<input type="hidden" name="Place" value="<%=sPlace%>">
<input type="hidden" name="LeagueDrop" value="<%=LeagueSelected%>">
<input type="hidden" name="TeamDrop" value="<%=TeamSelected%>"><%

DisplayScoreDetails 

%>
<br><br>
<TABLE align=center border=0 width="60%">
  <tr>
    <td align=center>

	<input type=submit style="width:10em" name="Action" value="Save Scores">
    </td>
    <td align=center>
	<input type=submit style="width:10em" name="Action" value="Cancel">

    </td>
  </tr>
</form>
</TABLE>
<%

END SUB






' ------------------
  SUB UpdateScoreTable
' ------------------

'response.write("<br><br>In UPDATE function")
'response.end


' -------------------------------------------
' Verify all fields and save to database.
' -------------------------------------------

IF session("userlevel") > 39 then

    ' -------------------------------------------------
    ' --- Loops thru up to 3 rounds of the tournament ---
    ' -------------------------------------------------

    FOR RoundLoop = 1 TO sTRounds(Session("TotEv"))

'response.write("<br><br>RoundLoop="&Roundloop)
'response.write("<br>TRIM(sScore(RoundLoop))="&TRIM(sScore(RoundLoop)))
'response.write("<br>TRIM(sScore(RoundLoop))<>null=")
'response.write(TRIM(sScore(RoundLoop))<>"")

'response.write("<br><br>In Update - sDiv(1)="&sDiv(1))



	' -----------------------------------------------------
	' --- Checks for existence of scores for this round ---
	' -----------------------------------------------------

        IF sSPtsGrpID = "AWS" OR sSPtsGrpID = "NCW" THEN
		sSQL = "SELECT * from " & RawScoresTableName
	ELSE
		sSQL = "SELECT * from " & RawScoresOtherTableName
	END IF
        sSQL = sSQL + " WHERE MemberID = '" & sMemberID & "' AND"
        sSQL = sSQL + " TourID = '" & sTourID & "' and"
        sSQL = sSQL + " Event = '" & sEvent & "' and"
        sSQL = sSQL + " [Round] = '" & RoundLoop & "'"
        sSQL = sSQL + " ORDER BY MemberID, TourID, Event, [Round]"

        rsCheckDups.open sSQL, sConnectionToTRATable, 3,3


'response.write("<br><br>Existance of Scores - sSQL="&sSQL)
'response.end
'response.write("<br>If True then found Dupe - NOT(rsCheckDups.eof)=")
'response.write(NOT(rsCheckDups.eof))
'response.write("<br>TRIM(sScore(RoundLoop)) = ")
'response.write(TRIM(sScore(RoundLoop))="")
'response.write("<br>Err.Number="&Err.Number)

	' --- INSERT because there are 1) NO duplicates 2) No errors and 3) Score field is NOT null ---
        IF rsCheckDups.eof AND (Err.Number = 0) AND TRIM(sScore(RoundLoop)) <> "" THEN

		'response.write("<br>Round "&RoundLoop&" - INSERT - NO Dupes - so INSERT")
                 
	        IF sSPtsGrpID = "AWS" OR sSPtsGrpID = "NCW" THEN
        	  sSQL = "INSERT INTO " & RawScoresTableName
		ELSE
        	  sSQL = "INSERT INTO " & RawScoresOtherTableName
		END IF

		  sSQL = sSQL + " (LName, FName, MemberID, TourFed, TourID, [H-Class], EndDate, Event, Place, [Round], Class, Div, Perf_Qual1, Perf_Qual2, AltScore, Score, SptsGrpID)"
        	  sSQL = sSQL + " VALUES ("
	          sSQL = sSQL + "'" & TempMemLName & "',"
	          sSQL = sSQL + "'" & TempMemFName & "',"
        	  sSQL = sSQL + "'" & SQLClean(sMemberID) & "',"
	          sSQL = sSQL + "'" & "USA" & "',"
	          sSQL = sSQL + "'" & SQLClean(sTourID) & "',"
	          sSQL = sSQL + "'" & " " & "',"
	          sSQL = sSQL + "'" & sTDateE & "',"
	          sSQL = sSQL + "'" & sEvent & "',"

	          If isnumeric(sPlace) then
        	    sSQL = sSQL + "'"&sPlace&"',"
	          Else
        	    sSQL = sSQL + "NULL,"
	          End If
        	  sSQL = sSQL + "'" & RoundLoop & "',"
	          sSQL = sSQL + "'" & SQLClean(sClassArray(RoundLoop)) & "',"
        	  sSQL = sSQL + "'" & SQLClean(sDiv(RoundLoop)) & "',"
	          If isnumeric(sPQ1(RoundLoop)) then
        	    sSQL = sSQL + "'" & SQLClean(sPQ1(RoundLoop)) & "',"
	          Else
        	    sSQL = sSQL + "NULL,"
	          End If
	          If isnumeric(sPQ2(RoundLoop)) then
        	    sSQL = sSQL + "'" & SQLClean(sPQ2(RoundLoop)) & "',"
	          Else
        	    sSQL = sSQL + "NULL,"
	          End If
        	  If isnumeric(sAltScore(RoundLoop)) then
	            sSQL = sSQL + "'" & SQLClean(sAltScore(RoundLoop)) & "',"
        	  Else
	            sSQL = sSQL + "NULL,"
        	  End If
	          If isnumeric(sScore(RoundLoop)) Then
        	    sSQL = sSQL + "'" & SQLClean(sScore(RoundLoop)) & "',"
	          Else
	            sSQL = sSQL + "NULL,"
	          End If

	          sSQL = sSQL + "'" & sSptsGrpID & "')"

	          Con.Execute(sSQL)
		'markdebug(sSQL)

	ELSEIF NOT(rsCheckDups.eof) AND TRIM(sScore(RoundLoop)) = "" THEN	
	   ' --- DELETE because 1) Duplicate found and 2) the score input was null --- 

	   'response.write("<br>Round "&RoundLoop&" - DELETE - Dupes found and score is null")
	   'response.end	
	        IF sSPtsGrpID = "AWS" OR sSPtsGrpID = "NCW" THEN
        	  sSQL = "DELETE " & RawScoresTableName
		ELSE
        	  sSQL = "DELETE " & RawScoresOtherTableName
		END IF

	        sSQL = sSQL + " WHERE MemberID='"&SQLClean(sMemberID)&"' AND TourID='"&SQLClean(sTourID)&"'"
		sSQL = sSQL + " AND [Round]='"&RoundLoop&"' AND Event='"&sEvent&"'"
                
		'response.write("<br>sSQL="&sSQL)

	        Con.Execute(sSQL)

	ELSE 	

	   ' --- UPDATE because 1) Duplicate found and 2) the score input is NOT null --- 

	   'response.write("<br>Round "&RoundLoop&" - UPDATE - Dupes found and score is NOT null")


	        IF sSPtsGrpID = "AWS" OR sSPtsGrpID = "NCW" THEN
        	  sSQL = "UPDATE " & RawScoresTableName
		ELSE
        	  sSQL = "UPDATE " & RawScoresOtherTableName
		END IF

	          sSQL = sSQL + " SET"

	          IF isnumeric(request("Place")) THEN
			sSQL = sSQL + " Place='"&sPlace&"',"
	          ELSE
			sSQL = sSQL + " Place=NULL,"
	          END IF
        	  sSQL = sSQL + " [Round]='"&RoundLoop&"', Class='"&SQLClean(sClassArray(RoundLoop))&"', Div='"&SQLClean(sDiv(RoundLoop))&"',"

	          IF isnumeric(sPQ1(RoundLoop)) THEN
        	    	sSQL = sSQL + " Perf_Qual1='"&SQLClean(sPQ1(RoundLoop))&"',"
	          ELSE
        	    	sSQL = sSQL + " Perf_Qual1=NULL,"
	          END IF

	          IF isnumeric(sPQ2(RoundLoop)) THEN
			sSQL = sSQL + " Perf_Qual2='" & SQLClean(sPQ2(RoundLoop)) & "',"
	          ELSE
			sSQL = sSQL + " Perf_Qual2=NULL,"
	          END IF

        	  IF isnumeric(sAltScore(RoundLoop)) THEN
			sSQL = sSQL + " AltScore='"&SQLClean(sAltScore(RoundLoop))&"',"
        	  ELSE
			sSQL = sSQL + " AltScore=NULL,"
        	  END IF

	          IF isnumeric(sScore(RoundLoop)) THEN
			sSQL = sSQL + " Score='"&SQLClean(sScore(RoundLoop))&"',"
	          Else
			sSQL = sSQL + " Score=NULL,"
	          End If


		  sSQL = sSQL + " SptsGrpID='"&sSptsGrpID&"'"

	          sSQL = sSQL + " WHERE MemberID='"&SQLClean(sMemberID)&"' AND TourID='"&SQLClean(sTourID)&"'"
		  sSQL = sSQL + " AND [Round]='"&RoundLoop&"' AND Event='"&sEvent&"'"

	          Con.Execute(sSQL)

        END IF	
	' --- Test for duplicates ---


        rsCheckDups.close



'      END IF	
	' --- Test of Null Value of Score or Placement Points ---

    NEXT  	' --- Loop through rounds ---



    ' --- Set the Session variable to whatever the last Div(1) was, even if this was a deletion --- 
    Session("sDiv")=sDiv(1)


    response.write ("<center><font color=blue>Score(s) Saved</font></center>")

ELSE

    response.write ("<center><font color=red>Scores Not Saved -- Your security level is not high enough to save scores.</font></center>")

END IF		' --- Session authority condition ---


END SUB



' -------------------------
   SUB DisplayScoreDetails
' -------------------------


'response.write("<br><br>DisplayScoreDetails")


%>
<TABLE class="innertable" align=center width="<%=TourTableWidth%>">
  <tr><%

     DisplayScoreLineHeader
  %>	
  </tr><% 

  FOR RoundLoop = 1 TO sTRounds(Session("TotEv"))
	    %>
	    <tr>
	      <td align=center width=130 color="black"><font color="<%=TextColor1%>" size="<%=fontsize1%>">Round <%=RoundLoop%>:</font></td>
	      <td align=center width=150 color="<%=TextColor1%>"><font color="<%=TextColor1%>" size="<%=fontsize1%>"><%=sDiv(RoundLoop)%></font></td>
	      <td align=center width=120 color="<%=TextColor1%>"><font color="<%=TextColor1%>" size="<%=fontsize1%>">&nbsp;<%=sScore(RoundLoop)%></font></td><%

	     SELECT CASE sEvent
		CASE "S"  %>	
			<td align=center><input type=text size=7 Maxlength=7 name="AltScore_<%=RoundLoop%>" value="<%=sAltScore(RoundLoop)%>"></td>
	        	<td align=center><input type=text size=7 Maxlength=7 name="PQ1_<%=RoundLoop%>" value="<%=sPQ1(RoundLoop)%>"></td>
      			<td align=center><input type=text size=7 Maxlength=7 name="PQ2_<%=RoundLoop%>" value="<%=sPQ2(RoundLoop)%>"></td><%

		CASE "J" %>

			<td align=center><input type=text size=7 Maxlength=7 name="AltScore_<%=RoundLoop%>" value="<%=sAltScore(RoundLoop)%>"></td><%

			RampFieldName="PQ1_"&RoundLoop %>
			<td align=center>
			  <font color="<%=TextColor1%>" size="<%=fontsize1%>"><%

				' --- SUB in Tools_Definitions.asp --
				LoadRampPulldown_Nov1_2010 sDiv(RoundLoop), RampFieldName, sPQ1(RoundLoop), SRampStatus  %>	
			  </font>
			</td>
	      		<td align=center>
				<input type=text name="PQ2_<%=RoundLoop%>" size=7 MaxLength=7 value="<%=sPQ2(RoundLoop)%>">
			</td><%
	     END SELECT %>	

	      <td align=center><font color="<%=TextColor1%>" size="<%=fontsize1%>"><%=sClassArray(RoundLoop)%></font></td>
    	    </tr>

	      <input type="hidden" name="Div_<%=RoundLoop%>" value="<%=sDiv(RoundLoop)%>">
	      <input type="hidden" name="Score_<%=RoundLoop%>" value="<%=sScore(RoundLoop)%>">
	      <input type="hidden" name="Class_<%=RoundLoop%>" value="<%=sClassArray(RoundLoop)%>"><%

  NEXT %>
</TABLE><%


END SUB




' ----------------------
  SUB GetScoreDetail
' ----------------------

response.write("<br><br>GETScoreDetails - sDiv(1)="&sDiv(1))

%>
<TABLE class="innertable" align=center width="<%=TourTableWidth%>">
  <tr><%
	DisplayScoreLineHeader %>
  </tr><% 

  FOR RoundLoop = 1 TO sTRounds(Session("TotEv")) 
	IF sScore(RoundLoop) >= 0 THEN
	    %>
	    <tr>
	      <td align=center width=130><font color="<%=TextColor1%>" size="<%=fontsize1%>">Round <%=RoundLoop%>:</font></td>
	      <td align=center width=150><font color="<%=TextColor1%>" size="<%=fontsize1%>"><%=sDiv(RoundLoop)%></font></td>
	      <td align=center><font color="<%=TextColor1%>" size="<%=fontsize1%>"><%=sScore(RoundLoop)%></font></td>
	      <input type="hidden" name="Div_<%=RoundLoop%>" value="<%=sDiv(RoundLoop)%>"><%



	     SELECT CASE TRIM(sEvent)
		CASE "S"  %>	
			<td align=center><input type=text size=7 MaxLength=7 name="AltScore_<%=RoundLoop%>" value="<%=sAltScore(RoundLoop)%>"></td>
	        	<td align=center><input type=text size=7 MaxLength=7 name="PQ1_<%=RoundLoop%>" value="<%=sPQ1(RoundLoop)%>"></td>
      			<td align=center><input type=text size=7 MaxLength=7 name="PQ2_<%=RoundLoop%>" value="<%=sPQ2(RoundLoop)%>"></td><%
		CASE "J" %>	
			<td align=center><input type=text size=7 MaxLength=7 name="AltScore_<%=RoundLoop%>" value="<%=sAltScore(RoundLoop)%>"></td>
		        <td align=center><input type=text size=7 MaxLength=7 name="PQ1_<%=RoundLoop%>" value="<%=sPQ1(RoundLoop)%>"></td>
	      		<td align=center><input type=text size=7 MaxLength=7 name="PQ2_<%=RoundLoop%>" value="<%=sPQ2(RoundLoop)%>"></td><%
	     END SELECT %>	
	      <td align=center><%
		     	' --- Define parameters and run sub in tools_define.asp ---
			sClassArray(RoundLoop)=""
			sClassFieldName="Class_"&RoundLoop

			BuildClassDrop_NEW sClassFieldName, sClassArray(RoundLoop), TRIM(sEvent), ScoreFieldStatus %>
	      </td>
    	    </tr>
  	    <%
	END IF
  NEXT %>
</TABLE><%



END SUB








' ----------------------------------
   SUB BuildEventDrop (EventStatus)
' ----------------------------------


%><select name="Event" onchange=submit() style="width:10em" <%=EventStatus%>><%



  IF TRIM(sTEvent(1))<>"" THEN 
	%><option value="<%=sTEvent(1)%>" <%IF sTEvent(1) = sEvent THEN Response.Write("Selected")%>><%=sTEventName(1)%></option><%
  END IF

  IF TRIM(sTEvent(2))<>"" THEN 
	%><option value="<%=sTEvent(2)%>" <%IF sTEvent(2) = sEvent THEN Response.Write("Selected")%>><%=sTEventName(2)%></option><%
  END IF

  IF TRIM(sTEvent(3))<>"" THEN 
	%><option value="<%=sTEvent(3)%>" <%IF sTEvent(3) = sEvent THEN Response.Write("Selected")%>><%=sTEventName(3)%></option><%
  END IF

  IF TRIM(sTEvent(4))<>"" THEN 
	%><option value="<%=sTEvent(4)%>" <%IF sTEvent(4) = sEvent THEN Response.Write("Selected")%>><%=sTEventName(4)%></option><%
  END IF

%></select><%         



END SUB


' -----------------------------------
  SUB DisplayPageBanner (BannerTitle)
' -----------------------------------

    %>
    <TABLE class="droptable" align=center width="<%=TourTableWidth%>"px ><% '---Table to hold image --- %>
	<tr><%

	  ' --- Defines Logo and Logo dimensions ---
	  SetLogoParameters sSptsGrpID

	  %>
	  <td height="<%=MainLogoHeight%>" width="<%=MainLogoWidth%>"px align=center vAlign=bottom noWrap background="<%=MainLogo%>">			
	  </td>
	  <td align=center>
    	    <font size=4 color="<%=TextColor2%>"><B><I><%=SD_Heading%></I></B></font>
	  </td>

	  <td align=center>
    	    <font size=4 color="<%=TextColor2%>"><B><I><%=BannerTitle%></I></B></font>
	  </td>

	</tr>
    </TABLE><% 	
END SUB


' ---------------------------
' --- COMMENTED OUT TO TEST WHETHER OK TO REMOVE THIS CODE ---

'  SUB CreateDivisionDropDown
' ---------------------------

SkipThis="N"
IF SkipThis="Y" THEN


IF sDiv = "" THEN
' We don't want to default to any old random division
' First we try to figure out what division they SHOULD
' be in.  If we can't find that, then we look through
' our existing scores and see what the last division 
' this member was in and use that.
' Unless the user has already selected a division.
  sDate = CDate("01/01/"&(2000+(left(sTourID,2))))
            
  sSQL = "Select top 1 * from " & SkiYearTableName & " where '"&sDate&"' BETWEEN BeginDate and EndDate "
  rs.open sSQL, SConnectionToTRATable, 3, 1
            

  ' We have to redo this because division is not based on REAL age, it's based on age relative to ski year.
  If NOT IsNull(TempBirthDate) THEN
    ' get absolute number of years 
    AgeInYears = cint(datediff("YYYY", TempBirthDate, sDate)) - 1

    IF AgeInYears > 16 THEN
       sSQL = "SELECT TOP 1 div FROM "&DivisionsTableName&" WHERE LEFT(Div,1) ='"&UCASE(left(sMembSex,1)) &"' AND "&AgeInYears&" <= Up_Age and "&AgeInYears&" >= Low_Age and SkiYearID = "&rs("SkiYearID")&" ORDER BY Div"
    ELSE
       IF ucase(left(sMembSex,1)) = "M" Then
         sSQL = "SELECT TOP 1 div FROM " & DivisionsTableName & " WHERE left(Div,1) = 'B' and "&AgeInYears&" <= Up_Age AND "&AgeInYears&" >= Low_Age AND SkiYearID = "& rs("SkiYearID") &" ORDER BY Div"
       ELSE
         sSQL = "SELECT TOP 1 div FROM " & DivisionsTableName & " WHERE left(Div,1) = 'G' and "&AgeInYears&" <= Up_Age AND "&AgeInYears&" >= Low_Age AND SkiYearID = "& rs("SkiYearID") &" ORDER BY Div"
       END IF
    END IF

  ELSE
        sSQL = "Select top 1 div from " & DivisionsTableName & " where 0=1"
  END IF

  rs.close 	' --- Close the Ski Year Table ---


  rs.open sSQL, SConnectionToTRATable, 3, 1      ' --- Open the Division table ---

  IF rs.EOF THEN
	rs.close

	IF sSPtsGrpID = "AWS" OR sSPtsGrpID = "NCW" THEN
	   sSQL = "SELECT TOP 1 * FROM " & RawScoresTableName
	ELSE	
	   sSQL = "SELECT TOP 1 * FROM " & RawScoresOtherTableName
	END IF	

	sSQL = sSQL + " WHERE memberid='" & sqlclean(sMemberID) &"' ORDER BY EndDate DESC"
	rs.open sSQL, SConnectionToTRATable

    	IF rs.eof THEN 
      	   sDiv = ""
    	ELSE
           sDiv = rs("Div")
    	END IF  
  ELSE
    sDiv = rs("Div")
  END IF
  rs.close


END IF   ' --- If sDiv="" ---


END IF

'END SUB


' -----------------------------
' --- COMMENTED OUT TO TEST WHETHER OK TO REMOVE THIS CODE ---

'  SUB CreateDivisionDropDown2
' -----------------------------

SkipThis="N"
IF SkipThis="Y" THEN

      sSQL = "SELECT DISTINCT div, div_name FROM " & DivisionsTableName & " ORDER BY div"
      rsSelectFields.open sSQL, SConnectionToTRATable

      IF NOT rsSelectFields.eof THEN 
        DO WHILE NOT rsSelectFields.eof
          %>
            <option value="<%=rsSelectFields("Div")%>" <%If sDiv(RoundLoop) = rsSelectFields("Div") Then Response.Write("selected")%>><%=rsSelectFields("Div")%> - <%=rsSelectFields("Div_Name")%></option><br>
          <%
          rsSelectFields.movenext
        LOOP
      ELSE 
          response.write("<option value="" "" selected> </option>")
      END IF
      rsSelectFields.close

'END SUB

END IF

' --------------------------
  SUB WriteHeaders(sTitle)
' --------------------------

' Write Headers for DB Page
%>
<TABLE BORDER="0" CELLPADDING="6" CELLSPACING="0" WIDTH="100%" BGCOLOR="#C0C0C0" BORDERCOLOR="#C0C0C0" BORDERCOLORDARK="#C0C0C0" BORDERCOLORLIGHT="#C0C0C0" >
<TR>
<TD ALIGN="Left"><Font Face="courier" COLOR="#000000" SIZE="4"><B><% Response.Write(sTitle) %></B></FONT></TD>
</TR>
</TABLE>
<BR>

<%
END SUB




' --------------------
  SUB SearchMember
' --------------------


If sMemberID = "" and sLastName = "" and sFirstName = "" Then 


%>


<br><br>
<center> <center><font size=4 color="<%=TextColor2%>"><B><I>Find a Member To Add <%=sSptsGrpID%> Scores</I></B></font>
<form action="/rankings/addscores.asp" method="post">
<input type="hidden" name="pvar" value="SearchMember">
<input type="hidden" name="Tour_ID" value="<%=sTourID%>">
<br>

<TABLE class="innertable" BORDER="1" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#FFFFFF" width=60%>
<TR>
<TH ALIGN="Left" vAlign="top"><Center><FONT COlOR="#FFFFFF" SIZE="1">Member ID</FONT></Center></TH>
<TH ALIGN="Left" vAlign="top"><Center><FONT COlOR="#FFFFFF" SIZE="1">Last Name</FONT></Center></TH>
<TH ALIGN="Left" vAlign="top"><Center><FONT COlOR="#FFFFFF" SIZE="1">First Name</FONT></Center></TH>
</TR>

<TR>
<TD ALIGN="Left" vAlign="top" bgcolor="#C0C0C0"><Center><FONT SIZE="1"><input type="text" name="sMemberID" size=9></input></FONT></Center></TD>
<TD ALIGN="Left" vAlign="top" bgcolor="#C0C0C0"><Center><FONT SIZE="1"><input type="text" name="Last_Name" size=15></input></FONT></Center></TD>
<TD ALIGN="Left" vAlign="top" bgcolor="#C0C0C0"><Center><FONT SIZE="1"><input type="text" name="First_Name" size=15></input></FONT></Center></TD>
</TR>
</table>
<br><br>
<input type="submit" value="Begin Search"></form>

<% 
else
  sSQL = "Select top 10 PersonIDwithCheckDigit,LastName,FirstName,City,State,BirthDate,Sex from " & MemberTableName & " where "
  If sLastName <> "" Then
    sSQL = sSQL + "lower(lastname) LIKE '%" & SQLClean(lcase(sLastName)) & "%'"
    If sFirstName <> "" or sMemberID <> "" Then
      sSQL = sSQL + " and "
    End If
  End If
  If sFirstName <> "" Then
    sSQL = sSQL + "lower(firstname) LIKE '%" & SQLClean(lcase(sFirstName)) & "%'"
    if sMemberID <> "" then
      sSQL = sSQL + " and "
    end if
  End If
  If sMemberID <> "" Then
    sSQL = sSQL + "PersonIDwithCheckDigit LIKE '%" & SQLClean(sMemberID) & "%'"
  End If
  sSQL = sSQL + " and membertypeid <> 2 order by PersonIDWithCheckDigit"
      
  rs.open sSQL, sConnectionToTRATable, 3, 1
 
  if rs.EOF then 

%>

   <br><br>
    <center><h2>Add Scores<br></h2>
    <form action="/rankings/addscores.asp" method="post">
      <input type="hidden" name="pvar" value="SearchMember">
      <input type="hidden" name="Tour_ID" value="<%=sTourID%>">
    
      <br>
      <br><font color="red"> No Records Found, Please Search Again </font><br>
        
      <TABLE BORDER="1" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#FFFFFF" width=60%>
        <TR>
          <TD ALIGN="Left" vAlign="top"><Center><FONT SIZE="1">Member ID</FONT></Center></TD>
          <TD ALIGN="Left" vAlign="top"><Center><FONT SIZE="1">Last Name</FONT></Center></TD>
          <TD ALIGN="Left" vAlign="top"><Center><FONT SIZE="1">First Name</FONT></Center></TD>
          <TD ALIGN="Left" vAlign="top"><Center><FONT SIZE="1">Age/Gender</FONT></Center></TD>
        </TR>
    
        <TR>
          <TD ALIGN="Left" vAlign="top" bgcolor="#C0C0C0"><Center><FONT SIZE="1"><input type="text" name="sMemberID" value="<%=sMemberID%>" size=9></input></FONT></Center></TD>
          <TD ALIGN="Left" vAlign="top" bgcolor="#C0C0C0"><Center><FONT SIZE="1"><input type="text" name="Last_Name" value="<%=sLastName%>" size=15></input></FONT></Center></TD>
          <TD ALIGN="Left" vAlign="top" bgcolor="#C0C0C0"><Center><FONT SIZE="1"><input type="text" name="First_Name" value="<%=sFirstName%>" size=15></input></FONT></Center></TD>
          <TD ALIGN="Left" vAlign="top" bgcolor="#C0C0C0"><Center><FONT SIZE="1"><input type="text" name="First_Name" value="<%=sMembAge%>" size=15></input></FONT></Center></TD>
        </TR>
      </table>
      <br>
      <input type="submit" value="Search Member Database">
    </form>
   
 <%
  Else




    If rs.recordcount > 1 then %>
  
      <br><br>
      <center><h2>Add Scores<br></h2>
      <form action="/rankings/addscores.asp" method="post">
        <input type="hidden" name="pvar" value="SearchMember">
        <input type="hidden" name="Tour_ID" value="<%=sTourID%>">
        <b><center>Search Results</b>
        <br><small>Click on an ID or Name to Select that Member</small>
<%  
       If rs.recordcount > 9 Then %>
         <br><font color="red"><small>More then ten records found.  Only the top ten records were displayed.<br>
         Please refine your search parameters.</small></font><br><br>
<% 
       End If 
%>
      <TABLE class="innertable" BORDER="1" width=50%>
        <TR>
          <TH ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" SIZE="1">Member ID</FONT></TH>
          <TH ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" SIZE="1">Last Name</FONT></TH>
          <TH ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" SIZE="1">First Name</FONT></TH>
          <TH ALIGN="center" vAlign="top"><FONT COlOR="#FFFFFF" SIZE="1">Age/Gender</FONT></TH>
        </tr>
<% 
        Do While Not rs.EOF 
		IF sTDateS<>"" THEN
	    		sMembAge = AgeAtDate(sTDateS, rs("PersonIDWithCheckDigit"))
		ELSE
	    		sMembAge = AgeAtDate(date, rs("PersonIDWithCheckDigit"))
		END IF


'markdebug("now="&date&"   -  sTDateS = "&sTDateS& " - rs(PersonIDWithCheckDigit)="&rs("PersonIDWithCheckDigit"))

%>
          <tr>
            <TD ALIGN="Center" vAlign="top"><FONT COlOR="#000000" SIZE="1"><a href="/rankings/addscores.asp?sMemberID=<%=rs("PersonIDwithCheckDigit")%>&tour_id=<%=sTourID%>"><%=rs("PersonIDwithCheckDigit")%></a></FONT></TD>
            <TD ALIGN="Center" vAlign="top"><FONT COlOR="#000000" SIZE="1"><a href="/rankings/addscores.asp?sMemberID=<%=rs("PersonIDwithCheckDigit")%>&tour_id=<%=sTourID%>"><%=rs("LastName")%></a></FONT></TD>
            <TD ALIGN="Center" vAlign="top"><FONT COlOR="#000000" SIZE="1"><a href="/rankings/addscores.asp?sMemberID=<%=rs("PersonIDwithCheckDigit")%>&tour_id=<%=sTourID%>"><%=rs("FirstName")%></a></FONT></TD>
            <TD ALIGN="Center" vAlign="top"><FONT COlOR="#000000" SIZE="1"><%=sMembAge%>/<%=rs("Sex")%></FONT></TD>
          </tr>
<% 
          rs.MoveNext %>
<% 
        Loop 
%>
        </table>
        <br><br>
        <TABLE class="innertable" BORDER="1" width=60%>
        <TR>
          <TH ALIGN="Left"><Center><FONT COlOR="#FFFFFF" SIZE="1">Member ID</FONT></Center></TH>
          <TH ALIGN="Left"><Center><FONT COlOR="#FFFFFF" SIZE="1">Last Name</FONT></Center></TH>
          <TH ALIGN="Left"><Center><FONT COlOR="#FFFFFF" SIZE="1">First Name</FONT></Center></TH>
        </TR>
    
        <TR>
          <TD ALIGN="Left" ><FONT SIZE="1"><input type="text" name="sMemberID" value="<%=sMemberID%>" size=9></input></FONT></TD>
          <TD ALIGN="Left" ><FONT SIZE="1"><input type="text" name="Last_Name" value="<%=sLastName%>" size=15></input></FONT></TD>
          <TD ALIGN="Left" ><FONT SIZE="1"><input type="text" name="First_Name" value="<%=sFirstName%>" size=15></input></FONT></TD>
        </TR>
        </table>
        <br>
        <input type="submit" value="Search Member Database">
      </form>
      </center>
<% 
    Else
      '-----------------------------------------------------
      ' We found a unique Member and now we can find a Tour.
      '-----------------------------------------------------

      WriteHeader

      BannerTitle = "Confirm Selection"	
      DisplayPageBanner (BannerTitle)

      %>
      <br><br>	
      <TABLE width="<%=TourTableWidth%>"px align=center>
	 <tr>
	  <td align=center colspan=2>	
		<FONT SIZE="3" color="<%=TextColor2%>"><B>You have selected</B></font> 
	  </td>
	 </tr>
	 <tr><td colspan=2>&nbsp;</td></tr>
	 <tr>
	  <td align=right><FONT SIZE="1"><b>Member Name:&nbsp;</b></font></td> 
	  <td align=left><FONT SIZE="1" color="<%=TextColor2%>">&nbsp;<%=rs("FirstName")%>&nbsp;<%=rs("LastName")%></font></td> 
 	 </tr>
	 <tr>
	  <td align=right><FONT SIZE="1"><b>MemberID:</b>&nbsp;</font></td> 
	  <td align=left><FONT SIZE="1" color="<%=TextColor2%>">&nbsp;<%=rs("PersonIDWithCheckDigit")%></font></td>
	 </tr>
	 <tr>
	  <td align=center colspan=2>	
	  <br>
      	  <FONT SIZE="3" color="red"><I><B>Is this the correct member? </B></I></font> 
	  </td>
 	 </tr>
      </TABLE>
	<br><br>
        <table border=0 align=center>
        <tr><td align=center>
      <form action="/rankings/addscores.asp" method="post">
        <input type="hidden" name="sMemberID" value="<%=rs("PersonIDwithCheckDigit")%>">
        <input type="hidden" name="Tour_ID" value="<%=sTourID%>">
        <input type="submit" style="width:9em" value="Yes">
      </form>
      </td>
      <td>&nbsp;</td>
      <td>
      <form action="/rankings/addscores.asp" method="post">
        <input type="hidden" name="pvar" value="SearchMember">
        <input type="hidden" name="Tour_ID" value="<%=sTourID%>">
        <input type="submit" style="width:9em" value="No">
      </form>
      </td><tr></table>
      <%      
      
      WriteFooter
 
     End If
  End If

END IF 

END SUB



' -----------------
  SUB SearchTour
' -----------------

DateGood = 0

IF (isnumeric(left(sTourDate,2)) and isnumeric(right(left(sTourDate,5),2)) and isnumeric(right(sTourDate,4)) and right(left(sTourDate,3),1) = "/" and right(left(sTourDate,6),1) = "/" and isDate(sTourDate)) or (sTourDate = "") then
  	DateGood = 1
ELSE
  	DateGood = 0
END IF

 
IF (sTourID = "" and sTourName = "" and sTourDate = "") or (DateGood = 0) then ' Check to see if the search fields have been used.

      WriteHeader
	%>
	<center><font size=4 color="<%=TextColor2%>"><B><I>Set Tournament Search Parameters</I></B></font>
      <form action="/rankings/addscores.asp" method="post">
        <input type="hidden" name="pvar" value="SearchTour">
        <input type="hidden" name="sMemberID" value="<%=sMemberID%>">

      <TABLE class="innertable" align="center" BORDER="1" width="<%=TourTableWidth%>"px>
        <TR>
          <TH ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" SIZE="1">Tournament ID</FONT></TH>
          <TH ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" SIZE="1">Name</FONT></TH>
          <TH ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" SIZE="1">Date (mm/dd/yyyy)</FONT></TH>
        </tr>

        <TR>
          <TD ALIGN="Center" ><FONT SIZE="1"><input type="text" name="Tour_ID" value="<%=sTourID%>" size=15></input></FONT></TD>
          <TD ALIGN="Center" ><FONT SIZE="1"><input type="text" name="Tour_Name" value="<%=sTourName%>" size=25></input></FONT></TD>
          <TD ALIGN="Center" ><FONT SIZE="1"><input type="text" name="Tour_Date" value="<%=sTourDate%>" size=20></input></FONT></TD>
        </TR>

	</TABLE>
	<br>
	<center><font size=2 color="<%=TextColor3%>"><B><I>IMPORTANT: For best results, use only snibbits of tournament name.</I></B></font>
	<br><br>	
        <input type="submit" value="Begin Search"><br><br><br>
      </form>
      <%      
      
      WriteFooter

ELSE 		' ---  Check to see if the search fields have been used.


	sSQL = "SELECT TOP 10 TSanction, TDateS, TDateE, TName, TCity, TState,"	
	sSQL = sSQL + " TRoundsWakeBd, TRoundsWSkate, TRoundsWSurf, KRoundsFlip, KRoundsFree,"
	sSQL = sSQL + " TRoundsS, TRoundsT, TRoundsJ, TRoundsF"
	
	sSQL = sSQL + " FROM "&SanctionTableName&" WHERE 1=1"

	IF sTourID <> "" THEN
		sSQL = sSQL + " AND lower(left(TSanction," & len(sTourID) & ")) = '" & SQLClean(lcase(sTourID)) & "'"
	END IF

	IF sTourName <> "" THEN
		sSQL = sSQL + " AND lower(TName) LIKE '%" & SQLClean(lcase(sTourName)) & "%'"
	END IF

	IF sTourDate <> "" THEN
		sSQL = sSQL + " AND (TDateE = '" & sTourDate & "' or TDateS = '" & sTourDate & "')"
	END IF


	SELECT CASE sSptsGrpID
		CASE "AWS"
			sSQL = sSQL + " AND SptsGrpID='AWS'"			
		CASE "USW"
			sSQL = sSQL + " AND SptsGrpID='USW'"
		CASE "AKA"
			sSQL = sSQL + " AND SptsGrpID='AKA'"
		CASE "NCW"
			sSQL = sSQL + " AND SptsGrpID='NCW'"

	END SELECT



	sSQL = sSQL + " and TStatus in (2,4,5) ORDER BY TDateS DESC"
	rs.open sSQL, sConnectionToTRATable, 3, 1




IF rs.EOF THEN 		' --- Tour ID Not Found ---

      'WriteHeader
      'WriteHeaders("Add Scores")

	BannerTitle="New Search"
	DisplayPageBanner (BannerTitle)

      %>
   <br><br>
      <TABLE width="<%=TourTableWidth%>"px align=center>
	 <tr>
     	   <form action="/rankings/addscores.asp" method="post">
		<input type="hidden" name="pvar" value="SearchTour">
        	<input type="hidden" name="sMemberID" value="<%=sMemberID%>">
	   <td align=center colspan=2>	
      		<font size="2" color="red"> <b>No Matching Tournaments Found, Please Search Again</b></font>
	   </td>
	 </tr>
	 <tr><td colspan=2>&nbsp;</td></tr>
	 <tr>
	  <td align=right><FONT SIZE="1"><b>Tournament Name:&nbsp;</b></font></td> 
	  <td align=left><input type="text" name="Tour_Name" value="<%=sTourName%>" size=20></td> 
 	 </tr>
	 <tr>
	  <td align=right><FONT SIZE="1"><b>TourID:</b>&nbsp;</font></td> 
	  <td align=left><input type="text" name="Tour_ID" value="<%=sTourID%>" size=20></td>
	 </tr>

	 <tr>
	  <td align=right><FONT SIZE="1"><b>Tour Date:</b>&nbsp;</font></td> 
	  <td align=left><input type="text" name="Tour_Date" value="<%=sTourDate%>" size=20><br><small>(mm/dd/yyyy)</small></td>
	 </tr>

	 <tr>
	  <td colspan=2 align=center>
		<br><br>
	        <input type="submit" value="Start Search">
	  </td>
	 <tr>
      </TABLE>


      </form>
      <%      
      
      WriteFooter

 ELSE 	' --- Tour ID Found ---

   ' ----------------------------------------------	
   ' --- Display the Tournament Search input table 	
   ' ---------------------------------------------- 

    If rs.recordcount > 1 Then 
      WriteHeader

      %>
      <center><font size=4 color="<%=TextColor2%>"><B><I>Set Tournament Search Parameters</I></B></font>
      <form action="/rankings/addscores.asp" method="post">
        <input type="hidden" name="pvar" value="SearchTour">
        <input type="hidden" name="sMemberID" value="<%=sMemberID%>">

      <TABLE class="innertable" align="center" BORDER="1" width="<%=TourTableWidth%>"px>
        <TR>
          <TH ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" SIZE="1">Tournament ID</FONT></TH>
          <TH ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" SIZE="1">Name</FONT></TH>
          <TH ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" SIZE="1">Date (mm/dd/yyyy)</FONT></TH>
        </tr>
        <TR>
          <TD ALIGN="Center" ><FONT SIZE="1"><input type="text" name="Tour_ID" value="<%=sTourID%>" size=15></input></FONT></TD>
          <TD ALIGN="Center" ><FONT SIZE="1"><input type="text" name="Tour_Name" value="<%=sTourName%>" size=25></input></FONT></TD>
          <TD ALIGN="Center" ><FONT SIZE="1"><input type="text" name="Tour_Date" value="<%=sTourDate%>" size=20></input></FONT></TD>
        </TR>
      </TABLE>
	<br>
	<center><font size=2 color="<%=TextColor3%>"><B><I>IMPORTANT: For best results, use only snibbits of tournament name.</I></B></font>
	<br><br>	
        <input type="submit" value="Begin Search"><br><br><br>
      </form>

        <b><center>Search Results</b>
        <br><small>Click on an ID or Name to Select that Tournament</small><%  

       If rs.recordcount > 9 THEN %>
        	<br><font color="red"><small>More then ten records found.  Only the top ten records were displayed.<br>
	        Please refine your search parameters.</small></font><br><br><% 
       END IF  

	%>
      <TABLE class="innertable" BORDER="1" width="<%=TourTableWidth%>"px>
        <TR>
          <TH ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" SIZE="1">Tournament ID</FONT></TH>
          <TH ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" SIZE="1">Tournament Name</FONT></TH>
          <TH ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" SIZE="1">Tournament Date</FONT></TH>
        </tr><% 

        DO WHILE NOT rs.EOF %>
	          <tr>
        	    <TD ALIGN="Center" ><FONT SIZE="1"><a href="/rankings/addscores.asp?sMemberID=<%=sMemberID%>&tour_id=<%=rs("TSanction")%>"><%=rs("TSanction")%></a></FONT></TD>
	            <TD ALIGN="Center" ><FONT SIZE="1"><a href="/rankings/addscores.asp?sMemberID=<%=sMemberID%>&tour_id=<%=rs("TSanction")%>"><%=rs("TName")%></a></FONT></TD>
        	    <TD ALIGN="Center" ><FONT SIZE="1"><a href="/rankings/addscores.asp?sMemberID=<%=sMemberID%>&tour_id=<%=rs("TSanction")%>"><%=rs("TDateE")%></a></FONT></TD>
	          </tr><% 
          	rs.MoveNext 
        LOOP 	

	%>
        </table>
      <%      
      
      WriteFooter

    ELSE

        '-----------------------------------------------------
        ' We found a tour ID and now we can add scores.
        '-----------------------------------------------------

        WriteHeader

      BannerTitle = "Confirm Selection"	
      DisplayPageBanner (BannerTitle)

      %>
      <br><br>	
      <TABLE width="<%=TourTableWidth%>"px align=center>
	 <tr>
	  <td align=center colspan=2>	
		<FONT SIZE="3" color="<%=TextColor2%>"><B>You have selected</B></font> 
	  </td>
	 </tr>
	 <tr><td colspan=2>&nbsp;</td></tr>
	 <tr>
	  <td align=right><FONT SIZE="1"><b>Tournament Name:&nbsp;</b></font></td> 
	  <td align=left><FONT SIZE="1" color="<%=TextColor2%>"><%=rs("TName")%></font></td> 
 	 </tr>
	 <tr>
	  <td align=right><FONT SIZE="1"><b>TourID:</b>&nbsp;</font></td> 
	  <td align=left><FONT SIZE="1" color="<%=TextColor2%>"><%=rs("TSanction")%></font></td>
	 </tr>
	 <tr>
	  <td align=right><FONT SIZE="1"><b>City/ST:</b>&nbsp;</font></td> 
	  <td align=left><FONT SIZE="1" color="<%=TextColor2%>"><%=rs("TCity")%>,&nbsp;<%=rs("TState")%></font></td>
	 </tr>
	 <tr>
	  <td align=center colspan=2>	
	  <br>
      	  <FONT SIZE="3" color="red"><I><B>Is this the correct tournament? </B></I></font> 
	  </td>
 	 </tr>
      </TABLE>
	<br><br>
        <table border=0 align=center>
        <tr><td align=center>
        <form action="/rankings/addscores.asp" method="post">
          <input type="hidden" name="sMemberID" value="<%=sMemberID%>">
          <input type="hidden" name="Tour_ID" value="<%=rs("TSanction")%>">
          <input type="submit" style="width:9em" value="Yes">
        </form>
        </td>
        <td>&nbsp;</td>
        <td align=center>
        <form action="/rankings/addscores.asp" method="post">
          <input type="hidden" name="pvar" value="SearchTour">
          <input type="hidden" name="sMemberID" value="<%=sMemberID%>">
          <input type="submit" style="width:9em"value="No">
        </form>
        </td><tr></table>
        
        <%                
        WriteFooter

    end if

  end if ' --- Tour ID Found ---

END IF ' ---- Check to see if the search fields have been used. ---


END SUB












Sub WriteHeader
%>
<HTML>
<HEAD><TITLE>Add Scores</TITLE>
</HEAD>

<BODY BGCOLOR="#FFFFFE" Text="#0A0D0A" LINK="#375AE2" VLINK="#36566D" ALINK="#3E85BB">
<style TYPE="text/css">
<!--  A:link {text-decoration: none; color:#375AE2}  A:visited {text-decoration: none; color:#375AE2}  A:active {text-decoration: none}   A:hover {text-decoration: ; color:#3E85BB; }-->
</style>
<%
End Sub

Sub WriteFooter
%>
<hr>
</BODY>
</HTML>
<%
End Sub



%>










