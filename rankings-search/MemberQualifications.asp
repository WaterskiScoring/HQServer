<!--#include file="settingsHQ.asp"-->
<!--#include file="tools_registration.asp"-->
<!--#include file="tools_include.asp"-->
<%


DefineTRAStyles 


Dim sMemberID, sTourID, TQList, pvar

Dim sMembCity, sMembState, sFullName, sMembAge, sTourCOD, sHomoType
Dim sEventSelected, sDivSelected, sCOAStat, sCOAStat2

Dim fRankByCOD, fRankAfterCOD, fScoreByCOD, fLCQHighScore
Dim DecPlaces
Dim ThisFileName, MarkCroneEmail
Dim sPlaceAHead, sPlaceBHead, sPlaceCHead, sPlaceDHead

ThisFileName="MemberQualifications.asp"

'MarkCroneEmail = "mark@productdesign-biz.com"
MarkCroneEmail = "RankingsErrors@usawaterski.org"





sMemberID=TRIM(Request("sMemberID"))
sTourID=LEFT(TRIM(Request("sTourID")),6)
sDivSelected=TRIM(Request("sDivSelected"))
sEventSelected=TRIM(Request("sEventSelected"))
sLeagueSelected=TRIM(Request("sLeagueSelected"))
pvar=TRIM(Request("pvar"))


'response.write("<br>sMemberID="&sMemberID)
'response.write("<br>sTourID="&sTourID)
'response.write("<br>sDivSelected="&sDivSelected)
'response.write("<br>sEventSelected="&sEventSelected)

'IF sMemberID="800073189" THEN response.write("ABOVE - sTourID="&sTourID)
'IF sMemberID="800073189" THEN response.write("<br>sLeagueSelected="&sLeagueSelected)

IF sLeagueSelected<>"" AND  LCASE(sLeagueSelected)<>"none" AND sTourID="" THEN 
	sTourID=sLeagueSelected
END IF


'response.write("<br>Under Reset - sTourID="&sTourID)






'IF sMemberID="800073189" THEN 
'	response.write("<br>MIDDLE - sTourID="&sTourID)
'	response.write("<br>sLeagueSelected="&sLeagueSelected)
'END IF



'WriteIndexPageHeader
WriteIndexPageHeader_NoMenu



IF sTourID="" AND sMemberID<>"" THEN

	' --- Determine tournament parameters ---
	' --- Returns list of tournaments where member has qualifications ---
	SET rs=Server.CreateObject("ADODB.recordset")
	sSQL = "SELECT TOP 1 TourID FROM "&RegQualifyTableName
	sSQL = sSQL +" WHERE MemberID='"&sMemberID&"'"
	rs.open sSQL, SConnectionToTRATable

	IF NOT(rs.eof) THEN 
		sTourID=LEFT(rs("TourID"),6)
	ELSE
		'--- Display message
		NoTourFound
	END IF


ELSEIF sMemberID="" THEN
	NoMemberFound

ELSE
	' --- In tools_registration.asp ---
	DefineTourVariables_New

	' --- Determines EVENTS and EVENT NAMES for this tourament - in tools_include.asp ---
	RegistrationEventsOffered (sTSptsGrpID)

	' --- Defines all member variables  ---
	DefineMemberVars

	' --- Returns tournament date ---
	SET rs=Server.CreateObject("ADODB.recordset")
	sSQL = "SELECT ST.TName, ST.TDateS, ST.TDateE, LT.COD FROM "&SanctionTableName&" AS ST" 
	sSQL = sSQL + " JOIN "&LeagueTableName&" AS LT ON LEFT(ST.TournAppID,6)=LEFT(LT.QualifyTour,6)"
	sSQL = sSQL + " WHERE LEFT(ST.TournAppID,6)='"&LEFT(sTourID,6)&"'" 
	rs.open sSQL, SConnectionToTRATable

	IF NOT rs.EOF THEN
		sTourName=rs("TName")
		sTDateS=rs("TDateS")
		'sTDateS=CDate("06/01/2008")
		sTDateE=rs("TDateE")
		sTourCOD=rs("COD")
	ELSE
		sTDateS=Date
	END IF

'IF sMemberID="800073189" THEN response.write("<br>BELOW - sTourID="&sTourID)

	' --- Displays the Member and Tournament info in with drop downs ---
	DisplayDropOptions

	' --- Displays the qualifications by event for this member ---	
	ListEventQualifications
END IF











WriteIndexPageFooter



' ------------------------------------------------------------------------------------------------------------------
' ----  END OF MAIN PROGRAM  ----
' ------------------------------------------------------------------------------------------------------------------




' -------------------
  SUB NoMemberFound
' -------------------

%>
<br><br><br>
<center><font size="<%=fontsize4%>" color="red"><b><i>A Member has not been selected or is not found in Master file</i></b></font></center>
<br>

     <form action="http://www.usawaterski.org" method="post">
	<center><input type="submit" value="Return to Menu"></center>
     </form>
<%	

response.end

END SUB



' -------------------
  SUB NoTourFound
' -------------------

%>
<br><br><br>
<center><font size="<%=fontsize4%>" color="red"><b><i>The tournament has not been selected or is not found in Master file</i></b></font></center>
<br>

     <form action="http://www.usawaterski.org" method="post">
	<center><input type="submit" value="Return to Menu"></center>
     </form><%	

END SUB



' ----------------------
  SUB DisplayDropOptions
' ----------------------

Dim NewsPageNum
NewsPageNum="FAQ_MembQual"


' --- Determine tournament parameters ---
sSQL = "SELECT LT.* FROM "&LeagueTableName&" AS LT"
sSQL = sSQL + " WHERE LEFT(LT.QualifyTour,6)='"&LEFT(sTourID,6)&"'"
SET rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable

sHomoType=""
IF NOT rs.eof THEN 
	sHomoType=rs("HomoType")
	sTourCOD=rs("COD")
END IF



%>

<TABLE class="innertable" align="center" height=125px width="100%">
<TR>
  <% ' --- Skier Name, MemberID, City/St and Age  --- %> 
   
  <th colspan=8 align ="center">
	<font size=<%=fontsize4%> color="#FFFFFF">&nbsp;&nbsp;<b>Qualifications Status</b></font>
	<font size=<%=fontsize2%> color="#FFFFFF">&nbsp;&nbsp;<b>ver 1.0</b></font>
  </th>

</TR>

<form action="<%=ThisFileName%>" method="post">
<TR>
  <td colspan=6 align ="left" style="border-style:none;">
	<font size=2 color="<%=TextColor2%>"><a title="TourID = <%=sTourID%>"><b><I>Tournament/League</I></b></a></font>
    <br><%
	BuildRegTourDropDown %>
    <br>
	<font size=<%=fontsize2%> color="<%=TextColor2%>">&nbsp;<b><%=sTourCity%>, <%=sTourState %></b></font>
	&nbsp;&nbsp;&nbsp;<font size=<%=fontsize2%> color="<%=TextColor2%>">&nbsp;<b><%=sTDateS%><% IF TRIM(sTDateE)<>"" THEN response.write("-")%><%=sTDateE%></b></font>
    <br><br>
	<font size=4 COlOR="<%=TextColor2%>"><b><I><a title="MemberID: <%=sMemberID%>">&nbsp;<%=sFullName%></a></I></b></font>
    <br>
	<font size=<% =fontsize2 %> COlOR="<%=TextColor2%>"><b>&nbsp;<%=sMembCity%>, <%=sMembState %></b></font>
	&nbsp;&nbsp;&nbsp;<font size=<% =fontsize2 %> COlOR="<%=TextColor2%>"><b>&nbsp;&nbsp;&nbsp;Age: <%=sMembAge%></b></font>
    <br>
  </td>	
  <td align=center colspan=2 style="border-style:none;">
	<font size=<%=fontsize2%> color="<%=TextColor2%>">&nbsp;<b>Cut-Off Date<br>(COD)</font>
	<br>
	<font size=<%=fontsize2%> color="<%=TextColor3%>">&nbsp;<b><%=sTourCOD%></b></font>
	<br><br>
        <input type="submit" style="width:9em" value="Update Display" title="Submit and reset this form">

	<input type="hidden" name="sMemberID" value="<%=sMemberID%>"> 
	<input type="hidden" name="sEventSelected" value="<%=sEventSelected%>"> 
	<input type="hidden" name="sDivSelected" value="<%=sDivSelected%>"> 
	<input type="hidden" name="pvar" value="<%=pvar%>"> 
  </td>
</TR>
</form>

<TR>
  <td align=center colspan=2 style="border-style:none;">
   	<a href="mailto:<%=markcroneemail%>?subject=Qualifications Message for MemberID: <%=sMemberID%> - TourID:<%=sTourID%>" title="Click here to Email problems or recommendations">Report Errors or Feedback</a>
  </td>
  <td align=center colspan=6 style="border-style:none;">
	<form action="/rankings/news/FAQ_MembQual.htm" method="post" target="_blank">
	    <input type="submit" style="width:9em" value="FAQ/Tips"
		title="Tips and Explanations and Insights and &#13;Answers to Frequently Asked Questions">
        </form>
  </td>
</TR>

</TABLE><%


END SUB

' -----------------------------
  SUB ListEventQualifications
' -----------------------------


'IF Session("AdminMenuLevel")>=50 THEN
'	response.write("<br>sHomoType = "&sHomoType)
'END IF


SELECT CASE sHomoType
   CASE "A"
	sPlaceAHead="Prev<br>Natls"
	sPlaceBHead="This<br>Regls"
	sPlaceCHead="N/A"
	sPlaceDHead="N/A"
   CASE "B"
	sPlaceAHead="Prev<br>Natls"
	sPlaceBHead="Prev<br>Regls"
	sPlaceCHead="This<br>States"
	sPlaceDHead="Other"
   CASE "C"
	sPlaceAHead="Place A"
	sPlaceBHead="Place B"
	sPlaceCHead="Place C"
	sPlaceCHead="Place D"
END SELECT


' --- Returns all records for this MemberID where 

sSQL = "SELECT RQ.*, RT.Reg_Ski, LQ.COA FROM "&RegQualifyTableName&" AS RQ"

sSQL = sSQL + " JOIN "&RankTableName&" AS RT ON RT.Event=RQ.Event AND RT.Div=RQ.Div AND RT.MemberID=RQ.MemberID AND (NOT RT.SC_1 IS NULL OR RT.Event='O')"  
sSQL = sSQL + "   AND RT.SkiYearID=1" 

sSQL = sSQL + " JOIN "&LeagueTableName&" AS LT ON LEFT(LT.QualifyTour,6)=LEFT(RQ.TourID,6)"

sSQL = sSQL + " JOIN "&LeagueQfyTableName&" AS LQ ON LQ.LeagueID=LT.LeagueID AND LQ.Event=RQ.Event AND LQ.Div=RQ.Div"

sSQL = sSQL + " WHERE RQ.MemberID='"&sMemberID&"' AND LEFT(RQ.TourID,6)='"&LEFT(sTourID,6)&"'"

sSQL = sSQL + " ORDER BY RQ.Div,"

sSQL = sSQL + " CASE WHEN RQ.Event='S' THEN '1' WHEN RQ.Event='T' THEN '2' WHEN RQ.Event='J' THEN '3' WHEN RQ.Event='O' THEN '4' END" 

SET rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable

%>

<%


Dim sCurrDiv

   IF NOT rs.eof THEN	

	sCurrDiv=rs("Div")

	%><TABLE class="innertable" align=center width="100%"><%
	' --- Displays the qualifications headings
	DisplayHeading


	DO WHILE NOT rs.eof


'Y=999
'IF sMemberID="300043260" AND Session("AdminMenuLevel") THEN 
'	response.write(rs("QfyByMElite"))
'END IF


		sQfyLocked=""
		sQfyPend=""
		sCOAStat="currently"
		sCOAStat2="current"
		IF DateDiff("d", Date, sTourCOD)<=0 THEN
			sCOAStat="locked in at"
			sCOAStat2="locked in"
		END IF

		' --- Qualified by Elite Ranking ---
		' --- Note:  Since Elite Ranking flags are calculated considering the expiration date of the last performance
		' --- 		(QualThru) and the qualifications are assumed to be needed at the time of the tournament
		' --- 		then a person's qualifications will be locked if a flag is found in the RegQfy table for this tournament
		
		sQfyByElite="--"
		sQfyByElite="--"

 	IF Session("AdminMenuLevel")>=50 THEN
		IF rs("QfyByOElite")=true AND rs("QfyByMElite")=true THEN
			sQfyByElite="Open/MM"
			'sQfyLocked=true
		ELSEIF rs("QfyByOElite")=true THEN
			sQfyByElite="Open"
			'sQfyLocked=true
		ELSEIF rs("QfyByMElite")=true THEN
			sQfyByElite="MM"
			'sQfyLocked=true
		END IF
	END IF

		' --- Qualification by RANKING exceeding COA  ---
		' --- Flag is true and the current date is AFTER COD ---
		IF rs("QfyByRankByCOD")=true AND DateDiff("d", Date, sTourCOD)<=0 THEN 
			sQfyByRankByCODStatus="OK"
			sQfyByRankByCODColor=textcolor2
			sQfyLocked=true
		' --- Flag is true and BEFORE COD 
		ELSEIF rs("QfyByRankByCOD")=true AND DateDiff("d", Date, sTourCOD)>0 THEN 
			sQfyByRankByCODStatus="Pending"
			sQfyByRankByCODColor="orange"
			sQfyPend=true
		' --- NOT Qualified by Level and BEFORE COD ---
		ELSEIF rs("QfyByRankByCOD")=false AND DateDiff("d", Date, sTourCOD)>0 THEN 
			sQfyByRankByCODStatus="--"
			sQfyByRankByCODColor=textcolor2

		' --- NOT Qualified by Level and AFTER COD but qualified by other means ---
		ELSEIF rs("QfyByRankByCOD")=false AND DateDiff("d", Date, sTourCOD)<=0 AND (TRIM(rs("QfyStatus"))="Qualified" OR TRIM(rs("QfyStatus"))="QFY_RPR") THEN 
			sQfyByRankByCODStatus="--"
			sQfyByRankByCODColor=textcolor2

		ELSE
			sQfyByRankByCODStatus="NCQ"
			sQfyByRankByCODColor=textcolor3
		END IF


		' --- Qualified by 3rd  ---
		' --- Flag is true and the current date is AFTER COD ---
		IF rs("QfyBy3rdEvt")=true  AND DateDiff("d", Date, sTourCOD)<=0 THEN 
			sQfyBy3rdEvtStatus="OK"
			sQfyBy3rdEvtColor=textcolor2
			sQfyLocked=true

		' --- Qualified by 3rd and BEFORE COD ---
		ELSEIF rs("QfyBy3rdEvt")=true  AND DateDiff("d", Date, sTourCOD)>0 THEN 
			sQfyBy3rdEvtStatus="Pending"
			sQfyBy3rdEvtColor="orange"
			sQfyPend=true
		' --- NOT Qualified by 3rd and BEFORE COD ---
		ELSEIF rs("QfyBy3rdEvt")=false AND DateDiff("d", Date, sTourCOD)>0 THEN 
			sQfyBy3rdEvtStatus="--"
			sQfyBy3rdEvtColor=textcolor2
		ELSE
			sQfyBy3rdEvtStatus="--"
			sQfyBy3rdEvtColor=textcolor2
		END IF
	
		' --- OVERALL qualifications ---
		' --- Flag is true and the current date is AFTER COD ---
		IF rs("QfyByOverall")=true AND DateDiff("d", Date, sTourCOD)<=0 THEN 
			sQfyByOverLevelStatus="OK"
			sQfyByOverLevelColor=textcolor2
			sQfyLocked=true
		' --- OVERALL and BEFORE COD ---
		ELSEIF rs("QfyByOverall")=true AND DateDiff("d", Date, sTourCOD)>0 THEN 
			sQfyByOverLevelStatus="Pending"
			sQfyByOverLevelColor="orange"
			sQfyPend=true
		' --- NOT Qualified by Overall and BEFORE COD ---
		ELSEIF rs("QfyByOverall")=false AND DateDiff("d", Date, sTourCOD)>0 THEN 
			sQfyByOverLevelStatus="--"
			sQfyByOverLevelColor=textcolor2
		ELSE
			sQfyByOverLevelStatus="--"
			sQfyByOverLevelColor=textcolor1
		END IF


		' --- Qualification LCQ by SCORE  ---
		' --- Flag is true and the current date is AFTER COD ---
		IF rs("QfyByScrAfter")=true AND DateDiff("d", Date, sTourCOD)<=0 THEN 
			sQfyByScrAfterStatus="OK"
			sQfyByScrAfterColor=textcolor2
			sQfyLocked=true
		' --- Flag is false and the current date is AFTER COD ---
		ELSEIF rs("QfyByScrAfter")=false AND DateDiff("d", Date, sTourCOD)<=0 THEN 
			sQfyByScrAfterStatus="--"
			sQfyByScrAfterColor=textcolor2
		ELSEIF DateDiff("d", Date, sTourCOD)>0 THEN 
			sQfyByScrAfterStatus="--"
			sQfyByScrAfterColor=textcolor2
		ELSEIF DATE<sTDateS THEN
			sQfyByScrAfterStatus="--"
			sQfyByScrAfterColor=textcolor2
		ELSE
			sQfyByScrAfterStatus="--"
			sQfyByScrAfterColor=textcolor3
		END IF


		' --- Qualification by LCQ By RANK ---
		' --- Flag is true and the current date is AFTER COD ---
		IF rs("QfyByRankAfter")=true AND DateDiff("d", Date, sTourCOD)<=0 THEN 
			sQfyByRankAfterStatus="OK"
			sQfyByRankAfterColor=textcolor2
			sQfyLocked=true
		' --- Flag is false and the current date is AFTER COD ---
		ELSEIF rs("QfyByRankAfter")=false AND DateDiff("d", Date, sTourCOD)<=0 THEN 
			sQfyByRankAfterStatus="--"
			sQfyByRankAfterColor=textcolor2
		ELSEIF DateDiff("d", Date, sTourCOD)>0 THEN 
			sQfyByRankAfterStatus="--"
			sQfyByRankAfterColor=textcolor2
		ELSEIF DATE<sTDateS THEN
			sQfyByRankAfterStatus="--"
			sQfyByRankAfterColor=textcolor2
		ELSE
			sQfyByRankAfterStatus="--"
			sQfyByRankAfterColor=textcolor3
		END IF


		' --- Qualification by LCQ By OVERALL (Rank or Score) ---
		' --- Flag is true and the current date is AFTER COD ---
		IF (rs("QfyOverLCQByScr")=true OR rs("QfyOverLCQByRank")=true) AND DateDiff("d", Date, sTourCOD)<=0 THEN 
			sQfyOverLCQAllStatus="OK"
			sQfyOverLCQAllColor=textcolor2
			sQfyLocked=true
		' --- Flag is false and the current date is AFTER COD ---
		ELSEIF (rs("QfyOverLCQByScr")=false AND rs("QfyOverLCQByRank")=false)  AND DateDiff("d", Date, sTourCOD)<=0 THEN 
			sQfyOverLCQAllStatus="--"
			sQfyOverLCQAllColor=textcolor2
		ELSEIF DateDiff("d", Date, sTourCOD)>0 THEN 
			sQfyOverLCQAllStatus="--"
			sQfyOverLCQAllColor=textcolor2
		ELSEIF DATE<sTDateS THEN
			sQfyOverLCQAllStatus="--"
			sQfyOverLCQAllColor=textcolor2
		ELSE
			sQfyOverLCQAllStatus="---"
			sQfyOverLCQAllColor=textcolor3
		END IF



		' --- Qualified by Placement A ---
		IF rs("QfyByPlaceA")=true THEN 
			sQfyByPlaceAStatus="OK"
			sQfyByPlaceAColor=textcolor2
			sQfyLocked=true
		' --- NOT Qualified by Placement A and BEFORE COD ---
		ELSEIF rs("QfyByPlaceA")<>true AND DateDiff("d", Date, sTourCOD)>0 THEN 
			sQfyByPlaceAStatus="--"
			sQfyByPlaceAColor=textcolor2
		ELSE
			sQfyByPlaceAStatus="--"
			sQfyByPlaceAColor=textcolor2
		END IF

		' --- Qualified by Placement B ---
		IF rs("QfyByPlaceB")=true THEN 
			sQfyByPlaceBStatus="OK"
			sQfyByPlaceBColor=textcolor2
			sQfyLocked=true
		' --- NOT Qualified by Placement B and BEFORE COD ---
		ELSEIF rs("QfyByPlaceB")<>true AND DateDiff("d", Date, sTourCOD)>0 THEN 
			sQfyByPlaceBStatus="--"
			sQfyByPlaceBColor=textcolor2
		ELSE
			sQfyByPlaceBStatus="--"
			sQfyByPlaceBColor=textcolor2
		END IF

		' --- Qualified by Placement C ---
		IF rs("QfyByPlaceC")=true THEN 
			sQfyByPlaceCStatus="OK"
			sQfyByPlaceCColor=textcolor2
			sQfyLocked=true
		' --- NOT Qualified by Placement C and BEFORE COD ---
		ELSEIF rs("QfyByPlaceC")<>true AND DateDiff("d", Date, sTourCOD)>0 THEN 
			sQfyByPlaceCStatus="--"
			sQfyByPlaceCColor=textcolor2
		ELSE
			sQfyByPlaceCStatus="--"
			sQfyByPlaceCColor=textcolor2
		END IF

		' --- Qualified by Placement D ---
		IF rs("QfyByPlaceD")=true THEN 
			sQfyByPlaceDStatus="OK"
			sQfyByPlaceDColor=textcolor2
			sQfyLocked=true
		' --- NOT Qualified by Placement D and BEFORE COD ---
		ELSEIF rs("QfyByPlaceD")<>true AND DateDiff("d", Date, sTourCOD)>0 THEN 
			sQfyByPlaceDStatus="--"
			sQfyByPlaceDColor=textcolor2
		ELSE
			sQfyByPlaceDStatus="--"
			sQfyByPlaceDColor=textcolor2
		END IF


		' --- Qualified by Overall Participation Only in STATES ---
		IF rs("QfyByState_3EvPart")=true THEN 
			sQfyBy_AnyOverall_InStates_Status="OK"
			sQfyBy_AnyOverall_InStates_Color=textcolor2
			sQfyLocked=true
		' --- NOT Qualified by Placement D and BEFORE COD ---
		ELSEIF rs("QfyByPlaceD")<>true AND DateDiff("d", Date, sTourCOD)>0 THEN 
			sQfyBy_AnyOverall_InStates_Status="--"
			sQfyBy_AnyOverall_InStates_Color=textcolor2
		ELSE
			sQfyBy_AnyOverall_InStates_Status="--"
			sQfyBy_AnyOverall_InStates_Color=textcolor2
		END IF




		' --- Determines color if any LOCKED qualifications ---
'		IF sQfyLocked=true THEN
'			sCurrentStatus="Qualified"		
'			sCurrentStatusColor=textcolor2
'		ELSEIF sQfyLocked<>true AND sQfyPend=true THEN
'			sCurrentStatus="Pending"		
'			sCurrentStatusColor="orange"
'		ELSE
'			sCurrentStatus="NCQ"		
'			sCurrentStatusColor="red"
'		END IF	

		' --- NEW 8-2-2008 ---
		IF TRIM(rs("QfyStatus"))="Qualified" OR TRIM(rs("QfyStatus"))="QFY-RPR" THEN
			sCurrentStatusColor=textcolor2
		ELSEIF TRIM(rs("QfyStatus"))="Pending" THEN
			sCurrentStatusColor="orange"
		ELSEIF TRIM(rs("QfyStatus"))="NCQ" THEN
			sCurrentStatusColor="red"
		END IF


		FOR EvtNo=1 TO TotEv
			IF TRIM(rs("Event"))=TRIM(sTEvent(EvtNo)) THEN sThisEventName=sTEventName(EvtNo)
			IF TRIM(rs("Event"))="O" THEN sThisEventName="Overall"
		NEXT

		SELECT CASE TRIM(rs("Event"))
		  CASE "S"
			sDecPlaces=2
			sEventSuffix="buoys"
		  CASE "T"
			sDecPlaces=0
			sEventSuffix="points"
		  CASE "J"
			sDecPlaces=1
			sEventSuffix="feet"
		  CASE "O"
			sDecPlaces=1
			sEventSuffix="points"
		END SELECT

		IF rs("RankByCOD")>=0 THEN fRankByCOD=formatnumber(rs("RankByCOD"),sDecPlaces) ELSE fRankByCOD="--" 
		IF rs("RankAfterCOD")>=0 THEN fRankAfterCOD=formatnumber(rs("RankAfterCOD"),sDecPlaces) ELSE fRankAfterCOD="--" 
		IF rs("ScoreByCOD")>=0 THEN fScoreByCOD=formatnumber(rs("ScoreByCOD"),sDecPlaces) ELSE fScoreByCOD="--" 
		IF rs("ScoreAfterCOD")>0 THEN fLCQHighScore=formatnumber(rs("ScoreAfterCOD"),sDecPlaces) ELSE fLCQHighScore="--" 
		IF rs("COA")>=0 THEN fCOA=formatnumber(rs("COA"),sDecPlaces) ELSE fCOA="Unknown" 



'IF sMemberID="900134058" THEN 
'	response.write("<br>TESTING VARIABLE VALUES WITH DOUG's RECORD - Mark Crone<br>")
'	response.write("<br>fRankAfterCOD="&fRankAfterCOD&"<br>")
'	response.write("<br>"&rs("RankByCOD"))
'	response.write("<br>rs(ScoreAfterCOD) = "&rs("ScoreAfterCOD"))
'	response.write("<br>fScoreByCOD="&fScoreByCOD&"<br>")
'	response.write(rs("ScoreAfterCOD")>=0)
'	response.write("<br>sQfyByRankAfterCOD="&sQfyByRankAfterCOD)
'END IF

		' ---Defines TITLE of statement ---

		Other_Title="LCQ Ranking is highest Ranking achieved between "&sTourCOD&" (COD) and "&sTDateS&" the tournament start date."

		EliteStatusTitle="Elite Status is calculated according to AWSA Rule 3.03.  The qualification for this specific tournament is based on the expiration date of the Elite status.   The Qualified Thru date must be after the Start Date of the tournament."
		CurrentStatusTitle="If the Current Status column indicates 'Qualified', then your qualification is locked in.  If the Current Status column indicates 'QFY-RPR', then you are qualified but Regionals Participation is Required. Otherwise, if this column indicates 'Pending' your qualification will not lock in until the COD for this tournament. NCQ indicates Not Currently Qualified (see below)."

		RankValueTitle="See Rank Value explanation below.  &#13;The "&rs("Div")&" - "&sThisEventName&" Cut-Off-Average (COA) is "&sCOAStat&"&nbsp; at "&fCOA&"&nbsp;"&sEventSuffix&" for the  &#13;"&sTourName&"."
		RankOverCOATitle="If 'OK' then you are qualified based on your Ranking Score being higher than the COA as of the COD (for this tournament).  Status of this and other 'Qualification by Level' methods will display as PENDING until "&sTourCOD&"."
		ThirdEventTitle="If 'OK' then you are qualified in two events and meet the 3rd event qualification standard in the 3rd event."
		OverallTitle="If 'OK' then you are qualified based on your Ranking Level in Overall."


		LCQHighRankValueTitle="Highest Ranking Value achieved during the period beginning with the Cut-off-Date (COD) and the Start Date of the tournament."
		LCQRankOverCOATitle="IF 'OK' then you are qualified based on the 'LCQ by Ranking Value' qualification method."
		LCQHighScoreTitle="Highest Score achieved in any of the specified LCQ tournaments during the period following the Cut-off-Date (COD)."
		LCQScoreOverCOATitle="IF 'OK' then you are qualified based on the 'LCQ by Score' qualification method."
		LCQOverallTitle="IF 'OK' then you are qualified based on the 'LCQ by Score' OR 'LCQ By Rank' based on your Overall Performance"


		' --- If the DIVISION changes then reprint the Heading ---
		IF rs("Div")<> sCurrDiv THEN 
			%><TABLE class="innertable" align=center width="100%"><%
			' --- Displays the qualifications headings
			DisplayHeading
			sCurrDiv=rs("Div")
		END IF


		' --- Displays Event Name, Div, Elite Status, COA and QFY information for each EVENT - one per line ---
		%>
		<TR>
		<TD><font size="<%=fontsize2%>"><b><%=sThisEventName%></b></font></TD>
		<TD align="center"><font size="<%=fontsize2%>"><%=rs("Div")%></font></TD>
		<TD align="center"><font size="<%=fontsize2%>"><a title="<%=EliteStatusTitle%>"><%=sQfyByElite%></a></font></TD>

		<TD align="center"><font size="<%=fontsize2%>" color="<%=sCurrentStatusColor%>"><a title="<%=CurrentStatusTitle%>"><b><%=rs("QfyStatus")%></b></a></font></TD>

		<TD align="center"><font size="<%=fontsize2%>" color="<%=TextColor2%>">
			<a title="<%=sCOAStat%>"><%=fCOA%></a></font>
		</TD>
		<TD align="center"><font size="<%=fontsize2%>" color="<%=TextColor2%>">
			<a title="<%=RankValueTitle%>"><%=fRankByCOD%></a></font>
		</TD>
		<TD align="center"><font size="<%=fontsize2%>" color="<%=sQfyByRankByCODColor%>">
			<a title="<%=RankOverCOATitle%>"><%=sQfyByRankByCODStatus%></a></font>
		</TD>
		<TD align="center"><font size="<%=fontsize2%>" color="<%=sQfyBy3rdEvtColor%>">
			<a title="<%=ThirdEventTitle%>"><%=sQfyBy3rdEvtStatus%></a></font></TD>
		<TD align="center"><font size="<%=fontsize2%>" color="<%=sQfyByOverLevelColor%>">
			<a title="<%=OverallTitle%>"><%=sQfyByOverLevelStatus%></a></font></TD><%


		' --- Displays the LCQ qualification status info ---
		IF DATE<sTourCOD THEN %>
			<TD align="center"><font size="<%=fontsize2%>" color="<%=TextColor2%>"><a title="<%=LCQHighRankValueTitle%>">--</a></font></TD>
			<TD align="center"><font size="<%=fontsize2%>" color="<%=TextColor2%>"><a title="<%=LCQRankOverCOATitle%>">--</a></font</TD>
			<TD align="center"><font size="<%=fontsize2%>" color="<%=TextColor2%>"><a title="<%=LCQHighScoreTitle%>">--</a></font></TD>
			<TD align="center"><font size="<%=fontsize2%>" color="<%=TextColor2%>"><a title="<%=LCQScoreOverCOATitle%>">--</a></font></TD>
			<TD align="center"><font size="<%=fontsize2%>" color="<%=TextColor2%>"><a title="<%=LCQOverallTitle%>">--</a></font></TD><%
	
		ELSE %>
			<TD align="center"><font size="<%=fontsize2%>" color="<%=TextColor2%>"><a title="<%=LCQHighRankValueTitle%>"><%=fRankAfterCOD%></a></font></TD>
			<TD align="center"><font size="<%=fontsize2%>" color="<%=sQfyByRankAfterColor%>"><a title="<%=LCQRankOverCOATitle%>"><%= sQfyByRankAfterStatus %></a></font></TD>
			<TD align="center"><font size="<%=fontsize2%>" color="<%=sQfyByScrAfterColor%>"><a title="<%=LCQHighScoreTitle%>"><%=fLCQHighScore %></a></font></TD>
			<TD align="center"><font size="<%=fontsize2%>" color="<%=sQfyByScrAfterColor%>"><a title="<%=LCQScoreOverCOATitle%>"><%=sQfyByScrAfterStatus%></a></font></TD>
			<TD align="center"><font size="<%=fontsize2%>" color="<%=sQfyByScrAfterColor%>"><a title="<%=LCQOverallTitle%>"><%=sQfyOverLCQAllStatus%></a></font></TD><%
		END IF  
	
		' --- Sets the popup title for each placement depending on the type of tournament being displayed ---
		SELECT CASE sHomoType
		   CASE "A"
			PlaceATitle="IF 'OK' then you are qualified based on placement in Previous NATIONAL Championships"
			PlaceBTitle="IF 'OK' then you are qualified based on placement in Current REGIONAL Championships"
			PlaceCTitle="Not Active for Nationals"
			PlaceDTitle="Not Active for Nationals"
		   CASE "B"
			PlaceATitle="IF 'OK' then you are qualified based on placement in Previous National Championships if applicable for this Region"
			PlaceBTitle="IF 'OK' then you are qualified based on placement in Previous Regional Championships if applicable for this Region"
			PlaceCTitle="IF 'OK' then you are qualified based on placement in Current Year State tournament if applicable for this Region"
			PlaceDTitle="IF 'OK' then you are qualified based on placement in Junior Development tournaments (applies to South Central Regionals only). "
		   CASE "C"
			PlaceATitle="Not Active"
			PlaceBTitle="Not Active"
			PlaceCTitle="Not Active"
			PlaceDTitle="Not Active"
		END SELECT 

		' --- Displays the placement qualification status ---
		%>
		<TD align="center"><font size="<%=fontsize2%>" color="<%=sQfyByPlaceAColor%>"><a title="<%=PlaceATitle%>"><%=sQfyByPlaceAStatus%></a></font></TD>
		<TD align="center"><font size="<%=fontsize2%>" color="<%=sQfyByPlaceBColor%>"><a title="<%=PlaceBTitle%>"><%=sQfyByPlaceBStatus%></a></font></TD>
		<TD align="center"><font size="<%=fontsize2%>" color="<%=sQfyByPlaceCColor%>"><a title="<%=PlaceCTitle%>"><%=sQfyByPlaceCStatus%></a></font></TD>
		<TD align="center"><font size="<%=fontsize2%>" color="<%=sQfyByPlaceDColor%>"><a title="<%=PlaceDTitle%>"><%=sQfyByPlaceDStatus%></a></font></TD>

		<TD align="center"><font size="<%=fontsize2%>" color="<%=sQfyBy_AnyOverall_InStates_Color%>"><a title="<%=Overall_InStates_Title%>"><%=sQfyBy_AnyOverall_InStates_Status%></a></font></TD>		
		</TR><%

		rs.movenext
	LOOP 

   END IF





	%>
    </TD>
  </TR>
</TABLE> 
<br>
<TABLE align=center width="850px">
<TR>
  <td align="left" colspan=2>
	<font size="<%=fontsize3%>" color="red"><i><b>NOTICE:</font>
	<br>
	<font size="<%=fontsize2%>" color="<%=textcolor1%>">1) National Championships requires regional participation (except Open Divs).
	<br>2) Regional Qualification see '10-week Rule' 4.03(a). All Regionals listed to accomodate (OOR) Out of [home] Region skiers.
	<br>3) Placement data for 'Overall' is not available. If you are qualified by Overall placement, bring proof of qualification.
	<br>4) Hover your mouse over each field to see additional information about that item.</b></i></font>
	<br><br>
  </td>
</TR>
<TR>
  <td align="left" colspan=2>
	<font size="<%=fontsize2%>" color="<%=TextColor1%>"><b><u>Current Status</u></b> - If the Current Status column indicates 'Qualified', then your qualification is locked in.  Otherwise, if this column indicates 'Pending'<br> 
	 your qualification will not lock in until you have meet the qualifications as of the COD for this tournament, or you have achieved a qualification by another method. <br> A status of Qfy-RO indicates member is qualified and
	 an administrative Override of the regional participation requirement was made. </font>
	<br><br>
	<font size="<%=fontsize2%>" color="<%=TextColor1%>"><b><u>NCQ (Not Currently Qualified)</u></b> - If your Current Status field shows NCQ, then at the present time your Ranking is below the current Level (COA) for this tournament.  Without raising your Ranking you will not be qualified for this tournament using this qualification method. No Rank Value will display if your current status is NCQ.</font>
	<br><br>
	<font size="<%=fontsize2%>" color="<%=TextColor1%>"><b><u>QFY-RPR (Regionals Participation Required)</u></b> - If your Current Status shows QFY-RPR, then you have qualified by some method, but final qualification is dependent on participation in a Regional tournament.</font>
	<br><br>
	<font size="<%=fontsize2%>" color="<%=TextColor1%>"><b><u>COA (Cut-Off Average)</u></b> - This is determined by the Ranking [score] of the last person of the 'Ranking Level' needed for qualification to this Tournament (and Division).  The Cut-off-Date becomes 'locked' on the Cut-off-Date for this tournament and in turn locks in anyone meeting the qualification by 'Rank Value' methods.</font>
	<br><br>
	<font size="<%=fontsize2%>" color="<%=TextColor1%>"><b><u>COD (Cut-Off Date)</u></b> - Date after which the Cut-off-Average for this tournament is 'Locked' and does not change.</font>
	<br><br>
	<font size="<%=fontsize2%>" color="<%=TextColor1%>"><b><u>LCQ (Last Chance Qualification)</u></b> - After the Cut-off-Date, 'LCQ' qualification occurs when, either: a) your Rank Value [Ranking Score] exceeds the 'Locked' COA for this tournament or b) when you get a score (in any of the specified qualifying tournaments that exceeds the 'Locked' COA for this tournament .  Also see the Regional Web Sites since <u>LCQ Qualification methods vary by Tournament and Region.</u></font>
	<br><br>
	<font size="<%=fontsize2%>" color="<%=TextColor1%>"><b><u>Rank Value</u></b> - Prior to the Cut-off-Date, this column (4) displays your current 'Ranking Value.'  After the COD for this tournament, this column does not update and retains your 'Ranking Score' as of the Cut-off-Date.</font>
	<br><br>
	<font size="<%=fontsize2%>" color="<%=TextColor1%>"><b><u>High Rank Value</u></b> - After the Cut-off-Date, this LCQ field tracks the highest 'Ranking Value' you achieved during the period between the Cut-off-Date and the tournament date.  <b>Note: SC Regionals does not use this qualification method.</b></font>
	<br><br>
	<font size="<%=fontsize2%>" color="<%=TextColor1%>"><b><u>High Score</u></b> - After the Cut-off-Date, this tracks the highest score you achieved in any of the qualifying tournaments between the Cut-off-Date and the tournament date.  <b>Note: E Regionals does not use this qualification method.<br></b></font>
  </td>
</TR>
</TABLE>
<br><br><%


END SUB



' ------------------
  SUB DisplayHeading
' ------------------
'IF Session("AdminMenuLevel")>=50 THEN
'	response.write("<br>sPlaceAHead = "&sPlaceAHead)
'END IF



%>
<br>

<TABLE class="innertable" align=center width="100%">
<TR>
	<th colspan=3 align="center"><font size="<%=fontsize2%>" color="#FFFFFF">&nbsp;</font></th>
	<th colspan=2 align="center"><font size="<%=fontsize2%>" color="#FFFFFF">&nbsp;</font></th>
	<th colspan=4 align="center"><font size="<%=fontsize2%>" color="#FFFFFF"><br>Qualification By Cut-Off-Date</font></th>
	<th colspan=5 align="center"><font size="<%=fontsize2%>" color="#FFFFFF"><br>Last Chance Qualification (LCQ)</font></th>
	<th colspan=5 align="center"><font size="<%=fontsize2%>" color="#FFFFFF"><br>Participation or Place Qualify</font></th>
</TR>

<TR>
	<th align="center"><font size="<%=fontsize2%>" color="#FFFFFF">Event</font></th>
	<th align="center" width="40px"><font size="<%=fontsize2%>" color="#FFFFFF">Div</font></th>
	<th align="center" width="60px"><font size="<%=fontsize2%>" color="#FFFFFF">Elite<br>Status</font></th>
	<th align="center" width="60px"><font size="<%=fontsize2%>" color="red"><b>Current<br>Status<br>Summary</b></font></th>
	<th align="center" width="60px"><font size="<%=fontsize2%>" color="#FFFFFF">Tourn<br>COA</font></th>
	<th align="center" width="60px"><font size="<%=fontsize2%>" color="#FFFFFF">Rank<br>Value</font></th>
	<th align="center" width="50px"><font size="<%=fontsize2%>" color="#FFFFFF">Event<br>Rank<br>Over<br>COA</font></th>
	<th align="center" width="50px"><font size="<%=fontsize2%>" color="#FFFFFF">3rd Evt<br>Rank<br>Over<br>COA-3</font></th>
	<th align="center" width="60px"><font size="<%=fontsize2%>" color="#FFFFFF">Overall<br>Rank<br>Over<br>COA</font></th>

	<th align="center" width="60px"><font size="<%=fontsize2%>" color="#FFFFFF">High<br>Rank<br>Value</font></th>
	<th align="center" width="50px"><font size="<%=fontsize2%>" color="#FFFFFF">Event<br>Rank<br>Over<br>COA</font></th>
	<th align="center" width="60px"><font size="<%=fontsize2%>" color="#FFFFFF">High<br>Score</font></th>
	<th align="center" width="50px"><font size="<%=fontsize2%>" color="#FFFFFF">Score<br>Over<br>COA</font></th>
	<th align="center" width="60px"><font size="<%=fontsize2%>" color="#FFFFFF">Overall<br>Rank or<br>Score<br>Over<br>COA</font></th>

	<th align="center" width="50px"><font size="<%=fontsize2%>" color="#FFFFFF"><%=sPlaceAHead%></font></th>
	<th align="center" width="50px"><font size="<%=fontsize2%>" color="#FFFFFF"><%=sPlaceBHead%></font></th>
	<th align="center" width="50px"><font size="<%=fontsize2%>" color="#FFFFFF"><%=sPlaceCHead%></font></th>
	<th align="center" width="50px"><font size="<%=fontsize2%>" color="#FFFFFF"><%=sPlaceDHead%></font></th>

	<th align="center" width="50px"><font size="<%=fontsize2%>" color="#FFFFFF">This<br>State<br>Overall<br>Score</font></th>
</TR>
<%



END SUB




' ------------------------
  SUB BuildRegTourDropDown
' ------------------------

' --- Returns list of tournaments where member has qualifications ---
SET rs=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT RQ.TourID, LT.LeagueName"
sSQL = sSQL +" FROM "&RegQualifyTableName&" AS RQ, "&SanctionTableName&" AS ST, "&LeagueTableName&" AS LT,"
sSQL = sSQL +" "&SkiYearTableName&" AS SY"
sSQL = sSQL +" WHERE LEFT(RQ.TourID,6)=LEFT(ST.TournAppID,6) AND LEFT(LT.QualifyTour,6)=LEFT(RQ.TourID,6)"
sSQL = sSQL +" AND SY.DefaultYear='1' AND RIGHT(SY.SkiYearName,2)=LEFT(RQ.TourID,2)" 
sSQL = sSQL +" AND RQ.MemberID='"&sMemberID&"' AND LT.Status<>'X'"
sSQL = sSQL +" GROUP BY RQ.MemberID, RQ.TourID, LT.LeagueName"
sSQL = sSQL +" ORDER BY LT.LeagueName DESC" 

rs.open sSQL, SConnectionToTRATable

'response.write(sSQL)
'response.end

' --- NEED TO ADD RESTRICTION TO SKI YEAR

'IF NOT(rs.eof) AND TRIM(sTourID)="" THEN sTourID=LEFT(rs("TourID"),6)


TQList="Select"
TQNameList="Select"
DO WHILE NOT rs.eof
	TQList=TQList+","&LEFT(rs("TourID"),6)
	TQNameList=TQNameList+","&rs("LeagueName")
	rs.movenext
LOOP

EnteredArray = Split(TQList,",")  
EnteredArray2 = Split(TQNameList,",")  

IF ThisFileName="MemberQualifications.asp" THEN  %>
	<select name="sTourID" style="width:25em"><%
ELSE %>
	<select name="sTourID" style="width:25em" onchange="submit()" ><%
END IF 

  FOR kvar = 1 TO UBOUND(EnteredArray)

    IF LEFT(sTourID,6) = LEFT(EnteredArray(kvar),6) THEN
	response.write("<option value = """&sTourID&""" SELECTED>"&EnteredArray2(kvar)&"</option>")
    ELSE
	response.write("<option value = """&EnteredArray(kvar)&""">"&EnteredArray2(kvar)&"</option>")
    END IF
  NEXT  %>
</select><%



END SUB



' -----------------------
  SUB BuildRegisteredList
' -----------------------

' --- Returns list of tournaments where member has qualifications ---
SET rs=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT TourID FROM "&RegGenTableName&" AS ST WHERE ST.MemberID='"&sMemberID&"'" 
' --- NEED TO ADD RESTRICTION TO SKI YEAR

'response.write("<br>"&sSQL)
'response.end

rs.open sSQL, SConnectionToTRATable


TQList="("
DO WHILE NOT rs.eof
	IF TRIM(TQList)<>"(" THEN TQList=TQList&", '"&LEFT(rs("TourID"),6)&"'"

	IF TRIM(TQList)="(" THEN TQList="('"&LEFT(rs("TourID"),6)&"'"

	rs.movenext
LOOP
TQList=TQList+")"


END SUB



' -----------------------------
  SUB DefineMemberVars
' -----------------------------

	
' ----------------------------------------------------------------------------------------------
' --------------------------  Define MEMBER Variables from MemberTrak --------------------------
' ----------------------------------------------------------------------------------------------

set rsMemb=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT TOP 1 * FROM "&MemberTableName
'sSQL = sSQL + " JOIN "&MemberTypeTableName&" ON "&MemberTypeTableName&".MembershipTypeID = "&MemberTableName&".MembershipTypeCode"
sSQL = sSQL + " WHERE PersonIDwithCheckDigit = "&sqlclean(sMemberID)


rsMemb.open sSQL, sConnectionToTRATable, 3, 1

IF rsMemb.eof THEN
	'--- Subroutine in this program
	NoMemberFound
END IF

sFirstName = rsMemb("FirstName")
sLastName = rsMemb("LastName")
sFullName = rsMemb("FirstName")&" "&rsMemb("LastName")

sMembCity = rsMemb("City")
sMembState = rsMemb("State")
sMembSex = rsMemb("Sex")
'sMembTypeID = rsMemb("MembershipTypeID")
sMembBirth = rsMemb("Birthdate")

'sCanSkiTour = rsMemb("CanSkiInTournaments")
'sMembTypeCode = rsMemb("TypeCode")
sEffectiveto = rsMemb("Effectiveto")

'Session("sCanSkiTour") = rsMemb("CanSkiInTournaments")

' ++++++++++  TEST VARIABLE  ++++++++++
'Session("sCanSkiTour") = false
' +++++++++++++++++++++++++++++++++++++




' ---- Needs both Member and Tournament information to define sMembAge  ----
IF TRIM(sMemberID)<>"" THEN
	sAgeDate=sTDateS
	IF TRIM(sTDateS)="" THEN sAgeDate=DATE 
	sMembAge = AgeAtDate(sAgeDate, sMemberID)		' Function finds Member Age
END IF


END SUB
%>


