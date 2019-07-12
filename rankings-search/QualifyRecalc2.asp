<!--#include file="settingsHQ.asp"-->
<%

Server.ScriptTimeout = 360 


Dim TypeAList, TypeBList, TypeCList, TypeDList
Dim LeagueList, TourList, CODList, MinClassList
Dim CurrLeagueID, CurrTourID, sHomoType, PastCOD, sCOD, sCOAMinClass, sTDateS
Dim LeagueArray, TourArray, CODArray, MinClassArray
Dim kvar, sSQL, rsType
Dim DisplayTestMarkers

DisplayTestMarkers="Y"

ty=1
IF ty=2 THEN
	Set myMail=CreateObject("CDO.Message")
	myMail.Subject="Flag Indicating Reached Member Qualify Recalc - Auto Start"
	myMail.From="USAWS.Rankings@USAWaterSki.ORG"
	myMail.To="AWSATechDude@comcast.net; mark@productdesign-biz.com"
	myMail.Send
	set myMail=nothing
END IF


OpenCon



' -----------------------------------------------------------------------------------
' --- Finds list of all LeagueID's that have tournament qualifications associated ---
' -----------------------------------------------------------------------------------

sSQL = "SELECT LeagueID, QualifyTour, COD, COAMinClass, HomoType, ST.TDateS, LT.Status"
sSQL = sSQL + " FROM "&LeagueTableName&" AS LT"
sSQL = sSQL + " JOIN "&SanctionTableName&" AS ST" 
sSQL = sSQL + " 	ON LEFT(ST.TournAppID,6)=LEFT(LT.QualifyTour,6)"
sSQL = sSQL + "  WHERE (QualifyTour<>''" 
sSQL = sSQL + "  AND LEFT(LT.QualifyTour,2)>='"&CInt(RIGHT(CStr(Year(DATE)),2))&"')" 
'sSQL = sSQL + "  OR  QualifyTour='12S999'"
'sSQL = sSQL + "  WHERE  QualifyTour='12S999'"
sSQL = sSQL + "  AND  QualifyTour='13S999'"
sSQL = sSQL + "  ORDER BY QualifyTour"

'response.write(sSQL)
'response.end
	 
SET rsTList=Server.CreateObject("ADODB.recordset")
rsTList.open sSQL, SConnectionToTRATable



IF rsTList.eof THEN 
		response.write("EOF")
		response.end
END IF


IF DisplayTestMarkers="Y" THEN response.write("<br><br>Line 30 "&CDate(Now()))
'response.end



' ----------------------------------
' --- Loops therough all Leagues ---
' ----------------------------------

DO WHILE NOT rsTList.eof


		'CurrTourID=TourArray(kvar)
		'CurrLeagueID=LeagueArray(kvar)

		CurrTourID=TRIM(rsTList("QualifyTour"))
		CurrLeagueID=rsTList("LeagueID")
		sHomoType=rsTList("HomoType")
		sCOD=rsTList("COD")
		sCOAMinClass=rsTList("COAMinClass")
		sTDateS=rsTList("TDateS")
		sStatus=TRIM(rsTList("Status"))


		' ----------------------------------------------------------------------------------------------------------------------------------
		' --- Defines list of current year tournaments from the LeagueTours table from which placement and score qualification can occur ---
		' ----------------------------------------------------------------------------------------------------------------------------------
		DefineTypeTourList


		' --- Testing only
		' --- Updates the Qualification by 3 Event Participation in STATES
		'IF TRIM(TypeCList)<>"()" THEN Update_3EventPart_InStates

		IF DisplayTestMarkers="Y" THEN	
				response.write("<br><br>CurrLeagueID="&CurrLeagueID)
				response.write("<br>CurrTourID="&CurrTourID)
				response.write("<br>sCOD="&sCOD)
				response.write("<br>sStatus="&sStatus)		
				response.write("<br>TypeCList="&TypeCList)
				response.write("<br>Test=")
				response.write(DateDiff("d", DATE, sCOD)>=0 AND sStatus="A")
			'	response.end
		END IF


	' -------------------------------------------------------------------
	' --- Checks if the current date is before the COD of this League ---
	' -------------------------------------------------------------------
	IF DateDiff("d", DATE, sCOD)>=0 AND sStatus="A" THEN


jumphere=true
IF jumphere=true THEN

			' --- Updates current COA for the CurrLeague in LeagueQualify from Rankings table
			IF DisplayTestMarkers="Y" THEN response.write("<br>Line 112 - Before-SetLeagueQualifyCOA")
			SetLeagueQualifyCOA



			' --- Starts Table Fresh for this CurrTourID ---
			EmptyTableContents
			PastCOD=false

			' --- Checks type of tournament to determine if all records or only those registered
			SELECT CASE sHomoType
					' --- Nationals and Regionals ---
					CASE "A", "B"

							' --- Creates Qfy records for ALL members in ranking table ---
							IF DisplayTestMarkers="Y" THEN response.write("<br><br>Line 127 - CreateRegQfyRecordsNoPreRegister Date: "&NOW)
							CreateRegQfyRecordsNoPreRegister

					' --- State Tournaments ---
					CASE "C"
							IF DisplayTestMarkers="Y" THEN response.write("<br><br>Line 132 - CreateRegQfyRecords_StateResidency Date: "&NOW)
							CreateRegQfyRecords_StateResidency

			END SELECT

			' --- Calculates Qualifications by LEVEL - but only runs prior to COD ---
			IF DisplayTestMarkers="Y" THEN response.write("<br><br>Line 138 - QualifyByRankByCOD Date: "&NOW)
			QualifyByRankByCOD


			' --- Calculates Qualifications by 3rd Event
			IF DisplayTestMarkers="Y" THEN response.write("<br><br>Line 143 - QualifyBy3rdEvent Date: "&NOW)
			QualifyBy3rdEvent

			
			' --- Calculates Qualifications by Overall 
			IF DisplayTestMarkers="Y" THEN response.write("<br><br>Line 148 - QualifyByRankByCOD_FromOverallRank Date: "&NOW)
			QualifyByRankByCOD_FromOverallRank


			' --- Updates the Qualification by placement based on TourType ---
			IF DisplayTestMarkers="Y" THEN response.write("<br><br>Line 153 - UpdatePlacementA-D: "&NOW)
			IF TRIM(TypeAList)<>"()" THEN UpdatePlacementA
			IF TRIM(TypeBList)<>"()" THEN UpdatePlacementB
			IF TRIM(TypeCList)<>"()" THEN UpdatePlacementC
			IF TRIM(TypeDList)<>"()" THEN UpdatePlacementD


			' --- Updates the Qualification by 3 Event Participation in STATES
			IF DisplayTestMarkers="Y" THEN response.write("<br><br>Line 161 - Update_3EventPart_InStates Date: "&NOW)
			IF TRIM(TypeCList)<>"()" THEN Update_3EventPart_InStates


			' --- Runs the Elite qualifications test ---
			IF DisplayTestMarkers="Y" THEN response.write("<br><br>Line 166 - QualifyByElite Date: "&NOW)
			QualifyByElite


			' --- Performs a CURRENT STATUS update for all qualifications methods
			IF DisplayTestMarkers="Y" THEN response.write("<br><br>Line 171 - QfyStatusUpdateNEW Date: "&NOW)
			QfyStatusUpdateNEW

END IF

IF DisplayTestMarkers="Y" THEN response.write("<br><br>Line 180 - QfyStatusUpdateNEW_08072013 Date: "&NOW)
QfyStatusUpdateNEW_08072013
						
			' --- Counts and Displays the numbers of updates ----
			'		CountUpdates



	' ---------------------------------------------------------------------------
	' --- ELSE Condition is when DATE is past COD 		---
	' --- Continue to update until START DATE of tournament ---
	' --- 4-2-2012 - Changed to add requirement for AFTER COD and Active status
	' ---------------------------------------------------------------------------

	ELSEIF DateDiff("d", DATE, sCOD)<0 AND DateDiff("d", DATE, sTDateS)>=0 AND sStatus="A" THEN


RunThisPart=false
IF RunThisPart=true THEN


			' --- Checks type of tournament to determine if all records or only those registered ---
			SELECT CASE sHomoType
				' --- Nationals and Regionals ---
				CASE "A", "B"
							' --- Creates Qfy records for ALL members in ranking table ---
							IF DisplayTestMarkers="Y" THEN response.write("<br><br>Line 197 - CreateRegQfyRecordsNoPreRegister Date: "&CurrTourID&" - "&NOW)
							CreateRegQfyRecordsNoPreRegister

					' --- State Tournaments ---
					CASE "C"
							IF DisplayTestMarkers="Y" THEN response.write("<br><br>Line 202 - CreateRegQfyRecords_StateResidency Date: "&CurrTourID&" - "&NOW)
							CreateRegQfyRecords_StateResidency
			END SELECT


			' --- Updates LCQ QfyByRankAfter by achieving Ranking greater than locked COA at any time between COD and tournament.
			IF DisplayTestMarkers="Y" THEN response.write("<br><br>Line 208 - QualifyByRankAfterCOD - Date: "&CurrTourID&" - "&NOW)
			QualifyByRankAfterCOD

			' --- Updates LCQ by achieving a Score greater than the locked COA (after COD and before tournament) in a) specified class or b) specified tournament list.
			IF DisplayTestMarkers="Y" THEN Response.write("<br><br>Line 213 - Qualify_LCQByScore "&CDate(Now()))
			Qualify_LCQByScore

			' --- Updates LCQ OVERALL [as an event] by achieving a Score greater than the locked COA (after COD and before tournament) in a) specified class or b) specified tournament list.
			IF DisplayTestMarkers="Y" THEN Response.write("<br><br>Line 218 - Qualify_LCQByScore_Overall "&CDate(Now()))
			Qualify_LCQByScore_Overall

			' --- Updates Event flags based on LCQ OVERALL qualifications ---
			IF DisplayTestMarkers="Y" THEN Response.write("<br><br>Line 221 - QualifyLCQByOverall_AllTypes "&CDate(Now()))
			QualifyLCQByOverall_AllTypes

			' --- Determines any required participation ---
			IF DisplayTestMarkers="Y" THEN Response.write("<br><br>Line 225 - RequiredParticipation "&CDate(Now()))
			RequiredParticipation

			PastCOD=true


			' --- Updates the Qualification by placement based on TourType ---
			IF DisplayTestMarkers="Y" THEN Response.write("<br><br>Line 232 - QualifyLCQByOverall_AllTypes "&CDate(Now()))
			IF TRIM(TypeAList)<>"()" THEN UpdatePlacementA
			IF TRIM(TypeBList)<>"()" THEN UpdatePlacementB
			IF TRIM(TypeCList)<>"()" THEN UpdatePlacementC
			IF TRIM(TypeDList)<>"()" THEN UpdatePlacementD

			' --- Updates the Qualification by 3 Event Participation in STATES
			IF DisplayTestMarkers="Y" THEN Response.write("<br><br>Line 236 - Update_3EventPart_InStates "&CDate(Now()))
			IF TRIM(TypeCList)<>"()" THEN Update_3EventPart_InStates


			' --- Runs the Elite qualifications test ---
			IF DisplayTestMarkers="Y" THEN Response.write("<br><br>Line 244 - QualifyByElite "&CDate(Now()))
			QualifyByElite

			' --- Performs a CURRENT STATUS update for all qualifications methods
			IF DisplayTestMarkers="Y" THEN Response.write("<br><br>Line 248 - QfyStatusUpdateNEW "&CDate(Now()))
			QfyStatusUpdateNEW

END IF

			IF DisplayTestMarkers="Y" THEN response.write("<br><br>Line 261a - QfyStatusUpdateNEW_08072013 Date: "&NOW)
			QfyStatusUpdateNEW_08072013


'IF DisplayTestMarkers="Y" THEN response.write("<br><br>Line 264 - ApplyAdminOverrides Date: "&NOW)
'ApplyAdminOverrides



			' --- Counts and Displays the numbers of updates ----
			'		CountUpdates

	END IF	

	' --- If date is less than Start Date of tournament and the status of the tournament is Active then write to history table 
	IF DateDiff("d", DATE, sTDateS)>=0 AND sStatus="A" THEN
			' --- Writes a summary of the current values to the History table ---	
			WriteToHistoryTable
	END IF	

	rsTList.movenext


LOOP


CloseCon
rsTList.close

ty=1
IF ty=2 THEN
	Set myMail=CreateObject("CDO.Message")
	myMail.Subject="Flag Indicating Reached Member Qualify Recalc - Automatic Done!"
	myMail.From="USAWS.Rankings@USAWaterSki.ORG"
	myMail.To="AWSATechDude@comcast.net"
	myMail.Send
	set myMail=nothing
END IF


%>
<br><br><br>
<center>
<font size=3 >Qualifications Recalc Complete</font>
<br><br>
<form action="/rankings/defaultHQ.asp" method="post">
	<input type="submit" name="Done" Value="Return to Menu">
</form>
</center>
<%



' ----------------------------------------------------------------------------------------------------------------
'--- END OF MAIN PROGRAM
' ----------------------------------------------------------------------------------------------------------------







' ------------------------
  SUB EmptyTableContents
' ------------------------

sSQL = "DELETE FROM "&RegQualifyTableName
sSQL = sSQL + " WHERE Left(TourID,6) = '"&LEFT(CurrTourID,6)&"'"
con.execute(sSQL)

END SUB




' -------------------------
  SUB SetLeagueQualifyCOA
' -------------------------

' --- Sets the CURRENT COA from each division/event in the LeagueQualify table ---

sSQL = "UPDATE LQ1 SET LQ1.COA=MS.RTCOA"   
sSQL = sSQL + " FROM "&LeagueQfyTableName&" AS LQ1"
sSQL = sSQL + " JOIN"
sSQL = sSQL + "    (SELECT RT.Event, RT.Div, MIN(RT.RankScore) AS RTCOA"
sSQL = sSQL + "    FROM "&RankTableName&" RT"
sSQL = sSQL + "    JOIN "
sSQL = sSQL + "    		(SELECT Level_A, Event, Div FROM "&LeagueQfyTableName&" WHERE LeagueID='"&CurrLeagueID&"') AS LQ" 
sSQL = sSQL + "    ON LQ.Event=RT.Event AND LQ.Div=RT.Div" 

' --- Changed 7-4-2008 to include overall settings ---
sSQL = sSQL + "    WHERE RIGHT(RT.AWSA_Rat,1)=LQ.Level_A AND RT.SkiYearID='1'"

sSQL = sSQL + "    GROUP BY RT.Div, RT.Event) AS MS"
sSQL = sSQL + " ON MS.Event=LQ1.Event AND MS.Div=LQ1.Div AND LQ1.LeagueID='"&CurrLeagueID&"'"

'response.write("<br><br>Line 339<br>"&sSQL)
'response.end

con.execute(sSQL)


END SUB



' -------------------------------------
  SUB CreateRegQfyRecordsNoPreRegister
' -------------------------------------

' --- Inserts the current LEVEL qualifications from Rankings into RegisterQualify for a Tournament with no Pre-Registration criteria --- 
' --- Those already in table as of COD do not get their QfyByRankByCOD value updated after the COD ---

sSQL = "INSERT INTO "&RegQualifyTableName&" (TourID, MemberID, Event, Div)"

sSQL = sSQL + " SELECT '"&CurrTourID&"', RT.MemberID, RT.Event, RT.Div"
sSQL = sSQL + " 	FROM "&RankTableName&" AS RT"

' --- Removed tables 7-8-2013 ---
'sSQL = sSQL + "	JOIN "&LeagueQfyTableName&" AS LQ ON LQ.LeagueID='"&CurrLeagueID&"' AND RT.Event=LQ.Event AND RT.Div=LQ.Div"
'sSQL = sSQL + "	JOIN "&MemberTableName&" AS MT ON RT.MemberID=MT.PersonIDWithCheckDigit"
'sSQL = sSQL + "	JOIN "&MemberShortTableName&" AS MT ON CAST(RIGHT(RT.MemberID,8) AS INT)=MT.PersonID"
'sSQL = sSQL + "	LEFT JOIN "&RegionTableName&" AS RG ON lower(MT.state) = lower(RG.state)"

sSQL = sSQL + "	WHERE RT.RankScore IS NOT NULL AND RT.SkiYearID=1"


' --- Designed to limit the insert to only those members not already in the table, so once COD is reached you no longer delete all records 
' --- 	   and rebuild the table before updating the various fields.
sSQL = sSQL + "	AND RT.MemberID NOT IN (SELECT MemberID"
sSQL = sSQL + "		FROM "&RegQualifyTableName&" AS RQ1"
sSQL = sSQL + "		WHERE LEFT(TourID,6)='"&CurrTourID&"' AND MemberID=RT.MemberID AND Event=RT.Event AND Div=RT.Div)"

'response.write("<br><br>Line 376<br>"&sSQL)
'response.end

con.execute(sSQL)





END SUB


' ---------------------------------------
  SUB CreateRegQfyRecords_StateResidency
' ---------------------------------------

' --- Inserts the current LEVEL qualifications from Rankings into RegisterQualify for a Tournament with no Pre-Registration criteria --- 
' --- Those already in table as of COD do not get their QfyByRankByCOD value updated after the COD ---

sSQL = "INSERT INTO "&RegQualifyTableName&" (TourID, MemberID, Event, Div)"
sSQL = sSQL + " SELECT '"&CurrTourID&"', RT.MemberID, RT.Event, RT.Div"
sSQL = sSQL + " 	FROM "&RankTableName&" AS RT"
sSQL = sSQL + "	LEFT JOIN "&LeagueQfyTableName&" AS LQ ON LQ.LeagueID='"&CurrLeagueID&"' AND RT.Event=LQ.Event AND RT.Div=LQ.Div"
sSQL = sSQL + " JOIN "&LeagueTableName&" AS LT ON LEFT(LT.QualifyTour,6)='"&CurrTourID&"'" 
sSQL = sSQL + "	JOIN "&MemberShortTableName&" AS MT ON CAST(RIGHT(RT.MemberID,8) AS INT)=MT.PersonID"
sSQL = sSQL + "	LEFT JOIN "&RegionTableName&" AS RG ON lower(MT.state) = lower(RG.state)"
sSQL = sSQL + "	WHERE RT.RankScore IS NOT NULL AND RT.SkiYearID=1"

' --- Limits to those with residency in State of League ---
sSQL = sSQL + " AND lower(MT.State)=lower(LT.State)"	

' --- Designed to limit the insert to only those members not already in the table, so once COD is reached you no longer delete all records 
' --- 	   and rebuild the table before updating the various fields.
sSQL = sSQL + "	AND RT.MemberID NOT IN (SELECT MemberID"
sSQL = sSQL + "		FROM "&RegQualifyTableName&" AS RQ1"
sSQL = sSQL + "		WHERE LEFT(TourID,6)='"&CurrTourID&"' AND MemberID=RT.MemberID AND Event=RT.Event AND Div=RT.Div)"

con.execute(sSQL)

response.write("<br>Line 411<br>"&sSQL)
response.end

END SUB



' ----------------------------
  SUB QualifyByRankByCOD
' ----------------------------

' --- Inserts the current RANKSCORE from Rankings into RegisterQualify  --- 

sSQL = "UPDATE RQ"
sSQL = sSQL + " SET RQ.RankByCOD=RQ1.RankScore"
sSQL = sSQL + ", QfyByRankByCOD=CASE WHEN RIGHT(isnull(RQ1.AWSA_Rat,0),1)>=RQ1.Level_A OR RQ1.Level_A IS NULL THEN 1 ELSE 0 END"

sSQL = sSQL + " FROM "&RegQualifyTableName&" AS RQ" 
sSQL = sSQL + "   JOIN "

sSQL = sSQL + "     (SELECT RT.MemberID, RT.Event, RT.Div, RT.AWSA_Rat, LQ.Level_A, RT.RankScore"
sSQL = sSQL + " 			FROM "&RankTableName&" AS RT"

sSQL = sSQL + "		INNER JOIN "
sSQL = sSQL + "		   (SELECT Level_A, Event, Div FROM " &LeagueQfyTableName&" WHERE LeagueID='"&CurrLeagueID&"') AS LQ"
sSQL = sSQL + "		ON RT.Event=LQ.Event AND RT.Div=LQ.Div"

sSQL = sSQL + "	WHERE RT.SkiYearID=1"
sSQL = sSQL + ") AS RQ1"

sSQL = sSQL + " ON RQ1.MemberID=RQ.MemberID AND RQ1.Event=RQ.Event AND RQ1.Div=RQ.Div"
sSQL = sSQL + "	WHERE RQ.TourID='"&CurrTourID&"'" 


'response.write("<br><br>Line 449<br>"&sSQL)
'response.end

con.execute(sSQL)



END SUB




' ---------------------------------------
  SUB QualifyByRankByCOD_FromOverallRank
' ---------------------------------------


' --- Updates QfyByOverall flag in each event when the member's OVERALL [Overall as an event] Level is above the COA Level for Overall By COD --- 

sSQL = "UPDATE RQ"
sSQL = sSQL + " SET QfyByOverall=1" 

sSQL = sSQL + " FROM "&RegQualifyTableName&" AS RQ"

sSQL = sSQL + " JOIN ("
sSQL = sSQL + " SELECT "
sSQL = sSQL + " 	R1.MemberID, R1.TourID, R1.Div, R1.Event, "
sSQL = sSQL + " 	RIGHT(isnull(RT.AWSA_Rat,0),1) AS MembOVLev, "
sSQL = sSQL + " 	LeagueID, isnull(Level_A,0) AS LQOVLevA"
sSQL = sSQL + " FROM "&RegQualifyTableName&" AS R1"

sSQL = sSQL + " 	JOIN "&RankTableName&" AS RT "
sSQL = sSQL + " 		ON RT.MemberID=R1.MemberID AND RT.Event=R1.Event AND RT.Div=R1.Div AND SkiYearID=1"
sSQL = sSQL + " 	JOIN "&LeagueQfyTableName&" AS LQ "
sSQL = sSQL + " 		ON LQ.Event=RT.Event AND LQ.Div=RT.Div "
sSQL = sSQL + " 	WHERE LQ.LeagueID='"&CurrLeagueID&"' AND R1.Event='O' AND LEFT(R1.TourID,6)='"&LEFT(CurrTourID,6)&"') AS RQ1"
			
sSQL = sSQL + " ON RQ.MemberID=RQ1.MemberID  AND LEFT(RQ.TourID,6)=LEFT(RQ1.TourID,6) AND RQ.Div=RQ1.Div"

sSQL = sSQL + " WHERE LEFT(RQ.TourID,6)='"&LEFT(CurrTourID,6)&"' AND RQ.Event<>'O'" 
sSQL = sSQL + " 		AND RQ1.MembOVLev>=RQ1.LQOVLevA"

'response.write("<br><br>Line 491<br>"&sSQL)
'response.end
' --- NOTE: 7-9-2013 Took 5 seconds to complete ---

con.execute(sSQL)




END SUB




'-----------------------
  SUB QualifyBy3rdEvent
'-----------------------

' --- Question as to whether this is actually applicable or allowed under present rules ---

sSQL = "UPDATE 	RQ SET RQ.QfyBy3rdEvt=1"

sSQL = sSQL + "	FROM "&RegQualifyTableName&" AS RQ"
	 
' --- Slalom 
sSQL = sSQL + " JOIN ("
sSQL = sSQL + " SELECT" 
sSQL = sSQL + " 	R1.MemberID, R1.TourID, R1.Div, R1.Event, RT.SkiYearID,"
sSQL = sSQL + " 	QfyByRankByCOD AS SLQfy, RIGHT(isnull(RT.AWSA_Rat,0),1) AS MembSLLev," 
sSQL = sSQL + " 	LeagueID, isnull(Level_A,0) AS LQSLLevA, LevelBy3rdEvt AS LQSL3rd"  
sSQL = sSQL + " FROM "&RegQualifyTableName&" R1"
sSQL = sSQL + " 	INNER JOIN "
sSQL = sSQL + " 		(SELECT MemberID, Event, Div, SC_1, SkiYearID, AWSA_Rat FROM "&RankTableName&" WHERE SC_1 IS NOT NULL AND SkiYearID=1) AS RT" 
sSQL = sSQL + " 	ON RT.MemberID=R1.MemberID AND RT.Event=R1.Event AND RT.Div=R1.Div"
sSQL = sSQL + " 	INNER JOIN "
sSQL = sSQL + " 		(SELECT LeagueID, Level_A, LevelBy3rdEvt, Event, Div FROM "&LeagueQfyTableName&" WHERE LeagueID='"&CurrLeagueID&"') AS LQ" 
sSQL = sSQL + " 	ON LQ.Event=RT.Event AND LQ.Div=RT.Div" 
sSQL = sSQL + " 	WHERE R1.Event='S' AND LEFT(R1.TourID,6)='"&LEFT(CurrTourID,6)&"') AS RQ1"
sSQL = sSQL + " ON RQ.MemberID=RQ1.MemberID  AND LEFT(RQ.TourID,6)=LEFT(RQ1.TourID,6) AND RQ.Div=RQ1.Div"


' --- Tricks
sSQL = sSQL + " JOIN ("
sSQL = sSQL + " SELECT" 
sSQL = sSQL + " 	R1.MemberID, R1.TourID, R1.Div, R1.Event,	RT.SkiYearID,"
sSQL = sSQL + " 	QfyByRankByCOD AS TRQfy, RIGHT(isnull(RT.AWSA_Rat,0),1) AS MembTRLev," 
sSQL = sSQL + " 	LeagueID, isnull(Level_A,0) AS LQTRLevA, LevelBy3rdEvt AS LQTR3rd"  
sSQL = sSQL + " FROM "&RegQualifyTableName&" R1"
sSQL = sSQL + " 	INNER JOIN "
sSQL = sSQL + " 		(SELECT MemberID, Event, Div, SC_1, SkiYearID, AWSA_Rat FROM "&RankTableName&" WHERE SC_1 IS NOT NULL AND SkiYearID=1) AS RT" 
sSQL = sSQL + " 	ON RT.MemberID=R1.MemberID AND RT.Event=R1.Event AND RT.Div=R1.Div"
sSQL = sSQL + " 	INNER JOIN "
sSQL = sSQL + " 		(SELECT LeagueID, Level_A, LevelBy3rdEvt, Event, Div FROM "&LeagueQfyTableName&" WHERE LeagueID='"&CurrLeagueID&"') AS LQ" 
sSQL = sSQL + " 	ON LQ.Event=RT.Event AND LQ.Div=RT.Div" 
sSQL = sSQL + " 	WHERE R1.Event='T' AND LEFT(R1.TourID,6)='"&LEFT(CurrTourID,6)&"') AS RQ2"
sSQL = sSQL + " ON RQ.MemberID=RQ2.MemberID AND LEFT(RQ.TourID,6)=LEFT(RQ2.TourID,6) AND RQ.Div=RQ2.Div"


'sSQL = sSQL + " JOIN ("
'sSQL = sSQL + " SELECT" 
'sSQL = sSQL + " 	R1.MemberID, R1.TourID, R1.Div, R1.Event," 
'sSQL = sSQL + " 	QfyByRankByCOD AS TRQfy, RIGHT(isnull(RT.AWSA_Rat,0),1) AS MembTRLev," 
'sSQL = sSQL + " 	LeagueID, isnull(Level_A,0) AS LQTRLevA, LevelBy3rdEvt AS LQTR3rd"  
'sSQL = sSQL + " FROM "&RegQualifyTableName&" R1"
'sSQL = sSQL + " 	INNER JOIN "&RankTableName&" AS RT" 
'sSQL = sSQL + " 		ON RT.MemberID=R1.MemberID AND RT.Event=R1.Event AND RT.Div=R1.Div"
'sSQL = sSQL + " 	INNER JOIN "&LeagueQfyTableName&" AS LQ" 
'sSQL = sSQL + " 		ON LQ.Event=RT.Event AND LQ.Div=RT.Div" 
'sSQL = sSQL + " 	WHERE LQ.LeagueID='"&CurrLeagueID&"' AND R1.Event='T' AND SC_1 IS NOT NULL AND SkiYearID=1 AND LEFT(R1.TourID,6)='"&LEFT(CurrTourID,6)&"') AS RQ2"
'sSQL = sSQL + " ON RQ.MemberID=RQ2.MemberID AND LEFT(RQ.TourID,6)=LEFT(RQ2.TourID,6) AND RQ.Div=RQ2.Div"

' --- Jump
sSQL = sSQL + " JOIN ("
sSQL = sSQL + " SELECT" 
sSQL = sSQL + " 	R1.MemberID, R1.TourID, R1.Div, R1.Event, RT.SkiYearID,"
sSQL = sSQL + " 	QfyByRankByCOD AS JUQfy, RIGHT(isnull(RT.AWSA_Rat,0),1) AS MembJULev," 
sSQL = sSQL + " 	LeagueID, isnull(Level_A,0) AS LQJULevA, LevelBy3rdEvt AS LQJU3rd"  
sSQL = sSQL + " FROM "&RegQualifyTableName&" R1"
sSQL = sSQL + " 	INNER JOIN "
sSQL = sSQL + " 		(SELECT MemberID, Event, Div, SC_1, SkiYearID, AWSA_Rat FROM "&RankTableName&" WHERE SC_1 IS NOT NULL AND SkiYearID=1) AS RT" 
sSQL = sSQL + " 	ON RT.MemberID=R1.MemberID AND RT.Event=R1.Event AND RT.Div=R1.Div"
sSQL = sSQL + " 	INNER JOIN "
sSQL = sSQL + " 		(SELECT LeagueID, Level_A, LevelBy3rdEvt, Event, Div FROM "&LeagueQfyTableName&" WHERE LeagueID='"&CurrLeagueID&"') AS LQ" 
sSQL = sSQL + " 	ON LQ.Event=RT.Event AND LQ.Div=RT.Div" 
sSQL = sSQL + " 	WHERE R1.Event='J' AND LEFT(R1.TourID,6)='"&LEFT(CurrTourID,6)&"') AS RQ3"
sSQL = sSQL + " ON RQ.MemberID=RQ3.MemberID AND LEFT(RQ.TourID,6)=LEFT(RQ3.TourID,6) AND RQ.Div=RQ3.Div" 

' --- When two events are above the required COA for the League and the 3rd is above the COA for 3rd event ---
sSQL = sSQL + " WHERE LEFT(RQ.TourID,6)='"&LEFT(CurrTourID,6)&"'"
sSQL = sSQL + " AND "
sSQL = sSQL + " ("
sSQL = sSQL + " 	(RQ2.MembTRLev>=RQ2.LQTRLevA AND RQ3.MembJULev>=RQ3.LQJULevA AND RQ1.MembSLLev>=RQ1.LQSL3rd)"
sSQL = sSQL + " OR" 
sSQL = sSQL + " 	(RQ1.MembSLLev>=RQ1.LQSLLevA AND RQ3.MembJULev>=RQ3.LQJULevA AND RQ2.MembTRLev>=RQ2.LQTR3rd)"
sSQL = sSQL + " OR"
sSQL = sSQL + " 	(RQ1.MembSLLev>=RQ1.LQSLLevA AND RQ2.MembTRLev>=RQ2.LQTRLevA AND RQ3.MembJULev>=RQ3.LQJU3rd)"
sSQL = sSQL + " )"
sSQL = sSQL + " AND isnull(RQ.QfyByRankByCOD,0)<>1"

'response.write("<br><br>Line 591<br>"& sSQL)
'response.end
' NOTE: 7-9-2013 Took 5 seconds to complete on MW Regionals
con.execute(sSQL)

END SUB



' ------------------------------------------------------------------------------------------------------
' ------- THE FOLLOWING SECTION BEGINS THE LCQ CALCULATIONS 	----------------------------------------
' ------------------------------------------------------------------------------------------------------

' ---------------------------
  SUB Qualify_LCQByScore
' ---------------------------

' --- Clears Table since it's OK to rebuild all values from original data ---
sSQL = "UPDATE "&RegQualifyTableName
sSQL = sSQL + " SET QfyByScrAfter=0, ScoreAfterCOD=0 "
sSQL = sSQL + "   WHERE LEFT(TourID,6)='"&LEFT(CurrTourID,6)&"'" 
con.execute(sSQL)


sSQL = "UPDATE RQ1 SET ScoreAfterCOD=MaxScore, QfyByScrAfter=CASE WHEN MS.MaxScore>=MS.COA THEN 1 ELSE 0 END" 
sSQL = sSQL + " FROM "&RegQualifyTableName&" AS RQ1"

' --- JOIN of Main Query that finds the maximum score of each member during period from COD to TDateS of CurrTourID ---
sSQL = sSQL + " JOIN "
sSQL = sSQL + " (SELECT RS.MemberID, RS.Event, RS.Div, MAX(RS.Score) AS MaxScore, LQ.COA"
sSQL = sSQL + " 	FROM "&RawScoresTableName&" AS RS" 

sSQL = sSQL + " 	LEFT JOIN "
sSQL = sSQL + " 	 (SELECT Event, Div, COA FROM "&LeagueQfyTableName&" WHERE LeagueID='"&CurrLeagueID&"') AS LQ" 
sSQL = sSQL + " 		ON LQ.Event=RS.Event AND LQ.Div=RS.Div" 

IF sCOAMinClass="R" THEN
	sSQL = sSQL + " 	WHERE RS.Class IN ('E', 'L', 'R')"
ELSE
	sSQL = sSQL + " 	WHERE RS.Class IN ('C', 'E', 'L', 'R')"
END IF

sSQL = sSQL + " 				AND" 
sSQL = sSQL + " 			LEFT(RS.TourID,6) IN" 
sSQL = sSQL + " 			( SELECT TournAppID FROM "&SanctionTableName
sSQL = sSQL + " 					WHERE (TDateE>(SELECT COD FROM "&LeagueTableName&" WHERE LeagueID='"&CurrLeagueID&"' AND UseLCQScore=1)" 
sSQL = sSQL + " 					AND TDateE<=(SELECT TDateE FROM "&SanctionTableName&" WHERE LEFT(TournAppID,6)='"&CurrTourID&"')" 


' --- Explicitely includes scores from Type B, C and D TourType from LeagueTours table ---
' --- Type A is not used as this implies a score from the previous Nationals ---
IF TRIM(TypeBList)<>"()" THEN
		sSQL = sSQL + " 		OR TournAppID IN "&TypeBList
END IF
IF TRIM(TypeCList)<>"()" THEN
		sSQL = sSQL + " 		OR TournAppID IN "&TypeCList
END IF
IF TRIM(TypeDList)<>"()" THEN
			sSQL = sSQL + " 		OR TournAppID IN "&TypeDList
END IF

' --- Close bracket from select list for TourID ---
sSQL = sSQL + " )"

' --- Close bracket for UPDATE select ---
sSQL = sSQL + " 		)"

sSQL = sSQL + " 	GROUP BY RS.MemberID, RS.Event, RS.Div, LQ.COA) AS MS"
sSQL = sSQL + " ON RQ1.MemberID=MS.MemberID AND RQ1.Event=MS.Event AND RQ1.Div=MS.Div AND LEFT(RQ1.TourID,6)='"&CurrTourID&"'" 
sSQL = sSQL + " WHERE LEFT(RQ1.TourID,6)='"&CurrTourID&"'"

'response.write("<br><br>Line 661<br>"&sSQL)
'response.end

con.execute(sSQL)


END SUB




' ---------------------------
  SUB Qualify_LCQByScore_OLD
' ---------------------------

' --- Clears Table since it's OK to rebuild all values from original data ---
sSQL = "UPDATE "&RegQualifyTableName
sSQL = sSQL + " SET QfyByScrAfter=0, ScoreAfterCOD=0 "
sSQL = sSQL + "   WHERE LEFT(TourID,6)='"&LEFT(CurrTourID,6)&"'" 
con.execute(sSQL)



sSQL = "UPDATE RQ1 SET ScoreAfterCOD=MaxScore, QfyByScrAfter=CASE WHEN MS.MaxScore>=MS.COA THEN 1 ELSE 0 END"
sSQL = sSQL + " FROM "&RegQualifyTableName&" AS RQ1"

' --- JOIN of Main Query that finds the maximum score of each member during period from COD to TDateS of CurrTourID ---
sSQL = sSQL + " JOIN "
sSQL = sSQL + " (SELECT RS.MemberID, RS.Event, RS.Div, MAX(RS.Score) AS MaxScore, LQ.COA"
sSQL = sSQL + " 	FROM "&RawScoresTableName&" AS RS" 
sSQL = sSQL + " 		JOIN "&SanctionTableName&" AS ST" 
sSQL = sSQL + " 			ON LEFT(ST.TournAppID,6)=LEFT(RS.TourID,6)" 
sSQL = sSQL + " 		JOIN "&LeagueQfyTableName&" AS LQ" 
sSQL = sSQL + " 			ON LQ.Event=RS.Event AND LQ.Div=RS.Div" 
sSQL = sSQL + " 		JOIN "&LeagueTableName&" AS LT" 
sSQL = sSQL + " 			ON LT.LeagueID=LQ.LeagueID" 

' --- Uses Tours table to determine which tournaments are used for 
sSQL = sSQL + " 		LEFT JOIN "&LeagueToursTableName&" AS LTOUR" 
sSQL = sSQL + " 			ON LTOUR.LeagueID=LQ.LeagueID AND LEFT(LTOUR.TourID,6)=LEFT(RS.TourID,6)" 


IF sCOAMinClass="R" THEN
	sSQL = sSQL + " 	WHERE ( RS.Class IN ('E', 'L', 'R')"
ELSE
	sSQL = sSQL + " 	WHERE ( RS.Class IN ('C', 'E', 'L', 'R')"
END IF
sSQL = sSQL + " 		AND LQ.LeagueID= '"&CurrLeagueID&"'" 
sSQL = sSQL + " 		AND LT.UseLCQScore=1" 
sSQL = sSQL + " 		AND ST.TDateE>LT.COD" 
sSQL = sSQL + " 		AND ST.TDateE<=( SELECT ST2.TDateE FROM "&SanctionTableName&" AS ST2 WHERE LEFT(ST2.TournAppID,6)='"&CurrTourID&"')"  
sSQL = sSQL + " 		)"


' --- Explicitely includes scores from Type B, C and D TourType from LeagueTours table ---
' --- Type A is not used as this implies a score from the previous Nationals ---
  
IF TRIM(TypeBList)<>"()" THEN
	sSQL = sSQL + " 		OR (LEFT(RS.TourID,6) IN "&TypeBList&" AND LTOUR.UseForLCQScr=1)"
END IF
IF TRIM(TypeCList)<>"()" THEN
	sSQL = sSQL + " 		OR (LEFT(RS.TourID,6) IN "&TypeCList&" AND LTOUR.UseForLCQScr=1)"
END IF
IF TRIM(TypeDList)<>"()" THEN
	sSQL = sSQL + " 		OR (LEFT(RS.TourID,6) IN "&TypeDList&" AND LTOUR.UseForLCQScr=1)" 
END IF
			
sSQL = sSQL + " 	GROUP BY RS.MemberID, RS.Event, RS.Div, LQ.COA) AS MS"
sSQL = sSQL + " ON RQ1.MemberID=MS.MemberID AND RQ1.Event=MS.Event AND RQ1.Div=MS.Div AND LEFT(RQ1.TourID,6)='"&CurrTourID&"'" 
		
sSQL = sSQL + " WHERE LEFT(RQ1.TourID,6)='"&CurrTourID&"'"

response.write("<br><br>Line 734 <br>"&sSQL)
response.end

con.execute(sSQL)

'response.write("<br>"&sSQL)

END SUB







' -------------------------------
  SUB Qualify_LCQByScore_Overall
' -------------------------------

' --- Updates the ScoreAfterCOD value and the QfyByScrAfter flag in OVERALL [overall as an event] --- 
' --- NOTE:  Secondary processing necessary to update each [non overall] event qualifications to reflect qualification status from overall ---

sSQL = "UPDATE RQ1 SET ScoreAfterCOD=MaxScore, QfyByScrAfter=CASE WHEN MS.MaxScore>=MS.COA THEN 1 ELSE 0 END"
sSQL = sSQL + " FROM "&RegQualifyTableName&" AS RQ1"

' --- JOIN of Main Query that finds the maximum score of each member during period from COD to TDateS of CurrTourID ---
sSQL = sSQL + " JOIN "
sSQL = sSQL + " (SELECT RS.MemberID, 'O' AS Event, RS.Div, MAX(RS.TotalOverAll) AS MaxScore, COALESCE(LQ.COA,0) AS COA "

sSQL = sSQL + " 	FROM "&OverAllScoresTableName&" AS RS" 

' --- Join the LeagueQualifications table where the divisions match and Event is Overall ---
sSQL = sSQL + " 		LEFT JOIN" 
sSQL = sSQL + "				(SELECT COA, Div FROM "&LeagueQfyTableName&" WHERE LeagueID= '"&CurrLeagueID&"' AND Event='O') AS LQ" 
sSQL = sSQL + " 		ON LQ.Div=RS.Div" 

' --- If the class requirement for LCQ By Score is E/L/R otherwise all classes above C --- 
IF sCOAMinClass="R" THEN
	sSQL = sSQL + " 	WHERE RS.Class IN ('E', 'L', 'R')"
ELSE
	sSQL = sSQL + " 	WHERE RS.Class IN ('C', 'E', 'L', 'R')"
END IF

sSQL = sSQL + " 				AND" 
sSQL = sSQL + " 			LEFT(RS.TourID,6) IN" 
sSQL = sSQL + " 			( SELECT TournAppID FROM "&SanctionTableName
sSQL = sSQL + " 					WHERE (TDateE>(SELECT COD FROM "&LeagueTableName&" WHERE LeagueID='"&CurrLeagueID&"' AND UseLCQScore=1)" 
sSQL = sSQL + " 					AND TDateE<=(SELECT TDateE FROM "&SanctionTableName&" WHERE LEFT(TournAppID,6)='"&CurrTourID&"')" 


' --- Explicitely includes scores from Type B, C and D TourType from LeagueTours table ---
' --- Type A is not used as this implies a score from the previous Nationals ---
IF TRIM(TypeBList)<>"()" THEN
		sSQL = sSQL + " 		OR TournAppID IN "&TypeBList
END IF
IF TRIM(TypeCList)<>"()" THEN
		sSQL = sSQL + " 		OR TournAppID IN "&TypeCList
END IF
IF TRIM(TypeDList)<>"()" THEN
			sSQL = sSQL + " 		OR TournAppID IN "&TypeDList
END IF
' --- Close bracket from select list for TourID ---
sSQL = sSQL + " )"

' --- Close bracket for UPDATE select ---
sSQL = sSQL + " 		)"
			
sSQL = sSQL + " 	GROUP BY RS.MemberID, RS.Div, LQ.COA) AS MS"
sSQL = sSQL + " ON RQ1.MemberID=MS.MemberID AND RQ1.Event='O' AND RQ1.Div=MS.Div AND LEFT(RQ1.TourID,6)='"&LEFT(CurrTourID,6)&"'" 
		
sSQL = sSQL + " WHERE LEFT(RQ1.TourID,6)='"&LEFT(CurrTourID,6)&"'"


'response.write("<br><br>Line 806<br>"&sSQL)
'response.end

con.execute(sSQL)

END SUB




' ---------------------------
  SUB QualifyByRankAfterCOD
' ---------------------------

' --- Updates RankAfterCOD value and QfyByRankAfter flag in all events and overall [as an event] when Ranking goes above locked COA after COD ---
' --- NOTE:  Secondary processing necessary to update each [non overall] event qualifications to reflect qualification status from overall ---

sSQL = " UPDATE RQ" 
sSQL = sSQL + " SET RankAfterCOD = CASE WHEN RT1.RankScore>RQ1.RankAfterCOD THEN RT1.RankScore ELSE RQ1.RankAfterCOD END,"
sSQL = sSQL + "	QfyByRankAfter=CASE WHEN RT1.RankScore>=RT1.COA THEN 1 ELSE 0 END"

sSQL = sSQL + " 	FROM "&RegQualifyTableName&" AS RQ"
	
' --- Finds League Qualification LEVEL required and the member's current RankScore ---
sSQL = sSQL + " 	JOIN"
sSQL = sSQL + " 		( SELECT RT.MemberID, RT.Event, RT.Div, RT.AWSA_Rat, LQ.COA, LQ.Level_A, RT.RankScore " 
sSQL = sSQL + " 		FROM "&RankTableName&" AS RT" 
sSQL = sSQL + " 			JOIN "&LeagueQfyTableName&" AS LQ" 
sSQL = sSQL + " 				ON RT.Event=LQ.Event AND RT.Div=LQ.Div AND LQ.LeagueID='"&CurrLeagueID&"'" 
sSQL = sSQL + " 		WHERE RT.SkiYearID=1) AS RT1"
sSQL = sSQL + " 	ON RT1.MemberID=RQ.MemberID AND RT1.Div=RQ.Div AND RT1.Event=RQ.Event"

' --- Joins the LeagueTable if LCQ By Rank is turned on for this LeagueID ---
sSQL = sSQL + " 	JOIN "&LeagueTableName&" AS LT" 
sSQL = sSQL + " 		ON LT.LeagueID='"&CurrLeagueID&"' AND LT.UseLCQRank=1" 

sSQL = sSQL + " 	JOIN"
sSQL = sSQL + " 		( SELECT MemberID, TourID, Event, Div, RankAfterCOD"
sSQL = sSQL + " 		FROM "&RegQualifyTableName&") AS RQ1"
sSQL = sSQL + " 	ON RQ1.MemberID=RQ.MemberID AND LEFT(RQ1.TourID,6)='"&LEFT(CurrTourID,6)&"' AND RQ1.Div=RQ.Div AND RQ1.Event=RQ.Event"

sSQL = sSQL + " WHERE RQ.TourID='"&LEFT(CurrTourID,6)&"' AND NOT(RQ.Div IN ('OM','OW','MM'))"

'response.write("<br><br>Line 850<br>"&sSQL)
'response.end

con.execute(sSQL)


END SUB








' ---------------------------------
  SUB QualifyLCQByOverall_AllTypes
' --------------------------------- 

' --- Updates QfyOverLCQByScr and QfyOverLCQByRank flags in the non-overall events ---
' --- NOTE:  This is the secondary operation required to update the events based on overall qualifications ---

sSQL =  " UPDATE RQ SET QfyOverLCQByScr=CASE WHEN RQ1.OverLCQByScr=1 THEN '1' ELSE '0' END," 
sSQL = sSQL + " 		QfyOverLCQByRank=CASE WHEN RQ1.OverLCQByRank=1 THEN '1' ELSE '0' END"
sSQL = sSQL + " 	FROM "&RegQualifyTableName&" AS RQ"
sSQL = sSQL + " 		JOIN"
sSQL = sSQL + " 		( SELECT MemberID, TourID, Event, Div, QfyByScrAfter AS OverLCQByScr, QfyByRankAfter AS OverLCQByRank"
sSQL = sSQL + " 			FROM "&RegQualifyTableName
sSQL = sSQL + " 			WHERE Event='O' ) AS RQ1"
sSQL = sSQL + " 		ON LEFT(RQ.TourID,6)=LEFT(RQ1.TourID,6) AND RQ1.MemberID=RQ.MemberID"
			
sSQL = sSQL + " WHERE LEFT(RQ.TourID,6)='"&LEFT(CurrTourID,6)&"' AND RQ.Event<>'O' AND RQ.Div=RQ1.Div"

'response.write("<br><br>Line883<br>"&sSQL)
'response.end

con.execute(sSQL)

END SUB




' ---------------------------------
  SUB QualifyByElite
' --------------------------------- 

' --- Writes the QfyByMElite and QfyByOElite flags indicating qualification by Elite with DateThru greater than the tournament start date

sSQL =  " UPDATE RQ SET QfyByOElite=CASE WHEN OQ.QualThru>='"&sTDateS&"' THEN '1' ELSE '0' END,"
sSQL = sSQL + " QfyByMElite=CASE WHEN MQ.QualThru>='"&sTDateS&"' THEN '1' ELSE '0' END" 
sSQL = sSQL + " 	FROM "&RegQualifyTableName&" AS RQ"

sSQL = sSQL + " 		LEFT JOIN"
sSQL = sSQL + " 		( SELECT MemberID, DivElite, DivOrig, Event, SkiYearID, QualThru"
sSQL = sSQL + " 			FROM "&EliteDateTableName&") AS OQ"
sSQL = sSQL + " 		ON OQ.MemberID=RQ.MemberID AND OQ.Event=RQ.Event AND OQ.DivElite IN ('OM','OW')"
sSQL = sSQL + " 		AND OQ.SkiYearID='1'"

sSQL = sSQL + " 		LEFT JOIN"
sSQL = sSQL + " 		( SELECT MemberID, DivElite, DivOrig, Event, SkiYearID, QualThru"
sSQL = sSQL + " 			FROM "&EliteDateTableName&") AS MQ"
sSQL = sSQL + " 		ON MQ.MemberID=RQ.MemberID AND MQ.Event=RQ.Event AND MQ.DivElite = 'MM'"
sSQL = sSQL + " 		AND MQ.SkiYearID='1'"

sSQL = sSQL + " WHERE LEFT(RQ.TourID,6)='"&LEFT(CurrTourID,6)&"'"

'response.write("<br><br>Line 917<br>"&sSQL)
'response.end

con.execute(sSQL)



END SUB







' ------------------------------------------------------------------------------------------------------
' ------- THE FOLLOWING SECTION BEGINS THE BY PLACEMENT CALCULATIONS 	--------------------------------
' ------------------------------------------------------------------------------------------------------



' ---------------------
  SUB UpdatePlacementA
' ---------------------


' --- Updates Regional Placement (A) ---
sSQL = " UPDATE RQ" 
sSQL = sSQL + " SET RQ.QfyByPlaceA=1 "
sSQL = sSQL + " FROM "&RegQualifyTableName&" AS RQ" 
sSQL = sSQL + " 	LEFT JOIN "&RawScoresTableName&" AS ST"
sSQL = sSQL + " 		ON RQ.MemberID=ST.MemberID AND RQ.Event=ST.Event AND LEFT(ST.TourID,6) IN "&TypeAList
sSQL = sSQL + " 	JOIN "&LeagueQfyTableName&" AS LQ"
sSQL = sSQL + " 		ON LQ.LeagueID='"&CurrLeagueID&"' AND LQ.Event=ST.Event AND LQ.Div=ST.Div" 
sSQL = sSQL + " WHERE LEFT(ST.Place,1)<=LQ.Place_TourA AND ST.Score>0 AND (NOT LEN(ST.Place)>1) AND (NOT ST.Place='')" 
sSQL = sSQL + " AND LEFT(RQ.TourID,6)='"&LEFT(CurrTourID,6)&"'" 
sSQL = sSQL + " AND (ST.Div=RQ.Div OR ST.Div=(SELECT DT.Div FROM "&DivisionsTableName&" AS DT WHERE DT.Next_Div=RQ.Div AND DT.SkiYearID=1))"

' NOTE: 7-9-2013 No updated for MW Regionals ---
'response.write("<br><br>Line 956<br>"&sSQL)
'response.end

con.execute(sSQL)


END SUB


' ---------------------
  SUB UpdatePlacementB
' ---------------------


' --- Updates Regional Placement (B) ---
sSQL = " UPDATE RQ" 
sSQL = sSQL + " SET RQ.QfyByPlaceB=1 "
sSQL = sSQL + " FROM "&RegQualifyTableName&" AS RQ" 
sSQL = sSQL + " 	LEFT JOIN "&RawScoresTableName&" AS ST"
sSQL = sSQL + " 		ON RQ.MemberID=ST.MemberID AND RQ.Event=ST.Event AND LEFT(ST.TourID,6) IN "&TypeBList
sSQL = sSQL + " 	JOIN "&LeagueQfyTableName&" AS LQ"
sSQL = sSQL + " 		ON LQ.LeagueID='"&CurrLeagueID&"' AND LQ.Event=ST.Event AND LQ.Div=ST.Div" 
sSQL = sSQL + " WHERE LEFT(ST.Place,1)<=LQ.Place_TourB AND ST.Score>0 AND (NOT LEN(ST.Place)>1) AND (NOT ST.Place='')" 
sSQL = sSQL + " AND LEFT(RQ.TourID,6)='"&LEFT(CurrTourID,6)&"'" 
sSQL = sSQL + " AND (ST.Div=RQ.Div OR ST.Div=(SELECT DT.Div FROM "&DivisionsTableName&" AS DT WHERE DT.Next_Div=RQ.Div AND DT.SkiYearID=1))"

'response.write("<br><br>Line 984<br>"&sSQL)
'response.end

con.execute(sSQL)

END SUB



' ---------------------
  SUB UpdatePlacementC
' ---------------------


sSQL = " UPDATE RQ SET RQ.QfyByPlaceC=1 FROM "&RawScoresTableName&" AS ST"
sSQL = sSQL + " JOIN "&RegQualifyTableName&" AS RQ ON LEFT(RQ.TourID,6)='"&LEFT(CurrTourID,6)&"'" 
sSQL = sSQL + 	" AND RQ.MemberID=ST.MemberID AND RQ.Event=ST.Event AND RQ.Div=ST.Div"
sSQL = sSQL + " JOIN "&LeagueQfyTableName&" AS LQ ON LQ.LeagueID='"&CurrLeagueID&"'"
sSQL = sSQL + 	" AND LQ.Event=ST.Event AND LQ.Div=ST.Div" 
sSQL = sSQL + " WHERE LEFT(ST.Place,1)<=LQ.Place_TourC AND (NOT LEN(ST.Place)>1) AND (NOT ST.Place='')" 
sSQL = sSQL + " AND LEFT(ST.TourID,6) IN "&TypeCList

'response.write("<br><br>Line 1006<br>"&sSQL)
'response.end

con.execute(sSQL)

END SUB


' ---------------------
  SUB UpdatePlacementD
' ---------------------

' --- Updates Other Tournament (D) Placement ---
sSQL = " UPDATE RQ SET RQ.QfyByPlaceD=1 FROM "&RawScoresTableName&" AS ST"
sSQL = sSQL + " JOIN "&RegQualifyTableName&" AS RQ ON LEFT(RQ.TourID,6)='"&LEFT(CurrTourID,6)&"'" 
sSQL = sSQL + 	" AND RQ.MemberID=ST.MemberID AND RQ.Event=ST.Event AND RQ.Div=ST.Div"
sSQL = sSQL + " JOIN "&LeagueQfyTableName&" AS LQ ON LQ.LeagueID='"&CurrLeagueID&"'"
sSQL = sSQL + 	" AND LQ.Event=ST.Event AND LQ.Div=ST.Div" 
sSQL = sSQL + " WHERE LEFT(ST.Place,1)<=LQ.Place_TourD AND (NOT LEN(ST.Place)>1) AND (NOT ST.Place='')" 
sSQL = sSQL + " AND LEFT(ST.TourID,6) IN "&TypeDList

'response.write("<br><br>Line 1025<br>"&sSQL)
'response.end

con.execute(sSQL)

END SUB



' ------------------------------
  SUB Update_3EventPart_InStates
' ------------------------------

sSQL = " UPDATE SEL SET QfyByState_3EvPart=Qfy_ByStatePart"
sSQL = sSQL + 	" FROM "&RegQualifyTableName&" AS SEL" 
sSQL = sSQL + " JOIN" 
sSQL = sSQL + " ( SELECT RQ.MemberID, CASE WHEN COUNT(EventCount)='3' AND LT.Qfy_By_AnyOverall_InStates=1 THEN 1 END AS Qfy_ByStatePart"  
sSQL = sSQL + 	" FROM "&RegQualifyTableName&" AS RQ "
sSQL = sSQL + " 	JOIN "&LeagueTableName&" AS LT" 
sSQL = sSQL + " 		ON LT.LeagueID='"&CurrLeagueID&"'"
sSQL = sSQL + " 	LEFT JOIN" 
sSQL = sSQL + " 	   (SELECT DISTINCT Event AS EventCount, MemberID FROM "&RawScoresTableName 
sSQL = sSQL + " 		WHERE LEFT(TourID,6) IN "&TypeCList&" ) AS ECT"
sSQL = sSQL + " 	ON ECT.MemberID=RQ.MemberID"
sSQL = sSQL + " WHERE LEFT(RQ.TourID,6) ='"&LEFT(CurrTourID,6)&"'"
sSQL = sSQL + " GROUP BY RQ.MemberID, RQ.TourID, RQ.Event, LT.Qfy_By_AnyOverall_InStates"
sSQL = sSQL + " ) AS OSEL"
sSQL = sSQL + " ON OSEL.MemberID=SEL.MemberID AND TourID='"&LEFT(CurrTourID,6)&"'"

'response.write("<br><br>Line 1053<br>"&sSQL)
'response.end

con.execute(sSQL)

END SUB








' ----------------------
  SUB QfyStatusUpdateNEW
' ----------------------

sSQL = " UPDATE RQ SET QfyStatus="
sSQL = sSQL + 	" 	CASE WHEN (RQ1.MemberID IS NOT NULL AND LT.COD<=GETDATE()) THEN 'Qualified'"
sSQL = sSQL + 	" 	WHEN RQ2.MemberID IS NOT NULL THEN 'Qualified'"
sSQL = sSQL + 	" 	WHEN RQ3.MemberID IS NOT NULL THEN 'Qualified'" 

' --- New 3-26-2012 to make anyone in a division in State tournament with Level_A=0 qualified even before COD ---
sSQL = sSQL + 	" 		  WHEN LQ.Level_A=0 AND LQ.LeagueID IS NOT NULL AND LT.HomoType='C' THEN 'Qualified' "

sSQL = sSQL + 	" 		  WHEN RQ1.MemberID IS NOT NULL AND LT.COD>GETDATE() THEN 'Pending' "
sSQL = sSQL + 	" 		   ELSE 'NCQ' END" 
	
sSQL = sSQL + 	" 	FROM "&RegQualifyTableName&" AS RQ"

sSQL = sSQL + 	" 	LEFT JOIN "
sSQL = sSQL + 	" 		( SELECT MemberID, TourID, Event, Div"
sSQL = sSQL + 	" 			FROM "&RegQualifyTableName
sSQL = sSQL + 	" 				WHERE ( QfyByRankByCOD=1 OR QfyBy3rdEvt=1 OR QfyByOverall=1)"  
sSQL = sSQL + 	" 					) AS RQ1"
sSQL = sSQL + 	" 		ON RQ1.MemberID=RQ.MemberID AND RQ1.TourID=RQ.TourID AND RQ1.Event=RQ.Event AND RQ1.Div=RQ.Div" 

sSQL = sSQL + 	" 	LEFT JOIN "
sSQL = sSQL + 	" 		( SELECT MemberID, TourID, Event, Div"
sSQL = sSQL + 	" 			FROM "&RegQualifyTableName
sSQL = sSQL + 	" 				WHERE ( QfyByRankAfter=1 OR QfyByScrAfter=1 OR QfyOverLCQByScr=1 OR QfyOverLCQByRank=1)"  
sSQL = sSQL + 	" 					) AS RQ2"
sSQL = sSQL + 	" 		ON RQ2.MemberID=RQ.MemberID AND RQ2.TourID=RQ.TourID AND RQ2.Event=RQ.Event AND RQ2.Div=RQ.Div" 

sSQL = sSQL + 	" 	LEFT JOIN "
sSQL = sSQL + 	" 		( SELECT MemberID, TourID, Event, Div"
sSQL = sSQL + 	" 			FROM "&RegQualifyTableName
sSQL = sSQL + 	"  				WHERE  QfyByPlaceA=1 OR QfyByPlaceB=1 OR QfyByPlaceC=1 OR QfyByPlaceD=1 OR QfyByOElite=1 OR QfyByMElite=1) AS RQ3"
sSQL = sSQL + 	" 		ON RQ3.MemberID=RQ.MemberID AND RQ3.TourID=RQ.TourID AND RQ3.Event=RQ.Event AND RQ3.Div=RQ.Div"

sSQL = sSQL + 	" 	LEFT JOIN "
sSQL = sSQL + 	" 		(SELECT LeagueID, COD, HomoType"
sSQL = sSQL + 	" 			FROM "&LeagueTableName&") AS LT"
sSQL = sSQL + 	" 				ON LT.LeagueID='"&CurrLeagueID&"'"

' --- New 3-26-2012 ---
sSQL = sSQL + 	" 	LEFT JOIN "
sSQL = sSQL + 	" 		(SELECT LeagueID, Level_A, Event, Div"
sSQL = sSQL + 	" 			FROM "&LeagueQfyTableName&") AS LQ"
sSQL = sSQL + 	" 		ON LQ.LeagueID='"&CurrLeagueID&"' AND LQ.Event=RQ.Event AND LQ.Div=RQ.Div"

sSQL = sSQL + 	" WHERE LEFT(RQ.TourID,6)='"&CurrTourID&"'"

'response.write("<br><br>Line 1120<br>"&sSQL)
'response.end

con.execute(sSQL)

END SUB







' --------------------------------
  SUB QfyStatusUpdateNEW_08072013
' --------------------------------


sSQL = " UPDATE RQ SET QfyStatus_Test="
sSQL = sSQL + 	" 	CASE"
' --- Qualified by ranking and after COD and participated in that event/div in regionals ---
sSQL = sSQL + 	" 	WHEN RQ1.MemberID IS NOT NULL AND LT.COD<=GETDATE() AND (RequirePart<>'B' OR SkiedRegls IN ('C','E','M','S','W','O')) THEN 'Qualified'"
' --- Qualified by RankAfter or ScoreAfter or OverallAfter and participated in that event/div in regionals ---
sSQL = sSQL + 	" 	WHEN RQ2.MemberID IS NOT NULL AND (RequirePart<>'B' OR SkiedRegls IN ('C','E','M','S','W','O')) THEN 'Qualified'"
' --- Qualified by Placement and participated in that event/div in regionals ---
sSQL = sSQL + 	" 	WHEN RQ3.MemberID IS NOT NULL AND (RequirePart<>'B' OR SkiedRegls IN ('C','E','M','S','W','O')) THEN 'Qualified'" 
' --- Qualified by OElite and no event/div participation required ---
sSQL = sSQL + 	" 	WHEN RQ4.MemberID IS NOT NULL AND RQ.Div IN ('OM','OW') THEN 'Qualified'" 
' --- Qualified by MElite and no event/div participation required ---
sSQL = sSQL + 	" 	WHEN RQ5.MemberID IS NOT NULL AND RQ.Div IN ('MM') THEN 'Qualified'" 
' --- Qualified in Overall Event and participated in that event/div in regionals ---
sSQL = sSQL + 	" 	WHEN RQ1.MemberID IS NOT NULL AND RQ.Event='O' AND SkiedRegls_S=1 AND SkiedRegls_T=1 AND SkiedRegls_J=1 THEN 'Qualified'" 
' --- Qualified by any means and has Admin Regional Override ---
sSQL = sSQL + 	" 	WHEN ( (RQ1.MemberID IS NOT NULL AND LT.COD<=GETDATE() ) OR RQ2.MemberID IS NOT NULL OR RQ3.MemberID IS NOT NULL ) AND RegionalOverride='Y' THEN 'Qfy w-RO'" 

' --- New 3-26-2012 to make anyone in a division in State tournament with Level_A=0 qualified even before COD ---
sSQL = sSQL + 	" 		  WHEN LQ.Level_A=0 AND LQ.LeagueID IS NOT NULL AND LT.HomoType='C' THEN 'Qualified' "

sSQL = sSQL + 	" 		  WHEN RQ1.MemberID IS NOT NULL AND LT.COD>GETDATE() THEN 'Pending' "
sSQL = sSQL + 	" 		   ELSE 'NCQ' END" 
	
sSQL = sSQL + 	" 	FROM "&RegQualifyTableName&" AS RQ"

sSQL = sSQL + 	" 	LEFT JOIN "
sSQL = sSQL + 	" 		( SELECT MemberID, TourID, Event, Div"
sSQL = sSQL + 	" 			FROM "&RegQualifyTableName
sSQL = sSQL + 	" 				WHERE ( QfyByRankByCOD=1 OR QfyBy3rdEvt=1 OR QfyByOverall=1)"  
sSQL = sSQL + 	" 					) AS RQ1"
sSQL = sSQL + 	" 		ON RQ1.MemberID=RQ.MemberID AND LEFT(RQ1.TourID,6)=LEFT(RQ.TourID,6) AND RQ1.Event=RQ.Event AND RQ1.Div=RQ.Div" 

sSQL = sSQL + 	" 	LEFT JOIN "
sSQL = sSQL + 	" 		( SELECT MemberID, TourID, Event, Div"
sSQL = sSQL + 	" 			FROM "&RegQualifyTableName
sSQL = sSQL + 	" 				WHERE ( QfyByRankAfter=1 OR QfyByScrAfter=1 OR QfyOverLCQByScr=1 OR QfyOverLCQByRank=1)"  
sSQL = sSQL + 	" 					) AS RQ2"
sSQL = sSQL + 	" 		ON RQ2.MemberID=RQ.MemberID AND LEFT(RQ2.TourID,6)=LEFT(RQ.TourID,6) AND RQ2.Event=RQ.Event AND RQ2.Div=RQ.Div" 

sSQL = sSQL + 	" 	LEFT JOIN "
sSQL = sSQL + 	" 		( SELECT MemberID, TourID, Event, Div"
sSQL = sSQL + 	" 			FROM "&RegQualifyTableName
sSQL = sSQL + 	"  				WHERE  QfyByPlaceA=1 OR QfyByPlaceB=1 OR QfyByPlaceC=1 OR QfyByPlaceD=1) AS RQ3"
sSQL = sSQL + 	" 		ON RQ3.MemberID=RQ.MemberID AND LEFT(RQ3.TourID,6)=LEFT(RQ.TourID,6) AND RQ3.Event=RQ.Event AND RQ3.Div=RQ.Div"

sSQL = sSQL + 	" 	LEFT JOIN "
sSQL = sSQL + 	" 		( SELECT MemberID, TourID, Event, Div"
sSQL = sSQL + 	" 			FROM "&RegQualifyTableName
sSQL = sSQL + 	"  				WHERE QfyByOElite=1) AS RQ4"
sSQL = sSQL + 	" 		ON RQ4.MemberID=RQ.MemberID AND LEFT(RQ4.TourID,6)=LEFT(RQ.TourID,6) AND RQ4.Event=RQ.Event AND RQ4.Div=RQ.Div"

sSQL = sSQL + 	" 	LEFT JOIN "
sSQL = sSQL + 	" 		( SELECT MemberID, TourID, Event, Div"
sSQL = sSQL + 	" 			FROM "&RegQualifyTableName
sSQL = sSQL + 	"  				WHERE QfyByMElite=1) AS RQ5"
sSQL = sSQL + 	" 		ON RQ5.MemberID=RQ.MemberID AND LEFT(RQ5.TourID,6)=LEFT(RQ.TourID,6) AND RQ5.Event=RQ.Event AND RQ5.Div=RQ.Div"

' --- New 8-8-2013 - Determines whether the Member skied all three events in this division in the Regionals ---
sSQL = sSQL + 	" 	LEFT JOIN"
sSQL = sSQL + 	" 		(SELECT MemberID, TourID, Div, Event, CASE WHEN SkiedRegls IN ('C','E','M','S','W','O') THEN 1 ELSE 0 END AS SkiedRegls_S"
sSQL = sSQL + 	" 			FROM "&RegQualifyTableName&" WHERE Event='S') AS OES"
sSQL = sSQL + 	" 	ON OES.MemberID=RQ.MemberID AND LEFT(OES.TourID,6)=LEFT(RQ.TourID,6) AND OES.Div=RQ.Div AND RQ.Event='O'" 		
sSQL = sSQL + 	" 	LEFT JOIN"
sSQL = sSQL + 	" 		(SELECT MemberID, TourID, Div, Event, CASE WHEN SkiedRegls IN ('C','E','M','S','W','O') THEN 1 ELSE 0 END AS SkiedRegls_T"
sSQL = sSQL + 	" 			FROM "&RegQualifyTableName&" WHERE Event='T') AS OET"
sSQL = sSQL + 	" 	ON OET.MemberID=RQ.MemberID AND LEFT(OET.TourID,6)=LEFT(RQ.TourID,6) AND OET.Div=RQ.Div AND RQ.Event='O'" 		
sSQL = sSQL + 	" 	LEFT JOIN"
sSQL = sSQL + 	" 		(SELECT MemberID, TourID, Div, Event, CASE WHEN SkiedRegls IN ('C','E','M','S','W','O') THEN 1 ELSE 0 END AS SkiedRegls_J"
sSQL = sSQL + 	" 			FROM "&RegQualifyTableName&" WHERE Event='J') AS OEJ"
sSQL = sSQL + 	" 	ON OEJ.MemberID=RQ.MemberID AND LEFT(OEJ.TourID,6)=LEFT(RQ.TourID,6) AND OEJ.Div=RQ.Div AND RQ.Event='O'"		

' --- Regional override ---
sSQL = sSQL + 	" 	LEFT JOIN"
sSQL = sSQL + 	" 	  ( SELECT MemberID, TourID, 'Y' AS RegionalOverride"
sSQL = sSQL + 	"         FROM "&RegGenTableName
sSQL = sSQL + 	" 		       WHERE RegionalOverride>'A' AND LEFT(TourID,6)='"&CurrTourID&"') AS RGN"
sSQL = sSQL + 	" 	ON RGN.MemberID=RQ.MemberID AND RGN.TourID=RQ.TourID"

' --- League information --- 
sSQL = sSQL + 	" 	LEFT JOIN "
sSQL = sSQL + 	" 		(SELECT LeagueID, COD, HomoType, RequirePart"
sSQL = sSQL + 	" 			FROM "&LeagueTableName&") AS LT"
sSQL = sSQL + 	" 				ON LT.LeagueID='"&CurrLeagueID&"'"

' --- New 3-26-2012 ---
sSQL = sSQL + 	" 	LEFT JOIN "
sSQL = sSQL + 	" 		(SELECT LeagueID, Level_A, Event, Div"
sSQL = sSQL + 	" 			FROM "&LeagueQfyTableName&") AS LQ"
sSQL = sSQL + 	" 		ON LQ.LeagueID='"&CurrLeagueID&"' AND LQ.Event=RQ.Event AND LQ.Div=RQ.Div"

sSQL = sSQL + 	" WHERE LEFT(RQ.TourID,6)='"&CurrTourID&"'"



response.write("<br><br>Line 1241<br>"&sSQL)
'response.end

con.execute(sSQL)

END SUB




' -----------------------------
   SUB ApplyAdminOverrides
' -----------------------------

sSQL = "UPDATE RGN SET QfyStatus_Test="
sSQL = sSQL + 	" 	CASE WHEN RegionalOverride='Y' AND IsQualified='Y' THEN 'Pend Regls' END"

sSQL = sSQL + 	" 	  ( SELECT MemberID, TourID, 'Y' AS RegionalOverride"
sSQL = sSQL + 	"         FROM "&RegGenTableName
sSQL = sSQL + 	" 		       WHERE RegionalOverride>'A' AND LEFT(TourID,6)='"&CurrTourID&"') AS RGN"

sSQL = sSQL + 	"			LEFT JOIN"
sSQL = sSQL + 	"		 		(	SELECT MemberID, TourID, 'Y' AS IsQualified"
sSQL = sSQL + 	" 					FROM "&RegQualifyTableName
sSQL = sSQL + 	"							WHERE QfyByRankByCOD=1 OR QfyBy3rdEvt=1 OR QfyByOverall=1 OR QfyByRankAfter=1 OR QfyByScrAfter=1"
sSQL = sSQL + 	"				 				OR QfyOverLCQByScr=1 OR QfyOverLCQByRank=1"
sSQL = sSQL + 	"				 				OR QfyByPlaceA=1 OR QfyByPlaceB=1 OR QfyByPlaceC=1 OR QfyByPlaceD=1) AS QFY1" 
sSQL = sSQL + 	" 		ON QFY1.MemberID=RGN.MemberID AND QFY1.TourID=RGN.TourID"		

response.write("<br><br>Line 1250<br>"&sSQL)
response.end

con.execute(sSQL)

END SUB





' -----------------------------
   SUB RequiredParticipation
' -----------------------------

sSQL = " UPDATE RQ SET SkiedRegls="
sSQL = sSQL + 	" (CASE WHEN RE.RegionalOverride<>'' THEN 'X'"
sSQL = sSQL + 	" 	WHEN LT.RequirePart='' OR LT.RequirePart IS NULL THEN '-'"
sSQL = sSQL + 	" 	WHEN LEFT(RQ.Div,1)='O' THEN 'O'"
sSQL = sSQL + 	"  	WHEN LE.TourID IS NOT NULL THEN RIGHT(LEFT(LE.TourID,3),1)"

' --- Gives 2 day grace period from the date of the LAST required tournament (regionals) ---
' --- Last [regionals] date used rather than assuming 3rd digit of sanction matches home region ---
sSQL = sSQL + 	"  	WHEN ST.MaxRegDate>'"&DATE+2&"' THEN 'P'"
sSQL = sSQL + 	" ELSE '' END)" 

' --- Determines which tournament type is required, if any ---
sSQL = sSQL + 	"	FROM"
sSQL = sSQL + 	"	(SELECT RequirePart, LeagueID"
sSQL = sSQL + 	"		FROM "&LeagueTableName 
sSQL = sSQL + 	"	WHERE LeagueID='"&CurrLeagueID&"') AS LT,"

' --- Finds the LAST date of any of the required tournaments ---
sSQL = sSQL + 	"	(SELECT MAX(TDateE) AS MaxRegDate"
sSQL = sSQL + 	"		 FROM "&SanctionTableName
sSQL = sSQL + 	"		WHERE LEFT(TournAppID,6) "
sSQL = sSQL + 	"	IN (SELECT LEFT(TourID,6) FROM "&LeagueToursTableName&" WHERE LeagueID='"&CurrLeagueID&"') ) AS ST,"

sSQL = sSQL + 	"	"&RegQualifyTableName&" AS RQ"

' --- Determines if member has any scores from any of the required tournaments ---
sSQL = sSQL + 	"	LEFT JOIN"	
sSQL = sSQL + 	"	( SELECT MemberID, TourID, Event, Div FROM "&RawScoresTableName
sSQL = sSQL + 	"	    WHERE" 
sSQL = sSQL + 	"		LEFT(TourID,6) IN ( SELECT LEFT(TourID,6)" 
sSQL = sSQL + 	"			FROM "&LeagueToursTableName 
sSQL = sSQL + 	"				WHERE LeagueID='"&CurrLeagueID&"'" 
sSQL = sSQL + 	"					AND TourType=(SELECT RequirePart FROM "&LeagueTableName&" WHERE LeagueID='"&CurrLeagueID&"')" 
sSQL = sSQL + 	"	)" 
sSQL = sSQL + 	"	) AS LE"
sSQL = sSQL + 	"	ON RQ.MemberID=LE.MemberID AND RQ.Event=LE.Event AND RQ.Div=LE.Div"

' --- Determines if member has an override value in the RegisterEvents table for this Event/Div/Tour ---
sSQL = sSQL + 	"	LEFT JOIN" 
sSQL = sSQL + 	"		(SELECT RegionalOverride, TourID, MemberID" 
sSQL = sSQL + 	"			FROM "&RegGenTableName&" ) AS RE"
sSQL = sSQL + 	"	ON LEFT(RE.TourID,6)=LEFT(RQ.TourID,6) AND RE.MemberID=RQ.MemberID" 


sSQL = sSQL + 	" WHERE LEFT(RQ.TourID,6)='"&CurrTourID&"' AND RQ.Event<>'O'"

'response.write("<br><br>Line 1179<br>"&sSQL)
'response.end

con.execute(sSQL)

END SUB



' ------------------------
  SUB WriteToHistoryTable
' ------------------------

sSQL = "INSERT INTO "&RegQfyHistoryTable 
sSQL = sSQL + 	" (TourID, CalcDate, TotalRegQfy, QfyByRankByCOD, QfyBy3rdEvt, QfyByOverall, QfyByRankAfter, QfyByScrAfter, QfyByPlaceA, QfyByPlaceB, QfyByPlaceC, QfyByPlaceD)"
sSQL = sSQL + 	" SELECT '"&LEFT(CurrTourID,6)&"', GETDATE(),"
sSQL = sSQL + 	"	( SELECT Count(MemberID) FROM "&RegQualifyTableName&" WHERE LEFT(TourID,6)='"&LEFT(CurrTourID,6)&"'),"
sSQL = sSQL + 	"	( SELECT Count(MemberID) FROM "&RegQualifyTableName&"  WHERE LEFT(TourID,6)='"&LEFT(CurrTourID,6)&"' AND QfyByRankByCOD=1),"
sSQL = sSQL + 	"	( SELECT Count(MemberID) FROM "&RegQualifyTableName&"  WHERE LEFT(TourID,6)='"&LEFT(CurrTourID,6)&"' AND QfyBy3rdEvt=1),"
sSQL = sSQL + 	"	( SELECT Count(MemberID) FROM "&RegQualifyTableName&"  WHERE LEFT(TourID,6)='"&LEFT(CurrTourID,6)&"' AND QfyByOverall=1),"
sSQL = sSQL + 	"	( SELECT Count(MemberID) FROM "&RegQualifyTableName&"  WHERE LEFT(TourID,6)='"&LEFT(CurrTourID,6)&"' AND QfyByRankAfter=1),"
sSQL = sSQL + 	"	( SELECT Count(MemberID) FROM "&RegQualifyTableName&"  WHERE LEFT(TourID,6)='"&LEFT(CurrTourID,6)&"' AND QfyByScrAfter=1),"
sSQL = sSQL + 	"	( SELECT Count(MemberID) FROM "&RegQualifyTableName&"  WHERE LEFT(TourID,6)='"&LEFT(CurrTourID,6)&"' AND QfyByPlaceA=1),"
sSQL = sSQL + 	"	( SELECT Count(MemberID) FROM "&RegQualifyTableName&"  WHERE LEFT(TourID,6)='"&LEFT(CurrTourID,6)&"' AND QfyByPlaceB=1),"
sSQL = sSQL + 	"	( SELECT Count(MemberID) FROM "&RegQualifyTableName&"  WHERE LEFT(TourID,6)='"&LEFT(CurrTourID,6)&"' AND QfyByPlaceC=1),"
sSQL = sSQL + 	"	( SELECT Count(MemberID) FROM "&RegQualifyTableName&"  WHERE LEFT(TourID,6)='"&LEFT(CurrTourID,6)&"' AND QfyByPlaceD=1)"

con.execute(sSQL)

END SUB




' --------------------
  SUB CountUpdates
' --------------------

response.write("<br><br>")
response.write("<br><b>LeagueID: &nbsp;"&CurrLeagueID&"</b>")
response.write("<br>")
response.write("<br>TourID Associated with this League: "&CurrTourID) 

response.write("<br>TypeAList="&TypeAList)
response.write("<br>TypeBList="&TypeBList)
response.write("<br>TypeCList="&TypeCList)
response.write("<br>TypeDList="&TypeDList)


' --- Count for COA update in the LeagueQualify table ---
SET rs=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT Count(LeagueID) AS TotCount FROM "&LeagueQfyTableName&" WHERE LeagueID='"&CurrLeagueID&"'"
rs.open sSQL, SConnectionToTRATable

IF NOT rs.eof THEN 
	response.write("<br><br>Total records in LeagueQualify table ="&rs("TotCount"))
ELSE
	response.write("<br><br>No records updated in LeagueQualify")
END IF  


' --- From Create Qualify Records ---
SET rs=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT Count(MemberID) AS TotCount FROM "&RegQualifyTableName&" WHERE TourID='"&CurrTourID&"'"
rs.open sSQL, SConnectionToTRATable

IF NOT rs.eof THEN 
	response.write("<br>Total records in RegisterQualify table for this LeagueID ="&rs("TotCount"))
ELSE
	response.write("<br>No records in RegisterQualify")
END IF  


' --- By Level prior to COD ---
SET rs=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT Count(MemberID) AS TotCount FROM "&RegQualifyTableName&" WHERE LEFT(TourID,6)='"&CurrTourID&"' AND QfyByRankByCOD=1"
rs.open sSQL, SConnectionToTRATable

IF NOT rs.eof THEN 
	response.write("<br>Total qualified by current Ranking exceeding COA  ="&rs("TotCount"))
ELSE
	response.write("<br>No records by ranking exceeding COA")
END IF  


' --- By 3rd Event ---
SET rs=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT Count(MemberID) AS TotCount FROM "&RegQualifyTableName&" WHERE LEFT(TourID,6)='"&CurrTourID&"' AND QfyBy3rdEvt=1"
rs.open sSQL, SConnectionToTRATable

IF NOT rs.eof THEN 
	response.write("<br>Total qualified by 3rd Event Score higher than 3rd event COA  ="&rs("TotCount"))
ELSE
	response.write("<br>No records by 3rd Event Qualification")
END IF  


' --- By Overall ---
SET rs=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT Count(MemberID) AS TotCount FROM "&RegQualifyTableName&" WHERE LEFT(TourID,6)='"&CurrTourID&"' AND QfyByOverall=1"
rs.open sSQL, SConnectionToTRATable

IF NOT rs.eof THEN 
	response.write("<br>Total qualified by Overall Ranking Level  ="&rs("TotCount"))
ELSE
	response.write("<br>No records by Overall Ranking Level")
END IF  



' --- Counts updated from QualifyByScoreRankCOD ---
SET rs=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT Count(MemberID) AS TotCount FROM "&RegQualifyTableName&" WHERE LEFT(TourID,6)='"&CurrTourID&"' AND QfyByRankAfter=1"
rs.open sSQL, SConnectionToTRATable

IF NOT rs.eof THEN 
	response.write("<br>Total by LCQ Rank above COA after COD and before Tournament ="&rs("TotCount"))
ELSE
	response.write("<br>No Records from LCQ Rank")
END IF  


' --- Counts updated from Qualify_LCQByScore ---
SET rs=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT Count(MemberID) AS TotCount FROM "&RegQualifyTableName&" WHERE LEFT(TourID,6)='"&CurrTourID&"' AND QfyByScrAfter=1"
rs.open sSQL, SConnectionToTRATable

IF NOT rs.eof THEN 
	response.write("<br>Total by LCQ Score in Current Regionals or Defined Class between COD and Tournament ="&rs("TotCount"))
ELSE
	response.write("<br>No Records from LCQ Score")
END IF  



response.write("<br>")


' --- Counts what was just updated ---
SET rs=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT Count(MemberID) AS TotCount FROM "&RegQualifyTableName&" WHERE LEFT(TourID,6)='"&CurrTourID&"' AND QfyByPlaceA=1"
rs.open sSQL, SConnectionToTRATable

IF NOT rs.eof THEN 
	response.write("<br>Total Placement A - Previous Nationals ="&rs("TotCount"))
ELSE
	response.write("<br>No Records from Placement A")
END IF  


' --- Counts what was just updated ---
SET rs=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT Count(MemberID) AS PlaceBTotCount FROM "&RegQualifyTableName&" WHERE LEFT(TourID,6)='"&CurrTourID&"' AND QfyByPlaceB=1"
rs.open sSQL, SConnectionToTRATable

IF NOT rs.eof THEN 
	response.write("<br>Total Placement B - Previous Regionals ="&rs("PlaceBTotCount"))
ELSE
	response.write("<br>No Records from Placement B ")
END IF  


' --- Counts what was just updated ---
SET rs=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT Count(MemberID) AS TotCount FROM "&RegQualifyTableName&" WHERE LEFT(TourID,6)='"&CurrTourID&"' AND QfyByPlaceC=1"
rs.open sSQL, SConnectionToTRATable

IF NOT rs.eof THEN 
	response.write("<br>Total Placement C (used for States or Other) ="&rs("TotCount"))
ELSE
	response.write("<br>No Records from Placement C")
END IF  


' --- Counts what was just updated ---
SET rs=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT Count(MemberID) AS TotCount FROM "&RegQualifyTableName&" WHERE LEFT(TourID,6)='"&CurrTourID&"' AND QfyByPlaceD=1"
rs.open sSQL, SConnectionToTRATable

IF NOT rs.eof THEN 
	response.write("<br>Total Placement D (Other Qualifier) ="&rs("TotCount"))
ELSE
	response.write("<br>No Records from Placement D")
END IF  


END SUB



' ------------------------
  SUB DefineTypeTourList
' ------------------------

' --- Creates and comma delimiated list of tours with TourType=A ---

SET rsType=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT TourID, TourType FROM "&LeagueToursTableName
sSQL = sSQL + " WHERE LeagueID='"&CurrLeagueID&"'"
sSQL = sSQL + " ORDER BY TourID"
rsType.open sSQL, SConnectionToTRATable

TypeAList="("
TypeBList="("
TypeCList="("
TypeDList="("

IF NOT rsType.eof THEN rsType.MoveFirst

DO WHILE NOT rsType.eof
	IF TRIM(rsType("TourType"))="A" THEN TypeAList=TypeAList&"'"&LEFT(rsType("TourID"),6)&"',"
	IF TRIM(rsType("TourType"))="B" THEN TypeBList=TypeBList&"'"&LEFT(rsType("TourID"),6)&"',"
	IF TRIM(rsType("TourType"))="C" THEN TypeCList=TypeCList&"'"&LEFT(rsType("TourID"),6)&"',"
	IF TRIM(rsType("TourType"))="D" THEN TypeDList=TypeDList&"'"&LEFT(rsType("TourID"),6)&"',"

	rsType.movenext
LOOP


IF TRIM(TypeAList)<>"(" THEN TypeAList=LEFT(TypeAList,LEN(TypeAList)-1)&")" ELSE TypeAList=TypeAList&")"
IF TRIM(TypeBList)<>"(" THEN TypeBList=LEFT(TypeBList,LEN(TypeBList)-1)&")" ELSE TypeBList=TypeBList&")"
IF TRIM(TypeCList)<>"(" THEN TypeCList=LEFT(TypeCList,LEN(TypeCList)-1)&")" ELSE TypeCList=TypeCList&")"
IF TRIM(TypeDList)<>"(" THEN TypeDList=LEFT(TypeDList,LEN(TypeDList)-1)&")" ELSE TypeDList=TypeDList&")"


END SUB




%>
