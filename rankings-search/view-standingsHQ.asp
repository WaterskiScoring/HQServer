<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_leagues.asp"-->
<!--#include virtual="/rankings/tools_include16.asp"-->
<!--#include virtual="/rankings/tools_registration.asp"-->
<title>Rankings List</title>
<%

Dim currentPage, rowCount, i
Dim MemoryScore, MemoryPlc, MemoryRank, RecordNum, RankValueWithTies
Dim tRankScore, tRankPct, tFmtScore, tRnkScoBkup, tMemberID
DIM tOpenDiv, tMastDiv, tEliteStat, tEliteBkup, tRating
Dim tName, tState, tRegion, tNatPlace, tRegSki, tRegPlace, tMemberFed
Dim tTeam, tTeamStat, tNCWRegn, tNCWConf, tDefaultYear, LastRegion
Dim DefineRowcolor, DefineLevelcolor, LastLevelcolor, tPerc3, tPerc4, tPerc5, tPerc6, tPerc7, tPerc8, tPerc9
Dim tLevelNo, LastLevelNo
Dim tBirth, tAge, tBirthday, tRCU
Dim sMemberID, sFullName, sEventName
Dim TourDisplaywidth, ScorePageBorderDark, ScorePageBorderLight
Dim MainImage
Dim COALevel9, COALevel8, COALevel7, COALevel6, COALevel5, COALevel4, COALevel3
Dim sDefaultNationals
Dim sShowSQL


Dim ThisFileName, SearchFileName
ThisFileName="View-StandingsHQ.asp"
SearchFileName = "search-memberHQ.asp"


tLevelNo = 999999

'response.write("---")
TourDisplaywidth=725
ScorePageBorderDark = HQSitecolor1
ScorePageBorderLight = HQSitecolor2


IF TRIM(Session("NewRankVis"))="" THEN
	KickTrafficCounter("NewRankVis")	
	Session("NewRankVis")="YES"
END IF


' --- This is a TEMPORARY fix.  Actual cut-off logic needs to be dynamic.
CutOffDate = "07/23/2008"


sShowSQL=Request("sShowSQL")
sRunByWhat = TRIM(Request("pvar"))
pvar = sRunByWhat

IF LCASE(pvar)="grassroots" THEN
		'response.end
		response.redirect("/rankings/view-grranking.asp?whatheadfoot=rs&pvar=grassroots")
END IF


' --- Temporary ---
adminmenulevel = TRIM(Request("adminmenulevel"))
'IF Session("adminmenulevel")>=50 THEN
'		response.write("Line 60 Shows to Admin only - pvar="&pvar)
'END IF




FilterSelected = TRIM(Request("FilterSelected"))
IF FilterSelected = "" THEN FilterSelected = "All"
IF Len(FilterSelected) > 3 then FilterSelected = Left(FilterSelected,3)


sLeagueSelected = TRIM(Request("sLeagueSelected"))
IF sLeagueSelected = "" THEN sLeagueSelected = "None"


' --- Note - The following request line must be above the definition within the CURRENT member section below. ---
FederationSelected = TRIM(Request("Include_International"))


' ----------------------------------------------
' --- Define Member information if available ---
' ----------------------------------------------
sMemberID=TRIM(Request("sMemberID"))
IF Len(sMemberID) > 9 then sMemberID = Left(sMemberID,9)


' --------------------------------------
' --- Determine Name of CURRENT user ---
' --------------------------------------
SET rs=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT * FROM "&MemberLiveTableName&" AS MT"
sSQL = sSQL + " WHERE PersonID = '" &RIGHT(sMemberID,8)& "'"
ChoosePagesSQL sSQL,currentPage, 30

IF NOT rs.EOF THEN
		sFullName=rs("FirstName")&" "&rs("LastName")
		' --- Only reset FED DropDown if the person has just selected a NEW member as determined by MyDiv or MyEvent NOT Null. ---
		IF UCASE(TRIM(rs("FederationCode")))<>"USA" AND (TRIM(Request("MyDiv"))<>"" OR TRIM(Request("MyEvent"))<>"") THEN
				FederationSelected=rs("FederationCode")
		END IF
ELSE
		sFullName="None Selected"
END IF


RecordNum = TRIM(Request("RecordNum"))    
IF RecordNum = "" THEN RecordNum = 1



' ----------------------
' --- Define SkiYear ---
' ----------------------
SkiYearSelected = TRIM(Request("SkiYear"))
IF TRIM(SkiYearSelected) = "" AND TRIM(Session("SkiYear"))<>"" THEN SkiYearSelected=Session("SkiYear")


' --------------------
' --- Define Event ---
' --------------------
EventSelected = TRIM(Request("event"))
IF TRIM(Request("MyEvent"))<>"" THEN EventSelected=TRIM(Request("MyEvent"))
IF EventSelected = "" THEN EventSelected = "S"
IF Len(EventSelected) > 1 then EventSelected = Left(EventSelected,1)

	
' --- Define Division ---
DivSelected = TRIM(Request("DivSelected"))
IF TRIM(Request("MyDiv"))<>"" THEN DivSelected=TRIM(Request("MyDiv"))
IF Len(DivSelected) > 2 then DivSelected = Left(DivSelected,2)


' --- Defines the image to be displayed in the drop downs box background ---

WhatDropdownImage EventSelected





' ---------------------------------------------
' --- Writes header portion of HQ main page ---
' ---------------------------------------------


  WriteIndexPageHeader


'response.write("<br> TOP OF PAGE")
'response.end
' --------------------------------------------------------------------------------- 
' Creates Radio Buttons to select LIST TYPE in case NOT selected from Settings menu
' ---------------------------------------------------------------------------------


IF sRunByWhat = "" THEN
    %>
    <br><br>
    <center><h2>View Rankings<br></h2>
    <br><br>
    <form action="/rankings/<%=ThisFileName%>?sMemberID=<%s=MemberID%>&NewType=Yes&rid=<%=rid%>" method="post">
    <input type="radio" name="pvar" value="National">National&nbsp;<br><br>
    <input type="radio" name="pvar" value="Junior">Junior&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br><br>
    <input type="radio" name="pvar" value="NCWSA">NCWSA<br><br>
    <input type="radio" name="pvar" value="NSL">Grassroots&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br><br>
    <input type="submit" value="Continue"><br><br><br>
    </form>
    <%

ELSE


    ' -----------------------------------------------------------------------------------------------------------
    ' ----------------   Sets Session("SkiYear") to request string from form   ------------------
    ' -----------------------------------------------------------------------------------------------------------
    ' --- NCWSA test is done first 
		IF (SkiYearSelected = "1" OR SkiYearSelected = "") AND sRunByWhat="NCWSA" THEN 

    		OpenCon
				Set rs = Server.CreateObject("ADODB.recordset")
				sSQL = "SELECT top 1 * from " & SkiYearTableName & " WHERE DefaultYear = 1"
    		rs.open sSQL, SConnectionToTRATable, 3, 3  

				IF NOT rs.EOF THEN
						Session("SkiYear")=rs("SkiYearID")
				END IF		


   	' --- Assigns SkiYear to whatever current setting is if there is a variable set on form
    ELSEIF SkiYearSelected <> "" THEN 
				Session("SkiYear") = SkiYearSelected

    
    ' --- If nothing is assigned, then set it to 12 month rankings
    ELSE 	
		
				Session("SkiYear")="1"	
    
    END IF	




	' --- Checks to make sure the DivSelected is in the divisions found in Rankings Table for this sRunByWhat ---
	'	commented out for now since we're going to key on no Division indicator to request first.
	'	CheckValidDivision



    ' -------- If NCWSA then select ALL Federations, otherwise only USA --------

    IF FederationSelected = "" AND sRunByWhat="NCWSA" THEN 
				FederationSelected = "ALL"	                
    ELSEIF FederationSelected = "" THEN 
				FederationSelected = "USA"
    END IF	


    currentPage = TRIM(Request("currentPage"))
    IF currentPage = "" THEN currentPage = 1
    
    sID = TRIM(Request("id"))
    IF sID = "" THEN sID = 0
            
            
	ThisPage = Request.ServerVariables("SCRIPT_NAME")
            

	' ------------------------------------------------------------------------------------------------           
	' -------------------------------  BEGINS WRITING HEADERS AND RANKINGS  --------------------------
	' ------------------------------------------------------------------------------------------------

	tempSkiYear = Session("SkiYear")


	' --- If User pressed Find My Rankings button and MemberID was not set OR user pressed get a New Member button ---
	IF (TRIM(Request("SingleRanking"))<>"" AND TRIM(sMemberID)="") OR TRIM(Request("NewMember"))<>"" THEN
			' --- This is where I would branch to get a member if not set ---

		
			' --- Sends user to search-member routine to selected member
			Session("SkiYear")="1"
			Session("sSendingPage")="/rankings/"&ThisFileName&"?SingleRanking=Find&pvar="&sRunByWhat
			Response.Redirect("/rankings/"&SearchFileName&"?rid="&rid&"&formstatus=search")

	ELSEIF Trim(Request("SingleRanking"))<>"" THEN

			SingleRanking="Find My Rankings"
			FindRankingInstances

	ELSE 		' --- Displays page ---


			' --- Displays picture box with drop downs ---
			DisplayDropDowns  

			' -- Displays divs containing levels between drop downs and rankings list --
			IF 2=1 AND Session("AdminMenuLevel")>=50 THEN
					FindCOAScoreAll
					
					DisplayCOAJava
			END IF

			' -------------------------------------------------------------------------------
			' -----  Check Recalculation Underway Flag for the Ski Year selected.  ----------
			' -----  If it's currently on, issue Come Back Later -- otherwise proceed.  -----
			' -------------------------------------------------------------------------------
		
			OpenCon
			sSQL = "SELECT Case when RecalcUnderway=1 THEN 'Y' ELSE 'N' END as RCUFlag FROM " & SkiYearTableName & " WHERE SkiYearID = " & tempSkiYear
			Set rs = Server.CreateObject("ADODB.recordset")
			rs.open sSQL, SConnectionToTRATable, 3, 3  
			
			IF rs.EOF THEN tRCU = "N" ELSE tRCU = RS("RCUFlag")
					rs.close

					IF tRCU = "Y" AND Session("AdminMenuLevel") = 0 and sRunByWhat <> "NCWSA" THEN   ' --- Calc underway - Tell them to try again later
							%><b><font color="red" size="2">
			  			<br>&nbsp;&nbsp;&nbsp;
			  			Ranking Recalculations are currently underway For the Ski Year requested.&nbsp; Please try
			  			<br>&nbsp;&nbsp;&nbsp;
						  your request again in a few minutes.&nbsp; We apologize for the temporary inconvenience.</font></b><% 	
   				ELSE

							' -------------------------------------------------------------------------------
							' -----  Check for presence of a Selected Division Code.  If there has been 
							' -----  none specified yet, then ask user to Select one -- otherwise proceed.  -----
							' -------------------------------------------------------------------------------
	
							IF DivSelected = "" THEN   
									' --- New Ranking Type -- Ask to select then hit Display
									
									%><b><font color="red" size="2">
				  				<br>&nbsp;&nbsp;&nbsp;
			  					Please Select a Division and Event (and any desired filters) using the drop-down
			  					<br>&nbsp;&nbsp;&nbsp;
			  					boxes above, then click the Display Rankings button to display that selection.</font></b><% 	
							ELSE

									' ---------------------------
									' --- Runs Rankings Query ---
									' ---------------------------
									RunStandingsQueryNew

			 						%>
									<TABLE width=98% height="500px" align=center>
									<TR>
			  						<TD style="white-space:nowrap">
			  							<%

											' ----------------------------------------------
											' --- Displays table header and ranking list ---
											' ----------------------------------------------
											DisplayRankList 

											KickTrafficCounter("NewRankPgs")  
											
											%>
		  	  						</TD>
									</TR>
									</TABLE>
									<%


									
									' --- Gets the Level Percentiles and COA Scores for the selected Division/Event ---
									FindCOAScoreAll

									' ------------------------------------------------------------------
									' --- Displays grid at bottom with Key to Percentiles and ranges ---
									' ------------------------------------------------------------------
									DisplayPercentilesandPageFooter

							END IF

					END IF

			END IF

  END IF




' ---------------------------------------------
' --- Writes header portion of HQ main page ---
' ---------------------------------------------

  WriteIndexPageFooter


' ---------------------------------------------------------------------------------------------------------------
' ----------------------   END OF MAIN CODE HERE  ---------------------------------------------------------------
' ---------------------------------------------------------------------------------------------------------------



%>
<script language="javascript" type="text/JavaScript">
	function OpenCloseCOAWindow(OpenCloseStatus) {
			if (OpenCloseStatus == "open") {
					document.getElementById('COAPanelClosed').style.display = 'none'; 		
					document.getElementById('COAPanelOpen').style.display = 'inline-block'; 
				}
			else if (OpenCloseStatus == "close") {
					document.getElementById('COAPanelOpen').style.display = 'none'; 		
					document.getElementById('COAPanelClosed').style.display = 'inline-block'; 
			}
	}
</script>
<%


' -------------------
  SUB DisplayCOAJava
' -------------------
	' -- Displays the div with COA
	
PlusSignImageURL="images/icons/Plus-icon.png"
' background-color:#ffe866; 	


%>
<div id="COAPanelClosed" style="display:inline-block; width:100%;">
	<a href="javascript:OpenCloseCOAWindow('open');" style="text-decoration:none">
		<TABLE class="innertable" align="center" style="height:25px; width:98%;">
			<tr>
				<td colspan="3" align="center" style="background-color:#ffe866">
					<font size=2 color="#000000">Cut-Off Average: 95.75</font> 
				</td>
				<td colspan="4" align="center" style="background-color:#ffe866">
					<font size=2 color="#000000" >COD: 8/15/2017</font> 
				</td>
				<td colspan="1" align="center" style="background: transparent url(<%= PlusSignImageURL %>) no-repeat center top;">&nbsp;</td>
			</tr>
		</TABLE>
	</a>
</div>	
<div id="COAPanelOpen" style="display:none; width:100%;">
	<a href="javascript:OpenCloseCOAWindow('close');" style="text-decoration:none">
	<TABLE class="innertable" align="center" style="height:25px; width:98%; background-color:#ffe866;">
		<tr>
			<td colspan="8" align="center" style="background-color:#ffe866;">
				<font size=2 color="#000000" >Close Cut-Off-Averages Window</font> 
			</td>
		</tr>
	</TABLE>
	</a>
	<%
	DisplayPercentilesandPageFooter
	%>
</div>
<%
END SUB





' -----------------------
   SUB DisplayRankList
' -----------------------

IF rs.eof THEN
		%>
		<TABLE class="innertable" width=98% align=center>
	  	<TR>
	  		<TD>
	  			<br><br>
					<font color="red">No Rankings Found With These Filter Settings.</font>
        </TD>
      </TR>
    </TABLE>
    <%
ELSE 


		' --- INITIALIZES the Ranking related memory fields for deal with ties.

		' --- RecordNum is essentially the record count
		' --- MemoryScore is the Score of the 
		' --- MemoryRank stores the highest value of placement - for which subsequent records may be tied 
		' --- tRankScore is the Score of the current record

		RecordNum = 1
		MemoryRank = 1
		IF sRunByWhat = "NSL" THEN MemoryScore = rs("sc_3") ELSE MemoryScore = rs("RankScore")
   
		' ---  After storing the values from the FIRST record then move to the second record to see if tied to know
		' ---     whether the FIRST record should have a T after it.  All others


		' --------------------------
		' --- Move to 2nd record ---
		' --------------------------
		rs.MoveNEXT


		IF NOT rs.EOF THEN
				' --- Initializes 2nd record in query --- 
				DefineRankingDataLine
		END IF


		' --- If the score from last tied record is same as current score 
		IF MemoryScore = tRankScore THEN
				RankValueWithTies = "1T"
		ELSE
				RankValueWithTies = "1"
		END IF


		' --------------------------------------------------------------------------
		' --- Now move back to FIRST record and initialize First record in query ---
		' --------------------------------------------------------------------------
		rs.MoveFIRST
		DefineRankingDataLine
  

		' -----------------------------------------------------------------------------------------------
		' -----  BEGINNING OF LOOPING FOR DISPLAYING ALL RECORDS MATCHING QUERY  ------------------------
		' -----   Loops thru the remaining 2,3...nth records
		' -----------------------------------------------------------------------------------------------

		%>
		<TABLE class="innertable" width=100% align=center style="border-width:2px;"><%

		' --- Displays the header on the top of the table ---
		DisplayRankHeader


		DO WHILE NOT rs.eof

				' --- Displays one line of ranking list ---
				DisplayRankingLine


				' --- Initializes NEXT record in query --- 
				rs.moveNEXT
				RecordNum = RecordNum + 1

				IF NOT rs.eof THEN
						' --- Defines the CURRENT record ---
						DefineRankingDataLine
			

						' --- If the score from PREVIOUS record is same as current score ---
						' ------------------------------------------------------------------
						IF cdbl(MemoryScore) = cdbl(tRankScore) THEN
								RankValueWithTies = MemoryRank&"T"
						ELSE
								MemoryRank = RecordNum
								MemoryScore = tRankScore
				
								' --- Move to NEXT record to see if tied---
								rs.MoveNEXT
			    			IF NOT rs.eof THEN
					
										' --- Initializes the record beyond the current record in query --- 
										DefineRankingDataLine

										' --- If the score from last tied record is same as current score 
										IF MemoryScore = tRankScore THEN
												RankValueWithTies = RecordNum&"T"
										ELSE
												RankValueWithTies = RecordNum
										END IF

								ELSE
										
										' --- Can't be tied with EOF so set it to the current record ---
										RankValueWithTies = RecordNum
								END IF


								' --- Now move back to CURRENT record and initialize ---
								IF NOT(rs.bof) THEN
										rs.MovePREVIOUS
										DefineRankingDataLine
								END IF
						END IF

				ELSE

				END IF

		LOOP  




	IF sRunByWhat = "NCWSA" THEN %>
		<form action="/rankings/view-TeamStdgsHQ.asp?NSL=&Event=<%=EventSelected%>&DivSelected=<%=DivSelected%>&SkiYear=<%=Session("SkiYear")%>" method="post">
		<td colspan=7 align=center>&nbsp;<br><input type="submit" style="width:12em" value="Team Rankings"
		title="Take me to the NCWSA TEAM&#13;Rankings for this Division / Event"><br>&nbsp;
		</td></form><%
	END IF

	%>
    </TABLE><%


END IF

CloseCon

END SUB




' ----------------------
  SUB DisplayRankHeader
' ----------------------


Headcolor1="#FFFFFF"

' ---------------   Top of large condition of branching to most of rest of code   ------------------

'	First conditional posts bold disclaimer if current incomplete ski year selected.

	IF sRunByWhat <> "NSL" AND sRunByWhat <> "NCWSA" AND tDefaultYear <> 0 THEN 
		
		%><TR><Td colspan="11" align="Center"><font size=<% =fontsize4 %> color="#FF0000"><b>
		Ski Year Rankings shown below are NOT for Qualifying.<br>Official Qualification Rankings Period is Last 12 months.
		</b></font></Td></TR><%
	
	END IF

      %><TR>
    	<Th align="Center" width=9% ><font size=<% =fontsize2 %> color="#FFFFFF"><b> <br>Rank</b></FONT></th>
    	<Th align="Left" ><font size=<% =fontsize2 %>  color="#FFFFFF"><b><br>&nbsp;&nbsp; Member Name</b></FONT></th><% 

	IF sRunByWhat = "NSL" THEN 
    		%><Th align="Center" ><font size=<% =fontsize2 %>  color="#FFFFFF"><b><br>Grassroots Placement Points</b></FONT></th><% 
	ELSE 
    		%><Th align="Center" ><font size=<% =fontsize2 %>  color="#FFFFFF"><b><br>Ranking<br>Score</b></FONT></TD>
    		<Th align="Center" ><font size=<% =fontsize2 %> color="#FFFFFF"><b><br>Elite<br>Status</b></FONT></TD><% 
	END IF 

	%><Th align="Center" ><font size=<% =fontsize2 %>  color="#FFFFFF"><b><br>Home<br>State</b></FONT></TD>
	
	<% IF sRunByWhat = "NCWSA" THEN %>
		<Th align="Center" ><font size=<% =fontsize2 %>  color="#FFFFFF"><b><br>Team</b></FONT></TD>
		<Th align="Center" ><font size=<% =fontsize2 %>  color="#FFFFFF"><b><br>Regn</b></FONT></TD>
		<Th align="Center" ><font size=<% =fontsize2 %>  color="#FFFFFF"><b><br>Conf</b></FONT></TD>
	<% ELSE %>
		<Th align="Center" ><font size=<% =fontsize2 %>  color="#FFFFFF"><b>Home<br>Region</b></FONT></TD>
	<% END IF

	IF sRunByWhat <> "NSL" AND sRunByWhat <> "NCWSA" AND EventSelected <> "O" THEN 
    		%><Th align="Center" ><font size=<% =fontsize2 %>  color="#FFFFFF"><b>Regl<br>Place</b></FONT></TD>
		<Th align="Center" ><font size=<% =fontsize2 %>  color="#FFFFFF"><b>Natl<br>Place</b></FONT></TD><% 
	END IF 

	IF sRunByWhat<>"NCWSA" AND sRunByWhat<>"NSL" THEN %>
		<Th align="Center" ><font size=<% =fontsize2 %>  color="#FFFFFF"><b><br>Fed</b></FONT></TD>
		<Th align="Center" ><font size=<% =fontsize2 %> color="#FFFFFF"><b><br>Pctile</b></FONT></Th>
	<% END IF 

    	IF sRunByWhat <> "NSL" AND sRunByWhat <> "NCWSA" THEN 
	    	IF LCASE(TRIM(sLeagueSelected))<>"none" THEN	
		     %><Th align="Center"><font size=<% =fontsize2 %> color="#FFFFFF"><a title="View Qualifications"><b>Event<br>Status<br><%=sLeagueSelected%></b></a></FONT></Th><%
		ELSE %>
		     <Th align="Center"><font size=<% =fontsize2 %> color="#FFFFFF"><a title="View Qualifications"><b>Event<br>Status</b></a></FONT></Th><%
		END IF
	END IF
    %></TR><%



END SUB




' -----------------------------
   SUB DefineRankingDataLine
' -----------------------------

' Extract key items from the Query Answerset for this Member's Ranking

tName = TRIM(rs("LastName")) & ", " & rs("FirstName")
tMemberID = rs("MemberID")
tState = rs("State")
tRegion = rs("Region")
tMemberFed = rs("FederationCode")
tDefaultYear = rs("DefaultYear")
tTeam = rs("Team")
tTeamStat = rs("TeamStat")
tNCWRegn = rs("NCWRegion")
tNCWConf = rs("NCWConf")
tL10Div = rs("L10Div")

If len(trim(rs("OpenDate"))) > 0 THEN
   if len(trim(rs("MastDate"))) > 0 THEN
      tEliteStat = tOpenDiv & "/" & tMastDiv
      tEliteBkup = tOpenDiv & " Qual fm " & rs("OpenDivOrig") & " thru " & rs("OpenDate") & Chr(13) & Chr(10) & tMastDiv & " Qual fm " & rs("MastDivOrig") & " thru " & rs("MastDate")
   ELSE
      tEliteStat = tOpenDiv
      tEliteBkup = tOpenDiv & " Qual fm " & rs("OpenDivOrig") & " thru " & rs("OpenDate")
   END IF
ELSE   
   if len(trim(rs("MastDate"))) > 0 THEN
      tEliteStat = tMastDiv
      tEliteBkup = tMastDiv & " Qual fm " & rs("MastDivOrig") & " thru " & rs("MastDate")
   ELSEIF tDefaultYear = 0 and (DivSelected = "OM" or DivSelected = "OW" or DivSelected = "MM" or DivSelected = "MW") THEN
      tEliteStat = "None"
      tEliteBkup = "Not Elite Qualified in this Event"
   ELSE
      tEliteStat = ""
      tEliteBkup = ""
   END IF
END IF

' -- Adds marker for Level 10
IF LEN(tL10Div)>0 THEN 
		IF LEN(tEliteStat)>0 THEN
				tEliteStat = tEliteStat & "/" & tL10Div
		ELSE
				tEliteStat = tL10Div
		END IF		 
END IF

IF sRunByWhat = "NSL" THEN tRankScore = rs("sc_3") ELSE tRankScore = rs("RankScore")

tRnkScoBkup = rs("RnkScoBkup")
tRating = rs("AWSA_Rat")

IF rs("RankPct")>0 THEN
		tRankPct=FormatNumber(rs("RankPct"),2) 
	ELSE
		tRankPct=""
	END IF		

IF left(tRating,1)= EventSelected THEN
		tLevelNo=RIGHT(tRating,1)
	ELSE
		tLevelNo=0
	END IF		

tNatPlace = TRIM(rs("natl_plc"))
tRegPlace = TRIM(rs("regl_plc"))
tRegSki = rs("reg_ski")

SELECT CASE tRegion
	CASE 1
		tRegion = "C"
	CASE 2
		tRegion = "M"
	CASE 3
		tRegion = "W"
	CASE 4
		tRegion = "S"
	CASE 5
		tRegion = "E"
END SELECT


IF EventSelected = "S" THEN
	tFmtScore = FormatNumber(tRankScore,2)
ELSE
	IF EventSelected = "O" or EventSelected = "J" THEN
		tFmtScore = FormatNumber(tRankScore,1)
	ELSE
		tFmtScore = FormatNumber(tRankScore,0)
	END IF
END IF
       



' --- Establishes background color for the current record
' --- Also determines the current Level Number, as a function
' --- of the "Rank_Level" value and the DT Percents

IF pvar="NCWSA" THEN 
		DefineLevelcolor=scolor06
	ELSE 
		SELECT CASE tLevelNo
			CASE 0
				DefineLevelcolor=scolor01
			CASE 1
				DefineLevelcolor=scolor01
			CASE 2
				DefineLevelcolor=scolor02
			CASE 3
				DefineLevelcolor=scolor03
			CASE 4
				DefineLevelcolor=scolor04
			CASE 5
				DefineLevelcolor=scolor05
			CASE 6
				DefineLevelcolor=scolor06
			CASE 7
				DefineLevelcolor=scolor07
			CASE 8
				DefineLevelcolor=scolor08
			CASE 9
				DefineLevelcolor=scolor09
		END SELECT
	END IF

END SUB




' ------------------------
  SUB CheckValidDivision
' ------------------------

' ----------------------------------------------------------------------------------------------------------------
' -------- If division is not in query list for sRunByWhat then set DivisionSelected to the FIRST division
' --------   found in Rankings Table meeting the filtering parameters for the current sRunByWhat 
' ----------------------------------------------------------------------------------------------------------------

SET rsSelectFields=Server.CreateObject("ADODB.recordset")
RunDivQuery
sSQL = sSQL + " AND DIV='"&DivSelected&"'"
rsSelectFields.open sSQL, SConnectionToTRATable

' --- Not found so reset DivisionSelected to first one found in Rankings table---
IF rsSelectFields.eof THEN 
	SET rsSelectFields=Server.CreateObject("ADODB.recordset")
	RunDivQuery
	sSQL = sSQL + " order by div"
	rsSelectFields.open sSQL, SConnectionToTRATable

	IF NOT rsSelectFields.eof THEN DivSelected=rsSelectFields("Div")
END IF

rsSelectFields.close


END SUB


' -----------------
  SUB RunDivQuery
' -----------------

sSQL = "Select top 1 div from " & RankTableName

SELECT CASE sRunByWhat
	CASE "Junior"
			sSQL = sSQL + " WHERE LOWER(left(div,1)) in ('b','g')"
	CASE "NSL"
			sSQL = sSQL + " WHERE LOWER(left(div,1)) in ('x','y')"
	CASE "NCWSA"
			sSQL = sSQL + " WHERE LOWER(left(div,1)) = 'c'"
	CASE ELSE
			IF Session("AdminMenuLevel")>0 THEN
					sSQL = sSQL + " WHERE (lower(left(RT.div,1)) in ('b','g','m','w','o','e') or lower(RT.Div) = 'sm')"
			ELSE
					sSQL = sSQL + " WHERE lower(left(RT.div,1)) in ('b','g','m','w','o')"
			END IF
END SELECT

END SUB




'--------------------------- 
   SUB RunStandingsQueryNew
'--------------------------- 

'	Creates and execute SQL query against Rankings Table for Selected Division/Event
'	Newly restructured Feb 2008, one single universal query for all ranking types


    IF instr("MW",Right(DivSelected,1)) > 0 THEN
       tOpenDiv = "O" & Right(DivSelected,1)
    ELSEIF instr("BM",left(DivSelected,1)) > 0 THEN
       tOpenDiv = "OM"
    ELSE
       tOpenDiv = "OW"
    END IF
    tMastDiv = "M" & right(tOpenDiv,1)

    sSQL = "Select distinct MT.lastname, MT.firstname, MT.federationcode, MT.state, MT.Region," 

	' Explore removing reference to DT entry -- Up_Age appears unused. -- Ditto for SY table mebbe ????

    sSQL = sSQL & " RT.*, SY.DefaultYear, TT.NCWRegion, TT.NCWConf,"
    sSQL = sSQL & " Coalesce(Convert(char(10),OQ.QualThru,111),'') as OpenDate,"
    sSQL = sSQL & " Coalesce(Convert(char(10),MQ.QualThru,111),'') as MastDate,"
		sSQL = sSQL & " OQ.DivOrig as OpenDivOrig, MQ.DivOrig as MastDivOrig,"
		sSQL = sSQL & " CASE WHEN L10.MemberID IS NOT NULL THEN 'L10' ELSE '' END AS L10Div,"		
'    sSQL = sSQL & " RQ.QfyStatus, RQ.QfyStatusTEMP, RQ.TourID AS RQTourID, Coalesce(LT.HomoType,'-') AS HomoType, LT.Status AS TourStatus"
    sSQL = sSQL & " RQ.QfyStatus, RQ.TourID AS RQTourID, Coalesce(LT.HomoType,'-') AS HomoType, LT.Status AS TourStatus"

    sSQL = sSQL & " FROM "&RankTableName&" as RT"
    ' sSQL = sSQL & " JOIN "&MemberTableName&" as MT on RT.memberid = MT.personidwithcheckdigit "
		
		sSQL = sSQL + "	JOIN "&MemberShortTableName&" AS MT ON CAST(RIGHT(RT.MemberID,8) AS INT)=MT.PersonID"
		

    sSQL = sSQL & " JOIN " & SkiYearTableName & " as SY on SY.SkiYearID = RT.SkiYearID"

    sSQL = sSQL & " LEFT JOIN " & TeamTableName & " as TT on TT.TeamID = RT.Team"

    sSQL = sSQL & " LEFT JOIN " & EliteDateTableName & " as OQ on OQ.MemberID = RT.MemberID"
    sSQL = sSQL & " and OQ.Event = RT.Event and OQ.DivElite = '" & tOpenDiv & "'"
    sSQL = sSQL & " and OQ.SkiYearID = " & Session("SkiYear")
    
    sSQL = sSQL & " LEFT JOIN " & EliteDateTableName & " as MQ on MQ.MemberID = RT.MemberID"
    sSQL = sSQL & " and MQ.Event = RT.Event and MQ.DivElite = '" & tMastDiv & "'"
    sSQL = sSQL & " and MQ.SkiYearID = " & Session("SkiYear")

		sSQL = sSQL & " LEFT JOIN " 
		sSQL = sSQL & " 	( SELECT MemberID, Event, Div FROM " & EquivLevel10TableName & " AS L" 
		sSQL = sSQL & " 			JOIN " & SkiYearTableName & " AS SY2 ON SY2.SkiYearID = L.SkiYearID"
' 		sSQL = sSQL & " 				WHERE DefaultYear=1 AND ( (Sent_Notice='Y' AND Event<>'O') OR (Sent_Notice='N' AND Event='O') ) ) L10 "
 		sSQL = sSQL & " 				WHERE DefaultYear=1 AND Sent_Notice='Y') L10 "
    sSQL = sSQL & " ON L10.MemberID=RT.MemberID AND L10.Event=RT.Event"
    sSQL = sSQL & "      AND L10.Div=CASE WHEN RT.Div='OM' THEN 'EM'"
    sSQL = sSQL & "           WHEN RT.Div='OW' THEN 'EW'"
    sSQL = sSQL & "           WHEN RT.Div='MM' THEN 'SM'"
    sSQL = sSQL & "           WHEN RT.Div='MW' THEN 'SW' END"    
            
'    sSQL = sSQL & " LEFT JOIN "&RegGenTableName&" AS RGEN ON RGEN.MemberID=RT.MemberID AND LEFT(RGEN.TourID,6)='07W999'" 	
'    sSQL = sSQL & " LEFT JOIN "&RegionTableName&" as RG ON lower(MT.state) = lower(RG.state) "

    sSQL = sSQL & " LEFT JOIN "&LeagueTableName&" AS LT ON LT.LeagueID='"&sLeagueSelected&"'"
    sSQL = sSQL & " LEFT JOIN "&RegQualifyTableName&" AS RQ ON LEFT(RQ.TourID,6)=LEFT(LT.QualifyTour,6) AND RQ.MemberID=RT.MemberID AND RQ.Event=RT.Event AND RQ.Div=RT.Div"
	
    sSQL = sSQL & " WHERE RT.div = '" & DivSelected & "'"
'    sSQL = sSQL & " AND RT.DivType <> 'D'"
    sSQL = sSQL & " AND RT.[event] = '" & EventSelected & "'"
    sSQL = sSQL & " AND RT.SkiYearID = " & Session("SkiYear")

		IF Left(FilterSelected,1) = "1" THEN
				sSQL = sSQL & " AND TT.NCWRegion = '" & Mid(FilterSelected,2) & "'"
		ELSEIF Left(FilterSelected,1) = "2" THEN
				sSQL = sSQL & " AND TT.NCWConf = '" & Mid(FilterSelected,2) & "'"
		ELSEIF Left(FilterSelected,1) = "3" THEN
				sSQL = sSQL & " AND MT.Region = '" & Mid(FilterSelected,2) & "'"
		ELSEIF Left(FilterSelected,1) = "4" THEN
				sSQL = sSQL & " AND MT.State = '" & Mid(FilterSelected,2) & "'"
		END IF

    IF FederationSelected = "USA" THEN sSQL = sSQL & " AND MT.federationcode = 'USA'"

    ' --- NSL (still needed ?) --- 
    IF sRunByWhat = "NSL" THEN 
    		sSQL = sSQL & " AND RT.sc_3 is not NULL ORDER by RT.sc_3 DESC" 
    ELSE
    		sSQL = sSQL & " AND RT.RankScore is not NULL ORDER by RT.RankScore DESC" 
    END IF


		IF Session("Adminmenulevel")=50 AND sShowSQL="on" THEN 
				response.write(sSQL)
				'response.end
		END IF


		' -------------------------
		' --- Execute SQL query ---
		' -------------------------
		OpenCon
		SET rs=Server.CreateObject("ADODB.recordset")
		'rs.open sSQL, SConnectionToTRATable
    ChoosePagesSQL sSQL,currentPage, 30  



rowCount = 0
'Response.Write("<BR>")



END SUB



' --------------------
  SUB FindCOAScoreAll
' --------------------

' --------------------------------------------------------------------------------
' --- Look up the Level Percents and COA Scores for the selected Division / Event
' --------------------------------------------------------------------------------


IF sRunByWhat <>"NCWSA" AND sRunByWhat<>"NSL" THEN

' Here is the new logic for this section -- DC 20080131

'	Get Values from new CutOffTable

	IF EventSelected="S" THEN i=2 ELSE IF EventSelected="J" THEN i=1 ELSE i=0

	Set rs = Server.CreateObject("ADODB.recordset")

	sSQL = " SELECT TOP 1 * FROM " & CutOffTableName  
	sSQL = sSQL + " WHERE Div = '" & DivSelected & "' AND Event = '" & EventSelected & "' AND SkiYearID = " & Session("SkiYear")

	'WriteDebugSQL(sSQL)

	rs.open sSQL, SConnectionToTRATable

	IF NOT rs.EOF THEN 
		tPerc3 = rs("Pct3"): COALevel3 = FormatNumber(rs("COA3"),i)
		tPerc4 = rs("Pct4"): COALevel4 = FormatNumber(rs("COA4"),i)
		tPerc5 = rs("Pct5"): COALevel5 = FormatNumber(rs("COA5"),i)
		tPerc6 = rs("Pct6"): COALevel6 = FormatNumber(rs("COA6"),i)
		tPerc7 = rs("Pct7"): COALevel7 = FormatNumber(rs("COA7"),i)
		tPerc8 = rs("Pct8"): COALevel8 = FormatNumber(rs("COA8"),i)
		tPerc9 = rs("Pct9"): COALevel9 = FormatNumber(rs("COA9"),i)
	ELSE
		tPerc3 = 0: COALevel3 = "0"
		tPerc4 = 0: COALevel4 = "0"
		tPerc5 = 0: COALevel5 = "0"
		tPerc6 = 0: COALevel6 = "0"
		tPerc7 = 0: COALevel7 = "0"
		tPerc8 = 0: COALevel8 = "0"
		tPerc9 = 0: COALevel9 = "0"
	END IF

	rs.close

END IF

END SUB




' --------------------------------------------------------------------------------------------
   SUB DisplayDropdowns      ' -----   Begin form for selection /  filtering parameters ------
' --------------------------------------------------------------------------------------------

Titlecolor=Textcolor2

%>

<TABLE class="droptable" align=center width=98% height=215 background="<%=MainImage%>" ><% '---Table to hold image --- %>
  <TR>
    <TD >
    <% ' ------ OUTER TABLE TO HOLD BACKGROUND IMAGE %>

<TABLE width=100% align=center>

<%

SELECT CASE sRunByWhat
  CASE "National"
    RankType="National"
  CASE "NSL"
    RankType="Grassroots"
  CASE "NCWSA"
    RankType="NCWSA "
  CASE "Junior"
    RankType="Junior"
END SELECT	


' -----------------------------  Build Ranking List Type Radio Button list  ----------------------------
' If this form kicks off then it posts ONLY a new sRunByWhat value, that in turn delivers new drop-downs

%>

<form method=post action="<%=ThisFileName%>?SkiYear=<%=session("SkiYear")%>">
  <input type="hidden" name="adminmenulevel" value="<%=adminmenulevel%>">
  <input type="hidden" name="sMemberID" value="<%=sMemberID%>">


<tr>
  <td colspan=3 valign="top" align="left">
	<FONT size=4 face=<% =font1 %> color=<% =textcolor2 %>><b>&nbsp;<I><%=RankType%> Rankings</I></b></font>
  </td>

  <td colspan=3 align="left">
	<input type=radio NAME=pvar VALUE="National" <% IF pvar="National" THEN response.write "checked" %> onclick=submit()>
	<FONT size=<% =fontsize3 %> face=<% =font1 %> color=<% =textcolor2 %>><b>National&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></font>
	<%
	IF 2=1 THEN
			%>
			<input type=radio NAME=pvar VALUE="Junior" <% IF pvar="Junior" THEN response.write "checked" %> onclick=submit()>
			<FONT size=<% =fontsize3 %> face=<% =font1 %> color=<% =textcolor2 %> checked><b>Junior&nbsp;&nbsp;&nbsp;</b></font>
			<br>
			<%
	END IF
	%>
	<input type=radio NAME=pvar VALUE="NCWSA" <% IF pvar="NCWSA" THEN response.write "checked" %> onclick=submit()>
	<FONT size=<% =fontsize3 %> face=<% =font1 %> color=<% =textcolor2 %> checked><b>Collegiate&nbsp;&nbsp;&nbsp;</b></font>
	<%

	' -- Dropped ability to access grassroots 6-28-2014 --
	ShowGrassRootsRadio="N"
	IF ShowGrassRootsRadio="Y" THEN
			%>
			<input type=radio NAME=pvar VALUE="Grassroots" <% IF pvar="grassroots" THEN response.write "checked" %> onclick=submit()>
			<FONT size=<% =fontsize3 %> face=<% =font1 %> color=<% =textcolor2 %> checked><b>Grassroots&nbsp;&nbsp;&nbsp;</b></font>
			<%
	ELSE
		%>
			&nbsp;
			<%
	END IF
	%>
	<br><br>
  </td>
</tr>
<% 
IF x=1 THEN
%>
	<input type=radio NAME=pvar VALUE="NSL" <% IF pvar="NSL" THEN response.write "checked" %> onclick=submit()>
	<FONT size=<% =fontsize3 %> face=<% =font1 %> color=<% =textcolor2 %>><b>Grassroots</b></font>
<%
END IF 

%>

</form>

<form method=post action="<%=ThisFileName%>?pvar=<%=sRunByWhat%>">
  <input type="hidden" name="adminmenulevel" value="<%=adminmenulevel%>">
  <input type="hidden" name="sMemberID" value="<%=sMemberID%>">

<% ' --------------------------------- Build SKI YEAR dropdown list  ------------------- %>

<tr>
  <td width=11% align="center">
    <font size="<%=fontsize2%>" face=<% =font1 %> color=<%=Titlecolor%>><b>Period:</b></font>
  </td>

  <td colspan=2>	
	<SELECT name='SkiYear'><%

		SET rsSelectFields=Server.CreateObject("ADODB.recordset")
		sSQL = "SELECT DISTINCT RT.SkiYearID, SY.SkiYearName"
		sSQL = sSQL + " FROM " &RankTableName&" AS RT"
		sSQL = sSQL + " JOIN " &SkiYearTableName&" AS SY ON RT.SkiYearID = SY.SkiYearID"

		' --- NCWSA does not display 12 Month Rankings
		IF sRunByWhat="NCWSA" THEN
				sSQL = sSQL + " WHERE SY.SkiYearID <> 1"
		END IF

		rsSelectFields.open sSQL, SConnectionToTRATable

		' Loads dropdown and sets default to Session("SkiYear")
		DO WHILE NOT rsSelectFields.eof

			IF TRIM(rsSelectFields("SkiYearID")) = TRIM(session("SkiYear")) THEN
					response.write("<option value =""" & rsSelectFields("SkiYearID") &""" selected>")
					response.write(rsSelectFields("SkiYearName"))
					response.write("</option><br>")
			ELSE
			response.write("<option value =""" & rsSelectFields("SkiYearID") &""">")
			response.write(rsSelectFields("SkiYearName"))
			response.write("</option><br>")
		END IF 

		rsSelectFields.moveNEXT

	LOOP

	rsSelectFields.close %>
	</select>
  </td><%



  IF sRunByWhat="NSL" OR sRunByWhat="NCWSA" THEN %>
		<td>&nbsp;</td>
		<td>&nbsp;</td><%
  ELSE 
		IF TRIM(session("SkiYear"))=1 THEN %>
		  <td colspan=2 align=left>
			 <font size=<% =fontsize2 %> face=<% =font1 %> color=<%=Textcolor2%>><b>Set Event Qualifications Filter&nbsp;</b></font><%
			%>
		  </td><%
		ELSE %>
		  <td>&nbsp;</td>
		  <td>&nbsp;</td><%
		END IF
  END IF %>	

<td>&nbsp;</td> 
</tr>


<% ' ---- Preloads Event dropdown with values based on sRunByWhat variable passed from Menu Link ---- %>

<tr>

  <td align="center"> 
     <font size=<% =fontsize2 %> face=<% =font1 %> color=<%=Titlecolor%>><b>Event:</b></font>
  </td>

  <td colspan=2 width=25%>	
	<select name='event'>
	  <option value ='S' <%IF EventSelected="S" THEN response.write(" selected")%>>Slalom</Option><br>
	  <option value ='J' <%IF EventSelected="J" THEN response.write(" selected")%>>Jump</Option><br>
	  <option value ='T' <%IF EventSelected="T" THEN response.write(" selected")%>>Trick</Option><br>

	  <% IF sRunByWhat <> "NSL" AND sRunByWhat <> "NCWSA" THEN %>
	  	<option value ='O' <%IF EventSelected = "O" THEN Response.Write(" selected")%>>Overall</option><br>
	  <% END IF %>

	</select>
  </td>


  <td align=left colspan=2><%
	IF sRunByWhat <> "NSL" AND sRunByWhat <> "NCWSA" AND TRIM(session("SkiYear"))=1 THEN 
		' --- Procedure found in Tools_Leagues.asp ---
		BuildLeagueDrop true, "None" 
	ELSE %>
		&nbsp;<%
	END IF  %>
  </td>

  <td>&nbsp;</td><%


' ---- OBSOLETE ?? The value of the dropdown is established based on Session("SkiYear") variable ---- 




' ------------------------------  Build DIVISION dropdown list  ---------------------------------- 

%>

<tr>
  <td align="center">
    <font size=<% =fontsize2 %> face=<% =font1 %> color=<%=Titlecolor%>><b>Division:</b></font>
  </td>

  <td colspan=2>
		<select name='DivSelected'><%

		SET rsSelectFields=Server.CreateObject("ADODB.recordset")
		sSQL = "Select distinct RT.div, DT.div_name from "&RankTableName&" as RT JOIN "&DivisionsTableName&" as DT ON RT.div = DT.div"

		SELECT CASE sRunByWhat
  			CASE "National"
						IF Session("AdminMenuLevel")>0 THEN
								sSQL = sSQL + " WHERE (lower(left(RT.div,1)) in ('b','g','m','w','o','e') or lower(RT.Div) in ('sm','sw'))"
						ELSE
								sSQL = sSQL + " WHERE lower(left(RT.div,1)) in ('b','g','m','w','o')"
						END IF
				CASE "NSL"
    				sSQL = sSQL + " WHERE lower(left(RT.div,1)) in ('y','x')"
	  		CASE "Junior"
    				sSQL = sSQL + " WHERE lower(left(RT.div,1)) in ('b','g')"
	  		CASE "NCWSA"
						sSQL = sSQL + " WHERE lower(left(RT.div,1)) = 'c'"
		END SELECT

		sSQL = sSQL + " order by RT.div"
		rsSelectFields.open sSQL, SConnectionToTRATable


		' ---  This section deals with case WHERE no scores exist for any of the divisions  ---
		IF NOT rsSelectFields.eof THEN 
	  		rsSelectFields.movefirst
  			DO WHILE NOT rsSelectFields.eof
	    			IF TRIM(rsSelectFields("Div")) = DivSelected THEN
      					response.write("<option value ="""&rsSelectFields("Div")&""" selected>"&rsSelectFields("Div")&" - "&rsSelectFields("Div_Name")&"</option><br>")
    				ELSE
      					response.write("<option value ="""&rsSelectFields("Div")&""">"&rsSelectFields("Div")&" - "&rsSelectFields("Div_Name")&"</option><br>")
	    			END IF	
						rsSelectFields.moveNEXT
  			LOOP
		ELSE
  			response.write("<option value =""None"" selected>None</option>")
		END IF

		rsSelectFields.close 
		%>
		</select>
  </td>
  <% 

	CutOffAvg = "98.76"
	CutOffDate = "8/15/2017"
	

	sSQL = "SELECT * FROM " &LeagueQfyTableName&" AS LQ"
	sSQL = sSQL + " JOIN " &LeagueTableName&" AS L ON L.LeagueID=LQ.LeagueID"
	sSQL = sSQL + " WHERE LQ.LeagueID='"&sLeagueSelected&"'"
	sSQL = sSQL + " AND Div='"&DivSelected&"'"	
	sSQL = sSQL + " AND Event='"&EventSelected&"'"
	
	SET rsCOA=Server.CreateObject("ADODB.recordset")
	rsCOA.open sSQL, SConnectionToTRATable
	
	' CutOffAvg=0.00
	IF NOT rsCOA.eof THEN
			CutOffAvg = rsCOA("COA")
			CutOffDate = rsCOA("COD")
	END IF

		
		
	IF TRIM(session("SkiYear"))=1 AND TRIM(LCASE(sLeagueSelected))<>"none" AND TRIM(sLeagueSelected)<>"" THEN 
			%>
			<td colspan=2 align="left">
		  	<font size=<% =fontsize2 %> face=<% =font1 %> color=<%=Titlecolor%>><b>Cut-Off Avg:</b> </font>
    		<font size=<% =fontsize2 %> face=<% =font1 %> color="<%=Textcolor3%>">&nbsp;<b><%=CutOffAvg%></b></font>
			</td>
			<%
	ELSE 
			%>
			<td colspan=2>&nbsp;</td>
			<td>&nbsp;</td>
			<%
	END IF

	IF Session("AdminMenuLevel")>=50 THEN  
			%>	
  		<td colspan=2 width=350 valign=top align="left">
				<FONT color="<% =Titlecolor %>" size="<% =fontsize2 %>"><b>Show SQL</b></font>
				<input type=checkbox name="sShowSQL" <% IF sShowSQL="on" THEN response.write "checked" %>>
			</td>
			<%
	ELSE 
			%><td colspan=2>&nbsp;</td><%
	END IF 
	%>



</tr>


<% ' ----  Build FilterSelected  REGION or STATE (AWSA)  or  REGION OR CONFERENCE (NCWSA)  dropdown list  ---- %>

<tr>
  <td align="center">
	  <font size=<% =fontsize2 %> face=<% =font1 %> color=<%=Titlecolor%>>
	  <b>
		<% 
		IF sRunByWhat = "NCWSA" THEN 
				%>Region or<br>Conference:<% 
		ELSE 
				%>Region<br>or State:<% 
		END IF 
		%>
		</b>
		</font>
	</td>
   	
  <td colspan=2>
		<select name='FilterSelected'>
			<% 
			IF sRunByWhat = "NCWSA" THEN 
					%>
					<option value ='All'  <%IF FilterSelected="All"  THEN response.write(" selected")%>>All</Option><br>
	  			<option value ='1E'  <%IF FilterSelected="1E"  THEN response.write(" selected")%>>Eastern Region</Option><br>
	  			<option value ='2NE' <%IF FilterSelected="2NE" THEN response.write(" selected")%>>.. Northeast Conf</Option><br>
	  			<option value ='2SA' <%IF FilterSelected="2SA" THEN response.write(" selected")%>>.. So Atlantic Conf</Option><br>
	  			<option value ='2SO' <%IF FilterSelected="2SO" THEN response.write(" selected")%>>.. Southern Conf</Option><br>
	  			<option value ='1M'  <%IF FilterSelected="1M"  THEN response.write(" selected")%>>Midwest Region</Option><br>
	  			<option value ='2GL' <%IF FilterSelected="2GL" THEN response.write(" selected")%>>.. Great Lakes Conf</Option><br>
	  			<option value ='2GP' <%IF FilterSelected="2GP" THEN response.write(" selected")%>>.. Great Plains Conf</Option><br>
	  			<option value ='1SC' <%IF FilterSelected="1SC" THEN Response.Write(" selected")%>>South Central Region</option><br>
	  			<option value ='1W'  <%IF FilterSelected="1W"  THEN response.write(" selected")%>>Western Region</Option><br>
	  			<option value ='2NW' <%IF FilterSelected="2NW" THEN response.write(" selected")%>>.. Northwest Conf</Option><br>
	  			<option value ='2PC' <%IF FilterSelected="2PC" THEN response.write(" selected")%>>.. Pacific Conf</Option><br>
	  			<% 
	  	ELSE 
	  			%>
	  			<option value ='All'  <%IF FilterSelected="All"  THEN response.write(" selected")%>>All</Option><br>
					<% 

					sSQL = "SELECT CASE When Region = '5' then 'Eastern' When Region = '2' then 'Midwest'"
					sSQL = sSQL + " When Region = '1' then 'South Central' When Region = '4' then 'Southern'"
					sSQL = sSQL + " When Region = '3' then 'Western' else 'Unknown' end as RegionName,"
					sSQL = sSQL + " Region, State, StateName FROM " & RegionTableName 
					sSQL = sSQL + " Order by Case When Region = '5' then 'E' When Region = '2' then 'M'"
					sSQL = sSQL + " When Region = '1' then 'P' when Region = '4' then 'S'"
					sSQL = sSQL + " When Region = '3' then 'W' else 'Z' end, StateName;"

					rsSelectFields.open sSQL, SConnectionToTRATable

					LastRegion = "0"
					DO WHILE NOT rsSelectFields.eof

							IF Trim(rsSelectFields("Region")) <> LastRegion THEN
									LastRegion = Trim(rsSelectFields("Region")) 
									%><option value ='3<%=LastRegion%>'<%IF FilterSelected = "3"&LastRegion THEN Response.Write(" selected ")%>><%=rsSelectFields("RegionName")&" Region"%></Option><br><% 
							END IF 
							%>
							<option value ='4<%=Trim(rsSelectFields("State"))%>'<%IF FilterSelected = "4"&Trim(rsSelectFields("State")) THEN Response.Write(" selected ")%>>... <%=rsSelectFields("StateName")%></Option><br>
							<% 
							
							rsSelectFields.moveNEXT

					LOOP

					rsSelectFields.close

			END IF 
			
			
			%>    
		</select>
  </td>
	<%
	IF TRIM(session("SkiYear"))=1 AND TRIM(LCASE(sLeagueSelected))<>"none" AND TRIM(sLeagueSelected)<>"" THEN 
			%>
			<td colspan=2 align="left">
		  	<font size=<% =fontsize2 %> face=<% =font1 %> color=<%=Titlecolor%>><b>Cut-Off Date:</b> </font>
    		<font size=<% =fontsize2 %> face=<% =font1 %> color="<%=Textcolor3%>">&nbsp;<b><%=CutOffDate%></b></font>
			</td>
			<%
	ELSE 
			%>
			<td colspan=2>&nbsp;</td>
			<td>&nbsp;</td>
			<%
	END IF
	%>
</tr>



<tr>
  <td align="center">
    <font size=<% =fontsize2 %> face=<% =font1 %> color=<%=Titlecolor%>><b>Federation:</b></font>
  </td>

  <td colspan=2>

	<%' --------------------------------  Build FEDERATION dropdown list  -------------------- %>
	<select name="Include_International">
	<option value="ALL"<%IF FederationSelected = "ALL" THEN Response.Write(" selected")%>>All Federations</option>
	<option value="USA"<%IF FederationSelected = "USA" THEN Response.Write(" selected")%>>USA Only</option>
	</select>
  </td>
  <td>&nbsp;</td>
  <td>&nbsp;</td>
  <td>&nbsp;</td>
</tr>


<tr>
	<td align="center"><font size=<% =fontsize2 %> face=<% =font1 %> color=<%=Titlecolor%>><b>Member:</b></font>
	<td colspan=2 align="left"><font size=<% =fontsize2 %> face=<% =font1 %> color=<%=Textcolor3%>><b>&nbsp;&nbsp;<%=sFullName%></b></font></td>
	<% 
	IF DivSelected = "" THEN 
			%>
   		<td align="center">
   			<input type=submit style="width:10em" value="Display Rankings" title="Display Rankings for the selection parameters above/left">
   		</td>
			<% 
	ELSE 
			%>
   		<td align="center">
   			<input type=submit style="width:10em" value="Update Display" title="Display a revised Rankings page after you have &#13;changed the selection parameters above/left">
   		</td>
   		<% 
  END IF 


	%>
  <td align="center">
  	<input type=submit style="width:10em" name="SingleRanking" value="Find My Rankings" title="TEMPORARILY OUT OF SERVICE (Look up my Rankings in all Events)" enabled>
	</td>
	<%


	IF sRunByWhat = "NCWSA" THEN
			%><td align=center><a title="Tips and Explanations and Insights and &#13;Answers to Frequently Asked Questions" onclick="window.open('/rankings/news/FAQ_NCWRankings.htm', '_blank');"><input type="submit" style="width:10em" name="thisaction" value="FAQ/Tips"></a></td><%
	ELSE 
			%><td align=center><a title="Tips and Explanations and Insights and &#13;Answers to Frequently Asked Questions" onclick="window.open('/rankings/news/FAQ_Rankings.htm', '_blank');"><input type="submit" style="width:10em" name="thisaction" value="FAQ/Tips"></a></td><%
	END IF 
	%>


  	</form>
	</tr>
</table>

</TD>
</TR>
</TABLE><% ' --- Table to hold picture ---


END SUB





' --------------------------
  SUB FindRankingInstances
' -------------------------- 

'response.write("<br>MADE IT - Line 1340 Top of FindRankingInstances")
'response.end


sSkiYearID=TRIM(Session("SkiYear"))

sSQL = "SELECT * FROM "&RankTableName&" AS RT"
'sSQL = sSQL + " JOIN "&MemberTableName&" AS MT ON MT.PersonIDwithCheckDigit=RT.MemberID"
sSQL = sSQL + "	JOIN "&MemberShortTableName&" AS MT ON CAST(RIGHT(RT.MemberID,8) AS INT)=MT.PersonID"
sSQL = sSQL + " WHERE MemberID='"&sMemberID&"' AND SkiYearID='"&sSkiYearID&"'"

SELECT CASE sRunByWhat
	CASE "National"
			IF Session("AdminMenuLevel")>50 THEN
				sSQL = sSQL + " AND (lower(left(RT.div,1)) in ('b','g','m','w','o','e') or lower(RT.Div) = 'sm')"
			ELSE
				sSQL = sSQL + " AND lower(left(RT.div,1)) in ('b','g','m','w','o')"
			END IF
	CASE "NSL"
    	sSQL = sSQL + " AND lower(left(RT.div,1)) in ('y','x')"
	CASE "Junior"
  			sSQL = sSQL + " AND lower(left(RT.div,1)) in ('b','g')"
	CASE "NCWSA"
			sSQL = sSQL + " AND lower(left(RT.div,1)) = 'c'"
END SELECT

sSQL = sSQL + " ORDER BY RT.Div, CASE when RT.Event='S' then 1 when RT.event='T' then 2 when RT.Event='J' then 3 else 4 end"


'response.write("<br>Line 1369 After SQL build")
'response.write("<br><br>"& sSQL)
'response.end

' --- Error trap for missing Tour and Member ---
Dim QuerySuccess
QuerySuccess=false
SET rsRankList=Server.CreateObject("ADODB.recordset")
IF TRIM(sMemberID)<>"" AND TRIM(sSkiYearID)<>"" THEN
		rsRankList.open sSQL, SConnectionToTRATable
		IF NOT(rsRankList.eof) THEN QuerySuccess=true
ELSE
		ErrorSubject="ERROR -  Line 1374 - View-StandingsHQ - TRIM(sMemberID) IS NULL OR TRIM(sSkiYearID) IS NULL"
		ErrorSQL = sSQL
		SendErrorEmailToMark ErrorSubject, ErrorSQL
END IF

'response.write("<br><br>Line 1392 After Query - QuerySuccess="&QuerySuccess)
'response.end




IF QuerySuccess=true THEN 

		'response.write("<br><br>Line 1393 Inside QuerySuccess=true IF <br>")
'		response.end
	
		%>
		<TABLE class="innertable" width=500px align="center"><% '--- Table for border --- %>
		<TR>
  		<TH>
    		<center><font size=<% =fontsize3 %> color="#FFFFFF"><b> National Rankings For</font>
      		<font size=4 color="white"><br><%=rsRankList("FirstName")%>&nbsp;<%=rsRankList("LastName")%></b></FONT>
      		<font size=2 color="#FFFFFF"><br><b><%=rsRankList("City")%>, <%=rsRankList("State")%></b></FONT>
    		</center>
  		</TH>
		</TR>
		<TR>
  		<TD>  
    		<TABLE class="innertable" align="Center" width="90%">
	    		<tr>
    				<th align="Center" width=20% valign="center"><font size=<% =fontsize2 %> color="#FFFFFF"><b>National<br>Rank</b></FONT></th>
    				<th align="Center" width=30% valign="center"><font size=<% =fontsize2 %>  color="#FFFFFF"><b>Division<br>and Event</b></FONT></th> 	
    				<th align="Center" width=30% valign="center"><font size=<% =fontsize2 %>  color="#FFFFFF"><b>Ranking<br>Score</b></FONT></th> 	
    			</tr>
    			<%
 	
					DO WHILE NOT rsRankList.eof 

							sEvent=TRIM(rsRankList("Event"))

							IF TRIM(rsRankList("RankScore"))<>"" THEN
				  				SELECT CASE TRIM(rsRankList("Event"))
										CASE "J"
											sEventName="Jump"
											sRankScore=formatnumber(rsRankList("RankScore"),1)		
										CASE "S"
											sEventName="Slalom"
											sRankScore=formatnumber(rsRankList("RankScore"),2)		
										CASE "T"
											sEventName="Trick"
											sRankScore=formatnumber(rsRankList("RankScore"),0)
										CASE "O"
											sEventName="Overall"
											sRankScore=formatnumber(rsRankList("RankScore"),0)
	  								END SELECT
        			ELSE
									sEventName="No Ranking"
									sRankScore=formatnumber(0,0)
 							END IF

							%>
    					<tr>
	  						<td align="Center" width=9% valign="top" bgcolor="<%=Tablecolor1%>">
	  							<font size=<% =fontsize2 %> color="#000000"><b><%=rsRankList("RankNum")%></b></font>
	  						</td>
	  						<td align="Center" width=9% valign="top" bgcolor="<%=Tablecolor1%>">
									<font size=<% =fontsize2 %> color="#000000">
		  							<b><a href="/rankings/<%=ThisFileName%>?MyEvent=<%=sEvent%>&MyDiv=<%=rsRankList("Div")%>&pvar=<%=sRunByWhat%>&sMemberID=<%=sMemberID%>&SkiYear=<%=sSkiYear%>"><%=rsRankList("Div")%>&nbsp;&nbsp;<%=sEventName%></a></b>
									</font>
								</td>
	  						<td align="Center" valign="top" bgcolor="<%=Tablecolor1%>">
	  							<font size=<% =fontsize2 %>  color="#000000"><b><%=sRankScore%></b></font>
	  						</td> 	
							</tr>
							<%
        
        
        			rsRankList.MoveNext	
    			LOOP 
    			
    			
    			%>
		   <br>
  		</TD>
		</TR>
		</TABLE>
		<%


		'response.write("<br><br>Line 1476 Below primary table ")
		'response.end


		%>
		<TABLE class="innertable" align="center" style="border-width:0px;" width="70%">
			<tr>
  			<td colspan=3 align="center" style="border-style:none;">
					<font size=<% =fontsize2 %> face=<% =font3 %> color="#000000"><b>&nbsp;<br>Rankings include USA skiers in Div/Event.<br> Regional or Filtered rankings may be different.</b></font>
					<br><br>
					<font size=<% =fontsize2 %> face=<% =font3 %> color="#000000"><b>Click on Links above to display Rankings for that Div/Event</b></font>
					<br><br>
  			</td>
			</tr>
			<tr>
  			<td width=20% align="right" style="border-style:none;">
    			<font size="<%=fontsize2%>" face=<% =font1 %> color=<%=Titlecolor%>><b>Range:&nbsp;&nbsp;</b></font>
  			</td>

  			<form method=post action="<%=ThisFileName%>?pvar=<%=sRunByWhat%>">
  				<td width=30% align="left" style="border-style:none;">
						<select OnChange=submit() name='SkiYear'>
						<%

							SET rsSelectFields=Server.CreateObject("ADODB.recordset")
							sSQL = "SELECT DISTINCT RT.SkiYearID, SY.SkiYearName"
							sSQL = sSQL + " FROM " &RankTableName&" AS RT"
							sSQL = sSQL + " JOIN " &SkiYearTableName&" AS SY ON RT.SkiYearID = SY.SkiYearID"


							' --- NCWSA does not display 12 Month Rankings
							IF sRunByWhat="NCWSA" THEN
									sSQL = sSQL + " WHERE SY.SkiYearID <> 1"
							END IF
	
							'response.write("<br>Line 1511 - Before Open Statement of Query")
							'response.write("<br>"&sSQL)
							'response.end
							rsSelectFields.open sSQL, SConnectionToTRATable

							' Loads dropdown and sets default to Session("SkiYear")
							DO WHILE NOT rsSelectFields.eof

									IF TRIM(rsSelectFields("SkiYearID")) = TRIM(session("SkiYear")) THEN
											response.write("<option value =""" & rsSelectFields("SkiYearID") &""" selected>")
											response.write(rsSelectFields("SkiYearName"))
											response.write("</option><br>")
									ELSE
											response.write("<option value =""" & rsSelectFields("SkiYearID") &""">")
											response.write(rsSelectFields("SkiYearName"))
											response.write("</option><br>")
									END IF 

									rsSelectFields.moveNEXT
							LOOP

							rsSelectFields.close 
							
						%>
						</select>
   					<input type="hidden" name="SingleRanking" value="Find My Rankings">
   					<input type="hidden" name="sMemberID" value="<%=sMemberID%>">
  				</td>
			</form>

  		<form method=post action="<%=ThisFileName%>?pvar=<%=sRunByWhat%>">
  			<td align="center" style="border-style:none;">
					<input type=submit name="NewMember" value="New Member" title="Select another member and Display their Personal Rankings">
			  </td>
  		</form>
		</tr>
		</TABLE>

  </TD>
	</TR>
</TABLE> 
<% 
'--- Table for border --- 

		'response.write("<br><br>Line 1553 Below border table ")
		'response.end




ELSE  
	
	'response.write("<br><br>Line 1561 ELSE Statement of QuerySuccess=true IF ")
	'response.end

	%>
	<br>
	<TABLE class="innertable" align=center width=60% border=3>
	<TR>
	  <th Colspan=2 align=center>
		<font size=4 color="<%=Textcolor5%>"><b> Notice</b></font>
	  </th>
	</TR>

	<TR>
	<TD>
	<TABLE align=center width=90% border=0>

	<TR>
	  <TD Colspan=2 align=center>
		<font size=<% =fontsize3 %> color="<%=Textcolor3%>"><b> No Records Found For This Member In This Range</b></font>
	  <br><br>
	  </TD>
	</TR>

	<TR>

	  <form method=post action="<%=ThisFileName%>?pvar=<%=sRunByWhat%>&sMemberID=<%=sMemberID%>">
	    <TD align=center>
				<input type=submit name="Continue" style="width:9em" value="Continue"></center>
	    </TD>
	  </form>

	  <form method=post action="<%=ThisFileName%>?pvar=<%=sRunByWhat%>">
	    <TD align=center>
  		<input type=submit name="NewMember" style="width:9em" value="New Member"></center>
	    </TD>
	  </form>
	</TR>
	</TABLE>
	</TD>
	</TR>
	</TABLE>



	<%
END IF

'response.write("<br>Line 1576 Bottom of SUBROUTINE ")
'response.end

END SUB







' ----------------------------------------------------------------------------------
    SUB DisplayRankingLine	' --- Displays a single line of the ranking list ---
' ----------------------------------------------------------------------------------

Dim sbgcolor




IF tLevelNo > 0 AND tLevelNo <> LastLevelNo AND sRunByWhat <> "NSL" AND sRunByWhat <> "NCWSA" THEN 	
		' --- Put a blank row in to separate from heading ---

		IF LastLevelNo="12" THEN
				%><TR>
		  		<td colspan=11>&nbsp;</td>
				</TR><%
		END IF 
		%>
		<TR>
	  	<td style="background-color:<%=DefineLevelcolor%>;">&nbsp;</td>
	  	<TD style="background-color:<%=DefineLevelcolor%>;" align="center">
	  		<font size=<% =fontsize3 %>  color="<%=Textcolor1%>"><b>Level <%=tLevelNo%></b></font>
	  	</TD>
	  	<TD style="background-color:<%=DefineLevelcolor%>;" align="center" colspan=9>&nbsp;</td>
		</TR><%
END IF 


%>
<TR><%	

    ' --- Changes background to red if current member is set and found in this ranking list ---	
    IF rs("MemberID")=sMemberID THEN
				sbgcolor=Textcolor3
    ELSE
				sbgcolor=DefineLevelcolor	
    END IF 


  %>
  <TD style="background-color:<%=sbgcolor%>;" align="Center" valign="top">
	<font size=<% =fontsize2 %> color="<%=Textcolor1%>"><%=RankValueWithTies%></font>
  </TD>
  <TD align="Left" valign="top">&nbsp;<a
  	 href="/rankings/view-scoresHQ.asp?NSL=<%=NSL%>&sMemberID=<%=tMemberID%>&EventSelected=<%=EventSelected%>&pvar=ByMember"
		 title="Click here to Display ALL of&#13;<%=mid(tName,instr(tName,", ")+2)%>'s scores in this Event"><font
		 size=<% =fontsize2 %>><%=tName%></FONT>
	</a>
  </TD>
  <TD align="Center" valign="top"><%

    ' --- Present Ranking Score and Backup Detail
    IF sRunByWhat <> "NSL" THEN  	
  	%><font size=<% =fontsize2 %>  color="<%=Textcolor1%>"><a title="<%=tRnkScoBkup%>"><%=tFmtScore%></a><%
    ELSE
	%><font size=<% =fontsize2 %>  color="<%=Textcolor1%>"><%=tFmtScore%></font><%
    END IF

    ' --- Tack on red asterisk unless Backup includes "NO Penalty"
    IF (sRunByWhat <>"NSL" AND instr(tRnkScoBkup,"Rule 1.13")<>0 AND instr(tRnkScoBkup,"Click Skier")=0) THEN
	%><font color="red"> #</font><%
    ELSEIF (sRunByWhat <>"NSL" AND instr(tRnkScoBkup,"NO Penalty")=0 AND instr(tRnkScoBkup,"Click Skier")=0) THEN
	%><font color="red"> *</font><%
    ELSE
	%>&nbsp;&nbsp;&nbsp; <%
    END IF  %>

  </TD>

  <% IF sRunByWhat <> "NSL" THEN %>
	  <TD align="Center" valign="top"><font size=<% =fontsize2 %> 

	  	<%  IF tEliteStat = "None" THEN %>
				color="Red">&nbsp;
			<% ELSE %>
		  	color="<%=Textcolor1%>">&nbsp;
		  <% END IF %>	  

		  <%  IF len(tEliteStat) > 0 THEN %>	
			  <a title="<%=tEliteBkup%>"><%=tEliteStat%></a>
		  <% ELSE %>
			  <% =tEliteStat %>
		  <% END IF %>	  
	  &nbsp</FONT></TD>
  <%  END IF  %>  

  <TD align="Center" valign="top"><font size=<% =fontsize2 %> color="<%=Textcolor1%>">&nbsp;<% =tState %>&nbsp</FONT></TD>
  
  <% IF sRunByWhat = "NCWSA" THEN %>
	<TD align="Center" valign="top"><font size=<% =fontsize2 %> color="<%=Textcolor1%>">&nbsp;<% =tTeam %> (<% =tTeamStat %>)&nbsp</FONT></TD>
	<TD align="Center" valign="top"><font size=<% =fontsize2 %> color="<%=Textcolor1%>">&nbsp;<% =tNCWRegn %>&nbsp</FONT></TD>
	<TD align="Center" valign="top"><font size=<% =fontsize2 %> color="<%=Textcolor1%>">&nbsp;<% =tNCWConf %>&nbsp</FONT></TD>
  <% ELSE %>
	<TD align="Center" valign="top"><font size=<% =fontsize2 %> color="<%=Textcolor1%>">&nbsp;<% =tRegion %>&nbsp</FONT></TD>
  <% END IF

    IF sRunByWhat <> "NSL" AND sRunByWhat <> "NCWSA" AND EventSelected <> "O" THEN               

			IF tRegPlace <> "" THEN 
				%><TD align="Center" valign="top"><font size=<% =fontsize2 %> color="<%=Textcolor1%>"><%=tRegPlace%></font><%
				IF ucase(tRegSki) <> tRegion THEN
         	%><font color="red" size=<% =fontsize2 %> >&nbsp;[<%=ucase(tRegSki)%>]</font><%
	  		END IF                  
				%></TD><%
			ELSE
				%><TD>&nbsp</TD><%
			END IF 


			IF tNatPlace <> "" THEN
				%><TD align="Center" valign="top"><font size=<% =fontsize2 %> color="<%=Textcolor1%>"><%=tNatPlace%></font></TD><%
			ELSE
				%><TD>&nbsp</TD><%
			END IF 

	  END IF 

   ' -----------  Temporary for displaying LEVELS during testing  MAIN SECTION OF RECORDS --------------	 

    IF sRunByWhat <> "NSL" AND sRunByWhat <> "NCWSA" THEN
	%>
		<TD align="Center" valign="top"><font size=<% =fontsize2 %> color="<%=Textcolor1%>">&nbsp;<% =tMemberFed %></FONT></TD>
		<TD align="Center" valign="top"><font size=<% =fontsize2 %> color="<%=Textcolor1%>">&nbsp;<%=tRankPct%></FONT></TD>
	<%
    END IF 

    IF sRunByWhat <> "NSL" AND sRunByWhat <> "NCWSA" THEN %>	

	  <TD align="Center" valign="top"><% 

'IF Session("Adminmenulevel")=50 THEN 
'	response.write("HomeT="&TRIM(rs("HomoType")))
'	response.write("<br>TReg="&TRIM(tRegion))
'	response.write("<br>RQ="&RIGHT(LEFT(rs("RQTourID"),3),1))
'	response.write("<br>ST="&rs("QfyStatus"))
'	response.write("<br>LS="&LCASE(TRIM(sLeagueSelected)))
'END IF

	
	
	IF TRIM(rs("TourStatus"))<>"X" AND ( TRIM(rs("HomoType"))="A" OR TRIM(tRegion)=RIGHT(LEFT(rs("RQTourID"),3),1) ) THEN  
			RankingQfyTitle="Check details of Qualifications for this Member"
			SELECT CASE TRIM(rs("QfyStatus"))
				CASE "QFY-RPR" 
						RankingQfyTitle="QUALIFIED PENDING REGIONAL PARTICIPATION - Check details of Qualifications for this Member" 
				CASE "Qualified" 
						RankingQfyTitle="QUALIFIED - Check details of Qualifications for this Member" 
			END SELECT
			%>
			<a href="/rankings/MemberQualifications.asp?sMemberID=<%=tMemberID%>&sTourID=<%=TRIM(rs("RQTourID"))%>" title="<%=RankingQfyTitle%>" target="_blank">
	     	<font size=<% =fontsize2 %>><%=rs("QfyStatus")%></FONT>
	    </a>
	    <%
 	ELSEIF TRIM(rs("TourStatus"))<>"X" AND TRIM(tRegion)<>RIGHT(LEFT(rs("RQTourID"),3),1) THEN 
 			%>
			<a href="/rankings/MemberQualifications.asp?sMemberID=<%=tMemberID%>&sTourID=<%=TRIM(rs("RQTourID"))%>" title="<%=RankingQfyTitle%>  This Member's HOME REGION is <%=tRegion%>" target="_blank">
	    	<font size=<% =fontsize2 %>>OOR</FONT>
	    </a>
	    <%
 	ELSEIF TRIM(rs("TourStatus"))<>"X" AND LCASE(TRIM(sLeagueSelected))="none" AND LCASE(TRIM(sLeagueSelected))<>"" THEN %>
			<a href="/rankings/MemberQualifications.asp?sMemberID=<%=tMemberID%>" title="Check details Qualifications for various Tournament/Leagues for this Member." target="_blank">
	      <font size=<% =fontsize2 %>>View</FONT>
	    </a>
	    <%
	ELSE 
		%>
     	<font size=<% =fontsize2 %>>---</FONT>
     	</a>
    <%
	END IF %>
	  </TD><%
	

    END IF  %>

</TR><%



' --- Saves last color for drawing bar across screen at level break, Only if Not Zero ---

IF tLevelNo > 0 THEN LastLevelNo = tLevelNo

END SUB




' -------------------------------------
  SUB DisplayPercentilesandPageFooter
' -------------------------------------


' Writes percentages in text at bottom of list	

SELECT CASE EventSelected
  CASE "S"
    sEventName="Slalom"	
  CASE "T"
    sEventName="Tricks"	
  CASE "J"
    sEventName="Jumping"
  CASE "O"
    sEventName="Overall"
  CASE "WB"
    sEventName="Wakeboard"	
END SELECT

mac=true
IF mac=true AND sRunByWhat = "National" OR sRunByWhat = "Junior" THEN 

		IF Session("adminmenulevel")<10 THEN
		%>
		<br>
		<TABLE class="innertable" align=CENTER width=500 BORDER=0>
			<tr>
				<th colspan=3 align="center"><font size="<%=fontsize1%>" color="<%=Textcolor5%>">&nbsp;&nbsp;Percentiles and COA Scores For&nbsp; <%=DivSelected%>&nbsp; <%=sEventName%></font></th>
			</tr>
			<tr>
				<th align="center"><font size="<%=fontsize1%>" color="<%=Textcolor5%>">&nbsp;&nbsp;Percentiles and COA Scores For&nbsp; <%=DivSelected%>&nbsp; <%=sEventName%></font></th>
			</tr>
    
			<tr>
       	<td align="center" style="background-color:<%=scolor09%>"><font size="<%=fontsize1%>">Level 9</font></td>
       	<td align="center"><font size="<%=fontsize1%>">&nbsp;&nbsp;( 93-100 ) Percentiles</font></td>
       	<td align="center"><font size="<%=fontsize1%>">&nbsp; COA Score:&nbsp; <%=COALevel9%>
	  			<% IF EventSelected = "O" THEN %><font color="red"> **</font><% END IF %>
       		</font>
       	</td>
			</tr>
			<tr>
				<td align="center" style="background-color:<%=scolor08%>"><font size="<%=fontsize1%>">Level 8</font></td>
				<td align="center"><font size="<%=fontsize1%>">&nbsp;&nbsp;( <%=tPerc8%>-92 ) Percentiles </font></td>
				<td align="center"><font size="<%=fontsize1%>">&nbsp; COA Score:&nbsp; <%=COALevel8%></font></td>
			</tr>
	     <tr>
	     	<td align="center" style="background-color:<%=scolor07%>"><font size="<%=fontsize1%>">Level 7</font></td>
	     	<td align="center"><font size="<%=fontsize1%>">&nbsp;&nbsp;( <%=tPerc7%>-<%=(tPerc8)-1%> ) Percentiles</font></td>
	     	<td align="center"><font size="<%=fontsize1%>">&nbsp; COA Score:&nbsp; <%=COALevel7%></font></td>
	    </tr>
      <tr>
      	<td align="center" style="background-color:<%=scolor06%>"><font size="<%=fontsize1%>">Level 6</font></td>
				<td align="center"><font size="<%=fontsize1%>">&nbsp;&nbsp;( <%=tPerc6%>-<%=(tPerc7)-1%> ) Percentiles</font></td>
        <td align="center"><font size="<%=fontsize1%>">&nbsp; COA Score:&nbsp; <%=COALevel6%></font></td>
      </tr>
			<tr>
	    	<td align="center" style="background-color:<%=scolor05%>"><font size="<%=fontsize1%>">Level 5</font></td>
	      <td align="center"><font size="<%=fontsize1%>">&nbsp;&nbsp;( <%=tPerc5%>-<%=(tPerc6)-1%> ) Percentiles</font></td>
	      <td align="center"><font size="<%=fontsize1%>">&nbsp; COA Score:&nbsp; <%=COALevel5%></font></td></tr>
      <tr>
				<td align="center" style="background-color:<%=scolor04%>"><font size="<%=fontsize1%>">Level 4</font></td>
				<td align="center"><font size="<%=fontsize1%>">&nbsp;&nbsp;( <%=tPerc4%>-<%=(tPerc5)-1%> ) Percentiles</font></td>
				<td align="center"><font size="<%=fontsize1%>">&nbsp; COA Score:&nbsp; <%=COALevel4%></font></td>
			</tr>
			<tr>
				<td align="center" style="background-color:<%=scolor03%>"><font size="<%=fontsize1%>">Level 3</font></td>
				<td align="center"><font size="<%=fontsize1%>">&nbsp;&nbsp;( <%=tPerc3%>-<%=(tPerc4)-1%> ) Percentiles</font></td>
				<td align="center"><font size="<%=fontsize1%>">&nbsp; All Others</font></td>
			</tr><%


		' --- Displays the last re-calculation date/time at bottom of screen	

	  sSQL = "SELECT * FROM " & SkiYearTableName & " WHERE "
		IF session("SkiYear") = "0" THEN
				sSQL = sSQL + "DefaultYear = 1"
		ELSE
				sSQL = sSQL + "SkiYearID = " + SQLClean(session("skiyear"))
		END IF

		SET rsSelectFields=Server.CreateObject("ADODB.recordset")
		rsSelectFields.open sSQL, SConnectionToTRATable

		' -- Last Recalc --
		IF not rsSelectFields.eof THEN 
				%>
				<tr>
					<td colspan=3 align="center">
						<small>Rankings last recalculated at <%=rsSelectFields("LastRecalc")%>.</small>
					</td>
				</tr>
				<%
		END IF 

		%>
		<tr>
			<td align="center" colspan=3>
				<font color="red">
					<small>
					<% 
					IF EventSelected = "O" THEN 
							%>** Level 9 uses Overall Scores recalculated to Elite O/A basis;<br><% 
					END IF 
					%>
					* Indicates Penalty;&nbsp;&nbsp;  # Indicates Rule 1.13;&nbsp;&nbsp;  See FAQ/Tips.
					</small>
				</font>
			</td>
		</tr>
		<%

		rsSelectFields.close  
		
		%>
	</TABLE>
	<br><br>
	<%

  ELSE


sSQL="SELECT " 
sSQL=sSQL+" (Select COD FROM "&LeagueTableName&" WHERE RIGHT(LEFT(QualifyTour,3),1)='E' AND HomoType='B' AND RIGHT(LeagueID,4)='2010') AS EADate"
sSQL=sSQL+", (Select COD FROM "&LeagueTableName&" WHERE RIGHT(LEFT(QualifyTour,3),1)='M' AND HomoType='B' AND RIGHT(LeagueID,4)='2010') AS MWDate"
sSQL=sSQL+", (Select COD FROM "&LeagueTableName&" WHERE RIGHT(LEFT(QualifyTour,3),1)='C' AND HomoType='B' AND RIGHT(LeagueID,4)='2010') AS SCDate"
sSQL=sSQL+", (Select COD FROM "&LeagueTableName&" WHERE RIGHT(LEFT(QualifyTour,3),1)='S' AND HomoType='B' AND RIGHT(LeagueID,4)='2010') AS SODate"
sSQL=sSQL+", (Select COD FROM "&LeagueTableName&" WHERE RIGHT(LEFT(QualifyTour,3),1)='W' AND HomoType='B' AND RIGHT(LeagueID,4)='2010') AS WEDate"
sSQL=sSQL+" FROM "&LeagueTableName
sSQL=sSQL+" WHERE RIGHT(LeagueID,4)='2010'"
SET rsReg=Server.CreateObject("ADODB.recordset")
rsReg.open sSQL, SConnectionToTRATable

EADate=(rsReg("EADate"))
MWDate=(rsReg("MWDate"))
SCDate=(rsReg("SCDate"))
SODate=(rsReg("SODate"))
WEDate=(rsReg("WEDate"))


'EADate="07/10/2008"
'SCDate="07/10/2008"
'MWDate="07/10/2008"
'SODate="07/1/2008"
'WEDate="07/1/2008"


	%>
	<br>
	<TABLE class="innertable" align=CENTER width="<%=TourTablewidth%>">
		<tr>
		  <th colspan=3 align="center"><font size="<%=fontsize1%>" color="<%=Textcolor5%>">&nbsp;<%=DivSelected%> - <%=sEventName%></font></th>
		  <th align="center" colspan=2><font size=<% =fontsize2 %> color="<%=Textcolor5%>">&nbsp;Qualification Cut-Off&nbsp;&nbsp;&nbsp;&nbsp;** Does not freeze</font></th>
		</tr>
		<tr>
		  <th width=80px colspan=1 align="center"><font size="<%=fontsize1%>" color="<%=Textcolor5%>">&nbsp;Tournament</font></th>
		  <th colspan=1 align="center"><font size="<%=fontsize1%>" color="<%=Textcolor5%>">&nbsp;Level</font></th>
		  <th colspan=1 align="center"><font size="<%=fontsize1%>" color="<%=Textcolor5%>">&nbsp;Percentiles</font></th>
		  <th colspan=1 align="center"><font size="<%=fontsize1%>" color="<%=Textcolor5%>">&nbsp;COA Score</font></th>
		  <th colspan=1 align="center"><font size="<%=fontsize1%>" color="<%=Textcolor5%>">&nbsp;Date</font></th>
		</tr>

		<tr>
		  <td align=center><font size="<%=fontsize1%>">Elite</font></td>
		  <td align=center style="background-color:<%=scolor09%>"><font size="<%=fontsize1%>">Level 9</font></td>
		  <td align="center"><font size="<%=fontsize1%>">&nbsp;&nbsp;&nbsp;( 93 - 100 ) Percentiles</font></td>
	  	  <td align="center"><font size="<%=fontsize2%>" color="<%=Textcolor1%>">&nbsp;<%=formatnumber(COALevel9,2)%>
	  	  		<% IF EventSelected = "O" THEN %>
				<font color="red"> **</font>
			   <% END IF %>
	  	  </font></td>
		  <td align="center"><font size="<%=fontsize2%>" color="<%=Textcolor1%>">&nbsp; Always 12 Mo&nbsp; </font></td>
		</tr>

		<tr>
		  <td align=center><font size="<%=fontsize1%>">Nationals</font></td>
		  <td align=center style="background-color:<%=scolor08%>"><font size="<%=fontsize1%>">Level 8</font></td>
		  <td align="center"><font size="<%=fontsize1%>">&nbsp;&nbsp;&nbsp;( <%=tPerc8%> - 92 ) Percentiles</font></td>
	  	  <td align="center"><font size="<%=fontsize2%>" color="<%=Textcolor1%>">&nbsp;<%=formatnumber(COALevel8,2)%></font></td>
		  <td align="center"><font size="<%=fontsize2%>" color="<%=Textcolor1%>">&nbsp;<%=CutOffDate%></font></td>
		</tr>

	        <tr>
		  <td align=center ><font size="<%=fontsize1%>">Regionals</font></td>
		  <td align=center style="background-color:<%=scolor07%>"><font size="<%=fontsize1%>">Level 7</font></td>
		  <td align=center><font size="<%=fontsize1%>">&nbsp;&nbsp;&nbsp;( <%=tPerc7%> - <%=(tPerc8)-1%> ) Percentiles</font></td>
	  	  <td align="center"><font size="<%=fontsize2%>" color="<%=Textcolor1%>">&nbsp;<%=formatnumber(COALevel7,2)%></font></td>
		  <td align="center"><font size=<% =fontsize2 %> color="<%=Textcolor1%>">EA - <%=EADate%><br>MW - <%=MWDate%><br>SC - <%=SCDate%><br>SO - <%=SODate%><br>WE - <%=WEDate%></font></td>
		</tr>

        	<tr>
		  <td align="center"><font size="<%=fontsize1%>"><a title="A list of States here as mouse hover-over">States</a></font></td>
		  <td align="center" style="background-color:<%=scolor06%>"><font size="<%=fontsize1%>">Level 6</font></td>
		  <td align=center><font size="<%=fontsize1%>">&nbsp;&nbsp;&nbsp;( <%=tPerc6%> - <%=(tPerc7)-1%> ) Percentiles</font></td>
	  	  <td align="center"><font size="<%=fontsize2%>" color="<%=Textcolor1%>">&nbsp;<%=formatnumber(COALevel6,2)%></font></td>
		  <td> </td>
		</tr>

	        <tr>
		  <td> </td>
		  <td align=center style="background-color:<%=scolor05%>"><font size="<%=fontsize1%>">Level 5</font></td>
		  <td align=center><font size="<%=fontsize1%>">&nbsp;&nbsp;&nbsp;( <%=tPerc5%> - <%=(tPerc6)-1%> ) Percentiles</font></td>
		  <td> </td>
		  <td> </td>
		</tr>

        	<tr>
		  <td> </td>
		  <td align=center style="background-color:<%=scolor04%>"><font size="<%=fontsize1%>">Level 4</font></td>
		  <td align=center><font size="<%=fontsize1%>">&nbsp;&nbsp;&nbsp;( <%=tPerc4%> - <%=(tPerc5)-1%> ) Percentiles</font></td>
		  <td> </td>
		  <td> </td>
		</tr>

	        <tr>
		  <td> </td>
		  <td align=center style="background-color:<%=scolor03%>"><font size="<%=fontsize1%>">Level 3</font></td>
		  <td align=center><font size="<%=fontsize1%>">&nbsp;&nbsp;&nbsp;( <%=tPerc3%> - <%=(tPerc4)-1%> ) Percentiles</font></td>
		  <td> </td>
		  <td> </td>
		</tr><%


		' Displays the last re-calculation date/time at bottom of screen	
		SET rsSelectFields=Server.CreateObject("ADODB.recordset")
	    	sSQL = "SELECT * FROM " & SkiYearTableName & " WHERE "

		IF session("SkiYear") = "0" THEN
        		sSQL = sSQL + "DefaultYear = 1"
		ELSE
		sSQL = sSQL + "SkiYearID = " + SQLClean(session("skiyear"))
		END IF

		rsSelectFields.open sSQL, SConnectionToTRATable

		IF not rsSelectFields.eof THEN 
			%><tr><td colspan=8 align="center">
			<small>Rankings Recalculated <%=rsSelectFields("LastRecalc")%>.</small>
			</td></tr><%
		END IF 

		%><tr><td align="center" colspan=8>
		<font color="red"><small>
	  	  		<% IF EventSelected = "O" THEN %>
				** Level 9 uses Overall Scores recalculated to Elite O/A basis;<br>
			   <% END IF %>
		* Indicates Penalty;&nbsp;&nbsp;  # Indicates Rule 1.13;&nbsp;&nbsp;  See FAQ/Tips.</small></font>
		</td></tr><%

		rsSelectFields.close  %>
	</TABLE><%

   END IF

   ELSE 
		%><br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<font color="red"><small>* Indicates Penalty;&nbsp;&nbsp;  # Indicates Rule 1.13;&nbsp;&nbsp;  See FAQ/Tips.</small></font><%

END IF

END SUB








Sub ChoosePagesSQL(sSQL, sStart, sSize)
	SET rs=Server.CreateObject("ADODB.recordset")

	'WriteDebugSQL(sSQL)

	sqlstmt = sSQL
	rs.CursorType = 3
'  rs.PageSize = cint(sSize)
	rs.open sqlstmt, SConnectionToTRATable
'  IF isrecordsetempty = false THEN
'    rs.AbsolutePage = cINT(sStart)
'  END IF
END SUB



Function IsRecordSetEmpty

IF rs.bof = true AND rs.eof = true THEN
    IsRecordSetEmpty = true
ELSE
    IsRecordSetEmpty = false
END IF
end Function



Sub WriteLink(sParms,sDisplay,sBreak)
%>
<A HREF="<% Response.Write(ThisPage & sParms) %>"><% Response.Write(sDisplay) %></A><% Response.Write(sBreak) %>
<%
END SUB


Sub DoCount(currentPage) 
h = 0

for i = 1 to rs.PageCount
 Response.Write(" <a href=" & chr(34) & ThisPage & "?DivSelected=" & DivSelected & "&RecordNum=" & RecordNum & "&EventSelected=" & EventSelected & "&currentpage=" &  i  & "&action=" & sAction & chr(34) & ">" & i & "</a>")
h = h +1
NEXT
IF h = 0 THEN h = 1
Response.Write("<BR><Small>Page " & currentPage & " of  " & h & "</SMALL></center><BR><BR>")
END SUB

%>




