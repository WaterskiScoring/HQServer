<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->

<%


' -----------------------------------------------------------------------------------------------------------
' ----------------------------------------- MAIN CODE -------------------------------------------------------
' -----------------------------------------------------------------------------------------------------------
'



Dim currentPage, rowCount, i, Tempcolor, sMemberID, sRunByWhat, Grassroots
Dim EventSelected, DivSelected, DivName, ClassSelected
Dim sShowSQL, sSQL, Action
Dim sTourName, sTourID, sTourClass, sTourCity, StateSelected, sTourRegion, sTourSportsGroup, sTourDate
Dim sIncludeScores, sSptsGrpID
Dim TourDisplayWidth, MainImage
NewsPageNum="FAQ_Scores"

TourDisplayWidth=700

ThisFileName="view-scoresHQ_AKA.asp"

DefineTRAStyles



IF TRIM(Session("NewScorVis"))="" THEN
	KickTrafficCounter("NewScorVis")	
	Session("NewScorVis")="YES"
END IF

KickTrafficCounter("NewScorPgs")	






' --------------------------------------------
' --- pvar and sRunByWhat define branching ---
' --------------------------------------------


sSptsGrpID=TRIM(Request("SptsGrpID"))
IF sSptsGrpID<>"" THEN 
	Session("sSptsGrpID")=sSptsGrpID
ELSE
	IF TRIM(Session("sSptsGrpID"))<>"" THEN
		sSptsGrpID=Session("sSptsGrpID")	
	ELSE
		sSptsGrpID="AWS"
	END IF
END IF

sRunByWhat = TRIM(Request("pvar"))

sMemberID = RIGHT(TRIM(request("sMemberID")),9)

EventSelected = left(TRIM(Request("EventSelected")),1)
IF TRIM(EventSelected) = "" THEN EventSelected="S" 

DivSelected = TRIM(Request("DivSelected"))
DivName = TRIM(Request("divname"))

ClassSelected = TRIM(Request("ClassSelected"))
'IF TRIM(ClassSelected)="" THEN ClassSelected="CELR" 

adminmenulevel = Session("adminmenulevel")
IF TRIM(adminmenulevel)="" THEN adminmenulevel=1




' -------------------------------------------------
' ---- Define variables for Tounament Search  -----
' -------------------------------------------------

sShowSQL = Request("sShowSQL")

sTourID = TRIM(Request("Tour_ID")) 
IF len(sTourID) > 6 THEN sTourID = left(sTourID,6)

sTourName = TRIM(Request("Tour_Name"))
sTourCity = TRIM(Request("Tour_City"))
StateSelected = TRIM(Request("StateSelected"))
sTourDate = TRIM(Request("Tour_Date"))
sTourClass = TRIM(Request("Tour_Class"))
sTourRegion = TRIM(Request("Tour_Region"))
sTourSportsGroup = TRIM(Request("sTourSportsGroup"))
sIncludeScores = TRIM(Request("sIncludeScores"))

IF sTourSportsGroup = "" THEN sTourSportsGroup = "AWS"
IF sTourSportsGroup = "NCW" THEN sTourRegion = ""
IF lcase(StateSelected) = "all" THEN StateSelected=""


' ----------------------------------------
' --- Runs Image definition subroutine ---
' ----------------------------------------
WhatDropDownImage EventSelected



' --------------------------------------------------------------------------------------------------------------------
' --- IF the user specified a year, save the year in the session so it applies automatically to all subsequent pages.
' --------------------------------------------------------------------------------------------------------------------

IF TRIM(Request("SkiYear")) <> "" THEN
		Session("SkiYear") = TRIM(Request("SkiYear"))
END IF


' IF the user picked NSL, THEN display the special NSL news page and set
' the session variable so we know they always want NSL reports.
IF Request("NSL") = "0" or Session("NSL") <> "1" THEN
		Session("NSL") = "0"
END IF

IF Request("NSL") = "1" THEN
		Session("NSL") = "1"
		Grassroots = "1"
END IF





' -----------------------------------------------------------------------------------------
' ---- IF we are doing the Print Tours function, THEN go ahead and display the page header.
' -----------------------------------------------------------------------------------------
IF sRunByWhat="PrintTours" THEN

ELSE
		IF sRunByWhat <> "OfficialScores" THEN
				'WriteIndexPageHeader
		END IF
END IF


Action=Request("Action")
IF sRunByWhat="ByTour" THEN Action="ByTour"

'response.write("<br>Line 134 - sRunByWhat="&sRunByWhat)
'response.write("<br>Action="&Action)



' ----------------------------------------------------------------------------
' ------------------  Main Branching SELECT statement here -------------------
' ----------------------------------------------------------------------------

'SELECT CASE sRunByWhat
SELECT CASE LCASE(Action)

  CASE ""  

			' --- This CASE is when no action was selected --- 
			%>
			<br><br>
			<form action="/rankings/<%=ThisFileName%>?pvar=ByTour" method="post">
			<TABLE class="innertable" width=400px align=center>
	  		<tr>
					<th align=center colspan=2 align=center>
						<font face=<%=font1%> size="4" color="<%=Textcolor5%>"><b>Select Search Method</b></font>
       			<br>
	     		</th>
	  		</tr>  
		  	<tr>
	  			<td height=80px align=center style="border-style:none;">
						<input type="submit" style="width:9em" Action="ByTour" value="By Tournament">
	  			</td>
				  <td align=center style="border-style:none;">
		  			<input type="submit" style="width:9em" Action="ByTour" value="By Member">
				  </td>
		  	</tr>
			</TABLE>
	 		</form>
	  	<%

  CASE "officialscores"
  		%>
			<HTML>
				<HEAD>
					<STYLE TYPE="text/css"><!--
      			#bgimg  
      				{
        			background-color: #FFFFFF;
        			background-image: url(/images/logos/USAWatermark80.jpg);
        			background-position: center;
        			background-repeat: no-repeat;
        			background-attachment: fixed;
        			width:100%;
        			height:100%;
        			margin:0px;
      				}
    				--></STYLE>
    			<TITLE>Scores</TITLE>
    		</HEAD>
    		<body>
    			<div id="bgimg">
						<%
						ScoresByMember  
						%>
					</div>
				</body>
		</html><%

    
		CASE "tourscores", "update search"
				WriteIndexPageHeader
   			ScoresByTour
				WriteIndexPageFooter
  		
		CASE "printtours"
   			DisplayTourList	

		CASE "print results"
  			%>
				<form method=post action="/rankings/<%=ThisFileName%>?action=ByTour">
  			<TABLE align=center width=30%>
  				<tr>
  					<td align=center>
  						<a href='#' onclick='window.print()' title="Click here to Print">
  							<input type=button style="width:9em;" value="Print Screen">
  						</a>
  					</td>
		  			<td align=center>
		  					<input type=submit style="width:9em;" value="Tour List" title="Return to Tournament List">
  					</td>
					</tr>
				</TABLE>
			</form>
		  	<%
   			DisplayTourList	


		CASE "bytour"
		'response.write("<br>Line216")
				WriteIndexPageHeader
				DisplayTourSearchFilters
				DisplayTourList	
				WriteIndexPageFooter

		CASE "begin search", "tour list", "new tournament"
		'response.write("<br>Line223")
				WriteIndexPageHeader
				DisplayTourSearchFilters
				DisplayTourList	
				WriteIndexPageFooter

		CASE "faq/tips"
				DisplayTourList	


		CASE "bymember", "by member"
      ' -----  User selected link to View Scores by Member  ----
				IF TRIM(sMemberID)="" THEN
						' --- Sends user to search-member routine to selected member
						Session("sSendingPage")="/rankings/"&ThisFileName&"?pvar=FoundMember"
						Response.Redirect("/rankings/search-memberHQ.asp?formstatus=search")
				ELSE
						ScoresByMember
						'	Response.Redirect("/rankings/"&ThisFileName&"?sMemberID="&sMemberID&"&pvar=FoundMember&EventSelected="&EventSelected)			
				END IF

   CASE "foundmember"   ' --- User successfully selected a member
				ScoresByMember
	

END SELECT



' IF we are doing the Print Tours function, THEN go ahead and display
' the page footer.

IF sRunByWhat="PrintTours" THEN
		IF sRunByWhat <> "OfficialScores" THEN
				'WriteIndexPageFooter
		END IF
END IF



' ---------------------------------------------------
' ----------  Writes the footer for the HQ site  ----
' ---------------------------------------------------







' -----------------------------------------------------------------------------------------
' ---------------------------   END OF MAIN CODE   ----------------------------------------
' -----------------------------------------------------------------------------------------





' ----------------------
    SUB DisplayTourList
' ----------------------

'response.write("sTourSportsGroup = "&sTourSportsGroup)

' ---------------------------------------------------------------------------------
' -------------------------  DISPLAY TOURNAMENTS  ---------------------------------
' ---------------------------------------------------------------------------------

	Dim DateGood
	DateGood = 0

	OpenCon
	set rs=Server.CreateObject("ADODB.recordset")
	
	' --- Validates sTourDate (Start) ---
	IF (isnumeric(LEFT(sTourDate,2)) AND isnumeric(RIGHT(LEFT(sTourDate,5),2)) AND isnumeric(RIGHT(sTourDate,4)) AND RIGHT(left(sTourDate,3),1) = "/" AND RIGHT(left(sTourDate,6),1) = "/" And isDate(sTourDate)) Or (sTourDate = "") THEN
			DateGood = 1
	ELSE
			DateGood = 0
	END IF


	' --------------------------------------
	' --- User has not selected anything ---
	' --------------------------------------
	IF (TRIM(request("SkiYear")) = "" AND sTourID = "" AND sTourName = "" AND sTourCity = "" AND StateSelected = "" AND sTourClass = "" AND sTourRegion = "" AND sTourDate = "") OR (DateGood = 0) THEN

	ELSE
			CreateTourListQuery
	END IF

END SUB



' ----------------------------
  SUB CreateTourListQuery
' ----------------------------

			' ---------------------------------------------------------------------------------
			' --- Creates array consisting of a unique list of TourID's from RawScoresTable ---
			' ---------------------------------------------------------------------------------
			IF sTourSportsGroup="AWS" OR sTourSportsGroup="NCW" THEN
					sSQL = "SELECT DISTINCT left(TourID,6) FROM "&RawScoresTableName&" RS"
			ELSE
					sSQL = "SELECT DISTINCT left(TourID,6) FROM "&RawScoresOtherTableName&" RS"
			END IF
			sSQL = sSQL + ", "&SanctionTableName&" ST"
			sSQL = sSQL + " WHERE LEFT(ST.TournAppID,6)=LEFT(RS.TourID,6)"
			'1=1"

			' -------------------------------------------------------------------------
			' --- Add Condition for selecting SptsGrp for non-AWS/NCW Sports Groups ---
			' -------------------------------------------------------------------------
			SELECT CASE sTourSportsGroup
				CASE "AWS", "NCW" 
						' Do nothing
						' sSQL = sSQL + " AND SptsGrpID='AWS'"			
				CASE "NCW"
						' Do nothing
						' sSQL = sSQL + " AND SptsGrpID='NCW'"
				CASE "USW"
						sSQL = sSQL + " AND RS.SptsGrpID='USW'"
				CASE "AKA"
						sSQL = sSQL + " AND RS.SptsGrpID='AKA'"
			END SELECT
  	
		IF Session("SkiYear") = 0 THEN
				' --- IF 0 THEN do the rolling 12 month calc.
				' --- IF the default year is not a valid id (basically, big problems!), THEN just poison the search with 1=0.
				' --- IF the default info is found, THEN use the begin and end dates to filter the query.
				set rsSelectFields=Server.CreateObject("ADODB.recordset")
				rsSelectFields.open ("SELECT TOP 1 * FROM " & SkiYearTableName & " WHERE SkiYearID = 1"), SConnectionToTRATable, 3, 3  
				IF rsSelectFields.eof THEN
						sSQL = sSQL + " and 1 = 0"
				ELSE
						sSQL = sSQL + " AND (ST.TDateE <= '" & FormatDateTime(rsSelectFields("EndDate"),2) & "' AND ST.TDateE >= '" & FormatDateTime(rsSelectFields("BeginDate"),2) & "')"
				END IF
				rsSelectFields.close
		ELSE

				' IF not 0, THEN do whatever ski year is indicated by the id provided.
				' IF the year provided is not a valid id, THEN just poison the search with 1=0.
				' IF the year is found, THEN use the begin and end dates to filter the query.
				set rsSelectFields=Server.CreateObject("ADODB.recordset")
				rsSelectFields.open ("Select top 1 * from " & SkiYearTableName & " where SkiYearID = " & sqlclean(Session("SkiYear"))), SConnectionToTRATable, 3, 3  
				IF rsSelectFields.eof THEN
						sSQL = sSQL + " and 1 = 0"
				ELSE
						sSQL = sSQL + " and (ST.TDateE <= '" & FormatDateTime(rsSelectFields("EndDate"),2) & "' and ST.TDateE >= '" & FormatDateTime(rsSelectFields("BeginDate"),2) & "')"
				END IF
				rsSelectFields.close
		END IF


		sSQL = sSQL + " GROUP BY left(TourID,6)"
		sSQL = sSQL + " ORDER BY left(TourID,6) DESC"

		set rs=Server.CreateObject("ADODB.recordset")
		rs.open sSQL, sConnectionToTRATable, 3, 1

		RSArray = RS.Getrows()
		rs.close

		ScoredTours = "("
		FOR j = 0 TO ubound(RSArray,2)
			ScoredTours = ScoredTours + "'" + TRIM(RSArray(0,j)) + "'"
			IF j < ubound(RSArray,2) THEN
					ScoredTours = ScoredTours + ","
			END IF
		NEXT
		ScoredTours = ScoredTours + ")"


		' ***********************************
		' *** Create the TOURNAMENT query ***
    ' ***********************************
    
		' --------------------------------------------------
		' ---- Creates the query for the tournament list ---
		' --------------------------------------------------
		sSQL = "SELECT TSanction, TDateS, TDateE, TName, TCity, TState, TStatus FROM "&SanctionTableName&" AS ST"
		sSQL = sSQL + " WHERE 1 <> 2"

	  ' --------------------------------------------------------
	  ' --- Selects specific tournaments based on SptgsGrpID ---
	  ' --------------------------------------------------------
	  IF sTourSportsGroup = "NSL" THEN
				sSQL = sSQL + " AND (TEventFun <> 0 or TEventFHF <> 0 or TEventFKB <> 0 or TEventFDA <> 0 or TEventF3ev <> 0 or TEventFB <> 0 or TEventFW <> 0)"
		ELSE
				sSQL = sSQL + " AND lower(ST.SptsGrpID) = '" & sqlclean(lCASE(sTourSportsGroup)) & "'"
		END IF

		' --------------------------------------
		' --- Filter for TSanction = sTourID ---
		' --------------------------------------
		IF sTourID <> "" THEN
			sSQL = sSQL + " AND lower(left(ST.TSanction," & len(sTourID) & ")) = '" & sqlclean(lCASE(sTourID)) & "'"
		END IF

		' -----------------------
		' --- Tournament Name ---
		' -----------------------
		IF sTourName <> "" THEN
			sSQL = sSQL + " AND lower(ST.TName) LIKE '%" & sqlclean(lCASE(sTourName)) & "%'"
		END IF

    ' -------------
    ' --- CLASS ---
    ' -------------
		IF sTourClass <> "" THEN
		  IF sTourClass = "LR" THEN
					sSQL = sSQL + " AND lower(right(ST.TSanction,1)) in ('l','r')"
		  ELSEIF sTourClass = "ELR" THEN
					sSQL = sSQL + " AND lower(right(ST.TSanction,1)) in ('e','l','r')"
	  	ELSEIF sTourClass = "CELR" THEN
					sSQL = sSQL + " AND lower(right(ST.TSanction,1)) in ('c','e','l','r')"
		  ELSEIF sTourClass = "F" or Session("NSL") = "1" THEN
		  		sSQL = sSQL + " AND lower(right(ST.TSanction,1)) in ('f','n','i')"
	  	ELSEIF sTourClass = "T" THEN
					sSQL = sSQL + " AND lower(right(ST.TSanction,1)) in ('t')"
	  	ELSEIF sTourClass = "Q" THEN
					sSQL = sSQL + " AND lower(right(ST.TSanction,1)) in ('q')"
	  	ELSEIF sTourClass = "W" THEN
					sSQL = sSQL + " AND lower(right(ST.TSanction,1)) in ('w')"
	  	END IF
		END IF

		' -----------------------
		' --- Tournament CITY ---
		' -----------------------
		IF sTourCity <> "" THEN
	  		sSQL = sSQL + " AND lower(ST.TCity) LIKE '%" & sqlclean(lCASE(sTourCity)) & "%'"
		END IF

		' ------------------------
		' --- Tournament STATE ---
		' ------------------------
		IF StateSelected <> "" THEN
		  	sSQL = sSQL + " AND lower(ST.TState) LIKE '%" & sqlclean(lCASE(StateSelected)) & "%'"
		END IF

		' -------------------------
		' --- Tournament REGION ---
		' -------------------------
		IF sTourRegion <> "" THEN
		  	sSQL = sSQL + " AND lower(right(left(ST.TSanction,3),1)) = '" & sqlclean(lCASE(sTourRegion)) & "'"
		END IF

		' -----------------------
		' --- Tournament DATE ---
		' -----------------------
		IF sTourDate <> "" THEN
			sSQL = sSQL + " AND (ST.TDateE >= '" & sTourDate & "' AND ST.TDateS <= '" & sTourDate & "')"
		END IF


		IF sIncludeScores = "with" THEN
	  		sSQL = sSQL + " and left(ST.TSanction,6) in " + ScoredTours
		ELSEIF sIncludeScores = "without" THEN
		    	sSQL = sSQL + " and left(ST.TournAppID,6) not in " + ScoredTours
		END IF


		IF Session("SkiYear") = 0 THEN
				' --- IF 0 THEN do the rolling 12 month calc.
				' --- IF the default year is not a valid id (basically, big problems!), THEN just poison the search with 1=0.
				' --- IF the default info is found, THEN use the begin and end dates to filter the query.

				set rsSelectFields=Server.CreateObject("ADODB.recordset")
				rsSelectFields.open ("Select top 1 * from " & SkiYearTableName & " where SkiYearID = 1"), SConnectionToTRATable, 3, 3  
				IF rsSelectFields.eof THEN
						sSQL = sSQL + " and 1 = 0"
				ELSE
						sSQL = sSQL + " and (ST.TDateE <= '" & FormatDateTime(rsSelectFields("EndDate"),2) & "' and ST.TDateE >= '" & FormatDateTime(rsSelectFields("BeginDate"),2) & "')"
				END IF
				rsSelectFields.close
		ELSE

				' IF not 0, THEN do whatever ski year is indicated by the id provided.
				' IF the year provided is not a valid id, THEN just poison the search with 1=0.
				' IF the year is found, THEN use the begin and end dates to filter the query.

				set rsSelectFields=Server.CreateObject("ADODB.recordset")
				rsSelectFields.open ("Select top 1 * from " & SkiYearTableName & " where SkiYearID = " & sqlclean(Session("SkiYear"))), SConnectionToTRATable, 3, 3  
				IF rsSelectFields.eof THEN
						sSQL = sSQL + " and 1 = 0"
				ELSE
						sSQL = sSQL + " and (ST.TDateE <= '" & FormatDateTime(rsSelectFields("EndDate"),2) & "' and ST.TDateE >= '" & FormatDateTime(rsSelectFields("BeginDate"),2) & "')"
				END IF
				rsSelectFields.close
		END IF


		' -------------------------
		' --- Tournament STATUS ---
		' -------------------------

		IF sTourSportsGroup="AWS" THEN
			sSQL = sSQL + " AND ST.TStatus IN (2,4,5)"
		END IF

		' ----------------
		' --- ORDER BY ---
		' ----------------

		sSQL = sSQL + " ORDER BY ST.TDateS " 
		rs.open sSQL, sConnectionToSanctionTable, 3, 1


	' ----------------------------------------------------------------
	' --- No records found in SWIFT tables meeting search criteria --- 
	' ----------------------------------------------------------------
	IF rs.EOF THEN   
	
			%><br>
				<h4>
				<center>
					<font size=3 face=<%=font2%> color="red"><I>No Records Found - Please Change The Filtering Parameters </I></font>
				</center>
				</h4>
				<%

	' ---------------------------------------------------------------------------------------		
	' --- Found MULTIPLE records in swift for search condition so display tournament list ---
	' ---------------------------------------------------------------------------------------	
	ELSE	
			
			'IF rs.recordcount > 1 THEN 
			' --- Formerly did something different if only one record found - too complicated ---
			'ad=1
			'IF ad=1 THEN
					DisplayTourListResults
			'ELSE
				'	sTourID = left(rs("TSanction"),6)
       	'	EventSelected = TRIM(Request("EventSelected"))
       	'	DivSelected = TRIM(Request("DivSelected"))
	      '	RankNum = TRIM(Request("ranknum"))    
       		'rs.close
        
	     ' 	IF RankNum = "" THEN RankNum = 1

			'		IF EventSelected = "" THEN
				'			sSQL = "Select top 1 [event] from " & RawScoresTableName & " where lower(left(TourID,6)) = '" & sqlclean(lCASE(sTourID)) & "' order by div, event, score DESC, altscore DESC"
				'			rsSelectFields.open sSQL, SConnectionToTRATable

				'			IF NOT rsSelectFields.EOF THEN 
				'					EventSelected = RsSelectFields.Fields(0).Value  
				'			END IF
				'			rsSelectFields.Close
				'	END IF

				'	IF DivSelected = "" THEN
				'			sSQL = "Select top 1 div from " & RawScoresTableName & " where lower(left(TourID,6)) = '" & sqlclean(lCASE(sTourID)) & "' order by div, event, score DESC, altscore DESC"
				'			rsSelectFields.open sSQL, SConnectionToTRATable

				'			IF NOT rsSelectFields.EOF THEN 
				'					DivSelected = RsSelectFields.Fields(0).Value  
       	'			END IF
        '			rsSelectFields.Close

			  '			currentPage = TRIM(Request("currentPage"))
       	'			IF currentPage = "" THEN currentPage = 1
        
				'			sID = TRIM(Request("id"))
       	'			IF sID = "" THEN sID = 0
        '			ThisPage = Request.ServerVariables("SCRIPT_NAME")

							' -----------------------------------
							' --- Display ScoresByTour format ---
							' -----------------------------------
							
				'	END IF



			'END IF	

		'ScoresByTour

	END IF

  ' --- Close conection ---
	CloseCon
	set rs = nothing
	set rsSelectFields = nothing


END SUB
	
   
' ---------------------------  
  SUB DisplayTourListResults  
' ---------------------------    

	IF Session("NSL")="1" THEN 
			TourType="Grassroots" 
	ELSE 
			TourType="AWSA & Collegiate"
	END IF	

	' ---------------------------------------------------------------------------------------------
	' Write Headers for DB Page
	' --------------------------------------------------------------------------------------------- 
			 
	%>
	<TABLE class="innertable" WIDTH=98% align=center><%

	IF adminmenulevel>=19 THEN 
			TotalColWidth=6 
	ELSE 
			TotalColWidth=5 
	END IF

	' ------------------------------
	' --- Display table headings ---
	' ------------------------------
	%>
	<tr>
		<th ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor5%>"><b>Tour ID</b></FONT></th>
		<th ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor5%>"><b>Tour Name</b></FONT></th>
		<th ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor5%>"><b>City</b></FONT></th>
		<th ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor5%>"><b>State</b></FONT></th>
		<th ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor5%>"><b>Start Date</b></FONT></th><%
			
		IF Adminmenulevel>=19 THEN 
		 		%><th ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor5%>">TStatus</FONT></th><%
		 END IF %>
	 </tr><% 

		' ----------------------------------------------
		' --- Display each tournament record summary ---
		' ----------------------------------------------
		DO WHILE Not rs.EOF 
						%>
						<tr>
							<TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>"><%=rs("TSanction")%></FONT></TD>
							<TD ALIGN="center" vAlign="top" bgcolor="#C0C0C0"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>"><a href="/rankings/<%=ThisFileName%>?tour_id=<%=rs("TSanction")%>&sTourSportsGroup=<%=sTourSportsGroup%>&action=tourscores"><%=rs("TName")%></a></FONT></TD>
							<TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>"><%=rs("TCity")%></FONT></TD>
							<TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>"><%=rs("TState")%></FONT></TD>
							<TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>"><%=rs("TDateS")%></FONT></TD>
							<%
				      IF Adminmenulevel>=19 THEN 
									%><TD ALIGN="Center" vAlign="top" bgcolor="#C0C0C0"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>"><a href="/rankings/<%=ThisFileName%>?tour_id=<%=rs("TSanction")%>&sTourSportsGroup=<%=sTourSportsGroup%>&pvar=ByTour"><%=rs("TStatus")%></a></FONT></TD><%
				      END IF 
				      %>
				    </tr>
				    <% 
				    rs.MoveNext
		LOOP

		rs.Close 
				
		%>
	  </TABLE>
		<br><br><br>
		<% 
		
		
		' -----------------------------------------------------------
		' --- Displays buttons for continuing and Printing Screen --- 
		' -----------------------------------------------------------

			IF Action="Print Results" THEN 
					%>
					<form method=post action="/rankings/<%=ThisFileName%>">
			  	<TABLE align=center>
			  		<tr>	
			    		<td width=25% Align="left">
							  <input type="hidden" name="pvar" value="ByTour">
							  <input type="hidden" name="sTourSportsGroup" value="<%=sTourSportsGroup%>">
				  			<input type="hidden" name="sIncludeScores" value="<%=sIncludeScores%>">
				  			<input type=submit style="width:9em" Action="ByTour" value="Continue">
			    		</td>
							<td width=25% Align="left">
				  			<input type="hidden" name="sTourSportsGroup" value="<%=sTourSportsGroup%>">
				  			<input type="hidden" name="sIncludeScores" value="<%=sIncludeScores%>">
		 		  			<a href='#' onclick='window.print()' title="Click here to Print">
									<input type=button value="Print Screen" style="width:9em">
				  			</a>
			    		</td>
					  </tr>
			  	</TABLE>
					</form>
					<%
			END IF


END SUB



' --------------------------
    SUB ScoresByTour
' --------------------------

'response.write("<br>Line742 - ScoresByTour")



sSQL = "Select TOP 1 TSanction, TDateS, TDateE, TName, TCity, TState from "& SanctionTableName
IF sSPtsGrpID="AWS" OR sSptsGrpID="NCW" THEN
		sSQL = sSQL + " WHERE lower(left(TSanction,6)) = '" & sqlclean(lCASE(sTourID)) & "' and TStatus in (2,4,5)"
ELSE
		sSQL = sSQL + " WHERE lower(left(TSanction,6)) = '" & sqlclean(lCASE(sTourID)) &"'"
END IF
set rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, sConnectionToTRATable, 3, 1


IF not rs.eof THEN
		sTName=rs("TName")
		sTSanction=rs("TSanction")
		sTCity=rs("TCity")
		sTState=rs("TState")
		sTDateS=rs("TDateS")
		sTDateE=rs("TDateE")

		IF Session("NSL")="1" AND sTourSportsGroup="AWS" THEN 
				TourType="Grassroots" 
		ELSEIF sTourSportsGroup="AWS" THEN
				TourType="AWSA"
		ELSEIF sTourSportsGroup="NCW" THEN
				TourType="Collegiate"
		ELSEIF sTourSportsGroup="USW" THEN
				TourType="Wakeboard"
		ELSEIF sTourSportsGroup="AKA" THEN
			TourType="Kneeboard"
		END IF	   
ELSE
		session("message") = "Tourname " & sTourID & " was not found."
		Response.Redirect("/?process=logout")
END IF

rs.close 



' -----------  Query to collect scores for this Tournament --------------------

IF sTourSportsGroup="AWS" OR sTourSportsGroup="NCW" THEN
		sSQL = "SELECT DISTINCT RAW.*, MEM.*, DT.*  FROM " & RawScoresTableName&" AS RAW"
		sSQL = sSQL + ", " & DivisionsTableName & " AS DT, " &SkiYearTableName& " AS SY, "&MemberTableName&" AS MEM"
ELSE
		sSQL = "SELECT DISTINCT RAW.*, MEM.*, DT.*  FROM " & RawScoresOtherTableName&" AS RAW"
		sSQL = sSQL + ", " & DivisionsOtherTableName & " AS DT, " &SkiYearTableName& " AS SY, "&MemberTableName&" AS MEM"
END IF

sSQL = sSQL + " WHERE lower(left(TourID,6)) = '" & sqlclean(lCASE(sTourID)) & "'"
IF TRIM(DivSelected)<>"" THEN
		sSQL = sSQL + " AND RAW.div = '" & sqlclean(DivSelected) & "'"
END IF
IF TRIM(EventSelected)<>"" THEN
		sSQL = sSQL + " AND [event] = '" & sqlclean(EventSelected) & "'"
END IF
sSQL = sSQL + " AND RAW.Div = DT.Div AND SY.skiyearid = DT.skiyearid AND SY.skiyearid<>'1'"

IF sTourSportsGroup<>"AWS" AND sTourSportsGroup<>"NCW" THEN
		sSQL = sSQL + " AND DT.SptsGrpID='"&sTourSportsGroup&"'"
END IF

sSQL = sSQL + " AND RAW.EndDate between SY.BeginDate AND SY.EndDate"
sSQL = sSQL + " AND RAW.MemberID = MEM.PersonIDwithCheckDigit"
sSQL = sSQL + " ORDER BY event, score DESC, PLACE, altscore DESC"

'response.write(sSQL)
'response.end

set rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, sConnectionToTRATable, 3, 1

%>
	<form method=post action="/rankings/<%=ThisFileName%>">   
		<input type="hidden" name="Tour_ID" value="<%=sTourID%>">
		<input type="hidden" name="sTourSportsGroup" value="<%=sTourSportsGroup%>">
 		<input type="hidden" name="sIncludeScores" value="<%=sIncludeScores%>">
	<TABLE class="droptable" height="225px" WIDTH="<%=TourDisplayWidth%>" background="<%=MainImage%>">
	  <tr>
	    <td colspan=8 align="left" style="vertical-align:top;">
				<font size="<%=fontsize4%>" face="<%=font2%>" color="<%=Textcolor2%>"><b><%=sTName%></b></font>
				<br>
				<font size="<%=fontsize4%>" face="<%=font2%>" color="<%=Textcolor1%>"><b>TourID: <%=sTSanction%></b></font>
				<br>
				<font size="<%=fontsize3%>" face="<%=font2%>" color="<%=Textcolor1%>"><%=sTCity%>, <%=sTState%></font>
				<br>
				<font size="<%=fontsize3%>" face="<%=font2%>" color="<%=Textcolor1%>"><%=sTDateS%> to <%=sTDateE%></font>
	    </td>
	  </tr>
  	<tr>
    	<td Align="right" width="100px" height="20px">
		  	<font size="<%=fontsize2%>" face="<%=font2%>" color="<%=Textcolor2%>"><b>Event:&nbsp;</b></font>
    	</td>
    	<td align="left" colspan=3>
				<%

				' ----------------------------
				' ---  Load EVENT dropdown ---
				' ----------------------------
				LoadEventDropForScores

				%>				
  		</td>
  		<td colspan=4 width="50%">&nbsp;</td> 
  	</tr>
    <tr>
    	<td align="right" height="20px">
				<font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor2%>"><b>Div:&nbsp;</b></font>
    	</td>
	    <td align="left" colspan=3>
				<%

				' -------------------------------
				' --- Loads DIVISION Dropdown ---
				' -------------------------------				 
				LoadDivisionDropForScores

				%>
    	</td>
	  	<td align="center" colspan=2 width="25%">
	  		<input type="submit" style="width:9em" name="Action" value="Update Search">
	  	</td>
    	<td align="center" colspan=2 width="25%">
	  		<input type="submit" style="width:9em" name="Action" value="New Tournament">
    	</td>
  </TR>
</TABLE>
</form>
<%


' ------------------------------------------------
' ----------  BEGIN Display of Scores  -----------
' ------------------------------------------------

IF rs.eof THEN  
		%>
		<br>
    <center><font color="red">No Scores Found In This Event and Division</font></center>
    <% 
ELSE 
		%>
		<TABLE class="innertable" WIDTH="<%=TourDisplayWidth%>px">
			<tr>
	  		<th ALIGN="Center"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor5%>">Member Name</FONT></th>
	  		<% 
	
				IF Session("NSL") <> "1" THEN 
						%>
      			<th ALIGN="Center"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor5%>">Score</FONT></th><%

						IF DivSelected = "CM" OR DivSelected = "CW" THEN 
								%><th ALIGN="Center"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor5%>">Team</FONT></th><%
						END IF	
						%>
						<th ALIGN="Center"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor5%>">Round</FONT></th>
    				<th ALIGN="Center"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor5%>">Div</FONT></th>
    				<th ALIGN="Center"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor5%>">Place</FONT></th>
    				<%
        		IF left(EventSelected,1) = "S" THEN 
        				%>
								<th ALIGN="Center"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor5%>">Buoys</FONT></th>
								<th ALIGN="Center"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor5%>">Line</FONT></th>
								<%
						END IF

	      		IF left(EventSelected,1) = "J" THEN 
	      				%><th ALIGN="Center"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor5%>">Ramp</FONT></th><%
						END IF

						IF left(EventSelected,1) <> "T" THEN 
								%><th ALIGN="Center"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor5%>">Speed</FONT></th><%
						END IF 
						%>
        		<th ALIGN="Center"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor5%>">Class</FONT></th>
        		<% 
		
				ELSE 
						%>
        		<th ALIGN="Center"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor5%>">Placement Points</FONT></th>
	      		<th ALIGN="Center"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor5%>">Div</FONT></th>
        		<th ALIGN="Center"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor5%>">Place</FONT></th>
	      		<th ALIGN="Center"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor5%>">Round</FONT></th>
        		<th ALIGN="Center"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor5%>">Score</FONT></th>
        		<% 
				END IF 
				%>
		</TR>
		<%

 

	' ------------------------------------------------
	' ---  Beginning of LOOP for displaying scores ---
	' ------------------------------------------------
	DO WHILE not rs.eof  
			%>
			<TR>
				<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<a href="/rankings/<%=ThisFileName%>?pvar=ByMember&sMemberID=<%=rs("MemberID")%>&DivSelected=<%=DivSelected%>&EventSelected=<%=EventSelected%>"><% Response.Write(rs("LastName") & ", " & rs("FirstName")) %></a></FONT></TD>
				<% 
				IF Session("NSL") <> "1" THEN 
						IF rs("Score") <> "" THEN
								IF EventSelected = "T" THEN 
                		%><TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<% Response.Write formatnumber(rs("score"),0) %></FONT></TD><% 
								ELSE 
										%><TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<% IF (Request("ZBSAdjustToOldStyle") = "on") And (rs("Event") = "S") THEN Response.Write formatnumber((rs("score") - rs("ZBSConversion")),2) ELSE Response.Write formatnumber(rs("score"),2) %></FONT></TD><% 
								END IF 
						ELSE 
								%><TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;</FONT></TD><% 
						END IF 


						' --- Display Round, link on Div, Place ---
						%>
    				<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<% =rs("round") %></FONT></TD>
						<TD Align="center" valign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<a href="/rankings/view-standingsHQ.asp?pvar=National&DivSelected=<%=rs("div")%>&EventSelected=<%=left(EventSelected,1)%>"><% =rs("div") %></a></font></td>
						<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<% =rs("place") %></FONT></TD>
						<%

						IF left(EventSelected,1) = "S" THEN 
								%><TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<%=rs("AltScore")%></FONT></TD><%
						END IF


						' --- Display Rope, Boat Speed, Ramp Height ----
						IF rs("Perf_Qual1") <> "" THEN
	              IF left(EventSelected,1) = "S" THEN 
	              		%><TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<%=formatnumber(rs("Perf_Qual1")/100, 2)%></FONT></TD><%
								END IF
						ELSE
								IF left(EventSelected,1) = "S" THEN 
										%><TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;</FONT></TD><%
								END IF
						END IF

						IF left(EventSelected,1) = "J" THEN 
								%><TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<%=rs("Perf_Qual1")%></FONT></TD><%
						END IF
	
						IF left(EventSelected,1) <> "T" THEN 
								%><TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<%=rs("Perf_Qual2")%></FONT></TD><%
						END IF  
						%>
						<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<%=rs("class") %></FONT></TD><% 
	    	ELSE    ' --- Display NSL Data ---  
	    			%>
     				<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<% =rs("NSL_Placement_Points") %></FONT></TD>
						<TD Align="center" valign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<a href="/rankings/view-standingsHQ.asp?pvar=National&DivSelected=<%=rs("div")%>&EventSelected=<%=left(EventSelected,1)%>"><% =rs("div") %></a></font></td>
						<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<% =rs("place") %></FONT></TD>
						<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<% =rs("round") %></FONT></TD>
						<%

						IF rs("Score") <> "" THEN
								IF EventSelected = "T" THEN 
										%><TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<% Response.Write formatnumber(rs("score"),0) %></FONT></TD><% 
								ELSE 
										%><TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<% IF (Request("ZBSAdjustToOldStyle") = "on") And (rs("Event") = "S") THEN Response.Write formatnumber((rs("score") - rs("ZBSConversion")),2) ELSE Response.Write formatnumber(rs("score"),2) %></FONT></TD><%
								END IF 
						ELSE 
								%><TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;</FONT></TD><% 
						END IF

        END IF 
        
        %>
		</TR><% 

	rs.movenext
  LOOP 
  
  %>
  </TABLE>
  <br><br>
  </form>
  <%


END IF

END SUB





' -----------------------------
  SUB LoadEventDropForScores
' -----------------------------

			%>
			<select name='EventSelected'style="Width:12em">
				<%

				IF sTourSportsGroup="AWS" OR sTourSportsGroup="NCW" THEN
						sSQL = "Select distinct [event] from " & RawScoresTableName & " where lower(left(TourID,6)) = '" & sqlclean(lCASE(sTourID)) & "' order by [event]"
				ELSE
						sSQL = "Select distinct [event] from " & RawScoresOtherTableName & " where lower(left(TourID,6)) = '" & sqlclean(lCASE(sTourID)) & "' order by [event]"
				END IF

				set rsSelectFields=Server.CreateObject("ADODB.recordset")
				rsSelectFields.open sSQL, SConnectionToTRATable

				IF not rsSelectFields.eof THEN 
						rsSelectFields.movefirst

						DO WHILE not rsSelectFields.eof
								IF TRIM(rsSelectFields.Fields(0).value) = EventSelected THEN
										response.write("<option value =""" & rsSelectFields.Fields(0).value &""" selected>")
										SELECT CASE rsSelectFields.Fields(0).value
												CASE "T"
														EventName = "Trick"
												CASE "S"
														EventName = "Slalom"
												CASE "J"
														EventName = "Jump"
												END SELECT

												response.write(EventName)
												response.write("</option><br>")
    						ELSE
      							response.write("<option value =""" & rsSelectFields.Fields(0).value &""">")

										SELECT CASE rsSelectFields.Fields(0).value
												CASE "T"
														response.write("Trick")
												CASE "S"
														response.write("Slalom")
												CASE "J"
        										response.write("Jump")
										END SELECT
										response.write("</option><br>")
								END IF

								rsSelectFields.movenext
						LOOP
				ELSE
						response.write("<option value =""None"" selected>None</option>")
				END IF

				rsSelectFields.close  
				%>
				</select>
<%
END SUB


' --------------------------------
  SUB LoadDivisionDropForScores
' --------------------------------  
  %>
  				<select name='DivSelected' style="Width:12em">
					<%
					IF sTourSportsGroup="AWS" OR sTourSportsGroup="NCW" THEN
							sSQL = "SELECT DISTINCT RS.div, DT.div_name FROM " & RawScoresTableName & " AS RS, "& DivisionsTableName & " AS DT "
							sSQL = sSQL + " WHERE RS.div = DT.div AND lower(left(TourID,6)) = '" & sqlclean(lCASE(sTourID)) & "'"
					ELSE
							sSQL = "SELECT DISTINCT RS.div, DT.div_name FROM " & RawScoresOtherTableName & " AS RS, "& DivisionsOtherTableName & " AS DT "
							sSQL = sSQL + " WHERE RS.div = DT.div AND lower(left(TourID,6)) = '" & sqlclean(lCASE(sTourID)) & "'"
					END IF
					sSQL = sSQL + " ORDER BY RS.div"

					set rsSelectFields=Server.CreateObject("ADODB.recordset")
					rsSelectFields.open sSQL, SConnectionToTRATable

					IF not rsSelectFields.eof THEN 
							rsSelectFields.movefirst

							DO WHILE not rsSelectFields.eof
									IF TRIM(rsSelectFields("div")) = DivSelected THEN
											'DivSelected = rsSelectFields("div_name")
											DivSelected = rsSelectFields("div")
											response.write("<option value =""" & rsSelectFields("div") &""" selected>" & rsSelectFields("div") & " - " & rsSelectFields("div_name") & "</option><br>")
									ELSE
											response.write("<option value =""" & rsSelectFields("div") &""">" & rsSelectFields("div") & " - " & rsSelectFields("div_name") & "</option><br>")
									END IF
									rsSelectFields.movenext
							LOOP

					ELSE
							response.write("<option value =""None"" selected>None</option>")
					END IF

					rsSelectFields.close 
					%>
				</select>
			<%
END SUB



' --------------------------
    SUB ScoresByMember
' --------------------------

'sMemberID="http://www.marsbook.co.kr/main/created/product/2/upu/ohoqoh/"



' --- To keep spammers out test for numeric value ---
IF NOT IsNumeric(sMemberID) THEN Response.redirect("/rankings/defaultHQ.asp")

' --- Checks to see if there are any scores for this MemberID ---
SET rsMemb=Server.CreateObject("ADODB.recordset")
sSQL = "Select top 1 PersonIDwithCheckDigit, LastName, FirstName, City, State, BirthDate from "&MemberTableName&" WHERE PersonIDwithCheckDigit="&sqlclean(sMemberID)
rsMemb.open sSQL, sConnectionToTRATable, 3, 1

IF NOT rsMemb.eof THEN
	FullName=rsMemb("FirstName")&" "&rsMemb("LastName")
	sMembCity = rsMemb("City")
	StateSelected = rsMemb("state")
	sMembAge = AgeAtDate(Date, sMemberID)
ELSE
	FullName="Not Defined"
	sMembCity = "Not Defined"
	StateSelected = "Not Defined"
	sMembAge = 150
END IF

'response.write(Session("sSptsGrpID"))

' ------ OUTER TABLE TO HOLD BACKGROUND IMAGE ---- %>
<TABLE border=1 class="droptable" align=center height=225px width="<%=TourDisplayWidth%>" background="<%=MainImage%>" >

<TR >
  <% ' --- Skier Name, MemberID, City/St and Age  --- %> 
   
  <td style="cell-padding:3px" colspan=1 align ="left">
	<font size=4 face="<%=font2%>" color="<%=Textcolor2%>"><b><I><a title="MemberID: <%=sMemberID%>">&nbsp;&nbsp;<%=FullName%></a></I></b></font>
  </td>
  <td colspan=2 align="left">	
	<font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor2%>"><b><%=sMembCity%>, <%=StateSelected %></b></font>
	&nbsp;&nbsp;&nbsp;<font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor2%>"><b>&nbsp;&nbsp;&nbsp;Age: <%=sMembAge%></b></font>
  </td>	
</TR>

<TR>
  <TD colspan=1 width=225px>


  <table height=120px> <% '--- Dropdowns Divider Table --- %>
  <form action="/rankings/<%=ThisFileName%>" method="post">
  <tr> 

    <td width=50px align="right">
	<font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor2%>"><b>Range:</b></font> 
    </td><%

    ' --- Displays Ski Year Drop Down --- %>	
    <td align="left">
    <select name='SkiYear' style="width: 150px"><%

	set rsSelectFields=Server.CreateObject("ADODB.recordset")
	sSQL = "SELECT * FROM " & SkiYearTableName
	rsSelectFields.open sSQL, SConnectionToTRATable

	DO WHILE not rsSelectFields.eof
		IF TRIM(rsSelectFields("SkiYearID")) = session("SkiYear") THEN
			response.write("<option value =""" & rsSelectFields("SkiYearID") &""" selected>")
			response.write(rsSelectFields("SkiYearName"))
			response.write("</option><br>")
  		ELSE
			response.write("<option value =""" & rsSelectFields("SkiYearID") &""">")
			response.write(rsSelectFields("SkiYearName"))
			response.write("</option><br>")
		END IF
		rsSelectFields.movenext
  	LOOP
	rsSelectFields.close %>

      </select>
    </td>
  </tr>

  <% '-----------  Loads Division Pulldown with only divisions found in RawScoresTable ----------- %>
  <tr>
    <td align="right">
 	<font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor2%>"><b>Div:</b></font>
    </td>

    <td>
      <select name='DivSelected' style="width: 150px">
        <option value=""<%IF DivSelected = "" THEN Response.Write(" selected ")%>>All Divisions</option><%

	sSQL = "Select DISTINCT RAW.div, DT.div_name FROM " & RawScoresTableName& " AS RAW"
	sSQL = sSQL + " JOIN " & DivisionsTableName & " AS DT on RAW.div = DT.div where MemberID = '" & sqlclean(sMemberID) & "' ORDER BY RAW.div"
	rsSelectFields.open sSQL, SConnectionToTRATable

	IF not rsSelectFields.eof THEN 
		DO WHILE not rsSelectFields.eof
			IF TRIM(rsSelectFields("div")) = DivSelected THEN
      				DivName = rsSelectFields("div_name")
		      		response.write("<option value =""" & rsSelectFields("div") &""" selected>" & rsSelectFields("div") & " - " & DivName & "</option><br>")
			ELSE
				response.write("<option value =""" & rsSelectFields("div") &""">" & rsSelectFields("div") & " - " & rsSelectFields("div_name") & "</option><br>")
			END IF
			rsSelectFields.movenext
		LOOP
	ELSE
		response.write("<option value =""None"" selected>None</option>")
	END IF

	rsSelectFields.close %>

      </select>
    </td>
  </tr>


  <% ' -------------------------  Loads EVENT Pulldown with only events found in RawScoresTable -------------------- %>
  <tr>
    <td align="right">
	<font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor2%>"><b>Event:</b></font>
    </td>

    <td align=left>
      <select name='EventSelected' style="width: 150px"><%

	sSQL = "Select distinct [event] from " & RawScoresTableName & " where MemberID = '" & sqlclean(sMemberID) & "' order by [event]"
	rsSelectFields.open sSQL, SConnectionToTRATable
	DO WHILE not rsSelectFields.eof
		IF TRIM(rsSelectFields.Fields(0).value) = EventSelected THEN
			response.write("<option value =""" & rsSelectFields.Fields(0).value &""" selected>")
			SELECT CASE rsSelectFields.Fields(0).value
				CASE "T"
					EventName = "Trick"
				CASE "S"
					EventName = "Slalom"
				CASE "J"
					EventName = "Jump"
				CASE "WB"
					EventName = "Wakeboard"
				CASE "WS"
					EventName = "Wake Skate"
				CASE "WB"
					EventName = "Wake Surf"
				CASE "KP"
					EventName = "Flip Out"
				CASE "KR"
					EventName = "Freestyle"

			END SELECT

			%><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor2%>"> <%=EventName%> </font><%
			response.write("</option><br>")

		ELSE
			response.write("<option value =""" & rsSelectFields.Fields(0).value &""">")
			SELECT CASE rsSelectFields.Fields(0).value
				CASE "T"
					response.write("Trick")
				CASE "S"
					response.write("Slalom")
				CASE "J"
					response.write("Jump")
				CASE "WB"
					response.write("Wakeboard")
				CASE "WS"
					response.write("Wake Skate")
				CASE "WB"
					response.write("Wake Surf")
				CASE "KP"
					response.write("Flip Out")
				CASE "KR"
					response.write("Freestyle")


			END SELECT


			response.write("</option><br>")
		END IF
		rsSelectFields.movenext
	LOOP

	rsSelectFields.close

	IF Request("EventSelected") = "O" THEN Response.Write("<option value =""O"" selected>Overall Scores</option><br>")
	IF Request("EventSelected") <> "O" THEN Response.Write("<option value =""O"">Overall Scores</option><br>")  %>
      </select>
    </td>
  </tr>


  <% ' ----------  Loads Class Pulldown ----------  %>
  <tr>
    <td align="right">
	<font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor2%>"><b>Class:</b></font>
    </td> 

    <td>
      <select name="ClassSelected" style="width: 150px"><%
	SELECT CASE sSptsGrpID	
	   CASE "AWS", "NCW"	
		IF Session("NSL") = "0" THEN %>
			<option value="" <%IF ClassSelected = "" THEN Response.Write(" selected ")%>>All Classes</option>
			<option value="LR"<%IF ClassSelected = "LR" THEN Response.Write(" selected ")%>>L or R</option>
			<option value="ELR"<%IF ClassSelected = "ELR" THEN Response.Write(" selected ")%>>E, L or R</option>
			<option value="CELR"<%IF ClassSelected = "CELR" THEN Response.Write(" selected ")%>>C, E, L or R</option> 
			<option value="FNI" <%IF ClassSelected = "FNI" THEN Response.Write(" selected ")%>>F, N or I</option><%
		ELSE %>
			<option value="F" selected>F, N or I</option><%
		END IF 

	   CASE "AKA"  %>
		<option value=""<%IF ClassSelected = "" THEN Response.Write(" selected ")%>>All Classes</option>
		<option value="T"<%IF ClassSelected = "T" THEN Response.Write(" selected ")%>>T</option>
		<option value="Q"<%IF ClassSelected = "Q" THEN Response.Write(" selected ")%>>Q</option><%

	   CASE "USW"  %>
		<option value=""<%IF ClassSelected = "" THEN Response.Write(" selected ")%>>All Classes</option>
		<option value="T"<%IF ClassSelected = "W" THEN Response.Write(" selected ")%>>W</option>
		<option value="Q"<%IF ClassSelected = "F" THEN Response.Write(" selected ")%>>F</option><%
	END SELECT %>

      </select>
      </td>
  </tr>
  </table>

</TD>

<TD align=left> <% ' --- Second column of table --- %>

  <TABLE height=110px border=0>  <% ' ---- Table of Buttons ----  %>	

    <% ' --- Get Scores or Print Screen Button --- %>
    <TR>	
      <TD align="center">


	<input type="hidden" name="pvar" value="ByMember">
	<input type="hidden" name="sMemberID" value="<%=sMemberID%>"><%

	IF sRunByWhat <> "OfficialScores" THEN %>	
	  	<input type=submit style="width:9em" value="Get Scores"><%
	ELSE %>
		<a href='#' onclick='window.print()' title="Click here to Print"><input type=submit value="Print Screen"></a><%
		
		'Response.Write("<a href='#' onclick='window.print()'>Print</a>")
	END IF  %>

      </TD>
      </form>	
     </TR>

    <% ' --- New Member Button or Cancel Button --- %>
    <TR>
      <TD Align=center><%
	IF sRunByWhat <> "OfficialScores" THEN %>	
     <form method=post action="/rankings/<%=ThisFileName%>">
			<input type=submit style="width:9em" value="New Member">
  		<input type="hidden" name="pvar" value="ByMember"><%
	ELSE 
			%><form method=post action="/rankings/<%=ThisFileName%>?sMemberID=<%=sMemberID%>">
					<input type=submit style="width:9em" value="Cancel">
  				<input type="hidden" name="EventSelected" value="<%=EventSelected%>">
  				<input type="hidden" name="DivSelected" value="<%=DivSelected%>">
  				<input type="hidden" name="ClassSelected" value="<%=ClassSelected%>">
  				<input type="hidden" name="pvar" value="ByMember"><%
	END IF 
	%>
   </TD>	
      </form>
    </TR>

    <% ' --- Print Scores Button --- %>
    <TR> 

      <TD align=center>
	 <form action="/rankings/<%=ThisFileName%>" method="post">
          <input type="hidden" name="pvar" value="OfficialScores">
          <input type="hidden" name="sMemberID" value="<%=sMemberID%>">
          <input type="hidden" name="DivSelected" value="<%=DivSelected%>">
          <input type="hidden" name="divname" value="<%=DivName%>">
          <input type="hidden" name="EventSelected" value="<%=EventSelected%>">
          <input type="hidden" name="ClassSelected" value="<%=ClassSelected%>"><%

	IF sRunByWhat <> "OfficialScores" THEN %>	
		<input type="submit" style="width:9em" value="Print Scores"><%
	END IF  

		%>
     </TD>
     </form>
  	</TR>
    <% 

    ' ------------------
    ' --- FAQ Button --- 
    ' ------------------

    %>
    <TR>		
	  <form action="/rankings/tools.asp?svar=FAQ&np=<%=NewsPageNum%>" method="post" target="_blank">
	    <td align=center><%	 
				Session("sSendingPage")="/rankings/"&ThisFileName&"?pvar="&sRunByWhat&"&sMemberID="&sMemberID 
				%>
				<input type="submit" style="width:9em" value="FAQ/Tips">
	    </td>
	  </form>

    </TR>
    </TABLE> 
    <% 
    ' -----------------------------------
    ' --- Bottom of table for buttons --- 
    ' -----------------------------------
    %>
  </TD>
</TR>
</TABLE>

<style TYPE="text/css">
<!--  A:link {text-decoration: none; color:#375AE2}  A:visited {text-decoration: none; color:#375AE2}  A:active {text-decoration: none}   A:hover {text-decoration: ; color:#3E85BB; }-->
</style>
<%









' ------------------------------------------------------------------------------------------------
' -----------  Query for ByMember display  -------------------------------------------------------
' ------------------------------------------------------------------------------------------------



set rsSelectFields=Server.CreateObject("ADODB.recordset")
SET rs=Server.CreateObject("ADODB.recordset")


IF EventSelected = "O" THEN

	sSQL = "Select * from " & OverAllScoresTableName & " as OA"
	sSQL = sSQL + " join " & DivisionsTableName & " as DT on OA.Div = DT.Div and OA.skiyearid = DT.skiyearid"
	sSQL = sSQL + " where MemberID = '" & sqlclean(sMemberID) & "' and substring(OA.Div,1,1) in ('B','G','M','W','O','N')"

	' --- IF not 0, THEN do whatever ski year is indicated by the id provided.
	rsSelectFields.open ("Select top 1 * from " & SkiYearTableName & " where SkiYearID = " & sqlclean(Session("SkiYear"))), SConnectionToTRATable, 3, 3  

	' --- IF the year provided is not a valid id, THEN just ruin the search with 1=0.
        IF rsSelectFields.eof THEN
		sSQL = sSQL + " and 1 = 0"
        ELSE
		' --- IF the year is found, THEN use the begin and end dates to filter the query.
		sSQL = sSQL + " and OA.SkiYearID = " & rsSelectFields("SkiYearID") 
        END IF
	rsSelectFields.close

	IF DivSelected <> "" THEN
		sSQL = sSQL + " and OA.[div] = '" & sqlclean(DivSelected) & "'"
	END IF

	sSQL = sSQL + " order by OA.TotalOverall DESC, OA.round"    
	rs.open sSQL, SConnectionToTRATable
ELSE


	
	' ------ REVISED FORMAT --------
	IF sSptsGrpID="AWS" OR sSptsGrpID="NCW" THEN
		sSQL = "SELECT DISTINCT RAW.*, DT.ZBSConversion FROM "&RawScoresTableName&" AS RAW" 
		sSQL = sSQL + ", "&MemberTableName&" AS MEM, "&SkiYearTableName&" AS SY, "&DivisionsTableName&" AS DT"
	ELSE
		sSQL = "SELECT DISTINCT RAW.*, DT.ZBSConversion FROM "&RawScoresOtherTableName&" AS RAW" 
		sSQL = sSQL + ", "&MemberTableName&" AS MEM, "&SkiYearTableName&" AS SY, "&DivisionsOtherTableName&" AS DT"
	END IF

		sSQL = sSQL + " WHERE RAW.MemberID = '" & sqlclean(sMemberID) & "' AND RAW.event = '" & sqlclean(EventSelected) & "'"    
		sSQL = sSQL + " AND MEM.PersonIDWithCheckDigit = RAW.MemberID AND RAW.EndDate BETWEEN SY.BeginDate AND SY.EndDate"
		sSQL = sSQL + " AND RAW.Div = DT.Div AND SY.skiyearid = DT.skiyearid"


    	' IF not 0, THEN do whatever ski year is indicated by the id provided.
    	rsSelectFields.open ("Select top 1 * from " & SkiYearTableName & " where SkiYearID = " & sqlclean(Session("SkiYear"))), SConnectionToTRATable, 3, 3  

	' IF the year provided is not a valid id, THEN just ruin the search with 1=0.
        IF rsSelectFields.eof THEN

        	sSQL = sSQL + " and 1 = 0"
        ELSE
      		' IF the year is found, THEN use the begin and end dates to filter the query.
          	sSQL = sSQL + " AND (RAW.Enddate <= '" & FormatDateTime(rsSelectFields("EndDate"),2) & "' AND RAW.EndDate >= '" & FormatDateTime(rsSelectFields("BeginDate"),2) & "')"
        END IF
	rsSelectFields.close
    
	IF DivSelected <> "" THEN
      		sSQL = sSQL + " AND RAW.div = '" & sqlclean(DivSelected) & "'"
	END IF


	SELECT CASE ClassSelected
      		CASE "LR"
			sSQL = sSQL + " and lower(RAW.class) in ('l','r')"
		CASE "ELR"
			sSQL = sSQL + " and lower(RAW.class) in ('e','l','r')"
		CASE "CELR"
			sSQL = sSQL + " and lower(RAW.class) in ('c','e','l','r')"
		CASE "FNI"
			sSQL = sSQL + " and lower(RAW.class) in ('f','n','i')"
      		CASE "Q"
			sSQL = sSQL + " and lower(RAW.class) = 'q'"
      		CASE "T"
			sSQL = sSQL + " and lower(RAW.class) = 't'"
      		CASE "W"
			sSQL = sSQL + " and lower(RAW.class) = 'w'"
      		CASE "F"
			sSQL = sSQL + " and lower(RAW.class) = 'f'"

	END SELECT 
	sSQL = sSQL + " ORDER by RAW.Event, RAW.score DESC, RAW.altscore DESC"

'IF adminmenulevel>=30 THEN markdebug(sSQL)
	rs.open sSQL, sConnectionToTRATable, 3, 1

END IF


'response.write(sSQL)

DisplayScoresData



END SUB







' ------------------------
    SUB DisplayScoresData
' ------------------------

'response.write(rs.eof)

IF rs.eof AND request("EventSelected") <> "O" THEN  %>
	<br><br>
	<center><font color="red">No Scores Found For These Search Criteria</font></center>
	<br><br><% 


ELSE 

IF sRunByWhat<>"OfficialScores" THEN
	ScoreBackground=Tablecolor1 
ELSE
	ScoreBackground="" 
END IF


' ---------  TOP OF TABLE FOR DISPLAYING SCORES ------------------ %>

<br>
<TABLE class="innertable" align="center" WIDTH="<%=TourDisplayWidth%>px">

<TR>
  <Th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Tour ID</FONT></th><%



	IF EventSelected <> "O" THEN

		IF Session("NSL") <> "1" THEN %>

			<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Score</FONT></th><%

			IF DivSelected = "CM" OR DivSelected = "CW"THEN %>
				<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Team</FONT></th><%
			END IF %>

			<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Round</FONT></th><%

			IF DivName = "" THEN %>
				<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Div</FONT></th><%
			END IF %>

			<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="<%=Textcolor1%>">Place</FONT></th>
			<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="<%=Textcolor1%>">PPts</FONT></th><%

		        IF left(EventSelected,1) = "S" THEN %>
				<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="<%=Textcolor1%>">Buoys</FONT></th>
				<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="<%=Textcolor1%>">Line</FONT></th><%
			END IF
		        IF left(EventSelected,1) = "J" THEN %>
				<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Ramp</FONT></th><%
			END IF
		        IF left(EventSelected,1) <> "T" THEN %>
				<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Speed</FONT></th><%
			END IF %>

			<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Class</FONT></th><%

		ELSE ' End of NSL Stuff %> 
      
			<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Placement Points</FONT></th>
			<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Div</FONT></th>
			<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Place</FONT></th>
			<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Round</FONT></th>
			<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Score</FONT></th><%
    
		END IF

	ELSE ' Overall Scores Stuff %>
      
		<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Round</FONT></th>
		<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Div</FONT></th>
		<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Slalom</FONT></th>
		<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Trick</FONT></th>
		<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">Jump</FONT></th>
		<th ALIGN="Center" vAlign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>"><b>Total Score</b></FONT></th><%
    
	END IF %>

    
</TR><%


    

    ' ------------------  Loop to begin displaying SCORES for Tournament  ------------------	
    DO WHILE not rs.eof  %>

	<TR>
	  <TD ALIGN="Center" vAlign="top">
	    <font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>"><% 

		SET rsSelectFields=Server.CreateObject("ADODB.recordset")
	     	sSQL = "Select top 1 TSanction,TName,TCity,TState,TDateE from "& SanctionTableName &" where lower(TournAppID) = '" & sqlclean(lCASE(TRIM(left(rs("TourID"),6)))) & "'"
		rsSelectFields.open sSQL, sConnectionToSanctionTable, 3, 1
  

		IF rsSelectFields.EOF THEN %>
			<a href="/rankings/<%=ThisFileName%>?tour_id=<% =TRIM(rs("TourID")) %>&pvar=ByTour&DivSelected=<% =TRIM(rs("Div")) %>"><% =rs("TourID") %></a></FONT></TD><% 
		ELSE %>
			<a href="/rankings/<%=ThisFileName%>?tour_id=<% =TRIM(rs("TourID")) %>&pvar=ByTour&DivSelected=<% =TRIM(rs("Div")) %>&EventSelected=<%=EventSelected%>&sTourSportsGroup=<%=sSptsGrpID%>"
			title="<% =rsSelectFields("tname") %>&#13;<% =rsSelectFields("tcity")%>, <% =rsSelectFields("tstate")%>&#13;<% =rsSelectFields("tdatee")%>"> <% =rs("TourID") %> </a></FONT></TD><%
		END IF 

		rsSelectFields.Close



		IF EventSelected <> "O" THEN
		   IF Session("NSL") <> "1" THEN 

			IF rs("Score") <> "" THEN
				'---  You can not "formatnumber" IF the value is null ... so we throw this check in just to prevent errors.

		        	IF EventSelected = "Trick" THEN %>
	        			<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<% Response.Write formatnumber(rs("score"),0) %></FONT></TD><%
				ELSE %>
        				<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<% IF (Request("ZBSAdjustToOldStyle") = "on") And (rs("Event") = "S") THEN Response.Write formatnumber((rs("score") - rs("ZBSConversion")),2) ELSE Response.Write formatnumber(rs("score"),2) %></FONT></TD><%
				END IF 

			ELSE %>
        			<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;</FONT></TD><%
			END IF 


			IF DivSelected = "CM" OR DivSelected = "CW" THEN %>
				<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<% =rs("team") %></FONT></TD><%
			END IF %>
			<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<% =rs("round") %></FONT></TD><%


		      	IF DivName = "" THEN %>
				<TD Align="center" valign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<a href="/rankings/view-standingsHQ.asp?pvar=National&DivSelected=<%=rs("div")%>&EventSelected=<%=left(EventSelected,1)%>"><% =rs("div") %></a></font></td><% 
			END IF %>
			<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<% =rs("place") %></FONT></TD>
			<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<% =rs("NSL_Placement_Points") %></FONT></TD><%

			IF left(EventSelected,1) = "S" THEN %>
				<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<%=rs("AltScore")%></FONT></TD><%
			END IF


			' --- Display Rope, Boat, Line, Class, etc -----
			IF rs("Perf_Qual1") <> "" THEN
				' --- You can not "formatnumber" IF the value is null ... so we throw this check in just to prevent errors.
				IF left(EventSelected,1) = "S" THEN  %>
					<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<%= formatnumber(rs("Perf_Qual1")/100, 2)%></FONT></TD><%
				ELSE %>
					<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;</FONT></TD><%
				END IF
			ELSE %>
				<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;</FONT></TD><%
		      	END IF

			IF left(EventSelected,1) = "J" THEN %>
				<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<%=rs("Perf_Qual1")%></FONT></TD><%
			END IF
			IF left(EventSelected,1) <> "T" THEN %>
				<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<%=rs("Perf_Qual2")%></FONT></TD><%
			END IF %>
			<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<% =rs("class") %></FONT></TD><% 




		ELSE  ' --- End of non-NSL Stuff  %>

			<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<% =rs("NSL_Placement_Points") %></FONT></TD>
			<TD Align="center" valign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<a href="/rankings/view-standingsHQ.asp?pvar=National&DivSelected=<%=rs("div")%>&EventSelected=<%=left(EventSelected,1)%>"><% =rs("div") %></a></font></td>
			<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<% =rs("place") %></FONT></TD>
			<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<% =rs("round") %></FONT></TD><% 

			IF rs("Score") <> "" THEN
      				'You can not "formatnumber" IF the value is null ... so we throw this check in just to prevent errors.
		
			        IF EventSelected = "T" THEN %>
        	  			<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<% Response.Write formatnumber(rs("score"),0) %></FONT></TD><% 
				ELSE %>
        				<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;<% IF (Request("ZBSAdjustToOldStyle") = "on") And (rs("Event") = "S") THEN Response.Write formatnumber((rs("score") - rs("ZBSConversion")),2) ELSE Response.Write formatnumber(rs("score"),2) %></FONT></TD><% 
				END IF 

			ELSE %>
        			<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=Textcolor1%>">&nbsp;</FONT></TD><%
			END IF 

		END IF  ' -- End of NSL Stuff

	   ELSE  ' Overall Stuff

		IF rs("Div")=rs("DivOrig") THEN 
			Tempcolor = Textcolor1
		ELSE 
			Tempcolor="#FF0000"
		END IF %>

		<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=tempcolor%>">&nbsp;<% =rs("round") %></FONT></TD>
		<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=tempcolor%>">&nbsp;<% =rs("Div") %></FONT></TD>
		<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=tempcolor%>">&nbsp;<a title="<%=rs("S_OrigScore")%>"><% IF rs("SlalomOverall") <> "" THEN Response.Write formatnumber(rs("SlalomOverAll"),1) %></a></FONT></TD>
		<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=tempcolor%>">&nbsp;<a title="<%=rs("T_OrigScore")%>"><% IF rs("TrickOverall") <> "" THEN Response.Write formatnumber(rs("TrickOverAll"),1) %></a></FONT></TD>
		<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=tempcolor%>">&nbsp;<a title="<%=rs("J_OrigScore")%>"><% IF rs("JumpOverall") <> "" THEN Response.Write formatnumber(rs("JumpOverAll"),1) %></a></FONT></TD>
		<TD ALIGN="Center" vAlign="top"><font size=<%=fontsize2%> face="<%=font2%>" color="<%=tempcolor%>">&nbsp;<b><% IF rs("TotalOverall") <> "" THEN Response.Write formatnumber(rs("TotalOverAll"),1) %></b></FONT></TD><%


	    END IF %>

    </TR><% 

    rs.movenext
    LOOP  %>



    </TABLE>
    <br>
    <br>
<%


END IF  ' End if for test of existence of scores

rs.close




END SUB






' ----------------------------------------------------------------------------------------------
   SUB DisplayTourSearchFilters
' ----------------------------------------------------------------------------------------------


					SELECT CASE sTourSportsGroup
	   				CASE "AWS" 	
								IF Session("NSL") = "1" THEN
										TourType = "Grassroots"
								ELSE
										TourType = "AWSA"
								END IF  
						CASE "NCW"
								TourType="Collegiate"
	   				CASE "AKA"
								TourType="Kneeboard"
	   				CASE "USW"
								TourType="Wakeboard"
					END SELECT


	IF Session("adminmenulevel")>=50 AND TRIM(sShowSQL)<>"" THEN
			%>
				<%=sSQL%>
			<%
			'response.end
	END IF

	%>
  <form action="/rankings/<%=ThisFileName%>" method="post">
	<TABLE class="droptable" height=215px ALIGN=center width="98%" background="<%=MainImage%>">
	<%
		IF sIncludeScores = "with" THEN 
				%><input type="hidden" name="sIncludeScores" value="with"><% 
		ELSEIF sIncludeScores = "without" THEN 
				%><input type="hidden" name="sIncludeScores" value="without"><% 
		ELSEIF sIncludeScores = "all" THEN 
				%><input type="hidden" name="sIncludeScores" value="all"><% 
		END IF 

	%>
    <tr>
      <td colspan=8 align=left>
				<font color="<%=textcolor2%>" size=4 face=<%=font1%>>
					<%
					' -----------------------------------------------
					' --- Displays heading on Search Filters Box ---
					' -----------------------------------------------
					%>
					<B><I><%=TourType%> Scores Search By Tournament</I></B>
				</font>
			</td>
    </tr>
    <% 
    
    ' -------------------------
    ' --- RANGE of Ski Year ---
    ' -------------------------
    
    %>
    <tr>
      <td align="right" width=60px>
				<font color="<%=textcolor2%>" size=<%=fontsize2%> face=<%=font1%>><b>Range:</b></font>
      </td>
      <td align=left width=155px colspan=2>
      	<%
				' --- Creates a select based on values in the Scores Table ---
				SET rsSelectFields=Server.CreateObject("ADODB.recordset")
				sSQL = "SELECT DISTINCT RT.SkiYearID, SY.SkiYearName"
				sSQL = sSQL + " FROM " &RankTableName&" AS RT"
				sSQL = sSQL + " JOIN " &SkiYearTableName&" AS SY ON RT.SkiYearID = SY.SkiYearID"

				' --- NCWSA does not display 12 Month Rankings
				IF sTourSportsGroup="NCW" THEN
						sSQL = sSQL + " WHERE RT.SkiYearID <> 1"
				END IF
				sSQL = sSQL + " ORDER BY RT.SkiYearID DESC"
        rsSelectFields.open sSQL, SConnectionToTRATable 

				' --- Builds Ski Year Dropdown ---
				%>
				<select name='SkiYear'>
					<%
		       DO WHILE Not rsSelectFields.EOF
    		      IF TRIM(rsSelectFields("SkiYearID")) = Session("SkiYear") THEN
        			    Response.Write("<option value =""" & rsSelectFields("SkiYearID") &""" selected>")
            			Response.Write(rsSelectFields("SkiYearName"))
            			Response.Write("</option><br>")
          		ELSE
            			Response.Write("<option value =""" & rsSelectFields("SkiYearID") &""">")
            			Response.Write(rsSelectFields("SkiYearName"))
            			Response.Write("</option><br>")
          		END IF
          		rsSelectFields.MoveNext
        	LOOP
        	rsSelectFields.Close
        	%>
        </select>
      </td>
			<%
			' --------------------------------
			' ---	TOURNAMENT NAME Text Box ---
      ' --------------------------------
      %>
      <td align="right" width=80px>
        <font color="<%=textcolor2%>" size=<%=fontsize2%> face=<%=font1%>><b>Tour Name:</b></font>
      </td>
      <td colspan=4>			
				<input type="text" name="Tour_Name" value="<%=sTourName%>" size=20><br>
      </td>
    </tr>
		<%

		' --------------------------
		' --- REGION and TOUR ID ---
		' --------------------------

		%>
    <tr>
	    <td align="right">
        <font color="<%=textcolor2%>" size=<%=fontsize2%> face=<%=font1%>><b>Region:</b></font> 
  	  </td>
    	<td colspan=2>	
      	<select name="Tour_Region">
        	<option value=""<%IF sTourRegion = "" THEN Response.Write(" selected ")%>>All Regions</option>
        	<%
					' --- Excludes any region selections from NCWSA ---
					IF sTourSportsGroup<>"NCW" THEN 
							%>
	      		  <option value="C"<%IF sTourRegion = "C" THEN Response.Write(" selected ")%>>S. Central</option>
        			<option value="M"<%IF sTourRegion = "M" THEN Response.Write(" selected ")%>>Midwest</option>
	        		<option value="W"<%IF sTourRegion = "W" THEN Response.Write(" selected ")%>>West</option>
        			<option value="S"<%IF sTourRegion = "S" THEN Response.Write(" selected ")%>>South</option>
	        		<option value="E"<%IF sTourRegion = "E" THEN Response.Write(" selected ")%>>East</option>
	        		<%
					END IF 
					%>
				</select>
    	</td>
    	<% 
    	' --- TourID Text box ---
    	%>
    	<td align="right">
        <font color="<%=textcolor2%>" size=<%=fontsize2%> face=<%=font1%>><b>Tour ID:</font>
    	</td>
    	<td colspan=4>	
				<input type="text" name="Tour_ID" value="<%=sTourID%>" maxlength=8 size=10><br>
    	</td>
    </tr>
		<%

		' ----------------------------------------------
		' --- CLASS Dropdown and City/State Text Box ---
		' ----------------------------------------------

		%>
    <tr>
	    <td align="right">
  		  <font color="<%=textcolor2%>" size=<%=fontsize2%> face=<%=font1%>><b>Class:</b></font> 
    	</td>
	    <td colspan=2>  
	    	<% 
	    	'--- Class of Scores ---  
	    	%>	
    		<select name="Tour_Class">
    			<% 
    			SELECT CASE sTourSportsGroup
						CASE "AWS", "NCW"
								IF Session("NSL") = "0" THEN 
										%><option value=""<%IF sTourClass = "" THEN Response.Write(" selected ")%>>All Classes</option>
        	  				<option value="LR"<%IF sTourClass = "LR" THEN Response.Write(" selected ")%>>L or R</option>
	          				<option value="ELR"<%IF sTourClass = "ELR" THEN Response.Write(" selected ")%>>E, L or R</option>
        	  				<option value="CELR"<%IF sTourClass = "CELR" THEN Response.Write(" selected ")%>>C, E, L or R</option>
        	  				<option value="FNI"<%IF sTourClass = "FNI" THEN Response.Write(" selected ")%>>F, N or I</option><%
								ELSE 
										%><option value="FNI"<%IF sTourClass = "FNI" THEN Response.Write(" selected ")%>>F, N or I</option><%
								END IF
						CASE "AKA" 
								%><option value=""<%IF sTourClass = "" THEN Response.Write(" selected ")%>>All Classes</option>
        				<option value="T"<%IF sTourClass = "T" THEN Response.Write(" selected ")%>>T</option>
	        			<option value="Q"<%IF sTourClass = "Q" THEN Response.Write(" selected ")%>>Q</option><%
						CASE "USW" 
								%><option value=""<%IF sTourClass = "" THEN Response.Write(" selected ")%>>All Classes</option>
        				<option value="W"<%IF sTourClass = "W" THEN Response.Write(" selected ")%>>W</option>
	        			<option value="F"<%IF sTourClass = "F" THEN Response.Write(" selected ")%>>F</option><%
    				END SELECT 
    
    			IF Session("NSL") = "1" THEN 
    					%><option value="F" selected>F or N or I</option><% 
    			END IF 
    			%>
				</select>
			</td>
    	<td align="right">
        <font color="<%=textcolor2%>" size=<%=fontsize2%> face=<%=font1%>><b>City/ST:</b></font>
    	</td>
    	<td colspan=4>	
				<input type="text" name="Tour_City" value="<%=sTourCity%>" size=12>
				<%

				' ---  Build STATE dropdown list ---
				Dim kvar, statearray
				StateArray = Split(USStatesList3,",") 
				%>  
 				<select name="StateSelected">
 					<%
	  			FOR kvar = 0 TO UBOUND(StateArray)
							IF StateSelected = TRIM(StateArray(kvar)) THEN
									response.write("<option value = """&StateSelected&""" SELECTED>"&StateSelected&"</option>")
							ELSE
									response.write("<option value = """&StateArray(kvar)&""">"&StateArray(kvar)&"</option>")
							END IF
	  			NEXT  
	  		
	  			%>
  			</select>
  		</td>
  	</tr>

		<%
		
		' -------------------------------------
		' --- SPORT DIVISION and Start Date ---
  	' -------------------------------------
  	
  	%>
  	<tr>
    	<td align="right">
        <font color="<%=textcolor2%>" size=<%=fontsize2%> face=<%=font1%>><b>Spt Div: </b></font>
    	</td>
      <td colspan=2>	
        <select name="sTourSportsGroup">
    	    <option value="AWS"<%IF sTourSportsGroup = "AWS" Or Session("NSL") = "1" THEN Response.Write(" selected ")%>>AWSA 3-Event</option>
      	 	<option value="NSL"<%IF sTourSportsGroup = "NSL" THEN Response.Write(" selected ")%>>Grassroots</option>
					<option value="NCW"<%IF sTourSportsGroup = "NCW" THEN Response.Write(" selected ")%>>Collegiate</option>
					<option value="AKA"<%IF sTourSportsGroup = "AKA" THEN Response.Write(" selected ")%>>Kneeboard</option>
        </select>
      </td>

      <td align="right">
        <font color="<%=textcolor2%>" size=<%=fontsize2%> face=<%=font1%>><b>Start Date:</b></font> 
      </td>
      <td colspan=4>	
				<input type="text" name="Tour_Date" value="<%=sTourDate%>" size=10>
				<font color="<%=textcolor2%>" size=1 face=<%=font1%>>(mm/dd/yyyy)</font>  
      </td>
  	</tr>
		<%

		' -----------------------------------------------------
		' --- Include Scores Dropdown and table for buttons ---
  	' -----------------------------------------------------

  	%>
  	<tr>
    	<td align="right" >
				<font color="<%=textcolor2%>" size=<%=fontsize2%> face=<%=font1%>><b>Include: </b></font>
    	</td>
      <td align="left" colspan=2>
        <select name="sIncludeScores">
        	<option value="with" <%IF sIncludeScores = "with" THEN Response.Write(" selected ")%>>With Scores</option>
       		<option value="without" <%IF sIncludeScores = "without" THEN Response.Write(" selected ")%>>Without Scores</option>
					<option value="all" <%IF sIncludeScores = "all" THEN Response.Write(" selected ")%>>All Tournaments</option>
        </select>
				<%	
				IF AdminMenuLevel>=50 THEN  
						%>	
						<br>
						<font color="<% =Titlecolor %>" size="<% =fontsize2 %>"><b>Show SQL</b></font>
						<input type=checkbox name="sShowSQL" <% IF sShowSQL="on" THEN response.write "checked" %>>
						<%
				END IF
				%>
      </td>
	 		<td>&nbsp;</td>
	 		<td align=center>	 
				<input type="submit" style="width:9em" name="Action" value="Begin Search">
			</td>
			<td align=center>	 
				<input type="submit" style="width:9em" name="Action" value="By Member">
			</td>
		  <td align=center>
				<input type="submit" style="width:9em" name="Action" value="Print Results">
 			</td>
			<td align=center>
				<a href='/rankings/tools.asp?svar=FAQ&np=<%=NewsPageNum%>' target='_blank' title="Frequently Asked Questions">
					<font color="<%=textcolor2%>" size=2 face=<%=font1%>><b>FAQ</b></font>
				</a>
      </td>
		</tr>


	</TABLE>
	</form>
	<%




	iu=2
	IF iu=1 THEN
		%>

							<form action="/rankings/<%=ThisFileName%>?rid=<%=rid%>" method="post">
								<input type="hidden" name="Tour_Name" value=<%=sTourName%>>
								<input type="hidden" name="Tour_ID" value=<%=sTourID%>>
								<input type="hidden" name="Tour_Class" value=<%=sTourClass%>>
								<input type="hidden" name="Tour_City" value=<%=sTourCity%>> 
								<input type="hidden" name="StateSelected" value=<%=StateSelected%>>
								<input type="hidden" name="Tour_Region" value=<%=sTourRegion%>>
								<input type="hidden" name="sTourSportsGroup" value=<%=sTourSportsGroup%>>
								<input type="hidden" name="Tour_Date" value=<%=sTourDate%>>
		<%
	END IF

END SUB













SUB ChoosePagesSQL(sSQL,sStart, sSize)
  set rs=Server.CreateObject("ADODB.recordset")
  sqlstmt = sSQL
  rs.CursorType = 3
  rs.open sqlstmt, SConnectionToTRATable
END SUB



Function IsRecordSetEmpty
IF rs.bof = true and rs.eof = true THEN
    IsRecordSetEmpty = true
ELSE
    IsRecordSetEmpty = false
END IF
end Function





SUB DoCount(currentPage) 
h = 0

for i = 1 to rs.PageCount
 Response.Write(" <a href=" & chr(34) & ThisPage & "?DivSelected=" & DivSelected & "&ranknum=" & RankNum & "&EventSelected=" & EventSelected & "&currentpage=" &  i  & "&action=" & sAction & chr(34) & ">" & i & "</a>")
h = h +1
next
IF h = 0 THEN h = 1
Response.Write("<BR><Small>Page " & currentPage & " of  " & h & "</SMALL></center><BR><BR>")
END SUB
%>
