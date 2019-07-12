<!--#include file="settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include16.asp"-->
<!--#include virtual="/rankings/tools_Definitions.asp"-->
<!--#include virtual="/rankings/tools_leagues.asp"-->




<%
' --- Last update 2-6-2013 ---


DefineTRAStyles

Dim ThisFileName, sPriorYear, process, MainImage, AdminMenuLevel

Dim RatingLevel
Dim rsList

Dim TeamTypeIDSelected, EventSelected, DivSelected, EventName, RegionSelected, QualifiedCheckbox
Dim SkiYearSelected, ClassTypeSelected
'Dim sLeagueSelected 


Dim ThisTournAppID, LastTournAppID, ThisStartDate, LastStartDate, DiffBetweenStartDates
Dim StartDateSelected, EndDateSelected, EliteSelected

TourTableWidth=675
TabWidth = 1000  	' --- Used in case where report does not have specific parameters

ThisFileName="Report_Generic.asp"
AdminMenuLevel=Session("AdminMenuLevel")




ReadFormVariables 

' --- Process Control ---

sAction=Request("Action")
IF sAction="Return to Menu" THEN 
		process="return"
END IF



' --------------------------------------------------------------------------
' --- Defines the image to be displayed in the drop downs box background ---
' --------------------------------------------------------------------------

WhatDropdownImage EventSelected






' --- Control execution ---
Dim sShowSQL, sStop
sShowSQL = Request("sShowSQL")
sStop = Request("sStop")
IF sStop="on" THEN
		response.write("<br>Stopped program flow")
   	response.end
END IF





Set rs=Server.CreateObject("ADODB.recordset")



SELECT CASE process
	CASE "return"
		response.redirect("/rankings/defaultHQ.asp")

	CASE "awsefloc"
		PageTitle="AWSEF Tournament List Where OLR Donors Exist"
		PageSubTitle="2007 and 2008 Ski Year"
		AWSEFDonorsByLOC
		IF rs.eof THEN 
			DisplayNoRecordsMessage
		ELSE	
			CreatePageHead 1000
			IF NOT rs.eof THEN DisplayResult 1000
		END IF

	CASE "donorlist"
		PageTitle="AWSEF Donor List From Online Registration Program"
		PageSubTitle="2007 and 2008 Ski Year"
		AWSEFDonorList
		CreatePageHead 1000
		IF NOT rs.eof THEN DisplayResult 1000

	CASE "pwlist" 
		PageTitle="User PW List"
		PageSubTitle="Taken From SWIFT"
		UserPWList
		CreatePageHead 1000
		IF NOT rs.eof THEN DisplayResult 1000

	CASE "refund_OLD"

		GetPriorSkiYear

		PageTitle="Refund Report for <br>"&sPriorYear&" Nationals"
		PageSubTitle="Beta Version"
		Refunds
		CreatePageHead 1000
		IF NOT rs.eof THEN DisplayResult 1000

	CASE "refund"

		DIM sTName, sTourID
		GetPriorNationals
		' sTourID="15S999"
		PageTitle="Refund Report for "&sTourID&"<br>"&LEFT(sTName,35)
		PageSubTitle="Beta Version"
		Refunds
		CreatePageHead 730
		IF NOT rs.eof THEN DisplayResult 1000


	CASE "ratinglist"
		PageTitle="Skiers with Rating Level"
		PageSubTitle="Beta Version"
		DisplaySkiersWithRating
		CreatePageHead 730
		IF NOT rs.eof THEN DisplayResult 1000


	CASE "ratinglist_deduped"
		PageTitle="Skiers with Rating Level"
		PageSubTitle="Deduped on PersonID"
		SkiersWithRatings_DedupedNoDivisionListed_NEW
		CreatePageHead 730
		IF NOT rs.eof THEN DisplayResult 1000

	CASE "qualifylist"
		PageTitle="Skiers Qualified For League"
		PageSubTitle="Beta Version"
		DisplayQualifiedSkiers
		CreatePageHead 730
		IF NOT rs.eof THEN DisplayResult 1000

	CASE "eliteskiers"
		PageTitle="Elite Skiers List"
		PageSubTitle="Last 12 Months - USA only"
		EliteSkiers
		CreatePageHead 730
		IF NOT rs.eof THEN DisplayResult 1000




			
	CASE "pblisting"
		PageTitle="Personal Best Requests"
		PageSubTitle="To avoid duplication End Date should be <br> no greater than current date -1 day"
		DisplayPersonalBestList
		CreatePageHead 730
		IF NOT rs.eof THEN DisplayResult 1000

	CASE "grrank"

		PageTitle="Sample GR Ranking Report"
		PageSubTitle="Version with No Division Selection"
		GRRanking
		CreatePageHead 1000
		IF NOT rs.eof THEN DisplayResult 1000

	CASE "loccontacts"
		PageTitle="LOC Contact List"
		PageSubTitle="Tournaments with "&EventName

		LOCContacts
		CreatePageHead 1000
		IF NOT rs.eof THEN DisplayResult 1000

	CASE "leaguequalsummary"
		WriteIndexPageHeader
		PageTitle="League Qualifications Summary"
		PageSubTitle="LeagueID: "&sLeagueSelected
    		LeagueQualSummary
		CreatePageHead 700
		IF NOT rs.eof THEN DisplayResult 700
		WriteIndexPageFooter
	
	CASE "surveyresults"
		WriteIndexPageHeader
		sTourID="19S999"
		PageTitle="Survey Results - 2019 Goode National Championships"
		PageSubTitle="Okeeheelee, FL"
    SurveyResults
		CreatePageHead 700
		IF NOT rs.eof THEN DisplayResult 700
		WriteIndexPageFooter

	CASE "hotelcount"
		WriteIndexPageHeader
		sTourID="15S999"
		PageTitle="Nights Stayed By Hotel"
		PageSubTitle="2015 Goode National Championships"
		Survey_CountByHotel
		CreatePageHead 700
		IF NOT rs.eof THEN DisplayResult 700
		WriteIndexPageFooter

	CASE "hotellist"
		WriteIndexPageHeader
		sTourID="15S999"
		PageTitle="Hotel Answer Options"
		PageSubTitle="2015 Goode National Championships"
		Survey_HotelList
		CreatePageHead 700
		IF NOT rs.eof THEN DisplayResult 700
		WriteIndexPageFooter

	CASE "hoteldetail"
		WriteIndexPageHeader
		sTourID="15S999"
		PageTitle="Hotel Detail"
		PageSubTitle="2015 Goode National Championships"
		Survey_HotelDetail
		CreatePageHead 700
		IF NOT rs.eof THEN DisplayResult 700
		WriteIndexPageFooter


	CASE "bioinfo"
		WriteIndexPageHeader
		sTourID="18M999"
		PageTitle="Skier Bio Info - Alpha Sequence"
		PageSubTitle="2018 Goode National Championships"
		GetBioInfo
		CreatePageHead 700
		IF NOT rs.eof THEN DisplayResult 700
		WriteIndexPageFooter

	CASE "bioinfo-evt"
		WriteIndexPageHeader
		sTourID="18M999"
		PageTitle="Skier Bio Info - Div/Event Sequence"
		PageSubTitle="2018 Goode National Championships"
		GetBioInfo_Evt
		CreatePageHead 700
		IF NOT rs.eof THEN DisplayResult 700
		WriteIndexPageFooter

	CASE "skierlist"
		WriteIndexPageHeader
		'sTourID="13S999"
		PageTitle="Participant List"
		PageSubTitle="Selected League"
		SkierList
		CreatePageHead 700
		IF NOT rs.eof THEN DisplayResult 700
		WriteIndexPageFooter


		
	CASE "nationals"

		' Dim DiffBetweenStartDates
		
		PageTitle="Nationals Entry Flow"
		PageSubTitle="This Year vs. Last Year"

		NationalEntries

		CreatePageHead 700
		IF NOT rs.eof THEN DisplayNationalsResult 700
		NationalTotals

'	CASE "olrtours"
'		WriteIndexPageHeader
'		PageTitle="OLR Tournaments By Year"
'		PageSubTitle="ver 1"
'		OLRToursByYear
'		CreatePageHead 600
'		DisplayResult 600
'		WriteIndexPageFooter

	CASE "olrentries"
		WriteIndexPageHeader
		PageTitle="Participation By Year"
		PageSubTitle="OLR vs All"
		'OLREntriesByYear
		OLRandALLTourStats
		CreatePageHead 600
		IF NOT rs.eof THEN DisplayResult 600
		WriteIndexPageFooter

	CASE "classnf"
		WriteIndexPageHeader
		PageTitle="Tournaments With Class N or F Included"
		PageSubTitle="By Year"

		ClassNorFToursByYear
		CreatePageHead 725
		IF NOT rs.eof THEN DisplayResult 725
		WriteIndexPageFooter

	CASE "olr_ipn_analysis"
		WriteIndexPageHeader
		PageTitle="OLR PayPal IPN Payment Analysis"
		PageSubTitle="Entries After 689410"

		OLR_IPN_Analysis
		CreatePageHead 725
		IF NOT rs.eof THEN DisplayResult 725
		WriteIndexPageFooter

	CASE "olr_ipn_analysis_summary"
		WriteIndexPageHeader
		PageTitle="OLR Pay Pal IPN Payment Summary"
		PageSubTitle="Entries After 689410"

		OLR_IPN_Analysis_Summary
		CreatePageHead 725
		IF NOT rs.eof THEN DisplayResult 725
		WriteIndexPageFooter

	CASE "ridescount"

		WriteIndexPageHeader
		PageTitle="Rides Sount By Year"
		PageSubTitle="Last 4 Years"

		RidesCountByYear
		CreatePageHead 300
		IF NOT rs.eof THEN DisplayResult 300
		WriteIndexPageFooter

	CASE "v_teammemberstatus"
		'WriteIndexPageHeader
		PageTitle="Virtual Team Member Ranking Detail"
		PageSubTitle="Rankings Based on Scoring Members"

		v_TeamMemberStatus
		CreatePageHead 725
		IF NOT rs.eof THEN DisplayResult 725
		'WriteIndexPageFooter

	CASE "v_teamstatus"
		'WriteIndexPageHeader
		PageTitle="V-Team Scoring Status"
		PageSubTitle="Sum of Member Improvement as Team Improvement for Ranking<br>&nbsp;&nbsp;All & Top 2 Members Scoring"

		v_TeamStatus
		CreatePageHead 725
		IF NOT rs.eof THEN DisplayResult 900
		'WriteIndexPageFooter

	CASE "v_teamtypes"
		'WriteIndexPageHeader
		PageTitle="V-Team Type List"
		PageSubTitle="Defines Make-up of each Team Type"

		v_TeamType
		CreatePageHead 725
		IF NOT rs.eof THEN DisplayResult 900
		'WriteIndexPageFooter

		
	CASE ELSE 
		WriteIndexPageHeader
    response.write("Invalid Report")
    
		WriteIndexPageFooter     		
END SELECT






' ---------------------------------------------------------------------------------------
' ------------------  BOTTOM OF MAIN PROGRAM CODE  	---------------------------------
' ---------------------------------------------------------------------------------------



' -----------------------
  SUB ReadFormVariables
' -----------------------  
  
process=TRIM(LCASE(request("process")))

' --- Event and League ---
EventSelected=TRIM(Request("EventSelected"))
IF EventSelected="" THEN EventSelected="J"
RegionSelected = trim(Request("RegionSelected"))
sLeagueSelected=TRIM(Request("sLeagueSelected"))
QualifiedCheckbox = Request("QualifiedCheckbox")
DivSelected = Request("DivSelected")
StartDateSelected = Request("StartDateSelected")
EndDateSelected = Request("EndDateSelected")
EliteSelected = TRIM(Request("EliteSelected"))
SkiYearSelected = TRIM(Request("SkiYearSelected"))
ClassTypeSelected = TRIM(Request("ClassTypeSelected")) 


SELECT CASE EventSelected
	CASE "S"
		EventName="Slalom"
	CASE "T"
		EventName="Trick"
	CASE "J"
		EventName="Jump"
END SELECT

TeamTypeIDSelected = Request("TeamTypeIDSelected")
IF TRIM(TeamTypeIDSelected)="" THEN TeamTypeIDSelected=0

RatingLevel=TRIM(Request("RatingLevel"))


END SUB  




' ------------------------
   SUB LoadRegionPulldown
' ------------------------

%>
<td align=left>
  <select name="RegionSelected">
	<option value=""<%IF RegionSelected = "" THEN Response.Write(" SELECTED ")%>>All Regions</option>
	<option value="1"<%IF RegionSelected = "1" THEN Response.Write(" SELECTED ")%>>S. Central</option>
	<option value="2"<%IF RegionSelected = "2" THEN Response.Write(" SELECTED ")%>>Midwest</option>
	<option value="3"<%IF RegionSelected = "3" THEN Response.Write(" SELECTED ")%>>West</option>
	<option value="4"<%IF RegionSelected = "4" THEN Response.Write(" SELECTED ")%>>South</option>
	<option value="5"<%IF RegionSelected = "5" THEN Response.Write(" SELECTED ")%>>East</option>
  </select>
</td><%


END SUB





' ------------------------
   SUB LoadClassTypePulldown
' ------------------------

%>
<td align=left>
  <select name="ClassTypeSelected">
	<option value=""<%IF ClassTypeSelected = "" THEN Response.Write(" SELECTED ")%>>All Classes</option>
	<option value="GorF"<%IF ClassTypeSelected = "GorF" THEN Response.Write(" SELECTED ")%>>Grassroots or F</option>
	<option value="NoGorF"<%IF ClassTypeSelected = "NoGorF" THEN Response.Write(" SELECTED ")%>>No Grassroots or F</option>
  </select>
</td><%


END SUB





' --------------------
  SUB GetPriorSkiYear
' --------------------


' --- Get prior SkiYear
sSQL = " SELECT SkiYear FROM "&SkiYearTableName&" WHERE SkiYearID=(SELECT SkiYearID FROM "&SkiYearTableName&" WHERE DefaultYear='1')-1"
rs.open sSQL, SConnectionToTRATable
IF NOT rs.eof THEN sPriorYear=rs("SkiYear")
rs.close


END SUB



' -----------------------
  SUB GetPriorNationals 
' ------------------------

' --- Get prior SkiYear
sSQL = " SELECT TournAppID, TName"
sSQL = sSQL + "   FROM "&SanctionTableName&" AS ST"
sSQL = sSQL + " WHERE LEFT(ST.TournAppID,2) = ( SELECT RIGHT(SkiYearName,2)-1 FROM "&SkiYearTableName&" WHERE DefaultYear='1')"
sSQL = sSQL + " AND RIGHT(LEFT(ST.TournAppID,6),3) = '999'"
sSQL = sSQL + " AND RIGHT(LEFT(ST.TournAppID,3),1) IN ('C','E','M','S','W')"
'response.write(sSQL)
'response.end

rs.open sSQL, SConnectionToTRATable
IF NOT rs.EOF THEN
		sTName = rs("TName")	
		sTourID = rs("TournAppID")
END IF
rs.close


END SUB

 

' ---------------------
  SUB DisplayResult (tabwidth)
' ---------------------



	rs.movefirst

	' ---------------  Displays table HEADINGS  ----------------------

	%>
	<TABLE class="innertable" Align=center WIDTH=<%=tabwidth%>px >
	  <TR>
	  <%

		FOR i = 0 TO rs.fields.count - 1
				TempFN = rs.fields(i).name
				j = 0 
				IF trim(rs.fields(i).name)="Team Type" THEN
						 %><th ALIGN="center" width=10% style="<%=RowColor%>"><font color="#FFFFFF" face="<%=font1%>" SIZE="<%=fontsize1%>">Team Rank</font></th><%
				END IF
				
				%><th ALIGN="Center" vAlign="top" nowrap><FONT COlOR="#FFFFFF" FACE="<%=font1%>" SIZE="<%=fontsize1%>"><%=Rs.Fields(i).name%></FONT></th><%
		NEXT 
		
		%>
	  </TR>
	  <%

	' --------------  Display table data here with paging --------------------------
	RowCount = 1 
	DO WHILE NOT rs.eof
			
			
			%>
			<TR>
			<%

			AllowEdit=true

			FOR i = 0 TO rs.fields.count - 1
	
					RowColor=""
					TempFN = rs.fields(i).name
					IF TempFN="TourID" AND process<>"pblisting" THEN
							IF RIGHT(LEFT(rs.Fields(i).value,6),3)="001" OR ( RIGHT(LEFT(rs.Fields(i).value,6),3)="999" AND ThisYear<>LEFT(rs.Fields(i).value,2) ) THEN
									RowColor="background-color:"&scolor08
							ELSEIF ThisYear<>LEFT(rs.Fields(i).value,2) THEN
									RowColor="background-color:"&scolor04
							END IF
					END IF

					IF trim(rs.fields(i).name)="COA Avg" THEN
							%><TD ALIGN="right" width="25%" style="<%= RowColor %>"><font SIZE="1"></font></td><%
					ELSEIF trim(rs.fields(i).name)="Team Type" THEN
							%><TD ALIGN="center" width="10%" style="<%= RowColor %>"><font SIZE="1"><%= RowCount %></font></td><%
					END IF 
			



	    		IF isnull(rs.Fields(i).value) THEN
							%><td ALIGN="center" style="<%=RowColor%>"><font SIZE="1">&nbsp;</font></td><%
    			ELSE
							%><td ALIGN="center" style="<%=RowColor%>"><font SIZE="1"><%= trim(Rs.Fields(i).Value) %></font></td><% 
					END IF  
	

		NEXT	
		
		%>
		</TR>
		<% 

		rowCount = rowCount + 1
		rs.movenext
	
LOOP 
	
		%>
	</TABLE>
<br><br>
<%

IF process="v_teamstatus" THEN
		%>
		<TABLE class="innertable" Align=center style="width:<%=tabwidth%>px"; height:"200px; ">
			<tr>
				<td align="left" style="width:100%; height:250px; vertical-align:top;">
				<b>TERMS & DEFINITIONS:</b>
					<br><br>
					<i>Team Rank</i> - Based on Team Improvement for Scoring Members
					<br><br>
					<i>Team Improvement</i> - Calculated by summing the Member Improvement (for Scoring Members)
					<br><br>
					<i>Member Improvement</i> - Calculated as the difference between Current Ranking Score and Member Event Benchmark
					<br><br>
					<i>Member Event Benchmark</i> - Current formular uses 2014 Ski Year Ranking Score
					<br><br>
					<i># of Scoring Members</i> - As defined in Team Type Rules table (2 used in BOD & Sibling ranking)
				</td>
			</tr>
		</TABLE>		
		<%
ELSEIF process="v_teammemberstatus" THEN
		%>
		<TABLE class="innertable" Align=center style="width:<%=tabwidth%>px"; height:"200px; ">
			<tr>
				<td align="left" style="width:100%; height:250px; vertical-align:top;">
				<b>TERMS & DEFINITIONS:</b>
					<br><br>
					<i>Rank Score</i> - Current 12 Month Rankings
					<br><br>
					<i>Rank Score Last SY Benchmark</i> - Ranking Score Last Ski Year [used as] Member Event Benchmark score
					<br><br>
					<i>Memb Improve</i> - Difference between Current Ranking Score and Member Event Benchmark Score
					<br><br>
					<i>In Team Member Rank</i> - Member ranking within Member's Team
					<br><br>
					<i>Rank Score 2SY Benchmark</i> - Ranking Score from 2nd Prior Ski Year (shown for comparison to Member Event Benchmark score)
				</td>
			</tr>
		</TABLE>		
		<%
END IF



END SUB



' ---------------------------------------
  SUB DisplayNationalsResult (tabwidth)
' ---------------------------------------


	rs.movefirst

'Response.write("<br>Date = "&Date)
'Response.write("<br>LastStartDate = "&LastStartDate)
	' ---------------  Displays table HEADINGS  ----------------------

	%>
	<TABLE class="innertable" Align=center WIDTH=<%=tabwidth%>px >
	  <TR>
	  	<th>
	  		<font color="<%=TextColor5%>" size=<%=fontsize1%>>This Date</font>
	  	</th>
	  	<th>
	  		<font color="<%=TextColor5%>" size=<%=fontsize1%>>Days To Start</font>
	  	</th>
	  	<th>
	  		<font color="<%=TextColor5%>" size=<%=fontsize1%>>Entries</font>
	  	</th>
	  	<th>
	  		<font color="<%=TextColor5%>" size=<%=fontsize1%>>Sub-Total</font>
	  	</th>
	  	<th>
	  		<font color="<%=TextColor5%>" size=<%=fontsize1%>>Last Date</font>
	  	</th>
	  	<th>
	  		<font color="<%=TextColor5%>" size=<%=fontsize1%>>Days To Start</font>
	  	</th>
	  	<th>
	  		<font color="<%=TextColor5%>" size=<%=fontsize1%>>Entries</font>
	  	</th>
	  	<th>
	  		<font color="<%=TextColor5%>" size=<%=fontsize1%>>Sub-Total</font>
	  	</th>
	  </TR><%

	' --------------  Display table data here with paging --------------------------

	TotalThis = CInt(0) 
	TotalLast = CInt(0)

	DO WHILE NOT rs.eof

		ThisTextColor="#000000"
		RowColor="#ffffff"
		IF Rs.Fields(0).Value=ThisStartDate OR DateAdd("d",-DiffBetweenStartDates,Rs.Fields(0).Value)=LastStartDate THEN
				ThisTextColor="Red"
		ELSEIF Rs.Fields(0).Value=Date THEN
				' ThisTextColor="red"
				RowColor="#8EC8FF"
		END IF



		
		%>
 		<TR>
				<td style="background-color:<%=RowColor%>;">
					<font color="<%=ThisTextColor%>" size=<%=fontsize1%>><%=Rs.Fields(0).Value%></font>
				</td>
				<td style="background-color:<%=RowColor%>;">
					<font color="<%=ThisTextColor%>" size=<%=fontsize1%>><%= DateDiff("d", Rs.Fields(0).Value, ThisStartDate) %></font> 
				</td>
				<td style="background-color:<%=RowColor%>;">
					<font color="<%=ThisTextColor%>" size=<%=fontsize1%>><%=Rs.Fields(1).Value%></font>
				</td> 
				<td style="background-color:<%=RowColor%>;">
					<% TotalThis = TotalThis + CInt(Rs.Fields(1).Value)	%>
					<font color="<%=ThisTextColor%>" size=<%=fontsize1%>><%=TotalThis%></font>
				</td>
				<td style="background-color:<%=RowColor%>;">
					<font color="<%=ThisTextColor%>" size=<%=fontsize1%>><%= DateAdd("d",-DiffBetweenStartDates-366,Rs.Fields(0).Value) %></font>
				</td>
				<td style="background-color:<%=RowColor%>;">
					<font color="<%=ThisTextColor%>" size=<%=fontsize1%>><%= DateDiff("d", Rs.Fields(0).Value, LastStartDate+366+DiffBetweenStartDates) %></font> 
				</td>
				<td style="background-color:<%=RowColor%>;">
					<font color="<%=ThisTextColor%>" size=<%=fontsize1%>><%=Rs.Fields(2).Value%></font>
				</td> 
				<td style="background-color:<%=RowColor%>;">
					<% TotalLast = TotalLast + CInt(Rs.Fields(2).Value)	%>
					<font color="<%=ThisTextColor%>" size=<%=fontsize1%>><%=TotalLast%></font>
				</td>
		</TR>
		<% 
		rowCount = rowCount + 1
		rs.movenext
	LOOP 
	
	%>
	</TABLE>
<br><br><%

END SUB




' ----------------------------
  SUB DisplayNoRecordsMessage
' ----------------------------

%>
<br>
<TABLE class="innertable" Align=center WIDTH=1000px height=100>
  <TR>
	<td style="border-style:none;">
		<font color="<%=TextColor2%>" size="3"><b>No Records Found</b></font>
	</td>
  </TR>
</TABLE><%


END SUB



' ----------------------------------
  SUB CreatePageHead (PageHeadWidth)
' ----------------------------------

SetEventImage

Dim backgroundcolor
backgroundcolor=""
IF PageHeadWidth>700 THEN 
		MainImage=""
		backgroundcolor="#FFFFFF"
END IF		 

'response.write("<br>AdminMenuLevel="&AdminMenuLevel)

' drop
%>
<form action="/rankings/<%=ThisFileName%>?process=<%= process %>" method="post">


<TABLE class="droptable" Align=center style="width:<%=PageHeadWidth%>px; height:175px; background-image:<%=MainImage%>; background-color:<%=backgroundcolor%>">

  <% ' --- Total width 8 columns --- %>	
  <TR>
		<td colspan=6 align=left>
			<font color="<%=TextColor2%>" size="3">&nbsp;&nbsp;<b><%=PageTitle%></b></font>
			<br>
			<font color="<%=TextColor1%>" size="2">&nbsp;&nbsp;<b><%=PageSubTitle%></b></font>
		</td>
		<%

		IF AdminMenuLevel>=50 THEN  
				%>	
  			<td colspan=1 valign=top align="left">
					<font COlOR="<% =TitleColor %>" size="<% =fontsize2 %>"><b>Show SQL</b></font>
						<input type=checkbox name="sShowSQL" <% IF sShowSQL="on" THEN response.write "checked" %>>
				</td>
  			<td colspan=1 valign=top align="left">
					<font color="<% =TitleColor %>" size="<% =fontsize2 %>"><b>Stop</b></font>
						<input type=checkbox name="sStop" <% IF sStop="on" THEN response.write "checked" %>>
				</td>
				<%
		ELSE  
			%><td colspan=2 width="25%" valign=top align="left">&nbsp;</td><%
		
		END IF 


		%>
  	</TR>
  	<TR>
  	<%


	' --- Adds specific drop downs depending on PROCESS ---
	SELECT CASE process
		
		CASE "loccontacts"
				%>
				<td align=right>&nbsp;&nbsp;Select Event:</td>
				<td align=left>
					<%
					LoadAWSAEvents 
					%>
				</td>
				<td align=right>&nbsp;&nbsp;SkiYear:</td>
				<td align=left>
					<%
					LoadSkiYearDropdown
					%>
				</td>
				<td align=right>&nbsp;&nbsp;Class:</td>
				<td align=left>
					<%
					LoadClassTypePulldown
					%>
				</td>
				
				
				<td colspan=2>&nbsp;</td>
				<%

		CASE "leaguequalsummary"  
				%>
				<td align=right>&nbsp;&nbsp;League: </td>
				<td align=left>
					<%
					LeagueDropBuild_07162010  
					%>
				</td>
				<td colspan=6>&nbsp;</td>
				<%

		CASE "ratinglist"
				%>
				<td align=right>&nbsp;&nbsp;Min Rating: </td>
				<td align=left>
					<% 
					RatingLevelDropBuild 
					%>
				</td>
				<td align=right>&nbsp;&nbsp;Event:</td>
				<td align=left>
					<% 
					LoadAWSAEvents_AndAll 
					%>
				</td>
				<td align=left>
					<% 
					LoadDivPulldown
					%>
				</td>

				<td colspan=4>&nbsp;</td>
				<%

		CASE "ratinglist_deduped"  
 				%>
				<td align=right>&nbsp;&nbsp;Rating: </td>
				<td align=left>
					<% 
					RatingLevelDropBuild 
					%>
				</td>
				<td align=right>&nbsp;&nbsp;Event:</td>
				<td align=left>
					<% 
					LoadAWSAEvents_AndAll 
					%>
				</td>
				<td align=right>&nbsp;&nbsp;Region:</td>
				<td align=left>
					<% 
					LoadRegionPulldown 
					%>
				</td>
				<td align=left>
					<% 
					LoadDivPulldown
					%>
				</td>
			</tr>
			<tr>
				</td>
				<td align=right>&nbsp;&nbsp;Qualified:</td>
				<td align=left>
					<%
					
					%>
					<input type="checkbox" name="QualifiedCheckbox" <% IF QualifiedCheckbox="on" THEN response.write "checked" %> >
				</td>
				<td colspan=6>&nbsp;</td>
				<%

		CASE "pblisting"  

				%>
				<td align=right>&nbsp;&nbsp;Start:</td>
				<td align=left>
 					 <input type="date" name="StartDateSelected" value="<%=StartDateSelected %>">
				</td>
				<td align=right>&nbsp;&nbsp;End:</td>
				<td align=left>
 					 <input type="date" name="EndDateSelected" value="<%=EndDateSelected %>">
				</td>
				<td colspan=4>&nbsp;</td>
				<%


		CASE "qualifylist"  
 				%>
				<td align=right>&nbsp;&nbsp;League: </td>
				<td align=left>
					<% 
					Session("sSptsGrpID")="AWS"
					LeagueDropBuild_07162010 %>
				</td>
				<td align=right>&nbsp;&nbsp;Event:</td>
				<td align=left>
					<% 
					LoadAWSAEvents_AndAll 
					%>
				</td>

				<td align=right>&nbsp;Home Region:</td>
				<td align=left>
					<% 
					LoadRegionPulldown 
					%>
				</td>
				<td colspan=2>&nbsp;</td>
				<%


		CASE "skierlist"
				%>
				<td align=right>&nbsp;&nbsp;LeagueID: </td>
				<td align=left colspan=2>
					<%
					LeagueDropBuild_SelectFromAll  
					%>
				</td>
				<td align=right colspan=1>&nbsp;&nbsp;Region:</td>
				<td align=left colspan=2>
					<% 
					LoadRegionPulldown 
					%>
				</td>
				<td colspan=2>&nbsp;</td>
				<%

		CASE "eliteskiers"
				%>
				<td align=right>&nbsp;&nbsp;Open/Mast: </td>
				<td align=left>
					<SELECT name='EliteSelected'>
						<option value ='' <%IF EliteSelected = "" THEN Response.Write(" selected ")%>>All </Option><br>
						<option value ='O'<%IF EliteSelected = "O" THEN Response.Write(" selected ")%>>Open</Option><br>
						<option value ='M'<%IF EliteSelected = "M" THEN Response.Write(" selected ")%>>Masters</Option><br>
					</SELECT>
				</td>
				<td colspan=6>&nbsp;</td>
				<%
				
		CASE "v_teammemberstatus", "v_teamstatus"
				%>
				<td align=right colspan=1 width="15%"><font size=1>&nbsp;&nbsp;Team Type: </font></td>
				<td align=left colspan=1 width="25%">
					<%
					BuildTeamType_DropDown 14,12,""  
					%>
				</td>
				<td colspan=4 width="50%">&nbsp;</td>
				<%

		CASE ELSE 
				%>
				<td colspan=2>&nbsp;</td>
				<%
		END SELECT 
		
		%>
  </TR>

  <TR>
		<td colspan=2 width=25% align=center>
			<input type="submit" name="Action" style="width:9em" value="Update Display" title="Submit and reset this form">
		</td>

		<td colspan=2 width=25%  align=center>
			<input type="submit" name="Action" value="Return to Menu">
		</td>
		<td colspan=4>&nbsp;</td>
  </TR>	

</TABLE>

</form>

<%

END SUB





'------------------
 SUB LoadDivPulldown
'------------------

' --- Loads applicable divisions into a division pulldown for each event selected ---
' -- Added 9/1/2016 to avoid errors when sSptsGrpID is not known for undetermined reason ---
Dim ThisSptsGrp
ThisSptsGrp = sSptsGrpID
IF TRIM(sSptsGrpID)="" THEN ThisSptsGrp = "AWS"

' --- Selects division table based on sSptsGrpID ---
 SELECT CASE ThisSptsGrp
		CASE "AKA", "USH", "USW", "ABC"
				ThisDivTable = DivisionsOtherTableName
   	CASE "AWS","NCW"
				ThisDivTable = DivisionsTableName 	
 END SELECT


    opencon
    SET rsSelectFields=Server.CreateObject("ADODB.recordset")
    sSQL = "SELECT DISTINCT DT.div, DT.div_name FROM "&ThisDivTable&" AS DT"

    ' ///////  NOTE - Need to add filter to filter to current SkiYear

	

    SELECT CASE ThisSptsGrp
  	CASE "AWS"
  			sSQL = sSQL + " WHERE lower(left(DT.div,1)) IN ('b','g','m','w','o')"
				'sSQL = sSQL + " WHERE lower(left(DT.div,1)) <> 'i' AND lower(left(DT.div,1)) <> 'y' AND lower(left(DT.div,1)) <> 'x'"
				'sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'n' AND lower(left(DT.div,1)) <> 'c'"
				'sSQL = sSQL + " AND lower(left(DT.div,1)) <> 'l' AND lower(left(DT.div,1)) <> 'e' AND lower(left(DT.div,1)) <> 's'"		
		CASE "AKA"
				sSQL = sSQL + " WHERE lower(left(DT.div,1)) = 'b' or lower(left(DT.div,1)) = 'g'"
		CASE "NCW"
				sSQL = sSQL + " WHERE lower(DT.div) = 'cm' or lower(DT.div) = 'cw'"
		CASE "USH"
				sSQL = sSQL + " WHERE SptsGrpID='USH'"
		CASE "ABC"
				sSQL = sSQL + " WHERE SptsGrpID='ABC'"
    END SELECT
    sSQL = sSQL + " order by DT.div"


'response.write("<br>"&sSQL)
		Session("sSQL 4310") = sSQL
   	rsSelectFields.open sSQL, SConnectionToTRATable



%>
<td align=right>
  <font size=<% =fontsize3 %> color="<% =TextDropcolor2 %>">Division:</font>
</td>
<td align=left>
  <SELECT name="DivSelected" style="width:6em"><%


    IF NOT rsSelectFields.eof THEN 
  	rsSelectFields.movefirst

	Dim DivCounter
	DivCounter = 1

'SkipOver=false
'IF SkipOver=false THEN
	%>
	<option value ="ALL" <% IF DivSelected = "ALL" THEN Response.Write(" SELECTED ") %>>All</option><br>
  	<%
'END IF
  	
	DO WHILE NOT rsSelectFields.eof
		DivCounter = DivCounter + 1
		%>
		<option value = "<%=rsSelectFields("Div")%>"  <% IF DivSelected = rsSelectFields("Div") THEN Response.Write(" SELECTED ") %>> <%=rsSelectFields("Div")%></option><br>")
		<% 
		rsSelectFields.moveNEXT
	LOOP
    ELSE
	response.write("<option value =""None"" selected>None Available</option>")
    END IF 

    rsSelectFields.close  %>

  </select>
</td><%

END SUB


' ---------------------
 SUB v_TeamMemberStatus
' ---------------------

sSQL = "SELECT "
sSQL = sSQL + " RANK() OVER(Partition By tmem.Team_ID ORDER BY RankScore - Rank_LSY_BM DESC) [In Team<br>Member Rank]"
sSQL = sSQL + ", tmem.Team_ID AS [Team ID]"
sSQL = sSQL + ", FirstName+' '+LastName AS [Member Name], Team_Name AS [Team Name]"

sSQL = sSQL + ", State, tmem.Event"
sSQL = sSQL + ", STR(RankScore,5,2) AS [Ranking<br>Score]"
sSQL = sSQL + ", STR(RankScore - Rank_LSY_BM,7,3) AS [Ranking Score<br>Advancement]"
sSQL = sSQL + ", CASE WHEN Rank_LSY_BM/RankScore<0.8 THEN STR(Rank_2PYSY_BM,5,2) ELSE STR(Rank_LSY_BM,5,2) END AS [Rank Score<br>Benchmark]"
' sSQL = sSQL + ", STR(Rank_2PYSY_BM,5,2) AS [Rank Score<br>2SY Benchmark]"
	
sSQL = sSQL + " FROM "&V_TeamMembersTableName&" tmem"
sSQL = sSQL + " JOIN "&MemberShortTableName&" AS memsht ON CAST(RIGHT(tmem.MemberID,8) AS INT)=memsht.PersonID"
	
sSQL = sSQL + " LEFT JOIN"
sSQL = sSQL + " ( SELECT *"
sSQL = sSQL + "	FROM "&V_TeamTableName&") t"
sSQL = sSQL + " ON t.Team_ID=tmem.Team_ID"

sSQL = sSQL + " LEFT JOIN"
sSQL = sSQL + " ( SELECT MemberID, Event, MAX(RankScore) AS Rank_LSY_BM"
sSQL = sSQL + " FROM "&RankTableName 
sSQL = sSQL + " WHERE SkiYearID=20"
sSQL = sSQL + " GROUP BY MemberID, Event) rPY"
sSQL = sSQL + " ON rPY.MemberID=tmem.MemberID AND rPY.Event=tmem.Event"

sSQL = sSQL + " LEFT JOIN"
sSQL = sSQL + " ( SELECT MemberID, Event, MAX(RankScore) AS Rank_2PYSY_BM"
sSQL = sSQL + " FROM "&RankTableName
sSQL = sSQL + " WHERE SkiYearID=17"
sSQL = sSQL + " GROUP BY MemberID, Event) r2PY" 
sSQL = sSQL + " ON r2PY.MemberID=tmem.MemberID AND r2PY.Event=tmem.Event"

sSQL = sSQL + " LEFT JOIN"
sSQL = sSQL + " ( SELECT MemberID, Event, MAX(RankScore) AS RankScore"
sSQL = sSQL + " FROM "&RankTableName
sSQL = sSQL + " WHERE SkiYearID=1"
sSQL = sSQL + " GROUP BY MemberID, Event) r1" 
sSQL = sSQL + " ON r1.MemberID=tmem.MemberID AND r1.Event=tmem.Event"

sSQL = sSQL + " JOIN usawsrank.V_Team_Type ttype ON ttype.Team_Type_ID=t.Team_Type_ID"	

sSQL = sSQL + " WHERE t.Team_Type_ID="&TeamTypeIDSelected

sSQL = sSQL + " ORDER BY RankScore - Rank_LSY_BM DESC, LastName"

'response.write(sSQL)
'response.end

Set rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable


END SUB







' ---------------------
  SUB v_TeamStatus
' ---------------------

sSQL = "SELECT"
sSQL = sSQL + " RANK() OVER(Partition By Team_Type_ID ORDER BY ROUND(SUM(MemberDeltaApplied),2) DESC) AS Rank"
sSQL = sSQL + ", Team_ID AS [Team ID], Team_Name AS [Team Name]"

sSQL = sSQL + ", ROUND(SUM(MemberDeltaApplied),2) AS [Team Improvement<br>Top 2 Scoring]"
sSQL = sSQL + ", SUM(NumMembersApplied) AS [# Scoring<br>Members]"
sSQL = sSQL + ", ROUND(SUM(MemberDelta),2) AS [Team Improvement<br>All Scoring]"
sSQL = sSQL + ", COUNT(MemberID) AS [# Team<br>Members]"

sSQL = sSQL + ", ROUND(SUM(MemberRankApplied),2) AS [Total Score<br>Top 2 Scoring]"
sSQL = sSQL + ", ROUND(SUM(RankScore),2) AS [Total Score<br>All Scoring]"

sSQL = sSQL + " FROM"
sSQL = sSQL + " ("

sSQL = sSQL + " SELECT Team_Type_Description, ttype.Team_Type_ID"
sSQL = sSQL + ", tmem.Team_ID, Team_Name, tmem.MemberID, FirstName, LastName, State"
sSQL = sSQL + ", tmem.Event"
sSQL = sSQL + ", RankScore"
sSQL = sSQL + ", Rank_LSY_BM"
sSQL = sSQL + ", Rank_2PYSY_BM"
sSQL = sSQL + ", RankScore - Rank_LSY_BM AS MemberDelta"
sSQL = sSQL + ", RANK() OVER(Partition By tmem.Team_ID ORDER BY RankScore - Rank_LSY_BM DESC) AS TeamRank"

sSQL = sSQL + ", CASE WHEN RANK() OVER(Partition By tmem.Team_ID ORDER BY RankScore)<=2 THEN RankScore ELSE 0 END AS MemberRankApplied"
sSQL = sSQL + ", CASE WHEN RANK() OVER(Partition By tmem.Team_ID ORDER BY RankScore - Rank_LSY_BM DESC)<=2 THEN RankScore - Rank_LSY_BM ELSE 0 END AS MemberDeltaApplied"
sSQL = sSQL + ", CASE WHEN RANK() OVER(Partition By tmem.Team_ID ORDER BY RankScore - Rank_LSY_BM DESC)<=2 THEN 1 ELSE 0 END AS NumMembersApplied"
	
sSQL = sSQL + " FROM usawsrank.V_Team_Members tmem"
sSQL = sSQL + " JOIN usawaterski.dbo.membershort AS memsht ON CAST(RIGHT(tmem.MemberID,8) AS INT)=memsht.PersonID"
	
sSQL = sSQL + " LEFT JOIN"
sSQL = sSQL + " ( SELECT *"
sSQL = sSQL + "	FROM usawsrank.V_Team) t"
sSQL = sSQL + " ON t.Team_ID=tmem.Team_ID"

sSQL = sSQL + " LEFT JOIN"
sSQL = sSQL + " ( SELECT MemberID, Event, MAX(RankScore) AS Rank_LSY_BM"
sSQL = sSQL + " FROM usawsrank.Rankings" 
sSQL = sSQL + " WHERE SkiYearID=20"
sSQL = sSQL + " GROUP BY MemberID, Event) rPY"
sSQL = sSQL + " ON rPY.MemberID=tmem.MemberID AND rPY.Event=tmem.Event"

sSQL = sSQL + " LEFT JOIN"
sSQL = sSQL + " ( SELECT MemberID, Event, MAX(RankScore) AS Rank_2PYSY_BM"
sSQL = sSQL + " FROM usawsrank.Rankings" 
sSQL = sSQL + " WHERE SkiYearID=17"
sSQL = sSQL + " GROUP BY MemberID, Event) r2PY" 
sSQL = sSQL + " ON r2PY.MemberID=tmem.MemberID AND r2PY.Event=tmem.Event"

sSQL = sSQL + " LEFT JOIN"
sSQL = sSQL + " ( SELECT MemberID, Event, MAX(RankScore) AS RankScore"
sSQL = sSQL + " FROM usawsrank.Rankings" 
sSQL = sSQL + " WHERE SkiYearID=1"
sSQL = sSQL + " GROUP BY MemberID, Event) r1" 
sSQL = sSQL + " ON r1.MemberID=tmem.MemberID AND r1.Event=tmem.Event"


sSQL = sSQL + " JOIN usawsrank.V_Team_Type ttype ON ttype.Team_Type_ID=t.Team_Type_ID"	

sSQL = sSQL + " WHERE t.Team_Type_ID="&TeamTypeIDSelected

sSQL = sSQL + " ) Summary"

sSQL = sSQL + " GROUP BY Team_Type_ID, Team_Type_Description, Team_ID, Team_Name"
sSQL = sSQL + " ORDER BY Team_Type_ID, Team_Type_Description, ROUND(SUM(MemberDeltaApplied),2) DESC"

'response.write(sSQL)
'response.end
Set rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable

END SUB




' --------------
  SUB v_TeamType
' --------------

sSQL = "SELECT Team_Type_Description, Max_Members, Min_Members, Max_Male, Min_Male, Max_Female, Min_Female"
sSQL = sSQL + ", Max_Scoring, Min_Scoring_Male, Min_Scoring_Female, Max_Age, Min_Age, Admin_Level"
sSQL = sSQL + " FROM usawsrank.V_Team_Type"
sSQL = sSQL + " ORDER BY Team_Type_ID"

Set rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable

END SUB







' ---------------------
  SUB SurveyResults
' ---------------------

sSQL = " SELECT DISTINCT RSQ.QuestionDesc" 
sSQL = sSQL + ", RSQ.QuestionID, CASE WHEN RSA1.Answer ='' THEN 'N/A' ELSE RSA1.Answer END AS Answer, RSA1.Count" 
sSQL = sSQL + ", CAST(CAST(RSA1.Count AS Real)/CAST(RSA2.MembCount AS Real) * 100 AS Decimal(5,1)) AS Perc"
'sSQL = sSQL + ", RSA2.MembCount"
sSQL = sSQL + " FROM usawsrank.RegSurveyQuestions RSQ"
sSQL = sSQL + " LEFT JOIN"
sSQL = sSQL + " 	( SELECT TourID,  Answer, QuestionID, Count(QuestionID) AS Count"
sSQL = sSQL + " 	 		FROM usawsrank.RegSurveyAnswers"
sSQL = sSQL + " 	 		WHERE TourID='"&sTourID&"'" 
sSQL = sSQL + " 	 		GROUP BY TourID, QuestionID, Answer"
sSQL = sSQL + " 	 		) AS RSA1"
sSQL = sSQL + " ON RSA1.TourID=RSQ.TourID AND RSA1.QuestionID=RSQ.QuestionID"
sSQL = sSQL + " LEFT JOIN"
sSQL = sSQL + " 	( SELECT TourID, Count(MemberID) AS MembCount"
sSQL = sSQL + " 	 		FROM usawsrank.RegisterGenNew"
sSQL = sSQL + " 	 		WHERE TourID='"&sTourID&"'" 
sSQL = sSQL + " 	 		GROUP BY TourID"
sSQL = sSQL + " 	 		) AS RSA2"
sSQL = sSQL + " ON RSA1.TourID=RSQ.TourID AND RSA1.QuestionID=RSQ.QuestionID"

sSQL = sSQL + " WHERE RSA1.TourID='"&sTourID&"'"
sSQL = sSQL + " ORDER BY RSQ.QuestionID, RSA1.Answer"

'response.write(sSQL)
'response.end
rs.open sSQL, SConnectionToTRATable

END SUB



' -------------------------
  SUB Survey_HotelDetail
' -------------------------

sSQL = " SELECT RSA1.MemberID, RSA1.Answer AS Hotel"
sSQL = sSQL + ", Nights"
sSQL = sSQL + " 	 		FROM usawsrank.RegSurveyAnswers RSA1"
sSQL = sSQL + " LEFT JOIN"
sSQL = sSQL + " ( SELECT MemberID, COALESCE(Answer,0) AS Nights"
sSQL = sSQL + " 	FROM usawsrank.RegSurveyAnswers"
sSQL = sSQL + " 		WHERE QuestionID='6' AND LEFT(TourID,6)='"&sTourID&"') AS RSA2"
sSQL = sSQL + " ON RSA1.MemberID=RSA2.MemberID"

sSQL = sSQL + " WHERE LEFT(RSA1.TourID,6)='"&sTourID&"'"
sSQL = sSQL + " 		AND RSA1.QuestionID='5'"
sSQL = sSQL + " 		AND LEFT(RSA1.Answer,1)<>' '"
sSQL = sSQL + " 		AND RSA1.MemberID<>'000001151'"
sSQL = sSQL + " ORDER BY Hotel"

'response.write(sSQL)
'response.end

rs.open sSQL, SConnectionToTRATable

END SUB


' -------------------------
  SUB Survey_HotelList
' -------------------------

sSQL = " SELECT Distinct RSA1.Answer AS Hotel"
sSQL = sSQL + " 	 		FROM usawsrank.RegSurveyAnswers RSA1"
sSQL = sSQL + " WHERE RSA1.QuestionID='5'"
sSQL = sSQL + " ORDER BY RSA1.Hotel"

'response.write(sSQL)
'response.end

rs.open sSQL, SConnectionToTRATable

END SUB



' -------------------------
   SUB Survey_CountByHotel
' -------------------------

sSQL = "SELECT Hotel, SUM(Nights) AS TotalNights"
sSQL = sSQL + " FROM"
sSQL = sSQL + " ("
sSQL = sSQL + " SELECT RSA1.MemberID, RSA1.Answer AS Hotel, Nights"
sSQL = sSQL + " 	 		FROM usawsrank.RegSurveyAnswers RSA1"
sSQL = sSQL + " LEFT JOIN"
sSQL = sSQL + " ( SELECT MemberID, COALESCE(Answer,0) AS Nights"
sSQL = sSQL + " 	FROM usawsrank.RegSurveyAnswers"
sSQL = sSQL + " 		WHERE QuestionID='6' AND LEFT(TourID,6)='"&sTourID&"') AS RSA2"
sSQL = sSQL + " ON RSA1.MemberID=RSA2.MemberID"
sSQL = sSQL + " WHERE LEFT(RSA1.TourID,6)='"&sTourID&"'"
sSQL = sSQL + " 		AND RSA1.QuestionID='5'"
sSQL = sSQL + " 		AND RSA1.MemberID<>'000001151'"
sSQL = sSQL + " ) RSA"
sSQL = sSQL + " GROUP BY Hotel"
sSQL = sSQL + " ORDER BY Hotel "

'response.write(sSQL)
'response.end

rs.open sSQL, SConnectionToTRATable

END SUB




' -------------------------------
  SUB DisplaySkiersWithRating 
' -------------------------------


sSQL = " SELECT DISTINCT MT.lastname, MT.firstname, Address1, Address2, City, State, Zip, Div, AWSA_Rat AS LvL"
sSQL = sSQL + " , MT.Email, MT.DoNotEmail"
sSQL = sSQL + " FROM "&RankTableName&" as RT"	

sSQL = sSQL + " JOIN "&MemberShortTableName&" AS MT "
sSQL = sSQL + " ON CAST(RIGHT(RT.MemberID,8) AS INT)=MT.PersonID" 

sSQL = sSQL + " JOIN "&SkiYearTableName&" as SY "
sSQL = sSQL + " ON SY.SkiYearID = RT.SkiYearID " 


sSQL = sSQL + " WHERE RT.SkiYearID = 1 AND MT.federationcode = 'USA' AND RT.RankScore is not NULL "
sSQL = sSQL + " AND RIGHT(AWSA_Rat,1) >='"&RatingLevel&"'"
sSQL = sSQL + " AND FederationCode ='USA'"
IF EventSelected="S" OR EventSelected="J" OR EventSelected="T" THEN
		sSQL = sSQL + " AND RT.[event] = '"&EventSelected&"'"
END IF
IF DivSelected<>"ALL" THEN
		sSQL = sSQL + " AND RT.[div] = '"&DivSelected&"'"
END IF


	
'-- SELECT MAX(MT.lastname) AS lastname, MAX(MT.firstname) AS firstname
'-- 	, MAX(Address1) AS Address1, MAX(Address2) AS Address2, MAX(City) AS City, MAX(State) AS State, MAX(Zip) AS Zip
'-- 	, MAX(MT.Email) AS Email, MAX(MT.DoNotEmail) AS DoNotEmail 
'-- 	, CASE WHEN LEFT(AWSA_Rat,1)='S' THEN Div END AS S_Div
'-- 	, CASE WHEN LEFT(AWSA_Rat,1)='T' THEN Div END AS T_Div
'--   , CASE WHEN LEFT(AWSA_Rat,1)='J' THEN Div END AS J_Div

'--  , CASE WHEN LEFT(AWSA_Rat,1)='S' THEN RIGHT(Div,1) END AS S_Level  
'--  , CASE WHEN LEFT(AWSA_Rat,1)='T' THEN RIGHT(Div,1) END AS T_Level
'--  , CASE WHEN LEFT(AWSA_Rat,1)='J' THEN RIGHT(Div,1) END AS J_Level

'-- FROM usawsrank.Rankings as RT 
'-- JOIN USAWaterski.dbo.membershort AS MT ON CAST(RIGHT(RT.MemberID,8) AS INT)=MT.PersonID 
'-- JOIN usawsrank.SkiYear as SY ON SY.SkiYearID = RT.SkiYearID 

'-- WHERE RT.SkiYearID = 1 
'-- 	AND LOWER(LEFT(Div,1)) IN ('b','g','m','w','o')
'-- 	AND MT.federationcode = 'USA' 
'-- 		AND RT.RankScore is not NULL AND RIGHT(AWSA_Rat,1) >='8' AND FederationCode ='USA'
'-- GROUP BY RT.MemberID
'-- ORDER BY MT.lastname, MT.firstname


' response.write(sSQL)
' response.end

rs.open sSQL, SConnectionToTRATable


END SUB



' -----------------------------------------------
  SUB SkiersWithRatings_DedupedNoDivisionListed
' -----------------------------------------------   

' IF TRIM(RatingLevel)="" THEN RatingLevel=8
	
' -- Alternate selection - used to select Members for Nationals Qualification email --
sSQL = "  SELECT *"
sSQL = sSQL + "  FROM ("
sSQL = sSQL + "  	SELECT" 
sSQL = sSQL + "  		CASE WHEN LEN(m.Email)-LEN(REPLACE(m.Email,'@',''))=1 AND CHARINDEX('..',m.Email)=0 AND CHARINDEX('.',m.Email)>0 AND CHARINDEX(')',m.Email)=0 AND CHARINDEX('/',m.Email)=0 AND CHARINDEX(':',m.Email)=0 AND CHARINDEX(' ',m.Email)=0 THEN m.Email ELSE '' END AS Email"
sSQL = sSQL + "  		, PersonID, MemberID, FirstName, LastName, Div"
sSQL = sSQL + "  		, COALESCE( CAST(CAST(Slalom_Level AS INTEGER) AS CHAR),' ') AS Slalom_Level"
sSQL = sSQL + "  		, COALESCE( CAST(CAST(Trick_Level AS INTEGER) AS CHAR),' ') AS Trick_Level"
sSQL = sSQL + "  		, COALESCE( CAST(CAST(Jump_Level AS INTEGER) AS CHAR),' ') AS Jump_Level"

sSQL = sSQL + "  		FROM "&MemberLiveTableName&" m"
sSQL = sSQL + "  	LEFT JOIN" 
sSQL = sSQL + "  		(	"	
sSQL = sSQL + "  			SELECT MemberID"
sSQL = sSQL + "  				, MAX(CASE WHEN Event='S' THEN Rank_Level END) AS Slalom_Level"
sSQL = sSQL + "  				, MAX(CASE WHEN Event='T' THEN Rank_Level END) AS Trick_Level"
sSQL = sSQL + "  				, MAX(CASE WHEN Event='J' THEN Rank_Level END) AS Jump_Level" 
sSQL = sSQL + "  				FROM "&RankTableName
sSQL = sSQL + "  				  WHERE ( (Event='S' AND Rank_Level>='"&RatingLevel&"') OR (Event='T' AND Rank_Level>='"&RatingLevel&"') OR (Event='J' AND Rank_Level>='"&RatingLevel&"') )"
sSQL = sSQL + "  				  		AND SkiYearID=1"
sSQL = sSQL + " 				GROUP BY MemberID"
sSQL = sSQL + "  		) st ON CAST(RIGHT(st.MemberID,8) AS INTEGER)=PersonID"
sSQL = sSQL + "  ) a"
sSQL = sSQL + "  	WHERE MemberID IS NOT NULL AND Email IS NOT NULL AND LOWER(Email)<>'deceased' AND LEN(Email)>5"
IF DivSelected<>"ALL" THEN
		sSQL = sSQL + " AND RT.[div] = '"&DivSelected&"'"
END IF
sSQL = sSQL + "  ORDER BY LastName ASC, FirstName ASC;"

'response.write(sSQL)
'response.end

rs.open sSQL, SConnectionToTRATable

END SUB




' ---------------------------------------------------
  SUB SkiersWithRatings_DedupedNoDivisionListed_NEW
' ---------------------------------------------------  

' IF TRIM(RatingLevel)="" THEN RatingLevel=8
	
' -- Alternate selection - used to select Members for Nationals Qualification email --
sSQL = "  SELECT *"
sSQL = sSQL + "  FROM ("
sSQL = sSQL + "  	SELECT" 
sSQL = sSQL + "  		CASE WHEN LEN(m.Email)-LEN(REPLACE(m.Email,'@',''))=1 AND CHARINDEX('..',m.Email)=0 AND CHARINDEX('.',m.Email)>0 AND CHARINDEX(')',m.Email)=0 AND CHARINDEX('/',m.Email)=0 AND CHARINDEX(':',m.Email)=0 AND CHARINDEX(' ',m.Email)=0 THEN m.Email ELSE '' END AS Email"
sSQL = sSQL + "  		, PersonID, st.MemberID AS QfyMemb, ra.MemberID AS RankMemb, FirstName, LastName, m.State, r.Region"
sSQL = sSQL + "  		, COALESCE( Slalom_Qfy,0) AS Slalom_Qfy"
sSQL = sSQL + "  		, COALESCE( Trick_Qfy,0) AS Trick_Qfy"
sSQL = sSQL + "  		, COALESCE( Jump_Qfy,0) AS Jump_Qfy"
sSQL = sSQL + "  		, COALESCE( Slalom_Rat,0) AS Slalom_Rat"
sSQL = sSQL + "  		, COALESCE( Trick_Rat,0) AS Trick_Rat"
sSQL = sSQL + "  		, COALESCE( Jump_Rat,0) AS Jump_Rat"

sSQL = sSQL + "  		FROM "&MemberLiveTableName&" m"
sSQL = sSQL + "  	JOIN "&RegionTableName&" r ON m.State=r.State"
sSQL = sSQL + "  	LEFT JOIN" 
sSQL = sSQL + "  		(	"	
sSQL = sSQL + "  			SELECT MemberID"
sSQL = sSQL + "  				, SUM(CASE WHEN Event='S' THEN 1 ELSE 0 END) AS Slalom_Qfy"
sSQL = sSQL + "  				, SUM(CASE WHEN Event='T' THEN 1 ELSE 0 END) AS Trick_Qfy"
sSQL = sSQL + "  				, SUM(CASE WHEN Event='J' THEN 1 ELSE 0 END) AS Jump_Qfy" 
sSQL = sSQL + "  				FROM "&RegQualifyTableName&" q"
sSQL = sSQL + "  				JOIN "&SkiYearTableName&" sy ON LEFT(q.TourID,2)=RIGHT(sy.SkiYear,2) AND DefaultYear=1"
sSQL = sSQL + "  						WHERE SUBSTRING(TourID,4,3)='999'"
sSQL = sSQL + "  				    	AND QfyStatus IN ('QFY-RPR','Qualified')"


sSQL = sSQL + "  			GROUP BY MemberID"
sSQL = sSQL + "  		) st ON CAST(RIGHT(st.MemberID,8) AS INTEGER)=PersonID"

sSQL = sSQL + "  	JOIN" 
sSQL = sSQL + "  		(	"	
sSQL = sSQL + "  			SELECT MemberID, MAX(Div) AS Div"
sSQL = sSQL + "  				, MAX(CASE WHEN Event='S' AND RIGHT(AWSA_Rat,1)>='"&RatingLevel&"' THEN RIGHT(AWSA_Rat,1) ELSE '' END) AS Slalom_Rat"
sSQL = sSQL + "  				, MAX(CASE WHEN Event='T' AND RIGHT(AWSA_Rat,1)>='"&RatingLevel&"' THEN RIGHT(AWSA_Rat,1) ELSE '' END) AS Trick_Rat"
sSQL = sSQL + "  				, MAX(CASE WHEN Event='J' AND RIGHT(AWSA_Rat,1)>='"&RatingLevel&"' THEN RIGHT(AWSA_Rat,1) ELSE '' END) AS Jump_Rat" 
sSQL = sSQL + "  				FROM "&RankTableName&" r"
sSQL = sSQL + "  				JOIN "&SkiYearTableName&" sy ON r.SkiYearID = sy.SkiYearID AND DefaultYear=1"
IF DivSelected<>"ALL" AND NOT(ISNULL(DivSelected)) THEN sSQL = sSQL + " AND div = '"&DivSelected&"'"
sSQL = sSQL + "  			GROUP BY MemberID) ra "
sSQL = sSQL + "  		ON CAST(RIGHT(ra.MemberID,8) AS INTEGER)=PersonID"


sSQL = sSQL + "  	     WHERE FederationCode='USA'"

sSQL = sSQL + "  ) a"
sSQL = sSQL + "  	WHERE RankMemb IS NOT NULL AND Email IS NOT NULL AND LOWER(Email)<>'deceased' AND LEN(Email)>5"

IF RegionSelected<>"ALL" AND RegionSelected<>"" THEN sSQL = sSQL + "  	        AND Region = '"&RegionSelected&"'"



' -- Qualified for League --
IF QualifiedSelected = "on" THEN 
		IF TRIM(EventSelected) = "S" THEN 
				sSQL = sSQL + "  	        AND Slalom_Qfy>=1"
		ELSEIF TRIM(EventSelected) = "T" THEN 
				sSQL = sSQL + "  	        AND Trick_Qfy>=1"
		ELSEIF TRIM(EventSelected) = "J" THEN 
				sSQL = sSQL + "  	        AND Jump_Qfy>=1"
		END IF
END IF


' -- Ranking Level applies to all events --
IF TRIM(RatingLevel)<>"" THEN 
		IF TRIM(EventSelected) = "S" THEN 
				sSQL = sSQL + "  	        AND Slalom_Rat>=1"
		ELSEIF TRIM(EventSelected) = "T" THEN 
				sSQL = sSQL + "  	        AND Trick_Rat>=1"
		ELSEIF TRIM(EventSelected) = "J" THEN 
				sSQL = sSQL + "  	        AND Jump_Rat>=1"
		ELSE
				sSQL = sSQL + "  	        AND (Slalom_Rat>=1 OR Trick_Rat>=1 OR Jump_Rat>=1)"
		END IF
		
END IF


sSQL = sSQL + "  ORDER BY LastName ASC, FirstName ASC;"


' response.write(sSQL)
' response.end

rs.open sSQL, SConnectionToTRATable

END SUB




' --------------------------
  SUB DisplayQualifiedSkiers
' --------------------------

sLeagueSelected = request("sLeagueSelected")

sSQL = "SELECT Email, MemberID, FirstName, LastName, m.State, rq.*"

sSQL = sSQL + "FROM"
sSQL = sSQL + "	( SELECT MemberID"
	
sSQL = sSQL + "			, MAX(CASE WHEN Event='S' AND (QfyStatus='Qualified' OR QfyStatus='QFY-RPR' OR QfyStatus='Pending') THEN Div ELSE '' END) AS Slalom"
sSQL = sSQL + "			, MAX(CASE WHEN Event='T' AND (QfyStatus='Qualified' OR QfyStatus='QFY-RPR' OR QfyStatus='Pending') THEN Div ELSE '' END) AS Trick"
sSQL = sSQL + "			, MAX(CASE WHEN Event='J' AND (QfyStatus='Qualified' OR QfyStatus='QFY-RPR' OR QfyStatus='Pending') THEN Div ELSE '' END) AS Jump"  
	
sSQL = sSQL + "			FROM "&RegQualifyTableName&" rq1"
sSQL = sSQL + "     LEFT JOIN "&LeagueTableName&" lt ON lt.QualifyTour=rq1.TourID" 
sSQL = sSQL + "						WHERE LeagueID='"&sLeagueSelected&"'" 
IF TRIM(EventSelected) <> "ALL" AND NOT(ISNULL(EventSelected)) THEN sSQL = sSQL + " AND rq1.event = '"&EventSelected&"'"
sSQL = sSQL + "			GROUP BY MemberID, Div ) rq"
	
sSQL = sSQL + "		JOIN "&MemberLiveTableName&" m ON m.PersonID=RIGHT(rq.MemberID,8)"
sSQL = sSQL + "  	JOIN "&RegionTableName&" r ON m.State=r.State"

sSQL = sSQL + "		WHERE SUBSTRING(Email,1,1)<>' ' AND Email IS NOT NULL" 
sSQL = sSQL + "			AND (Slalom IS NOT NULL OR Trick IS NOT NULL OR Jump IS NOT NULL)"
sSQL = sSQL + "			AND (Slalom<>'' OR Trick<>'' OR Jump<>'')" 
sSQL = sSQL + "			AND (FederationCode ='USA' OR FederationCode IS NULL)" 
IF RegionSelected<>"ALL" AND RegionSelected<>"" THEN sSQL = sSQL + "  	        AND m.Region = '"&RegionSelected&"'"

sSQL = sSQL + "	ORDER BY LastName, FirstName"

' response.write(sSQL)

rs.open sSQL, SConnectionToTRATable

END SUB
		




' ------------------
  SUB EliteSkiers 
' ------------------


sSQL = "	SELECT MAX(FirstName) AS FirstName, MAX(LastName) AS LastName"
sSQL = sSQL + "		, MAX(Email) AS Email, ed.MemberID"
	
sSQL = sSQL + "		, CASE WHEN COALESCE(sl.MemberID,0)<>0 THEN MAX(CAST(sl.QualThru AS CHAR)) ELSE '' END AS Sl_Open"
sSQL = sSQL + "		, CASE WHEN COALESCE(tr.MemberID,0)<>0 THEN MAX(CAST(tr.QualThru AS CHAR)) ELSE '' END AS Tr_Open"
sSQL = sSQL + "		, CASE WHEN COALESCE(ju.MemberID,0)<>0 THEN MAX(CAST(ju.QualThru AS CHAR)) ELSE '' END AS Ju_Open"
	
sSQL = sSQL + "		, CASE WHEN COALESCE(slm.MemberID,0)<>0 THEN MAX(CAST(slm.QualThru AS CHAR)) ELSE '' END AS Sl_Masters"
sSQL = sSQL + "		, CASE WHEN COALESCE(trm.MemberID,0)<>0 THEN MAX(CAST(trm.QualThru AS CHAR)) ELSE '' END AS Tr_Masters"
sSQL = sSQL + "		, CASE WHEN COALESCE(jum.MemberID,0)<>0 THEN MAX(CAST(jum.QualThru AS CHAR)) ELSE '' END AS Ju_Masters"

' sSQL = sSQL + "		, CASE WHEN COALESCE(sl.MemberID,0)<>0 THEN MAX(sl.DivElite) ELSE '' END AS Sl_Div"
'sSQL = sSQL + "		, CASE WHEN COALESCE(tr.MemberID,0)<>0 THEN MAX(tr.DivElite) ELSE '' END AS Tr_Div"
'sSQL = sSQL + "		, CASE WHEN COALESCE(ju.MemberID,0)<>0 THEN MAX(ju.DivElite) ELSE '' END AS Ju_Div"


sSQL = sSQL + "	FROM usawsrank.EliteDates ed"
sSQL = sSQL + "	JOIN usawaterski.dbo.MemberShort m ON RIGHT(ed.MemberID,8)=PersonID"

sSQL = sSQL + "	LEFT JOIN"
sSQL = sSQL + "		( SELECT MemberID, CAST(MAX(QualThru) AS DATE) AS QualThru"
sSQL = sSQL + "			, MAX(DivElite) AS DivElite, MAX(DivOrig) AS DivOrig"
sSQL = sSQL + "			FROM usawsrank.EliteDates"
sSQL = sSQL + "			WHERE Event='S' AND SkiYearID=1"
sSQL = sSQL + "				AND LEFT(DivElite,1)='O'" 
sSQL = sSQL + "			GROUP BY MemberID) sl"
sSQL = sSQL + "	ON sl.MemberID=ed.MemberID"

sSQL = sSQL + "	LEFT JOIN"
sSQL = sSQL + "		( SELECT MemberID, CAST(MAX(QualThru) AS DATE) AS QualThru, MAX(DivElite) AS DivElite, MAX(DivOrig) AS DivOrig"
sSQL = sSQL + "			FROM usawsrank.EliteDates"
sSQL = sSQL + "			WHERE Event='T' AND SkiYearID=1"
sSQL = sSQL + "				AND LEFT(DivElite,1)='O'"
sSQL = sSQL + "			GROUP BY MemberID) tr"
sSQL = sSQL + "	ON tr.MemberID=ed.MemberID"

sSQL = sSQL + "	LEFT JOIN"
sSQL = sSQL + "		( SELECT MemberID, CAST(MAX(QualThru) AS DATE) AS QualThru, MAX(DivElite) AS DivElite, MAX(DivOrig) AS DivOrig"
sSQL = sSQL + "			FROM usawsrank.EliteDates"
sSQL = sSQL + "			WHERE Event='J' AND SkiYearID=1" 
sSQL = sSQL + "				AND LEFT(DivElite,1)='O'"
sSQL = sSQL + "			GROUP BY MemberID) ju"
sSQL = sSQL + "	ON ju.MemberID=ed.MemberID"

sSQL = sSQL + "	LEFT JOIN"
sSQL = sSQL + "		( SELECT MemberID, CAST(MAX(QualThru) AS DATE) AS QualThru"
sSQL = sSQL + "			, MAX(DivElite) AS DivElite, MAX(DivOrig) AS DivOrig"
sSQL = sSQL + "			FROM usawsrank.EliteDates"
sSQL = sSQL + "			WHERE Event='S' AND SkiYearID=1"
sSQL = sSQL + "				AND LEFT(DivElite,1)='M'" 
sSQL = sSQL + "			GROUP BY MemberID) slm"
sSQL = sSQL + "	ON slm.MemberID=ed.MemberID"

sSQL = sSQL + "	LEFT JOIN"
sSQL = sSQL + "		( SELECT MemberID, CAST(MAX(QualThru) AS DATE) AS QualThru, MAX(DivElite) AS DivElite, MAX(DivOrig) AS DivOrig"
sSQL = sSQL + "			FROM usawsrank.EliteDates"
sSQL = sSQL + "			WHERE Event='T' AND SkiYearID=1"
sSQL = sSQL + "				AND LEFT(DivElite,1)='M'"
sSQL = sSQL + "			GROUP BY MemberID) trm"
sSQL = sSQL + "	ON trm.MemberID=ed.MemberID"

sSQL = sSQL + "	LEFT JOIN"
sSQL = sSQL + "		( SELECT MemberID, CAST(MAX(QualThru) AS DATE) AS QualThru, MAX(DivElite) AS DivElite, MAX(DivOrig) AS DivOrig"
sSQL = sSQL + "			FROM usawsrank.EliteDates"
sSQL = sSQL + "			WHERE Event='J' AND SkiYearID=1" 
sSQL = sSQL + "				AND LEFT(DivElite,1)='M'"
sSQL = sSQL + "			GROUP BY MemberID) jum"
sSQL = sSQL + "	ON jum.MemberID=ed.MemberID"



sSQL = sSQL + "	WHERE skiYearID=1"
IF EliteSelected<>"" THEN sSQL = sSQL + "	AND LEFT(ed.DivElite,1)='"&EliteSelected&"'"
sSQL = sSQL + "		AND FederationCode='USA'"

sSQL = sSQL + "	GROUP BY ed.MemberID, sl.MemberID, tr.MemberID, ju.MemberID, slm.MemberID, trm.MemberID, jum.MemberID"

sSQL = sSQL + "	ORDER BY ed.lastName, ed.FirstName"
  

rs.open sSQL, SConnectionToTRATable


END SUB





' ----------------------------
  SUB DisplayPersonalBestList
' ----------------------------  

' -- Personal Best Sticker Requests --
sSQL = "SELECT FirstName, LastName, Address1, City, State, Zip" 
sSQL = sSQL + "		 , MemberID, TourID, Event, Div, Score, Created_Date" 
sSQL = sSQL + "		  FROM "&PBStickerTableName&" pb"
sSQL = sSQL + "		  JOIN " &MemberTableName& " m ON m.PersonIDWithCheckDigit=pb.MemberID"
sSQL = sSQL + "		  WHERE created_date>='" &StartDateSelected& "' AND created_date<='" &EndDateSelected& "'"

SortSequence = "membcreatedate"
IF SortSequence = "membcreatedate" THEN 
		sSQL = sSQL + "		 ORDER BY Created_Date DESC, MemberID"
ELSE 
		sSQL = sSQL + "		 ORDER BY MemberID, Created_Date DESC"	
END IF

rs.open sSQL, SConnectionToTRATable

END SUB




		
' -------------------------
   SUB OLR_IPN_Analysis
' -------------------------

sSQL = "SELECT TourID"
sSQL = sSQL + ", CASE WHEN PayStatus='N' THEN 'N-IPN Only'"
sSQL = sSQL + " WHEN PayStatus='O' THEN '0-OLR Receipt'"
sSQL = sSQL + " WHEN PayStatus='C' THEN 'C-OLR Other'"
sSQL = sSQL + " ELSE 'No Payment' END AS Paystatus" 
sSQL = sSQL + " , Message, Result"
sSQL = sSQL + " , Count(Result) AS CountResult, Count(PayStatus) AS CountPayStatus"
sSQL = sSQL + " FROM usawsrank.RegPaymentLog"
sSQL = sSQL + " WHERE OrderNo>689410"
sSQL = sSQL + " GROUP BY Result, PayStatus, TourID, Message"
sSQL = sSQL + " ORDER BY TourID, Result, PayStatus, Message "

'response.write(sSQL)
'response.end

rs.open sSQL, SConnectionToTRATable

END SUB


' ---------------------------------
   SUB OLR_IPN_Analysis_Summary
' ---------------------------------

sSQL = "SELECT CASE WHEN PayStatus='N' THEN 'N-IPN Only'"
sSQL = sSQL + " WHEN PayStatus='O' THEN '0-OLR Receipt'"
sSQL = sSQL + " WHEN PayStatus='C' THEN 'C-OLR Other'"
sSQL = sSQL + " ELSE 'No Payment' END AS Paystatus" 
sSQL = sSQL + " , Message, Result"
sSQL = sSQL + " , Count(Result) AS CountResult, Count(PayStatus) AS CountPayStatus"
sSQL = sSQL + " FROM usawsrank.RegPaymentLog"
sSQL = sSQL + " WHERE OrderNo>689410"
sSQL = sSQL + " GROUP BY Result, PayStatus, Message"
sSQL = sSQL + " ORDER BY Result, PayStatus, Message "

'response.write(sSQL)
'response.end

rs.open sSQL, SConnectionToTRATable

END SUB




' ----------------------
   SUB GetBioInfo
' ----------------------

'sTourID = "14C999"

sSQL = sSQL + " SELECT LastUpdate, LastName, FirstName "
sSQL = sSQL + " 	, m.City, m.State"
'sSQL = sSQL + " 	, DateDiff(yy,BirthDate, GETDATE()) AS Age	"
sSQL = sSQL + " 	, SLDiv, TRDiv, JUDiv"
sSQL = sSQL + " 	, (HgtFeet+'-'+HgtInch) AS Height"
sSQL = sSQL + " 	, SkiSinceAge, CompSinceAge, MembSinceAge"
sSQL = sSQL + " 	, Club, School, Occup AS Occupation, Career"
sSQL = sSQL + " 	, Hobby, Paper, Sponsors"
sSQL = sSQL + " 	, BestSlal AS BestSlalom, BestTrick AS BestTrick, BestJump AS BestJump"
sSQL = sSQL + " 	, Records, Titles"
sSQL = sSQL + " 	, Fav_Sports, Fav_Boat, Fav_Slalom, Fav_Jump, Fav_Trick"
sSQL = sSQL + " 	, Accomplish AS Biggest_Accomplishment, Mentors"
sSQL = sSQL + " 	FROM "&BioTableName&" b"
sSQL = sSQL + " 	JOIN "
sSQL = sSQL + " 		(SELECT MemberID"
sSQL = sSQL + " 			FROM "&RegGenTableName
sSQL = sSQL + " 			WHERE LEFT(TourID,6)='"&sTourID&"') rg"
sSQL = sSQL + " 	ON b.MemberID=rg.MemberID"
sSQL = sSQL + " 	LEFT JOIN"
sSQL = sSQL + " 		( SELECT PersonID, FirstName, LastName, City, State, BirthDate"
sSQL = sSQL + " 				FROM "&MemberShortTableName& ") m"
sSQL = sSQL + " 	ON RIGHT(rg.MemberID,8)=m.PersonID"
sSQL = sSQL + " 	LEFT JOIN"
sSQL = sSQL + " 		( SELECT MemberID, Div AS SLDiv "
sSQL = sSQL + " 				FROM "&RegDetailTableName
sSQL = sSQL + " 				WHERE TourID='"&sTourID&"' AND Event='S') RE1"
sSQL = sSQL + " 	ON RIGHT(RE1.MemberID,8)=m.PersonID"
sSQL = sSQL + " 	LEFT JOIN"
sSQL = sSQL + " 		( SELECT MemberID, Div AS TRDiv "
sSQL = sSQL + " 				FROM "&RegDetailTableName
sSQL = sSQL + " 				WHERE TourID='"&sTourID&"' AND Event='T') RE2"
sSQL = sSQL + " 	ON RIGHT(RE2.MemberID,8)=m.PersonID"
sSQL = sSQL + " 	LEFT JOIN"
sSQL = sSQL + " 		( SELECT MemberID, Div AS JUDiv "
sSQL = sSQL + " 				FROM "&RegDetailTableName
sSQL = sSQL + " 				WHERE TourID='"&sTourID&"' AND Event='J') RE3"
sSQL = sSQL + " 	ON RIGHT(RE3.MemberID,8)=m.PersonID"
sSQL = sSQL + " 	ORDER BY LastName, FirstName"

rs.open sSQL, SConnectionToTRATable

END SUB





' ----------------------
   SUB GetBioInfo_Evt
' ----------------------

'sTourID = "14C999"

sSQL = sSQL + " SELECT RE.Div, RE.Event, LastUpdate, LastName, FirstName "
sSQL = sSQL + " 	, m.City, m.State"
sSQL = sSQL + " 	, (HgtFeet+'-'+HgtInch) AS Height"
sSQL = sSQL + " 	, SkiSinceAge, CompSinceAge, MembSinceAge"
sSQL = sSQL + " 	, Club, School, Occup AS Occupation, Career"
sSQL = sSQL + " 	, Hobby, Paper, Sponsors"
sSQL = sSQL + " 	, BestSlal AS BestSlalom, BestTrick AS BestTrick, BestJump AS BestJump"
sSQL = sSQL + " 	, Records, Titles"
sSQL = sSQL + " 	, Fav_Sports, Fav_Boat, Fav_Slalom, Fav_Jump, Fav_Trick"
sSQL = sSQL + " 	, Accomplish AS Biggest_Accomplishment, Mentors"

sSQL = sSQL + " 	FROM "&RegDetailTableName&" RE"

sSQL = sSQL + " 	LEFT JOIN"
sSQL = sSQL + " 		( SELECT *"
sSQL = sSQL + " 				FROM "&BioTableName&") B"
sSQL = sSQL + " 	ON B.MemberID=RE.MemberID "

sSQL = sSQL + " 	JOIN "
sSQL = sSQL + " 		(SELECT MemberID, TourID"
sSQL = sSQL + " 			FROM "&RegGenTableName
sSQL = sSQL + " 			WHERE LEFT(TourID,6)='"&sTourID&"') RG"
sSQL = sSQL + " 	ON RG.MemberID=RE.MemberID AND RG.TourID=RE.TourID"

sSQL = sSQL + " 	LEFT JOIN"
sSQL = sSQL + " 		( SELECT PersonID, FirstName, LastName, City, State, BirthDate"
sSQL = sSQL + " 				FROM "&MemberShortTableName& ") M"
sSQL = sSQL + " 	ON RIGHT(RE.MemberID,8)=M.PersonID"

sSQL = sSQL + " 	LEFT JOIN" 
sSQL = sSQL + " 	(SELECT MemberID, Event, Div, RankScore" 
sSQL = sSQL + " 	FROM usawsrank.Rankings WHERE SkiYearID='1') RK" 
sSQL = sSQL + " 	ON RK.MemberID=RE.MemberID AND RK.Event=RE.Event AND RK.Div=RE.Div"

sSQL = sSQL + " 	WHERE RE.TourID='"&sTourID&"'"

sSQL = sSQL + " 	ORDER BY RE.Div, RE.Event, RankScore"



'response.write(sSQL)
'response.end

rs.open sSQL, SConnectionToTRATable


END SUB












' -----------------
  SUB LeagueList
' -----------------

sSQL = "SELECT LeagueID, LeagueName"
sSQL = sSQL + " FROM "&LeagueTableName
sSQL = sSQL + " WHERE QualifyTour<>''"
sSQL = sSQL + " AND Status<>'X'"
sSQL = sSQL + " AND RIGHT(LeagueID,4)="
sSQL = sSQL + " 		(SELECT MAX(RIGHT(LeagueID,4)) FROM "&LeagueTableName&")"
sSQL = sSQL + " ORDER BY LeagueName DESC" 

response.write(sSQL)
'response.end

rs.open sSQL, SConnectionToTRATable

END SUB



' -----------------
  SUB SkierList
' -----------------

Set rs=Server.CreateObject("ADODB.recordset")

'response.write("sLeagueSelected = "&sLeagueSelected)

sSQL = " SELECT QualifyTour, LeagueID"
sSQL = sSQL + " FROM "&LeagueTableName
sSQL = sSQL + " WHERE LeagueID = '"&sLeagueSelected&"'"
rs.open sSQL, SConnectionToTRATable
sTourID=""
IF NOT rs.eof THEN
		sTourID=rs("QualifyTour")
END IF
'response.write(sSQL)
'response.end


Set rs=Server.CreateObject("ADODB.recordset")
sSQL = " SELECT Email, First, Last, State"
sSQL = sSQL + " , CASE WHEN Region='1' THEN 'SC' WHEN Region='2' THEN 'MW' WHEN Region='3' THEN 'WE' WHEN Region='4' THEN 'SO' WHEN Region='5' THEN 'EA' END AS Region" 
sSQL = sSQL + " FROM"
sSQL = sSQL + " ( "
sSQL = sSQL + " SELECT DISTINCT MemberID, First, Last, Email, Region, State"
sSQL = sSQL + " FROM "&RegGenTableName&" RG"

sSQL = sSQL + " JOIN"
sSQL = sSQL + " ( SELECT PersonID, FirstName AS First, LastName AS Last, Address1, City, State, Zip, Email, Region"
sSQL = sSQL + " 		FROM "&MemberLiveTableName&" ) MT"
sSQL = sSQL + " ON CAST(RIGHT(RG.MemberID,8) AS INT)=MT.PersonID"

sSQL = sSQL + " WHERE TourID='"&sTourID&"'"
sSQL = sSQL + "  AND Email IS NOT NULL"
IF TRIM(RegionSelected)<>"" THEN
	sSQL = sSQL + "  AND MT.Region='"&RegionSelected&"'"
END IF
sSQL = sSQL + " ) AS A"
sSQL = sSQL + " ORDER BY Last, First" 

' response.write(sSQL)
' response.end

rs.open sSQL, SConnectionToTRATable

END SUB




' --------------------
  SUB RidesCountByYear
' --------------------


sSQL = "Select Region, Year, count(*) as Skiers, Sum(Rides) as Rides, CAST(CAST(Sum(Rides) AS REAL)/CAST(count(*) AS REAL) AS DECIMAL(7,2)) AS [Rides per Skier]"
sSQL = sSQL + " FROM"
sSQL = sSQL + "    (Select MemberID, substring(TourID,3,1) as Region, datepart(yyyy,enddate) as Year, count(*) as Rides"
sSQL = sSQL + "    From USAWSRank.Scores"
sSQL = sSQL + "    where enddate >= '2006-01-01'"
sSQL = sSQL + "    and substring(TourID,7,1) <> 'A'"
sSQL = sSQL + "    group by MemberID, substring(TourID,3,1), datepart(yyyy,enddate)) InnerSel"
sSQL = sSQL + " Group by Region, Year" 
sSQL = sSQL + " Order by Region, Year;"

'CAST(CAST((TourCnt-1)*5 AS DECIMAL(7,2)) AS Char(10))

'response.write(sSQL)
'response.end
rs.open sSQL, SConnectionToTRATable

END SUB


' ---------------------------
  SUB ClassNorFToursByYear
' ---------------------------

sSQL = "SELECT LEFT(TourID,2) AS Year, COUNT(TourID) AS [Grassroots Class N or F]"  
sSQL = sSQL + "		FROM "&RawScoresTableName
sSQL = sSQL + "			WHERE Class IN ('N', 'F')"
sSQL = sSQL + "		GROUP BY LEFT(TourID,2)"
sSQL = sSQL + "		ORDER BY LEFT(TourID,2) DESC"
rs.open sSQL, SConnectionToTRATable



END SUB



' ----------------------
  SUB OLREntriesByYear
' ----------------------

sSQL =  "SELECT LEFT(TourID,2) AS Year, Count(TourID) AS [Skier Count]" 
sSQL = sSQL + "	FROM "&RegGenTableName
sSQL = sSQL + "	GROUP BY LEFT(TourID,2)"  
sSQL = sSQL + "	ORDER BY LEFT(TourID,2)  DESC"
rs.open sSQL, SConnectionToTRATable

END SUB


' -----------------------
SUB OLRandALLTourStats
' -----------------------

sSQL = "SELECT DISTINCT LEFT(TourID,2) AS [Ski Year]" 
sSQL = sSQL + ",	[Tour Count OLR] AS [Tour Count<br>OLR], COALESCE([Tour Count All],0) AS [Tour Count<br>All]" 
sSQL = sSQL + ",	[Unique Members OLR] AS [Members Unique<br>OLR],	[Unique Members All] AS [Members Unique<br>All]"
sSQL = sSQL + ",	[Entry Count OLR] AS [Entry Count<br>OLR], COALESCE([Entry Count All],0) AS [Entry Count<br>All]"
sSQL = sSQL + ",	[Event Count OLR] AS [Event Count<br>OLR], COALESCE([Event Count All],0) AS [Event Count<br>All]"
sSQL = sSQL + " FROM "&RegGenTableName&" Y"
	
sSQL = sSQL + "	LEFT JOIN"
sSQL = sSQL + "	( 	SELECT LEFT(TourID,2) AS Year, Count(Distinct TourID) AS [Tour Count OLR]"
sSQL = sSQL + "			FROM "&RegGenTableName
sSQL = sSQL + "			GROUP BY LEFT(TourID,2) ) TourCntOLR"
sSQL = sSQL + "	ON LEFT(Y.TourID,2) = TourCntOLR.Year"

sSQL = sSQL + "	LEFT JOIN"
sSQL = sSQL + "	( 	SELECT LEFT(TourID,2) AS Year, Count(Distinct TourID) AS [Tour Count All]"
sSQL = sSQL + "			FROM "&RawScoresTableName
sSQL = sSQL + "			GROUP BY LEFT(TourID,2) ) TourCntAll"
sSQL = sSQL + "	ON LEFT(Y.TourID,2) = TourCntAll.Year "

sSQL = sSQL + "	LEFT JOIN"
sSQL = sSQL + "	( 	SELECT LEFT(TourID,2) AS Year, Count(Distinct MemberID) AS [Unique Members All]"
sSQL = sSQL + "			FROM "&RawScoresTableName
sSQL = sSQL + "			GROUP BY LEFT(TourID,2) ) EntCntAllUni"
sSQL = sSQL + "	ON LEFT(Y.TourID,2) = EntCntAllUni.Year "

sSQL = sSQL + "	LEFT JOIN"
sSQL = sSQL + "	( SELECT LEFT(TourID,2) AS Year, Count(DISTINCT MemberID) AS [Unique Members OLR] "
sSQL = sSQL + "		FROM "&RegGenTableName
sSQL = sSQL + "			GROUP BY LEFT(TourID,2)  ) EntCntOLRUni"
sSQL = sSQL + "	ON LEFT(Y.TourID,2) = EntCntOLRUni.Year"

sSQL = sSQL + "	LEFT JOIN"
sSQL = sSQL + "	( SELECT LEFT(TourID,2) AS Year, Count(TourID) AS [Entry Count OLR] "
sSQL = sSQL + "		FROM "&RegGenTableName
sSQL = sSQL + "			GROUP BY LEFT(TourID,2)  ) EntCntOLR "
sSQL = sSQL + "	ON LEFT(Y.TourID,2) = EntCntOLR.Year"

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "	( SELECT LEFT(TourID,2) AS Year, Count(DISTINCT MemberID+TourID) AS [Entry Count All] "
sSQL = sSQL + "		FROM "&RawScoresTableName
sSQL = sSQL + "			GROUP BY LEFT(TourID,2)  ) EntCntAll"
sSQL = sSQL + "	ON LEFT(Y.TourID,2) = EntCntAll.Year"

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "	( SELECT LEFT(TourID,2) AS Year, Count(MemberID) AS [Event Count OLR]"
sSQL = sSQL + "		FROM "&RegDetailTableName
sSQL = sSQL + "			GROUP BY LEFT(TourID,2)  ) EvtCntOLR "
sSQL = sSQL + "	ON LEFT(Y.TourID,2) = EvtCntOLR.Year "

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "	( SELECT LEFT(TourID,2) AS Year, Count(TourID) AS [Event Count All] "
sSQL = sSQL + "		FROM "&RawScoresTableName
sSQL = sSQL + "			GROUP BY LEFT(TourID,2)  ) EvtCntAll"
sSQL = sSQL + "	ON LEFT(Y.TourID,2) = EvtCntAll.Year"

sSQL = sSQL + "	ORDER BY LEFT(TourID,2)"
'response.write(sSQL)
'response.end
rs.open sSQL, SConnectionToTRATable


END SUB

' ----------------------
  SUB OLREntriesByYear_2
' ----------------------

sSQL =  "SELECT LEFT(SC.TourID,2) AS Year, Count(SC.TourID) AS [Skier Count]" 
sSQL = sSQL + "	        FROM "&RawScoresTableName&" AS SC"
sSQL = sSQL +  "     JOIN"
sSQL = sSQL +  "     (SELECT LEFT(TourID,2) AS Year, Count(TourID) AS [Skier Count]" 
sSQL = sSQL + "	        FROM "&RegGenTableName&") AS Reg"
sSQL = sSQL + "	     ON LEFT(SC.TourID,6)=LEFT(Reg.TourID,6)"   
sSQL = sSQL + "	GROUP BY LEFT(SC.TourID,2)"  
sSQL = sSQL + "	ORDER BY LEFT(SC.TourID,2)  DESC"

response.write(sSQL)
response.end
rs.open sSQL, SConnectionToTRATable

END SUB




' ----------------------
  SUB OLREntriesByYear_incomplete
' ----------------------

' --- NOT COMPLETE

sSQL =  "SELECT LEFT(TourID,2) AS Year, Count(TourID) AS [Skier Count], [SkierF Count]" 
sSQL = sSQL + "	FROM "&RegGenTableName&" AS RG1"

sSQL = sSQL + "	   JOIN"
sSQL = sSQL + "	     (SELECT LEFT(TourID,2) AS Year, Count(TourID) AS [SkierF Count]"  
sSQL = sSQL + "		  FROM "&RegGenTableName&" AS T2"
sSQL = sSQL + "		 JOIN"
sSQL = sSQL + "		     (SELECT" 	
sSQL = sSQL + "		 ON" 
sSQL = sSQL + "	     WHERE Class IN ('F','N')) AS RG2"  
sSQL = sSQL + "	   ON RG2.TourID=RG1.TourID"  

sSQL = sSQL + "	GROUP BY LEFT(TourID,2)"  
sSQL = sSQL + "	ORDER BY LEFT(TourID,2)  DESC"
rs.open sSQL, SConnectionToTRATable

END SUB




' ----------------------
  SUB OLRToursByYear
' ----------------------
 
sSQL =  "SELECT LEFT(TournAppID,2) AS Year, Count(TournAppID) AS [Tour Count]" 
sSQL = sSQL + "	FROM sanctions.dbo.Registration"
sSQL = sSQL + "	WHERE PayPalOK='1'"
sSQL = sSQL + "	GROUP BY LEFT(TournAppID,2)"
sSQL = sSQL + "	ORDER BY LEFT(TournAppID,2) DESC"
rs.open sSQL, SConnectionToTRATable

END SUB



' ----------------------
  SUB LeagueQualSummary
' ----------------------

' CAST(CAST(SumScore/NumScores AS DECIMAL(7,2)) AS Char(10)) AS [Avg Score] 

sSQL =  "SELECT Event, Div, CAST(CAST(COA AS DECIMAL(7,2)) AS Char(10)) AS [COA Avg] " 
sSQL = sSQL + "	FROM "&LeagueQfyTableName
sSQL = sSQL + "	WHERE LeagueID='"&sLeagueSelected&"'"
sSQL = sSQL + "	ORDER BY Div, Event"

'response.write(sSQL)
'response.end

rs.open sSQL, SConnectionToTRATable

'response.write("<br>EOF1=")
'response.write(rs.eof)
'response.end

END SUB



' -------------------
  SUB NationalEntries
' -------------------

'response.write("<br>Line 1559<br>")

SET rs=Server.CreateObject("ADODB.recordset")
sSQL =  "SELECT SkiYear, SkiYearName FROM "&SkiYearTableName&" WHERE DefaultYear='1'"
rs.open sSQL, SConnectionToTRATable, 3, 3
sSkiYear=rs("SkiYear")
ThisSkiYear_2Digit=RIGHT(rs("SkiYearName"),2)
rs.close

sSQL =  "SELECT tr.TournAppID, ts.SptsGrpID FROM "&TRegSetupTableName&" tr"
sSQL = sSQL + " JOIN "&SanctionTableName&" ts ON ts.TournAppID=tr.TournAppID" 
sSQL = sSQL + " WHERE LEFT(tr.TournAppID,2) = '"&ThisSkiYear_2Digit&"' AND RIGHT(LEFT(tr.TournAppID,6),3) = '999' AND ts.SptsGrpID='AWS'"
' sSQL = sSQL + " WHERE LEFT(TournAppID,2) = '"&ThisSkiYear_2Digit&"' AND RIGHT(LEFT(TournAppID,6),3) = '999'"

rs.open sSQL, SConnectionToTRATable, 3, 3
ThisTournAppID=rs("TournAppID")
rs.close




LastSkiYear_2Digit = CInt(ThisSkiYear_2Digit)-CInt(1)
sSQL =  "SELECT tr.TournAppID, ts.SptsGrpID FROM "&TRegSetupTableName&" tr"
sSQL = sSQL + " JOIN "&SanctionTableName&" ts ON ts.TournAppID=tr.TournAppID" 
sSQL = sSQL + " WHERE LEFT(tr.TournAppID,2) = '"&LastSkiYear_2Digit&"' AND RIGHT(LEFT(tr.TournAppID,6),3) = '999' AND ts.SptsGrpID='AWS'"

'sSQL = sSQL + " WHERE LEFT(TournAppID,2) = '"&LastSkiYear_2Digit&"' AND RIGHT(LEFT(TournAppID,6),3) = '999'"
rs.open sSQL, SConnectionToTRATable, 3, 3
LastTournAppID=rs("TournAppID")

'response.write("<br>Line 1589<br>")
'response.write(sSQL)
'response.write("<br>")
'response.end



sSQL = "SELECT TDateS AS ThisStartDate FROM "&SanctionTableName
sSQL = sSQL + " WHERE LEFT(TournAppID,6)='"&ThisTournAppID&"'"
SET rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable, 3, 3
ThisStartDate = rs("ThisStartDate")

sSQL = "SELECT TDateS AS LastStartDate FROM "&SanctionTableName
sSQL = sSQL + " WHERE LEFT(TournAppID,6)='"&LastTournAppID&"'"
SET rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable, 3, 3
LastStartDate = rs("LastStartDate")

DiffBetweenStartDates = DateDiff("d",LastStartDate+366, ThisStartDate)

'response.write("<br>DiffBetweenStartDates = "&DiffBetweenStartDates)
'response.write("<br>")
rs.close





sSQL = "SELECT RC.RegisterDate"
sSQL = sSQL + ", Coalesce(RC1.CntThis,0) AS CntThis"
sSQL = sSQL + ", Coalesce(RC2.CntLast,0) AS CntLast" 
sSQL = sSQL + " FROM "
sSQL = sSQL + " ( "


sSQL = sSQL + " SELECT DISTINCT CAST(RegisterDate AS DATE) AS RegisterDate FROM"

sSQL = sSQL + " (SELECT CAST(RegisterDate AS DATE) AS RegisterDate" 
sSQL = sSQL + " FROM "&RegGenTableName&" WHERE LEFT(TourID,6)='"&ThisTournAppID&"'"
sSQL = sSQL + " UNION"

sSQL = sSQL + " SELECT DateAdd(d,(366+"&DiffBetweenStartDates&"),RegisterDate) AS RegisterDate"
sSQL = sSQL + "    FROM "&RegGenTableName&" WHERE LEFT(TourID,6)='"&LastTournAppID&"' ) AS A"

sSQL = sSQL + " ) AS RC" 

sSQL = sSQL + "	LEFT JOIN"
sSQL = sSQL + "	(SELECT CAST(RegisterDate AS DATE) AS RegisterDate, Count(MemberID) AS CntThis"
sSQL = sSQL + "		FROM "&RegGenTableName
sSQL = sSQL + " WHERE LEFT(TourID,6)='"&ThisTournAppID&"'"
sSQL = sSQL + "	 GROUP BY CAST(RegisterDate AS DATE) ) AS RC1" 
sSQL = sSQL + "	ON RC.RegisterDate=RC1.RegisterDate"

sSQL = sSQL + "	LEFT JOIN"
sSQL = sSQL + "	(SELECT CAST(RegisterDate AS DATE) AS RegisterDate, Count(MemberID) AS CntLast"
sSQL = sSQL + "		FROM "&RegGenTableName
sSQL = sSQL + "	WHERE LEFT(TourID,6)='"&LastTournAppID&"'"
sSQL = sSQL + "		GROUP BY CAST(RegisterDate AS DATE) ) AS RC2 "
sSQL = sSQL + "	ON RC.RegisterDate=DateAdd(d,(366+"&DiffBetweenStartDates&"),RC2.RegisterDate)"
'"&DiffBetweenStartDates&"
sSQL = sSQL + " ORDER BY RC.RegisterDate"

'response.write(sSQL)
'response.end
rs.open sSQL, SConnectionToTRATable


END SUB




' --------------------
  SUB NationalTotals
' --------------------

rs.Close
sSQL = "	SELECT Count(MemberID) AS TotalCntLastYear"
sSQL = sSQL + "			FROM "&RegGenTableName
sSQL = sSQL + "		WHERE LEFT(TourID,6)='"&LastTournAppID&"'"
rs.open sSQL, SConnectionToTRATable
IF NOT rs.eof THEN sTotalCntLastYear =rs("TotalCntLastYear")
rs.Close


sSQL = "	SELECT Count(MemberID) AS TotalCntThisYear"
sSQL = sSQL + "			FROM "&RegGenTableName
sSQL = sSQL + "		WHERE LEFT(TourID,6)='"&ThisTournAppID&"'"
rs.open sSQL, SConnectionToTRATable
IF NOT rs.eof THEN sTotalCntThisYear =rs("TotalCntThisYear")
rs.Close


%>
<center>
<font size=2 color="white">Total Nationals Entries</font>
<br>
<font size=2 color="white">Last Year Total - <%=sTotalCntLastYear%></font>
<br>
<font size=2 color="white">This Year To Date - <%=sTotalCntThisYear%></font>
</center>
<br><br><br>
<%



END SUB



' ---------------------
  SUB Refunds_OLD
' ---------------------


sSQL = "SELECT RG.MemberID, MT.FirstName, MT.LastName, MT.Address1, MT.City, MT.State, MT.Zip, REC.EnterCount AS [Events<br>Entered]"
sSQL = sSQL + ", ST1.SkiedCount AS [Events<br>Skied], CAST(COALESCE(Payments,0) AS money) AS Payments"
sSQL = sSQL + "			FROM "&RegGenTableName&" AS RG"

sSQL = sSQL + "	  LEFT JOIN "&MemberShortTableName&" AS MT ON CAST(RIGHT(RG.MemberID,8) AS INT)=MT.PersonID"

sSQL = sSQL + "		JOIN"
sSQL = sSQL + "			(SELECT MemberID, TourID, Count(MemberID) AS EnterCount FROM "&RegDetailTableName
sSQL = sSQL + "				GROUP BY MemberID, TourID) AS REC"
sSQL = sSQL + "			ON REC.MemberID=RG.MemberID AND LEFT(REC.TourID,6)=LEFT(RG.TourID,6)"

sSQL = sSQL + "		LEFT JOIN"
sSQL = sSQL + "			(SELECT MemberID, TourID, Count(MemberID) AS SkiedCount FROM "&RawScoresTableName
sSQL = sSQL + "				GROUP BY MemberID, TourID) AS ST1"
sSQL = sSQL + "			ON ST1.MemberID=RG.MemberID AND LEFT(ST1.TourID,6)=LEFT(RG.TourID,6)"

sSQL = sSQL + "			LEFT JOIN (SELECT MemberID, SUM(Amount) AS Payments FROM "&RegPaymentTableName
sSQL = sSQL + "				WHERE  LEFT(TourID,6) = '13S999' AND Result = '0'"
sSQL = sSQL + "				GROUP BY MemberID) AS TP"
sSQL = sSQL + "				ON TP.MemberID = RG.MemberID"
	
sSQL = sSQL + "			WHERE LEFT(RG.TourID,6)='13S999'"
sSQL = sSQL + "				AND ( (ST1.SkiedCount <> REC.EnterCount AND ST1.SkiedCount<>2)"
sSQL = sSQL + "				OR ST1.SkiedCount IS NULL)"

sSQL = sSQL + "				AND TP.Payments<>'0' AND Payments<>'355'"



sSQL = sSQL + "			ORDER BY RG.MemberID"

response.write(sSQL)
response.end
rs.open sSQL, SConnectionToTRATable


END SUB





' ---------------------
  SUB Refunds
' ---------------------


sSQL = " SELECT RP.MemberID, COALESCE(RE.EnterCount,0) AS [EventsEntered], COALESCE(ST.SkiedCount,0) AS [EventsSkied]"
sSQL = sSQL + ", MT.FirstName, MT.LastName, MT.Address1, MT.City, MT.State, MT.Zip"
sSQL = sSQL + ", CAST(RP.Payments AS decimal(6,2)) AS Payments "
sSQL = sSQL + " FROM"

sSQL = sSQL + " ( SELECT MemberID, TourID, SUM(Amount) AS Payments, Result"
sSQL = sSQL + "	FROM usawsrank.RegPaymentLog" 
sSQL = sSQL + "	GROUP BY MemberID, TourID, Result ) AS RP"	 

sSQL = sSQL + " JOIN"	 
sSQL = sSQL + "	(SELECT MemberID, TourID, Count(MemberID) AS EnterCount"	 
sSQL = sSQL + "	FROM usawsrank.RegisterEvents"	 
sSQL = sSQL + "		GROUP BY MemberID, TourID) AS RE"	 
sSQL = sSQL + " ON RE.MemberID=RP.MemberID AND LEFT(RE.TourID,6)=LEFT(RP.TourID,6)"	 

sSQL = sSQL + " JOIN"	 
sSQL = sSQL + "	(SELECT MemberID, TourID, TotalEntry AS Payments"
sSQL = sSQL + "	FROM usawsrank.RegisterGen_05042014) AS RG"	 
sSQL = sSQL + " ON RG.MemberID=RP.MemberID AND LEFT(RG.TourID,6)=LEFT(RP.TourID,6)"	 

sSQL = sSQL + " LEFT JOIN"	 
sSQL = sSQL + "	(SELECT MemberID, TourID, Count(MemberID) AS SkiedCount" 
sSQL = sSQL + "			FROM usawsrank.Scores"	 
sSQL = sSQL + "			GROUP BY MemberID, TourID) AS ST"	 
sSQL = sSQL + " ON ST.MemberID=RP.MemberID AND LEFT(ST.TourID,6)=LEFT(RP.TourID,6)"	 

sSQL = sSQL + " LEFT JOIN"	 
sSQL = sSQL + "	( SELECT TournAppID, EntryFeeFamily" 
sSQL = sSQL + "			FROM sanctions.dbo.Registration) SR"	 
sSQL = sSQL + " ON SR.TournAppID=LEFT(RP.TourID,6)"	 

sSQL = sSQL + "	  LEFT JOIN "&MemberShortTableName&" AS MT ON CAST(RIGHT(RG.MemberID,8) AS INT)=MT.PersonID"

sSQL = sSQL + " WHERE LEFT(RP.TourID,6) = '"&sTourID&"' AND Result = '0'" 
sSQL = sSQL + "	AND ( (ST.SkiedCount <> RE.EnterCount AND ST.SkiedCount<>2) OR ST.SkiedCount IS NULL )"	 
sSQL = sSQL + "	AND RG.Payments<>EntryFeeFamily"	 

sSQL = sSQL + "			ORDER BY RP.MemberID"

'response.write(sSQL)
'response.end
rs.open sSQL, SConnectionToTRATable


END SUB




' ---------------------
  SUB LOCContacts
' ---------------------

IF TRIM(SkiYearSelected)="" THEN

		sSQL = "SELECT SkiYearID FROM " &SkiYearTableName& " WHERE DefaultYear=1"
		rs.open sSQL, SConnectionToTRATable
 
		SkiYearSelected = rs("SkiYearID")
		rs.close
END IF


Dim sSQL

sSQL = "SELECT ts.TournAppID AS TourID, RIGHT(LEFT(Rs.TourID,3),1) AS Reg, TName AS Tournament, TCity AS City, TState AS ST, TDirName AS [Tour Director], TDirEmail AS [Tour Dir Email], TSponsor AS [Tour Sponsor]"  
sSQL = sSQL + "	FROM sanctions.dbo.Tschedul AS TS"
sSQL = sSQL + "	LEFT JOIN sanctions.dbo.Registration r ON r.TournAppID=ts.TournAppID"
sSQL = sSQL + "	JOIN " & SkiYearTableName & " sy ON RIGHT(sy.SkiYear,2)=LEFT(ts.TournAppID,2)"	
sSQL = sSQL + "	JOIN"
sSQL = sSQL + "		  (SELECT Distinct TourID"
sSQL = sSQL + "			FROM usawsrank.Scores"
sSQL = sSQL + "			WHERE Score IS NOT NULL" 
IF EventSelected<>"" THEN
	sSQL = sSQL + "				AND Event='"&EventSelected&"'"
END IF

sSQL = sSQL + ") AS RS"
sSQL = sSQL + "	ON LEFT(Rs.TourID, 6)=LEFT(TS.TournAppID, 6)"

sSQL = sSQL + " WHERE RIGHT(LEFT(Rs.TourID,3),1) <> 'U'"
sSQL = sSQL + " AND sy.SkiYearID = '"&SkiYearSelected&"'"		
IF TRIM(ClassTypeSelected)<>"" THEN
		IF ClassTypeSelected="GorF" THEN sSQL = sSQL + " AND (GREntryFee1<>0 OR SClassN>0 OR TClassN>0 OR JClassN>0)"					
		IF ClassTypeSelected="NoGorF" THEN sSQL = sSQL + " AND GREntryFee1=0 AND SClassN=0 AND TClassN=0 AND JClassN=0"						
END IF
sSQL = sSQL + " ORDER BY ts.TournAppID"

'response.write(sSQL)
'response.end


rs.open sSQL, SConnectionToTRATable


END SUB





' ---------------------
  SUB AWSEFDonorsByLOC
' ---------------------

Dim sSQL

sSQL = "SELECT TName AS [Tournament Name], SUM(AWSEFDonation) AS Donations, Left(Convert(char,TDateE,111),10) as [End Date], TDirName AS Contact, TDirAddress AS Address, TDirCity AS City, TDirState AS State, TDirZip AS Zip, TDirEmail AS Email"  
sSQL = sSQL + "		FROM "&RegGenTableName&" AS RG, "&SanctionTableName&" AS ST"
sSQL = sSQL + "	    JOIN "&MemberShortTableName&" AS MT ON CAST(RIGHT(RG.MemberID,8) AS INT)=MT.PersonID"

sSQL = sSQL + "			WHERE AWSEFDonation>0" 
sSQL = sSQL + "				AND MT.PersonIDWIthCheckDigit=RG.MemberID"
sSQL = sSQL + "				AND LEFT(ST.TournAppID,6)=LEFT(RG.TourID,6)"
sSQL = sSQL + "			GROUP BY TName, TDirName, TDirAddress, TDirCity, TDirState, TDirZip, TDateE, TDirEmail"  
sSQL = sSQL + "		ORDER BY TDateE"

rs.open sSQL, SConnectionToTRATable

END SUB


' ---------------------
  SUB AWSEFDonorList
' ---------------------

sSQL = "SELECT Left(Convert(char,RegisterDate,111),10) AS DonorDate, FirstName AS First, LastName AS Last, Address1, City, State, Zip, Email, Left(Convert(char,BirthDate,111),10) AS DOB, MemberID, CAST(AWSEFDonation AS money) AS Donation, TName AS [Tournament Name]" 
sSQL = sSQL + "		FROM "&RegGenTableName&" AS RG, "&SanctionTableName&" AS ST"
sSQL = sSQL + "	    JOIN "&MemberShortTableName&" AS MT ON CAST(RIGHT(RG.MemberID,8) AS INT)=MT.PersonID"
sSQL = sSQL + "		WHERE AWSEFDonation>0" 
sSQL = sSQL + "			AND MT.PersonIDWIthCheckDigit=RG.MemberID"
sSQL = sSQL + "			AND LEFT(ST.TournAppID,6)=LEFT(RG.TourID,6)"
sSQL = sSQL + "		ORDER BY RegisterDate, MemberID"

rs.open sSQL, SConnectionToTRATable

END SUB


' ---------------------
  SUB UserPWList
' ---------------------

sSQL = "SELECT LastName, FirstName, SptsGrpID AS SD, RegnID AS Reg, UserName AS [User Name], PWord AS Password, HQUser, SecLevel, Sanctions AS Sanc, Seeding AS Seed, Officials AS [Offic]" 
sSQL = sSQL + " FROM sanctions.dbo.users" 
sSQL = sSQL + " ORDER BY LastName"
rs.open sSQL, SConnectionToTRATable

END SUB



' ---------------------
  SUB GRRanking
' ---------------------


sSQL = " SELECT MT.FirstName, MT.LastName, ST.MemberID, ST.Event, ST.Div, ST.[# of Scores], ST.[Avg Score], ST.[Bonus Points], ST.Ranking"
sSQL = sSQL + " 	FROM "
sSQL = sSQL + " 		(SELECT MemberID, Event, Div, COUNT(Score) AS [# of Scores], SUM(Score)/COUNT(Score) AS [Avg Score], COUNT(Score)*5 AS [Bonus Points], (SUM(Score)/COUNT(Score) + COUNT(Score)*5) AS Ranking"
sSQL = sSQL + " 			FROM usawsrank.scores "
sSQL = sSQL + " 				WHERE Class IN ('F', 'N') AND LEFT(TourID,2)='08' AND LEFT(Div,1) NOT IN ('Y', 'X', 'C')  AND IsNull(score,1)<>1"
sSQL = sSQL + " 			GROUP BY MemberID, Event, Div) AS ST"

sSQL = sSQL + " 		LEFT JOIN "
sSQL = sSQL + " 			( SELECT FirstName, LastName, PersonIDwithCheckDigit"
sSQL = sSQL + " 				FROM usawaterski.dbo.Members) AS MT"
sSQL = sSQL + " 		ON MT.PersonIDwithCheckDigit=ST.MemberID"			

sSQL = sSQL + " 		ORDER BY ST.Event, ST.Div, ST.[Avg Score] DESC"
rs.open sSQL, SConnectionToTRATable

END SUB



' ------------------------------
   SUB LeagueDropBuild_07162010
' ------------------------------

' ------------   Builds Ski Year Drop Down list ----------------- 


set rsList=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT DISTINCT LeagueID, LeagueName FROM "&LeagueTableName&" WHERE SptsGrpID='"&Session("sSptsGrpID")&"' ORDER BY LeagueID"
rsList.open sSQL, SConnectionToTRATable

' response.write("<br>"&sSQL)
'response.write(rsList.eof)
'response.write("<br>sLeagueSelected="&sLeagueSelected)
 %>
<SELECT name='sLeagueSelected' style="width:8em"><%

  response.write("<option value ='ALL'")
  IF sLeagueSelected = "ALL" THEN response.write(" Selected")
  response.write(">ALL</option><br>")

  IF NOT rsList.eof THEN
	rsList.movefirst
	DO WHILE not rsList.eof
	  response.write(" <option value ="""&rsList("LeagueID")&""" ")
	  response.write(" <a title="""&rsList("LeagueName")&"""")

	  IF trim(rsList("LeagueID")) = sLeagueSelected THEN
	    response.write(" selected")
	  END IF

	  response.write(">")
	  response.write(rsList("LeagueID"))
	  response.write("</a></option><br>")
	  rsList.movenext
	LOOP
  END IF %>

</SELECT><%

rsList.close



END SUB




' ------------------------------
   SUB LeagueDropBuild_SelectFromAll
' ------------------------------

' ------------   Builds Ski Year Drop Down list ----------------- 

set rsList=Server.CreateObject("ADODB.recordset")

'ChooseSQL("SELECT DISTINCT LeagueID, LeagueName FROM "&LeagueTableName&" WHERE SptsGrpID='"&Session("sSptsGrpID")&"' ORDER BY LeagueID")

ChooseSQL("SELECT LeagueID, LeagueName FROM "&LeagueTableName&" WHERE QualifyTour<>'' AND Status<>'X' AND RIGHT(LeagueID,4)>=(SELECT MAX(RIGHT(LeagueID,4))-1 FROM "&LeagueTableName&") ORDER BY LeagueName DESC") 



'response.write("<br>EOF2=")
'response.write(rsList.eof)
'response.write("<br>sLeagueSelected="&sLeagueSelected)
 %>
<SELECT name='sLeagueSelected' style="width:8em"><%

  response.write("<option value ='Select'")
  IF sLeagueSelected = "Select" THEN response.write(" Selected")
  response.write(">Select</option><br>")

  IF NOT rsList.eof THEN
	rsList.movefirst
	DO WHILE not rsList.eof
	  response.write(" <option value ="""&rsList("LeagueID")&""" ")
	  response.write(" <a title="""&rsList("LeagueName")&"""")

	  IF trim(rsList("LeagueID")) = sLeagueSelected THEN
	    response.write(" selected")
	  END IF

	  response.write(">")
	  response.write(rsList("LeagueID"))
	  response.write("</a></option><br>")
	  rsList.movenext
	LOOP
  END IF %>

</SELECT><%

rsList.close

END SUB







' --------------------
   SUB ChooseSQL(sSQL)
' --------------------

'response.write("In ChooseSQL "&sSQL)

rsList.open sSQL, sConnectionToTRATable, 3, 3

END SUB


%>