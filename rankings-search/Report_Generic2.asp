<!--#include file="settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->
<!--#include virtual="/rankings/tools_Definitions.asp"-->

<%
' --- Last update 2-6-2013 ---


DefineTRAStyles

Dim ThisFileName, EventSelected, EventName, sPriorYear, process, MainImage, AdminMenuLevel
Dim sLeagueSelected
Dim rsList, Misc1, Misc2, Misc3

Dim ThisTournAppID, LastTournAppID, ThisStartDate, LastStartDate, DiffBetweenStartDates

TourTableWidth=675
TabWidth = 1000  	' --- Used in case where report does not have specific parameters

ThisFileName="Report_Generic2.asp"
AdminMenuLevel=Session("AdminMenuLevel")



' --- Process Control ---
process=TRIM(LCASE(request("process")))
sAction=Request("Action")
IF sAction="Return to Menu" THEN 
	process="return"
END IF


' --- Event and League ---
EventSelected=TRIM(Request("EventSelected"))
IF EventSelected="" THEN EventSelected="J"

sLeagueSelected=TRIM(Request("sLeagueSelected"))
'IF sLeagueSelected="" THEN sLeagueSelected="NATL2010"

SELECT CASE EventSelected
	CASE "S"
		EventName="Slalom"
	CASE "T"
		EventName="Trick"
	CASE "J"
		EventName="Jump"
END SELECT


Misc1=Request("Misc1")
Misc2=Request("Misc2")
Misc3=Request("Misc3")



' --- Defines the image to be displayed in the drop downs box background ---
WhatDropdownImage EventSelected




'response.write("<br>sLeagueSelected="&sLeagueSelected)
'response.write("<br>Process="&process)
'response.write("<br>EventSelected="&EventSelected)
'response.write("<br>MainImage="&MainImage)


' --- TEMPORARY
'sLeagueSelected="NATL2010"


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

	CASE "refund"

		GetPriorSkiYear

		PageTitle="Refund Report for - "&sPriorYear&" Nationals"
		PageSubTitle="Beta Version"
		Refunds
		CreatePageHead 1000
		IF NOT rs.eof THEN DisplayResult 1000




	CASE "grrank"

		PageTitle="Sample GR Ranking Report"
		PageSubTitle="Version with No Division Selection"
		GRRanking
		CreatePageHead 1000
		IF NOT rs.eof THEN DisplayResult 1000

	CASE "loccontacts"
		PageTitle="LOC Contact List"
		PageSubTitle="Tournaments with "&EventName&" Events 2008"

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
		sTourID="13S999"
		PageTitle="Survey Results - 2013 Goode National Championships"
		PageSubTitle="Okeeheelee, FL"
    SurveyResults
		CreatePageHead 700
		IF NOT rs.eof THEN DisplayResult 700
		WriteIndexPageFooter

	CASE "hotelcount"
		WriteIndexPageHeader
		sTourID="13S999"
		PageTitle="Nights Stayed By Hotel"
		PageSubTitle="2013 Goode National Championships"
		Survey_CountByHotel
		CreatePageHead 700
		IF NOT rs.eof THEN DisplayResult 700
		WriteIndexPageFooter

	CASE "hotellist"
		WriteIndexPageHeader
		sTourID="13S999"
		PageTitle="Hotel Answer Options"
		PageSubTitle="2013 Goode National Championships"
		Survey_HotelList
		CreatePageHead 700
		IF NOT rs.eof THEN DisplayResult 700
		WriteIndexPageFooter

	CASE "hoteldetail"
		WriteIndexPageHeader
		sTourID="13S999"
		PageTitle="Hotel Detail"
		PageSubTitle="2013 Goode National Championships"
		Survey_HotelDetail
		CreatePageHead 700
		IF NOT rs.eof THEN DisplayResult 700
		WriteIndexPageFooter


	CASE "bioinfo"
		WriteIndexPageHeader
		sTourID="13S999"
		PageTitle="Skier Bio Info"
		PageSubTitle="2013 Goode National Championships"
		GetBioInfo
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

	CASE "regpl"
		'WriteIndexPageHeader
		PageTitle="Registration Payment Log"
		PageSubTitle="By Date"
    PaymentLogByDate
		CreatePageHead 725
		IF NOT rs.eof THEN DisplayResult 1400
		'WriteIndexPageFooter


	CASE "regpldet"
		'WriteIndexPageHeader
		PageTitle="Registration Payment DETAIL"
		PageSubTitle="By OrderNo"
    PaymentTransByOrderNo
		CreatePageHead 725
		IF NOT rs.eof THEN DisplayResult 1000
		'WriteIndexPageFooter



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
		
	CASE ELSE 
		WriteIndexPageHeader
    response.write("Invalid Report")
    
		WriteIndexPageFooter     		
END SELECT






' ---------------------------------------------------------------------------------------
' ------------------  BOTTOM OF MAIN PROGRAM CODE  	---------------------------------
' ---------------------------------------------------------------------------------------



' --------------------
  SUB GetPriorSkiYear
' --------------------


' --- Get prior SkiYear
sSQL = " SELECT SkiYear FROM "&SkiYearTableName&" WHERE SkiYearID=(SELECT SkiYearID FROM "&SkiYearTableName&" WHERE DefaultYear='1')-1"
rs.open sSQL, SConnectionToTRATable
IF NOT rs.eof THEN sPriorYear=rs("SkiYear")
rs.close


END SUB


 

' ---------------------
  SUB DisplayResult (tabwidth)
' ---------------------



	rs.movefirst

	' ---------------  Displays table HEADINGS  ----------------------

	%>
	<TABLE class="innertable" Align=center WIDTH=<%=tabwidth%>px >
	  <TR><%

		FOR i = 0 TO rs.fields.count - 1
			TempFN = rs.fields(i).name
			j = 0 %>

	   		<th ALIGN="Center" vAlign="top" nowrap>
			  <FONT COlOR="#FFFFFF" FACE="<%=font1%>" SIZE="<%=fontsize1%>"><%=Rs.Fields(i).name%></FONT>
			</th><%
		NEXT %>
	  </TR><%

	' --------------  Display table data here with paging --------------------------

	DO WHILE NOT rs.eof

		'IF rowCount = rs.PageSize THEN EXIT DO	%>

 		<TR><%

		AllowEdit=true

		FOR i = 0 TO rs.fields.count - 1
	
			RowColor=""
			TempFN = rs.fields(i).name
			IF TempFN="TourID" THEN
					IF RIGHT(LEFT(rs.Fields(i).value,6),3)="001" OR (RIGHT(LEFT(rs.Fields(i).value,6),3)="999" AND ThisYear<>LEFT(rs.Fields(i).value,2)) THEN
							RowColor="background-color:"&scolor08
					ELSEIF ThisYear<>LEFT(rs.Fields(i).value,2) THEN
							RowColor="background-color:"&scolor04
					END IF
			END IF

			IF trim(rs.fields(i).name)="COA Avg" THEN
					%><TD ALIGN="right" width=25% style="<%=RowColor%>"><%
			ELSE
					%><TD ALIGN="center" style="<%=RowColor%>"><%
			END IF 
			
			%>
			<FONT SIZE="1">&nbsp;
			<%
	    IF isnull(rs.Fields(i).value) THEN
					response.write ("&nbsp;")
    	ELSEIF process=regpl AND trim(rs.fields(i).name)="OrderNo" THEN
    			%>
    			<a href="/rankings/<%= ThisFileName %>?process=regpldet&misc1=<%= Rs.Fields(i).Value %>"><%=Rs.Fields(i).Value%></a>
    			<% 
    	ELSE
					Response.Write(trim(Rs.Fields(i).Value)) 
			END IF  
			%>&nbsp;
			 </FONT>
			</TD><%

		NEXT	%>

		</TR><% 
		rowCount = rowCount + 1
		rs.movenext
	LOOP %>

	</TABLE>
<br><br><%

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

		ThisTextColor=""
		IF Rs.Fields(0).Value=ThisStartDate OR DateAdd("d",-DiffBetweenStartDates,Rs.Fields(0).Value)=LastStartDate THEN
				ThisTextColor="Red"
		ELSEIF Rs.Fields(0).Value=Date THEN
				ThisTextColor="Blue"
		END IF
		
		%>
 		<TR>
				<td bgcolor="<%=RowColor%>">
					<font color="<%=ThisTextColor%>" size=<%=fontsize1%>><%=Rs.Fields(0).Value%></font>
				</td>
				<td>
					<font color="<%=ThisTextColor%>" size=<%=fontsize1%>><%= DateDiff("d", Rs.Fields(0).Value, ThisStartDate) %></font> 
				</td>
				<td>
					<font color="<%=ThisTextColor%>" size=<%=fontsize1%>><%=Rs.Fields(1).Value%></font>
				</td> 
				<td>
					<% TotalThis = TotalThis + CInt(Rs.Fields(1).Value)	%>
					<font color="<%=ThisTextColor%>" size=<%=fontsize1%>><%=TotalThis%></font>
				</td>
				<td>
					<font color="<%=ThisTextColor%>" size=<%=fontsize1%>><%= DateAdd("d",-DiffBetweenStartDates,Rs.Fields(0).Value) %></font>
				</td>
				<td>
					<font color="<%=ThisTextColor%>" size=<%=fontsize1%>><%= DateDiff("d", Rs.Fields(0).Value, LastStartDate+366+DiffBetweenStartDates) %></font> 
				</td>
				<td>
					<font color="<%=ThisTextColor%>" size=<%=fontsize1%>><%=Rs.Fields(2).Value%></font>
				</td> 
				<td>
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

'response.write("<br>AdminMenuLevel="&AdminMenuLevel)

' drop
%>
<form action="/rankings/<%=ThisFileName%>?misc1=<%=Misc1%>" method="post">
  <input type="hidden" name="process" value="<%=Process%>">

<TABLE class="droptable" Align=center WIDTH=<%=PageHeadWidth%>px height=175 background="<%=MainImage%>">

  <% ' --- Total width 8 columns --- %>	
  <TR>
	<td colspan=6 align=left>
		<font color="<%=TextColor2%>" size="3">&nbsp;&nbsp;<b><%=PageTitle%></b></font>
		<br>
		<font color="<%=TextColor1%>" size="2">&nbsp;&nbsp;<b><%=PageSubTitle%></b></font>
	</td><%
	IF AdminMenuLevel>=50 THEN  %>	
  		<td colspan=1 valign=top align="left">
			<FONT COlOR="<% =TitleColor %>" size="<% =fontsize2 %>"><b>Show SQL</b></font>
			<input type=checkbox name="sShowSQL" <% IF sShowSQL="on" THEN response.write "checked" %>>

		</td>
  		<td colspan=1 valign=top align="left">
			<FONT COlOR="<% =TitleColor %>" size="<% =fontsize2 %>"><b>Stop</b></font>
			<input type=checkbox name="sStop" <% IF sStop="on" THEN response.write "checked" %>>

		</td><%
	ELSE  %>
  		<td colspan=2 width=350 valign=top align="left">&nbsp;</td><%
	END IF %>
  </TR>
  <TR><%
'response.write("<br>process="&process)
'response.write("<br>")
'response.write(process="leaguequalsummary")


	IF process="loccontacts" THEN  
		'response.write("<br>LOC")
		%>
		<td align=right>&nbsp;&nbsp;Select Event:</td>
		<td align=left><%
			LoadAWSAEvents %>
		</td><%

	ELSEIF process="leaguequalsummary" THEN  
		'response.write("<br>LQS")
		%>
		<td align=right>&nbsp;&nbsp;LeagueID: </td>
		<td align=left><%
			LeagueDropBuild_07162010  %>
		</td><%

	ELSEIF process="skierlist" THEN  
		'response.write("<br>LQS")
		%>
		<td align=right>&nbsp;&nbsp;LeagueID: </td>
		<td align=left><%
			LeagueDropBuild_SelectFromAll  %>
		</td><%

	ELSE 
		'response.write("<br>ELSE")
		%>
		<td colspan=2>&nbsp;</td><%
	END IF %>

	<td colspan=6>&nbsp;</td>
  </TR>
<% 'response.end %>
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
sSQL = sSQL + " FROM "&RegPaymentTableName
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
sSQL = sSQL + " FROM "&RegPaymentTableName
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

sTourID = "13S999"

sSQL = sSQL + " SELECT LastUpdate, FirstName, LastName "
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
sSQL = sSQL + " 	ON RIGHT(RE1.MemberID,6)=m.PersonID"
sSQL = sSQL + " 	LEFT JOIN"
sSQL = sSQL + " 		( SELECT MemberID, Div AS TRDiv "
sSQL = sSQL + " 				FROM "&RegDetailTableName
sSQL = sSQL + " 				WHERE TourID='"&sTourID&"' AND Event='T') RE2"
sSQL = sSQL + " 	ON RIGHT(RE2.MemberID,6)=m.PersonID"
sSQL = sSQL + " 	LEFT JOIN"
sSQL = sSQL + " 		( SELECT MemberID, Div AS JUDiv "
sSQL = sSQL + " 				FROM "&RegDetailTableName
sSQL = sSQL + " 				WHERE TourID='"&sTourID&"' AND Event='J') RE3"
sSQL = sSQL + " 	ON RIGHT(RE3.MemberID,6)=m.PersonID"
sSQL = sSQL + " 	ORDER BY LastName, FirstName"



rs.open sSQL, SConnectionToTRATable

END SUB


' -----------------------
  SUB PaymentLogByDate
' -----------------------

sSQL = " SELECT TransDate, OrderNo, TourID, MemberID, FirstName, LastName, Address1, City, State, ZipCode, Email"
sSQL = sSQL + " , Amount, PayType, PayStatus"
'sSQL = sSQL + " , Message"  
sSQL = sSQL + " FROM "&RegPaymentTableName&" m"
sSQL = sSQL + " WHERE TransDate <='"&NOW&"'"
sSQL = sSQL + " AND TransDate>='"&NOW-30&"'"   
sSQL = sSQL + " AND Result='0'"
sSQL = sSQL + " ORDER BY OrderNo DESC"

response.write(sSQL)
'response.end

rs.open sSQL, SConnectionToTRATable

END SUB



' ---------------------------
  SUB PaymentTransByOrderNo
' ---------------------------

Dim sOrderNo
sOrderNo=Request("Misc1")
sSQL = " SELECT TransDate, TransNo, OrderNo, MemberID, TourID, TransCode"
sSQL = sSQL + ", CASE "
sSQL = sSQL + " WHEN TransCode='FEF' THEN 'Entry Fee'"
sSQL = sSQL + " WHEN TransCode='CEF' THEN 'CREDIT Entry Fee'"
sSQL = sSQL + " WHEN TransCode='FLF' THEN 'Late Fee'"
sSQL = sSQL + " WHEN TransCode='OBF' THEN 'AWSEF Donation'"

sSQL = sSQL + " WHEN TransCode='BAN' THEN 'Banquet'"
sSQL = sSQL + " WHEN TransCode='DOF' THEN 'Officials Discount'"
sSQL = sSQL + " WHEN TransCode='DSR' THEN 'Senior Discount'"
sSQL = sSQL + " WHEN TransCode='DFR' THEN 'Juniors Discount'"
sSQL = sSQL + " WHEN TransCode='DCL' THEN 'Club Discount'"
sSQL = sSQL + " WHEN TransCode='OF1' THEN 'Optional Fee 1'"
sSQL = sSQL + " WHEN TransCode='OF2' THEN 'Optional Fee 2'"
sSQL = sSQL + " WHEN TransCode='OF3' THEN 'Optional Fee 3'"
sSQL = sSQL + " WHEN TransCode='OF4' THEN 'Optional Fee 4'"
sSQL = sSQL + " WHEN TransCode='OF5' THEN 'Optional Fee 5'"
sSQL = sSQL + " WHEN TransCode='OF6' THEN 'Optional Fee 6'"
sSQL = sSQL + " WHEN TransCode='OF7' THEN 'Optional Fee 7'"
sSQL = sSQL + " WHEN TransCode='OF8' THEN 'Optional Fee 8'"
sSQL = sSQL + " WHEN TransCode='OF9' THEN 'Optional Fee 9'"
sSQL = sSQL + " WHEN TransCode='OF10' THEN 'Optional Fee 10'"
sSQL = sSQL + " END AS Description"

sSQL = sSQL + ", Amount"
sSQL = sSQL + " 	FROM "&RegTransTableName
'sSQL = sSQL + " 	WHERE OrderNo='689572'"
sSQL = sSQL + " 	WHERE OrderNo='"&sOrderNo&"'"
sSQL = sSQL + " 			ORDER BY TransDate DESC"				

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
sSQL = " SELECT First, Last, Email"
sSQL = sSQL + " FROM"
sSQL = sSQL + " ( "
sSQL = sSQL + " SELECT DISTINCT MemberID, First, Last, Email"
sSQL = sSQL + " FROM "&RegGenTableName&" RG"

sSQL = sSQL + " LEFT JOIN"
sSQL = sSQL + " ( SELECT PersonID, FirstName AS First, LastName AS Last, Address1, City, State, Zip, Email"
sSQL = sSQL + " 		FROM "&MemberLiveTableName&" ) MT"
sSQL = sSQL + " ON CAST(RIGHT(RG.MemberID,8) AS INT)=MT.PersonID"

sSQL = sSQL + " WHERE TourID='"&sTourID&"'"
sSQL = sSQL + " AND Email IS NOT NULL"
sSQL = sSQL + " ) AS A"
sSQL = sSQL + " ORDER BY Last, First" 

'response.write(sSQL)
'response.end

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
sSQL = sSQL + ",	[Entry Count OLR] AS [Entry Count<br>OLR], COALESCE([Entry Count All],0) AS [Entry Count<br>All]"
sSQL = sSQL + ",	[Event Count OLR] AS [Event Count<br>OLR], COALESCE([Event Count All],0) AS [Event Count<br>All]"
sSQL = sSQL + " FROM usawsrank.RegisterGenNew Y"
	
sSQL = sSQL + "	LEFT JOIN"
sSQL = sSQL + "	( 	SELECT LEFT(TourID,2) AS Year, Count(Distinct TourID) AS [Tour Count OLR]"
sSQL = sSQL + "			FROM usawsrank.RegisterGenNew"
sSQL = sSQL + "			GROUP BY LEFT(TourID,2) ) TourCntOLR"
sSQL = sSQL + "	ON LEFT(Y.TourID,2) = TourCntOLR.Year"

sSQL = sSQL + "	LEFT JOIN"
sSQL = sSQL + "	( 	SELECT LEFT(TourID,2) AS Year, Count(Distinct TourID) AS [Tour Count All]"
sSQL = sSQL + "			FROM usawsrank.Scores "
sSQL = sSQL + "			GROUP BY LEFT(TourID,2) ) TourCntAll"
sSQL = sSQL + "	ON LEFT(Y.TourID,2) = TourCntAll.Year "

sSQL = sSQL + "	LEFT JOIN"
sSQL = sSQL + "	( SELECT LEFT(TourID,2) AS Year, Count(TourID) AS [Entry Count OLR] "
sSQL = sSQL + "		FROM usawsrank.RegisterGenNew "
sSQL = sSQL + "			GROUP BY LEFT(TourID,2)  ) EntCntOLR "
sSQL = sSQL + "	ON LEFT(Y.TourID,2) = EntCntOLR.Year"

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "	( SELECT LEFT(TourID,2) AS Year, Count(DISTINCT MemberID+TourID) AS [Entry Count All] "
sSQL = sSQL + "		FROM usawsrank.Scores "
sSQL = sSQL + "			GROUP BY LEFT(TourID,2)  ) EntCntAll"
sSQL = sSQL + "	ON LEFT(Y.TourID,2) = EntCntAll.Year"

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "	( SELECT LEFT(TourID,2) AS Year, Count(MemberID) AS [Event Count OLR]"
sSQL = sSQL + "		FROM usawsrank.RegisterEvents "
sSQL = sSQL + "			GROUP BY LEFT(TourID,2)  ) EvtCntOLR "
sSQL = sSQL + "	ON LEFT(Y.TourID,2) = EvtCntOLR.Year "

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "	( SELECT LEFT(TourID,2) AS Year, Count(TourID) AS [Event Count All] "
sSQL = sSQL + "		FROM usawsrank.Scores "
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

SET rs=Server.CreateObject("ADODB.recordset")
sSQL =  "SELECT SkiYear, SkiYearName FROM "&SkiYearTableName&" WHERE DefaultYear='1'"
rs.open sSQL, SConnectionToTRATable, 3, 3
sSkiYear=rs("SkiYear")
ThisSkiYear_2Digit=RIGHT(rs("SkiYearName"),2)
rs.close

sSQL =  "SELECT TournAppID FROM "&TRegSetupTableName
sSQL = sSQL + " WHERE LEFT(TournAppID,2) = '"&ThisSkiYear_2Digit&"' AND RIGHT(LEFT(TournAppID,6),3) = '999'"
rs.open sSQL, SConnectionToTRATable, 3, 3
ThisTournAppID=rs("TournAppID")
rs.close

LastSkiYear_2Digit = CInt(ThisSkiYear_2Digit)-CInt(1)
sSQL =  "SELECT TournAppID FROM "&TRegSetupTableName
sSQL = sSQL + " WHERE LEFT(TournAppID,2) = '"&LastSkiYear_2Digit&"' AND RIGHT(LEFT(TournAppID,6),3) = '999'"
rs.open sSQL, SConnectionToTRATable, 3, 3
LastTournAppID=rs("TournAppID")

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
sSQL = sSQL + " , Coalesce(RC1.CntThis,0) AS CntThis"
sSQL = sSQL + " , Coalesce(RC2.CntLast,0) AS CntLast" 
sSQL = sSQL + " FROM "
sSQL = sSQL + " ( "
sSQL = sSQL + " SELECT DISTINCT RegisterDate FROM"
sSQL = sSQL + " (SELECT RegisterDate FROM usawsrank.RegisterGenNew WHERE LEFT(TourID,6)='"&ThisTournAppID&"'"
sSQL = sSQL + " UNION"
sSQL = sSQL + " SELECT RegisterDate+366+("&DiffBetweenStartDates&") AS RegisterDate FROM usawsrank.RegisterGenNew WHERE LEFT(TourID,6)='"&LastTournAppID&"' ) AS A"
sSQL = sSQL + " ) AS RC" 

sSQL = sSQL + "	LEFT JOIN"
sSQL = sSQL + "	(SELECT RegisterDate, Count(MemberID) AS CntThis"
sSQL = sSQL + "		FROM usawsrank.RegisterGenNew"
sSQL = sSQL + " WHERE LEFT(TourID,6)='"&ThisTournAppID&"'"
sSQL = sSQL + "	 GROUP BY RegisterDate) AS RC1" 
sSQL = sSQL + "	ON RC.RegisterDate=RC1.RegisterDate"

sSQL = sSQL + "	LEFT JOIN"
sSQL = sSQL + "	(SELECT RegisterDate, Count(MemberID) AS CntLast"
sSQL = sSQL + "		FROM usawsrank.RegisterGenNew"
sSQL = sSQL + "	WHERE LEFT(TourID,6)='"&LastTournAppID&"'"
sSQL = sSQL + "		GROUP BY RegisterDate) AS RC2 "
sSQL = sSQL + "	ON RC.RegisterDate=(RC2.RegisterDate+366+("&DiffBetweenStartDates&"))"
'"&DiffBetweenStartDates&"
sSQL = sSQL + " ORDER BY RC.RegisterDate"

response.write(sSQL)
response.end
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
  SUB Refunds
' ---------------------


sSQL = "SELECT RG.MemberID, MT.FirstName, MT.LastName, MT.Address1, MT.City, MT.State, MT.Zip, REC.EnterCount AS [Events<br>Entered], ST1.SkiedCount AS [Events<br>Skied], CAST(COALESCE(Payments,0) AS money) AS Payments"
sSQL = sSQL + "			FROM "&RegGenTableName&" AS RG"

sSQL = sSQL + "		LEFT JOIN "&MemberTableName&" AS MT"
sSQL = sSQL + "			ON MT.PersonIDWithCheckDigit=RG.MemberID"

sSQL = sSQL + "		JOIN"
sSQL = sSQL + "			(SELECT MemberID, TourID, Count(MemberID) AS EnterCount FROM "&RegDetailTableName
sSQL = sSQL + "				GROUP BY MemberID, TourID) AS REC"
sSQL = sSQL + "			ON REC.MemberID=RG.MemberID AND LEFT(REC.TourID,6)=LEFT(RG.TourID,6)"

sSQL = sSQL + "		LEFT JOIN"
sSQL = sSQL + "			(SELECT MemberID, TourID, Count(MemberID) AS SkiedCount FROM "&RawScoresTableName
sSQL = sSQL + "				GROUP BY MemberID, TourID) AS ST1"
sSQL = sSQL + "			ON ST1.MemberID=RG.MemberID AND LEFT(ST1.TourID,6)=LEFT(RG.TourID,6)"

sSQL = sSQL + "			LEFT JOIN (SELECT MemberID, SUM(Amount) AS Payments FROM "&RegPaymentTableName
sSQL = sSQL + "				WHERE  LEFT(TourID,6) = '09S999' AND Result = '0'"
sSQL = sSQL + "				GROUP BY MemberID) AS TP"
sSQL = sSQL + "				ON TP.MemberID = RG.MemberID"
	
sSQL = sSQL + "			WHERE LEFT(RG.TourID,6)='09S999'"
sSQL = sSQL + "				AND ( (ST1.SkiedCount <> REC.EnterCount AND ST1.SkiedCount<>2)"
sSQL = sSQL + "				OR ST1.SkiedCount IS NULL)"

sSQL = sSQL + "				AND TP.Payments<>'0' AND Payments<>'355'"



sSQL = sSQL + "			ORDER BY RG.MemberID"

'response.write(sSQL)
rs.open sSQL, SConnectionToTRATable


END SUB


' ---------------------
  SUB LOCContacts
' ---------------------

Dim sSQL

sSQL = "SELECT TournAppID AS TourID, RIGHT(LEFT(TourID,3),1) AS Reg, TName AS Tournament, TCity AS City, TState AS ST, TDirName AS [Tour Director], TDirEmail AS [Tour Dir Email], TSponsor AS [Tour Sponsor]"  
sSQL = sSQL + "	FROM sanctions.dbo.Tschedul AS TS"
	
sSQL = sSQL + "	JOIN"
sSQL = sSQL + "		  (SELECT Distinct TourID"
sSQL = sSQL + "			FROM usawsrank.Scores"
sSQL = sSQL + "			WHERE Score IS NOT NULL" 
IF EventSelected<>"" THEN
	sSQL = sSQL + "				AND Event='"&EventSelected&"'"
END IF

sSQL = sSQL + ") AS RS"
sSQL = sSQL + "	ON LEFT(Rs.TourID, 6)=LEFT(TS.TournAppID, 6)"

sSQL = sSQL + " WHERE LEFT(TourID,2) = '08'"
sSQL = sSQL + " AND RIGHT(LEFT(TourID,3),1) <> 'U'"		

sSQL = sSQL + "ORDER BY TournAppID"

rs.open sSQL, SConnectionToTRATable


END SUB





' ---------------------
  SUB AWSEFDonorsByLOC
' ---------------------

Dim sSQL

sSQL = "SELECT TName AS [Tournament Name], SUM(AWSEFDonation) AS Donations, Left(Convert(char,TDateE,111),10) as [End Date], TDirName AS Contact, TDirAddress AS Address, TDirCity AS City, TDirState AS State, TDirZip AS Zip, TDirEmail AS Email"  
sSQL = sSQL + "		FROM "&RegGenTableName&" AS RG, "&MemberTableName&" AS MT, "&SanctionTableName&" AS ST"
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
sSQL = sSQL + "		FROM "&RegGenTableName&" AS RG, "&MemberTableName&" AS MT, "&SanctionTableName&" AS ST"
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

ChooseSQL("SELECT DISTINCT LeagueID, LeagueName FROM "&LeagueTableName&" WHERE SptsGrpID='"&Session("sSptsGrpID")&"' ORDER BY LeagueID")

'response.write("<br>EOF2=")
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

ChooseSQL("SELECT LeagueID, LeagueName FROM "&LeagueTableName&" WHERE QualifyTour<>'' AND Status<>'X' AND RIGHT(LeagueID,4)=(SELECT MAX(RIGHT(LeagueID,4)) FROM "&LeagueTableName&") ORDER BY LeagueName DESC") 



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