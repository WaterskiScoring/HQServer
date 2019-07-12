<!--#include file="settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->
<!--#include virtual="/rankings/tools_Definitions.asp"-->
<%

' -----------------------------------------------------------
' --- This display the results for US Elite Team Selection
' --- 
' --- Last updated:  	6-5-2009
' --- Developer:	Mark Crone
' -----------------------------------------------------------



DefineTRAStyles

Dim ThisFileName, TeamXTableName
Dim EventSelected, EventName, process, SubProcess, Whataction
Dim sTourID, sTeamNo, FedSelected, sSLMin, sTRMin, sJUMin
Dim sUSBeginDate, sUSEndDate, sIntBeginDate, sIntEndDate, sIntFileName
Dim sPassword, validpassword, sTName
Dim action
Dim TurnOnList
Dim sEmail

Dim USAScore1Title, USAScore2Title, USAScore3Title
Dim INTScore1Title, INTScore2Title, INTScore3Title
Dim WinLossTitle, RiskFactTitle, DWinTitle, DRiskTitle, RiskRewardTitle


USAScore1Title="Best 2017 Team Trials Score"
USAScore2Title="Second 2017 Team Trials Score"
USAScore3Title="Not Used"
USATrialsTitle="Trials Score"

INTScore1Title="Best Ranking List Score"
INTScore2Title="Second Ranking List Score"
INTScore3Title="Median of All Ranking List Scores"
INTTrialsTitle="Trials Score"

WinLossTitle="Difference between Team Oveall score of TeamX and the highest scoring International team" 
RiskFactTitle="Change in Team Overall score when 4th score is substituted for 1st score in each event" 
DWinTitle="Difference between the WinLoss of the TeamX and the highest scoring team"
DRiskTitle="Difference between the RiskFactor of the TeamX and the highest scoring team"
RiskRewardTitle="(Risk Factor of 1st Team minus the Risk Factor of TeamX)/(Win Factor of 1st Team minus the Win Factor of TeamX)"
BlendedTitle="Composite calculation of Win Loss less Risk Factor" 








' --- Hides team 
TurnOnList=false


ThisFileName="IAC_Reports.asp"

TeamXTableName="usawsrank.IAC_TeamX_IndivScrs_2017"
CalcTourID="17TEAM"




WhatAction=TRIM(request("WhatAction"))

process=TRIM(LCASE(request("process")))
sEmail=TRIM(LCASE(request("sEmail")))
sPassword=TRIM(LCASE(request("sPassword")))
sListSeq=TRIM(UCASE(request("sListSeq")))
IF sListSeq="" THEN sListSeq="W"
sOverlay=request("sOverlay")

'sListSeq="RR"

sMemberID=TRIM(request("sMemberID"))




'response.write("<br>ADM = "&Session("adminmenulevel"))
IF process="validpw" THEN
		process="validpw"
ELSEIF (Session("validpassword")="" OR Session("validpassword")="no") AND (Session("adminmenulevel")<40 OR process="") THEN
		process="getpw"
END IF









TourTableWidth=675
FedSelected="USA"


sTourID=TRIM(Request("sTourID"))

IF sTourID="" THEN

	' --- Finds DefaultTour setting if ---
	Set rs=Server.CreateObject("ADODB.recordset")
	sSQL = " SELECT CT.*, ST.TName"
	sSQL = sSQL + " FROM "&IAC_ControlTableName&" AS CT"
	sSQL = sSQL + " LEFT JOIN "&Sanctiontablename&" AS ST"
	sSQL = sSQL + " ON ST.TournAppID=LEFT(CT.TourID,6)"
	sSQL = sSQL + " WHERE CT.DefaultTour='1'"

	'response.write(sSQL)
	'response.end

	rs.open sSQL, SConnectionToTRATable

	IF NOT rs.eof THEN 
		sTourID=rs("TourID")
		sTName=rs("TName")
	END IF

	rs.close

END IF





IF Request("sTeamNo")="" THEN 
		'response.write("<br>Did not find Team Number")
		'response.end
		SetDefaultTeam
ELSE 
		'response.write("<br>Found Team Number")
		'response.end

	sTeamNo=cdbl(Request("sTeamNo")) 
END IF





EventSelected=TRIM(Request("EventSelected"))
SELECT CASE EventSelected
	CASE "S"
		EventName="Slalom"
	CASE "T"
		EventName="Trick"
	CASE "J"
		EventName="Jump"
END SELECT

IF Request("sSLMin")<>"" THEN sSLMin=cdbl(Request("sSLMin")) ELSE sSLMin=4 END IF
IF Request("sTRMin")<>"" THEN sTRMin=cdbl(Request("sTRMin")) ELSE sTRMin=4 END IF
IF Request("sJUMin")<>"" THEN sJUMin=cdbl(Request("sJUMin")) ELSE sJUMin=4 END IF







'response.write("<br>WhatAction = "&WhatAction)
'response.write("<br>")
'response.write(WhatAction = "Return to Menu")

' --- Redefines process based on button push instead of Report Selection ---
SELECT CASE WhatAction
	CASE "Control Panel", "controlpanel"
			process="controlpanel"
	CASE "Save Changes"
			process="saveiac"
	CASE "Return to Trials"
			process="teamlist"
	CASE "Display Scores"
			process="indivscores"
	CASE "Return to Menu"
			response.redirect("/rankings/defaultHQ.asp")
END SELECT

SELECT CASE process
	CASE "getpw"
			IACGetPW

	CASE "validpw"	
			'--- Validate sPassword
			NowValidatePW

	CASE "foreignsummary"
		PageTitle="<br> International Team Score Summary "&Session("adminmenulevel")
		PageSubTitle=" Tournament: "&sTName&" - "&sTourID
		CreatePageHead
		foreignsummary

	CASE "foreigncandidates"
		PageTitle="<br> International Team Candidates CAN AUS FRA GBR BLR ITA"&Session("adminmenulevel")
		PageSubTitle=" Tournament: "&sTName&" - "&sTourID&" - Scores Exceeding: MS/WS 60/55 bouys MT/WT 8000/6500 MJ/WJ 55/45)" 
		CreatePageHead
		ForeignTeamCandidates
		
	CASE "ustrialssummary"
		PageTitle="<br> US Team Trials Participants Score Summary Report"
		PageSubTitle=" Tournament: "&sTName&" - "&sTourID
		CreatePageHead
		ustrialssummary

	CASE "teamdetail"
		PageTitle="<br> US and International Team Overall Scoring Detail - Team No: "&sTeamNo
		PageSubTitle=" Tournament: "&sTName&" - "&sTourID
		CreatePageHead
		teamdetail

	CASE "indivscores"
		PageTitle="<br> USA Team Selection Candidates"
		PageSubTitle=" Member Score Detail"
		indivscores

	CASE "teamlist"
		IF sListSeq="R" THEN
				ListSeqName="Risk Factor"			
		ELSEIF sListSeq="B" THEN
				ListSeqName="Blended"			
		ELSEIF sListSeq="RR" THEN
				ListSeqName="Risk Reward"			
		ELSE
				ListSeqName="Win Factor"			
		END IF
		PageTitle="<br>TeamX Combinations List "
		PageSubTitle="Order By "&ListSeqName&" - Top 300 Teams <br> Tournament: "&sTName&" - "&sTourID
		CreatePageHead
		teamlist

	CASE "medianlist"
		PageTitle="<br>Median Scores List "
		PageSubTitle="All US Skiers for Specified Date Range <br> Tournament: "&sTName&" - "&sTourID
		CreatePageHead
		DisplayMedianList

	CASE "teamstatistics"
		PageTitle="<br>Team Trials Statistics "
		PageSubTitle="Count of Team Combinations w/Minimum of 3 Scores <br> Tournament: "&sTName&" - "&sTourID
		CreatePageHead
		TeamStatistics
		
		
	CASE "ratiolist"
		PageTitle="<br>Formula Comparison "
		PageSubTitle="All Trials Skiers With International Formula (Best+Median)/2 <br> Tournament: "&sTName&" - "&sTourID
		CreatePageHead

		DisplayRatioList

	CASE "controlpanel"
		response.write("<br>In Control Panel")

		PageTitle="Elite Team Trials - Selection Program"
		PageSubTitle="Control Panel"
		CreatePageHead
		controlpanel

	CASE "editoverlay"
		PageTitle="Edit Skiers Showing on Overlay"
		PageSubTitle="Control Panel Function"
		CreatePageHead
		ChangeOverlay

	CASE "saveiac"		
		response.write("<br>Save IAC")
		'response.end
		saveiac
		PageTitle="Elite Team Trials - Simulation Program"
		PageSubTitle="Control Panel - Tournament: "&sTourID
		controlpanel


	CASE ELSE
		PageTitle="Select Report Option"
		PageSubTitle="Press Update Display when Ready - Tournament: "&sTourID
		CreatePageHead

		%><br><center><font size=3 color="#FFFFFF"><b>Select a Report and Update Display</b></font></center><%	

END SELECT






' ---------------------------------------------------------------------------------------
' ------------------  BOTTOM OF MAIN PROGRAM CODE  	---------------------------------
' ---------------------------------------------------------------------------------------


' ----------------------
  SUB NowValidatePW
' ---------------------- 

'response.write("IN NOW")

Session("validpassword")="no"

'response.write("sPassword="&sPassword)
'response.end

IF (sPassword="iac2017" AND sEmail<>"") OR sPassword="zzzaaa"  THEN
		Session("validpassword")=sPassword

		IF sEmail<>"" THEN
				eMailTo="cronemarka@gmail.com"
				eMailSubj=" IAC Login" 
				eMailFrom="Competition@USAWaterski.org"
				eMailBody = "Login complete for Email: "&sEmail	



				' ---------------------------------------------------------------
				' --- Now assign the components to the standard email objects ---
				' ---------------------------------------------------------------

				SetupEmailService

				objMessage.Subject = eMailSubj
				objMessage.From = eMailFrom
				objMessage.To = eMailTo
				objMessage.HTMLBody = eMailBody
 		
				' --- Finally send the message, and then clear that object
				objMessage.Send
				SET objMessage = Nothing
		END IF
END IF

response.redirect("/rankings/IAC_Reports.asp?spassword="&spassword&" ")		


END SUB








' ---------------------
  SUB DisplayResult
' ---------------------


IF NOT rs.eof THEN
	rs.movefirst

	' ---------------  Displays table HEADINGS  ----------------------

	%>
	<BR>
	<TABLE class="innertable" Align=center WIDTH=1200px >
		<%

		IF SubProcess="usdetail" THEN 
				SubProcessHead = "US Team Score and Individual Overall Detail - TeamX No: "&sTeamNo	
		ELSEIF SubProcess="intdetail" THEN 
				SubProcessHead="Country: "&FedSelected&" - Team & Individual Overall Scores - TeamX No: "&sTeamNo
		END IF


		IF SubProcess="usdetail" OR SubProcess="intdetail" THEN 
				%>	
		  	<TR>
		    	<th ALIGN="Center" colspan=17>
						<font size=2 color="#FFFFFF"><b><%=SubProcessHead%></b></font>
		    	</th>
		  	</TR>
		  	<%
		END IF 
		

		%>
	  <TR><%
		' --- Colors cell if scores is used ---
		SLHighlight="N"
		TRHighlight="N"
		JUHighlight="N"


		' --- Display this member row ---
		FOR i = 0 TO rs.fields.count - 1
				TempFN = rs.fields(i).name
				
				CellTitle=""
				IF process="ustrialssummary" THEN
						IF TempFN="Slalom<br>1st" OR TempFN="Trick<br>1st" OR TempFN="Jump<br>1st" THEN CellTitle=USAScore1Title
						IF TempFN="Slalom<br>2nd" OR TempFN="Trick<br>2nd" OR TempFN="Jump<br>2nd" THEN CellTitle=USAScore2Title
						IF TempFN="Slalom<br>3rd" OR TempFN="Trick<br>3rd" OR TempFN="Jump<br>3rd" THEN CellTitle=USAScore3Title
				ELSEIF process="foreignsummary" THEN
						IF TempFN="Slalom<br>Score1" OR TempFN="Tricks<br>Score1" OR TempFN="Jump<br>Score1" THEN CellTitle=INTScore1Title
						IF TempFN="Slalom<br>Score2" OR TempFN="Tricks<br>Score2" OR TempFN="Jump<br>Score2" THEN CellTitle=INTScore2Title
						IF TempFN="Slalom<br>Score3" OR TempFN="Tricks<br>Score3" OR TempFN="Jump<br>Score3" THEN CellTitle=INTScore3Title
				ELSEIF process="teamlist" THEN 
						IF TempFN="M" THEN CellTitle="Number of Male Skiers"
						IF TempFN="F" THEN CellTitle="Number of Female Skiers"
						IF TempFN="#<br>SL" THEN CellTitle="Number of Slalom Skiers"
						IF TempFN="#<br>TR" THEN CellTitle="Number of Trick Skiers"
						IF TempFN="#<br>JU" THEN CellTitle="Number of Jumpers"

						IF TempFN="Win<br>Loss<br>Diff" THEN CellTitle=WinLossTitle
						IF TempFN="Risk<br>Factor" THEN CellTitle=RiskFactTitle
						IF TempFN="DWin" THEN CellTitle=DWinTitle							
						IF TempFN="DRisk" THEN CellTitle=RRiskTitle
						IF TempFN="Risk<br>Reward<br>Ratio" THEN CellTitle=RiskRewardTitle
						IF TempFN="Blended" THEN CellTitle=BlendedTitle
				END IF
		
				j = 0 
				IF Session("AdminMenuLevel")>=50 AND process="ustrialssummary" AND i=0 THEN
					%>
					<Th ALIGN="center"><FONT COlOR="#FFFFFF" SIZE="<%=fontsize1%>">Add</FONT></Th>
					<Th ALIGN="center"><FONT COlOR="#FFFFFF" SIZE="<%=fontsize1%>">Remove</FONT></Th>
					<%
				END IF
				IF RIGHT(rs.fields(i).name,3) <> "Pos" AND LEFT(rs.fields(i).name,2) <> "MK" AND (NOT (sOverlay<>"" AND rs.fields(i).name="Marked")) THEN 
						%>
			   		<th ALIGN="Center" vAlign="top" bgcolor="<%=CellColor%>" nowrap>
					  	<FONT COlOR="#FFFFFF" SIZE="<%=fontsize1%>" title="<%=CellTitle%>"><%=Rs.Fields(i).name%></FONT>
						</th>
						<%
				END IF
		NEXT %>
	  </TR><%



Dim sColorSelected, sOverlayON
sColorSelected="red"

MemberIDWidth="10"	' -- 1 column 
NameWidth="16"	' -- 2 columns 
GenderWidth="5"	' -- 1 column 
TrialsScoreWidth="8"	' -- 3 columns 
OverallScoreWidth="10"	' -- 3 columns 




	' --------------  Display table data here with paging --------------------------
	DO WHILE NOT rs.eof

		%>

 		<TR><%

		AllowEdit=true

		IF Session("AdminMenuLevel")>=50 AND process="ustrialssummary" THEN
			%>
			<TD ALIGN="center" vAlign="top"><FONT SIZE="<%=fontsize1%>"><% WriteLink "?Whataction=addtoover&sMemberID="&rs.fields(3).Value&"&process=ustrialssummary&sOverlay="&sOverlay&" ","Add","" %></FONT></TD>
			<TD ALIGN="center" vAlign="top"><FONT SIZE="<%=fontsize1%>"><% WriteLink "?Whataction=delfromover&sMemberID="&rs.fields(3).Value&"&process=ustrialssummary&sOverlay="&sOverlay&" ","Remove","" %></FONT></TD>
			<%
		END IF


		' --- Reset for every row ---
		SLHighlight="N"
		TRHighlight="N"
		JUHighlight="N"


		FOR i = 0 TO rs.fields.count - 1


			' --- Test for Athlete MemberID
			CellColor=""
			IF sOverlay="on" AND (LEFT(rs.Fields(i).name,2)="MK" AND rs.Fields(i).value="Y") THEN
					' --- Change the background to GREEN if this member is marked
					CellColor=tcolor03
					'i = i + 1	
			ELSEIF sOverlay<>"" AND Rs.Fields(i).value = "N" THEN
					'i = i + 1
			END IF
			
			IF process="teamlist" AND LEFT(rs.Fields(i).name,2)="MK" THEN i=i+1


			TempFN = rs.fields(i).name  

			' --- Determines whether the SL, TR and JU cells need highlights ---
			IF rs.fields(i).name="SLPos" THEN
					IF rs.fields(i).value<=3 THEN SLHighlight="Y"
			END IF
			IF rs.fields(i).name="TRPos" THEN
					IF rs.fields(i).value<=3 THEN TRHighlight="Y"
			END IF
			IF rs.fields(i).name="JUPos" THEN
					IF rs.fields(i).value<=3 THEN JUHighlight="Y"
			END IF
			

			' --- Loops thru the columns ---
			IF RIGHT(rs.fields(i).name,3) <> "Pos" THEN
					ColWidth=""					
					' IF process="teamdetail" OR SubProcess="intdetail" THEN
					IF process="teamdetail" THEN
							CellColor="#FFFFFF"
							IF (rs.fields(i).name="Slalom<br>Overall" AND SLHighlight="Y") OR (rs.fields(i).name="Trick<br>Overall" AND TRHighlight="Y") OR (rs.fields(i).name="Jump<br>Overall" AND JUHighlight="Y") THEN CellColor="#D6ECF2"

							IF rs.fields(i).name="MemberID" THEN ColWidth=MemberIDWidth
							IF rs.fields(i).name="Sex" THEN ColWidth=GenderWidth
							IF rs.fields(i).name="Last" OR rs.fields(i).name="First" THEN ColWidth=NameWidth
							IF rs.fields(i).name="Slalom<br>Score" OR rs.fields(i).name="Trick<br>Score" OR rs.fields(i).name="Jump<br>Score" THEN ColWidth=TrialsScoreWidth
							IF rs.fields(i).name="Slalom<br>Overall" OR rs.fields(i).name="Trick<br>Overall" OR rs.fields(i).name="Jump<br>Overall" THEN ColWidth=OverallScoreWidth
							
							' --- Changes text to red if setting highest score ---
							TextColor="#000000"
							IF rs.fields(i).name="Slalom<br>Overall" OR rs.fields(i).name="Trick<br>Overall" OR rs.fields(i).name="Jump<br>Overall" THEN 
									IF rs.fields(i).value=1000 THEN TextColor="#FF0000"
							END IF
					END IF


					%>
					
					<TD ALIGN="center" width="<%=ColWidth%>%" style="background-color:<%=CellColor%>">
			  		<font size="1" color="<%=TextColor%>">&nbsp;<%
		
						' --- Displays link on TTNo to the team detail page for that TTNo ---
						IF process="teamlist" AND TempFN="TTNo" THEN  
								%><a href="<%=ThisFileName%>?process=teamdetail&sTeamNo=<%=rs.fields(i).value%>&sTourID=<%=sTourID%>"><%=rs.fields(i).value%></a><%

						' --- Displays link on MemberID to the score detail for that member ---
						ELSEIF SubProcess="usdetail" AND i=0 THEN 
								%><a href="<%=ThisFileName%>?process=indivscores&sTourID=<%=sTourID%>&sMemberID=<%=rs.fields(i).value%>"><%=rs.fields(i).value%></a><%
						ELSE

								' --- Displays the fields in query other than TTNo
								SELECT CASE Rs.Fields(i).type
										CASE 3 'numeric'
												Response.Write(Rs.Fields(i).value)
										CASE 4  'numeric'
												Response.Write(formatnumber(Rs.Fields(i).value,2))
										CASE 5  'numeric'
												Response.Write(formatnumber(Rs.Fields(i).value,2))
										CASE 200  'char'
												Response.Write(LEFT(Rs.Fields(i).value,15))
										CASE 131 'numeric'
												Response.Write(formatnumber(Rs.Fields(i).value,2))
										CASE ELSE 'not handled by this function'
			        					Response.Write(Rs.Fields(i).value)
								END SELECT
						END IF  
						%>	
			  		</FONT>
					</TD>
					<%
			END IF
			


			NEXT


		%>

		</TR><% 
		rowCount = rowCount + 1
		rs.movenext
	LOOP %>

	</TABLE>
<br><%

END IF


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



' ----------------------
  SUB CreatePageHead
' ----------------------

%>

<br>
<form action="/rankings/<%=ThisFileName%>" method="post">

  <input type="hidden" name="sPassword" value="<%= sPassword %>">
  <input type="hidden" name="sTourID" value="<%= sTourID %>">

<TABLE class="innertable" Align=center WIDTH=1200px height=100>
  <TR>
		<td colspan=8 style="border-style:none;">
			<font color="<%=TextColor2%>" size="3"><b><%=PageTitle%></b></font>
			<br>
			<font color="<%=TextColor1%>" size="2"><b><%=PageSubTitle%></b></font>
		</td>
  </TR>
  <TR>
	<td colspan=8 style="border-style:none;">&nbsp;</td>
  </TR>



  <TR>

		<td style="border-style:none;"  align=right>Select Report:</td>

		<td colspan=3 style="border-style:none;" align=left>
	  	<select name="process" style="width:20em" >
				<option value="teamlist"<%IF process = "teamlist" THEN Response.Write(" SELECTED ")%>>TeamX Listing - Sequential Sort</option>
				<option value="teamdetail"<%IF process = "teamdetail" THEN Response.Write(" SELECTED ")%>>Specific TeamX - Detail</option>
				<option value="ustrialssummary"<%IF process = "ustrialssummary" THEN Response.Write(" SELECTED ")%>>US Trials Skiers - Score Summary</option>
				<%

				cvg=1
				IF cvg=1 THEN
						%>
						<option value="indivscores"<%IF process = "indivscores" THEN Response.Write(" SELECTED ")%>>US Trials Skier Score - Detail for Skier</option>
						<%
				END IF

				%>
	    	<option value="foreignsummary" <%IF process = "foreignsummary" THEN Response.Write(" SELECTED ")%>>Foreign Skiers - Score Summary</option>
	    	<option value="foreigncandidates" <%IF process = "foreigncandidates" THEN Response.Write(" SELECTED ")%>>Foreign Skiers - All Candidates</option>
	    	<option value="teamstatistics" <%IF process = "teamstatistics" THEN Response.Write(" SELECTED ")%>>Teams Statistics</option>	
				<%

				

				medscr=1
				IF medscr=2 THEN 
						%><option value="medianlist"<%IF process = "medianlist" THEN Response.Write(" SELECTED ")%>>Median Score List - All USA Skiers</option><%
				END IF

				IF Session("adminmenulevel")>=50 THEN 
						%>
						<option value="ratiolist"<%IF process = "ratiolist" THEN Response.Write(" SELECTED ")%>>Selection Score Analysis</option>
						<option value="editoverlay"<%IF process = "editoverlay" THEN Response.Write(" SELECTED ")%>>Change Marked For Overlay</option>
						<%
				END IF 
				%>
      	</select>
		</td>
		<%
		IF process="teamdetail" THEN  
				%>
				<td coalspan=1 style="border-style:none;">&nbsp;Team No: 
					<%
					
					TeamNoDropDown 
					
					%>
				</td>
				<td colspan=3 style="border-style:none;">&nbsp;</td>
				<%
		ELSEIF process="teamlist" THEN  
				%>
				<td style="border-style:none;">&nbsp;</td>
				<td width=150 style="border-style:none;">Slalom: 
					<%

			  	 LoadValuePulldown "sSLMin", sSLMin, 0, 6, 1, enable, false  
			  
			  	%>
				</td>
				<td width=150  style="border-style:none;">Trick: 
					<%
			   
			   	LoadValuePulldown "sTRMin", sTRMin, 0, 6, 1, enable, false  
			   	
			   	%>
				</td>
				<td width=150  style="border-style:none;">Jump: 
					<%
			   
			   	LoadValuePulldown "sJUMin", sJUMin, 0, 6, 1, enable, false  
			   	
			   	%>
				</td>
				<%
		ELSE  
				%><td colspan=4 style="border-style:none;">&nbsp;</td><%
		END IF  
		%>
		</td>
</TR>
<%
IF process = "teamlist" THEN
		%>
		<TR>	
			<td colspan=1 style="border-style:none;">&nbsp;</td>
			<td style="border-style:none;"  align=right><b>Show Overlay:</b>
				<input type=checkbox NAME="sOverlay" title="Check to Display Colored Overlay" <% IF sOverlay<>"" THEN response.write("checked") %>>
			</td>
			<td colspan=2 style="border-style:none;">&nbsp;</td>
			<td style="border-style:none;"  align=right><b>Sort Sequence:</b></td>
			<td style="border-style:none;"  align=center>WinLoss Diff
				<input type=radio NAME="sListSeq" title="Order By WinLoss Differential" VALUE="W" <% IF sListSeq="W" THEN response.write("checked") %>>
			</td>
			<td style="border-style:none;"  align=center>Risk Factor
				<input type=radio NAME="sListSeq" title="Order By Risk Factor" VALUE="R" <% IF sListSeq="R" THEN response.write("checked") %>>
			</td>
			<td style="border-style:none;"  align=center>Blended
				<input type=radio NAME="sListSeq" title="Order By Blended Method = WinLoss - Risk Factor" VALUE="B" <% IF sListSeq="B" THEN response.write("checked") %>>
			</td>
  	</TR>
  	<%
ELSE
		%>
		<TR><td colspan=8 style="border-style:none;">&nbsp;</td></TR>
		<%
END IF


	%>
	<TR>
		<td colspan=2 align="center" style="border-style:none;">
			<input type="submit" style="width:9em" value="Update Display" title="Submit and reset this form">
		</td>
		<td colspan=2 style="border-style:none;" align=center>
	  	<input type="submit" name="WhatAction" value="Return to Menu" style="background-color:yellow;">
		</td>
		<%

		IF Session("adminmenulevel")>=50 THEN 
				%>
				<td colspan=2 style="border-style:none;"align=center>
		  		<input type="submit" name="WhatAction" value="Control Panel">
				</td>
				<td colspan=2 style="border-style:none;"align=center>&nbsp;</td>
				<%
		ELSE 
				%>
				<td colspan=4 style="border-style:none;"align=center>&nbsp;</td>
				<%
		END IF 
		%>	
	</TR>

</TABLE>

</form>
<%



END SUB


' ---------------------
   SUB SetDefaultTeam
' ---------------------

	Set rs=Server.CreateObject("ADODB.recordset")

	sSQL = " SELECT TOP 1 TTNo AS DefTeam"
	sSQL = sSQL + " 	FROM usawsrank.IAC_USTeamCombos AS TC"
	sSQL = sSQL + " 	WHERE TourID='"&sTourID&"'"	
	sSQL = sSQL + " ORDER BY Win_Fact DESC"	
	rs.open sSQL, SConnectionToTRATable
	
	IF NOT rs.eof THEN
		sTeamNo=rs("DefTeam") 
	ELSE
		sTeamNo=999999
		'response.end
	END IF


	rs.close

END SUB



' ---------------------
  SUB BuildMinDrop
' ---------------------


 


END SUB



' ---------------------
  SUB TeamNoDropDown
' ---------------------

Dim sTotalTeams


Set rs=Server.CreateObject("ADODB.recordset")

sSQL="SELECT COUNT(TTNo) AS TotalTeams FROM (SELECT DISTINCT TTNo FROM "&TeamXTableName&") AS TX"
'response.write(sSQL)
'response.end

rs.open sSQL, SConnectionToTRATable

'response.write("<br>"&sSQL)
'response.end

sTotalTeams=0
IF NOT rs.eof THEN sTotalTeams=rs("TotalTeams")
rs.close

%>
<select name="sTeamNo" style="width:6em;" <%=PulldownStatus%>><%

FOR iCounter = 1 TO sTotalTeams STEP 1

	IF iCounter = sTeamNo THEN
		response.write("<option value = """&iCounter&""" SELECTED>"&iCounter&"</option>")
	ELSE
		response.write("<option value = """&iCounter&""">"&iCounter&"</option>")
	END IF
NEXT %>
</select><%


END SUB



' ---------------------
  SUB foreignsummary
' ---------------------


Set rs=Server.CreateObject("ADODB.recordset")

sSQL = "SELECT RWT.First, RWT.Last, RWT.Sex, RWT.Fed" 

sSQL = sSQL + "	, COALESCE(SLCount,0) AS [# of<br>Slalom<br>Scores]"
sSQL = sSQL + "	, COALESCE(TRCount,0) AS [# of<br>Trick<br>Scores]"
sSQL = sSQL + "	, COALESCE(JUCount,0) AS [# of<br>Jump<br>Scores]"

sSQL = sSQL + "	, CASE WHEN SL_Phantom_Added='Y' THEN '*' + CAST( CAST(COALESCE(SLTrialsScr,0) AS decimal(6,2)) AS varchar(10)) + '*' ELSE CAST( CAST(COALESCE(SLTrialsScr,0) AS decimal(6,2)) AS varchar(10)) END AS [Slalom<br>Trials]" 
' sSQL = sSQL + "	, CASE WHEN SL_Phantom_Added='Y' THEN CAST(  CAST(COALESCE(SLTrialsScr,0) AS decimal(6,2)) AS VARCHAR(6)) END AS [Slalom<br>Trials]" 

sSQL = sSQL + "	, CAST(COALESCE(SLScore1,0) AS decimal(6,2)) AS [Slalom<br>Score1]"
sSQL = sSQL + " , CAST(COALESCE(SLScore2,0) AS decimal(6,2)) AS [Slalom<br>Score2]" 
sSQL = sSQL + "	, CAST(COALESCE(SLScore3,0) AS decimal(6,2)) AS [Slalom<br>Score3]"

sSQL = sSQL + "	, CASE WHEN TR_Phantom_Added='Y' THEN '*' + CAST( CAST(COALESCE(TRTrialsScr,0) AS INTEGER) AS varchar(10)) + '*' ELSE CAST( CAST(COALESCE(TRTrialsScr,0) AS INTEGER) AS varchar(10)) END AS [Trick<br>Trials]" 
' sSQL = sSQL + "	, CAST(COALESCE(TRTrialsScr,0) AS INTEGER) AS [Trick<br>Trials]" 
sSQL = sSQL + "	, CAST(COALESCE(TRScore1,0) AS INTEGER) AS [Tricks<br>Score1]"
sSQL = sSQL + "	, CAST(COALESCE(TRScore2,0) AS INT) AS [Tricks<br>Score2]"
sSQL = sSQL + " , CAST(COALESCE(TRScore3,0) AS INTEGER) AS [Tricks<br>Score3]"

sSQL = sSQL + "	, CASE WHEN JU_Phantom_Added='Y' THEN '*' + CAST( CAST(COALESCE(JUTrialsScr,0) AS decimal(6,2)) AS varchar(10)) + '*' ELSE CAST( CAST(COALESCE(JUTrialsScr,0) AS decimal(6,2)) AS varchar(10)) END AS [Jump<br>Trials]" 
' sSQL = sSQL + "	, CAST(COALESCE(JUTrialsScr,0) AS decimal(6,2)) AS [Jump<br>Trials]" 
sSQL = sSQL + "	, CAST(COALESCE(JUScore1,0) AS decimal(6,2)) AS [Jump<br>Score1]"
sSQL = sSQL + "	, CAST(COALESCE(JUScore2,0) AS decimal(6,2)) AS [Jump<br>Score2]"
sSQL = sSQL + "	, CAST(COALESCE(JUScore3,0) AS decimal(6,2)) AS [Jump<br>Score3]"

sSQL = sSQL + "	FROM usawsrank.IAC_RegWorldTeams AS RWT"

sSQL = sSQL + "	WHERE LEFT(TourID,6)='"&sTourID&"'"

sSQL = sSQL + " ORDER BY Fed, Sex, RWT.Last, RWT.First"	

' response.write("<br>"&sSQL)
' response.end

rs.open sSQL, SConnectionToTRATable

DisplayResult

rs.close

IF LEFT(sTourID,6)="07S151" OR LEFT(sTourID,6)="09S083" THEN %>
	<font size=3 color="#FFFFFF"><b>2007 Foreign Athlete Score Formula</b></font>
	<br><font size=2 color="#FFFFFF">Trials Score = TopScore/2 + MedianScore/2</font><%
END IF



END SUB




' ---------------------
  SUB indivscores
' ---------------------

' --- Displays the score detail, best score of touraments and score summary for each member --

%>

<br>
<form action="/rankings/<%=ThisFileName%>" method="post">
  

  <input type="hidden" name="sPassword" value="<%= sPassword %>">
  <input type="hidden" name="sTourID" value="<%= sTourID %>">
  <input type="hidden" name="sOverlay" value="<%= sOverlay %>">


<TABLE class="innertable" Align=center WIDTH=1000px height=100>
  <TR>
	<td colspan=8 style="border-style:none;">
		<font color="<%=TextColor2%>" size="3"><b><%=PageTitle%></b></font>
		<br>
		<font color="<%=TextColor1%>" size="2"><b><%=PageSubTitle%></b></font>
	</td>
  </TR>
  <TR>
	<td colspan=8 style="border-style:none;">&nbsp;</td>
  </TR>



  <TR>

	<td style="border-style:none;"  align=right>Select Member</td>
	<td colspan=3 style="border-style:none;" align=left>

	<%


	sSQL = "SELECT DISTINCT FirstName AS [First Name], LastName AS [Last Name], MemberID" 
	sSQL = sSQL + "	FROM usawsrank.IAC_TrialsSkiers AS ES, "&MemberShortTableName&" M"
	sSQL = sSQL + "	WHERE RIGHT(es.MemberID,8)=m.PersonID AND TourID='"&sTourID&"'"

	Set rs=Server.CreateObject("ADODB.recordset")
	rs.open sSQL, SConnectionToTRATable


	%><select name="sMemberID" style="width:20em"><%

	IF NOT rs.eof THEN 
  		rs.movefirst

	  	DO WHILE NOT rs.eof
			IF TRIM(rs("MemberID")) = sMemberID THEN %>
				<option value="<%=rs("MemberID")%>" selected><%=rs("MemberID")%> - <%=rs("First Name")%>&nbsp;<%=rs("Last Name")%> </option><br><%
			ELSE %>
				<option value="<%=rs("MemberID")%>"><%=rs("MemberID")%> - <%=rs("First Name")%>&nbsp;<%=rs("Last Name")%> </option><br><%
			END IF	

			rs.moveNEXT
		LOOP
	END IF  %>
	</select>
	
	</td>
	<td colspan=2 style="border-style:none;"align=center>
	  <input type="submit" name="WhatAction" value="Display Scores">
	</td>
	<td colspan=2 style="border-style:none;"align=center>
	  <input type="submit" name="WhatAction" value="Return to Trials">
	</td>

  </TR>

  <TR>
	<td colspan=8 style="border-style:none;">&nbsp;</td>
  </TR>
 
</TABLE>

</form>
<%

IF sMemberID<>"" THEN
	displayindivscores
ELSE %>
<br><br>
<center><h1>Select Member and Press Display Scores</h1></center><%

END IF


END SUB





'-----------------------
  SUB displayindivscores
'-----------------------

Set rs=Server.CreateObject("ADODB.recordset")

sSQL = "SELECT es.TourID"

sSQL = sSQL + ", CASE WHEN ts.TName IS NOT NULL THEN ts.TName"
sSQL = sSQL + " WHEN it.Tour_Desc IS NOT NULL THEN  it.Tour_Desc"	
sSQL = sSQL + " ELSE '** Not Defined **' END AS [Tournament Name]"

' sSQL = sSQL + "	, COALESCE(TName,'** Foreign - Manual Add **') AS [Tournament Name]"
sSQL = sSQL + ", Site_ID, Event, Round, Score, Place, Best_Tour_Score" 
sSQL = sSQL + "	FROM usawsrank.IAC_EventScores AS es"
sSQL = sSQL + "	LEFT JOIN"
sSQL = sSQL + "	 (SELECT TName, TournAppID FROM sanctions.dbo.TSchedul) ts"
sSQL = sSQL + "	ON LEFT(ts.TournAppID,6)=LEFT(es.TourID,6)"
sSQL = sSQL + " LEFT JOIN usawsrank.IAC_Tournaments_IWWF it ON it.TourID=es.TourID"
sSQL = sSQL + "	WHERE MemberID='"&sMemberID&"'"
sSQL = sSQL + " ORDER BY Event, es.TourID, Round"	

'response.write("sSQL = "&sSQL)
'response.end

rs.open sSQL, SConnectionToTRATable

%>
<br>
<TABLE class="innertable" Align=center WIDTH=1000px>
  <TR>
	<th colspan=8><font color="#FFFFFF">All Scores</font></th>
  </TR>

  <TR>
	<td width=10%><FONT SIZE="1"><b>TourID</b></font></td>
	<td><FONT SIZE="1"><b>Tournament Name</b></font></td>
	<td><FONT SIZE="1"><b>Site_ID</b></font></td>
	<td><FONT SIZE="1"><b>Event</b></font></td>
	<td><FONT SIZE="1"><b>Round</b></font></td>
	<td><FONT SIZE="1"><b>Score</b></font></td>
	<td><FONT SIZE="1"><b>Place</b></font></td>
	<td><FONT SIZE="1"><b>Best Tour</b></font></td>
	
  </TR><%

  DO WHILE NOT rs.eof %>	
	  <TR>
		<td><FONT SIZE="1"><%=rs("TourID")%></font></td>
		<td><FONT SIZE="1"><%=rs("Tournament Name")%></font></td>
		<td><FONT SIZE="1"><%=rs("Site_ID")%></font></td>
		<td><FONT SIZE="1"><%=rs("Event")%></font></td>
		<td><FONT SIZE="1"><%=rs("Round")%></font></td>
		<td><FONT SIZE="1"><%=rs("Score")%></font></td>
		<td><FONT SIZE="1"><%=rs("Place")%></font></td>
		<td><FONT SIZE="1"><%=rs("Best_Tour_Score")%></font></td>
	  </TR><%

	rs.movenext
   LOOP %>	

</TABLE>
<br><%


rs.close



'--- Builds Top Score of Each Tournament Set ---

Set rs=Server.CreateObject("ADODB.recordset")

sSQL = "SELECT es.TourID, Site_ID, Event, Round, Score, Place" 
sSQL = sSQL + ", CASE WHEN ts.TName IS NOT NULL THEN ts.TName"
sSQL = sSQL + " WHEN it.Tour_Desc IS NOT NULL THEN  it.Tour_Desc"	
sSQL = sSQL + " ELSE '** Not Defined **' END AS TName"
sSQL = sSQL + ", TournAppID, TDateS, TDateE" 

sSQL = sSQL + "	FROM usawsrank.IAC_EventScores AS ES"
sSQL = sSQL + "	LEFT JOIN sanctions.dbo.TSchedul ts ON LEFT(ts.TournAppID,6)=LEFT(es.TourID,6)"
sSQL = sSQL + " LEFT JOIN usawsrank.IAC_Tournaments_IWWF it ON it.TourID=es.TourID"

sSQL = sSQL + "	WHERE MemberID='"&sMemberID&"' AND Best_Tour_Score IN ('Y','1','2','3')"

sSQL = sSQL + "	ORDER BY Event, Score DESC"


rs.open sSQL, SConnectionToTRATable


%>
<TABLE class="innertable" Align=center WIDTH=1000px>
  <TR>
	<th colspan=8><font color="#FFFFFF">Best Scores By Tournament</font></th>
  </TR>

  <TR>
	<td width=10%><FONT SIZE="1"><b>TourID</b></font></td>
	<td><FONT SIZE="1"><b>Tournament Name</b></font></td>
	<td><FONT SIZE="1"><b>Site_ID</b></font></td>
	<td><FONT SIZE="1"><b>Dates</b></font></td>
	<td><FONT SIZE="1"><b>Event</b></font></td>
	<td><FONT SIZE="1"><b>Score</b></font></td>
	<td><FONT SIZE="1"><b>Place</b></font></td>
  </TR><%

  DO WHILE NOT rs.eof %>	
	  <TR>
		<td><FONT SIZE="1"><%=rs("TourID")%></font></td>
		<td><FONT SIZE="1"><%=rs("TName")%></font></td>
		<td><FONT SIZE="1"><%=rs("Site_ID")%></font></td>
		<td><FONT SIZE="1"><%=rs("TDateS")%> - <%=rs("TDateE")%></font></td>
		<td><FONT SIZE="1"><%=rs("Event")%></font></td>
		<td><FONT SIZE="1"><%=rs("Score")%></font></td>
		<td><FONT SIZE="1"><%=rs("Place")%></font></td>
	  </TR><%

	rs.movenext
   LOOP %>	

</TABLE>
<br><%

rs.close






'--- Builds Summary Set - One line per event ---


Set rs=Server.CreateObject("ADODB.recordset")

sSQL = "SELECT First, Last," 
sSQL = sSQL + "	Coalesce(SLScore1,0) AS SLScore1," 
sSQL = sSQL + "	Coalesce(SLScore2,0) AS SLScore2,"
sSQL = sSQL + "	Coalesce(SLScore3,0) AS SLScore3,"
sSQL = sSQL + "	Coalesce(SLScore_Median,0) AS SLScore_Median,"
sSQL = sSQL + "	Coalesce(SlNumSco,0) AS SlNumSco,"
sSQL = sSQL + "	Coalesce(SlNumTour,0) AS SlNumTour,"

sSQL = sSQL + "	CAST(Coalesce(TRScore1,0) AS INTEGER) AS TRScore1," 
sSQL = sSQL + "	CAST(Coalesce(TRScore2,0) AS INTEGER) AS TRScore2,"
sSQL = sSQL + "	CAST(Coalesce(TRScore3,0) AS INTEGER) AS TRScore3,"
sSQL = sSQL + "	CAST(Coalesce(TRScore_Median,0) AS INTEGER) AS TRScore_Median,"
sSQL = sSQL + "	CAST(Coalesce(TrNumSco,0) AS INTEGER) AS TrNumSco,"
sSQL = sSQL + "	CAST(Coalesce(TrNumTour,0) AS INTEGER) AS TrNumTour,"

sSQL = sSQL + "	CAST(Coalesce(JUScore1,0) AS decimal(6,2)) AS JUScore1," 
sSQL = sSQL + "	CAST(Coalesce(JUScore2,0) AS decimal(6,2)) AS JUScore2,"
sSQL = sSQL + "	CAST(Coalesce(JUScore3,0) AS decimal(6,2)) AS JUScore3,"
sSQL = sSQL + "	CAST(Coalesce(JUScore_Median,0) AS decimal(6,2)) AS JUScore_Median,"
sSQL = sSQL + "	Coalesce(JuNumSco,0) as JuNumSco,"
sSQL = sSQL + "	Coalesce(JuNumTour,0) as JuNumTour"

sSQL = sSQL + "	FROM USAWSRank.IAC_SkierSummary"
sSQL = sSQL + "	WHERE MemberID='"&sMemberID&"'"

rs.open sSQL, SConnectionToTRATable

sSLScore1=rs("SLScore1")
sSLScore2=rs("SLScore2")
sSLScore3=rs("SLScore3")
sSlMedAll=rs("SLScore_Median")
sSlNumSco=rs("SlNumSco")
sSlNumTour=rs("SlNumTour")

sTrScore1=rs("TRScore1")
sTrScore2=rs("TRScore2")
sTrScore3=rs("TRScore3")
sTrMedAll=rs("TRScore_Median")
sTrNumSco=rs("TrNumSco")
sTrNumTour=rs("TrNumTour")

sJuScore1=rs("JUScore1")
sJuScore2=rs("JUScore2")
sJuScore3=rs("JUScore3")
sJuMedAll=rs("JUScore_Median")
sJuNumSco=rs("JuNumSco")
sJuNumTour=rs("JuNumTour")



Set rsTS=Server.CreateObject("ADODB.recordset")

sSQL = "SELECT MemberID, TourID, Sex," 
sSQL = sSQL + "	Coalesce(SlTrialsScr,0) as SlTrialsScr," 
sSQL = sSQL + "	Coalesce(TrTrialsScr,0) as TrTrialsScr," 
sSQL = sSQL + "	Coalesce(JuTrialsScr,0) as JuTrialsScr" 
sSQL = sSQL + "	FROM USAWSRank.IAC_TrialsSkiers"
sSQL = sSQL + "	WHERE MemberID='"&sMemberID&"' AND LEFT(TourID,6)='"&LEFT(sTourID,6)&"'"

rsTS.open sSQL, SConnectionToTRATable

sSlTrialsScr=rsTS("SLTrialsScr")
sTrTrialsScr=rsTS("trTrialsScr")
sJuTrialsScr=rsTS("JuTrialsScr")



%>
<TABLE class="innertable" Align=center style="text-align:center;" WIDTH=1000px>
  <TR>
	<th colspan=8><font color="#FFFFFF">Summary Scores By Event </font></th>
  </TR>

  <TR>
	<td width=10%><FONT SIZE="1"><b>Event</b></font></td>
	<td width=12%><FONT SIZE="1"><b># Scores</b></font></td>
	<td width=12%><FONT SIZE="1"><b># Tours</b></font></td>
	<td width=12%><FONT SIZE="1"><b>Trials</b></font></td>
	<td width=12%><FONT SIZE="1"><b>Score1</b></font></td>
	<td width=12%><FONT SIZE="1"><b>Score2</b></font></td>
	<td width=12%><FONT SIZE="1"><b>Score3</b></font></td>
	<td width=12%><FONT SIZE="1"><b>Median</b></font></td>

  </TR><%

'  DO WHILE NOT rs.eof 
	IF rs("SlNumTour")>0 THEN %>	
	  <TR>
		<td><FONT SIZE="1">Slalom</font></td>
		<td><FONT SIZE="1"><%=sSlNumSco%></font></td>
		<td><FONT SIZE="1"><%=sSlNumTour%></font></td>
		<td><FONT SIZE="1"><%=sSlTrialsScr%></font></td>
		<td><FONT SIZE="1"><%=sSLScore1%></font></td>
		<td><FONT SIZE="1"><%=sSLScore2%></font></td>
		<td><FONT SIZE="1"><%=sSLScore3%></font></td>
		<td><FONT SIZE="1"><%=sSlMedAll%></font></td>

	  </TR><%
	END IF  

	IF sTrNumTour>0 THEN %>	
	  <TR>
		<td><FONT SIZE="1">Tricks</font></td>
		<td><FONT SIZE="1"><%=sTrNumSco%></font></td>
		<td><FONT SIZE="1"><%=sTrNumTour%></font></td>
		<td><FONT SIZE="1"><%=sTrTrialsScr%></font></td>
		<td><FONT SIZE="1"><%=sTrScore1%></font></td>
		<td><FONT SIZE="1"><%=sTrScore2%></font></td>
		<td><FONT SIZE="1"><%=sTrScore3%></font></td>
		<td><FONT SIZE="1"><%=sTrMedAll%></font></td>
	  </TR><%
	END IF  

	IF sJuNumTour>0 THEN %>	
	  <TR>
		<td><FONT SIZE="1">Jump</font></td>
		<td><FONT SIZE="1"><%=sJuNumSco%></font></td>
		<td><FONT SIZE="1"><%=sJuNumTour%></font></td>
		<td><FONT SIZE="1"><%=sJuTrialsScr%></font></td>
		<td><FONT SIZE="1"><%=sJuScore1%></font></td>
		<td><FONT SIZE="1"><%=sJuScore2%></font></td>
		<td><FONT SIZE="1"><%=sJuScore3%></font></td>		
		<td><FONT SIZE="1"><%=sJuMedAll%></font></td>
	  </TR><%
	END IF  

'	rs.movenext

'   LOOP 

%>	

</TABLE>
<br><br><%



rs.close
rsTS.close

DisplayInfoFooter


END SUB






' ---------------------
  SUB ustrialssummary
' ---------------------

OpenCon

Set rs=Server.CreateObject("ADODB.recordset")

IF Whataction="delfromover" THEN
	sSQL = "UPDATE TS SET Marked='N' FROM usawsrank.IAC_TrialsSkiers AS TS WHERE MemberID='"&sMemberID&"' AND LEFT(TourID,6)='"&LEFT(sTourID,6)&"'"
  	con.execute(sSQL)
	CloseCon
ELSEIF Whataction="addtoover" THEN
	sSQL = "UPDATE TS SET Marked='Y' FROM usawsrank.IAC_TrialsSkiers AS TS WHERE MemberID='"&sMemberID&"' AND LEFT(TourID,6)='"&LEFT(sTourID,6)&"'"
  	con.execute(sSQL)
	CloseCon
END IF


'response.write("<br>sSQL="&sSQL)
'response.end



sSQL = "	SELECT Marked, MEM.FirstName AS First, MEM.LastName AS Last, TS.MemberID, TS.Sex," 

sSQL = sSQL + "	  CASE WHEN COALESCE(SL_Suppress_Score,' ')='Y' THEN '*' + CAST(COALESCE(CAST(SLCount AS Integer),0) AS nVarChar) + '*' ELSE CAST(COALESCE(CAST(SLCount AS Integer),0) AS nVarChar) END AS [# of<br>Slalom<br>Scores],"
sSQL = sSQL + "	  CASE WHEN COALESCE(TR_Suppress_Score,' ')='Y' THEN '*' + CAST(COALESCE(CAST(TRCount AS Integer),0) AS nVarChar) + '*' ELSE CAST(COALESCE(CAST(TRCount AS Integer),0) AS nVarChar) END AS [# of<br>Trick<br>Scores],"
sSQL = sSQL + "	  CASE WHEN COALESCE(JU_Suppress_Score,' ')='Y' THEN '*' + CAST(COALESCE(CAST(JUCount AS Integer),0) AS nVarChar) + '*' ELSE CAST(COALESCE(CAST(JUCount AS Integer),0) AS nVarChar) END AS [# of<br>Jump<br>Scores],"

' sSQL = sSQL + "	  SL_Suppress_Score AS [# of<br>Slalom<br>Suppress],"
' sSQL = sSQL + "	  TR_Suppress_Score AS [# of<br>Trick<br>Suppress],"
' sSQL = sSQL + "	  JU_Suppress_Score AS [# of<br>Jump<br>Suppress],"

sSQL = sSQL + "		CAST(COALESCE(SLTrialsScr,0) AS decimal(6,2)) AS [Slalom<br>Trials], "
sSQL = sSQL + "		COALESCE(SLScore1,0) AS [Slalom<br>1st],"
sSQL = sSQL + "		COALESCE(SLScore2,0) AS [Slalom<br>2nd],"
sSQL = sSQL + "		COALESCE(SLScore3,0) AS [Slalom<br>3rd],"

sSQL = sSQL + "		CAST(COALESCE(TRTrialsScr,0) AS INTEGER) AS [Trick<br>Trials],"
sSQL = sSQL + "		CAST(COALESCE(TRScore1,0) AS INTEGER) AS [Trick<br>1st]," 
sSQL = sSQL + "		CAST(COALESCE(TRScore2,0) AS INTEGER) AS [Trick<br>2nd],"
sSQL = sSQL + "	  CAST(COALESCE(TRScore3,0) AS INTEGER) AS [Trick<br>3rd],"

sSQL = sSQL + "	  CAST(COALESCE(JUTrialsScr,0) AS decimal(6,2)) AS [Jump<br>Trials],"
sSQL = sSQL + "		COALESCE(JUScore1,0) AS [Jump<br>1st]," 
sSQL = sSQL + "		COALESCE(JUScore2,0) AS [Jump<br>2nd],"
sSQL = sSQL + "		COALESCE(JUScore3,0) AS [Jump<br>3rd]" 

sSQL = sSQL + "		FROM usawsrank.IAC_TrialsSkiers AS TS"
	
sSQL = sSQL + "		LEFT JOIN"
sSQL = sSQL + "			(SELECT FirstName, LastName, PersonID"
sSQL = sSQL + "				FROM "&MemberShortTableName&") AS MEM"
sSQL = sSQL + "		ON RIGHT(TS.MemberID,8)=MEM.PersonID"

sSQL = sSQL + "	WHERE LEFT(TourID,6)='"&sTourID&"'"
sSQL = sSQL + "	ORDER BY Sex, MEM.LastName, MEM.FirstName"

'response.write(sSQL)
'response.end

rs.open sSQL, SConnectionToTRATable


DisplayResult

rs.close

DisplayInfoFooter 

END SUB




' ------------------------
   SUB DisplayInfoFooter 
' ------------------------
%>
<table align=center class="innertable" Align=center WIDTH=1020px height=100>
  <tr>
    <td><%

IF LEFT(sTourID,6)="07S151" THEN %>
	<font size=3 color="<%=Textcolor2%>"><b>2007 US Athlete Score Formula</b></font>
	<br><font size=2 color="<%=Textcolor1%>">Team Trials Score = Score1/2 + Score2/4 + Score3/4</font>
	<br><font size=2 color="<%=Textcolor1%>">Score1 = High Score from Team Trials</font>
	<br><font size=2 color="<%=Textcolor1%>">Score2 = Middle Score from Team Trials</font>
	<br><font size=2 color="<%=Textcolor1%>">Score3 = Low Score from Team Trials</font>

	<br><br><center><font size=2 color="white"><a href="http://www.usawaterski.org/rankings/IAC/Team_Selection_2007_Elite_SUN_DOCS_v2.pdf" target="_blank">2007 Elite Trials Results</a></font></center> 
	<br><center><font size=2 color="white"><a href="http://www.usawaterski.org/rankings/IAC/Method_for_Selecting_Alternates_DRAFT.pdf" target="_blank">Alternate Selection (DRAFT)</a></font></center> 
	<%


ELSEIF LEFT(sTourID,6)="09S083" THEN %>
	<font size=3 color="<%=Textcolor2%>"><b>2009 US Athlete Score Formula</b></font>
	<br><font size=2 color="<%=Textcolor1%>">Team Trials Score = Score1/2 + Score2/4 + Score3/4</font>
	<br><font size=2 color="<%=Textcolor1%>">Score1 = High Score from Team Trials</font>
	<br><font size=2 color="<%=Textcolor1%>">Score2 = Low Score from Team Trials</font>
	<br><font size=2 color="<%=Textcolor1%>">Score3 = Median Score from period 6-19-2008 to 6-18-2009</font>		<br><br><center><font size=2 color="white"><a href="http://www.usawaterski.org/rankings/iwsf/SQL_Code/IAC_Programming_Summary_Step_by_Step_2009.pdf" target="_blank">Programming Summary</a></font></center>
	<br><center><font size=2 color="white"><a href="http://www.usawaterski.org/rankings/IAC/Method_for_Selecting_Alternates_DRAFT.pdf" target="_blank">Alternate Selection (DRAFT)</a></font></center> 
	<%
ELSEIF LEFT(sTourID,6)="11TEAM" THEN %>
	<font size=3 color="<%=Textcolor2%>"><b>2011 US Athlete Score Formula</b></font>
	<br><font size=2 color="<%=Textcolor1%>">Team Trials Score = TopScore/2 + 2ndScore/4 + MedScore/4</font>
	<br><font size=2 color="<%=Textcolor1%>">Qualifying Period from Jan1-2011 to June20-2011</font>
	<br><font size=2 color="<%=Textcolor1%>">Second Highest Score not at same Site_ID as High Score)</font>
	<br><font size=2 color="<%=Textcolor1%>">Median Score includes of all L & R rounds </font><%
	y=1
	IF y=2 THEN %>
		<br><br><center><font size=2 color="white"><a href="http://www.usawaterski.org/rankings/iwsf/SQL_Code/IAC_Programming_Summary_Step_by_Step_2011.pdf" target="_blank">Programming Summary</a></font></center><%
	END IF %> 
	<br><center><font size=2 color="white"><a href="http://www.usawaterski.org/rankings/IAC/Method_for_Selecting_Alternates_DRAFT.pdf" target="_blank">Alternate Selection (DRAFT)</a></font></center> 
	<%
ELSEIF LEFT(sTourID,6)="13TEAM" THEN %>
	<font size=3 color="<%=Textcolor2%>"><b>2013 US Athlete Score Formula</b></font>
	<br><font size=2 color="<%=Textcolor1%>">USA Team Trials Score = TopScore*0.6 + 2ndScore*0.4</font>
	<br><font size=2 color="<%=Textcolor1%>">Qualifying Method = 2 Round Team Trials</font>
	<%
	y=2
	IF y=2 THEN %>
		<br><br><center><font size=2 color="white"><a href="http://www.usawaterski.org/rankings/iwsf/SQL_Code/IAC_2013_Elite_WorldTeam_Selection_Process_Summary.pdf" target="_blank">Programming Summary</a></font></center><%
	END IF %> 
	<br><center><font size=2 color="white"><a href="http://www.usawaterski.org/rankings/IAC/Method_for_Selecting_Alternates_DRAFT.pdf" target="_blank">Alternate Selection (DRAFT)</a></font></center> 
	<%

ELSEIF LEFT(sTourID,6)="15TEAM" THEN %>
	<font size=3 color="<%=Textcolor2%>"><b>2015 Trials Score Formula</b></font>
	<br><font size=2 color="<%=Textcolor1%>"><b>USA Athletes:</b> Team Trials Score = (TopScore*0.6) + (2ndScore*0.4) performed in 2 Round 2015 Team Elite Team Trials</font>
	<br><font size=2 color="<%=Textcolor1%>"><b>Foreign Athletes:</b> IWWF Raw Scores (May 2015 WRL)- Team Trials Score = (TopScore*0.6) + (Median Score*0.4)</font>
	<%
	y=2
	IF y=2 THEN %>
		<br><br><center><font size=2 color="white"><a href="http://www.usawaterski.org/rankings/iwsf/SQL_Code/2015/IAC 2015 - Elite Team Selection Process Summary.pdf" target="_blank">Programming Summary 2015</a></font></center><%
	END IF 
	ELSEIF LEFT(sTourID,6)="17TEAM" THEN %>
	<font size=3 color="<%=Textcolor2%>"><b>2017 Trials Score Formula</b></font>
	<br><font size=2 color="<%=Textcolor1%>"><b>USA Athletes:</b> Team Trials Score = (TopScore*0.5) + Avg(Top + 2nd + 3rd Scores)*0.5 performed between 1-1-17 and 7-23-17</font>
	<br><font size=2 color="<%=Textcolor1%>"><b>Foreign Athletes:</b> IWWF Raw Scores (Dynamic WRL)- Team Trials Score = (TopScore*0.5) + Avg(Top + 2nd + 3rd Scores)*0.5 performed between 8-1-16 and 7-23-17</font>
	<%
	y=2
	IF y=2 THEN %>
		<br><br><center><font size=2 color="white"><a href="http://www.usawaterski.org/rankings/IAC/2017/IAC_2017_Elite_Team_Selection_Process_Summary.pdf" target="_blank">Programming Summary 2017</a></font></center><%
	END IF %>  
	<br><center><font size=2 color="white"><a href="http://www.usawaterski.org/rankings/IAC/Method_for_Selecting_Alternates_DRAFT.pdf" target="_blank">Alternate Selection (DRAFT)</a></font></center> 
	<%

END IF  %>
    </td>
  </tr>
</table><%
	

END SUB



' ---------------------
  SUB teamdetail
' ---------------------

' --- Displays Score and Overall Detail for a selected TeamX

Set rs=Server.CreateObject("ADODB.recordset")

sSQL = "SELECT TX.MemberID,"

sSQL = sSQL + "	LastName, FirstName, Sex"
sSQL = sSQL + "	, SLPos, TRPos, JUPos"

sSQL = sSQL + "	, SLTrialsScr AS [Slalom<br>Score]"
sSQL = sSQL + "	, CAST(TRTrialsScr AS INT) AS [Trick<br>Score]"
sSQL = sSQL + "	, JUTrialsScr AS [Jump<br>Score]"  

sSQL = sSQL + "	, IndOvrSL AS [Slalom<br>Overall]"
sSQL = sSQL + "	, IndOvrTR AS [Trick<br>Overall]"
sSQL = sSQL + "	, IndOvrJU AS [Jump<br>Overall]"


sSQL = sSQL + "	FROM "&TeamXTableName&" AS TX"
	
sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		(SELECT FirstName, LastName, PersonID"
sSQL = sSQL + "			FROM "&MemberShortTableName&") AS MEM" 
sSQL = sSQL + "	ON RIGHT(TX.MemberID,8)=MEM.PersonID" 

	
sSQL = sSQL + " WHERE LEFT(TX.TourID,6)='"&sTourID&"' AND TTNo='"&sTeamNo&"'"

sSQL = sSQL + " AND TX.Fed='USA'"

sSQL = sSQL + " ORDER BY Sex, MEM.LastName, MEM.FirstName"

rs.open sSQL, SConnectionToTRATable


'response.write(rs.eof)
'response.write(sSQL)

'response.end


SubProcess="usdetail"

IF rs.eof THEN	
	%><center><font size="3" color="<%=TextColor3%>">No Data In <%=TeamXTableName%></font></center><%

ELSE
	DisplayResult


  JumpOverUSAResultsStandAlone="Y"
	IF NOT(JumpOverUSAResultsStandAlone="Y") THEN

		rs.close

		sSQL = " SELECT TTNo, 'USA' AS Fed, TeamOvrSL, TeamOvrTR, TeamOvrJU, TeamGrand, Win_Fact, Rsk_Fact"
		sSQL = sSQL + " 	FROM usawsrank.IAC_USTeamCombos AS TC"
		sSQL = sSQL + " 	WHERE TTNo='"&sTeamNo&"' AND TourID='"&sTourID&"'"
		rs.open sSQL, SConnectionToTRATable


			%>
			<TABLE class="innertable" Align=center WIDTH=1000px>
	  		<TR>
					<th colspan=7><font color="#FFFFFF">Overall Results - US TeamX No: <%=sTeamNo%></font></th>
	  		</TR>

	  		<TR>
					<th><FONT SIZE="1" color="#FFFFFF">Fed</font></th>
					<th><FONT SIZE="1" color="#FFFFFF">Slalom</font></th>
					<th><FONT SIZE="1" color="#FFFFFF">Tricks</font></th>
					<th><FONT SIZE="1" color="#FFFFFF">Jump</font></th>
					<th><FONT SIZE="1" color="#FFFFFF">Total</font></th>
					<th><FONT SIZE="1" color="#FFFFFF">Win Factor</font></th>
					<th><FONT SIZE="1" color="#FFFFFF">Risk Factor</font></th>
	  		</TR>

	  		<TR>
					<td><FONT SIZE="1"><%=rs("Fed")%></font></td>
					<td><FONT SIZE="1"><%=rs("TeamOvrSL")%></font></td>
					<td><FONT SIZE="1"><%=rs("TeamOvrTR")%></font></td>
					<td><FONT SIZE="1"><%=rs("TeamOvrJU")%></font></td>
					<td><FONT SIZE="1"><%=rs("TeamGrand")%></font></td>
					<td><FONT SIZE="1"><%=rs("Win_Fact")%></font></td>
					<td><FONT SIZE="1"><%=rs("Rsk_Fact")%></font></td>
  			</TR>

			</TABLE>
			<br><br>
			<%
 			
 			rs.close
	END IF


	' --- Prints 1 line of summary for each FED
	ForeignTeamResults

	' --- Displays all 6 skiers for each FED with score and Overall detail
	ForeignDetail


END IF


END SUB



' -------------------
  SUB foreigndetail
' -------------------



' --- Layout of USA score data ---
'sSQL = sSQL + "	LastName, FirstName, Sex"
'sSQL = sSQL + "	, SLPos, TRPos, JUPos"
'sSQL = sSQL + "	, SLTrialsScr AS [Slalom<br>Score], CAST(TRTrialsScr AS INT) AS [Trick<br>Score], JUTrialsScr AS [Jump<br>Score]"  
'sSQL = sSQL + "	, IndOvrSL AS [Slalom<br>Overall], IndOvrTR AS [Trick<br>Overall], IndOvrJU AS [Jump<br>Overall]"


sSQL="SELECT DISTINCT Fed FROM usawsrank.IAC_RegWorldTeams AS TX"
Set rsFed=Server.CreateObject("ADODB.recordset")
rsFed.open sSQL, SConnectionToTRATable

DO WHILE NOT rsFed.eof

		FedSelected=rsFed("Fed")
	
		sSQL = "SELECT TX.MemberID, Last, First, TX.Sex"
		sSQL = sSQL + ", SLPos, TRPos, JUPos"

		sSQL = sSQL + "	, CASE WHEN SL_Phantom_Added='Y' THEN '*' + CAST( CAST(COALESCE(SLTrialsScr,0) AS decimal(6,2)) AS varchar(10)) + '*' ELSE CAST( CAST(COALESCE(SLTrialsScr,0) AS decimal(6,2)) AS varchar(10)) END AS [Slalom<br>Score]" 		
		sSQL = sSQL + "	, CASE WHEN TR_Phantom_Added='Y' THEN '*' + CAST( CAST(COALESCE(TRTrialsScr,0) AS INTEGER) AS varchar(10)) + '*' ELSE CAST( CAST(COALESCE(TRTrialsScr,0) AS INTEGER) AS varchar(10)) END AS [Trick<br>Score]" 
		sSQL = sSQL + "	, CASE WHEN JU_Phantom_Added='Y' THEN '*' + CAST( CAST(COALESCE(JUTrialsScr,0) AS decimal(6,2)) AS varchar(10)) + '*' ELSE CAST( CAST(COALESCE(JUTrialsScr,0) AS decimal(6,2)) AS varchar(10)) END AS [Jump<br>Score]" 

		sSQL = sSQL + ", IndOvrSL AS [Slalom<br>Overall]"
		sSQL = sSQL + ", IndOvrTR AS [Trick<br>Overall]"
		sSQL = sSQL + ", IndOvrJU AS [Jump<br>Overall]"

		sSQL = sSQL + "	FROM "&TeamXTableName&" AS TX"

		sSQL = sSQL + "	LEFT JOIN" 
		sSQL = sSQL + "		(SELECT TourID, First, Last, Sex, MemberID, SL_Phantom_Added, TR_Phantom_Added, JU_Phantom_Added"
		sSQL = sSQL + "			FROM usawsrank.IAC_RegWorldTeams) AS RWT" 
		sSQL = sSQL + "	ON  LEFT(RWT.TourID,6)=LEFT(TX.TourID,6) AND TX.MemberID=RWT.MemberID" 
	
		sSQL = sSQL + " WHERE LEFT(TX.TourID,6)='"&sTourID&"' AND TTNo='"&sTeamNo&"'"


		IF FedSeleted="FFF" THEN
				sSQL = sSQL + " AND TX.Fed<>'USA'"
		ELSEIF TRIM(FedSelected)<>"" THEN
				sSQL = sSQL + " AND TX.Fed='"&FedSelected&"'"
		END IF

		sSQL = sSQL + " ORDER BY TX.Sex, RWT.Last, RWT.First"

		'response.write(sSQL)
		'response.end
		
		
		Set rs=Server.CreateObject("ADODB.recordset")
		rs.open sSQL, SConnectionToTRATable


		SubProcess="intdetail"
		DisplayResult

	rs.close

	rsFed.movenext
LOOP

END SUB




' ----------------------------
  SUB ForeignTeamCandidates
' ----------------------------  

sSQL = "SELECT rs.First, rs.Last, rs.Country, rs.Sex, rs.HomeFed_ID"
sSQL = sSQL + "	, COALESCE(SLScore,0) AS SLScore, COALESCE(TRScore,0) AS TRScore, COALESCE(JUScore,0) AS JUScore"
sSQL = sSQL + "	FROM"
sSQL = sSQL + "	("
sSQL = sSQL + "		SELECT First, Last, Country, Sex, MAX(HomeFed_ID) AS HomeFed_ID"
sSQL = sSQL + "			FROM [usawsrank].[IAC_IWSFRaw2017]"
sSQL = sSQL + "				WHERE Country IN ('CAN','AUS','FRA','GBR','BLR','ITA')"
sSQL = sSQL + "					AND End_Date>='08/01/2016'"
sSQL = sSQL + "					AND (" 
sSQL = sSQL + "						(Sex='M' AND SLScore>60)" 
sSQL = sSQL + "							OR" 
sSQL = sSQL + "						(Sex='M' AND TRScore>8000)"
sSQL = sSQL + "							OR" 
sSQL = sSQL + "						(Sex='M' AND JUScore>55)"
sSQL = sSQL + "							OR" 
sSQL = sSQL + "						(Sex='F' AND SLScore>50)"
sSQL = sSQL + "							OR" 
sSQL = sSQL + "						(Sex='F' AND TRScore>6500)"
sSQL = sSQL + "							OR" 
sSQL = sSQL + "						(Sex='F' AND JUScore>45) )"
sSQL = sSQL + "				GROUP BY First, Last, Country, Sex"
sSQL = sSQL + "		) rs"		
sSQL = sSQL + "		LEFT JOIN"
sSQL = sSQL + "			( SELECT First, Last, Country, Sex, MAX(HomeFed_ID) AS HomeFed_ID"
sSQL = sSQL + "				, MAX(SLScore) AS SLScore, MAX(TRScore) AS TRScore, MAX(JUScore) AS JUScore"
sSQL = sSQL + "					FROM [usawsrank].[IAC_IWSFRaw2017] rs"
sSQL = sSQL + "						WHERE Country IN ('CAN','AUS','FRA','GBR','BLR','ITA')"
sSQL = sSQL + "					GROUP BY First, Last, Country, Sex"
sSQL = sSQL + "			) sc"
sSQL = sSQL + "		ON sc.First=rs.First AND sc.Last=rs.Last AND sc.Country=rs.Country AND sc.Sex=rs.Sex"
'  AND sc.HomeFed_ID=rs.HomeFed_ID"	

sSQL = sSQL + "		ORDER BY rs.Country, rs.Sex, rs.Last, rs.First"  


'response.write(sSQL)
'response.end



Set rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable




SubProcess="foreigncandidates"
DisplayResult

	rs.close


END SUB




'----------------------------
  SUB ForeignTeamResults
'----------------------------

' --- Finds TOP 3 and 2,3,4 Team Overall scores

Set rsFedTot=Server.CreateObject("ADODB.recordset")


sSQL =  "	SELECT TTB.TourID, TTB.TTNo, TTB.Fed," 
sSQL = sSQL + " CAST(COALESCE(SLTeamTot,0) AS decimal(6,2)) AS SLTeamTot,"
sSQL = sSQL + " CAST(COALESCE(TRTeamTot,0) AS decimal(6,2)) AS TRTeamTot," 
sSQL = sSQL + " CAST(COALESCE(JUTeamTot,0) AS decimal(6,2)) AS JUTeamTot,"
sSQL = sSQL + "	CAST(COALESCE((SLTeamTot+TRTeamTot+JUTeamTot),0) AS decimal(6,2)) AS TeamGrand,"
sSQL = sSQL + "	COALESCE(Rsk_Fact,0) AS Rsk_Fact"

sSQL = sSQL + "		FROM "

sSQL = sSQL + "		(SELECT TourID, TTNo, Fed"
sSQL = sSQL + "			FROM "&TeamXTableName
sSQL = sSQL + "			GROUP BY TourID, TTNo, Fed"
sSQL = sSQL + "		) AS TTB"
 
sSQL = sSQL + "		LEFT JOIN"
sSQL = sSQL + "			(SELECT TourID, TTNo, Fed, SUM(IndOvrSL) AS SLTeamTot"
sSQL = sSQL + "				FROM "&TeamXTableName
sSQL = sSQL + "				WHERE SLPos<=3"
sSQL = sSQL + "				GROUP BY TourID, TTNo, Fed) AS SL"
sSQL = sSQL + "		ON SL.TourID=TTB.TourID AND SL.TTNo=TTB.TTNo AND SL.Fed=TTB.Fed"

sSQL = sSQL + "		LEFT JOIN"
sSQL = sSQL + "			(SELECT TourID, TTNo, Fed, SUM(IndOvrSL) AS SLAltTot"
sSQL = sSQL + "				FROM "&TeamXTableName
sSQL = sSQL + "				WHERE SLPos>=2 AND SLPos<=4"
sSQL = sSQL + "				GROUP BY TourID, TTNo, Fed) AS SL2"
sSQL = sSQL + "		ON SL2.TourID=TTB.TourID AND SL2.TTNo=TTB.TTNo AND SL2.Fed=TTB.Fed"


sSQL = sSQL + "		LEFT JOIN"
sSQL = sSQL + "			(SELECT TourID, TTNo, Fed, SUM(IndOvrTR) AS TRTeamTot"
sSQL = sSQL + "				FROM "&TeamXTableName
sSQL = sSQL + "				WHERE TRPos<=3"
sSQL = sSQL + "				GROUP BY TourID, TTNo, Fed) AS TR"
sSQL = sSQL + "		ON TR.TourID=TTB.TourID AND TR.TTNo=TTB.TTNo AND TR.Fed=TTB.Fed"

sSQL = sSQL + "		LEFT JOIN"
sSQL = sSQL + "			(SELECT TourID, TTNo, Fed, SUM(IndOvrTR) AS TRAltTot"
sSQL = sSQL + "				FROM "&TeamXTableName
sSQL = sSQL + "				WHERE TRPos>=2 AND TRPos<=4"
sSQL = sSQL + "				GROUP BY TourID, TTNo, Fed) AS TR2"
sSQL = sSQL + "		ON TR2.TourID=TTB.TourID AND TR2.TTNo=TTB.TTNo AND TR2.Fed=TTB.Fed"


sSQL = sSQL + "		LEFT JOIN"
sSQL = sSQL + "			(SELECT TourID, TTNo, Fed, SUM(IndOvrJU) AS JUTeamTot"
sSQL = sSQL + "				FROM "&TeamXTableName
sSQL = sSQL + "				WHERE JUPos<=3"
sSQL = sSQL + "				GROUP BY TourID, TTNo, Fed) AS JU"
sSQL = sSQL + "		ON JU.TourID=TTB.TourID AND JU.TTNo=TTB.TTNo AND JU.Fed=TTB.Fed"

sSQL = sSQL + "		LEFT JOIN"
sSQL = sSQL + "			(SELECT TourID, TTNo, Fed, SUM(IndOvrJU) AS JUAltTot"
sSQL = sSQL + "				FROM "&TeamXTableName
sSQL = sSQL + "				WHERE JUPos>=2 AND JUPos<=4"
sSQL = sSQL + "				GROUP BY TourID, TTNo, Fed) AS JU2"
sSQL = sSQL + "		ON JU2.TourID=TTB.TourID AND JU2.TTNo=TTB.TTNo AND JU2.Fed=TTB.Fed"

sSQL = sSQL + " 	LEFT JOIN"
sSQL = sSQL + " 		( SELECT TourID, TTNo, Rsk_Fact"
sSQL = sSQL + " 				FROM usawsrank.IAC_USTeamCombos AS TC"
sSQL = sSQL + " 			WHERE TTNo='"&sTeamNo&"' AND TourID='"&sTourID&"') AS RFC"
sSQL = sSQL + "		ON RFC.TourID=TTB.TourID AND RFC.TTNo=TTB.TTNo AND TTB.Fed='USA'"


'sSQL = sSQL + "	WHERE LEFT(TTB.TourID,6)='"&sTourID&"' AND TTB.Fed<>'USA' AND TTB.TTNo='"&sTeamNo&"'"
sSQL = sSQL + "	WHERE LEFT(TTB.TourID,6)='"&sTourID&"' AND TTB.TTNo='"&sTeamNo&"'"

'sSQL = sSQL + "	ORDER BY TTB.TourID, TTB.TTNo, TTB.TeamGrand DESC"
sSQL = sSQL + "	ORDER BY TTB.TeamGrand DESC, TTB.TourID, TTB.Fed "


'response.write(sSQL)
'response.end


rsFedTot.open sSQL, SConnectionToTRATable

%>
<TABLE class="innertable" Align=center WIDTH=1200px>
  <TR>
	<th colspan=7><font color="#FFFFFF">Overall Summary By Event - International Teams </font></th>
  </TR>

  <TR>
	<td><FONT SIZE="1">Fed</font></td>
	<td><FONT SIZE="1">Slalom</font></td>
	<td><FONT SIZE="1">Tricks</font></td>
	<td><FONT SIZE="1">Jump</font></td>
	<td><FONT SIZE="1">Total</font></td>
	<td><FONT SIZE="1">Win Factor</font></td>
	<td><FONT SIZE="1">Risk Factor</font></td>
  </TR><%

	FedCounter=0
	HighestTeamScore=0
  DO WHILE NOT rsFedTot.eof 
  	CellColor=""
		FedCounter=FedCounter+1
		IF FedCounter=1 THEN HighestTeamScore=rsFedTot("TeamGrand")
  	IF rsFedTot("Fed")="USA" THEN CellColor="#D6ECF2"
  		'"#B0E0E6"
  		
  	WinDif = CDbl(rsFedTot("TeamGrand"))-CDbl(HighestTeamScore)
  		
  	%>	
	  <TR>
		<td style="background-color:<%=CellColor%>"><FONT SIZE="1"><%=rsFedTot("Fed")%></font></td>
		<td style="background-color:<%=CellColor%>"><FONT SIZE="1"><%=formatnumber(rsFedTot("SLTeamTot"),2)%></font></td>
		<td style="background-color:<%=CellColor%>"><FONT SIZE="1"><%=formatnumber(rsFedTot("TRTeamTot"),2)%></font></td>
		<td style="background-color:<%=CellColor%>"><FONT SIZE="1"><%=formatnumber(rsFedTot("JUTeamTot"),2)%></font></td>
		<td style="background-color:<%=CellColor%>"><FONT SIZE="1"><%=formatnumber(rsFedTot("TeamGrand"),2)%></font></td>
		<td style="background-color:<%=CellColor%>"><FONT SIZE="1"><%=formatnumber(WinDif,1) %></font></td>
		<td style="background-color:<%=CellColor%>"><FONT SIZE="1"><%=formatnumber(rsFedTot("Rsk_Fact"),1)%></font></td>
	  </TR><%

	rsFedTot.movenext
   LOOP %>	

</TABLE>
<br><%

rsFedTot.close


END SUB


' ----------------
   SUB TeamList
' ---------------


' --- Determines first team in list Win_Fact and Rsk_Fact for determining differential to subsequent values

Dim onColumn, FirstWin, FirstRisk
FirstWin=0
FirstRisk=0

onColumn="yes"
IF onColumn="yes" AND sListSeq<>"B" THEN
	Set rs=Server.CreateObject("ADODB.recordset")

	sSQL = " SELECT Top 1 Rsk_Fact AS FirstRisk, Win_Fact AS FirstWin"
	sSQL = sSQL + "		FROM usawsrank.IAC_USTeamCombos AS TC"
	sSQL = sSQL + " WHERE LEFT(TC.TourID,6)='"&sTourID&"'"
	sSQL = sSQL + " AND SNo_Scores>='"&sSLMin&"' AND TNo_Scores>='"&sTRMin&"' AND JNo_Scores>='"&sJUMin&"'"
	sSQL = sSQL + " AND TotFemale>=2 AND TotMale>=2"

	IF sListSeq="W" THEN
			sSQL = sSQL + " ORDER BY Win_Fact DESC"
	ELSEIF sListSeq="B" THEN  		
			sSQL = sSQL + " ORDER BY ((1 * Win_Fact)-RSk_Fact) DESC"
	ELSE
			sSQL = sSQL + " ORDER BY Rsk_Fact" 		
	END IF 

	rs.open sSQL, SConnectionToTRATable

	IF NOT rs.eof THEN
		FirstWin=rs("FirstWin")
		FirstRisk=rs("FirstRisk")
	END IF
	rs.close

END IF



'--- Begin Listing query build ---

Set rs=Server.CreateObject("ADODB.recordset")

sSQL = " SELECT Top 300 TTNo, MK1, M1.First+' '+M1.Last AS [Athlete 1], MK2, M2.First+' '+M2.Last AS [Athlete 2], MK3, M3.First+' '+M3.Last AS [Athlete 3]"
sSQL = sSQL + ", MK4, M4.First+' '+M4.Last AS [Athlete 4], MK5, M5.First+' '+M5.Last AS [Athlete 5], MK6, M6.First+' '+M6.Last AS [Athlete 6]"
sSQL = sSQL + ", COALESCE(TotMale,0) AS [M], COALESCE(TotFemale,0) AS [F]"
sSQL = sSQL + ", CAST(SNo_Scores AS INTEGER) AS [#<br>SL], CAST(TNo_Scores AS INTEGER) AS [#<br>TR], CAST(JNo_Scores AS INTEGER) AS [#<br>JU]"
sSQL = sSQL + ", CAST(COALESCE(Win_Fact,0) AS DECIMAL(6,2)) AS [Win<br>Loss<br>Diff], CAST(COALESCE(Rsk_Fact,0) AS DECIMAL(6,2)) AS [Risk<br>Factor]"
sSQL = sSQL + ", CAST(COALESCE(Win_Fact-Rsk_Fact,0) AS DECIMAL(6,2)) AS [Blended]"


'--- Note: Runs differential from Risk Factor and Win Loss for 1st team - Mannually Input
IF FirstWin<>0 AND FirstRisk<>0 THEN
		sSQL = sSQL + ", CAST(COALESCE(Win_Fact-("&FirstWin&"),0) AS DECIMAL(6,2)) AS [DWin],"
		sSQL = sSQL + " CAST(COALESCE(("&FirstRisk&")-Rsk_Fact,0) AS DECIMAL(6,2)) AS [DRisk],"
		sSQL = sSQL + " CASE WHEN Win_Fact-("&FirstWin&") BETWEEN -0.01 AND 0.01 THEN 0 ELSE CAST(COALESCE(-(("&FirstRisk&")-Rsk_Fact)/(Win_Fact-("&FirstWin&")),0) AS DECIMAL(6,2)) END AS [Risk<br>Reward<br>Ratio]"

		' sSQL = sSQL +" CASE WHEN Win_Fact-("&FirstWin&")<>0 THEN CAST(COALESCE(-(("&FirstRisk&")-Rsk_Fact)/(Win_Fact-("&FirstWin&")),0) AS DECIMAL(6,2)) ELSE 0 END AS [Risk<br>Reward<br>Ratio]"
		' CASE WHEN Win_Fact-("&FirstWin&")<>0 THEN -(("&FirstRisk&")-Rsk_Fact)/(Win_Fact-("&FirstWin&")) ELSE 0 END
END IF
 
	
sSQL = sSQL + "		FROM usawsrank.IAC_USTeamCombos AS TC"
	

sSQL = sSQL + " 		LEFT JOIN" 
sSQL = sSQL + " 			(SELECT FirstName AS First, LastName  AS Last,  PersonID"
sSQL = sSQL + " 				FROM "&MemberShortTableName&") AS M1"
sSQL = sSQL + " 		ON M1.PersonID=RIGHT(TC.Member1,8)"

sSQL = sSQL + " 		LEFT JOIN" 
sSQL = sSQL + " 			(SELECT TourID, MemberID, Marked AS MK1"
sSQL = sSQL + " 				FROM usawsrank.IAC_TrialsSkiers) AS TS1"
sSQL = sSQL + " 		ON TS1.MemberID=TC.Member1 AND TS1.TourID=TC.TourID"


sSQL = sSQL + " 		LEFT JOIN" 
sSQL = sSQL + " 			(SELECT FirstName AS First, LastName  AS Last,  PersonID"
sSQL = sSQL + " 				FROM "&MemberShortTableName&") AS M2"
sSQL = sSQL + " 		ON M2.PersonID=RIGHT(TC.Member2,8)"

sSQL = sSQL + " 		LEFT JOIN" 
sSQL = sSQL + " 			(SELECT TourID, MemberID, Marked AS MK2"
sSQL = sSQL + " 				FROM usawsrank.IAC_TrialsSkiers) AS TS2"
sSQL = sSQL + " 		ON TS2.MemberID=TC.Member2 AND TS2.TourID=TC.TourID"


sSQL = sSQL + " 		LEFT JOIN "
sSQL = sSQL + " 			(SELECT FirstName AS First, LastName  AS Last,  PersonID"
sSQL = sSQL + " 				FROM "&MemberShortTableName&") AS M3"
sSQL = sSQL + " 		ON M3.PersonID=RIGHT(TC.Member3,8)"

sSQL = sSQL + " 		LEFT JOIN" 
sSQL = sSQL + " 			(SELECT TourID, MemberID, Marked AS MK3"
sSQL = sSQL + " 				FROM usawsrank.IAC_TrialsSkiers) AS TS3"
sSQL = sSQL + " 		ON TS3.MemberID=TC.Member3 AND TS3.TourID=TC.TourID"


sSQL = sSQL + " 		LEFT JOIN" 
sSQL = sSQL + " 			(SELECT FirstName AS First, LastName  AS Last,  PersonID"
sSQL = sSQL + " 				FROM "&MemberShortTableName&") AS M4"
sSQL = sSQL + " 		ON M4.PersonID=RIGHT(TC.Member4,8)"

sSQL = sSQL + " 		LEFT JOIN" 
sSQL = sSQL + " 			(SELECT TourID, MemberID, Marked AS MK4"
sSQL = sSQL + " 				FROM usawsrank.IAC_TrialsSkiers) AS TS4"
sSQL = sSQL + " 		ON TS4.MemberID=TC.Member4 AND TS4.TourID=TC.TourID"


sSQL = sSQL + " 		LEFT JOIN "
sSQL = sSQL + " 			(SELECT FirstName AS First, LastName  AS Last,  PersonID"
sSQL = sSQL + " 				FROM "&MemberShortTableName&") AS M5"
sSQL = sSQL + " 		ON M5.PersonID=RIGHT(TC.Member5,8)"

sSQL = sSQL + " 		LEFT JOIN" 
sSQL = sSQL + " 			(SELECT TourID, MemberID, Marked AS MK5"
sSQL = sSQL + " 				FROM usawsrank.IAC_TrialsSkiers) AS TS5"
sSQL = sSQL + " 		ON TS5.MemberID=TC.Member5 AND TS5.TourID=TC.TourID"


sSQL = sSQL + " 		LEFT JOIN" 
sSQL = sSQL + " 			(SELECT FirstName AS First, LastName  AS Last,  PersonID"
sSQL = sSQL + " 				FROM "&MemberShortTableName&") AS M6"
sSQL = sSQL + " 		ON M6.PersonID=RIGHT(TC.Member6,8)"

sSQL = sSQL + " 		LEFT JOIN" 
sSQL = sSQL + " 			(SELECT TourID, MemberID, Marked AS MK6"
sSQL = sSQL + " 				FROM usawsrank.IAC_TrialsSkiers) AS TS6"
sSQL = sSQL + " 		ON TS6.MemberID=TC.Member6 AND TS6.TourID=TC.TourID"

			
sSQL = sSQL + " WHERE LEFT(TC.TourID,6)='"&sTourID&"'"
sSQL = sSQL + " AND SNo_Scores>='"&sSLMin&"' AND TNo_Scores>='"&sTRMin&"' AND JNo_Scores>='"&sJUMin&"'"
sSQL = sSQL + " AND TotFemale>=2 AND TotMale>=2"



IF sListSeq="W" THEN
		sSQL = sSQL + " ORDER BY Win_Fact DESC"
ELSEIF sListSeq="B" THEN  		
		sSQL = sSQL + " ORDER BY ((1 * Win_Fact)-RSk_Fact) DESC"
ELSEIF sListSeq="RR" THEN  		
		'  sSQL = sSQL + " ORDER BY -(("&FirstRisk&")-Rsk_Fact)/(Win_Fact-("&FirstWin&")) DESC"
		sSQL = sSQL + " ORDER BY CASE WHEN Win_Fact-("&FirstWin&") BETWEEN -0.01 AND 0.01 THEN 0 ELSE -(("&FirstRisk&")-Rsk_Fact)/(Win_Fact-("&FirstWin&")) END DESC"
ELSE
		sSQL = sSQL + " ORDER BY Rsk_Fact" 		
END IF 

'response.write(sSQL)
'response.end


rs.open sSQL, SConnectionToTRATable





IF rs.eof THEN
	'response.write("<br>End of File<br><br><br>")
	'response.write(sSQL)
	NoresultsToDisplay
	
ELSE
	DisplayResult
END IF

END SUB



' --------------------
  SUB DisplayMedianList
' --------------------

Set rs=Server.CreateObject("ADODB.recordset")


sSQL = "SELECT MemberID, Last, First, Sex, SLScore3, TRScore3, JUScore3, SLNumTour, TRNumTour, JUNumTour, SLNumSco, TRNumSco, JUNumSco" 
sSQL = sSQL + " FROM usawsrank.IAC_SkierSummary"
sSQL = sSQL + " 	WHERE Country='USA' ORDER BY LAST, First"

rs.open sSQL, SConnectionToTRATable

DisplayResult


END SUB



' --------------------
  SUB TeamStatistics
' --------------------   



sSQL = " 	SELECT COALESCE(SNo_Scores,0) AS [# Slalom Scores]"
sSQL = sSQL + " 	, COALESCE(TNo_Scores,0) AS [# Trick Scores]"
sSQL = sSQL + " 	, COALESCE(JNo_Scores,0) AS [# Jump Scores]"
' sSQL = sSQL + " 	, COUNT(SNo_Scores) AS [Slalom Count]"
' sSQL = sSQL + " 	, COUNT(TNo_Scores) AS [Trick Count]"
' sSQL = sSQL + " 	, COUNT(JNo_Scores) AS [Jump Count]"
sSQL = sSQL + " 	,COUNT(*) AS [Total Number of Teams]" 

sSQL = sSQL + " 	FROM [usawsrank].[IAC_USTeamCombos]"
sSQL = sSQL + " 	WHERE SNo_Scores>=3 AND TNo_Scores>=3 AND JNo_Scores>=3"
sSQL = sSQL + " 	GROUP BY SNo_Scores, TNo_Scores, JNo_Scores"
sSQL = sSQL + " 	ORDER BY SNo_Scores, TNo_Scores, JNo_Scores"

Set rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable

DisplayResult

rs.close


END SUB



' ---------------------
  SUB DisplayRatioList
' ---------------------

Set rs=Server.CreateObject("ADODB.recordset")

sSQL = "SELECT TS.MemberID, TS.Sex," 
'sSQL = sSQL + " 	TS.SLTrialsScr, SS.SLBest, SS.SLMedAll, (SS.SLBest+SS.SLMedAll)/2 AS IWSFSL,"
sSQL = sSQL + " 	TS.SLTrialsScr, (SS.SLBest+SS.SLMedAll)/2 AS IWSFSL,"
sSQL = sSQL + " 	  CASE WHEN TS.SLTrialsScr>=0.5 * SS.SLBest AND (SS.SLBest+SS.SLMedAll)>0 THEN TS.SLTrialsScr/((SS.SLBest+SS.SLMedAll)/2)*100 ELSE 0 END AS SLScrRatio,"

'sSQL = sSQL + " 	TS.TRTrialsScr, SS.TRBest, SS.TRMedAll, (SS.TRBest+SS.TRMedAll)/2 AS IWSFTR,"
sSQL = sSQL + " 	TS.TRTrialsScr, (SS.TRBest+SS.TRMedAll)/2 AS IWSFTR,"
sSQL = sSQL + " 	  CASE WHEN TS.TRTrialsScr>=0.5 * SS.TRBest AND (SS.TRBest+SS.TRMedAll)>0 THEN TS.TRTrialsScr/((SS.TRBest+SS.TRMedAll)/2)*100 ELSE 0 END AS TRScrRatio,"

'sSQL = sSQL + " 	TS.JUTrialsScr, SS.JUBest, SS.JUMedAll, (SS.JUBest+SS.JUMedAll)/2 AS IWSFJU,"
sSQL = sSQL + " 	TS.JUTrialsScr, (SS.JUBest+SS.JUMedAll)/2 AS IWSFJU,"
sSQL = sSQL + " 	  CASE WHEN TS.JUTrialsScr>=0.5 * SS.JUBest AND (SS.JUBest+SS.JUMedAll)>0 THEN TS.JUTrialsScr/((SS.JUBest+SS.JUMedAll)/2)*100 ELSE 0 END AS JUScrRatio"


sSQL = sSQL + " 	FROM usawsrank.IAC_TrialsSkiers AS TS"
	
sSQL = sSQL + " 	JOIN "
sSQL = sSQL + " 		(SELECT MemberID, First, Last, SLBest, SLMedAll, TRBest, TRMedAll, JUBest, JUMedAll"
sSQL = sSQL + " 			FROM usawsrank.IAC_SkierSummary ) AS SS"
sSQL = sSQL + " 	ON TS.MemberID=SS.MemberID"

sSQL = sSQL + " WHERE LEFT(TS.TourID,6)='"&sTourID&"'"	

rs.open sSQL, SConnectionToTRATable

DisplayResult


END SUB



' ----------------------
  SUB NoResultsToDisplay
' ----------------------

%>
<table class="innertable" align=center width=1000px>
  <tr><td style="border-style:none;">&nbsp;</td></tr>
  <tr>
	<td align="center" style="border-style:none;"><font size="3" color="red">No Results For This Tournament</font></center></td>
  </tr>
  <tr><td style="border-style:none;">&nbsp;</td></tr>

</table>	

<%


END SUB




' -----------------
  SUB ControlPanel
' -----------------




Set rs=Server.CreateObject("ADODB.recordset")
sSQL = " SELECT IntFileName"
sSQL = sSQL + " FROM usawsrank.IAC_Control"
sSQL = sSQL + " WHERE LEFT(TourID,6)='"&sTourID&"'"
rs.open sSQL, SConnectionToTRATable

IF NOT(rs.EOF) THEN
		sIntFileName=rs("IntFileName")
END IF
rs.close


Set rs=Server.CreateObject("ADODB.recordset")

'sSQL = " SELECT *"

sSQL = " SELECT *,"
sSQL = sSQL + "	(SELECT COUNT(Last) FROM usawsrank."&sIntFileName&") AS IWSFScrs,"
sSQL = sSQL + "	(SELECT COUNT(Last) FROM usawsrank.IAC_EventScores) AS EventScrs,"
sSQL = sSQL + "	(SELECT COUNT(Last) FROM usawsrank.IAC_SkierSummary) AS SkiSum,"
sSQL = sSQL + "	(SELECT COUNT(TourID) FROM usawsrank.IAC_TrialsSkiers WHERE LEFT(TourID,6)='"&sTourID&"') AS Trials,"
sSQL = sSQL + "	(SELECT COUNT(TourID) FROM "&TeamXTableName&" WHERE LEFT(TourID,6)='"&sTourID&"') AS IndScores,"
sSQL = sSQL + "	(SELECT COUNT(TourID) FROM usawsrank.IAC_USTeamCombos WHERE LEFT(TourID,6)='"&sTourID&"') AS Combos"

sSQL = sSQL + " FROM usawsrank.IAC_Control"
sSQL = sSQL + " WHERE LEFT(TourID,6)='"&sTourID&"'"


'response.write("<br><br>")
'response.write(sSQL)
'response.end

rs.open sSQL, SConnectionToTRATable


'response.write("<br><br>Found Select = ")
'response.write(NOT(rs.EOF))


' --- Reads the file name from IAC_Control then looks up the date range
IF NOT rs.eof THEN
	Dim RawDataFile
	RawDataFile=rs("IntFileName") 

	' --- Reads Lowest and highest date range from Current IWSF Raw Score File
	Set rsRD=Server.CreateObject("ADODB.recordset")
	sSQL = " SELECT MIN(CAST(End_Date AS DateTime)) AS RawMinDate, MAX(CAST(End_Date AS DateTime)) AS RawMaxDate" 
	sSQL = sSQL + " FROM usawsrank."&RawDataFile
	sSQL = sSQL + " WHERE End_Date<>'01/01/1900'"	


' --- PROBLEM IS THAT UNTIL THERE IS DATA IN THE RAW DATA FILE - THERE IS NO END DATE ---

'response.write("<br><br>")
'response.write(sSQL)
'response.end

	rsRD.open sSQL, SConnectionToTRATable


	IF NOT rsRD.eof THEN
		sRawMinDate = rsRD("RawMinDate")
		sRawMaxDate = rsRD("RawMaxDate")	
	ELSE
		sRawMinDate = "No IWWF Scores Found"
		sRawMaxDate = "No IWWF Scores Found"	
	END IF

	sUSBeginDate = rs("USBeginDate")
	sUSEndDate = rs("USEndDate")
	sIntBeginDate = rs("IntBeginDate")
	sIntEndDate = rs("IntEndDate")
	sIntFileName = rs("IntFileName")

	sIWSFScrs = rs("IWSFScrs")
	sEventScrs = rs("EventScrs")
	sSkiSum = rs("SkiSum")
	sTrials = rs("Trials")
	sIndScores = rs("IndScores")
	sUSTrialsScores = rs("USTrialsScores")
	sForeignScores = rs("ForeignScores")
	sCombos = rs("Combos")

	sSkierSummary = rs("SkierSummary")
	sTeamIndividual = rs("TeamIndividual")
	sTeamSummary = rs("TeamSummary")
	sTeamCombos = rs("TeamCombos")
	sUSTrialsScores = rs("USTrialsScores")
	sForeignScores = rs("ForeignScores")
ELSE
	sUSBeginDate = "01/01/1900"
	sUSEndDate = "01/01/1900"
	sIntBeginDate = "01/01/1900"
	sIntEndDate = "01/01/1900"
	sIntFileName = "Not Defined"

	sIWSFScrs = 0
	sEventScrs = 0
	sSkiSum = 0
	sTrials = 0
	sIndScores = 0
	sUSTrialsScores = 0
	sForeignScores = 0
	sCombos = 0

	sSkierSummary = 0
	sTeamIndividual = 0
	sTeamSummary = 0
	sTeamCombos = 0
	sUSTrialsScores = 0
	sForeignScores = 0

END IF



%>

<br>

<form action="/rankings/<%=ThisFileName%>" method="post">


<TABLE class="innertable" Align=center WIDTH=1000px height=100>
  <TR>	
	<td colspan=8 >
		<br>
		<font color="<%=TextColor2%>" size="3"><b><%=PageTitle%></b></font>
		<br>
		<font color="<%=TextColor1%>" size="2"><b><%=PageSubTitle%></b></font>
	</td>
  </TR>
  <TR><td colspan=8 >&nbsp;</td></TR>

  <TR>
    <td colspan=4 align="right">
	<font color="<%=TextColor1%>" size=2>Default Tournament</font>
    </td>
    <td colspan=4 align="left" valign="center">
	<% BuildTrialsDropDownList %>
    </td>
  </TR>

	  <TR>
	    <td colspan=4 align="center">
		<font color="<%=TextColor1%>" size=2><b>US Date Range</font>
	    </td>
	    <td colspan=4 align="center">
		<font color="<%=TextColor1%>" size=2><b>International Date Range</b></font>
	    </td>
	  </TR>

	  <TR>
	    <td colspan=2 align="right" width=15%>
		<font color="<%=TextColor1%>" size=2>Begin&nbsp;</font>
	    </td>
	    <td colspan=2 align="left" width=35%>
		<input type="text" name="sUSBeginDate" value="<%= sUSBeginDate %>">
	    </td>
	    <td colspan=2 align="right" width=15% >
		<font color="<%=TextColor1%>" size=2>Begin&nbsp;</font>
	    </td>
	    <td colspan=2 align="left" width=35%>
		<input type="text" name="sIntBeginDate" value="<%= sIntBeginDate %>">
	    </td>
	  </TR>

	  <TR>	
	    <td colspan=2 align="right" >
		<font color="<%=TextColor1%>" size=2>End&nbsp;</font>
	    </td>
	    <td colspan=2 align="left">
		<input type="text" name="sUSEndDate" value="<%= sUSEndDate %>">
	    </td>
	    <td colspan=2 align="right" >
		<font color="<%=TextColor1%>" size=2>End&nbsp;</font>
	    </td>
	    <td colspan=2 align="left">
		<input type="text" name="sIntEndDate" value="<%= sIntEndDate %>">
	    </td>
	  </TR>

	  <TR>	
	    <td colspan=4>&nbsp;</td>
	    <td colspan=2 align="right" >
		<font color="<%=TextColor1%>" size=2>File Name&nbsp;</font>
	    </td>
	    <td colspan=2 align="left">
		<input type="text" name="sIntFileName" value="<%= sIntFileName %>">
	    </td>
	  </TR> 

	  <TR> 	
	    <td colspan=8 align=center >
		<input type="submit" style=width:20em name="WhatAction" value="Save Changes">
	    </td>
	  </TR>

	  <TR><td colspan=8>&nbsp;</td></TR>

	  <TR>
	    <td colspan=4 align="center">
		<font color="<%=TextColor1%>" size=2><b>&nbsp;</b></font>
	    </td>

	    <td colspan=4 align="center">
		<font color="<%=TextColor1%>" size=2><b>IWSF Raw Score Data File</b></font>
	    </td>
	  </TR>

	   <TR>	
	    <td colspan=4>&nbsp;</td>
	    <td colspan=2 align="right" >
		<font color="<%=TextColor1%>" size=2>First&nbsp;</font>
	    </td>
	    <td colspan=2 align="left">
		<font color="<%=TextColor2%>" size=2>&nbsp;<%= sRawMinDate %></font>
	    </td>
	  </TR>

	   <TR>	
	    <td colspan=4>&nbsp;</td>
	    <td colspan=2 align="right" >
		<font color="<%=TextColor1%>" size=2>Last&nbsp;</font>
	    </td>
	    <td colspan=2 align="left">
		<font color="<%=TextColor2%>" size=2>&nbsp;<%= sRawMaxDate %></font>
	    </td>
	  </TR>




  <TR><td colspan=8>&nbsp;</td></TR>
  <TR><td colspan=8>&nbsp;</td></TR>
  <TR>
    <td colspan=8 >
	<font color="<%=TextColor3%>" size=3><b>Recalc Values</b></font>
	<br>
    </td>
  </TR>

  <TR>
    <td colspan=2 >
	<font color="<%=TextColor1%>" size=2><b>Last Recalced</b></font>
    </td>
    <td colspan=2 >
	<font color="<%=TextColor1%>" size=2><b>Operation</b></font>
    </td>
    <td colspan=2 >
	<font color="<%=TextColor1%>" size=2><b>Table</b></font>
    </td>
    <td colspan=2 >
	<font color="<%=TextColor1%>" size=2><b>Quantity</b></font>
    </td>
  </TR>

  <TR>
	<td colspan=2 align=center >Manual Import</td>
	<td colspan=2 align=center >
	  <input type="submit" style=width:12em name="WhatAction" value="IWSF Import" disabled>
	</td>
    <td colspan=2>IWSF Scores</td>
    <td colspan=2><font color="<%=TextColor2%>" size=2><%= sIWSFScrs %></font></td>
  </TR>

  <TR>
	<td colspan=2 align=center ><%= sSkierSummary %>
	</td>
	<td colspan=2 align=center >
	  <input type="submit" style=width:12em name="WhatAction" value="Median & Top All">
	</td>
    <td colspan=2>Event Scores</td>
    <td colspan=2><font color="<%=TextColor2%>" size=2><%= sEventScrs %></font></td>

  </TR>


  <TR>
	<td colspan=2 align=center ><%= sForeignScores %></td>
	<td colspan=2 align=center >
	  <input type="submit" style=width:12em name="WhatAction" value="Foreign Skier Calcs">
	</td>
    <td colspan=2>Skier Summary</td>
    <td colspan=2><font color="<%=TextColor2%>" size=2><%= sSkiSum %></font></td>

  </TR>

  <TR>
	<td colspan=2 align=center ><%= sUSTrialsScores %></td>
	<td colspan=2 align=center >
	  <input type="submit" style=width:12em name="WhatAction" value="USA Skier Calcs">
	</td>
    <td colspan=2>Trials Skiers</td>
    <td colspan=2><font color="<%=TextColor2%>" size=2><%= sTrials %></font></td>


  </TR>
  <TR>
	<td colspan=2 align=center ><%= sTeamCombos %></td>
	<td colspan=2 align=center >
	  <input type="submit" style=width:12em name="WhatAction" value="Team Combinations">
	</td>
    <td colspan=2># of Team Combos</td>
    <td colspan=2><font color="<%=TextColor2%>" size=2><%= sCombos %></font></td>


  </TR>

  <TR>
	<td colspan=2 align=center ><%= sTeamIndividual %></td>
	<td colspan=2 align=center >
	  <input type="submit" style=width:12em name="WhatAction" value="Individual Overall">
	</td>
    <td colspan=2>Individual Scores</td>
    <td colspan=2><font color="<%=TextColor2%>" size=2><%= sIndScores %></font></td>

  </TR>

  <TR>
	<td colspan=2 align=center ><%= sTeamSummary %></td>
	<td colspan=2 align=center >
	  <input type="submit" style=width:12em name="WhatAction" value="Team Summaries">
	</td>
    <td colspan=2>TBD</td>
    <td colspan=2><font color="<%=TextColor2%>" size=2>&nbsp;</font></td>

  </TR>

  <TR><td colspan=8 >&nbsp;</td></TR>

  <TR>
	<td colspan=8 align=center >
	  <input type="submit" style="width:20em; background-color:yellow;" name="WhatAction" value="Return to Menu" >
	</td>
  </TR>

</TABLE>

</form>

<br><%




END SUB



' ------------------------
  SUB ChangeOverlay_NEw
' ------------------------


OpenCon

'Set rs=Server.CreateObject("ADODB.recordset")
IF Whataction="delfromover" THEN
	sSQL = "UPDATE TS SET Marked='N' FROM usawsrank.IAC_TrialsSkiers AS TS WHERE MemberID='"&sMemberID&"' AND TourID='"&LEFT(sTourID,6)&"'"
ELSEIF Whataction="addtoover" THEN
	sSQL = "UPDATE TS SET Marked='Y' FROM usawsrank.IAC_TrialsSkiers AS TS WHERE MemberID='"&sMemberID&"' AND TourID='"&LEFT(sTourID,6)&"'"
END IF

'response.write("<br>")
'response.write(sSQL)


con.execute(sSQL)
CloseCon

END SUB




' --------------
  SUB ChangeOverlay
' --------------

'response.write("made it")
' response.end




IF Whataction="delfromover" THEN
		OpenCon
		sSQL = "UPDATE TS SET Marked='N' FROM usawsrank.IAC_TrialsSkiers AS TS WHERE MemberID='"&sMemberID&"'"
		con.execute(sSQL)
		CloseCon
ELSEIF Whataction="addtoover" THEN
		OpenCon
		sSQL = "UPDATE TS SET Marked='Y' FROM usawsrank.IAC_TrialsSkiers AS TS WHERE MemberID='"&sMemberID&"'"
		con.execute(sSQL)
		CloseCon
END IF




'response.write("<br>action="&action&"<br>")
'response.write(sSQL)
'response.end

Set rs=Server.CreateObject("ADODB.recordset")

sSQL = "SELECT Marked, FirstName, LastName, MemberID, Sex"
sSQL = sSQL + " FROM usawsrank.IAC_TrialsSkiers AS TS"
sSQL = sSQL + " JOIN"
sSQL = sSQL + " (SELECT PersonID, FirstName, LastName"
sSQL = sSQL + " 	FROM "&MemberShortTableName&") AS MT"
sSQL = sSQL + " ON MT.PersonID=RIGHT(TS.MemberID,8)"	 

sSQL = sSQL + " WHERE LEFT(TourID,6)='"&LEFT(sTourID,6)&"'"


rs.open sSQL, SConnectionToTRATable


rowCount = 0

' ---------------  Displays table HEADINGS  ----------------------

%>
&nbsp;<BR>&nbsp;<BR>
<TABLE Align=center BORDER="1" bgcolor="<%=TableColor1%>" WIDTH="<%=TourTableWidth%>" >
<TR>
<TD ALIGN="Center" bgcolor="<%=HQSiteColor2%>"><FONT COlOR="#FFFFFF" SIZE="<%=fontsize1%>">Delete</FONT></TD>
<TD ALIGN="Center" bgcolor="<%=HQSiteColor2%>"><FONT COlOR="#FFFFFF" SIZE="<%=fontsize1%>">Edit</FONT></TD>
<%

FOR i = 0 TO rs.fields.count - 1
	TempFN = rs.fields(i).name
	j = 0
	
	%>
   	<TD ALIGN="Center" vAlign="top" bgcolor="<%=HQSiteColor2%>" nowrap>
	  <FONT COlOR="#FFFFFF" FACE="<%=font1%>" SIZE="<%=fontsize1%>"><% Response.Write(Rs.Fields(i).name) %></FONT>
	</TD><%
NEXT

%>
</TR>
<%

'response.write("PW=")
'response.write(Session("validpassword"))
sPassword=Session("validpassword")

' --------------  Display table data here with paging --------------------------

DO WHILE NOT rs.eof
	
	Currcolor="#000000"
	IF rs.fields(3).value="Y" THEN Currcolor=sColor3

	%>
 	<TR>
	<TD ALIGN="center" vAlign="top"><FONT SIZE="<%=fontsize1%>"><% WriteLink "?Whataction=addtoover&sMemberID="&rs.fields(3).Value&"&process=editoverlay","Add","" %></FONT></TD>
	<TD ALIGN="center" vAlign="top"><FONT SIZE="<%=fontsize1%>"><% WriteLink "?Whataction=delfromover&sMemberID="&rs.fields(3).Value&"&process=editoverlay","Remove","" %></FONT></TD>
	<%

	FOR i = 0 TO rs.fields.count - 1
	
		TempFN = rs.fields(i).name
		
		%><TD ALIGN="center" vAlign="top" nowrap>
			<FONT COlOR="#000000" FACE="<%=font1%>" SIZE="<%=fontsize1%>">&nbsp;<%

		    IF isnull(rs.Fields(i).value) THEN
			response.write ("&nbsp;")
    		    ELSE
			Response.Write(trim(Rs.Fields(i).Value)) 
		    END IF  

		%>&nbsp;
		  </FONT>
		</TD>
	<%

	

	NEXT

		%>
		</TR>
		<% 
		rowCount = rowCount + 1
		rs.movenext
LOOP

%>
</TABLE>
<br><br><br><br>
<%

rs.close
set rs = nothing

END SUB






' ---------------------------------------
  SUB WriteLink(sParms,sDisplay,sBreak)
' ---------------------------------------

%>
<A HREF="<% Response.Write(ThisPage & sParms) %>"><% Response.Write(sDisplay) %></A><%

END SUB




' ------------
  SUB SaveIAC
' ------------

sUSBeginDate=Request("sUSBeginDate")
sUSEndDate=Request("sUSEndDate")
sIntBeginDate=Request("sIntBeginDate")
sIntEndDate=Request("sIntEndDate")
sIntFileName=Request("sIntFileName")



opencon
'response.write("<br>Top of Save")

' Logic for validating date numbers
'IF (isnumeric(left(sMembRegDate,2)) And isnumeric(right(left(sMembRegDate,5),2)) And isnumeric(right(sMembRegDate,4)) And right(left(sMembRegDate,3),1) = "/" And right(left(sMembRegDate,6),1) = "/" And isDate(sMembRegDate)) THEN



' --- Sets all defaults to 0 ---
sSQL ="UPDATE CT SET DefaultTour='0'"
sSQL = sSQL + " FROM "&IAC_ControlTableName&" AS CT" 
con.execute(sSQL)

' --- Now sets the selected one ---
'sSQL ="UPDATE CT SET DefaultTour='1'"
sSQL ="UPDATE CT SET USBeginDate='"&sUSBeginDate&"', USEndDate='"&sUSEndDate&"', DefaultTour='1',"
sSQL = sSQL + " IntBeginDate='"&sIntBeginDate&"', IntEndDate='"&sIntEndDate&"',"
sSQL = sSQL + " IntFileName='"&sIntFileName&"'"
sSQL = sSQL + " FROM "&IAC_ControlTableName&" AS CT" 
sSQL = sSQL + " WHERE TourID='"&sTourID&"'"

response.write("<br>"&sSQL)
'response.end
 	
con.execute(sSQL)

END SUB





' ----------------------------
  SUB BuildTrialsDropDownList
' ----------------------------

sSQL = "SELECT TourID, DefaultTour FROM "&IAC_ControlTableName&" as CT"
SET rsCT=Server.CreateObject("ADODB.recordset")
rsCT.open sSQL, SConnectionToTRATable

'response.write("<br>In drop sTourID="&sTourID)


IF NOT rsCT.eof THEN  %>

	<select name="sTourID" style="width:10em" onchange=submit()><%

	DO WHILE NOT rsCT.eof 
		IF sTourID="" THEN %>
			<option value = "<%=rsCT("TourID")%>" <% IF rsCT("DefaultTour")="1" THEN Response.Write(" selected ") %>><%=rsCT("TourID")%></Option><br><%
		ELSE %>
			<option value = "<%=rsCT("TourID")%>" <% IF TRIM(rsCT("TourID"))=TRIM(sTourID) THEN Response.Write(" selected ") %>><%=rsCT("TourID")%></Option><br><%
		END IF 
		rsCT.movenext
	LOOP %>

	</select><%

END IF

rsCT.close


END SUB



 




' ------------
  SUB IACGetPW
' ------------

'response.write("sPassword="&sPassword)
'response.end

	' ------------------------------------------------------------
	' ----------  Display initial request for Password  ----------
	' ------------------------------------------------------------
	%>

	<br><br><br><br><br>
	<TABLE class="innertable" BORDER="4" ALIGN="CENTER" width="325px" >
	  <TR>
	      <TH align=center><font face=<% =font1 %> size="4" Color="<%=TextColor5%>"><b>Enter Password</b></font><br></TH>
	  </TR>  


	  <TR>
	      <form action="/rankings/<%=ThisFileName%>?process=validpw" method="post">
	     <TD>
	     	<TABLE class="innertable" ALIGN="CENTER" width="90%" >

		  <tr>
	    	    <br>
		    <TH ALIGN="center" vAlign="top"><FONT COlOR="#FFFFFF" size=<% =fontsize2 %> face=<% =font1 %>><b>Password (Up to 10 digits)</b></FONT></th>
	  	  </tr>

	          <tr>	
        	    <TD ALIGN="center" vAlign="top" bgcolor="#FFFFFF"><input type="text" name="sPassword" maxlength=12 size=14></TD>
	          </tr>  

		  <tr>
	    	    <br>
		    <TH ALIGN="center" vAlign="top"><FONT COlOR="#FFFFFF" size=<% =fontsize2 %> face=<% =font1 %>><b>Email Address</b></FONT></th>
	  	  </tr>

	          <tr>	
        	    <TD ALIGN="center" vAlign="top" bgcolor="#FFFFFF"><input type="text" name="sEmail" maxlength=50 size=50>
		    </TD>
	          </tr>  

		</TABLE>
 
   		<TABLE BORDER="0" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#FFFFFF" width="90%" ><%

		    ' --- PW was entered and FOUND in PW table and NOT a match
		    IF sPassword <> "" THEN  %>
	          	<tr>	
        	    	  <TD colspan=2 ALIGN="center" style="border-style:none;"><FONT COlOR="<% =textcolor3 %>" size=<% =fontsize3 %> face=<% =font1 %>><% response.write("** Invalid Password **") %></FONT></TD>
			</tr><%
		    END IF  %>	


		  <tr>	<%
		    ' --- If this MemberID has a password, then display button to email password  --- 
			%>
			<td Align="center" style="border-style:none;">			
				<input type="submit" style="width:11em" value="Continue">
			</td>

    		  </tr>	
		</TABLE>
	    </TD>
		</form>

	  </TR>
	</TABLE><% 



END SUB





'---------------------
SUB WriteType(I)
'---------------------

SELECT CASE ucase(Rs.Fields(i).name)

	CASE "ID" %>
	   	<input type="hidden" name="id" value="<% Response.Write(sID) %>"> Auto Number<% 
		IF sid = 0 THEN 
     			response.write("(new)")
	   	ELSE
     			response.write(sID)
   	END IF

	CASE "SEX" %>
		<SELECT name="Sex">
		<option value="M" <%IF GetValue(i) = "M" THEN Response.Write("SELECTed")%>>Male</option>
		<option value="F" <%IF GetValue(i) = "F" THEN Response.Write("SELECTed")%>>Female</option>
		</SELECT>
		<%

	CASE "SKIYEARID" 
	    response.write(sSYID)
	    %><input type="hidden" name="SkiYearID" value="<%=sSYID%>"><%

	CASE "OLDSKIYEARID" 
		response.write("  <SELECT name=""SkiYearID"">   ")

		set rsSELECTFields=Server.CreateObject("ADODB.recordset")
    
    		sSQL = "SELECT * FROM " & SkiYearTableName
		rsSELECTFields.open sSQL, SConnectionToTRATable
  
    		DO WHILE not rsSELECTFields.eof
      			response.write("<option value =""" & rsSELECTFields("SkiYearID") & """")

			IF trim(rsSELECTFields("SkiYearID")) = trim(GetValue(i)) THEN
				response.write(" SELECTed")
			END IF

			IF GetValue(i) = "" and rsSELECTFields("DefaultYear") THEN
				response.write(" SELECTed")
			END IF

			response.write(">")
			response.write(rsSELECTFields("SkiYearName"))
			response.write("</option><br>")

			rsSELECTFields.movenext
		LOOP

		rsSELECTFields.close
		set rsSELECTFields = nothing
        
		response.write("  </SELECT>  ")

	CASE ELSE

		SELECT CASE Rs.Fields(i).type
			CASE 3 'primary key / auto number ?'
				%><input type="text" name="<% Response.Write(Rs.Fields(i).name) %>" SIZE="25" value="<% GetFieldValue i %>"><%
			CASE 20 'primary key / auto number ?'
				%><input type="text" name="<% Response.Write(Rs.Fields(i).name) %>" MAXLENGTH="<%=sLen(i)%>" SIZE="<%=sLen(i)%>" value="<% GetFieldValue i %>"><%
			CASE 11 'boolean'
        			%><INPUT TYPE="checkbox" NAME="<% Response.Write(Rs.Fields(i).name) %>" VALUE="0"<% GetcheckValue i %>><%
			CASE 203 'memo'
        			%><TEXTAREA NAME="<% Response.Write(Rs.Fields(i).name) %>" ROWS="20" COLS="56"><% GetFieldValue i %></TEXTAREA><%
			CASE ELSE 'not handled by this function'
			        %><input type="text" name="<% Response.Write(Rs.Fields(i).name) %>" MAXLENGTH="10" SIZE="10" value="<% GetFieldValue i %>"><%
		END SELECT

END SELECT 

END SUB

%>