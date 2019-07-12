<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/Bio-Form3.asp"-->
<!--#include virtual="/rankings/Tools_Definitions.asp"-->
<!--#include virtual="/rankings/qualifications.asp"-->
<!--#include virtual="/rankings/Tools_include16.asp"-->
<!--#include virtual="/rankings/tools_registration16.asp"-->
<%

Server.ScriptTimeout = 3000 

Dim ThisFileName
ThisFileName="view-registration2.asp"





' ---------------------------------------------------------------------------------------------
' --- This module displays various reports associated with REGISTRATION functions 
' --- Original module created by Mark Crone
' --- LAST updated: 7/4/2009
' ---------------------------------------------------------------------------------------------


Dim RegionSelected, EventSelected, DivSelected, StateSelected, sWhatReport, WhatPayments, WhatNotify, WhatLetter
Dim SequenceSelected, sBioFilter, sQualFilter, sWaivFilter, sFeeFilter, sResendEmail, sSentBioEmail
Dim emailbuttonstatus, FileLetter, currentline, ECount, MaxECount, ebody
Dim EVT1_TIME, EVT2_TIME, EVT3_TIME

' --- Used in OLR Listing ---
Dim sSortBy, sIncludePast, sIncludeFuture
Dim PreviousDiv, PreviousEvent, SkiYearID

Dim sDiv1, sDiv2, sDiv3, sDiv4, sEvent1, sEvent2, sEvent3, sEvent4
Dim sMemberID, LastMemb, sFirstName, sLastName, sFullName

Dim sLast4
Dim EmailCount, sState, sNoEmail, sEmail

Dim MembStatusTitle, MembStatusText, MembStatuscolor, FeesText, Feescolor, FeesTitle, regstatuscolor, regstatusText, regstatusTitle
Dim QualStatusEvent1, QualStatusEvent2, QualStatusEvent3, QualStatusEvent4, sRequirePart

Dim WaiverTitle, Waivercolor, WaiverText, TrickTitle, Trickcolor, TrickText
Dim BioText, Biocolor, BioTitle, BioLink, sRefreshData, SeedCount

Dim StartCharSelected, EndCharSelected
Dim PrintButton
Dim TextDropcolor1, TextDropcolor2

Dim sSQL, sSkiYearID

Dim sShowSQL
Dim sTestMode




' --- For debugging ---
sShowSQL = Request("sShowSQL")


'sTestMode="<br>**TEST MODE**"



'Session("sTourID")="07W999A"
'sTourID="07W999A"

'Session("sTourID")="08S093"
'sTourID="08S093"





' --- Resets to blank for testing purposes
'IF TRIM(Request("process"))="reset" THEN Session("sTourID")=""
IF TRIM(Request("process"))="viewreg" THEN Session("sTourID")=""


adminmenulevel = Session("adminmenulevel")
IF adminmenulevel = "" THEN adminmenulevel = 0

sMemberID = TRIM(Request("sMemberID"))
RegionSelected = trim(Request("RegionSelected"))
EventSelected = trim(Request("EventSelected"))
DivSelected = trim(Request("DivSelected"))
StateSelected = trim(Request("StateSelected"))
sWhatReport = LCASE(trim(Request("WhatReport")))





SequenceSelected = LCASE(trim(Request("SequenceSelected")))
WhatNotify=TRIM(Request("WhatNotify"))
WhatLetter=TRIM(Request("WhatLetter"))
sBioFilter=Request("sBioFilter")
sQualFilter=Request("sQualFilter")
sWaivFilter=Request("sWaivFilter")
sFeeFilter=Request("sFeeFilter")
sResendEmail=Request("sResendEmail")
sRefreshData=Request("sRefreshData")

sSkiYearID = Request("sSkiYearID")




StartCharSelected=Request("StartPulldown")
EndCharSelected=Request("EndPulldown")
IF StartCharSelected="" OR sWhatReport = "seeding" THEN StartCharSelected="All"
IF EndCharSelected="" OR sWhatReport = "seeding" THEN EndCharSelected="All"

sPrintDate=Request("sPrintDate")
PrintButton=Request("PrintButton")
ReturnButton=Request("ReturnButton")
IF ReturnButton="Main Menu" THEN
	Response.redirect("/rankings/defaultHQ.asp")
END IF


'----------------------------------------
' --- Sets Default values for report  ---
'----------------------------------------

'IF sWhatReport="" THEN sWhatReport="regstat"
IF sWhatReport = "" THEN sWhatReport = "noreportselected"
IF sWhatReport = "seeding" THEN StateSelected="" 

IF TRIM(Request("DivSelected")) = "" THEN DivSelected = "ALL"
IF TRIM(RegionSelected) = "" THEN RegionSelected = 6
IF TRIM(Request("StateSelected")) = "" THEN StateSelected = "All"
IF TRIM(SkiYearSelected) = "" THEN SkiYearSelected = 1

' ---- This will need a condition depending on which sports division  ----
IF TRIM(Request("EventSelected")) = "" THEN EventSelected = "ALL"
IF SequenceSelected = "" THEN SequenceSelected = "seed"



'response.write("sWhatReport="&sWhatReport)
'response.write("<br>sSkiYearID="&sSkiYearID)
'response.end




'response.write("sWhatNotify="&sWhatNotify)
'response.write("sWhatLetter="&sWhatLetter)
'response.write("sBioFilter="&sBioFilter)

'response.end




IF sWhatReport="tourstatus" THEN
	' --- Do nothing
ELSEIF TRIM(Request("sTourID"))<>"" THEN
	Session("sTourID")=TRIM(Request("sTourID"))
	sTourID=TRIM(Request("sTourID"))
ELSE
	IF TRIM(Session("sTourID"))="" THEN
		' - Go get tournament
		Session("sSendingPage") = "/rankings/"&ThisFileName&"?rid="&rid
		Session("sTourID") = ""

		sl=Request("sl")
		tr=Request("tr")
		ju=Request("ju")
		wb=Request("wb")
		ws=Request("ws")
		wu=Request("wu")
		bf=Request("bf")
		kb=Request("kb")
		hy=Request("hy")
		hf=Request("hf")
		jd=Request("jd")
		ad=Request("ad")

	   	IF sl="on" OR tr="on" OR ju="on" THEN		
			response.redirect("/rankings/view-tournamentsHQ.asp?sl=on&tr=on&ju=on&process=viewreg&rid="&rid)
		ELSEIF wb="on" OR ws="on" OR wu="on" THEN
			response.redirect("/rankings/view-tournamentsHQ.asp?wb=on&ws=on&wu=on&process=viewreg&rid="&rid)		
		ELSE
			response.redirect("http://www.usawaterski.org")	
		END IF		
	ELSE
		sTourID = Session("sTourID")
	END IF
END IF



' --- In SUB Tools_registration16.asp - Sets all tournament variables ---
IF sWhatReport<>"tourstatus" THEN
	DefineTourVariables_New

	' --- SUB found in tools_include.asp - Defines what events this sTSptsGrpID offers ---
	RegistrationEventsOffered (sTSptsGrpID)
END IF





'response.write("<br>sWhatReport = "&sWhatReport)

'sTourID="12S998B"
'sWhatReport="othersales"

SELECT CASE sWhatReport


  CASE "viewbio"
	
			Session("sSendingPage") = "/rankings/"&ThisFileName
	
			rdState = "/rankings/bio-form.asp?FormStatus=new&EditStatus=disabled&sMemberID="&sMemberID&"&sTourID="&sTourID
			response.redirect("/rankings/bio-form.asp?FormStatus=new&BioStatus=disabled&sMemberID="&sMemberID&"&sTourID="&sTourID)


  CASE "seeding"

			IF PrintButton<>"Printer Friendly" THEN WriteIndexPageHeader
			DropBoxFormat1
			BuildSeedingQuery
			EndofReportLine
			SeedingHeading
			RunSeedingStyleReport

			IF PrintButton<>"Printer Friendly" THEN WriteIndexPageFooter

  CASE "regstat"

			IF PrintButton<>"Printer Friendly" THEN  WriteIndexPageHeader
			DropBoxFormat1
			EndofReportLine
			BuildSeedingQuery
			RegStatHeading
			RunSeedingStyleReport
			DisplayBanquetCount
			IF PrintButton<>"Printer Friendly" THEN  WriteIndexPageFooter

  CASE "skierpayments"

			IF PrintButton<>"Printer Friendly" THEN  WriteIndexPageHeader_NoMenu	
			DropBoxFormat1
			BuildSeedingQuery
			SkiPayHeading
			RunSeedingStyleReport
			IF PrintButton<>"Printer Friendly" THEN  WriteIndexPageFooter

  CASE "scratched"

			IF PrintButton<>"Printer Friendly" THEN  WriteIndexPageHeader
			DropBoxFormat1
			EndofReportLine
			RegStatHeading
			BuildSeedingQuery
			RunSeedingStyleReport
			IF PrintButton<>"Printer Friendly" THEN  WriteIndexPageFooter

  CASE "bystate"

			IF PrintButton<>"Printer Friendly" THEN  WriteIndexPageHeader
			DropBoxFormat1
			ByStateHeading
			BuildSeedingQuery
			IF NOT rs.eof THEN
					RunByStateReport
			ELSE
					DisplayNoDataLine
			END IF

	CASE "othersales"	
			IF PrintButton<>"Printer Friendly" THEN  WriteIndexPageHeader
			DropBoxFormat1
			DisplayBanquet_andOther_Buyers
			DisplayBanquetCount	
			IF PrintButton<>"Printer Friendly" THEN  WriteIndexPageFooter



  CASE "notifications"
	
			BuildSeedingQuery
			MailNotices	

  CASE "financial"

			WriteIndexPageHeader_NoMenu
	PaymentReport	
	WriteIndexPageFooter

  CASE "divisiontotals"
			IF PrintButton<>"Printer Friendly" THEN  WriteIndexPageHeader
			SchedulingByDivision
			DisplayBanquetCount	
			IF PrintButton<>"Printer Friendly" THEN  WriteIndexPageFooter


  CASE "noreportselected"

'response.write("EXIT")
'response.end
	IF PrintButton<>"Printer Friendly" THEN  WriteIndexPageHeader
	DropBoxFormat1
	%><br><br><center><font size=<% =fontsize3 %> color="<% =Textcolor3 %>"><b>Select Report and Other Settings <br>then Press 'Display Report'</b></FONT></center><%
	IF PrintButton<>"Printer Friendly" THEN  WriteIndexPageFooter

   CASE "endoffile"

	IF PrintButton<>"Printer Friendly" THEN  WriteIndexPageHeader
	%><br><br><br><center><font size=<% =fontsize3 %> color="<% =Textcolor3 %>"><b>No Data Matching Search Settings</b></FONT></center><%
	IF PrintButton<>"Printer Friendly" THEN  WriteIndexPageFooter

END SELECT



' ------------------------------------------------------------------------------------------------------------------------------
' ---------------------   END OF MAIN SECTION OF PROGRAM    --------------------------------------------------------------------
' ------------------------------------------------------------------------------------------------------------------------------





' -----------------------------
  SUB RunSeedingStyleReport
' -----------------------------


'response.write("aml="&adminmenulevel)

IF NOT rs.eof THEN

	rs.movefirst
	PreviousEvent="XX"
	PreviousDiv="XX"


	DO WHILE Not rs.EOF 
		ScratchCount = 0
		sMemberID  = rs("MemberID")

		' --- Sets formatting based on expected data from that Event ---
		IF rs("Rank")<>"" THEN
				SELECT CASE TRIM(rs("event"))
					CASE "T", "WB", "WS"
							FRank=FormatNumber(rs("Rank"),0)
					CASE "S"
							FRank=FormatNumber(rs("Rank"),2)
					CASE "J"
							FRank=FormatNumber(rs("Rank"),2)
					CASE ELSE
							FRank=rs("Rank")
		  	END SELECT
		ELSE
				FRank=0.00
		END IF


		' ---------------------------------------------------------------
		' --- Test Membership Types and Expiration Date of Membership ---
		' ---------------------------------------------------------------

		sMemberOverride=rs("MembOverride")
		sCanSkiInTournaments=rs("CanSkiInTournaments")		
		sCanSkiInGRTournaments=rs("CanSkiInGRTournaments")	
		sEffectiveTo=rs("EffectiveTo")
		sMemberType=rs("MemberType")
		sMembTypeDesc=rs("MembTypeDesc")
		sFeeClassEvt=rs("FeeClass")


	   	MembershipStatus sMemberOverride, sCanSkiInTournaments, sCanSkiInGRTournaments, sFeeClassEvt, sTDateE, sEffectiveTo, sMemberType, sMembTypeDesc


		' --------------------------------------------
		' --- Total Fees are greater than Payments ---
		' --------------------------------------------

		sPayments=rs("Payments")
		sMoneyOverride=rs("MoneyOverride")
		sTotalEntry=rs("TotalEntry")
		sEntryType=rs("EntryType")

		IF 1=2 AND Session("AdminMenuLevel")>=50 THEN
				response.write("<br><br>rs(MemberID) = "&rs("MemberID"))
				response.write("<br>sPayments = "&sPayments)
				response.write("<br>sMoneyOverride = "&sMoneyOverride)
				response.write("<br>sTotalEntry = "&sTotalEntry)
				response.write("<br>sEntryType = "&sEntryType)

		END IF
		
		PaymentStatus2 sPayments, sMoneyOverride, sTotalEntry, sEntryType, rs("MemberID")

		' ---------------------------------------------------------------
		' --- Checks for Participation in Current SKI YEAR REGIONALS  ---
		' ---------------------------------------------------------------


	

' +++ LOOK AT THIS +++
		sdiv(1)=rs("div")


		' -------------------------------------------------------------------------------------------------------------------
		' --- Actually represents ANY required tournament participation, not just regionals ---
		' --- Condition added to allow overrides to take affect instantly, instead of only after recalc of qualificaitons ---
		' -------------------------------------------------------------------------------------------------------------------



		IF TRIM(rs("RegionalOverride"))<>"" THEN 
				sReglPartStat="X"
		ELSE
				sReglPartStat=TRIM(rs("SkiedRegls"))
				sRequirePart = TRIM(rs("RequirePart"))
		END IF	

		RequiredParticipation sReglPartStat

		'RegionalOverride


		' ------------------------------------------------------------------------------------
		' --- Rating, Masters as 3rd, Regional & National Placement Tested 	--------------
		' --- Uses Regional Participation results from ABOVE 			--------------
		' ------------------------------------------------------------------------------------
		sNatl_Plc=rs("Natl_Plc")
		sRegl_Plc=rs("Regl_Plc")
		sRating=rs("Rating")
		sRank=rs("Rank")

		EvtNo=1


		' ----------------------------------------------------------------------------
		' --- SUB ChkQualALL in Qualifications.asp module ---
		ChkQualALL2 sTourID, sMemberID, rs("Event"), rs("Div"), rs("QfyOverride") 				




		' --- 4-4-2012 New method of getting qualification data ---
		' ----------------------------------------------------------------------------
		IF TRIM(rs("LT_LeagueID"))<>"" THEN 
				QfyStatusTextNew = TRIM(rs("QfyStatusTextNew"))
				QfyStatusTitleNew = "Participant qualification status: " & rs("QfyStatusTextNew")

'response.write("<br>Line 455 - TRIM(rs(QfyOverride)) = "&TRIM(rs("QfyOverride")))
				IF TRIM(rs("QfyOverride")) <> "" THEN
						QfyStatusTextNew="(OV)"
						QfyStatusTitleNew = "&nbsp; OK - Administrative Override: "&rs("QfyOverride")
	
				ELSEIF IsNull(QfyStatusTextNew) THEN 
						ScratchCount = 1
						QfyStatusTextNew="None"
						QfyStatusTitleNew = "Participant has no qualification data"
				' --- Administrative Override ---
				ELSEIF QfyStatusTextNew="NCQ" THEN 
						ScratchCount = 1
						QfyStatusTextNew="NCQ"
						QfyStatusTitleNew = "Participant is not currently qualified"
				END IF

		ELSE 
				QfyStatusTextNew = "N/A"
				QfyStatusTitleNew = "Tournament has no qualification requirement"
		END IF



'		IF TRIM(rs("LT_LeagueID"))<>"" THEN 
'				QfyStatusTextNew = TRIM(rs("QfyStatusTextNew"))
'				QfyStatusTitleNew = "Participant qualification status: " & rs("QfyStatusTextNew")
'				IF IsNull(QfyStatusTextNew) THEN 
'						ScratchCount = 1
'						QfyStatusTextNew="None"
'						QfyStatusTitleNew = "Participant has no qualification data"
				' --- Administrative Override ---
'				ELSEIF TRIM(rs("QfyOverride")) <> "" THEN
'						QfyStatusTextNew="(OV)"
'						QfyStatusTitleNew = "&nbsp; OK - Administrative Override: "&rs("QfyOverride")
'				ELSEIF QfyStatusTextNew="NCQ" THEN 
'						ScratchCount = 1
'						QfyStatusTextNew="NCQ"
'						QfyStatusTitleNew = "Participant is not currently qualified"
'				END IF

'		ELSE 
'				QfyStatusTextNew = "N/A"
'				QfyStatusTitleNew = "Tournament has no qualification requirement"
'		END IF







'IF sMemberID="000153324" THEN
'		response.write("<br>QfyStatusTextNew = "&QfyStatusTextNew)
'END IF
		' --------------------
		' --- WAIVER Field ---  
		' --------------------

		sWaiverCode=rs("WaiverCode")

		VerifyWaiver (sWaiverCode)



		' ------------------------------------------------------------------------------------------
		' Trick Form Requirement - NEED TO ADD logic for looking at TouRGTTable to see if required.
		' ------------------------------------------------------------------------------------------

		sForm2Name=rs("Form2Name")

		VerifyTrickList sForm2Name, rs("Event"), EventSelected


		' -----------------
		' --- Bio Field ---
		' -----------------

		sForm1Name=rs("Form1Name")
		sBio_Reqd=rs("Bio_Reqd")
		sBioMemberID=rs("BioMemberID")
		ThisDiv = rs("div")
		ThisEvent	= TRIM(rs("event"))

		VerifyPersonalBio sForm1Name, sBioMemberID, sMemberID, sTourID 



		' ----------------------------------------------------------------------------------------
		' ------------------------------    BEGIN DISPLAY OF DATA   ------------------------------
		' ----------------------------------------------------------------------------------------

		IF sWhatReport="seeding" OR (sWhatReport="scratched" AND ScratchCount>0) OR sWhatReport="regstat" OR sWhatReport="override" THEN


		   IF SequenceSelected<>"regdateall" AND (rs("div")<>PreviousDiv OR rs("event")<>PreviousEvent) THEN  
				
					SELECT CASE ThisEvent
						CASE sTEvent(1)
								HeadEvent=sTEventName(1)
						CASE sTEvent(2)
								HeadEvent=sTEventName(2)
						CASE sTEvent(3)
								HeadEvent=sTEventName(3)
						CASE sTEvent(4)
								HeadEvent=sTEventName(4)
						CASE sTEvent(5)
								HeadEvent=sTEventName(5)
						CASE sTEvent(6)
								HeadEvent=sTEventName(6)
					END SELECT 
				
					' ---------------------------------------------
					' --- Displays the Division and Event break ---
					' ---------------------------------------------
					%>
			  	<tr>
						<TD>&nbsp;</TD>
						<TD align="Center" style="background-color:<%=scolor8%>;" valign="top"><font size=<% =fontsize3 %> color="<% =Textcolor1 %>"><%=rs("div")%></FONT></TD>
						<TD colspan=2 align="Center" style="background-color:<%=scolor8%>" >
							<font size=<% =fontsize3 %> color="<% =Textcolor1 %>"><%=HeadEvent%></font>
						</TD>
						<TD colspan=3 style="background-color:<%=scolor8%>;">&nbsp; </TD>
						<%
						
						' --------------------------------------------------------------------
						' --- Displays the Print Bio Button to run bios for this event/div ---
						' --------------------------------------------------------------------
						testing="N"
						IF testing="Y" OR Session("adminmenulevel")>=2 THEN
								%>
								<TD colspan=3 align="Center" style="background-color:<%=scolor8%>" >
									<a href="http://www.usawaterski.org/rankings/registration-email.asp?WhatReport=notifications&WhatLetter=Bio&WhatNotify=PrintBio&sTourID=<%=sTourID%>&EventSelected=<%=ThisEvent%>&DivSelected=<%=ThisDiv%>" target="_blank">
										<input type="submit" value="<%=ThisDiv%>-<%=ThisEvent%> Bios" style="width:6em; height:1.5em" title="Print Bios For <%=ThisDiv%> - <%=HeadEvent%>">
									</a>
								</TD>
			  				<%
			  		ELSE
			  				%><TD colspan=3 style="background-color:<%=scolor8%>;">&nbsp;</TD><%
			  		END IF
			  		%>
			  		<TD colspan=1 style="background-color:<%=scolor8%>;">&nbsp;</TD>		
			   		<%
		   			IF Session("adminmenulevel")>=2 THEN
			  				%><TD colspan=2 style="background-color:<%=scolor8%>;">&nbsp;</TD><%
						END IF
						%>
			  		</tr> 
						<%
				
				SeedCount=0	

		   	END IF


		   SeedCount=SeedCount+1
		   IF sWhatReport="seeding" THEN %>

			<TR>
			  <TD align="Center" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=SeedCount%></font></TD>
			  <TD align="Left" >
			    <font size=<%=fontsize2%> face=<% =font2 %>>&nbsp;<% =rs("LastName")&", "&rs("FirstName") %></font>
			  </TD>
				<TD align="center"><font size=<%=fontsize2%>><%=rs("MemberID")%></font></TD>
				<TD align="Center" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=rs("state")%></FONT></TD>
				<TD align="Center" style="background-color:<%=Tablecolor1%>;">
					<font size=<%=fontsize2%> color="<% =Textcolor1 %>">
						<a title="G - Grassroots&#13;S - Premier (Class C or Alternate to Record)&#13;R - Record">
						<%
				
						IF rs("FeeClass")="G" THEN 
								response.write(LEFT(sTGRClassText,5)) 
						ELSEIF rs("FeeClass")="S" THEN 
								response.write(LEFT(sTBaseClassText,5)) 
						ELSEIF rs("FeeClass")="R" THEN 
								response.write(LEFT(sTUpgradeClassText,5)) 
						END IF
				
						%>
						</a>
					</font>
				</TD>
			  <TD align="Center"><font size=<%=fontsize2%>>
					<a TITLE="Skill Level is used for ability-based grouping of competitors"><%=SkillDecode(rs("Skill"))%></a></font>
				</td>
		    <TD align="Center" style="background-color:<%=Tablecolor1%>;"><font size=<%=fontsize2%> color="<% =Textcolor1 %>"><% =FRank %></FONT></TD> 
		    <TD align="Center" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =RIGHT(rs("Rating"),1) %></TD>
				<TD align="Center" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>"><% =formatnumber(rs("RankPct"),2) %></FONT></TD>
		    <TD align="Center" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =rs("Natl_plc") %></FONT></TD>
		    <TD align="Center" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =rs("Regl_plc") %></FONT></TD>
			</TR>
			<%

 
		   ELSEIF sWhatReport="regstat" OR sWhatReport="scratched" THEN  


'IF rs("MemberID")="000153324" THEN
'		response.write("<br>1 RegStatusText="&RegStatusText)
		'sReglPartStat="X"
'		response.write("<br>sReglPartStat="&sReglPartStat)
'		response.write("<br>")
'		response.write(TRIM(RegStatusText)="DNS" AND TRIM(sReglPartStat)<>"X")
'		response.write("<br>QfyStatusTextNew="&QfyStatusTextNew)
'		response.write("<br>Feescolor="&Textcolor)		
'END IF
					' --- Show Record only if Fees are Paid (not red) OR (TestCalidAdminCode=true OR AdminMenuLevel>=20)  
		      IF sWhatReport="scratched" OR (Feescolor<>Textcolor3 OR LEFT(FeesText,1)="(") OR TestValidAdminCode OR adminmenulevel>=10 OR TRIM(QfyStatusTextNew)="NCQ" OR (TRIM(RegStatusText)="DNS" AND sReglPartStat<>"X") THEN 
		      		%>
							<TR>
			  				<TD align="Center" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=SeedCount%></FONT></TD>
			  				<TD align="Left" >
			  				<%
								IF TRIM(rs("MembEmail"))<>"" AND (TestValidAdminCode OR adminmenulevel>=10) THEN 
										%>
										<font size=<%=fontsize2%> face=<% =font2 %>><a title="Send Email to <% =rs("FirstName")%> <%=rs("LastName")%>" href="mailto:<%=rs("MembEmail")%>?subject=Registration issue for <%=sTourName%> - Member <% =rs("LastName")&", "&rs("FirstName") %>"><%=sMembEmail%>&nbsp;<% =rs("LastName")&", "&rs("FirstName") %></a>
										</font>
										<%
								ELSE 
										%>
										<font size=<%=fontsize2%> face=<% =font2 %>>&nbsp;<% =rs("LastName")%>, <%=rs("FirstName")%><font>
										<%
								END IF 
								%>
			  				</TD>
		          	<TD>
								<% 
								
								IF (adminmenulevel >= 30 OR TestValidAdminCode=true) THEN 
										%>
										<font size=<%=fontsize2%>>
											<a href = "/rankings/registration16.asp?sMemberID=<%=rs("MemberID")%>&sTourID=<%=sTourID%>" title="Link to open Registration record for <% =rs("FirstName")&" "&rs("LastName") %>" target="_blank"><%=rs("MemberID")%></a>
										</font>
										<% 
								ELSE 
										%>
										<font size=<%=fontsize2%>><%=rs("MemberID")%></font>
										<% 
								END IF 

								%>
							  </TD>	
		    				<TD align="Center" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=rs("state")%></FONT></TD>
								<%
								
								' -----------------------------------------
								' --- Registration status data
								' -----------------------------------------

								%>
			  				<TD align="Center" >
			  					<font size=<%=fontsize2%> color="<% =MembStatuscolor %>">
										<a TITLE="<% =MembStatusTitle %>"><% =MembStatusText %></a>
									</FONT>
								</TD>
								<%

								' ----------------------
								' --- Fee Status     ---
								' ----------------------
								%>
			  				<TD align="Center" >
			  					<font size=<%=fontsize2%> color="<% =Feescolor %>">
			  						<a TITLE="<% =FeesTitle %>"><% =FeesText %></a>
			  					</FONT>
			  				</TD>
			  				<%

'IF rs("MemberID")="000153324" THEN
'			response.write("<br>QfyStatusTextNew="&QfyStatusTextNew)
'END IF				
								' ----------------------
								' --- Qualifications ---
								' ----------------------
								IF TRIM(QfyStatusTextNew)="Qualified" OR TRIM(QfyStatusTextNew)="QFY-RPR" OR TRIM(QfyStatusTextNew)="Pending" OR TRIM(QfyStatusTextNew)="NCQ" THEN 
										%>
			  						<TD align="Center" >
				 							<font size="<%=fontsize2%>" color="<%=Textcolor2%>">
												<a TITLE="<%= QfyStatusTitleNew %>" href="/rankings/MemberQualifications.asp?sMemberID=<%=rs("MemberID")%>&sTourID=<%=sTourID%>" 	target="_blank"><%= QfyStatusTextNew %></a>
											</FONT>
										</TD>
										<%
								ELSE 
										%>
			  						<TD align="Center" >
			 								<font size="<%=fontsize2%>" color="<%=Textcolor2%>">
												<a TITLE="<%=QfyStatusTitleNew%>"><%=QfyStatusTextNew%></a>
											</font>
										</TD>
										<%
								END IF 
				
								' ----------------------
								' --- Waiver         ---
								' ----------------------
								%> 
								<TD align="Center" >
									<font size=<%=fontsize2%> color="<% =Waivercolor %>">
										<a TITLE="<% =WaiverTitle %>"><% = WaiverText %></a>
									</FONT>
								</TD>
								<% 

								' ----------------------
								' --- Trick Forms    ---
								' ----------------------
				
								IF TRIM(lCase(rs("Form2Name")))="list" AND (TRIM(uCase(rs("Event"))) = "T" OR (TRIM(uCase(rs("Event"))) <> "T" AND EventSelected = "ALL")) THEN 
										%>
					  				<TD align="Center" >
					  					<font size=<%=fontsize2%> color="<% =Trickcolor %>">
					 							<a TITLE="<% =TrickTitle %>">&nbsp;<% =TrickText %></a>
					 						</FONT>
					 					</TD>
					 					<% 
								END IF 


								' ----------------------
								' --- Bio Form       ---
								' ----------------------
								IF (adminmenulevel >= 30 OR TestValidAdminCode=true) AND BioLink <> "" THEN 
										%>
					  				<TD align="Center" >
					  					<font size="<%=fontsize2%>" color = "<% =Biocolor %>" face="<% =font1 %>">
												<a title="Link to bio for <%=rs("FirstName")%> <%=rs("LastName")%>" href = "/rankings/bio-form.asp?FormStatus=new&BioStatus=disabled&sMemberID=<%=rs("MemberID")%>&sTourID=<%=sTourID%>" target="_blank"><% =BioText %>
												</a>
											</font>
										</td>
										<% 
								ELSE 
										%>
					  				<TD align="Center" >
					  					<font size="<%=fontsize2%>" color = "<% =Biocolor %>" face="<% =font1 %>">
												<a TITLE="<% =BioTitle %>">&nbsp; <% =BioText %></a>
											</font>
										</td>
										<% 
								END IF 
	

								' ----------------------
								' --- Register Date  ---
								' ----------------------

								%>
  							<TD align="Center" >
  								<font size=<%=fontsize2%> color="<% =regstatuscolor %>">
			 							<a TITLE="<% =regstatusTitle %>"><% =regstatusText %></a>
			 						</FONT>
			 					</TD>
        				<TD align="Center" >
        					<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=rs("RegDate")%></FONT>
        				</TD>
        				<%

			  				IF (adminmenulevel>=20 OR TestValidAdminCode=true) THEN

										IF rs("BanquetQty")>0 THEN 
												%>
												<TD align="Center" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=rs("BanquetQty")%></FONT></TD>
												<%
										ELSE 
												%>
												<TD align="Center" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">--</FONT></TD><%				
										END IF  


										IF TRIM(rs("WaiverCode"))<>"Paper" AND TRIM(rs("WaiverCode"))<>"" THEN 
												%>
			      				  	<TD align="Center"><font size=<%=fontsize2%>>Elec</FONT></TD>
			      				  	<%
										ELSEIF TRIM(rs("WaiverCode"))="Paper" THEN 
												%>
				        				<TD align="Center"><font size=<%=fontsize2%>>Paper</FONT></TD>
				        				<%
										ELSEIF TRIM(rs("WaiverCode"))="" THEN 
												%>
				        				<TD align="Center"><font size=<%=fontsize2%> color="red">None</FONT></TD>
				        				<%
										END IF  
				
			  				END IF

			  				%>
							</TR><%
		    	ELSE 
		    			%>
							<TR>
			  				<TD align="Center" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=SeedCount%></FONT></TD>
			  				<TD align="Left" >
			  				<%
								IF TRIM(rs("MembEmail"))<>"" AND (TestValidAdminCode OR adminmenulevel>=10) THEN 
										%>
										<font size=<%=fontsize2%> face=<% =font2 %>>
											<a title="Send Email to <% =rs("FirstName")%> <%=rs("LastName")%>" href="mailto:<%=rs("MembEmail")%>?subject=Registration issue for <%=sTourName%> - Member <% =rs("LastName")&", "&rs("FirstName") %>"><%=sMembEmail%>&nbsp;<% =rs("LastName")&", "&rs("FirstName") %></a>
										</font>
										<%
								ELSE 
										%>
										<font size=<%=fontsize2%> face=<% =font2 %>>&nbsp;<% =rs("LastName")%>, <%=rs("FirstName")%><font>
										<%
								END IF 
								%>
							  </TD>
		    	      <TD>
								<% 
								IF (adminmenulevel >= 30 OR TestValidAdminCode=true) THEN 
										%>
										<font size=<%=fontsize2%>>
											<a href = "/rankings/registration16.asp?sMemberID=<%=rs("MemberID")%>&sTourID=<%=sTourID%>" title="Link to open Registration record for <% =rs("FirstName")&" "&rs("LastName") %>" target="_blank"><%=rs("MemberID")%></a>
										</font>
										<% 
								ELSE 
										%>
										<font size=<%=fontsize2%>><%=rs("MemberID")%></font>
										<% 
								END IF 
								
								IF sMemberID="000001151" THEN
										FeesMessage="ADMINISTRATION ONLY - TESTING"
								ELSE
										FeesMessage="Not Registered - Fees Not Paid"
								END IF			
								
								%>
			  				</TD>	
		          	<TD align="Center"><font size="<%=fontsize2%>" color="<%=Textcolor1%>"><%=rs("state")%></FONT></TD>
		          	<TD align="Center" >
		          		<font size=<%=fontsize2%> color="<%=MembStatuscolor%>">
										<a TITLE="<% =MembStatusTitle %>"><% =MembStatusText %></a>
									</FONT>
								</TD>
			  				<TD align="center" >
			  					<font size=<%=fontsize2%> color="<% =Feescolor %>">
			  						<a TITLE="<% =FeesTitle %>"><% =FeesText %></a>
			  					</FONT>
			  				</TD>
			 					<TD align="center">
			 						<font size="<%=fontsize2%>" color="<%=Textcolor2%>">
			 							<a TITLE="<%=QfyStatusTitleNew%>"><%=QfyStatusTextNew%></a>
			 						</font>
			 					</td>			  
								<TD align="center" >
									<font size=<%=fontsize2%> color="<% =Waivercolor %>">
										<a TITLE="<% =WaiverTitle %>"><% = WaiverText %></a>
									</FONT>
								</TD>
			  				<TD colspan=5>
			 						<font size=<%=fontsize2%> color="<%=Feescolor%>"><%=FeesMessage%></font>
			  				</TD>
								</TR>
								<% 	

	
		     		END IF  ' --- Bottom of exlcusion for not paid ---

		
		   ELSEIF sWhatReport="override" THEN %>

			<TR>
			   <TD align="Center" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=SeedCount%></FONT></TD>
			   <TD align="Center" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">
				<a href="/rankings/view-regist.asp?member_id=<%=sMemberID%>&event=<%=rs("Event")%>&pvar=ByMember">Override</a>
				</FONT>
			   </TD>

			  <TD align="Left" >
			    <font size=<%=fontsize2%> color="<% =Textcolor1 %>">
			    <font size=<%=fontsize2%> face=<% =font2 %>>&nbsp;<% =rs("LastName")&", "&rs("FirstName") %></FONT>
			  </TD>
		          <TD><font size=<%=fontsize2%>><%=rs("MemberID")%></font></TD>
		          <TD align="Center" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=rs("state")%></FONT></TD>

		          <TD align="Center" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;-- </FONT></TD>
		          <TD align="Center" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;-- </FONT></TD>
		          <TD align="Center" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =rs("QfyOverride") %></FONT></TD>
		          <TD align="Center" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;-- </FONT></TD>
		          <TD align="Center" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =rs("WaiverCode") %></FONT></TD>
			</TR><%
		   END IF  


		END IF 	' --- Bottom of a variety of report types 




		IF sWhatReport="skierpayments" THEN  

			IF rs("div")<>PreviousDiv OR rs("event")<>PreviousEvent THEN  %>
			   <tr>
				<TD align="Center" ><font size=<% =fontsize3 %> color="<% =Textcolor1 %>"><%=rs("div")%></FONT></TD><%
				SELECT CASE TRIM(rs("event"))
					CASE sTEvent(1)
						HeadEvent=sTEventName(1)
					CASE sTEvent(2)
						HeadEvent=sTEventName(2)
					CASE sTEvent(3)
						HeadEvent=sTEventName(3)
					CASE sTEvent(4)
						HeadEvent=sTEventName(4)
					CASE sTEvent(5)
						HeadEvent=sTEventName(5)
					CASE sTEvent(6)
						HeadEvent=sTEventName(6)
				END SELECT 
				
				
				NumFeeColumns=TotNumOptItems+4
				%>
				<TD colspan=2 align="Center" ><font size=<% =fontsize3 %> color="<% =Textcolor1 %>"><%=HeadEvent%></FONT></TD>
				<TD colspan="<%=NumFeeColumns%>" style="background-color:<%=Tablecolor1%>;">&nbsp; </TD>
				<TD colspan=4 style="background-color:<%=Tablecolor1%>;">&nbsp; </TD>
				<TD colspan=2 style="background-color:<%=Tablecolor1%>;">&nbsp; </TD>

			   </tr> <%	
			END IF


			SET rsRegTrans=Server.CreateObject("ADODB.recordset")
			sSQL = "(SELECT MAX(TransDate) AS maxdate FROM "&RegTransTableName
			sSQL = sSQL + " WHERE LEFT(TourID,6) = '"&LEFT(sTourID,6)&"' AND MemberID = '"&rs("MemberID")&"') AS d"
			rsRegTrans.open sSQL, SConnectionToTRATable, 3, 3

			mDate = rsRegTrans("maxdate")     ' --- Latest date
			rsRegTrans.close

			' ----  Reads all transactions with matching date/time  ----
			SET rsRegTrans=Server.CreateObject("ADODB.recordset")
			sSQL = "SELECT MemberID, TourID, TransCode, Amount FROM "&RegTransTableName
			sSQL = sSQL + " WHERE TransDate = '"&mDate&"' AND LEFT(TourID,6) = '"&LEFT(sTourID,6)&"' AND MemberID = '"&rs("MemberID")&"'"
			rsRegTrans.open sSQL, SConnectionToTRATable, 3, 3

			%>
			<TR>
			  <TD align="Left" valign="top" style="background-color:<%=Tablecolor1%>;"><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =TRIM(rs("lastname"))&", "&rs("firstname")%></FONT></TD>
			  <TD align="Left" valign="top" style="background-color:<%=Tablecolor1%>;"><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =rs("MemberID") %></FONT></TD>
        		  <TD align="Center" valign="top" style="background-color:<%=Tablecolor1%>;"><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =rs("EntryType")%></FONT></TD>
			<%
	
			sEntryFees = FormatNumber(0,2) 
			sLateFee = FormatNumber(0,2)
			sAWSEFDonation = FormatNumber(0,2)
			sBanquetTot = FormatNumber(0,2)

			sOF1Fee = FormatNumber(0,2)
			sOF2Fee = FormatNumber(0,2)
			sOF3Fee = FormatNumber(0,2)
			sOF4Fee = FormatNumber(0,2)
			sOF5Fee = FormatNumber(0,2)
			sOF6Fee = FormatNumber(0,2)
			sOF7Fee = FormatNumber(0,2)
			sOF8Fee = FormatNumber(0,2)
			sOF9Fee = FormatNumber(0,2)
			sOF10Fee = FormatNumber(0,2)

			sOffDiscAmt = FormatNumber(0,2)
			sJrDiscAmt = FormatNumber(0,2)
			sSrDiscAmt = FormatNumber(0,2)
			sClubDiscAmt = FormatNumber(0,2)

		    IF NOT rsRegTrans.eof THEN

			rsRegTrans.movefirst	

			DO WHILE NOT rsRegTrans.eof 
				SELECT CASE TRIM(rsRegTrans("TransCode"))
					CASE "FEF"
							sEntryFees = FormatNumber(rsRegTrans("Amount"),2)
					CASE "FLF"
							sLateFee = FormatNumber(rsRegTrans("Amount"))
					CASE "OBF"
							sAWSEFDonation = FormatNumber(rsRegTrans("Amount"))
					CASE "BAN"
							sBanquetTot = FormatNumber(rsRegTrans("Amount"))
					CASE "FLF"
							sLateFee = FormatNumber(rsRegTrans("Amount"))
					CASE "OF1"
							sOF1Fee = FormatNumber(rsRegTrans("Amount"))
					CASE "OF2"
							sOF2Fee = FormatNumber(rsRegTrans("Amount"))
					CASE "OF3"
							sOF3Fee = FormatNumber(rsRegTrans("Amount"))
					CASE "OF4"
							sOF4Fee = FormatNumber(rsRegTrans("Amount"))
					CASE "OF5"
							sOF5Fee = FormatNumber(rsRegTrans("Amount"))
					CASE "OF6"
							sOF6Fee = FormatNumber(rsRegTrans("Amount"))
					CASE "OF7"
							sOF7Fee = FormatNumber(rsRegTrans("Amount"))
					CASE "OF8"
							sOF8Fee = FormatNumber(rsRegTrans("Amount"))
					CASE "OF9"
							sOF9Fee = FormatNumber(rsRegTrans("Amount"))
					CASE "OF10"
							sOF10Fee = FormatNumber(rsRegTrans("Amount"))


					CASE "DOF"
							sOffDiscAmt = FormatNumber(rsRegTrans("Amount"))
					CASE "DJR"
							sJrDiscAmt = FormatNumber(rsRegTrans("Amount"))
					CASE "DSR"
							sSrDiscAmt = FormatNumber(rsRegTrans("Amount"))
					CASE "DCL"
							sClubDiscAmt = FormatNumber(rsRegTrans("Amount"))
				END SELECT  
			  rsRegTrans.movenext
			LOOP


			sTotalFees = FormatNumber(cdbl(sEntryFees) + cdbl(sAWSEFDonation) + cdbl(sLateFee) + cdbl(sBanquetTot) + cdbl(sOffDiscAmt) + cdbl(sJrDiscAmt) + cdbl(sSrDiscAmt) + cdbl(sClubDiscAmt) + cdbl(sOF1Fee) + cdbl(sOF2Fee) + cdbl(sOF3Fee) + cdbl(sOF4Fee) + cdbl(sOF5Fee) + cdbl(sOF6Fee) + cdbl(sOF7Fee) + cdbl(sOF8Fee) + cdbl(sOF9Fee) + cdbl(sOF10Fee) ,2)
			sBalDue = FormatNumber(cdbl(sTotalFees) - cdbl(rs("Payments")),2)
			

			%>
			  <TD align="Center" style="background-color:<%=Tablecolor1%>;" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =sEntryFees %></FONT></TD>
			  <TD align="Center" style="background-color:<%=Tablecolor1%>;" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =sLateFee %></FONT></TD>
			  <TD align="Center" style="background-color:<%=Tablecolor1%>;" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =sAWSEFDonation %></FONT></TD>
			  <TD align="Center" style="background-color:<%=Tablecolor1%>;" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =sBanquetTot %></FONT></TD>
				<%
				
				IF TRIM(sOF1Desc)<>"" THEN 	
			  		%><TD align="Center" style="background-color:<%=Tablecolor1%>;" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =sOF1Fee %></FONT></TD><%
			  END IF		
				IF TRIM(sOF2Desc)<>"" THEN 				  
			  		%><TD align="Center" style="background-color:<%=Tablecolor1%>;" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =sOF2Fee %></FONT></TD><%
			  END IF		
				IF TRIM(sOF3Desc)<>"" THEN 				  
			  		%><TD align="Center" style="background-color:<%=Tablecolor1%>;" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =sOF3Fee %></FONT></TD><%
			  END IF		
				IF TRIM(sOF4Desc)<>"" THEN 				  
			  		%><TD align="Center" style="background-color:<%=Tablecolor1%>;" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =sOF4Fee %></FONT></TD><%
			  END IF		
				IF TRIM(sOF5Desc)<>"" THEN 				  
			  		%><TD align="Center" style="background-color:<%=Tablecolor1%>;" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =sOF5Fee %></FONT></TD><%
			  END IF		
				IF TRIM(sOF6Desc)<>"" THEN 				  
			  		%><TD align="Center" style="background-color:<%=Tablecolor1%>;" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =sOF6Fee %></FONT></TD><%
			  END IF		
				IF TRIM(sOF7Desc)<>"" THEN 				  
			  		%><TD align="Center" style="background-color:<%=Tablecolor1%>;" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =sOF7Fee %></FONT></TD><%
			  END IF		
				IF TRIM(sOF8Desc)<>"" THEN 				  
			  		%><TD align="Center" style="background-color:<%=Tablecolor1%>;" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =sOF8Fee %></FONT></TD><%
			  END IF		
				IF TRIM(sOF9Desc)<>"" THEN 				  
			  		%><TD align="Center" style="background-color:<%=Tablecolor1%>;" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =sOF9Fee %></FONT></TD><%
			  END IF		
				IF TRIM(sOF10Desc)<>"" THEN 				  
			  		%><TD align="Center" style="background-color:<%=Tablecolor1%>;" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =sOF10Fee %></FONT></TD><%
			  END IF		

				%>
			  <TD align="Center" style="background-color:<%=Tablecolor1%>;" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =sOffDiscAmt %></FONT></TD>
			  <TD align="Center" style="background-color:<%=Tablecolor1%>;" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =sJrDiscAmt %></FONT></TD>
			  <TD align="Center" style="background-color:<%=Tablecolor1%>;" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =sSrDiscAmt %></FONT></TD>
			  <TD align="Center" style="background-color:<%=Tablecolor1%>;" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =sClubDiscAmt %></FONT></TD>
			  <TD align="Center" style="background-color:<%=Tablecolor1%>;" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =sTotalFees %></FONT></TD>
			  <TD align="Center" style="background-color:<%=Tablecolor1%>;" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =sBalDue %></FONT></TD>
			</TR><%

			rsRegTrans.close
		   ELSE  %>
			  <TD colspan=16 align="left" style="background-color:<%=Tablecolor1%>;" ><font size=<%=fontsize2%> color="<% =Textcolor3 %>">&nbsp;No Payment Reported to Registration System - Check PayPal Receipts</FONT></TD><%
		   END IF  


		END IF		' --- Bottom of general IF for a variety of report types 


		IF sWhatReport="scratched" AND ScratchCount>0 THEN
			PreviousEvent=rs("event")
			PreviousDiv=rs("div")
		ELSEIF sWhatReport<>"scratched" THEN
			PreviousEvent=rs("event")
			PreviousDiv=rs("div")
		END IF

'		IF sWhatReport="scratched" AND (ScratchCount>1 OR (ScratchCount=1 AND regstatusText="PEND")) THEN
'			PreviousEvent=rs("event")
'			PreviousDiv=rs("div")
'		ELSEIF	sWhatReport<>"scratched" THEN
'			PreviousEvent=rs("event")
'			PreviousDiv=rs("div")
'		END IF


		rs.MoveNext	



	LOOP %>

	</TABLE><%

	rs.close
	Set rs = nothing
	CloseCon

	
ELSE  

	DisplayNoDataLine

END IF





END SUB


' ----------------------
  SUB DisplayNoDataLine
' ----------------------
	%>
	<table align=center width=100%>
	  <td align=center><font size=<% =fontsize3 %> color=<% =textcolor3 %>><br><b>No Data For These Settings.</b></font></td>
	</table><%


END SUB


' ----------------------
  SUB EndofReportLine
' ----------------------
	%>
	<center><font size=<% =fontsize3 %> color=<% =textcolor3 %>><b>Hold your cursor over certain fields or headings to view details about that item.</b></font></center><%


END SUB


' -----------------------
  SUB DisplayBanquetCount
' -----------------------


'sSQL = " SELECT OF1Desc, OF2Desc, OF3Desc, OF4Desc, OF5Desc, OF6Desc, OF7Desc, OF8Desc, OF9Desc, OF10Desc"
'sSQL = sSQL + ", BTickCost, BTickWithE"

Set rs=Server.CreateObject("ADODB.recordset")
sSQL = " SELECT SumBanquetFee, BanquetCount"
sSQL = sSQL + ", OF1QtyCount, OF2QtyCount, OF3QtyCount, OF4QtyCount, OF5QtyCount, OF6QtyCount, OF7QtyCount, OF8QtyCount, OF9QtyCount, OF10QtyCount" 			
sSQL = sSQL + ", OF1FeeSum, OF2FeeSum, OF3FeeSum, OF4FeeSum, OF5FeeSum, OF6FeeSum, OF7FeeSum, OF8FeeSum, OF9FeeSum, OF10FeeSum" 			
sSQL = sSQL + "  			FROM "&TRegSetupTableName&" SRT"

sSQL = sSQL + " LEFT JOIN "
sSQL = sSQL + " (SELECT TourID, SUM(BanquetQty) AS BanquetCount, SUM(BanquetFee) AS SumBanquetFee"
sSQL = sSQL + ", SUM(OF1Qty) AS OF1QtyCount, SUM(OF2Qty) AS OF2QtyCount, SUM(OF3Qty) AS OF3QtyCount" 
sSQL = sSQL + ", SUM(OF4Qty) AS OF4QtyCount, SUM(OF5Qty) AS OF5QtyCount, SUM(OF6Qty) AS OF6QtyCount"
sSQL = sSQL + ", SUM(OF7Qty) AS OF7QtyCount, SUM(OF8Qty) AS OF8QtyCount, SUM(OF9Qty) AS OF9QtyCount"
sSQL = sSQL + ", SUM(OF10Qty) AS OF10QtyCount"
sSQL = sSQL + ", SUM(OF1Fee) AS OF1FeeSum, SUM(OF2Fee) AS OF2FeeSum, SUM(OF3Fee) AS OF3FeeSum" 
sSQL = sSQL + ", SUM(OF4Fee) AS OF4FeeSum, SUM(OF5Fee) AS OF5FeeSum, SUM(OF6Fee) AS OF6FeeSum"
sSQL = sSQL + ", SUM(OF7Fee) AS OF7FeeSum, SUM(OF8Fee) AS OF8FeeSum, SUM(OF9Fee) AS OF9FeeSum"
sSQL = sSQL + ", SUM(OF10Fee) AS OF10FeeSum"
sSQL = sSQL + " FROM "&RegGenTableName
sSQL = sSQL + " GROUP BY TourID) RG1"
sSQL = sSQL + " ON LEFT(RG1.TourID,6)=LEFT(SRT.TournAppID,6)"

sSQL = sSQL + " WHERE LEFT(SRT.TournAppID,6)='"&LEFT(sTourID,6)&"'"

'response.write(sSQL)
'response.end

rs.open sSQL, sConnectionToTRATable, 3, 1

	
IF (adminmenulevel>=20 OR TestValidAdminCode=true) THEN  
		'BTickCost=rs("BTickCost")
		'BTickWithE=rs("BTickWithE")
		ShowBanquetColumn=false
		IF sBTickCost>0 OR sBTickWithE=true THEN ShowBanquetColumn=true
		TotalTableWidth=400
		IF ShowBanquetColumn=true THEN TotalTableWidth=TotalTableWidth+50
		IF TRIM(sOF1Desc)<>"" THEN TotalTableWidth=TotalTableWidth+50
		IF TRIM(sOF2Desc)<>"" THEN TotalTableWidth=TotalTableWidth+50
		IF TRIM(sOF3Desc)<>"" THEN TotalTableWidth=TotalTableWidth+50
		IF TRIM(sOF4Desc)<>"" THEN TotalTableWidth=TotalTableWidth+50
		IF TRIM(sOF5Desc)<>"" THEN TotalTableWidth=TotalTableWidth+50									
		IF TRIM(sOF6Desc)<>"" THEN TotalTableWidth=TotalTableWidth+50
		IF TRIM(sOF7Desc)<>"" THEN TotalTableWidth=TotalTableWidth+50
		IF TRIM(sOF8Desc)<>"" THEN TotalTableWidth=TotalTableWidth+50
		IF TRIM(sOF9Desc)<>"" THEN TotalTableWidth=TotalTableWidth+50									
		IF TRIM(sOF10Desc)<>"" THEN TotalTableWidth=TotalTableWidth+50

		%>
		<br>
		<TABLE align=center class="innertable" width=<%=TotalTableWidth%>px>
				<TR>
					<th align=left colspan=8>
						<font size=<% =fontsize3 %> color="<% =Textcolor5 %>"><b>Other Sales Summary</b></font> 
					</th>
				</TR>
				<TR>
					<td align=left colspan=4>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><b>Item Description</b></font> 
					</td>
					<td align=right colspan=2>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><b>Count</b></font> 
					</td>
					<td align=right colspan=2>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><b>Sales $$</b></font> 
					</td>
				</TR>
				<%
		IF rs("BanquetCount")>0 THEN 	
				%>
				<TR>
					<td align=left colspan=4>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>">Banquet Tickets</font> 
					</td>
					<td align=right colspan=2>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=rs("BanquetCount")%></font>
					</td>
					<td align=right colspan=2>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=rs("SumBanquetFee")%></font>
					</td>
				</TR>
				<%
		END IF

		IF rs("OF1QtyCount")>0 THEN 	
				%>
				<TR>
					<td align=left colspan=4>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=sOF1Desc%></font> 
					</td>
					<td align=right colspan=2>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=rs("OF1QtyCount")%></font>
					</td>
					<td align=right colspan=2>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=formatNumber(rs("OF1FeeSum"),2)%></font>
					</td>
				</TR>
				<%
		END IF

		IF rs("OF2QtyCount")>0 THEN 	
				%>
				<TR>
					<td align=left colspan=4>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=sOF2Desc%></font> 
					</td>
					<td align=right colspan=2>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=rs("OF2QtyCount")%></font>
					</td>
					<td align=right colspan=2>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=formatNumber(rs("OF2FeeSum"),2)%></font>
					</td>
				</TR>
				<%
		END IF

		IF rs("OF3QtyCount")>0 THEN 	
				%>
				<TR>
					<td align=left colspan=4>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=sOF3Desc%></font> 
					</td>
					<td align=right colspan=2>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=rs("OF3QtyCount")%></font>
					</td>
					<td align=right colspan=2>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=formatNumber(rs("OF3FeeSum"),2)%></font>
					</td>
				</TR>
				<%
		END IF
		IF rs("OF4QtyCount")>0 THEN 	
				%>
				<TR>
					<td align=left colspan=4>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=sOF4Desc%></font> 
					</td>
					<td align=right colspan=2>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=rs("OF4QtyCount")%></font>
					</td>
					<td align=right colspan=2>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=formatNumber(rs("OF4FeeSum"),2)%></font>
					</td>
				</TR>
				<%
		END IF
		IF rs("OF5QtyCount")>0 THEN 	
				%>
				<TR>
					<td align=left colspan=4>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=sOF5Desc%></font> 
					</td>
					<td align=right colspan=2>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=rs("OF5QtyCount")%></font>
					</td>
					<td align=right colspan=2>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=formatNumber(rs("OF5FeeSum"),2)%></font>
					</td>
				</TR>
				<%
		END IF

		IF rs("OF6QtyCount")>0 THEN 	
				%>
				<TR>
					<td align=left colspan=4>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=sOF6Desc%></font> 
					</td>
					<td align=right colspan=2>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=rs("OF6QtyCount")%></font>
					</td>
					<td align=right colspan=2>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=formatNumber(rs("OF6FeeSum"),2)%></font>
					</td>
				</TR>
				<%
		END IF

		IF rs("OF7QtyCount")>0 THEN 	
				%>
				<TR>
					<td align=left colspan=4>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=sOF7Desc%></font> 
					</td>
					<td align=right colspan=2>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=rs("OF7QtyCount")%></font>
					</td>
					<td align=right colspan=2>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=formatNumber(rs("OF7FeeSum"),2)%></font>
					</td>
				</TR>
				<%
		END IF
		IF rs("OF8QtyCount")>0 THEN 	
				%>
				<TR>
					<td align=left colspan=4>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=sOF8Desc%></font> 
					</td>
					<td align=right colspan=2>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=rs("OF8QtyCount")%></font>
					</td>
					<td align=right colspan=2>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=formatNumber(rs("OF8FeeSum"),2)%></font>
					</td>
				</TR>
				<%
		END IF
		IF rs("OF9QtyCount")>0 THEN 	
				%>
				<TR>
					<td align=left colspan=4>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=sOF9Desc%></font> 
					</td>
					<td align=right colspan=2>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=rs("OF9QtyCount")%></font>
					</td>
					<td align=right colspan=2>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=formatNumber(rs("OF9FeeSum"),2)%></font>
					</td>
				</TR>
				<%
		END IF

		IF rs("OF10QtyCount")>0 THEN 	
				%>
				<TR>
					<td align=left colspan=4>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=sOF10Desc%></font> 
					</td>
					<td align=right colspan=2>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=rs("OF10QtyCount")%></font>
					</td>
					<td align=right colspan=2>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=formatNumber(rs("OF10FeeSum"),2)%></font>
					</td>
				</TR>
				<%
		END IF

%>
</TABLE>
<%

END IF 	 

END SUB



' -------------------------------------
  SUB DisplayBanquet_andOther_Buyers
' -------------------------------------


		ShowBanquetColumn=false
		IF sBTickCost>0 OR sBTickWithE=true THEN ShowBanquetColumn=true
		TotalTableWidth=400
		IF ShowBanquetColumn=true THEN TotalTableWidth=TotalTableWidth+50
		IF TRIM(sOF1Desc)<>"" THEN TotalTableWidth=TotalTableWidth+50
		IF TRIM(sOF2Desc)<>"" THEN TotalTableWidth=TotalTableWidth+50
		IF TRIM(sOF3Desc)<>"" THEN TotalTableWidth=TotalTableWidth+50
		IF TRIM(sOF4Desc)<>"" THEN TotalTableWidth=TotalTableWidth+50
		IF TRIM(sOF5Desc)<>"" THEN TotalTableWidth=TotalTableWidth+50									
		IF TRIM(sOF6Desc)<>"" THEN TotalTableWidth=TotalTableWidth+50
		IF TRIM(sOF7Desc)<>"" THEN TotalTableWidth=TotalTableWidth+50
		IF TRIM(sOF8Desc)<>"" THEN TotalTableWidth=TotalTableWidth+50
		IF TRIM(sOF9Desc)<>"" THEN TotalTableWidth=TotalTableWidth+50									
		IF TRIM(sOF10Desc)<>"" THEN TotalTableWidth=TotalTableWidth+50



sSQL = " SELECT RG.TourID, RG.MemberID, BanquetQty, BanquetFee"
sSQL = sSQL + ", CASE WHEN OF1Qty>0 THEN OF1Qty ELSE '-' END AS OF1Qty, COALESCE(OF2Qty,'-') AS OF2Qty, COALESCE(OF3Qty,'-') AS OF3Qty"
sSQL = sSQL + ", COALESCE(OF4Qty,0) AS OF4Qty, COALESCE(OF5Qty,0) AS OF5Qty, COALESCE(OF6Qty,0) AS OF6Qty"
sSQL = sSQL + ", COALESCE(OF7Qty,0) AS OF7Qty, COALESCE(OF8Qty,0) AS OF8Qty, COALESCE(OF9Qty,0) AS OF9Qty"
sSQL = sSQL + ", COALESCE(OF10Qty,0) AS OF10Qty"
sSQL = sSQL + ", COALESCE(OF1Fee,0) AS OF1Fee, COALESCE(OF2Fee,0) AS OF2Fee, COALESCE(OF3Fee,0) AS OF3Fee"
sSQL = sSQL + ", COALESCE(OF4Fee,0) AS OF4Fee, COALESCE(OF5Fee,0) AS OF5Fee, COALESCE(OF6Fee,0) AS OF6Fee"
sSQL = sSQL + ", COALESCE(OF7Fee,0) AS OF7Fee, COALESCE(OF8Fee,0) AS OF8Fee, COALESCE(OF9Fee,0) AS OF9Fee"
sSQL = sSQL + ", COALESCE(OF10Fee,0) AS OF10Fee"
sSQL = sSQL + ", FirstName, LastName"
sSQL = sSQL + " , CASE WHEN Payments<TotalEntry THEN 'Fees Due' ELSE '' END AS Fee_Status" 
sSQL = sSQL + " FROM "&RegGenTableName&" RG"

sSQL = sSQL + " LEFT JOIN"
sSQL = sSQL + " (SELECT PersonID, FirstName, LastName"
sSQL = sSQL + " FROM "&MemberShortTableName&") MT"
sSQL = sSQL + " ON MT.PersonID=CAST(RIGHT(RG.MemberID,8) AS INT)"

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		( SELECT MemberID, SUM(Amount) AS Payments" 
sSQL = sSQL + "				FROM "&RegPaymentTableName 
sSQL = sSQL + "					WHERE LEFT(TourID,6) = '"&sTourID&"' and Result = '0'" 
sSQL = sSQL + "				GROUP BY MemberID) AS TP" 
sSQL = sSQL + "	ON TP.MemberID = RG.MemberID"

sSQL = sSQL + " WHERE LEFT(RG.TourID,6)='"&LEFT(sTourID,6)&"'"
sSQL = sSQL + " ORDER BY LastName, FirstName"

' response.write(sSQL)
'response.end


Set rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, sConnectionToTRATable, 3, 1

	%>
	<TABLE class="innertable" align="Center" BORDER="1" CELLPADDING="3" CELLSPACING="0" style="background-color:<%=Tablecolor1%>;" width=<%=TotalTableWidth%>px>
		<TR>
			<th align=center width=100px ColSpan="2" valign="top" bgcolor="<%=Headcolor1%>">
				<font size=<%=fontsize2%> color="#FFFFFF">Item</FONT>
			</th>      
			<th align=left ColSpan="6" valign="top" bgcolor="<%=Headcolor1%>">
				<font size=<%=fontsize2%> color="#FFFFFF">&nbsp;Description</FONT>
			</th>      
		</TR>	
		<%
		IF ShowBanquetColumn=true THEN
				%>
				<TR>
					<td align=center colspan=2>
						<font size=<%=fontsize2%> color="<%=Textcolor1%>">BQT</font> 
					</td>
					<td align=left colspan=6>
						<font size=<%=fontsize2%> color="<%=Textcolor1%>">Tournament Banquet</font> 
					</td>
				</TR>
				<%
		END IF

		IF TRIM(sOF1Desc)<>"" THEN
				%>
				<TR>
					<td align=center colspan=2>
						<font size=<%=fontsize2%> color="<%=Textcolor1%>">1</font> 
					</td>
					<td align=left colspan=6>
						<font size=<%=fontsize2%> color="<%=Textcolor1%>"><%=sOF1Desc%></font> 
					</td>
				</TR>
				<%
		END IF
		IF TRIM(sOF2Desc)<>"" THEN
				%>
				<TR>
					<td align=center colspan=2>
						<font size=<%=fontsize2%> color="<%=Textcolor1%>">2</font> 
					</td>
					<td align=left colspan=6>
						<font size=<%=fontsize2%> color="<%=Textcolor1%>"><%=sOF2Desc%></font> 
					</td>
				</TR>
				<%
		END IF
		IF TRIM(sOF3Desc)<>"" THEN
				%>
				<TR>
					<td align=center colspan=2>
						<font size=<%=fontsize2%> color="<%=Textcolor1%>">3</font> 
					</td>
					<td align=left colspan=6>
						<font size=<%=fontsize2%> color="<%=Textcolor1%>"><%=sOF3Desc%></font> 
					</td>
				</TR>
				<%
		END IF
		IF TRIM(sOF4Desc)<>"" THEN
				%>
				<TR>
					<td align=center colspan=2>
						<font size=<%=fontsize2%> color="<%=Textcolor1%>">4</font> 
					</td>
					<td align=left colspan=6>
						<font size=<%=fontsize2%> color="<%=Textcolor1%>"><%=sOF4Desc%></font> 
					</td>
				</TR>
				<%
		END IF
		IF TRIM(sOF5Desc)<>"" THEN
				%>
				<TR>
					<td align=center colspan=2>
						<font size=<%=fontsize2%> color="<%=Textcolor1%>">5</font> 
					</td>
					<td align=left colspan=6>
						<font size=<%=fontsize2%> color="<%=Textcolor1%>"><%=sOF5Desc%></font> 
					</td>
				</TR>
				<%
		END IF
		IF TRIM(sOF6Desc)<>"" THEN
				%>
				<TR>
					<td align=center colspan=2>
						<font size=<%=fontsize2%> color="<%=Textcolor1%>">6</font> 
					</td>
					<td align=left colspan=6>
						<font size=<%=fontsize2%> color="<%=Textcolor1%>"><%=sOF6Desc%></font> 
					</td>
				</TR>
				<%
		END IF
		IF TRIM(sOF7Desc)<>"" THEN
				%>
				<TR>
					<td align=center colspan=2>
						<font size=<%=fontsize2%> color="<%=Textcolor1%>">3</font> 
					</td>
					<td align=left colspan=6>
						<font size=<%=fontsize2%> color="<%=Textcolor1%>"><%=sOF7Desc%></font> 
					</td>
				</TR>
				<%
		END IF
		IF TRIM(sOF8Desc)<>"" THEN
				%>
				<TR>
					<td align=center colspan=2>
						<font size=<%=fontsize2%> color="<%=Textcolor1%>">4</font> 
					</td>
					<td align=left colspan=6>
						<font size=<%=fontsize2%> color="<%=Textcolor1%>"><%=sOF8Desc%></font> 
					</td>
				</TR>
				<%
		END IF
		IF TRIM(sOF9Desc)<>"" THEN
				%>
				<TR>
					<td align=center colspan=2>
						<font size=<%=fontsize2%> color="<%=Textcolor1%>">5</font> 
					</td>
					<td align=left colspan=6>
						<font size=<%=fontsize2%> color="<%=Textcolor1%>"><%=sOF9Desc%></font> 
					</td>
				</TR>
				<%
		END IF
		IF TRIM(sOF10Desc)<>"" THEN
				%>
				<TR>
					<td align=center colspan=2>
						<font size=<%=fontsize2%> color="<%=Textcolor1%>">6</font> 
					</td>
					<td align=left colspan=6>
						<font size=<%=fontsize2%> color="<%=Textcolor1%>"><%=sOF10Desc%></font> 
					</td>
				</TR>
				<%
		END IF

		%>
	</TABLE>
	<br>
	<TABLE class="innertable" align="Center" BORDER="1" CELLPADDING="3" CELLSPACING="0" style="background-color:<%=Tablecolor1%>;" width=<%=TotalTableWidth%>px>
	  <TR>	
	      <th align="Center" ColSpan="2" valign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Name</FONT></th>      
	      <th align="Center" ColSpan="1" valign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">MemberID</FONT></th>      
				<%
				IF ShowBanquetColumn=true THEN
	      		%><th align="Center" ColSpan="1" valign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Banquet</FONT></th><%
	      END IF
				IF TRIM(sOF1Desc)<>"" THEN
	      		%><th align="Center" ColSpan="1" valign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Item 1</FONT></th><%
	      END IF
				IF TRIM(sOF2Desc)<>"" THEN
	      		%><th align="Center" ColSpan="1" valign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Item 2</FONT></th><%
	      END IF
				IF TRIM(sOF3Desc)<>"" THEN
			      %><th align="Center" ColSpan="1" valign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Item 3</FONT></th><%
	      END IF
				IF TRIM(sOF4Desc)<>"" THEN
	      		%><th align="Center" ColSpan="1" valign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Item 4</FONT></th><%
	      END IF
				IF TRIM(sOF5Desc)<>"" THEN
	      		%><th align="Center" ColSpan="1" valign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Item 5</FONT></th><%
	      END IF
				IF TRIM(sOF6Desc)<>"" THEN
	      		%><th align="Center" ColSpan="1" valign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Item 6</FONT></th><%
				END IF
				IF TRIM(sOF7Desc)<>"" THEN
			      %><th align="Center" ColSpan="1" valign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Item 7</FONT></th><%
	      END IF
				IF TRIM(sOF8Desc)<>"" THEN
	      		%><th align="Center" ColSpan="1" valign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Item 8</FONT></th><%
	      END IF
				IF TRIM(sOF9Desc)<>"" THEN
	      		%><th align="Center" ColSpan="1" valign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Item 9</FONT></th><%
	      END IF
				IF TRIM(sOF10Desc)<>"" THEN
	      		%><th align="Center" ColSpan="1" valign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Item 10</FONT></th><%
				END IF
				%><th align="Center" ColSpan="1" valign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Pay Status</FONT></th>
	  </TR>
		<%
		
		OF1QtyTotal=0
		OF2QtyTotal=0
		OF3QtyTotal=0
		OF4QtyTotal=0
		OF5QtyTotal=0
		OF6QtyTotal=0
		OF7QtyTotal=0
		OF8QtyTotal=0
		OF9QtyTotal=0
		OF10QtyTotal=0
		Fee_Status=""
		
		BanquetQtyTotal=0
		DO WHILE NOT(rs.eof)
				MemberID=rs("MemberID")
				FirstName=rs("FirstName")
				LastName=rs("LastName")
				BanquetQty=rs("BanquetQty")
				OF1Qty=rs("OF1Qty")
				OF2Qty=rs("OF2Qty")
				OF3Qty=rs("OF3Qty")
				OF4Qty=rs("OF4Qty")
				OF5Qty=rs("OF5Qty")
				OF6Qty=rs("OF6Qty")
				OF7Qty=rs("OF7Qty")
				OF8Qty=rs("OF8Qty")
				OF9Qty=rs("OF9Qty")
				OF10Qty=rs("OF10Qty")

				OF1Fee=rs("OF1Fee")
				OF2Fee=rs("OF2Fee")
				OF3Fee=rs("OF3Fee")	
				OF4Fee=rs("OF4Fee")
				OF5Fee=rs("OF5Fee")
				OF6Fee=rs("OF6Fee")
				OF7Fee=rs("OF7Fee")	
				OF8Fee=rs("OF8Fee")
				OF9Fee=rs("OF9Fee")
				OF10Fee=rs("OF10Fee")
				Fee_Status=rs("Fee_Status")
			
			' --- Display row only if this condition is true	
			IF BanquetQty>0 OR OF1Qty>0 OR OF2Qty>0 OR OF3Qty>0 OR OF4Qty>0 OR OF5Qty>0 OR OF6Qty>0 OR OF7Qty>0 OR OF8Qty>0 OR OF9Qty>0 OR OF10Qty>0 THEN

				%>
				<TR>
					<td align=left width=150px colspan=2>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=LastName%>,&nbsp;<%=FirstName%></font> 
					</td>
					<td align=center width=100px>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=MemberID%></font> 
					</td>
					<%
					IF ShowBanquetColumn=true THEN
							BanquetQtyTotal=BanquetQtyTotal+BanquetQty
							%>
							<td align=center width=50px>
								<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=BanquetQty%></font> 
							</td>
							<%
					END IF		
					IF TRIM(sOF1Desc)<>"" THEN
							OF1QtyTotal=OF1QtyTotal+OF1Qty
							%>
							<td align=center width=50px>
								<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=OF1Qty%></font> 
							</td>
							<%
					END IF		
					IF TRIM(sOF2Desc)<>"" THEN
							OF2QtyTotal=OF2QtyTotal+OF2Qty
							%>
							<td align=center width=50px>
								<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=OF2Qty%></font> 
							</td>
							<%
					END IF		
					IF TRIM(sOF3Desc)<>"" THEN
							OF3QtyTotal=OF3QtyTotal+OF3Qty
							%>
							<td align=center width=50px>
								<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=OF3Qty%></font> 
							</td>
							<%
					END IF		
					IF TRIM(sOF4Desc)<>"" THEN
							OF4QtyTotal=OF4QtyTotal+OF4Qty
							%>
							<td align=center width=50px>
								<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=OF4Qty%></font> 
							</td>
							<%
					END IF		
					IF TRIM(sOF5Desc)<>"" THEN
							OF5QtyTotal=OF5QtyTotal+OF5Qty
							%>
							<td align=center width=50px>
								<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=OF5Qty%></font> 
							</td>
							<%
					END IF		
					IF TRIM(sOF6Desc)<>"" THEN
							OF6QtyTotal=OF6QtyTotal+OF6Qty
							%>
							<td align=center width=50px>
								<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=OF6Qty%></font> 
							</td>
							<%
					END IF
					IF TRIM(sOF7Desc)<>"" THEN
							OF7QtyTotal=OF7QtyTotal+OF7Qty
							%>
							<td align=center width=50px>
								<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=OF7Qty%></font> 
							</td>
							<%
					END IF		
					IF TRIM(sOF8Desc)<>"" THEN
							OF8QtyTotal=OF8QtyTotal+OF8Qty
							%>
							<td align=center width=50px>
								<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=OF8Qty%></font> 
							</td>
							<%
					END IF		
					IF TRIM(sOF9Desc)<>"" THEN
							OF9QtyTotal=OF9QtyTotal+OF9Qty
							%>
							<td align=center width=50px>
								<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=OF9Qty%></font> 
							</td>
							<%
					END IF		
					IF TRIM(sOF10Desc)<>"" THEN
							OF10QtyTotal=OF10QtyTotal+OF10Qty
							%>
							<td align=center width=50px>
								<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=OF10Qty%></font> 
							</td>
							<%
					END IF
					%>
					<td align=center width=70px>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=Fee_Status%></font> 	
					</td>
				</TR>
				<%
			END IF
			rs.MoveNext	

		LOOP
				
		%>
				<TR>
					<td align=left width=150px colspan=2>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><b>TOTAL ALL</b></font> 
					</td>
					<td align=center width=100px>
						<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><b>&nbsp;</b></font> 
					</td>
					<%
					IF ShowBanquetColumn=true THEN
							%>
							<td align=center width=50px>
								<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><b><%=BanquetQtyTotal%></b></font> 
							</td>
							<%
					END IF		
					IF TRIM(sOF1Desc)<>"" THEN
							%>
							<td align=center width=50px>
								<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><b><%=OF1QtyTotal%></b></font> 
							</td>
							<%
					END IF		
					IF TRIM(sOF2Desc)<>"" THEN
							%>
							<td align=center width=50px>
								<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><b><%=OF2QtyTotal%></b></font> 
							</td>
							<%
					END IF		
					IF TRIM(sOF3Desc)<>"" THEN
							%>
							<td align=center width=50px>
								<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><b><%=OF3QtyTotal%></b></font> 
							</td>
							<%
					END IF		
					IF TRIM(sOF4Desc)<>"" THEN
							%>
							<td align=center width=50px>
								<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><b><%=OF4QtyTotal%></b></font> 
							</td>
							<%
					END IF		
					IF TRIM(sOF5Desc)<>"" THEN
							%>
							<td align=center width=50px>
								<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><b><%=OF5QtyTotal%></b></font> 
							</td>
							<%
					END IF		
					IF TRIM(sOF6Desc)<>"" THEN
							%>
							<td align=center width=50px>
								<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><b><%=OF6QtyTotal%></b></font> 
							</td>
							<%
					END IF
					IF TRIM(sOF7Desc)<>"" THEN
							%>
							<td align=center width=50px>
								<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><b><%=OF7QtyTotal%></b></font> 
							</td>
							<%
					END IF		
					IF TRIM(sOF8Desc)<>"" THEN
							%>
							<td align=center width=50px>
								<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><b><%=OF8QtyTotal%></b></font> 
							</td>
							<%
					END IF		
					IF TRIM(sOF9Desc)<>"" THEN
							%>
							<td align=center width=50px>
								<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><b><%=OF9QtyTotal%></b></font> 
							</td>
							<%
					END IF		
					IF TRIM(sOF10Desc)<>"" THEN
							%>
							<td align=center width=50px>
								<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><b><%=OF10QtyTotal%></b></font> 
							</td>
							<%
					END IF
					%>
					<td>&nbsp;</td>
				</TR>		
		</TABLE>	
		<%



END SUB



' -------------------
  SUB SeedingHeading
' -------------------



	%>
	<TABLE class="innertable" align="Center" BORDER="1" CELLPADDING="3" CELLSPACING="0" style="background-color:<%=Tablecolor1%>;" width=100%>

	  <TR>	
	      <th align="Center" ColSpan="4" valign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Personal</FONT></th>      
	      <th align="Center" ColSpan="2" valign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Entry Selections</FONT></th>      
	      <th align="Center" ColSpan="5" valign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Seeding & Placement</FONT></th>
	  </TR>

	  <TR><%
		SeedColWidth=55	%>		
	      <th align="Center"  bgcolor="<%=Headcolor1%>" valign="top"><font size=<%=fontsize2%> color="#FFFFFF">Num</FONT></th>
	      <th align="Left"  bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Name</FONT></th>
	      <th align="center" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">MemberID</FONT></th>
	      <th align="Center" width="<%=SeedColWidth%>" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF"><a TITLE="State">ST</a></FONT></th>

	      <th align="Center" width="<%=SeedColWidth%>" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">
			<a TITLE="Entry Classification is coded G-Grassroots S-ClassC (or Base Class) and R-Record (Upgrade)">Class</a></FONT></th>
	      <th align="Center" width="<%=SeedColWidth%>" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">
			<a TITLE="Skill Level is used for ability-based grouping of competitors">Skill</a></FONT></th>
	      <th align="Center" width="<%=SeedColWidth%>" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">
			<a TITLE="Seeding Value - Same as Ranking Value - Also the order of participation, highest value skis last">Rank<br>Score</a></FONT></th>
	      <th align="Center" width="<%=SeedColWidth%>" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">
			<a TITLE="Current Qualification Level based on position on Rankings List">Rank<br>Level</a></FONT></th>
	      <th align="Center" width="<%=SeedColWidth%>" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">
			<a TITLE="Level from Ranking List to be Used as Qualification Method beginning in 2008">Rank<br>Pctl</a></FONT></th>
	      <th align="Center" width="<%=SeedColWidth%>" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">
			<a TITLE="Most Recent National Placement">Natl<br>Place</a></FONT></th>
	      <th align="Center" width="<%=SeedColWidth%>" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">
			<a TITLE="Most Recent Regional Placement">Regl<br>Place</a></FONT></th>
	  </TR><%

END SUB



' --------------------
  SUB RegStatHeading
' --------------------

%>
	<TABLE class="innertable" align="Center" BORDER="1" CELLPADDING="3" CELLSPACING="0" width=100%>

	  <TR>	
	      <th align="Center" ColSpan="4" valign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Personal</FONT></th>      
	      <th align="Center" ColSpan="10" valign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Registration Status</FONT></th>
	  </TR>
	  <TR><%		

	     IF SequenceSelected="seed" THEN%>			
	      <th align="Center" valign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Seed</FONT></th><%
	     ELSE %>
	      <th align="Center" bgcolor="<%=Headcolor1%>" valign="top"><font size=<%=fontsize2%> color="#FFFFFF">Num</FONT></th><%
	     END IF %>	
	      <th align="Left" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Name</FONT></th>
	      <th align="Left" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">MemberID</FONT></th>
	      <th align="Center" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF"><a TITLE="State">ST</a></FONT></th>

	      <th align="Center" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">
			<a TITLE="Status of Membership including Expiration Date and Membership Type">Mem</a></FONT></th>
	      <th align="Center" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">
			<a TITLE="Status of Fees, Charges and Payments">Fees</a></FONT></th>
	      <th align="Center" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">
			<a TITLE="Qualifications by Rating or Placement">Qlfy</a></FONT></th>
	      <th align="Center" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">
			<a TITLE="Waiver & Release of Liability">Waiv</a></FONT></th>
		<% IF sForm2Name="list" AND (EventSelected = "T" OR EventSelected="ALL") THEN %>
		      <th align="Center" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">
			<a TITLE="Trick List">Trick</a></FONT></th>
		<% END IF %>
	      <th align="Center" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">
			<a TITLE="Skier Personal Bio Sheet">Bio</a></FONT></th>
	      <th align="Center" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">
			<a TITLE="Participation in most recent Regional tournament - This is a Requirement for all divisions except Open">Ski Reg</a></FONT></th>
	      <th align="Center" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">
			<a TITLE="Date the Registration was received/entered">Reg Date</a></FONT></th><%

		IF (adminmenulevel>=20 OR TestValidAdminCode=true) THEN %>

		      <th align="Center" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">
			<a TITLE="Number of Banquet Tickets if Tournament offers banquet">Banq</a></FONT></th><%
			IF LEFT(sTourID,6)="08M103" THEN %>
			      <th align="Center" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">
				<a TITLE="Fee Class">Code</a></FONT></th><%
			END IF %>
			
		      <th align="Center" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Form</FONT></th><%
		END IF %>
	   </TR><%

END SUB



' -------------------
  SUB SkiPayHeading
' -------------------


NumFeeColumns=TotNumOptItems+4
  	  %>
	<TABLE class="innertable" width=100%>
	  <TR>	
	      <th align="Center" ColSpan="3" valign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Personal</FONT></th>      
	      <th align="Center" ColSpan="<%=NumFeeColumns%>" valign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Fees</FONT></th>
	      <th align="Center" ColSpan="4" valign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Discounts</FONT></th>
	      <th align="Center" ColSpan="2" valign="top" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF">Summary</FONT></th>
	  </TR>

	  <TR>			
	      <th align="Left"><font size=<%=fontsize2%> color="#FFFFFF">Name</FONT></th>
	      <th align="Left"><font size=<%=fontsize2%> color="#FFFFFF">MemberID</FONT></Center></th>
	      <th align="Left"><font size=<%=fontsize2%> color="#FFFFFF">Type</FONT></Center></th>

	      <th align="Left"><font size=<%=fontsize2%> color="#FFFFFF">Entry</FONT></Center></th>
	      <th align="Left"><font size=<%=fontsize2%> color="#FFFFFF">Late</FONT></Center></th>
	      <th align="Left"><font size=<%=fontsize2%> color="#FFFFFF">AWSEF</FONT></Center></th>
	      <th align="Left"><font size=<%=fontsize2%> color="#FFFFFF">Banq</FONT></Center></th>
				<%

	      IF TRIM(sOF1Desc)<>"" THEN 
	      		%><th align="Left"><font size=<%=fontsize2%> color="#FFFFFF">Opt1</FONT></Center></th><%
	      END IF
	      IF TRIM(sOF2Desc)<>"" THEN 	      
	      		%><th align="Left"><font size=<%=fontsize2%> color="#FFFFFF">Opt2</FONT></Center></th><%
	      END IF
	      IF TRIM(sOF3Desc)<>"" THEN 	      
	      		%><th align="Left"><font size=<%=fontsize2%> color="#FFFFFF">Opt3</FONT></Center></th><%
	      END IF
	      IF TRIM(sOF4Desc)<>"" THEN 	      
	      		%><th align="Left"><font size=<%=fontsize2%> color="#FFFFFF">Opt4</FONT></Center></th><%
	      END IF
	      IF TRIM(sOF5Desc)<>"" THEN 	      
			      %><th align="Left"><font size=<%=fontsize2%> color="#FFFFFF">Opt5</FONT></Center></th><%
	      END IF
	      IF TRIM(sOF6Desc)<>"" THEN 	      
						%><th align="Left"><font size=<%=fontsize2%> color="#FFFFFF">Opt6</FONT></Center></th><%
	      END IF
	      IF TRIM(sOF7Desc)<>"" THEN 	      
						%><th align="Left"><font size=<%=fontsize2%> color="#FFFFFF">Opt7</FONT></Center></th><%
	      END IF
	      IF TRIM(sOF8Desc)<>"" THEN 	      
						%><th align="Left"><font size=<%=fontsize2%> color="#FFFFFF">Opt8</FONT></Center></th><%
	      END IF
	      IF TRIM(sOF9Desc)<>"" THEN 	      
						%><th align="Left"><font size=<%=fontsize2%> color="#FFFFFF">Opt9</FONT></Center></th><%
	      END IF
	      IF TRIM(sOF10Desc)<>"" THEN 	      
						%><th align="Left"><font size=<%=fontsize2%> color="#FFFFFF">Opt10</FONT></Center></th><%
	      END IF
				
				%>
	      <th align="Left"><font size=<%=fontsize2%> color="#FFFFFF">Off Disc</FONT></Center></th>
	      <th align="Left"><font size=<%=fontsize2%> color="#FFFFFF">Jr Disc</FONT></Center></th>
	      <th align="Left"><font size=<%=fontsize2%> color="#FFFFFF">Sr Disc</FONT></Center></th>
	      <th align="Left"><font size=<%=fontsize2%> color="#FFFFFF">Club Disc</FONT></Center></th>
	      <th align="Left"><font size=<%=fontsize2%> color="#FFFFFF">Total</FONT></Center></th>

	      <th align="Left"><font size=<%=fontsize2%> color="#FFFFFF">Bal Due</FONT></Center></th>
	  </TR><%

END SUB


' --------------------
  SUB ByStateHeading
' --------------------
  	  %>
	<TABLE align="Center" class="innertable" width=100%>

	  <TR>			
	      <th align="Left"><font size=<%=fontsize2%> color="#FFFFFF"><b>&nbsp;Name</b></FONT></th>
	      <th align="center"><font size=<%=fontsize2%> color="#FFFFFF"><b>&nbsp;MemberID</b></FONT></Center></th>
	      <th align="Left"><font size=<%=fontsize2%> color="#FFFFFF"><b>&nbsp;City/State</b></FONT></Center></th>
	      <th align="center" width=60px><font size=<%=fontsize2%> color="#FFFFFF"><b>&nbsp;<%=sTEventName(1)%></b></FONT></Center></th>
	      <th align="center" width=60px><font size=<%=fontsize2%> color="#FFFFFF"><b>&nbsp;<%=sTEventName(2)%></b></FONT></Center></th>
	      <th align="center" width=60px><font size=<%=fontsize2%> color="#FFFFFF"><b>&nbsp;<%=sTEventName(3)%></b></FONT></Center></th>
	      <% IF sTEvent4<>"" THEN %>
	         <th align="center" width=60px><font size=<%=fontsize2%> color="#FFFFFF"><b>&nbsp;<%=sTEventName(4)%></b></FONT></Center></th>
	      <% END IF %>	
	  </TR><%

END SUB


' --------------------
  SUB DropBoxFormat1
' --------------------

	
	' --- Defines the Report Title ---
	SELECT CASE sWhatReport
		CASE "seeding"
				ReportTitle = "Seeding Summary"&sTestMode	
		CASE "skierpayments"
				ReportTitle = "Payments By Type"&sTestMode	
		CASE "regstat"
				ReportTitle = "Registration Status"&sTestMode	
		CASE "bystate"
				ReportTitle = "Participants By State"&sTestMode	
		CASE "scratched"
				ReportTitle = "Scratch List"&sTestMode	
		CASE "othersales"
				ReportTitle = "Other Sales Report"&sTestMode	
		
	END SELECT


	' --- Looks up image to display in header based on link between TSiteID in TSchedul

	Set rs=Server.CreateObject("ADODB.recordset")

' --- Old ---
'	sSQL= " SELECT HeaderImage FROM usawsrank.TourExtras WHERE LEFT(TourID,6)='"&LEFT(sTourID,6)&"'"

' --- New ---
	sSQL = "SELECT HeaderImage FROM sanctions.dbo.TSchedul AS TS"
	sSQL = sSQL + " JOIN usawsrank.TourExtras AS TE"
	sSQL = sSQL + "   ON SiteID=TS.TSiteID"
	sSQL = sSQL + " WHERE TournAppID='"&sTourID&"'"

	rs.open sSQL, sConnectionToTRATable, 3, 1


	' --- Uses a default image if it does not find a site image ---
	IF (NOT rs.eof) AND (NOT PrintButton="Printer Friendly") THEN 
			MainImage="images\LOCSites\"&TRIM(rs("HeaderImage"))
			TextDropcolor1="#FFFFFF"
			TextDropcolor2=Textcolor2
	ELSE
			TextDropcolor1="#000000"	
			TextDropcolor2=Textcolor2
			MainImage="images\LOCSites\AMFog.jpg"			
	END IF


	%>
     <form action="/rankings/<%=ThisFileName%>" method="post">

	<TABLE align="center" class="droptable" WIDTH="740px" height=180px background="<%=MainImage%>">
	  <tr>
			<td align="left" colspan=4 width=60%>
				<font size=3 color="<% =TextDropcolor1 %>"><b>&nbsp;&nbsp;<% =sTourName %></b>&nbsp;&nbsp;&nbsp;&nbsp</font>
				<font size=<%=fontsize2%> color="<% =TextDropcolor2 %>"><br>&nbsp;&nbsp;<% =sTDateS %>-<% =sTDateE %> - <%=sTourID%></b></font>
			</td>	
			<td colspan=2 align="left" width=40%>
				<FONT color="<%=TextDropcolor1%>" size=3><B><% Response.Write(ReportTitle) %></B></FONT>
				<br><br>
			</td>
	 </tr>

	 <tr><%
		' --- Loads List of report options ---
		LoadReportPullDown  

		' --- Loads divisions offered in this event ---
	   LoadDivPulldown 
	    	%>
	    <td width=100px>&nbsp;</td>
	    <td width=100px>&nbsp;</td>
	 </tr>

	 <tr><%

	   ' --- Loads drop down to set sequence ---
	   LoadDropSequence

	   ' --- Loads Event Drop down ---	
	   LoadEventPulldownNew %>
	    <td>&nbsp;</td>
	    <td>&nbsp;</td>
	  </tr>

	 <tr><%

		' --- State or Region dropdown depending on sWhatReport selected
		IF sWhatReport="bystate" THEN 
			LoadStatePulldown 
		ELSE 	
			LoadRegionPulldown 
		END IF %>
	    <td>&nbsp;</td>
	    <td>&nbsp;</td><%

		IF AdminMenuLevel>=50 THEN  %>	
  			<td colspan=1 valign=top align="left">
				<FONT color="<% =Titlecolor %>" size="<%=fontsize2%>"><b>Show SQL</b></font>
				<input type=checkbox name="sShowSQL" <% IF sShowSQL="on" THEN response.write "checked" %>>

			</td><%
		ELSE  %>
			<td>&nbsp;</td><%
		END IF %>

	    <td>&nbsp;</td>
	  </tr>

	  <tr>	
	   <td align=center colspan=2><%
	      IF PrintButton="Printer Friendly" THEN %>	
		      <input type="submit" align="center" style="width:10em;"  value="Report Update"><%
	      ELSE %>
		      <input type="submit" align="center" style="width:10em;"  value="Display Report"><%
	      END IF %>	
	   </td>

	   <td align=center colspan=2><%
	      IF PrintButton="Printer Friendly" THEN %>	
			<a href='#' onclick='window.print()' title="Click here to Print"><input type=submit value="Print Now" style="width:9em"></a><%
	      ELSE %>
			<input type="submit" align="center" style="width:10em;" name="PrintButton" value="Printer Friendly"><%
	      END IF %>
	   </td>

	   <td align=center colspan=2><%
	      IF PrintButton="Printer Friendly" THEN 
			' --- Don't Display
	      ELSE %>
			<input type="submit" align="center" style="width:10em;" name="ReturnButton" value="Main Menu"><%
	      END IF %>
	   </td>


	 </tr>
	</table>
	</form>

	<%

END SUB



' ----------------------
  SUB RunByStateReport
' ----------------------


	'Dim sDiv(10)	
	rs.movefirst
	LastFullName=rs("FirstName")&" "&rs("LastName")
	LastMemb=rs("MemberID")
	LastCity=rs("City")
	LastState=rs("State")
	LastEvent=rs("Event")



	' --- sDiv(X) is actually representing LASTDiv ---
	SELECT CASE TRIM(rs("Event"))
		CASE sTEvent(1)
			sDiv(1)=rs("Div")
		CASE sTEvent(2)
			sDiv(2)=rs("Div")
		CASE sTEvent(3)
			sDiv(3)=rs("Div")
		CASE sTEvent(4)
			sDiv(4)=rs("Div")
	END SELECT

	
	Dim C
	C=1
	DO WHILE C=1

		IF NOT rs.eof THEN
			CurMemb=rs("MemberID")
		ELSE
			CurMemb=" "
		END IF

		IF LastMemb<>CurMemb THEN %>

			<TR>
			  <TD align="Left" >
			    <font size=<%=fontsize2%> color="<% =Textcolor1 %>">
			    <a href="/rankings/view-scoresHQ.asp?NSL=<% IF Request("Pvar") = "NSL" THEN Response.Write("1") ELSE Response.Write("0")%>&sMemberID=<%=LastMemb%>&event=<%=LastEvent%>&pvar=ByMember">
				<font size=<%=fontsize2%> face=<% =font2 %>>&nbsp;<% =LastFullName %><% IF LastMemb = "000001151" OR LastMemb = "100000850" THEN response.write(" - TEST") %></a></FONT>
			  </TD>
		          <TD align="Center" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>"><%=LastMemb%></FONT></TD>
	        	  <TD align="Left" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;&nbsp;<%=LastCity%>, <%=LastState%></FONT></TD>
	        	  <TD align="Center" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<%=sDiv(1)%></FONT></TD>
	        	  <TD align="Center" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<%=sDiv(2)%></FONT></TD>
	        	  <TD align="Center" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<%=sDiv(3)%></FONT></TD>
			  <% IF sTEvent(4)<>"" THEN %>
	        	  <TD align="Center" ><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<%=sDiv(4)%></FONT></TD>
			  <% END IF %>
			</TR><%
			

			IF rs.eof THEN 
				EXIT DO
			ELSE
				LastFullName=rs("FirstName")&" "&rs("LastName")
				LastMemb=rs("MemberID")
				LastCity=rs("City")
				LastState=rs("State")
				LastEvent=rs("Event")
				sDiv(1)=""
				sDiv(2)=""
				sDiv(3)=""
				sDiv(4)=""

				SELECT CASE TRIM(rs("Event"))
					CASE sTEvent(1)
						sDiv(1)=rs("Div")
					CASE sTEvent(2)
						sDiv(2)=rs("Div")
					CASE sTEvent(3)
						sDiv(3)=rs("Div")
					CASE sTEvent(4)
						sDiv(4)=rs("Div")
				END SELECT

			END IF
		
		END IF

		rs.MoveNext		

	LOOP %>

	</TABLE><%

	rs.close
	Set rs = nothing
'	CloseCon

'response.end

	NewsPageNum="10RegRep"
	WriteIndexPageFooter


END SUB


' -----------------
  SUB SendMarkEmail
' -----------------


eMailTo = marksemailaddress
eMailCC = ""
eMailBCC = ""
eMailFrom = "USA Water Ski<competition@usawaterski.org>"
eMailSubj = "Error in determining SkiYearID"
eMailBody="There was an error in determining the Ski Year ID - sTourID = "&sTourID  


SetupEmailService

objMessage.To = eMailTo
objMessage.cc = eMailCC
objMessage.bcc = eMailBCC
objMessage.From = eMailFrom
objMessage.Subject = eMailSubj
objMessage.HTMLBody = eMailBody
 
 ' --- Finally send the message, and then clear that object
IF TRIM(sMailTo)<>"" THEN
		objMessage.Send
END IF
set objMessage = Nothing




END SUB



' ----------------------
  SUB BuildSeedingQuery
' ----------------------


	' ----------------------------------------------------------------------------------------------------------
	' -----------  Builds SQL string to define display values  -------------------------------------------------
	' ----------------------------------------------------------------------------------------------------------

	Set rs=Server.CreateObject("ADODB.recordset")

	' --- Set SkiYear first and if tournament is in 12 Month range then SkiYear=1 will be the top row of answerset ---
	Dim RankSkiYear
	sSQL = "SELECT SkiYearID FROM usawsrank.SkiYear"
	sSQL = sSQL + "	WHERE BeginDate<='"&sTDateS&"' AND EndDate>='"&sTDateS&"'"
	rs.open sSQL, sConnectionToTRATable, 3, 1

	IF rs.eof THEN
		SendMarkEmail		
		RankSkiYear=1
	ELSE
		RankSkiYear=rs("SkiYearID")
			
	END IF


	IF DateDiff("d",Date,sTDateS)<=365 AND DateDiff("d",Date,sTDateS)>=0 THEN RankSkiYear=1
	
'	response.write("<br>"& DateDiff("d",Date,sTDateS)) 
'	response.write("<br>RankSkiYear="&RankSkiYear)
'	response.write("<br>sSQL="&sSQL)
'	response.end

	rs.close


' --------------------------------
' --- Begin SQL Query contruct ---
' --------------------------------

sSQL = " SELECT EVT.MemberID, EVT.div, EVT.event, EVT.QfyOverride, EVT.FeeClass, EVT.Skill" 

sSQL = sSQL + ", RGEN.RegDate, RGEN.EntryType, RGEN.WaiverCode, RGEN.TotalEntry, RGEN.BanquetQty, RGEN.MembOverride, RGEN.RegionalOverride" 
sSQL = sSQL + ", RGEN.MoneyOverride, coalesce(RGEN.SentBioEmail,'N') AS SentBioEmail" 

sSQL = sSQL + ", RT.Rank, coalesce(RT.RankPct,0) as RankPct, coalesce(RT.natl_plc,' ') AS natl_plc, coalesce(RT.regl_plc,' ') AS regl_plc, RT.Reg_Ski" 
sSQL = sSQL + ", RT.Rating, RT.Reg_Ski" 

sSQL = sSQL + ", coalesce(RQ.SkiedRegls,' ') AS SkiedRegls, RQ.QfyStatusTextNew" 
sSQL = sSQL + ", coalesce(TP.Payments, 0) AS Payments" 

sSQL = sSQL + ", PW.Email" 

sSQL = sSQL + ", MEM.firstname, MEM.lastname, MEM.EffectiveTo, MEM.MemberShipTypeCode, UPPER(MEM.[state]) AS 'state', MEM.City, MEM.NoEmail" 
sSQL = sSQL + ", MEM.MembEmail" 

sSQL = sSQL + ", MTT.MemberType, MTT.MemberShipTypeID, MTT.CanSkiInTournaments, MTT.CanSkiInGRTournaments, MTT.MembTypeDesc" 

sSQL = sSQL + ", BIO.BioMemberID" 

sSQL = sSQL + ", TGEN.Form1Name, TGEN.Form2Name, TGEN.Form3Name, TGEN.Form4Name, TGEN.Form5Name, TGEN.Form6Name, TGEN.EmailAddress, TGEN.QualLevel, TGEN.Bio_Reqd" 

sSQL = sSQL + ", REGION.region" 

sSQL = sSQL + ", coalesce(LT.RequirePart,'-') AS RequirePart, LT.LT_LeagueID" 


sSQL = sSQL + "	FROM "&RegDetailTableName&" AS EVT" 

sSQL = sSQL + "	JOIN" 
sSQL = sSQL + "		( SELECT MemberID, TourID, RegisterDate as 'RegDate', EntryType, WaiverCode, TotalEntry, BanquetQty, MembOverride, RegionalOverride," 
sSQL = sSQL + "				MoneyOverride, coalesce(SentBioEmail,'N') AS SentBioEmail" 
sSQL = sSQL + "			FROM "&RegGenTableName&") AS RGEN" 
sSQL = sSQL + "	ON EVT.MemberID = RGEN.MemberID AND LEFT(EVT.TourID,6) = LEFT(RGEN.TourID,6)" 

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		( SELECT MemberID, Event, Div, SC_3, SkiYearID, RankScore as 'Rank', coalesce(RankPct,0) as RankPct," 
sSQL = sSQL + "			coalesce(natl_plc,' ') AS natl_plc, coalesce(regl_plc,' ') AS regl_plc, Reg_Ski, AWSA_Rat AS 'Rating'" 
sSQL = sSQL + "			FROM "&RankTableName
sSQL = sSQL + "				WHERE SkiYearID='1' AND SC_3 IS NULL AND Event IS NOT NULL AND Div IS NOT NULL) AS RT" 
sSQL = sSQL + "	ON EVT.MemberID=RT.MemberID AND EVT.Div=RT.Div AND EVT.Event=RT.Event"  

' --- Gets total amount paid by this member in this tournament
sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		( SELECT MemberID, SUM(Amount) AS Payments" 
sSQL = sSQL + "				FROM "&RegPaymentTableName 
sSQL = sSQL + "					WHERE LEFT(TourID,6) = '"&sTourID&"' and Result = '0'" 
sSQL = sSQL + "				GROUP BY MemberID) AS TP" 
sSQL = sSQL + "	ON TP.MemberID = EVT.MemberID"

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		( SELECT MemberID, TourID, Event, Div, coalesce(SkiedRegls,' ') AS SkiedRegls, QfyStatus AS QfyStatusTextNew" 
sSQL = sSQL + "			FROM "&RegQualifyTableName&" ) AS RQ" 
sSQL = sSQL + "	ON EVT.MemberID=RQ.MemberID AND LEFT(EVT.TourID,6)=LEFT(RQ.TourID,6) AND EVT.Event=RQ.Event AND EVT.Div=RQ.Div" 

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		( SELECT PersonID, firstname, lastname, EffectiveTo, MemberShipTypeCode, UPPER(state) AS 'state', City, DoNotEmail AS NoEmail, Email AS MembEmail" 
sSQL = sSQL + "			FROM "&MemberLiveTableName&") AS MEM" 
sSQL = sSQL + "	ON CAST(RIGHT(RGEN.MemberID,8) AS INT) = MEM.PersonID" 

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		( SELECT MembershipTypeID, TypeCode AS 'MemberType', CanSkiInTournaments, CanSkiInGRTournaments, Description AS MembTypeDesc" 
sSQL = sSQL + "			FROM "&MemberTypeOLRTableName&") AS MTT" 
sSQL = sSQL + "	ON MEM.MembershipTypeCode = MTT.MembershipTypeID" 
	
sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		( SELECT MemberID, Email"
sSQL = sSQL + "			FROM "&RegPWTableName&" WHERE MemberID IS NOT NULL) AS PW" 
sSQL = sSQL + "	ON PW.MemberID = RGEN.MemberID" 
	
sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		( SELECT State, region" 
sSQL = sSQL + "			FROM "&RegionTableName&") AS REGION" 
sSQL = sSQL + "	ON LOWER(MEM.[state]) = LOWER(REGION.[state])" 
	
sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		( SELECT MemberID AS BioMemberID"
sSQL = sSQL + "			FROM "&BioTableName&") AS BIO" 
sSQL = sSQL + "	ON RGEN.MemberID = BIO.BioMemberID" 

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		( SELECT TournAppID, Form1Name, Form2Name, Form3Name, Form4Name, Form5Name, Form6Name, EmailAddress, QualLevel, Bio_Reqd"
sSQL = sSQL + "			FROM "&TRegSetupTableName&") AS TGEN" 
sSQL = sSQL + "	ON LEFT(TGEN.TournAppID,6) = LEFT(RGEN.TourID,6)" 

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		( SELECT QualifyTour, coalesce(RequirePart,'-') AS RequirePart, LeagueID AS LT_LeagueID"  
sSQL = sSQL + "			FROM "&LeagueTableName&") AS LT" 
sSQL = sSQL + "	ON TGEN.TournAppID = LEFT(LT.QualifyTour,6)" 




	' -----------------------------------
	' ------ Begin WHERE condition ------
	' -----------------------------------

	sSQL = sSQL + " WHERE LEFT(RGEN.[TourID],6) = '"&LEFT(sTourID,6)&"'"

	IF DivSelected = "ALL" THEN
		'sSQL = sSQL + " AND (EVT.div IN ('MM', 'OM', 'OW', 'B1', 'B2','B3', 'G1', 'G2', 'G3', 'M1', 'M2', 'M3', 'M4', 'M5', 'M6', 'M7', 'M8', 'M9', 'MA', 'MB', 'W1', 'W2', 'W3','W4', 'W5', 'W6','W7', 'W8', 'W9', 'WA', 'WB'))"
	ELSE
		sSQL = sSQL + " AND EVT.div = '"&DivSelected&"'"
	END IF

	IF EventSelected = "ALL" THEN 
		'sSQL = sSQL + " AND (EVT.event IN ('S', 'T', 'J', 'WB', 'WS', 'WU', 'KS', 'KT', 'KF', 'KR'))"
	ELSE
		sSQL = sSQL + " AND (EVT.event = '"&EventSelected&"')"
	END IF

	' --- First and Last Characters when report gets too big
	IF StartCharSelected<>"All" THEN
		sSQL = sSQL + " AND (LEFT(MEM.LastName,1)>='"&StartCharSelected&"')"		
	END IF
	
	IF EndCharSelected<>"All" THEN
		sSQL = sSQL + " AND (LEFT(MEM.LastName,1)<='"&EndCharSelected&"')"		
	END IF

	IF TRIM(sPrintDate)<>"" THEN
		sSQL = sSQL + " AND RGEN.RegisterDate='"&sPrintDate&"'"			
	END IF


	IF RegionSelected <> "6" THEN sSQL = sSQL + " AND REGION.[region] = '"&RegionSelected&"'"
	IF StateSelected <> "All" THEN sSQL = sSQL + " AND MEM.State = '"&StateSelected&"'"


	' ------------------------------------
	' ------ Sets ORDER of Display  ------
	' ------------------------------------

	IF sWhatReport="notifications" THEN
		sSQL = sSQL + " ORDER BY MEM.LastName, MEM.FirstName, EVT.event"
	ELSEIF sWhatReport="bystate" THEN
		sSQL = sSQL + " ORDER BY MEM.State, MEM.LastName, MEM.FirstName"
	ELSEIF SequenceSelected="alpha" THEN
		sSQL = sSQL + " ORDER BY EVT.div, EVT.event, MEM.LastName, MEM.FirstName"
	ELSEIF SequenceSelected="seed" OR SequenceSelected="regstat" THEN
			sSQL = sSQL + " ORDER BY EVT.div, EVT.event, RT.Rank DESC, MEM.LastName, MEM.FirstName"
	ELSEIF SequenceSelected="regdate" THEN
			sSQL = sSQL + " ORDER BY EVT.div, EVT.event, RegDate, MEM.LastName, MEM.FirstName"
	ELSEIF SequenceSelected="regdateall" THEN
			sSQL = sSQL + " ORDER BY RegDate, MEM.LastName, MEM.FirstName"
	ELSE
		' sSQL = sSQL + " ORDER BY EVT.div, EVT.event, MEM.LastName, MEM.FirstName"
	END IF

	Set rs=Server.CreateObject("ADODB.recordset")
	'rs.CommandTimeout = 90
	'rs.ConnectionTimeout = 90
	rs.open sSQL, sConnectionToTRATable, 3, 1

  'Set rs = Server.CreateObject("ADODB.Connection")
  'rs.ConnectionTimeout = 2000
  'rs.Open Application("sConnectionToTRATable")
  'rs.CommandTimeout = 2000


	IF Session("adminmenulevel")>=50 AND TRIM(sShowSQL)<>"" THEN
		response.write("<br>"&sSQL)
'		response.end
	END IF

	IF rs.eof THEN sWhatReport = "EndofFile"




END SUB





' ----------------------
  SUB BuildPaymentsQuery
' ----------------------

	' ----------------------------------------------------------------------------------------------------------
	' -----------  Builds SQL string to define display values  -------------------------------------------------
	' ----------------------------------------------------------------------------------------------------------

	Set rs=Server.CreateObject("ADODB.recordset")

	' --- Set SkiYear first and if tournament is in 12 Month range then SkiYear=1 will be the top row of answerset ---
	Dim RankSkiYear
	sSQL = "SELECT SkiYearID FROM usawsrank.SkiYear"
	sSQL = sSQL + "	WHERE BeginDate<='"&sTDateS&"' AND EndDate>='"&sTDateS&"'"
	rs.open sSQL, sConnectionToTRATable, 3, 1

	IF rs.eof THEN
			SendMarkEmail		
			RankSkiYear=1
	ELSE
			RankSkiYear=rs("SkiYearID")
	END IF

	IF DateDiff("d",Date,sTDateS)<=365 AND DateDiff("d",Date,sTDateS)>=0 THEN RankSkiYear=1
	
	rs.close
	Set rs=Server.CreateObject("ADODB.recordset")

	sSQL = "SELECT EVT.div, EVT.event, EVT.QfyOverride, EVT.FeeClass, EVT.Skill"

	sSQL = sSQL + ", RGEN.RegisterDate as 'RegDate', RGEN.EntryType, RGEN.MemberID, RGEN.WaiverCode, RGEN.TotalEntry"
	sSQL = sSQL + ", RGEN.BanquetQty"
	sSQL = sSQL + ", RGEN.MembOverride, RGEN.RegionalOverride, RGEN.MoneyOverride, coalesce(RGEN.SentBioEmail,'N') AS SentBioEmail"

	sSQL = sSQL + ", coalesce(TP.Payments, 0) as Payments"

	sSQL = sSQL + ", MEM.firstname, MEM.lastname, MEM.EffectiveTo, MEM.MemberShipTypeCode, UPPER(MEM.[state]) AS 'state', MEM.City, MEM.DoNotEmail AS NoEmail"
	sSQL = sSQL + ", MEM.Email AS MembEmail"

	sSQL = sSQL + " FROM "&RegDetailTableName&" AS EVT"

	sSQL = sSQL + " JOIN "&RegGenTableName&" AS RGEN ON EVT.MemberID = RGEN.MemberID AND LEFT(EVT.TourID,6) = LEFT(RGEN.TourID,6)" 

	' --- Gets total amount paid by this member in this tournament
	sSQL = sSQL + " LEFT JOIN (SELECT MemberID, SUM(Amount) AS Payments FROM "&RegPaymentTableName
	sSQL = sSQL + " 	WHERE LEFT(TourID,6) = '"&LEFT(sTourID,6)&"' and Result = '0'"
	sSQL = sSQL + " 	GROUP BY MemberID) AS TP"
	sSQL = sSQL + "	ON TP.MemberID = RGEN.MemberID"

	'sSQL = sSQL + " LEFT JOIN "&MemberTableName&" AS MEM ON RGEN.MemberID = MEM.personidwithcheckdigit" 
  sSQL = sSQL + " LEFT JOIN "&MemberLiveTableName&" AS MEM ON CAST(RIGHT(RGEN.MemberID,8) AS INT) = MEM.PersonID" 
	sSQL = sSQL + " LEFT JOIN "&TRegSetupTableName&" AS TGEN ON LEFT(TGEN.TournAppID,6) = LEFT(RGEN.TourID,6)"
	sSQL = sSQL + " LEFT JOIN "&LeagueTableName&" AS LT ON TGEN.TournAppID = LEFT(LT.QualifyTour,6)"



	' -----------------------------------
	' ------ Begin WHERE condition ------
	' -----------------------------------

	sSQL = sSQL + " WHERE LEFT(RGEN.[TourID],6) = '"&LEFT(sTourID,6)&"'"

	IF DivSelected = "ALL" THEN
		'sSQL = sSQL + " AND (EVT.div IN ('MM', 'OM', 'OW', 'B1', 'B2','B3', 'G1', 'G2', 'G3', 'M1', 'M2', 'M3', 'M4', 'M5', 'M6', 'M7', 'M8', 'M9', 'MA', 'MB', 'W1', 'W2', 'W3','W4', 'W5', 'W6','W7', 'W8', 'W9', 'WA', 'WB'))"
	ELSE
		sSQL = sSQL + " AND EVT.div = '"&DivSelected&"'"
	END IF

	IF EventSelected = "ALL" THEN 
			'sSQL = sSQL + " AND (EVT.event IN ('S', 'T', 'J', 'WB', 'WS', 'WU', 'KS', 'KT', 'KF', 'KR'))"
	ELSE
			sSQL = sSQL + " AND (EVT.event = '"&EventSelected&"')"
	END IF

	' --- First and Last Characters when report gets too big
	IF StartCharSelected<>"All" THEN
			sSQL = sSQL + " AND (LEFT(MEM.LastName,1)>='"&StartCharSelected&"')"		
	END IF
	
	IF EndCharSelected<>"All" THEN
			sSQL = sSQL + " AND (LEFT(MEM.LastName,1)<='"&EndCharSelected&"')"		
	END IF

	IF TRIM(sPrintDate)<>"" THEN
			sSQL = sSQL + " AND RGEN.RegisterDate='"&sPrintDate&"'"			
	END IF

	IF RegionSelected <> "6" THEN sSQL = sSQL + " AND REGION.[region] = '"&RegionSelected&"'"
	IF StateSelected <> "All" THEN sSQL = sSQL + " AND MEM.State = '"&StateSelected&"'"


	' ------------------------------------
	' ------ Sets ORDER of Display  ------
	' ------------------------------------

	sSQL = sSQL + " ORDER BY EVT.div, EVT.event, RT.RankScore DESC, MEM.LastName, MEM.FirstName"

	IF Session("adminmenulevel")>=50 AND TRIM(sShowSQL)<>"" THEN
			response.write("<br>"&sSQL)
			response.end
	END IF

	rs.open sSQL, sConnectionToTRATable, 3, 1
	IF rs.eof THEN sWhatReport = "EndofFile"




END SUB








' -------------------------------
  SUB MailNotices
' -------------------------------

MaxECount=6


'IF NOT WhatNotify="PrintBio" THEN WriteIndexPageHeader
' --- Changed 8-9-2013 ---
IF NOT WhatNotify="PrintBio" THEN 
	
		WriteIndexPageHeader
		' --- Bottom of IF-THEN for not displaying the drop down ---

		' ---- Heading Section for all reports ----
		ReportTitle = "Bios and Notifications"

		%>
		<TABLE class="innertable" WIDTH=750px  >
			<tr>
	   		<td align="left" colspan=4 width=60%>
					<font size=3 color="<% =Textcolor2 %>"><b><% =sTourName %></b>&nbsp;&nbsp;&nbsp;&nbsp</font>
					<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><br><% =sTDateS %>-<% =sTDateE %></b></font>
	   		</td>	
	   		<td colspan=2 align="left" width=40%>
		  		<font color="<%=Textcolor2%>" size=3><B><% Response.Write(ReportTitle) %></B></FONT>
					<br><br>
	   		</td>
	 		</tr>

	  	<form action="/rankings/<%=ThisFileName%>" method="post">

	  	<tr><%
			
			' --- Loads report Type 
			LoadReportPulldown

			' --- Loads divisions offered in this event - SUB in view-registration.asp --  
			LoadDivPulldown 

			LoadStartPulldown %>

	  	</tr>
	  	<tr>	
	    	<td align="right">		
					<font size=<% =fontsize3 %> color="<% =Textcolor1 %>">Action:</font>
	    	</td>	
	    	<td align="left">		
    			<select name="WhatNotify">
						<option value=""<%IF WhatNotify = "" THEN Response.Write(" SELECTED ")%>>Not Selected</option>
						<option value="ViewNotices"<%IF WhatNotify = "ViewNotices" THEN Response.Write(" SELECTED ")%>>Preview Document</option>
	        	<option value="ViewList"<%IF WhatNotify = "ViewList" THEN Response.Write(" SELECTED ")%>>Show Target List</option>
	        	<%
						IF AdminMenuLevel>=30 THEN %>
	        		<option value="SendList"<%IF WhatNotify = "SendList" THEN Response.Write(" SELECTED ")%>>Send Emails</option><%
						END IF 	
						%>
	        	<option value="PrintBio"<%IF WhatNotify = "PrintBio" THEN Response.Write(" SELECTED ")%>>Print Bios</option>
        	</select>
	    	</td><%

				LoadEventPulldownNew  	

				LoadEndPulldown %>
	  	</tr>

	  	<tr>
	    	<td align="right">		
					<font size=<% =fontsize3 %> color="<% =Textcolor1 %>">Document:</font>
	    	</td><%

	    	IF AdminMenuLevel>=30 THEN %>
						<td align="left">		
							<select name="WhatLetter">
								<option value=""<%IF WhatLetter = "" THEN Response.Write(" SELECTED ")%>>Not Selected</option>
								<option value="reg_deficiency"<%IF WhatLetter = "reg_deficiency" THEN Response.Write(" SELECTED ")%>>Deficiency</option>
        				<option value="reg_bioincomplete"<%IF WhatLetter = "reg_bioincomplete" THEN Response.Write(" SELECTED ")%>>Bio Incomplete</option>
	      				<option value="reg_confirm"<%IF WhatLetter = "reg_confirm" THEN Response.Write(" SELECTED ")%>>Confirmation</option>
	      				<option value="reg_eventchange-m3-4slalom-2014"<%IF WhatLetter = "reg_eventchange-m3-4slalom-2014" THEN Response.Write(" SELECTED ")%>>M3-M4 Event Change</option>
        				<option value="Bio"<%IF WhatLetter = "Bio" THEN Response.Write(" SELECTED ")%>>Personal Bio</option>
        				<%
								IF adminmenulevel>=50 THEN %>
	      					<option value="custom"<%IF WhatLetter = "custom" THEN Response.Write(" SELECTED ")%>>Custom</option><%
								END IF 
								%>
		    			</select>
						</td><%
	    	ELSE %>
					<input type="hidden" name="WhatLetter" value="Bio">		
					<td align="left">
						<font size=<% =fontsize3 %> color="<% =Textcolor2 %>">Bio Form</font>
					</td><%
	    	END IF %>
	
		
	    	<td>&nbsp;</td>	
	    	<td>&nbsp;</td><%

	    	IF AdminMenuLevel>=30 THEN %>
		    	<td align="right">
						<font size=<% =fontsize3 %> color="<% =Textcolor1 %>">Register Date:</font>
		    	</td>
		    	<td align="left">	
						<input type="text" name="sPrintDate" value= "<% =sPrintDate %>" maxlength="10" size="10" >
						<font size=<% =fontsize3 %> color="<% =Textcolor1 %>">mm/dd/yyyy</font>
		    	</td><%
	    	ELSE  %>
		    	<td>&nbsp;</td>
		    	<td>&nbsp;</td><%
	    	END IF  %>

	  	</tr><% 

	  	IF adminmenulevel >= 19  THEN %>
		  	<tr>
		    	<td align="right">
		        <font size=<% =fontsize3 %> color="<% =Textcolor1 %>">Scratch Filter&nbsp;</FONT>
		    	</td>
		    	<td align="left" colspan=3>	
						<font size=<% =fontsize3 %>>Bio(<%=MaxECount%>)</font>
						<input type=checkbox name="sBioFilter" <% IF sBioFilter = "on" THEN Response.Write("Checked") %>>
						<font size=<% =fontsize3 %>>Qualify</font>
						<input type=checkbox name="sQualFilter" <% IF sQualFilter = "on" THEN Response.Write("Checked") %>>
						<font size=<% =fontsize3 %>>Waiver</font>
						<input type=checkbox name="sWaivFilter" <% IF sWaivFilter = "on" THEN Response.Write("Checked") %>>
						<font size=<% =fontsize3 %>>Fee Due</font>
						<input type=checkbox name="sFeeFilter" <% IF sFeeFilter = "on" THEN Response.Write("Checked") %>>
		    	</td>

		    	<td align="right">
						<font size=<% =fontsize3 %>><b>Resend:</b></font>
		    	</td>
		    	<td align="left">	
						<input type=checkbox name="sResendEmail" <% IF sResendEmail = "on" THEN Response.Write("Checked") %>>
						<font size=<% =fontsize3 %>>If previously printed</font>
		    	</td>
		  	</tr><%
	  	END IF %>


	  	<tr>
	    	<td align=center colspan=6>
	      	<input type="submit" align="center" value="Submit Action">
	    	</td>
	   	</form>
	  	</tr>
		</TABLE>

		<br>

		<TABLE class="innertable" WIDTH=750px >
  		<tr>
   			<td width=10%>&nbsp;</td>	
   			<td align="left" colspan=4 width=60%>
					<br>
					<font size=4><b>NOTICE:</b></font>
					<br>
					<font size=<%=fontsize2%>>Authorized announcers may view Skier Bio online from links on the Registration Status report.</font>
					<br><br>
					<font size=<%=fontsize2%>>Bio printing follows three (3) steps. </font>
					<br>
					<font size=<%=fontsize2%>>&nbsp;&nbsp;&nbsp;1) Select the document design and Preview Document as the Action.  Confirm the bio format is what you desire to print.   </font>
					<br>
					<font size=<%=fontsize2%>>&nbsp;&nbsp;&nbsp;2) Display the target list of recipients. </font>
					<br>
					<font size=<%=fontsize2%>>&nbsp;&nbsp;&nbsp;3) Print the selected document or Send It to recipient list.  </font>
					<br><br>

					<font size=<%=fontsize2%>>Bios received for registrations on specified dates may be printed.  Larger tournaments may need to split the print run to <br>avoid buffer overflow by selecting ranges of names using the <b>Start With</b> and <b>End With</b> selections</font>
					<br><br>
					<font size=<%=fontsize2%>>It is recommended to print all events for a division as this makes the bio more versatile.</font>
					<br><br>
					<font size=<%=fontsize2%>>The scratch notice filter selections (Bio6, Qualify, Waiver and Fee Due) are used in conjunction with Deficiency <br>Notice document.</font>

					<br>
					<br>
	    	</td>
	  	</tr>
		</TABLE>



		<%
END IF   ' --- Bottom of IF-THEN for not displaying the drop down ---


' --- Performes function ---

If NOT rs.eof THEN
	GenerateList
ELSE  %>
	<br>
	<center><font size=<% =fontsize3 %> color=<% =textcolor3 %>><b>No Output For These Settings.</b></font></center><%
END IF

IF NOT WhatNotify="PrintBio" THEN  WriteIndexPageFooter




END SUB




' -----------------
   SUB GenerateList
' -----------------

sDiv(1)=""
sDiv(2)=""
sDiv(3)=""
sDiv(4)=""

rs.movefirst

IF NOT rs.eof THEN 

	IF WhatNotify="ViewList" THEN  %>

		<br>
		<center><font size=5 font=<%=font1%> ><b>TARGET LIST</b></font></center>
		<br>
		<TABLE align="Center" class="innertable" width=100%>
 	    <TR>
	    	<th align="left"><font size=<%=fontsize2%> color="#FFFFFF"><b>&nbsp;Member</b></FONT></th>
	      <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>State</b></FONT></th>
		    <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>&nbsp;<%=sTEventName(1)%></b></FONT></th>
		    <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>&nbsp;<%=sTEventName(2)%></b></FONT></th>
		    <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>&nbsp;<%=sTEventName(3)%></b></FONT></th>
		    <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>&nbsp;<%=sTEventName(4)%></b></FONT></th>
		    <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>OK to<br>Email</b></FONT></th>
		    <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>Email Address</b></FONT></th>
		    <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>BioScore</b></FONT></th>
		    <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>Bio<br>Sent</b></FONT></th>
		  </TR><%	

	END IF

	IF WhatNotify="PrintBio" THEN
	        %><center>
	        	<a href='#' onclick='window.print()' title="Click here to Print">
	        		<input type=submit value="Print Now" style="width:9em">
	        	</a>
		  <form method="post"> 	
	          <input type="submit" value="Return to Menu" title="Return to screen with selection options"></center>
		  </form>
	
	  
			<%

	END IF

	LastMemb=rs("MemberID")

	EMailCount=0


	DO WHILE NOT rs.eof  

		  SELECT CASE TRIM(rs("Event"))
				CASE sTEvent(1) 
						sDiv(1)=rs("Div")
				CASE sTEvent(2)
						sDiv(2)=rs("Div")
				CASE sTEvent(3) 
						sDiv(3)=rs("Div")
				CASE sTEvent(4) 
						sDiv(4)=rs("Div")
		  END SELECT

		  sState=rs("State")
		  sFullName=rs("LastName")&", "&rs("FirstName")
		  sNoEmail=rs("NoEmail")
		  sEmail=rs("Email")
		  'sTourEmail=rs("EmailAddress")
		  sSentBioEmail=TRIM(rs("SentBioEmail"))	


		  rs.movenext
		  IF NOT rs.eof THEN sMemberID=rs("MemberID")




		  ' ---- After collecting each event information then display this members info	---
		  IF rs.eof OR sMemberID<>LastMemb THEN 

			' --- View the list of selected ---
		 	IF WhatNotify="ViewList" THEN  
				
			    ' --- Checks whether bio form is complete ---
			    HowEmptyIsForm LastMemb, sTourID

					'response.write("INSIDE PONT 1")			
					'response.write("sBioFilter= "&sBioFilter)
					'response.write("ECount= "&ECount)
					'response.write("MaxECount= "&MaxECount)
			    ' --- Display info to screen ---
			    IF WhatLetter<>"reg_bioincomplete" OR (WhatLetter="reg_bioincomplete" AND sBioFilter="on" AND ECount >= MaxECount AND (sSentBioEmail="N" OR (TRIM(sSentBioEmail)="Y" AND sResendEmail="on"))) THEN
							'response.write("INSIDE PONT 2a")
				 			%>	 
			      	<TR>
			        	<TD align="left" >
									<font size=<%=fontsize2%>>
										<a title="MemberID: <%=sMemberID%>"><%=sFullName%></a>
									</font>
				      	</TD>
	        		  <TD align="Center" ><font size=<%=fontsize2%>><%=sState%></FONT></TD>
				      	<TD align="Center" ><font size=<%=fontsize2%>>&nbsp;<%=sDiv(1)%></FONT></TD>
				      	<TD align="Center" ><font size=<%=fontsize2%>>&nbsp;<%=sDiv(2)%></FONT></TD>
				      	<TD align="Center" ><font size=<%=fontsize2%>>&nbsp;<%=sDiv(3)%></FONT></TD>
				      	<TD align="Center" ><font size=<%=fontsize2%>>&nbsp;<%=sDiv(4)%></FONT></TD><%

		  	        IF sNoEmail="True" THEN  %>
										<TD align="Center" ><font size=<%=fontsize2%>>NO</FONT></TD><%
			         	ELSE %>	
										<TD>&nbsp;</TD><%
			         END IF	

		  	       IF sEmail<>"" THEN  %>
										<TD align="Center" ><font size=<%=fontsize2%>><%=sEmail%></FONT></TD><%
			         ELSE %>	
										<TD>&nbsp;</TD><%
			         END IF %>	
				      <TD align="Center" ><font size=<%=fontsize2%>>&nbsp;<%=ECount%></FONT></TD>
				      <TD align="Center" ><font size=<%=fontsize2%>>&nbsp;<%=sSentBioEmail%></FONT></TD>
				    </TR><%	

			    ELSE
							'response.write("INSIDE PONT 2b")
			    END IF

			' --- View Samples of Email Letters or item to be displayed/printed in bulk ---
			ELSEIF WhatNotify="ViewNotices" THEN

					'response.write("INSIDE PONT 3")
					'response.write("IN ELSEIF WhatLetter="&WhatLetter)

					SELECT CASE WhatLetter
						CASE "Bio"
								%><br><center><font size=5 font=<%=font1%> ><b>SAMPLE</b></font></center><br><%
								DisplayBioForm LastMemb, sTourID, sDiv(1), sDiv(2), sDiv(3), sDiv(4)
								EXIT DO

						CASE ELSE
								%><br><center><font size=5 font=<%=font1%> ><b>SAMPLE LETTER</b></font></center><br><%
								'markdebug("Inside")
								CreateTheLetter
								response.write(ebody)
								ebody=""
								EXIT DO
					END SELECT

				


			' --- Send the next email
		 	ELSEIF WhatNotify="SendList" THEN  
					'response.write("INSIDE PONT 4")

			    ' --- Checks whether bio form is complete ---
			    HowEmptyIsForm LastMemb, sTourID

					' --- Letter must be the Bio Letter update
					' --- Checkbox for bio must be on
					' --- Analysis of missing field information must be greater than threshold
					' --- Member must not have NO EMAIL shown in Member file
					' --- Bio Letter must not have been previously sent or ResendEmail checkbox must be checked

			    IF WhatLetter="reg_bioincomplete" AND sBioFilter="on" AND ECount >= MaxECount AND sNoEmail<>"True" AND (sSentBioEmail="N" OR (sSentBioEmail="Y" AND sReSendEmail="on")) THEN 	 
							NowSendEmail


							sSQL = "UPDATE "&RegGenTableName
							sSQL = sSQL + " SET SentBioEmail = 'Y'"
							sSQL = sSQL + " WHERE LEFT(TourID,6) = '"&LEFT(sTourID,6)&"' AND MemberID = '"&LastMemb&"'"
							con.execute(sSQL)

					ELSEIF (WhatLetter="reg_eventchange-m3-4slalom-2014" OR WhatLetter="reg_westnile" OR WhatLetter="custom") AND sNoEmail<>"True" THEN

							ebody=""
							NowSendEmail
					END IF



			' --- Write the next bio to the end of the display string
		 	ELSEIF WhatNotify="PrintBio" THEN  
					'response.write("INSIDE PONT 5")
					%><br><%

					DisplayBioForm LastMemb, sTourID, sDiv(1), sDiv(2), sDiv(3), sDiv(4)
			
			
			END IF

			'response.write("INSIDE PONT 6")
			sDiv(1)=""	
			sDiv(2)=""
			sDiv(3)=""	
			sDiv(4)=""	

		END IF

		'response.write("INSIDE PONT 7")

		IF NOT rs.eof THEN LastMemb=rs("MemberID")		  		    

	LOOP

	IF WhatNotify="ViewList" THEN %>
		</TABLE> <%

	ELSEIF WhatNotify="SendList" THEN
		
		%><br><br><center><font size=4 color="<% =Textcolor1 %>"><b><%=EmailCount%> - Emails Sent</b></center></font<%
	END IF

ELSE  %>

	<br>
	<center><font size=4 font=<%=font1%> ><b>No Data Meets Search Criteria</b></font></center>
	<br><%


END IF



END SUB




' ---------------------------
    SUB NowSendEmail
' ---------------------------


Dim eMailSubj, eMailFrom, eMailTo, eMailCC, eMailBCC, eMailBody
Dim SendAddress, HQErrorEmail

	' --- Build the HTML for the messsage ---	
	CreateTheLetter

'response.write("<br>sEmail = "&sEmail)

	ByPassEmails="NO"
	IF TRIM(sEmail)<>"" AND ByPassEmails <> "YES" THEN
			EmailCount=EmailCount+1			
	
			eMailTo = sEmail
			'eMailTo = "mark.crone@bonniercorp.com"

			IF EmailCount=1 THEN
					eMailBCC = marksemailaddress
			END IF
			eMailCC="competition@usawaterski.org"
			eMailFrom = "USA Water Ski<competition@usawaterski.org>"
			eMailSubj = "USA Waterski - Event Notification - "&sFullName

			eMailBody = ebody	

 			' --- SEND the message, and then clear that object
			SetupEmailService

			objMessage.Subject = eMailSubj
			objMessage.From = eMailFrom
			objMessage.To = eMailTo
			objMessage.cc = eMailCC
			objMessage.bcc = eMailBCC
			objMessage.HTMLBody = eMailBody
 

			IF TRIM(eMailTo)<>"" THEN
					objMessage.Send
			END IF
			set objMessage = Nothing


	END IF



END SUB



' ---------------------
   SUB CreateTheLetter
' ---------------------

	ebody = "<HTML><HEAD>"

	ebody = ebody & "<style>div.break {page-break-before:always}</style>" 

	ebody = ebody & "</HEAD><BODY>"

 	ebody = ebody & "<TABLE BORDER=4 align=CENTER CELLPADDING=5 CELLSPACING=0 BGcolor="&Tablecolor1&" width=75% >"

	' Reads and displays text from communications folder
	Set objfso = CreateObject("Scripting.FileSystemObject")
	IF objFSO.FileExists(PathToCommune & "\"&WhatLetter&".txt") THEN
			set objstream=objFSO.opentextfile(PathToCommune & "\"&WhatLetter&".txt")
		   	IF NOT objstream.atendofstream THEN
				DO WHILE not objstream.atendofstream
					currentline=" "&objstream.readline
					LenCurLine=Len(currentline)

					IF InStr(currentline, "+HEADLINE+") > 0 THEN
						headcolor=LEFT(RIGHT(currentline,LenCurLine-11),3)
						headcolor="red"
						ebody = ebody & "<TR>"
						IF LEFT(RIGHT(currentline,LenCurLine-11),3) ="RED" THEN
							ebody = ebody & "<TD BGcolor=red ><center><font face="&font1&" color=#FFFFFF size=4><b>"&RIGHT(currentline,LenCurLine-15)&"</b></font></TD>"
						ELSE
							ebody = ebody & "<TD BGcolor=blue ><center><font face="&font1&" color=#FFFFFF size=4><b>"&RIGHT(currentline,LenCurLine-15)&"</b></font></TD>"					
						END IF
	
						ebody = ebody & "</TR>"
						ebody = ebody & "<TR>"
						ebody = ebody & "<TD align=center Valign=top>"

					ELSEIF InStr(currentline, "+EVENT+") > 0 THEN

						ebody = ebody & "<br><br>"
						ebody = ebody & "<font face="&font1&" size=2><b>Events Entered</b></font>"
						ebody = ebody & "<br>"

						IF TRIM(sDiv(1)) <> "" THEN
							ebody = ebody & "<font color="&Textcolor2&" face="&font1&" size=2>"&sDiv(1)&" - "&sTEventName(1)&"</font>"
							ebody = ebody & "<br>"
						END IF
						IF TRIM(sDiv(2)) <> "" THEN
							ebody = ebody & "<font color="&Textcolor2&" face="&font1&" size=2>"&sDiv(2)&" - "&sTEventName(2)&"</font>"
							ebody = ebody & "<br>"
						END IF
						IF TRIM(sDiv(3)) <> "" THEN
							ebody = ebody & "<font color="&Textcolor2&" face="&font1&" size=2>"&sDiv(3)&" - "&sTEventName(3)&"</font>"
							ebody = ebody & "<br>"
						END IF
						IF TRIM(sDiv(4)) <> "" THEN
							ebody = ebody & "<font color="&Textcolor2&" face="&font1&" size=2>"&sDiv(4)&" - "&sTEventName(4)&"</font>"
							ebody = ebody & "<br>"
						END IF
		
					ELSEIF InStr(currentline, "+MEMBER+") > 0 THEN
						ebody = ebody & "<br>"
						ebody = ebody & "<font color="&Textcolor2&" face="&font1&" size=4><b>"&sFullName&"</b></font>"
						ebody = ebody & "<br>"
						ebody = ebody & "<font face="&font1&" size=2><b>MemberID = </b></font><font color="&Textcolor2&" face="&font1&" size=2>"&LastMemb&"</font>"

					ELSEIF InStr(currentline, "+TOURNAMENT+") > 0 THEN
						ebody = ebody & "<br>"
						ebody = ebody & "<font color=red face="&font1&" size=4><b>"&sTourName&"</b></font>"
						ebody = ebody & "<br>"
						ebody = ebody & "<font face="&font1&" size=2><b>SanctionID = </b></font><font color="&Textcolor2&" face="&font1&" size=2>"&sTourID&"</font>"
						ebody = ebody & "<br>"
						ebody = ebody & "<font face="&font1&" size=2><b>Date = </b></font><font color="&Textcolor2&" face="&font1&" size=2>"&sTDateS&" to "&sTDateE&"</font></b>"
						ebody = ebody & "<br><br>"


					ELSE
						ebody = ebody & "<br><font color="&Textcolor2&" face="&font1&" size=2>"&currentline&"</font>"
					END IF
				LOOP
			   END IF
			objstream.close
	END IF


	ebody = ebody & "<br>"
	ebody = ebody & "</td></tr>"

	ebody = ebody & "</TABLE>"

	IF NOT rs.eof THEN
		ebody = ebody & "<div class=break />"
	END IF

	ebody = ebody & "</BODY></HTML>"


	

END SUB






' -------------------------------
  SUB SchedulingByDivision
' -------------------------------

ReportTitle = "Scheduling Time By Division"
MembersOrPulls = "members"
IF TRIM(Request("MembersOrPulls"))<>"" THEN MembersOrPulls = Request("MembersOrPulls")

'response.write("<br>sTPandC = "&sTPandC)


%>
<TABLE class="droptable" height="120px" width="<%=TourTableWidth%>px" align=center>
	  <tr>
	   <td align="left" colspan=4 width=60%>
		<font size=3 color="<% =Textcolor2 %>"><b><% =sTourName %></b>&nbsp;&nbsp;&nbsp;&nbsp</font>
		<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><br><% =sTDateS %>-<% =sTDateE %></b></font>
	   </td>	
	   <td colspan=2 align="left" width=40%>
		  <FONT color="<%=Textcolor2%>" size=3><B><% Response.Write(ReportTitle) %></B></FONT>
		<br><br>
	   </td>
	 </tr>

  <form action="/rankings/<%=ThisFileName%>" method="post">

  	<tr>
    	<td colspan=4>
    		<%
    		IF sTPandC=true THEN
    				%>
    				<input type="radio" name="MembersOrPulls" value="members" <% IF MembersOrPulls="members" THEN response.write("checked") %> onchange="submit();">Members
    				<input type="radio" name="MembersOrPulls" value="pulls" <% IF MembersOrPulls="pulls" THEN response.write("checked") %> onchange="submit();">Pulls
						<%
				ELSE
						%>&nbsp;<%
				END IF
				%>			
    	</td>	
    	<td align=center colspan=2>
    		<%
	    	 IF PrintButton="Printer Friendly" THEN 
	    	 		%><a href='#' onclick='window.print()' title="Click here to Print"><input type=submit value="Print Now" style="width:9em"></a><%
	      ELSE 
	      		%><input type="submit" align="center" style="width:10em;" name="PrintButton" value="Printer Friendly"><%
	      END IF 
	      %>
   		</td>
  	</tr>
		<tr>
			<%
			LoadReportPulldown 
			%>
    	<td colspan=3>&nbsp</td>	
    	<td align=center COLSPAN=1>
    		<%
				IF PrintButton="Printer Friendly" THEN 
						%><input type="submit" align="center" style="width:10em;"  value="Report Update"><%
				ELSE 
						%><input type="submit" align="center" style="width:10em;"  value="Display Report"><%
				END IF 
				%>	
    	</td>
  	</tr>


</TABLE> <%




' --- This produces totals by division
SET rsTot=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT  COALESCE(EVT1.DIV, EVT2.DIV, EVT3.DIV) AS DIV"
sSQL = sSQL + " , COALESCE(EVT1.EVT1_CNT,0) AS EVT1_CNT, COALESCE(EVT2.EVT2_CNT,0) AS EVT2_CNT, COALESCE(EVT3.EVT3_CNT,0) AS EVT3_CNT"
sSQL = sSQL + " , COALESCE(EVT1.EVT1_SUM,0) AS EVT1_SUM, COALESCE(EVT2.EVT2_SUM,0) AS EVT2_SUM, COALESCE(EVT3.EVT3_SUM,0) AS EVT3_SUM"

sSQL = sSQL + " FROM "

sSQL = sSQL + " (SELECT DIV, COUNT(MemberID) AS EVT1_CNT, SUM(FeeRounds) AS EVT1_SUM FROM "&RegDetailTableName&" AS EVT WHERE Event='"&sTEvent(1)&"' AND LEFT(TourID,6)='"&LEFT(sTourID,6)&"' GROUP BY Div) AS EVT1"
sSQL = sSQL + " FULL OUTER JOIN "
sSQL = sSQL + " (SELECT div, COUNT(MemberID) AS EVT2_CNT, SUM(FeeRounds) AS EVT2_SUM FROM "&RegDetailTableName&" AS EVT WHERE Event='"&sTEvent(2)&"' AND LEFT(TourID,6)='"&LEFT(sTourID,6)&"' GROUP BY Div) AS EVT2 ON EVT2.DIV=EVT1.DIV"
sSQL = sSQL + " FULL OUTER JOIN "
sSQL = sSQL + " (SELECT div, COUNT(MemberID) AS EVT3_CNT, SUM(FeeRounds) AS EVT3_SUM FROM "&RegDetailTableName&" AS EVT WHERE Event='"&sTEvent(3)&"' AND LEFT(TourID,6)='"&LEFT(sTourID,6)&"' GROUP BY Div) AS EVT3 ON EVT3.DIV=EVT1.DIV"

sSQL = sSQL + " UNION "
sSQL = sSQL + " SELECT 'z TOTAL ALL', EVT1.EVT1_CNT, EVT2.EVT2_CNT, COALESCE(EVT3.EVT3_CNT,0)"
sSQL = sSQL + " , EVT1.EVT1_SUM, EVT2.EVT2_SUM, EVT3.EVT3_SUM"
sSQL = sSQL + "  FROM " 
sSQL = sSQL + " (SELECT COUNT(MemberID) AS EVT1_CNT, SUM(FeeRounds) AS EVT1_SUM FROM "&RegDetailTableName&" AS EVT WHERE Event='"&sTEvent(1)&"' AND LEFT(TourID,6)='"&LEFT(sTourID,6)&"') AS EVT1" 
sSQL = sSQL + ", (SELECT COUNT(MemberID) AS EVT2_CNT, SUM(FeeRounds) AS EVT2_SUM FROM "&RegDetailTableName&" AS EVT WHERE Event='"&sTEvent(2)&"' AND LEFT(TourID,6)='"&LEFT(sTourID,6)&"') AS EVT2" 
sSQL = sSQL + ", (SELECT COUNT(MemberID) AS EVT3_CNT, SUM(FeeRounds) AS EVT3_SUM FROM "&RegDetailTableName&" AS EVT WHERE Event='"&sTEvent(3)&"' AND LEFT(TourID,6)='"&LEFT(sTourID,6)&"') AS EVT3" 

sSQL = sSQL + " ORDER BY Div" 

'response.write(sSQL)
'response.end


rsTot.open sSQL, sConnectionToTRATable, 3, 1


EVT1_TIME=Request("EVT1_TIME")
EVT2_TIME=Request("EVT2_TIME")
EVT3_TIME=Request("EVT3_TIME")


IF TRIM(EVT1_TIME) = "" THEN EVT1_TIME = 360
IF TRIM(EVT2_TIME) = "" THEN EVT2_TIME = 300
IF TRIM(EVT3_TIME) = "" THEN EVT3_TIME = 420


IF MembersOrPulls="members" THEN
		ThisMetricHeading = "# of Skiers"
		ThisTimeHeading = "Time (min:sec) Per Skier"
ELSE
		ThisMetricHeading = "# of Pulls"	
		ThisTimeHeading = "Time (min:sec) Per Pull"
END IF

IF NOT rsTot.eof THEN

	%>
	<br>
	<TABLE align="Center" class="innertable" width="<%=TourTableWidth%>px">

		<tr>
		  <th bgcolor="<%=Headcolor1%>">&nbsp;</th>
		  <th align="Center" colspan=3 bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF"><b><% =ThisMetricHeading %></b></th>	
		  <th align="Center" colspan=3 bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF"><b><% =ThisTimeHeading %></b></th>	
		</tr>

		<tr>
		  <th align="Center"  bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF"><b>Division</b></th>	
		  <th align="Center" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF"><b><%=sTEventName(1)%></b></th>	
		  <th align="Center" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF"><b><%=sTEventName(2)%></b></th>	
		  <th align="Center" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF"><b><%=sTEventName(3)%></b></th>
		  <th align="Center" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF"><b><%=sTEventName(1)%></b></th>	
		  <th align="Center" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF"><b><%=sTEventName(2)%></b></th>	
		  <th align="Center" bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="#FFFFFF"><b><%=sTEventName(3)%></b></th>

		</tr>

		<tr>    	
		  <td align=center bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="<%=Textcolor1%>"><b>Set Time/Skier</b></font></td>
		<td bgcolor="<%=Headcolor1%>">&nbsp;</td>
		<td bgcolor="<%=Headcolor1%>">&nbsp;</td>
		<td bgcolor="<%=Headcolor1%>">&nbsp;</td>
		    <td align ="center" bgcolor="<%=Headcolor1%>">
			 <select name="EVT1_TIME" ><%
			   LoadTimePullDown EVT1_TIME, 240, 600, 10
			  %>
			</select>
		    </td>
		    <td align ="center" bgcolor="<%=Headcolor1%>">
			 <select name="EVT2_TIME" ><%
			   LoadTimePullDown EVT2_TIME, 240, 600, 10
			  %>
			</select>
		    </td>
		    <td align ="center" bgcolor="<%=Headcolor1%>">
			 <select name="EVT3_TIME" ><%
			   LoadTimePullDown EVT3_TIME, 240, 600, 10
			  %>
			</select>
		    </td>
		</tr>

		<%

	  rsTot.movefirst

	  DO WHILE Not rsTot.EOF 

				IF MembersOrPulls="members" THEN
						ThisMetric_EVT1 = rsTot("EVT1_CNT")
						ThisMetric_EVT2 = rsTot("EVT2_CNT")
						ThisMetric_EVT3 = rsTot("EVT3_CNT")
				ELSE
						ThisMetric_EVT1 = rsTot("EVT1_SUM")
						ThisMetric_EVT2 = rsTot("EVT2_SUM")
						ThisMetric_EVT3 = rsTot("EVT3_SUM")
				END IF




		IF NOT LEFT(formatNumber((((ThisMetric_EVT1 * EVT1_TIME/60) - fix(ThisMetric_EVT1 * EVT1_TIME/60))*60),0),2) > 0 THEN
				DivTimeEvt1 = fix(ThisMetric_EVT1 * EVT1_TIME/60)&":00"
		ELSE
				DivTimeEvt1 = fix(ThisMetric_EVT1 * EVT1_TIME/60)&":"&formatNumber((((ThisMetric_EVT1 * EVT1_TIME/60) - fix(ThisMetric_EVT1 * EVT1_TIME/60))*60),0)
		END IF



		IF NOT LEFT(formatNumber((((ThisMetric_EVT2 * EVT2_TIME/60) - fix(ThisMetric_EVT2 * EVT2_TIME/60))*60),0),2) > 0 THEN
			DivTimeEvt2 = fix(ThisMetric_EVT2 * EVT2_TIME/60)&":00"
		ELSE
			DivTimeEvt2 = fix(ThisMetric_EVT2 * EVT2_TIME/60)&":"&formatNumber((((ThisMetric_EVT2 * EVT2_TIME/60) - fix(ThisMetric_EVT2 * EVT2_TIME/60))*60),0)
		END IF

		IF NOT LEFT(formatNumber((((ThisMetric_EVT3 * EVT3_TIME/60) - fix(ThisMetric_EVT3 * EVT3_TIME/60))*60),0),2) > 0 THEN
			DivTimeEvt3 = fix(ThisMetric_EVT3 * EVT3_TIME/60)&":00"
		ELSE
			DivTimeEvt3 = fix(ThisMetric_EVT3 * EVT3_TIME/60)&":"&formatNumber((((ThisMetric_EVT3 * EVT3_TIME/60) - fix(ThisMetric_EVT3 * EVT3_TIME/60))*60),0)
		END IF




		%>
		<tr>
		  <TD align="Center" style="background-color:<%=Tablecolor1%>;"><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =rsTot("Div") %></TD>	
		  <TD align="Center" style="background-color:<%=Tablecolor1%>;"><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =ThisMetric_EVT1 %></TD>	
		  <TD align="Center" style="background-color:<%=Tablecolor1%>;"><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =ThisMetric_EVT2 %></TD>	
		  <TD align="Center" style="background-color:<%=Tablecolor1%>;"><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =ThisMetric_EVT3 %></TD>	

		  <TD align="Center" style="background-color:<%=Tablecolor1%>;"><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =DivTimeEvt1 %></TD>	
		  <TD align="Center" style="background-color:<%=Tablecolor1%>;"><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =DivTimeEvt2 %></TD>	
		  <TD align="Center" style="background-color:<%=Tablecolor1%>;"><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =DivTimeEvt3 %></TD>	

		</tr><%

		rsTot.movenext

	  LOOP 

	  rsTot.close
	  Set rsTot=nothing %>

	</TABLE>
	
  </form>
<% 


  ' -------------------------------------------------------
	' ---- This section displays the summaries by Region ----
  ' -------------------------------------------------------
	 %>	
		<br>
		<TABLE align="center" class="innertable" width=50%>
		<tr>
		  <th colspan=2 align=center><font size=<%=fontsize2%> color="#FFFFFF"><b>Number of Skiers By Region</b></font></th>
		</tr>

		<tr>
		  <td bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<b>Region</b></font></td>
		  <td align=center  bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<b>Skiers</b></font></td>
		</tr>


		<%
		SET rsTot=Server.CreateObject("ADODB.recordset")
		sSQL = "SELECT COUNT(MemberID) AS SC_CNT FROM "&RegGenTableName&" AS RGEN"
    sSQL = sSQL + " LEFT JOIN "&MemberLiveTableName&" AS MEM ON CAST(RIGHT(RGEN.MemberID,8) AS INT) = MEM.PersonID AND MEM.FederationCode = 'USA'" 
		sSQL = sSQL + " LEFT JOIN "&RegionTableName&" AS REGN ON LOWER(MEM.[state]) = LOWER(REGN.[state])"
		sSQL = sSQL + " WHERE REGN.[region] = '1' AND LEFT(TourID,6)='"&LEFT(sTourID,6)&"'" 
		rsTot.open sSQL, sConnectionToTRATable, 3, 1  %>
		<tr>
		  <TD align="Left" style="background-color:<%=Tablecolor1%>;"><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;S. Central</font></TD>	
		  <TD align="Center" style="background-color:<%=Tablecolor1%>;"><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =rsTot("SC_CNT") %></font></TD>	
		</tr><%

		SET rsTot=Server.CreateObject("ADODB.recordset")
		sSQL = "SELECT COUNT(MemberID) AS MW_CNT FROM "&RegGenTableName&" AS RGEN"
		sSQL = sSQL + " LEFT JOIN "&MemberLiveTableName&" AS MEM ON CAST(RIGHT(RGEN.MemberID,8) AS INT) = MEM.PersonID AND MEM.FederationCode = 'USA'" 

		sSQL = sSQL + " LEFT JOIN "&RegionTableName&" AS REGN ON LOWER(MEM.[state]) = LOWER(REGN.[state])"
		sSQL = sSQL + " WHERE REGN.[region] = '2' AND LEFT(TourID,6)='"&LEFT(sTourID,6)&"'" 
		rsTot.open sSQL, sConnectionToTRATable, 3, 1  %>

		<tr>
		  <TD align="Left" style="background-color:<%=Tablecolor1%>;"><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;Midwest</font></TD>	
		  <TD align="Center" style="background-color:<%=Tablecolor1%>;"><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =rsTot("MW_CNT") %></font></TD>	
		</tr><%

		SET rsTot=Server.CreateObject("ADODB.recordset")
		sSQL = "SELECT COUNT(MemberID) AS WE_CNT FROM "&RegGenTableName&" AS RGEN"
		sSQL = sSQL + " LEFT JOIN "&MemberLiveTableName&" AS MEM ON CAST(RIGHT(RGEN.MemberID,8) AS INT) = MEM.PersonID AND MEM.FederationCode = 'USA'" 
		sSQL = sSQL + " LEFT JOIN "&RegionTableName&" AS REGN ON LOWER(MEM.[state]) = LOWER(REGN.[state])"
		sSQL = sSQL + " WHERE REGN.[region] = '3' AND LEFT(TourID,6)='"&LEFT(sTourID,6)&"'" 
		rsTot.open sSQL, sConnectionToTRATable, 3, 1  %>

		<tr>
		  <TD align="Left" style="background-color:<%=Tablecolor1%>;"><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;West</font></TD>	
		  <TD align="Center" style="background-color:<%=Tablecolor1%>;"><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =rsTot("WE_CNT") %></font></TD>	
		</tr><%

		SET rsTot=Server.CreateObject("ADODB.recordset")
		sSQL = "SELECT COUNT(MemberID) AS SO_CNT FROM "&RegGenTableName&" AS RGEN"
		sSQL = sSQL + " LEFT JOIN "&MemberLiveTableName&" AS MEM ON CAST(RIGHT(RGEN.MemberID,8) AS INT) = MEM.PersonID AND MEM.FederationCode = 'USA'" 
		sSQL = sSQL + " LEFT JOIN "&RegionTableName&" AS REGN ON LOWER(MEM.[state]) = LOWER(REGN.[state])"
		sSQL = sSQL + " WHERE REGN.[region] = '4' AND LEFT(TourID,6)='"&LEFT(sTourID,6)&"'" 
		rsTot.open sSQL, sConnectionToTRATable, 3, 1  
		%>
		<tr>
		  <TD align="Left" style="background-color:<%=Tablecolor1%>;"><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;South</font></TD>	
		  <TD align="Center" style="background-color:<%=Tablecolor1%>;"><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =rsTot("SO_CNT") %></font></TD>	
		</tr>
		<%

		SET rsTot=Server.CreateObject("ADODB.recordset")
		sSQL = "SELECT COUNT(MemberID) AS EA_CNT FROM "&RegGenTableName&" AS RGEN"
		sSQL = sSQL + " LEFT JOIN "&MemberLiveTableName&" AS MEM ON CAST(RIGHT(RGEN.MemberID,8) AS INT) = MEM.PersonID AND MEM.FederationCode = 'USA'" 
		sSQL = sSQL + " LEFT JOIN "&RegionTableName&" AS REGN ON LOWER(MEM.[state]) = LOWER(REGN.[state])"
		sSQL = sSQL + " WHERE REGN.[region] = '5' AND LEFT(TourID,6)='"&LEFT(sTourID,6)&"'" 
		rsTot.open sSQL, sConnectionToTRATable, 3, 1  
		%>
		<tr>
		  <TD align="Left" style="background-color:<%=Tablecolor1%>;"><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;East</font></TD>	
		  <TD align="Center" style="background-color:<%=Tablecolor1%>;"><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =rsTot("EA_CNT") %></font></TD>	
		</tr>
		<%

	SET rsTot=Server.CreateObject("ADODB.recordset")
	sSQL = "SELECT COUNT(MemberID) AS OT3_CNT" 
	sSQL = sSQL + " FROM "&RegGenTableName&" AS RGEN" 
	sSQL = sSQL + " LEFT JOIN "
	'sSQL = sSQL + "  (SELECT FederationCode,  FirstName, LastName, personidwithcheckdigit, State"
	'sSQL = sSQL + "   	 FROM "&MemberTableName&") AS MEM"
	'sSQL = sSQL + " ON RGEN.MemberID = personidwithcheckdigit"
	sSQL = sSQL + "  (SELECT FederationCode,  FirstName, LastName, personid, State"
	sSQL = sSQL + "      FROM "&MemberLiveTableName&") AS MEM"
	sSQL = sSQL + " ON CAST(RIGHT(RGEN.MemberID,8) AS INT) = MEM.PersonID " 
	sSQL = sSQL + " WHERE LEFT(TourID,6)='"&LEFT(sTourID,6)&"' AND FederationCode<>'USA'"
	
	rsTot.open sSQL, sConnectionToTRATable, 3, 1  

  ' response.write(sSQL)

		%>
		<tr>
		  <TD align="Left" style="background-color:<%=Tablecolor1%>;"><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;Not HQ Fed: USA </font></TD>	
		  <TD align="Center" style="background-color:<%=Tablecolor1%>;"><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<% =rsTot("OT3_CNT") %></font></TD>	
		</tr>
		<%

'a=35
'IF a=34 THEN


		SET rsTot=Server.CreateObject("ADODB.recordset")
		sSQL = "SELECT COUNT(MemberID) AS TT_CNT FROM "&RegGenTableName&" AS RGEN WHERE LEFT(TourID,6)='"&LEFT(sTourID,6)&"'"
		rsTot.open sSQL, sConnectionToTRATable, 3, 1  
		%>
		<tr>
		  <TD align="Left" style="background-color:<%=Tablecolor1%>;"><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<b>TOTAL ALL</b></font></TD>	
		  <TD align="Center" style="background-color:<%=Tablecolor1%>;"><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<b><%=rsTot("TT_CNT")%></b></font></TD>	
		</tr>
	</TABLE>
<%


  ' --------------------------------------------------------
	' ---- This section displays the Out of Region Skiers ----
  ' --------------------------------------------------------

	IF TRIM(Session("AdminMenuLevel"))<>"" THEN
			%>
			<br>
			<TABLE align="center" class="innertable" width=50%>
			<tr>
	  		<th colspan=4 align=center><font size=<%=fontsize2%> color="#FFFFFF"><b>Out of Federation Skiers</b></font></th>
			</tr>
			<tr>
	  		<td colspan=1 align=center><font size=<%=fontsize2%> color="<%=TextColor1%>"><b>Federation</b></font></th>
	  		<td colspan=1 align=center><font size=<%=fontsize2%> color="<%=TextColor1%>"><b>First</b></font></th>
	  		<td colspan=1 align=center><font size=<%=fontsize2%> color="<%=TextColor1%>"><b>Last</b></font></th>
	  		<td colspan=1 align=center><font size=<%=fontsize2%> color="<%=TextColor1%>"><b>MemberID</b></font></th>
			</tr>
			<%
			SET rsTot=Server.CreateObject("ADODB.recordset")
			sSQL = "SELECT COUNT(MemberID) AS OT3_CNT, MEM.FederationCode, State, Mem.FirstName, Mem.LastName, MemberID" 
			sSQL = sSQL + " FROM "&RegGenTableName&" AS RGEN" 
			sSQL = sSQL + " LEFT JOIN "
			'sSQL = sSQL + "  (SELECT FederationCode,  FirstName, LastName, personidwithcheckdigit, State"
			'sSQL = sSQL + "   	 FROM "&MemberTableName&") AS MEM"
			'sSQL = sSQL + " ON RGEN.MemberID = personidwithcheckdigit"

			sSQL = sSQL + "  (SELECT FederationCode,  FirstName, LastName, PersonID, State"
			sSQL = sSQL + "      FROM "&MemberLiveTableName&") AS MEM"
			sSQL = sSQL + " ON CAST(RIGHT(RGEN.MemberID,8) AS INT) = MEM.PersonID " 

			sSQL = sSQL + " WHERE LEFT(TourID,6)='"&sTourID&"' AND FederationCode<>'USA'"
			sSQL = sSQL + " GROUP BY  MEM.FederationCode,  Mem.State, Mem.FirstName, Mem.LastName, MemberID"
			rsTot.open sSQL, sConnectionToTRATable, 3, 1  

			DO WHILE NOT(rsTot.EOF)
					%>
					<tr>
		  			<td bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<%=rsTot("FederationCode")%></font></td>
		  			<td align=center  bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<%=rsTot("FirstName")%></font></td>
		  			<td align=center  bgcolor="<%=Headcolor1%>"><font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;<%=rsTot("LastName")%></font></td>
		  			<td align=center  bgcolor="<%=Headcolor1%>">
		  				<font size=<%=fontsize2%> color="<% =Textcolor1 %>">&nbsp;
								<a href = "/rankings/registration16.asp?sMemberID=<%=rsTot("MemberID")%>&sTourID=<%=sTourID%>" title="Link to open Registration record for <% =rsTot("FirstName")&" "&rsTot("LastName") %>" target="_blank"><%=rsTot("MemberID")%></a>
							</font>			
						</td>
					</tr>
					<%
					rsTot.movenext
			LOOP
			%>
			</TABLE>
			<%
	END IF


ELSE
	response.write("No Data Found in Table")
END IF


END SUB










' -------------------------------
  SUB PaymentReport
' -------------------------------



' --- Displays Transactions from Payments Table CCLog ---

WhatDates = trim(Request("WhatDates"))
WhatPayments = trim(Request("WhatPayments"))
SequenceSelected=trim(Request("SequenceSelected"))
IF SequenceSelected = "" OR SequenceSelected = "seed" THEN SequenceSelected = "chgdate"

IF trim(WhatDates) = "" THEN WhatDates = "Yesterday"
IF trim(WhatPayments) = "" THEN WhatPayments = "ALL"
ReportTitle = "Payments Received"

%>

<form action="/rankings/<%=ThisFileName%>" method="post">

<TABLE class="droptable" align="center" width=100%>
	  <tr>
	   <td align="left" colspan=4 width=60%>
		<font size=3 color="<% =Textcolor2 %>"><a title="TourID: <%=sTourID%>"><b><% =sTourName %></b></a>&nbsp;&nbsp;&nbsp;&nbsp</font>
		<font size=<%=fontsize2%> color="<% =Textcolor1 %>"><br><% =sTDateS %>-<% =sTDateE %></b></font>
	   </td>	
	   <td colspan=2 align="left" width=40%>
		  <FONT color="<%=Textcolor2%>" size=3><B><% Response.Write(ReportTitle) %></B></FONT>
		<br><br>
	   </td>
	 </tr>




  	<tr>
  		<%
			LoadReportPulldown  
			%>
    <td align="right">		
			<font size=<% =fontsize3 %> color="<% =Textcolor1 %>">Paid By:</font>
    </td>
    <td align="left">		
    	<select name="WhatPayments">
				<option value="ALL"<%IF WhatPayments = "ALL" THEN Response.Write(" SELECTED ")%>>All</option>
        <option value="PayPal"<%IF WhatPayments = "PayPal" THEN Response.Write(" SELECTED ")%>>Pay Pal</option>
				<option value="Card"<%IF WhatPayments = "Card" THEN Response.Write(" SELECTED ")%>>Card</option>
        <option value="Check"<%IF WhatPayments = "Check" THEN Response.Write(" SELECTED ")%>>Check</option>
        <option value="Cash"<%IF WhatPayments = "Cash" THEN Response.Write(" SELECTED ")%>>Cash</option>
        <option value="Refund"<%IF WhatPayments = "Refund" THEN Response.Write(" SELECTED ")%>>Refunds</option>
			</select>
    </td>	
    <td>&nbsp;</td>
  </tr>

  <tr>
    <td align="right">		
	<font size=<% =fontsize3 %> color="<% =Textcolor1 %>">Range:</font>
    </td>
    <td align="left">		
    	<select name="WhatDates">
				<option value="Today"<%IF WhatDates = "Today" THEN Response.Write(" SELECTED ")%>>Today</option>
        <option value="Yesterday"<%IF WhatDates = "Yesterday" THEN Response.Write(" SELECTED ")%>>Yesterday</option>
        <option value="Ago2"<%IF WhatDates = "Ago2" THEN Response.Write(" SELECTED ")%>>2 Days Ago</option>
        <option value="Ago3"<%IF WhatDates = "Ago3" THEN Response.Write(" SELECTED ")%>>3 Days Ago</option>
        <option value="Ago4"<%IF WhatDates = "Ago4" THEN Response.Write(" SELECTED ")%>>4 Days Ago</option>
        <option value="Last7"<%IF WhatDates = "Last7" THEN Response.Write(" SELECTED ")%>>Last 7 Days</option>
        <option value="Last30"<%IF WhatDates = "Last30" THEN Response.Write(" SELECTED ")%>>Last 30 Days</option>
        <option value="Last90"<%IF WhatDates = "Last90" THEN Response.Write(" SELECTED ")%>>Last 90 Days</option>
        <option value="ALL"<%IF WhatDates = "ALL" THEN Response.Write(" SELECTED ")%>>All Dates</option>
        </select>
    </td>	
    <td align="right">
			<font size=<% =fontsize3 %> color="<% =Textcolor1 %>">Order By:</font>
    </td>
    <td align="left"">		
			<select name="SequenceSelected">
				<option value="alpha"<%IF SequenceSelected = "alpha" THEN Response.Write(" SELECTED ")%>>Alphabetic</option>
				<option value="chgdate"<%IF SequenceSelected = "chgdate" THEN Response.Write(" SELECTED ")%>>Charge Date</option>
			</select>
    </td>
    <td><a href="mailto:<%=Marksemail%>?subject=Payments Received Report for TourID: <%=sTourID%>" title="Click here to Email problems or recommendations">Report Errors or Feedback</a></td>
  </tr>

  <tr>
    <td COLSPAN=3 align="center">
      <br>	
      <input type="submit" style="width:9em" name="ReturnButton" value="Reset Report">
    </td>

    <td COLSPAN=3 align="center">
      <br>	
      <input type="submit" style="width:9em" name="ReturnButton" value="Main Menu">
    </td>
  </tr>
</TABLE> 

</form>
<%


'WhatDates="Last7"

SELECT CASE WhatDates
	CASE "Today"
		FirstDate = CDate(DateAdd("d", 0, DATE))
		LastDate = CDate(DateAdd("d", 1, DATE))
	CASE "Yesterday"
		FirstDate = CDate(DateAdd("d", -1, DATE))
		LastDate = CDate(DateAdd("d", 0, DATE))
	CASE "Ago2"
		FirstDate = CDate(DateAdd("d", -2, DATE))
		LastDate = CDate(DateAdd("d", -1, DATE))
	CASE "Ago3"
		FirstDate = CDate(DateAdd("d", -3, DATE))
		LastDate = CDate(DateAdd("d", -2, DATE))
	CASE "Ago4"
		FirstDate = CDate(DateAdd("d", -4, DATE))
		LastDate = CDate(DateAdd("d", -3, DATE))
	CASE "Last7"
		FirstDate = CDate(DateAdd("d", -7, DATE))
		LastDate = CDate(DateAdd("d", 1, DATE))
	CASE "Last30"
		FirstDate = CDate(DateAdd("d", -30, DATE))
		LastDate = CDate(DateAdd("d", 1, DATE))
	CASE "Last90"
		FirstDate = CDate(DateAdd("d", -360, DATE))
		LastDate = CDate(DateAdd("d", 1, DATE))
	CASE ELSE
		FirstDate = CDate(DateAdd("d", -3600, DATE))
		LastDate = CDate(DateAdd("d", 1, DATE))
END SELECT 

%>
<br>
<TABLE align="Center" class="innertable" width="<%=TourTableWidth%>">
<%


' ----  Creates query for all transactions with matching date/time  ----

SET rsRT=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT RTN.MemberID, MT.FirstName, MT.LastName, RPL.TransDate, RPL.PayType, RTN.OrderNo, RT1.TransCode, Coalesce(EntryFee,0) AS EntryFee"
sSQL = sSQL + "	, RT2.TransCode, Coalesce(LateFee,0) AS LateFee, RT3.TransCode, Coalesce(AWSEF,0) AS AWSEF"
sSQL = sSQL + "	, RT4.TransCode, Coalesce(OffDisc,0) AS OffDisc, RT5.TransCode, Coalesce(JrDisc,0) AS JrDisc, RT6.TransCode, Coalesce(SrDisc,0) AS SrDisc"
sSQL = sSQL + "	, RT7.TransCode, Coalesce(ClubDisc,0) AS ClubDisc"

sSQL = sSQL + "	, RT8.TransCode, Coalesce(Banquet,0) AS Banquet"
sSQL = sSQL + "	, RT11.TransCode, Coalesce(OptFee1,0) AS OptFee1, RT12.TransCode, Coalesce(OptFee2,0) AS OptFee2, RT13.TransCode, Coalesce(OptFee3,0) AS OptFee3"
sSQL = sSQL + "	, RT14.TransCode, Coalesce(OptFee4,0) AS OptFee4, RT15.TransCode, Coalesce(OptFee5,0) AS OptFee5, RT16.TransCode, Coalesce(OptFee6,0) AS OptFee6"
sSQL = sSQL + "	, RT17.TransCode, Coalesce(OptFee7,0) AS OptFee7, RT18.TransCode, Coalesce(OptFee8,0) AS OptFee8, RT19.TransCode, Coalesce(OptFee9,0) AS OptFee9"
sSQL = sSQL + "	, RT20.TransCode, Coalesce(OptFee10,0) AS OptFee10"

sSQL = sSQL + "	, Coalesce(EntryFee,0)+Coalesce(LateFee,0)+Coalesce(AWSEF,0)+Coalesce(OffDisc,0)+Coalesce(JrDisc,0)+Coalesce(SrDisc,0)+Coalesce(ClubDisc,0)"
sSQL = sSQL + " + Coalesce(Banquet,0)+Coalesce(OptFee1,0)+Coalesce(OptFee2,0)+Coalesce(OptFee3,0)+Coalesce(OptFee4,0)+Coalesce(OptFee5,0)"
sSQL = sSQL + " + Coalesce(Banquet,0)+Coalesce(OptFee6,0)+Coalesce(OptFee7,0)+Coalesce(OptFee8,0)+Coalesce(OptFee9,0)+Coalesce(OptFee10,0) AS TotalFees"

sSQL = sSQL + "	, CCFirst+' '+CCLast AS CCName, CCAmount, Last4Card" 

sSQL = sSQL + "	FROM" 

sSQL = sSQL + "	(SELECT MemberID, OrderNo, TourID"
sSQL = sSQL + "		FROM usawsrank.regtransactions"
sSQL = sSQL + "			WHERE LEFT(TourID,6) = '"&LEFT(sTourID,6)&"'" 
sSQL = sSQL + "			GROUP BY MemberID, TourID, OrderNo) AS RTN"
		
sSQL = sSQL + "	LEFT JOIN"
sSQL = sSQL + "		(SELECT OrderNo, FirstName AS CCFirst, LastName AS CCLast, Last4Card, Amount AS CCAmount, PayType, TransDate"
sSQL = sSQL + "			FROM usawsrank.RegPaymentLog) AS RPL"
sSQL = sSQL + "	ON RPL.OrderNo=RTN.OrderNo" 

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		(SELECT TransCode, Coalesce(Amount,0.00) AS EntryFee, OrderNo" 
sSQL = sSQL + "			FROM usawsrank.regtransactions"
sSQL = sSQL + "				WHERE TransCode='FEF' OR TransCode='CEF')  AS RT1" 
sSQL = sSQL + "	ON RT1.OrderNo=RTN.OrderNo"

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		(SELECT TransCode, Amount AS LateFee, OrderNo" 
sSQL = sSQL + "			FROM usawsrank.regtransactions"
sSQL = sSQL + "				WHERE TransCode='FLF' OR TransCode='CLF')  AS RT2"
sSQL = sSQL + "	ON RT2.OrderNo=RTN.OrderNo"

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		(SELECT TransCode, Amount AS AWSEF, OrderNo" 
sSQL = sSQL + "			FROM usawsrank.regtransactions"
sSQL = sSQL + "				WHERE TransCode='OBF' OR TransCode='CBF' )  AS RT3"
sSQL = sSQL + "	ON RT3.OrderNo=RTN.OrderNo"

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		(SELECT TransCode, Amount AS OffDisc, OrderNo" 
sSQL = sSQL + "			FROM usawsrank.regtransactions"
sSQL = sSQL + "				WHERE TransCode='DOF' OR TransCode='COF' )  AS RT4"
sSQL = sSQL + "	ON RT4.OrderNo=RTN.OrderNo"

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		(SELECT TransCode, Amount AS JrDisc, OrderNo" 
sSQL = sSQL + "			FROM usawsrank.regtransactions"
sSQL = sSQL + "				WHERE TransCode='DJR' OR TransCode='CJR' )  AS RT5"
sSQL = sSQL + "	ON RT5.OrderNo=RTN.OrderNo"

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		(SELECT TransCode, Amount AS SrDisc, OrderNo" 
sSQL = sSQL + "			FROM usawsrank.regtransactions"
sSQL = sSQL + "				WHERE TransCode='DSR' OR TransCode='CSR' )  AS RT6"
sSQL = sSQL + "	ON RT6.OrderNo=RTN.OrderNo"

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		(SELECT TransCode, Amount AS ClubDisc, OrderNo" 
sSQL = sSQL + "			FROM usawsrank.regtransactions"
sSQL = sSQL + "				WHERE TransCode='DCL' OR TransCode='CCL' )  AS RT7"
sSQL = sSQL + "	ON RT7.OrderNo=RTN.OrderNo"

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		(SELECT TransCode, Amount AS Banquet, OrderNo" 
sSQL = sSQL + "			FROM usawsrank.regtransactions"
sSQL = sSQL + "				WHERE TransCode='BAN' OR TransCode='CAN' )  AS RT8"
sSQL = sSQL + "	ON RT8.OrderNo=RTN.OrderNo"

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		(SELECT TransCode, Amount AS OptFee1, OrderNo" 
sSQL = sSQL + "			FROM usawsrank.regtransactions"
sSQL = sSQL + "				WHERE TransCode='OF1' OR TransCode='CF1' )  AS RT11"
sSQL = sSQL + "	ON RT11.OrderNo=RTN.OrderNo"

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		(SELECT TransCode, Amount AS OptFee2, OrderNo" 
sSQL = sSQL + "			FROM usawsrank.regtransactions"
sSQL = sSQL + "				WHERE TransCode='OF2' OR TransCode='CF2' )  AS RT12"
sSQL = sSQL + "	ON RT12.OrderNo=RTN.OrderNo"

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		(SELECT TransCode, Amount AS OptFee3, OrderNo" 
sSQL = sSQL + "			FROM usawsrank.regtransactions"
sSQL = sSQL + "				WHERE TransCode='OF3' OR TransCode='CF3' )  AS RT13"
sSQL = sSQL + "	ON RT13.OrderNo=RTN.OrderNo"

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		(SELECT TransCode, Amount AS OptFee4, OrderNo" 
sSQL = sSQL + "			FROM usawsrank.regtransactions"
sSQL = sSQL + "				WHERE TransCode='OF4' OR TransCode='CF4' )  AS RT14"
sSQL = sSQL + "	ON RT14.OrderNo=RTN.OrderNo"

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		(SELECT TransCode, Amount AS OptFee5, OrderNo" 
sSQL = sSQL + "			FROM usawsrank.regtransactions"
sSQL = sSQL + "				WHERE TransCode='OF5' OR TransCode='CF5' )  AS RT15"
sSQL = sSQL + "	ON RT15.OrderNo=RTN.OrderNo"

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		(SELECT TransCode, Amount AS OptFee6, OrderNo" 
sSQL = sSQL + "			FROM usawsrank.regtransactions"
sSQL = sSQL + "				WHERE TransCode='OF6' OR TransCode='CF6' )  AS RT16"
sSQL = sSQL + "	ON RT16.OrderNo=RTN.OrderNo"

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		(SELECT TransCode, Amount AS OptFee7, OrderNo" 
sSQL = sSQL + "			FROM usawsrank.regtransactions"
sSQL = sSQL + "				WHERE TransCode='OF7' OR TransCode='CF7' )  AS RT17"
sSQL = sSQL + "	ON RT17.OrderNo=RTN.OrderNo"

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		(SELECT TransCode, Amount AS OptFee8, OrderNo" 
sSQL = sSQL + "			FROM usawsrank.regtransactions"
sSQL = sSQL + "				WHERE TransCode='OF8' OR TransCode='CF8' )  AS RT18"
sSQL = sSQL + "	ON RT18.OrderNo=RTN.OrderNo"

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		(SELECT TransCode, Amount AS OptFee9, OrderNo" 
sSQL = sSQL + "			FROM usawsrank.regtransactions"
sSQL = sSQL + "				WHERE TransCode='OF9' OR TransCode='CF9' )  AS RT19"
sSQL = sSQL + "	ON RT19.OrderNo=RTN.OrderNo"

sSQL = sSQL + "	LEFT JOIN" 
sSQL = sSQL + "		(SELECT TransCode, Amount AS OptFee10, OrderNo" 
sSQL = sSQL + "			FROM usawsrank.regtransactions"
sSQL = sSQL + "				WHERE TransCode='OF10' OR TransCode='CF10' )  AS RT20"
sSQL = sSQL + "	ON RT20.OrderNo=RTN.OrderNo"

sSQL = sSQL + "	LEFT JOIN"
'sSQL = sSQL + "		( SELECT FirstName, LastName, PersonIDWithCheckDigit"
'sSQL = sSQL + "	ON MT.PersonIDWithCheckDigit=RTN.MemberID"
sSQL = sSQL + "		( SELECT FirstName, LastName, PersonID"
sSQL = sSQL + "			FROM usawaterski.dbo.Members) AS MT" 
sSQL = sSQL + "	ON MT.PersonID=CAST(RIGHT(RTN.MemberID,8) AS INT)"
	

sSQL = sSQL + " WHERE RPL.TransDate <= '"&LastDate&"' AND RPL.TransDate >= '"&FirstDate&"'"
sSQL = sSQL + " AND LEFT(RTN.TourID,6) = '"&LEFT(sTourID,6)&"'"

IF WhatPayments <> "ALL" THEN sSQL = sSQL + " AND RPL.PayType = '"&WhatPayments&"'"
SELECT CASE SequenceSelected 
	CASE "alpha"
		sSQL = sSQL + " ORDER BY MT.LastName, MT.FirstName, RPL.TransDate"
	CASE "chgdate"
		sSQL = sSQL + " ORDER BY RPL.TransDate DESC, RTN.MemberID"
END SELECT

'IF sTourID="08S999A" THEN 
	'response.write(sSQL)
	'response.end
'END IF

rsRT.open sSQL, SConnectionToTRATable, 3, 3




IF NOT rsRT.eof THEN 

	rsRT.movefirst		


	%>
	<TR>
	  <th align="Left"><font size=<%=fontsize2%> color="#FFFFFF"><b>Participant</b></FONT></th>
	  <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>Date</b></FONT></th>
	  <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>Type</b></FONT></th>
	  <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>OrderNo</b></FONT></th>
	  <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>EntryFee</b></FONT></th>
	  <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>OffDisc</b></FONT></th>
	  <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>JrDisc</b></FONT></th>
	  <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>SrDisc</b></FONT></th>
	  <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>ClubDisc</b></FONT></th>
	  <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>LateFee</b></FONT></th>
	  <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>AWSEF</b></FONT></th>

	  <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>Banqt</b></FONT></th>
		<%
		IF TRIM(sOF1Desc)<>"" THEN 
	  		%><th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>OptFee1</b></FONT></th><%
	  END IF		
		IF TRIM(sOF2Desc)<>"" THEN 	  
			  %><th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>OptFee2</b></FONT></th><%
	  END IF		
		IF TRIM(sOF3Desc)<>"" THEN 	  
	  		%><th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>OptFee3</b></FONT></th><%
	  END IF		
		IF TRIM(sOF4Desc)<>"" THEN 
	  		%><th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>OptFee4</b></FONT></th><%
	  END IF		
		IF TRIM(sOF5Desc)<>"" THEN 	  
			  %><th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>OptFee5</b></FONT></th><%
	  END IF		
		IF TRIM(sOF6Desc)<>"" THEN 	  
	  		%><th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>OptFee6</b></FONT></th><%
	  END IF		
		IF TRIM(sOF7Desc)<>"" THEN 
	  		%><th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>OptFee7</b></FONT></th><%
	  END IF		
		IF TRIM(sOF8Desc)<>"" THEN 	  
			  %><th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>OptFee8</b></FONT></th><%
	  END IF		
		IF TRIM(sOF9Desc)<>"" THEN 	  
	  		%><th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>OptFee9</b></FONT></th><%
	  END IF		
		IF TRIM(sOF10Desc)<>"" THEN 	  
	  		%><th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>OptFee10</b></FONT></th><%
	  END IF		

		%>
	  <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>TotalFees</b></FONT></th>
	  <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>CardHolder</b></FONT></th>
	  <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>Payment</b></FONT></th>
	  <th align="Center"><font size=<%=fontsize2%> color="#FFFFFF"><b>Last-4</b></FONT></th>
	</TR><%

	
	C=1	
	DO WHILE NOT rsRT.eof	%>
		<TR>
		  <TD align="Left" valign="top">
			<font size="<%=fontsize1%>">&nbsp;<a title="MemberID: <%=rsRT("MemberID")%>"><% =TRIM(rsRT("LastName"))%>, <%=TRIM(rsRT("FirstName")) %></a></FONT>
		  </TD>
		  <TD align="Center" ><font size="<%=fontsize1%>">&nbsp;<% =rsRT("TransDate") %></FONT></TD>
		  <TD align="Center" ><font size="<%=fontsize1%>">&nbsp;<% =rsRT("PayType") %></FONT></TD>
		  <TD align="Center" ><font size="<%=fontsize1%>">&nbsp;<% =rsRT("OrderNo") %></FONT></TD>
		  <TD align="Center" ><font size="<%=fontsize1%>"><%= FormatNumber(rsRT("EntryFee"),2) %></FONT></TD>
		  <TD align="Center" ><font size="<%=fontsize1%>"><% =FormatNumber(rsRT("OffDisc"),2) %></FONT></TD>
		  <TD align="Center" ><font size="<%=fontsize1%>"><% =FormatNumber(rsRT("JrDisc"),2) %></FONT></TD>
		  <TD align="Center" ><font size="<%=fontsize1%>"><% =FormatNumber(rsRT("SrDisc"),2) %></FONT></TD>
		  <TD align="Center" ><font size="<%=fontsize1%>"><% =FormatNumber(rsRT("ClubDisc"),2) %></FONT></TD>
		  <TD align="Center" ><font size="<%=fontsize1%>"><% =FormatNumber(rsRT("LateFee"),2) %></FONT></TD>
		  <TD align="Center" ><font size="<%=fontsize1%>"><% =FormatNumber(rsRT("AWSEF"),2) %></FONT></TD>

		  <TD align="Center" ><font size="<%=fontsize1%>"><%=FormatNumber(rsRT("Banquet"),2)%></FONT></TD>
			<%	
			
			IF TRIM(sOF1Desc)<>"" THEN 	  
		  		%><TD align="Center"><font size="<%=fontsize1%>"><%=FormatNumber(rsRT("OptFee1"),2)%></FONT></TD><%
		  END IF		
			IF TRIM(sOF2Desc)<>"" THEN 	  
		  		%><TD align="Center"><font size="<%=fontsize1%>"><%=FormatNumber(rsRT("OptFee2"),2)%></FONT></TD><%
		  END IF		
			IF TRIM(sOF3Desc)<>"" THEN 	  
		  		%><TD align="Center"><font size="<%=fontsize1%>"><%=FormatNumber(rsRT("OptFee3"),2)%></FONT></TD><%
		  END IF		
			IF TRIM(sOF4Desc)<>"" THEN 	  
		  		%><TD align="Center"><font size="<%=fontsize1%>"><%=FormatNumber(rsRT("OptFee1"),2)%></FONT></TD><%
		  END IF		
			IF TRIM(sOF5Desc)<>"" THEN 	  
		  		%><TD align="Center"><font size="<%=fontsize1%>"><%=FormatNumber(rsRT("OptFee2"),2)%></FONT></TD><%
		  END IF		
			IF TRIM(sOF6Desc)<>"" THEN 	  
		  		%><TD align="Center"><font size="<%=fontsize1%>"><%=FormatNumber(rsRT("OptFee6"),2)%></FONT></TD><%
		  END IF		
			IF TRIM(sOF7Desc)<>"" THEN 	  
		  		%><TD align="Center"><font size="<%=fontsize1%>"><%=FormatNumber(rsRT("OptFee7"),2)%></FONT></TD><%
		  END IF		
			IF TRIM(sOF8Desc)<>"" THEN 	  
		  		%><TD align="Center"><font size="<%=fontsize1%>"><%=FormatNumber(rsRT("OptFee8"),2)%></FONT></TD><%
		  END IF		
			IF TRIM(sOF9Desc)<>"" THEN 	  
		  		%><TD align="Center"><font size="<%=fontsize1%>"><%=FormatNumber(rsRT("OptFee9"),2)%></FONT></TD><%
		  END IF		
			IF TRIM(sOF10Desc)<>"" THEN 	  
		  		%><TD align="Center"><font size="<%=fontsize1%>"><%=FormatNumber(rsRT("OptFee10"),2)%></FONT></TD><%
		  END IF		

			%>
		  <TD align="Center" ><font size="<%=fontsize1%>"><% =FormatNumber(rsRT("TotalFees"),2) %></FONT></TD>
		  <TD align="Center" ><font size="<%=fontsize1%>">&nbsp;<% =rsRT("CCName") %></a></FONT></TD>
		  <TD align="Center" ><font size="<%=fontsize1%>"><% =rsRT("CCAmount") %></FONT></TD>
		  <TD align="Center" ><font size="<%=fontsize1%>">&nbsp;<%=rsRT("Last4Card")%></FONT></TD>
		</TR><%

		rsRT.MoveNext

	LOOP

ELSE
	%><TD align="Center" style="background-color:<%=Tablecolor1%>;" >
		<font size=<% =fontsize3 %> color="<% =Textcolor3 %>"><i><b>No Data For These Settings</b></i></FONT>
	</TD><%

END IF  %>

</TABLE><%

rsRT.close
Set rsRT = nothing



END SUB

' -----------------------
  SUB PayReportDataLine
' -----------------------

%>
<TR>
  <TD align="Left" valign="top">
	<font size="<%=fontsize1%>">&nbsp;<a title="MemberID: <%=sMemberID%>"><% =FullName %></a></FONT>
  </TD>
  <TD align="Center" ><font size="<%=fontsize1%>">&nbsp;<% =sTransDate %></FONT></TD>
  <TD align="Center" ><font size="<%=fontsize1%>">&nbsp;<% =sPayType %></FONT></TD>
  <TD align="Center" ><font size="<%=fontsize1%>">&nbsp;<% =sOrderNo %></FONT></TD>
  <TD align="Center" ><font size="<%=fontsize1%>"><% =sEntryFees %></FONT></TD>
  <TD align="Center" ><font size="<%=fontsize1%>"><% =sOffDiscAmt %></FONT></TD>
  <TD align="Center" ><font size="<%=fontsize1%>"><% =sJrDiscAmt %></FONT></TD>
  <TD align="Center" ><font size="<%=fontsize1%>"><% =sSrDiscAmt %></FONT></TD>
  <TD align="Center" ><font size="<%=fontsize1%>"><% =sClubDiscAmt %></FONT></TD>
  <TD align="Center" ><font size="<%=fontsize1%>"><% =sLateFee %></FONT></TD>
  <TD align="Center" ><font size="<%=fontsize1%>"><% =sAWSEFDonation %></FONT></TD>
  <TD align="Center" ><font size="<%=fontsize1%>"><% =sTotalFees %></FONT></TD>
  <TD align="Center" ><font size="<%=fontsize1%>">&nbsp;<% =CCFullName %></a></FONT></TD>
  <TD align="Center" ><font size="<%=fontsize1%>"><% =Charge %></FONT></TD>
  <TD align="Center" ><font size="<%=fontsize1%>">&nbsp;<%=sLast4%></FONT></TD>
</TR><%


END SUB




' --------------------
  SUB LoadDropSequence
' --------------------

%>
  <td align=right>
	<font size=<% =fontsize3 %> color="<% =TextDropcolor2 %>">Order By:</font>
  </td>

  <td align=left>	
	<select name="SequenceSelected" style="width:11em">
	        <option value="alpha"<%IF SequenceSelected = "alpha" THEN Response.Write(" SELECTED ")%>>Alphabetic</option>
        	<option value="seed"<%IF SequenceSelected = "seed" THEN Response.Write(" SELECTED ")%>>Seeding Value</option>
	        <option value="regdate"<%IF SequenceSelected = "regdate" THEN Response.Write(" SELECTED ")%>>Register Date by Div</option>
	        <option value="regdateall"<%IF SequenceSelected = "regdateall" THEN Response.Write(" SELECTED ")%>>Register Date - All</option>
	</select>
  </td><%

END SUB


' ----------------------
   SUB LoadReportPullDown 
' ----------------------

'response.write("<br>Line 4216 - sWhatReport= "&sWhatReport)

%>
  <td align=right>
    <font size=<% =fontsize3 %> color="<% =TextDropcolor2 %>">Type:</font>
  </td>
  <td align=left>
     <select name="WhatReport" style="width:11em;">
			<option value="noreportselected"<% IF sWhatReport = "noreportselected" THEN Response.Write(" SELECTED ") %>>Select Report</option>
      <option value="regstat"<% IF sWhatReport = "regstat" THEN Response.Write(" SELECTED ") %>>Registration Status</option>
			<option value="seeding"<% IF sWhatReport = "seeding" THEN Response.Write(" SELECTED ") %>>Seeding</option>
      <option value="scratched"<% IF sWhatReport = "scratched" THEN Response.Write(" SELECTED ") %>>Not Ready To Ski</option>
			<% 
			IF adminmenulevel>=1 OR LCASE(Session("UserAdminPW"))=LCASE(Session("AdminCode")) THEN 
					%> 	
       	  <option value="skierpayments" <% IF sWhatReport = "skierpayments" THEN Response.Write(" SELECTED ")%> >Payments by Type</option>
          <option value="financial"<% IF sWhatReport = "financial" THEN Response.Write(" SELECTED ")%> >Payments Received</option>
					<option value="othersales" <% IF sWhatReport = "othersales" THEN Response.Write(" SELECTED ")%> >Other Sales</option>
        	<% 
			END IF 
			IF adminmenulevel >= 19  THEN 
       	  %><option value="notifications" <% IF sWhatReport = "notifications" THEN Response.Write(" SELECTED ")%> >Notifications</option><%
      END IF
			IF adminmenulevel >= 19 OR TestValidAdminCode=true  THEN 
					%><option value="notifications" <% IF sWhatReport = "PrintBio" THEN Response.Write(" SELECTED ")%> >Print Bios</option><% 
			END IF 
			%>
			<option value="divisiontotals"<% IF sWhatReport = "divisiontotals" THEN Response.Write(" SELECTED ")%> >Scheduling</option>
			<option value="bystate"<% IF sWhatReport = "bystate" THEN Response.Write(" SELECTED ")%> >Skiers By State</option>
     </select>
  </td>
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



' ------------------------
   SUB LoadEventPulldown
' ------------------------

' ***************  OBSOLETE  ***************


%>
<td align=right>
  <font size=<% =fontsize3 %> color="<% =TextDropcolor2 %>">Event:</font>
</td>
<td align=left>
<select name="EventSelected" style="width:6em">
<option value="ALL" <% IF EventSelected = "ALL" THEN Response.Write(" SELECTED ")%> >All</option><%

SELECT CASE sSptsGrpID
	CASE "AWS"
		%>
		<option value="S" <%IF EventSelected = "S" THEN Response.Write(" SELECTED ")%>>Slalom</option>
		<option value="T" <%IF EventSelected = "T" THEN Response.Write(" SELECTED ")%>>Tricks</option>
		<option value="J" <%IF EventSelected = "J" THEN Response.Write(" SELECTED ")%>>Jump</option>
		<%
	CASE "ABC"
		%>
		<option value="W" <%IF EventSelected = "W" THEN Response.Write(" SELECTED ")%>>Wake Cross</option>
		<option value="T" <%IF EventSelected = "T" THEN Response.Write(" SELECTED ")%>>Tricks</option>
		<option value="J" <%IF EventSelected = "J" THEN Response.Write(" SELECTED ")%>>Jump</option>
		<%
	CASE "USW"
		%>
		<option value="W" <%IF EventSelected = "W" THEN Response.Write(" SELECTED ")%>>Wakeboard</option>
		<option value="WS" <%IF EventSelected = "WS" THEN Response.Write(" SELECTED ")%>>Wake Skate</option>
		<option value="WU" <%IF EventSelected = "WU" THEN Response.Write(" SELECTED ")%>>Wake Surf</option>
		<%
	CASE "USH"
		%>
		<option value="W" <%IF EventSelected = "HB" THEN Response.Write(" SELECTED ")%>>Big Air</option>
		<option value="WS" <%IF EventSelected = "WS" THEN Response.Write(" SELECTED ")%>>Free Ride</option>
		<option value="WU" <%IF EventSelected = "WU" THEN Response.Write(" SELECTED ")%>>Jump Out</option>
		<%

	END SELECT
	%>
</select>
</td><%

END SUB



' ------------------------
   SUB LoadEventPulldownNew
' ------------------------

'response.write("<br>EventSelected = "&EventSelected)
'response.write("<br>EventSelected = "&EventSelected)
%>
<td align=right>
  <font size=<% =fontsize3 %> color="<% =TextDropcolor2 %>">Event:</font>
</td>
<td align=left>
<select name="EventSelected" style="width:6em">
<option value="ALL" <% IF EventSelected = "ALL" THEN Response.Write(" SELECTED ")%> >All</option><%

	FOR EvtNo = 1 TO TotEv 
		IF TRIM(sTEvent(EvtNo)) <> "" THEN %>
			<option value="<%=sTEvent(EvtNo)%>" <% IF EventSelected = ""&sTEvent(EvtNo)&"" THEN Response.Write(" SELECTED ") %>><%=sTEventName(EvtNo)%></option><%
		END IF
	NEXT %>

</select>
</td><%

END SUB




' ------------------------
   SUB LoadRegionPulldown
' ------------------------

%>
<td align=right>
	<font size=<% =fontsize3 %> color="<% =TextDropcolor2 %>"><a TITLE="Region is based on State of Membership record.">Region:<a></font>
</td>
<td align=left>
  <select name="RegionSelected">
	<option value=""<%IF RegionSelected = "" THEN Response.Write(" SELECTED ")%>>All Regions</option>
	<option value="1"<%IF RegionSelected = "1" THEN Response.Write(" SELECTed ")%>>S. Central</option>
	<option value="2"<%IF RegionSelected = "2" THEN Response.Write(" SELECTED ")%>>Midwest</option>
	<option value="3"<%IF RegionSelected = "3" THEN Response.Write(" SELECTED ")%>>West</option>
	<option value="4"<%IF RegionSelected = "4" THEN Response.Write(" SELECTED ")%>>South</option>
	<option value="5"<%IF RegionSelected = "5" THEN Response.Write(" SELECTED ")%>>East</option>
  </select>
</td><%


END SUB


' ------------------------
   SUB LoadStatePulldown
' ------------------------

StateArray = Split(USStatesList2,",")  %>

<td align=right>
	<font size=<% =fontsize3 %> color="<% =Textcolor1 %>">State:</font>
</td>
<td align=left>
    <select name="StateSelected"><%
	  response.write("<option value = ""All"" SELECTED>All</option>")
	
	  FOR kvar = 0 TO UBOUND(StateArray)
	    IF TRIM(StateArray(kvar)) = "" THEN
		' - Blank do nothing	
	    ELSEIF TRIM(StateSelected) = TRIM(StateArray(kvar)) THEN
		response.write("<option value = """&StateSelected&""" SELECTED>"&StateSelected&"</option>")
	    ELSE
		response.write("<option value = """&StateArray(kvar)&""">"&StateArray(kvar)&"</option>")
	    END IF
	  NEXT  %>
    </select>
</td><%

END SUB




' ---------------------------------
   SUB LoadStartPulldown
' ---------------------------------

AlphaArray = Split(AlphaList,",")  %>
<td align="right">
	<font size=<% =fontsize3 %> color="<% =Textcolor1 %>">Start With:</font>
</td>
<td align="left">
	<select name="StartPulldown"><%
	    response.write("<option value = ""All"" SELECTED>All</option>")
	
	  FOR kvar = 0 TO UBOUND(AlphaArray)
	    IF TRIM(AlphaArray(kvar)) = "" THEN
		' - Blank do nothing	
	    ELSEIF TRIM(StartCharSelected) = TRIM(AlphaArray(kvar)) THEN
		response.write("<option value = """&StartCharSelected&""" SELECTED>"&StartCharSelected&"</option>")
	    ELSE
		response.write("<option value = """&AlphaArray(kvar)&""">"&AlphaArray(kvar)&"</option>")
	    END IF
	  NEXT  %>
	</select>
</td><%

END SUB


' ---------------------------------
   SUB LoadEndPulldown
' ---------------------------------

AlphaArray = Split(AlphaList,",")  %>
<td align="right">
	<font size=<% =fontsize3 %> color="<% =Textcolor1 %>">End With:</font>
</td>
<td align="left">
	 <select name="EndPulldown"><%
	    response.write("<option value = ""All"" SELECTED>All</option>")
	
	  FOR kvar = 0 TO UBOUND(AlphaArray)
	    IF TRIM(AlphaArray(kvar)) = "" THEN
		' - Blank do nothing	
	    ELSEIF TRIM(StartCharSelected) = TRIM(AlphaArray(kvar)) THEN
		response.write("<option value = """&EndCharSelected&""" SELECTED>"&EndCharSelected&"</option>")
	    ELSE
		response.write("<option value = """&AlphaArray(kvar)&""">"&AlphaArray(kvar)&"</option>")
	    END IF
	  NEXT  %>
</select>
</td><%

END SUB





' ------------------------------------------------------------
  SUB LoadTimePulldown (TimeAmount, MinTime, MaxTime, StepTime)
' ------------------------------------------------------------

Dim iCounter

TimeAmount = Cint(TimeAmount)

'response.write("<option value = 0 >NA</option>")

FOR iCounter = MinTime TO MaxTime STEP StepTime

	mymin=Fix(iCounter/60)

	mysec=iCounter - 60*mymin
	IF cdbl(mysec) = 0 THEN mySec = "00"
	myMinSec=mymin&":"&mysec

	IF iCounter = TimeAmount THEN
		response.write("<option value = """&iCounter&""" SELECTED>"&myMinSec&"</option>")
	ELSE
		response.write("<option value = """&iCounter&""">"&myMinSec&"</option>")
	END IF
NEXT


END SUB


%>




