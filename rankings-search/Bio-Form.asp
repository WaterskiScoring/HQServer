<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->

<%

Dim sFirstName, sLastName, sMembAge, sMembSex, sMemberID, sTourID
Dim sAddress1, sregion, sCity, sState, sEmail, sPhone, sWeight, sHgtFeet, sHgtInch, sSkiSinceAge, sCompSinceAge, sMembSinceAge, sSponsors
Dim sClub, sSchool, sOccup, sCareer, sHobby, sPaper, sBestSlal, sBestHydro, sBestTrick, sBestFree, sBestJump, sBestKnee, sBestWake, sBestMara
Dim sLastUpdate
Dim sRunByWhat, FormStatus, BioStatus
Dim sFav_Slalom, sFav_Jump, sFav_Trick, sFav_Wake, sFav_Boat, sMentors, sAccomplish, sFav_Sports, sTitles, sRecords


DefineTRAStyles

sRunByWhat = TRIM(Request("sRunByWhat"))
FormStatus = TRIM(Request("FormStatus"))
BioStatus = TRIM(Request("BioStatus"))
ThisPrint=TRIM(Request("ThisPrint"))





' --- This is the section that screens for Session("sMemberID") set
sMemberID = trim(Request("sMemberID"))


IF sMemberID = "" THEN 
		IF Session("sMemberID")="" THEN
				sRunByWhat="TimeoutNotice"
		ELSE
				sMemberID = Session("sMemberID")
		END IF
END IF

sTourID = TRIM(Request("sTourID"))
IF sTourID = "" THEN 
		IF Session("sTourID")="" THEN
				sRunByWhat="TimeoutNotice"
		ELSE
				sTourID = Session("sTourID")
		END IF
END IF

IF 2=1 AND sMemberID="000001151" THEN
		response.write("<br>Line 47 - sMemberID =" &sMemberID)
		response.write("<br>sTourID =" &TRIM(Request("sTourID")))
		response.write("<br>sSendingPage =" &TRIM(Request("sTourID")))
		'response.end
END IF





' --- Add this once things are settled so default is NO access  -----
'IF TRIM(Request("BioStatus")) = "" THEN BioStatus = "disabled"




IF sRunByWhat<>"TimeoutNotice" THEN

		' ----------------------------------------
		' ----  Read tournament information  -----
		' ----------------------------------------
		IF TRIM(Request("sTourID"))<>"999999" AND TRIM(Session("sTourID"))<>"999999" THEN
				SET rsSanc=Server.CreateObject("ADODB.recordset")
				sSQL = "SELECT TOP 1 * from " & SanctionTableName
				sSQL = sSQL + " LEFT JOIN "&GuideBookTableName&" ON "&SanctionTableName&".TournAppID = "&GuideBookTableName&".GTournAppID"
				sSQL = sSQL + " WHERE LEFT(TournAppID,6) = '" & SQLClean(left(sTourID,6)) & "'"
				'response.write(sSQL)
				'response.end
				rsSanc.open sSQL, SConnectionToTRATable

				' -- Define page variables from SWIFT Guidebook table	
				sTourName = rsSanc("TName")
				sTourCity = rsSanc("TCity")
				sTourState = rsSanc("TState")
				sTourSDate = rsSanc("TDateS")
				sTourEDate = rsSanc("TDateE")
				sSptsGrpID = rsSanc("SptsGrpID")


				sMembAge = AgeAtDate_New(sTourSDate, sMemberID)
		ELSE
		
				sMembAge=0
		END IF


		IF FormStatus="new" THEN
				SET rsBio=Server.CreateObject("ADODB.recordset")
				sSQL = "SELECT * FROM "&BioTableName
				sSQL = sSQL + " WHERE MemberID = '"&sMemberID&"'"
				rsBio.open sSQL, SConnectionToTRATable, 3, 3

				IF NOT rsBio.eof	THEN	
						' --- Found EXISTING record --
						ReadFromTable
				ELSE
						' --- No BIO record found --
						sSQL = "SELECT TOP 1 * FROM "&MemberLiveTableName&" AS MT"
						sSQL = sSQL + " LEFT JOIN "&MemberTypeOLRTableName&" AS MTT ON MTT.MembershipTypeID = MT.MembershipTypeCode"

						sSQL = sSQL + " WHERE PersonID = "&RIGHT(sMemberID,8)
						set rsMemb=Server.CreateObject("ADODB.recordset")
						rsMemb.open sSQL, sConnectionToTRATable, 3, 1

						sFirstName = rsMemb("FirstName")
						sLastName = rsMemb("LastName")
						sMembSex = rsMemb("Sex")
						'sMembAge = Session("MembAge")
				END IF
		END IF

END IF







SELECT CASE sRunByWhat

		CASE "DoneWindow"	
				DisplayDoneWindow

   CASE "TimeoutNotice"
				DisplayTimeoutNotice

   CASE "SaveRecord"
				ReadFromForm

				IF BioStatus <> "disabled" THEN SaveBio

				DisplayDoneWindow


   CASE ELSE

   			' ---------------------------------------------------------------------------
   			' --- IF value passed from Registration.asp then load previous if exists ----
				' ---------------------------------------------------------------------------

				'--- Writes header unless the user intended to print the page.
				IF ThisPrint<>"YES" THEN
						'	WriteIndexPageHeader
				END IF  

				' ---- TOP OF MAIN FORM  ---- 
				%> 
		  	<form action="/rankings/bio-form.asp?sRunByWhat=SaveRecord" method="post">
		  		<input type="hidden" name="sMemberID" value="<% =sMemberID %>">
		  		<input type="hidden" name="sTourID" value="<% =sTourID %>">
		  		<%

					IF BioStatus="disabled" THEN 
							%><input type="hidden" name="BioStatus" value="<%=BioStatus%>"><%
					END IF 

				%>
				<br>
      	<TABLE class="innertable" ALIGN="CENTER" width=80%>
					<TR>
	  				<TH colspan="8" align="center"><FONT size="3" COlOR="#FFFFFF"><strong>Personal Bio Form</strong></FONT></TH>
					</TR>  
 					<TR>
	  				<TD VALIGN="top">
							<table ALIGN="center" BORDER="1" CELLPADDING="3" CELLSPACING="0" width=100% BGCOLOR="<%=TableColor1%>">   	

		   					<tr> 
	    						<td ALIGN="right"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Name:</b>&nbsp;&nbsp</FONT></td>
	    						<td align="left"><font size="<% =fontsize3 %>" face=<% =font1 %> COlOR="<% =textcolor2 %>"><% =sFirstName %>&nbsp;<% =sLastName %></font></td>

	    						<td ALIGN="right"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Last Updated:</b>&nbsp;&nbsp</FONT></td>
	    						<td align="left"><font size="<% =fontsize3 %>" face=<% =font1 %> COlOR="<% =textcolor2 %>">&nbsp;<% =sLastUpdate %></font></td>

   	    					<td ALIGN="right"><FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR=<% =textcolor1 %>><b>Age/Gender:</b>&nbsp;&nbsp</FONT></td>
	    						<td align="left"><font size=<% =fontsize3 %> face=<% =font1 %> COlOR="<% =textcolor2 %>"><% =sMembAge %>/<% =sMembSex %></font></td> 
	  						</tr>
								<tr>
	    						<td ALIGN="right"><FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR=<% =textcolor4 %>>&nbsp;&nbsp;&nbsp;<b>City:</b></FONT></td>
	    						<td align="left">
										<input type="text" <% =BioStatus %> name="fCity" value= "<% =sCity %>"  maxlength=20 size="20">
										<%
										IF ChargeStatus="confirm" AND TRIM(sCity)="" THEN
												FieldErr = FieldErr +1
												%><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="red">* Required</font><%
										END IF  
										%>			
	    						</td>
									<td ALIGN="right"><FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR=<% =textcolor4 %>><b>State:</b></FONT></td>
    							<td align="left">
    								<%
										IF ChargeStatus="confirm" AND TRIM(sState)="" THEN
												FieldErr = FieldErr +1
												%><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="red">* Required</font><%
										ELSEIF ChargeStatus="confirm" THEN 
												%><input type="text" <% =BioStatus %> name="fState" value= "<% =sState %>" size="2"><%
										ELSE  
												StateArray = Split(USStatesList2,",")  %>
												<select name="fState" <% =BioStatus %>><%
			  									FOR kvar = 0 TO UBOUND(StateArray)
			    										IF TRIM(sState) = TRIM(StateArray(kvar)) THEN
																	response.write("<option value = """&sState&""" SELECTED>"&sState&"</option>")
			   											ELSE
																	response.write("<option value = """&StateArray(kvar)&""">"&StateArray(kvar)&"</option>")
			    										END IF
			  									NEXT  %>
												</select><%
										END IF  %>
	    						</td>

	    						<td align="right"><font size=<% =fontsize2 %> face=<% =font1 %>><b>Home Region:</b></font></td> 
	    						<td align="left">
	    							<select name='fregion' <% =BioStatus %> value=<% =sRegion %> >
		  								<option value ='0'<%IF sRegion = "0" THEN Response.Write(" selected ")%>>Region</Option><br>
											<option value ='1'<%IF sRegion = "1" THEN Response.Write(" selected ")%>>S. Central</Option><br>
		  								<option value ='2'<%IF sRegion = "2" THEN Response.Write(" selected ")%>>MidWest</Option><br>
		  								<option value ='3'<%IF sRegion = "3" THEN Response.Write(" selected ")%>>West</Option><br>
		  								<option value ='4'<%IF sRegion = "4" THEN Response.Write(" selected ")%>>South</Option><br>
		  								<option value ='5'<%IF sRegion = "5" THEN Response.Write(" selected ")%>>East</Option><br>
										</select>
	    						</td>
	    					</tr>
								<tr> 
	    						<td ALIGN="right"><FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Email:</b></FONT></td>
									<td colspan="3" align="left">
										<input type="text" <% =BioStatus %> name="fEmail" value= "<% =sEmail %>" maxlength=30 size="40">
									</td>
									<%
									IF ChargeStatus="confirm" AND TRIM(sEmail)="" THEN
											' FieldErr = FieldErr +1  Not required
											%><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor2 %>">&nbsp</font><%
									END IF 
									 %>
									<td ALIGN="right"><FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR=<% =textcolor4 %>>&nbsp;&nbsp;&nbsp;<b>Phone:</b></FONT></td>
									<td align="left">
										<input type="text" <% =BioStatus %> name="fPhone" value= "<% =sPhone %>" maxlength=12 size=12>
										<%
										IF ChargeStatus="confirm" AND TRIM(sCity)="" THEN
												FieldErr = FieldErr +1
												%><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="red">* Required</font><%
										END IF  
										%>			
									</td>
								</tr>	
								<tr>  
									<td ALIGN="right"><FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR=<% =textcolor4 %>>&nbsp;&nbsp;&nbsp;<b>Weight:</b></FONT></td>
										<%
										IF ChargeStatus="confirm" THEN
												FieldErr = FieldErr +1
												%><td><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="red">* Required</font></td><%
										ELSEIF ChargeStatus="confirm" THEN
												%><td><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor2 %>"><% =sWeight %></font></td><%
										ELSE
												%>
												<td align="left"><select <% =BioStatus %> name="fWeight" ><% LoadDropDown "LBS", sWeight,50,300,5  %></select>
													<font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"> lbs</font>
												</td>
												<%
										END IF  
										%>			
									<td><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Height: </b></font>
										<select name="fHgtFeet" <% =BioStatus %>><% LoadDropDown "FT", sHgtFeet,3,6,1  %></select>
										<font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>">Ft&nbsp;&nbsp;&nbsp;</font>
										<select name="fHgtInch" <% =BioStatus %>><% LoadDropDown "IN", sHgtInch,1,12,1 %></select>
										<font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>">In</font>
									</td>
									<td colspan=3>&nbsp</td>
								</tr>

	   						<tr>
									<td align="right"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Skiing Since:</b></font></td>
									<td align="left">
										<select name="fSkiSinceAge" <% =BioStatus %>><% LoadDropDown "Age", sSkiSinceAge,2,90,1  %></select>
										<font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>">Years Old</font>
									</td>
									<td align="right"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Competing Since:</b></font></td>
									<td align="left">
										<select name="fCompSinceAge" <% =BioStatus %>><% LoadDropDown "Age", sCompSinceAge,2,90,1  %></select>
										<font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>">Years Old</font>
									</td>
									<td align="right"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Member Since:</b></font></td>
									<td align="left">
										<select name="fMembSinceAge" <% =BioStatus %>><% LoadDropDown "Age", sMembSinceAge,2,90,1  %></select>
										<font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>">Years Old</font>
									</td>
	    					</tr>
	    				</table>
	    				<%
	    				
	    				' --- Second Section --
	    				
	    				%>
							<br>
							<table ALIGN="center" BORDER="1" CELLPADDING="3" CELLSPACING="0" width=100% BGCOLOR="<%=TableColor1%>">   	
								<tr>
									<td ALIGN="right" colspan=1><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Occupation:</b></font></td>
									<td colspan=2 align="left"><input type="text" <% =BioStatus %> name="fOccup" value= "<% =sOccup %>" maxlength="25" size="30"></td>
									<td ALIGN="right" colspan=1><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Career Plans:</b></font></td>
									<td colspan=2 align="left"><input type="text" <% =BioStatus %> name="fCareer" value= "<% =sCareer %>" maxlength="25" size="30"></td>
	    					</tr>
								<tr>
									<td ALIGN="right" colspan=1><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Hobbies:</b></font></td>
									<td colspan=2 align="left"><input type="text" <% =BioStatus %> name="fHobby" value= "<% =sHobby %>" maxlength="49" size="55"></td>
									<td ALIGN="right" colspan=1><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>School:</b></font></td>
									<td colspan=2 align="left"><input type="text" <% =BioStatus %> name="fSchool" value= "<% =sSchool %>"  maxlength="25" size="30"></td>
	     					</tr>
								<tr>
									<td ALIGN="right" colspan=1><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Mentors:</b></font></td>
									<td colspan=2 align="left"><input type="text" <% =BioStatus %> name="fMentors" value= "<% =sMentors %>" maxlength="59" size="65"></td>
									<td ALIGN="right" colspan=1><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Ski Club:</b></font></td>
									<td colspan=2 align="left"><input type="text" <% =BioStatus %> name="fClub" value= "<% =sClub %>" maxlength="25" size="30"></td>
	     					</tr>
								<tr>
									<td ALIGN="right"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Accomplishments:</b></font></td>
									<td colspan=2 align="left"><input type="text" <% =BioStatus %> name="fAccomplish" value= "<% =sAccomplish %>" maxlength="59" size="65"></td>
									<td ALIGN="right" colspan=1><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Local Paper:</b></font></td>
									<td colspan=2 align="left"><input type="text" <% =BioStatus %> name="fPaper" value= "<% =sPaper %>" maxlength="34" size="40"></td>
								</tr>
								<tr>
									<td ALIGN="right"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Records:</b></font></td>
									<td colspan=3 align="left"><input type="text" <% =BioStatus %> name="fRecords" value= "<% =sRecords %>" maxlength="79" size="85"></td>
									<td>&nbsp</td>
								</tr>
								<tr>
									<td ALIGN="right"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Titles:</b></font></td>
									<td colspan=3 align="left"><input type="text" <% =BioStatus %> name="fTitles" value= "<% =sTitles %>" maxlength="79" size="85"></td>
									<td>&nbsp</td>
								</tr>
								<tr>
									<td ALIGN="right" colspan=1><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Favorite Other Sports:</b></font></td>
									<td colspan=2 align="left"><input type="text" <% =BioStatus %> name="fFav_Sports" value= "<% =sFav_Sports %>" maxlength="49" size="55"></td>
									<td>&nbsp</td>
									<td>&nbsp</td>
								</tr>
								<tr>
									<td align="right"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Sponsors:</b></font></td>
									<td colspan=2 align="left"><input type="text" <% =BioStatus %> name="fSponsors" value= "<% =sSponsors %>"  maxlength="59" size="65"></td>
									<td ALIGN="right"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Favorite Boat:</b></font></td>
									<td align="left"><input type="text" <% =BioStatus %> name="fFav_Boat" value= "<% =sFav_Boat %>" maxlength="29" size="33"></td>
	    					</tr>
							</table>
							<%
							
							' --- Third Section --
							
							%>
							<br>
							<table ALIGN="center" BORDER="1" CELLPADDING="3" CELLSPACING="0" width=100% BGCOLOR="<%=TableColor1%>">   	
								<tr>
									<td WIDTH=10%>&nbsp</td>
									<td WIDTH=15% ALIGN="left"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Tournament Best</b></font></td>
									<td ALIGN="left"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Ski/Board (Model/Length)</b></font></td>
	    					</tr>
	    					<tr>
									<td ALIGN="right" colspan=1><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Slalom:</b></font></td>
									<td align="left"><input type="text" <% =BioStatus %> name="fBestSlal" value= "<% =sBestSlal %>" maxlength="15" size="15"></td>
									<td align="left"><input type="text" <% =BioStatus %> name="fFav_Slalom" value= "<% =sFav_Slalom %>" maxlength="20" size="20"></td>
	    					</tr>
								<tr>
									<td ALIGN="right" colspan=1><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Trick:</b></font></td>
									<td align="left"><input type="text" <% =BioStatus %> name="fBestTrick" value= "<% =sBestTrick %>" maxlength="15" size="15"></td>
									<td align="left"><input type="text" <% =BioStatus %> name="fFav_Trick" value= "<% =sFav_Trick %>" maxlength="20" size="20"></td>
	    					</tr>
								<tr>
									<td ALIGN="right" colspan=1><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Jump:</b></font></td>
									<td align="left"><input type="text" <% =BioStatus %> name="fBestJump" value= "<% =sBestJump %>" maxlength="15" size="15"></td>
									<td align="left"><input type="text" <% =BioStatus %> name="fFav_Jump" value= "<% =sFav_Jump %>" maxlength="20" size="20"></td>
	    					</tr>
	    					<tr>
									<td ALIGN="right" colspan=1><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Wakeboard:</b></font></td>
									<td align="left"><input type="text" <% =BioStatus %> name="fBestWake" value= "<% =sBestWake %>" maxlength="15" size="15"></td>
									<td align="left"><input type="text" <% =BioStatus %> name="fFav_Wake" value= "<% =sFav_Wake %>" maxlength="20" size="20"></td>
	    					</tr>
								<tr>
									<td ALIGN="right"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Hydrofoil</b></font></td>
									<td align="left"><input type="text" <% =BioStatus %> name="fBestHydro" value= "<% =sBestHydro %>" maxlength="15" size="15"></td>
									<td align="left">&nbsp</td>
	    					</tr>
	    					<tr>
									<td ALIGN="right"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Freestyle</b></font></td>
									<td align="left"><input type="text" <% =BioStatus %> name="fBestFree" value= "<% =sBestFree %>" maxlength="15" size="15"></td>
									<td align="left">&nbsp</td>
								</tr>
								<tr>
									<td ALIGN="right"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Kneeboard</b></font></td>
									<td align="left"><input type="text" <% =BioStatus %> name="fBestKnee" value= "<% =sBestKnee %>" maxlength="15" size="15"></td>
									<td>&nbsp</td>
								</tr>
								<tr>
	    						<td ALIGN="right"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Marathon</b></font></td>
									<td align="left"><input type="text" <% =BioStatus %> name="fBestMara" value= "<% =sBestMara %>" maxlength="15" size="15"></td>
									<td>&nbsp</td>
								</tr>
							</table>

							<table border=0 Align="center">
		   					<tr>	
		   						<td colspan=4 align="center" style="border-style:none;">
		   							<br>
			  						<input type=submit value="Save" style="width:9em" name="Done">
										<br>
									</td>	
								</tr>
							</table>
						</TD>	
					</TR>
				</TABLE>
			</form>		
			<%

		IF ThisPrint<>"YES" THEN
				'WriteIndexPageFooter
		END IF

END SELECT


' -------------------------------------------------------------------------------
' --------------------------   END OF MAIN PROGRAM  -----------------------------
' -------------------------------------------------------------------------------






' ----------------------
  SUB DisplayDoneWindow
' ----------------------



%>
<br><br><br><br><br><br>
<TABLE class="innertable" border="4" align="center" width="25%">
<TR>
  <TH align="center">
	<FONT size="3" COlOR="#FFFFFF">Press Continue to Return</FONT>
  </TH>
</TR>  
 
<TR>
  <TD align=center>
	<br><br>
	<input type=button value="Continue" onclick="javascript:window.close();">
	<br><br><br>
  </TD>
</TR>  
</TABLE><%

response.end
'style="width:9em" 

END SUB




' ------------------------------------------------------------
  SUB LoadDropDown (ZeroValue, DefaultNum, MinNum, MaxNum, StepNum)
' ------------------------------------------------------------

Dim iCounter

DefaultNum = Cint(DefaultNum)

response.write("<option value = 0 >"&ZeroValue&"</option>")

FOR iCounter = MinNum TO MaxNum STEP StepNum
		IF iCounter = DefaultNum THEN
				response.write("<option value = """&iCounter&""" SELECTED>"&iCounter&"</option>")
		ELSE
				response.write("<option value = """&iCounter&""">"&iCounter&"</option>")
		END IF
NEXT

END SUB



' -----------------
   SUB SaveBio
' -----------------



   OpenCon

   SET rsBio=Server.CreateObject("ADODB.recordset")
   sSQL = "SELECT * FROM "&BioTableName
   sSQL = sSQL + " WHERE MemberID = '"&sMemberID&"'"
   rsBio.open sSQL, SConnectionToTRATable, 3, 3

   IF NOT rsBio.eof	THEN	' --- Found EXISTING record
				OpenCon
				sSQL = "UPDATE "&BioTableName
				sSQL = sSQL + " SET MemberID = '"&sMemberID&"'"
							sSQL = sSQL + " , Region = '"&sRegion&"', Address1 = '"&sAddress1&"', City = '"&sCity&"', State = '"&sState&"', Email = '"&sEmail&"'"
				sSQL = sSQL + " , Phone = '"&sPhone&"', Weight = '"&sWeight&"', HgtFeet = '"&sHgtFeet&"', HgtInch = '"&sHgtInch&"'"

				sSQL = sSQL + " , SkiSinceAge = '"&sSkiSinceAge&"', CompSinceAge = '"&sCompSinceAge&"', MembSinceAge = '"&sMembSinceAge&"'"
				sSQL = sSQL + " , Sponsors = '"&sSponsors&"', Club = '"&sClub&"', School = '"&sSchool&"', Occup = '"&sOccup&"', Career = '"&sCareer&"',Hobby = '"&sHobby&"'"
				sSQL = sSQL + " , Paper = '"&sPaper&"', BestSlal = '"&sBestSlal&"', BestHydro = '"&sBestHydro&"', BestTrick = '"&sBestTrick&"'"
				sSQL = sSQL + " , BestFree = '"&sBestFree&"', BestJump = '"&sBestJump&"', BestKnee = '"&sBestKnee&"', BestWake = '"&sBestWake&"', BestMara = '"&sBestMara&"'"

				sSQL = sSQL + " WHERE MemberID = '"&sMemberID&"'"
				con.execute(sSQL)

				OpenCon
				sSQL = "UPDATE "&BioTableName
				sSQL = sSQL + " SET Fav_Slalom = '"&sFav_Slalom&"', Fav_Jump = '"&sFav_Jump&"', Fav_Trick = '"&sFav_Trick&"', Fav_Wake = '"&sFav_Wake&"'"
				sSQL = sSQL + " , Fav_Boat = '"&sFav_Boat&"', Fav_Sports = '"&sFav_Sports&"', Mentors = '"&sMentors&"', Accomplish = '"&sAccomplish&"'"
				sSQL = sSQL + " , Titles = '"&sTitles&"', Records = '"&sRecords&"'"
				sSQL = sSQL + " , LastUpDate = '"&DATE&"'"

				sSQL = sSQL + " WHERE MemberID = '"&sMemberID&"'"
				
				Session("sSQL") = sSQL
				con.execute(sSQL)

   ELSE			' --- No existing so ADD new record  ---
				OpenCon
				sSQL = "INSERT INTO "&BioTableName
				sSQL = sSQL + " (MemberID, Region, Address1, City, State, Email, Phone, Weight, HgtFeet, HgtInch, SkiSinceage, CompSinceAge, MembSinceAge"
				sSQL = sSQL + ", Sponsors, Club, School, Occup, Career, Hobby, Paper, BestSlal, BestHydro, BestTrick, BestFree, BestJump, BestKnee, BestWake"
				sSQL = sSQL + ", BestMara, Fav_Slalom, Fav_Jump, Fav_Trick, Fav_Wake, Fav_Boat, Fav_Sports, Mentors, Accomplish, Titles, Records"
				sSQL = sSQL + ", LastUpdate)" 

				sSQL = sSQL + " VALUES ('"&sMemberID&"', '"&sRegion&"', '"&sAddress1&"', '"&sCity&"', '"&sState&"', '"&sEmail&"'"
				sSQL = sSQL + " , '"&sPhone&"', '"&sWeight&"', '"&sHgtFeet&"', '"&sHgtInch&"', '"&sSkiSinceAge&"', '"&sCompSinceAge&"', '"&sMembSinceAge&"'"
				sSQL = sSQL + " , '"&sSponsors&"', '"&sClub&"', '"&sSchool&"', '"&sOccup&"', '"&sCareer&"', '"&sHobby&"', '"&sPaper&"', '"&sBestSlal&"', '"&sBestHydro&"'"
				sSQL = sSQL + " , '"&sBestTrick&"', '"&sBestFree&"', '"&sBestJump&"', '"&sBestKnee&"', '"&sBestWake&"', '"&sBestMara&"'"
				sSQL = sSQL + " , '"&sFav_Slalom&"', '"&sFav_Jump&"', '"&sFav_Trick&"', '"&sFav_Wake&"', '"&sFav_Boat&"', '"&sFav_Sports&"', '"&sMentors&"'"
				sSQL = sSQL + " , '"&sAccomplish&"', '"&sTitles&"', '"&sRecords&"'" 
				sSQL = sSQL + " , '"&DATE&"')"
				
				Session("sSQL") = sSQL
				
				con.execute(sSQL)
				
   END IF
   closecon

		sBioDone = "Y"
		Session("sBioDoneText")="Complete"
		Session("sBioDoneTextColor")="blue"


END SUB



' ------------------------
   SUB ReadFromForm
' ------------------------

SFirstName = sqlclean(TRIM(Request("fFirstName")))
SFirstName = sqlclean(TRIM(Request("fFirstName")))
SMembSex = sqlclean(TRIM(Request("fMembSex")))

sAddress1 = sqlclean(TRIM(Request("fAddress1")))
sregion = sqlclean(TRIM(Request("fregion")))
sCity = sqlclean(TRIM(Request("fCity")))
sState = sqlclean(TRIM(Request("fState")))
sEmail = sqlclean(TRIM(Request("fEmail")))
sPhone = sqlclean(TRIM(Request("fPhone")))
sWeight = sqlclean(TRIM(Request("fWeight")))
sHgtFeet = sqlclean(TRIM(Request("fHgtFeet")))
sHgtInch = sqlclean(TRIM(Request("fHgtInch")))
sSkiSinceAge = sqlclean(TRIM(Request("fSkiSinceAge")))
sCompSinceAge = sqlclean(TRIM(Request("fCompSinceAge")))
sMembSinceAge = sqlclean(TRIM(Request("fMembSinceAge")))

IF sWeight = "" THEN sWeight = 0
IF sHgtFeet = "" THEN sHgtFeet = 0
IF sHgtInch = "" THEN sHgtInch = 0
IF sSkiSinceAge = "" THEN sSkiSinceAge = 0
IF sCompSinceAge = "" THEN sCompSinceAge = 0
IF sMembSinceAge = "" THEN sMembSinceAge = 0


sSponsors = sqlclean(TRIM(Request("fSponsors")))
sClub = sqlclean(TRIM(Request("fClub")))
sSchool = sqlclean(TRIM(Request("fSchool")))
sOccup = sqlclean(TRIM(Request("fOccup")))
sCareer = sqlclean(TRIM(Request("fCareer")))
sHobby = sqlclean(TRIM(Request("fHobby")))
sPaper = sqlclean(TRIM(Request("fPaper")))
sBestSlal = sqlclean(TRIM(Request("fBestSlal")))
sBestHydro = sqlclean(TRIM(Request("fBestHydro")))
sBestTrick = sqlclean(TRIM(Request("fBestTrick")))
sBestFree = sqlclean(TRIM(Request("fBestFree")))
sBestJump = sqlclean(TRIM(Request("fBestJump")))
sBestKnee = sqlclean(TRIM(Request("fBestKnee")))
sBestWake = sqlclean(TRIM(Request("fBestWake")))
sBestMara = sqlclean(TRIM(Request("fBestMara")))

sFav_Slalom = sqlclean(TRIM(Request("fFav_Slalom")))
sFav_Jump = sqlclean(TRIM(Request("fFav_Jump")))
sFav_Trick = sqlclean(TRIM(Request("fFav_Trick")))
sFav_Wake = sqlclean(TRIM(Request("fFav_Wake")))
sFav_Boat = sqlclean(TRIM(Request("fFav_Boat")))
sMentors = sqlclean(TRIM(Request("fMentors")))
sAccomplish = sqlclean(TRIM(Request("fAccomplish")))
sFav_Sports = sqlclean(TRIM(Request("fFav_Sports")))
sTitles = sqlclean(TRIM(Request("fTitles")))
sRecords = sqlclean(TRIM(Request("fRecords")))


END SUB



' ------------------------
   SUB ReadFromTable
' ------------------------

'sMemberID = Session("sMemberID")
'sTourID = Session("sTourID")

set rsMemb=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT TOP 1 * FROM "&MemberLiveTableName
sSQL = sSQL + " WHERE PersonID = "&RIGHT(sMemberID,8)
rsMemb.open sSQL, sConnectionToTRATable, 3, 1

sFirstName = rsMemb("FirstName")
sLastName = rsMemb("LastName")
sMembSex = rsMemb("Sex")
'sMembAge = Session("MembAge")

SET rsBio=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT * FROM "&BioTableName
sSQL = sSQL + " WHERE MemberID = '"&sMemberID&"'"
rsBio.open sSQL, SConnectionToTRATable, 3, 3


IF NOT rsBio.eof THEN

		' -- LOGIC changed 2-27-2016 to blank out certain values to keep BIO more current --
		' ----------------------------------------------------------------------------------
		
		' --- When bio is out of date --
				' sBioDone = "N"
				' Session("sBioDoneText")="Out of Date"
				' Session("sBioDoneTextColor")="red"
		
		' --- When bio is complete --
				' sBioDone = "Y"
				' Session("sBioDoneText")="Complete"
				' Session("sBioDoneTextColor")="blue"

		sLastUpDate = rsBio("LastUpdate")
		
		sAddress1 = rsBio("Address1")
		sregion = TRIM(rsBio("region"))
		sCity = rsBio("City")
		sState = rsBio("State")
		sEmail = rsBio("Email")
		sPhone = rsBio("Phone")
		sWeight = rsBio("Weight")
		sHgtFeet = rsBio("HgtFeet")
		sHgtInch = rsBio("HgtInch")
		sSkiSinceAge = rsBio("SkiSinceAge")
		sCompSinceAge = rsBio("CompSinceAge")
		sMembSinceAge = rsBio("MembSinceAge")

		' -- These values are blanked out when form is out-of-date --

' IF sMemberID="000001151" THEN response.write("<br>Session(sBioDoneText) = "&Session("sBioDoneText"))	
		
		IF Session("sBioDoneText")="Out of Date" THEN
				sSponsors = ""
				sClub = ""
				sSchool = ""
				sOccup = ""
				sCareer = ""
				sHobby = ""
				sPaper = ""
				sBestSlal = ""
				sBestHydro = ""
				sBestTrick = ""
				sBestFree = ""
				sBestJump = ""
				sBestKnee = ""
				sBestWake = ""
				sBestMara = ""

				sFav_Slalom = ""
				sFav_Jump = ""
				sFav_Trick = ""
				sFav_Wake = ""
				sFav_Boat = ""
				sMentors = ""
				sAccomplish = ""
				sFav_Sports = ""
				sTitles = ""
				sRecords = ""
				
		ELSE		
				sSponsors = rsBio("Sponsors")
				sClub = rsBio("Club")
				sSchool = rsBio("School")
				sOccup = rsBio("Occup")
				sCareer = rsBio("Career")
				sHobby = rsBio("Hobby")
				sPaper = rsBio("Paper")
				sBestSlal = rsBio("BestSlal")
				sBestHydro = rsBio("BestHydro")
				sBestTrick = rsBio("BestTrick")
				sBestFree = rsBio("BestFree")
				sBestJump = rsBio("BestJump")
				sBestKnee = rsBio("BestKnee")
				sBestWake = rsBio("BestWake")
				sBestMara = rsBio("BestMara")

				sFav_Slalom = rsBio("Fav_Slalom")
				sFav_Jump = rsBio("Fav_Jump")
				sFav_Trick = rsBio("Fav_Trick")
				sFav_Wake = rsBio("Fav_Wake")
				sFav_Boat = rsBio("Fav_Boat")
				sMentors = rsBio("Mentors")
				sAccomplish = rsBio("Accomplish")
				sFav_Sports = rsBio("Fav_Sports")
				sTitles = rsBio("Titles")
				sRecords = rsBio("Records")

		END IF

ELSE
		sAddress1 = ""
		sregion = ""
		sCity = rsMemb("City")
		sState = rsMemb("State")
		sEmail = ""
		sPhone = ""
		sWeight = ""
		sHgtFeet = ""
		sHgtInch = ""
		sSkiSinceAge = ""
		sCompSinceAge = ""
		sMembSinceAge = ""

		sSponsors = ""
		sClub = ""
		sSchool = ""
		sOccup = ""
		sCareer = ""
		sHobby = ""
		sPaper = ""
		sBestSlal = ""
		sBestHydro = ""
		sBestTrick = ""
		sBestFree = ""
		sBestJump = ""
		sBestKnee = ""
		sBestWake = ""
		sBestMara = ""


		sFav_Slalom = ""
		sFav_Jump = ""
		sFav_Trick = ""
		sFav_Wake = ""
		sFav_Boat = ""
		sMentors = ""
		sAccomplish = ""
		sFav_Sports = ""
		sTitles = ""
		sRecords = ""

		sLastUpDate = DATE

END IF

rsMemb.Close
rsBio.Close

END SUB


%>






