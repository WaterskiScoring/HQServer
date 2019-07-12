<%



	Dim rsQues
	Dim QuestionID, AnswerDesc1, AnswerDesc2, AnswerDesc3, AnswerDesc4, AnswerDesc5			
	Dim AnswerValue, NumberQuestions, StartQuestionID, sSurveyThankYouText

' ***************************************************************************************************************************
' ***************************************************************************************************************************
' --- This program operates the survey which can be embedded in the Registration program by specifying certain parameters ---
' ***************************************************************************************************************************
' ***************************************************************************************************************************

	


' ---------------------------
  SUB BeginTournamentSurvey
' ---------------------------


	ThisPageName="Register_Survey.asp"

	sSurveyThankYouText = "Thank you for taking the time to complete this survey. This assists Palm Beach County in understanding the benefits of hosting water skiing at our great facilities at Okeeheelee Park." 


	surveyaction = TRIM(LCASE(Request("surveyaction")))

	' --- Writes th Javascript code ---
	CreateJavaSection

	SELECT CASE surveyaction
		CASE "continue"
				SaveSurveyFormVariables
				
				ConfirmSurveyComplete

		CASE ELSE
				ListSurveyQuestions			

	END SELECT


END SUB





' -------------------------
  SUB ListSurveyQuestions
' -------------------------



	' ----  Reads all transactions with matching date/time  ----
	SET rsQues=Server.CreateObject("ADODB.recordset")
	sSQL = "SELECT * FROM "&RegSurveyQuestionsTableName
	sSQL = sSQL + " WHERE LEFT(TourID,6) = '"&sTourID&"'"
	rsQues.open sSQL, SConnectionToTRATable, 3, 3

'response.write(sSQL)
'response.write("Found=")
'response.write(NOT(rsQues.eof))

		%>
		<form id="surveyform" name="surveyform" method="post" action="/rankings/<%=RegFileName%>">
			<input type="hidden" name="nav" value="5">
			<input type="hidden" name="waivernav" value="6">

		

		<TABLE class="innertable" width=90% align=center>
			<TR>
				<th colspan=8 align=center>
					<font face="<%=font1%>" size="<%=fontsize3%>" color="<%=textcolor5%>">Required Survey for <%=sTourName%></font>
				</th>
			</TR>
			<TR>		
				<td colspan=8>&nbsp;</td>
			</TR>
			<%


			NumberQuestions = 0
			StartQuestionID=""
			DO WHILE NOT rsQues.eof
					QuestionID=rsQues("QuestionID")
					IF StartQuestionID="" THEN StartQuestionID=QuestionID		
					DisplaySurveyQuestions
			
					rsQues.movenext
			LOOP
	
			%>
			<TR>
				<td colspan=8 align=center>
						<br>
						<input type="submit" name="surveyaction" value="Continue" style="width:9em">
						<br><br>
				</td>
			</TR>	
			</TABLE>

			<input type="hidden" name="NumberQuestions" value="<%=NumberQuestions%>">
			<input type="hidden" name="StartQuestionID" value="<%=StartQuestionID%>">


		</form>
		<%

END SUB




' -----------------------------
  SUB SaveSurveyFormVariables
' -----------------------------  
  

	NumberQuestions = Request("NumberQuestions")
	StartQuestionID = Request("StartQuestionID")
	
	IF TRIM(NumberQuestions)="" THEN NumberQuestions=8
	IF TRIM(StartQuestionID)="" THEN StartQuestionID=1
			 
	LastQuestionID = Cint(StartQuestionID) + CInt(NumberQuestions)-CInt(1)
	
	sSQL = "SELECT MemberID FROM "&RegSurveyAnswersTableName
	sSQL = sSQL + " WHERE LEFT(TourID,6) = '"&sTourID&"' AND MemberID = '"&sMemberID&"'"
	
	SET rs=Server.CreateObject("ADODB.recordset")
	rs.open sSQL, SConnectionToTRATable, 3, 3
	
	IF rs.eof THEN
			OpenCon

			FOR ThisQuestionID=StartQuestionID to LastQuestionID
					ThisAnswer = TRIM(Request(ThisQuestionID&"_Answer"))
					IF ThisQuestionID=4 AND TRIM(ThisAnswer)="" THEN ThisAnswer = "No"
					IF ThisQuestionID=7 AND TRIM(ThisAnswer)="" THEN ThisAnswer = "Car"
	
					sSQL = "INSERT INTO "&RegSurveyAnswersTableName
					sSQL = sSQL + " VALUES ('"&sTourID&"','"&sMemberID&"','"&ThisQuestionID&"','"&ThisAnswer&"')"
					con.execute(sSQL)
			NEXT

			CloseCon
	END IF

END SUB



' ------------------------
  SUB DisplaySurveyQuestions
' ------------------------

	
	AnswerType=rsQues("AnswerType")
	AnswerQuantity=rsQues("AnswerQuantity")	
	AnswerCode=TRIM(rsQues("AnswerCode"))
	QuestionID=rsQues("QuestionID")
	QuestionDesc=TRIM(rsQues("QuestionDesc"))
	AnswerDesc1=TRIM(rsQues("AnswerDesc1"))
	AnswerDesc2=TRIM(rsQues("AnswerDesc2"))
	AnswerDesc3=TRIM(rsQues("AnswerDesc3"))
	AnswerDesc4=TRIM(rsQues("AnswerDesc4"))
	AnswerDesc5=TRIM(rsQues("AnswerDesc5"))
	
		

	%>	
	<TR>
		<td align=right height="30px">
			<font face="<%=font1%>" size="<%=fontsize2%>" color="<%=textcolor1%>"><%=QuestionDesc%>&nbsp&nbsp</font>
		</td>
		<%
			
	AnswerValue=QuestionID&"_Answer"
	
	SELECT CASE AnswerType
			CASE "TXT"
						OrigAnswer = ""
						IF TRIM(AnswerCode)="OrigCity" THEN OrigAnswer = sMembCity & ", "&sMembState

						%>	
						<td align=left colspan=7 width=50%>
							<input type=text id="<%=AnswerCode%>" name=<%=AnswerValue%> value="<%=OrigAnswer%>" MaxLength=40 size="40">					
						</td>
						<%	
					NumberQuestions = NumberQuestions +1
			CASE "RAD"

						%>	
						<td align=left>
							<font face="<%=font1%>" size="<%=fontsize2%>" color="<%=textcolor2%>">
					    	<%=AnswerDesc1%>
					    	<input type="radio" id="<%=AnswerValue%>_1" name="<%=AnswerValue%>" value="<%=AnswerDesc1%>">
							</font>
						</td>
						<td align=left>
							<font face="<%=font1%>" size="<%=fontsize2%>" color="<%=textcolor2%>">
					    	<%=AnswerDesc2%>
    						<input type="radio" id="<%=AnswerValue%>_2" name="<%=AnswerValue%>" value="<%=AnswerDesc2%>">
							</font>
						</td>
 						<%

  					IF TRIM(AnswerDesc3)<>"" THEN
   								%>
									<td align=left>
										<font face="<%=font1%>" size="<%=fontsize2%>" color="<%=textcolor2%>">
						  		  	<%=AnswerDesc3%>
   										<input type="radio" id="<%=AnswerValue%>_3" name="<%=AnswerValue%>" value="<%=AnswerDesc3%>">
										</font>
									</td>
  								<%
						END IF
  					IF TRIM(AnswerDesc4)<>"" THEN
  								%>
									<td align=left>
										<font face="<%=font1%>" size="<%=fontsize2%>" color="<%=textcolor2%>">
						  		  	<%=AnswerDesc4%>
    									<input type="radio" id="<%=AnswerValue%>_4" name="<%=AnswerValue%>" value="<%=AnswerDesc4%>">
										</font>
									</td>
   								<%
						END IF
    				IF TRIM(AnswerDesc5)<>"" THEN
    							%>
									<td align=left>
										<font face="<%=font1%>" size="<%=fontsize2%>" color="<%=textcolor2%>">
						  		  	<%=AnswerDesc5%>
    									<input type="radio" id="<%=AnswerValue%>_5" name="<%=AnswerValue%>" value="<%=AnswerDesc5%>">
										</font>
									</td>
   								<%
						END IF
		
						%>
						<td colspan=<%=7-AnswerQuantity%>>&nbsp;</td>
						<%	
  					NumberQuestions = NumberQuestions +1
  								
			CASE "STD"
					%>
					<td colspan=7 align=left>
						<%
						SELECT CASE AnswerCode
								CASE "RentalBrand"
										BuildDropRentalBrand
								CASE "NumberNights"
										BuildDropQuantity
								CASE "NumberRooms"
										BuildRoomCount		
								CASE "RegionCode"
										BuildDropRegionCode 
								CASE "DistTravel"
										
										BuildDropDistMiles
								CASE "HotelName"
										BuildDropHotelName								
						END SELECT
						%>
					</td>
					<%
			  	NumberQuestions = NumberQuestions +1
			
	END SELECT

		%>
		</TR>
		<%

END SUB



' ----------------------------
  SUB CreateJavaSection
' ---------------------------- 


%>
<script type="text/javascript">

	function DisableNonLocal () 
		{

		//document.write("HERE");
			// -----------------------------------------------
			// --- Disable options if participant is local ---
			// -----------------------------------------------
		
			if (document.getElementById('DistTravel').value == "Local")
				{
					document.getElementById('HotelName').value="None";
					document.getElementById('NumberNights').value="0";
					document.getElementById('RentalBrand').value="None";
					document.getElementById('HotelName').disabled=true;
					document.getElementById('4_Answer_1').disabled=true;
					document.getElementById('4_Answer_2').disabled=true;
					document.getElementById('NumberNights').disabled=true;
					document.getElementById('7_Answer_1').disabled=true;
					document.getElementById('7_Answer_2').disabled=true;
					document.getElementById('7_Answer_3').disabled=true;
					document.getElementById('RentalBrand').disabled=true;
										
				}
			else
				{
					document.getElementById('HotelName').value="";
					document.getElementById('NumberNights').value="";
					document.getElementById('RentalBrand').value="";
					document.getElementById('HotelName').disabled=false;
					document.getElementById('4_Answer_1').disabled=false;
					document.getElementById('4_Answer_2').disabled=false;
					document.getElementById('NumberNights').disabled=false;
					document.getElementById('7_Answer_1').disabled=false;
					document.getElementById('7_Answer_2').disabled=false;
					document.getElementById('7_Answer_3').disabled=false;
					document.getElementById('RentalBrand').disabled=false;

				}
				
		}

</script>
<%

END SUB


' ---------------------------
  SUB ConfirmSurveyComplete
' ---------------------------

		%>
		<form id="surveyformcomplete" name="surveyformcomplete" method="post" action="/rankings/<%=RegFileName%>">
	  	<input type="hidden" name="nav" value=6>

		<TABLE class="innertable" width=90% align=center>
			<TR>
				<th colspan=8 align=center>
					<font face="<%=font1%>" size="<%=fontsize3%>" color="<%=textcolor5%>">Survey Complete</font>
				</th>
			</TR>
			<TR>		
				<td align=center>
					<br>
					<font face="<%=font1%>" size="<%=fontsize3%>" color="<%=textcolor1%>">&nbsp;<%=sSurveyThankYouText%>&nbsp;</font>
				</td>
			</TR>
			<TR>
	  		<td align="center">
					<br><br>
	  			<input type="submit" name="SurveyContinue" value="Continue To Payment" style="width:12em" title="Continue to Payment Page">
	  			<br><br>
	  		</td>
	  	</TR>
	  </TABLE>
			
		</form>

		<%



END SUB

					
	
' --------------------------
   SUB BuildDropRentalBrand 
' --------------------------

	%>
	<select id='RentalBrand' name='<%=AnswerValue%>' style="width:8em" >
		<option value ='' <%IF AnswerDesc1 = "" THEN Response.Write(" selected ")%>>Select</Option><br>
		<option value ='Alamo' <%IF AnswerDesc1 = "Alamo" THEN Response.Write(" selected ")%>>Alamo</Option><br>
		<option value ='Avis' <%IF AnswerDesc1 = "Avis" THEN Response.Write(" selected ")%>>Avis</Option><br>
		<option value ='Budget' <%IF AnswerDesc1 = "Budget" THEN Response.Write(" selected ")%>>Budget</Option><br>
		<option value ='Dollar' <%IF AnswerDesc1 = "Dollar" THEN Response.Write(" selected ")%>>Dollar</Option><br>		
		<option value ='Enterprise' <%IF AnswerDesc1 = "Enterprise" THEN Response.Write(" selected ")%>>Enterprise</Option><br>		
		<option value ='Hertz' <%IF AnswerDesc1 = "Hertz" THEN Response.Write(" selected ")%>>Hertz</Option><br>
		<option value ='Other' <%IF AnswerDesc1 = "Other" THEN Response.Write(" selected ")%>>Other</Option><br>
		<option value ='Unknown' <%IF AnswerDesc1 = "Unknown" THEN Response.Write(" selected ")%>>Unknown</Option><br>
		<option value ='None' <%IF AnswerDesc1 = "None" THEN Response.Write(" selected ")%>>None</Option><br>
	</select>
	<%
END SUB


' --------------------------
   SUB BuildDropHotelName 
' --------------------------

	%>
	<select id='HotelName' name='<%=AnswerValue%>' style="width:10em" >
		<option value ='' <%IF AnswerDesc1 = "" THEN Response.Write(" selected ")%>>Select</Option><br>
		<option value ='None' <%IF AnswerDesc1 = "None" THEN Response.Write(" selected ")%>>None</Option><br>
		<option value ='Embassy Suites' <%IF AnswerDesc1 = "Embassy Suites" THEN Response.Write(" selected ")%>>Embassy Suites</Option><br>
		<option value ='AirBnb' <%IF AnswerDesc1 = "AirBnb" THEN Response.Write(" selected ")%>>AirBnb</Option><br>
		<option value ='Best Western' <%IF AnswerDesc1 = "Best Western" THEN Response.Write(" selected ")%>>Best Western</Option><br>
		<option value ='Comfort Inn' <%IF AnswerDesc1 = "Comfort Inn" THEN Response.Write(" selected ")%>>Comfort Inn</Option><br>		
		<option value ='Courtyard' <%IF AnswerDesc1 = "Courtyard" THEN Response.Write(" selected ")%>>Courtyard</Option><br>
		<option value ='Days Inn' <%IF AnswerDesc1 = "Days Inn" THEN Response.Write(" selected ")%>>Days Inn</Option><br>
		<option value ='DoubleTree' <%IF AnswerDesc1 = "DoubleTree" THEN Response.Write(" selected ")%>>DoubleTree</Option><br>
		<option value ='Econo Lodge' <%IF AnswerDesc1 = "Econo Lodge" THEN Response.Write(" selected ")%>>Econo Lodge</Option><br>		
		<option value ='Fairfield Inn' <%IF AnswerDesc1 = "Fairfield Inn" THEN Response.Write(" selected ")%>>Fairfield Inn</Option><br>		
		<option value ='Hampton Inn' <%IF AnswerDesc1 = "Hampton Inn" THEN Response.Write(" selected ")%>>Hampton Inn</Option><br>		
		<option value ='Holiday Inn' <%IF AnswerDesc1 = "Holiday Inn" THEN Response.Write(" selected ")%>>Holiday Inn</Option><br>		
		<option value ='Howard Johnson' <%IF AnswerDesc1 = "Howard Johnson" THEN Response.Write(" selected ")%>>Howard Johnsonr</Option><br>		
		<option value ='La Quinta' <%IF AnswerDesc1 = "La Quinta" THEN Response.Write(" selected ")%>>La Quinta</Option><br>		
		<option value ='Motel 6' <%IF AnswerDesc1 = "Motel 6" THEN Response.Write(" selected ")%>>Motel 6</Option><br>		
		<option value ='Marriott' <%IF AnswerDesc1 = "Marriott" THEN Response.Write(" selected ")%>>Marriott</Option><br>		
		<option value ='Quality' <%IF AnswerDesc1 = "Quality" THEN Response.Write(" selected ")%>>Quality</Option><br>		
		<option value ='Ramada' <%IF AnswerDesc1 = "Ramada" THEN Response.Write(" selected ")%>>Ramada</Option><br>		
		<option value ='Residence' <%IF AnswerDesc1 = "Residence" THEN Response.Write(" selected ")%>>Residence Inn</Option><br>		
		<option value ='Super 8' <%IF AnswerDesc1 = "Super 8" THEN Response.Write(" selected ")%>>Super 8</Option><br>		
		<option value ='Travel Lodge' <%IF AnswerDesc1 = "Travel Lodge" THEN Response.Write(" selected ")%>>Travel Lodge</Option><br>		
		<option value ='Other' <%IF AnswerDesc1 = "Other" THEN Response.Write(" selected ")%>>Other</Option><br>
		<option value ='Unknown' <%IF AnswerDesc1 = "Unknown" THEN Response.Write(" selected ")%>>Unknown</Option><br>
	</select>
	<%
END SUB


' --------------------------
   SUB BuildDropQuantity
' --------------------------

	%>
	<select id='NumberNights' name='<%=AnswerValue%>' style="width:8em" >
		<option value ='' <%IF AnswerDesc1 = "" THEN Response.Write(" selected ")%>>Select</Option><br>
		<option value ='0' <%IF AnswerDesc1 = "0" THEN Response.Write(" selected ")%>>0</Option><br>
		<option value ='1' <%IF AnswerDesc1 = "1" THEN Response.Write(" selected ")%>>1</Option><br>
		<option value ='2' <%IF AnswerDesc1 = "2" THEN Response.Write(" selected ")%>>2</Option><br>
		<option value ='3' <%IF AnswerDesc1 = "3" THEN Response.Write(" selected ")%>>3</Option><br>
		<option value ='4' <%IF AnswerDesc1 = "4" THEN Response.Write(" selected ")%>>4</Option><br>		
		<option value ='5' <%IF AnswerDesc1 = "5" THEN Response.Write(" selected ")%>>5</Option><br>		
		<option value ='6' <%IF AnswerDesc1 = "6" THEN Response.Write(" selected ")%>>6</Option><br>
		<option value ='7' <%IF AnswerDesc1 = "7" THEN Response.Write(" selected ")%>>7</Option><br>
		<option value ='8' <%IF AnswerDesc1 = "8" THEN Response.Write(" selected ")%>>8+</Option><br>
	</select>
	<%
END SUB



' --------------------------
   SUB BuildRoomCount
' --------------------------

	%>
	<select id='RoomCount' name='<%=AnswerValue%>' style="width:8em" >
		<option value ='' <%IF AnswerDesc1 = "" THEN Response.Write(" selected ")%>>Select</Option><br>
		<option value ='0' <%IF AnswerDesc1 = "0" THEN Response.Write(" selected ")%>>0</Option><br>
		<option value ='1' <%IF AnswerDesc1 = "1" THEN Response.Write(" selected ")%>>1</Option><br>
		<option value ='2' <%IF AnswerDesc1 = "2" THEN Response.Write(" selected ")%>>2</Option><br>
		<option value ='3' <%IF AnswerDesc1 = "3" THEN Response.Write(" selected ")%>>3</Option><br>
		<option value ='4' <%IF AnswerDesc1 = "4" THEN Response.Write(" selected ")%>>4+</Option><br>		
	</select>
	<%
END SUB



' --------------------------
   SUB BuildDropRegionCode 
' --------------------------

	%>
	<select id='RegionCode' name='<%=AnswerValue%>' style="width:8em" >
		<option value ='' <%IF AnswerDesc1 = "" THEN Response.Write(" selected ")%>>Select</Option><br>
		<option value ='S' <%IF AnswerDesc1 = "S" THEN Response.Write(" selected ")%>>SCentral</Option><br>
		<option value ='M' <%IF AnswerDesc1 = "M" THEN Response.Write(" selected ")%>>Midwest</Option><br>
		<option value ='S' <%IF AnswerDesc1 = "S" THEN Response.Write(" selected ")%>>South</Option><br>
		<option value ='W' <%IF AnswerDesc1 = "W" THEN Response.Write(" selected ")%>>West</Option><br>		
		<option value ='E' <%IF AnswerDesc1 = "E" THEN Response.Write(" selected ")%>>East</Option><br>		
	</select>
	<%
END SUB


' --------------------------
   SUB BuildDropDistMiles 
' --------------------------

	%>
	<select id='DistTravel' name='<%=AnswerValue%>' style="width:8em" onchange='javascript:DisableNonLocal();'>
		<option value ='' <%IF AnswerDesc1 = "" THEN Response.Write(" selected ")%>>Select</Option><br>
		<option value ='Local' <%IF AnswerDesc1 = "Local" THEN Response.Write(" selected ")%>>Local</Option><br>
		<option value ='0-100' <%IF AnswerDesc1 = "0-100" THEN Response.Write(" selected ")%>>Under 100</Option><br>
		<option value ='101-250' <%IF AnswerDesc1 = "101-250" THEN Response.Write(" selected ")%>>101 to 250</Option><br>
		<option value ='251-500' <%IF AnswerDesc1 = "251-500" THEN Response.Write(" selected ")%>>251 to 500</Option><br>
		<option value ='501-1000' <%IF AnswerDesc1 = "501-1000" THEN Response.Write(" selected ")%>>501 to 1000</Option><br>		
		<option value ='1001+' <%IF AnswerDesc1 = "1001+" THEN Response.Write(" selected ")%>>More than 1000</Option><br>		
	</select>
	<%
END SUB


%>

