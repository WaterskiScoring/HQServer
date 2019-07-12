<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->
<!--#include virtual="/rankings/tools_registration.asp"-->
<%

Dim sTourID, rsQues
Dim QuestionID, AnswerDesc1, AnswerDesc2, AnswerDesc3, AnswerDesc4, AnswerDesc5			
Dim AnswerValue, NumberQuestions, StartQuestionID


ThisPageName="tools_surveys.asp"
sTourID="12S999"
sMemberID="000001151"


DefineTRAStyles 




WriteIndexPageHeader


surveyaction = TRIM(LCASE(Request("surveyaction")))

'response.write("<br>surveyaction="&surveyaction)

SELECT CASE surveyaction
		CASE "continue"
				ReadSurveyFormVariables
				
				Response.write("SURVEY SUBMITTED")
		CASE ELSE
				ListQuestions			
END SELECT





' --------------------
  SUB ListQuestions
' --------------------

	'RegSurveyQuestionsTableName="usawsrank.RegSurveyQuestions"

	' ----  Reads all transactions with matching date/time  ----
	SET rsQues=Server.CreateObject("ADODB.recordset")
	sSQL = "SELECT * FROM "&RegSurveyQuestionsTableName
	sSQL = sSQL + " WHERE LEFT(TourID,6) = '"&sTourID&"'"
	rsQues.open sSQL, SConnectionToTRATable, 3, 3



	TName="2012 Goode National Championships"

		%>
		<form id="surveyform" name="surveyform" method="post" action="<%=ThisFileName%>">

		<TABLE class="innertable" width=90% align=center>
			<th colspan=8 align=center>
				<font face="<%=font1%>" size="<%=fontsize3%>" color="<%=textcolor5%>">Required Survey for <%=TName%></font>
			</th>
		</TR>
		<TR>		
			<td align=left width=50%>
				<font face="<%=font1%>" size="<%=fontsize2%>" color="<%=textcolor1%>"><%=QuestionDesc%></font>
			</td>
		</TR>
		<%

		NumberQuestions = 0
		StartQuestionID=""
		DO WHILE NOT rsQues.eof
				QuestionID=rsQues("QuestionID")
				IF StartQuestionID="" THEN StartQuestionID=QuestionID		
				'response.write("<br>QuestionID= "&QuestionID)
				DisplayQuestions
			
				rsQues.movenext
		LOOP
	
		%>
		<TR>
			<td colspan=8 align=center>
					<input type="submit" name="surveyaction" value="Continue">
			</td>
		</TR>	
		</TABLE>

		<input type="hidden" name="NumberQuestions" value="<%=NumberQuestions%>">
		<input type="hidden" name="StartQuestionID" value="<%=StartQuestionID%>">


	</form>
	<%

END SUB


' -----------------------------
  SUB ReadSurveyFormVariables
' -----------------------------  
  
  'RegSurveyAnswersTableName = "usawsrank.RegSurveyAnswers"

	NumberQuestions = Request("NumberQuestions")
	StartQuestionID = Request("StartQuestionID")
	
	LastQuestionID = Cint(StartQuestionID) + CInt(NumberQuestions)-CInt(1)
	
	sSQL = "SELECT MemberID FROM "&RegSurveyAnswersTableName
	sSQL = sSQL + " WHERE LEFT(TourID,6) = '"&sTourID&"' AND MemberID = '"&sMemberID&"'"
	
	SET rs=Server.CreateObject("ADODB.recordset")
	rs.open sSQL, SConnectionToTRATable, 3, 3
	
'	Response.write("<br> "&sSQL)
'	Response.write("<br>Found = ")
'	Response.write(NOT(rs.eof))
	
	IF rs.eof THEN
			OpenCon

			FOR ThisQuestionID=StartQuestionID to LastQuestionID
					ThisAnswer = TRIM(Request(ThisQuestionID&"_Answer"))
	
					sSQL = "INSERT INTO "&RegSurveyAnswersTableName
					sSQL = sSQL + " VALUES ('"&sTourID&"','"&sMemberID&"','"&ThisQuestionID&"','"&ThisAnswer&"')"
					'response.write("<br>"&sSQL)
					con.execute(sSQL)
			NEXT

			CloseCon
	END IF




END SUB


' ------------------------
  SUB DisplayQuestions
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
		<td>
			<font face="<%=font1%>" size="<%=fontsize2%>" color="<%=textcolor1%>"><%=QuestionDesc%></font>
		</td>
		<%
			
	AnswerValue=QuestionID&"_Answer"
	
	SELECT CASE AnswerType
			CASE "TXT"
						%>	
						<td align=left colspan=7 width=50%>
							<input type=text name=<%=AnswerValue%> value="" MaxLength=40 size="40">					
						</td>
						<%	
					NumberQuestions = NumberQuestions +1
			CASE "RAD"
						%>	
						<td align=left>
							<font face="<%=font1%>" size="<%=fontsize2%>" color="<%=textcolor1%>">
					    	<%=AnswerDesc1%>
					    	<input type="radio" name="<%=AnswerValue%>" value="<%=AnswerDesc1%>">
							</font>
						</td>
						<td align=left>
							<font face="<%=font1%>" size="<%=fontsize2%>" color="<%=textcolor1%>">
					    	<%=AnswerDesc2%>
    						<input type="radio" name="<%=AnswerValue%>" value="<%=AnswerDesc2%>">
							</font>
						</td>
 						<%

  					IF TRIM(AnswerDesc3)<>"" THEN
   								%>
									<td align=left>
										<font face="<%=font1%>" size="<%=fontsize2%>" color="<%=textcolor1%>">
						  		  	<%=AnswerDesc3%>
   										<input type="radio" name="<%=AnswerValue%>" value="<%=AnswerDesc3%>">
										</font>
									</td>
  								<%
						END IF
  					IF TRIM(AnswerDesc4)<>"" THEN
  								%>
									<td align=left>
										<font face="<%=font1%>" size="<%=fontsize2%>" color="<%=textcolor1%>">
						  		  	<%=AnswerDesc4%>
    									<input type="radio" name="<%=AnswerValue%>" value="<%=AnswerDesc4%>">
										</font>
									</td>
   								<%
						END IF
    				IF TRIM(AnswerDesc5)<>"" THEN
    							%>
									<td align=left>
										<font face="<%=font1%>" size="<%=fontsize2%>" color="<%=textcolor1%>">
						  		  	<%=AnswerDesc5%>
    									<input type="radio" name="<%=AnswerValue%>" value="<%=AnswerDesc5%>">
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








' --------------------------------------------------------------------------------------------
   SUB BuildDropRentalBrand 
' --------------------------------------------------------------------------------------------

	%>
	<select id='<%=AnswerValue%>' name='<%=AnswerValue%>' style="width:8em" >
		<option value ='' <%IF AnswerDesc1 = "" THEN Response.Write(" selected ")%>>Select</Option><br>
		<option value ='Alamo' <%IF AnswerDesc1 = "Alamo" THEN Response.Write(" selected ")%>>Alamo</Option><br>
		<option value ='Avis' <%IF AnswerDesc1 = "Avis" THEN Response.Write(" selected ")%>>Avis</Option><br>
		<option value ='Budget' <%IF AnswerDesc1 = "Budget" THEN Response.Write(" selected ")%>>Budget</Option><br>
		<option value ='Dollar' <%IF AnswerDesc1 = "Dollar" THEN Response.Write(" selected ")%>>Dollar</Option><br>		
		<option value ='Enterprise' <%IF AnswerDesc1 = "Enterprise" THEN Response.Write(" selected ")%>>Enterprise</Option><br>		
		<option value ='Hertz' <%IF AnswerDesc1 = "Hertz" THEN Response.Write(" selected ")%>>Hertz</Option><br>
		<option value ='Other' <%IF AnswerDesc1 = "Other" THEN Response.Write(" selected ")%>>Other</Option><br>
		<option value ='Unknown' <%IF AnswerDesc1 = "Unknown" THEN Response.Write(" selected ")%>>Unknown</Option><br>
	</select>
	<%
END SUB


' --------------------------------------------------------------------------------------------
   SUB BuildDropQuantity
' --------------------------------------------------------------------------------------------

	%>
	<select id='<%=AnswerValue%>' name='<%=AnswerValue%>' style="width:8em" >
		<option value ='' <%IF AnswerDesc1 = "" THEN Response.Write(" selected ")%>>Select</Option><br>
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



%>

