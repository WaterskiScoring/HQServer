<%
' ----------------------------------------------------------------------------------------------------------------
' --- FILE:  Bio-Form-Print.asp 
'
' --- This is used only as a virtual include ---
' ----------------------------------------------------------------------------------------------------------------






' ----------------------------------------------------------------------
   SUB DisplayBioForm  (sMemberID, sTourID, sDiv1, sDiv2, sDiv3, sDiv4)
' ----------------------------------------------------------------------


Dim sAddress1, sregion, sCity, sState, sEmail, sPhone, sWeight, sHgtFeet, sHgtInch, sSkiSinceAge, sCompSinceAge, sMembSinceAge, sSponsors
Dim sClub, sSchool, sOccup, sCareer, sHobby, sPaper, sBestSlal, sBestHydro, sBestTrick, sBestFree, sBestJump, sBestKnee, sBestWake, sBestMara
Dim sRunByWhat, FormStatus, BioStatus
Dim sFav_Slalom, sFav_Jump, sFav_Trick, sFav_Wake, sFav_Boat, sMentors, sAccomplish, sFav_Sports, sTitles, sRecords


' --------------------------
' --- Sets TABLE designs ---
' --------------------------
DefineTRAStyles 



' ----------------------------
' --- Get Tour information ---
' ----------------------------
set rsTour=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT * FROM "&SanctionTABLEName
sSQL = sSQL + " WHERE TournAppID = '"&LEFT(sTourID,6)&"'"
rsTour.open sSQL, sConnectionToTRATABLE, 3, 1
IF NOT rsTour.eof THEN
		sTName=rsTour("TName")
END IF


' ------------------------------
' --- Get Member information ---
' ------------------------------
set rsMemb=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT TOP 1 * FROM "&MemberTABLEName
sSQL = sSQL + " WHERE PersonIDwithCheckDigit = "&sqlclean(sMemberID)
rsMemb.open sSQL, sConnectionToTRATABLE, 3, 1

sFirstName = rsMemb("FirstName")
sLastName = rsMemb("LastName")
sMembSex = rsMemb("Sex")
'sMembAge = Session("MembAge")
sMembAge = AgeAtDate_New(Date, sMemberID)

' --- Get Bio information ---
SET rsBio=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT * FROM "&BioTABLEName
sSQL = sSQL + " WHERE MemberID = '"&sMemberID&"'"
rsBio.open sSQL, SConnectionToTRATABLE, 3, 3


' -----------------------------
' --- Found EXISTING record ---
' -----------------------------
IF NOT rsBio.eof THEN	

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
		sLastUpDate = rsBio("LastUpdate")

		Dim ThisEventName1, ThisEventName2, ThisEventName3, ThisDiv1, ThisDiv2, ThisDiv3, ThisHomeRegion
		
		ThisEventName1 = sTEventName(1)
		ThisEventName2 = sTEventName(2)
		ThisEventName3 = sTEventName(3)						
				
		ThisDiv1 = "***"
		ThisDiv2 = "***"
		ThisDiv3 = "***"
		IF TRIM(sDiv(1))<>"" THEN ThisDiv1 = TRIM(sDiv(1))	
		IF TRIM(sDiv(2))<>"" THEN ThisDiv2 = TRIM(sDiv(2))	
		IF TRIM(sDiv(3))<>"" THEN ThisDiv3 = TRIM(sDiv(3))	
		
		ThisHomeRegion="Unknown"
		SELECT CASE sregion
				CASE "1"
						ThisHomeRegion="S Central"
				CASE "2"
						ThisHomeRegion="Midwest"
				CASE "3"
						ThisHomeRegion="West"
				CASE "4"
						ThisHomeRegion="Southern"
				CASE "5"
						ThisHomeRegion="East"
				CASE "6"
						ThisHomeRegion="Foreign"
		END SELECT


'response.write("<br>font1 = "&font1)
'response.write("<br>fontsize2 = "&fontsize2)

	%>

	<HTML>
		<HEAD>
			<style>div.break {page-break-before:always}</style>
		</HEAD>
		<BODY>
		<TABLE class="innertable" align="center" width=100%>   	
      	<tr> 
	    		<td align="center">
						<font size="<%=fontsize4%>" face="<%=font1%>" color="red"><b><%=sTName%></b></font>
						<br>
						<font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>"><%=sTDateS%>-<%=sTDateE%></font>
						<br><br>
						<font size="<%=fontsize4%>" face="<%=font1%>" color="<% =textcolor1 %>"><b>Division:
						<font size="<%=fontsize4%>" face="<%=font1%>" color="<%=textcolor2%>">&nbsp;<%=DivSelected%>&nbsp;&nbsp;&nbsp;&nbsp;
						<font size="<%=fontsize4%>" face="<%=font1%>" color="<% =textcolor1 %>">Event:
						<font size="<%=fontsize4%>" face="<%=font1%>" color="<%=textcolor2%>">&nbsp;<%=EventSelected%></b></font>
						<br><br>
	    		</td>
				</tr>  
		</TABLE>

    <TABLE class="innertable" align="center" width=100%>
			<tr>
  			<th colspan="8" align="center"> <font size="3" face="<%=font1%>" color="#FFFFFF"><b>Personal Information</b></font></th>
			</tr>
			<tr>  
	    	<td align="right" width="10%"><font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>"><b>Name:</b>&nbsp;&nbsp</font></td>
	    	<td align="left" width="15%"><font size="<% =fontsize3 %>" face="<%=font1%>" color="<%=textcolor2%>"><% =sFirstName %>&nbsp;<% =sLastName %></font></td>
   	    <td align="right" width="10%"><font size="<%=fontsize2%>" face="<%=font1%>" color="<%=textcolor1%>"><b>Age/Gender:</b>&nbsp;&nbsp</font></td>
		    <td align="left" width="15%"><font size="<%=fontsize3%>" face="<%=font1%>" color="<%=textcolor2%>"><%=sMembAge%>/<%=sMembSex%></font></td> 
			</tr>
			<tr>  
	    	<td align="right" width="10%"><font size="<%=fontsize2%>" face="<%=font1%>" color="<%=textcolor4%>">&nbsp;&nbsp;&nbsp;<b>City/ST:</b></font></td>
	    	<td align="left" width="15%"><font size="<% =fontsize3 %>" face="<%=font1%>" color="<%=textcolor2%>"><%=sCity%>,&nbsp;<%=sState%></font></td>
		    <td align="right" width="10%"><font size="<%=fontsize2%>" face="<%=font1%>"><b>Home Region:</b></font></td> 
		    <td align="left" width="15%"><font size="<% =fontsize3 %>" face="<%=font1%>" color="<%=textcolor2%>"><%=ThisHomeRegion%></font></td>
		  </tr>

		  <tr>  
		    <td align="right"><font size="<%=fontsize2%>" face="<%=font1%>" color="<%=textcolor4%>">&nbsp;&nbsp;&nbsp;<b>Weight:</b></font></td>
				<td align="left"><font size="<% =fontsize3 %>" face="<%=font1%>" color="<%=textcolor2%>"><%=sWeight%></font></td>
		    <td align="right"><font size="<%=fontsize2%>" face="<%=font1%>" color="<%=textcolor4%>">&nbsp;&nbsp;&nbsp;<b>Height:</b></font></td>
		    <td>
					<font size="<% =fontsize3 %>" face="<%=font1%>" color="<%=textcolor2%>"><%=sHgtFeet%></font>
					<font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>">ft&nbsp;&nbsp;&nbsp;</font>
	     		<font size="<% =fontsize3 %>" face="<%=font1%>" color="<%=textcolor2%>"><%=sHgtInch%></font>
					<font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>">in</font>
	    	</td>
			</tr>
			<tr>  
	    	<td colspan=2>&nbsp</td>
	    	<td colspan=2>&nbsp</td>
	  	</tr>

			<tr>
				<td align="right" ><font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>"><b>Ski Club:</b></font></td>
				<td align="left"><font size="<% =fontsize3 %>" face="<%=font1%>" color="<%=textcolor2%>"><%=sClub%></font></td>
				<td align="right"><font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>"><b>Skiing Since:</b></font></td>
				<td align="left">
					<font size="<% =fontsize3 %>" face="<%=font1%>" color="<%=textcolor2%>"><%=sSkiSinceAge%></font>
					<font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>">years old</font>
				</td>
			</tr>

			<tr>  
				<td align="right"><font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>"><b>Occupation:</b></font></td>
				<td align="left"><font size="<% =fontsize3 %>" face="<%=font1%>" color="<%=textcolor2%>"><%=sOccup%></font></td>
				<td align="right"><font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>"><b>Competing Since:</b></font></td>
				<td align="left">
					<font size="<% =fontsize3 %>" face="<%=font1%>" color="<%=textcolor2%>"><%=sCompSinceAge%></font>
					<font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>">years old</font>
				</td>
			</tr>

			<tr>  
				<td align="right"><font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>"><b>Career Plans:</b></font></td>
				<td align="left"><font size="<% =fontsize3 %>" face="<%=font1%>" color="<%=textcolor2%>"><%=sCareer%></font></td>
				<td align="right"><font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>"><b>Member Since:</b></font></td>
				<td align="left">
					<font size="<% =fontsize3 %>" face="<%=font1%>" color="<%=textcolor2%>"><%=sMembSinceAge%></font>
					<font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>">years old</font>
				</td>
	    </tr>
			<tr>
				<td align="right"><font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>"><b>School:</b></font></td>
				<td align="left"><font size="<% =fontsize3 %>" face="<%=font1%>" color="<%=textcolor2%>"><%=sSchool%></font></td>
				<td align="right"><font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>"><b>Local Paper:</b></font></td>
				<td align="left"><font size="<% =fontsize3 %>" face="<%=font1%>" color="<%=textcolor2%>"><%=sPaper%></font></td>
	    </tr>
	    <tr><td colspan=4>&nbsp;</td></tr>
		</TABLE>

    <TABLE class="innertable" align="center" width=100%>
			<tr>
  			<th colspan="8" align="center"> <font size="3" face="<%=font1%>" color="#FFFFFF"><b>Events Entered</b></font></th>
			</tr>
	    <tr>
	  	  <td align="right" width="20%"><font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>"><b>Events Entered:</b></font></td>
				<td align="left" width="20%"><font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>"><b>Division</b></font></td>
				<td align="left" width="30%"><font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>"><b>Tournament Best</b></font></td>
				<td align="left" width="30%"><font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>"><b>Ski/Board (Model/Length)</b></font></td>
	    </tr>	
	    <tr>
				<td align="right"><font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>"><b>Slalom:</b></font></td>
	   		<td align="left"><font size="<%=fontsize3%>" face="<%=font1%>" color="<%=textcolor2%>"><%=ThisDiv1%></b></font></td>
				<td align="left"><font size="<% =fontsize3 %>" face="<%=font1%>" color="<%=textcolor2%>"><% =sBestSlal %></font></td>
				<td align="left"><font size="<% =fontsize3 %>" face="<%=font1%>" color="<%=textcolor2%>"><% =sFav_Slalom %></font></td>
	    </tr>
	    <tr>
				<td align="right"><font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>"><b>Trick:</b></font></td>
	   		<td align="left"><font size="<%=fontsize3%>" face="<%=font1%>" color="<%=textcolor2%>"><%=ThisDiv2%></b></font></td>
				<td align="left"><font size="<% =fontsize3 %>" face="<%=font1%>" color="<%=textcolor2%>"><% =sBestTrick %></font></td>
				<td align="left"><font size="<% =fontsize3 %>" face="<%=font1%>" color="<%=textcolor2%>"><% =sFav_Trick %></font></td>
	    </tr>
	    <tr>
				<td align="right"><font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>"><b>Jump:</b></font></td>
	   		<td align="left"><font size="<%=fontsize3%>" face="<%=font1%>" color="<%=textcolor2%>"><%=ThisDiv3%></b></font></td>
				<td align="left"><font size="<% =fontsize3 %>" face="<%=font1%>" color="<%=textcolor2%>"><% =sBestJump %></font></td>
				<td align="left"><font size="<% =fontsize3 %>" face="<%=font1%>" color="<%=textcolor2%>"><% =sFav_Jump %></font></td>
	    </tr>
	    <tr><td colspan=4>&nbsp;</td></tr>
		</TABLE>
	
	  <TABLE class="innertable" align="center" width=100%>
			<tr>
  			<th colspan="8" align="center"> <font size="3" face="<%=font1%>" color="#FFFFFF"><b>Records Titles and Accomplishments</b></font></th>
			</tr>
     	<tr>
				<td align="right" width="20%"><font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>"><b>Records:</b></font></td>
				<td align="left" colspan=3><font size="<% =fontsize3 %>" face="<%=font1%>" color="<%=textcolor2%>"><%=sRecords%></font></td>
    	</tr>
    	<tr>
				<td align="right"><font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>"><b>Titles:</b></font></td>
				<td align="left" colspan=3><font size="<% =fontsize3 %>" face="<%=font1%>" color="<%=textcolor2%>"><%=sTitles%></font></td>
			</tr>
			<tr>
				<td align="right"><font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>"><b>Accomplishments:</b></font></td>
				<td align="left" colspan=3 width="40%"><font size="<% =fontsize3 %>" face="<%=font1%>" color="<%=textcolor2%>"><%=sAccomplish%></font></td>
	    </tr>
	    <tr>
				<td align="right"><font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>"><b>Sponsors:</b></font></td>
				<td align="left"><font size="<% =fontsize3 %>" face="<%=font1%>" color="<%=textcolor2%>"><%=sSponsors%></font></td>
				<td align="right" width="20%"><font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>"><b>Mentors:</b></font></td>
				<td align="left" ><font size="<% =fontsize3 %>" face="<%=font1%>" color="<%=textcolor2%>"><%=sMentors%></font></td>
	    </tr>
	    <tr>
				<td align="right"><font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>"><b>Favorite Boat:</b></font></td>
				<td align="left"><font size="<% =fontsize3 %>" face="<%=font1%>" color="<%=textcolor2%>"><%=sFav_Boat%></font></td>
				<td colspan=2 width="50%">&nbsp</td>
	    <tr>
				<td align="right"><font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>"><b>Hobbies:</b></font></td>
				<td align="left" colspan=3><font size="<% =fontsize3 %>" face="<%=font1%>" color="<%=textcolor2%>"><%=sHobby%></font></td>
	    </tr>
	    <tr>
				<td align="right"><font size="<%=fontsize2%>" face="<%=font1%>" color="<% =textcolor1 %>"><b>Favorite Sports:</b></font></td>
				<td align="left"><font size="<% =fontsize3 %>" face="<%=font1%>" color="<%=textcolor2%>"><%=sFav_Sports%></font></td>
				<td colspan=2>&nbsp</td>
			</tr>
		</TABLE>
		<%

		IF NOT rs.eof THEN
			%><div class=break /><%
		END IF  



ELSE
	
	' ----  Do nothing - Watch for page advance - blank page

END IF
		








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






' -----------------------------------------
   SUB HowEmptyIsForm (sMemberID, sTourID) 
' -----------------------------------------


'markdebug("sMemberID in ReadFromTABLE = "&sMemberID)


'sMemberID = Session("sMemberID")
'sTourID = Session("sTourID")

set rsMemb=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT TOP 1 * FROM "&MemberTABLEName
sSQL = sSQL + " WHERE PersonIDwithCheckDigit = "&sqlclean(sMemberID)
rsMemb.open sSQL, sConnectionToTRATABLE, 3, 1

sFirstName = rsMemb("FirstName")
sLastName = rsMemb("LastName")
sMembSex = rsMemb("Sex")
'sMembAge = Session("MembAge")

SET rsBio=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT * FROM "&BioTABLEName
sSQL = sSQL + " WHERE MemberID = '"&sMemberID&"'"
rsBio.open sSQL, SConnectionToTRATABLE, 3, 3

ECount=0

IF NOT rsBio.eof THEN
'	sAddress1 = rsBio("Address1")

	sregion = TRIM(rsBio("region"))
	sCity = rsBio("City")
	IF TRIM(sCity)="" THEN ECount=ECount+5

	sState = rsBio("State")
	IF TRIM(sState)="" THEN ECount=ECount+5

	sEmail = rsBio("Email")
	sPhone = rsBio("Phone")
	sWeight = rsBio("Weight")
	IF TRIM(sWeight)="" THEN ECount=ECount+2

	sHgtFeet = rsBio("HgtFeet")
	sHgtInch = rsBio("HgtInch")
	IF TRIM(sHgtFeet)="" OR TRIM(sHgtInch)="" THEN ECount=ECount+5

	sSkiSinceAge = rsBio("SkiSinceAge")
	IF TRIM(sSkiSinceAge)="" THEN ECount=ECount+3

	sCompSinceAge = rsBio("CompSinceAge")
	IF TRIM(sCompSinceAge)="" THEN ECount=ECount+3

	sMembSinceAge = rsBio("MembSinceAge")

	sSponsors = rsBio("Sponsors")
	sClub = rsBio("Club")
	sSchool = rsBio("School")
	sOccup = rsBio("Occup")
	IF TRIM(sOccup)="" THEN ECount=ECount+3

	sCareer = rsBio("Career")
	sHobby = rsBio("Hobby")
	IF TRIM(sHobby)="" THEN ECount=ECount+2
	sPaper = rsBio("Paper")
	IF TRIM(sPaper)="" THEN ECount=ECount+4

	sBestSlal = rsBio("BestSlal")
	sBestHydro = rsBio("BestHydro")
	sBestTrick = rsBio("BestTrick")
	IF TRIM(sBestSlal)="" AND TRIM(sBestTrick)="" AND TRIM(sBestJump)="" THEN ECount=ECount+5

	sBestFree = rsBio("BestFree")
	sBestJump = rsBio("BestJump")
	sBestKnee = rsBio("BestKnee")
	sBestWake = rsBio("BestWake")
	sBestMara = rsBio("BestMara")

	sFav_Slalom = rsBio("Fav_Slalom")
	sFav_Jump = rsBio("Fav_Jump")
	sFav_Trick = rsBio("Fav_Trick")
	IF TRIM(sFav_Slalom)="" AND TRIM(sFav_Jump)="" AND TRIM(sFav_Trick)="" THEN ECount=ECount+5

	sFav_Wake = rsBio("Fav_Wake")
	sFav_Boat = rsBio("Fav_Boat")
	IF TRIM(sFav_Boat)="" THEN ECount=ECount+3

	sMentors = rsBio("Mentors")
	sAccomplish = rsBio("Accomplish")
	sFav_Sports = rsBio("Fav_Sports")
	sTitles = rsBio("Titles")

	sRecords = rsBio("Records")

	sLastUpDate = rsBio("LastUpdate")


END IF

rsMemb.Close
rsBio.Close

END SUB


%>






