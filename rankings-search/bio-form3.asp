
<%

' ----------------------------------------------------------------------------------------------------------------
' --- FILE:  Bio-form3.asp 
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

sTEvent1Name="Slalom"
sTEvent2Name="Tricks"
sTEvent3Name="Jump"
sTEvent4Name=""





set rsMemb=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT TOP 1 * FROM "&MemberTableName
sSQL = sSQL + " WHERE PersonIDwithCheckDigit = "&sqlclean(sMemberID)
rsMemb.open sSQL, sConnectionToTRATable, 3, 1

sFirstName = rsMemb("FirstName")
sLastName = rsMemb("LastName")
sMembSex = rsMemb("Sex")
'sMembAge = Session("MembAge")

SET rsBio=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT * FROM "&BioTableName
sSQL = sSQL + " WHERE MemberID = '"&sMemberID&"'"
rsBio.open sSQL, SConnectionToTRATable, 3, 3


' --- Found EXISTING record
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


	%>

	<HTML><HEAD>

	<style>div.break {page-break-before:always}</style>

	</HEAD><BODY>

 	<TABLE BORDER=4 ALIGN=CENTER CELLPADDING=3 CELLSPACING=0 BGCOLOR="#FFFFFF" width=100% >

	  <table ALIGN="center" BORDER="0" CELLPADDING="3" CELLSPACING="0" width=100%>   	
      	  <tr> 
	    <td ALIGN="center">
		<font size=<% =fontsize4 %> face=<% =font1 %> COlOR="red"><b><%=sTName%></b></FONT>
		<br>
		<font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><%=sTDateS%>-<%=sTDateE%></FONT>
		<br><br>
		<font size=<% =fontsize4 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Division:
		<font size=<% =fontsize4 %> face=<% =font1 %> COlOR="<% =textcolor2 %>">&nbsp;<%=DivSelected%>&nbsp;&nbsp;&nbsp;&nbsp;
		<font size=<% =fontsize4 %> face=<% =font1 %> COlOR="<% =textcolor1 %>">Events:
		<font size=<% =fontsize4 %> face=<% =font1 %> COlOR="<% =textcolor2 %>">&nbsp;<%=EventSelected%></b></FONT>
		<br><br>
	    </td>

	</tr>  
	</TABLE>


      <TABLE BORDER="4" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="<% =TableColor1 %>" width=100%>
	<TR>
	  <TD colspan="8" align="center" BGCOLOR="<% =HeadColor1 %>"><FONT size="3" face=<% =font1 %> COlOR="<% =textcolor1 %>"><strong>Personal Bio Form</strong></FONT></td></tr>
	</TR>  
 
	<TR>
	  <TD VALIGN="top">


	  <table ALIGN="center" BORDER="1" CELLPADDING="3" CELLSPACING="0" width=100% BGCOLOR="<%=TableColor1%>">   	
      	  <tr> 
	    <td ALIGN="right"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Name:</b>&nbsp;&nbsp</FONT></td>
	    <td colspan=3><font size="<% =fontsize3 %>" face=<% =font1 %> COlOR="<% =textcolor2 %>"><% =sFirstName %>&nbsp;<% =sLastName %></font></td>

   	    <td ALIGN="right"><FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR=<% =textcolor1 %>><b>Age/Gender:</b>&nbsp;&nbsp</FONT></td>
	    <td><font size=<% =fontsize3 %> face=<% =font1 %> COlOR="<% =textcolor2 %>"><% =sMembAge %>/<% =sMembSex %></font></td> 
	  </tr>


      	  <tr> 
	    <td ALIGN="right"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Events Entered:</b></FONT></td><%
		IF TRIM(sDiv1)<>"" THEN %>
			<td ALIGN="right">
			  <font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b><%=sTEvent1Name%></b> - </FONT> 
		    	  <font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor2 %>"><b><%=sDiv1%></b></FONT>
	        	</td><%
	    	ELSE %>
			<td ALIGN="right">
			  <font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b><%=sTEvent1Name%></b> - </FONT> 
		    	  <font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor2 %>"><b>***</b></FONT>
			</td><%
	    	END IF

		IF TRIM(sDiv2)<>"" THEN %>
			<td ALIGN="right">
			  <font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b><%=sTEvent2Name%></b> - </FONT>
		    	  <font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor2 %>"><b><%=sDiv2%></b></FONT>
	        	</td><%
	    	ELSE %>
			<td ALIGN="right">
			  <font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b><%=sTEvent2Name%></b> - </FONT>
		    	  <font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor2 %>"><b>***</b></FONT>
			</td><%
	    	END IF

	    	IF TRIM(sDiv3)<>"" THEN %>
			<td ALIGN="right">
			  <font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b><%=sTEvent3Name%></b> - </FONT>
	    		  <font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor2 %>"><b><%=sDiv3%></b></FONT>
			</td><% 
	    	ELSE %>
			<td ALIGN="right">
			  <font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b><%=sTEvent3Name%></b> - </FONT>
		    	  <font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor2 %>"><b>***</b></FONT>
			</td><%
		END IF %>
	     <td colspan=2>&nbsp;</td>
	  </tr>	

	  <tr>
	    <td ALIGN="right"><FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR=<% =textcolor4 %>>&nbsp;&nbsp;&nbsp;<b>City:</b></FONT></td>
	    <td>
		<input type="text" <% =BioStatus %> name="fCity" value= "<% =sCity %>"  MAXLENGTH=20 size="20"><%
		IF ChargeStatus="confirm" AND TRIM(sCity)="" THEN
			FieldErr = FieldErr +1
			%><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="red">* Required</font><%
		END IF  %>			
	    </td>

	    <td ALIGN="right"><FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR=<% =textcolor4 %>><b>State:</b></FONT></td>
    	
	    <td><%
		IF ChargeStatus="confirm" AND TRIM(sState)="" THEN
			FieldErr = FieldErr +1
			%><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="red">* Required</font><%
		ELSEIF ChargeStatus="confirm" THEN %>			
			<input type="text" <% =BioStatus %> name="fState" value= "<% =sState %>" size="2"><%
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
	    <td align="left"><select name='fregion' <% =BioStatus %> value=<% =sRegion %> >
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
	    <td colspan="3">
		<input type="text" <% =BioStatus %> name="fEmail" value= "<% =sEmail %>" MAXLENGTH=30 size="40">
	    </td><%
		IF ChargeStatus="confirm" AND TRIM(sEmail)="" THEN
			' FieldErr = FieldErr +1  Not required
			%><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor2 %>">&nbsp</font><%
		END IF  %>
	    

	    <td ALIGN="right"><FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR=<% =textcolor4 %>>&nbsp;&nbsp;&nbsp;<b>Phone:</b></FONT></td>

	    <td>
		<input type="text" <% =BioStatus %> name="fPhone" value= "<% =sPhone %>" MAXLENGTH=12 size=12><%
		IF ChargeStatus="confirm" AND TRIM(sCity)="" THEN
			FieldErr = FieldErr +1
			%><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="red">* Required</font><%
		END IF  %>			
	    </td>
	  </tr>	

	  <tr>  
	    <td ALIGN="right"><FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR=<% =textcolor4 %>>&nbsp;&nbsp;&nbsp;<b>Weight:</b></FONT></td><%
		IF ChargeStatus="confirm" THEN
			FieldErr = FieldErr +1
			%><td><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="red">* Required</font></td><%
		ELSEIF ChargeStatus="confirm" THEN
			%><td><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor2 %>"><% =sWeight %></font></td><%
		ELSE
			%><td><select <% =BioStatus %> name="fWeight" ><% LoadDropDown "LBS", sWeight,50,300,5  %></select>
			<font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"> lbs</font></td><%
		END IF  %>			


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
		<td><select name="fSkiSinceAge" <% =BioStatus %>><% LoadDropDown "Age", sSkiSinceAge,2,90,1  %></select>
		<font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>">Years Old</font></td>

		<td align="right"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Competing Since:</b></font></td>
		<td><select name="fCompSinceAge" <% =BioStatus %>><% LoadDropDown "Age", sCompSinceAge,2,90,1  %></select>
		<font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>">Years Old</font></td>

		<td align="right"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Member Since:</b></font></td>
		<td><select name="fMembSinceAge" <% =BioStatus %>><% LoadDropDown "Age", sMembSinceAge,2,90,1  %></select>
		<font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>">Years Old</font></td>
	    </tr>


	</table>
	<br>
	<table ALIGN="center" BORDER="1" CELLPADDING="3" CELLSPACING="0" width=100% BGCOLOR="<%=TableColor1%>">   	

	    <tr>
		<td ALIGN="right" colspan=1><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Occupation:</b></font></td>
		<td colspan=2><input type="text" <% =BioStatus %> name="fOccup" value= "<% =sOccup %>" MAXLENGTH=25 size="30"></td>
		<td ALIGN="right" colspan=1><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Career Plans:</b></font></td>
		<td colspan=2><input type="text" <% =BioStatus %> name="fCareer" value= "<% =sCareer %>" MAXLENGTH=25 size="30"></td>
	    </tr>

	     <tr>
		<td ALIGN="right" colspan=1><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Hobbies:</b></font></td>
		<td colspan=2><input type="text" <% =BioStatus %> name="fHobby" value= "<% =sHobby %>" MAXLENGTH=50 size="55"></td>
		<td ALIGN="right" colspan=1><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>School:</b></font></td>
		<td colspan=2><input type="text" <% =BioStatus %> name="fSchool" value= "<% =sSchool %>"  MAXLENGTH=25 size="30"></td>
	     </tr>

	     <tr>
		<td ALIGN="right" colspan=1><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Mentors:</b></font></td>
		<td colspan=2><input type="text" <% =BioStatus %> name="fMentors" value= "<% =sMentors %>" MAXLENGTH=60 size="65"></td>
		<td ALIGN="right" colspan=1><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Ski Club:</b></font></td>
		<td colspan=2><input type="text" <% =BioStatus %> name="fClub" value= "<% =sClub %>" MAXLENGTH=25 size="30"></td>
	     </tr>

	     <tr>
		<td ALIGN="right"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Accomplishments:</b></font></td>
		<td colspan=2><input type="text" <% =BioStatus %> name="fAccomplish" value= "<% =sAccomplish %>" MAXLENGTH=60 size="65"></td>
		<td ALIGN="right" colspan=1><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Local Paper:</b></font></td>
		<td colspan=2><input type="text" <% =BioStatus %> name="fPaper" value= "<% =sPaper %>" MAXLENGTH=35 size="40"></td>
	     </tr>

	     <tr>
		<td ALIGN="right"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Records:</b></font></td>
		<td colspan=3><input type="text" <% =BioStatus %> name="fRecords" value= "<% =sRecords %>" MAXLENGTH=80 size="85"></td>
		<td>&nbsp</td>
	     </tr>

	     <tr>
		<td ALIGN="right"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Titles:</b></font></td>
		<td colspan=3><input type="text" <% =BioStatus %> name="fTitles" value= "<% =sTitles %>" MAXLENGTH=80 size="85"></td>
		<td>&nbsp</td>
	     </tr>
	     <tr>

		<td ALIGN="right" colspan=1><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Favorite Other Sports:</b></font></td>
		<td colspan=2><input type="text" <% =BioStatus %> name="fFav_Sports" value= "<% =sFav_Sports %>" MAXLENGTH=50 size="55"></td>
		<td>&nbsp</td>
		<td>&nbsp</td>
	     </tr>

			
	    <tr>
		<td align="right"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Sponsors:</b></font></td>
		<td colspan=2><input type="text" <% =BioStatus %> name="fSponsors" value= "<% =sSponsors %>"  MAXLENGTH=60 size="65"></td>
		<td ALIGN="right"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Favorite Boat:</b></font></td>
		<td><input type="text" <% =BioStatus %> name="fFav_Boat" value= "<% =sFav_Boat %>" MAXLENGTH=30 size="33"></td>
	    </tr>

	
	</table>
	<br>
	<table ALIGN="center" BORDER="1" CELLPADDING="3" CELLSPACING="0" width=100% BGCOLOR="<%=TableColor1%>">   	

	    <tr>
		<td WIDTH=10%>&nbsp</td>
		<td WIDTH=15% ALIGN="left"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Tournament Best</b></font></td>
		<td ALIGN="left"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Ski/Board (Model/Length)</b></font></td>
	    </tr>	

	    <tr>
		<td ALIGN="right" colspan=1><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Slalom:</b></font></td>
		<td><input type="text" <% =BioStatus %> name="fBestSlal" value= "<% =sBestSlal %>" MAXLENGTH=15 size="15"></td>
		<td><input type="text" <% =BioStatus %> name="fFav_Slalom" value= "<% =sFav_Slalom %>" MAXLENGTH=20 size="20"></td>
	    </tr>

	    <tr>
		<td ALIGN="right" colspan=1><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Trick:</b></font></td>
		<td><input type="text" <% =BioStatus %> name="fBestTrick" value= "<% =sBestTrick %>" MAXLENGTH=15 size="15"></td>
		<td><input type="text" <% =BioStatus %> name="fFav_Trick" value= "<% =sFav_Trick %>" MAXLENGTH=20 size="20"></td>
	    </tr>

	    <tr>
		<td ALIGN="right" colspan=1><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Jump:</b></font></td>
		<td><input type="text" <% =BioStatus %> name="fBestJump" value= "<% =sBestJump %>" MAXLENGTH=15 size="15"></td>
		<td><input type="text" <% =BioStatus %> name="fFav_Jump" value= "<% =sFav_Jump %>" MAXLENGTH=20 size="20"></td>
	    </tr>

	    <tr>
		<td ALIGN="right" colspan=1><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Wakeboard:</b></font></td>
		<td><input type="text" <% =BioStatus %> name="fBestWake" value= "<% =sBestWake %>" MAXLENGTH=15 size="15"></td>
		<td><input type="text" <% =BioStatus %> name="fFav_Wake" value= "<% =sFav_Wake %>" MAXLENGTH=20 size="20"></td>
	    </tr>

	    <tr>
		<td ALIGN="right"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Hydrofoil</b></font></td>
		<td><input type="text" <% =BioStatus %> name="fBestHydro" value= "<% =sBestHydro %>" MAXLENGTH=15 size="15"></td>
		<td>&nbsp</td>
	    </tr>

	    <tr>
	    	<td ALIGN="right"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Freestyle</b></font></td>
		<td><input type="text" <% =BioStatus %> name="fBestFree" value= "<% =sBestFree %>" MAXLENGTH=15 size="15"></td>
		<td>&nbsp</td>
	    </tr>

	    <tr>
	    	<td ALIGN="right"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Kneeboard</b></font></td>
		<td><input type="text" <% =BioStatus %> name="fBestKnee" value= "<% =sBestKnee %>" MAXLENGTH=15 size="15"></td>
		<td>&nbsp</td>
	    </tr>

	   <tr>
	    	<td ALIGN="right"><font size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor1 %>"><b>Marathon</b></font></td>
		<td><input type="text" <% =BioStatus %> name="fBestMara" value= "<% =sBestMara %>" MAXLENGTH=15 size="15"></td>
		<td>&nbsp</td>
	   </tr>


		</table>

		<table Align="center">
		   <tr>
		     <td>&nbsp</td>
		   </tr>
		   <tr>	
		   <td colspan=4 align="center">&nbsp;

		   </td>	
		   </tr>
		</table>

		</TD>	
		</TR>
		</TABLE><%

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


'markdebug("sMemberID in ReadFromTable = "&sMemberID)


'sMemberID = Session("sMemberID")
'sTourID = Session("sTourID")

set rsMemb=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT TOP 1 * FROM "&MemberTableName
sSQL = sSQL + " WHERE PersonIDwithCheckDigit = "&sqlclean(sMemberID)
rsMemb.open sSQL, sConnectionToTRATable, 3, 1

sFirstName = rsMemb("FirstName")
sLastName = rsMemb("LastName")
sMembSex = rsMemb("Sex")
'sMembAge = Session("MembAge")

SET rsBio=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT * FROM "&BioTableName
sSQL = sSQL + " WHERE MemberID = '"&sMemberID&"'"
rsBio.open sSQL, SConnectionToTRATable, 3, 3

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










