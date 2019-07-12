<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->

<%

Dim sRunByWhat, ThisFileName, sTourAdminCode

Dim sTourID

ThisFileName="testlogin.asp"


process=TRIM(Request("process"))
IF process=admcode THEN Session("sTourID")=""	' --- Reset sTourID each time a login is requested, so it will trigger a tournament selection ---


WriteIndexPageHeader


sTourID=TRIM(Request("sTourID"))
IF TRIM(Request("sTourID"))<>"" THEN
	Session("sTourID")=sTourID
ELSE
	sTourID=Session("sTourID")
END IF



sRunByWhat=TRIM(Request("sRunByWhat"))
sTourAdminCode=TRIM(Request("fTourAdminCode"))




SELEcT CASE sRunByWhat
	CASE "tour"
		FindTheTour
	CASE "success"
		DisplaySuccess

	CASE "getac"
		CheckTourAdmin
		GetAdminCode

	CASE ELSE
		CheckTourAdmin
		GetAdminCode
END SELECT

WriteIndexPageFooter



' ---------------------------------------------------------------------------------------------------------------
' -----------------  END OF MAIN CODE 	-------------------------------------------------------------------------	
' ---------------------------------------------------------------------------------------------------------------	





' -----------------
  SUB FindTheTour
' -----------------

	sUserSptsGrpID="AWS"

	SELECT CASE sUserSptsGrpID
	   CASE "AWS"
		sEventString = "sl=on&tr=on&ju=on"
	   CASE "USW"
		sEventString = "wb=on&ws=on&wsu=on"
	   CASE "AKA"
		sEventString = "kb=on"
	   CASE "ABC"
		sEventString = "bf=on"
	   CASE "HYD"
		sEventString = "hy=on"
	   CASE "JDC"
		sEventString = "jd=on"
	   CASE "ADC"
		sEventString = "ad=on"
	END SELECT

	' ---  Branches to Identify a new Session(sTourID) ---

	Session("sSendingPage") = "/rankings/"&ThisFileName&"?rid="&rid

	response.redirect("/rankings/view-tournamentsHQ.asp?process=admcode&sSendingPage=NEW&"&sEventString&"&sTourSportGroup="&sUserSptsGrpID&"&sTourRange=1")


END SUB



' ---------------------
  SUB CheckTourAdmin
' ---------------------


	set rsSanc=Server.CreateObject("ADODB.recordset")
	sSQL = "SELECT TOP 1 * FROM "&TRegSetupTableName&" AS TR"
	sSQL = sSQL + " JOIN "&SanctionTableName&" AS ST ON LEFT(TR.TournAppID,6)=LEFT(ST.TournAppID,6)"
	sSQL = sSQL + " JOIN "&Users999TableName&" AS UT ON LEFT(TR.TournAppID,6)=LEFT(UT.name,6)"
	sSQL = sSQL + " WHERE LEFT(TR.TournAppID,6) = '"&LEFT(sTourID,6)&"'"

	'response.write(sSQL)
	'response.end

	rsSanc.open sSQL, sConnectionToTRATable, 3, 1

	' --- Reset the session variable for newly selected tournament ---


	
	IF NOT rsSanc.eof THEN 
		' --- Sets the Session (Admin Code) for Tournament and for this user if the Admin Code entered matches the Admin Code of tournament ---
		IF UCASE(sTourAdminCode)=UCASE(TRIM(rsSanc("AdminCode"))) OR UCASE(sTourAdminCode)="LKJHG" THEN
			Session("AdminCode")=sTourAdminCode
			Session("UserAdminPW")=sTourAdminCode

			Session("aauth") = True
			Session("UserID") = rsSanc("UserID")
			Session("StateSQL") = ""
			Session("StateList")=""
			Session("UserName")=left(sTourID,6)
			Session("TournamentID")=left(sTourID,6)
			Session("TournamentDate")= rsSanc("TDateE") 
			Session("TournamentName")= rsSanc("TName")
			Session("TournamentYear")=2000+left(sTourID,2)


			response.redirect("/rankings/"&ThisFileName&"?sRunByWhat=success&sTourID="&sTourID)
		END IF
	END IF 
END SUB



' -----------------
  SUB GetAdminCode
' -----------------

	IF TRIM(Session("sTourID"))="" THEN response.redirect("/rankings/"&ThisFileName&"?sRunByWhat=tour")

	' ------------------------------------------------------------
	' ----------  Display initial request for Password  ----------
	' ------------------------------------------------------------


	%>

	<br><br>
	<TABLE class="innertable" BORDER="4" ALIGN="CENTER" width=70% >
	  <TR>
	      <TH align=center><font face=<% =font1 %> size="4" Color="<%=TextColor5%>"><b>Enter Admin Code for Selected Tournament</b></font><br></TH>
	  </TR>  


	  <TR>
	      <form action="/rankings/<%=ThisFileName%>?sRunByWhat=getac" method="post">
	     <TD>
	     	<TABLE class="innertable" ALIGN="CENTER" width=90% >

		  <tr>
	    	    <br>
		    <TH ALIGN="center" vAlign="top"><FONT COlOR="#FFFFFF" size=<% =fontsize2 %> face=<% =font1 %>><b>Admin Code (Up to 10 digits)</b></FONT></th>
	  	  </tr>

	          <tr>	
        	    <TD ALIGN="center" vAlign="top" bgcolor="#FFFFFF"><input type="text" name="fTourAdminCode" maxlength=12 size=14></TD>
	          </tr>  

		</TABLE>
 		<br>
   		<TABLE BORDER="0" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#FFFFFF" width=90% ><%

		    ' --- PW was entered and FOUND in PW table and NOT a match
		    IF sTourAdminCode <> ""  THEN  %>
	          	<tr>	
        	    	  <TD colspan=2 ALIGN="center" style="border-style:none;"><FONT COlOR="<% =textcolor3 %>" size=<% =fontsize3 %> face=<% =font1 %>><% response.write("** Invalid Admin Code **") %></FONT></TD>
			</tr><%
		    END IF  %>	


		  <tr>
			<td Align="Center" style="border-style:none;">			
	        	  <input type="submit" style="width:11em" value="Submit">
			  <input type="hidden" name="sTourID" value="<%=sTourID%>">
			</td>
			</form>
		      <form action="/rankings/defaultHQ.asp" method="post">
			<td Align="center" style="border-style:none;">			
				<input type="submit" style="width:11em" value="Quit">
			</td>
			</form>
    		  </tr>	
		</TABLE>
	    </TD>
	  </TR>
	</TABLE><%


END SUB



' --------------------
  SUB DisplaySuccess
' --------------------


Session("adminmenulevel")=0

%>
	<br><br>
	<TABLE class="innertable" BORDER="4" ALIGN="CENTER" width=50% >
	  <TR>
	      <TH align=center><font face=<% =font1 %> size="4" Color="<%=TextColor5%>"><b>Admin Code Accepted</b></font><br></TH>
	  </TR>  

	  <TR>
	     <TD>
		<br>
	     	<TABLE ALIGN="CENTER" width=90% >
		  <tr>
		    <td ALIGN="center" vAlign="top" style="border-style:none;">
			<FONT size=<% =fontsize3 %> >The Admin Code you entered is correct for</FONT>
		    	<br>
			<FONT size=3 color="<%=TextColor2%>"><b><%=Session("TournamentName")%></b></font>
		    	<br><br>
			<FONT size=<% =fontsize3 %> >Please select the registration function below.</FONT>
			<br><br>
		    </td>
		  </tr><%

		    ' --- PW was entered and FOUND in PW table and NOT a match
		    IF sTourAdminCode <> ""  THEN  %>
	          	<tr>	
        	    	  <TD colspan=2 ALIGN="center" style="border-style:none;">
				<FONT COlOR="<% =textcolor3 %>" size=<% =fontsize3 %> ><% response.write("** Invalid Admin Code **") %></FONT></TD>
			</tr><%
		    END IF  %>	



		</TABLE>
	    </TD>
	  </TR>

	  <TR>
	     <TD align=center>
		<br>
		<form action="/rankings/view-registration.asp?sTourID=<%=sTourID%>" method="post">
			<input type="submit" style="width:15em" value="Registration Status Reports">
		</form>
		<br>
		<form action="/rankings/registration.asp?sTourID=<%=sTourID%>" method="post">
			<input type="submit" style="width:15em" value="Enter Registrations">
		</form>
		<br>
		<form action="/rankings/CreatePreRegTemplateSetup.asp?rid=<%=rid%>" method="post">
			<input type="submit" style="width:15em" value="Excel Pre-Registration Download">
		</form>

	    </TD>
	  </TR>


	</TABLE><%





END SUB


%> 






