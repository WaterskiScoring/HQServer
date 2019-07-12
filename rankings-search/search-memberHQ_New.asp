<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->


<%

Dim ThisFileName
ThisFileName="Search-MemberHQ_New.asp"

' --- TEMP
'PWreq = "no"


Dim currentPage, sMemberID, sLastName, sFirstName, RunByWhat, sSendingPage, pVar
'Dim sPassword
Dim adminmenulevel
Dim sUserAdminPW

Dim Ebody, sIncludeClubs
Dim DialogHeader

Dim FormStatus
Dim sTname, sAddressList, sEmailText, sTempEmail1, sTempEmail2, sTempPW, sPassword, sTourAdminPW, NowDate, sTableStatus, sTableEmail, sTableAC
Dim DisplayHeadFoot
Dim ScorePageBorderLight, ScorePageBorderDark

Dim sPasswordEmail, sPasswordEmailAdm, sPasswordEmailHQ
Dim sDispDebugButtons, sDispDebugButtonsAdm, sDispDebugButtonsHQ






' --- Reads settings from email and display control table ---
ReadContDispTableValues



DialogHeader=HQSiteColor2

ScorePageBorderDark = HQSiteColor1
ScorePageBorderLight = HQSiteColor2


FormStatus = TRIM(Request("FormStatus"))
sMemberID = TRIM(sqlclean(Request("sMemberID")))
sLastName = TRIM(sqlclean(Request("sLast_Name")))
sFirstName = TRIM(sqlclean(Request("sFirst_Name")))

sPassword = TRIM(sqlclean(Request("fPassword")))
'sTempEmail1 = TRIM(sqlclean(Request("fTempEmail1")))
'sTempEmail2 = TRIM(sqlclean(Request("fTempEmail2")))
'sTempPW = TRIM(sqlclean(Request("fTempPW")))

sIncludeClubs=Request("sIncludeClubs")


' ------------------------------
' --- Reads the sending page ---
' ------------------------------
sSendingPage = TRIM(Request("sSendingPage"))
IF TRIM(Session("sSendingPage"))<>"" THEN sSendingPage=TRIM(Session("sSendingPage"))

' ----------------------------------------------
' --- AdminCode as been established as Admin ---
' ----------------------------------------------
sTourAdminPW=TRIM(Request("sTourAdminPW"))
IF TRIM(Session("AdminCode"))<>"" THEN sTourAdminPW = Session("AdminCode")
IF TRIM(sTourAdminPW)="" THEN sTourAdminPW="ZYXW1"


sUserAdminPW=TRIM(Request("sUserAdminPW"))
IF TRIM(Session("UserAdminPW"))<>"" THEN sUserAdminPW = Session("UserAdminPW")



' ---------------------
' --- Pulls sTourID ---
' ---------------------
sTourID=TRIM(Request("sTourID"))
IF TRIM(Session("sTourID"))<>"" THEN sTourID=Session("sTourID")

' -------------------------------------------------
' --- adminmenu level ---
' -------------------------------------------------
adminmenulevel=TRIM(Request("adminmenulevel"))
IF TRIM(Session("adminmenulevel"))<>"" THEN adminmenulevel=Session("adminmenulevel")
IF TRIM(adminmenulevel)="" THEN adminmenulevel=1







' -----------------------------------------------------------------------------------------------------------------------
' --- This code determines whether the calling program requires a password.
' --- If so, then the PWreq is set so this module will remember that the PW is required.
' --- The sSendingPage is established by the calling program.  This is an added measure of security to prevent 
' ---    users from attempting to access the program by monkeying with the querystring in the URL.
' -----------------------------------------------------------------------------------------------------------------------


IF sSendingPage = "/Bio-Form.asp?formstatus=new&sReturnto=default&biostatus=enabled" THEN
		'PWreg="NO"
ELSEIF LCASE(LEFT(sSendingPage,21)) = "/rankings/view-scores" THEN
		PWreq="no"
ELSEIF LCASE(LEFT(sSendingPage,24)) = "/rankings/view-standings" THEN
		PWreq="no"
ELSEIF LCASE(LEFT(sSendingPage,23)) = "/rankings/sample_insert" THEN
		PWreq = "no"
		DisplayHeadFoot="no"
ELSEIF LCASE(LEFT(sSendingPage,25)) = "/rankings/member-personal" THEN
		PWreq = "no"
ELSEIF LCASE(LEFT(sSendingPage,22)) = "/rankings/registration" THEN
		PWreq = "yes"
END IF



IF DisplayHeadFoot<>"no" THEN
		' -------------------------------------------
		' --- Writes the HQ main page with menus  ---
		' -------------------------------------------
		WriteIndexPageHeader
END IF



'response.write("Line 132 - Sending Page = "&sSendingPage)
'response.write("<br>Session(sSendingPage) = "&Session("sSendingPage"))
'response.write("<br>FormStatus = "&FormStatus)
'response.write("<br>PWReq = "&PWReq)
'response.write("<br>sMemberID = "&sMemberID)
'response.write("<br>Session(sMemberID) = "&Session("sMemberID"))



' --- Use this to force it to go to new search
IF FormStatus = "search" THEN
		FindaMemberID

ELSEIF FormStatus = "found" THEN
		' --- User has chosen a Member so now make user confirm
		DisplayMemberData

'ELSEIF FormStatus = "getpw" AND sMemberID<>"" THEN
ELSEIF FormStatus = "getpw" THEN
		' --- Dialog boxes and logic for password input
		'response.write("<br> Line 151 IN ELSEIF ")
		'response.end
		GetPassword


ELSEIF FormStatus = "confirmed" THEN

		'response.write("<br> Line 158 IN CONFIRMED ")

		' --- User has confirmed, so branch back to Session variable sSendingPage
		Session("Know_Orig_Trans") = ""

		IF TRIM(sMemberID)<>"" THEN Session("sMemberID") = sMemberID
		IF TRIM(sTourID)<>"" THEN Session("sTourID") = sTourID
		IF TRIM(sTourAdminPW)<>"" THEN Session("AdminCode")=sTourAdminPW
		IF TRIM(sUserAdminPW)<>"" THEN Session("UserAdminPW")=sUserAdminPW
		IF TRIM(adminmenulevel)<>"" THEN Session("AdminMenuLevel")=adminmenulevel

		' --- If PWReg is YES then establish the Session variable for sMemberID
		' --- This allows the calling program to receive the sMemberID without passing as a querystring or hidden.

		IF PWreq="yes" THEN
				PWreq=""
				response.redirect(sSendingPage)

		ELSEIF LCASE(LEFT(sSendingPage,21)) = "/rankings/view-scores" THEN
				sSendingPage = sSendingPage&"&sMemberID="&sMemberID
				response.redirect(sSendingPage)
		ELSEIF LCASE(LEFT(sSendingPage,24)) = "/rankings/view-standings" THEN
				sSendingPage = sSendingPage&"&sMemberID="&sMemberID
				response.redirect(sSendingPage)

		ELSEIF LCASE(LEFT(sSendingPage,25)) = "/rankings/member-personal" THEN
				sSendingPage = sSendingPage&"&sMemberID="&sMemberID
				response.redirect(sSendingPage)

		ELSEIF LCASE(LEFT(sSendingPage,24)) = "/rankings/pw_update_tool" THEN
				sSendingPage = sSendingPage&"?sMemberID="&sMemberID
				Session("sMemberID") = sMemberID
				response.redirect(sSendingPage)

		ELSEIF TRIM(sSendingPage)="" THEN
				DisplayTimeOutNotice

		ELSE
				' --- User has confirmed, so branch back to Session variable sSendingPage
				response.redirect(sSendingPage)
	
		END IF

ELSE
		DisplayTimeOutNotice

END IF




' -------------------------------------------
' --- Writes the HQ main page with menus  ---
' -------------------------------------------

WriteIndexPageFooter




' ------------------------------------------------------------------------------------------
' -------------------------------   END OF MAIN PROGRAM   ----------------------------------
' ------------------------------------------------------------------------------------------



' ----------------------------------
  SUB ReadContDispTableValues
' ----------------------------------

' --- Read transactions from Credit Card Table to determine Total Fees actually completed ----
SET rsContDisp=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT * FROM "&ControlDisplayTableName
rsContDisp.open sSQL, SConnectionToTRATable, 3, 3

IF NOT rsContDisp.eof THEN

	' --- Values read from table are convert back to control the form checkboxes ---
	sPasswordEmail=rsContDisp("PasswordEmail")
	sPasswordEmailAdm=rsContDisp("PasswordEmailAdm")
	sPasswordEmailHQ=rsContDisp("PasswordEmailHQ")

	sDispDebugButtons=rsContDisp("DispDebugButtons")
	sDispDebugButtonsAdm=rsContDisp("DispDebugButtonsAdm")
	sDispDebugButtonsHQ=rsContDisp("DispDebugButtonsHQ")

END IF


END SUB







' ---------------------
  SUB SetupDone
' ---------------------


%>
<br>
<TABLE class="innertable" BORDER="4" ALIGN="CENTER" width="60%" >
<TR>
  <TH align="center"><font color="#FFFFFF" size="4"><b>Notice End of Session</b></font><br></TH>
</TR> 
 
  <br>
<TR>
  <TD>

	<TABLE ALIGN="CENTER" width="80%" >
	<tr>
    	  <td style="border-style:none;">&nbsp</td>			
	</tr>

	<tr>
	  <td ALIGN="center" style="border-style:none;">
		<font face=<% =font1 %> size="<% =fontsize3 %>"><b>Your password must be activated to proceed.</b></font>
		<br><br>
		<font face=<% =font1 %> size="<% =fontsize2 %>">Please check your email service.  An automated message has been sent to the email address you provided,
		<br> with a link which you must follow in order to activate your password. </font>  
		<br><br>
		<font face=<% =font1 %> size="<% =fontsize2 %>">Activation of your password will give you protected access to your online registration account.<br> Once the activation is complete, you will have full access to the online registration system. </font>
		<br><br>
		<font face=<% =font1 %> size="<% =fontsize1 %>">Please email us at competition@usawaterski.org if you do not receive the email notice, or <br>if the link does cause the activation of your account.</font>
	  </td>
	</tr>  

	<tr>
    	  <td style="border-style:none;">&nbsp</td>			
	</tr>
	</TABLE>



    </TD>
  </TR>
</TABLE><% 

END SUB








' -----------------
  SUB GetPassword
' -----------------


'response.write("<br> Line 309 GET PASSWORD - sSendingPage = "&sSendingPage )


' --- Revised 11-26-2013 to use MemberShort View ---
set rsPW=Server.CreateObject("ADODB.recordset")
sSQL = " SELECT FirstName, LastName, Password "
sSQL = sSQL + " FROM "&MemberShortTableName
sSQL = sSQL + " WHERE PersonID='"&RIGHT(sMemberID,8)&"'"

'response.write("<br>"&sSQL)
'response.end

rsPW.open sSQL, sConnectionToTRATable, 3, 1

sFirstName = rsPW("FirstName")
sLastName = rsPW("LastName")
sMemberPW = LCASE(TRIM(rsPW("Password")))



' ----------------------------------------------------------
' ----------- REDIRECTS if PW input is correct  ------------
' ----------------------------------------------------------

' --- Password matches the AdminCode of this tournament ---
IF adminmenulevel>= 20 THEN
		IF LCASE(sPassword)=LCASE(sTourAdminPW) THEN Session("UserAdminPW")=TRIM(sPassword)
		response.redirect("/rankings/"&ThisFileName&"?FormStatus=confirmed&sMemberID="&sMemberID)

' --- Session Password or Password entered matches the AdminCode of this tournament ---
ELSEIF TRIM(sTourAdminPW)<>"" AND ( sUserAdminPW=LCASE(sTourAdminPW) OR LCASE(sPassword)=LCASE(sTourAdminPW) ) THEN
		IF LCASE(sPassword)=LCASE(sTourAdminPW) THEN Session("UserAdminPW") = TRIM(sPassword)
		response.redirect("/rankings/"&ThisFileName&"?FormStatus=confirmed&sMemberID="&sMemberID)

' --- Password info entered and found in Member table and entered information matches member table
ELSEIF sPassword<>"" AND (NOT rsPW.eof) AND LCASE(sMemberPW)=LCASE(sPassword) THEN
		IF LCASE(sPassword) = LCASE(sTourAdminPW) THEN Session("UserAdminPW") = TRIM(sPassword)
		response.redirect("/rankings/"&ThisFileName&"?FormStatus=confirmed&sMemberID="&sMemberID)


ELSE
		' ----------  Display initial request for Password  ----------

'response.write("<br> Line 343 - sSendingPage = "&sSendingPage)

		Dim PWType
		IF sPassword="manual" THEN
				PWType="Activation Code"
		ELSE
				PWType="Password"
		END IF
		%>
		<form action="/rankings/<%=ThisFileName%>?FormStatus=getpw" method="post">
			<input type="hidden" name="sMemberID" value="<%=sMemberID%>">
			<input type="hidden" name="sSendingPage" value="<%=sSendingPage%>">
			<input type="hidden" name="sTourAdminPW" value="<%=sTourAdminPW%>">
			<input type="hidden" name="sTourID" value="<%=sTourID%>">
			<input type="hidden" name="adminmenulevel" value="<%=adminmenulevel%>">			
		<br><br>
		<TABLE class="innertable" BORDER="4" ALIGN="CENTER" width="60%" >
			<TR>
				<TH align=center><font face=<% =font1 %> size="4" Color="<%=TextColor5%>"><b>Enter <%=PWType%></b></font><br></TH>
			</TR>  
	  	<TR>
				<TD style="border-style:none;">
					<br>
					<TABLE class="innertable" ALIGN="CENTER" width="90%" >
						<tr>
					    <TH ALIGN="center" vAlign="top"><FONT COlOR="#FFFFFF" size=<% =fontsize2 %> face=<% =font1 %>><b><%=PWType%> (Up to 10 digits)</b></FONT></th>
		  		  </tr>
		        <tr>	
  	          <TD ALIGN="center" vAlign="top" bgcolor="#FFFFFF"><input type="text" name="fPassword" maxlength=12 size=14></TD>
	  	      </tr>  
				</TABLE>
   			<TABLE BORDER="0" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#FFFFFF" width="90%" >
   				<%

		    	' --- PW was entered and FOUND in PW table and NOT a match
		    	IF sPassword <> "" AND sPassword<>"manual" AND (NOT rsPW.eof) AND LCASE(sPassword) <> LCASE(sPassword) THEN  
			    		%>
	  	       	<tr>	
      	  	    <TD colspan=2 ALIGN="center" style="border-style:none;"><FONT COlOR="<% =textcolor3 %>" size=<% =fontsize3 %> face=<% =font1 %>><% response.write("** Invalid "&PWType&" **") %></FONT></TD>
							</tr>
							<%
		    	END IF  
		    	%>	
	        <tr>	
       	    <TD colspan=2 ALIGN="center" style="border-style:none;">
							<FONT size=<% =fontsize2 %> ><b>Tournament registrars should use the assigned Admin Code.</b></FONT>
		      		<br>
		    		</TD>
		  		</tr>
	        <tr>	
       	    <TD colspan=2 ALIGN="center" style="border-style:none;">
       	    	<br>
							<font size="<%=fontsize2%>" color="red"><b>IMPORTANT</b></font>
							<br>
							<font size="<%=fontsize2%>" > Effective Dec 1 2013, <b>Online Registration Passwords</b> are managed from the 'Members Only' section of the USA Waterski site.  Click to 
								<br>
								<a href='../members/login' title='Manage Passwords' target="_blank">Setup, Lookup or Manage Your Password</a></b>
							</font>   
		      		<br>
		    		</TD>
		  		</tr>
		  		<tr>
						<td Align="Center" style="border-style:none;">			
		  				<br>
        	  	<input type="submit" style="width:11em" value="Continue">
						</td>  
 		  		</tr>	
				</TABLE>

		    </TD>
		  </TR>
		</TABLE>
		</form>
		<% 
END IF


rsPW.close

END SUB













' -----------------
  SUB FindaMemberID
' -----------------


' User selected link to View Scores by Member

set rs=Server.CreateObject("ADODB.recordset")
sMemberID = replace(TRIM(request("Member_ID")),"-","")
sLastName = TRIM(request("Last_Name"))
sFirstName = TRIM(Request("First_Name"))
sCompanyName = TRIM(Request("Last_Name"))


' ********************************************
'   Nothing was put in any of the 3 fields yet.
' *********************************************

IF ((sMemberID = "" or NOT IsNumeric(sMemberID)) AND sLastName = "" AND sFirstName = "") THEN 
		DisplayMemberSearchFilters

		' ************************************************
		'   User entered something in at least one field
		' ************************************************

ELSE
		' -------------------------------------------------------------------
    ' --- Set up Query for member
		' -------------------------------------------------------------------
		
		sSQL = "SELECT DISTINCT TOP 20 PersonID, MT.LastName, MT.FirstName, MT.CompanyName"
		sSQL = sSQL + ", MT.City, MT.State, MT.BirthDate, MT.Sex" 
		sSQL = sSQL + " FROM "&MemberLiveTableName&" AS MT WHERE 1=1"

		IF sMemberID <> "" and IsNumeric(sMemberID) THEN
				tMemberID = PersonIDwChkDgt(RIGHT(sqlclean(sMemberID),8))
    		sSQL = sSQL + " AND PersonID = '" &RIGHT(tMemberID,8)& "'"
		END IF
	
		IF sLastName <> "" AND sIncludeClubs<>"on" THEN
				sSQL = sSQL + " AND lower(left(lastname," & len(sLastName) & ")) = '" & sqlclean(lCASE(sLastName)) & "'"
		ELSEIF sLastName <> "" AND sIncludeClubs="on" THEN
    		sSQL = sSQL + " AND lower(left(CompanyName," & len(sCompanyName) & ")) = '" & sqlclean(lCASE(sCompanyName)) & "'"
		END IF
	
		IF sFirstName <> "" THEN
    		sSQL = sSQL + " AND lower(left(firstname," & len(sFirstName) & ")) = '" & sqlclean(lCASE(sFirstName)) & "'"
		END IF


		' ---------------------------------------------------
		' --- Initial search based on user input in boxes ---
		' ---------------------------------------------------
		IF sMemberID <> "" THEN
				sSQL = sSQL + " ORDER BY PersonID"
		ELSEIF sLastName <> "" AND sIncludeClubs="on" THEN 
				sSQL = sSQL + " ORDER BY CompanyName"
		ELSEIF sLastName <> "" AND sIncludeClubs<>"on" THEN 
				sSQL = sSQL + " ORDER BY LastName, FirstName"
		ELSEIF sFirstName <> "" THEN 
				sSQL = sSQL + " ORDER BY FirstName, LastName"
		END IF
 		rs.open sSQL, sConnectionToTRATable, 3, 1







	' ******************************************
 	' No records found matching search criteria
	' ******************************************
	
	IF rs.EOF THEN %>
		<br>
		<TABLE class="innertable" BORDER="4" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="<%=TableColor1%>" width="50%" >
		  <TR>
				<Th align=center><font size="4" Color="<%=TextColor5%>"><b>Search Failed</b></font><br></Th>
		  </TR>  
		  <TR>
		    <TD>
					<TABLE BORDER="0" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="<%=TableColor1%>" width="90%" >
			  		<tr>
   			    	<td colspan=2 align=center>     		
	    					<br>
				  			<FONT COlOR="red" size=2 >No Records Found - Please Search Again </font>
 			        	<br><br>
			    		</td>
			  		</tr>
			  		<tr>
   			  		<td width=50% align=center>     		
								<form action="/rankings/<%=ThisFileName%>?rid=<%=rid%>&FormStatus=search" method="post">
									<input type="hidden" name="sSendingPage" value="<%=sSendingPage%>">
									<input type="hidden" name="sTourAdminPW" value="<%=sTourAdminPW%>">
									<input type="hidden" name="sTourID" value="<%=sTourID%>">
									<input type="hidden" name="adminmenulevel" value="<%=adminmenulevel%>">			
				  				<input type="hidden" name="pvar" value="SelectMember">
				  				<input type="submit" style="width:9em" value="New Search">
								</form>
			    		</td>
   			  		<td width=50% align=center>     				
								<form action="/defaultHQ.asp" method="post">
									<input type="hidden" name="sSendingPage" value="<%=sSendingPage%>">
									<input type="hidden" name="sTourAdminPW" value="<%=sTourAdminPW%>">
									<input type="hidden" name="sTourID" value="<%=sTourID%>">
									<input type="hidden" name="adminmenulevel" value="<%=adminmenulevel%>">			
				  				<input type="submit" style="width:9em" value="Quit">
								</form>
			    		</td>
			  		</tr>	
					</TABLE>
		  	</TD>
		  </TR>
		</TABLE>

<%

	ELSE			' Found at least ONE or MANY members in MemberTrak


			' *********************************************************************************************
			' Found MORE THAN ONE member in MemberTrak meeting search criteria - so display list of matches
			' *********************************************************************************************


			IF rs.recordcount > 1 THEN 
					%>
					<br><br>
					<TABLE class="innertable" ALIGN="CENTER" width="80%" >
			  		<tr>
			    		<th align=center><font face=<% =font1 %> size="4" Color="<%=TextColor5%>"><b>Select Member Link Below</b></font></th>
			  		</tr>  
			  		<tr>
			  			<td>
				  			<br>
								<TABLE class="innertable" ALIGN="CENTER" BGCOLOR="<%=TableColor1%>" width=95% >
			    				<tr>
			    				<%
				    			IF sIncludeClubs="on" AND LCASE(rs("Sex"))<>"male" AND LCASE(rs("Sex"))<>"female" THEN %>
				       	    	<th ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" size=<% =fontsize2 %> ><b>Member ID</b></FONT></th>
				            	<th ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" size=<% =fontsize2 %> ><b>Club</b></FONT></th>
        				    	<th ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" size=<% =fontsize2 %> ><b>Contact</b></FONT></th>
        				    	<th ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" size=<% =fontsize2 %> ><b>City/ST</b></FONT></th>
        				    	<th ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" size=<% =fontsize2 %> ><b>Blank</b></FONT></th><%
				    			ELSEIF LCASE(rs("Sex"))="male" OR LCASE(rs("Sex"))="female" THEN %>
				       	  	  <th ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" size=<% =fontsize2 %> ><b>Member ID</b></FONT></th>
				            	<th ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" size=<% =fontsize2 %> ><b>Last Name</b></FONT></th>
        				    	<th ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" size=<% =fontsize2 %> ><b>First Name</b></FONT></th>
        				    	<th ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" size=<% =fontsize2 %> ><b>City/ST</b></FONT></th>
        				    	<th ALIGN="Center" vAlign="top"><FONT COlOR="#FFFFFF" size=<% =fontsize2 %> ><b>Age/Gender</b></FONT></th><%
				    			END IF 
				    	    %>
 				  				</tr><%

		        			DO WHILE NOT rs.EOF 
											sMemberID = PersonIDwChkDgt(rs("PersonID"))
											sMembAge = AgeAtDate_New(Date, sMemberID)
											t=1
											IF t=2 AND sMemberID="900009554" THEN
													response.write("<br><br>sMemberID = "&sMemberID)
													response.write("<br><br>sMembAge = "&sMembAge)
													response.write("<br><br>LCASE(rs(Sex)) = "&LCASE(rs("Sex")))
													response.write("<br><br>LCASE(rs(Sex)) = "&LCASE(rs("Sex")))
											END IF		

											%>
											<tr>
											<%
				    					IF sIncludeClubs="on" AND LCASE(rs("Sex"))<>"male" AND LCASE(rs("Sex"))<>"female" THEN %>
	            			  		<TD ALIGN="Center" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>><%= sMemberID %></a></FONT></TD>
		            		  		<TD ALIGN="Center" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>>&nbsp;<%=rs("CompanyName")%></FONT></TD>
		            		  		<TD ALIGN="Center" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>>&nbsp;<%=rs("FirstName")%>&nbsp;<%=rs("LastName")%></FONT></TD>
		            		  		<TD ALIGN="Center" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>>&nbsp;<%=rs("City")&", "&rs("State")%></a></FONT></TD>
		            		  		<TD ALIGN="Center" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>>&nbsp;---</a></FONT></TD><%
				    					ELSEIF sMembAge="Unk" AND (LCASE(rs("Sex"))="male" OR LCASE(rs("Sex"))="female")  THEN %>
	            			  		<TD ALIGN="Center" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>><%= sMemberID %></a></FONT></TD>
		            		  		<TD ALIGN="Center" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>><%=rs("LastName")%></FONT></TD>
		            		  		<TD ALIGN="Center" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>><%=rs("FirstName")%></FONT></TD>
	  	          		  		<TD ALIGN="Center" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>>&nbsp;<%=rs("City")&", "&rs("State")%></a></FONT></TD>
		            		  		<TD ALIGN="Center" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>>&nbsp;<% response.write(sMembAge&"/"&rs("Sex")) %></a></FONT></TD><%
				    					ELSEIF sIncludeClubs<>"on" AND (LCASE(rs("Sex"))="male" OR LCASE(rs("Sex"))="female")  THEN %>
	            			  		<TD ALIGN="Center" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>><%= sMemberID %></a></FONT></TD>
		            		  		<TD ALIGN="Center" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>><a href="/rankings/<%=ThisFileName%>?sMemberID=<%= sMemberID %>&FormStatus=found"><%=rs("LastName")%></a></FONT></TD>
		            		  		<TD ALIGN="Center" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>><a href="/rankings/<%=ThisFileName%>?sMemberID=<%= sMemberID %>&FormStatus=found"><%=rs("FirstName")%></a></FONT></TD>
	  	          		  		<TD ALIGN="Center" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>>&nbsp;<%=rs("City")&", "&rs("State")%></a></FONT></TD>
		            		  		<TD ALIGN="Center" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>>&nbsp;<% response.write(sMembAge&"/"&rs("Sex")) %></a></FONT></TD><%
				    					END IF 
				    					%>
        	  					</tr>
											<% 
	        
											rs.MoveNext 
											
  			    			LOOP 
  			    
  			    			%>
		       			</table> 
								<br>
								<TABLE ALIGN="CENTER" style="border-style:none;" width="80%" >
			  					<tr>
	        					<TD colspan=5 ALIGN="Center" vAlign="top" style="border-style:none;" ><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>>IMPORTANT: Age and Gender must be correct for on-line entry system.</a></FONT></TD>
		    					</tr>
			  					<tr>
			  					<%  

									' --- Found more than 9 records in file - ie encourage to modify search  
									IF rs.recordcount > 9 THEN 
											%>
	    	   						<td colspan=2 align=center style="border-style:none;">
					  						<font COlOR="red" size=<% =fontsize3 %> face=<% =font1 %>>More than twenty records found.  Only the top twenty records were displayed.
					  						<br>
			        	  			Please refine your search parameters.
					  						</font>
					  						<br><br>
											</td>
											<%
									END IF 
				
									%>
			  					</tr>
									<form action="/rankings/<%=ThisFileName%>?FormStatus=search" method="post">
			  					<tr>
			    					<td width=50% align=center style="border-style:none;">
												<input type="submit" style="width:9em" value="New Search"></form>
												<input type="hidden" name="sSendingPage" value="<%=sSendingPage%>">
												<input type="hidden" name="sTourAdminPW" value="<%=sTourAdminPW%>">
												<input type="hidden" name="sTourID" value="<%=sTourID%>">
												<input type="hidden" name="adminmenulevel" value="<%=adminmenulevel%>">			
			    					</td>
 			    					<td width=50% align=center style="border-style:none;">     				
											<form action="http://usawaterski.org" method="post">
						  					<input type="submit" style="width:9em" value="Quit">
												<input type="hidden" name="sSendingPage" value="<%=sSendingPage%>">
												<input type="hidden" name="sTourAdminPW" value="<%=sTourAdminPW%>">
												<input type="hidden" name="sTourID" value="<%=sTourID%>">
												<input type="hidden" name="adminmenulevel" value="<%=adminmenulevel%>">			
								    </td>
	      					</tr>
									</form>
       					</TABLE> 
					</TABLE>
					<%    

					rs.Close
			
		ELSE
					' ******************************************************************************************************************
					' Only ONE MEMBER was found or the temporary variables have been set indicating a Member had been selected from list
					' ******************************************************************************************************************

				'response.write("<br><br>Line 1284")
				'response.end
      			
				sMemberID=PersonIDwChkDgt(rs("PersonID"))
				'sMemberID=rs("MemberID")

				rs.close
		
				' --- Pull the info from MemberTrak where the selected-searched matches MemberTrak
				sSQL = "SELECT TOP 1 * FROM "&MemberLiveTableName
   			sSQL = sSQL + " WHERE PersonID = '" & RIGHT(sMemberID,8) & "'"
   			sSQL = sSQL + " ORDER BY PersonID"
      	rs.open sSQL, sConnectionToTRATable, 3, 1

				' --- Display banner across the top with name of selected member
				sMemberID = PersonIDwChkDgt(rs("PersonID")) 
				sFirstName = rs("FirstName")
				sLastName = rs("LastName")
				sBirthDate = rs("BirthDate")

				rs.close

				DisplayMemberData

		 END IF 
	 END IF

END IF



set rs = nothing



END SUB





' ----------------------- 
   SUB DisplayMemberData
' ----------------------- 

tMemberID=sqlclean(sMemberID)

set rs=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT TOP 1 * FROM "&MemberLiveTableName
sSQL = sSQL + " WHERE PersonID = '" & RIGHT(tMemberID,8) & "'"
rs.open sSQL, sConnectionToTRATable, 3, 1

sFirstName = rs("FirstName")
sLastName = rs("LastName")
sMembCity = rs("City")
sMembState = rs("State")
sSex = rs("Sex")

sMembAge = AgeAtDate_New(DATE, tMemberID) 

%>

<br>
<TABLE class="innertable" ALIGN="CENTER" width="70%" >
  <tr>
      <th align=center><font size="4" Color="<%=TextColor5%>"><b>Confirm Member Name</b></font><br></th>
  </tr>  
  <br>

  <tr>
		<td>
     	<TABLE class="innertable" ALIGN="CENTER" width="90%" >
	 			<tr>
    	    <br>
     	    <TH ALIGN="left" vAlign="top"><FONT COlOR="#FFFFFF" size=<% =fontsize2 %> face=<% =font1 %>><b>Member ID</b></FONT></TH>
     	    <TH ALIGN="left" vAlign="top"><FONT COlOR="#FFFFFF" size=<% =fontsize2 %> face=<% =font1 %>><b>Name</b></FONT></TH>
     	    <TH ALIGN="left" vAlign="top"><FONT COlOR="#FFFFFF" size=<% =fontsize2 %> face=<% =font1 %>><b>City/ST</b></FONT></TH>
     	    <TH ALIGN="left" vAlign="top"><FONT COlOR="#FFFFFF" size=<% =fontsize2 %> face=<% =font1 %>><b>Age/Gender</b></FONT></TH>
  	  	</tr>
        <tr>	
            <TD ALIGN="left" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#0000CD" size=<% =fontsize2 %> face=<% =font1 %>><% =sMemberID %></FONT></TD>
            <TD ALIGN="left" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#0000CD" size=<% =fontsize2 %> face=<% =font1 %>><% =sFirstName&" "&sLastName %></FONT></TD>
            <TD ALIGN="left" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#0000CD" size=<% =fontsize2 %> face=<% =font1 %>><% =sMembCity&", "&sMembState %></FONT></TD>
            <TD ALIGN="left" vAlign="top" BGCOLOR="#FFFFFF"><FONT COlOR="#0000CD" size=<% =fontsize2 %> face=<% =font1 %>><% =sMembAge&"/"&sSex %></FONT></TD>
				</tr>
			</TABLE>
 			<br>
			<TABLE ALIGN="CENTER" width="80%" >
	  		<tr>
	    		<td align ="left" style="border-style:none;"> 
	    			<%	
						IF PWreq = "yes" THEN
		  					%><form action="/rankings/<%=ThisFileName%>?FormStatus=getpw" method="post"><%
						ELSE
								%><form action="/rankings/<%=ThisFileName%>?FormStatus=confirmed" method="post"><%
						END IF  
						%>
							<input type="hidden" name="sMemberID" value="<%=sMemberID%>">
							<input type="hidden" name="sSendingPage" value="<%=sSendingPage%>">
							<input type="hidden" name="sTourAdminPW" value="<%=sTourAdminPW%>">        			
							<input type="hidden" name="sTourID" value="<%=sTourID%>">
							<input type="hidden" name="adminmenulevel" value="<%=adminmenulevel%>">			
        			<input type="submit" value="Confirm This Member">
	      		</form>
					</td>	    
					<td Align="right" style="border-style:none;">			
						<form action="/rankings/<%=ThisFileName%>?FormStatus=search" method="post">
        			<input type="submit" value="Select New Member">
        			<input type="hidden" name="sSendingPage" value="<%=sSendingPage%>">
        			<input type="hidden" name="sTourAdminPW" value="<%=sTourAdminPW%>">
							<input type="hidden" name="sTourID" value="<%=sTourID%>">
							<input type="hidden" name="adminmenulevel" value="<%=adminmenulevel%>">			
        		</form>
	      		<% ' --- Formerly second form tag here --- form %>
	    		</td>
    	  </tr>	
			</TABLE>
    </td>
  </tr>
</TABLE><% 


END SUB






' ---------------------------------
   SUB DisplayMemberSearchFilters
' ---------------------------------


%>
<br><br>
<form action="/rankings/<%=ThisFileName%>?rid=<%=rid%>&FormStatus=search" method="post">
	<input type="hidden" name="pvar" value="SelectMember">
	<input type="hidden" name="sSendingPage" value="<%=sSendingPage%>">
	<input type="hidden" name="sTourAdminPW" value="<%=sTourAdminPW%>">
	<input type="hidden" name="sTourID" value="<%=sTourID%>">
	<input type="hidden" name="adminmenulevel" value="<%=adminmenulevel%>">			

<TABLE ALIGN="CENTER" class="innertable" width="70%" >
  <tr>
  	<th align=center>
			<font face=<% =font1 %> size="4" Color="<%=TextColor5%>"><b>Search for Member</b></font>
      <br>
    </th>
  </tr>  
  <tr>
		<td bgcolor="<%=TableColor1%>">
			<br>
			<TABLE Class="innertable" ALIGN="CENTER" width=95% >
      	<TR>
        	<TH ALIGN="center"><FONT COlOR="#FFFFFF" size=<% =fontsize2 %>>Member ID</FONT></Center></TH>
          <TH ALIGN="center"><FONT COlOR="#FFFFFF" size=<% =fontsize2 %>>Last Name or Club</FONT></Center></TH>
          <TH ALIGN="center"><FONT COlOR="#FFFFFF" size=<% =fontsize2 %>>First Name</FONT></Center></TH>
				</TR>
				<TR>
            <TD ALIGN="Left" bgcolor="#FFFFFF"><Center><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>><input type="text" name="Member_ID" maxlength=9 size=11></input></FONT></Center></TD>
            <TD ALIGN="Left" bgcolor="#FFFFFF"><Center><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>><input type="text" name="Last_Name" maxlength=15 size=18></input></FONT></Center></TD>
            <TD ALIGN="Left" bgcolor="#FFFFFF"><Center><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>><input type="text" name="First_Name" maxlength=15 size=18></input></FONT></Center></TD>
				</TR>
			</TABLE>

			<TABLE align=center width=95% BORDER="0">
			  <tr>	
          <td ALIGN="center" style="border-style:none;"><FONT size=<% =fontsize2 %>>Club MemberID</FONT>
						<input type="checkbox" name="sIncludeClubs" <% IF sIncludeClubs="on" THEN Response.write("Checked")%>>
	    		</td>
	  		</tr>	
	  		<tr>	
	    		<td colspan=3 align=center style="border-style:none;">
						<br>

	          <input type="submit" value="Begin Search">
	      		<br><br>
					</td>
				</tr>
			</TABLE>

     </td>
  </tr>
</TABLE>
</form>
<%


END SUB




' ------------------------------
   SUB DisplayTimeOutNotice
' ------------------------------

sSendingPage = "/rankings/defaultHQ.asp" 

%>
<br><br>

<TABLE class="droptable" ALIGN="CENTER" width=65%>
  <TR>
      <TD BGCOLOR="red"><center><font face=<% =font1 %> color="#FFFFFF" size="4"><b>Important Notice !!</b></font></TD>
  </TR>  

  <TR>
     <TD VALIGN="top">
	<TABLE BORDER="0" VALIGN="top" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="<%=TableColor1%>" width="100%">
	   <tr>
	      <td VALIGN="top" ALIGN="center">
		<br>
		<font color="<% =TextColor1 %>" face="<% =font1 %>" size="3"><b><i>Your Session Timed Out</i></b></font>
		<br><br>
		<font face="<% =font1 %>" size="1">We are sorry for the inconvenience, but you must start over.</font>
		<br><br>	
		<font face="<% =font1 %>" size="1">The inactivity caused our server to reach the maximum time limit for maintaining your member and tournament selections.  The record you were working on is no longer active. Please try again.  
		<br><br>
		If you have any questions, please contact:
		<br>
		USA Water Ski - Competition Dept at 800-533-5972</b></font>
	    </td>
	  </tr>
	<tr>
	   <td align="center">
		<br>
		<form action="<% =sSendingPage %>" method="post">
		  <center><input type="submit" value=" Continue "></center>
		</form>
		</TABLE>
		   </td>	
	</tr>
    </TD>
  </TR>
</TABLE><% 


END SUB


%>









