<!--#include virtual="/admin/MemberRegFunctions.asp"-->

<% 
If not Session("aauth") then response.redirect "Login.asp"

Dim curTraceMsg, sTourID, sTourDate, sStateSQL, sTourName, sUserName
Dim curSanctionId, curMemberId, curMemberFirstName, curMemberLastName
Dim monthNode, dayNode, yearNode, delim1, delim2
Dim memberData, ContactDtls

sStateSQL = "State IN ('')"

sUserName = session("UserName")
sTourID = Session("TournamentID")
sTourDate = session("tournamentdate")
sTourName = session("TournamentName")

curSanctionId = left(sTourID, 6)
curMemberId = PersonIDwChkDgt(Request.QUERYSTRING("PersonID"))
curMemberFirstName = ""
curMemberLastName = ""

If len(curSanctionId) < 6 THEN
    sTourDate = Date()    
    
    delim1 = instr(sTourDate, "/")
    delim2 = instr(delim1 + 1, sTourDate, "/")

    monthNode = Left(sTourDate, delim1 - 1)
    yearNode = Right(sTourDate, 4)
    dayNode = Mid(sTourDate, delim1 + 1, delim2 - delim1 - 1)
    
    curSanctionId = yearNode - 2000 & "Z000"
END IF

%>

<html>

<head>
    <title>Display One Member</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" >

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#42639F">
      <p>&nbsp;</p>
      <p align="center"><font face="Verdana" size="6" color="#FFFFFF">
      	USA Water Ski Registration Details</font></p>
      <p align="center"><font face="Verdana" size="4" color="#FFFFFF">
      	Registration Support for -- <%=session("TournamentName")%></font></p>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>  
  
<table border="0" cellspacing="0" cellpadding="0">  
    <tr> 
        <td width="185" valign="top" bgcolor="#42639F">
	        <font face="Verdana" size="2" COLOR="#FFFFFF"><br>&nbsp;Currently Logged in as: </font><br>
	        <font face="Verdana" size="2" COLOR="#FFFFFF">&nbsp;<%=Session("UserName")%>&nbsp;&nbsp;
		    <%=session("tournamentdate")%></font><br>
	        <br>

			<font face="Verdana" size="2"> 
                <br>&nbsp;<a href="logout.asp"><font face="arial" COLOR="#FFFFFF">Log Out</font></a>&nbsp;<br>
			</font>
			<br>
	        &nbsp;<a href="/admin/index.asp"><font face="arial" size="2" COLOR="#FFFFFF">Back to Admin Index</font></a><br>&nbsp;<br>
	        &nbsp;<a href="http://www.usawaterski.org"><font face="arial" size="2" COLOR="#FFFFFF">USA Water Ski Home</font></a><br>&nbsp;<br>
			<br>
            <font face="Verdana" size="1">&nbsp;<font COLOR="#FFFFFF">Powered by</font> <a href="http://www.epolk.com"><font COLOR="#FFFFFF">ePolk.com</font></a><br></font>
	    </td>

	    <td width="640" >
	        <TABLE BORDER="0">
			    <TR>
				    <td width="14">&nbsp;</td>
				    <td width="14">&nbsp;</td>
				    <td>&nbsp;</td>
			    </TR>

			    <TR>
				    <td>&nbsp;</td>
				    <td colspan=2><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
                        Here is the information for the selected member, to be copied into
					    an additional row on your Registration Template spreadsheet.&nbsp;
					    Brief instructions appear with the member's information below.&nbsp; 
					    <a href="OneMemberInstructions.htm" target="_blank">Click Here</a>
					    to display more comprehensive instructions 
					    (in a separate window).
                        <br /><br /><b>Note: </b>There is now a new feature in WSTIMS Registration Window > Add icon to directly perform a similar search and directly import to your tournament
                        <br /><br />
				        </font>
				    </td>
			    </TR>

    			<TR> 
				    <td>&nbsp;</td>
				    <td colspan=2><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
                        Hilite/copy the values below, then <b>Paste Special (text)</b> into column <b>A</b> cell on a new row, that you have inserted into your 
                        Registration Template.  If the data doesn't propagate to all the columns then you might need to 
                        use the <b>Text To Columns</b> feature in the Excel Data menu <br>&nbsp;</font>
				    </td>
	    		</TR>

			    <TR> 
				    <td>&nbsp;</td>
				    <td colspan=2>

            <%
            Dim column24, column25, column26
            Dim rsMember
            Set WaterskiConnect = Server.CreateObject("ADODB.Connection")
            WaterskiConnect.Open Application("WaterSkiConn")
            Set rsMember = Server.CreateObject("ADODB.RecordSet")
            rsMember.ActiveConnection = WaterskiConnect

            curSqlStmt = buildQueryMemberRegEntries(curSanctionId, sTourDate, sStateSQL, curMemberId, curMemberFirstName, curMemberLastName)
            rsMember.Open curSqlStmt
            Do until rsMember.EOF
	            memberData =  "<br />" & rsMember("MemberID")
	            memberData = memberData & "," & rsMember("LastName")
	            memberData = memberData & "," & rsMember("FirstName")
	            memberData = memberData & ",," & rsMember("Div")
	            memberData = memberData & "," & rsMember("Age")
	            memberData = memberData & "," & rsMember("City")
	            memberData = memberData & "," & rsMember("State")
	            memberData = memberData & ",,,,," & rsMember("SlalomRank")
	            memberData = memberData & "," & rsMember("TrickRank")
	            memberData = memberData & "," & rsMember("JumpRank")
	            memberData = memberData & "," & rsMember("SlalomRating")
	            memberData = memberData & "," & rsMember("TrickRating")
	            memberData = memberData & "," & rsMember("JumpRating")
	            memberData = memberData & "," & rsMember("OverallRating")

		        column24 = ""
                column25 = "" 
                column26 = ""
                IF rsMember("EffTo") >= cdate(sTourDate) and rsMember("CanSki") = True and rsMember("Waiver") > 0 THEN
		            column24 = "Yes"
			        column25 = "Ready"
			        column26 = FormatNumber(0,2)
		        ELSE
			        column24 = "No"
                    column25 = rsMember("MembershipRate")
                    column26 = rsMember("CostToUpgrade")

			        ' Figure applicable Renewal / Upgrade Amount based on MemType & Status
			        IF rsMember("EffTo") < cdate(sTourDate) THEN
				        IF rsMember("CanSki") = False THEN
					        column25 = "Needs Renew/Upgrade"
					        column26 = rsMember("MembershipRate")
				        ELSE
					        column25 = "Needs Renew"
					        column26 = rsMember("MembershipRate")
				        END IF
			        ELSE
				        IF rsMember("CanSkiGR") = True THEN
					        column25 = "** Grass Roots Only"
                            column26 = rsMember("CostToUpgrade")
				        ELSEIF rsMember("CanSki") = False THEN
					        column25 = "Needs Upgrade"
					        column26 = rsMember("MembershipRate")
				        ELSE
					        column25 = "Needs Annual Waiver"
					        column26 = FormatNumber(0,2)
				        END IF
			        END IF
		        END IF

	            memberData = memberData & ",,,,,," & column24
	            memberData = memberData & "," & column25
	            memberData = memberData & "," & column26
	            memberData = memberData & "," & rsMember("EffTo")

	            memberData = memberData & ",,,,," & rsMember("JudgeSlalom")
	            memberData = memberData & "," & rsMember("JudgeTrick")
	            memberData = memberData & "," & rsMember("JudgeJump")

	            memberData = memberData & "," & rsMember("DriverSlalom")
	            memberData = memberData & "," & rsMember("DriverTrick")
	            memberData = memberData & "," & rsMember("DriverJump")

	            memberData = memberData & "," & rsMember("ScorerSlalom")
	            memberData = memberData & "," & rsMember("ScorerTrick")
	            memberData = memberData & "," & rsMember("ScorerJump")

	            memberData = memberData & "," & rsMember("Safety")
	            memberData = memberData & "," & rsMember("TechController")

            	rsMember.MoveNext
            Loop

            rsMember.Close
            Set rsMember = Nothing
    	    %>

                        <TABLE BORDER="1" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#DDDDFF">
                            <TR>
					            <td VALIGN="Left">
                                    <font size="2"><pre><%=memberData %></font>
					            </td>
                            </TR>
            			</TABLE>
                    
                        <TABLE ALIGN="CENTER" WIDTH=70%>
                            <TR><TD>&nbsp;</TD></TR>

                            <TR>
                                <TD width=25% align=center>
		                            <form action="LookupMembers.asp?FormStatus=newsearch" method="post">
		                                <input type="submit" style="width:9em" value="New Search">
		                            </form>
	                            </TD>

                                <td width=25% align=center>     				
		                            <form action="Index.asp" method="post">
	                                    <input type="submit" style="width:9em" value="Quit">
		                            </form>
                                </td>

	                            <td width=25% align=center valign=center><font size="2" face="Verdana, Arial, Helvetica, sans-serif">     				
		                            <form action="OneMemberInstructions.htm" method="post" target="_blank">
		                                <input type="submit" style="width:9em" value="Instructions">
		                            </form>
                                </td>
                            </TR>

                        </TABLE>
                        <%                    
                        '	Finally Display additional contact details for selected member, if authorized

                        IF left(Session("UserName"),1) > "9" THEN
                            ContactDtls = "Contact details will go here"

                            Dim rsContact
                            Set rsContact = Server.CreateObject("ADODB.RecordSet")
                            rsContact.ActiveConnection = WaterskiConnect

                            curSqlStmt = "SELECT Mem.PersonID, Mem.LastName, Mem.FirstName, Mem.City," 
                            curSqlStmt = curSqlStmt & " Mem.State, Left(Mem.Sex,1) as Gender, Mem.EffectiveTo as ExpDt," 
                            curSqlStmt = curSqlStmt & " Mem.BirthDate, Convert(char(10),Mem.Birthdate,111) as Bdate,"
                            curSqlStmt = curSqlStmt & " Mem.DivisionCode1 + '/' + Mem.DivisionCode2 as SptsDiv,"
                            curSqlStmt = curSqlStmt & " Mem.Address1, Mem.Address2, Mem.Zip, Mem.FederationCode as FedCode," 
                            curSqlStmt = curSqlStmt & " Mem.ForeignFederationID as ForFedID, Mem.Email,"
                            curSqlStmt = curSqlStmt & " Mem.Phone as HomePhone, Mem.BusinessPhone as BusPhone, Mem.MobilePhone,"
                            curSqlStmt = curSqlStmt & " Mem.MembershipTypeCode as MemType, Mem.WaiverStatusID as Waiver, Typ.TypeCode,"
                            curSqlStmt = curSqlStmt & " Typ.CanSkiInTournaments as CanSki, Typ.CanSkiInGRTournaments as CanSkiGR"
                            curSqlStmt = curSqlStmt & " FROM USAWaterski.dbo.memberslive as Mem, USAWaterski.dbo.MembershipTypes as Typ"
                            curSqlStmt = curSqlStmt & " WHERE Mem.MembershipTypeCode = Typ.MemberShipTypeID"
                            curSqlStmt = curSqlStmt &+ " AND Mem.PersonID='" & Request("PersonID") & "';" 

                            rsContact.Open curSqlStmt

                            If rsContact.EOF THEN
                                ContactDtls = ""
                            ELSE
	                            ContactDtls = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Member Name:&nbsp;&nbsp; " & rsContact("firstname") & " " & rsContact("lastname") 
	                            ContactDtls = ContactDtls & "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Address:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; " & rsContact("Address1") & "&nbsp; " & rsContact("Address2")
	                            ContactDtls = ContactDtls & "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; City/ST/Zip:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; " & rsContact("city") & ", " & rsContact("State") & "&nbsp; " & rsContact("Zip")
	                            ContactDtls = ContactDtls & "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Birthdate:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; " & rsContact("BDate") & "&nbsp;&nbsp;&nbsp;&nbsp; Gender: &nbsp;" & Left(rsContact("Gender"),1)
	                            ContactDtls = ContactDtls & "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Federation:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; " & rsContact("FedCode") & "&nbsp;&nbsp;&nbsp; Fed ID #: " & rsContact("ForFedID")
	                            ContactDtls = ContactDtls & "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Member Type:&nbsp;&nbsp;&nbsp; " & rsContact("TypeCode") & "&nbsp;&nbsp;&nbsp; SptsDiv: " & rsContact("SptsDiv")
	                            ContactDtls = ContactDtls & "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; E-Mail Adrs:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; " & rsContact("EMail")
	                            ContactDtls = ContactDtls & "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Home Phone:&nbsp;&nbsp;&nbsp;&nbsp; " & rsContact("HomePhone")
	                            ContactDtls = ContactDtls & "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Bus Phone:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; " & rsContact("BusPhone")
	                            ContactDtls = ContactDtls & "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Mobile Phone:&nbsp;&nbsp;&nbsp; " & rsContact("MobilePhone")
                            END IF

                            rsContact.Close
                            Set rsContact = Nothing

                            %>
                            <TABLE ALIGN="CENTER" WIDTH=70%>
                                <TR><TD>&nbsp;</TD></TR>

                			    <TR>
                                    <td>&nbsp;</td>
				                    
                                    <td colspan=2>
                                        <font size="2" face="Verdana, Arial, Helvetica, sans-serif">
                                        <br>Additional Contact Details for this Member:<br>&nbsp;<br>
				                        <%=ContactDtls%></font>
                                    </td>
                                </TR>
                            </TABLE>
                            <%
                        END IF
                        %>
                    </td>
                </TR>
            
            </TABLE>
        </td>
    </TR>

</TABLE>
	
<%
WaterskiConnect.Close
%>
</body>
</html>





