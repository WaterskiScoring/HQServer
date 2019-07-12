<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->


<%

Dim currentPage, rowCount, i
Dim sMonth, sSendingPage, SendingURL, sTourID, TStatus, STourRange
Dim MainImage


NewsPageNum = "10"
SkiYear = "2007"

TStatus = TRIM(Request("TStatus"))
sTourID = TRIM(Request("sTourID"))
sSendingPage = TRIM(Request("sSendingPage"))
sTourRange = TRIM(Request("Tour_Range"))
process=TRIM(Request("process"))





adminmenulevel = Session("adminmenulevel")
IF TRIM(adminmenulevel) = "" THEN adminmenulevel = 1



' Is this NEEDED?
IF TStatus="Quit" THEN 
	' User pressed Select New tournament 
	Pvar="TourInfo"
END IF	

' --- Sets what the filter does and (MARK TO DO - what the button displays)
IF InStr(LCASE(sSendingPage), "new") > 0 THEN 
	TStatus="TourInfo"
	sTourRange = "2"
	IF process = "register"  THEN 
		sTourRange = "2"
	ELSEIF process = "tourmgr" THEN 
		sTourRange = "2"
	ELSEIF process = "viewreg" THEN 
		sTourRange = "3"
	ELSE
		sTourRange = "1"
	END IF
END IF





IF TStatus = "TourInfo" THEN
	TourSearch
END IF 

IF TStatus = "Details" THEN
	TourDetail
END IF


IF TStatus="Confirmed" THEN 

	Session("sTourID") = sTourID
	sSendingPage = Session("sSendingPage") + "&sTourID="&sTourID
	response.redirect(sSendingPage)
END IF	






' -------------------------- END OF MAIN PROGRAM  ---------------------




' ---------------------------- Display Tour Information --------------------------------


' ---------------------
  SUB TourDetail
' ---------------------


'WriteIndexPageHeader
HQHead1

sTourID = TRIM(Request("sTourID"))
sTourRange = TRIM(Request("sTourRange"))


sSQL = "SELECT TOP 1 * from " & SanctionTableName & " LEFT JOIN "&GuideBookTableName&" ON "&SanctionTableName&".TournAppID = "&GuideBookTableName&".GTournAppID WHERE LEFT("&SanctionTableName&".TournAppID,6) = '" & SQLClean(left(sTourID,6)) & "'"


set rs=Server.CreateObject("ADODB.recordset")

IF TRIM(Request("SkiYear")) <> "" THEN
	Session("SkiYear") = TRIM(Request("SkiYear"))
END IF
rs.open sSQL, SConnectionToTRATable



IF rs.EOF THEN
	WriteLog(date() &"  "& time() &" *** ERROR *** -  User tried to view Tour Info for tournament '"& sTourID &"' but that tour was not found in SWIFT.")
	Response.write ("<br><br><br><font color=red>Tournament Not Found</font>")
ELSE  %>


      <TABLE class="droptable" ALIGN="center" BGCOLOR="<% =tablecolor1 %>" BORDER="3" CELLPADDING="2" CELLSPACING="4" BGCOLOR="#FFFFFF" width=600>
	<tr><td>

	<TABLE ALIGN="center" BGCOLOR="#FFFFFF" BORDER="1" CELLPADDING="2" CELLSPACING="1" BGCOLOR="#FFFFFF" width=100%>
	<tr>
  	  <TD Align="Center" ColSpan="3" BGCOLOR="<%=HQSiteColor2%>" vAlign="top"><FONT COlOR="#000000" size="4" face=<% =font1 %>><% Response.Write("<b>" & rs("TName") & "</b>") %></FONT></td>
	</tr>
		
 	<tr>
	  <TD ColSpan="3" ALIGN="center"><FONT COlOR="#000000" size="3" face=<% =font1 %>><% IF rs("TDateS") = rs("TDateE") THEN Response.Write ("&nbsp;&nbsp;&nbsp;" & rs("TDateS")) ELSE Response.Write ("&nbsp;&nbsp;&nbsp;" & rs("TDateS") & " to " & rs("TDateE")) %></FONT>
	   <%
	   NoteOn="Y"
	   IF NoteOn = "Y" THEN
		IF LEFT(rs("TSanction"),3) = LEFT(rs("TournAppID"),3) THEN 
			%><right><FONT COLOR="Red" size="2" face=<% =font1 %>>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; ACTIVE </FONT></right></TD><% 
		ELSE
			%></td><%
		END IF
	   END IF %>
  
       	</tr>
	<tr><td colspan=3>&nbsp;</td></tr>

	<form action="/rankings/search-tourHQ.asp?rid=<%=rid%>&TStatus=Confirmed&sTourID=<%=sTourID%>" method="post">
        <TR Align="left" valign="center">
	  <td>&nbsp;</td>


	  <td ALIGN="center"><%
	    IF sTourRange="2" THEN  %>
			<input type="hidden" name="sTourRange" value="<%=sTourRange%>">
	        	<input style="width:10em" type="submit" value="Continue" >
		<%
	    END IF %>

          </td>
        </form>

         <form action="/rankings/search-tourHQ.asp?rid=<%=rid%>&TStatus=TourInfo&Tour_Range=<%=sTourRange%>" method="post">
          <td ALIGN="center">
              <input style="width:10em" type="submit" value="Change Tournament" >

	  </td>
            </form> 

	<TR>
           <TD vAlign="top"><FONT COLOR="#000000" size=<% =fontsize1 %> face=<% =font1 %>><b>Description:</b></FONT></TD>
           <TD ColSpan="2" vAlign="top">
		<%

	     IF TRIM(rs("TDescription"))<>"" THEN
		ThisDescription = TRIM(rs("TDescription"))
	     ELSEIF TRIM(rs("FDescription"))<>"" THEN
		ThisDescription = rs("FDescription")
	     ELSEIF TRIM(rs("WDescription"))<>"" THEN
		ThisDescription = rs("WDescription")
	     ELSEIF TRIM(rs("KDescription"))<>"" THEN
		ThisDescription = rs("KDescription")
	     ELSEIF TRIM(rs("CDescription"))<>"" THEN 
		ThisDescription = rs("CDescription")
	     END IF %>	

		<FONT COlOR="<% =textcolor2 %>" size="<% =fontsize1 %>"><%=ThisDescription%></font>
	  </TD>
        </TR>
       <TR>
           <TD width="150" vAlign="top"><FONT COLOR="#000000" size=<% =fontsize1 %> face=<% =font1 %>><b>Divisions:</b></FONT></TD>
           <TD  ColSpan="2"width="400" vAlign="top"><FONT COLOR="<% =textcolor2 %>" size=<% =fontsize1 %> face=<% =font1 %>><% Response.Write (rs("TDvOffered")) %></FONT></TD>
       </tr>
       <TR>
         <TD vAlign="top"><FONT COLOR="#000000" size=<% =fontsize1 %> face=<% =font1 %>><b>Sponsor:</b></FONT></TD>
         <TD ColSpan="2" vAlign="top"><FONT COLOR="<% =textcolor2 %>" size=<% =fontsize1 %> face=<% =font1 %>><% Response.Write (rs("TSponsor")) %></FONT></TD>
       </tr>
       <TR>
         <TD vAlign="top"><FONT COLOR="#000000" size=<% =fontsize1 %> face=<% =font1 %>><b>Site:</b></FONT></TD>
         <TD ColSpan="2" vAlign="top"><FONT COLOR="<% =textcolor2 %>" size=<% =fontsize1 %> face=<% =font1 %>><% Response.Write (rs("TSite")) %></FONT></TD>
       </tr>
       <TR>
         <TD vAlign="top"><FONT COLOR="#000000" size=<% =fontsize1 %> face=<% =font1 %>><b>Location:</b></FONT></TD>
         <TD ColSpan="2" vAlign="top"><FONT COLOR="<% =textcolor2 %>" size=<% =fontsize1 %> face=<% =font1 %>><% Response.Write (rs("TCity") & ", " & rs("TState")) %></FONT></TD>
       </tr>
       <TR>
         <TD vAlign="top"><FONT COLOR="#000000" size=<% =fontsize1 %> face=<% =font1 %>><b>Directions:</b></FONT></TD>
         <TD ColSpan="2" vAlign="top"><FONT COLOR="<% =textcolor2 %>" size=<% =fontsize1 %> face=<% =font1 %>><% Response.Write (rs("GTSDirections")) %></FONT></TD>
       </tr>
       <tr>    
         <td align="left" colspan="3"><hr width="550"></td>
       </tr>
       <TR>
         <TD vAlign="top"><FONT COLOR="#000000" size=<% =fontsize1 %> face=<% =font1 %>><b>Entry:</b></FONT></TD>
         <TD ColSpan="2" vAlign="top"><FONT COLOR="<% =textcolor2 %>" size=<% =fontsize1 %> face=<% =font1 %>><% IF rs("TOpenClosed") THEN Response.Write ("Closed") ELSE Response.Write ("Open") %></FONT></TD>
       </tr>
       <TR>
         <TD vAlign="top"><FONT COLOR="#000000" size=<% =fontsize1 %> face=<% =font1 %>><b>Tow Boat:</b></FONT></TD>
         <TD ColSpan="2" vAlign="top"><FONT COLOR="<% =textcolor2 %>" size=<% =fontsize1 %> face=<% =font1 %>><% IF rs("TTowBoatClosed") THEN Response.Write ("Closed") ELSE Response.Write ("Open") %></FONT></TD>
       </tr>
       <tr>    
         <td align="left" colspan="3"><hr width="550"></td>
       </tr>
       <TR>
         <TD vAlign="top"><FONT COLOR="#000000" size=<% =fontsize1 %> face=<% =font1 %>><b>Entry Limit:</b></FONT></TD>
         <TD ColSpan="2" vAlign="top"><FONT COLOR="<% =textcolor2 %>" size=<% =fontsize1 %> face=<% =font1 %>><% Response.Write (rs("TEntryLimit")) %></FONT></TD>
       </tr>
       <TR>
         <TD vAlign="top"><FONT COLOR="#000000" size=<% =fontsize1 %> face=<% =font1 %>><b>Entry Fees:</b></FONT></TD>
         <TD ColSpan="2" vAlign="top"><FONT COLOR="<% =textcolor2 %>" size=<% =fontsize1 %> face=<% =font1 %>><% Response.Write (rs("TEntryFees")) %></FONT></TD>
       </tr>
       <TR>
         <TD vAlign="top"><FONT COLOR="#000000" size=<% =fontsize1 %> face=<% =font1 %>><b>Entry Deadline:</b></FONT></TD>
         <TD ColSpan="2" vAlign="top"><FONT COLOR="<% =textcolor2 %>" size=<% =fontsize1 %> face=<% =font1 %>><% Response.Write (rs("TLateDate")) %></FONT></TD>
       </tr>
       <TR>
         <TD vAlign="top"><FONT COLOR="#000000" size=<% =fontsize1 %> face=<% =font1 %>><b>Late Fee:</b></FONT></TD>
         <TD ColSpan="2" vAlign="top"><FONT COLOR="<% =textcolor2 %>" size=<% =fontsize1 %> face=<% =font1 %>><% Response.Write (rs("TLateFee")) %></FONT></TD>
       </tr>
       <tr>    
         <td align="left" colspan="3"><hr width="550"></td>
       </tr>
       <TR>
         <TD vAlign="top"><FONT COLOR="#000000" size=<% =fontsize1 %> face=<% =font1 %>><b>Send Entries To:</b></FONT></TD>
         <TD ColSpan="2" vAlign="top"><FONT COLOR="<% =textcolor2 %>" size=<% =fontsize1 %> face=<% =font1 %>><% Response.Write (rs("TRegistrarName") & "<br>" & rs("TRegistrarAddr") & "<br>" & rs("TRegistrarCity") & ", " & rs("TRegistrarState") & "  " & rs("TRegistrarZip")) %></FONT></TD>
       </tr>
       <tr>    
         <td align="left" colspan="3"><hr width="550"></td>
       </tr>
       <TR>
         <TD vAlign="top"><FONT COLOR="#000000" size=<% =fontsize1 %> face=<% =font1 %>><b>Accommodations:</b></FONT></TD>
         <TD ColSpan="2" vAlign="top"><FONT COLOR="<% =textcolor2 %>" size=<% =fontsize1 %> face=<% =font1 %>><% Response.Write (rs("GTAccommodation")) %></FONT></TD>
       </tr>
       <TR>
         <TD vAlign="top"><FONT COLOR="#000000" size=<% =fontsize1 %> face=<% =font1 %>><b>Awards:</b></FONT></TD>
         <TD ColSpan="2" vAlign="top"><FONT COLOR="<% =textcolor2 %>" size=<% =fontsize1 %> face=<% =font1 %>><% Response.Write (rs("GTAwards")) %></FONT></TD>
       </tr>
       <TR>
         <TD vAlign="top"><FONT COLOR="#000000" size=<% =fontsize1 %> face=<% =font1 %>><b>Practice:</b></FONT></TD>
         <TD ColSpan="2" vAlign="top"><FONT COLOR="<% =textcolor2 %>" size=<% =fontsize1 %> face=<% =font1 %>><% Response.Write (rs("GTPractice")) %></FONT></TD>
       </tr>
       <TR>
         <TD vAlign="top"><FONT COLOR="#000000" size=<% =fontsize1 %> face=<% =font1 %>><b>Start Time:</b></FONT></TD>
         <TD ColSpan="2" vAlign="top"><FONT COLOR="<% =textcolor2 %>" size=<% =fontsize1 %> face=<% =font1 %>><% Response.Write (rs("GTStartTime")) %></FONT></TD>
       </tr>
       <TR>
         <TD vAlign="top"><FONT COLOR="#000000" size=<% =fontsize1 %> face=<% =font1 %>><b>Schedule of Events:</b></FONT></TD>
         <TD ColSpan="2" vAlign="top"><FONT COLOR="<% =textcolor2 %>" size=<% =fontsize1 %> face=<% =font1 %>><% Response.Write (rs("GTSofE")) %></FONT></TD>
       </tr>
       <TR>
         <TD vAlign="top"><FONT COLOR="#000000" size=<% =fontsize1 %> face=<% =font1 %>><b>Comments:</b></FONT></TD>
         <TD ColSpan="2" vAlign="top"><FONT COLOR="<% =textcolor2 %>" size=<% =fontsize1 %> face=<% =font1 %>><% Response.Write (rs("GTComments")) %></FONT></TD>
       </tr>
     </TABLE>
     </td></tr>	
   </TABLE>

     <%

SET rs=nothing

END IF  

WriteIndexPageFooter


END SUB



' -------------------------
  SUB TourSearch
' -------------------------

' ---------------------------- Display Tour Listing and Search Filters --------------------------------


'WriteIndexPageHeader
HQHead1

Dim Date1Good, Date2Good
Date1Good = 0
Date2Good = 0



sSptsGrpID = TRIM(Request("SptsGrpSelect"))
sTourState = TRIM(Request("State"))
sTourDate1 = TRIM(Request("Tour_Date1"))
sTourDate2 = TRIM(Request("Tour_Date2"))
sTourRegion = TRIM(Request("Region"))
'sTourRange = TRIM(Request("Tour_Range"))



SELECT CASE sSptsGrpID
  CASE "AWS"
	sl="on"	
  CASE "NCW"
	ju="on"	
  CASE "USW"
	wb="on"	
END SELECT


SetEventImage


sMonth = 0
%>


<br>
<form action="/rankings/search-tourHQ.asp?rid=<%=rid%>&TStatus=TourInfo" method="post">

<TABLE class="droptable" WIDTH="<%=TourTableWidth%>px" height=225px background="<%=MainImage%>">
  <TR>
    <TD>


<table ALIGN="center" width=100%>

<tr><td colspan=6 align="left"><font size=4 color="<%=textcolor2%>"><B><I>&nbsp;&nbsp;&nbsp;Search For A Tournament</I></B></font>
</td></tr>

<% 
IF Session("adminmenulevel")>=40 THEN %> 
<tr>
   <td width="120px" align="right">
     <font size=<% =fontsize2 %> >Select View:&nbsp;</font>
   </td>

   <td>
	<select name='Tour_Range'>
	<option value="0"<%If sTourRange = "0" Then Response.Write(" selected ")%>>Future Active & Pending</option>
	<option value="1"<%If sTourRange = "1" Then Response.Write(" selected ")%>>Future - Active</option>
	<option value="2"<%If sTourRange = "2" Then Response.Write(" selected ")%>>On Line Registration Available</option>
	<option value="3"<%If sTourRange = "3" Then Response.Write(" selected ")%>>Previous Year</option>
  	</select>
  </td>
</tr>

<%
END IF
%>


<tr>
<td align="right" width="120px">
<font size=<% =fontsize2 %> face=<% =font1 %>>Sports Division:&nbsp;</font> 
</td>

<td>
<SELECT name="SptsGrpSelect">
            <option value=""<%IF sSptsGrpID = "" THEN Response.Write(" SELECTed ")%>>All Sports</option>
            <option value="AWS"<%IF sSptsGrpID = "AWS" THEN Response.Write(" SELECTed ")%>>Water Skiing</option>
            <option value="USW"<%IF sSptsGrpID = "USW" THEN Response.Write(" SELECTed ")%>>Wakeboard</option>
            <option value="AKA"<%IF sSptsGrpID = "AKA" THEN Response.Write(" SELECTed ")%>>Kneeboard</option>
            <option value="NCW"<%IF sSptsGrpID = "NCW" THEN Response.Write(" SELECTed ")%>>Collegiate *</option>
</SELECT>
</td>
</tr>


<tr>
<td align="right" width="120px">
<font size=<% =fontsize2 %> face=<% =font1 %>>Region:&nbsp;</font> 
</td>

<td>
<SELECT name="Region">
            <option value=""<%IF sTourRegion = "" THEN Response.Write(" SELECTed ")%>>All Regions</option>
            <option value="C"<%IF sTourRegion = "C" THEN Response.Write(" SELECTed ")%>>S. Central</option>
            <option value="M"<%IF sTourRegion = "M" THEN Response.Write(" SELECTed ")%>>Midwest</option>
            <option value="W"<%IF sTourRegion = "W" THEN Response.Write(" SELECTed ")%>>West</option>
            <option value="S"<%IF sTourRegion = "S" THEN Response.Write(" SELECTed ")%>>South</option>
            <option value="E"<%IF sTourRegion = "E" THEN Response.Write(" SELECTed ")%>>East</option>
</SELECT>
</td>
</tr>




<tr>
<td align="right" width="120px">
<font size=<% =fontsize2 %> face=<% =font1 %>>State:&nbsp;</font> 
  <td colspan=1 valign=top align="left"><%
    StateArray = Split(USStatesList3,",")  %>
    <select name="State"><%
      FOR kvar = 0 TO UBOUND(StateArray)
        IF TRIM(sTourState) = TRIM(StateArray(kvar)) THEN
	  response.write("<option value = """&sTourState&""" SELECTED>"&sTourState&"</option>")
        ELSE
	  response.write("<option value = """&StateArray(kvar)&""">"&StateArray(kvar)&"</option>")
        END IF
      NEXT  %>
    </select>
   </td>
</tr>




<tr><td align="right">
<% 
IF Date1Good = 0 OR Date2Good = 0 THEN
       	Response.Write ("<font color=Red>")
      END IF 

%>

<font color="000000" size=<% =fontsize2 %> face=<% =font1 %>>&nbsp;&nbsp;Tournament Date:&nbsp;</font> 
</td>

<td>
<input type="text" name="Tour_Date1" value="<%=sTourDate1%>" color="000000" size=10><font color="000000" size=<% =fontsize2 %> face=<% =font1 %>> to </font>
<input type="text" name="Tour_Date2" value="<%=sTourDate2%>" size=10> <small><font color="000000" size=<% =fontsize2 %> face=<% =font1 %>>(mm/dd/yyyy)</font></small>
<% 
IF Date1Good = 0 OR Date2Good = 0 THEN
       	Response.Write ("</font>")
END IF
%>
</td></tr>

<tr><td colspan=2 align="center">
<br>
<input type="submit" value="Start Search">
</form>
</td></tr>
</table>


    </td>
  </TR>
</TABLE>
<%


' --- Runs the query against SWIFT ----
RunTournamentQuery






' No TOURNAMENTS FOUND
' --------------------	
IF rs.eof THEN
	%>
        <br><br>
	<center><font size="3" color=red><b><i>No Tournaments Found - Change Settings and Press Start Search</i></b></font></center><br>
        <%
ELSE
        %>
        <TABLE class="innertable" width="<%=TourTableWidth%>">
        <tr>
          <Th ALIGN="Center" ><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>><br>Date(s)</FONT></th>
          <th ALIGN="Center" ><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>>Tournament Name (ID)<br>Event Info</FONT></th>
          <th ALIGN="Center" colspan=20% ><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>>City<br>State</FONT></th>
	  <th ALIGN="Center" ><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>><br>Comment </FONT></th>
	</tr>
	<% 

	DO WHILE NOT rs.EOF 
            	IF sMonth <> Month(rs("TDateS")) THEN %>
              		<tr>
                	<td ALIGN="Left" vAlign="top" colspan=100% <%
                	SELECT CASE Month(rs("TDateS"))
                		CASE 1,5,9
                  			Response.Write (" BGCOLOR=""#EEDDDD"">")
                		CASE 2,6,10
                  			Response.Write (" BGCOLOR=""#CCCCFF"">")
                		CASE 3,7,11
                  			Response.Write (" BGCOLOR=""#FFFF66"">")
                		CASE 4,8,12
                  			Response.Write (" BGCOLOR=""#99FF66"">")
                	END SELECT %>

			<font size=<% =fontsize2 %> face=<% =font1 %>><% Response.Write(MonthName(Month(rs("TDateS")))) %></font>
             		 </td>
              		</tr> <%
              		
			sMonth = Month(rs("TDateS")) 
            	END IF
            	%>
            <tr>
            <TD height="40" width="20" vAlign="middle" <%
              		SELECT CASE Month(rs("TDateS"))
                		CASE 1,5,9
                  			Response.Write (" BGCOLOR=""#EEDDDD"">")
                		CASE 2,6,10
                  			Response.Write (" BGCOLOR=""#CCCCFF"">")
                		CASE 3,7,11
                  			Response.Write (" BGCOLOR=""#FFFF66"">")
                		CASE 4,8,12
                  			Response.Write (" BGCOLOR=""#99FF66"">")
              		END SELECT
            		%>  
	    
            	<font size=<% =fontsize2 %> face=<% =font1 %>><% IF rs("TDateS") = rs("TDateE") THEN Response.Write (rs("TDateS")) ELSE Response.Write (left(rs("TDateS"), len(TRIM(rs("TDateS"))) - 5) & "&nbsp;to " & rs("TDateE")) %></a></FONT></TD><% 


		' -----------------------   MARK - define the relation and activity between TourAppID and TSanction  ------------------


'		Response.write("LEFT TourAppID = "&LEFT(rs("TournAppID"), 6))
'		Response.write("LEFT TourID ="&LEFT(rs("TourID"), 6))
'		Response.write("sTourRange = "&sTourRange)

'		IF LEFT(rs("TournAppID"), 6) <> LEFT(rs("TourID"), 6) THEN

		SELECT CASE sTourRange

		    CASE "2"
			' Tournaments with OnLine Registration %>
			<TD ALIGN="Center" height="40" vAlign="middle"><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>><a href="/rankings/search-tourHQ.asp?rid=<%=rid%>&TStatus=Details&sTourID=<%=rs("TournAppID")%>&sTourRange=<%=sTourRange%>"><% Response.Write (rs("TName")  &nbsp&nbsp&nbsp&  " (" & rs("TournAppID") & ")") %></a></font><br>
               		<font color="000000" size=0 face=<% =font1 %>><% Response.Write (rs("TDescription")) %></font></TD>
	            	<TD ALIGN="Center" colspan=20% height="40" vAlign="middle"><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>><% Response.Write (rs("TCity") & "<br>" & ucase(rs("TState"))) %></a></FONT></TD>
			<TD ALIGN="Center" height="40" vAlign="middle"><FONT COLOR="<% =textcolor3 %>" size=<% =fontsize1 %> face=<% =font1 %>>Register</a></FONT></TD><% 

		    CASE ELSE %>
	    		<TD ALIGN="Center" height="40" vAlign="middle"><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>><% Response.Write (rs("TName")  &nbsp&nbsp&nbsp&  " (" & rs("TournAppID") & ")") %></a></font><br>
        	    	<% Response.Write ("<font size=""-1"">" & rs("TDescription") & "</font>") %></TD>
       	    	    	<TD ALIGN="Center" colspan=20% height="40" vAlign="middle"><FONT COlOR="#000000" size=<% =fontsize2 %> face=<% =font1 %>><% Response.Write (rs("TCity") & "<br>" & ucase(rs("TState"))) %></a></FONT></TD>
			<td>&nbsp</td><%

		END SELECT %>

            </tr>

            <% rs.MoveNext %>
        <%
	LOOP
	%>


        </table>
        </center>
        <%


	rs.close

END IF

WriteIndexPageFooter



END SUB



' ------------------------
   SUB RunTournamentQuery
' ------------------------



sSQL = "SELECT "&SanctionTableName&".TournAppID, "&SanctionTableName&".TSanction, "&SanctionTableName&".TName, "&SanctionTableName&".TCity"  
sSQL = sSQL + ", "&SanctionTableName&".TDateS, "&SanctionTableName&".TDateE, "&SanctionTableName&".TName, "&SanctionTableName&".TDescription"
sSQL = sSQL + ", "&SanctionTableName&".SptsGrpID AS SprtsGrpID, "&SanctionTableName&".TState"

sSQL = sSQL + ", "&TourGenTableName&".TourID"

sSQL = sSQL + " FROM " &SanctionTableName&" "


SELECT CASE sTourRange

	' All Active and Pending Tournaments

	' Active Tournaments
	CASE "1"
		sSQL = sSQL + " LEFT OUTER JOIN "&TourGenTableName&" ON LEFT("&SanctionTableName&".TournAppID, 6) = LEFT("&TourGenTableName&".TourID, 6)"
		sSQL = sSQL + " WHERE LEFT("&SanctionTableName&".TSanction,3) = LEFT("&SanctionTableName&".TournAppID,3) AND "
		sSQL = sSQL + " (TDateE >= '" & Date() & "')"

	' Tournaments with OnLine Registration
	CASE "2" 
		sSQL = sSQL + " LEFT OUTER JOIN "&TourGenTableName&" ON LEFT("&SanctionTableName&".TournAppID, 6) = LEFT("&TourGenTableName&".TourID, 6)"
		sSQL = sSQL + " WHERE LEFT("&SanctionTableName&".TournAppID, 6) = LEFT("&TourGenTableName&".TourID, 6)"
		IF adminmenulevel>=19 THEN
		
		ELSE
			sSQL = sSQL + " AND (TDateE >= '" & Date() & "')"
		END IF

	' Year Ago Tournaments
	CASE "3"
		sSQL = sSQL + " LEFT OUTER JOIN "&TourGenTableName&" ON LEFT("&SanctionTableName&".TournAppID, 6) = LEFT("&TourGenTableName&".TourID, 6)"
		Dim YearAgo
		YearAgo = DateAdd("d", -365, Date()) 
		sSQL = sSQL + " WHERE ("&SanctionTableName&".TDateE >= '"&YearAgo&"') AND ("&SanctionTableName&".TDateE <= '"&Date()&"')" 

	CASE ELSE
		sSQL = sSQL + " LEFT OUTER JOIN "&TourGenTableName&" ON LEFT("&SanctionTableName&".TournAppID, 6) = LEFT("&TourGenTableName&".TourID, 6)"
		sSQL = sSQL + " WHERE "
		sSQL = sSQL + " (TDateE >= '" & Date() & "')"


END SELECT



' ----- Restricts display of anyone who is NOT logged in to Nationals only ----

IF adminmenulevel <= 50 THEN
	sSQL = sSQL + " AND LEFT("&TourGenTableName&".TourID, 6) = '07W999'"
END IF


' ----- Here is the start of the filter for Sport Group

IF sSptsGrpID <> "" THEN
	IF sSptsGrpID = "NSL" THEN
		sSQL = sSQL + " AND LOWER("&SanctionTableName&".SptsGrpID) = 'AWS' and lower(right(TSanction,1)) in ('f','n','i')"
	ELSE
		sSQL = sSQL + " AND LOWER("&SanctionTableName&".SptsGrpID) = '"& sqlclean(lcase(sSptsGrpID)) & "'"
	END IF
END IF


IF LCASE(sTourState) <> "all" AND TRIM(sTourState)<>"" THEN
      	sSQL = sSQL + " AND LOWER(TState) = '" & sqlclean(lcase(sTourState)) & "'"
END IF


IF sTourRegion <> "" THEN
	sSQL = sSQL + " AND LOWER(RIGHT(LEFT(TournAppID,3),1)) = '" & sqlclean(lcase(sTourRegion)) & "'"
END IF
        
IF sTourDate1 <> "" THEN
'	sSQL = sSQL + " AND (TDateE >= '" & sTourDate1 & "' OR TDateS >= '" & sTourDate1 & "')"
      	IF sTourDate2 <> "" THEN
'       		sSQL = sSQL + " AND "
       	END IF
END IF
        
IF sTourDate2 <> "" THEN
'       	sSQL = sSQL + "(TDateE <= '" & sTourDate2 & "' OR TDateS <= '" & sTourDate2 & "')"
END IF


sSQL = sSQL + " ORDER BY TDateS "

OpenCon
set rs=Server.CreateObject("ADODB.recordset")


IF TRIM(Request("SkiYear")) <> "" THEN
	Session("SkiYear") = TRIM(Request("SkiYear"))
END IF



rs.open sSQL, SConnectionToTRATable


END SUB










Sub WriteHeaders(sTitle)
' Write Headers for DB Page

%>


<TABLE BORDER="0" CELLPADDING="6" CELLSPACING="0" WIDTH="100%" BGCOLOR="#C0C0C0" BORDERCOLOR="#C0C0C0" BORDERCOLORDARK="#C0C0C0" BORDERCOLORLIGHT="#C0C0C0" >
<TR>
<TD ALIGN="Left"><Font Face="courier" COLOR="#000000" SIZE="4"><B><% Response.Write(sTitle) %></B></FONT></TD>
</TR>
</TABLE>
<BR>

<%
End Sub




Sub WriteHeader
%>
<HTML>
<HEAD><TITLE>TRA Report Viewer</TITLE>
</HEAD>

<BODY BGCOLOR="#FFFFFE" Text="#0A0D0A" LINK="#375AE2" VLINK="#36566D" ALINK="#3E85BB">
<style TYPE="text/css">
<!--  A:link {text-decoration: none; color:#375AE2}  A:visited {text-decoration: none; color:#375AE2}  A:active {text-decoration: none}   A:hover {text-decoration: ; color:#3E85BB; }-->
</style>
<%
End Sub



Sub WriteFooter
%>
<hr>
</BODY>
</HTML>
<%
End Sub




Sub ChoosePagesSQL(sSQL,sStart, sSize)

  set rs=Server.CreateObject("ADODB.recordset")
  sqlstmt = sSQL
  rs.CursorType = 3
'  rs.PageSize = cint(sSize)
  rs.open sqlstmt, SConnectionToTRATable
'  IF isrecordsetempty = false THEN
'    rs.AbsolutePage = cINT(sStart)
'  END IF
End Sub



Function IsRecordSetEmpty

IF rs.bof = true and rs.eof = true THEN
    IsRecordSetEmpty = true
ELSE
    IsRecordSetEmpty = false
END IF
end Function




Sub WriteLink(sParms,sDisplay,sBreak)

%>
<A HREF="<% Response.Write(ThisPage & sParms) %>"><% Response.Write(sDisplay) %></A><% Response.Write(sBreak) %>
<%
End Sub




Sub DoCount(currentPage) 

h = 0

for i = 1 to rs.PageCount
 Response.Write(" <a href=" & chr(34) & ThisPage & "?div=" & DivSELECTed & "&ranknum=" & RankNum & "&event=" & EventSELECTed & "&currentpage=" &  i  & "&action=" & sAction & chr(34) & ">" & i & "</a>")
h = h +1
next
IF h = 0 THEN h = 1
Response.Write("<BR><Small>Page " & currentPage & " of  " & h & "</SMALL></center><BR><BR>")

END SUB

%>




