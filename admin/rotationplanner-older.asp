<!--#include virtual="epl/functions.asp" -->

<% 

Dim objConn, RS, HQConn, RSHQ
Set HQConn = Server.CreateObject("ADODB.Connection")
HQConn.Open Application("HQSQLConn")
Set RSHQ = Server.CreateObject("ADODB.RecordSet")
RSHQ.ActiveConnection = HQConn
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open Application("WaterSkiConn")
Set RS = Server.CreateObject("ADODB.RecordSet")
RS.ActiveConnection = objConn

Set objCDO = Server.CreateObject("CDO.Message")

Dim Usages(2,3,5), ErrStat, sSex

Function RemInvChr(strInput)
    dim workingstring
	On Error Resume Next
	For i = 1 to Len(strInput)
		If isNumeric(Mid(strInput, i, 1)) then
			workingstring = workingstring & Mid(strInput, i, 1)
		End If
		If (Mid(strInput, i, 1)) => "a" and (Mid(strInput, i, 1)) <=  "z" then
			workingstring = workingstring & Mid(strInput, i, 1)
		End If
		If (Mid(strInput, i, 1)) => "A" and (Mid(strInput, i, 1)) <=  "Z" then
			workingstring = workingstring & Mid(strInput, i, 1)
		End If
		If (Mid(strInput, i, 1)) = "@" Or (Mid(strInput, i, 1)) = "." Or (Mid(strInput, i, 1)) = " " Then
				workingstring = workingstring & Mid(strInput, i, 1)
		End If
	Next
	RemInvChr = workingstring
End Function


'	***** Bailout to Members Login if not auth or no Session("TeamId") value

IF not Session("auth") or Session("TeamID") = "" then response.redirect "https://www.usawaterski.org/members/login/index.asp"

'	*****	Store Selected Tournament to Session Variables, if re-entering at "NewPlan" of "SeePlan" points

IF Request("FormStatus") = "NewPlan" or Request("FormStatus") = "SeePlan" THEN 
	Session("TourID") = left(request("TourID"),6)
	Session("TourDate") = mid(request("TourID"),7,10)	
	Session("TourName") = mid(request("TourID"),17,len(request("TourID"))-16)
END IF

%>

<html>

<head>
<title>NCWSA Team Rotation Planning</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" >

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#42639F">
      <p>&nbsp;</p>
      <p align="center"><font face="Verdana" size="5" color="#FFFFFF">
      	<b>NCWSA Team Rotation Planning</b></font></p>
      <p align="center"><font face="Verdana" size="4" color="#FFFFFF">
			<% IF request("FormStatus") = "PickTour" THEN %>
	      	<%=Session("MemName")%>&nbsp;&nbsp; as Administrator for:&nbsp;&nbsp;&nbsp;&nbsp;<%=Session("TeamName")%>&nbsp;&nbsp; ( <%=Session("TeamID")%> ) </font></p>
			<% ELSE %>
	      	Team:&nbsp;&nbsp; <%=Session("TeamID")%> -- Rotation Plan for:&nbsp;&nbsp; <%=Session("TourName")%>&nbsp;&nbsp; ( <%=Session("TourID")%> ) </font></p>
			<% END IF %>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>

<table border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="180" valign="top" bgcolor="#42639F">

			<br>
	        &nbsp;<a href="https://www.usawaterski.org/members/"><font face="arial" size="2" COLOR="#FFFFFF">Member's Only Home</font></a><br>&nbsp;<br>
			<br>
            <font face="Verdana" size="1">&nbsp;<font COLOR="#FFFFFF">Powered by</font> <a href="http://www.epolk.com"><font COLOR="#FFFFFF">ePolk.com</font></a><br></font>
	
    </td>

<td width="720" >

<%

'	********** If Option is "PickTour" then present list of upcoming tournaments to pick from

IF Request("FormStatus") = "PickTour" THEN 

	%>

  <table>
    <tr> 
      <td width="20">&nbsp;</td>
      <td width="560">&nbsp;</td>
      <td width="20">&nbsp;</td>
    </tr>
 
    <tr> 
      <td>&nbsp;</td>
      <td valign="top">
        <font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        Following is a list of upcoming NCWSA Tournaments.&nbsp; Please select the 
        tournament for which you are going to create or revise your Rotation Plan
        from this drop-down list below.<br>&nbsp;
        </font></td>
      <td>&nbsp;</td>
    </tr>

    <form action="RotationPlanner.asp?FormStatus=NewPlan" method="post">
 
    <tr> 
      <td>&nbsp;</td>
      <td>
				<select name="TourID" size="11" onclick=submit()><%
		
				'	Create SQL to pull upcoming tournament list.

				sSQL = "Select ST.TournAppID, ST.TName,"
				sSQL = sSQL & " convert(char(10),ST.TDateE, 111) as TDateE,"
				sSQL = sSQL & " ST.TCity+', '+ST.TState as TLocation,"
				sSQL = sSQL & " ST.TStatus, ST.TSanType, CASE WHEN left(ST.TSanction,6) <> ST.TournAppID"
				sSQL = sSQL & " THEN ST.TournAppID+'?' ELSE ST.TSanction end as TSanction"
				sSQL = sSQL & " FROM Sanctions.dbo.TSchedul ST WHERE ST.TStatus in (0,1,2)"
				sSQL = sSQL & " AND ST.TDateE > GetDate() and ST.TSanType = 1" 
				sSQL = sSQL & " ORDER BY TDateE"
				
        RS.open sSQL

				DO WHILE NOT rs.eof

					response.write("<option value =""" & rs("TournAppID") & rs("TDateE") & RemInvChr(rs("TName")))
					response.write("""> " & rs("TDateE") & "&nbsp;&nbsp; " & rs("TournAppID"))
					response.write("&nbsp;&nbsp; " & rs("TName") & "&nbsp;&nbsp; ( " & rs("TLocation"))
					response.write(" ) </option>")

					rs.moveNEXT

				LOOP

				rs.close %>
				</select></form>
      </td>
      <td>&nbsp;</td>
    </tr>

	 <tr><td>&nbsp;</td></tr>
	   
  </table>
  
  </td></tr></table>

  <%


'	**********	IF Option is "SeePlan", then display existing plan.  
'	**********	This area will look a lot like "Validate", except pulls from Table.

ELSEIF Request("FormStatus") = "SeePlan" THEN 
	
	'	This table will offer an "Edit Plan" button, and that is 
	'	the ONLY outside entry to the later "EditPlan" section.





' **********	IF option is "NewPlan", then pull list of previous rotation plans for this team.
' **********	IF recordset includes selected tournament, then redirect to "EditPlan" for that event.
' **********	IF that list is empty, then redirect to "EditPlan" to build entirely new plan.
' **********	Otherwise offer this list of existing events from which to select for copy.

ELSEIF Request("FormStatus") = "NewPlan" THEN 
	
	'	Create SQL to pull existing rotation plans for this team.

	sSQL = "Select ST.TournAppID, ST.TName,"
	sSQL = sSQL & " convert(char(10),ST.TDateE, 111) as TDateE,"
	sSQL = sSQL & " ST.TCity+', '+ST.TState as TLocation,"
	sSQL = sSQL & " ST.TStatus, ST.TSanType, CASE WHEN left(ST.TSanction,6) <> ST.TournAppID"
	sSQL = sSQL & " THEN ST.TournAppID+'?' ELSE ST.TSanction end as TSanction"
	sSQL = sSQL & " FROM Sanctions.dbo.TSchedul ST WHERE ST.TournAppID in"
	sSQL = sSQL & " (Select distinct TournAppID from Cobra00025.USAWSRank.TeamRotations"
	sSQL = sSQL & " WHERE Team = '" & Session("TeamID") & "') ORDER BY TDateE desc"
				
  RS.open sSQL

	IF rs.eof THEN response.redirect "RotationPlanner.asp?FormStatus=EditPlan&TourID=" & Session("TourID") 

	' Next we loop over recordset to see if we already have a plan for the selected tournament.

	EditPlan = "N"
	DO WHILE NOT rs.eof
		IF rs("TournAppID") = Session("TourID") THEN EditPlan = "Y"
		rs.moveNEXT
	LOOP
	IF EditPlan = "Y" THEN response.redirect "RotationPlanner.asp?FormStatus=EditPlan&TourID=" & Session("TourID") 
	
	RS.MoveFirst

	' No existing plan for this tournament, but do have others, so offer list ...
	
	%>

  <table>
    <tr> 
      <td width="20">&nbsp;</td>
      <td width="560">&nbsp;</td>
      <td width="20">&nbsp;</td>
    </tr>
 
    <tr> 
      <td>&nbsp;</td>
      <td valign="top">
        <font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        Following is a list of existing rotation plans for your team, for 
        other recent NCWSA Tournaments.&nbsp; You may elect to copy one of 
        those other rotation plans, to use as a starting point for your new 
        plan for <%=Session("TourID")%>&nbsp; <%=Session("TourName")%>.&nbsp; 
        Select the "New Plan" choice to begin an entirely new rotation plan 
        for this tournament; otherwise select the specific tournament whose 
        rotation plan you wish to copy, from the list below.<br>&nbsp;
        </font></td>
      <td>&nbsp;</td>
    </tr>

    <form action="RotationPlanner.asp?FormStatus=EditPlan" method="post">
 
    <tr> 
      <td>&nbsp;</td>
      <td>
				<select name="TourID" size="11" onclick=submit()><%
		
				response.write("<option value =""" & Session("TourID") & """> ")
				response.write(Session("TourDate") & "&nbsp;&nbsp; " & Session("TourID") & "&nbsp;&nbsp; Create an entirely New Plan </option>")

				DO WHILE NOT rs.eof

					response.write("<option value =""" & rs("TournAppID"))
					response.write("""> " & rs("TDateE") & "&nbsp;&nbsp; " & rs("TournAppID"))
					response.write("&nbsp;&nbsp; " & rs("TName") & "&nbsp;&nbsp; ( " & rs("TLocation"))
					response.write(" ) </option>")

					rs.moveNEXT

				LOOP

				rs.close %>
				</select></form>
      </td>
      <td>&nbsp;</td>
    </tr>

	 <tr><td>&nbsp;</td></tr>
	   
  </table>
  
  </td></tr></table>

  <%

'	********** Enter here when user has chosen the plan to work from.
'	********** This is the primary Rotation Plan Edit / Update screen.

ELSEIF Request("FormStatus") = "EditPlan" THEN 

	' ********** IF EditPlan is a different ID, then copy that other plan first
	'	********** -- but only if we still have zero rows for this TourID --
	'	********** -- this avoids problem if user uses Back button after copy.
	'	********** Copied Plan will have WaiverStat = 'C' on all detail rows.

	IF Request("TourID") <> Session("TourID") THEN

		sSQL = "Select count(*) as Kount FROM Cobra00025.USAWSRank.TeamRotations"
		sSQL = sSQL & " Where Team = '" & Session("TeamID") & "' and TournAppID = '"	
		sSQL = sSQL & Session("TourID") & "'" 
	  RS.open sSQL
	  Kount = rs("Kount")
	  RS.close

	  IF Kount = 0 THEN
			sSQL = "INSERT INTO Cobra00025.USAWSRank.TeamRotations (Team, TournAppID, MemberID,"
			sSQL = sSQL & " Sex, SlalomEnt, TrickEnt, JumpEnt, DateUpdated, WaiverStat) SELECT"
			sSQL = sSQL & " Team, '" & Session("TourID") & "', MemberID, Sex, SlalomEnt, TrickEnt,"
			sSQL = sSQL & " JumpEnt, GetDate(), 'C' FROM Cobra00025.USAWSRank.TeamRotations"
			sSQL = sSQL & " Where Team = '" & Session("TeamID") & "' and TournAppID = '"	
			sSQL = sSQL & request("TourID") & "'"
			objConn.Execute (sSQL)
		END IF

	END IF
		
				
  ' Set up query to pull Team Rotation Plan Table for current TeamID and TourID

	sSQL = "SELECT Mem.PersonIDWithCheckDigit AS MemberID," 
	sSQL = sSQL & " Mem.LastName, Mem.FirstName, Mem.Email, Left(Mem.Sex,1) as Sex,"

	sSQL = sSQL & " CASE when DateAdd(mm,8,Mem.EffectiveTo) < TS.TDateE then 'X'"

'	sSQL = sSQL & " CASE when Mem.EffectiveTo < TS.TDateE then 'X'"

	sSQL = sSQL & " when Typ.CanSkiInTournaments = 0 then 'U'" 
	sSQL = sSQL & " when Mem.WaiverStatusID = 0 then 'W' else 'G' end as MemStat,"
	sSQL = sSQL & " CASE when patindex('%@%',Mem.Email) <= 0 then 'N' else 'Y' end as EMStat,"
	sSQL = sSQL & " Coalesce(RP.SlalomEnt,'  ') as SlalomEnt, Coalesce(RP.TrickEnt,'  ') as TrickEnt,"
	sSQL = sSQL & " Coalesce(RP.JumpEnt,'  ') as JumpEnt, Coalesce(RP.WaiverStat,' ') as WaiverStat,"
	sSQL = sSQL & " Coalesce(SX.SlmSco,'') as SlmRank, Coalesce(TX.TrkSco,'') as TrkRank,"
	sSQL = sSQL & " Coalesce(JX.JmpSco,'') as JmpRank"
	sSQL = sSQL & " FROM Sanctions.dbo.TSchedul TS, Cobra00025.USAWSRank.TeamRoster as TR"
	sSQL = sSQL & " JOIN USAWaterski.dbo.members as Mem"
	sSQL = sSQL & " on Mem.PersonIDWithCheckDigit = TR.MemberID"
	sSQL = sSQL & " JOIN USAWaterski.dbo.MembershipTypes as Typ"
	sSQL = sSQL & " ON Typ.MembershipTypeID = Mem.MemberShipTypeCode"

	sSQL = sSQL & " LEFT JOIN Cobra00025.USAWSRank.TeamRotations as RP"	
	sSQL = sSQL & " on RP.MemberID = TR.MemberID and RP.Team = TR.Team"
	sSQL = sSQL & " and RP.TournAppID = '" & Session("TourID") & "'"
	
	sSQL = sSQL & " Left Join	(Select MemberID, AWSA_Rat as SlmRat,"
	sSQL = sSQL & " Left(Cast(Cast(RankScore as Decimal(7,2)) as Varchar(8)),6) as SlmSco"
	sSQL = sSQL & " From Cobra00025.USAWSRank.Rankings Where SkiYearID = 1"
	sSQL = sSQL & " and Left(Div,1) = 'C' and Event = 'S' and RankScore is not null)"
	sSQL = sSQL & " as SX on SX.MemberID = TR.MemberID"

	sSQL = sSQL & " Left Join	(Select MemberID, AWSA_Rat as TrkRat,"
	sSQL = sSQL & " Left(Cast(Cast(RankScore as Decimal(7,1)) as Varchar(8)),6) as TrkSco"
	sSQL = sSQL & " From Cobra00025.USAWSRank.Rankings Where SkiYearID = 1"
	sSQL = sSQL & " and Left(Div,1) = 'C' and Event = 'T' and RankScore is not null)"
	sSQL = sSQL & " as TX on TX.MemberID = TR.MemberID"

	sSQL = sSQL & " Left Join	(Select MemberID, AWSA_Rat as JmpRat,"
	sSQL = sSQL & " Left(Cast(Cast(RankScore as Decimal(6,2)) as Varchar(8)),6) as JmpSco"
	sSQL = sSQL & " From Cobra00025.USAWSRank.Rankings Where SkiYearID = 1"
	sSQL = sSQL & " and Left(Div,1) = 'C' and Event = 'J' and RankScore is not null)"
	sSQL = sSQL & " as JX on JX.MemberID = TR.MemberID"

	sSQL = sSQL & " WHERE TR.Team = '" & Session("TeamID") & "' and TR.DateInactive is NULL"
	sSQL = sSQL & " and TS.TournAppID = '" & Session("TourID") & "'"
	sSQL = sSQL & " order by Mem.Sex, Mem.LastName, Mem.FirstName"
		
 	RS.open sSQL

	%>

	<TABLE BORDER="4" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#DDDDFF" width=95% >
		<tr><br>
			<td BGCOLOR="#42639F">
	        <center><font face="Verdana, Arial, Helvetica, sans-serif" size="4" COLOR="#FFFFFF"><b>Design your Rotation Plan with the drop-downs below</b></font>
			<br></td>
		</tr>  

		<tr>
			<td>
  			  <br>

				<form action="RotationPlanner.asp?FormStatus=Validate" method="post">

				<TABLE BORDER="1" ALIGN="CENTER" CELLPADDING="1" CELLSPACING="0" BGCOLOR="#F5F5F5" width=95% >

	<%

	sSex = "?": index = 0

	DO WHILE NOT RS.EOF 
								
		IF rs("Sex") <> sSex THEN
			IF rs("Sex") = "F" THEN %>
				<TR><TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><b> Women's Team </b></FONT></TD>
			<% ELSE %>
				<TR><TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><b> Men's Team </b></FONT></TD>
			<% END IF %>

	    <TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><b> Slalom </b></FONT></TD>
 	    <TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><b> Trick </b></FONT></TD>
	    <TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><b> Jump </b></FONT></TD>
	    <TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>Wvr</b></FONT></TD>
			
		<%
		
		END IF
		sSex = rs("Sex")

		sMemStat = rs("Memstat")

		'	**********	Now if Membership Status is not "G" (good), 
		'	**********	then pull latest status from HQ server and re-classify

		IF sMemStat <> "G" THEN

			sSQL = "Select CASE when convert(char(10),DateAdd(mm,8,MH.EffectiveTo),111) < '" 

'			sSQL = "Select CASE when convert(char(10),MH.EffectiveTo,111) < '" 

			sSQL = sSQL & Session("TourDate") & "' then 'X' when MT.CanCompete = 0 then 'U'"
			sSQL = sSQL & " when MH.WaiverStatusID = 0 then 'W' else 'G' end as MemStat"
			sSQL = sSQL & " FROM Waterski.dbo.[Membership History] as MH JOIN"
			sSQL = sSQL & " (Select [Person ID] as PersonID, Max(EffectiveTo) as MaxEffTo"
			sSQL = sSQL & " From Waterski.dbo.[Membership History] where [Person id] = "
			sSQL = sSQL & right(rs("MemberID"),8) & " group by [Person ID]) as ME"
			sSQL = sSQL & " ON MH.[Person ID] = ME.PersonID AND MH.EffectiveTo = ME.MaxEffTo"			
			sSQL = sSQL & " JOIN	waterski.dbo.tblMembershipTypeCodes as MT"
			sSQL = sSQL & " on MT.[Membership Type Code] = MH.[Membership Type Code]"

			RSHQ.open sSQL
			IF NOT RSHQ.eof THEN sMemStat = RSHQ("MemStat")
			RSHQ.close
		 	
		END IF

		'	**********	Now classify email and membership status 

		SELECT CASE sMemStat&rs("EMStat")
			CASE "UN": sNote = "Needs Upgrade &amp; eMail Address"
			CASE "WN": sNote = "Needs Annual Waiver &amp; eMail Address"
			CASE "XN": sNote = "Needs Renewal &amp; eMail Address"
			CASE "GN": sNote = "Needs eMail Address"
			CASE "UY": sNote = "Needs Membership Upgrade"
			CASE "WY": sNote = "Needs Annual USA Water Ski Waiver"
			CASE "XY": sNote = "Membership will be Expired; must Renew"
			CASE ELSE: sNote = ""
		END SELECT
		
		IF sNote <> "" THEN 

			%>

			<tr>
  		  <TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("LastName")%>, <%=rs("FirstName")%></FONT></TD>
				<TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF" colspan=4><FONT COlOR="red" size="2" face="Verdana, Arial, Helvetica, sans-serif"><b> <%=sNote%> </b></FONT></TD>
				</tr>

			<%

		ELSE 

			sMemberName = rs("FirstName") & " " & rs("LastName")			
			index = index + 1 %>

			<tr>
			<input type="hidden" name="MemberID<%=index%>" value="<%=rs("MemberID")%>">
			<input type="hidden" name="MemberName<%=index%>" value="<%=sMemberName%>">
			<input type="hidden" name="MemberEmail<%=index%>" value="<%=RemInvChr(rs("Email"))%>">
			<input type="hidden" name="MemberSex<%=index%>" value="<%=rs("Sex")%>">
			<input type="hidden" name="WaiverStat<%=index%>" value="<%=rs("WaiverStat")%>">

			<TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=rs("LastName")%>, <%=rs("FirstName")%></FONT></TD>

			<%
			
			OfferRotation "SlalomEnt", rs("SlalomEnt"), index, rs("SlmRank"), rs("FirstName"), rs("MemberID")
			OfferRotation "TrickEnt", rs("TrickEnt"), index, rs("TrkRank"), rs("FirstName"), rs("MemberID")
			OfferRotation "JumpEnt", rs("JumpEnt"), index, rs("JmpRank"), rs("FirstName"), rs("MemberID")

			IF rs("WaiverStat") = "E" THEN %>
				<td align="Center"><a title="Event Waiver email request sent,&#13; but not yet acted upon"><img src="/admin/Envelope6.gif"></a></td>
			<% ELSEIF rs("WaiverStat") = "X" THEN %>
				<td align="Center"><a title="Event Waiver Accepted and Filed"><img src="/admin/Smile19.gif"></a></td>
			<% ELSE %>
				<td>&nbsp;</td>
			<% END IF %>
				
			</tr>
			
			<% 

		END IF

		RS.MoveNext

	LOOP

	rs.Close

	'	**********	Done -- now drop the Number of populated rows into a hidden form
	'	**********	variable, then finally present option buttons at bottom of form.

 	  %></table>

		<input type="hidden" name="FormRows" value="<%=index%>">

		<TABLE BORDER="0" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#DDDDFF" width=95% >
		  <tr><td>&nbsp;</td></tr></table>

				<TABLE BORDER="0" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#DDDDFF" width=95% >
				  <tr>
	
				    <TD width=33% align=center>
					<input type="submit" style="width:9em" value="Validate Plan"
						title="Analyze and Validate details of the above&#13;Rotation Plan for the specified tournament"></form>
			    	</TD>

					</form>

   			    <td width=33% align=center>     				
					<form action="rostermanager.asp" method="link">
				   <input type="submit" style="width:9em" value="Back to Roster"
				   	title="Return to Team Roster"></form>
			   	</td>

   			    <td width=33% align=center>     				
					<form action="FAQ_NCWRosters.htm" method="post" target="_blank">
					<input type="submit" style="width:7em" value="Tips/Help"
						title="Tips and Explanations and Insights and &#13;Answers to Frequently Asked Questions"></form>
			  	   </td>
	
				</TR>
		     </table> 

 	  </td></tr>
 	</table><%    


'	**********	Return here with User-Edited Plan details -- update plan table and validate rotations

ELSEIF Request("FormStatus") = "Validate" THEN 

	%>

	<TABLE BORDER="4" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#DDDDFF" width=95% >
		<tr><br>
			<td BGCOLOR="#42639F">
	        <center><font face="Verdana, Arial, Helvetica, sans-serif" size="4" COLOR="#FFFFFF"><b>Evaluation of your Rotation Plan ...</b></font>
			<br></td>
		</tr>  

		<tr>
			<td>
  			  <br>

				<TABLE BORDER="1" ALIGN="CENTER" CELLPADDING="1" CELLSPACING="0" BGCOLOR="#F5F5F5" width=95% >

	<%

	sSex = "?": ErrStat = "N"
	
	For Index = 1 to Request("FormRows")
	
		InMemID = Request("MemberID" & Index)
		InName = Request("MemberName" & Index)
		InEmail = Request("MemberEmail" & Index)
		InSex = Request("MemberSex" & Index)
		InWaiver = Request("WaiverStat" & Index)
		InSlm = Request("SlalomEnt" & Index)
		InTrk = Request("TrickEnt" & Index)
		InJmp = Request("JumpEnt" & Index)
	
		IF sSex <> InSex THEN

			IF InSex = "F" THEN %>
				<TR><TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><b> Women's Team </b></FONT></TD>
			<% ELSE 
				IF sSex = "F" THEN RecapErrors "Women" %>
				<TR><TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><b> Men's Team </b></FONT></TD>
			<% END IF %>

			<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><b> Slalom </b></FONT></TD>
			<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><b> Trick </b></FONT></TD>
			<TD bgcolor="#42639F" ALIGN="Center" vAlign="Center"><FONT COlOR="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><b> Jump </b></FONT></TD>
			</TR>

			<%

			For I=1 to 2
				For J=1 to 3
					For K=1 to 5
						Usages(I,J,K) = "N"
					NEXT
				NEXT
			NEXT
		
		END IF
		sSex = InSex

		'	**********	IF this form row is all empty, and WaiverStat IS set, then reset this row in table.

		IF InSlm = "  " and InTrk = "  " and InJmp = "  " THEN

			IF InWaiver <> " " THEN
			
				sSQL = "Update Cobra00025.USAWSRank.TeamRotations Set"
				sSQL = sSQL & " SlalomEnt=' ', TrickEnt=' ', JumpEnt=' ',"
				sSQL = sSQL & " DateUpdated=GetDate() WHERE MemberID='"
				sSQL = sSQL & InMemID & "' and Team='" & Session("TeamID") 
				sSQL = sSQL & "' and TournAppID='" & Session("TourID") & "'"
				objConn.Execute (sSQL)

			END IF
			
		'	**********	Otherwise we now have an active entry.  So check whether this entry already
		'	**********	exists in the TeamRotation table -- if so, update that row

		ELSE
		
			sSQL = "Select count(*) as Kount FROM Cobra00025.USAWSRank.TeamRotations"
			sSQL = sSQL & " Where Team = '" & Session("TeamID") & "' and TournAppID = '"	
			sSQL = sSQL & Session("TourID") & "' and MemberID = '" & InMemID & "'"
			RS.open sSQL
			Kount = rs("Kount")
			RS.close

			IF Kount > 0 THEN
		
				sSQL = "Update Cobra00025.USAWSRank.TeamRotations Set"
				sSQL = sSQL & " SlalomEnt='" & InSlm & "', TrickEnt='" & InTrk 
				sSQL = sSQL& "', JumpEnt='" & InJmp	& "', DateUpdated=GetDate()"			
				sSQL = sSQL & " WHERE MemberID='" & InMemID
				sSQL = sSQL & "' and Team='" & Session("TeamID") 
				sSQL = sSQL & "' and TournAppID='" & Session("TourID") & "'"
				objConn.Execute (sSQL)

			'	**********	This is a new Entry -- so insert into TeamRotation table, WaiverStat
			'	**********	initially set to 'C', which means eMail request has not been sent.

			ELSE
			
				sSQL = "Insert into Cobra00025.USAWSRank.TeamRotations (Team, TournAppID,"
				sSQL = sSQL & " MemberID, Sex, SlalomEnt, TrickEnt, JumpEnt, DateUpdated,"
				sSQL = sSQL & " WaiverStat) Values ('" & Session("TeamID") & "', '"
				sSQL = sSQL & Session("TourID") & "', '" & InMemID & "', '" & InSex
				sSQL = sSQL & "', '" & InSlm & "', '" & InTrk & "', '" & InJmp 
				sSQL = sSQL & "', GetDate(), 'C')"
				objConn.Execute (sSQL)
				
				InWaiver = "C"

			END IF

		%>
			<TR>
			<TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=InName%></FONT></TD>

		<%
			ShowRotation InSlm, 1
			ShowRotation InTrk, 2
			ShowRotation InJmp, 3
		%>

			</TR>
		<% 
		
		END IF

	NEXT
	
	'	*****	Now Create Male Team Error / Warning Strings and tack onto bottom, unless no Men 

	IF sSex = "M" THEN RecapErrors "Men": ELSE RecapErrors "Women"

 	  %></table>

		<TABLE BORDER="0" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#DDDDFF" width=95% >
		  <tr><td>&nbsp;</td></tr></table>

				<TABLE BORDER="0" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#DDDDFF" width=95% >
				  <tr>
	
				    <TD width=24% align=center>
					<form action="RotationPlanner.asp?FormStatus=PlanOK" method="post">
					
					<% 
					
					'	*****	Build pass-along form values, for next "PlanOK" stage, with active entries only, for Emails etc.
					
					NewIndex = 0
					For Index = 1 to Request("FormRows")
						IF Request("SlalomEnt"&index) <> "  " OR Request("TrickEnt"&index) <> "  " OR (Request("JumpEnt"&index)) <> "  " THEN 
							NewIndex = NewIndex + 1
							InWaiver = Request("WaiverStat" & Index)
							IF InWaiver = " " THEN InWaiver = "C"
							%>
							<input type="hidden" name="MemberID<%=NewIndex%>" value="<%=Request("MemberID" & Index)%>">
							<input type="hidden" name="MemberName<%=NewIndex%>" value="<%=Request("MemberName" & Index)%>">
							<input type="hidden" name="MemberEmail<%=NewIndex%>" value="<%=Request("MemberEmail" & Index)%>">
							<input type="hidden" name="WaiverStat<%=NewIndex%>" value="<%=InWaiver%>">
							<%
						END IF
					NEXT 

					%>

					<input type="hidden" name="FormRows" value="<%=NewIndex%>">
							
					<input <% IF ErrStat = "Y" then response.write("disabled") %> type="submit" style="width:9em" value="Submit Plan"
						title="Submit this validated Rotation Plan for the specified tournament"></form>
			    	</TD>

				    <TD width=24% align=center>
					<form action="RotationPlanner.asp?FormStatus=EditPlan&TourID=<%=Session("TourID")%>" method="post">
					<input type="submit" style="width:9em" value="Revise Plan"
						title="Revise this Rotation Plan for the specified tournament"></form>
			    	</TD>

   			    <td width=24% align=center>     				
					<form action="rostermanager.asp" method="link">
				   <input type="submit" style="width:9em" value="Back to Roster"
				   	title="Return to Team Roster"></form>
			   	</td>

   			    <td width=24% align=center>     				
					<form action="FAQ_NCWRosters.htm" method="post" target="_blank">
					<input type="submit" style="width:7em" value="Tips/Help"
						title="Tips and Explanations and Insights and &#13;Answers to Frequently Asked Questions"></form>
			  	   </td>
	
				</TR>
		     </table> 

 	  </td></tr>
 	</table><%    


'	**********	Return here with Validated Plan details -- 
'	**********	Email entrants and update table, where WaiverStat = "C"

ELSEIF Request("FormStatus") = "PlanOK" THEN 
	
	For Index = 1 to Request("FormRows")

		Response.write ("<br>&nbsp;&nbsp;&nbsp;&nbsp; " & Request("WaiverStat"&Index) & " / " & Request("MemberID"&Index) & " / " & Request("MemberName"&Index) & " / " & Request("MemberEmail"&Index))
	
		IF Request("WaiverStat"&Index) = "C" THEN
	
			ebody = "<html><head><title>Waiver and Release</title>"
			ebody = ebody & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
			ebody = ebody & "</head><body bgcolor=""#FFFFFF"" text=""#000000""><div align=""left"">"
			ebody = ebody & "<font face=""Verdana, Arial, Helvetica, sans-serif"" size=""2""><p>Subj: "
			ebody = ebody & "Event Waiver for " & Request("MemberName"&Index) & " for " & Session("TourID") 
			ebody = ebody & " " & Session("TourName") & " " & Session("TourDate") & ".</p><p>Dear "
			ebody = ebody & left(Request("MemberName"&Index),instr(Request("MemberName"&Index)," ")-1)
			ebody = ebody & ",</p><p>You have been entered by your Team Captain as a member of the ("
			ebody = ebody & Session("TeamID") & ") " & Session("TeamName") & " Collegiate Waterski team, " 
			ebody = ebody & "to be a participant in the following upcoming NCWSA tournament:</p>"
			ebody = ebody & "<p><font color=""blue"" size=""3""><b>" & Session("TourID") & "&nbsp;&nbsp;&nbsp; "
			ebody = ebody & Session("TourName") & "&nbsp;&nbsp;&nbsp; " & Session("TourDate") & "</b></font></p>"
			ebody = ebody & "<p>All USA Water Ski sanctioned competitions require that each participant "
			ebody = ebody & "execute an Event Waiver and Release form, which can be done most easily "
			ebody = ebody & "through the USA Water Ski Online system.&nbsp; Clicking on the "
			ebody = ebody & "<b>Online Waiver</b> link you see below will take you directly to a "
			ebody = ebody & "customized online Waiver and Release form, prepared for you and for this "
			ebody = ebody & "particular event.&nbsp; Please do so at your earliest convenience.</p>"
			ebody = ebody & "<p>Thank you on behalf of your team captain " & Session("MemName") & ".</p>"
			ebody = ebody & "<p>USA Water Ski Competition Dept.&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; "			

			ebody = ebody & "<a href=""https://www.usawaterski.org/admin/ncwsawaiver.asp?FormStatus=Tour&TourID="
			ebody = ebody & Session("TourID") & "&PID=" & Request("MemberID"&Index) mod 100000000
			ebody = ebody & """	title=""Execute the Waiver and Release form Online""><b>Online Waiver</b></a>"

			ebody = ebody & "</p><p>&nbsp;</p><p>P. S.&nbsp; If your email client does not show you "
			ebody = ebody & "an <b>Online Waiver</b> button above, then you can reach this form by "
			ebody = ebody & "signing onto the <b>Member's Only</b> area of the USA Water Ski website "
			ebody = ebody & "(Member Login link located at the upper right corner of the USA Water "
			ebody = ebody & "Ski home page).&nbsp; Once logged in there, then you will find the "
			ebody = ebody & "<b>Participant Waiver</b> link under the <b>Collegiate Registration</b> "		
			ebody = ebody & "menu heading, in the left navigation panel.</p></font></div></body></html>"
									
			objCDO.Subject = "Event Waiver for " & Request("MemberName"&Index) & " for " & Session("TourID") & " " & Session("TourName") & " " & Session("TourDate")
			objCDO.To = """" & Request("MemberName"&Index) & """ <" & Request("MemberEmail"&Index) & ">"
			objCDO.From = "competition@usawaterski.org"
			objCDO.CC = """Dave Clark"" <awsatechdude@comcast.net>"
			objCDO.HTMLBody = ebody	

			objCDO.Send

			sSQL = "Update Cobra00025.USAWSRank.TeamRotations set WaiverStat = 'E'"
			sSQL = sSQL & " Where MemberID = '" & Request("MemberID"&Index) 
			sSQL = sSQL & "' and TournAppID = '" & Session("TourID") & "'"
			objConn.Execute (sSQL)

		END IF 

	NEXT

	%>
		
  <table>
    <tr> 
      <td width="20">&nbsp;</td>
      <td width="560">&nbsp;</td>
      <td width="20">&nbsp;</td>
    </tr>

    <tr> 
      <td>&nbsp;</td>
      <td valign="top">
    
        <font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        Your validated Rotation Plan for the <%=Session("TeamName")%> Team (<%=Session("TeamID")%>)
        has been submitted to the Tournament Registrar for the above-noted NCWSA
        tournament.&nbsp; Each entered member of your team has been sent an eMail 
        note, requesting them to execute an Event
        Waiver and Release form for their participation in this event, through
        the online system, as soon as possible.<br>&nbsp;<br>
        
        Having your team members do those Waivers through the Online
        system now, will greatly simplify the few remaining steps you and
        your team need to carry out to participate in this event.&nbsp; So
        please remind your team members to get those waivers done soon.
        <br>&nbsp;<br>
        
        Thank you for using this system to prepare and submit your team
        entry and rotation plan.<br>&nbsp;<br>
        </font></td>
      <td>&nbsp;</td>
   </tr> 

	</table>

	<table align="center" width=75%>
	<tr>	
   	<td align="center">
   		<form action="rostermanager.asp" method="link">
		   <input type="submit" style="width:9em" value="Back to Roster"
		   	title="Return to your Team Roster"></form>
		</td>

		<td align="center">
			<form action="RotationPlanner.asp?FormStatus=PickTour" method="post">
			<input type="submit" style="width:11em" value="Do another Plan"
				title="Create or Revise a Rotation Plan for another tournament"></form>
		</td>					
	
   </tr> 
	</table>
  
  </td></tr></table>

	<%

END IF

set rs = nothing


'	*******************************************
SUB OfferRotation ( FieldName, FieldValue, FormIndex, FormRank, tName, tMemberID )
'	*******************************************

	%>
		
	<td align="Center"><select name="<%=FieldName%><%=FormIndex%>">
	  <option value ="  " <%IF FieldValue="  " THEN response.write(" selected")%>> &nbsp; </option>
	  <option value ="A1" <%IF FieldValue="A1" THEN response.write(" selected")%>> A Tm R 1 </option>
	  <option value ="A2" <%IF FieldValue="A2" THEN response.write(" selected")%>> A Tm R 2 </option>
	  <option value ="A3" <%IF FieldValue="A3" THEN response.write(" selected")%>> A Tm R 3 </option>
	  <option value ="A4" <%IF FieldValue="A4" THEN response.write(" selected")%>> A Tm R 4 </option>
	  <option value ="A5" <%IF FieldValue="A5" THEN response.write(" selected")%>> A Tm R 5 </option>
	  <option value ="B1" <%IF FieldValue="B1" THEN response.write(" selected")%>> B Tm R 1 </option>
	  <option value ="B2" <%IF FieldValue="B2" THEN response.write(" selected")%>> B Tm R 2 </option>
	  <option value ="B3" <%IF FieldValue="B3" THEN response.write(" selected")%>> B Tm R 3 </option>
	  <option value ="B4" <%IF FieldValue="B4" THEN response.write(" selected")%>> B Tm R 4 </option>
	  <option value ="B5" <%IF FieldValue="B5" THEN response.write(" selected")%>> B Tm R 5 </option>
	  <option value ="DD" <%IF FieldValue="DD" THEN response.write(" selected")%>> Age Div </option>
	</select>&nbsp;&nbsp;&nbsp;
	<a href="/rankings/view-scoresHQ.asp?NSL=&sMemberID=<%=tMemberID%>&EventSelected=<%=left(FieldName,1)%>&pvar=ByMember" target="_blank"
		 title="<%IF FormRank <> "" THEN Response.write("NCWSA Ranking: " & FormRank & "&#13;")%>Click here to Display ALL&#13;of <%=tName%>'s <%=left(FieldName,len(FieldName)-3)%> scores"><img src="/admin/Magnifier6.gif" STYLE="border-style:none"></a>

	</td>

	<%
	
END SUB


'	*******************************************
SUB ShowRotation ( FieldValue, EventIndex )
'	*******************************************

IF Left(FieldValue,1) = "A" OR Left(FieldValue,1) = "B" THEN
	Rotation = Left(FieldValue,1) & " Tm R " & Right(FieldValue,1)
	Position = Right(FieldValue,1)
	IF Left(FieldValue,1) = "A" THEN TeamType = 1: ELSE TeamType = 2
	IF Usages(TeamType,EventIndex,Position) = "Y" THEN 
		Rotation = "<font color=red><b>" & Rotation & "</b></font>"
		Usages(TeamType,EventIndex,Position) = "D"	
	ELSEIF Left(FieldValue,1) = "B" THEN
		Rotation = "<font color=blue>" & Rotation & "</font>"
	END IF
	IF Usages(TeamType,EventIndex,Position) = "N" THEN 
		Usages(TeamType,EventIndex,Position) = "Y"	
	END IF
ELSEIF FieldValue = "DD" THEN 
	Rotation = "<font color=magenta>Age Div<font>"
ELSE
	Rotation = "&nbsp;"
END IF
	
%>
	<TD ALIGN="Center" vAlign="Center" BGCOLOR="#FFFFFF"><FONT COlOR="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=Rotation%></FONT></TD>
<%

END SUB



'	*******************************************
SUB RecapErrors ( TeamName )
'	*******************************************

%>
	<TR>
	<TD ALIGN="Center" vAlign="Center" BGCOLOR="#DDFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=TeamName%>'s Team<br>Errors & Warnings:</FONT></TD>
<%
	
For J=1 to 3
	ErrStr = ""
	For I=1 to 2
		DupErr = "": MsgErr = ""
		For K=1 to 5
			IF Usages(I,J,K) = "D" Then
				IF DupErr <> "" THEN DupErr = DupErr & "/"
				DupErr = DupErr & K
				ErrStat = "Y"
			END IF
			IF Usages(I,J,K) = "N" then 
				IF MsgErr <> "" THEN MsgErr = MsgErr & "/"
				MsgErr = MsgErr & K
			END IF
		NEXT
		ErrStr = ErrStr & Mid("AB",I,1) & " Tm: "
		IF DupErr = "" and MsgErr = "" THEN 
			ErrStr = ErrStr & "OK"
		ELSE
			IF DupErr <> "" THEN ErrStr = ErrStr & "<font color=""red""><b>Dup: " & DupErr & "</b></font> "
			IF MsgErr <> "" THEN ErrStr = ErrStr & "Skip: " & MsgErr & " "
		END IF
		IF I=1 then ErrStr = ErrStr & "<br>"
	NEXT
	
	%>
		<TD ALIGN="Center" vAlign="Center" BGCOLOR="#DDFFFF"><FONT COlOR="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=ErrStr%></FONT></TD>
	<%

NEXT

%>
	</TR>
<%
	
END SUB



%>
 