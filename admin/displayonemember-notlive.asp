<!--#include virtual="epl/functions.asp" -->

<% If not Session("aauth") then response.redirect "Login.asp" %>

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

<%  	
Dim objConn
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open Application("WaterSkiConn")
%>
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

    <%

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
		If (Mid(strInput, i, 1)) = "@" Or (Mid(strInput, i, 1)) = "." Then
				workingstring = workingstring & Mid(strInput, i, 1)
		End If
	Next
	RemInvChr = workingstring
	
End Function

Set SQLConnect = CreateObject("ADODB.Connection")
SQLConnect.Open Application("HQSQLConn")

Function GetMemberShipPrice(EffectiveDate, MembershipTypeCode, RUCode)
	strSql = "SELECT * FROM [Membership Types with pricing] " 
	strSql = strSql & "WHERE EffectiveFrom <= CONVERT(DATETIME, '" & EffectiveDate & " 00:00:00', 102) AND "
	strSql = strSql & "EffectiveTo >= CONVERT(DATETIME, '" & EffectiveDate & " 00:00:00', 102) AND "
	strSql = strSql & "[Membership Type Code] =  " & MembershipTypeCode
	Set MemPrices = SQLConnect.Execute(strSql)
	IF MemPrices.EOF then
		GetMemberShipPrice = 0 
	ELSE
		IF RUCode = "R" THEN
			GetMemberShipPrice = MemPrices("MemberShipTypeRates")
		ELSEIF RUCode = "U" THEN
			GetMemberShipPrice = MemPrices("CostToUpgrade")
		ELSE
			GetMemberShipPrice = MemPrices("MemberShipTypeRates") + Memprices("CostToUpgrade")
		END IF			
	END IF
	MemPrices.Close
	Set MemPrices = Nothing
End Function

Function CalculateDivision(SkiAge, Gender)
Dim AgeDivision
if len(SkiAge) = 0 then
	AgeDivision = "-"
elseif SkiAge >= 0 AND SkiAge < 10 THEN '1' 
	AgeDivision = "1"
elseif  SkiAge >= 10 AND SkiAge < 14 THEN '2' 
	AgeDivision = "2"
elseif  SkiAge >= 14 AND SkiAge < 18 THEN '3' 
	AgeDivision = "3"
elseif  SkiAge >= 18 AND SkiAge < 25 THEN '1' 
	AgeDivision = "1"
elseif  SkiAge >= 25 AND SkiAge < 35 THEN '2' 
	AgeDivision = "2"
elseif  SkiAge >= 35 AND SkiAge < 45 THEN '3' 
	AgeDivision = "3"
elseif  SkiAge >= 45 AND SkiAge < 53 THEN '4' 
	AgeDivision = "4"
elseif  SkiAge >= 53 AND SkiAge < 60 THEN '5' 
	AgeDivision = "5"
elseif  SkiAge >= 60 AND SkiAge < 65 THEN '6' 
	AgeDivision = "6"
elseif  SkiAge >= 65 AND SkiAge < 70 THEN '7' 
	AgeDivision = "7"
elseif  SkiAge >= 70 AND SkiAge < 75 THEN '8' 
	AgeDivision = "8"
elseif  SkiAge >= 75 AND SkiAge < 80 THEN '9' 
	AgeDivision = "9"
elseif  SkiAge >= 80 AND SkiAge < 85 THEN 'A' 
	AgeDivision = "A"
elseif  SkiAge >= 85 THEN 'B' 
	AgeDivision = "B"
else
	AgeDivision = "-"
end if
					  
if Gender = "M" AND SkiAge < 18 THEN 'B' 
	SkiGender = "B"
elseif Gender = "M" AND SkiAge >= 18 THEN 'M' 
	SkiGender = "M"
elseif Gender = "F" AND SkiAge < 18 THEN 'G' 
	SkiGender = "G"
elseif Gender = "F" AND SkiAge >= 18 THEN 'W' 
	SkiGender = "W"
else 
	SkiGender = "-"
end if					  

CalculateDivision = SkiGender & AgeDivision
				  
End Function


Function MyFormatNumber(InNumber, NumberofPositions)	
	if len(InNumber) = 0 then
		MyFormatNumber = ""
	elseif isnull(InNumber) then
		MyFormatNumber = ""	
	elseif not isnumeric(InNumber) then
		MyFormatNumber = ""	
	else
		WorkingString = formatnumber(InNumber,6)
		'now remove any commas
		WorkingString2 = ""
		for x = 1 to len(WorkingString) 
			if mid(WorkingString, x ,1) = "," then
				'skip
			else
				WorkingString2 = WorkingString2 & mid(WorkingString,x,1)
			end if
		next
		MyFormatNumber = left(WorkingString2,6)
	end if
end Function


'	Pull Member Data for selected MemberID

Dim RS, sSQL, MemData, ContactDtls, OfclRtg, CurrDiv, SlmSco, SlmRtg, TrkSco, TrkRtg, JmpSco, JmpRtg
Set RS = Server.CreateObject("ADODB.RecordSet")
RS.ActiveConnection = objConn

'	RS.open "SELECT * FROM [Export Members to Excel] Where PersonIDWithCheckDigit='" & Request("MemberID") & "';" 

sSQL = "SELECT Mem.PersonIDWithCheckDigit as MemberID, Mem.LastName, Mem.FirstName, Mem.City," 
sSQL = sSQL & " Mem.State, Left(Mem.Sex,1) as Gender, Mem.EffectiveTo as ExpDt," 
sSQL = sSQL & " Mem.BirthDate, Convert(char(10),Mem.Birthdate,111) as Bdate,"
sSQL = sSQL & " Mem.DivisionCode1 + '/' + Mem.DivisionCode2 as SptsDiv,"
sSQL = sSQL & " Mem.Address1, Mem.Address2, Mem.Zip, Mem.FederationCode as FedCode, Mem.Email,"
sSQL = sSQL & " Mem.Phone as HomePhone, Mem.BusinessPhone as BusPhone, Mem.MobilePhone,"
sSQL = sSQL & " Mem.MembershipTypeCode as MemType, Mem.WaiverStatusID as Waiver, Typ.TypeCode,"
sSQL = sSQL & " Typ.CanSkiInTournaments as CanSki, Typ.CanSkiInGRTournaments as CanSkiGR"
sSQL = sSQL & " FROM USAWaterski.dbo.members as Mem, USAWaterski.dbo.MembershipTypes as Typ"
sSQL = sSQL & " WHERE Mem.MembershipTypeCode = Typ.MemberShipTypeID"
sSQL = sSQL &+ " AND Mem.PersonIDWithCheckDigit='" & Request("MemberID") & "';" 

RS.open sSQL

	SkiAge = Session("TournamentYear") - DATEPART("yyyy", RS("BirthDate")) - 1

	MemData = left(RS("MemberID"),3) & "-" & mid(RS("MemberID"),4,2) & "-" & right(RS("MemberID"),4)
	MemData = Memdata & "&#9;" & left(RS("lastname"),13)
	MemData = Memdata & "&#9;" & left(RS("firstname"),11)
	MemData = Memdata & "&#9;&#9;" & CalculateDivision(SkiAge, Left(RS("Gender"),1))
	MemData = Memdata & "&#9;" & SkiAge
	MemData = Memdata & "&#9;" & left(RS("city"),13)
	MemData = Memdata & "&#9;" & RS("State")

	ContactDtls = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Member Name:&nbsp;&nbsp; " & RS("firstname") & " " & RS("lastname") 
	ContactDtls = ContactDtls & "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Address:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; " & RS("Address1") & "&nbsp; " & RS("Address2")
	ContactDtls = ContactDtls & "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; City/ST/Zip:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; " & RS("city") & ", " & RS("State") & "&nbsp; " & RS("Zip")
	ContactDtls = ContactDtls & "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Birthdate:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; " & RS("BDate") & "&nbsp;&nbsp;&nbsp;&nbsp; Gender: &nbsp;" & Left(RS("Gender"),1)
	ContactDtls = ContactDtls & "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Federation:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; " & RS("FedCode") & "&nbsp;&nbsp;&nbsp; Type: " & RS("TypeCode") & "&nbsp;&nbsp;&nbsp; SD: " & RS("SptsDiv")
	ContactDtls = ContactDtls & "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; E-Mail Adrs:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; " & RS("EMail")
	ContactDtls = ContactDtls & "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Home Phone:&nbsp;&nbsp;&nbsp;&nbsp; " & RS("HomePhone")
	ContactDtls = ContactDtls & "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Bus Phone:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; " & RS("BusPhone")
	ContactDtls = ContactDtls & "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Mobile Phone:&nbsp;&nbsp;&nbsp; " & RS("MobilePhone")
	
	%>

	<TABLE BORDER="0">

			<TR>
				<td width="14">&nbsp;</td>
				<td width="14">&nbsp;</td>
				<td>&nbsp;</td>
			</TR>

			<TR>
				<td>&nbsp;</td>
				<td colspan=2><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Here
					is the information for the selected member, to be copied into
					an additional row in your Registration Template spreadsheet.&nbsp;
					Brief instructions appear with the member's information below.&nbsp; 
					<a href="OneMemberInstructions.htm" target="_blank">Click Here</a>
					to display more comprehensive instructions 
					(in a separate window).<br>&nbsp;</font></td>
			</TR>

			<TR> 
				<td>&nbsp;</td>
				<td colspan=2><font size="2" face="Verdana, Arial, Helvetica, sans-serif">1.&nbsp;&nbsp; 
					Hilite/copy the values below, then <b>Paste Special (text)</b> into the 
					col <b>A</b> cell in a new row, that you have inserted into your Registration 
					Template.<br>&nbsp;</font></td>
			</TR>

			<TR> 
				<td>&nbsp;</td>
				<td colspan=2><TABLE BORDER="1" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#DDDDFF"><TR>
					<td VALIGN="Middle"><font size="2"><pre><%=MemData%></pre></font></td>
				</TR></TABLE></td>
			</TR>

<%


IF RS("ExpDt") < cdate(session("tournamentdate")) THEN 
		MemData = "Membership expired " & datepart("m",RS("ExpDt")) & "/" & datepart("d",RS("ExpDt")) & "/" & datepart("yyyy",RS("ExpDt"))
		IF RS("CanSki") = False THEN
			MemData = MemData + ", Cost to Renew/Upgrade is $" & GetMemberShipPrice(session("tournamentdate"), rs("MemType"), "B")
		ELSE
			MemData = MemData + ", Cost to Renew is $" & GetMemberShipPrice(session("tournamentdate"), rs("MemType"), "R")
		END IF
ELSE 
		IF RS("CanSkiGR") = True THEN
			MemData = "GrassRts OK, Upgr Reqd for Prem, Cost is $" & GetMemberShipPrice(session("tournamentdate"), rs("MemType"), "U")
		ELSEIF RS("CanSki") = False THEN
			MemData = "Membership Upgrade required, Cost is $" & GetMemberShipPrice(session("tournamentdate"), rs("MemType"), "U")
		ELSE
			MemData = ""
		END IF
END IF

IF MemData = "" and RS("Waiver") = 0 THEN
		MemData = "USA-WS Annual Participation Waiver Needs To Be Signed."
END IF
			
IF MemData > "" THEN %>

			<TR> 
				<td>&nbsp;</td>
				<td colspan=2><font size="2" color="red" face="Verdana, Arial, Helvetica, sans-serif"><br>2.&nbsp;&nbsp; 
					<b><%=MemData%></b><br>&nbsp;</font></td>
			</TR>

<% ELSE 
			MemData = " through " & datepart("m",RS("ExpDt")) & "/" & datepart("d",RS("ExpDt")) & "/" & datepart("yyyy",RS("ExpDt"))	
	%>

			<TR> 
				<td>&nbsp;</td>
				<td colspan=2><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><br>2.&nbsp;&nbsp; 
					Membership Status is <b>OK to Ski</b><%=MemData%>.<br>&nbsp;</font></td>
			</TR>

<% END IF 
		
RS.Close

'	Get Officials Ratings for this member

sSQL = "SELECT case when OD.RatingType_ID = 3 then 1 when OD.RatingType_ID = 1 then 2"
sSQL = sSQL + " when OD.RatingType_ID = 2 then 3 else 4 end as RtgType,"
sSQL = sSQL + " MAX(convert(char(1),LV.LevelOrderforTemplate) + "
sSQL = sSQL + " LV.LevelAbbreviationforTemplate) AS RtgLvl"
sSQL = sSQL + " FROM USAWaterski.dbo.Officials OD"
sSQL = sSQL + " INNER JOIN USAWaterski.dbo.Level LV ON OD.Level_ID = LV.Level_ID"
sSQL = sSQL + " WHERE OD.DivisionCode in ('AWS','USA') AND	OD.RatingType_ID in (1,2,3,9)"
sSQL = sSQL + " AND LV.LevelOrderforTemplate IS NOT NULL AND OD.PersonID = "
sSQL = sSQL + Right(Request("MemberID"),8) & "GROUP BY OD.RatingType_ID"

OfclRtg = "----"

RS.Open sSQL

DO WHILE NOT RS.EOF 
	OfclRtg = Left(OfclRtg, RS("RtgType")-1) & Right(RS("RtgLvl"),1) & Right(OfclRtg, 4-RS("RtgType"))
	RS.MoveNext 
 	LOOP

RS.Close


'	Prepare Query to get Ranking Scores and Levels for this member

sSQL = "SELECT RT.Event, RT.Div, Coalesce(ED.DivElite,RT.AWSA_Rat) as Rtg,"
sSQL = sSQL & " RankScore FROM Cobra00025.USAWSRank.Rankings as RT"
sSQL = sSQL & " LEFT JOIN (Select Event, Max(DivElite) as DivElite"
sSQL = sSQL & " FROM Cobra00025.USAWSRank.EliteDates WHERE SkiYearID = 1"
sSQL = sSQL & " and MemberID = '" & Request("MemberID") & "' group by Event)"
sSQL = sSQL & " as ED on Left(RT.Event,1) = ED.Event"
sSQL = sSQL & " WHERE RT.MemberID = '" & Request("MemberID")
sSQL = sSQL & "' AND RT.RankScore is not null AND RT.SkiYearID = 1"
sSQL = sSQL & " AND left(Div,1) in ('B','G','M','W','O') Order by RT.Div, RT.Event;"

RS.Open sSQL

IF RS.EOF THEN 
	
	IF OfclRtg = "----" THEN %>

			<TR> 
				<td>&nbsp;</td>
				<td colspan=2><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><br>3.&nbsp;&nbsp; 
					No Ranking / Seeding information exists for this Member.<br>&nbsp;</font></td>
			</TR>

	<%	ELSE	%>

			<TR> 
				<td>&nbsp;</td>
				<td colspan=2><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><br>3.&nbsp;&nbsp; 
					No Ranking / Seeding information exists for this Member.&nbsp; 
					But please hilite/copy the Official Rating values below, then <b>Paste Special 
					(text)</b> that into the col <b>L</b> cell in the template.<br>&nbsp;</font></td>
			</TR>

			<TR> 
				<td>&nbsp;</td>
				<td colspan=2><TABLE BORDER="1" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#DDDDFF"><TR>
					<td VALIGN="Middle"><font size="2"><pre><%=OfclRtg%></pre></font></td>
				</TR></TABLE></td>
			</TR>

	<%	END IF	%>



<%	ELSE	%>
	
			<TR> 
				<td>&nbsp;</td>
				<td colspan=2><font size="2" face="Verdana, Arial, Helvetica, sans-serif">3.&nbsp;&nbsp; 
					Hilite/copy the Ranking/Level values below for the Applicable Division,
					then <b>Paste Special (text)</b> into the col <b>L</b> cell in the template.<br>&nbsp;</font></td>
			</TR>

<%

' Loop over Divisions for this Member 

DO UNTIL RS.EOF

	CurrDiv = RS("Div"): SlmSco = "": SlmRtg = "": TrkSco = "": TrkRtg = "": JmpSco = "": JmpRtg = "": OvrRtg = ""
	
' Loop over Events for this Member / Division

	DO WHILE RS("Div") = CurrDiv
		
		IF left(RS("Event"),1) = "S" THEN 
			SlmSco = MyFormatNumber(RS("RankScore"),6): SlmRtg = RS("Rtg")
		ELSEIF left(RS("Event"),1) = "T" THEN 
			TrkSco = MyFormatNumber(RS("RankScore"),6): TrkRtg = RS("Rtg")
		ELSEIF left(RS("Event"),1) = "J" THEN 
			JmpSco = MyFormatNumber(RS("RankScore"),6): JmpRtg = RS("Rtg")
		ELSE
			OvrRtg = RS("Rtg")
		END IF
	
		RS.moveNEXT
	  IF RS.EOF THEN EXIT DO

  	LOOP

' Now present the Rankings/Levels for this Member / Division

	MemData = OfclRtg & "&#9;" & SlmSco & "&#9;" & TrkSco & "&#9;" & JmpSco
	MemData = MemData & "&#9;" & SlmRtg & "&#9;" & TrkRtg & "&#9;" & JmpRtg & "&#9;" & OvrRtg
		
%>

			<TR> 
				<td>&nbsp;</td>
				<td>&nbsp;<b><%=CurrDiv%></b></td>
				<td><TABLE BORDER="1" CELLPADDING="3" CELLSPACING="0" BGCOLOR="#DDDDFF"><TR>
					<td VALIGN="Middle"><font size="2"><pre><%=MemData%></pre></font></td>
				</TR></TABLE></td>
			</TR>

<% 

	LOOP

	END IF
	
RS.Close

'	Finally Display additional contact details for selected member, if authorized

IF left(Session("UserName"),1) > "9" THEN
%>
			<TR> 
				<td>&nbsp;</td>
				<td colspan=2><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><br>4.&nbsp;&nbsp; 
					Additional Contact Details for this Member:<br>&nbsp;<br>
				<%=ContactDtls%></font></td>
			</TR>
<%
END IF


%>

	</TABLE>

		
	<TABLE ALIGN="CENTER" WIDTH=70%><TR><TD>&nbsp;</TD></TR>

			<TD width=25% align=center>
		<form action="LookupMembers.asp?FormStatus=newsearch" method="post">
		<input type="submit" style="width:9em" value="New Search"></form>
			</TD>

	    <td width=25% align=center>     				
		<form action="Index.asp" method="post">
	  <input type="submit" style="width:9em" value="Quit"></form>
  	  </td>

	    <td width=25% align=center valign=center><font size="2" face="Verdana, Arial, Helvetica, sans-serif">     				
		<form action="OneMemberInstructions.htm" method="post" target="_blank">
		<input type="submit" style="width:9em" value="Instructions"></form>
	    </td>

	  </tr></table>
	
	</td>
  </tr>
</table>
</body>
</html>





