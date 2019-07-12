<!--#include file="settingsHQ.asp"-->


<%
'Dim sID, i
'Dim sTourDate
'Dim testno, sRunByWhat, sTableName

Dim sRunByWhat

Dim RunByWhat, FormStatus, sSequence

Dim sMemberID, sLastName, sFirstName, sMembSex, sMembCity, sMembState, sMembAge, sMembTypeID, sCanSkiTour, sMembTypeCode, sEffectiveTo, sMembBirth
Dim TourSelected
Dim TourCount




sTableName = TRIM(Request("fTableName"))
sSequence = Request("sSequence")
IF sSequence="" THEN sSequence=2

sTourDate = "08/21/2005"
sTourID = "07W999A"
sMemberID = "000001151" 

TourSelected=Request("TourDrop")
'Markdebug("TourSelected="&TourSelected)



' Execute SUBroutine

sRunByWhat=TRIM(Request("sRunByWhat"))
'sRunByWhat = 8
IF sRunByWhat="" THEN sRunByWhat=2

'markdebug("sRunByWhat = "&sRunByWhat)


DisplayDrops

SELECT CASE sRunByWhat

	CASE 1
		UpdateFile

	CASE 2

		IF sTableName <> "" THEN
			FindOutAllData
		END IF
	CASE 3
		IF sTableName <> "" THEN
			FindOutFieldNames
		END IF
	CASE 4
		QuickTest

	CASE 5
		QuickTest2

	CASE 6
		DeleteData

	
	CASE 7
		AlterMyTable

	CASE 8
		CreateMyTable
		'FindOutFieldNames
	CASE 9
		AppendFile
	CASE 10
		NCWSATable

	CASE 11
		TestONE

	CASE 12
		TestTWO


END SELECT









' -------------
   SUB TESTTWO
' -------------


' --- original

PathtoTRA = Server.mappath("/")&"\rankings\" 

' filein=PathToTRA & "uploads\" & request("file")  
' CSVfile=PathToTRA & "uploads\" & left(Request("file"),7) & ".csv"

'filein="http://usawaterski.org\rankings\uploads\" & request("file")  
'CSVfile="http://usawaterski.org\rankings\uploads\" & left(Request("file"),7) & ".csv"

filein="http://usawaterski.org\rankings\uploads\07S160E.wsp" 
CSVfile="http://usawaterski.org\rankings\uploads\07S160E.csv"



markdebug("From Utility_Mark - filein= "& filein)
markdebug("From Utility_Mark - CSVfile= "& CSVFile)


  ' We will delete the CSV when we are done processing.
    set objFSO=server.createobject("scripting.filesystemObject")
  objfso.CopyFile filein, CSVfile, 1




END SUB




' ------------
  SUB TESTONE
' ------------

Dim CurrLeagueID, TypeAList
CurrLeagueID="NATL2008"

SET rsTypeA=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT TourID FROM "&LeagueToursTableName
sSQL = sSQL + " WHERE LeagueID='"&CurrLeagueID&"' AND TourType='A'"

'response.write(sSQL)
'response.end
rsTypeA.open sSQL, SConnectionToTRATable

TypeAList="("
DO WHILE NOT rsTypeA.eof
	IF TRIM(TypeAList)="(" THEN
		TypeAList=TypeAList&rsTypeA("TourID")		
	ELSE
		TypeAList=TypeAList&", "&rsTypeA("TourID")
	END IF
	rsTypeA.movenext
LOOP
TypeAList=TypeAList+")"

response.write("<br>TypeAList="&TypeAList)


END SUB






' ------------------
   SUB DisplayDrops
' ------------------


%>

<TABLE ALIGN="CENTER" WIDTH=100%>

  <form action="utility_mark.asp" method="post">


   <td>
	<font size=<% =fontsize2 %> face=<% =font1 %>>Table Name: fTableName</font> 
	<select name='fTableName'>
	<option value ='0'<%IF sTableName = "0" THEN Response.Write(" selected ")%>>Select Table </Option><br>
	<option value ='1'<%IF sTableName = "1" THEN Response.Write(" selected ")%>>RegGenTableName</Option><br>
	<option value ='2'<%IF sTableName = "2" THEN Response.Write(" selected ")%>>RegTransTableName</Option><br>
	<option value ='3'<%IF sTableName = "3" THEN Response.Write(" selected ")%>>RegTempTableName</Option><br>
	<option value ='4'<%IF sTableName = "4" THEN Response.Write(" selected ")%>>RegPWTableName</Option><br>
	<option value ='5'<%IF sTableName = "5" THEN Response.Write(" selected ")%>>SptsGrpTableName</Option><br>
	<option value ='6'<%IF sTableName = "6" THEN Response.Write(" selected ")%>>TourGenTableName</Option><br>
	<option value ='7'<%IF sTableName = "7" THEN Response.Write(" selected ")%>>CCLogTableName</Option><br>
	<option value ='8'<%IF sTableName = "8" THEN Response.Write(" selected ")%>>BioTableName</Option><br>
	<option value ='9'<%IF sTableName = "9" THEN Response.Write(" selected ")%>>DivisionsTableName</Option><br>
	<option value ='10'<%IF sTableName = "10" THEN Response.Write(" selected ")%>>RankTableName</Option><br>
	<option value ='11'<%IF sTableName = "11" THEN Response.Write(" selected ")%>>TrafficTableName</Option><br>
	<option value ='12'<%IF sTableName = "12" THEN Response.Write(" selected ")%>>12 TEST CODE</Option><br>
	<option value ='13'<%IF sTableName = "13" THEN Response.Write(" selected ")%>>13 Non-Entered</Option><br>
	<option value ='14'<%IF sTableName = "14" THEN Response.Write(" selected ")%>>14 Member History Table</Option><br>
	<option value ='15'<%IF sTableName = "15" THEN Response.Write(" selected ")%>>15 Overall Scores</Option><br>
	<option value ='16'<%IF sTableName = "16" THEN Response.Write(" selected ")%>>16 Raw Other Scores</Option><br>
	<option value ='17'<%IF sTableName = "17" THEN Response.Write(" selected ")%>>17 Users999</Option><br>
	<option value ='18'<%IF sTableName = "18" THEN Response.Write(" selected ")%>>18 Users Sanction</Option><br>
	<option value ='19'<%IF sTableName = "19" THEN Response.Write(" selected ")%>>19 Sanction Table</Option><br>
	<option value ='20'<%IF sTableName = "20" THEN Response.Write(" selected ")%>>20 Sanction VIEW</Option><br>
	<option value ='21'<%IF sTableName = "21" THEN Response.Write(" selected ")%>>21 Leagues</Option><br>
	<option value ='22'<%IF sTableName = "22" THEN Response.Write(" selected ")%>>22 Teams</Option><br>
	</select>
    </td>

  <td>
	<font size=<% =fontsize2 %> face=<% =font1 %>>Operation: sRunByWhat</font> 
	<select name='sRunByWhat'>
	<option value =''<%IF sRunByWhat = "" THEN Response.Write(" selected ")%>>0 Select Operation</Option><br>
	<option value ='2'<%IF sRunByWhat = "2" THEN Response.Write(" selected ")%>>2 Data in Table</Option><br>
	<option value ='3'<%IF sRunByWhat = "3" THEN Response.Write(" selected ")%>>3 Field Names</Option><br>
	<option value ='4'<%IF sRunByWhat = "4" THEN Response.Write(" selected ")%>>4 QuickTest</Option><br>
	<option value ='6'<%IF sRunByWhat = "6" THEN Response.Write(" selected ")%>>6 Delete Data</Option><br>
	<option value ='7'<%IF sRunByWhat = "7" THEN Response.Write(" selected ")%>>7 Alter Table</Option><br>
	<option value ='8'<%IF sRunByWhat = "8" THEN Response.Write(" selected ")%>>8 Create Table</Option><br>
	<option value ='9'<%IF sRunByWhat = "9" THEN Response.Write(" selected ")%>>9 Append to Table</Option><br>
	<option value ='1'<%IF sRunByWhat = "1" THEN Response.Write(" selected ")%>>1 Update Table</Option><br>
	<option value ='11'<%IF sRunByWhat = "11" THEN Response.Write(" selected ")%>>11 Test Sub 1</Option><br>

	</select>
  </td>

  <td>
	<font size=<% =fontsize2 %> face=<% =font1 %>>Sort By:</font>
	 <select name='sSequence'>
	<option value ='0'<%IF sSequence = "0" THEN Response.Write(" selected ")%>>Alphabetic</Option><br>
	<option value ='1'<%IF sSequence = "1" THEN Response.Write(" selected ")%>>MemberID</Option><br>
	<option value ='2'<%IF sSequence = "2" THEN Response.Write(" selected ")%>>Date</Option><br>
	</select> 
  </td>
  <tr>
    <td><%
      'LoadTourDropDown	%>
    </td>
  </tr>	

  <center><input type="submit" value="Continue"></center>
</form>

</tr>
</TABLE>

<%

END SUB




'----------------------------
  SUB LoadTourDropDown
'----------------------------



SET rsTourList=Server.CreateObject("ADODB.recordset")

sSQL = "SELECT ST.TSanction, ST.TName FROM "&SanctionTableName&" AS ST"
sSQL = sSQL + " JOIN "&TourGenTableName&" AS TG ON LEFT(TG.TourID,6)=LEFT(ST.TSanction,6) WHERE LEFT(ST.TSanction,2)='07'" 
sSQL = sSQL + " ORDER BY ST.TSanction DESC"

rsTourList.open sSQL, SConnectionToTRATable

%>
<select name='TourDrop'>
<%

IF NOT rsTourList.eof THEN 
  	rsTourList.movefirst

  	DO WHILE NOT rsTourList.eof
		IF TRIM(rsTourList("TSanction")) = TourSelected THEN %>
			<option value=<%=rsTourList("TSanction")%> selected></option>
			<br><%
    		ELSE  %>
			<option value=<%=rsTourList("TSanction")%>></option>
			<br><%
		END IF	

		rsTourList.moveNEXT
	LOOP
ELSE
	response.write("<option value =""None"" selected>None Available</option>")
END IF  %>

</select><%

rsTourList.close
















END SUB






' ------------------------------------
    SUB FindOutAllData 
' ------------------------------------

Dim tTable

Dim RowCount, i
rowcount = -1

sMemberID="000001151"
sTourID="07W999A"



SET rs=Server.CreateObject("ADODB.recordset")

SELECT CASE sTableName
	CASE "1"
		sSQL = "SELECT RGEN.MemberID, MEM.FirstName, MEM.LastName, Div1, Div2, Div3, Div4, SignWaiver, WaiverCode, RegisterDate"
		sSQL = sSQL + ", EntryType, RGEN.Status, TotalCharges"
		sSQL = sSQL + ", PW.email"
		sSQL = sSQL + "  FROM "&RegGenTableName&" AS RGEN"
		sSQL = sSQL + " JOIN "&MemberTableName&" AS MEM ON RGEN.MemberID = MEM.PersonIDWithCheckDigit" 
		sSQL = sSQL + " JOIN "&RegPWTableName&" AS PW ON PW.MemberID = RGEN.MemberID"
 
		IF TourSelected<>"" THEN
'			sSQL = sSQL + " WHERE RGEN.TourID="&TourSelected
		END IF
		

		IF sSequence="0" THEN
			sSQL = sSQL + " ORDER BY MEM.LastName, MEM.FirstName"
		ELSEIF sSequence="1" THEN
			sSQL = sSQL + " ORDER BY RGEN.MemberID"
		ELSE
			sSQL = sSQL + " ORDER BY RGEN.RegisterDate DESC"
		END IF


		rs.open sSQL, sConnectionToTRATable, 3, 1
	CASE "2"

		sSQL = "SELECT * FROM "&RegTransTableName&" AS RTRAN"
		sSQL = sSQL + " JOIN "&MemberTableName&" AS MEM ON RTRAN.MemberID = MEM.PersonIDWithCheckDigit" 

'		sSQL = sSQL + " WHERE Left(TourID,6) = '" & SQLClean(left(sTourID,6)) & "' AND MemberID = '"&sMemberID&"'"
'		sSQL = sSQL + " WHERE WaiverCode <> ''"


		IF sSequence="0" THEN
			sSQL = sSQL + " ORDER BY MEM.LastName, MEM.FirstName"
		ELSEIF sSequence="1" THEN
			sSQL = sSQL + " ORDER BY RTRAN.MemberID"
		ELSEIF sSequence="2" THEN
			sSQL = sSQL + " ORDER BY RTRAN.TransDate DESC, RTRAN.TransNo ASC"
		END IF



		rs.open sSQL, sConnectionToTRATable, 3, 1

	CASE "3"
		sSQL = "SELECT * FROM "&RegTempTableName&" AS TEMP" 
		sSQL = sSQL + " JOIN "&MemberTableName&" AS MEM ON TEMP.MemberID = MEM.PersonIDWithCheckDigit" 
		'sSQL = sSQL + " WHERE Left(TEMP.TourID,6) = '" & SQLClean(left(sTourID,6)) & "' AND TEMP.MemberID = '"&sMemberID&"'"

		IF sSequence="0" THEN
			sSQL = sSQL + " ORDER BY MEM.LastName, MEM.FirstName"
		ELSEIF sSequence="1" THEN
			sSQL = sSQL + " ORDER BY TEMP.MemberID"
		ELSEIF sSequence="2" THEN
			sSQL = sSQL + " ORDER BY TEMP.RegisterDate DESC"
		END IF

		rs.open sSQL, sConnectionToTRATable, 3, 1

	CASE "4"
		sTourID = "07W999A"
		sSQL = "SELECT MEM.FirstName, MEM.LastName, PW.MemberID, PW.Password, PW.CreateDate, PW.Email, PW.Status,"
		sSQL = sSQL + " RGEN.TourID, RGEN.Div1, RGEN.Div2, RGEN.Div3, RGEN.Div4"
		sSQL = sSQL + " FROM "&RegPWTableName&" AS PW "

		sSQL = sSQL + " JOIN "&MemberTableName&" AS MEM ON PW.MemberID = MEM.PersonIDWithCheckDigit" 
		sSQL = sSQL + " LEFT JOIN "&RegGenTableName&" AS RGEN ON RGEN.MemberID = PW.MemberID "
		sSQL = sSQL + " AND LEFT(RGEN.TourID,6) = '"&LEFT(sTourID,6)&"'"
		IF sSequence="0" THEN
			sSQL = sSQL + " ORDER BY MEM.LastName"
		ELSEIF sSequence="1" THEN
			sSQL = sSQL + " ORDER BY PW.MemberID"
		ELSE
			sSQL = sSQL + " ORDER BY PW.CreateDate DESC"
		END IF
		rs.open sSQL, sConnectionToTRATable, 3, 1

	CASE "5"
		sSQL = "SELECT * FROM " &SptsGrpTableName 
		rs.open sSQL, sConnectionToSanctionTable, 3, 1

	CASE "6"
		sSQL = "SELECT * FROM "&TourGenTableName
		rs.open sSQL, sConnectionToTRATable, 3, 1

	CASE "7"
		sSQL = "SELECT CC.MemberID, MEM.FirstName, MEM.LastName, Last4Card, Amount, OrderNo, CC.FirstName, CC.LastName, TransDate, Result, Checkno, PayType"
		sSQL = sSQL + " FROM "&CCLogTableName&" AS CC"
		sSQL = sSQL + " JOIN "&MemberTableName&" AS MEM ON CC.MemberID = MEM.PersonIDWithCheckDigit"

		 
		IF sSequence="0" THEN
			sSQL = sSQL + " ORDER BY MEM.LastName, MEM.FirstName"
		ELSEIF sSequence="1" THEN
			sSQL = sSQL + " ORDER BY CC.MemberID"
		ELSEIF sSequence="2" THEN
			sSQL = sSQL + " ORDER BY CC.OrderNo DESC"
		END IF

		rs.open sSQL, sConnectionToTRATable, 3, 1

	CASE "8"
		sSQL = "SELECT * FROM "&BioTableName&" AS BIO"
		sSQL = sSQL + " JOIN "&MemberTableName&" AS MEM ON BIO.MemberID = MEM.PersonIDWithCheckDigit" 
		IF sSequence="0" THEN
			sSQL = sSQL + " ORDER BY MEM.LastName, MEM.FirstName"
		ELSEIF sSequence="1" THEN
			sSQL = sSQL + " ORDER BY BIO.MemberID"
		END IF
		rs.open sSQL, sConnectionToTRATable, 3, 1

	CASE "9"
		sSQL = "SELECT * FROM "&DivisionsTableName
		rs.open sSQL, sConnectionToTRATable, 3, 1

	CASE "10"
		sSQL = "SELECT * FROM "&RankTableName
		sSQL = sSQL + " WHERE MemberID='600118259'"
		rs.open sSQL, sConnectionToTRATable, 3, 1


	CASE "11"
		sSQL = "SELECT * FROM "&TrafficTableName&" Order by ActivityDate"
		rs.open sSQL, sConnectionToTRATable, 3, 1

	CASE "12"


		SET rs=Server.CreateObject("ADODB.recordset")

		sSQL = "SELECT * FROM "&SanctionTableName&" AS ST"

		sSQL = sSQL + " WHERE left(TSanction,6) = '07M039'"
		'sSQL = sSQL + " WHERE left(TSanction,6) = '07W155'"

		rs.open sSQL, sConnectionToTRATable, 3, 1	






	CASE "13"

		' ----- Start of query to determine who is not completing their registration  -----
		sSQL = "SELECT RTEMP.MemberID, MEM.FirstName, MEM.LastName, RTEMP.RegisterDate, PW.Email, RTEMP.Div1, RTEMP.Div2, RTEMP.Div3, RGEN.Event1, RGEN.Event2, RGEN.Event3, RGEN.WaiverCode"
		sSQL = sSQL + " FROM "&RegTempTableName&" AS RTEMP"

		sSQL = sSQL + " LEFT JOIN "&RegGenTableName&" AS RGEN ON RGEN.MemberID = RTEMP.MemberID"
		sSQL = sSQL + " LEFT JOIN "&RegPWTableName&" AS PW ON RTEMP.MemberID = PW.MemberID"
		sSQL = sSQL + " LEFT JOIN "&MemberTableName&" AS MEM ON RTEMP.MemberID = MEM.PersonIDWithCheckDigit"
		sSQL = sSQL + " ORDER BY RTEMP.RegisterDate DESC"
		rs.open sSQL, sConnectionToTRATable, 3, 1

	CASE "14"



	Set SQLConnect = CreateObject("ADODB.Connection")
	SQLConnect.Open Application("HQSQLConn")
	Set RS = SQLConnect.Execute("SELECT [Person ID], Email FROM tblPeople")



	CASE "16"
		SET rs=Server.CreateObject("ADODB.recordset")

		sSQL = "SELECT * FROM "&RawScoresTableName&" AS RSO"
		sSQL = sSQL + " WHERE MemberID = '000001151'"

		rs.open sSQL, sConnectionToTRATable, 3, 1	

	CASE "17"
		SET rs=Server.CreateObject("ADODB.recordset")
		sSQL = "SELECT TOP 500 * FROM usawaterski.dbo.Users999 AS RT "
		sSQL = sSQL + " ORDER BY LastName"
		rs.open sSQL, sConnectionToTRATable, 3, 1	

	CASE "18"
		sSQL = "SELECT TOP 500 * FROM sanctions.dbo.users"
		sSQL = sSQL + " ORDER BY LastName"
		rs.open sSQL, sConnectionToSanctionTable, 3, 1

	CASE "19"
		sSQL = "SELECT TOP 100 * FROM "&SanctionTableName
		rs.open sSQL, sConnectionToSanctionTable, 3, 1

	CASE "20"
		sTAID="08U033"
			SQL = "SELECT * FROM fn_TschedulRegFieldsXTournAppID('" & sTAID & "')"

			sConn = "Provider=SQLOLEDB;SERVER=jaguar.epolk.net;Database=Sanctions;uid=Sanctions_Admin;pwd=qej8h7w34w"
			'sConn = Application("SanctionConn")
			
			Set Conn = Server.CreateObject("ADODB.Connection")
			Set RS = server.CreateObject("ADODB.Recordset")
			Conn.Open sConn
			Set RS = Conn.Execute (SQL)

	CASE "21"
		sSQL = "SELECT TOP 100 * FROM "&LeagueTableName
		rs.open sSQL, sConnectionToTRATable, 3, 1
	CASE "22"
		sSQL = "SELECT TOP 100 * FROM "&TeamTableName
		rs.open sSQL, sConnectionToTRATable, 3, 1


END SELECT


RowCount = 0
IF NOT rs.eof THEN

	rs.movefirst  %>
	
	<table width=100% BORDER="1" >
	<tr>
	  <td><FONT Size="1">Row</font></td><%
	
	  FOR i = 0 TO rs.fields.count - 1
		%><TD ALIGN="Left" vAlign="top" nowrap><FONT COlOR="#000000" Size="1">&nbsp;<%
		Response.Write(trim(Rs.Fields(i).Name)) 
	  NEXT %>
	
	</tr><%

	  DO WHILE NOT rs.eof

		rowCount = rowCount + 1   %>

		<tr>
		  <td> 
			<FONT Size="1"><% =RowCount %></font>
		  </td><%

		FOR i = 0 TO rs.fields.count - 1

			%><TD ALIGN="Left" vAlign="top" nowrap><FONT COlOR="#000000" Size="1">&nbsp;<%

			Response.Write(trim(Rs.Fields(i).Value))

			%>&nbsp;</FONT></TD><%

		NEXT %>

		</TR><%

		rs.movenext

	  LOOP  %>
	</tr>
	<table><% 

	rs.close
	Set rs = nothing

ELSE
	response.write("No Data Found in Table")

END IF



END SUB 






' --------------------
    SUB CreateMyTable
' --------------------

OpenCon

'sSQL = "CREATE TABLE usawsrank.SkierBios (MemberID nvarCHAR(9), Region nvarCHAR(2), Address1 nvarchar(30), City nvarchar(20), State nvarchar(2), Zip nvarchar(6)"
'sSQL = sSQL + ", Email nvarCHAR(30), Phone nvarCHAR(10), Weight nvarCHAR(3), HgtFeet nvarCHAR(1), HgtInch nvarCHAR(2), SkiSinceAge nvarCHAR(2)"
'sSQL = sSQL + ", CompSinceAge nvarCHAR(2), MembSinceAge nvarCHAR(2), Club nvarCHAR(25), School nvarCHAR(25), Occup nvarCHAR(25), Career nvarCHAR(25)"
'sSQL = sSQL + ", Hobby nvarCHAR(25), Paper nvarCHAR(25), BestSlal nvarCHAR(10), BestTrick nvarCHAR(10), BestJump nvarCHAR(10), BestKnee nvarCHAR(10)"
'sSQL = sSQL + ", BestFree nvarCHAR(10), BestWake nvarCHAR(10), BestMara nvarCHAR(10), BestHydro nvarCHAR(10))"


'sSQL = sSQL + ", Event1 nvarCHAR(2), Event2 nvarCHAR(2), Event3 nvarCHAR(2), Event4 nvarCHAR(2)"
'sSQL = sSQL + ", Div1 nvarCHAR(2), Div2 nvarCHAR(2), Div3 nvarCHAR(2), Div4 nvarCHAR(2), Ramp nvarCHAR(4), Weight nvarCHAR(3)"
'sSQL = sSQL + ", EntryFee decimal(6,3), LateFee decimal(6,3), AWSEFDonation decimal(6,3), OffDisc decimal(6,3), JrDisc decimal(6,3)"
'sSQL = sSQL + ", SrDisc decimal(6,3), ClubDisc decimal(6,3), ClubCode nvarCHAR(3), TotalEntry decimal(6,3))"

'sSQL = "CREATE TABLE usawsrank.RegPasswords (MemberID nvarchar(9) NOT NULL, Password nvarchar(10), CreateDate DATETIME)"
'sSQL = sSQL + ", Address1 nvarchar(25), City nvarchar(20), State nvarchar(2), ZipCode nvarchar(6)"
'sSQL = sSQL + ", Email nvarchar(35), Last4Card nvarchar(4), ExpMonth nvarchar(2), ExpYear nvarchar(4), Amount decimal(6,2), OrderNo nvarchar(20) )"


'sSQL = "CREATE TABLE usawsrank.RegTransactions (MemberID nvarCHAR(9) NOT NULL, TourID nvarCHAR(7), TransCode nvarCHAR(3), Amount decimal(7,2), TransDate DATETIME)"
'sSQL = "DROP TABLE usawsrank.RegTransactions"


sSQL = "CREATE TABLE usawsrank.Leagues (LeagueID varCHAR(4), SptsGrpID varCHAR(3), LeagueName varchar(20), Status varchar(1))"


con.execute(sSQL)
closecon

END SUB




' --------------------
    SUB AlterMyTable
' --------------------

OpenCon
'sSQL = "ALTER TABLE "&DivisionsTableName&" DROP COLUMN SptsGrpIP"
'con.execute(sSQL)

sSQL = "ALTER TABLE "&TeamTableName&" ADD Address varchar(25), Address2 varchar(25), City varchar(25), State char(2), Zip char(5)"

' --- Use this ---
' Alter table USAWSRank.DivisionOther Alter Column DIV varchar(3)



'closecon

'opencon
'sSQL = "ALTER TABLE "&RegTemporary&" ADD Rank_Level Decimal(6,3), Natl_plc varchar(4), Regl_plc varchar(4), Reg_Ski varchar(1), RankDiv varchar(2)"
'sSQL = sSQL + " MembOverride varCHAR(3), RegionalOverride varCHAR(3)"
'sSQL = sSQL + "QfyOverride varCHAR(3)"






con.execute(sSQL)
closecon


END SUB



' --------------------
    SUB DeleteData
' --------------------

OpenCon
SELECT CASE sTableName
	CASE "1"
'		sSQL = "DELETE FROM "&RegGenTableName
'		sSQL = sSQL + " WHERE MemberID IN ('000001151')"
	CASE "2"
'		sSQL = "DELETE FROM "&RegTransTableName
'		sSQL = sSQL + " WHERE MemberID IN ('000001151')"
	CASE "3"
'		sSQL = "DELETE FROM "&RegTempTableName
'		sSQL = sSQL + " WHERE MemberID IN ('000001151')"
	CASE "4"
'		sSQL = "DELETE FROM "&RegPWTableName
'		sSQL = sSQL + " WHERE MemberID IN ('000001151')"
	CASE "5"

	CASE "6"
'		sSQL = "DELETE FROM "&TourGenTableName

	CASE "7"
'		sSQL = "DELETE FROM "&CCLogTableName
'		sSQL = sSQL + " WHERE MemberID IN ('000001151')"

	CASE "8"
'		sSQL = "DELETE FROM "&BioTableName

END SELECT


sSQL = "DELETE FROM "&TeamTableName
sSQL = sSQL + " WHERE TeamID='XXX'"
'sSQL = sSQL + " WHERE LEFT(Div,1)='Y' OR LEFT(Div,1)='X' OR LEFT(Div,1)='G' OR LEFT(Div,1)='B'"


con.execute(sSQL)
closecon

END SUB





' --------------------
  SUB AppendFile
' --------------------


OpenCon

sSQL = "INSERT INTO " & LeagueTableName
sSQL = sSQL + " (LeagueID, SptsGrpID, LeagueName, Status)"

sSQL = sSQL + " VALUES ("

sSQL = sSQL + "'CIND',"
sSQL = sSQL + "'AWS',"
sSQL = sSQL + "'Cindonway Series',"
sSQL = sSQL + "'A')"


'sSQL = sSQL + "'OM',"
'sSQL = sSQL + "'AWS')"


'sSQL = sSQL + "'S',"
'sSQL = sSQL + "'T',"
'sSQL = sSQL + "'J',"
'sSQL = sSQL + "'',"
'sSQL = sSQL + "'',"
'sSQL = sSQL + "'',"

'sSQL = sSQL + "'Wakeboard',"
'sSQL = sSQL + "'Slalom',"
'sSQL = sSQL + "'Trick',"
'sSQL = sSQL + "'Jumping',"
'sSQL = sSQL + "'',"
'sSQL = sSQL + "'',"

'sSQL = sSQL + "'')"

con.execute(sSQL)




set rs = Server.CreateObject("ADODB.Recordset")
sSQL = "SELECT TOP 1 * " 
sSQL = sSQL + " FROM " & RegPWTableName
sSQL = sSQL + " WHERE MemberID = '000001151'"
rs.open sSQL, sConnectionToTRATable, 3, 3



%><table width=100% BORDER="1" ><%

rs.movefirst


%><tr><%
FOR i = 0 TO rs.fields.count - 1
	%><TD ALIGN="Left" vAlign="top" nowrap><FONT COlOR="#000000" Size="1">&nbsp;<%
	Response.Write(trim(Rs.Fields(i).Name)) 
NEXT
%></tr><%

DO WHILE NOT rs.eof

	rowCount = rowCount + 1


	%><tr><%
	FOR i = 0 TO rs.fields.count - 1

		%><TD ALIGN="Left" vAlign="top" nowrap><FONT COlOR="#000000" Size="1">&nbsp;<%

		Response.Write(trim(Rs.Fields(i).Value))

		%>&nbsp;</FONT></TD><%

	NEXT


	%></TR><%


	rs.movenext

LOOP


%></tr>
<table>
<% 
rs.close
Set rs = nothing



END SUB





' --------------------
  SUB UpdateFile
' --------------------


sMemberID="000001151"
sTourID="07W999A"

'on error resume next

'if err.num <> 0 then
'        error handler
'end if


OpenCon

sSQL = "UPDATE "&TeamTableName

sSQL = sSQL + " SET "
sSQL = sSQL + " TeamID = CASE TeamName "


'sSQL = sSQL + " WHEN 'Arizona State Univ' THEN 'ASU'"  

sSQL = sSQL + " WHEN 'Baylor University' THEN 'BAY'" 
sSQL = sSQL + " WHEN 'Cal Poly' THEN 'CPO'"
sSQL = sSQL + " WHEN 'Cal State Northridge' THEN 'CSN'"
sSQL = sSQL + " WHEN 'Cal State Univ, Chico' THEN 'CSC'"
sSQL = sSQL + " WHEN 'Central Washington Univ' THEN 'CWU'"
sSQL = sSQL + " WHEN 'Chico State' THEN 'JJJ'"
sSQL = sSQL + " WHEN 'Clemson University' THEN 'CLE'"
sSQL = sSQL + " WHEN 'East Carolina Univ' THEN 'ECU'"
sSQL = sSQL + " WHEN 'Elon' THEN 'ELU'"
sSQL = sSQL + " WHEN 'FL Gulf Coast Univ' THEN 'FGC'"                            
sSQL = sSQL + " WHEN 'Florida Southern College' THEN 'FSC'"                            
sSQL = sSQL + " WHEN 'Grand Valley ST Univ' THEN 'GVS'"                           
sSQL = sSQL + " WHEN 'Hope College' THEN 'HOP'"                           
sSQL = sSQL + " WHEN 'Illinois State Univ' THEN 'ILS'"                           
sSQL = sSQL + " WHEN 'Indiana Univ Purdue-FW' THEN 'IND'"                           
sSQL = sSQL + " WHEN 'Iowa State University' THEN 'IWS'"                           
sSQL = sSQL + " WHEN 'Kettering University' THEN 'KET'"                           
sSQL = sSQL + " WHEN 'Long Beach State' THEN 'LBS'"                            
sSQL = sSQL + " WHEN 'Louisiana Tech Univ' THEN 'LAT'"                            
sSQL = sSQL + " WHEN 'Marquette University' THEN 'MAR'"                           
sSQL = sSQL + " WHEN 'Miami University' THEN 'MIA'"                           
sSQL = sSQL + " WHEN 'Michigan State Univ' THEN 'MST'"                           
sSQL = sSQL + " WHEN 'Missouri State Univ' THEN 'MOS'"                           
sSQL = sSQL + " WHEN 'Ohio State University' THEN 'OSU'"                           
sSQL = sSQL + " WHEN 'Ohio University' THEN 'OHI'"                           
sSQL = sSQL + " WHEN 'Purdue University' THEN 'PUR'"                           
sSQL = sSQL + " WHEN 'Rollins College' THEN 'ROL'"                            
sSQL = sSQL + " WHEN 'Sacramento State' THEN 'AU'"                            
sSQL = sSQL + " WHEN 'San Diego State Univ' THEN 'SDS'"                            
sSQL = sSQL + " WHEN 'Southern IL Univ' THEN 'SIU'"                           
sSQL = sSQL + " WHEN 'St. Olaf College' THEN 'STO'"                           
sSQL = sSQL + " WHEN 'Stephen F. Austin' THEN 'SFA'"
sSQL = sSQL + " WHEN 'Univ of NC Wilmington' THEN 'UNC'"
sSQL = sSQL + " WHEN 'Univ of WI - La Crosse' THEN 'LAX'"
sSQL = sSQL + " WHEN 'Univ of WI Oshkosh' THEN 'UWO'"
sSQL = sSQL + " WHEN 'Univ of WI-Madison' THEN 'MAD'"
sSQL = sSQL + " WHEN 'University of Dayton' THEN 'DAY'"
sSQL = sSQL + " WHEN 'University of Iowa' THEN 'IOW'"   
sSQL = sSQL + " WHEN 'UCLA' THEN 'CLA'"  
sSQL = sSQL + " WHEN 'UC San Diego' THEN 'USD'"  
sSQL = sSQL + " WHEN 'UC Santa Barbara' THEN 'CSB'"   

sSQL = sSQL + " ELSE 'XXX'"
sSQL = sSQL + " END"
con.execute(sSQL)
closecon


END SUB



' ----------------------
   SUB FindOutFieldNames
' ----------------------


' rs.fields(i).type


' Determines fields names of whatever table is shown in the FROM statement

SET rs=Server.CreateObject("ADODB.recordset")

SELECT CASE sTableName
	CASE "1"
		sSQL = "SELECT * FROM "&RegGenTableName
		sSQL = sSQL + " WHERE Left(TourID,6) = '" & SQLClean(left(sTourID,6)) & "' AND MemberID = '"&sMemberID&"'"
		rs.open sSQL, sConnectionToTRATable, 3, 1
	CASE "2"
		sSQL = "SELECT * FROM "&RegTransTableName
		sSQL = sSQL + " WHERE Left(TourID,6) = '" & SQLClean(left(sTourID,6)) & "' AND MemberID = '"&sMemberID&"'"
		sSQL = sSQL + " ORDER BY TransDate, TransNo DESC"
		rs.open sSQL, sConnectionToTRATable, 3, 1

	CASE "3"
		sSQL = "SELECT * FROM "&RegTempTableName
		sSQL = sSQL + " WHERE Left(TourID,6) = '" & SQLClean(left(sTourID,6)) & "' AND MemberID = '"&sMemberID&"'"
		rs.open sSQL, sConnectionToTRATable, 3, 1
	CASE "4"
		sSQL = "SELECT * FROM "&RegPWTableName
		rs.open sSQL, sConnectionToTRATable, 3, 1

	CASE "5"
		sSQL = "SELECT * FROM " &SptsGrpTableName 
		rs.open sSQL, sConnectionToSanctionTable, 3, 1

	CASE "6"
		sSQL = "SELECT * FROM "&TourGenTableName
		rs.open sSQL, sConnectionToTRATable, 3, 1

	CASE "7"
		sSQL = "SELECT * FROM "&CCLogTableName
		rs.open sSQL, sConnectionToTRATable, 3, 1
	CASE "8"
		sSQL = "SELECT * FROM "&BioTableName
		rs.open sSQL, sConnectionToTRATable, 3, 1

	CASE "9"
		sSQL = "SELECT * FROM "&DivisionsTableName
		rs.open sSQL, sConnectionToTRATable, 3, 1

	CASE "10"
		sSQL = "SELECT * FROM "&RankTableName
		rs.open sSQL, sConnectionToTRATable, 3, 1

	CASE "14"

		Set SQLConnect = CreateObject("ADODB.Connection")
		SQLConnect.Open Application("HQSQLConn")
		Set RS = SQLConnect.Execute("SELECT * FROM [Membership History]")

	CASE "15"
		sSQL = "SELECT * FROM "&OverallScoresTableName
		rs.open sSQL, sConnectionToTRATable, 3, 1

	CASE "16"
		sSQL = "SELECT * FROM "&RawScoresTableName
		rs.open sSQL, sConnectionToTRATable, 3, 1

	CASE "17"
		sSQL = "SELECT * FROM usawaterski.dbo.users999"
		rs.open sSQL, sConnectionToMemberTable, 3, 1

	CASE "18"
		sSQL = "SELECT * FROM sanctions.dbo.users"
		rs.open sSQL, sConnectionToMemberTable, 3, 1

	CASE "19"
		sSQL = "SELECT * FROM "&SanctionTableName
		rs.open sSQL, sConnectionToSanctionTable, 3, 1

	CASE"20"


END SELECT



'IF NOT rs.eof THEN
	response.write(rs.fields.count)

	%><TABLE width=100%><%
	FOR i = 0 TO rs.fields.count-1
		response.write(i &"  =  "& rs.fields(i).name&" - " &rs.fields(i).type)
'		IF rs.fields(i).type = 202 THEN
'			response.write (" - " &rs.fields(i).length)
'		END IF

		%><tr></tr><%
	NEXT      

	%>
	</TABLE><% 

'ELSE
'	response.write("No fields Found")

'END IF

rs.close
Set rs = nothing

END SUB 




' ----------------------
   SUB MakeNewArray
' ----------------------

CurrentYear = YEAR(DATE) + 1
YearString = 2006
FOR i = 2006 to CurrentYear
	YearString = YearString + ", "&i
	
NEXT

YearArray = Split(YearString,",")

FOR kvar = 0 TO UBOUND(EventArray)

	response.write(kvar)
	%><br><%
NEXT


END SUB







%>







