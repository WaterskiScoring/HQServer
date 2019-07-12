<!--#include virtual="/rankings/settingsHQ.asp"--><%

Phase4




SUB Phase4

Set SQLConnect = CreateObject("ADODB.Connection")
SQLConnect.Open Application("HQSQLConn")
Set RS = SQLConnect.Execute("SELECT [Sub Members].PrimaryPersonID, [Sub Members].SubMemberID FROM [Sub Members] WHERE ([Sub Members].SubMemberPersonID='74054')")

DO WHILE NOT rs.eof

	response.write(rs("PrimaryPersonID")&" - "&rs("SubMemberID"))%><br><%
	rs.movenext

LOOP


END SUB





SUB Phase3


sMemberID="900043479"

C=1
DO WHILE C<10
	C=C+1

	IF LEFT(RIGHT(sMemberID,10-C),1) <>"0" THEN
		sPeopleID=RIGHT(sMemberID,10-C)
		EXIT DO
	END IF

LOOP

response.write("sPeopleID = "&sPeopleID)

END SUB


SUB Phase1

	Set SQLConnect = CreateObject("ADODB.Connection")
	SQLConnect.Open Application("HQSQLConn")
	Set RS = SQLConnect.Execute("SELECT [Person ID] AS MemberID, Email, password FROM tblPeople WHERE [Date Updated]>01/01/2007 AND email <> '' AND password <> ''")

	opencon
	dim c
	rs.movefirst
	c=0	
	DO WHILE NOT rs.eof
		c=c+1
		IF rs("MemberID")<>"" AND LEFT(rs("Email"),6)<>"REMOVE" THEN
			sSQL = "INSERT INTO usawsrank.Reg_PW"
			sSQL = sSQL + " (MemberID, Email, password)"
			sSQL = sSQL + " VALUES ('"&rs("MemberID")&"', '"&rs("Email")&"', '"&SQLClean(rs("password"))&"')"
		'IF c=657 THEN 
		'	response.write(sSQL)
		'	response.end
		'END IF
	
			con.execute(sSQL)
		END IF
		rs.movenext
	LOOP 


	closecon

END SUB



SUB Phase2



	SET rs=Server.CreateObject("ADODB.recordset")
	sSQL = "SELECT MT.PersonIDWithCheckDigit AS MyMemberID, RPW.MemberID, RPW.Email, MT.Email AS NewEmail, MT.FirstName,MT.LastName, RPW.Password FROM usawsrank.RegPasswords AS RPW"
	sSQL = sSQL + " JOIN usawaterski.dbo.Members AS MT ON MT.PersonIDWithCheckDigit=RPW.MemberID"
	sSQL = sSQL + " JOIN usawsrank.Reg_PW AS PWHQ ON PWHQ.MemberID=MT.PersonID"

	rs.open sSQL, sConnectionToTRATable, 3, 1	

	opencon
	dim c
	rs.movefirst
	c=0	
	DO WHILE NOT rs.eof
		c=c+1
		sSQL = "UPDATE usawsrank.RegPasswords"
		sSQL = sSQL + " Set Email='"&rs("NewEmail")&"' WHERE MemberID='"&rs("MyMemberID")&"'"

		con.execute(sSQL)
		rs.movenext
	LOOP 




END SUB

%>





