<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>

<% 

'	-----------------------------------------------------------------------
'	Utility Process to Display a Table from the HQ Membership System Server
'	-----------------------------------------------------------------------


Set SQLConnect = CreateObject("ADODB.Connection")
SQLConnect.Open Application("HQSQLConn")

Dim sSQL, TotalRows, RowCount

'	sSQL = "SELECT * FROM [Membership History] Where EffectiveFrom >= '2008-01-01'" 
sSQL = "SELECT * FROM [Members] Where LastName = 'Clark' and State = 'IL' Order By FirstName" 

Set RS = SQLConnect.Execute(sSQL)

%><br><br><table width=10% BORDER="1" ><%

DO WHILE NOT rs.eof

	rowCount = rowCount + 1

		IF RowCount = 1 THEN
	  		%><tr><%
			FOR i = 0 TO rs.fields.count -1 
				%><TD ALIGN="Left" vAlign="top" nowrap><FONT COlOR="#000000" Size="1">&nbsp;<%
				Response.Write(trim(Rs.Fields(i).Name))%>
				</td><% 
			NEXT

	  		%></tr><%
		END IF		

	%><tr><%
	FOR i = 0 TO rs.fields.count -1 

		
			%><TD ALIGN="Left" vAlign="top" nowrap><FONT COlOR="#000000" Size="1">&nbsp;<%

			SELECT CASE i
				CASE ELSE
					Response.Write(rs.fields(i).value)
			END SELECT

			%>&nbsp;</TD></FONT><%

	
	NEXT

	
	IF rs.eof THEN
		EXIT DO
	END IF

	%></TR><%

	rs.movenext

LOOP

%>
<table>

<% 

rs.close

%>









