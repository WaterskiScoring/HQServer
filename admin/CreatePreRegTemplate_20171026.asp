<% 

If not Session("aauth") then response.redirect "Login.asp"

Server.ScriptTimeout = 300

Dim curTraceMsg, sTourID

' Validate TourID value for scores to be Exported.
sTourID = Request.QueryString("TourID")
IF len(sTourID) <=0 THEN
    sTourID = Session("TourID")
END IF	

IF len(sTourID) <=0 THEN
    ErrMsg = ErrMsg & "<br />Tournament ID not provided.  Setting default"
    sTourID = "18S041R"
END IF	
TraceMsg = TraceMsg & "<br />TourId=" & sTourID


Function RemoveInvalidChars(strInput)
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
	RemoveInvalidChars = workingstring
	
End Function


' The following lines of HTML display the "opening please wait" banner.
%>
    
<html>
    <head>
        <title>USA Water Ski Registration Template</title>
        <SCRIPT LANGUAGE="JavaScript">
        // First we detect the browser type
        if(document.getElementById) { // IE 5 and up, NS 6 and up
    	    var upLevel = true;
    	    }
        else if(document.layers) { // Netscape 4
    	    var ns4 = true;
    	    }
        else if(document.all) { // IE 4
    	    var ie4 = true;
    	    }
    
        function showObject(obj) {
        if (ns4) {
    	    obj.visibility = "show";
    	    }
        else if (ie4 || upLevel) {
    	    obj.style.visibility = "visible";
    	    }
        }
    
        function hideObject(obj) {
        if (ns4) {
    	    obj.visibility = "hide";
    	    }
        if (ie4 || upLevel) {
    	    obj.style.visibility = "hidden";
    	    }
        }
    
        </SCRIPT>

    </head>

    <body>

        <DIV ID="splashScreen" STYLE="position:absolute;z-index:5;top:30%;left:35%;">
            <TABLE BGCOLOR="#000000" BORDER=1 BORDERCOLOR="#000000"	CELLPADDING=0 CELLSPACING=0 HEIGHT=150 WIDTH=300>
                <TR>
                    <TD WIDTH="100%" HEIGHT="100%" BGCOLOR="#CCCCCC" ALIGN="CENTER" VALIGN="MIDDLE">
                        <BR>
                        <FONT FACE="Helvetica,Verdana,Arial" SIZE=2 COLOR="#000066">
                        <B>Preparing your Registration Template.<br><br>
                        This may take a minute or so ...<br><br><br>  
                        </B></FONT>
                        <IMG SRC="includes/wait.gif" BORDER=1 WIDTH=150 HEIGHT=15><BR><BR>
                    </TD>
                </TR>
            </TABLE>
        </DIV>

<% 
' Once the above "please wait" banner is written to HTML, we flush the response
' buffer to make the page appear to the users browser.  That sits on their display
' while the rest of the template preparation script processing takes place.
    
response.flush

'	-----------------------------------------------------------------------
'	Start by sucking Membership Pricing Info from HQ Table into local Array
'	-----------------------------------------------------------------------
TraceMsg = TraceMsg & "<br />Trace message start"
Dim MT, MemPrice(200), MemUpgrd(200)
FOR MT = 1 to 200: MemPrice(MT) = 0: MemUpgrd(MT) = 0: NEXT

'Open connection to HQ Database
Set SQLConnect = CreateObject("ADODB.Connection")
SQLConnect.Open Application("HQSQLConn")

curSqlStmt = "SELECT * FROM [Membership Types with pricing]" 
curSqlStmt = curSqlStmt & " WHERE EffectiveFrom <= CONVERT(DATETIME, '" & session("tournamentdate") & " 00:00:00', 102)"
curSqlStmt = curSqlStmt & " AND EffectiveTo >= CONVERT(DATETIME, '" & session("tournamentdate") & " 00:00:00', 102)"
Set rsMemType = SQLConnect.Execute(curSqlStmt)
DO UNTIL rsMemType.EOF
	MT = rsMemType("Membership Type Code")
	MemPrice(MT) = rsMemType("MemberShipTypeRates")
	MemUpgrd(MT) = rsMemType("CostToUpgrade")

    curTraceMsg = curTraceMsg & "<br />Membership Type=" & rsMemType("Membership Type Code") & ", ShipTypeRates=" & rsMemType("MemberShipTypeRates") & ", CostToUpgrade="  & rsMemType("CostToUpgrade")

	rsMemType.MoveNext
LOOP

rsMemType.Close
Set rsMemType = Nothing
SQLConnect.Close

%>

        <DIV ID="debugMsg">
            <br />This is a test and only a test
            <br /><%=curTraceMsg %>

        </DIV>
<%
curTraceMsg = ""

'Open connection to WaterSki Database
Set SQLConnect = Server.CreateObject("ADODB.Connection")
SQLConnect.Open Application("WaterSkiConn")

'Open connection to Sanction Database
Dim rsSanction
Set rsSanction = Server.CreateObject("ADODB.RecordSet")
rsSanction.ActiveConnection = SQLConnect

'Setup to reference blank registration template file
Dim fileRegTempBlank
Set fileRegTempBlank = Server.CreateObject("Scripting.FileSystemObject")
Dim pathExcelFiles
pathExcelFiles = Server.MapPath("Excel/")
curTraceMsg = curTraceMsg & "<br />pathExcelFiles=" & pathExcelFiles

'Something
Dim QfyNum, DateRaw, DateFmt, I1, I2, RowNo
DateRaw = Date(): I1 = instr(DateRaw,"/"): I2 = instr(I1+1,DateRaw,"/")
DateFmt = Mid(DateRaw,I2+1): ' Start with Year value
IF I1=2 THEN DateFmt = DateFmt + "-0" + Left(DateRaw,1): ELSE DateFmt = DateFmt + "-" + Left(DateRaw,2)
IF I2-I1=2 THEN DateFmt = DateFmt + "-0" + Mid(DateRaw,I1+1,1): ELSE DateFmt = DateFmt + "-" + Mid(DateRaw,I1+1,2)

'Get TStatus and TSanction from TSchedul table
Dim strTStatus, strTSanction
curSqlStmt = "Select Top 1 TSanction, TStatus from Sanctions.dbo.TSchedul where TournAppID = '" & left(sTourID,6) & "'"
rsSanction.Open curSqlStmt
If rsSanction.EOF THEN
	strTStatus = -1: strTSanction = sTourID
ELSE 
	strTStatus = rsSanction("TStatus"): strTSanction = rsSanction("TSanction")
	IF left(strTSanction,6) <> left(sTourID,6) THEN
		strTSanction = sTourID
	END IF
END IF

curTraceMsg = curTraceMsg & "<br />sTourID=" & sTourID & ", strTStatus=" & strTStatus & ", strTSanction=" & strTSanction

rsSanction.Close

' Then we establish if there are qualifications for this TourID
curSqlStmt = "Select count(*) as QfyNum From Cobra00025.USAWSRank.RegisterQualify_TEST"
curSqlStmt = curSqlStmt & " Where left(TourID,6) = '" & left(sTourID,6) & "';"
rsSanction.Open curSqlStmt
QfyNum = rsSanction("QfyNum")
rsSanction.Close

curTraceMsg = curTraceMsg & "<br />QfyNum=" & QfyNum
%>
        <DIV ID="debugMsg">
            <br /><%=curTraceMsg %>

        </DIV>
<%
curTraceMsg = ""


%>
    </body>

</html>