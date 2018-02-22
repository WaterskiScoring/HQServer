<!--#include virtual="/admin/MemberRegFunctions.asp"-->
<% 

If not Session("aauth") then response.redirect "Login.asp"

Server.ScriptTimeout = 30

' The following lines of HTML display the "opening please wait" banner.
''''http://usawaterski.org/admin/CreateNCWSATemplate.asp

Dim curTraceMsg, sTourID, sTourDate, sStateSQL, sTourName, sUserName, AllowAccess
Dim curSqlStmt, curSanctionId, curMemberId, curMemberFirstName, curMemberLastName

'	-----------------------------------------------------------------------
' Validate TourID value for scores to be Exported.
'	-----------------------------------------------------------------------
sTourDate = ""
sTourName = ""
AllowAccess = false

sUserName = session("UserName")
sTourID = Session("TournamentID")

IF len(sTourID) > 0 THEN
    sTourID = Session("TournamentID")
    sTourDate = session("tournamentdate")
    sTourName = session("TournamentName")
ELSE
    sTourID = Request.QueryString("TourID")
END IF

curSanctionId = left(sTourID, 6)

'	-----------------------------------------------------------------------
'	Format current date for using in file name
'	-----------------------------------------------------------------------
Dim DateRaw, DateFmt, DateFmtForFile, I1, I2
DateRaw = Date(): I1 = instr(DateRaw,"/"): I2 = instr(I1+1,DateRaw,"/")
DateFmt = Mid(DateRaw,I2+1): ' Start with Year value
IF I1=2 THEN DateFmt = DateFmt + "-0" + Left(DateRaw,1): ELSE DateFmt = DateFmt + "-" + Left(DateRaw,2)
IF I2-I1=2 THEN DateFmt = DateFmt + "-0" + Mid(DateRaw,I1+1,1): ELSE DateFmt = DateFmt + "-" + Mid(DateRaw,I1+1,2)
DateFmtForFile = Mid(DateFmt, 1, 4) + Mid(DateFmt, 6, 2) + Mid(DateFmt, 9, 2)
%>
    
<html>
    <head><title>USA Water Ski NCWSA Registration Template</title>
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
'Open connection to Sanction Database
'Get tournament attributes from TSchedul table
'	-----------------------------------------------------------------------
Set WaterskiConnect = Server.CreateObject("ADODB.Connection")
WaterskiConnect.Open Application("WaterSkiConn")
Dim rsWaterski
Set rsWaterski = Server.CreateObject("ADODB.RecordSet")
rsWaterski.ActiveConnection = WaterskiConnect

Dim strTStatus, strTSanction, strTourName, strTourDate
curSqlStmt = "Select Distinct TSanction, TStatus, TournAppID, TDateE, TName, TCity, TState from " & SanctionTableName & " where TournAppID = '" & curSanctionId & "'"
rsWaterski.Open curSqlStmt
If rsWaterski.EOF THEN
	response.write "Invalid sanction number (" & curSanctionId & "), unable to complete request"
	response.status = "401 Unauthorized"
	response.flush
	response.end
ELSE
	strTStatus = rsWaterski("TStatus")
    strTSanction = rsWaterski("TSanction")
    strTourDate = rsWaterski("TDateE")
    sTourDate = strTourDate
    strTourName = rsWaterski("TName")
    sTourName = strTourName
    AllowAccess = true
END IF

''''response.write "<br /><br />sTourID=" & sTourID & ", strTStatus=" & strTStatus & ", strTSanction=" & strTSanction & ", strTourDate=" & strTourDate & ", sTourDate=" & sTourDate

rsWaterski.Close
Set rsWaterski = Nothing

'	-----------------------------------------------------------------------
'Now open a connection to the new XLS file
'Setup to reference blank registration template file
'	-----------------------------------------------------------------------
Dim path
path = Server.MapPath("Excel/")

Dim fileRegXls
Set fileRegXls = Server.CreateObject("Scripting.FileSystemObject")
Dim pathExcelFiles
pathExcelFiles = Server.MapPath("Excel/")
dim copyFileSour, copyFileDest
''''response.write "<br /><br />pathExcelFiles=" & pathExcelFiles

copyFileSour = pathExcelFiles & "/Templates/NCWSATemplateBlank.xls"
copyFileDest = pathExcelFiles & "/template.xls"
''''response.write "<br />copyFileSour=" & copyFileSour & "<br />copyFileDest=" & copyFileDest

fileRegXls.CopyFile copyFileSour, copyFileDest , True

'	-----------------------------------------------------------------------
'Now open a connection to the new XLS file
'	-----------------------------------------------------------------------
Set objExcelConn = Server.CreateObject("ADODB.Connection")
objExcelConn.Provider = "Microsoft.ACE.OLEDB.12.0"
objExcelConn.ConnectionString = "Data Source=" & copyFileDest & ";Extended Properties=""Excel 8.0;"""
    On Error Resume Next
objExcelConn.Open
    If Err.Number <> 0 Then
        %>
            <DIV ID="debugMsg">
                <br />Error creating registration template file
                <br />Err.Number=<%=Err.Number %>
                <br />Err.Description=<%=Err.Description %>
                <br />
            </DIV>
        <%
        On Error Goto 0 ' But don't let other errors hide!
    End If
''''response.write "<br />Open Excel file=" & copyFileDest

Set objExcelSingleFields = Server.CreateObject("ADODB.Recordset")
objExcelSingleFields.ActiveConnection = objExcelConn 
objExcelSingleFields.CursorType = 3                    'Static cursor.
objExcelSingleFields.LockType = 2                      'Pessimistic Lock.

objExcelSingleFields.Source = "Select * from RegistTourName"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = sTourName
objExcelSingleFields.update
objExcelSingleFields.close
		
objExcelSingleFields.Source = "Select * from RegistTourID"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = strTSanction
objExcelSingleFields.update
objExcelSingleFields.close
		
objExcelSingleFields.Source = "Select * from AllOthrTourName"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = sTourName
objExcelSingleFields.update
objExcelSingleFields.close
		
objExcelSingleFields.Source = "Select * from AllOthrTourID"
objExcelSingleFields.Open
objExcelSingleFields.Fields(0).Value = strTSanction
objExcelSingleFields.update
objExcelSingleFields.close
		
Set objExcelRegist = Server.CreateObject("ADODB.Recordset")
objExcelRegist.ActiveConnection = objExcelConn 
objExcelRegist.CursorType = 3                    'Static cursor.
objExcelRegist.LockType = 2                      'Pessimistic Lock.
objExcelRegist.Source = "Select * from RegistRange"
objExcelRegist.Open

Set objExcelAllOthr = Server.CreateObject("ADODB.Recordset")
objExcelAllOthr.ActiveConnection = objExcelConn 
objExcelAllOthr.CursorType = 3                    'Static cursor.
objExcelAllOthr.LockType = 2                      'Pessimistic Lock.
objExcelAllOthr.Source = "Select * from AllOthrRange"
objExcelAllOthr.Open

'	-----------------------------------------------------------------------
' Refresh the list of chief and appointed officials for a tournament
' The data is stored in a temporary work table for use in build tournament registration entries
'	-----------------------------------------------------------------------
refreshApptOfficials(curSanctionId)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Now build a Query to Extract the Desired Members, joining in data 
''' pulled from the Rankings and Officials and Membership Type tables.
''' Note that we prefix each team ID with "E" if the team has entries,
''' or "Z" if no entries, so that all the entered teams list at the top,
''' then finally all those without any team affiliation last with Zzzz.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
curSqlStmt = ""
curSqlStmt = buildQueryMemberRegNcwsaEntries(curSanctionId, sTourDate)
''''response.write "<br /><br />" & curSqlStmt & "<br /><br />"

Set WaterskiConnect = Server.CreateObject("ADODB.Connection")
WaterskiConnect.Open Application("WaterSkiConn")
Set rsWaterski = Server.CreateObject("ADODB.RecordSet")
rsWaterski.ActiveConnection = WaterskiConnect
    On Error Resume Next
rsWaterski.Open curSqlStmt
    If Err.Number <> 0 Then
        %>
            <DIV ID="debugMsg">
                <br />Error creating registration template file
                <br />Err.Number=<%=Err.Number %>
                <br />Err.Description=<%=Err.Description %>
                <br />SQL Statement:<br /><%=curSqlStmt %>
            </DIV>
        <%
        On Error Goto 0 ' But don't let other errors hide!
    End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' Now we loop through the Extract of Members, merging the
''' Team Headers into All sections, by Team ID.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'	Now here's the scoop on the primary variables coming in and how we split things up
' Sorter is now four positions, with "E" at the front if member of Entered Team, otherwise "Z"
' Rest of Sorter indicates team, if any.  "zzz" for non-team skiers.  " OF" are apptd Officials

' as to section splits.  Sorter Prefix is "Z" then all go to OTHERS section
' Where sorter prefix is "E", then if TeamStat is "I" then goes to OTHERS,
' For xxxEnt stat "DD", we need to compute age div code.

Dim curTeamId, curTeamName, curMemType, curSortValue
Dim prevTeamId, prevTeamName, prevSortValue
Dim NextRotSlmMen, NextRotTrkMen, NextRotJmpMen
Dim NextRotSlmWom, NextRotTrkWom, NextRotJmpWom

curSortValue = ""
prevTeamId = "E0FF"
prevTeamName = "Appointed Officials"
prevSortValue = "E0FF"
TeamSlm = 0
TeamTrk = 0
TeamJmp = 0
TeamTot = 0
GrandSlm = 0
GrandTrk = 0
GrandJmp = 0
GrandTot = 0

''''response.write "<br /><br />Select members "

DO until rsWaterski.EOF
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Determine the membership status 
    ' e.g. "OK to Ski" OR "Reason not ready to ski" and renew/upgrade amount strings
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''curTraceMsg = curTraceMsg & "<br /><br />FirstName=" & rsWaterski("FirstName") & ", LastName=" & rsWaterski("LastName") & ", Team=" & rsWaterski("Team") & ", OffCode=" & rsWaterski("ApptdOfficial") & ", Sorter=" + rsWaterski("Sorter") & ", TeamStat=" & rsWaterski("TeamStat")
    ''''curTraceMsg = curTraceMsg & ", JudgeSlalom=" & rsWaterski("JudgeSlalom") & ", DriverSlalom=" & rsWaterski("DriverSlalom") & ", ScorerSlalom=" & rsWaterski("ScorerSlalom")

	curMemType = rsWaterski("MemType")
	IF curMemType < 1 OR curMemType > 200 THEN curMemType = 1

	IF rsWaterski("EffTo") < cdate(sTourDate) THEN 
		IF rsWaterski("CanSki") = False THEN
			OKtoSki = "Needs Renew/Upgrade" 
            UpgrdAmt = FormatNumber(rsWaterski("MembershipRate") + rsWaterski("CostToUpgrade"), 2)
		ELSE
			OKtoSki = "Needs Renew" 
            UpgrdAmt = FormatNumber(rsWaterski("MembershipRate"), 2)
		END IF
	ELSE 
		IF rsWaterski("CanSki") = False THEN
			OKtoSki = "Needs Upgrade" 
            UpgrdAmt = FormatNumber(rsWaterski("CostToUpgrade"), 2)
		ELSEIF rsWaterski("Waiver") = 0 THEN
			OKtoSki = "Needs Annual Waiver" 
			UpgrdAmt = ""
		ELSEIF rsWaterski("EventWaiver") <> "X" THEN
			OKtoSki = "Needs Event Waiver" 
			UpgrdAmt = ""
		ELSE
			OKtoSki = "" 
			UpgrdAmt = ""
		END IF				
	END IF
	
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Determine if current record is for a new team
    ' When a new team is detected then write totals for previous team
    ' Write team header for new team
	' Put into registration section if Prefix is "E", otherwise only into All Other
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	curTeamId = rsWaterski("Team")
	curTeamName = rsWaterski("TeamName")
	curSortValue = rsWaterski("sorter")

    IF curSortValue <> prevSortValue THEN
		IF left(prevSortValue, 1) = "E" THEN
			objExcelRegist.addnew
			objExcelRegist.Fields(0).Value = "Team Totals"
			objExcelRegist.Fields(1).Value = prevTeamName
			objExcelRegist.Fields(3).Value = prevTeamId
			objExcelRegist.Fields(8).Value = TeamSlm
			objExcelRegist.Fields(9).Value = TeamTrk
			objExcelRegist.Fields(10).Value = TeamJmp
			objExcelRegist.Fields(11).Value = " <Rides  Skiers>"
			objExcelRegist.Fields(15).Value = TeamTot
			objExcelRegist.Update	
			
            objExcelRegist.addnew
			objExcelRegist.Fields(0).Value = " "
			objExcelRegist.Update	
	
			GrandSlm = GrandSlm + TeamSlm
			GrandTrk = GrandTrk + TeamTrk
			GrandJmp = GrandJmp + TeamJmp
			GrandTot = GrandTot + TeamTot
			TeamSlm = 0
			TeamTrk = 0
			TeamJmp = 0
			TeamTot = 0			

			IF curSortValue <> "E0FF" THEN
				objExcelRegist.addnew
				objExcelRegist.Fields(0).Value = "Team Header"
				objExcelRegist.Fields(1).Value = curTeamName
				objExcelRegist.Fields(3).Value = curTeamId
				objExcelRegist.Fields(4).Value = "CM"
				objExcelRegist.Fields(8).Value = "RD"
				objExcelRegist.Update

				objExcelRegist.addnew
				objExcelRegist.Fields(0).Value = "Team Header"
				objExcelRegist.Fields(1).Value = curTeamName
				objExcelRegist.Fields(3).Value = curTeamId
				objExcelRegist.Fields(4).Value = "CW"
				objExcelRegist.Fields(8).Value = "RD"
				objExcelRegist.Update
			END IF

        ELSE
			IF curSortValue <> "E0FF" THEN
				objExcelAllOthr.addnew
				objExcelAllOthr.Fields(0).Value = "Team Header"
				objExcelAllOthr.Fields(1).Value = curTeamName
				objExcelAllOthr.Fields(3).Value = curTeamId
				objExcelAllOthr.Fields(4).Value = "CM"
				objExcelAllOthr.Fields(8).Value = "RD"
				objExcelAllOthr.Update

				objExcelAllOthr.addnew
				objExcelAllOthr.Fields(0).Value = "Team Header"
				objExcelAllOthr.Fields(1).Value = curTeamName
				objExcelAllOthr.Fields(3).Value = curTeamId
				objExcelAllOthr.Fields(4).Value = "CW"
				objExcelAllOthr.Fields(8).Value = "RD"
				objExcelAllOthr.Update
			END IF

        END IF
    END IF

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	' Store current skier in the "Registration" section, if an active member of an entered team.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	IF left(curSortValue, 1) = "E" and (rsWaterski("TeamStat") = "A" or rsWaterski("ApptdOfficial") <> "") THEN
		NumEvts = 0
		
        objExcelRegist.addnew
		objExcelRegist.Fields(0).Value = rsWaterski("MemberID")
		objExcelRegist.Fields(1).Value = rsWaterski("LastName")
		objExcelRegist.Fields(2).Value = rsWaterski("FirstName")

		IF curSortValue = "E0FF" THEN
			objExcelRegist.Fields(3).Value = "OFF"
		ELSE
			objExcelRegist.Fields(3).Value = trim(rsWaterski("Team"))
		END IF

		IF rsWaterski("EventSlalom") = "DD" or rsWaterski("EventTrick") = "DD" or rsWaterski("EventJump") = "DD" or Instr(Ucase(sTourName),"ALUMNI") > 0 THEN
			objExcelRegist.Fields(4).Value = rsWaterski("AgeDiv")
		ELSE
			objExcelRegist.Fields(4).Value = rsWaterski("Div")
		END IF

		objExcelRegist.Fields(5).Value = rsWaterski("Age")
		objExcelRegist.Fields(6).Value = rsWaterski("City")
		objExcelRegist.Fields(7).Value = rsWaterski("State")
		objExcelRegist.Fields(8).Value = rsWaterski("EventSlalom")
		objExcelRegist.Fields(9).Value = rsWaterski("EventTrick")
		objExcelRegist.Fields(10).Value = rsWaterski("EventJump")

    	objExcelRegist.Fields(11).Value = rsWaterski("ApptdOfficial")

    	''''objExcelRegist.Fields(12).Value = "xxx"

		objExcelRegist.Fields(13).Value = rsWaterski("TrickBoat")
		objExcelRegist.Fields(14).Value = rsWaterski("JumpHeight")

		objExcelRegist.Fields(16).Value = rsWaterski("SptsDiv")
		objExcelRegist.Fields(17).Value = OKtoSki
		objExcelRegist.Fields(18).Value = UpgrdAmt

		IF rsWaterski("EventSlalom") > "  " THEN 
			NumEvts = NumEvts + 1
			TeamSlm = TeamSlm + 1
		END IF

		IF rsWaterski("EventTrick") > "  " THEN 
			NumEvts = NumEvts + 1
			TeamTrk = TeamTrk + 1
		END IF

		IF rsWaterski("EventJump") > "  " THEN 
			NumEvts = NumEvts + 1
			TeamJmp = TeamJmp + 1
		END IF

		
		IF NumEvts > 0 THEN 
			objExcelRegist.Fields(19).Value = NumEvts
			TeamTot = TeamTot + 1
		END IF

		objExcelRegist.Fields(32).Value = rsWaterski("JudgeSlalom")
		objExcelRegist.Fields(33).Value = rsWaterski("JudgeTrick")
		objExcelRegist.Fields(34).Value = rsWaterski("JudgeJump")
		objExcelRegist.Fields(35).Value = rsWaterski("DriverSlalom")
		objExcelRegist.Fields(36).Value = rsWaterski("DriverTrick")
		objExcelRegist.Fields(37).Value = rsWaterski("DriverJump")
		objExcelRegist.Fields(38).Value = rsWaterski("ScorerSlalom")
		objExcelRegist.Fields(39).Value = rsWaterski("ScorerTrick")
		objExcelRegist.Fields(40).Value = rsWaterski("ScorerJump")
		objExcelRegist.Fields(41).Value = rsWaterski("Safety")
		objExcelRegist.Fields(42).Value = rsWaterski("TechController")
    	''''objExcelRegist.Fields(12).Value = "zzz"

		objExcelRegist.Update

	ELSE
		objExcelAllOthr.addnew
		objExcelAllOthr.Fields(0).Value = rsWaterski("MemID")
		objExcelAllOthr.Fields(1).Value = rsWaterski("LastName")
		objExcelAllOthr.Fields(2).Value = rsWaterski("FirstName")
		objExcelAllOthr.Fields(3).Value = trim(rsWaterski("Team"))
		objExcelAllOthr.Fields(4).Value = rsWaterski("Div")
		objExcelAllOthr.Fields(5).Value = rsWaterski("Age")
		objExcelAllOthr.Fields(6).Value = rsWaterski("City")
		objExcelAllOthr.Fields(7).Value = rsWaterski("State")
		objExcelAllOthr.Fields(11).Value = rsWaterski("OffRat")
		objExcelAllOthr.Fields(16).Value = rsWaterski("SptsDiv")
		objExcelAllOthr.Fields(17).Value = OKtoSki
		objExcelAllOthr.Fields(18).Value = UpgrdAmt

		objExcelAllOthr.Fields(32).Value = rsWaterski("JudgeSlalom")
		objExcelAllOthr.Fields(33).Value = rsWaterski("JudgeTrick")
		objExcelAllOthr.Fields(34).Value = rsWaterski("JudgeJump")
		objExcelAllOthr.Fields(35).Value = rsWaterski("DriverSlalom")
		objExcelAllOthr.Fields(36).Value = rsWaterski("DriverTrick")
		objExcelAllOthr.Fields(37).Value = rsWaterski("DriverJump")
		objExcelAllOthr.Fields(38).Value = rsWaterski("ScorerSlalom")
		objExcelAllOthr.Fields(39).Value = rsWaterski("ScorerTrick")
		objExcelAllOthr.Fields(40).Value = rsWaterski("ScorerJump")
		objExcelAllOthr.Fields(41).Value = rsWaterski("Safety")
		objExcelAllOthr.Fields(42).Value = rsWaterski("TechController")

		objExcelAllOthr.Update	
            		
	END IF

	rsWaterski.MoveNext

    prevSortValue = curSortValue
    IF curSortValue <> "E0FF" THEN
        prevTeamId = curTeamId
        prevTeamName = curTeamName
	END IF


LOOP


'"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

'	Finally tack on Overall Registration Total Rides and Skiers

objExcelRegist.addnew
objExcelRegist.Fields(0).Value = " "
objExcelRegist.Update	

objExcelRegist.addnew
objExcelRegist.Fields(0).Value = "Grand Tots"
objExcelRegist.Fields(1).Value = "Across All Teams"
objExcelRegist.Fields(3).Value = "Tot"
objExcelRegist.Fields(8).Value = GrandSlm
objExcelRegist.Fields(9).Value = GrandTrk
objExcelRegist.Fields(10).Value = GrandJmp
objExcelRegist.Fields(11).Value = " <Rides  Skiers>"

' objExcelRegist.Fields(13).Value = "Skiers>"
objExcelRegist.Fields(15).Value = GrandTot
objExcelRegist.Update	

objExcelRegist.close
set objExcelRegist = nothing
objExcelAllOthr.close
set objExcelAllOthr = nothing
objExcelConn.close
set objExcelConn = nothing
'
rsWaterski.Close
Set rsWaterski = Nothing
WaterskiConnect.Close

'	-----------------------------------------------------------------------
'Now copy the file from Template to a file with the tournamentid
'	-----------------------------------------------------------------------
Dim regTemplateFilename
regTemplateFilename = "Entries-" & DateFmtForFile

'	-----------------------------------------------------------------------
'Add the Tournament Name to the start of the file name
'	-----------------------------------------------------------------------
if len(sTourName) > 0 then
	regTemplateFilename = RemoveInvalidChars(sTourName) & "-" & regTemplateFilename
end if

'	-----------------------------------------------------------------------
'Append the username
'	-----------------------------------------------------------------------
if len(strTSanction) > 0 then
	regTemplateFilename = regTemplateFilename & "-" & strTSanction & ".xls"
else
	regTemplateFilename = regTemplateFilename & ".xls"
end if

''''response.write "<br /><br />copyFileDest=" + pathExcelFiles & "/" & regTemplateFilename

fileRegXls.CopyFile copyFileDest, pathExcelFiles & "/" & regTemplateFilename , True

'	-----------------------------------------------------------------------
' Clean up old files
'	-----------------------------------------------------------------------
Set dataFolder = objFSO.GetFolder(pathExcelFiles)
Set folderFileList = dataFolder.Files
For Each curFile in folderFileList
	Set myfile = objFSO.GetFile(pathExcelFiles & "/" & curFile.name)
	if datediff("d",myfile.DateCreated,date()) > 2 and left(myfile.name,8) <> "Template" then
		myfile.delete
	end if
Next

Set dataFolder = nothing
Set folderFileList = nothing
Set objFSO = Nothing
Set fileRegXls = Nothing

Response.Flush
      
' This final bit of HTML is written after processing is successfully completed
' to tell the user how to download their template, and where to go from here.
      
%>
    
    <SCRIPT LANGUAGE="JavaScript">
    if(upLevel) {
      var splash = document.getElementById("splashScreen");
    }
    else if(ns4) {
      var splash = document.splashScreen;
    }
    else if(ie4) {
      var splash = document.all.splashScreen;
    }
      
    hideObject(splash);
    </SCRIPT>  


<html>

<head>
<title>Create Pre-Registration Export v1.5</title>

<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" background = "/images/TopBackground.jpg" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" >

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#42639F">
      <p>&nbsp;</p>
      <p align="center"><font face="Verdana" size="6" color="#FFFFFF">
      	USA Water Ski NCWSA Registration Template</font></p>
      <p align="center"><font face="Verdana" size="4" color="#FFFFFF">
      	Registration Support for -- <%=sTourName%></font></p>
      <p>&nbsp;</p>
    </td>
  </tr>
</table>  
  
<table border="0" cellspacing="0" cellpadding="0">  
  <tr> 
    <td width="185" valign="top" bgcolor="#42639F">

	<% If Session("aauth") then %>
	<font face="Verdana" size="2" COLOR="#FFFFFF"><br>&nbsp;Currently Logged in as: </font><br>
	<font face="Verdana" size="2" COLOR="#FFFFFF">&nbsp;<%=Session("UserName")%>&nbsp;&nbsp;
		<%=sTourDate%></font><br>
	<br>
	<% Else %>
	<font face="Verdana" size="2" COLOR="#FFFFFF"><br>&nbsp;Not currently logged in.</font>
	<% End If %>
	
			<font face="Verdana" size="2"> 
         <br>&nbsp;<a href="logout.asp"><font face="arial" COLOR="#FFFFFF">Log Out</font></a>&nbsp;<br>
			</font>
			<br>
	        &nbsp;<a href="/admin/index.asp"><font face="arial" size="2" COLOR="#FFFFFF">Back to Admin Index</font></a><br>&nbsp;<br>
	        &nbsp;<a href="http://www.usawaterski.org"><font face="arial" size="2" COLOR="#FFFFFF">USA Water Ski Home</font></a><br>&nbsp;<br>
			<br>
            <font face="Verdana" size="1">&nbsp;<font COLOR="#FFFFFF">Powered by</font> <a href="http://www.epolk.com"><font COLOR="#FFFFFF">ePolk.com</font></a><br></font>

  </td>

	<td>

  <table>
      <tr> 
         <td width="14">&nbsp;</td>
         <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><br>Your NCWSA
         Registration Export Excel Workbook is now complete and ready to download.&nbsp;</font>
         <br>&nbsp;<br>

         <font color="#FF0000" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>
         !! New for 2011 -- Online Team Entry and Rotation Plan details now included !!
         </strong></font><br>
         <font size="2" face="Verdana, Arial, Helvetica, sans-serif">Details of Team Entry
         and Rotation Plans that have been prepared and submitted by the respective team
         captains through the new online team entry system, are now incorporated into this 
         Excel workbook.&nbsp; See the revised instructions section of your Excel workbook 
         for details.&nbsp; <font color="#FF0000"><strong>New for Fall 2011 -- A Registrar 
         Recap is now included.</strong></font>&nbsp; This new section in the Excel workbook 
         makes it easier for Registrars to see each team's overall entry status, to see which 
         entered skiers still need to execute event waivers locally, and to assess each team's 
         total entry fees.</font>
         <br>&nbsp;<br>

         <a href="excel/<% response.write regTemplateFilename %>"><font face="Arial" size="2"><b>RIGHT 
         Click Here</b></font></a>&nbsp; <font size="2" face="Verdana, Arial, Helvetica, 
         sans-serif">to download your NCWSA Registration Template workbook, then select the 
         "Save As" option from that menu, and then choose a suitable location to store the 
         download file in your PC.&nbsp; After your Registration Template download has 
         completed, then open the Excel file from that location on your PC.&nbsp; It will 
         open automatically to an Instructions Tab section.&nbsp; Please review the material 
         in that section for the latest information on contents and usage.</font>
         <br>&nbsp;<br>

         <% IF AllowAccess THEN %>

         <font size="2" face="Verdana, Arial, Helvetica, sans-serif">If you are now doing 
         your <b><i>final and official</i></b> download of entries for this tournament, then 
         <b><i>after</i></b> downloading your Excel workbook (see the <b>RIGHT Click Here</b> 
         link in paragraph above), then click the <b>Close Registration</b> button that you
         see below.&nbsp; That will block any further modifications to existing Team Entry and 
         Rotation Plans, and refer team captains to the Tournament Registrar at the tournament 
         site for any last-minute changes.</font>
         <br>&nbsp;<br>
         
         <% ELSE %>

         <font size="2" face="Verdana, Arial, Helvetica, sans-serif">Online Entry to this 
         tournament is currently set to <b>Closed</b>.&nbsp; If you have not yet done your 
         <b><i>final and official</i></b> download of entries for this tournament, then 
         you may want to re-open Online Entry status, by clicking the <b>Re-Open Registration</b> 
         button below.</font>
         <br>&nbsp;<br>
         
         <% END IF %>

         </td>
      </tr>

 	</table>

	<TABLE ALIGN="CENTER" WIDTH=80%>
		
		<TR>

    <% IF AllowAccess THEN %>

		    <TD width=35% align=center>
			<form action="NCWSAChgRegStat.asp?TourID=<%=curSanctionId%>&Status=Close" method="post">
			<input type="submit" style="width:12em" value="Close Registration"
			title="Close Online Registration -- No further Changes by Captains allowed"></form>
 		   	</TD>

    <% ELSE %>

		    <TD width=35% align=center>
			<form action="NCWSAChgRegStat.asp?TourID=<%=curSanctionId%>&Status=Open" method="post">
			<input type="submit" style="width:12em" value="Re-Open Registration"
			title="Close Online Registration -- No further Changes by Captains allowed"></form>
 		   	</TD>

    <% END IF %>

	    <TD width=30% align=center>
		<form action="LookupMembers.asp?FormStatus=newsearch" method="post">
		<input type="submit" style="width:10em" value="Lookup Members"></form>
    	</TD>

	    <td width=25% align=center>     				
		<form action="Index.asp" method="post">
    <input type="submit" style="width:7em" value="Quit"></form>
 	    </td>
  	    
 	  </TR>

 	</TABLE>

  	  </td>
	  </tr>
</table>

</body>
</html>
