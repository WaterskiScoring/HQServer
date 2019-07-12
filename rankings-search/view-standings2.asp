<!--#include virtual="/settings.asp"-->

<%
Dim currentPage, rowCount, i
Dim MemoryScore, MemoryPlc, MemoryRank, RecordNum, RankValueWithTies, ThisCOA



Dim tName, tRankScore, tFmtScore, tRnkScoBkup, tMemberID, tRating, tState, tRegion, tNatPlace, tRegSki, tRegPlace, tMemberFed, tTeam

Dim DefineRowColor, tPerc1, tPerc2, tPerc3, tPerc4, tPerc5, tPerc6, tPerc7, tPerc8, tPerc9, tPerc10

Dim tBirth, tAge, tBirthday, tSkiYearEndDate, tLevel, tASC3, tRCU

Dim sMemberID, sFullName


SAdminMenuLevel=Session("AdminMenuLevel")

' --- This is a TEMPORARY fix.  Actual cut-off logic needs to be dynamic.
CutOffDate = "07/23/2008"


sRunByWhat = TRIM(Request("pvar"))
pvar = TRIM(Request("pvar"))
adminmenulevel = TRIM(Request("adminmenulevel"))

EventSelected = TRIM(Request("event"))
IF Request("MyEvent")<>"" THEN EventSelected=Request("MyEvent")
DivSelected = TRIM(Request("div"))
IF Request("MyDiv")<>"" THEN DivSelected=Request("MyDiv")



RegionSelected = TRIM(Request("region"))
StateSelected = TRIM(Request("state"))

FederationSelected = TRIM(Request("Include_International"))
RecordNum = TRIM(Request("RecordNum"))    
RemoveCTF = TRIM(Request("CTF"))
SkiYearSelected = TRIM(Request("SkiYear"))
IF TRIM(SkiYearSelected) = "" AND TRIM(Session("SkiYear"))<>"" THEN SkiYearSelected=Session("SkiYear")

' --- Maybe?
'IF TRIM(Request("SkiYear")) <> "" THEN
'	Session("SkiYear") = TRIM(Request("SkiYear"))
'END IF


IF RecordNum = "" THEN RecordNum = 1
    
IF EventSelected = "" THEN EventSelected = "S"
IF RegionSelected = "" THEN RegionSelected = "All"
IF StateSelected = "" THEN StateSelected = "All"


Dim ThisFileName
ThisFileName="view-standings2.asp"



sMemberID=TRIM(Request("sMemberID"))


' --- Determine Name of CURRENT user ---
SET rs=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT * FROM "&MemberTableName&" AS MT"
sSQL = sSQL + " WHERE MT.PersonIDWithCheckDigit='"&sMemberID&"'"
ChoosePagesSQL sSQL,currentPage, 30

IF NOT rs.EOF THEN
	sFullName=rs("FirstName")&" "&rs("LastName")
ELSE
	sFullName="None"
END IF



' --- If User pressed Find My Ranking button and MemberID was not set OR user pressed get a New Member button ---
IF (TRIM(Request("SingleRanking"))<>"" AND TRIM(sMemberID)="") OR TRIM(Request("NewMember"))<>"" THEN
	' --- This is where I would branch to get a member if not set ---

	' --- Sends user to search-member routine to selected member
	Session("sSendingPage")="/"&ThisFileName&"?SingleRanking=Find&pvar="&sRunByWhat
	Response.Redirect("/search-member.asp?rid="&rid&"&formstatus=search")

	'SingleRanking=TRIM(Request("SingleRanking"))
ELSEIF Trim(Request("SingleRanking"))<>"" THEN
	SingleRanking="Find My Ranking"
END IF










NewsPageNum = "4m"
IF Request("PVar") = "NSL" THEN
    NewsPageNum = "4m-nsl"
END IF


' --------- Defines the images and text for news box at right of screen --------
' ------------------------------------------------------------------------------

NewsHead_01 = rankhead_01
NewsHead_02 = rankhead_02

SELECT CASE EventSelected
	CASE "S"
		NewsImage_01 = rankimage_SL
		NewsImageCaption_01 = rankimagecaption_SL
		Newstitle_01 = ranktitle_SL
	CASE "T"
		NewsImage_01 = rankimage_TR
		NewsImageCaption_01 = rankimagecaption_TR
		Newstitle_01 = ranktitle_TR
	CASE "J"
		NewsImage_01 = rankimage_JU
		NewsImageCaption_01 = rankimagecaption_JU
		Newstitle_01 = ranktitle_JU
	CASE ELSE

END SELECT






WriteIndexPageHeader


' --------------------------------------------------------------------------------- 
' Creates Radio Buttons to select LIST TYPE in case NOT selected from Settings menu


IF sRunByWhat = "" THEN
    %>
    <br><br>
    <center><h2>View Rankings<br></h2>
    <br><br>
    <form action="/"&ThisFileName&"?sMemberID=<%s=MemberID%>&rid=<%=rid%>" method="post">
    <input type="radio" name="pvar" value="National">National&nbsp;<br><br>
    <input type="radio" name="pvar" value="Regional">Regional&nbsp;<br><br>
    <input type="radio" name="pvar" value="Junior">Junior&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br><br>
    <input type="radio" name="pvar" value="NCWSA">NCWSA<br><br>
    <input type="radio" name="pvar" value="NSL">NSL&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br><br>
    <input type="submit" value="Continue"><br><br><br>
    </form>
    <%

ELSE





    ' -----------------------------------------------------------------------------------------------------------
    ' ----------------   Sets Session("SkiYear") to request string from form   ------------------
    ' -----------------------------------------------------------------------------------------------------------
    ' --- NCWSA test is done first 
    'IF (TRIM(Request("SkiYear")) = "1" OR TRIM(Request("SkiYear")) = "") AND sRunByWhat="NCWSA" THEN 
	IF (SkiYearSelected = "1" OR SkiYearSelected = "") AND sRunByWhat="NCWSA" THEN 

    	OpenCon
	Set rs = Server.CreateObject("ADODB.recordset")
	sSQL = "SELECT top 1 * from " & SkiYearTableName & " WHERE DefaultYear = 1"
    	rs.open sSQL, SConnectionToTRATable, 3, 3  

	IF NOT rs.EOF THEN
		Session("SkiYear")=rs("SkiYearID")
	END IF		


    ' --- Assigns SkiYear to whatever current setting is if there is a variable set on form
    ELSEIF SkiYearSelected <> "" THEN 
	Session("SkiYear") = SkiYearSelected

    ' --- If nothing is assigned, then set it to 12 month rankings
    ELSE 	
	Session("SkiYear")="1"	
    END IF	




    	
    SET rsSelectFields = Server.CreateObject("ADODB.recordset")

    ' -----------------------------------------------------------------------------------------------------------
    ' -------- If no division is selected then populate dropdown with divisions found in Rankings Table ---------
    ' -----------------------------------------------------------------------------------------------------------
    IF DivSelected = "" THEN
      opencon
      SET rsSelectFields=Server.CreateObject("ADODB.recordset")
      sSQL = "Select top 1 div from " & RankTableName

      SELECT CASE sRunByWhat
        CASE "Junior"
          sSQL = sSQL + " WHERE LOWER(left(div,1)) = 'b' OR LOWER(left(div,1)) = 'g'"
        CASE "NCWSA"
          sSQL = sSQL + " WHERE LOWER(div) = 'cm' OR LOWER(div) = 'cw'"
        CASE ELSE
          sSQL = sSQL + " WHERE LOWER(div) <> 'cm' AND LOWER(div) <> 'cw'"  ' AND LOWER(left(div,1)) <> 'i'"  Potential Upgrade Later On per Mark Crone
      END SELECT

      sSQL = sSQL + " order by div"
      rsSelectFields.open sSQL, SConnectionToTRATable

      IF not rsSelectFields.eof THEN DivSelected = RsSelectFields.Fields(0).value  
      rsSelectFields.close
      Closecon
    END IF


    ' -------- If NCWSA then select ALL Federations, otherwise only USA --------

    IF FederationSelected = "" AND sRunByWhat="NCWSA" THEN 
		FederationSelected = "ALL"	                
    ELSEIF FederationSelected = "" THEN 
		FederationSelected = "USA"
    END IF	


    currentPage = TRIM(Request("currentPage"))
    IF currentPage = "" THEN currentPage = 1
    
    sID = TRIM(Request("id"))
    IF sID = "" THEN sID = 0
            
            
    ThisPage = Request.ServerVariables("SCRIPT_NAME")
            

    ' ------------------------------------------------------------------------------------------------           
    ' -------------------------------  BEGINS WRITING HEADERS AND RANKINGS  --------------------------
    ' ------------------------------------------------------------------------------------------------

    tempSkiYear = Session("SkiYear")
    WriteHeader

    ' -------------- EXAMINE THIS CODE TO GET BRANCHING AND RECEIPT OF PARAMETERS  -------------------	


    ' ---- Calls subroutine to display header  ----	
    WriteNewHeader sRunByWhat&" Rankings", "Range ID = " &Session("SkiYear")&""

	' Check Recalculation Underway Flag for the Ski Year selected.
	' If it's currently on, issue Come Back Later -- otherwise proceed.
	
   	OpenCon
		Set rs = Server.CreateObject("ADODB.recordset")
		sSQL = "SELECT Case when RecalcUnderway=1 THEN 'Y' ELSE 'N' END as RCUFlag FROM " & SkiYearTableName & " WHERE SkiYearID = " & tempSkiYear
   	rs.open sSQL, SConnectionToTRATable, 3, 3  
    IF rs.EOF THEN tRCU = "N" ELSE tRCU = RS("RCUFlag")
    rs.close

    IF tRCU = "Y" THEN 
		%><br><br><H2><center><font color="red">Ranking Recalculations are currently 
		  <br>underway For the Ski Year requested.<br>&nbsp;
		  <br>Please try your request again in a few minutes.
		  <br>We apologize for the temporary inconvenience.</font></H2><% 	
    ELSE
	   	Standings
    END IF

    WriteFooter

END IF

WriteIndexPageFooter





'--------------------------------------
  SUB WriteHeaders (sTitle, sSubTitle)
'--------------------------------------

' ------------------  WHY DEFINE THIS WAY?  COULD THE NEW TEXT/IMAGE DEFINITION BENEFIT BY THIS APPROACH?
' Write Headers for DB Page
%>

<TABLE BORDER="0" CELLPADDING="6" CELLSPACING="0" WIDTH="100%" BGCOLOR="#C0C0C0" BORDERCOLOR="#C0C0C0" BORDERCOLORDARK="#C0C0C0" BORDERCOLORLIGHT="#C0C0C0" >
  <TR>
	<TD align="center" vAlign=bottom noWrap background="/images/buttons/Vertical_Shade_564x152_New.jpg">
		    <FONT face="Verdana, Arial, Helvetica, sans-serif" color=#ffffff size=5><B><%= sTitle %></B></FONT><br>
		    <FONT size=<% =fontsize3 %> face=<% =font2 %> color=#ffffff size=3><B><%= sSubTitle %></B></FONT>
		<br>
	</TD>	
  </TR>
</TABLE><%

END SUB






'----------------- 
   SUB Standings
'----------------- 

' Creates SQL queries
OpenCon

SET rs=Server.CreateObject("ADODB.recordset")

SELECT CASE sRunByWhat

CASE "NSL"

    sSQL = "Select distinct MT.lastname, MT.firstname, RT.memberid, MT.federationcode, MT.state, MT.federationcode, MT.BirthDate," 
    sSQL = sSQL + " RG.region, RT.sc_3 as Score,"
    sSQL = sSQL + " '' as P1, '' as QR1, '' as P2, '' as QR2, '' as P3, '' as QR3, '' as P4, '' as QR4, '' as P5, '' as QR5,"
    sSQL = sSQL + " '' as P6, '' as QR6, '' as P7, '' as QR7, '' as P8, '' as QR8, '' as P9, '' as QR9, '' as P10, '' as QR10,"
    sSQL = sSQL + " '' as AWSA_Rat, '' as reg_ski, '' as regl_plc, '' as natl_plc, '' as asc1, '' as asc2, RT.div, RT.event,"
    sSQL = sSQL + " Coalesce(RT.Rank_Level, 0) AS Rank_Level, RT.Team, RT.ASC3, RT.TestField3, RT.TestField4,"

    sSQL = sSQL + " DT.Up_Age, DT.SkiYearID, SY.EndDate,"
    sSQL = sSQL + " 0 as Perc1, 0 as Perc2, 0 as Perc3, 0 as Perc4, 0 as Perc5,"  
    sSQL = sSQL + " 0 as Perc6, 0 as Perc7, 0 as Perc8, 0 as Perc9, 0 as Perc10" 

    sSQL = sSQL + " FROM "&RankTableName&" as RT JOIN "&MemberTableName&" as MT ON RT.memberid = MT.personidwithcheckdigit"
    sSQL = sSQL + " JOIN "&SkiYearTableName&" as SY ON SY.SkiYearID = '"&Session("SkiYear")&"'"
    sSQL = sSQL + " JOIN "&DivisionsTableName&" as DT ON SY.skiyearid = DT.skiyearid AND RT.Div = DT.Div"
    sSQL = sSQL + " LEFT JOIN "&RegionTableName&" as RG on lower(MT.state) = lower(RG.state)"
    sSQL = sSQL + " WHERE RT.sc_3 IS NOT NULL AND RT.div = '"&DivSelected&"' AND RT.[event] = '"&EventSelected&"'"
    sSQL = sSQL + " AND RT.SkiYearID = "&Session("SkiYear")

  IF RegionSelected <> "All" THEN
	sSQL = sSQL + " AND RG.[region] = '"&RegionSelected&"'"
  END IF

  IF StateSelected <> "All" AND StateSelected <> "XX" THEN
    	sSQL = sSQL + " AND MT.[state] = '"&StateSelected&"'"
  END IF

  IF StateSelected = "XX" THEN
    	sSQL = sSQL + " AND MT.[state] NOT IN " & USStatesList
  END IF

  SELECT CASE FederationSelected
	  CASE "USA"
		sSQL = sSQL + " AND MT.federationcode = 'USA'"
  END SELECT

  sSQL = sSQL + " order by RT.div, RT.event, RT.sc_3 DESC"
  ChoosePagesSQL sSQL,currentPage, 30



CASE "NCWSA"

    sSQL = "Select distinct MT.lastname, MT.firstname, RT.memberid, MT.federationcode, MT.state, MT.BirthDate," 

'  sSQL = "Select distinct MT.lastname, MT.firstname, RT.memberid, MT.federationcode, MT.state,"
    sSQL = sSQL + " RT.asc1, RT.asc2, RT.ASC3, RG.region, RT.awsa_rat, RT.RankScore as Score, RT.RnkScoBkup,"
    sSQL = sSQL + " RT.div, RT.event, RT.reg_ski, RT.regl_plc, RT.natl_plc,"
    sSQL = sSQL + " Coalesce(RT.Rank_Level, 0) AS Rank_Level, RT.Team, RT.TestField3, RT.TestField4,"

    sSQL = sSQL + " 0 as P1, '' as QR1, 0 as P2, '' as QR2, 0 as P3, '' as QR3, 0 as P4, '' as QR4, 0 as P5, '' as QR5,"
    sSQL = sSQL + " '0' as P6, '' as QR6, '0' as P7, '' as QR7, '0' as P8, '' as QR8, '0' as P9, '' as QR9, '0' as P10, '' as QR10,"

    sSQL = sSQL + " 0 as Perc1, 0 as Perc2, 0 as Perc3, 0 as Perc4, 0 as Perc5,"  
    sSQL = sSQL + " 0 as Perc6, 0 as Perc7, 0 as Perc8, 0 as Perc9, 0 as Perc10," 

' new 2/8/2007
    sSQL = sSQL + " DT.Up_Age, DT.SkiYearID, SY.EndDate,"
    sSQL = sSQL + " DT.SkiYearID as DTSYID"

    sSQL = sSQL + " FROM "&RankTableName&" as RT JOIN "&MemberTableName&" as MT on RT.memberid = MT.personidwithcheckdigit "
    sSQL = sSQL + " JOIN "&SkiYearTableName&" as SY ON SY.SkiYearID = '"&Session("SkiYear")&"'"
    sSQL = sSQL + " JOIN "&DivisionsTableName&" as DT ON SY.skiyearid = DT.skiyearid AND RT.Div = DT.Div"
    sSQL = sSQL + " LEFT JOIN "&RegionTableName&" as RG on lower(MT.state) = lower(RG.state) "
    sSQL = sSQL + " WHERE RT.RankScore is not null AND RT.div = '"&DivSelected&"' AND RT.[event] = '"&EventSelected&"'"
    sSQL = sSQL + " AND RT.SkiYearID = "&Session("SkiYear")

  IF RegionSelected <> "All" THEN
    	sSQL = sSQL + " AND RG.[region] = '"&RegionSelected&"'"
  END IF

  IF StateSelected <> "All" AND StateSelected <> "XX" THEN
    	sSQL = sSQL + " AND MT.[state] = '"&StateSelected&"'"
  END IF

  IF StateSelected = "XX" THEN
    	sSQL = sSQL + " AND MT.[state] NOT IN "&USStatesList
  END IF

  SELECT CASE FederationSelected
  	CASE "USA"
    		sSQL = sSQL + " AND MT.federationcode = 'USA'"
  END SELECT

  sSQL = sSQL + " order by RT.div, RT.event, RT.RankScore DESC"
  ChoosePagesSQL sSQL,currentPage, 30



CASE ELSE
' Represents AWSA National Ranking List 

  ' When &SkiYearTableName.defaultYear is "1", which is the 12 Month Ranking
  ' This IF is for any event EXCEPT Overall
  ' 
  IF EventSelected <> "O" THEN
    sSQL = "Select distinct MT.lastname, MT.firstname, RT.memberid, MT.federationcode, MT.state, MT.BirthDate," 
    sSQL = sSQL + " RT.asc1, RT.asc2, RG.region, RT.awsa_rat, RT.RankScore as Score, RT.RnkScoBkup,"
    sSQL = sSQL + " RT.Percent_01 as P1, RT.Percent_02 as P2, RT.Percent_03 as P3, RT.Percent_04 as P4, RT.Percent_05 as P5,"  
    sSQL = sSQL + " RT.Percent_06 as P6, RT.Percent_07 as P7, RT.Percent_08 as P8, RT.Percent_09 as P9, RT.Percent_10 as P10," 

'    sSQL = sSQL + " RT.Qual_Result_01 as QR1, RT.Qual_Result_02 as QR2, RT.Qual_Result_03 as QR3, RT.Qual_Result_04 as QR4, RT.Qual_Result_05 as QR5,"
'    sSQL = sSQL + " RT.Qual_Result_06 as QR6, RT.Qual_Result_07 as QR7, RT.Qual_Result_08 as QR8, RT.Qual_Result_09 as QR9, RT.Qual_Result_10 as QR10,"

    sSQL = sSQL + " RT.div, RT.event, RT.reg_ski, RT.regl_plc, coalesce(RT.natl_plc, '999') as natl_plc, DT.ZBSConversion,"
    sSQL = sSQL + " RT.Rank_Level, RT.Team, RT.ASC3, RT.TestField3, RT.TestField4,"
    sSQL = sSQL + " DT.Up_Age, DT.SkiYearID, SY.EndDate,"
    sSQL = sSQL + " RGEN.MemberID AS RGENMemberID,"

    SELECT CASE EventSelected
	CASE "S"	
	    sSQL = sSQL + " DT.Percent_01_S as Perc1, DT.Percent_02_S as Perc2, DT.Percent_03_S as Perc3, DT.Percent_04_S as Perc4, DT.Percent_05_S as Perc5,"  
	    sSQL = sSQL + " DT.Percent_06_S as Perc6, DT.Percent_07_S as Perc7, DT.Percent_08_S as Perc8, DT.Percent_09_S as Perc9, DT.Percent_10_S as Perc10" 
	CASE "T"
	    sSQL = sSQL + " DT.Percent_01_T as Perc1, DT.Percent_02_T as Perc2, DT.Percent_03_T as Perc3, DT.Percent_04_T as Perc4, DT.Percent_05_T as Perc5,"  
	    sSQL = sSQL + " DT.Percent_06_T as Perc6, DT.Percent_07_T as Perc7, DT.Percent_08_T as Perc8, DT.Percent_09_T as Perc9, DT.Percent_10_T as Perc10" 
	CASE "J"
	    sSQL = sSQL + " DT.Percent_01_J as Perc1, DT.Percent_02_J as Perc2, DT.Percent_03_J as Perc3, DT.Percent_04_J as Perc4, DT.Percent_05_J as Perc5,"  
	    sSQL = sSQL + " DT.Percent_06_J as Perc6, DT.Percent_07_J as Perc7, DT.Percent_08_J as Perc8, DT.Percent_09_J as Perc9, DT.Percent_10_J as Perc10" 
	CASE "O"
	    sSQL = sSQL + " DT.Percent_01_O as Perc1, DT.Percent_02_O as Perc2, DT.Percent_03_O as Perc3, DT.Percent_04_O as Perc4, DT.Percent_05_O as Perc5,"  
	    sSQL = sSQL + " DT.Percent_06_O as Perc6, DT.Percent_07_O as Perc7, DT.Percent_08_O as Perc8, DT.Percent_09_O as Perc9, DT.Percent_10_O as Perc10" 
    END SELECT	

    sSQL = sSQL + " FROM "&RankTableName&" as RT"

    sSQL = sSQL + " JOIN "&MemberTableName&" as MT on RT.memberid = MT.personidwithcheckdigit "

    sSQL = sSQL + " LEFT JOIN "&RegGenTableName&" AS RGEN ON RGEN.MemberID=RT.MemberID AND LEFT(RGEN.TourID,6)='07W999'" 	

    sSQL = sSQL + " JOIN "&SkiYearTableName&" as SY ON SY.SkiYearID = '"&Session("SkiYear")&"'"

    sSQL = sSQL + " JOIN "&DivisionsTableName&" as DT ON SY.skiyearid = DT.skiyearid AND RT.Div = DT.Div"

    sSQL = sSQL + " LEFT JOIN "&RegionTableName&" as RG on lower(MT.state) = lower(RG.state) "

    sSQL = sSQL + " WHERE RT.RankScore is not null AND RT.div = '"&DivSelected&"' AND RT.[event] = '"&EventSelected&"'"

    sSQL = sSQL + " AND RT.SkiYearID = "&Session("SkiYear")

    ' Region, State and Federation 
    IF RegionSelected <> "All" THEN
      	sSQL = sSQL + " AND RG.[region] = '"&RegionSelected&"'"
    END IF

    IF StateSelected <> "All" AND StateSelected <> "XX" THEN
      	sSQL = sSQL + " AND MT.[state] = '"&StateSelected&"'"
    END IF
    
'    IF StateSelected = "XX" THEN
'      	sSQL = sSQL + " AND MT.[state] NOT IN "&USStatesList
'    END IF
    
    SELECT CASE FederationSelected
    	CASE "USA"
      		sSQL = sSQL + " AND MT.federationcode = 'USA'"
	CASE ELSE

    END SELECT

    sSQL = sSQL + " ORDER BY RT.div, RT.event, RT.RankScore DESC, natl_plc"



    ChoosePagesSQL sSQL,currentPage, 30

  ' ----------  If Event is OVERALL  ----------------
  ELSE
    sSQL = "Select distinct MT.lastname, MT.firstname, MT.federationcode, MT.state, MT.BirthDate," 
    sSQL = sSQL + " RG.Region, RT.memberid, RT.RankScore as Score, RT.RnkScoBkup, RT.Div, RT.Event,"
    sSQL = sSQL + " RT.Percent_01 as P1, RT.Percent_02 as P2, RT.Percent_03 as P3, RT.Percent_04 as P4, RT.Percent_05 as P5,"  
    sSQL = sSQL + " RT.Percent_06 as P6, RT.Percent_07 as P7, RT.Percent_08 as P8, RT.Percent_09 as P9, RT.Percent_10 as P10," 

'    sSQL = sSQL + " RT.Qual_Result_01 as QR1, RT.Qual_Result_02 as QR2, RT.Qual_Result_03 as QR3, RT.Qual_Result_04 as QR4, RT.Qual_Result_05 as QR5,"
'    sSQL = sSQL + " RT.Qual_Result_06 as QR6, RT.Qual_Result_07 as QR7, RT.Qual_Result_08 as QR8, RT.Qual_Result_09 as QR9, RT.Qual_Result_10 as QR10,"

    sSQL = sSQL + " RT.Rank_Level, RT.Team, RT.ASC3, RT.TestField3, RT.TestField4,"

	
    sSQL = sSQL + " DT.Up_Age, DT.SkiYearID, SY.EndDate,"
    sSQL = sSQL + " DT.Percent_01_O as Perc1, DT.Percent_02_O as Perc2, DT.Percent_03_O as Perc3, DT.Percent_04_O as Perc4, DT.Percent_05_O as Perc5,"  
    sSQL = sSQL + " DT.Percent_06_O as Perc6, DT.Percent_07_O as Perc7, DT.Percent_08_O as Perc8, DT.Percent_09_O as Perc9, DT.Percent_10_O as Perc10" 

    sSQL = sSQL + " FROM "&RankTableName&" as RT JOIN "&MemberTableName&" as MT on RT.memberid = MT.personidwithcheckdigit "
    sSQL = sSQL + " JOIN "&SkiYearTableName&" as SY ON SY.SkiYearID = '"&Session("SkiYear")&"'"
    sSQL = sSQL + " JOIN "&DivisionsTableName&" as DT ON SY.skiyearid = DT.skiyearid AND RT.Div = DT.Div"
    sSQL = sSQL + " LEFT JOIN "&RegionTableName&" as RG on lower(MT.state) = lower(RG.state) "

    sSQL = sSQL + " WHERE RT.RankScore is not null AND RT.div = '"&DivSelected&"' AND RT.[event] = '"&EventSelected&"'"
    sSQL = sSQL + " AND RT.SkiYearID = "&Session("SkiYear")

    IF RegionSelected <> "All" THEN
      	sSQL = sSQL + " AND RG.[region] = '"&RegionSelected&"'"
    END IF
    
    IF StateSelected <> "All" AND StateSelected <> "XX" THEN
      	sSQL = sSQL + " AND MT.[state] = '"&StateSelected&"'"
    END IF

    IF StateSelected = "XX" THEN
      	sSQL = sSQL + " AND MT.[state] NOT IN " & USStatesList
    END IF

    SELECT CASE FederationSelected
	CASE "USA"
      	    sSQL = sSQL + " AND MT.federationcode = 'USA'"
    END SELECT

    sSQL = sSQL + " ORDER by RT.div, RT.RankScore DESC"

    ChoosePagesSQL sSQL,currentPage, 30  

  END IF



END SELECT

rowCount = 0
Response.Write("<BR>")



' -----------------------------------------------------------------------------------------------------------
' --- Established the COA for top of screen
' --- More work is required here because at COD, the COA becomes static not dynamic.
' -----------------------------------------------------------------------------------------------------------


IF NOT rs.eof AND sRunByWhat <>"NCWSA" AND sRunByWhat<>"NSL" THEN

	rs.MoveFirst
	LastRank_Level = rs("Rank_Level")
	LastScore = rs("Score")
	rs.MoveNext

  	DO WHILE NOT rs.eof

		IF LastRank_Level >= rs("Perc8")/10 AND rs("Rank_Level") < rs("Perc8")/10 THEN
			ThisCOA = LastScore
		END IF

		LastRank_Level = rs("Rank_Level")
		LastScore = rs("Score")
		
		rs.MoveNext
	LOOP
	rs.MoveFirst
END IF

'markdebug("ThisCOA = "&ThisCOA)

IF SingleRanking="Find My Ranking" THEN
	FindRankingInstances
ELSE
	DisplayDropDowns
	DisplayRankList
END IF



END SUB








' -----------------------
   SUB DisplayDropdowns
' -----------------------

'SELECT CASE sRunByWhat
 ' CASE "National"
'	radiostate="on"


' -------------------------------------------------------------------------------------------------
' ------------------   Begin form to select filtering parameters  --------------------------------
' -------------------------------------------------------------------------------------------------
'
%>
<form method=post action="<%=ThisFileName%>">
<input type="hidden" name="adminmenulevel" value="<%=adminmenulevel%>">
<input type="hidden" name="sMemberID" value="<%=sMemberID%>">

<TABLE width=100% border="0" align=center><%

'markdebug("sRunByWhat="&sRunByWhat)

' -----------------------------  Build EVENT dropdown list  ----------------------------%>

<tr>
  <td colspan=6 align="center">
	<input type=radio NAME=pvar VALUE="National" <% IF pvar="National" THEN response.write "checked" %> onclick=submit()>
	<FONT size=<% =fontsize3 %> face=<% =font1 %> COlOR=<% =textcolor2 %>><b>National&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></font>
	<input type=radio NAME=pvar VALUE="Junior" <% IF pvar="Junior" THEN response.write "checked" %> onclick=submit()>
	<FONT size=<% =fontsize3 %> face=<% =font1 %> COlOR=<% =textcolor2 %> checked><b>Junior&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></font>
	<input type=radio NAME=pvar VALUE="NCWSA" <% IF pvar="NCWSA" THEN response.write "checked" %> onclick=submit()>
	<FONT size=<% =fontsize3 %> face=<% =font1 %> COlOR=<% =textcolor2 %> checked><b>Collegiate&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b></font>
	<input type=radio NAME=pvar VALUE="NSL" <% IF pvar="NSL" THEN response.write "checked" %> onclick=submit()>
	<FONT size=<% =fontsize3 %> face=<% =font1 %> COlOR=<% =textcolor2 %>><b>Grassroots</b></font>
	<br><br>
  </td>
</tr>
<tr>
  <td width=8%>
     <font size=<% =fontsize2 %> face=<% =font1 %>><b>Event:</b></font>
  </td>

  <td width=25%>	
	<select name='event'>
	  <option value ='S' <%IF EventSelected="S" THEN response.write(" selected")%>>Slalom</Option><br>
	  <option value ='J' <%IF EventSelected="J" THEN response.write(" selected")%>>Jump</Option><br>
	  <option value ='T' <%IF EventSelected="T" THEN response.write(" selected")%>>Trick</Option><br><%

	  IF Request("event") = "O" THEN Response.Write("<option value =""O"" selected>Overall Scores</option><br>")
	  IF Request("event") <> "O" THEN Response.Write("<option value =""O"">Overall Scores</option><br>")%>
	</select>
  </td><%


' -----------------------------  Build REGION dropdown list  ----------------------------%>

  <td width=8%>
     <font size=<% =fontsize2 %> face=<% =font1 %>><b>Region:</b></font> 
  </td>
  <td width=25%>
	<select name='region'>
	<option value ='All'<%IF RegionSelected = "All" THEN Response.Write(" selected ")%>>All </Option><br>
	<option value ='1'<%IF RegionSelected = "1" THEN Response.Write(" selected ")%>>S. Central</Option><br>
	<option value ='2'<%IF RegionSelected = "2" THEN Response.Write(" selected ")%>>MidWest</Option><br>
	<option value ='3'<%IF RegionSelected = "3" THEN Response.Write(" selected ")%>>West</Option><br>
	<option value ='4'<%IF RegionSelected = "4" THEN Response.Write(" selected ")%>>South</Option><br>
	<option value ='5'<%IF RegionSelected = "5" THEN Response.Write(" selected ")%>>East</Option><br>
	</select>
  </td><%
  IF sRunByWhat="NSL" OR sRunByWhat="NCWSA" THEN %>
		<td>&nbsp;</td>
		<td>&nbsp;</td><%
  ELSE 
	IF TRIM(session("SkiYear"))=1 THEN %>
	  <td colspan=2 align=left><font size=<% =fontsize2 %> face=<% =font1 %>><b>Nationals Qualification</b></font></td><%
	ELSE %>
	  <td colspan=2>&nbsp;</td><%
	END IF
  END IF %>	

 
</tr><%



' ------------------------------  Build DIVISION dropdown list  ----------------------------------

' Preloads dropdown with values based on RunByWhat variable passed from Menu Link 

%>
<tr>
  <td>
    <font size=<% =fontsize2 %> face=<% =font1 %>><b>Division:</b></font>
  </td>
  <td>
	<select name='div'><%
	sSQL = "Select distinct RT.div, DT.div_name from "&RankTableName&" as RT JOIN "&DivisionsTableName&" as DT ON RT.div = DT.div"

	SELECT CASE sRunByWhat
  		CASE "National", "NSL"
    			sSQL = sSQL + " WHERE lower(left(RT.div,1)) <> 'i' AND lower(left(RT.div,1)) <> 'n' AND lower(left(RT.div,1)) <> 'c'"
	  	CASE "Junior"
    			sSQL = sSQL + " WHERE lower(left(RT.div,1)) = 'b' or lower(left(RT.div,1)) = 'g'"
	  	CASE "NCWSA"
			sSQL = sSQL + " WHERE lower(RT.div) = 'cm' or lower(RT.div) = 'cw'"
	END SELECT

	sSQL = sSQL + " order by RT.div"
	rsSelectFields.open sSQL, SConnectionToTRATable


	' ---  This section deals with case WHERE no scores exist for any of the divisions  ---
	IF NOT rsSelectFields.eof THEN 
	  	rsSelectFields.movefirst
  		DO WHILE NOT rsSelectFields.eof
	    		IF TRIM(rsSelectFields("Div")) = DivSelected THEN
      				response.write("<option value ="""&rsSelectFields("Div")&""" selected>"&rsSelectFields("Div")&" - "&rsSelectFields("Div_Name")&"</option><br>")
    			ELSE
      				response.write("<option value ="""&rsSelectFields("Div")&""">"&rsSelectFields("Div")&" - "&rsSelectFields("Div_Name")&"</option><br>")
	    		END IF	
			rsSelectFields.moveNEXT
  		LOOP
	ELSE
  		response.write("<option value =""None"" selected>None Available</option>")
	END IF

	rsSelectFields.close %>
	</select>
  </td><%


' -------------------------------------  Build STATE dropdown list --------------------------------

' Preloads dropdown with values based on RunByWhat variable passed from Menu Link %>

  <td>
    <font size=<% =fontsize2 %> face=<% =font1 %>><b>State:</b></font>
  </td>

  <td><%
	Dim kvar, statearray
	StateArray = Split(USStatesList2,",") %>  

  	<select name="state"><%
	FOR kvar = 0 TO UBOUND(StateArray)
		IF StateSelected = TRIM(StateArray(kvar)) THEN
			response.write("<option value = """&StateSelected&""" SELECTED>"&StateSelected&"</option>")
		ELSE
			response.write("<option value = """&StateArray(kvar)&""">"&StateArray(kvar)&"</option>")
		END IF
	NEXT  %>
  </select>
  </td><%


  IF sRunByWhat="NSL" OR sRunByWhat="NCWSA" THEN %>
		<td>&nbsp;</td>
		<td>&nbsp;</td><%
  ELSE 
	IF TRIM(session("SkiYear"))=1 THEN %>
	  <td width=15%>
		<font size=<% =fontsize2 %> face=<% =font1 %>><b>Cut-Off Avg</b></font>
	  </td>

	  <td width=8% bgcolor="<%=tcolor03%>" align="center">
	  	<font size=<% =fontsize2 %> face=<% =font1 %>>&nbsp;<%=formatnumber(ThisCOA,2)%></font>
	  </td><%
	ELSE %>
		<td>&nbsp;</td>
		<td>&nbsp;</td><%
	END IF 
  END IF %>


</tr><%



' --------------------------------- Build SKI YEAR dropdown list  -------------------

' ---- The value of the dropdown is established based on Session("SkiYear") variable ----  %>



<tr>
  <td>
    <font size=<% =fontsize2 %> face=<% =font1 %>><b>Range:</b></font>&nbsp 
  </td>

  <td>	
	<select name='SkiYear'><%

	sSQL = "SELECT DISTINCT RT.SkiYearID, SY.SkiYearName"
	sSQL = sSQL + " FROM " &RankTableName&" AS RT"
	sSQL = sSQL + " JOIN " &SkiYearTableName&" AS SY ON RT.SkiYearID = SY.SkiYearID"

	' --- NCWSA does not display 12 Month Rankings
	IF sRunByWhat="NCWSA" THEN
		sSQL = sSQL + " WHERE SY.SkiYearID <> 1"
	END IF

	rsSelectFields.open sSQL, SConnectionToTRATable

	' Loads dropdown and sets default to Session("SkiYear")
	DO WHILE NOT rsSelectFields.eof


		IF TRIM(rsSelectFields("SkiYearID")) = TRIM(session("SkiYear")) THEN
			response.write("<option value =""" & rsSelectFields("SkiYearID") &""" selected>")
			response.write(rsSelectFields("SkiYearName"))
			response.write("</option><br>")
		ELSE
			response.write("<option value =""" & rsSelectFields("SkiYearID") &""">")
			response.write(rsSelectFields("SkiYearName"))
			response.write("</option><br>")
		END IF 

		rsSelectFields.moveNEXT
	LOOP

	rsSelectFields.close %>
	</select>
  </td>

 <% ' --------------------------------  Build FEDERATION dropdown list  -------------------- %>

 <td>
    <font size=<% =fontsize2 %> face=<% =font1 %>><b>Federation:</b></font> 
  </td> 	
	
  <td>
	<select name="Include_International">
	<option value="ALL"<%IF FederationSelected = "ALL" THEN Response.Write(" selected")%>>All Feds</option>
	<option value="USA"<%IF FederationSelected = "USA" THEN Response.Write(" selected")%>>USA</option>
	</select>
  </td><%

  IF sRunByWhat="NSL" OR sRunByWhat="NCWSA" THEN %>
		<td>&nbsp;</td>
		<td>&nbsp;</td><%
  ELSE 
   	IF TRIM(session("SkiYear"))=1 THEN %>
		<td>
		  <font size=<% =fontsize2 %> face=<% =font1 %>><b>Cut-Off Date:</b> </font>
		</td>

		<td>
    		  <font size=<% =fontsize2 %> face=<% =font1 %>>&nbsp;<%=CutOffDate%></font>
		</td><%
	ELSE %>
		<td>&nbsp;</td>
		<td>&nbsp;</td><%
	END IF

  END IF %>


</tr><%


' ----------------------  Option to REMOVE CTF Overlay  --------------------------------------

%>
<tr>
   <td>&nbsp;</td>
   <td colspan=1 align="left"><input type=submit value="Update Display"></td>
   <td colspan=1>&nbsp;</td>
   <td colspan=1 align="left"><input type=submit name="SingleRanking" value="Find My Ranking"></td><%

  IF sRunByWhat="NSL" OR sRunByWhat="NCWSA" THEN %>
		<td>&nbsp;</td>
		<td>&nbsp;</td><%
  ELSE 
   	IF TRIM(session("SkiYear"))=1 THEN %>
		<td>
		  <font size=<% =fontsize2 %> face=<% =font1 %>><b>Regl/Natl Place</b> </font>
		</td>

		<td>
    		  <font size=<% =fontsize2 %> face=<% =font1 %>>Top 5</font>
		</td><%
	ELSE %>
		<td>&nbsp;</td>
		<td>&nbsp;</td><%
	END IF

  END IF %>

</tr>

<tr>

   <td colspan=2><font color="red"><small>*</font><font face=<% =font2 %>><small> - Missing Score Penalty Applied</small></font></td>
   <td colspan=2 align="center"><font size=<% =fontsize2 %> face=<% =font1 %>>Active Member: <%=sFullName%></font></td>
   <td colspan=2>&nbsp;</td>

</tr>

</form>
</table><%




END SUB


' --------------------------
  SUB FindRankingInstances
' -------------------------- 

sSkiYearID=TRIM(Session("SkiYear"))

SET rsRankList=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT * FROM "&RankTableName&" AS RT"
sSQL = sSQL + " JOIN "&MemberTableName&" AS MT ON MT.PersonIDwithCheckDigit=RT.MemberID"
sSQL = sSQL + " WHERE MemberID='"&sMemberID&"' AND SkiYearID='"&sSkiYearID&"'"
sSQL = sSQL + " AND LOWER(LEFT(Div,1))<>'i' AND LOWER(LEFT(Div,1))<>'x' AND LOWER(LEFT(Div,1))<>'y'"
rsRankList.open sSQL, SConnectionToTRATable

IF NOT rsRankList.eof THEN %>

    <center><font size=<% =fontsize3 %> face=<% =font2 %> COlOR="<%=TextColor1%>"><b> Rankings For</font>
      <font size=4 face=<% =font2 %> COlOR="<%=TextColor2%>"><br><%=rsRankList("FirstName")%>&nbsp;<%=rsRankList("LastName")%></b></FONT>
      <font size=2 face=<% =font2 %> COlOR="<%=TextColor1%>"><br><%=rsRankList("City")%>, <%=rsRankList("State")%></b></FONT>
    </center><br>	

    <TABLE ALIGN="Center" BORDER="1" CELLPADDING="3" CELLSPACING="0" WIDTH="60%" BGCOLOR="#FFFFFF" >
	
    <TR>
    	<TD ALIGN="Center" Width=20% vAlign="top" bgcolor="<%=HeadColor1%>"><font size=<% =fontsize2 %> face=<% =font2 %> COlOR="#000000"><b>Universal<br>Rank</b></FONT></TD>
    	<TD ALIGN="Center" Width=25% vAlign="top" bgcolor="<%=HeadColor1%>"><font size=<% =fontsize2 %> face=<% =font2 %>  COlOR="#000000"><b><br>Event</b></FONT></TD> 	
    	<TD ALIGN="Center" Width=20% vAlign="top" bgcolor="<%=HeadColor1%>"><font size=<% =fontsize2 %> face=<% =font2 %>  COlOR="#000000"><b><br>Division</b></FONT></TD> 	
    	<TD ALIGN="Center" Width=35% vAlign="top" bgcolor="<%=HeadColor1%>"><font size=<% =fontsize2 %> face=<% =font2 %>  COlOR="#000000"><b>Ranking<br>Score</b></FONT></TD> 	
    </TR><%
 	
    DO WHILE NOT rsRankList.eof 

	sEvent=TRIM(rsRankList("Event"))
	sDiv=TRIM(rsRankList("Div"))

	sRankScore = rsRankList("RankScore")
	SELECT CASE TRIM(rsRankList("Event"))
		CASE "J"
			sEventName="Jump"
			sRankScore=formatnumber(rsRankList("RankScore"),2)		
		CASE "S"
			sEventName="Slalom"
			'sRankScore=formatnumber(rsRankList("RankScore"),2)		

		CASE "T"
			sEventName="Trick"
			sRankScore=formatnumber(rsRankList("RankScore"),0)
		CASE "O"
			sEventName="Overall"
			sRankScore=formatnumber(rsRankList("RankScore"),0)
	END SELECT
	%>

    	<TR>
	  <TD ALIGN="Center" Width=9% vAlign="top" bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font2 %> COlOR="#000000"><b><%=rsRankList("RankNum")%></b></FONT></TD>
	  <TD ALIGN="Center" Width=9% vAlign="top" bgcolor="<%=TableColor1%>">
		<font size=<% =fontsize2 %> face=<% =font2 %> COlOR="#000000">
		  <b><a href="/<%=ThisFileName%>?MyEvent=<%=sEvent%>&MyDiv=<%=rsRankList("Div")%>&pvar=<%=sRunByWhat%>&sMemberID=<%=sMemberID%>"><%=sEventName%></a></b>
		</FONT></TD>
	  <TD ALIGN="Center" vAlign="top" bgcolor="<%=TableColor1%>">
		<font size=<% =fontsize2 %> face=<% =font2 %>  COlOR="#000000">
		  <b><a href="/<%=ThisFileName%>?MyEvent=<%=sEvent%>&MyDiv=<%=rsRankList("Div")%>&pvar=<%=sRunByWhat%>&sMemberID=<%=sMemberID%>"><%=rsRankList("Div")%></a></b>
		</FONT></TD> 	
	  <TD ALIGN="Center" vAlign="top" bgcolor="<%=TableColor1%>"><font size=<% =fontsize2 %> face=<% =font2 %>  COlOR="#000000"><b><%=sRankScore%></b></FONT></TD> 	
	</TR><%
        rsRankList.MoveNext	
    LOOP %>

    </TABLE>
    <br>
<center><font size=<% =fontsize2 %> face=<% =font2 %> COlOR="#000000"><b> Universal Rank includes International skiers with US scores.  Actual ranking may be different.</b></font></center>
<br><br>
<center><font size=<% =fontsize3 %> face=<% =font2 %> COlOR="#000000"><b> Click on Event Link Above to Display Rankings </b></font></center>
<br>
<form method=post action="<%=ThisFileName%>">
  <center><input type=submit name="NewMember" value="New Member"></center>
</form>
<%	
ELSE  %>

	<center>
	<font size=<% =fontsize3 %> face=<% =font2 %> COlOR="<%=TextColor3%>"><b> No Records Found For This Member In This Range</font>
	</center>
	<form method=post action="<%=ThisFileName%>?pvar=<%=sRunByWhat%>&sMemberID=<%=sMemberID%>">
	<center><input type=submit name="Continue" value="Continue"></center>
	</form>
	<%
END IF


END SUB




' -----------------------
   SUB DisplayRankList
' -----------------------



' ---------------   Top of large condition of branching to most of rest of code   ------------------

IF rs.eof THEN
	%><br><br>
    	<font color="red">No Rankings Found With These Filter Settings.</font><% 
ELSE 

	KickTrafficCounter("RankPages")

  %>

    <TABLE BORDER="1" CELLPADDING="3" CELLSPACING="0" WIDTH="100%" BGCOLOR="#FFFFFF" >
    <TR>
    	<TD ALIGN="Center" Width=9% vAlign="top" bgcolor="<%=HeadColor1%>"><font size=<% =fontsize2 %> face=<% =font2 %> COlOR="#000000"><b> Rank</b></FONT></TD>
    	<TD ALIGN="Center" vAlign="top" bgcolor="<%=HeadColor1%>"><font size=<% =fontsize2 %> face=<% =font2 %>  COlOR="#000000"><b>Name</b></FONT></TD><% 

	IF Request("PVar") = "NSL" THEN %>
    		<TD ALIGN="Center" vAlign="top" bgcolor="<%=HeadColor1%>"><font size=<% =fontsize2 %> face=<% =font2 %>  COlOR="#000000"><b>NSL Placement Points</b></FONT></TD><% 
	ELSE %>
    		<TD ALIGN="Center" vAlign="top" bgcolor="<%=HeadColor1%>"><font size=<% =fontsize2 %> face=<% =font2 %>  COlOR="#000000"><b>Score</b></FONT></TD><% 
	END IF %>

    	<TD ALIGN="Center" vAlign="top" bgcolor="<%=HeadColor1%>"><font size=<% =fontsize2 %> face=<% =font2 %>  COlOR="#000000"><b>Member #</b></FONT></TD><% 

	' --- Changed 9-7-2007 to make condition impossible
	IF 1=2 AND Request("PVar") <> "NSL" AND EventSelected <> "O" THEN %>
    		<TD ALIGN="Center" vAlign="top" bgcolor="<%=HeadColor1%>"><font size=<% =fontsize2 %> face=<% =font2 %> COlOR="#000000"><b>Rating</b></FONT></TD><% 
	END IF %>

	<TD ALIGN="Center" vAlign="top" bgcolor="<%=HeadColor1%>"><font size=<% =fontsize2 %> face=<% =font2 %>  COlOR="#000000"><b>ST</b></FONT></TD><%

	IF Request("Pvar") = "NCWSA" THEN
		%><TD ALIGN="Center" vAlign="top" bgcolor="<%=HeadColor1%>"><font size=<% =fontsize2 %> face=<% =font2 %>  COlOR="#000000"><b>Team</b></FONT></TD><% 
	ELSE
		%><TD ALIGN="Center" vAlign="top" bgcolor="<%=HeadColor1%>"><font size=<% =fontsize2 %> face=<% =font2 %>  COlOR="#000000"><b>Reg</b></FONT></TD><% 
	END IF

	IF Request("PVar") <> "NSL" AND Request("PVar") <> "NCWSA" AND EventSelected <> "O" THEN %>
    		<TD ALIGN="Center" vAlign="top" bgcolor="<%=HeadColor1%>"><font size=<% =fontsize2 %> face=<% =font2 %>  COlOR="#000000"><b>Rgnl Plc</b></FONT></TD>
		<TD ALIGN="Center" vAlign="top" bgcolor="<%=HeadColor1%>"><font size=<% =fontsize2 %> face=<% =font2 %>  COlOR="#000000"><b>Ntnl Plc</b></FONT></TD><% 
	END IF 

	%><TD ALIGN="Center" vAlign="top" bgcolor="<%=HeadColor1%>"><font size=<% =fontsize2 %> face=<% =font2 %>  COlOR="#000000"><b>Mem Fed</b></FONT></TD><%
	IF sRunByWhat<>"NCWSA" AND sRunByWhat<>"NSL" THEN %>
		<TD ALIGN="Center" vAlign="top" bgcolor="<%=HeadColor1%>"><font size=<% =fontsize2 %> face=<% =font2 %> COlOR="#000000"><b>Pctl</b></FONT></TD><%
	END IF 
	IF 1=2 AND EventSelected <> "O" AND RunByWhat <> "NSL" AND RunByWhat<>"NCWSA" THEN %>
	   <TD ALIGN="Center" vAlign="top" bgcolor="<%=HeadColor1%>"><font size=<% =fontsize2 %> face=<% =font2 %> COlOR="#000000"><a title="Entered in 2007 Goode National Championships"><b>Natls</b></a></FONT></TD><%
	END IF

	' ------  May be temporary -----
 	%>
		<% IF sAdminMenuLevel >= "40" THEN %>
			<TD ALIGN="Center" vAlign="top" bgcolor="<%=HeadColor1%>"><font size=<% =fontsize2 %> face=<% =font2 %> COlOR="#000000"><b>Birth</b></FONT></TD>
		<% END IF %>

    </TR><%



	' ----------------------------------------------------------------------------------
	' ------------------------  BEGIN DISPLAYING DATA ----------------------------------
	' ----------------------------------------------------------------------------------


	' --- Use first record to intialize the percentages from DivionsTableName ---
	tPerc1=rs("Perc1")/10
	tPerc2=rs("Perc2")/10
	tPerc3=rs("Perc3")/10
	tPerc4=rs("Perc4")/10
	tPerc5=rs("Perc5")/10
	tPerc6=rs("Perc6")/10
	tPerc7=rs("Perc7")/10
	tPerc8=rs("Perc8")/10
	tPerc9=rs("Perc9")/10
	tPerc10=rs("Perc10")/10



	' --- INITIALIZES the Ranking related memory fields for deal with ties.

	' --- RecordNum is essentially the record count
	' --- MemoryScore is the Score of the 
    	' --- MemoryRank stores the highest value of placement - for which subsequent records may be tied 
	' --- tRankScore is the Score of the current record

	RecordNum = 1
	MemoryRank = 1
	MemoryScore = rs("Score")

   
    	' ---  After storing the values from the FIRST record then move to the second record to see if tied to know
	' ---     whether the FIRST record should have a T after it.  All others

	' --- Move to 2nd record ---
    	rs.MoveNEXT

	IF NOT rs.EOF THEN
		' --- Initializes 2nd record in query --- 
		DefineRankingDataLine
	END IF

	' --- If the score from last tied record is same as current score 
	IF MemoryScore = tRankScore THEN
		RankValueWithTies = "1T"
	ELSE
		RankValueWithTies = "1"
	END IF

	' --- Now move back to FIRST record and initialize First record in query ---
    	rs.MoveFIRST
	DefineRankingDataLine
  



        ' -----  BEGINNING OF LOOPING FOR DISPLAYING ALL RECORDS MATCHING QUERY  ------------------------
	' -----   Loops thru the remaining 2,3...nth records

    	DO WHILE NOT rs.eof

		'IF rs("MemberID")=sMemberID THEN

			' --- Displays one line of ranking list ---
			DisplayRankingLine

		'END IF

		' --- Initializes NEXT record in query --- 
		rs.moveNEXT
		RecordNum = RecordNum + 1

		IF NOT rs.eof THEN
			' --- Defines the CURRENT record ---
			DefineRankingDataLine

			' --- If the score from PREVIOUS record is same as current score 
			IF cdbl(MemoryScore) = cdbl(tRankScore) THEN
				RankValueWithTies = MemoryRank&"T"
			ELSE

				MemoryRank = RecordNum
				MemoryScore = rs("Score")
				
				' --- Move to NEXT record to see if tied---
				rs.MoveNEXT
			    	IF NOT rs.eof THEN
					' --- Initializes the record beyond the current record in query --- 
					DefineRankingDataLine

					' --- If the score from last tied record is same as current score 
					IF MemoryScore = tRankScore THEN
						RankValueWithTies = RecordNum&"T"
					ELSE
						RankValueWithTies = RecordNum
					END IF

				ELSE
					' --- Can't be tied with EOF so set it to the current record ---
					RankValueWithTies = RecordNum
				END IF

				' --- Now move back to CURRENT record and initialize ---
				rs.MovePREVIOUS
				DefineRankingDataLine

			END IF
		ELSE


		END IF

	LOOP  %>

       </TR>
    </TABLE><%

    DisplayPercentilesandPageFooter


END IF

CloseCon

END SUB





' -------------------------------------
  SUB DisplayPercentilesandPageFooter
' -------------------------------------

    ' Writes percentages in text at bottom of list	
    IF RemoveCTF <> "on" AND (sRunByWhat = "National" OR sRunByWhat = "Junior") THEN

'        Response.Write ("<table border=0><tr><td bgcolor=" & tcolor01 &">Level 10 &nbsp;&nbsp;&nbsp;<font size=2> (" & tPerc10*10 & ") Percentile</font></td></tr></table>")
'        Response.Write ("<table border=0><tr><td bgcolor=" & tcolor02 &">Level 9 &nbsp;&nbsp;&nbsp;<font size=2> (" & tPerc9*10 & ") Percentile</font></td></tr></table>")
        Response.Write ("<table border=0><tr><td bgcolor=" & tcolor03 &">Level 8 &nbsp;&nbsp;&nbsp;<font size=2> (" & tPerc8*10 & ") Percentile</font></td></tr></table>")
        Response.Write ("<table border=0><tr><td bgcolor=" & tcolor04 &">Level 7 &nbsp;&nbsp;&nbsp;<font size=2> (" & tPerc7*10 & ") Percentile</font></td></tr></table>")
        Response.Write ("<table border=0><tr><td bgcolor=" & tcolor05 &">Level 6 &nbsp;&nbsp;&nbsp;<font size=2> (" & tPerc6*10 & ") Percentile</font></td></tr></table>")
        Response.Write ("<table border=0><tr><td bgcolor=" & tcolor06 &">Level 5 &nbsp;&nbsp;&nbsp;<font size=2> (" & tPerc5*10 & ") Percentile</font></td></tr></table>")
        Response.Write ("<table border=0><tr><td bgcolor=" & tcolor07 &">Level 4 &nbsp;&nbsp;&nbsp;<font size=2> (" & tPerc4*10 & ") Percentile</font></td></tr></table>")
        Response.Write ("<table border=0><tr><td bgcolor=" & tcolor08 &">Level 3 &nbsp;&nbsp;&nbsp;<font size=2> (" & tPerc3*10 & ") Percentile</font></td></tr></table>")
'        Response.Write ("<table border=0><tr><td bgcolor=" & tcolor09 &">Level 2 &nbsp;&nbsp;&nbsp;<font size=2> (" & tPerc2*10 & ") Percentile</font></td></tr></table>")
'        Response.Write ("<table border=0><tr><td bgcolor=" & tcolor10 &">Level 1 &nbsp;&nbsp;&nbsp;<font size=2> (" & tPerc1*10 & ") Percentile</font></td></tr></table>")

    END IF

    %>
    <font color=<%=textcolor1%>><small>T Indicates a Tied Score.</small></font><br>
    <font color="red"><small>* Indicates a penalty was applied.</small></font>
    <br>
    <%



      ' Displays the last re-calculation date/time at bottom of screen	
      sSQL = "SELECT * FROM " & SkiYearTableName & " WHERE "

      IF session("SkiYear") = "0" THEN
        sSQL = sSQL + "DefaultYear = 1"
      ELSE
        sSQL = sSQL + "SkiYearID = " + SQLClean(session("skiyear"))
      END IF

      rsSelectFields.open sSQL, SConnectionToTRATable

      IF not rsSelectFields.eof THEN
        response.write ("<small><small>Rankings last updated at " & rsSelectFields("LastRecalc") & ".</small></small>")
      END IF

      rsSelectFields.close  %>
    <br><br>
    <%



END SUB




' ----------------------------------------------------------------------------------
    SUB DisplayRankingLine	' --- Displays a single line of the ranking list ---
' ----------------------------------------------------------------------------------


IF RemoveCTF <> "on" THEN
	Response.Write (DefineRowColor)
END IF 

IF rs("MemberID")=sMemberID THEN
	sTextColor=TextColor3
END IF %>

<TD ALIGN="Center" vAlign="top" bgcolor="<%=sTextColor%>">
<font size=<% =fontsize2 %> face=<% =font2 %> COlOR="<%=TextColor1%>"><%=RankValueWithTies%></font>
		                  
<TD ALIGN="Left" vAlign="top"><a href="/<%=ThisFileName%>?NSL=<% IF Request("Pvar") = "NSL" THEN Response.Write("1") ELSE Response.Write("0")%>&sMemberID=<%=tMemberID%><% IF sRunByWhat = "NCWSA" THEN Response.Write ("&div=" & DivSelected) %>&event=<%=EventSelected%>&pvar=ByMember"><font size=<% =fontsize2 %> face=<% =font2 %>>&nbsp;<% Response.Write(tName) %></FONT></a></TD>
<TD ALIGN="Center" vAlign="top">
  <font size=<% =fontsize2 %> face=<% =font2 %>  COlOR="<%=TextColor1%>">&nbsp;<%


  ' --- Present Ranking Score and Backup Detail
      
  IF sRunByWhat <> "NSL" THEN
	Response.Write ("<a title='" & tRnkScoBkup & "'>" & tFmtScore & "</a>")
  ELSE
	Response.Write tFmtScore
  END IF

  ' --- Tack on red asterisk unless Backup includes "NO Penalty"

  IF (sRunByWhat <>"NSL" AND instr(tRnkScoBkup,"NO Penalty")=0) THEN
	Response.Write("<font color=""red""><small>*</small></font>")
  END IF  %>
                
</TD>

<TD ALIGN="Center" vAlign="top"><a href="/<%=ThisFileName%>?NSL=<% IF Request("Pvar") = "NSL" THEN Response.Write("1") ELSE Response.Write("0")%>&sMemberID=<%=tMemberID%><% IF sRunByWhat = "NCWSA" THEN Response.Write ("&div=" & DivSelected) %>&event=<%=EventSelected%>&pvar=ByMember"><font size=<% =fontsize2 %> face=<% =font2 %>>&nbsp;<% =tMemberID %></FONT></a></TD>
<TD ALIGN="Center" vAlign="top"><font size=<% =fontsize2 %> face=<% =font2 %> COlOR="<%=TextColor1%>">&nbsp;<% =tState %>&nbsp</FONT></TD><%

IF Request("PVar") = "NCWSA" THEN
	%><TD ALIGN="Center" vAlign="top"><font size=<% =fontsize2 %> face=<% =font2 %> COlOR="<%=TextColor1%>">&nbsp;<% =tTeam %>&nbsp</FONT></TD><%
ELSE			
	%><TD ALIGN="Center" vAlign="top"><font size=<% =fontsize2 %> face=<% =font2 %> COlOR="<%=TextColor1%>">&nbsp;<% =tRegion %>&nbsp</FONT></TD><%
END IF

IF Request("PVar") <> "NSL" AND Request("PVar") <> "NCWSA" AND EventSelected <> "O" THEN %>               
	<TD ALIGN="Center" vAlign="top"><font size=<% =fontsize2 %> face=<% =font2 %> COlOR="<%=TextColor1%>"><%

	' Since we coalesced the value of 999 in place of nulls, we have to pull
	' the 999 out during the display process.
		                
	  IF tRegPlace <> "999" AND tRegPlace <> "0" AND tRegPlace <> "" THEN 
		Response.Write (tRegPlace)
                   
		IF ucase(tRegSki) <> tRegion THEN
                	Response.Write ("<font color=darkgreen> (" & ucase(tRegSki) & ")</font>")
	  	END IF                  
          ELSE
		Response.Write ("&nbsp;")
	  END IF %>
	</FONT></TD>
	<TD ALIGN="Center" vAlign="top"><font size=<% =fontsize2 %> face=<% =font2 %> COlOR="<%=TextColor1%>"><%
                 

	IF tNatPlace <> "999" AND tNatPlace <> "0" AND tNatPlace <> "" THEN
		Response.Write (tNatPlace)
	ELSE
       		Response.Write ("&nbsp;")
	END IF 

END IF %> 

<TD ALIGN="Center" vAlign="top"><font size=<% =fontsize2 %> face=<% =font2 %> COlOR="<%=TextColor1%>">&nbsp;<% =tMemberFed %></FONT></TD><%


' -----------  Temporary for displaying LEVELS during testing  MAIN SECTION OF RECORDS --------------	 

IF Request("PVar") <> "NSL" AND Request("PVar") <> "NCWSA" THEN
	%><TD ALIGN="Center" vAlign="top"><font size=<% =fontsize2 %> face=<% =font2 %> COlOR="<%=TextColor1%>">&nbsp;<% =formatnumber(tLevel,3) %></FONT></TD><%

	IF  1=2 AND EventSelected <> "O" AND Request("pvar") <> "NCWSA" AND Request("pvar") <> "NSL" AND RunByWhat <> "NSL" AND RunByWhat<>"NCWSA" THEN
		IF tRGENMemberID <> "" THEN
			 %><TD align=center><a title="Entered in 2007 Goode National Championships"><img src="/images/tools/yellowcheck12.jpg"></a></td><%
		ELSE
			 %><TD>&nbsp;</td><%
		END IF
	END IF	

END IF 

IF sAdminMenuLevel >= "40" THEN 
	%><TD ALIGN="Center" vAlign="top"><font size=<% =fontsize2 %> face=<% =font2 %> COlOR="<%=TextColor1%>">&nbsp;<% =tBirthday %></FONT></TD><%
END IF %>	

</TR><%


END SUB






' -----------------------------
   SUB DefineRankingDataLine
' -----------------------------


tName = TRIM(rs("LastName")) & ", " & rs("FirstName")
tRankScore = rs("Score")
tMemberID = rs("MemberID")
tState = rs("state")
tRegion = rs("Region")
tMemberFed = rs("federationcode")
tTeam = rs("Team")
tBirthday = rs("Birthdate")

IF EventSelected <> "O" THEN   
	tPenalty = TRIM(rs("asc1") & rs("asc2"))    
       	tRating = rs("AWSA_Rat")
       	tNatPlace = TRIM(rs("natl_plc"))
       	tRegSki = rs("reg_ski")
       	tRegPlace = TRIM(rs("regl_plc"))
END IF

SELECT CASE tRegion
	CASE 1
		tRegion = "C"
	CASE 2
		tRegion = "M"
	CASE 3
		tRegion = "W"
	CASE 4
		tRegion = "S"
	CASE 5
		tRegion = "E"
END SELECT




IF sRunByWhat <> "NSL" THEN tRnkScoBkup = rs("RnkScoBkup"): ELSE tRnkScoBkup = ""

IF Request("ZBSAdjustToOldStyle") = "on" AND EventSelected = "S" THEN
	tRankScore = tRankScore - rs("ZBSConversion")
END IF       

IF EventSelected = "S" THEN
	tFmtScore = FormatNumber(tRankScore,2)
ELSE
	IF EventSelected = "O" or EventSelected = "J" THEN
		tFmtScore = FormatNumber(tRankScore,1)
	ELSE
		tFmtScore = FormatNumber(tRankScore,0)
	END IF
END IF
       
' --- Used to allow display of checkmark for indicating entered in Nationals ---
IF EventSelected <> "O" AND Request("pvar") <>"NCWSA" AND Request("pvar") <> "NSL" AND RunByWhat <> "NSL" AND RunByWhat<>"NCWSA" THEN
	tRGENMemberID = rs("RGENMemberID")	 
ELSE
	tRGENMemberID = ""
END IF


' --- Sets levels from DivisionTableName      
tLevel=0
IF rs("Rank_Level")<>0 THEN		
	tLevel=Cdbl(rs("Rank_Level"))
END IF



' ---------------------   WHAT IS THIS - RnkScoBkup  ---------------------------------

IF EventSelected = "O" THEN
      	tRnkScoBkup = rs("RnkScoBkup")
END IF  



' --- Establishes background color for the current record
IF tLevel >= 0  AND tLevel <= tPerc1 THEN 
        DefineRowColor = "<TR bgcolor=" & tcolor10 &">"
ELSEIF  tLevel > tPerc1 AND tLevel <= tPerc2  THEN 
	DefineRowColor = "<TR bgcolor=" & tcolor10 &">"
ELSEIF  tLevel > tPerc2 AND tLevel <= tPerc3  THEN 
        DefineRowColor = "<TR bgcolor=" & tcolor09 &">"
ELSEIF  tLevel > tPerc3 AND tLevel <= tPerc4  THEN 
	DefineRowColor = "<TR bgcolor=" & tcolor08 &">"
ELSEIF  tLevel > tPerc4 AND tLevel <= tPerc5  THEN 
	DefineRowColor = "<TR bgcolor=" & tcolor07 &">"
ELSEIF  tLevel > tPerc5 AND tLevel <= tPerc6  THEN 
	DefineRowColor = "<TR bgcolor=" & tcolor06 &">"
ELSEIF  tLevel > tPerc6 AND tLevel <= tPerc7  THEN 
	DefineRowColor = "<TR bgcolor=" & tcolor05 &">"
ELSEIF  tLevel > tPerc7 AND tLevel <= tPerc8  THEN 
	DefineRowColor = "<TR bgcolor=" & tcolor04 &">"
ELSEIF  tLevel > tPerc8 AND tLevel <= tPerc9  THEN 
	DefineRowColor = "<TR bgcolor=" & tcolor03 &">"
ELSEIF  tLevel > tPerc9 AND tLevel <= tPerc10  THEN 
	DefineRowColor = "<TR bgcolor=" & tcolor02 &">"
ELSE
	DefineRowColor = "<TR>"
END IF



END SUB








' -------------------
   Sub WriteHeader
' -------------------

' Runs As First Subroutine
'
%>
<HTML>
<HEAD><TITLE>TRA Report Viewer</TITLE>
</HEAD>

<BODY BGCOLOR="#FFFFFE" Text="#0A0D0A" LINK="#375AE2" VLINK="#36566D" ALINK="#3E85BB">
<style TYPE="text/css">
<!--  A:link {text-decoration: none; color:#375AE2}  A:visited {text-decoration: none; color:#375AE2}  A:active {text-decoration: none}   A:hover {text-decoration: ; color:#3E85BB; }-->
</style>
<%
END SUB


' -----------------
   Sub WriteFooter
' -----------------

' Runs As Last Subroutine
'

%>
<br><br>
<hr>
</BODY>
</HTML>
<%
END SUB



Sub ChoosePagesSQL(sSQL,sStart, sSize)
  SET rs=Server.CreateObject("ADODB.recordset")

' WriteDebugSQL(sSQL)

  sqlstmt = sSQL
  rs.CursorType = 3
'  rs.PageSize = cint(sSize)
  rs.open sqlstmt, SConnectionToTRATable
'  IF isrecordsetempty = false THEN
'    rs.AbsolutePage = cINT(sStart)
'  END IF
END SUB



Function IsRecordSetEmpty

IF rs.bof = true AND rs.eof = true THEN
    IsRecordSetEmpty = true
ELSE
    IsRecordSetEmpty = false
END IF
end Function



Sub WriteLink(sParms,sDisplay,sBreak)
%>
<A HREF="<% Response.Write(ThisPage & sParms) %>"><% Response.Write(sDisplay) %></A><% Response.Write(sBreak) %>
<%
END SUB


Sub DoCount(currentPage) 
h = 0

for i = 1 to rs.PageCount
 Response.Write(" <a href=" & chr(34) & ThisPage & "?div=" & DivSelected & "&RecordNum=" & RecordNum & "&event=" & EventSelected & "&currentpage=" &  i  & "&action=" & sAction & chr(34) & ">" & i & "</a>")
h = h +1
NEXT
IF h = 0 THEN h = 1
Response.Write("<BR><Small>Page " & currentPage & " of  " & h & "</SMALL></center><BR><BR>")
END SUB

%>