<!--#include file="settingsHQ.asp"-->
<!--#include virtual="/rankings/Tools_Include16_Testing.asp"-->
<!--#include virtual="/rankings/Tools_Registration16.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Tournament Search</title>
<%


'response.write("AML="&Session("AdminMenuLevel"))

Dim Date1Good, Date2Good
Dim MainImage
Dim currentPage, rowCount, i, Monthspan
Dim sMonth, sTourSportsGroup, sTourLevel
Dim sl_check, wb_check, kb_check, bf_check, adminmenulevel
Dim StartMonth, EndMonth
Dim process, OLRButtonEntryStatus, DisableOLRButtons
'Dim sExclude
Dim sTourID

Dim sAllowRegistrationsCheck
Dim sSL_Offered, sTR_Offered, sJU_Offered
Dim TNameWidth
Dim ThisFileName, sShowSQL, OpenNewForm, thisaction
Dim sTourDate1, sTourDate2
' sTourRange, sTourRegion, sTourState

' --- Used in setting images, etc ---
Dim sl, tr, ju, wb, ws, wu, bf, kb, hy, da, jd, ad
Dim s_greenflag, s_yellowflag, s_redflag, FlagMessage, ThisFlag


' -------------------------------------------------------------
' --- Name the file temporarily here for editing/debugging ---
' -------------------------------------------------------------

'ThisFileName="/rankings/View-TournamentsHQ_07032011.asp"
' ThisFileName="/rankings/View-TournamentsHQ16.asp"
ThisFileName="/rankings/View-TournamentsHQ_Test.asp"

' ---------------------------------------------------------------------------------------

Dim objFSO
Set objFSO = Server.CreateObject("Scripting.FileSystemObject")



'-------------------------------------------------------------------------------
' --- Displays the variable values for debugging for TestTour1 and TestTour2
'-------------------------------------------------------------------------------

Display_OLR_variables_Yes="N"
Display_Listing_Variables_Yes="N"
TestTour1="11M128"
TestTour2="11S105"

'-------------------------------------------------------------------------------

'-------------------------------------------------------------------------------




' ---------------------------------------------------------------------------------------
' --- Prevents the ENTER ONLINE registration button from appearing unless set to "on" ---
' ---------------------------------------------------------------------------------------

sAllowRegistrationsCheck="on"
'sAllowRegistrationsCheck="off"

IF Session("AdminMenuLevel")>=50 THEN
  DisableOLRButtons=false
ELSE
  'DisableOLRButtons=true
  DisableOLRButtons=false
END IF

MarkDisableOLR=0
IF MarkDisableOLR=1 THEN DisableOLRButtons=true



' -------------------------------------------------------------------------
' --- Defines the CSS classes so tables and other objects work properly ---
' -------------------------------------------------------------------------

DefineTRAStyles

ScorePageBorderDark = HQSiteColor1
ScorePageBorderLight = HQSiteColor2
ScorePageBorder = HQSiteColor2
NewsPageNum="FAQ_Tours"




' -----------------------------------------------------
' --- Performs requests for all variables 
' -----------------------------------------------------

TourID=TRIM(Request("TourID"))

sTourSportGroup=Request("sTourSportGroup")
sTourRange = TRIM(Request("sTourRange"))
' --- If resulting from link on tournament listing --
IF Request("rg")<>"" THEN sTourRange=Request("rg")

sTourState = TRIM(Request("State"))
sTourDate1 = TRIM(Request("Tour_Date1"))
sTourDate2 = TRIM(Request("Tour_Date2"))
sTourRegion = TRIM(Request("Region"))
sClass=TRIM(Request("sClass"))
StartMonth = TRIM(Request("StartMonth"))	
EndMonth = TRIM(Request("EndMonth"))
process=TRIM(Request("process"))

pvar=Request("pvar")
thisaction=TRIM(Request("thisaction"))
IF thisaction="Update Search" THEN pvar=""
OpenNewForm=TRIM(Request("OpenNewForm"))

sShowSQL = Request("sShowSQL")
adminmenulevel=Session("adminmenulevel")

IF TRIM(Request("SkiYear")) <> "" THEN Session("SkiYear") = TRIM(Request("SkiYear"))



' -----------------------------------------
' --- Runs Dave Clark's traffic counter ---
' -----------------------------------------

IF TRIM(Session("NewTourVis"))="" THEN
	KickTrafficCounter("NewTourVis")	
	Session("NewTourVis")="YES"
END IF

KickTrafficCounter("NewTourPgs")	




' -------------------------------------------------------------------------------------
' --- Request and set all the parameters needed for setting header background image ---
' -------------------------------------------------------------------------------------

RequestBackGroundImageParameters 

' --- Runs SUBROUTINE in tools_include.asp to define the header background image ---
SetEventImage






' -----------------------------------
' --- Opens connection to servers ---
' -----------------------------------

OpenCon
    


' -----------------------------------------------------------------
' --- Check for which Date Range to use for Drop Down and Query ---
' -----------------------------------------------------------------

ValidateDateRangeFromDropDown




' ---------------------------------------------
' --- Writes header portion of HQ main page ---
' ---------------------------------------------

		
IF Thisaction<>"Print Results" THEN 

		WriteIndexPageHeader
END IF

' ----------------------------------------------
' --- Displays box with image and drop downs ---
' ----------------------------------------------
DisplayTourDropdowns




' -------------------------------------------------------------------------------------
' --- Determines program branching - either listing or details on single tournament ---
' -------------------------------------------------------------------------------------

Dim InRange

' --- Changed 4-20-2013 to eliminate CInt error generated when the value included invalid characters ---
IF sTourRange="1" OR sTourRange="2" OR sTourRange="3" OR sTourRange="4" OR sTourRange="5" OR sTourRange="6" OR sTourRange="7" THEN InRange=true
'IF CInt(sTourRange)>=0 AND CInt(sTourRange)<=7 THEN InRange=true 

IF InRange=true AND pvar<> "TourInfo" THEN

	' -----------------------------------------------------------------------------------------
	' --- Runs the appropriate SQL query depending on the range selected from the drop-down ---
	' -----------------------------------------------------------------------------------------

  'set rs=Server.CreateObject("ADODB.recordset")

	IF sTourRange = "0" OR sTourRange = "1" OR sTourRange = "2" OR sTourRange = "3" OR sTourRange = "5" OR sTourRange = "6" THEN
			' --- Executes query for displaying tournaments ---
			PerformSQLQuery_2010
	ELSEIF sTourRange = "4" OR sTourRange = "7" THEN
			' --- Executes query for displaying tournaments ---
			'PerformSQLQuery_Pre2009
			PerformSQLQuery_2010
	ELSE
			InRange=false
	END IF




'response.write("thisaction: "&thisaction)

Dim TNameWidthHead, ThisTableWidth

DataHeaderColor="#FFFFFF"
colcount=0
IF sl="on" THEN colcount=colcount+1
IF tr="on" THEN colcount=colcount+1
IF ju="on" THEN colcount=colcount+1
IF wb="on" THEN colcount=colcount+1
IF ws="on" THEN colcount=colcount+1
IF wu="on" THEN colcount=colcount+1
IF bf="on" THEN colcount=colcount+1
IF kb="on" THEN colcount=colcount+1
IF hy="on" THEN colcount=colcount+1
IF da="on" THEN colcount=colcount+1
IF jd="on" THEN colcount=colcount+1
IF ad="on" THEN colcount=colcount+1


TNameWidth=450-INT(15*colcount)
TNameWidthHead=TNameWidth+5
Monthspan=3+colcount

ThisTableWidth=TourTableWidth+8


	' ----------------------------------------
	' --- Begin Display of Tournament List ---
	' ----------------------------------------

	IF InRange = false THEN
			Display_NoRecords_ErrorMessage

	ELSEIF rs.eof THEN 	
			Display_NoRecords_ErrorMessage

	ELSE 	
			' --- Displays the column headings above the data table ---
			DisplayTableHeader 
			%>
			<TABLE class="innertable" align="center" WIDTH=<%=ThisTableWidth%>px style="padding:2px; border-collapse:collapse; border:1px solid <%=HQSiteColor2%>;">
			<%
			Dim TourCount, OLRCount
			TourCount=0          
			OLRCount=0
			DO WHILE NOT rs.EOF 

					' ------------------------------------------------
					' --- Displays Listing Variables for dedugging ---
					' ------------------------------------------------
					IF adminmenulevel >= 50 AND Display_Listing_Variables_Yes="Y" AND ( rs("TournAppID")=TestTour1 OR rs("TournAppID")=TestTour2 ) THEN
							response.write("<br>*** Listing Variables ***")
							response.write("<br>TournAppID="&rs("TournAppID"))
							response.write("<br>TDateE="&rs("TDateE"))
							response.write("<br>STRegion="&rs("STRegion"))
							response.write("<br>SptsGrpID="&rs("SptsGrpID"))
							response.write("<br>Pending="&rs("Pending"))
							response.write("<br>ShowPSched="&rs("ShowPSched"))
							response.write("<br>ShowRegistrar="&rs("ShowRegistrar")) 
							response.write("<br>GBPolicy="&rs("GBPolicy"))
							response.write("<br>TKitOKGuideBookAd="&rs("TKitOKGuideBookAd"))
							response.write("<br>OK2Publish="&rs("OK2Publish"))
					END IF

					' --------------------------------------------------------------------------
					' --- Determines all the conditions where a data line could be displayed ---
					' --------------------------------------------------------------------------
					DispDataLineYorN = "N"
					IF adminmenulevel > 19 THEN
							DispDataLineYorN = "Y"
					ELSEIF (rs("SptsGrpID")="AWS" AND rs("STRegion")="D") THEN
							DispDataLineYorN = "Y"				
					ELSEIF (rs("SptsGrpID")="ABC" AND rs("STRegion")="B") THEN
							DispDataLineYorN = "Y"			
					ELSEIF (rs("SptsGrpID")="USH" AND rs("Pending")=0) THEN
							DispDataLineYorN = "Y"			
					ELSEIF (rs("Pending") = 0 AND rs("ShowPSched") <> 0 AND rs("ShowRegistrar") <> 0 AND ( (rs("TSTATUS") > 0 AND rs("TSTATUS") <> 3) OR ( rs("GBPolicy") <> 0 AND (rs("TKitOKGuideBookAd") OR rs("OK2Publish")) <> 0) ) ) THEN
							DispDataLineYorN = "Y"	
					ELSEIF rs("TSantype") = "6" AND rs("OK2Publish") = true THEN
							DispDataLineYorN = "Y"
					END IF
			
			
					'IF ( LEFT(rs("TournAppID"),6)="14M085" OR LEFT(rs("TournAppID"),6)="14M086" OR LEFT(rs("TournAppID"),6)="15M029" ) AND Session("adminmenulevel") > 19 THEN
					IF  LEFT(rs("TournAppID"),6)="999999" AND Session("adminmenulevel") > 19 THEN
							response.write("<br>DispDataLineYorN = "&DispDataLineYorN)
							response.write("<br>rs(TSantype) = "&rs("TSantype"))
							response.write("<br>rs(TSantype) = "&rs("TSantype"))
							response.write("<br>rs(OK2Publish) = "&rs("OK2Publish"))
							response.write("<br>rs(SptsGrpID) = "&rs("SptsGrpID"))
							response.write("<br>rs(WWakeW) = "&rs("WWakeW"))
							response.write("<br>rs(Gr2USW_WPulls) = "&rs("Gr2USW_WPulls"))
							response.write("<br>rs(Gr2USW_SurfPulls) = "&rs("Gr2USW_SurfPulls"))
							response.write("<br>Gr2AWS_SPulls = "&rs("Gr2AWS_SPulls")) 
							response.write("<br>Gr1AWSPulls = "&rs("Gr1AWSPulls")) 
							response.write("<br>THSClassN = "&rs("THSClassN")) 
							'response.write("<br>THSClassF = "&rs("THSClassF")) 
							response.write("<br>TEventF3ev = "&rs("TEventF3ev")) 
							response.write("<br>Gr2AWS_TPulls = "&rs("Gr2AWS_TPulls")) 
							response.write("<br>THTClassN = "&rs("THTClassN")) 
							'response.write("<br>THTClassF = "&rs("THTClassF")) 
							response.write("<br>THJClassN = "&rs("THJClassN"))
					END IF	
					
					' ---------------------------------------------------
					' --- Displays as single line of data for listing ---
					' ---------------------------------------------------

					IF DispDataLineYorN = "Y" THEN
							TourCount=TourCount+1
							DisplayDataLine
					END IF

					rs.MoveNext 
			LOOP 

		
			' ---------------------------------------------------------------------
			' --- Displays the total records meeting criteria at bottom of page --- 
			' ---------------------------------------------------------------------	
			DisplaySummaryLine  
			
			%>
			</TABLE>
			<%

			rs.close
			set rs = nothing

	END IF  ' --- Bottom of if NOT eof portion of condition




ELSEIF pvar="TourInfo" THEN

		' ----------------------------------------------
		' --- Displays the DETAILS of the tournament
		' ----------------------------------------------
		DisplaySingleListing	

ELSE
	
		Display_NoRecords_ErrorMessage

END IF     



set rsSelectFields = nothing
CloseCon

' ---------------------------------------------
' --- Writes footer portion of HQ main page ---
' ---------------------------------------------
IF thisaction<>"Print Results" THEN WriteIndexPageFooter








' ----------------------------------------------------------------------------------------------------------------
' ------------------        END OF MAIN PROGRAM       ------------------------------------------------------------
' ----------------------------------------------------------------------------------------------------------------


' ****************************************
  SUB Display_NoRecords_ErrorMessage
' ****************************************

			%>
			<TABLE style="border-width:0px" align=center CELLPADDING="1" CELLSPACING="0" width="<%=TourTableWidth%>px">
				<tr>
					<td align=center>
						<br><br><br>
						<font color="<% =textcolor3 %>" size="<% =fontsize3 %>"><b>No Records Found <br>Please Modify Your Selections and Press Update Search</b></FONT>
					</td>
				</tr>
			</TABLE>
			<%


END SUB



' *************************************
  SUB ValidateDateRangeFromDropDown
' *************************************

Date1Good = 0
Date2Good = 0
sMonth = 0 
    
' -------------------------------------------------
' --- Checks if END Date of TourDate is valid	---
' -------------------------------------------------

IF (isnumeric(left(sTourDate1,2)) And isnumeric(right(left(sTourDate1,5),2)) And isnumeric(right(sTourDate1,4)) And right(left(sTourDate1,3),1) = "/" And right(left(sTourDate1,6),1) = "/" And isDate(sTourDate1)) Or (sTourDate1 = "") THEN
	Date1Good = 1
ELSE
	Date1Good = 0
END IF

' -------------------------------------------------
' --- Checks if START Date of TourDate is valid	---
' -------------------------------------------------
IF (isnumeric(left(sTourDate2,2)) And isnumeric(right(left(sTourDate2,5),2)) And isnumeric(right(sTourDate2,4)) And right(left(sTourDate2,3),1) = "/" And right(left(sTourDate2,6),1) = "/" And isDate(sTourDate2)) Or (sTourDate2 = "") THEN
	Date2Good = 1
ELSE
	Date2Good = 0
END IF
    
IF (Date1Good = 0 or Date2Good = 0) or (sTourDate1 = "" and sTourDate2 = "" and sTourRange = "" and sTourRegion = "" and sTourState = "") THEN
       	sTourRange = "1"   
END IF


END SUB
 




' *****************************************
   SUB RequestBackGroundImageParameters 
' *****************************************

' ---------------------------------------------------
' --- This is temporary to change the links over
' ---------------------------------------------------
evt=Request("evt")


s_greenflag = "images\buttons\Flag-green16.png"
s_yellowflag = "images\buttons\Flag-yellow16.png"
s_redflag = "images\buttons\Flag-red16.png"



	IF evt="on" THEN
		sl="on"
		tr="on"
		ju="on"
		wb="on"
		ws="on"
		wu="on"
		bf="on"
		kb="on"
		hy="on"
		da="on"
		jd="on"
		ad="on"
	ELSE

		sl=Request("sl")
		tr=Request("tr")
		ju=Request("ju")
		wb=Request("wb")
		ws=Request("ws")
		wu=Request("wu")
		bf=Request("bf")
		kb=Request("kb")
		hy=Request("hy")
		IF TRIM(hy)="" THEN hy=Request("hf")
		da=Request("da")
		jd=Request("jd")
		ad=Request("ad")
	END IF
'END IF

IF bf="on" THEN sSptsGrpID="ABC"
IF kb="on" THEN sSptsGrpID="AKA"


sTourLevel=LCASE(Request("sTourLevel"))
IF sTourLevel = "" THEN sTourLevel = "all"


IF sTourLevel = "clinic" THEN 
	sl=""
	tr=""
	ju=""
	wb=""
	ws=""
	wu=""
	bf=""
	hy=""
	kb=""
	da=""
ELSEIF sTourLevel = "cash" THEN 
	ad=""
	jd=""
	wb=""
	ws=""
	wu=""
	bf=""
	hy=""
	kb=""
	da=""
ELSEIF sTourLevel = "premier" THEN 
	ad=""
	jd=""
ELSEIF sTourLevel = "grass" THEN 
	ad=""
	jd=""
	ju=""
ELSEIF sTourLevel = "collegiate" THEN 
	ad=""
	jd=""
	bf=""
	hy=""
	kb=""
END IF


END SUB

' --------------------------------------------------------------------------------------------------------------
    SUB DisplayTourDropdowns  	' ---- Displays all the dropdowns background image and page title
' --------------------------------------------------------------------------------------------------------------


TitleColor=TextColor2
' TourTableWidth

%>
<TABLE class="droptable" align=center width=<%=TourTableWidth%>px height=215 background="<%=MainImage%>"><% '---Table to hold image --- %>
<tr>
<td>
<TABLE width=<%=TourTableWidth%> align=center CELLPADDING="4" CELLSPACING="1"> 

   <form action="<%=ThisFileName%>?rid=<%=rid%>" method="post">
   <input type="hidden" name="TourID" value="<%=TourID%>">
   <input type="hidden" name="pvar" value="<%=pvar%>">


<tr>
  <td colspan=7 align="left">
	<FONT size=4 COlOR=<% =TitleColor %>><b><i>&nbsp;Tournament Search</i></b></font>
  </td>
</tr><%

' /table
' table width="TourTableWidth"
%>
<tr><%
  
     IF sTourLevel="grass" THEN %>
      <td colspan=2 valign=top width=80 align="left">
	<input type=checkbox name=sl <% IF sl="on" THEN response.write "checked" %>>
	<FONT size=<% =fontsize2 %> COlOR=<% =textcolor2 %>><b>Slalom</b></font>
	<br>
	<input type=checkbox name=tr <% IF tr="on" THEN response.write "checked" %>>
	<FONT size=<% =fontsize2 %> COlOR=<% =textcolor2 %>><b>Tricks</b></font>
      </td>	
      <td colspan=2 valign=top width=110 align="left">
	<input type=checkbox name=wb <% IF wb="on" THEN response.write "checked" %>>
	<FONT size=<% =fontsize2 %> COlOR=<% =textcolor2 %>><b>Wakeboard</b></font>
	<br>
	<input type=checkbox name=ws <% IF ws="on" THEN response.write "checked" %>>
	<FONT size=<% =fontsize2 %> COlOR=<% =textcolor2 %>><b>Wake Skate</b></font>
	<br>
	<input type=checkbox name=wu <% IF wu="on" THEN response.write "checked" %>>
	<FONT size=<% =fontsize2 %> COlOR=<% =textcolor2 %>><b>Wake Surfing</b></font>
      </td>	
      <td colspan=2 valign=top width=100 align="left">
	<input type=checkbox name=bf <% IF bf="on" THEN response.write "checked" %>>
	<FONT size=<% =fontsize2 %> COlOR=<% =textcolor2 %>><b>Barefooting</b></font>
	<br>
	<input type=checkbox name=kb <% IF kb="on" THEN response.write "checked" %>>
	<FONT size=<% =fontsize2 %> COlOR=<% =textcolor2 %>><b>Kneeboard</b></font>
	<br>
	<input type=checkbox name=hy <% IF hy="on" THEN response.write "checked" %>>
	<FONT size=<% =fontsize2 %> COlOR=<% =textcolor2 %>><b>Hydrofoil</b></font>
      </td>
      <td colspan=1 valign=top width=400 align="left">&nbsp;</td><%

     ELSEIF sTourLevel="premier" OR sTourLevel="all" THEN  %>
      <td colspan=2 valign=top width=80 align="left">
	<input type=checkbox name=sl <% IF sl="on" THEN response.write "checked" %>>
	<FONT size=<% =fontsize2 %> COlOR=<% =textcolor2 %>><b>Slalom</b></font>
	<br>
	<input type=checkbox name=tr <% IF tr="on" THEN response.write "checked" %>>
	<FONT size=<% =fontsize2 %> COlOR=<% =textcolor2 %>><b>Tricks</b></font>
	<br>
	<input type=checkbox name=ju <% IF ju="on" THEN response.write "checked" %>>
	<FONT size=<% =fontsize2 %> COlOR=<% =textcolor2 %>><b>Jump</b></font>
      </td>	

      <td colspan=2 valign=top width=110 align="left">
	<input type=checkbox name=wb <% IF wb="on" THEN response.write "checked" %>>
	<FONT size=<% =fontsize2 %> COlOR=<% =textcolor2 %>><b>Wakeboard</b></font>
	<br>
	<input type=checkbox name=ws <% IF ws="on" THEN response.write "checked" %>>
	<FONT size=<% =fontsize2 %> COlOR=<% =textcolor2 %>><b>Wake Skate</b></font>
	<br>
	<input type=checkbox name=wu <% IF wu="on" THEN response.write "checked" %>>
	<FONT size=<% =fontsize2 %> COlOR=<% =textcolor2 %>><b>Wake Surfing</b></font>
      </td>	

      <td colspan=2 valign=top width=100 align="left">
	<input type=checkbox name=bf <% IF bf="on" THEN response.write "checked" %>>
	<FONT size=<% =fontsize2 %> COlOR=<% =textcolor2 %>><b>Barefooting</b></font>
	<br>
	<input type=checkbox name=kb <% IF kb="on" THEN response.write "checked" %>>
	<FONT size=<% =fontsize2 %> COlOR=<% =textcolor2 %>><b>Kneeboard</b></font>
	<br>
	<input type=checkbox name=hy <% IF hy="on" THEN response.write "checked" %>>
	<FONT size=<% =fontsize2 %> COlOR=<% =textcolor2 %>><b>Hydrofoil</b></font>
      </td>

      <td colspan=1 valign=top width=300 align="left">&nbsp;</td><%

     ELSEIF sTourLevel="cash" THEN  %>
      <td colspan=2 valign=top width=80 align="left">
	<input type=checkbox name=sl <% IF sl="on" THEN response.write "checked" %>>
	<FONT size=<% =fontsize2 %> COlOR=<% =textcolor2 %>><b>Slalom</b></font>
	<br>
	<input type=checkbox name=tr <% IF tr="on" THEN response.write "checked" %>>
	<FONT size=<% =fontsize2 %> COlOR=<% =textcolor2 %>><b>Tricks</b></font>
	<br>
	<input type=checkbox name=ju <% IF ju="on" THEN response.write "checked" %>>
	<FONT size=<% =fontsize2 %> COlOR=<% =textcolor2 %>><b>Jump</b></font>
      </td>	
      <td colspan=1 valign=top width=300 align="left">&nbsp;</td>
      <td colspan=1 valign=top width=300 align="left">&nbsp;</td><%


     ELSEIF sTourLevel="collegiate" THEN  %>

      <td colspan=2 valign=top width=80 align="left">
	<input type=checkbox name=sl <% IF sl="on" THEN response.write "checked" %>>
	<FONT size=<% =fontsize2 %> COlOR=<% =textcolor2 %>><b>Slalom</b></font>
	<br>
	<input type=checkbox name=tr <% IF tr="on" THEN response.write "checked" %>>
	<FONT size=<% =fontsize2 %> COlOR=<% =textcolor2 %>><b>Tricks</b></font>
	<br>
	<input type=checkbox name=ju <% IF ju="on" THEN response.write "checked" %>>
	<FONT size=<% =fontsize2 %> COlOR=<% =textcolor2 %>><b>Jump</b></font>
      </td>

      <td colspan=2 valign=top width=110 align="left">
	<input type=checkbox name=wb <% IF wb="on" THEN response.write "checked" %>>
	<FONT size=<% =fontsize2 %> COlOR=<% =textcolor2 %>><b>Wakeboard</b></font>
	<br>
	<input type=checkbox name=ws <% IF ws="on" THEN response.write "checked" %>>
	<FONT size=<% =fontsize2 %> COlOR=<% =textcolor2 %>><b>Wake Skate</b></font>
	<br>
	<input type=checkbox name=wu <% IF wu="on" THEN response.write "checked" %>>
	<FONT size=<% =fontsize2 %> COlOR=<% =textcolor2 %>><b>Wake Surfing</b></font>
      </td>	

      <td colspan=1 valign=top width=100 align="left">&nbsp;</td>
      <td colspan=1 valign=top width=100 align="left">&nbsp;</td><%

     ELSEIF sTourLevel="clinic" THEN  %>

      <td colspan=2 valign=top width=200 align="left">
	<input type=checkbox name=jd <% IF jd="on" THEN response.write "checked" %>>
	<FONT size=<% =fontsize2 %> COlOR=<% =textcolor2 %>><b>Junior Development</b></font>
        <br>
	<input type=checkbox name=ad <% IF ad="on" THEN response.write "checked" %>>
	<FONT size=<% =fontsize2 %> COlOR=<% =textcolor2 %>><b>Athlete Development</b></font>
	<br>
      </td>	

      <td colspan=4 valign=top width=150 align="left">&nbsp;</td><%	

     END IF %>	

  </td>

</tr>
</table>



<table WIDTH=100%>

<tr>
  <td colspan=1 width=60 valign=top align="right">
	<font color="<% =TitleColor %>" size="<% =fontsize2 %>"><b>Type:</b></font>
  </td>
  <td colspan=2 width=130 align="left">
      <select name="sTourLevel" onchange=submit()>
            <option value="premier" <%IF sTourLevel = "premier" THEN Response.Write(" selected ")%>>Premier</option>
            <option value="grass" <%IF sTourLevel = "grass" THEN Response.Write(" selected ")%>>GrassRoots</option>
            <option value="all" <%IF sTourLevel = "all" THEN Response.Write(" selected ")%>>Premier & GrassRoots</option>
            <option value="collegiate" <%IF sTourLevel = "collegiate" THEN Response.Write(" selected ")%>>Collegiate</option>
            <option value="cash" <%IF sTourLevel = "cash" THEN Response.Write(" selected ")%>>Cash Prize</option>
            <option value="clinic" <%IF sTourLevel = "clinic" THEN Response.Write(" selected ")%>>Clinics</option>
    </select>
  </td>
  <td colspan=1 width=60 valign=top align="right">
	<font color="<% =TitleColor %>" size="<% =fontsize2 %>"><b>Region:</b></font>
  </td>
  <td width=60 colspan=1 valign=top align="left">
    <select name="Region">
  	<option value=""<%IF sTourRegion = "" THEN Response.Write(" selected ")%>>All</option><%
	IF sTourLevel<>"collegiate" THEN %>
		<option value="C"<%IF sTourRegion = "C" THEN Response.Write(" selected ")%>>SC</option>
		<option value="M"<%IF sTourRegion = "M" THEN Response.Write(" selected ")%>>MW</option>
		<option value="W"<%IF sTourRegion = "W" THEN Response.Write(" selected ")%>>WE</option>
		<option value="S"<%IF sTourRegion = "S" THEN Response.Write(" selected ")%>>SO</option>
		<option value="E"<%IF sTourRegion = "E" THEN Response.Write(" selected ")%>>EA</option><%
	END IF  %>
    </select>
  </td>

  <td colspan=2 width=350 valign=top align="left">&nbsp;</td>
</tr>

<tr>
  <td colspan=1 width=60 align="right">
	<font color="<% =TitleColor %>" size="<% =fontsize2 %>"><b>Range:</b></font>
  </td>

  <td colspan=2 align="left">
    <select name='sTourRange'>
  	<option value="0"<%IF sTourRange = "0" THEN Response.Write(" selected ")%>>Custom</option>
  	<option value="1"<%IF sTourRange = "1" THEN Response.Write(" selected ")%>>Future</option><%

        set rsSelectFields=Server.CreateObject("ADODB.recordset")
	rsSelectFields.open "Select Top 3 * FROM "&SkiYearTableName&" WHERE SkiYearID<>'1' ORDER BY BeginDate DESC", SConnectionToTRATable

	IF NOT rsSelectFields.eof THEN %>
	  	<option value="2"<%IF sTourRange = "2" THEN Response.Write(" selected ")%>>Ski Year <%=right(right(TRIM(rsSelectFields("SkiYearName")),4),4)%></option><%
	  	rsSelectFields.movenext 
	  	IF NOT rsSelectFields.eof THEN %>
					<option value="3"<%IF sTourRange = "3" THEN Response.Write(" selected ")%>>Ski Year <%=right(right(TRIM(rsSelectFields("SkiYearName")),4),4)%></option><%
					rsSelectFields.movenext 
					IF NOT rsSelectFields.eof THEN %>
		   				<option value="4"<%IF sTourRange = "4" THEN Response.Write(" selected ")%>>Ski Year <%=right(right(TRIM(rsSelectFields("SkiYearName")),4),4)%></option>
	  	  			<option value="5"<%IF sTourRange = "5" THEN Response.Write(" selected ")%>>Calendar <%=Year(Date())%></option>
	  					<option value="6"<%IF sTourRange = "6" THEN Response.Write(" selected ")%>>Calendar <%=Year(Date())-1%></option><%
					END IF 
	  	END IF	
	ELSE  ' --- Applies only if no SkiYears are found in Ski Year table  %>
			<option value="2"<%IF sTourRange = "2" THEN Response.Write(" selected ")%>>Ski Year <%=Year(Date())%></option>
			<option value="3"<%IF sTourRange = "3" THEN Response.Write(" selected ")%>>Ski Year <%=Year(Date())-1%></option>
			<option value="4"<%IF sTourRange = "4" THEN Response.Write(" selected ")%>>Ski Year <%=Year(Date())-2%></option>
	  	<option value="5"<%IF sTourRange = "5" THEN Response.Write(" selected ")%>>Calendar <%=Year(Date())%></option>
	  	<option value="6"<%IF sTourRange = "6" THEN Response.Write(" selected ")%>>Calendar <%=Year(Date())-1%></option><%
	END IF
	rsSelectFields.close

	'IF adminmenulevel>19 OR sTourRange = "5" THEN 
	IF adminmenulevel>19 THEN %>

	  	<option value="7"<%IF sTourRange = "7" THEN Response.Write(" selected ")%>>Calendar <%=Year(Date())-2%></option><%
	END IF %>	
    </select>
  </td>

  <td colspan=1 valign=top align="right">
        <font color="<% =TitleColor %>" size="<% =fontsize2 %>"><b>State:</b></font> 
  </td>

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
  <td colspan=2 width=350 valign=top align="left">&nbsp;</td>
</tr>

<tr>

  <td colspan=1 valign=top align="right">
	<font color="<% =TitleColor %>" size="<% =fontsize2 %>"><b>Start/End:</b></font>
  </td>

  <td colspan=1 valign=top align="left"><%
    LoadMonthsPulldown "StartMonth", StartMonth %>
  </td>

  <td colspan=1 valign=top align="left"><%
    LoadMonthsPulldown "EndMonth", EndMonth %>
  </td>


  <td colspan=1 valign=top width=60 align="right">
	<font color="<% =TitleColor %>" size="<% =fontsize2 %>"><b>Class:</b></font>
  </td>

  <td colspan=1 valign=top align="left">
      <select name="sClass">
            <option value="All" <%IF sClass = "All" THEN Response.Write(" selected ")%>>All</option>
            <option value="R" <%IF sClass = "R" THEN Response.Write(" selected ")%>>R</option>
            <option value="L" <%IF sClass = "L" THEN Response.Write(" selected ")%>>L</option>
            <option value="E" <%IF sClass = "E" THEN Response.Write(" selected ")%>>E</option>
            <option value="C" <%IF sClass = "C" THEN Response.Write(" selected ")%>>C</option>
            <option value="N" <%IF sClass = "N" THEN Response.Write(" selected ")%>>N</option>
            <option value="F" <%IF sClass = "F" THEN Response.Write(" selected ")%>>F</option>
            <option value="F_O" <%IF sClass = "F_O" THEN Response.Write(" selected ")%>>F W/O</option>
    </select>
  </td><%

	IF AdminMenuLevel>=50 THEN  %>	
  		<td colspan=2 width=350 valign=top align="left">
			<font color="<% =TitleColor %>" size="<% =fontsize2 %>"><b>Show SQL</b></font>
			<input type=checkbox name="sShowSQL" <% IF sShowSQL="on" THEN response.write "checked" %>>

		</td><%
	ELSE  %>
  		<td colspan=2 width=350 valign=top align="left">&nbsp;</td><%
	END IF %>
</tr>

<tr>
  <td colspan=7 valign=top>&nbsp;</td> 
</tr>
<% 'response.write("<br>ThisFileName="&ThisFileName)
%>
<tr>
  <td colspan=2 align="center">
	<input type="hidden" name="process" value="<%=process%>">
	<input type="submit" style="width:9em" name="thisaction" value="Update Search">
  </td>

  <td colspan=2 align="center" width="135px"><%

	IF thisaction="Print Results" THEN  ' ---- Print Tour Listing ---- %>
		<a href='#' onclick='window.print()' title="Click here to Print"><input type=submit name="thisaction" value="Print Screen" style="width:9em; background-color:red; color:white"></a><%		
	ELSE  %>
		<input type="submit" style="width:9em" name="thisaction" value="Print Results"><%
	END IF %>
  </td>
  <td colspan=3 align="left"> 
	<a title="View FAQ for Tournaments" onclick="window.open('/rankings/tools.asp?svar=FAQ&np=FAQ_Tours_Detail', '_blank', 'width=600,height=500');"><input type="submit" style="width:9em" name="thisaction" value="FAQ/Tips"></a>
  </td>
</tr>


  </td>
</tr>
</TABLE>

  </td>
</tr>

</form>

</TABLE><%   ' --- Bottom of Picture Table


END SUB




SUB Ttemp


%> 
	<a href='/rankings/tools.asp?svar=FAQ&np=FAQ_Tours_Detail' onclick='target=_blank' title="View FAQ for Tournaments"><input type="submit" style="width:9em" name="thisaction" value="FAQ/Tips"></a>
<%
	IF pvar= "TourInfo" THEN %>
	    <form action="/rankings/tools.asp?svar=FAQ&np=FAQ_Tours_Detail" method="post" target="_blank"><%
	ELSE %>
	    <form action="/rankings/tools.asp?svar=FAQ&np=FAQ_Tours" method="post" target="_blank"><%
	
		Session("sSendingPage")=ThisFileName&"?sTourLevel="&sTourLevel&"&sl="&sl&"&tr="&tr&"&ju="&ju&"&wb="&wb&"&ws="&ws&"&wu="&wu&"&bf="&bf&"&kb="&kb&"&hy="&hy

	END IF
END SUB




' -----------------------------------------------------------------------
   SUB DisplaySummaryLine  	' Displays total line after data table
' -----------------------------------------------------------------------

Dim OLRPerc
IF TourCount>0 THEN 
		OLRPerc=OLRCount/TourCount 
ELSE 
		OLRPerc=0 
END IF

%>
<tr>
  <td colspan=1>&nbsp;</td>
  <td>
	<font size=<% =fontsize3 %> COlOR=<% =textcolor1 %>>&nbsp;Total All: <%=TourCount%></font>
	<font size=<% =fontsize3 %> COlOR=<% =textcolor1 %>>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;OLR Active: <%=OLRCount%></font>
	<font size=<% =fontsize3 %> COlOR=<% =textcolor1 %>>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Perc: <% Response.write(FormatNumber(OLRPerc*100,1)) %>%</font>  
  </td>
  <td colspan=6>&nbsp;</td>
</tr><%

END SUB



' -----------------------------------------------------------------------
   SUB DisplaySummaryLine_OLD  	' Displays total line after data table
' -----------------------------------------------------------------------

%>
<tr>
  <td>
	<FONT size=<% =fontsize3 %> COlOR=<% =textcolor1 %>>&nbsp;Total:</font>
  </td>
  <td>
	<FONT size=<% =fontsize3 %> COlOR=<% =textcolor1 %>><%=TourCount%></font>
  </td>
  <td>
	<FONT size=<% =fontsize3 %> COlOR=<% =textcolor1 %>>&nbsp;OLR Entry:</font>
  </td>
  <td>
	<FONT size=<% =fontsize3 %> COlOR=<% =textcolor1 %>><%=OLRCount%></font>
  </td>
  <td colspan=3>&nbsp;</td>
</tr><%

END SUB







' -------------------------
    SUB DisplayTableHeader
' -------------------------

'DataHeaderColor="#FFFFFF"
'colcount=0
'IF sl="on" THEN colcount=colcount+1
'IF tr="on" THEN colcount=colcount+1
'IF ju="on" THEN colcount=colcount+1
'IF wb="on" THEN colcount=colcount+1
'IF ws="on" THEN colcount=colcount+1
'IF wu="on" THEN colcount=colcount+1
'IF bf="on" THEN colcount=colcount+1
'IF kb="on" THEN colcount=colcount+1
'IF hy="on" THEN colcount=colcount+1
'IF da="on" THEN colcount=colcount+1
'IF jd="on" THEN colcount=colcount+1
'IF ad="on" THEN colcount=colcount+1


'TNameWidth=450-INT(15*colcount)
'TNameWidthHead=TNameWidth+5
'Monthspan=3+colcount
'response.write("TNameWidth="&TNameWidth)

'ThisTableWidth=TourTableWidth+8
%>
<TABLE class="scores" style="border-collapse:collapse; border:1px solid <%=HQSiteColor2%>" width=<%=ThisTableWidth%>px align=center>

<tr>
  <th align="Center" width=75px><font color="FFFFFF" size="<%=fontsize2%>"><br><b>Date(s)</b></FONT></th>
  <th align="Center" width=<%=TNameWidthHead%>px  valign="top"><font color="FFFFFF" size="<% =fontsize2 %>">
	<br>
	<b>Tournament Name (ID)<br>Event Info</b></FONT>
  </th>
  <th align="Center" width=115px valign="top"><font color="FFFFFF" size="<% =fontsize2 %>"><b>City<br><b>State</b></FONT></Th><%

	    IF sl="on" THEN 
		%><th align="Center" width=40px valign="top"><font color="FFFFFF" size="<% =fontsize2 %>"><br><b><a title="Slalom">SL</b></FONT></th><%
	    END IF
	    IF tr="on" THEN 
		%><th align="Center" width=40px valign="top"><font color="FFFFFF" size="<% =fontsize2 %>"><br><b><a title="Tricks">TR</b></FONT></th><%
	    END IF
	    IF ju="on" THEN 
		%><th align="Center" width=40px valign="top"><font color="FFFFFF" size="<% =fontsize2 %>"><br><b><a title="Jumping">JU</b></FONT></th><%
	    END IF
    	    IF wb="on" THEN 
		%><th align="Center" width=40px valign="top"><font color="FFFFFF" size="<% =fontsize2 %>"><br><b><a title="Wakeboard">WB</b></FONT></th><%
	    END IF
	    IF ws="on" THEN 
		%><th align="Center" width=40px valign="top"><font color="FFFFFF" size="<% =fontsize2 %>"><br><b><a title="Wake Skate">SK</b></FONT></th><%
	    END IF
	    IF wu="on" THEN 
		%><th align="Center" width=40px valign="top"><font color="FFFFFF" size="<% =fontsize2 %>"><br><b><a title="Wake Surf">SU</b></FONT></th><%
	    END IF
	    IF bf="on" THEN 
	    	%><th align="Center" width=40px valign="top"><font color="FFFFFF" size="<% =fontsize2 %>"><br><b><a title="Barefoot">BF</b></FONT></th><%
	    END IF 
	    IF kb="on" THEN 
	    	%><th align="Center" width=40px valign="top"><font color="FFFFFF" size="<% =fontsize2 %>"><br><b><a title="Kneeboard">KB</a></b></FONT></th><%
	    END IF 
	    IF hy="on" THEN 
	    	%><th align="Center" width=40px valign="top"><font color="FFFFFF" size="<% =fontsize2 %>"><br><b><a title="Hydrofoiling">HY</b></FONT></th><%
	    END IF 
	    IF da="on" THEN
	    	%><th align="Center" width=40px valign="top"><font color="FFFFFF" size="<% =fontsize2 %>"><br><b><a title="Disabled">DA</b></FONT></th><%
	    END IF 
	    IF jd="on" THEN 
	    	%><th align="Center" width=40px valign="top"><font color="FFFFFF" size="<% =fontsize2 %>"><br><b><a title="Junior Development">JD</b></FONT></th><%
	    END IF 
	    IF ad="on" THEN 
	    	%><th align="Center" width=40px valign="top"><font color="FFFFFF" size="<% =fontsize2 %>"><br><b><a title="Athlete Development - Define Brandon">AD</b></FONT></th><%
	    END IF %>

  <% ' <th>&nbsp;</th>
  %>
</tr>
</TABLE>

<% 


END SUB





' ---------------------------
    SUB DisplayDataLine
' ---------------------------
	
Dim MonthColor

sTourID = LEFT(rs("TournAppID"),6)
sTDateS = rs("TDateS")
sTDateE = rs("TDateE")

' --- From Display ONE TOURNAMENT SECTION ---
'OLRButtonEntryStatus="enabled"
'EntryButtonTitle="Enter this tournament with our online entry form"

' --- Tournament reached it's entry limit ---
'IF request("olrds") ="disabled" and sTStatus>0 THEN 
'	OLRButtonEntryStatus="disabled"
'	EntryButtonTitle="This tournament is no longer open to Online Registration"

' --- PayPal is setup but tournament has not received approval ---
'ELSEIF sPayPalOK=true AND sTStatus=0 THEN
'	sExclude="no"
'	OLRButtonEntryStatus="disabled"
'	EntryButtonTitle="This tournament is not yet available for Online Registration"
'END IF



' --------------------------------------------------------------------
' --- Define what color to use based on the month so it alternates ---
' --------------------------------------------------------------------
SELECT CASE Month(sTDateS)
		CASE 1,5,9
				MonthColor="#EEDDDD"
		CASE 2,6,10
				MonthColor="#CCCCFF"
		CASE 3,7,11
				MonthColor="#FFFF66"
		CASE 4,8,12
				MonthColor="#CCFFCC"
END SELECT 

IF sTDateS = sTDateE THEN 
		DisplayDate = Month(sTDateS) & "/" & Day(sTDateS) & "/" & RIGHT(cStr(Year(sTDateS)),2)
		'DisplayDate = Month(sTDateS) & "/" & Day(sTDateS) 
ELSE
		DisplayDate = Month(sTDateS) & "/" & Day(sTDateS) & "-" & Day(sTDateE) & "/" & RIGHT(cStr(Year(sTDateS)),2)
		'DisplayDate = Month(sTDateS) & "/" & Day(sTDateS) & "-" & Day(sTDateE)
END IF




	' ------------------------------------------------
	' --- DEFINE ALL VARIABLES USED IN THE DISPLAY ---
	' ------------------------------------------------


	' ---------------------------------------------------------------
	' --- CONTROLS FOR BUTTONS FOR ONLINE ENTRY, ADMIN LOGIN, ETC ---
	' ---------------------------------------------------------------

	' --- TStatus > 1 means the tournament has been sanctioned by HQ
	' --- sAllowRegistrationsCheck="on" is set at the top of this file as an override
 	' --- adminmenulevel is defined by user profile in SWIFT LOGIN
	' --- TestValidAdminCode is a function in Tools_Registration that verifies AdminCode for this tournament

	sTourID = rs("TournAppID")
	sTSanction = rs("TSanction")
	IF left(sTSanction,6) <> sTourID THEN sTSanction = sTourID & "-"
	sTourName=rs("TName")
	sTCity=rs("TCity")
	sTState=rs("TState")

	' -----------------------------------------------------------------------------
	' --- Adds all descriptions into one string with breaks between SptsGrpID's ---
	' -----------------------------------------------------------------------------
	IF TRIM(rs("TDescription"))<>"" THEN ThisDescription = TRIM(rs("TDescription")) + "<br>" 
	IF TRIM(rs("FDescription"))<>"" THEN ThisDescription = ThisDescription + TRIM(rs("FDescription")) + "<br>" 
	IF TRIM(rs("WDescription"))<>"" THEN ThisDescription = ThisDescription + TRIM(rs("WDescription")) + "<br>" 
	IF TRIM(rs("KDescription"))<>"" THEN ThisDescription = ThisDescription + TRIM(rs("KDescription")) + "<br>" 
	IF TRIM(rs("CDescription"))<>"" THEN ThisDescription = ThisDescription + TRIM(rs("CDescription")) + "<br>" 

	sTStatus=rs("TSTATUS")
	ThisFlag=""
	FlagMessage=""
	IF sTStatus > 1 THEN
			ThisFlag=s_greenflag
			FlagMessage="All approvals received"
	ELSEIF sTStatus = 1 THEN
			ThisFlag=s_yellowflag
			FlagMessage="Final sanction approvals pending"			
	ELSEIF sTStatus = 3 THEN
			ThisFlag=s_redflag			 		
			FlagMessage="Tournament cancelled"			
	ELSEIF sTStatus = 0 THEN
			ThisFlag=s_redflag			 		
			FlagMessage="No sanction approvals received"			
	END IF	
	
	sTDeleted=rs("Deleted")   
	sPayPalOK=rs("PayPalOK") 
	sPayPalAct=rs("PayPalAct")
	sUseOLReg=rs("UseOLReg")
	sOLR_PD=rs("OLR_PD")


	sSL_Offered="N"
	sTR_Offered="N"
	sJU_Offered="N"

	sWB_Offered="N"
	sWS_Offered="N"
	sWU_Offered="N"

	sBF_Offered="N"
	sKB_Offered="N"
	sHY_Offered="N"

	sDA_Offered="N"
	sJD_Offered="N"
	sAD_Offered="N"

	' --- Begins in 2010
	IF sTourRange = "0" OR sTourRange = "1" OR sTourRange = "2" OR sTourRange >= "5" THEN
		IF sl="on" AND ( ( rs("sClassC") + rs("sClassE") + rs("sClassL") + rs("sClassR") + rs("sClassCash") + rs("sClassX") > 0 )  OR rs("Gr2AWS_SPulls")<>0 OR rs("Gr1AWSPulls") ) THEN sSL_Offered="Y"
		IF tr="on" AND (  (rs("tClassC") + rs("tClassE") + rs("tClassL") + rs("tClassR") + rs("tClassCash") + rs("tClassX") > 0 )  OR rs("Gr2AWS_TPulls")<>0  ) THEN sTR_Offered ="Y"
		IF ju="on" AND (  (rs("jClassC") + rs("jClassE") + rs("jClassL") + rs("jClassR") + rs("jClassCash") + rs("jClassX") > 0 )  ) THEN sJU_Offered ="Y"

		IF wb="on" AND (  rs("WWakeW")>0 OR rs("Gr2USW_WPulls")<>0 OR rs("Gr2USW_RailJamPulls")<>0 OR rs("Gr1USWPulls") <> 0 ) THEN sWB_Offered="Y"
		IF ws="on" AND rs("WSkateW")>0 OR rs("Gr2USW_SkatePulls") THEN sWS_Offered="Y"
		IF wu="on" AND rs("WSurfW")>0 OR rs("Gr2USW_SurfPulls") THEN sWU_Offered="Y"

		IF bf="on" AND (  rs("SptsGrpID")="ABC" OR rs("Gr1ABCPulls")<>0  ) THEN sBF_Offered="Y" 
		IF kb="on" AND (  rs("SptsGrpID")="AKA" OR rs("Gr2AKA_SPulls")<>0 OR rs("Gr2AKA_TPulls")<>0 OR rs("Gr2AKA_FreePulls")<>0 OR rs("Gr2AKA_FlipPulls")<>0 OR rs("Gr1AKAPulls")<>0  ) THEN sKB_Offered="Y"	
		IF hy="on" AND (  rs("TEventFHF")<>0 OR rs("Gr2USH_FreeRidePulls")<>0 OR rs("Gr2USH_JumpOutPulls")<>0 OR rs("Gr2USH_BigAirPulls")<>0 OR rs("Gr2USH_3TrickPulls")<>0 OR rs("Gr1USHPulls")<>0  ) THEN sHY_Offered="Y"

		IF da="on" AND (rs("Gr1WSDPulls")) THEN sDA_Offered="Y"
		IF jd="on" AND (rs("JDClin")<>0) THEN sJD_Offered="Y"
		IF ad="on" AND (rs("ADClin")<>0) THEN sAD_Offered="Y"

	ELSE
		IF sl="on" AND (rs("TEventSlalom")<>0 OR rs("TEventF3ev")<>0 OR rs("Gr2AWS_SPulls")<>0 OR rs("Gr1AWSPulls")<>0) THEN sSL_Offered ="Y"
		IF tr="on" AND (rs("TEventTrick")<>0 OR rs("Gr2AWS_TPulls")<>0) THEN sTR_Offered ="Y"
		IF ju="on" AND (rs("TEventJump")<>0) THEN sJU_Offered ="Y"

		IF wb="on" AND (rs("TEventWake")<>0 OR rs("TEventFW")<>0 OR rs("Gr2USW_WPulls")<>0 OR rs("Gr2USW_RailJamPulls")<>0 OR rs("Gr1USWPulls")<>0) THEN sWB_Offered="Y"
		IF ws="on" AND (rs("TEventWSkate")<>0 OR rs("Gr2USW_SkatePulls")<>0) THEN sWS_Offered="Y"
		IF wu="on" AND (rs("TEventWSurf")<>0 OR rs("Gr2USW_SurfPulls")<>0) THEN sWU_Offered="Y"

		IF bf="on" AND (rs("SptsGrpID")="ABC" OR rs("TEventNBL")<>0 OR rs("Gr1ABCPulls")<>0) THEN sBF_Offered="Y" 
		IF kb="on" AND (rs("SptsGrpID")="AKA" OR rs("TEventFKB")<>0 OR rs("Gr2AKA_SPulls")<>0 OR rs("Gr2AKA_TPulls")<>0 OR rs("Gr2AKA_FreePulls")<>0 OR rs("Gr2AKA_FlipPulls")<>0 OR rs("Gr1AKAPulls")<>0) THEN sKB_Offered="Y"	
		IF hy="on" AND (rs("TEventFHF")<>0 OR rs("Gr2USH_FreeRidePulls")<>0 OR rs("Gr2USH_JumpOutPulls")<>0 OR rs("Gr2USH_BigAirPulls")<>0 OR rs("Gr2USH_3TrickPulls")<>0 OR rs("Gr1USHPulls")<>0) THEN sHY_Offered="Y"

		IF da="on" AND (rs("Gr1WSDPulls")) THEN sDA_Offered="Y"
		IF jd="on" AND (rs("JDClin")<>0) THEN sJD_Offered="Y"
		IF ad="on" AND (rs("ADClin")<>0) THEN sAD_Offered="Y"
	END IF



	'----------------------------------------------------------------
	' --- Determines whether or not to GREY out OLR button ---
	'----------------------------------------------------------------

	OLRButtonEntryStatus="enabled"
	EntryButtonTitle="Enter this tournament with our online entry form"

	
	' --- Determines if corresponding record exists in Ski Year table - avoids errors until record is entered --
	set rsSkiYear=Server.CreateObject("ADODB.recordset")
	rsSkiYearSQL = "SELECT * FROM "&SkiYearTableName&" WHERE EndDate>='"&sTDateE&"' AND BeginDate<='"&sTDateS&"'" 
	rsSkiYear.open rsSkiYearSQL, SConnectionToTRATable
	
	whiletesting=0	
	
	IF rsSkiYear.eof THEN 
			OLRButtonEntryStatus="disabled"
			EntryButtonTitle="Ski Year Administrative table must be updated for this Ski Year before available for OLR "
	ELSEIF whiletesting=1 AND adminmenulevel < 50 AND RIGHT(cStr(Year(sTDateS)),2)>="16" THEN 
  		RegFileForLink = "registration16.asp"
  		OLRButtonEntryStatus="disabled"
			EntryButtonTitle="Online Registration has not been activated for 2016 - Thanks for your patience"
	ELSEIF NOT(rs("OLRDisplayStatus")) THEN
			sExclude="no"
			OLRButtonEntryStatus="disabled"
			EntryButtonTitle="This tournament is no longer open to Online Registration"
	ELSEIF (rs("UseOLReg")=true AND (rs("PayPalOK")=0 OR rs("OLR_PD")=0 OR TRIM(rs("PayPalAct"))="" ) ) THEN
			sExclude="no"
			OLRButtonEntryStatus="disabled"
			EntryButtonTitle="This tournament is not yet available for Online Registration"
	END IF


	' --- OLR Button status ---
	IF DisableOLRButtons=true THEN
			OLRButtonEntryStatus="disabled"
			EntryButtonTitle="Online Registration is Temporarily Disabled - Thanks for your patience"
	END IF


	' --- Determines which OLR program is used ---  
  IF RIGHT(cStr(Year(sTDateS)),2)>="16" THEN 
  		RegFileForLink = "registration16.asp"
  ELSE
  		RegFileForLink = "registration.asp"
  END IF		


	'---------------------------------------------------------
	' --- For DEBUGGING purposes - set displayvariables_Yes=Y
	'---------------------------------------------------------

	IF adminmenulevel >= 50 AND Display_OLR_variables_Yes="Y" AND ( sTourID=TestTour1 OR sTourID=TestTour2 ) THEN
			response.write("<br>*** OLR Variables ***")
			response.write("<br>TournAppID="&rs("TournAppID"))
			response.write("<br>TStatus="&sTSTATUS)
			response.write("<br>sUseOLReg="&sUseOLReg)
			response.write("<br>PayPalAct="&sPayPalAct)
			response.write("<br>OLRButtonEntryStatus="&OLRButtonEntryStatus)
			response.write("<br><br>")
	END IF

'style="word-wrap:break-word;" 



' -----------------------------------------------------------
' --- Display the colored row each time the month changes ---
' -----------------------------------------------------------
IF sMonth <> Month(sTDateS) THEN 
		sMonth = Month(sTDateS) 
		ThisMonthName = MonthName(MONTH(sTDateS)) & " - " & YEAR(sTDateS) 
		%>
		<tr>
	  	<td align="Left" valign="top" colspan=<%=Monthspan%> style="background-color:<%=MonthColor%>;">
				<font color="<% =textcolor1 %>" size="<% =fontsize3 %>" face="<% =font1 %>"><b>&nbsp;<%=ThisMonthName%></b></Font>
	  	</td>
		</tr>
		<%
END IF 

%> 
<tr>
<%


' ---------------------------------------------
' --- Display the Date(s) of the tournament ---
' ---------------------------------------------

	%>
  <td width=70px align="center" style="background-color:<%=MonthColor%>;">
  	<img src="<%= ThisFlag %>" title="<%= FlagMessage %>" height="13px" width="13px">
		<br>
 		<font color="<%=textcolor1%>" size="<%=fontsize2%>" face="<%=font1%>"><%=DisplayDate%></font>
	</td>

	<td align="Center" width=<%=TNameWidth%>px ><%

'response.write("sTStatus = "&sTStatus)

		Dim TournamentLinkTitle, redbutton
		TournamentLinkTitle="Check details for "&sTourname







	  IF sTStatus="3" THEN 
	  		' --- Tournament has been cancelled ---
	  		%><font color="<%=textcolor1%>" size="<%=fontsize2%>"><% Response.Write(sTourName&"</a> ("&sTSanction&" ) - ") %></font><font color="red" size="<% =fontsize2 %>"><% Response.write("CANCELLED !!") %></font><%
	  ELSEIF process<>"viewreg" THEN 
	  		%>
	  		<font color="<% =textcolor1 %>" size="<% =fontsize2 %>">
	  			<a href="<%=ThisFileName%>?pvar=TourInfo&TourID=<%=sTourID%>&olrds=<%=OLRButtonEntryStatus%>&rg=<%=sTourRange%>&sl=<%=sl%>&tr=<%=tr%>&ju=<%=ju%>&wb=<%=wb%>&ws=<%=ws%>&wu=<%=wu%>&bf=<%=bf%>&kb=<%=kb%>&hy=<%=hy%>&sExclude=<%=sExclude%>" title="<%=TournamentLinkTitle%>" >
	  				<%=sTourName%>
	  			</a> 
	  				( <%=sTSanction%> )  
	  		</font>
	  		<%
	  ELSE 
	  		%>
	  		<font color="<% =textcolor1 %>" size="<% =fontsize2 %>">
	  			<a href="<%=ThisFileName%>?pvar=TourInfo&TourID=<%=sTourID%>&olrds=<%=OLRButtonEntryStatus%>&rg=<%=sTourRange%>&sl=<%=sl%>&tr=<%=tr%>&ju=<%=ju%>&wb=<%=wb%>&ws=<%=ws%>&wu=<%=wu%>&bf=<%=bf%>&kb=<%=kb%>&hy=<%=hy%>&sExclude=<%=sExclude%>" title="Check tournament details" onclick="window.open(this.href, popupwindow, width=900;height=900;left=100;top=50;scrollbars;resizable); return false;">
	  				<%=sTourName%>
	  			</a>
	  			<%=sTSanction%>
	  		</font>
	  		<%
	  END IF 
	  
	  %>
	  <br>
	  <font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><%= ThisDescription %></font>
	</td>
	<%



	' --------------------------------------------------------------------
	' --- ADMIN LOGIN to access entry status report and other features ---
	' --------------------------------------------------------------------

' width=110px
	IF process="admcode" AND (sAllowRegistrationsCheck="on" OR adminmenulevel>=50) AND sTStatus >= 0 AND sPayPalOK=true AND sPayPalAct<>"" THEN
			%>
			<form action="/rankings/login_registrar.asp?sTourID=<%=sTourID%>" method="post">
		  		<td align="Center" style="word-wrap:break-word;" >
					<font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><%=sTCity%>,&nbsp;<%=ucase(sTState)%></FONT>
		    	<br>
					<input type="submit" style="width:7em; height:1.7em" value="Select" title="Click here to select this tournament.">
		  	</td>
	  	</form>
	  	<%


	' -------------------------------------------------------------- 
	' --- Select tournament for Check Registration Status button ---
	' -------------------------------------------------------------- 

	'ELSEIF process="viewreg" AND sTStatus > 1 AND ( sAllowRegistrationsCheck="on" OR adminmenulevel>=50 ) AND sPayPalOK=true AND sPayPalAct<>"" AND (  ( sTDateE>=Date ) OR adminmenulevel>=20 OR TestValidAdminCode=true  ) THEN 
	ELSEIF process="viewreg" AND ( sAllowRegistrationsCheck="on" OR adminmenulevel>=50 ) AND sPayPalOK=true AND sPayPalAct<>"" AND (  ( sTDateE>=Date ) OR adminmenulevel>=20 OR TestValidAdminCode=true  ) THEN 
			OLRCount = OLRCount + 1
			%>
			<form action="/rankings/view-registration.asp" method="post">
		  	<td  width=110px align="Center">
					<font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><%=sTCity%>,&nbsp;<%=ucase(sTState)%></FONT>
		    	<br>
					<input type="submit" style="width:7em; height:1.7em" value="Entry Status" title="Click here to view the status of your entry for this tournament." >
					<input type="hidden" name="sTourID" value="<%=sTourID%>">
		  	</td>
	  	</form><%

			' --- Removed 7-23-2013
			' OLRButtonEntryStatus


	' -----------------------------
	' --- View Scorebook Button ---
	' -----------------------------
	ELSEIF objFSO.FileExists (PathToScorebks & "\" & sTourID & "CS.HTM") THEN
		%>
		  <td  width=110px align="Center">
			<font color="<% =textcolor1 %>" size="<% =fontsize1 %>">
		    	(<a href="/rankings/ScoreBks/<%=sTourID%>CS.HTM" Target="_blank"
		    	title="View the Scorebook for this tournament in a separate window">ViewScorebk</a>)</FONT><br>
			<font color="<% =textcolor1 %>" size="<% =fontsize1 %>"><%=sTCity%>,&nbsp;<%=ucase(sTState)%></font>
		  <%
' img src="/rankings/images/buttons/Scorebook.bmp" alt="View the Scorebook" 

	 ' --- Show OLR button if same conditions as when future (See below)
	 IF process<>"viewreg" AND sTStatus >= 0 AND (sAllowRegistrationsCheck="on" OR adminmenulevel>=50) AND sUseOLReg=True AND sPayPalAct<>"" AND ( sTDateE>=Date OR adminmenulevel>=20 ) THEN 
			OLRCount = OLRCount + 1
			%>
			<form action="/rankings/<% =RegFileForLink %>" method="post">
		    	<%
					
					redbutton="no"
					IF redbutton="yes" AND RIGHT(TRIM(sTourID),3)="999" THEN %>
							<input type="submit" style="width:8em; height:1.7em; background-color:red; color:white" value="Enter Online" title="<%=EntryButtonTitle%>" <%=OLRButtonEntryStatus%>>
							<input type="hidden" name="sTourID" value="<%=sTournID%>"><%
					ELSE %>
							<input type="submit" style="width:8em; height:1.7em; font-size: 9pt" value="Enter Online" title="<%=EntryButtonTitle%>" <%=OLRButtonEntryStatus%>>
							<input type="hidden" name="sTourID" value="<%=sTourID%>"><%
					END IF 
					
					%>
		  </form>
	  	</td><%
		END IF

	' ---------------------------------------------  
	' --- Online Entry button in list for user ---
	' ---------------------------------------------   


	ELSEIF process<>"viewreg" AND sTStatus >= 0 AND (LEFT(sTourID,6)="16E036" OR sAllowRegistrationsCheck="on" OR adminmenulevel>=50) AND sUseOLReg=True AND TRIM(sPayPalAct)<>"" AND ( sTDateE>=Date OR adminmenulevel>=20 ) THEN 
' --- 16C037 - RELoRHXKCJ
' --- UN: GuidBkAWSE
' --- PW: ashley


TestThis=2	
IF TestThis=1 AND LEFT(sTourID,6)="16E036" THEN
response.write("<br>process = "&process)

response.write("<br>LEFT(sTourID,6) = "&LEFT(sTourID,6))
response.write("<br>sUseOLReg = "&sUseOLReg)
response.write("<br>sPayPalAct = "&sPayPalAct&"<br>")

response.write("<br><br>This is start of conditions")
response.write("<br>")
response.write(process<>"viewreg")
response.write("<br>sTStatus = "&sTStatus)
response.write("<br>LEFT(sTourID,6)=16E036 OR sAllowRegistrationsCheck=on OR adminmenulevel>=50")
response.write(LEFT(sTourID,6)="16E036" OR sAllowRegistrationsCheck="on" OR adminmenulevel>=50)
response.write("<br>sUseOLReg = "&sUseOLReg)
response.write("<br>sPayPalAct = "&sPayPalAct)
response.write("<br>")
response.write(sTDateE>=Date OR adminmenulevel>=20)


END IF

			OLRCount = OLRCount + 1
			%>
			<form action="/rankings/<% =RegFileForLink %>" method="post">
		  	<td  width=110px align="Center">
					<font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><%=sTCity%>,&nbsp;<%=ucase(sTState)%></FONT>
		    	<br>
		    	<%
					
					redbutton="no"
					IF redbutton="yes" AND RIGHT(TRIM(sTourID),3)="999" THEN %>
							<input type="submit" style="width:8em; height:1.7em; background-color:red; color:white" value="Enter Online" title="<%=EntryButtonTitle%>" <%=OLRButtonEntryStatus%>>
							<input type="hidden" name="sTourID" value="<%=sTournID%>"><%
					ELSE %>
							<input type="submit" style="width:8em; height:1.7em; font-size: 9pt" value="Enter Online" title="<%=EntryButtonTitle%>" <%=OLRButtonEntryStatus%>>
							<input type="hidden" name="sTourID" value="<%=sTourID%>"><%
					END IF 
					
					%>
		  	</td>
	  	</form><%





	ELSE %>
		  <td width=110px align="Center">
			<font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><%=sTCity%>,&nbsp;<%=ucase(sTState)%></FONT>
		  </td><%
	END IF 



	' ------------------------------------------------------------------------------------
	' --- Displays signal "YES" to indicate each event is available at that Tournament. ---
	' ------------------------------------------------------------------------------------

	IF sSL_Offered="Y" THEN 
		colcount=colcount+1
		%><td width=35px align="Center" valign="top" bgcolor=<%=TableColor1%>><font color="<% =textcolor1 %>" size="<% =fontsize2 %>">YES</FONT></td><%
	ELSEIF sl="on" THEN
		%><td width=35px bgcolor=<%=TableColor1%>>&nbsp;</td><%
	END IF

	IF sTR_Offered="Y" THEN
		colcount=colcount+1
		%><td width=35px align="Center" valign="top" bgcolor=<%=TableColor1%>><font color="<% =textcolor1 %>" size="<% =fontsize2 %>">YES</FONT></td><%
	ELSEIF tr="on" THEN
		%><td width=35px bgcolor=<%=TableColor1%>>&nbsp;</td><%
	END IF

	IF sJU_Offered="Y" THEN
		colcount=colcount+1
		%><td width=35px align="Center" valign="top" bgcolor=<%=TableColor1%>><font color="<% =textcolor1 %>" size="<% =fontsize2 %>">YES</FONT></td><%
	ELSEIF ju="on" THEN
		%><td width=35px bgcolor=<%=TableColor1%>>&nbsp;</td><%
	END IF

	IF sWB_Offered="Y" THEN
		colcount=colcount+1
		%><td width=35px align="Center" valign="top" bgcolor=<%=TableColor1%>><font color="<% =textcolor1 %>" size="<% =fontsize2 %>">YES</FONT></td><%
	ELSEIF wb="on" THEN
		%><td width=35px bgcolor=<%=TableColor1%>>&nbsp;</td><%
	END IF

	IF sWS_Offered="Y" THEN
		colcount=colcount+1
		%><td width=35px align="Center" valign="top" bgcolor=<%=TableColor1%>><font color="<% =textcolor1 %>" size="<% =fontsize2 %>">YES</FONT></td><%
	ELSEIF wu="on" THEN
		%><td width=35px bgcolor=<%=TableColor1%>>&nbsp;</td><%
	END IF

	IF sWU_Offered="Y" THEN
		colcount=colcount+1
		%><td width=35px align="Center" valign="top" bgcolor=<%=TableColor1%>><font color="<% =textcolor1 %>" size="<% =fontsize2 %>">YES</FONT></td><%
	ELSEIF ws="on" THEN
		%><td width=35px bgcolor=<%=TableColor1%>>&nbsp;</td><%
	END IF

	IF sBF_Offered="Y" THEN 
		colcount=colcount+1
	    	%><td width=35px align="Center" valign="top" bgcolor=<%=TableColor1%>><font color="<% =textcolor1 %>" size="<% =fontsize2 %>">YES</FONT></td><%
	ELSEIF bf="on" THEN
		%><td width=35px bgcolor=<%=TableColor1%>>&nbsp;</td><%
	END IF 

	IF sKB_Offered="Y" THEN 
		colcount=colcount+1
	    	%><td width=35px align="Center" valign="top" bgcolor=<%=TableColor1%>><font color="<% =textcolor1 %>" size="<% =fontsize2 %>">YES</FONT></td><%
	ELSEIF kb="on" THEN
		%><td width=35px bgcolor=<%=TableColor1%>>&nbsp;</td><%
	END IF 
	IF sHY_Offered="Y" THEN 
		colcount=colcount+1
	    	%><td width=35px align="Center" valign="top" bgcolor=<%=TableColor1%>><font color="<% =textcolor1 %>" size="<% =fontsize2 %>">YES</FONT></td><%
	ELSEIF hy="on" THEN
		%><td width=35px bgcolor=<%=TableColor1%>>&nbsp;</td><%
	END IF 

	IF sDA_Offered="Y" THEN 
		colcount=colcount+1
	    	%><td width=35px align="Center" valign="top" bgcolor=<%=TableColor1%>><font color="<% =textcolor1 %>" size="<% =fontsize2 %>">YES</FONT></td><%
	ELSEIF da="on" THEN
		%><td width=35px bgcolor=<%=TableColor1%>>&nbsp;</td><%
	END IF 

	IF sJD_Offered="Y" THEN 
		colcount=colcount+1
	    	%><td width=35px align="Center" valign="top" bgcolor=<%=TableColor1%>><font color="<% =textcolor1 %>" size="<% =fontsize2 %>">YES</FONT></td><%
	ELSEIF jd="on" THEN
		%><td width=35px bgcolor=<%=TableColor1%>>&nbsp;</td><%
	END IF 

	IF sAD_Offered="Y" THEN 
		colcount=colcount+1
	    	%><td width=35px align="Center" valign="top" bgcolor=<%=TableColor1%>><font color="<% =textcolor1 %>" size="<% =fontsize2 %>">YES</FONT></td><%
	ELSEIF ad="on" THEN
		%><td width=35px bgcolor=<%=TableColor1%>>&nbsp;</td><%
	END IF 

%>
</tr><%

END SUB





' ------------------------------
    SUB DisplaySingleListing
' ------------------------------


sTourID = TRIM(Request("TourID"))
sExclude = Request("sExclude")


' --- Uses variable definition from Tools_Registration.asp ---
DefineTourVariables_New


OLRButtonEntryStatus="enabled"
EntryButtonTitle="Enter this tournament with our online entry form"
RegFileForLink = "registration16.asp"



	' --- Determines if corresponding record exists in Ski Year table - avoids errors until record is entered --
	set rsSkiYear=Server.CreateObject("ADODB.recordset")
	rsSkiYearSQL = "SELECT * FROM "&SkiYearTableName&" WHERE EndDate>='"&sTDateE&"' AND BeginDate<='"&sTDateS&"'" 
	rsSkiYear.open rsSkiYearSQL, SConnectionToTRATable


	
	whiletesting=0
	
	IF rsSkiYear.eof THEN 
			OLRButtonEntryStatus="disabled"
			EntryButtonTitle="Ski Year Administrative table must be updated for this Ski Year before available for OLR "
	ELSEIF request("olrds") ="disabled" and sTStatus>=0 THEN 
			OLRButtonEntryStatus="disabled"
			EntryButtonTitle="This tournament is no longer open to Online Registration"

	ELSEIF whiletesting=1 AND adminmenulevel < 50 AND RIGHT(cStr(Year(sTDateS)),2)>="16" THEN 
  		OLRButtonEntryStatus="disabled"
			EntryButtonTitle="Online Registration has not been activated for 2016 - Thanks for your patience"
	'ELSEIF NOT(rs("OLRDisplayStatus")) THEN
	'		sExclude="no"
	'		OLRButtonEntryStatus="disabled"
	'		EntryButtonTitle="This tournament is no longer open to Online Registration"
	ELSEIF sUseOLReg=true AND sPayPalOK=0 OR sOLR_PD=0 OR TRIM(sPayPalAct)="" THEN
			sExclude="no"
			OLRButtonEntryStatus="disabled"
			EntryButtonTitle="This tournament is not yet available for Online Registration"
	END IF


	' --- OLR Button status ---
	IF DisableOLRButtons=true THEN
			OLRButtonEntryStatus="disabled"
			EntryButtonTitle="Online Registration is Temporarily Disabled - Thanks for your patience"
	END IF






		%>
		<TABLE width="<%=TourTableWidth+10%>" class="noborder" align=center>
			<tr>
				<td width=100px>&nbsp;</td>
        <td width=225px>&nbsp;</td>
        <td width=150px>&nbsp;</td>
        <td width=175px>&nbsp;</td>
      </tr>
      <tr>
        <td colspan=2>
	    		<font color="<% =textcolor2 %>" size=3><b><%= sTourName %></b></FONT>
	     		<font color="<% =textcolor1 %>" size="<%= fontsize3 %>">&nbsp;&nbsp; <%=sTSanction%></FONT>
	 			</td>
        <td colspan=2>
	    		<font color="<% =textcolor1 %>" size="<%= fontsize3 %>"><b>SiteID:</b></FONT>
	     		<font color="<% =textcolor1 %>" size="<% =fontsize3 %>">&nbsp; <%=sTSiteID%></FONT>
	 			</td>
      </tr>
			<tr>
				<td colSpan=2 align="left" width="350px">
					<font color="<% =textcolor1 %>" size="<% =fontsize3 %>"><b><%=sTourCity%>, <%=sTourState%></b></FONT>
	 			</td>

	 			<td colspan=2 align=right width="350px">
	 				<% 
					IF sTDateS = sTDateE THEN 
							%><font color="<% =textcolor1 %>" size=<% =fontsize3 %>><b><%= sTDateS %></b></font><%
					ELSE 
							%><font color="<% =textcolor1 %>" size=<% =fontsize3 %>><b><%= sTDateS %> to <%= sTDateE %></b></font><%
					END IF 
					%>
	 			</td>
			</tr>
      <tr>
        <td colspan=4>&nbsp;</td>
      </tr>

      <tr>
				<td width="100px" align=left><font color="<% =textcolor1 %>" size="<% =fontsize2 %>" ><b>Description:</b></font></td>
				<td colspan=3 align=left>
					<font color="<% =textcolor1 %>" size="<% =fontsize2 %>">
					<%

					IF TRIM(sTDescription)<>"" THEN
							Response.write(sTDescription)
					ELSEIF TRIM(sFDescription)<>"" THEN 
							Response.write(sFDescription)
					ELSEIF TRIM(sWDescription)<>"" THEN 
							Response.write(sWDescription)
					ELSEIF TRIM(KTDescription)<>"" THEN 
							Response.write(sKDescription)
					ELSEIF TRIM(sTDescription)<>"" THEN 
							Response.write(sCDescription)
					END IF 
	
					%>
					</font>
				</td>
      </tr>

      <tr>
        <td valign="top"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><b>Sponsor:</b></FONT></td>
        <td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><% Response.Write (sTSponsor) %></FONT></td>
      </tr>
      <tr>
         <td valign="top"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><b>Site:</b></FONT></td>
         <td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><% Response.Write (sTSite) %></FONT></td>
       </tr>
       <tr>
         <td valign="top"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><b>Location:</b></FONT></td>
         <td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><% Response.Write (sTourCity & ", " & sTourState) %></FONT></td>
       </tr>
       <tr>
         <td valign="top"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><b>Directions:</b></FONT></td>
         <td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><% Response.Write (sGTSDirections) %></FONT></td>
       </tr>

       <tr>    
         <td align="left" colSpan="4"><hr width="650"></td>
       </tr>

       <tr>
         <td valign="top"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><b>Entry:</b></FONT></td>
         <td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><% IF sTOpenClosed=True THEN Response.Write ("Closed") ELSE Response.Write ("Open") %></FONT></td>
       </tr>
       <tr>
         <td valign="top"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><b>Tow Boat:</b></FONT></td>
         <td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><% IF sTTowBoatClosed=True THEN Response.Write ("Closed") ELSE Response.Write ("Open") %></FONT></td>
       </tr>
       <tr>    
         <td align="left" colSpan="4"><hr width="650"></td>
       </tr>
       <tr>
         <td valign="top"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><b>Entry Limit:</b></FONT></td>
         <%
         tyu=1
         IF tyu=2 THEN
         		%>	
         		<td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><% If sMaxPulls>0 THEN Response.Write(sMaxPulls&" Pulls") ELSE Response.write("None")%></FONT></td>
         		<%
         ELSE
						Dim MaxPullsText
						MaxPullsText = "None"
						IF sMaxPulls>0 THEN 
								MaxPullsText = CStr(sMaxPulls)+ " Pulls"
						ELSEIF sTEntryLimit <> "None" THEN
								MaxPullsText = sTEntryLimit
						END IF
         		%>	
         		<td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><%= MaxPullsText %></FONT></td>
         		<%
         END IF
				 %>	
       </tr>
       <tr>
         <td valign="top"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><b>Entry Fees:</b></FONT></td>
         <td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><% Response.Write (sTEntryFees) %></FONT></td>
       </tr>
       <tr>
         <td valign="top"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><b>Entry Deadline:</b></FONT></td>
         <td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><% Response.Write (sTLateDate) %></FONT></td>
       </tr>
       <tr>
         <td valign="top"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><b>Late Fee:</b></FONT></td><%
	 IF sTLFPerDay=true THEN 
	 		%><td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>">$<% Response.Write (sTLateFee) %> Per Day</FONT></td><%
	 ELSE  
	 		%><td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>">$<% Response.Write (sTLateFee) %> </FONT></td><%
	 END IF 
	 %>
			</tr>
			<tr>    
				<td align="left" colSpan="4"><hr width="650"></td>
			</tr>
			<tr>
				<td valign="top">
         	<font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><b>Send Entries To:</b></FONT>
				</td>
				<td valign="top" colSpan="2">
         	<font color="<% =textcolor1 %>" size="<% =fontsize2 %>">
						<% 
						Response.Write (sTRegistrarName& "<br>" & sTRegistrarAddr & "<br>" & sTRegistrarCity & ", " & sTRegistrarState & "  " & sTRegistrarZip&"<br>"&sTRegistrarPhone&"<br>"&sTRegistrarEmail) 
						%>
					</font>
				</td>
	 			<td align=right>
	  			<form action="http://www.usawaterski.org/rankings/news/AWSA_Entry_Form_2019.PDF" method="get" target="_blank">
	  				<%
						IF TRIM(sTSptsGrpID)="AWS" OR sGrassroots=true THEN
		  					%><input type="submit" style="width:10em; height:1.8em; font-size: 9pt" value="Print Entry" title="Display printable Entry Form"><%
						END IF  
						%>
			   	</form>
   			<%



	 IF process<>"viewreg" AND sTStatus >= 0 AND ( LEFT(sTourID,6)="16E036" OR sAllowRegistrationsCheck="on" OR adminmenulevel>=50) AND sUseOLReg=True AND sPayPalAct<>"" AND ( sTDateE>=Date OR adminmenulevel>=50 ) THEN 
		%>
		<br>
		<form action="/rankings/<% =RegFileForLink %>" method="post">
		   <input type="submit" style="width:10em; height:1.8em; font-size: 9pt" value="Enter Online" title="<%=EntryButtonTitle%>" <%=OLRButtonEntryStatus%>>
		   <input type="hidden" name="sTourID" value="<%=sTourID%>">
		</form>
		<%


	' --- View Scorebook Button ---
	' ------------------------------
	ELSEIF objFSO.FileExists (PathToScorebks & "\" & sTourID & "CS.HTM") THEN
			%>
      <br>
			<form action="/rankings/ScoreBks/<%=sTourID%>CS.HTM" method="post" Target="_blank">
				<input type="submit" style="width:10em; height:1.7em" value="View Scorebook" title="View the Scorebook for this tournament in a separate window">
	  	</form><%
	END IF 
		%>
		</td>
	</tr>
	<tr>    
		<td align="left" colSpan="4"><hr width="650"></td>
	</tr>
	<tr>
    <td valign="top"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><b>Accommodations:</b></FONT></td>
    <td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><% Response.Write (sGTAccommodation) %></FONT></td>
  </tr>
       <tr>
         <td valign="top"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><b>Awards:</b></FONT></td>
         <td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><% Response.Write (sGTAwards) %></FONT></td>
       </tr>
       <tr>
         <td valign="top"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><b>Practice:</b></FONT></td>
         <td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><% Response.Write (sGTPractice) %></FONT></td>
       </tr>
       <tr>
         <td valign="top"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><b>Start Time:</b></FONT></td>
         <td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><% Response.Write (sGTStartTime) %></FONT></td>
       </tr>
       <tr>
         <td valign="top"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><b>Sched of Events:</b></FONT></td>
         <td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><% Response.Write (sGTSofE) %></FONT></td>
       </tr>
       <tr>
         <td valign="top"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><b>Entry Reqts:</b></FONT></td>
         <td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><% Response.Write (sG_IWWF_req) %></FONT></td>
       </tr>
       <tr>
         <td valign="top"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><b>Comments:</b></FONT></td>
         <td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><% Response.Write (sGTComments) %></FONT></td>
       </tr>
       <%

	' -------------------------------
	' --- BEGIN OFFICIALS SECTION --- 
	' ------------------------------- %>
       <tr>    
         <td align="left" colSpan="4"><hr width="650"></td>
       </tr>

       <tr>
         <td valign="top"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><b>Tourn Director:</b></FONT></td>
         <td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><% Response.Write (sTDirName) %></FONT></td>
       </tr>

       <tr>
         <td valign="top"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><b>Chief Judge:</b></FONT></td>
	 <td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><% Response.Write (sCJudge) %></FONT></td>
       </tr>
       <tr>
         <td valign="top"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><b>Chief Scorer:</b></FONT></td>
         <td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><% Response.Write (sCScorer) %></FONT></td>
       </tr>
       <tr>
         <td valign="top"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><b>Chief Boat Driver:</b></FONT></td>
         <td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><% Response.Write (sCDriver) %></FONT></td>
       </tr>
       <tr>
         <td valign="top"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><b>Chief Safety:</b></FONT></td>
         <td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><% Response.Write (sCSafety) %></FONT></td>
       </tr>
       <tr>
         <td valign="top"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><b>Announcer:</b></FONT></td>
         <td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><% Response.Write (sAnnouncer) %></FONT></td>
       </tr>

       <tr>
         <td valign="top"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><b>Tech Controller:</b></FONT></td>
         <td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><% Response.Write (sTechCont) %></FONT></td>
       </tr>
       <tr>
         <td valign="top"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><b>Pan Am Judge:</b></FONT></td>
         <td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><% Response.Write (sPanAmJudge) %></FONT></td>
       </tr>
       <tr>
         <td valign="top"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><b>Appt'd Judges:</b></FONT></td>
         <td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><% Response.Write (sAp1Judge) %></FONT></td>
       </tr>
       <tr>
         <td valign="top">&nbsp;</td>
         <td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><% Response.Write (sAp2Judge) %></FONT></td>
       </tr>
       <tr>
         <td valign="top">&nbsp;</td>
         <td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><% Response.Write (sAp3Judge) %></FONT></td>
       </tr>
       <tr>
         <td valign="top">&nbsp;</td>
         <td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><% Response.Write (sAp4Judge) %></FONT></td>
       </tr>
       <tr>
         <td valign="top">&nbsp;</td>
         <td valign="top" colSpan="3"><font color="<% =textcolor1 %>" size="<% =fontsize2 %>"><% Response.Write (sAp5Judge) %></FONT></td>
       </tr>
     </TABLE>
     <%


END SUB











' ---------------------------------------------------------------------------------------------------------
    SUB PerformSQLQuery_2010  ' ----------------  BUILD SQL statement   -----------------------------------
' ---------------------------------------------------------------------------------------------------------

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' 			--- IMPORTANT ---

' --- This modules is only applicable to tournaments in 2010 and after

' ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++





' --- From Jim Meis 10-29-2007
' Class I (and N) ARE "traditional" AWSA/NCWSA classes.  NCWSA is supposed to use Class I, and Class I is expected by WSTIMS for collegiate events.
' Class I should only be used by NCWSA so I should probably remove it as an option on the AWSA sanction form.
' Classes I and N predate Grassroots, have different officials requirements, allow different officials work credits, and have different sanction fees.
' Classes I and N also require AWSA or NCWSA Region Sanction Approval which Grassroots technically does not.   
' If it has a traditional event, the Sports Division admin and HQ give approval as part of the traditional approval. 
'     If all the traditional requirements are in place any Grassroots program is automatically OK without much thought

' TEventSlalom, TEventTrick, and TEventJump -   Barefoot traditional tournaments (ABC) as users of those fields.  
' Users are:  AWS, NCW, ABC use all 3 and AKA uses TEventSlalom and TEventTrick

' THSClassF is the original Fun field.  It is and always was distinct from THSClassN.   Officials required are different etc.
' Current designation for fun 3 event is TEventF3ev=1.  More specifically this indicates a Grassroots event that the sponsor is characterizing
'    as 3 event type.  Could be sanctioned under any of the sports divisions.
' All the fields beginning with TEventF, except TEventFun, are the most recent Grassroots fields.
' Sponsor can offer multiple classes and skier can pick what level he wants to ski at - so THSClassR is 1 if R is offered and 0 if it is not.  


' --- From Jim Meis 10/28/2007
' In response to the question "why so many fields?"
' They came about because of the umpteen revisions to the "Fun", NWL, NBL, NSL, and Grassroots programs in the past 3 or 4 years.
' Started out with THSClassF, then added TEventFun to allow stand alone fun, then dropped THSClassF to separate FUN from 3 event 
'   and to allow fun to be sanctioned by other sports divisions, then added NSL, NWL, NBL, then added Grassroots, 
'   then dropped NSL, NWL, and NBL from sanction form but left it on the adverts whenever grassroots was selected 
'   for 3ev, Barefoot or wakeboard).  Latest directive says drop NSL, NWL, and NBL entirely, change the fun "events" 
'   already offered and add new ones


' --- From Jim Meis 3/1/2007 ---
' There are separate description fields in swift for traditional, fun, clinics, Wakeboard and Kneeboard events.  
' They have zero length strings if there is no description  (no matching events).
' When a sanction includes more than one of these categories you need to concatenate the description fields to get all the information.


' Tschedul.TDescription - AWS, ABC, or NCW standard events
' Tschedul.WDescription - Wakeboard standard events
' Tschedul.KDescription - Kneeboard standard events
' Tschedul.FDescription - Fun Events including NSL 
' Tschedul.CDescription - Clinics

' Tschedul.TStatus = 0     Application received
' Tschedul.TStatus = 1     Region approved
' Tschedul.TStatus = 2    USAWS Approved
 
' Tschedul.TPending   True until first save by an administrator - To publish must be TPending must be false and the other conditions below 
'     must be met.
' RegnSetup.ShowPSchedule - On Off switch for the entire schedule for a Region.
' RegnSetup.ShowGBLink  -  controls display of tournament schedule as pick list for sponsors. Generally set same as ShowPSchedule
' RegnSetup.GBPolicy = true if ad is allowed to be displayed before full Region approval of the sanction.  Necessary but not sufficient.
' Tschedul.TKitOKGuidebookAd  -   Set by Region Admin on each sanction application - Gives Region approval of content of the Ad.  Allows 
'     publication if GBPolicy is true and ShowPSchedule is true. True = Bit 1 False = Bit 0
 
' Tschedule.TPending is true by default - it is changed to false after the first review and save by an administrator.  Nothing should be 
'    displayed unless TPending is false (has received its first review and save).

' Some regions particularly Western Region do not want tournament information posted at all until the Guidebook is published.  Regions can 
'    toggle ShowPSchedule on and off in their Region Preferences. True = OK to show as long as the rest of the conditions are met.  
'    False means do not show under any conditions.
 
' ShowGBLink is related but only important for SWIFT - it determines if the tournament schedule is used as a pick list for sponsors revising 
'    tournaments or if they have to supply their tournAppID and Edit Code blind.
 
' GBPolicy -    Guidebook Policy determines if advertisement is allowed to display before the Region has given their sanction approval.
' Some regions require that the region part of the sanction process be complete (fees paid) and approved before displaying the advertisement.  
'    If guidebook = false then don't display unless TStatus >= 1 - NOT APPLICABLE AFTER 1-16-2014
 
' Other Regions Others don't care and allow publication prior to region approval.   They only require that the ad itself be approved.
' In this case TStatus could be 0 or higher, Guidebook must be true, and TKitOKGuidebookAd must be true (ad itself has region approval)

' The  ShowReg, ShowAppointed, etc control display of specific parts of an ad - also set in region preferences. Some regions don't want 
'    registrar information published online until the guidebook is published on the theory that it levels the playing field for entries.
' ---------------------------------------------------------------------------------------------------------------------------------------



sSQL = "SELECT TOP 800 "
sSQL = sSQL + "ST.TournAppID, TName, ST.SptsGrpID, TDescription, WDescription, ST.FDescription, KDescription, CDescription" 
sSQL = sSQL + ", TSanction, TSanType, TDateE, TDateS, TCity, Tstate, Pending, Deleted"
sSQL = sSQL + ", ShowPSched, TKitOKGuideBookAd, GBPolicy, TStatus, Deleted, ShowRegistrar"
sSQL = sSQL + ", OK2Publish"

sSQL = sSQL + ", Gr1AWSPulls, Gr1ABCPulls, Gr1USWPulls, Gr1AKAPulls, Gr1USHPulls, Gr1WSDPulls"
sSQL = sSQL + ", Gr2USH_FreeRidePulls, Gr2USH_JumpOutPulls, Gr2USH_BigAirPulls, Gr2USH_3TrickPulls"
sSQL = sSQL + ", Gr2AWS_SPulls, Gr2AWS_TPulls, Gr2ABC_SPulls, Gr2ABC_TPulls, Gr2USW_WPulls, Gr2USW_SkatePulls, Gr2USW_SurfPulls, Gr2USW_RailJamPulls" 
sSQL = sSQL + ", Gr2AKA_SPulls, Gr2AKA_TPulls, Gr2AKA_FreePulls, Gr2AKA_FlipPulls"
sSQL = sSQL + ", OLRDisplayStatus, UseOLReg, OLR_PD"

sSQL = sSQL + ", TRS.PayPalAct, TRS.PayPalOK"

sSQL = sSQL + ", ST.SptsGrpID AS sSptsGrpID, ST.TRegion AS STRegion"
sSQL = sSQL + ", ST.TEventNWL, ST.TEventNBL, ST.TEventNSL, ST.THSClassN, ST.THTClassN, ST.THJClassN"
sSQL = sSQL + ", TEventF3ev"
sSQL = sSQL + ", ST.TEventWake, ST.TEventWSkate, ST.TEventWSurf, ST.TEventFW"
sSQL = sSQL + ", WWakeW, WSkateW, WSurfW"
sSQL = sSQL + ", TRS.sClassC, TRS.sClassE, TRS.sClassL, TRS.sClassR, TRS.sClassCash, TRS.sClassX"
sSQL = sSQL + ", TRS.tClassC, TRS.tClassE, TRS.tClassL, TRS.tClassR, TRS.tClassCash, TRS.tClassX"
sSQL = sSQL + ", TRS.jClassC, TRS.jClassE, TRS.jClassL, TRS.jClassR, TRS.jClassCash, TRS.jClassX"
sSQL = sSQL + ", USClassC, UTClassC, UJClassC"

' --- Fields obsolete beginning in 2010
sSQL = sSQL + ", ST.TEventSlalom, ST.TEventTrick, ST.TEventJump"
'sSQL = sSQL + ", ST.TEventFun"



sSQL = sSQL + ", ST.THSClassI, ST.THJClassI, ST.THTClassI"
sSQL = sSQL + ", ST.JDClin, ST.ADClin, ST.TEventFHF, ST.TEventFKB"

sSQL = sSQL + " FROM " &SanctionTableName&" AS ST"

sSQL = sSQL + " LEFT JOIN "&RegnSetupTableName&" AS RT ON ST.SptsGrpID = RT.SptsGrpID AND ST.TRegion = RT.TRegion"
sSQL = sSQL + " LEFT JOIN "&TRegSetupTableName&" AS TRS ON TRS.TournAppID = ST.TournAppID"




	sSQL = sSQL + " WHERE (11=12 "   ' --- This is the top of the bracket of all event inclusions ---


	IF sTourLevel="cash" THEN
		sSQL = sSQL +" OR (ST.THSClassCash<>0 OR ST.THTClassCash<>0 OR ST.THJClassCash<>0)"
	END IF

	IF sTourLevel="premier" OR sTourLevel="all" THEN

			' --- 3 Event Premier ---		
			IF sl="on" OR tr="on" OR ju="on" THEN 		

					' --- Top of AWS bracket "OR" ---
					' -----------------------------
					sSQL = sSQL + " OR (ST.SptsGrpID='AWS' AND (3=4" 	' --- Top of AWS stuff

					' ---  ST.TEventSlalom etc maintained to allow fall 2009 tournaments to display in 2010 criteria
					IF sl="on" THEN 
							' --- Changed 1-23-2010
							sSQL = sSQL + " OR (TRS.sClassC + TRS.sClassE + TRS.sClassL + TRS.sClassR + TRS.sClassCash + TRS.sClassX)>0  OR ST.TEventSlalom<>0"
					END IF
					IF tr="on" THEN 
							' --- Changed 1-23-2010
							sSQL = sSQL + " OR (TRS.tClassC + TRS.tClassE + TRS.tClassL + TRS.tClassR + TRS.tClassCash + TRS.tClassX)>0  OR ST.TEventTrick<>0"
					END IF
					IF ju="on" THEN
							' --- Changed 1-23-2010
							sSQL = sSQL + " OR (TRS.jClassC + TRS.jClassE + TRS.jClassL + TRS.jClassR + TRS.jClassCash + TRS.jClassX)>0  OR ST.TEventJump<>0"
					END IF
					sSQL = sSQL + "))"				' --- Bottom of AWS stuff ---
			END IF		

		' --- Wakeboard Premier ---
			IF wb="on" OR ws="on" OR wu="on" THEN 
					sSQL = sSQL + " OR (1=2"
					' --- Changed 1-23-2010 ---
					IF wb="on" THEN sSQL = sSQL + " OR ST.TEventWake<>0 OR WWakeW<>0"
					IF ws="on" THEN sSQL = sSQL + " OR ST.TEventWSkate<>0 OR WSkateW<>0"
					IF wu="on" THEN sSQL = sSQL + " OR ST.TEventWSurf<>0 OR WSurfW<>0"

					sSQL = sSQL + ")" 	
			END IF	
	END IF

	IF sTourLevel="grass" OR sTourLevel="all" THEN
		
			sSQL = sSQL +" OR (5=6"  	'---- Open bracket Grassroots

			' --- Grassroots 3 Event ---
			IF sl="on" THEN 
					sSQL = sSQL + " OR ST.Gr2AWS_SPulls<>0 OR Gr1AWSPulls<>0" 
					sSQL = sSQL + " OR ST.THSClassN<>0"
					' --- TEventFun and THSClassF included for legacy Pre-2009 system ---
					sSQL = sSQL + " OR ST.THSClassF<>0 OR ST.TEventF3ev<>0"
			END IF
			IF tr="on" THEN 		
					sSQL = sSQL + " OR ST.Gr2AWS_TPulls<>0" 
					sSQL = sSQL + " OR ST.THTClassN<>0"
					' --- TEventFun and THTClassF included for legacy Pre-2009 system ---
					sSQL = sSQL + " OR ST.THTClassF<>0"
			END IF
			IF ju="on" THEN 		
					sSQL = sSQL + " OR ST.THJClassN<>0"
			END IF


			' --- Grassroots Wakeboard ---		
			IF wb="on" OR ws="on" OR wu="on" THEN 
					' --- Legacy from Pre-2009 system ---
					sSQL = sSQL + " OR (ST.TEventFW<>0 OR ST.TEventNWL<>0"

					IF wb="on" THEN 
							' --- Changed 1-23-2010 ---
							sSQL = sSQL + " OR Gr2USW_WPulls<>0 OR Gr2USW_RailJamPulls<>0 OR Gr1USWPulls<>0 OR WWakeW>0" 
					END IF
					IF ws="on" THEN 
							' --- Changed 1-23-2010 ---
							sSQL = sSQL + " OR Gr2USW_SkatePulls<>0 OR WSkateW>0"
					END IF 
					IF wu="on" THEN 
							' --- Changed 1-23-2010 ---
							sSQL = sSQL + " OR Gr2USW_SurfPulls<>0 OR WSurfW>0"
					END IF 
					sSQL = sSQL + ")" 	
		END IF

		sSQL = sSQL + ")" 		'---- Close bracket Grass

	END IF



	' --- Collegiate ---
	IF sTourLevel="collegiate" THEN 
			' --- 3 Event ---
			IF sl="on" OR tr="on" OR ju="on" THEN 		
					sSQL = sSQL + " OR (ST.SptsGrpID='NCW' AND (1=2 "
					' ---  ST.TEventSlalom etc maintained to allow fall 2009 tournaments to display in 2010 criteria
					IF sl="on" THEN sSQL = sSQL + " OR (TRS.USClassC>0 OR ST.TEventSlalom<>0)"
					IF tr="on" THEN sSQL = sSQL + " OR (TRS.UTClassC>0 OR ST.TEventTrick<>0)"
					IF ju="on" THEN sSQL = sSQL + " OR (TRS.UJClassC>0 OR ST.TEventTrick<>0)"
					sSQL = sSQL + "))" 
			END IF

			' --- Wakeboard ---		
			IF wb="on" OR ws="on" OR wu="on" THEN 
					sSQL = sSQL + " OR (ST.SptsGrpID='NCW' AND ST.TEventFW<>0 OR ST.TEventNWL<>0 OR WWakeW>0  OR WSurfW>0 OR WSkateW>0)" 
			END IF
	END IF
	
	' --- Barefoot ---
	IF bf="on" THEN sSQL = sSQL + " OR (ST.SptsGrpID='ABC' OR ST.TEventNBL<>0 OR Gr1ABCPulls<>0 OR Gr2ABC_SPulls<>0 OR Gr2ABC_TPulls<>0)"

	' --- Kneeboard ---
	IF kb="on" THEN sSQL = sSQL + " OR (ST.SptsGrpID='AKA' OR ST.TEventFKB<>0) OR Gr1AKAPulls<>0 OR Gr2AKA_SPulls<>0 OR Gr2AKA_TPulls<>0 OR Gr2AKA_FreePulls<>0 OR Gr2AKA_FlipPulls<>0"

	' --- Hydrofoil ---
	IF hy="on" THEN 
		' --- Legacy from Pre-2009 ---
		sSQL = sSQL + " OR (ST.TEventFHF<>0"
		sSQL = sSQL + " OR ST.Gr2USH_FreeRidePulls<>0 OR ST.Gr2USH_JumpOutPulls<>0 OR ST.Gr2USH_BigAirPulls<>0 OR ST.Gr2USH_3TrickPulls<>0 OR ST.Gr1USHPulls<>0)"
	END IF

	' --- Clinic ---
	IF ad="on" THEN sSQL = sSQL + " OR ST.ADClin<>0"
	IF jd="on" THEN sSQL = sSQL + " OR ST.JDClin<>0"

	sSQL = sSQL + ")"    ' --- This is the bottom of the bracket of all event inclusions ---


	


	' --- Filters for highest homologation class ---
	HighClass=99
	IF sClass="R" AND (sl="on" OR tr="on" OR ju="on") THEN 
			sSQL = sSQL + " AND (THSClassR<>0 OR THTClassR<>0 OR THJClassR<>0 OR SClassR>0 OR TClassR>0 OR JClassR>0)" 
			HighClass=1

	ELSEIF sClass="L" AND HighClass>1 AND (sl="on" OR tr="on" OR ju="on") THEN
			sSQL = sSQL + " AND (THSClassL<>0 OR THTClassL<>0 OR THJClassL<>0 OR SClassL>0 OR TClassL>0 OR JClassL>0)" 
			HighClass=2

	ELSEIF sClass="E" AND HighClass>2 AND (sl="on" OR tr="on" OR ju="on") THEN
			sSQL = sSQL + " AND (THSClassE<>0 OR THTClassE<>0 OR THJClassE<>0 OR SClassE>0 OR TClassE>0 OR JClassE>0)" 
			HighClass=3

	ELSEIF sClass="C" AND HighClass>3 AND (sl="on" OR tr="on" OR ju="on") THEN
			sSQL = sSQL + " AND (THSClassC<>0 OR THTClassC<>0 OR THJClassC<>0 OR SClassC>0 OR TClassC>0 OR JClassC>0)" 
			HighClass=4

	ELSEIF sClass="N" AND HighClass>4 AND (sl="on" OR tr="on" OR ju="on") THEN
			sSQL = sSQL + " AND (THSClassN<>0 OR THTClassN<>0 OR THJClassN<>0)" 
			HighClass=5

	ELSEIF sClass="F" AND HighClass>5 AND (sl="on" OR tr="on" OR ju="on") THEN
			sSQL = sSQL + " AND (THSClassF<>0 OR THTClassF<>0 OR THJClassF<>0)" 
			HighClass=6

	ELSEIF sClass="F_O" AND (sl="on" OR tr="on" OR ju="on") THEN
			sSQL = sSQL + " AND TEventF3ev<>0" 
			HighClass=7

	END IF



		IF sTourRange <> "" AND sTourRange <> "0" THEN
				IF sTourRange = "1" THEN
						sSQL = sSQL + " AND (ST.TDateE >= '" & Date() & "')"
				END IF

		' --- Ski Year defined as Latest in DivisionTable ---
		IF sTourRange = "2" THEN
				set rsSelectFields=Server.CreateObject("ADODB.recordset")
				rsSelectFields.open "Select Top 1 * FROM "&SkiYearTableName&" WHERE SkiYearID<>'1' ORDER BY BeginDate DESC", SConnectionToTRATable
				IF NOT rsSelectFields.eof THEN
						sSQL = sSQL + " AND (left(ST.TournAppID,2) = '" & right(right(TRIM(rsSelectFields("SkiYearName")),4),2) & "')"
				END IF
				rsSelectFields.close

		' --- Ski Year defined as SECOND latest in DivisionTable ---
		ELSEIF sTourRange = "3" THEN 
				set rsSelectFields=Server.CreateObject("ADODB.recordset")
				rsSelectFields.open "Select Top 2 * FROM "&SkiYearTableName&" WHERE SkiYearID<>'1' ORDER BY BeginDate DESC", SConnectionToTRATable
				IF NOT rsSelectFields.eof THEN
						rsSelectFields.movenext
						sSQL = sSQL + " AND (left(ST.TournAppID,2) = '" & right(right(TRIM(rsSelectFields("SkiYearName")),4),2) & "')"
				END IF
				rsSelectFields.close

		' --- Ski Year defined as THIRD latest in DivisionTable ---
		ELSEIF sTourRange = "4" THEN 
				set rsSelectFields=Server.CreateObject("ADODB.recordset")
				rsSelectFields.open "Select Top 3 * FROM "&SkiYearTableName&" WHERE SkiYearID<>'1' ORDER BY BeginDate DESC", SConnectionToTRATable
				IF NOT rsSelectFields.eof THEN
					  rsSelectFields.movenext
				  	IF NOT rsSelectFields.eof THEN
								rsSelectFields.movenext
								IF NOT rsSelectFields.eof THEN
				  					sSQL = sSQL + " AND (left(ST.TournAppID,2) = '" & right(right(TRIM(rsSelectFields("SkiYearName")),4),2) & "')"
			  				END IF
			  		END IF
				END IF
				rsSelectFields.close

		' --- Current Calendar year if the year is nearly over otherwise last calendar year ---
		ELSEIF sTourRange = "5" THEN
				sSQL = sSQL + " AND Year(TDateS) = '"&Year(Date())&"'"
	
		' --- Last Calendar year if this year is nearly over otherwise two calendar years ago ---
		ELSEIF sTourRange = "6" THEN
				sSQL = sSQL + " AND Year(TDateS) = '"&Year(Date())-1&"'"
	
		' --- Two calendar years ago if this year is nearly over otherwise three calendar years ago ---
		ELSEIF sTourRange = "7" THEN
				sSQL = sSQL + " AND Year(TDateS) = '"&Year(Date())-2&"'"
	
		END IF
	END IF

	IF StartMonth<>0 THEN
		sSQL = sSQL + " AND Month(TDateS) >= '"&StartMonth&"'"
	END IF

	IF EndMonth<>0 THEN
		sSQL = sSQL + " AND Month(TDateE) <= '"&EndMonth&"'"
	END IF

	IF sTourState <> "" AND LCASE(sTourState) <> "all" THEN sSQL = sSQL + " AND lower(TState) = '" & sqlclean(lcase(sTourState)) & "'"

	IF sTourRegion <> "" THEN sSQL = sSQL + " AND lower(right(left(ST.TournAppID,3),1)) = '" & sqlclean(lcase(sTourRegion)) & "'"

	IF sTourDate1 <> "" THEN sSQL = sSQL + " AND (TDateE >= '" & sTourDate1 & "' or TDateS >= '" & sTourDate1 & "')"

	IF sTourDate2 <> "" THEN sSQL = sSQL + " AND (TDateE <= '" & sTourDate2 & "' or TDateS <= '" & sTourDate2 & "')"


	IF process="register" OR process="viewreg" OR process="admcode" THEN sSQL = sSQL + " AND PayPalOK=1 AND PayPalAct<>''"
		
	sSQL = sSQL + " AND Deleted=0"	
	
  sSQL = sSQL + " ORDER BY TDateS"


	IF Session("adminmenulevel")>=50 AND TRIM(sShowSQL)<>"" THEN
			response.write("<br>"&sSQL)
			'response.end
	END IF

	set rs=Server.CreateObject("ADODB.recordset")
	rs.open sSQL, SConnectionToTRATable

	IF Session("adminmenulevel")>=50 AND TRIM(sShowSQL)<>"" THEN
			IF NOT rs.eof THEN response.write("<br><br>FOUND")
	END IF

END SUB







' ---------------------------------------------------------------------------------------------------------
    SUB PerformSQLQuery_Pre2009  ' ---------------------  BUILD SQL statement   -----------------------------------
' ---------------------------------------------------------------------------------------------------------

'IF Session("AdminMenuLevel")>=50 THEN
'	response.write("<br> Mark's TEMP stop")
'	response.write("<br> sTourRange="&sTourRange)

'	response.end

'END IF

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' 			*** IMPORTANT ***

' --- This module is only applicable for tournaments prior to 2010 

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++



' --- From Jim Meis 10-29-2007
' Class I (and N) ARE "traditional" AWSA/NCWSA classes.  NCWSA is supposed to use Class I, and Class I is expected by WSTIMS for collegiate events.
' Class I should only be used by NCWSA so I should probably remove it as an option on the AWSA sanction form.
' Classes I and N predate Grassroots, have different officials requirements, allow different officials work credits, and have different sanction fees.
' Classes I and N also require AWSA or NCWSA Region Sanction Approval which Grassroots technically does not.   
' If it has a traditional event, the Sports Division admin and HQ give approval as part of the traditional approval. 
'     If all the traditional requirements are in place any Grassroots program is automatically OK without much thought

' TEventSlalom, TEventTrick, and TEventJump -   Barefoot traditional tournaments (ABC) as users of those fields.  
' Users are:  AWS, NCW, ABC use all 3 and AKA uses TEventSlalom and TEventTrick

' THSClassF is the original Fun field.  It is and always was distinct from THSClassN.   Officials required are different etc.
' Current designation for fun 3 event is TEventF3ev=1.  More specifically this indicates a Grassroots event that the sponsor is characterizing
'    as 3 event type.  Could be sanctioned under any of the sports divisions.
' All the fields beginning with TEventF, except TEventFun, are the most recent Grassroots fields.
' Sponsor can offer multiple classes and skier can pick what level he wants to ski at - so THSClassR is 1 if R is offered and 0 if it is not.  


' --- From Jim Meis 10/28/2007
' In response to the question "why so many fields?"
' They came about because of the umpteen revisions to the "Fun", NWL, NBL, NSL, and Grassroots programs in the past 3 or 4 years.
' Started out with THSClassF, then added TEventFun to allow stand alone fun, then dropped THSClassF to separate FUN from 3 event 
'   and to allow fun to be sanctioned by other sports divisions, then added NSL, NWL, NBL, then added Grassroots, 
'   then dropped NSL, NWL, and NBL from sanction form but left it on the adverts whenever grassroots was selected 
'   for 3ev, Barefoot or wakeboard).  Latest directive says drop NSL, NWL, and NBL entirely, change the fun "events" 
'   already offered and add new ones


' --- From Jim Meis 3/1/2007 ---
' There are separate description fields in swift for traditional, fun, clinics, Wakeboard and Kneeboard events.  
' They have zero length strings if there is no description  (no matching events).
' When a sanction includes more than one of these categories you need to concatenate the description fields to get all the information.


' Tschedul.TDescription - AWS, ABC, or NCW standard events
' Tschedul.WDescription - Wakeboard standard events
' Tschedul.KDescription - Kneeboard standard events
' Tschedul.FDescription - Fun Events including NSL 
' Tschedul.CDescription - Clinics

' Tschedul.TStatus = 0     Application received
' Tschedul.TStatus = 1     Region approved
' Tschedul.TStatus = 2    USAWS Approved
 
' Tschedul.TPending   True until first save by an administrator - To publish must be TPending must be false and the other conditions below 
'     must be met.
' RegnSetup.ShowPSchedule - On Off switch for the entire schedule for a Region.
' RegnSetup.ShowGBLink  -  controls display of tournament schedule as pick list for sponsors. Generally set same as ShowPSchedule
' RegnSetup.GBPolicy = true if ad is allowed to be displayed before full Region approval of the sanction.  Necessary but not sufficient.
' Tschedul.TKitOKGuidebookAd  -   Set by Region Admin on each sanction application - Gives Region approval of content of the Ad.  Allows 
'     publication if GBPolicy is true and ShowPSchedule is true. True = Bit 1 False = Bit 0
 
' Tschedule.TPending is true by default - it is changed to false after the first review and save by an administrator.  Nothing should be 
'    displayed unless TPending is false (has received its first review and save).

' Some regions particularly Western Region do not want tournament information posted at all until the Guidebook is published.  Regions can 
'    toggle ShowPSchedule on and off in their Region Preferences. True = OK to show as long as the rest of the conditions are met.  
'    False means do not show under any conditions.
 
' ShowGBLink is related but only important for SWIFT - it determines if the tournament schedule is used as a pick list for sponsors revising 
'    tournaments or if they have to supply their tournAppID and Edit Code blind.
 
' GBPolicy -    Guidebook Policy determines if advertisement is allowed to display before the Region has given their sanction approval.
' Some regions require that the region part of the sanction process be complete (fees paid) and approved before displaying the advertisement.  
'    If guidebook = false then don't display unless TStatus >= 1
 
' Other Regions Others don't care and allow publication prior to region approval.   They only require that the ad itself be approved.
' In this case TStatus could be 0 or higher, Guidebook must be true, and TKitOKGuidebookAd must be true (ad itself has region approval)

' The  ShowReg, ShowAppointed, etc control display of specific parts of an ad - also set in region preferences. Some regions don't want 
'    registrar information published online until the guidebook is published on the theory that it levels the playing field for entries.
' ---------------------------------------------------------------------------------------------------------------------------------------



sSQL = "SELECT TOP 800 "
sSQL = sSQL + "ST.TournAppID, TName, ST.SptsGrpID, TDescription, WDescription, ST.FDescription, KDescription, CDescription" 
sSQL = sSQL + ", TSanction, TSanType, TDateE, TDateS, TCity, Tstate, Pending"
sSQL = sSQL + ", ShowPSched, TKitOKGuideBookAd, GBPolicy, TStatus, ShowRegistrar"
sSQL = sSQL + ", OK2Publish"
sSQL = sSQL + ", sClassC, sClassE, sClassL, sClassR, sClassX, sClassCash"
sSQL = sSQL + ", tClassC, tClassE, tClassL, tClassR, tClassX, tClassCash"
sSQL = sSQL + ", jClassC, jClassE, jClassL, jClassR, jClassX, jClassCash"
sSQL = sSQL + ", WWakeW, WSkateW, WSurfW"
sSQL = sSQL + ", Gr1AWSPulls, Gr1ABCPulls, Gr1USWPulls, Gr1AKAPulls, Gr1USHPulls, Gr1WSDPulls"
sSQL = sSQL + ", Gr2USH_FreeRidePulls, Gr2USH_JumpOutPulls, Gr2USH_BigAirPulls, Gr2USH_3TrickPulls"
sSQL = sSQL + ", Gr2AWS_SPulls, Gr2AWS_TPulls, Gr2ABC_SPulls, Gr2ABC_TPulls, Gr2USW_WPulls, Gr2USW_SkatePulls, Gr2USW_SurfPulls, Gr2USW_RailJamPulls" 
sSQL = sSQL + ", Gr2AKA_SPulls, Gr2AKA_TPulls, Gr2AKA_FreePulls, Gr2AKA_FlipPulls"
sSQL = sSQL + ", OLRDisplayStatus, UseOLReg, OLR_PD"


'		IF wb="on" AND (  rs("WWakeW")>0 OR rs("Gr2USW_WPulls")<>0 OR rs("Gr2USW_RailJamPulls")<>0 OR rs("Gr1USWPulls") <> 0 ) THEN sWB_Offered="Y"
'		IF ws="on" AND rs("WSkateW")>0 OR rs("Gr2USW_SkatePulls") THEN sWS_Offered="Y"
'		IF wu="on" AND rs("WSurfW")>0 OR rs("Gr2USW_SurfPulls") THEN sWU_Offered="Y"

'		IF bf="on" AND (  rs("SptsGrpID")="ABC" OR rs("Gr1ABCPulls")<>0  ) THEN sBF_Offered="Y" 
'		IF kb="on" AND (  rs("SptsGrpID")="AKA" OR rs("Gr2AKA_SPulls")<>0 OR rs("Gr2AKA_TPulls")<>0 OR rs("Gr2AKA_FreePulls")<>0 OR rs("Gr2AKA_FlipPulls")<>0 OR rs("Gr1AKAPulls")<>0  ) THEN sKB_Offered="Y"	
'		IF hy="on" AND (  rs("TEventFHF")<>0 OR rs("Gr2USH_FreeRidePulls")<>0 OR rs("Gr2USH_JumpOutPulls")<>0 OR rs("Gr2USH_BigAirPulls")<>0 OR rs("Gr2USH_3TrickPulls")<>0 OR rs("Gr1USHPulls")<>0  ) THEN sHY_Offered="Y"

'		IF da="on" AND (rs("Gr1WSDPulls")) THEN sDA_Offered="Y"
'		IF jd="on" AND (rs("JDClin")<>0) THEN sJD_Offered="Y"
'		IF ad="on" AND (rs("ADClin")<>0) THEN sAD_Offered="Y"




sSQL = sSQL + ", TRS.PayPalAct, TRS.PayPalOK"

sSQL = sSQL + ", ST.SptsGrpID AS sSptsGrpID, ST.TRegion AS STRegion"
sSQL = sSQL + ", ST.TEventNWL, ST.TEventNBL, ST.TEventNSL, ST.THSClassN, ST.THTClassN, ST.THJClassN"
sSQL = sSQL + ", TEventF3ev"
sSQL = sSQL + ", ST.TEventWake, ST.TEventWSkate, ST.TEventWSurf, ST.TEventFW"
sSQL = sSQL + ", ST.TEventSlalom, ST.TEventTrick, ST.TEventJump, ST.TEventFun"
sSQL = sSQL + ", ST.THSClassI, ST.THJClassI, ST.THTClassI"
sSQL = sSQL + ", ST.JDClin, ST.ADClin, ST.TEventFHF, ST.TEventFKB"

sSQL = sSQL + " FROM " &SanctionTableName&" AS ST"

sSQL = sSQL + " LEFT JOIN "&RegnSetupTableName&" AS RT ON ST.SptsGrpID = RT.SptsGrpID AND ST.TRegion = RT.TRegion"
sSQL = sSQL + " LEFT JOIN "&TRegSetupTableName&" AS TRS ON TRS.TournAppID = ST.TournAppID"






	sSQL = sSQL + " WHERE (1=2 "   ' --- This is the top of the bracket of all event inclusions ---

	IF sTourLevel="cash" THEN
		sSQL = sSQL +" OR (ST.THSClassCash<>0 OR ST.THTClassCash<>0 OR ST.THJClassCash<>0)"
	END IF

	IF sTourLevel="premier" OR sTourLevel="all" THEN

		' --- 3 Event Premier ---		
		IF sl="on" OR tr="on" OR ju="on" THEN 		

			' --- Top of AWS bracket "OR" ---
			' -----------------------------
			sSQL = sSQL + " OR (ST.SptsGrpID='AWS' AND (3=4" 	' --- Top of AWS stuff

			IF sl="on" THEN 
				sSQL = sSQL + " OR ST.TEventSlalom<>0"
			END IF
			IF tr="on" THEN 
				sSQL = sSQL + " OR ST.TEventTrick<>0"
			END IF
			IF ju="on" THEN
				sSQL = sSQL + " OR ST.TEventJump<>0"
			END IF
			sSQL = sSQL + "))"				' --- Bottom of AWS stuff ---
		END IF		

		' --- Wakeboard Premier ---
		IF wb="on" OR ws="on" OR wu="on" THEN 
			sSQL = sSQL + " OR (1=2"
			IF wb="on" THEN sSQL = sSQL + " OR ST.TEventWake<>0"
			IF ws="on" THEN sSQL = sSQL + " OR ST.TEventWSkate<>0"
			IF wu="on" THEN sSQL = sSQL + " OR ST.TEventWSurf<>0"
			sSQL = sSQL + ")" 	
		END IF	
	END IF

	IF sTourLevel="grass" OR sTourLevel="all" THEN
		
		sSQL = sSQL +" OR (5=6"  	'---- Open bracket Grass

		' --- Grassroots 3 Event ---
		IF sl="on" THEN 
			sSQL = sSQL + " OR ST.Gr2AWS_SPulls<>0 OR Gr1AWSPulls<>0" 
			sSQL = sSQL + " OR ST.THSClassN<>0"
			' --- TEventFun and THSClassF included for legacy Pre-2009 system ---
			sSQL = sSQL + " OR ST.TEventFun<>0 OR ST.THSClassF<>0 OR ST.TEventF3ev<>0"
		END IF
		IF tr="on" THEN 		
			sSQL = sSQL + " OR ST.Gr2AWS_TPulls<>0" 
			sSQL = sSQL + " OR ST.THTClassN<>0"
			' --- TEventFun and THTClassF included for legacy Pre-2009 system ---
			sSQL = sSQL + " OR ST.THTClassF<>0"
		END IF
		IF ju="on" THEN 		
			sSQL = sSQL + " OR ST.THJClassN<>0"
		END IF


		' --- Grassroots Wakeboard ---		
		IF wb="on" OR ws="on" OR wu="on" THEN 
			' --- Legacy from Pre-2009 system ---
			sSQL = sSQL + " OR (ST.TEventFW<>0 OR ST.TEventNWL<>0"
			IF wb="on" THEN 
				sSQL = sSQL + " OR Gr2USW_WPulls<>0 OR Gr2USW_RailJamPulls<>0 OR Gr1USWPulls<>0" 
			END IF
			IF ws="on" THEN 
				sSQL = sSQL + " OR Gr2USW_SkatePulls<>0"
			END IF 
			IF wu="on" THEN 
				sSQL = sSQL + " OR Gr2USW_SurfPulls<>0"
			END IF 
			sSQL = sSQL + ")" 	
		END IF

		sSQL = sSQL + ")" 		'---- Close bracket Grass


	END IF

	' --- Collegiate ---
	IF sTourLevel="collegiate" THEN 
		' --- 3 Event ---
		IF sl="on" OR tr="on" OR ju="on" THEN 		
			sSQL = sSQL + " OR (ST.SptsGrpID='NCW' AND (1=2 "
			IF sl="on" THEN sSQL = sSQL + " OR ST.TEventSlalom<>0 OR ST.TEventFun<>0"
			IF tr="on" THEN sSQL = sSQL + " OR ST.TEventTrick<>0"
			IF ju="on" THEN sSQL = sSQL + " OR ST.TEventJump<>0"
			sSQL = sSQL + "))" 
		END IF

		' --- Wakeboard ---		
		IF wb="on" OR ws="on" OR wu="on" THEN 
			sSQL = sSQL + " OR (ST.SptsGrpID='NCW' AND ST.TEventFW<>0 OR ST.TEventNWL<>0)" 
		END IF

	END IF


IF Session("AdminMenuLevel")>=50 THEN
	' response.write("SptsGrpID=")
END IF
	' --- Barefoot ---
	IF bf="on" THEN sSQL = sSQL + " OR (ST.SptsGrpID='ABC' OR ST.TEventNBL<>0 OR Gr1ABCPulls<>0 OR Gr2ABC_SPulls<>0 OR Gr2ABC_TPulls<>0)"

	' --- Kneeboard ---
	IF kb="on" THEN sSQL = sSQL + " OR (ST.SptsGrpID='AKA' OR ST.TEventFKB<>0) OR Gr1AKAPulls<>0 OR Gr2AKA_SPulls<>0 OR Gr2AKA_TPulls<>0 OR Gr2AKA_FreePulls<>0 OR Gr2AKA_FlipPulls<>0"


	' --- Hydrofoil ---
	IF hy="on" THEN 
		' --- Legacy from Pre-2009 ---
		sSQL = sSQL + " OR (ST.TEventFHF<>0"
		sSQL = sSQL + " OR ST.Gr2USH_FreeRidePulls<>0 OR ST.Gr2USH_JumpOutPulls<>0 OR ST.Gr2USH_BigAirPulls<>0 OR ST.Gr2USH_3TrickPulls<>0 OR ST.Gr1USHPulls<>0)"
	END IF


	' --- Clinic ---
	IF ad="on" THEN sSQL = sSQL + " OR ST.ADClin<>0"
	IF jd="on" THEN sSQL = sSQL + " OR ST.JDClin<>0"


	sSQL = sSQL + ")"    ' --- This is the bottom of the bracket of all event inclusions ---
	

	' --- Filters for highest homologation class ---


	HighClass=99
	IF sClass="R" AND (sl="on" OR tr="on" OR ju="on") THEN 
		sSQL = sSQL + " AND (THSClassR<>0 OR THTClassR<>0 OR THJClassR<>0)" 
		HighClass=1

	ELSEIF sClass="L" AND HighClass>1 AND (sl="on" OR tr="on" OR ju="on") THEN
		sSQL = sSQL + " AND (THSClassL<>0 OR THTClassL<>0 OR THJClassL<>0)" 
		HighClass=2

	ELSEIF sClass="E" AND HighClass>2 AND (sl="on" OR tr="on" OR ju="on") THEN
		sSQL = sSQL + " AND (THSClassE<>0 OR THTClassE<>0 OR THJClassE<>0)" 
		HighClass=3

	ELSEIF sClass="C" AND HighClass>3 AND (sl="on" OR tr="on" OR ju="on") THEN
		sSQL = sSQL + " AND (THSClassC<>0 OR THTClassC<>0 OR THJClassC<>0)" 
		HighClass=4

	ELSEIF sClass="N" AND HighClass>4 AND (sl="on" OR tr="on" OR ju="on") THEN
		sSQL = sSQL + " AND (THSClassN<>0 OR THTClassN<>0 OR THJClassN<>0)" 
		HighClass=5

	ELSEIF sClass="F" AND HighClass>5 AND (sl="on" OR tr="on" OR ju="on") THEN
		sSQL = sSQL + " AND (THSClassF<>0 OR THTClassF<>0 OR THJClassF<>0)" 
		HighClass=6

	ELSEIF sClass="F_O" AND (sl="on" OR tr="on" OR ju="on") THEN
		sSQL = sSQL + " AND TEventF3ev<>0" 
		HighClass=7

	END IF



	IF sTourRange <> "" AND sTourRange <> "0" THEN
        	IF sTourRange = "1" THEN
			sSQL = sSQL + " AND (ST.TDateE >= '" & Date() & "')"
		END IF

		' --- Ski Year defined as Latest in DivisionTable ---
		IF sTourRange = "2" THEN
		        set rsSelectFields=Server.CreateObject("ADODB.recordset")
            		rsSelectFields.open "Select Top 1 * FROM "&SkiYearTableName&" WHERE SkiYearID<>'1' ORDER BY BeginDate DESC", SConnectionToTRATable
			IF NOT rsSelectFields.eof THEN
				sSQL = sSQL + " AND (left(ST.TournAppID,2) = '" & right(right(TRIM(rsSelectFields("SkiYearName")),4),2) & "')"
			END IF
			rsSelectFields.close

		' --- Ski Year defined as SECOND latest in DivisionTable ---
		ELSEIF sTourRange = "3" THEN 
		        set rsSelectFields=Server.CreateObject("ADODB.recordset")
            		rsSelectFields.open "Select Top 2 * FROM "&SkiYearTableName&" WHERE SkiYearID<>'1' ORDER BY BeginDate DESC", SConnectionToTRATable
			IF NOT rsSelectFields.eof THEN
				rsSelectFields.movenext
				sSQL = sSQL + " AND (left(ST.TournAppID,2) = '" & right(right(TRIM(rsSelectFields("SkiYearName")),4),2) & "')"
			END IF
			rsSelectFields.close

		' --- Ski Year defined as THIRD latest in DivisionTable ---
		ELSEIF sTourRange = "4" THEN 
		        set rsSelectFields=Server.CreateObject("ADODB.recordset")
            		rsSelectFields.open "Select Top 3 * FROM "&SkiYearTableName&" WHERE SkiYearID<>'1' ORDER BY BeginDate DESC", SConnectionToTRATable
			IF NOT rsSelectFields.eof THEN
			  rsSelectFields.movenext
			  IF NOT rsSelectFields.eof THEN
				rsSelectFields.movenext
				IF NOT rsSelectFields.eof THEN
				  sSQL = sSQL + " AND (left(ST.TournAppID,2) = '" & right(right(TRIM(rsSelectFields("SkiYearName")),4),2) & "')"
			  	END IF
			  END IF
			END IF
			rsSelectFields.close

		' --- Current Calendar year if the year is nearly over otherwise last calendar year ---
		ELSEIF sTourRange = "5" THEN
'			IF month(date())>10 THEN 
				sSQL = sSQL + " AND Year(TDateS) = '"&Year(Date())&"'"
'			ELSE
'response.write("<br>Year(Date)="&Year(Date()))

'response.end
'				sSQL = sSQL + " AND Year(TDateS) = '"&Year(Date())-1&"'"
'			END IF

		' --- Last Calendar year if this year is nearly over otherwise two calendar years ago ---
		ELSEIF sTourRange = "6" THEN
			'IF month(date())>10 THEN 
				sSQL = sSQL + " AND Year(TDateS) = '"&Year(Date())-1&"'"
			'ELSE
			'	sSQL = sSQL + " AND Year(TDateS) = '"&Year(Date())-2&"'"
			'END IF

		' --- Two calendar years ago if this year is nearly over otherwise three calendar years ago ---
		ELSEIF sTourRange = "7" THEN
			'IF month(date())>10 THEN 
				sSQL = sSQL + " AND Year(TDateS) = '"&Year(Date())-2&"'"
			'ELSE
			'	sSQL = sSQL + " AND Year(TDateS) = '"&Year(Date())-3&"'"
			'END IF

		END IF
	END IF

	IF StartMonth<>0 THEN
		sSQL = sSQL + " AND Month(TDateS) >= '"&StartMonth&"'"
	END IF

	IF EndMonth<>0 THEN
		sSQL = sSQL + " AND Month(TDateE) <= '"&EndMonth&"'"
	END IF

	IF sTourState <> "" AND LCASE(sTourState) <> "all" THEN sSQL = sSQL + " AND lower(TState) = '" & sqlclean(lcase(sTourState)) & "'"

	IF sTourRegion <> "" THEN sSQL = sSQL + " AND lower(right(left(ST.TournAppID,3),1)) = '" & sqlclean(lcase(sTourRegion)) & "'"

	IF sTourDate1 <> "" THEN sSQL = sSQL + " AND (TDateE >= '" & sTourDate1 & "' or TDateS >= '" & sTourDate1 & "')"

	IF sTourDate2 <> "" THEN sSQL = sSQL + " AND (TDateE <= '" & sTourDate2 & "' or TDateS <= '" & sTourDate2 & "')"


	IF process="register" OR process="viewreg" OR process="admcode" THEN sSQL = sSQL + " AND PayPalOK=1"

	ShowCancelled="no"
	IF ShowCancelled = "no" THEN  sSQL = sSQL + " AND TStatus<>'3'" 
	
        sSQL = sSQL + " ORDER BY TDateS"


IF Session("adminmenulevel")>=50 AND TRIM(sShowSQL)<>"" THEN
response.write("<br>2 - "&Session("adminmenulevel"))
	response.write("<br>"&sSQL)
'	response.end
END IF


        set rs=Server.CreateObject("ADODB.recordset")
        rs.open sSQL, SConnectionToTRATable

END SUB









' -----------------------------
   Sub WriteHeaders(sTitle)
' -----------------------------

' Write Headers for DB Page

%>


<TABLE BORDER="0" CELLPADDING="6" CELLSPACING="0" WIDTH="100%" BGCOLOR="#C0C0C0" BORDERCOLOR="#C0C0C0" BORDERCOLORDARK="#C0C0C0" BORDERCOLORLIGHT="#C0C0C0" >
<tr>
<td align="Left"><Font Face="courier" COLOR="#000000" SIZE="4"><B><% Response.Write(sTitle) %></B></FONT></td>
</tr>
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
 Response.Write(" <a href=" & chr(34) & ThisPage & "?div=" & DivSelected & "&ranknum=" & RankNum & "&event=" & EventSelected & "&currentpage=" &  i  & "&action=" & sAction & chr(34) & ">" & i & "</a>")
h = h +1
next
IF h = 0 THEN h = 1
Response.Write("<BR><Small>Page " & currentPage & " of  " & h & "</SMALL></center><BR><BR>")
end sub
%>
