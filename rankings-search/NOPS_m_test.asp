<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_leagues.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->
<!--#include virtual="/rankings/tools_registration.asp"-->
<!--#include virtual="/rankings/Tools_TournamentListQuery.asp"-->
<!--#include virtual="/rankings/tools_mobile_version_TEST.asp"-->
<%



' ------------------------------------------------
' --- Dimensions variables used in this module ---
' ------------------------------------------------

Dim ThisFileName



Dim DivArray, SlalomNOPSArray, TrickNOPSArray, JumpNOPSArray
Dim SlalomExpArray, TrickExpArray, JumpExpArray, OverPtsBySArray


MenuItemPath = "images\icons\"
FAQImagePath = "images\mobile_faq\"
ThisFileName = "NOPS_m.asp"
ThisSitePath = "/rankings"

' --- Names related programs for linking ---
SearchFileName = "search-memberHQ.asp"
RankingsMobileFilename="view-standings_m.asp"
TournamentsMobileFilename="view-tournaments_m.asp"
TeamsMobileRankingFilename="View-vteamstatus_m.asp"
LocalVarFileName="User_Set.asp"
MenuFileName = "mainmenu_m.asp"
MyStatsFilename = "view-mystats_m.asp"







' --- Displays the html, head and opening body tag ---
OpenState="nops"
DisplayHeadOpenBodyAndBannerTags OpenState



' --- Reads NOPS values from Division table ---
ReadNOPSFromTable

' --- Displays the Preliminary NOPS page --
DisplayNOPSPage



' --- Writes the Closing tags for HTML ---
DisplayCloseBodyAndHTMLTags






' ---------------------------------------------------
' --- BOTTOM OF MAIN CODE ---
' ---------------------------------------------------





' ----------------------
  SUB ReadNOPSFromTable
' ----------------------

sRunByWhat = "National"
sSkiYearSelected = Request("sSkiYear")

sSkiYearSelected=1

sSQL = "SELECT Div, Over_S, Over_T, Over_J, OverExp_S, OverExp_T, OverExp_J, OverPtsBy_S"
sSQL = sSQL + " FROM "&DivisionsTableName
sSQL = sSQL + " WHERE SkiYearID = "&sSkiYearSelected
SELECT CASE sRunByWhat
 		CASE "National"
				sSQL = sSQL + " AND lower(left(Div,1)) in ('b','g','m','w','o')"
  	CASE "NCWSA"
				sSQL = sSQL + " AND lower(left(Div,1)) = 'c'"
END SELECT
sSQL = sSQL + " ORDER BY Div"

SET rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable

'response.write("</div></div><div style='color:black; background-color:white;'>"&sSQL)

DivArray = ""
'SlalomNOPSArray = ""
'TrickNOPSArray = ""
'JumpNOPSArray = ""

ArrayIndex=0

DO WHILE NOT rs.eof
		Div = rs("Div")
		Over_S = rs("Over_S")
		Over_T = rs("Over_T")
		Over_J = rs("Over_J")		
		OverExp_S = rs("OverExp_S")		
		OverExp_T = rs("OverExp_T")		
		OverExp_J = rs("OverExp_J")				
		OverPtsBy_S = rs("OverPtsBy_S")

		'IF ArrayIndex=0 THEN DivArray = "'"&Div&"'" ELSE DivArray = DivArray & ",'" &Div&"'"
		'IF ArrayIndex=0 THEN SlalomNOPSArray = "'"&Over_S&"'" ELSE SlalomNOPSArray = SlalomNOPSArray & ",'" &Over_S&"'"
		'IF ArrayIndex=0 THEN TrickNOPSArray = "'"&Over_T&"'" ELSE TrickNOPSArray = TrickNOPSArray & ",'" &Over_T&"'"
		'IF ArrayIndex=0 THEN JumpNOPSArray = "'"&Over_J&"'" ELSE JumpNOPSArray = JumpNOPSArray & ",'" &Over_J&"'"		

		IF ArrayIndex=0 THEN DivArray = "(XX," &Div ELSE DivArray = DivArray & "," &Div
		IF ArrayIndex=0 THEN SlalomNOPSArray = "(0,"&Over_S ELSE SlalomNOPSArray = SlalomNOPSArray & "," &Over_S
		IF ArrayIndex=0 THEN TrickNOPSArray = "(0,"&Over_T ELSE TrickNOPSArray = TrickNOPSArray & "," &Over_T
		IF ArrayIndex=0 THEN JumpNOPSArray = "(0,"&Over_J ELSE JumpNOPSArray = JumpNOPSArray & "," &Over_J

		IF ArrayIndex=0 THEN SlalomExpArray = "(0,"&OverExp_S ELSE SlalomExpArray = SlalomExpArray & "," &OverExp_S
		IF ArrayIndex=0 THEN TrickExpArray = "(0,"&OverExp_T ELSE TrickExpArray = TrickExpArray & "," &OverExp_T
		IF ArrayIndex=0 THEN JumpExpArray = "(0,"&OverExp_J ELSE JumpExpArray = JumpExpArray & "," &OverExp_J					
		
		IF ArrayIndex=0 THEN OverPtsBySArray = "(0"&OverPtsBy_S ELSE OverPtsBySArray = OverPtsBySArray & "," &OverPtsBy_S				


		
		rs.MoveNEXT
		ArrayIndex = ArrayIndex + 1
		
LOOP

DivArray = DivArray& ")"

SlalomNOPSArray = SlalomNOPSArray& ")"
TrickNOPSArray = TrickNOPSArray& ")"
JumpNOPSArray = JumpNOPSArray& ")"

SlalomExpArray = SlalomExpArray& ")"
TrickExpArray = TrickExpArray& ")"
JumpExpArray = JumpExpArray& ")"

OverPtsBySArray = OverPtsBySArray& ")"

'response.write("</div></div><div style='color:red;'><br><br>DivArray = "&DivArray)
'response.write("</div><div style='color:red;'><br>JumpNOPSArray = "&JumpNOPSArray)
'response.write("</div><div style='color:red;'><br>JumpExpArray = "&JumpExpArray)
' response.write("</div><div style='color:red;'><br>OverPtsBySArray = "&OverPtsBySArray)
' JumpNOPSArray = (999,143,191,999,120,154,218,216,198,179,157,143,112,96,98,47,22,215,155,250,184,157,158,155,133,108,95,69,57,45,32,32)
' JumpExpArray = (0,0.669,1.148,0,0.738,0.796,0.805,1.184,1.02,1.227,1.062,1.31,1.39,0.827,0.673,1.582,1.674,1.266,0.976,1.022,0.904,0.663,0.991,0.956,0.768,0.976,1.348,1.533,1.934,2.645,9.401,5.409)
' OverPtsBySArray = (024,8,8,40,8,8,8,8,8,8,8,8,8,8,8,8,8,8,8,8,8,8,8,8,8,8,8,8,8,8,8,8)


rs.close




END SUB


' ---------------------
  SUB DisplayNOPSPage
' ---------------------

' --- Retains values by requesting in case Ski Year is changed after inputting scores
RawScore_S = Request("RawScore_S")
RawScore_T = Request("RawScore_T")
RawScore_J = Request("RawScore_J")

%>
<div id="createrankingsfilters" class="errorbox" style="height:440px; margin:7px 0px 0px 0px; padding-right:0px; padding-left:10px;">
	<form action="<%=ThisSitePath%>/<%=ThisFileName%>" method="post">
		<input type="hidden" id="DivArray" name="DivArray" value="<%= DivArray %>">
		<input type="hidden" id="SlalomNOPSArray" name="SlalomNOPSArray" value="<%= SlalomNOPSArray %>">
		<input type="hidden" id="TrickNOPSArray" name="TrickNOPSArray" value="<%= TrickNOPSArray %>">
		<input type="hidden" id="JumpNOPSArray" name="JumpNOPSArray" value="<%= JumpNOPSArray %>">
		<input type="hidden" id="SlalomExpArray" name="SlalomExpArray" value="<%= SlalomExpArray %>">
		<input type="hidden" id="TrickExpArray" name="TrickExpArray" value="<%= TrickExpArray %>">
		<input type="hidden" id="JumpExpArray" name="JumpExpArray" value="<%= JumpExpArray %>">				
		<input type="hidden" id="OverPtsBySArray" name="OverPtsBySArray" value="<%= OverPtsBySArray %>">				

	<div style="width:96%; margin-top:10px; padding-left:10px; text-align:left; border:0px solid red;">		
			<span class="span95" style="margin-left:0px; padding-left:0px; text-align:center; font-size:16px; color:yellow; border:0px solid white;">NOPS Overall Calculator</span> 
	</div>
		<div style="margin:15px 0px 0px 0px; padding:0px 0px 0px 0px; border:0px solid white; text-align:left;">
			<span class="span20" style="text-align:right;">Ski Year:</span>
			<span class="span70" style="text-align:left;">
				<%
				BuildSkiYearDropDown
				%>
			</span>	
		</div>
		<div style="margin:10px 0px 0px 0px; padding:0px 0px 0px 0px; border:0px solid white; text-align:left;">
			<span class="span20" style="text-align:right;">Div:</span>
			<span class="span70" style="text-align:left;">
				<%
				BuildDivisionDropDownNOPS
				%>
			</span>	
		</div>
		
		<div style="width:96%; margin:25px 0px 0px 0px; padding:0px 0px 0px 0px; border:0px solid white; text-align:left;">
			<span class="span20" style="white; text-align:right; color:yellow; border:0px solid red;" >Event</span>
			<span class="span25" style="text-align:center; color:yellow; border:0px solid white;">&nbsp;&nbsp;&nbsp;Input</span>
			<span class="span20" style="text-align:center; color:yellow; margin-left:10px;">Record</span>
			<span class="span20" style="text-align:center; color:yellow; margin-left:10px;">NOPS</span>
		</div>	
		<div style="width:96%; margin:0px 0px 0px 0px; padding:0px 0px 0px 0px; border:0px solid white; text-align:left;">
			<span class="span20" style="text-align:right; border:0px solid white;" >Slalom</span>
			<span class="span25" style="text-align:right;">
				<input type="Tel" id="RawScore_S" name="RawScore_S" value="<%=RawScore_S%>" style="font-size:12pt; width:50px; text-align:right;">
			</span>
			<span class="span20" style="text-align:right;">
				<input type="text" class="textbox_hidden_banner" id="Record_S" name="Record_S" value="" style="font-size:12pt; color:yellow; width:50px; text-align:right; border:1px solid white;">
			</span>
			<span class="span20" style="text-align:right; margin-left:10px;">
				<input type="text" class="textbox_hidden_banner" id="NOPS_S" name="NOPS_S" value="" style="font-size:12pt; color:yellow; width:60px; text-align:right; border:1px solid white;">
			</span>
		</div>	
		<div style="width:96%; margin:15px 0px 0px 0px; padding:0px 0px 0px 0px; border:0px solid white; text-align:left;">
			<span class="span20" style="text-align:right;">Trick</span>
			<span class="span25" style="text-align:right;">
				<input type="Tel" id="RawScore_T" name="RawScore_T" value="<%=RawScore_T%>" style="font-size:12pt; width:50px; text-align:right;">
			</span>
			<span class="span20" style="text-align:right;">
				<input type="text" class="textbox_hidden_banner" id="Record_T" name="Record_T" value="" style="font-size:12pt; color:yellow; width:50px; text-align:right; border:1px solid white;">
			</span>
			<span class="span20" style="text-align:right; margin-left:10px;">
				<input type="text" class="textbox_hidden_banner" id="NOPS_T" name="NOPS_T" value="" style="font-size:12pt; color:yellow; width:60px; text-align:right; border:1px solid white;">
			</span>
		</div>	
		<div style="width:96%;  margin:15px 0px 0px 0px; padding:0px 0px 0px 0px; border:0px solid white; text-align:left;">
			<span class="span20" style="text-align:right;">Jump</span>
			<span class="span25" style="text-align:right;">
				<input type="Tel" id="RawScore_J" name="RawScore_J" value="<%=RawScore_J%>" style="font-size:12pt; width:50px; text-align:right;">
			</span>
			<span class="span20" style="text-align:right;">
				<input type="text" class="textbox_hidden_banner" id="Record_J" name="Record_J" value="" style="font-size:12pt; color:yellow; width:50px; text-align:right; border:1px solid white;">
			</span>
			<span class="span20" style="text-align:right; margin-left:10px;">
				<input type="text" class="textbox_hidden_banner" id="NOPS_J" name="NOPS_J" value="" style="font-size:12pt; color:yellow; width:60px; text-align:right; border:1px solid white;">
			</span>
		</div>	
		<div style="width:96%;  margin:15px 0px 0px 0px; padding:0px 0px 0px 0px; border:0px solid white; text-align:left;">
			<span class="span20" style="text-align:right;">Overall</span>
			<span class="span25" style="text-align:right;">&nbsp;</span>
			<span class="span20" style="text-align:right;">&nbsp;</span>
			<span class="span25" style="text-align:right; margin-left:15px;">
				<input type="text" class="textbox_hidden_banner" id="NOPS_O" name="NOPS_O" value="" style="font-size:12pt; color:yellow; width:60px; text-align:right; border:1px solid white;">
			</span>
		</div>	
		<div style="width:96%; margin:35px 0px 0px 0px; padding:0px 0px 0px 10px; text-align:left; border:0px solid red;">		
			<span class="span95" style="margin-left:0px; padding-left:0px; text-align:center; font-size:10pt; color:#FFFFFF; border:0px solid white;">Enter SCORES and press 'Recalculate'</span> 
		</div>

		<div style="height:50px; margin:15px 0px 0px 0px;">
			<input type="button" id="recalcNOPS" name="recalcNOPS" value="Recalculate" style="width:9em; font-size:12pt;" onclick="Javascript:UpdateNOPSField();">
		</div>
	<form>
</div>
<%


END SUB



' -------------------------------
  SUB BuildDivisionDropDownNOPS
' -------------------------------

sRunByWhat = "National"
DivSelected = Request("DivSelected")

	%>
	<select id='DivSelected' name='DivSelected' style="width:12em; font-size:12pt"><%

	SET rsSelectFields=Server.CreateObject("ADODB.recordset")
	sSQL = "SELECT DISTINCT Div, Div_name FROM "&DivisionsTableName&" AS DT"

	SELECT CASE sRunByWhat
  		CASE "National"
					sSQL = sSQL + " WHERE lower(left(Div,1)) in ('b','g','m','w','o')"
	  	CASE "NCWSA"
					sSQL = sSQL + " WHERE lower(left(Div,1)) = 'c'"
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
  		response.write("<option value =""None"" selected>None</option>")
	END IF

	rsSelectFields.close %>
	</select>
	<%


END SUB




' ---------------------------
  SUB BuildSkiYearDropDown
' ---------------------------

sRunByWhat="National"
SkiYearIDSelected = Request("SkiYearIDSelected")
	%>
	<SELECT id='SkiYearIDSelected' name='SkiYearIDSelected' style="width:12em; font-size:12pt;" onchange="submit()"><%

		
		sSQL = "SELECT DISTINCT SkiYearID, SkiYearName"
		sSQL = sSQL + ", CASE WHEN SkiYearID=1 THEN 1 ELSE 999-SkiYearID END AS MyOrder"
		sSQL = sSQL + " FROM " &SkiYearTableName&" AS SY"
		sSQL = sSQL + " ORDER BY CASE WHEN SkiYearID=1 THEN 1 ELSE 999-SkiYearID END"

		' --- NCWSA does not display 12 Month Rankings
		IF sRunByWhat="NCWSA" THEN
				sSQL = sSQL + " WHERE SY.SkiYearID <> 1"
		END IF

		SET rsSelectFields=Server.CreateObject("ADODB.recordset")
		rsSelectFields.open sSQL, SConnectionToTRATable

		' -- Loads dropdown and sets default to Session("SkiYear")
		DO WHILE NOT rsSelectFields.eof

			IF TRIM(rsSelectFields("SkiYearID")) = SkiYearIDSelected THEN
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

	rsSelectFields.close 
	%>
	</select>
	<%
  

END SUB


%>

