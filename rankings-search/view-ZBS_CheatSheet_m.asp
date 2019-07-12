<!--#include file="settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include16.asp"-->
<!--#include virtual="/rankings/tools_Definitions.asp"-->
<!--#include virtual="/rankings/tools_leagues.asp"-->
<!--#include virtual="/rankings/tools_mobile_version.asp"-->
<%



Dim ThisFileName



Dim sKPH, sMPH, sL_2300, sL_1825, sL_1600, sL_1425, sL_1300, sL_1200, sL_1125, sL_1075, sL_1025, sL_0975, sL_0950, sL_0925



ThisFileName = "view-ZBS_CheatSheet_m.asp"



OpenState="zbs_cheat"

DisplayHeadOpenBodyAndBannerTags OpenState


GetZBSOutput




' --- Writes the Closing tags for HTML - in tools_mobile_version.asp ---
DisplayCloseBodyAndHTMLTags




' ---------------------------------------------------------------------------------------------
' --- Bottom of MAIN PROGRAM ---
' ---------------------------------------------------------------------------------------------









' -----------------------
  SUB GetZBSOutput
' -----------------------  


%>
<div id="myteamlisting" style="padding:0px; border:0px solid white;">
	<form method="post">
		<input type="hidden" name="sWatchMemberIDs" value="<%=sWatchMemberIDs%>">
		<input type="hidden" name="sMemberID" value="<%=sMemberID%>">
	<div style="width:100%; margin-top:10px; padding-left:0px; text-align:left;">		
			<span class="span95" style="margin-left:0px; padding-left:0px; text-align:center; font-size:14px; color:yellow; border:0px solid white;">ZBS Slalom Scoring Cheat Sheet</span> 
	</div>	

		<div style="width:100%; margin-top:0px; padding-left:0px; text-align:center; display:inline-block" >
		<%
			
			' BuildStatsMemberIDDropdown 
		
		%>
		
		</div>			
	</form>
	<div class="scroll" style="margin-top:5px; padding:0px; margin-left:0px; height:435px; border:0px solid white;">
		<%   

		' --- Displays SUMMARY Information ---
		RunZBSQuery

		IF NOT rs.eof THEN 
				LoopThruZBS
		END IF
		

		
		%>
	</div> <! -- Bottom of scroll box -- ->
</div> <! -- Bottom of div for hiding and displaying -- ->
<%



END SUB




' -----------------------------
  SUB DisplayNoUserSetScreen
' -----------------------------

'response.write("</div><div style=""color:red;"">HERE lin 516</div>")
%>
<div class="tabrankings" style="height:50px; margin-top:50px; padding-top:4px; background-color:<%=DefineLevelcolor%>; text-color:<%=Textcolor1%>; font-size:16px;">
	<span class="span90" style="color:yellow; text-align:center;"><b>This function may not be accessed until an Authorized User has been set up on this device.</b></span>
</div>
<div class="span95" style="margin-top:50px; text-align:center;">
	<form action="/rankings/<%= MenuFileName %>" method="post">
		<input type="submit" value="Return to Main Menu" title="Go to Main Menu" style="font-size:12pt; size:12em;">
	</form>
</div>
<%

END SUB



' -----------------------------------------
  SUB DisplayNoListingFound (whichdisplay) 
' ------------------------------------------  

'response.write("</div><div style=""color:red;"">HERE lin 516</div>")
%>
<div class="rankingsbody" style="width:97%; border:0px solid white; padding:0px 3px 0px 2px; margin:0px 0px 0px 2px;>
	<span class="span90" style="color:black; text-align:center; font-size:14pt; font-weight:bold; margin:0px 0px 0px 30px;">No <%= whichdisplay %> Found for these Settings</span>
</div>
<%

END SUB



' --------------------------
  SUB LoopThruZBS
' --------------------------
 

ZBSCount = 1


DisplayZBSTab

DO WHILE NOT rs.eof

		GetZBSLine		

		DisplayZBSDataLine
		
		rs.movenext
		ZBSCount = ZBSCount + 1
	
LOOP

DisplayZBSBottomLine

END SUB










' --------------------------------
  SUB GetZBSLine
' --------------------------------  
  
sKPH = rs("KPH")
sMPH = rs("MPH")
sL_2300 = rs("L_2300")
sL_1825 = rs("L_1825")
sL_1600 = rs("L_1600")
sL_1425 = rs("L_1425")
sL_1300 = rs("L_1300")
sL_1200 = rs("L_1200")
sL_1125 = rs("L_1125")
sL_1075 = rs("L_1075")
sL_1025 = rs("L_1025")
sL_0975 = rs("L_0975")
sL_0950 = rs("L_0950")
sL_0925 = rs("L_0925")

END SUB  




' --------------------------
  SUB DisplayZBSDataLine
' --------------------------

%> 
  <div class="rankingsbody" style="width:97%; border:0px solid white; padding:0px 3px 0px 2px; margin:0px 0px 0px 2px;">
		<span class="span10" style="text-align:left; font-size:9pt; margin-top:5px; font-weight:normal;"><b><% =sKPH %></b>/<% = FormatNumber(sMPH,1) %></span>
		<span class="span7" style="text-align:right; font-size:9pt; font-weight:normal;"><%= sL_2300 %></span>
		<span class="span7" style="text-align:right; font-size:9pt; font-weight:normal;"><%= sL_1825 %></span>
		<span class="span7" style="text-align:right; font-size:9pt; font-weight:normal;"><%= sL_1600 %></span>
		<span class="span7" style="text-align:right; font-size:9pt; font-weight:normal;"><%= sL_1425 %></span>
		<span class="span7" style="text-align:right; font-size:9pt; font-weight:normal;"><%= sL_1300 %></span>
		<span class="span7" style="text-align:right; font-size:9pt; font-weight:normal;"><%= sL_1200 %></span>
		<span class="span7" style="text-align:right; font-size:9pt; font-weight:normal;"><%= sL_1125 %></span>
		<span class="span7" style="text-align:right; font-size:9pt; font-weight:normal;"><%= sL_1075 %></span>
		<span class="span7" style="text-align:right; font-size:9pt; font-weight:normal;"><%= sL_0975 %></span>
		<span class="span7" style="text-align:right; font-size:9pt; font-weight:normal;"><%= sL_0950 %></span>										
		<span class="span7" style="text-align:right; font-size:9pt; font-weight:normal;"><%= sL_0925 %></span>			
	</div>
<%	

END SUB








' -------------------------
  SUB DisplayZBSTab
' -------------------------

TabColor = tcolor01
%> 
  <div class="tabrankings" style="width:97%; height:38px; background-color:<% =TabColor %>; padding:0px 0px 0px 5px; margin:0px 0px 0px 0px;">
		<span class="span80" style="font-size:12pt; font-weight:bold;">Speed (KPH/MPH) & Line Length</span>
		<br>
		<span class="span10" style="text-align:left; font-size:9pt; margin-top:5px; font-weight:normal;">&nbsp;</span>
		<span class="span7" style="text-align:right; font-size:9pt; font-weight:normal;">2300</span>
		<span class="span7" style="text-align:right; font-size:9pt; font-weight:normal;">1825</span>
		<span class="span7" style="text-align:right; font-size:9pt; font-weight:normal;">1600</span>		
		<span class="span7" style="text-align:right; font-size:9pt; font-weight:normal;">1425</span>
		<span class="span7" style="text-align:right; font-size:9pt; font-weight:normal;">1300</span>
		<span class="span7" style="text-align:right; font-size:9pt; font-weight:normal;">1200</span>		
		<span class="span7" style="text-align:right; font-size:9pt; font-weight:normal;">1125</span>
		<span class="span7" style="text-align:right; font-size:9pt; font-weight:normal;">1075</span>		
		<span class="span7" style="text-align:right; font-size:9pt; font-weight:normal; padding-left:3px;">975</span>
		<span class="span7" style="text-align:right; font-size:9pt; font-weight:normal; padding-left:1px;">950</span>		
		<span class="span7" style="text-align:right; font-size:9pt; font-weight:normal; padding-left:1px;">925</span>
	</div>
<%	

END SUB



' -------------------------------
  SUB DisplayZBSBottomLine
' -------------------------------  
' style="width:97%; border:0px solid white; padding:0px 3px 0px 2px; margin:0px 0px 0px 2px;"
' style="width:97%; border:0px solid white; padding:0px 2px 0px 3px; margin:0px 0px 0px 2px;"
' -- Orig style="width:97%; background-color:#FFFFFF; height:10px; margin:0px 0px 0px 2px; padding:0px 3px 0px 5px;"
%>
<div class="rankingsbottom" style="width:97%; background-color:#FFFFFF; border-right:1px solid #FFFFFF; height:10px; padding:0px 2px 0px 2px; margin:0px 0px 0px 2px;">
		<span class="span100">&nbsp;</span>
</div>
<%

END SUB



' --------------------
  SUB RunZBSQuery
' --------------------


sSQL = " SELECT *"
sSQL = sSQL + " FROM [usawsrank].[ZBS_Cheat_Sheet]"
sSQL = sSQL + " ORDER BY KPH DESC"  

Set rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable

END SUB



%>