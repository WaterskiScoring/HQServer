<!--#include file="settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include16.asp"-->
<!--#include virtual="/rankings/tools_Definitions.asp"-->
<!--#include virtual="/rankings/tools_leagues.asp"-->
<!--#include virtual="/rankings/tools_mobile_version.asp"-->
<%



Dim ThisFileName



Dim sTolerance, sKPH, sMPH, sDisplay, sSeg0, sSeg1, sSeg2, sSeg3, sSeg4, sSeg5, sSeg6



ThisFileName = "view-TimingCharts_m.asp"



OpenState="timing_chart"

DisplayHeadOpenBodyAndBannerTags OpenState


GetTimingChartOutput




' --- Writes the Closing tags for HTML - in tools_mobile_version.asp ---
DisplayCloseBodyAndHTMLTags




' ---------------------------------------------------------------------------------------------
' --- Bottom of MAIN PROGRAM ---
' ---------------------------------------------------------------------------------------------









' -----------------------
  SUB GetTimingChartOutput
' -----------------------  


%>
<div id="myteamlisting" style="padding:0px; border:0px solid white;">
	<form method="post">
		<input type="hidden" name="sWatchMemberIDs" value="<%=sWatchMemberIDs%>">
		<input type="hidden" name="sMemberID" value="<%=sMemberID%>">
	<div style="width:100%; margin-top:10px; padding-left:0px; text-align:left;">		
			<span class="span95" style="margin-left:0px; padding-left:0px; text-align:center; font-size:14px; color:yellow; border:0px solid white;">Timing Charts - Record & Standard Tolerance</span> 
	</div>	

		<div style="width:100%; margin-top:0px; padding-left:0px; text-align:center; display:inline-block" >
		<%
			
			' BuildStatsMemberIDDropdown 
		
		%>
		
		</div>			
	</form>
	<div class="scroll" style="margin-top:5px; padding:0px; margin-left:0px; border:0px solid white;">
		<%   

		' --- Displays SUMMARY Information ---
		RunTimingChartQuery

		IF NOT rs.eof THEN 
				LoopThruTimingChart
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
  SUB LoopThruTimingChart
' --------------------------
 

TimingCount = 1

TabColor = tcolor01
DisplayTimingTab "Record Tolerance", TabColor

DO WHILE NOT rs.eof

		GetTimingLine		

		IF rs("Tolerance") = "S" AND TimingCount = 1 THEN
				TimingCount = 2
				
				DisplayTimingBottomLine
				TabColor = tcolor02
				response.write("<br>")
				DisplayTimingTab "Standard Tolerance", TabColor 
		END IF		 
		DisplayTimingDataLine
		
		rs.movenext
		' TimingCount = TimingCount + 1
	
LOOP

DisplayTimingBottomLine

END SUB










' --------------------------------
  SUB GetTimingLine
' --------------------------------  

sTolerance = rs("tolerance")  
sKPH = rs("KPH")
sMPH = rs("MPH")
sDisplay = rs("Display")
sSeg0 = FormatNumber(rs("Seg0"),2)
sSeg1 = FormatNumber(rs("Seg1"),2)
sSeg2 = FormatNumber(rs("Seg2"),2)
sSeg3 = FormatNumber(rs("Seg3"),2)
sSeg4 = FormatNumber(rs("Seg4"),2)
sSeg5 = FormatNumber(rs("Seg5"),2)
sSeg6 = FormatNumber(rs("Seg6"),2)

END SUB  




' -------------------------------------------
  SUB DisplayTimingTab (Tolerance, TabColor)
' -------------------------------------------


%> 
  <div class="tabrankings" style="width:97%; height:38px; background-color:<% =TabColor %>; padding:0px 0px 0px 5px; margin:0px 0px 0px 0px;">
		<span class="span90" style="font-size:12pt; font-weight:bold;">Speed (KPH/MPH) - <%= Tolerance %></span>
		<br>
		<span class="span10" style="text-align:left; font-size:9pt; margin-top:5px; font-weight:normal;">&nbsp;</span>
		<span class="span10" style="text-align:right; font-size:9pt; font-weight:normal;">Seg</span>
		<span class="span10" style="text-align:right; font-size:9pt; font-weight:normal;">0</span>
		<span class="span10" style="text-align:right; font-size:9pt; font-weight:normal;">1</span>		
		<span class="span10" style="text-align:right; font-size:9pt; font-weight:normal;">2</span>
		<span class="span10" style="text-align:right; font-size:9pt; font-weight:normal;">3</span>
		<span class="span10" style="text-align:right; font-size:9pt; font-weight:normal;">4</span>		
		<span class="span10" style="text-align:right; font-size:9pt; font-weight:normal;">5</span>
		<span class="span10" style="text-align:right; font-size:9pt; font-weight:normal;">6</span>
	</div>
<%	

END SUB




' --------------------------
  SUB DisplayTimingDataLine
' --------------------------


sThisSpeed = ""
IF sDisplay="Ideal" THEN
		sThisSpeed = "<b>" & sKPH & "</b>/" & FormatNumber(sMPH,1)			
END IF
IF sDisplay="Fast" THEN response.write("<hr style=""width:97%; padding:0px 0px 0px 3px; margin:0px 2px 0px 2px; height:2px; background-color:#FFFFFF;"">")
%> 
  <div class="rankingsbody" style="width:97%; border:0px solid white; padding:0px 3px 0px 2px; margin:0px 0px 0px 2px;">
		<span class="span10" style="height:9px; text-align:left; font-size:9pt; margin-top:5px; font-weight:normal;"><% =sThisSpeed %></span>
		<span class="span10" style="height:9px; text-align:right; font-size:9pt; font-weight:normal;"><%= sDisplay %></span>
		<span class="span10" style="height:9px; text-align:right; font-size:9pt; font-weight:normal;"><%= sSeg0 %></span>
		<span class="span10" style="height:9px; text-align:right; font-size:9pt; font-weight:normal;"><%= sSeg1 %></span>
		<span class="span10" style="height:9px; text-align:right; font-size:9pt; font-weight:normal;"><%= sSeg2 %></span>
		<span class="span10" style="height:9px; text-align:right; font-size:9pt; font-weight:normal;"><%= sSeg3 %></span>
		<span class="span10" style="height:9px; text-align:right; font-size:9pt; font-weight:normal;"><%= sSeg4 %></span>
		<span class="span10" style="height:9px; text-align:right; font-size:9pt; font-weight:normal;"><%= sSeg5 %></span>
		<span class="span10" style="height:9px; text-align:right; font-size:9pt; font-weight:normal;"><%= sSeg6 %></span>
	</div>
<%	

END SUB










' -------------------------------
  SUB DisplayTimingBottomLine
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
  SUB RunTimingChartQuery
' --------------------


sSQL = " SELECT *"
sSQL = sSQL + " FROM [usawsrank].[Timing_Charts]"
sSQL = sSQL + " ORDER BY Tolerance, KPH DESC"  

Set rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable

END SUB



%>