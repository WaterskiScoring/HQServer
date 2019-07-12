<% '---This is the menu for Adminstrative Level login on TRA --- 

Dim adminmenulevel, aml, usg, admcode
adminmenulevel=Session("adminmenulevel")
aml=Session("adminmenulevel")
usg=Session("UserSptsGrpID")
admcode=Session("AdminCode")

%>

	<p><a href="javascript:toggle('scores_subnav')">Tournament Scores</a></p>
   <div class="subnav" id="scores_subnav" style="display:none;">
    <a href="javascript:toggle('ScrByTour_subnav')">By Tournament </a>
    <div class="subnav" id="ScrByTour_subnav" style="display:none;">
     <a href="/rankings/view-scoreshq_AKA.asp?pvar=ByTour&sTourSportsGroup=AWS&rid=<%=rid%>">Water Skiing - AWSA</a>
     <a href="/rankings/view-scoreshq_AKA.asp?pvar=ByTour&sTourSportsGroup=NCW&rid=<%=rid%>">Collegiate - NCWSA</a>
     <a href="/rankings/view-scoreshq_AKA.asp?pvar=ByTour&sTourSportsGroup=AKA&rid=<%=rid%>">Kneeboard - AKA</a>
     <a href="http://barefoot.org/Scores/FindTournament.asp" target="_blank">Barefoot - ABC</a>
     <a href="http://www.nwsra.net/07marathon.htm" target="_blank">Ski Race - NWSRA</a>
     <a href="/rankings/defaultHQ.asp?process=login&rid=<%=rid%>">Administrative Log In</a>
    </div>
  <a href="javascript:toggle('ScrByMemb_subnav')">By Member</a>
    <div class="subnav" id="ScrByMemb_subnav" style="display:none;">
       <a href="/rankings/view-scoreshq_AKA.asp?pvar=ByMember&SptsGrpID=AWS&rid=<%=rid%>">Water Skiing - AWSA</a>
       <a href="/rankings/view-scoreshq_AKA.asp?pvar=ByMember&SptsGrpID=NCW&rid=<%=rid%>">Collegiate - NCWSA</a>
       <a href="/rankings/view-scoreshq_AKA.asp?pvar=ByMember&SptsGrpID=AKA&rid=<%=rid%>">Kneeboard - AKA</a>
       <a href="http://barefoot.org/results/findskier.asp" target="_blank">Barefoot - ABC</a>
       <a href="http://www.nwsra.net" target="_blank">Ski Race - NWSRA</a>
       <a href="/rankings/defaultHQ.asp?process=login&rid=<%=rid%>">Administrative Log In</a>
    </div>
  <a href="javascript:toggle('ViewScrBks_subnav')">View Scorebooks</a>
    <div class="subnav" id="ViewScrBks_subnav" style="display:none;">
       <a href="/rankings/view-tournamentshq.asp?sl=on&tr=on&ju=on&sTourLevel=Premier&sTourRange=5">Water Skiing - AWSA</a>
       <a href="/rankings/view-tournamentshq.asp?sl=on&tr=on&ju=on&sTourLevel=Collegiate&sTourRange=5">Collegiate - NCWSA</a>
    </div>
   </div>
  <p><a href="javascript:toggle('Tour_SptsDiv_subnav')">Events & Registration </a></p>
  <div class="subnav" id="Tour_SptsDiv_subnav" style="display:none;">
    <a href="javascript:toggle('findtour_subnav')">Event Search or Register</a>
      <div class="subnav" id="findtour_subnav" style="display:none;">
    <a href="/rankings/view-tournamentshq.asp?sl=on&tr=on&ju=on&rid=<%=rid%>">Water Skiing</a>
    <a href="/rankings/view-tournamentsHQ.asp?sl=on&tr=on&ju=on&wb=on&ws=on&wu=on&sTourLevel=Collegiate&rid=<%=rid%>">Collegiate</a>
    <a href="/rankings/view-tournamentsHQ.asp?wb=on&ws=on&wu=on&rid=<%=rid%>">Wakeboarding</a>
    <a href="/rankings/view-tournamentshq.asp?bf=on&rid=<%=rid%>">Barefooting</a>
    <a href="/rankings/view-tournamentshq.asp?hf=on&rid=<%=rid%>">Hydrofoiling</a>
    <a href="/rankings/view-tournamentshq.asp?kb=on&rid=<%=rid%>">Kneeboard</a>
    <a href="/rankings/login_registrar.asp">Registrar Login</a>
      </div>
    <a href="javascript:toggle('regtools_subnav')">Check Registration Status</a>
      <div class="subnav" id="regtools_subnav" style="display:none;">
    	<a href="/rankings/view-registration.asp?sl=on&tr=on&ju=on&process=viewreg&rid=<%=rid%>">Water Skiing</a><%
    	IF aml>=30 OR (aml>=30 AND usg="USW") THEN response.write("<a href='/rankings/view-registration.asp?wb=on&ws=on&wu=on&process=viewreg&rid="&rid&"'>Wakeboard</a>") %>
        <a href="http://www.usawaterski.org/admin/CreatePreRegTemplateSetup.asp?&rid=<%=rid%>">OLR Excel Download</a>
    	<a href="/rankings/login_registrar.asp">Registrar Login</a>
      </div>
  </div>
  <p><a href="javascript:toggle('Rank_SptsDiv_subnav')">Rankings Lists</a></p>
   <div class="subnav" id="Rank_SptsDiv_subnav" style="display:none;">
    <a href="/rankings/view-standingshq.asp?pvar=National&rid=<%=rid%>">Water Skiing - AWSA</a>
    <a href="/rankings/view-standingshq.asp?pvar=Junior&rid=<%=rid%>">Junior Ski - AWSA</a>
    <a href="javascript:toggle('RankNCWSA_subnav')">Collegiate - NCWSA </a>
      <div class="subnav" id="RankNCWSA_subnav" style="display:none;">
    	<a href="/rankings/View-TeamStdgsHQ.asp?rid=<%=rid%>">By Team</a>
    	<a href="/rankings/view-standingshq.asp?pvar=NCWSA&rid=<%=rid%>">By Individual</a>
      </div>	
    <a href="http://barefoot.org/ranking/" target="_blank">Barefoot - ABC</a>
    <a href="http://www.nwsra.net/points.htm" target="_blank">Ski Race - NWSRA</a>
    <a href="/rankings/view-grranking.asp?rid="&rid>GR Rankings TEST</a>
    <a href="/rankings/defaultHQ.asp?process=login&rid=<%=rid%>">Administrative Log In</a>
  </div>
 </div>
<br><br>
<%

IF adminmenulevel>=1 THEN

   IF aml>=20 THEN response.write("<p><a href='/rankings/login_registrar.asp?pvar=member'>OLR Registrar Functions</a></p>")
   	%>
	<p><a href="javascript:toggle('upanddown_subnav')">Manage Scores</a></p>
  <div class="subnav" id="upanddown_subnav" style="display:none;"><%
   IF aml>=10 THEN response.write("<a href='/rankings/defaultHQ.asp?process=uploadany&rid="&rid&"'>Upload a ZIP or Report File</a>")
   IF aml>=10 THEN response.write("<a href='/rankings/defaultHQ.asp?process=badscores&rid="&rid&"'>Fix Score Exceptions</a>") %>
    <a href="javascript:toggle('addeditscores_subnav')">Manually Add Scores</a>
      <div class="subnav" id="addeditscores_subnav" style="display:none;"><%
   	IF aml>=30 OR (aml>=10 AND (usg="AWS" OR usg="NCW")) THEN response.write("<a href='/rankings/addscores.asp?SD_Desc=AWS&rid="&rid&"'>Water Skiing</a>")  
   	IF aml>=30 OR (aml>=10 AND (usg="AWS" OR usg="NCW")) THEN response.write("<a href='/rankings/addscores.asp?SD_Desc=NCW&rid="&rid&"'>Collegiate</a>")  
        IF aml>=30 OR (aml>=10 AND usg="USW") THEN response.write("<a href='/rankings/addscores.asp?SD_Desc=USW&rid="&rid&"'>Wakeboard</a>")  
        IF aml>=30 OR (aml>=10 AND usg="AKA") THEN response.write("<a href='/rankings/addscores.asp?SD_Desc=AKA&rid="&rid&"'>Kneeboard</a>") %> 
      </div>
    <a href="javascript:toggle('broweditscores_subnav')">Search-Edit Scores</a>
      <div class="subnav" id="broweditscores_subnav" style="display:none;"><%
	IF aml>=30 OR (aml>=10 AND (usg="AWS" OR usg="NCW")) THEN response.write("<a href='/rankings/editscores.asp?search=0&pvar=AWS&rid="&rid&"'>Water Skiing</a>")
	IF aml>=30 OR (aml>=10 AND (usg="AWS" OR usg="NCW")) THEN response.write("<a href='/rankings/editscores.asp?search=0&pvar=NCW&rid="&rid&"'>Collegiate Skiing</a>")
	IF aml>=30 OR (aml>=10 AND usg="USW") THEN response.write("<a href='/rankings/editscores.asp?search=0&pvar=USW&rid="&rid&"'>Wakeboard</a>")
	IF aml>=30 OR (aml>=10 AND usg="AKA") THEN response.write("<a href='/rankings/editscores.asp?search=0&pvar=AKA&rid="&rid&"'>Kneeboard</a>") %>
      </div>
  </div>
<p><a href="javascript:toggle('filemanage_subnav')">Manage Tours and Files</a></p>
  <div class="subnav" id="filemanage_subnav" style="display:none;"><%
     IF aml>=30 OR (aml>=10 AND (usg="AWS" OR usg="NCW")) THEN response.write("<a href='/rankings/defaultHQ.asp?process=deletetour&rid="&rid&"'>Delete Tournament</a>")
     IF aml>=30 OR (aml>=2 AND (usg="AWS" OR usg="NCW")) THEN response.write("<a href='/rankings/defaultHQ.asp?process=uploadedzips&rid="&rid&"'>Uploaded ZIP Files</a>")
     IF aml>=30 OR (aml>=2 AND (usg="AWS" OR usg="NCW")) THEN response.write("<a href='/rankings/defaultHQ.asp?process=uploadedjmptms&rid="&rid&"'>Uploaded Jump CSV Files</a>")
     IF aml>=30 OR (aml>=10 AND (usg="AWS" OR usg="NCW")) THEN response.write("<a href='/rankings/defaultHQ.asp?process=fixsanction&rid="&rid&"'>Replace Sanction ID</a>")
     IF aml>=10 THEN response.write("<a href='/rankings/OffNameRepair.asp?process=Start&rid="&rid&"'>Find Sanction Officials IDs</a>")
     IF aml>=30 OR (aml>=10 AND (usg="AWS" OR usg="NCW")) THEN response.write("<a href='/rankings/PostTourStatus.asp?process=Start&rid="&rid&"'>Post Tournament Status</a>")
     IF aml>=30 OR (aml>=10 AND (usg="AWS" OR usg="NCW")) THEN response.write("<a href='/rankings/PTFollowUpStandalone2013.asp?rid="&rid&"'>P/T Followup Simulation</a>") %>
  </div>
<p><a href="javascript:toggle('reports1_subnav')">Reports & Downloads</a></p>
 <div class="subnav" id="reports1_subnav" style="display:none;"><%
   IF aml>=30 THEN response.write("<a href='/admin/Index.asp'>Registration Downloads</a>")
   IF aml>=25 THEN response.write("<a href='/rankings/defaultHQ.asp?process=trados&rid="&rid&"'>Download TRA-DOS Files</a>") 
   IF aml>=25 THEN response.write("<a href='/rankings/defaultHQ.asp?process=corson&rid="&rid&"'>Download IWSF R/L Scores</a>") 
   IF aml>=10 THEN response.write("<a href='/rankings/Guidebook.asp?process=Guide&rid="&rid&"'>Guidebook Report</a>")  
   IF aml>=10 THEN response.write("<a href='/rankings/RankChampions.asp?rid="&rid&"'>Ranking Champions</a>")  
   IF aml>=10 THEN response.write("<a href='/rankings/defaultHQ.asp?process=traffic&rid="&rid&"'>Website Traffic Stats</a>") 
   IF aml>=10 THEN response.write("<a href='/rankings/defaultHQ.asp?process=orphanreport&rid="&rid&"'>Orphaned Records</a>") %>
    <a href="javascript:toggle('AWSEF_subnav')">AWSEF Reports</a>
      <div class="subnav" id="AWSEF_subnav" style="display:none;"><%
	IF aml>=30 THEN response.write("<a href='/rankings/report_generic.asp?process=AWSEFLOC&rid="&rid&"'>Tours w/Donors</a>")
	IF aml>=30 THEN response.write("<a href='/rankings/report_generic.asp?process=DonorList&rid="&rid&"'>Donor Lists</a>") %>
      </div><%
   IF aml>=50 THEN response.write("<a href='/rankings/report_generic.asp?process=PWList&rid="&rid&"'>User PW List</a>")
   IF aml>=30 THEN response.write("<a href='/rankings/report_generic.asp?process=LOCContacts&rid="&rid&"'>LOC Contacts</a>") %>
 </div>
<p><a href="javascript:toggle('sysmanage_subnav')">System Management</a></p>
 <div class="subnav" id="sysmanage_subnav" style="display:none;"><%
   IF aml>=30 THEN response.write("<a href='/rankings/defaultHQ.asp?process=recalc&rid="&rid&"'>Recalculate Rankings</a>")  
   IF aml>=50 THEN response.write("<a href='/rankings/defaultHQ.asp?process=resetrcuf&rid="&rid&"'>Reset Recalc Flags</a>")  
   IF aml>=50 THEN response.write("<a href='/rankings/defaultHQ.asp?process=addyear&rid="&rid&"'>Add Ski Year</a>")  
   IF aml>=50 THEN response.write("<a href='/rankings/defaultHQ.asp?process=defaultyear&rid="&rid&"'>Set Default Year</a>") 
   IF aml>=50 THEN response.write("<a href='/rankings/tools_admin.asp?rid="&rid&"'>Create Divs for new SkiYear</a>") 	
   IF aml>=50 THEN response.write("<a href='/rankings/tools_admin_copy_default_to_12mo.asp?rid="&rid&"'>Update 12mo Divs by Default SY</a>") 	
   	%> 
     <a href="javascript:toggle('editdivfile_subnav')">Edit Division Tables</a>
      <div class="subnav" id="editdivfile_subnav" style="display:none;"><%
        IF aml>=30 THEN response.write("<a href='/rankings/editdivisions.asp?SptsGrpID=AWS&rid="&rid&"'>&nbsp; AWSA</a>")  
        IF aml>=30 THEN response.write("<a href='/rankings/editdivisions.asp?SptsGrpID=NCW&rid="&rid&"'>&nbsp; Collegiate</a>")  
        IF aml>=30 THEN response.write("<a href='/rankings/editdivisions.asp?SptsGrpID=USW&rid="&rid&"'>&nbsp; Wakeboard</a>") 
        IF aml>=30 THEN response.write("<a href='/rankings/editdivisions.asp?SptsGrpID=AKA&rid="&rid&"'>&nbsp; Kneeboard</a>")%>
     </div>
     <a href="javascript:toggle('editother_subnav')">Edit Other Tables</a>
      <div class="subnav" id="editother_subnav" style="display:none;"><%
        IF aml>=10 OR (aml>=2 AND usg="NCW") THEN response.write("<a href='/rankings/editteams.asp?sSptsGrpID=NCW&rid="&rid&"'>College Teams</a>")  %>
     </div>
     <a href="javascript:toggle('leagues_subnav')">AWSA Leagues</a>
      <div class="subnav" id="leagues_subnav" style="display:none;"><%
        IF aml>=20 OR (aml>=10 AND usg="AWS") THEN response.write("<a href='/rankings/editleagues.asp?sSptsGrpID=AWS&rid="&rid&"'>General Settings</a>")  
        IF aml>=20 OR (aml>=10 AND usg="AWS") THEN response.write("<a href='/rankings/editleaguetours.asp?sSptsGrpID=AWS&rid="&rid&"'>Touraments</a>") 
        IF aml>=20 OR (aml>=10 AND usg="AWS") THEN response.write("<a href='/rankings/EditLeagueQualifications.asp?sSptsGrpID=AWS&rid="&rid&"'>Qualifications</a>")  
        IF aml>=20 OR (aml>=10 AND usg="AWS") THEN response.write("<a href='/rankings/report_generic.asp?process=leaguequalsummary&rid="&rid&"'>Display COA</a>")  
        IF aml>=50 OR (aml>=50 AND usg="AWS") THEN response.write("<a href='/rankings/qualifyrecalc.asp'>Recalc Qualifications</a>") %> 
     </div>
     <%	IF aml>=50 THEN response.write("<a href='/rankings/Register_longlat_widget.asp?rid="&rid&"'>GPS Widget</a>") %>
     <a href="javascript:toggle('logfile_subnav')">View Log Files</a>
      <div class="subnav" id="logfile_subnav" style="display:none;"><%
	   IF aml>=25 THEN response.write("<a href='/rankings/defaultHQ.asp?process=download&file=tra-log.txt&rid="&rid&"'>Current Log File</a>")
	   IF aml>=25 THEN response.write("<a href='/rankings/defaultHQ.asp?process=download&file=tra-log-2016.txt&rid="&rid&"'>2016 Log File</a>")
	   IF aml>=25 THEN response.write("<a href='/rankings/defaultHQ.asp?process=download&file=tra-log-2015.txt&rid="&rid&"'>2015 Log File</a>")
	   IF aml>=25 THEN response.write("<a href='/rankings/defaultHQ.asp?process=download&file=tra-log-2014.txt&rid="&rid&"'>2014 Log File</a>")
	   IF aml>=25 THEN response.write("<a href='/rankings/defaultHQ.asp?process=download&file=tra-log-2013.txt&rid="&rid&"'>2013 Log File</a>")
	   IF aml>=25 THEN response.write("<a href='/rankings/defaultHQ.asp?process=download&file=tra-log-2012.txt&rid="&rid&"'>2012 Log File</a>")
	   IF aml>=25 THEN response.write("<a href='/rankings/defaultHQ.asp?process=download&file=tra-log-2011.txt&rid="&rid&"'>2011 Log File</a>")
	   IF aml>=25 THEN response.write("<a href='/rankings/defaultHQ.asp?process=download&file=tra-log-2010.txt&rid="&rid&"'>2010 Log File</a>")
	   IF aml>=25 THEN response.write("<a href='/rankings/defaultHQ.asp?process=download&file=tra-log-2009.txt&rid="&rid&"'>2009 Log File</a>")
	   IF aml>=25 THEN response.write("<a href='/rankings/defaultHQ.asp?process=download&file=tra-log-2008.txt&rid="&rid&"'>2008 Log File</a>")
	   IF aml>=25 THEN response.write("<a href='/rankings/defaultHQ.asp?process=download&file=tra-log-2007.txt&rid="&rid&"'>2007 Log File</a>")
	   IF aml>=25 THEN response.write("<a href='/rankings/defaultHQ.asp?process=download&file=tra-log-2006.txt&rid="&rid&"'>2006 Log File</a>") %>
     </div>

 </div>
<p><a href="javascript:toggle('registeradmin_subnav')">OLR Admin Tools</a></p>
 <div class="subnav" id="registeradmin_subnav" style="display:none;"><%  

   IF aml>=1 AND usg="AWS" THEN response.write("<a href='/rankings/view-tournamentsHQ.asp?sl_check=on&tr_check=on&ju_check=on&process=viewreg&rid="&rid&"'>Check Your Status Report</a>") %> 
     <a href="javascript:toggle('OLRStats_subnav')">OLR HQ Reports</a>
      <div class="subnav" id="OLRStats_subnav" style="display:none;"><%
	IF aml>=10 THEN response.write("<a href='/rankings/stats_tour.asp?pvar=entrystat&rid="&rid&"'>OLR Tournaments Status</a>")

	IF aml>=20 THEN response.write("<a href='/rankings/report_generic.asp?process=olrentries&rid="&rid&"'>Participation By Year</a>") 
	IF aml>=20 THEN response.write("<a href='/rankings/report_generic.asp?process=pblisting&rid="&rid&"'>Personal Best List</a>")		
	IF aml>=20 THEN response.write("<a href='/rankings/report_generic.asp?process=skierlist&rid="&rid&"'>Participant Email List</a>")
	IF aml>=20 THEN response.write("<a href='/rankings/report_generic.asp?process=ratinglist_deduped&rid="&rid&"'>Emails w-Rating Deduped</a>")
	IF aml>=20 THEN response.write("<a href='/rankings/report_generic.asp?process=ratinglist&rid="&rid&"'>Skiers With Rating</a>")	
	IF aml>=20 THEN response.write("<a href='/rankings/report_generic.asp?process=eliteskiers&rid="&rid&"'>Elite Skiers List</a>")	
	IF aml>=20 THEN response.write("<a href='/rankings/report_generic.asp?process=qualifylist&rid="&rid&"'>Qualified Skiers</a>")	
	IF aml>=1 THEN response.write("<a href='/rankings/report_generic.asp?process=bioinfo&rid="&rid&"'>Skier Bio List - Alpha</a>")
	IF aml>=1 THEN response.write("<a href='/rankings/report_generic.asp?process=bioinfo-evt&rid="&rid&"'>Skier Bio List - DivEvt</a>")
	IF aml>=20 THEN response.write("<a href='/rankings/report_generic.asp?process=classnf&rid="&rid&"'>Tounaments With N or F</a>")

		
		%> 
	  </div>
     <a href="javascript:toggle('devtools_subnav')">Developer Tools</a>
      <div class="subnav" id="devtools_subnav" style="display:none;"><%
	   IF aml>=30 THEN response.write("<a href='/rankings/PW_Update_Tool.asp?pvar=member'>PW Update Tool</a>") 
	   IF aml>=30 AND usg="AWS" THEN response.write("<a href='/rankings/Bio-Form2.asp?formstatus=search&rid="&rid&"'>Update Bio</a>") 
	   IF aml>=25 THEN response.write("<a href='/rankings/SiteImageTool.asp?pvar=member'>Site Image Tool</a>")  
	   IF aml>=50 AND usg="AWS" THEN response.write("<a href='/rankings/cont_disp_edit.asp?&rid="&rid&"'>Display Control Settings</a>")
		 IF aml>=20 THEN response.write("<a href='/rankings/report_generic.asp?process=olr_ipn_analysis_summary&rid="&rid&"'>OLR IPN - Summary</a>") 
	   IF aml>=20 THEN response.write("<a href='/rankings/report_generic.asp?process=olr_ipn_analysis&rid="&rid&"'>OLR IPN - Detail</a>") 
	   	%>
     </div>
     <a href="javascript:toggle('natls_subnav')">Nationals Reports</a>
      <div class="subnav" id="natls_subnav" style="display:none;"><%
 	IF aml>=10 THEN response.write("<a href='/rankings/report_generic.asp?process=refund&rid="&rid&"'>Nationals Refunds</a>") 
	IF aml>=10 THEN response.write("<a href='/rankings/report_generic.asp?process=nationals&rid="&rid&"'>Nationals Entry Flow</a>") 
	IF aml>=10 THEN response.write("<a href='/rankings/report_generic.asp?process=surveyresults&rid="&rid&"'>Nationals Survey Results</a>")%> 
     </div> 
 </div>
<p><a href="javascript:toggle('Grassroots_subnav')">Grassroots</a></p>
     <div class="subnav" id="Grassroots_subnav" style="display:none;"><%
	IF aml>=30 THEN response.write("<a href='/rankings/GRUploadform.asp?rid="&rid&"'>Upload GR Excel doc</a>")
	IF aml>=30 THEN response.write("<a href='/rankings/view-grranking.asp?WhatHeadFoot=gr&rid="&rid&"'>Prototype Rankings Page</a>")
	IF aml>=30 THEN response.write("<a href='/rankings/view-grscores.asp?rid="&rid&"'>Prototype Score Page</a>")%>
     </div> 

<p><a href="javascript:toggle('IAC_subnav')">International Activities</a></p>
 <div class="subnav" id="IAC_subnav" style="display:none;"><%
   IF aml>=50 THEN response.write("<a href='/rankings/IAC_reports.asp?rid="&rid&"'>Team Selection</a>")
   IF aml>=50 AND usg="AWS" THEN response.write("<a href='/rankings/IAC_Reports.asp?process=controlpanel&rid="&rid&"'>Control Panel</a>") %> 
 </div>
<br>
<p><A href='/rankings/defaultHQ.asp?process=logout&rid=<%=rid%>'>Log Out</A></p>  
<%
ELSE %>
    <p><a href='/rankings/defaultHQ.asp?process=login&rid=<%=rid%>'>Administrative Login</A></p><%
END IF
%>





