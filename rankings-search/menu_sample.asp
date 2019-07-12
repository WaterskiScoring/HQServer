<% '---This is the menu for Adminstrative Level login on TRA --- 

Dim adminmenulevel, aml, usg
adminmenulevel=Session("adminmenulevel")
aml=Session("adminmenulevel")
usg=Session("UserSptsGrpID")


%>
<p><a href="javascript:toggle('basic_subnav')">Member Sections</a></p>
 <div class="subnav" id="basic_subnav" style="display:none;">
  <a href="javascript:toggle('scores_subnav')">Results - Tournament Scores</a>
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
   </div>
  <a href="javascript:toggle('Tour_SptsDiv_subnav')">Tournament Schedules </a>
  <div class="subnav" id="Tour_SptsDiv_subnav" style="display:none;">
    <a href="/rankings/view-tournamentshq.asp?sl_check=on&tr_check=on&ju_check=on&rid=<%=rid%>">Water Skiing</a>
    <a href="/rankings/view-tournamentsHQ.asp?sl_check=on&tr_check=on&ju_check=on&wb_check=on&ws_check=on&wsu_check=on&sTourLevel=Collegiate&rid=<%=rid%>">Collegiate</a>
    <a href="/rankings/view-tournamentsHQ.asp?wb_check=on&ws_check=on&wsu_check=on&rid=<%=rid%>">Wakeboarding</a>
    <a href="/rankings/view-tournamentshq.asp?bf_check=on&rid=<%=rid%>">Barefooting</a>
    <a href="/rankings/view-tournamentshq.asp?hf_check=on&rid=<%=rid%>">Hydrofoiling</a>
    <a href="/rankings/view-tournamentshq.asp?kb_check=on&rid=<%=rid%>">Kneeboard</a>
  </div>
  <a href="javascript:toggle('Rank_SptsDiv_subnav')">Rankings Lists</a>
   <div class="subnav" id="Rank_SptsDiv_subnav" style="display:none;">
    <a href="/rankings/view-standingshq.asp?pvar=National&rid=<%=rid%>">Water Skiing - AWSA</a>
    <a href="/rankings/view-standingshq.asp?pvar=Junior&rid=<%=rid%>">Junior Ski - AWSA</a>
    <a href="/rankings/view-standingshq.asp?pvar=NCWSA&rid=<%=rid%>">Collegiate - NCWSA</a>
    <a href="http://barefoot.org/ranking/" target="_blank">Barefoot - ABC</a>
    <a href="http://www.nwsra.net/points.htm" target="_blank">Ski Race - NWSRA</a>
    <a href="/rankings/defaultHQ.asp?process=login&rid=<%=rid%>">Administrative Log In</a>
  </div>
 </div>
<br><%

IF adminmenulevel>=1 THEN
%>
<p><a href="javascript:toggle('upanddown_subnav')">Manage Scores</a></p>
  <div class="subnav" id="upanddown_subnav" style="display:none;"><%
   IF aml>=10 THEN response.write("<a href='/rankings/defaultHQ.asp?process=uploadfiles&rid="&rid&"'>Upload a WSP File</a>") 
   IF aml>=10 THEN response.write("<a href='/rankings/defaultHQ.asp?process=badscores&rid="&rid&"'>Fix Upload Exceptions</a>") %>
    <a href="javascript:toggle('addeditscores_subnav')">Manually Add Scores</a>
      <div class="subnav" id="addeditscores_subnav" style="display:none;"><%
   	IF aml>=30 OR (aml>=10 AND usg="AWS") THEN response.write("<a href='/rankings/addscores.asp?SD_Desc=AWS&rid="&rid&"'>Water Skiing</a>")  
   	IF aml>=30 OR (aml>=10 AND usg="NCW") THEN response.write("<a href='/rankings/addscores.asp?SD_Desc=NCW&rid="&rid&"'>Collegiate</a>")  
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
     IF aml>=30 OR (aml>=10 AND (usg="AWS" OR usg="NCW")) THEN response.write("<a href='/rankings/defaultHQ.asp?process=invalidfiles&rid="&rid&"'>Invalid WSP Names</a>")
     IF aml>=30 OR (aml>=10 AND (usg="AWS" OR usg="NCW")) THEN response.write("<a href='/rankings/defaultHQ.asp?process=uploadedfiles&rid="&rid&"'>Uploaded WSP Files</a>")
     IF aml>=30 OR (aml>=10 AND (usg="AWS" OR usg="NCW")) THEN response.write("<a href='/rankings/defaultHQ.asp?process=fixsanction&rid="&rid&"'>Replace Sanction ID</a>") %>
  </div>
<p><a href="javascript:toggle('reports1_subnav')">Reports & Downloads</a></p>
 <div class="subnav" id="reports1_subnav" style="display:none;"><%
   IF aml>=30 THEN response.write("<a href='/rankings/admin/CreateRegTemplateStep1.asp'>Download Excel Template</a>")
   IF aml>=25 THEN response.write("<a href='/rankings/defaultHQ.asp?process=trados&rid="&rid&"'>Download TRA-DOS Files</a>") 
   IF aml>=25 THEN response.write("<a href='/rankings/defaultHQ.asp?process=corson&rid="&rid&"'>Download IWSF R/L Scores</a>") 
   IF aml>=10 THEN response.write("<a href='/rankings/Guidebook.asp?process=Guide&rid="&rid&"'>Guidebook Report</a>")  
   IF aml>=10 THEN response.write("<a href='/rankings/defaultHQ.asp?process=traffic&rid="&rid&"'>Website Traffic Stats</a>") 
   IF aml>=20 THEN response.write("<a href='/rankings/defaultHQ.asp?process=orphanreport&rid="&rid&"'>Orphaned Records</a>") %>
 </div>
<p><a href="javascript:toggle('sysmanage_subnav')">System Management</a></p>
 <div class="subnav" id="sysmanage_subnav" style="display:none;"><%
   IF aml>=50 THEN response.write("<a href='/rankings/defaultHQ.asp?process=recalc&rid="&rid&"'>Recalculate Rankings</a>")  
   IF aml>=50 THEN response.write("<a href='/rankings/defaultHQ.asp?process=addyear&rid="&rid&"'>Add Ski Year</a>")  
   IF aml>=50 THEN response.write("<a href='/rankings/defaultHQ.asp?process=defaultyear&rid="&rid&"'>Set Default Year</a>") %> 
     <a href="javascript:toggle('editdivfile_subnav')">Edit Division Tables</a>
      <div class="subnav" id="editdivfile_subnav" style="display:none;"><%
        IF aml>=30 THEN response.write("<a href='/rankings/editdivisions.asp?SptsGrpID=AWS&rid="&rid&"'>&nbsp; AWSA</a>")  
        IF aml>=30 THEN response.write("<a href='/rankings/editdivisions.asp?SptsGrpID=NCW&rid="&rid&"'>&nbsp; Collegiate</a>")  
	IF aml>=30 THEN response.write("<a href='/rankings/editdivisions.asp?SptsGrpID=USW&rid="&rid&"'>&nbsp; Wakeboard</a>") 
        IF aml>=30 THEN response.write("<a href='/rankings/editdivisions.asp?SptsGrpID=AKA&rid="&rid&"'>&nbsp; Kneeboard</a>")%>
     </div>
     <a href="javascript:toggle('editother_subnav')">Edit Other Tables</a>
      <div class="subnav" id="editother_subnav" style="display:none;"><%
        IF aml>=50 OR (aml>=20 AND usg="NCW") THEN response.write("<a href='/rankings/editteams.asp?SptsGrpID=NCW&rid="&rid&"'>College Teams</a>")  
        IF aml>=50 OR (aml>=20 AND (usg="AWS" OR usg="NCW")) THEN response.write("<a href='/rankings/editteams.asp?SptsGrpID=NCW&rid="&rid&"'>League/Series</a>") %> 
     </div><%
   IF aml>=25 THEN response.write("<a href='/rankings/defaultHQ.asp?process=download&file=tra-log.txt&rid="&rid&"'>View Log File</a>") %>  
 </div>
<p><a href="javascript:toggle('registeradmin_subnav')">On Line Registration</a></p>
 <div class="subnav" id="registeradmin_subnav" style="display:none;"><%
   IF aml>=30 AND usg="AWS" THEN response.write("<a href='/rankings/registration_bywizard.asp?process=register&rid="&rid&"'>Online Registration</a>")  
   IF aml>=30 AND usg="AWS" THEN response.write("<a href='/rankings/view-registration.asp?process=regreport&sTourID=07W999A&rid="&rid&"'>Check Your Status Report</a>")  
   IF aml>=50 THEN response.write("<a href='/rankings/Bio-Form2.asp?formstatus=search&rid="&rid&"'>Update Bio</a>") %>
     <a href="javascript:toggle('regtools_subnav')">Admin Tools</a>
      <div class="subnav" id="regtools_subnav" style="display:none;"><%
        IF aml>=50 THEN response.write("<a href='/rankings/PW_Update_Tool.asp?&rid="&rid&"'>&nbsp;Password Update Tool</a>")  
     </div>
 </div>
<br>
<p><A href='/rankings/defaultHQ.asp?process=logout&rid=<%=rid%>'>Log Out</A></p>  
<%
ELSE %>
    <p><a href='/rankings/defaultHQ.asp?process=login&rid=<%=rid%>'>Administrative Login</A></p><%
END IF
%>





