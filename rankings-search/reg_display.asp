<!--#include virtual="/rankings/settingsHQ.asp"--><%


Reg_Disp_Test


SUB Reg_Disp_Test

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
    <style type="text/css">
    /* Required Fields */
        .reqd_blu  {
            font-weight: bold;
            color: blue;
        }
    /* Accordion */
        .accordionHeader {
            border: 1px solid #2F4F4F;
            color: white;
            background-color: #2E4d7B;
	        font-family: Arial, Sans-Serif;
	        font-size: 12px;
	        font-weight: bold;
            padding: 5px;
            margin-top: 5px;
            cursor: pointer;
        }
        .accordionHeaderSelected {
            border: 1px solid #2F4F4F;
            color: white;
            background-color: #5078B3;
	        font-family: Arial, Sans-Serif;
	        font-size: 12px;
	        font-weight: bold;
            padding: 5px;
            margin-top: 5px;
            cursor: pointer;
        }
        .accordionHeader a:hover {
	        background: none;
	        text-decoration: underline;
        }
        .accordionHeader a {
	        color: #FFFFFF;
	        background: none;
	        text-decoration: none;
        }
        .accordionHeaderSelected a {
	        color: #FFFFFF;
	        background: none;
	        text-decoration: none;
        }
        .accordionContent {
            background-color: #D3DEEF;
            border: 1px dashed #2F4F4F;
            border-top: none;
            padding: 5px;
            padding-top: 10px;
        }
        /************ MaskedEdit Related Styles ***********************/
.MaskedEditFocus
{
    background-color: #ffffcc;
    color: #000000;
}
.MaskedEditMessage
{
	color: #ff0000;
	font-weight: bold;
}
.MaskedEditError
{
    background-color: #ffcccc;
}
.MaskedEditFocusNegative
{
    background-color: #ffffcc;
    color: #ff0000;
}
.MaskedEditBlurNegative
{
    color: #ff0000;
}

.MyCalendar .ajax__calendar_container {
    border:1px solid #646464;
    background-color: lemonchiffon;
    color: red;
}
.MyCalendar .ajax__calendar_other .ajax__calendar_day,
.MyCalendar .ajax__calendar_other .ajax__calendar_year {
    color: black;
}
.MyCalendar .ajax__calendar_hover .ajax__calendar_day,
.MyCalendar .ajax__calendar_hover .ajax__calendar_month,
.MyCalendar .ajax__calendar_hover .ajax__calendar_year {
    color: black;
}
.MyCalendar .ajax__calendar_active .ajax__calendar_day,
.MyCalendar .ajax__calendar_active .ajax__calendar_month,
.MyCalendar .ajax__calendar_active .ajax__calendar_year {
    color: black;
    font-weight:bold;
}
    </style>

<title>
	Untitled Page
</title>
<link href="http://www.usawaterski.org/css/styles.css" rel="stylesheet" type="text/css" />
<link href="/sanctions/WebResource.axd?d=zo9GqCX3ABOTgZ72LXU1yjy_y6fKPfpQ1RoJzRZZcvmf3pt3SjXNXTQDV__Z7jhvXjKxh1sxLbqdFOxoeZfrew2&amp;t=633323210281093750" type="text/css" rel="stylesheet" />
</head>
<body>

    <form name="form1" method="post" action="/rankings/registration_ByWizard.asp" onsubmit="javascript:return WebForm_OnSubmit();" id="form1">

    <div>



        <script type="text/javascript">
//<![CDATA[
Sys.WebForms.PageRequestManager._initialize('ScriptManager1', document.getElementById('form1'));
Sys.WebForms.PageRequestManager.getInstance()._updateControls(['tctl04$UpdatePanel1','tctl06$UpdatePanel2','tctl08$UpdatePanel3','tctl10$UpdatePanel4','tctl12$UpdatePanel5','tctl14$UpdatePanel6'], [], [], 90);
//]]>
</script>

   
        <table class="contentcontainer" width="<%=TourTableWidth%>px">
            <tr><td class="content_2_col1"><h1>USA Water Ski & USA Wakeboard Tournament Registration</h1>
 	Click on the headings to display fields.

	<% 	' -------------------------------------------------------------------
	  	' ---------------------  MEMBER PERSONAL DATA  ----------------------
	  	' ------------------------------------------------------------------- %>

        <div id="Accordion1">
		<input type="hidden" name="Accordion1_AccordionExtender_ClientState" id="Accordion1_AccordionExtender_ClientState" value="0" />
	   <div class="accordionHeaderSelected">
		<% sMemberID="000001151" %>
		<a href="">STEP 1 - Membership Information &nbsp;&nbsp;&nbsp;&nbsp;MemberID:&nbsp;<font color=Yellow><%=sMemberID%></font></a>
	</div>

	<div class="accordionContent" style="display:block;">
	  <div id="ctl04_UpdatePanel1">	 

	  <TABLE BORDER="4" CELLPADDING="3" CELLSPACING="0" width=100% BGCOLOR="<%=TableColor1%>">
	  <tr  valign=top>  
	    <TH ALIGN="left" width=20% ><font size=<% =fontsize2 %>  COlOR="#000000">Name</FONT>&nbsp;&nbsp;
	    <br><FONT size=<% =fontsize2 %>  COlOR="#0000CD"><% =sFirstName&" "&sLastName %></FONT></TH>

	    <TH ALIGN="left" width=15% vAlign="top"><font size=<% =fontsize2 %>  COlOR="#000000">Member ID</FONT>&nbsp;&nbsp;
	    <br><FONT size=<% =fontsize2 %>  COlOR="#0000CD"><% =sMemberID %></FONT></TH>
  
	    <TH ALIGN="left" width=15% vAlign="top"><font size=<% =fontsize2 %>  COlOR="#000000">City/ST</FONT>&nbsp;&nbsp;
	    <br><FONT size=<% =fontsize2 %>  COlOR="#0000CD"><% =sMembCity&", "&sMembState %></FONT></TH>

	    <TH ALIGN="left" vAlign="top"><font size=<% =fontsize2 %>  COlOR="#000000">Age/Gender</FONT>&nbsp;&nbsp;
	    <br><FONT size=<% =fontsize2 %>  COlOR="#0000CD"><a TITLE="Correct Age and Gender is required. Please make sure this information is correct.  Contact USA Water Ski Membership Department at 800-533-2972 to correct any missing or inaccurate information."> 
		<% =sMembAge&"/"&sMembSex %></a></FONT></TH><%

	    
	    ' -------------------------------------------------------
	    ' ----------- Team Selection (collegiate)  --------------
	    ' ------------------------------------------------------- %>

		<TH ALIGN="left" vAlign="top"><font size=<% =fontsize2 %>  COlOR="#000000">Team</FONT>&nbsp;&nbsp;
	 	  <br>
		  <FONT size=<% =fontsize2 %>  COlOR="<% =textcolor2 %>">&nbsp;<% =MembTeam %></FONT>
 		</TH><%

	    ' Loads Team drop-down list
	    ' LoadTeam MembTeam 
 
		%>

	  </tr><%

		' -----------------------------------------------------------------------------------
		' ----------------------  MEMBERSHIP AND ENTRY STATUS  ------------------------------ 
		' ----------------------------------------------------------------------------------- %> 		
	      <tr valign=top>
	  	<TH ALIGN="left" vAlign="top"><font size=<% =fontsize2 %>  COlOR="#000000">Competition Status</FONT><%
	  	IF sCanSkiTour = TRUE THEN  
	    		%><br><FONT size=<% =fontsize2 %>  COlOR="#0000CD">OK - <% =sMembTypeCode %></FONT></TH><%
	  	ELSE
	   	 	%><br><FONT size=<% =fontsize2 %>  COlOR=red><% =sMembTypeCode %> - Upgrade Required</FONT></TH><%
	  	END IF 

		%><TH ALIGN="left" vAlign="top"><font size=<% =fontsize2 %>  COlOR="#000000">Expiration</FONT><%


		' ---  Checks End Date of tournament against Expiration Date of membership record  ---
		IF DateDiff("d", sEffectiveto, sTourEDate) <= 0  THEN
	    		%><br><FONT size=<% =fontsize2 %>  COlOR=blue>OK - <% =sEffectiveto %></FONT></TH><%
	  	ELSE
	    		%><br><FONT size=<% =fontsize2 %>  COlOR=red>Renew - <% =sEffectiveto %></FONT></TH><%
	  	END IF 

		' -------------------------------
		' ------  Payment Status  -------
		' -------------------------------

		%><TH ALIGN="left" vAlign="top"><font size=<% =fontsize2 %>  COlOR="#000000">Payment Status</FONT>&nbsp;&nbsp;<%

		' ----  Override of Fees has been set  -----
		IF TRIM(sMoneyOverride) <> "" THEN
		 	%><br><FONT size=<% =fontsize2 %>  COlOR="red">Fee Override - <%=sMoneyOverride%></FONT></TH><%		 

		' ----  Fees from CCLogTablName (Previous Charges) are less than current form values  -----
		ELSEIF cdbl(sTotalFormFees) <> 0 AND cdbl(sTotalPreviousFees) < cdbl(sTotalFormFees) THEN
		 	%><br><FONT size=<% =fontsize2 %>  COlOR="red">Balance Due</FONT></TH><%		 

		' ----  Fees from RegGenTable (Previous) are greater than current form values  -----
		ELSEIF sTotalFormFees <> 0 AND cdbl(sTotalPreviousFees) > cdbl(sTotalFormFees) THEN
		 	%><br><FONT size=<% =fontsize2 %>  COlOR="red">Refund Due</FONT></TH><%		 

		' ---- confirm has not been pressed 
		ELSEIF sTotalFormFees = 0 AND cdbl(sTotalPreviousFees) = cdbl(0) THEN
		 	%><br><FONT size=<% =fontsize2 %>  COlOR="red">Not Entered</FONT></TH><%

		' ----------------------------------------------------------------------------------------------
		' ---- *****  MARK - DO WE NEED A NEW CONDITION?  when FORM has never been confirmed and displaying original information 
		  ELSE
			%><br><FONT size=<% =fontsize2 %>  COlOR="#0000CD">Paid in Full</FONT></TH><%
		END IF 

		' ----------------------------------------
		' -------- Liability Waiver --------------
		' ----------------------------------------  
		%>
		<TH ALIGN="left" vAlign="top"><font size=<% =fontsize2 %>  COlOR="#000000">Release</FONT>&nbsp;&nbsp;<%

		IF TRIM(Session("sRelease")) = "" THEN
		 	%><br><FONT size=<% =fontsize2 %>  COlOR="red">Not Signed</FONT></TH><%
		ELSE
			%><br><FONT size=<% =fontsize2 %> face="<% =font1 %>" COlOR="#0000CD">Complete</FONT></TH><%
		END IF 
	
		' ----------------------------------------
		' ----------- Personal Bio  --------------
		' ----------------------------------------
		%>
		<TH ALIGN="left" vAlign="top"><font size=<% =fontsize2 %>  COlOR="#000000">Pers Bio</FONT>&nbsp;&nbsp;
		<%

		IF sBioDone = "Y" THEN 

			%><br><FONT size=<% =fontsize2 %>  COlOR="<% =TextColor2 %>">Complete</FONT></TH><%
		ELSE  
			%><br><FONT size=<% =fontsize2 %>  COlOR="red">Incomplete</FONT></TH><% 
			sErrorNo = sErrorNo + 1
		END IF %>

		</TH>
	       </tr>


		<tr>

		<th colspan=1>  <%
		' --- Button to select NEW MEMBER  --- 

		IF adminmenulevel >= 50 THEN  %>
			<form action="/rankings/RegistrationHQ.asp?rid=<%=rid%>&sRunByWhat=Member" method="post">
		  	   <input type="submit" border-top=0 border-bottom=0  padding-top=0px padding-bottom=0px value="New Member">
			</form> <%
		END IF  
		
		%>
		</th>

		<th align=left colspan=4><%

		' ---------------------------
		' ---- Top of main form  ----
		' ---------------------------

		IF FormStatus = "modify" THEN
			%><form action="/rankings/RegistrationHQ.asp?sRunByWhat=Edit&formstatus=confirm" method="post"><% 
		ELSE 
			%><form action="/rankings/RegistrationHQ.asp?sRunByWhat=Edit&formstatus=modify" method="post"><%
		END IF 



		IF adminmenulevel >= 50 THEN  %>	
		    <FONT size=<% =fontsize2 %>  COlOR="<% =TextColor1 %>">Override</FONT>
			<select name="sMembOverride" value="<% =sMembOverride %>" <% =OverrideStatus %> >
			  <option value ="" <%IF sMembOverride = "" THEN Response.Write(" selected ")%> >None</Option><br>
			  <option value ="JOS" <%IF sMembOverride = "JOS" THEN Response.Write(" selected ")%> >Joined On Site</Option><br>
			  <option value ="PRF" <%IF sMembOverride = "PRF" THEN Response.Write(" selected ")%> >Proof Supplied</Option><br>
			</select><%
		ELSE %>
			<input type="hidden" name="sMembOverride" value="<% =sMembOverride %>"><%
		END IF  %>

		</th>

		</tr>
	     </table>                         
                            
	  </div>
	</div>




	<% ' ---------------------------------------------------------------------------
	' ----------------------  DISPLAY TOURNAMENT INFORMATION  ---------------------- 
	' ------------------------------------------------------------------------------ %>


	<div class="accordionHeader">
		<a href="">STEP 2 - Tournament </a>&nbsp;&nbsp;&nbsp;&nbsp;
	</div>

	<div class="accordionContent" style="display:none;">
	  <div id="ctl06_UpdatePanel2">

		<br>
	  	<TABLE BORDER="4" CELLPADDING="3" CELLSPACING="0" width=100% BGCOLOR="<%=TableColor1%>">

		<tr>  
		  <TD WIDTH = 10% ALIGN="left" vAlign="top"><FONT COlOR="#000000" size=<% =fontsize2 %>  >Tour ID</FONT>
		    <br><FONT COlOR="#0000CD" size=<% =fontsize2 %>  ><% =sTourID %></FONT></TD>

		  <TD Width = 30% colspan=4 ALIGN="left" vAlign="top"><FONT COlOR="#000000" size=<% =fontsize2 %> >Tour Name</FONT>
		    <br><FONT COlOR="#0000CD" size=<% =fontsize2 %> ><% =sTourName %></FONT></TD>
  
		  <TD Width = 20% ALIGN="left" vAlign="top"><FONT COlOR="#000000" size=<% =fontsize2 %> >City/ST</FONT>
		    <br><FONT COlOR="#0000CD" size=<% =fontsize2 %> ><% =sTourCity&", "&sTourState %></FONT></TD>

		  <TD Width = 15% colspan=2 ALIGN="left" vAlign="top"><FONT COlOR="#000000" size=<% =fontsize2 %> >Dates</FONT>
	 	    <br><FONT COlOR="#0000CD" size=<% =fontsize2 %> ><% =sTourSDate&"-"&sTourEDate %></FONT></TD>

		  <TD Width = 15% ALIGN="left" vAlign="top"><FONT COlOR="#000000" size=<% =fontsize2 %> >SptDiv</FONT>
		    <br><FONT COlOR="#0000CD" size=<% =fontsize2 %> ><% =sSptsGrpID %></FONT></TD>
		</tr>

		<tr>		
		  <TD WIDTH=10% ALIGN="left" vAlign="top"><FONT COlOR="#000000" size=<% =fontsize2 %> >1 Event Fee</FONT><br>
			<FONT COlOR="#0000CD" size=<% =fontsize2 %> ><% =FormatCurrency(sTEntryFee1,2) %></FONT></TD>		

		  <TD WIDTH=10% ALIGN="left" vAlign="top"><FONT COlOR="#000000" size=<% =fontsize2 %> >2 Event Fee</FONT><%
		  IF Cdbl(sTEntryFee2) > 0.00 THEN
			%><br><FONT COlOR="#0000CD" size=<% =fontsize2 %> ><% =FormatCurrency(sTEntryFee2,2) %></FONT></TD><%
		  ELSE 
			%>&nbsp;&nbsp</td><%		
		  END IF

		  %><TD WIDTH = 10% ALIGN="left" vAlign="top"><FONT COlOR="#000000" size=<% =fontsize2 %> >3 Event Fee</FONT><%
		  IF Cdbl(sTEntryFee3) > 0.00 THEN	
		 	%><br><FONT COlOR="#0000CD" size=<% =fontsize2 %> ><% =FormatCurrency(sTEntryFee3,2) %></FONT></TD><%
		  ELSE 
			%>&nbsp; &nbsp</td><%		
		  END IF
	
		  %><TD WIDTH = 10% ALIGN="left" vAlign="top"><FONT COlOR="#000000" size=<% =fontsize2 %> ><a title="Family Entry must be entered by paper form">Family Fee</a></FONT><%
		  IF Cdbl(sTEntryFeeFamily) > 0.00 THEN
			%><br><FONT COlOR="#0000CD" size=<% =fontsize2 %> ><% =FormatCurrency(sTEntryFeeFamily,2) %></FONT></TD><%
		  ELSE 
			%>&nbsp; &nbsp</td><%		
		  END IF

		  %><TD WIDTH = 15% ALIGN="left" vAlign="top"><FONT COlOR="#000000" size=<% =fontsize2 %> >Late Fee/day</FONT><%
		  IF Cdbl(sLateFee) > 0.00 THEN
			%><br><FONT COlOR="#0000CD" size=<% =fontsize2 %> ><% =FormatCurrency(sLateFee,2) %></FONT></TD><%
		  ELSE 
			%>&nbsp; &nbsp</td><%		
		  END IF %>
			<TD ALIGN="left" vAlign="top"><FONT COlOR="#000000" size="<% =fontsize2 %>" face="<% =font1 %>">Registration Deadline</FONT>
			<br><FONT COlOR="#0000CD" size=<% =fontsize2 %> ><% =sTourDeadline %></FONT></td> <%

		EntryTypeStatus = ""
		IF FormStatus = "confirm" OR adminmenulevel < 50 THEN EntryTypeStatus = "disabled" %>
		
		  <td valign=top colspan=2 align="left">
			<FONT COlOR="#000000" size=<% =fontsize2 %>  >Entry Type</font>
			<br>
			<select name="sEntryType" value="<% =sEntryType %>" <% =EntryTypeStatus %> >
			  <option value ="IND" <%IF sEntryType = "IND" THEN Response.Write(" selected ")%> >Individual</Option><br>
			  <option value ="HOH" <%IF sEntryType = "HOH" THEN Response.Write(" selected ")%> >Family HOH</Option><br>
			  <option value ="MEM" <%IF sEntryType = "MEM" THEN Response.Write(" selected ")%> >Family Member</Option><br>
			</select>
		  </td>
  		  <td> &nbsp; &nbsp</td>

		  

	 	 </tr>
		</table>

	  </div>
	</div>




	<div class="accordionHeader">
		<a href="">STEP 3 - Enter Events</a>
	</div>

	<div class="accordionContent" style="display:none;">
	  <div id="ctl08_UpdatePanel3">
			
                        <table>
                                <tr><td><span title="AWSA, NCWSA "><input id="ctl06_GrSl_L1" type="checkbox" name="ctl06$GrSl_L1" onclick="javascript:setTimeout('__doPostBack(\'ctl06$GrSl_L1\',\'\')', 0)" /></span>Slalom</td><td align="left" style="width: 271px"><br />
                                    </td></tr>
                                <tr><td><span title="AWSA, NCWSA"><input id="ctl06_GrTr_L1" type="checkbox" name="ctl06$GrTr_L1" onclick="javascript:setTimeout('__doPostBack(\'ctl06$GrTr_L1\',\'\')', 0)" /></span>Trick</td><td align="left" style="width: 271px"><br />
                                    </td></tr>
                                <tr><td><span title="AWSA, NCWSA"><input id="ctl06_GrTr_L1" type="checkbox" name="ctl06$GrTr_L1" onclick="javascript:setTimeout('__doPostBack(\'ctl06$GrJu_L1\',\'\')', 0)" /></span>Jumping</td><td align="left" style="width: 271px"><br />
                                    </td></tr>
                                <tr><td><span title="ABC"><input id="ctl06_GrBSl_L1" type="checkbox" name="ctl06$GrBSl_L1" onclick="javascript:setTimeout('__doPostBack(\'ctl06$GrBSl_L1\',\'\')', 0)" /></span>Barefoot Slalom</td><td align="left" style="width: 271px"><br />
                                    </td></tr>
                                <tr><td><span title="ABC"><input id="ctl06_GrBTr_L1" type="checkbox" name="ctl06$GrBTr_L1" onclick="javascript:setTimeout('__doPostBack(\'ctl06$GrBTr_L1\',\'\')', 0)" /></span>Barefoot Trick</td><td align="left" style="width: 271px"><br />
                                    </td></tr>
                                <tr><td><span title="USW"><input id="ctl06_GRWakebd_L1" type="checkbox" name="ctl06$GRWakebd_L1" onclick="javascript:setTimeout('__doPostBack(\'ctl06$GRWakebd_L1\',\'\')', 0)" /></span>Wakeboard</td><td align="left" style="width: 271px"><br />
                                    </td></tr>
                                <tr><td><span title="USW"><input id="ctl06_GRWSkate_L1" type="checkbox" name="ctl06$GRWSkate_L1" onclick="javascript:setTimeout('__doPostBack(\'ctl06$GRWSkate_L1\',\'\')', 0)" /></span>Wakeskate</td><td align="left" style="width: 271px"><br />
                                    </td></tr>
                                <tr><td><span title="USW"><input id="ctl06_GRWSurf_L1" type="checkbox" name="ctl06$GRWSurf_L1" onclick="javascript:setTimeout('__doPostBack(\'ctl06$GRWSurf_L1\',\'\')', 0)" /></span>Wakesurf</td><td align="left" style="width: 271px"><br />
                                    </td></tr>
                                <tr><td><span title="USW"><input id="ctl06_GrRailJ_L1" type="checkbox" name="ctl06$GrRailJ_L1" onclick="javascript:setTimeout('__doPostBack(\'ctl06$GrRailJ_L1\',\'\')', 0)" /></span>Rail Jam</td><td align="left" style="width: 271px"><br />
                                    </td></tr>
                                <tr><td><span title="AKA"><input id="ctl06_GrKneeSL_L1" type="checkbox" name="ctl06$GrKneeSL_L1" onclick="javascript:setTimeout('__doPostBack(\'ctl06$GrKneeSL_L1\',\'\')', 0)" /></span>Kneeboard Slalom</td><td align="left" style="width: 271px"><br />
                                    </td></tr>
                                <tr><td><span title="AKA"><input id="ctl06_GrKneeTr_L1" type="checkbox" name="ctl06$GrKneeTr_L1" onclick="javascript:setTimeout('__doPostBack(\'ctl06$GrKneeTr_L1\',\'\')', 0)" /></span>Kneeboard Trick</td><td align="left" style="width: 271px"><br />
                                    </td></tr>
                                <tr><td><span title="USH"><input id="ctl06_GrHfoil_L1" type="checkbox" name="ctl06$GrHfoil_L1" onclick="javascript:setTimeout('__doPostBack(\'ctl06$GrHfoil_L1\',\'\')', 0)" /></span>Hydrofoil</td><td align="left" style="width: 271px"><br />
                                    </td></tr>
                                <tr><td><span title="Use for unconventional formats or events"><input id="ctl06_GrOther" type="checkbox" name="ctl06$GrOther" onclick="javascript:setTimeout('__doPostBack(\'ctl06$GrOther\',\'\')', 0)" /></span>Other</td><td align="left" style="width: 271px">
                                    &nbsp;&nbsp;
                                    </td></tr>
                                <tr><td><span id="ctl06_labelFDescription">List of Events:</span></td><td align="left" ><textarea name="ctl06$TB_FDescription" rows="3" cols="20" id="ctl06_TB_FDescription" title="For online advertisement. 150 characters max." style="width:450px;"></textarea></p>
                                    </td></tr>
                            </table>                            
	  </div>
	</div>



	<div class="accordionHeader">
		<a href="">STEP 3a - Original Enter Events</a>
	</div>

	<div class="accordionContent" style="display:none;">
	  <div id="ctl08_UpdatePanel3a">

	<% 	' ----------------------------------------------------------------------
		' ---------------------  EVENT ENTRY  ----------------------------------
		' ----------------------------------------------------------------------
		%>

		<br>
		<TABLE BORDER="4" CELLPADDING="3" CELLSPACING="0" width=100% BGCOLOR="<%= TableColor1 %>">

 		<tr>
		  <TD WIDTH = 10% ALIGN="left" vAlign="top"><FONT COlOR="#000000" size=<% =fontsize2 %> ><br>Event</FONT></TD><%

		  IF FormStatus = "modify" THEN
			%><TD WIDTH = 8% ALIGN="left" vAlign="top"><FONT COlOR="red" size=<% =fontsize2 %> >Check<br>To Enter</FONT></TD>
			  <TD WIDTH = 20% ALIGN="left" vAlign="top"><FONT COlOR="red" size=<% =fontsize2 %> ><br>Divisions Offered</FONT></TD><%
		  ELSE
			%><TD WIDTH = 8% ALIGN="left" vAlign="top"><FONT COlOR="#000000" size=<% =fontsize2 %> >Enter<br>Event</FONT></TD>
			  <TD WIDTH = 20% ALIGN="left" vAlign="top"><FONT COlOR="#000000" size=<% =fontsize2 %> ><br>Divisions Entered</FONT></TD><%
		  END IF %>
		  <TD WIDTH = 22% ALIGN="left" vAlign="top"><FONT COlOR="#000000" size=<% =fontsize2 %> ><br>Qualifications Status</FONT></TD>
		  <TD WIDTH = 20% ALIGN="left" vAlign="top"><FONT COlOR="#000000" size=<% =fontsize2 %> ><br>Boat/Ramp/Weight</FONT></TD>
		  <TD WIDTH = 20% ALIGN="left" vAlign="top"><FONT COlOR="#000000" size=<% =fontsize2 %> ><br>Comments</FONT></TD>
		  <% IF adminmenulevel >=19 THEN %>		
		    	  <TD WIDTH = 20% ALIGN="left" vAlign="top"><FONT size=<% =fontsize2 %>  COlOR="<% =TextColor1 %>">Override</FONT></TD>
		  <% END IF %>
	    	</tr>  


		<tr><% 
        	  ' --------------------------------------------------------------------------------------------
	 	  ' ------------  Displays checkbox OPTION TourGenTable shows data in Event1 field  ------------
        	  ' --------------------------------------------------------------------------------------------

		  ' ---- Loads Qualifications and indicates if skier has required qualifications
		  sEvent1 = sTEvent1
		  sEvent2 = sTEvent2
		  sEvent3 = sTEvent3
		  sEvent4 = sTEvent4


		' ---temporarily disabled 
		  'LoadQualifications sEvent1, sDiv1, 1    
		  'LoadQualifications sEvent2, sDiv2, 2   
		  'LoadQualifications sEvent3, sDiv3, 3   
		  'LoadQualifications sEvent4, sDiv4, 4 
		  'OtherQualifications  


		  IF TRIM(sTEvent1) <> "" AND TRIM(sTEvent1) <> "X" THEN %>
			<td><font size=<% =fontsize2 %> ><%= sTEvent1Name %></td><%

			sEventNo=1
			sEvent1 = sTEvent1

			IF FormStatus = "modify" THEN  %>
				<td><input type=checkbox name="fSelectEvent1" <% IF sSelectEvent1 = "on" THEN Response.Write("Checked") %>></font></td>
				<td><% LoadDivPulldown sDiv1, sEvent1 %></td>
				<td>&nbsp</td>  <%

			ELSEIF FormStatus = "confirm" AND sSelectEvent1 ="on" THEN %>
				<td><font color="<% = textcolor2 %>" size=<% =fontsize2 %> >YES</font></td>
				<td><font color="<% =textcolor2 %>" size=<% =fontsize2 %> ><%= sDiv1 %></td><%

				IF LEFT(QualStatEvent1,2) = "OK" THEN
					%><td><font color="<% = textcolor2 %>" size=<% =fontsize2 %> ><% response.write(QualStatEvent1&" - "&QualEvent1) %></font></td><%
				ELSE
					%><td><font color=red size=<% =fontsize2 %> ><% response.write(QualStatEvent1&" - "&QualEvent1) %></font></td><%
				END IF
			ELSE  %>
				<td>&nbsp</td>
				<td>&nbsp</td>
				<td>&nbsp</td><%
			END IF %>

			  <td><font color="<% =textcolor2 %>" size=<% =fontsize2 %> >&nbsp;<%= EV1_BRWText %></td>
			  <td><font color="<% =textcolor2 %>" size=<% =fontsize2 %> >&nbsp;<%= CommentText1 %></td>

			  <% IF adminmenulevel >=19 THEN %>
			  <td>		
				<select name="sQfyOverride_E1" value="<% =sQfyOverride_E1 %>" <% =OverrideStatus %> >
				  <option value ="" <%IF sQfyOverride_E1 = "" THEN Response.Write(" selected ")%> >None</Option><br>
				  <option value ="EP" <%IF sQfyOverride_E1 = "EP" THEN Response.Write(" selected ")%> >Proof of EP</Option><br>
				  <option value ="3MS" <%IF sQfyOverride_E1 = "3MS" THEN Response.Write(" selected ")%> >3rd Masters</Option><br>
				  <option value ="OVR" <%IF sQfyOverride_E1 = "OVR" THEN Response.Write(" selected ")%> >NOPS EP</Option><br>
				  <option value ="ALT" <%IF sQfyOverride_E1 = "ALT" THEN Response.Write(" selected ")%> >Alt Div Qfy</Option><br>
				  <option value ="OAP" <%IF sQfyOverride_E4 = "OAP" THEN Response.Write(" selected ")%> >Reg/Nat OA Plc</Option><br>
				  <option value ="OTH" <%IF sQfyOverride_E1 = "OTH" THEN Response.Write(" selected ")%> >Other Qfy</Option><br>
				  <option value ="DNS" <%IF sEntryType = "DNS" THEN Response.Write(" selected ")%> >Did Not Ski</Option><br>
				</select>
			  </td>
			  <% ELSE %>
				<input type="hidden" name="sQfyOverride_E1" value="<% =sQfyOverride_E1 %>">		
			  <% END IF 

		  END IF %>


		</tr>

		<tr><%
        	  ' ----------------------------------------------------------------------------
	 	  ' Displays checkbox OPTION TourGenTable shows data in EVENT2 field
        	  ' ----------------------------------------------------------------------------

	 	  IF TRIM(sTEvent2) <> "" AND TRIM(sTEvent2) <> "X" THEN %>
			<td><font  size=<% =fontsize2 %> ><%=sTEvent2Name%></td><%

			sEventNo=2
			sEvent2 = sTEvent2

			IF FormStatus = "modify" THEN 

				BoatStatus =""  %>
				<td><input type=checkbox name="fSelectEvent2" <% IF sSelectEvent2 = "on" THEN Response.Write("Checked") %>></font></td>
				<td><% LoadDivPulldown sDiv2, sEvent2 %></td>
				<td>&nbsp</td>

				<td>
				  <FONT COlOR="red" size=<% =fontsize2 %> >Boat: </FONT>
				   <% LoadBoatPulldown sEvent1 %>
				</td><%

			ELSEIF FormStatus = "confirm" AND sSelectEvent2 ="on" THEN %>
				 <td><font color="<% = textcolor2 %>" size=<% =fontsize2 %> >YES</font></td>
				 <td><font color="<% =textcolor2 %>" size=<% =fontsize2 %> ><%= sDiv2 %></td><%

				' LoadQualifications sEvent2, sDiv2   ' ---- Loads Qualifications and indicates if skier has required qualifications for this div/event  ---
				IF LEFT(QualStatEvent2,2) = "OK" THEN
					%><td><font color="<% = textcolor2 %>" size=<% =fontsize2 %> ><% response.write(QualStatEvent2&" - "&QualEvent2) %></font></td><%
				ELSE
					%><td><font color=red size=<% =fontsize2 %> ><% response.write(QualStatEvent2&" - "&QualEvent2) %></font></td><%
				END IF   

				BoatStatus ="disabled" %>   
				<td>
				  <FONT COlOR="<%=Textcolor1%>" size=<% =fontsize2 %> >Boat: </FONT>
				  <% LoadBoatPulldown sEvent2 %>
				</td><%
			ELSE  
				BoatStatus ="disabled"    %>
				<td>&nbsp</td>
				<td>&nbsp</td>
				<td>&nbsp</td>
				<td>&nbsp</td><%
			END IF  %>

			<td><font color="<% =textcolor2 %>" size=<% =fontsize2 %> >&nbsp;<%= CommentText2 %></td>

			  <% IF adminmenulevel >=19 THEN %>
			  <td>		
				<select name="sQfyOverride_E2" value="<% =sQfyOverride_E2 %>" <% =OverrideStatus %> >
				  <option value ="" <%IF sQfyOverride_E2 = "" THEN Response.Write(" selected ")%> >None</Option><br>
				  <option value ="EP" <%IF sQfyOverride_E2 = "EP" THEN Response.Write(" selected ")%> >Proof of EP</Option><br>
				  <option value ="3MS" <%IF sQfyOverride_E2 = "3MS" THEN Response.Write(" selected ")%> >3rd Masters</Option><br>
				  <option value ="OVR" <%IF sQfyOverride_E2 = "OVR" THEN Response.Write(" selected ")%> >Overall EP</Option><br>
				  <option value ="ALT" <%IF sQfyOverride_E2 = "ALT" THEN Response.Write(" selected ")%> >Alt Div Qfy</Option><br>
				  <option value ="OAP" <%IF sQfyOverride_E4 = "OAP" THEN Response.Write(" selected ")%> >Reg/Nat OA Plc</Option><br>
				  <option value ="OTH" <%IF sQfyOverride_E2 = "OTH" THEN Response.Write(" selected ")%> >Other Qfy</Option><br>
				  <option value ="DNS" <%IF sEntryType = "DNS" THEN Response.Write(" selected ")%> >Did Not Ski</Option><br>
				</select>
			  </td>		
			  <% ELSE %>
				<input type="hidden" name="sQfyOverride_E2" value="<% =sQfyOverride_E2 %>">		
			  <% END IF

		  END IF %>
		</tr>
	    	
	    	<tr><%

        	  ' ----------------------------------------------------------------------------
	 	  ' Displays checkbox OPTION TourGenTable shows data in EVENT3 field
        	  ' ----------------------------------------------------------------------------

		  IF TRIM(sTEvent3) <> "" AND TRIM(sTEvent3) <> "X" THEN %>
			<td><font  size=<% =fontsize2 %> ><%=sTEvent3Name%></td><%

			sEventNo=3
			sEvent3 = sTEvent3

			IF FormStatus = "modify" THEN %>
				<td><input type=checkbox name="fSelectEvent3" <% IF sSelectEvent3 = "on" THEN Response.Write("Checked") %>></font></td>
				<td><% LoadDivPulldown sDiv3, sEvent3 %></td>
				<td>&nbsp</td><%
				RampStatus ="" %>
				<td>
				    <font color="red" size=<% =fontsize2 %> >&nbsp;Ramp Height</font>
				    <font color="<% =textcolor2 %>" size=<% =fontsize2 %> >&nbsp;<% LoadRampPulldown sDiv3 %></font>
				    <font color="<% =textcolor1 %>" size=<% =fontsize2 %> >-Ft</font>
				</td><%


			ELSEIF FormStatus = "confirm" AND sSelectEvent3 ="on" THEN %>
				<td><font color="<% = textcolor2 %>" size=<% =fontsize2 %> >YES</font></td>
				<td><font color="<% =textcolor2 %>" size=<% =fontsize2 %> ><%= sDiv3 %></td><%

				' LoadQualifications sEvent3, sDiv3   ' ---- Loads Qualifications and indicates if skier has required qualifications for this div/event  ---
				IF LEFT(QualStatEvent3,2) = "OK" THEN
					%><td><font color="<% = textcolor2 %>" size=<% =fontsize2 %> ><% response.write(QualStatEvent3&" - "&QualEvent3) %></font></td><%
				ELSE
					%><td><font color=red size=<% =fontsize2 %> ><% response.write(QualStatEvent3&" - "&QualEvent3) %></font></td><%
				END IF  
				RampStatus ="disabled" %>
				<td>
				    <font color="<% =textcolor1 %>" size=<% =fontsize2 %> >&nbsp;Ramp Height</font>
				    <font color="<% =textcolor2 %>" size=<% =fontsize2 %> >&nbsp;<% LoadRampPulldown sDiv3 %></font>
				    <font color="<% =textcolor1 %>" size=<% =fontsize2 %> >-Ft</font>
				</td><%

			ELSE  %>
				<td>&nbsp</td>
				<td>&nbsp</td>
				<td>&nbsp</td>
				<td>&nbsp</td><%
			END IF  %>


			<td><font color="<% =textcolor2 %>" size=<% =fontsize2 %> >&nbsp;<%= CommentText3 %></td>

			  <% IF adminmenulevel >=19 THEN %>
			  <td>		
				<select name="sQfyOverride_E3" value="<% =sQfyOverride_E3 %>" <% =OverrideStatus %> >
				  <option value ="" <%IF sQfyOverride_E3 = "" THEN Response.Write(" selected ")%> >None</Option><br>
				  <option value ="EP" <%IF sQfyOverride_E3 = "EP" THEN Response.Write(" selected ")%> >Proof of EP</Option><br>
				  <option value ="3MS" <%IF sQfyOverride_E3 = "3MS" THEN Response.Write(" selected ")%> >3rd Masters</Option><br>
				  <option value ="OVR" <%IF sQfyOverride_E3 = "OVR" THEN Response.Write(" selected ")%> >Overall EP</Option><br>
				  <option value ="ALT" <%IF sQfyOverride_E3 = "ALT" THEN Response.Write(" selected ")%> >Alt Div Qfy</Option><br>
				  <option value ="OAP" <%IF sQfyOverride_E4 = "OAP" THEN Response.Write(" selected ")%> >Reg/Nat OA Plc</Option><br>
				  <option value ="OTH" <%IF sQfyOverride_E3 = "OTH" THEN Response.Write(" selected ")%> >Other Qfy</Option><br>
				  <option value ="DNS" <%IF sEntryType = "DNS" THEN Response.Write(" selected ")%> >Did Not Ski</Option><br>
				</select>
			  </td>		
			  <% ELSE %>
				<input type="hidden" name="sQfyOverride_E3" value="<% =sQfyOverride_E3 %>">		
		  	  <% END IF 

		  END IF %>
		</tr>

	    	<tr><%
        	  ' ----------------------------------------------------------------------------
	 	  ' Displays checkbox OPTION TourGenTable shows data in EVENT4 field
        	  ' ----------------------------------------------------------------------------

		  IF TRIM(sTEvent4)<>"" AND TRIM(sTEvent4) <> "X" THEN %>
			<td><font size=<% =fontsize2 %> ><%=sTEvent4Name%></td><%

			sEventNo=4


			IF FormStatus = "modify" THEN
				%><td><input type=checkbox name="fSelectEvent4" <% IF sSelectEvent4 = "on" THEN Response.Write("Checked") %>></font></td>
				<td><% LoadDivPulldown sDiv4, sEvent4 %></td>
				<td>&nbsp</td>
				<td>&nbsp</td>
				<td>&nbsp</td><%

			ELSEIF FormStatus = "confirm" AND sSelectEvent4 ="on" THEN 
				%><td><font color="<% = textcolor2 %>" size=<% =fontsize2 %> >YES</font></td>
				 <td><font color="<% =textcolor2 %>" size=<% =fontsize2 %> ><%= sDiv4 %></td><%

				' LoadQualifications sEvent4, sDiv4   ' ---- Loads Qualifications and indicates if skier has required qualifications for this div/event  ---
				IF LEFT(QualStatEvent4,2) = "OK" THEN
					%><td><font color="<% = textcolor2 %>" size=<% =fontsize2 %> ><% response.write(QualStatEvent4&" - "&QualEvent4) %></font></td><%
				ELSE
					%><td><font color=red size=<% =fontsize2 %> ><% response.write(QualStatEvent4&" - "&QualEvent4) %></font></td><%
				END IF   %>

			ELSE
				%><td>&nbsp</td>
				<td>&nbsp</td>
				<td>&nbsp</td>
				<td>&nbsp</td>
				<td>&nbsp</td><%
			END IF  %>

			<td><font color="<% =textcolor2 %>" size=<% =fontsize2 %> >&nbsp;<%= EV4_BRWText %></td>
			<td><font color="<% =textcolor2 %>" size=<% =fontsize2 %> >&nbsp;<%= CommentText4 %></td>


		  	<% IF adminmenulevel >=19 THEN %>
			  <td>		
				<select name="sQfyOverride_E4" value="<% =sQfyOverride_E4 %>" <% =OverrideStatus %> >
				  <option value ="" <%IF sQfyOverride_E4 = "" THEN Response.Write(" selected ")%> >None</Option><br>
				  <option value ="EP" <%IF sQfyOverride_E4 = "EP" THEN Response.Write(" selected ")%> >Proof of EP</Option><br>
				  <option value ="3MS" <%IF sQfyOverride_E4 = "3MS" THEN Response.Write(" selected ")%> >3rd Masters</Option><br>
				  <option value ="OVR" <%IF sQfyOverride_E4 = "OVR" THEN Response.Write(" selected ")%> >Overall EP</Option><br>
				  <option value ="ALT" <%IF sQfyOverride_E4 = "ALT" THEN Response.Write(" selected ")%> >Alt Div Qfy</Option><br>
				  <option value ="OAP" <%IF sQfyOverride_E4 = "OAP" THEN Response.Write(" selected ")%> >Reg/Nat OA Plc</Option><br>
				  <option value ="OTH" <%IF sQfyOverride_E4 = "OTH" THEN Response.Write(" selected ")%> >Other Qfy</Option><br>
				  <option value ="DNS" <%IF sEntryType = "DNS" THEN Response.Write(" selected ")%> >Did Not Ski</Option><br>
				</select>
			  </td>		
			  <% ELSE %>
				<input type="hidden" name="sQfyOverride_E4" value="<% =sQfyOverride_E4 %>">		
			  <% END IF 

	     	  END IF %>
    		</tr>
		</table> 

	  </div>
	</div>







 	<% ' ----------------------------------------------------------------------------------------------------------
	   ' --------------------------------  BEGIN FINANCIAL SECTION  -----------------------------------------------
	   ' ----------------------------------------------------------------------------------------------------------
	   ' ----------------------------------------------------------------------------------------------------------
	   ' -----------  Does NOT DISPLAY financial section UNLESS at least one event has been selected   ------------	

	%>
	<div class="accordionHeader">
		<a href="">STEP 4 - Financial Summary</a>
	</div>

	<div class="accordionContent" style="display:none;">
	  <div id="ctl10_UpdatePanel4">

	<% 'IF TotEvents > 0 THEN %>		

	  	<TABLE WIDTH=100% BORDER="4" CELLPADDING="3" CELLSPACING="0" width=100% BGCOLOR="<%=TableColor1%>">
		<tr align=center valign=center> 
		  <th align=center HEIGHT=135 WIDTH=30 nowrap rowspan=7 valign=center background="/images/buttons/Banner_Financial6.jpg"></th>		
		<%	 
		  IF FormStatus="modify" THEN
			%><td align="left" width=65% colspan=2><font color=red size=<% =fontsize2 %> >&nbsp;&nbsp;Check all that apply</font></td><%
		  ELSE
			%><td align="left" width=65% colspan=2><font size=<% =fontsize2 %> >&nbsp;&nbsp;Press 'Change Settings' to modify information</font></td><%
		  END IF  %>
	  	  <td width=20% align="right"><font color="<% =textcolor1 %>" size=<% =fontsize2 %> >Sub-Total Entry Fees</font></td><% 
		 

			

		 ' ---------------------   NEED TO DEAL WITH FAMILY MEMBERSHIP   ----------------------------


		  %><td width=15% align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> ><%=FormatCurrency(sTotalEntryFees,2)%></font></td>
		</tr><%

		' ---------------------------------------------
		' --------  LATE FEES --------------------------
		' ---------------------------------------------  %>

		<tr>
		  <% IF adminmenulevel >= 50 THEN %>
		    <td>
		    <FONT size=<% =fontsize2 %>  COlOR="<% =TextColor1 %>">Regionals Override</FONT>
			<select name="sRegionalOverride" value="<% =sRegionalOverride %>" <% =OverrideStatus %> >
			  <option value ="" <%IF sRegionalOverride = "" THEN Response.Write(" selected ")%> >None</Option><br>
			  <option value ="MED" <%IF sRegionalOverride = "MED" THEN Response.Write(" selected ")%> >Medical Excuse</Option><br>
			  <option value ="OTH" <%IF sRegionalOverride = "OTH" THEN Response.Write(" selected ")%> >Other</Option><br>
			</select>
		    </td>		
		    <td align="center">
		  <% ELSE %>
			  <td colspan=2 align="right">  
			  <input type="hidden" name="sRegionalOverride" value="<% =sRegionalOverride %>">
		  <% END IF

		    Dim MRDate	
		    MRDate = formatDateTime(sMembRegDate,2)
		    IF adminmenulevel >= 50 AND FormStatus = "modify" THEN  %>
			<font size=<% =fontsize2 %> >Date Entered (mm/dd/yyyy): &nbsp </font>
			<input type="text" name="sMembRegDate" value="<% =MRDate %>" MAXLENGTH=10 size="10" length="10"><%
		    ELSE  %>
			<font size=<% =fontsize2 %> >Date Entered</font>
			<font color="<% = textcolor2 %>" size="<% =fontsize2 %>" face="<% =font1 %>"><% =sMembRegDate %></font>
			<input type="hidden" name="sMembRegDate" value="<% =sMembRegDate %>"><%
		    END IF %>
		</td>  <%


'sLateDays = 2
'sLateFeeTot = 20
		IF Cint(sLateDays) > 0 AND cdbl(sTotalEntryFees) > 0 THEN %>
			  <td align="right"><font size=<% =fontsize2 %> >Late Fee - <%=sLateDays%> Days</font></td>
			  <td align="right"><font color="<% = textcolor2 %>" size=<% =fontsize2 %> ><%= FormatCurrency(sLateFeeTot,2) %></font></td>  <%
		ELSE
			%><td>&nbsp</td><%
			%><td>&nbsp</td><%
		END IF %>

		<tr><%

		  ' -------------------------------------------	
		  ' ---- Donation to AWSEF Building Fund  -----
		  ' -------------------------------------------

		  IF FormStatus="modify" THEN 
		       	  %><td colspan="2"><input type=checkbox name="fAWSEFCheck" <%IF sAWSEFCheck = "on" THEN Response.Write("Checked") %>>
		     	  <font size=<% =fontsize2 %> >&nbsp;&nbsp;Check to Donate $10.00 to AWSEF Building Fund</font></td><%
		  ELSEIF FormStatus="confirm" AND TRIM(sAWSEFCheck)="" THEN
			  %><td colspan="2"><font size=<% =fontsize2 %> >&nbsp;&nbsp;<b>NO</b>, I do not want to Donate $10.00 to the Building Fund</font></td><%
		  ELSE
			  %><td colspan="2"><font size=<% =fontsize2 %> >&nbsp;&nbsp;<b>YES</b>, I wish to Donate $10.00 to the Building Fund</font></td><%
		  END IF  

		  IF sAWSEFCheck = "on" THEN  
			  %><td align="right"><font size=<% =fontsize2 %> >Donation</font></td>
			  <td colspan=20% align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> ><%=FormatCurrency(sAWSEFDonation,2)%></font></td><%
		  ELSE	
			%><td>&nbsp</td><%
			%><td>&nbsp</td><%
		  END IF %>

	
		</tr><%
 

		  ' ------------------------------------------------------------	
		  ' ---- Discount to Junior B/G 1-3 per Tour_Manager.asp   -----
		  ' ------------------------------------------------------------


		  IF cdbl(sJrDiscPerc) > 0 AND sMembAge < 18 AND cdbl(sTotalEntryFees) > 0 THEN %>
			<tr>
			  <td colspan="2">&nbsp</td>
			  <td align="right"><font size=<% =fontsize2 %> >Junior Discount</font></td>
			  <td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> ><%= FormatCurrency(sJrDiscAmt,2) %></font></td>
			</tr><%
		  END IF 	

' --------------------------------------------------------------------------------------------------------------------
' ----  "NOTE:"  FUTURE - Make AGE for Senior Discount established by division setting in Tour and DivisionTable ----- 
' --------------------------------------------------------------------------------------------------------------------


		  ' -------------------------------------------------------------------------	
		  ' ---- Discount to divisions M/W-6 if specified in Tour_Manager.asp   -----
		  ' -------------------------------------------------------------------------

		  IF cdbl(sSrDiscPerc) > 0 AND sMembAge > 59 AND cdbl(sTotalEntryFees) > 0 THEN  %>
			<tr>
			  <td colspan="2">&nbsp</td>
			  <td align="right"><font size=<% =fontsize2 %> >Senior Discount</font></td>
			  <td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> ><%= FormatCurrency(sSrDiscAmt,2) %></font></td>
			</tr><%
		  END IF


'sOffDiscPerc = 10
'sOfficial = "on"
		' -------------------------------------------------------------------------	
		' ---------- Discount to OFFICIALS if specified in Tour_Manager.asp   -----
		' -------------------------------------------------------------------------  

		IF FormStatus="modify" AND sOffDiscPerc > 0 THEN %>
			<tr>	
			  <td colspan="2"><input type=checkbox name="fOfficial" <%IF sOfficial = "on" THEN Response.Write("Checked") %>>
		     	  <font size=<% =fontsize2 %> >&nbsp;&nbsp;Check if you are a rated official willing to work the tournament.</font></td>
			  <td>&nbsp;</td>
			  <td>&nbsp;</td><%  

		      ELSEIF FormStatus="confirm" AND TRIM(sOfficial)=""  AND sOffDiscPerc > 0 THEN  %>
			<tr>	
			  <td colspan="2"><font size=<% =fontsize2 %> >&nbsp;&nbsp;<b>NO</b>...I cannot work the tournament as a rated official.</font></td>
			  <td>&nbsp;</td>
			  <td>&nbsp;</td><%  
		      ELSEIF FormStatus="confirm" AND sOffDiscPerc > 0 THEN  %>
			<tr>
			  <td colspan="2"><font size=<% =fontsize2 %> >&nbsp;&nbsp;<b>YES</b>, I am a rated official willing to work the tournament.</font></td>
			<%  
		END IF

		IF cdbl(sOffDiscPerc) > 0 AND sOfficial = "on" THEN  %>
			<td align="right"><font color="<% =textcolor1 %>" size=<% =fontsize2 %> >Officials Discount</font></td>
			<td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> ><%= FormatCurrency(sOffDiscAmt,2) %></font></td>
			</tr><%	
		END IF 

		  ' -------------------------------------------------------------------------------------------------	
		  ' ---------- Discount to CLUB MEMBERS if match to ClubCode as specified in Tour_Manager.asp   -----
		  ' -------------------------------------------------------------------------------------------------  


'sClubDiscPerc = 10
'sClubMemb = "on"
'sClubCode = "123"
		      IF FormStatus="modify" AND sClubDiscPerc > 0 THEN %>
		       	  <tr> 
			     <td colspan="2"><input type=checkbox name="fClubMemb" <%IF sClubMemb = "on" THEN Response.Write("Checked") %>>
		     	  	<font size=<% =fontsize2 %> >&nbsp;&nbsp;Check if you are a Member of the Host Club.  CLUB CODE</font>
			  	<input type="text" name="fClubCode" value="<% =sClubCode %>" size="3" ></td>
				  <td>&nbsp;</td>
				  <td>&nbsp;</td><%  

		      ELSEIF FormStatus="confirm" AND TRIM(sClubMemb)="" AND sClubDiscPerc > 0 THEN   ' Checkbox NOT checked  %>
			  <tr>
			    <td colspan="2"><font size=<% =fontsize2 %> >&nbsp;&nbsp;<b>NO</b>...I am not a Member of the Host Club.</font></td>
				  <td>&nbsp;</td>
				  <td>&nbsp;</td><%  
		      ELSEIF FormStatus="confirm" AND TRIM(sClubCode) <> "" AND TRIM(sClubCode)<>TRIM(sTourClubCode) THEN %>
			  <tr> 
			    <td colspan="2"><font color="<% =textcolor1 %>" size=<% =fontsize2 %> > &nbsp;&nbsp;<b>YES</b>, I am a club Member but my</font> 
			  <font color="red" size=<% =fontsize2 %> > &nbsp;&nbsp;Club Code is Invalid</font></td><% 

		      ELSEIF  FormStatus="confirm" AND TRIM(sClubCode) <> "" AND TRIM(sClubCode)=TRIM(sTourClubCode) THEN  %>
			  <tr>
			    <td colspan="2"><font size=<% =fontsize2 %> >&nbsp;&nbsp;<b>YES</b>, I am a Member of the Host Club.</font></td><%
		      END IF 

	  	    	IF cdbl(sClubDiscPerc) > 0 AND sClubMemb = "on" AND cdbl(sTotalEntryFees) > 0 THEN
				IF TRIM(sClubCode) <> "" AND TRIM(sClubCode)=TRIM(sTourClubCode) THEN  %>
					<td align="right"><font size=<% =fontsize2 %> >Club Member Discount</font></td>
					<td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> ><%= FormatCurrency(sClubDiscAmt,2) %></font></td><%	
				ELSE
					%><td>&nbsp</td><td>&nbsp</td><%						
				END IF
			 ELSE   %>
				</td><%
			 END IF	%>
		  </tr>

		  <tr><%


			' -------------------------------------------------------------------------------------------------
			' -----  Calculate Applied Discount depending on which discount method was selected  --------------
			' -------------------------------------------------------------------------------------------------

			IF adminmenulevel >= 50 THEN  %>	
			    <td align=left>
			      <FONT size=<% =fontsize2 %>  COlOR="<% =TextColor1 %>">Fee Override</FONT>
				<select name="sMoneyOverride" value="<% =sMoneyOverride %>" <% =OverrideStatus %> >
				  <option value ="" <%IF sMoneyOverride = "" THEN Response.Write(" selected ")%> >None</Option><br>
				  <option value ="FAM" <%IF sMoneyOverride = "FAM" THEN Response.Write(" selected ")%> >Family Membership</Option><br>
				  <option value ="OTH" <%IF sMoneyOverride = "OTH" THEN Response.Write(" selected ")%> >Other</Option><br>
				</select>
			    </td><%
			ELSE %>
				<td>&nbsp;<input type="hidden" name="sMoneyOverride" value="<% =sMoneyOverride %>"></td><%
			END IF  


			IF cdbl(sJrDiscAmt)+cdbl(sSrDiscAmt)+cdbl(sClubDiscAmt)+cdbl(sOffDiscAmt) = 0 THEN
				%><td>&nbsp</td><%
			ELSEIF sDiscMeth ="M" AND cdbl(sTotalFormFees) > 0 THEN 
				%><td colspan="1" align="right"><font color="#000000" size=<% =fontsize2 %> >NOTE: Discount based on largest single discount (N/A to Late Fees)</font></td><%				   
			ELSE
				%><td colspan="1" align="right"><font color="#000000" size=<% =fontsize2 %> >NOTE: Cummulative discount does NOT apply to Late Fees !</font></td><%				   
			END IF %>

		    <td align="right"><font color="#000000" size=<% =fontsize2 %> >TOTAL ALL</font></td>
		    <td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> ><%= FormatCurrency(sTotalFormFees,2) %></font></td>
		  </tr>



		  <table border="1" width=35% align="right" CELLPADDING="3" CELLSPACING="0" width=100% BGCOLOR="<%=TableColor1%>">  <%

			IF cdbl(sTotalPreviousFees) < cdbl(sTotalFormFees) THEN %>
				<tr>
				  <td width=57% align="right"><font color="<% =textcolor1 %>" size=<% =fontsize2 %> >Previous Payments</font></td>
				  <td width=43% align="right"><font color="#000000" size=<% =fontsize2 %> ><%= FormatCurrency(cdbl(sTotalPreviousFees),2) %></font></td>
				</tr>
				
				<tr>
				  <td width=57% align="right"><font color="<% =textcolor3 %>" size="<% =fontsize2 %>" face="<% =font1 %>">BALANCE DUE</font></td>
				  <td width=43% align="right"><font color="<% =textcolor2 %>" size="<% =fontsize2 %>" face="<% =font1 %>"><%=(FormatCurrency(cdbl(sTotalFormFees)-cdbl(sTotalPreviousFees),2))%></font></td>
				</tr><%

			ELSEIF cdbl(sTotalPreviousFees) > cdbl(sTotalFormFees) THEN  %>
				<tr>
				  <td width=57% align="right"><font color="<% =textcolor1 %>" size=<% =fontsize2 %> >Previous Payments</font></td>
				  <td width=43% align="right"><font color="#000000" size=<% =fontsize2 %> ><%= FormatCurrency(cdbl(sTotalPreviousFees),2) %></font></td>
				</tr>
				<tr>
				  <td align="right"><font color="<% =textcolor3 %>" size=<% =fontsize2 %> >CREDIT DUE</font></td>
				  <td align="right"><font color="<% =textcolor3 %>" size=<% =fontsize2 %> ><%= FormatCurrency(cdbl(sTotalFormFees)-cdbl(sTotalPreviousFees),2) %></font></td>
				</tr><%
			ELSE


			END IF  %>

	
		</table>  
                            
	  </div>
	</div>



	<% '---------  WAIVER PAGE ------------  %>

	<div class="accordionHeader">
		<a href="">STEP 5 - Waiver</a>
	</div>
	<div class="accordionContent" style="display:none;">
           <div id="ctl12_UpdatePanel5">

		
             <TABLE BORDER="4" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="<% =TableColor1 %>" width=100%>
		     <TR>
			<TD BGCOLOR="red"><center><font  color="#FFFFFF" size="4"><b>Waiver and Release Form</b></font></TD>
		     </TR>  
 
		     <TR>
			<TD VALIGN="top">
  			   <TABLE BORDER="0" VALIGN="top" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="<%=TableColor1%>" width=100%>
				<tr>
				   <% ' ----- BEGINNING OF CELL -------- %>	
				   <td><%
		'END IF

		' ------------------------------------------------------------------------------------------
		' ----------  Displays release and sets FORM ACTION to rerun this condition  ---------------	
		' ------------------------------------------------------------------------------------------

		'DefineMemberVariables
		'DefineTourVariables

		   IF Session("MembAge") < 18 THEN
			ReleaseVersion = minor_waiver
			subTitle="Waiver for MINOR Participant - WaiverID: "&minor_waiver
		   ELSE
			ReleaseVersion = adult_waiver
			subTitle="Waiver for ADULT Participant - WaiverID: "&adult_waiver
		   END IF  %>
	

		<form action = "/rankings/RegistrationHQ.asp?sRunByWhat=release" method="post"><%

				
		IF adminmenulevel >= 50 THEN  %>
			<center> <%
			IF sReleaseType ="Electronic" THEN 
				%><input type=radio NAME="sReleaseType" VALUE="Electronic" ><%
			ELSE
				%><input type=radio NAME="sReleaseType" VALUE="Electronic" ><%
			END IF  %>
			<FONT size=<% =fontsize2 %>  COlOR=<% =textcolor1 %> ><b>Electronic</b></font>
			<input type=radio NAME="sReleaseType" VALUE="Paper">
			<FONT size=<% =fontsize2 %>  COlOR=<% =textcolor1 %> ><b>Paper W/Signature</b></font>
			<input type=radio NAME="sReleaseType" VALUE="None">
			<FONT size=<% =fontsize2 %>  COlOR=<% =textcolor1 %> ><b>No Waiver</b></font>
			</center>
			<br><br> <%
		ELSE  %>
			<INPUT type="hidden" NAME="sReleaseType" VALUE="Electronic" ><%
		END IF 	%>
	

		   <center>	
	 	   <font  size="4" ><b>PARTICIPANT WAIVER AND RELEASE OF LIABILITY,</b></font><br>
		   <font  size="4"><b>ASSUMPTION OF RISK AND INDEMNITY AGREEMENT</b></font>
		   <br>
		   <font  size="2"><b><% =subTitle %></b></font>
		   <br><br>

		   <font  color="<% =TextColor2 %>" size="3"><b><% =sTourName %></font></b>
		   <br><br>
		   <font  size="2"><b>MemberID = </font><font color="<% =textcolor2 %>"  size="2"><% =Session("sMemberID") %>
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="<% =textcolor1 %>"  size="2">Participant:</font>
			<font color="<% =textcolor2 %>"  size="2"><% =sFirstname %>&nbsp;<% =sLastName %></font></b><br>
		   </center><br>
		   <P><font color="<% =textcolor1 %>" size="1" ><left><%
	
		  Set objfso = CreateObject("Scripting.FileSystemObject")
		  IF objfso.FileExists(PathtoWaivers & "\waiver-"&ReleaseVersion&".txt") THEN
			SET objstream=objFSO.opentextfile(PathtoWaivers & "\waiver-"&ReleaseVersion&".txt")

			IF NOT objstream.atendofstream THEN
				DO WHILE not objstream.atendofstream
					response.write(objstream.readline)
				   	response.write("<br>")
				LOOP
			END IF

		  END IF
		  objstream.close  %>

		  </left></font></P>
			<center>
			<font size="4" color="red" ><b>The name listed above must be the person completing this form.</b></font>
			<br>
			<font size="4" color="red" ><b>Minors under 18 Years may NOT accept liability waiver.</b></font>
			</center>
		<center>
		<%
		   IF Session("MembAge") < 18 THEN  %>
			  <br><font color="<% =textcolor3 %>"  size="2"><b>Name of Parent or Guardian acccepting this waiver on behalf of this minor.</b></font>&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="fSignWaiver" value= "<% =sSignWaiver %>" size="30" ><%
		   ELSE  %>
			<br>
			<font color="<% =textcolor1 %>"  size="2"><b>By acccepting this waiver I acknowledge that I am the 'PARTICIPANT' listed above.</b></font><br><%
		   END IF %>
			<br><br>
		    	<font color="<% =textcolor1 %>"  size="2"><b>Accept:</b></font><input type="radio" name="fRelease" <%IF sRelease="Accept" THEN Response.write("checked")%> value="Accept">
		    	<font color="<% =textcolor1 %>"  size="2"><b>Decline:</b></font><input type="radio" name="fRelease" <%IF sRelease="Decline" THEN Response.write("checked")%> value="Decline">
			&nbsp;&nbsp;&nbsp;&nbsp
			<input type="submit" value="Submit Waiver">
			&nbsp;&nbsp;&nbsp;&nbsp
			<font color="<% =textcolor1 %>"  size="2"><b>Date: <% =DATE %></b></font>
		<br>
		</center>


		</form>

		</td> 

		<br>
		</tr>
		</TABLE>

		</TD></TR>
	     </TABLE>
                         

	  </div>
	</div>




	<div class="accordionHeader">
		<a href="">STEP 6 - Payment</a>
	</div>

	<div class="accordionContent" style="display:none;">
	  <div id="ctl14_UpdatePanel6">

                            <table>
                                <tr><td colspan="2"><b>THE ORGANIZING CLUB HEREBY REPRESENTS, CERTIFIES AND AGREES THAT:</b>
                                    <ol>
                                        <li>No competitor shall be permitted to ski in the tournament unless his/her USA Water Ski &quot;Active&quot; membership dues are paid in full.  A foreign skier may ski if he is a member of an IWSF-affiliated federation and can provide proof of USA Water Ski insurance coverage.</li>
                                        <li>The Organizing Club shall provide the site, equipment and the personnel required for the tournament.</li>
                                        <li>The tournament shall be operated in strict conformity with the Official Tournament Rules and  USA Water Ski policies and safety procedures.</li>
                                        <li>All USA Water Ski membership registration forms shall be properly documented and forwarded to USA Water Ski's Competition Department with all membership fees collected.</li>
                                        <li>The tournament shall be officiated exclusively by USA Water Ski-rated or IWSF-affiliated judges, drivers, scorers, safety directors and technical controllers/homologators.</li>
                                        <li>The <a href="http://www.usawaterski.org/pages/TournKit/AWSA/Event%20Organizer/Pre%20Tournament%20Safety%20Checklist.PDF" target="_blank"">USA Water Ski Tournament Organizer's Safety Checklist</a> will be followed where applicable.</li>
                                        <li>The undersigned is authorized to make these agreements on behalf of the Organizing Club.</li>
                                    </ol></td></tr>
                             

			<% ' --- SOME BUTTONS HERE --- %>

                            </table>
                            
	  </div>
	</div>                        

  </td></tr>
</table>
</div>




</form>
</body>
</html>

<%

END SUB


%>





