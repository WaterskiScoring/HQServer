<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include16.asp"-->
<!--#include virtual="/rankings/tools_registration16.asp"-->
<!--#include virtual="/rankings/tools_Definitions.asp"-->
<!--#include virtual="/rankings/RegFormDisplay16.asp"-->

<style>

	.span5 { width:5%; display:inline-block; }
	.span7 { width:7%; display:inline-block; }	
	.span10 { width:10%; display:inline-block; }
	.span15 { width:15%; display:inline-block; }
	.span20 { width:20%; display:inline-block; }
	.span25 { width:25%; display:inline-block; }
	.span30 { width:30%; display:inline-block; }
	.span35 { width:35%; display:inline-block; }
	.span40 { width:40%; display:inline-block; }
	.span45 { width:45%; display:inline-block; }
	.span50 { width:50%; display:inline-block; }
	.span55 { width:55%; display:inline-block; }
	.span60 { width:60%; display:inline-block; }
	.span65 { width:65%; display:inline-block; }
	.span70 { width:70%; display:inline-block; }
	.span75 { width:75%; display:inline-block; }
	.span80 { width:80%; display:inline-block; }	
	.span85 { width:85%; display:inline-block; }
	.span90 { width:90%; display:inline-block; }
	.span95 { width:95%; display:inline-block; }
	.span100 { width:100%; display:inline-block}
	
	
.details {
		font-size:11px;
		font-family: Arial, Sans-Serif;
		margin:0px 0px 0px 0p;
    padding:0px 2px 0px 5px;
    border-bottom-left-radius:10px;
    border-bottom-right-radius:10px;    
    text-align:left;
    width:96%;
    min-height:300px;
    height:auto;
	} 
	
.detailline {
		padding-top:5px;
		padding-bottom:5px;			
		verical-align:top;
	}	

.sectiondiv {
	border: 1px solid black;
	
	}
	
.headline {
	text-align: center;
	border: 1px solid #0f77da;
  color: white;
  background-color: #0f77da;
	font-family: Arial, Sans-Serif;
	font-size: 12pt;
	font-weight: bold;
	padding: 5px;
 	margin-top: 5px;
  width:100%;		
	
	}	
	
</style>
<%


' HQSiteColor1="#203f5e"
' HQSiteColor2="#0f77da"
' HQSiteColor3="#2F4F4F"

Dim RegFileName, CardFileName, DisplayFileName
Dim action, ClubID, TID, sOrderNo, CCAmount, sLP, HQF, OLR, LF, PA, IW, RF, PF




RegFileName="Test_CC.asp"
CardFileName="CCSJ_Sanctions.asp"


action = "new"
ClubID = "000001151"
TID = "20N001"
' sOrderNo = "5000000"
CCAmount = "5"
' sLP = "Test"
sLP = ""
HQF = "3"
OLR = "0" 
LF = "0"
PA = "0"
IW = "0"
RF = "2"
PF = "0"



' -- Increments OrderNo by one
' -- Assumes multiple people will not be using this page simultaneously 

sSQL = "SELECT MAX(OrderNo)+1 AS sOrderNo"
sSQL = sSQL + " FROM [usawsrank].[SanctionPaymentLog]"
sSQL = sSQL + " WHERE OrderNo BETWEEN 5000000 AND 5000099"

' response.write (sSQL)
' response.end


SET rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, sConnectionToTRATable, 3, 1
	
sOrderNo = ""
IF NOT(rs.EOF) THEN
	sOrderNo = rs("sOrderNo")	
END IF	

rs.close



WriteIndexPageHeader




Dim href
href = "http://www.usawaterski/rankings/"&CardFileName&"?action="&action&"&ClubID="&ClubID&"&TID="&TID&"&sOrderNo="&sOrderNo&"&CCAmount="&CCAmount&"&sLP="&sLP&"&HQF="&HQF&"&OLR="&OLR&"&LF="&LF&"&PA="&PA&"&IW="&IW&"&RF="&RF	

TestValues="N"
IF TestValues="Y" THEN 
		response.write("<br>TOP OF PROGRAM")
		response.write("<br>Action="&Action)
		response.write("<br>ClubID="&ClubID)
		response.write("<br>TID="&TID)
		response.write("<br>sOrderNo="&sOrderNo)
		response.write("<br>CCAmount="&CCAmount)
		response.write("<br>sLP="&sLP)
		response.write("<br>HQF="&HQF)
		response.write("<br>OLR="&OLR)
		response.write("<br>LF="&LF)
		response.write("<br>PA="&PA)
		response.write("<br>IW="&IW)
		response.write("<br>RF="&RF)
		response.write("<br>PF="&PF)			
		response.write("<br><br>href="&href)

END IF





%>

<form action="/rankings/<%=CardFileName%>" method=get>
<input type="hidden" id="sOrderNo_hide" name="sOrderNo" value="<%=sOrderNo%>">
<input type="hidden" id="action" name="action" value="new">
<input type="hidden" id="sLP" name="sLP" value="<%=sLP%>">



<div id="sectiondiv">
	<div id="headline" class="headline">		
		NSSA Sanctioning Payment Processing
	</div>  


	<div class="details" style="background-color:#FFFFFF; margin:25px;">
		<div style="border:1px solid blue; padding-bottom:10px;">
			<div class="detailline" style="padding-left:10px; background-color:#0f77da; color:#FFFFFF;">
				<span class="span25" style="text-align:left; font-size:12pt; font-weight:bold;">REQUIRED</span>
			</div>
			<div class="detailline" style="margin-top:10px;">
				<span class="span25" style="text-align:right">ClubID:</span>
				<span class="span70"><input type="text" id="ClubID" name="ClubID" value="<%=ClubID%>" maxlegnth="11"></span>
			</div>
			<div class="detailline">
				<span class="span25" style="text-align:right">SanctionID:</span>
				<span class="span70"><input type="text" id="TID" name="TID" value="<%=TID%>"></span>
			</div>
			<div class="detailline">
				<span class="span25" style="text-align:right">Order #:</span>
				<span class="span70"><input type="text" id="sOrderNo" name="sOrderNo" value="<%=sOrderNo%>" disabled></span>
			</div>

			<div class="detailline">
				<span class="span25" style="text-align:right">Total Fees $:</span>
				<span class="span70"><input type="text" id="CCAmount" name="CCAmount" value="<%=CCAmount%>"></span>
			</div>
			
		</div>

		<div style="border:1px solid blue; margin-top:15px; padding-bottom:10px;">
			<div class="detailline" style="padding-left:10px; background-color:#0f77da; color:#FFFFFF;">
				<span class="span70" style="text-align:left; font-size:12pt; font-weight:bold;">OPTIONAL DETAILS</span>
			</div>
			<div class="detailline" style="margin-top:10px;">
				<span class="span25" style="text-align:right">HQ Fee $:</span>
				<span class="span70"><input type="text" id="HQF" name="HQF" value="<%=HQF%>"></span>
			</div>
			<div class="detailline">
				<span class="span25" style="text-align:right">Late Fee $:</span>
				<span class="span70"><input type="text" id="LF" name="LF" value="<%=LF%>"></span>
			</div>
			<div class="detailline">
				<span class="span25" style="text-align:right">PanAm Fee $:</span>
				<span class="span70"><input type="text" id="PA" name="PA" value="<%=PA%>"></span>
			</div>
			<div class="detailline">
				<span class="span25" style="text-align:right">IWWF Fee $:</span>
				<span class="span70"><input type="text" id="IW" name="IW" value="<%=IW%>"></span>
			</div>
			<div class="detailline">
				<span class="span25" style="text-align:right">Region Fee $:</span>
				<span class="span70"><input type="text" id="RF" name="RF" value="<%=RF%>"></span>
			</div>
			<div class="detailline">
				<span class="span25" style="text-align:right">Practice Fee $:</span>
				<span class="span70"><input type="text" id="PF" name="PF" value="<%=PF%>"></span>
			</div>


		</div>

	</div>	

	<div style="margin-top:5px; font-weight:bold;">INSTRUCTIONS:</div>
	<div style="font-size:8pt;"> 1) Sanction ID is not generated by Sanction system and standard format must be 20N001, 20N002, etc where N is in 3rd position and</div>
	<div style="font-size:8pt;">  the 2-digit year is in the first two positions as is the standard for USAWS sanctioning.</div>
	<div style="font-size:8pt;"> 2) Total Fee must be an even dollar amount. Do not use $ sign or decimal - ex: use 10 for $10.</div>
	<div style="font-size:8pt;"> 3) Details of payment are for references only.  No validation is made on details or whether they match the Total Amount. </div>
	<div style="font-size:8pt;"> The details entered will only appear in confirmation email. </div>
	<div style="font-size:8pt;"> 4) Only one staff member should use the processor at any one time.</div>

	<div style="width:100%; margin-top:20px; margin-left:0px; text-align:center; padding-bottom:20px;">

			<span class="gentext" style="text-align:center;">
				<input type="submit" name="CreditCard" value="Continue" style="width:9em" title="Click here to proceed to secure Payment Page">
			</span>
	</div>


</div>


</form>






<%

WriteIndexPageFooter

					  
%>

					  			    