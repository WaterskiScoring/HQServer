<%

' --- This is the display program for the Registration Module
' --- Originally intended to be only displays to limit computation and logic in this module.
' --- Written by Mark Crone


' ---------------------------
   SUB DisplayAccordion
' ---------------------------


%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
  <link href="http://www.usawaterski.org/css/styles.css" rel="stylesheet" type="text/css" />

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

</style>
<title>
	Online Registration Display Page
</title>

</head>
<body><%



Dim AdminButtonStyle
AdminButtonStyle="width:9em; background-color:red; color:white"
UserButtonStyle="width:9em"
SpecialUserButtonStyle="width:9em; background-color:yellow; color:black"

IF DisplayVars="on" THEN DisplayPertinentVariables



OFDescArray = Array (sOF1Desc,sOF2Desc,sOF3Desc,sOF4Desc,sOF5Desc,sOF6Desc,sOF7Desc,sOF8Desc,sOF9Desc,sOF10Desc)
OFAmtArray = Array (sOF1Amt,sOF2Amt,sOF3Amt,sOF4Amt,sOF5Amt,sOF6Amt,sOF7Amt,sOF8Amt,sOF9Amt,sOF10Amt)
OFRequiredArray = Array (sOF1Required,sOF2Required,sOF3Required,sOF4Required,sOF5Required,sOF6Required,sOF7Required,sOF8Required,sOF9Required,sOF10Required)			
OFMaxQtyArray = Array (sOF1MaxQty,sOF2MaxQty,sOF3MaxQty,sOF4MaxQty,sOF5MaxQty,sOF6MaxQty,sOF7MaxQty,sOF8MaxQty,sOF9MaxQty,sOF10MaxQty)			
OFQtyArray = Array (sOF1Qty,sOF2Qty,sOF3Qty,sOF4Qty,sOF5Qty,sOF6Qty,sOF7Qty,sOF8Qty,sOF9Qty,sOF10Qty)
OFFeeArray = Array (sOF1Fee,sOF2Fee,sOF3Fee,sOF4Fee,sOF5Fee,sOF6Fee,sOF7Fee,sOF8Fee,sOF9Fee,sOF10Fee)






	' ----------------------------------------------------------------------
	' -----------------------   TOURNAMENT INFORMATION   -----------------------
	' ----------------------------------------------------------------------
'response.write("<br>TEST RegDisp 91 - sTourID="&sTourID)

%>


<div id="Accordion1">
  <div class="<% IF nav=1 THEN response.write("accordionHeaderSelected") ELSE response.write("accordionHeader") END IF %>">
		<TABLE>
			<TR>
				<td width="150px" align="left">
					<a href="/rankings/<%=RegFileName%>?nav=1" title="TourID: <%=sTourID%>">STEP 1 - Tournament</a>
				</td>
				<td align="left">
					<font style="color:Yellow;"><%=sTourName%></font>
				</td>
			</TR>
		</TABLE>
  </div>
  
  <div class="innertable" style="display:<% IF nav=1 THEN response.write("block") ELSE response.write("none") END IF %>;">
    <div id="RegPanel1">	 
  	  <br>
    	  <table class="spacetable" width=100%>
					<tr><td align="left" colspan=7><br><b>&nbsp;TOURNAMENT DETAILS</b></td><tr>
					<tr>  
	  				<TD ALIGN="right" width=120px><FONT COlOR="<%=TextColor1%>" size=<% =fontsize2 %>>Tour Name&nbsp;</FONT></TD>
	  				<TD ALIGN="left" width=300px colspan=4 width=300px><FONT COlOR="<%=TextColor2%>" size=<% =fontsize2 %> >&nbsp;<% =sTourName %></FONT></TD>
	  				<TD Align="right" width=100px><FONT COlOR="<% =textcolor1 %>" size="<% =fontsize2 %>">Registrar&nbsp;</FONT></TD>
	  				<TD align="left" width=200px><FONT COlOR="<% =textcolor2 %>" size="<% =fontsize2 %>">&nbsp;<%=sTRegistrarName%></FONT></TD>
					</tr>
	
					<tr>  
	  				<TD ALIGN="right" vAlign="top"><FONT COlOR="<%=TextColor1%>" size=<% =fontsize2 %>  >Tour ID&nbsp;</FONT></TD>
	  				<TD ALIGN="left" colspan=4 ><FONT COlOR="<%=TextColor2%>" size=<% =fontsize2 %>  >&nbsp;<% =sTourID %></FONT></TD>
	  				<TD Align=right><FONT COlOR="<% =textcolor1 %>" size="<% =fontsize2 %>">Address&nbsp;</FONT></TD>
	  				<TD align=left><FONT COlOR="<% =textcolor2 %>" size="<% =fontsize2 %>">&nbsp;<%=sTRegistrarAddr%></FONT></TD>
					</tr>

	<tr>  
	  <TD ALIGN="right"><FONT color="<%=TextColor1%>" size=<% =fontsize2 %>>City/ST&nbsp;</FONT></TD>
	  <TD ALIGN="left" colspan=4><FONT color="<%=TextColor2%>" size=<% =fontsize2 %>>&nbsp;<% =sTourCity&", "&sTourState %></FONT></TD>
	  <TD ALIGN="left">&nbsp;</td>		
	  <TD align=left><FONT COlOR="<% =textcolor2 %>" size="<% =fontsize2 %>">&nbsp;<%=sTRegistrarCity%>, <%=sTRegistrarState%>&nbsp;<%=sTRegistrarZip %></FONT></TD>
	</tr>
	
	<tr>  
	  <TD ALIGN="right"><FONT COlOR="<%=TextColor1%>" size=<% =fontsize2 %> >Dates&nbsp;</FONT></TD>
	  <TD ALIGN="left" colspan=4><FONT COlOR="<%=TextColor2%>" size=<% =fontsize2 %> >&nbsp;<% =sTDateS&"-"&sTDateE %></FONT></TD>
	  <TD Align=right><FONT COlOR="<% =textcolor1 %>" size="<% =fontsize2 %>">Phone&nbsp;</FONT></TD>
	  <TD Align="left"><FONT COlOR="<% =textcolor2 %>" size="<% =fontsize2 %>">&nbsp;<%=sTRegistrarPhone%></FONT></TD>
	</tr>
	
	<tr>  
	  <TD ALIGN="right"><FONT COlOR="<%=TextColor1%>" size=<% =fontsize2 %>>SptDiv&nbsp;</FONT></TD>
	  <TD ALIGN="left" colspan=4><FONT COlOR="<%=TextColor2%>" size=<% =fontsize2 %> >&nbsp;<%=sTSptsGrpID%></FONT></TD>
	  <TD Align=right><FONT COlOR="<% =textcolor1 %>" size="<% =fontsize2 %>">Email&nbsp;</FONT></TD>
	  <TD Align="left"><FONT COlOR="<% =textcolor2 %>" size="<% =fontsize2 %>">&nbsp;<%=sTRegistrarEmail%></FONT></TD>
	</tr>

	<tr>		
	  <td>&nbsp;</td>	 
	  <td colspan=4>&nbsp;</td>	 
	  <TD ALIGN="left">&nbsp;</td>		
	  <TD ALIGN="left">&nbsp;</td>		
 	 </tr>
       <TR>
	<TD align=right><FONT COlOR="<% =textcolor1 %>" size="<% =fontsize2 %>" >Description</font></td>

	<td colspan=4 align=left style="word-wrap:break-word">
		<FONT COlOR="<% =textcolor2 %>" size="<% =fontsize2 %>" >&nbsp; <%
		     	IF TRIM(sTDescription)<>"" THEN ThisDescription = TRIM(sTDescription) + "<br>" 
		     	IF TRIM(sFDescription)<>"" THEN ThisDescription = TRIM(sFDescription) + "<br>" 
		     	IF TRIM(sWDescription)<>"" THEN ThisDescription = TRIM(sWDescription) + "<br>" 
		     	IF TRIM(sKDescription)<>"" THEN ThisDescription = TRIM(sKDescription) + "<br>" 
		     	IF TRIM(sCDescription)<>"" THEN ThisDescription = TRIM(sCDescription) + "<br>" 
			response.write(ThisDescription)  %>
			
		</font>
	</td>
	 <TD>&nbsp;</TD>		
	 <TD>&nbsp;</TD>		
       </tr>	 

       <TR>
       	 <TD Align=right><FONT COlOR="<% =textcolor1 %>" size="<% =fontsize2 %>" face=<% =font1 %>>Divisions</FONT></TD>
         <TD colSpan="4" align=left><FONT COlOR="<% =textcolor2 %>" size="<% =fontsize2 %>">&nbsp;<%=sTDvOffered%></FONT></TD>
	 <TD>&nbsp;</TD>		
	 <TD>&nbsp;</TD>		
       </tr>

       <TR>
       	 <TD Align=right><FONT COlOR="<% =textcolor1 %>" size="<% =fontsize2 %>" face=<% =font1 %>>Directions</FONT></TD>
         <TD align=left colSpan="6"><FONT COlOR="<% =textcolor2 %>" size="<% =fontsize2 %>" face=<% =font1 %>>&nbsp;<%=GTSDirections %></FONT></TD>
       </TR>
       <TR>
         <TD align="right"><FONT COlOR="<% =textcolor1 %>" size="<% =fontsize2 %>">Schedule:</FONT></TD>
         <TD align="left" colSpan="6"><FONT COlOR="<% =textcolor2 %>" size="<% =fontsize2 %>">&nbsp;<%= GTSofE %></FONT></TD>
       </tr>
       <TR>
         <TD align="right"><FONT COlOR="<% =textcolor1 %>" size="<% =fontsize2 %>">Comments:</FONT></TD>
         <TD align="left" colSpan="6"><FONT COlOR="<% =textcolor2 %>" size="<% =fontsize2 %>">&nbsp;<%= GTComments %></FONT></TD>
       </tr>


	<tr><td align="left" colspan=7><br><b>&nbsp;ENTRY FEES</b></td><tr><%

	' --- Sets column width for fees ---
	sClassWidth=83



'response.write("<br>sGREntryFee1 = "&sGREntryFee1)
'response.write("<br>sGREntryFee2 = "&sGREntryFee2)
'response.write("<br>sGREntryFee3 = "&sGREntryFee3)



	  %>
	  <tr>	
	    <TD width=120px>&nbsp;</TD><%
		IF Cdbl(sGREntryFee1)>cdbl(0.00) OR Cdbl(sGREntryFee2)>cdbl(0.00) OR Cdbl(sGREntryFee3)>cdbl(0.00) THEN %>
			<td align="center" width="<%=sClassWidth%>px"><font size="<%=fontsize2%>"><b>&nbsp;<%=sTGRClassText%></font></td>
			<td align="center" width="<%=sClassWidth%>px"><font size="<%=fontsize2%>"><b><%=sTBaseClassText%></b></font></td><%
		ELSE %>
			<td width="<%=sClassWidth%>px">&nbsp;</td>
			<td align="center" width="<%=sClassWidth%>px"><font size="<%=fontsize2%>"><b><%=sTBaseClassText%></b></font></td><%
		END IF 			
		IF sRSurchg>0 THEN %>
			<td align="center" width="<%=sClassWidth%>px"><font size="<%=fontsize2%>"><b>&nbsp;<%=sTUpgradeClassText%></b></font></td><%
		ELSE %>
			<TD ALIGN="center" width="<%=sClassWidth%>px">&nbsp;</td><%
		END IF %>
		<td width=100px>&nbsp;</td>
		<td width=100px>&nbsp;</td>
		<td align="left" width="150px"><font size="<%=fontsize2%>"><b>&nbsp;Deadline</b></font></td>
	</tr>

	<tr><%


		IF sTPandC<>true THEN %> 	  
			<TD ALIGN="right"><FONT COlOR="<%=TextColor1%>" size=<%=fontsize2%>>1 Event &nbsp;</FONT></TD><%
		ELSE %>
			<TD ALIGN="right"><FONT COlOR="<%=TextColor1%>" size=<%=fontsize2%>>1st Ride/Event&nbsp;</FONT></TD><%
		END IF 

		IF Cdbl(sGREntryFee1)>cdbl(0.00) THEN %>
			<TD ALIGN="center"><FONT COlOR="<%=TextColor2%>" size=<% =fontsize2 %> >&nbsp;<% =FormatCurrency(sGREntryFee1,2) %></FONT></TD><%
		ELSE %>
			<TD ALIGN="center">&nbsp;</td><%
		END IF

		IF cdbl(sTEntryFee1)>cdbl(0.00) THEN %>
			<TD ALIGN="center"><FONT COlOR="<%=TextColor2%>" size=<% =fontsize2 %> >&nbsp;<% =FormatCurrency(sTEntryFee1,2) %></FONT></TD><%
		ELSE %>
			<TD ALIGN="center">&nbsp;</td><%
		END IF

		IF sRSurchg>0 THEN %>
			<TD ALIGN="center"><FONT COlOR="<%=TextColor2%>" size=<% =fontsize2 %> >&nbsp;<% =FormatCurrency(sTEntryFee1+sRSurchg,2) %></FONT></TD><%		
		ELSE %>
			<TD ALIGN="center">&nbsp;</td><%
		END IF %>

		<td>&nbsp;</td><%

		IF sTLFPerDay=true AND sTLateFee>cdbl(0.00) THEN %>
			<TD ALIGN="right"><FONT COlOR="<%=TextColor1%>" size=<% =fontsize2 %> >Late Fee&nbsp;</FONT>
			<TD vAlign="top"><FONT COlOR="<% =textcolor2 %>" size="<% =fontsize2 %>" face=<% =font1 %>>&nbsp;<%=FormatCurrency(sTLateFee,2)%> Per Day</FONT></TD><%
		ELSEIF sTLFPerDay<>true AND sTLateFee>0.00 THEN  %>
			<TD ALIGN="right"><FONT COlOR="<%=TextColor1%>" size=<% =fontsize2 %> >Late Fee&nbsp;</FONT>
			<TD vAlign="top"><FONT COlOR="<% =textcolor2 %>" size="<% =fontsize2 %>" face=<% =font1 %>>&nbsp;<%=FormatCurrency(sTLateFee,2)%></FONT></TD><%
		ELSE  %>
			<TD>&nbsp;</td>		
			<TD>&nbsp;</td><%	
		END IF  %>
	</tr>
	<tr><%
		IF sTPandC<>true AND (Cdbl(sTEntryFee2)>cdbl(0.00) OR Cdbl(sGREntryFee2)>cdbl(0.00) ) THEN  
				%> 	  	
				<TD ALIGN="right"><FONT COlOR="<%=TextColor1%>" size=<%=fontsize2%>>2 Events &nbsp;</FONT></TD><%
		ELSEIF Cdbl(sTEntryFee2)>0.00 THEN %>
				<TD ALIGN="right"><FONT COlOR="<%=TextColor1%>" size=<%=fontsize2%>>2nd Ride/Events&nbsp;</FONT></TD><%
		ELSE  %>
				<TD>&nbsp;</td><%		
		END IF 

'IF MarkTester THEN response.write("<br>sGREntryFee2="& sGREntryFee2)
'IF MarkTester THEN response.write("<br>sTEntryFee2="& sTEntryFee2)

		IF Cdbl(sGREntryFee2)>cdbl(0.00) THEN %>		
			<TD ALIGN="center"><FONT COlOR="<%=TextColor2%>" size=<% =fontsize2 %> >&nbsp;<% =FormatCurrency(sGREntryFee2-sGREntryFee1,2) %></FONT></TD><%
			 'response.write("GR2="&sGREntryFee2) 
		ELSE %>
			<TD>&nbsp;</td><%				
		END IF

		IF Cdbl(sTEntryFee2)>0.00 THEN %>
			<TD ALIGN="center"><FONT COlOR="<%=TextColor2%>" size=<% =fontsize2 %> >&nbsp;<% =FormatCurrency(sTEntryFee2,2) %></FONT></TD><%		
		ELSE %>
			<TD>&nbsp;</td><%				
		END IF

		' --- Was sTEntryFee2+sRSurchg until 7/16/2014 ---
		IF sRSurchg>0 AND Cdbl(sTEntryFee2)>cdbl(0.00) THEN %>
			<TD ALIGN="center"><FONT COlOR="<%=TextColor2%>" size=<% =fontsize2 %> >&nbsp;<% =FormatCurrency(sTEntryFee2,2) %></FONT></TD><%		
		ELSE %>
			<TD ALIGN="center">&nbsp;</td><%
		END IF %>

		<td>&nbsp;</td>
	  	<TD ALIGN="right"><FONT COlOR="<%=TextColor1%>" size="<% =fontsize2 %>">Reg Deadline&nbsp;</FONT></td>
		<TD ALIGN="left"><FONT COlOR="<%=TextColor2%>" size=<% =fontsize2 %> >&nbsp;<% =sTLateDate %></FONT></td>
	</tr>
	<tr><%
		IF sTPandC<>true AND ( Cdbl(sTEntryFee3)>0.00 OR Cdbl(sGREntryFee3)>0.00 ) THEN %> 	  
			<TD ALIGN="right"><FONT COlOR="<%=TextColor1%>" size=<%=fontsize2%>>3+ Events&nbsp;</FONT></TD><%
		ELSEIF Cdbl(sTEntryFee3)>0.00 THEN %>
			<TD ALIGN="right"><FONT COlOR="<%=TextColor1%>" size=<%=fontsize2%>>3+ Per Ride/Events&nbsp;</FONT></TD><%
		ELSE  %>
			<TD>&nbsp;</td><%		
		END IF 

		IF Cdbl(sGREntryFee3)>0.00 THEN %>
			<TD ALIGN="center"><FONT COlOR="<%=TextColor2%>" size=<% =fontsize2 %> >&nbsp;<% =FormatCurrency(sGREntryFee3-sGREntryFee1,2) %></FONT></TD><%		
		ELSE %>
			<TD ALIGN="center">&nbsp;</td><%
		END IF

		IF Cdbl(sTEntryFee3)>cdbl(0.00) THEN %>
			<TD ALIGN="center"><FONT COlOR="<%=TextColor2%>" size=<% =fontsize2 %> >&nbsp;<% =FormatCurrency(sTEntryFee3,2) %></FONT></TD><%		
		ELSE %>
			<TD ALIGN="center">&nbsp;</td><%
		END IF
		' --- Was sTEntryFee3+sRSurchg until 7/16/2014 
		IF sRSurchg>0 AND Cdbl(sTEntryFee3)>cdbl(0.00) THEN %>
			<TD ALIGN="center"><FONT COlOR="<%=TextColor2%>" size=<% =fontsize2 %> >&nbsp;<% =FormatCurrency(sTEntryFee3,2) %></FONT></TD><%		
		ELSE %>
			<TD ALIGN="center">&nbsp;</td><%
		END IF %>

		<td>&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	<%
	IF sTEntryFeeFamily>cdbl(0.00) THEN 
			%> 	  
			<tr>
				<TD ALIGN="right"><FONT COlOR="<%=TextColor1%>" size=<%=fontsize2%>>Family<br> (Base-<%=sMaxFamMembers%>)&nbsp;</FONT></TD><%
				IF sGrassOffered=true THEN %>
						<TD ALIGN="center"><FONT COlOR="<%=TextColor2%>" size=<% =fontsize2 %> >&nbsp;<% =FormatCurrency(sTEntryFeeFamily,2) %></FONT></TD><%
				ELSE %>
						<TD ALIGN="center">&nbsp;</td><%
				END IF
				%>
				<TD ALIGN="center"><FONT COlOR="<%=TextColor2%>" size=<% =fontsize2 %> >&nbsp;<% =FormatCurrency(sTEntryFeeFamily,2) %></FONT></TD>
				<%
				IF sRSurchg>0 AND sTEntryFeeFamily>cdbl(0.00) THEN %>
						<TD ALIGN="center"><FONT COlOR="<%=TextColor2%>" size=<% =fontsize2 %> >&nbsp;<% =FormatCurrency(sTEntryFeeFamily,2) %></FONT></TD><%
				ELSE %>
						<TD ALIGN="center">&nbsp;</td><%
				END IF %>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			</tr>
			<%
	END IF		
	IF sTEntryFeeFamExtra>cdbl(0.00) THEN	
			%>
			<tr>
				<TD ALIGN="right"><FONT COlOR="<%=TextColor1%>" size=<%=fontsize2%>>Family Fee (Addl) &nbsp;</FONT></TD>
				<%
				IF sGrassOffered=true THEN %>
						<TD ALIGN="center"><FONT COlOR="<%=TextColor2%>" size=<% =fontsize2 %> >&nbsp;<% =FormatCurrency(sTEntryFeeFamExtra,2) %></FONT></TD><%
				ELSE %>
						<TD ALIGN="center">&nbsp;</td><%
				END IF
				%>
				<TD ALIGN="center"><FONT COlOR="<%=TextColor2%>" size=<% =fontsize2 %> >&nbsp;<% =FormatCurrency(sTEntryFeeFamExtra,2) %></FONT></TD>
				<%
				IF sRSurchg>0 THEN %>
						<TD ALIGN="center"><FONT COlOR="<%=TextColor2%>" size=<% =fontsize2 %> >&nbsp;<% =FormatCurrency(sTEntryFeeFamExtra,2) %></FONT></TD><%
				ELSE %>
						<TD ALIGN="center">&nbsp;</td><%
				END IF %>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			</tr>
			<%
	END IF		

	IF sTPandC=true AND Cdbl(sGREntryFee1)>cdbl(0.00) THEN 
			%>
			<tr>
				<td ALIGN="center" colspan=4><FONT COlOR="<%=TextColor2%>" size="<%=fontsize2%>"><b>NOTE:</b> Grassroots entry fee by # of events entered not rounds</FONT></td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
			</tr>
			<%		
	END IF		

	' --- If any Optional Fees then display the Option Fee HEADING and DATA ---
	IF TRIM(sOF1Desc)<>"" OR TRIM(sOF2Desc)<>"" OR TRIM(sOF3Desc)<>"" OR TRIM(sOF4Desc)<>"" OR TRIM(sOF5Desc)<>"" OR TRIM(sOF6Desc)<>"" OR TRIM(sOF7Desc)<>"" OR TRIM(sOF8Desc)<>"" OR TRIM(sOF9Desc)<>"" OR TRIM(sOF10Desc)<>"" THEN
			%>
			<tr><td align="left" colspan=7><br><b>&nbsp;OTHER OPTIONAL OR REQUIRED ITEMS/FEES</b></td><tr>
			<tr>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td>&nbsp;</td>
				<td align=center><FONT COlOR="<%=TextColor1%>" size=<%=fontsize2%>><b>Required</b></font></td>
				<td align=left><FONT COlOR="<%=TextColor1%>" size=<%=fontsize2%>><b>Amount</b></font></td>
			</tr>
			<%

			
			' --- Loop thru the Optional Fees and display the items ---
			FOR OFItem=0 TO 9
					IF TRIM(OFDescArray(OFItem))<>"" THEN 
							%> 	  
							<tr>
								<td>&nbsp;</td>
								<td ALIGN="right" Colspan=4><FONT COlOR="<%=TextColor1%>" size=<%=fontsize2%>><%=OFDescArray(OFItem)%> &nbsp;</FONT></TD>
								<td ALIGN="center"><FONT COlOR="<%=TextColor2%>" size=<%=fontsize2%>><% IF OFRequiredArray(OFItem)=true THEN response.write("Yes") ELSE response.write("No") END IF %> </FONT></TD>
								<td ALIGN="left"><FONT COlOR="<%=TextColor2%>" size=<%=fontsize2%>> &nbsp;<%=FormatCurrency(OFAmtArray(OFItem),2)%></FONT></TD>
							</tr>
							<%
					END IF
			NEXT

	END IF

%>
	<tr>
		<td colspan=7>&nbsp;</td>
	</tr>
</table>
<%

' --- Begin form portion of the TOURNAMENT page ---
%>
<br>
<table align="center" width=100%>
	<tr>
		<form name="TournamentForm" method="post" action="/rankings/<%=RegFileName%>" id="TournamentForm">
	  	<input type="hidden" name="nav" value=1>
	  	<td width=20% align="center">
				<input type="submit" name="MainStatus" value="Continue" style="width:9em" title="Continue" <%=MainButtonStatus%>>
	  	</td>
	</form><%

	IF adminmenulevel>=20 OR LCASE(Session("UserAdminPW"))=LCASE(Session("AdminCode")) THEN 
			%>
			<form name="TournamentForm3" method="post" action="/rankings/<%=RegFileName%>" id="TournamentForm3">
		  	<td width=20% align="center">
					<input type="submit" name="SkipToPayment" value="Payment Page" style="<%=AdminButtonStyle%>" title="Press to Skip Forward to Payment Page" <%=MainButtonStatus%>>
		  	</td>
		  	<input type="hidden" name="sTourID" value="<%=sTourID%>">
		  	<input type="hidden" name="sMemberID" value="<%=sMemberID%>">
		  	<input type="hidden" name="nav" value="6">
			</form>
			<%
	ELSE 
			%><td width=20% align="center">&nbsp;</td><%
	END IF 

	IF adminmenulevel>=20 OR LCASE(Session("UserAdminPW"))=LCASE(Session("AdminCode")) THEN %>
	      	<form name="TournamentForm4" method="post" action="/rankings/view-Registration.asp?sTourID=<%=sTourID%>" id="TournamentForm4" target="_blank">
		  <td width=20% align="center">
			<input type="submit" name="Registration Report" value="Registrar Report" style="<%=AdminButtonStyle%>" title="Open Registrars reports function for this tournament." >
		  </td>
		</form><%
	ELSE %>
	  <td width=20% align="center">&nbsp;</td><%
	END IF 

	'IF adminmenulevel>=20 OR LCASE(Session("UserAdminPW"))=LCASE(Session("AdminCode")) THEN %>
	      	<form name="TournamentForm2" method="post" action="/rankings/<%=RegFileName%>?sRunByWhat=Tour" id="TournamentForm2">
		  <td width=20% align="center"><input type="submit" name="NewTour" value="New Tournament" style="<%=AdminButtonStyle%>" title="Admin Users can select a new tournament offering online registration.  Note: You will be required to enter the AdminCode for the new tournament." <%=MainButtonStatus%>></td>
		</form><%
	'ELSE %>
	<%'response.write("  <td width=20% align="center">&nbsp;</td> ")<%
	'END IF 

	IF adminmenulevel>=20 OR LCASE(Session("UserAdminPW"))=LCASE(Session("AdminCode")) THEN %>
	      	<form name="TournamentForm5" method="post" action="http://usawaterski.org/rankings/news/FAQ_Register1.htm" id="TournamentForm5" target="_blank">
		  <td width=20% align="center"><input type="submit" name="FAQ" value="Registrar FAQ" style="<%=AdminButtonStyle%>" title="Admin Users can view FAQ."></td>
		</form><%
	ELSE %>
	  <td width=20% align="center">&nbsp;</td><%
	END IF %>

	</tr>
      </table>

      <br>

    </div>
  </div><%



	' **********************************************************************
	' **********************************************************************
	' **********************************************************************
	' -----------------------   MEMBER INFORMATION   -----------------------
	' **********************************************************************
	' **********************************************************************
	' **********************************************************************	

%>
<div class="<% IF nav=2 THEN response.write("accordionHeaderSelected") ELSE response.write("accordionHeader") END IF %>">
	<TABLE>
		<TR>
			<td width="150px" align="left">
				<a href="/rankings/<%=RegFileName%>?nav=2" title="MemberID: <%=sMemberID%>">STEP 2 - Member</a>
			</td>
			<td align="left">
				<font style="color:Yellow;"><%=sFullName%></font>
				<%
				IF sBioDone="N" THEN %>
						&nbsp;&nbsp;&nbsp;&nbsp;
						<font style="color:red;">Bio Requires Update</font><%
				END IF 
				%>
			</td>
		</TR>
	</TABLE>
</div>

  <div class="innertable" style="display:<% IF nav=2 THEN response.write("block") ELSE response.write("none") END IF %>;">
    <div id="RegPanel2">	 
      <br>	

    <table class="spacetable" width=100%>
	<tr>
	  <td align="left" colspan=2><br><b>&nbsp;PERSONAL INFORMATION</b></td>
	  <td align="left" colspan=2><br><b>&nbsp;COMPETITION STATUS</b></td>
	</tr>
	<tr>  
	  <TD width="120px" ALIGN="right"><font size="<%=fontsize2 %>" COlOR="<%=TextColor1%>">Name&nbsp;</FONT></TD>
	  <TD width="250px" ALIGN="left"><FONT size="<%=fontsize2 %>" COlOR="<%=textcolor2%>">&nbsp;<%=sFullName%></FONT></TD>
	  <TD width="120px" ALIGN="right"><font size=<% =fontsize2 %> COlOR="<%=TextColor1%>">Comp Status&nbsp;</FONT></td>
	  <TD width="250px" ALIGN="left"><FONT size="<%=fontsize2%>" COlOR="<%=Session("sMembCanSkiColor")%>">&nbsp;<%=Session("sMembCanSkiText")%></FONT></td>
	</tr>	

	<tr>
	  <TD ALIGN="right"><font size=<% =fontsize2 %> COlOR="<%=TextColor1%>">Member ID&nbsp;</FONT></TD>
	  <TD ALIGN="left"><FONT size=<% =fontsize2 %> COlOR="<%=textcolor2%>"><FONT size=<% =fontsize2 %> COlOR="<%=TextColor2%>">&nbsp;<% =Session("sMemberID") %></FONT></TD>
	  <TD ALIGN="right" vAlign="top"><font size=<% =fontsize2 %>  COlOR="<%=TextColor1%>">Expiration:&nbsp;</FONT></td>
	  <TD ALIGN="left"><FONT size=<% =fontsize2 %> COlOR="<%=Session("sExpirationStatusColor")%>">&nbsp;<%=Session("sExpirationStatusText")%></FONT></td>
	</tr>	

	<tr>
	  <TD ALIGN="right"><font size="<%=fontsize2%>" COlOR="<%=TextColor1%>">City/ST&nbsp;</FONT></TD>
	  <TD ALIGN="left"><FONT size="<%=fontsize2%>" COlOR="<%=TextColor2%>">&nbsp;<%=sMembCity%>,&nbsp;<%=sMembState%></FONT></TD>
	  <TD ALIGN="right" vAlign="top"><font size="<%=fontsize2%>" COlOR="<%=TextColor1%>">Personal Bio &nbsp;</FONT></td>
	  <TD ALIGN="left"><FONT size="<%=fontsize2%>" COlOR="<%=Session("sBioDoneTextColor")%>">&nbsp;<%=Session("sBioDoneText")%></FONT></TD>
	</tr>	

	<tr>
	  <TD ALIGN="right"><font size="<%=fontsize2%>" COlOR="<%=TextColor1%>">Age/Gender&nbsp;</FONT></TD>
	  <TD ALIGN="left"><FONT size="<%=fontsize2%>" COlOR="<%=TextColor2%>"><a TITLE="Correct Age and Gender is required. Please make sure this information is correct.  Contact USA Water Ski Membership Department at 800-533-2972 to correct any missing or inaccurate information.">&nbsp;<%=sMembAge%>/<%=sMembSex%></a></FONT></TD><%
	  
	  ' Session("sMembAge")	  
	  ' -------------------------------------------------------
	  ' ----------- Team Selection (collegiate)  --------------
	  ' ------------------------------------------------------- 

	  %>
	  <TD ALIGN="right"><font size="<%=fontsize2%>" COlOR="<%=TextColor1%>">Team</FONT>&nbsp;</td>
	  <TD ALIGN="left"><%
		IF Session("SptsGrpID")="NCW" THEN 
			TeamStatus="disabled"
			' --- LoadTeam in Tools_Definitions.asp ---
			TeamSelected="ROL"
			'LoadTeam TeamSelected, TeamStatus  
		ELSE
			%><FONT size="<%=fontsize2%>" COlOR="<%=TextColor2%>">&nbsp;N/A</FONT><%
		END IF  %>
	  </TD>
	</tr>
	<tr>
	  <TD ALIGN="right"><font size="<%=fontsize2%>" COlOR="<%=TextColor1%>">Phone&nbsp;</FONT></TD>
	  <TD ALIGN="left"><FONT size="<%=fontsize2%>" COlOR="<%=TextColor2%>">&nbsp;<%=sMembPhone%></FONT></TD>
	  <td>&nbsp;</td>
	  <td>&nbsp;</td>
	</tr>
	<tr>
	  <TD ALIGN="right"><font size="<%=fontsize2%>" COlOR="<%=TextColor1%>">Email&nbsp;</FONT></TD>
	  <TD ALIGN="left"><FONT size="<%=fontsize2%>" COlOR="<%=TextColor2%>">&nbsp;<a href="mailto:<%=sMembEmail%>?subject=Registration issue for <%=sTourName%> - Member <%=sFullName%>"><%=sMembEmail%></a></FONT></TD>
	  <td>&nbsp;</td>
	  <td>&nbsp;</td>
	</tr><%

	IF (LEFT(Session("sMembCanSkiText"),2)<>"OK" OR LEFT(Session("sExpirationStatusText"),2)<>"OK") THEN
			'response.write("<br>sMemberID="&sMemberID)
			'response.write("<br>Session(sMemberID)="&Session("sMemberID"))
			'response.write(AdminmenuLevel=50)
			IF AdminmenuLevel<50 AND TRIM(sMemberID)<>"000001151" THEN  
						MainButtonStatus="disabled"
			ELSE
						MainButtonStatus="enabled"
			END IF			
			%>
			<tr><td colspan=4>&nbsp;</td></tr>
			<tr>
		  	<td ALIGN="center" colspan=4><font size="3" color="red"><b>IMPORTANT NOTICE:</b></FONT></td>
			</tr>
			<tr>
		  	<td ALIGN="center" colspan=4><%
					Dim UpgradeText
					IF LEFT(Session("sMembCanSkiText"),2)<>"OK" THEN
			   			UpgradeText="Upgrade"  %>
			   			<FONT size="<%=fontsize3%>" COlOR="<%=TextColor1%>">Your present <b>Membership Type</b> does not permit participation in sanctioned events of USA Water Ski.  To upgrade your membership to a 'Competition Status', press the <b>Upgrade Membership</b> button below.  <br><br>Once you have completed your membership upgrade, return to this form and press <b>Verify Upgrade</b> to activate your membership in this registration form.</FONT><%
					ELSEIF LEFT(Session("sExpirationStatusText"),2)<>"OK" THEN 
			   			UpgradeText="Renewal"  %>
			   			<FONT size="<%=fontsize3%>" COlOR="<%=TextColor1%>">Your Membership has <b><u>expired</u></b> and must be renewed before you can participate in this tournament.  To renew your membership, press the <b>Membership Renewal</b> button below.  <br><br>Once you have completed your membership renewal, return to this page by restarting the registration process.  If you are entering the same day that you upgraded your membership, then Press <b>Verify Renewal</b> to continue using the online registration form.</FONT><%
					END IF%>			  
		  	</td>
			</tr>
			<tr><td colspan=4>&nbsp;</td></tr>
			<tr>
	    
	    <form name="UpgradeForm" method="post" action="https://www.usawaterski.org/renew/" id="UpgradeForm" target="_blank">
		  	<td width=50% align="center" colspan=2><input type="submit" name="Upgrade" value="Membership <%=UpgradeText%>" style="width:15em" title="Upgrade or Renew your membership status"></td>
			</form>
		
    	<form name="VerifyUpgradeForm" method="post" action="/rankings/<%=RegFileName%>?sRunByWhat=VerifyUpgrade&nav=2" id="VerifyUpgradeForm">
		  	<td width=50% align="center" colspan=2>
		  		<input type="submit" name="Verify Renewal" value="Verify <%=UpgradeText%>" style="width:15em" title="Use this link to cause the entry form to verify that your renewal or upgrade is complete. ">
		  	</td>
			</form>
		
			</tr>
			<tr><td colspan=4>&nbsp;</td></tr>
			<%
	
	END IF  

ep=1
IF ep=2 AND sMemberID="000001151" THEN
	  Response.write("<br>action=/rankings/"&RegFileName&"?sRunByWhat=VerifyUpgrade&nav=2")
END IF
	%>
  </TABLE>
  
  <br>
  <%
  IF sBioDone="N" THEN 
  		'response.write("<br>sBioDone = "&sBioDone)
			'IF AdminmenuLevel<50 AND sMemberID<>"000001151" THEN  
			IF AdminmenuLevel<50 THEN  
						MainButtonStatus="disabled"
			ELSE
						MainButtonStatus="enabled"
			END IF		
			
			'BioButtonStyle = "width:9em; background-color:red; color:white"
			BioButtonStyle = "width:9em; background-color:yellow; color:black"
	ELSE
			BioButtonStyle="width:9em"
	END IF
  
  'response.write("MBS="&MainButtonStatus)
  %>
  
  <TABLE align="center" width=100%>
		<tr>
     	<form name="TournamentForm" method="post" action="/rankings/<%=RegFileName%>" id="TournamentForm">
		  		<input type="hidden" name="nav" value=2>
	  		<td width=25% align="center">
	  			<input type="submit" name="MainStatus" value="Continue" style="width:9em" title="Continue" <%=MainButtonStatus%>>
	  		</td>
	  		<td width=25% align="center">
	  			<input type="submit" name="Previous" value="Previous" style="width:9em" title="Previous Page" <%=PreviousButtonStatus%>>
	  		</td>	
			</form>
				<%


  ' ------------------------------------
  ' --- Button to select UPDATE BIO  --- 
	' ------------------------------------	%>

	<form nowrap action="/rankings/bio-form.asp?FormStatus=new" method="post" target="_blank">
	   <td width=25% align="center" >
				<input type="submit" style="width:9em" style="<%=BioButtonStyle%>" color " <%=BioButtonStatus%> value="Update Bio" title="Create or Update your Personal Bio. &#13; Bio is used for all tournaments. &#13; Keep your bio up-to-date so announcers have current information.">
	   </td>
	</form>
	<%

		IF sBioDone="N" THEN
				%>
	  		<form name="VerifyBioUpdate" method="post" action="/rankings/<%=RegFileName%>?nav=2" id="VerifyBioUpdate">
		 			<td width=25% align="center" colspan=2>
		 				<input type="submit" name="Verify Bio Update" value="Verify Bio Update" style="width:11em" title="Press this button after updating personal bio to confirm the update is complete. ">
		 			</td>
				</form>
				<%
		END IF


  ' ------------------------------------
  ' --- Button to select NEW MEMBER  --- 
	' ------------------------------------	

	' --- Only displays this row if Adminmenulevel correct ---- 
	IF adminmenulevel>=20 OR TestValidAdminCode=true THEN  %>
	   	<form action="/rankings/<%=RegFileName%>?rid=<%=rid%>&sRunByWhat=NewMember" method="post">
				<td  width=25% align="center">  <%
					' --- Button to select NEW MEMBER  --- %>
		  	  <input type="submit" style="<%=AdminButtonStyle%>" value="New Member" title="Admin users may select a new member">
				</td>
			</form><%
	ELSE 
			%>
			<td width=25% align="center">&nbsp;</td>
			<%
	END IF  
		
		%>	
		</tr>
  </table>

  <br>

    </div>
  </div><%





' **********************************************************************
' **********************************************************************
' **********************************************************************
' ---------------------  BEGIN ENTRY FORM DISPLAY ----------------------
' **********************************************************************
' **********************************************************************
' **********************************************************************


  %>
  <div class="<% IF nav=3 THEN response.write("accordionHeaderSelected") ELSE response.write("accordionHeader") END IF %>">
	<TABLE>
		<TR>
		<%
    IF LEFT(Session("sMembCanSkiText"),2)="OK" AND LEFT(Session("sExpirationStatusText"),2)="OK" AND sBioDone="Y" THEN 	
	   		%>
				<td width="150px" align="left">
		  		<a href="/rankings/<%=RegFileName%>?MainStatus=Continue&nav=2"><b>STEP 3 - Entry Form</b></a>
	  		</td>
	  		<%
    ELSE  
    		%>
				<td align=left ><b>STEP 3 - Entry Form </b></td>
	  		<%
    END IF 
	  %>	
		</TR>
	</TABLE>
  </div>

  <div class="innertable" style="display:<% IF nav=3 THEN response.write("block") ELSE response.write("none") END IF %>;">
		<div id="RegPanel3">	 

    <table class="spacetable" width=100%>

      <form name="EntryForm" method="post" action="/rankings/<%=RegFileName%>" id="EntryForm">
			<input type="hidden" name="nav" value=3><%

			' --- Determines column width which varies depending on 


			sClassCols=cdbl(0)
			sClassWidth=70

			' --- Sets hidden variables for those inputs not in this form ---
			SetHiddenFinancialVariables %>

			<tr>
				<td align="left" colspan=8><br><b>&nbsp;ENTRY INFORMATION</b></td>
			</tr> 
			<tr>
	  		<td>&nbsp;</td>
	  		<td>&nbsp;</td>
	  		<td>&nbsp;</td><%

	  		IF sTPandC=true THEN %>	  
		  			<td align="center" width=110px><font size="<%=fontsize2%>"><b>&nbsp;<a title="Select the number of rounds you wish to ski in each event">PICK & CHOOSE</a></b></font></td><%
	  		ELSE %>	
	  				<td width=110px>&nbsp;</td><%
	  		END IF  %>

	  		<td>&nbsp;</td>
 	  		<td align="center" colspan="3" width="210px"><font size="<%=fontsize2%>"><b>&nbsp;ENTRY CLASSIFICATION<br></b></font></td>
			</tr>

			<tr> 
	  		<td align="left" width=120px><font size="<%=fontsize2%>"><b>&nbsp;EVENT</b></font></td>
	  		<td align="center" width=100px><font size="<%=fontsize2%>"><b>&nbsp;ENTER</b></font></td>
	  		<td align="left" width=180px><font size="<%=fontsize2%>"><b>&nbsp;&nbsp;DIVISION </b></font></td><%
		
				IF sTPandC=true THEN %>	  
						<td align="center"><font size="<%=fontsize2%>"><b>&nbsp;# OF ROUNDS</b></font></td><%
				ELSE %>
						<td>&nbsp;</td><%
				END IF  

				IF sShowSkills=true THEN %>	  
						<td align="left"><font size="<%=fontsize2%>"><b>&nbsp;SKILL</b></font></td><%
				ELSE %>
						<td>&nbsp;</td><%
				END IF  

				'ShowStdHead=true
				'ShowRecHead=true
				IF ShowGRHead=true THEN %>
						<td align="center" width="<%=sClassWidth%>px"><font size="<%=fontsize2%>"><b>&nbsp;<%=sTGRClassText%></b></font></td><%
				ELSE %>
						<td width="<%=sClassWidth%>px">&nbsp;</td><%
				END IF 

				IF ShowStdHead=true THEN %>
						<td align="center" width="<%=sClassWidth%>px"><font size="<%=fontsize2%>" title="Premier can indicate classes C,E,L or R (or equivalent) as may be offered by the LOC.  If offered, record classes are typically presented as a separate selection and may be subject to a higher fee structure.">&nbsp;<b><%=sTBaseClassText%></b></font></td><%
				ELSE %>
						<td width="<%=sClassWidth%>px">&nbsp;</td><%
				END IF 

				IF ShowRecHead=true THEN %>
						<td align="center" width="<%=sClassWidth%>px"><font size="<%=fontsize2%>">&nbsp;<b><%=sTUpgradeClassText%></b></font></td><%
				ELSE %>
						<td width="<%=sClassWidth%>px">&nbsp;</td><%
				END IF %>

      </tr>
      <%


' --------------------------------------------------------------------------------------------------------
' --- JAVASCRIPT as patch for GR only being offered as traditional format even when tournament is PandC
' --------------------------------------------------------------------------------------------------------

%>
<script type="text/javascript">
	
	function ResetRoundsToOne(EvtNo)
		{
			alert('Number of pulls not applicable in Grassroots events.  Fees are based on number of events entered.  Setting value to 1')
			if (document.getElementById('GRRadioButton').value=="G")
				{
    			var PulldownName = "fFeeRounds" + EvtNo;
     			document.getElementById(PulldownName).value=1;
     		}	
		}
</script>
<%	






  ' --------------------------------------------------------------------------------------------
	' ------------  Displays checkbox OPTION TourGenTable shows data in Event1 field  ------------
  ' -------------------------------------------------------------------------------------------- 

	Dim fSelectEvent, fDiv, fFeeClass, fFeeRounds, fQfyOverride, fBoat
	Dim TrickEvtNo, JumpEvtNo
	TrickEvtNo=0
	JumpEvtNo=0


	FOR EvtNo = 1 TO TotEv

		  IF TRIM(sTEvent(EvtNo))<>"" THEN  %>
			<tr>
			<td align="left"><font size=<% =fontsize2 %> >&nbsp;<%= sTEventName(EvtNo) %></td>
			<td align="center"><%
			  fSelName = "fSelectEvent"&EvtNo %>
			  <input type=checkbox name="<%=fSelName%>" <% IF sSelectEvent(EvtNo) <> "" THEN Response.Write("Checked "&AllObjectStatus) ELSE Response.write(AllObjectStatus)%>>
			  <font size=<% =fontsize2 %>>&nbsp; Enter</font></td>
			<td align="left">&nbsp;<% 
					' --- SUB in tools_include.asp ---
					LoadDivDropWithAgeGender sDiv(EvtNo), sTEvent(EvtNo), "fDiv"&EvtNo, AllObjectStatus %>
			</td><%

			' ----------------------------------------------
			' --- Displays PICK & CHOOSE rounds dropdown ---	
			' ----------------------------------------------
			fFeeRounds="fFeeRounds"&EvtNo			
			IF sTPandC=true THEN %>
		 		<td align=center><%
					' --- SUB in Tools_Definitions.asp --- 
					SELECT CASE sTEvent(EvtNo)
					  CASE "S"
								ThisMax=sMaxSLPulls
					  CASE "T"
								ThisMax=sMaxTRPulls
					  CASE "J"
								ThisMax=sMaxJUPulls
					END SELECT
					LoadRoundSkiedPulldown fFeeRounds, sFeeRounds(EvtNo), 0, ThisMax, 1, AllObjectStatus, "false" %>
				</td><%
			ELSE %>
				<input type="hidden" name="<%=fFeeRounds%>" value="<%=sFeeRounds(EvtNo)%>">
				<td>&nbsp;</td><%
			END IF  	        

'IF sMemberID="000001151" THEN
'		response.write("<br>sTEvent(EvtNo)= " &sTEvent(EvtNo))
'		response.write("<br>ThisMax= " &ThisMax)
'END IF

			' -------------------------------
			' --- Displays SKILL dropdown ---
			' -------------------------------
			fSkill="fSkill"&EvtNo
			IF sShowSkills=true AND sTEvent(EvtNo)="WB" THEN %>
		 		<td align=center><%
					' --- SUB in Tools_Definitions.asp --- 
					LoadGRSkillPulldown fSkill, sSkill(EvtNo), AllObjectStatus %>
				</td><%
			ELSE %>
				<input type="hidden" name="<%=fSkill%>" value="<%=sSkill(EvtNo)%>">
				<td>&nbsp;</td><%
			END IF  



			' --------------------------------------------------------------------------------
			' --- Using RIGHT function because system sets Barefoot Tricks eventcode to BT
			' --------------------------------------------------------------------------------

			IF RIGHT(TRIM(sTEvent(EvtNo)),1)="T" THEN TrickEvtNo=EvtNo
			IF TRIM(sTEvent(EvtNo))="J" THEN JumpEvtNo=EvtNo


			' ----------------------------------------------------------------------------
			' --- Control over display of radio buttons to select class of competition ---
			' ----------------------------------------------------------------------------

'IF Session("AdminMenuLevel")>49 THEN
'	IF LEFT(sTourID,6)="13M081" THEN
'			response.write("<br>Line 964 RegFormDisplay - sShowGR("&EvtNo&") = "&sShowGR(EvtNo))
'	END IF
'END IF	



			' -----------------------------------------------
			' --- Show radio button for GRASSROOTS events ---
			' -----------------------------------------------

'response.write("<br>sFeeClass(EvtNo) = "&sFeeClass(EvtNo))
'response.write("<br>sShowGR(EvtNo) = "&sShowGR(EvtNo))
			
			'IF sTPandC=true THEN
			'		WhatOnChange="Javascript:ResetRoundsToOne("&EvtNo&");"
			'ELSE
					WhatOnChange=""
			'END IF
			fFeeClass = "fFeeClass"&EvtNo

			' --- Changed 7-13-2013 ---
			IF sShowGR(EvtNo)=true THEN 
					IF sFeeClass(EvtNo)="G" THEN
							StatusToWrite="checked"
					END IF
					%>
					<td align="center" width="<%=sClassWidth%>px">
					<input type=radio id="GRRadioButton" NAME="<%=fFeeClass%>" VALUE="G" <% IF sFeeClass(EvtNo)="G" THEN response.write(StatusToWrite&" "&AllObjectStatus) %> onclick="<%=WhatOnChange%>">
					</td>
					<%
			ELSE %>
				<td>&nbsp;</td><%
			END IF

			' ---------------------------------------------
			' --- Show radio button for BASE/STD events ---
			' ---------------------------------------------


			' --- Changed 7-13-2013 ---
			'IF sShowStd(EvtNo)=true AND TRIM(Session("sEnableStd"))="Y" THEN 
			IF sShowStd(EvtNo)=true THEN 
					IF sFeeClass(EvtNo)="S" THEN
							StatusToWrite="checked"
					END IF
					%>
			   	<td align="center" width="<%=sClassWidth%>px">
						<input type=radio NAME="<%=fFeeClass%>" VALUE="S" title="Premier can indicate classes C,E,L or R (or equivalent) as may be offered by the LOC.  If offered, record classes are typically presented as a separate selection and may be subject to a higher fee structure." <% IF sFeeClass(EvtNo)="S" THEN response.write(StatusToWrite&" "&AllObjectStatus) %>>
					</td><%	
			ELSEIF sShowStd(EvtNo)=true AND TRIM(Session("sEnableStd"))="" THEN %>
			   	<td align="center" width="<%=sClassWidth%>px">
			   		<input type=radio NAME="<%=fFeeClass%>" VALUE="S" disabled>
			   	</td><%	
			ELSE %>
					<td>&nbsp;</td><%
			END IF 

			' -------------------------------------------
			' --- Show radio button for RECORD events ---
			' -------------------------------------------

			' --- Changed 7-13-2013 ---
			'IF sShowRec(EvtNo)=true AND TRIM(Session("sEnableRec"))="Y" THEN 

			IF sShowRec(EvtNo)=true THEN 
					IF sFeeClass(EvtNo)="R" THEN
							StatusToWrite="checked"
					END IF
					%>
					<td align="center" width="<%=sClassWidth%>px">
						<input type=radio NAME="<%=fFeeClass%>" VALUE="R" <% IF sFeeClass(EvtNo)="R" THEN response.write(StatusToWrite&" "&AllObjectStatus) %>>
					</td><%
			ELSEIF sShowRec(EvtNo)=true AND TRIM(Session("sEnableRec"))="" THEN %>
					<td align="center" width="<%=sClassWidth%>px">
						<input type=radio NAME="<%=fFeeClass%>" VALUE="R" disabled>
					</td><%
			ELSE %>
					<td>&nbsp;</td><%
			END IF %>

			</tr><%

		  END IF 

	NEXT 



	' --------------------------------------------------------
	' --- Displays a notice you must select an entry class ---
	' --------------------------------------------------------
	IF (ShowGRHead=true AND ShowStdHead=true) OR (ShowGRHead=true AND ShowRecHead=true) OR (ShowStdHead=true AND ShowRecHead=true) THEN %>
		<tr><td colspan=8>&nbsp;</td></tr>
		<tr>
		  <td colspan=8><font size=<% =fontsize2 %>>&nbsp; If this tournament offers multiple classes. You must select the ENTRY CLASSIFICATION for each event.</font></td>
		</tr><%
	END IF %>
	<tr>
	  <td colspan=8>&nbsp;</td>
	</tr><%


  ' ----------------------------
  ' --- Notice of FORM ERROR ---
  ' ----------------------------
	IF TRIM(sFormError)<>"" THEN %>
		<tr><td colspan=8><font size=<%=fontsize3%> color="red">&nbsp; <b>IMPORTANT:</b> <%=sFormError%></font></td></tr><%
	END IF %>

</table>
<br><%



' -------------------------------------
' --- At least one event is offered ---
' -------------------------------------

IF (TRIM(sTEvent(1)) <> "" OR TRIM(sTEvent(2)) <> "" OR TRIM(sTEvent(3)) <> "" OR TRIM(sTEvent(4)) <> "" OR TRIM(sTEvent(5)) <> "" OR TRIM(sTEvent(6)) <> "") THEN  %>

	<table class="spacetable" width=100%>
	  <tr><td colspan=4>&nbsp;</td></tr>
	  <tr>	
	  <td align="left"><font size="<%=fontsize2%>"><b>&nbsp;EVENT</b></font></td><%

		' --- Headings for adminstrative override ---
	  IF TestValidAdminCode=true OR adminmenulevel>=20 THEN %>
		  <td align="left" width=200px ><font size="<%=fontsize2%>" color="red"><b>&nbsp;ADMINISTRATIVE OVERRIDE</b></font></td><%
	  ELSE %>
		  <td width=200px>&nbsp;</td><%
	  END IF  %>		
	  <td align="center" colspan=2><font size="<%=fontsize2%>"><b>&nbsp;OTHER INFORMATION</b></font></td>
	  </tr><%




		' -------------------------------------
		' --- Displays OVERRIDE information ---
		' -------------------------------------

	  FOR EvtNo = 1 TO TotEv

	    IF TRIM(sTEvent(EvtNo))<>"" THEN

				%>
				<tr>
		  		<td align="left" height=20px><font size=<%=fontsize2%>>&nbsp;<%=sTEventName(EvtNo)%></font></td>
		  		<td align="left">&nbsp;<% 

						' -------------------------------------------
						' --- Displays qualifications information ---
						' -------------------------------------------
						fQfyOverride="fQfyOverride"&EvtNo
						IF TestValidAdminCode=true OR adminmenulevel>=20 THEN
								' --- SUB in tools_Definitions.asp --- 
								LoadQualificationsOverrideDropDown fQfyOverride, sQfyOverride(EvtNo), AllObjectStatus 
						ELSE %> 
								<input type="hidden" name="<%=fQfyOverride%>" value="<%=sQfyOverride(EvtNo)%>"><%
						END IF %>
		  		</td><%


					' -----------------------------------
					' --- Placeholder only for slalom ---
					' -----------------------------------
  				IF TRIM(sTEvent(EvtNo))="S" AND sSelectEvent(EvtNo)="on" THEN %>   
							<td width=120px>&nbsp;</td>
							<td>&nbsp;</td><%
					ELSEIF 	TRIM(sTEvent(EvtNo))="S" AND sSelectEvent(EvtNo)<>"on" THEN %>
							<td>&nbsp;</td>
							<td>&nbsp;</td><%
					END IF

					' ------------------------------------------------------------------------------------------------------------
					' --- Loads TRICK Boat pulldown on NEW ROW or carries variable - LoadBoatPulldown in tools_Definitions.asp ---
					' ------------------------------------------------------------------------------------------------------------
  				IF TRIM(sTEvent(EvtNo))="T" AND sSelectEvent(EvtNo)="on" THEN %>   
							<td align=right width=120px><FONT COlOR="<%=Textcolor1%>" size=<% =fontsize2 %> face=<% =font1 %>>Trick Boat&nbsp;</FONT></td>
							<td>&nbsp;<%
			  			fBoat="fBoat"&TrickEvtNo	
			   			LoadBoatPulldown fBoat, sBoat(TrickEvtNo), AllObjectStatus %>
							</td><%
					ELSEIF TRIM(sTEvent(EvtNo))="T" AND sSelectEvent(EvtNo)<>"on" THEN %>
							<td>&nbsp;</td>
							<td>&nbsp;</td><%
					END IF 

'response.end
'IF sMemberID="000001151" THEN response.write("InRegFormDisplay")
IF sMemberID="000001151" THEN sDiv(JumpEvtNo)="M9"
					' ----------------------------------------------------------------------------------
					' --- Loads JUMP RAMP pulldown - LoadRampPullDownNew is in tools_Definitions.asp ---	
					' ----------------------------------------------------------------------------------
					IF TRIM(sTEvent(EvtNo))="J" AND sSelectEvent(EvtNo)="on" THEN %>
		      		<td align=right width=120px><font color="<%=TextColor1%>" size="<% =fontsize2 %>">Ramp Height&nbsp;</font></td>
		      		<td align=left>
		      			<font color="<% =textcolor2 %>" size="<% =fontsize2 %>">&nbsp;
		      				<% LoadRampPulldownRegister sDiv(JumpEvtNo), "sRampHeight", sRampHeight, AllObjectStatus %>
		      			</font>
		      		<font color="<% =textcolor1 %>" size=<% =fontsize2 %> >-Ft</font></td><%
					ELSEIF TRIM(sTEvent(EvtNo))="J" AND sSelectEvent(EvtNo)<>"on" THEN %>
							<td>&nbsp;</td>
							<td>&nbsp;</td><%
					END IF 


					' --- Accounts for Events NOT S/T/J ---
					IF TRIM(sTEvent(EvtNo))<>"S" AND TRIM(sTEvent(EvtNo))<>"T" AND TRIM(sTEvent(EvtNo))<>"J" THEN %>
							<td>&nbsp;</td>
							<td>&nbsp;</td><%
					END IF 		%>

				</tr><%

	    END IF
  
  	NEXT 

END IF  

		
		' --------------------------
  	' --- Regionals Override ---
		' --------------------------
	  IF TestValidAdminCode=true OR adminmenulevel >= 20 THEN %>
	    <tr>	
	    <td align="left"><FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">&nbsp;Regionals Override</FONT></td>
	    <td>&nbsp;
		<select name="sRegionalOverride" value="<% =sRegionalOverride %>" <% =AllObjectStatus  %>  style="width:12em">
		  <option value ="" <%IF sRegionalOverride = "" THEN Response.Write(" selected ")%> >None</Option><br>
		  <option value ="MED" <%IF sRegionalOverride = "MED" THEN Response.Write(" selected ")%> >Medical Excuse</Option><br>
		  <option value ="OTH" <%IF sRegionalOverride = "OTH" THEN Response.Write(" selected ")%> >Other</Option><br>
		</select>
	    </td>
	    <td colspan=3>&nbsp;</td> 
	    </tr><% 
	  ELSE %>
	    <tr>	
		  <td colspan=4>&nbsp;</td> 
		  <input type="hidden" name="sRegionalOverride" value="<% =sRegionalOverride %>">
	    </tr><%
	  END IF %>
	   

	  <tr><td colspan=4>&nbsp;</td></tr>
	</table>

	<br>
	
	<TABLE align="center" width=100%>
		<tr>
	  	<td width=25% align="center"><input type="submit" name="MainStatus" value="<%=MainStatusValue%>" style="width:9em" title="<%=MainStatusValue%>" <%=MainButtonStatus%>></td>
  </form>

		<form name="EntryForm2" method="post" action="/rankings/<%=RegFileName%>" id="EntryForm2">
	  	<input type="hidden" name="nav" value=3>
	  	<%
	   	MakeHiddenEntryForm	
	   	%>
	  	<td width=25% align="center"><input type="submit" name="Edit" value="Edit" style="width:9em" title="Edit the settings on this page" <%=EditButtonStatus%>></td>
    </form>


		<form name="EntryForm2" method="post" action="/rankings/<%=RegFileName%>?" id="EntryForm2">
	  	<input type="hidden" name="nav" value=1>
	  	<input type="hidden" name="MainStatus" value=Continue>
	  	<%
	  	MakeHiddenEntryForm	
	  	%>
		  <td width=25% align="center"><input type="submit" name="Previous" value="Previous" style="width:9em" title="Back up to previous page" <%=PreviousButtonStatus%>></td>
		</form><%

	IF sQualLevel>0 THEN %>	
			<form name="Qualifications" method="post" action="/rankings/MemberQualifications.asp?sTourID=<%=sTourID%>&sMemberID=<%=sMemberID%>" id="EntryForm2" target="_blank">
		  	<td align="center" width=25%>
					<input type="submit" style="width:9em" value="Qualifications" title="Check your qualification status for this tournament.">
		  	</td>
		  </form><%
	ELSE %>
			<td>&nbsp;</td><%
	END IF	%>

	</tr>
      </table>
      <br>

    </div>
  </div>
<%








' **********************************************************************************************************
' **********************************************************************************************************
' --------------------------------  BEGIN FINANCIAL SECTION  -----------------------------------------------
' **********************************************************************************************************
' **********************************************************************************************************
	%>
  <div class="<% IF nav=4 THEN response.write("accordionHeaderSelected") ELSE response.write("accordionHeader") END IF %>">
  	<%
    IF nav<4 THEN 
    		%>
	  		<TABLE>
	  			<TR>
	  				<td width="25%" align=left>STEP 4 - Financial Summary</td>
	  			</TR>
	  		</TABLE>
	  		<%
    ELSE 
				%>
				<a href="/rankings/<%=RegFileName%>?nav=4">STEP 4 - Financial Summary</a>
				<% 
    END IF 


    %>
  </div>

  <div class="innertable" style="display:<% IF nav=4 THEN response.write("block") ELSE response.write("none") END IF %>;">
    <div id="RegPanel4">	 
      <br>
      <form name="FinancialForm" method="post" action="/rankings/<%=RegFileName%>" id="FinancialForm">	
				<input type="hidden" name="nav" value=4>
				<%
	
				MakeHiddenEntryForm	 

				' --- TEST --- 
				RecalcFormValues

 				' DisplayCurrentValues "Inside Financials of Form" 
 				
 				%>
      	<table class="spacetable" width="100%">
					<tr>
		  			<td align="left" colspan=4><b>&nbsp;GENERAL INFORMATION</b></td>
		  			<td colspan=2 align=right>
				  	<%
				  	'response.write("<br>AllObjectStatus="&AllObjectStatus)
				  	IF (Adminmenulevel>=20 OR TestValidAdminCode=true) AND TRIM(AllObjectStatus)<>"disabled" THEN
				  			%>
		  					<FONT size=<% =fontsize1 %> face=<% =font1 %> COlOR="<% =TextColor1 %>">mm/dd/yyyy hh:mm:ss PM</font>
		  					<%
						END IF
						%>
	 					</td>
		  		</tr>
					
					<tr>
					<%

					IF sTEntryFeeFamily<>cdbl(0) THEN 
							%>
		  				<td align="right" width=100px><FONT COlOR="<%=TextColor1%>" size="<%=fontsize2%>">Entry Type&nbsp;</font></td>
		  				<td align=left width=150px>&nbsp;		
								<select name="sEntryType" value="<%=sEntryType %>" <% =AllObjectStatus %> >
			  					<option value ="IND" <%IF sEntryType = "IND" THEN Response.Write(" selected ")%> >Individual</Option><br>
			  					<option value ="FAM" <%IF sEntryType = "FAM" THEN Response.Write(" selected ")%> >Family Entry</Option><br>
								</select>
		  				</td>
		  				<%
					ELSE 
							%>
		  				<input type="hidden" name="sEntryType" value="<% =sEntryType %>">		  
		  				<td width=100px>&nbsp;</td>
		  				<td width=150px>&nbsp;</td>
		  				<%
					END IF

					IF adminmenulevel>=20 OR TestValidAdminCode=true THEN 
							%>	
			    		<td align=right width=100px><FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =TextColor3 %>">Fee Override&nbsp;</FONT></td>
			    		<td align=left width=150px>&nbsp;
							<select name="sMoneyOverride" title="Test" value="<% =sMoneyOverride %>" <% =AllObjectStatus %>>
				  			<option value ="" <%IF sMoneyOverride = "" THEN Response.Write(" selected ")%> >None</Option><br>
				  			<option value ="OTH" <%IF sMoneyOverride = "OTH" THEN Response.Write(" selected ")%> >Other</Option><br>
				  			<option value ="FAM" <%IF sMoneyOverride = "FAM" THEN Response.Write(" selected ")%> >Family</Option><br>
							</select>
			    		</td>
			    		<%
					ELSE 
							%>
							<td width=100px>&nbsp;</td>
							<td width=150px>&nbsp;</td>
							<input type="hidden" name="sMoneyOverride" value="<%=sMoneyOverride%>" <%=AllObjectStatus%>>
							<%
					END IF

		  		EntryColor=TextColor1		  
		
		  		IF adminmenulevel>=20 OR TestValidAdminCode=true THEN EntryColor="red"
				  IF (Adminmenulevel<20 AND TestValidAdminCode<>true) OR AllObjectStatus="disabled" THEN
							MembRegDateStatus="disabled"  
							%>
							<input type="hidden" name="sMembRegDate" value="<%=sMembRegDate%>" <%=AllObjectStatus%>>
							<%
		  		END IF 
		  		%>
		  		<td align=right width=150px><font size=<% =fontsize2 %> color=<%=EntryColor%>>&nbsp;Date Entered</font></td>
		  		<td align=left width=100px>
		  			<input type="text" name="sMembRegDate" value="<% =sMembRegDate %>" MAXLENGTH=22 size="22" <% =MembRegDateStatus %>>
		  		</td>
				</tr>
				<%


				' --- Notice for Family Entries ---
				IF sEntryType="FAM" THEN

		   			' --- Number of people in this family membership group ---
		   			IF MainStatus<>"Verify" OR TotQualifyingFamMemb>0 THEN 
		   					%>
			  				<tr>
			    				<td colspan=6 align=left>&nbsp;<font color="<% =textcolor2 %>" size=<% =fontsize2 %>><b>IMPORTANT</b></font></td>
			  				</tr> 
			  				<tr>
			    			<td colspan=6>
			    				<%
			      			IF TRIM(Session("sWhichFamilyMemberPaid"))<>"" AND sMaxFamMembers>1 THEN 
			      					%>
											<font color="<% =textcolor1 %>" size=<% =fontsize2 %>>&nbsp;<%=Session("sWhichFamilyMemberPaid")%> was charged for the 'Family Entry Fee'.<br>&nbsp;Late entry fees and other charges are not included in Family Entry Fee.</font>
											<%
			   					ELSEIF TRIM(Session("sWhichFamilyMemberPaid"))<>"" AND sMaxFamMembers=1 THEN 
			   							%>
											<font color="<% =textcolor1 %>" size=<% =fontsize2 %>>&nbsp;<%=Session("sWhichFamilyMemberPaid")%> was charged for the 'Family Entry Fee'. All other entries for family members will be charged the 'Additional Family Member' fee.&nbsp;Late entry fees and other charges are not included in Family Entry Fee.</font><%
			      			ELSEIF TRIM(Session("sWhichFamilyMemberPaid"))="" AND sMaxFamMembers>0 THEN %>
											<font color="<% =textcolor1 %>" size=<% =fontsize2 %>>&nbsp;The first family member registering will be charged the 'Family Entry Fee', which will pay for up to <%=sMaxFamMembers%> entries.&nbsp;  All other entries for family members will be charged the 'Additional Family Member' fee.&nbsp;Late entry fees and other charges are not included in Family Entry Fee.</font><%
			      			ELSE %>
											<font color="<% =textcolor1 %>" size=<% =fontsize2 %>>&nbsp;The first family member registering will be charged the 'Family Entry Fee.'  All other entries for family members are free.<br>&nbsp;Late entry fees and other charges are not included in Family Entry Fee.</font><%
			      			END IF  
			      			%>
			    			</td>
		  				</tr>

			  			<tr>	
			    			<td colspan=6 align=left><br>&nbsp;<font color="<% =textcolor2 %>" size=<% =fontsize2 %>><b>FAMILY MEMBERS INCLUDE - Total <%=TotQualifyingFamMemb%></b></font></td>
			  			</tr>
			  			<%

			  			' --- Displays the list of family members for this member ---
			  			MembNo=0
 			  			DO WHILE MembNo<10 
									MembNo=MembNo+1 
									%>
									<tr>
			    					<td colspan=6 align=left>&nbsp;<font color="<% =textcolor1 %>" size=<% =fontsize1 %>><%=MembListName(MembNo)%></font></td>
									</tr>
									<%
									IF TRIM(MembList(MembNo))="" THEN EXIT DO
			  			LOOP
		   				' --- Not a family membership even though user set it that way 
		   		ELSE  
		   				%>
			  			<tr>
			    			<td colspan=6 align=left><br>&nbsp;
									<font color="<% =textcolor3 %>" size=<% =fontsize4 %>><b>WARNING</b></font>
			      			<br><br>
									<font color="<% =textcolor3 %>" size=<% =fontsize2 %>><b>&nbsp;&nbsp;THIS IS NOT A FAMILY MEMBERSHIP TYPE</b></font>
									<br>
			    			</td>
			  			</tr><%
		   		END IF

		END IF 

		
		%> 		 
		<tr>
		  <td colspan=6>&nbsp;</td>
		</tr> 
		<tr>
		  <td align="left" colspan=4><b>&nbsp;FEES AND CHARGES</b></td>
		  <td>&nbsp;</td>
		  <td>&nbsp;</td>
		</tr>

		<tr>
		  <td colspan=4>&nbsp;</td>
	  	  <td align="right" width=150px><font color="<% =textcolor1 %>" size=<% =fontsize2 %> face=<% =font1 %>><b>ENTRY FEES</b></font></td>
		  <td align="right" width=150px><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>><b><%=FormatCurrency(sEntryFee,2)%>&nbsp;&nbsp;</b></font></td>
		</tr><%

 

		  ' ------------------------------------------------------------	
		  ' ---- Discount to Junior B/G 1-3 per Tour_Manager.asp   -----
		  ' ------------------------------------------------------------


		  IF sJrDiscPerc <> 0 AND sMembAge < 18 AND sEntryFee > 0 THEN %>
			<tr><%	 
			  IF MainStatus="Verify" THEN
				%><td align="left" colspan=4><font size=<% =fontsize2 %> color=<%=TextColor2%>>&nbsp;&nbsp;Press <b>'Edit'</b> to modify information</font></td><%
			  ELSE
				%><td align="left" colspan=4><font color=red size="<%=fontsize2 %>">&nbsp;&nbsp;Check all that apply and press <b>'Verify'</b></font></td><%
			  END IF  %>

			  <td align="right"><font size=<% =fontsize2 %> face=<% =font1 %>>Junior Discount</font></td>
			  <td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>><%= FormatCurrency(sJrDiscAmt,2) %>&nbsp;</font></td>
			</tr><%
		  END IF 	


		  ' -------------------------------------------------------------------------	
		  ' ---- Discount to divisions M/W-6 if specified in Tour_Manager.asp   -----
		  ' -------------------------------------------------------------------------

'response.write("<br>sSrDiscPerc="&sSrDiscPerc)
'response.write("<br>sMembAge="&sMembAge)

		  IF cdbl(sSrDiscPerc) <> 0 AND sMembAge > 59 AND cdbl(sEntryFee) > 0 THEN  %>
			<tr>
			  <td colspan="4">&nbsp</td>
			  <td align="right"><font size=<% =fontsize2 %> face=<% =font1 %>>Senior Discount</font></td>
			  <td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>><%=FormatCurrency(sSrDiscAmt,2)%>&nbsp;</font></td>
			</tr><%
		  END IF


		' -------------------------------------------------------------------------	
		' ---------- Discount to OFFICIALS if specified in Tour_Manager.asp   -----
		' -------------------------------------------------------------------------  

		IF sOffDiscPerc <> cdbl(0) THEN %>
			<tr>	
			  <td colspan="4">
			    <input type=checkbox name="fOfficial" <%IF sOfficial = "on" THEN Response.Write("Checked "&AllObjectStatus) ELSE Response.write(AllObjectStatus) %>>
			    <font size=<% =fontsize2 %> >&nbsp;&nbsp;Check here if you are an invited official.</font>
			  </td><%  
			
			  IF sOfficial = "on" THEN  %>
				<td align="right"><font color="<% =textcolor1 %>" size=<% =fontsize2 %> face=<% =font1 %>>Officials Discount</font></td>
				<td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>><%=FormatCurrency(sOffDiscAmt,2)%>&nbsp;</font></td><%	
			  ELSE %>
				<td>&nbsp</td>
				<td>&nbsp</td><%
			  END IF %>
			</tr><%
		END IF

		  ' -------------------------------------------------------------------------------------------------	
		  ' ---------- Discount to CLUB MEMBERS if match to ClubCode as specified in Tour_Manager.asp   -----
		  ' -------------------------------------------------------------------------------------------------  

		      IF sClubDiscPerc <> cdbl(0) THEN %>
		       	  <tr> 
			     <td colspan="4"><input type=checkbox name="fClubMemb" <%IF sClubMemb = "on" THEN Response.Write("Checked "&AllObjectStatus) ELSE Response.write(AllObjectStatus) %>>
		     	  	<font size=<% =fontsize2 %> face=<% =font1 %>>&nbsp;&nbsp;Check if Member of the Host Club.  CLUB CODE</font>
			  	<input type="text" name="fClubCode" value="<% =sClubCode %>" maxlength=5 size="5" <%=AllObjectStatus%>>
			     </td><%  

	  	    		IF cdbl(sClubDiscPerc) <> 0 AND sClubMemb = "on" AND cdbl(sEntryFee) > 0 THEN
				   IF TRIM(sClubCode) <> "" AND TRIM(sClubCode)=TRIM(sTourClubCode) THEN  %>
					<td align="right"><font size=<% =fontsize2 %> face=<% =font1 %>>Club Member Discount</font></td>
					<td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>><%=FormatCurrency(sClubDiscAmt,2)%>&nbsp;</font></td><%	
				   ELSEIF MainStatus="Verify" AND (TRIM(sClubCode) = "" OR TRIM(sClubCode)<>TRIM(sTourClubCode)) THEN  %>
					  <td align="right"><font color="red" size=<% =fontsize2 %> face=<% =font1 %>> &nbsp;&nbsp;Club Code is Invalid</font></td>
					  <td>&nbsp;</td><%						
				   ELSE
					%><td>&nbsp;</td>
					  <td>&nbsp;</td><%						
				   END IF
		  		ELSE  %>
				   <td>&nbsp</td>
			  	   <td>&nbsp</td><%
				END IF	%>
			  </tr><%
		     END IF  


		IF sAWSEFDon_OK=true THEN
		  ' -------------------------------------------	
		  ' ---- Donation to AWSEF Building Fund  -----
		  ' -------------------------------------------  %>

	      <tr>
		    	<td colspan="4" style="height:2em;">
		    		<input type=checkbox name="fAWSEFCheck" <% IF sAWSEFCheck = "on" THEN Response.Write("Checked "&AllObjectStatus) ELSE Response.write(AllObjectStatus) %>>
		   			<font size=<% =fontsize2 %> face=<% =font1 %>>&nbsp;&nbsp;Check to Donate $10.00 to American Water Ski Educational Foundation</font>
		    	</td>
		    	<%

		  		IF sAWSEFCheck = "on" THEN  
		  				%>
							<td align="right"><font size=<% =fontsize2 %> face=<% =font1 %>>Donation</font></td>
							<td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>><%=FormatCurrency(sAWSEFDonation,2)%>&nbsp;&nbsp;</font></td>
							<%
		  		ELSE	
		  				%>
							<td>&nbsp</td>
							<td>&nbsp</td>
							<%
		  		END IF %>
		</tr><%
		END IF

		' ---------------------------------------------
		' --------  LATE FEES --------------------------
		' ---------------------------------------------  
		IF Cdbl(sLateFeeTot)>Cdbl(0.00) THEN  %>
			<tr>
			 <td colspan=4 style="height:2em;">&nbsp</td>
			 <td align="right"><font size=<% =fontsize2 %> face=<% =font1 %>>Late Fee</font><%
			    IF sTLFPerDay=true THEN %>
				 <font size=<% =fontsize2 %> face=<% =font1 %>>- <%=sLateDays%> Days</font><%
			    END IF  %>
			 </td>
			 <td align="right"><font color="<% = textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>><%=FormatCurrency(sLateFeeTot)%>&nbsp;&nbsp;</font></td>
			</tr><%
		END IF



	  ' ----------------------------
	  ' ---- Banquet Tickets  -----
	  ' ----------------------------

		IF sBTickCost>0 THEN %>
		  <tr>
		    <td colspan=4 style="height:2em;">
			<font color="<%=textcolor1%>" size=<% =fontsize2 %>>&nbsp;&nbsp;&nbsp;&nbsp;Banquet tickets at <%=FormatCurrency(sBTickCost)%>/ticket.</font><%

			'AllObjectStatus ="disabled"
			LoadValuePulldown "sBanquetQty", sBanquetQty, 0, 10, 1, AllObjectStatus, "true"

			IF sBTickWithE=true THEN %>
				<font color="<%=textcolor3%>" size=<% =fontsize2 %>><br>&nbsp;&nbsp;&nbsp;&nbsp;IF ATTENDING THE BANQUET CHOOSE ONE (1) TICKET. THIS TICKET WILL BE<br>&nbsp;&nbsp;&nbsp;&nbsp;FREE AND CHARGES START @ 2 AND UP</font><%
			END IF %>
		    </td>
		    <td align="right"><font color="<%=textcolor1%>" size=<% =fontsize2 %>>Banquet Tickets</font></td>
		    <td align="right"><font color="<%=textcolor2%>" size=<% =fontsize2 %>><%=FormatCurrency(sBanquetTot)%>&nbsp;&nbsp;</font></td>
		  </tr><%
		END IF 
	
	
	
		' ----------------------------
		' --- Optional Items --------- 
		' ----------------------------
		
			FOR OFItem=0 TO 9
					IF TRIM(OFDescArray(OFItem))<>"" THEN
							%>
		  				<tr>
		    				<td colspan=4 style="height:2em;">
									<font color="<%=textcolor1%>" size=<% =fontsize2 %>>&nbsp;&nbsp;&nbsp;&nbsp; <%=OFDescArray(OFItem)%> at <%=FormatCurrency(OFAmtArray(OFItem))%>:</font>
										&nbsp;&nbsp;
		    					<%
									' --- Array starts at zero so +1 required to get correct element NAME --- 
									LoadValuePulldown "sOF"&OFItem+1&"Qty", OFQtyArray(OFItem), 0, OFMaxQtyArray(OFItem), 1, AllObjectStatus, "true" 
									%>
		    				</td>
		    				<td align="right"><font color="<%=textcolor1%>" size=<% =fontsize2 %>>Cost</font></td>
		    				<td align="right"><font color="<%=textcolor2%>" size=<% =fontsize2 %>><%=FormatCurrency(OFFeeArray(OFItem))%>&nbsp;&nbsp;</font></td>
		  				</tr>
		  				<%
					END IF
			NEXT


		' -----------------------------------------
		' --- Totals at the lower right of page ---
		' -----------------------------------------

		%>
		<tr>
		  <td colspan=4>&nbsp</td>
		  <td>&nbsp</td>
		  <td>&nbsp</td>
		</tr>
		<tr>
			<td colspan=4 align="right" style="height:2em;"><font color="<%=TextColor1%>" size=<% =fontsize2 %> face=<% =font1 %>><%=sDiscNote%></font></td>				   
		    	<td align="right"><font color="<%=TextColor1%>" size=<% =fontsize2 %> face=<% =font1 %>><b>TOTAL ALL</b></font></td>
		    	<td align="right"><font color="<% =textcolor2 %>" size=<% =fontsize2 %> face=<% =font1 %>><b><%=FormatCurrency(sTotalFormFees,2)%>&nbsp;&nbsp;</b></font></td>
		  </tr>
		  <tr>
			<td colspan=4 style="height:2em;">&nbsp;</td>
			<td align="right"><font color="<% =textcolor1 %>" size=<% =fontsize2 %> face=<% =font1 %>>Previous Payments</font></td>
			<td align="right"><font color="<%=TextColor1%>" size=<% =fontsize2 %> face=<% =font1 %>><%=FormatCurrency(cdbl(sTotalPreviousPayments),2)%>&nbsp;&nbsp;</font></td>
		  </tr>
		  <tr>
			<td colspan=4>&nbsp;</td><%

'IF Session("AdminMenuLevel")<>"" THEN
'	RESPONSE.WRITE("<br>Session(sWhichFamilyMemberPaid) = "&Session("sWhichFamilyMemberPaid"))
'	RESPONSE.WRITE("<br>maxfam="&sMaxFamMembers)
'END IF
			' --- Session("sWhichFamilyMemberPaid")<>"" When the family fee has been paid by another member
			IF (TRIM(Session("sWhichFamilyMemberPaid"))<>"" AND sMaxFamMembers>Session("TotRegisteredFamMembers")) THEN  %>
				  <td align="right" style="height:2em;"><font color="<% =textcolor3 %>" size="<% =fontsize2 %>" face="<% =font1 %>"><b>FAMILY MEMBER</b></font></td>
				  <td align="right"><font color="<% =textcolor2 %>" size="<% =fontsize2 %>" face="<% =font1 %>"><b>Paid</b></font></td><%
			
			ELSEIF cdbl(sTotalPreviousPayments) <= cdbl(sTotalFormFees) THEN %> 
				  <td align="right" style="height:2em;"><font color="<% =textcolor3 %>" size="<% =fontsize2 %>" face="<% =font1 %>"><b>BALANCE DUE</b></font></td>
				  <td align="right"><font color="<% =textcolor2 %>" size="<% =fontsize2 %>" face="<% =font1 %>"><b><%=(FormatCurrency(cdbl(sTotalFormFees)-cdbl(sTotalPreviousPayments),2))%>&nbsp;&nbsp;</b></font></td><%

			ELSE %>
				  <td align="right" style="height:2em;"><font color="<% =textcolor3 %>" size=<% =fontsize2 %> face=<% =font1 %>><b>CREDIT DUE</b></font></td>
				  <td align="right"><font color="<% =textcolor3 %>" size=<% =fontsize2 %> face=<% =font1 %>><b><%= FormatCurrency(cdbl(sTotalFormFees)-cdbl(sTotalPreviousPayments),2)%>&nbsp;&nbsp;</b></font></td><%
			END IF  %>
		  </tr><%
		IF MainStatusValue="Verify" AND TRIM(sFormError)="" THEN 
				%>
		     <tr>
					<td align="left" colspan=6>
						<br>
						<font color="<% =textcolor3 %>" size=<% =fontsize2 %>>&nbsp;&nbsp;You must press <b>Verify</b> to calculate your total fees and apply any applicable discounts.</font>
						<br>
					</td>
		     </tr>
		     <%
		ELSEIF TRIM(sFormError)<>"" THEN
		  	' --- Notice of FORM ERROR ---
				%>
				<tr>
					<td colspan=8><font size=<%=fontsize3%> color="red">&nbsp; <b>INPUT ERROR:</b> <%=sFormError%></font></td>
				</tr>
				<%
				MainButtonStatus="disabled"
		END IF 
  

		
		%>
		</table>
		<br>

	        <table align="center" width=100%>
		<tr>
		  <td width=33% align="center"><input type="submit" name="MainStatus" value="<%=MainStatusValue%>" style="width:9em" title="<%=MainStatusValue%>" <%=MainButtonStatus%>></td>
		</form>

	        <form name="FinancialForm2" method="post" action="/rankings/<%=RegFileName%>" id="FinancialForm2">
		  <input type="hidden" name="nav" value=4>
		  <td width=33% align="center"><input type="submit" name="Edit" value="Edit" style="width:9em" title="Edit the settings on this page" <%=EditButtonStatus%>></td>
		  <td width=33% align="center"><input type="submit" name="Previous" value="Previous" style="width:9em" title="Previous to previous page" <%=PreviousButtonStatus%>></td>
		</form>
		</tr>
	      </table>
	      <br>	

    </div>
  </div>

<%




	' *****************************************************************************
	' *****************************************************************************
  ' *****************************************************************************
  ' -----------------------------  RELEASE  -------------------------------------
	' *****************************************************************************
	' *****************************************************************************
  ' *****************************************************************************
	%>
  <div class="<% IF nav=5 THEN response.write("accordionHeaderSelected") ELSE response.write("accordionHeader") END IF %>">
 		<TABLE>
 			<TR>
 				<td width="150px" align=left>STEP 5 - Waiver</td>
 				<td align=left>
 					<font style="color:<%=Session("sReleaseTextColor")%>"><%=Session("sReleaseText")%></font>
 				</td>
 			</TR>
 		</TABLE>
  </div>

  <div class="innertable" style="display:<% IF nav=5 THEN response.write("block") ELSE response.write("none") END IF %>;">
    <div id="RegPanel5">
    	<%	 
			AccordRelease 
			%>
    </div>
  </div>
  <%






	' *****************************************************************************
	' *****************************************************************************
  ' *****************************************************************************
  ' -----------------------------  PAYMENT PAGE  --------------------------------
	' *****************************************************************************
	' *****************************************************************************
  ' *****************************************************************************
 

Dim Item_Name(10), Amount(10), Quantity(10), ItemNo, RegFeeAmount, NETAmountDue

FOR ItemNo=1 TO 10
	Amount(ItemNo)=cdbl(0)
NEXT

' --- Get previous charges from Transaction Table ---
ReadFromTransTable

' --- TEST 
RecalcFormValues





' --- Find just the registration fee amount ---
RegFeeAmount=sTotalFormFees-cdbl(sTotalPreviousPayments)-sBanquetTot-sAWSEFDonation+sBanquetTotTrans+sAWSEFDonationTrans-sOF1Fee+sOF1FeeTrans-sOF2Fee+sOF2FeeTrans-sOF3Fee+sOF3FeeTrans-sOF4Fee+sOF4FeeTrans-sOF5Fee+sOF5FeeTrans-sOF6Fee+sOF6FeeTrans-sOF7Fee+sOF7FeeTrans-sOF8Fee+sOF8FeeTrans-sOF9Fee+sOF9FeeTrans-sOF10Fee+sOF10FeeTrans
sAllDisc=sSrDiscAmt+sJrDiscAmt+sOffDiscAmt+sClubDiscAmt

'response.write("<br>HERE")
'response.write("<br>sTotalPreviousPayments="&sTotalPreviousPayments)
'response.write("<br>sOF1Fee="&sOF1Fee)
'response.write("<br>sTotalFormFees="&sTotalFormFees)
'response.write("<br>RegFeeAmount="&RegFeeAmount)
'response.end

ty=2
IF ty=2 AND sMemberID="000001151" THEN
	RegFeeAmount=cdbl(1)
	sAWSEFDonation=cdbl(0.23)
END IF


ItemNo=0
IF RegFeeAmount<0 OR sBanquetTot-sBanquetTotTrans<0 OR sAWSEFDonation-sAWSEFDonationTrans<0 OR sOF1Fee-sOF1FeeTrans<0 OR sOF2Fee-sOF2FeeTrans<0 OR sOF3Fee-sOF3FeeTrans<0 OR sOF4Fee-sOF4FeeTrans<0 OR sOF5Fee-sOF5FeeTrans<0 OR sOF6Fee-sOF6FeeTrans<0 OR sOF7Fee-sOF7FeeTrans<0 OR sOF8Fee-sOF8FeeTrans<0 OR sOF9Fee-sOF9FeeTrans<0 OR sOF10Fee-sOF10FeeTrans<0 THEN
	ItemNo=ItemNo+1
	Item_Name(ItemNo)="Changes to Registration for Member # "&sMemberID&" at "&sTourName
	Quantity(ItemNo)="1"
	Amount(ItemNo)=sTotalFormFees-cdbl(sTotalPreviousPayments)

' --- TEST




ELSE
	' --- If the total fees less banquet tickets and AWSEF are greater than zero ---
	IF RegFeeAmount>0 THEN
		ItemNo=ItemNo+1
		Item_name(ItemNo)="Registration for Member # "&sMemberID&" at "&sTourName
		Quantity(ItemNo)="1"
		Amount(ItemNo)=cdbl(RegFeeAmount)	
	END IF


	' --- If the Banquet $$ is greater than zero and ticket payments are less than Banquet $$ --- 
	IF sBanquetTot>0 AND sBanquetTotTrans<sBanquetTot THEN
		ItemNo=ItemNo+1
		IF sBTickWithE=true THEN 
			Item_Name(ItemNo)="Banquet Ticket(s) - NOTE: 1 extra ticket is included with entry"
		ELSE
			Item_name(ItemNo)="Banquet Ticket(s)"
		END IF	
		Quantity(ItemNo)=(sBanquetTot-sBanquetTotTrans)/sBTickCost
		Amount(ItemNo)=cdbl(sBTickCost)
	END IF


	' --- If AWSEF donation ---
	IF sAWSEFDonation>0 AND sAWSEFDonation<>sAWSEFDonationTrans THEN 
		ItemNo=ItemNo+1
		Item_Name(ItemNo)="AWSEF Donation"
		Quantity(ItemNo)="1"
		Amount(ItemNo)=cdbl(sAWSEFDonation-sAWSEFDonationTrans)
	END IF

NewOptDisplay="N"
IF NewOptDisplay="Y" THEN
	FOR OFItem=0 TO 9
			IF TRIM(OFDescArray(OFItem))<>"" AND OFQtyArray(OFItem)>0 THEN
					ItemNo=ItemNo+1
					Item_Name(ItemNo)=OFDescArray(OFItem)
					Quantity(ItemNo)=OFQtyArray(OFItem)
					Amount(ItemNo)=cdbl(OFFeeArray(OFItem)/OFQtyArray(OFItem))
			END IF
	NEXT

ELSE

	' --- These define OPTIONAL CUSTOM ITEMS ---
	IF TRIM(sOF1Desc)<>"" AND sOF1Qty>0  THEN
		ItemNo=ItemNo+1
		Item_Name(ItemNo)=sOF1Desc
		Quantity(ItemNo)=sOF1Qty
		Amount(ItemNo)=cdbl(sOF1Fee/sOF1Qty)
	END IF

	IF TRIM(sOF2Desc)<>"" AND sOF2Qty>0 THEN
		ItemNo=ItemNo+1
		Item_Name(ItemNo)=sOF2Desc
		Quantity(ItemNo)=sOF2Qty
		Amount(ItemNo)=cdbl(sOF2Fee/sOF2Qty)
	END IF

	IF TRIM(sOF3Desc)<>"" AND sOF3Qty>0 THEN
		ItemNo=ItemNo+1
		Item_Name(ItemNo)=sOF3Desc
		Quantity(ItemNo)=sOF3Qty
		Amount(ItemNo)=cdbl(sOF3Fee/sOF3Qty)
	END IF

	IF TRIM(sOF4Desc)<>"" AND sOF4Qty>0  THEN
		ItemNo=ItemNo+1
		Item_Name(ItemNo)=sOF4Desc
		Quantity(ItemNo)=sOF4Qty
		Amount(ItemNo)=cdbl(sOF4Fee/sOF4Qty)
	END IF

	IF TRIM(sOF5Desc)<>"" AND sOF5Qty>0  THEN
		ItemNo=ItemNo+1
		Item_Name(ItemNo)=sOF5Desc
		Quantity(ItemNo)=sOF5Qty
		Amount(ItemNo)=cdbl(sOF5Fee/sOF5Qty)
	END IF

	IF TRIM(sOF6Desc)<>"" AND sOF6Qty>0  THEN
		ItemNo=ItemNo+1
		Item_Name(ItemNo)=sOF6Desc
		Quantity(ItemNo)=sOF6Qty
		Amount(ItemNo)=cdbl(sOF6Fee/sOF6Qty)
	END IF

	IF TRIM(sOF7Desc)<>"" AND sOF7Qty>0 THEN
		ItemNo=ItemNo+1
		Item_Name(ItemNo)=sOF7Desc
		Quantity(ItemNo)=sOF7Qty
		Amount(ItemNo)=cdbl(sOF7Fee/sOF7Qty)
	END IF

	IF TRIM(sOF8Desc)<>"" AND sOF8Qty>0  THEN
		ItemNo=ItemNo+1
		Item_Name(ItemNo)=sOF8Desc
		Quantity(ItemNo)=sOF8Qty
		Amount(ItemNo)=cdbl(sOF8Fee/sOF8Qty)
	END IF

	IF TRIM(sOF9Desc)<>"" AND sOF9Qty>0  THEN
		ItemNo=ItemNo+1
		Item_Name(ItemNo)=sOF9Desc
		Quantity(ItemNo)=sOF9Qty
		Amount(ItemNo)=cdbl(sOF9Fee/sOF9Qty)
	END IF

	IF TRIM(sOF10Desc)<>"" AND sOF10Qty>0  THEN
		ItemNo=ItemNo+1
		Item_Name(ItemNo)=sOF10Desc
		Quantity(ItemNo)=sOF10Qty
		Amount(ItemNo)=cdbl(sOF10Fee/sOF10Qty)
	END IF
END IF




END IF


ReturnURL="http://usawaterski.org/rankings/"&RegFileName&"?sTourID="&sTourID&"&sMemberID="&sMemberID&"&sOrderNo="&sOrderNo&"&nav=7&sPayType=PayPal"
ReturnURLBad="http://usawaterski.org/rankings/"&RegFileName&"?sTourID="&sTourID&"&sMemberID="&sMemberID&"&sOrderNo="&sOrderNo&"&nav=6&sPayType=PPErr"

simage_url="https://www.usawaterski.org/rankings/images/logos/usawslogo_no_sub.jpg"



%>

  <div class="<% IF nav=6 THEN response.write("accordionHeaderSelected") ELSE response.write("accordionHeader") END IF %>">
 		<TABLE>
 			<TR>
 				<td width="150px" align=left>STEP 6 - Payment</td>
 				<td align=left>
 					<font style="color:<%=Session("FeeStatusTextColor")%>;"><%=Session("FeeStatusText")%></font>
 				</td>
 			</TR>
 		</TABLE>
	</div>

  <div class="innertable" style="display:<% IF nav=6 THEN response.write("block") ELSE response.write("none") END IF %>">
    <div id="RegPanel6">

	<br>
	<TABLE Align=center width="90%">
	  <TR>
	   	<TD Align="Left"><font size="3"><b>Review Your Order</b></font></TD>
	  </TR>
	  <TR>
	  <%
		' --- CODE FOR TESTING New REFUND POLICY - 3/26/2015 
		'IF adminmenulevel>49 AND LEFT(sTourID,6)="15S119" THEN
		'		sHQAccount=true
		'		sPayType="Card"
		'		sTourID="15S999A"
		'		sTourName="2015 Goode National Water Ski Championships"
		'END IF 

	  IF sHQAccount=true THEN 
	   		%>
	   		<TD Align="Left">
		 			<font size="1">When you click the 'Pay Now' button, you will be directed to make a payment through a secure server from USA Water Ski.  You may use a credit card only. Once you begin the payment process, do not stop until you reach the 'Receipt" page as this registration session may expire.  If this occurs, just restart your registration for this tournament and make the required payment.  Your registration information is stored, but <b><u> your registration will not be active until payment has been made.</u></b> </font>
	   	  	<br><br>
		 			<font size="1">A USA Water Ski registration receipt will be emailed to you following payment.  <b>Please retain this receipt as this is the only proof of payment you will receive.</b>  For refunds, credits or other matters relating to entry fees and payments please contact USA Water Ski Competition Dept at 800-533-2972.</font>
				</TD>
				<%
	   ELSE 
	   		%>	
	   		<TD Align="Left">
		 			<font size="1">When you click the 'Pay Now' button, you will be directed to make a payment to the PayPal account for the <%=sTourName%>.  You can pay using your PayPal account, or you may use a credit card.  If you elect to set up a new PayPal account, this registration session may expire.  If this occurs, once your PayPal account has been verified, just restart your registration for this tournament.  Your registration information is stored, but <b><u> your registration will not be active until payment has been made. </u></b></font>
		 			<br><br>
	   	  	<font size="1">In addition to the USA Water Ski registration receipt that will be emailed to you following payment, you will also receive a separate PayPal receipt.  <b>Please retain the PayPal receipt as this is the only proof of payment you will receive.</b>  Refunds, credits or other matters relating to entry fees and payments should be directed to the tournament organizer, or to the contact information on your PayPal receipt.</font>
				</TD>
				<%
	   END IF 
	  %>	
		</TR>
		<%
		' IF adminmenulevel>49 AND LEFT(sTourID,6)="15S119" THEN
		' IF adminmenulevel>49 AND RIGHT(LEFT(sTourID,6),3)="999" THEN
		IF RIGHT(LEFT(sTourID,6),3)="999" THEN
	   		
	   		%>	
	   		<TR>
	   			<TD Align="center">
						<font size="3" color="red"><b><br>REFUND POLICY</b> </font>
						<br><br>
					</TD>
				</TR>	
	   		<TR>
	   			<TD Align="Left">
						<font size="2" color="<%=TextColor3%>" ><b>$35 of the entry fee (per person)</b> is an administration and processing fee and is non-refundable. Late fees are also non-refundable. If you registered for the <%=sTourName%>, and are you are unable to participate you must <b>submit a cancelation request to USA Water Ski in writing</b> prior to the start of your first event. If you do not submit a cancelation request, you will not receive a refund. Cancelation requests will be honored due to lack of qualification or medical with documentation provided. All other excuses will be evaluated by the president of AWSA when the cancellation notice is received to determine if a refund will be issued. USA Water Ski will pay refunds <b>within 60 days following the conclusion of Nationals.</b> Please send cancellation requests and any documentation to competition@usawaterski.org</font>
						<br><br>
					</TD>
				</TR>	
				<%
   	END IF 
	  %>	

	</TABLE>
	<br>
	<TABLE class="innertable" Align=center width="90%">
	  <TR>
	    <TH ><font size="1" color="#FFFFFF"> Item #</font></th>
	    <TH><font size="1" color="#FFFFFF">&nbsp;Description</font></th>
	    <TH align=center><font size="1" color="#FFFFFF">&nbsp;Quantity</font></th>
	    <TH align="right"><font size="1" color="#FFFFFF">Amount&nbsp;</font></TH>
	  </TR><%
	
	FOR ItemNo=1 TO 9
		  IF TRIM(Item_Name(ItemNo))<>"" THEN %> 	  
		  <TR>
		    <TD align=center><font size="1"><%=ItemNo%></td>
		    <TD><font size="1">&nbsp;<%=Item_Name(ItemNo)%></td>
		    <TD align=center><font size="1">&nbsp;<%=Quantity(ItemNo)%></td>
		    <TD align="right"><font size="1">&nbsp;<%=formatcurrency(Amount(ItemNo)*Quantity(ItemNo),2)%>&nbsp;</td>
		  </TR><%
		  END IF
	NEXT  %>

	  <TR><TD align=center colspan=4>&nbsp;</TD></TR>
	  <TR>
	    <TD align=center>&nbsp;</td>
	    <TD>&nbsp;</td>
	    <TD align=right><font size="1">&nbsp;TOTAL ALL</td>
	    <TD align="right"><font size="1">&nbsp;<%=formatcurrency( (Amount(1)*Quantity(1)) + (Amount(2)*Quantity(2)) + (Amount(3)*Quantity(3)) + (Amount(4)*Quantity(4)) + (Amount(5)*Quantity(5)) + (Amount(6)*Quantity(6)) + (Amount(7)*Quantity(7)) + (Amount(8)*Quantity(8)) + (Amount(9)*Quantity(9)),2)%>&nbsp;</td>
	  </TR>

	</TABLE>
	<br>

   	<br><%	

'response.write("<br>Amount(1)="&Amount(1))
'response.write("<br>Amount(2)="&Amount(2))
'response.write("<br>Amount(3)="&Amount(3))
'response.write("<br>Amount(4)="&Amount(4))
'response.write("<br>Amount(5)="&Amount(5))
'response.write("<br>Amount(6)="&Amount(6))
'response.write("<br>Item_Name(6)="&Item_Name(6))

'response.write("<br>")
'response.write(Amount(2)+Amount(3)+Amount(4)+Amount(5)+Amount(6)+Amount(7)+Amount(8)+Amount(9))
'response.write(Amount(1)+Amount(2)+Amount(3)+Amount(4)+Amount(5)+Amount(6)+Amount(7)+Amount(8)+Amount(9))

'response.end

	
   	Dim ThisInvAmt
   	ThisInvAmt = Amount(1)+Amount(2)+Amount(3)+Amount(4)+Amount(5)+Amount(6)+Amount(7)+Amount(8)+Amount(9)
	
	IF ThisInvAmt > 0 THEN 
			%>
      <table align="center" width=90%>
				<tr>
					<td align=center>
		    		<hr>
		    		<%
		    		IF sHQAccount=true THEN 
		    				%><font size="2" color="<%=TextColor1%>" ><b>Click 'Pay Now' to pay for your Online Registration using your Credit Card.</b></font><%
		    		ELSE 
		    				%>
								<font size="2" color="<%=TextColor1%>" ><b>Click 'Pay Now' to pay for your Online Registration using your PayPal account or Credit Card.<br><u>If you do have a PayPal account</u> and do not wish to establish one for future transactions, click on the ‘Continue’ button near the credit card images on the first PayPal screen.</b></font>
								<br><br>
								<font size="3" color="red"><b>IMPORTANT</b> </font>
								<br><br>			
								<font size="2" color="<%=TextColor3%>" >Once you have completed the PayPal transaction, to complete your registration you must press the button titled <u><b>Finalize Your Registration and Get Receipt</b></u> located on the final PayPal screen.</font>
								<br><br><%
		    		END IF 
		    		%>	
		  		</td>
				</tr>
			</table>
			<%


		IF sPayType="Card" THEN 	

				' ------------------------------- 
				' --- Send to HQ CC processor ---
				' ------------------------------- 
				IF LCASE(LEFT(Session("sRelease"),4))="adlt" OR LCASE(LEFT(Session("sRelease"),4))="pbct" THEN ppf=1
						%>
			  		<br>
						<table align="center" width=90%>
							<tr>
								<form action="/rankings/<%=CardFileName%>?action=new&ppf=<%=ppf%>&CCAmount=<%=ThisInvAmt%>&sOrderNo=<%=sOrderNo%>&sMemberID=<%=SMemberID%>&sTourID=<%=sTourID%>" method=post>
	        				<td colspan=3 align="center">
				  					<input type="submit" name="CreditCard" value="Pay Now" style="width:9em" title="Click here to proceed to secure Payment Page">
			    	  			<br>
			    	  			<br>
				  					<hr>
									</td>
				      	</form>
					    </tr>
			  		</table>
			  		<%

		ELSE	

				' ------------------------------
				' --- PayPal Button and POST --- 
				' ------------------------------

				notify_URL="http://usawaterski.org/rankings/PayPal_IPN.asp?sMemberID="&sMemberID&"&sTourID="&sTourID
		    %>
				<table align="center" width=90%>
					<tr>
						<td colspan=3 align="center">
						  <form action="<%=sPayPalActionURL%>" method=POST name="PPForm">
								<input type=hidden value="<%=sPayPalAct%>" name="business">
								<input type=hidden value="_cart" name="cmd">
								<input type=hidden value="1" name="upload">
								<%
							
								FOR ItemNo=1 TO 9
			  						IF TRIM(Item_Name(ItemNo))<>"" THEN	
												thisitemname="item_name_"&ItemNo
												thisamountname="amount_"&ItemNo
												thisquantityname="quantity_"&ItemNo 
												%>
												<input type=hidden value="<%=item_name(ItemNo)%>" name="<%=thisitemname%>"> 	
												<input type=hidden value="<%=round(amount(ItemNo),2)%>" name="<%=thisamountname%>"> 	
												<input type=hidden value="<%=quantity(ItemNo)%>" name="<%=thisquantityname%>"><% 
			   						END IF 
								NEXT  
							
								%>
								<input type=hidden value="<%=sOrderNo%>" name="invoice">
								<input type=hidden value="<%=ReturnURL%>" name="return">
								<input type=hidden value="<%=ReturnURLBad%>" name="cancel_return">
								<input type=hidden value="<%=notify_URL%>" name="notify_URL">

								<input type=hidden value="2" name="rm">
					 			<input type=hidden value="####  IMPORTANT - CLICK HERE TO FINALIZE REGISTRATION ####" name="cbt">


								<input type=hidden value="<%=simage_url%>" name="cpp_header_image">
								<input type="hidden" name="no_shipping" value="0">
								<input type="hidden" name="no_note" value="1">
								<input type="hidden" name="lc" value="US">
								<input type="hidden" name="currency_code" value="USD">
								<input type="hidden" name="bn" value="PP-BuyNowBF">

					
								<input type="image" src="https://www.paypal.com/en_US/i/btn/btn_paynow_LG.gif" border="0" name="submit" alt="Make payments with PayPal - it's fast, free and secure!">
								<img alt="" border="0" src="https://www.paypal.com/en_US/i/scr/pixel.gif" width="1" height="1">
			    			<br>
			    			<br>
							</form>
						</td>
					</tr>
				</table>
				<%


				' --------------------------------------------
				' --- Displays optional Pay on Site button ---
				' --------------------------------------------

	
				IF sAllowOfflinePmt<>0 THEN
						%>
						<table align="center" width=90%>
							<tr>
								<td colspan=3 align="center">
									<font size="3" color="<%=TextColor1%>"><b>OPTIONAL PAYMENT METHOD</b> </font>
									<br><br>	
									<font size="2" color="<%=TextColor1%>" >This tournament has elected to allow payment by mail or at the site.  <b>If you elect to pay in this manner your registration will NOT be complete</b> and you may not be allowed to compete in this tournament. Depending on when your payment is received, late fees may apply.  It is your responsibility to make payment and confirm your eligibility.</font>
									<br><br>
								</td>
							</tr>	
							<tr>
								<td colspan=3 align="center">
				    		  <form method="post" action="/rankings/<%=RegFileName%>" id="OfflinePaymentForm">
										<input type="hidden" name="nav" value=7>
										<input type="submit" value="Pay On Site" style="width:10em;">
									</form>
								</td>
							</tr>
						</table>	
						<br><br>
						<%
						END IF



						%>	
					<hr>
					<%

		END IF		' --- Bottom of condition for PayPal or Card ---

	ELSEIF ThisInvAmt <= 0 THEN %>

		<table align="center" width=90%>
		  <tr>
		    <td align=center>
			<hr><%
			IF ThisInvAmt<0 THEN %>	
				<font size="2" color="<%=TextColor2%>" ><b>PayPal refunds cannot be initiated through this system. Please contact the tournament registrar, LOC or the contact listed on your PayPal receipt.</b></font><% 
				sPayType="ByPass"
			ELSE  %>	
				<font size="2" color="<%=TextColor2%>" ><b>No Payment is Due.  Press the 'Continue' button to go to receipt tab.</b></font><% 
				sPayType="NoSale" 
			END IF  %>	
		    </td>
		  </tr><%


		' --- ByPasses Payment altogether --- %>
		  <tr>
		  <form action="http://usawaterski.org/rankings/<%=RegFileName%>" method=POST name="defaultform">
		    <td width=25% align="center">
			<br>
			<input type="submit" name="Continue" value="Continue" style="width:9em" title="Continue to Print Receipt Page">
			  <input type="hidden" name="sPayType" value="<%=sPayType%>">
			  <input type="hidden" name="sTourID" value="<%=sTourID%>">
			  <input type="hidden" name="sMemberID" value="<%=sMemberID%>">
			  <input type="hidden" name="sOrderNo" value="<%=sOrderNo%>">
			  <input type="hidden" name="nav" value="7">
		    <hr>
		  </td>
		</form>
		</tr>
	      </table><%

	END IF 




	' -----------------------------------------------------------------------------------------------------------------------
	' --- Displays button and additional fields for confirming payment, checks, cash and credits when AdminCode is active ---
	' -----------------------------------------------------------------------------------------------------------------------

' WAS - sPayAmount
			
	IF (sDispDebugButtons=true OR adminmenulevel>=20 OR TestValidAdminCode) THEN

		' --- Simulates payment variables returned from PayPal SUCCESS--- %>
		<table align="center" width=90%>

	   <form name="PaymentForm3" method="post" action="/rankings/<%=RegFileName%>" id="PaymentlForm3">
		  <input type="hidden" name="sTourID" value="<%=sTourID%>">
		  <input type="hidden" name="sMemberID" value="<%=sMemberID%>">
		  <input type="hidden" name="sOrderNo" value="<%=sOrderNo%>">
		  <input type="hidden" name="nav" value="7">
		  <input type="hidden" name="SpecialAction" value="Y">
		  <tr>
		    <td align="center" colspan=3>
					<FONT size=<%=fontsize3%> color="<%=textcolor3%>"><b>ATTENTION REGISTRAR</b></FONT>
		    </td>
		  </tr>
		  <tr><td colspan=3>&nbsp;</td></tr>
		  <tr>
		    <td align="left" colspan=3>
					<FONT size=<%=fontsize2%>>Use the field and drop down below to record payments made by Cash or Check.  If you received an email from PayPal acknowledging this member's payment to your PayPal account, but the online registration system still shows a 'Balance Due', set the dropdown to 'Confirm PayPal Payment' and press 'Submit' to finalize the member's registration and acknowledge that you received the payment amount shown in the box below.</FONT>			    
		    </td>
		  </tr>
		  <tr><td colspan=3>&nbsp;</td></tr>
		  <tr>
		    <td width=20% align="center">
					<FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor3 %>"><b>Amount:&nbsp;</b></FONT>
					<input type="text" name="sPayAmount" value="<%=formatnumber(ThisInvAmt,2)%>" MAXLENGTH=7 size=7 Align="right">
		    </td>
		    <td width=40% align="center">
					<FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR="<% =textcolor3 %>"><b>Payment Type:&nbsp;</font>
					<select name="sPayType" value="<%=sPayType%>" style="width:15em">
					  <option value ="" <%IF sPayType = "" THEN Response.Write(" selected ")%> >No Payment</Option><br>
			  		<option value ="PayPal" <%IF sPayType = "PayPal" THEN Response.Write(" selected ")%> >Confirm PayPal Payment</Option><br>
			  		<option value ="Check" <%IF sPayType = "Check" THEN Response.Write(" selected ")%> >Receive Check</Option><br>
			  		<option value ="Cash" <%IF sPayType = "Cash" THEN Response.Write(" selected ")%> >Receive Cash</Option><br>
			  		<option value ="Refund" <%IF sPayType = "Refund" THEN Response.Write(" selected ")%> >Issue Refund</Option><br>
					</select>
		    </td>
		    <td width=20% align="center">
					<input type="submit" name="PayTestSuccess " value="Submit" style="<%=AdminButtonStyle%>" title="Press SUBMIT to record amount shown according to drop down setting" <%=ByPassButtonStatus%>>
		    </td>
		  </tr>
		  <tr>
		    <td>&nbsp;</td>
		  </tr>

		</form>
		</table><%
	END IF	



	' ---------------  BEBUGGING BUTTONS  ------------------

	IF sDispDebugButtons=true OR (sDispDebugButtonsAdm=true AND adminmenulevel>=20) THEN   %>
		  <table align=center width="90%">
		    <tr>
			<hr>
			<% ' ---- This is for the PayPal Sandbox Site --- %>
		    <form action=https://www.sandbox.paypal.com/cgi-bin/webscr method=post name="PPForm">
					<INPUT type=hidden value="mark_api1.kingsbridgehomes.com" name="business">
					<INPUT type=hidden value="_cart" name="cmd">
					<INPUT type=hidden value="1" name="upload">

		   		<INPUT type=hidden value="<%=item_name_1%>" name="item_name_1"> 	
					<INPUT type=hidden value="<%=amount_1%>" name="amount_1"> 	
					<INPUT type=hidden value="<%=quantity_1%>" name="quantity_1">
					<% 
					IF item2="true" THEN	%>
							<INPUT type=hidden value="<%=item_name_2%>" name="item_name_2"> 	
							<INPUT type=hidden value="<%=amount_2%>" name="amount_2"> 	
							<INPUT type=hidden value="<%=quantity_2%>" name="quantity_2"><% 
					END IF
					IF item3="true" THEN	%>
							<INPUT type=hidden value="<%=item_name_3%>" name="item_name_3"> 	
							<INPUT type=hidden value="<%=amount_3%>" name="amount_3"> 	
							<INPUT type=hidden value="<%=quantity_3%>" name="quantity_3"><% 
					END IF 
					
					%>
					<INPUT type=hidden value="<%=sOrderNo%>" name="invoice">
					<INPUT type=hidden value="<%=ReturnURL%>" name="return">
					<INPUT type=hidden value="<%=ReturnURLBad%>" name="cancel_return">
					<INPUT type=hidden value="<%=simage_url%>" name="cpp_header_image">
					<INPUT type=hidden value="2" name="rm"> 
					<input type="hidden" name="no_shipping" value="0">
					<input type="hidden" name="no_note" value="1">
					<input type="hidden" name="lc" value="US">
					<input type="hidden" name="currency_code" value="USD">
					<input type="hidden" name="bn" value="PP-BuyNowBF">

					<td width=25% align="center">
						<input type="submit" name="Sandbox" value="Sandbox" style="width:9em" title="Sandbox">
					</td>
		    </form><%

			' --- Simulates payment variables returned from PayPal FAIL--- %>
		        <form name="PaymentForm4" method="post" action="/rankings/<%=RegFileName%>" id="PaymentlForm4">
			  <input type="hidden" name="sTourID" value="<%=sTourID%>">
			  <input type="hidden" name="sMemberID" value="<%=sMemberID%>">
			  <input type="hidden" name="sOrderNo" value="<%=sOrderNo%>">
			  <input type="hidden" name="nav" value="7">
			  <input type="hidden" name="sPayType" value="PPErr">
			  <td width=25% align="center">
				<input type="submit" name="PayTestFail" value="Pay Test FAIL" style="width:9em" title="Bypass PayPal or Merchant Account and Simulate FAILED Payment" <%=ByPassButtonStatus%>>
			  </td>
			</form><%

			' --- ByPasses Payment altogether --- %>
			<form action="http://usawaterski.org/rankings/<%=RegFileName%>" method=POST name="PPForm">
			  <input type="hidden" name="sTourID" value="<%=sTourID%>">
			  <input type="hidden" name="sMemberID" value="<%=sMemberID%>">
			  <input type="hidden" name="sOrderNo" value="<%=sOrderNo%>">
			  <input type="hidden" name="nav" value="7">
			  <td width=25% align="center"><input type="submit" name="ByPassButton" value="ByPass Payment" style="width:9em" title="Bypass Payment Screen" <%=ByPassButtonStatus%>></td>
			</form>
		  </tr>
		 </table>
	      <br><%
		END IF 
		%>

    </div>
  </div><%
  
 




 ' -----------------------------  RECEIPT AND NOTICES  -------------------------------- %>


<% IF sMemberID="001001151" THEN response.end %>

  <div class="<% IF nav=7 THEN response.write("accordionHeaderSelected") ELSE response.write("accordionHeader") END IF %>"><%
    IF nav<7 THEN 
	%>STEP 7 - Receipt<% 
    ELSE 
	%><a href="/rankings/<%=RegFileName%>?nav=7">STEP 7 - Receipt</a><% 
    END IF %>
   
  </div>




  <div class="innertable" style="display:<% IF nav=7 THEN response.write("block") ELSE response.write("none") END IF %>;">
    <div id="RegPanel7">

	<br>
	<TABLE ALIGN="Center" class="innertable" width=80% >

	  <tr>
	    <th colspan=8 align=center>
		<font size="3" color="#FFFFFF"><b>Registration Complete</b></font>
	    </th>
	  </tr>


	  <tr>
	    <td colspan=8 align="center">
		<br>
		<font face="<% =font1 %>" size="2"><b>You have completed your registration for</b></font>
		<br><br>
		<font face="<% =font1 %>" size="3" color="<% =textcolor2 %>"><b><% =sTourName %></b></font>
		<br><br>   
	    </td>
	  </tr> 	


	  <tr>
	    <td align="center" colspan=8><%
		IF TRIM(ReceiptNote1)<>"" THEN %>
			<font color="<% =textcolor1 %>" size=<% =fontsize2 %> >1. <%=ReceiptNote1%></font>
			<br><%
		END IF
		IF TRIM(ReceiptNote2)<>"" THEN %>
			<font color="<% =textcolor1 %>" size=<% =fontsize2 %> >2. <%=ReceiptNote2%></font>
			<br><%
		END IF
		IF TRIM(ReceiptNote3)<>"" THEN %>
			<font color="<% =textcolor1 %>" size=<% =fontsize2 %> >3. <%=ReceiptNote3%></font>
			<br><%
		END IF
		IF TRIM(ReceiptNote4)<>"" THEN %>
			<font color="<% =textcolor1 %>" size=<% =fontsize2 %> >4. <%=ReceiptNote4%></font>
			<br><%
		END IF
		IF TRIM(ReceiptNote5)<>"" THEN %>
			<font color="<% =textcolor1 %>" size=<% =fontsize2 %> >5. <%=ReceiptNote5%></font>
			<br><%
		END IF %> 
		<br><br>
	    </td>
	  </tr> 	

	  <tr>
	    <td align=center colspan=8>
		<font face="<% =font1 %>" size="2" color="red">Session Ended - Do not use expired pages!</font>
		<br><br>
		<font face="<% =font1 %>" size="2"><b>For questions, please contact the Tournament Registrar at <%=sTRegistrarPhone%>.</b></font>
		<br><br> 
		<font face="<% =font1 %>" size="3"><b>Thank you.</b>
		<br>
	    </td>
	  </tr>
	  <tr><td colspan=8>&nbsp;</td></tr>


	  <tr>
	    <td align="center" colspan=2 width=25%>
		<form action="/rankings/<%=RegFileName%>?sRunByWhat=ReturnToMainMenu" method="post">
		  <input type="submit" value=" Main Menu "  style="width:9em" title="Leave registration and return to the main menu. IMPORTANT: Once you return to main menu, your 'session' will end and you must log in again to print receipt or view information.">
		</form>
	    </td><%	

  	' --- Lets Admin Users return to Member Page ---
	  IF TestValidAdminCode=true OR adminmenulevel >= 20 THEN 

	  		%>
	      <form name="ReceiptForm3" method="post" action="/rankings/<%=RegFileName%>" id="ReceiptForm3">
	  			<input type="hidden" name="sTourID" value="<%=sTourID%>">
	  			<input type="hidden" name="sMemberID" value="<%=sMemberID%>">
	  			<input type="hidden" name="nav" value="2">
		  		<td align="center" colspan=2 width=25%>
						<input type="submit" name="BackToMember" value="New Member" style="<%=AdminButtonStyle%>" title="Admin users - Return to Member tab to select a NEW member without losing your Administrative login information." <%=MainButtonStatus%>>
			  	</td>
		  	</form>
				<%
	  ELSE 
	  		%>
	      <form name="ReceiptForm3" method="post" action="/rankings/<%=RegFileName%>" id="ReceiptForm3">
	  			<input type="hidden" name="sTourID" value="<%=sTourID%>">
	  			<input type="hidden" name="sRunByWhat" value="NewMember">
		  		<td align="center" colspan=2 width=25%>
						<input type="submit" name="NewMember" value="New Member" style="<%=UserButtonStyle%>" title="Register a different member." <%=MainButtonStatus%>>
			  	</td>
		  	</form>
				<%
	  END IF

	    ' --- Debugging buttons display (not used any more) ---	
	    IF sDispDebugButtons=true OR (sDispDebugButtonsAdm=true AND adminmenulevel>=20) THEN 
	    		%>
					<td align="center" colspan=2 width=25%>
						<form action="/rankings/<%=RegFileName%>?sRunByWhat=DeletePayments" method="post">
		    		   <input type="submit" value="Delete Pays" style="width:9em" title="Delete all record of payments and transactions for Trans >=2000">
						</form>	
					</td>
					<%
	    END IF %>	

	  	<td width=25% align="center">
				<form name="NewEntryForm" method="post" action="/rankings/<%=RegFileName%>?sRunByWhat=Tour" id="NewEntryForm">
		    	 <input type="submit" name="NewTour" value="New Entry" style="<%=SpecialUserButtonStyle%>" title="Create another online entry for the same member" <%=MainButtonStatus%>>
				</form>
	  	</td>

	  	<td align="center" colspan=2 width=25%>
				<form action="/rankings/<%=RegFileName%>?sRunByWhat=Print" method="post" target="_blank">
		   		<input type="submit" value="Print Receipt"  style="width:9em" title="Displays complete entry receipt on screen, which may be printed for your records.">
				</form>	
	  	</td>	
	  </tr>

	  <tr><td colspan=8>&nbsp;</td></tr> 	

	</TABLE>

    </div>
  </div>

</div><% ' --- Outer Accordion div --- %>

</body><%


END SUB




' --------------------------------
   SUB SetHiddenFinancialVariables
' --------------------------------

	' --- These are the variables that are on the Financial Form --- %>
	<input type="hidden" name="fAWSEFCheck" value="<%=sAWSEFCheck%>">
	<input type="hidden" name="fOfficial" value="<%=sOfficial%>">
	<input type="hidden" name="fClubMemb" value="<%=sClubMemb%>">
	<input type="hidden" name="fClubCode" value="<%=sClubCode%>">
	<input type="hidden" name="sMoneyOverride" value="<%=sMoneyOverride%>">
	<input type="hidden" name="sBanquetQty" value="<% =sBanquetQty %>">


	<input type="hidden" name="sOF1Qty" value="<% =sOF1Qty %>">
	<input type="hidden" name="sOF2Qty" value="<% =sOF2Qty %>">
	<input type="hidden" name="sOF3Qty" value="<% =sOF3Qty %>">
	<input type="hidden" name="sOF4Qty" value="<% =sOF4Qty %>">
	<input type="hidden" name="sOF5Qty" value="<% =sOF5Qty %>">
	<input type="hidden" name="sOF6Qty" value="<% =sOF6Qty %>">
	<input type="hidden" name="sOF7Qty" value="<% =sOF7Qty %>">
	<input type="hidden" name="sOF8Qty" value="<% =sOF8Qty %>">
	<input type="hidden" name="sOF9Qty" value="<% =sOF9Qty %>">
	<input type="hidden" name="sOF10Qty" value="<% =sOF10Qty %>">


	<input type="hidden" name="sOF1Fee" value="<% =sOF1Fee %>">
	<input type="hidden" name="sOF2Fee" value="<% =sOF2Fee %>">
	<input type="hidden" name="sOF3Fee" value="<% =sOF3Fee %>">
	<input type="hidden" name="sOF4Fee" value="<% =sOF4Fee %>">
	<input type="hidden" name="sOF5Fee" value="<% =sOF5Fee %>">
	<input type="hidden" name="sOF6Fee" value="<% =sOF6Fee %>">
	<input type="hidden" name="sOF7Fee" value="<% =sOF7Fee %>">
	<input type="hidden" name="sOF8Fee" value="<% =sOF8Fee %>">
	<input type="hidden" name="sOF9Fee" value="<% =sOF9Fee %>">
	<input type="hidden" name="sOF10Fee" value="<% =sOF10Fee %>">

	<input type="hidden" name="sEntryType" value="<% =sEntryType %>">
	<input type="hidden" name="sMembRegDate" value="<% =sMembRegDate %>">
	<input type="hidden" name="sWaiverCode" value="<% =sWaiverCode %>">
	<input type="hidden" name="sSignWaiver" value="<% =sSignWaiver %>"><%

END SUB


' ------------------------
   SUB MakeHiddenEntryForm	  
' ------------------------

	  FOR EvtNo = 1 TO TotEv
		fSelectEvent="fSelectEvent"&EvtNo
		fDiv="fDiv"&EvtNo  				 
		fFeeClass="fFeeClass"&EvtNo
		fFeeRounds="fFeeRounds"&EvtNo
		fQfyOverride="fQfyOverride"&EvtNo  
		fBoat="fBoat"&EvtNo  
		fSkill="fSkill"&EvtNo  %>
		  <input type="hidden" name="<%= fSelectEvent %>" value="<% =sSelectEvent(EvtNo) %>">
		  <input type="hidden" name="<%=fDiv%>" value="<% =sDiv(EvtNo) %>">
		  <input type="hidden" name="<%=fFeeClass%>" value="<% =sFeeClass(EvtNo) %>">
		  <input type="hidden" name="<%=fFeeRounds%>" value="<% =sFeeRounds(EvtNo) %>">
		  <input type="hidden" name="<%=fQfyOverride%>" value="<% =sQfyOverride(EvtNo) %>">
		  <input type="hidden" name="<%=fBoat%>" value="<%=sBoat(EvtNo)%>">
		  <input type="hidden" name="<%=fSkill%>" value="<%=sSkill(EvtNo)%>"><%
	  NEXT %>

	  <input type="hidden" name="sRegionalOverride" value="<% =sRegionalOverride %>">
	  <input type="hidden" name="sRampHeight" value="<%=sRampHeight%>">
	  <input type="hidden" name="sWaiverCode" value="<% =sWaiverCode %>">
	  <input type="hidden" name="sSignWaiver" value="<% =sSignWaiver %>"><%


END SUB



' ------------------------------
  SUB DisplayPertinentVariables
' ------------------------------

response.write("IN Display ")%><br><%
response.write("sEntryFee = "&sEntryFee)%><br><%
response.write("sLateFeeTot = "&sLateFeeTot)%><br><%
response.write("sAWSEFDonation ="&sAWSEFDonation)%><br><%
response.write("sBanquetTot = "&sBanquetTot)%><br><%
response.write("sJrDiscAmt = "&sJrDiscAmt)%><br><%
response.write("sSrDiscAmt = "&sSrDiscAmt)%><br><%
response.write("sClubDiscAmt = "&sClubDiscAmt)%><br><%
response.write("sOffDiscAmt = "&sOffDiscAmt)%><br><%

response.write("sTotalFormFees = "&sTotalFormFees)%><br><%
response.write("cdbl(sTotalPreviousPayments) = "&cdbl(sTotalPreviousPayments))
'response.end

END SUB



' -----------------------
    SUB AccordRelease
' -----------------------

	waivernav=Request("waivernav")
	IF waivernav="" THEN waivernav=1



' --- New 4-28-2013 - Gets SPECIAL WAIVER info from table based on SiteID rather than hard coding specific tournaments ---
Dim swaiverSQL, sSpecialWaiverHeadline, sSpecialReleaseBannerText
swaiverSQL = "SELECT SpecialWaiverCode, SpecialWaiverHeadline, SpecialReleaseBannerText FROM usawsrank.TourExtras TE"
swaiverSQL = swaiverSQL + " JOIN sanctions.dbo.TSchedul AS TS"
swaiverSQL = swaiverSQL + "   ON SiteID=TS.TSiteID"
swaiverSQL = swaiverSQL + " WHERE LEFT(TS.TournAppID,6)='"&LEFT(sTourID,6)&"'"

Set rswaiver=Server.CreateObject("ADODB.recordset")
rswaiver.open swaiverSQL, sConnectionToTRATable, 3, 1

testwaiver=false
IF testwaiver=true AND sMemberID="000001151" THEN
		Response.write("<br>Found = ")
		response.write(NOT(rswaiver.eof))
		response.write("<br>rswaiver(SpecialWaiverHeadline) = "&rswaiver("SpecialWaiverHeadline"))
END IF

IF NOT(rswaiver.EOF) THEN
		sSpecialWaiverCode=rswaiver("SpecialWaiverCode")
		sSpecialWaiverHeadline=rswaiver("SpecialWaiverHeadline")
		sSpecialReleaseBannerText=rswaiver("SpecialReleaseBannerText")
		'response.write("<br>Line 2526 - sSpecialWaiverCode="&sSpecialWaiverCode)
END IF


' --- TESTING CODE and code to add a waiver for a new site ---
' SELECT * FROM usawsrank.RegisterGenNew WHERE TourID='13S999' AND MemberID='000001151'
' DELETE FROM usawsrank.RegisterGenNew WHERE TourID='13S999' AND MemberID='000001151'
' ALTER TABLE usawsrank.TourExtras ADD SpecialWaiverCode CHAR(8)
' ALTER TABLE usawsrank.TourExtras ADD SpecialWaiverHeadline VARCHAR(70)
' ALTER TABLE usawsrank.TourExtras ADD SpecialReleaseBannerText VARCHAR(70)

' UPDATE TE SET SpecialWaiverCode='pbct2012' FROM usawsrank.TourExtras TE WHERE SiteID='USAS0364'
' UPDATE TE SET SpecialWaiverHeadline='AMATEUR ATHLETIC WAIVER AND RELEASE OF LIABILITY' FROM usawsrank.TourExtras TE WHERE SiteID='USAS0364'
' UPDATE TE SET SpecialReleaseBannerText='Palm Beach County Waiver - Read Carefully' FROM usawsrank.TourExtras TE WHERE SiteID='USAS0364'

' --- Changed 4-28-2013 to use SiteID approach ---
'	IF LEFT(sTourID,6)="13S999" OR LEFT(sTourID,6)="12S999" OR LEFT(sTourID,6)="13S090" OR LEFT(sTourID,6)="13S091" OR LEFT(sTourID,6)="13S148" OR LEFT(sTourID,6)="13S130" OR LEFT(sTourID,6)="13S150" OR LEFT(sTourID,6)="13S151" OR LEFT(sTourID,6)="14S036" OR LEFT(sTourID,6)="14S037" OR LEFT(sTourID,6)="14S041" OR LEFT(sTourID,6)="13S195" OR LEFT(sTourID,6)="13S196" OR LEFT(sTourID,6)="13S197" OR LEFT(sTourID,6)="13S198" THEN

		' --- OR LEFT(sTourID,6)="11M999" 
		' --- Per Dale Stevens WPBSC 1-24-2013 ---
		' 13S090, 13S091, 13S148, 13S130, 13S150, 13S151, 14S036, 14S037, 14S041
		' --- New ones
		' 13S195, 13S196, 13S197, 13S198

		'	sSpecialWaiverCode="pbct2012"
		'	sSpecialWaiverHeadline="AMATEUR ATHLETIC WAIVER AND RELEASE OF LIABILITY"
		'	sSpecialReleaseBannerText="Palm Beach County Waiver - Read Carefully"
'	END IF



	' --- Test for tournaments to use special survey form ---
	sSurveyForm=""
	SELECT CASE LEFT(sTourID,6)
			CASE "12S999"
					sSurveyForm="12S999"			
			CASE "11M999"
					sSurveyForm="12S999"
			CASE "13S999"
					sSurveyForm="13S999"
	END SELECT

'Response.write("<br>sSurveyForm = "&sSurveyForm)


	' --- Request variables related to waiver ---
	sRelease = TRIM(Request("fRelease"))
	sSignWaiver = sqlclean(TRIM(Request("sSignWaiver")))
	sReleaseType = TRIM(Request("sReleaseType"))

	' --- If blank and NOT an Admin User then OK to default to Electronic ---
	IF sReleaseType = "" AND (NOT TestValidAdminCode) THEN 
			sReleaseType="Electronic"
	ELSEIF sReleaseType = "" AND TestValidAdminCode THEN 
			sReleaseType="None"
	END IF

	' - CASE 1 - Notice
	' - CASE 3 - Std USA Water Ski waiver 
	' - CASE 3 - Special Waiver
	' - CASE 4 - Waiver Acknowledgement
	' - CASE 5 - Save and if survey display page 1 otherwise set waivernav=10 to redirect
	' - CASE 6 - If survey display thank you
	' - CASE 10 - Redirect (used?)


	IF sMemberID="700144639" THEN
			response.write("<br>sReleaseType = "&sReleaseType)
			response.write("<br>sSignWaiver = "&sSignWaiver) 
	END IF

	' --- If Electronic then check to make sure either age>=18 or someone has signed the waiver as the guardian ---
	IF sReleaseType = "Electronic" AND (Session("sMembAge") >= 18 OR (Session("sMembAge") < 18 AND TRIM(sSignWaiver)<>"")) THEN sElectronicOK = true	


	ReleaseAccepted = "N"	
	IF sRelease="Accept" AND (sElectronicOK OR sReleaseType = "Paper" OR sReleaseType = "None") THEN
			ReleaseAccepted = "Y"
	END IF		
	IF waivernav=3 AND TRIM(sSpecialWaiverCode)="" AND ReleaseAccepted="Y" THEN
			' --- IF Release accepted AND one of the methods is met ---
			waivernav=4

	ELSEIF waivernav=4 AND TRIM(sSpecialWaiverCode)<>"" AND ReleaseAccepted="Y" THEN
			' --- Release accepted AND SPECIAL waiver AND one of the methods is met ---
			waivernav=4			

	ELSEIF waivernav=4 AND TRIM(sSpecialWaiverCode)<>"" THEN
			' --- IF Release NOT accepted AND in SPECIAL waiver section then redisplay waiver ---
			waivernav=3	

	ELSEIF waivernav=3 AND TRIM(sSpecialWaiverCode)="" THEN
			' --- Release NOT accepted and no Special waiver then redisplay waiver---
			waivernav=2	

	ELSEIF waivernav=5 THEN
			' --- Saves the registration waiver values to the RegGen table ---
			SaveWaiverValuesToTable

			' --- If survey then display survey page, else redirect to next tab in registration (Payment)
			IF TRIM(sSurveyForm)<>"" THEN
					BeginTournamentSurvey
			ELSE
					' --- No survey so exit waiver section to Payment page ---
					waivernav=10
			END IF
	ELSEIF waivernav=6 THEN		' --- Thank you page of survey which could only occur if survey ---
			BeginTournamentSurvey

	END IF		





SELECT CASE LCASE(TRIM(waivernav))
		CASE 1

			' -----------------------------------------------------------
			' ---------  Initial Heads Up Dialog Box on Warning  --------	
			' -----------------------------------------------------------
			%>
			<br>

			<TABLE BORDER="2" BORDERCOLOR="black" ALIGN="CENTER" BGCOLOR="<% =TableColor1 %>" width="100%">
	  		<TR>
	      	<TD BGCOLOR="red"><center><font face=<% =font1 %> color="#FFFFFF" size="4"><b>Important Notice !!</b></font></TD>
	  		</TR>  
 
			  <TR>
					<TD VALIGN="top">
  					<TABLE BORDER="0" VALIGN="top" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="<%=TableColor1%>" width=100%>
			   			<tr>
			      		<td VALIGN="top" ALIGN="center">
									<br>
									<font color="<% =TextColor1 %>" face="<% =font1 %>" size="4"><b><i>Read This Before Continuing !!</i></b></font>
									<br><br>
									<font face="<% =font1 %>" size="1">You may not participate in any USA Water Ski sanctioned event without accepting the terms of the following PARTICIPANT WAIVER AND RELEASE OF LIABILITY, ASSUMPTION OF RISK AND INDEMNITY.</font>
									<br><br>
									<font color="red" face="<% =font1 %>" size="2"><b>You must be 18 years old.</b></font> 			
									<br><br>	
									<font face="<% =font1 %>" size="1"><b>Unless you are accepting it as a parent of a minor, you may NOT accept the 'RELEASE' on behalf of another person.</b></font> 
									<br><br>
									<font face="<% =font1 %>" size="1">Accepting the 'RELEASE' for another adult (even a spouse), or accepting it on behalf of a minor for whom you are not the legal guardian, is strictly prohibited.  Such action may be subject to civil liability of criminal prosecution under <b> state and federal laws.</b></font>
			    			</td>
			  			</tr>
			  			<tr>
			    			<td align="center">
									<br>
									<form action="/rankings/<%=RegFileName%>?nav=5&waivernav=2" method="post">
				  						<center><input type="submit" value=" Continue "></center>
									</form>
			    			</td>	
			  			</tr>
						</TABLE>
		    	</TD>
		  	</TR>
			</TABLE>   
			<br>
			<% 

	
		CASE 2

			' --------------------------------------------
			'--- Display standard USA Water Ski Waiver ---
			' --------------------------------------------
			DefineWaiverSpecs
		
			%>
			<br>
			<TABLE BORDER="2" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="<% =TableColor1 %>" width=100%>
		  	<TR>
		      <TD BGCOLOR="red"><center><font face=<% =font1 %> color="#FFFFFF" size="4"><b>Waiver and Release Form</b></font></TD>
		  	</TR>  
 		  	<TR>
					<TD VALIGN="top">
  					<TABLE BORDER="0" VALIGN="top" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="<%=TableColor1%>" width=100%><%

						' ------------------------------------------------------------------------------------------
						' ----------  Displays release and sets FORM ACTION to rerun this condition  ---------------	
						' ------------------------------------------------------------------------------------------ 



						' --- Determines if a special waiver is to be used for the tournament in addition to std USA Waterski waiver ---
						'IF LEFT(sTourID,6)="12S999" OR LEFT(sTourID,6)="11M999" THEN 
						IF sSpecialWaiverCode="pbct2012" THEN
								WhichWaiverNav="3"
						ELSE
								WhichWaiverNav="4"
						END IF
						
						
						%>
						<form action = "/rankings/<%=RegFileName%>?waivernav=<%=WhichWaiverNav%>&nav=5" method="post">
	  		  		<tr>
								<td align=center>	
	 	   						<font face=<% =font1 %> size="4" ><b>PARTICIPANT WAIVER AND RELEASE OF LIABILITY,</b></font><br>
		   						<font face=<% =font1 %> size="4"><b>ASSUMPTION OF RISK AND INDEMNITY AGREEMENT</b></font>
		   						<br>
		   						<font face=<% =font1 %> size="2"><b><% =sWaiverSubTitle %></b></font>
		   						<br><br>

									<font face=<% =font1 %> color="<% =TextColor2 %>" size="3"><b><% =sTourName %></font></b>
									<br><br>
									<font face=<% =font1 %> size="2"><b>MemberID = </font><font color="<% =textcolor2 %>" face=<% =font1 %> size="2"><%=sMemberID%>
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="<% =textcolor1 %>" face=<% =font1 %> size="2">Participant:</font>
									<font color="<% =textcolor2 %>" face=<% =font1 %> size="2"><% =sFirstname %>&nbsp;<% =sLastName %></font></b><br>
								</td>
							</tr>

  		  			<tr>
		   					<td>
									<br>
		     					<p>
 										<font color="<% =textcolor1 %>" size="1" face=<% =font1 %>><left>
 											<%

			  						Set objfso = CreateObject("Scripting.FileSystemObject")
			  						IF objfso.FileExists(PathtoWaivers & "\waiver-"&sWaiverCode&".txt") THEN
												SET objstream=objFSO.opentextfile(PathtoWaivers & "\waiver-"&sWaiverCode&".txt")

												IF NOT objstream.atendofstream THEN
														DO WHILE not objstream.atendofstream
																response.write(objstream.readline)
					   										response.write("<br>")
														LOOP
												END IF
												objstream.close
			  						END IF 

			  						%>
										</font>
		    					</p>
		   					</td>
  		  			</tr>
							<tr>
		    				<td align="center">
		    				<%
									IF TestValidAdminCode OR adminmenulevel>=20 THEN 
											%>
			   							<br>
			   							<font size="3" color="red" ><b>As 'Tournament Registrar', you are responsible for collecting 
											<br>waivers and returning them to USA Water Ski competition services dept.</b></font>
											<%
									ELSE 
											%>
			   							<br>
			   							<font size="3" color="red" ><b>The name listed above must be the person completing this form.</b></font>
			   							<br>
			   							<font size="3" color="red" ><b>Minors under 18 Years may NOT accept liability waiver.</b></font>
			   							<%
									END IF 
									%>
		    				</td>
		  				</tr>

		  				<tr>
		  					<td align="center">
		     					<br>
		     					<%
		     					IF Session("sMembAge") < 18 AND (NOT TestValidAdminCode) THEN  
		     							%>
											<font color="<% =textcolor3 %>" size="2"><b>Name of Parent or Guardian acccepting waiver on behalf of </b></font>
											<font color="<% =textcolor2 %>" size="2"><b><%=sFirstName%>&nbsp;<%=sLastName%></b></font>
											<br>
											<input type="text" name="sSignWaiver" value= "<% =sSignWaiver %>" size="30" >
											<%
		     					ELSEIF Session("sMembAge") >= 18 AND (NOT TestValidAdminCode) THEN  
		     							%>
											<font color="<% =textcolor1 %>" size="2"><b>By acccepting this waiver I acknowledge that I am the 'PARTICIPANT' listed above.	</b></font>
											<br>
											<%
		     					ELSEIF TestValidAdminCode THEN  
		     							%>
											<font color="<% =textcolor1 %>" size="2"><b>Please select current 'Status' of Waiver form.</b></font>
											<br><%
		     					END IF 
		     					%>
		   					</td>
		   				</tr><%

							' -----    DECLINE USA Water Ski WAIVER ------
							IF sRelease="Decline" THEN 
									%>
		   						<tr>
		   							<td align="center">
			   							<font size="3" color="red" ><b>You May Not Enter Tournament Without Accepting the Waiver and Release</b></font>
		   							</td>
		   						</tr>
		   						<%
							END IF

							' -----    Admin Level and No Waiver Type Selected  ------
							IF sReleaseType="" AND adminmenulevel >=20  THEN  
									%>		  
		   						<tr>
		   							<td align="center">
											<br>
											<font size="4" color="red" ><b>Please select a WAIVER TYPE</b></font>
		   							</td>
		   						</tr>
		   						<%
							END IF  

							' --- Displays options for type of waiver ---
							IF adminmenulevel >= 10 OR TestValidAdminCode THEN  
									%>
  		  					<tr>
										<td align="center">
											<%
											IF adminmenulevel >= 10 THEN 
													%> 
			   									<input type=radio NAME="sReleaseType" VALUE="Electronic" <% IF sReleaseType="Electronic" THEN response.write "checked"%> >
			   									<FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR=<% =textcolor1 %> ><b>Electronic</b></font>
			   									<%
											END IF 
											%>
											<input type=radio NAME="sReleaseType" VALUE="Paper" <% IF sReleaseType="Paper" THEN response.write "checked"%>>
											<FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR=<% =textcolor1 %> ><b>Paper W/Signature</b></font>
											<input type=radio NAME="sReleaseType" VALUE="None" <% IF sReleaseType="None" THEN response.write "checked"%>>
											<FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR=<% =textcolor1 %> ><b>No Waiver</b></font>
										</td> 
  		  					</tr><%
							ELSE  
									%>
									<INPUT type="hidden" NAME="sReleaseType" VALUE="<%=sReleaseType%>" >
									<%
							END IF 	
							%>
		   				<tr>
		   					<td align="center">
									<br>
									<%
									IF TestValidAdminCode<>true THEN 
											%>	
											<font color="<% =textcolor1 %>" face=<% =font1 %> size="2"><b>Accept:</b></font><input type="radio" name="fRelease" <%IF sRelease="Accept" THEN Response.write("checked")%> value="Accept">
											<font color="<% =textcolor1 %>" face=<% =font1 %> size="2"><b>Decline:</b></font><input type="radio" name="fRelease" <%IF sRelease="Decline" THEN Response.write("checked")%> value="Decline">
											&nbsp;&nbsp;&nbsp;&nbsp 
											<%
									ELSE 
											%>
											<input type="hidden" name="fRelease" value="Accept">
											<%
									END IF 
									%>
									<input type="submit" value="Submit" style="width:9em">
									&nbsp;&nbsp;&nbsp;&nbsp
									<font color="<% =textcolor1 %>" face=<% =font1 %> size="2"><b>Date: <% =DATE %></b></font>
									<br><br>
		  					</td>
		  				</tr>

							</form>
			  		</TABLE>
						<br>

			  	</TD>
 		  	</TR>
			</TABLE>
		<%


	CASE 3

		
					
			' -----------------------------------------------------------------
			' --- SPECIAL WAIVER option - Does not apply to all tournaments ---
			' -----------------------------------------------------------------
			DefineWaiverSpecs			
			sRelease=""

			%>
			<br>
			<TABLE BORDER="2" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="<% =TableColor1 %>" width=100%>
		  	<TR>
		      <TD BGCOLOR="orange"><center><font face=<% =font1 %> color="#FFFFFF" size="4"><b><%=sSpecialReleaseBannerText%></b></font></TD>
		  	</TR>  
 		  	<TR>
					<TD VALIGN="top">
  					<TABLE BORDER="0" VALIGN="top" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="<%=TableColor1%>" width=100%><%

						' ------------------------------------------------------------------------------------------
						' ----------  Displays release and sets FORM ACTION to rerun this condition  ---------------	
						' ------------------------------------------------------------------------------------------ 
						%>
						<form action = "/rankings/<%=RegFileName%>?waivernav=4&nav=5" method="post">
							<input type="hidden" name="sSignWaiver" value="<%=sSignWaiver%>">
	  		  		<tr>
								<td align=center>	
	 	   						<font face=<% =font1 %> size="4" ><b><%=sSpecialWaiverHeadline%></b></font><br>
		   						<br>

									<font face=<% =font1 %> color="<% =TextColor2 %>" size="3"><b><% =sTourName %></font></b>
									<br><br>
									<font face=<% =font1 %> size="2"><b>MemberID = </font><font color="<% =textcolor2 %>" face=<% =font1 %> size="2"><%=sMemberID%>
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="<% =textcolor1 %>" face=<% =font1 %> size="2">Participant:</font>
									<font color="<% =textcolor2 %>" face=<% =font1 %> size="2"><% =sFirstname %>&nbsp;<% =sLastName %></font></b><br>
								</td>
							</tr>

  		  			<tr>
		   					<td>
									<br>
		     					<p>
 										<font color="<% =textcolor1 %>" size="1" face=<% =font1 %>><left>
 											<%

			  						Set objfso = CreateObject("Scripting.FileSystemObject")
			  						IF objfso.FileExists(PathtoWaivers & "\waiver-"&sSpecialWaiverCode&".txt") THEN
												SET objstream=objFSO.opentextfile(PathtoWaivers & "\waiver-"&sSpecialWaiverCode&".txt")

												IF NOT objstream.atendofstream THEN
														DO WHILE not objstream.atendofstream
																response.write(objstream.readline)
					   										response.write("<br>")
														LOOP
												END IF
												objstream.close
			  						END IF 

			  						%>
										</font>
		    					</p>
		   					</td>
  		  			</tr>
   						<tr>
   							<td align="center">
	   							<font size="3" color="red" ><b>I agree to be fully responsible for my conduct at the tournament and/or for the conduct of the minor on whose behalf I sign.</b></font>
   							</td>
   						</tr>
							<%

							' -----    DECLINE USA Water Ski WAIVER ------
							IF sRelease="Decline" THEN 
									%>
		   						<tr>
		   							<td align="center">
			   							<font size="3" color="red" ><b>You May Not Enter Without Accepting the Waiver and Release</b></font>
		   							</td>
		   						</tr>
		   						<%
							END IF


							' -----    Admin Level and No Waiver Type Selected  ------
							IF sReleaseType="" AND adminmenulevel >=20  THEN  
									%>		  
		   						<tr>
		   							<td align="center">
											<br>
											<font size="4" color="red" ><b>Please select a WAIVER TYPE</b></font>
		   							</td>
		   						</tr>
		   						<%
							END IF  

							' --- Displays options for type of waiver ---
							IF adminmenulevel >= 10 OR TestValidAdminCode THEN  
									%>
  		  					<tr>
										<td align="center">
											<%
											IF adminmenulevel >= 10 THEN 
													%> 
			   									<input type=radio NAME="sReleaseType" VALUE="Electronic" <% IF sReleaseType="Electronic" THEN response.write "checked"%> >
			   									<FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR=<% =textcolor1 %> ><b>Electronic</b></font>
			   									<%
											END IF 
											%>
											<input type=radio NAME="sReleaseType" VALUE="Paper" <% IF sReleaseType="Paper" THEN response.write "checked"%>>
											<FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR=<% =textcolor1 %> ><b>Paper W/Signature</b></font>
											<input type=radio NAME="sReleaseType" VALUE="None" <% IF sReleaseType="None" THEN response.write "checked"%>>
											<FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR=<% =textcolor1 %> ><b>No Waiver</b></font>
										</td> 
  		  					</tr><%
							ELSE  
									%>
									<INPUT type="hidden" NAME="sReleaseType" VALUE="<%=sReleaseType%>" >
									<%
							END IF 	
							%>
		   				<tr>
		   					<td align="center">
									<br>
									<%
									IF TestValidAdminCode<>true THEN 
											%>	
											<font color="<% =textcolor1 %>" face=<% =font1 %> size="2"><b>Accept:</b></font><input type="radio" name="fRelease" <%IF sRelease="Accept" THEN Response.write("checked")%> value="Accept">
											<font color="<% =textcolor1 %>" face=<% =font1 %> size="2"><b>Decline:</b></font><input type="radio" name="fRelease" <%IF sRelease="Decline" THEN Response.write("checked")%> value="Decline">
											&nbsp;&nbsp;&nbsp;&nbsp 
											<%
									ELSE 
											%>
											<input type="hidden" name="fRelease" value="Accept">
											<%
									END IF 
									%>
									<input type="submit" value="Submit" style="width:9em">
									&nbsp;&nbsp;&nbsp;&nbsp
									<font color="<% =textcolor1 %>" face=<% =font1 %> size="2"><b>Date: <% =DATE %></b></font>
									<br><br>
		  					</td>
		  				</tr>

							</form>
			  		</TABLE>
						<br>

			  	</TD>
 		  	</TR>
			</TABLE>
			<%


	CASE 4

			' ------------------------------------------
			' --- Notice that waiver has been signed ---
			' ------------------------------------------
	
		%>
		<br>
		<TABLE BORDER="2" BORDERCOLOR="black" ALIGN="CENTER" BGCOLOR="<% =TableColor1 %>" width=75%>
		  <TR>
		      <TD BGCOLOR="red"><center><font face=<% =font1 %> color="#FFFFFF" size="4"><b>Important Notice !!</b></font></TD>
		  </TR>  
 
		  <TR>
		     <TD VALIGN="top">
  			<TABLE BORDER="0" VALIGN="top" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="<%=TableColor1%>" width=100%>
			   <tr>
			      <td colspan=2 VALIGN="top" ALIGN="center">
				<br>
				<font color="<% =TextColor1 %>" face="<% =font1 %>" size="4"><b>Release Accepted !</b></font>
				<br><br>
				<font face="<% =font1 %>" size="2"><p>You have accepted the terms and conditions of the USA Water Ski PARTICIPANT WAIVER AND RELEASE OF LIABILITY, ASSUMPTION OF RISK AND INDEMNITY.</font>
				<br>
				<font color="red" face="<% =font1 %>" size="3"><p><b>You must be 18 years old.</b></font> 			
				<br><br>
				<font face="<% =font1 %>" size="2">If you did not intend to accept this WAIVER and RELEASE, or you are less than 18 years old, do not continue. </font>
				<font face="<% =font1 %>" size="2">A digital record of your acceptance will be created.<br></font>
				<br>
			    </td>
			  </tr>
			<tr>
			   <td align="center">

				<form action="/rankings/<%=RegFileName%>?nav=5&waivernav=5" method="post">

					<INPUT type="hidden" NAME="sReleaseType" VALUE="<%=sReleaseType%>" >
					<INPUT type="hidden" NAME="sSignWaiver" VALUE="<%=sSignWaiver%>" >

					<center><input type="submit" value=" Continue "></center>
				</form>
			   </td>	
			   <td align="center">
				<form action="/rankings/<%=RegFileName%>?nav=5&waivernav=2" method="post">
				  <center><input type="submit" value="Previous"></center>
				</form>
			   </td>	
			</tr>
		 	</TABLE>   
		   </TD>		
		</TR>
	</TABLE>
	<br>   
	<% 


	CASE 10
				' --- Sends navigation to the next tab ---
				response.redirect("/rankings/"&RegFileName&"?nav=6")


	END SELECT

END SUB


' ---------------------------
  SUB SaveWaiverValuesToTable
' ---------------------------

		' --- SUB is NOT located in this RegFormDisplay.asp ---
		DefineWaiverSpecs

		' --- SUB is Located in RegFileName (updates Session variables) ---
		CheckWaiverStatus


		' ---- Read WhichTable for existing record ----
		SET rs=Server.CreateObject("ADODB.recordset")
		sSQL = "SELECT * FROM "&RegGenTableName
		sSQL = sSQL + " WHERE Left(TourID,6) = '" & SQLClean(left(sTourID,6)) & "' AND MemberID = '"&sMemberID&"'"
		rs.open sSQL, SConnectionToTRATable, 3, 3


		' -----------------------------------------------------------------------------------------
		' -----  Stores Waiver Code in RegTempTable then branches to complete the transaction -----
		' -----------------------------------------------------------------------------------------
		OpenCon
		sSQL = "UPDATE "&RegTempTableName
		sSQL = sSQL + " SET WaiverCode = '"&sWaiverCode&"', SignWaiver = '"&SQLClean(sSignWaiver)&"'"
		sSQL = sSQL + " WHERE Left(TourID,6) = '"&left(sTourID,6)&"' AND MemberID = '"&sMemberID&"'"
		con.execute(sSQL)
		closecon	

		Session("sRelease")=sWaiverCode


		' ---- Sends email to user and competition services  ----
		IF sReleaseType = "Electronic" THEN
				' --- SUB Located in Registration.asp --- 
				SendWaiverEmail
				
				' --- If the tournament has its own waiver to sign ---
				IF TRIM(sSpecialWaiverCode)<>"" THEN
						' --- SUB Located in Registration.asp ---
						SendSPECIALWaiverEmail sSpecialWaiverCode, sSpecialWaiverHeadline, sSpecialReleaseBannerText
				END IF
		END IF

END SUB






' ----------------------
  SUB DefineWaiverSpecs
' ----------------------


	' ---------  Displays RELEASE & WAIVER   --------	
	IF sReleaseType = "Electronic" AND Session("sMembAge") < 18 THEN
		' --- Value of minor_waiver set at the top of this page
		sWaiverCode = minor_waiver
		sWaiverSubTitle="Waiver for MINOR Participant - WaiverID: "&minor_waiver

	ELSEIF  sReleaseType = "Electronic" AND Session("sMembAge") >= 18 THEN
		' --- Value of adult_waiver set at the top of this page
		sWaiverCode = adult_waiver	
		sSignWaiver = SQLClean(LEFT(sFirstName,12))&" "&SQLClean(LEFT(sLastName,18))
		sWaiverSubTitle="Waiver for ADULT Participant - WaiverID: "&adult_waiver

	ELSEIF sReleaseType = "Paper" AND Session("sMembAge") >= 18 THEN 
		sWaiverCode = "Paper"	
		sSignWaiver = SQLClean(LEFT(sFirstName,12))&" "&SQLClean(LEFT(sLastName,18))
		sWaiverSubTitle="Waiver for ADULT Participant - WaiverID: "&adult_waiver

	ELSEIF sReleaseType = "Paper" AND Session("sMembAge") < 18 THEN 
		sWaiverCode = "Paper"	
		sSignWaiver = SQLClean(LEFT(sFirstName,12))&" "&SQLClean(LEFT(sLastName,18))
		sWaiverSubTitle=" Waiver for MINOR Participant Required "

	ELSEIF sReleaseType = "None" AND Session("sMembAge") >= 18 THEN 
		sWaiverCode = "None"	
		sSignWaiver = ""
		sWaiverSubTitle=" Waiver for ADULT Participant Required "

	ELSEIF sReleaseType = "None" AND Session("sMembAge") < 18 THEN 
		sWaiverCode = "None"	
		sSignWaiver = ""
		sWaiverSubTitle=" Waiver for MINOR Participant Required "

	ELSE		
		sWaiverCode = "None"	
		sSignWaiver = ""
		sWaiverSubTitle=" Paper Waiver for ADULT Participant Required "
	END IF

END SUB


Function CheckClass(ThisClass, WhichField)
	CheckClass=""
	IF WhichField=ThisClass THEN CheckClass="checked"
'response.write("sClass1="&sClass1)
'response.write("ThisClass="&ThisClass)
'response.write("WhichField="&WhichField)

End Function	

		

%>

