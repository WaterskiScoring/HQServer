<!--#include virtual="/rankings/Eago_Settings.asp"-->
<!--#include virtual="/rankings/Eago_Tools.asp"-->


<%



DefineEagoStyles

Dim ThisPageName

Dim action
Dim sFromMemberID, sFromFirst, sFromLast, sFromCompany, sFromAddress1, sFromAddress2, sFromCity, sFromState, sFromZip
Dim sFromWorkPhone, sFromCellPhone, sFromEmail, sFromCategory
Dim sToMemberID, sToFirst, sToLast, sToCompany, sToAddress1, sToAddress2, sToCity, sToState, sToZip
Dim sToWorkPhone, sToCellPhone, sToEmail, sToCategory
Dim sLeadMemberID, sLeadFirst, sLeadLast, sLeadCompany, sLeadAddress1, sLeadAddress2, sLeadCity, sLeadState, sLeadZip
Dim sLeadWorkPhone, sLeadCellPhone, sLeadEmail
Dim sShowFromDetails, sShowToDetails

Dim sContactUrgency, sPurchaseUrgency, sLeadType, sBusinessType, sLeadReference, sLeadCategory

ThisPageName="/rankings/Eago_LeadForm.asp"

MemberDropStatus="Active"





' --- Begin Form Display

ReadFormData

IF sFromMemberID<>"" THEN
		GetMemberFromData
END IF

IF sToMemberID<>"" THEN
		GetMemberToData
END IF



IF action="Submit Lead" THEN
		SendLeadSubmitEMail
ELSE
		DisplayMainForm
END IF




' *********************************************************************
' ---     									END OF MAIN PROGRAM 										---
' *********************************************************************




' ----------------------
   SUB ReadFormData
' ----------------------

action=Request("action")


sToMemberID=TRIM(Request("sToMemberID"))
sToCompany=TRIM(Request("sToCompany"))

sFromMemberID=TRIM(Request("sFromMemberID"))
sFromCompany=TRIM(Request("sFromCompany"))

sLeadCompany=TRIM(Request("sLeadCompany"))
sLeadFirst=TRIM(Request("sLeadFirst"))
sLeadLast=TRIM(Request("sLeadLast"))
sLeadAddress1=TRIM(Request("sLeadAddress1"))
sLeadAddress1=TRIM(Request("sLeadAddress1"))
sLeadCity=TRIM(Request("sLeadCity"))
sLeadState=TRIM(Request("sLeadState"))
sLeadZip=TRIM(Request("sLeadZip"))
sLeadWorkPhone=TRIM(Request("sLeadWorkPhone"))
sLeadCellPhone=TRIM(Request("sLeadCellPhone"))
sLeadEmail=TRIM(Request("sLeadEmail"))

sContactUrgency=Request("sContactUrgency")
sPurchaseUrgency=Request("sPurchaseUrgency")
sBusinessType=Request("sBusinessType")
sLeadType=Request("sLeadType")	
sLeadReference=Request("sLeadReference")

sShowFromDetails=Request("sShowFromDetails")
sShowToDetails=Request("sShowToDetails")
sLeadCategory=Request("sLeadCategory")

' --- Define defaults 
IF action="" THEN action="Modify"
IF sLeadState="" THEN sLeadState="FL"


END SUB


' -----------------------
   SUB ValidateFormData
' -----------------------   

sLeadCompanyValid=true
sLeadFirstValid=true
sLeadPhoneValid=true

IF sLeadCompany="" THEN sLeadCompanyValid=false
IF sLeadFirstValid="" THEN sLeadFirstValid=false
IF sLeadWorkPhone="" AND sLeadCellPhone="" THEN sLeadPhoneValid=false


END SUB



' ------------------------
   SUB GetMemberFromData
' ------------------------
   
sSQL = "SELECT * FROM "&MembersTableName
sSQL = sSQL + " WHERE MemberID='"&sFromMemberID&"'"
   
SET rs=Server.CreateObject("ADODB.recordset")

'response.write("<br>sSQL="&sSQL)
'response.end

rs.open sSQL, SConnectionToTRATable

IF NOT rs.eof THEN
	sFromCompany=rs("Company")
	sFromFirst=rs("First")
	sFromLast=rs("Last")
	sFromAddress1=rs("Address1")
	sFromAddress2=rs("Address2")
	sFromCity=rs("City")
	sFromState=rs("State")
	sFromZip=rs("Zip")
	sFromWorkPhone=rs("WorkPhone")
	sFromCellPhone=rs("CellPhone")
	sFromEmail=rs("Email")
	sFromCategory=rs("Category")
END IF

END SUB





' ------------------------
   SUB GetMemberToData
' ------------------------
   
sSQL = "SELECT * FROM "&MembersTableName
sSQL = sSQL + " WHERE MemberID='"&sToMemberID&"'"
   
SET rs=Server.CreateObject("ADODB.recordset")

'response.write("<br>sSQL="&sSQL)
'response.end

rs.open sSQL, SConnectionToTRATable

IF NOT rs.eof THEN
	sToCompany=rs("Company")
	sToFirst=rs("First")
	sToLast=rs("Last")
	sToAddress1=rs("Address1")
	sToAddress2=rs("Address2")
	sToCity=rs("City")
	sToState=rs("State")
	sToZip=rs("Zip")
	sToWorkPhone=rs("WorkPhone")
	sToCellPhone=rs("CellPhone")
	sToEmail=rs("Email")
	sToCategory=rs("Category")
END IF


END SUB





' ***********************************
    SUB DisplayMainForm
' ***********************************

Heading1="Sender"
Heading2="Recipient "
Heading3="Lead Information"



ContactUrgencyDropStatus="Active"
PurchaseUrgencyDropStatus="Active"
LeadTypeDropStatus="Active"
StateListStatus="Active"
IF action="Confirm" THEN LeadTextFieldStatus="Disabled"


' --- Begin table design ---
	MemberSelected=sFromMemberID
	MemberDropName="sFromMemberID"

	%>
<body BGColor="<%=BackgroundColor1%>">

<TABLE align=center height=150px WIDTH=900px background=<%=HeaderImage1%>>
		<TR>
				<td>&nbsp;</td>
		</TR>		
</TABLE>

<form action="<%= ThisPageName %>" method="post">
	<TABLE class="innertable" Align=center WIDTH=<%=StdTableWidth%> >
		  <TR>
		    <th ALIGN="Left" colspan=8>
					<font size=4 color="<%=HeaderTextColor1%>"><b>&nbsp;<%=Heading1%></b></font>
					<right>
						<font color="<%=HeaderTextColor1%>" size=2><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Show Details</b></font>
						<input type=checkbox name="sShowFromDetails" <% IF sShowFromDetails="on" THEN response.write "checked" %> onclick=submit()>
					</right>
			 </th>
		  </TR>

			<TR>
      	<td Colspan=1  width=15% align=right >
					<font color="<%=TextColor1%>" size="2"><b>Company</b></font>
				</td>
				<td Colspan=3 width=35% align=left >
					<font color="<%=TextColor1%>" size="2"><%
						LoadMemberListDropDown MemberSelected, MemberDropName, MemberDropStatus
		  			%>
					</font>
				</td>
      	<td align=right width=15% >
					<font color="<%=TextColor1%>" size="2"><b>Category</b></font>
				</td>
				<td Colspan=3 align=left width=35% >
					<font color="<%=TextColor2%>" size="2"><% = sFromCategory %></font>
				</td>
		  </TR>

		<%
		IF sShowFromDetails="on" THEN	%>
		  <TR>
      	<td align=right>
					<font color="<%=TextColor1%>" size="2"><b>First</b></font>
				</td>
				<td Colspan=3 align=left>
					<font color="<%=TextColor1%>" size="2"><% = sFromFirst %></font>
				</td>
      	<td align=right>
					<font color="<%=TextColor1%>" size="2"><b>Last</b></font>
				</td>
				<td Colspan=3 align=left>
					<font color="<%=TextColor1%>" size="2"><% = sFromLast %></font>
				</td>
		  </TR>

		  <TR>
      	<td align=right>
					<font color="<%=TextColor1%>" size="2"><b>Address1</b></font>
				</td>
				<td Colspan=3 align=left>
					<font color="<%=TextColor1%>" size="2"><% = sFromAddress1 %></font>
				</td>
      	<td align=right>
					<font color="<%=TextColor1%>" size="2"><b>&nbsp;</b></font>
				</td>
				<td Colspan=3 align=left>
					<font color="<%=TextColor1%>" size="2">&nbsp;</font>
				</td>
		  </TR>

		  <TR>
      	<td align=right>
					<font color="<%=TextColor1%>" size="2"><b>Address2</b></font>
				</td>
				<td Colspan=3 align=left>
					<font color="<%=TextColor1%>" size="2"><% = sFromAddress2 %></font>
				</td>
      	<td align=right>
					<font color="<%=TextColor1%>" size="2"><b>&nbsp;</b></font>
				</td>
				<td Colspan=3 align=left>
					<font color="<%=TextColor1%>" size="2">&nbsp;</font>
				</td>
		  </TR>

		  <TR>
      	<td align=right >
					<font color="<%=TextColor1%>" size="2"><b>City</b></font>
				</td>
				<td Colspan=3 align=left>
					<font color="<%=TextColor1%>" size="2"><% = sFromCity %></font>
				</td>
      	<td align=right>
					<font color="<%=TextColor1%>" size="2"><b>State</b></font>
				</td>
				<td Colspan=1 align=left >
					<font color="<%=TextColor1%>" size="2"><% = sFromState %></font>
				</td>
				<td Colspan=1 align=right >
					<font color="<%=TextColor1%>" size="2"><b>Zip</b></font>
				</td>
				<td Colspan=1 align=left >
					<font color="<%=TextColor1%>" size="2"><% = sFromZip %></font>
				</td>
		  </TR>
		  <TR>
      	<td align=right>
					<font color="<%=TextColor1%>" size="2"><b>Work Phone</b></font>
				</td>
				<td Colspan=3 align=left>
					<font color="<%=TextColor1%>" size="2"><% = sFromWorkPhone %></font>
				</td>
      	<td align=right>
					<font color="<%=TextColor1%>" size="2"><b>Cell Phone</b></font>
				</td>
				<td Colspan=3 align=left >
					<font color="<%=TextColor1%>" size="2"><% = sFromCellPhone %></font>
				</td>
		  </TR>
		  <TR>
      	<td align=right>
					<font color="<%=TextColor1%>" size="2"><b>EMail</b></font>
				</td>
				<td Colspan=3 align=left>
					<font color="<%=TextColor1%>" size="2"><% = sFromEmail %></font>
				</td>
				<td colspan=4>&nbsp;</td>
		  </TR><%

		END IF		' --- Bottom of Show Details
		%>
		
 	</TABLE>
	<%


	MemberSelected=sToMemberID
	MemberDropName="sToMemberID"

	%>
  <br> 
	<TABLE class="innertable" Align=center WIDTH=<%=StdTableWidth%> >
		  <TR>
		    <th ALIGN="left" colspan=8>
					<font size=4 color="<%=HeaderTextColor1%>"><b>&nbsp;<%=Heading2%></b></font>
					<right>
						<font color="<%=HeaderTextColor1%>" size=2><b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Show Details</b></font>
						<input type=checkbox name="sShowToDetails" <% IF sShowToDetails="on" THEN response.write "checked" %> onclick=submit()>
					</right>
		    </th>
		  </TR>
			<TR>
      	<td Colspan=1  width=15% align=right >
					<font color="<%=TextColor1%>" size="2"><b>Company</b></font>
				</td>
				<td Colspan=3 width=35% align=left ><%
						LoadMemberListDropDown MemberSelected, MemberDropName, MemberDropStatus
		  			%>
				</td>
      	<td align=right width=15% >
					<font color="<%=TextColor1%>" size="2"><b>Category</b></font>
				</td>
				<td Colspan=3 align=left width=35% >
					<font color="<%=TextColor2%>" size="2"><% = sToCategory %></font>
				</td>
		  </TR>

		<%
		IF sShowToDetails="on" THEN	%>
			<TR>
      	<td align=right>
					<font color="<%=TextColor1%>" size="2"><b>First</b></font>
				</td>
				<td Colspan=3 align=left>
					<font color="<%=TextColor2%>" size="2"><% = sToFirst %></font>
				</td>
      	<td align=right width=15%>
					<font color="<%=TextColor1%>" size="2"><b>Last</b></font>
				</td>
				<td Colspan=3 align=left>
					<font color="<%=TextColor2%>" size="2"><% = sToLast %></font>
				</td>
		  </TR>
		  <TR>
      	<td align=right>
					<font color="<%=TextColor1%>" size="2"><b>Address1</b></font>
				</td>
				<td Colspan=3 align=left>
					<font color="<%=TextColor2%>" size="2"><% = sToAddress1 %></font>
				</td>
      	<td align=right width=15%>
					<font color="<%=TextColor1%>" size="2"><b>&nbsp;</b></font>
				</td>
				<td Colspan=3 align=left>
					<font color="<%=TextColor2%>" size="2">&nbsp;</font>
				</td>
		  </TR>
		  <TR>
      	<td align=right>
					<font color="<%=TextColor1%>" size="2"><b>Address2</b></font>
				</td>
				<td Colspan=3 align=left>
					<font color="<%=TextColor2%>" size="2"><% = sToAddress2 %></font>
				</td>
      	<td align=right width=15%>
					<font color="<%=TextColor1%>" size="2"><b>&nbsp;</b></font>
				</td>
				<td Colspan=3 align=left>
					<font color="<%=TextColor2%>" size="2">&nbsp;</font>
				</td>
		  </TR>
		  <TR>
      	<td align=right>
					<font color="<%=TextColor1%>" size="2"><b>City</b></font>
				</td>
				<td Colspan=3 align=left width=35% >
					<font color="<%=TextColor2%>" size="2"><% = sToCity %></font>
				</td>
      	<td align=right width=15%>
					<font color="<%=TextColor1%>" size="2"><b>State</b></font>
				</td>
				<td Colspan=1 align=left width=10% >
					<font color="<%=TextColor2%>" size="2"><% = sToState %></font>
				</td>
				<td Colspan=1 align=right >
					<font color="<%=TextColor1%>" size="2"><b>Zip</b></font>
				</td>
				<td Colspan=1 align=left>
					<font color="<%=TextColor2%>" size="2"><% = sToZip %></font>
				</td>
			</TR>
			<TR>
      	<td align=right>
					<font color="<%=TextColor1%>" size="2"><b>Work Phone</b></font>
				</td>
				<td Colspan=3 align=left>
					<font color="<%=TextColor2%>" size="2"><% = sToWorkPhone %></font>
				</td>
      	<td align=right>
					<font color="<%=TextColor1%>" size="2"><b>Cell Phone</b></font>
				</td>
				<td Colspan=3 align=left >
					<font color="<%=TextColor2%>" size="2"><% = sToCellPhone %></font>
				</td>
		  </TR>
		  <TR>
      	<td align=right>
					<font color="<%=TextColor1%>" size="2"><b>EMail</b></font>
				</td>
				<td Colspan=3 align=left>
					<font color="<%=TextColor2%>" size="2"><% = sToEmail %></font>
				</td>
				<td colspan=4>&nbsp;</td>
		  </TR><%
		
		END IF 		' --- Bottom of ShowToDetails 
		%> 
	</TABLE>



  <br> 
	<TABLE class="innertable" Align=center WIDTH=<%=StdTableWidth%> >
		  <TR>
		    <th ALIGN="left" colspan=8>
					<font size=4 color="<%=HeaderTextColor1%>"><b>&nbsp;<%=Heading3%></b></font>
		    </th>
		  </TR>
			<TR>
      	<td align=right width=15%>
					<font color="<%=TextColor1%>" size="2"><b>Company</b></font>
				</td>
				<td colspan=3 align=left width=35%><%
					IF action="Confirm" THEN %>
						<font color="<%=TextColor2%>" size="2"><%=sLeadCompany%></font>
						<input type=hidden name="sLeadCompany" value="<%=sLeadCompany%>" ><%
					ELSE %>
						<input type="Text" name="sLeadCompany" value="<%=sLeadCompany%>" MaxLength=30 size="30"><%
					END IF %>
				</td>
				<td colspan=4 width=50%>&nbsp;</td>
		  </TR>
			<TR>
      	<td align=right>
					<font color="<%=TextColor1%>" size="2"><b>First</b></font>
				</td>
      	<td Colspan=3 align=left width=35%><%
					IF action="Confirm" THEN %>
						<font color="<%=TextColor2%>" size="2"><%=sLeadFirst%></font>
						<input type=hidden name="sLeadFirst" value="<%=sLeadFirst%>" ><%
					ELSE %>
						<input type="Text" name="sLeadFirst" value="<%=sLeadFirst%>" MaxLength=30 size="30"><%
					END IF %>
				</td>
      	<td align=right>
					<font color="<%=TextColor1%>" size="2"><b>Last</b></font>
				</td>
				<td Colspan=3 align=left width=35% ><%
					IF action="Confirm" THEN %>
						<font color="<%=TextColor2%>" size="2"><%=sLeadLast%></font>
						<input type=hidden name="sLeadLast" value="<%=sLeadLast%>" ><%
					ELSE %>
						<input type="Text" name="sLeadLast" value="<%=sLeadLast%>" MaxLength=30 size="30"><%
					END IF %>
				</td>
		  </TR>
		  <TR>
      	<td align=right>
					<font color="<%=TextColor1%>" size="2"><b>Address1</b></font>
				</td>
				<td Colspan=3 align=left width=35% ><%
					IF action="Confirm" THEN %>
						<font color="<%=TextColor2%>" size="2"><%=sLeadAddress1%></font>
						<input type=hidden name="sLeadAddress1" value="<%=sLeadAddress1%>" ><%
					ELSE %>
						<input type="Text" name="sLeadAddress1" value="<%=sLeadAddress1%>" MaxLength=30 size="30"><%
					END IF %>
				</td>
      	<td align=right width=15%>
					<font color="<%=TextColor1%>" size="2"><b>&nbsp;</b></font>
				</td>
				<td Colspan=3 align=left width=35% >
					<font color="<%=TextColor1%>" size="2">&nbsp;</font>
				</td>
		  </TR>

		  <TR>
      	<td align=right>
					<font color="<%=TextColor1%>" size="2"><b>Address2</b></font>
				</td>
				<td Colspan=3 align=left width=35% ><%
					IF action="Confirm" THEN %>
						<font color="<%=TextColor2%>" size="2"><%=sLeadAddress2%></font>
						<input type=hidden name="sLeadAddress2" value="<%=sLeadAddress2%>" ><%
					ELSE %>
						<input type="Text" name="sLeadAddress2" value="<%=sLeadAddress2%>" MaxLength=30 size="30"><%
					END IF %>
				</td>
      	<td align=right width=15%>
					<font color="<%=TextColor1%>" size="2"><b>&nbsp;</b></font>
				</td>
				<td Colspan=3 align=left width=35% >
					<font color="<%=TextColor1%>" size="2">&nbsp;</font>
				</td>
		  </TR>

		  <TR>
      	<td align=right>
					<font color="<%=TextColor1%>" size="2"><b>City</b></font>
				</td>
				<td Colspan=3 align=left width=35% ><%
					IF action="Confirm" THEN %>
						<font color="<%=TextColor2%>" size="2"><%=sLeadCity%></font>
						<input type=hidden name="sLeadCity" value="<%=sLeadCity%>" ><%
					ELSE %>
						<input type="Text" name="sLeadCity" value="<%=sLeadCity%>" MaxLength=20 size="20"><%
					END IF %>
				</td>
      	<td align=right width=15%>
					<font color="<%=TextColor1%>" size="2"><b>State</b></font>
				</td>
				<td Colspan=1 align=left><%
					IF action="Confirm" THEN %>
						<font color="<%=TextColor2%>" size="2"><%=sLeadState%></font>
						<input type=hidden name="sLeadState" value="<%=sLeadState%>" ><%
					ELSE 
						BuildStateDropDown StateListStatus  
					END IF %>
				</td>
				<td Colspan=1 align=right>
					<font color="<%=TextColor1%>" size="2"><b>Zip</b></font>
				</td>
				<td Colspan=1 align=left><%
					IF action="Confirm" THEN %>
						<font color="<%=TextColor2%>" size="2"><%=sLeadZip%></font>
						<input type=hidden name="sLeadZip" value="<%=sLeadZip%>" ><%
					ELSE %>
						<input type="Text" name="sLeadZip" value="<%=sLeadZip%>" MaxLength=5 size=6><%
					END IF %>
				</td>
			</TR>	
			<TR>
      	<td align=right>
					<font color="<%=TextColor1%>" size="2"><b>Work Phone</b></font>
				</td>
				<td Colspan=3 align=left><%
					IF action="Confirm" THEN %>
						<font color="<%=TextColor2%>" size="2"><%=sLeadWorkPhone%></font>
						<input type=hidden name="sLeadWorkPhone" value="<%=sLeadWorkPhone%>" ><%
					ELSE %>
						<input type="Text" name="sLeadWorkPhone" value="<%=sLeadWorkPhone%>" MaxLength=12 size=12><%
					END IF %>
				</td>
      	<td align=right>
					<font color="<%=TextColor1%>" size="2"><b>Cell Phone</b></font>
				</td>
				<td Colspan=3 align=left ><%
					IF action="Confirm" THEN %>
						<font color="<%=TextColor2%>" size="2"><%=sLeadCellPhone%></font>
						<input type=hidden name="sLeadCellPhone" value="<%=sLeadCellPhone%>" ><%
					ELSE %>
						<input type="Text" name="sLeadCellPhone" value="<%=sLeadCellPhone%>" MaxLength=12 size=12><%
					END IF %>
				</td>
		  </TR>
		  <TR>
      	<td align=right>
					<font color="<%=TextColor1%>" size="2"><b>EMail</b></font>
				</td>
				<td Colspan=3 align=left><%
					IF action="Confirm" THEN %>
						<font color="<%=TextColor2%>" size="2"><%=sLeadEmail%></font>
						<input type=hidden name="sLeadEmail" value="<%=sLeadEmail%>" ><%
					ELSE %>
						<input type="Text" name="sLeadEmail" value="<%=sLeadEmail%>" MaxLength=50 size="50"><%
					END IF %>
				</td>
      	<td align=right>
					<font color="<%=TextColor1%>" size="2"><b>Category</b></font>
				</td>
				<td colspan=3><%
					IF action="Confirm" THEN %>
						<font color="<%=TextColor2%>" size="2"><%=sLeadCategory%></font>
						<input type=hidden name="sLeadCategory" value="<%=sLeadCategory%>" ><%
					ELSE 
						LoadLeadCategoryDropDown LeadCategoryStatus 
					END IF %>
				</td>
		  </TR>
			<TR>
      	<td align=right >
					<font color="<%=TextColor1%>" size="2"><b>Contact Urgency</b></font>
				</td>
				<td colspan=3><%
					IF action="Confirm" THEN %>
						<font color="<%=TextColor2%>" size="2"><%=sContactUrgency%></font>
						<input type=hidden name="sContactUrgency" value="<%=sContactUrgency%>" ><%
					ELSE 
						LoadContactUrgencyDropDown ContactUrgencyDropStatus
					END IF %>
				</td>
      	<td align=right >
					<font color="<%=TextColor1%>" size="2"><b>Business Type</b></font>
				</td>
				<td colspan=3><%
					IF action="Confirm" THEN %>
						<font color="<%=TextColor2%>" size="2"><%=sBusinessType%></font>
						<input type=hidden name="sBusinessType" value="<%=sBusinessType%>" ><%
					ELSE %>
						<font color="<%=TextColor1%>" size="2">
    					<input type="radio" name="sBusinessType" value="New" <% IF sBusinessType="New" THEN response.write("Checked") %> <% IF action="Confirm" THEN response.write("Disabled") %>>New&nbsp;&nbsp;&nbsp;&nbsp;
    					<input type="radio" name="sBusinessType" value="Repeat" <% IF sBusinessType="Repeat" THEN response.write("Checked") %> <% IF action="Confirm" THEN response.write("Disabled") %>>Repeat&nbsp;
						</font><%
					END IF %>
				</td>
		  </TR>

			<TR>
      	<td align=right >
					<font color="<%=TextColor1%>" size="2"><b>Purchase Urgency</b></font>
				</td>
				<td colspan=3><%
					IF action="Confirm" THEN %>
						<font color="<%=TextColor2%>" size="2"><%=sPurchaseUrgency%></font>
						<input type=hidden name="sPurchaseUrgency" value="<%=sPurchaseUrgency%>" ><%
					ELSE 
						LoadPurchaseUrgencyDropDown LeadTextFieldStatus
					END IF %>
				</td>
      	<td align=right >
					<font color="<%=TextColor1%>" size="2"><b>Lead Type</b></font>
				</td>
				<td colspan=3><%
					IF action="Confirm" THEN %>
						<font color="<%=TextColor2%>" size="2"><%=sLeadType%></font>
						<input type=hidden name="sLeadType" value="<%=sLeadType%>" ><%
					ELSE %>
						<font color="<%=TextColor1%>" size="2">
	    				<input type="radio" name="sLeadType" value="General" <% IF sLeadType="General" THEN response.write("Checked") %>>General&nbsp;
  	  				<input type="radio" name="sLeadType" value="Direct" <% IF sLeadType="Direct" THEN response.write("Checked") %>>Direct&nbsp;
    					<input type="radio" name="sLeadType" value="Confidential" <% IF sLeadType="Confidential" THEN response.write("Checked") %>>Confidential&nbsp;
    					<input type="radio" name="sLeadType" value="Feedback" <% IF sLeadType="Feedback" THEN response.write("Checked") %>>Feedback&nbsp;
						</font><%
					END IF %>
				</td>
		  </TR>
			<TR>
      	<td align=right >
					<font color="<%=TextColor1%>" size="2"><b>Use My Name (reference)</b></font>
				</td>
				<td colspan=3 align=left ><%
					IF action="Confirm" THEN %>
						<font color="<%=TextColor2%>" size="2"><%
							IF sLeadReference="N" THEN Response.write("No") ELSE Response.write("Yes") END IF %>
						</font>
						<input type=hidden name="sLeadReference" value="<%=sLeadReference%>" ><%
					ELSE %>
						<font color="<%=TextColor1%>" size="2">
    				<input type="radio" name="sLeadReference" value="Y" <% IF sLeadReference="Y" THEN response.write("Checked") %>>Yes&nbsp;
    				<input type="radio" name="sLeadReference" value="N" <% IF sLeadReference="N" THEN response.write("Checked") %>>No&nbsp;
						</font><%
					END IF %>
				</td>
		  	<td colspan=4 >&nbsp;</td>
		  </TR>


			<TR>
      	<td colspan=8 >&nbsp;</td>
		  </TR>

			<TR>
      	<td colspan=4 align=center width=50%><%
					IF action="Modify" THEN %>
						<input type="submit" name="action" value="Confirm" style="width:9em"><%
					ELSE %>
						<input type="submit" name="action" value="Modify" style="width:9em"><%
					END IF %>	
				</td>
      	<td colspan=4 align=center width=50%><%
					IF action="Modify" THEN %>
						<input type="submit" name="action" value="Submit Lead" style="width:9em" Disabled ><%
					ELSE %>
						<input type="submit" name="action" value="Submit Lead" style="width:9em"><%
					END IF %>	
				</td>

		  </TR>

	</TABLE>
	<br><br><br><br><br><br>

</form>
</body>
<%



END SUB







' ---------------------
  SUB DisplayResult
' ---------------------


IF NOT rs.eof THEN

	rs.movefirst

	' ---------------  Displays table HEADINGS  ----------------------

	%>
	<BR>
	<TABLE class="innertable" Align=center WIDTH=1020px ><%

		IF SubProcess="usdetail" THEN 
			SubProcessHead = "US Team Score and Individual Overall Detail - TeamX No: "&sTeamNo	
		ELSEIF SubProcess="intdetail" THEN 
			SubProcessHead="Country: "&FedSelected&" - Team & Individual Overall Scores - TeamX No: "&sTeamNo
		END IF
		IF SubProcess="usdetail" OR SubProcess="intdetail" THEN %>	
		  <TR>
		    <th ALIGN="Center" colspan=17>
			<font size=2 color="#FFFFFF"><b><%=SubProcessHead%></b></font>
		    </th>
		  </TR><%
		END IF %>

	  <TR><%
		FOR i = 0 TO rs.fields.count - 1
			TempFN = rs.fields(i).name
			j = 0 

			IF Session("AdminMenuLevel")>=50 AND process="ustrialssummary" AND i=0 THEN
				%>
				<Th ALIGN="center"><FONT COlOR="#FFFFFF" SIZE="<%=fontsize1%>">Add</FONT></Th>
				<Th ALIGN="center"><FONT COlOR="#FFFFFF" SIZE="<%=fontsize1%>">Remove</FONT></Th>
				<%
			END IF
			IF LEFT(rs.fields(i).name,2) <> "MK" AND (NOT (sOverlay<>"" AND rs.fields(i).name="Marked")) THEN %>
		   		<th ALIGN="Center" vAlign="top" nowrap>
				  <FONT COlOR="#FFFFFF" SIZE="<%=fontsize1%>"><%=Rs.Fields(i).name%></FONT>
				</th><%

			END IF
		NEXT %>
	  </TR><%


'response.write("Point A")
'response.end

Dim sColorSelected, sOverlayON
sColorSelected="red"
'sOverlayON = "yes"


	' --------------  Display table data here with paging --------------------------

	DO WHILE NOT rs.eof

		%>

 		<TR><%

		AllowEdit=true

		IF Session("AdminMenuLevel")>=50 AND process="ustrialssummary" THEN
			%>
			<TD ALIGN="center" vAlign="top"><FONT SIZE="<%=fontsize1%>"><% WriteLink "?Whataction=addtoover&sMemberID="&rs.fields(3).Value&"&process=ustrialssummary&sOverlay="&sOverlay&" ","Add","" %></FONT></TD>
			<TD ALIGN="center" vAlign="top"><FONT SIZE="<%=fontsize1%>"><% WriteLink "?Whataction=delfromover&sMemberID="&rs.fields(3).Value&"&process=ustrialssummary&sOverlay="&sOverlay&" ","Remove","" %></FONT></TD>
			<%
		END IF

		FOR i = 0 TO rs.fields.count - 1
	
			' --- Test for Athlete MemberID
			RowColor=""
			IF sOverlay<>"" AND Rs.Fields(i).value = "Y" THEN
				' --- Change the background to GREEN if this member is marked
				i = i + 1
				RowColor="background-color:"&tcolor03&";"
			ELSEIF sOverlay<>"" AND Rs.Fields(i).value = "N" THEN
				i = i + 1
			END IF

			TempFN = rs.fields(i).name  


			IF LEFT(Rs.Fields(i).name,2)<>"MK" THEN
				' --- Do nothing this is a flag  %>
		
			<TD ALIGN="center" style="<%=RowColor%>">
			  <FONT SIZE="1">&nbsp;<%
		
				' --- Displays link on TTNo to the team detail page for that TTNo ---
				IF process="teamlist" AND TempFN="TTNo" THEN  %>
					<a href="<%=ThisFileName%>?process=teamdetail&sTeamNo=<%=rs.fields(i).value%>&sTourID=<%=sTourID%>"><%=rs.fields(i).value%></a><%

				' --- Displays link on MemberID to the score detail for that member ---
				ELSEIF SubProcess="usdetail" AND i=0 THEN %>
					<a href="<%=ThisFileName%>?process=indivscores&sTourID=<%=sTourID%>&sMemberID=<%=rs.fields(i).value%>"><%=rs.fields(i).value%></a><%

				ELSE
					' --- Displays the fields in query other than TTNo
					SELECT CASE Rs.Fields(i).type
						CASE 3 'numeric'
							Response.Write(Rs.Fields(i).value)
						CASE 4  'numeric'
							'response.end
							Response.Write(formatnumber(Rs.Fields(i).value,2))
						CASE 5  'numeric'
							Response.Write(formatnumber(Rs.Fields(i).value,2))
	
						CASE 200  'char'
				
							Response.Write(LEFT(Rs.Fields(i).value,15))

						CASE 131 'numeric'
							Response.Write(formatnumber(Rs.Fields(i).value,2))

						CASE ELSE 'not handled by this function'
							'response.write("Y")
			        			Response.Write(Rs.Fields(i).value)
					END SELECT
				END IF  %>	
			  </FONT>
			</TD><%
			END IF
			


		NEXT
'response.end

		%>

		</TR><% 
		rowCount = rowCount + 1
		rs.movenext
	LOOP %>

	</TABLE>
<br><%

END IF


END SUB



%>








