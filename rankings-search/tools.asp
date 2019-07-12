<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include file="tools_include.asp"-->
<%



' -----------------------------------------------------------------------------------------------
' --------------  Commonly Used Subroutines that require branching str NOT in INCLUDE  ----------
' -----------------------------------------------------------------------------------------------

Dim svar
svar=TRIM(Request("svar"))
NewsPageNum=TRIM(Request("np"))
TextDisplayWidth=675

'svar="tourreglist"

PathToNews=Server.mappath("/")&"/rankings/news"

SELECT CASE svar
	CASE "FAQ"
		DisplayFAQ
		'response.write("PathToNews="&PathToNews&"/"&NewsPageNum&".txt")

	CASE "reject"
		AuthorityRejectionNotice

	CASE "tourreglist"
		TourOnlineSetup

	CASE "close"
		close()
END SELECT




' -----------------------------
  SUB AuthorityRejectionNotice
' -----------------------------

WriteIndexPageHeader
%>
<br>

	<TABLE BORDER="4" class="droptable" ALIGN="CENTER" CELLPADDING="0" CELLSPACING="0" BGCOLOR="<%=TableColor1%>" width=70% >
		  <TR>
		      <TD BGCOLOR="<%=HQSiteColor2%>" ALIGN="center"><font face=<% =font1 %> size="4"><b>ATTENTION</b></font></TD>
		  </TR>  
		  <TR>
		     <TD>
			<TABLE ALIGN="Center" BORDER="0" BGCOLOR="<% =tablecolor1 %>" CELLPADDING="6" CELLSPACING="3" width=100% BGCOLOR="#FFFFFF">
			<tr>
			   <td colspan="2" align="center">
			    <br> 
				<font face="<% =font1 %>" size="2"><b>Your LOGIN does not have authority in this Sports Discipline</b></font>
			    <br><br> 
				<font face="<% =font1 %>" size="2"><b>Sports Discipline ID: <%=Session("SptsGrpID")%></b></font>

			    <br> 
				<font face="<% =font1 %>" size="1">For questions, contact</font>
			    <br> 			
				<font face="<% =font1 %>" size="1">USA Water Ski Competition Dept. at 800-533-2972.</font>
			    <br><br> 
			  </td>
			</tr>
			<tr>
			   <td align="center">
				<form action="/rankings/defaultHQ.asp" method="post">
				  <input type="submit" value=" Continue "  title="Press Continue to return to main menu.">
				</form>
			   </td>	
			</tr>
			</TABLE>

		    </TD>
		  </TR>
			
		</TABLE><%

WriteIndexPageFooter
END SUB



' ------------------------------
   SUB DisplayFAQ
' ------------------------------

Dim currentline, TextDisplayWidth


'WriteIndexPageHeader


%>
<TABLE border=0 width=100%>
<tr>
  <td align="center">
        <font size="5" face=<% =font2 %> COlOR="<%=TextColor3%>">
	<b>Frequently Asked Questions<b>
	</font>
	<br>
  </td>
</tr>


<tr>
  <td align="left">


	<body>
        <font size=<% =fontsize2 %> face=<% =font2 %> COlOR="<%=TextColor1%>">
	<%

	  Set objfso = CreateObject("Scripting.FileSystemObject")
	  set objstream=objFSO.opentextfile(PathToNews&"/"&NewsPageNum&".txt")

	  IF NOT objstream.atendofstream THEN
  		DO WHILE not objstream.atendofstream
			currentline=objstream.readline
			response.write(currentline)
			response.write(chr(10))
		LOOP
	  END IF

	  objstream.close %>

	</font>
	</body>

  </td>
</tr>

</TABLE>

<%

'WriteIndexPageFooter

END SUB




SUB HoldItHere
%>

<tr>
  <td align=center>
    <br>
     <form action="/rankings/tools.asp?svar=close" method="post">
     	<input type="submit" value="Continue" >
   </form>
  </td>
</tr>



<%
END SUB


' --------------------
 SUB TourOnlineSetup
' --------------------


set rsTList=Server.CreateObject("ADODB.recordset")
sSQL = "SELECT * FROM sanctions.dbo.Registration AS TRT"
sSQL = sSQL + " JOIN sanctions.dbo.TSchedul AS TS ON TRT.TournAppID=TS.TournAppID"
sSQL = sSQL + " WHERE LEFT(TRT.TournAppID,2)='08'"
rsTList.open sSQL, sConnectionToTRATable, 3, 1

%>
<TABLE align="center" class="innertable"><%

DO WHILE NOT rsTList.eof  %>
  <TR>
    <TD><font fontsize="<%=fontsize1%>"><%=rsTList("TournAppID")%></font></TD>	
    <TD><font fontsize="<%=fontsize1%>"><%=rsTList("TDateS")%></font></TD>	
  </TR><%

	rsTList.movenext
LOOP %>

</TABLE><%




END SUB

%>









