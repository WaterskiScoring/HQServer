<!--#include virtual="/admin/includes/security.asp" -->

<% If not Session("aauth") then response.redirect "Login.asp" %>

<html>

<head>
    <title>USA Water Ski Admin Index</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" >
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
            <td bgcolor="#42639F">
                <p>&nbsp;</p>
                <p align="center"><font face="Verdana" size="6" color="#FFFFFF">
                    USA Water Ski Admin Index</font></p>
                <p align="center"><font face="Verdana" size="4" color="#FFFFFF">
                    Registration Support for -- <%=session("TournamentName")%></font></p>
                <p>&nbsp;</p>
            </td>
        </tr>
    </table>

    <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
            <td width="185" bgcolor="#42639F" valign="top">

                <!--#include virtual="/admin/includes/menu.asp" -->
            </td>

            <td valign="top" >
        	    <table border="0" cellspacing="1" cellpadding="1">
                    <tr>
                      <td>&nbsp;&nbsp;&nbsp;</td>
                      <td>&nbsp;</td>
                      <td>&nbsp;&nbsp;&nbsp;</td>
                    </tr>

                    <tr>
                        <td>&nbsp;</td>
                        <td valign="top"> 
                            <% 
		 		            Dim UserRS
				            Set UserRS = Server.CreateObject("ADODB.RecordSet")
				            UserRS.ActiveConnection = objConn
				            UserRS.Open "SELECT * FROM Users999 where Name = '" & Session("UserName") & "'"
				            if not UserRS.EOF then
					            
                                if UserRS("CreateRegistrationTemplate") then
                                    %>
                                    <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><p>You can now 
                                    create your &#8220;Registration Template&#8221; (provided in Microsoft Excel 
                                    format) for your event.&nbsp; This Excel file contains a listing of members, 
                                    including their names, membership numbers, age divisions, home city and state, 
                                    officials ratings codes, ranking list scores and ranking levels, as well as 
                                    membership status information.&nbsp; Age Divisions and Membership Status 
                                    information will be presented relative to
            
                                    <%
                                    IF left(Session("UserName"),1) <= "9" THEN
                                    %>
                                        the applicable Ski Year and Tournament Start Date listed in SWIFT for the 
                                        specific event that you have just signed in for.</p>
                                    <%
                                    ELSE
                                    %>
                                        the applicable Ski Year and Tournament Start Date listed in SWIFT for a 
                                        specific Tournament.&nbsp; <b><i>As an Administrative User</i></b>, if you 
                                        will be pulling a Registration Template or Pre-Registration Export, or be
                                        looking up individual members, you will need to supply the Tournament 
                                        ID for the specific competition desired, on the setup page.</p>
            
					                <%
						            END IF
					                %>
                                    
                                    <p>
                                    Members may be selected geographically for up to five states, and foreign 
                                    and/or open/elite skiers may be selected as well.&nbsp; A section of detailed
                                    instructions will also be included in your Excel file, directing you on how 
                                    to use this template and the membership information included therein for the 
                                    registration and scoring set-up of your tournament.
                                    </p>
            
                                    <p>
                                    In addition to the geographic/status selection of skiers in a downloaded
                                    Excel file, this site also provides a facility through which you can look up 
                                    additional scattered single members, one at a time.&nbsp; The information 
                                    on those additional members can then be copied and pasted into a Registration 
                                    Template that you have downloaded earlier.
                                    </p>

                                    <p><font color="red"><b>A new feature added for 2010</b></font></p>
            
                                    <p>
                                    <b>Chief and Appointed Officials</b>.&nbsp; Chief and Appointed officials
                                    coded in the Sanction system for the selected tournament, will now be 
                                    included in your Excel file.&nbsp; Additional information can be found in
                                    the instructions section of your downloaded Excel file.
                                    </p>

					                <%
					                IF left(Session("UserName"),1) > "9" THEN
					                %>
						                <p>
                                        <b><i>Note: As an Administrative User</i></b>, I will show you detailed 
						                contact information -- address, phone numbers and email address -- for 
						                each member you look up.&nbsp; This additional information will appear 
						                below the primary member data that you would copy and paste into your 
						                registration template.&nbsp; Please respect the confidential nature of 
						                this added contact information.
						                </p>
					                <%
					                END IF
					                %>

						            </p>

                                    <p>
                                    To begin accessing these features, click on either the &quot;Create 
                                    Registration Template&quot; or &quot;Look Up Individual Members&quot; 
                                    links, that appears to the left.
                                    </p>

                                    </font><br>
                                    <%  
					            END IF
			                END IF

				            UserRS.close
				            set UserRS = nothing
		                    %>

                        </td>
   		                
                        <td>&nbsp;</td>
                    </tr>

                </table>

            </td>

      </tr>
    </table>
</body>

</html>





