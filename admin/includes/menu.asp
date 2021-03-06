<%

' response.redirect "../../rankings/CompetitionMaintSplashPage.htm": ' transfer to maintenance announcement page.

Dim objConn
Set objConn = Server.CreateObject("ADODB.Connection")
objConn.Open Application("WaterSkiConn")
%>
	<% 
    If Session("aauth") then 
        %>
	    <font face="Verdana" size="2" COLOR="#FFFFFF"><br>&nbsp;Currently Logged in as: </font>
        <br>
	    <font face="Verdana" size="2" COLOR="#FFFFFF">&nbsp;
            <%=Session("UserName")%>&nbsp;&nbsp;
            <%=session("TournamentDate")%></font>
            <br><br>
	<% 
    Else 
    %>
        <font face="Verdana" size="2" COLOR="#FFFFFF">
            <br>&nbsp;Not currently logged in.
            <br>&nbsp;<br>
	    </font>

	<% 
    End If 
	
    IF Session("aauth") THEN
        Dim TopUser, DownloadMembers1, DownloadDBF, FromUSAWS, SanIDType
        Dim CreateRegistrationTemplate, AdminUsers, EditDivisions
        Set TopUser = Server.CreateObject("ADODB.RecordSet")
        TopUser.ActiveConnection = objConn
        TopUser.Open "SELECT * FROM Users999 where UserID = " & Session("UserID")
        DownloadMembers1 = TopUser("DownloadMembers1")
        DownloadDBF = TopUser("DownloadDBF")
        CreateRegistrationTemplate = TopUser("CreateRegistrationTemplate")
        AllowAccess = TopUser("AllowAccess")
        AdminUsers = TopUser("AdminUsers")
        EditDivisions = TopUser("EditDivisions")
        FromUSAWS = TopUser("FromUSAWS")
        TopUser.Close
        Set TopUser = Nothing

        %>
        <font face="Verdana" size="2"> 
		    <% 
            IF DownloadMembers1 THEN 
            %>
                <a href="/admin/FindExpiresReEnrolling.asp"><FONT face="arial" COLOR="#FFFFFF">Expires Re-Enrolling Rept</font></a>
                <br>&nbsp;<br>
                <a href="/admin/FindPossDupMembers.asp"><FONT face="arial" COLOR="#FFFFFF">Possible Dup Member Rept</font></a>
                <br>&nbsp;<br>
		    <% 
            end if 
            %>

			<% 
            IF CreateRegistrationTemplate THEN
				
				IF Left(Session("UserName"),1) > "9" THEN 
                %>
                    <a href="/admin/CreateRegTemplateSetup.asp"> <FONT face="arial" COLOR="#FFFFFF">Create Registration Template</font></a>
                    <br>&nbsp;<br>
                    <a href="/admin/CreatePreRegTemplateSetup.asp"><FONT face="arial" COLOR="#FFFFFF">Create Pre-Reg Template</font></a>
                    <br>&nbsp;<br>

				<% 
                ELSE 
					' We have a Sanction ID as UserName -- Check to see if we have any online entries for this event
					' If not staged for OLR, then Separate Processes for NCWSA (3rd = "U") versus all others.

					Dim RegTable
					Set RegTable = Server.CreateObject("ADODB.RecordSet")
					RegTable.ActiveConnection = objConn
					RegTable.Open "SELECT count(*) as Entries FROM Cobra00025.USAWSRank.RegisterGen_05042014 where Left(TourID,6) = '" & Session("UserName") & "'"

					IF RegTable("Entries") > 0 THEN 
                    %>
  	      	            <a href="/admin/CreatePreRegTemplateSetup.asp"><FONT face="arial" COLOR="#FFFFFF">Create Pre-Registration Template</font></a><br>&nbsp;<br>
					<% 
                    ELSE 
					    IF Mid(Session("UserName"),3,1) = "U" THEN 
                        %>
                            <a href="/admin/CreateNCWSATemplate.asp"><FONT face="arial" COLOR="#FFFFFF">Create Registration Template</font></a><br>&nbsp;<br>
		      	            <% 
                            IF AllowAccess THEN 
                            %>
			      	            <a href="/admin/NCWSAChgRegStat.asp?TourID=<%=Session("UserName")%>&Status=Close"><FONT face="arial" COLOR="#FFFFFF">Close Online Registration</font></a><br>&nbsp;<br>
			      	        <% 
                            ELSE 
                            %>
			      	            <a href="/admin/NCWSAChgRegStat.asp?TourID=<%=Session("UserName")%>&Status=Open"><FONT face="arial" COLOR="#FFFFFF">Re-Open Online Registration</font></a><br>&nbsp;<br>
			      	        <% 
                            END IF 
                            %>
						
                        <% 
                        ELSE 
                        %>
		      	            <a href="/admin/CreateRegTemplateSetup.asp"><FONT face="arial" COLOR="#FFFFFF">Create Registration Template</font></a><br>&nbsp;<br>
						<% 
                        END IF 
					
                    END IF

					RegTable.Close
					Set RegTable = Nothing 
                    %>

                <% 
                END IF 
                %>

    	    	<a href="/admin/LookupMembers.asp"><FONT face="arial" COLOR="#FFFFFF">Look Up Individual Members</font></a><br>&nbsp;<br>

            <% END IF %>

			<% 
            IF AdminUsers THEN 
            %>
                <a href="/admin/useradmin.asp"><FONT face="arial" COLOR="#FFFFFF">Admin Users</font></a><br>&nbsp;<br>
			<% 
            END IF 
            %>

            <% 
            IF EditDivisions THEN 
            %>
                <a href="/admin/divisionsadmin.asp"><FONT face="arial" COLOR="#FFFFFF">Admin Divisions</font></a><br>&nbsp;<br>
            <% 
            END IF 
            %>

            <br><a href="/admin/logout.asp"><font face="arial" COLOR="#FFFFFF">Log Out</font></a><br>&nbsp;<br>
        </font>

    <% 
    Else 
    %>
        <br>
    <% 
    End If 
    %>

    <a href=/admin/index.asp><font face="arial" size="2" COLOR="#FFFFFF">Back to Admin Index</font></a><br>&nbsp;<br>

    <a href="http://usaws.usawaterski.org"><font face="arial" size="2" COLOR="#FFFFFF">Back to Online Sanctioning</font></a><br>&nbsp;<br>

    <a href="http://www.usawaterski.org"><font face="arial" size="2" COLOR="#FFFFFF">USA Water Ski Home</font></a><br>&nbsp;<br>

    <br>
    <font face="Verdana" size="1">&nbsp;<font COLOR="#FFFFFF">Powered by</font> 
        <a href="http://www.epolk.com"><font COLOR="#FFFFFF">ePolk.com</font></a>
    </font>
    
    <br>&nbsp;<br>
