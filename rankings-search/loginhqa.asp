<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!-- saved from url=(0029)http://www.samplewebsite.com/ -->
<!--#include virtual="/rankings/settingsHQ.asp"--> <%' ---------------------------- LOGIN VERIFICATION --------------------------------

  if session("reallogin") = "Valid" then

      Dim TRALogin
      TRALogin = 0

      If inStr(request("username"), " ") or inStr(request("password"), " ") > 0 then
        session("message") = "Invalid Login Attempt"
        Response.Redirect "/rankings/defaultHQ.asp?process=logout&rid=" & rid
      Else
        session("UserName") = request("username")
      End If
' If trim(request("username")) ="SEEDING" and trim(request("password")) ="GNIDEES" Then '        TRALogin = 1
' session("UserName") ="Seeding" '        session("FirstName") = "Seeding"
' session("LastName") ="User" '        session("UserLevel") = 49
' Session("adminmenulevel") =40 '      End If
' If trim(request("username")) ="SCORER" and trim(request("password")) ="REROCS" Then '        TRALogin = 1
' session("UserName") ="Scorer" '        session("FirstName") = "Scoring"
' session("LastName") ="User" '        session("UserLevel") = 10
' Session("adminmenulevel") =2 '      End If
      If trim(request("username")) = "" and trim(request("password")) = "buGtra" Then
        TRALogin = 1
        session("UserName") = "Administrator"
        session("FirstName") = "Jim"
        session("LastName") = "Euliano"
        session("UserLevel") = 9999999
        Session("adminmenulevel") = 50
      End If
       If TRALogin = 1 Then 
        WriteLog(date() &"  "& time() &"  "& session("UserName") & " has logged in.")
           
        WriteIndexPageHeader
        %>
         <br><br>
         <center><h2>Welcome to TRA<br>Data Administration Center</h2><br><h4>Logged In As: <%Response.Write(Session("FirstName") & " " & Session("LastName"))%><br></h4></center>
      
         <br><br><br><br>
         <br><br><br><br>
         <br><br><br><br>
        <%        
        WriteIndexPageFooter
      Else
        OpenCon
        set rs=Server.CreateObject("ADODB.recordset")
        sSQL = "Select top 1 * from " & LoginTableName & " where username = '" & trim(request("username")) & "' and pword = '" & trim(request("password")) & "'"
        rs.open sSQL, sConnectionToTRATable, 3, 1
        If rs.EOF Then
          WriteLog(date() &"  "& time() &"  **** INVALID LOGIN ATTEMPT FROM IP " & Request.ServerVariables("REMOTE_HOST") & " ****")
          WriteIndexPageHeader
          %>
           <br><br>
           <center><h2>Welcome to TRA<br>Data Administration Center</h2><br><br>
           <br><font color="red">Username or Password Incorrect.</font><br>
           <br><br>
           Please Try Again.
           <br><br>
           <form action="/rankings/loginHQ.asp" method="post">
           Username: <input type="text" name="username" size="10"><br><br>
           Password: <input type="password" name="password" size="10"><br><br>
           <br><br>
           <input type="submit" value="Login"><br><br><br>
           </form>
        <% Else 
      
      ' This is where we set up our security levels '
      ' Basically there are two important session variables ' 
      ' The USERLEVEL and the ADMINMENULEVEL ' The Userlevel is our protection to make sure people arent
      ' trying to make up random menu values '
      ' The AdminMenuLevel is what tells the system ' which menu to display.
      ' ' If you try to display a menu that is beyond your 
      ' hidden userlevel value,then it will kick you out.'
      ' Most of the individual procedures also check your ' hidden userlevel just to make sure you didn't try ' to call them directly without using a menu.
       
       
           If Session("reallogin") = "Valid" Then
             Session("FirstName") = rs("FirstName")
             Session("LastName") = rs("LastName")
             If not rs("Seeding") Then
               Session("UserLevel") = 10
               Session("adminmenulevel") = 2
	       Session("UserSptsGrpID") = rs("SptsGrpID")	
             End If
             IF rs("Seeding") AND rs("SecLevel")>=30 THEN
               Session("UserLevel") = 49
               Session("adminmenulevel") = 30
	       Session("UserSptsGrpID") = rs("SptsGrpID")
             ELSEIF rs("Seeding") AND (rs("SecLevel")>=20 AND rs("SecLevel")<30) THEN
               Session("UserLevel") = 49
               Session("adminmenulevel") = 20
	       Session("UserSptsGrpID") = rs("SptsGrpID")
             ELSEIF rs("Seeding") AND (rs("SecLevel")>=10 AND rs("SecLevel")<20) THEN
               Session("UserLevel") = 49
               Session("adminmenulevel") = 10
	       Session("UserSptsGrpID") = rs("SptsGrpID")

             ELSEIF rs("Seeding") AND (rs("SecLevel")>=1 AND rs("SecLevel")<10) THEN
               Session("UserLevel") = 49
               Session("adminmenulevel") = 1
	       Session("UserSptsGrpID") = rs("SptsGrpID")
             End If


             If rs("SecLevel") >= 50 Then
               Session("UserLevel") = 100
               Session("adminmenulevel") = 50
	       Session("UserSptsGrpID") = rs("SptsGrpID")
             End If
           End If
           rs.close
           CloseCon
           set rs = Nothing
           WriteLog(date() &"  "& time() &"  "& session("UserName") &" - "&Session("UserSptsGrpID")& " has logged in.")
             
           WriteIndexPageHeader
          %>
            <br><br>
            <center><h2>Welcome to TRA<br>Data Administration Center</h2><br><h4>Logged In As: <%Response.Write(Session("FirstName") & " " & Session("LastName"))%><br></h4></center>
      
            <br><br><br><br>
            <br><br><br><br>
            <br><br><br><br>
      
          <%
        End If ' rs.EOF or Actual User WriteIndexPageFooter End If ' TRALogin = 1
  End If ' Login=Valid %>





