<%
Session("aauth") = false
Session("UserID") = ""
Session("UserName") = ""
Session("FullName") = ""
Session("FromUSAWS") = ""
Session("TournamentName") = ""
Session("TournamentDate") = ""
Session("TournamentYear") = ""
Session.Abandon
Response.Redirect("/admin/login.asp")
%>





