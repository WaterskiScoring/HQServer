<%

' ------------------------------------------------
' This quick check only allows access to the
' settings if the user has actually logged in.
' ------------------------------------------------


If trim(Session("adminmenulevel")) = "" Then
  Session("message") = "You Must Login Before Accessing This System"
  Response.Redirect("/rankings/defaultHQ.asp?process=logout&rid=" & rid)
Else
%>   <!--#include virtual="/rankings/settingsHQ.asp"--> <%
End If


%>





