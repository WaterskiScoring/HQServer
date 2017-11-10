<%
If Application("isSecure") then
	If Request.ServerVariables("HTTPS") = "off" then
		URL = "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("SCRIPT_NAME") 
		If Request.QueryString <> "" then
			URL = URL & "?" & Request.QueryString
		End IF
		Response.Redirect URL
		Response.End
	End If
End If
If not Session("aauth") then
	Response.Redirect "/admin/login.asp"
	Response.End
End If 
%>



