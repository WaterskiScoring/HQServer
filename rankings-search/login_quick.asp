<%




session("reallogin") = "Valid"
session("membermenulevel") = ""

username=TRIM(request("username"))

SELECT CASE username
	CASE "mcrone"
		response.redirect("/rankings/loginHQ.asp?username=mcrone&password=1050slsd")

	CASE "dclark"
		response.redirect("/rankings/loginHQ.asp?username=dclark&password=techdude")
END SELECT

%>





