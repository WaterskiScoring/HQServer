<%


LastName = "O'Connor"

response.write("Test - ") 
response.write(INSTR(LCASE(LastName),"connor"))

response.write("<br>")
response.write(INSTR(LCASE(LastName),"connor")>0)

%>