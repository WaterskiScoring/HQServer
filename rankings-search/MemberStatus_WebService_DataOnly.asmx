<%@ WebService Language="VBScript" Class="Service1" %>
Imports System.Web.Services 
Imports System.Web.Services.Protocols 
Imports System.ComponentModel 
Imports System.Data.SqlClient 
' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line. 
' <System.Web.Script.Services.ScriptService()> _ 
<System.Web.Services.WebService(Namespace:="http://localhost/")> _ 
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _ 
<ToolboxItem(False)> _ 
	Public Class Service1 
	    Inherits System.Web.Services.WebService 
	    <WebMethod()> Public Function searchname(ByVal srctrm As System.String) As System.String 
	        Dim con As SqlConnection 
	        Dim sql As SqlCommand 
	        Dim dr As SqlDataReader 
	        Dim name, age, res As String 
	        
	        Try 
						
						// sConnectionToTRATable = Application("sConnectionToTRATable")
						// Set Con = Server.CreateObject("ADODB.Connection")
  					// Con.ConnectionTimeout = 3000
  					// Con.Open Application("sConnectionToTRATable")
  					// Con.CommandTimeout = 3000
						
						
						// sSQL = "INSERT INTO usawsrank.A_Test (MarkData) VALUES ('1')"
						// con.execute(sSQL)

	        	var res = "<Dim>Test</Dim>"
	        	Return res
	       	End Try  

	    End Function 

	End Class 

