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

	      XmlDocument xmlDoc = new XmlDocument();
	       	
				// Write down the XML declaration
        //XmlDeclaration xmlDeclaration = xmlDoc.CreateXmlDeclaration("1.0","utf-8",null); 
				xmlDeclaration = xmlDoc.CreateXmlDeclaration("1.0","utf-8",null); 

        // Create the root element
        XmlElement rootNode  = xmlDoc.CreateElement("USAWaterski");
        xmlDoc.InsertBefore(xmlDeclaration, xmlDoc.DocumentElement); 
        xmlDoc.AppendChild(rootNode);

        // Create a new <MemberStatus> element and add it to the root node
        XmlElement parentNode  = xmlDoc.CreateElement("MemberStatus");
        xmlDoc.DocumentElement.PrependChild(parentNode);

        // Create the nodes
        XmlElement idNode  = xmlDoc.CreateElement("MemberID");
        XmlElement foundNode  = xmlDoc.CreateElement("MemberFound");
        XmlElement firstNode  = xmlDoc.CreateElement("MemberFirst");
        XmlElement lastNode  = xmlDoc.CreateElement("MemberLast");
        XmlElement expireNode  = xmlDoc.CreateElement("MemberExpireDate");
        XmlElement typeNode  = xmlDoc.CreateElement("MembershipType");
        XmlElement canskitourNode  = xmlDoc.CreateElement("CanSkiTour");

        // Define the values 
        XmlText idText= xmlDoc.CreateTextNode("000001151");
        XmlText foundText  = xmlDoc.CreateTextNode("Yes");
        XmlText firstText  = xmlDoc.CreateTextNode("Mark");
				XmlText lastText  = xmlDoc.CreateTextNode("Crone");
				XmlText expireText  = xmlDoc.CreateTextNode("2/1/2016");
				XmlText typeText  = xmlDoc.CreateTextNode("Individual Active");
				XmlText canskitourText  = xmlDoc.CreateTextNode("Yes");
				
        // Append the nodes to the parentNode without the value
        parentNode.AppendChild(idNode);
        parentNode.AppendChild(foundNode);
        parentNode.AppendChild(firstNode);
        parentNode.AppendChild(lastNode);
        parentNode.AppendChild(expireNode);
        parentNode.AppendChild(typeNode);
        parentNode.AppendChild(canskitourNode);
        
        // Save the value of the fields into the nodes
        idNode.AppendChild(idText);
        foundNode.AppendChild(foundText);
        firstNode.AppendChild(firstText);
        lastNode.AppendChild(lastText);
        expireNode.AppendChild(expireText);
        typeNode.AppendChild(typeText);
        canskitourNode.AppendChild(canskitourText);
        
        res = "<?xml version="1.0" encoding="utf-8" ?><USAWaterski><MemberStatus><MemberID>000001151</MemberID><MemberFound>yes</MemberFound><MemberFirst>Mark</MemberFirst><MemberLast>Crone</MemberLast><MemberExpireDate>4/1/2016</MemberExpireDate><MembershipType>Individual Active</MembershipType><CanSkiTour>yes</CanSkiTour></MemberStatus></USAWaterski>"
	      return res
	      //Return xmlDoc 
	    
	    End Function 

	End Class 

