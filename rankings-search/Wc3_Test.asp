<% @language=vbscript %>



<html>
<body>
	

	
<script type="text/javascript">
txt="<outer>";
txt=txt+"<book>";
txt=txt+"<title>Everyday Italian</title>";
txt=txt+"<author>Giada De Laurentiis</author>";
txt=txt+"<year>2005</year>";
txt=txt+"</book>";
txt=txt+"<book>";
txt=txt+"<title>Mark Recipes</title>";
txt=txt+"<author>Sybelia</author>";
txt=txt+"<year>2011</year>";
txt=txt+"</book>";
txt=txt+"</outer>";
if (window.DOMParser)
  {
  parser=new DOMParser();
  xmlDoc=parser.parseFromString(txt,"text/xml");
  }
else // Internet Explorer
  {
  xmlDoc=new ActiveXObject("Microsoft.XMLDOM");
  xmlDoc.async="false";
  xmlDoc.loadXML(txt); 
  } 


document.write("<table align=center width=50% border='1'>");
document.write("<tr><td align=center colspan=3>");
document.write("<font size=3><b>Silverpop Mailings</b></font>");
document.write("</td></tr>");

var x=xmlDoc.getElementsByTagName("book");
for (i=0;i<x.length;i++)
  { 
  document.write("<tr><td>");
  document.write(x[i].getElementsByTagName("title")[0].childNodes[0].nodeValue);
  document.write("</td><td>");
  document.write(x[i].getElementsByTagName("author")[0].childNodes[0].nodeValue);
  document.write("</td><td>");
  document.write(x[i].getElementsByTagName("year")[0].childNodes[0].nodeValue);
  document.write("</td></tr>");
  }
document.write("</table>");


</script>

</body>
</html>



