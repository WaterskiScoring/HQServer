
<br><br><br>
<center></><h3>This is the Confirmation Page</h3></center>
<center></><p>Please check Salesforce for Conversion Tracking for this message</p></center>


<%

' memberemail="cronemarka@gmail.com"
memberemail=""
'xmlstring = "<system><system_name>tracking</system_name><action>conversion</action><member_id>"& Session("MID") &"</member_id><job_id>"& Session("JobID") &"</job_id><email>"& memberemail &"</email><sub_id>"& Session("SubID") &"</sub_id><list>"& Session("ListID") &"</list><original_link_id>"& Session("LinkID") &"</original_link_id><BatchID>"& Session("BatchID") &"</BatchID><conversion_link_id>1001</conversion_link_id><link_alias>RenewalConfirmation</link_alias><display_order>2</display_order><data_set><data amt=""80"" unit=""Dollars"" accumulate=""true""/></data_set></system>"

xmlstring = "<system><system_name>tracking</system_name><action>conversion</action><member_id>"& Session("MID") &"</member_id><job_id>"& Session("JobID") &"</job_id><email>"& memberemail &"</email><sub_id>"& Session("SubID") &"</sub_id><list>"& Session("ListID") &"</list><original_link_id>"& Session("LinkID") &"</original_link_id><BatchID>"& Session("BatchID") &"</BatchID><conversion_link_id>1001</conversion_link_id><link_alias>RenewalConfirmation</link_alias><display_order>2</display_order><data_set><data amt=""80"" unit=""MembershipFee"" accumulate=""true""/><data amt=""25"" unit=""Products"" accumulate=""true""/></data_set></system>"



response.write("xml = " & xmlstring)
%>

<img src='http://click.exacttarget.com/conversion.aspx?xml=<%= xmlstring %>' width = "1" height ="1">





