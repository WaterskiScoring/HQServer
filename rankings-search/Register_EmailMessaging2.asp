<!--#include virtual="/rankings/settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include16.asp"-->
<!--#include virtual="/rankings/tools_registration16.asp"-->
<!--#include virtual="/rankings/tools_mobile_version.asp"-->


<style type="text/css">
	
	var HQSiteColor1="#203f5e";
	var HQSiteColor2="#0f77da";
	var HQSiteColor3="#2F4F4F";
	
	
	html {
		/*background:#1a1a1a;*/
		margin:0px 0px 0px 0px;
		padding:0px 0px 0px 0px;
		background:#FFFFFF;
		/* background:#262626; */
		/* yellow;  */
		height:800px;
	}

	
.container {
			font-family: Arial, Helvetica, sans-serif; 
			font-weight:normal; 
			font-size:12px;
			font-weight:bold;
			color:#000000;
		  background-color: #FFFFFF;
			display:inline-block;
			margin:0px;
			padding:0px;
			height:100%;
		}
		
body {
		/* -ms-text-size-adjust:120%; */
		/* -webkit-text-size-adjust:120; */
		/* -moz-text-size-adjust:120%; */
		margin:0px 0px 0px 0px;
		padding:0px 0px 0px 0px;
 		background-color:grey;
 		height:100%
	}

input[type="text"] {
		margin:0px;
		padding:0px;
		display:inline-block;		
	}

div {
		padding:0px;
		margin:0px;
	}

		
.accordionheader	{
		color:#FFFFFF;
		background-color: #2E4d7B;
		height:30px;
		width:100%;
		font-size:11pt;
		border:1px solid black;
		display:inline-block;
		margin:0px 0px 0px 0px;
		padding:0px 0px 0px 0px;
		-moz-box-sizing: border-box;
    -webkit-box-sizing: border-box;
    -ms-box-sizing: border-box;
    box-sizing: border-box;	
		}	

.accordionbody	{
		color:#000000;
		border:1px solid yellow;
		background-color: #FFFFFF;
		width:100%;
		display:inline-block;
		border:1px solid #0f77da;
		margin:0px 0px 0px 0px;
		padding:15px 0px 10px 0px;			
		}	

.textbox  {
		background-color:#FFF8DC;
		border: 1px solid blue;	
		display:inline-block;		
}

.textlabel {
		width:15%;
		height:22px;
		color:blue;
		text-align:right;
		border:0px solid red;
		display:inline-block;
		margin:0px 0px 0px 0px;
		padding:0px 0px 0px 0px;
		-moz-box-sizing: border-box;
    -webkit-box-sizing: border-box;
    -ms-box-sizing: border-box;			
		}

.textdata {
		width:80%;
		height:22px;
		color:#000000;
		text-align:left;
		border:0px solid green;
		display:inline-block;
		margin:0px 0px 0px 0px;
		padding:0px 0px 0px 10px;
		-moz-box-sizing: border-box;
    -webkit-box-sizing: border-box;
    -ms-box-sizing: border-box;		
		}

textarea {
		font-family: Arial, Helvetica, sans-serif; 
		font-weight:normal; 
		color:#000000;
	}
			
.stdbutton {
	width:9em;
	height:2em;
	-webkit-appearance: none;
	padding:0px;
	margin:0px;
	border-radius: 10px; 
	background-color:#D3D3D3; 
	}

.greenbutton {
	width:9em;
	height:2em;
	-webkit-appearance: none;
	padding:0px;
	margin:0px;	
	border-radius: 10px; 
	background-color:#ADFF2F;
	}

.yellowbutton {
	width:9em;
	height:2em;
	-webkit-appearance: none;
	padding:0px;
	margin:0px;	
	border-radius: 10px; 
	background-color:#FFD700;
	}
	
.redbutton {
	width:9em;
	height:2em;
	-webkit-appearance: none;
	padding:0px;
	margin:0px;	
	border-radius: 10px; 
	background-color:red;
	color:#FFFFFF;
	}

.buttonrowemail {
		width:100%;
		display:inline-block;
		text-align:center;
		margin:15px 0px 0px 0px; 
		height:28px; 
		padding:0px;
		border:0px solid green;
		-moz-box-sizing: border-box;
    -webkit-box-sizing: border-box;
    -ms-box-sizing: border-box;		
	}

	.span5 { width:5%; display:inline-block; }
	.span10 { width:10%; display:inline-block; }
	.span15 { width:15%; display:inline-block; }
	.span20 { width:20%; display:inline-block; }
	.span25 { width:25%; display:inline-block; }
	.span30 { width:30%; display:inline-block; }
	.span35 { width:35%; display:inline-block; }
	.span40 { width:40%; display:inline-block; }
	.span45 { width:45%; display:inline-block; }
	.span50 { width:50%; display:inline-block; }
	.span55 { width:55%; display:inline-block; }
	.span60 { width:60%; display:inline-block; }
	.span65 { width:65%; display:inline-block; }
	.span70 { width:70%; display:inline-block; }
	.span75 { width:75%; display:inline-block; }
	.span80 { width:80%; display:inline-block; }	
	.span85 { width:85%; display:inline-block; }
	.span90 { width:90%; display:inline-block; }
	.span95 { width:95%; display:inline-block; }
	.span100 { width:100%; display:inline-block}
	
			
</style>




<script type="text/javascript">
	
	
	function UpdateTabDisplay(WhichTab) {
						
			document.getElementById('EmailCMS_TournamentBody').style.display = 'none'; 		
			document.getElementById('EmailCMS_MessageBody').style.display = 'none'; 		
			document.getElementById('EmailCMS_RecipientBody').style.display = 'none'; 
			document.getElementById('EmailCMS_SendBody').style.display = 'none'; 			
			document.getElementById('EmailCMS_ConfirmBody').style.display = 'none'; 			
			
			document.getElementById('EmailCMS_TournamentHeader').style.backgroundColor = '#2E4d7B';
			document.getElementById('EmailCMS_MessageHeader').style.backgroundColor = '#2E4d7B';					
			document.getElementById('EmailCMS_RecipientHeader').style.backgroundColor = '#2E4d7B';
			document.getElementById('EmailCMS_SendHeader').style.backgroundColor = '#2E4d7B';
			document.getElementById('EmailCMS_ConfirmHeader').style.display = '#2E4d7B'; 
			
			if (WhichTab == 'tournament') { 
					document.getElementById('EmailCMS_TournamentBody').style.display = 'inline-block'; 
					document.getElementById('EmailCMS_TournamentHeader').style.backgroundColor = '#5078B3';
					}
			if (WhichTab == 'message') { 
					document.getElementById('EmailCMS_MessageBody').style.display = 'inline-block'; 	
					document.getElementById('EmailCMS_MessageHeader').style.backgroundColor = '#5078B3';
					ValidateMessageForm()
					}
			if (WhichTab == 'recipients') { 
					document.getElementById('EmailCMS_RecipientBody').style.display = 'inline-block'; 
					document.getElementById('EmailCMS_RecipientHeader').style.backgroundColor = '#5078B3';
					}
			if (WhichTab == 'send') { 
					document.getElementById('EmailCMS_SendBody').style.display = 'inline-block'; 
					document.getElementById('EmailCMS_SendHeader').style.backgroundColor = '#5078B3';
					}
			if (WhichTab == 'confirm') { 
					document.getElementById('EmailCMS_ConfirmBody').style.display = 'inline-block'; 
					document.getElementById('EmailCMS_ConfirmHeader').style.backgroundColor = '#5078B3';
					}					
		}
		

	function ValidateUseAcceptance() {

					var AcceptRelease = document.getElementById('AcceptRelease').value;	
			
				
					//alert('AcceptRelease = ' + AcceptRelease);

					// var AC_Len = AdminCode_Tournament.length;
					if ( document.getElementById('AcceptRelease').checked == true ) {
							// alert('CHECKED')
							// document.getElementById('TournamentValidate').style.display = 'none';					
							// document.getElementById('TournamentContinue').style.display = 'inline-block';
							document.getElementById('TournamentContinue').style.backgroundColor = '#ADFF2F';	
							document.getElementById('TournamentContinue').disabled = false;
						}
					else {
							//alert('ELSE')
							// document.getElementById('TournamentValidate').style.display = 'inline-block';	
							document.getElementById('TournamentContinue').style.backgroundColor = '#FFD700';
							document.getElementById('TournamentContinue').disabled = true;				
							// document.getElementById('TournamentContinue').style.display = 'none';		
							alert('You must accept the Terms of Use');
						}	
		}

	
	function ValidateMessageForm() {

			var Subject = document.getElementById('ThisTemplateSubject').value
			if( Subject == "" ) {
					document.getElementById('MessageContinue').disabled = true;
					document.getElementById('MessagePreviewTemplate').disabled = true;
					document.getElementById('MessageEditTemplate').disabled = true;
					document.getElementById('ValidateTemplate').disabled = true;
					document.getElementById('MessageNewTemplate').disabled = false;
					document.getElementById('MessageNewTemplate').value = 'New Template';
				}
			
			// --- A TEMPLATE has been selected ---
			if (document.getElementById('TemplateIDSelected').value == "90000001") {
					document.getElementById('MessageContinue').disabled = true;
					document.getElementById('ValidateTemplate').disabled = true;
					document.getElementById('MessagePreviewTemplate').disabled = true;
					document.getElementById('MessageEditTemplate').disabled = true;
					document.getElementById('MessageNewTemplate').style.display = 'none';
					document.getElementById('MessageCopyTemplate').style.display = 'inline-block';							
				}
			

			
			if (document.getElementById('ThisReplyToEmail').value == "YourReplyTo EMail Address") {
				
				}
			
			
		}





	function ConfirmSelectedElement(whatelement) {
				// alert('In Function');
				var ThisElement = whatelement + 'Selected';
				var ThisElementText = whatelement + 'SelectedText';
				var x=document.getElementById(ThisElement);
     		var ValuesSelected;
     		ValuesSelected = '';
  			for (var i = 0; i < x.options.length; i++) {
						ThisOption = x.options[i].value
     				if(x.options[i].selected == true) {
 								ValuesSelected = ValuesSelected + ThisOption.trim() + ' ';
          	}
  				}
    		if ( ValuesSelected.length == 0 )
    				{ alert('No ' + whatelement + ' has been selected'); }
    		else
    				{ alert(whatelement + 's you selected in the dropdown are: ' + ValuesSelected); }
			}





	function UpdateTemplateFormAction(WhichAction) {
			//alert('WhichAction = ' + WhichAction);

			if (WhichAction == 'edittemplate') {
					//alert('In edit mode');
					document.getElementById('MessageForm').action = location.pathname + '?action=updatecurrenttemplate';
					
					// -- Exposes and enables text input and hides dropdown --  
					document.getElementById('ThisTemplateSubject').style.display = 'inline-block';
					document.getElementById('ThisTemplateSubject').disabled = false;
					document.getElementById('TemplateIDSelected').style.display = 'none'; 
					
					TurnFieldsRedAndDisable()
					
					document.getElementById('MessageEditTemplate').style.display = 'none';					
					document.getElementById('MessageUpdateTemplate').style.display = 'inline-block';
					
					document.getElementById('MessageContinue').disabled = true;
					// document.getElementById('MessagePreviewTemplate').disabled = true;	
					document.getElementById('MessagePreviewTemplate').style.display = 'none';					
					document.getElementById('MessageCancelTemplate').style.display = 'inline-block';					
					document.getElementById('MessageBack').disabled = true;
					document.getElementById('MessageNewTemplate').disabled = true;	
				}

			if (WhichAction == 'copytemplate') {
					//alert('In edit mode');
					
					// -- Exposes and enables text input and hides dropdown --  
					document.getElementById('ThisTemplateSubject').style.display = 'inline-block';
					document.getElementById('ThisTemplateSubject').disabled = false;
					document.getElementById('TemplateIDSelected').style.display = 'none';
					var origsubject = document.getElementById('ThisTemplateSubject').value;
					var newsubject = origsubject.replace(" TEMPLATE", ""); 
					var tourname = document.getElementById('ThisTournamentName').value
					
					document.getElementById('ThisTemplateSubject').value = tourname.substr(0,20) + ' - '+newsubject;
					document.getElementById('TemplateSubjectInHeader').value = tourname.substr(0,20) + ' - '+newsubject;					
					
					TurnFieldsRedAndDisable()
		
					document.getElementById('MessageEditTemplate').style.display = 'none';					
					document.getElementById('MessageUpdateTemplate').style.display = 'none';
					document.getElementById('MessageSaveTemplate').style.display = 'inline-block';
					document.getElementById('MessagePreviewTemplate').style.display = 'none';					
					document.getElementById('MessageCancelTemplate').style.display = 'inline-block';
										
					document.getElementById('MessageSaveTemplate').disabled = false;
					document.getElementById('MessageContinue').disabled = true;
					document.getElementById('MessagePreviewTemplate').disabled = true;	
					document.getElementById('MessageBack').disabled = false;
					document.getElementById('MessageNewTemplate').disabled = true;
					document.getElementById('MessageCopyTemplate').disabled = true;
				}

			if (WhichAction == 'newtemplate') {
					//alert('In edit mode');
					
					// -- Exposes and enables text input and hides dropdown --  
					document.getElementById('ThisTemplateSubject').style.display = 'inline-block';
					document.getElementById('ThisTemplateSubject').disabled = false;
					document.getElementById('TemplateIDSelected').style.display = 'none';
					var newsubject = 'Your Subject';
					var tourname = document.getElementById('ThisTournamentName').value
					
					//document.getElementById('ThisTemplateSubject').value = tourname.substr(0,20) + ' - '+newsubject;
					//document.getElementById('TemplateSubjectInHeader').value = tourname.substr(0,20) + ' - '+newsubject;					
					document.getElementById('ThisTemplateSubject').value = tourname.substr(0,25);
					document.getElementById('TemplateSubjectInHeader').value = tourname.substr(0,25);					

					
					TurnFieldsRedAndDisable()
					
					// alert('ThisTourDirector = ' + document.getElementById('ThisTourDirector').innerHTML)
					document.getElementById('ThisSenderSignature').value = document.getElementById('ThisTourDirector').innerHTML;
					document.getElementById('ThisReplyToEmail').value = document.getElementById('ThisTourDirEmail').innerHTML;
					document.getElementById('ThisSalutation').value = 'Dear Skiers';
					document.getElementById('ThisSenderTitle').value = 'Tournament Director';
					//document.getElementById('ThisTemplate_Body').value = 'Your message here.  Use HELP to learn about HTML markup options';					
					// document.getElementById('TemplateIDSelected').value
												
					document.getElementById('MessageEditTemplate').style.display = 'none';					
					document.getElementById('MessageUpdateTemplate').style.display = 'none';
					document.getElementById('MessageSaveTemplate').style.display = 'inline-block';
					document.getElementById('MessagePreviewTemplate').style.display = 'none';					
					document.getElementById('MessageCancelTemplate').style.display = 'inline-block';
										
					document.getElementById('MessageContinue').disabled = true;
					document.getElementById('MessagePreviewTemplate').disabled = true;	
					document.getElementById('MessageBack').disabled = false;
					document.getElementById('MessageNewTemplate').disabled = true;
					document.getElementById('MessageCopyTemplate').disabled = true;
				}
				
			if (WhichAction == 'canceledittemplate') {
					//alert('In edit mode');
					
					// -- Exposes and enables text input and hides dropdown --  
					document.getElementById('ThisTemplateSubject').style.display = 'none';
					document.getElementById('ThisTemplateSubject').disabled = true;
					document.getElementById('TemplateIDSelected').style.display = 'inline-block'; 
					
					document.getElementById('ThisSalutation').disabled = true;
					document.getElementById('ThisTemplate_Body').disabled = true;	
					document.getElementById('ThisSenderSignature').disabled = true;	
					document.getElementById('ThisSenderTitle').disabled = true;	
					document.getElementById('ThisReplyToEmail').disabled = true;	
					document.getElementById('ValidateTemplate').disabled = false;
					//document.getElementById('messagelist_existing').disabled = false;
					//document.getElementById('messagelist_template').disabled = false;				

					document.getElementById('ThisTemplateSubject').style.color = 'black'
					document.getElementById('ThisSalutation').style.color = 'black'
					document.getElementById('ThisTemplate_Body').style.color = 'black'										
					document.getElementById('ThisSenderSignature').style.color = 'black'
					document.getElementById('ThisSenderTitle').style.color = 'black'
					document.getElementById('ThisReplyToEmail').style.color = 'black'
					
		
					document.getElementById('MessageEditTemplate').style.display = 'inline-block';					
					document.getElementById('MessageUpdateTemplate').style.display = 'none';
					document.getElementById('MessageSaveTemplate').style.display = 'none';
					document.getElementById('MessagePreviewTemplate').style.display = 'inline-block';					
					document.getElementById('MessageCancelTemplate').style.display = 'none';
															
					document.getElementById('MessageContinue').disabled = true;
					document.getElementById('ValidateTemplate').disabled = true;					
					document.getElementById('MessagePreviewTemplate').disabled = false;	
					document.getElementById('MessageBack').disabled = true;
					document.getElementById('MessageNewTemplate').disabled = false;
					document.getElementById('MessageCopyTemplate').disabled = false;
						
				}


			else if (WhichAction == 'templateselectedbyuser') {
				 	//alert('IN selected');
					document.getElementById('MessageForm').action = location.pathname + '?action=templateselectedbyuser';
				}
		}


		function TurnFieldsRedAndDisable() {
					document.getElementById('ThisSalutation').disabled = false;
					document.getElementById('ThisTemplate_Body').disabled = false;	
					document.getElementById('ThisSenderSignature').disabled = false;	
					document.getElementById('ThisSenderTitle').disabled = false;	
					document.getElementById('ThisReplyToEmail').disabled = false;	
					document.getElementById('ValidateTemplate').disabled = true;
					//document.getElementById('messagelist_existing').disabled = true;
					//document.getElementById('messagelist_template').disabled = true;
					
					

					document.getElementById('ThisTemplateSubject').style.color = 'red'
					document.getElementById('ThisSalutation').style.color = 'red'
					document.getElementById('ThisTemplate_Body').style.color = 'red'										
					document.getElementById('ThisSenderSignature').style.color = 'red'
					document.getElementById('ThisSenderTitle').style.color = 'red'
					document.getElementById('ThisReplyToEmail').style.color = 'red'
	
				}
	




		
	function ValidateTemplateCriteria() {

			var TemplateSubject = document.getElementById('ThisTemplateSubject').value;
			var Salutation = document.getElementById('ThisSalutation').value;
			var SenderSignature = document.getElementById('ThisSenderSignature').value;
			var SenderTitle = document.getElementById('ThisSenderTitle').value;
			var ReplyToEmail = document.getElementById('ThisReplyToEmail').value;	
			var ThisTemplate_Body = document.getElementById('ThisTemplate_Body').value;		
			//alert('SenderCopySelect = ' + SenderCopySelect);

			if ( ThisTemplate_Body.trim() == '' || TemplateSubject.trim() == '' || Salutation.trim() == '' || SenderSignature.trim() == '' || SenderTitle.trim() == '' || ReplyToEmail.trim() == '') {
					alert('All template fields must be complete to proceed');
					}
			else {
					document.getElementById('MessageContinue').style.display = 'inline-block';
					document.getElementById('ValidateTemplate').style.display = 'none';	
					}			
		}



	function ValidateRecipientCriteria() {

				var x=document.getElementById('DivSelected');
     		var DivSelected;
     		DivSelected = '';
  			for (var i = 0; i < x.options.length; i++) {
						ThisOption = x.options[i].value
     				if(x.options[i].selected == true) {
 								DivSelected = DivSelected + ThisOption.trim() + '|';
          	}
  				}
  		
				var x=document.getElementById('EventSelected');
     		var EventSelected;
     		EventSelected = '';
  			for (var i = 0; i < x.options.length; i++) {
						ThisOption = x.options[i].value
     				if(x.options[i].selected == true) {
 								EventSelected = EventSelected + ThisOption.trim() + '|';
          	}
  				}
  		
  				
			if ( DivSelected.trim() == '' || EventSelected.trim() == '' ) {
					alert('You must select at least one Div and Event');
					}
			else {
					document.getElementById('RecipientContinue').style.display = 'inline-block';
					document.getElementById('ValidateRecipient').style.display = 'none';	
					}			
		}


	function ValidateSendCriteria() {

			var SenderCopySelect = document.getElementById('SenderCopySelect').value;
			var Send_DivSelected = document.getElementById('Send_DivSelected').value;
			var Send_EventSelected = document.getElementById('Send_EventSelected').value;
			//alert('SenderCopySelect = ' + SenderCopySelect);
			//if ( AdminCode_Tournament.trim() == '' || Send_DivSelected.trim() == '' || Send_DivSelected.trim() == '' || SenderCopySelect.trim() == '') {
			// if ( AdminCode_Tournament.trim() == '' || Send_DivSelected.trim() == '' || Send_DivSelected.trim() == '') {
			if ( Send_DivSelected.trim() == '' || Send_EventSelected.trim() == '' || SenderCopySelect.trim() == '' ) {
					alert('You must a) Select your Copy Method');
					}
			else {
					document.getElementById('SendMessageNow').style.display = 'inline-block';
					document.getElementById('ValidateSend').style.display = 'none';
					alert('When you press Send Now the current template will deploy to the recipients');	
					}			
		}


	function DeactivateSendButton() {
					document.getElementById('ValidateTemplate').disabled = true;		
		
		}


function ShowMessagePreview()
	{
 		var emailPreviewWindow = window.open("","Preview","width=450px, height=500px, scrollbars=yes, resizable=yes, status=0");
		var USAWS_Logo ='http://www.usawaterski.org/rankings/images/logos/usawslogo_no_sub.jpg'

    emailPreviewWindow.document.open();
    emailPreviewWindow.document.writeln('<HTML><HEAD><TITLE>Message Preview</TITLE></HEAD>');
    emailPreviewWindow.document.writeln('<BODY style="font-family: Arial, Helvetica, sans-serif; text-align:left; font-size:10pt;">');
    // emailPreviewWindow.document.writeln('<div style="width:100%; margin-top:30px;">Please <b>DO NOT REPLY</b> to this email as this is not a monitored email address. For questions, please contact me at: <a href=mailto:' + document.getElementById('ThisReplyToEmail').value + ' ?subject=Question About ' + document.getElementById('ThisTournamentName').value + 'style=text-decoration:none;>' + document.getElementById('ThisReplyToEmail').value + '</a></div>');

 
    emailPreviewWindow.document.writeln('<div style="width:100%; margin-top:40px;">Re: ' + document.getElementById('ThisTournamentName').value + ':</div>');
    emailPreviewWindow.document.writeln('<div style="width:100%;">TourID: ' + document.getElementById('sTourID').value + '</div>');
    emailPreviewWindow.document.writeln('<div style="width:100%; margin-top:0px;">Event Date: ' + document.getElementById('ThisTourDate').value + '</div>');
                   
    emailPreviewWindow.document.writeln('<div style="width:100%; margin-top:20px;">' + document.getElementById('ThisSalutation').value + ':</div>');
    emailPreviewWindow.document.writeln('<div style="width:100%; margin-top:20px;">' + document.getElementById('ThisTemplate_Body').value + '</div>');
    emailPreviewWindow.document.writeln('<div style="width:100%; margin-top:20px;">Sincerely,</div>');
    emailPreviewWindow.document.writeln('<div style="width:100%; margin-top:20px;">' + document.getElementById('ThisSenderSignature').value + '</div>');
    emailPreviewWindow.document.writeln('<div style="width:100%; margin:0px 0px 30px 0px;">' + document.getElementById('ThisSenderTitle').value + '</div>'); 

    emailPreviewWindow.document.writeln('<div style="width:100%; text-align:center; font-size:8pt; margin:0px 0px 0px 0px;">A Service of</div>'); 
    emailPreviewWindow.document.writeln('<div style="width:100%; text-align:center; margin:10px 0px 0px 0px;"><img src=' + USAWS_Logo + ' style=width:100px;></div>'); 
    emailPreviewWindow.document.writeln('<div style="width:100%; text-align:center; font-size:8pt; font-style:bold margin:10px 0px 0px 0px;">180 Holy Cow Rd<br>Polk City, FL 33883</div>'); 
          
    emailPreviewWindow.document.writeln('<A HREF="javascript:window.close()">Close Preview</A><BR>');

		//emailPreviewWindow.document.writeln(document.SampleForm.MessageText.value);
    emailPreviewWindow.document.writeln('</BODY></HTML>');
    emailPreviewWindow.document.close();
	}



function ShowFormatHelp()
	{
 		var emailHelpWindow = window.open("","Preview","width=325px, height=500px, scrollbars=yes, resizable=yes, status=0");

    emailHelpWindow.document.open();
    emailHelpWindow.document.writeln('<HTML><HEAD><TITLE>Format Help</TITLE></HEAD>');
    emailHelpWindow.document.writeln('<BODY style="font-family: Arial, Helvetica, sans-serif; text-align:left; font-size:10pt;">');

    emailHelpWindow.document.writeln('<div style="width:100%; margin:40px 0px 0px 0px;"><h3>MESSAGE BODY FORMATTING HELP</h3></div>');  
    emailHelpWindow.document.writeln('<div style="width:100%; margin:20px 0px 0px 0px;"><b>USING A TEMPLATE:</b> Select View Templates from the subject dropdown.  The available template options will then appear in the listing.  Before you can use one of the templates, you must first press the COPY TEMPLATE button and then edit it to apply your tournament information. When you do this the template will be saved in the templates for the tournament.</div>');
    emailHelpWindow.document.writeln('<div style="width:100%; margin:20px 0px 0px 0px;"><b>EXISTING MESSAGE:</b> Select Existing Message. This will display any messages you have previously created for this tournament.</div>');   
    emailHelpWindow.document.writeln('<div style="width:100%; margin:20px 0px 0px 0px;"><b>NEW MESSAGE:</b> Select New Template button. This will create a blank template associated with this tournament</div>');   
    emailHelpWindow.document.writeln('<div style="width:100%; margin:20px 0px 0px 0px;"><b>HTML:</b> A limited amount of html markup may be used in the MESSAGE BODY of the template.  Markup MAY NOT be used in Subject Line, Salutation, Sender Name or Sender Title. <br><br>Line feeds you try to add to your text will NOT show up in the email template.  You must explicitly add the html code to your text.  Please see the allowed markup below.</div>');     
    emailHelpWindow.document.writeln('<div style="width:100%; margin:20px 0px 0px 0px;"><b>LINE BREAK:</b> Insert &#60;br&#62; before the new line. Insert 2 times &#60;br&#62;&#60;br&#62; to leave space between paragraphs. You can also insert a &#60;br&#62 in the Title line. No breaks are permitted in the Subject, Salute, Signature or ReplyTo</div>');        
    emailHelpWindow.document.writeln('<div style="width:100%; margin:10px 0px 0px 0px;"><b>BOLD:</b> Make a section <b>bold</b> by inserting &#60;b&#62; before the section and &#60;/b&#62; at the end of the section. Example: &#60;b&#62;<b>Your Text Here</b>&#60;/b&#62;</div>');
    emailHelpWindow.document.writeln('<div style="width:100%; margin:10px 0px 30px 0px;"><b>HEADLINE:</b> Create a headline by inserting &#60;h3&#62; before the headline and &#60;/h3&#62; after the headline. Please note that the h3 markup will automatically create a line feed after the headline.</div>');

    emailHelpWindow.document.writeln('<A HREF="javascript:window.close()">Close Preview</A><BR>');

		// emailHelpWindow.document.writeln(document.SampleForm.MessageText.value);
    emailHelpWindow.document.writeln('</BODY></HTML>');
    emailHelpWindow.document.close();
	}
	


function WhereToStartHelp() {
 		var emailHelpWindow = window.open("","Preview","width=325px, height=500px, scrollbars=yes, resizable=yes, status=0");

    emailHelpWindow.document.open();
    emailHelpWindow.document.writeln('<HTML><HEAD><TITLE>Getting Started</TITLE></HEAD>');
    emailHelpWindow.document.writeln('<BODY style="font-family: Arial, Helvetica, sans-serif; text-align:left; font-size:10pt;">');

    emailHelpWindow.document.writeln('<div style="width:100%; margin:40px 0px 0px 0px;"><h3>GETTING STARTED</h3></div>');  
    emailHelpWindow.document.writeln('<div style="width:100%; margin:20px 0px 0px 0px;"><b>GENERAL:</b> Hover over fields, buttons and drop downs to see the pop up explaining more about each function. <br><br>This program may be used to communicate with tournament entrants about this tournament or about an upcoming USA Water Ski sanctioned event.</div>');
     emailHelpWindow.document.writeln('<div style="width:100%; margin:20px 0px 0px 0px;"><b>HERE IS WHAT TO EXPECT ON EACH TAB</b><br></div>');
    emailHelpWindow.document.writeln('<div style="width:100%; margin:20px 0px 0px 0px;"><b>1. TOURNAMENT:</b> On this tab you must agree to the Terms of Use by checking the box and pressing Continue.</div>');
    emailHelpWindow.document.writeln('<div style="width:100%; margin:20px 0px 0px 0px;"><b>2. CONTENT:</b> Here you will build the message you want to send and specify the Subject Line, Salutation, Signature Lines (2) and ReplyToEmail.  You may start with a blank message or Copy a Template from the ones provided. Existing messages may be selected from the drop down for sending now or at a future date.<br><br>Once you have saved your new or copied message it may be Edited and Previewed.  Make sure if you have copied a template that you change (Edit) the places in the Subject and Message Body where customization is intended. Formatting options may be accessed from the Help button.</div>');   
    emailHelpWindow.document.writeln('<div style="width:100%; margin:20px 0px 0px 0px;"><b>3. RECIPIENTS:</b> This is where you select WHO you want to send to.  Select those divisions and events you want included in the distribution.  Only those divisions/events with entrants for this tournament will appear in the list.  If you have many divisions/events included in your selection, you can press the Check Divisions/Events buttons and a pop up will display a list your selections.  You must select at least one division and one event.</div>');   
    emailHelpWindow.document.writeln('<div style="width:100%; margin:20px 0px 0px 0px;"><b>4. SENDING:</b> This is your last opportunity to check the settings.  Make sure everything is correct before proceeding.<br><br>The Copy method allows you to specify whether you receive: a) One message per send or b) a CC on every message or c) a BCC on every message or d) No copies of the message.</div>');     
    emailHelpWindow.document.writeln('<div style="width:100%; margin:20px 0px 0px 0px;"><b>5. CONFIRMATION:</b> When you reach this page, your message(s) will have been sent.  A summary record of the send is saved containing the parameters of the mailing (NOTE: Report for this not yet available).  A record of who was sent is also stored (NOTE: Report for this not yet available).  Once you return to the Tournament tab it will show the summary information about the mailing you just sent.</div>');        
    emailHelpWindow.document.writeln('<br>'); 
    emailHelpWindow.document.writeln('<A HREF="javascript:window.close()">Close Preview</A><BR>');

		// emailHelpWindow.document.writeln(document.SampleForm.MessageText.value);
    emailHelpWindow.document.writeln('</BODY></HTML>');
    emailHelpWindow.document.close();
	}	
		
</script>
<%




' ===============================================================================================
' -- MAIN SECTION OF PROGRAM --

' ===============================================================================================


Dim Action, WhichTab, OpenState 
Dim SendErrorMessage, SendErrorCode_DisplayStatus
Dim DivSelected, EventSelected, TemplateIDSelected, SpecialSendSelect
Dim EventSelected_ForSQL, DivSelected_ForSQL

Dim ThisTournamentName, sTourID, ThisTourDate, ThisTourDirector, ThisTourDirEmail
Dim ThisTemplateSubject, ThisSalutation, ThisSenderSignature, ThisSenderTitle, ThisReplyToEmail
Dim AdminCode_Tournament, SenderCopySelect
Dim ThisTemplate_Body, eMailBody

Dim SpecialSendSelectStatus, DivDropStatus
Dim TemplateFieldStatus
Dim MessageStatus_Found, MessageID_Found, Recipient_Count, NumberRecipientsSent

Dim ebody, RPText

Dim ThisFileName


ThisSitePath="/rankings"
ThisFileName = "Register_EmailMessaging2.asp"
MenuFilename = "DefaultHQ.asp"







SpecialSendSelectStatus="enabled"
DivDropStatus="enabled"



IF TRIM(Request("sMemberID"))<>"" THEN
		sMemberID=Request("sMemberID")
		Session("sMemberID")=sMemberID
ELSE
		sMemberID = Session("sMemberID")
END IF

IF TRIM(sMemberID)="" THEN
		Session("sSendingPage")="/rankings/"&ThisFileName
		Response.Redirect("/rankings/search-memberHQ.asp?rid="&rid&"&formstatus=search")
END IF



' -- Set from login_registrar.asp --
' IF TRIM(Session("sTourID"))<>"" THEN

IF Session("sTourID")="" THEN 
		response.redirect("/rankings/"&RegFileName&"?sRunByWhat=Tour")
ELSE
		sTourID = Session("sTourID")
END IF




' --- TOURNAMENT DATA - Pulls from J Meis function --
DefineTourVariables_New

ThisTournamentName = sTourName
ThisTourDate = sTDateS & "-" & sTDateE 
ThisTourDate = MID(sTDateS,1,Instr(4,sTDateS,"/")-1) & "-" & MID(sTDateE,1,Instr(4,sTDateE,"/")) & RIGHT(sTDateE,2) 
ThisTourDirector = sTDirName
ThisTourDirEmail = sTDirEmail





' --- Controlling variable ---
Action = TRIM(LCASE(Request("Action")))
messagelist = TRIM(Request("messagelist"))
IF messagelist="" THEN messagelist="template"


' --- Control action without using querystring --
TemplateIDSelected = Request("TemplateIDSelected")
IF Request("TemplateButton")="Save" THEN action="insertcurrenttemplate"
IF Request("TemplateButton")="none" THEN action="message"






SendErrorCode_DisplayStatus = "none"

' -- Selected Divisions and Events --
DivSelected = TRIM(Request("DivSelected"))
EventSelected = TRIM(Request("EventSelected"))

DivSelected_ForSQL1 = "'" &Replace(DivSelected,"," ,"','")&"'" 
DivSelected_ForSQL = Replace(DivSelected_ForSQL1," ","")
EventSelected_ForSQL1 = "'" &Replace(EventSelected,"," ,"','")&"'" 
EventSelected_ForSQL = Replace(EventSelected_ForSQL1," ","")


SpecialSendSelect = TRIM(Request("SpecialSendSelect"))
AdminCode_Tournament = TRIM(Request("AdminCode_Tournament"))
SenderCopySelect = TRIM(Request("SenderCopySelect"))







'IF Action="sendmessage" AND UCASE(TRIM(AdminCode_Tournament)) <> UCASE(TRIM(Session("AdminCode"))) THEN
'		Action = "ontosend"	
'		SendErrorMessage = "Invalid AdminCode for this tournament"
'		SendErrorCode_DisplayStatus = "inline-block"
' END IF	



MessageSentStatus = ""




' -- Sets the page the program opens on --
SELECT CASE Action 
		CASE "message", "templateselectedbyuser", "updatecurrenttemplate", "buildnewtemplate", "savenewtemplateinfo", "insertcurrenttemplate"
				OpenState = "message"
		CASE "backtorecipients"
				OpenState = "recipients"		
		CASE "ontosend"
				OpenState = "send"
		CASE "sendmessage"
				OpenState = "confirm"		
	CASE ELSE
				OpenState = "tournament"
END SELECT

IF Action<> "dispsent" THEN
		DisplayHeadOpenBodyAndBannerTags_EmailMessaging (OpenState)
END IF



' -- Determines which module runs --
SELECT CASE Action

		CASE "dispsent"
				DisplaySentRecipientList


		CASE "templateselectedbyuser"
				TemplateFieldStatus = "disabled"
				
				UpdateTemplateFormDataFromTable
				Display_EmailCMS_Main	

		CASE "buildnewtemplate"
				TemplateFieldStatus = "enabled"
				
				Display_EmailCMS_Main

		CASE "insertcurrenttemplate"

				ReadTemplateFormValues
				
				SaveTemplateToTable
				
				TemplateFieldStatus = "enabled"
				
				Display_EmailCMS_Main

		
		CASE "updatecurrenttemplate"

				' -- Read the values from the browser --			
				ReadTemplateFormValues
				
				' -- Update the values into the Template table --
				UpdateTemplateToTable
				 	
				' -- Read the values back into variables for consistency --
				UpdateTemplateFormDataFromTable					

				' -- Display everything in form --
				TemplateFieldStatus = "disabled"
				Display_EmailCMS_Main				
				

		CASE "ontosend"
		
				'response.write("Line 811 - TemplateIDSelected = "&TemplateIDSelected)
				
				' -- Update variables because there was a form submit from the previous page
				UpdateTemplateFormDataFromTable
				
				' -- Creates a record in Email_Summary table if Pending for this templateID does not exist --
				FindOrCreate_MessageID

				' -- Counts recipients based on current criteria --
				CountRecipients
								
				' --- Updates the send table with PENDING status --
				Update_SendSummary_Table ("Pending")
				
				' -- Display everything --					
				Display_EmailCMS_Main		


		CASE "sendmessage"
				' response.write("<div style=color:red;><br>SEND MESSAGE HERE</div>")

				' -- Update variables because there was a form submit from the previous page
				UpdateTemplateFormDataFromTable
				
				' -- Creates a record in Email_Summary table if Pending for this templateID does not exist --
				FindOrCreate_MessageID
					
				' -- Run the query to send messages to recipients --
				SelectRecipients

				' -- Counts recipients based on current criteria --
				CountRecipients
								
				' -- Send the messages --
				SendMessagesToRecipients
				
				rs.close
				
				
				' -- Update variables because there was a form submit from the previous page
				UpdateTemplateFormDataFromTable
				
				MessageSentStatus = "Sent"
				Display_EmailCMS_Main

				
		CASE ELSE
				UpdateTemplateFormDataFromTable
					
				Display_EmailCMS_Main				
					
END SELECT				



' --- Closing BODY and HTML tags --
DisplayCloseBodyAndHTMLTags




' ---------------------------------------------------------------------------------------
' --- END OF PROGRAM   written by: Mark Crone - Jan 2017 ---
' ---------------------------------------------------------------------------------------




' ---------------------------
  SUB ReadTemplateFormValues
' ---------------------------  

ThisTemplateSubject = TRIM(Request("ThisTemplateSubject"))
ThisSalutation = TRIM(Request("ThisSalutation"))
ThisTemplate_Body = TRIM(Request("ThisTemplate_Body"))

ThisSenderSignature = TRIM(Request("ThisSenderSignature"))
ThisSenderTitle = TRIM(Request("ThisSenderTitle"))
ThisReplyToEmail = TRIM(Request("ThisReplyToEmail")) 

IF TRIM(ThisSenderSignature) = "" AND (TRIM(sTRegistrarName) <> "" OR TRIM(sTDirName) <> "") THEN 
		ThisSenderSignature = sTRegistrarName
		IF TRIM(sTRegistrarName) = "" THEN 
				ThisSenderSignature = sTDirName
				ThisReplyToEmail = sTDirEmail
		END IF 
END IF

IF TRIM(ThisReplyToEmail) = "" AND (TRIM(sTRegistrarEmail) <> "" OR TRIM(sTDirEmail) <> "") THEN 
		ThisReplyToEmail = sTRegistrarEmail
		IF TRIM(sTRegistrarName) = "" THEN 
				ThisReplyToEmail = sTDirEmail
		END IF 
END IF


'response.write("<br>Line 479 - ThisSenderTitle = "&ThisSenderTitle)

END SUB




' --------------------------------
  SUB ValidateAdminCode_ForEmail
' --------------------------------
  
	
	sSQL = "SELECT TOP 1 * FROM "&TRegSetupTableName&" AS TR"
	sSQL = sSQL + " JOIN "&SanctionTableName&" AS ST ON LEFT(TR.TournAppID,6)=LEFT(ST.TournAppID,6)"
	sSQL = sSQL + " LEFT JOIN "&Users999TableName&" AS UT ON LEFT(TR.TournAppID,6)=LEFT(UT.name,6)"
	sSQL = sSQL + " WHERE LEFT(TR.TournAppID,6) = '"&LEFT(sTourID,6)&"'"  

	set rsSanc=Server.CreateObject("ADODB.recordset")
	rsSanc.open sSQL, sConnectionToTRATable, 3, 1



END SUB



' ------------------------------------
  SUB UpdateTemplateFormDataFromTable 
' ------------------------------------

' response.write("Line 977 - TemplateIDSelected = "&TemplateIDSelected)

IF messagelist="template" THEN
		sSQL = " SELECT DISTINCT TemplateID, '"&sTourID&"' AS TourID, TemplateSubject, Created_Date"
		sSQL = sSQL + " , Salutation, Message_Body, '"&ThisTourDirector&"' AS SenderSignature, SenderTitle"
		sSQL = sSQL + " , '"&ThisTourDirEmail&"' AS ReplyToEmail"
		sSQL = sSQL + " FROM usawsrank.Register_Email_Template_Samples as ets"
		sSQL = sSQL + " WHERE TemplateID = '"&TemplateIDSelected&"'"
	
ELSE
		sSQL = "SELECT DISTINCT TemplateID, TourID, TemplateSubject, Created_Date"
		sSQL = sSQL + ", Salutation, Message_Body, SenderSignature, SenderTitle, ReplyToEmail, Created_Date, Created_MemberID" 
		sSQL = sSQL + " FROM "&EmailTemplateTableName&" as et"
		sSQL = sSQL + " WHERE LEFT(et.TourID,6) = '"&LEFT(sTourID,6)&"'"
		sSQL = sSQL + " AND TemplateID = '"&TemplateIDSelected&"'"
		sSQL = sSQL + " ORDER BY rd.Created_Date DESC"	   
END IF



'response.write("<div style=color:red;> Line 997 - sSQL = " &sSQL)
'response.end

ThisTemplateSubject = ""
ThisTemplate_Body = ""
ThisSalutation = ""
ThisSenderSignature = ""
ThisSenderTitle = ""
ThisReplyToEmail = ""

SET rsTemplates=Server.CreateObject("ADODB.recordset")
rsTemplates.open sSQL, SConnectionToTRATable

IF NOT rsTemplates.eof THEN
		ThisTemplateSubject = rsTemplates("TemplateSubject")
		ThisSalutation = rsTemplates("Salutation")
		ThisTemplate_Body = rsTemplates("Message_Body")
		ThisSenderSignature = rsTemplates("SenderSignature")
		ThisSenderTitle = rsTemplates("SenderTitle")
		ThisReplyToEmail = rsTemplates("ReplyToEmail")
END IF	

rsTemplates.close


END SUB




' --------------------------
  SUB SaveTemplateToTable
' --------------------------

messagelist="existing"

' -- Get the Max ID of existing templates -- 
sSQL = "SELECT MAX(TemplateID)+1 AS NextTemplateID FROM "&EmailTemplateTableName
SET rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable
NextTemplateID = rs("NextTemplateID")


' -- Insert the new Template -- 
sSQL = "INSERT INTO "&EmailTemplateTableName
sSQL = sSQL + " (TemplateID, TourID, TemplateSubject, Salutation, Message_Body, SenderSignature, SenderTitle, ReplyToEmail"
sSQL = sSQL + " , Created_Date, Created_MemberID)"  
sSQL = sSQL + " VALUES ('"&NextTemplateID&"', '"&sTourID&"', '"&SQLCleanEmail(ThisTemplateSubject)&"', '"&SQLCleanEmail(ThisSalutation)&"', '"&SQLCleanEmail(ThisTemplate_Body)&"'"
sSQL = sSQL + ", '"&SQLCleanEmail(ThisSenderSignature)&"', '"&SQLCleanEmail(ThisSenderTitle)&"', '"&SQLCleanEmail(ThisReplyToEmail)&"'"
sSQL = sSQL + ", '"& Date() &"', '"&sMemberID&"')"

TemplateIDSelected = NextTemplateID
action="member"

'response.write("<div style=background-color:white; color:red;>Line 1051 - "&sSQL&"</div>")
'response.end

OpenCon
con.execute(sSQL)
CloseCon

SendMarkEmail_SaveTemplate "Save"

END SUB



' -----------------------------
   Function SQLCleanEmail(str)
' ------------------------------

' This function cleans variables to remove any SQL protected symbols which might 
' be used to hack our SQL or crash the program.

Dim tempString

tempString = str
IF tempString <> "" THEN
	' --- A single apostrophe is replaced by double - WORKING ?	
	tempString = replace(tempString,"'","''")

	' --- double pluses 
	tempString = replace(tempString,"++","'")

	' --- semicolon
	tempString = replace(tempString,";","")

	' --- comma
	tempString = replace(tempString,",","")


END IF
SQLCleanEmail = tempString
End Function




' --------------------------
  SUB UpdateTemplateToTable
' --------------------------

sSQL = "UPDATE et"
sSQL = sSQL + " SET TemplateSubject = '"&SQLCleanEmail(ThisTemplateSubject)&"', Salutation = '"&SQLCleanEmail(ThisSalutation)&"', Message_Body = '"&SQLCleanEmail(ThisTemplate_Body)&"'"
sSQL = sSQL + ", SenderSignature = '"&SQLCleanEmail(ThisSenderSignature)&"', SenderTitle = '"&SQLCleanEmail(ThisSenderTitle)&"', ReplyToEmail = '"&SQLCleanEmail(ThisReplyToEmail)&"'"
sSQL = sSQL + ", ModifiedLast = '" & NOW() &"'"
sSQL = sSQL + " FROM "&EmailTemplateTableName&" et"
sSQL = sSQL + " WHERE TemplateID = '"&TemplateIDSelected&"'"

' response.write("<div style=color:red;>Line 1072 - "&sSQL&"</div>")
' response.end

OpenCon
con.execute(sSQL)
CloseCon

SendMarkEmail_SaveTemplate "Update"


END SUB



' --------------------------------------------
  SUB SendMarkEmail_SaveTemplate (subroutine)
' --------------------------------------------

eMailTo = "cronemarka@gmail.com"
eMailCC = ""
eMailBCC = ""
eMailFrom = "competition@usawaterski.org"
eMailReplyTo = "cronemarka@gmail.com"
eMailSubj = "Email Messaging being used - " &sTourID
eMailBody = "Email Messaging " &sTourID& " - "&sTourName& "<br><br>Using Subroutine: "&subroutine&"<br><br>Member: "&sMemberID&"<br><br>ReplyTo: "&eMailReplyTo



SendEmailFromGenericMethodAndReplyTo eMailTo,eMailCC,eMailBCC,eMailFrom,eMailReplyTo,eMailSubj,eMailBody


END SUB



' -------------------------
  SUB CreateHTMLForMessage 
' -------------------------

 		USAWS_Logo ="http://www.usawaterski.org/rankings/images/logos/usawslogo_no_sub.jpg"

    eMailBody = "<HTML><HEAD><TITLE>Message Preview</TITLE></HEAD>"
    eMailBody = eMailBody + "<BODY style='font-family: Arial, Helvetica, sans-serif; text-align:left; font-size:10pt;'>"
		eMailBody = eMailBody + "<div style='width:auto; height:35px; margin-top:40px; padding-top:9px; color:#FFFFFF; background-color:#0000b3; border:1px solid #0000b3; border-radius:20px 20px 0px 0px; text-align:center; font-size:16pt; font-weight:bold;'>Tournament Notification</div>"

    eMailBody = eMailBody + "<div style='width:auto; padding:0px 10px 0px 10px; border:1px solid black; border-radius:0px 0px 20px 20px;'>"

    eMailBody = eMailBody + " <div style='width:100%; margin-top:30px;'>Re: <b>"&ThisTournamentName&"</b></div>"
    eMailBody = eMailBody + "	<div style='width:100%;'>TourID: "&sTourID&"</div>"
    eMailBody = eMailBody + "	<div style='width:100%; margin-top:0px;'>Event Date: "&ThisTourDate&"</div>"
                   
    eMailBody = eMailBody + "	<div style='width:100%; margin-top:20px;'>"&ThisSalutation&":</div>"
    eMailBody = eMailBody + "	<div style='width:100%; margin-top:20px;'>"&ThisTemplate_Body&"</div>"
    eMailBody = eMailBody + "	<div style='width:100%; margin-top:20px;'>Sincerely,</div>"
    eMailBody = eMailBody + "	<div style='width:100%; margin-top:20px;'>"&ThisSenderSignature&"</div>"
    eMailBody = eMailBody + "	<div style='width:100%; margin:0px 0px 30px 0px;'>"&ThisSenderTitle&"</div>" 
		eMailBody = eMailBody + "</div>"

    eMailBody = eMailBody + "<div style='width:100%; text-align:center; font-size:8pt; margin:15px 0px 0px 0px;'>A Service of</div>" 
    eMailBody = eMailBody + "<div style='width:100%; text-align:center; margin:10px 0px 0px 0px;'><img src='"&USAWS_Logo&"' style='width:100px;'></div>" 
    eMailBody = eMailBody + "<div style='width:100%; text-align:center; font-size:8pt; font-style:bold; margin:10px 0px 0px 0px;'>180 Holy Cow Rd<br>Polk City, FL 33883</div>" 

    eMailBody = eMailBody + "</BODY></HTML>"


END SUB



' -----------------------------
  SUB SendMessagesToRecipients
' -----------------------------  


eMailCC = ""
IF SenderCopySelect="cc_all" THEN eMailCC = ThisReplyToEmail
eMailBCC = ""
IF SenderCopySelect="bcc_all" THEN eMailBCC = ThisReplyToEmail
	
eMailFrom = "competition@usawaterski.org"
eMailReplyTo = ThisReplyToEmail
eMailSubj = ThisTemplateSubject

CreateHTMLForMessage



' -- Loop thru the registrants selected and send email --
NumberRecipientsSent = 0
	
IF NOT rs.eof THEN 
		DO WHILE NOT rs.eof

				NumberRecipientsSent = NumberRecipientsSent +1 
				
				tMemberID = rs("MemberID")
				eMailTo = rs("Email")

				' -- Live or Test send
				IF sMemberID="000001151" OR LEFT(sTourID,6)="16W999" THEN Test1x="Y"
				
				IF Test1x="Y" THEN
						' response.write("</div><div style=background-color:white; color:red;> MemberID = "&sMemberID)
						' response.end
						' -- TESTING mode --
						' sTourID = "16W999"
						' - GHVuXYNXHH
						
						eMailTo = "mark.crone@bonniercorp.com"
						eMailCC = ""
						eMailBCC = ""
						SendEmailFromGenericMethodAndReplyTo eMailTo,eMailCC,eMailBCC,eMailFrom,eMailReplyTo,eMailSubj,eMailBody
						UpdateRecipients_SentList MessageID_Found, tMemberID
						EXIT DO
				ELSE
						' -- Normal sending mode
						' SendEmailFromGenericMethodAndReplyTo eMailTo,eMailCC,eMailBCC,eMailFrom,eMailReplyTo,eMailSubj,eMailBody
				END IF
				
				UpdateRecipients_SentList MessageID_Found, tMemberID
				rs.moveNEXT
		LOOP
END IF  

' --- Sends one copy of the message to the Sender --
IF SenderCopySelect="to_1x" THEN 
		eMailTo = ThisReplyToEmail
		eMailBCC = ""
		eMailCC = ""
		SendEmailFromGenericMethodAndReplyTo eMailTo,eMailCC,eMailBCC,eMailFrom,eMailReplyTo,eMailSubj,eMailBody
END IF

' -- Mark gets copy of every message - set to N to turn off --
SendMark="Y"
IF SendMark="Y" THEN
		eMailTo = "cronemarka@gmail.com"
		eMailBCC = ""
		eMailCC = ""
		SendEmailFromGenericMethodAndReplyTo eMailTo,eMailCC,eMailBCC,eMailFrom,eMailReplyTo,eMailSubj,eMailBody
END IF
	
Update_SendSummary_Table ("Sent")



END SUB




' ---------------------
  SUB SelectRecipients
' ---------------------


sSQL = "SELECT MemberID, MAX(FirstName) AS FirstName, MAX(LastName) AS LastName"
sSQL = sSQL + ", MAX(Email) AS Email"
sSQL = sSQL + ", MAX(Div) AS Div"
sSQL = sSQL + ", CASE WHEN SUM(Sl)>0 THEN 'S' ELSE ' ' END AS Slalom"
sSQL = sSQL + ", CASE WHEN SUM(Tr)>0 THEN 'T' ELSE ' ' END AS Trick"
sSQL = sSQL + ", CASE WHEN SUM(Ju)>0 THEN 'J' ELSE ' ' END AS Jump"
sSQL = sSQL + " FROM ("
sSQL = sSQL + "   SELECT rd.MemberID, mt.FirstName, mt.LastName, mt.Email, Div"
sSQL = sSQL + "    , CASE WHEN Event='S' THEN 1 ELSE 0 END AS Sl" 
sSQL = sSQL + "    , CASE WHEN Event='T' THEN 1 ELSE 0 END AS Tr" 
sSQL = sSQL + "    , CASE WHEN Event='J' THEN 1 ELSE 0 END AS Ju" 
sSQL = sSQL + "       FROM "&RegDetailTableName&" rd"
sSQL = sSQL + "   LEFT JOIN "&RegGenTableName&" rg ON rg.TourID=rd.TourID AND rg.MemberID=rd.MemberID"
sSQL = sSQL + "   LEFT JOIN "&MemberShortTableName&" mt ON mt.PersonID = RIGHT(rd.MemberID,8)"
sSQL = sSQL + "   WHERE LEFT(rd.TourID,6) = '"&LEFT(sTourID,6)&"'"
sSQL = sSQL + "     AND rd.Div IN ("&DivSelected_ForSQL&")"
sSQL = sSQL + "     AND rd.Event IN ("&EventSelected_ForSQL&")"
sSQL = sSQL + ") DivEntered"
sSQL = sSQL + "  WHERE LEN(Email)-LEN(REPLACE(Email,'@',''))=1 AND CHARINDEX('..',Email)=0 AND CHARINDEX('.',Email)>0 AND CHARINDEX(')',Email)=0 AND CHARINDEX('/',Email)=0 AND CHARINDEX(':',Email)=0 AND CHARINDEX(' ',Email)=0"
' sSQL = sSQL + "  WHERE LEFT(Email,1)<>' ' AND LEN(Email)>=10 AND charindex('@',Email)>0 AND charindex('.',Email)>0"AND MemberID NOT IN ('500159147

sSQL = sSQL + "  AND MemberID NOT IN ('500159147','600159146','800168993','300174883','000177481','400151784','500151783')"

sSQL = sSQL + " GROUP BY MemberID"
sSQL = sSQL + " ORDER BY LastName"



' --- Add criteria for NOT SENT --

' --- Add criteria for NOT PAID, NOT QUALIFIED and other standard issues --
'response.write("</div><div style=color:red; background-color:white;>Linr 1080 - EventSelected = "&EventSelected&"</div>")

'response.write("<div style=color:red;>Line 1071 - "&sSQL&"</div>")
'response.end
 
SET rs=Server.CreateObject("ADODB.recordset")
rs.open sSQL, SConnectionToTRATable


END SUB




' -------------------
  SUB CountRecipients 
' -------------------  

sSQL = "SELECT COUNT(DISTINCT rd.MemberID) AS Recipient_Count"
sSQL = sSQL + "       FROM "&RegDetailTableName&" rd"
sSQL = sSQL + "   LEFT JOIN "&RegGenTableName&" rg ON rg.TourID=rd.TourID AND rg.MemberID=rd.MemberID"
sSQL = sSQL + "   LEFT JOIN "&MemberShortTableName&" mt ON mt.PersonID = RIGHT(rd.MemberID,8)"
sSQL = sSQL + "   WHERE LEFT(rd.TourID,6) = '"&LEFT(sTourID,6)&"'"
sSQL = sSQL + "     AND rd.Div IN ("&DivSelected_ForSQL&")"
sSQL = sSQL + "     AND rd.Event IN ("&EventSelected_ForSQL&")"
sSQL = sSQL + "     AND LEFT(Email,1)<>' '"

Session("CountRecipients sSQL = "&sSQL)

SET rsCNT=Server.CreateObject("ADODB.recordset")
rsCNT.open sSQL, SConnectionToTRATable


Recipient_Count = 0
IF NOT rsCNT.eof THEN Recipient_Count = rsCNT("Recipient_Count")
rsCNT.close	 

END SUB



' ----------------------------
  SUB DisplaySentRecipientList
' ----------------------------

ThisMessageID = request("mid")

sSQL = "SELECT FirstName, LastName, City, State, MemberID_Recipient"
sSQL = sSQL + " , ss.TemplateSubject, Sent_Date, Message_Body, ss.MessageID"
sSQL = sSQL + " FROM "&EmailSendDetailTableName&" sd"
sSQL = sSQL + " LEFT JOIN "&MemberShortTableName&" m ON m.PersonID=CAST(RIGHT(sd.MemberID_Recipient,8) AS INT)"
sSQL = sSQL + " LEFT JOIN "&EmailSendSummaryTableName&" ss ON ss.MessageID=sd.MessageID"
sSQL = sSQL + " LEFT JOIN "&EmailTemplateTableName&" st ON st.TemplateID=ss.TemplateID"

sSQL = sSQL + " WHERE sd.MessageID ='"&ThisMessageID&"' AND Status='Sent'"

SET rsPrior=Server.CreateObject("ADODB.recordset")
rsPrior.open sSQL, SConnectionToTRATable

'response.write(sSQL)
'response.end

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml" style="height:auto">
<head>
<meta charset="utf-8">
<title>OLR Email Messaging</title>
<link rel="stylesheet" href="css/stylesheet_mob_tours.css" media="screen">
<meta charset="utf-8"> 		
<meta name="apple-touch-fullscreen" content="yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="viewport" content="width=device-width, height=device-height, minimum-scale=1, maximum-scale=1, user-scalable=no, minimal-ui">
<meta name="apple-mobile-web-app-status-bar-style" content="black">
<meta name="apple-mobile-web-app-title" content="AWSA Mob">
<meta name="format-detection" content="telephone=no">
<link rel="apple-touch-icon" href="http://www.usawaterski.org/rankings/images/icons/AWSA_HomeScreen_57.png">
<! '--- For iPad --- ->
<link rel="apple-touch-icon" sizes="72x72" href="http://www.usawaterski.org/rankings/images/icons/AWSA_HomeScreen_57.png">
<! --- For pre-retina iPhone, iPod Touch, and Android 2.1+ devices --- ->
<link rel="apple-touch-icon" href="http://www.usawaterski.org/rankings/images/icons/AWSA_HomeScreen_57.png">
<script language="javascript" type="text/javascript" src="js/view-tours-mobile.js"></script>
<script language="javascript" type="text/JavaScript" src="/jscripts/scripts.js"></script>
<script language="javascript" type="text/javascript" src="/jscripts/swfobject.js"></script>
</head>


<div class="container" style="background-color:#FFFFFF; height:auto;">
	<body style="background:#6666CC; height:auto;">
<%	

reccnt = 0

IF NOT rsPrior.EOF THEN
		DO WHILE NOT rsPrior.EOF 

				reccnt = reccnt + 1

				FirstName = rsPrior("FirstName")
				LastName = rsPrior("LastName")
				City = rsPrior("City")
				State = rsPrior("State") 
				MemberID_Recipient = rsPrior("MemberID_Recipient") 

				MessageID = rsPrior("MessageID")
				TemplateSubject = rsPrior("TemplateSubject")
				Sent_Date = rsPrior("Sent_Date")  				
				Message_Body = rsPrior("Message_Body")
					
				IF reccnt=1 THEN 
						%>
						<br><br>
						<div class="textlabel" style="width:100%; text-align:center; padding:0px 0px 0px 10px; border:0px solid black; font-size:14pt;">Email Message Recipient List</div>
						<div class="textdata" style="width:100%; text-align:center; padding:0px 0px 0px 10px; border:0px solid black; color:red;">(Use the print option in your browser to output)</div> 
						<div class="textlabel" style="width:12%; text-align:left; padding:0px 0px 0px 10px; border:0px solid black;">Tour:</div> 
						<div class="textdata" style="width:45%; text-align:left; padding:0px 0px 0px 10px; border:0px solid black; margin:20px 0px 0px 0px;"><% =sTourName %></div>
						<div class="textlabel" style="width:12%; text-align:left; padding:0px 0px 0px 10px; border:0px solid black;">TourID:</div> 
						<div class="textdata" style="width:25%; text-align:left; padding:0px 0px 0px 10px; border:0px solid black; margin:20px 0px 0px 0px;"><% =sTourID %></div>

						<div class="textlabel" style="width:12%; text-align:left; padding:0px 0px 0px 10px; border:0px solid black;">Sent Date:</div> 
						<div class="textdata" style="width:45%; text-align:left; padding:0px 0px 0px 10px; border:0px solid black;"><% =Sent_Date %></div>
						<div class="textlabel" style="width:12%; text-align:left; padding:0px 0px 0px 10px; border:0px solid black;">MessageID:</div> 
						<div class="textdata" style="width:25%; text-align:left; padding:0px 0px 0px 10px; border:0px solid black;"><% =MessageID %></div>

						<div class="textlabel" style="width:12%; text-align:left; padding:0px 0px 0px 10px; border:0px solid black;">Subject:</div> 
						<div class="textdata" style="width:80%; text-align:left; padding:0px 0px 0px 10px; border:0px solid black;"><% =TemplateSubject %></div>

						<div class="textlabel" style="width:12%; text-align:left; padding:0px 0px 0px 10px; border:0px solid black; margin:20px 0px 0px 0px">Message Text</div>	
						<div class="textdata" style="width:96%; text-align:left; padding:0px 0px 0px 10px; border:0px solid black; min-height:120px; overflow: hidden;"><% =Message_Body %></div>
												
						<div style="width:98%; margin:20px 0px 0px 0px">
							<div class="textlabel" style="width:20%; text-align:center; border:0px solid black;">MemberID</div>
							<div class="textlabel" style="width:30%; text-align:left; padding:0px 0px 0px 10px; border:0px solid black;">Name</div>				
							<div class="textlabel" style="width:20%; text-align:left; padding:0px 0px 0px 10px; border:0px solid black;">City/State</div>
						</div>
						<%
				END IF	
						
				%>
				<div style="width:98%;">
					<div class="textdata" style="background:#FFFFFF; width:20%; text-align:center; border:0px solid black;"><% =MemberID_Recipient %></div>					
					<div class="textdata" style="background:#FFFFFF; width:30%; text-align:left; padding:0px 0px 0px 10px; border:0px solid black;"><% =FirstName %>&nbsp;<% =LastName %></div>				
					<div class="textdata" style="background:#FFFFFF; width:20%; text-align:left; padding:0px 0px 0px 10px; border:0px solid black;"><% =City %>,&nbsp;<% =State %></div>
				</div>
				<%
				
				rsPrior.movenext
		LOOP
ELSE
		%><div style="width:100%;">No Recipients Were Sent</div><%
	
END IF

rsPrior.close

%>
<br><br>
</div>
</body>
</html>
<%



END SUB






' ----------------------------
  SUB FindOrCreate_MessageID
' ----------------------------

' --- Look for this message info in Send Summary file --

sSQL = "SELECT MessageID, Status AS MessageStatus"
sSQL = sSQL + " FROM "&EmailSendSummaryTableName
sSQL = sSQL + " WHERE TemplateID = "&TemplateIDSelected&""
sSQL = sSQL + 		" AND MemberID_Sender = '"&sMemberID&"'"
'sSQL = sSQL + 		" AND Divisions_Selected = '"&DivSelected&"'"
'sSQL = sSQL + 		" AND Events_Selected '"&EventSelected&"'"
sSQL = sSQL + 		" AND Status = 'Pending'"

'response.write("<div style=color:green;>"&sSQL&"</div>")
'response.end

SET rsSend=Server.CreateObject("ADODB.recordset")
rsSend.open sSQL, SConnectionToTRATable


MessageStatus_Found = ""
MessageID_Found = 0

IF NOT rsSend.eof THEN
		MessageStatus_Found = rsSend("MessageStatus")
		MessageID_Found = rsSend("MessageID")
	
ELSE
	
		' -- Add a new MessageID to Summary Table if a Pending message doesn't exist --
		sSQL = "SELECT MAX(MessageID) AS Max_MessageID"
		sSQL = sSQL + " FROM "&EmailSendSummaryTableName
		SET rsNew=Server.CreateObject("ADODB.recordset")
		rsNew.open sSQL, SConnectionToTRATable
		
		MessageID_New = 1000
		IF NOT rsNew.eof THEN MessageID_New = rsNew("Max_MessageID") + 1
		
		sSQL = "INSERT INTO "&EmailSendSummaryTableName
		sSQL = sSQL + " (MessageID, TemplateID, TourID, Status)"
		sSQL = sSQL + " VALUES ("&MessageID_New&", "&TemplateIDSelected&", '"&sTourID&"', 'Pending')"   		
		
		Session("FindOrCreate_MessageID sSQL = "&sSQL)
		'response.write("</div><div style=background-color:white; color:red;>Line 1275"&sSQL&"</div>")
		'response.end
		OpenCon
		con.execute(sSQL)
		CloseCon
		
		MessageID_Found = MessageID_New
		
END IF	

'response.write("<div style=color:red;>MessageID_Found = "&MessageID_Found&"</div>")
'response.end


rsSend.close


END SUB





' -----------------------------------------
  SUB Update_SendSummary_Table (SendStatus)
' -----------------------------------------  

 
sSQL = "UPDATE ss "
sSQL = sSQL + " SET MemberID_Sender = '"&sMemberID&"', ReplyToEmail = '"&ThisReplyToEmail&"'"
sSQL = sSQL + ", TemplateSubject = '"&ThisTemplateSubject&"', Divisions_Selected = '"&DivSelected&"', Events_Selected = '"&EventSelected&"'"
sSQL = sSQL + ", Special_Selected = '"&SpecialSendSelect&"', Status = '"&SendStatus&"'"
IF SendStatus="Sent" THEN sSQL = sSQL + ", Sent_Date = '"&DATE()&"'"
IF SendStatus="Sent" THEN sSQL = sSQL + ", Message_Count = "&NumberRecipientsSent	
sSQL = sSQL + " FROM "&EmailSendSummaryTableName&" ss"
sSQL = sSQL + " WHERE MessageID = "&MessageID_Found
sSQL = sSQL + "    AND Status<>'Sent'" 

Session("Update_SendSummary_Table sSQL = "&sSQL)
'response.write("<div style=color:red;>"&sSQL&"</div>")
'response.end

OpenCon
con.execute(sSQL)
CloseCon
		

END SUB



' ---------------------------------------------------
  SUB UpdateRecipients_SentList (MessageID, sMemberID)
' ---------------------------------------------------  

sSQL = "INSERT INTO "&EmailSendDetailTableName
sSQL = sSQL + " (TourID, MessageID, MemberID_Recipient)" 
sSQL = sSQL + " VALUES ('"&sTourID&"', "&MessageID&", '"&sMemberID&"')"

Session("UpdateRecipients_SentList sSQL = "&sSQL)

'response.write("<div style=color:red;>Line 1334 "&sSQL&"</div>")
'response.end

OpenCon
con.execute(sSQL)
CloseCon



END SUB



' --------------------------------
  SUB Display_PriorMessageListing 
' --------------------------------  

sSQL = "SELECT Sent_Date, MessageID, TemplateID, TemplateSubject, Message_Count, Status"
sSQL = sSQL + " FROM "&EmailSendSummaryTableName
sSQL = sSQL + " WHERE LEFT(TourID,6) = '"&LEFT(sTourID,6)&"'"

SET rsPrior=Server.CreateObject("ADODB.recordset")
rsPrior.open sSQL, SConnectionToTRATable


IF NOT rsPrior.EOF THEN
		DO WHILE NOT rsPrior.EOF 
				'Sent_Date="8/16"

				MessageID = rsPrior("MessageID")
				TemplateID = rsPrior("TemplateID")
				TemplateSubject = LEFT(rsPrior("TemplateSubject"),40)
				Message_Count = rsPrior("Message_Count")
				Message_Status = TRIM(rsPrior("Status")) 
				' response.write("<br>Line 1212 - Sent_Date = "&Sent_Date)
				Sent_Date="99/99/99"
				IF Message_Status="Sent" THEN Sent_Date = rsPrior("Sent_Date")

						
				%>
				<div class="textdata" style="width:13%; text-align:left; padding:0px 0px 0px 10px; border:0px solid black;"><% =Sent_Date %></div>
				<div class="textdata" style="width:63%; text-align:left; padding:0px 0px 0px 5px; border:0px solid black;"><a href="<%=ThisSitePath%>/<%=ThisFileName%>?action=dispsent&mid=<%=MessageID%>" title="Display details of send on <%=Sent_Date%> - MessageID: <%=MessageID%>" target="_blank"><% =TemplateSubject %></a></div>				
				<div class="textdata" style="width:7%; text-align:center; border:0px solid black;"><% =Message_Count %></div>
				<div class="textdata" style="width:12%; text-align:center; border:0px solid black;"><% =LEFT(Message_Status,4) %></div>

				<%
				
				rsPrior.movenext
		LOOP
ELSE
		%><div style="width:100%;">No Prior Messages Found</div><%
	
END IF

rsPrior.close


END SUB












' --------------------------------
  SUB BuildRecipientTextForPreview 
' --------------------------------

' *********************************
' **** MOVE TO SEPARATE .asp ?? ***
' *********************************

RPText = ""
RecipientCountEstimate = 0
DO WHILE NOT rs.EOF 
		ThisName = LEFT(rs("FirstName"),20)& " " &LEFT(rs("LastName"),20)   
		ThisDiv = rs("Div")& " - " &RPText& " - " &ThisName 
		ThisEvent = rs("Event")
		
		' RPText = RPText & ThisDiv& "   "& ThisEvent& "   "& ThisName  	
		' RecipientCountEstimate = RecipientCountEstimate + 1
LOOP

rs.close


END SUB



' -------------------------
  SUB BuildPreviewHTMLLine
' -------------------------

%>
<div style="width:10%"><% =ThisDiv %></div>
<div style="width:10%"><% =ThisEvent %></div>
<div style="width:60%"><% =ThisName %></div>
<%


END SUB




' ----------------------------------------------------------------
  SUB DisplayHeadOpenBodyAndBannerTags_EmailMessaging (OpenState)
' ----------------------------------------------------------------  
  
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<HTML xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta charset="utf-8">
<title>OLR Email Messaging</title>
<link rel="stylesheet" href="css/stylesheet_mob_tours.css" media="screen">
<meta charset="utf-8"> 		
<meta name="apple-touch-fullscreen" content="yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="viewport" content="width=device-width, height=device-height, minimum-scale=1, maximum-scale=1, user-scalable=no, minimal-ui">
<meta name="apple-mobile-web-app-status-bar-style" content="black">
<meta name="apple-mobile-web-app-title" content="AWSA Mob">
<meta name="format-detection" content="telephone=no">
<link rel="apple-touch-icon" href="http://www.usawaterski.org/rankings/images/icons/AWSA_HomeScreen_57.png">
<! '--- For iPad --- ->
<link rel="apple-touch-icon" sizes="72x72" href="http://www.usawaterski.org/rankings/images/icons/AWSA_HomeScreen_57.png">
<! --- For pre-retina iPhone, iPod Touch, and Android 2.1+ devices --- ->
<link rel="apple-touch-icon" href="http://www.usawaterski.org/rankings/images/icons/AWSA_HomeScreen_57.png">
<script language="javascript" type="text/javascript" src="js/view-tours-mobile.js"></script>
<script language="javascript" type="text/JavaScript" src="/jscripts/scripts.js"></script>
<script language="javascript" type="text/javascript" src="/jscripts/swfobject.js"></script>
</head>
<%
' response.write("OpenState = "&OpenState)
' response.end
%>
<body onLoad="javascript:UpdateTabDisplay('<%=OpenState%>');">
<%



END SUB	





' --------------------------
  SUB Display_EmailCMS_Main
' --------------------------

Dim ContentStatus, AWSA_Logo, SponsorHeader
Dim UserReleaseAcceptanceMessage, UserReleaseAcceptanceMessage_DisplayStatus

UserReleaseAcceptanceMessage = "By accepting, you agree to use this email messaging platform only for purposes related to the successful hosting of a USA Water Ski sanctioned event.  The sending Member and LOC shall be in good standing and must have authorization to use the ReplyTo email address provided." 
UserReleaseAcceptanceMessage_DisplayStatus = "enabled"


ContentStatus = "enabled"
AWSA_Logo = "AWSA_Oval_BlueSquare_197x83.png"
SponsorHeader = "VisitCentralFlorida.png"

' --- Displays Banner Line --- 

%>
<div class="container" style="border:0px solid red;">
	<a name="TopTop" title="Page Navigation"></a>
	<div id="bannerheader" style="width:100%; background-color:<%=HQSiteColor2%>; height:70px; margin:0px; padding:0px; border:0px solid white;">
		<a href='<%=ThisSitePath%>/<%=MenuFilename%>' title="Rankings" style="text-decoration:none;" >
			<span class="span45" style="width:45%; height:100%; border:0px solid white;">
				<img src="images/logos/<%=AWSA_Logo%>" style="height:57px; margin:7px 0px 0px 3px; padding:0px 0px 0px 0px; border:0px solid green;" alt="AWSA New Logo" />
			</span>
			<span class="span55" style="width:50%; height:100%; vertical-align:top; text-align:center; padding:0px 0px 0px 0px; margin:0px 0px 0px 0px; border:0px solid white;">
					<img src="images/logos/<%=SponsorHeader%>" style="width:135px; margin:13px 0px 0px 13px; border:1px solid green;" alt="Banner Ad" />
					<span class="span95" style="text-align:center; padding:5px 0px 0px 0px; margin:0px 0px 0px 8px; color:#FFFFFF; border:0px solid red;">Share Life On The Water</span>
			</span>
		</a>	
	</div>


<! ---------------------- ->
<! -- For javascript only ->
<! ---------------------- ->	
<input type="hidden" id="ThisTournamentName" value="<% =ThisTournamentName %>">
<input type="hidden" id="sTourID" value="<% =sTourID %>">
<input type="hidden" id="ThisTourDate" value="<% =ThisTourDate %>">





	<! ----------------- ->	
	<! -- TOURNAMENT Tab ->
	<! ----------------- ->		
	<div class="accordionheader" id="EmailCMS_TournamentHeader">
		<div class="textdata" style="width:100%; text-align:center; padding:7px 0px 0px 0px; margin-left:0px; color:yellow; border:0px solid red; font-size:11pt;"><% =LEFT(ThisTournamentName,45) %></div>
	</div>
	
	<div class="accordionbody" id="EmailCMS_TournamentBody">
		<div class="textlabel">TourID:</div> 
		<div class="textdata" style="width:30%;"><% =sTourID %></div>	

		<div class="textlabel">Date:</div> 
		<div class="textdata" style="width:30%;"><% =ThisTourDate %></div>			

		<div class="textlabel">Tour Dir:</div> 
		<div class="textdata" id="ThisTourDirector"><% =ThisTourDirector %></div>	

		<div class="textlabel">Email:</div> 

		<div class="textdata" id="ThisTourDirEmail"><% =ThisTourDirEmail %></div>

		<div class="textlabel">Mem ID:</div> 
		<div class="textdata" style="width:20%;"><% =sMemberID %></div>	

		<div style="width:98%;">						
			<div class="textlabel" style="width:100%; text-align:left; margin-top:20px; padding:0px 0px 0px 10px; border:0px solid black;">Messages previously sent</div> 
			<div class="textlabel" style="width:13%; text-align:left; padding:0px 0px 0px 10px; margin:0px; border:0px solid black;">Date</div> 
			<div class="textlabel" style="width:63%; text-align:left; padding:0px 0px 0px 5px; margin:0px; border:0px solid black;">Subject</div> 		
			<div class="textlabel" style="width:7%; text-align:center; padding:0px 0px 0px 0px; border:0px solid black;">Qty</div> 
			<div class="textlabel" style="width:12%; text-align:center; border:0px solid black;">Status</div>
			<hr style="margin:0px;"> 		
		</div>
		<div class="scroll" style="height:100px; width:99.8%; padding:0px 0px 0px 0px; border:0px solid black; display:inline-block;">
				<%
				' -- Displays Listing of pending and sent message summary --
				Display_PriorMessageListing
				%>	
		</div>	

   			
		<div class="textdata" style="margin-top:20px; width:100%; height:20px; text-align:center; font-size:12pt; display:<%= UserReleaseAcceptanceMessage_DisplayStatus %>;">TERMS OF USE</div> 		
		<div style="margin:0px; padding:0px 10px 10px 10px; width:100%; text-align:center; font-size:8pt; border:0px solid red; -moz-box-sizing: border-box; -webkit-box-sizing: border-box; -ms-box-sizing: border-box; display:<%= UserReleaseAcceptanceMessage_DisplayStatus %>;">
			<%=UserReleaseAcceptanceMessage%>
		</div> 		

		<div class="buttonrowemail">
			<div class="textdata" style="width:20%; height:28px; margin:0px; vertical-align:bottom;">
				<input class="yellowbutton" type="button" value="Validate" id="TournamentValidate" style="display:none;" onclick="javascript:ValidateUseAcceptance();">		
				<input class="yellowbutton" type="button" value="Continue" id="TournamentContinue" disabled onclick="javascript:UpdateTabDisplay('message');">		
			</div>
 			<div class="textdata" style="width:30%; height:28px;">
				<input type="checkbox" id="AcceptRelease" name="AcceptRelease" onclick="javascript:ValidateUseAcceptance();"> Check to Accept Terms
			</div>
			<div class="textdata" style="width:20%; height:28px; margin:0px; vertical-align:bottom;">
				<input class="stdbutton" type="button" value="Getting Started" id="GettingStarted" onclick="javascript:WhereToStartHelp();">		
			</div>	    						
		</div>
	
	</div>	




<! ----------------- ->
<! -- MESSAGE TAB -- ->	
<! ----------------- ->		

	<! -- MESSAGE HEADER -- ->	
	<div class="accordionheader" id="EmailCMS_MessageHeader" style="display:inline-block;">
		<div class="textlabel" style="width:20%; padding:7px 0px 0px 10px; color:white; border:0px solid red; font-size:10pt;">Content: </div>
		<div class="textdata" style="width:70%; padding:7px 0px 0px 5px; color:yellow; border:0px solid yellow; text-align:left; font-size:9pt;"><div id="TemplateSubjectInHeader"><% =LEFT(ThisTemplateSubject,40) %></div></div>		
	</div>
	
	<! -- MESSAGE FORM -- ->
<form method="post" id="MessageForm" action="" style="margin:0px; padding:0px;">
	
	<input type="hidden" name="ThisTemplateID" id="ThisTemplateID" value="<% =TemplateIDSelected %>">

	<div class="accordionbody" id="EmailCMS_MessageBody" style="display:none;">


		
		<div class="textlabel">Which messages?</div>
		<div class="textdata" style="height:25px; vertical-align:middle;">				
			<input type="radio" name="messagelist" id="messagelist_existing" value="existing" title="Choose from messages previously created for this tournament" onclick="javascript:UpdateTemplateFormAction('templateselectedbyuser'); submit();" <% IF messagelist="existing" THEN  response.write("checked") %>>Existing &nbsp;&nbsp;
			<input type="radio" name="messagelist" id="messagelist_template" value="template" title="Select template then Copy Template to create a new message for this tournament" onclick="javascript:UpdateTemplateFormAction('templateselectedbyuser'); submit();" <% IF messagelist="template" THEN  response.write("checked") %>>Templates
		</div>
				
		<div class="textlabel">Subject:</div> 

		<div class="textdata" style="height:25px; vertical-align:middle;">
				<%
				Build_MessageList_Dropdown			' -- Load Message List Dropdown --
				%>
				<input type="text" class="textbox" name="ThisTemplateSubject" id="ThisTemplateSubject" value="<% =ThisTemplateSubject %>" title="Subject Line of message - keep to 35 characters for mobile devices" style="display:none;" size="45" maxlength="50" onchange="javascript:ValidateMessageForm();" placeholder="Subject Line - Max 45 char" <% =TemplateFieldStatus %>>
		</div>

		<div class="textlabel" style="padding:3px 0px 0px 0px;">Salute:</div> 
		<div class="textdata" style="vertical-align:bottom; border:0px solid red;">
			<input type="text" class="textbox" name="ThisSalutation" id="ThisSalutation" value="<% =ThisSalutation %>" title="How you want to address the message - typically Dear Skiers" size="30" height="20px" maxlength="30" placeholder="Salutation - Max 30 char" <% =TemplateFieldStatus %>>
		</div>
	
		<div class="textlabel" style="padding:3px 0px 0px 0px;">Signature Line1:</div> 
		<div class="textdata" style="vertical-align:bottom;">
			<input type="text" class="textbox" name="ThisSenderSignature" id="ThisSenderSignature" value="<% =ThisSenderSignature %>" title="Email will be signed Sincerely <signature>" size="30" maxlength="30" placeholder="Name for Signature - Max 30 char" <% =TemplateFieldStatus %>>
		</div>	
		<div class="textlabel" style="padding:3px 0px 0px 0px;">Signature Line2:</div> 
		<div class="textdata" style="vertical-align:bottom;">
			<input type="text" class="textbox" name="ThisSenderTitle"  id="ThisSenderTitle" value="<%=ThisSenderTitle%>" title="Your title for this tournament - ex: Tournament Director" size="45" maxlength="50" placeholder="Title - Max 50 char" <% =TemplateFieldStatus %>>
		</div>	
		
		<div class="textlabel" style="margin-top:10px; padding:3px 0px 0px 0px;" title="Email address where replies are to be sent">Reply To:</div> 
		<div class="textdata" style="margin-top:10px; vertical-align:bottom;">
			<input type="text" class="textbox" name="ThisReplyToEmail" id="ThisReplyToEmail" value="<%=ThisReplyToEmail%>" title="Email address where replies are to be sent" size="45" maxlength="50" placeholder="Reply To Email - Max 50 char" <% =TemplateFieldStatus %>>
		</div>			
	
	
		<div class="textlabel" style="vertical-align:top; padding:0px 0px 0px 20px; margin:10px 0px 0px 0px; width:50%; text-align:left; border:0px solid black;">Message Content:</div>
		<div class="scroll" style="height:150px; width:95%; margin:0px; 0px 0px 0px; padding:0px 10px 0px 10px; border:0px solid black; display:inline-block; -moz-box-sizing:border-box; -webkit-box-sizing:border-box; -ms-box-sizing:border-box;">
			<textarea name="ThisTemplate_Body" id="ThisTemplate_Body" style="width:100%;" placeholder="Write your message here.  Use HELP to learn about HTML markup options" rows=10 wrap=virtual maxlength=999 <% =TemplateFieldStatus %>><% =ThisTemplate_Body %></textarea>
		</div>			

	
		<div class="buttonrowemail">
			<div style="width:30%; display:inline-block; padding:0px; margin:0px;">
				<input class="stdbutton" type="button" value="Edit" id="MessageEditTemplate" title="Edit the details of this message" onclick="javascript:UpdateTemplateFormAction('edittemplate');">
				<input class="stdbutton" type="submit" value="Update" id="MessageUpdateTemplate" style="display:none;" onclick="javascript:ValidateTemplateCriteria();">						
				<input class="stdbutton" type="submit" name="TemplateButton" value="Save" title="Save the changes to this message" id="MessageSaveTemplate" style="display:none;">						
			</div>
			<div style="width:30%; display:inline-block;">
				<input class="stdbutton" type="button" value="Preview" id="MessagePreviewTemplate" title="Preview how the message will look when it is assembled" onclick="javascript:ShowMessagePreview();">
				<input class="stdbutton" type="button" value="Cancel" id="MessageCancelTemplate" style="display:none;" onclick="javascript:UpdateTemplateFormAction('canceledittemplate');">						
			</div>
			<div style="width:30%; display:inline-block;">
				<input class="stdbutton" type="button" value="New Message" id="MessageNewTemplate" title="Create a NEW message from scratch for this tournament" onclick="javascript:UpdateTemplateFormAction('newtemplate');">
				<input class="greenbutton" type="button" value="Copy Template" id="MessageCopyTemplate" style="display:none;" title="Copies this Master Template into this tournament where it may be edited before sending" onclick="javascript:UpdateTemplateFormAction('copytemplate');">					
			</div>
		</div>
								
		<div class="buttonrowemail">
			<div style="width:30%; display:inline-block;">
				<input class="yellowbutton" type="button" value="Validate" id="ValidateTemplate" onclick="javascript:ValidateTemplateCriteria();">
				<input class="greenbutton" type="button" value="Continue" id="MessageContinue" style="display:none;" onclick="javascript:UpdateTabDisplay('recipients');">								
			</div>		
			<div style="width:30%; display:inline-block;">
				<input class="stdbutton" type="button" value="Back"  id="MessageBack" onclick="javascript:UpdateTabDisplay('tournament');">
			</div>	
			<div style="width:30%; display:inline-block;">
				<input class="stdbutton" type="button" value="Help" id="MessageFormat" onclick="javascript:ShowFormatHelp();">
			</div>		
		</div>		
	</div>	<! -- EmailCMS_MessageBody ->
</form>

<% 


StatusOfRecipientHeader = "Criteria Not Defined"
IF TRIM(DivSelected) <> "" AND TRIM(EventSelected) <> "" THEN 
		StatusOfRecipientHeader = "Selected"
END IF


%>
<! ------------------- ->
<! -- RECIPIENT TAB -- ->	
<! ------------------- ->		

	<! -- RECIPIENT LIST ->
	<div class="accordionheader" id="EmailCMS_RecipientHeader" style="display:inline-block;">
		<div class="textlabel" style="width:20%; padding:7px 0px 0px 10px; color:white; font-size:10pt; border:0px solid red;">Recipient:</div>
		<div class="textdata" style="width:70%; padding:7px 0px 1px 5px; color:yellow; border:0px solid yellow; text-align:left; font-size:9pt;"><% =StatusOfRecipientHeader %></div>	
	</div>


<form method="post" id="recipientform" style="margin:0px; padding:0px;"> 
	<input type="hidden" name="TemplateIDSelected" value="<%= TemplateIDSelected %>">	
	<input type="hidden" name="messagelist" value="existing">	
		
	<div class="accordionbody" id="EmailCMS_RecipientBody" style="display:none;">
		<div class="textlabel">Subject:</div> 
		<div class="textdata"><%=ThisTemplateSubject%></div>

		<div class="textlabel" style="vertical-align:top;">TIP:</div> 
		<div class="textlabel" style="width:75%; height:22px; padding:0px 0px 0px 10px; color:red; text-align:left; border:0px solid black;">Select Divs & Events. [cntl]-Click (on PC) for multiple. Check buttons verify selections.
		</div> 
		<br>

		<! ** DIV ** ->
		<div class="textlabel" style="vertical-align:top; margin:20px 0px 0px 0px; border:0px solid black">Div:</div>
		<div class="textdata" style="width:45%; border:0px solid green; margin:20px 0px 0px 0px;">
			<%
			LoadDivDrop_ForEntered			' -- Load Division Dropdown --
			%>
		</div>
		<div style="width:30%; display:inline-block; text-align:left; vertical-align:top; margin:20px 0px 0px 0px; padding:0px 0px 0px 0px;">
			<input class="stdbutton" type="button" value="Check Divisions" id="VerifyDivs" title="Press to see the list of Divisions you have selected" onclick="javascript:ConfirmSelectedElement('Div');">
		</div>	


		<! ** EVENT ** ->		
		<div class="textlabel" style="vertical-align:top; margin:20px 0px 0px 0px;">Event:</div> 
		<div class="textdata" style="width:45%; margin:20px 0px 0px 0px;">
			<%
			LoadEventDrop_ForEntered			' -- Load Event Dropdown --
			%>
		</div>
		<div style="width:30%; display:inline-block; text-align:left; vertical-align:top; margin:20px 0px 0px 0px; padding:0px 0px 0px 0px;">
			<input class="stdbutton" type="button" value="Check Events" id="VerifyEvents" title="Press to see the list of Events you have selected" onclick="javascript:ConfirmSelectedElement('Event');">
		</div>	


		<! ** SPECIAL ** ->		
		<div class="textlabel" style="margin:20px 0px 0px 0px;">Special:</div> 
		<div class="textdata" style="width:80%; margin:20px 0px 0px 0px;">
			<%
			Build_MessageSelect_Dropdown			' -- Load Special Select Dropdown --
			%>
		</div>
	
		<div class="buttonrowemail">
			<div style="width:30%; display:inline-block;">
				<input class="greenbutton" type="submit" value="Continue" id="RecipientContinue" style="display:none;" formaction="<%=ThisSitePath%>/<%=ThisFileName%>?action=OnToSend">
				<input class="yellowbutton" type="button" value="Validate" id="ValidateRecipient" onclick="javascript:ValidateRecipientCriteria();">					
			</div>	
			<div style="width:30%; display:inline-block;">
				<input class="stdbutton" type="button" value="Back"  id="RecipientBack" onclick="javascript:UpdateTabDisplay('message');">	
			</div>
			<div style="width:30%; display:inline-block;">
				<input class="stdbutton" type="button" value="List" id="RecipientPreviewRecipients" title="List of Recipients function - NOT ACTIVATED" onclick="">		
			</div>
		</div>		
	</div>	<! -- EmailCMS_RecipientBody ->
</form>	




<! ------------------- ->
<! --   SEND TAB    -- ->	
<! ------------------- ->		

	<! -- SEND MESSAGE ->

	<div class="accordionheader" id="EmailCMS_SendHeader" style="display:inline-block;">
		<div class="textlabel" style="width:20%; padding:7px 0px 0px 10px; color:white; font-size:10pt; border:0px solid red;">Sending:</div>		
		<div class="textdata" style="width:70%; padding:7px 0px 1px 5px; color:yellow; border:0px solid yellow; text-align:left; font-size:9pt;">Template ID - <% =TemplateIDSelected %></div>
	</div>

<form method="post" style="margin:0px; padding:0px;">
	<input type="hidden" name="sTourID" value="<% =sTourID %>">
	<input type="hidden" id="Send_TemplateIDSelected" name="TemplateIDSelected" value="<%= TemplateIDSelected %>">	
	<input type="hidden" id="Send_DivSelected" name="DivSelected" value="<%= DivSelected %>">
	<input type="hidden" id="Send_EventSelected" name="EventSelected" value="<%= EventSelected %>">
	<input type="hidden" id="Send_SpecialSendSelect" name="SpecialSendSelect" value="<%= SpecialSendSelect %>">
	<input type="hidden" name="messagelist" value="existing">		
		 	
	<div class="accordionbody" id="EmailCMS_SendBody" style="display:none;">

		<div class="textlabel">Subject:</div> 
		<div class="textdata"><%= ThisTemplateSubject %></div>
		
		<div class="textlabel" style="margin-top:5px;">Divs:</div> 
		<div class="textdata" style="width:70%; margin-top:5px;"><% =DivSelected %></div>		
	
		<div class="textlabel">Events:</div> 
		<div class="textdata" style="width:70%;"><% =EventSelected %>&nbsp;</div>		

		<div class="textlabel">Special:</div> 
		<div class="textdata" style="width:70%;"><% =SpecialSendSelect %>&nbsp;</div>	
				
		<div class="textlabel">Quantity:</div> 
		<div class="textdata" style="width:70%;"><% =Recipient_Count %>&nbsp;<black> Recipients (Estimated)</black></div>		

		<div class="textlabel" style="margin-top:10px;">Sender:</div> 
		<div class="textdata" style="margin-top:10px;"><%=ThisSenderSignature%></div>

		<div class="textlabel">Reply To:</div> 
		<div class="textdata"><%=ThisReplyToEmail%></div>		

		<div class="textlabel">From:</div> 
		<div class="textdata">competition@usawaterski.org</div>		
				
		<! ** SPECIAL ** ->		
		<div class="textlabel" style="margin:0px 0px 0px 0px;">Want A Copy?:</div> 
		<div class="textdata" style="width:80%; margin:0px 0px 0px 0px;">
			<%
			Build_SenderCopySelect_Dropdown			' -- Load SenderCopy Dropdown --
			%>
		</div>		

		<div class="textlabel" style="margin:15px 0px 0px 0px; border:0px solid red;">Mem ID:</div>
		<div class="textdata" style="width:30%; margin:15px 0px 0px 10px; padding:0px 0px 0px 0px; border:0px solid green;"><% =sMemberID %></div>

		<div class="textlabel" style="margin-top:10px; color:red; display:<%= SendErrorCode_DisplayStatus %>;">ERROR:</div>
		<div class="textdata" style="margin-top:10px; width:70%; display:<%= SendErrorCode_DisplayStatus %>;"><%=SendErrorMessage%></div> 		

	
		<div class="buttonrowemail">
			<div style="width:45%; display:inline-block;">
				<input class="yellowbutton" type="button" value="Validate" id="ValidateSend" onclick="javascript:ValidateSendCriteria();">
				<input class="redbutton" type="submit" value="Send Now" id="SendMessageNow" formaction="<%=ThisSitePath%>/<%=ThisFileName%>?action=SendMessage" style="display:none;">				
			</div>		
			<div style="width:45%; display:inline-block;">
				<input class="stdbutton" type="submit" value="Back" id="SendBack" formaction="<%=ThisSitePath%>/<%=ThisFileName%>?action=BackToRecipients" onclick="javascript:DeactivateSendButton();">	
			</div>
			<div style="width:30%; display:inline-block;">&nbsp;</div>
		</div>		
	</div>	<! -- EmailCMS_SendBody ->	
</form>	



	<! -- CONFIRM SEND ->

	<div class="accordionheader" id="EmailCMS_ConfirmHeader" style="display:inline-block;">
		<div class="textlabel" style="width:20%; padding:7px 0px 0px 10px; color:white; font-size:10pt; border:0px solid red;">Confirmation</div>
		<div class="textdata" style="width:70%; padding:7px 0px 1px 5px; color:yellow; border:0px solid yellow; text-align:left; font-size:9pt;">&nbsp;<% =MessageSentStatus %></div>			
	</div>

<form method="post" style="margin:0px; padding:0px;">
	<input type="hidden" name="sTourID" value="<% =sTourID %>">
	<input type="hidden" id="Confirm_TemplateIDSelected" name="TemplateIDSelected" value="<%= TemplateIDSelected %>">	
	<input type="hidden" id="Confirm_DivSelected" name="DivSelected" value="<%= DivSelected %>">
	<input type="hidden" id="Confirm_EventSelected" name="EventSelected" value="<%= EventSelected %>">
	<input type="hidden" id="Confirm_SpecialSendSelect" name="SpecialSendSelect" value="<%= SpecialSendSelect %>">
	
		 	
	<div class="accordionbody" id="EmailCMS_ConfirmBody" style="display:none;">
		<div class="textlabel">Subject:</div> 
		<div class="textdata"><%= ThisTemplateSubject %></div>
		
		<div class="textlabel">Reply To:</div> 
		<div class="textdata"><%=ThisReplyToEmail%></div>		

		<div class="textlabel" style="margin-top:20px;">Divs:</div> 
		<div class="textdata" style="width:70%; margin-top:20px;"><% =DivSelected %></div>		
	
		<div class="textlabel">Events:</div> 
		<div class="textdata" style="width:70%;"><% =EventSelected %></div>		

		<div class="textlabel">Special:</div> 
		<div class="textdata" style="width:70%;"><% =SpecialSendSelect %></div>	
				
		<div class="textlabel">Recipients:</div> 
		<div class="textdata" style="width:70%;"><% =NumberRecipientsSent %> Sent</div>		
	
	
		<div class="buttonrowemail">
			<div style="width:45%; display:inline-block;">
				<input class="stdbutton" type="submit" value="Back To Start" id="Complete" formaction="<%=ThisSitePath%>/<%=ThisFileName%>" style="width:13em;">
			</div>		
		</div>		
	</div>	<! -- EmailCMS_SendBody ->	
</form>	

<%		

END SUB








' -----------------------------------
  SUB DisplayCloseBodyAndHTMLTags
' -----------------------------------

%>
</div><! -- Container -- ->
</body>
</html>
<%

END SUB	



' --------------------------------
  SUB Build_MessageSelect_Dropdown
' --------------------------------
%>
<SELECT name='SpecialSendSelect' id='SpecialSendSelect' <%=SpecialSendSelectStatus%> style="width:14em; color:red; background-color:#FFF8DC;" disabled>
	<option value="All" <%IF SpecialSendSelect = "All" THEN Response.Write(" selected ")%>>All</option><br>
	<option value="Not Sent" <%IF SpecialSendSelect = "Not Sent" THEN Response.Write(" selected ")%>>Not Previously Sent</option><br>
	<option value="Not Paid" <%IF SpecialSendSelect = "Not Paid" THEN Response.Write(" selected ")%>>Not Paid (Standard)</option><br>
</SELECT>
<%

END SUB



' -----------------------------------
  SUB Build_SenderCopySelect_Dropdown
' ------------------------------------
%>
<SELECT name='SenderCopySelect' id='SenderCopySelect' <%=SpecialSendSelectStatus%> title="Select where you want " style="width:14em; color:red;  background-color:#FFF8DC;">
	<option value="" <%IF SenderCopySelect = "" THEN Response.Write(" selected ")%>>Select Option</option><br>
	<option value="to_1x" <%IF SenderCopySelect = "to_1x" THEN Response.Write(" selected ")%>>Sender receives 1 copy per mailing</option><br>
	<option value="cc_all" <%IF SenderCopySelect = "cc_all" THEN Response.Write(" selected ")%>>Sender is CC'd on every email sent</option><br>
	<option value="bcc_all" <%IF SenderCopySelect = "bcc_all" THEN Response.Write(" selected ")%>>Sender is BCC'd on every email sent</option><br>
	<option value="None" <%IF SenderCopySelect = "None" THEN Response.Write(" selected ")%>>No Copies to sender</option><br>
</SELECT>
<%

END SUB



'----------------------------------------------------------------------------------------------
 SUB LoadDivDrop_ForEntered
'----------------------------------------------------------------------------------------------


' -- Loads applicable divisions into a division pulldown for those entered in tournament --


sSQL = "SELECT DISTINCT rd.div, dt.div_name"
sSQL = sSQL + " FROM "&RegDetailTableName&" as rd"
sSQL = sSQL + " LEFT JOIN "&DivisionsTableName&" dt ON dt.div=rd.div" 
sSQL = sSQL + " WHERE dt.SkiYearID=1"
sSQL = sSQL + " AND	LEFT(rd.TourID,6) = '"&LEFT(sTourID,6)&"'"
sSQL = sSQL + " ORDER BY rd.div"	   

SET rsDivisions=Server.CreateObject("ADODB.recordset")
rsDivisions.open sSQL, SConnectionToTRATable

'response.write("<div style=color:red;><br> DivSelected = " & DivSelected & "</div>")
'response.write("<div style=color:red;><br>  Line 744 TEST = " & Instr(DivSelected,"W8"))
'response.end

'  onchange="javascript:UpdateSelectedText('Div');


%>
<select name='DivSelected' id='DivSelected' <%=DivDropStatus%> multiple size=3 style="width:14em"">
<%

IF NOT rsDivisions.eof THEN 
  	rsDivisions.movefirst

  	DO WHILE NOT rsDivisions.eof
  			IF Instr(DivSelected,TRIM(rsDivisions("Div"))) > 0 THEN
				' IF TRIM(rsDivisions("Div")) = DivSelected THEN 
						%>
						<option value="<%=rsDivisions("Div")%>" selected><%=rsDivisions("Div")%> - <%=rsDivisions("Div_Name")%></option><br>
						<%
				ELSE 
						%>
						<option value="<%=rsDivisions("Div")%>"><%=rsDivisions("Div")%> - <%=rsDivisions("Div_Name")%></option><br>
						<%
				END IF	

				rsDivisions.moveNEXT
	LOOP
ELSE
		%>
		<option value = "" selected>No Registrations</option>
		<%
END IF  
%>
</select>
<%

rsDivisions.close

END SUB





'----------------------------------------------------------------------------------------------
 SUB LoadEventDrop_ForEntered
'----------------------------------------------------------------------------------------------


' -- Loads applicable divisions into a division pulldown for those entered in tournament --


sSQL = "SELECT DISTINCT rd.event"
sSQL = sSQL + " , CASE WHEN rd.event='S' THEN 'Slalom'" 
sSQL = sSQL + "       WHEN rd.event='T' THEN 'Trick'"
sSQL = sSQL + "       WHEN rd.event='J' THEN 'Jump'"
sSQL = sSQL + " END AS Event_Name"
sSQL = sSQL + " FROM "&RegDetailTableName&" as rd"
sSQL = sSQL + " WHERE LEFT(rd.TourID,6) = '"&LEFT(sTourID,6)&"'"
sSQL = sSQL + " ORDER BY rd.event"	   

'response.write("<div style=background-color:white; color:#000000;>" &sSQL)
'response.end

SET rsEvent=Server.CreateObject("ADODB.recordset")
rsEvent.open sSQL, SConnectionToTRATable


' onchange="javascript:UpdateSelectedText('Event');"

%>
<select name='EventSelected' id='EventSelected' <%=EventDropStatus%> multiple size=3 style="width:14em">
<%

s=0
IF NOT rsEvent.eof THEN 
  	rsEvent.movefirst

  	DO WHILE NOT rsEvent.eof
				s=s+1
				IF Instr(EventSelected,TRIM(rsEvent("Event"))) > 0 THEN 
						%>
						<option value="<%=rsEvent("Event")%>" selected><%=rsEvent("Event")%> - <%=rsEvent("Event_Name")%></option><br>
						<%
				ELSE 
						%>
						<option value="<%=rsEvent("Event")%>"><%=rsEvent("Event")%> - <%=rsEvent("Event_Name")%></option><br>
						<%
				END IF	

				rsEvent.moveNEXT
	LOOP
ELSE
		%>
		<option value = "" selected>No Registrations</option>
		<%
END IF  
%>
</select>
<%

rsEvent.close

END SUB


' -- Dynamic URL in PADI --




'----------------------------------------------------------------------------------------------
 SUB Build_MessageList_Dropdown
'----------------------------------------------------------------------------------------------

' response.write("</div><div style=background-color:white; color:red;> Line 1909 - TemplateIDSelected = "&TemplateIDSelected&"</div>")

' IF TemplateIDSelected>"90000000" THEN
IF messagelist="template" THEN
		sSQL = " SELECT DISTINCT TemplateID, '"&sTourID&"' AS TourID, TemplateSubject, Created_Date"
		sSQL = sSQL + " FROM usawsrank.Register_Email_Template_Samples as ets"
ELSE
		sSQL = " SELECT DISTINCT TemplateID, TourID, TemplateSubject, Created_Date"
		sSQL = sSQL + " FROM "&EmailTemplateTableName&" as et"
		sSQL = sSQL + " WHERE LEFT(TourID,6) = '"&LEFT(sTourID,6)&"'"
END IF	
sSQL = sSQL + " ORDER BY Created_Date DESC"	   
	 
SET rsTemplates=Server.CreateObject("ADODB.recordset")
rsTemplates.open sSQL, SConnectionToTRATable

' response.write("<div style=height:150px; font-size:8pt; color:red; background-color:white;>" &sSQL)
' response.end


%>
<select name='TemplateIDSelected' id='TemplateIDSelected' <%=TemplateDropStatus%> title="Subject Line of message - keep to 35 characters for mobile devices" style="width:25em; padding:0px; margin:0px; background-color:#FFF8DC;" onchange="javascript:UpdateTemplateFormAction('templateselectedbyuser'); submit();">
<%
IF messagelist="existing" THEN 
		%><option value="">Select Existing Message</option><%
ELSE
		%><option value="">Select From Templates</option><%
END IF

 			
IF NOT rsTemplates.eof THEN 
  	rsTemplates.movefirst

  	DO WHILE NOT rsTemplates.eof
				IF TRIM(rsTemplates("TemplateID")) = TRIM(TemplateIDSelected) THEN 
						%>
						<option value="<%=rsTemplates("TemplateID")%>" selected><% =rsTemplates("TemplateSubject") %></option><br>
						<%
				
				ELSE 
						%>
						<option value="<%=rsTemplates("TemplateID")%>"><%=rsTemplates("TemplateSubject")%></option><br>
						<%
				END IF	

				rsTemplates.moveNEXT
	LOOP
END IF  
%>
</select>
<%

rsTemplates.close

END SUB





' ---------------------------
  SUB Display_TemplatePreview
' ---------------------------

  


END SUB




' -------------------
  SUB CreateEmailHTML
' -------------------

SQT = "'"
ecss = "<style type=text/css>"
ecss = ecss & " body { font-family: Arial, Helvetica, sans-serif; text-align:center;}"
ecss = ecss & " .outer {color:white; font-size:14pt; background-color:#FFFFFF; text-align:center; min-width:320px; max-width:500px; height:500px; border:1px solid;}"
ecss = ecss & " p {color:black; font-size:12pt; text-align:left; font-style:normal; position:relative;}"
ecss = ecss & " .pblue {color:blue; font-size:12pt; text-align:left;}"
ecss = ecss & " .pblack {color:#000000; font-size:12pt; text-align:left;}"
ecss = ecss & " .actionbutton {background-color:#006400; color:white; -moz-border-radius:15px; -webkit-border-radius:15px; border:5px solid; padding:5px;}"
ecss = ecss & " .psuedobuttoncellgreen {width:175px; text-align:center; -webkit-border-radius: 3px; -moz-border-radius: 3px; border-radius: 3px; background-color:#006400;}"
ecss = ecss & " .psuedobuttoncellred {width:175px; text-align:center; -webkit-border-radius: 3px; -moz-border-radius: 3px; border-radius: 3px; background-color:#DC143C;}"
ecss = ecss & " .psuedobuttongreen {width:100%; font-size:16pt; font-family:Helvetica, Arial, sans-serif; color:#ffffff; text-decoration:none; color:#ffffff; text-decoration:none; -webkit-border-radius:3px; -moz-border-radius:3px; border-radius:3px; padding:12px 0px; border: 1px solid #7FFF00; display: inline-block;}"
ecss = ecss & " .psuedobuttonred {width:100%; font-size:16pt; font-family:Helvetica, Arial, sans-serif; color:#ffffff; text-decoration:none; color:#ffffff; text-decoration:none; -webkit-border-radius:3px; -moz-border-radius:3px; border-radius:3px; padding:12px 0px; border: 1px solid #FFA500; display: inline-block;}"
ecss = ecss & " </style>"


' --- Data for Invite Email ---

sInviteBannerText = "Please Join My Team"
sInviteName = Request("sInviteName")
sInviteMemberID = Request("sInviteMemberID")
sInviteEmail = Request("sInviteEmail")
sTeam_ID = Request("sTeam_ID")

' --- Create Email message ---
ebody = ecss & "<html>"
ebody = ebody & "<head>"
ebody = ebody & "<title>Invite to Join My Team</title>"
ebody = ebody & "<meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">"
ebody = ebody & "</head>"
ebody = ebody & "<body bgcolor=""#FFFFFF"" text=""#000000"">"

ebody = ebody & "<div class=outer>"

ebody = ebody & "<div style="&SQT&"text-align:center;"&SQT&">"
ebody = ebody & "<img style='width:300px;' name=BannerLogo src='http://usawaterski.com/rankings/images/General/JoinTheFun.JPG' alt=Accept_Join>"
ebody = ebody & "</div>"
ebody = ebody & "<div style="&SQT&"margin-top:20px; text-align:center; font-size:32px; color:red;"&SQT&"><i>Please Join My Team</i></div>"

ebody = ebody & "<div class=pblack style="&SQT&"margin-top:20px;"&SQT&">Hi "&sInviteName&":</div>"
ebody = ebody & "<div class=pblack style=""margin-top:20px;"">"
ebody = ebody & "  I am building a waterski team using the new mobile app from the <b>American Water Ski Association</b>. This system uses your real scores together with scores of other team members to establish a Team Ranking.  The ranking is based on each member's improvement throughout the year."
ebody = ebody & "</div>"

ebody = ebody & "<div class=pblack style=""margin-top:20px;"">"
ebody = ebody & "  To join my team you have to accept my invitation so my team can become active."
ebody = ebody & "  Once everyone I have invited has accepted my invitation, the team will appear under the Team Rankings in a League called <b>"&sThisTeamTypeDescription&"</b>. We will be competing against other teams in the same league throughout the year."  
ebody = ebody & "</div>"
ebody = ebody & "<div style=margin-top:15px; text-align:center;><span class=pblack style=font-size:14pt;>Team Name:&nbsp;</span><br><span class=pblue style=font-size:16pt;>"&sThisTeamName&"</span></div>"
ebody = ebody & "<div style='margin-top:5px; text-align:center;' ><span class=pblack>Team ID: </span><span class=pblue>"&sTeam_ID&"</span></div>"

ebody = ebody & "<div class=pblack style='margin-top:15px; text-align:center;'>To <b>Accept</b> and be part of my team, click below</div>"
ebody = ebody & "<div class=pblack style=""margin-top:10px; text-align:center;"">"
ebody = ebody & "      <table align=""center"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
ebody = ebody & "        <tr>"
ebody = ebody & "          <td class=psuedobuttoncellgreen><a href='http://usawaterski.org/rankings/vteams_manage.asp?action=acceptinvite&team_id="&sTeam_ID&"&sMemberID="&sInviteMemberID&"' target='_blank' class=psuedobuttongreen>Join My Team</a></td>"
ebody = ebody & "        </tr>"
ebody = ebody & "      </table>"
ebody = ebody & "</div>"

ebody = ebody & "<div class=pblack style='margin-top:15px; text-align:center;'>To <b>Decline</b> participation with this team, click below.</div>"
ebody = ebody & "<div class=pblack style='margin-top:10px; text-align:center;'>"
ebody = ebody & "      <table align=""center"" border=""0"" cellspacing=""0"" cellpadding=""0"">"
ebody = ebody & "        <tr>"
ebody = ebody & "          <td class=psuedobuttoncellred><a href='http://usawaterski.org/rankings/vteams_manage.asp?action=declineinvite&team_id="&sTeam_ID&"&sMemberID="&sInviteMemberID&"' target='_blank' class=psuedobuttonred>Decline Invitation</a></td>"
ebody = ebody & "        </tr>"
ebody = ebody & "      </table>"
ebody = ebody & "</div>"

ebody = ebody & "<div class=pblack style='margin-top:20px; text-align:center;'>"
ebody = ebody & "I am looking forward to having you on my team."
ebody = ebody & "<br><b>"&sThisTeamManagerName&"</b>"
ebody = ebody & "</div>"

ebody = ebody & "<div class=pblack style='text-align:center; margin-top:15px; padding-bottom;30px'>Click the image below from your phone to try out the new AWSA mobile App."
ebody = ebody & "<br>" 
ebody = ebody & " <a href='http://usawaterski.org/rankings/mainmenu_m.asp' style=""text-decoration:none;"">"
ebody = ebody & "  <img style=""width:57px;"" name=""MobileAppIcon"" src=""http://www.usawaterski.com/rankings/images/icons/AWSA_HomeScreen_57.PNG"" alt=""Mobile App"">"
ebody = ebody & " </a>"
ebody = ebody & "</div>"

ebody = ebody & "</div>"
ebody = ebody & "<br><br><br>"
ebody = ebody & "</body></html>"

eMailBody = ebody

'response.write("</div><br>eMailTo = "&eMailTo&"<br>eMailCC = "&eMailCC&"<br>eMailBCC = "&eMailBCC&"<br>eMailFrom = "&eMailFrom&"")
'response.write("<br>eMailSubj = "&eMailSubj&"<br><br>"&eMailBody)
'response.end




END SUB



%>


