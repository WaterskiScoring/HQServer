<%

' --- This is the display program for the Registration Module
' --- Originally intended to be only displays to limit computation and logic in this module.
' --- Written by Mark Crone


' ---------------------------
   SUB DisplayAccordion
' ---------------------------





%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<%
' <html xmlns="http://www.w3.org/1999/xhtml" >

' <head>
' --------------------------------------------------------------------------------------------------------
' --- JAVASCRIPT as patch for GR only being offered as traditional format even when tournament is PandC
' --------------------------------------------------------------------------------------------------------

%>
<script type="text/javascript">
	
	function ResetRoundsToOne(EvtNo)
		{
			alert('Number of pulls not applicable in Grassroots events.  Fees are based on number of events entered.  Setting value to 1')
			if (document.getElementById('GRRadioButton').value=="G")
				{
    			var PulldownName = "fFeeRounds" + EvtNo;
     			document.getElementById(PulldownName).value=1;
     		}	
		}
</script>






  <link href="http://www.usawaterski.org/css/styles.css" rel="stylesheet" type="text/css" />

   <style type="text/css">
    		/* Required Fields */
        .reqd_blu  {
            font-weight: bold;
            color: blue;
        		}
    		/* Accordion */
        
        .accordionHeader {
            border: 1px solid #2F4F4F;
            color: white;
            background-color: #2E4d7B;
	        	font-family: Arial, Sans-Serif;
	        	font-size: 12px;
	        	font-weight: bold;
            padding: 5px;
            margin-top: 5px;
            width:100%;
            cursor: pointer;
		        }
        .accordionHeaderSelected {
            border: 1px solid #2F4F4F;
            color: white;
            background-color: #5078B3;
	        	font-family: Arial, Sans-Serif;
	        	font-size: 12px;
	        	font-weight: bold;
            padding: 5px;
            margin-top: 5px;
            cursor: pointer;
            width:100%;
    		    }
        .accordionHeader a:hover {
	        	background: none;
	        	text-decoration: underline;
            width:100%;
        		}
        .accordionHeader a {
	        	color: #FFFFFF;
	        	background: none;
	        	text-decoration: none;
          	width:100%;
        		}
        .accordionHeaderSelected a {
	        	color: #FFFFFF;
	        	background: none;
	        	text-decoration: none;
          	width:100%;
        		}
        .accordionContent {
            background-color: #D3DEEF;
            border: 1px dashed #2F4F4F;
            border-top: none;
            padding: 5px;
            padding-top: 10px;
            width:100%;            
        		}


/* -- New Classes for Tournament tab - 10/25/2015 */
	.container { width:100%; display:inline-block; position:relative;	padding: 0 0 0 0; font-size:0;}
	.feerowdiv { display:inline-block; margin: 0 0 0 0; padding: 0 0 0 0; }
	.olrsecheading { border:0px solid black; width:98%; height:20px; padding-right:2px; display:inline-block; text-align:left; font-size:12px; font-weight:bold;}
	.feeheading {  border:0px solid black; width:80px; height:15px; padding-right:2px; display:inline-block; text-align:right; font-size:10px; font-weight:bold;}
	.tourrowlabel { border:0px solid black; margin-top:0px; margin-bottom:0px; margin-right:0px; padding-right:2px; padding-left:10px; padding-top:0px; padding-bottom:0px; width:90px; height:12px; display:inline-block; text-align:right; font-size:10px; font-weight:normal; }
	.tourrowlabel2 { border:0px solid black; margin-right:0px; margin-left:40px; padding-right:2px; padding-left:10px; padding-top:0px; padding-bottom:0px; width:90px; height:12px; position:absolute; left:363px; display:inline-block; text-align:right; font-size:10px; font-weight:normal; }
	.tourgendataC1 { border:0px solid black; margin-top:0px; margin-bottom:0px; margin-right:0px; padding-left:10px; padding-right:2px; padding-top:0px; padding-bottom:0px; width:277px; height:12px; display:inline-block; text-align:left; color:blue; font-size:10px; font-weight:normal; }
	.tourgendataC2 { border:0px solid black; margin 0 0 0 0; padding-left:10px; padding-right:2px; padding-top:0px; padding-bottom:0px; width:155px; height:12px; position:absolute; left:506px; display:inline-block; text-align:left; color:blue; font-size:10px; font-weight:normal; }
	.feeamt { border:0px solid black; margin-top:0px; margin-bottom:0px; padding-right:2px; color:blue; width:80px; padding-top:0px; padding-bottom:0px; height:12px; display:inline-block; text-align:right; font-size:10px; font-weight:normal; }
	.eventline { border:0px solid black; margin-top:0px; margin-bottom:0px; margin-right:0px; padding-right:2px; padding-left:0px; padding-top:0px; padding-bottom:0px; height:22px; display:inline-block; text-align:right; font-size:10px; font-weight:normal; }

	form { display:inline-block;}


/* -- Squeezes div tags together */
/* margin-right:-4px;	*/
/* vertical-align:top; 	*/
	
/* this style applies to the SpaceTable table */
.tour_div {width:100%; display:inline-block; position:relative; margin-left:0px; padding-left:10px; border:1px solid <%=HQSiteColor2%>; background-color:<%=TableColor1%>; }

.AdminButtonStyle { width:9em; margin-left:0px; background-color:red; color:white; display:inline-block; }
.UserButtonStyle { width:9em; margin-left:0px; background-color:white; display:inline-block; }
.SpecialUserButtonStyle { width:9em; margin-left:0px; background-color:yellow; color:black; display:inline-block; }
.YellowButtonStyle { width:9em; margin-left:0px; background-color:yellow; color:black; display:inline-block; }
.spanbuttons { border:0px solid red; width:130px; height:25px; padding-left:0px; padding-right:0px; text-align:center; display:inline-block; }
.buttonrow { border:0px solid black; width:98%; height:25px; padding-top:10px; padding-bottom:10px; margin-top:10px; }

.gentext { border:0px solid black; margin-top:5px; width:98%; display:inline-block; text-align:left; font-size:10px; font-weight:normal; }

</style>
<%






IF DisplayVars="on" THEN DisplayPertinentVariables


' --- Controls whether or not Admin buttons are visible
FormAdminStatus="none"
IF adminmenulevel>=20 OR LCASE(Session("UserAdminPW"))=LCASE(Session("AdminCode")) THEN FormAdminStatus="inline-block"

' --- TEMP ---
'FormAdminStatus="inline-block"


IF TRIM(sTDescription)<>"" THEN ThisDescription = TRIM(sTDescription) + "<br>" 
IF TRIM(sFDescription)<>"" THEN ThisDescription = TRIM(sFDescription) + "<br>" 
IF TRIM(sWDescription)<>"" THEN ThisDescription = TRIM(sWDescription) + "<br>" 
IF TRIM(sKDescription)<>"" THEN ThisDescription = TRIM(sKDescription) + "<br>" 
IF TRIM(sCDescription)<>"" THEN ThisDescription = TRIM(sCDescription) + "<br>" 
 





OFDescArray = Array (sOF1Desc,sOF2Desc,sOF3Desc,sOF4Desc,sOF5Desc,sOF6Desc,sOF7Desc,sOF8Desc,sOF9Desc,sOF10Desc)
OFAmtArray = Array (sOF1Amt,sOF2Amt,sOF3Amt,sOF4Amt,sOF5Amt,sOF6Amt,sOF7Amt,sOF8Amt,sOF9Amt,sOF10Amt)
OFRequiredArray = Array (sOF1Required,sOF2Required,sOF3Required,sOF4Required,sOF5Required,sOF6Required,sOF7Required,sOF8Required,sOF9Required,sOF10Required)			
OFMaxQtyArray = Array (sOF1MaxQty,sOF2MaxQty,sOF3MaxQty,sOF4MaxQty,sOF5MaxQty,sOF6MaxQty,sOF7MaxQty,sOF8MaxQty,sOF9MaxQty,sOF10MaxQty)			
OFQtyArray = Array (sOF1Qty,sOF2Qty,sOF3Qty,sOF4Qty,sOF5Qty,sOF6Qty,sOF7Qty,sOF8Qty,sOF9Qty,sOF10Qty)
OFFeeArray = Array (sOF1Fee,sOF2Fee,sOF3Fee,sOF4Fee,sOF5Fee,sOF6Fee,sOF7Fee,sOF8Fee,sOF9Fee,sOF10Fee)





' --- Sets column width for fees ---
sClassWidth=83



sClassFeeXStatus="none"
sClassFeeCashStatus="none"

sClassXHeadingText="&nbsp"
sClassCashHeadingText="&nbsp"

' --- Sets the value shown in the left margin for the meaning of the fee row ---
' --- *** UPDATE *** Add other sports disciplines as ELSEIF ---
' IF Gr2AWS_SPulls>0 THEN sClassRow1HeadingText="Grassroots"

' --- Row 1 ---
sClassRow1HeadingText="Grassroots"

' --- Row 2 ---
IF KSClassT>0 OR KTClassT>0 THEN 
	sClassRow2HeadingText="Class T"
ELSE 		' -- formerly IF SClassC>0 OR TClassC>0 OR JClassC>0 OR BSClassC>0 OR BTClassC>0 OR BJClassC>0 THEN 
		sClassRow2HeadingText="Class C"
END IF		

' --- Row 3 ---
IF KSClassQ>0 OR KTClassQ>0 THEN 
		sClassRow3HeadingText="Class Q"	
ELSE		' -- Formerly IF SClassE>0 OR TClassE>0 OR JClassE>0 THEN 
		sClassRow3HeadingText="Class E"
END IF

' --- Row 4 ---
'IF SClassL>0 OR TClassL>0 OR JClassL>0 OR BSClassL>0 OR BTClassL>0 OR BJClassL>0 THEN 
		sClassRow4HeadingText="Class L"
'END IF

' --- Row 5 ---
'IF SClassR>0 OR TClassR>0 OR JClassR>0 OR BSClassR>0 OR BTClassR>0 OR BJClassR>0 THEN 
		sClassRow5HeadingText="Class R"
'END IF

' --- Row 6 Experimental ---
IF SClassX>0 OR TClassX>0 OR JClassX>0 THEN 
		sClassFeeXStatus="inline-block"			
		sClassXHeadingText="Experimental"
END IF

' --- Row 7 Cash ---
IF SClassCash>0 OR TClassCash>0 OR JClassCash>0 THEN 
		sClassFeeCashStatus="inline-block"			
		sClassCashHeadingText="Cash"
END IF









' --------------------------------------
' --- Fee Column Heading Description ---
' --------------------------------------

Fee1DescriptionText="1 Event"
Fee2DescriptionText="&nbsp;"
Fee3DescriptionText="&nbsp;"

Dim SL_YorN, TR_YorN, JU_YorN
SL_YorN=0
TR_YorN=0
JU_YorN=0

IF SCGr2AWS_SPulls>=1 OR lassC>=1 OR SClassE>=1 OR SClassL>=1 OR SClassR>=1 OR SClassX>=1 OR SClassCash>=1 THEN SL_YorN=1 
IF TClassC>=1 OR TClassE>=1 OR TClassL>=1 OR TClassR>=1 OR TClassX>=1 OR TClassCash>=1 THEN TR_YorN=1 
IF JClassC>=1 OR JClassE>=1 OR JClassL>=1 OR JClassR>=1 OR JClassX>=1 OR JClassCash>=1 THEN JU_YorN=1 
		

IF sTPandC=true THEN Fee1DescriptionText="1st Pull"

'IF Gr2AWS_SPulls>=2 OR SClassC>=2 OR TClassC>=2 OR JClassC>=2 OR SClassE>=2 OR TClassE>=2 OR JClassE>=2 OR SClassL>=2 OR TClassL>=2 OR JClassL>=2 OR SClassR>=2 OR TClassR>=2 OR JClassR>=2 OR SClassX>=2 OR TClassX>=2 OR JClassX>=2 OR SClassCash>=2 OR TClassCash>=2 OR JClassCash>=2 THEN
IF SL_YorN + TR_YorN + JU_YorN >=2 THEN
		Fee2DescriptionText="2 Events"
		IF sTPandC=true THEN Fee2DescriptionText="2nd Pull"
END IF
'IF Gr2AWS_SPulls>=3 OR SClassC>=3 OR TClassC>=3 OR JClassC>=3 OR SClassE>=3 OR TClassE>=3 OR JClassE>=3 OR SClassL>=3 OR TClassL>=3 OR JClassL>=3 OR SClassR>=3 OR TClassR>=3 OR JClassR>=3 OR SClassX>=3 OR TClassX>=3 OR JClassX>=3 OR SClassCash>=3 OR TClassCash>=3 OR JClassCash>=3 THEN
IF SL_YorN + TR_YorN + JU_YorN >=3 THEN
		Fee3DescriptionText="3 Events"
		IF sTPandC=true THEN Fee3DescriptionText="Addl Pulls"
END IF



' ---------------------------------------------------------------------------------------
' --- Define whether to display the Fees for each class and number of pulls or events ---
' ---------------------------------------------------------------------------------------

IF Gr2AWS_SPulls>=1 THEN sClassFeeG1Text = FormatCurrency(sClassFeeG1,2) ELSE sClassFeeG1Text = "--" END IF
IF Gr2AWS_SPulls>=2 THEN sClassFeeG2Text = FormatCurrency(sClassFeeG2,2) ELSE sClassFeeG2Text = "--" END IF
IF Gr2AWS_SPulls>=3 THEN sClassFeeG3Text = FormatCurrency(sClassFeeG3,2) ELSE sClassFeeG3Text = "--" END IF

'IF SClassC>=1 OR TClassC>=1 OR JClassC>=1 THEN sClassFeeC1Text = FormatCurrency(sClassFeeC1,2) ELSE sClassFeeC1Text = "--" END IF
'IF SClassC>=2 OR TClassC>=2 OR JClassC>=2 THEN sClassFeeC2Text = FormatCurrency(sClassFeeC2,2) ELSE sClassFeeC2Text = "--" END IF
'IF SClassC>=3 OR TClassC>=3 OR JClassC>=3 THEN sClassFeeC3Text = FormatCurrency(sClassFeeC3,2) ELSE sClassFeeC3Text = "--" END IF

'IF SClassE>=1 OR TClassE>=1 OR JClassE>=1 THEN sClassFeeE1Text = FormatCurrency(sClassFeeE1,2) ELSE sClassFeeE1Text = "--" END IF
'IF SClassE>=2 OR TClassE>=2 OR JClassE>=2 THEN sClassFeeE2Text = FormatCurrency(sClassFeeE2,2) ELSE sClassFeeE2Text = "--" END IF
'IF SClassE>=3 OR TClassE>=3 OR JClassE>=3 THEN sClassFeeE3Text = FormatCurrency(sClassFeeE3,2) ELSE sClassFeeE3Text = "--" END IF

'IF SClassL>=1 OR TClassL>=1 OR JClassL>=1 THEN sClassFeeL1Text = FormatCurrency(sClassFeeL1,2) ELSE sClassFeeL1Text = "--" END IF
'IF SClassL>=2 OR TClassL>=2 OR JClassL>=2 THEN sClassFeeL2Text = FormatCurrency(sClassFeeL2,2) ELSE sClassFeeL2Text = "--" END IF
'IF SClassL>=3 OR TClassL>=3 OR JClassL>=3 THEN sClassFeeL3Text = FormatCurrency(sClassFeeL3,2) ELSE sClassFeeL3Text = "--" END IF

'IF SClassR>=1 OR TClassR>=1 OR JClassR>=1 THEN sClassFeeR1Text = FormatCurrency(sClassFeeR1,2) ELSE sClassFeeR1Text = "--" END IF
'IF SClassR>=2 OR TClassR>=2 OR JClassR>=2 THEN sClassFeeR2Text = FormatCurrency(sClassFeeR2,2) ELSE sClassFeeR2Text = "--" END IF
'IF SClassR>=3 OR TClassR>=3 OR JClassR>=3 THEN sClassFeeR3Text = FormatCurrency(sClassFeeR3,2) ELSE sClassFeeR3Text = "--" END IF

'IF SClassX>=1 OR TClassX>=1 OR JClassX>=1 THEN sClassFeeX1Text = FormatCurrency(sClassFeeX1,2) ELSE sClassFeeX1Text = "--" END IF
'IF SClassX>=2 OR TClassX>=2 OR JClassX>=2 THEN sClassFeeX2Text = FormatCurrency(sClassFeeX2,2) ELSE sClassFeeX2Text = "--" END IF
'IF SClassX>=3 OR TClassX>=3 OR JClassX>=3 THEN sClassFeeX3Text = FormatCurrency(sClassFeeX3,2) ELSE sClassFeeX3Text = "--" END IF

'IF SClassCash>=1 OR TClassCash>=1 OR JClassCash>=1 THEN sClassFeeCash1Text = FormatCurrency(sClassFeeCash1,2) ELSE sClassFeeCash1Text = "--" END IF
'IF SClassCash>=2 OR TClassCash>=2 OR JClassCash>=2 THEN sClassFeeCash2Text = FormatCurrency(sClassFeeCash2,2) ELSE sClassFeeCash2Text = "--" END IF
'IF SClassCash>=3 OR TClassCash>=3 OR JClassCash>=3 THEN sClassFeeCash3Text = FormatCurrency(sClassFeeCash3,2) ELSE sClassFeeCash3Text = "--" END IF




IF SClassC>=1 OR TClassC>=1 OR JClassC>=1 THEN sClassFeeC1Text = FormatCurrency(sClassFeeC1,2) ELSE sClassFeeC1Text = "--" END IF
IF (SClassC>=1 AND TClassC>=1) OR (SClassC>=1 AND JClassC>=1) OR (TClassC>=1 AND JClassC>=1) THEN sClassFeeC2Text = FormatCurrency(sClassFeeC2,2) ELSE sClassFeeC2Text = "--" END IF
IF SClassC>=1 AND TClassC>=1 AND JClassC>=1 THEN sClassFeeC3Text = FormatCurrency(sClassFeeC3,2) ELSE sClassFeeC3Text = "--" END IF

IF SClassE>=1 OR TClassE>=1 OR JClassE>=1 THEN sClassFeeE1Text = FormatCurrency(sClassFeeE1,2) ELSE sClassFeeE1Text = "--" END IF
IF (SClassE>=1 AND TClassE>=1) OR (SClassE>=1 AND JClassE>=1) OR (TClassE>=1 AND JClassE>=1) THEN sClassFeeE2Text = FormatCurrency(sClassFeeE2,2) ELSE sClassFeeE2Text = "--" END IF
IF SClassE>=1 AND TClassE>=1 AND JClassE>=1 THEN sClassFeeE3Text = FormatCurrency(sClassFeeE3,2) ELSE sClassFeeE3Text = "--" END IF

IF SClassL>=1 OR TClassL>=1 OR JClassL>=1 THEN sClassFeeL1Text = FormatCurrency(sClassFeeL1,2) ELSE sClassFeeL1Text = "--" END IF
IF (SClassL>=1 AND TClassL>=1) OR (SClassL>=1 AND JClassL>=1) OR (TClassL>=1 AND JClassL>=1) THEN sClassFeeL2Text = FormatCurrency(sClassFeeL2,2) ELSE sClassFeeL2Text = "--" END IF
IF SClassL>=1 AND TClassL>=1 AND JClassL>=1 THEN sClassFeeL3Text = FormatCurrency(sClassFeeL3,2) ELSE sClassFeeL3Text = "--" END IF

IF SClassR>=1 OR TClassR>=1 OR JClassR>=1 THEN sClassFeeR1Text = FormatCurrency(sClassFeeR1,2) ELSE sClassFeeR1Text = "--" END IF
IF (SClassR>=1 AND TClassR>=1) OR (SClassR>=1 AND JClassR>=1) OR (TClassR>=1 AND JClassR>=1) THEN sClassFeeR2Text = FormatCurrency(sClassFeeR2,2) ELSE sClassFeeR2Text = "--" END IF
IF SClassR>=1 AND TClassR>=1 AND JClassR>=1 THEN sClassFeeR3Text = FormatCurrency(sClassFeeR3,2) ELSE sClassFeeR3Text = "--" END IF

IF SClassX>=1 OR TClassX>=1 OR JClassX>=1 THEN sClassFeeX1Text = FormatCurrency(sClassFeeX1,2) ELSE sClassFeeX1Text = "--" END IF
IF (SClassX>=1 AND TClassX>=1) OR (SClassX>=1 AND JClassX>=1) OR (TClassX>=1 AND JClassX>=1) THEN sClassFeeX2Text = FormatCurrency(sClassFeeX2,2) ELSE sClassFeeX2Text = "--" END IF
IF SClassX>=1 AND TClassX>=1 AND JClassX>=1 THEN sClassFeeX3Text = FormatCurrency(sClassFeeX3,2) ELSE sClassFeeX3Text = "--" END IF

IF SClassCash>=1 OR TClassCash>=1 OR JClassCash>=1 THEN sClassFeeCash1Text = FormatCurrency(sClassFeeCash1,2) ELSE sClassFeeCash1Text = "--" END IF
IF (SClassCash>=1 AND TClassCash>=1) OR (SClassCash>=1 AND JClassCash>=1) OR (TClassCash>=1 AND JClassCash>=1) THEN sClassFeeCash2Text = FormatCurrency(sClassFeeCash2,2) ELSE sClassFeeCash2Text = "--" END IF
IF SClassCash>=1 AND TClassCash>=1 AND JClassCash>=1 THEN sClassFeeCash3Text = FormatCurrency(sClassFeeCash3,2) ELSE sClassFeeCash3Text = "--" END IF

'IF SClassE>=1 OR TClassE>=1 OR JClassE>=1 THEN sClassFeeE1Text = FormatCurrency(sClassFeeE1,2) ELSE sClassFeeE1Text = "--" END IF
'IF SClassE>=2 OR TClassE>=2 OR JClassE>=2 THEN sClassFeeE2Text = FormatCurrency(sClassFeeE2,2) ELSE sClassFeeE2Text = "--" END IF
'IF SClassE>=3 OR TClassE>=3 OR JClassE>=3 THEN sClassFeeE3Text = FormatCurrency(sClassFeeE3,2) ELSE sClassFeeE3Text = "--" END IF

'IF SClassL>=1 OR TClassL>=1 OR JClassL>=1 THEN sClassFeeL1Text = FormatCurrency(sClassFeeL1,2) ELSE sClassFeeL1Text = "--" END IF
'IF SClassL>=2 OR TClassL>=2 OR JClassL>=2 THEN sClassFeeL2Text = FormatCurrency(sClassFeeL2,2) ELSE sClassFeeL2Text = "--" END IF
'IF SClassL>=3 OR TClassL>=3 OR JClassL>=3 THEN sClassFeeL3Text = FormatCurrency(sClassFeeL3,2) ELSE sClassFeeL3Text = "--" END IF

'IF SClassR>=1 OR TClassR>=1 OR JClassR>=1 THEN sClassFeeR1Text = FormatCurrency(sClassFeeR1,2) ELSE sClassFeeR1Text = "--" END IF
'IF SClassR>=2 OR TClassR>=2 OR JClassR>=2 THEN sClassFeeR2Text = FormatCurrency(sClassFeeR2,2) ELSE sClassFeeR2Text = "--" END IF
'IF SClassR>=3 OR TClassR>=3 OR JClassR>=3 THEN sClassFeeR3Text = FormatCurrency(sClassFeeR3,2) ELSE sClassFeeR3Text = "--" END IF

'IF SClassX>=1 OR TClassX>=1 OR JClassX>=1 THEN sClassFeeX1Text = FormatCurrency(sClassFeeX1,2) ELSE sClassFeeX1Text = "--" END IF
'IF SClassX>=2 OR TClassX>=2 OR JClassX>=2 THEN sClassFeeX2Text = FormatCurrency(sClassFeeX2,2) ELSE sClassFeeX2Text = "--" END IF
'IF SClassX>=3 OR TClassX>=3 OR JClassX>=3 THEN sClassFeeX3Text = FormatCurrency(sClassFeeX3,2) ELSE sClassFeeX3Text = "--" END IF

'IF SClassCash>=1 OR TClassCash>=1 OR JClassCash>=1 THEN sClassFeeCash1Text = FormatCurrency(sClassFeeCash1,2) ELSE sClassFeeCash1Text = "--" END IF
'IF SClassCash>=2 OR TClassCash>=2 OR JClassCash>=2 THEN sClassFeeCash2Text = FormatCurrency(sClassFeeCash2,2) ELSE sClassFeeCash2Text = "--" END IF
'IF SClassCash>=3 OR TClassCash>=3 OR JClassCash>=3 THEN sClassFeeCash3Text = FormatCurrency(sClassFeeCash3,2) ELSE sClassFeeCash3Text = "--" END IF


byu=2
IF sMemberID="000001151" AND byu=1 THEN
		response.write("<br>SClassL = "&SClassL)
		response.write("<br>SClassR = "&SClassR)
		response.write("<br>JClassL = "&JClassL)
		response.write("<br>JClassR = "&JClassR)
		response.write("<br>TClassL = "&TClassL)
		response.write("<br>TClassR = "&TClassR)
response.write("<br>")
		response.write("<br>sClassFeeL1 = "&sClassFeeL1)				
		response.write("<br>sClassFeeL2 = "&sClassFeeL2)
		response.write("<br>sClassFeeL3 = "&sClassFeeL3)
		response.write("<br>sClassFeeR1 = "&sClassFeeR1)				
		response.write("<br>sClassFeeR2 = "&sClassFeeR2)
		response.write("<br>sClassFeeR3 = "&sClassFeeR3)
	
	
END IF
	
	

' --------------------------------
' --- Set late fee information ---
' --------------------------------

sLateFeeText = "--"
IF sTLFPerDay=true AND sTLateFee>cdbl(0.00) THEN 
		sLateFeeText = FormatCurrency(sTLateFee,2) & " Per Day"
ELSEIF sTLFPerDay<>true AND sTLateFee>cdbl(0.00) THEN  
		sLateFeeText = FormatCurrency(sTLateFee,2) & ""
END IF



' -------------------------
' --- Family Entry Fees ---
' -------------------------

sTEntryFeeFamilyAmtText="--"
sMaxFamMembersAmtText="--"
sMaxFamMembersExtraAmtText="--"

IF sTEntryFeeFamily>cdbl(0.00) THEN sTEntryFeeFamilyAmtText = FormatCurrency(sTEntryFeeFamily,2) & " Base: " &sMaxFamMembers& " Members"
IF sTEntryFeeFamExtra>cdbl(0.00) THEN	sMaxFamMembersExtraAmtText = FormatCurrency(sTEntryFeeFamExtra,2) & " Ea Addl Member"




' -------------------------------------------------------------------------
' --- If any Optional Fees then display the Option Fee HEADING and DATA ---
' -------------------------------------------------------------------------

sOptionalDivDisplay = "none"	
IF TRIM(sOF1Desc)<>"" OR TRIM(sOF2Desc)<>"" OR TRIM(sOF3Desc)<>"" OR TRIM(sOF4Desc)<>"" OR TRIM(sOF5Desc)<>"" OR TRIM(sOF6Desc)<>"" OR TRIM(sOF7Desc)<>"" OR TRIM(sOF8Desc)<>"" OR TRIM(sOF9Desc)<>"" OR TRIM(sOF10Desc)<>"" THEN
		' --- At least one item so show heading --
		sOptionalDivDisplay = "inline-block"	
END IF




' ----------------------------------------------------------------------
' ----------------------------------------------------------------------
' -----------------------   TOURNAMENT INFORMATION   -------------------
' ----------------------------------------------------------------------
' ----------------------------------------------------------------------


%>

<div class="container"> <% ' -- Surrounds everything %>

  <div class="<% IF nav=1 THEN response.write("accordionHeaderSelected") ELSE response.write("accordionHeader") END IF %>" style="width:100%;">
		<span style="text-align:left; width:30%;">
			<a href="/rankings/<%=RegFileName%>?nav=1" title="TourID: <%=sTourID%>">STEP 1 - Tournament</a>
		</span>
		<span style="width:80%; text-align:left; padding-left:30px; color:yellow" ><% =sTourName %></span>
  </div>



  <div id="RegPanel1" class="tour_div" style="padding-top:10px; display:<% IF nav=1 THEN response.write("block") ELSE response.write("none") END IF %>;">
			<div class="olrsecheading" style="width:665px"><b>TOURNAMENT DETAILS</b></div>
				<br>
			<span class="tourrowlabel">Tour Name</span>
	  	<span class="tourgendataC1"><% =sTourName %></span>
	  	<span class="tourrowlabel2">Registrar</span>
	  	<span class="tourgendataC2"><%=sTRegistrarName%></span>
				<br>
	  	<span class="tourrowlabel">Tour ID</span>
	  	<span class="tourgendataC1"><% =sTourID %></span>
	  	<span class="tourrowlabel2">Address</span>
	  	<span class="tourgendataC2"><%=sTRegistrarAddr%></span>
				<br>  
	  	<span class="tourrowlabel">City/ST</span>
 			<span class="tourgendataC1"><% =sTourCity%>, <%=sTourState %></span>
	  	<span class="tourrowlabel2">&nbsp;</span>		
	  	<span class="tourgendataC2"><%= sTRegistrarCity %>, <%= sTRegistrarState %>&nbsp;<%= sTRegistrarZip %></span>
				<br>  
 			<span class="tourrowlabel">Dates</span>
	  	<span class="tourgendataC1"><% =sTDateS%>-<%=sTDateE %></span>
	  	<span class="tourrowlabel2">Phone</span>
			<span class="tourgendataC2"><%=sTRegistrarPhone%></span>
				<br>  
			<span class="tourrowlabel">Sptspan</span>
			<span class="tourgendataC1"><%=sTSptsGrpID%></span>
			<span class="tourrowlabel2">Email</span>
			<span class="tourgendataC2"><%=sTRegistrarEmail%></span>
				<br>
				<br>
			<span class="tourrowlabel" style="vertical-align:top;">Description</span>
			<span class="tourgendataC1" style="word-wrap:break-word; width:555px"><% =ThisDescription %></span>
     		<br>
      <span class="tourrowlabel" style="vertical-align:top;">Divisions</span>
      <span class="tourgendataC1" style="width:555px"><%=sTDvOffered%>&nbsp;</span>
				<br>
      <span class="tourrowlabel" style="vertical-align:top;">Directions</span>
      <span class="tourgendataC1" style="width:555px"><%=GTSDirections %>&nbsp;</span>
				<br>
      <span class="tourrowlabel" style="vertical-align:top;">Schedule</span>
      <span class="tourgendataC1" style="width:555px;"><%= GTSofE %>&nbsp;</span>
				<br>
			<span class="tourrowlabel" style="vertical-align:top;">Comments</span>
      <span class="tourgendataC1" style="width:555px;"><%= GTComments %>&nbsp;</span>
				<br>
				<br>
			<div style="width:100%;">
				<span class="olrsecheading" style="width:100px; margin-top:10px;"><b>ENTRY FEES</b></span>
				<span id="sClassFeeH1" class="feeheading">&nbsp;<%= Fee1DescriptionText %></span>
				<span id="sClassFeeH2" class="feeheading">&nbsp;<%= Fee2DescriptionText %></span>
				<span id="sClassFeeH3" class="feeheading">&nbsp;<%= Fee3DescriptionText %></span>
				<span id="sClassFeeT1" class="tourrowlabel2" style="font-weight:bold; position:absolute; left:363px; margin-top:12px; height:15px; display:inline-block;">Other</span>
			</div>
				<br>
		  <div class="feerowdiv">
				<span id="sClassFeeG0" class="tourrowlabel"><%= sClassRow1HeadingText %></span>
	  	 	<span id="sClassFeeG1" class="feeamt"><%= sClassFeeG1Text %></span>
	   		<span id="sClassFeeG1" class="feeamt"><%= sClassFeeG2Text %></span>
	   		<span id="sClassFeeG1" class="feeamt"><%= sClassFeeG3Text %></span>
	   		<span id="sClassFeeSA" class="tourrowlabel2">Late Fee</span>
	   		<span id="sClassFeeSB" class="tourgendataC2" ><%= sLateFeeText %></span>
			</div>
			<br>
				<br>
			<div class="feerowdiv">
				<span id="sClassFeeC0" class="tourrowlabel"><%= sClassRow2HeadingText %></span>
				<span id="sClassFeeC1" class="feeamt"><%= sClassFeeC1Text %></span>
				<span id="sClassFeeC2" class="feeamt"><%= sClassFeeC2Text %></span>
				<span id="sClassFeeC3" class="feeamt"><%= sClassFeeC3Text %></span>
	   		<span id="sClassFeeS1" class="tourrowlabel2" >Entry Deadline</span>
	   		<span id="sClassFeeS2" class="tourgendataC2" ><%= sTLateDate %></span>
			</div>
				<br>
			<div class="feerowdiv">	
				<span id="sClassFeeE0" class="tourrowlabel"><%= sClassRow3HeadingText %></span>
				<span id="sClassFeeE1" class="feeamt"><%= sClassFeeE1Text %></span>
				<span id="sClassFeeE2" class="feeamt"><%= sClassFeeE2Text %></span>
				<span id="sClassFeeE3" class="feeamt"><%= sClassFeeE3Text %></span>
			</div>
				<br>
	  	<div class="feerowdiv">	
				<span id="sClassFeeL0" class="tourrowlabel"><% =sClassRow4HeadingText %></span>
				<span id="sClassFeeL1" class="feeamt"><% =sClassFeeL1Text %></span>
				<span id="sClassFeeL2" class="feeamt"><% =sClassFeeL2Text %></span>
				<span id="sClassFeeL3" class="feeamt"><% =sClassFeeL3Text %></span>
	   		<span id="sClassFeeS5" class="tourrowlabel2" >Family Fee</span>
	   		<span id="sClassFeeS6" class="tourgendataC2"><%= sTEntryFeeFamilyAmtText %></span>
			</div>
				<br>
	  	<div class="feerowdiv">	
				<span id="sClassFeeR0" class="tourrowlabel"><% =sClassRow5HeadingText %></span>
				<span id="sClassFeeR1" class="feeamt"><% =sClassFeeR1Text %></span>
				<span id="sClassFeeR2" class="feeamt"><% =sClassFeeR2Text %></span>
				<span id="sClassFeeR3" class="feeamt"><% =sClassFeeR3Text %></span>
	   		<span id="sClassFeeS7" class="tourrowlabel2" >Extra/Member</span>
	   		<span id="sClassFeeS8" class="tourgendataC2" style="text-align:left;"><%= sMaxFamMembersExtraAmtText %></span>
			</div>
				<br>
	  	<div style="display:<%= sClassFeeXStatus %>;">	
				<span id="sClassFeeX0" class="tourrowlabel"><%=sClassXHeadingText%></span>
				<span id="sClassFeeX1" class="feeamt"><%=sClassFeeX1Text %></span>
				<span id="sClassFeeX2" class="feeamt"><%=sClassFeeX2Text %></span>
				<span id="sClassFeeX3" class="feeamt"><%=sClassFeeX3Text %></span>
			</div>
	  	<br>
	  	<div style="display:<% =sClassFeeCashStatus %>">	
				<span id="sClassFeeCash0" class="tourrowlabel"><% =sClassCashHeadingText %></span>
				<span id="sClassFeeCash1" class="feeamt"><% =sClassFeeCash1Text %></span>
				<span id="sClassFeeCash2" class="feeamt"><% =sClassFeeCash2Text %></span>
				<span id="sClassFeeCash3" class="feeamt"><% =sClassFeeCash3Text %></span>
			</div>
		<% 

		' -------------------------------------------------------------------------
		' --- If any Optional Fees then display the Option Fee HEADING and DATA --- 
		' -------------------------------------------------------------------------

		%>
	  	<div>	
			<br>			
			<br>
			<span class="olrsecheading" style="margin-top:10px; font-weight:bold; display:<%= sOptionalDivDisplay %>;">OTHER OPTIONAL OR REQUIRED ITEMS/FEES</span>
				<br>
			<span class="feeheading" style="width:400px; text-align:left; font-weight:bold; display:<%= sOptionalDivDisplay %>">Description</span>
			<span class="feeheading" style="text-align:center; display:<%= sOptionalDivDisplay %>;"><b>Required</b></span>
			<span class="feeheading" style="display:<%= sOptionalDivDisplay %>;"><b>Amount</b></span>
			<br>
			<%		
			
			' --- Loop thru the Optional Fees and display the items ---
			FOR OFItem=0 TO 9
					IF TRIM(OFDescArray(OFItem))<>"" THEN 
							%> 	  
								<span class="feeamt" style="width:400px; text-align:left;"><%=OFDescArray(OFItem)%>&nbsp;</span>
								<span class="feeamt" style="text-align:center;"><% IF OFRequiredArray(OFItem)=true THEN response.write("Yes") ELSE response.write("No") END IF %> </span>
								<span class="feeamt">&nbsp;<%=FormatCurrency(OFAmtArray(OFItem),2)%></span>
							<%
					END IF
			NEXT
			%>
		</div>
		<%



		' -------------------------------------------------
		' --- Begin FORM portion of the TOURNAMENT page ---
		' -------------------------------------------------
	
		%>
		<div class="buttonrow">
			<form name="TournamentForm3" method="post" id="TournamentForm3">
		  		<input type="hidden" name="sTourID" value="<%=sTourID%>">
		  		<input type="hidden" name="sMemberID" value="<%=sMemberID%>">
		  		<input type="hidden" name="nav" value="6">
		  	
		  	<span class="spanbuttons">
					<input type="submit" class="UserButtonStyle" name="MainStatus" value="Continue" formaction="/rankings/<%=RegFileName%>?nav=1" title="Continue" <% =MainButtonStatus %>>
				</span>
		  	<span class="spanbuttons" style="display:<% =FormAdminStatus %>">
					<input type="submit" class="AdminButtonStyle" name="SkipToPayment" value="Payment Page" formaction="/rankings/<%=RegFileName%>?nav=6" title="Press to Skip Forward to Payment Page" <% response.write(MainButtonStatus) %>>
				</span>
				<span class="spanbuttons" style="display:<% =FormAdminStatus %>" >
					<input type="submit" class="AdminButtonStyle" name="Registration Report" value="Registrar Report" formtarget="_blank" formaction="/rankings/view-Registration16.asp?sTourID=<%=sTourID%>" title="Open Registrars reports function for this tournament.">
				</span>
				<span class="spanbuttons" style="display:<% =FormAdminStatus %>" >
					<input type="submit" class="UserButtonStyle" name="NewTour" value="New Tournament" formaction="/rankings/<%=RegFileName%>?sRunByWhat=Tour" title="Admin Users can select a new tournament offering online registration.  Note: You will be required to enter the AdminCode for the new tournament." <%=MainButtonStatus%>>
				</span>
				<span class="spanbuttons" style="display:<% =FormAdminStatus %>" >				
		  		<input type="submit" class="AdminButtonStyle" name="FAQ" value="Registrar FAQ" formaction="http://usawaterski.org/rankings/news/FAQ_Register1.htm" formtarget="_blank" title="Admin Users can view FAQ.">
				</span>
			</form>
		</div><%' -- tournament FORM section ---
		
		%>				
  </div><%' -- tour_div -- 






	' **********************************************************************
	' **********************************************************************
	' **********************************************************************
	' -----------------------   MEMBER INFORMATION   -----------------------
	' **********************************************************************
	' **********************************************************************
	' **********************************************************************	


MailLinkText = "mailto:"&sMembEmail&"?subject=Registration issue for "&sTourName&" - Member "&sFullName&""


BioStatusDisplay = "none"
IF sBioDone="N" THEN BioStatusDisplay = "inline-block"


' --- Team (NOT USED?)
IF Session("SptsGrpID")="NCW" THEN 
		TeamStatus="disabled"
		' --- LoadTeam in Tools_Definitions.asp ---
		TeamSelected="ROL"
		'LoadTeam TeamSelected, TeamStatus  
ELSE
		TeamSelected="N/A"
END IF 

IF (LEFT(Session("sMembCanSkiText"),2)<>"OK" OR LEFT(Session("sExpirationStatusText"),2)<>"OK") THEN
		IF AdminmenuLevel<50 AND TRIM(sMemberID)<>"000001151" THEN  
					MainButtonStatus="disabled"
		ELSE
					MainButtonStatus="enabled"
		END IF			
END IF

' --- TEST REMOVE ---
'MainButtonStatus="enabled"



' --------------------------
' --- Membership Status ---
' --------------------------


Dim MembershipUpgradeButtonText, UpgradeRemedyText, MembershipDisplayStatus
MembershipDisplayStatus="none"




MembershipUpgradeButtonText=""
IF LEFT(Session("sMembCanSkiText"),2)<>"OK" THEN
		MainButtonStatus="disabled"
		MembershipDisplayStatus="inline-block"
		MembershipUpgradeButtonText="Upgrade"  
		UpgradeRemedyText = "Your present <b>Membership Type</b> does not permit participation in competition events of USA Water Ski.  To upgrade your membership to a 'Competition Status', press the <b>Member Upgrade</b> button below.  <br><br>Once you have completed your membership upgrade, return to this form and press <b>Verify Upgrade</b> to activate your membership in this registration form."
ELSEIF LEFT(Session("sExpirationStatusText"),2)<>"OK" THEN 
		MainButtonStatus="disabled"
		MembershipDisplayStatus="inline-block"
		MembershipUpgradeButtonText="Renewal"  
		UpgradeRemedyText = "Your Membership has <b><u>expired</u></b> and must be renewed before you can participate in this tournament.  To renew your membership, press the <b>Member Renewal</b> button below.  <br><br>Once you complete your membership renewal, return to this page by restarting the registration process.  If you are entering the same day you upgraded your membership, press <b>Verify Renewal</b> to continue with online registration."
END IF



' ----------------
' --- Bio Form ---
' ----------------

'BioDisplayStatus="none"
BioUpdateDisplayStatus="inline-block"
BioUpdateButtonClass="UserButtonStyle"

BioVerifyDisplayStatus="none"
BioVerifyButtonClass="YellowButtonStyle"


' IF sMemberID="000001151" THEN sBioDone="N"
IF sBioDone="N" THEN 

		BioVerifyDisplayStatus="inline-block"
		BioUpdateButtonClass="YellowButtonStyle"
		
		IF AdminmenuLevel<50 THEN  
				MainButtonStatus="disabled"
		ELSE
				MainButtonStatus="enabled"
		END IF		
		
END IF

BioFormASP="bio-form.asp"

' --- TESTING ---
tyu=2
IF tyu=1 THEN
		IF sMemberID="000001151" THEN
				'BioFormASP="bio-form_TEST.asp"
				Session("sBioDoneText")="Out of Date"
				'Session("sBioDoneText")=""
		ELSE		
				'BioFormASP="bio-form.asp"
		END IF
END IF



AdminNewMemberButtonStatus="none"
IF adminmenulevel>=20 OR TestValidAdminCode=true THEN AdminNewMemberButtonStatus="inline-block"


				
%>
  <div style="width:100%" class="<% IF nav=2 THEN response.write("accordionHeaderSelected") ELSE response.write("accordionHeader") END IF %>">
		<span style="text-align:left; width:30%;">
			<a href="/rankings/<%=RegFileName%>?nav=2" title="MemberID: <%=sMemberID%>">STEP 2 - Member</a>
		</span>
		<span style="text-align:left; padding-left:30px; color:yellow"><% =sFullName %></span>
		<span style="text-align:left; padding-left:30px; color:red; display:<%= BioStatusDisplay %>;">Bio Requires Update</span>
  </div>
  <div id="RegPanel2" class="tour_div" style="padding-top:10px; display:<% IF nav=2 THEN response.write("block") ELSE response.write("none") END IF %>;">
			<div class="olrsecheading" style="width:98%;"><b>PERSONAL INFORMATION</b></div>
				<br>
			<span class="tourrowlabel">Name</span>
	  	<span class="tourgendataC1"><% =sFullName %></span>
	  	<span class="tourrowlabel">Comp Status</span>
	  	<span class="tourgendataC2"><% =Session("sMembCanSkiText") %></span>
	
				<br>
			<span class="tourrowlabel">Member ID</span>
	  	<span class="tourgendataC1"><%= Session("sMemberID") %></span>
	  	<span class="tourrowlabel">Expiration</span>
	  	<span class="tourgendataC2"><% =Session("sExpirationStatusText") %></span>
				<br>
			<span class="tourrowlabel">City/ST</span>
	  	<span class="tourgendataC1"><%= sMembCity %>,&nbsp;<%= sMembState %>&nbsp;</span>
	  	<span class="tourrowlabel">Personal Bio</span>
	  	<span class="tourgendataC2"><% =Session("sBioDoneText") %></span>
				<br>
			<span class="tourrowlabel">Age/Gender</span>
	  	<span class="tourgendataC1"><% =sMembAge %>/<% =sMembSex %>&nbsp;</span>
	  	<span class="tourrowlabel">Team</span>
	  	<span class="tourgendataC2">&nbsp;</span>
				<br>
			<span class="tourrowlabel">Phone</span>
	  	<span class="tourgendataC1"><%= sMembPhone %>&nbsp;</span>
				<br>
			<span class="tourrowlabel">Email</span>
	  	<span class="tourgendataC1"><a href="<% =MailLinkText %>"><% =sMembEmail %></a>&nbsp;</span>

			<div style="height:125px; margin-top:15px; margin-right:10px; padding-left:5px; border:0px solid red; display:<%= MembershipDisplayStatus %>;">
				<span class="olrsecheading" style="color:red; border:0px solid red; padding-top:5px;"><b>IMPORTANT NOTICE:</b></span>
				<br>
				<span class="olrsecheading" style="font-size:10px; font-style:normal; color:red; margin-bottom:10px; height:60px; text-align:left;"><% =UpgradeRemedyText %></span>
				<br>
	    	<form name="UpgradeForm" method="post" id="UpgradeForm">
		  		<span class="spanbuttons">
	  				<input type="submit" class="YellowButtonStyle" name="Upgrade" value="Member <% =MembershipUpgradeButtonText %>" formaction="https://www.usawaterski.org/renew/" formtarget="_blank" title="Upgrade or Renew your membership status">
		  		</span>
		  		<span class="spanbuttons">
						<input type="submit" class="YellowButtonStyle" name="Verify Renewal" value="Verify <% =MembershipUpgradeButtonText %>" formaction="/rankings/<%=RegFileName%>?sRunByWhat=VerifyUpgrade&nav=2" formtarget="_blank" title="Use this link to cause the entry form to verify that your renewal or upgrade is complete.">
					</span>
				</form>
			</div>
	
			<div class="buttonrow">
  			<form name="MemberForm" method="post" id="MemberForm">
			  	<span class="spanbuttons">
						<input type="submit" class="UserButtonStyle" name="MainStatus" value="Continue" formaction="/rankings/<%=RegFileName%>?nav=2" title="Continue" <%=MainButtonStatus%>>
			  	</span>
			  	<span class="spanbuttons">
	  				<input type="submit" class="UserButtonStyle" name="Previous" value="Previous" formaction="/rankings/<%=RegFileName%>?nav=2" title="Previous Page" <%=PreviousButtonStatus%>>
			  	</span>
			  	<span class="spanbuttons">
						<input type="submit" class="<% =BioUpdateButtonClass %>" name="Update_Bio" style="display:<%= BioUpdateDisplayStatus %>" formaction="/rankings/<%= BioFormASP %>?FormStatus=new" formtarget="_blank" value="Update Bio" title="Create or Update your Personal Bio. &#13; Bio is used for all tournaments. &#13; Keep your bio up-to-date so announcers have current information.">
			  	</span>
			  	<span class="spanbuttons">
						<input type="submit" class="<% =BioVerifyButtonClass %>" name="Verify Bio Update" style="display:<%= BioVerifyDisplayStatus %>" formaction="/rankings/<%=RegFileName%>?nav=2" value="Verify Bio Update" title="Press this button after updating personal bio to confirm the update is complete. ">
			  	</span>
			  	<span class="spanbuttons">
			  	  <input type="submit" class="AdminButtonStyle" name="New_Member" style="display:<% =AdminNewMemberButtonStatus %>;" formaction="/rankings/<%=RegFileName%>?rid=<%=rid%>&sRunByWhat=NewMember" value="New Member" title="Admin users may select a new member">
			  	</span>
				</form>
			</div>
	</div>
<% ' -- tour_div









' **********************************************************************
' **********************************************************************
' **********************************************************************
' ---------------------  BEGIN ENTRY FORM DISPLAY ----------------------
' **********************************************************************
' **********************************************************************
' **********************************************************************



Dim fSelectEvent, fDiv, fFeeClass, fFeeRounds, fQfyOverride, fBoat
Dim TrickEvtNo, JumpEvtNo
TrickEvtNo=0
JumpEvtNo=0


' --- Determines column width which varies depending on 
sClassCols=cdbl(0)
sClassWidth=70

sTPandCTitle = ""
sTPandCText = ""

IF sTPandC=true THEN
		sTPandCTitle = "Select the number of rounds you wish to ski in each event"
		sTPandCText = "PICK & CHOOSE"
END IF


' --- Defines the text and href of header link ---
Step3HeadingLink = ""
Step3HeadingTitle = ""
IF LEFT(Session("sMembCanSkiText"),2)="OK" AND LEFT(Session("sExpirationStatusText"),2)="OK" AND sBioDone="Y" THEN 	
		Step3HeadingLink = "/rankings/"&RegFileName&"?MainStatus=Continue&nav=2"
		Step3HeadingTitle = "Return to Step 3"
END IF  



ClassHeaderDisplayValue="inline-block"			
sClassHeaderText="ENTRY CLASSIFICATION"		


FormErrorDisplayStatus="none"
IF TRIM(sFormError)<>"" THEN 
		FormErrorDisplayStatus="inline-block"
END IF

AdminOverrideSectionDisplayStatus="none"
IF TestValidAdminCode=true OR adminmenulevel>=20 THEN AdminOverrideDisplayStatus="inline-block"

QualificationsButtonDisplaystatus = "none"
IF sQualLevel>0 THEN QualificationsButtonDisplaystatus = "inline-block"




' ------------------------------------------------------------
' --- Controls COLUMN WIDTH and [Start] POSITION settings ----
' ------------------------------------------------------------
EventWidth = "80"
EntryWidth = "45"
DivisionWidth = "170"
RoundsWidth = "55"
ClassWidth = "110"
SkillWidth = "115"

EventPosition = "10"
EntryPosition = "93"
DivisionPosition = "141"
RoundsPosition = "314"
ClassPosition = "372"
SkillPosition = "485"


' --- Allocates the room for the div tag surrounding the Admin Override section since it must have a div to display:none when not logged in ---
AdminSectionDivHeight = 60 + (TotNumEvents * 35)



	' --------------------------------
	' --- Forms TAB for ENTRY FORM ---
	' --------------------------------

  %>
  <div style="width:100%" class="<% IF nav=3 THEN response.write("accordionHeaderSelected") ELSE response.write("accordionHeader") END IF %>">
		<span style="text-align:left; width:30%;">
			<a href="<% =Step3HeadingLink %>" title="<% =Step3HeadingTitle %>">STEP 3 - Entry Form</a>
		</span>
	</div>	

  <div id="RegPanel3" class="tour_div" style="padding-top:10px; display:<% IF nav=3 THEN response.write("block") ELSE response.write("none") END IF %>;">
     <form name="EntryForm" style="width:98%;" action="/rankings/<%=RegFileName%>?nav=3" method="post" id="EntryForm">
		<%


			' ********************************************************************************
			' --- Sets HIDDEN INPUT variables for when disabled and those not in this form ---
			' ********************************************************************************
			
			' --- Sets hidden variable for form tools that are not visible ---
			SetHiddenFinancialVariables
			SetHiddenFinancialOverrideVariables 
   		SetHiddenWaiverVariables
			IF MainStatusValue="Continue" AND Edit<>"Edit" THEN SetHiddenEntryVariables	


			' --- Controls display of Admin section and associated hidden variables ---
			AdminOverrideSectionDisplayStatus = "none"
			IF TestValidAdminCode=true OR adminmenulevel >= 20 THEN 		' -- Admin User set hidden when elements disabled ---
					AdminOverrideSectionDisplayStatus = "inline-block"
		   		IF MainStatusValue="Continue" THEN SetHiddenEntryOverrideVariables

			ELSEIF TestValidAdminCode=false AND adminmenulevel < 20 THEN 	'-- Regular user - set hidden even when fields are enabled
   				SetHiddenEntryOverrideVariables
			END IF


			' -------------------------------------
			' --- Top of Content for ENTRY FORM ---
			' -------------------------------------
			%>
			<div class="olrsecheading" style="width:98%;"><a title="CalcCode:  <%= sRegFeeCalcCode %>"><b>ENTRY INFORMATION</b></a></div>
			<%

			' --- TESTING - Displays values of form elements ---
			yu=1
			IF yu=2 THEN
					%>
					<div class="olrsecheading" style="width:98%;"><b>MainStatusValue = <% =MainStatusValue %></b></div>
					<div class="olrsecheading" style="width:98%;"><b>sSelectEvent(1) = <% =sSelectEvent(1) %></b></div>
					<div class="olrsecheading" style="width:98%;"><b>sRegionalOverride = <% =sRegionalOverride %></b></div>
					<div class="olrsecheading" style="width:98%;"><b>MainButtonStatus = <% =MainButtonStatus %></b></div>
					<%
			END IF

			%>
			<br>
			<div class="feerowdiv" style="height:20px; margin-top:5px;">
					<span class="eventline" style="position:absolute; text-align:left; font-weight:bold; left:<% =EventPosition %>px; width:<% =EventWidth %>px; height:20px;">Event</span>
	  			<span class="eventline" style="position:absolute; text-align:center; font-weight:bold; left:<% =EntryPosition %>px;width:<% =EntryWidth %>px; height:20px;">Enter</span>
	  			<span class="eventline" style="position:absolute; text-align:left; font-weight:bold; left:<% =DivisionPosition %>px;width:<% =DivisionWidth %>px; height:20px;">Division</span>
					<%
					IF sTPandC=true THEN 
							%>					
							<span class="eventline" style="position:absolute; text-align:left; font-weight:bold; left:<% =RoundsPosition %>px; width:<% =RoundsWidth %>px; height:20px;" >Pulls</span>
							<%
					END IF
	  					%>
	  					<span class="eventline" style="position:absolute; text-align:left; font-weight:bold; left:<% =ClassPosition %>px;width:<% =ClassWidth %>px; height:20px; display:<% =ClassHeaderDisplayValue %>">Class</span>
	  					<%
					IF JClassC>0 OR JClassE>0 OR JClassL>0 OR JClassR>0 OR JClassCash>0 OR JClassX>0 OR TClassC>0 OR TClassE>0 OR TClassL>0 OR TClassR>0 OR TClassCash>0 OR TClassX>0 OR (sShowSkills=true AND WWakeW>0) THEN
							%>
							<span class="eventline" style="position:absolute; text-align:left; font-weight:bold; left:<% =SkillPosition %>px;width:<% =SkillWidth %>px; height:20px;">Boat/Ramp/Skill</span>
							<%
					END IF
					%>
	  	</div> 


      <%


  ' -----------------------------------------------------
	' ------------  Displays checkbox OPTION   ------------
  ' ----------------------------------------------------- 

	' --- Set default outside loop --- 
	JumpRampEnabledStatus = "disabled"
	TrickBoatEnabledStatus = "disabled"


	FOR EvtNo = 1 TO TotEv

			SELECT CASE TRIM(sTEvent(EvtNo))
				CASE "S"
						ThisMax=sMaxSLPulls
				CASE "T"
						ThisMax=sMaxTRPulls
				CASE "J"
						ThisMax=sMaxJUPulls
			END SELECT


		  IF TRIM(sTEvent(EvtNo))<>"" THEN  

					' --- Values used in reading variables for calculations ---
		  		fSelName = "fSelectEvent"&EvtNo
		  		fQfyOverride="fQfyOverride"&EvtNo
		  		fFeeRounds="fFeeRounds"&EvtNo			
		  		fSkill="fSkill"&EvtNo		
					fFeeClass = "fFeeClass"&EvtNo
		  		
		  		
		  		
		  		
		  		
		  		' --- TESTING ---
		  		IF sMemberID = "000001151"  THEN
		 					' response.end
		 					' response.write("</div></div><div><br>RFD995 = "&TRIM(sFeeClass(EvtNo)))
		 					' response.end
		 			END IF
					
					
					
					
					' --- Using RIGHT function because system sets Barefoot Tricks eventcode to BT
					IF RIGHT(TRIM(sTEvent(EvtNo)),1)="T" THEN TrickEvtNo=EvtNo
					IF TRIM(sTEvent(EvtNo))="J" THEN JumpEvtNo=EvtNo
 
					fBoat="fBoat"&TrickEvtNo

								

					' -------------------------------------------
					' --- Top of ENTRY FORM variables display ---
					' -------------------------------------------					

					' --- LoadDivDropWithAgeGender - in tools_include.asp ---
					' --- LoadRoundSkiedPulldown - in tools_definitions.asp ---
					' --- LoadClassEnteredDropDownByEvent - in tools_include.asp ---


					EventSelectionCheckedStatus = ""
					IF TRIM(sSelectEvent(EvtNo)) <> "" THEN EventSelectionCheckedStatus = "checked"
					IF TRIM(sTEvent(EvtNo))="J" THEN JumpRampEnabledStatus = AllObjectStatus
					IF TRIM(sTEvent(EvtNo))="T" THEN TrickBoatEnabledStatus = AllObjectStatus
					
		
		  		%>
					<div style="height:30px; border:0px solid red; padding-top:5px;">
						<span class="eventline" style="position:absolute; left:<% =EventPosition %>px; width:<% =EventWidth %>px; text-align:left;"><%= sTEventName(EvtNo) %></span>
						<span class="eventline" style="position:absolute; left:<% =EntryPosition %>px; width:<% =EntryWidth %>px; text-align:center;">
							<input type=checkbox id="<% =fSelName %>" name="<% =fSelName %>" <% IF sSelectEvent(EvtNo) <> "" THEN Response.Write("Checked "&AllObjectStatus) ELSE Response.write(AllObjectStatus) %>>
						</span>
						<span class="eventline" style="position:absolute; left:<% =DivisionPosition %>px; width:<% =DivisionWidth %>px; text-align:left;"><% LoadDivDropWithAgeGender sDiv(EvtNo), sTEvent(EvtNo), "fDiv"&EvtNo, AllObjectStatus %></span>
						<%
						

						' --- Displays PICK & CHOOSE rounds dropdown ---
						IF sTPandC=true THEN 				
								%><span class="eventline" style="position:absolute; left:<% =RoundsPosition %>px; width:<% =RoundsWidth %>px; text-align:left;"><% LoadRoundSkiedPulldown fFeeRounds, sFeeRounds(EvtNo), 0, ThisMax, 1, AllObjectStatus, "false" %></span><%	
						END IF  	        

						%>
						<span class="eventline" style="position:absolute; left:<% =ClassPosition %>px; width:<% =ClassWidth %>px; text-align:left; display: <% =ClassDivDisplayStatus %>;"><% LoadClassEnteredDropDownByEvent fFeeClass, sTEvent(EvtNo), sFeeClass(EvtNo), AllObjectStatus %></span>
						<%
			
						' --- Displays SKILL dropdown in Tools_Definitions.asp ---
						IF sShowSkills=true AND sTEvent(EvtNo)="WB" THEN 			
								%><span class="eventline" style="position:absolute; left:<% =SkillPosition %>px; width:<% =SkillWidth %>px; text-align:left;"><% LoadGRSkillPulldown fSkill, sSkill(EvtNo), AllObjectStatus %></span><%
						ELSEIF RIGHT(TRIM(sTEvent(EvtNo)),1)="T" THEN
							%><span class="eventline" style="position:absolute; left:<% =SkillPosition %>px; width:<% =SkillWidth %>px; text-align:left;"><% LoadBoatPulldown fBoat, sBoat(EvtNo), AllObjectStatus %></span><%
						ELSEIF sTEvent(EvtNo)="J" THEN
							' response.end
								%><span class="eventline" style="position:absolute; left:<% =SkillPosition %>px; width:<% =SkillWidth %>px; text-align:left;"><% LoadRampPulldownRegister_11072015 sDiv(EvtNo), "sRampHeight", sRampHeight, AllObjectStatus %></span><%		
						END IF
						%>
						</div>
						<%

		  END IF 

	NEXT 


  ' ----------------------------
  ' --- Notice of FORM ERROR ---
  ' ----------------------------

' FormErrorDisplayStatus = "inline-block"
'sFormError = "Form Error Test Message"
	
	IF TRIM(sFormError)<>"" THEN
			%>
			<div id="FormErrorMessage" style="width:100%; height:40px; margin-top:10px;">
				<span class="olrsecheading" style="color:red;"><%=sFormError%></span>
			</div>
			<%
	END IF




' -------------------------------
' --- Admin override Headings --- 
' -------------------------------


' --- At least one event is offered ---
IF (TestValidAdminCode=true OR adminmenulevel >= 20) AND (TRIM(sTEvent(1)) <> "" OR TRIM(sTEvent(2)) <> "" OR TRIM(sTEvent(3)) <> "" OR TRIM(sTEvent(4)) <> "" OR TRIM(sTEvent(5)) <> "" OR TRIM(sTEvent(6)) <> "") THEN  

	%>
	<div style="margin-top:20px; width:99%; height:<% =AdminSectionDivHeight %>px; border:0px solid; display:<% =AdminOverrideSectionDisplayStatus %>;">
		<div style="margin-top:5px; width:99%; height:50px;">
			<span class="olrsecheading">ADMINISTRATIVE OVERRIDE</span>
			<br>	
			<span class="eventline" style="margin-top:5px; position:absolute; left:<%=EventPosition%>px; width:<%=EventWidth%>px; text-align:left; font-weight:bold;">Event</span>
			<span class="eventline" style="margin-top:5px; position:absolute; left:<% =DivisionPosition %>px; width:<% =DivisionWidth %>px; text-align:left; font-weight:bold;">Qualification Reason</span>
			<span class="eventline" style="margin-top:5px; position:absolute; left:<% =ClassPosition %>px; width:<% =ClassWidth %>px; text-align:left; font-weight:bold;">Regl Participation</span>
			<%


			' ---------------------------------
  		' --- Regionals Excuse Override ---
			' ---------------------------------
	  	%>
			<span class="eventline" style="margin-top:5px; position:absolute; left:<% =SkillPosition %>px; width:<% =DivisionWidth %>px;">
				<select name="sRegionalOverride" value="<% =sRegionalOverride %>" style="width:11em; text-align:left;" <% =AllObjectStatus %>>
		  		<option value ="" <% IF sRegionalOverride = "" THEN Response.Write(" selected ") %>>None</Option><br>
		  		<option value ="MED" <% IF sRegionalOverride = "MED" THEN Response.Write(" selected ") %>>Medical Excuse</Option><br>
		  		<option value ="OTH" <% IF sRegionalOverride = "OTH" THEN Response.Write(" selected ") %>>Other</Option><br>
				</select>
	    </span>
		</div><!-- Enclosing headings -->
		<%


		' -------------------------------------------
		' --- EVENT level QUALIFICATIONS Override ---
		' -------------------------------------------

	  FOR EvtNo = 1 TO TotEv

	    	IF TRIM(sTEvent(EvtNo))<>"" THEN
						fQfyOverride="fQfyOverride"&EvtNo

						%>
		  			<div style="height:30px; width:320px; border:0px solid red;">
		  				<span class="eventline" style="position:absolute; left:<% =EventPosition %>px; width:<% =EventWidth %>px; text-align:left;"><% =sTEventName(EvtNo) %></span>
							<span class="eventline" style="position:absolute; left:<% =DivisionPosition %>px; width:<% =DivisionWidth %>px; text-align:left;">
							<%
										
							' --- Displays qualifications information - in Tools_Definitions.asp ---
							LoadQualificationsOverrideDropDown fQfyOverride, sQfyOverride(EvtNo), AllObjectStatus 
										
							%>
							</span>
						</div><!-- Bottom of single line of event override -->
						<%

				END IF		
		
  	NEXT 

		%>
		</div><!-- This is the bottom of ADMIN OVERRIDE - Entry Form section -->
		<%

END IF	' --- For Admin Override section ---




		' -------------------------------------------
		' --- STEP 3 ENTRY FORM - BUTTONS Section ---
		' -------------------------------------------

			%>
			<br>
			<span class="spanbuttons" style="position:absolute; margin-top:10px;">
				<input type="submit" name="MainStatus" value="<% =MainStatusValue %>" style="width:9em;" title="<% =MainStatusValue %>" <% =MainButtonStatus %>>
  		</span>
	</form>

	<span class="spanbuttons" style="margin-top:10px; padding-bottom:10px; width:98%;">
		<form action="/rankings/<%=RegFileName%>?nav=3" method="post" id="Edit">
	  		<input type="submit" name="Edit" value="Edit" style="width:9em; position:absolute; left:170px; "  title="Edit the settings on this page" <%=EditButtonStatus%>>
		</form>

		<form action="/rankings/<%=RegFileName%>?nav=3" method="post" id="Previous">
	  		<input type="submit" name="Previous" value="Previous"  style="width:9em; position:absolute; left:330px;" title="Back up to previous page" <%=PreviousButtonStatus%>>
		</form>
		<form action="/rankings/MemberQualifications.asp?sTourID=<%=sTourID%>&sMemberID=<%=sMemberID%>" method="post" target="_blank" id="Qualifications">
				<input type="submit" style="width:9em; position:absolute; left:490px; display:<%=QualificationsButtonDisplaystatus%>;" value="Qualifications"  title="Check your qualification status for this tournament.">
		</form>
	</span>	
<br><br>

	


</div><!-- RegPanel3 -->
<% 






' **********************************************************************************************************
' **********************************************************************************************************
' --------------------------------  BEGIN FINANCIAL SECTION  -----------------------------------------------
' **********************************************************************************************************
' **********************************************************************************************************




' --------------------------------------------------
' --- Establishes position and width of div tags ---
' --------------------------------------------------

FinWidth1=60
FinWidth2=128
FinWidth3=81
FinWidth4=79
FinWidth5=72
FinWidth6=70

FinPos1=0
FinPos2=63
FinPos3=234
FinPos4=318
FinPos5=580
FinPos6=655


' --- Tab Link and Title (Not used for 2016) ---
Step4HeadingTitle=""
Step4HeadingLink=""

IF nav>4 THEN 	
		Step4HeadingLink = "/rankings/"&RegFileName&"?MainStatus=Continue&nav=4"
		Step4HeadingTitle = "Return to Step 4"
END IF  



' --- Session("sWhichFamilyMemberPaid")<>"" When the family fee has been paid by another member
IF (TRIM(Session("sWhichFamilyMemberPaid"))<>"" AND sMaxFamMembers>Session("TotRegisteredFamMembers")) THEN  
		TotalLineStatusText = "FAMILY MEMB"
		TotalLineAmountText = "Paid"
		TotalLineColor = "blue"
ELSEIF cdbl(sTotalPreviousPayments) <= cdbl(sTotalFormFees) THEN 
		TotalLineStatusText = "BAL DUE"
		TotalLineAmountText = (FormatCurrency(cdbl(sTotalFormFees)-cdbl(sTotalPreviousPayments),2))
		TotalLineColor = "red"
ELSE
		TotalLineStatusText = "CREDIT DUE"
		TotalLineAmountText = FormatCurrency(cdbl(sTotalFormFees)-cdbl(sTotalPreviousPayments),2)
		TotalLineColor = "blue"
END IF  







' *** TESTING AGE ***
'sMembAge=13


' ********** TESTING **********
TestAllObject=2
IF TestAllObject=1 THEN 
		%></div><%
 		response.write("<br><br>Line 1252 - AllObjectStatus = "&AllObjectStatus)
 		response.write("<br>nav = "&nav)
 		response.write("<br>request(MainStatus) = "&request("MainStatus"))

		Response.write("<br><br>TotalLineAmountText ="&TotalLineAmountText)
		Response.write("<br>cdbl(sTotalFormFees) ="&cdbl(sTotalFormFees))
		Response.write("<br>cdbl(sTotalPreviousPayments ="&cdbl(sTotalPreviousPayments))
END IF




	' -----------------------------------------
	' --- Creates TAB for FINANCIAL SUMMARY ---
	' -----------------------------------------
  %>
  <div style="width:100%" class="<% IF nav=4 THEN response.write("accordionHeaderSelected") ELSE response.write("accordionHeader") END IF %>">
		<span style="text-align:left; width:30%;">
			<a title="Please use Navigation Buttons">STEP 4 - Financial Summary</a>
		</span>
	</div>	

  <div id="RegPanel4" class="tour_div" style="padding-top:10px; display:<% IF nav=4 THEN response.write("block") ELSE response.write("none") END IF %>;">
		
		<form name="FinancialForm" style="width:98%;" action="/rankings/<%=RegFileName%>?nav=4&thisform=financialform&sRegFeeCalcCode=<%=sRegFeeCalcCode%>" method="post" id="FinancialForm">
				<%
																		

				
				' --- Sets hidden variables when form tools are disabled ---	
				SetHiddenEntryVariables	 
				SetHiddenEntryOverrideVariables
				SetHiddenWaiverVariables

				IF MainStatusValue="Continue" AND Edit<>"Edit" AND nav=4 AND TRIM(Request("thisform"))="financialform" THEN 
						SetHiddenFinancialVariables
						SetHiddenFinancialOverrideVariables
						EntryDateObjectStatus = "disabled"
				ELSE
						IF adminmenulevel < 10 AND TestValidAdminCode=false THEN 
								SetHiddenFinancialOverrideVariables	
								EntryDateObjectStatus = "disabled"
								%><input type="hidden" name="sMembRegDate" value="<%=sMembRegDate%>"><%
						END IF		
				END IF	

						

				' --- TEST - Active in pre-2016 version --- 
				RecalcFormValues



 				
 				' ----------------------------------------------------------
 				' --- Membership Type, Fee Override and Date Entered row ---
 				' ---------------------------------------------------------- 
 				%>
				<div class="olrsecheading" style="width:98%; margin-top:10px;">
					<span class="olrsecheading" style="position:absolute; text-align:left; font-weight:bold; left:0px; padding-left:10px; width:425px; height:20px;">GENERAL INFORMATION</span>
					<span class="eventline" style="position:absolute; text-align:left; padding-top:5px; font-size:7.5px; left:605px; width:120px; height:15px;">mm/dd/yyyy hh:mm:ss PM</span>
				</div>
				<br>
				<div class="feerowdiv" style="width:98%; border:0px solid; height:20px; margin-top:5px;">
					<span class="eventline" style="position:absolute; text-align:right; color:red; left:<% =FinPos1 %>px; width:<% =FinWidth1 %>px; height:20px;">Type</span>
					<%
					' --- Selection of Entry Type ---
					IF sTEntryFeeFamily<>cdbl(0) THEN 
							%>
							<span class="eventline" style="position:absolute; text-align:left; left:<% =FinPos2 %>px; width:<% =FinWidth2 %>px; height:20px;">
								<select name="sEntryType" value="<%=sEntryType %>" style="width:8em;" <%=AllObjectStatus %> >
	  							<option value ="IND" <%IF sEntryType = "IND" THEN Response.Write(" selected ")%> >Individual</Option><br>
	  							<option value ="FAM" <%IF sEntryType = "FAM" THEN Response.Write(" selected ")%> >Family Entry</Option><br>
								</select>
	 						</span>
	 						<%
					ELSE
							%>
							<span class="eventline" style="position:absolute; text-align:left; color:blue; font-weight:bold; left:<% =FinPos2 %>px; width:<% =FinWidth2 %>px; height:20px;">Individual</span>
	 						<%
	 				END IF				  

					' --- Fee Override if Admin User ---
					IF adminmenulevel>=20 OR TestValidAdminCode=true THEN 
							%>	
							<span class="eventline" style="position:absolute; color:red; text-align:right; left:<% =FinPos3 %>px; width:<% =FinWidth3 %>px; height:20px;">Fee Override</span>
			    		<span class="eventline" style="position:absolute; text-align:left; font-weight:bold; left:<% =FinPos4 %>px; width:<% =FinWidth4 %>px; height:20px;">
								<select name="sMoneyOverride" title="Money_Override" style="width:6em;" value="<% =sMoneyOverride %>" <%=AllObjectStatus %>>
				  				<option value ="" <%IF sMoneyOverride = "" THEN Response.Write(" selected ")%> >None</Option><br>
				  				<option value ="OTH" <%IF sMoneyOverride = "OTH" THEN Response.Write(" selected ")%> >Other</Option><br>
				  				<option value ="FAM" <%IF sMoneyOverride = "FAM" THEN Response.Write(" selected ")%> >Family</Option><br>
								</select>
			    		</span>
			    		<%
					END IF

		  		%>
		  		<span class="eventline" style="position:absolute; color:red; text-align:right; left:440px; width:162px; height:20px;">Entry Date</span>
		  		<span class="eventline" style="position:absolute; text-align:left; left:605px; width:120px; height:20px;">
		  			<input type="text" name="sMembRegDate" value="<% =sMembRegDate %>" style="width:9em;" MAXLENGTH=22 size="13" <% =EntryDateObjectStatus %>>
		  		</span>

				</div><!-- Entry Type Fee Override and Date Entered Row -->
				<%
  
  
  			' ---------------------------------
				' --- Notice for Family Entries ---
				' ---------------------------------
				IF sEntryType="FAM" THEN

		   			' --- Number of people in this family membership group ---
		   			IF MainStatus<>"Verify" OR TotQualifyingFamMemb>0 THEN 
		   					%>
		   					<br>
			  				<div class="olrsecheading" style="border:1px solid; width:100%; margin-top:10px; color:red;">IMPORTANT INFORMATION ABOUT FAMILY ENTRIES</div>
								<div class="feerowdiv" style="border:1px solid black; width:100%; height:50px; margin-top:5px;">
			    				<%
			      			IF TRIM(Session("sWhichFamilyMemberPaid"))<>"" AND sMaxFamMembers>1 THEN 
			      					%><span class="eventline" style="text-align:left; width:100%; height:20px;"><%=Session("sWhichFamilyMemberPaid")%> was charged for the 'Family Entry Fee'.<br>Late entry fees and other charges are not included in Family Entry Fee.</span><%
			   					ELSEIF TRIM(Session("sWhichFamilyMemberPaid"))<>"" AND sMaxFamMembers=1 THEN 
			   							%><span class="eventline" style="text-align:left; width:100%; height:20px;"><%=Session("sWhichFamilyMemberPaid")%> was charged for the 'Family Entry Fee'. All other entries for family members will be charged the 'Additional Family Member' fee.&nbsp;Late entry fees and other charges are not included in Family Entry Fee.</span><%
			      			ELSEIF TRIM(Session("sWhichFamilyMemberPaid"))="" AND sMaxFamMembers>0 THEN 
			      					%><span class="eventline" style="text-align:left; width:100%; height:20px;">The first family member registering will be charged the 'Family Entry Fee', which pays for up to <%=sMaxFamMembers%> entries.&nbsp;  Other family members will be charged the 'Additional Family Member' fee.&nbsp;Late fees and other charges not included in Family Entry Fee.</span><%
			      			ELSE 
			      					%><span class="eventline" style="text-align:left; width:100%; height:20px;"><font color="<% =textcolor1 %>" size=<% =fontsize2 %>>&nbsp;The first family member registering will be charged the 'Family Entry Fee.'  All other entries for family members are free.<br>&nbsp;Late entry fees and other charges are not included in Family Entry Fee.</span<%
			      			END IF  
			      			%>
									<br>
									<span class="eventline" style="text-align:left; width:100%; height:20px; margin-top:10px;"><b>FAMILY MEMBERS ALLOWED UNDER THIS MEMBERSHIP - Total: <%=TotQualifyingFamMemb%></b></span>
			  				</div>
			  				<%

			  			' --- Displays the list of family members for this member ---
			  			MembNo=0
 			  			DO WHILE MembNo<10 
									MembNo=MembNo+1 
									%>
									<div class="feerowdiv" style="height:20px; margin-top:5px;">
										<span class="eventline" style="text-align:left; width:400px; height:20px;"><%=MembListName(MembNo)%></span>
									</div>
									<%
									IF TRIM(MembList(MembNo))="" THEN EXIT DO
			  			LOOP
		   				 
		   		ELSE		' --- Not a family membership even though user set it that way  
		   				%>
			  			<br>
			  			<div class="feerowdiv" style="border:1px solid; width:98%; height:20px; margin-top:10px;">
			    			<span class="eventline" style="text-align:left; width:400px; height:20px; color:red;"><b>WARNING - THIS IS NOT A FAMILY MEMBERSHIP TYPE</b></span>
			  			</div><%
		   		END IF

		END IF 




		' ----------------------------------
		' --- Second Section SUB-HEADING ---
		' ----------------------------------		

		%>
		<div class="olrsecheading" style="width:98%; margin-top:15px;">
			<span class="olrsecheading" style="position:absolute; text-align:left; font-weight:bold; left:0px; padding-left:10px; width:427px; height:20px;">FEES & CHARGES</span>
			<span class="eventline" style="position:absolute; text-align:right; font-weight:bold; left:<% =FinPos5 %>px; width:<% =FinWidth5 %>px; height:20px;"><a title="CalcCode:  <%= Request("sRegFeeCalcCode") %>">ENTRY</a></span>	 		 
			<span class="eventline" style="position:absolute; text-align:right; color:blue; font-weight:bold; left:<% =FinPos6 %>px; width:<% =FinWidth6 %>px; height:20px;" ><%=FormatCurrency(sEntryFee,2)%></span>	 		 
		</div>
		<%


	  ' ------------------------------------------------------------	
	  ' ---- Discount to Junior B/G 1-3 per Tour_Manager.asp   -----
	  ' ------------------------------------------------------------

	  IF sJrDiscPerc <> 0 AND sMembAge < 18 AND sEntryFee > 0 THEN 
	  		%>
				<br>
				<div class="feerowdiv" style="height:20px; margin-top:5px;">
					<span class="eventline" style="position:absolute; text-align:right; left:<% =FinPos5 %>px; width:<% =FinWidth5 %>px; height:20px;">Junior Disc</span>
					<span class="eventline" style="position:absolute; text-align:right; color:blue; left:<% =FinPos6 %>px; width:<% =FinWidth6 %>px; height:20px;"><%= FormatCurrency(sJrDiscAmt,2) %></span>
				</div>
				<%
		END IF 	


		' -------------------------------------------------------------------------	
		' ---- Discount to divisions M/W-6 if specified in Tour_Manager.asp   -----
		' -------------------------------------------------------------------------

		IF cdbl(sSrDiscPerc) <> 0 AND sMembAge > 59 AND cdbl(sEntryFee) > 0 THEN  
				%>
				<br>
				<div class="feerowdiv" style="height:20px; margin-top:5px;">
					<span class="eventline" style="position:absolute; text-align:right; left:<% =FinPos5 %>px; width:<% =FinWidth5 %>px; height:20px;">Senior Disc</span>
					<span class="eventline" style="position:absolute; text-align:right; color:blue; left:<% =FinPos6 %>px; width:<% =FinWidth6 %>px; height:20px;"><%=FormatCurrency(sSrDiscAmt,2)%></span>
				</div>
				<%
		END IF


		' -------------------------------------------------------------------------	
		' ---------- Discount to OFFICIALS if specified in Tour_Manager.asp   -----
		' -------------------------------------------------------------------------  

		IF sOffDiscPerc <> cdbl(0) THEN
				%>
				<br>
				<div class="feerowdiv" style="height:20px; margin-top:5px;">
			    <span class="eventline" style="position:absolute; height:2px; text-align:right; padding-top:0px; left:0px; width:<%=FinWidth1%>px; height:20px;">
			    	<input type=checkbox name="fOfficial" <%IF sOfficial = "on" THEN Response.Write("Checked "&AllObjectStatus) ELSE Response.write(AllObjectStatus) %>>
			    </span>
			    <span class="eventline" style="position:absolute; text-align:left; padding-top:2px; left:<%=FinPos2%>px; width:252px; height:20px;">Yes, I am an Invited Official</span>
					<%
				  IF sOfficial = "on" THEN  
				  		%>
							<span class="eventline" style="position:absolute; text-align:right; left:<% =FinPos5 %>px; width:<% =FinWidth5 %>px; height:20px;">Official Disc</span>
							<span class="eventline" style="position:absolute; text-align:right; color:blue; left:<% =FinPos6 %>px; width:<% =FinWidth6 %>px; height:20px;"><%=FormatCurrency(sOffDiscAmt,2)%></span>
							<%
			  	END IF 
			  	%>
				</div>
				<%
		END IF


		' -------------------------------------------------------------------------------------------------	
		' ---------- Discount to CLUB MEMBERS if match to ClubCode as specified in Tour_Manager.asp   -----
		' -------------------------------------------------------------------------------------------------  

		IF sClubDiscPerc <> cdbl(0) THEN 
				%>
				<br>
				<div class="feerowdiv" style="height:20px; margin-top:5px;">
			    <span class="eventline" style="position:absolute; text-align:right; height:2px; padding-top:0px; font-weight:bold; left:0px; width:<%=FinWidth1%>px; height:20px;">
			    	<input type=checkbox name="fClubMemb" <%IF sClubMemb = "on" THEN Response.Write("checked "&AllObjectStatus) ELSE Response.write(AllObjectStatus) %>>
			    </span>
			    <span class="eventline" style="position:absolute; text-align:left; padding-top:2px; left:<%=FinPos2%>px; width:178px; height:20px;">Member of Host Club - Code</span>
	  		  <span class="eventline" style="position:absolute; text-align:center; padding-top:0px; margin-bottom:2px; left:<% =FinPos3 %>px; width:<% =FinWidth3 %>px; height:20px;">
						<input type="text" name="fClubCode" value="<% =sClubCode %>" maxlength=5 size="7" <%=AllObjectStatus%>>
			    </span>
					<%  
					IF cdbl(sClubDiscPerc) <> 0 AND sClubMemb = "on" AND cdbl(sEntryFee) > 0 THEN
							IF TRIM(sClubCode) <> "" AND TRIM(sClubCode)=TRIM(sTourClubCode) THEN  
									%>
									<span class="eventline" style="position:absolute; text-align:right; left:<% =FinPos5 %>px; width:<% =FinWidth5 %>px; height:20px;">Club Disc</span>
									<span class="eventline" style="position:absolute; text-align:right; color:blue; left:<% =FinPos6 %>px; width:<% =FinWidth6 %>px; height:20px;"><%=FormatCurrency(sClubDiscAmt,2)%></span>
									<%	
							ELSEIF MainStatus="Verify" AND (TRIM(sClubCode) = "" OR TRIM(sClubCode)<>TRIM(sTourClubCode)) THEN  
				   				%>
					  			<span class="eventline" style="position:absolute; text-align:right; color:red; font-weight:bold; left:<% =FinPos5 %>px; width:<% =FinWidth5 %>px; height:20px;">Invalid Code</span>
					  			<%						
				  		END IF
					END IF
					%>
				</div>
				<%	
		END IF  
  

		' -------------------------------------------	
		' ---- Donation to AWSEF Building Fund  -----
		' -------------------------------------------  
		IF sAWSEFDon_OK=true THEN
		  	%>
		  	<br>
				<div class="feerowdiv" style="height:20px; margin-top:5px;">
		    	<span class="eventline" style="position:absolute; text-align:center; left:0px; width:<%=FinWidth1%>px; height:20px;">
		    		<% LoadValuePulldown "sAWSEFDonation", sAWSEFDonation, 10, 300, 10, AllObjectStatus, "true" %>
		    		<%
		    		chg=2
		    		IF chg=1 THEN
		    				%>		
		    				<input type=checkbox name="fAWSEFCheck" <% IF sAWSEFCheck = "on" THEN Response.Write("Checked "&AllObjectStatus) ELSE Response.write(AllObjectStatus) %>>
		    				<%
		    		END IF
		    		%>		
		    	</span>	
	  		  <span class="eventline" style="position:absolute; text-align:left; padding-top:2px; left:<%=FinPos2%>px; width:252px; height:20px;">Add Donation to USA Water Ski Foundation</span>
		    	<%
		  		' IF sAWSEFCheck = "on" THEN  
		  				%>
							<span class="eventline" style="position:absolute; text-align:right; left:<% =FinPos5 %>px; width:<% =FinWidth5 %>px; height:20px;">Donation</span>
							<span class="eventline" style="position:absolute; text-align:right; color:blue; left:<% =FinPos6 %>px; width:<% =FinWidth6 %>px; height:20px;">
								<% =FormatCurrency(sAWSEFDonation,2) %>
							</span>
							<%
		  		' END IF
					%>
				</div>
				<%	
		ELSE
				%><input type="hidden" name="sAWSEFDonation" value="<%=sAWSEFDonation%>"><%
		END IF



		' ---------------------------------------------
		' --------  LATE FEES --------------------------
		' ---------------------------------------------  

IF 1=1 AND (sMemberID="000001151" OR sMemberID="000040569") THEN
		%>
		</div></div><div style="color:red; background-color:yellow;">
		<%
		response.write("<br>sLateDays = "&sLateDays)
		response.write("<br>sTLFPerDay = "&sTLFPerDay)
		response.write("<br>sLateFeeTot = "&sLateFeeTot)
		response.write("<br>sAWSEFDonation) = "&FormatCurrency(sAWSEFDonation,2))
		
END IF

		LateFeeDaysText = ""
		IF sTLFPerDay=true THEN LateFeeDaysText = "- "&sLateDays&" Days"
		IF Cdbl(sLateFeeTot)>Cdbl(0.00) THEN  
				%><br>
				<div class="feerowdiv" style="height:20px; margin-top:5px;">
					<span class="eventline" style="position:absolute; text-align:right; left:<% =FinPos5 %>px; width:<% =FinWidth5 %>px; height:20px;">Late Fee <% =LateFeeDaysText %></span>
					<span class="eventline" style="position:absolute; text-align:right; color:blue; left:<% =FinPos6 %>px; width:<% =FinWidth6 %>px; height:20px;"><%=FormatCurrency(sLateFeeTot)%></span>
				</div>
			<%
		END IF



	  ' ----------------------------
	  ' ---- Banquet Tickets  -----
	  ' ----------------------------
		
		ExtraBanquetTicketText=""	
		BanquetTicketRowHeight=20

		IF sBTickWithE=true THEN 
				ExtraBanquetTicketText = "<br>If Attending Select (1) for FREE Ticket - Charge for 2 or more"
				BanquetTicketRowHeight=30
		END IF

		IF sBTickCost>0 THEN 
				%>
				<br>
		  	<div class="feerowdiv" style="height:<% =BanquetTicketRowHeight %>px; margin-top:5px;">
		   		<span class="eventline" style="position:absolute; text-align:center; left:0px; width:<%=FinWidth1%>px; height:<% =BanquetTicketRowHeight %>px;">
						<%
						LoadValuePulldown "sBanquetQty", sBanquetQty, 0, 10, 1, AllObjectStatus, "true"
						%>
					</span>
		    	<span class="eventline" style="position:absolute; text-align:left; left:<% =FinPos2 %>px; width:252px; height:<% =BanquetTicketRowHeight %>px;">Banquet <%=FormatCurrency(sBTickCost)%>/ticket <%= ExtraBanquetTicketText %>.</span>
					<span class="eventline" style="position:absolute; text-align:right; left:<% =FinPos5 %>px; width:<% =FinWidth5 %>px; height:<% =BanquetTicketRowHeight %>px;">Banquet</span>
					<span class="eventline" style="position:absolute; text-align:right; color:blue; left:<% =FinPos6 %>px; width:<% =FinWidth6 %>px; height:<% =BanquetTicketRowHeight %>px;"><%=FormatCurrency(sBanquetTot)%></span>
				</div>
				<%
		END IF 
  
  
  	
		' ----------------------------
		' --- Optional Items --------- 
		' ----------------------------
		
			FOR OFItem=0 TO 9
					IF TRIM(OFDescArray(OFItem))<>"" THEN
							%>
							<br>
							<div class="feerowdiv" style="height:20px; margin-top:5px;">
			    			<span class="eventline" style="position:absolute; text-align:center; left:0px; width:<%=FinWidth1%>px; height:20px;">
		    					<%
									' --- Array starts at zero so +1 required to get correct element NAME --- 
									LoadValuePulldown "sOF"&OFItem+1&"Qty", OFQtyArray(OFItem), 0, OFMaxQtyArray(OFItem), 1, AllObjectStatus, "true" 
									%>
								</span>
		    				<span class="eventline" style="position:absolute; text-align:left; left:<% =FinPos2 %>px; width:252px; height:20px;"><%=OFDescArray(OFItem)%> at <%=FormatCurrency(OFAmtArray(OFItem))%>:</span>
		    				<span class="eventline" style="position:absolute; text-align:right; left:<% =FinPos5 %>px; width:<% =FinWidth5 %>px; height:20px;">Item Cost</span>
		    				<span class="eventline" style="position:absolute; text-align:right; color:blue; left:<% =FinPos6 %>px; width:<% =FinWidth6 %>px; height:20px;"><%=FormatCurrency(OFFeeArray(OFItem))%></span>>
		  				</div>
		  				<%
					END IF
			NEXT


		' -----------------------------------------
		' --- Totals at the lower right of page ---
		' -----------------------------------------

		%>
		<br>
		<div class="feerowdiv" style="height:20px; margin-top:5px;">
			<span class="eventline" style="position:absolute; text-align:left; color:red; padding-left:10px; left:0px; width:427px; height:20px;"><% =sOtherNote %></span>				   
		  <span class="eventline" style="position:absolute; text-align:right; left:<% =FinPos5 %>px; width:<% =FinWidth5 %>px; height:20px;"><b>TOTAL ALL</b></span>
		  <span class="eventline" style="position:absolute; text-align:right; color:blue; left:<% =FinPos6 %>px; width:<% =FinWidth6 %>px; height:20px;"><b><%=FormatCurrency(sTotalFormFees,2)%></b></span>
		</div>
		<br>
		<div class="feerowdiv" style="height:20px; margin-top:5px;">
			<span class="eventline" style="position:absolute; text-align:right; left:<% =FinPos5 %>px; width:<% =FinWidth5 %>px; height:20px;">Prev Paid</span>
			<span class="eventline" style="position:absolute; text-align:right; color:blue; left:<% =FinPos6 %>px; width:<% =FinWidth6 %>px; height:20px;"><%=FormatCurrency(cdbl(sTotalPreviousPayments),2)%></span>
		</div>
		
		<br>
		<div class="feerowdiv" style="height:20px; margin-top:5px;">
			<%


			IF MainStatusValue="Verify" AND TRIM(sFormError)="" THEN 
					%> 
					<span class="eventline" style="position:absolute; text-align:left; color:red; padding-left:10px; left:0px; width:427px; height:20px;">Press <b>Verify</b> to calculate your total fees and apply any applicable discounts.</span>
		     	<%
			ELSEIF TRIM(sFormError)<>"" THEN
		  	' --- Notice of FORM ERROR ---
					%> 
					<span class="eventline" style="position:absolute; text-align:right; color:red; padding-left:10px; left:0px; width:427px; height:20px;"><b>INPUT ERROR:</b> <%=sFormError%></span>
					<%
					MainButtonStatus="disabled"
			END IF 

			
			%>
			<span class="eventline" style="position:absolute; text-align:right; left:<% =FinPos5 %>px; width:<% =FinWidth5 %>px; height:20px;"><b><%= TotalLineStatusText %></b></span>
			<span class="eventline" style="position:absolute; text-align:right; color:blue; left:<% =FinPos6 %>px; width:<% =FinWidth6 %>px; height:20px;"><b><%= TotalLineAmountText %></b></span>
		</div>

		<br>
		<span class="spanbuttons" style="position:absolute; margin-top:10px;">
			<input type="submit" name="MainStatus" value="<%=MainStatusValue%>" style="width:9em" title="<%=MainStatusValue%>" <%=MainButtonStatus%>>
  	</span>
	</form>
	<%
	
	' ---------------------------------
	' --- FINANCIAL SUMMARY BUTTONS ---
	' ---------------------------------

	%>
	<span class="spanbuttons" style="margin-top:10px; padding-bottom:10px; width:98%;">
		<form name="FinancialForm2" method="post" action="/rankings/<%=RegFileName%>?nav=4" id="FinancialForm2">
	  		<input type="submit" name="Edit" value="Edit" style="position:absolute; left:165px; width:9em" title="Edit the settings on this page" <%=EditButtonStatus%>>
		</form>
		<form action="/rankings/<%=RegFileName%>?nav=4" method="post" id="Previous">
	  		<input type="submit" name="Previous" value="Previous" style="position:absolute; left:320px; width:9em" title="Back to Entry Form tab" <%=PreviousButtonStatus%>>
		</form>
		<form action="/rankings/MemberQualifications.asp?sTourID=<%=sTourID%>&sMemberID=<%=sMemberID%>" method="post" target="_blank" id="Qualifications">
				<input type="submit" style="width:9em; position:absolute; left:490px; display:<%=QualificationsButtonDisplaystatus%>;" value="Qualifications"  title="Check your qualification status for this tournament.">
		</form>
	</span>	
		<br><br>


    </div>
  </div>
<%






	' *****************************************************************************
	' *****************************************************************************
  ' *****************************************************************************
  ' -----------------------------  RELEASE  -------------------------------------
	' *****************************************************************************
	' *****************************************************************************
  ' *****************************************************************************
	
	
	%>
  <div class="<% IF nav=5 THEN response.write("accordionHeaderSelected") ELSE response.write("accordionHeader") END IF %>">
 		<TABLE>
 			<tr>
 				<td width="150px" align=left>STEP 5 - Waiver</td>
 				<td align=left>
 					<font style="color:<%=Session("sReleaseTextColor")%>"><%=Session("sReleaseText")%></font>
 				</td>
 			</TR>
 		</TABLE>
  </div>

  <div class="innertable" style="display:<% IF nav=5 THEN response.write("block") ELSE response.write("none") END IF %>;">
    <div id="RegPanel5">
    	<%	 
			AccordRelease 
			%>
    </div>
  </div>
  <%






	' *****************************************************************************
	' *****************************************************************************
  ' *****************************************************************************
  ' -----------------------------  PAYMENT PAGE  --------------------------------
	' *****************************************************************************
	' *****************************************************************************
  ' *****************************************************************************
 

Dim Item_Name(10), Amount(10), Quantity(10), ItemNo, RegFeeAmount, NETAmountDue

FOR ItemNo=1 TO 10
	Amount(ItemNo)=cdbl(0)
NEXT

' --- Get previous charges from Transaction Table ---
ReadFromTransTable

' --- TEST - Really?? probably not 12/21/2015 --- 
RecalcFormValues



' --- Find just the registration fee amount ---
RegFeeAmount=sTotalFormFees-cdbl(sTotalPreviousPayments)-sBanquetTot-sAWSEFDonation+sBanquetTotTrans+sAWSEFDonationTrans-sOF1Fee+sOF1FeeTrans-sOF2Fee+sOF2FeeTrans-sOF3Fee+sOF3FeeTrans-sOF4Fee+sOF4FeeTrans-sOF5Fee+sOF5FeeTrans-sOF6Fee+sOF6FeeTrans-sOF7Fee+sOF7FeeTrans-sOF8Fee+sOF8FeeTrans-sOF9Fee+sOF9FeeTrans-sOF10Fee+sOF10FeeTrans
sAllDisc=sSrDiscAmt+sJrDiscAmt+sOffDiscAmt+sClubDiscAmt

'response.write("<br>HERE")
'response.write("<br>sTotalPreviousPayments="&sTotalPreviousPayments)
'response.write("<br>sOF1Fee="&sOF1Fee)
'response.write("<br>sTotalFormFees="&sTotalFormFees)
'response.write("<br>RegFeeAmount="&RegFeeAmount)
'response.end

ty=2
IF ty=2 AND (sMemberID="000001151" OR sMemberID="1151") THEN
	RegFeeAmount=cdbl(1)
	sAWSEFDonation=cdbl(0.23)
END IF


ItemNo=0

' --- Fees due are less than the amount paid (REFUND OWED) ---
IF RegFeeAmount<0 OR sBanquetTot-sBanquetTotTrans<0 OR sAWSEFDonation-sAWSEFDonationTrans<0 OR sOF1Fee-sOF1FeeTrans<0 OR sOF2Fee-sOF2FeeTrans<0 OR sOF3Fee-sOF3FeeTrans<0 OR sOF4Fee-sOF4FeeTrans<0 OR sOF5Fee-sOF5FeeTrans<0 OR sOF6Fee-sOF6FeeTrans<0 OR sOF7Fee-sOF7FeeTrans<0 OR sOF8Fee-sOF8FeeTrans<0 OR sOF9Fee-sOF9FeeTrans<0 OR sOF10Fee-sOF10FeeTrans<0 THEN
		ItemNo=ItemNo+1
		Item_Name(ItemNo)="Changes to Registration for Member # "&sMemberID&" at "&sTourName
		Quantity(ItemNo)="1"
		Amount(ItemNo)=sTotalFormFees-cdbl(sTotalPreviousPayments)

' --- Fees du are GREATER or EQUAL TO than amount paid
ELSE
		' --- If the total fees less banquet tickets and AWSEF are greater than zero ---
		IF RegFeeAmount>0 THEN
				ItemNo=ItemNo+1
				Item_name(ItemNo)="Registration for Member # "&sMemberID&" at "&sTourName
				Quantity(ItemNo)="1"
				Amount(ItemNo)=cdbl(RegFeeAmount)	
		END IF


		' --- If the Banquet $$ is greater than zero and ticket payments are less than Banquet $$ --- 
		IF sBanquetTot>0 AND sBanquetTotTrans<sBanquetTot THEN
				ItemNo=ItemNo+1
				IF sBTickWithE=true THEN 
						Item_Name(ItemNo)="Banquet Ticket(s) - NOTE: 1 extra ticket is included with entry"
				ELSE
						Item_name(ItemNo)="Banquet Ticket(s)"
				END IF	
				Quantity(ItemNo)=(sBanquetTot-sBanquetTotTrans)/sBTickCost
				Amount(ItemNo)=cdbl(sBTickCost)
		END IF


		' --- If AWSEF donation ---
		IF sAWSEFDonation>0 AND sAWSEFDonation<>sAWSEFDonationTrans THEN 
				ItemNo=ItemNo+1
				Item_Name(ItemNo)="AWSEF Donation"
				Quantity(ItemNo)="1"
				Amount(ItemNo)=cdbl(sAWSEFDonation-sAWSEFDonationTrans)
		END IF

		' --- These define OPTIONAL CUSTOM ITEMS ---
		NewOptDisplay="N"
		IF NewOptDisplay="Y" THEN
				FOR OFItem=0 TO 9
						IF TRIM(OFDescArray(OFItem))<>"" AND OFQtyArray(OFItem)>0 THEN
								ItemNo=ItemNo+1
								Item_Name(ItemNo)=OFDescArray(OFItem)
								Quantity(ItemNo)=OFQtyArray(OFItem)
								Amount(ItemNo)=cdbl(OFFeeArray(OFItem)/OFQtyArray(OFItem))
						END IF
				NEXT

		ELSE
		
				IF TRIM(sOF1Desc)<>"" AND sOF1Qty>0  THEN
						ItemNo=ItemNo+1
						Item_Name(ItemNo)=sOF1Desc
						Quantity(ItemNo)=sOF1Qty
						Amount(ItemNo)=cdbl(sOF1Fee/sOF1Qty)
				END IF

				IF TRIM(sOF2Desc)<>"" AND sOF2Qty>0 THEN
						ItemNo=ItemNo+1
						Item_Name(ItemNo)=sOF2Desc
						Quantity(ItemNo)=sOF2Qty
						Amount(ItemNo)=cdbl(sOF2Fee/sOF2Qty)
				END IF

				IF TRIM(sOF3Desc)<>"" AND sOF3Qty>0 THEN
						ItemNo=ItemNo+1
						Item_Name(ItemNo)=sOF3Desc
						Quantity(ItemNo)=sOF3Qty
						Amount(ItemNo)=cdbl(sOF3Fee/sOF3Qty)
				END IF

				IF TRIM(sOF4Desc)<>"" AND sOF4Qty>0  THEN
						ItemNo=ItemNo+1
						Item_Name(ItemNo)=sOF4Desc
						Quantity(ItemNo)=sOF4Qty
						Amount(ItemNo)=cdbl(sOF4Fee/sOF4Qty)
				END IF

				IF TRIM(sOF5Desc)<>"" AND sOF5Qty>0  THEN
						ItemNo=ItemNo+1
						Item_Name(ItemNo)=sOF5Desc
						Quantity(ItemNo)=sOF5Qty
						Amount(ItemNo)=cdbl(sOF5Fee/sOF5Qty)
				END IF

				IF TRIM(sOF6Desc)<>"" AND sOF6Qty>0  THEN
						ItemNo=ItemNo+1
						Item_Name(ItemNo)=sOF6Desc
						Quantity(ItemNo)=sOF6Qty
						Amount(ItemNo)=cdbl(sOF6Fee/sOF6Qty)
				END IF

				IF TRIM(sOF7Desc)<>"" AND sOF7Qty>0 THEN
						ItemNo=ItemNo+1
						Item_Name(ItemNo)=sOF7Desc
						Quantity(ItemNo)=sOF7Qty
						Amount(ItemNo)=cdbl(sOF7Fee/sOF7Qty)
				END IF

				IF TRIM(sOF8Desc)<>"" AND sOF8Qty>0  THEN
						ItemNo=ItemNo+1
						Item_Name(ItemNo)=sOF8Desc
						Quantity(ItemNo)=sOF8Qty
						Amount(ItemNo)=cdbl(sOF8Fee/sOF8Qty)
				END IF

				IF TRIM(sOF9Desc)<>"" AND sOF9Qty>0  THEN
						ItemNo=ItemNo+1
						Item_Name(ItemNo)=sOF9Desc
						Quantity(ItemNo)=sOF9Qty
						Amount(ItemNo)=cdbl(sOF9Fee/sOF9Qty)
				END IF

				IF TRIM(sOF10Desc)<>"" AND sOF10Qty>0  THEN
						ItemNo=ItemNo+1
						Item_Name(ItemNo)=sOF10Desc
						Quantity(ItemNo)=sOF10Qty
						Amount(ItemNo)=cdbl(sOF10Fee/sOF10Qty)
				END IF

		END IF 	' -- Discrete Optional versus loop --

END IF			' -- Fees greater or LE to amount paid --




IF sHQAccount=true THEN 
		PaymentMethodNoticeText = "When you click the 'Pay Now' button, you will be directed to make a payment through a secure server from USA Water Ski.  You may use a credit card only. Once you begin the payment process, do not stop until you reach the 'Receipt page as this registration session may expire.  If this occurs, just restart your registration for this tournament and make the required payment.  Your registration information is stored, but <b><u> your registration will not be active until payment has been made.</u><br><br>A USA Water Ski registration receipt will be emailed to you following payment.  <b>Please retain this receipt as this is the only proof of payment you will receive.</b>  For refunds, credits or other matters relating to entry fees and payments please contact USA Water Ski Competition Dept at 800-533-2972."
ELSE 
 		PaymentMethodNoticeText = "When you click the 'Pay Now' button, you will be directed to make a payment to the PayPal account for <b>"&sTourName&"</b>.  You can pay using your PayPal account, or you may use a credit card.  If you elect to set up a new PayPal account, this registration session may expire.  If this occurs, once your PayPal account has been verified, just restart your registration for this tournament.  Your registration information is stored, but <b><u> your registration will not be active until payment has been made.</u></b><br><br>In addition to the USA Water Ski registration receipt that will be emailed to you following payment, you will also receive a separate PayPal receipt.  <b>Please retain the PayPal receipt as this is the only proof of payment you will receive.</b>  Refunds, credits or other matters relating to entry fees and payments should be directed to the tournament organizer, or to the contact information on your PayPal receipt."
END IF 




ReturnURL="http://usawaterski.org/rankings/"&RegFileName&"?sTourID="&sTourID&"&sMemberID="&sMemberID&"&sOrderNo="&sOrderNo&"&nav=7&sPayType=PayPal"
ReturnURLBad="http://usawaterski.org/rankings/"&RegFileName&"?sTourID="&sTourID&"&sMemberID="&sMemberID&"&sOrderNo="&sOrderNo&"&nav=6&sPayType=PPErr"

simage_url="https://www.usawaterski.org/rankings/images/logos/usawslogo_no_sub.jpg"



Step6HeadingLink = ""
Step6HeadingTitle = ""




' --- Sets Position and Width of Order Summary (Payment) tab columns --- 
OrderPos1=10
OrderPos2=88
OrderPos3=572
OrderPos4=652

OrderWidth1=76
OrderWidth2=550
OrderWidth3=78
OrderWidth4=78		





  %>
  <div style="width:100%" class="<% IF nav=6 THEN response.write("accordionHeaderSelected") ELSE response.write("accordionHeader") END IF %>">
		<span style="text-align:left; width:30%;">STEP 6 - Payment</span>
		<span style="text-align:left; padding-left:30px; color:<%=Session("FeeStatusTextColor")%>"><%=Session("FeeStatusText")%></span>
	</div>	

  <!-- <div id="RegPanel6" class="tour_div" style="padding-top:10px; display:block;"> -->
  <div id="RegPanel6" class="tour_div" style="padding-top:10px; display:<% IF nav=6 THEN response.write("block") ELSE response.write("none") END IF %>;">
		<div class="olrsecheading" style="width:98%; margin-top:10px;">
			<span class="eventline" style="position:absolute; text-align:left; font-weight:bold; font-size:16px; left:0px; padding-left:10px; height:20px;">Final Review of Your Order</span>
		</div>

 		<br>
		<div style="margin-left:0px; padding-left:0px;">
			<span class="gentext"><%= PaymentMethodNoticeText %></span>	
		</div>
		<%


		' --- Nationals Refund policy ---
		t=1
		IF RIGHT(LEFT(sTourID,6),3)="999" OR t=2 THEN
	   		%>
	   		<br>	
	   		<div class="feerowdiv" style="margin-top:10px;">
	   			<span class="eventline" style="height:14px; margin-top:5px; font-size:12px; color:red; font-weight:bold;">REFUND POLICY</span>
					<br>
					<span class="gentext" ><b>$35 of the entry fee (per person)</b> is an administration and processing fee and is non-refundable. Late fees are also non-refundable. If you registered for the <%=sTourName%>, and are you are unable to participate you must <b>submit a cancelation request to USA Water Ski in writing</b> prior to the start of your first event. If you do not submit a cancelation request, you will not receive a refund. Cancelation requests will be honored due to lack of qualification or medical with documentation provided. All other excuses will be evaluated by the president of AWSA when the cancellation notice is received to determine if a refund will be issued. USA Water Ski will pay refunds <b>within 60 days following the conclusion of Nationals.</b> Please send cancellation requests and any documentation to competition@usawaterski.org</span>
				</div>
				<%
   	END IF 



	  %>	
		<div style="width:100%; margin-top:15px;">
		  <span class="eventline" style="position:absolute; padding-top:2px; text-align:center; font-weight:bold; left:<%= OrderPos1 %>px; width:<% =OrderWidth1 %>px; color:#FFFFFF; border:1px solid <%=HQSiteColor2%>; background-color:<%=HQSiteColor2%>;">Item #</span>
	    <span class="eventline" style="position:absolute; padding-top:2px; padding-left:2px; text-align:left; font-weight:bold; left:<%= OrderPos2 %>px; width:<% =OrderWidth2 %>px; color:#FFFFFF; border:1px solid <%=HQSiteColor2%>; background-color:<%=HQSiteColor2%>;">Description</span>
	    <span class="eventline" style="position:absolute; padding-top:2px; text-align:center; font-weight:bold; left:<%= OrderPos3 %>px; width:<% =OrderWidth3 %>px; color:#FFFFFF; border:1px solid <%=HQSiteColor2%>; background-color:<%=HQSiteColor2%>;">Quantity</span>
	    <span class="eventline" style="position:absolute; padding-top:2px; text-align:right; font-weight:bold; left:<%= OrderPos4 %>px; width:<% =OrderWidth4 %>px; color:#FFFFFF; border:1px solid <%=HQSiteColor2%>; background-color:<%=HQSiteColor2%>;">Amount</span>
		</div>
		<div style="width:100%; margin-top:0px;">
	  	<%

			FOR ItemNo=1 TO 9
					IF TRIM(Item_Name(ItemNo))<>"" THEN 
		  				%>
		  				<br> 	  
							<span class="eventline" style="position:absolute; border:1px solid <%=HQSiteColor2%>; text-align:center; left:<% =OrderPos1 %>px; width:<% =OrderWidth1 %>px; height:15px; color:#000000; background-color:<%= TableColor1 %>;"><%= ItemNo %></span>
							<span class="eventline" style="position:absolute; border:1px solid <%=HQSiteColor2%>; padding-left:2px; text-align:left; left:<% =OrderPos2 %>px; width:<% =OrderWidth2 %>px; height:15px; color:#000000; background-color:<%= TableColor1 %>;"><%= Item_Name(ItemNo) %></span>
							<span class="eventline" style="position:absolute; border:1px solid <%=HQSiteColor2%>; text-align:center; left:<% =OrderPos3 %>px; width:<% =OrderWidth3 %>px; height:15px; color:#000000; background-color:<%= TableColor1 %>"><%= Quantity(ItemNo) %></span>
							<span class="eventline" style="position:absolute; border:1px solid <%=HQSiteColor2%>; text-align:right; left:<% =OrderPos4 %>px; width:<% =OrderWidth4 %>px; height:15px; color:#000000; background-color:<%= TableColor1 %>"><%= formatcurrency(Amount(ItemNo)*Quantity(ItemNo),2) %></span>
		  				<%
					END IF
			NEXT  
			
			%>
			<br>
			<span class="eventline" style="position:absolute; border:1px solid <%=HQSiteColor2%>; text-align:center; left:<% =OrderPos3 %>px; width:<% =OrderWidth3 %>px; height:15px; color:#000000; background-color:<%= TableColor1 %>">TOTAL ALL</span>
			<span class="eventline" style="position:absolute; border:1px solid <%=HQSiteColor2%>; text-align:right; left:<% =OrderPos4 %>px; width:<% =OrderWidth4 %>px; height:15px; color:#000000; background-color:<%= TableColor1 %>"><% =formatcurrency( (Amount(1)*Quantity(1)) + (Amount(2)*Quantity(2)) + (Amount(3)*Quantity(3)) + (Amount(4)*Quantity(4)) + (Amount(5)*Quantity(5)) + (Amount(6)*Quantity(6)) + (Amount(7)*Quantity(7)) + (Amount(8)*Quantity(8)) + (Amount(9)*Quantity(9)),2) %></span>
			<br>	
		</div>	
		<%



		' --- Changed 7-30-2015 because ThisInvAmt was not accounting for when the person purchased multiple quantity of the same item --
   	Dim ThisInvAmt
		ThisInvAmt = (Amount(1)*Quantity(1)) + (Amount(2)*Quantity(2)) + (Amount(3)*Quantity(3)) + (Amount(4)*Quantity(4)) + (Amount(5)*Quantity(5)) + (Amount(6)*Quantity(6)) + (Amount(7)*Quantity(7)) + (Amount(8)*Quantity(8)) + (Amount(9)*Quantity(9))
	
	
		' --- TEST ---
		' ThisInvAmt=0
		IF ThisInvAmt > 0 THEN 
				%>
				<div style="width:100%; margin-top:5px; margin-left:0px;">
					<%
					IF sHQAccount=true THEN 
							%>
							<span class="gentext" ><b>Click 'Pay Now' to pay for your Online Registration using your Credit Card.</b></span>
							<%
					ELSE
							%>
							<span class="gentext" ><b>Click 'Pay Now' to pay for your Online Registration using your PayPal account or Credit Card. <u> If you do have a PayPal account</u> and do not wish to establish one for future transactions, click on the ‘Continue’ button near the credit card images on the first PayPal screen.</b></span>		
							<br>
							<span class="gentext" style="color:red; font-size:14px; text-align:center;"><b>IMPORTANT - TO FINALIZE REGISTRATION !!!</b></span>		
							<br>
							<span class="gentext" style="color:red; font-size:11px; text-align:center;">Once you have made payment in the PayPal processor, to complete your registration <u>you must press the button</u> located on the final PayPal screen titled: <b>####  IMPORTANT - CLICK HERE TO FINALIZE REGISTRATION ####</b> .</span>		
							<%
					END IF
					%>
				</div>
				<%						


				' --- TEST ---
				' sPayType="Card"
				IF sPayType="Card" THEN 		' --- Send to HQ CC processor ---

						IF LCASE(LEFT(Session("sRelease"),4))="adlt" OR LCASE(LEFT(Session("sRelease"),4))="pbct" THEN ppf=1
						
						
						%>
			  		<div style="width:100%; margin-top:5px; margin-left:0px; text-align:center; padding-bottom:20px;">
							<form action="/rankings/<%=CardFileName%>?action=new&ppf=<%=ppf%>&CCAmount=<%=ThisInvAmt%>&sOrderNo=<%=sOrderNo%>&sMemberID=<%=SMemberID%>&sTourID=<%=sTourID%>" method=post>
		        		<span class="gentext" style="text-align:center;">
				  				<input type="submit" name="CreditCard" value="Pay Now" style="width:9em" title="Click here to proceed to secure Payment Page">
			  	  	  </span>
					    </form>
					  </div>
					  <%


				ELSE								' --- PayPal Button and POST --- 

						notify_URL="http://usawaterski.org/rankings/PayPal_IPN.asp?sMemberID="&sMemberID&"&sTourID="&sTourID
				    %>
						<div style="width:100%; margin-top:10px; margin-left:0px;">
				 			<span class="gentext" style="text-align:center;">
								<form action="<%=sPayPalActionURL%>" method=POST name="PPForm">
									<input type=hidden value="<%=sOrderNo%>" name="invoice">
									<input type=hidden value="<%=ReturnURL%>" name="return">
									<input type=hidden value="<%=ReturnURLBad%>" name="cancel_return">
									<input type=hidden value="<%=notify_URL%>" name="notify_URL">

									<input type=hidden value="2" name="rm">
									<input type=hidden value="####  IMPORTANT - CLICK HERE TO FINALIZE REGISTRATION ####" name="cbt">

									<input type=hidden value="<%=simage_url%>" name="cpp_header_image">
									<input type="hidden" name="no_shipping" value="0">
									<input type="hidden" name="no_note" value="1">
									<input type="hidden" name="lc" value="US">
									<input type="hidden" name="currency_code" value="USD">
									<input type="hidden" name="bn" value="PP-BuyNowBF">
							
									<input type=hidden value="<%=sPayPalAct%>" name="business">
									<input type=hidden value="_cart" name="cmd">
									<input type=hidden value="1" name="upload">
									<%
							
									FOR ItemNo=1 TO 9
			  									IF TRIM(Item_Name(ItemNo))<>"" THEN	
															thisitemname="item_name_"&ItemNo
															thisamountname="amount_"&ItemNo
															thisquantityname="quantity_"&ItemNo 
															%>
															<input type=hidden value="<%=item_name(ItemNo)%>" name="<%=thisitemname%>"> 	
															<input type=hidden value="<%=round(amount(ItemNo),2)%>" name="<%=thisamountname%>"> 	
															<input type=hidden value="<%=quantity(ItemNo)%>" name="<%=thisquantityname%>">
															<% 
			   									END IF 
									NEXT  

									%>
		

									<input type="image" src="https://www.paypal.com/en_US/i/btn/btn_paynow_LG.gif" border="0" name="submit" alt="Make payments with PayPal - it's fast, free and secure!">
									<img alt="" border="0" src="https://www.paypal.com/en_US/i/scr/pixel.gif" width="1" height="1">
								</span>
							</div>
						</form>
						<%


						' --------------------------------------------
						' --- Displays optional Pay on Site button ---
						' --------------------------------------------

						' --- TEST ---
						' sAllowOfflinePmt=1
						IF sAllowOfflinePmt<>0 THEN
									%>
									<br>
									<div style="width:100%; margin-top:0px; margin-left:0px;">
				    		  	<form method="post" action="/rankings/<%=RegFileName%>" id="OfflinePaymentForm">
											<input type="hidden" name="nav" value=7>
		
											<span class="gentext" style="text-align:center; font-size:11px;"><b>OPTIONAL PAYMENT METHOD</b></span>
											<br>
											<span class="gentext" >This tournament has elected to allow payment by mail or at the site.  <b>If you elect to pay in this manner your registration will NOT be complete</b> and you may not be allowed to compete in this tournament. Depending on when your payment is received, late fees may apply.  It is your responsibility to make payment and confirm your eligibility.</span>
											<span class="gentext" style="text-align:center; margin-top:10px;">
												<input type="submit" value="Pay On Site" style="width:10em;">
											</span>	
										</form>	
									</div>
									<%
						END IF
						%>	
						<hr>
						<%

				END IF		' --- Bottom of condition for PayPal or Card ---



		ELSEIF ThisInvAmt <= 0 THEN 
		
				%>
				<div style="width:100%; margin-top:0px; margin-left:0px;">
    			<span class="gentext" style="font-weight:bold; text-align:center;">
						<%
						IF ThisInvAmt<0 THEN 
								%>	
								PayPal refunds cannot be initiated through this system. Please contact the tournament registrar, LOC or the contact listed on your PayPal receipt.
								<% 
								sPayType="ByPass"
						ELSE  
								%>	
								No Payment is Due.  Press the 'Continue' button to go to receipt tab.
								<% 
								sPayType="NoSale" 
						END IF  
						%>	
		    	</span>
			  	<%


				' --- ByPasses Payment altogether --- %>
		  		<span class="gentext" style="text-align:center; width="100%;">
		  			<form action="http://usawaterski.org/rankings/<%=RegFileName%>" method=POST name="defaultform">
							<input type="submit" name="Continue" value="Continue" style="width:9em" title="Continue to Print Receipt Page">
			  			<input type="hidden" name="sPayType" value="<%=sPayType%>">
			  			<input type="hidden" name="sTourID" value="<%=sTourID%>">
			  			<input type="hidden" name="sMemberID" value="<%=sMemberID%>">
			  			<input type="hidden" name="sOrderNo" value="<%=sOrderNo%>">
			  			<input type="hidden" name="nav" value="7">
		    			<br>
		    			
						</form>
		  		</span>
					<hr>
				</div>
	      <%

		END IF 		' --- Bottom of ThisInvAmt condition ---





	' -----------------------------------------------------------------------------------------------------------------------
	' --- Displays button and additional fields for confirming payment, checks, cash and credits when AdminCode is active ---
	' -----------------------------------------------------------------------------------------------------------------------

			
	IF (sDispDebugButtons=true OR adminmenulevel>=20 OR TestValidAdminCode) THEN

			' --- Simulates payment variables returned from PayPal SUCCESS--- 
			%>
			<div style="width:100%; margin-top:10px; margin-left:0px; padding-bottom:10px;">
		 		<form name="PaymentForm3" method="post" action="/rankings/<%=RegFileName%>" id="PaymentlForm3">
  				<input type="hidden" name="sTourID" value="<%=sTourID%>">
	  			<input type="hidden" name="sMemberID" value="<%=sMemberID%>">
	  			<input type="hidden" name="sOrderNo" value="<%=sOrderNo%>">
	  			<input type="hidden" name="nav" value="7">
	  			<input type="hidden" name="SpecialAction" value="Y">


			  	<span class="gentext" style="color:red; font-size:12px; text-align:center;"><b>ATTENTION REGISTRAR</b></span>
			    <br>
		  		<span class="gentext" style="text-align:left;">Use the field and drop down below to record payments made by Cash or Check.  If you received an email from PayPal acknowledging this member's payment to your PayPal account, but the online registration system still shows a <u>Balance Due</u>, set the dropdown to 'Confirm PayPal Payment' and press 'Submit' to finalize the member's registration and acknowledge that you received the payment amount shown in the box below.</span>			    
			    <br>
			    <span class="gentext" style="width:100px; text-align:right; color:red;"><b>Amount:</b></span>
					<span class="gentext" style="width:100px; text-align:left;">	
						<input type="text" name="sPayAmount" value="<%=formatnumber(ThisInvAmt,2)%>" MAXLENGTH=7 size=7 style="text-align:right">
		    	</span>
		    	<span class="gentext" style="width:120px; text-align:right; color:red;"><b>Payment Type:</span>
					<span class="gentext" style="width:200px; text-align:left;">
						<select name="sPayType" value="<%=sPayType%>" style="width:15em">
					  	<option value ="" <%IF sPayType = "" THEN Response.Write(" selected ")%> >No Payment</Option><br>
			  			<option value ="PayPal" <%IF sPayType = "PayPal" THEN Response.Write(" selected ")%> >Confirm PayPal Payment</Option><br>
			  			<option value ="Check" <%IF sPayType = "Check" THEN Response.Write(" selected ")%> >Receive Check</Option><br>
			  			<option value ="Cash" <%IF sPayType = "Cash" THEN Response.Write(" selected ")%> >Receive Cash</Option><br>
			  			<option value ="Refund" <%IF sPayType = "Refund" THEN Response.Write(" selected ")%> >Issue Refund</Option><br>
						</select>
		    	</span>
		    	<span class="gentext" style="width:170px; text-align:center; ">
						<input type="submit" class="AdminButtonStyle" name="PayTestSuccess " value="Submit" title="Press SUBMIT to record amount shown according to drop down setting" <%=ByPassButtonStatus%>>
		    	</span>
				</form>
			</div>
			<%
	END IF	


	%>
	</div>	
  <%
  
 
 
 
 


' ------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------
' -----------------------------  RECEIPT AND NOTICES  -------------------------------- 
' ------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------
' ------------------------------------------------------------------------------------


	%>
 	<div style="width:100%" class="<% IF nav=7 THEN response.write("accordionHeaderSelected") ELSE response.write("accordionHeader") END IF %>">
		<span style="text-align:left; width:30%;">STEP 7 - Receipt</span>
	</div>	

  <!-- <div id="RegPanel6" class="tour_div" style="padding-top:10px; display:block;"> -->
  <div id="RegPanel7" class="tour_div" style="padding-top:10px; display:<% IF nav=7 THEN response.write("block") ELSE response.write("none") END IF %>;">
		<div class="olrsecheading" style="width:98%; margin-top:10px;">
			<span class="eventline" style="position:absolute; text-align:left; font-weight:bold; font-size:16px; left:0px; padding-left:10px; height:20px;">Registration Complete</span>
		</div>
		
		<div style="width:100%; margin-top:15px;">
			<span style="font-size:12px;"><b>You have completed your registration for:</b></span>
			<span style="font-size:12px; color:blue; width:300px;"><b><% =sTourName %></b></span>
	  </div> 	
		<div style="width:95%; padding-left:15px; margin-top:15px;  border:0px solid;">
			<span class="gentext" >1. <%=ReceiptNote1%></span>
			<br>
			<span class="gentext" >2. <%=ReceiptNote2%></span>
			<br>
			<span class="gentext" >3. <%=ReceiptNote3%></span>
			<br>
			<span class="gentext" >4. <%=ReceiptNote4%></span>
			<br>
			<span class="gentext" >5. <%=ReceiptNote5%></span>
			<br>
	  </div> 	
		<div style="width:98%; margin-top:15px; border:0px solid;">
			<span class="gentext" style="text-align:center; font-size:12px; color:red; font-weight:bold;">Session Ended - Do not use expired pages!</span>
			<br><br>
			<span class="gentext" style="text-align:center; font-size:12px;"><b>For questions, please contact the Tournament Registrar at <%=sTRegistrarPhone%>.</b></span>
			<br><br> 
			<span class="gentext" style="text-align:center; font-size:14px; ><b>Thank you.</b></span>
		</div>
		<%

		' --- RECEIPT BUTTONS ---
		%>
		<div style="width:98%; margin-top:15px; border:0px solid;">
			<span class="spanbuttons" style="border:0px solid red; padding-bottom:15px; margin-bottom:10px; width:98%;">
				<form name="ReceiptForm" method="post" action="/rankings/<%=RegFileName%>?sRunByWhat=ReturnToMainMenu" id="ReceiptForm">
	  			<input type="submit" name="Main Menu" value="Main Menu" style="position:absolute; left:165px; width:9em" title="Leave registration and return to the main menu. IMPORTANT: Once you return to main menu, your 'session' will end and you must log in again to print receipt or view information.">
				</form>
				<%	

  			' --- Lets Admin Users return to Member Page ---
	  		IF TestValidAdminCode=true OR adminmenulevel >= 20 THEN 
	  				%>
	      		<form name="ReceiptForm3" method="post" action="/rankings/<%=RegFileName%>" id="ReceiptForm3">
	  					<input type="hidden" name="sTourID" value="<%=sTourID%>">
	  					<input type="hidden" name="sMemberID" value="<%=sMemberID%>">
	  					<input type="hidden" name="nav" value="2">
							<input type="submit" class="AdminButtonStyle" name="BackToMember" value="New Member" style="position:absolute; left:320px; width:9em" title="Admin users - Return to Member tab to select a NEW member without losing your Administrative login information." <%=MainButtonStatus%>>
		  			</form>
						<%
	  		ELSE 
	  				%>
	      		<form name="ReceiptForm3" method="post" action="/rankings/<%=RegFileName%>" id="ReceiptForm3">
	  					<input type="hidden" name="sTourID" value="<%=sTourID%>">
	  					<input type="hidden" name="sRunByWhat" value="NewMember">
							<input type="submit" class="UserButtonStyle" name="NewMember" value="New Member" style="position:absolute; left:320px; width:9em" title="Register a different member." <%=MainButtonStatus%>>
		  			</form>
						<%
	  		END IF

 				%>	
				<form name="NewEntryForm" method="post" action="/rankings/<%=RegFileName%>?sRunByWhat=Tour" id="NewEntryForm">
		 	 		<input type="submit" class="" name="NewTour" value="New Entry" style="width:9em; position:absolute; left:490px;" title="Create another online entry for the same member" <%=MainButtonStatus%>>
				</form>

				<form action="/rankings/<%=RegFileName%>?sRunByWhat=Print" method="post" target="_blank">
					<input type="submit" value="Print Receipt" style="width:9em; position:absolute; left:490px;" title="Displays complete entry receipt on screen, which may be printed for your records.">
				</form>	
			</span>
		</div>		<!-- Buttons -->
	
	</div>		<!-- Bottom of Tab Div -->

</div>			<!-- Outer Accordion div -->
<%




END SUB




 



' --------------------------------
   SUB SetHiddenFinancialVariables
' --------------------------------

	' --- These are the variables that are on the Financial Form --- 
	' --- For 2016 removed fAWSEFCheck
	%>
	<input type="hidden" name="sAWSEFDonation" value="<%=sAWSEFDonation%>">
	<input type="hidden" name="fOfficial" value="<%=sOfficial%>">
	<input type="hidden" name="fClubMemb" value="<%=sClubMemb%>">
	<input type="hidden" name="fClubCode" value="<%=sClubCode%>">
	<input type="hidden" name="sBanquetQty" value="<% =sBanquetQty %>">
	
	<input type="hidden" name="sOF1Qty" value="<% =sOF1Qty %>">
	<input type="hidden" name="sOF2Qty" value="<% =sOF2Qty %>">
	<input type="hidden" name="sOF3Qty" value="<% =sOF3Qty %>">
	<input type="hidden" name="sOF4Qty" value="<% =sOF4Qty %>">
	<input type="hidden" name="sOF5Qty" value="<% =sOF5Qty %>">
	<input type="hidden" name="sOF6Qty" value="<% =sOF6Qty %>">
	<input type="hidden" name="sOF7Qty" value="<% =sOF7Qty %>">
	<input type="hidden" name="sOF8Qty" value="<% =sOF8Qty %>">
	<input type="hidden" name="sOF9Qty" value="<% =sOF9Qty %>">
	<input type="hidden" name="sOF10Qty" value="<% =sOF10Qty %>">
	<input type="hidden" name="sOF1Fee" value="<% =sOF1Fee %>">
	<input type="hidden" name="sOF2Fee" value="<% =sOF2Fee %>">
	<input type="hidden" name="sOF3Fee" value="<% =sOF3Fee %>">
	<input type="hidden" name="sOF4Fee" value="<% =sOF4Fee %>">
	<input type="hidden" name="sOF5Fee" value="<% =sOF5Fee %>">
	<input type="hidden" name="sOF6Fee" value="<% =sOF6Fee %>">
	<input type="hidden" name="sOF7Fee" value="<% =sOF7Fee %>">
	<input type="hidden" name="sOF8Fee" value="<% =sOF8Fee %>">
	<input type="hidden" name="sOF9Fee" value="<% =sOF9Fee %>">
	<input type="hidden" name="sOF10Fee" value="<% =sOF10Fee %>">
	<input type="hidden" name="sEntryType" value="<% =sEntryType %>">
	<input type="hidden" name="sMembRegDate" value="<% =sMembRegDate %>">
	<%

END SUB



' ------------------------------------------
  SUB SetHiddenFinancialOverrideVariables 
' ------------------------------------------  
  	
%><input type="hidden" name="sMoneyOverride" value="<%=sMoneyOverride%>"><%


END SUB






' -----------------------------
   SUB SetHiddenEntryVariables	  
' -----------------------------

	  FOR EvtNo = 1 TO TotEv
				fSelectEvent="fSelectEvent"&EvtNo
				fDiv="fDiv"&EvtNo  				 
				fFeeClass="fFeeClass"&EvtNo
				fFeeRounds="fFeeRounds"&EvtNo
				fQfyOverride="fQfyOverride"&EvtNo  
				fBoat="fBoat"&EvtNo  
				fSkill="fSkill"&EvtNo  
				%>
		  	<input type="hidden" name="<%= fSelectEvent %>" value="<% =sSelectEvent(EvtNo) %>">
		  	<input type="hidden" name="<%=fDiv%>" value="<% =sDiv(EvtNo) %>">
		  	<input type="hidden" name="<%=fFeeClass%>" value="<% =sFeeClass(EvtNo) %>">
		  	<input type="hidden" name="<%=fFeeRounds%>" value="<% =sFeeRounds(EvtNo) %>">
		  	<input type="hidden" name="<%=fSkill%>" value="<%=sSkill(EvtNo)%>">
  			<input type="hidden" name="<%=fBoat%>" value="<%=sBoat(EvtNo)%>">
		  	<%
	  NEXT 
	  %>
	  <input type="hidden" name="sRampHeight" value="<% =sRampHeight %>">
	  <input type="hidden" name="sRegFeeCalcCode" value="<% =sRegFeeCalcCode %>">
	  <%


END SUB


' ------------------------------------
   SUB SetHiddenEntryOverrideVariables
' -------------------------------------

	  FOR EvtNo = 1 TO TotEv
				fQfyOverride="fQfyOverride"&EvtNo  
				%>
		  	<input type="hidden" name="<%=fQfyOverride%>" value="<% =sQfyOverride(EvtNo) %>">
		  	<%
	  NEXT 
	  %>
	  <input type="hidden" name="sRegionalOverride" value="<% =sRegionalOverride %>">
	  <%


END SUB




' -------------------------------
  SUB SetHiddenWaiverVariables 
' -------------------------------

		%>
	  <input type="hidden" name="sWaiverCode" value="<% =sWaiverCode %>">
	  <input type="hidden" name="sSignWaiver" value="<% =sSignWaiver %>">
		<%
  
END SUB



' ------------------------------
  SUB DisplayPertinentVariables
' ------------------------------

response.write("IN Display ")%><br><%
response.write("sEntryFee = "&sEntryFee)%><br><%
response.write("sLateFeeTot = "&sLateFeeTot)%><br><%
response.write("sAWSEFDonation ="&sAWSEFDonation)%><br><%
response.write("sBanquetTot = "&sBanquetTot)%><br><%
response.write("sJrDiscAmt = "&sJrDiscAmt)%><br><%
response.write("sSrDiscAmt = "&sSrDiscAmt)%><br><%
response.write("sClubDiscAmt = "&sClubDiscAmt)%><br><%
response.write("sOffDiscAmt = "&sOffDiscAmt)%><br><%

response.write("sTotalFormFees = "&sTotalFormFees)%><br><%
response.write("cdbl(sTotalPreviousPayments) = "&cdbl(sTotalPreviousPayments))
'response.end

END SUB



' -----------------------
    SUB AccordRelease
' -----------------------

	waivernav=Request("waivernav")
	IF waivernav="" THEN waivernav=1



' --- New 4-28-2013 - Gets SPECIAL WAIVER info from table based on SiteID rather than hard coding specific tournaments ---
Dim swaiverSQL, sSpecialWaiverHeadline, sSpecialReleaseBannerText
swaiverSQL = "SELECT SpecialWaiverCode, SpecialWaiverHeadline, SpecialReleaseBannerText FROM usawsrank.TourExtras TE"
swaiverSQL = swaiverSQL + " JOIN sanctions.dbo.TSchedul AS TS"
swaiverSQL = swaiverSQL + "   ON SiteID=TS.TSiteID"
swaiverSQL = swaiverSQL + " WHERE LEFT(TS.TournAppID,6)='"&LEFT(sTourID,6)&"'"

Set rswaiver=Server.CreateObject("ADODB.recordset")
rswaiver.open swaiverSQL, sConnectionToTRATable, 3, 1

testwaiver=false
IF testwaiver=true AND sMemberID="000001151" THEN
		Response.write("<br>Found = ")
		response.write(NOT(rswaiver.eof))
		response.write("<br>rswaiver(SpecialWaiverHeadline) = "&rswaiver("SpecialWaiverHeadline"))
END IF

IF NOT(rswaiver.EOF) THEN
		sSpecialWaiverCode=rswaiver("SpecialWaiverCode")
		sSpecialWaiverHeadline=rswaiver("SpecialWaiverHeadline")
		sSpecialReleaseBannerText=rswaiver("SpecialReleaseBannerText")
		'response.write("<br>Line 2526 - sSpecialWaiverCode="&sSpecialWaiverCode)
END IF


' --- TESTING CODE and code to add a waiver for a new site ---
' SELECT * FROM usawsrank.RegisterGenNew WHERE TourID='13S999' AND MemberID='000001151'
' DELETE FROM usawsrank.RegisterGenNew WHERE TourID='13S999' AND MemberID='000001151'
' ALTER TABLE usawsrank.TourExtras ADD SpecialWaiverCode CHAR(8)
' ALTER TABLE usawsrank.TourExtras ADD SpecialWaiverHeadline VARCHAR(70)
' ALTER TABLE usawsrank.TourExtras ADD SpecialReleaseBannerText VARCHAR(70)

' UPDATE TE SET SpecialWaiverCode='pbct2012' FROM usawsrank.TourExtras TE WHERE SiteID='USAS0364'
' UPDATE TE SET SpecialWaiverHeadline='AMATEUR ATHLETIC WAIVER AND RELEASE OF LIABILITY' FROM usawsrank.TourExtras TE WHERE SiteID='USAS0364'
' UPDATE TE SET SpecialReleaseBannerText='Palm Beach County Waiver - Read Carefully' FROM usawsrank.TourExtras TE WHERE SiteID='USAS0364'

' --- Changed 4-28-2013 to use SiteID approach ---
'	IF LEFT(sTourID,6)="13S999" OR LEFT(sTourID,6)="12S999" OR LEFT(sTourID,6)="13S090" OR LEFT(sTourID,6)="13S091" OR LEFT(sTourID,6)="13S148" OR LEFT(sTourID,6)="13S130" OR LEFT(sTourID,6)="13S150" OR LEFT(sTourID,6)="13S151" OR LEFT(sTourID,6)="14S036" OR LEFT(sTourID,6)="14S037" OR LEFT(sTourID,6)="14S041" OR LEFT(sTourID,6)="13S195" OR LEFT(sTourID,6)="13S196" OR LEFT(sTourID,6)="13S197" OR LEFT(sTourID,6)="13S198" THEN

		' --- OR LEFT(sTourID,6)="11M999" 
		' --- Per Dale Stevens WPBSC 1-24-2013 ---
		' 13S090, 13S091, 13S148, 13S130, 13S150, 13S151, 14S036, 14S037, 14S041
		' --- New ones
		' 13S195, 13S196, 13S197, 13S198

		'	sSpecialWaiverCode="pbct2012"
		'	sSpecialWaiverHeadline="AMATEUR ATHLETIC WAIVER AND RELEASE OF LIABILITY"
		'	sSpecialReleaseBannerText="Palm Beach County Waiver - Read Carefully"
'	END IF



	' --- Test for tournaments to use special survey form ---
	sSurveyForm=""
	SELECT CASE LEFT(sTourID,6)
			CASE "12S999"
					sSurveyForm="12S999"			
			CASE "11M999"
					sSurveyForm="12S999"
			CASE "13S999"
					sSurveyForm="13S999"
			CASE "15S999"
					sSurveyForm="15S999"
	END SELECT


'Response.write("<br>sSurveyForm = "&sSurveyForm)


	' --- Request variables related to waiver ---
	sRelease = TRIM(Request("fRelease"))
	sSignWaiver = sqlclean(TRIM(Request("sSignWaiver")))
	sReleaseType = TRIM(Request("sReleaseType"))

	' --- If blank and NOT an Admin User then OK to default to Electronic ---
	IF sReleaseType = "" AND (NOT TestValidAdminCode) THEN 
			sReleaseType="Electronic"
	ELSEIF sReleaseType = "" AND TestValidAdminCode THEN 
			sReleaseType="None"
	END IF

	' - CASE 1 - Notice
	' - CASE 3 - Std USA Water Ski waiver 
	' - CASE 3 - Special Waiver
	' - CASE 4 - Waiver Acknowledgement
	' - CASE 5 - Save and if survey display page 1 otherwise set waivernav=10 to redirect
	' - CASE 6 - If survey display thank you
	' - CASE 10 - Redirect (used?)


	IF sMemberID="700144639" THEN
			response.write("<br>sReleaseType = "&sReleaseType)
			response.write("<br>sSignWaiver = "&sSignWaiver) 
	END IF

	' --- If Electronic then check to make sure either age>=18 or someone has signed the waiver as the guardian ---
	IF sReleaseType = "Electronic" AND (Session("sMembAge") >= 18 OR (Session("sMembAge") < 18 AND TRIM(sSignWaiver)<>"")) THEN sElectronicOK = true	


	ReleaseAccepted = "N"	
	IF sRelease="Accept" AND (sElectronicOK OR sReleaseType = "Paper" OR sReleaseType = "None") THEN
			ReleaseAccepted = "Y"
	END IF		
	IF waivernav=3 AND TRIM(sSpecialWaiverCode)="" AND ReleaseAccepted="Y" THEN
			' --- IF Release accepted AND one of the methods is met ---
			waivernav=4

	ELSEIF waivernav=4 AND TRIM(sSpecialWaiverCode)<>"" AND ReleaseAccepted="Y" THEN
			' --- Release accepted AND SPECIAL waiver AND one of the methods is met ---
			waivernav=4			

	ELSEIF waivernav=4 AND TRIM(sSpecialWaiverCode)<>"" THEN
			' --- IF Release NOT accepted AND in SPECIAL waiver section then redisplay waiver ---
			waivernav=3	

	ELSEIF waivernav=3 AND TRIM(sSpecialWaiverCode)="" THEN
			' --- Release NOT accepted and no Special waiver then redisplay waiver---
			waivernav=2	

	ELSEIF waivernav=5 THEN
			' --- Saves the registration waiver values to the RegGen table ---
			SaveWaiverValuesToTable

			' --- If survey then display survey page, else redirect to next tab in registration (Payment)
			IF TRIM(sSurveyForm)<>"" THEN
					BeginTournamentSurvey
			ELSE
					' --- No survey so exit waiver section to Payment page ---
					waivernav=10
			END IF
	ELSEIF waivernav=6 THEN		' --- Thank you page of survey which could only occur if survey ---
			BeginTournamentSurvey

	END IF		





SELECT CASE LCASE(TRIM(waivernav))
		CASE 1

			' -----------------------------------------------------------
			' ---------  Initial Heads Up Dialog Box on Warning  --------	
			' -----------------------------------------------------------
			%>
			<br>

			<TABLE BORDER="2" BORDERCOLOR="black" ALIGN="CENTER" BGCOLOR="<% =TableColor1 %>" width="100%">
	  		<tr>
	      	<td BGCOLOR="red"><center><font face=<% =font1 %> color="#FFFFFF" size="4"><b>Important Notice !!</b></font></TD>
	  		</TR>  
 
			  <tr>
					<td VALIGN="top">
  					<TABLE BORDER="0" VALIGN="top" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="<%=TableColor1%>" width=100%>
			   			<tr>
			      		<td VALIGN="top" ALIGN="center">
									<br>
									<font color="<% =TextColor1 %>" face="<% =font1 %>" size="4"><b><i>Read This Before Continuing !!</i></b></font>
									<br><br>
									<font face="<% =font1 %>" size="1">You may not participate in any USA Water Ski sanctioned event without accepting the terms of the following PARTICIPANT WAIVER AND RELEASE OF LIABILITY, ASSUMPTION OF RISK AND INDEMNITY.</font>
									<br><br>
									<font color="red" face="<% =font1 %>" size="2"><b>You must be 18 years old.</b></font> 			
									<br><br>	
									<font face="<% =font1 %>" size="1"><b>Unless you are accepting it as a parent of a minor, you may NOT accept the 'RELEASE' on behalf of another person.</b></font> 
									<br><br>
									<font face="<% =font1 %>" size="1">Accepting the 'RELEASE' for another adult (even a spouse), or accepting it on behalf of a minor for whom you are not the legal guardian, is strictly prohibited.  Such action may be subject to civil liability of criminal prosecution under <b> state and federal laws.</b></font>
			    			</td>
			  			</tr>
			  			<tr>
			    			<td align="center">
									<br>
									<form action="/rankings/<%=RegFileName%>?nav=5&waivernav=2" method="post">
				  						<center><input type="submit" value=" Continue "></center>
									</form>
			    			</td>	
			  			</tr>
						</TABLE>
		    	</TD>
		  	</TR>
			</TABLE>   
			<br>
			<% 

	
		CASE 2

			' --------------------------------------------
			'--- Display standard USA Water Ski Waiver ---
			' --------------------------------------------
			DefineWaiverSpecs
		
			%>
			<br>
			<TABLE BORDER="2" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="<% =TableColor1 %>" width=100%>
		  	<tr>
		      <td BGCOLOR="red"><center><font face=<% =font1 %> color="#FFFFFF" size="4"><b>Waiver and Release Form</b></font></TD>
		  	</TR>  
 		  	<tr>
					<td VALIGN="top">
  					<TABLE BORDER="0" VALIGN="top" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="<%=TableColor1%>" width=100%><%

						' ------------------------------------------------------------------------------------------
						' ----------  Displays release and sets FORM ACTION to rerun this condition  ---------------	
						' ------------------------------------------------------------------------------------------ 



						' --- Determines if a special waiver is to be used for the tournament in addition to std USA Waterski waiver ---
						'IF LEFT(sTourID,6)="12S999" OR LEFT(sTourID,6)="11M999" THEN 
						IF sSpecialWaiverCode="pbct2012" THEN
								WhichWaiverNav="3"
						ELSE
								WhichWaiverNav="4"
						END IF
						
						
						%>
						<form action = "/rankings/<%=RegFileName%>?waivernav=<%=WhichWaiverNav%>&nav=5" method="post">
	  		  		<tr>
								<td align=center>	
	 	   						<font face=<% =font1 %> size="4" ><b>PARTICIPANT WAIVER AND RELEASE OF LIABILITY,</b></font><br>
		   						<font face=<% =font1 %> size="4"><b>ASSUMPTION OF RISK AND INDEMNITY AGREEMENT</b></font>
		   						<br>
		   						<font face=<% =font1 %> size="2"><b><% =sWaiverSubTitle %></b></font>
		   						<br><br>

									<font face=<% =font1 %> color="<% =TextColor2 %>" size="3"><b><% =sTourName %></font></b>
									<br><br>
									<font face=<% =font1 %> size="2"><b>MemberID = </font><font color="<% =textcolor2 %>" face=<% =font1 %> size="2"><%=sMemberID%>
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="<% =textcolor1 %>" face=<% =font1 %> size="2">Participant:</font>
									<font color="<% =textcolor2 %>" face=<% =font1 %> size="2"><% =sFirstname %>&nbsp;<% =sLastName %></font></b><br>
								</td>
							</tr>

  		  			<tr>
		   					<td>
									<br>
		     					<p>
 										<font color="<% =textcolor1 %>" size="1" face=<% =font1 %>><left>
 											<%

			  						Set objfso = CreateObject("Scripting.FileSystemObject")
			  						IF objfso.FileExists(PathtoWaivers & "\waiver-"&sWaiverCode&".txt") THEN
												SET objstream=objFSO.opentextfile(PathtoWaivers & "\waiver-"&sWaiverCode&".txt")

												IF NOT objstream.atendofstream THEN
														DO WHILE not objstream.atendofstream
																response.write(objstream.readline)
					   										response.write("<br>")
														LOOP
												END IF
												objstream.close
			  						END IF 

			  						%>
										</font>
		    					</p>
		   					</td>
  		  			</tr>
							<tr>
		    				<td align="center">
		    				<%
									IF TestValidAdminCode OR adminmenulevel>=20 THEN 
											%>
			   							<br>
			   							<font size="3" color="red" ><b>As 'Tournament Registrar', you are responsible for collecting 
											<br>waivers and returning them to USA Water Ski competition services dept.</b></font>
											<%
									ELSE 
											%>
			   							<br>
			   							<font size="3" color="red" ><b>The name listed above must be the person completing this form.</b></font>
			   							<br>
			   							<font size="3" color="red" ><b>Minors under 18 Years may NOT accept liability waiver.</b></font>
			   							<%
									END IF 
									%>
		    				</td>
		  				</tr>

		  				<tr>
		  					<td align="center">
		     					<br>
		     					<%
		     					IF Session("sMembAge") < 18 AND (NOT TestValidAdminCode) THEN  
		     							%>
											<font color="<% =textcolor3 %>" size="2"><b>Name of Parent or Guardian acccepting waiver on behalf of </b></font>
											<font color="<% =textcolor2 %>" size="2"><b><%=sFirstName%>&nbsp;<%=sLastName%></b></font>
											<br>
											<input type="text" name="sSignWaiver" value= "<% =sSignWaiver %>" maxlength=30 size="30" >
											<%
		     					ELSEIF Session("sMembAge") >= 18 AND (NOT TestValidAdminCode) THEN  
		     							%>
											<font color="<% =textcolor1 %>" size="2"><b>By acccepting this waiver I acknowledge that I am the 'PARTICIPANT' listed above.	</b></font>
											<br>
											<%
		     					ELSEIF TestValidAdminCode THEN  
		     							%>
											<font color="<% =textcolor1 %>" size="2"><b>Please select current 'Status' of Waiver form.</b></font>
											<br><%
		     					END IF 
		     					%>
		   					</td>
		   				</tr><%

							' -----    DECLINE USA Water Ski WAIVER ------
							IF sRelease="Decline" THEN 
									%>
		   						<tr>
		   							<td align="center">
			   							<font size="3" color="red" ><b>You May Not Enter Tournament Without Accepting the Waiver and Release</b></font>
		   							</td>
		   						</tr>
		   						<%
							END IF

							' -----    Admin Level and No Waiver Type Selected  ------
							IF sReleaseType="" AND adminmenulevel >=20  THEN  
									%>		  
		   						<tr>
		   							<td align="center">
											<br>
											<font size="4" color="red" ><b>Please select a WAIVER TYPE</b></font>
		   							</td>
		   						</tr>
		   						<%
							END IF  

							' --- Displays options for type of waiver ---
							IF adminmenulevel >= 10 OR TestValidAdminCode THEN  
									%>
  		  					<tr>
										<td align="center">
											<%
											IF adminmenulevel >= 10 THEN 
													%> 
			   									<input type=radio NAME="sReleaseType" VALUE="Electronic" <% IF sReleaseType="Electronic" THEN response.write "checked"%> >
			   									<FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR=<% =textcolor1 %> ><b>Electronic</b></font>
			   									<%
											END IF 
											%>
											<input type=radio NAME="sReleaseType" VALUE="Paper" <% IF sReleaseType="Paper" THEN response.write "checked"%>>
											<FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR=<% =textcolor1 %> ><b>Paper W/Signature</b></font>
											<input type=radio NAME="sReleaseType" VALUE="None" <% IF sReleaseType="None" THEN response.write "checked"%>>
											<FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR=<% =textcolor1 %> ><b>No Waiver</b></font>
										</td> 
  		  					</tr><%
							ELSE  
									%>
									<INPUT type="hidden" NAME="sReleaseType" VALUE="<%=sReleaseType%>" >
									<%
							END IF 	
							%>
		   				<tr>
		   					<td align="center">
									<br>
									<%
									IF TestValidAdminCode<>true THEN 
											%>	
											<font color="<% =textcolor1 %>" face=<% =font1 %> size="2"><b>Accept:</b></font><input type="radio" name="fRelease" <%IF sRelease="Accept" THEN Response.write("checked")%> value="Accept">
											<font color="<% =textcolor1 %>" face=<% =font1 %> size="2"><b>Decline:</b></font><input type="radio" name="fRelease" <%IF sRelease="Decline" THEN Response.write("checked")%> value="Decline">
											&nbsp;&nbsp;&nbsp;&nbsp 
											<%
									ELSE 
											%>
											<input type="hidden" name="fRelease" value="Accept">
											<%
									END IF 
									%>
									<input type="submit" value="Submit" style="width:9em">
									&nbsp;&nbsp;&nbsp;&nbsp
									<font color="<% =textcolor1 %>" face=<% =font1 %> size="2"><b>Date: <% =DATE %></b></font>
									<br><br>
		  					</td>
		  				</tr>

							</form>
			  		</TABLE>
						<br>

			  	</TD>
 		  	</TR>
			</TABLE>
		<%


	CASE 3

		
					
			' -----------------------------------------------------------------
			' --- SPECIAL WAIVER option - Does not apply to all tournaments ---
			' -----------------------------------------------------------------
			DefineWaiverSpecs			
			sRelease=""

			%>
			<br>
			<TABLE BORDER="2" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="<% =TableColor1 %>" width=100%>
		  	<tr>
		      <td BGCOLOR="orange"><center><font face=<% =font1 %> color="#FFFFFF" size="4"><b><%=sSpecialReleaseBannerText%></b></font></TD>
		  	</TR>  
 		  	<tr>
					<td VALIGN="top">
  					<TABLE BORDER="0" VALIGN="top" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="<%=TableColor1%>" width=100%><%

						' ------------------------------------------------------------------------------------------
						' ----------  Displays release and sets FORM ACTION to rerun this condition  ---------------	
						' ------------------------------------------------------------------------------------------ 
						%>
						<form action = "/rankings/<%=RegFileName%>?waivernav=4&nav=5" method="post">
							<input type="hidden" name="sSignWaiver" value="<%=sSignWaiver%>">
	  		  		<tr>
								<td align=center>	
	 	   						<font face=<% =font1 %> size="4" ><b><%=sSpecialWaiverHeadline%></b></font><br>
		   						<br>

									<font face=<% =font1 %> color="<% =TextColor2 %>" size="3"><b><% =sTourName %></font></b>
									<br><br>
									<font face=<% =font1 %> size="2"><b>MemberID = </font><font color="<% =textcolor2 %>" face=<% =font1 %> size="2"><%=sMemberID%>
									&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="<% =textcolor1 %>" face=<% =font1 %> size="2">Participant:</font>
									<font color="<% =textcolor2 %>" face=<% =font1 %> size="2"><% =sFirstname %>&nbsp;<% =sLastName %></font></b><br>
								</td>
							</tr>

  		  			<tr>
		   					<td>
									<br>
		     					<p>
 										<font color="<% =textcolor1 %>" size="1" face=<% =font1 %>><left>
 											<%

			  						Set objfso = CreateObject("Scripting.FileSystemObject")
			  						IF objfso.FileExists(PathtoWaivers & "\waiver-"&sSpecialWaiverCode&".txt") THEN
												SET objstream=objFSO.opentextfile(PathtoWaivers & "\waiver-"&sSpecialWaiverCode&".txt")

												IF NOT objstream.atendofstream THEN
														DO WHILE not objstream.atendofstream
																response.write(objstream.readline)
					   										response.write("<br>")
														LOOP
												END IF
												objstream.close
			  						END IF 

			  						%>
										</font>
		    					</p>
		   					</td>
  		  			</tr>
   						<tr>
   							<td align="center">
	   							<font size="3" color="red" ><b>I agree to be fully responsible for my conduct at the tournament and/or for the conduct of the minor on whose behalf I sign.</b></font>
   							</td>
   						</tr>
							<%

							' -----    DECLINE USA Water Ski WAIVER ------
							IF sRelease="Decline" THEN 
									%>
		   						<tr>
		   							<td align="center">
			   							<font size="3" color="red" ><b>You May Not Enter Without Accepting the Waiver and Release</b></font>
		   							</td>
		   						</tr>
		   						<%
							END IF


							' -----    Admin Level and No Waiver Type Selected  ------
							IF sReleaseType="" AND adminmenulevel >=20  THEN  
									%>		  
		   						<tr>
		   							<td align="center">
											<br>
											<font size="4" color="red" ><b>Please select a WAIVER TYPE</b></font>
		   							</td>
		   						</tr>
		   						<%
							END IF  

							' --- Displays options for type of waiver ---
							IF adminmenulevel >= 10 OR TestValidAdminCode THEN  
									%>
  		  					<tr>
										<td align="center">
											<%
											IF adminmenulevel >= 10 THEN 
													%> 
			   									<input type=radio NAME="sReleaseType" VALUE="Electronic" <% IF sReleaseType="Electronic" THEN response.write "checked"%> >
			   									<FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR=<% =textcolor1 %> ><b>Electronic</b></font>
			   									<%
											END IF 
											%>
											<input type=radio NAME="sReleaseType" VALUE="Paper" <% IF sReleaseType="Paper" THEN response.write "checked"%>>
											<FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR=<% =textcolor1 %> ><b>Paper W/Signature</b></font>
											<input type=radio NAME="sReleaseType" VALUE="None" <% IF sReleaseType="None" THEN response.write "checked"%>>
											<FONT size=<% =fontsize2 %> face=<% =font1 %> COlOR=<% =textcolor1 %> ><b>No Waiver</b></font>
										</td> 
  		  					</tr><%
							ELSE  
									%>
									<INPUT type="hidden" NAME="sReleaseType" VALUE="<%=sReleaseType%>" >
									<%
							END IF 	
							%>
		   				<tr>
		   					<td align="center">
									<br>
									<%
									IF TestValidAdminCode<>true THEN 
											%>	
											<font color="<% =textcolor1 %>" face=<% =font1 %> size="2"><b>Accept:</b></font><input type="radio" name="fRelease" <%IF sRelease="Accept" THEN Response.write("checked")%> value="Accept">
											<font color="<% =textcolor1 %>" face=<% =font1 %> size="2"><b>Decline:</b></font><input type="radio" name="fRelease" <%IF sRelease="Decline" THEN Response.write("checked")%> value="Decline">
											&nbsp;&nbsp;&nbsp;&nbsp 
											<%
									ELSE 
											%>
											<input type="hidden" name="fRelease" value="Accept">
											<%
									END IF 
									%>
									<input type="submit" value="Submit" style="width:9em">
									&nbsp;&nbsp;&nbsp;&nbsp
									<font color="<% =textcolor1 %>" face=<% =font1 %> size="2"><b>Date: <% =DATE %></b></font>
									<br><br>
		  					</td>
		  				</tr>

							</form>
			  		</TABLE>
						<br>

			  	</TD>
 		  	</TR>
			</TABLE>
			<%


	CASE 4

			' ------------------------------------------
			' --- Notice that waiver has been signed ---
			' ------------------------------------------
	
		%>
		<br>
		<TABLE BORDER="2" BORDERCOLOR="black" ALIGN="CENTER" BGCOLOR="<% =TableColor1 %>" width=75%>
		  <tr>
		      <td BGCOLOR="red"><center><font face=<% =font1 %> color="#FFFFFF" size="4"><b>Important Notice !!</b></font></TD>
		  </TR>  
 
		  <tr>
		     <td VALIGN="top">
  			<TABLE BORDER="0" VALIGN="top" ALIGN="CENTER" CELLPADDING="3" CELLSPACING="0" BGCOLOR="<%=TableColor1%>" width=100%>
			   <tr>
			      <td colspan=2 VALIGN="top" ALIGN="center">
				<br>
				<font color="<% =TextColor1 %>" face="<% =font1 %>" size="4"><b>Release Accepted !</b></font>
				<br><br>
				<font face="<% =font1 %>" size="2"><p>You have accepted the terms and conditions of the USA Water Ski PARTICIPANT WAIVER AND RELEASE OF LIABILITY, ASSUMPTION OF RISK AND INDEMNITY.</font>
				<br>
				<font color="red" face="<% =font1 %>" size="3"><p><b>You must be 18 years old.</b></font> 			
				<br><br>
				<font face="<% =font1 %>" size="2">If you did not intend to accept this WAIVER and RELEASE, or you are less than 18 years old, do not continue. </font>
				<font face="<% =font1 %>" size="2">A digital record of your acceptance will be created.<br></font>
				<br>
			    </td>
			  </tr>
			<tr>
			   <td align="center">

				<form action="/rankings/<%=RegFileName%>?nav=5&waivernav=5" method="post">

					<INPUT type="hidden" NAME="sReleaseType" VALUE="<%=sReleaseType%>" >
					<INPUT type="hidden" NAME="sSignWaiver" VALUE="<%=sSignWaiver%>" >

					<center><input type="submit" value=" Continue "></center>
				</form>
			   </td>	
			   <td align="center">
				<form action="/rankings/<%=RegFileName%>?nav=5&waivernav=2" method="post">
				  <center><input type="submit" value="Previous"></center>
				</form>
			   </td>	
			</tr>
		 	</TABLE>   
		   </TD>		
		</TR>
	</TABLE>
	<br>   
	<% 


	CASE 10
				' --- Sends navigation to the next tab ---
				response.redirect("/rankings/"&RegFileName&"?nav=6")


	END SELECT

END SUB


' ---------------------------
  SUB SaveWaiverValuesToTable
' ---------------------------

		' --- SUB is NOT located in this RegFormDisplay.asp ---
		DefineWaiverSpecs

		' --- SUB is Located in RegFileName (updates Session variables) ---
		CheckWaiverStatus


		' ---- Read WhichTable for existing record ----
		SET rs=Server.CreateObject("ADODB.recordset")
		sSQL = "SELECT * FROM "&RegGenTableName
		sSQL = sSQL + " WHERE Left(TourID,6) = '" & SQLClean(left(sTourID,6)) & "' AND MemberID = '"&sMemberID&"'"
		rs.open sSQL, SConnectionToTRATable, 3, 3


		' -----------------------------------------------------------------------------------------
		' -----  Stores Waiver Code in RegTempTable then branches to complete the transaction -----
		' -----------------------------------------------------------------------------------------
		OpenCon
		sSQL = "UPDATE "&RegTempTableName
		sSQL = sSQL + " SET WaiverCode = '"&sWaiverCode&"', SignWaiver = '"&SQLClean(sSignWaiver)&"'"
		sSQL = sSQL + " WHERE Left(TourID,6) = '"&left(sTourID,6)&"' AND MemberID = '"&sMemberID&"'"
		con.execute(sSQL)
		closecon	

		Session("sRelease")=sWaiverCode


		' ---- Sends email to user and competition services  ----
		IF sReleaseType = "Electronic" THEN
				' --- SUB Located in Registration16.asp --- 
				SendWaiverEmail
				
				' --- If the tournament has its own waiver to sign ---
				IF TRIM(sSpecialWaiverCode)<>"" THEN
						' --- SUB Located in Registration16.asp ---
						SendSPECIALWaiverEmail sSpecialWaiverCode, sSpecialWaiverHeadline, sSpecialReleaseBannerText
				END IF
		END IF

END SUB






' ----------------------
  SUB DefineWaiverSpecs
' ----------------------


	' ---------  Displays RELEASE & WAIVER   --------	
	IF sReleaseType = "Electronic" AND Session("sMembAge") < 18 THEN
		' --- Value of minor_waiver set at the top of this page
		sWaiverCode = minor_waiver
		sWaiverSubTitle="Waiver for MINOR Participant - WaiverID: "&minor_waiver

	ELSEIF  sReleaseType = "Electronic" AND Session("sMembAge") >= 18 THEN
		' --- Value of adult_waiver set at the top of this page
		sWaiverCode = adult_waiver	
		sSignWaiver = SQLClean(LEFT(sFirstName,12))&" "&SQLClean(LEFT(sLastName,18))
		sWaiverSubTitle="Waiver for ADULT Participant - WaiverID: "&adult_waiver

	ELSEIF sReleaseType = "Paper" AND Session("sMembAge") >= 18 THEN 
		sWaiverCode = "Paper"	
		sSignWaiver = SQLClean(LEFT(sFirstName,12))&" "&SQLClean(LEFT(sLastName,18))
		sWaiverSubTitle="Waiver for ADULT Participant - WaiverID: "&adult_waiver

	ELSEIF sReleaseType = "Paper" AND Session("sMembAge") < 18 THEN 
		sWaiverCode = "Paper"	
		sSignWaiver = SQLClean(LEFT(sFirstName,12))&" "&SQLClean(LEFT(sLastName,18))
		sWaiverSubTitle=" Waiver for MINOR Participant Required "

	ELSEIF sReleaseType = "None" AND Session("sMembAge") >= 18 THEN 
		sWaiverCode = "None"	
		sSignWaiver = ""
		sWaiverSubTitle=" Waiver for ADULT Participant Required "

	ELSEIF sReleaseType = "None" AND Session("sMembAge") < 18 THEN 
		sWaiverCode = "None"	
		sSignWaiver = ""
		sWaiverSubTitle=" Waiver for MINOR Participant Required "

	ELSE		
		sWaiverCode = "None"	
		sSignWaiver = ""
		sWaiverSubTitle=" Paper Waiver for ADULT Participant Required "
	END IF

END SUB


Function CheckClass(ThisClass, WhichField)
	CheckClass=""
	IF WhichField=ThisClass THEN CheckClass="checked"
'response.write("sClass1="&sClass1)
'response.write("ThisClass="&ThisClass)
'response.write("WhichField="&WhichField)

End Function	

		

%>

