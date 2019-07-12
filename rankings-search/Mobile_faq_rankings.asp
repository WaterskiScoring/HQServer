<!--#include file="settingsHQ.asp"-->
<!--#include virtual="/rankings/tools_include.asp"-->
<!--#include virtual="/rankings/tools_Definitions.asp"-->
<!--#include virtual="/rankings/tools_mobile_version.asp"-->
<%

ThisFileName = "Mobile_faq_rankings.asp"


OpenState="faq_rankings"
DisplayHeadOpenBodyAndBannerTags OpenState
' class="scroll" 
' height:500px; 
tyt=1
IF tyt=1 THEN
%>
<div id="faqscroll" style="background-color:white; font-family: Arial, Helvetica, sans-serif; font-size:12pt; width:100%;">
	<%
	WriteFAQData
	%>
</div><%
END IF

'WriteFAQData


' --- Writes the Closing tags for HTML ---
DisplayCloseBodyAndHTMLTags










' -------------------
  SUB WriteFAQData 
' -------------------

' javascript:scrollTo('LevsPcts')
%>
<div style="color:red; font-size:16pt; text-align:center;">Rankings FAQ
	<a name="TopTop"></a>
</div>

<div style="color:#000000; font-size:10pt; text-align:left;">
<ul><b>
<li><a href="#WhatsNew"><font color="red"> What's new for 2010 -- Elite Qualification</font></a></li>
<br><li><a href="#LevsPcts"> Explanation of Levels and Percentiles</a></li>
<br><li><a href="#InfoExpl"> Understanding the Information Presented</a></li>
<br><li><a href="#SelFiltr"> Selections &amp; Filters -- What Do I Want to See</a></li>
<br><li><a href="#RankCalc"> How the Ranking Scores are Calculated</a></li>
<br><li><a href="#WhaDaHek"> What are those *'s and #'s all about?</a></li>
<br><li><a href="#MemNoIss"> Memberships and Member Number Issues</a></li>
<br><li><a href="#NatQual"> AWSA Nationals Qualifying &amp; Rankings</a></li>
<br><li><a href="#RegQual"> AWSA Regionals Qualifying &amp; Rankings</a></li>
<br><li><a href="#EliteQual"> Elite Qualification -- Open &amp; Masters Divs</a></li>
<br><li><a href="#MiscInfo"> Other Miscellaneous Rankings Topics</a></li>
</b></ul>
</div>

<div style="color:red; font-size:12pt; margin:25px 10px 0px 10px; text-align:center;"><a name="WhatsNew"></a>What's new -- Elite Qualification</div>
<div style="color:#000000; font-size:10pt; text-align:center;"><a href="#TopTop">(back to top)</a></div>

<div style="color=#000000; font-size:10pt; margin:10px 10px 0px 10px; text-align:left;">
<p>New for 2010 is conversion from the former fixed-score-target-based 
Elite Qualification methodology, to a new Rankings-based framework.&nbsp; 
This completes the transition of AWSA's qualifications mechanisms to the 
rankings platform.</p>

<p>What does this mean?&nbsp; Succinctly put, arrival of a skier into 
Level 9 of the rankings in an event, now establishes Elite qualification 
for that skier in that event.&nbsp; Rankings-based qualifications are
good for 12 months.&nbsp; Rankings displays now include an Elite Status
column.&nbsp; The content of this column will tell you which skiers 
are Elite qualified, where their qualification comes from, and how long 
the qualification is good for.</p>

<p>During this 2010 transition year, previous Open and Masters ratings
-- achieved under the former fixed-score qualification rules -- are also 
indicated by those Elite Status codes, until they eventually expire.&nbsp; 
New Elite qualifications, now being achieved under the newly-revised AWSA 
Rule 3.03, are being folded in each day.&nbsp; So just like the other 
levels in the Rankings, the Elite qualification thresholds will now 
automatically shift from day to day to keep pace with the capabilities 
of the skier population in each event.</p>

<p>Details on how this all works appears further down on this page.&nbsp; 
<a href="#EliteQual">Click Here</a> to jump to that section.</p>
</div>



<div style="color:red; font-size:12pt; margin:25px 10px 0px 10px; padding:10px 0px 0px 0px;"><a name="LevsPcts"></a><br>Explanation of Levels and Percentiles</div>
<div style="color:#000000; font-size:10pt; text-align:center;"><a name="#TopTop" onclick="javascript:scrollTo('TopTop')">(back to top)</a></div>

<div style="color=#000000; font-size:10pt; margin:10px 10px 0px 10px;"">
<p>Skiers in the rankings for any given Division and Event have been
classified into <b>Levels</b>.&nbsp; These levels are defined in terms 
of each skiers' <b>Percentile</b> within that division and event, working
from the bottom upwards.&nbsp; So the Skier with the highest ranking score
within a Division and Event will be at the 100-th percentile.</p>

<p>The classification specifics -- the percentile break points that define
the levels -- for each Division and Event are established annually by the 
AWSA Skier's Qualification Committee.&nbsp; Those specifications can be 
found in a table that will be displayed at the bottom of any such rankings 
display.</p>

<p>The separations between levels in the rankings display are indicated
by a color-coded heading line.&nbsp; For each skier shown in the display,
The 'RANK' column value on each skier row is color coded to indicate the 
corresponding level, according to the following key:</p>
</div>


<div style="text-align:left; width;100%">
  <span style=background-color:#FFCCCC; width:15%;">&nbsp;Level 9&nbsp;</span>
  <span style=background-color:#CCFFCC; width:15%;">&nbsp;Level 8&nbsp;</span>
  <span style=background-color:#FFFF66; width:15%;">&nbsp;Level 7&nbsp;</span>
  <span style=background-color:#CCCCFF; width:15%;">&nbsp;Level 6&nbsp;</span>
  <span style=background-color:#CC99FF; width:15%;">&nbsp;Level 5&nbsp;</span>
  <span style=background-color:#F5DEB3; width:15%;">&nbsp;Level 4&nbsp;</span>
  <span style=background-color:#FFFACD; width:15%;">&nbsp;Level 3&nbsp;</span>
</div>

<br>&nbsp;<br>

<div style="color=#000000; font-size:10pt; margin:10px 10px 0px 10px;">
<p><font color="blue"><b>Foreign Skiers:</b></font>&nbsp; It is 
important to note that only the distribution of scores of USA skiers 
are used to determine where those level breaks will fall for each 
Division and Event.&nbsp; The scores of Non-USA skiers are ignored in 
developing those level break points.&nbsp; So if you choose an "All 
Federations" display, you will find that the Non-USA skiers in that 
display will not show either a percentile nor a level code -- but they 
will of course appear at their proper place in the overall list of 
scores for that Division and Event.</p>
</div>




<div style="color:red; font-size:12pt; margin:25px 10px 0px 10px;"><a name="InfoExpl"></a>Understanding Information Presented</div>
<div style="color:#000000; font-size:10pt; text-align:center;"><a href="#TopTop">(back to top)</a></div>

<div style="color=#000000; font-size:10pt; margin:10px 10px 0px 10px;">

<p>Each row in the rankings table presents a number of attributes for 
that named skier.&nbsp; Column-by-column, here is a brief explanation on 
each of those elements:</p>

<ul>

<li><font color="blue"><b>Rank:</b></font>&nbsp; Displays each skier's
position in the list, with #1 at the top and working downwards.&nbsp;
If two or more skiers have identical ranking scores, they will be shown
as tied, with the same Rank number and the letter <b>T</b> indicating
that tie on both of their rankings.&nbsp; The color background in this 
column indicates the Qualification Level for this skier.&nbsp; 
<a href="#LevsPcts">Click Here</a> to jump to more detailed information
on the subject of Levels and Percentiles.
<br>&nbsp;</li>

<li><font color="blue"><b>Member Name:</b></font>&nbsp; Displays the
member's name, exactly as it appears in the USA Waterski membership
database.
<br>&nbsp;</li>

<li><font color="blue"><b>Score:</b></font>&nbsp; Displays the calculated
Ranking Score value for the member in this Division and Event.&nbsp; 
Details on how these are calculated is presented further down on this 
page.&nbsp; <a href="#RankCalc">Click Here</a> to jump directly to that 
explanatory material.&nbsp; If you hold your mouse pointer over the Score
value for any particular skier, a box will pop up that details the specific
performances which contribute to that skier's ranking score.&nbsp; If you 
are interested in seeing more detail on <b><i>all</i></b> the scores for 
a specific skier, clicking on the skier's name will switch you to a 
score detail display for that particular skier, for the division and 
event currently selected and displayed.
<br>&nbsp;</li>

<li><font color="blue"><b>Elite Status:</b></font>&nbsp; An Elite
Division code will appear in this column for any skier who is currently
Elite qualified in the event.&nbsp; If you hold your mouse pointer over 
the code(s) for any particular skier, a box will pop up listing which 
division's ranking that qualification stems from, and how long the 
qualification is good for.&nbsp; <a href="#EliteQual">Click Here</a> 
to jump to more detailed material on how Elite Qualifications are
determined.
<br>&nbsp;</li>

<li><font color="blue"><b>Home State:</b></font>&nbsp; Displays the
State code for the member's address, as it appears in the USA Waterski
membership database.
<br>&nbsp;</li>

<li><font color="blue"><b>Home Region:</b></font>&nbsp; This is derived
from the above-mentioned State code, based on the mapping of the 50 
states into the 5 AWSA Regions.
<br>&nbsp;</li>

<li><font color="blue"><b>Regional Place:</b></font>&nbsp; Displays 
the placement earned by this skier in their Regional Tournament for 
the selected time period.&nbsp; If the Regionals in which the skier 
competed was somewhere <i>other</i> than the region to which their
state belongs, then the code for the Regionals actually skied in will 
be shown adjacent to the placement value.
<br>&nbsp;</li>

<li><font color="blue"><b>National Place:</b></font>&nbsp;  Displays 
the placement earned by this skier in the National Tournament for 
the selected time period.
<br>&nbsp;</li>

<li><font color="blue"><b>Federation:</b></font>&nbsp; Displays the 
Federation code entered by the skier into the Membership system.
<br>&nbsp;</li>

<li><font color="blue"><b>Percentile:</b></font>&nbsp; Displays the
percentile level at which this particular skier falls within the overall
distribution of USA skiers in the Division and Event.&nbsp; 
<a href="#LevsPcts">Click Here</a> to jump to more detailed information
on the subject of Levels and Percentiles.
<br>&nbsp;</li>

</ul>
</div>





<div style="color:red; font-size:12pt; margin:25px 10px 0px 10px; text-align:center;"><a name="SelFiltr"></a>Rankings Selection - Filters</div>
<div style="color:#000000; font-size:10pt; text-align:center;"><a href="#TopTop">(back to top)</a></div>

<div style="color=#000000; font-size:10pt; margin:10px 10px 0px 10px; text-align:left;">

<p><font color="blue"><b>National / Junior / Collegiate / 
Grassroots:</b></font>:&nbsp; 
This major choice is controlled by the radio buttons presented across the 
top of the controls box that sits above the rankings display area.

<p><font color="blue"><b>Range (time period):</b></font>&nbsp; is specified 
by the first drop-down at the top of the controls box.&nbsp; Please note 
that the Range chosen is a <b><i>global</i></b> choice, and so a change 
here will also affect which scores are displayed in any Score detail 
displays you may subsequently choose.&nbsp; The choices offered will vary 
depending on the chosen Sports Division.&nbsp; For the National and Junior 
choices, there are two major types of timeframes that you may choose:</p>

<ul>

<li><font color="blue"><b>Qualifications Ranking List -- Rolling 12 
Months:</b></font>:&nbsp; This is the default choice, and the resultant 
ranking lists are the ones used in determining qualification for AWSA 
Regional and National tournaments, and for Elite status.&nbsp; This 
timeframe represents exactly 12 months on any given day, counted 
backwards from the current date.&nbsp; Scores older than 12 months 
ago are NOT included in the Last 12 Months ranking calculations.
<br>&nbsp;</li> 

<li><font color="blue"><b>Ski Year Ranking Lists:</b></font>&nbsp; Rankings 
lists for the Range of any <b>Ski Year</b> may be accessed by changing 
the Range option from Last 12 Months to a particular Ski Year.&nbsp; 
Note the Ski Year period begins on the day after the conclusion of an AWSA 
Nationals and extends through the closing date of the next Nationals.&nbsp;
The Nationals is the last tournament of each ski year, and the scores 
therefrom are credited to that ski year.
</li>

</ul>

<p><font color="blue"><b>Division &amp; Event:</b></font>&nbsp; These two
elements are the primary selections, and are controlled by the drop-down
boxes near the top of the controls box.&nbsp; The specific choices available
in those drop-downs, will be a function of the Sports Division chosen.</p> 

<p><font color="blue"><b>Geographic Filters:</b></font>&nbsp; The defaults 
for Region and State selection are both "All", and so the initial display 
you see will be for the entire country.&nbsp; You may change the Region 
and/or State selections to display rankings for geographic subsets.&nbsp;
Such Region and State selections are based on the State code stored in the 
USA Waterski Membership database.&nbsp; Updating that code (including 
the National Federation code) can only be done by USA Water Ski Membership 
services staff.&nbsp; Consequently an address change may also change the 
region in which a member appears in filtered rankings, and whether a member 
is considered a USA or Non-USA skier.</p>
</div>



<div style="color:red; font-size:12pt; margin:25px 10px 0px 10px; text-align:center;"><a name="RankCalc"></a>How Rankings are Calculated</div>
<div style="color:#000000; font-size:10pt; text-align:center;"><a href="#TopTop">(back to top)</a></div>

<div style="color=#000000; font-size:10pt; margin:10px 10px 0px 10px; text-align:left;">

<p><font color="blue"><b>Official Ranking Formula:</b></font>&nbsp; 
AWSA Rule 1.13, newly revised in 2008, specifies the official Rankings formula
that is used for the AWSA National and Junior Ranking Lists:</p>

<p><b><i>"Skiers will be ranked based on the average of their top three 
tournament scores in slalom, tricks, jumping and overall.&nbsp; For 
tournaments having multiple rounds only the best single round score will 
be taken for that tournament.&nbsp; A skier having fewer than three 
tournament scores will have a penalty applied to the average of the 
scores they do have, in the respective classes reported, based on the 
table below:</i></b></p>

<table cellPadding=0 Border=0>
<tr>
	<th Align="left"><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></td>
	<th Align="left"><font size="2"><i>Scores and Classes Situation</i></font></th>
	<th Align="left"><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></td>
	<th Align="center"><font size="2"><i>Penalty</i></font></th>
</tr>
<tr>
	<td Align="left"><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></td>
	<td Align="left"><font size="2"><i>One score class <b>C</b></i></font></td>
	<td Align="left"><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></td>
	<td Align="Center"><font size="2"><i>10%</i></font></td>
</tr>
<tr>
	<td Align="left"><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></td>
	<td Align="left"><font size="2"><i>One score class <b>E</b> or above</i></font></td>
	<td Align="left"><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></td>
	<td Align="Center"><font size="2"><i>5%</i></font></td>
</tr>
<tr>
	<td Align="left"><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></td>
	<td Align="left"><font size="2"><i>Two scores both class <b>C</b></i></font></td>
	<td Align="left"><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></td>
	<td Align="Center"><font size="2"><i>5%</i></font></td>
</tr>
<tr>
	<td Align="left"><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></td>
	<td Align="left"><font size="2"><i>One score class <b>E</b> or above and one 
		score class <b>C</b></i></font></td>
	<td Align="left"><font size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font></td>
	<td Align="Center"><font size="2"><i>2.5%</i></font></td>
</tr>
</TABLE>

<b><i>
<p>If a skier has two or more tournament scores, then the official Ranking
Score will be the highest of the following possibilities:</p>

<ol>
	<li>The best single score with penalty for that one score according to the 
		chart above.</li>
	<li>The average of the best two scores with penalty for those two scores 
		according to the chart above.</li>
	<li>The average of the best three scores with no penalty."</li>
</ol>
</i></b>

<p>In addition to the above provisions in Rule 1.13, a new provision has been
added to Rule 3.06 which also affects the ranking list score calculations:</p>

<p><b><i>"Scores below Class C may be included in the Ranking calculations, 
but may not exceed the applicable previous ski year’s Level 5 Cut Off Average."
</i></b></p>

<p><font color="blue"><b>Ranking Calculation Examples:</b></font></p>
	
<ol>

<li>You slalomed in 2 Class C tournaments with scores of 50 
&amp; 46 buoys.&nbsp; Your average of those two would be 50 + 46 = 
96 / 2 = 48.&nbsp; Less 5% penalty (2.4) yields an official ranking 
score of 45.6 buoys.<br>&nbsp;</li>

<li>You jumped in 2 Class C tournaments with scores of 60 
&amp; 50 feet.&nbsp; Your average of those two would be 60 + 50 = 
110 / 2 = 55.&nbsp; Less 5% penalty (2.75) yields a ranking 
score of 52.25 feet.&nbsp;However, applying a 10% penalty to the 
highest single score of 60 feet would yield a result of 54 feet,
and since that is higher than 52.25, then the 54 becomes the 
official ranking score value.</li>

</ol>


<p><font color="blue"><b>Equivalent Scores:</b></font>&nbsp; Each skier
will appear in the particular ranking list that corresponds to their
official age division, for the selected timeframe.&nbsp; The AWSA optional
Divisions -- Open Men and Women and Masters Men -- represent additional
separate ranking lists.&nbsp; A skier in the age range of 18-24 with scores 
recorded in <b>M1</b> will appear in the <b>M1</b> ranking list.&nbsp; If 
such a skier has scores reported in <b>OM</b>, then they will appear in
the OM ranking list.&nbsp; If they have scores in both <b>M1</b> 
<b><i>and</i></b> in <b>OM</b>, then they will appear in both lists.</p>

<p>Ranking Score values for any particular Division and Event are calculated 
following an Equivalent Scores methodology.&nbsp; This combines scores 
across divisions that a member has reported performances in, provided the 
performance conditions of Boat Speed (S/J) and Ramp Height (J) are 
comparable.&nbsp; Tricks is always equivalent.</p>

<p><font color="blue"><b>Graduating Skiers:</b></font>&nbsp; For the Rolling 
12 Month timeframe, those skiers who have changed Age Divisions during that 
period will be shown in their new Age Division.&nbsp; In most instances, 
performances recorded in the previous age division will be equivalent and 
will contribute to the skier's rankings in their new age division.&nbsp;
Where the maximum slalom boat speed in the new Age Division is lower than
the old one, those scores will be adjusted in accordance with AWSA rule
10.06(c) to be compatible with the new Age Division.</p>

<p><font color="blue"><b>Posting and Updating Scores:</b></font>&nbsp; 
Scores are uploaded to the Rankings website, immediately upon receipt from 
the tournament scorer.&nbsp; Any newly reported or updated scores will show
up immediately in Score Detail displays.&nbsp; However, the Ranking Scores
are recalculated only once each day, during the overnight hours.&nbsp; As 
a result, any changes or additions to the score database on any given day, 
will not be reflected in the rankings until the following morning.</p>

</div>




<div style="color:red; font-size:12pt; margin:25px 10px 0px 10px; text-align:center;"><a name="WhaDaHek"></a>Meaning of *'s and #'s</div>
<div style="color:#000000; font-size:10pt; text-align:center;"><a href="#TopTop">(back to top)</a></div>

<div style="color=#000000; font-size:10pt; margin:10px 10px 0px 10px; text-align:left;">

<p>Ranking scores that are accompanied with special symbols, signify
rankings that are based on fewer than three tournaments:</p>

<ul>

<li>a <font color="red" size="4"><b>*</b></font> next to a ranking score, 
signifies that skier has less than the requisite three tournaments, 
and that they have been assessed a penalty on the average of what they 
<b><i>do</i></b> have, according to the Rule 1.13 table (which appears
in the Ranking Calculations section above).
<br>&nbsp;</li>

<li>a <font color="red" size="3"><b>#</b></font> next to a ranking score, 
on the other hand, means the calculation has chosen to use fewer scores
than the skier <b><i>does</i></b> have, because bringing in the next 
lower score would actually serve to reduce their ranking score, compared 
to using fewer scores with the applicable penalty.&nbsp; This is commonly 
referred to as the "Do No Harm" provision, spelled out in Rule 1.13 --
explained in more detail in the preceding section.</li>

</ul>
</div>



<div style="color:red; font-size:12pt; margin:25px 10px 0px 10px; text-align:center;"><a name="MemNoIss"></a>Memberships and MemberID Issues</div>
<div style="color:#000000; font-size:10pt; text-align:center;"><a href="#TopTop">(back to top)</a></div>

<div style="color=#000000; font-size:10pt; margin:10px 10px 0px 10px; text-align:left;">

<p><font color="blue"><b>Birthdate / Gender information:</b></font>&nbsp; 
It is important to recognize that the classification of scores and the
calculation of rankings are dependent on a number of factors that come from
the USA Water Ski Membership database.&nbsp; And these factors are connected
to the Scores database, using Membership Numbers as the primary key.</p>

<p>Therefore, if the Birthdate or Gender information in the USA Water Ski
membership database is not correct for any particular member, then their 
derived rankings may not be correct either.&nbsp; If you find your scores 
or rankings showing up where you believe is the wrong place, then please
start by contacting the Membership department at USA Water Ski and validate
the birthdate they have on file for you.&nbsp; Any corrections they make in 
the Membership Database will be reflected in the rankings the next day.</p>

<p><font color="blue"><b>Renewing Memberships versus Enrolling again:</b></font>&nbsp;
We have recently found that some members appear to be Enrolling anew, instead
of renewing their existing membership.&nbsp; And this results in their being
issued a new membership number.&nbsp; As a consequence, then any subsequent
scores reported under that new membership number will create a new ranking list
entry.&nbsp; As a result, a member may show up twice in the rankings.&nbsp;
Please recognize that such seeming duplications in the Rankings are really
the result of such schitzoid memberships, rather than scorekeeping errors.&nbsp;
If you see something like this for yourself, then please contact the USA 
Water Ski Membership department with the particulars, including both the 
new and the old member numbers, and they should be able to consolidate your
memberships and member history.&nbsp; And when they do so, then the two sets
of scores will also be brought together under the surviving member number, in 
the next daily Rankings update.</p>

</div>


<div style="color:red; font-size:12pt; margin:25px 10px 0px 10px; text-align:center;"><a name="NatQual"></a>Nationals Qualifying and Rankings</div>
<div style="color:#000000; font-size:10pt; text-align:center;"><a href="#TopTop">(back to top)</a></div>

<div style="color=#000000; font-size:10pt; margin:10px 10px 0px 10px; text-align:left;">


<p>Qualification to compete in the AWSA Nationals is based primarily on
top placements in the current year's Regionals, plus top placements in
the previous year's Nationals.&nbsp; Additional skiers may qualify based 
on Ranking List Levels. Skiers ranked Level 8 or better in any
event or overall on the Thursday three weeks prior to the Thursday 
immediately before the start of the National Tournament (the official
Cut Off Date) are automatically qualified.&nbsp; At the Cut Off Date,
the Rankings List Score from the lowest skier in Level 8 in each event
and overall will be locked in for each division.&nbsp; Those particular
scores on that date become the Cut Off Averages (COA).</p>  

<p>After the Cut Off Date, skiers not already qualified as of that date, 
may qualify for Nationals by any one of following three <b>Last Chance 
Qualifying</b> methods (LCQs):

<ol>

<li>Increase your ranking value to the COA or higher, at any time during 
the period beginning with the Cut Off Date through the close of Nationals 
registration for that event;&nbsp; or
<br>&nbsp;</li>

<li>Record a score equal to or above the Level 8 COA for respective event and division at Regionals, or any class C or above tournament between the Cut Off Date and the first day of Nationals;&nbsp; or
<br>&nbsp;</li>

<li>Place in the top five (5) at Regionals, regardless of ranking. 

</ol>

<p>Skiers placing in the top five (5) at Nationals automatically qualify 
for the next year's Nationals.&nbsp; Any skier who is Elite Qualified on
the Rankings is automatically qualified.</p>

<p>More details on the subject of Nationals Qualifications can be found
in section 4.02 of the AWSA Rulebook.&nbsp; A link to that rulebook can
be found in the Left Navigation section on the Water Skiing section on 
this USA Water Ski Web site.</p> 

</div>


<div style="color:red; font-size:12pt; margin:25px 10px 0px 10px; text-align:center;"><a name="RegQual"></a>Regionals Qualifying and Rankings</div>
<div style="color:#000000; font-size:10pt; text-align:center;"><a href="#TopTop">(back to top)</a></div>

<div style="color=#000000; font-size:10pt; margin:10px 10px 0px 10px; text-align:left;">

<p>Qualification for most of the Regional tournaments are also being done
in terms of the Ranking List and Levels.&nbsp; However, the specific levels 
and dates do vary from one Region to another.&nbsp; Consequently, you 
should consult the applicable Region's website to get the pertinent rules 
for that particular Region.&nbsp; You can find links to those Regional 
websites, under the <b>Region Websites</b> heading, which can be found 
in the Left Navigation area on the <b>Water Skiing</b> section on this USA 
Water Ski Web site.
</div>


<div style="color:red; font-size:12pt; margin:25px 10px 0px 10px; text-align:center;"><a name="EliteQual"></a>Elite Qualification -- Open &amp; Masters</div>
<div style="color:#000000; font-size:10pt; text-align:center;"><a href="#TopTop">(back to top)</a></div>

<div style="color=#000000; font-size:10pt; margin:10px 10px 0px 10px; text-align:left;">

<p>Qualification for the Elite Divisions -- Open and Masters -- is now based
on Rankings rather than on fixed score targets, as it used to be done in past
seasons.&nbsp; This section presents information on the following key features 
of this new protocol:</p>

<ul><b>
<li><a href="#EQBkgrnd"> Overview of Rankings-based Elite:&nbsp; Level 9</a></li>
<li><a href="#EQByEvent"> Level 9 Cutoff Scores by Event, S / T / J</a></li>
<li><a href="#EQOverall"> Level 9 for Overall -- special considerations</a></li>
<li><a href="#EQDisplay"> Elite Qualification Info in Rankings Displays</a></li>
<li><a href="#EQOpen35"> Open Qualification for Men Age 35 and up</a></li>
<li><a href="#EQOldFix"> Transition of Existing fixed-score Ratings</a></li>
</b></ul>

<p>AWSA Rule 3.03, as revised January 2010, defines Elite Division
competition.&nbsp; While we will be discussing certain features of this
new protocol here, we strongly encourage anyone interested in the finer
points to consult the official AWSA rulebook for those details.&nbsp;
Meanwhile, here are the high points.</p>

<p>Rule 3.03 defines two Elite Divisions of competition.&nbsp; The 
<b>Open</b> Division may be entered by any skier of <b>any age</b> who 
is Open Qualified.&nbsp; The <b>Masters</b> Division may be entered by 
any skier <b>age 35 or greater</b> who is Masters Qualified.&nbsp; Entry 
into either Elite Division is optional.&nbsp; Masters or Open qualification 
for an event is attained by a skier reaching Level 9 in the rankings for 
that event, in the applicable skier pool.</p>

<p>The Elite Level 9 is fundamentally different from the other lower 
Levels in the rankings.&nbsp; All Ranking Levels up to Level 8 in any 
Division/Event are based on skiers having scores <b><i>solely in that 
particular Division</i></b> for that Event.&nbsp; In contrast, Level 9 
is determined <b><i>across Consolidated Pools</i></b> of skiers.&nbsp; 
The Open and Masters pools each consist of an age range that encompasses 
more than one division.&nbsp; We deliberately use the word <b>Pool</b> 
to refer to each of these age bands, rather than division, to avoid
possible confusion.</p>

<p>The age band for the pool used to determine the cutoffs for Open is 
17-34, and for Men includes skiers with scores in M1 and M2 and OM.&nbsp; 
Likewise for Women, where the pool is those with scores in W1 and W2 and 
OW.&nbsp; The age band for the pool used to determine the cutoffs for 
Masters Men is 35-52, and includes skiers with scores in M3 and M4 and 
MM.&nbsp; Those specifications appear in sub-paragraphs 3&amp;4 of rule 
3.03(c).</p>
</div>


<div style="color:red; font-size:12pt; margin:25px 10px 0px 10px; text-align:center;"><a name="EQByEvent"></a>Level 9 Cutoff Scores by Event, S/T/J</div>
<div style="color:#000000; font-size:10pt; text-align:center;"><a href="#TopTop">(back to top)</a></div>

<div style="color=#000000; font-size:10pt; margin:10px 10px 0px 10px; text-align:left;">

<p>The Level 8 (and lower) Slalom cutoff scores in M3 and M4 will always 
be completely independent of one another -- The M3 cutoffs are derived 
from the ranking of just the M3 slalom skiers, while the M4 cutoffs are 
derived from the ranking of just the M4 slalom skiers.&nbsp; In contrast, 
the Level 9 slalom cutoff scores for M3 and M4 will always be <b><i>exactly 
the same</i></b> -- that's the MM cutoff.&nbsp; That Level 9 cutoff for 
Masters Men is derived from the ranking of <b><i>the consolidation</i></b> 
of all male slalom skiers with scores in M3 and M4 and MM.&nbsp; The same 
pattern exists for the M1 and M2 and OM cutoffs for Level 8 (and lower), 
which are independent by division, versus the cutoffs for Level 9, which 
you will find are the same.&nbsp; And likewise for the W1 and W2 and OW 
divisions, as well.<p>

<p>Cutoff Scores in Tricks have exactly the same relationship as noted 
above for Slalom, with the Level 8 (and lower) cutoff scores being independent 
for each associated division, whereas the Level 9 cutoff scores are the same.</p> 
	
<p>For the Jumping event, however, the Level 9 cutoff scores across the set
of associated divisions will <b><i>not</i></b> be exactly the same, because 
there are speed/ramp condition differences between those divisions.&nbsp; 
The Level 9 cutoff for Masters Men Jumping is derived from the consolidation
of scores of M3 and M4 and MM jumpers, adjusted for condition differences
as specified in 3.03(c)6.&nbsp; The AWSA Skier's Qualification Committee has 
defined those adjustments to be 12 feet of distance for each half foot of ramp
height difference, and 8 feet of distance for each 3 kph difference in the
maximum boat speed allowed.</p>

<p>An example will help illustrate this.&nbsp; Let's say that the Level 9 
cutoff for Masters Men Jumping comes out to be 166.2 feet.&nbsp; That 
cutoff applies to jumpers with scores actually performed in MM, at 57 kph 
and on the five and a half foot ramp.&nbsp; But that needs to be adjusted 
before we can apply it to the rankings of the M3 and M4 jumpers, since they
jump at different ramp heights and/or boat speeds.&nbsp; So for M3, the 
Level 9 cutoff score will be adjusted to 146.2 feet, and the M4 cutoff to 
138.2 feet, and so on.&nbsp; This treatment is consistent with the way 
these differences in jumping conditions have always been dealt with in the 
past, as defined by the AWSA Skier's Qualification Committee.</p>

<p>If you weren't already aware of this, please note that a box appears at 
the bottom of each rankings page, which displays the actual percentiles and 
cutoff scores in effect for that day, for that Division/Event.&nbsp; The 
Level 9 cutoff score that is applicable for that division, now appears at 
the top of that box.&nbsp; So the Level 8 and lower cutoffs you see in that
box for any particular division's ranking display are derived from the
ranking of just that division, whereas the Level 9 cutoff at the top of the
box is derived from the ranking of the associated pool, instead.</p>
</div>


<div style="color:red; font-size:12pt; margin:25px 10px 0px 10px; text-align:center;"><a name="EQOverall"></a>Level 9 for Overall -- special considerations</div>
<div style="color:#000000; font-size:10pt; text-align:center;"><a href="#TopTop">(back to top)</a></div>

<div style="color=#000000; font-size:10pt; margin:10px 10px 0px 10px; text-align:left;">
<a href="#EliteQual">(back to Elite)</a>


<p>The handling of the Consolidated Pools for Elite Overall qualification,
has one additional wrinkle.&nbsp; That is, each candidate skier's Overall 
scores will have been completely recalculated, using the applicable base 
NOPS standards for the particular Elite division(s) for which that skier 
might be eligible.&nbsp; So skiers in OW or W1 or W2 will have had their 
overall scores recalculated according to OW NOPS, and skiers in OM or M1 
or M2 will have had their overall scores recalculated according to OM 
NOPS.&nbsp; And then finally skiers in M3 or older will have had their 
overall scores recalculated <b><i>twice</i></b>, once in candidacy for 
Masters using the MM NOPS, and then again for consideration for Open, 
using the OM NOPS.</p>

<p>Therefore, each candidate Elite Overall skier will have at least two
different views of each overall performance.&nbsp; What you will see when
you look at the M2 Overall rankings, are the overall rankings of those M2 
skiers, based on the M2 NOPS calculations, which of course are used to
determine all of the M2 Overall cutoffs from Level 8 on down.&nbsp; 
However, the Level 9 cutoff score that you will see displayed at the bottom 
of the M2 Overall ranking, is the OM Overall Cutoff -- which reflects the
pool of overall scores all calculated using OM NOPS.&nbsp; And so the
overall scores for each skier that are being used to determine if they 
are Elite Open overall or not, are those recalculated overall scores.</p>

<p>At this point we have not found any neat way to make those recalculated
Elite-division-based overall scores visible on the website, but please do 
recognize that those are what are being used internally, to determine each 
overall skier's Elite qualification status.&nbsp; The NOPS calculator tool, 
an Excel spreadsheet template, is available on the AWSA Website under the 
Scoring link, and can be used by anyone interesting in exploring their 
overall scores in more detail.</p>
</div>


<div style="color:red; font-size:12pt; margin:25px 10px 0px 10px; text-align:center;"><a name="EQDisplay"></a>Elite Qualification Info in Rankings Displays</div>
<div style="color:#000000; font-size:10pt; text-align:center;"><a href="#TopTop">(back to top)</a></div>

<div style="color=#000000; font-size:10pt; margin:10px 10px 0px 10px; text-align:left;">


<a name="EQDisplay"></a><br><b>Elite Qualification Info in Rankings Displays</b>
<a href="#EliteQual">(back to Elite)</a>

<p>Elite status for any particular skier in any particular event can be seen
in the rankings of any division for which that skier has recorded scores in 
that event.&nbsp; The Elite Status column for a skier will display a code 
if that skier is Elite Qualified in that event.&nbsp; Please recognize that 
a skier might be either above or below the Level 9 cutoff in that particular 
division's ranking, but if they have a code showing in that column, then 
they are Elite Qualified for the Division(s) coded.&nbsp; A skier with such 
a code that appears below Level 9 in that day's ranking, may have previously 
been in Level 9 on an earlier date within the past 12 months, or may be 
currently in Level 9 in the rankings for another division.&nbsp; Skiers 
who have scores in two or more divisions for an event, will show the 
<b><i>same</i></b> Elite Status coding in <b><i>each</i></b> of those 
division's rankings display.</p>

<p>Please note that a skier may appear in the rankings for OM or OW or MM,
<b><i>in addition to</i></b>, or instead of, the rankings for their age 
division.&nbsp; Presence in an OM or OW or MM ranking list <b><i>does 
not</i></b> signify or imply Elite qualification in that event.&nbsp; The 
presence of a skier in such a division's ranking merely indicates that the 
skier had recorded scores in that division/event.&nbsp; Conversely, the 
absence of a skier from the OM or OW or MM rankings for an event does not 
signify or imply that skier is not Elite qualified in that event, either.&nbsp; 
Indeed, they may be so qualified based on scores achieved in their age 
division, and merely have not chosen to ski in that Elite division yet.&nbsp; 
Current Elite qualification will always be unequivocally indicated by the 
codes you see in the Elite Status column -- do not be misled about such status 
based on a skier's presence on or absence from a particular Elite Division's 
ranking list.</p>

<p>Holding your mouse pointer over a code in the Elite Status column in any 
ranking display, will pop up a box showing where that Qualification stems
from, and when it will expire.</p>
</div>


<div style="color:red; font-size:12pt; margin:25px 10px 0px 10px; text-align:center;"><a name="EQOpen35"></a>Open Qualification for Men Age 35 and up</div>
<div style="color:#000000; font-size:10pt; text-align:center;"><a href="#TopTop">(back to top)</a></div>

<div style="color=#000000; font-size:10pt; margin:10px 10px 0px 10px; text-align:left;">
<a href="#EliteQual">(back to Elite)</a>

<p>Top male skiers over the age of 35 might be Masters qualified, or could 
possibly be both Masters <b><i>and</i></b> Open qualified, if their ranking 
score meets the Level 9 cutoff score for Open in their event.&nbsp; Taking 
Tricks as an example, let's say the Level 9 cutoff score for Open is 5400 
points, and the Level 9 cutoff score for Masters is 3600 points.&nbsp;
Then a M4 tricker with a ranking score of 4500 points would be Masters
qualified, and a M4 tricker with a ranking score of 6500 points would be 
<b><i>both</i></b> Masters <b><i>and</i></b> Open qualified.</p>

<p>The same applies in Slalom and Jumping as well, with a couple of small
wrinkles for each.&nbsp; In Jumping, the Open Cutoff score being used as
the criteria for a 35+ male jumper, may have been further adjusted to match 
the conditions for his age division, in accordance with the adjustment
factors explained earlier in this section.</p>

<p>In Slalom, Zero-Based Slalom scoring automatically accounts for the 
difference between the 55 kph and 58 kph maximum slalom boat speeds of 
the Open and Masters divisions.&nbsp; So let's say that on a given day
the Level 9 cutoff score for Open is 106 buoys, which for skiers in M1 
or M2 or OM (or B3) is 4 at 11.25 meters (38 off) at 58 kph.&nbsp; And 
that the Masters Level 9 cutoff score is 104 buoys, which for M3 or M4 
or MM etc is 2 at 10.75 meters (39-1/2 off) at 55 kph.&nbsp; So a M4 
slalom skier with a ranking score of 105 buoys -- 3 at 39-1/2 off at
55 kph -- would be Masters qualified.&nbsp; and another M4 slalom skier 
with a ranking score of 107 buoys -- 5 at 39-1/2 off at 55 kph -- would 
be <b><i>both</i></b> Masters <b><i>and</i></b> Open qualified.</p>

<p>Finally, a 35+ male skier could be both Masters and Open qualified in
Overall if their Overall performances -- as recalculated using the Open
division NOPS standards -- exceeds the OM Overall Level 9 cutoff score.&nbsp;
See the material on Elite Overall, which appears earlier in this section,
for more details on this subject.</p>

<p>A key point here to keep in mind, is that the Level 9 cutoff score
which displays at the bottom of any of the Male divisions for skiers age 
35 or older, will always be the <b>Masters Men cutoff</b>.&nbsp; The 
Level 9 <b>Open cutoff</b> score that is applicable for those 35+ skiers 
to attain Open Elite qualification can be found at the bottom of any of 
the M1 or M2 or OM division rankings for that event.&nbsp; Remember that
in Jumping that cutoff may be subject to a further adjustment.&nbsp; Where 
such a skier is coded as OM/MM in the Elite Status column, then recognize 
that the OM qualification has been determined by using the associated 
<b>Open cutoff</b>, rather than the Masters cutoff which appears in the 
box at the bottom of that particular age division ranking page.</p>

<p>Finally, for those skiers showing both OM and MM Elite status, that 
the expiration dates for those two qualifications may not be the same, 
depending of course on when and how those two qualifications were 
respectively earned.</p>
</div>


<div style="color:red; font-size:12pt; margin:25px 10px 0px 10px; text-align:center;"><a name="EQOldFix"></a>Transition of Existing fixed-score Ratings</div>
<div style="color:#000000; font-size:10pt; text-align:center;"><a href="#TopTop">(back to top)</a></div>

<div style="color=#000000; font-size:10pt; margin:10px 10px 0px 10px; text-align:left;">
<a href="#EliteQual">(back to Elite)</a>

<p>Open and Masters ratings that had been earned through the end of the 
2009 calendar year, under then-current rules, are good for 15 and 12 months, 
respectively.&nbsp; All of those qualifications were identified as of the
end of the 2009 year, and were pre-loaded into the new Elite Qualification 
Management tables in the rankings system.&nbsp; Any such skiers who now 
appear in Level 9 in their event(s), will of course have their Qualified 
Through information updated based on the new rules.&nbsp; Therefore, any 
Qualified Through dates you see that cite April 2011 or later, are the 
result of that skier having attained Level 9 under the new rules.&nbsp; 
Conversely, any skiers found in the rankings below level 9, but showing 
Elite Status codes, and Qualified Through dates of February 2011 and 
earlier, are the result of those previous fixed-score qualifications.</p>
</div>


<div style="color:red; font-size:12pt; margin:25px 10px 0px 10px; text-align:center;"><a name="MiscInfo"></a>Other Miscellaneous Rankings Topics</div>
<div style="color:#000000; font-size:10pt; text-align:center;"><a href="#TopTop">(back to top)</a></div>

<div style="color=#000000; font-size:10pt; margin:10px 10px 0px 10px; text-align:left;">
<a href="#TopTop">(back to top)</a>

<p><font color="blue"><b>Zero-Based Scoring in Slalom:</b></font>&nbsp; 
All divisions in Slalom are now ranked based on Zero-Based Scoring 
(ZBS).&nbsp; Slalom scoring now begins at 25 kph and the 23 meter line 
length, for all divisions.&nbsp; All reported scores and rankings values 
from previous periods have been adjusted to the new ZBS system.</p>
 
<p><font color="blue"><b>All Division Handicapping:</b></font>&nbsp; 
The ZBS system scores slalom on a continuous scale from 0 buoys to the 
best scores recorded, so special handicapping strategies across divisions 
in slalom are no longer needed.&nbsp; Even when divisions have different 
maximum boat speeds, the difficulty levels for a particular total score 
recorded in different divisions will be very similar.</p>

<p><font color="blue"><b>Ability Based Groupings:</b></font>&nbsp; 
Grouping events based on ability is encouraged.&nbsp; And of course the 
underlying division for your skier's age and gender is carried forwards,
and so those scores will contribute to rankings exactly like those from
other competitions that may be divisionally-oriented.&nbsp; Lots of folks 
are finding that creating competitions featuring ability-based groups makes
for great fun!</p>

<p><font color="blue"><b>Registration Downloads:</b></font>&nbsp; 
Downloadable Excel Registration Templates include the current Ranking 
Scores and Ranking Levels for all skiers included.&nbsp; Such a template 
can be downloaded by a sponsor or scorer and then used as the framework 
for Registration for a tournament.&nbsp; The requestor can specify criteria
to define the scope of the skiers to be included in their template.&nbsp;
USA Waterski's new online Registration system will also deliver an Excel 
template for the pre-registered skiers, and can also include other members 
according to user-specified criteria.&nbsp; These templates make it easy 
for sponsors to group skiers, or to seed their events, based on this data.</p>    

<p><font color="blue"><b>Reporting Errors or Questions:</b></font>&nbsp; 
If anything you see in these Rankings displays strikes you as being 
incorrect, or if you have questions that aren't answered by the information 
on this page, then please drop a note to
<A href="mailto:competition@usawaterski.org?subject=Ranking 
List Question">USA Water Ski Competition Department</a><p>

<p>&nbsp;</p>
</div>    
<%

END SUB    




%>


