
/***	Set up and do some preliminary work on	***/
/***	an NCWSA Team Roster Table, derived	from	***/
/***	actual NCWSA scores data over 2008-2010.	***/



/***	Before beginning, here are some queries to explore possible errata	***/
/***	in the Scores table, relative to the Official TeamsList table		***/

/***	Identify usages of Invalid Team codes for A team Skiers	***/

Select	NCWSA.TourID, Convert(char(10),NCWSA.EndDate,111) as EndDate,
		ST.TName, ST.TCity, ST.TState, NCWSA.Team, NScr
FROM		(Select	TourID, Team, EndDate, count(*) as NScr
		 FROM		USAWSRank.Scores
		 Where	left(TourID,2) >= '08'
			and	left(div,1) = 'C'
			and	left(Team,1) >= 'A'
			and	TeamStat = 'A'
			and	Team Not in (select Distinct TeamID
						 from usawsrank.teamslist)
		 group by TourID, Team, EndDate)	NCWSA
JOIN		Sanctions.dbo.TSchedul			ST
	on	ST.TSanction = NCWSA.TourID
Order by NCWSA.EndDate
;

/***	Identify usages of Invalid Team codes for B team Skiers	***/

Select	NCWSA.TourID, Convert(char(10),NCWSA.EndDate,111) as EndDate,
		ST.TName, ST.TCity, ST.TState, NCWSA.Team, NScr
FROM		(Select	TourID, Team, EndDate, count(*) as NScr
		 FROM		USAWSRank.Scores
		 Where	left(TourID,2) >= '08'
			and	left(div,1) = 'C'
			and	left(Team,1) >= 'A'
			and	TeamStat = 'B'
			and	Team Not in (select Distinct TeamID
						 from usawsrank.teamslist)
		 group by TourID, Team, EndDate)	NCWSA
JOIN		Sanctions.dbo.TSchedul			ST
	on	ST.TSanction = NCWSA.TourID
Order by NCWSA.EndDate
;


/***	Now let's create our official TeamRoster table	***/

Drop table USAWSRank.TeamRoster
;

Create table USAWSRank.TeamRoster
(
    Team char(3),
    MemberID char(9),
    DateAdded DateTime,
    FirstEvent DateTime,
    LastEvent DateTime,
    NumEvents Int default 0,
    DateInactive DateTime
);
    
Delete from USAWSRank.TeamRoster
;

/***	Populate the TeamRoster table from derivatives of actual NCWSA scores	***/
/***	Validate by left(Div,1)='C' and valid team codes and SkiYears 08+		***/

Insert Into USAWSRank.TeamRoster
		(Team, MemberID, DateAdded, FirstEvent, LastEvent, NumEvents, DateInactive)
Select	Team, MemberID, Min(EndDate), Min(EndDate), Max(EndDate), count(*), NULL
FROM		(Select	Team, MemberID, TourID, EndDate
		 FROM		USAWSRank.Scores
		 Where	substring(TourID,3,1) = 'U'
			and	left(div,1) = 'C'
			and	left(TourID,2) >= '08'
			and	Team in (select Distinct TeamID 
						from usawsrank.teamslist)
		 group by Team, MemberID, TourID, EndDate)	NCWSA
group by Team, memberID
;


/***	Now Update the team roster and set DateInactive to LastEvent	***/
/***	When LastEvent Date is previous year or earlier			***/

Update USAWSRank.TeamRoster
Set DateInactive = LastEvent
where LastEvent < '2010-01-01'
;


/***	Now let's do some analysis -- first look for	***/
/***	skiers with multiple team occurrances.		***/

Select	NumTeam, Count(*)
FROM		(Select	MemberID, count(*) as NumTeam
		 FROM		USAWSRank.TeamRoster
		 group by 	MemberID)		NCWSA
Group by	NumTeam
Order by	NumTeam
;

/*** Produce listing of skiers with multiple team occurrances	***/

SELECT	TR.MemberID, MT.LastName, MT.FirstName, TR.Team, 
		Convert(char(10),TR.FirstEvent,111) as FirstEvent,
		Convert(char(10),TR.LastEvent,111) as LastEvent,
		NumEvents
FROM		USAWSRank.TeamRoster		TR
JOIN		USAWaterski.dbo.Members		MT
	on	MT.PersonIDWithCheckDigit = TR.MemberID
WHERE		TR.MemberID in
		(Select	MemberID
		 FROM		USAWSRank.TeamRoster
		 group by 	MemberID
		 having count(*) > 1)
order by MT.LastName, MT.FirstName, TR.LastEvent
;


/***	Now count up the number of Men and Women by Team,	***/
/***	and present as a report sorted by Region/Conference	***/

Select	TT.NCWRegion, TT.NCWConf, TT.TeamID, TT.TeamName, 
		TS.Men, TS.Women, TS.Men+TS.Women as Total
from		(Select	TR.Team, 
				sum(case when Upper(left(MT.Sex,1)) = 'M' then 1 else 0 end) as Men,
				sum(case when Upper(left(MT.Sex,1)) = 'F' then 1 else 0 end) as Women
		 FROM		USAWSRank.TeamRoster	TR
		 JOIN		USAWaterski.dbo.Members	MT
		 	on	MT.PersonIDWithCheckDigit = TR.MemberID
		 WHERE	TR.LastEvent > '2010-01-01'
		 group by Team)			TS
JOIN		USAWSRank.TeamsList		TT
	on	TS.Team = TT.TeamID
Order by TT.NCWRegion, TT.NCWConf, TT.TeamID
;


/***	Now let's create our official TeamRotations table	***/

Drop table USAWSRank.TeamRotations
;

Create table USAWSRank.TeamRotations
(
    Team char(3) NOT NULL,
    TournAppID char(6) NOT NULL,
    MemberID char(9) NOT NULL,
    Sex char(1) NOT NULL,
    SlalomEnt Char(2),
    TrickEnt Char(2),
    JumpEnt Char(2),
    DateUpdated DateTime,
    WaiverStat char(1)
);
    
Delete from USAWSRank.TeamRotations
;

