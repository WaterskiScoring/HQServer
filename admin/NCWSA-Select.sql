 Select Substring(MX.MemberID,1,3) + '-' + Substring(MX.MemberID,4,2) + '-' + 
 Substring(MX.MemberID,6,4) as MemID, MX.LastName, MX.FirstName,
 Case when MX.Sex = 'F' Then 'CW' else 'CM' end as Div,

 Case when SO.OffCode is not Null then '0FF' else MX.Sorter end as Sorter,

 MX.Team, MX.Sex, MX.Age, MX.City, MX.State,

 Case when OD.PersonID is Null then '-' else Right(OD.RtgLvl,1) end +
 Case when OJ.PersonID is Null then '-' else Right(OJ.RtgLvl,1) end +
 Case when OC.PersonID is Null then '-' else Right(OC.RtgLvl,1) end +
 Case when OS.PersonID is Null then '-' else Right(OS.RtgLvl,1) end as OffRat,

 Coalesce(SO.OffCode,'') as OffCode,

 MX.EffTo, MX.Memtype, MX.MemCode, MX.CanSki, MX.SptsDiv, MX.AnnWvr, MX.EvtWvr


'	This begins the major MX Sub-query, which pulls membership and team and entry information
		
 From (Select MT.PersonIDWithCheckDigit as MemberID, MT.PersonID,
 Left(MT.LastName,12) as LastName, Left(MT.FirstName,10) as FirstName,

 ( 2010 -Year(MT.BirthDate)-1) as Age,

 Left(MT.City,12) as City, Left(MT.State,2) as State,
 MT.EffectiveTo as EffTo, MT.MembershipTypeCode as MemType,
 Typ.TypeCode as MemCode, Typ.CanSkiInTournaments as CanSki,
 MT.DivisionCode1 + '/' + MT.DivisionCode2 as SptsDiv,
 Upper(Left(MT.Sex,1)) as Sex, MT.WaiverStatusID as AnnWvr,

 Case when TE.Team is not null then 'E' else 'Z' end + Case 
 when TR.Team is not null then TR.Team else 'zzz' end as Sorter,

 Case when TR.DateInactive is not null then 'I' else 'A' end as TeamStat,

 Coalesce(TR.Team,'   ') as Team,
 Coalesce(RP.SlalomEnt,'  ') as SlmEnt, 
 Coalesce(RP.TrickEnt,'  ') as TrkEnt, 
 Coalesce(RP.JumpEnt,'  ') as JmpEnt, 
 Coalesce(RP.WaiverStat,' ') as EvtWvr, 

/*	Begin FROM and JOIN table list for MX Sub-Query */

 FROM USAWaterski.dbo.Members as MT Inner Join
 USAWaterski.dbo.MembershipTypes as Typ
 ON MT.MembershipTypeCode = Typ.MemberShipTypeID


/*	Here's the subquery which now pulls Team ID's from the Team Roster Extract. */
/*	Identify Latest Team affiliation for Member -- new version */
 Left Join (Select RX.MemberID, RX.Team, RX.DateInactive
 from Cobra00025.USAWSRank.TeamRoster as RX
 join (select MemberID, Max(LastEvent) as MaxEvt
 from Cobra00025.USAWSRank.TeamRoster group by MemberID) as ME 
 on ME.MemberID = RX.MemberID and ME.MaxEvt = RX.LastEvent) as TR
 on TR.MemberID = MT.PersonIDWithCheckDigit

/*	This subquery identifies Teams that are Entered, used to preface Sorter extract column */
 left join (Select distinct team 
 from Cobra00025.USAWSRank.TeamRotations where
 TournAppID = '11U042') as TE
 on TE.Team = TR.Team

/* This subquery pulls Rotation Plan information for this Person/Team */
 left join Cobra00025.USAWSRank.TeamRotations as RP
 on RP.TournAppID = ' 11U042'
 and RP.MemberID = MT.PersonIDWithCheckDigit and RP.Team = TR.Team

/* Now here's the WHERE condition clause for the Primary MX Sub-Query */
 Where Typ.ExporttoTouramentRegistrationTemplate = 1
 AND MT.Deceased = 0 AND ( ( 2010 
 - Year(MT.BirthDate) - 1) between 16 and 29 OR
 MT.DivisionCode1 = 'NCW' OR MT.DivisionCode2 = 'NCW' OR

 PersonID in (Select PersonID from USAWaterski.dbo.TempApptdOfcls
 Where TournAppID = '11U042') OR

 PersonIDWithCheckDigit IN (Select Distinct MemberID from
 Cobra00025.USAWSRank.TeamRoster) ) ) as MX 

/*	End of MX Primary MX Select Subquery.  Appended Info Subqueries follow. */

 Left Join (Select OT.PersonID,
 Max(convert(char(1),LV.LevelOrderforTemplate)
 + LV.LevelAbbreviationforTemplate) AS RtgLvl
 FROM USAWaterski.dbo.Officials OT INNER JOIN
 USAWaterski.dbo.Level LV ON OT.Level_ID = LV.Level_ID
 WHERE OT.DivisionCode in ('AWS','USA')
 AND LV.LevelOrderforTemplate IS NOT NULL
 AND OT.RatingType_ID = 3 GROUP BY OT.PersonID) as OD
 on OD.PersonID = MX.PersonID

 Left Join (Select OT.PersonID,
 Max(convert(char(1),LV.LevelOrderforTemplate)
 + LV.LevelAbbreviationforTemplate) AS RtgLvl
 FROM USAWaterski.dbo.Officials OT INNER JOIN
 USAWaterski.dbo.Level LV ON OT.Level_ID = LV.Level_ID
 WHERE OT.DivisionCode in ('AWS','USA')
 AND LV.LevelOrderforTemplate IS NOT NULL
 AND OT.RatingType_ID = 1 GROUP BY OT.PersonID) as OJ
 on OJ.PersonID = MX.PersonID

 Left Join (Select OT.PersonID,
 Max(convert(char(1),LV.LevelOrderforTemplate)
 + LV.LevelAbbreviationforTemplate) AS RtgLvl
 FROM USAWaterski.dbo.Officials OT INNER JOIN
 USAWaterski.dbo.Level LV ON OT.Level_ID = LV.Level_ID
 WHERE OT.DivisionCode in ('AWS','USA')
 AND LV.LevelOrderforTemplate IS NOT NULL
 AND OT.RatingType_ID = 2 GROUP BY OT.PersonID) as OC
 on OC.PersonID = MX.PersonID

 Left Join (Select OT.PersonID,
 Max(convert(char(1),LV.LevelOrderforTemplate)
 + LV.LevelAbbreviationforTemplate) AS RtgLvl
 FROM USAWaterski.dbo.Officials OT INNER JOIN
 USAWaterski.dbo.Level LV ON OT.Level_ID = LV.Level_ID
 WHERE OT.DivisionCode in ('AWS','USA')
 AND LV.LevelOrderforTemplate IS NOT NULL
 AND OT.RatingType_ID = 9 GROUP BY OT.PersonID) as OS
 on OS.PersonID = MX.PersonID

 Left Join	(Select PersonID, OffCode from USAWaterski.dbo.TempApptdOfcls
 Where TournAppID = '11U042'
 as SO on SO.PersonID = MX.PersonID

 Order By Case when SO.OffCode is not Null then 'E OF' else MX.Sorter end,
 MX.LastName, MX.FirstName, MX.MemberID
