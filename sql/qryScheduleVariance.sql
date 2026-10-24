SELECT
	s.SiteCode,
	s.SiteName,
	m.MilestoneName,
    d.[Start Plan],
	d.[Finish Plan],
	d.[Start Real],
	d.[Finish Real],
    DateDiff("d", d.[Finish Plan], d.[Finish Real]) AS VarianceDays,
	IIf(d.[Finish Real] Is Null, "Not Completed",
		IIf(DateDiff("d", d.[Finish Plan], d.[Finish Real]) > 0, "Late", "On Time"))
	AS Status
FROM (tblDates AS d
	INNER JOIN tblSites AS s ON d.SiteID = s.SiteID)
	INNER JOIN tblMilestones AS m ON d.Milestone = m.MilestoneID
ORDER BY DateDiff("d", d.[Finish Plan], d.[Finish Real]) DESC;
