SELECT
	scl.SiteID,
	s.SiteCode,
	s.SiteName,
	s.Province,
	Sum(scl.Quantity*uc.[Unit Cost USD]) AS SiteTotalCost
FROM tblSites AS s
	INNER JOIN (tblUnitCosts AS uc
		INNER JOIN tblSiteCostLine AS scl
		ON uc.UnitCostsID = scl.UnitCostID)
	ON s.SiteID = scl.SiteID
GROUP BY
	scl.SiteID,
	s.SiteCode,
	s.SiteName,
	s.Province
ORDER BY Sum(scl.Quantity * uc.[Unit Cost USD]) DESC;
