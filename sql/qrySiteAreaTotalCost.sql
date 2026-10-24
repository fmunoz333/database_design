SELECT
	scl.SiteID,
	s.SiteCode,
	s.SiteName,
	s.Province,
	a.AreaName,
	Sum(scl.Quantity * uc.[Unit Cost USD]) AS AreaGrandTotal
FROM 
	(((tblSites AS s
	INNER JOIN tblSiteCostLine AS scl ON s.SiteID = scl.SiteID)
	INNER JOIN tblUnitCosts AS uc ON scl.UnitCostID = uc.UnitCostsID)
	INNER JOIN tblArea AS a ON uc.Area = a.AreaID)
GROUP BY 
	scl.SiteID, 
	s.SiteCode, 
	s.SiteName, 
	s.Province, 
	a.AreaName
ORDER BY 
	scl.SiteID;