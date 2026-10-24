SELECT
	ar.AreaName,
	Sum(scl.Quantity * uc.[Unit Cost USD]) AS TotalByArea
FROM (tblSiteCostLine AS scl
	INNER JOIN tblUnitCosts AS uc
		ON scl.UnitCostID = uc.UnitCostsID)
	INNER JOIN tblArea AS ar
		ON uc.Area = ar.AreaID
GROUP BY ar.AreaName
ORDER BY Sum(scl.Quantity * uc.[Unit Cost USD]) DESC;
