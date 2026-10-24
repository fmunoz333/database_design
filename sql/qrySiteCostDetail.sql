SELECT
	scl.SiteID,
	uc.UnitCostItem,
	scl.Quantity,
	uc.[Unit Cost USD],
	scl.Quantity * uc.[Unit Cost USD] AS ExtendedCost
FROM
	tblSiteCostLine AS scl
INNER JOIN
	tblUnitCosts AS uc 
	ON scl.UnitCostsID = uc.UnitCostsID;