TRANSFORM Count(s.SiteID) AS SiteCount
SELECT
	s.Province
FROM tblStructureType AS st
	INNER JOIN (tblSites AS s
		INNER JOIN tblInfrastructure AS inf
		ON s.SiteID = inf.SiteID)
	ON st.StructureTypeID = inf.StructureTypeID
GROUP BY s.Province
PIVOT st.StructureTypeName;
