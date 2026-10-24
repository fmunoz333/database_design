SELECT
	s.SiteID,
	s.SiteCode,
	s.SiteName,
	s.Province,
	s.Municipality,
	at.AccessTypeName, o.OwnerCode
FROM (tblSites AS s
	INNER JOIN tblAccessType AS at
		ON s.AccessTypeID = at.AccessTypeID)
	INNER JOIN tblOwner AS o
		ON s.OwnerID = o.OwnerID
WHERE (s.Province = [Enter Province or leave blank:]
	OR [Enter Province or leave blank:] Is Null)
AND (at.AccessTypeName = [Enter Access Type or leave blank:]
	OR [Enter Access Type or leave blank:] Is Null);
