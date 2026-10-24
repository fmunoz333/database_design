SELECT
	tblSites.Province,
	Count(tblSites.SiteID) AS NumberOfSites,
	Sum(tblSites.Population) AS TotalPopulation,
	Avg(tblSites.Population) AS AvgPopulationPerSite
FROM tblSites
GROUP BY tblSites.Province
ORDER BY Count(tblSites.SiteID) DESC;
