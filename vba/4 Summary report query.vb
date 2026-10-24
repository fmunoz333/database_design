Sub CreateSummaryReportQuery()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim sql As String
    
    Set db = CurrentDb
    
    ' Drop existing query if it exists
    On Error Resume Next
    db.QueryDefs.Delete "qrySummaryReport"
    On Error GoTo 0
    
    ' Build SQL for summary report
    sql = "SELECT " & _
          "s.SiteID, s.SiteCode, s.SiteName, s.Province, s.Municipality, s.Parrish, " & _
          "d.Milestone, d.[Start Plan], d.[Finish Plan], d.[Start Real], d.[Finish Real], " & _
          "u.UnitCostItem, u.[Unit Cost (Euro)], u.[Unit Cost USD], " & _
          "p.municipal, p.construction, p.environmental, p.cultural " & _
          "FROM ((tblSites AS s " & _
          "LEFT JOIN tblDates AS d ON s.SiteCode = d.Code) " & _
          "LEFT JOIN tblUnitCosts AS u ON u.Area = s.AccessTypeID) " & _
          "LEFT JOIN tblPermits AS p ON s.SiteID = p.SiteID;"
    
    ' Create query
    Set qdf = db.CreateQueryDef("qrySummaryReport", sql)
    
    MsgBox "Summary report query 'qrySummaryReport' created successfully!"
End Sub
