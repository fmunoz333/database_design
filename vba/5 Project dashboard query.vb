Sub CreateProjectDashboardQuery()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim sql As String
    
    Set db = CurrentDb
    
    ' Drop existing query if it exists
    On Error Resume Next
    db.QueryDefs.Delete "qryProjectDashboard"
    On Error GoTo 0
    
    ' Build SQL for comprehensive dashboard (6 opening parens for 7 JOINs)
    sql = "SELECT " & _
          "s.SiteID, s.SiteCode, s.SiteName, s.Province, s.Municipality, s.Parrish, " & _
          "d.Milestone, d.[Start Plan], d.[Finish Plan], d.[Start Real], d.[Finish Real], " & _
          "u.UnitCostItem, u.[Unit Cost (Euro)], u.[Unit Cost USD], " & _
          "p.municipal, p.construction, p.environmental, p.cultural, " & _
          "r.spectrum, r.radioelectric, r.notary_fee, " & _
          "i.structure, i.eng_layout, i.civil_work, i.logistics, i.uW_poles, i.cel_poles, i.RRU_poles, i.EDP_support, i.beacon, i.add_civil_works, i.sat_pad, " & _
          "e.AC, e.DC, e.grd, " & _
          "n.d_link, n.hwTx, n.sectors, n.hwrf, n.rf_services, n.licences " & _
          "FROM ((((((tblSites AS s " & _
          "LEFT JOIN tblDates AS d ON s.SiteCode = d.Code) " & _
          "LEFT JOIN tblUnitCosts AS u ON u.Area = s.AccessTypeID) " & _
          "LEFT JOIN tblPermits AS p ON s.SiteID = p.SiteID) " & _
          "LEFT JOIN tblRegulatory AS r ON s.SiteID = r.SiteID) " & _
          "LEFT JOIN tblInfrastructure AS i ON s.SiteID = i.SiteID) " & _
          "LEFT JOIN tblElectrical AS e ON s.SiteID = e.SiteID) " & _
          "LEFT JOIN tblNetwork AS n ON s.SiteID = n.SiteID;"
    
    ' Create query definition
    Set qdf = db.CreateQueryDef("qryProjectDashboard", sql)
    
    ' Refresh Navigation Pane display
    Application.RefreshDatabaseWindow
    
    MsgBox "Comprehensive dashboard query 'qryProjectDashboard' created successfully!", vbInformation
End Sub

