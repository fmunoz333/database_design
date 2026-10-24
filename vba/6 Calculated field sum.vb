Sub CreateProjectDashboardQueryWithTotal()
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim sql As String
    
    Set db = CurrentDb
    
    ' Drop existing query if it exists
    On Error Resume Next
    db.QueryDefs.Delete "qryProjectDashboard"
    On Error GoTo 0
    
    ' Build SQL for comprehensive dashboard with total cost
    sql = "SELECT " & _
          "s.SiteID, s.SiteCode, s.SiteName, s.Province, s.Municipality, s.Parrish, " & _
          "d.Milestone, d.[Start Plan], d.[Finish Plan], d.[Start Real], d.[Finish Real], " & _
          "u.UnitCostItem, u.[Unit Cost (Euro)], u.[Unit Cost USD], " & _
          "p.municipal, p.construction, p.environmental, p.cultural, " & _
          "r.spectrum, r.radioelectric, r.notary_fee, " & _
          "i.structure, i.eng_layout, i.civil_work, i.logistics, i.uW_poles, i.cel_poles, i.RRU_poles, i.EDP_support, i.beacon, i.add_civil_works, i.sat_pad, " & _
          "e.AC, e.DC, e.grd, " & _
          "n.d_link, n.hwTx, n.sectors, n.hwrf, n.rf_services, n.licences, " & _
          "Nz(u.[Unit Cost (Euro)],0) + Nz(p.municipal,0) + Nz(p.construction,0) + Nz(p.environmental,0) + Nz(p.cultural,0) + " & _
          "Nz(r.spectrum,0) + Nz(r.radioelectric,0) + Nz(r.notary_fee,0) + " & _
          "Nz(i.structure,0) + Nz(i.eng_layout,0) + Nz(i.civil_work,0) + Nz(i.logistics,0) + Nz(i.uW_poles,0) + Nz(i.cel_poles,0) + Nz(i.RRU_poles,0) + Nz(i.EDP_support,0) + Nz(i.beacon,0) + Nz(i.add_civil_works,0) + Nz(i.sat_pad,0) + " & _
          "Nz(e.AC,0) + Nz(e.DC,0) + Nz(e.grd,0) + " & _
          "Nz(n.hwTx,0) + Nz(n.hwrf,0) + Nz(n.rf_services,0) + Nz(n.licences,0) AS TotalProjectCost " & _
          "FROM ((((((((tblSites AS s " & _
          "LEFT JOIN tblDates AS d ON s.SiteCode = d.Code) " & _
          "LEFT JOIN tblUnitCosts AS u ON u.Area = s.AccessTypeID) " & _
          "LEFT JOIN tblPermits AS p ON s.SiteID = p.SiteID) " & _
          "LEFT JOIN tblRegulatory AS r ON s.SiteID = r.SiteID) " & _
          "LEFT JOIN tblInfrastructure AS i ON s.SiteID = i.SiteID) " & _
          "LEFT JOIN tblElectrical AS e ON s.SiteID = e.SiteID) " & _
          "LEFT JOIN tblNetwork AS n ON s.SiteID = n.SiteID);"
    
    ' Create query
    Set qdf = db.CreateQueryDef("qryProjectDashboard", sql)
    
    MsgBox "Comprehensive dashboard query 'qryProjectDashboard' with TotalProjectCost created successfully!"
End Sub
