Option Compare Database
Option Explicit

Sub BuildTerranovaDB()
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim xlPath As String
    
    ' Path to your Excel workbook
    xlPath = CurrentProject.Path & "\Terranova 100 4G - Database Design.xlsx"
    
    Set db = CurrentDb
    
    ' === Example: Create tblSites ===
    On Error Resume Next
    db.Execute "DROP TABLE tblSites"
    On Error GoTo 0
    
    db.Execute "CREATE TABLE tblSites (" & _
               "SiteID INTEGER PRIMARY KEY, " & _
               "SiteCode TEXT(10), " & _
               "SiteName TEXT(50), " & _
               "Province TEXT(50), " & _
               "Municipality TEXT(50), " & _
               "Parrish TEXT(50), " & _
               "Latitude DOUBLE, " & _
               "Longitude DOUBLE, " & _
               "Population LONG, " & _
               "AccessTypeID INTEGER, " & _
               "OwnerID INTEGER)"
    
    ' === Import data from Excel sheet tblSites ===
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, _
        "tblSites", xlPath, True, "tblSites!"
    
    ' === Repeat for other tables ===
    ' Example: tblDates
    On Error Resume Next
    db.Execute "DROP TABLE tblDates"
    On Error GoTo 0
    
    db.Execute "CREATE TABLE tblDates (" & _
               "Code TEXT(10), " & _
               "Milestone INTEGER, " & _
               "[Start Plan] DATE, " & _
               "[Finish Plan] DATE, " & _
               "[Start Real] DATE, " & _
               "[Finish Real] DATE, " & _
               "PRIMARY KEY (Code, Milestone))"
    
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, _
        "tblDates", xlPath, True, "tblDates!"
    
    ' Continue with tblUnitCosts, tblRegulatory, tblPermits, etc.
    
    MsgBox "Terranova 100 4G database structure and data imported successfully!"
End Sub
