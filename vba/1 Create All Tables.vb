Option Compare Database
Option Explicit

Sub BuildTerranovaDB()
    Dim db As DAO.Database
    Dim xlPath As String
    
    xlPath = CurrentProject.Path & "\Terranova 100 4G - Database Design.xlsx"
    Set db = CurrentDb
    
    ' =========================================================================
    ' 1. CREATE AND IMPORT TABLES
    ' =========================================================================
    
    ' --- tblSites ---
    DropTableIfExists db, "tblSites"
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
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "tblSites", xlPath, True, "tblSites!A1:K101"
    
    ' --- tblDates ---
    DropTableIfExists db, "tblDates"
    db.Execute "CREATE TABLE tblDates (" & _
               "DateID INTEGER PRIMARY KEY, " & _
               "SiteID INTEGER, " & _
               "Milestone INTEGER, " & _
               "[Start Plan] DATE, " & _
               "[Finish Plan] DATE, " & _
               "[Start Real] DATE, " & _
               "[Finish Real] DATE)"
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "tblDates", xlPath, True, "tblDates!A1:G1101"
    
    ' --- tblUnitCosts ---
    DropTableIfExists db, "tblUnitCosts"
    db.Execute "CREATE TABLE tblUnitCosts (" & _
               "UnitCostsID AUTOINCREMENT PRIMARY KEY, " & _
               "Area INTEGER, " & _
               "UnitCostItem TEXT(60), " & _
               "CapexItem TEXT(20), " & _
               "[Unit Cost (Euro)] CURRENCY, " & _
               "[Unit Cost USD] CURRENCY)"
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "tblUnitCosts", xlPath, True, "tblUnitCosts!A1:F46"
    
    ' --- tblRegulatory ---
    DropTableIfExists db, "tblRegulatory"
    db.Execute "CREATE TABLE tblRegulatory (" & _
               "SiteID INTEGER PRIMARY KEY, " & _
               "SpectrumTax YESNO, " & _
               "SpectrumCost CURRENCY, " & _
               "RadioelectricTax YESNO, " & _
               "RadioelectricCost CURRENCY, " & _
               "ContractTypeID INTEGER, " & _
               "NotaryFee CURRENCY)"
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "tblRegulatory", xlPath, True, "tblRegulatory!A1:G101"
    
    ' --- tblPermits ---
    DropTableIfExists db, "tblPermits"
    db.Execute "CREATE TABLE tblPermits (" & _
               "SiteID INTEGER PRIMARY KEY, " & _
               "OwnerID INTEGER, " & _
               "MunicipalCost CURRENCY, " & _
               "ConstructionCost CURRENCY, " & _
               "EnvAvailabilityID INTEGER, " & _
               "EnvironmentalCost CURRENCY, " & _
               "CulturalAvailabilityID INTEGER, " & _
               "CulturalCost CURRENCY, " & _
               "RowAvailabilityID INTEGER, " & _
               "RowCost CURRENCY, " & _
               "GridAvailabilityID INTEGER, " & _
               "GridCost CURRENCY)"
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "tblPermits", xlPath, True, "tblPermits!A1:L101"
    
    ' --- tblInfrastructure ---
    DropTableIfExists db, "tblInfrastructure"
    db.Execute "CREATE TABLE tblInfrastructure (" & _
               "SiteID INTEGER PRIMARY KEY, " & _
               "StructureAvailability INTEGER, " & _
               "StructureTypeID INTEGER, " & _
               "Height INTEGER, " & _
               "OwnerID INTEGER, " & _
               "StructureCost CURRENCY, " & _
               "LayoutCost CURRENCY, " & _
               "CivilWorkCost CURRENCY, " & _
               "LogisticsCost CURRENCY, " & _
               "uWAvailabilityID INTEGER, " & _
               "uWPolesCost CURRENCY, " & _
               "CelPolesCost CURRENCY, " & _
               "RRUPolesCost CURRENCY, " & _
               "EDPAvailabilityID INTEGER, " & _
               "EDPSupportCost CURRENCY, " & _
               "BeaconCost CURRENCY, " & _
               "AddCivilWorksCost CURRENCY, " & _
               "SatPadCost CURRENCY)"
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "tblInfrastructure", xlPath, True, "tblInfrastructure!A1:R101"
    
    ' --- tblElectrical ---
    DropTableIfExists db, "tblElectrical"
    db.Execute "CREATE TABLE tblElectrical (" & _
               "SiteID INTEGER PRIMARY KEY, " & _
               "ACAvailabilityID INTEGER, " & _
               "ACCost CURRENCY, " & _
               "DCAvailabilityID INTEGER, " & _
               "DCCost CURRENCY, " & _
               "GrdAvailabilityID INTEGER, " & _
               "GrdCost CURRENCY)"
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "tblElectrical", xlPath, True, "tblElectrical!A1:G101"
    
    ' --- tblNetwork ---
    DropTableIfExists db, "tblNetwork"
    db.Execute "CREATE TABLE tblNetwork (" & _
               "SiteID INTEGER PRIMARY KEY, " & _
               "TxID INTEGER, " & _
               "d_link DOUBLE, " & _
               "HwTxCost CURRENCY, " & _
               "sectors INTEGER, " & _
               "HwRFCost CURRENCY, " & _
               "RFServicesCost CURRENCY, " & _
               "LicencesCost CURRENCY)"
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, "tblNetwork", xlPath, True, "tblNetwork!A1:H101"
    
    ' =========================================================================
    ' 2. IMPORT SUPPORT TABLES, CONVERT TYPES & SET PRIMARY KEYS
    ' =========================================================================
    Dim supportTables As Variant
    Dim i As Long
    supportTables = Array("tblArea!A1:B9", "tblAccessType!A1:B5", "tblOwner!A1:C5", _
                          "tblContractType!A1:B4", "tblAvailability!A1:B4", "tblStructureType!A1:B7", _
                          "tblElectricalAC!A1:B4", "tblElectricalDC!A1:B5", "tblTx!A1:B6", _
                          "tblMilestones!A1:B12")
                          
    For i = LBound(supportTables) To UBound(supportTables)
        Dim fullSpec As String
        Dim tName As String
        fullSpec = CStr(supportTables(i))
        tName = Left(fullSpec, InStr(fullSpec, "!") - 1)
        
        DropTableIfExists db, tName
        DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, tName, xlPath, True, fullSpec
    Next i

    ' --- Step A: Convert imported Double ID fields to INTEGER to match main tables ---
    db.Execute "ALTER TABLE tblAccessType ALTER COLUMN AccessTypeID INTEGER"
    db.Execute "ALTER TABLE tblOwner ALTER COLUMN OwnerID INTEGER"
    db.Execute "ALTER TABLE tblContractType ALTER COLUMN ContractTypeID INTEGER"
    db.Execute "ALTER TABLE tblAvailability ALTER COLUMN AvailabilityID INTEGER"
    db.Execute "ALTER TABLE tblStructureType ALTER COLUMN StructureTypeID INTEGER"
    db.Execute "ALTER TABLE tblElectricalAC ALTER COLUMN ACID INTEGER"
    db.Execute "ALTER TABLE tblElectricalDC ALTER COLUMN DCID INTEGER"
    db.Execute "ALTER TABLE tblTx ALTER COLUMN TxID INTEGER"

    ' --- Step B: Add Primary Key constraints ---
    db.Execute "ALTER TABLE tblAccessType ADD CONSTRAINT PK_tblAccessType PRIMARY KEY (AccessTypeID)"
    db.Execute "ALTER TABLE tblOwner ADD CONSTRAINT PK_tblOwner PRIMARY KEY (OwnerID)"
    db.Execute "ALTER TABLE tblContractType ADD CONSTRAINT PK_tblContractType PRIMARY KEY (ContractTypeID)"
    db.Execute "ALTER TABLE tblAvailability ADD CONSTRAINT PK_tblAvailability PRIMARY KEY (AvailabilityID)"
    db.Execute "ALTER TABLE tblStructureType ADD CONSTRAINT PK_tblStructureType PRIMARY KEY (StructureTypeID)"
    db.Execute "ALTER TABLE tblElectricalAC ADD CONSTRAINT PK_tblElectricalAC PRIMARY KEY (ACID)"
    db.Execute "ALTER TABLE tblElectricalDC ADD CONSTRAINT PK_tblElectricalDC PRIMARY KEY (DCID)"
    db.Execute "ALTER TABLE tblTx ADD CONSTRAINT PK_tblTx PRIMARY KEY (TxID)"
    
    ' =========================================================================
    ' 3. DEFINE RELATIONSHIPS
    ' =========================================================================
    
    ' tblSites -> Lookups
    AddRelation db, "rel_Sites_AccessType", "tblAccessType", "AccessTypeID", "tblSites", "AccessTypeID"
    AddRelation db, "rel_Sites_Owner", "tblOwner", "OwnerID", "tblSites", "OwnerID"
    
    ' tblRegulatory -> Lookups
    AddRelation db, "rel_Regulatory_ContractType", "tblContractType", "ContractTypeID", "tblRegulatory", "ContractTypeID"
    
    ' tblPermits -> Lookups
    AddRelation db, "rel_Permits_Owner", "tblOwner", "OwnerID", "tblPermits", "OwnerID"
    AddRelation db, "rel_Permits_EnvAvailability", "tblAvailability", "AvailabilityID", "tblPermits", "EnvAvailabilityID"
    AddRelation db, "rel_Permits_CulturalAvailability", "tblAvailability", "AvailabilityID", "tblPermits", "CulturalAvailabilityID"
    
    ' tblInfrastructure -> Lookups
    AddRelation db, "rel_Infrastructure_StructureType", "tblStructureType", "StructureTypeID", "tblInfrastructure", "StructureTypeID"
    AddRelation db, "rel_Infrastructure_Owner", "tblOwner", "OwnerID", "tblInfrastructure", "OwnerID"
    AddRelation db, "rel_Infrastructure_uWAvailability", "tblAvailability", "AvailabilityID", "tblInfrastructure", "uWAvailabilityID"
    AddRelation db, "rel_Infrastructure_EDPAvailability", "tblAvailability", "AvailabilityID", "tblInfrastructure", "EDPAvailabilityID"
    
    ' tblElectrical -> Lookups
    AddRelation db, "rel_Electrical_AC", "tblElectricalAC", "ACID", "tblElectrical", "ACAvailabilityID"
    AddRelation db, "rel_Electrical_DC", "tblElectricalDC", "DCID", "tblElectrical", "DCAvailabilityID"
    AddRelation db, "rel_Electrical_GrdAvailability", "tblAvailability", "AvailabilityID", "tblElectrical", "GrdAvailabilityID"
    
    ' tblNetwork -> Lookups
    AddRelation db, "rel_Network_Tx", "tblTx", "TxID", "tblNetwork", "TxID"

    MsgBox "Terranova 100 4G database created, populated, and relationships defined successfully!", vbInformation
    
End Sub

' =============================================================================
' 4. HELPER PROCEDURES (Placed outside the main Sub, below End Sub)
' =============================================================================

Private Sub AddRelation(ByRef db As DAO.Database, _
                        ByVal strRelName As String, _
                        ByVal strPrimaryTable As String, _
                        ByVal strPrimaryKeyField As String, _
                        ByVal strForeignTable As String, _
                        ByVal strForeignKeyField As String)
    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    
    On Error Resume Next
    db.Relations.Delete strRelName
    Err.Clear
    On Error GoTo 0
    
    Set rel = db.CreateRelation(strRelName, strPrimaryTable, strForeignTable, dbRelationUpdateCascade)
    Set fld = rel.CreateField(strPrimaryKeyField)
    fld.ForeignName = strForeignKeyField
    rel.Fields.Append fld
    
    db.Relations.Append rel
End Sub

Private Sub DropTableIfExists(ByRef db As DAO.Database, ByVal strTableName As String)
    On Error Resume Next
    db.Execute "DROP TABLE [" & strTableName & "]"
    Err.Clear
    On Error GoTo 0
End Sub



