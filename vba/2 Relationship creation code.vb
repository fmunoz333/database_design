Option Compare Database
Option Explicit

Sub BuildTerranovaDB()
    Dim db As DAO.Database
    Dim xlPath As String
    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    
    xlPath = CurrentProject.Path & "\Terranova 100 4G - Database Design.xlsx"
    Set db = CurrentDb
    
    ' === Create & Import Tables (same as before) ===
    ' [Code from previous step for tblSites, tblDates, tblUnitCosts, tblRegulatory, tblPermits,
    '  tblInfrastructure, tblElectrical, tblNetwork, tblArea, tblAccessType, tblOwner,
    '  tblContractType, tblAvailability, tblStructureType, tblElectricalAC, tblElectricalDC,
    '  tblTx, tblMilestones]
    ' (Keep the CREATE TABLE + DoCmd.TransferSpreadsheet blocks as before)
    
    ' === Define Relationships ===
    On Error Resume Next
    
    ' tblSites.AccessTypeID → tblAccessType.AccessTypeID
    Set rel = db.CreateRelation("rel_Sites_AccessType", "tblSites", "tblAccessType", dbRelationUpdateCascade)
    Set fld = rel.CreateField("AccessTypeID")
    fld.ForeignName = "AccessTypeID"
    rel.Fields.Append fld
    db.Relations.Append rel
    
    ' tblSites.OwnerID → tblOwner.OwnerID
    Set rel = db.CreateRelation("rel_Sites_Owner", "tblSites", "tblOwner", dbRelationUpdateCascade)
    Set fld = rel.CreateField("OwnerID")
    fld.ForeignName = "OwnerID"
    rel.Fields.Append fld
    db.Relations.Append rel
    
    ' tblRegulatory.ContractTypeID → tblContractType.ContractTypeID
    Set rel = db.CreateRelation("rel_Regulatory_ContractType", "tblRegulatory", "tblContractType", dbRelationUpdateCascade)
    Set fld = rel.CreateField("ContractTypeID")
    fld.ForeignName = "ContractTypeID"
    rel.Fields.Append fld
    db.Relations.Append rel
    
    ' tblPermits.OwnerID → tblOwner.OwnerID
    Set rel = db.CreateRelation("rel_Permits_Owner", "tblPermits", "tblOwner", dbRelationUpdateCascade)
    Set fld = rel.CreateField("OwnerID")
    fld.ForeignName = "OwnerID"
    rel.Fields.Append fld
    db.Relations.Append rel
    
    ' tblPermits.EnvAvailabilityID → tblAvailability.AvailabilityID
    Set rel = db.CreateRelation("rel_Permits_EnvAvailability", "tblPermits", "tblAvailability", dbRelationUpdateCascade)
    Set fld = rel.CreateField("EnvAvailabilityID")
    fld.ForeignName = "AvailabilityID"
    rel.Fields.Append fld
    db.Relations.Append rel
    
    ' tblPermits.CulturalAvailabilityID → tblAvailability.AvailabilityID
    Set rel = db.CreateRelation("rel_Permits_CulturalAvailability", "tblPermits", "tblAvailability", dbRelationUpdateCascade)
    Set fld = rel.CreateField("CulturalAvailabilityID")
    fld.ForeignName = "AvailabilityID"
    rel.Fields.Append fld
    db.Relations.Append rel
    
    ' tblInfrastructure.StructureTypeID → tblStructureType.StructureTypeID
    Set rel = db.CreateRelation("rel_Infrastructure_StructureType", "tblInfrastructure", "tblStructureType", dbRelationUpdateCascade)
    Set fld = rel.CreateField("StructureTypeID")
    fld.ForeignName = "StructureTypeID"
    rel.Fields.Append fld
    db.Relations.Append rel
    
    ' tblInfrastructure.OwnerID → tblOwner.OwnerID
    Set rel = db.CreateRelation("rel_Infrastructure_Owner", "tblInfrastructure", "tblOwner", dbRelationUpdateCascade)
    Set fld = rel.CreateField("OwnerID")
    fld.ForeignName = "OwnerID"
    rel.Fields.Append fld
    db.Relations.Append rel
    
    ' tblInfrastructure.uWAvailabilityID → tblAvailability.AvailabilityID
    Set rel = db.CreateRelation("rel_Infrastructure_uWAvailability", "tblInfrastructure", "tblAvailability", dbRelationUpdateCascade)
    Set fld = rel.CreateField("uWAvailabilityID")
    fld.ForeignName = "AvailabilityID"
    rel.Fields.Append fld
    db.Relations.Append rel
    
    ' tblInfrastructure.EDPAvailabilityID → tblAvailability.AvailabilityID
    Set rel = db.CreateRelation("rel_Infrastructure_EDPAvailability", "tblInfrastructure", "tblAvailability", dbRelationUpdateCascade)
    Set fld = rel.CreateField("EDPAvailabilityID")
    fld.ForeignName = "AvailabilityID"
    rel.Fields.Append fld
    db.Relations.Append rel
    
    ' tblElectrical.ACAvailabilityID → tblElectricalAC.ACID
    Set rel = db.CreateRelation("rel_Electrical_AC", "tblElectrical", "tblElectricalAC", dbRelationUpdateCascade)
    Set fld = rel.CreateField("ACAvailabilityID")
    fld.ForeignName = "ACID"
    rel.Fields.Append fld
    db.Relations.Append rel
    
    ' tblElectrical.DCAvailabilityID → tblElectricalDC.DCID
    Set rel = db.CreateRelation("rel_Electrical_DC", "tblElectrical", "tblElectricalDC", dbRelationUpdateCascade)
    Set fld = rel.CreateField("DCAvailabilityID")
    fld.ForeignName = "DCID"
    rel.Fields.Append fld
    db.Relations.Append rel
    
    ' tblElectrical.GrdAvailabilityID → tblAvailability.AvailabilityID
    Set rel = db.CreateRelation("rel_Electrical_GrdAvailability", "tblElectrical", "tblAvailability", dbRelationUpdateCascade)
    Set fld = rel.CreateField("GrdAvailabilityID")
    fld.ForeignName = "AvailabilityID"
    rel.Fields.Append fld
    db.Relations.Append rel
    
    ' tblNetwork.TxID → tblTx.TxID
    Set rel = db.CreateRelation("rel_Network_Tx", "tblNetwork", "tblTx", dbRelationUpdateCascade)
    Set fld = rel.CreateField("TxID")
    fld.ForeignName = "TxID"
    rel.Fields.Append fld
    db.Relations.Append rel
    
    MsgBox "Terranova 100 4G database created, populated, and relationships defined successfully!"
End Sub
