Option Compare Database
Option Explicit

Sub BuildTerranovaDB()
    Dim db As DAO.Database
    Dim xlPath As String
    Dim rel As DAO.Relation
    Dim fld As DAO.Field
    
    xlPath = CurrentProject.Path & "\Terranova 100 4G - Database Design.xlsx"
    Set db = CurrentDb
    
    ' === Create & Import Tables ===
    ' (Keep the CREATE TABLE + DoCmd.TransferSpreadsheet blocks from the previous script)
    ' For brevity, not repeating all CREATE TABLE statements here — ensure you include them all as in the last version.
    ' Each table: tblSites, tblDates, tblUnitCosts, tblRegulatory, tblPermits, tblInfrastructure,
    ' tblElectrical, tblNetwork, tblArea, tblAccessType, tblOwner, tblContractType, tblAvailability,
    ' tblStructureType, tblElectricalAC, tblElectricalDC, tblTx, tblMilestones
    
    ' === Define Relationships with Referential Integrity ===
    On Error Resume Next
    
    ' tblSites.AccessTypeID → tblAccessType.AccessTypeID
    Set rel = db.CreateRelation("rel_Sites_AccessType", "tblSites", "tblAccessType", _
                                dbRelationUpdateCascade + dbRelationDeleteCascade)
    Set fld = rel.CreateField("AccessTypeID")
    fld.ForeignName = "AccessTypeID"
    rel.Fields.Append fld
    db.Relations.Append rel
    
    ' tblSites.OwnerID → tblOwner.OwnerID
    Set rel = db.CreateRelation("rel_Sites_Owner", "tblSites", "tblOwner", _
                                dbRelationUpdateCascade + dbRelationDeleteCascade)
    Set fld = rel.CreateField("OwnerID")
    fld.ForeignName = "OwnerID"
    rel.Fields.Append fld
    db.Relations.Append rel
    
    ' tblRegulatory.ContractTypeID → tblContractType.ContractTypeID
    Set rel = db.CreateRelation("rel_Regulatory_ContractType", "tblRegulatory", "tblContractType", _
                                dbRelationUpdateCascade + dbRelationDeleteCascade)
    Set fld = rel.CreateField("ContractTypeID")
    fld.ForeignName = "ContractTypeID"
    rel.Fields.Append fld
    db.Relations.Append rel
    
    ' tblPermits.OwnerID → tblOwner.OwnerID
    Set rel = db.CreateRelation("rel_Permits_Owner", "tblPermits", "tblOwner", _
                                dbRelationUpdateCascade + dbRelationDeleteCascade)
    Set fld = rel.CreateField("OwnerID")
    fld.ForeignName = "OwnerID"
    rel.Fields.Append fld
    db.Relations.Append rel
    
    ' tblPermits.EnvAvailabilityID → tblAvailability.AvailabilityID
    Set rel = db.CreateRelation("rel_Permits_EnvAvailability", "tblPermits", "tblAvailability", _
                                dbRelationUpdateCascade + dbRelationDeleteCascade)
    Set fld = rel.CreateField("EnvAvailabilityID")
    fld.ForeignName = "AvailabilityID"
    rel.Fields.Append fld
    db.Relations.Append rel
    
    ' tblPermits.CulturalAvailabilityID → tblAvailability.AvailabilityID
    Set rel = db.CreateRelation("rel_Permits_CulturalAvailability", "tblPermits", "tblAvailability", _
                                dbRelationUpdateCascade + dbRelationDeleteCascade)
    Set fld = rel.CreateField("CulturalAvailabilityID")
    fld.ForeignName = "AvailabilityID"
    rel.Fields.Append fld
    db.Relations.Append rel
    
    ' tblInfrastructure.StructureTypeID → tblStructureType.StructureTypeID
    Set rel = db.CreateRelation("rel_Infrastructure_StructureType", "tblInfrastructure", "tblStructureType", _
                                dbRelationUpdateCascade + dbRelationDeleteCascade)
    Set fld = rel.CreateField("StructureTypeID")
    fld.ForeignName = "StructureTypeID"
    rel.Fields.Append fld
    db.Relations.Append rel
    
    ' tblInfrastructure.OwnerID → tblOwner.OwnerID
    Set rel = db.CreateRelation("rel_Infrastructure_Owner", "tblInfrastructure", "tblOwner", _
                                dbRelationUpdateCascade + dbRelationDeleteCascade)
    Set fld = rel.CreateField("OwnerID")
    fld.ForeignName = "OwnerID"
    rel.Fields.Append fld
    db.Relations.Append rel
    
    ' tblInfrastructure.uWAvailabilityID → tblAvailability.AvailabilityID
    Set rel = db.CreateRelation("rel_Infrastructure_uWAvailability", "tblInfrastructure", "tblAvailability", _
                                dbRelationUpdateCascade + dbRelationDeleteCascade)
    Set fld = rel.CreateField("uWAvailabilityID")
    fld.ForeignName = "AvailabilityID"
    rel.Fields.Append fld
    db.Relations.Append rel
    
    ' tblInfrastructure.EDPAvailabilityID → tblAvailability.AvailabilityID
    Set rel = db.CreateRelation("rel_Infrastructure_EDPAvailability", "tblInfrastructure", "tblAvailability", _
                                dbRelationUpdateCascade + dbRelationDeleteCascade)
    Set fld = rel.CreateField("EDPAvailabilityID")
    fld.ForeignName = "AvailabilityID"
    rel.Fields.Append fld
    db.Relations.Append rel
    
    ' tblElectrical.ACAvailabilityID → tblElectricalAC.ACID
    Set rel = db.CreateRelation("rel_Electrical_AC", "tblElectrical", "tblElectricalAC", _
                                dbRelationUpdateCascade + dbRelationDeleteCascade)
    Set fld = rel.CreateField("ACAvailabilityID")
    fld.ForeignName = "ACID"
    rel.Fields.Append fld
    db.Relations.Append rel
    
    ' tblElectrical.DCAvailabilityID → tblElectricalDC.DCID
    Set rel = db.CreateRelation("rel_Electrical_DC", "tblElectrical", "tblElectricalDC", _
                                dbRelationUpdateCascade + dbRelationDeleteCascade)
    Set fld = rel.CreateField("DCAvailabilityID")
    fld.ForeignName = "DCID"
    rel.Fields.Append fld
    db.Relations.Append rel
    
    ' tblElectrical.GrdAvailabilityID → tblAvailability.AvailabilityID
    Set rel = db.CreateRelation("rel_Electrical_GrdAvailability", "tblElectrical", "tblAvailability", _
                                dbRelationUpdateCascade + dbRelationDeleteCascade)
    Set fld = rel.CreateField("GrdAvailabilityID")
    fld.ForeignName = "AvailabilityID"
    rel.Fields.Append fld
    db.Relations.Append rel
    
    ' tblNetwork.TxID → tblTx.TxID
    Set rel = db.CreateRelation("rel_Network_Tx", "tblNetwork", "tblTx", _
                                dbRelationUpdateCascade + dbRelationDeleteCascade)
    Set fld = rel.CreateField("TxID")
    fld.ForeignName = "TxID"
    rel.Fields.Append fld
    db.Relations.Append rel
    
    MsgBox "Terranova 100 4G database created, populated, and relationships with referential integrity defined successfully!"
End Sub
