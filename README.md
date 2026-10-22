# Terranova 100-Site 4G Rural Rollout — Relational Database & Reporting System

**Microsoft Access · SQL · Visual Basic (VBA) · Data Modelling · Automated Reporting**

*Author: Fernando Muñoz — Project Management, Data Analytics & Automation*

---

## 1. Project Title & Overview

**Terranova** is a normalized relational database and reporting system built to plan, cost, and track a **100-site rural 4G network rollout** across four provinces of northern Spain (A Coruña, Asturias, Lugo, Pontevedra), serving an estimated **99,320 inhabitants**.

The project replaces scattered spreadsheets with a single source of truth that connects site data, regulatory and permitting status, infrastructure, power, transmission/RF design, a 45-item unit-cost catalogue, and an 11-milestone schedule for every site. On top of that model sit parameterized SQL queries, a multi-subreport Site Profile report, and a Visual Basic automation layer that turns raw project data into management-ready dashboards and recurring reports.

### Project at a glance

| Metric | Value |
|---|---|
| Sites modelled | **100** (4 provinces, 98 municipalities) |
| Population served | **99,320** inhabitants |
| Database tables | **19** (9 core + 10 lookup), 16 defined relationships |
| Cost lines (site × catalogue item) | **4,600** (46 per site) |
| Unit-cost catalogue | **45 items** across **8 cost areas**, EUR and USD |
| Schedule records | **1,100** (11 milestones × 100 sites, planned vs. actual) |
| Total modelled CAPEX | **€9.90 M / US$11.48 M** (catalogue-priced) |
| Saved analytical queries | **7** (SQL, including parameterized search) |
| Report | **Site Profile** report with **7 subreports** |

---

## 2. Role & Responsibilities

**Role:** Database Designer & Project Reporting Lead

- **Design** — Conceived the data architecture: entity-relationship model, normalization to third normal form, naming conventions, lookup/reference tables, primary and foreign keys, and the cost-model redesign described below.
- **Planning** — Defined the reporting requirements and KPIs (CAPEX by area/province/site, schedule variance by milestone, permit and infrastructure readiness) and mapped each one to a query, a report, or a dashboard view before building.
- **Execution** — Built the tables, relationships, SQL queries, report layouts, and VBA automation; loaded and reconciled the data; validated totals.
- **Regular reporting** — Produced recurring status reports and refreshed dashboards as new data arrived, translating findings into conclusions and recommended actions for project stakeholders.

---

## 3. Challenge / Problem

A rural rollout of 100 sites is a data-integration problem as much as an engineering one. Each site carries information owned by different teams: land and negotiation, permits, civil works, power, transmission, RF, and regulatory fees. Key challenges:

1. **Fragmented data.** Cost, schedule, and technical attributes lived in separate sheets, making site-level and portfolio-level answers slow and error-prone.
2. **Unstructured cost data.** The first design stored roughly 27 cost values as separate columns across five domain tables, with the unit-cost catalogue unlinked to them. Totals could not be traced back to a price and a quantity.
3. **No consistent view of schedule performance.** Planned and actual dates for 11 milestones per site needed to be compared systematically to expose delays early.
4. **Repeatable reporting.** Stakeholders needed the same trusted numbers every reporting cycle, without rebuilding reports by hand.

**Goal:** a normalized, auditable database with query-driven dashboards and automated reports, able to answer *where is the money going, what is late, and which sites need attention* in seconds.

---

## 4. Actions / Process

### 4.1 Methodology

- **Requirements-first design** aligned with **PMBOK** knowledge areas (cost, schedule, scope) so every table and query traces to a project-management question.
- **Relational modelling** with an ERD, normalization to 3NF, surrogate keys, and lookup tables to enforce consistent vocabularies (owner type, access type, structure type, availability status, transmission medium, and so on).
- **Structured improvement cycle** (define → measure → analyze → improve → control), applied to the reporting process itself.

### 4.2 Steps performed

1. **Data modelling.** Designed 9 core tables (`tblSites`, `tblDates`, `tblRegulatory`, `tblPermits`, `tblInfrastructure`, `tblElectrical`, `tblNetwork`, `tblUnitCosts`, `tblSiteCostLine`) and 10 lookup tables, with foreign keys enforcing consistency.
2. **Cost-model redesign.** Replaced ~27 scattered cost columns with a single **`tblSiteCostLine`** fact table (Site × Unit Cost × Quantity), linked by foreign key to `tblSites` and `tblUnitCosts`. Every cost is now `Quantity × Unit Cost`, fully traceable, and the catalogue can be repriced in one place. Totals were reconciled against the original model.
3. **Query layer (SQL).** Built reusable queries for cost roll-ups, schedule variance, geographic and technical distributions, and a parameterized site search.
4. **Report layer.** Designed the **Site Profile** report, a one-page-per-site view composed of 7 subreports (dates, regulatory, permits, infrastructure, electrical, network, cost lines).
5. **Automation (VBA).** Added Visual Basic routines to generate reports on demand, export them to PDF per site or per province, and standardize the recurring reporting cycle.
6. **Validation & reporting.** Cross-checked totals, tested edge cases (null actual dates, zero-quantity lines), and issued regular reports with findings and conclusions.

### 4.3 Key SQL (excerpts)

**Total CAPEX per site**

```sql
SELECT s.SiteID, s.SiteCode, s.SiteName, s.Province,
       Sum(scl.Quantity * uc.[Unit Cost USD]) AS TotalCostUSD
FROM (tblSites AS s
      INNER JOIN tblSiteCostLine AS scl ON s.SiteID = scl.SiteID)
      INNER JOIN tblUnitCosts AS uc ON scl.UnitCostID = uc.UnitCostsID
GROUP BY s.SiteID, s.SiteCode, s.SiteName, s.Province
ORDER BY Sum(scl.Quantity * uc.[Unit Cost USD]) DESC;
```

**Schedule variance with status classification**

```sql
SELECT s.SiteCode, s.SiteName, m.MilestoneName,
       d.[Finish Plan], d.[Finish Real],
       DateDiff("d", d.[Finish Plan], d.[Finish Real]) AS VarianceDays,
       IIf(d.[Finish Real] Is Null, "Not Completed",
           IIf(DateDiff("d", d.[Finish Plan], d.[Finish Real]) > 0,
               "Late", "On Time")) AS Status
FROM (tblDates AS d
      INNER JOIN tblSites AS s ON d.SiteID = s.SiteID)
      INNER JOIN tblMilestones AS m ON d.Milestone = m.MilestoneID
ORDER BY DateDiff("d", d.[Finish Plan], d.[Finish Real]) DESC;
```

**Parameterized site search (prompts for Province / Access Type, blank = all)**

```sql
SELECT s.SiteID, s.SiteCode, s.SiteName, s.Province, s.Municipality,
       at.AccessTypeName, o.OwnerCode
FROM (tblSites AS s
      INNER JOIN tblAccessType AS at ON s.AccessTypeID = at.AccessTypeID)
      INNER JOIN tblOwner AS o ON s.OwnerID = o.OwnerID
WHERE (s.Province = [Enter Province or leave blank:]
       OR [Enter Province or leave blank:] Is Null)
  AND (at.AccessTypeName = [Enter Access Type or leave blank:]
       OR [Enter Access Type or leave blank:] Is Null);
```

### 4.4 Key VBA (representative automation module)

```vba
' Export the Site Profile report to PDF for every site in a province.
Public Sub ExportProvinceReports(ByVal strProvince As String, ByVal strFolder As String)
    Dim rs As DAO.Recordset
    On Error GoTo ErrHandler

    Set rs = CurrentDb.OpenRecordset( _
        "SELECT SiteID, SiteCode FROM tblSites WHERE Province = '" & _
        Replace(strProvince, "'", "''") & "' ORDER BY SiteCode")

    Do While Not rs.EOF
        DoCmd.OpenReport "rptSiteProfile", acViewPreview, , _
                         "SiteID = " & rs!SiteID, acHidden
        DoCmd.OutputTo acOutputReport, "rptSiteProfile", acFormatPDF, _
                       strFolder & "\SiteProfile_" & rs!SiteCode & ".pdf"
        DoCmd.Close acReport, "rptSiteProfile", acSaveNo
        rs.MoveNext
    Loop

    rs.Close
    MsgBox "Reports exported to " & strFolder, vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Export failed: " & Err.Description, vbExclamation
End Sub
```

### 4.5 Tools

Microsoft Access (ACE engine) · SQL (Access/Jet dialect) · Visual Basic for Applications · Entity-relationship modelling · PMBOK cost & schedule baselines · Excel (data staging and dashboard front-ends)

---

## 5. Results / Impact

All figures below are computed directly from the database.

### Delivered system

- **One source of truth** for 100 sites, 4,600 cost lines, and 1,100 schedule records, with 16 defined relationships and no orphaned site or schedule records.
- **Fully traceable costing:** every dollar of the **US$11.48 M** CAPEX baseline resolves to *quantity × catalogue price*, replacing ~27 hard-coded cost columns with one fact table.
- **Instant roll-ups** by site, area, province, owner, and access type through saved queries.

### Insights produced

| Finding | Evidence |
|---|---|
| **RF and Infrastructure drive 69% of CAPEX** | RF US$4.49 M (39.1%) + Infrastructure US$3.40 M (29.6%) of US$11.48 M |
| **Average site CAPEX of US$114.8 K**, ranging US$49.4 K to US$174.1 K | `qrySiteTotalCost`; the 3.5× spread points to where cost control matters most |
| **Cost per inhabitant covered: ≈ US$116** | US$11.48 M ÷ 99,320 inhabitants |
| **Standardized structure strategy** | 83 of 100 sites use monopoles; mean structure height 21.1 m |
| **Access type affects cost** | Pedestrian-access sites average US$102 K vs. US$120 K for aerial-access sites |
| **Land-tenure mix** | 35 sites on company-owned property, 24 on public schools, 21 private, 20 local-government facilities; 44 agreements, 35 own sites, 21 leases |
| **Schedule is tightly controlled but consistently slightly late** | Of 736 milestones with a recorded finish, **612 (83%) closed after plan by an average of 2.6 days; none by more than 5 days** |
| **Transmission, Infrastructure, and Permits slip the most** | Average slip of 3.5, 3.3, and 3.2 days respectively, flagging them as the milestones to protect |
| **Progress snapshot** | 60 sites with construction finished and 45 sites ready, from 1,100 tracked milestones |

### Business value

- **Faster, repeatable reporting** with one-click PDF generation per site or province, replacing manual assembly.
- **Early-warning visibility** on milestones that slip, supporting corrective action before delays compound.
- **Better budget decisions** through transparent cost drivers and a single editable price catalogue (EUR and USD).
- **Audit-ready numbers**: one definition of cost and status, reused across every report.

---

## 6. Visuals / Artifacts

### Generated from the database

| | |
|---|---|
| ![CAPEX by area](assets/capex_by_area.png) | ![Schedule slip by milestone](assets/schedule_slip_by_milestone.png) |
| ![Site cost distribution](assets/site_cost_distribution.png) | |

### Screenshots to add (recommended)

| # | Visual | Why it helps |
|---|---|---|
| 1 | **Relationships window / ERD** | Shows data-modelling skill at a glance |
| 2 | **Query design view of `qrySiteTotalCost` + its datasheet result** | Proves SQL joins and aggregation |
| 3 | **`qryScheduleVariance` datasheet** (sorted by variance) | Shows analytical thinking, not just structure |
| 4 | **Site Profile report (page 1) with its subreports** | The most visual deliverable; use a clean, anonymized site |
| 5 | **VBA editor with the export routine** | Evidence of Visual Basic skills |
| 6 | **Exported PDF** opened beside the Access window | Demonstrates the end-to-end automated workflow |
| 7 | **Before/after cost model** (scattered columns → `tblSiteCostLine`) | Tells the design-improvement story |
| 8 | **Navigation Pane** showing tables, queries, and reports | Conveys scope and organization |

Suggested folder structure for this repository:

```
/Terranova-4G-Database
├── README.md
├── Terranova_Project_Description.docx
├── assets/            (charts and screenshots)
├── sql/               (exported query SQL)
├── vba/               (exported .bas modules)
└── docs/              (ERD, data dictionary)
```

---

## 7. Lessons Learned

**Technical**

- **Model costs as facts, not columns.** Moving to a quantity × unit-price structure made totals auditable and repricing trivial. Designing the catalogue relationship from day one avoids a mid-project redesign.
- **Enforce referential integrity early.** Lookup tables and foreign keys prevent inconsistent data entry and catch errors at load time rather than at reporting time.
- **Treat nulls deliberately.** Milestones without actual dates must be classified ("Not Completed"), never silently dropped, or schedule metrics are overstated.
- **Parameterized queries beat duplicated queries.** One search query with optional parameters replaced a family of near-identical filters.

**Management**

- **Start from the decisions, not the data.** Defining the questions stakeholders ask first kept the database lean and the reports useful.
- **Small, consistent delays add up.** A 2–3 day average slip looks harmless per milestone but compounds across 11 sequential milestones; tracking by milestone type shows where to intervene.

**Benefits to the organization**

1. A **single source of truth** shared by engineering, finance, permitting, and project control.
2. **Cost transparency:** clear visibility of what drives CAPEX (RF and Infrastructure ≈ 69%) and where per-site variation comes from.
3. **Proactive schedule control** through objective, milestone-level variance reporting.
4. **Time savings and consistency** from automated, repeatable reports.
5. **Scalability:** the same model extends to additional sites by adding rows, not redesigning the structure.

---

*Data and figures are derived from the Terranova project database. Costs are catalogue-priced estimates (EUR converted to USD at 1.16).*
