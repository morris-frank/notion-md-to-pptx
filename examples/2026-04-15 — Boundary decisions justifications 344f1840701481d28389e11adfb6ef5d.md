# Boundary decisions: justifications


<aside>
<img src="/icons/checkmark-square_orange.svg" alt="/icons/checkmark-square_orange.svg" width="40px" />

### Canonical domain truth is not UI or orchestration state

</aside>

**Final boundary**

- The canonical model is independent of UI edit state, client grouping, and ad hoc dependency resolution.
- Stable identity and server-side rule resolution are required from draft onward.

**What this absorbed**

- Rejected UI arrays and edit groupings as source of truth.
- Rejected synthetic draft IDs like `non-geo-{index}`.
- Rejected client-side dependency closure for analyses.

---

## Material

<aside>
<img src="/icons/checkmark-square_orange.svg" alt="/icons/checkmark-square_orange.svg" width="40px" />

### Material provenance is not assay execution

</aside>

**Final boundary**

- `SampleLineage` / provenance is the durable scientific root for material identity.
- Assay orders, analysis selections, and execution attempts operate on material; they do not define it.

**What this absorbed**

- Corrected `project -> analyses` ownership toward order/material/sample-set assignment.
- Corrected the simple 1:1 field sample -> lab sample model.
- Preserved composite and transformed sample chains.

---

## Material

<aside>
<img src="/icons/checkmark-square_orange.svg" alt="/icons/checkmark-square_orange.svg" width="40px" />

### Spatial assets are not investigation or engagement containers

</aside>

**Final boundary**

- `Farm` / `Field`, and later `AreaAsset` / `ManagementZone`, are durable spatial roots.
- They persist independently of any one engagement, study, or sampling campaign.

**What this absorbed**

- Corrected farm-specific language toward a broader area/asset abstraction.
- Rejected modeling land identity as owned by studies, campaigns, or commercial work.

---

## Request

<aside>
<img src="/icons/checkmark-square_orange.svg" alt="/icons/checkmark-square_orange.svg" width="40px" />

### Commercial engagement is not the scientific root

</aside>

**Final boundary**

- Commercial objects explain why work exists, who pays, what is approved, and what is promised for delivery.
- They do not own samples, lineage, evidence, or reports as truth containers.

**What this absorbed**

- Rejected one overloaded `project` as the domain root.
- Narrowed `Program` into a more explicit commercial decomposition such as `Engagement`, intake/session, and order objects.
- Rejected evidence as a child of engagement, order, or report.

---

## Request

<aside>
<img src="/icons/checkmark-square_orange.svg" alt="/icons/checkmark-square_orange.svg" width="40px" />

### Investigation framing is not the owner of land, material, or commerce

</aside>

**Final boundary**

- Trial/study/investigation concepts frame questions, protocol, and observation windows.
- They reference commercial, spatial, and material roots; they are not the owner of those roots.

**What this absorbed**

- Narrowed `Trial` as the stable aggregate name.
- Evolved toward `Study` and later `InvestigationFrame` as a framing construct rather than a universal root.

---

## Work

<aside>
<img src="/icons/checkmark-square_orange.svg" alt="/icons/checkmark-square_orange.svg" width="40px" />

### Planned and requested work is not actual execution

</aside>

**Final boundary**

- Draft, submission, approval, requested work, and actual execution are separate truths.
- Lifecycle state is multi-axis, not one linear enum.

**What this absorbed**

- Corrected the single lifecycle state model.
- Corrected execution from a thin operational detail into first-class `ExecutionRun` / `ExecutionAttempt`.
- Rejected collapsing intended analysis or package selection into execution state.

---

## Evidence

<aside>
<img src="/icons/checkmark-square_orange.svg" alt="/icons/checkmark-square_orange.svg" width="40px" />

### Reference corpus is not a project artifact

</aside>

**Final boundary**

- Benchmark/reference assets are governed and versioned independently of projects and reports.
- Benchmark outputs must pin the corpus version used.

**What this absorbed**

- Corrected benchmarking from “charting” into a first-class subsystem.
- Rejected treating the corpus as just another project artifact.

---

## Evidence

<aside>
<img src="/icons/checkmark-square_orange.svg" alt="/icons/checkmark-square_orange.svg" width="40px" />

### Evidence is not interpretation or publication

</aside>

**Final boundary**

- `EvidenceSet` is separate from `FindingSetVersion`, `Recommendation`, and `ReportDeliverableVersion`.
- Publication, reissue, and revocation are separate decisions over interpreted output.

**What this absorbed**

- Rejected reports as truth containers.
- Corrected raw measurement vs interpretation vs publication into distinct layers.