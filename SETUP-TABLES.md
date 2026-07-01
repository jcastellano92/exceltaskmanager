# Workbook tables — required columns

The add-in renders a UI over Excel **named tables** (Insert → Table, then Table Design → name it).
Column **header text must match exactly** (case-sensitive). Extra columns are ignored; the app
only reads/writes the columns below.

> **🆕 New in the redesign** — create these before using the new features:
> - **UpdatesTable** — the chronological Updates feed on every task **and** subtask.
> - **MilestonesTable** — standalone dated milestone lines on the Roadmap (milestones are
>   **no longer** a task flag — the old `IsMilestone` task column is unused now).
> - Two new optional columns on **TasksTable**: **`Health`** and **`Slips`**.

---

## TasksTable
The core table. `CostOfDelay`, `WSJF`, and `AttachmentCount` are **formula columns** — keep your
sheet formulas there; the app reads them but never overwrites them.

| Column | Notes |
|---|---|
| `TaskID` | number, unique (auto = max+1 on create) |
| `Title` | |
| `WorkstreamID` | links to WorkstreamsTable (e.g. `WS05`) |
| `GoalID` | one or more goal IDs, `;`-separated (e.g. `G1;G2`) — drives roadmap bar colour |
| `Owner` | single name (matches OwnersTable) |
| `Contributors` | extra people, `;`-separated |
| `StartDate` | ISO date — roadmap bar start |
| `Quarter` | `Q1`–`Q4` (matches QuartersTable) |
| `Status` | matches StatusesTable / board columns |
| `LinkedEpic` | optional |
| `UserBusinessValue` | WSJF input (Fibonacci) |
| `TimeCriticality` | WSJF input |
| `RiskReduction` | WSJF input |
| `CostOfDelay` | **formula** = UBV+TC+RR |
| `JobSize` | WSJF input (effort) |
| `WSJF` | **formula** = CostOfDelay ÷ JobSize |
| `PercentComplete` | 0–100 |
| `ColumnOrder` | board ordering |
| `DueDate` | ISO date — roadmap bar end |
| `Tags` | `;`-separated |
| `BlockedByTaskIDs` | `;`-separated TaskIDs |
| `Description` | |
| `Source` | e.g. `Manual` |
| `Archived` | `Yes`/`No` |
| `CreatedDate`, `CreatedBy` | |
| `LastUpdated`, `UpdatedBy` | concurrency stamps (app-written) |
| `AttachmentCount` | **formula** (count from AttachmentsTable) |
| `Health` | 🆕 `On Track` / `At Risk` / `Off Track` (blank = none) |
| `Slips` | 🆕 signed integer — net quarter moves. `+1` per Delay, `−1` per Accelerate. Drives the «/» chip |
| `RoadmapGroup` | 🆕 text — name of the roadmap group this task belongs to (groups are per-workstream). Optional; powers roadmap grouping, the List column/filter, and bulk |
| `CommitmentID` | optional (written through if the column exists) |

## SubtasksTable
| Column | Notes |
|---|---|
| `SubtaskID` | number, unique |
| `ParentTaskID` | the owning TaskID |
| `Text` | |
| `Done` | `Yes`/`No` |
| `Order` | |
| `DueDate` | |
| `CompletedDate` | set when marked done |
| `Owner` | |
| `LastUpdated`, `UpdatedBy` | |

## UpdatesTable 🆕
Chronological notes for a task **or** a subtask (the Updates tab + subtask modal).
| Column | Notes |
|---|---|
| `UpdateID` | number, unique |
| `ParentType` | `Task` or `Subtask` |
| `ParentID` | the TaskID or SubtaskID |
| `Text` | |
| `AddedBy` | app-written (current user) |
| `AddedDate` | app-written (ISO timestamp) |

## MilestonesTable 🆕
Standalone dated lines on the Roadmap (not tasks).
| Column | Notes |
|---|---|
| `MilestoneID` | number, unique |
| `Title` | |
| `Date` | ISO date — where the line sits |
| `Quarter` | auto-derived from `Date` on save |
| `GoalID` | optional — colours the marker by goal |
| `Color` | optional hex override (e.g. `#7a5cff`); blank = goal colour |
| `Notes` | optional |

## AttachmentsTable
| Column | Notes |
|---|---|
| `AttachmentID`, `ParentTaskID`, `Label`, `Url`, `Type`, `AddedBy`, `AddedDate` | `Type` ∈ web/confluence/sharepoint/file |

## ActivityLogTable
| Column |
|---|
| `LogID`, `Timestamp`, `EntityType`, `EntityID`, `Action`, `ChangedBy`, `FieldChanged`, `OldValue`, `NewValue`, `Notes` |

## WorkstreamsTable
| Column | Notes |
|---|---|
| `WorkstreamID` | e.g. `WS05` |
| `Name` | |
| `Owner`, `Status` | |
| `Goals` | `;`/`,`-separated GoalIDs |
| `Quarters` | optional, `,`-separated |
| `UserStory`, `Metric1`, `Metric2`, `Metric3` | optional |

## GoalsTable
| Column | Notes |
|---|---|
| `GoalID` | e.g. `G1` |
| `ShortName` | label shown in the UI |
| `GoalName` | optional long name |

## QuartersTable
Doubles as the quarter **list** (first column) and the quarter **date bounds** (drives the roadmap axis & "days left").
| Column | Notes |
|---|---|
| `Quarter` | `Q1`–`Q4` |
| `StartDate`, `EndDate` | ISO dates |

## CapacityTable
Per-person planning capacity.
| Column | Notes |
|---|---|
| `Person` | matches OwnersTable |
| `Availability` | 0–1 default availability |
| `AvailQ1`–`AvailQ4` | optional per-quarter overrides (0 is valid) |
| `BaselineOverride` | optional; else demonstrated Q1 velocity |

## Config single-column lists
Each is a one-column table whose **first column** holds the values:
`OwnersTable`, `TypesTable`, `YesNoTable`.

## StatusesTable (board columns)
The first column holds the status names (these are your Kanban columns). Two **optional** columns make it richer — all editable from the in-app **Config → Board columns** screen:

| Column | Notes |
|---|---|
| `Status` (first col) | the status name, e.g. `Backlog`, `Planned/Ready`, `In Progress`, `Blocked`, `Done` |
| `Order` 🆕 | number — left→right board-column order. With this column, reordering in the app saves here (otherwise order is per-device only) |
| `Color` 🆕 | hex colour (e.g. `#6366f1`) — the **one** colour used for that status everywhere (board, list, roadmap bars, workstream mix) |
| `Bucket` 🆕 | `ready` / `active` / `blocked` / `done` — categorises the status so the app **adapts to renamed statuses**. Drives the My-Tasks stat tiles + the "complete"/"blocked" behaviour (auto-100%, at-risk, etc.). If blank, inferred from the name |
| `Priority` 🆕 | `Yes`/`No` — does this status feed the **My Priorities** list (default: the `active` bucket) |
