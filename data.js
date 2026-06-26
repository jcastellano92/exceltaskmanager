/* data.js — Office.js data layer.
 * Exposes window.WsjfData. All functions return Promises.
 *
 * Sheet/table contract (see spec):
 *   Tasks / TasksTable           — 27 cols
 *   Goals / GoalsTable
 *   Workstreams / WorkstreamsTable
 *   Subtasks / SubtasksTable
 *   Attachments / AttachmentsTable
 *   ActivityLog / ActivityLogTable
 *   Config sheet: cells K2..N2 hold NextTaskID, NextSubtaskID, NextAttachmentID, NextLogID
 */
(function (root) {
  "use strict";

  const TASKS_TABLE = "TasksTable";
  const GOALS_TABLE = "GoalsTable";
  const WS_TABLE    = "WorkstreamsTable";
  const SUB_TABLE   = "SubtasksTable";
  const ATT_TABLE   = "AttachmentsTable";
  const LOG_TABLE   = "ActivityLogTable";
  const UPD_TABLE   = "UpdatesTable";
  const MS_TABLE    = "MilestonesTable";
  const CFG_SHEET   = "Config";

  // Next-ID is computed as MAX(existing id) + 1 straight from each table. This is
  // immune to the Config "NextID" helper cells losing their formulas or being
  // shifted by added columns (both of which have happened).
  const ID_SOURCE = {
    Task:       { table: TASKS_TABLE, col: "TaskID" },
    Subtask:    { table: SUB_TABLE,   col: "SubtaskID" },
    Attachment: { table: ATT_TABLE,   col: "AttachmentID" },
    Log:        { table: LOG_TABLE,   col: "LogID" },
    Update:     { table: UPD_TABLE,   col: "UpdateID" },
    Milestone:  { table: MS_TABLE,    col: "MilestoneID" }
  };

  // ────────────────────────────────────────────────────────────
  // Low-level: read full table as array of objects by header
  // ────────────────────────────────────────────────────────────
  async function _readTable(tableName) {
    return Excel.run(async (ctx) => {
      const tbl = ctx.workbook.tables.getItem(tableName);
      const headerRange = tbl.getHeaderRowRange();
      const bodyRange   = tbl.getDataBodyRange();
      headerRange.load("values");
      bodyRange.load("values");
      await ctx.sync();

      const headers = headerRange.values[0];
      const rows = bodyRange.values || [];
      return rows.map((row, idx) => {
        const obj = { _rowIndex: idx }; // body-relative
        headers.forEach((h, i) => { obj[h] = row[i]; });
        return obj;
      });
    });
  }

  // ────────────────────────────────────────────────────────────
  // User
  // ────────────────────────────────────────────────────────────
  // Identity is stored per-user in localStorage (no SSO/Graph needed). The user
  // picks their name once; getCurrentUser returns it for stamping + "My Tasks".
  const IDENTITY_KEY = "wsjf.identity";

  function _readStoredIdentity() {
    try {
      if (typeof localStorage === "undefined") return null;
      const raw = localStorage.getItem(IDENTITY_KEY);
      if (!raw) return null;
      const o = JSON.parse(raw);
      return (o && o.name) ? o : null;
    } catch (_) { return null; }
  }

  async function getCurrentUser() {
    const stored = _readStoredIdentity();
    if (stored) {
      return { email: stored.email || "", name: stored.name, initials: initialsFromName(stored.name) };
    }
    // Best-effort fallback (Outlook mailbox isn't available in Excel).
    let email = "", name = "Unknown User";
    try {
      if (Office.context && Office.context.mailbox && Office.context.mailbox.userProfile) {
        email = Office.context.mailbox.userProfile.emailAddress || email;
        name  = Office.context.mailbox.userProfile.displayName || name;
      }
    } catch (_) {}
    return { email: email, name: name, initials: initialsFromName(name) };
  }

  async function setCurrentUser(name, email) {
    name = String(name == null ? "" : name).trim();
    try {
      if (typeof localStorage !== "undefined") {
        if (name) localStorage.setItem(IDENTITY_KEY, JSON.stringify({ name: name, email: email || "" }));
        else localStorage.removeItem(IDENTITY_KEY);
      }
    } catch (_) {}
    const finalName = name || "Unknown User";
    return { email: email || "", name: finalName, initials: initialsFromName(finalName) };
  }

  function initialsFromName(name) {
    if (!name) return "?";
    const parts = String(name).replace(/[,\[\]]/g, " ").trim().split(/\s+/);
    if (parts.length === 1) return parts[0].substring(0, 2).toUpperCase();
    return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
  }

  // ────────────────────────────────────────────────────────────
  // Tasks
  // ────────────────────────────────────────────────────────────
  async function readAllTasks() {
    const all = await _readTable(TASKS_TABLE);
    // Skip phantom rows that have no TaskID (e.g. a stray Description pasted into
    // an empty table row). They can't be opened or saved and would otherwise
    // render as blank cards/rows.
    return all.filter((t) =>
      String(t.TaskID == null ? "" : t.TaskID).trim() !== "" &&
      String(t.Archived || "").toLowerCase() !== "yes");
  }
  async function readArchivedTasks() {
    const all = await _readTable(TASKS_TABLE);
    return all.filter((t) => String(t.Archived || "").toLowerCase() === "yes");
  }
  async function readTaskById(taskId) {
    const all = await _readTable(TASKS_TABLE);
    return all.find((t) => Number(t.TaskID) === Number(taskId)) || null;
  }

  // ────────────────────────────────────────────────────────────
  // Subtasks / Attachments / Activity
  // ────────────────────────────────────────────────────────────
  async function readSubtasksForTask(parentTaskId) {
    const all = await _readTable(SUB_TABLE);
    return all
      .filter((s) => Number(s.ParentTaskID) === Number(parentTaskId))
      .sort((a, b) => (Number(a.Order) || 0) - (Number(b.Order) || 0));
  }
  async function readAttachmentsForTask(parentTaskId) {
    const all = await _readTable(ATT_TABLE);
    return all.filter((a) => Number(a.ParentTaskID) === Number(parentTaskId));
  }
  async function readActivityForTask(parentTaskId, limit) {
    limit = limit || 10;
    const all = await _readTable(LOG_TABLE);
    return all
      .filter((e) => e.EntityType === "Task" && Number(e.EntityID) === Number(parentTaskId))
      .sort((a, b) => String(b.Timestamp).localeCompare(String(a.Timestamp)))
      .slice(0, limit);
  }

  // ────────────────────────────────────────────────────────────
  // Config lists  (OwnersTable / StatusesTable / QuartersTable / TypesTable / YesNoTable)
  // ────────────────────────────────────────────────────────────
  async function readConfigList(tableName) {
    const arr = await _readTable(tableName);
    // first column is the value
    return arr.map((r) => Object.values(r).find((v, idx) => idx > 0 /* skip _rowIndex */));
  }

  // ────────────────────────────────────────────────────────────
  // Next ID helper  (Config sheet K2/L2/M2/N2 contain formulas)
  // ────────────────────────────────────────────────────────────
  async function _readAndBumpNextId(kind) {
    // MAX(existing id) + 1, read directly from the source table.
    const src = ID_SOURCE[kind];
    if (!src) throw new Error("Unknown id kind: " + kind);
    const rows = await _readTable(src.table);
    let max = 0;
    rows.forEach((r) => {
      const n = Number(r[src.col]);
      if (!isNaN(n) && n > max) max = n;
    });
    return max + 1;
  }

  // ────────────────────────────────────────────────────────────
  // Write helpers
  // ────────────────────────────────────────────────────────────
  async function _findRowIndexById(tableName, idColumn, idValue) {
    const all = await _readTable(tableName);
    // Compare as trimmed strings so it works for numeric IDs (TaskID 5) AND
    // string IDs (WorkstreamID "WS09", GoalID "G1"). The previous Number()
    // comparison made NaN === NaN, so string IDs were never found.
    const target = String(idValue == null ? "" : idValue).trim();
    const r = all.find((row) => String(row[idColumn] == null ? "" : row[idColumn]).trim() === target);
    return r ? r._rowIndex : -1;
  }

  function _nowIso() { return new Date().toISOString(); }

  async function _updateRow(tableName, rowIndex, columnHeader, value) {
    return Excel.run(async (ctx) => {
      const tbl = ctx.workbook.tables.getItem(tableName);
      const header = tbl.getHeaderRowRange();
      header.load("values");
      await ctx.sync();
      const colIdx = header.values[0].indexOf(columnHeader);
      if (colIdx < 0) throw new Error(`Column ${columnHeader} not found in ${tableName}`);
      const cell = tbl.getDataBodyRange().getCell(rowIndex, colIdx);
      cell.values = [[value]];
      await ctx.sync();
    });
  }

  // Bulk-update multiple columns in same row in one sync. Faster.
  async function _updateRowMulti(tableName, rowIndex, updates /* {header: value} */) {
    return Excel.run(async (ctx) => {
      const tbl = ctx.workbook.tables.getItem(tableName);
      const header = tbl.getHeaderRowRange();
      header.load("values");
      await ctx.sync();
      const headers = header.values[0];
      const body = tbl.getDataBodyRange();
      Object.keys(updates).forEach((h) => {
        const idx = headers.indexOf(h);
        if (idx < 0) {
          console.warn(`Column ${h} not found in ${tableName} — skipping`);
          return;
        }
        const cell = body.getCell(rowIndex, idx);
        cell.values = [[updates[h]]];
      });
      await ctx.sync();
    });
  }

  async function _deleteRow(tableName, rowIndex) {
    return Excel.run(async (ctx) => {
      const tbl = ctx.workbook.tables.getItem(tableName);
      tbl.rows.getItemAt(rowIndex).delete();
      await ctx.sync();
    });
  }

  // ────────────────────────────────────────────────────────────
  // Task writes
  // ────────────────────────────────────────────────────────────
  // The task headers in declared order (must match sheet):
  const TASK_HEADERS = [
    "TaskID","Title","WorkstreamID","GoalID","Owner","Contributors","StartDate","Quarter","Status","LinkedEpic",
    "UserBusinessValue","TimeCriticality","RiskReduction","CostOfDelay",
    "JobSize","WSJF","PercentComplete","ColumnOrder","DueDate","Tags",
    "BlockedByTaskIDs","Description","Source","Archived","CreatedDate","CreatedBy",
    "LastUpdated","UpdatedBy","AttachmentCount"
  ];

  async function writeTask(taskObj, opts) {
    opts = opts || {};
    const me = await getCurrentUser();
    const rowIndex = await _findRowIndexById(TASKS_TABLE, "TaskID", taskObj.TaskID);
    if (rowIndex < 0) throw new Error("Task not found: " + taskObj.TaskID);

    // Concurrency guard — compare LastUpdated.
    if (!opts.force) {
      const fresh = await readTaskById(taskObj.TaskID);
      if (fresh && taskObj._loadedLastUpdated && fresh.LastUpdated &&
          String(fresh.LastUpdated) > String(taskObj._loadedLastUpdated)) {
        const err = new Error("CONFLICT");
        err.code = "CONFLICT";
        err.serverRow = fresh;
        throw err;
      }
    }

    const ts = _nowIso();
    const updates = {};
    TASK_HEADERS.forEach((h) => {
      if (h === "CostOfDelay" || h === "WSJF" || h === "AttachmentCount") return; // formulas
      if (Object.prototype.hasOwnProperty.call(taskObj, h)) updates[h] = taskObj[h];
    });
    // Optional columns not in the canonical header list (written only if present).
    ["StartDate", "CommitmentID", "Health", "Slips"].forEach((h) => {
      if (Object.prototype.hasOwnProperty.call(taskObj, h)) updates[h] = taskObj[h];
    });
    updates.LastUpdated = ts;
    updates.UpdatedBy   = me.name;

    await _updateRowMulti(TASKS_TABLE, rowIndex, updates);
    // If this save sets the task to Done, close out its subtasks too (best-effort).
    if (String(taskObj.Status) === "Done") {
      try { await closeSubtasksForTask(taskObj.TaskID); }
      catch (e) { console.warn("closeSubtasksForTask failed (non-fatal):", e); }
    }
    // opts.silent suppresses the activity entry (used for derived/secondary writes
    // like syncing % from a subtask toggle, which would otherwise be noise).
    if (!opts.silent) await logActivity("Task", taskObj.TaskID, "Updated", "", "", "", opts.notes || "");
    return ts;
  }

  async function writeTaskStatus(taskId, newStatus, newColumnOrder) {
    const me = await getCurrentUser();
    const rowIndex = await _findRowIndexById(TASKS_TABLE, "TaskID", taskId);
    if (rowIndex < 0) throw new Error("Task not found: " + taskId);

    const prior = await readTaskById(taskId);
    const oldStatus = prior ? prior.Status : "";
    const ts = _nowIso();
    const updates = {
      Status: newStatus,
      ColumnOrder: Number(newColumnOrder) || 0,
      LastUpdated: ts,
      UpdatedBy: me.name
    };
    // Moving a task to Done marks it 100% complete. Reopening it (Done → other)
    // drops it to its subtask-completion ratio, or 90% if it has no subtasks.
    if (String(newStatus) === "Done") {
      updates.PercentComplete = 100;
    } else if (String(oldStatus) === "Done") {
      const subs = await readSubtasksForTask(taskId);
      updates.PercentComplete = subs.length
        ? Math.round(subs.filter((s) => String(s.Done).toLowerCase() === "yes").length / subs.length * 100)
        : 90;
    }
    await _updateRowMulti(TASKS_TABLE, rowIndex, updates);
    // ...and closes out all its subtasks with a completion date (best-effort:
    // a subtask hiccup must not undo the status change the user just made).
    if (String(newStatus) === "Done") {
      try { await closeSubtasksForTask(taskId); }
      catch (e) { console.warn("closeSubtasksForTask failed (non-fatal):", e); }
    }
    // Only log real status changes — not pure within-column reordering.
    if (oldStatus !== newStatus) {
      const action = String(newStatus) === "Done" ? "Completed" : "StatusChanged";
      await logActivity("Task", taskId, action, "Status", oldStatus, newStatus, "");
    }
    return ts;
  }

  async function createTask(taskObj) {
    const me = await getCurrentUser();
    const nextId = await _readAndBumpNextId("Task");
    const ts = _nowIso();
    const today = new Date().toISOString().substring(0, 10);

    // Build an object keyed by header, then append matched to the table's ACTUAL
    // columns. This keeps creation working even if extra columns (StartDate,
    // CommitmentID, …) get added to TasksTable later.
    const obj = {};
    TASK_HEADERS.forEach((h) => {
      switch (h) {
        case "TaskID":          obj[h] = nextId; break;
        case "CostOfDelay":     obj[h] = `=[@UserBusinessValue]+[@TimeCriticality]+[@RiskReduction]`; break;
        case "WSJF":            obj[h] = `=IFERROR([@CostOfDelay]/[@JobSize],0)`; break;
        case "AttachmentCount": obj[h] = `=COUNTIF(AttachmentsTable[ParentTaskID],[@TaskID])`; break;
        case "CreatedDate":     obj[h] = taskObj.CreatedDate || today; break;
        case "CreatedBy":       obj[h] = taskObj.CreatedBy || me.name; break;
        case "LastUpdated":     obj[h] = ts; break;
        case "UpdatedBy":       obj[h] = me.name; break;
        case "Archived":        obj[h] = taskObj.Archived || "No"; break;
        case "Source":          obj[h] = taskObj.Source || "Manual"; break;
        default: obj[h] = Object.prototype.hasOwnProperty.call(taskObj, h) ? (taskObj[h] ?? "") : "";
      }
    });
    // Optional columns not in the canonical header list (written only if the
    // column exists; harmlessly ignored otherwise).
    ["StartDate", "CommitmentID", "Health", "Slips"].forEach((h) => {
      if (Object.prototype.hasOwnProperty.call(taskObj, h)) obj[h] = taskObj[h];
    });

    await _appendObjByHeaders(TASKS_TABLE, obj);
    await logActivity("Task", nextId, "Created", "", "", taskObj.Title || "(untitled)", "");
    return nextId;
  }

  async function archiveTask(taskId) {
    const me = await getCurrentUser();
    const rowIndex = await _findRowIndexById(TASKS_TABLE, "TaskID", taskId);
    if (rowIndex < 0) throw new Error("Task not found: " + taskId);
    const ts = _nowIso();
    await _updateRowMulti(TASKS_TABLE, rowIndex, {
      Archived: "Yes",
      LastUpdated: ts,
      UpdatedBy: me.name
    });
    await logActivity("Task", taskId, "Archived", "Archived", "No", "Yes", "");
    return ts;
  }

  // ────────────────────────────────────────────────────────────
  // Subtask writes
  // ────────────────────────────────────────────────────────────
  const SUB_HEADERS = ["SubtaskID","ParentTaskID","Text","Done","Order","DueDate","CompletedDate","Owner","LastUpdated","UpdatedBy"];

  async function createSubtask(subObj) {
    const me = await getCurrentUser();
    const id = await _readAndBumpNextId("Subtask");
    const ts = _nowIso();
    // Header-driven append (resilient to added columns like CompletedDate/DueDate).
    const obj = {};
    SUB_HEADERS.forEach((h) => {
      switch (h) {
        case "SubtaskID":   obj[h] = id; break;
        case "Done":        obj[h] = subObj.Done || "No"; break;
        case "Order":       obj[h] = subObj.Order != null ? subObj.Order : 999; break;
        case "LastUpdated": obj[h] = ts; break;
        case "UpdatedBy":   obj[h] = me.name; break;
        default: obj[h] = subObj[h] != null ? subObj[h] : "";
      }
    });
    ["CompletedDate", "DueDate", "Owner"].forEach((h) => {
      if (Object.prototype.hasOwnProperty.call(subObj, h)) obj[h] = subObj[h];
    });
    await _appendObjByHeaders(SUB_TABLE, obj);
    await logActivity("Subtask", id, "Created", "", "", subObj.Text || "", `parent=${subObj.ParentTaskID}`);
    return id;
  }
  async function writeSubtask(subObj) {
    const me = await getCurrentUser();
    const rowIndex = await _findRowIndexById(SUB_TABLE, "SubtaskID", subObj.SubtaskID);
    if (rowIndex < 0) throw new Error("Subtask not found: " + subObj.SubtaskID);
    const ts = _nowIso();
    const updates = { LastUpdated: ts, UpdatedBy: me.name };
    SUB_HEADERS.forEach((h) => {
      if (h === "SubtaskID") return;
      if (Object.prototype.hasOwnProperty.call(subObj, h)) updates[h] = subObj[h];
    });
    // CompletedDate isn't in SUB_HEADERS (so creates stay safe), but allow updating
    // it. _updateRowMulti skips it harmlessly if the column doesn't exist yet.
    if (Object.prototype.hasOwnProperty.call(subObj, "CompletedDate")) updates.CompletedDate = subObj.CompletedDate;
    await _updateRowMulti(SUB_TABLE, rowIndex, updates);
    return ts;
  }
  async function deleteSubtask(subtaskId) {
    const rowIndex = await _findRowIndexById(SUB_TABLE, "SubtaskID", subtaskId);
    if (rowIndex < 0) return;
    await _deleteRow(SUB_TABLE, rowIndex);
    await logActivity("Subtask", subtaskId, "Deleted", "", "", "", "");
  }
  async function toggleSubtask(subtaskId, done) {
    const today = new Date().toISOString().substring(0, 10);
    // Stamp a completion date when checked; clear it when unchecked.
    return writeSubtask({ SubtaskID: subtaskId, Done: done ? "Yes" : "No", CompletedDate: done ? today : "" });
  }

  // Close out every open subtask of a task and stamp a completion date.
  // Called when the parent task moves to Done. Only touches subtasks that
  // aren't already done, so it won't overwrite an existing completion date.
  async function closeSubtasksForTask(taskId) {
    const me = await getCurrentUser();
    const subs = await readSubtasksForTask(taskId);
    const today = new Date().toISOString().substring(0, 10);
    const ts = _nowIso();
    for (const s of subs) {
      if (String(s.Done).toLowerCase() !== "yes") {
        await _updateRowMulti(SUB_TABLE, s._rowIndex, {
          Done: "Yes", CompletedDate: today, LastUpdated: ts, UpdatedBy: me.name
        });
      }
    }
  }

  // ────────────────────────────────────────────────────────────
  // Attachments
  // ────────────────────────────────────────────────────────────
  const ATT_HEADERS = ["AttachmentID","ParentTaskID","Label","Url","Type","AddedBy","AddedDate"];

  async function createAttachment(attObj) {
    const me = await getCurrentUser();
    const id = await _readAndBumpNextId("Attachment");
    const ts = new Date().toISOString().substring(0, 10);
    const obj = {};
    ATT_HEADERS.forEach((h) => {
      switch (h) {
        case "AttachmentID": obj[h] = id; break;
        case "AddedBy":      obj[h] = me.name; break;
        case "AddedDate":    obj[h] = ts; break;
        default: obj[h] = attObj[h] != null ? attObj[h] : "";
      }
    });
    await _appendObjByHeaders(ATT_TABLE, obj);
    await logActivity("Attachment", id, "Created", "", "", attObj.Label || "", `parent=${attObj.ParentTaskID}`);
    return id;
  }
  async function writeAttachment(attObj) {
    const rowIndex = await _findRowIndexById(ATT_TABLE, "AttachmentID", attObj.AttachmentID);
    if (rowIndex < 0) throw new Error("Attachment not found: " + attObj.AttachmentID);
    const updates = {};
    ATT_HEADERS.forEach((h) => {
      if (h === "AttachmentID") return;
      if (Object.prototype.hasOwnProperty.call(attObj, h)) updates[h] = attObj[h];
    });
    await _updateRowMulti(ATT_TABLE, rowIndex, updates);
  }
  async function deleteAttachment(attachmentId) {
    const rowIndex = await _findRowIndexById(ATT_TABLE, "AttachmentID", attachmentId);
    if (rowIndex < 0) return;
    await _deleteRow(ATT_TABLE, rowIndex);
    await logActivity("Attachment", attachmentId, "Deleted", "", "", "", "");
  }

  // ────────────────────────────────────────────────────────────
  // Updates — manual, chronological notes on a Task or Subtask.
  // Requires an UpdatesTable: UpdateID, ParentType, ParentID, Text, AddedBy, AddedDate
  // ────────────────────────────────────────────────────────────
  async function readUpdatesForParent(parentType, parentId) {
    const all = await _readTable(UPD_TABLE);
    return all
      .filter((u) => String(u.ParentType) === String(parentType) && Number(u.ParentID) === Number(parentId))
      .sort((a, b) => String(a.AddedDate).localeCompare(String(b.AddedDate)));   // oldest → newest
  }
  async function createUpdate(obj) {
    const me = await getCurrentUser();
    const id = await _readAndBumpNextId("Update");
    await _appendObjByHeaders(UPD_TABLE, {
      UpdateID: id,
      ParentType: obj.ParentType || "Task",
      ParentID: obj.ParentID,
      Text: obj.Text || "",
      AddedBy: me.name,
      AddedDate: _nowIso()
    });
    return id;
  }

  // ────────────────────────────────────────────────────────────
  // Activity Log
  // ────────────────────────────────────────────────────────────
  const LOG_HEADERS = ["LogID","Timestamp","EntityType","EntityID","Action","ChangedBy","FieldChanged","OldValue","NewValue","Notes"];

  async function logActivity(entityType, entityId, action, fieldChanged, oldValue, newValue, notes) {
    try {
      const me = await getCurrentUser();
      const id = await _readAndBumpNextId("Log");
      await _appendObjByHeaders(LOG_TABLE, {
        LogID: id,
        Timestamp: _nowIso(),
        EntityType: entityType || "",
        EntityID: entityId || 0,
        Action: action || "",
        ChangedBy: me.name,
        FieldChanged: fieldChanged || "",
        OldValue: oldValue == null ? "" : String(oldValue),
        NewValue: newValue == null ? "" : String(newValue),
        Notes: notes || ""
      });
    } catch (err) {
      console.warn("logActivity failed (non-fatal):", err);
    }
  }

  // ────────────────────────────────────────────────────────────
  // Config / dimension WRITES  (admin "Config" screen)
  // ────────────────────────────────────────────────────────────
  function _firstColValue(rowObj) {
    const keys = Object.keys(rowObj).filter((k) => k !== "_rowIndex");
    return rowObj[keys[0]];
  }

  // Append a row built from an object, matched to the table's headers.
  async function _appendObjByHeaders(tableName, obj) {
    return Excel.run(async (ctx) => {
      const tbl = ctx.workbook.tables.getItem(tableName);
      const header = tbl.getHeaderRowRange();
      header.load("values");
      await ctx.sync();
      const headers = header.values[0];
      const row = headers.map((h) => (Object.prototype.hasOwnProperty.call(obj, h) ? obj[h] : ""));
      tbl.rows.add(null, [row]);
      await ctx.sync();
    });
  }

  // ----- single-value config lists (Owners / Quarters / Statuses / Types) -----
  async function addConfigValue(tableName, value) {
    value = String(value == null ? "" : value).trim();
    if (!value) throw new Error("Value required");
    const existing = await readConfigList(tableName);
    if (existing.map((v) => String(v).toLowerCase()).includes(value.toLowerCase())) {
      throw new Error('"' + value + '" already exists');
    }
    return Excel.run(async (ctx) => {
      const tbl = ctx.workbook.tables.getItem(tableName);
      const header = tbl.getHeaderRowRange();
      header.load("values");
      await ctx.sync();
      const width = header.values[0].length;
      const row = new Array(width).fill("");
      row[0] = value;
      tbl.rows.add(null, [row]);
      await ctx.sync();
    });
  }

  async function renameConfigValue(tableName, oldValue, newValue) {
    oldValue = String(oldValue);
    newValue = String(newValue == null ? "" : newValue).trim();
    if (!newValue) throw new Error("New value required");
    const all = await _readTable(tableName);
    const row = all.find((r) => String(_firstColValue(r)) === oldValue);
    if (!row) throw new Error('"' + oldValue + '" not found');
    const firstHeader = Object.keys(row).filter((k) => k !== "_rowIndex")[0];
    await _updateRow(tableName, row._rowIndex, firstHeader, newValue);
  }

  async function deleteConfigValue(tableName, value) {
    const all = await _readTable(tableName);
    const row = all.find((r) => String(_firstColValue(r)) === String(value));
    if (row) await _deleteRow(tableName, row._rowIndex);
  }

  // ----- usage counts (guard before delete) -----
  async function countTasksByField(field, value, multi) {
    const tasks = await _readTable(TASKS_TABLE);
    let n = 0;
    tasks.forEach((t) => {
      const cur = String(t[field] || "");
      if (multi) {
        if (cur.split(/[;,]/).map((s) => s.trim()).includes(String(value))) n++;
      } else if (cur === String(value)) n++;
    });
    return n;
  }
  async function countTasksByWorkstream(id) { return countTasksByField("WorkstreamID", id, false); }
  async function countTasksByGoal(id) { return countTasksByField("GoalID", id, true); }

  // ----- cascade a renamed value into TasksTable -----
  async function _cascadeTaskField(field, oldValue, newValue, multi) {
    const tasks = await _readTable(TASKS_TABLE);
    let changed = 0;
    for (const t of tasks) {
      const cur = String(t[field] || "");
      let next = null;
      if (multi) {
        const parts = cur.split(";").map((s) => s.trim()).filter(Boolean);
        if (!parts.includes(String(oldValue))) continue;
        next = parts.map((p) => (p === String(oldValue) ? newValue : p)).join("; ");
      } else {
        if (cur !== String(oldValue)) continue;
        next = newValue;
      }
      if (next !== null && next !== cur) {
        await _updateRow(TASKS_TABLE, t._rowIndex, field, next);
        changed++;
      }
    }
    return changed;
  }

  // ----- Owners / Quarters: rename cascades to TasksTable -----
  async function renameOwner(oldName, newName) {
    await renameConfigValue("OwnersTable", oldName, newName);
    return _cascadeTaskField("Owner", oldName, String(newName).trim(), true);
  }
  async function renameQuarter(oldQ, newQ) {
    await renameConfigValue("QuartersTable", oldQ, newQ);
    return _cascadeTaskField("Quarter", oldQ, String(newQ).trim(), false);
  }

  // ----- Workstreams / Goals (tasks link by ID; renaming Name auto-propagates) -----
  // These log meaningful create/update/delete activity (not trivial UI actions).
  async function createWorkstream(obj) {
    await _appendObjByHeaders(WS_TABLE, obj);
    await logActivity("Workstream", obj.WorkstreamID, "Created", "", "", obj.Name || obj.WorkstreamID, "");
  }
  async function updateWorkstream(id, fields) {
    const idx = await _findRowIndexById(WS_TABLE, "WorkstreamID", id);
    if (idx < 0) throw new Error("Workstream not found: " + id);
    await _updateRowMulti(WS_TABLE, idx, fields);
    await logActivity("Workstream", id, "Updated", "", "", fields.Name || "", "");
  }
  async function deleteWorkstream(id) {
    const idx = await _findRowIndexById(WS_TABLE, "WorkstreamID", id);
    if (idx >= 0) { await _deleteRow(WS_TABLE, idx); await logActivity("Workstream", id, "Deleted", "", "", "", ""); }
  }
  async function createGoal(obj) {
    await _appendObjByHeaders(GOALS_TABLE, obj);
    await logActivity("Goal", obj.GoalID, "Created", "", "", obj.ShortName || obj.GoalName || obj.GoalID, "");
  }
  async function updateGoal(id, fields) {
    const idx = await _findRowIndexById(GOALS_TABLE, "GoalID", id);
    if (idx < 0) throw new Error("Goal not found: " + id);
    await _updateRowMulti(GOALS_TABLE, idx, fields);
    await logActivity("Goal", id, "Updated", "", "", fields.ShortName || fields.GoalName || "", "");
  }
  async function deleteGoal(id) {
    const idx = await _findRowIndexById(GOALS_TABLE, "GoalID", id);
    if (idx >= 0) { await _deleteRow(GOALS_TABLE, idx); await logActivity("Goal", id, "Deleted", "", "", "", ""); }
  }

  // ----- Milestones (standalone dated lines on the roadmap — NOT tasks) -----
  // Requires a MilestonesTable: MilestoneID, Title, Date, Quarter, GoalID, Color, Notes.
  async function readMilestones() {
    return _readTable(MS_TABLE);
  }
  async function createMilestone(obj) {
    const id = await _readAndBumpNextId("Milestone");
    await _appendObjByHeaders(MS_TABLE, {
      MilestoneID: id,
      Title: obj.Title || "",
      Date: obj.Date || "",
      Quarter: obj.Quarter || "",
      GoalID: obj.GoalID || "",
      Color: obj.Color || "",
      Notes: obj.Notes || ""
    });
    await logActivity("Milestone", id, "Created", "", "", obj.Title || "", "");
    return id;
  }
  async function updateMilestone(id, fields) {
    const idx = await _findRowIndexById(MS_TABLE, "MilestoneID", id);
    if (idx < 0) throw new Error("Milestone not found: " + id);
    await _updateRowMulti(MS_TABLE, idx, fields);
    await logActivity("Milestone", id, "Updated", "", "", fields.Title || "", "");
  }
  async function deleteMilestone(id) {
    const idx = await _findRowIndexById(MS_TABLE, "MilestoneID", id);
    if (idx >= 0) { await _deleteRow(MS_TABLE, idx); await logActivity("Milestone", id, "Deleted", "", "", "", ""); }
  }

  // ────────────────────────────────────────────────────────────
  // Public API
  // ────────────────────────────────────────────────────────────
  root.WsjfData = {
    // user
    getCurrentUser, setCurrentUser,
    // tasks
    readAllTasks, readArchivedTasks, readTaskById,
    writeTask, writeTaskStatus, createTask, archiveTask,
    // children
    readSubtasksForTask, readAttachmentsForTask, readActivityForTask,
    writeSubtask, createSubtask, deleteSubtask, toggleSubtask,
    writeAttachment, createAttachment, deleteAttachment,
    // updates
    readUpdatesForParent, createUpdate,
    // config (read)
    readConfigList,
    // config admin (write)
    addConfigValue, renameConfigValue, deleteConfigValue,
    countTasksByField, countTasksByWorkstream, countTasksByGoal,
    renameOwner, renameQuarter,
    createWorkstream, updateWorkstream, deleteWorkstream,
    createGoal, updateGoal, deleteGoal,
    // milestones (standalone roadmap lines)
    readMilestones, createMilestone, updateMilestone, deleteMilestone,
    // activity
    logActivity,
    // expose for tests
    _internal: { _readTable }
  };
})(window);