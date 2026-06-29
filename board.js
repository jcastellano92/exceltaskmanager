/* board.js — Kanban dialog logic */
(function () {
  "use strict";

  const COLUMNS = ["Backlog", "In Progress", "On Track", "Blocked", "Done"];   // default if StatusesTable is empty
  const POLL_MS = 30000;

  // Board columns are driven by the StatusesTable (editable in Config), with an
  // optional per-user column order remembered in localStorage. Falls back to the
  // built-in defaults when the table is empty.
  function boardColumns() {
    let base = (State.config.Statuses && State.config.Statuses.length) ? State.config.Statuses.slice() : COLUMNS.slice();
    try {
      const saved = JSON.parse(lsGet("boardColOrder", "[]"));
      if (saved && saved.length) {
        base = saved.filter((s) => base.includes(s)).concat(base.filter((s) => !saved.includes(s)));
      }
    } catch (e) {}
    return base;
  }
  function saveBoardColOrder(order) { lsSet("boardColOrder", JSON.stringify(order)); }

  // ───── State ─────
  const State = {
    me: { name: "Unknown", email: "", initials: "?" },
    tasks: [],
    subtasksByParent: {},
    attsByParent: {},
    activityByParent: {},
    config: { Owners: [], Statuses: [], Quarters: [], Types: [], YesNo: [] },
    goals: [],
    workstreams: [],
    view: "all",                         // always show all tasks; narrow with the filters
    viewMode: "board",                   // "board" | "list" | "mine" | "ws" | "config"
    wsQuarter: "",                       // focused quarter on the Workstreams dashboard
    quarterDates: {},                    // { Q1: {start, end}, … } from QuartersTable
    capacity: {},                        // { person: {availability, baseline} } from CapacityTable
    expandedSubtasks: new Set(),         // taskIds whose subtask checklist is expanded inline (board + list)
    milestones: [],                      // standalone roadmap milestone lines (MilestonesTable)
    allUpdates: [],                      // every Updates row (for the My Tasks "recent updates" feed)
    roadmap: { showDates: true, expanded: new Set() }, // roadmap prefs: drag date tooltip + expanded subtask lanes
    selected: new Set(),                 // taskIds selected in List view for bulk actions
    sort: "wsjf",                        // board column sort: "wsjf" | "due" | "manual"
    groupBy: "",                         // List view grouping: "" | WorkstreamID | Owner | Quarter | Status | GoalID
    listSort: { key: "WSJF", dir: "desc" }, // list view column sort
    listColFilters: {},                  // per-column filters (Sheets-style): key -> Set of allowed display values
    filters: { owner: "", workstream: "", quarter: "", status: "", goal: "", health: "", subtasks: "",
               startFrom: "", startTo: "", dueFrom: "", dueTo: "", tags: new Set(), search: "" },
    selectedTags: new Set(),
    modalOpen: false,
    dragging: false,                     // a card is being dragged right now
    pendingWrites: 0,                    // in-flight Excel writes (poll defers while > 0)
    lastWriteTs: 0,                      // ms timestamp of the last write settling
    lastSyncTs: null,
    pollTimer: null,
    sortableInstances: []
  };

  Office.onReady(async () => {
    try {
      if (/[?&]host=dialog/.test(location.search)) {
        // Running inside the full-screen dialog, which can't call Excel.run.
        // Route all data access through the task pane via message-passing.
        window.WsjfData = makeRemoteWsjf();
      }
      await bootstrap();
      bindUi();
      await ensureIdentity();   // first-run gate: must pick/add a name
      render();
      startPolling();
    } catch (err) {
      console.error("Bootstrap failed", err);
      showFatalError(err);
    }
  });

  // Remote data proxy used ONLY in dialog mode: mirrors window.WsjfData but
  // forwards every call to the task pane (which has a working Excel session),
  // since Office dialogs cannot run Excel.run directly.
  function makeRemoteWsjf() {
    let seq = 0;
    const pending = {};
    Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, (arg) => {
      let msg;
      try { msg = JSON.parse(arg.message); } catch (_) { return; }
      const p = pending[msg.id];
      if (!p) return;
      delete pending[msg.id];
      if (msg.error) {
        const e = new Error(msg.error.message || "Remote error");
        if (msg.error.code) e.code = msg.error.code;
        if (msg.error.serverRow) e.serverRow = msg.error.serverRow;
        p.reject(e);
      } else {
        p.resolve(msg.result);
      }
    });
    function call(method, args) {
      return new Promise((resolve, reject) => {
        const id = ++seq;
        const timer = setTimeout(() => {
          if (pending[id]) {
            delete pending[id];
            reject(new Error("Lost connection to the Product Management Tool sidebar. Reopen it from the “Product Mgmt” button on the Home ribbon, then try again."));
          }
        }, 20000);
        pending[id] = {
          resolve: function (v) { clearTimeout(timer); resolve(v); },
          reject: function (e) { clearTimeout(timer); reject(e); }
        };
        Office.context.ui.messageParent(JSON.stringify({ id: id, method: method, args: args || [] }));
      });
    }
    const methods = [
      "getCurrentUser", "setCurrentUser", "readAllTasks", "readArchivedTasks", "readTaskById",
      "writeTask", "writeTaskStatus", "createTask", "archiveTask",
      "readSubtasksForTask", "readAttachmentsForTask", "readActivityForTask",
      "readUpdatesForParent", "createUpdate",
      "writeSubtask", "createSubtask", "deleteSubtask", "toggleSubtask",
      "writeAttachment", "createAttachment", "deleteAttachment",
      "readConfigList", "addConfigValue", "renameConfigValue", "deleteConfigValue",
      "countTasksByField", "countTasksByWorkstream", "countTasksByGoal",
      "renameOwner", "renameQuarter", "renameStatus",
      "createWorkstream", "updateWorkstream", "deleteWorkstream",
      "createGoal", "updateGoal", "deleteGoal", "logActivity",
      "createMilestone", "updateMilestone", "deleteMilestone"
    ];
    const api = {};
    methods.forEach((m) => {
      api[m] = function () { return call(m, Array.prototype.slice.call(arguments)); };
    });
    api._internal = { _readTable: function (name) { return call("_readTable", [name]); } };
    return api;
  }

  // Serialize every mutating data call into a single queue. Firing concurrent
  // Excel.run writes (e.g. dragging fast, or a write landing on top of a poll)
  // can overload the Excel API and return 500s — which is what makes a dropped
  // card "pop back" to where it was. One-write-at-a-time fixes that, and the
  // pendingWrites/lastWriteTs counters let polling defer while writes are active.
  function installWriteQueue() {
    const mutating = [
      "writeTask", "writeTaskStatus", "createTask", "archiveTask",
      "writeSubtask", "createSubtask", "deleteSubtask", "toggleSubtask",
      "writeAttachment", "createAttachment", "deleteAttachment",
      "addConfigValue", "renameConfigValue", "deleteConfigValue",
      "renameOwner", "renameQuarter", "renameStatus",
      "createWorkstream", "updateWorkstream", "deleteWorkstream",
      "createGoal", "updateGoal", "deleteGoal",
      "createMilestone", "updateMilestone", "deleteMilestone"
    ];
    let chain = Promise.resolve();
    mutating.forEach((m) => {
      const orig = window.WsjfData[m];
      if (typeof orig !== "function") return;
      window.WsjfData[m] = function () {
        const args = arguments;
        State.pendingWrites++;
        const run = chain.then(() => orig.apply(window.WsjfData, args));
        chain = run.then(function () {}, function () {});   // keep the queue alive on error
        run.then(
          function () { State.pendingWrites--; State.lastWriteTs = Date.now(); },
          function () { State.pendingWrites--; State.lastWriteTs = Date.now(); }
        );
        return run;
      };
    });
  }

  // ───── Bootstrap ─────
  async function bootstrap() {
    installWriteQueue();
    State.me = await window.WsjfData.getCurrentUser();
    // Read sequentially (not Promise.all) — firing many concurrent Excel.run
    // calls at once can overload the Excel API service and return 500 / internal errors.
    await loadConfigAndDimensions();
    await reloadTasks();
    validateSchema();
    render();
  }

  // Verify the workbook's columns still line up with what the app reads/writes,
  // so a renamed/removed column surfaces immediately instead of breaking silently.
  // Uses already-loaded rows (no extra Excel reads); skips empty tables.
  function validateSchema() {
    const anySub = Object.values(State.subtasksByParent).find((a) => a && a.length);
    const checks = [
      { name: "TasksTable", sample: State.tasks[0],
        critical: ["TaskID", "Title", "Status", "Owner", "WorkstreamID", "Quarter", "WSJF", "PercentComplete", "DueDate"],
        optional: ["StartDate", "Contributors", "GoalID", "Tags", "Health", "Slips"] },
      { name: "WorkstreamsTable", sample: State.workstreams[0],
        critical: ["WorkstreamID", "Name"], optional: ["Owner", "Status", "Goals", "Quarters", "Metric1"] },
      { name: "GoalsTable", sample: State.goals[0], critical: ["GoalID"], optional: ["ShortName", "GoalName"] },
      { name: "SubtasksTable", sample: anySub && anySub[0],
        critical: ["SubtaskID", "ParentTaskID", "Text", "Done"], optional: ["Order", "DueDate", "CompletedDate", "Owner"] }
    ];
    const blocking = [];
    checks.forEach((c) => {
      if (!c.sample) return;                         // empty table — nothing to verify
      const keys = Object.keys(c.sample);
      const missCrit = c.critical.filter((k) => !keys.includes(k));
      const missOpt = (c.optional || []).filter((k) => !keys.includes(k));
      if (missOpt.length) console.warn(`[Schema] ${c.name} optional column(s) not present: ${missOpt.join(", ")}`);
      if (missCrit.length) blocking.push(`${c.name}: ${missCrit.join(", ")}`);
    });
    if (blocking.length) {
      console.error("[Schema] Missing required column(s) — " + blocking.join(" | "));
      toast("⚠ Workbook schema mismatch — missing " + blocking.join("; ") + ". Some features may not work.", "error");
    } else {
      console.log("[Schema] OK — workbook columns match the app.");
    }
  }

  async function loadConfigAndDimensions() {
    // Sequential reads (one Excel.run at a time) to avoid overloading the API service.
    const owners      = await window.WsjfData.readConfigList("OwnersTable");
    const statuses    = await window.WsjfData.readConfigList("StatusesTable");
    const quarters    = await window.WsjfData.readConfigList("QuartersTable");
    const types       = await window.WsjfData.readConfigList("TypesTable");
    const yesNo       = await window.WsjfData.readConfigList("YesNoTable");
    const goals       = await window.WsjfData._internal._readTable("GoalsTable");
    const workstreams = await window.WsjfData._internal._readTable("WorkstreamsTable");
    State.config = {
      Owners: owners.filter(Boolean),
      Statuses: statuses.filter(Boolean),
      Quarters: quarters.filter(Boolean),
      Types: types.filter(Boolean),
      YesNo: yesNo.filter(Boolean)
    };
    State.goals = goals;
    State.workstreams = workstreams;

    // Quarter start/end dates (if the QuartersTable carries them) drive
    // "days left" and the new-task due-date default instead of hardcoded ones.
    try {
      const qRows = await window.WsjfData._internal._readTable("QuartersTable");
      const qd = {};
      qRows.forEach((r) => { if (r.Quarter) qd[r.Quarter] = { start: r.StartDate, end: r.EndDate }; });
      State.quarterDates = qd;
    } catch (e) { console.warn("QuartersTable dates unavailable:", e); }

    // Standalone milestones (dated lines on the roadmap; optional table).
    try {
      State.milestones = await window.WsjfData._internal._readTable("MilestonesTable");
    } catch (e) { State.milestones = []; }

    // Per-person capacity (availability dial + optional baseline override).
    try {
      const capRows = await window.WsjfData._internal._readTable("CapacityTable");
      const cap = {};
      const numOrNull = (v) => (v === "" || v == null ? null : Number(v));  // "" → fall back; 0 stays 0
      capRows.forEach((r) => {
        if (!r.Person) return;
        cap[r.Person] = {
          availability: r.Availability === "" || r.Availability == null ? 1 : Number(r.Availability),
          byQuarter: { Q1: numOrNull(r.AvailQ1), Q2: numOrNull(r.AvailQ2), Q3: numOrNull(r.AvailQ3), Q4: numOrNull(r.AvailQ4) },
          baseline: numOrNull(r.BaselineOverride)
        };
      });
      State.capacity = cap;
    } catch (e) { State.capacity = {}; }

    fillSelect("filter-owner", ["", ...State.config.Owners], "All Owners");
    fillSelect("filter-workstream",
      [{ value: "", label: "All Workstreams" }].concat(
        State.workstreams.map((w) => ({ value: w.WorkstreamID, label: w.Name || w.WorkstreamID }))
      )
    );
    fillSelect("filter-quarter", ["", ...State.config.Quarters], "All Quarters");
    fillSelect("filter-status", ["", ...boardColumns()], "All Statuses", State.filters.status);
    fillSelect("filter-goal",
      [{ value: "", label: "All Goals" }].concat(
        State.goals.map((g) => ({ value: g.GoalID, label: g.ShortName || g.GoalName || g.GoalID }))
      )
    );
  }

  async function reloadTasks() {
    State.tasks = await window.WsjfData.readAllTasks();
    await loadAllSubtasks();
    await loadAllUpdates();
    State.lastSyncTs = Date.now();
    updateSyncLabel();
    refreshTagFilterMenu();
  }

  async function loadAllUpdates() {
    try { State.allUpdates = await window.WsjfData._internal._readTable("UpdatesTable"); }
    catch (e) { State.allUpdates = []; }
  }

  async function loadAllSubtasks() {
    try {
      const all = await window.WsjfData._internal._readTable("SubtasksTable");
      const byParent = {};
      all.forEach((s) => {
        const pid = Number(s.ParentTaskID);
        (byParent[pid] = byParent[pid] || []).push(s);
      });
      State.subtasksByParent = byParent;
    } catch (e) { console.warn("loadAllSubtasks failed", e); }
  }

  // ───── Render ─────
  function render() {
    const board = document.getElementById("kanban");
    const list = document.getElementById("listview");
    const mineEl = document.getElementById("mineview");
    const wsEl = document.getElementById("wsview");
    const config = document.getElementById("configview");
    const mid = document.querySelector(".topbar-middle");
    const addBtn = document.getElementById("add-task-btn");
    const sortToggle = document.getElementById("sort-toggle");
    const af = document.getElementById("active-filters");
    if (af) af.hidden = true;   // only task views (board/list) show filter chips

    updateIdentityChip();
    const vt = document.getElementById("view-mode-toggle");
    if (vt) Array.from(vt.querySelectorAll("button")).forEach((b) => b.classList.toggle("active", b.dataset.mode === State.viewMode));
    const cfgBtn = document.getElementById("config-btn");
    if (cfgBtn) cfgBtn.classList.toggle("active", State.viewMode === "config");

    const roadmapEl = document.getElementById("roadmapview");
    board.hidden = true;
    list.hidden = true;
    if (mineEl) mineEl.hidden = true;
    if (wsEl) wsEl.hidden = true;
    if (roadmapEl) roadmapEl.hidden = true;
    if (config) config.hidden = true;

    if (State.viewMode === "config") {
      if (config) config.hidden = false;
      if (mid) mid.hidden = true;       // filters/search/sort are task-only
      if (addBtn) addBtn.hidden = true;
      renderConfig();
      return;
    }
    if (State.viewMode === "mine") {
      if (mineEl) mineEl.hidden = false;
      if (mid) mid.hidden = true;       // dashboard is self-contained, no global filters
      if (addBtn) addBtn.hidden = false;
      if (sortToggle) sortToggle.hidden = true;
      renderMyDashboard();
      return;
    }
    if (State.viewMode === "ws") {
      if (wsEl) wsEl.hidden = false;
      if (mid) mid.hidden = true;
      if (addBtn) addBtn.hidden = true;
      if (sortToggle) sortToggle.hidden = true;
      renderWorkstreams();
      return;
    }

    if (mid) mid.hidden = false;
    if (addBtn) addBtn.hidden = false;
    renderActiveFilters();

    const sortRow = sortToggle ? sortToggle.closest(".fp-row") : null;
    if (State.viewMode === "roadmap") {
      if (roadmapEl) roadmapEl.hidden = false;
      if (sortRow) sortRow.hidden = true;
      renderRoadmap();
      return;
    }
    if (State.viewMode === "list") {
      list.hidden = false;
      if (sortRow) sortRow.hidden = true;     // list has its own column sort
      renderList();
      return;
    }

    board.hidden = false;
    if (sortRow) sortRow.hidden = false;
    renderBoard();
  }

  function renderBoard() {
    const board = document.getElementById("kanban");
    board.innerHTML = "";

    const visible = computeVisibleTasks();

    boardColumns().forEach((status) => {
      const col = document.createElement("section");
      col.className = "kcol kcol-" + statusSlug(status);
      col.dataset.status = status;

      const tasksInCol = visible
        .filter((t) => t.Status === status)
        .sort(sortComparator);

      col.innerHTML = `
        <header class="kcol-head">
          <span class="kcol-title">${escapeHtml(status)}</span>
          <span class="kcol-count">(${tasksInCol.length})</span>
        </header>
        <div class="kcol-body" data-status="${escapeHtml(status)}"></div>
      `;
      const body = col.querySelector(".kcol-body");

      if (tasksInCol.length === 0) {
        const empty = document.createElement("div");
        empty.className = "kcol-empty";
        empty.textContent = "Drop here";
        body.appendChild(empty);
      } else {
        tasksInCol.forEach((t) => body.appendChild(renderCard(t)));
      }
      board.appendChild(col);
    });

    initSortable();
  }

  function renderCard(task) {
    const el = document.createElement("article");
    el.className = "kcard";
    el.dataset.taskId = task.TaskID;

    const wsjf = Number(task.WSJF) || 0;
    const pct  = Math.max(0, Math.min(100, Number(task.PercentComplete) || 0));
    const wsjfClass = wsjfPillClass(wsjf);

    const ownerHtml = peopleAvatars(task, 4);

    const wsName = workstreamName(task.WorkstreamID);
    const wsColor = colorHash(task.WorkstreamID || "");
    const tags = String(task.Tags || "").split(";").map((s) => s.trim()).filter(Boolean);

    const dueLabel = task.DueDate ? formatDateShort(task.DueDate) : "";
    const subCount = (State.subtasksByParent[task.TaskID] || []).length;
    const subDone  = (State.subtasksByParent[task.TaskID] || []).filter((s) => String(s.Done).toLowerCase() === "yes").length;
    const attCount = Number(task.AttachmentCount) || (State.attsByParent[task.TaskID] || []).length;
    const health = task.Health || "";
    const healthDot = health ? `<span class="health-dot health-${statusSlug(health)}" title="${escapeAttr(health)}"></span>` : "";
    if (health) el.classList.add("khealth-" + statusSlug(health));   // left-edge accent

    el.innerHTML = `
      <div class="kcard-head">
        <span class="kcard-title" title="${escapeAttr(task.Title)}">${healthDot}${scheduleChip(task)}${escapeHtml(task.Title || "")}</span>
        <span class="wsjf-pill ${wsjfClass}">${wsjf.toFixed(1)}</span>
      </div>`;
    el.innerHTML += `
      <div class="kcard-meta">
        <div class="avatars">${ownerHtml}</div>
        <span class="ws-pill" style="background:${wsColor}">${escapeHtml(wsName)}</span>
        ${goalDotsHtml(task)}
        <span class="q-badge">${escapeHtml(task.Quarter || "")}</span>
      </div>
      <div class="kcard-foot">
        ${dueLabel ? `<span class="due">📅 ${escapeHtml(dueLabel)}</span>` : ""}
        <div class="progress"><div class="progress-fill" style="width:${pct}%"></div></div>
        <span class="pct">${pct}%</span>
        ${attCount > 0 ? `<span class="att">📎 ${attCount}</span>` : ""}
        ${subCount > 0 ? `<button class="sub-toggle" title="Show subtasks">✓ ${subDone}/${subCount} ${State.expandedSubtasks.has(Number(task.TaskID)) ? "▾" : "▸"}</button>` : ""}
      </div>
      ${tags.length ? `<div class="kcard-tags">${tags.map((t) => `<span class="tag-pill">${escapeHtml(t)}</span>`).join("")}</div>` : ""}
    `;

    el.addEventListener("click", (e) => {
      if (e.target.closest(".kcard-drag-handle")) return;
      openEditModal(task.TaskID);
    });

    const subToggle = el.querySelector(".sub-toggle");
    if (subToggle) subToggle.addEventListener("click", (e) => {
      e.stopPropagation();
      const tid = Number(task.TaskID);
      if (State.expandedSubtasks.has(tid)) State.expandedSubtasks.delete(tid);
      else State.expandedSubtasks.add(tid);
      render();
    });
    if (State.expandedSubtasks.has(Number(task.TaskID))) {
      el.appendChild(buildSubtaskAccordion(Number(task.TaskID)));
    }

    return el;
  }

  // Inline subtask checklist used by both Board cards and List rows.
  function buildSubtaskAccordion(taskId) {
    const wrap = document.createElement("div");
    wrap.className = "subtask-accordion";
    const subs = (State.subtasksByParent[taskId] || []).slice()
      .sort((a, b) => (Number(a.Order) || 0) - (Number(b.Order) || 0));
    if (!subs.length) { wrap.innerHTML = '<div class="sa-empty">No subtasks</div>'; return wrap; }
    const doneN = subs.filter((s) => String(s.Done).toLowerCase() === "yes").length;
    const head = document.createElement("div");
    head.className = "sa-head";
    head.innerHTML = "Subtasks <span class=\"sa-head-count\">" + doneN + "/" + subs.length + "</span>";
    wrap.appendChild(head);
    subs.forEach((s) => {
      const row = document.createElement("label");
      row.className = "sa-row";
      row.addEventListener("click", (e) => e.stopPropagation());
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = String(s.Done).toLowerCase() === "yes";
      cb.addEventListener("click", (e) => e.stopPropagation());
      cb.addEventListener("change", async () => {
        const newVal = cb.checked;
        s.Done = newVal ? "Yes" : "No";
        try {
          await window.WsjfData.toggleSubtask(s.SubtaskID, newVal);
          await syncParentPctFromSubtasks(taskId);
          render();
        } catch (err) {
          cb.checked = !newVal; s.Done = !newVal ? "Yes" : "No";
          toast("Subtask save failed: " + err.message, "error");
        }
      });
      const span = document.createElement("span");
      span.className = "sa-text" + (cb.checked ? " done" : "");
      span.textContent = s.Text || "";
      span.title = s.Text || "";   // full text on hover (board cards truncate)
      row.appendChild(cb);
      row.appendChild(span);

      // Owner / due / completed metadata (display only here; edit in the modal).
      const meta = document.createElement("span");
      meta.className = "sa-meta";
      const bits = [];
      if (s.Owner) bits.push('<span class="sa-avatar" style="background:' + colorHash(s.Owner) + '" title="' + escapeAttr(s.Owner) + '">' + initialsFromName(s.Owner) + '</span>');
      if (s.DueDate) {
        const overdue = !cb.checked && isoDate(s.DueDate) && isoDate(s.DueDate) < new Date().toISOString().slice(0, 10);
        bits.push('<span class="sa-due' + (overdue ? " overdue" : "") + '">📅 ' + escapeHtml(formatDateShort(s.DueDate)) + '</span>');
      }
      if (cb.checked && s.CompletedDate) bits.push('<span class="sa-done">✓ ' + escapeHtml(formatDateShort(s.CompletedDate)) + '</span>');
      meta.innerHTML = bits.join("");
      row.appendChild(meta);
      wrap.appendChild(row);
    });
    return wrap;
  }

  // % complete derived from a task's subtasks (e.g. 4 of 5 done = 80). null if none.
  function subtaskRatio(taskId) {
    const subs = State.subtasksByParent[taskId] || [];
    if (!subs.length) return null;
    const done = subs.filter((s) => String(s.Done).toLowerCase() === "yes").length;
    return Math.round((done / subs.length) * 100);
  }

  // Keep a task's % in step with its subtasks (used after inline subtask toggles).
  // Doesn't touch a Done task — that stays 100 until a user moves it.
  async function syncParentPctFromSubtasks(taskId) {
    const task = State.tasks.find((t) => Number(t.TaskID) === Number(taskId));
    if (!task || task.Status === "Done") return;
    const ratio = subtaskRatio(taskId);
    if (ratio == null || Number(task.PercentComplete) === ratio) return;
    task.PercentComplete = ratio;
    try { await window.WsjfData.writeTask({ TaskID: task.TaskID, PercentComplete: ratio }, { force: true, silent: true }); }
    catch (e) { console.warn("syncParentPctFromSubtasks failed:", e); }
  }

  function initSortable() {
    // Tear down previous instances
    State.sortableInstances.forEach((s) => { try { s.destroy(); } catch (_) {} });
    State.sortableInstances = [];

    if (typeof Sortable === "undefined") {   // library blocked/failed to load — don't break the board
      console.warn("SortableJS not available — drag-and-drop disabled.");
      toast("Drag-and-drop unavailable (Sortable failed to load). Status changes still work via the card.", "warn");
      return;
    }
    const manual = State.sort === "manual";
    document.querySelectorAll(".kcol-body").forEach((body) => {
      const s = Sortable.create(body, {
        group: "wsjf-board",
        animation: 150,
        ghostClass: "kcard-ghost",
        filter: "input, button, a, .subtask-accordion",
        preventOnFilter: false,
        // Reordering within a column only sticks in Manual sort. In any other sort
        // the order is computed, so within-column drags are disabled (cards snap
        // back); cross-column drags still work (they change Status).
        sort: manual,
        onStart: () => { State.dragging = true; },
        onEnd: handleDrop
      });
      State.sortableInstances.push(s);
    });
  }

  async function handleDrop(evt) {
    const taskId = Number(evt.item.dataset.taskId);
    const newStatus = evt.to.dataset.status;
    const sourceTask = State.tasks.find((t) => Number(t.TaskID) === taskId);
    if (!sourceTask) { State.dragging = false; return; }

    // In a computed sort, a within-column drop changes nothing — ignore it
    // (cross-column drops still change Status; Manual sort persists reorders).
    if (evt.from === evt.to && State.sort !== "manual") { State.dragging = false; return; }

    const oldStatus = sourceTask.Status;
    const oldPct    = sourceTask.PercentComplete;
    const newIndex  = evt.newIndex; // position in destination column
    const destCards = Array.from(evt.to.querySelectorAll(".kcard"));

    // Optimistic update locally (the DOM is already where Sortable dropped it).
    sourceTask.Status = newStatus;
    if (newStatus === "Done") sourceTask.PercentComplete = 100;   // Done ⇒ 100%
    destCards.forEach((cardEl, idx) => {
      const tid = Number(cardEl.dataset.taskId);
      const t = State.tasks.find((x) => Number(x.TaskID) === tid);
      if (t) t.ColumnOrder = idx + 1;
    });

    try {
      await window.WsjfData.writeTaskStatus(taskId, newStatus, newIndex + 1);
      if (newStatus === "Done") { try { await loadAllSubtasks(); } catch (_) {} }
      else if (oldStatus === "Done") {
        // Reopening: mirror the server rule (subtask ratio, or 90% if none).
        const subs = State.subtasksByParent[taskId] || [];
        sourceTask.PercentComplete = subs.length
          ? Math.round(subs.filter((s) => String(s.Done).toLowerCase() === "yes").length / subs.length * 100)
          : 90;
      }
      State.lastSyncTs = Date.now();
      updateSyncLabel();

      if (State.sort === "manual") {
        // Manual order: keep the dropped position, just patch % + counts in place.
        const pct = sourceTask.PercentComplete;
        const fill = evt.item.querySelector(".progress-fill"); if (fill) fill.style.width = pct + "%";
        const pctEl = evt.item.querySelector(".pct"); if (pctEl) pctEl.textContent = pct + "%";
        refreshBoardCounts();
      } else {
        // Computed sort: re-render so the card lands in its sorted position.
        render();
      }
    } catch (err) {
      console.error("Drop write failed", err);
      sourceTask.Status = oldStatus;            // revert state…
      sourceTask.PercentComplete = oldPct;
      toast("Could not update task: " + err.message + " — reverting.", "error");
      try { await reloadTasks(); } catch (_) {}  // resync sibling ColumnOrder from server
      render();                                  // …and rebuild so the card snaps back
    } finally {
      State.dragging = false;
    }
  }

  // Lightweight board refresh after a drag: update column counts + the empty
  // "Drop here" placeholders without wiping/rebuilding the DOM (avoids the flash
  // and the Sortable re-init that made drags feel janky).
  function refreshBoardCounts() {
    const visible = computeVisibleTasks();
    document.querySelectorAll(".kcol").forEach((col) => {
      const status = col.dataset.status;
      const n = visible.filter((t) => t.Status === status).length;
      const cnt = col.querySelector(".kcol-count");
      if (cnt) cnt.textContent = "(" + n + ")";
      const body = col.querySelector(".kcol-body");
      if (!body) return;
      const hasCard = !!body.querySelector(".kcard");
      const empty = body.querySelector(".kcol-empty");
      if (hasCard && empty) empty.remove();
      if (!hasCard && !empty) {
        const e = document.createElement("div");
        e.className = "kcol-empty"; e.textContent = "Drop here";
        body.appendChild(e);
      }
    });
  }

  // ───── List view ─────
  const LIST_COLS = [
    { key: "Title",           label: "Title" },
    { key: "Status",          label: "Status" },
    { key: "Health",          label: "Health" },
    { key: "Owner",           label: "Owner" },
    { key: "WorkstreamID",    label: "Workstream" },
    { key: "GoalID",          label: "Goal" },
    { key: "Quarter",         label: "Quarter" },
    { key: "WSJF",            label: "WSJF", num: true },
    { key: "DueDate",         label: "Due" },
    { key: "PercentComplete", label: "%", num: true }
  ];

  // Displayed string for a column — used for both the cell and the filter values.
  function listCellValue(t, key) {
    switch (key) {
      case "WorkstreamID":    return workstreamName(t.WorkstreamID);
      case "GoalID":          return taskGoalIds(t).map(goalShortName).join(", ");
      case "Health":          return String(t.Health || "");
      case "WSJF":            return (Number(t.WSJF) || 0).toFixed(1);
      case "PercentComplete": return (Math.max(0, Math.min(100, Number(t.PercentComplete) || 0))) + "%";
      case "DueDate":         return t.DueDate ? formatDateShort(t.DueDate) : "";
      default:                return String(t[key] == null ? "" : t[key]);
    }
  }

  function applyListColFilters(rows) {
    const keys = Object.keys(State.listColFilters);
    if (!keys.length) return rows;
    return rows.filter((t) =>
      keys.every((k) => {
        const set = State.listColFilters[k];
        return set && set.has(listCellValue(t, k));
      })
    );
  }

  // One List row (+ its inline subtask-detail row when expanded).
  function listRowHtml(t) {
    const wsjf = Number(t.WSJF) || 0;
    const pct  = Math.max(0, Math.min(100, Number(t.PercentComplete) || 0));
    const subs = State.subtasksByParent[t.TaskID] || [];
    const expanded = State.expandedSubtasks.has(Number(t.TaskID));
    const caret = subs.length
      ? '<button class="row-expand" data-exp="' + t.TaskID + '" title="Subtasks">' + (expanded ? "▾" : "▸") + '</button> '
      : '';
    const health = String(t.Health || "");
    const healthCell = health
      ? '<span class="health-chip health-' + statusSlug(health) + '"><i class="health-dot health-' + statusSlug(health) + '"></i>' + escapeHtml(health) + '</span>'
      : '<span class="muted">—</span>';
    const goalIds = taskGoalIds(t);
    const goalCell = goalIds.length
      ? '<span class="list-goals">' + goalIds.map((gid) =>
          '<span class="goal-pill" style="background:' + goalColor(gid) + '22;border-color:' + goalColor(gid) + '">' +
          '<i class="goal-dot" style="background:' + goalColor(gid) + '"></i>' + escapeHtml(goalShortName(gid)) + '</span>').join("")
        + '</span>'
      : '<span class="muted">—</span>';
    const ownerName = String(t.Owner || "").trim();
    const ownerCell = ownerName
      ? '<span class="owner-cell"><span class="avatar avatar-sm avatar-owner" style="background:' + colorHash(ownerName) + '">' + initialsFromName(ownerName) + '</span>' + escapeHtml(ownerName) + '</span>'
      : '<span class="muted">—</span>';
    let h = '<tr data-task-id="' + t.TaskID + '">' +
      '<td class="sel-col"><input type="checkbox" class="sel-row" data-id="' + t.TaskID + '"' + (State.selected.has(Number(t.TaskID)) ? " checked" : "") + ' /></td>' +
      '<td class="list-title">' + caret + scheduleChip(t) + escapeHtml(t.Title || "") + '</td>' +
      '<td><span class="status-chip status-' + statusSlug(t.Status) + '">' + escapeHtml(t.Status || "") + '</span></td>' +
      '<td>' + healthCell + '</td>' +
      '<td>' + ownerCell + '</td>' +
      '<td>' + escapeHtml(workstreamName(t.WorkstreamID)) + '</td>' +
      '<td>' + goalCell + '</td>' +
      '<td>' + escapeHtml(t.Quarter || "") + '</td>' +
      '<td class="num"><span class="wsjf-pill ' + wsjfPillClass(wsjf) + '">' + wsjf.toFixed(1) + '</span></td>' +
      '<td>' + escapeHtml(t.DueDate ? formatDateShort(t.DueDate) : "") + '</td>' +
      '<td class="num">' + pct + '%</td>' +
      '</tr>';
    if (expanded) {
      h += '<tr class="subtask-detail" data-detail="' + t.TaskID + '"><td colspan="' + (LIST_COLS.length + 1) + '"></td></tr>';
    }
    return h;
  }

  // Group label for the List "Group by" option.
  function groupKeyLabel(t, key) {
    if (key === "WorkstreamID") return workstreamName(t.WorkstreamID) || "(none)";
    if (key === "GoalID") {
      const ids = String(t.GoalID || "").split(/[;,]/).map((s) => s.trim()).filter(Boolean);
      if (!ids.length) return "(none)";
      return ids.map((id) => {
        const g = State.goals.find((x) => String(x.GoalID) === id);
        return g ? (g.ShortName || g.GoalName || g.GoalID) : id;
      }).join(", ");
    }
    if (key === "Health") { const h = String(t.Health || "").trim(); return h || "(no health)"; }
    const v = t[key];
    return (v == null || v === "") ? "(none)" : String(v);
  }

  function renderList() {
    closeColumnFilter();
    const root = document.getElementById("listview");
    const visible = applyListColFilters(computeVisibleTasks().slice()).sort(listComparator);
    const activeFilters = Object.keys(State.listColFilters).length;

    let html = '<div class="list-toolbar">' +
      '<span class="list-count">' + visible.length + ' task' + (visible.length === 1 ? "" : "s") + '</span>' +
      (activeFilters ? '<button class="btn-link" id="list-clear-filters">✕ Clear filters (' + activeFilters + ')</button>' : "") +
      '<span id="bulk-bar" class="bulk-bar"></span>' +
      '</div>';

    html += '<table class="tasklist"><thead><tr>';
    html += '<th class="sel-col"><input type="checkbox" class="sel-all" title="Select all" /></th>';
    LIST_COLS.forEach((c) => {
      const sorted = State.listSort.key === c.key;
      const arrow = sorted ? (State.listSort.dir === "asc" ? "▲" : "▼") : "";
      const filtered = !!State.listColFilters[c.key];
      html += '<th class="' + (c.num ? "num" : "") + '">' +
        '<span class="th-cell">' +
          '<span class="th-label" data-sort="' + c.key + '">' + escapeHtml(c.label) +
            (arrow ? ' <span class="th-arrow">' + arrow + '</span>' : "") + '</span>' +
          '<button class="th-filter' + (filtered ? " active" : "") + '" data-filter="' + c.key + '" title="Filter">⏷</button>' +
        '</span>' +
        '</th>';
    });
    html += '</tr></thead><tbody>';

    if (visible.length === 0) {
      html += '<tr><td colspan="' + (LIST_COLS.length + 1) + '" class="list-empty">No tasks match your filters.</td></tr>';
    } else if (State.groupBy) {
      const groups = new Map();
      visible.forEach((t) => {
        const label = groupKeyLabel(t, State.groupBy);
        if (!groups.has(label)) groups.set(label, []);
        groups.get(label).push(t);
      });
      Array.from(groups.keys()).sort((a, b) => {
        if (a === "(none)") return 1;
        if (b === "(none)") return -1;
        return String(a).localeCompare(String(b), undefined, { numeric: true });
      }).forEach((label) => {
        const rows = groups.get(label);
        const dn = rows.filter((t) => t.Status === "Done").length;
        html += '<tr class="group-row"><td colspan="' + (LIST_COLS.length + 1) + '">' +
          escapeHtml(label) + ' <span class="group-count">' + dn + '/' + rows.length + ' done</span></td></tr>';
        rows.forEach((t) => { html += listRowHtml(t); });
      });
    } else {
      visible.forEach((t) => { html += listRowHtml(t); });
    }
    html += '</tbody></table>';
    root.innerHTML = html;

    Array.from(root.querySelectorAll(".sel-row")).forEach((cb) => {
      cb.addEventListener("click", (e) => e.stopPropagation());
      cb.addEventListener("change", (e) => {
        e.stopPropagation();
        const id = Number(cb.dataset.id);
        if (cb.checked) State.selected.add(id); else State.selected.delete(id);
        const all = root.querySelector(".sel-all");
        const rows = root.querySelectorAll(".sel-row");
        if (all) all.checked = rows.length > 0 && Array.from(rows).every((r) => r.checked);
        renderBulkBar();
      });
    });
    const selAll = root.querySelector(".sel-all");
    if (selAll) {
      const rows0 = root.querySelectorAll(".sel-row");
      selAll.checked = rows0.length > 0 && Array.from(rows0).every((r) => r.checked);
      selAll.addEventListener("change", () => {
        Array.from(root.querySelectorAll(".sel-row")).forEach((cb) => {
          cb.checked = selAll.checked;
          const id = Number(cb.dataset.id);
          if (selAll.checked) State.selected.add(id); else State.selected.delete(id);
        });
        renderBulkBar();
      });
    }
    renderBulkBar();

    Array.from(root.querySelectorAll(".th-label[data-sort]")).forEach((el) => {
      el.addEventListener("click", () => {
        const key = el.dataset.sort;
        if (State.listSort.key === key) State.listSort.dir = State.listSort.dir === "asc" ? "desc" : "asc";
        else { State.listSort.key = key; State.listSort.dir = "asc"; }
        renderList();
      });
    });
    Array.from(root.querySelectorAll(".th-filter[data-filter]")).forEach((el) => {
      el.addEventListener("click", (e) => { e.stopPropagation(); openColumnFilter(el.dataset.filter, el); });
    });
    const clr = document.getElementById("list-clear-filters");
    if (clr) clr.addEventListener("click", () => { State.listColFilters = {}; renderList(); });
    Array.from(root.querySelectorAll(".row-expand[data-exp]")).forEach((b) => {
      b.addEventListener("click", (e) => {
        e.stopPropagation();
        const id = Number(b.dataset.exp);
        if (State.expandedSubtasks.has(id)) State.expandedSubtasks.delete(id);
        else State.expandedSubtasks.add(id);
        renderList();
      });
    });
    Array.from(root.querySelectorAll("tr.subtask-detail[data-detail]")).forEach((tr) => {
      tr.querySelector("td").appendChild(buildSubtaskAccordion(Number(tr.dataset.detail)));
    });
    Array.from(root.querySelectorAll("tr[data-task-id]")).forEach((tr) => {
      tr.addEventListener("click", () => openEditModal(Number(tr.dataset.taskId)));
    });
  }

  // Google-Sheets-style per-column filter popup.
  function openColumnFilter(key, anchorEl) {
    closeColumnFilter();
    const col = LIST_COLS.find((c) => c.key === key);
    const base = computeVisibleTasks();
    const values = Array.from(new Set(base.map((t) => listCellValue(t, key))))
      .sort((a, b) => String(a).localeCompare(String(b), undefined, { numeric: true }));
    const current = State.listColFilters[key];
    const selected = new Set(current ? Array.from(current) : values);

    const pop = document.createElement("div");
    pop.className = "col-filter-pop";
    pop.innerHTML =
      '<div class="cf-head">Filter: ' + escapeHtml(col.label) + '</div>' +
      '<input type="text" class="cf-search" placeholder="Search values…" />' +
      '<label class="cf-all-row"><input type="checkbox" class="cf-all" /> <strong>(Select all)</strong></label>' +
      '<div class="cf-list"></div>' +
      '<div class="cf-actions">' +
        '<button class="btn btn-secondary btn-sm" data-cf="clear">Clear</button>' +
        '<button class="btn btn-primary btn-sm" data-cf="ok">OK</button>' +
      '</div>';
    document.body.appendChild(pop);

    const listEl = pop.querySelector(".cf-list");
    const allCb = pop.querySelector(".cf-all");
    function renderValues(ft) {
      ft = (ft || "").toLowerCase();
      listEl.innerHTML = "";
      values.filter((v) => !ft || String(v).toLowerCase().includes(ft)).forEach((v) => {
        const row = document.createElement("label");
        row.className = "cf-item";
        const cb = document.createElement("input");
        cb.type = "checkbox";
        cb.checked = selected.has(v);
        cb.addEventListener("change", () => {
          if (cb.checked) selected.add(v); else selected.delete(v);
          allCb.checked = values.every((x) => selected.has(x));
        });
        const span = document.createElement("span");
        span.textContent = v === "" ? "(blank)" : v;
        row.appendChild(cb);
        row.appendChild(span);
        listEl.appendChild(row);
      });
    }
    renderValues("");
    allCb.checked = values.every((v) => selected.has(v));

    pop.querySelector(".cf-search").addEventListener("input", (e) => renderValues(e.target.value));
    allCb.addEventListener("change", () => {
      if (allCb.checked) values.forEach((v) => selected.add(v)); else selected.clear();
      renderValues(pop.querySelector(".cf-search").value);
    });
    pop.querySelector('[data-cf="clear"]').addEventListener("click", () => {
      delete State.listColFilters[key];
      renderList();
    });
    pop.querySelector('[data-cf="ok"]').addEventListener("click", () => {
      if (selected.size === values.length) delete State.listColFilters[key];
      else State.listColFilters[key] = new Set(selected);
      renderList();
    });

    const r = anchorEl.getBoundingClientRect();
    pop.style.position = "fixed";
    pop.style.top = Math.min(r.bottom + 4, window.innerHeight - 380) + "px";
    pop.style.left = Math.max(8, Math.min(r.left, window.innerWidth - 288)) + "px";

    setTimeout(() => document.addEventListener("mousedown", pop._closer = function (e) {
      if (!pop.contains(e.target)) closeColumnFilter();
    }, true), 0);
    State._openColFilter = pop;
  }

  function closeColumnFilter() {
    const pop = State._openColFilter;
    if (!pop) return;
    if (pop._closer) document.removeEventListener("mousedown", pop._closer, true);
    pop.remove();
    State._openColFilter = null;
  }

  // ───── Bulk actions (List view multi-select) ─────
  function renderBulkBar() {
    const bar = document.getElementById("bulk-bar");
    if (!bar) return;
    const n = State.selected.size;
    if (!n) { bar.innerHTML = ""; bar.classList.remove("active"); return; }
    bar.classList.add("active");
    bar.innerHTML =
      '<span class="bulk-count">' + n + ' selected</span>' +
      '<div class="bulk-edit-wrap">' +
        '<button class="btn btn-secondary btn-sm" id="bulk-edit-btn">✎ Edit field ▾</button>' +
        '<div class="bulk-edit-pop" id="bulk-edit-pop" hidden>' +
          '<select id="bulk-field" class="filter">' +
            '<option value="">Choose field…</option>' +
            '<option value="status">Status</option>' +
            '<option value="health">Health</option>' +
            '<option value="owner">Owner</option>' +
            '<option value="addcontrib">Add contributor</option>' +
            '<option value="workstream">Workstream</option>' +
            '<option value="goal">Goal</option>' +
            '<option value="quarter">Quarter</option>' +
            '<option value="pct">% complete</option>' +
            '<option value="startdate">Start date</option>' +
            '<option value="duedate">Due date</option>' +
          '</select>' +
          '<div id="bulk-value" class="bulk-value"></div>' +
          '<button class="btn btn-primary btn-sm" id="bulk-apply" disabled>Apply</button>' +
        '</div>' +
      '</div>' +
      '<button class="btn btn-secondary btn-sm" data-bulk="earlier" title="Pull in one quarter — completing sooner (logs Accelerated)">⏪ Accelerate</button>' +
      '<button class="btn btn-secondary btn-sm" data-bulk="later" title="Roll over one quarter — a slip (logs Delayed)">Delay ⏩</button>' +
      '<button class="btn btn-archive btn-sm" data-bulk="archive">Archive</button>' +
      '<button class="btn btn-secondary btn-sm" data-bulk="clear">Clear</button>';

    const editBtn = bar.querySelector("#bulk-edit-btn");
    const pop = bar.querySelector("#bulk-edit-pop");
    const fieldSel = bar.querySelector("#bulk-field");
    const valWrap = bar.querySelector("#bulk-value");
    const applyBtn = bar.querySelector("#bulk-apply");
    editBtn.addEventListener("click", (e) => { e.stopPropagation(); pop.hidden = !pop.hidden; });
    document.addEventListener("mousedown", (e) => { if (!pop.hidden && !bar.querySelector(".bulk-edit-wrap").contains(e.target)) pop.hidden = true; });
    fieldSel.addEventListener("change", () => { renderBulkValue(fieldSel.value, valWrap); applyBtn.disabled = !fieldSel.value; });
    applyBtn.addEventListener("click", () => {
      const field = fieldSel.value; if (!field) return;
      const ctrl = valWrap.querySelector("select, input");
      const val = ctrl ? ctrl.value : "";
      pop.hidden = true;
      applyBulk(field, val);
    });
    bar.querySelector('[data-bulk="earlier"]').addEventListener("click", () => applyBulk("earlier"));
    bar.querySelector('[data-bulk="later"]').addEventListener("click", () => applyBulk("later"));
    bar.querySelector('[data-bulk="archive"]').addEventListener("click", () => applyBulk("archive"));
    bar.querySelector('[data-bulk="clear"]').addEventListener("click", () => { State.selected.clear(); renderList(); });
  }

  // Build the value control for the chosen bulk field.
  function renderBulkValue(field, wrap) {
    const opt = (v, label, sel) => '<option value="' + escapeAttr(v) + '"' + (sel ? " selected" : "") + '>' + escapeHtml(label) + '</option>';
    const sel = (opts) => '<select class="filter">' + opts + '</select>';
    if (!field) { wrap.innerHTML = ""; return; }
    switch (field) {
      case "status": wrap.innerHTML = sel((State.config.Statuses || []).map((s) => opt(s, s)).join("")); break;
      case "health": wrap.innerHTML = sel(opt("", "— clear —") + ["On Track", "At Risk", "Off Track"].map((h) => opt(h, h)).join("")); break;
      case "owner":
      case "addcontrib": wrap.innerHTML = sel((State.config.Owners || []).map((o) => opt(o, o)).join("")); break;
      case "workstream": wrap.innerHTML = sel((State.workstreams || []).map((w) => opt(w.WorkstreamID, w.Name || w.WorkstreamID)).join("")); break;
      case "goal": wrap.innerHTML = sel((State.goals || []).map((g) => opt(g.GoalID, g.ShortName || g.GoalName || g.GoalID)).join("")); break;
      case "quarter": wrap.innerHTML = sel((State.config.Quarters || []).map((q) => opt(q, q)).join("")); break;
      case "pct": wrap.innerHTML = sel([0, 10, 20, 25, 30, 40, 50, 60, 70, 75, 80, 90, 100].map((p) => opt(String(p), p + "%")).join("")); break;
      case "startdate":
      case "duedate": wrap.innerHTML = '<input type="date" class="filter" />'; break;
      default: wrap.innerHTML = ""; break;
    }
  }

  async function applyBulk(kind, val) {
    const ids = Array.from(State.selected);
    if (!ids.length) return;
    const labels = {
      status: 'set status to "' + val + '"',
      health: val ? 'set health to "' + val + '"' : "clear health",
      owner: 'set owner to "' + val + '"',
      addcontrib: 'add contributor "' + val + '"',
      workstream: 'set workstream to "' + workstreamName(val) + '"',
      goal: 'set goal',
      quarter: 'set quarter to "' + val + '"',
      pct: 'set % complete to ' + val + '%',
      startdate: 'set start date to ' + val,
      duedate: 'set due date to ' + val,
      earlier: 'accelerate one quarter (pull in)',
      later: 'delay one quarter (roll over)',
      archive: 'archive'
    };
    const label = labels[kind] || kind;
    if (!(await uiConfirm("Apply to " + ids.length + " task" + (ids.length === 1 ? "" : "s") + ": " + label + "?", { okText: "Apply" }))) return;
    toast("Updating " + ids.length + "…", "info");
    // Field-per-kind for the simple writeTask cases.
    const FIELD = { owner: "Owner", health: "Health", workstream: "WorkstreamID", goal: "GoalID", quarter: "Quarter", pct: "PercentComplete", startdate: "StartDate", duedate: "DueDate" };
    let ok = 0, fail = 0;
    for (const id of ids) {
      try {
        if (kind === "status") await window.WsjfData.writeTaskStatus(id, val, 999);
        else if (kind === "archive") await window.WsjfData.archiveTask(id);
        else if (kind === "earlier" || kind === "later") {
          const r = await rescheduleTask(id, kind);
          if (!r || r.error) { fail++; continue; }
        }
        else if (kind === "addcontrib") {
          const task = State.tasks.find((x) => Number(x.TaskID) === Number(id));
          const set = new Set(String(task && task.Contributors || "").split(/[;,]/).map((s) => s.trim()).filter(Boolean));
          if (val && val !== (task && task.Owner)) set.add(val);
          if (task) set.delete(task.Owner);
          await window.WsjfData.writeTask({ TaskID: id, Contributors: Array.from(set).join("; ") }, { force: true });
        }
        else if (FIELD[kind]) {
          const value = kind === "pct" ? Number(val) : val;
          await window.WsjfData.writeTask({ TaskID: id, [FIELD[kind]]: value }, { force: true });
        }
        ok++;
      } catch (e) { fail++; console.warn("bulk fail", id, e); }
    }
    State.selected.clear();
    await reloadTasks();
    render();
    toast(ok + " updated" + (fail ? ", " + fail + " failed" : "") + ".", fail ? "warn" : "info");
  }

  function listComparator(a, b) {
    const k = State.listSort.key;
    const dir = State.listSort.dir === "asc" ? 1 : -1;
    if (k === "WSJF" || k === "PercentComplete") {
      return ((Number(a[k]) || 0) - (Number(b[k]) || 0)) * dir;
    }
    if (k === "DueDate") {
      const ax = a.DueDate ? new Date(a.DueDate).getTime() : Infinity;
      const bx = b.DueDate ? new Date(b.DueDate).getTime() : Infinity;
      return (ax - bx) * dir;
    }
    let av = a[k], bv = b[k];
    if (k === "WorkstreamID") { av = workstreamName(av); bv = workstreamName(bv); }
    else if (k === "GoalID") { av = taskGoalIds(a).map(goalShortName).join(", "); bv = taskGoalIds(b).map(goalShortName).join(", "); }
    else if (k === "Health") {
      // order by severity, not alphabet: Off Track → At Risk → On Track → (none)
      const rank = { "Off Track": 0, "At Risk": 1, "On Track": 2, "": 3 };
      return ((rank[String(a.Health || "")] ?? 3) - (rank[String(b.Health || "")] ?? 3)) * dir;
    }
    return String(av || "").localeCompare(String(bv || "")) * dir;
  }

  // ───── Filtering / Sorting ─────
  function computeVisibleTasks() {
    return State.tasks.filter((t) => {
      if (State.filters.owner && !String(t.Owner || "").includes(State.filters.owner)) return false;
      if (State.filters.workstream && t.WorkstreamID !== State.filters.workstream) return false;
      if (State.filters.quarter && t.Quarter !== State.filters.quarter) return false;
      if (State.filters.status && String(t.Status || "") !== State.filters.status) return false;
      if (State.filters.goal) {
        const gids = String(t.GoalID || "").split(/[;,]/).map((s) => s.trim());
        if (!gids.includes(State.filters.goal)) return false;
      }
      if (State.filters.health && String(t.Health || "") !== State.filters.health) return false;
      if (State.filters.startFrom || State.filters.startTo) {
        const sIso = isoDate(t.StartDate);
        if (!sIso) return false;
        if (State.filters.startFrom && sIso < State.filters.startFrom) return false;
        if (State.filters.startTo && sIso > State.filters.startTo) return false;
      }
      if (State.filters.dueFrom || State.filters.dueTo) {
        const dIso = isoDate(t.DueDate);
        if (!dIso) return false;
        if (State.filters.dueFrom && dIso < State.filters.dueFrom) return false;
        if (State.filters.dueTo && dIso > State.filters.dueTo) return false;
      }
      if (State.filters.subtasks) {
        const n = (State.subtasksByParent[t.TaskID] || []).length;
        if (State.filters.subtasks === "has" && n === 0) return false;
        if (State.filters.subtasks === "none" && n > 0) return false;
      }
      if (State.filters.search) {
        const q = State.filters.search.toLowerCase();
        const hay = [
          t.Title, t.Owner, t.Tags, t.Description,
          workstreamName(t.WorkstreamID), t.Quarter, t.Status
        ].map((x) => String(x == null ? "" : x).toLowerCase()).join(" ");
        if (!hay.includes(q)) return false;
      }
      if (State.selectedTags.size > 0) {
        const tags = new Set(String(t.Tags || "").split(";").map((s) => s.trim()).filter(Boolean));
        let any = false;
        State.selectedTags.forEach((tg) => { if (tags.has(tg)) any = true; });
        if (!any) return false;
      }
      return true;
    });
  }

  function sortComparator(a, b) {
    switch (State.sort) {
      case "wsjf":
        return (Number(b.WSJF) || 0) - (Number(a.WSJF) || 0);
      case "due": {
        const ax = a.DueDate ? new Date(a.DueDate).getTime() : Infinity;
        const bx = b.DueDate ? new Date(b.DueDate).getTime() : Infinity;
        return ax - bx;
      }
      case "pct":
        return (Number(b.PercentComplete) || 0) - (Number(a.PercentComplete) || 0);
      case "title":
        return String(a.Title || "").localeCompare(String(b.Title || ""));
      case "owner":
        return String(a.Owner || "").localeCompare(String(b.Owner || ""));
      default:
        return (Number(a.ColumnOrder) || 0) - (Number(b.ColumnOrder) || 0);
    }
  }

  // ───── My Tasks dashboard ─────
  function byWsjfDesc(a, b) { return (Number(b.WSJF) || 0) - (Number(a.WSJF) || 0); }

  function renderMyDashboard() {
    const root = document.getElementById("mineview");
    root.innerHTML = "";

    const known = State.me && State.me.name && State.me.name !== "Unknown User";
    if (!known) {
      const setup = document.createElement("div");
      setup.className = "mine-setup";
      setup.innerHTML = '<h2>Who are you?</h2><p class="muted">Set your name to get a personalized dashboard of your tasks.</p>';
      const b = document.createElement("button");
      b.className = "btn btn-primary";
      b.textContent = "Set your name";
      b.addEventListener("click", async () => { const me = await openIdentityPicker(); if (me) render(); });
      setup.appendChild(b);
      root.appendChild(setup);
      return;
    }

    const meName = State.me.name.toLowerCase();
    const meEmail = State.me.email ? String(State.me.email).toLowerCase() : "";
    const matchMe = (s) => { const v = String(s || "").toLowerCase(); return v.includes(meName) || (meEmail && v.includes(meEmail)); };
    // "Connected" = I own it OR I'm a contributor.
    const isConnected = (t) => matchMe(t.Owner) || matchMe(t.Contributors);
    const mine = State.tasks.filter(isConnected);
    const open = mine.filter((t) => t.Status !== "Done");
    const todayIso = new Date().toISOString().slice(0, 10);
    const in7Iso = new Date(Date.now() + 7 * 86400000).toISOString().slice(0, 10);
    const dueIsoOf = (t) => isoDate(t.DueDate);
    const overdue = open.filter((t) => { const d = dueIsoOf(t); return d && d < todayIso; });
    const dueSoon = open.filter((t) => { const d = dueIsoOf(t); return d && d >= todayIso && d <= in7Iso; });
    const blocked = mine.filter((t) => t.Status === "Blocked");
    const doneThisQ = mine.filter((t) => t.Status === "Done" && t.Quarter === currentQuarter());

    const header = document.createElement("div");
    header.className = "mine-header";
    header.innerHTML = '<h2>My Tasks</h2><span class="mine-who"></span>';
    header.querySelector(".mine-who").textContent = State.me.name + " · owned + contributing";
    root.appendChild(header);

    const stats = document.createElement("div");
    stats.className = "mine-stats";
    const cards = [
      { label: "Open", value: open.length, cls: "" },
      { label: "Overdue", value: overdue.length, cls: overdue.length ? "stat-danger" : "" },
      { label: "Due ≤ 7 days", value: dueSoon.length, cls: dueSoon.length ? "stat-warn" : "" },
      { label: "Blocked", value: blocked.length, cls: blocked.length ? "stat-danger" : "" },
      { label: "Done this " + currentQuarter(), value: doneThisQ.length, cls: "stat-good" }
    ];
    stats.innerHTML = cards.map((c) =>
      '<div class="mine-stat ' + c.cls + '"><div class="ms-value">' + c.value + '</div><div class="ms-label">' + escapeHtml(c.label) + '</div></div>'
    ).join("");
    root.appendChild(stats);

    // ⚠ Needs attention = overdue or blocked. Only shown when there's something —
    // so it's never a confusing empty "(0)" section.
    const attention = overdue.concat(blocked.filter((b) => overdue.indexOf(b) < 0)).sort(byDueThenWsjf);
    if (attention.length) {
      root.appendChild(mineSection("⚠ Needs attention — overdue or blocked", attention, ""));
    }

    // Recent updates on tasks I'm connected to.
    const updSec = buildRecentUpdates(mine);
    if (updSec) root.appendChild(updSec);

    // Priorities = active work (excludes Backlog and Done), soonest due first, then WSJF.
    const attSet = new Set(attention.map((t) => Number(t.TaskID)));
    const active = open.filter((t) => t.Status !== "Backlog" && !attSet.has(Number(t.TaskID)));
    const priorities = active.sort(byDueThenWsjf);
    root.appendChild(mineSection("My priorities — by due date, then WSJF", priorities, "No active tasks. 🎉"));

    const done = mine.filter((t) => t.Status === "Done").sort(byWsjfDesc).slice(0, 6);
    root.appendChild(mineSection("Recently completed", done, "Nothing completed yet."));
  }

  // Soonest due first (no due date = last), then highest WSJF.
  function byDueThenWsjf(a, b) {
    const ad = isoDate(a.DueDate) || "9999-12-31";
    const bd = isoDate(b.DueDate) || "9999-12-31";
    if (ad !== bd) return ad < bd ? -1 : 1;
    return (Number(b.WSJF) || 0) - (Number(a.WSJF) || 0);
  }

  // "Recent updates" feed for the My Tasks view — newest updates on any task the
  // user owns or contributes to (including updates posted on its subtasks).
  function buildRecentUpdates(mineTasks) {
    if (!State.allUpdates || !State.allUpdates.length) return null;
    const myTaskIds = new Set(mineTasks.map((t) => Number(t.TaskID)));
    // subtaskId → parentTaskId
    const subParent = {};
    Object.keys(State.subtasksByParent).forEach((pid) => {
      (State.subtasksByParent[pid] || []).forEach((s) => { subParent[Number(s.SubtaskID)] = Number(pid); });
    });
    const titleOf = (tid) => { const t = State.tasks.find((x) => Number(x.TaskID) === Number(tid)); return t ? t.Title : ""; };
    const rows = State.allUpdates.map((u) => {
      let tid = null;
      if (String(u.ParentType) === "Task") tid = Number(u.ParentID);
      else if (String(u.ParentType) === "Subtask") tid = subParent[Number(u.ParentID)];
      return { u: u, tid: tid };
    }).filter((r) => r.tid != null && myTaskIds.has(r.tid))
      .sort((a, b) => String(b.u.AddedDate).localeCompare(String(a.u.AddedDate)))
      .slice(0, 6);
    if (!rows.length) return null;

    const sec = document.createElement("section");
    sec.className = "mine-section";
    const h = document.createElement("h3");
    h.textContent = "Recent updates (" + rows.length + ")";
    sec.appendChild(h);
    rows.forEach((r) => {
      const div = document.createElement("div");
      div.className = "mine-update";
      div.innerHTML =
        '<div class="mu-task">' + escapeHtml(titleOf(r.tid)) +
          (String(r.u.ParentType) === "Subtask" ? ' <span class="mu-on">· subtask</span>' : '') + '</div>' +
        '<div class="mu-text">' + escapeHtml(r.u.Text || "") + '</div>' +
        '<div class="mu-meta">' + escapeHtml(r.u.AddedBy || "?") + ' · ' + relTime(r.u.AddedDate) + '</div>';
      div.addEventListener("click", () => openEditModal(Number(r.tid)));
      sec.appendChild(div);
    });
    return sec;
  }

  function mineSection(title, tasks, emptyMsg) {
    const sec = document.createElement("section");
    sec.className = "mine-section";
    const h = document.createElement("h3");
    h.textContent = title + " (" + tasks.length + ")";
    sec.appendChild(h);
    if (!tasks.length) {
      const e = document.createElement("div");
      e.className = "mine-empty";
      e.textContent = emptyMsg;
      sec.appendChild(e);
      return sec;
    }
    tasks.forEach((t) => sec.appendChild(mineTaskRow(t)));
    return sec;
  }

  function mineTaskRow(t) {
    const wsjf = Number(t.WSJF) || 0;
    const pct = Math.max(0, Math.min(100, Number(t.PercentComplete) || 0));
    const subs = State.subtasksByParent[t.TaskID] || [];
    const subDone = subs.filter((s) => String(s.Done).toLowerCase() === "yes").length;
    const todayIso = new Date().toISOString().slice(0, 10);
    const dIso = isoDate(t.DueDate);
    const isOverdue = dIso && dIso < todayIso && t.Status !== "Done";
    const row = document.createElement("div");
    row.className = "mine-row";
    row.innerHTML =
      '<span class="wsjf-pill ' + wsjfPillClass(wsjf) + '">' + wsjf.toFixed(1) + '</span>' +
      '<span class="mine-row-title">' + escapeHtml(t.Title || "") + '</span>' +
      '<span class="status-chip status-' + statusSlug(t.Status) + '">' + escapeHtml(t.Status || "") + '</span>' +
      (dIso ? '<span class="mine-due' + (isOverdue ? " overdue" : "") + '">📅 ' + escapeHtml(formatDateShort(t.DueDate)) + '</span>' : "") +
      (subs.length ? '<span class="mine-sub">✓ ' + subDone + '/' + subs.length + '</span>' : "") +
      '<span class="mine-pct">' + pct + '%</span>';
    row.addEventListener("click", () => openEditModal(Number(t.TaskID)));
    return row;
  }

  // ───── Workstreams view ─────
  function goalLabel(token) {
    token = String(token).trim();
    const g = State.goals.find((x) => String(x.GoalID) === token);
    return g ? (g.ShortName || g.GoalName || g.GoalID) : token;
  }

  // Is an open task in trouble? Blocked, health-flagged, or overdue.
  // (The quarter dashboard adds a "behind pace" check on top of this.)
  function isAtRisk(t) {
    if (t.Status === "Done") return false;
    if (t.Status === "Blocked") return true;
    const h = String(t.Health || "");
    if (h === "At Risk" || h === "Off Track") return true;
    const d = isoDate(t.DueDate);
    return d && d < new Date().toISOString().slice(0, 10);
  }

  function renderWorkstreams() {
    const root = document.getElementById("wsview");
    root.innerHTML = "";
    if (!State.wsQuarter) State.wsQuarter = currentQuarter();

    root.appendChild(renderWsDashboard());

    const header = document.createElement("div");
    header.className = "mine-header ws-grid-header";
    header.innerHTML = '<h3>All workstreams</h3>' +
      '<span class="ws-grid-head-right"><span class="mine-who">' + State.workstreams.length + ' streams</span>' +
      '<button class="btn btn-primary btn-sm" id="ws-add-btn">+ Add workstream</button></span>';
    root.appendChild(header);
    header.querySelector("#ws-add-btn").addEventListener("click", () => openWorkstreamEditor(null));

    if (!State.workstreams.length) {
      const e = document.createElement("div");
      e.className = "mine-empty";
      e.textContent = "No workstreams yet. Click “+ Add workstream”.";
      root.appendChild(e);
      return;
    }

    // Group workstreams by their optional Group column (insertion order preserved).
    const groups = [];
    const byName = {};
    State.workstreams.forEach((w) => {
      const g = String(w.Group || "").trim() || "Workstreams";
      if (!byName[g]) { byName[g] = { name: g, items: [] }; groups.push(byName[g]); }
      byName[g].items.push(w);
    });
    // Order within a group by optional Order column, then name.
    groups.forEach((g) => g.items.sort((a, b) =>
      ((Number(a.Order) || 9999) - (Number(b.Order) || 9999)) ||
      String(a.Name || a.WorkstreamID).localeCompare(String(b.Name || b.WorkstreamID))));

    const single = groups.length === 1 && groups[0].name === "Workstreams";
    groups.forEach((g) => root.appendChild(renderWsGroup(g, single)));
  }

  // Collapsible group of workstream cards. Default closed; remembers last state.
  function renderWsGroup(g, forceOpen) {
    const key = "wsGroupOpen:" + g.name;
    const open = forceOpen || lsGet(key, "0") === "1";
    const sec = document.createElement("section");
    sec.className = "ws-group";
    const head = document.createElement("button");
    head.className = "ws-group-head";
    head.innerHTML = '<span class="ws-group-caret">' + (open ? "▾" : "▸") + '</span>' +
      '<span class="ws-group-name">' + escapeHtml(g.name) + '</span>' +
      '<span class="ws-group-count">' + g.items.length + ' workstream' + (g.items.length === 1 ? "" : "s") + '</span>';
    sec.appendChild(head);
    const grid = document.createElement("div");
    grid.className = "ws-grid";
    grid.hidden = !open;
    g.items.forEach((w) => grid.appendChild(workstreamCard(w)));
    sec.appendChild(grid);
    if (!forceOpen) head.addEventListener("click", () => {
      const nowOpen = grid.hidden;
      grid.hidden = !nowOpen;
      head.querySelector(".ws-group-caret").textContent = nowOpen ? "▾" : "▸";
      lsSet(key, nowOpen ? "1" : "0");
    });
    return sec;
  }

  // Quarter "bird's-eye" dashboard above the workstream cards.
  function renderWsDashboard() {
    const q = State.wsQuarter;
    const qtasks = State.tasks.filter((t) => t.Quarter === q);
    const total = qtasks.length;
    const done = qtasks.filter((t) => t.Status === "Done").length;
    const pct = total ? Math.round((done / total) * 100) : 0;
    const open = qtasks.filter((t) => t.Status !== "Done");

    // Pace: how far through the quarter are we vs. how much is done?
    const dleft = daysLeftInQuarter(q);
    const qDays = { Q1: 90, Q2: 91, Q3: 92, Q4: 92 }[q] || 91;
    const elapsedPct = dleft == null ? null : Math.max(0, Math.min(100, Math.round(((qDays - dleft) / qDays) * 100)));
    const behind = elapsedPct != null && total > 0 && pct < elapsedPct - 10;

    // At risk = blocked / off-track / overdue (isAtRisk) OR "behind pace": an open
    // task whose % trails the quarter's elapsed % by >15 pts. This makes the list
    // reflect the "behind pace" banner instead of staying empty until tasks go overdue.
    const behindPace = (t) => t.Status !== "Done" && elapsedPct != null &&
      (Number(t.PercentComplete) || 0) < elapsedPct - 15;
    const atRisk = qtasks.filter((t) => isAtRisk(t) || behindPace(t)).sort(byWsjfDesc);

    // Per-workstream at-risk counts.
    const wsRisk = {};
    atRisk.forEach((t) => { wsRisk[t.WorkstreamID] = (wsRisk[t.WorkstreamID] || 0) + 1; });
    const riskWsIds = Object.keys(wsRisk).sort((a, b) => wsRisk[b] - wsRisk[a]);

    const sec = document.createElement("section");
    sec.className = "ws-dashboard";

    const daysLabel = dleft == null ? "—" :
      dleft < 0 ? "ended" : dleft + " day" + (dleft === 1 ? "" : "s") + " left";

    // Header: title + quarter picker + days-left.
    const head = document.createElement("div");
    head.className = "ws-dash-head";
    head.innerHTML =
      '<h2>Quarter progress</h2>' +
      '<div class="ws-dash-controls">' +
        '<label class="ws-dash-q">Quarter <select id="ws-dash-quarter"></select></label>' +
        '<span class="ws-dash-daysleft' + (dleft != null && dleft >= 0 && dleft <= 14 ? " urgent" : "") + '">⏳ ' + daysLabel + '</span>' +
      '</div>';
    sec.appendChild(head);
    const qsel = head.querySelector("#ws-dash-quarter");
    (State.config.Quarters || []).forEach((qq) => {
      const o = document.createElement("option"); o.value = qq; o.textContent = qq;
      if (qq === q) o.selected = true; qsel.appendChild(o);
    });
    qsel.addEventListener("change", () => { State.wsQuarter = qsel.value; render(); });

    // Stat tiles.
    const stats = document.createElement("div");
    stats.className = "mine-stats";
    const tiles = [
      { label: q + " progress", value: pct + "%", cls: behind ? "stat-warn" : "stat-good" },
      { label: "Done", value: done + " / " + total, cls: "" },
      { label: "Open", value: open.length, cls: "" },
      { label: "At risk", value: atRisk.length, cls: atRisk.length ? "stat-danger" : "" }
    ];
    stats.innerHTML = tiles.map((c) =>
      '<div class="mine-stat ' + c.cls + '"><div class="ms-value">' + escapeHtml(String(c.value)) + '</div><div class="ms-label">' + escapeHtml(c.label) + '</div></div>'
    ).join("");
    sec.appendChild(stats);

    // Pace line.
    if (elapsedPct != null) {
      const pace = document.createElement("div");
      pace.className = "ws-pace " + (behind ? "pace-bad" : "pace-good");
      pace.textContent = behind
        ? "⚠ Behind pace — " + pct + "% done but " + elapsedPct + "% of the quarter has elapsed."
        : "✓ On pace — " + pct + "% done, " + elapsedPct + "% of the quarter elapsed.";
      sec.appendChild(pace);
    }

    // Workstreams at risk.
    if (riskWsIds.length) {
      const wrap = document.createElement("div");
      wrap.className = "ws-risk-block";
      wrap.innerHTML = '<div class="ws-sub-label">Workstreams at risk this ' + escapeHtml(q) + '</div>';
      const chips = document.createElement("div");
      chips.className = "ws-risk-chips";
      riskWsIds.forEach((id) => {
        const b = document.createElement("button");
        b.className = "ws-risk-chip";
        b.innerHTML = escapeHtml(workstreamName(id)) + ' <span class="rc-n">' + wsRisk[id] + '</span>';
        b.addEventListener("click", () => {
          State.filters.workstream = id; State.filters.quarter = q;
          State.viewMode = "list"; syncFilterControls(); render();
        });
        chips.appendChild(b);
      });
      wrap.appendChild(chips);
      sec.appendChild(wrap);
    }

    // Open & at-risk task list for the quarter — collapsible, default closed.
    const listSec = document.createElement("div");
    listSec.className = "ws-risk-tasks";
    if (!atRisk.length) {
      listSec.innerHTML = '<div class="ws-sub-label">Open &amp; at-risk tasks · ' + escapeHtml(q) + '</div>' +
        '<div class="mine-empty">Nothing at risk this quarter. 🎉</div>';
      sec.appendChild(listSec);
      return sec;
    }
    const key = "wsAtRiskOpen";
    const accOpen = lsGet(key, "0") === "1";
    const topNames = atRisk.slice(0, 3).map((t) => t.Title).join(" · ");
    const accHead = document.createElement("button");
    accHead.className = "ws-acc-head";
    accHead.innerHTML = '<span class="ws-acc-caret">' + (accOpen ? "▾" : "▸") + '</span>' +
      '<span class="ws-acc-title">Open &amp; at-risk tasks · ' + escapeHtml(q) + '</span>' +
      '<span class="ws-acc-count danger">' + atRisk.length + '</span>' +
      '<span class="ws-acc-summary">' + escapeHtml(topNames) + (atRisk.length > 3 ? " …" : "") + '</span>';
    const accBody = document.createElement("div");
    accBody.className = "ws-acc-body";
    accBody.hidden = !accOpen;
    atRisk.slice(0, 15).forEach((t) => accBody.appendChild(mineTaskRow(t)));
    if (atRisk.length > 15) {
      const more = document.createElement("div"); more.className = "muted ws-more";
      more.textContent = "+ " + (atRisk.length - 15) + " more — open the List view to see all.";
      accBody.appendChild(more);
    }
    accHead.addEventListener("click", () => {
      const nowOpen = accBody.hidden;
      accBody.hidden = !nowOpen;
      accHead.querySelector(".ws-acc-caret").textContent = nowOpen ? "▾" : "▸";
      lsSet(key, nowOpen ? "1" : "0");
    });
    listSec.appendChild(accHead);
    listSec.appendChild(accBody);
    sec.appendChild(listSec);
    return sec;
  }

  function workstreamCard(w) {
    const wsId = w.WorkstreamID;
    const tasks = State.tasks.filter((t) => t.WorkstreamID === wsId);
    const total = tasks.length;
    const done = tasks.filter((t) => t.Status === "Done").length;
    const blocked = tasks.filter((t) => t.Status === "Blocked").length;
    const todayIso = new Date().toISOString().slice(0, 10);
    const overdue = tasks.filter((t) => t.Status !== "Done" && isoDate(t.DueDate) && isoDate(t.DueDate) < todayIso).length;
    const pctDone = total ? Math.round((done / total) * 100) : 0;
    const avgWsjf = total ? (tasks.reduce((s, t) => s + (Number(t.WSJF) || 0), 0) / total) : 0;
    const cols = boardColumns();
    const counts = {};
    cols.forEach((s) => { counts[s] = tasks.filter((t) => t.Status === s).length; });

    const metrics = [w.Metric1, w.Metric2, w.Metric3].map((m) => String(m == null ? "" : m).trim()).filter(Boolean);
    const goals = String(w.Goals || "").split(/[;,]/).map((s) => s.trim()).filter(Boolean).map(goalLabel);
    const avgPct = total ? Math.round(tasks.reduce((s, t) => s + (Number(t.PercentComplete) || 0), 0) / total) : 0;

    // Per-quarter progress within this workstream (only quarters that have tasks).
    const quarters = (State.config.Quarters || []).filter((q) =>
      tasks.some((t) => t.Quarter === q));
    const qRowsHtml = quarters.map((q) => {
      const qt = tasks.filter((t) => t.Quarter === q);
      const qd = qt.filter((t) => t.Status === "Done").length;
      const qOver = qt.filter((t) => t.Status !== "Done" && isoDate(t.DueDate) && isoDate(t.DueDate) < todayIso).length;
      const qpct = qt.length ? Math.round((qd / qt.length) * 100) : 0;
      return '<button class="ws-q-row" data-q="' + escapeAttr(q) + '" title="View ' + escapeAttr(q) + ' tasks">' +
        '<span class="ws-q-name">' + escapeHtml(q) + (qOver ? ' <span class="ws-q-over">⚠' + qOver + '</span>' : '') + '</span>' +
        '<span class="ws-q-bar"><span class="ws-q-fill" style="width:' + qpct + '%"></span></span>' +
        '<span class="ws-q-num">' + qd + '/' + qt.length + '</span>' +
        '</button>';
    }).join("");

    const card = document.createElement("section");
    card.className = "ws-card";
    card.innerHTML =
      '<div class="ws-card-head">' +
        '<span class="ws-name">' + escapeHtml(w.Name || wsId) + '</span>' +
        '<span class="ws-head-right">' +
          (w.Status ? '<span class="status-chip status-' + statusSlug(w.Status) + '">' + escapeHtml(w.Status) + '</span>' : '') +
          '<button class="ws-edit-btn" title="Edit workstream">✎</button>' +
        '</span>' +
      '</div>' +
      (w.Owner ? '<div class="ws-owner">👤 ' + escapeHtml(w.Owner) + '</div>' : '') +
      (w.UserStory ? '<p class="ws-story">' + escapeHtml(w.UserStory) + '</p>' : '') +
      '<div class="ws-statbar" title="Status mix across this workstream’s tasks">' +
        cols.map((s) => counts[s] ? '<span class="ws-seg ws-seg-' + statusSlug(s) + '" style="flex:' + counts[s] + '" title="' + escapeAttr(s + ": " + counts[s]) + '"></span>' : '').join('') +
      '</div>' +
      '<div class="ws-progress-label"><b>' + done + ' / ' + total + '</b> done · ' + pctDone + '% · avg ' + avgPct + '% complete</div>' +
      '<div class="ws-legend">' + cols.filter((s) => counts[s]).map((s) =>
        '<span class="ws-leg"><i class="ws-seg-' + statusSlug(s) + '"></i>' + escapeHtml(s) + ' ' + counts[s] + '</span>').join('') + '</div>' +
      '<div class="ws-chips">' +
        '<span class="ws-chip">' + total + ' tasks</span>' +
        (blocked ? '<span class="ws-chip danger">' + blocked + ' blocked</span>' : '') +
        (overdue ? '<span class="ws-chip danger">' + overdue + ' overdue</span>' : '') +
        '<span class="ws-chip">WSJF ' + avgWsjf.toFixed(1) + ' avg</span>' +
      '</div>' +
      (qRowsHtml ? '<div class="ws-quarters"><div class="ws-sub-label">Progress by quarter</div>' + qRowsHtml + '</div>' : '') +
      (metrics.length ? '<div class="ws-metrics"><div class="ws-sub-label">Metrics</div><ul>' + metrics.map((m) => '<li>' + escapeHtml(m) + '</li>').join('') + '</ul></div>' : '') +
      (goals.length ? '<div class="ws-goals"><div class="ws-sub-label">Goals</div>' + goals.map((g) => '<span class="ws-goal-pill">' + escapeHtml(g) + '</span>').join('') + '</div>' : '') +
      '<button class="btn btn-secondary btn-sm ws-view-tasks">View ' + total + ' tasks →</button>';

    const drillTo = (quarter) => {
      State.filters.workstream = wsId;
      State.filters.quarter = quarter || "";
      State.viewMode = "list";
      syncFilterControls();   // reflect the drill-down in the filter chips + popover
      render();
    };
    card.querySelector(".ws-view-tasks").addEventListener("click", () => drillTo(""));
    Array.from(card.querySelectorAll(".ws-q-row")).forEach((b) =>
      b.addEventListener("click", () => drillTo(b.dataset.q)));
    const editBtn = card.querySelector(".ws-edit-btn");
    if (editBtn) editBtn.addEventListener("click", (e) => { e.stopPropagation(); openWorkstreamEditor(w); });
    return card;
  }

  // Edit a workstream's info (owner, status, story, metrics, goals, quarters)
  // directly from the Workstreams page.
  function openWorkstreamEditor(w) {
    const isNew = !w;
    w = w || {};
    const goals = State.goals || [];
    const quarters = State.config.Quarters || [];
    const curGoals = new Set(String(w.Goals || "").split(/[;,]/).map((s) => s.trim()).filter(Boolean));
    const curQuarters = new Set(String(w.Quarters || "").split(/[;,]/).map((s) => s.trim()).filter(Boolean));
    const curOwners = new Set(String(w.Owner || "").split(/[;,]/).map((s) => s.trim()).filter(Boolean));
    // People list = configured owners + any already-assigned owner not in the list.
    const ownerList = (State.config.Owners || []).slice();
    curOwners.forEach((o) => { if (ownerList.indexOf(o) < 0) ownerList.push(o); });
    const hasQuartersCol = State.workstreams[0] && Object.prototype.hasOwnProperty.call(State.workstreams[0], "Quarters");

    const host = document.createElement("div");
    host.className = "confirm-overlay";
    const box = document.createElement("div");
    box.className = "confirm-box ws-editor";
    box.innerHTML =
      '<div class="wse-head"><h3>' + (isNew ? "New workstream" : "Edit workstream") + '</h3><span class="wse-id">' + escapeHtml(w.WorkstreamID || "auto-ID") + '</span></div>' +
      '<label class="wse-label">Name<input id="wse-name" type="text" /></label>' +
      '<div class="wse-metrics">' +
        '<label class="wse-label">Group <span class="muted">(optional)</span><input id="wse-group" type="text" placeholder="e.g. Portfolio" /></label>' +
        '<label class="wse-label">Order <span class="muted">(optional)</span><input id="wse-order" type="number" placeholder="1" /></label>' +
      '</div>' +
      '<div class="wse-label">Owner(s)<div class="wse-checks" id="wse-owners"></div></div>' +
      '<label class="wse-label">Status<select id="wse-status"></select></label>' +
      '<label class="wse-label">User story<textarea id="wse-story" rows="3"></textarea></label>' +
      '<div class="wse-metrics">' +
        '<label class="wse-label">Metric 1<input id="wse-m1" type="text" /></label>' +
        '<label class="wse-label">Metric 2<input id="wse-m2" type="text" /></label>' +
        '<label class="wse-label">Metric 3<input id="wse-m3" type="text" /></label>' +
      '</div>' +
      '<div class="wse-label">Goals<div class="wse-checks" id="wse-goals"></div></div>' +
      '<div class="wse-label">Quarters it runs' +
        (hasQuartersCol ? '' : ' <span class="muted">(add a "Quarters" column to WorkstreamsTable to save this)</span>') +
        '<div class="wse-checks" id="wse-quarters"></div></div>' +
      '<div class="confirm-actions"><button class="btn btn-secondary" id="wse-cancel">Cancel</button><button class="btn btn-primary" id="wse-save">Save</button></div>';
    host.appendChild(box);
    document.body.appendChild(host);

    box.querySelector("#wse-name").value = w.Name || "";
    box.querySelector("#wse-group").value = w.Group || "";
    box.querySelector("#wse-order").value = (w.Order != null && w.Order !== "") ? w.Order : "";
    box.querySelector("#wse-story").value = w.UserStory || "";
    box.querySelector("#wse-m1").value = w.Metric1 || "";
    box.querySelector("#wse-m2").value = w.Metric2 || "";
    box.querySelector("#wse-m3").value = w.Metric3 || "";
    fillSelect("wse-status", ["", ...State.config.Statuses], "—", w.Status || "");

    const oWrap = box.querySelector("#wse-owners");
    ownerList.forEach((o) => {
      const l = document.createElement("label"); l.className = "wse-check";
      l.innerHTML = '<input type="checkbox" value="' + escapeAttr(o) + '" ' + (curOwners.has(o) ? "checked" : "") + ' /> ' + escapeHtml(o);
      oWrap.appendChild(l);
    });
    if (!ownerList.length) oWrap.innerHTML = '<span class="muted contrib-empty">No owners — add people in Config.</span>';

    const gWrap = box.querySelector("#wse-goals");
    goals.forEach((g) => {
      const l = document.createElement("label"); l.className = "wse-check";
      l.innerHTML = '<input type="checkbox" value="' + escapeAttr(g.GoalID) + '" ' + (curGoals.has(g.GoalID) ? "checked" : "") + ' /> ' +
        escapeHtml(g.ShortName || g.GoalName || g.GoalID);
      gWrap.appendChild(l);
    });
    const qWrap = box.querySelector("#wse-quarters");
    quarters.forEach((qq) => {
      const l = document.createElement("label"); l.className = "wse-check";
      l.innerHTML = '<input type="checkbox" value="' + escapeAttr(qq) + '" ' + (curQuarters.has(qq) ? "checked" : "") + ' /> ' + escapeHtml(qq);
      qWrap.appendChild(l);
    });

    const close = () => host.remove();
    box.querySelector("#wse-cancel").addEventListener("click", close);
    host.addEventListener("click", (e) => { if (e.target === host) close(); });
    box.querySelector("#wse-save").addEventListener("click", async () => {
      const getChecks = (sel) => Array.from(box.querySelectorAll(sel + " input:checked")).map((c) => c.value).join(",");
      const ownerVal = Array.from(box.querySelectorAll("#wse-owners input:checked")).map((c) => c.value).join("; ");
      const orderVal = box.querySelector("#wse-order").value.trim();
      const fields = {
        Name: box.querySelector("#wse-name").value.trim(),
        Group: box.querySelector("#wse-group").value.trim(),
        Order: orderVal === "" ? "" : Number(orderVal),
        Owner: ownerVal,
        Status: box.querySelector("#wse-status").value,
        UserStory: box.querySelector("#wse-story").value.trim(),
        Metric1: box.querySelector("#wse-m1").value.trim(),
        Metric2: box.querySelector("#wse-m2").value.trim(),
        Metric3: box.querySelector("#wse-m3").value.trim(),
        Goals: getChecks("#wse-goals")
      };
      if (hasQuartersCol) fields.Quarters = getChecks("#wse-quarters");
      if (!fields.Name) { toast("Name is required.", "warn"); return; }
      const saveBtn = box.querySelector("#wse-save"); saveBtn.disabled = true; saveBtn.textContent = "Saving…";
      try {
        if (isNew) {
          const id = nextSequentialId(State.workstreams, "WorkstreamID", "WS");
          await window.WsjfData.createWorkstream(Object.assign({ WorkstreamID: id }, fields));
          toast("Workstream " + id + " created.", "info");
        } else {
          await window.WsjfData.updateWorkstream(w.WorkstreamID, fields);
          toast("Workstream updated.", "info");
        }
        close();
        await afterConfigChange();
      } catch (e) {
        saveBtn.disabled = false; saveBtn.textContent = "Save";
        toast("Save failed: " + e.message, "error");
      }
    });
  }

  // ───── Roadmap (timeline) ─────
  function renderRoadmap() {
    const root = document.getElementById("roadmapview");
    const visible = computeVisibleTasks();
    const yr = new Date().getFullYear();
    const yearStart = isoDate((State.quarterDates.Q1 || {}).start) || (yr + "-01-01");
    const yearEnd = isoDate((State.quarterDates.Q4 || {}).end) || (yr + "-12-31");
    const t0 = new Date(yearStart + "T00:00:00").getTime();
    const t1 = new Date(yearEnd + "T00:00:00").getTime();
    const span = Math.max(1, t1 - t0);
    const pctOf = (iso) => {
      const tt = new Date(iso + "T00:00:00").getTime();
      if (isNaN(tt)) return 0;
      return Math.max(0, Math.min(100, ((tt - t0) / span) * 100));
    };
    const isoFromPct = (p) => new Date(t0 + (Math.max(0, Math.min(100, p)) / 100) * span).toISOString().slice(0, 10);
    const todayPct = pctOf(new Date().toISOString().slice(0, 10));

    // A task's bar range — explicit dates, else the quarter's bounds.
    function rng(task) {
      let s = isoDate(task.StartDate), e = isoDate(task.DueDate);
      const q = State.quarterDates[task.Quarter];
      if (!s) s = q ? isoDate(q.start) : (e || yearStart);
      if (!e) e = q ? isoDate(q.end) : (s || yearEnd);
      if (e < s) e = s;
      return { s: s, e: e };
    }

    const order = State.workstreams.map((w) => w.WorkstreamID)
      .filter((id) => visible.some((t) => t.WorkstreamID === id));
    if (visible.some((t) => !order.includes(t.WorkstreamID))) order.push("");

    // Goal legend (distinct colour per goal drives the bar colours).
    const usedGoals = [];
    visible.forEach((t) => taskGoalIds(t).forEach((g) => { if (!usedGoals.includes(g)) usedGoals.push(g); }));
    const legend = usedGoals.map((g) =>
      '<span class="rm-leg"><i style="background:' + goalColor(g) + '"></i>' + escapeHtml(goalShortName(g)) + '</span>').join("") +
      '<span class="rm-leg"><i style="background:#cbd5e1"></i>No goal</span>';

    let html = '<div class="rm-head"><h2>Roadmap</h2>' +
      '<div class="rm-head-right">' +
        '<label class="rm-toggle"><input type="checkbox" id="rm-show-dates"' + (State.roadmap.showDates ? " checked" : "") + '> Date tooltip on drag</label>' +
        '<span class="mine-who">' + visible.length + ' items · ' + escapeHtml(yearStart.slice(0, 4)) + '</span>' +
      '</div></div>';
    html += '<div class="rm-legend">' + legend + '<span class="rm-leg-hint">Drag a bar to move it · drag its edges to resize · click to open</span></div>';
    html += '<div class="rm-scroll"><div class="rm">';

    // Milestone vertical guide-lines spanning every track (overlay).
    if (State.milestones && State.milestones.length) {
      html += '<div class="rm-mslines">' + State.milestones.map((m) => {
        const d = isoDate(m.Date); if (!d) return "";
        return '<div class="rm-msline" style="left:' + pctOf(d) + '%;border-color:' + (m.Color || goalColor(m.GoalID)) + '"></div>';
      }).join("") + '</div>';
    }

    // Axis (quarter columns + today).
    html += '<div class="rm-row rm-axis-row"><div class="rm-row-label"></div><div class="rm-row-track rm-axis-track">' +
      ["Q1", "Q2", "Q3", "Q4"].map((q) => '<span class="rm-qcol">' + q + '</span>').join("") +
      '<div class="rm-today" style="left:' + todayPct + '%" title="Today"></div></div></div>';

    // Milestone lane (dated diamonds + labels).
    if (State.milestones && State.milestones.length) {
      const flags = State.milestones.slice()
        .sort((a, b) => (isoDate(a.Date) < isoDate(b.Date) ? -1 : 1))
        .map((m) => {
          const d = isoDate(m.Date); if (!d) return "";
          const col = m.Color || goalColor(m.GoalID);
          return '<div class="rm-ms-flag" style="left:' + pctOf(d) + '%;--mc:' + col + '" data-ms-id="' + escapeAttr(m.MilestoneID) +
            '" title="' + escapeAttr((m.Title || "Milestone") + " · " + formatDateShort(d) + (m.Notes ? " · " + m.Notes : "")) + '">' +
            '<span class="rm-ms-diostamp">◆</span><span class="rm-ms-flaglabel">' + escapeHtml(m.Title || "") + '</span></div>';
        }).join("");
      html += '<div class="rm-row rm-ms-row"><div class="rm-row-label rm-ms-rowlabel">◆ Milestones <button class="rm-ms-add" id="rm-add-ms" title="Add a dated milestone line">＋ Add</button></div>' +
        '<div class="rm-row-track">' + flags + '<div class="rm-today" style="left:' + todayPct + '%"></div></div></div>';
    } else {
      html += '<div class="rm-row rm-ms-row"><div class="rm-row-label rm-ms-rowlabel">◆ Milestones <button class="rm-ms-add" id="rm-add-ms" title="Add a dated milestone line">＋ Add</button></div>' +
        '<div class="rm-row-track"><span class="rm-ms-empty">No milestones — add a dated line</span><div class="rm-today" style="left:' + todayPct + '%"></div></div></div>';
    }

    let any = false;
    order.forEach((wid) => {
      const tasks = visible.filter((t) => t.WorkstreamID === wid)
        .sort((a, b) => { const ra = rng(a).s, rb = rng(b).s; return ra < rb ? -1 : ra > rb ? 1 : 0; });
      if (!tasks.length) return;
      any = true;
      const done = tasks.filter((t) => t.Status === "Done").length;
      html += '<div class="rm-group"><span class="rm-group-name">' +
        escapeHtml(workstreamName(wid) || "(no workstream)") + '</span><span class="rm-group-meta">' +
        done + '/' + tasks.length + ' done</span></div>';
      tasks.forEach((t) => {
        const r = rng(t);
        const left = pctOf(r.s), width = Math.max(1.5, pctOf(r.e) - left);
        const pct = Math.max(0, Math.min(100, Number(t.PercentComplete) || 0));
        const gc = goalColor(taskGoalIds(t)[0] || "");
        const hl = t.Health ? " rm-h-" + statusSlug(t.Health) : "";
        const blocked = t.Status === "Blocked" ? " rm-blocked" : "";
        const subs = State.subtasksByParent[t.TaskID] || [];
        const expanded = State.roadmap.expanded.has(Number(t.TaskID));
        const caret = subs.length
          ? '<button class="rm-caret" data-rm-exp="' + t.TaskID + '" title="Subtasks">' + (expanded ? "▾" : "▸") + '</button> '
          : '';
        const bar = '<div class="rm-bar' + hl + blocked + '" style="left:' + left + '%;width:' + width + '%;background:' + gc + '" data-task-id="' + t.TaskID + '">' +
          '<span class="rm-handle rm-handle-l" data-grip="l"></span>' +
          '<span class="rm-bar-fill" style="width:' + pct + '%"></span>' +
          '<span class="rm-bar-label">' + escapeHtml(t.Title || "") + '</span>' +
          '<span class="rm-handle rm-handle-r" data-grip="r"></span></div>';
        html += '<div class="rm-row" data-task-id="' + t.TaskID + '">' +
          '<div class="rm-row-label" title="' + escapeAttr(t.Title) + '">' + caret + scheduleChip(t) + escapeHtml(t.Title || "") + '</div>' +
          '<div class="rm-row-track">' + bar + '<div class="rm-today" style="left:' + todayPct + '%"></div></div></div>';

        // Subtask lane (accordion) — thin bars from parent start to each subtask due.
        if (expanded) {
          subs.slice().sort((a, b) => (Number(a.Order) || 0) - (Number(b.Order) || 0)).forEach((s) => {
            const done2 = String(s.Done).toLowerCase() === "yes";
            const due = isoDate(s.DueDate);
            const sl = left;
            const sr = due ? pctOf(due) : pctOf(r.e);
            const sw = Math.max(1.2, sr - sl);
            const stip = escapeAttr((s.Text || "Subtask") + (due ? " · due " + formatDateShort(due) : "") + (done2 ? " · done" : ""));
            html += '<div class="rm-row rm-subrow"><div class="rm-row-label rm-sublabel" title="' + escapeAttr(s.Text || "") + '">↳ ' +
              escapeHtml(s.Text || "") + '</div><div class="rm-row-track">' +
              '<div class="rm-subbar' + (done2 ? " done" : "") + '" style="left:' + sl + '%;width:' + sw + '%" title="' + stip + '"></div>' +
              '<div class="rm-subdot' + (done2 ? " done" : "") + '" style="left:' + sr + '%" title="' + stip + '"></div>' +
              '<div class="rm-today" style="left:' + todayPct + '%"></div></div></div>';
          });
        }
      });
    });
    html += '</div></div>';
    if (!any) html += '<div class="mine-empty">No items match your filters. Milestones still show above.</div>';
    html += '<div class="rm-hovercard" id="rm-hovercard" hidden></div>';
    html += '<div class="rm-dragtip" id="rm-dragtip" hidden></div>';
    root.innerHTML = html;

    // Date-tooltip toggle.
    const showDatesCb = document.getElementById("rm-show-dates");
    if (showDatesCb) showDatesCb.addEventListener("change", () => { State.roadmap.showDates = showDatesCb.checked; });

    // Add-milestone button.
    const addMs = document.getElementById("rm-add-ms");
    if (addMs) addMs.addEventListener("click", (e) => { e.stopPropagation(); openMilestoneEditor(null); });
    Array.from(root.querySelectorAll(".rm-ms-flag[data-ms-id]")).forEach((f) => {
      f.addEventListener("click", (e) => {
        e.stopPropagation();
        const m = (State.milestones || []).find((x) => String(x.MilestoneID) === String(f.dataset.msId));
        if (m) openMilestoneEditor(m);
      });
    });

    // Row label click → open modal (track clicks are for dragging).
    Array.from(root.querySelectorAll(".rm-row[data-task-id] .rm-row-label")).forEach((lab) => {
      lab.addEventListener("click", (e) => {
        if (e.target.closest(".rm-caret")) return;
        openEditModal(Number(lab.closest(".rm-row").dataset.taskId));
      });
    });
    // Subtask accordion carets.
    Array.from(root.querySelectorAll(".rm-caret[data-rm-exp]")).forEach((b) => {
      b.addEventListener("click", (e) => {
        e.stopPropagation();
        const id = Number(b.dataset.rmExp);
        if (State.roadmap.expanded.has(id)) State.roadmap.expanded.delete(id);
        else State.roadmap.expanded.add(id);
        renderRoadmap();
      });
    });
    // Bars: hover card + drag/resize.
    Array.from(root.querySelectorAll(".rm-bar[data-task-id]")).forEach((bar) => {
      const task = State.tasks.find((x) => Number(x.TaskID) === Number(bar.dataset.taskId));
      if (!task) return;
      wireRoadmapHover(bar, task);
      wireRoadmapDrag(bar, task, { pctOf: pctOf, isoFromPct: isoFromPct });
    });
  }

  // Hover card with the full task detail (the bar label truncates).
  function wireRoadmapHover(bar, task) {
    const card = document.getElementById("rm-hovercard");
    if (!card) return;
    bar.addEventListener("mouseenter", () => {
      if (State._rmDragging) return;
      const owner = String(task.Owner || "").trim();
      const goals = taskGoalIds(task).map(goalShortName).join(", ") || "—";
      const pct = Math.max(0, Math.min(100, Number(task.PercentComplete) || 0));
      card.innerHTML =
        '<div class="rm-hc-title">' + escapeHtml(task.Title || "") + '</div>' +
        '<div class="rm-hc-row"><span>Dates</span><b>' + (task.StartDate ? formatDateShort(task.StartDate) : "?") + ' – ' + (task.DueDate ? formatDateShort(task.DueDate) : "?") + '</b></div>' +
        '<div class="rm-hc-row"><span>Quarter</span><b>' + escapeHtml(task.Quarter || "—") + '</b></div>' +
        '<div class="rm-hc-row"><span>Status</span><b>' + escapeHtml(task.Status || "—") + '</b></div>' +
        (task.Health ? '<div class="rm-hc-row"><span>Health</span><b class="health-' + statusSlug(task.Health) + '-txt">' + escapeHtml(task.Health) + '</b></div>' : "") +
        '<div class="rm-hc-row"><span>Owner</span><b>' + escapeHtml(owner || "—") + '</b></div>' +
        '<div class="rm-hc-row"><span>Goals</span><b>' + escapeHtml(goals) + '</b></div>' +
        '<div class="rm-hc-row"><span>Progress</span><b>' + pct + '%</b></div>';
      card.hidden = false;
      const r = bar.getBoundingClientRect();
      const top = r.bottom + 8;
      card.style.top = Math.min(top, window.innerHeight - 180) + "px";
      card.style.left = Math.max(8, Math.min(r.left, window.innerWidth - 250)) + "px";
    });
    bar.addEventListener("mouseleave", () => { card.hidden = true; });
  }

  // Drag a bar to move it, or its edges to resize — live, with an optional date
  // tooltip. On drop, commit the new dates (and prompt on a quarter change).
  function wireRoadmapDrag(bar, task, ctx) {
    bar.addEventListener("pointerdown", (ev) => {
      if (ev.button !== 0) return;
      ev.preventDefault();
      const grip = ev.target.closest(".rm-handle");
      const mode = grip ? (grip.dataset.grip === "l" ? "resize-l" : "resize-r") : "move";
      const track = bar.parentElement;
      const trackW = track.getBoundingClientRect().width;
      const startX = ev.clientX;
      const left0 = parseFloat(bar.style.left) || 0;
      const width0 = parseFloat(bar.style.width) || 1;
      let moved = false;
      State._rmDragging = false;
      const tip = document.getElementById("rm-dragtip");
      const card = document.getElementById("rm-hovercard"); if (card) card.hidden = true;
      bar.classList.add("rm-bar-dragging");

      function onMove(e) {
        const dPx = e.clientX - startX;
        if (Math.abs(dPx) > 3) { moved = true; State._rmDragging = true; }
        const dPct = (dPx / trackW) * 100;
        let nl = left0, nw = width0;
        if (mode === "move") { nl = Math.max(0, Math.min(100 - width0, left0 + dPct)); }
        else if (mode === "resize-l") { nl = Math.max(0, Math.min(left0 + width0 - 1, left0 + dPct)); nw = (left0 + width0) - nl; }
        else { nw = Math.max(1, Math.min(100 - left0, width0 + dPct)); }
        bar.style.left = nl + "%"; bar.style.width = nw + "%";
        if (State.roadmap.showDates && tip) {
          const ns = ctx.isoFromPct(nl), ne = ctx.isoFromPct(nl + nw);
          tip.textContent = formatDateShort(ns) + " – " + formatDateShort(ne);
          tip.hidden = false;
          tip.style.top = (e.clientY - 34) + "px";
          tip.style.left = Math.min(e.clientX + 12, window.innerWidth - 130) + "px";
        }
      }
      function onUp() {
        document.removeEventListener("pointermove", onMove);
        document.removeEventListener("pointerup", onUp);
        bar.classList.remove("rm-bar-dragging");
        if (tip) tip.hidden = true;
        const nl = parseFloat(bar.style.left) || 0, nw = parseFloat(bar.style.width) || 1;
        setTimeout(() => { State._rmDragging = false; }, 0);   // let the click handler see it
        if (!moved) return;                                    // a plain click — row-label handler opens modal
        commitRoadmapDates(task, ctx.isoFromPct(nl), ctx.isoFromPct(nl + nw));
      }
      document.addEventListener("pointermove", onMove);
      document.addEventListener("pointerup", onUp);
    });
  }

  // Persist new roadmap dates. Prompts (Accelerated / Delayed / just move) when an
  // in-progress task changes quarter; cascades the shift to subtask due dates.
  async function commitRoadmapDates(task, ns, ne) {
    if (ne < ns) { const tmp = ns; ns = ne; ne = tmp; }
    const oldStart = isoDate(task.StartDate) || isoDate(task.DueDate) || ns;
    const oldQ = task.Quarter;
    const newQ = quarterOf(ns);
    const deltaMs = new Date(ns + "T00:00:00").getTime() - new Date(oldStart + "T00:00:00").getTime();
    const order = ["Q1", "Q2", "Q3", "Q4"];
    const movable = task.Status !== "Backlog" && task.Status !== "Done";
    let slipsDelta = 0, action = null;

    if (newQ && oldQ && newQ !== oldQ && movable) {
      const fwd = order.indexOf(newQ) > order.indexOf(oldQ);
      const choice = await uiChoose(
        "Moved " + oldQ + " → " + newQ,
        "This in-progress item changed quarter. How should it be logged?",
        [
          fwd ? { value: "delay", label: "⏩ Delayed — a slip (roll-over)", cls: "btn-loss" }
              : { value: "accel", label: "⏪ Accelerated — a win (pulled in)", cls: "btn-win" },
          { value: "move", label: "Just move it — no schedule flag", cls: "btn-secondary" }
        ]
      );
      if (choice === null) { renderRoadmap(); return; }     // cancelled → snap back
      if (choice === "delay") { slipsDelta = 1; action = "Delayed"; }
      else if (choice === "accel") { slipsDelta = -1; action = "Accelerated"; }
    }

    const updates = { TaskID: task.TaskID, StartDate: ns, DueDate: ne };
    if (newQ) updates.Quarter = newQ;
    if (slipsDelta) updates.Slips = (Number(task.Slips) || 0) + slipsDelta;
    // optimistic local mirror
    task.StartDate = ns; task.DueDate = ne;
    if (newQ) task.Quarter = newQ;
    if (slipsDelta) task.Slips = (Number(task.Slips) || 0) + slipsDelta;

    try {
      await window.WsjfData.writeTask(updates, { force: true, silent: true });
      if (deltaMs) {
        const subs = State.subtasksByParent[task.TaskID] || [];
        for (const s of subs) {
          if (s.SubtaskID && s.DueDate) {
            const nd = shiftIso(s.DueDate, deltaMs);
            try { await window.WsjfData.writeSubtask({ SubtaskID: s.SubtaskID, DueDate: nd }); s.DueDate = nd; } catch (_) {}
          }
        }
      }
      if (action) { try { await window.WsjfData.logActivity("Task", task.TaskID, action, "Quarter", oldQ, newQ, ""); } catch (_) {} }
      State.lastSyncTs = Date.now(); updateSyncLabel();
      toast((action ? action + ": " : "Rescheduled: ") + formatDateShort(ns) + " – " + formatDateShort(ne) +
        (newQ && newQ !== oldQ ? " (" + newQ + ")" : ""), "info");
    } catch (e) {
      toast("Save failed: " + e.message + " — reverting.", "error");
      await reloadTasks();
    }
    renderRoadmap();
  }

  async function reloadMilestones() {
    try { State.milestones = await window.WsjfData._internal._readTable("MilestonesTable"); }
    catch (e) { State.milestones = []; }
  }

  // Create/edit a standalone milestone (a dated line on the roadmap — NOT a task).
  function openMilestoneEditor(m) {
    const isNew = !m;
    m = m || {};
    const host = document.createElement("div");
    host.className = "confirm-overlay";
    const box = document.createElement("div");
    box.className = "confirm-box ms-editor";
    const goalOpts = '<option value="">— no goal —</option>' +
      State.goals.map((g) => '<option value="' + escapeAttr(g.GoalID) + '"' +
        (String(m.GoalID) === String(g.GoalID) ? " selected" : "") + '>' +
        escapeHtml(g.ShortName || g.GoalName || g.GoalID) + '</option>').join("");
    box.innerHTML =
      '<h3 class="confirm-title">' + (isNew ? "Add milestone" : "Edit milestone") + '</h3>' +
      '<label class="ms-fl">Title<input id="ms-title" type="text" value="' + escapeAttr(m.Title || "") + '" placeholder="e.g. Beta launch" /></label>' +
      '<div class="ms-two">' +
        '<label class="ms-fl">Date<input id="ms-date" type="date" value="' + escapeAttr(isoDate(m.Date) || "") + '" /></label>' +
        '<label class="ms-fl">Goal<select id="ms-goal">' + goalOpts + '</select></label>' +
      '</div>' +
      '<label class="ms-fl">Notes<input id="ms-notes" type="text" value="' + escapeAttr(m.Notes || "") + '" placeholder="optional" /></label>' +
      '<div class="confirm-actions">' +
        (isNew ? "" : '<button class="btn btn-archive" id="ms-del">Delete</button>') +
        '<span style="flex:1"></span>' +
        '<button class="btn btn-secondary" id="ms-cancel">Cancel</button>' +
        '<button class="btn btn-primary" id="ms-save">Save</button>' +
      '</div>';
    host.appendChild(box);
    document.body.appendChild(host);
    function close() { host.remove(); }
    host.addEventListener("click", (e) => { if (e.target === host) close(); });
    box.querySelector("#ms-cancel").addEventListener("click", close);
    const delBtn = box.querySelector("#ms-del");
    if (delBtn) delBtn.addEventListener("click", async () => {
      if (!(await uiConfirm('Delete milestone "' + (m.Title || "") + '"?', { okText: "Delete" }))) return;
      try { await window.WsjfData.deleteMilestone(m.MilestoneID); await reloadMilestones(); renderRoadmap(); toast("Milestone deleted.", "info"); close(); }
      catch (e) { toast("Delete failed: " + e.message, "error"); }
    });
    box.querySelector("#ms-save").addEventListener("click", async () => {
      const fields = {
        Title: box.querySelector("#ms-title").value.trim(),
        Date: box.querySelector("#ms-date").value,
        GoalID: box.querySelector("#ms-goal").value,
        Notes: box.querySelector("#ms-notes").value.trim()
      };
      if (!fields.Title) { toast("Title is required.", "warn"); return; }
      if (!fields.Date) { toast("Date is required.", "warn"); return; }
      fields.Quarter = quarterOf(fields.Date);
      const saveBtn = box.querySelector("#ms-save"); saveBtn.disabled = true; saveBtn.textContent = "Saving…";
      try {
        if (isNew) await window.WsjfData.createMilestone(fields);
        else await window.WsjfData.updateMilestone(m.MilestoneID, fields);
        await reloadMilestones();
        renderRoadmap();
        toast("Milestone saved.", "info");
        close();
      } catch (e) {
        saveBtn.disabled = false; saveBtn.textContent = "Save";
        toast("Save failed: " + e.message, "error");
      }
    });
    setTimeout(() => { const ti = box.querySelector("#ms-title"); if (ti) ti.focus(); }, 0);
  }

  // ───── Config screen ─────
  function renderConfig() {
    const root = document.getElementById("configview");
    root.innerHTML = "";

    const intro = document.createElement("div");
    intro.className = "config-intro";
    intro.innerHTML =
      '<h2>Configuration</h2>' +
      '<p class="muted">Add or rename Board columns, Workstreams, Goals, Owners and Quarters. ' +
      'Changes flow through to connected tasks.</p>';
    root.appendChild(intro);

    root.appendChild(configSectionStatuses());

    root.appendChild(configSectionIdName({
      title: "Workstreams",
      items: State.workstreams,
      idField: "WorkstreamID",
      idPrefix: "WS",
      nameOf: (w) => w.Name || w.WorkstreamID,
      note: "Tasks link to a workstream by ID, so renaming updates every task automatically.",
      onAdd: (id, name) => window.WsjfData.createWorkstream({ WorkstreamID: id, Name: name }),
      onRename: (id, name) => window.WsjfData.updateWorkstream(id, { Name: name }),
      onDelete: async (id) => {
        const n = await window.WsjfData.countTasksByWorkstream(id);
        if (n > 0) throw new Error(n + " task(s) use this workstream — reassign them first.");
        await window.WsjfData.deleteWorkstream(id);
      }
    }));

    root.appendChild(configSectionIdName({
      title: "Goals",
      items: State.goals,
      idField: "GoalID",
      idPrefix: "G",
      nameOf: (g) => g.ShortName || g.GoalName || g.GoalID,
      note: "Tasks link to a goal by ID, so renaming updates every task automatically.",
      onAdd: (id, name) => window.WsjfData.createGoal({ GoalID: id, ShortName: name, GoalName: name }),
      onRename: (id, name) => window.WsjfData.updateGoal(id, { ShortName: name, GoalName: name }),
      onDelete: async (id) => {
        const n = await window.WsjfData.countTasksByGoal(id);
        if (n > 0) throw new Error(n + " task(s) use this goal — reassign them first.");
        await window.WsjfData.deleteGoal(id);
      }
    }));

    root.appendChild(configSectionValue({
      title: "Owners",
      items: State.config.Owners,
      note: "Renaming an owner rewrites that owner on every task that has it.",
      onAdd: (v) => window.WsjfData.addConfigValue("OwnersTable", v),
      onRename: (oldV, newV) => window.WsjfData.renameOwner(oldV, newV),
      onDelete: async (v) => {
        const n = await window.WsjfData.countTasksByField("Owner", v, true);
        if (n > 0) throw new Error(n + " task(s) have this owner — rename, or reassign them first.");
        await window.WsjfData.deleteConfigValue("OwnersTable", v);
      }
    }));

    root.appendChild(configSectionValue({
      title: "Quarters",
      items: State.config.Quarters,
      note: "Renaming a quarter rewrites it on every task that uses it.",
      onAdd: (v) => window.WsjfData.addConfigValue("QuartersTable", v),
      onRename: (oldV, newV) => window.WsjfData.renameQuarter(oldV, newV),
      onDelete: async (v) => {
        const n = await window.WsjfData.countTasksByField("Quarter", v, false);
        if (n > 0) throw new Error(n + " task(s) use this quarter — rename, or reassign them first.");
        await window.WsjfData.deleteConfigValue("QuartersTable", v);
      }
    }));
  }

  async function configGuard(fn) {
    try { await fn(); }
    catch (e) { console.error(e); toast((e && e.message) || "Operation failed", "error"); }
  }

  async function afterConfigChange() {
    await loadConfigAndDimensions();
    await reloadTasks();
    render();
  }

  // Board columns (Statuses): rename (cascades to tasks), reorder (remembered),
  // add, and delete-with-reassign (move that column's tasks elsewhere first).
  function configSectionStatuses() {
    const sec = document.createElement("section");
    sec.className = "config-section";
    const cols = boardColumns();
    const h = document.createElement("h3");
    h.textContent = "Board columns (" + cols.length + ")";
    sec.appendChild(h);
    const p = document.createElement("p");
    p.className = "config-note";
    p.textContent = "These are your Kanban columns. Rename cascades to every task. Use ▲▼ to set the column order (left→right). Delete lets you move that column's tasks to another column first.";
    sec.appendChild(p);

    const listEl = document.createElement("div");
    listEl.className = "config-list";
    const counts = {};
    cols.forEach((s) => { counts[s] = State.tasks.filter((t) => t.Status === s).length; });

    cols.forEach((val, idx) => {
      const row = document.createElement("div");
      row.className = "config-row";
      const up = document.createElement("button");
      up.className = "btn btn-secondary btn-sm"; up.textContent = "▲"; up.disabled = idx === 0; up.title = "Move left";
      const down = document.createElement("button");
      down.className = "btn btn-secondary btn-sm"; down.textContent = "▼"; down.disabled = idx === cols.length - 1; down.title = "Move right";
      const input = document.createElement("input"); input.type = "text"; input.value = val;
      const tag = document.createElement("span"); tag.className = "config-id"; tag.textContent = counts[val] + " task" + (counts[val] === 1 ? "" : "s");
      const saveBtn = document.createElement("button"); saveBtn.className = "btn btn-secondary btn-sm"; saveBtn.textContent = "Rename";
      const delBtn = document.createElement("button"); delBtn.className = "btn btn-archive btn-sm"; delBtn.textContent = "Delete";

      const reorder = (from, to) => {
        const arr = cols.slice();
        const [m] = arr.splice(from, 1);
        arr.splice(to, 0, m);
        saveBoardColOrder(arr);
        renderConfig();
        renderBoardIfActive();
      };
      up.addEventListener("click", () => reorder(idx, idx - 1));
      down.addEventListener("click", () => reorder(idx, idx + 1));

      saveBtn.addEventListener("click", async () => {
        const newV = input.value.trim();
        if (!newV || newV === val) return;
        if (!(await uiConfirm('Rename column "' + val + '" → "' + newV + '"? This updates every task with that status.', { okText: "Rename" }))) return;
        configGuard(async () => {
          const n = await window.WsjfData.renameStatus(val, newV);
          // keep the saved order in sync with the new name
          const arr = boardColumns().map((s) => (s === val ? newV : s));
          saveBoardColOrder(arr);
          toast("Renamed" + (typeof n === "number" ? " (" + n + " task" + (n === 1 ? "" : "s") + " moved)" : "") + ".", "info");
          await afterConfigChange();
        });
      });

      delBtn.addEventListener("click", async () => {
        const n = counts[val];
        if (n > 0) {
          const others = cols.filter((s) => s !== val);
          if (!others.length) { toast("Can't delete the only column.", "warn"); return; }
          const target = await uiChoose(
            'Delete "' + val + '" (' + n + ' task' + (n === 1 ? "" : "s") + ')',
            "Move its tasks to another column, then delete it:",
            others.map((s) => ({ value: s, label: "Move to “" + s + "”", cls: "btn-secondary" }))
          );
          if (!target) return;
          configGuard(async () => {
            toast("Moving " + n + " task" + (n === 1 ? "" : "s") + "…", "info");
            for (const t of State.tasks.filter((x) => x.Status === val)) {
              await window.WsjfData.writeTask({ TaskID: t.TaskID, Status: target }, { force: true, silent: true });
            }
            await window.WsjfData.deleteConfigValue("StatusesTable", val);
            saveBoardColOrder(boardColumns().filter((s) => s !== val));
            toast("Moved " + n + " to “" + target + "” and deleted “" + val + "”.", "info");
            await afterConfigChange();
          });
        } else {
          if (!(await uiConfirm('Delete empty column "' + val + '"?', { okText: "Delete" }))) return;
          configGuard(async () => {
            await window.WsjfData.deleteConfigValue("StatusesTable", val);
            saveBoardColOrder(boardColumns().filter((s) => s !== val));
            toast("Deleted “" + val + "”.", "info");
            await afterConfigChange();
          });
        }
      });

      row.appendChild(up); row.appendChild(down);
      row.appendChild(input); row.appendChild(tag);
      row.appendChild(saveBtn); row.appendChild(delBtn);
      listEl.appendChild(row);
    });
    sec.appendChild(listEl);

    const addRow = document.createElement("div"); addRow.className = "config-row config-add";
    const addInput = document.createElement("input"); addInput.type = "text"; addInput.placeholder = "New column (e.g. Planned)…";
    const addBtn = document.createElement("button"); addBtn.className = "btn btn-primary btn-sm"; addBtn.textContent = "Add column";
    const doAdd = () => {
      const v = addInput.value.trim();
      if (!v) return;
      configGuard(async () => {
        await window.WsjfData.addConfigValue("StatusesTable", v);
        saveBoardColOrder(boardColumns().concat([v]));   // new column lands on the right; reorder with ▲
        toast('Added “' + v + '”. Use ▲ to position it.', "info");
        await afterConfigChange();
      });
    };
    addBtn.addEventListener("click", doAdd);
    addInput.addEventListener("keydown", (e) => { if (e.key === "Enter") doAdd(); });
    addRow.appendChild(addInput); addRow.appendChild(addBtn);
    sec.appendChild(addRow);
    return sec;
  }

  // Re-render the board only if it's the active view (used after a column reorder
  // so the change is visible immediately behind the Config screen too).
  function renderBoardIfActive() {
    if (State.viewMode === "board") renderBoard();
  }

  function configSectionValue(opts) {
    const sec = document.createElement("section");
    sec.className = "config-section";
    const h = document.createElement("h3");
    h.textContent = opts.title + " (" + opts.items.length + ")";
    sec.appendChild(h);
    if (opts.note) {
      const p = document.createElement("p"); p.className = "config-note"; p.textContent = opts.note; sec.appendChild(p);
    }
    const listEl = document.createElement("div"); listEl.className = "config-list";
    opts.items.slice().sort((a, b) => String(a).localeCompare(String(b))).forEach((val) => {
      const row = document.createElement("div"); row.className = "config-row";
      const input = document.createElement("input"); input.type = "text"; input.value = val;
      const saveBtn = document.createElement("button"); saveBtn.className = "btn btn-secondary btn-sm"; saveBtn.textContent = "Save";
      const delBtn = document.createElement("button"); delBtn.className = "btn btn-archive btn-sm"; delBtn.textContent = "Delete";
      saveBtn.addEventListener("click", async () => {
        const newV = input.value.trim();
        if (!newV || newV === val) return;
        if (!(await uiConfirm('Rename "' + val + '" to "' + newV + '"? This updates connected tasks.', { okText: "Rename" }))) return;
        configGuard(async () => {
          const n = await opts.onRename(val, newV);
          toast("Renamed" + (typeof n === "number" ? " (" + n + " task" + (n === 1 ? "" : "s") + " updated)" : "") + ".", "info");
          await afterConfigChange();
        });
      });
      delBtn.addEventListener("click", async () => {
        if (!(await uiConfirm('Delete "' + val + '"?', { okText: "Delete" }))) return;
        configGuard(async () => { await opts.onDelete(val); toast("Deleted.", "info"); await afterConfigChange(); });
      });
      row.appendChild(input); row.appendChild(saveBtn); row.appendChild(delBtn);
      listEl.appendChild(row);
    });
    sec.appendChild(listEl);

    const addRow = document.createElement("div"); addRow.className = "config-row config-add";
    const addInput = document.createElement("input"); addInput.type = "text";
    addInput.placeholder = "New " + opts.title.replace(/s$/, "") + "…";
    const addBtn = document.createElement("button"); addBtn.className = "btn btn-primary btn-sm"; addBtn.textContent = "Add";
    addBtn.addEventListener("click", () => {
      const v = addInput.value.trim();
      if (!v) return;
      configGuard(async () => { await opts.onAdd(v); toast("Added.", "info"); await afterConfigChange(); });
    });
    addRow.appendChild(addInput); addRow.appendChild(addBtn);
    sec.appendChild(addRow);
    return sec;
  }

  function configSectionIdName(opts) {
    const sec = document.createElement("section"); sec.className = "config-section";
    const h = document.createElement("h3"); h.textContent = opts.title + " (" + opts.items.length + ")"; sec.appendChild(h);
    if (opts.note) { const p = document.createElement("p"); p.className = "config-note"; p.textContent = opts.note; sec.appendChild(p); }
    const listEl = document.createElement("div"); listEl.className = "config-list";
    opts.items.slice().forEach((item) => {
      const id = item[opts.idField];
      const row = document.createElement("div"); row.className = "config-row";
      const idTag = document.createElement("span"); idTag.className = "config-id"; idTag.textContent = id;
      const input = document.createElement("input"); input.type = "text"; input.value = opts.nameOf(item);
      const saveBtn = document.createElement("button"); saveBtn.className = "btn btn-secondary btn-sm"; saveBtn.textContent = "Save";
      const delBtn = document.createElement("button"); delBtn.className = "btn btn-archive btn-sm"; delBtn.textContent = "Delete";
      saveBtn.addEventListener("click", () => {
        const name = input.value.trim();
        if (!name || name === opts.nameOf(item)) return;
        configGuard(async () => { await opts.onRename(id, name); toast("Renamed.", "info"); await afterConfigChange(); });
      });
      delBtn.addEventListener("click", async () => {
        if (!(await uiConfirm('Delete "' + opts.nameOf(item) + '"?', { okText: "Delete" }))) return;
        configGuard(async () => { await opts.onDelete(id); toast("Deleted.", "info"); await afterConfigChange(); });
      });
      row.appendChild(idTag); row.appendChild(input); row.appendChild(saveBtn); row.appendChild(delBtn);
      listEl.appendChild(row);
    });
    sec.appendChild(listEl);

    // ID is auto-generated (next in sequence) — users only type a name.
    const addRow = document.createElement("div"); addRow.className = "config-row config-add";
    const nameInput = document.createElement("input"); nameInput.type = "text";
    nameInput.placeholder = "New " + opts.title.replace(/s$/, "") + " name…";
    const addBtn = document.createElement("button"); addBtn.className = "btn btn-primary btn-sm"; addBtn.textContent = "Add";
    const doAdd = () => {
      const name = nameInput.value.trim();
      if (!name) { toast("Name required", "warn"); return; }
      const id = nextSequentialId(opts.items, opts.idField, opts.idPrefix);
      configGuard(async () => { await opts.onAdd(id, name); toast("Added " + id + ".", "info"); await afterConfigChange(); });
    };
    addBtn.addEventListener("click", doAdd);
    nameInput.addEventListener("keydown", (e) => { if (e.key === "Enter") doAdd(); });
    addRow.appendChild(nameInput); addRow.appendChild(addBtn);
    sec.appendChild(addRow);
    return sec;
  }

  // Auto-generate the next ID in a "PREFIX + zero-padded number" sequence
  // (e.g. WS01..WS10 -> WS11, G1..G5 -> G6) so users never type IDs.
  function nextSequentialId(items, idField, fallbackPrefix) {
    let prefix = fallbackPrefix || "ID", width = 1, max = 0;
    (items || []).forEach((it) => {
      const m = String(it[idField] || "").match(/^([A-Za-z]+)(\d+)$/);
      if (m) { prefix = m[1]; width = Math.max(width, m[2].length); const n = Number(m[2]); if (n > max) max = n; }
    });
    return prefix + String(max + 1).padStart(width, "0");
  }

  // ───── UI bindings ─────
  function bindUi() {
    const backBtn = document.getElementById("back-to-stats");
    const inDialog = /[?&]host=dialog/.test(location.search);
    if (backBtn) {
      if (inDialog) backBtn.hidden = true;   // dialog has its own close (X); no "back to stats"
      else backBtn.addEventListener("click", goBackToStats);
    }

    const modeToggle = document.getElementById("view-mode-toggle");
    if (modeToggle) modeToggle.addEventListener("click", (e) => {
      const btn = e.target.closest("button");
      if (!btn) return;
      State.viewMode = btn.dataset.mode;
      Array.from(modeToggle.querySelectorAll("button")).forEach((b) =>
        b.classList.toggle("active", b === btn));
      render();
    });

    const idChip = document.getElementById("identity-chip");
    if (idChip) idChip.addEventListener("click", async () => { await openIdentityPicker(); render(); });

    // Settings gear → Config view (Config is no longer a tab)
    const cfgBtn = document.getElementById("config-btn");
    if (cfgBtn) cfgBtn.addEventListener("click", () => {
      State.viewMode = State.viewMode === "config" ? "board" : "config";
      render();
    });

    document.getElementById("filter-owner").addEventListener("change", (e) => {
      State.filters.owner = e.target.value; render();
    });
    document.getElementById("filter-workstream").addEventListener("change", (e) => {
      State.filters.workstream = e.target.value; render();
    });
    document.getElementById("filter-quarter").addEventListener("change", (e) => {
      State.filters.quarter = e.target.value; render();
    });
    document.getElementById("filter-subtasks").addEventListener("change", (e) => {
      State.filters.subtasks = e.target.value; render();
    });
    document.getElementById("filter-goal").addEventListener("change", (e) => {
      State.filters.goal = e.target.value; render();
    });
    document.getElementById("filter-health").addEventListener("change", (e) => {
      State.filters.health = e.target.value; render();
    });
    document.getElementById("filter-status").addEventListener("change", (e) => {
      State.filters.status = e.target.value; render();
    });
    [["filter-start-from", "startFrom"], ["filter-start-to", "startTo"],
     ["filter-due-from", "dueFrom"], ["filter-due-to", "dueTo"]].forEach(([id, key]) => {
      document.getElementById(id).addEventListener("change", (e) => {
        State.filters[key] = e.target.value; render();
      });
    });
    document.getElementById("group-toggle").addEventListener("change", (e) => {
      State.groupBy = e.target.value; render();
    });

    // Filters popover toggle + outside-click close
    const filterBtn = document.getElementById("filter-btn");
    const filterPanel = document.getElementById("filter-panel");
    const filterWrap = filterBtn ? filterBtn.closest(".filter-wrap") : null;
    if (filterBtn && filterPanel) {
      filterBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        filterPanel.hidden = !filterPanel.hidden;
      });
      document.addEventListener("mousedown", (e) => {
        if (!filterPanel.hidden && filterWrap && !filterWrap.contains(e.target)) filterPanel.hidden = true;
      });
    }
    const filterDone = document.getElementById("filter-done");
    if (filterDone) filterDone.addEventListener("click", () => { filterPanel.hidden = true; });
    const filterClear = document.getElementById("filter-clear");
    if (filterClear) filterClear.addEventListener("click", clearAllFilters);

    document.getElementById("search-box").addEventListener("input", debounce((e) => {
      State.filters.search = e.target.value;
      render();
    }, 200));

    document.getElementById("sort-toggle").addEventListener("change", (e) => {
      State.sort = e.target.value; render();
    });

    document.getElementById("refresh-btn").addEventListener("click", async () => {
      try {
        await reloadTasks();
        render();
        toast("Refreshed.", "info");
      } catch (e) {
        toast("Refresh failed: " + e.message, "error");
      }
    });

    document.getElementById("add-task-btn").addEventListener("click", openCreateModal);

    // Tag filter dropdown
    const tagWrap = document.getElementById("filter-tags");
    tagWrap.addEventListener("click", (e) => {
      const menu = document.getElementById("tag-multi-menu");
      if (e.target.closest(".tag-multi-item")) return;
      menu.hidden = !menu.hidden;
    });
    document.addEventListener("click", (e) => {
      if (!tagWrap.contains(e.target)) document.getElementById("tag-multi-menu").hidden = true;
    });

    // Modal bindings (created once, populated each open)
    bindModalUi();
  }

  function refreshTagFilterMenu() {
    const tagsSet = new Set();
    State.tasks.forEach((t) => {
      String(t.Tags || "").split(";").map((s) => s.trim()).filter(Boolean)
        .forEach((tg) => tagsSet.add(tg));
    });
    const menu = document.getElementById("tag-multi-menu");
    menu.innerHTML = "";
    Array.from(tagsSet).sort().forEach((tg) => {
      const item = document.createElement("div");
      item.className = "tag-multi-item";
      const checked = State.selectedTags.has(tg);
      item.innerHTML = `<label><input type="checkbox" ${checked ? "checked" : ""} /> ${escapeHtml(tg)}</label>`;
      item.querySelector("input").addEventListener("change", (e) => {
        if (e.target.checked) State.selectedTags.add(tg);
        else State.selectedTags.delete(tg);
        render();
      });
      menu.appendChild(item);
    });
    if (tagsSet.size === 0) {
      menu.innerHTML = `<div class="tag-multi-empty">No tags yet</div>`;
    }
  }

  // ───── Active-filter chips (always-visible "what am I looking at") ─────
  function clearAllFilters() {
    const f = State.filters;
    f.owner = ""; f.workstream = ""; f.quarter = ""; f.status = ""; f.goal = ""; f.health = ""; f.subtasks = "";
    f.startFrom = ""; f.startTo = ""; f.dueFrom = ""; f.dueTo = ""; f.search = "";
    State.selectedTags.clear();
    State.listColFilters = {};        // also clear the List per-column filters
    syncFilterControls();
    render();
  }

  // Push filter state back into the form controls so the popover always reflects
  // the active filters (e.g. after a drill-down or a chip removal).
  function syncFilterControls() {
    const f = State.filters;
    const set = (id, val) => { const el = document.getElementById(id); if (el) el.value = val; };
    set("filter-owner", f.owner);
    set("filter-workstream", f.workstream);
    set("filter-quarter", f.quarter);
    set("filter-status", f.status);
    set("filter-goal", f.goal);
    set("filter-health", f.health);
    set("filter-subtasks", f.subtasks);
    set("filter-start-from", f.startFrom); set("filter-start-to", f.startTo);
    set("filter-due-from", f.dueFrom); set("filter-due-to", f.dueTo);
    set("search-box", f.search);
    refreshTagFilterMenu();
  }

  function goalShort(id) {
    const g = State.goals.find((x) => String(x.GoalID) === String(id));
    return g ? (g.ShortName || g.GoalName || g.GoalID) : id;
  }

  function renderActiveFilters() {
    const bar = document.getElementById("active-filters");
    if (!bar) return;
    const f = State.filters;
    const chips = [];
    // Whole-view filters
    if (f.search)     chips.push({ label: 'Search: "' + f.search + '"', clear: () => { f.search = ""; } });
    if (f.owner)      chips.push({ label: "Owner: " + f.owner, clear: () => { f.owner = ""; } });
    if (f.workstream) chips.push({ label: "Workstream: " + workstreamName(f.workstream), clear: () => { f.workstream = ""; } });
    if (f.goal)       chips.push({ label: "Goal: " + goalShort(f.goal), clear: () => { f.goal = ""; } });
    if (f.quarter)    chips.push({ label: "Quarter: " + f.quarter, clear: () => { f.quarter = ""; } });
    if (f.status)     chips.push({ label: "Status: " + f.status, clear: () => { f.status = ""; } });
    if (f.health)     chips.push({ label: "Health: " + f.health, clear: () => { f.health = ""; } });
    if (f.startFrom)  chips.push({ label: "Start ≥ " + f.startFrom, clear: () => { f.startFrom = ""; } });
    if (f.startTo)    chips.push({ label: "Start ≤ " + f.startTo, clear: () => { f.startTo = ""; } });
    if (f.dueFrom)    chips.push({ label: "Due ≥ " + f.dueFrom, clear: () => { f.dueFrom = ""; } });
    if (f.dueTo)      chips.push({ label: "Due ≤ " + f.dueTo, clear: () => { f.dueTo = ""; } });
    if (f.subtasks)   chips.push({ label: f.subtasks === "has" ? "Has subtasks" : "No subtasks", clear: () => { f.subtasks = ""; } });
    Array.from(State.selectedTags).forEach((tg) => chips.push({ label: "Tag: " + tg, clear: () => State.selectedTags.delete(tg) }));

    // Per-column (List) filters render as the SAME chips, only in List view.
    if (State.viewMode === "list") {
      Object.keys(State.listColFilters).forEach((key) => {
        const col = LIST_COLS.find((c) => c.key === key);
        const vals = Array.from(State.listColFilters[key]);
        const shown = vals.slice(0, 2).map((v) => v === "" ? "(blank)" : v).join(", ") + (vals.length > 2 ? " +" + (vals.length - 2) : "");
        chips.push({ label: (col ? col.label : key) + ": " + shown, clear: () => { delete State.listColFilters[key]; } });
      });
    }

    // Badge on the Filters button counts panel filters (not search, not column filters).
    const panelCount = ["owner", "workstream", "quarter", "status", "goal", "health", "subtasks", "startFrom", "startTo", "dueFrom", "dueTo"]
      .reduce((n, k) => n + (f[k] ? 1 : 0), 0) + State.selectedTags.size;
    const badge = document.getElementById("filter-count");
    if (badge) { if (panelCount > 0) { badge.textContent = panelCount; badge.hidden = false; } else badge.hidden = true; }

    if (!chips.length) { bar.hidden = true; bar.innerHTML = ""; return; }
    bar.hidden = false;
    bar.innerHTML = '<span class="af-label">Showing:</span>';
    chips.forEach((c) => {
      const chip = document.createElement("button");
      chip.className = "af-chip"; chip.type = "button";
      chip.innerHTML = escapeHtml(c.label) + ' <span class="af-x">×</span>';
      chip.addEventListener("click", () => { c.clear(); syncFilterControls(); render(); });
      bar.appendChild(chip);
    });
    const clearAll = document.createElement("button");
    clearAll.className = "af-clear"; clearAll.type = "button"; clearAll.textContent = "Clear all";
    clearAll.addEventListener("click", clearAllFilters);
    bar.appendChild(clearAll);
  }

  // ───── Polling / sync display ─────
  function startPolling() {
    State.pollTimer = setInterval(async () => {
      // Don't poll while the user is mid-action or a write is settling — otherwise
      // stale server data can snap a just-moved card back to its old spot.
      if (State.modalOpen || State.dragging || State.pendingWrites > 0 ||
          (Date.now() - State.lastWriteTs < 6000)) return;
      try {
        const fresh = await window.WsjfData.readAllTasks();
        diffAndAnimate(State.tasks, fresh);
        State.tasks = fresh;
        State.lastSyncTs = Date.now();
        updateSyncLabel();
        refreshTagFilterMenu();
      } catch (e) {
        console.warn("Poll failed", e);
      }
    }, POLL_MS);

    setInterval(updateSyncLabel, 5000);
  }

  function diffAndAnimate(prev, next) {
    const prevMap = new Map(prev.map((t) => [Number(t.TaskID), t]));
    const nextMap = new Map(next.map((t) => [Number(t.TaskID), t]));

    let changed = prev.length !== next.length;
    nextMap.forEach((t, id) => {
      const p = prevMap.get(id);
      if (!p) { changed = true; flashTaskNew(id); }
      else if (p.Status !== t.Status) { changed = true; }
      else if (Object.keys(t).some((k) => k !== "_rowIndex" && p[k] !== t[k])) { changed = true; flashTask(id); }
    });
    prevMap.forEach((t, id) => { if (!nextMap.has(id)) { changed = true; fadeOutTask(id); } });

    // Only rebuild when something actually changed (and the user isn't busy) —
    // a no-op re-render every poll was part of the "pop in/out" jank.
    if (changed && !State.dragging && !State.modalOpen) setTimeout(render, 80);
  }

  function flashTask(id) {
    const el = document.querySelector(`.kcard[data-task-id="${id}"]`);
    if (!el) return;
    el.classList.add("flash");
    setTimeout(() => el.classList.remove("flash"), 1100);
  }
  function flashTaskNew(id) { /* will appear on re-render with fade-in via CSS */ }
  function fadeOutTask(id) {
    const el = document.querySelector(`.kcard[data-task-id="${id}"]`);
    if (el) el.classList.add("fade-out");
  }

  function updateSyncLabel() {
    const el = document.getElementById("last-sync");
    if (!State.lastSyncTs) { el.textContent = "Updated —"; return; }
    const s = Math.floor((Date.now() - State.lastSyncTs) / 1000);
    let label = "Updated " + (s < 60 ? `${s}s ago` : s < 3600 ? `${Math.floor(s/60)}m ago` : `${Math.floor(s/3600)}h ago`);
    el.textContent = label;
  }

  // ───── Modal ─────
  let currentEditingTask = null; // working copy
  let originalTaskSnapshot = null;
  let currentSubtasks = [];
  let currentAttachments = [];
  let currentTagList = []; // for autocomplete
  let modalDirty = false;
  let pendingSubtaskToggleTimers = {};

  function bindModalUi() {
    document.getElementById("m-close").addEventListener("click", closeModalAsk);
    document.getElementById("m-cancel").addEventListener("click", closeModalAsk);
    document.getElementById("m-save").addEventListener("click", saveModal);
    document.getElementById("m-archive").addEventListener("click", archiveFromModal);
    document.getElementById("modal-overlay").addEventListener("click", (e) => {
      if (e.target.id === "modal-overlay") closeModalAsk();
    });

    // WSJF live recompute (selects fire change; input covers number-like UAs)
    ["m-ubv","m-tc","m-rr","m-js"].forEach((id) => {
      const handler = () => { recomputeWsjfDisplay(); if (id === "m-js") updateCapacityReadout(); markDirty(); };
      document.getElementById(id).addEventListener("input", handler);
      document.getElementById(id).addEventListener("change", handler);
    });
    // Right-rail Updates/Activity toggle
    Array.from(document.querySelectorAll("#m-rail-tabs button")).forEach((b) => {
      b.addEventListener("click", () => switchRailTab(b.dataset.rail));
    });
    // Subtasks collapsible (under description)
    document.getElementById("m-subtasks-toggle").addEventListener("click", toggleSubtasksBlock);
    document.getElementById("m-quarter").addEventListener("change", () => { updateCapacityReadout(); markDirty(); });
    document.getElementById("m-pct").addEventListener("input", (e) => {
      document.getElementById("m-pct-readout").textContent = e.target.value;
      markDirty();
    });

    // Subtask add → opens the subtask modal in create mode
    document.getElementById("m-add-subtask").addEventListener("click", () => openSubtaskModal(null));

    // Attachments
    document.getElementById("m-add-att").addEventListener("click", () => {
      document.getElementById("m-att-form").hidden = false;
      document.getElementById("m-att-label").focus();
    });
    document.getElementById("m-att-cancel").addEventListener("click", () => {
      document.getElementById("m-att-form").hidden = true;
      clearAttForm();
    });
    document.getElementById("m-att-save").addEventListener("click", saveNewAttachment);

    // Post an update (chronological note on the task)
    document.getElementById("m-update-add").addEventListener("click", async () => {
      const ta = document.getElementById("m-update-text");
      const text = ta.value.trim();
      if (!text) return;
      if (!currentEditingTask || !currentEditingTask.TaskID) { toast("Save the task first, then add updates.", "warn"); return; }
      try {
        await window.WsjfData.createUpdate({ ParentType: "Task", ParentID: currentEditingTask.TaskID, Text: text });
        ta.value = "";
        renderUpdates(await window.WsjfData.readUpdatesForParent("Task", currentEditingTask.TaskID));
        toast("Update posted.", "info");
      } catch (e) {
        toast("Couldn't post update — add an 'UpdatesTable' to the workbook. (" + e.message + ")", "error");
      }
    });

    // Tags input
    const tagInput = document.getElementById("m-tag-input");
    tagInput.addEventListener("keydown", handleTagKey);
    tagInput.addEventListener("input", showTagSuggestions);
    tagInput.addEventListener("blur", () => setTimeout(() => {
      document.getElementById("m-tag-suggest").hidden = true;
    }, 150));

    // Blocked-by picker
    document.getElementById("m-blocked-add").addEventListener("click", openBlockedPicker);

    // Auto % from subtasks toggle
    document.getElementById("m-autopct").addEventListener("change", () => { refreshAutoPctUI(); markDirty(); });

    // Owner change re-renders contributors (owner can't also be a contributor)
    document.getElementById("m-owner").addEventListener("change", () => { renderContributors(currentContributors()); updateCapacityReadout(); });

    // Pull earlier / Slip later (immediate, logged)
    document.getElementById("m-pull-earlier").addEventListener("click", () => modalReschedule("earlier"));
    document.getElementById("m-slip-later").addEventListener("click", () => modalReschedule("later"));

    // Generic dirty tracking
    ["m-title","m-description","m-owner","m-workstream","m-goal","m-quarter","m-status","m-due","m-start"]
      .forEach((id) => document.getElementById(id).addEventListener("change", markDirty));
    document.getElementById("m-title").addEventListener("input", markDirty);
    document.getElementById("m-description").addEventListener("input", markDirty);
  }

  function markDirty() { modalDirty = true; }

  // Right-rail feed toggle (Updates / Activity).
  function switchRailTab(rail) {
    Array.from(document.querySelectorAll("#m-rail-tabs button")).forEach((b) => b.classList.toggle("active", b.dataset.rail === rail));
    Array.from(document.querySelectorAll(".mb-rail .rail-pane")).forEach((p) => { p.hidden = p.dataset.railpane !== rail; });
  }
  function setSubtasksExpanded(open) {
    const body = document.getElementById("m-subtasks-body");
    const caret = document.getElementById("m-subtasks-caret");
    if (!body || !caret) return;
    body.hidden = !open;
    caret.textContent = open ? "▾" : "▸";
  }
  function toggleSubtasksBlock() {
    const body = document.getElementById("m-subtasks-body");
    if (body) setSubtasksExpanded(body.hidden);
  }

  // Health as clickable flip pills (writes the hidden #m-health input).
  const HEALTH_OPTS = [["", "—"], ["On Track", "🟢 On Track"], ["At Risk", "🟡 At Risk"], ["Off Track", "🔴 Off Track"]];
  function renderHealthFlip() {
    const wrap = document.getElementById("m-health-flip");
    const cur = document.getElementById("m-health").value || "";
    wrap.innerHTML = "";
    HEALTH_OPTS.forEach(([val, label]) => {
      const b = document.createElement("button");
      b.type = "button";
      b.className = "hp" + (cur === val ? " active" : "") + (val ? " hp-" + statusSlug(val) : "");
      b.textContent = label;
      b.addEventListener("click", () => {
        document.getElementById("m-health").value = val;
        renderHealthFlip();
        markDirty();
      });
      wrap.appendChild(b);
    });
  }

  async function openEditModal(taskId) {
    State.modalOpen = true;
    try {
      const t = await window.WsjfData.readTaskById(taskId);
      if (!t) { toast("Task not found.", "error"); return; }
      const [subs, atts, acts] = await Promise.all([
        window.WsjfData.readSubtasksForTask(taskId),
        window.WsjfData.readAttachmentsForTask(taskId),
        window.WsjfData.readActivityForTask(taskId, 10)
      ]);
      State.subtasksByParent[taskId] = subs;
      State.attsByParent[taskId] = atts;
      State.activityByParent[taskId] = acts;

      currentEditingTask = Object.assign({}, t, { _loadedLastUpdated: t.LastUpdated });
      originalTaskSnapshot = Object.assign({}, t);
      currentSubtasks = subs.slice();
      currentAttachments = atts.slice();
      currentTagList = collectAllTags();
      modalDirty = false;
      populateModal();
      renderSubtasks();
      renderAttachments();
      renderActivity(acts);
      renderUpdates([]);
      loadUpdates(taskId);   // async; UpdatesTable may not exist yet
      showModal();
    } catch (err) {
      console.error(err);
      toast("Could not open task: " + err.message, "error");
    }
  }

  async function loadUpdates(taskId) {
    try { renderUpdates(await window.WsjfData.readUpdatesForParent("Task", taskId)); }
    catch (e) { renderUpdates([]); }   // no UpdatesTable → just show empty
  }

  function renderUpdates(list) {
    const ul = document.getElementById("m-updates-list");
    if (!ul) return;
    ul.innerHTML = "";
    if (!list || !list.length) { ul.innerHTML = '<div class="updates-empty">No updates yet.</div>'; }
    else list.forEach((u) => {
      const div = document.createElement("div");
      div.className = "update-item";
      div.innerHTML = '<div class="update-text">' + escapeHtml(u.Text || "") + '</div>' +
        '<div class="update-meta">' + escapeHtml(u.AddedBy || "?") + ' · ' +
        escapeHtml(formatDateShort(u.AddedDate)) + ' · ' + relTime(u.AddedDate) + '</div>';
      ul.appendChild(div);
    });
    const c = document.getElementById("m-updates-count");
    if (c) c.textContent = (list && list.length) ? "(" + list.length + ")" : "";
  }

  async function openCreateModal() {
    State.modalOpen = true;
    const me = State.me;
    // New tasks inherit whatever you're currently filtered to, so adding a task
    // in a filtered view drops it into that context. With filters cleared there's
    // no default — except Quarter (current) and Due date (that quarter's end).
    const q = State.filters.quarter || currentQuarter();
    currentEditingTask = {
      TaskID: null,
      Title: "",
      WorkstreamID: State.filters.workstream || "",
      GoalID: "",
      Owner: State.filters.owner || "",
      Quarter: q,
      Status: "Backlog",
      LinkedEpic: "",
      UserBusinessValue: 8, TimeCriticality: 5, RiskReduction: 5,
      JobSize: 5,
      PercentComplete: 0,
      ColumnOrder: maxColumnOrderForStatus("Backlog") + 1,
      DueDate: quarterEndDate(q),
      Tags: Array.from(State.selectedTags).join("; "),
      BlockedByTaskIDs: "",
      Description: "",
      Source: "Manual",
      Archived: "No",
      CreatedDate: new Date().toISOString().substring(0, 10),
      CreatedBy: me.name
    };
    originalTaskSnapshot = Object.assign({}, currentEditingTask);
    currentSubtasks = [];
    currentAttachments = [];
    currentTagList = collectAllTags();
    modalDirty = true; // new task — needs save
    populateModal();
    renderSubtasks();
    renderAttachments();
    renderActivity([]);
    renderUpdates([]);
    showModal();
    setTimeout(() => document.getElementById("m-title").focus(), 50);
  }

  function showModal() {
    document.getElementById("modal-overlay").hidden = false;
  }
  function hideModal() {
    document.getElementById("modal-overlay").hidden = true;
    State.modalOpen = false;
    currentEditingTask = null;
  }
  async function closeModalAsk() {
    if (modalDirty) {
      if (!(await uiConfirm("Discard unsaved changes?", { okText: "Discard" }))) return;
    }
    hideModal();
  }

  function populateModal() {
    const t = currentEditingTask;

    document.getElementById("m-title").value = t.Title || "";
    document.getElementById("m-description").value = t.Description || "";
    fillFib("m-ubv", t.UserBusinessValue, 8);
    fillFib("m-tc",  t.TimeCriticality, 5);
    fillFib("m-rr",  t.RiskReduction, 5);
    fillFib("m-js",  t.JobSize, 5, true);
    document.getElementById("m-pct").value = t.PercentComplete || 0;
    document.getElementById("m-pct-readout").textContent = t.PercentComplete || 0;
    document.getElementById("m-start").value = isoDate(t.StartDate);
    document.getElementById("m-due").value = isoDate(t.DueDate);
    // Auto-% on by default only when the stored % already matches the subtask
    // ratio (so a manually-set % isn't silently overwritten on open+save).
    (function () {
      const n = currentSubtasks.length;
      const done = currentSubtasks.filter((s) => String(s.Done).toLowerCase() === "yes").length;
      const ratio = n ? Math.round((done / n) * 100) : null;
      document.getElementById("m-autopct").checked = (ratio != null && Number(t.PercentComplete || 0) === ratio);
    })();
    document.getElementById("m-epic").textContent = t.LinkedEpic || "—";
    document.getElementById("m-source").textContent = t.Source || "—";

    // Preserve the stored owner even if it's not in the Owners list (e.g. a
    // legacy multi-owner "A; B" value) so opening+saving never wipes it.
    const ownerOpts = ["", ...State.config.Owners];
    if (t.Owner && ownerOpts.indexOf(t.Owner) < 0) ownerOpts.push(t.Owner);
    fillSelect("m-owner", ownerOpts, "— Unassigned —", t.Owner);
    renderContributors(String(t.Contributors || "").split(/[;,]/).map((x) => x.trim()).filter(Boolean));
    fillSelect("m-workstream",
      [{ value: "", label: "— No workstream —" }].concat(
        State.workstreams.map((w) => ({ value: w.WorkstreamID, label: w.Name || w.WorkstreamID }))),
      null, t.WorkstreamID);
    renderTaskGoals(String(t.GoalID || "").split(/[;,]/).map((x) => x.trim()).filter(Boolean));
    fillSelect("m-quarter", State.config.Quarters, null, t.Quarter);
    fillSelect("m-status", State.config.Statuses, null, t.Status);
    document.getElementById("m-health").value = t.Health || "";
    renderHealthFlip();
    updateCapacityReadout();
    const reschedRow = document.getElementById("m-reschedule");
    if (reschedRow) reschedRow.hidden = !t.TaskID;   // reschedule only existing tasks
    switchRailTab("updates");
    setSubtasksExpanded(currentSubtasks.length > 0);
    refreshAutoPctUI();

    // Tags
    renderTagPills(String(t.Tags || "").split(";").map((x) => x.trim()).filter(Boolean));
    // Blocked
    renderBlockedPills(String(t.BlockedByTaskIDs || "").split(",").map((x) => x.trim()).filter(Boolean));

    recomputeWsjfDisplay();
  }

  function recomputeWsjfDisplay() {
    const ubv = numericVal("m-ubv");
    const tc  = numericVal("m-tc");
    const rr  = numericVal("m-rr");
    const js  = Math.max(1, numericVal("m-js"));
    const cod = ubv + tc + rr;
    const wsjf = cod / js;
    document.getElementById("m-cod").textContent = cod;
    document.getElementById("m-ubv-readout").textContent = ubv;
    document.getElementById("m-tc-readout").textContent  = tc;
    document.getElementById("m-rr-readout").textContent  = rr;
    document.getElementById("m-js-readout").textContent  = js;
    document.getElementById("m-wsjf-score").textContent  = wsjf.toFixed(1);
  }

  function numericVal(id) {
    const v = Number(document.getElementById(id).value);
    return isNaN(v) ? 0 : v;
  }

  // Live capacity guardrail in the modal: the selected owner's projected load for
  // the selected quarter (including this task) vs their effective capacity.
  function updateCapacityReadout() {
    const el = document.getElementById("m-capacity-readout");
    if (!el) return;
    const owner = document.getElementById("m-owner").value;
    const quarter = document.getElementById("m-quarter").value;
    if (!owner || !quarter) { el.hidden = true; return; }
    const js = Math.max(1, numericVal("m-js"));
    const exceptId = currentEditingTask ? currentEditingTask.TaskID : null;
    const projected = plannedPoints(owner, quarter, exceptId) + js;
    const base = personBaseline(owner);
    const cap = personCapacity(owner, quarter);
    const who = owner.split(" ")[0];
    let cls, text;
    if (base === 0) {
      cls = "cap-warn";
      text = who + " · " + quarter + ": " + projected + " pts planned — set a baseline to see capacity";
    } else if (cap === 0) {
      cls = projected > 0 ? "cap-over" : "cap-ok";
      text = who + " · " + quarter + ": " + projected + " pts planned — no capacity this quarter";
    } else {
      const ratio = projected / cap;
      const pctTxt = Math.round(ratio * 100) + "%";
      cls = ratio > 1 ? "cap-over" : ratio > 0.85 ? "cap-warn" : "cap-ok";
      text = (ratio > 1 ? "⚠ " : "") + who + " " + pctTxt + " · " + quarter + " (" + projected + "/" + cap + " pts)" + (ratio > 1 ? " OVER" : "");
    }
    el.hidden = false;
    el.className = "cap-readout " + cls;
    el.title = text;
    el.textContent = text;
  }

  // Pull earlier / Slip later from inside the modal (acts on the loaded task).
  async function modalReschedule(dir) {
    if (!currentEditingTask || !currentEditingTask.TaskID) return;
    const r = await rescheduleTask(currentEditingTask.TaskID, dir);
    if (!r || r.error) {
      toast(
        r && r.error === "no-quarter" ? "Set a quarter (Q1–Q4) first." :
        r && r.error === "at-start" ? "Already in Q1 — can't pull earlier." :
        r && r.error === "at-end" ? "Already in Q4 — can't slip later." : "Couldn't reschedule.",
        "warn");
      return;
    }
    currentEditingTask.Quarter = r.newQ;
    if (r.startDate !== undefined) currentEditingTask.StartDate = r.startDate;
    if (r.dueDate !== undefined) currentEditingTask.DueDate = r.dueDate;
    if (r.ts) currentEditingTask._loadedLastUpdated = r.ts;  // avoid a false conflict on next Save
    const qs = document.getElementById("m-quarter"); if (qs) qs.value = r.newQ;
    document.getElementById("m-start").value = isoDate(currentEditingTask.StartDate);
    document.getElementById("m-due").value = isoDate(currentEditingTask.DueDate);
    updateCapacityReadout();
    try { renderActivity(await window.WsjfData.readActivityForTask(currentEditingTask.TaskID, 10)); } catch (_) {}
    await reloadTasks();
    if (dir === "earlier") toast("🎉 Accelerated: " + r.oldQ + " → " + r.newQ, "info");
    else toast("↻ Delayed: " + r.oldQ + " → " + r.newQ, "warn");
  }

  // Populate a WSJF dropdown with the Fibonacci scale. Job Size tops out lower
  // than the value fields. Any non-standard saved value is preserved as an option.
  const FIB_VALUE = [1, 2, 3, 5, 8, 13, 20];
  const FIB_SIZE  = [1, 2, 3, 5, 8, 13];
  function fillFib(id, current, fallback, isSize) {
    const el = document.getElementById(id);
    if (!el) return;
    let cur = Number(current);
    if (!cur || isNaN(cur)) cur = fallback;
    const scale = (isSize ? FIB_SIZE : FIB_VALUE).slice();
    if (scale.indexOf(cur) < 0) scale.push(cur);          // keep legacy values selectable
    scale.sort((a, b) => a - b);
    el.innerHTML = scale.map((v) => '<option value="' + v + '">' + v + '</option>').join("");
    el.value = String(cur);
  }

  // Show/refresh the "Set % from subtasks" control. When checked, the % slider is
  // driven by subtask completion (e.g. 4/5 = 80%) and disabled for manual edits.
  function refreshAutoPctUI() {
    const row = document.getElementById("m-autopct-row");
    const cb = document.getElementById("m-autopct");
    const slider = document.getElementById("m-pct");
    if (!row || !cb || !slider) return;
    const has = currentSubtasks.length > 0;
    row.hidden = !has;
    if (has && cb.checked) {
      const done = currentSubtasks.filter((s) => String(s.Done).toLowerCase() === "yes").length;
      const ratio = Math.round((done / currentSubtasks.length) * 100);
      slider.value = ratio;
      slider.disabled = true;
      document.getElementById("m-pct-readout").textContent = ratio;
    } else {
      slider.disabled = false;
    }
  }

  function renderSubtasks() {
    const ul = document.getElementById("m-subtask-list");
    ul.innerHTML = "";
    currentSubtasks
      .sort((a, b) => (Number(a.Order) || 0) - (Number(b.Order) || 0))
      .forEach((s, i) => ul.appendChild(makeSubtaskRow(s, i)));
    document.getElementById("m-subtask-count").textContent =
      currentSubtasks.length ? `(${currentSubtasks.filter((x) => String(x.Done).toLowerCase() === "yes").length}/${currentSubtasks.length})` : "";
    refreshAutoPctUI();

    // Sortable for reorder
    if (typeof Sortable === "undefined") return;
    Sortable.create(ul, {
      handle: ".drag",
      animation: 120,
      onEnd: () => {
        Array.from(ul.querySelectorAll(".subtask-item")).forEach((el, idx) => {
          const sid = el.dataset.sid;
          const item = currentSubtasks.find((x) => String(x.SubtaskID || x._tempId) === String(sid));
          if (item) item.Order = idx + 1;
        });
        // Persist order changes
        currentSubtasks.filter((s) => s.SubtaskID).forEach((s) => {
          window.WsjfData.writeSubtask({ SubtaskID: s.SubtaskID, Order: s.Order })
            .catch((e) => console.warn("subtask reorder failed", e));
        });
      }
    });
  }

  // Compact single-line subtask row. Click the text (or ⤢) to open the subtask
  // modal for full detail + updates; checkbox toggles done; 🗑 deletes.
  function makeSubtaskRow(s, idx) {
    const li = document.createElement("li");
    const sid = s.SubtaskID || (s._tempId = "tmp-" + Math.random().toString(36).slice(2));
    li.dataset.sid = sid;
    const checked = String(s.Done).toLowerCase() === "yes";
    li.className = "subtask-item" + (checked ? " st-done" : "");
    const due = isoDate(s.DueDate);
    const overdue = !checked && due && due < new Date().toISOString().slice(0, 10);
    const dueHtml = due ? '<span class="st-due-chip' + (overdue ? " overdue" : "") + '">📅 ' + escapeHtml(formatDateShort(due)) + '</span>' : "";
    const ownHtml = s.Owner ? '<span class="st-avatar" style="background:' + colorHash(s.Owner) + '" title="' + escapeAttr(s.Owner) + '">' + initialsFromName(s.Owner) + '</span>' : "";
    li.innerHTML =
      '<span class="drag" title="Drag to reorder">⋮⋮</span>' +
      '<input type="checkbox"' + (checked ? " checked" : "") + ' title="Mark done" />' +
      '<span class="subtask-text" title="Open subtask">' + (escapeHtml(s.Text || "") || '<span class="muted">Untitled subtask</span>') + '</span>' +
      '<span class="st-rowmeta">' + dueHtml + ownHtml + '</span>' +
      '<button class="icon-btn st-open" title="Open subtask (detail + updates)">⤢</button>' +
      '<button class="icon-btn icon-trash" title="Delete">🗑</button>';

    const cb = li.querySelector('input[type="checkbox"]');
    const txt = li.querySelector(".subtask-text");
    const trash = li.querySelector(".icon-trash");
    const openBtn = li.querySelector(".st-open");

    cb.addEventListener("click", (e) => e.stopPropagation());
    cb.addEventListener("change", () => {
      s.Done = cb.checked ? "Yes" : "No";
      li.classList.toggle("st-done", cb.checked);
      if (s.SubtaskID) {
        clearTimeout(pendingSubtaskToggleTimers[s.SubtaskID]);
        pendingSubtaskToggleTimers[s.SubtaskID] = setTimeout(() => {
          window.WsjfData.toggleSubtask(s.SubtaskID, cb.checked).catch((e) => toast("Subtask save failed: " + e.message, "error"));
        }, 500);
      }
      document.getElementById("m-subtask-count").textContent =
        `(${currentSubtasks.filter((x) => String(x.Done).toLowerCase() === "yes").length}/${currentSubtasks.length})`;
      refreshAutoPctUI();
    });

    txt.addEventListener("click", () => openSubtaskModal(s));
    openBtn.addEventListener("click", (e) => { e.stopPropagation(); openSubtaskModal(s); });

    trash.addEventListener("click", async (e) => {
      e.stopPropagation();
      if (!(await uiConfirm("Delete this subtask?", { okText: "Delete" }))) return;
      try {
        if (s.SubtaskID) await window.WsjfData.deleteSubtask(s.SubtaskID);
        currentSubtasks = currentSubtasks.filter((x) => x !== s);
        renderSubtasks();
      } catch (e) {
        toast("Delete failed: " + e.message, "error");
      }
    });

    return li;
  }

  // Dedicated subtask modal — full detail + its own chronological Updates feed.
  // Opens on top of the task modal; saves back into currentSubtasks.
  function openSubtaskModal(s) {
    const isNew = !s;
    if (isNew) s = { SubtaskID: null, Text: "", Done: "No", Owner: "", DueDate: "", Order: currentSubtasks.length + 1 };
    const hasId = !!s.SubtaskID;
    const host = document.createElement("div");
    host.className = "modal-overlay st-modal-overlay";
    const ownerList = (State.config.Owners || []).slice();
    if (s.Owner && ownerList.indexOf(s.Owner) < 0) ownerList.push(s.Owner);
    const ownerOpts = '<option value="">— owner —</option>' +
      ownerList.map((o) => '<option value="' + escapeAttr(o) + '"' + (s.Owner === o ? " selected" : "") + '>' + escapeHtml(o) + '</option>').join("");
    const checked = String(s.Done).toLowerCase() === "yes";
    const updatesHtml = hasId
      ? '<div class="st-m-body">' +
          '<h4 class="st-m-h">Updates</h4>' +
          '<div id="st-m-updates" class="updates-list"></div>' +
          '<div class="update-add">' +
            '<textarea id="st-m-update-text" rows="2" placeholder="Add an update…"></textarea>' +
            '<button class="btn btn-primary btn-sm" id="st-m-update-add">Post update</button>' +
          '</div>' +
        '</div>'
      : '<div class="st-m-body"><p class="muted" style="font-size:12px;">Save this subtask first to start its updates feed.</p></div>';
    host.innerHTML =
      '<div class="modal st-modal">' +
        '<header class="modal-header">' +
          '<input id="st-m-text" class="modal-title-input" value="' + escapeAttr(s.Text || "") + '" placeholder="' + (isNew ? "New subtask…" : "Subtask…") + '" />' +
          '<button class="close-btn" id="st-m-close">×</button>' +
        '</header>' +
        '<div class="st-m-meta">' +
          '<button type="button" id="st-m-done" class="st-done-toggle' + (checked ? " on" : "") + '">' + (checked ? "✓ Done" : "Mark done") + '</button>' +
          '<label class="st-m-fl">👤 <select id="st-m-owner">' + ownerOpts + '</select></label>' +
          '<label class="st-m-fl">📅 <input type="date" id="st-m-due" value="' + escapeAttr(isoDate(s.DueDate)) + '" /></label>' +
          '<span class="st-m-doneon" id="st-m-doneon">' + (checked && s.CompletedDate ? "Completed " + escapeHtml(formatDateShort(s.CompletedDate)) : "") + '</span>' +
        '</div>' +
        updatesHtml +
        '<footer class="modal-footer">' +
          '<span style="flex:1"></span>' +
          '<button class="btn btn-secondary" id="st-m-cancel">Cancel</button>' +
          '<button class="btn btn-primary" id="st-m-save">' + (isNew ? "Add subtask" : "Save") + '</button>' +
        '</footer>' +
      '</div>';
    document.body.appendChild(host);
    function close() { host.remove(); }
    host.addEventListener("click", (e) => { if (e.target === host) close(); });
    host.querySelector("#st-m-close").addEventListener("click", close);
    host.querySelector("#st-m-cancel").addEventListener("click", close);

    let doneState = checked;
    const doneBtn = host.querySelector("#st-m-done");
    doneBtn.addEventListener("click", () => {
      doneState = !doneState;
      doneBtn.classList.toggle("on", doneState);
      doneBtn.textContent = doneState ? "✓ Done" : "Mark done";
    });

    // Updates feed (existing subtasks only).
    if (hasId) {
      const renderStUpdates = (list) => {
        const ul = host.querySelector("#st-m-updates");
        ul.innerHTML = "";
        if (!list || !list.length) { ul.innerHTML = '<div class="updates-empty">No updates yet.</div>'; return; }
        list.forEach((u) => {
          const div = document.createElement("div");
          div.className = "update-item";
          div.innerHTML = '<div class="update-text">' + escapeHtml(u.Text || "") + '</div>' +
            '<div class="update-meta">' + escapeHtml(u.AddedBy || "?") + ' · ' +
            escapeHtml(formatDateShort(u.AddedDate)) + ' · ' + relTime(u.AddedDate) + '</div>';
          ul.appendChild(div);
        });
      };
      const loadStUpdates = async () => {
        try { renderStUpdates(await window.WsjfData.readUpdatesForParent("Subtask", s.SubtaskID)); }
        catch (e) { renderStUpdates([]); }
      };
      loadStUpdates();
      const addBtn = host.querySelector("#st-m-update-add");
      addBtn.addEventListener("click", async () => {
        const ta = host.querySelector("#st-m-update-text");
        const txt = ta.value.trim();
        if (!txt) return;
        addBtn.disabled = true;
        try {
          await window.WsjfData.createUpdate({ ParentType: "Subtask", ParentID: s.SubtaskID, Text: txt });
          ta.value = "";
          await loadStUpdates();
        } catch (e) { toast("Update failed: " + e.message, "error"); }
        addBtn.disabled = false;
      });
    }

    host.querySelector("#st-m-save").addEventListener("click", async () => {
      const newText = host.querySelector("#st-m-text").value.trim();
      const newOwner = host.querySelector("#st-m-owner").value;
      const newDue = host.querySelector("#st-m-due").value;
      if (!newText) { toast("Subtask text is required.", "warn"); return; }
      const saveBtn = host.querySelector("#st-m-save"); saveBtn.disabled = true; saveBtn.textContent = "Saving…";
      try {
        if (isNew) {
          const parentId = currentEditingTask && currentEditingTask.TaskID;
          if (parentId) {
            const id = await window.WsjfData.createSubtask({
              ParentTaskID: parentId, Text: newText, Done: doneState ? "Yes" : "No",
              Order: currentSubtasks.length + 1, DueDate: newDue, Owner: newOwner
            });
            currentSubtasks.push({ SubtaskID: id, ParentTaskID: parentId, Text: newText,
              Done: doneState ? "Yes" : "No", Order: currentSubtasks.length + 1, DueDate: newDue, Owner: newOwner,
              CompletedDate: doneState ? new Date().toISOString().slice(0, 10) : "" });
          } else {
            // Unsaved task — keep in memory; flushed when the task is saved.
            currentSubtasks.push({ SubtaskID: null, Text: newText, Done: doneState ? "Yes" : "No",
              Order: currentSubtasks.length + 1, DueDate: newDue, Owner: newOwner });
          }
        } else {
          if (hasId) {
            if (doneState !== checked) await window.WsjfData.toggleSubtask(s.SubtaskID, doneState);
            await window.WsjfData.writeSubtask({ SubtaskID: s.SubtaskID, Text: newText, Owner: newOwner, DueDate: newDue });
          }
          s.Text = newText; s.Owner = newOwner; s.DueDate = newDue; s.Done = doneState ? "Yes" : "No";
          if (doneState && !s.CompletedDate) s.CompletedDate = new Date().toISOString().slice(0, 10);
          if (!doneState) s.CompletedDate = "";
        }
        renderSubtasks();
        setSubtasksExpanded(true);
        if (currentEditingTask && currentEditingTask.TaskID) await syncParentPctFromSubtasks(currentEditingTask.TaskID);
        toast(isNew ? "Subtask added." : "Subtask saved.", "info");
        close();
      } catch (e) {
        saveBtn.disabled = false; saveBtn.textContent = isNew ? "Add subtask" : "Save";
        toast("Save failed: " + e.message, "error");
      }
    });

    setTimeout(() => { const ti = host.querySelector("#st-m-text"); if (ti) ti.focus(); }, 0);
  }

  // Capture the only inline-editable subtask field (the Done checkbox) back into
  // currentSubtasks before a re-render or save. Text/due/owner are edited in the
  // subtask modal now, so they're already on the object — not read from the DOM.
  function syncSubtasksFromDom() {
    const ul = document.getElementById("m-subtask-list");
    if (!ul) return;
    Array.from(ul.querySelectorAll(".subtask-item")).forEach((li) => {
      const sid = li.dataset.sid;
      const cbEl = li.querySelector('input[type="checkbox"]');
      const item = currentSubtasks.find((x) => String(x.SubtaskID || x._tempId) === String(sid));
      if (item && cbEl) item.Done = cbEl.checked ? "Yes" : "No";
    });
  }

  function renderAttachments() {
    const ul = document.getElementById("m-att-list");
    ul.innerHTML = "";
    currentAttachments.forEach((a) => {
      const li = document.createElement("li");
      li.className = "att-item";
      const icon = attIcon(a.Type);
      li.innerHTML = `
        <span class="att-icon">${icon}</span>
        <a href="${escapeAttr(a.Url || "#")}" target="_blank" rel="noopener">${escapeHtml(a.Label || a.Url || "")}</a>
        <button class="icon-btn icon-x" title="Remove">×</button>
      `;
      li.querySelector(".icon-x").addEventListener("click", async () => {
        if (!(await uiConfirm("Remove attachment?", { okText: "Remove" }))) return;
        try {
          if (a.AttachmentID) await window.WsjfData.deleteAttachment(a.AttachmentID);
          currentAttachments = currentAttachments.filter((x) => x !== a);
          renderAttachments();
        } catch (e) {
          toast("Delete failed: " + e.message, "error");
        }
      });
      ul.appendChild(li);
    });
  }

  async function saveNewAttachment() {
    const label = document.getElementById("m-att-label").value.trim();
    const url = document.getElementById("m-att-url").value.trim();
    const type = document.getElementById("m-att-type").value;
    if (!label || !url) { toast("Label and URL required", "warn"); return; }

    if (currentEditingTask.TaskID) {
      try {
        const id = await window.WsjfData.createAttachment({
          ParentTaskID: currentEditingTask.TaskID, Label: label, Url: url, Type: type
        });
        currentAttachments.push({ AttachmentID: id, ParentTaskID: currentEditingTask.TaskID, Label: label, Url: url, Type: type });
      } catch (e) {
        toast("Attachment save failed: " + e.message, "error"); return;
      }
    } else {
      currentAttachments.push({ AttachmentID: null, Label: label, Url: url, Type: type });
    }
    document.getElementById("m-att-form").hidden = true;
    clearAttForm();
    renderAttachments();
  }
  function clearAttForm() {
    document.getElementById("m-att-label").value = "";
    document.getElementById("m-att-url").value = "";
    document.getElementById("m-att-type").value = "web";
  }

  function renderActivity(entries) {
    const ul = document.getElementById("m-activity-list");
    ul.innerHTML = "";
    entries.forEach((e) => {
      const li = document.createElement("li");
      const old = e.OldValue ? `${e.OldValue} → ` : "";
      const newv = e.NewValue || "";
      const field = e.FieldChanged ? ` ${e.FieldChanged}: ` : " ";
      li.innerHTML = `<strong>${escapeHtml(e.ChangedBy || "?")}</strong> ${escapeHtml(e.Action || "")}${field}${escapeHtml(old)}<strong>${escapeHtml(newv)}</strong> · <span class="muted">${relTime(e.Timestamp)}</span>`;
      ul.appendChild(li);
    });
    if (entries.length === 0) {
      ul.innerHTML = `<li class="muted">No activity yet.</li>`;
    }
  }

  // ───── Contributors (people chips + ＋, excludes the Owner) ─────
  var currentContribList = [];
  function currentContributors() {
    const owner = document.getElementById("m-owner").value;
    return currentContribList.filter((p) => p && p !== owner);
  }
  function renderContributors(selected) {
    if (selected) currentContribList = selected.filter(Boolean);
    const wrap = document.getElementById("m-contributors");
    if (!wrap) return;
    const owner = document.getElementById("m-owner").value;
    currentContribList = currentContribList.filter((p) => p && p !== owner);  // owner can't be a contributor
    wrap.innerHTML = "";
    currentContribList.forEach((p) => {
      const chip = document.createElement("span");
      chip.className = "person-chip";
      chip.innerHTML = '<span class="pc-avatar" style="background:' + colorHash(p) + '">' + initialsFromName(p) + '</span>' +
        escapeHtml(p) + ' <span class="pc-x" title="Remove">×</span>';
      chip.querySelector(".pc-x").addEventListener("click", () => {
        currentContribList = currentContribList.filter((x) => x !== p);
        renderContributors(); markDirty();
      });
      wrap.appendChild(chip);
    });
    // ＋ add menu
    const addWrap = document.createElement("span");
    addWrap.className = "pc-add-wrap";
    const addBtn = document.createElement("button");
    addBtn.type = "button"; addBtn.className = "pc-add"; addBtn.textContent = "＋";
    const menu = document.createElement("div");
    menu.className = "pc-menu"; menu.hidden = true;
    const avail = (State.config.Owners || []).filter((o) => o && o !== owner && currentContribList.indexOf(o) < 0);
    if (!avail.length) menu.innerHTML = '<div class="pc-menu-empty">No more people</div>';
    avail.forEach((o) => {
      const item = document.createElement("div");
      item.className = "pc-menu-item"; item.textContent = o;
      item.addEventListener("click", () => {
        currentContribList.push(o); menu.hidden = true; renderContributors(); markDirty();
      });
      menu.appendChild(item);
    });
    addBtn.addEventListener("click", (e) => { e.stopPropagation(); menu.hidden = !menu.hidden; });
    document.addEventListener("mousedown", (e) => { if (!addWrap.contains(e.target)) menu.hidden = true; });
    addWrap.appendChild(addBtn); addWrap.appendChild(menu);
    wrap.appendChild(addWrap);
  }

  // ───── Goals (multi, like workstreams) ─────
  function currentTaskGoals() {
    return Array.from(document.querySelectorAll("#m-goal input:checked")).map((c) => c.value);
  }
  function renderTaskGoals(selected) {
    const wrap = document.getElementById("m-goal");
    if (!wrap) return;
    const sel = new Set((selected || []).map(String).filter(Boolean));
    const goals = (State.goals || []).slice();
    // Keep already-assigned goals not in the Goals list (renamed/deleted goal).
    sel.forEach((id) => { if (!goals.some((g) => String(g.GoalID) === id)) goals.push({ GoalID: id, ShortName: id + " (unknown)" }); });
    wrap.innerHTML = "";
    if (!goals.length) { wrap.innerHTML = '<span class="muted contrib-empty">No goals defined.</span>'; return; }
    goals.forEach((g) => {
      const l = document.createElement("label");
      l.className = "contrib-check";
      const cb = document.createElement("input");
      cb.type = "checkbox"; cb.value = g.GoalID; cb.checked = sel.has(String(g.GoalID));
      cb.addEventListener("change", markDirty);
      l.appendChild(cb);
      l.appendChild(document.createTextNode(" " + (g.ShortName || g.GoalName || g.GoalID)));
      wrap.appendChild(l);
    });
  }

  // ───── Tags ─────
  function collectAllTags() {
    const set = new Set();
    State.tasks.forEach((t) => {
      String(t.Tags || "").split(";").map((s) => s.trim()).filter(Boolean).forEach((tg) => set.add(tg));
    });
    return Array.from(set).sort();
  }

  function renderTagPills(tags) {
    const wrap = document.getElementById("m-tag-pills");
    wrap.innerHTML = "";
    tags.forEach((tg) => {
      const pill = document.createElement("span");
      pill.className = "tag-pill removable";
      pill.innerHTML = `${escapeHtml(tg)} <span class="x">×</span>`;
      pill.querySelector(".x").addEventListener("click", () => {
        const list = currentTags();
        renderTagPills(list.filter((t) => t !== tg));
        markDirty();
      });
      wrap.appendChild(pill);
    });
  }
  function currentTags() {
    return Array.from(document.querySelectorAll("#m-tag-pills .tag-pill"))
      .map((el) => el.firstChild.textContent.trim());
  }
  function handleTagKey(e) {
    const input = e.target;
    if (e.key === "," || e.key === "Enter") {
      e.preventDefault();
      const v = input.value.trim().replace(/,$/, "");
      if (v) {
        const list = currentTags();
        if (!list.includes(v)) {
          list.push(v);
          renderTagPills(list);
          markDirty();
        }
      }
      input.value = "";
      document.getElementById("m-tag-suggest").hidden = true;
    } else if (e.key === "Backspace" && !input.value) {
      const list = currentTags();
      if (list.length) {
        list.pop();
        renderTagPills(list);
        markDirty();
      }
    }
  }
  function showTagSuggestions(e) {
    const v = e.target.value.trim().toLowerCase();
    const box = document.getElementById("m-tag-suggest");
    if (!v) { box.hidden = true; return; }
    const taken = new Set(currentTags());
    const matches = currentTagList.filter((t) => t.toLowerCase().includes(v) && !taken.has(t)).slice(0, 8);
    if (!matches.length) { box.hidden = true; return; }
    box.innerHTML = matches.map((m) => `<div class="suggest-item">${escapeHtml(m)}</div>`).join("");
    box.hidden = false;
    Array.from(box.querySelectorAll(".suggest-item")).forEach((el) => {
      el.addEventListener("mousedown", () => {
        const list = currentTags();
        list.push(el.textContent);
        renderTagPills(list);
        document.getElementById("m-tag-input").value = "";
        box.hidden = true;
        markDirty();
      });
    });
  }

  // ───── Blocked-by (continued) ─────
  function renderBlockedPills(taskIdList) {
    const wrap = document.getElementById("m-blocked-pills");
    wrap.innerHTML = "";
    taskIdList.forEach((tid) => {
      const t = State.tasks.find((x) => Number(x.TaskID) === Number(tid));
      const label = t ? `#${tid} ${t.Title}` : `#${tid}`;
      const pill = document.createElement("span");
      pill.className = "blocked-pill";
      pill.innerHTML = `🚧 ${escapeHtml(label)} <span class="x">×</span>`;
      pill.querySelector(".x").addEventListener("click", () => {
        const list = currentBlockedIds().filter((x) => String(x) !== String(tid));
        renderBlockedPills(list);
        markDirty();
      });
      wrap.appendChild(pill);
    });
  }

  function currentBlockedIds() {
    return Array.from(document.querySelectorAll("#m-blocked-pills .blocked-pill"))
      .map((el) => {
        const m = el.textContent.match(/#(\d+)/);
        return m ? Number(m[1]) : null;
      })
      .filter(Boolean);
  }

  function openBlockedPicker() {
    const picker = document.getElementById("m-blocked-picker");
    const already = new Set(currentBlockedIds().map(Number));
    const candidates = State.tasks
      .filter((t) => Number(t.TaskID) !== Number(currentEditingTask.TaskID || -1))
      .filter((t) => !already.has(Number(t.TaskID)));

    picker.innerHTML = `
      <input type="text" class="blocked-search" placeholder="Search tasks…" />
      <div class="blocked-results"></div>
    `;
    picker.hidden = false;

    const search = picker.querySelector(".blocked-search");
    const results = picker.querySelector(".blocked-results");

    function refresh() {
      const q = search.value.toLowerCase();
      const matches = candidates
        .filter((t) => !q || String(t.Title || "").toLowerCase().includes(q) || String(t.TaskID).includes(q))
        .slice(0, 25);
      results.innerHTML = matches.map((t) =>
        `<div class="blocked-result" data-tid="${t.TaskID}">#${t.TaskID} ${escapeHtml(t.Title || "")}</div>`
      ).join("") || `<div class="muted" style="padding:8px;">No matches</div>`;

      Array.from(results.querySelectorAll(".blocked-result")).forEach((el) => {
        el.addEventListener("click", () => {
          const list = currentBlockedIds().concat([Number(el.dataset.tid)]);
          renderBlockedPills(list);
          picker.hidden = true;
          markDirty();
        });
      });
    }
    search.addEventListener("input", refresh);
    refresh();
    setTimeout(() => search.focus(), 30);

    // close on outside click
    setTimeout(() => {
      const onDoc = (e) => {
        if (!picker.contains(e.target) && e.target.id !== "m-blocked-add") {
          picker.hidden = true;
          document.removeEventListener("click", onDoc);
        }
      };
      document.addEventListener("click", onDoc);
    }, 50);
  }

  // ───── Save / Archive ─────
  async function saveModal() {
    const t = currentEditingTask;

    // Capture any subtask text the user typed but didn't blur out of.
    syncSubtasksFromDom();

    // Gather form values
    t.Title              = document.getElementById("m-title").value.trim();
    t.Description        = document.getElementById("m-description").value;
    t.UserBusinessValue  = numericVal("m-ubv");
    t.TimeCriticality    = numericVal("m-tc");
    t.RiskReduction      = numericVal("m-rr");
    t.JobSize            = Math.max(1, numericVal("m-js"));
    // % comes from subtasks when "Set % from subtasks" is on; otherwise the slider.
    const autoCb = document.getElementById("m-autopct");
    if (currentSubtasks.length && autoCb && autoCb.checked) {
      const sd = currentSubtasks.filter((s) => String(s.Done).toLowerCase() === "yes").length;
      t.PercentComplete = Math.round((sd / currentSubtasks.length) * 100);
    } else {
      t.PercentComplete = numericVal("m-pct");
    }
    t.Owner              = document.getElementById("m-owner").value;
    t.Contributors       = currentContributors().join("; ");
    t.WorkstreamID       = document.getElementById("m-workstream").value;
    t.GoalID             = currentTaskGoals().join("; ");
    t.Quarter            = document.getElementById("m-quarter").value;
    const newStatus      = document.getElementById("m-status").value;
    t.Health             = document.getElementById("m-health").value;
    t.StartDate          = document.getElementById("m-start").value || "";
    t.DueDate            = document.getElementById("m-due").value || "";
    t.Tags               = currentTags().join("; ");
    t.BlockedByTaskIDs   = currentBlockedIds().join(",");

    if (!t.Title) { toast("Title is required.", "warn"); return; }

    // Status changed? maintain ColumnOrder accordingly
    if (t.Status !== newStatus) {
      t.ColumnOrder = maxColumnOrderForStatus(newStatus) + 1;
      t.Status = newStatus;
    }
    // Done ⇒ 100%. Reopening (Done → other) drops to the subtask ratio, or 90% if
    // none — mirroring the board-drag rule — but only if % is still at the Done
    // value (don't override a % the user just set manually).
    if (newStatus === "Done") {
      t.PercentComplete = 100;
    } else if (originalTaskSnapshot && originalTaskSnapshot.Status === "Done" && Number(t.PercentComplete) === 100) {
      const sd = currentSubtasks.filter((s) => String(s.Done).toLowerCase() === "yes").length;
      t.PercentComplete = currentSubtasks.length ? Math.round((sd / currentSubtasks.length) * 100) : 90;
    }

    const saveBtn = document.getElementById("m-save");
    saveBtn.disabled = true; saveBtn.textContent = "Saving…";

    try {
      if (t.TaskID) {
        // Update path
        try {
          await window.WsjfData.writeTask(t);
        } catch (err) {
          if (err && err.code === "CONFLICT") {
            const who = err.serverRow && err.serverRow.UpdatedBy || "someone";
            const when = relTime(err.serverRow && err.serverRow.LastUpdated);
            const choice = await uiConfirm(`Card was updated by ${who} ${when}.\nOverwrite their changes, or keep theirs?`, { okText: "Overwrite", cancelText: "Keep theirs" });
            if (!choice) {
              toast("Save cancelled — refresh to see latest.", "info");
              saveBtn.disabled = false; saveBtn.textContent = "Save";
              return;
            }
            await window.WsjfData.writeTask(t, { force: true });
          } else {
            throw err;
          }
        }
        // Persist any subtask text/done edits that weren't flushed on blur.
        // If the task itself is now Done, every subtask closes with it (stay in
        // sync with the server-side cascade rather than re-opening them).
        const taskDone = t.Status === "Done";
        const doneToday = new Date().toISOString().slice(0, 10);
        for (const s of currentSubtasks) {
          const wasDone = String(s.Done).toLowerCase() === "yes";
          const done = taskDone ? "Yes" : (s.Done || "No");
          if (s.SubtaskID) {
            const payload = { SubtaskID: s.SubtaskID, Text: s.Text || "", Done: done, Order: s.Order || 1, DueDate: s.DueDate || "", Owner: s.Owner || "" };
            // only stamp a completion date on the transition — don't clobber an existing one
            if (taskDone && !wasDone) payload.CompletedDate = doneToday;
            await window.WsjfData.writeSubtask(payload);
          } else if (s.Text) {
            await window.WsjfData.createSubtask({ ParentTaskID: t.TaskID, Text: s.Text, Done: done, Order: s.Order || 1, DueDate: s.DueDate || "", Owner: s.Owner || "" });
          }
        }
        toast("Saved.", "info");
      } else {
        // Create path
        const newId = await window.WsjfData.createTask(t);
        t.TaskID = newId;
        // Flush any in-memory subtasks/attachments collected before the task existed
        const createdDone = t.Status === "Done";
        for (const s of currentSubtasks) {
          if (!s.SubtaskID && s.Text) {
            await window.WsjfData.createSubtask({
              ParentTaskID: newId, Text: s.Text, Done: createdDone ? "Yes" : (s.Done || "No"), Order: s.Order || 1,
              DueDate: s.DueDate || "", Owner: s.Owner || ""
            });
          }
        }
        for (const a of currentAttachments) {
          if (!a.AttachmentID && a.Url) {
            await window.WsjfData.createAttachment({
              ParentTaskID: newId, Label: a.Label, Url: a.Url, Type: a.Type
            });
          }
        }
        toast("Task created.", "info");
      }

      modalDirty = false;
      await reloadTasks();
      render();
      hideModal();
    } catch (err) {
      console.error(err);
      toast("Save failed: " + err.message, "error");
    } finally {
      saveBtn.disabled = false; saveBtn.textContent = "Save";
    }
  }

  async function archiveFromModal() {
    if (!currentEditingTask || !currentEditingTask.TaskID) {
      hideModal();
      return;
    }
    if (!(await uiConfirm("Archive this task? It will be hidden from the board.", { okText: "Archive" }))) return;
    try {
      await window.WsjfData.archiveTask(currentEditingTask.TaskID);
      toast("Archived.", "info");
      modalDirty = false;
      await reloadTasks();
      render();
      hideModal();
    } catch (err) {
      toast("Archive failed: " + err.message, "error");
    }
  }

  function maxColumnOrderForStatus(status) {
    let max = 0;
    State.tasks.forEach((t) => {
      if (t.Status === status) {
        const c = Number(t.ColumnOrder) || 0;
        if (c > max) max = c;
      }
    });
    return max;
  }

  // ───── Utilities ─────
  function fillSelect(id, options, placeholder, selectedValue) {
    const el = document.getElementById(id);
    if (!el) return;
    el.innerHTML = "";
    options.forEach((opt) => {
      const o = document.createElement("option");
      if (typeof opt === "object" && opt !== null) {
        o.value = opt.value;
        o.textContent = opt.label;
      } else {
        o.value = opt;
        o.textContent = opt === "" && placeholder ? placeholder : opt;
      }
      el.appendChild(o);
    });
    if (selectedValue !== undefined && selectedValue !== null) el.value = selectedValue;
  }

  function statusSlug(s) {
    return String(s).toLowerCase().replace(/[^a-z0-9]+/g, "-");
  }

  // Tiny localStorage helpers (remember accordion open/closed state, etc.).
  function lsGet(k, d) { try { const v = localStorage.getItem(k); return v == null ? d : v; } catch (e) { return d; } }
  function lsSet(k, v) { try { localStorage.setItem(k, v); } catch (e) {} }

  // Schedule chip (PI planning): green » when net accelerated, red « when net
  // delayed (rolled over). Driven by the optional signed Slips counter.
  function scheduleChip(t) {
    const s = Number(t && t.Slips) || 0;
    if (s > 0) return '<span class="sched-chip sched-late" title="Delayed ' + s + ' quarter' + (s === 1 ? "" : "s") + ' (net)">« ' + s + '</span> ';
    if (s < 0) return '<span class="sched-chip sched-early" title="Accelerated ' + (-s) + ' quarter' + (-s === 1 ? "" : "s") + ' (net)">» ' + (-s) + '</span> ';
    return "";
  }

  function wsjfPillClass(v) {
    if (v >= 15) return "wsjf-red";
    if (v >= 10) return "wsjf-orange";
    if (v >= 5)  return "wsjf-yellow";
    return "wsjf-gray";
  }

  function colorHash(str) {
    let h = 0;
    const s = String(str || "");
    for (let i = 0; i < s.length; i++) h = (h << 5) - h + s.charCodeAt(i);
    const hue = Math.abs(h) % 360;
    return `hsl(${hue}, 55%, 70%)`;
  }

  function initialsFromName(name) {
    if (!name) return "?";
    const parts = String(name).replace(/[,\[\]]/g, " ").trim().split(/\s+/);
    if (parts.length === 1) return parts[0].substring(0, 2).toUpperCase();
    return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
  }

  function workstreamName(id) {
    const w = State.workstreams.find((x) => x.WorkstreamID === id);
    return w ? (w.Name || w.WorkstreamID) : (id || "—");
  }

  // ───── Goal helpers (distinct colour per goal, used on cards + roadmap) ─────
  const GOAL_PALETTE = ["#6366f1", "#0ea5e9", "#10b981", "#f59e0b", "#ef4444",
                        "#ec4899", "#8b5cf6", "#14b8a6", "#f97316", "#64748b"];
  function goalColor(goalId) {
    if (!goalId) return "#cbd5e1";
    const idx = State.goals.findIndex((g) => String(g.GoalID) === String(goalId));
    return idx >= 0 ? GOAL_PALETTE[idx % GOAL_PALETTE.length] : colorHash(goalId);
  }
  function goalShortName(goalId) {
    const g = State.goals.find((x) => String(x.GoalID) === String(goalId));
    return g ? (g.ShortName || g.GoalName || g.GoalID) : goalId;
  }
  function taskGoalIds(t) {
    return String(t.GoalID || "").split(/[;,]/).map((s) => s.trim()).filter(Boolean);
  }

  // Is a person committed beyond their effective capacity in a quarter?
  function isOverAllocated(person, quarter) {
    if (!person || !quarter) return false;
    const cap = personCapacity(person, quarter);
    if (!cap) return false;                       // no baseline → can't judge
    return plannedPoints(person, quarter) > cap;
  }

  // Owner + contributor avatars (owner first, ringed). Used on board + list.
  function peopleAvatars(task, max) {
    max = max || 4;
    const owner = String(task.Owner || "").trim();
    const contribs = String(task.Contributors || "").split(/[;,]/)
      .map((s) => s.trim()).filter(Boolean).filter((c) => c !== owner);
    const people = [];
    if (owner) people.push({ name: owner, owner: true });
    contribs.forEach((c) => people.push({ name: c, owner: false }));
    const shown = people.slice(0, max);
    const extra = people.length - shown.length;
    let h = shown.map((p) => {
      const over = isOverAllocated(p.name, task.Quarter);
      return '<span class="avatar' + (p.owner ? " avatar-owner" : "") + (over ? " avatar-overalloc" : "") + '" style="background:' +
        colorHash(p.name) + '" title="' + escapeAttr(p.name + (p.owner ? " · owner" : " · contributor") + (over ? " · ⚠ over-allocated in " + task.Quarter : "")) +
        '">' + initialsFromName(p.name) + '</span>';
    }).join("");
    if (extra > 0) {
      h += '<span class="avatar avatar-more" title="' +
        escapeAttr(people.slice(max).map((p) => p.name).join(", ")) + '">+' + extra + '</span>';
    }
    return h;
  }

  // Goal colour dots (one per goal the task serves).
  function goalDotsHtml(task) {
    const ids = taskGoalIds(task);
    if (!ids.length) return "";
    return '<span class="goal-dots">' + ids.map((gid) =>
      '<span class="goal-dot" style="background:' + goalColor(gid) +
      '" title="' + escapeAttr(goalShortName(gid)) + '"></span>').join("") + '</span>';
  }

  function attIcon(type) {
    switch (String(type).toLowerCase()) {
      case "confluence": return "📄";
      case "sharepoint": return "📁";
      case "file":       return "📎";
      default:           return "🌐";
    }
  }

  function isoDate(d) {
    if (!d) return "";
    if (typeof d === "string") {
      // Already iso-like? trim any time portion.
      const m = d.match(/^(\d{4}-\d{2}-\d{2})/);
      if (m) return m[1];
      const dt = new Date(d);
      if (!isNaN(dt.getTime())) return dt.toISOString().substring(0, 10);
      return "";
    }
    if (typeof d === "number") {
      // Excel serial date — convert (Excel epoch is 1899-12-30)
      const ms = (d - 25569) * 86400 * 1000;
      const dt = new Date(ms);
      if (!isNaN(dt.getTime())) return dt.toISOString().substring(0, 10);
    }
    return "";
  }

  function formatDateShort(d) {
    const iso = isoDate(d);
    if (!iso) return "";
    const dt = new Date(iso + "T00:00:00");
    if (isNaN(dt.getTime())) return "";
    return dt.toLocaleDateString(undefined, { month: "short", day: "numeric" });
  }

  function currentQuarter() {
    const m = new Date().getMonth() + 1;
    if (m <= 3) return "Q1";
    if (m <= 6) return "Q2";
    if (m <= 9) return "Q3";
    return "Q4";
  }

  // Quarter-end date — prefers the QuartersTable EndDate, falls back to calendar.
  function quarterEndDate(q) {
    const d = State.quarterDates && State.quarterDates[q];
    if (d && d.end) { const iso = isoDate(d.end); if (iso) return iso; }
    const year = new Date().getFullYear();
    const ends = { Q1: "-03-31", Q2: "-06-30", Q3: "-09-30", Q4: "-12-31" };
    const suffix = ends[q];
    return suffix ? (year + suffix) : "";
  }

  // ───── Capacity helpers (planning + at-assignment) ─────
  function ownerMatchesPerson(t, person) {
    return String(t.Owner || "").split(/[;,]/).map((s) => s.trim()).includes(person);
  }
  // Demonstrated baseline = points completed in Q1 (fixed all year, per plan).
  function personVelocity(person) {
    let pts = 0;
    State.tasks.forEach((t) => {
      if (t.Quarter === "Q1" && t.Status === "Done" && ownerMatchesPerson(t, person)) pts += Number(t.JobSize) || 0;
    });
    return pts;
  }
  // Baseline = override, else demonstrated Q1 velocity.
  function personBaseline(person) {
    const c = State.capacity[person] || {};
    return (c.baseline != null && !isNaN(c.baseline)) ? c.baseline : personVelocity(person);
  }
  // Availability for a quarter: that quarter's AvailQ* if set (0 is valid), else
  // the single Availability column, else 1.
  function personAvailability(person, quarter) {
    const c = State.capacity[person] || {};
    let a = (quarter && c.byQuarter && c.byQuarter[quarter] != null && !isNaN(c.byQuarter[quarter]))
      ? c.byQuarter[quarter] : c.availability;
    return (a == null || isNaN(a)) ? 1 : a;
  }
  // Effective capacity = baseline × availability(quarter).
  function personCapacity(person, quarter) {
    return Math.round(personBaseline(person) * personAvailability(person, quarter));
  }
  // Committed points for a person in a quarter (optionally excluding one task).
  function plannedPoints(person, quarter, exceptTaskId) {
    let pts = 0;
    State.tasks.forEach((t) => {
      if (quarter && t.Quarter !== quarter) return;
      if (exceptTaskId != null && Number(t.TaskID) === Number(exceptTaskId)) return;
      if (!ownerMatchesPerson(t, person)) return;
      pts += Number(t.JobSize) || 0;
    });
    return pts;
  }

  // Days remaining until the end of a quarter (negative if past). null if unknown.
  function daysLeftInQuarter(q) {
    const end = quarterEndDate(q);
    if (!end) return null;
    const today = new Date(new Date().toISOString().slice(0, 10) + "T00:00:00");
    const endDt = new Date(end + "T00:00:00");
    return Math.round((endDt - today) / 86400000);
  }

  function shiftIso(iso, deltaMs) {
    const base = isoDate(iso);
    if (!base) return iso;
    const t = new Date(base + "T00:00:00").getTime();
    if (isNaN(t)) return iso;
    return new Date(t + deltaMs).toISOString().slice(0, 10);
  }

  // Move a task one quarter earlier ("Pulled earlier" / accelerated) or later
  // ("Slipped later" / delayed), shifting its dates, bumping the roll-over count,
  // and logging the event distinctly. dir = "earlier" | "later".
  async function rescheduleTask(taskId, dir) {
    const t = State.tasks.find((x) => Number(x.TaskID) === Number(taskId));
    if (!t) return { error: "not-found" };
    const order = ["Q1", "Q2", "Q3", "Q4"];
    const idx = order.indexOf(t.Quarter);
    if (idx < 0) return { error: "no-quarter" };
    const ni = dir === "earlier" ? idx - 1 : idx + 1;
    if (ni < 0) return { error: "at-start" };
    if (ni >= order.length) return { error: "at-end" };
    const oldQ = order[idx], newQ = order[ni];

    const updates = { TaskID: taskId, Quarter: newQ };
    // Shift dates by the gap between the two quarters' starts (keeps duration).
    let delta = null;
    const oqd = State.quarterDates[oldQ], nqd = State.quarterDates[newQ];
    if (oqd && nqd && oqd.start && nqd.start) {
      delta = new Date(isoDate(nqd.start) + "T00:00:00").getTime() -
        new Date(isoDate(oqd.start) + "T00:00:00").getTime();
      if (t.StartDate) updates.StartDate = shiftIso(t.StartDate, delta);
      if (t.DueDate) updates.DueDate = shiftIso(t.DueDate, delta);
    }
    // Net schedule counter (optional Slips column): +1 delayed (later), −1 accelerated
    // (earlier). Signed: >0 net delayed, <0 net accelerated.
    updates.Slips = (Number(t.Slips) || 0) + (dir === "later" ? 1 : -1);

    const ts = await window.WsjfData.writeTask(updates, { force: true, silent: true });
    // Cascade the same shift to subtask due dates so they move with the task.
    if (delta) {
      const subs = State.subtasksByParent[taskId] || [];
      for (const s of subs) {
        if (s.SubtaskID && s.DueDate) {
          const nd = shiftIso(s.DueDate, delta);
          try { await window.WsjfData.writeSubtask({ SubtaskID: s.SubtaskID, DueDate: nd }); s.DueDate = nd; }
          catch (_) {}
        }
      }
    }
    const action = dir === "earlier" ? "Accelerated" : "Delayed";
    await window.WsjfData.logActivity("Task", taskId, action, "Quarter", oldQ, newQ, "");

    // Local mirror so the board/list reflects it without a full reload.
    t.Quarter = newQ; t.Slips = updates.Slips;
    if (updates.StartDate !== undefined) t.StartDate = updates.StartDate;
    if (updates.DueDate !== undefined) t.DueDate = updates.DueDate;
    return { oldQ, newQ, action, ts, startDate: updates.StartDate, dueDate: updates.DueDate };
  }

  function relTime(iso) {
    if (!iso) return "";
    const t = new Date(iso).getTime();
    if (isNaN(t)) return String(iso);
    const s = Math.floor((Date.now() - t) / 1000);
    if (s < 0)     return "just now";   // guard against clock skew / future stamps
    if (s < 60)    return s + "s ago";
    if (s < 3600)  return Math.floor(s / 60) + "m ago";
    if (s < 86400) return Math.floor(s / 3600) + "h ago";
    return Math.floor(s / 86400) + "d ago";
  }

  function escapeHtml(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
  }
  function escapeAttr(s) { return escapeHtml(s); }

  function debounce(fn, ms) {
    let to;
    return function () {
      clearTimeout(to);
      const args = arguments, ctx = this;
      to = setTimeout(() => fn.apply(ctx, args), ms);
    };
  }

  // ───── Identity ("who am I") ─────
  function updateIdentityChip() {
    const chip = document.getElementById("identity-chip");
    if (!chip) return;
    const known = State.me && State.me.name && State.me.name !== "Unknown User";
    chip.textContent = known ? State.me.initials : "👤";
    chip.title = known ? ("You: " + State.me.name + " — click to change") : "Set your name";
    chip.classList.toggle("unset", !known);
  }

  // On first run (no saved identity) we gate the app until the user identifies
  // themselves — their name stamps edits and powers "My Tasks".
  async function ensureIdentity() {
    const known = State.me && State.me.name && State.me.name !== "Unknown User";
    if (known) return;
    // Loop until they choose (mustChoose hides Cancel / outside-close).
    let me = null;
    while (!me) { me = await openIdentityPicker({ mustChoose: true }); }
  }

  function openIdentityPicker(opts) {
    opts = opts || {};
    return new Promise((resolve) => {
      const host = document.createElement("div");
      host.className = "confirm-overlay";
      const box = document.createElement("div");
      box.className = "confirm-box identity-box";
      const title = document.createElement("h3");
      title.className = "identity-title";
      title.textContent = opts.mustChoose ? "👋 Welcome — who are you?" : "Who are you?";
      const msg = document.createElement("p");
      msg.className = "confirm-msg";
      msg.textContent = 'This stamps your edits and powers your "My Tasks" dashboard. Saved on this device only. Don\'t see your name? Pick "Add someone new" and type it.';
      const sel = document.createElement("select");
      sel.className = "identity-select";
      const owners = (State.config && State.config.Owners) ? State.config.Owners : [];
      sel.innerHTML = '<option value="">— Select your name —</option>' +
        owners.map((o) => '<option>' + escapeHtml(o) + '</option>').join("") +
        '<option value="__other__">➕ Add someone new…</option>';
      const txt = document.createElement("input");
      txt.type = "text"; txt.className = "identity-text"; txt.placeholder = "Type the new name"; txt.style.display = "none";
      const cur = (State.me && State.me.name && State.me.name !== "Unknown User") ? State.me.name : "";
      if (cur && owners.indexOf(cur) >= 0) sel.value = cur;
      sel.addEventListener("change", () => {
        txt.style.display = sel.value === "__other__" ? "block" : "none";
        if (sel.value === "__other__") txt.focus();
      });
      const actions = document.createElement("div");
      actions.className = "confirm-actions";
      const cancel = document.createElement("button"); cancel.className = "btn btn-secondary"; cancel.textContent = "Cancel";
      const save = document.createElement("button"); save.className = "btn btn-primary"; save.textContent = "Save";
      if (!opts.mustChoose) actions.appendChild(cancel);
      actions.appendChild(save);
      box.appendChild(title); box.appendChild(msg); box.appendChild(sel); box.appendChild(txt); box.appendChild(actions);
      host.appendChild(box);
      function done(v) { host.remove(); resolve(v); }
      cancel.addEventListener("click", () => done(null));
      if (!opts.mustChoose) host.addEventListener("click", (e) => { if (e.target === host) done(null); });
      save.addEventListener("click", async () => {
        const isNew = sel.value === "__other__";
        const name = isNew ? txt.value.trim() : sel.value;
        if (!name) { toast("Pick or type your name.", "warn"); return; }
        save.disabled = true;
        try {
          // A typed-in new person is added to the Owners list so they "exist"
          // everywhere (filters, assignee dropdowns) going forward.
          if (isNew && owners.map((o) => o.toLowerCase()).indexOf(name.toLowerCase()) < 0) {
            try { await window.WsjfData.addConfigValue("OwnersTable", name); await loadConfigAndDimensions(); }
            catch (e) { console.warn("addConfigValue(owner) failed (non-fatal):", e); }
          }
          const me = await window.WsjfData.setCurrentUser(name);
          State.me = me;
          updateIdentityChip();
          toast("Hi, " + me.name + "!", "info");
          done(me);
        } catch (e) {
          save.disabled = false;
          toast("Could not save name: " + e.message, "error");
        }
      });
      document.body.appendChild(host);
    });
  }

  // ───── Confirm dialog ─────
  // Native confirm()/alert() are blocked inside Office dialogs (sandboxed iframe),
  // so use a custom DOM-based confirm that works everywhere. Returns a Promise<boolean>.
  // Multi-choice dialog. choices = [{ value, label, cls }]. Resolves chosen value
  // (or null on cancel). Used by the roadmap cross-quarter drop prompt.
  function uiChoose(title, message, choices) {
    return new Promise((resolve) => {
      const host = document.createElement("div");
      host.className = "confirm-overlay";
      const box = document.createElement("div");
      box.className = "confirm-box";
      box.innerHTML = '<h3 class="confirm-title">' + escapeHtml(title) + '</h3>' +
        (message ? '<p class="confirm-msg">' + escapeHtml(message) + '</p>' : "");
      const actions = document.createElement("div");
      actions.className = "confirm-actions confirm-actions-col";
      function done(val) { host.remove(); resolve(val); }
      choices.forEach((c) => {
        const b = document.createElement("button");
        b.className = "btn " + (c.cls || "btn-secondary");
        b.textContent = c.label;
        b.addEventListener("click", () => done(c.value));
        actions.appendChild(b);
      });
      const cancel = document.createElement("button");
      cancel.className = "btn btn-link";
      cancel.textContent = "Cancel";
      cancel.addEventListener("click", () => done(null));
      actions.appendChild(cancel);
      box.appendChild(actions);
      host.appendChild(box);
      host.addEventListener("click", (e) => { if (e.target === host) done(null); });
      document.body.appendChild(host);
    });
  }

  // Which quarter does an ISO date fall in? Prefers QuartersTable bounds, else month.
  function quarterOf(iso) {
    const d = isoDate(iso);
    if (!d) return "";
    const qd = State.quarterDates || {};
    for (const q of Object.keys(qd)) {
      const s = isoDate(qd[q].start), e = isoDate(qd[q].end);
      if (s && e && d >= s && d <= e) return q;
    }
    const m = Number(d.slice(5, 7));
    return m <= 3 ? "Q1" : m <= 6 ? "Q2" : m <= 9 ? "Q3" : "Q4";
  }

  function uiConfirm(message, opts) {
    opts = opts || {};
    return new Promise((resolve) => {
      const host = document.createElement("div");
      host.className = "confirm-overlay";
      const box = document.createElement("div");
      box.className = "confirm-box";
      const msg = document.createElement("p");
      msg.className = "confirm-msg";
      msg.textContent = message;
      const actions = document.createElement("div");
      actions.className = "confirm-actions";
      const cancelBtn = document.createElement("button");
      cancelBtn.className = "btn btn-secondary";
      cancelBtn.textContent = opts.cancelText || "Cancel";
      const okBtn = document.createElement("button");
      okBtn.className = "btn btn-primary";
      okBtn.textContent = opts.okText || "OK";
      actions.appendChild(cancelBtn);
      actions.appendChild(okBtn);
      box.appendChild(msg);
      box.appendChild(actions);
      host.appendChild(box);
      function done(val) { host.remove(); resolve(val); }
      cancelBtn.addEventListener("click", () => done(false));
      okBtn.addEventListener("click", () => done(true));
      host.addEventListener("click", (e) => { if (e.target === host) done(false); });
      document.body.appendChild(host);
      okBtn.focus();
    });
  }

  // ───── Toasts ─────
  function toast(msg, kind) {
    const host = document.getElementById("toast-host");
    const el = document.createElement("div");
    el.className = "toast toast-" + (kind || "info");
    el.textContent = msg;
    host.appendChild(el);
    requestAnimationFrame(() => el.classList.add("in"));
    setTimeout(() => {
      el.classList.remove("in");
      el.classList.add("out");
      setTimeout(() => el.remove(), 400);
    }, kind === "error" ? 6000 : 3500);
  }

  // ───── Fatal error display ─────
  function showFatalError(err) {
    const board = document.getElementById("kanban");
    const msg = (err && (err.message || String(err))) || "Unknown error";
    const code = (err && err.code) ? err.code : "";
    let detail = "";
    try { if (err && err.debugInfo) detail = JSON.stringify(err.debugInfo, null, 2); } catch (_) {}
    const stack = (err && err.stack) ? err.stack : "";
    if (!board) { alert((code ? "[" + code + "] " : "") + msg); return; }
    board.innerHTML =
      '<div style="max-width:720px;margin:40px auto;padding:24px;background:#fff;border-radius:8px;box-shadow:0 1px 3px rgba(0,0,0,.1);">' +
        '<h2 style="margin-top:0;color:#CA001B;">⚠️ Could not load the board</h2>' +
        '<p class="fe-msg" style="font-weight:600;font-size:15px;"></p>' +
        '<pre class="fe-detail" style="white-space:pre-wrap;background:#f5f5f5;padding:12px;border-radius:6px;font-size:12px;overflow:auto;max-height:320px;"></pre>' +
        '<button id="fe-back" style="margin-top:12px;padding:8px 14px;border:none;border-radius:6px;background:#f0f0f0;cursor:pointer;font-weight:600;">‹ Back to stats</button>' +
      '</div>';
    board.querySelector(".fe-msg").textContent = (code ? "[" + code + "] " : "") + msg;
    board.querySelector(".fe-detail").textContent = (detail ? detail + "\n\n" : "") + stack;
    const b = document.getElementById("fe-back");
    if (b) b.addEventListener("click", goBackToStats);
  }

  function goBackToStats() {
    window.location.href = new URL("taskpane.html", window.location.href).href;
  }

  // expose for debugging
  window.__WSJF = { State, render, reload: reloadTasks };
})();