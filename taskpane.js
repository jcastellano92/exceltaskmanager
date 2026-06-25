/* taskpane.js — side panel launcher + quick stats */
(function () {
  let dialog = null;
  let lastSyncIso = null;

  Office.onReady((info) => {
    if (info.host !== Office.HostType.Excel) {
      console.warn("Product Management Tool only supports Excel.");
    }
    document.getElementById("open-board-btn").addEventListener("click", openBoardDialog);
    document.getElementById("sync-now-btn").addEventListener("click", refreshStats);
    refreshStats();
    setInterval(updateRelativeSync, 5000);
  });

  function openBoardDialog() {
    // CONFIRMED: Office dialogs cannot run Excel.run in this tenant — ctx.sync
    // throws GeneralException even for a single sequential read. So render the
    // board in THIS task pane, where the Excel session is valid.
    window.location.href = new URL("board.html", window.location.href).href;
  }

  function onDialogMessage(arg) {
    // Dialog can post messages like "refresh-stats" to ask side panel to update.
    try {
      const msg = JSON.parse(arg.message);
      if (msg.type === "refresh-stats") refreshStats();
    } catch (e) {
      console.warn("Bad dialog message", arg.message);
    }
  }

  function onDialogEvent(arg) {
    if (arg.error === 12006) {
      // user closed dialog
      dialog = null;
      refreshStats();
    }
  }

  async function refreshStats() {
    showStatus("Syncing…", "amber");
    try {
      const tasks = await window.WsjfData.readAllTasks();
      const me = await window.WsjfData.getCurrentUser();
      const q = currentQuarter();
      document.getElementById("stat-total").textContent = tasks.length;
      document.getElementById("stat-mine").textContent = tasks.filter(
        (t) => (t.Owner || "").toLowerCase().includes(me.name.toLowerCase()) ||
               (t.Owner || "").toLowerCase().includes((me.email || "").toLowerCase())
      ).length;
      document.getElementById("stat-blocked").textContent = tasks.filter((t) => t.Status === "Blocked").length;
      document.getElementById("stat-done").textContent = tasks.filter((t) => t.Status === "Done" && t.Quarter === q).length;
      lastSyncIso = new Date().toISOString();
      updateRelativeSync();
      showStatus("Ready", "green");
    } catch (err) {
      console.error(err);
      showStatus("Sync failed", "red");
    }
  }

  function updateRelativeSync() {
    if (!lastSyncIso) {
      document.getElementById("last-sync-time").textContent = "—";
      return;
    }
    document.getElementById("last-sync-time").textContent = relTime(lastSyncIso);
  }

  function relTime(iso) {
    const s = Math.floor((Date.now() - new Date(iso).getTime()) / 1000);
    if (s < 60) return s + "s ago";
    if (s < 3600) return Math.floor(s / 60) + "m ago";
    return Math.floor(s / 3600) + "h ago";
  }

  function currentQuarter() {
    const m = new Date().getMonth() + 1;
    if (m <= 3) return "Q1";
    if (m <= 6) return "Q2";
    if (m <= 9) return "Q3";
    return "Q4";
  }

  function showStatus(text, color) {
    document.getElementById("sync-text").textContent = text;
    const dot = document.getElementById("sync-dot");
    dot.className = "dot dot-" + color;
  }
})();