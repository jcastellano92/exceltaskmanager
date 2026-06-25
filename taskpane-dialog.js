/* taskpane-dialog.js — sidebar launcher + data broker.
 *
 * Opens board.html as a large in-app dialog (displayInIframe:true). Because an
 * Office dialog can't run Excel.run itself, the board (running with ?host=dialog)
 * proxies every WsjfData call to here via messageParent; we execute it against
 * the real data layer (Excel.run works in this task pane) and reply via
 * messageChild. See makeRemoteWsjf() in board.js for the dialog side.
 */
(function () {
  let dialog = null;

  Office.onReady((info) => {
    if (info.host !== Office.HostType.Excel) {
      console.warn("Product Management Tool only supports Excel.");
    }
    document.getElementById("open-fs-btn").addEventListener("click", openFullScreen);
  });

  function openFullScreen() {
    const url = new URL("board.html?host=dialog", window.location.href).href;
    showStatus("Opening…", "amber");
    Office.context.ui.displayDialogAsync(
      url,
      { height: 97, width: 99, displayInIframe: true, promptBeforeOpen: false },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          showStatus("Could not open: " + result.error.message, "red");
          return;
        }
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, onDialogRpc);
        dialog.addEventHandler(Office.EventType.DialogEventReceived, onDialogEvent);
        showStatus("Board open", "green");
      }
    );
  }

  // The board (dialog) sends { id, method, args }. Run it against the real data
  // layer here and reply with { id, result } or { id, error }.
  async function onDialogRpc(arg) {
    let req;
    try { req = JSON.parse(arg.message); } catch (_) { return; }
    if (!req || !req.method) return;
    try {
      let result;
      if (req.method === "_readTable") {
        result = await window.WsjfData._internal._readTable((req.args || [])[0]);
      } else if (typeof window.WsjfData[req.method] === "function") {
        result = await window.WsjfData[req.method].apply(null, req.args || []);
      } else {
        throw new Error("Unknown method: " + req.method);
      }
      reply({ id: req.id, result: result === undefined ? null : result });
    } catch (err) {
      reply({
        id: req.id,
        error: {
          message: String((err && err.message) || err),
          code: err && err.code,
          serverRow: err && err.serverRow
        }
      });
    }
  }

  function reply(obj) {
    if (!dialog) return;
    try { dialog.messageChild(JSON.stringify(obj)); }
    catch (e) { console.error("messageChild failed", e); }
  }

  function onDialogEvent(arg) {
    // 12006 = dialog closed by user
    if (arg.error === 12006) {
      dialog = null;
      showStatus("Ready", "green");
    }
  }

  function showStatus(text, color) {
    const t = document.getElementById("sync-text");
    if (t) t.textContent = text;
    const d = document.getElementById("sync-dot");
    if (d) d.className = "dot dot-" + color;
  }
})();
