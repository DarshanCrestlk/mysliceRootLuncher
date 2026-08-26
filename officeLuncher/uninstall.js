const { app, dialog } = require("electron");
const { removeMySlice } = require("./cleanup");

function isSilent() {
  return process.argv.some((arg) => arg === "--silent" || arg === "/S");
}

app.whenReady().then(() => {
  if (!isSilent()) {
    const choice = dialog.showMessageBoxSync(null, {
      type: "warning",
      title: "MySlice Uninstall",
      message: "Remove MySlice from this PC?",
      detail:
        "This removes the launcher, mysliceLTS:// protocol, network share, Office catalog, and ProgramData files. Close Word and Excel first.",
      buttons: ["Uninstall", "Cancel"],
      defaultId: 1,
      cancelId: 1,
      noLink: true,
    });

    if (choice !== 0) {
      app.quit();
      return;
    }
  }

  const { ok, failedSteps } = removeMySlice();

  if (ok) {
    dialog.showMessageBoxSync(null, {
      type: "info",
      title: "MySlice Uninstall – Complete",
      message: "MySlice was removed from this PC.",
      detail: "Restart Word/Excel if they were open.",
    });
  } else {
    dialog.showErrorBox(
      "MySlice Uninstall – Issue",
      `Failed steps:\n• ${failedSteps.join(
        "\n• "
      )}\n\nRun MySlice Uninstall as Administrator, and close Word/Excel.`
    );
  }

  app.quit();
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
