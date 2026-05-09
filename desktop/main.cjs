const { app, BrowserWindow, dialog, Menu } = require("electron");
const path = require("node:path");

function createWindow() {
  const win = new BrowserWindow({
    width: 1100,
    height: 760,
    backgroundColor: "#0f0322",
    webPreferences: {
      preload: path.join(__dirname, "preload.cjs"),
      sandbox: true,
      contextIsolation: true,
      nodeIntegration: false,
    },
  });
  win.loadFile(path.join(__dirname, "renderer", "index.html"));

  const template = [
    {
      label: "File",
      submenu: [
        {
          label: "Open…",
          accelerator: "CmdOrCtrl+O",
          click: async () => {
            const res = await dialog.showOpenDialog(win, {
              properties: ["openFile"],
              filters: [{ name: "Office files", extensions: ["docx", "xlsx", "pptx", "zip"] }],
            });
            if (!res.canceled && res.filePaths[0]) {
              win.webContents.send("open-file", res.filePaths[0]);
            }
          },
        },
        { role: "quit" },
      ],
    },
    { role: "editMenu" },
    { role: "viewMenu" },
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}

app.whenReady().then(() => {
  createWindow();
  app.on("activate", () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
