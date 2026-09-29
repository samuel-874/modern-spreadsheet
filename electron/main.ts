import { app, BrowserWindow, dialog, ipcMain } from "electron";
import { autoUpdater } from "electron-updater";
import { fileURLToPath } from "node:url";
import path from "node:path";
import { promises as fs } from "node:fs";
import fsSync from "node:fs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// The built directory structure
//
// ├─┬─┬ dist
// │ │ └── index.html
// │ │
// │ ├─┬ dist-electron
// │ │ ├── main.js
// │ │ └── preload.mjs
// │
process.env.APP_ROOT = path.join(__dirname, "..");

// 🚧 Use ['ENV_NAME'] avoid vite:define plugin - Vite@2.x
export const VITE_DEV_SERVER_URL = process.env["VITE_DEV_SERVER_URL"];
export const MAIN_DIST = path.join(process.env.APP_ROOT, "dist-electron");
export const RENDERER_DIST = path.join(process.env.APP_ROOT, "dist");

process.env.VITE_PUBLIC = VITE_DEV_SERVER_URL
  ? path.join(process.env.APP_ROOT, "public")
  : RENDERER_DIST;

let win: BrowserWindow | null;
let defaultSaveFolder = "";
const settingsPath = () => path.join(app.getPath("userData"), "settings.json");

function createWindow() {
  const isMac = process.platform === "darwin";
  const iconPath = process.platform === "win32"
    ? path.join(process.env.VITE_PUBLIC, "icon.ico")
    : path.join(process.env.VITE_PUBLIC, "icon.png");

  const resolvedIcon = fsSync.existsSync(iconPath)
    ? iconPath
    : path.join(process.env.VITE_PUBLIC, "electron-vite.svg");

  win = new BrowserWindow({
    title: "Paperclip",
    icon: resolvedIcon,
    width: 1240,
    height: 780,
    minWidth: 800,
    minHeight: 520,
    // Custom Title Bar configurations:
    titleBarStyle: isMac ? "hiddenInset" : "hidden",
    trafficLightPosition: isMac ? { x: 14, y: 13 } : undefined,
    webPreferences: {
      preload: path.join(__dirname, "preload.mjs"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
    },
  });

  // Hide default menu on Windows/Linux for modern custom title bar look
  if (!isMac) {
    win.setMenuBarVisibility(false);
  }

  // Set dock icon on macOS during development
  if (isMac && app.dock) {
    const dockIcon = path.join(process.env.VITE_PUBLIC, "icon.png");
    if (fsSync.existsSync(dockIcon)) {
      app.dock.setIcon(dockIcon);
    }
  }

  // Track maximize state and broadcast to renderer for titlebar icon toggle
  win.on("maximize", () => {
    win?.webContents.send("window:maximized-change", true);
  });
  win.on("unmaximize", () => {
    win?.webContents.send("window:maximized-change", false);
  });

  // Test active push message to Renderer-process.
  win.webContents.on("did-finish-load", () => {
    win?.webContents.send("main-process-message", new Date().toLocaleString());
  });

  if (VITE_DEV_SERVER_URL) {
    win.loadURL(VITE_DEV_SERVER_URL);
  } else {
    // win.loadFile('dist/index.html')
    win.loadFile(path.join(RENDERER_DIST, "index.html"));
  }
}

// Window control IPC handlers for Custom Title Bar
ipcMain.handle("window:minimize", () => {
  win?.minimize();
});

ipcMain.handle("window:maximize", () => {
  if (!win) return;
  if (win.isMaximized()) {
    win.unmaximize();
  } else {
    win.maximize();
  }
});

ipcMain.handle("window:close", () => {
  win?.close();
});

ipcMain.handle("window:is-maximized", () => {
  return win ? win.isMaximized() : false;
});

ipcMain.handle("window:get-platform", () => {
  return process.platform;
});

ipcMain.handle("file:open", async () => {
  const result = await dialog.showOpenDialog({
    properties: ["openFile"],
    filters: [
      { name: "Spreadsheets", extensions: ["csv", "xlsx", "xls"] },
      { name: "CSV Files", extensions: ["csv"] },
      { name: "Excel Workbooks", extensions: ["xlsx", "xls"] },
    ],
  });
  if (result.canceled || !result.filePaths[0]) return null;
  const filePath = result.filePaths[0];
  const ext = path.extname(filePath).toLowerCase();
  if (ext === ".xlsx" || ext === ".xls") {
    const buffer = await fs.readFile(filePath);
    return {
      filePath,
      name: path.basename(filePath),
      base64: buffer.toString("base64"),
      isExcel: true,
    };
  }
  return {
    filePath,
    name: path.basename(filePath),
    content: await fs.readFile(filePath, "utf8"),
    isExcel: false,
  };
});

ipcMain.handle("file:get-save-folder", () => {
  if (!defaultSaveFolder) defaultSaveFolder = app.getPath("desktop");
  return defaultSaveFolder;
});

ipcMain.handle("file:choose-save-folder", async () => {
  if (!defaultSaveFolder) defaultSaveFolder = app.getPath("desktop");
  const result = await dialog.showOpenDialog({
    defaultPath: defaultSaveFolder,
    properties: ["openDirectory", "createDirectory"],
  });
  if (result.canceled || !result.filePaths[0]) return defaultSaveFolder;
  defaultSaveFolder = result.filePaths[0];
  await fs.writeFile(
    settingsPath(),
    JSON.stringify({ defaultSaveFolder }),
    "utf8",
  );
  return defaultSaveFolder;
});

ipcMain.handle(
  "file:save",
  async (
    _event,
    payload: { filePath?: string; name: string; content: string },
  ) => {
    if (
      !payload ||
      typeof payload.name !== "string" ||
      typeof payload.content !== "string"
    )
      throw new Error("Invalid save request");
    const result = payload.filePath
      ? { canceled: false, filePath: payload.filePath }
      : await dialog.showSaveDialog({
          defaultPath: path.join(
            defaultSaveFolder || app.getPath("desktop"),
            payload.name,
          ),
          filters: [{ name: "CSV", extensions: ["csv"] }],
        });
    if (result.canceled || !result.filePath) return null;
    await fs.writeFile(result.filePath, payload.content, "utf8");
    return { filePath: result.filePath, name: path.basename(result.filePath) };
  },
);

// Auto-updater handlers and logic
ipcMain.handle("app:get-version", () => {
  return app.getVersion();
});

ipcMain.handle("update:check", async () => {
  if (!app.isPackaged) {
    return { status: "dev", message: "Updates are disabled in development mode." };
  }
  try {
    const result = await autoUpdater.checkForUpdates();
    return { status: "ok", updateInfo: result?.updateInfo };
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    return { status: "error", error: message };
  }
});

ipcMain.handle("update:restart-and-install", () => {
  autoUpdater.quitAndInstall();
});

function initAutoUpdater() {
  autoUpdater.autoDownload = true;
  autoUpdater.autoInstallOnAppQuit = true;

  autoUpdater.logger = {
    info: (...args: unknown[]) => console.log("[AutoUpdater]", ...args),
    warn: (...args: unknown[]) => console.warn("[AutoUpdater]", ...args),
    error: (...args: unknown[]) => console.error("[AutoUpdater]", ...args),
  };

  autoUpdater.on("checking-for-update", () => {
    console.log("[AutoUpdater] Checking for updates...");
    win?.webContents.send("update:status", { status: "checking" });
  });

  autoUpdater.on("update-available", (info) => {
    console.log("[AutoUpdater] Update available:", info.version);
    win?.webContents.send("update:status", { status: "available", version: info.version });
  });

  autoUpdater.on("update-not-available", (info) => {
    console.log("[AutoUpdater] App is up to date:", info.version);
    win?.webContents.send("update:status", { status: "not-available", version: info.version });
  });

  autoUpdater.on("error", (err) => {
    console.error("[AutoUpdater] Update error:", err);
    win?.webContents.send("update:status", { status: "error", error: err.message });
  });

  autoUpdater.on("download-progress", (progress) => {
    win?.webContents.send("update:progress", {
      percent: Math.round(progress.percent),
      bytesPerSecond: progress.bytesPerSecond,
      transferred: progress.transferred,
      total: progress.total,
    });
  });

  autoUpdater.on("update-downloaded", (info) => {
    console.log("[AutoUpdater] Update downloaded:", info.version);
    win?.webContents.send("update:status", { status: "downloaded", version: info.version });

    if (win && !win.isDestroyed()) {
      dialog
        .showMessageBox(win, {
          type: "info",
          title: "Update Ready",
          message: `A new version of Paperclip (${info.version}) has been downloaded.`,
          detail: "Restart now to apply the update, or it will be installed automatically the next time you quit.",
          buttons: ["Restart Now", "Later"],
          defaultId: 0,
          cancelId: 1,
        })
        .then((result) => {
          if (result.response === 0) {
            autoUpdater.quitAndInstall();
          }
        })
        .catch((dialogErr) => {
          console.error("[AutoUpdater] Dialog error:", dialogErr);
        });
    }
  });

  if (app.isPackaged) {
    // Initial check 3 seconds after launch to ensure smooth startup
    setTimeout(() => {
      autoUpdater.checkForUpdates().catch((err) => {
        console.error("[AutoUpdater] Initial check error:", err);
      });
    }, 3000);

    // Periodic check every 4 hours while app is running
    setInterval(() => {
      autoUpdater.checkForUpdates().catch((err) => {
        console.error("[AutoUpdater] Periodic check error:", err);
      });
    }, 4 * 60 * 60 * 1000);
  } else {
    console.log("[AutoUpdater] Skipping automatic update checks in development mode.");
  }
}

// Quit when all windows are closed, except on macOS. There, it's common
// for applications and their menu bar to stay active until the user quits
// explicitly with Cmd + Q.
app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
    win = null;
  }
});

app.on("activate", () => {
  // On OS X it's common to re-create a window in the app when the
  // dock icon is clicked and there are no other windows open.
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});

app.whenReady().then(async () => {
  try {
    const settings = JSON.parse(await fs.readFile(settingsPath(), "utf8"));
    if (typeof settings.defaultSaveFolder === "string") {
      defaultSaveFolder = settings.defaultSaveFolder;
    }
  } catch {
    // The settings file is optional on first launch.
  }
  createWindow();
  initAutoUpdater();
});
