import { app, BrowserWindow, dialog, ipcMain } from "electron";
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
});
