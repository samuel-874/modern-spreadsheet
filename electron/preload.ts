import { ipcRenderer, contextBridge } from "electron";

contextBridge.exposeInMainWorld("electronAPI", {
  openFile: () => ipcRenderer.invoke("file:open"),
  getSaveFolder: () => ipcRenderer.invoke("file:get-save-folder"),
  chooseSaveFolder: () => ipcRenderer.invoke("file:choose-save-folder"),
  saveFile: (payload: { filePath?: string; name: string; content: string }) =>
    ipcRenderer.invoke("file:save", payload),

  // Custom title bar and window controls
  minimize: () => ipcRenderer.invoke("window:minimize"),
  maximize: () => ipcRenderer.invoke("window:maximize"),
  close: () => ipcRenderer.invoke("window:close"),
  isMaximized: () => ipcRenderer.invoke("window:is-maximized"),
  getPlatform: () => ipcRenderer.invoke("window:get-platform"),
  onMaximizedChange: (callback: (isMaximized: boolean) => void) => {
    const handler = (_event: any, isMax: boolean) => callback(isMax);
    ipcRenderer.on("window:maximized-change", handler);
    return () => {
      ipcRenderer.off("window:maximized-change", handler);
    };
  },

  // Auto-updater and version info
  getVersion: () => ipcRenderer.invoke("app:get-version"),
  checkForUpdates: () => ipcRenderer.invoke("update:check"),
  restartAndInstall: () => ipcRenderer.invoke("update:restart-and-install"),
  onUpdateStatus: (
    callback: (data: { status: string; version?: string; error?: string }) => void,
  ) => {
    const handler = (
      _event: unknown,
      data: { status: string; version?: string; error?: string },
    ) => callback(data);
    ipcRenderer.on("update:status", handler);
    return () => {
      ipcRenderer.off("update:status", handler);
    };
  },
  onUpdateProgress: (
    callback: (progress: {
      percent: number;
      bytesPerSecond?: number;
      transferred?: number;
      total?: number;
    }) => void,
  ) => {
    const handler = (
      _event: unknown,
      progress: {
        percent: number;
        bytesPerSecond?: number;
        transferred?: number;
        total?: number;
      },
    ) => callback(progress);
    ipcRenderer.on("update:progress", handler);
    return () => {
      ipcRenderer.off("update:progress", handler);
    };
  },
});

// --------- Expose some API to the Renderer process ---------
contextBridge.exposeInMainWorld("ipcRenderer", {
  on(...args: Parameters<typeof ipcRenderer.on>) {
    const [channel, listener] = args;
    return ipcRenderer.on(channel, (event, ...args) =>
      listener(event, ...args),
    );
  },
  off(...args: Parameters<typeof ipcRenderer.off>) {
    const [channel, ...omit] = args;
    return ipcRenderer.off(channel, ...omit);
  },
  send(...args: Parameters<typeof ipcRenderer.send>) {
    const [channel, ...omit] = args;
    return ipcRenderer.send(channel, ...omit);
  },
  invoke(...args: Parameters<typeof ipcRenderer.invoke>) {
    const [channel, ...omit] = args;
    return ipcRenderer.invoke(channel, ...omit);
  },

  // You can expose other APTs you need here.
  // ...
});
