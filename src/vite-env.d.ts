/// <reference types="vite/client" />

type OpenFileResult = {
  filePath: string;
  name: string;
  content?: string;
  base64?: string;
  isExcel?: boolean;
};

type SaveFileResult = {
  filePath: string;
  name: string;
};

interface Window {
  electronAPI?: {
    openFile: () => Promise<OpenFileResult | null>;
    getSaveFolder: () => Promise<string>;
    chooseSaveFolder: () => Promise<string | null>;
    saveFile: (payload: {
      filePath?: string;
      name: string;
      content: string;
    }) => Promise<SaveFileResult | null>;
    minimize: () => Promise<void>;
    maximize: () => Promise<void>;
    close: () => Promise<void>;
    isMaximized: () => Promise<boolean>;
    getPlatform: () => Promise<string>;
    onMaximizedChange: (callback: (isMaximized: boolean) => void) => () => void;
    getVersion: () => Promise<string>;
    checkForUpdates: () => Promise<{ status: string; updateInfo?: unknown; error?: string; message?: string }>;
    restartAndInstall: () => Promise<void>;
    onUpdateStatus: (
      callback: (data: { status: string; version?: string; error?: string }) => void,
    ) => () => void;
    onUpdateProgress: (
      callback: (progress: {
        percent: number;
        bytesPerSecond?: number;
        transferred?: number;
        total?: number;
      }) => void,
    ) => () => void;
  };
}
