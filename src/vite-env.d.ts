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
  };
}
