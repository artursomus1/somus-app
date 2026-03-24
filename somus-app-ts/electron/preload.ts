import { contextBridge, ipcRenderer } from 'electron';

// ── Type definitions for the exposed API ─────────────────────────────────────

export interface FsResult<T = string> {
  success: boolean;
  data?: T;
  error?: string;
}

export interface DirEntry {
  name: string;
  isDirectory: boolean;
  isFile: boolean;
}

export interface DialogResult {
  success: boolean;
  canceled?: boolean;
  filePaths?: string[];
  filePath?: string;
  error?: string;
}

export interface ElectronAPI {
  fs: {
    readFile: (filePath: string, encoding?: string) => Promise<FsResult<string>>;
    writeFile: (filePath: string, data: string, encoding?: string) => Promise<FsResult>;
    readDir: (dirPath: string) => Promise<FsResult<DirEntry[]>>;
    exists: (filePath: string) => Promise<boolean>;
    mkdir: (dirPath: string) => Promise<FsResult>;
  };
  dialog: {
    openFile: (options?: Record<string, unknown>) => Promise<DialogResult>;
    saveFile: (options?: Record<string, unknown>) => Promise<DialogResult>;
    openDirectory: (options?: Record<string, unknown>) => Promise<DialogResult>;
  };
  shell: {
    openExternal: (url: string) => Promise<void>;
    openPath: (filePath: string) => Promise<void>;
  };
  app: {
    getVersion: () => Promise<string>;
    getPath: (name: string) => Promise<string>;
    quit: () => Promise<void>;
    minimize: () => Promise<void>;
    maximize: () => Promise<void>;
    close: () => Promise<void>;
  };
  updater: {
    checkForUpdates: () => Promise<void>;
    onUpdateAvailable: (callback: (info: unknown) => void) => () => void;
    onDownloadProgress: (callback: (progress: unknown) => void) => () => void;
    onUpdateDownloaded: (callback: (info: unknown) => void) => () => void;
    onError: (callback: (message: string) => void) => () => void;
    installUpdate: () => Promise<void>;
  };
}

// ── Helper to create removable IPC listeners ─────────────────────────────────

function createListener(channel: string, callback: (...args: unknown[]) => void): () => void {
  const handler = (_event: Electron.IpcRendererEvent, ...args: unknown[]) => callback(...args);
  ipcRenderer.on(channel, handler);
  return () => {
    ipcRenderer.removeListener(channel, handler);
  };
}

// ── Expose API ───────────────────────────────────────────────────────────────

const electronAPI: ElectronAPI = {
  fs: {
    readFile: (filePath, encoding) =>
      ipcRenderer.invoke('fs:readFile', filePath, encoding),
    writeFile: (filePath, data, encoding) =>
      ipcRenderer.invoke('fs:writeFile', filePath, data, encoding),
    readDir: (dirPath) =>
      ipcRenderer.invoke('fs:readDir', dirPath),
    exists: (filePath) =>
      ipcRenderer.invoke('fs:exists', filePath),
    mkdir: (dirPath) =>
      ipcRenderer.invoke('fs:mkdir', dirPath),
  },

  dialog: {
    openFile: (options) =>
      ipcRenderer.invoke('dialog:openFile', options),
    saveFile: (options) =>
      ipcRenderer.invoke('dialog:saveFile', options),
    openDirectory: (options) =>
      ipcRenderer.invoke('dialog:openDirectory', options),
  },

  shell: {
    openExternal: (url) =>
      ipcRenderer.invoke('shell:openExternal', url),
    openPath: (filePath) =>
      ipcRenderer.invoke('shell:openPath', filePath),
  },

  app: {
    getVersion: () =>
      ipcRenderer.invoke('app:getVersion'),
    getPath: (name) =>
      ipcRenderer.invoke('app:getPath', name),
    quit: () =>
      ipcRenderer.invoke('app:quit'),
    minimize: () =>
      ipcRenderer.invoke('app:minimize'),
    maximize: () =>
      ipcRenderer.invoke('app:maximize'),
    close: () =>
      ipcRenderer.invoke('app:close'),
  },

  updater: {
    checkForUpdates: () =>
      ipcRenderer.invoke('updater:checkForUpdates'),
    onUpdateAvailable: (callback) =>
      createListener('updater:update-available', callback),
    onDownloadProgress: (callback) =>
      createListener('updater:download-progress', callback),
    onUpdateDownloaded: (callback) =>
      createListener('updater:update-downloaded', callback),
    onError: (callback) =>
      createListener('updater:error', callback as (...args: unknown[]) => void),
    installUpdate: () =>
      ipcRenderer.invoke('updater:installUpdate'),
  },
};

contextBridge.exposeInMainWorld('electron', electronAPI);
