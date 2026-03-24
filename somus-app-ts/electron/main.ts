import { app, BrowserWindow, ipcMain, dialog, shell } from 'electron';
import { autoUpdater } from 'electron-updater';
import * as path from 'path';
import * as fs from 'fs';

let mainWindow: BrowserWindow | null = null;

const isDev = !app.isPackaged;

function createWindow(): void {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    minWidth: 1200,
    minHeight: 800,
    frame: false,
    titleBarStyle: 'hidden',
    titleBarOverlay: false,
    show: false,
    backgroundColor: '#0a0a0a',
    icon: path.join(__dirname, '..', 'assets', 'icon_somus.ico'),
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  });

  mainWindow.maximize();
  mainWindow.show();

  if (isDev) {
    const devServerUrl = process.env.VITE_DEV_SERVER_URL || 'http://localhost:5173';
    mainWindow.loadURL(devServerUrl);
    mainWindow.webContents.openDevTools({ mode: 'detach' });
  } else {
    mainWindow.loadFile(path.join(__dirname, '..', 'dist', 'index.html'));
  }

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

// ── App lifecycle ──────────────────────────────────────────────────────────────

app.whenReady().then(() => {
  createWindow();
  setupIpcHandlers();

  if (!isDev) {
    autoUpdater.checkForUpdatesAndNotify();
  }
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});

// ── IPC Handlers ───────────────────────────────────────────────────────────────

function setupIpcHandlers(): void {
  // ── Window controls ────────────────────────────────────────────────────────
  ipcMain.handle('app:minimize', () => {
    mainWindow?.minimize();
  });

  ipcMain.handle('app:maximize', () => {
    if (mainWindow?.isMaximized()) {
      mainWindow.unmaximize();
    } else {
      mainWindow?.maximize();
    }
  });

  ipcMain.handle('app:close', () => {
    mainWindow?.close();
  });

  ipcMain.handle('app:quit', () => {
    app.quit();
  });

  ipcMain.handle('app:getVersion', () => {
    return app.getVersion();
  });

  ipcMain.handle('app:getPath', (_event, name: string) => {
    return app.getPath(name as any);
  });

  // ── File system ────────────────────────────────────────────────────────────
  ipcMain.handle('fs:readFile', async (_event, filePath: string, encoding?: BufferEncoding) => {
    try {
      const content = await fs.promises.readFile(filePath, { encoding: encoding ?? 'utf-8' });
      return { success: true, data: content };
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle('fs:writeFile', async (_event, filePath: string, data: string, encoding?: BufferEncoding) => {
    try {
      const dir = path.dirname(filePath);
      await fs.promises.mkdir(dir, { recursive: true });
      await fs.promises.writeFile(filePath, data, { encoding: encoding ?? 'utf-8' });
      return { success: true };
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle('fs:readDir', async (_event, dirPath: string) => {
    try {
      const entries = await fs.promises.readdir(dirPath, { withFileTypes: true });
      const result = entries.map((e) => ({
        name: e.name,
        isDirectory: e.isDirectory(),
        isFile: e.isFile(),
      }));
      return { success: true, data: result };
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle('fs:exists', async (_event, filePath: string) => {
    try {
      await fs.promises.access(filePath);
      return true;
    } catch {
      return false;
    }
  });

  ipcMain.handle('fs:mkdir', async (_event, dirPath: string) => {
    try {
      await fs.promises.mkdir(dirPath, { recursive: true });
      return { success: true };
    } catch (err: any) {
      return { success: false, error: err.message };
    }
  });

  // ── Dialogs ────────────────────────────────────────────────────────────────
  ipcMain.handle('dialog:openFile', async (_event, options?: Electron.OpenDialogOptions) => {
    if (!mainWindow) return { success: false, error: 'No window' };
    const result = await dialog.showOpenDialog(mainWindow, {
      properties: ['openFile'],
      ...options,
    });
    if (result.canceled) return { success: false, canceled: true };
    return { success: true, filePaths: result.filePaths };
  });

  ipcMain.handle('dialog:saveFile', async (_event, options?: Electron.SaveDialogOptions) => {
    if (!mainWindow) return { success: false, error: 'No window' };
    const result = await dialog.showSaveDialog(mainWindow, {
      ...options,
    });
    if (result.canceled) return { success: false, canceled: true };
    return { success: true, filePath: result.filePath };
  });

  ipcMain.handle('dialog:openDirectory', async (_event, options?: Electron.OpenDialogOptions) => {
    if (!mainWindow) return { success: false, error: 'No window' };
    const result = await dialog.showOpenDialog(mainWindow, {
      properties: ['openDirectory'],
      ...options,
    });
    if (result.canceled) return { success: false, canceled: true };
    return { success: true, filePaths: result.filePaths };
  });

  // ── Shell ──────────────────────────────────────────────────────────────────
  ipcMain.handle('shell:openExternal', async (_event, url: string) => {
    await shell.openExternal(url);
  });

  ipcMain.handle('shell:openPath', async (_event, filePath: string) => {
    await shell.openPath(filePath);
  });

  // ── Auto-updater ───────────────────────────────────────────────────────────
  ipcMain.handle('updater:checkForUpdates', () => {
    autoUpdater.checkForUpdates();
  });

  ipcMain.handle('updater:installUpdate', () => {
    autoUpdater.quitAndInstall(false, true);
  });

  autoUpdater.on('update-available', (info) => {
    mainWindow?.webContents.send('updater:update-available', info);
  });

  autoUpdater.on('download-progress', (progress) => {
    mainWindow?.webContents.send('updater:download-progress', progress);
  });

  autoUpdater.on('update-downloaded', (info) => {
    mainWindow?.webContents.send('updater:update-downloaded', info);
  });

  autoUpdater.on('error', (err) => {
    mainWindow?.webContents.send('updater:error', err.message);
  });
}
