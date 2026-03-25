/**
 * Test setup file for Vitest
 * Somus Capital - Mesa de Produtos
 *
 * Mocks window.electron and other browser globals for Node environment.
 */

// Mock window.electron for tests running in Node
if (typeof globalThis.window === 'undefined') {
  (globalThis as any).window = {};
}

if (typeof (globalThis as any).window.electron === 'undefined') {
  (globalThis as any).window.electron = {
    ipcRenderer: {
      send: () => {},
      on: () => {},
      invoke: async () => null,
    },
    store: {
      get: async () => null,
      set: async () => {},
    },
  };
}
