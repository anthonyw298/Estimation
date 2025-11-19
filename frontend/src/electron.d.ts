export interface ElectronAPI {
  send: (channel: string, data: unknown) => void
  receive: (channel: string, func: (...args: unknown[]) => void) => void
}

export interface Versions {
  node: () => string
  chrome: () => string
  electron: () => string
}

declare global {
  interface Window {
    electron: ElectronAPI
    versions: Versions
  }
}
