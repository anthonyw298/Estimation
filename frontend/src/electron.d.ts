export interface ElectronAPI {
  send: (channel: string, data: any) => void
  receive: (channel: string, func: (...args: any[]) => void) => void
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
