/** Legacy account IPC compatibility. DVH Office never reads, writes or revokes Genspark credentials. */
import { homedir } from 'node:os'
import { join } from 'node:path'

export interface GskLoginProgress {
  phase: 'url' | 'success' | 'error'
  url?: string
  expiresInSec?: number
  /** 'network' | 'expired' | raw error text */
  error?: string
}

export function genofficeAuthPath(): string {
  return join(homedir(), '.dvh-office', 'disabled-auth.json')
}

export interface GenofficeAuth {
  apiKey: string
  keyId?: string
  accessToken?: string
}

export function loadGenofficeAuth(): GenofficeAuth | null {
  return null
}

export function reloadGenofficeAuth(): void {}

export function genofficeApiKey(): string {
  return ''
}

export function startGenofficeLogin(_onEvent?: (progress: GskLoginProgress) => void): boolean {
  return false
}

export function genofficeLoginInFlight(): boolean {
  return false
}

export function ensureGenofficeLogin(_openUrl: (url: string) => void): void {}

export async function genofficeLogout(): Promise<void> {}

export function resetGenofficeAuthCache(): void {}

export function genofficeProxyFallbackPreferred(): boolean {
  return false
}
