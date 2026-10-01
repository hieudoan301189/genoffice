import { execFile } from 'node:child_process'
import { existsSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { promisify } from 'node:util'
import { app, safeStorage } from 'electron'

const execFileAsync = promisify(execFile)
const keyPath = () => join(app.getPath('userData'), 'gemini-key.bin')

/** DVH-Tool stores Gemini in HKCU with DPAPI CurrentUser and no entropy. */
async function readDvhToolKey(): Promise<string> {
  if (process.platform !== 'win32') return ''
  const script = `
$key = [Microsoft.Win32.Registry]::CurrentUser.OpenSubKey('Software\\DVH_Tool\\ApiKeys')
if ($null -eq $key) { exit }
$encoded = $key.GetValue('Gemini')
if ([string]::IsNullOrWhiteSpace($encoded)) { exit }
try {
  Add-Type -AssemblyName System.Security
  $bytes = [Convert]::FromBase64String($encoded)
  $plain = [Security.Cryptography.ProtectedData]::Unprotect($bytes, $null, [Security.Cryptography.DataProtectionScope]::CurrentUser)
  [Console]::Out.Write([Text.Encoding]::UTF8.GetString($plain))
} catch { exit }
`
  try {
    const { stdout } = await execFileAsync('powershell.exe', [
      '-NoProfile', '-NonInteractive', '-EncodedCommand', Buffer.from(script, 'utf16le').toString('base64'),
    ], { windowsHide: true, timeout: 5_000, maxBuffer: 16_384 })
    return stdout.trim()
  } catch {
    return ''
  }
}

export async function loadDvhGeminiKey(): Promise<string> {
  if (existsSync(keyPath())) {
    try {
      return safeStorage.decryptString(readFileSync(keyPath())).trim()
    } catch {
      return ''
    }
  }
  return readDvhToolKey()
}

export function saveDvhGeminiKey(apiKey: string): void {
  if (!safeStorage.isEncryptionAvailable()) throw new Error('Cannot encrypt Gemini API key on this device.')
  writeFileSync(keyPath(), safeStorage.encryptString(apiKey.trim()))
}
