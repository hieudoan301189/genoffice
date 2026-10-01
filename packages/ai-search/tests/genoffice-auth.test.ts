import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  loadGenofficeAuth,
  genofficeApiKey,
  startGenofficeLogin,
  ensureGenofficeLogin,
  genofficeLogout,
} from '../src/genoffice-auth'
import {
  gskApiKey,
  hasGskAuth,
  gskLoginInfo,
  gskListPastProjects,
  resolveGskEntry,
} from '../src/gsk'
afterEach(() => {
  vi.unstubAllGlobals()
  vi.unstubAllEnvs()
})
describe('removed account integration', () => {
  it('ignores legacy environment credentials and never opens login or requests the service', async () => {
    vi.stubEnv('GSK_API_KEY', 'test-legacy-key')
    vi.stubEnv('GSK_CLI_PATH', 'test-legacy-cli')
    const fetch = vi.fn()
    const open = vi.fn()
    vi.stubGlobal('fetch', fetch)
    expect(loadGenofficeAuth()).toBeNull()
    expect(genofficeApiKey()).toBe('')
    expect(gskApiKey()).toBe('')
    expect(hasGskAuth()).toBe(false)
    expect(resolveGskEntry()).toBeNull()
    expect(startGenofficeLogin(open)).toBe(false)
    ensureGenofficeLogin(open)
    await genofficeLogout()
    expect(await gskLoginInfo()).toBeNull()
    await expect(gskListPastProjects()).rejects.toThrow('not available')
    expect(open).not.toHaveBeenCalled()
    expect(fetch).not.toHaveBeenCalled()
  })
})
