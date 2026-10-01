import { describe, it, expect, vi } from 'vitest'
vi.mock('../src/gsk', () => ({
  gskGenerateImage: vi.fn(),
  gskAnalyzeMedia: vi.fn(),
  hasGskAuth: vi.fn(() => true),
}))
import { generateImageTool } from '../src/media-tools'
import { gskGenerateImage } from '../src/gsk'
describe('removed image service', () => {
  it.each([true, false])(
    'never invokes the old service (transparent: %s)',
    async (transparentBackground) => {
      const result = await generateImageTool('/nonexistent/ai-settings.json', {
        prompt: 'test',
        transparentBackground,
      })
      expect(result.error).toBeTruthy()
      expect(gskGenerateImage).not.toHaveBeenCalled()
    },
  )
})
