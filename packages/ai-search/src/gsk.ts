/** Legacy API compatibility for DVH Office. Genspark execution and credential access are removed. */
import {
  asRecord,
  firstItem,
  isCopyrightHost,
  safeHost,
  type ImageSearchResult,
  type WebSearchResult,
} from './shared'
export { setGskProxyUrl, gskProxyUrl } from './shared'
export const MAX_GSK_RESULTS = 20
export const MAX_GSK_SNIPPET_CHARS = 2_000

export function resolveGskEntry(): string | null {
  return null
}

export function gskApiKey(): string {
  return ''
}

export function watchGskApiKey(_onChange: (key: string) => void, _intervalMs = 2000): () => void {
  return () => {}
}

export function hasGskAuth(): boolean {
  return false
}

export function gskChildEnv(base: NodeJS.ProcessEnv = process.env): NodeJS.ProcessEnv {
  return { ...base }
}

export function parseGskOutput(stdout: string): unknown {
  const trimmed = stdout.trim()
  try {
    return JSON.parse(trimmed)
  } catch {
    /* fall through to line-by-line scan */
  }
  // Pretty-printed output puts inner elements on their own `{` lines, so the
  // scan must start from the earliest candidate and take the longest parse.
  const lines = trimmed.split('\n')
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i]!.trim()
    if (line.startsWith('{') || line.startsWith('[')) {
      for (let j = lines.length; j > i; j--) {
        try {
          return JSON.parse(lines.slice(i, j).join('\n'))
        } catch {
          continue
        }
      }
    }
  }
  throw new Error(`No JSON found in gsk output: ${stdout.slice(0, 300)}`)
}

function normalizeMaxResults(n: number): number {
  if (!Number.isFinite(n)) return 6
  return Math.min(20, Math.max(1, Math.floor(n)))
}

function clipField(v: unknown): string {
  const s = String(v ?? '')
  return s.length > MAX_GSK_SNIPPET_CHARS ? s.slice(0, MAX_GSK_SNIPPET_CHARS) : s
}

export function parseGskWebSearch(
  raw: unknown,
  maxResults: number,
): { results: WebSearchResult[]; answer?: string } {
  const bounded = normalizeMaxResults(maxResults)
  const data = asRecord(asRecord(raw).data ?? raw)
  const organic: unknown[] = Array.isArray(data.organic_results) ? data.organic_results : []
  const results: WebSearchResult[] = organic.slice(0, bounded).map((item) => {
    const o = asRecord(item)
    return {
      title: clipField(o.title),
      url: clipField(o.link),
      snippet: clipField(o.snippet),
    }
  })
  const answerRaw = typeof data.answer === 'string' && data.answer ? data.answer : undefined
  const answer =
    answerRaw !== undefined && answerRaw.length > MAX_GSK_SNIPPET_CHARS
      ? `${answerRaw.slice(0, MAX_GSK_SNIPPET_CHARS)}…`
      : answerRaw
  return answer !== undefined ? { results, answer } : { results }
}

export async function gskWebSearch(
  _query: string,
  _maxResults = 6,
): Promise<{ results: WebSearchResult[]; answer?: string }> {
  throw new Error(
    'Genspark integration is not available in DVH Office. Configure your own AI provider in Settings.',
  )
}

export function parseGskImageSearch(raw: unknown, maxResults: number): ImageSearchResult[] {
  const dataRaw = asRecord(raw).data
  const data: unknown[] = Array.isArray(dataRaw) ? dataRaw : []
  const images: ImageSearchResult[] = []
  for (const item of data) {
    const img = asRecord(item)
    const imageUrl = String(img.image_url ?? img.imageUrl ?? '')
    if (!imageUrl) continue
    if (isCopyrightHost(imageUrl)) continue
    const width = Number(img.width)
    const height = Number(img.height)
    const entry: ImageSearchResult = {
      title: String(img.title ?? ''),
      imageUrl,
      sourceUrl: String(img.link ?? ''),
      source: String(img.source ?? safeHost(img.link)),
    }
    if (Number.isFinite(width) && width > 0) entry.width = width
    if (Number.isFinite(height) && height > 0) entry.height = height
    images.push(entry)
    if (images.length >= maxResults) break
  }
  return images
}

export async function gskImageSearch(
  _query: string,
  _maxResults = 8,
): Promise<ImageSearchResult[]> {
  throw new Error(
    'Genspark integration is not available in DVH Office. Configure your own AI provider in Settings.',
  )
}

export interface GskGenerateImageOptions {
  /** Image description (English works better; text that must appear in the image stays verbatim) */
  prompt: string
  /** Generation model; defaults to the CLI default (nano-banana-2). Special-purpose: fal-bria-rmbg background removal, fal-ai/recraft-clarity-upscale upscaling, flux-pro/outpaint outpainting, fal-ai/image-editing/text-removal watermark text removal */
  model?: string
  /** Reference/edit-target image URLs (local paths also supported; the CLI uploads them automatically) */
  referenceImageUrls?: string[]
  /** 1:1 | 4:3 | 16:9 | 9:16 | 3:4 | 2:3 | 3:2 | auto */
  aspectRatio?: string
  /** auto | 0.5k | 1k | 2k | 3k | 4k */
  imageSize?: string
}

export interface GskGeneratedImage {
  /** Watermark-free image URL (preferred); can be downloaded and inserted directly */
  url: string
  taskId: string
}

export function parseGskGeneratedImage(raw: unknown): GskGeneratedImage {
  const data = asRecord(asRecord(raw).data ?? raw)
  const img = asRecord(firstItem(data.generated_images ?? data.images ?? []))
  const url = firstItem(img.image_urls_nowatermark) ?? firstItem(img.image_urls) ?? img.url ?? ''
  if (!url) {
    throw new Error(`gsk image generation returned no result: ${JSON.stringify(raw).slice(0, 200)}`)
  }
  return { url: String(url), taskId: String(img.task_id ?? '') }
}

export async function gskGenerateImage(
  _options: GskGenerateImageOptions,
  _signal?: AbortSignal,
): Promise<GskGeneratedImage> {
  throw new Error(
    'Genspark integration is not available in DVH Office. Configure your own AI provider in Settings.',
  )
}

export async function gskResolveDownloadUrl(url: string): Promise<string> {
  return url
}

export interface GskSlideGenerateOptions {
  /** Content and layout brief for this page */
  brief: string
  title?: string
  /** Deck-level visual system / typography / palette rules (Style Skill text) */
  styleSkill?: string
  /** Deck topic, page index, total pages, neighboring-page context */
  deckContext?: Record<string, unknown>
  /** HTTPS image candidates for this page */
  images?: { url: string; caption?: string }[]
  width?: number
  height?: number
  /** Slides model tier: ultra = opus-class model, standard (server default) = lighter model */
  tier?: 'standard' | 'ultra'
  signal?: AbortSignal
}

export function parseToolCliNdjson(text: string): Record<string, unknown> {
  const lines = text.trim().split('\n')
  for (let i = lines.length - 1; i >= 0; i--) {
    const line = lines[i]!.trim()
    if (!line.startsWith('{')) continue
    try {
      const obj: unknown = JSON.parse(line)
      if (obj && typeof obj === 'object' && 'status' in obj) return obj as Record<string, unknown>
    } catch {
      continue
    }
  }
  throw new Error(`No result line in tool_cli response: ${text.slice(0, 300)}`)
}

export async function gskSlideGenerate(
  _options: GskSlideGenerateOptions,
): Promise<{ bytes: Uint8Array; model: string }> {
  throw new Error(
    'Genspark integration is not available in DVH Office. Configure your own AI provider in Settings.',
  )
}

export function extractGskText(raw: unknown): string {
  const data = asRecord(raw).data ?? raw
  if (typeof data === 'string') return data
  if (data && typeof data === 'object') {
    for (const key of [
      'analysis',
      'analysis_result',
      'result',
      'content',
      'text',
      'answer',
      'transcript',
    ]) {
      const v = (data as Record<string, unknown>)[key]
      if (typeof v === 'string' && v.trim()) return v
    }
  }
  return JSON.stringify(data ?? raw)
}

export interface GskAnalyzeMediaOptions {
  /** Media URLs or local paths (image/audio/video; the CLI uploads local files automatically) */
  mediaUrls: string[]
  /** Analysis requirements (in English): what info to extract and what it's for */
  requirements: string
}

export async function gskAnalyzeMedia(
  _options: GskAnalyzeMediaOptions,
  _signal?: AbortSignal,
): Promise<string> {
  throw new Error(
    'Genspark integration is not available in DVH Office. Configure your own AI provider in Settings.',
  )
}

export interface GskTranscribeOptions {
  /** Audio URLs or local paths */
  audioUrls: string[]
  /** Prompt (context / proper nouns; can improve recognition quality) */
  prompt?: string
  /** whisper-1 (default, with timestamps) | gemini-3-flash-preview | elevenlabs_scribe_v2 */
  model?: string
}

export async function gskTranscribe(
  _options: GskTranscribeOptions,
  _signal?: AbortSignal,
): Promise<string> {
  throw new Error(
    'Genspark integration is not available in DVH Office. Configure your own AI provider in Settings.',
  )
}

export interface GskPastProject {
  projectId: string
  /** raw project type, e.g. 'slides_agent_git' */
  type: string
  title: string
  /** creation time, ISO-like string from the API */
  ctime: string
  /** relative web URL, e.g. '/agents?id=...' — join with https://www.genspark.ai */
  projectUrl: string
}

export interface GskPastProjectsPage {
  projects: GskPastProject[]
  total: number
  hasMore: boolean
}

export function parseGskPastProjects(raw: unknown): GskPastProjectsPage {
  const rec = asRecord(raw)
  const data = asRecord(rec.data ?? raw)
  const urlById = new Map<string, string>()
  const sessionProjects = asRecord(asRecord(rec.session_state).past_projects).projects
  if (Array.isArray(sessionProjects)) {
    for (const item of sessionProjects) {
      const p = asRecord(item)
      if (p.project_id && typeof p.project_url === 'string' && p.project_url) {
        urlById.set(String(p.project_id), p.project_url)
      }
    }
  }
  const listRaw: unknown[] = Array.isArray(data.projects) ? data.projects : []
  const projects: GskPastProject[] = []
  for (const item of listRaw) {
    const p = asRecord(item)
    const projectId = String(p.project_id ?? '')
    if (!projectId) continue
    projects.push({
      projectId,
      type: String(p.type ?? ''),
      title: String(p.title ?? ''),
      ctime: String(p.ctime ?? ''),
      projectUrl: urlById.get(projectId) ?? `/agents?id=${projectId}`,
    })
  }
  const total = Number(data.total)
  return {
    projects,
    total: Number.isFinite(total) ? total : projects.length,
    hasMore: data.has_more === true,
  }
}

export interface GskListPastProjectsOptions {
  /** 'slides' | 'docs' | 'sheets' | ...; omit for all kinds */
  artifactTypes?: string[]
  /** page size, CLI default 20, max 100 */
  limit?: number
  offset?: number
  signal?: AbortSignal
}

export async function gskListPastProjects(
  _options: GskListPastProjectsOptions = {},
): Promise<GskPastProjectsPage> {
  throw new Error(
    'Genspark integration is not available in DVH Office. Configure your own AI provider in Settings.',
  )
}

export interface GskLoginInfo {
  email: string
  plan: string
  creditBalance?: number
}

export async function gskLoginInfo(): Promise<GskLoginInfo | null> {
  return null
}
