export { ProjectStore, canonicalPathKey, MAX_PROJECT_DATA_BYTES } from './store.js'
export type {
  ChatMessage,
  ChatMeta,
  ChatScope,
  ProjectData,
  ProjectIndex,
  ProjectInfo,
  ProjectSummary,
  TimelineEntry,
  ToolActivity,
} from './types.js'
export type {
  AppendChatArgs,
  LoadChatArgs,
  ProjectApi,
  RebindChatArgs,
  ResolveChatArgs,
  ResolveChatResult,
} from './ipc.js'
