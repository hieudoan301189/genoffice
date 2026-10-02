/**
 * Short names for DVH worksheet functions: every `DVH.X` executor is also
 * callable as `X` (DVH.Table → TABLE, DVH.Font.Color → FONT.COLOR).
 *
 * The alias is the same executor seen through a prototype link with its own
 * name, so per-call state (unitId, row, column) stays on the alias while the
 * function's logic and stores are shared. A short name that the engine
 * already knows (SUM, ROMAN, LOOKUP...) or that Excel reserves keeps only its
 * DVH. form: a workbook must never change what a built-in function means.
 */
import {
  FunctionType,
  IFunctionService,
  type BaseFunction,
  type IFunctionInfo,
} from '@univerjs/engine-formula'
import { IDescriptionService } from '@univerjs/sheets-formula'
import type { Injector } from '@univerjs/core'

const PREFIX = 'DVH.'

/**
 * Excel functions the engine may not implement; a short DVH alias must not
 * shadow them (the file would mean something else in Excel).
 */
const EXCEL_RESERVED = new Set([
  'TRANSLATE',
  'DETECTLANGUAGE',
  'FILTER',
  'LOOKUP',
  'SUM',
  'ROMAN',
  'DSUM',
  'DCOUNT',
  'DAVERAGE',
  'DMAX',
  'DMIN',
  'IMAGE',
  'WEBSERVICE',
  'FILTERXML',
  'ENCODEURL',
  'GETPIVOTDATA',
  'CUBEVALUE',
  'INFO',
  'CELL',
])

/** Short names registered in this session (upper case). */
const registeredAliases = new Set<string>()

/** `DVH.TABLE` → `TABLE`; null when the name is not a DVH function. */
export function dvhShortName(name: string): string | null {
  const upper = name.toUpperCase()
  return upper.startsWith(PREFIX) && upper.length > PREFIX.length
    ? upper.slice(PREFIX.length)
    : null
}

/** True for `DVH.X` and for the short aliases this session registered. */
export function isDvhFunctionName(name: string): boolean {
  const upper = name.toUpperCase()
  return upper.startsWith(PREFIX) || registeredAliases.has(upper)
}

/** True when a formula calls a DVH function under either name. */
export function formulaCallsDvh(formula: string): boolean {
  for (const match of formula.matchAll(/([A-Za-z_][A-Za-z0-9_.]*)\s*\(/g)) {
    if (isDvhFunctionName(match[1]!)) return true
  }
  return false
}

/** Same executor, own name; the engine keys executors by `name`. */
function aliasExecutor(executor: BaseFunction, alias: string): BaseFunction {
  const aliased = Object.create(executor) as BaseFunction
  ;(aliased as unknown as { _name: string })._name = alias
  return aliased
}

/**
 * Registers the short names of `executors` (with descriptions copied from
 * `infos`) and returns what to undo. Call after the DVH. names are registered.
 */
export function registerDvhAliases(
  injector: Injector,
  executors: readonly BaseFunction[],
  infos: readonly IFunctionInfo[],
): { dispose(): void } {
  const functions = injector.get(IFunctionService)
  const descriptions = injector.get(IDescriptionService)
  const infoByName = new Map(infos.map((info) => [info.functionName.toUpperCase(), info]))
  const aliases: BaseFunction[] = []
  const aliasInfos: IFunctionInfo[] = []
  for (const executor of executors) {
    const full = String(executor.name).toUpperCase()
    const short = dvhShortName(full)
    if (!short || EXCEL_RESERVED.has(short) || functions.hasExecutor(short)) continue
    if (descriptions.hasDescription(short)) continue
    aliases.push(aliasExecutor(executor, short))
    registeredAliases.add(short)
    const info = infoByName.get(full)
    aliasInfos.push({
      functionName: short,
      functionType: info?.functionType ?? FunctionType.User,
      abstract: info?.abstract ?? full,
      description: `${info?.description ?? 'DVH Tool'} (= ${full})`,
      functionParameter: info?.functionParameter ?? [],
    })
  }
  functions.registerExecutors(...aliases)
  const handle = descriptions.registerDescriptions(aliasInfos)
  return {
    dispose() {
      handle.dispose()
      functions.unregisterExecutors(...aliases.map((alias) => alias.name))
      for (const alias of aliases) registeredAliases.delete(String(alias.name))
    },
  }
}
