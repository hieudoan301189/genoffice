import {
  BaseFunction,
  ArrayValueObject,
  BooleanValueObject,
  ErrorType,
  ErrorValueObject,
  FunctionType,
  IFunctionService,
  NumberValueObject,
  StringValueObject,
  type BaseValueObject,
  type BaseReferenceObject,
  type IFunctionInfo,
} from '@univerjs/engine-formula'
import { IDescriptionService } from '@univerjs/sheets-formula'
import { dvhLunarDay, dvhSolarDay } from './dvh-lunar'
import { evaluateDvhExpression } from './dvh-evaluate'
import { dvhLastUsed } from './dvh-last-used'
import {
  deleteDvhRows, insertDvhChild, lookupDvhMany, lookupDvhTwoKeys,
  sortDvhRows, subDvhArray, sumDvhExpression,
  type DvhMatrix,
} from './dvh-matrix-functions'
import { registerDvhAliases } from './dvh-function-aliases'
import type { UniverRuntime } from './univer-state'

// Ported from DVH-Excel/Function/clsFun_Text.cs, clsFun_Color.cs,
// clsFun_Table.cs and clsFun_Array.cs. These functions read only their arguments; workbook
// mutations and network-backed functions are deliberately separate.
type Scalar = string | number | boolean
type Evaluator = (args: readonly Scalar[]) => Scalar
interface DvhFunctionSpec {
  name: string
  params: readonly string[]
  required: number
  description: string
  evaluate: Evaluator
}

const string = (value: Scalar | undefined): string => value == null ? '' : String(value)
const numeric = (value: Scalar | undefined): number => Number(value)
const bool = (value: Scalar | undefined, fallback = false): boolean =>
  value === undefined ? fallback : value === true || value === 1 || String(value).toUpperCase() === 'TRUE'

function date(value: Scalar | undefined): Date | null {
  if (value === undefined || value === '') return null
  if (typeof value === 'number') {
    const d = new Date(Date.UTC(1899, 11, 30) + Math.trunc(value * 86_400_000))
    return Number.isNaN(d.getTime()) ? null : d
  }
  if (typeof value !== 'string') return null
  const iso = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(value)
  const local = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(value)
  if (!iso && !local) return null
  const year = Number(iso?.[1] ?? local?.[3])
  const month = Number(iso?.[2] ?? local?.[2])
  const day = Number(iso?.[3] ?? local?.[1])
  const d = new Date(Date.UTC(year, month - 1, day))
  return d.getUTCFullYear() === year && d.getUTCMonth() + 1 === month && d.getUTCDate() === day ? d : null
}

function roman(input: Scalar | undefined): string {
  if (input === undefined || input === '') return ''
  const n = numeric(input)
  if (!Number.isInteger(n)) return '#VALUE!'
  if (n < 1) return '#NUM! (Must be >= 1)'
  if (n > 3999) return '#NUM! (Must be <= 3999)'
  const pairs: readonly [number, string][] = [
    [1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'],
    [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I'],
  ]
  let rest = n
  let result = ''
  for (const [amount, numeral] of pairs) {
    while (rest >= amount) {
      result += numeral
      rest -= amount
    }
  }
  return result
}

function numberToColumn(input: Scalar | undefined): string {
  let n = numeric(input)
  if (!Number.isSafeInteger(n) || n < 1) return '#VALUE!'
  let result = ''
  while (n > 0) {
    n--
    result = String.fromCharCode(65 + n % 26) + result
    n = Math.floor(n / 26)
  }
  return result
}

function columnToNumber(input: Scalar | undefined): Scalar {
  const column = string(input).toUpperCase()
  if (!/^[A-Z]+$/.test(column)) return '#VALUE!'
  let n = 0
  for (const char of column) {
    n = n * 26 + char.charCodeAt(0) - 64
    if (!Number.isSafeInteger(n)) return '#VALUE!'
  }
  return n
}

function colorParts(args: readonly Scalar[]): number[] | null {
  const parts = args.slice(0, 3).map(numeric)
  return parts.length === 3 && parts.every((n) => Number.isInteger(n) && n >= 0 && n <= 255)
    ? parts : null
}

const digits = ['không', 'một', 'hai', 'ba', 'bốn', 'năm', 'sáu', 'bảy', 'tám', 'chín']
const numberUnits = ['', 'nghìn', 'triệu', 'tỷ', 'nghìn tỷ', 'triệu tỷ', 'tỷ tỷ']

function readThreeDigits(group: string, higherEmitted: boolean): string {
  const a = Number(group[0]), b = Number(group[1]), c = Number(group[2])
  const parts: string[] = []
  if (a > 0) parts.push(digits[a]!, 'trăm')
  if (b === 0 && c > 0 && (a > 0 || higherEmitted)) parts.push('lẻ')
  else if (b === 1) parts.push('mười')
  else if (b > 1) parts.push(digits[b]!, 'mươi')
  if (c > 0) parts.push(b > 1 && c === 1 ? 'mốt' : b > 0 && c === 5 ? 'lăm' : b > 1 && c === 4 ? 'tư' : digits[c]!)
  return parts.join(' ')
}

function readIntegerText(value: string): string {
  if (!value || /^0+$/.test(value)) return 'không'
  const groups = Math.ceil(value.length / 3)
  const padded = value.padStart(groups * 3, '0')
  const parts: string[] = []
  let higherEmitted = false
  for (let index = 0; index < groups; index++) {
    const group = padded.slice(index * 3, index * 3 + 3)
    if (Number(group) === 0) continue
    const unit = numberUnits[groups - index - 1]
    parts.push(`${readThreeDigits(group, higherEmitted)}${unit ? ` ${unit}` : ''}`)
    higherEmitted = true
  }
  return parts.join(', ')
}

function readNumber(args: readonly Scalar[]): string {
  const [value, before, after, integerUnit, decimalUnit] = args
  if (value === undefined || value === '') return 'Số không hợp lệ'
  const raw = string(value).trim()
  const comma = raw.lastIndexOf(','), dot = raw.lastIndexOf('.')
  const separator = Math.max(comma, dot)
  let integer = separator < 0 ? raw.replace(/[,. ]/g, '') : raw.slice(0, separator).replace(/[,. ]/g, '')
  let fraction = separator < 0 ? '' : raw.slice(separator + 1).replace(/[,.]/g, '').replace(/0+$/, '')
  const negative = integer.startsWith('-')
  if (negative) integer = integer.slice(1)
  if (!/^\d*$/.test(integer) || (fraction && !/^\d+$/.test(fraction))) return 'Số không hợp lệ'
  if (integer.length > 27) return 'Số quá lớn'
  if ((!integer || /^0+$/.test(integer)) && !fraction)
    return `${string(before)}không${string(integerUnit) ? ` ${string(integerUnit).trim()}` : ''}${string(after)}`
  let result = readIntegerText(integer)
  if (negative) result = `âm ${result}`
  if (integer !== '0' || fraction) result = result[0]!.toUpperCase() + result.slice(1)
  const parts = [string(before).trim(), result, string(integerUnit).trim()].filter(Boolean)
  const units = string(decimalUnit).trim().split(/[ ,/-]+/).filter(Boolean)
  if (fraction && units.length) {
    fraction = fraction.slice(0, 2)
    if (units.length >= 2) {
      for (let index = 0; index < fraction.length; index++) {
        const digit = Number(fraction[index])
        if (digit) parts.push(digits[digit]!, units[index]!)
      }
    } else if (Number(fraction)) parts.push(readIntegerText(fraction.replace(/^0+/, '')), units[0]!)
  }
  if (string(after).trim()) parts.push(string(after).trim())
  return parts.join(' ').replace(/\s+/g, ' ').trim()
}

function readTime(value: Scalar | undefined): string {
  let hour: number, minute: number
  if (typeof value === 'number') {
    const totalMinutes = Math.trunc(((value % 1 + 1) % 1) * 1440)
    hour = Math.floor(totalMinutes / 60)
    minute = totalMinutes % 60
  } else {
    const match = /^(\d{1,2}):(\d{1,2})(?::\d{1,2})?$/.exec(string(value))
    if (!match) return '#VALUE!'
    hour = Number(match[1]); minute = Number(match[2])
    if (hour > 23 || minute > 59) return '#VALUE!'
  }
  const hourText = readIntegerText(String(hour))
  if (minute === 0) return `${hourText} giờ`
  if (minute === 30) return `${hourText} giờ rưỡi`
  if (minute <= 30) return `${hourText} giờ ${readIntegerText(String(minute))} phút`
  return `${readIntegerText(String((hour + 1) % 24))} giờ kém ${readIntegerText(String(60 - minute))} phút`
}

function nthIndex(text: string, search: string, order: number): number {
  if (!Number.isInteger(order) || order < 1) return -1
  const lower = text.toLowerCase(), needle = search.toLowerCase()
  let from = 0, found = -1
  for (let count = 0; count < order; count++) {
    found = lower.indexOf(needle, from)
    if (found < 0) return -1
    from = found + search.length
  }
  return found
}

const specs: DvhFunctionSpec[] = [
  { name: 'DVH.ReadNumber', params: ['number', 'text_before', 'text_after', 'integer_unit', 'decimal_unit'],
    required: 1, description: 'Đọc số thành chữ tiếng Việt', evaluate: readNumber },
  { name: 'DVH.ReadTime', params: ['time_value'], required: 1,
    description: 'Đọc giờ phút bằng tiếng Việt', evaluate: ([value]) => readTime(value) },
  { name: 'DVH.UnMark', params: ['text'], required: 1, description: 'Bỏ dấu tiếng Việt',
    evaluate: ([value]) => string(value).replace(/[Đđ]/g, (c) => c === 'Đ' ? 'D' : 'd')
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '') },
  { name: 'DVH.FirstChar', params: ['text', 'is_upper'], required: 1, description: 'Lấy chữ đầu mỗi từ',
    evaluate: ([value, upper]) => string(value).trim().split(' ').filter(Boolean)
      .map((word) => bool(upper, true) ? word[0]!.toUpperCase() : word[0]!).join('') },
  { name: 'DVH.Roman', params: ['number'], required: 1, description: 'Đổi số thành số La Mã',
    evaluate: ([value]) => roman(value) },
  { name: 'DVH.GetString', params: ['text'], required: 1, description: 'Bỏ chữ số khỏi chuỗi',
    evaluate: ([value]) => string(value).replace(/\p{Nd}/gu, '') },
  { name: 'DVH.GetNumber', params: ['text'], required: 1, description: 'Lấy số đầu tiên trong chuỗi',
    evaluate: ([value]) => Number((/-?\d+(?:[.,]\d+)?/.exec(string(value))?.[0] ?? '0').replace(',', '.')) },
  { name: 'DVH.IsInclude', params: ['text', 'text_inside', 'case_sensitive'], required: 2,
    description: 'Kiểm tra chuỗi con', evaluate: ([a, b, sensitive]) => {
      const text = string(a), search = string(b)
      if (!text || !search) return false
      return bool(sensitive) ? text.includes(search) : text.toLowerCase().includes(search.toLowerCase())
    } },
  { name: 'DVH.CountText', params: ['text', 'find_text'], required: 2,
    description: 'Đếm số lần xuất hiện của chuỗi con', evaluate: ([a, b]) => {
      const text = string(a), search = string(b)
      if (!text || !search) return 0
      let count = 0, index = 0
      while ((index = text.indexOf(search, index)) >= 0) { count++; index += search.length }
      return count
    } },
  { name: 'DVH.Evaluate', params: ['text_or_expression', 'decimal_digits'], required: 1,
    description: 'Tính biểu thức số học và trả về biểu_thức=kết_quả',
    evaluate: ([expression, digits]) => evaluateDvhExpression(string(expression),
      digits === undefined || digits === '' ? 2 : numeric(digits)) },
  { name: 'DVH.GetText_XY', params: ['text', 'text_find_x', 'text_find_y', 'order_x', 'order_y'],
    required: 3, description: 'Lấy văn bản giữa hai chuỗi mốc', evaluate: ([text, x, y, orderX, orderY]) => {
      const source = string(text), start = string(x), end = string(y)
      if (!source || !start || !end) return '#N/A!'
      const posX = nthIndex(source, start, orderX === undefined ? 1 : numeric(orderX))
      const posY = nthIndex(source, end, orderY === undefined ? 1 : numeric(orderY))
      if (posX < 0 || posY < 0) return '#N/A!'
      const from = posX + start.length
      return posY < from ? '#N/A! (Y is before X)' : source.slice(from, posY)
    } },
  { name: 'DVH.InsertText', params: ['text', 'text_insert', 'location'], required: 3,
    description: 'Chèn văn bản vào vị trí hoặc trước một chuỗi mốc', evaluate: ([a, b, location]) => {
      const text = string(a), insert = string(b)
      if (!text) return insert
      if (typeof location === 'number') {
        const index = Math.max(0, Math.min(text.length, Math.trunc(location) - 1))
        return text.slice(0, index) + insert + text.slice(index)
      }
      const index = text.toLowerCase().indexOf(string(location).toLowerCase())
      return index < 0 ? text : text.slice(0, index) + insert + text.slice(index)
    } },
  { name: 'DVH.NumToCol', params: ['number'], required: 1, description: 'Đổi số thành tên cột',
    evaluate: ([value]) => numberToColumn(value) },
  { name: 'DVH.ColToNum', params: ['column_name'], required: 1, description: 'Đổi tên cột thành số',
    evaluate: ([value]) => columnToNumber(value) },
  { name: 'DVH.RGBtoNum', params: ['red', 'green', 'blue'], required: 3,
    description: 'Đổi RGB thành số màu Excel (BGR)', evaluate: (args) => {
      const parts = colorParts(args)
      return parts ? parts[2]! * 65536 + parts[1]! * 256 + parts[0]! : '#NUM!'
    } },
  { name: 'DVH.RGBtoHex', params: ['red', 'green', 'blue'], required: 3,
    description: 'Đổi RGB thành mã HEX', evaluate: (args) => {
      const parts = colorParts(args)
      return parts ? `#${parts.map((n) => n.toString(16).padStart(2, '0')).join('').toUpperCase()}` : '#NUM!'
    } },
  { name: 'DVH.NumToHex', params: ['number'], required: 1, description: 'Đổi số màu Excel thành HEX',
    evaluate: ([value]) => {
      const n = numeric(value)
      if (!Number.isInteger(n) || n < 0) return '#NUM!'
      return `#${[n % 256, Math.floor(n / 256) % 256, Math.floor(n / 65536) % 256]
        .map((part) => part.toString(16).padStart(2, '0')).join('').toUpperCase()}`
    } },
  { name: 'DVH.WeekdayVN', params: ['date_value'], required: 0, description: 'Tên thứ trong tuần bằng tiếng Việt',
    evaluate: ([value]) => {
      const d = value === undefined || value === '' ? new Date() : date(value)
      return d ? ['Chủ Nhật', 'Thứ Hai', 'Thứ Ba', 'Thứ Tư', 'Thứ Năm', 'Thứ Sáu', 'Thứ Bảy'][d.getUTCDay()]! : '#VALUE!'
    } },
  { name: 'DVH.DayInMonth', params: ['date_value'], required: 1, description: 'Số ngày trong tháng',
    evaluate: ([value]) => {
      const d = date(value)
      return d ? new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth() + 1, 0)).getUTCDate() : '#VALUE!'
    } },
  { name: 'DVH.WeekInMonth', params: ['date_value'], required: 1, description: 'Số tuần trong tháng',
    evaluate: ([value]) => {
      const d = date(value)
      if (!d) return '#VALUE!'
      const year = d.getUTCFullYear(), month = d.getUTCMonth()
      const firstDayOfWeek = new Date(Date.UTC(year, month, 1)).getUTCDay() + 1
      const lastDay = new Date(Date.UTC(year, month + 1, 0)).getUTCDate()
      return Math.ceil((firstDayOfWeek + lastDay - 1) / 7)
    } },
  { name: 'DVH.LunarDay', params: ['solar_date', 'time_zone'], required: 1,
    description: 'Chuyển ngày dương sang ngày âm', evaluate: ([value, zone]) => {
      const input = date(value)
      return input ? dvhLunarDay(input, zone === undefined ? 7 : numeric(zone)) : '#VALUE!'
    } },
  { name: 'DVH.SolarDay', params: ['lunar_date', 'is_leap_month', 'time_zone'], required: 1,
    description: 'Chuyển ngày âm sang ngày dương', evaluate: ([value, leap, zone]) => {
      const input = date(value)
      return input ? dvhSolarDay(input, bool(leap), zone === undefined ? 7 : numeric(zone)) : '#VALUE!'
    } },
]

export const DVH_PURE_FUNCTION_NAMES = specs.map((spec) => spec.name)

export function evaluateDvhPure(name: string, args: readonly Scalar[]): Scalar | undefined {
  return specs.find((spec) => spec.name.toUpperCase() === name.toUpperCase())?.evaluate(args)
}

function scalarArg(value: BaseValueObject): Scalar | null {
  let cell = value
  if (cell.isReferenceObject()) {
    const reference = cell as BaseValueObject & {
      getRangePosition(): { startRow: number; endRow: number; startColumn: number; endColumn: number }
      toArrayValueObject(): BaseValueObject
    }
    const range = reference.getRangePosition()
    if (range.startRow !== range.endRow || range.startColumn !== range.endColumn) return null
    cell = reference.toArrayValueObject()
  }
  if (cell.isArray()) {
    const array = cell as BaseValueObject & {
      getRowCount(): number; getColumnCount(): number; getFirstCell(): BaseValueObject
    }
    if (array.getRowCount() !== 1 || array.getColumnCount() !== 1) return null
    cell = array.getFirstCell()
  }
  return cell.isNull() ? '' : cell.getValue()
}

function matrixArg(value: BaseValueObject): Scalar[][] | null {
  let array = value
  if (value.isReferenceObject()) {
    const reference = value as BaseValueObject & {
      getRangePosition(): { startRow: number; endRow: number; startColumn: number; endColumn: number }
      toArrayValueObject(): ArrayValueObject
    }
    const range = reference.getRangePosition()
    if ((range.endRow - range.startRow + 1) * (range.endColumn - range.startColumn + 1) > 100_000)
      return null
    array = reference.toArrayValueObject()
  }
  if (!array.isArray()) {
    return [[array.isNull() ? '' : array.getValue()]]
  }
  const cells = array as ArrayValueObject
  const rows = cells.getRowCount(), columns = cells.getColumnCount()
  if (rows * columns > 100_000) return null
  return Array.from({ length: rows }, (_, row) =>
    Array.from({ length: columns }, (_, column) => {
      const cell = cells.get(row, column)
      return cell == null || cell.isNull() ? '' : cell.getValue()
    }))
}

function joinText(values: Scalar[][], delimiter: string, only: boolean, empty: boolean, order: number): string {
  if (!Number.isInteger(order) || order < 0 || order > 3) return '#N/A!'
  const rows = order >= 2 ? [...values].reverse() : values
  const result: string[] = []
  const seen = new Set<string>()
  for (const row of rows) {
    for (const cell of order % 2 ? [...row].reverse() : row) {
      const value = string(cell)
      if (!empty && !value.trim()) continue
      if (only && seen.has(value)) continue
      seen.add(value)
      result.push(value)
    }
  }
  return result.join(delimiter)
}

function joinTextIf(values: Scalar[][], include: Scalar[][] | null, delimiter: string,
  only: boolean, empty: boolean, down: boolean): string {
  if (include && include.length !== values.length)
    return '#ERROR: Include array must have same row count\nMảng include phải cùng số hàng'
  const result: string[] = []
  const seen = new Set<string>()
  const indexes = Array.from({ length: values.length }, (_, index) => index)
  if (!down) indexes.reverse()
  for (const row of indexes) {
    const flag = include?.[row]?.[0]
    if (include && (flag === undefined || flag === '' || flag === false || flag === 0 ||
      string(flag).trim().toUpperCase() === 'FALSE' || string(flag).trim() === '0')) continue
    const value = string(values[row]?.[0])
    if (!empty && !value.trim()) continue
    if (only && seen.has(value)) continue
    seen.add(value)
    result.push(value)
  }
  return result.join(delimiter)
}

interface DvhArrayFunctionSpec {
  name: string
  params: readonly string[]
  required: number
  description: string
}

const arraySpecs: DvhArrayFunctionSpec[] = [
  { name: 'DVH.JoinText', params: ['text_join', 'delimiter', 'is_only', 'is_empty', 'style_order'],
    required: 1, description: 'Nối các ô trong vùng theo thứ tự' },
  { name: 'DVH.JoinTextIF', params: ['area_join', 'delimiter', 'include', 'is_only', 'is_empty', 'is_order_down'],
    required: 1, description: 'Nối các ô theo mảng điều kiện' },
  { name: 'DVH.SplitText', params: ['text', 'delimiter', 'is_array_result', 'number'],
    required: 2, description: 'Tách chuỗi hoặc trả về mảng dọc' },
  { name: 'DVH.Array.GetRow', params: ['array', 'row_number'], required: 2,
    description: 'Lấy một hàng từ mảng' },
  { name: 'DVH.Array.GetColumn', params: ['array', 'column_number'], required: 2,
    description: 'Lấy một cột từ mảng' },
  { name: 'DVH.SumHight', params: ['array', 'k'], required: 2,
    description: 'Tổng k số lớn nhất trong mảng' },
  { name: 'DVH.SumLow', params: ['array', 'k'], required: 2,
    description: 'Tổng k số nhỏ nhất trong mảng' },
  { name: 'DVH.Array.DeleteRow', params: ['array', 'index_begin', 'index_end'], required: 3,
    description: 'Xóa hàng trong kết quả mảng' },
  { name: 'DVH.Array.QSortNum', params: ['array', 'column', 'ascending'], required: 2,
    description: 'Sắp xếp mảng theo cột số' },
  { name: 'DVH.Array.QSortText', params: ['array', 'column', 'ascending', 'ignore_case'], required: 2,
    description: 'Sắp xếp mảng theo cột văn bản' },
  { name: 'DVH.Array.InsertChild', params: ['parent', 'child', 'position', 'mode'], required: 3,
    description: 'Chèn mảng con vào mảng mẹ' },
  { name: 'DVH.Array.Sub', params: ['array', 'start_row', 'start_column', 'num_rows', 'num_columns'], required: 5,
    description: 'Lấy vùng con của mảng' },
  { name: 'DVH.LookUp_ManyRes', params: ['lookup_value', 'lookup_array', 'result_array', 'delimiter', 'is_unique'], required: 3,
    description: 'Dò nhiều kết quả và nối chuỗi' },
  { name: 'DVH.LookUp_ManyLookUp', params: ['result_array', 'lookup1', 'array1', 'lookup2', 'array2'], required: 5,
    description: 'Dò theo hai điều kiện' },
  { name: 'DVH.Sum', params: ['range'], required: 1,
    description: 'Biểu diễn các số trong vùng thành chuỗi phép cộng' },
]

function sumEdge(values: Scalar[][], k: number, largest: boolean): number {
  const numbers = values.flat().flatMap((value) => {
    if (value === '' || typeof value === 'boolean') return []
    const parsed = typeof value === 'number' ? value : Number(value.trim().replace(',', '.'))
    return Number.isFinite(parsed) ? [parsed] : []
  })
  numbers.sort((a, b) => largest ? b - a : a - b)
  return numbers.slice(0, Math.max(0, Math.trunc(k))).reduce((sum, value) => sum + value, 0)
}

class DvhArrayFunction extends BaseFunction {
  override minParams: number
  override maxParams: number
  constructor(private readonly spec: DvhArrayFunctionSpec) {
    super(spec.name.toUpperCase())
    this.minParams = spec.required
    this.maxParams = spec.params.length
  }
  override calculate(...args: BaseValueObject[]): ReturnType<BaseFunction['calculate']> {
    for (const arg of args) if (arg.isError()) return arg
    const option = (index: number): Scalar | null => args[index] ? scalarArg(args[index]!) : ''
    const matrix = (index: number): DvhMatrix | null => args[index] ? matrixArg(args[index]!) : null
    const arrayResult = (value: DvhMatrix): BaseValueObject =>
      value.length * (value[0]?.length ?? 0) > 100_000
        ? ErrorValueObject.create(ErrorType.VALUE) : ArrayValueObject.createByArray(value)
    const scalarResult = (value: Scalar): BaseValueObject =>
      typeof value === 'number' ? NumberValueObject.create(value)
        : typeof value === 'boolean' ? BooleanValueObject.create(value)
          : StringValueObject.create(value)
    if (this.spec.name === 'DVH.SplitText') {
      const source = option(0), delimiter = option(1)
      if (source === null || delimiter === null) return ErrorValueObject.create(ErrorType.VALUE)
      if (!string(source)) return StringValueObject.create('')
      const parts = string(delimiter) ? string(source).split(string(delimiter)) : [string(source)]
      if (bool(option(2) ?? undefined)) return ArrayValueObject.createByArray(parts.map((part) => [part]))
      const index = Math.max(1, Math.trunc(numeric(option(3) || 1)))
      return StringValueObject.create(parts[index - 1] ?? '')
    }
    if (this.spec.name === 'DVH.LookUp_ManyRes') {
      const lookup = option(0), keys = matrix(1), results = matrix(2)
      const delimiter = option(3), unique = option(4)
      if (lookup === null || !keys || !results || delimiter === null || unique === null)
        return ErrorValueObject.create(ErrorType.VALUE)
      return StringValueObject.create(lookupDvhMany(lookup, keys, results,
        args[3] === undefined ? ', ' : string(delimiter), bool(unique)))
    }
    const values = matrix(0)
    if (!values) return ErrorValueObject.create(ErrorType.VALUE)
    if (this.spec.name === 'DVH.LookUp_ManyLookUp') {
      const key1 = option(1), array1 = matrix(2), key2 = option(3), array2 = matrix(4)
      if (key1 === null || !array1 || key2 === null || !array2)
        return ErrorValueObject.create(ErrorType.VALUE)
      return scalarResult(lookupDvhTwoKeys(values, key1, array1, key2, array2))
    }
    if (this.spec.name === 'DVH.Array.InsertChild') {
      const child = matrix(1), position = option(2), mode = option(3)
      if (!child || position === null || mode === null)
        return ErrorValueObject.create(ErrorType.VALUE)
      return arrayResult(insertDvhChild(values, child, numeric(position), args[3] === undefined ? 1 : numeric(mode)))
    }
    if (this.spec.name === 'DVH.Array.DeleteRow') {
      const begin = option(1), end = option(2)
      if (begin === null || end === null) return ErrorValueObject.create(ErrorType.VALUE)
      return arrayResult(deleteDvhRows(values, numeric(begin), numeric(end)))
    }
    if (this.spec.name === 'DVH.Array.Sub') {
      const numbers = [1, 2, 3, 4].map(option)
      if (numbers.some((value) => value === null)) return ErrorValueObject.create(ErrorType.VALUE)
      return arrayResult(subDvhArray(values, ...numbers.map((value) => numeric(value!)) as [number, number, number, number]))
    }
    if (this.spec.name === 'DVH.Array.QSortNum' || this.spec.name === 'DVH.Array.QSortText') {
      const column = option(1), ascending = option(2), ignoreCase = option(3)
      if (column === null || ascending === null || ignoreCase === null)
        return ErrorValueObject.create(ErrorType.VALUE)
      return arrayResult(sortDvhRows(values, numeric(column), args[2] === undefined ? true : bool(ascending),
        this.spec.name === 'DVH.Array.QSortText', args[3] === undefined ? true : bool(ignoreCase)))
    }
    if (this.spec.name === 'DVH.Sum') return StringValueObject.create(sumDvhExpression(values))
    if (this.spec.name === 'DVH.Array.GetRow' || this.spec.name === 'DVH.Array.GetColumn') {
      const index = option(1)
      if (index === null) return ErrorValueObject.create(ErrorType.VALUE)
      const n = numeric(index)
      const rowFunction = this.spec.name === 'DVH.Array.GetRow'
      const size = rowFunction ? values.length : values[0]!.length
      if (!Number.isInteger(n) || n < 1 || n > size)
        return StringValueObject.create(`#ERROR: ${rowFunction ? 'startRow' : 'startColumn'} must be 1-${size}`)
      return ArrayValueObject.createByArray(rowFunction
        ? [values[n - 1]!] : values.map((row) => [row[n - 1]!]))
    }
    if (this.spec.name === 'DVH.SumHight' || this.spec.name === 'DVH.SumLow') {
      const k = option(1)
      if (k === null || !Number.isFinite(numeric(k))) return ErrorValueObject.create(ErrorType.VALUE)
      return NumberValueObject.create(sumEdge(values, numeric(k), this.spec.name === 'DVH.SumHight'))
    }
    const delimiter = option(1)
    if (delimiter === null) return ErrorValueObject.create(ErrorType.VALUE)
    const separator = args[1] === undefined ? ',' : string(delimiter)
    if (this.spec.name === 'DVH.JoinText') {
      const only = option(2), empty = option(3), order = option(4)
      if (only === null || empty === null || order === null) return ErrorValueObject.create(ErrorType.VALUE)
      return StringValueObject.create(joinText(values, separator, bool(only), bool(empty), numeric(order || 0)))
    }
    const hasInclude = Boolean(args[2] && !args[2].isNull())
    const include = hasInclude ? matrixArg(args[2]!) : null
    const only = option(3), empty = option(4), down = option(5)
    if ((hasInclude && !include) || only === null || empty === null || down === null)
      return ErrorValueObject.create(ErrorType.VALUE)
    return StringValueObject.create(joinTextIf(values, include, separator, bool(only), bool(empty),
      args[5] === undefined ? true : bool(down)))
  }
}

export function evaluateDvhArray(name: string, values: Scalar[][], options: readonly Scalar[] = [], include?: Scalar[][]): string {
  if (name.toUpperCase() === 'DVH.JOINTEXT')
    return joinText(values, options[0] === undefined ? ',' : string(options[0]), bool(options[1]), bool(options[2]), numeric(options[3] || 0))
  if (name.toUpperCase() === 'DVH.JOINTEXTIF')
    return joinTextIf(values, include ?? null, options[0] === undefined ? ',' : string(options[0]),
      bool(options[1]), bool(options[2]), options[3] === undefined ? true : bool(options[3]))
  throw new Error(`Unknown DVH array function: ${name}`)
}

class DvhPureFunction extends BaseFunction {
  override minParams: number
  override maxParams: number
  constructor(private readonly spec: DvhFunctionSpec) {
    // Univer uppercases parsed function tokens before looking up executors.
    super(spec.name.toUpperCase())
    this.minParams = spec.required
    this.maxParams = spec.params.length
  }
  override calculate(...args: BaseValueObject[]): ReturnType<BaseFunction['calculate']> {
    const values: Scalar[] = []
    for (const arg of args) {
      if (arg.isError()) return arg
      const value = scalarArg(arg)
      if (value === null) return ErrorValueObject.create(ErrorType.VALUE)
      values.push(value)
    }
    const result = this.spec.evaluate(values)
    if (this.spec.name === 'DVH.LunarDay' || this.spec.name === 'DVH.SolarDay') {
      if (result === '#NUM!') return ErrorValueObject.create(ErrorType.NUM)
      if (result === '#VALUE!') return ErrorValueObject.create(ErrorType.VALUE)
    }
    return typeof result === 'number' ? NumberValueObject.create(result)
      : typeof result === 'boolean' ? BooleanValueObject.create(result)
        : StringValueObject.create(result)
  }
}

const lastUsedSpecs: Pick<DvhFunctionSpec, 'name' | 'params' | 'required' | 'description'>[] = [
  { name: 'DVH.LastRow', params: ['range'], required: 1,
    description: 'Dòng cuối có dữ liệu trong cột của vùng tham chiếu' },
  { name: 'DVH.LastCol', params: ['range'], required: 1,
    description: 'Cột cuối có dữ liệu trong hàng của vùng tham chiếu' },
]

class DvhLastUsedFunction extends BaseFunction {
  override minParams = 1
  override maxParams = 1
  override needsReferenceObject = true
  constructor(private readonly axis: 'row' | 'column', private readonly isLoaded: () => boolean) {
    super(axis === 'row' ? 'DVH.LASTROW' : 'DVH.LASTCOL')
  }
  override calculate(value: BaseValueObject): ReturnType<BaseFunction['calculate']> {
    if (value.isError()) return value
    if (!value.isReferenceObject() || !this.isLoaded()) return ErrorValueObject.create(ErrorType.NA)
    return NumberValueObject.create(dvhLastUsed(value as unknown as BaseReferenceObject, this.axis))
  }
}

export function installDvhPureFunctions(runtime: UniverRuntime, isLoaded: () => boolean): { dispose(): void } {
  const injector = runtime.univer.__getInjector()
  const functions = injector.get(IFunctionService)
  const descriptions = injector.get(IDescriptionService)
  const executors = [
    ...specs.map((spec) => new DvhPureFunction(spec)),
    ...arraySpecs.map((spec) => new DvhArrayFunction(spec)),
    new DvhLastUsedFunction('row', isLoaded),
    new DvhLastUsedFunction('column', isLoaded),
  ]
  functions.registerExecutors(...executors)
  const infos = [...specs, ...arraySpecs, ...lastUsedSpecs].map((spec): IFunctionInfo => ({
    functionName: spec.name.toUpperCase(),
    functionType: FunctionType.User,
    abstract: spec.description,
    description: `DVH Tool · ${spec.description}`,
    functionParameter: spec.params.map((name, index) => ({
      name, detail: name, example: '', require: index < spec.required ? 1 : 0, repeat: 0,
    })),
  }))
  const descriptionHandle = descriptions.registerDescriptions(infos)
  // DVH.Table and Table: every function also answers to its short name
  const aliases = registerDvhAliases(injector, executors, infos)
  return {
    dispose() {
      aliases.dispose()
      descriptionHandle.dispose()
      functions.unregisterExecutors(...executors.map((executor) => executor.name))
    },
  }
}
