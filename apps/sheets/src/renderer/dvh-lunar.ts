// Port of the worksheet conversion algorithms in
// DVH-Excel/Function/clsFun_Lunar_Solar.cs (the Ho Ngoc Duc/VB formula).
const PI = 3.14159265358979
const DR = PI / 180
const EPOCH = 2415021.07699869
const MONTH = 29.530588853
const MS_PER_DAY = 86_400_000
const oaEpoch = Date.UTC(1899, 11, 30)

function jdFromDate(day: number, month: number, year: number): number {
  const a = Math.floor((14 - month) / 12)
  const y = year + 4800 - a
  const m = month + 12 * a - 3
  let jd = day + Math.floor((153 * m + 2) / 5) + 365 * y + Math.floor(y / 4)
    - Math.floor(y / 100) + Math.floor(y / 400) - 32045
  if (jd < 2299161)
    jd = day + Math.floor((153 * m + 2) / 5) + 365 * y + Math.floor(y / 4) - 32083
  return jd
}

function jdToDate(jd: number): Date {
  let b: number, c: number
  if (jd > 2299160) {
    const a = jd + 32044
    b = Math.floor((4 * a + 3) / 146097)
    c = a - Math.floor(b * 146097 / 4)
  } else {
    b = 0
    c = jd + 32082
  }
  const d = Math.floor((4 * c + 3) / 1461)
  const e = c - Math.floor(1461 * d / 4)
  const m = Math.floor((5 * e + 2) / 153)
  return new Date(Date.UTC(b * 100 + d - 4800 + Math.floor(m / 10),
    m + 3 - 12 * Math.floor(m / 10) - 1,
    e - Math.floor((153 * m + 2) / 5) + 1))
}

function newMoon(k: number): number {
  const t = k / 1236.85, t2 = t * t, t3 = t2 * t
  let jd1 = 2415020.75933 + 29.53058868 * k + 0.0001178 * t2 - 0.000000155 * t3
  jd1 += 0.00033 * Math.sin((166.56 + 132.87 * t - 0.009173 * t2) * DR)
  const m = 359.2242 + 29.10535608 * k - 0.0000333 * t2 - 0.00000347 * t3
  const mp = 306.0253 + 385.81691806 * k + 0.0107306 * t2 + 0.00001236 * t3
  const f = 21.2964 + 390.67050646 * k - 0.0016528 * t2 - 0.00000239 * t3
  let c1 = (0.1734 - 0.000393 * t) * Math.sin(m * DR) + 0.0021 * Math.sin(2 * DR * m)
  c1 = c1 - 0.4068 * Math.sin(mp * DR) + 0.0161 * Math.sin(DR * 2 * mp)
  c1 = c1 - 0.0004 * Math.sin(DR * 3 * mp)
  c1 = c1 + 0.0104 * Math.sin(DR * 2 * f) - 0.0051 * Math.sin(DR * (m + mp))
  c1 = c1 - 0.0074 * Math.sin(DR * (m - mp)) + 0.0004 * Math.sin(DR * (2 * f + m))
  c1 = c1 - 0.0004 * Math.sin(DR * (2 * f - m)) - 0.0006 * Math.sin(DR * (2 * f + mp))
  c1 = c1 + 0.001 * Math.sin(DR * (2 * f - mp)) + 0.0005 * Math.sin(DR * (2 * mp + m))
  const deltaT = t < -11
    ? 0.001 + 0.000839 * t + 0.0002261 * t2 - 0.00000845 * t3 - 0.000000081 * t * t3
    : -0.000278 + 0.000265 * t + 0.000262 * t2
  return jd1 + c1 - deltaT
}

function sunLongitude(jdn: number): number {
  const t = (jdn - 2451545) / 36525, t2 = t * t
  const m = 357.5291 + 35999.0503 * t - 0.0001559 * t2 - 0.00000048 * t * t2
  const l0 = 280.46645 + 36000.76983 * t + 0.0003032 * t2
  let dl = (1.9146 - 0.004817 * t - 0.000014 * t2) * Math.sin(m * DR)
  dl += (0.019993 - 0.000101 * t) * Math.sin(2 * m * DR) + 0.00029 * Math.sin(3 * m * DR)
  const radians = (l0 + dl) * DR
  return radians - 2 * PI * Math.floor(radians / (2 * PI))
}

const newMoonDay = (k: number, zone: number): number => Math.floor(newMoon(k) + 0.5 + zone / 24)
const sunSector = (day: number, zone: number): number =>
  Math.floor(sunLongitude(day - 0.5 - zone / 24) / PI * 6)

function monthStartK(day: number, zone: number): number {
  let k = Math.trunc((day - EPOCH) / MONTH) + 1
  while (newMoonDay(k, zone) > day) k--
  return k
}

function lunarMonth11(year: number, zone: number): number {
  const off = jdFromDate(31, 12, year) - 2415021
  const k = Math.floor(off / MONTH)
  let moon = newMoonDay(k, zone)
  if (sunSector(moon, zone) >= 9) moon = newMoonDay(k - 1, zone)
  return moon
}

function leapMonthOffset(a11: number, zone: number): number {
  const k = Math.floor((a11 - EPOCH) / MONTH + 0.5)
  let i = 1, last = 0, arc = sunSector(newMoonDay(k + i, zone), zone)
  do {
    last = arc
    i++
    arc = sunSector(newMoonDay(k + i, zone), zone)
  } while (arc !== last && i < 14)
  return i - 1
}

const validZoneAndYear = (date: Date, zone: number): boolean =>
  Number.isInteger(zone) && zone >= -12 && zone <= 14 &&
  date.getUTCFullYear() >= 1800 && date.getUTCFullYear() <= 2200
const serial = (date: Date): number => Math.round((date.getTime() - oaEpoch) / MS_PER_DAY)
const twoDigits = (n: number): string => String(n).padStart(2, '0')

export function dvhSolarDay(lunarDate: Date, leap = false, zone = 7): number | '#NUM!' | '#VALUE!' {
  if (!validZoneAndYear(lunarDate, zone)) return '#NUM!'
  const day = lunarDate.getUTCDate(), month = lunarDate.getUTCMonth() + 1, year = lunarDate.getUTCFullYear()
  if (day < 1 || day > 30) return '#VALUE!'
  let a11: number, b11: number
  if (month < 11) {
    a11 = lunarMonth11(year - 1, zone)
    b11 = lunarMonth11(year, zone)
  } else {
    a11 = lunarMonth11(year, zone)
    b11 = lunarMonth11(year + 1, zone)
  }
  const k = Math.floor(0.5 + (a11 - EPOCH) / MONTH)
  let offset = month - 11
  if (offset < 0) offset += 12
  if (b11 - a11 > 365) {
    const leapOffset = leapMonthOffset(a11, zone)
    let leapMonth = leapOffset - 2
    if (leapMonth < 0) leapMonth += 12
    if (leap && month !== leapMonth) return '#VALUE!'
    if (leap || offset >= leapOffset) offset++
  } else if (leap) return '#VALUE!'
  const target = jdToDate(newMoonDay(k + offset, zone) + day - 1)
  return Number.isNaN(target.getTime()) ? '#VALUE!' : serial(target)
}

export function dvhLunarDay(solarDate: Date, zone = 7): number | string {
  if (!validZoneAndYear(solarDate, zone)) return '#NUM!'
  const day = solarDate.getUTCDate(), month = solarDate.getUTCMonth() + 1, year = solarDate.getUTCFullYear()
  const dayNumber = jdFromDate(day, month, year)
  const monthStart = newMoonDay(monthStartK(dayNumber, zone), zone)
  let a11 = lunarMonth11(year, zone), b11 = a11, lunarYear: number
  if (a11 >= monthStart) {
    lunarYear = year
    a11 = lunarMonth11(year - 1, zone)
  } else {
    lunarYear = year + 1
    b11 = lunarMonth11(year + 1, zone)
  }
  const lunarDay = dayNumber - monthStart + 1
  const diff = Math.trunc((monthStart - a11) / 29)
  let lunarMonth = diff + 11, isLeap = false
  if (b11 - a11 > 365) {
    const leapDiff = leapMonthOffset(a11, zone)
    if (diff >= leapDiff) {
      lunarMonth = diff + 10
      isLeap = diff === leapDiff
    }
  }
  if (lunarMonth > 12) lunarMonth -= 12
  if (lunarMonth >= 11 && diff < 4) lunarYear--
  if (lunarMonth < 1 || lunarMonth > 12 || lunarDay < 1 || lunarDay > 30) return '#VALUE!'
  const formatted = `${twoDigits(lunarDay)}/${twoDigits(lunarMonth)}/${lunarYear}`
  if (isLeap) return `Nhuận ${formatted}`
  const asDate = new Date(Date.UTC(lunarYear, lunarMonth - 1, lunarDay))
  if (asDate.getUTCDate() !== lunarDay || asDate.getUTCMonth() + 1 !== lunarMonth)
    return formatted
  return serial(asDate)
}
