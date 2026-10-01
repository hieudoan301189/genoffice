import { describe, expect, it } from 'vitest'
import { dvhLunarDay, dvhSolarDay } from '../src/renderer/dvh-lunar'

describe('DVH lunar and solar dates', () => {
  it('matches both examples in the DVH-Excel function catalog', () => {
    expect(dvhLunarDay(new Date('2024-02-10T00:00:00Z'))).toBe(45292) // 01/01/2024
    expect(dvhSolarDay(new Date('2024-01-01T00:00:00Z'))).toBe(45332) // 10/02/2024
  })

  it('handles leap lunar months as text and rejects a mismatched leap flag', () => {
    const leapSolar = dvhSolarDay(new Date('2023-02-01T00:00:00Z'), true)
    expect(leapSolar).toBe(45007) // 22/03/2023
    expect(dvhLunarDay(new Date('2023-03-22T00:00:00Z'))).toBe('Nhuận 01/02/2023')
    expect(dvhSolarDay(new Date('2024-01-01T00:00:00Z'), true)).toBe('#VALUE!')
  })

  it('keeps the original year and time-zone limits', () => {
    expect(dvhLunarDay(new Date('2024-02-10T00:00:00Z'), 15)).toBe('#NUM!')
    expect(dvhSolarDay(new Date('1700-01-01T00:00:00Z'))).toBe('#NUM!')
  })
})
