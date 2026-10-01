/** Arithmetic subset of Excel Application.Evaluate, without executing JavaScript or workbook mutations. */
export function evaluateDvhExpression(input: string, decimalDigits = 2): string {
  const expression = input.trim()
  if (!expression) return '#VALUE: Chuỗi rỗng'
  const source = expression.startsWith('=') ? expression.slice(1) : expression
  const tokens: string[] = []
  const token = /\s*((?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?|[()+\-*/^])/y
  let offset = 0
  while (offset < source.length) {
    token.lastIndex = offset
    const match = token.exec(source)
    if (!match) return '#VALUE!'
    tokens.push(match[1]!)
    offset = token.lastIndex
  }
  let cursor = 0
  const take = (text: string): boolean => {
    if (tokens[cursor] !== text) return false
    cursor++
    return true
  }
  const primary = (): number => {
    if (take('(')) {
      const result = add()
      if (!take(')')) throw new Error('parenthesis')
      return result
    }
    const value = tokens[cursor++]
    if (!value || !/^(?:\d|\.)/.test(value)) throw new Error('number')
    return Number(value)
  }
  const power = (): number => {
    const left = primary()
    return take('^') ? left ** unary() : left
  }
  const unary = (): number => take('+') ? unary() : take('-') ? -unary() : power()
  const multiply = (): number => {
    let result = unary()
    while (tokens[cursor] === '*' || tokens[cursor] === '/') {
      const operator = tokens[cursor++]
      const right = unary()
      result = operator === '*' ? result * right : result / right
    }
    return result
  }
  const add = (): number => {
    let result = multiply()
    while (tokens[cursor] === '+' || tokens[cursor] === '-') {
      const operator = tokens[cursor++]
      const right = multiply()
      result = operator === '+' ? result + right : result - right
    }
    return result
  }
  try {
    const result = add()
    if (cursor !== tokens.length || !Number.isFinite(result)) return '#VALUE!'
    const digits = Math.max(0, Math.min(15, Math.trunc(decimalDigits)))
    return `${expression}=${Number(result.toFixed(digits))}`
  } catch {
    return '#VALUE!'
  }
}
