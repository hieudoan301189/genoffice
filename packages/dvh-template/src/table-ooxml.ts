/**
 * A rendered DVH.Table as WordprocessingML (P6 generation): the same grid
 * Docs shows (dvh-model renderTable), written straight to `<w:tbl>` so
 * documents can be generated without an editor. Colors are document data.
 */
import {
  formatCellValue,
  type DvhCellStyle,
  type DvhTable,
  type RenderedTable,
} from '@genoffice/dvh-model'
import { escapeText } from './ooxml'

/** text column width of an A4 page with 2 cm margins, in twips */
const PAGE_TEXT_TWIPS = 9638

const hex = (color: string | undefined) => color?.replace(/^#/, '').toUpperCase()

function cellXml(
  value: string,
  style: DvhCellStyle,
  widthTwips: number,
  span: number,
  align: string | undefined,
): string {
  const borders = style.border
    ? `<w:tcBorders>${['top', 'left', 'bottom', 'right']
        .map(
          (side) =>
            `<w:${side} w:val="single" w:sz="${style.border!.width === 'medium' ? 8 : 4}" w:space="0" w:color="${hex(style.border!.color)}"/>`,
        )
        .join('')}</w:tcBorders>`
    : ''
  const fill = style.fill ? `<w:shd w:val="clear" w:color="auto" w:fill="${hex(style.fill)}"/>` : ''
  const tcPr =
    `<w:tcPr><w:tcW w:w="${Math.round(widthTwips)}" w:type="dxa"/>` +
    (span > 1 ? `<w:gridSpan w:val="${span}"/>` : '') +
    borders +
    fill +
    '</w:tcPr>'
  const rPr =
    (style.bold ? '<w:b/>' : '') +
    (style.italic ? '<w:i/>' : '') +
    (style.color ? `<w:color w:val="${hex(style.color)}"/>` : '') +
    (style.fontSize ? `<w:sz w:val="${Math.round(style.fontSize * 2)}"/>` : '')
  const jc = align
    ? `<w:pPr><w:jc w:val="${align === 'right' ? 'right' : align === 'center' ? 'center' : 'left'}"/></w:pPr>`
    : ''
  const run = value
    ? `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ''}<w:t xml:space="preserve">${escapeText(value)}</w:t></w:r>`
    : ''
  return `<w:tc>${tcPr}<w:p>${jc}${run}</w:p></w:tc>`
}

/** `<w:tbl>` for a rendered table; header rows repeat on every page when the table asks for it. */
export function renderedTableToOoxml(
  rendered: RenderedTable,
  table: Pick<DvhTable, 'layout'>,
  locale = 'en-US',
): string {
  const widths = rendered.widths.map((w) => w * PAGE_TEXT_TWIPS)
  const grid = `<w:tblGrid>${widths.map((w) => `<w:gridCol w:w="${Math.round(w)}"/>`).join('')}</w:tblGrid>`
  const rows = rendered.rows.map((row, r) => {
    const header = r < rendered.headerRowCount
    const trPr =
      (header && table.layout.repeatHeader ? '<w:tblHeader/>' : '') +
      (table.layout.keepRowsTogether ? '<w:cantSplit/>' : '')
    let col = 0
    const cells = row.map((cell) => {
      const width = widths.slice(col, col + cell.colSpan).reduce((a, b) => a + b, 0)
      col += cell.colSpan
      const text =
        cell.kind === 'body' || cell.kind === 'total'
          ? formatCellValue(cell.value, cell.numFmt, locale)
          : cell.value === null
            ? ''
            : String(cell.value)
      return cellXml(text, cell.style, width, cell.colSpan, cell.style.align)
    })
    return `<w:tr>${trPr ? `<w:trPr>${trPr}</w:trPr>` : ''}${cells.join('')}</w:tr>`
  })
  return (
    '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr>' +
    grid +
    rows.join('') +
    '</w:tbl>'
  )
}
