import ExcelJS from 'exceljs'
import type { DriverGroup } from './types'
import { excelDateToLabel, monthYearFromSerial } from './parser'

const RED = 'FFC62828'
const HEADER_BG = 'FF1565C0'


function styleHeader(row: ExcelJS.Row) {
  row.eachCell(cell => {
    cell.font = { bold: true, color: { argb: 'FFFFFFFF' } }
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEADER_BG } }
    cell.alignment = { horizontal: 'center' }
    cell.border = {
      bottom: { style: 'thin', color: { argb: 'FF90CAF9' } }
    }
  })
}

function styleTotalRow(row: ExcelJS.Row) {
  row.eachCell(cell => {
    cell.font = { bold: true }
  })
}

function setVarianceCell(cell: ExcelJS.Cell, value: number) {
  cell.value = value
  cell.font = { color: { argb: value < 0 ? RED : 'FF000000' }, bold: false }
  cell.numFmt = '0.000'
}

export function getMonthLabel(groups: DriverGroup[]): string {
  for (const g of groups) {
    for (const t of g.trips) {
      if (t.date) return monthYearFromSerial(t.date)
    }
  }
  return 'REPORT'
}

export async function generateReport(groups: DriverGroup[], monthLabel: string): Promise<void> {
  const wb = new ExcelJS.Workbook()
  const usedNames = new Set<string>()

  for (const group of groups) {
    const surname = group.driver.split('.').pop()?.trim().split(' ').pop()?.toUpperCase()
      ?? group.driver.toUpperCase()
    let base = `${surname} ${group.truck}`.slice(0, 31)
    let sheetName = base
    let counter = 2
    while (usedNames.has(sheetName)) {
      sheetName = `${base.slice(0, 28)} ${counter}`.slice(0, 31)
      counter++
    }
    usedNames.add(sheetName)

    const ws = wb.addWorksheet(sheetName)

    // Title rows
    ws.addRow([`${group.driver.toUpperCase()} (PLATE#: ${group.plate} / TRUCK#: ${group.truck})`])
    ws.addRow([`MOLASSES TRIP REPORT FOR THE MONTH OF ${monthLabel}`])
    ws.addRow([])

    // Header
    const headerRow = ws.addRow(['SOURCE', 'DESTINATION', 'CLIENT', 'LOADING DATE', 'WAYBILL#', 'MW', 'OUTTURN', 'VARIANCE'])
    styleHeader(headerRow)

    // Data rows
    for (const trip of group.trips) {
      const row = ws.addRow([
        trip.source,
        trip.destination,
        trip.client,
        excelDateToLabel(trip.date),
        trip.wb,
        trip.mw,
        trip.tonnage,
      ])
      setVarianceCell(row.getCell(8), trip.variance)
    }

    // Total row
    const totalRow = ws.addRow(['', '', '', '', '', '', 'TOTAL VARIANCE', ''])
    setVarianceCell(totalRow.getCell(8), group.totalVariance)
    styleTotalRow(totalRow)

    ws.addRow([])
    ws.addRow(['Prepared By:'])

    // Column widths
    ws.columns = [
      { width: 16 }, // SOURCE
      { width: 16 }, // DESTINATION
      { width: 20 }, // CLIENT
      { width: 14 }, // LOADING DATE
      { width: 22 }, // WAYBILL#
      { width: 10 }, // MW
      { width: 10 }, // OUTTURN
      { width: 12 }, // VARIANCE
    ]
  }

  // Ranking sheet
  const ranked = [...groups].sort((a, b) => a.totalVariance - b.totalVariance)
  const rankWs = wb.addWorksheet('RANKING')

  rankWs.addRow([`DRIVER EFFICIENCY RANKING — ${monthLabel}`])
  rankWs.addRow([])
  const rankHeader = rankWs.addRow(['RANK', 'DRIVER', 'PLATE#', 'TRUCK#', 'TOTAL VARIANCE'])
  styleHeader(rankHeader)

  ranked.forEach((g, i) => {
    const row = rankWs.addRow([i + 1, g.driver, g.plate, g.truck, ''])
    setVarianceCell(row.getCell(5), g.totalVariance)
  })

  rankWs.columns = [
    { width: 8 },
    { width: 24 },
    { width: 18 },
    { width: 10 },
    { width: 16 },
  ]

  // Trigger download
  const buffer = await wb.xlsx.writeBuffer()
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = `Driver Report - ${monthLabel}.xlsx`
  a.click()
  URL.revokeObjectURL(url)
}
