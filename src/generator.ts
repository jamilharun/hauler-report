import ExcelJS from 'exceljs'
import type { DriverGroup } from './types'
import { excelDateToLabel, monthYearFromSerial } from './parser'

const RED = 'FFC62828'


function styleHeader(row: ExcelJS.Row) {
  row.eachCell(cell => {
    cell.font = { bold: true }
    cell.alignment = { horizontal: 'center' }
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF000000' } },
      bottom: { style: 'thin', color: { argb: 'FF000000' } },
    }
  })
}

function styleTitleRow(row: ExcelJS.Row, colCount: number) {
  row.worksheet.mergeCells(row.number, 1, row.number, colCount)
  const cell = row.getCell(1)
  cell.font = { name: 'Calibri', bold: true, size: 16 }
  cell.alignment = { horizontal: 'center' }
}

function styleTotalRow(row: ExcelJS.Row) {
  row.eachCell(cell => {
    if (cell.value === null || cell.value === undefined || cell.value === '') return
    cell.font = { ...cell.font, bold: true }
    if (cell.value !== 'TOTAL')
      cell.border = {
        top: { style: 'thin', color: { argb: 'FF000000' } },
        bottom: { style: 'double', color: { argb: 'FF000000' } },
      }
  })
}

function autoFitColumns(ws: ExcelJS.Worksheet, count: number, startRow = 1) {
  const center = { alignment: { horizontal: 'center' as const } }
  const widths = Array(count).fill(8)
  ws.eachRow((row, rowNum) => {
    if (rowNum < startRow) return
    row.eachCell({ includeEmpty: false }, (cell, colIdx) => {
      if (colIdx <= count) {
        const len = cell.value?.toString().length ?? 0
        if (len > widths[colIdx - 1]) widths[colIdx - 1] = len
      }
    })
  })
  ws.columns = widths.map(w => ({ width: w + 2, style: center }))
}

function setVarianceCell(cell: ExcelJS.Cell, value: number) {
  cell.value = value
  cell.font = { color: { argb: value < 0 ? RED : 'FF000000' }, bold: false }
  cell.numFmt = '0.000;(0.000)'
  cell.alignment = { horizontal: 'center' }
}

export function getMonthLabel(groups: DriverGroup[]): string {
  for (const g of groups) {
    for (const t of g.trips) {
      if (t.date) return monthYearFromSerial(t.date)
    }
  }
  return 'REPORT'
}

export async function generateReport(groups: DriverGroup[], monthLabel: string, preparedBy = '', certifiedBy = ''): Promise<void> {
  const wb = new ExcelJS.Workbook()
  const usedNames = new Set<string>()
  const sortedGroups = [...groups].sort((a, b) => a.driver.localeCompare(b.driver))

  for (const group of sortedGroups) {
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
    styleTitleRow(ws.addRow([`${group.driver.toUpperCase()} (PLATE#: ${group.plate} / TRUCK#: ${group.truck})`]), 9)
    styleTitleRow(ws.addRow([`MOLASSES TRIP REPORT FOR THE MONTH OF ${monthLabel}`]), 9)
    ws.addRow([])

    // Header
    const headerRow = ws.addRow(['SOURCE', 'DESTINATION', 'CLIENT', 'LOADING DATE', 'WAYBILL#', 'MW', 'OUTTURN', 'OVERRAGE (SHORTAGE)', '%'])
    styleHeader(headerRow)

    // Data rows
    const sortedTrips = [...group.trips].sort((a, b) => (a.date ?? 0) - (b.date ?? 0))
    for (const trip of sortedTrips) {
      const row = ws.addRow([
        trip.source,
        trip.destination,
        trip.client,
        excelDateToLabel(trip.date),
        trip.wb,
        trip.mw,
        trip.tonnage,
      ])
      for (let i = 1; i <= 5; i++)
        row.getCell(i).alignment = { horizontal: 'center' }
      row.getCell(6).numFmt = '0.000'
      row.getCell(6).alignment = { horizontal: 'center' }
      row.getCell(7).numFmt = '0.000'
      row.getCell(7).alignment = { horizontal: 'center' }
      setVarianceCell(row.getCell(8), trip.variance)
      const pct = trip.mw ? (trip.variance / trip.mw) * 100 : 0
      row.getCell(9).value = pct
      row.getCell(9).numFmt = '0.000"%";(0.000"%")'
      row.getCell(9).font = { color: { argb: pct < 0 ? RED : 'FF000000' } }
      row.getCell(9).alignment = { horizontal: 'center' }
    }

    // Total row
    const totalMW = group.trips.reduce((s, t) => s + (t.mw ?? 0), 0)
    const totalOutturn = group.trips.reduce((s, t) => s + (t.tonnage ?? 0), 0)
    const totalPct = group.trips.reduce((s, t) => s + (t.mw ? (t.variance / t.mw) * 100 : 0), 0)

    const totalRow = ws.addRow(['', '', '', '', 'TOTAL', '', '', '', ''])
    totalRow.getCell(5).alignment = { horizontal: 'center' }
    totalRow.getCell(6).value = totalMW
    totalRow.getCell(6).numFmt = '0.000'
    totalRow.getCell(6).alignment = { horizontal: 'center' }
    totalRow.getCell(7).value = totalOutturn
    totalRow.getCell(7).numFmt = '0.000'
    totalRow.getCell(7).alignment = { horizontal: 'center' }
    setVarianceCell(totalRow.getCell(8), group.totalVariance)
    totalRow.getCell(9).value = totalPct
    totalRow.getCell(9).numFmt = '0.000"%";(0.000"%")'
    totalRow.getCell(9).font = { color: { argb: totalPct < 0 ? RED : 'FF000000' } }
    totalRow.getCell(9).alignment = { horizontal: 'center' }
    styleTotalRow(totalRow)

    ws.addRow([])
    ws.addRow(['Prepared By:', '', 'Certified Correct:'])
    ws.addRow([])
    ws.addRow([preparedBy, '', certifiedBy])

    autoFitColumns(ws, 9, 4)
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

  autoFitColumns(rankWs, 5, 3)

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
