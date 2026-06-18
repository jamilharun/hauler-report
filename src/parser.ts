import ExcelJS from 'exceljs'
import type { TripRow, DriverGroup } from './types'

const MONTHS = ['JANUARY','FEBRUARY','MARCH','APRIL','MAY','JUNE',
                'JULY','AUGUST','SEPTEMBER','OCTOBER','NOVEMBER','DECEMBER']

function findSummarySheet(wb: ExcelJS.Workbook): ExcelJS.Worksheet | undefined {
  return wb.worksheets.find(ws => {
    const upper = ws.name.trim().toUpperCase()
    return MONTHS.some(m => upper.includes(m))
  })
}

function cellText(cell: ExcelJS.Cell): string {
  const v = cell.value
  if (v === null || v === undefined) return ''
  if (typeof v === 'object' && 'richText' in v) {
    return (v as ExcelJS.CellRichTextValue).richText.map(r => r.text).join('')
  }
  return String(v).trim()
}


function excelSerialFromDate(d: Date): number {
  return Math.round((d.getTime() / 86400000) + 25569)
}

export function serialToDate(serial: number): Date {
  return new Date(Math.round((serial - 25569) * 86400 * 1000))
}

export function excelDateToLabel(serial: number): string {
  if (!serial) return ''
  const d = serialToDate(serial)
  return `${(d.getUTCMonth() + 1).toString().padStart(2, '0')}/${d.getUTCDate().toString().padStart(2, '0')}/${d.getUTCFullYear()}`
}

export function monthYearFromSerial(serial: number): string {
  const d = serialToDate(serial)
  return `${MONTHS[d.getUTCMonth()]} ${d.getUTCFullYear()}`
}

export function parseWorkbook(file: File): Promise<DriverGroup[]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = async (e) => {
      try {
        const buffer = e.target!.result as ArrayBuffer
        const wb = new ExcelJS.Workbook()
        await wb.xlsx.load(buffer)

        const ws = findSummarySheet(wb)
        if (!ws) {
          reject(new Error('No summary sheet found. Sheet name must contain the month (e.g. "SEPTEMBER 2025").'))
          return
        }

        // Find header row
        let headerRowNum = -1
        let headers: string[] = []
        ws.eachRow((row, rowNum) => {
          if (headerRowNum !== -1) return
          const vals = (row.values as ExcelJS.CellValue[]).slice(1).map(v => String(v ?? '').trim().toUpperCase())
          if (vals.includes('CLIENT') && vals.includes('DRIVER') && vals.includes('MW')) {
            headerRowNum = rowNum
            headers = vals
          }
        })

        if (headerRowNum === -1) {
          reject(new Error('Could not find header row with CLIENT, DRIVER, MW columns.'))
          return
        }

        const col = (name: string) => headers.indexOf(name)
        const clientIdx  = col('CLIENT')
        const dateIdx    = headers.findIndex(h => h === 'DATE')
        const plateIdx   = col('PLATE#')
        const truckIdx   = col('TRUCK #')
        const driverIdx  = col('DRIVER')
        const mwIdx      = col('MW')
        const sourceIdx  = col('SOURCE')
        const destIdx    = col('DESTINATION')
        const wbIdx      = col('WB #')
        const stIdx      = col('S/T #')
        const tonnageIdx = col('TONNAGE')

        const trips: TripRow[] = []

        ws.eachRow((row, rowNum) => {
          if (rowNum <= headerRowNum) return
          const vals = row.values as ExcelJS.CellValue[]
          const cells = (idx: number) => row.getCell(idx + 1)

          const client = String(vals[clientIdx + 1] ?? '').trim()
          const driver = String(vals[driverIdx + 1] ?? '').trim()
          if (!client || !driver) return

          const mw      = Number(vals[mwIdx + 1]) || 0
          const tonnage = Number(vals[tonnageIdx + 1]) || 0

          const rawDate = cells(dateIdx).value
          const dateSerial = rawDate instanceof Date
            ? excelSerialFromDate(rawDate)
            : Number(rawDate) || 0

          trips.push({
            client,
            date:        dateSerial,
            plate:       String(vals[plateIdx + 1] ?? '').trim(),
            truck:       String(vals[truckIdx + 1] ?? '').trim(),
            driver,
            mw,
            deliveryDate: dateSerial,
            source:      String(vals[sourceIdx + 1] ?? '').trim(),
            destination: String(vals[destIdx + 1] ?? '').trim(),
            wb:          cellText(cells(wbIdx)),
            st:          cellText(cells(stIdx)),
            tonnage,
            variance:    parseFloat((tonnage - mw).toFixed(4)),
          })
        })

        const map = new Map<string, DriverGroup>()
        for (const trip of trips) {
          const key = `${trip.driver}|${trip.truck}`
          if (!map.has(key)) {
            map.set(key, { driver: trip.driver, plate: trip.plate, truck: String(trip.truck), trips: [], totalVariance: 0 })
          }
          const group = map.get(key)!
          group.trips.push(trip)
          group.totalVariance = parseFloat((group.totalVariance + trip.variance).toFixed(4))
        }

        resolve(Array.from(map.values()))
      } catch (err) {
        reject(err)
      }
    }
    reader.onerror = () => reject(new Error('Failed to read file.'))
    reader.readAsArrayBuffer(file)
  })
}
