export interface TripRow {
  client: string
  date: number
  plate: string
  truck: string | number
  driver: string
  mw: number
  deliveryDate: number
  source: string
  destination: string
  wb: string | number
  st: string | number
  tonnage: number
  variance: number
}

export interface DriverGroup {
  driver: string
  plate: string
  truck: string
  trips: TripRow[]
  totalVariance: number
}

export interface DriverRank {
  rank: number
  driver: string
  plate: string
  truck: string
  totalVariance: number
}
