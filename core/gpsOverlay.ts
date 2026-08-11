// Parse the GPS-Map-Camera style overlay text stamped on a field photo into
// structured fields. The overlay looks like:
//   N 17°52'17.42196" LAT
//   E 77°59'25.8558" LON
//   Altitude: 1696 ft a.s.l
//   8/7/26 3:31 PM
//   Location provider: Fused
//   No street
//   Keroor 502270
//   Telangana
//   India
// The input is the set of OCR lines (from one or more preprocessing passes). OCR
// on white-on-photo text is imperfect, so `confidence` is 'high' only when both
// coordinates were read with their full degree/minute/second delimiters AND a
// decimal — the readings that parse exactly — and 'low' otherwise (a digit-run
// guess or a missing decimal), so the caller can flag those rows for a human check.

export interface GpsOverlay {
  lat?: number
  lon?: number
  latDMS?: string
  lonDMS?: string
  altitude?: string
  /** The full street address block from the stamp (door no, landmark, locality,
   * state, pincode, India) — the whole thing, not just the locality/pincode. */
  address?: string
  place?: string
  pincode?: string
  state?: string
  dateTime?: string
  confidence: 'high' | 'low'
  /** The raw OCR lines joined, so the user can verify a flagged row. */
  raw: string
}

interface Coord {
  dec: number
  dms: string
  fracLen: number
  strong: boolean
}

function toDecimal(deg: number, min: number, sec: number, negative: boolean): number {
  const d = deg + min / 60 + sec / 3600
  return negative ? -d : d
}

// Pull the best coordinate for one axis out of the OCR lines. A "strong" match
// has the °/'/" structure and a decimal (parses exactly); a "weak" match is a
// bare digit run split by the fixed DD MM SS field widths (a guess).
function findCoord(lines: string[], kind: 'lat' | 'lon'): Coord | undefined {
  const tag = kind === 'lat' ? /LAT|IAT|1AT/i : /LON|L0N|L1N|LONG/i
  const hemiRe = kind === 'lat' ? /\b([NS])\b|(?:^|\s)([NS])\s*\d/ : /\b([EW])\b|(?:^|\s)([EW])\s*\d/
  let best: Coord | undefined
  const consider = (c: Coord) => {
    if (!best) best = c
    else if (c.strong !== best.strong) best = c.strong ? c : best
    else if (c.fracLen > best.fracLen) best = c
  }
  for (const line of lines) {
    if (!tag.test(line)) continue
    const hm = line.match(hemiRe)
    const hemi = hm ? (hm[1] || hm[2] || '').toUpperCase() : ''
    const negative = hemi === 'S' || hemi === 'W'
    // Strong: DD ° MM ' SS . fff  (delimiters + decimal)
    const strong = line.match(/(\d{1,3})\s*[°ºo*]\s*(\d{1,2})\s*['’:]\s*(\d{1,2})\s*[.,]\s*(\d+)/)
    if (strong) {
      const sec = Number(`${strong[3]}.${strong[4]}`)
      if (Number(strong[2]) <= 59 && sec < 60) {
        consider({
          dec: toDecimal(Number(strong[1]), Number(strong[2]), sec, negative),
          dms: `${strong[1]}°${strong[2]}'${strong[3]}.${strong[4]}"`,
          fracLen: strong[4].length,
          strong: true
        })
        continue
      }
    }
    // Weak: strip to digits (+ one decimal) and split by fixed field widths.
    const cleaned = line
      .replace(/LAT|IAT|1AT|LON|L0N|L1N|LONG/gi, ' ')
      .replace(/[Oo]/g, '0')
      .replace(/[lI|]/g, '1')
    const numMatch = cleaned.match(/(\d[\d\s]*(?:[.,]\d+)?)/)
    if (!numMatch) continue
    const digitsAll = numMatch[1].replace(/[\s]/g, '').replace(',', '.')
    const dot = digitsAll.indexOf('.')
    const intPart = dot === -1 ? digitsAll : digitsAll.slice(0, dot)
    const frac = dot === -1 ? '' : digitsAll.slice(dot + 1)
    if (intPart.length < 6) continue
    // 2/2/2 field widths (degrees are 2 digits for this region's lat & long).
    const deg = Number(intPart.slice(0, 2))
    const min = Number(intPart.slice(2, 4))
    const sec = Number(`${intPart.slice(4, 6)}${frac ? '.' + frac : ''}`)
    if (min > 59 || sec >= 60) continue
    consider({
      dec: toDecimal(deg, min, sec, negative),
      dms: `${intPart.slice(0, 2)}°${intPart.slice(2, 4)}'${intPart.slice(4, 6)}${frac ? '.' + frac : ''}"`,
      fracLen: frac.length,
      strong: false
    })
  }
  return best
}

const STATE_RE =
  /telangana|andhra\s*pradesh|karnataka|maharashtra|tamil\s*nadu|kerala|odisha|gujarat|rajasthan|madhya\s*pradesh|uttar\s*pradesh|bihar|west\s*bengal|punjab|haryana|chhattisgarh|jharkhand|assam|goa/i

// Reconstruct the full street address from a GPS-Map-Camera stamp. The stamp
// lists it as, e.g.:
//   Ananthasagar, Telangana, India                <- short locality header
//   Sangareddy, 4-91/2, Near At Old Csi Church,   <- district / door no / landmark
//   Ananthasagar, Telangana 502306, India         <- locality, state, PINCODE, India
// We anchor on the pincode line (the address "tail"), then walk upward gathering
// the street/door/landmark lines, stopping at the short "locality, State, India"
// header (which has "India" but no pincode) and at any coord/date/label line — so
// the result is the whole address, not just the locality + pincode.
function findFullAddress(lines: string[]): string | undefined {
  const skip =
    /\blat\b|\blong\b|latitude|longitude|altitude|\bGMT\b|\bAM\b|\bPM\b|google|GPS\s*Map|map\s*camera|location\s*provider/i
  let pi = -1
  for (let i = 0; i < lines.length; i++) {
    if (/\b\d{6}\b/.test(lines[i]) && (/india/i.test(lines[i]) || STATE_RE.test(lines[i]) || lines[i].includes(','))) {
      pi = i
      break
    }
  }
  if (pi === -1) return undefined
  const parts: string[] = []
  for (let i = pi; i >= 0 && pi - i < 6; i--) {
    const L = lines[i].trim()
    if (!L) continue
    if (i !== pi) {
      // The short "…, State, India" header (India, no pincode) ends the address.
      if (/india/i.test(L) && !/\d{6}/.test(L)) break
      if (skip.test(L)) break
      if (!/[A-Za-z]/.test(L)) break // a separator (e.g. the flag) — stop
    }
    parts.unshift(L)
  }
  const addr = parts
    .join(' ')
    .replace(/\s*,\s*/g, ', ')
    .replace(/(?:,\s*)+,/g, ',')
    .replace(/\s+/g, ' ')
    .replace(/[\s,]+$/, '')
    .trim()
  return addr || undefined
}

// The other common overlay is plain decimal degrees — "Lat/Long: Lat 17.611674
// Long 78.162789" (bottom-corner GPS-camera stamp). It's faint, so the caller
// OCRs several contrast-enhanced variants and passes ALL their lines here; we
// tally the decimal values seen for each axis and take the consensus (the value
// two or more independent passes agree on). Two passes rarely make the identical
// digit error, so a value with ≥2 votes is trusted; a lone read (which may have
// lost a digit or the decimal point) has count 1 and the row is flagged instead.
function findDecimalCoords(lines: string[]): {
  lat?: number
  lon?: number
  latVotes: number
  lonVotes: number
  rebuilt: boolean
} {
  const latV = new Map<number, number>()
  const lonV = new Map<number, number>()
  const anyV = new Map<number, number>()
  // A separate pool of "rebuilt" candidates — values recovered from a faint
  // stamp where OCR dropped the decimal point or the value wrapped onto the next
  // line. Only consulted when the normal decimal read fails for that axis, and
  // when used it forces the row to 'low' confidence (a human should verify).
  const latRebuilt = new Map<number, number>()
  const lonRebuilt = new Map<number, number>()
  const bump = (m: Map<number, number>, v: number) => m.set(v, (m.get(v) ?? 0) + 1)
  // Insert the decimal into a decimal-LESS coordinate run: every India lat/long
  // has a 2-digit integer part, so "78164114" → 78.164114. undefined if implausible.
  const rebuild = (digits: string): number | undefined => {
    const d = digits.replace(/\D/g, '')
    if (d.length < 6 || d.length > 10) return undefined
    const v = parseFloat(`${d.slice(0, 2)}.${d.slice(2)}`)
    return Number.isFinite(v) ? v : undefined
  }
  // "Long" gets OCR'd as L0ng / L0n9 / Long; match all. (o→0 already applied.)
  const LON = 'L[o0]n[g9]'
  for (let i = 0; i < lines.length; i++) {
    const s = lines[i].replace(/[Oo]/g, '0')
    const lm = s.match(/Lat(?:itude)?\s*[:.\-/]*\s*(-?\d{1,2}\.\d{3,})/i)
    if (lm) {
      const v = parseFloat(lm[1])
      if (Math.abs(v) <= 90) bump(latV, v)
    }
    const om = s.match(new RegExp(`${LON}(?:itude)?\\s*[:.\\-/]*\\s*(-?\\d{1,3}\\.\\d{3,})`, 'i'))
    if (om) {
      const v = parseFloat(om[1])
      if (Math.abs(v) <= 180) bump(lonV, v)
    }
    for (const mm of s.matchAll(/(-?\d{1,3}\.\d{3,})/g)) bump(anyV, parseFloat(mm[1]))

    // ---- Faint-stamp recovery (context-gated so pincodes/dates/plus-codes
    // ---- can never be mistaken for coordinates) ----
    // (a) same line, decimal dropped: "Lat 17611379" / "Long 78164114".
    const latNoDot = s.match(/Lat(?:itude)?\s*[:.\-/]*\s*(\d{6,10})(?![\d.])/i)
    if (latNoDot) {
      const v = rebuild(latNoDot[1])
      if (v != null && v <= 90) bump(latRebuilt, v)
    }
    const lonNoDot = s.match(new RegExp(`${LON}(?:itude)?\\s*[:.\\-/]*\\s*(\\d{6,10})(?![\\d.])`, 'i'))
    if (lonNoDot) {
      const v = rebuild(lonNoDot[1])
      if (v != null && v <= 180) bump(lonRebuilt, v)
    }
    // (b) wrapped onto the next line: this line ENDS with a bare Lat/Long label,
    // the value is the start of the next line (with or without its decimal).
    const endsLon = new RegExp(`${LON}\\s*[:.\\-]?\\s*$`, 'i').test(s)
    const endsLat = /Lat(?:itude)?\s*[:.\-]?\s*$/i.test(s) && !endsLon
    if (endsLon || endsLat) {
      const next = (lines[i + 1] ?? '').replace(/[Oo]/g, '0').trim()
      const dot = next.match(/^(-?\d{1,3}\.\d{3,})/)
      const run = next.match(/^(\d{6,10})(?![\d.])/)
      const v = dot ? parseFloat(dot[1]) : run ? rebuild(run[1]) : undefined
      if (v != null) {
        if (endsLon && v <= 180) bump(lonRebuilt, v)
        else if (endsLat && v <= 90) bump(latRebuilt, v)
      }
    }
  }
  const mode = (m: Map<number, number>): { v: number; c: number } | undefined => {
    let best: { v: number; c: number } | undefined
    for (const [v, c] of m) if (!best || c > best.c) best = { v, c }
    return best
  }
  let lat = mode(latV)
  let lon = mode(lonV)
  // If an axis had no labelled read, fall back to a standalone decimal in range
  // (lat ≤ 90, lon ≤ 180) that isn't the other axis's value.
  if (!lon) {
    const cand = new Map<number, number>()
    for (const [v, c] of anyV) if ((!lat || v !== lat.v) && Math.abs(v) <= 180) cand.set(v, c)
    lon = mode(cand)
  }
  if (!lat) {
    const cand = new Map<number, number>()
    for (const [v, c] of anyV) if ((!lon || v !== lon.v) && Math.abs(v) <= 90) cand.set(v, c)
    lat = mode(cand)
  }
  // Last resort: a rebuilt (decimal-inserted / de-wrapped) value for whichever
  // axis is still missing. Marks the whole read 'low' so the user verifies it.
  let rebuilt = false
  if (!lon) {
    const cand = new Map<number, number>()
    for (const [v, c] of lonRebuilt) if (!lat || v !== lat.v) cand.set(v, c)
    lon = mode(cand)
    if (lon) rebuilt = true
  }
  if (!lat) {
    const cand = new Map<number, number>()
    for (const [v, c] of latRebuilt) if (!lon || v !== lon.v) cand.set(v, c)
    lat = mode(cand)
    if (lat) rebuilt = true
  }
  return { lat: lat?.v, lon: lon?.v, latVotes: lat?.c ?? 0, lonVotes: lon?.c ?? 0, rebuilt }
}

export function parseGpsOverlay(lines: string[]): GpsOverlay {
  const clean = lines.map((l) => l.trim()).filter(Boolean)
  const raw = clean.join(' | ')
  const lat = findCoord(clean, 'lat')
  const lon = findCoord(clean, 'lon')
  const dec = findDecimalCoords(clean)

  let altitude: string | undefined
  let dateTime: string | undefined
  let place: string | undefined
  let pincode: string | undefined
  let state: string | undefined

  for (const line of clean) {
    const alt = line.match(/Alt[a-z]*\s*[:.]?\s*([\d.,]+)\s*(ft|m)\b/i)
    if (alt && !altitude) altitude = `${alt[1].replace(',', '')} ${alt[2].toLowerCase()}`

    const dt = line.match(/(\d{1,2}\/\d{1,2}\/\d{2,4})\s*(\d{1,2}[:.]\d{2}\s*[AP]\.?M\.?)?/i)
    if (dt && !dateTime) dateTime = [dt[1], dt[2]].filter(Boolean).join(' ').replace(/\s+/g, ' ')

    const pin = line.match(/([A-Za-z][A-Za-z .]*?)\s*(\d{6})\b/)
    if (pin && !pincode) {
      place = pin[1].trim() || undefined
      pincode = pin[2]
    } else {
      const bare = line.match(/\b(\d{6})\b/)
      if (bare && !pincode) pincode = bare[1]
    }

    const st = line.match(STATE_RE)
    if (st && !state) state = st[0].replace(/\s+/g, ' ')

    // Plus-code + place line, e.g. "J5674+M4H,Erdanoor" → place "Erdanoor".
    const plus = line.match(/[A-Z0-9]{4,7}\+[A-Z0-9]{2,4}\s*[,.\s]\s*([A-Za-z][A-Za-z ]{2,})/)
    if (plus && !place) place = plus[1].trim()

    // "Time: 08.08.2026 13:31" (decimal-format stamp). OCR mangles the dots, so
    // capture it loosely as a fallback date/time.
    const tm = line.match(/Time\s*[:.]?\s*([\d.]{6,}\s*\d{1,2}[:.]\d{2})/i)
    if (tm && !dateTime) dateTime = tm[1].replace(/\s+/g, ' ')
  }

  // Prefer a consensus decimal read (≥2 passes agree = exact), then a strong DMS
  // read (exact); fall back to whatever partial value exists, marked low so the
  // row gets verified.
  // A rebuilt value (decimal reconstructed / de-wrapped from a faint stamp) can
  // never be 'high' — the user must verify it.
  const decHigh =
    dec.lat != null && dec.lon != null && dec.latVotes >= 2 && dec.lonVotes >= 2 && !dec.rebuilt
  const decAny = dec.lat != null && dec.lon != null
  const dmsStrong = !!(lat?.strong && lon?.strong)
  let outLat: number | undefined
  let outLon: number | undefined
  let confidence: 'high' | 'low'
  if (decHigh) {
    outLat = dec.lat
    outLon = dec.lon
    confidence = 'high'
  } else if (dmsStrong) {
    outLat = lat!.dec
    outLon = lon!.dec
    confidence = 'high'
  } else if (decAny) {
    outLat = dec.lat
    outLon = dec.lon
    confidence = 'low'
  } else {
    outLat = dec.lat ?? lat?.dec
    outLon = dec.lon ?? lon?.dec
    confidence = 'low'
  }

  return {
    lat: outLat,
    lon: outLon,
    latDMS: lat?.dms,
    lonDMS: lon?.dms,
    altitude,
    address: findFullAddress(clean),
    place,
    pincode,
    state,
    dateTime,
    confidence,
    raw
  }
}
