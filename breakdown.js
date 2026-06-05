import * as XLSX from 'xlsx'

export const ALL_STATES = [
  'Abia','Adamawa','Akwa Ibom','Anambra','Bauchi','Bayelsa','Benue','Borno',
  'Cross River','Delta','Ebonyi','Edo','Ekiti','Enugu','Gombe','Imo','Jigawa',
  'Kaduna','Kano','Katsina','Kebbi','Kogi','Kwara','Lagos','Nasarawa','Niger',
  'Ogun','Ondo','Osun','Oyo','Plateau','Rivers','Sokoto','Taraba','Yobe',
  'Zamfara','Abuja',
]

export const CHANNELS = ['Retail', 'Horeca', 'Modern Trade']

function canonical(s) {
  const sl = s.trim().toLowerCase()
  for (const st of ALL_STATES) if (st.toLowerCase() === sl) return st
  const aliases = {
    'phc': 'Rivers', 'rivers': 'Rivers', 'fct': 'Abuja',
    'akwa ibom': 'Akwa Ibom', 'akwa_ibom': 'Akwa Ibom',
    'cross river': 'Cross River', 'crossriver': 'Cross River',
  }
  return aliases[sl] || null
}

export function resolveStates(r) {
  const rl = r.trim().toLowerCase()
  if (rl === 'all location' || rl === 'all locations' || rl === 'all location ')
    return [...ALL_STATES]
  if (rl.includes('all locations except edo'))
    return ALL_STATES.filter(s => s !== 'Edo')
  if (rl.includes('other region except') || rl.includes('other regions except')) {
    const exclude = new Set()
    if (rl.includes('lagos')) exclude.add('Lagos')
    if (rl.includes('phc'))   exclude.add('Rivers')
    if (rl.includes('abuja')) exclude.add('Abuja')
    return ALL_STATES.filter(s => !exclude.has(s))
  }
  if (rl.includes('lagos and other regions except phc') ||
      rl.includes('lagos and other region except phc'))
    return ALL_STATES.filter(s => s !== 'Rivers')
  if (r.includes(',')) {
    const parts = r.split(',').map(p => p.trim())
    const resolved = parts.map(canonical).filter(Boolean)
    return resolved.length ? resolved : [r]
  }
  const c = canonical(r)
  return c ? [c] : [r]
}

export function isGeneric(r) {
  const rl = r.trim().toLowerCase()
  return (
    rl === 'all location' || rl === 'all locations' ||
    rl.includes('other region') || rl.includes('other regions') ||
    rl.includes('all locations except') || r.includes(',')
  )
}

export function formatDate(val) {
  if (!val) return ''
  try {
    const d = new Date(val)
    if (isNaN(d)) return String(val)
    return d.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' })
  } catch { return String(val) }
}

export function processWorkbook(arrayBuffer, sheetName) {
  const wb = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array', cellDates: true })

  // Use specified sheet name, fall back to first sheet
  const targetSheet = sheetName && wb.SheetNames.includes(sheetName)
    ? sheetName
    : wb.SheetNames[0]

  const ws = wb.Sheets[targetSheet]
  const rows = XLSX.utils.sheet_to_json(ws, { defval: '' })
  if (!rows.length) throw new Error(`Sheet "${targetSheet}" appears to be empty.`)

  const norm = r => {
    const out = {}
    Object.entries(r).forEach(([k, v]) => { out[k.trim()] = v })
    return out
  }
  const df = rows.map(norm)
  const cols = Object.keys(df[0])

  const stateCol = cols.find(c => c.toLowerCase() === 'state') || 'State'
  const dateCol  = cols.find(c => c.toLowerCase().includes('date')) || 'Date Updated'

  // Explode states + channels
  const expanded = []
  for (const row of df) {
    const raw = String(row[stateCol] || '').trim()
    const states = resolveStates(raw)
    for (const st of states) {
      for (const ch of CHANNELS) {
        expanded.push({ ...row, _Resolved_State: st, _Resolved_Channel: ch, _raw_state: raw })
      }
    }
  }

  // Sort: newest date first, specific-state beats generic on same date
  expanded.sort((a, b) => {
    const skuCmp = String(a['SKU Code'] || '').localeCompare(String(b['SKU Code'] || ''))
    if (skuCmp !== 0) return skuCmp
    const stCmp = a._Resolved_State.localeCompare(b._Resolved_State)
    if (stCmp !== 0) return stCmp
    const chCmp = a._Resolved_Channel.localeCompare(b._Resolved_Channel)
    if (chCmp !== 0) return chCmp
    const dA = new Date(a[dateCol] || 0), dB = new Date(b[dateCol] || 0)
    if (dB - dA !== 0) return dB - dA
    return (isGeneric(a._raw_state) ? 1 : 0) - (isGeneric(b._raw_state) ? 1 : 0)
  })

  // Latest: first occurrence per (SKU Code, state, channel)
  const seen = new Set()
  const latest = []
  for (const row of expanded) {
    const key = `${row['SKU Code']}||${row._Resolved_State}||${row._Resolved_Channel}`
    if (!seen.has(key)) { seen.add(key); latest.push(row) }
  }

  // Build output columns
  const dropCols = new Set(['State', 'Channel', 'State ID', 'Channel ID'])
  const cleanCols = cols.filter(c => !dropCols.has(c) && !c.startsWith('Unnamed'))
  const frontCols = [dateCol, 'Model', 'Manufacturer', 'Category', 'Brand',
                     '_Resolved_Channel', '_Resolved_State', 'SKU Code', 'SKU']
  const restCols = cleanCols.filter(c => !frontCols.includes(c))
  const allOutCols = [...frontCols, ...restCols]

  function toOutputRows(arr) {
    return arr.map(row => {
      const out = {}
      for (const c of allOutCols) {
        const key = c === '_Resolved_Channel' ? 'Channel'
                  : c === '_Resolved_State'   ? 'State'
                  : c
        let val = c === '_Resolved_Channel' ? row._Resolved_Channel
                : c === '_Resolved_State'   ? row._Resolved_State
                : row[c]
        if (val instanceof Date) val = formatDate(val)
        out[key] = val ?? ''
      }
      return out
    })
  }

  return {
    full:       toOutputRows(expanded),
    latest:     toOutputRows(latest),
    sourceRows: rows.length,
    sheetUsed:  targetSheet,
  }
}

export function buildOutputXLSX(fullRows, latestRows, sourceRows, sheetUsed) {
  const wb = XLSX.utils.book_new()

  const summaryData = [
    ['SKU Price Breakdown — Summary'],
    [],
    ['Metric', 'Value'],
    ['Source file', 'Master Price Update vF.xlsx'],
    ['Sheet read', sheetUsed],
    ['Source file rows', sourceRows],
    ['Full Expanded rows', fullRows.length],
    ['Latest Prices rows', latestRows.length],
    ['Channels expanded', 'Retail | Horeca | Modern Trade'],
    ['States expanded', '36 States + FCT Abuja (37 total)'],
    [],
    ['EXPANSION RULES'],
    ['"All Channels"', '→ 3 rows: Retail, Horeca, Modern Trade'],
    ['"All Location / All Locations"', '→ 37 rows (one per state)'],
    ['"Other region except X"', '→ All states MINUS X'],
    ['"Other region except X and PHC"', '→ All states MINUS X and Rivers'],
    ['"State A, State B, ..."', '→ Individual rows per state'],
    [],
    ['LATEST PRICE LOGIC'],
    ['Specific state wins', 'State-specific entry beats All Location on same date'],
    ['Recency rule', 'Most recent date always wins'],
    ['Fallback', 'No specific entry → use most recent All Location entry'],
    [],
    ['SHEETS'],
    ['Full Expanded', 'All historical records broken out by channel & state'],
    ['Latest Prices', 'Current active price per SKU × Channel × State'],
  ]
  const wsSummary = XLSX.utils.aoa_to_sheet(summaryData)
  wsSummary['!cols'] = [{ wch: 46 }, { wch: 58 }]
  XLSX.utils.book_append_sheet(wb, wsSummary, 'Summary')

  const wsFull = XLSX.utils.json_to_sheet(fullRows)
  if (fullRows.length) {
    const ncols = Object.keys(fullRows[0]).length
    wsFull['!cols'] = Array(ncols).fill({ wch: 22 })
    wsFull['!autofilter'] = { ref: `A1:${XLSX.utils.encode_col(ncols - 1)}1` }
  }
  XLSX.utils.book_append_sheet(wb, wsFull, 'Full Expanded')

  const wsLatest = XLSX.utils.json_to_sheet(latestRows)
  if (latestRows.length) {
    const ncols = Object.keys(latestRows[0]).length
    wsLatest['!cols'] = Array(ncols).fill({ wch: 22 })
    wsLatest['!autofilter'] = { ref: `A1:${XLSX.utils.encode_col(ncols - 1)}1` }
  }
  XLSX.utils.book_append_sheet(wb, wsLatest, 'Latest Prices')

  return XLSX.write(wb, { bookType: 'xlsx', type: 'array' })
}
