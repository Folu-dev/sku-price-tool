const XLSX = require('xlsx')

const ALL_STATES = [
  'Abia','Adamawa','Akwa Ibom','Anambra','Bauchi','Bayelsa','Benue','Borno',
  'Cross River','Delta','Ebonyi','Edo','Ekiti','Enugu','Gombe','Imo','Jigawa',
  'Kaduna','Kano','Katsina','Kebbi','Kogi','Kwara','Lagos','Nasarawa','Niger',
  'Ogun','Ondo','Osun','Oyo','Plateau','Rivers','Sokoto','Taraba','Yobe',
  'Zamfara','Abuja',
]
const CHANNELS = ['Retail', 'Horeca', 'Modern Trade']

function canonical(s) {
  const sl = s.trim().toLowerCase()
  for (const st of ALL_STATES) if (st.toLowerCase() === sl) return st
  const aliases = {
    'phc':'Rivers','fct':'Abuja','akwa ibom':'Akwa Ibom',
    'cross river':'Cross River','crossriver':'Cross River',
  }
  return aliases[sl] || null
}

function resolveStates(r) {
  const rl = r.trim().toLowerCase()
  if (rl === 'all location' || rl === 'all locations') return [...ALL_STATES]
  if (rl.includes('all locations except edo')) return ALL_STATES.filter(s => s !== 'Edo')
  if (rl.includes('other region except') || rl.includes('other regions except')) {
    const exclude = new Set()
    if (rl.includes('lagos')) exclude.add('Lagos')
    if (rl.includes('phc'))   exclude.add('Rivers')
    if (rl.includes('abuja')) exclude.add('Abuja')
    return ALL_STATES.filter(s => !exclude.has(s))
  }
  if (rl.includes('lagos and other regions except phc')) return ALL_STATES.filter(s => s !== 'Rivers')
  if (r.includes(',')) {
    const parts = r.split(',').map(p => p.trim())
    const resolved = parts.map(canonical).filter(Boolean)
    return resolved.length ? resolved : [r]
  }
  const c = canonical(r)
  return c ? [c] : [r]
}

function isGeneric(r) {
  const rl = r.trim().toLowerCase()
  return rl === 'all location' || rl === 'all locations' ||
    rl.includes('other region') || rl.includes('all locations except') || r.includes(',')
}

function formatDate(val) {
  if (!val) return ''
  try {
    const d = new Date(val)
    if (isNaN(d)) return String(val)
    return d.toLocaleDateString('en-GB', { day:'2-digit', month:'short', year:'numeric' })
  } catch { return String(val) }
}

function processBuffer(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true })
  const targetSheet = wb.SheetNames.includes('MMP New Structure') ? 'MMP New Structure' : wb.SheetNames[0]
const ws = wb.Sheets[targetSheet]
const rows = XLSX.utils.sheet_to_json(ws, { defval: '', range: 1 })
if (!rows.length) throw new Error(`Sheet "${targetSheet}" appears to be empty. Available sheets: ${wb.SheetNames.join(', ')}`)

  const norm = r => { const o = {}; Object.entries(r).forEach(([k,v]) => { o[k.trim()] = v }); return o }
  const df = rows.map(norm)
  const cols = Object.keys(df[0])

  const skuCol  = cols.find(c => c === 'SKU Code') || cols.find(c => c === 'A') || 'A'
  const stateCol = cols.find(c => c.toLowerCase() === 'state') || 'State'
  const dateCol  = cols.find(c => c.toLowerCase().includes('date')) || 'Date Updated'

  // Explode
  const expanded = []
  for (const row of df) {
    const raw = String(row[stateCol] || '').trim()
    for (const st of resolveStates(raw)) {
      for (const ch of CHANNELS) {
        expanded.push({ ...row, _State: st, _Channel: ch, _raw: raw })
      }
    }
  }

  // Sort: newest first, specific beats generic
  expanded.sort((a, b) => {
    const s = String(a[skuCol]||'').localeCompare(String(b[skuCol]||''))
    if (s) return s
    const t = a._State.localeCompare(b._State)
    if (t) return t
    const u = a._Channel.localeCompare(b._Channel)
    if (u) return u
    const dA = new Date(a[dateCol]||0), dB = new Date(b[dateCol]||0)
    if (dB-dA) return dB-dA
    return (isGeneric(a._raw)?1:0) - (isGeneric(b._raw)?1:0)
  })

  // Latest
  const seen = new Set()
  const latest = []
  for (const row of expanded) {
    const key = `${row[skuCol]}||${row._State}||${row._Channel}`
    if (!seen.has(key)) { seen.add(key); latest.push(row) }
  }

  const dropCols = new Set(['State','Channel','State ID','Channel ID'])
  const cleanCols = cols.filter(c => !dropCols.has(c) && !c.startsWith('Unnamed') && !c.startsWith('__EMPTY'))
  const frontCols = [dateCol,'Model','Manufacturer','Category','Brand','_Channel','_State', skuCol,'SKU']
  const restCols  = cleanCols.filter(c => !frontCols.includes(c))
  const allCols   = [...frontCols, ...restCols]

  function toRows(arr) {
    return arr.map(row => {
      const o = {}
      for (const c of allCols) {
        const key = c === '_Channel' ? 'Channel' : c === '_State' ? 'State' : c === skuCol && skuCol === 'A' ? 'SKU Code' : c
        let val = c === '_Channel' ? row._Channel : c === '_State' ? row._State : row[c]
        if (val instanceof Date) val = formatDate(val)
        o[key] = val ?? ''
      }
      return o
    })
  }

  const fullRows   = toRows(expanded)
  const latestRows = toRows(latest)

  // Build output workbook
  const outWb = XLSX.utils.book_new()

  const summary = [
    ['SKU Price Breakdown — Summary'],
    [],
    ['Metric','Value'],
    ['Source rows', rows.length],
    ['Full Expanded rows', fullRows.length],
    ['Latest Prices rows', latestRows.length],
    ['Channels','Retail | Horeca | Modern Trade'],
    ['States','36 States + FCT Abuja (37 total)'],
    [],
    ['RULES'],
    ['"All Location"','→ 37 rows'],
    ['"Other region except X"','→ All states minus X'],
    ['"State A, State B"','→ Individual rows per state'],
    ['Latest price','Specific state wins; most recent date wins'],
  ]
  XLSX.utils.book_append_sheet(outWb, XLSX.utils.aoa_to_sheet(summary), 'Summary')

  const wsFull = XLSX.utils.json_to_sheet(fullRows)
  if (fullRows.length) {
    const n = Object.keys(fullRows[0]).length
    wsFull['!cols'] = Array(n).fill({ wch: 22 })
    wsFull['!autofilter'] = { ref: `A1:${XLSX.utils.encode_col(n-1)}1` }
  }
  XLSX.utils.book_append_sheet(outWb, wsFull, 'Full Expanded')

  const wsLatest = XLSX.utils.json_to_sheet(latestRows)
  if (latestRows.length) {
    const n = Object.keys(latestRows[0]).length
    wsLatest['!cols'] = Array(n).fill({ wch: 22 })
    wsLatest['!autofilter'] = { ref: `A1:${XLSX.utils.encode_col(n-1)}1` }
  }
  XLSX.utils.book_append_sheet(outWb, wsLatest, 'Latest Prices')

  return {
    buffer: XLSX.write(outWb, { bookType: 'xlsx', type: 'buffer' }),
    stats: { sourceRows: rows.length, full: fullRows.length, latest: latestRows.length }
  }
}

// Parse multipart form data manually (no external deps)
function parseMultipart(body, boundary) {
  const parts = []
  const boundaryBuffer = Buffer.from('--' + boundary)
  let start = body.indexOf(boundaryBuffer) + boundaryBuffer.length + 2
  while (start < body.length) {
    const end = body.indexOf(boundaryBuffer, start)
    if (end === -1) break
    const part = body.slice(start, end - 2)
    const headerEnd = part.indexOf(Buffer.from('\r\n\r\n'))
    if (headerEnd === -1) { start = end + boundaryBuffer.length + 2; continue }
    const headers = part.slice(0, headerEnd).toString()
    const data = part.slice(headerEnd + 4)
    parts.push({ headers, data })
    start = end + boundaryBuffer.length + 2
  }
  return parts
}

module.exports = async (req, res) => {
  // CORS headers
  res.setHeader('Access-Control-Allow-Origin', '*')
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS')
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type')

  if (req.method === 'OPTIONS') { res.status(200).end(); return }
  if (req.method !== 'POST') { res.status(405).json({ error: 'Method not allowed' }); return }

  try {
    const chunks = []
    for await (const chunk of req) chunks.push(chunk)
    const body = Buffer.concat(chunks)

    const contentType = req.headers['content-type'] || ''
    const boundaryMatch = contentType.match(/boundary=(.+)/)
    if (!boundaryMatch) { res.status(400).json({ error: 'No boundary found' }); return }

    const parts = parseMultipart(body, boundaryMatch[1])
    const filePart = parts.find(p => p.headers.includes('filename'))
    if (!filePart) { res.status(400).json({ error: 'No file found in upload' }); return }

    const { buffer, stats } = processBuffer(filePart.data)

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    res.setHeader('Content-Disposition', 'attachment; filename="SKU_Price_Breakdown.xlsx"')
    res.setHeader('X-Stats', JSON.stringify(stats))
    res.status(200).send(buffer)
  } catch (e) {
    console.error(e)
    res.status(500).json({ error: e.message || 'Processing failed' })
  }
}
