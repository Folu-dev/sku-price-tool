const XLSX = require('xlsx')
const https = require('https')

const TENANT_ID      = process.env.AZURE_TENANT_ID
const CLIENT_ID      = process.env.AZURE_CLIENT_ID
const CLIENT_SECRET  = process.env.AZURE_CLIENT_SECRET
const SHAREPOINT_HOST = 'omnibiz0.sharepoint.com'
const SITE_NAME      = 'FinancialPlaninngAnalysis'
const FILE_NAME      = 'Master Price Update vF.xlsx'
const SHEET_NAME     = 'MMP New Structure'

function httpsPost(url, data, headers) {
  return new Promise((resolve, reject) => {
    const u = new URL(url)
    const body = Buffer.from(data)
    const options = {
      hostname: u.hostname, path: u.pathname + u.search,
      method: 'POST',
      headers: { ...headers, 'Content-Length': body.length }
    }
    const req = https.request(options, res => {
      const chunks = []
      res.on('data', c => chunks.push(c))
      res.on('end', () => resolve({ status: res.statusCode, body: Buffer.concat(chunks).toString() }))
    })
    req.on('error', reject)
    req.write(body)
    req.end()
  })
}

function httpsGet(url, headers) {
  return new Promise((resolve, reject) => {
    const u = new URL(url)
    const options = { hostname: u.hostname, path: u.pathname + u.search, method: 'GET', headers }
    const req = https.request(options, res => {
      const chunks = []
      res.on('data', c => chunks.push(c))
      res.on('end', () => resolve({ status: res.statusCode, body: Buffer.concat(chunks), headers: res.headers }))
    })
    req.on('error', reject)
    req.end()
  })
}

async function getToken() {
  const body = new URLSearchParams({
    grant_type: 'client_credentials',
    client_id: CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope: 'https://graph.microsoft.com/.default'
  }).toString()
  const res = await httpsPost(
    `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    body,
    { 'Content-Type': 'application/x-www-form-urlencoded' }
  )
  const data = JSON.parse(res.body)
  if (!data.access_token) throw new Error('Auth failed: ' + (data.error_description || res.body))
  return data.access_token
}

async function graphGet(url, token) {
  const res = await httpsGet(url, { Authorization: `Bearer ${token}` })
  if (res.status >= 400) throw new Error(`Graph API error ${res.status}: ${res.body.toString().slice(0,200)}`)
  return res
}

const ALL_STATES = [
  'Abia','Adamawa','Akwa Ibom','Anambra','Bauchi','Bayelsa','Benue','Borno',
  'Cross River','Delta','Ebonyi','Edo','Ekiti','Enugu','Gombe','Imo','Jigawa',
  'Kaduna','Kano','Katsina','Kebbi','Kogi','Kwara','Lagos','Nasarawa','Niger',
  'Ogun','Ondo','Osun','Oyo','Plateau','Rivers','Sokoto','Taraba','Yobe',
  'Zamfara','Abuja',
]
const CHANNELS = ['Retail','Horeca','Modern Trade']

function canonical(s) {
  const sl = s.trim().toLowerCase()
  for (const st of ALL_STATES) if (st.toLowerCase() === sl) return st
  const aliases = { 'phc':'Rivers','fct':'Abuja','akwa ibom':'Akwa Ibom','cross river':'Cross River','crossriver':'Cross River' }
  return aliases[sl] || null
}

function resolveStates(r) {
  const rl = r.trim().toLowerCase()
  if (rl === 'all location' || rl === 'all locations') return [...ALL_STATES]
  if (rl.includes('all locations except edo')) return ALL_STATES.filter(s => s !== 'Edo')
  if (rl.includes('other region except') || rl.includes('other regions except')) {
    const ex = new Set()
    if (rl.includes('lagos')) ex.add('Lagos')
    if (rl.includes('phc'))   ex.add('Rivers')
    if (rl.includes('abuja')) ex.add('Abuja')
    return ALL_STATES.filter(s => !ex.has(s))
  }
  if (rl.includes('lagos and other regions except phc')) return ALL_STATES.filter(s => s !== 'Rivers')
  if (r.includes(',')) { const p = r.split(',').map(canonical).filter(Boolean); return p.length ? p : [r] }
  const c = canonical(r); return c ? [c] : [r]
}

function isGeneric(r) {
  const rl = r.trim().toLowerCase()
  return rl === 'all location' || rl === 'all locations' || rl.includes('other region') || rl.includes('all locations except') || r.includes(',')
}

function formatDate(val) {
  if (!val) return ''
  try { const d = new Date(val); if (isNaN(d)) return String(val); return d.toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'}) }
  catch { return String(val) }
}

function processBuffer(buffer, sheetName) {
  const wb = XLSX.read(buffer, { type:'buffer', cellDates:true })
  const targetSheet = sheetName && wb.SheetNames.includes(sheetName) ? sheetName : wb.SheetNames[0]
  const ws = wb.Sheets[targetSheet]
  const rows = XLSX.utils.sheet_to_json(ws, { defval:'', range:1 })
if (!rows.length) throw new Error(`Sheet "${targetSheet}" is empty`)

  const norm = r => { const o={}; Object.entries(r).forEach(([k,v])=>{ o[k.trim()]=v }); return o }
  const df = rows.map(norm)
  const cols = Object.keys(df[0])
 const skuCol = cols.find(c=>c==='SKU Code')||cols.find(c=>c==='SKU')||cols[0]
  const stateCol = cols.find(c=>c.toLowerCase()==='state')||'State'
  const dateCol  = cols.find(c=>c.toLowerCase().includes('date'))||'Date Updated'

  const expanded = []
  for (const row of df) {
    const raw = String(row[stateCol]||'').trim()
    for (const st of resolveStates(raw))
      for (const ch of CHANNELS)
        expanded.push({...row, _State:st, _Channel:ch, _raw:raw})
  }

  expanded.sort((a,b) => {
    const s = String(a[skuCol]||'').localeCompare(String(b[skuCol]||'')); if(s) return s
    const t = a._State.localeCompare(b._State); if(t) return t
    const u = a._Channel.localeCompare(b._Channel); if(u) return u
    const dA=new Date(a[dateCol]||0),dB=new Date(b[dateCol]||0); if(dB-dA) return dB-dA
    return (isGeneric(a._raw)?1:0)-(isGeneric(b._raw)?1:0)
  })

  const seen=new Set(), latest=[]
  for (const row of expanded) {
    const key=`${row[skuCol]}||${row._State}||${row._Channel}`
    if (!seen.has(key)) { seen.add(key); latest.push(row) }
  }

  const dropCols=new Set(['State','Channel','State ID','Channel ID'])
  const cleanCols=cols.filter(c=>!dropCols.has(c)&&!c.startsWith('Unnamed')&&!c.startsWith('__EMPTY'))
  const frontCols=[dateCol,'Model','Manufacturer','Category','Brand','_Channel','_State',skuCol,'SKU']
  const restCols=cleanCols.filter(c=>!frontCols.includes(c))
  const allCols=[...frontCols,...restCols]

  function toRows(arr) {
    return arr.map(row => {
      const o={}
      for (const c of allCols) {
        const key=c==='_Channel'?'Channel':c==='_State'?'State':c===skuCol&&skuCol==='A'?'SKU Code':c
        let val=c==='_Channel'?row._Channel:c==='_State'?row._State:row[c]
        if (val instanceof Date) val=formatDate(val)
        o[key]=val??''
      }
      return o
    })
  }

  const fullRows=toRows(expanded), latestRows=toRows(latest)
  const outWb=XLSX.utils.book_new()
  const sum=[['SKU Price Breakdown — Summary'],[],['Source file',FILE_NAME],['Sheet',targetSheet],
    ['Source rows',rows.length],['Full Expanded',fullRows.length],['Latest Prices',latestRows.length],
    ['Channels','Retail | Horeca | Modern Trade'],['States','36 States + FCT (37 total)']]
  XLSX.utils.book_append_sheet(outWb,XLSX.utils.aoa_to_sheet(sum),'Summary')
  const wf=XLSX.utils.json_to_sheet(fullRows)
  if(fullRows.length){const n=Object.keys(fullRows[0]).length;wf['!cols']=Array(n).fill({wch:22});wf['!autofilter']={ref:`A1:${XLSX.utils.encode_col(n-1)}1`}}
  XLSX.utils.book_append_sheet(outWb,wf,'Full Expanded')
  const wl=XLSX.utils.json_to_sheet(latestRows)
  if(latestRows.length){const n=Object.keys(latestRows[0]).length;wl['!cols']=Array(n).fill({wch:22});wl['!autofilter']={ref:`A1:${XLSX.utils.encode_col(n-1)}1`}}
  XLSX.utils.book_append_sheet(outWb,wl,'Latest Prices')
  return { buffer:XLSX.write(outWb,{bookType:'xlsx',type:'buffer'}), stats:{sourceRows:rows.length,full:fullRows.length,latest:latestRows.length} }
}

module.exports = async (req, res) => {
  res.setHeader('Access-Control-Allow-Origin','*')
  res.setHeader('Access-Control-Allow-Methods','GET,OPTIONS')
  if (req.method==='OPTIONS') { res.status(200).end(); return }
  if (req.method!=='GET') { res.status(405).json({error:'Method not allowed'}); return }

  try {
    const token = await getToken()
    const siteRes = await graphGet(`https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_HOST}:/sites/${SITE_NAME}`, token)
    const siteId = JSON.parse(siteRes.body.toString()).id

    const searchRes = await graphGet(`https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root/search(q='${encodeURIComponent(FILE_NAME)}')`, token)
    const files = JSON.parse(searchRes.body.toString()).value || []
    const file = files.find(f => f.name === FILE_NAME)
    if (!file) throw new Error(`File "${FILE_NAME}" not found`)

    const fileMetaRes = await graphGet(`https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${file.parentReference.driveId}/items/${file.id}?$select=@microsoft.graph.downloadUrl`, token)
const downloadUrl = JSON.parse(fileMetaRes.body.toString())['@microsoft.graph.downloadUrl']
const fileRes = await httpsGet(downloadUrl, {})
    const { buffer, stats } = processBuffer(fileRes.body, SHEET_NAME)

    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    res.setHeader('Content-Disposition','attachment; filename="SKU_Price_Breakdown.xlsx"')
    res.setHeader('X-Stats', JSON.stringify(stats))
    res.status(200).send(buffer)
  } catch(e) {
    console.error(e)
    res.status(500).json({ error: e.message })
  }
}
