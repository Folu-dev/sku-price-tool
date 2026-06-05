// SharePoint → Microsoft Graph API integration
// Credentials are injected as Vercel environment variables at build time

const TENANT_ID   = import.meta.env.VITE_TENANT_ID
const CLIENT_ID   = import.meta.env.VITE_CLIENT_ID
const CLIENT_SECRET = import.meta.env.VITE_CLIENT_SECRET

const SHAREPOINT_HOST = 'omnibiz0.sharepoint.com'
const SITE_NAME       = 'FinancialPlaninngAnalysis'
const FILE_NAME       = 'Master Price Update vF.xlsx'
const SHEET_NAME      = 'MMP New Structure'

// ── Get OAuth2 token from Microsoft ─────────────────────────────────────────
async function getAccessToken() {
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`
  const body = new URLSearchParams({
    grant_type:    'client_credentials',
    client_id:     CLIENT_ID,
    client_secret: CLIENT_SECRET,
    scope:         'https://graph.microsoft.com/.default',
  })

  const res = await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body,
  })

  if (!res.ok) {
    const err = await res.json().catch(() => ({}))
    throw new Error(`Authentication failed: ${err.error_description || res.statusText}`)
  }

  const data = await res.json()
  return data.access_token
}

// ── Get SharePoint site ID ───────────────────────────────────────────────────
async function getSiteId(token) {
  const url = `https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_HOST}:/sites/${SITE_NAME}`
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  })

  if (!res.ok) {
    const err = await res.json().catch(() => ({}))
    throw new Error(`Could not find SharePoint site: ${err.error?.message || res.statusText}`)
  }

  const data = await res.json()
  return data.id
}

// ── Find the file in the site's drive ───────────────────────────────────────
async function getFileId(token, siteId) {
  const encoded = encodeURIComponent(FILE_NAME)
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root/search(q='${encoded}')`
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  })

  if (!res.ok) {
    const err = await res.json().catch(() => ({}))
    throw new Error(`Could not search for file: ${err.error?.message || res.statusText}`)
  }

  const data = await res.json()
  const file = data.value?.find(f => f.name === FILE_NAME)
  if (!file) throw new Error(`File "${FILE_NAME}" not found in SharePoint site`)
  return { fileId: file.id, driveId: file.parentReference?.driveId }
}

// ── Download file as binary (ArrayBuffer) ───────────────────────────────────
async function downloadFile(token, siteId, driveId, fileId) {
  const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${fileId}/content`
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  })

  if (!res.ok) {
    const err = await res.json().catch(() => ({}))
    throw new Error(`Could not download file: ${err.error?.message || res.statusText}`)
  }

  return await res.arrayBuffer()
}

// ── Main export: fetch the SharePoint file as ArrayBuffer ───────────────────
export async function fetchSharePointFile(onProgress) {
  onProgress?.('Connecting to Microsoft 365…')
  const token = await getAccessToken()

  onProgress?.('Locating SharePoint site…')
  const siteId = await getSiteId(token)

  onProgress?.('Finding price list file…')
  const { fileId, driveId } = await getFileId(token, siteId)

  onProgress?.('Downloading latest file…')
  const buffer = await downloadFile(token, siteId, driveId, fileId)

  return { buffer, sheetName: SHEET_NAME }
}

export { SHEET_NAME, FILE_NAME }
