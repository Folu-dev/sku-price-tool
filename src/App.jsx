import { useState, useRef, useCallback } from 'react'
import './App.css'

const RULES = [
  { icon: '⇄', label: 'All Channels', desc: 'Expanded into Retail, Horeca, and Modern Trade — one row each' },
  { icon: '⊙', label: 'All Location', desc: 'Expanded into all 36 states + FCT Abuja (37 rows per SKU)' },
  { icon: '∖', label: 'Other region except X', desc: 'All 37 states minus the excluded state(s)' },
  { icon: '⋮', label: 'Comma-separated states', desc: 'Each state in the list gets its own individual row' },
  { icon: '◷', label: 'Latest price logic', desc: 'Specific-state date wins over All Location; most recent date always takes priority' },
]

export default function App() {
  const [phase, setPhase] = useState('idle')
  const [message, setMessage] = useState('')
  const [fileName, setFileName] = useState('')
  const [stats, setStats] = useState(null)
  const [error, setError] = useState('')
  const [dragging, setDragging] = useState(false)
  const [spLoading, setSpLoading] = useState(false)
  const outputRef = useRef(null)
  const fileInputRef = useRef(null)

  const downloadBlob = (blob, filename) => {
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = filename
    a.click()
    URL.revokeObjectURL(url)
  }

  const runUpload = useCallback(async (file) => {
    if (!file) return
    if (!/\.(xlsx|xls)$/i.test(file.name)) {
      setError('Please upload an Excel file (.xlsx or .xls)')
      setPhase('error')
      return
    }

    setFileName(file.name)
    setPhase('processing')
    setMessage('Uploading file to server…')
    setError('')

    try {
      const formData = new FormData()
      formData.append('file', file)

      setMessage('Processing on server — this may take 20–30 seconds…')
      const res = await fetch('/api/process', { method: 'POST', body: formData })

      if (!res.ok) {
        const err = await res.json().catch(() => ({}))
        throw new Error(err.error || `Server error ${res.status}`)
      }

      const statsHeader = res.headers.get('X-Stats')
      if (statsHeader) setStats(JSON.parse(statsHeader))

      const blob = await res.blob()
      outputRef.current = blob
      setPhase('done')
    } catch (e) {
      setError(e.message || 'Processing failed')
      setPhase('error')
    }
  }, [])

  const runSharePoint = async () => {
    setSpLoading(true)
    setPhase('processing')
    setFileName('Master Price Update vF.xlsx (SharePoint)')
    setMessage('Connecting to SharePoint…')
    setError('')

    try {
      setMessage('Fetching latest file from SharePoint and processing…')
      const res = await fetch('/api/sharepoint')

      if (!res.ok) {
        const err = await res.json().catch(() => ({}))
        throw new Error(err.error || `Server error ${res.status}`)
      }

      const statsHeader = res.headers.get('X-Stats')
      if (statsHeader) setStats(JSON.parse(statsHeader))

      const blob = await res.blob()
      outputRef.current = blob
      setPhase('done')
    } catch (e) {
      setError(e.message || 'SharePoint fetch failed')
      setPhase('error')
    } finally {
      setSpLoading(false)
    }
  }

  const handleDrop = useCallback((e) => {
    e.preventDefault()
    setDragging(false)
    runUpload(e.dataTransfer.files[0])
  }, [runUpload])

  const handleDownload = () => {
    if (!outputRef.current) return
    downloadBlob(outputRef.current, 'SKU_Price_Breakdown.xlsx')
  }

  const reset = () => {
    setPhase('idle')
    setStats(null)
    setFileName('')
    setMessage('')
    setError('')
    outputRef.current = null
    if (fileInputRef.current) fileInputRef.current.value = ''
  }

  return (
    <div className="page">
      <aside className="sidebar">
        <div className="logo">
          <div className="logo-icon"><span /><span /><span /></div>
          <p className="logo-name">Price Breakdown</p>
        </div>
        <nav className="rules-nav">
          <p className="nav-label">Rules applied</p>
          {RULES.map((r, i) => (
            <div key={i} className="rule-row">
              <span className="rule-sym">{r.icon}</span>
              <div>
                <p className="rule-name">{r.label}</p>
                <p className="rule-desc">{r.desc}</p>
              </div>
            </div>
          ))}
        </nav>
        <div className="sidebar-footer">
          <p>Processing happens on the server — no browser memory limits. Works with any size file.</p>
        </div>
      </aside>

      <main className="main">
        <header className="main-header">
          <div>
            <h1>SKU Price Breakdown Tool</h1>
            <p>Upload your price list or fetch directly from SharePoint</p>
          </div>
        </header>

        <div className="content">
          {phase === 'idle' && (
            <>
              {/* SharePoint auto-sync button */}
              <div className="sp-card">
                <div className="sp-left">
                  <div className="sp-icon">
                    <svg width="24" height="24" viewBox="0 0 24 24" fill="none">
                      <rect x="2" y="3" width="20" height="18" rx="3" fill="#0078D4" opacity=".15"/>
                      <path d="M7 8h10M7 12h7M7 16h5" stroke="#0078D4" strokeWidth="1.8" strokeLinecap="round"/>
                    </svg>
                  </div>
                  <div>
                    <p className="sp-title">Fetch from SharePoint</p>
                    <p className="sp-sub">Auto-syncs Master Price Update vF.xlsx · Sheet: MMP New Structure</p>
                  </div>
                </div>
                <button className="btn-sp" onClick={runSharePoint} disabled={spLoading}>
                  {spLoading ? 'Connecting…' : 'Run breakdown'}
                </button>
              </div>

              <div className="divider"><span>or upload a file manually</span></div>

              {/* Manual upload */}
              <div
                className={`dropzone${dragging ? ' drag-over' : ''}`}
                onDragOver={e => { e.preventDefault(); setDragging(true) }}
                onDragLeave={() => setDragging(false)}
                onDrop={handleDrop}
                onClick={() => fileInputRef.current?.click()}
                role="button" tabIndex={0}
                onKeyDown={e => e.key === 'Enter' && fileInputRef.current?.click()}
              >
                <div className="dz-icon">
                  <svg width="40" height="40" viewBox="0 0 40 40" fill="none">
                    <rect x="6" y="4" width="22" height="28" rx="3" fill="var(--blue-light)" stroke="var(--blue)" strokeWidth="1.5"/>
                    <rect x="10" y="10" width="14" height="2.5" rx="1.25" fill="var(--blue)"/>
                    <rect x="10" y="15" width="10" height="2.5" rx="1.25" fill="var(--blue)" opacity=".5"/>
                    <rect x="10" y="20" width="12" height="2.5" rx="1.25" fill="var(--blue)" opacity=".3"/>
                    <circle cx="30" cy="30" r="8" fill="var(--blue)"/>
                    <path d="M30 26v8M26 30l4-4 4 4" stroke="white" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"/>
                  </svg>
                </div>
                <p className="dz-title">Drop your Excel file here</p>
                <p className="dz-sub">or click anywhere to browse</p>
                <div className="dz-btn">Choose file</div>
                <p className="dz-hint">.xlsx or .xls supported</p>
                <input ref={fileInputRef} type="file" accept=".xlsx,.xls"
                  style={{ display: 'none' }}
                  onChange={e => runUpload(e.target.files[0])} />
              </div>

              <div className="output-info">
                <p className="oi-title">What you'll get</p>
                <div className="oi-grid">
                  <div className="oi-item"><p className="oi-name">📋 Summary</p><p className="oi-desc">Row counts, rules applied, and logic explanation</p></div>
                  <div className="oi-item"><p className="oi-name">📊 Full Expanded</p><p className="oi-desc">All historical records broken out by channel and state</p></div>
                  <div className="oi-item"><p className="oi-name">✅ Latest Prices</p><p className="oi-desc">Current active price per SKU × channel × state</p></div>
                </div>
              </div>
            </>
          )}

          {phase === 'processing' && (
            <div className="proc-card">
              <div className="proc-file">
                <svg width="18" height="18" viewBox="0 0 18 18" fill="none">
                  <rect x="2" y="1" width="11" height="14" rx="2" fill="var(--blue-light)" stroke="var(--blue)" strokeWidth="1.2"/>
                  <rect x="4" y="5" width="7" height="1.5" rx=".75" fill="var(--blue)"/>
                </svg>
                <span>{fileName}</span>
              </div>
              <div className="proc-spinner-wrap">
                <div className="proc-spinner" />
              </div>
              <p className="proc-msg">{message}</p>
            </div>
          )}

          {phase === 'done' && (
            <div className="done-card">
              <div className="done-header">
                <div className="done-check">
                  <svg width="22" height="22" viewBox="0 0 22 22">
                    <circle cx="11" cy="11" r="11" fill="var(--green)"/>
                    <path d="M6 11l3.5 3.5 6.5-7" stroke="white" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"/>
                  </svg>
                </div>
                <div>
                  <h2>Breakdown complete</h2>
                  <p className="done-file">{fileName}</p>
                </div>
              </div>

              {stats && (
                <div className="stats-row">
                  <div className="stat"><p className="stat-n">{stats.sourceRows.toLocaleString()}</p><p className="stat-l">Source rows</p></div>
                  <div className="stat-arrow">→</div>
                  <div className="stat highlight"><p className="stat-n">{stats.full.toLocaleString()}</p><p className="stat-l">Full expanded</p></div>
                  <div className="stat"><p className="stat-n">{stats.latest.toLocaleString()}</p><p className="stat-l">Latest prices</p></div>
                </div>
              )}

              <div className="sheets-preview">
                <div className="sheet-chip"><span className="sheet-dot s1" />Summary</div>
                <div className="sheet-chip"><span className="sheet-dot s2" />Full Expanded</div>
                <div className="sheet-chip"><span className="sheet-dot s3" />Latest Prices</div>
              </div>

              <div className="done-actions">
                <button className="btn-download" onClick={handleDownload}>
                  <svg width="18" height="18" viewBox="0 0 18 18" fill="none">
                    <path d="M9 2v10M5 8l4 4 4-4" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"/>
                    <path d="M2 14h14" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round"/>
                  </svg>
                  Download SKU_Price_Breakdown.xlsx
                </button>
                <button className="btn-reset" onClick={reset}>Process another file</button>
              </div>
            </div>
          )}

          {phase === 'error' && (
            <div className="error-card">
              <div className="error-icon">
                <svg width="22" height="22" viewBox="0 0 22 22">
                  <circle cx="11" cy="11" r="11" fill="var(--red-light)" stroke="var(--red)" strokeWidth="1.2"/>
                  <path d="M11 7v5M11 15v.5" stroke="var(--red)" strokeWidth="2" strokeLinecap="round"/>
                </svg>
              </div>
              <div>
                <p className="error-title">Could not process file</p>
                <p className="error-msg">{error}</p>
                <button className="btn-reset" onClick={reset} style={{ marginTop: '12px' }}>Try again</button>
              </div>
            </div>
          )}
        </div>
      </main>
    </div>
  )
}
