import React, { useState, useCallback } from 'react';
import * as XLSX from 'xlsx';
import './App.css';

const CC = 'administrator@aaruush.org,secretary@aaruush.org,jointsecretary@aaruush.org,sponsorship@aaruush.org';
const DEFAULT_SUBJECT = "INVITATION FOR COLLABORATION WITH AARUUSH'26, SRM IST, CHENNAI";

const COMMITTEE_HEADS = [
  { name: 'Rudraksh Pardhi',  phone: '+91 7049523177', password: 'Aaruush@77' },
  { name: 'Brijesh Mohapatra', phone: '+91 72054 00424', password: 'Aaruush@24' },
  { name: 'Animesh Rai',       phone: '+91 6352 504 531', password: 'Aaruush@31' },
  { name: 'Shreyash Mishra',   phone: '+91 91198 64318', password: 'Aaruush@18' },
  { name: 'Mohak Dhawan',      phone: '+91 70215 88840', password: 'Aaruush@40' },
  { name: 'Krishn Raj',        phone: '+91 70048 16119', password: 'Aaruush@19' },
];

function getSignature(ch) {
  return `\nRegards,
${ch.name}
Committee Head | Sponsorship and Marketing | Aaruush'26
SRM Institute Of Science And Technology
p: ${ch.phone} | w: www.aaruush.org
e: sponsorship@aaruush.org | sponsorshipsrmuniv@gmail.com`;
}

function getDefaultBody(ch) {
  return `Respected Sir,
Greetings from Team Aaruush!
We are writing to propose a sponsorship collaboration between {{COMPANY}} and Aaruush, the annual national-level techno-management fest of SRM Institute of Science and Technology, Chennai. Aaruush'26 is scheduled for mid September 2026 and will bring together students, professionals, and innovators from across the country.
We believe that partnering with {{COMPANY}} would offer significant brand visibility and engagement with a large, diverse audience. Please find attached a detailed proposal outlining the sponsorship opportunities and benefits.
We look forward to the possibility of collaborating and await your positive response.
${getSignature(ch)}`;
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

function fillTemplate(template, row, headers) {
  let out = template;
  out = out.replace(/{{\s*([^}]+?)\s*}}/g, (match, key) => {
    const trimmed = key.trim().toLowerCase();
    const header = headers.find(h => h.trim().toLowerCase() === trimmed);
    return header ? (row[header] || '') : match;
  });
  return out;
}

async function aiPersonalise(row, emailCol, tplBody, headers) {
  const extraFields = Object.entries(row)
    .filter(([k]) => k !== emailCol)
    .map(([k, v]) => `${k}: ${v}`)
    .join('\n');

  const prompt = `You are a sponsorship outreach assistant for Aaruush'26, a national-level techno-management fest at SRM Institute of Science and Technology, Chennai.

Using the template below, write a personalized sponsorship email for this company. Naturally weave in the company name and any relevant details. Keep it professional and concise. Do NOT add new sections, do NOT change the signature, and do NOT change the structure. Return only the email body — no subject line, no preamble.

COMPANY INFO:
${extraFields}

TEMPLATE:
${tplBody}`;

  const res = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      model: 'claude-sonnet-4-6',
      max_tokens: 1000,
      messages: [{ role: 'user', content: prompt }],
    }),
  });
  if (!res.ok) throw new Error(`API ${res.status}`);
  const data = await res.json();
  if (data.content?.[0]?.text) return data.content[0].text.trim();
  throw new Error('Empty AI response');
}

function openGmailDraft(to, subject, body) {
  const p = new URLSearchParams({ view: 'cm', fs: '1', to, su: subject, body, cc: CC });
  window.open('https://mail.google.com/mail/?' + p.toString(), '_blank');
}

/* ─── LOGIN PAGE ─── */
function LoginPage({ onLogin }) {
  const [selected, setSelected] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [showPass, setShowPass] = useState(false);

  const handleLogin = () => {
    if (!selected) { setError('Please select your name.'); return; }
    const ch = COMMITTEE_HEADS.find(c => c.name === selected);
    if (!ch) { setError('Invalid selection.'); return; }
    if (password !== ch.password) { setError('Incorrect password. Try again.'); return; }
    setError('');
    onLogin(ch);
  };

  return (
    <div className="login-wrap">
      <div className="login-box">
        <div className="login-logo">
          <svg viewBox="0 0 24 24" fill="white" width="22" height="22">
            <path d="M20 4H4C2.9 4 2 4.9 2 6v12c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm0 4l-8 5-8-5V6l8 5 8-5v2z"/>
          </svg>
        </div>
        <div className="login-title">Aaruush Mail Agent</div>
        <div className="login-sub">Sponsorship & Marketing · A'26</div>

        <div className="login-field">
          <label className="login-label">Who are you?</label>
          <select
            value={selected}
            onChange={e => { setSelected(e.target.value); setError(''); }}
            className="login-select"
          >
            <option value="">— Select your name —</option>
            {COMMITTEE_HEADS.map(ch => (
              <option key={ch.name} value={ch.name}>{ch.name}</option>
            ))}
          </select>
        </div>

        <div className="login-field">
          <label className="login-label">Password</label>
          <div className="pass-wrap">
            <input
              type={showPass ? 'text' : 'password'}
              value={password}
              onChange={e => { setPassword(e.target.value); setError(''); }}
              onKeyDown={e => e.key === 'Enter' && handleLogin()}
              placeholder="Enter your password"
              className="login-input"
            />
            <button className="pass-toggle" onClick={() => setShowPass(s => !s)}>
              {showPass ? '🙈' : '👁️'}
            </button>
          </div>
        </div>

        {error && <div className="login-error">{error}</div>}

        <button className="login-btn" onClick={handleLogin}>
          Sign In →
        </button>

        <div className="login-hint">Forgot your password? Contact Rudraksh.</div>
      </div>
    </div>
  );
}

/* ─── SHARED COMPONENTS ─── */
function Dot({ color }) {
  const colors = { gray: 'rgba(255,255,255,0.2)', green: '#22C55E', amber: '#F59E0B', red: '#EF4444' };
  return (
    <span style={{
      display: 'inline-block', width: 6, height: 6, borderRadius: '50%',
      background: colors[color] || colors.gray, flexShrink: 0,
      boxShadow: color === 'green' ? '0 0 6px rgba(34,197,94,0.5)' : 'none'
    }} />
  );
}

function StatusBar({ color, msg }) {
  if (!msg) return null;
  return (
    <div className="status-bar">
      <Dot color={color} />
      <span>{msg}</span>
    </div>
  );
}

function Card({ num, title, tag, children, active }) {
  return (
    <div className={`card ${active ? 'card-active' : ''}`}>
      <div className="card-header">
        <div className="card-num">{num}</div>
        <span className="card-title">{title}</span>
        {tag && <span className={`card-tag tag-${tag.color}`}>{tag.label}</span>}
      </div>
      {children}
    </div>
  );
}

/* ─── MAIN AGENT ─── */
function Agent({ user, onLogout }) {
  const [rows, setRows] = useState([]);
  const [headers, setHeaders] = useState([]);
  const [fileName, setFileName] = useState('');
  const [uploadStatus, setUploadStatus] = useState({ color: 'gray', msg: 'No file loaded yet' });
  const [companyCol, setCompanyCol] = useState('');
  const [emailCol, setEmailCol] = useState('');
  const [subject, setSubject] = useState(DEFAULT_SUBJECT);
  const [body, setBody] = useState(getDefaultBody(user));
  const [previewIdx, setPreviewIdx] = useState(0);
  const [running, setRunning] = useState(false);
  const [logs, setLogs] = useState([]);
  const [progress, setProgress] = useState(0);
  const [metrics, setMetrics] = useState(null);
  const [runStatus, setRunStatus] = useState({ color: 'gray', msg: '' });

  const loaded = rows.length > 0;

  const processFile = useCallback((file) => {
    setUploadStatus({ color: 'amber', msg: 'Reading file…' });
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(ws, { defval: '' });
        if (!json.length) throw new Error('Sheet appears empty');
        const hdrs = Object.keys(json[0]);
        const parsed = json.map(r => { const o = {}; hdrs.forEach(h => { o[h] = String(r[h] || ''); }); return o; });
        setHeaders(hdrs);
        setRows(parsed);
        setFileName(file.name);
        setPreviewIdx(0);
        setLogs([]);
        setMetrics(null);
        setProgress(0);
        const cCol = hdrs.find(h => /company|name|brand|org/i.test(h)) || '';
        const eCol = hdrs.find(h => /email|mail/i.test(h)) || '';
        setCompanyCol(cCol);
        setEmailCol(eCol);
        setUploadStatus({ color: 'green', msg: `Loaded ${parsed.length} companies · ${hdrs.length} columns from "${file.name}"` });
      } catch (err) {
        setUploadStatus({ color: 'red', msg: 'Error: ' + err.message });
      }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const onDrop = useCallback((e) => {
    e.preventDefault();
    if (e.dataTransfer.files[0]) processFile(e.dataTransfer.files[0]);
  }, [processFile]);

  const previewRow = rows[previewIdx] || {};
  const previewBody = fillTemplate(body, previewRow, headers);
  const previewSubject = fillTemplate(subject, previewRow, headers);
  const previewTo = emailCol ? (previewRow[emailCol] || '(empty)') : '(select email column)';

  const runAgent = async () => {
    if (!emailCol) { alert('Please select the email column in Step 2.'); return; }
    setRunning(true);
    setLogs([]);
    setMetrics(null);
    setProgress(0);
    setRunStatus({ color: 'amber', msg: 'Starting agent…' });
    const addLog = (html, type) => setLogs(l => [...l, { html, type }]);
    let done = 0, skipped = 0;

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const to = row[emailCol];
      setProgress(Math.round(((i + 1) / rows.length) * 100));
      setRunStatus({ color: 'amber', msg: `Processing ${i + 1} of ${rows.length}…` });
      if (!to || !to.includes('@')) { addLog(`[${i+1}] SKIPPED — no valid email`, 'skip'); skipped++; await sleep(80); continue; }
      addLog(`[${i+1}] AI personalizing for ${to}…`, 'info');
      let emailBody;
      try {
        emailBody = await aiPersonalise(row, emailCol, body, headers);
        addLog(`[${i+1}] ✓ AI draft ready`, 'ok');
      } catch {
        emailBody = fillTemplate(body, row, headers);
        addLog(`[${i+1}] ⚠ AI unavailable — using template`, 'skip');
      }
      openGmailDraft(to, subject, emailBody);
      addLog(`[${i+1}] ✓ Gmail draft opened → ${to}`, 'ok');
      done++;
      await sleep(1000);
    }
    setMetrics({ total: rows.length, done, skipped });
    setRunStatus({ color: 'green', msg: `All done! ${done} drafts opened · ${skipped} skipped.` });
    setRunning(false);
  };

  const downloadCSV = () => {
    const lines = [['to','cc','subject','body'].map(h => `"${h}"`).join(',')];
    rows.forEach(row => {
      const to = emailCol ? (row[emailCol] || '') : '';
      const b = fillTemplate(body, row, headers);
      lines.push([to, CC, subject, b].map(v => `"${(v||'').replace(/"/g,'""')}"`).join(','));
    });
    const a = Object.assign(document.createElement('a'), {
      href: URL.createObjectURL(new Blob([lines.join('\n')], { type: 'text/csv' })),
      download: 'aaruush_mail_drafts.csv',
    });
    a.click();
  };

  return (
    <div className="app">
      <nav className="nav">
        <div className="nav-brand">
          <div className="nav-logo">
            <svg viewBox="0 0 24 24" fill="white" width="18" height="18">
              <path d="M20 4H4C2.9 4 2 4.9 2 6v12c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm0 4l-8 5-8-5V6l8 5 8-5v2z"/>
            </svg>
          </div>
          <div>
            <div className="nav-title">Aaruush Mail Agent</div>
            <div className="nav-sub">Sponsorship Outreach · A'26</div>
          </div>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
          <div className="nav-user">
            <span className="nav-user-dot" />
            {user.name}
          </div>
          <button className="logout-btn" onClick={onLogout}>Sign out</button>
        </div>
      </nav>

      <div className="page">
        <div className="hero">
          <h1>Send <span>smarter</span> sponsorship emails,<br />not more of them.</h1>
          <p>Drop your sponsor list, let AI personalize every pitch, and push them straight to Gmail drafts — ready to review and send.</p>
        </div>

        {/* STEP 1 */}
        <Card num="1" title="Upload Sponsor List" active={!loaded}>
          <div className="drop-zone" onClick={() => document.getElementById('file-input').click()} onDragOver={e => e.preventDefault()} onDrop={onDrop}>
            {loaded ? (
              <><div className="drop-icon" style={{ fontSize: 28 }}>✅</div><p style={{ color: 'var(--green)', fontWeight: 600 }}>{fileName}</p><span>{rows.length} rows · {headers.length} columns · click to change</span></>
            ) : (
              <><div className="drop-icon">📊</div><p>Click to upload or drag &amp; drop</p><span>.xlsx &nbsp;·&nbsp; .xls &nbsp;·&nbsp; .csv</span></>
            )}
          </div>
          <input type="file" id="file-input" accept=".xlsx,.xls,.csv" style={{ display: 'none' }} onChange={e => e.target.files[0] && processFile(e.target.files[0])} />
          <StatusBar {...uploadStatus} />
          {loaded && <div className="chip-wrap">{headers.map(h => <span key={h} className="chip">{`{{${h}}}`}</span>)}</div>}
        </Card>

        {/* STEP 2 */}
        {loaded && (
          <Card num="2" title="Map Columns" active>
            <div className="col-grid">
              <div>
                <label className="field-label">Company name column</label>
                <select value={companyCol} onChange={e => setCompanyCol(e.target.value)}>
                  <option value="">-- select --</option>
                  {headers.map(h => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
              <div>
                <label className="field-label">Recipient email column</label>
                <select value={emailCol} onChange={e => setEmailCol(e.target.value)}>
                  <option value="">-- select --</option>
                  {headers.map(h => <option key={h} value={h}>{h}</option>)}
                </select>
              </div>
            </div>
            <div className="table-wrap">
              <table>
                <thead><tr>{headers.slice(0,5).map(h => <th key={h}>{h}</th>)}{headers.length > 5 && <th>…</th>}</tr></thead>
                <tbody>{rows.slice(0,4).map((r,i) => <tr key={i}>{headers.slice(0,5).map(h => <td key={h} title={r[h]}>{r[h]}</td>)}{headers.length > 5 && <td style={{color:'var(--muted)'}}>…</td>}</tr>)}</tbody>
              </table>
            </div>
          </Card>
        )}

        {/* STEP 3 */}
        {loaded && (
          <Card num="3" title="Email Template" tag={{ color: 'green', label: '{{column_name}} placeholders' }}>
            <div className="inline-field">
              <span className="field-lbl">Subject</span>
              <input type="text" value={subject} onChange={e => setSubject(e.target.value)} />
            </div>
            <textarea value={body} onChange={e => setBody(e.target.value)} rows={12} />
            <div className="cc-wrap">
              <span className="cc-label">Auto CC:</span>
              {['administrator@aaruush.org','secretary@aaruush.org','jointsecretary@aaruush.org','sponsorship@aaruush.org'].map(cc => (
                <span key={cc} className="cc-tag">{cc}</span>
              ))}
            </div>
          </Card>
        )}

        {/* STEP 4 */}
        {loaded && (
          <Card num="4" title="Preview" tag={{ color: 'purple', label: '✦ AI personalized' }}>
            <div className="preview-nav">
              <span>Company {previewIdx + 1} of {rows.length}</span>
              <div className="nav-btns">
                <button className="btn-icon" onClick={() => setPreviewIdx(i => Math.max(0, i-1))}>←</button>
                <button className="btn-icon" onClick={() => setPreviewIdx(i => Math.min(rows.length-1, i+1))}>→</button>
              </div>
            </div>
            <div className="pf-row"><span className="pf-lbl">To</span><span className="pf-val">{previewTo}</span></div>
            <div className="pf-row"><span className="pf-lbl">CC</span><span className="pf-val" style={{color:'var(--muted)',fontSize:12}}>administrator@aaruush.org + 3 others</span></div>
            <div className="pf-row"><span className="pf-lbl">Subject</span><span className="pf-val">{previewSubject}</span></div>
            <div className="preview-body">{previewBody}</div>
          </Card>
        )}

        {/* STEP 5 */}
        {loaded && (
          <Card num="5" title="Generate Gmail Drafts">
            <div className="popup-warn">
              <strong>⚠ Allow popups from mail.google.com</strong> — Click the blocked popup icon in the address bar → <em>"Always allow popups from mail.google.com"</em>. Then click Generate again.
            </div>
            <div className="btn-row">
              <button className="btn-primary" onClick={runAgent} disabled={running || !emailCol}>
                {running ? '⏳ Running…' : '✦ Generate All Drafts ↗'}
              </button>
              <button className="btn-sec" onClick={downloadCSV}>↓ Download CSV backup</button>
            </div>
            {runStatus.msg && (
              <>
                <StatusBar {...runStatus} />
                <div className="prog-wrap"><div className="prog-bar" style={{ width: progress + '%' }} /></div>
              </>
            )}
            {logs.length > 0 && (
              <div className="log-box">{logs.map((l,i) => <div key={i} className={`log-${l.type}`}>{l.html}</div>)}</div>
            )}
            {metrics && (
              <div className="metrics">
                <div className="metric"><div className="metric-val">{metrics.total}</div><div className="metric-lbl">Total rows</div></div>
                <div className="metric"><div className="metric-val" style={{color:'var(--green)'}}>{metrics.done}</div><div className="metric-lbl">Drafts opened</div></div>
                <div className="metric"><div className="metric-val" style={{color:'var(--red)'}}>{metrics.skipped}</div><div className="metric-lbl">Skipped</div></div>
              </div>
            )}
          </Card>
        )}
      </div>
    </div>
  );
}

/* ─── ROOT ─── */
export default function App() {
  const [user, setUser] = useState(null);
  if (!user) return <LoginPage onLogin={setUser} />;
  return <Agent user={user} onLogout={() => setUser(null)} />;
}
