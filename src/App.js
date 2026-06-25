import React, { useState, useCallback, useRef } from 'react';
import * as XLSX from 'xlsx';
import { downloadCompanyDocs } from './docxTemplates';
import './App.css';

const CC = 'administrator@aaruush.org,secretary@aaruush.org,jointsecretary@aaruush.org,sponsorship@aaruush.org';
const DEFAULT_SUBJECT = "INVITATION FOR COLLABORATION WITH AARUUSH'26, SRM IST, CHENNAI";
const LOGO_URL = 'https://s3.ap-south-1.amazonaws.com/townscript-production/images/8bb5c3cb-bd88-4b2e-8f7c-434395f00045.jpg';

const SOCIAL_LINKS = [
  { label: 'Facebook',  url: 'https://www.facebook.com/aaruush.srm' },
  { label: 'Instagram', url: 'https://www.instagram.com/aaruush_srm/' },
  { label: 'Twitter',   url: 'https://x.com/aaruushsrmist' },
  { label: 'Website',   url: 'https://www.aaruush.net/' },
  { label: 'YouTube',   url: 'https://www.youtube.com/channel/UC6mwWpwkZchii-oyWz0v3dw' },
];

const COMMITTEE_HEADS = [
  { name: 'Rudraksh Pardhi', phone: '+91 7049523177', password: 'Aaruush@77', role: 'Committee Head' },
  { name: 'Brijesh Mohapatra', phone: '+91 72054 00424', password: 'Aaruush@24', role: 'Committee Head' },
  { name: 'Animesh Rai', phone: '+91 6352 504 531', password: 'Aaruush@31', role: 'Committee Head' },
  { name: 'Shreyash Mishra', phone: '+91 91198 64318', password: 'Aaruush@18', role: 'Committee Head' },
  { name: 'Mohak Dhawan', phone: '+91 70215 88840', password: 'Aaruush@40', role: 'Committee Head' },
  { name: 'Krishn Raj', phone: '+91 70048 16119', password: 'Aaruush@19', role: 'Committee Head' },
  { name: 'Siddharth Agarwal', phone: '+91 87450 72181', password: 'Aaruush@81', role: 'Organizer' },
  { name: 'Vansh Gupta', phone: '+91 84480 67969', password: 'Aaruush@69', role: 'Organizer' },
  { name: 'Devansh Gupta', phone: '+91 95205 34441', password: 'Aaruush@41', role: 'Organizer' },
  { name: 'Riddhi Shukla', phone: '+91 93614 60998', password: 'Aaruush@98', role: 'Committee Head' },
];

const ADMINS = ['Rudraksh Pardhi'];

function isOrganizer(user) {
  return user && user.role === 'Organizer';
}

function makePassword(phone) {
  const digits = (phone || '').replace(/\D/g, '');
  return 'Aaruush@' + (digits.slice(-2) || '00');
}

function generateCHCode(list) {
  const lines = list.map(ch =>
    `  { name: '${ch.name.replace(/'/g, "\\'")}', phone: '${ch.phone}', password: '${ch.password}', role: '${ch.role || 'Committee Head'}' },`
  ).join('\n');
  return `const COMMITTEE_HEADS = [\n${lines}\n];`;
}

function getBodySignature(ch) {
  return `\nRegards,`;
}

function getDefaultBody(ch) {
  return `Respected Sir,
Greetings from Team AARUUSH!
We are writing to propose a sponsorship collaboration between {{COMPANY}} and AARUUSH, the annual national-level techno-management fest of SRM Institute of Science and Technology, Chennai. AARUUSH'26 is scheduled for mid September 2026 and will bring together students, professionals, and innovators from across the country.
We believe that partnering with {{COMPANY}} would offer significant brand visibility and engagement with a large, diverse audience. Please find attached a detailed proposal outlining the sponsorship opportunities and benefits.
We look forward to the possibility of collaborating and await your positive response.
${getBodySignature(ch)}`;
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

// Render body text for preview with AARUUSH and company name bolded
function renderBoldedBody(text, companyName) {
  // Build a list of terms to bold (longest first to avoid partial overlaps)
  const terms = [];
  if (companyName && companyName.trim()) terms.push(companyName.trim());
  terms.push('AARUUSH');
  // Escape regex special chars
  const escaped = terms.map(t => t.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'));
  // Match AARUUSH optionally followed by 'XX (e.g. AARUUSH'26)
  const pattern = new RegExp(`(${escaped.join('|')})('?\\d*)?`, 'g');
  const parts = [];
  let lastIdx = 0, m, key = 0;
  while ((m = pattern.exec(text)) !== null) {
    if (m.index > lastIdx) parts.push(text.slice(lastIdx, m.index));
    parts.push(<strong key={key++}>{m[0]}</strong>);
    lastIdx = m.index + m[0].length;
  }
  if (lastIdx < text.length) parts.push(text.slice(lastIdx));
  return parts;
}

// Loose match: compares first name, case-insensitive
function chMatches(cellValue, loginName) {
  if (!cellValue) return false;
  const cell = cellValue.trim().toLowerCase();
  const login = loginName.trim().toLowerCase();
  if (!cell) return false;
  if (cell === login) return true;
  // First name of login
  const loginFirst = login.split(/\s+/)[0];
  const cellFirst = cell.split(/\s+/)[0];
  // Match if either side's first name appears in the other
  if (cellFirst === loginFirst) return true;
  if (cell.includes(loginFirst) || login.includes(cellFirst)) return true;
  return false;
}

// Find the CH column header (case-insensitive)
function findCHColumn(headers) {
  return headers.find(h => h.trim().toLowerCase() === 'ch') ||
         headers.find(h => /committee\s*head/i.test(h)) || null;
}

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

Using the template below, write a personalized sponsorship email for this company. Naturally weave in the company name and any relevant details. Keep it professional and concise. Do NOT change the signature block at the end. Return only the email body — no subject line, no preamble.

COMPANY INFO:
${extraFields}

TEMPLATE:
${tplBody}`;
  const res = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ model: 'claude-sonnet-4-6', max_tokens: 1000, messages: [{ role: 'user', content: prompt }] }),
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

/* ── SIGNATURE BLOCK (exact layout match) ── */
function SignatureBlock({ user }) {
  return (
    <div style={{ fontFamily: 'Arial, sans-serif', fontSize: 13, lineHeight: 1.5, color: '#333', marginTop: 8 }}>
      <div style={{ fontWeight: 700, fontSize: 15, marginBottom: 2 }}>{user.name}</div>
      <div style={{ marginBottom: 2 }}>
        <span style={{ color: '#E8540A', fontWeight: 700 }}>{user.role || 'Committee Head'}</span>
        <span style={{ color: '#333' }}> | <strong>Sponsorship and Marketing</strong></span>
      </div>
      <div style={{ color: '#333', marginBottom: 1 }}>Aaruush'26</div>
      <div style={{ color: '#333', marginBottom: 8 }}>SRM Institute Of Science And Technology</div>
      <div style={{ display: 'flex', gap: 24, marginBottom: 4, flexWrap: 'wrap' }}>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
          <span><span style={{ color: '#E8540A' }}>p:</span> {user.phone}</span>
          <span><span style={{ color: '#E8540A' }}>w:</span> <a href="https://www.aaruush.org" target="_blank" rel="noreferrer" style={{ color: '#E8540A' }}>www.aaruush.org</a></span>
          <span><span style={{ color: '#E8540A' }}>e:</span> <a href="mailto:sponsorship@aaruush.org" style={{ color: '#E8540A' }}>sponsorship@aaruush.org</a></span>
          <span><span style={{ color: '#E8540A' }}>e:</span> <a href="mailto:sponsorshipsrmuniv@gmail.com" style={{ color: '#E8540A' }}>sponsorshipsrmuniv@gmail.com</a></span>
          <span><span style={{ color: '#E8540A' }}>a:</span> The Aaruush Room, 8th Floor, Classroom Complex</span>
          <span style={{ paddingLeft: 14 }}>SRMIST Kattankulathur, Tamil Nadu - 603203</span>
        </div>
      </div>
      <div style={{ marginBottom: 6, color: '#333', fontSize: 12 }}>Follow us:</div>
      <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
        {SOCIAL_LINKS.map(s => (
          <a key={s.label} href={s.url} target="_blank" rel="noreferrer"
            style={{ fontSize: 11, color: '#E8540A', textDecoration: 'none', border: '1px solid #E8540A', borderRadius: 4, padding: '2px 8px' }}>
            {s.label}
          </a>
        ))}
      </div>
    </div>
  );
}

/* ── FULL PREVIEW EMAIL (with logo left, sig right like screenshot) ── */
function EmailPreview({ user, body, subject, to, companyName }) {
  const previewRef = useRef(null);
  const [copied, setCopied] = useState(false);

  const copyAll = async () => {
    try {
      // Use clipboard API with html format for rich copy
      const el = previewRef.current;
      const range = document.createRange();
      range.selectNodeContents(el);
      const sel = window.getSelection();
      sel.removeAllRanges();
      sel.addRange(range);
      document.execCommand('copy');
      sel.removeAllRanges();
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch {
      setCopied(false);
    }
  };

  return (
    <div>
      {/* Email header fields */}
      <div className="email-header-fields">
        <div className="pf-row"><span className="pf-lbl">To</span><span className="pf-val">{to}</span></div>
        <div className="pf-row"><span className="pf-lbl">CC</span><span className="pf-val" style={{color:'var(--muted)',fontSize:12}}>administrator@aaruush.org + 3 others</span></div>
        <div className="pf-row"><span className="pf-lbl">Subject</span><span className="pf-val">{subject}</span></div>
      </div>

      {/* Copy button */}
      <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 8 }}>
        <button className="btn-copy" onClick={copyAll}>
          {copied ? '✓ Copied!' : '⎘ Copy email body'}
        </button>
      </div>

      {/* Email body preview — white bg like real email */}
      <div ref={previewRef} className="email-render" style={{ background: '#ffffff', borderRadius: 8, padding: '1.25rem 1.5rem', border: '1px solid #ddd', color: '#1a1a1a' }}>
        {/* Body text */}
        <div style={{ fontFamily: 'Arial, sans-serif', fontSize: 14, lineHeight: 1.7, color: '#1a1a1a', whiteSpace: 'pre-wrap', marginBottom: 16 }}>
          {renderBoldedBody(body, companyName)}
        </div>

        {/* Divider */}
        <hr style={{ border: 'none', borderTop: '1px solid #ddd', margin: '12px 0' }} />

        {/* Signature row: logo left, details right */}
        <div style={{ display: 'flex', gap: 16, alignItems: 'flex-start' }}>
          <div style={{ flexShrink: 0, textAlign: 'center' }}>
            <img src={LOGO_URL} alt="Aaruush SRMIST" style={{ width: 80, height: 80, objectFit: 'contain', borderRadius: 6, display: 'block' }} />
            <div style={{ fontSize: 11, fontWeight: 700, color: '#1a1a1a', marginTop: 4 }}>AARUUSH,</div>
            <div style={{ fontSize: 11, fontWeight: 700, color: '#1a1a1a' }}>SRMIST</div>
          </div>
          <div style={{ borderLeft: '3px solid #E8540A', paddingLeft: 12 }}>
            <SignatureBlock user={user} />
          </div>
        </div>
      </div>
    </div>
  );
}

/* ── LOGIN ── */
function LoginPage({ onLogin, chList }) {
  const [selected, setSelected] = useState('');
  const [password, setPassword] = useState('');
  const [error, setError] = useState('');
  const [showPass, setShowPass] = useState(false);
  const handleLogin = () => {
    if (!selected) { setError('Please select your name.'); return; }
    const ch = chList.find(c => c.name === selected);
    if (!ch) { setError('Invalid selection.'); return; }
    if (password !== ch.password) { setError('Incorrect password. Try again.'); return; }
    setError(''); onLogin(ch);
  };
  return (
    <div className="login-wrap">
      <div className="login-box">
        <div className="login-logo"><svg viewBox="0 0 24 24" fill="white" width="22" height="22"><path d="M20 4H4C2.9 4 2 4.9 2 6v12c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm0 4l-8 5-8-5V6l8 5 8-5v2z"/></svg></div>
        <div className="login-title">Aaruush Mail Agent</div>
        <div className="login-sub">Sponsorship & Marketing · A'26</div>
        <div className="login-field">
          <label className="login-label">Who are you?</label>
          <select value={selected} onChange={e => { setSelected(e.target.value); setError(''); }} className="login-select">
            <option value="">— Select your name —</option>
            {chList.map(ch => <option key={ch.name} value={ch.name}>{ch.name} — {ch.role || 'Committee Head'}</option>)}
          </select>
        </div>
        <div className="login-field">
          <label className="login-label">Password</label>
          <div className="pass-wrap">
            <input type={showPass ? 'text' : 'password'} value={password}
              onChange={e => { setPassword(e.target.value); setError(''); }}
              onKeyDown={e => e.key === 'Enter' && handleLogin()}
              placeholder="Enter your password" className="login-input" />
            <button className="pass-toggle" onClick={() => setShowPass(s => !s)}>{showPass ? '🙈' : '👁️'}</button>
          </div>
        </div>
        {error && <div className="login-error">{error}</div>}
        <button className="login-btn" onClick={handleLogin}>Sign In →</button>
        <div className="login-hint">Forgot your password? Contact Rudraksh.</div>
        <div className="login-credit">
          <div className="credit-label">Built by</div>
          <div className="credit-names">Rudraksh Pardhi &nbsp;·&nbsp; Animesh Rai</div>
          <div className="credit-role">Committee Heads · Sponsorship & Marketing</div>
        </div>
      </div>
    </div>
  );
}

/* ── SHARED ── */
function Dot({ color }) {
  const colors = { gray:'rgba(255,255,255,0.2)', green:'#22C55E', amber:'#F59E0B', red:'#EF4444' };
  return <span style={{ display:'inline-block', width:6, height:6, borderRadius:'50%', background:colors[color]||colors.gray, flexShrink:0, boxShadow:color==='green'?'0 0 6px rgba(34,197,94,0.5)':'none' }} />;
}
function StatusBar({ color, msg }) {
  if (!msg) return null;
  return <div className="status-bar"><Dot color={color} /><span>{msg}</span></div>;
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

/* ── MANUAL ENTRY ── */
function ManualEntry({ onAdd }) {
  const [company, setCompany] = useState('');
  const [email, setEmail] = useState('');
  const [extra, setExtra] = useState('');
  const [error, setError] = useState('');
  const handleAdd = () => {
    if (!company.trim()) { setError('Company name is required.'); return; }
    if (!email.trim() || !email.includes('@')) { setError('Valid email is required.'); return; }
    onAdd({ COMPANY: company.trim(), EMAIL: email.trim(), NOTES: extra.trim() });
    setCompany(''); setEmail(''); setExtra(''); setError('');
  };
  return (
    <div className="manual-form">
      <div className="manual-grid">
        <div><label className="field-label">Company Name *</label><input type="text" value={company} onChange={e => { setCompany(e.target.value); setError(''); }} placeholder="e.g. Dabur India Ltd" /></div>
        <div><label className="field-label">Email Address *</label><input type="text" value={email} onChange={e => { setEmail(e.target.value); setError(''); }} placeholder="e.g. marketing@dabur.com" /></div>
      </div>
      <div style={{ marginTop: 8 }}>
        <label className="field-label">Extra Notes (optional — helps AI personalize)</label>
        <input type="text" value={extra} onChange={e => setExtra(e.target.value)} placeholder="e.g. FMCG brand, youth-focused, sponsors cricket" />
      </div>
      {error && <div className="manual-error">{error}</div>}
      <button className="btn-add" onClick={handleAdd}>+ Add Company</button>
    </div>
  );
}

/* ── MAIN AGENT ── */
/* ── ADMIN PANEL (Rudraksh only) ── */
function AdminPanel({ chList, setChList }) {
  const [open, setOpen] = useState(false);
  const [name, setName] = useState('');
  const [phone, setPhone] = useState('');
  const [role, setRole] = useState('Committee Head');
  const [error, setError] = useState('');
  const [copied, setCopied] = useState(false);

  const addCH = () => {
    if (!name.trim()) { setError('Name is required.'); return; }
    if (!phone.trim()) { setError('Phone is required.'); return; }
    if (chList.some(c => c.name.trim().toLowerCase() === name.trim().toLowerCase())) {
      setError('A member with this name already exists.'); return;
    }
    const newCH = { name: name.trim(), phone: phone.trim(), password: makePassword(phone), role };
    setChList([...chList, newCH]);
    setName(''); setPhone(''); setRole('Committee Head'); setError('');
  };

  const removeCH = (idx) => {
    if (chList[idx].name === 'Rudraksh Pardhi') { setError("Can't remove the admin account."); return; }
    setChList(chList.filter((_, i) => i !== idx));
  };

  const code = generateCHCode(chList);

  const copyCode = async () => {
    try { await navigator.clipboard.writeText(code); setCopied(true); setTimeout(() => setCopied(false), 2000); }
    catch { setCopied(false); }
  };

  return (
    <div className="admin-wrap">
      <button className="admin-toggle" onClick={() => setOpen(o => !o)}>
        <span>⚙️ Manage Committee Heads</span>
        <span>{open ? '▲' : '▼'}</span>
      </button>
      {open && (
        <div className="admin-body">
          <div className="admin-note">
            Add or remove CHs here, then copy the generated code and paste it into <code>src/App.js</code> on GitHub (replace the <code>COMMITTEE_HEADS</code> block) and commit. Changes apply for everyone after deploy.
          </div>

          <div className="admin-form">
            <div className="admin-grid">
              <div><label className="field-label">Full Name</label><input type="text" value={name} onChange={e => { setName(e.target.value); setError(''); }} placeholder="e.g. Aarav Sharma" /></div>
              <div><label className="field-label">Phone</label><input type="text" value={phone} onChange={e => { setPhone(e.target.value); setError(''); }} placeholder="e.g. +91 98765 43210" /></div>
            </div>
            <div style={{ marginTop: 8 }}>
              <label className="field-label">Role</label>
              <select value={role} onChange={e => setRole(e.target.value)}>
                <option value="Committee Head">Committee Head (sees only assigned companies)</option>
                <option value="Organizer">Organizer (sees all companies)</option>
              </select>
            </div>
            {phone.trim() && <div className="admin-pass-preview">Password will be: <strong>{makePassword(phone)}</strong></div>}
            {error && <div className="manual-error">{error}</div>}
            <button className="btn-add" onClick={addCH}>+ Add Member</button>
          </div>

          <div className="admin-list">
            {chList.map((ch, i) => (
              <div key={i} className="admin-row">
                <div className="admin-row-info">
                  <span className="admin-ch-name">{ch.name}{ch.name === 'Rudraksh Pardhi' && <span className="admin-badge">admin</span>}<span className="admin-role-tag">{ch.role || 'Committee Head'}</span></span>
                  <span className="admin-ch-detail">{ch.phone} · {ch.password}</span>
                </div>
                {ch.name !== 'Rudraksh Pardhi' && <button className="btn-remove" onClick={() => removeCH(i)}>✕</button>}
              </div>
            ))}
          </div>

          <div className="admin-code-header">
            <span>Code to paste into GitHub</span>
            <button className="btn-copy" onClick={copyCode}>{copied ? '✓ Copied!' : '⎘ Copy code'}</button>
          </div>
          <pre className="admin-code">{code}</pre>
        </div>
      )}
    </div>
  );
}

function Agent({ user, onLogout, chList, setChList }) {
  const [rows, setRows] = useState([]);
  const [headers, setHeaders] = useState(['COMPANY', 'EMAIL', 'NOTES']);
  const [fileName, setFileName] = useState('');
  const [uploadStatus, setUploadStatus] = useState({ color: 'gray', msg: 'No file loaded yet' });
  const [inputMode, setInputMode] = useState('upload');
  const [companyCol, setCompanyCol] = useState('COMPANY');
  const [emailCol, setEmailCol] = useState('EMAIL');
  const [subject, setSubject] = useState(DEFAULT_SUBJECT);
  const [body, setBody] = useState(getDefaultBody(user));
  const [previewIdx, setPreviewIdx] = useState(0);
  const [running, setRunning] = useState(false);
  const [logs, setLogs] = useState([]);
  const [progress, setProgress] = useState(0);
  const [metrics, setMetrics] = useState(null);
  const [runStatus, setRunStatus] = useState({ color: 'gray', msg: '' });
  const [chInfo, setChInfo] = useState(null); // { total, matched, filtered }
  const [selected, setSelected] = useState({}); // { rowIndex: true }
  const [includeDocs, setIncludeDocs] = useState(true);
  const [docsBusy, setDocsBusy] = useState(false);
  const [docsError, setDocsError] = useState('');

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
        let parsed = json.map(r => { const o = {}; hdrs.forEach(h => { o[h] = String(r[h] || ''); }); return o; });

        // Filter by CH column if it exists (organizers see all, no filter)
        const chCol = findCHColumn(hdrs);
        let chMsg = '';
        if (chCol && !isOrganizer(user)) {
          const total = parsed.length;
          const matched = parsed.filter(r => chMatches(r[chCol], user.name));
          parsed = matched;
          setChInfo({ total, matched: matched.length, chCol });
          chMsg = ` · filtered to ${matched.length} assigned to ${user.name.split(' ')[0]} (CH column found)`;
          if (matched.length === 0) {
            setHeaders(hdrs); setRows([]); setFileName(file.name);
            setUploadStatus({ color: 'amber', msg: `No companies assigned to ${user.name} in the CH column (found ${total} total rows).` });
            return;
          }
        } else if (chCol && isOrganizer(user)) {
          setChInfo({ total: parsed.length, matched: parsed.length, chCol, organizer: true });
          chMsg = ` · showing all (organizer access)`;
        } else {
          setChInfo(null);
        }

        setHeaders(hdrs); setRows(parsed); setFileName(file.name);
        setPreviewIdx(0); setLogs([]); setMetrics(null); setProgress(0);
        // Select all by default
        const sel = {}; parsed.forEach((_, i) => { sel[i] = true; });
        setSelected(sel);
        const cCol = hdrs.find(h => /company|name|brand|org/i.test(h)) || hdrs[0];
        const eCol = hdrs.find(h => /email|mail/i.test(h)) || '';
        setCompanyCol(cCol); setEmailCol(eCol);
        setUploadStatus({ color: 'green', msg: `Loaded ${parsed.length} companies${chMsg}` });
      } catch (err) { setUploadStatus({ color: 'red', msg: 'Error: ' + err.message }); }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const onDrop = useCallback((e) => { e.preventDefault(); if (e.dataTransfer.files[0]) processFile(e.dataTransfer.files[0]); }, [processFile]);

  const addManualRow = (row) => {
    setHeaders(['COMPANY', 'EMAIL', 'NOTES']);
    setCompanyCol('COMPANY'); setEmailCol('EMAIL');
    setRows(prev => {
      const next = [...prev, row];
      setSelected(s => ({ ...s, [next.length - 1]: true }));
      return next;
    });
    setUploadStatus({ color: 'green', msg: `${rows.length + 1} ${rows.length === 0 ? 'company' : 'companies'} added` });
  };

  const removeRow = (idx) => {
    setRows(prev => prev.filter((_, i) => i !== idx));
    setSelected(prev => {
      const next = {};
      Object.keys(prev).map(Number).filter(i => i !== idx).forEach(i => { next[i > idx ? i - 1 : i] = prev[i]; });
      return next;
    });
  };

  const toggleSelect = (idx) => setSelected(s => ({ ...s, [idx]: !s[idx] }));
  const selectAll = () => { const s = {}; rows.forEach((_, i) => { s[i] = true; }); setSelected(s); };
  const deselectAll = () => setSelected({});
  const selectedCount = rows.filter((_, i) => selected[i]).length;

  const previewRow = rows[previewIdx] || {};
  const previewBody = fillTemplate(body, previewRow, headers);
  const previewSubject = fillTemplate(subject, previewRow, headers);
  const previewTo = emailCol ? (previewRow[emailCol] || '(empty)') : '(select email column)';

  const runAgent = async () => {
    if (!emailCol) { alert('Please select the email column in Step 2.'); return; }
    const targets = rows.map((row, i) => ({ row, i })).filter(({ i }) => selected[i]);
    if (targets.length === 0) { alert('Please select at least one company to generate drafts for.'); return; }
    setRunning(true); setLogs([]); setMetrics(null); setProgress(0);
    setRunStatus({ color: 'amber', msg: 'Starting agent…' });
    const addLog = (html, type) => setLogs(l => [...l, { html, type }]);
    let done = 0, skipped = 0;
    for (let n = 0; n < targets.length; n++) {
      const { row, i } = targets[n];
      const to = row[emailCol];
      setProgress(Math.round(((n + 1) / targets.length) * 100));
      setRunStatus({ color: 'amber', msg: `Processing ${n + 1} of ${targets.length}…` });
      if (!to || !to.includes('@')) { addLog(`[${i+1}] SKIPPED — no valid email`, 'skip'); skipped++; await sleep(80); continue; }
      addLog(`[${i+1}] AI personalizing for ${to}…`, 'info');
      let emailBody;
      try { emailBody = await aiPersonalise(row, emailCol, body, headers); addLog(`[${i+1}] ✓ AI draft ready`, 'ok'); }
      catch { emailBody = fillTemplate(body, row, headers); addLog(`[${i+1}] ⚠ AI unavailable — using template`, 'skip'); }
      openGmailDraft(to, subject, emailBody);
      addLog(`[${i+1}] ✓ Gmail draft opened → ${to}`, 'ok');
      if (includeDocs) {
        try {
          await downloadCompanyDocs(row[companyCol]);
          addLog(`[${i+1}] ✓ Invite & Proposal docs downloaded`, 'ok');
        } catch (err) {
          addLog(`[${i+1}] ⚠ Could not generate docs — ${err.message}`, 'skip');
        }
      }
      done++; await sleep(1000);
    }
    setMetrics({ total: targets.length, done, skipped });
    setRunStatus({ color: 'green', msg: `All done! ${done} drafts opened · ${skipped} skipped.` });
    setRunning(false);
  };

  const handleDownloadDocs = async (companyName) => {
    setDocsBusy(true); setDocsError('');
    try { await downloadCompanyDocs(companyName); }
    catch (err) { setDocsError(err.message); }
    setDocsBusy(false);
  };

  const downloadCSV = () => {
    const lines = [['to','cc','subject','body'].map(h => `"${h}"`).join(',')];
    rows.forEach(row => {
      const to = emailCol ? (row[emailCol] || '') : '';
      const b = fillTemplate(body, row, headers);
      lines.push([to, CC, subject, b].map(v => `"${(v||'').replace(/"/g,'""')}"`).join(','));
    });
    const a = Object.assign(document.createElement('a'), { href: URL.createObjectURL(new Blob([lines.join('\n')], { type: 'text/csv' })), download: 'aaruush_mail_drafts.csv' });
    a.click();
  };

  const stepOffset = inputMode === 'manual' ? -1 : 0;

  return (
    <div className="app">
      <nav className="nav">
        <div className="nav-brand">
          <div className="nav-logo"><svg viewBox="0 0 24 24" fill="white" width="18" height="18"><path d="M20 4H4C2.9 4 2 4.9 2 6v12c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm0 4l-8 5-8-5V6l8 5 8-5v2z"/></svg></div>
          <div><div className="nav-title">Aaruush Mail Agent</div><div className="nav-sub">Sponsorship Outreach · A'26</div></div>
        </div>
        <div style={{ display:'flex', alignItems:'center', gap:12 }}>
          <div className="nav-user"><span className="nav-user-dot" />{user.name}<span className="nav-role">{user.role || 'Committee Head'}</span></div>
          <button className="logout-btn" onClick={onLogout}>Sign out</button>
        </div>
      </nav>

      <div className="page">
        <div className="hero">
          <h1>Send <span>smarter</span> sponsorship emails,<br />not more of them.</h1>
          <p>Drop your sponsor list or add manually, let AI personalize every pitch, push them to Gmail drafts.</p>
        </div>

        {ADMINS.includes(user.name) && <AdminPanel chList={chList} setChList={setChList} />}

        {/* STEP 1 */}
        <Card num="1" title="Add Companies" active={!loaded}>
          <div className="mode-tabs">
            <button className={`mode-tab ${inputMode === 'upload' ? 'active' : ''}`} onClick={() => { setInputMode('upload'); setRows([]); setUploadStatus({ color:'gray', msg:'No file loaded yet' }); }}>📂 Upload Excel / CSV</button>
            <button className={`mode-tab ${inputMode === 'manual' ? 'active' : ''}`} onClick={() => { setInputMode('manual'); setRows([]); setHeaders(['COMPANY','EMAIL','NOTES']); setCompanyCol('COMPANY'); setEmailCol('EMAIL'); setUploadStatus({ color:'gray', msg:'' }); }}>✏️ Add Manually</button>
          </div>

          {inputMode === 'upload' && (
            <>
              <div className="drop-zone" onClick={() => document.getElementById('file-input').click()} onDragOver={e => e.preventDefault()} onDrop={onDrop}>
                {fileName ? (<><div className="drop-icon" style={{fontSize:28}}>✅</div><p style={{color:'var(--green)',fontWeight:600}}>{fileName}</p><span>{rows.length} rows · click to change</span></>) : (<><div className="drop-icon">📊</div><p>Click to upload or drag &amp; drop</p><span>.xlsx &nbsp;·&nbsp; .xls &nbsp;·&nbsp; .csv</span></>)}
              </div>
              <input type="file" id="file-input" accept=".xlsx,.xls,.csv" style={{display:'none'}} onChange={e => e.target.files[0] && processFile(e.target.files[0])} />
              <StatusBar {...uploadStatus} />
              {loaded && <div className="chip-wrap">{headers.map(h => <span key={h} className="chip">{`{{${h}}}`}</span>)}</div>}
            </>
          )}

          {inputMode === 'manual' && (
            <>
              <ManualEntry onAdd={addManualRow} />
              <StatusBar {...uploadStatus} />
              {rows.length > 0 && (
                <div className="manual-list">
                  <div className="manual-list-header">Added companies ({rows.length})</div>
                  {rows.map((r, i) => (
                    <div key={i} className="manual-row">
                      <div className="manual-row-info">
                        <span className="manual-company">{r.COMPANY}</span>
                        <span className="manual-email">{r.EMAIL}</span>
                        {r.NOTES && <span className="manual-notes">{r.NOTES}</span>}
                      </div>
                      <button className="btn-remove" onClick={() => removeRow(i)}>✕</button>
                    </div>
                  ))}
                </div>
              )}
            </>
          )}
        </Card>

        {/* STEP 2 — column map (upload only) */}
        {loaded && inputMode === 'upload' && (() => {
          const chCol = findCHColumn(headers);
          // Build display columns: first 5, but ensure CH column included for organizers
          let displayCols = headers.slice(0, 5);
          if (isOrganizer(user) && chCol && !displayCols.includes(chCol)) {
            displayCols = [...headers.slice(0, 4), chCol];
          }
          const hiddenCount = headers.length - displayCols.length;
          return (
          <Card num="2" title="Map Columns" active>
            <div className="col-grid">
              <div><label className="field-label">Company name column</label><select value={companyCol} onChange={e => setCompanyCol(e.target.value)}><option value="">-- select --</option>{headers.map(h => <option key={h} value={h}>{h}</option>)}</select></div>
              <div><label className="field-label">Recipient email column</label><select value={emailCol} onChange={e => setEmailCol(e.target.value)}><option value="">-- select --</option>{headers.map(h => <option key={h} value={h}>{h}</option>)}</select></div>
            </div>
            {isOrganizer(user) && chCol && <div className="organizer-note">👁️ Organizer view — showing all companies. The <strong>{chCol}</strong> column shows who each is assigned to.</div>}
            <div className="table-wrap table-scroll">
              <table>
                <thead><tr>{displayCols.map(h => <th key={h} className={h === chCol ? 'ch-col' : ''}>{h}</th>)}{hiddenCount > 0 && <th>…</th>}</tr></thead>
                <tbody>{rows.map((r,i) => <tr key={i}>{displayCols.map(h => <td key={h} title={r[h]} className={h === chCol ? 'ch-col' : ''}>{r[h]}</td>)}{hiddenCount > 0 && <td style={{color:'var(--muted)'}}>…</td>}</tr>)}</tbody>
              </table>
            </div>
            <div style={{ fontSize: 11, color: 'var(--muted)', marginTop: 6 }}>Showing all {rows.length} rows{hiddenCount > 0 ? ` · ${displayCols.length} of ${headers.length} columns` : ''}</div>
          </Card>
          );
        })()}

        {/* STEP 3 — template */}
        {loaded && (
          <Card num={inputMode === 'upload' ? '3' : '2'} title="Email Template" tag={{ color: 'green', label: '{{column_name}} placeholders' }}>
            <div className="inline-field"><span className="field-lbl">Subject</span><input type="text" value={subject} onChange={e => setSubject(e.target.value)} /></div>
            <textarea value={body} onChange={e => setBody(e.target.value)} rows={10} />
            <div className="cc-wrap">
              <span className="cc-label">Auto CC:</span>
              {['administrator@aaruush.org','secretary@aaruush.org','jointsecretary@aaruush.org','sponsorship@aaruush.org'].map(cc => <span key={cc} className="cc-tag">{cc}</span>)}
            </div>
          </Card>
        )}

        {/* STEP 4 — rich preview */}
        {loaded && (
          <Card num={inputMode === 'upload' ? '4' : '3'} title="Preview" tag={{ color: 'purple', label: '✦ AI personalized' }}>
            <div className="preview-nav">
              <span>Company {previewIdx + 1} of {rows.length}</span>
              <div className="nav-btns">
                <button className="btn-icon" onClick={() => setPreviewIdx(i => Math.max(0, i-1))}>←</button>
                <button className="btn-icon" onClick={() => setPreviewIdx(i => Math.min(rows.length-1, i+1))}>→</button>
              </div>
            </div>
            <EmailPreview user={user} body={previewBody} subject={previewSubject} to={previewTo} companyName={previewRow[companyCol]} />
            <div style={{fontSize:11,color:'var(--muted)',marginTop:8}}>
              💡 The logo & signature shown here matches your Gmail signature. The agent writes the body text above the divider.
            </div>
            <div style={{ display:'flex', gap:8, marginTop:12, flexWrap:'wrap' }}>
              <button className="btn-sec" disabled={docsBusy} onClick={() => handleDownloadDocs(previewRow[companyCol])}>
                {docsBusy ? '⏳ Generating…' : '↓ Invite + Proposal docs'}
              </button>
            </div>
            {docsError && <div className="manual-error" style={{ marginTop: 6 }}>{docsError}</div>}
          </Card>
        )}

        {/* STEP 5 — generate */}
        {loaded && (
          <Card num={inputMode === 'upload' ? '5' : '4'} title="Generate Gmail Drafts">
            <div className="select-bar">
              <span className="select-count">{selectedCount} of {rows.length} selected</span>
              <div className="select-actions">
                <button className="select-link" onClick={selectAll}>Select all</button>
                <button className="select-link" onClick={deselectAll}>Deselect all</button>
              </div>
            </div>
            <div className="select-list">
              {rows.map((r, i) => (
                <label key={i} className={`select-item ${selected[i] ? 'checked' : ''}`}>
                  <input type="checkbox" checked={!!selected[i]} onChange={() => toggleSelect(i)} />
                  <span className="select-company">{r[companyCol] || '(no name)'}</span>
                  <span className="select-email">{r[emailCol] || '(no email)'}</span>
                </label>
              ))}
            </div>
            <label className="select-item" style={{ marginBottom: 10 }}>
              <input type="checkbox" checked={includeDocs} onChange={e => setIncludeDocs(e.target.checked)} />
              <span className="select-company">Also download Invite &amp; Proposal docs per company</span>
              <span className="select-email" style={{fontSize:11}}>2 files each, lands in Downloads — drag into the Gmail draft before sending</span>
            </label>
            <div className="popup-warn"><strong>⚠ Allow popups from mail.google.com</strong> — Click the blocked popup icon in the address bar → <em>"Always allow"</em>. Then click Generate again.</div>
            <div className="btn-row">
              <button className="btn-primary" onClick={runAgent} disabled={running || !emailCol || selectedCount === 0}>{running ? '⏳ Running…' : `✦ Generate ${selectedCount} Draft${selectedCount === 1 ? '' : 's'} ↗`}</button>
              <button className="btn-sec" onClick={downloadCSV}>↓ Download CSV backup</button>
            </div>
            {runStatus.msg && (<><StatusBar {...runStatus} /><div className="prog-wrap"><div className="prog-bar" style={{width:progress+'%'}} /></div></>)}
            {logs.length > 0 && <div className="log-box">{logs.map((l,i) => <div key={i} className={`log-${l.type}`}>{l.html}</div>)}</div>}
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

export default function App() {
  const [user, setUser] = useState(null);
  const [chList, setChList] = useState(COMMITTEE_HEADS);
  if (!user) return <LoginPage onLogin={setUser} chList={chList} />;
  return <Agent user={user} onLogout={() => setUser(null)} chList={chList} setChList={setChList} />;
}
