import PizZip from 'pizzip';

/*
 * Client-side .docx generator — ported from a teammate's Next.js tool.
 * No backend involved: we fetch the official Word templates as raw bytes,
 * string-replace the locked placeholder names inside the docx XML, and
 * hand back a Blob the browser can download. PizZip runs fine in the browser.
 */

const TEMPLATE_BASE = '/templates';
const INVITE_TEMPLATE_URL = `${TEMPLATE_BASE}/official-invite.docx`;
const PROPOSAL_TEMPLATE_URL = `${TEMPLATE_BASE}/official-proposal.docx`;

const DEFAULT_PLACEHOLDERS = [
  '{{COMPANY_NAME}}', '{COMPANY_NAME}', '[COMPANY_NAME]', 'COMPANY_NAME',
  '{{company_name}}', '{company_name}', '[company_name]', 'company_name',
];

// The official templates were originally drafted for sample companies —
// these are the literal strings baked into the Word XML that get swapped
// for the real company name on each generation.
const INVITE_SOURCE_NAMES = ['Mother Dairy', 'MOTHER DAIRY'];
const PROPOSAL_SOURCE_NAMES = ['Oyela', 'OYELA', 'oyela'];

// Leftover signature-block fixes inherited from the original template file.
const PROPOSAL_EXTRA_REPLACEMENTS = {
  HarSiddharth: 'Siddharth',
  'HarSiddharth Agarwal': 'Siddharth',
  Yas: '',
  'Yas          Devansh Gupta': 'Devansh Gupta',
};

const templateCache = new Map();

function fetchTemplateBytes(url) {
  if (templateCache.has(url)) return templateCache.get(url);
  const promise = fetch(url)
    .then((res) => {
      if (!res.ok) throw new Error(`Could not load template ${url} (HTTP ${res.status})`);
      return res.arrayBuffer();
    })
    .catch((err) => { templateCache.delete(url); throw err; });
  templateCache.set(url, promise);
  return promise;
}

export function sanitizeCompanyName(value) {
  return (value || '')
    .trim()
    .replace(/[<>:"/\\|?*\x00-\x1F]/g, '')
    .replace(/\s+/g, ' ');
}

function escapeXml(value) {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function replacePlainXmlPlaceholders(zip, replacements) {
  Object.keys(zip.files)
    .filter((fileName) => fileName.startsWith('word/') && fileName.endsWith('.xml'))
    .forEach((fileName) => {
      const file = zip.file(fileName);
      const current = file && file.asText();
      if (!current) return;
      let next = current;
      for (const [search, replacement] of Object.entries(replacements)) {
        if (!search) continue;
        next = next.split(search).join(escapeXml(replacement));
      }
      if (next !== current) zip.file(fileName, next);
    });
}

async function renderDocxTemplate(templateUrl, companyName, sourceNames, extraReplacements) {
  const bytes = await fetchTemplateBytes(templateUrl);
  const zip = new PizZip(bytes);
  const replacements = Object.fromEntries(
    [...DEFAULT_PLACEHOLDERS, ...sourceNames].map((placeholder) => [placeholder, companyName])
  );
  replacePlainXmlPlaceholders(zip, { ...replacements, ...(extraReplacements || {}) });
  const out = zip.generate({ type: 'blob', compression: 'DEFLATE' });
  return new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
}

export function downloadBlob(blob, filename) {
  const a = Object.assign(document.createElement('a'), {
    href: URL.createObjectURL(blob),
    download: filename,
  });
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 1000);
}

// Generates the Invite and/or Proposal doc for one company.
// `options.invite` / `options.proposal` (both default true) control which
// one(s) actually get rendered — the other comes back null.
// Returns { inviteBlob, inviteFileName, proposalBlob, proposalFileName }.
export async function generateCompanyDocs(rawCompanyName, options = {}) {
  const { invite = true, proposal = true } = options;
  const companyName = sanitizeCompanyName(rawCompanyName) || 'Sponsor';
  const [inviteBlob, proposalBlob] = await Promise.all([
    invite ? renderDocxTemplate(INVITE_TEMPLATE_URL, companyName, INVITE_SOURCE_NAMES) : Promise.resolve(null),
    proposal ? renderDocxTemplate(PROPOSAL_TEMPLATE_URL, companyName, PROPOSAL_SOURCE_NAMES, PROPOSAL_EXTRA_REPLACEMENTS) : Promise.resolve(null),
  ]);
  return {
    inviteBlob,
    inviteFileName: `AARUUSH X ${companyName} Invite.docx`,
    proposalBlob,
    proposalFileName: `AARUUSH X ${companyName} Proposal A'26.docx`,
  };
}

export async function downloadCompanyDocs(rawCompanyName, options = {}) {
  const { invite = true, proposal = true } = options;
  if (!invite && !proposal) return;
  const { inviteBlob, inviteFileName, proposalBlob, proposalFileName } = await generateCompanyDocs(rawCompanyName, { invite, proposal });
  if (inviteBlob) downloadBlob(inviteBlob, inviteFileName);
  // Stagger slightly so the browser doesn't drop the second download.
  if (inviteBlob && proposalBlob) await new Promise((r) => setTimeout(r, 400));
  if (proposalBlob) downloadBlob(proposalBlob, proposalFileName);
}
