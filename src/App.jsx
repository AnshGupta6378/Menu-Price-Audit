import { useState, useMemo, useRef, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import { SAMPLE_INSTORE, SAMPLE_DELIVERY } from "./samples.js";
import {
  Upload,
  FileSpreadsheet,
  FileText,
  Image as ImageIcon,
  X,
  Loader2,
  ArrowUpRight,
  ArrowDownRight,
  Link2,
  Trash2,
  Download,
  RotateCcw,
  AlertTriangle,
  Check,
  Repeat,
  Sparkles,
} from "lucide-react";

/* ------------------------------------------------------------------ */
/*  Tokens                                                             */
/* ------------------------------------------------------------------ */
const C = {
  paper: "#E7EBE5",
  card: "#F7F9F5",
  white: "#FFFFFF",
  ink: "#10261E",
  inkSoft: "#4B5F55",
  inkFaint: "#8A9A90",
  rule: "#C7D1C6",
  up: "#A61352",
  upSoft: "#F5DCE8",
  down: "#0E7C7B",
  downSoft: "#D7EDEC",
  warn: "#A9711A",
  warnSoft: "#F6E7CC",
  ai: "#3D4E9E",
  aiSoft: "#DFE3F5",
};
const SERIF = 'Georgia, "Iowan Old Style", "Times New Roman", serif';
const SANS =
  '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif';
const MONO = 'ui-monospace, SFMono-Regular, Menlo, Consolas, "Courier New", monospace';

/* ------------------------------------------------------------------ */
/*  Menu vocabulary                                                    */
/* ------------------------------------------------------------------ */
const STOP = new Set([
  "the", "with", "a", "an", "of", "in", "and", "our", "special", "specials",
  "new", "chefs", "chef", "house", "signature", "style", "served", "fresh",
  "homemade", "classic", "authentic", "famous",
]);

/* Same dish, different spelling. Applied to both menus so it stays symmetric. */
const WORD = {
  veg: "vegetable", vegs: "vegetable", vegetarian: "vegetable", vegg: "vegetable",
  nonveg: "chicken", chkn: "chicken", chiken: "chicken", chikan: "chicken",
  chickn: "chicken", murg: "chicken", murgh: "chicken", murgha: "chicken",
  pnr: "paneer", paneeer: "paneer", panner: "paneer", panir: "paneer",
  btr: "butter", makhani: "butter", makhanwala: "butter", mahkani: "butter",
  daal: "dal", dhal: "dal", dail: "dal", lentil: "dal", lentils: "dal",
  biriyani: "biryani", briyani: "biryani", biriyani2: "biryani", biryan: "biryani",
  chilly: "chili", chilli: "chili", chillie: "chili", chillies: "chili",
  schezuan: "schezwan", szechuan: "schezwan", sichuan: "schezwan", schezwaan: "schezwan",
  chapati: "roti", chapathi: "roti", phulka: "roti", tandooriroti: "roti",
  dahi: "curd", yoghurt: "curd", yogurt: "curd",
  aloo: "potato", alu: "potato", batata: "potato",
  gobi: "cauliflower", gobhi: "cauliflower",
  bhindi: "okra", ladyfinger: "okra", ladysfinger: "okra",
  baingan: "eggplant", brinjal: "eggplant", bengan: "eggplant",
  kaju: "cashew", kaaju: "cashew",
  jeera: "cumin", methi: "fenugreek", palak: "spinach", saag: "spinach",
  matar: "peas", mutter: "peas", mattar: "peas",
  anda: "egg", eggs: "egg",
  mutton: "lamb", gosht: "lamb",
  prawns: "prawn", jhinga: "prawn",
  fries: "frenchfries", finger: "frenchfries",
  momos: "momo", dumpling: "momo", dumplings: "momo",
  milkshake: "shake", thickshake: "shake",
  softdrink: "beverage", colddrink: "beverage", aerated: "beverage",
  mocktails: "mocktail", juices: "juice",
  wrap: "roll", wraps: "roll", rolls: "roll", frankie: "roll",
  starter: "starters", appetizer: "starters", appetiser: "starters",
  gravy: "curry", masaledar: "masala",
  kadai: "kadhai", karahi: "kadhai", kadahi: "kadhai",
  pcs: "", pc: "", piece: "", pieces: "", nos: "",
};

const SIZE_CANON = {
  half: "half", adha: "half", "1/2": "half",
  quarter: "qtr", "1/4": "qtr",
  full: "std", regular: "std", normal: "std", standard: "std", single: "std", plate: "std",
  large: "lg", jumbo: "lg", family: "lg", king: "lg", xl: "lg", sharing: "lg", double: "lg",
  small: "sm", mini: "sm", bite: "sm", "6inch": "sm", "6 inch": "sm",
};

function baseNorm(s) {
  return String(s || "")
    .toLowerCase()
    .replace(/&/g, " and ")
    .replace(/[^a-z0-9./\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

/* Pull portion / quantity out of the name so it can be compared separately. */
function sizeSignature(raw) {
  const s = " " + baseNorm(raw).replace(/\./g, " ") + " ";
  const sig = [];
  let m = s.match(/(\d+(?:\.\d+)?)\s*(ml|ltr|litre|liter|l)\b/);
  if (m) {
    let v = parseFloat(m[1]);
    if (/^(ltr|litre|liter|l)$/.test(m[2])) v *= 1000;
    sig.push("v" + v);
  }
  m = s.match(/(\d+(?:\.\d+)?)\s*(gm|gms|g|gram|grams|kg)\b/);
  if (m) {
    let v = parseFloat(m[1]);
    if (m[2] === "kg") v *= 1000;
    sig.push("w" + v);
  }
  m = s.match(/(\d+)\s*(pc|pcs|piece|pieces|nos)\b/);
  if (m) sig.push("p" + m[1]);
  m = s.match(/\b(half|full|quarter|regular|normal|standard|large|jumbo|family|king|small|mini|single|double|sharing)\b/);
  if (m) sig.push("s" + (SIZE_CANON[m[1]] || m[1]));
  return sig.length ? sig.sort().join("|") : null;
}

function stripSize(s) {
  return baseNorm(s)
    .replace(/\(?\s*\d+(\.\d+)?\s*(ml|ltr|litre|liter|gm|gms|g|gram|grams|kg|pc|pcs|piece|pieces|nos)\b\s*\)?/g, " ")
    .replace(/\b(half|full|quarter|regular|normal|standard|large|jumbo|family|king|size|small|mini|single|double|sharing|portion)\b/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function canonTokens(s) {
  return stripSize(s)
    .replace(/\//g, " ")
    .split(" ")
    .map((t) => (WORD[t] !== undefined ? WORD[t] : t))
    .filter((t) => t && !STOP.has(t) && t.length > 1);
}

function canonName(s) {
  return canonTokens(s).sort().join(" ");
}

function lev(a, b) {
  if (a === b) return 0;
  const m = a.length, n = b.length;
  if (!m) return n;
  if (!n) return m;
  let prev = Array.from({ length: n + 1 }, (_, i) => i);
  const cur = new Array(n + 1);
  for (let i = 1; i <= m; i++) {
    cur[0] = i;
    for (let j = 1; j <= n; j++) {
      cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    }
    prev = cur.slice();
  }
  return prev[n];
}

/* Returns { score, sizeClash, sizeOneSided, catClash } */
function compare(a, b) {
  const ca = canonName(a.name);
  const cb = canonName(b.name);
  const sa = sizeSignature(a.name);
  const sb = sizeSignature(b.name);
  const sizeClash = !!(sa && sb && sa !== sb);
  const sizeOneSided = !!((sa && !sb) || (!sa && sb));
  const catClash = !!(a.category && b.category && canonName(a.category) !== canonName(b.category));

  if (!ca || !cb) return { score: 0, sizeClash, sizeOneSided, catClash };

  const ta = new Set(canonTokens(a.name));
  const tb = new Set(canonTokens(b.name));
  let inter = 0;
  ta.forEach((t) => tb.has(t) && inter++);
  const jac = inter / (new Set([...ta, ...tb]).size || 1);
  const cover = inter / Math.max(1, Math.min(ta.size, tb.size)); // subset credit
  const ratio = 1 - lev(ca, cb) / Math.max(ca.length, cb.length);

  let score = 0.42 * jac + 0.28 * cover + 0.3 * ratio;
  if (ca === cb) score = 1;

  if (sizeClash) score *= 0.28;
  else if (sizeOneSided) score *= 0.9;
  if (catClash) score *= 0.93;

  return { score: Math.max(0, Math.min(1, score)), sizeClash, sizeOneSided, catClash };
}

function toNumber(v) {
  if (typeof v === "number" && isFinite(v)) return v;
  const s = String(v ?? "").replace(/[₹$,\s]/g, "").replace(/rs\.?/i, "");
  const m = s.match(/-?\d+(\.\d+)?/);
  if (!m) return null;
  const n = parseFloat(m[0]);
  return isFinite(n) ? n : null;
}

const inr = (n) =>
  n == null ? "-" : "₹" + Number(n).toLocaleString("en-IN", { maximumFractionDigits: 2 });
const pct = (n) => (n == null ? "-" : (n > 0 ? "+" : "") + n.toFixed(1) + "%");

/* ------------------------------------------------------------------ */
/*  File reading                                                       */
/* ------------------------------------------------------------------ */
const readAs = (file, how) =>
  new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(how === "b64" ? String(r.result).split(",")[1] : r.result);
    r.onerror = () => rej(new Error("Could not read " + file.name));
    if (how === "b64") r.readAsDataURL(file);
    else if (how === "buf") r.readAsArrayBuffer(file);
    else r.readAsText(file);
  });

function itemsFromGrid(grid) {
  if (!grid || !grid.length) return [];
  const NAME_RE = /(item|name|dish|product|menu|particular|description)/i;
  const PRICE_RE = /(price|rate|amount|mrp|cost|value|selling)/i;
  const CAT_RE = /(category|section|course|group|type|cuisine)/i;

  let hr = -1, nc = -1, pc = -1, cc = -1;
  for (let r = 0; r < Math.min(grid.length, 12); r++) {
    const row = grid[r] || [];
    let n = -1, p = -1, c = -1;
    row.forEach((cell, i) => {
      const s = String(cell ?? "");
      if (n === -1 && NAME_RE.test(s) && !CAT_RE.test(s)) n = i;
      if (p === -1 && PRICE_RE.test(s)) p = i;
      if (c === -1 && CAT_RE.test(s)) c = i;
    });
    if (n !== -1 && p !== -1) { hr = r; nc = n; pc = p; cc = c; break; }
  }

  const out = [];
  if (hr !== -1) {
    for (let r = hr + 1; r < grid.length; r++) {
      const row = grid[r] || [];
      const name = String(row[nc] ?? "").trim();
      const price = toNumber(row[pc]);
      if (name && /[a-z]/i.test(name) && price != null && price > 0)
        out.push({ name, price, category: cc !== -1 ? String(row[cc] ?? "").trim() : "" });
    }
    if (out.length >= 2) return out;
  }

  const alt = [];
  for (const row of grid) {
    if (!row) continue;
    let name = null, price = null;
    for (const cell of row) {
      const s = String(cell ?? "").trim();
      if (name == null && /[a-zA-Z]{2,}/.test(s) && toNumber(s) == null) name = s;
    }
    for (let i = row.length - 1; i >= 0; i--) {
      const n = toNumber(row[i]);
      if (n != null && n > 0) { price = n; break; }
    }
    if (name && price != null) alt.push({ name, price, category: "" });
  }
  return alt.length >= 2 ? alt : [];
}

function parseItemsJSON(raw) {
  const text = String(raw || "").replace(/```json/gi, "").replace(/```/g, "").trim();
  const shape = (arr) =>
    arr
      .map((e) => {
        if (Array.isArray(e))
          return { name: String(e[0] ?? ""), price: toNumber(e[1]), category: String(e[2] ?? "") };
        if (e && typeof e === "object")
          return {
            name: String(e.name ?? e.item ?? ""),
            price: toNumber(e.price ?? e.rate ?? e.amount),
            category: String(e.category ?? e.section ?? ""),
          };
        return null;
      })
      .filter((e) => e && e.name.trim() && e.price != null && e.price > 0);

  try {
    const j = JSON.parse(text);
    if (Array.isArray(j)) return shape(j);
    if (j && Array.isArray(j.items)) return shape(j.items);
  } catch (e) { /* salvage below */ }

  const out = [];
  const re = /\[\s*"((?:[^"\\]|\\.)*)"\s*,\s*(-?\d+(?:\.\d+)?)\s*(?:,\s*"((?:[^"\\]|\\.)*)"\s*)?\]/g;
  let m;
  while ((m = re.exec(text)) !== null) {
    const price = parseFloat(m[2]);
    if (m[1].trim() && price > 0)
      out.push({ name: m[1].trim(), price, category: (m[3] || "").trim() });
  }
  return out;
}

/* Bring-your-own key. Kept in this browser only, sent straight to Anthropic. */
const KEY_SLOT = "menu-price-audit:key";
const getKey = () => { try { return localStorage.getItem(KEY_SLOT) || ""; } catch (e) { return ""; } };
const saveKey = (k) => { try { k ? localStorage.setItem(KEY_SLOT, k) : localStorage.removeItem(KEY_SLOT); } catch (e) { /* storage blocked */ } };

async function callClaude(content) {
  const key = getKey();
  if (!key) throw new Error("Add an Anthropic API key to read photos and PDFs");
  const r = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": key,
      "anthropic-version": "2023-06-01",
      "anthropic-dangerous-direct-browser-access": "true",
    },
    body: JSON.stringify({
      model: "claude-sonnet-4-6",
      max_tokens: 4000,
      messages: [{ role: "user", content }],
    }),
  });
  if (!r.ok) throw new Error("Reader failed (" + r.status + ")");
  const data = await r.json();
  return (data.content || []).map((b) => (b.type === "text" ? b.text : "")).join("\n");
}

const EXTRACT_PROMPT = `You are reading a restaurant menu. List every item with its price and the section it sits under.

Rules:
- Reply with ONLY a JSON array. No preamble, no markdown.
- Format: [["Item name", price, "Section"]]
- Price is a plain number. No currency symbol, no commas.
- If an item has several portions or sizes, output one entry per size with the size in the name: ["Paneer Butter Masala (Half)", 180, "Main Course"].
- Copy the item name exactly as printed, including its spelling. Do not translate, expand or tidy it.
- "Section" is the heading the item appears under, e.g. Starters, Main Course, Breads, Beverages. Use "" if there is none.
- Skip headings, taxes, notes and anything without a price.`;

async function extractItems(file) {
  const name = file.name.toLowerCase();
  const type = file.type || "";

  if (/\.(csv|tsv)$/.test(name) || type === "text/csv") {
    const text = await readAs(file, "text");
    const items = itemsFromGrid(Papa.parse(text, { skipEmptyLines: true }).data);
    if (items.length) return items;
    return parseItemsJSON(await callClaude([{ type: "text", text: text.slice(0, 6000) }, { type: "text", text: EXTRACT_PROMPT }]));
  }

  if (/\.(xlsx|xlsm|xls)$/.test(name)) {
    const wb = XLSX.read(await readAs(file, "buf"), { type: "array" });
    let all = [];
    for (const s of wb.SheetNames)
      all = all.concat(itemsFromGrid(XLSX.utils.sheet_to_json(wb.Sheets[s], { header: 1, defval: "" })));
    if (all.length) return all;
    const g = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: "" });
    const txt = g.map((r) => r.join(" | ")).join("\n").slice(0, 6000);
    return parseItemsJSON(await callClaude([{ type: "text", text: txt }, { type: "text", text: EXTRACT_PROMPT }]));
  }

  if (name.endsWith(".pdf") || type === "application/pdf") {
    const b64 = await readAs(file, "b64");
    return parseItemsJSON(
      await callClaude([
        { type: "document", source: { type: "base64", media_type: "application/pdf", data: b64 } },
        { type: "text", text: EXTRACT_PROMPT },
      ])
    );
  }

  if (type.startsWith("image/")) {
    const b64 = await readAs(file, "b64");
    return parseItemsJSON(
      await callClaude([
        { type: "image", source: { type: "base64", media_type: type, data: b64 } },
        { type: "text", text: EXTRACT_PROMPT },
      ])
    );
  }

  const text = await readAs(file, "text");
  return parseItemsJSON(await callClaude([{ type: "text", text: String(text).slice(0, 6000) }, { type: "text", text: EXTRACT_PROMPT }]));
}

/* ------------------------------------------------------------------ */
/*  Stage 3: semantic reconciliation                                   */
/* ------------------------------------------------------------------ */
const RECONCILE_PROMPT = `Two menus from the SAME restaurant: list A is the in-store menu, list B is the delivery listing. The same dish is often written differently on each: different spelling, a shorter or longer name, a regional name, or a different section heading.

Pair each A item with the B item that is the same dish.

Pair when:
- The names differ but the dish is the same (Paneer Butter Masala = Paneer Makhani, Veg Manchurian Dry = Vegetable Manchurian, Chicken Biryani = Murgh Dum Biryani).
- One side abbreviates, misspells or adds words (Chicken Tikka Masala = Chkn Tikka Masala, Dal Tadka = Yellow Dal Tadka).

Do NOT pair when:
- The portion differs (Half vs Full, 250ml vs 500ml, 4pc vs 8pc). Leave those unpaired.
- The dish is only similar, not the same (Paneer Tikka vs Paneer Tikka Masala, Veg Fried Rice vs Egg Fried Rice, Butter Naan vs Garlic Naan). These are different items.
- The preparation differs in a way that changes the dish (Dry vs Gravy, Tandoori vs Curry).

Reply with ONLY a JSON array, no preamble:
[[A_id, B_id, confidence, "reason"]]
confidence is 0 to 1. reason is 4 words maximum. Each id used once. Skip anything you cannot pair confidently.`;

async function reconcile(aList, bList) {
  const fmt = (l) =>
    l.map((x) => `${x.id}|${x.name}|${x.price}${x.category ? "|" + x.category : ""}`).join("\n");
  const text =
    RECONCILE_PROMPT +
    "\n\nList A (id|name|price|section):\n" + fmt(aList) +
    "\n\nList B (id|name|price|section):\n" + fmt(bList);
  const out = await callClaude([{ type: "text", text }]);
  const clean = String(out).replace(/```json/gi, "").replace(/```/g, "").trim();
  const pairs = [];
  const re = /\[\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d*\.?\d+)\s*(?:,\s*"((?:[^"\\]|\\.)*)"\s*)?\]/g;
  let m;
  while ((m = re.exec(clean)) !== null) {
    pairs.push({
      a: parseInt(m[1], 10),
      b: parseInt(m[2], 10),
      conf: Math.min(1, parseFloat(m[3])),
      reason: (m[4] || "").trim(),
    });
  }
  return pairs;
}

/* ------------------------------------------------------------------ */
/*  Sample menus (made up, for the demo)                               */
/* ------------------------------------------------------------------ */

/* ------------------------------------------------------------------ */
/*  UI atoms                                                           */
/* ------------------------------------------------------------------ */
function Leader() {
  return <div className="flex-1 mx-2 mb-1 self-end" style={{ borderBottom: `1px dotted ${C.rule}` }} />;
}

function Chip({ children, tone = "neutral", title }) {
  const map = {
    neutral: { bg: "#DFE6DD", fg: C.inkSoft },
    up: { bg: C.upSoft, fg: C.up },
    down: { bg: C.downSoft, fg: C.down },
    warn: { bg: C.warnSoft, fg: C.warn },
    ai: { bg: C.aiSoft, fg: C.ai },
  };
  const t = map[tone] || map.neutral;
  return (
    <span
      title={title}
      className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded shrink-0"
      style={{ background: t.bg, color: t.fg, fontFamily: SANS, fontSize: 10.5, letterSpacing: "0.02em" }}
    >
      {children}
    </span>
  );
}

function fileIcon(f) {
  const n = f.name.toLowerCase();
  if ((f.type || "").startsWith("image/")) return <ImageIcon size={13} />;
  if (/\.(xlsx|xlsm|xls|csv|tsv)$/.test(n)) return <FileSpreadsheet size={13} />;
  return <FileText size={13} />;
}

function UploadCard({ label, hint, files, onAdd, onRemove, busy }) {
  const ref = useRef(null);
  const [over, setOver] = useState(false);
  const done = files.filter((f) => f.status === "done").length;
  const count = files.reduce((a, f) => a + (f.items?.length || 0), 0);

  return (
    <div
      className="rounded-lg p-4"
      style={{ background: C.card, border: `1px solid ${over ? C.ink : C.rule}`, transition: "border-color 120ms" }}
      onDragOver={(e) => { e.preventDefault(); setOver(true); }}
      onDragLeave={() => setOver(false)}
      onDrop={(e) => { e.preventDefault(); setOver(false); onAdd(Array.from(e.dataTransfer.files || [])); }}
    >
      <div className="flex items-baseline justify-between mb-1">
        <h3 style={{ fontFamily: SERIF, fontSize: 17, color: C.ink }}>{label}</h3>
        {count > 0 && (
          <span style={{ fontFamily: MONO, fontSize: 11, color: C.inkSoft }}>
            {count} items · {done}/{files.length} files
          </span>
        )}
      </div>
      <p style={{ fontFamily: SANS, fontSize: 12, color: C.inkFaint }} className="mb-3">{hint}</p>

      <input
        ref={ref}
        type="file"
        multiple
        accept="image/*,.pdf,.csv,.tsv,.xlsx,.xls,.xlsm"
        className="hidden"
        onChange={(e) => { onAdd(Array.from(e.target.files || [])); e.target.value = ""; }}
      />

      <button
        onClick={() => ref.current?.click()}
        disabled={busy}
        className="w-full flex items-center justify-center gap-2 py-2.5 rounded"
        style={{
          background: busy ? "#DFE6DD" : C.ink,
          color: busy ? C.inkSoft : "#F2F5F0",
          fontFamily: SANS, fontSize: 13, cursor: busy ? "default" : "pointer",
        }}
      >
        {busy ? <Loader2 size={14} className="animate-spin" /> : <Upload size={14} />}
        {busy ? "Reading menu…" : "Choose files"}
      </button>

      {files.length > 0 && (
        <ul className="mt-3 space-y-1.5">
          {files.map((f) => (
            <li key={f.id} className="flex items-center gap-2 px-2 py-1.5 rounded"
              style={{ background: C.white, fontFamily: SANS, fontSize: 12, color: C.inkSoft }}>
              <span style={{ color: C.inkFaint }}>{fileIcon(f.file)}</span>
              <span className="truncate flex-1">{f.file.name}</span>
              {f.status === "reading" && <Loader2 size={13} className="animate-spin" style={{ color: C.inkFaint }} />}
              {f.status === "done" && <span style={{ fontFamily: MONO, fontSize: 11, color: C.down }}>{f.items.length}</span>}
              {f.status === "error" && <span title={f.error} style={{ color: C.up }}><AlertTriangle size={13} /></span>}
              <button onClick={() => onRemove(f.id)} style={{ color: C.inkFaint }}><X size={13} /></button>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Main                                                               */
/* ------------------------------------------------------------------ */
const STRONG = 0.72;
const WEAK = 0.46;

export default function MenuPriceDelta() {
  const [partner, setPartner] = useState("");
  const [offFiles, setOffFiles] = useState([]);
  const [onFiles, setOnFiles] = useState([]);
  const [busy, setBusy] = useState({ offline: false, online: false });

  const [edits, setEdits] = useState({});
  const [dropped, setDropped] = useState({});
  const [manual, setManual] = useState({});
  const [unlinked, setUnlinked] = useState({});
  const [aiLinks, setAiLinks] = useState({});
  const [apiKey, setApiKey] = useState(getKey);
  const [aiOn, setAiOn] = useState(() => !!getKey());
  const [aiState, setAiState] = useState("idle");
  const [method, setMethod] = useState("blended");
  const [tab, setTab] = useState("matched");
  const [q, setQ] = useState("");
  const [sort, setSort] = useState("hike");
  const [swapping, setSwapping] = useState(null);
  const attempted = useRef("");

  const collect = (files) => {
    const seen = new Set();
    const out = [];
    files.forEach((f) =>
      (f.items || []).forEach((it) => {
        const k = canonName(it.name) + "#" + (sizeSignature(it.name) || "");
        if (!k.trim() || seen.has(k)) return;
        seen.add(k);
        out.push(it);
      })
    );
    return out;
  };

  const offline = useMemo(() => collect(offFiles), [offFiles]);
  const online = useMemo(() => collect(onFiles), [onFiles]);

  const addFiles = useCallback(async (side, list) => {
    if (!list.length) return;
    const setter = side === "offline" ? setOffFiles : setOnFiles;
    const staged = list.map((file) => ({
      id: Math.random().toString(36).slice(2), file, status: "reading", items: [], error: null,
    }));
    setter((p) => [...p, ...staged]);
    setBusy((b) => ({ ...b, [side]: true }));
    for (const s of staged) {
      try {
        const items = await extractItems(s.file);
        setter((p) => p.map((f) => f.id === s.id
          ? { ...f, status: items.length ? "done" : "error", items, error: items.length ? null : "No priced items found here" }
          : f));
      } catch (err) {
        setter((p) => p.map((f) => (f.id === s.id ? { ...f, status: "error", error: err.message } : f)));
      }
    }
    setBusy((b) => ({ ...b, [side]: false }));
  }, []);

  /* ---------- Stage 1+2: exact + gated fuzzy ---------- */
  const scored = useMemo(() => {
    const all = [];
    offline.forEach((o, i) =>
      online.forEach((n, j) => {
        const c = compare(o, n);
        if (c.score >= WEAK) all.push({ i, j, ...c });
      })
    );
    return all.sort((a, b) => b.score - a.score);
  }, [offline, online]);

  const strongLinks = useMemo(() => {
    const uo = new Set(), un = new Set(), out = new Map();
    for (const p of scored) {
      if (p.score < STRONG || p.sizeClash) continue;
      if (uo.has(p.i) || un.has(p.j)) continue;
      uo.add(p.i); un.add(p.j);
      out.set(p.i, { j: p.j, score: p.score, src: p.score === 1 ? "exact" : "name", ...p });
    }
    return out;
  }, [scored]);

  /* leftovers after stage 2: what the semantic pass gets to see */
  const leftovers = useMemo(() => {
    const takenN = new Set([...strongLinks.values()].map((v) => v.j));
    const a = offline.map((it, i) => ({ ...it, id: i })).filter((x) => !strongLinks.has(x.id));
    const b = online.map((it, j) => ({ ...it, id: j })).filter((x) => !takenN.has(x.id));
    return { a, b };
  }, [offline, online, strongLinks]);

  const runReconcile = useCallback(async (force) => {
    if (!getKey()) { setAiState("nokey"); return; }
    const sig = leftovers.a.map((x) => x.name).join("|") + "::" + leftovers.b.map((x) => x.name).join("|");
    if (!force && attempted.current === sig) return;
    if (!leftovers.a.length || !leftovers.b.length) return;
    attempted.current = sig;
    setAiState("running");
    try {
      const found = {};
      const takenB = new Set();
      for (let k = 0; k < leftovers.a.length; k += 18) {
        const chunk = leftovers.a.slice(k, k + 18);
        const pool = leftovers.b.filter((x) => !takenB.has(x.id)).slice(0, 70);
        if (!pool.length) break;
        const pairs = await reconcile(chunk, pool);
        pairs.forEach((p) => {
          if (p.conf < 0.5) return;
          if (takenB.has(p.b)) return;
          if (!offline[p.a] || !online[p.b]) return;
          if (!chunk.some((x) => x.id === p.a)) return;
          takenB.add(p.b);
          found[p.a] = { j: p.b, conf: p.conf, reason: p.reason };
        });
      }
      setAiLinks(found);
      setAiState("done");
    } catch (e) {
      setAiState("error");
    }
  }, [leftovers, offline, online]);

  useEffect(() => {
    if (!aiOn) return;
    if (busy.offline || busy.online) return;
    if (!offline.length || !online.length) return;
    runReconcile(false);
  }, [aiOn, busy, offline, online, runReconcile]);

  /* ---------- assemble ---------- */
  const { rows, onlyOffline, onlyOnline } = useMemo(() => {
    const link = new Map();
    const takenN = new Set();
    const claim = (i, v) => {
      if (link.has(i) || takenN.has(v.j)) return;
      link.set(i, v); takenN.add(v.j);
    };

    // 1. manual wins
    Object.entries(manual).forEach(([k, j]) => {
      const i = parseInt(k, 10);
      if (j == null || !offline[i] || !online[j]) return;
      const c = compare(offline[i], online[j]);
      claim(i, { j, score: c.score, src: "manual", ...c });
    });
    // 2. strong name matches
    strongLinks.forEach((v, i) => { if (!unlinked[i]) claim(i, v); });
    // 3. semantic matches
    Object.entries(aiLinks).forEach(([k, v]) => {
      const i = parseInt(k, 10);
      if (unlinked[i] || !offline[i] || !online[v.j]) return;
      const c = compare(offline[i], online[v.j]);
      claim(i, { j: v.j, score: v.conf, src: "ai", reason: v.reason, sizeClash: c.sizeClash, sizeOneSided: c.sizeOneSided, catClash: c.catClash });
    });
    // 4. weak name matches, flagged
    for (const p of scored) {
      if (p.score >= STRONG || unlinked[p.i]) continue;
      claim(p.i, { j: p.j, score: p.score, src: "weak", ...p });
    }

    const out = [];
    link.forEach((v, i) => {
      const key = String(i);
      if (dropped[key]) return;
      const e = edits[key] || {};
      const op = e.offline ?? offline[i].price;
      const np = e.online ?? online[v.j].price;
      const hike = op > 0 ? ((np - op) / op) * 100 : null;
      const suspicious = hike != null && Math.abs(hike) > 60;
      out.push({
        key, i, j: v.j,
        offName: offline[i].name, onName: online[v.j].name,
        offCat: offline[i].category || "", onCat: online[v.j].category || "",
        offPrice: op, onPrice: np, delta: np - op, hike,
        weight: e.weight ?? 1,
        score: v.score, src: v.src, reason: v.reason,
        sizeClash: v.sizeClash, sizeOneSided: v.sizeOneSided, catClash: v.catClash,
        suspicious,
        needsEye:
          v.src === "weak" ||
          (v.src === "ai" && v.score < 0.75) ||
          v.sizeClash ||
          (v.sizeOneSided && v.src !== "manual") ||
          suspicious,
      });
    });

    const oo = offline.map((it, i) => ({ ...it, i })).filter((x) => !link.has(x.i) || dropped[String(x.i)]);
    const on = online.map((it, j) => ({ ...it, j })).filter((x) => !takenN.has(x.j));
    return { rows: out, onlyOffline: oo, onlyOnline: on };
  }, [offline, online, manual, unlinked, aiLinks, strongLinks, scored, edits, dropped]);

  const review = rows.filter((r) => r.needsEye);
  const clean = rows.filter((r) => !r.needsEye);

  const stats = useMemo(() => {
    const valid = rows.filter((r) => r.offPrice > 0 && r.onPrice > 0);
    if (!valid.length) return null;
    const wOff = valid.reduce((a, r) => a + r.offPrice * (r.weight || 0), 0);
    const wOn = valid.reduce((a, r) => a + r.onPrice * (r.weight || 0), 0);
    const sorted = [...valid].sort((a, b) => a.hike - b.hike);
    return {
      blended: wOff > 0 ? ((wOn - wOff) / wOff) * 100 : null,
      simple: valid.reduce((a, r) => a + r.hike, 0) / valid.length,
      median: sorted[Math.floor(sorted.length / 2)].hike,
      weighted: valid.some((r) => (r.weight || 1) !== 1),
      wOff, wOn,
      hiked: valid.filter((r) => r.hike > 0.5).length,
      same: valid.filter((r) => Math.abs(r.hike) <= 0.5).length,
      cut: valid.filter((r) => r.hike < -0.5).length,
      top: [...valid].sort((a, b) => b.hike - a.hike).slice(0, 3),
      n: valid.length,
    };
  }, [rows]);

  const headline = stats ? (method === "blended" ? stats.blended : stats.simple) : null;

  const view = useMemo(() => {
    let list = tab === "review" ? review : clean;
    if (q.trim()) {
      const n = baseNorm(q);
      list = list.filter((r) => baseNorm(r.offName).includes(n) || baseNorm(r.onName).includes(n));
    }
    const s = [...list];
    if (sort === "hike") s.sort((a, b) => (b.hike ?? -999) - (a.hike ?? -999));
    if (sort === "drop") s.sort((a, b) => (a.hike ?? 999) - (b.hike ?? 999));
    if (sort === "name") s.sort((a, b) => a.offName.localeCompare(b.offName));
    if (sort === "value") s.sort((a, b) => b.onPrice - a.onPrice);
    if (sort === "confidence") s.sort((a, b) => a.score - b.score);
    return s;
  }, [clean, review, tab, q, sort]);

  const candidates = (i) =>
    online
      .map((o, j) => ({ o, j, c: compare(offline[i], o) }))
      .sort((a, b) => b.c.score - a.c.score)
      .slice(0, 8);

  const setEdit = (key, field, val) => setEdits((e) => ({ ...e, [key]: { ...(e[key] || {}), [field]: val } }));

  const relink = (i, j) => {
    setManual((m) => ({ ...m, [i]: j }));
    setUnlinked((u) => { const c = { ...u }; delete c[i]; return c; });
    setDropped((d) => { const c = { ...d }; delete c[String(i)]; return c; });
    setSwapping(null);
  };
  const breakLink = (i) => {
    setUnlinked((u) => ({ ...u, [i]: true }));
    setManual((m) => { const c = { ...m }; delete c[i]; return c; });
    setSwapping(null);
  };

  const loadSample = () => {
    reset();
    if (!partner) setPartner("Sample Kitchen");
    const mk = (name, items) => ({ id: Math.random().toString(36).slice(2), file: { name, type: "text/csv" }, status: "done", items, error: null });
    setOffFiles([mk("sample-instore-menu.csv", SAMPLE_INSTORE)]);
    setOnFiles([mk("sample-delivery-menu.csv", SAMPLE_DELIVERY)]);
  };

  const reset = () => {
    setOffFiles([]); setOnFiles([]); setEdits({}); setDropped({});
    setManual({}); setUnlinked({}); setAiLinks({}); setQ("");
    setAiState("idle"); attempted.current = "";
  };

  const exportCSV = () => {
    const head = ["Item (offline)", "Section (offline)", "Item (online)", "Section (online)",
      "Offline price", "Online price", "Delta", "Hike %", "Weight", "Matched by", "Confidence", "Flag"];
    const body = rows.map((r) => [
      r.offName, r.offCat, r.onName, r.onCat, r.offPrice, r.onPrice,
      r.delta.toFixed(2), r.hike?.toFixed(2), r.weight,
      { exact: "Exact name", name: "Name similarity", ai: "Meaning match", weak: "Loose name match", manual: "Set by user" }[r.src],
      Math.round(r.score * 100) + "%",
      r.needsEye ? "Review" : "",
    ]);
    const tail = [[], ["Restaurant", partner || "-"],
      ["Blended inflation %", stats?.blended?.toFixed(2) ?? ""],
      ["Simple average %", stats?.simple?.toFixed(2) ?? ""],
      ["Items compared", stats?.n ?? 0],
      ["Flagged for review", review.length],
      ["Only in offline menu", onlyOffline.length],
      ["Only in online menu", onlyOnline.length]];
    const csv = [head, ...body, ...tail]
      .map((r) => r.map((c) => `"${String(c ?? "").replace(/"/g, '""')}"`).join(",")).join("\n");
    const url = URL.createObjectURL(new Blob([csv], { type: "text/csv" }));
    const a = document.createElement("a");
    a.href = url;
    a.download = `${(partner || "partner").replace(/\s+/g, "_")}_price_delta.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const ready = offline.length > 0 && online.length > 0;
  const tone = headline == null ? "n" : headline > 0.5 ? "up" : headline < -0.5 ? "down" : "n";
  const toneColor = tone === "up" ? C.up : tone === "down" ? C.down : C.inkSoft;

  const srcChip = (r) => {
    if (r.src === "manual") return <Chip tone="neutral" title="You set this match">linked by you</Chip>;
    if (r.src === "ai")
      return <Chip tone="ai" title={r.reason || "Matched on meaning"}><Sparkles size={9} />{r.reason || "same dish"}</Chip>;
    if (r.src === "weak")
      return <Chip tone="warn" title="Names only loosely resemble each other">{Math.round(r.score * 100)}% name match</Chip>;
    return null;
  };

  return (
    <div style={{ background: C.paper, minHeight: "100%", color: C.ink }}>
      <div className="max-w-5xl mx-auto px-5 py-8">
        <header className="mb-7">
          <div style={{ fontFamily: MONO, fontSize: 10.5, letterSpacing: "0.18em", color: C.inkFaint }}>MENU AUDIT</div>
          <h1 style={{ fontFamily: SERIF, fontSize: 32, lineHeight: 1.1 }} className="mt-1">
            Offline vs online price delta
          </h1>
          <p style={{ fontFamily: SANS, fontSize: 13.5, color: C.inkSoft, maxWidth: 640 }} className="mt-2">
            Upload the in-store menu and the listing menu. Items get matched on meaning, not
            just spelling, then rolled up into one blended inflation number for the restaurant.
          </p>
          <div className="mt-4 flex flex-wrap items-center gap-2">
            <input
              value={partner} onChange={(e) => setPartner(e.target.value)}
              placeholder="Restaurant name"
              className="px-3 py-1.5 rounded outline-none"
              style={{ background: C.white, border: `1px solid ${C.rule}`, fontFamily: SANS, fontSize: 13, width: 240 }}
            />
            <button onClick={loadSample} className="flex items-center gap-1.5 px-3 py-1.5 rounded"
              style={{ fontFamily: SANS, fontSize: 12.5, color: C.ink, border: `1px solid ${C.rule}`, background: C.white }}>
              <Sparkles size={12} /> Try with sample menus
            </button>
            {(offFiles.length > 0 || onFiles.length > 0) && (
              <button onClick={reset} className="flex items-center gap-1.5 px-3 py-1.5 rounded"
                style={{ fontFamily: SANS, fontSize: 12.5, color: C.inkSoft, border: `1px solid ${C.rule}` }}>
                <RotateCcw size={12} /> Start over
              </button>
            )}
          </div>
          <div className="mt-3 flex flex-wrap items-center gap-2" style={{ fontFamily: SANS, fontSize: 12, color: C.inkFaint }}>
            <input
              type="password" value={apiKey} placeholder="Anthropic API key (optional)"
              onChange={(e) => { const k = e.target.value.trim(); setApiKey(k); saveKey(k); if (k) setAiOn(true); }}
              className="px-3 py-1.5 rounded outline-none"
              style={{ background: C.white, border: `1px solid ${C.rule}`, fontFamily: MONO, fontSize: 12, width: 240, color: C.ink }}
            />
            <span>Only needed for photos, PDFs and matching renamed dishes. Stays in this browser.</span>
          </div>
        </header>

        <div className="grid md:grid-cols-2 gap-4">
          <UploadCard
            label="Offline menu"
            hint="The printed or in-store menu. Excel or CSV work straight away. Photos and PDFs need an API key. Add as many files as you need."
            files={offFiles} busy={busy.offline}
            onAdd={(f) => addFiles("offline", f)}
            onRemove={(id) => setOffFiles((p) => p.filter((x) => x.id !== id))}
          />
          <UploadCard
            label="Online menu"
            hint="The delivery listing menu. Exports work straight away; app screenshots need an API key."
            files={onFiles} busy={busy.online}
            onAdd={(f) => addFiles("online", f)}
            onRemove={(id) => setOnFiles((p) => p.filter((x) => x.id !== id))}
          />
        </div>

        {!ready && (
          <p className="mt-4 text-center" style={{ fontFamily: SANS, fontSize: 12.5, color: C.inkFaint }}>
            Both menus are needed before anything can be compared.
          </p>
        )}

        {/* matching status */}
        {ready && (
          <div className="mt-4 flex flex-wrap items-center gap-3 px-4 py-2.5 rounded"
            style={{ background: C.card, border: `1px solid ${C.rule}` }}>
            <span style={{ fontFamily: SANS, fontSize: 12.5, color: C.inkSoft }}>
              {strongLinks.size} matched on name
              {Object.keys(aiLinks).length > 0 && `, ${Object.keys(aiLinks).length} matched on meaning`}
              {review.length > 0 && `, ${review.length} to check`}
            </span>
            {aiState === "running" && (
              <span className="flex items-center gap-1.5" style={{ fontFamily: SANS, fontSize: 12, color: C.ai }}>
                <Loader2 size={12} className="animate-spin" /> Reading the leftovers for renamed dishes…
              </span>
            )}
            {aiState === "nokey" && aiOn && (
              <span style={{ fontFamily: SANS, fontSize: 12, color: C.inkFaint }}>Add an API key above to match renamed dishes too.</span>
            )}
            {aiState === "error" && (
              <span style={{ fontFamily: SANS, fontSize: 12, color: C.up }}>Meaning match failed. Name matching still applied.</span>
            )}
            <div className="ml-auto flex items-center gap-2">
              <label className="flex items-center gap-1.5" style={{ fontFamily: SANS, fontSize: 12, color: C.inkSoft }}>
                <input type="checkbox" checked={aiOn} onChange={(e) => setAiOn(e.target.checked)} />
                Match renamed dishes
              </label>
              <button onClick={() => runReconcile(true)} disabled={aiState === "running"}
                className="flex items-center gap-1.5 px-2.5 py-1 rounded"
                style={{ fontFamily: SANS, fontSize: 12, color: C.ink, border: `1px solid ${C.rule}`, background: C.white }}>
                <Repeat size={12} /> Re-run
              </button>
            </div>
          </div>
        )}

        {/* headline */}
        {ready && stats && (
          <section className="mt-4 rounded-lg overflow-hidden" style={{ background: C.white, border: `1px solid ${C.rule}` }}>
            <div className="px-6 pt-6 pb-5">
              <div className="flex flex-wrap items-end gap-x-10 gap-y-5">
                <div>
                  <div style={{ fontFamily: MONO, fontSize: 10.5, letterSpacing: "0.16em", color: C.inkFaint }}>
                    {method === "blended" ? "BLENDED PRICE INFLATION" : "AVERAGE ITEM HIKE"}
                  </div>
                  <div className="flex items-baseline gap-2">
                    <span style={{ fontFamily: SERIF, fontSize: 58, lineHeight: 1, color: toneColor, fontVariantNumeric: "tabular-nums" }}>
                      {headline > 0 ? "+" : ""}{headline.toFixed(1)}<span style={{ fontSize: 30 }}>%</span>
                    </span>
                    {tone === "up" && <ArrowUpRight size={22} style={{ color: C.up }} />}
                    {tone === "down" && <ArrowDownRight size={22} style={{ color: C.down }} />}
                  </div>
                  <div style={{ fontFamily: SANS, fontSize: 12.5, color: C.inkSoft }} className="mt-1">
                    across {stats.n} matched items{stats.weighted && method === "blended" ? ", order-weighted" : ""}
                  </div>
                </div>

                <div className="flex-1 min-w-[260px]">
                  {[["Offline menu value", inr(stats.wOff)],
                    ["Online menu value", inr(stats.wOn)],
                    ["Difference", inr(stats.wOn - stats.wOff)],
                    ["Median item hike", pct(stats.median)]].map(([k, v], idx) => (
                    <div key={k} className="flex items-end">
                      <span style={{ fontFamily: SANS, fontSize: 12.5, color: idx === 2 ? C.ink : C.inkSoft }}>{k}</span>
                      <Leader />
                      <span style={{ fontFamily: MONO, fontSize: 12.5, color: idx === 2 ? toneColor : C.ink, fontVariantNumeric: "tabular-nums" }}>{v}</span>
                    </div>
                  ))}
                </div>
              </div>

              <div className="mt-6">
                <div className="flex h-2 rounded overflow-hidden" style={{ background: "#E4EAE2" }}>
                  <div style={{ width: `${(stats.hiked / stats.n) * 100}%`, background: C.up, transition: "width 300ms" }} />
                  <div style={{ width: `${(stats.same / stats.n) * 100}%`, background: C.inkFaint }} />
                  <div style={{ width: `${(stats.cut / stats.n) * 100}%`, background: C.down }} />
                </div>
                <div className="flex flex-wrap gap-x-5 gap-y-1 mt-2" style={{ fontFamily: SANS, fontSize: 11.5, color: C.inkSoft }}>
                  <span style={{ color: C.up }}>{stats.hiked} priced up</span>
                  <span>{stats.same} unchanged</span>
                  <span style={{ color: C.down }}>{stats.cut} priced down</span>
                  {onlyOffline.length > 0 && <span>{onlyOffline.length} only on offline menu</span>}
                  {onlyOnline.length > 0 && <span>{onlyOnline.length} only listed online</span>}
                </div>
              </div>
            </div>

            <div className="px-6 py-3 flex flex-wrap items-center gap-3" style={{ borderTop: `1px solid ${C.rule}`, background: C.card }}>
              <div className="flex rounded overflow-hidden" style={{ border: `1px solid ${C.rule}` }}>
                {[["blended", "Menu-value blended"], ["simple", "Simple average"]].map(([k, label]) => (
                  <button key={k} onClick={() => setMethod(k)} className="px-3 py-1.5"
                    style={{ fontFamily: SANS, fontSize: 12, background: method === k ? C.ink : "transparent", color: method === k ? "#F2F5F0" : C.inkSoft }}>
                    {label}
                  </button>
                ))}
              </div>
              <span style={{ fontFamily: SANS, fontSize: 11.5, color: C.inkFaint }}>
                {method === "blended"
                  ? "Total online value over total offline value. Expensive items pull harder, which is closest to what a customer actually pays extra."
                  : "Plain mean of every item's hike. Treats a ₹20 starter and a ₹800 platter equally."}
              </span>
              <button onClick={exportCSV} className="ml-auto flex items-center gap-1.5 px-3 py-1.5 rounded"
                style={{ fontFamily: SANS, fontSize: 12.5, color: C.ink, border: `1px solid ${C.rule}`, background: C.white }}>
                <Download size={13} /> Export CSV
              </button>
            </div>
          </section>
        )}

        {/* table */}
        {ready && rows.length > 0 && (
          <section className="mt-6">
            <div className="flex flex-wrap items-center gap-2 mb-3">
              {[["matched", `Compared (${clean.length})`], ["review", `Check these (${review.length})`]].map(([k, label]) => (
                <button key={k} onClick={() => setTab(k)} className="px-3 py-1.5 rounded"
                  style={{
                    fontFamily: SANS, fontSize: 12.5,
                    background: tab === k ? C.ink : C.card,
                    color: tab === k ? "#F2F5F0" : C.inkSoft,
                    border: `1px solid ${tab === k ? C.ink : C.rule}`,
                  }}>
                  {label}
                </button>
              ))}
              <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="Find an item"
                className="px-3 py-1.5 rounded outline-none"
                style={{ background: C.white, border: `1px solid ${C.rule}`, fontFamily: SANS, fontSize: 12.5, width: 160 }} />
              <select value={sort} onChange={(e) => setSort(e.target.value)} className="px-2 py-1.5 rounded outline-none"
                style={{ background: C.white, border: `1px solid ${C.rule}`, fontFamily: SANS, fontSize: 12.5 }}>
                <option value="hike">Biggest hike first</option>
                <option value="drop">Biggest drop first</option>
                <option value="value">Highest online price</option>
                <option value="confidence">Weakest match first</option>
                <option value="name">Item name</option>
              </select>
            </div>

            {tab === "review" && review.length > 0 && (
              <p className="mb-2 px-3 py-2 rounded" style={{ background: C.warnSoft, color: C.warn, fontFamily: SANS, fontSize: 12 }}>
                These pairs are the ones worth a look before the number goes anywhere: loose name
                matches, portion sizes that may not line up, and swings above 60%.
              </p>
            )}

            <div className="rounded-lg overflow-hidden" style={{ background: C.white, border: `1px solid ${C.rule}` }}>
              <div className="grid px-4 py-2"
                style={{ gridTemplateColumns: "1fr 80px 80px 88px 56px 52px", gap: 8,
                  borderBottom: `1px solid ${C.rule}`, background: C.card, fontFamily: MONO,
                  fontSize: 10, letterSpacing: "0.12em", color: C.inkFaint }}>
                <span>ITEM PAIR</span>
                <span className="text-right">OFFLINE</span>
                <span className="text-right">ONLINE</span>
                <span className="text-right">HIKE</span>
                <span className="text-right">WT</span>
                <span />
              </div>

              {view.map((r) => {
                const up = r.hike > 0.5, down = r.hike < -0.5;
                const open = swapping === r.i;
                return (
                  <div key={r.key} style={{ borderBottom: "1px solid #EDF1EC" }}>
                    <div className="grid px-4 py-2.5 items-center"
                      style={{ gridTemplateColumns: "1fr 80px 80px 88px 56px 52px", gap: 8 }}>
                      <div className="min-w-0">
                        <div style={{ fontFamily: SANS, fontSize: 13.5 }} className="truncate">{r.offName}</div>
                        <div className="flex items-center gap-1.5 mt-0.5 min-w-0">
                          {canonName(r.offName) !== canonName(r.onName) && (
                            <span className="truncate" style={{ fontFamily: SANS, fontSize: 11.5, color: C.inkFaint }}>
                              → {r.onName}
                            </span>
                          )}
                          {srcChip(r)}
                          {r.sizeClash && <Chip tone="warn" title="Portion sizes look different">portion differs</Chip>}
                          {r.sizeOneSided && !r.sizeClash && <Chip tone="warn" title="One side names a portion, the other does not">size unclear</Chip>}
                          {r.catClash && <Chip tone="neutral" title={`${r.offCat} vs ${r.onCat}`}>section differs</Chip>}
                          {r.suspicious && <Chip tone="up" title="Unusually large swing">big swing</Chip>}
                        </div>
                      </div>

                      <input value={r.offPrice} onChange={(e) => setEdit(r.key, "offline", toNumber(e.target.value) ?? 0)}
                        className="text-right px-1.5 py-1 rounded outline-none"
                        style={{ fontFamily: MONO, fontSize: 12.5, border: "1px solid transparent", background: "#F4F7F3", fontVariantNumeric: "tabular-nums" }} />
                      <input value={r.onPrice} onChange={(e) => setEdit(r.key, "online", toNumber(e.target.value) ?? 0)}
                        className="text-right px-1.5 py-1 rounded outline-none"
                        style={{ fontFamily: MONO, fontSize: 12.5, border: "1px solid transparent", background: "#F4F7F3", fontVariantNumeric: "tabular-nums" }} />

                      <div className="text-right">
                        <div style={{ fontFamily: MONO, fontSize: 13, color: up ? C.up : down ? C.down : C.inkSoft, fontVariantNumeric: "tabular-nums" }}>
                          {pct(r.hike)}
                        </div>
                        <div style={{ fontFamily: MONO, fontSize: 10.5, color: C.inkFaint }}>
                          {r.delta > 0 ? "+" : ""}{inr(r.delta)}
                        </div>
                      </div>

                      <input value={r.weight} onChange={(e) => setEdit(r.key, "weight", toNumber(e.target.value) ?? 0)}
                        className="text-right px-1.5 py-1 rounded outline-none"
                        style={{ fontFamily: MONO, fontSize: 12, background: "#F4F7F3", border: "1px solid transparent", color: r.weight === 1 ? C.inkFaint : C.ink }} />

                      <div className="flex items-center justify-end gap-1.5">
                        <button onClick={() => setSwapping(open ? null : r.i)} title="Pick a different online item"
                          style={{ color: open ? C.ink : C.inkFaint }}>
                          <Repeat size={13} />
                        </button>
                        <button onClick={() => setDropped((d) => ({ ...d, [r.key]: true }))} title="Remove from comparison"
                          style={{ color: C.inkFaint }}>
                          <Trash2 size={13} />
                        </button>
                      </div>
                    </div>

                    {open && (
                      <div className="px-4 pb-3" style={{ background: "#FAFCF9" }}>
                        <div style={{ fontFamily: SANS, fontSize: 11.5, color: C.inkSoft }} className="mb-1.5">
                          Match “{r.offName}” with:
                        </div>
                        <div className="flex flex-wrap gap-1.5">
                          {candidates(r.i).map(({ o, j, c }) => (
                            <button key={j} onClick={() => relink(r.i, j)}
                              className="px-2 py-1 rounded text-left"
                              style={{
                                fontFamily: SANS, fontSize: 12,
                                background: j === r.j ? C.ink : C.white,
                                color: j === r.j ? "#F2F5F0" : C.ink,
                                border: `1px solid ${j === r.j ? C.ink : C.rule}`,
                              }}>
                              {o.name} · {inr(o.price)}
                              <span style={{ fontFamily: MONO, fontSize: 10, opacity: 0.6 }}> {Math.round(c.score * 100)}%</span>
                            </button>
                          ))}
                          <button onClick={() => breakLink(r.i)} className="px-2 py-1 rounded"
                            style={{ fontFamily: SANS, fontSize: 12, color: C.up, border: `1px solid ${C.upSoft}`, background: C.upSoft }}>
                            Not the same dish
                          </button>
                        </div>
                      </div>
                    )}
                  </div>
                );
              })}

              {view.length === 0 && (
                <div className="px-4 py-8 text-center" style={{ fontFamily: SANS, fontSize: 13, color: C.inkFaint }}>
                  {tab === "review" ? "Nothing flagged. Every pair cleared the checks." : "No items match that search."}
                </div>
              )}
            </div>

            <p className="mt-2" style={{ fontFamily: SANS, fontSize: 11.5, color: C.inkFaint }}>
              Prices are editable, and the swap icon lets you repoint any pair. Weight defaults to 1;
              drop in order counts or GMV share to blend by what actually sells.
            </p>
          </section>
        )}

        {/* unmatched */}
        {ready && (onlyOffline.length > 0 || onlyOnline.length > 0) && (
          <section className="mt-6 grid md:grid-cols-2 gap-4">
            {[["Only on the offline menu", onlyOffline, "offline"], ["Only listed online", onlyOnline, "online"]].map(
              ([title, list, side]) =>
                list.length ? (
                  <div key={side} className="rounded-lg p-4" style={{ background: C.card, border: `1px solid ${C.rule}` }}>
                    <h3 style={{ fontFamily: SERIF, fontSize: 15 }} className="mb-1">{title}</h3>
                    <p style={{ fontFamily: SANS, fontSize: 11.5, color: C.inkFaint }} className="mb-3">
                      {side === "offline"
                        ? "No online counterpart found. Either delisted, or renamed past recognition."
                        : "Delivery-exclusive items sit here, and so do dishes whose offline twin wasn't found."}
                    </p>
                    <ul className="space-y-1">
                      {list.map((it) => {
                        const id = side + ":" + (side === "offline" ? it.i : it.j);
                        const open = swapping === id;
                        return (
                          <li key={id}>
                            <div className="flex items-center gap-2">
                              <span style={{ fontFamily: SANS, fontSize: 12.5 }} className="truncate flex-1">{it.name}</span>
                              <span style={{ fontFamily: MONO, fontSize: 12, color: C.inkSoft, fontVariantNumeric: "tabular-nums" }}>
                                {inr(it.price)}
                              </span>
                              {side === "offline" && (
                                <button onClick={() => setSwapping(open ? null : id)} title="Link to an online item"
                                  style={{ color: open ? C.ink : C.inkFaint }}>
                                  <Link2 size={13} />
                                </button>
                              )}
                            </div>
                            {open && side === "offline" && (
                              <div className="mt-1 mb-2 flex flex-wrap gap-1.5">
                                {candidates(it.i).map(({ o, j, c }) => (
                                  <button key={j} onClick={() => { relink(it.i, j); setSwapping(null); }}
                                    className="px-2 py-1 rounded"
                                    style={{ fontFamily: SANS, fontSize: 11.5, background: C.white, border: `1px solid ${C.rule}` }}>
                                    {o.name} · {inr(o.price)}
                                    <span style={{ fontFamily: MONO, fontSize: 10, color: C.inkFaint }}> {Math.round(c.score * 100)}%</span>
                                  </button>
                                ))}
                              </div>
                            )}
                          </li>
                        );
                      })}
                    </ul>
                  </div>
                ) : null
            )}
          </section>
        )}

        {/* talking points */}
        {ready && stats && stats.top.length > 0 && (
          <section className="mt-6 rounded-lg p-5" style={{ background: C.white, border: `1px solid ${C.rule}` }}>
            <h3 style={{ fontFamily: SERIF, fontSize: 16 }} className="mb-3">What to take into the conversation</h3>
            <ul className="space-y-2" style={{ fontFamily: SANS, fontSize: 13, color: C.inkSoft }}>
              <li className="flex gap-2">
                <Check size={14} style={{ color: C.down, marginTop: 3, flexShrink: 0 }} />
                <span>Blended online pricing sits <b style={{ color: toneColor }}>{pct(stats.blended)}</b> against the in-store menu across {stats.n} comparable items.</span>
              </li>
              {stats.top[0]?.hike > 0 && (
                <li className="flex gap-2">
                  <Check size={14} style={{ color: C.down, marginTop: 3, flexShrink: 0 }} />
                  <span>Sharpest markups: {stats.top.filter((t) => t.hike > 0).map((t) => `${t.offName} (${pct(t.hike)})`).join(", ")}.</span>
                </li>
              )}
              {stats.same > 0 && (
                <li className="flex gap-2">
                  <Check size={14} style={{ color: C.down, marginTop: 3, flexShrink: 0 }} />
                  <span>{stats.same} of {stats.n} items are held at parity, useful proof the restaurant isn't marking up across the board.</span>
                </li>
              )}
              {review.length > 0 && (
                <li className="flex gap-2">
                  <AlertTriangle size={14} style={{ color: C.warn, marginTop: 3, flexShrink: 0 }} />
                  <span>{review.length} pairs are still flagged. Clear those before quoting the blended number.</span>
                </li>
              )}
            </ul>
          </section>
        )}

        <p className="mt-8 text-center" style={{ fontFamily: SANS, fontSize: 11, color: C.inkFaint }}>
          CSV and Excel files are read in your browser. Photos and PDFs go to the Anthropic API with your own key. Nothing is stored. Always eyeball the flagged pairs before the numbers leave this screen.
        </p>
      </div>
    </div>
  );
}