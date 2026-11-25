if (typeof pdfjsLib !== "undefined" && pdfjsLib.GlobalWorkerOptions) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
}

function esc(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

const DEFAULT_DOC = `# Autonomous Solar Rover
Category: Planetary Exploration Robot
Power Source: Solar Array
Primary Sensor: LiDAR Scanner
Autonomous Solar Rover is located in Jezero Crater.
Autonomous Solar Rover uses RTG Backup Battery.

# LiDAR Scanner
LiDAR Scanner is part of Navigation Subsystem.
Navigation Subsystem is part of Autonomous Solar Rover.
LiDAR Scanner requires 24V DC Power.
Purpose: 3D Terrain Mapping

# Dust Storm Event
Dust Storm Event causes Solar Output Drop.
Solar Output Drop requires RTG Backup Battery.

# Logic Rules
IF ?X causes Solar Output Drop THEN ?X is Critical Hazard
IF ?X uses RTG Backup Battery THEN ?X is All-Weather Capable`;

const JSON_DOC = JSON.stringify({
    "Autonomous Solar Rover": {
        "Category": "Planetary Exploration Robot",
        "Power Source": "Solar Array",
        "location": "Jezero Crater",
        "uses": "RTG Backup Battery"
    },
    "LiDAR Scanner": {
        "part_of": "Navigation Subsystem",
        "requires": "24V DC Power",
        "Purpose": "3D Terrain Mapping"
    },
    "Navigation Subsystem": {
        "part_of": "Autonomous Solar Rover"
    },
    "Dust Storm Event": {
        "causes": "Solar Output Drop"
    },
    "Rules": [
        "IF ?X causes Solar Output Drop THEN ?X is Critical Hazard",
        "IF ?X uses RTG Backup Battery THEN ?X is All-Weather Capable"
    ]
}, null, 2);

class MultiFormatReader {
    static async parse(file) {
        if (!file) throw new Error("No file chosen.");
        const ext = file.name.toLowerCase().split(".").pop();
        if (ext === "pdf") return await this.fromPDF(file);
        if (ext === "docx") return await this.fromDOCX(file);
        if (ext === "doc") throw new Error("Legacy .doc files are not supported. Save as .docx or .pdf.");
        if (ext === "xlsx" || ext === "xls") return await this.fromExcel(file);
        if (ext === "csv" || ext === "tsv") return this.fromCSV(await file.text(), ext === "tsv" ? "\t" : ",");
        if (ext === "json") return this.fromJSON(await file.text());
        if (ext === "xml" || ext === "html" || ext === "htm") return this.fromMarkup(await file.text(), ext === "xml");
        return await file.text();
    }

    static async fromPDF(file) {
        if (typeof pdfjsLib === "undefined") throw new Error("PDF decoder unavailable offline.");
        const pdf = await pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
        let out = [];
        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const content = await page.getTextContent();
            let lastY = null, line = "", lines = [];
            for (const item of content.items) {
                const y = item.transform ? Math.round(item.transform[5]) : null;
                if (lastY !== null && y !== null && Math.abs(y - lastY) > 4) {
                    if (line.trim()) lines.push(line.trim());
                    line = item.str;
                } else {
                    line += (line && !line.endsWith(" ") ? " " : "") + item.str;
                }
                lastY = y;
            }
            if (line.trim()) lines.push(line.trim());
            out.push(lines.join("\n"));
        }
        const txt = out.join("\n").trim();
        if (!txt) throw new Error("Scanned image PDF has no selectable text.");
        return txt;
    }

    static async fromDOCX(file) {
        if (typeof mammoth === "undefined") throw new Error("Word decoder unavailable offline.");
        const res = await mammoth.extractRawText({ arrayBuffer: await file.arrayBuffer() });
        const txt = (res.value || "").trim();
        if (!txt) throw new Error("Word document is empty.");
        return txt;
    }

    static async fromExcel(file) {
        if (typeof XLSX === "undefined") throw new Error("Spreadsheet decoder unavailable offline.");
        const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
        let lines = [];
        wb.SheetNames.forEach(name => {
            const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { defval: "" });
            rows.forEach((r, idx) => {
                const keys = Object.keys(r);
                if (keys.length === 3 && keys.map(k => k.toLowerCase()).includes("subject")) {
                    lines.push(`${r[keys[0]]} | ${r[keys[1]]} | ${r[keys[2]]}`);
                } else if (keys.length >= 2) {
                    lines.push(`# ${String(r[keys[0]] || `Record ${idx + 1}`).trim()}`);
                    keys.slice(1).forEach(k => { if (String(r[k]).trim()) lines.push(`${k}: ${r[k]}`); });
                }
            });
        });
        return lines.join("\n");
    }

    static splitDelimited(line, sep) {
        const res = []; let cur = "", q = false;
        for (const ch of line) {
            if (ch === '"') q = !q;
            else if (ch === sep && !q) { res.push(cur.trim()); cur = ""; }
            else cur += ch;
        }
        res.push(cur.trim());
        return res;
    }

    static fromCSV(raw, sep) {
        const rows = raw.split(/\r?\n/).map(r => r.trim()).filter(Boolean);
        if (rows.length < 2) return raw;
        const headers = this.splitDelimited(rows[0], sep);
        let lines = [];
        for (let i = 1; i < rows.length; i++) {
            const cols = this.splitDelimited(rows[i], sep);
            if (cols.length >= 2) {
                lines.push(`# ${cols[0] || `Row ${i}`}`);
                for (let j = 1; j < headers.length; j++) if (cols[j]) lines.push(`${headers[j]}: ${cols[j]}`);
            }
        }
        return lines.join("\n");
    }

    static fromJSON(raw) {
        const obj = JSON.parse(raw);
        const lines = [];
        const addRecord = (name, value) => {
            lines.push(`# ${String(name)}`);
            if (Array.isArray(value)) {
                value.forEach(item => {
                    if (typeof item === "string" && /^IF\s+/i.test(item)) lines.push(item);
                    else if (item && typeof item === "object") addRecord("Record", item);
                    else lines.push(`note: ${String(item)}`);
                });
            } else if (value && typeof value === "object") {
                for (const [key, val] of Object.entries(value)) {
                    if (key.toLowerCase() === "rules" && Array.isArray(val)) {
                        val.forEach(rule => lines.push(String(rule)));
                    } else if (val && typeof val === "object") {
                        addRecord(key, val);
                    } else if (val !== null && val !== undefined) {
                        lines.push(`${key}: ${String(val)}`);
                    }
                }
            } else if (value !== null && value !== undefined) lines.push(`note: ${String(value)}`);
        };
        if (Array.isArray(obj)) {
            obj.forEach((item, i) => addRecord(`Record ${i + 1}`, item));
        } else if (obj && typeof obj === "object") {
            for (const [key, value] of Object.entries(obj)) {
                if (key.toLowerCase() === "rules" && Array.isArray(value)) value.forEach(rule => lines.push(String(rule)));
                else addRecord(key, value);
            }
        } else {
            throw new Error("JSON root must be an object or array.");
        }
        return lines.join("\n");
    }

    static fromMarkup(raw, isXML) {
        const doc = new DOMParser().parseFromString(raw, isXML ? "text/xml" : "text/html");
        if (doc.querySelector("parsererror")) throw new Error("Invalid XML/HTML tags.");
        let lines = [];
        if (!isXML) {
            doc.querySelectorAll("h1, h2, h3, h4, p, li, tr").forEach(el => {
                const tag = el.tagName.toLowerCase();
                const txt = el.textContent.replace(/\s+/g, " ").trim();
                if (!txt) return;
                if (tag.startsWith("h")) lines.push(`# ${txt}`);
                else if (tag === "tr") {
                    const cells = Array.from(el.querySelectorAll("td, th")).map(c => c.textContent.trim());
                    if (cells.length >= 2) lines.push(`${cells[0]}: ${cells.slice(1).join(" - ")}`);
                } else lines.push(txt);
            });
        } else {
            const walk = el => {
                const kids = Array.from(el.children);
                if (!kids.length) { const v = el.textContent.trim(); if (v) lines.push(`${el.tagName}: ${v}`); }
                else { lines.push(`# ${el.getAttribute("name") || el.getAttribute("id") || el.tagName}`); kids.forEach(walk); }
            };
            if (doc.documentElement) walk(doc.documentElement);
        }
        return lines.join("\n");
    }
}

class SymbolicEngine {
    constructor() {
        this.topics = {};
        this.facts = [];
        this.rules = [];
        this.deduced = [];
    }

    clear() {
        this.topics = {}; this.facts = []; this.rules = []; this.deduced = [];
    }

    addFact(s, r, d, deduced = false) {
        const subj = String(s || "").trim();
        const rel = String(r || "").trim().toLowerCase();
        const det = String(d || "").trim().replace(/\.$/, "");
        if (!subj || !rel || !det) return;

        if (!this.facts.some(f => f.s.toLowerCase() === subj.toLowerCase() && f.r === rel && f.d.toLowerCase() === det.toLowerCase())) {
            this.facts.push({ s: subj, r: rel, d: det, deduced });
            if (deduced) this.deduced.push({ s: subj, r: rel, d: det });
        }
        const k = subj.toLowerCase();
        if (!this.topics[k]) this.topics[k] = { name: subj, details: {} };
        if (!this.topics[k].details[rel]) this.topics[k].details[rel] = [];
        if (!this.topics[k].details[rel].includes(det)) this.topics[k].details[rel].push(det);
    }

    parseClause(txt) {
        const clean = txt.trim().replace(/\.$/, "");
        const verbs = [" is ", " has ", " requires ", " causes ", " uses ", " part_of ", " produces ", " enables ", " supports ", " subset_of "];
        for (const v of verbs) {
            const i = clean.toLowerCase().indexOf(v);
            if (i !== -1) return { s: clean.slice(0, i).trim(), r: v.trim().toLowerCase(), d: clean.slice(i + v.length).trim() };
        }
        return null;
    }

    ingest(raw) {
        this.clear();
        let text = (raw || "").trim();
        if (!text) return;
        if ((text.startsWith("{") && text.endsWith("}")) || (text.startsWith("[") && text.endsWith("]"))) {
            text = MultiFormatReader.fromJSON(text);
        }
        const lines = text.split(/\r?\n/).map(l => l.trim()).filter(Boolean);
        let curTopic = "General Notes";

        const patterns = [
            { rx: /^(.+?)\s+is\s+defined\s+as\s+(.+)$/i, rel: "definition" },
            { rx: /^(.+?)\s+is\s+located\s+in\s+(.+)$/i, rel: "location" },
            { rx: /^(.+?)\s+is\s+part\s+of\s+(.+)$/i, rel: "part_of" },
            { rx: /^(.+?)\s+is\s+a\s+subset\s+of\s+(.+)$/i, rel: "subset_of" },
            { rx: /^(.+?)\s+requires\s+(.+)$/i, rel: "requires" },
            { rx: /^(.+?)\s+uses\s+(.+)$/i, rel: "uses" },
            { rx: /^(.+?)\s+produces\s+(.+)$/i, rel: "produces" },
            { rx: /^(.+?)\s+causes\s+(.+)$/i, rel: "causes" },
            { rx: /^(.+?)\s+enables\s+(.+)$/i, rel: "enables" },
            { rx: /^(.+?)\s+supports\s+(.+)$/i, rel: "supports" },
            { rx: /^(.+?)\s+has\s+(.+)$/i, rel: "has" },
            { rx: /^(.+?)\s+is\s+a\s+(.+)$/i, rel: "category" },
            { rx: /^(.+?)\s+is\s+(.+)$/i, rel: "is" }
        ];

        for (const line of lines) {
            const rm = line.match(/^IF\s+(.+?)\s+THEN\s+(.+)$/i);
            if (rm) {
                const a = this.parseClause(rm[1]), c = this.parseClause(rm[2]);
                if (a && c) this.rules.push({ if: a, then: c });
                continue;
            }
            const pipe = line.split("|").map(p => p.trim());
            if (pipe.length === 3 && pipe[0] && pipe[1] && pipe[2]) {
                this.addFact(pipe[0], pipe[1], pipe[2]);
                continue;
            }
            if (line.startsWith("#") || (line.endsWith(":") && line.split(/\s+/).length <= 6)) {
                curTopic = line.replace(/^#+/, "").replace(/:$/, "").trim();
                continue;
            }
            const kv = line.match(/^[-*]?\s*([A-Za-z0-9 _/()-]{2,40}):\s*(.+)$/);
            if (kv) {
                this.addFact(curTopic, kv[1], kv[2]);
                continue;
            }
            const clean = line.replace(/^[-*\s]+/, "").replace(/\.$/, "");
            let ok = false;
            for (const p of patterns) {
                const m = clean.match(p.rx);
                if (m) { this.addFact(m[1], p.rel, m[2]); ok = true; break; }
            }
            if (!ok) this.addFact(curTopic, "note", clean);
        }
        this.forwardChain();
    }

    forwardChain() {
        const factKeys = new Set(this.facts.map(f => `${f.s.toLowerCase()}|${f.r}|${f.d.toLowerCase()}`));
        let frontier = this.facts.slice();
        let rounds = 0;
        while (frontier.length && rounds++ < 100) {
            const next = [];
            for (const rule of this.rules) {
                for (const fact of frontier) {
                    if (fact.r !== rule.if.r || fact.d.toLowerCase() !== rule.if.d.toLowerCase()) continue;
                    if (rule.if.s !== "?X" && rule.if.s.toLowerCase() !== fact.s.toLowerCase()) continue;
                    const subject = rule.then.s === "?X" ? fact.s : rule.then.s;
                    const key = `${subject.toLowerCase()}|${rule.then.r}|${rule.then.d.toLowerCase()}`;
                    if (factKeys.has(key)) continue;
                    factKeys.add(key);
                    this.addFact(subject, rule.then.r, rule.then.d, true);
                    next.push({ s: subject, r: rule.then.r, d: rule.then.d, deduced: true });
                }
            }
            frontier = next;
        }
    }

    bfsChain(a, b) {
        const start = a.toLowerCase(), end = b.toLowerCase();
        const adjacency = new Map();
        for (const fact of this.facts) {
            const key = fact.s.toLowerCase();
            if (!adjacency.has(key)) adjacency.set(key, []);
            adjacency.get(key).push(fact);
        }
        const queue = [start];
        let head = 0;
        const seen = new Set([start]);
        const previous = new Map();
        while (head < queue.length) {
            const current = queue[head++];
            if (current === end) break;
            for (const fact of adjacency.get(current) || []) {
                const next = fact.d.toLowerCase();
                if (seen.has(next)) continue;
                seen.add(next);
                previous.set(next, { from: current, label: `${fact.s} (${fact.r.replace(/_/g, " ")}) -> ${fact.d}` });
                queue.push(next);
            }
        }
        if (!seen.has(end) || start === end) return null;
        const path = [];
        let current = end;
        while (current !== start) {
            const step = previous.get(current);
            if (!step) return null;
            path.push(step.label);
            current = step.from;
        }
        return path.reverse();
    }

    ask(query) {
        const clean = String(query || "").trim().replace(/\?+$/, "").trim();
        const low = clean.toLowerCase();
        const steps = [];
        if (!clean) return { found: false, answer: "Please enter a question.", method: "Empty Query", steps: ["No text entered."], score: 0 };
        if (!this.facts.length) return { found: false, answer: "Knowledge Base is empty.", method: "No Data", steps: ["Load a document first."], score: 0 };

        const escRx = s => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
        const matched = Object.entries(this.topics)
            .filter(([k]) => k !== "general notes" && new RegExp(`(^|[^a-z0-9])${escRx(k)}([^a-z0-9]|$)`, "i").test(low))
            .map(([, v]) => v.name)
            .sort((a, b) => b.length - a.length);

        if (matched.length >= 2) {
            const [t1, t2] = matched;
            steps.push(`Identified two topics: "${t1}" and "${t2}".`);
            const forward = this.bfsChain(t1, t2);
            const reverse = this.bfsChain(t2, t1);
            const rankPath = path => path ? (path.filter(x => x.includes("part of")).length * 100 - path.length) : -Infinity;
            const chain = rankPath(reverse) > rankPath(forward) ? reverse : forward;
            if (chain) {
                steps.push("Connected topics via Breadth-First Search (BFS), prioritizing structural relationships.");
                return { found: true, answer: `Yes, connected via: ${chain.join(" -> ")}.`, method: "Multi-Hop BFS Search", steps: [...steps, ...chain], score: 100 };
            }
        }

        let targetRels = [];
        if (/^(what is|define|explain)/i.test(low)) { targetRels = ["definition", "is", "category", "note"]; steps.push("Intent: Definition / Category lookup."); }
        else if (/^(where|location)/i.test(low)) { targetRels = ["location", "part_of"]; steps.push("Intent: Location lookup."); }
        else if (/^(why|cause)/i.test(low)) { targetRels = ["causes", "requires"]; steps.push("Intent: Cause / Dependency lookup."); }
        else if (low.includes("require") || low.includes("need")) { targetRels = ["requires", "uses"]; steps.push("Intent: Requirement lookup."); }

        if (matched.length) {
            const topic = matched[0];
            const obj = this.topics[topic.toLowerCase()];
            steps.push(`Matched topic: "${topic}".`);

            for (const [rel, vals] of Object.entries(obj.details)) {
                const readable = rel.replace(/_/g, " ");
                if ((low.includes(rel) || low.includes(readable)) && rel !== "is" && rel !== "has") {
                    const isRule = vals.some(v => this.deduced.some(df => df.s === topic && df.r === rel && df.d === v));
                    steps.push(`Matched detail "${readable}": ${vals.join(", ")}.`);
                    return { found: true, answer: `The ${readable} of ${topic} is: ${vals.join(", ")}.`, method: isRule ? "Logic Rule Deduction" : "Direct Fact Lookup", steps, score: 100 };
                }
            }
            for (const rel of targetRels) {
                if (obj.details[rel]) {
                    const vals = obj.details[rel];
                    const isRule = vals.some(v => this.deduced.some(df => df.s === topic && df.r === rel && df.d === v));
                    steps.push(`Matched "${rel.replace(/_/g, " ")}" -> ${vals.join(", ")}.`);
                    return { found: true, answer: `${topic}: ${vals.join(", ")}.`, method: isRule ? "Logic Rule Deduction" : "Smart Topic Match", steps, score: 95 };
                }
            }
            const summary = Object.entries(obj.details).map(([k, v]) => `${k.replace(/_/g, " ")}: ${v.join(", ")}`);
            steps.push("Returned full topic summary.");
            return { found: true, answer: `${topic} — ${summary.join("; ")}.`, method: "Topic Summary", steps, score: 85 };
        }

        const stop = new Set(["what", "where", "when", "which", "does", "have", "about", "tell", "give", "the", "and", "for"]);
        const tokens = (low.match(/[a-z0-9]+/g) || []).filter(w => w.length > 3 && !stop.has(w));
        let best = null, hits = 0;
        for (const f of this.facts) {
            const str = `${f.s} ${f.r} ${f.d}`.toLowerCase();
            const h = tokens.reduce((a, t) => a + (str.includes(t) ? 1 : 0), 0);
            if (h > hits) { hits = h; best = f; }
        }
        if (best && hits > 0) {
            steps.push(`Matched ${hits} keyword(s) in saved fact.`);
            return { found: true, answer: `${best.s} ${best.r.replace(/_/g, " ")} ${best.d}.`, method: "Keyword Fact Match", steps, score: Math.min(85, 45 + hits * 15) };
        }
        return { found: false, answer: "No matching answer found in your document.", method: "Safe Abstention", steps: ["Checked all topics, facts, and logic rules; avoided guessing."], score: 0 };
    }

    benchmark() {
        const cases = [
            { q: "What is the Power Source of Autonomous Solar Rover?", exp: "Solar Array" },
            { q: "Where is Autonomous Solar Rover located?", exp: "Jezero Crater" },
            { q: "What does LiDAR Scanner require?", exp: "24V DC Power" },
            { q: "How is LiDAR Scanner related to Autonomous Solar Rover?", exp: "Navigation Subsystem" },
            { q: "What is Dust Storm Event?", exp: "Critical Hazard" },
            { q: "What is Autonomous Solar Rover?", exp: "All-Weather Capable" },
            { q: "What is the fuel capacity of Nuclear Submarine?", exp: "No matching answer" }
        ];
        let passed = 0;
        const rows = cases.map(c => {
            const r = this.ask(c.q);
            const ok = r.answer.toLowerCase().includes(c.exp.toLowerCase());
            if (ok) passed++;
            return { ...c, actual: r.answer, method: r.method, ok };
        });
        return { total: cases.length, passed, pct: Math.round((passed / cases.length) * 100), rows };
    }
}

const engine = new SymbolicEngine();

function switchView(id, btn) {
    document.querySelectorAll(".view").forEach(v => v.classList.remove("active"));
    document.querySelectorAll(".nav-item").forEach(b => b.classList.remove("active"));
    const target = document.getElementById(id);
    if (!target || !btn) return;
    target.classList.add("active");
    btn.classList.add("active");
}

function setMsg(txt, isErr = false) {
    const el = document.getElementById("status-msg");
    el.textContent = txt;
    el.className = isErr ? "status-line err" : "status-line";
}

async function handleFileSelect(ev) {
    const file = ev.target.files[0];
    if (!file) return;
    setMsg(`Reading "${file.name}"...`);
    try {
        const extracted = await MultiFormatReader.parse(file);
        document.getElementById("raw-editor").value = extracted;
        compileFacts();
        setMsg(`Loaded "${file.name}" — ${engine.facts.length} facts extracted.`);
    } catch (e) {
        setMsg(`Error: ${e.message}`, true);
    }
    ev.target.value = "";
}

function loadDefaultSample() {
    document.getElementById("raw-editor").value = DEFAULT_DOC;
    compileFacts();
    setMsg("Ready — Sample document loaded.");
}

function loadJsonSample() {
    document.getElementById("raw-editor").value = JSON_DOC;
    compileFacts();
    setMsg("Ready — JSON example parsed into facts.");
}

function compileFacts() {
    const source = document.getElementById("raw-editor").value;
    try { localStorage.setItem("logicAtlas.source", source); } catch (e) { /* storage may be disabled */ }
    engine.ingest(source);
    document.getElementById("m-topics").textContent = Object.keys(engine.topics).length;
    document.getElementById("m-facts").textContent = engine.facts.length;
    document.getElementById("m-rules").textContent = engine.rules.length;
    document.getElementById("m-deduced").textContent = engine.deduced.length;
    renderFactRows();
    renderTopicDirectory();
}

function renderFactRows() {
    const q = document.getElementById("filter-box").value.trim().toLowerCase();
    const tbody = document.getElementById("facts-list");
    tbody.innerHTML = "";
    const list = q ? engine.facts.filter(f => `${f.s} ${f.r} ${f.d}`.toLowerCase().includes(q)) : engine.facts;
    if (!list.length) {
        tbody.innerHTML = `<tr><td colspan="4" style="text-align:center;padding:1rem;">No matching facts.</td></tr>`;
        return;
    }
    list.forEach(f => {
        const origin = f.deduced ? '<span class="tag tag-amber">Rule Deduced</span>' : '<span class="tag tag-green">From File</span>';
        tbody.innerHTML += `<tr><td><strong>${esc(f.s)}</strong></td><td><span class="tag">${esc(f.r.replace(/_/g, " "))}</span></td><td>${esc(f.d)}</td><td>${origin}</td></tr>`;
    });
}

function renderTopicDirectory() {
    const box = document.getElementById("topic-cards");
    const items = Object.values(engine.topics);
    document.getElementById("topic-pill").textContent = `${items.length} Topics`;
    box.innerHTML = "";
    items.forEach(t => {
        let lines = "";
        for (const [k, vals] of Object.entries(t.details)) {
            lines += `<div class="kv-line"><span class="kv-k">${esc(k.replace(/_/g, " "))}</span><span class="kv-v">${esc(vals.join(", "))}</span></div>`;
        }
        box.innerHTML += `<div class="entry-card"><div class="entry-head"><span>${esc(t.name)}</span><span class="tag">${Object.keys(t.details).length}</span></div>${lines}</div>`;
    });
}

function askSample(q) {
    document.getElementById("q-box").value = q;
    runQuery();
}

function runQuery() {
    const res = engine.ask(document.getElementById("q-box").value);
    const sheet = document.getElementById("result-card");
    sheet.style.display = "block";
    const m = document.getElementById("res-method");
    m.textContent = res.method;
    m.className = res.found ? "tag tag-green" : "tag tag-amber";
    document.getElementById("res-conf").textContent = res.found ? `${res.score}% Match` : "Not Found";
    document.getElementById("res-text").textContent = res.answer;
    document.getElementById("res-steps").innerHTML = "<strong>Reasoning Steps:</strong><br>" + res.steps.map(s => `• ${esc(s)}`).join("<br>");
}

function executeBenchmark() {
    const rep = engine.benchmark();
    const banner = document.getElementById("bench-banner");
    banner.style.display = "block";
    banner.textContent = `Evaluation Score: ${rep.passed} / ${rep.total} Passed (${rep.pct}% Accuracy)`;
    const tbody = document.getElementById("bench-rows");
    tbody.innerHTML = "";
    rep.rows.forEach(r => {
        const badge = r.ok ? '<span class="tag tag-green">PASS</span>' : '<span class="tag tag-red">FAIL</span>';
        tbody.innerHTML += `<tr><td><strong>${esc(r.q)}</strong></td><td>${esc(r.exp)}</td><td>${esc(r.actual)}</td><td><span class="tag">${esc(r.method)}</span></td><td>${badge}</td></tr>`;
    });
}

window.addEventListener("DOMContentLoaded", () => {
    let saved = "";
    try { saved = localStorage.getItem("logicAtlas.source") || ""; } catch (e) { /* storage may be disabled */ }
    if (saved) {
        document.getElementById("raw-editor").value = saved;
        compileFacts();
        setMsg("Restored your last knowledge base from this browser.");
    } else loadDefaultSample();
    const drop = document.getElementById("drop-area");
    ["dragenter", "dragover"].forEach(e => drop.addEventListener(e, ev => { ev.preventDefault(); drop.classList.add("over"); }));
    ["dragleave", "drop"].forEach(e => drop.addEventListener(e, ev => { ev.preventDefault(); drop.classList.remove("over"); }));
    drop.addEventListener("drop", ev => {
        const f = ev.dataTransfer && ev.dataTransfer.files[0];
        if (f) {
            const dt = new DataTransfer();
            dt.items.add(f);
            const inp = document.getElementById("file-input");
            inp.files = dt.files;
            handleFileSelect({ target: inp });
        }
    });
});
