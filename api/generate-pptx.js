const pptxgen = require("pptxgenjs");
const { createClient } = require("@supabase/supabase-js");

// ── Supabase Config ──
const SUPABASE_URL = "https://hsmawbvxhlkssdowlbzx.supabase.co";
const SUPABASE_KEY = process.env.SUPABASE_SERVICE_KEY;

// ── Brand Colors (from real TOMAS TECH proposals) ──
const C = {
  headerBlue: "1F5BA8",
  navy:       "1C3F7F",
  orange:     "F7941D",
  orangeDark: "E8722A",
  teal:       "2BB5B8",
  white:      "FFFFFF",
  black:      "000000",
  gray:       "595959",
  ltgray:     "D9D9D9",
  green:      "70AD47",
};

// ═══════════════════════════════════════════════════════════
// ROBUST HELPERS — handle any JSON shape AI might return
// ═══════════════════════════════════════════════════════════

// Deeply flatten any value to a clean string
function toStr(val) {
  if (val === null || val === undefined) return "";
  if (typeof val === "string") return val;
  if (typeof val === "number" || typeof val === "boolean") return String(val);
  if (Array.isArray(val)) return val.map(toStr).filter(Boolean).join(", ");
  if (typeof val === "object") {
    // Try common keys first
    for (const k of ["text","title","name","description","summary","value","label","detail","role","item","service","function"]) {
      if (val[k] && typeof val[k] === "string") return val[k];
    }
    // Fallback: join all string values
    const parts = Object.values(val).map(v => {
      if (typeof v === "string") return v;
      if (typeof v === "number") return String(v);
      return null;
    }).filter(Boolean);
    if (parts.length > 0) return parts.join(" — ");
    return JSON.stringify(val);
  }
  return String(val);
}

// Extract a flat array from any content shape
function toArray(content) {
  if (!content) return [];
  if (Array.isArray(content)) return content;
  if (typeof content === "string") return content.split("\n").filter(l => l.trim());
  if (typeof content === "object") {
    // Search for any array-valued key
    for (const k of Object.keys(content)) {
      if (Array.isArray(content[k]) && content[k].length > 0) return content[k];
    }
    // If object has numbered keys or is a single-level map, convert to array
    const entries = Object.entries(content);
    if (entries.length > 0 && entries.every(([k]) => !["summary","description","overview","total","bar_title","bar_color","erp_name","server_name","connection_text"].includes(k))) {
      return entries.map(([k, v]) => {
        if (typeof v === "string") return { title: k, description: v };
        if (typeof v === "object" && !Array.isArray(v)) return { title: k, ...v };
        return { title: k, description: toStr(v) };
      });
    }
    return [];
  }
  return [];
}

// Get a simple text summary from content
function toText(content, fallback) {
  if (!content) return fallback || "";
  if (typeof content === "string") return content;
  if (typeof content === "object") {
    // Check known text keys
    for (const k of ["summary","description","overview","text","detail","content"]) {
      if (content[k] && typeof content[k] === "string") return content[k];
    }
    // If it has arrays, format them as bullet list
    const arr = toArray(content);
    if (arr.length > 0) {
      return arr.map((item, i) => {
        const t = typeof item === "string" ? item : item.title || item.name || toStr(item);
        const d = typeof item === "object" ? (item.description || item.detail || "") : "";
        return d ? `${i+1}. ${t}: ${d}` : `${i+1}. ${t}`;
      }).join("\n\n");
    }
    // Last resort: pretty print
    return JSON.stringify(content, null, 2);
  }
  return String(content);
}

// ═══════════════════════════════════════════════════════════
// SLIDE FRAMEWORK
// ═══════════════════════════════════════════════════════════

function addFooter(pptx, s, n) {
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 7.25, w: 13.3, h: 0.25,
    fill: { color: C.navy }, line: { color: C.navy },
  });
  s.addText("Copyright\u00A9 TOMAS TECH CORPORATION. All rights reserved.", {
    x: 0, y: 7.25, w: 13.3, h: 0.25,
    fontSize: 7, color: C.white, align: "center", valign: "middle",
  });
  if (n) s.addText(String(n), { x: 12.6, y: 7.0, w: 0.5, h: 0.25, fontSize: 10, color: C.gray, align: "right" });
  s.addText("Tomas Tech Co., Ltd.", { x: 0.3, y: 7.0, w: 3, h: 0.25, fontSize: 9, color: C.gray });
}

function addPegasusLogo(pptx, s) {
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.18, y: 0.1, w: 1.5, h: 0.38,
    fill: { color: C.headerBlue }, line: { color: C.headerBlue }, rectRadius: 0.04,
  });
  s.addText("\u2708 PEGASUS", {
    x: 0.18, y: 0.1, w: 1.5, h: 0.38,
    fontSize: 10, bold: true, color: C.white, align: "center", valign: "middle",
  });
}

function addHeader(pptx, s, title, n) {
  addPegasusLogo(pptx, s);
  s.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0.57, w: 13.3, h: 0.06, fill: { color: C.headerBlue }, line: { color: C.headerBlue } });
  s.addText("     " + title, { x: 0, y: 0.65, w: 13.3, h: 0.45, fontSize: 16, bold: true, color: C.headerBlue, valign: "middle" });
  s.addShape(pptx.shapes.LINE, { x: 0.18, y: 1.1, w: 12.94, h: 0, line: { color: C.headerBlue, pt: 2 } });
  addFooter(pptx, s, n);
}

function addSectionDivider(pptx, title, n) {
  const s = pptx.addSlide();
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 2.5, y: 3.0, w: 8.3, h: 0.9,
    fill: { color: C.navy }, line: { color: C.navy }, rectRadius: 0.1,
  });
  s.addText(title, {
    x: 2.5, y: 3.0, w: 8.3, h: 0.9,
    fontSize: 26, bold: true, color: C.white, align: "center", valign: "middle",
  });
  addFooter(pptx, s, n);
}

function addImageZone(pptx, s, x, y, w, h, label) {
  s.addShape(pptx.shapes.RECTANGLE, { x, y, w, h, fill: { color: "F0F4F8" }, line: { color: "AABBCC", pt: 1.5, dashType: "dash" } });
  s.addText(label, { x, y: y + h/2 - 0.15, w, h: 0.3, fontSize: 9, bold: true, color: "4A6FA5", align: "center", valign: "middle" });
}

// ═══════════════════════════════════════════════════════════
// SLIDE 1 — COVER
// ═══════════════════════════════════════════════════════════
function buildCoverSlide(pptx, data) {
  const s = pptx.addSlide();
  addPegasusLogo(pptx, s);

  s.addText((data.proposal_title || data.service_type || "SYSTEM PROPOSAL").toUpperCase(), {
    x: 0.3, y: 0.7, w: 12.7, h: 1.0, fontSize: 34, bold: true, color: C.black, fontFace: "Calibri",
  });

  s.addShape(pptx.shapes.LINE, { x: 0.3, y: 2.0, w: 12.7, h: 0, line: { color: C.ltgray, pt: 1.5 } });
  s.addText("PROPOSAL FOR :", { x: 0.3, y: 2.1, w: 5, h: 0.45, fontSize: 18, color: C.gray });
  s.addShape(pptx.shapes.LINE, { x: 0.3, y: 2.62, w: 12.7, h: 0, line: { color: C.ltgray, pt: 2 } });

  s.addText((data.client_name || "CLIENT").toUpperCase(), {
    x: 0.3, y: 2.72, w: 12.7, h: 0.72, fontSize: 28, bold: true, color: C.black, underline: true,
  });
  s.addShape(pptx.shapes.LINE, { x: 0.3, y: 3.52, w: 12.7, h: 0, line: { color: C.ltgray, pt: 1.5 } });

  s.addText("Tomas Tech Co., Ltd.", { x: 0.3, y: 3.62, w: 6, h: 0.38, fontSize: 14, color: C.gray });
  const now = new Date();
  const months = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  s.addText(`${months[now.getMonth()]} ${now.getFullYear()}`, { x: 10.5, y: 3.62, w: 2.5, h: 0.38, fontSize: 14, color: C.gray, align: "right" });

  addImageZone(pptx, s, 0.3, 5.5, 2.5, 1.5, "TOMAS TECH LOGO");
  addImageZone(pptx, s, 9.8, 4.3, 3.2, 2.7, "ILLUSTRATION");

  s.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 7.25, w: 13.3, h: 0.25, fill: { color: C.navy }, line: { color: C.navy } });
  s.addText("Copyright\u00A9 TOMAS TECH CORPORATION. All rights reserved.", { x: 0, y: 7.25, w: 13.3, h: 0.25, fontSize: 7, color: C.white, align: "center", valign: "middle" });
}

// ═══════════════════════════════════════════════════════════
// SLIDE — PROJECT PURPOSE (objectives + chevron flow)
// Dynamic Y: chevrons positioned BELOW objectives with gap
// ═══════════════════════════════════════════════════════════
function buildProjectPurpose(pptx, data, section) {
  const s = pptx.addSlide();
  addHeader(pptx, s, "Project Purpose", 2);

  const content = section.content;
  let mainText = "", objectives = [], steps = [];

  if (typeof content === "string") {
    mainText = content;
  } else if (content) {
    mainText = content.summary || content.description || content.overview || "";
    objectives = toArray(content.objectives || content.goals || content.key_points || []);
    steps = toArray(content.process_flow || content.steps || content.flow || []);
  }

  // --- Main description text ---
  let yPos = 1.25;
  if (mainText) {
    // Use autoFit to handle long text gracefully
    s.addText(mainText, { x: 0.3, y: yPos, w: 12.7, h: 0.75, fontSize: 11, color: C.black, valign: "top", shrinkText: true });
    yPos += 0.85;
  }

  // --- Objectives (numbered, with proper spacing) ---
  const maxObj = steps.length > 0 ? 3 : 5;
  const objCount = Math.min(objectives.length, maxObj);
  // Dynamic row height based on count and remaining space
  const objRowH = steps.length > 0 ? 0.65 : 0.75;

  objectives.slice(0, objCount).forEach((o, i) => {
    // Number badge
    s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 0.35, y: yPos + i * objRowH + 0.05, w: 0.32, h: 0.32, fill: { color: C.headerBlue }, rectRadius: 0.04 });
    s.addText(String(i + 1), { x: 0.35, y: yPos + i * objRowH + 0.05, w: 0.32, h: 0.32, fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle" });
    // Objective text with shrinkText to prevent overflow
    s.addText(toStr(o), {
      x: 0.78, y: yPos + i * objRowH, w: 11.7, h: objRowH - 0.08,
      fontSize: 11, color: C.black, valign: "middle", shrinkText: true,
    });
    // Light separator line
    if (i < objCount - 1) {
      s.addShape(pptx.shapes.LINE, { x: 0.78, y: yPos + (i + 1) * objRowH - 0.02, w: 11.5, h: 0, line: { color: "E8E8E8", pt: 0.5 } });
    }
  });
  yPos += objCount * objRowH + 0.25;

  // --- Process Flow chevrons ---
  if (steps.length > 0) {
    const flowY = Math.min(yPos, 5.5);
    const flowColors = [C.headerBlue, C.teal, C.headerBlue, C.teal, C.orange, C.orange];
    const maxSteps = Math.min(steps.length, 6);
    const stepW = Math.min(1.92, 12.4 / maxSteps - 0.18);

    steps.slice(0, maxSteps).forEach((st, i) => {
      const text = typeof st === "string" ? st : st.name || st.title || st.t || `Step ${i+1}`;
      const col = (typeof st === "object" && st.color) ? st.color : flowColors[i % flowColors.length];
      const x = 0.3 + i * (stepW + 0.18);

      s.addShape(pptx.shapes.CHEVRON, { x, y: flowY, w: stepW, h: 0.88, fill: { color: col }, line: { color: col } });
      s.addText(text, { x: x + 0.15, y: flowY, w: stepW - 0.3, h: 0.88, fontSize: 9.5, bold: true, color: C.white, align: "center", valign: "middle", shrinkText: true });
    });
  }
}

// ═══════════════════════════════════════════════════════════
// SLIDE — EXECUTIVE SUMMARY (text)
// ═══════════════════════════════════════════════════════════
function buildExecutiveSummary(pptx, data, section) {
  const s = pptx.addSlide();
  addHeader(pptx, s, "Executive Summary", 2);
  const text = toText(section.content, "");
  s.addText(text.substring(0, 3000), {
    x: 0.5, y: 1.3, w: 12.3, h: 5.5, fontSize: 13, color: "333333", valign: "top", lineSpacing: 22, paraSpaceAfter: 8,
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE — PAIN POINT (problem cards + illustration zone)
// ═══════════════════════════════════════════════════════════
function buildPainPoint(pptx, data, section) {
  const s = pptx.addSlide();
  addHeader(pptx, s, "Current Pain Point", 3);

  const problems = toArray(section.content);
  // Limit to 4 cards max to prevent overflow on 7.5" slide
  const maxP = Math.min(problems.length, 4);
  // Dynamic card height: fit within y=1.3 to y=6.85 (5.55" total)
  // Each card needs gap of 0.16 between them
  const totalH = 5.55;
  const cardH = Math.min(1.35, (totalH - (maxP - 1) * 0.16) / Math.max(maxP, 1));

  problems.slice(0, maxP).forEach((p, i) => {
    const y = 1.3 + i * (cardH + 0.16);
    const title = typeof p === "string" ? p : (p.title || p.name || p.problem || toStr(p));
    const pts = Array.isArray(p.points) ? p.points : (p.details ? (Array.isArray(p.details) ? p.details : [p.details]) : (p.description ? [p.description] : []));

    s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 0.3, y, w: 9.2, h: cardH, fill: { color: C.ltgray }, line: { color: "BBBBBB", pt: 1 }, rectRadius: 0.07 });
    // Title: starts 0.06" from top, height 0.32"
    s.addText(`${i+1}. ${title}`, { x: 0.5, y: y + 0.06, w: 8.8, h: 0.32, fontSize: 12, bold: true, color: C.black, valign: "middle" });

    // Bullets: start after title ends (0.06 + 0.32 + 0.04 gap = 0.42")
    // But limit bullets based on remaining card space
    const bulletStartY = y + 0.42;
    const remainingH = cardH - 0.42 - 0.04; // leave 0.04" bottom padding
    const bulletH = 0.24;
    const bulletGap = 0.27;
    const maxBullets = Math.min(pts.length, 3, Math.floor(remainingH / bulletGap));

    pts.slice(0, maxBullets).forEach((pt, pi) => {
      s.addText(`- ${toStr(pt)}`, { x: 0.7, y: bulletStartY + pi * bulletGap, w: 8.6, h: bulletH, fontSize: 10, color: "333333", valign: "top" });
    });
  });

  addImageZone(pptx, s, 9.7, 1.3, 3.3, 5.55, "Problem Illustration");
}

// ═══════════════════════════════════════════════════════════
// SLIDE — BENEFITS (3-column: Problem / Benefits / Use-Reduce)
// ═══════════════════════════════════════════════════════════
function buildBenefitsOverview(pptx, data, section) {
  const s = pptx.addSlide();
  addHeader(pptx, s, "Benefits Target", 4);

  const content = section.content;
  let problems = [], benefits = [], reductions = [];

  if (content) {
    problems = toArray(content.problems || content.pain_points || content.current_issues || []);
    benefits = toArray(content.benefits || content.improvements || content.solutions || []);
    reductions = toArray(content.reductions || content.use_reduce || content.metrics || content.kpis || []);
  }

  const colY = 1.25, colH = 5.65, colW = 3.8;
  const headerH = 0.4;
  const contentStartY = colY + headerH + 0.12; // 0.12" gap after header
  const contentAreaH = colH - headerH - 0.12;  // available space for items

  // === Col 1: Problem (gray) ===
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.18, y: colY, w: colW, h: colH, fill: { color: "F2F2F2" }, line: { color: C.ltgray, pt: 1 } });
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.18, y: colY, w: colW, h: headerH, fill: { color: "888888" } });
  s.addText("\u26A0  Problem of Normal process", { x: 0.18, y: colY, w: colW, h: headerH, fontSize: 10, bold: true, color: C.white, align: "center", valign: "middle" });

  // Dynamic: limit items so they fit in available space
  const pItemH = 0.42;
  const maxProblems = Math.min(problems.length, 6, Math.floor(contentAreaH / pItemH));
  const pSpacing = maxProblems > 0 ? Math.min(contentAreaH / maxProblems, 0.55) : 0.5;

  problems.slice(0, maxProblems).forEach((p, pi) => {
    s.addText("\u2717  " + toStr(p), { x: 0.28, y: contentStartY + pi * pSpacing, w: colW - 0.2, h: pItemH, fontSize: 9.5, color: "444444", valign: "middle" });
  });

  // Arrow 1
  s.addShape(pptx.shapes.CHEVRON, { x: 4.1, y: 3.6, w: 0.6, h: 0.65, fill: { color: C.orange }, line: { color: C.orange } });

  // === Col 2: Benefits (orange) ===
  s.addShape(pptx.shapes.RECTANGLE, { x: 4.78, y: colY, w: colW, h: colH, fill: { color: "FFF8F0" }, line: { color: "FFCCAA", pt: 1 } });
  s.addShape(pptx.shapes.RECTANGLE, { x: 4.78, y: colY, w: colW, h: headerH, fill: { color: C.orange } });
  s.addText("\u2714  Benefits", { x: 4.78, y: colY, w: colW, h: headerH, fontSize: 10, bold: true, color: C.white, align: "center", valign: "middle" });

  const bItemH = 0.42;
  const maxBenefits = Math.min(benefits.length, 6, Math.floor(contentAreaH / bItemH));
  const bSpacing = maxBenefits > 0 ? Math.min(contentAreaH / maxBenefits, 0.55) : 0.5;

  benefits.slice(0, maxBenefits).forEach((b, bi) => {
    s.addText("\u2714  " + toStr(b), { x: 4.9, y: contentStartY + bi * bSpacing, w: colW - 0.2, h: bItemH, fontSize: 9.5, color: "444444", valign: "middle" });
  });

  // Arrow 2
  s.addShape(pptx.shapes.CHEVRON, { x: 8.7, y: 3.6, w: 0.6, h: 0.65, fill: { color: C.headerBlue }, line: { color: C.headerBlue } });

  // === Col 3: Use/Reduce (blue) ===
  s.addShape(pptx.shapes.RECTANGLE, { x: 9.35, y: colY, w: colW, h: colH, fill: { color: "EEF4FF" }, line: { color: "C0D0E8", pt: 1 } });
  s.addShape(pptx.shapes.RECTANGLE, { x: 9.35, y: colY, w: colW, h: headerH, fill: { color: C.headerBlue } });
  s.addText("Use / Reduce", { x: 9.35, y: colY, w: colW, h: headerH, fontSize: 10, bold: true, color: C.white, align: "center", valign: "middle" });

  // Reductions: limit to 3, dynamically space within available area
  const maxRed = Math.min(reductions.length, 3);
  const redBlockH = maxRed > 0 ? Math.min(contentAreaH / maxRed, 1.7) : 1.65;

  reductions.slice(0, maxRed).forEach((r, ri) => {
    const ry = contentStartY + ri * redBlockH;
    const value = typeof r === "string" ? r : (r.value || r.metric || r.number || toStr(r));
    const label = typeof r === "object" ? (r.label || r.title || r.name || "") : "";
    const sub = typeof r === "object" ? (r.subtitle || r.description || r.detail || "") : "";

    s.addText(String(value), { x: 9.35, y: ry, w: colW, h: 0.6, fontSize: 26, bold: true, color: C.orange, align: "center", valign: "bottom" });
    if (label) s.addText(label, { x: 9.35, y: ry + 0.6, w: colW, h: 0.38, fontSize: 10, color: C.black, align: "center", valign: "top" });
    if (sub) s.addText(sub, { x: 9.35, y: ry + 0.96, w: colW, h: 0.28, fontSize: 9, color: C.gray, align: "center", italic: true, valign: "top" });
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE — SYSTEM OUTLINE (architecture diagram)
// ═══════════════════════════════════════════════════════════
function buildSystemOutline(pptx, data, section) {
  const s = pptx.addSlide();
  addHeader(pptx, s, "Suggestion | System Outline", null);

  const content = (typeof section.content === "object" && !Array.isArray(section.content)) ? section.content : {};

  // Left zone: Server and Database
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 0.18, y: 1.25, w: 3.6, h: 5.85, fill: { color: "F8FAFF" }, line: { color: "AAAAAA", pt: 1.5, dashType: "dash" }, rectRadius: 0.1 });
  s.addText("Server and Database", { x: 0.18, y: 1.25, w: 3.6, h: 0.4, fontSize: 11, bold: true, color: C.gray, align: "center", valign: "middle" });

  // ERP block
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.78, w: 2.9, h: 0.72, fill: { color: C.teal }, line: { color: C.teal }, rectRadius: 0.07 });
  s.addText(content.erp_name || "ERP Client Site", { x: 0.5, y: 1.78, w: 2.9, h: 0.72, fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle" });

  s.addText("\u2195", { x: 1.65, y: 2.56, w: 0.6, h: 0.38, fontSize: 18, color: C.gray, align: "center" });

  // WMS Server block
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 2.98, w: 2.9, h: 0.88, fill: { color: C.headerBlue }, line: { color: C.headerBlue }, rectRadius: 0.07 });
  s.addText(content.server_name || "System\nServer", { x: 0.5, y: 2.98, w: 2.9, h: 0.88, fontSize: 12, bold: true, color: C.white, align: "center", valign: "middle" });

  // PEGASUS badge
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 0.55, y: 3.92, w: 1.3, h: 0.3, fill: { color: C.orange }, line: { color: C.orange }, rectRadius: 0.04 });
  s.addText("PEGASUS", { x: 0.55, y: 3.92, w: 1.3, h: 0.3, fontSize: 8, bold: true, color: C.white, align: "center", valign: "middle" });

  // Hardware list below server
  const hwList = toArray(content.hardware || ["Label Printer", "Warehouse Control", "Barcode Scanner"]);
  hwList.slice(0, 4).forEach((hw, i) => {
    s.addText("\u{1F5A8}  " + toStr(hw), { x: 0.3, y: 4.38 + i * 0.37, w: 3.3, h: 0.35, fontSize: 10, color: C.black });
  });

  // Right zone: Operation System
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 4.05, y: 1.25, w: 5.95, h: 5.85, fill: { color: "F8FFF8" }, line: { color: "AAAAAA", pt: 1.5, dashType: "dash" }, rectRadius: 0.1 });
  s.addShape(pptx.shapes.RECTANGLE, { x: 4.05, y: 1.25, w: 5.95, h: 0.4, fill: { color: C.teal } });
  s.addText("Operation System", { x: 4.05, y: 1.25, w: 5.95, h: 0.4, fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle" });

  const depts = toArray(content.departments || ["Materials & FG Store", "Production", "Outbound Delivery", "Office"]);
  const deptColors = [C.headerBlue, C.teal, C.orange, C.gray];
  depts.slice(0, 4).forEach((d, i) => {
    const dx = 4.3 + i * 1.45;
    const name = toStr(d);
    s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: dx, y: 1.78, w: 1.28, h: 0.9, fill: { color: "E8EEF8" }, line: { color: "C0CADD", pt: 1 }, rectRadius: 0.06 });
    s.addText("\u{1F5A5}", { x: dx, y: 1.82, w: 1.28, h: 0.45, fontSize: 18, align: "center" });
    s.addText(name, { x: dx, y: 2.25, w: 1.28, h: 0.38, fontSize: 8, color: C.black, align: "center" });
    s.addText("\u2191", { x: dx + 0.44, y: 2.72, w: 0.4, h: 0.35, fontSize: 14, color: deptColors[i] || C.gray, align: "center" });
  });

  // Connection bar
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 4.25, y: 3.18, w: 5.5, h: 0.6, fill: { color: "EEF4FF" }, line: { color: C.headerBlue, pt: 1 }, rectRadius: 0.06 });
  s.addText(content.connection_text || "PEGASUS \u2014 Real-time Inventory | ERP Integration", {
    x: 4.25, y: 3.18, w: 5.5, h: 0.6, fontSize: 10, bold: true, color: C.headerBlue, align: "center", valign: "middle",
  });

  // Hardware requirements box
  const hwReqs = toArray(content.hardware_requirements || ["Handy Terminal (Android)", "Client PC \u2014 Web browser", "Label Printer", "WiFi Access Point", "Cloud or On-premise Server"]);
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 4.25, y: 3.9, w: 5.5, h: 3.0, fill: { color: C.white }, line: { color: C.ltgray, pt: 1 }, rectRadius: 0.07 });
  hwReqs.slice(0, 5).forEach((h, i) => {
    s.addText("\u{1F4F1}  " + toStr(h), { x: 4.45, y: 4.05 + i * 0.55, w: 5.1, h: 0.48, fontSize: 10.5, color: C.black, valign: "middle" });
  });

  addImageZone(pptx, s, 10.15, 1.25, 2.95, 5.85, "SYSTEM IMAGE");
}

// ═══════════════════════════════════════════════════════════
// SLIDE — SYSTEM FLOW DIAGRAM (chevrons + detail boxes)
// ═══════════════════════════════════════════════════════════
function buildSystemFlowDiagram(pptx, data, section) {
  const s = pptx.addSlide();
  const flowTitle = (typeof section.content === "object" && section.content && section.content.bar_title) || "Process Flow";
  addHeader(pptx, s, "Suggestion | " + flowTitle, null);

  const content = section.content;
  let steps = toArray(content);
  if (typeof content === "object" && !Array.isArray(content)) {
    steps = toArray(content.steps || content.flow || content.process || content);
  }

  const barColor = (typeof content === "object" && content && content.bar_color) || C.headerBlue;
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.3, y: 1.3, w: 12.7, h: 0.42, fill: { color: barColor }, line: { color: barColor } });
  s.addText(flowTitle, { x: 0.3, y: 1.3, w: 12.7, h: 0.42, fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle" });

  const flowColors = [C.headerBlue, C.teal, C.headerBlue, C.teal, C.orange];
  const maxSteps = Math.min(steps.length, 5);
  const stepW = maxSteps <= 3 ? 3.8 : (12.4 / maxSteps - 0.15);

  steps.slice(0, maxSteps).forEach((st, i) => {
    const x = 0.5 + i * (stepW + 0.15);
    const title = typeof st === "string" ? st : (st.title || st.name || st.t || `Step ${i+1}`);
    const detail = typeof st === "object" ? toStr(st.detail || st.description || st.s || "") : "";
    const col = (typeof st === "object" && st.color) ? st.color : flowColors[i % flowColors.length];

    // Chevron: ONLY title (no detail to avoid duplication)
    s.addShape(pptx.shapes.CHEVRON, { x, y: 1.88, w: stepW, h: 0.65, fill: { color: col }, line: { color: col } });
    s.addText(title, { x: x + 0.18, y: 1.88, w: stepW - 0.36, h: 0.65, fontSize: 9.5, bold: true, color: C.white, align: "center", valign: "middle", shrinkText: true });

    // Detail box below: shows full detail text (NOT duplicated from chevron)
    const boxY = 2.68;
    const boxH = 2.55;
    s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: x + 0.05, y: boxY, w: stepW - 0.1, h: boxH, fill: { color: C.white }, line: { color: col, pt: 1 }, rectRadius: 0.07 });
    s.addShape(pptx.shapes.RECTANGLE, { x: x + 0.05, y: boxY, w: stepW - 0.1, h: 0.28, fill: { color: col } });
    s.addText(title, { x: x + 0.05, y: boxY, w: stepW - 0.1, h: 0.28, fontSize: 8, bold: true, color: C.white, align: "center", valign: "middle" });
    // Detail content
    const detailText = detail || "Details to be provided";
    s.addText(detailText, { x: x + 0.12, y: boxY + 0.32, w: stepW - 0.24, h: boxH - 0.38, fontSize: 8.5, color: C.black, valign: "top", shrinkText: true });
  });

  // Image placeholder zones at bottom
  addImageZone(pptx, s, 0.3, 5.38, 4.0, 1.6, "HANDY SCREEN");
  addImageZone(pptx, s, 4.5, 5.38, 4.0, 1.6, "PC SCREEN");
  addImageZone(pptx, s, 8.7, 5.38, 4.3, 1.6, "SAMPLE OUTPUT");
}

// ═══════════════════════════════════════════════════════════
// SLIDE — OPERATION FLOW (numbered circle steps)
// ═══════════════════════════════════════════════════════════
function buildOperationFlow(pptx, data, section) {
  const items = toArray(section.content);
  const maxItems = Math.min(items.length, 6);
  const flowColors = [C.headerBlue, C.teal, C.orange, C.headerBlue, C.teal, C.orange];

  // ──────── PAGE 1: Summary overview ────────
  const s = pptx.addSlide();
  addHeader(pptx, s, "Suggestion | Explaining Operation Flow", null);

  s.addText("Summary operation flow", { x: 0.3, y: 1.28, w: 9, h: 0.35, fontSize: 13, bold: true, color: C.headerBlue });
  s.addShape(pptx.shapes.LINE, { x: 0.3, y: 1.65, w: 9, h: 0, line: { color: C.headerBlue, pt: 1.5 } });

  // Chevron flow bar at top (compact summary of all steps)
  const chevronW = maxItems > 0 ? Math.min(2.0, 12.4 / maxItems - 0.12) : 2.0;
  items.slice(0, maxItems).forEach((f, i) => {
    const title = typeof f === "string" ? f : (f.title || f.name || f.t || `Step ${i+1}`);
    const x = 0.3 + i * (chevronW + 0.12);
    const col = flowColors[i % flowColors.length];
    s.addShape(pptx.shapes.CHEVRON, { x, y: 1.78, w: chevronW, h: 0.55, fill: { color: col }, line: { color: col } });
    s.addText(title, { x: x + 0.15, y: 1.78, w: chevronW - 0.3, h: 0.55, fontSize: 8.5, bold: true, color: C.white, align: "center", valign: "middle", shrinkText: true });
  });

  // Summary list below
  const listStartY = 2.55;
  const totalH = 4.25;
  const itemH = Math.min(1.0, totalH / Math.max(maxItems, 1));

  items.slice(0, maxItems).forEach((f, i) => {
    const y = listStartY + i * itemH;
    const title = typeof f === "string" ? f : (f.title || f.name || f.t || `Step ${i+1}`);
    const desc = typeof f === "object" ? toStr(f.description || f.detail || f.d || "") : "";
    const col = flowColors[i % flowColors.length];

    s.addShape(pptx.shapes.OVAL, { x: 0.3, y: y + 0.03, w: 0.35, h: 0.35, fill: { color: col } });
    s.addText(String(i+1), { x: 0.3, y: y + 0.03, w: 0.35, h: 0.35, fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle" });
    s.addText(title, { x: 0.78, y, w: 8.5, h: 0.32, fontSize: 11, bold: true, color: C.black, valign: "middle" });
    if (desc && itemH > 0.5) {
      s.addText(desc, { x: 0.78, y: y + 0.33, w: 8.5, h: itemH - 0.38, fontSize: 9, color: C.gray, valign: "top", shrinkText: true });
    }
  });

  addImageZone(pptx, s, 9.8, 2.55, 3.2, 4.25, "FLOW DIAGRAM");

  // ──────── PAGE 2+: Individual detail pages per step ────────
  items.slice(0, maxItems).forEach((f, i) => {
    const title = typeof f === "string" ? f : (f.title || f.name || f.t || `Step ${i+1}`);
    const desc = typeof f === "object" ? toStr(f.description || f.detail || f.d || "") : "";
    const points = typeof f === "object" ? toArray(f.points || f.bullets || f.items || []) : [];
    const col = flowColors[i % flowColors.length];

    const ds = pptx.addSlide();
    addHeader(pptx, ds, `Operation flow Detail | ${title}`, null);

    // Top: chevron flow bar showing all steps (highlight current)
    items.slice(0, maxItems).forEach((st, si) => {
      const stTitle = typeof st === "string" ? st : (st.title || st.name || st.t || `Step ${si+1}`);
      const stCol = si === i ? flowColors[si % flowColors.length] : "CCCCCC";
      const sx = 0.3 + si * (chevronW + 0.12);
      ds.addShape(pptx.shapes.CHEVRON, { x: sx, y: 1.22, w: chevronW, h: 0.48, fill: { color: stCol }, line: { color: stCol } });
      ds.addText(stTitle, { x: sx + 0.15, y: 1.22, w: chevronW - 0.3, h: 0.48, fontSize: 8, bold: si === i, color: C.white, align: "center", valign: "middle", shrinkText: true });
    });

    // Left side: description content box
    ds.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 0.3, y: 1.88, w: 6.2, h: 4.8, fill: { color: C.white }, line: { color: col, pt: 1.2 }, rectRadius: 0.08 });

    // Step title inside box
    ds.addText("Operation Flow — " + title, { x: 0.5, y: 1.95, w: 5.8, h: 0.38, fontSize: 12, bold: true, color: C.black, valign: "middle" });

    // Description text
    let contentY = 2.4;
    if (desc) {
      ds.addText(desc, { x: 0.5, y: contentY, w: 5.8, h: 1.5, fontSize: 10.5, color: C.black, valign: "top", shrinkText: true });
      contentY += 1.6;
    }

    // Bullet points
    points.slice(0, 6).forEach((pt, pi) => {
      if (contentY > 6.2) return;
      ds.addText("•  " + toStr(pt), { x: 0.6, y: contentY, w: 5.6, h: 0.32, fontSize: 10, color: "444444", valign: "top" });
      contentY += 0.35;
    });

    // Right side: image placeholder zones (2 zones for screenshots)
    addImageZone(pptx, ds, 6.8, 1.88, 6.2, 2.25, title + " — Screenshot 1");
    addImageZone(pptx, ds, 6.8, 4.28, 6.2, 2.4, title + " — Screenshot 2");
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE — SCOPE OF WORK (numbered items)
// ═══════════════════════════════════════════════════════════
function buildScopeOfWork(pptx, data, section) {
  addSectionDivider(pptx, "SCOPE OF WORK", null);

  const items = toArray(section.content);
  // Split into pages if too many items (max ~6 items per page)
  const itemsPerPage = 6;
  const pages = [];
  for (let p = 0; p < items.length; p += itemsPerPage) {
    pages.push(items.slice(p, p + itemsPerPage));
  }
  if (pages.length === 0) pages.push([]);

  pages.forEach((pageItems, pageIdx) => {
    const s = pptx.addSlide();
    const pageLabel = pages.length > 1 ? ` (${pageIdx + 1}/${pages.length})` : "";
    addHeader(pptx, s, "Scope of Work - Details" + pageLabel, null);

    let yPos = 1.3;
    const baseIndex = pageIdx * itemsPerPage;

    pageItems.forEach((item, i) => {
      if (yPos > 6.3) return; // Safety limit
      const text = typeof item === "string" ? item : (item.title || item.name || toStr(item));
      const desc = typeof item === "object" ? toStr(item.description || item.detail || "") : "";

      // Number badge
      s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 0.3, y: yPos, w: 0.38, h: 0.38, fill: { color: C.headerBlue }, rectRadius: 0.05 });
      s.addText(String(baseIndex + i + 1), { x: 0.3, y: yPos, w: 0.38, h: 0.38, fontSize: 13, bold: true, color: C.white, align: "center", valign: "middle" });
      // Title
      s.addText(text, { x: 0.82, y: yPos, w: 11.5, h: 0.38, fontSize: 12, bold: true, color: C.black, valign: "middle" });
      yPos += 0.42;

      // Description (with proper gap from title)
      if (desc) {
        s.addText(desc, { x: 0.82, y: yPos, w: 11.5, h: 0.32, fontSize: 10.5, color: C.gray, valign: "top" });
        yPos += 0.36;
      }
      // Separator line between items
      if (i < pageItems.length - 1) {
        s.addShape(pptx.shapes.LINE, { x: 0.3, y: yPos + 0.06, w: 12.3, h: 0, line: { color: "E0E0E0", pt: 0.5 } });
        yPos += 0.18;
      }
    });
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE — TIMELINE (chevron phases)
// ═══════════════════════════════════════════════════════════
function buildTimeline(pptx, data, section) {
  const s = pptx.addSlide();
  addHeader(pptx, s, "Timeline & Milestones", null);

  const phases = toArray(section.content);
  const colors = [C.headerBlue, C.teal, C.orange, C.headerBlue, C.teal, C.orange, C.green, C.navy];
  const maxPhases = Math.min(phases.length, 8);

  // Dynamic layout: 1 or 2 rows depending on count
  const cols = maxPhases <= 4 ? maxPhases : 4;
  const rows = Math.ceil(maxPhases / cols);
  // Calculate column width dynamically
  const totalW = 12.7;
  const arrowW = 0.28;
  const colW = (totalW - (cols - 1) * arrowW) / cols;
  // Row heights
  const rowGap = rows > 1 ? 2.65 : 0;
  const boxH = 0.52;

  phases.slice(0, maxPhases).forEach((phase, i) => {
    const row = Math.floor(i / cols);
    const col = i % cols;
    const x = 0.3 + col * (colW + arrowW);
    const y = 1.4 + row * rowGap;
    const color = colors[i % colors.length];
    const text = typeof phase === "string" ? phase : (phase.name || phase.title || phase.phase || `Phase ${i+1}`);
    const duration = typeof phase === "object" ? toStr(phase.duration || phase.period || "") : "";
    const tasks = typeof phase === "object" ? toArray(phase.tasks || phase.details || phase.items || []) : [];

    // Phase box
    s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x, y, w: colW, h: boxH, fill: { color }, rectRadius: 0.06 });
    s.addText(text, { x, y, w: colW, h: boxH, fontSize: 10.5, bold: true, color: C.white, align: "center", valign: "middle" });

    // Duration below box
    let detailY = y + boxH + 0.06;
    if (duration) {
      s.addText(duration, { x, y: detailY, w: colW, h: 0.26, fontSize: 9.5, color: C.gray, align: "center", valign: "top" });
      detailY += 0.28;
    }

    // Task list (max 4 tasks, only show if space allows)
    const maxTasks = rows > 1 ? 3 : 5;
    tasks.slice(0, maxTasks).forEach((t, ti) => {
      if (detailY > y + 2.4) return; // Overflow guard
      s.addText("• " + toStr(t), { x: x + 0.05, y: detailY, w: colW - 0.1, h: 0.22, fontSize: 8.5, color: "555555", valign: "top" });
      detailY += 0.24;
    });

    // Arrow between columns (not after last column in row)
    if (col < cols - 1 && i < maxPhases - 1) {
      s.addText("\u25B6", { x: x + colW, y, w: arrowW, h: boxH, fontSize: 13, color, align: "center", valign: "middle" });
    }
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE — TECHNOLOGY STACK (grid with icons)
// ═══════════════════════════════════════════════════════════
function buildTechStack(pptx, data, section) {
  const s = pptx.addSlide();
  addHeader(pptx, s, "Technology Stack", null);

  let items = toArray(section.content);
  // Handle case where content is a flat object {key: value}
  if (items.length === 0 && typeof section.content === "object" && !Array.isArray(section.content)) {
    items = Object.entries(section.content).map(([k, v]) => ({ name: k, description: toStr(v) }));
  }

  items.slice(0, 12).forEach((item, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.3 + col * 6.5;
    const y = 1.4 + row * 0.85;
    const text = typeof item === "string" ? item : (item.name || item.technology || toStr(item));
    const desc = typeof item === "object" ? toStr(item.description || item.purpose || "") : "";

    s.addShape(pptx.shapes.OVAL, { x, y: y + 0.05, w: 0.4, h: 0.4, fill: { color: C.teal } });
    s.addText("\u2699", { x, y: y + 0.05, w: 0.4, h: 0.4, fontSize: 14, color: C.white, align: "center", valign: "middle" });
    s.addText(text, { x: x + 0.5, y, w: 5.5, h: 0.35, fontSize: 12, bold: true, color: C.black, valign: "middle" });
    if (desc) s.addText(desc, { x: x + 0.5, y: y + 0.35, w: 5.5, h: 0.3, fontSize: 10, color: C.gray });
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE — TEAM STRUCTURE (org chart cards)
// ═══════════════════════════════════════════════════════════
function buildTeamStructure(pptx, data, section) {
  const s = pptx.addSlide();
  addHeader(pptx, s, "Team Structure", null);

  const roles = toArray(section.content);

  roles.slice(0, 9).forEach((role, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.5 + col * 4.2;
    const y = 1.5 + row * 1.8;
    const text = typeof role === "string" ? role : (role.role || role.title || role.name || toStr(role));
    const count = typeof role === "object" ? (role.count ? `${role.count} person(s)` : "") : "";
    const desc = typeof role === "object" ? toStr(role.responsibilities || role.description || "") : "";
    const descText = count ? (count + (desc ? " — " + desc : "")) : desc;

    s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x, y, w: 3.8, h: 1.4, fill: { color: "F8F9FA" }, line: { color: C.headerBlue, pt: 1.5 }, rectRadius: 0.08 });
    s.addShape(pptx.shapes.RECTANGLE, { x: x + 0.05, y: y + 0.05, w: 3.7, h: 0.4, fill: { color: C.headerBlue } });
    s.addText(text, { x: x + 0.05, y: y + 0.05, w: 3.7, h: 0.4, fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle" });
    if (descText) s.addText(descText, { x: x + 0.15, y: y + 0.5, w: 3.5, h: 0.8, fontSize: 9, color: C.gray, valign: "top" });
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE — PRICING TABLE
// ═══════════════════════════════════════════════════════════
function buildPricing(pptx, data, section) {
  const content = section.content;

  // Simple string content
  if (typeof content === "string") {
    const s = pptx.addSlide();
    addHeader(pptx, s, "Pricing Estimate", null);
    s.addText(content, { x: 0.5, y: 1.4, w: 12.3, h: 5.5, fontSize: 13, color: "333333", valign: "top", lineSpacing: 22 });
    return;
  }

  const rows = toArray(content);
  if (rows.length === 0) {
    const s = pptx.addSlide();
    addHeader(pptx, s, "Pricing Estimate", null);
    s.addText(toText(content, "Pricing details to be confirmed."), { x: 0.5, y: 1.4, w: 12.3, h: 5.5, fontSize: 13, color: "333333", valign: "top" });
    return;
  }

  // Extract total before pagination
  const total = (typeof content === "object" && !Array.isArray(content)) ? content.total : null;

  // Paginate: max 10 data rows per page (header + 10 rows + optional total = 12 × 0.45 = 5.4")
  const rowsPerPage = 10;
  const pages = [];
  for (let p = 0; p < rows.length; p += rowsPerPage) {
    pages.push(rows.slice(p, p + rowsPerPage));
  }

  const headerRow = [
    { text: "#", options: { fill: C.headerBlue, color: C.white, bold: true, fontSize: 10, align: "center" } },
    { text: "Item", options: { fill: C.headerBlue, color: C.white, bold: true, fontSize: 10 } },
    { text: "Description", options: { fill: C.headerBlue, color: C.white, bold: true, fontSize: 10 } },
    { text: "Estimate", options: { fill: C.headerBlue, color: C.white, bold: true, fontSize: 10, align: "right" } },
  ];

  pages.forEach((pageRows, pageIdx) => {
    const s = pptx.addSlide();
    const pageLabel = pages.length > 1 ? ` (${pageIdx + 1}/${pages.length})` : "";
    addHeader(pptx, s, "Pricing Estimate" + pageLabel, null);

    const tableData = [headerRow];
    const baseIndex = pageIdx * rowsPerPage;

    pageRows.forEach((row, i) => {
      const item = typeof row === "string" ? row : toStr(row.item || row.name || row.service || "");
      const desc = typeof row === "object" ? toStr(row.description || row.detail || "") : "";
      const price = typeof row === "object" ? toStr(row.price || row.estimate || row.cost || row.amount || "") : "";
      const bg = (baseIndex + i) % 2 === 0 ? "FFFFFF" : "F5F7FA";

      tableData.push([
        { text: String(baseIndex + i + 1), options: { fill: bg, fontSize: 10, align: "center" } },
        { text: item, options: { fill: bg, fontSize: 10 } },
        { text: desc, options: { fill: bg, fontSize: 9, color: C.gray } },
        { text: price, options: { fill: bg, fontSize: 10, align: "right", bold: true } },
      ]);
    });

    // Total row only on the last page
    if (total && pageIdx === pages.length - 1) {
      tableData.push([
        { text: "", options: { fill: C.navy } },
        { text: "TOTAL", options: { fill: C.navy, color: C.white, bold: true, fontSize: 11 } },
        { text: "", options: { fill: C.navy } },
        { text: toStr(total), options: { fill: C.navy, color: C.orange, bold: true, fontSize: 12, align: "right" } },
      ]);
    }

    s.addTable(tableData, { x: 0.3, y: 1.4, w: 12.7, border: { pt: 0.5, color: "CCCCCC" }, colW: [0.6, 3.5, 5.5, 3.1], rowH: 0.45 });
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE — TERMS & CONDITIONS (TABLE FORMAT)
// ═══════════════════════════════════════════════════════════
function buildTerms(pptx, data, section) {
  const content = section.content;

  // Extract terms as array of {title, detail} pairs
  let terms = [];
  if (Array.isArray(content)) {
    terms = content.map((item, i) => {
      if (typeof item === "string") return { title: `Item ${i+1}`, detail: item };
      return {
        title: toStr(item.title || item.name || item.term || item.item || `Item ${i+1}`),
        detail: toStr(item.detail || item.description || item.content || item.value || item.condition || ""),
      };
    });
  } else if (typeof content === "object" && content) {
    // Object format: { "Payment": "50/50", "Warranty": "1 year", ... }
    // Or nested with arrays
    const arr = toArray(content);
    if (arr.length > 0) {
      terms = arr.map((item, i) => {
        if (typeof item === "string") return { title: `Item ${i+1}`, detail: item };
        return {
          title: toStr(item.title || item.name || item.term || item.item || `Item ${i+1}`),
          detail: toStr(item.detail || item.description || item.content || item.value || item.condition || ""),
        };
      });
    } else {
      // Direct key-value pairs
      for (const [k, v] of Object.entries(content)) {
        terms.push({ title: k, detail: toStr(v) });
      }
    }
  } else if (typeof content === "string") {
    // Try to parse lines as terms
    const lines = content.split("\n").filter(l => l.trim());
    terms = lines.map((line, i) => {
      const colonIdx = line.indexOf(":");
      if (colonIdx > 0 && colonIdx < 50) {
        return { title: line.substring(0, colonIdx).replace(/^\d+[\.\)]\s*/, "").trim(), detail: line.substring(colonIdx + 1).trim() };
      }
      return { title: `Item ${i+1}`, detail: line.replace(/^\d+[\.\)]\s*/, "").trim() };
    });
  }

  if (terms.length === 0) {
    terms = [{ title: "Terms", detail: "To be confirmed." }];
  }

  // Paginate: max 8 terms per page
  const termsPerPage = 8;
  const pages = [];
  for (let p = 0; p < terms.length; p += termsPerPage) {
    pages.push(terms.slice(p, p + termsPerPage));
  }

  const headerRow = [
    { text: "#", options: { fill: C.headerBlue, color: C.white, bold: true, fontSize: 10, align: "center" } },
    { text: "Item", options: { fill: C.headerBlue, color: C.white, bold: true, fontSize: 10 } },
    { text: "Details", options: { fill: C.headerBlue, color: C.white, bold: true, fontSize: 10 } },
  ];

  pages.forEach((pageTerms, pageIdx) => {
    const s = pptx.addSlide();
    const pageLabel = pages.length > 1 ? ` (${pageIdx + 1}/${pages.length})` : "";
    addHeader(pptx, s, "Terms & Conditions" + pageLabel, null);

    const tableData = [headerRow];
    const baseIndex = pageIdx * termsPerPage;

    pageTerms.forEach((term, i) => {
      const bg = (baseIndex + i) % 2 === 0 ? "FFFFFF" : "F5F7FA";
      tableData.push([
        { text: String(baseIndex + i + 1), options: { fill: bg, fontSize: 10, align: "center" } },
        { text: term.title, options: { fill: bg, fontSize: 10, bold: true } },
        { text: term.detail, options: { fill: bg, fontSize: 10, color: "333333" } },
      ]);
    });

    s.addTable(tableData, {
      x: 0.3, y: 1.4, w: 12.7,
      border: { pt: 0.5, color: "CCCCCC" },
      colW: [0.6, 3.2, 8.9],
      rowH: 0.55,
    });
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE — MAINTENANCE / SLA TABLE
// ═══════════════════════════════════════════════════════════
function buildMaintenance(pptx, data, section) {
  const items = toArray(section.content);
  if (items.length === 0) {
    const s = pptx.addSlide();
    addHeader(pptx, s, "System Maintenance & Support", null);
    s.addText(toText(section.content, "Maintenance details to be confirmed."), {
      x: 0.5, y: 1.4, w: 12.3, h: 5.5, fontSize: 12, color: "333333", valign: "top", lineSpacing: 20,
    });
    return;
  }

  // Split into pages: max 10 data rows per page (header row + 10 = 11 rows × 0.45 = 4.95")
  const rowsPerPage = 10;
  const pages = [];
  for (let p = 0; p < items.length; p += rowsPerPage) {
    pages.push(items.slice(p, p + rowsPerPage));
  }

  const headerRow = [
    { text: "Item", options: { fill: C.headerBlue, color: C.white, bold: true, fontSize: 10 } },
    { text: "Year 1", options: { fill: C.headerBlue, color: C.white, bold: true, fontSize: 10, align: "center" } },
    { text: "Year 2+", options: { fill: C.headerBlue, color: C.white, bold: true, fontSize: 10, align: "center" } },
    { text: "Notes", options: { fill: C.headerBlue, color: C.white, bold: true, fontSize: 10 } },
  ];

  pages.forEach((pageItems, pageIdx) => {
    const s = pptx.addSlide();
    const pageLabel = pages.length > 1 ? ` (${pageIdx + 1}/${pages.length})` : "";
    addHeader(pptx, s, "System Maintenance & Support" + pageLabel, null);

    const tableData = [headerRow];
    const baseIndex = pageIdx * rowsPerPage;

    pageItems.forEach((item, i) => {
      const name = typeof item === "string" ? item : toStr(item.name || item.item || item.service || "");
      const y1 = typeof item === "object" ? toStr(item.year1 || item.warranty || "Included") : "";
      const y2 = typeof item === "object" ? toStr(item.year2 || item.renewal || "") : "";
      const notes = typeof item === "object" ? toStr(item.notes || item.description || "") : "";
      const bg = (baseIndex + i) % 2 === 0 ? "FFFFFF" : "F5F7FA";

      tableData.push([
        { text: name, options: { fill: bg, fontSize: 10 } },
        { text: y1, options: { fill: bg, fontSize: 10, align: "center" } },
        { text: y2, options: { fill: bg, fontSize: 10, align: "center" } },
        { text: notes, options: { fill: bg, fontSize: 9, color: C.gray } },
      ]);
    });

    s.addTable(tableData, { x: 0.3, y: 1.4, w: 12.7, border: { pt: 0.5, color: "CCCCCC" }, colW: [4.0, 2.5, 2.5, 3.7], rowH: 0.45 });
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE — FUNCTION LIST TABLE
// ═══════════════════════════════════════════════════════════
function buildFunctionList(pptx, data, section) {
  const content = section.content;

  // Support two formats:
  // A) Array of functions: [{category, name, description}, ...]
  // B) Object with categories: {Inbound: [{name, description}], Outbound: [...]}
  let allRows = [];
  let categories = [];

  if (typeof content === "object" && !Array.isArray(content)) {
    // Format B: grouped by category keys
    for (const [catName, catItems] of Object.entries(content)) {
      if (Array.isArray(catItems)) {
        categories.push(catName);
        catItems.forEach((fn, i) => {
          allRows.push({
            category: catName,
            isCategoryHeader: i === 0,
            num: i + 1,
            name: typeof fn === "string" ? fn : toStr(fn.name || fn.function_name || fn.title || ""),
            desc: typeof fn === "object" ? toStr(fn.description || fn.detail || "") : "",
          });
        });
      }
    }
  }

  // Fallback to flat array
  if (allRows.length === 0) {
    const funcs = toArray(content);
    funcs.forEach((fn, i) => {
      const cat = typeof fn === "object" ? toStr(fn.category || fn.module || "") : "";
      allRows.push({
        category: cat,
        isCategoryHeader: false,
        num: i + 1,
        name: typeof fn === "string" ? fn : toStr(fn.name || fn.function_name || fn.title || ""),
        desc: typeof fn === "object" ? toStr(fn.description || fn.detail || "") : "",
      });
    });
  }

  // Paginate: max 12 data rows per page
  const rowsPerPage = 12;
  const pages = [];
  for (let p = 0; p < allRows.length; p += rowsPerPage) {
    pages.push(allRows.slice(p, p + rowsPerPage));
  }
  if (pages.length === 0) pages.push([]);

  const headerRow = [
    { text: "#", options: { fill: C.headerBlue, color: C.white, bold: true, fontSize: 9, align: "center" } },
    { text: "Category", options: { fill: C.headerBlue, color: C.white, bold: true, fontSize: 9 } },
    { text: "Function", options: { fill: C.headerBlue, color: C.white, bold: true, fontSize: 9 } },
    { text: "Description", options: { fill: C.headerBlue, color: C.white, bold: true, fontSize: 9 } },
  ];

  pages.forEach((pageRows, pageIdx) => {
    const s = pptx.addSlide();
    const pageLabel = pages.length > 1 ? ` (${pageIdx + 1}/${pages.length})` : "";
    addHeader(pptx, s, "Function | List" + pageLabel, null);

    const tableData = [headerRow];
    let prevCat = "";

    pageRows.forEach((row) => {
      // Insert category header row when category changes
      if (row.category && row.category !== prevCat) {
        tableData.push([
          { text: "#", options: { fill: "0088CC", color: C.white, bold: true, fontSize: 8.5, align: "center" } },
          { text: row.category, options: { fill: "0088CC", color: C.white, bold: true, fontSize: 9 } },
          { text: row.category, options: { fill: "0088CC", color: C.white, bold: true, fontSize: 9 } },
          { text: "", options: { fill: "0088CC" } },
        ]);
        prevCat = row.category;
      }

      const bg = row.num % 2 === 0 ? "F5F7FA" : "FFFFFF";
      tableData.push([
        { text: String(row.num), options: { fill: bg, fontSize: 8.5, align: "center" } },
        { text: row.name, options: { fill: bg, fontSize: 8.5 } },
        { text: row.name, options: { fill: bg, fontSize: 8.5, bold: true } },
        { text: row.desc, options: { fill: bg, fontSize: 8.5, color: C.gray } },
      ]);
    });

    s.addTable(tableData, { x: 0.3, y: 1.35, w: 12.7, border: { pt: 0.5, color: "CCCCCC" }, colW: [0.45, 2.5, 3.5, 6.25], rowH: 0.35 });
  });

  // ──── Web Screen Pages: one per category ────
  const screensContent = typeof content === "object" && !Array.isArray(content) ? content : null;
  const screenCategories = screensContent
    ? Object.keys(screensContent).filter(k => Array.isArray(screensContent[k]) && screensContent[k].length > 0)
    : categories.length > 0 ? [...new Set(categories)] : [];

  // Also check for explicit screen_pages in content
  const screenPages = (typeof content === "object" && content && content.screen_pages) ? toArray(content.screen_pages) : [];

  // Generate screen pages from categories
  const screensToRender = screenPages.length > 0
    ? screenPages
    : screenCategories.map(cat => ({ title: cat, description: `${cat} functions overview` }));

  screensToRender.forEach((screen) => {
    const screenTitle = typeof screen === "string" ? screen : (screen.title || screen.name || toStr(screen));
    const screenDesc = typeof screen === "object" ? toStr(screen.description || "") : "";
    buildScreenPage(pptx, screenTitle, screenDesc);
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE — WEB SCREEN PAGE (image placeholder for each function area)
// ═══════════════════════════════════════════════════════════
function buildScreenPage(pptx, title, description) {
  const s = pptx.addSlide();
  addHeader(pptx, s, title + " | Screen", null);

  // Description text
  if (description) {
    s.addText(description, { x: 0.3, y: 1.2, w: 12.7, h: 0.35, fontSize: 10, color: C.gray, valign: "middle" });
  }

  // Large image placeholder zone (main screenshot area)
  const imgY = description ? 1.65 : 1.3;
  const imgH = description ? 5.2 : 5.55;

  // PEGASUS badge
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 0.3, y: imgY, w: 1.3, h: 0.3, fill: { color: C.orange }, rectRadius: 0.04 });
  s.addText("PEGASUS", { x: 0.3, y: imgY, w: 1.3, h: 0.3, fontSize: 8, bold: true, color: C.white, align: "center", valign: "middle" });

  // Main screenshot placeholder
  addImageZone(pptx, s, 0.3, imgY + 0.4, 12.7, imgH - 0.4, title + " — Web Screen");
}

// ═══════════════════════════════════════════════════════════
// SLIDE — GENERIC FALLBACK (for unknown sections)
// ═══════════════════════════════════════════════════════════
function buildGenericSlide(pptx, data, section, slideNum) {
  const s = pptx.addSlide();
  const title = section.section_title || section.title || `Section ${slideNum}`;
  addHeader(pptx, s, title, slideNum);
  const text = toText(section.content, "");
  s.addText(text.substring(0, 3000), { x: 0.5, y: 1.3, w: 12.3, h: 5.7, fontSize: 12, color: "333333", valign: "top", lineSpacing: 20, paraSpaceAfter: 6 });
}

// ═══════════════════════════════════════════════════════════
// SLIDE — THANK YOU
// ═══════════════════════════════════════════════════════════
function buildThankYouSlide(pptx, data) {
  const s = pptx.addSlide();
  s.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 13.3, h: 7.5, fill: { color: C.navy } });
  s.addText("Thank You", { x: 0, y: 2.0, w: 13.3, h: 1.2, fontSize: 48, bold: true, color: C.white, align: "center", valign: "middle" });
  s.addText("TOMAS TECH CO., LTD.", { x: 0, y: 3.5, w: 13.3, h: 0.6, fontSize: 20, color: C.orange, align: "center", valign: "middle" });
  s.addText(
    "7/1(3C) Udomsuk 46 Alley, Khwaeng Bang Na Nuea, Khet Bang Na, Bangkok 10260\n" +
    "Tel: +66-98-271-9741  |  Email: info@tomastc.com  |  www.tomastc.com",
    { x: 1.5, y: 4.5, w: 10.3, h: 1.0, fontSize: 11, color: C.white, align: "center", lineSpacing: 18 }
  );
  addFooter(pptx, s, null);
}

// ═══════════════════════════════════════════════════════════
// MAIN: Generate PPTX from proposal JSON
// ═══════════════════════════════════════════════════════════
async function generatePPTX(proposalData) {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "TOMAS TECH CO., LTD.";
  pptx.company = "TOMAS TECH CO., LTD.";
  pptx.subject = proposalData.proposal_title || "Proposal";

  let data = proposalData;
  if (typeof data === "string") {
    try {
      const cleaned = data.replace(/```json\s*/g, "").replace(/```\s*/g, "").trim();
      const match = cleaned.match(/\{[\s\S]*\}/);
      if (match) data = JSON.parse(match[0]);
    } catch (e) {
      data = { proposal_title: "Proposal", client_name: "Client", sections: [] };
    }
  }

  // 1. Cover slide
  buildCoverSlide(pptx, data);

  // 2. Process sections
  const sections = data.sections || [];
  let slideNum = 2;

  // Section builder lookup — ordered by priority
  const sectionBuilders = [
    { kw: ["project purpose", "purpose", "proposal overview"], fn: buildProjectPurpose },
    { kw: ["executive", "summary", "บทสรุป"], fn: buildExecutiveSummary },
    { kw: ["pain point", "pain", "problem", "challenge", "ปัญหา", "current issue", "current situation"], fn: buildPainPoint },
    { kw: ["benefit", "target", "improvement", "advantage", "ประโยชน์"], fn: buildBenefitsOverview },
    { kw: ["system outline", "outline", "architecture", "infrastructure", "โครงสร้างระบบ"], fn: buildSystemOutline },
    { kw: ["system flow", "flow diagram", "process flow", "inbound", "outbound"], fn: buildSystemFlowDiagram },
    { kw: ["operation flow", "operation", "explaining", "ขั้นตอน"], fn: buildOperationFlow },
    { kw: ["function list", "function", "feature list", "ฟังก์ชัน"], fn: buildFunctionList },
    { kw: ["scope", "ขอบเขต"], fn: buildScopeOfWork },
    { kw: ["timeline", "milestone", "schedule", "go live", "phase", "แผนงาน", "ระยะเวลา"], fn: buildTimeline },
    { kw: ["technology", "tech_stack", "stack", "tech", "เทคโนโลยี"], fn: buildTechStack },
    { kw: ["team", "structure", "resource", "organization", "ทีมงาน", "โครงสร้าง"], fn: buildTeamStructure },
    { kw: ["pricing", "price", "cost", "budget", "estimate", "quotation", "ราคา", "ประมาณการ"], fn: buildPricing },
    { kw: ["maintenance", "support", "sla", "service level", "warranty", "บำรุงรักษา"], fn: buildMaintenance },
    { kw: ["terms", "condition", "payment", "เงื่อนไข", "ข้อกำหนด"], fn: buildTerms },
  ];

  sections.forEach((section) => {
    const title = (section.section_title || section.title || "").toLowerCase();

    // Section divider
    if (section.type === "divider" || title.includes("[divider]")) {
      const divTitle = (section.divider_title || title.replace("[divider]", "").replace(/^\d+\.\s*/, "")).trim().toUpperCase() || "SECTION";
      addSectionDivider(pptx, divTitle, slideNum);
      slideNum++;
      return;
    }

    // Match section to builder
    let built = false;
    for (const { kw, fn } of sectionBuilders) {
      if (kw.some(k => title.includes(k))) {
        fn(pptx, data, section);
        built = true;
        break;
      }
    }

    if (!built) buildGenericSlide(pptx, data, section, slideNum);
    slideNum++;
  });

  // Last: Thank you slide
  buildThankYouSlide(pptx, data);

  return await pptx.write({ outputType: "nodebuffer" });
}

// ═══════════════════════════════════════════════════════════
// API Handler
// ═══════════════════════════════════════════════════════════
module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  try {
    const input = req.body;
    const proposalContent = input.proposal_content || input.output || input.text || input;

    if (!proposalContent) return res.status(400).json({ error: "proposal_content is required" });

    let proposalData = proposalContent;
    if (typeof proposalData === "string") {
      try {
        const cleaned = proposalData.replace(/```json\s*/g, "").replace(/```\s*/g, "").trim();
        const match = cleaned.match(/\{[\s\S]*\}/);
        if (match) proposalData = JSON.parse(match[0]);
      } catch (e) { /* keep as string */ }
    }

    const requestId = input.request_id;
    const clientName = input.client_name || (typeof proposalData === "object" ? proposalData.client_name : null);

    const buffer = await generatePPTX(proposalData);

    const safeName = (clientName || "proposal").replace(/[^a-zA-Z0-9\u0E00-\u0E7F]/g, "_").substring(0, 30);
    const timestamp = new Date().toISOString().slice(0, 10).replace(/-/g, "");
    const filename = `TOMAS_TECH_Proposal_${safeName}_${timestamp}.pptx`;

    if (SUPABASE_KEY) {
      try {
        const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);
        const { error: uploadError } = await supabase.storage.from("proposals").upload(`pptx/${filename}`, buffer, {
          contentType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
          upsert: true,
        });

        if (uploadError) {
          console.error("Upload error:", uploadError);
          res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
          res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
          return res.status(200).send(buffer);
        }

        const { data: urlData } = supabase.storage.from("proposals").getPublicUrl(`pptx/${filename}`);

        if (requestId) {
          await supabase.from("proposal_requests").update({ pptx_url: urlData.publicUrl, status: "completed" }).eq("id", requestId);
        }

        return res.status(200).json({ success: true, filename, pptx_url: urlData.publicUrl, request_id: requestId || null });
      } catch (storageError) {
        console.error("Storage error:", storageError);
      }
    }

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    return res.status(200).send(buffer);

  } catch (error) {
    console.error("PPTX generation error:", error);
    return res.status(500).json({ error: "Failed to generate PPTX", details: error.message });
  }
};
