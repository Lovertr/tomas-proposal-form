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
// HELPERS
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
  if (n) {
    s.addText(String(n), {
      x: 12.6, y: 7.0, w: 0.5, h: 0.25,
      fontSize: 10, color: C.gray, align: "right",
    });
  }
  s.addText("Tomas Tech Co., Ltd.", {
    x: 0.3, y: 7.0, w: 3, h: 0.25, fontSize: 9, color: C.gray,
  });
}

function addPegasusLogo(pptx, s) {
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.18, y: 0.1, w: 1.5, h: 0.38,
    fill: { color: C.headerBlue }, line: { color: C.headerBlue }, rectRadius: 0.04,
  });
  s.addText("PEGASUS", {
    x: 0.18, y: 0.1, w: 1.5, h: 0.38,
    fontSize: 10, bold: true, color: C.white, align: "center", valign: "middle",
  });
}

function addHeader(pptx, s, title, n) {
  addPegasusLogo(pptx, s);
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 0.57, w: 13.3, h: 0.06,
    fill: { color: C.headerBlue }, line: { color: C.headerBlue },
  });
  s.addText("     " + title, {
    x: 0, y: 0.65, w: 13.3, h: 0.45,
    fontSize: 16, bold: true, color: C.headerBlue, valign: "middle",
  });
  s.addShape(pptx.shapes.LINE, {
    x: 0.18, y: 1.1, w: 12.94, h: 0,
    line: { color: C.headerBlue, pt: 2 },
  });
  addFooter(pptx, s, n);
}

function addSectionDivider(pptx, s, title, n) {
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

// Image placeholder zone (dashed border with landscape icon)
function addImageZone(pptx, s, x, y, w, h, label, hint) {
  s.addShape(pptx.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: "F0F4F8" }, line: { color: "AABBCC", pt: 1.5, dashType: "dash" },
  });
  s.addText(label, {
    x, y: y + h / 2 - 0.2, w, h: 0.28,
    fontSize: 9, bold: true, color: "4A6FA5", align: "center", valign: "middle",
  });
  if (hint) {
    s.addText(hint, {
      x, y: y + h / 2 + 0.1, w, h: 0.2,
      fontSize: 7.5, color: "8899AA", align: "center", italic: true,
    });
  }
}

// Safe text extraction from various content shapes
function extractText(content, fallback) {
  if (!content) return fallback || "";
  if (typeof content === "string") return content;
  if (content.summary) return content.summary;
  if (content.description) return content.description;
  if (content.text) return content.text;
  return JSON.stringify(content);
}

function extractArray(content) {
  if (!content) return [];
  if (Array.isArray(content)) return content;
  if (content.items) return content.items;
  if (content.scope_items) return content.scope_items;
  if (content.phases) return content.phases;
  if (content.milestones) return content.milestones;
  if (content.roles) return content.roles;
  if (content.team) return content.team;
  if (content.members) return content.members;
  if (content.technologies) return content.technologies;
  if (content.stack) return content.stack;
  if (content.line_items) return content.line_items;
  if (typeof content === "string") return content.split("\n").filter(l => l.trim());
  return [];
}

// ═══════════════════════════════════════════════════════════
// SLIDE 1 — COVER
// ═══════════════════════════════════════════════════════════
function buildCoverSlide(pptx, data) {
  const s = pptx.addSlide();
  addPegasusLogo(pptx, s);

  // System name — large bold
  s.addText((data.proposal_title || data.service_type || "SYSTEM PROPOSAL").toUpperCase(), {
    x: 0.3, y: 0.7, w: 12.7, h: 1.0,
    fontSize: 34, bold: true, color: C.black, fontFace: "Calibri",
  });

  s.addShape(pptx.shapes.LINE, { x: 0.3, y: 2.0, w: 12.7, h: 0, line: { color: C.ltgray, pt: 1.5 } });
  s.addText("PROPOSAL FOR :", { x: 0.3, y: 2.1, w: 5, h: 0.45, fontSize: 18, color: C.gray });
  s.addShape(pptx.shapes.LINE, { x: 0.3, y: 2.62, w: 12.7, h: 0, line: { color: C.ltgray, pt: 2 } });

  // Client name — underlined bold
  s.addText((data.client_name || "CLIENT").toUpperCase(), {
    x: 0.3, y: 2.72, w: 12.7, h: 0.72,
    fontSize: 28, bold: true, color: C.black, underline: true,
  });
  s.addShape(pptx.shapes.LINE, { x: 0.3, y: 3.52, w: 12.7, h: 0, line: { color: C.ltgray, pt: 1.5 } });

  // Company + Date
  s.addText("Tomas Tech Co., Ltd.", { x: 0.3, y: 3.62, w: 6, h: 0.38, fontSize: 14, color: C.gray });
  const now = new Date();
  const months = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  s.addText(`${months[now.getMonth()]} ${now.getFullYear()}`, {
    x: 10.5, y: 3.62, w: 2.5, h: 0.38, fontSize: 14, color: C.gray, align: "right",
  });

  // Image zones
  addImageZone(pptx, s, 0.3, 5.5, 2.5, 1.5, "TOMAS TECH LOGO", "logo");
  addImageZone(pptx, s, 9.8, 4.3, 3.2, 2.7, "ILLUSTRATION", "");

  // Footer
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 7.25, w: 13.3, h: 0.25,
    fill: { color: C.navy }, line: { color: C.navy },
  });
  s.addText("Copyright\u00A9 TOMAS TECH CORPORATION. All rights reserved.", {
    x: 0, y: 7.25, w: 13.3, h: 0.25,
    fontSize: 7, color: C.white, align: "center", valign: "middle",
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE — PROJECT PURPOSE / EXECUTIVE SUMMARY
// 3 objectives + 6-step chevron process flow
// ═══════════════════════════════════════════════════════════
function buildProjectPurpose(pptx, data, section) {
  const s = pptx.addSlide();
  addHeader(pptx, s, "Project Purpose", 2);

  const content = section.content;
  let mainText = "";
  let objectives = [];

  if (typeof content === "string") {
    mainText = content;
  } else if (content) {
    mainText = content.summary || content.description || content.overview || "";
    objectives = content.objectives || content.goals || content.key_points || [];
    if (!Array.isArray(objectives)) objectives = [];
  }

  // Main description
  if (mainText) {
    s.addText(mainText, {
      x: 0.3, y: 1.25, w: 12.7, h: 0.85,
      fontSize: 11.5, color: C.black, valign: "middle",
    });
  }

  // Objectives (up to 5)
  const objSlice = objectives.slice(0, 5);
  objSlice.forEach((o, i) => {
    const text = typeof o === "string" ? o : o.title || o.description || o.text || JSON.stringify(o);
    s.addText(`${i + 1}.  ${text}`, {
      x: 0.5, y: (mainText ? 2.22 : 1.3) + i * 0.55, w: 12.0, h: 0.48,
      fontSize: 12, color: C.black, valign: "middle",
    });
  });

  // 6-step chevron flow (from process_flow or auto-generated)
  let steps = [];
  if (content && content.process_flow) {
    steps = content.process_flow;
  } else if (content && content.steps) {
    steps = content.steps;
  } else if (content && content.flow) {
    steps = content.flow;
  }

  // If no steps provided, skip the chevron section
  if (steps.length > 0) {
    const flowColors = [C.headerBlue, C.teal, C.headerBlue, C.teal, C.orange, C.orange, C.green, C.navy];
    const flowY = 4.05, flowH = 0.95;
    const maxSteps = Math.min(steps.length, 6);
    const stepW = Math.min(1.92, (12.7 - 0.3) / maxSteps - 0.2);

    steps.slice(0, maxSteps).forEach((st, i) => {
      const text = typeof st === "string" ? st : st.name || st.title || st.text || `Step ${i + 1}`;
      const color = typeof st === "object" && st.color ? st.color : flowColors[i % flowColors.length];
      const x = 0.3 + i * (stepW + 0.2);

      s.addShape(pptx.shapes.CHEVRON, {
        x, y: flowY, w: stepW, h: flowH,
        fill: { color }, line: { color },
      });
      s.addText(text, {
        x, y: flowY, w: stepW, h: flowH,
        fontSize: 10, bold: true, color: C.white, align: "center", valign: "middle",
      });
    });

    // Icons below (if provided)
    const icons = content.icons || [];
    icons.slice(0, maxSteps).forEach((ic, i) => {
      const x = 0.3 + i * (stepW + 0.2);
      s.addText(ic, { x, y: 5.08, w: stepW, h: 0.45, fontSize: 20, align: "center" });
    });
  }
}

// ═══════════════════════════════════════════════════════════
// SLIDE — EXECUTIVE SUMMARY (text-focused version)
// ═══════════════════════════════════════════════════════════
function buildExecutiveSummary(pptx, data, section) {
  const s = pptx.addSlide();
  addHeader(pptx, s, "Executive Summary", 2);

  const content = typeof section.content === "string"
    ? section.content
    : section.content?.summary || section.content?.description || extractText(section.content);

  s.addText(content, {
    x: 0.5, y: 1.3, w: 12.3, h: 5.5,
    fontSize: 13, color: "333333", valign: "top", lineSpacing: 22,
    paraSpaceAfter: 8,
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE — CURRENT PAIN POINT / PROBLEM
// Numbered problem cards left + illustration zone right
// ═══════════════════════════════════════════════════════════
function buildPainPoint(pptx, data, section) {
  const s = pptx.addSlide();
  addHeader(pptx, s, "Current Pain Point", 3);

  const content = section.content;
  let problems = extractArray(content);
  if (problems.length === 0 && typeof content === "string") {
    problems = content.split("\n").filter(l => l.trim()).map(l => ({ title: l }));
  }

  const maxProblems = Math.min(problems.length, 5);
  const cardH = Math.min(1.28, 5.55 / maxProblems - 0.14);

  problems.slice(0, maxProblems).forEach((p, i) => {
    const y = 1.3 + i * (cardH + 0.14);
    const title = typeof p === "string" ? p : p.title || p.name || p.problem || JSON.stringify(p);
    const points = typeof p === "object" ? (p.points || p.details || p.description || []) : [];
    const pts = Array.isArray(points) ? points : [points];

    // Gray rounded card
    s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: 0.3, y, w: 9.2, h: cardH,
      fill: { color: C.ltgray }, line: { color: "BBBBBB", pt: 1 }, rectRadius: 0.07,
    });
    s.addText(`${i + 1}. ${title}`, {
      x: 0.5, y: y + 0.08, w: 8.8, h: 0.36,
      fontSize: 12, bold: true, color: C.black,
    });

    // Bullet points inside card
    pts.slice(0, 3).forEach((pt, pi) => {
      const ptText = typeof pt === "string" ? pt : pt.text || pt.description || "";
      if (ptText) {
        s.addText(`- ${ptText}`, {
          x: 0.7, y: y + 0.42 + pi * 0.28, w: 8.6, h: 0.26,
          fontSize: 10.5, color: "333333",
        });
      }
    });
  });

  // Illustration zone right
  addImageZone(pptx, s, 9.7, 1.3, 3.3, 5.55, "Problem Illustration", "");
}

// ═══════════════════════════════════════════════════════════
// SLIDE — BENEFITS OVERVIEW (3-column: Problem / Benefits / Use-Reduce)
// ═══════════════════════════════════════════════════════════
function buildBenefitsOverview(pptx, data, section) {
  const s = pptx.addSlide();
  const titleText = data.service_type
    ? `${data.service_type} | Benefits Target`
    : "Benefits Target";
  addHeader(pptx, s, titleText, 4);

  const content = section.content;
  let problems = [], benefits = [], reductions = [];

  if (content) {
    problems = content.problems || content.pain_points || content.current_issues || [];
    benefits = content.benefits || content.improvements || content.solutions || [];
    reductions = content.reductions || content.use_reduce || content.metrics || content.kpis || [];
    if (!Array.isArray(problems)) problems = [problems];
    if (!Array.isArray(benefits)) benefits = [benefits];
    if (!Array.isArray(reductions)) reductions = [reductions];
  }

  const colY = 1.25, colH = 5.85, colW = 3.8;

  // ── Column 1: Problem (gray) ──
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.18, y: colY, w: colW, h: colH, fill: { color: "F2F2F2" }, line: { color: C.ltgray, pt: 1 } });
  s.addShape(pptx.shapes.RECTANGLE, { x: 0.18, y: colY, w: colW, h: 0.4, fill: { color: "888888" }, line: { color: "888888" } });
  s.addText("Problem of Normal process", {
    x: 0.18, y: colY, w: colW, h: 0.4,
    fontSize: 10, bold: true, color: C.white, align: "center", valign: "middle",
  });

  // Problem icon row
  const probIcons = [
    { i: "\u{1F4CA}", l: "Manage Stock\nfailure" },
    { i: "\u{1F4E6}", l: "Wrong\nPicking" },
    { i: "\u{1F4CB}", l: "Wrong inventory\nand information" },
  ];
  probIcons.forEach((ic, pi) => {
    s.addText(ic.i, { x: 0.28 + pi * 1.24, y: colY + 0.5, w: 1.1, h: 0.5, fontSize: 22, align: "center" });
    s.addText(ic.l, { x: 0.28 + pi * 1.24, y: colY + 0.95, w: 1.1, h: 0.42, fontSize: 8, color: C.gray, align: "center" });
  });
  s.addShape(pptx.shapes.LINE, { x: 0.28, y: colY + 1.4, w: colW - 0.2, h: 0, line: { color: C.ltgray, pt: 0.5 } });

  problems.slice(0, 8).forEach((p, pi) => {
    const text = typeof p === "string" ? p : p.text || p.title || p.description || JSON.stringify(p);
    s.addText("\u2717  " + text, {
      x: 0.28, y: colY + 1.55 + pi * 0.5, w: colW - 0.2, h: 0.45,
      fontSize: 9.5, color: "444444", valign: "middle",
    });
  });

  // Arrow 1 (orange chevron)
  s.addShape(pptx.shapes.CHEVRON, { x: 4.1, y: 3.6, w: 0.6, h: 0.65, fill: { color: C.orange }, line: { color: C.orange } });

  // ── Column 2: Benefits (orange) ──
  s.addShape(pptx.shapes.RECTANGLE, { x: 4.78, y: colY, w: colW, h: colH, fill: { color: "FFF8F0" }, line: { color: "FFCCAA", pt: 1 } });
  s.addShape(pptx.shapes.RECTANGLE, { x: 4.78, y: colY, w: colW, h: 0.4, fill: { color: C.orange }, line: { color: C.orange } });
  s.addText("Benefits", {
    x: 4.78, y: colY, w: colW, h: 0.4,
    fontSize: 10, bold: true, color: C.white, align: "center", valign: "middle",
  });

  benefits.slice(0, 8).forEach((b, bi) => {
    const text = typeof b === "string" ? b : b.text || b.title || b.description || JSON.stringify(b);
    s.addText("\u2714  " + text, {
      x: 4.9, y: colY + 0.5 + bi * 0.55, w: colW - 0.2, h: 0.48,
      fontSize: 9.5, color: "444444", valign: "middle",
    });
  });

  // Arrow 2 (blue chevron)
  s.addShape(pptx.shapes.CHEVRON, { x: 8.7, y: 3.6, w: 0.6, h: 0.65, fill: { color: C.headerBlue }, line: { color: C.headerBlue } });

  // ── Column 3: Use/Reduce (blue) ──
  s.addShape(pptx.shapes.RECTANGLE, { x: 9.35, y: colY, w: colW, h: colH, fill: { color: "EEF4FF" }, line: { color: "C0D0E8", pt: 1 } });
  s.addShape(pptx.shapes.RECTANGLE, { x: 9.35, y: colY, w: colW, h: 0.4, fill: { color: C.headerBlue }, line: { color: C.headerBlue } });
  s.addText("Use / Reduce", {
    x: 9.35, y: colY, w: colW, h: 0.4,
    fontSize: 10, bold: true, color: C.white, align: "center", valign: "middle",
  });

  reductions.slice(0, 3).forEach((r, ri) => {
    const ry = colY + 0.6 + ri * 1.65;
    const value = typeof r === "string" ? r : r.value || r.metric || r.number || "";
    const label = typeof r === "object" ? r.label || r.title || r.name || "" : "";
    const sub = typeof r === "object" ? r.subtitle || r.description || r.detail || "" : "";

    s.addText(String(value), {
      x: 9.35, y: ry, w: colW, h: 0.7,
      fontSize: 28, bold: true, color: C.orange, align: "center", fontFace: "Calibri",
    });
    if (label) {
      s.addText(label, {
        x: 9.35, y: ry + 0.68, w: colW, h: 0.42,
        fontSize: 10, color: C.black, align: "center",
      });
    }
    if (sub) {
      s.addText(sub, {
        x: 9.35, y: ry + 1.08, w: colW, h: 0.32,
        fontSize: 9, color: C.gray, align: "center", italic: true,
      });
    }
    if (ri < 2) {
      s.addShape(pptx.shapes.LINE, { x: 9.45, y: ry + 1.45, w: colW - 0.2, h: 0, line: { color: "C0D0E8", pt: 0.5 } });
    }
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE — SCOPE OF WORK (numbered items with blue accent)
// ═══════════════════════════════════════════════════════════
function buildScopeOfWork(pptx, data, section) {
  // Divider
  const d = pptx.addSlide();
  addSectionDivider(pptx, d, "SCOPE OF WORK", null);

  const s = pptx.addSlide();
  addHeader(pptx, s, "Scope of Work - Details", null);

  const items = extractArray(section.content);
  let yPos = 1.3;
  let currentSlide = s;

  items.forEach((item, i) => {
    if (yPos > 6.2) {
      currentSlide = pptx.addSlide();
      addHeader(pptx, currentSlide, "Scope of Work - Details (cont.)", null);
      yPos = 1.3;
    }
    const text = typeof item === "string" ? item : item.title || item.name || JSON.stringify(item);
    const desc = typeof item === "object" ? item.description || item.detail || "" : "";

    // Blue numbered badge
    currentSlide.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: 0.3, y: yPos, w: 0.4, h: 0.4,
      fill: { color: C.headerBlue }, rectRadius: 0.05,
    });
    currentSlide.addText(String(i + 1), {
      x: 0.3, y: yPos, w: 0.4, h: 0.4,
      fontSize: 14, bold: true, color: C.white, align: "center", valign: "middle",
    });
    currentSlide.addText(text, {
      x: 0.85, y: yPos, w: 11.5, h: 0.4,
      fontSize: 13, bold: true, color: C.black, valign: "middle",
    });
    yPos += 0.45;

    if (desc) {
      currentSlide.addText(desc, {
        x: 0.85, y: yPos, w: 11.5, h: 0.35,
        fontSize: 11, color: C.gray, valign: "top",
      });
      yPos += 0.4;
    }
    yPos += 0.15;
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE — SYSTEM OUTLINE (Server+DB zone + Operation zone)
// ═══════════════════════════════════════════════════════════
function buildSystemOutline(pptx, data, section) {
  const s = pptx.addSlide();
  addHeader(pptx, s, "Suggestion | System Outline", null);

  const content = section.content;

  // Left zone: Server and Database
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.18, y: 1.25, w: 3.6, h: 5.85,
    fill: { color: "F8FAFF" }, line: { color: "AAAAAA", pt: 1.5, dashType: "dash" }, rectRadius: 0.1,
  });
  s.addText("Server and Database", {
    x: 0.18, y: 1.25, w: 3.6, h: 0.4,
    fontSize: 11, bold: true, color: C.gray, align: "center", valign: "middle",
  });

  // ERP block
  const erpName = (content && content.erp_name) || "ERP Client Site";
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.5, y: 1.78, w: 2.9, h: 0.72,
    fill: { color: C.teal }, line: { color: C.teal }, rectRadius: 0.07,
  });
  s.addText(erpName, {
    x: 0.5, y: 1.78, w: 2.9, h: 0.72,
    fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle",
  });

  // Arrow
  s.addText("\u2195", { x: 1.65, y: 2.56, w: 0.6, h: 0.38, fontSize: 18, color: C.gray, align: "center" });

  // WMS/System Server block
  const serverName = (content && content.server_name) || "System\nServer";
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.5, y: 2.98, w: 2.9, h: 0.88,
    fill: { color: C.headerBlue }, line: { color: C.headerBlue }, rectRadius: 0.07,
  });
  s.addText(serverName, {
    x: 0.5, y: 2.98, w: 2.9, h: 0.88,
    fontSize: 12, bold: true, color: C.white, align: "center", valign: "middle",
  });

  // PEGASUS badge on server
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 0.55, y: 3.92, w: 1.3, h: 0.3,
    fill: { color: C.orange }, line: { color: C.orange }, rectRadius: 0.04,
  });
  s.addText("PEGASUS", {
    x: 0.55, y: 3.92, w: 1.3, h: 0.3,
    fontSize: 8, bold: true, color: C.white, align: "center", valign: "middle",
  });

  // Hardware list below server
  const hwList = (content && content.hardware) || [
    "Label Printer", "Warehouse Control", "Barcode Scanner",
  ];
  hwList.slice(0, 4).forEach((hw, i) => {
    const hwText = typeof hw === "string" ? hw : hw.name || hw.title || "";
    s.addText("\u{1F5A8}  " + hwText, {
      x: 0.3, y: 4.38 + i * 0.37, w: 3.3, h: 0.35,
      fontSize: 10, color: C.black,
    });
  });

  // Right zone: Operation System
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 4.05, y: 1.25, w: 5.95, h: 5.85,
    fill: { color: "F8FFF8" }, line: { color: "AAAAAA", pt: 1.5, dashType: "dash" }, rectRadius: 0.1,
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 4.05, y: 1.25, w: 5.95, h: 0.4,
    fill: { color: C.teal }, line: { color: C.teal },
  });
  s.addText("Operation System", {
    x: 4.05, y: 1.25, w: 5.95, h: 0.4,
    fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle",
  });

  // Department icons
  const defaultDepts = [
    { n: "Materials &\nFG Store", col: C.headerBlue },
    { n: "Production", col: C.teal },
    { n: "Outbound\nDelivery", col: C.orange },
    { n: "Office", col: C.gray },
  ];
  const depts = (content && content.departments) || defaultDepts;
  const deptColors = [C.headerBlue, C.teal, C.orange, C.gray, C.green, C.navy];

  depts.slice(0, 4).forEach((d, i) => {
    const dx = 4.3 + i * 1.45;
    const name = typeof d === "string" ? d : d.name || d.n || `Dept ${i + 1}`;
    const col = typeof d === "object" ? d.col || deptColors[i] : deptColors[i];

    s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: dx, y: 1.78, w: 1.28, h: 0.9,
      fill: { color: "E8EEF8" }, line: { color: "C0CADD", pt: 1 }, rectRadius: 0.06,
    });
    s.addText("\u{1F5A5}", { x: dx, y: 1.82, w: 1.28, h: 0.45, fontSize: 18, align: "center" });
    s.addText(name, { x: dx, y: 2.25, w: 1.28, h: 0.38, fontSize: 8, color: C.black, align: "center" });
    s.addText("\u2191", { x: dx + 0.44, y: 2.72, w: 0.4, h: 0.35, fontSize: 14, color: col, align: "center" });
  });

  // Connection bar
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 4.25, y: 3.18, w: 5.5, h: 0.6,
    fill: { color: "EEF4FF" }, line: { color: C.headerBlue, pt: 1 }, rectRadius: 0.06,
  });
  const connText = (content && content.connection_text) || "PEGASUS \u2014 Real-time Inventory | ERP Integration";
  s.addText(connText, {
    x: 4.25, y: 3.18, w: 5.5, h: 0.6,
    fontSize: 10, bold: true, color: C.headerBlue, align: "center", valign: "middle",
  });

  // Hardware requirements box
  const hwReqs = (content && content.hardware_requirements) || [
    "Handy Terminal (Android)", "Client PC \u2014 Web browser", "Label Printer",
    "WiFi Access Point", "Cloud or On-premise Server",
  ];
  s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
    x: 4.25, y: 3.9, w: 5.5, h: 3.0,
    fill: { color: C.white }, line: { color: C.ltgray, pt: 1 }, rectRadius: 0.07,
  });
  hwReqs.slice(0, 5).forEach((h, i) => {
    const hText = typeof h === "string" ? h : h.name || h.title || "";
    s.addText("\u{1F4F1}  " + hText, {
      x: 4.45, y: 4.05 + i * 0.55, w: 5.1, h: 0.48,
      fontSize: 10.5, color: C.black, valign: "middle",
    });
  });

  // System image zone far right
  addImageZone(pptx, s, 10.15, 1.25, 2.95, 5.85, "SYSTEM IMAGE", "");
}

// ═══════════════════════════════════════════════════════════
// SLIDE — SYSTEM FLOW DIAGRAM (chevron steps + detail boxes)
// ═══════════════════════════════════════════════════════════
function buildSystemFlowDiagram(pptx, data, section) {
  const s = pptx.addSlide();
  const flowTitle = section.flow_title || section.title || "System Flow Diagram";
  addHeader(pptx, s, "Suggestion | " + flowTitle, null);

  const content = section.content;
  let steps = [];

  if (content && content.steps) steps = content.steps;
  else if (content && content.flow) steps = content.flow;
  else if (content && content.process) steps = content.process;
  else if (Array.isArray(content)) steps = content;
  else if (typeof content === "string") {
    steps = content.split("\n").filter(l => l.trim()).map(l => ({ title: l }));
  }

  // Title bar
  const barColor = (content && content.bar_color) || C.headerBlue;
  const barTitle = (content && content.bar_title) || flowTitle;
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.3, y: 1.3, w: 12.7, h: 0.42,
    fill: { color: barColor }, line: { color: barColor },
  });
  s.addText(barTitle, {
    x: 0.3, y: 1.3, w: 12.7, h: 0.42,
    fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle",
  });

  const flowColors = [C.headerBlue, C.teal, C.headerBlue, C.teal, C.orange, C.orange];
  const maxSteps = Math.min(steps.length, 5);

  steps.slice(0, maxSteps).forEach((st, i) => {
    const x = 0.5 + i * 2.52;
    const title = typeof st === "string" ? st : st.title || st.name || st.t || `Step ${i + 1}`;
    const detail = typeof st === "object" ? st.detail || st.description || st.s || "" : "";
    const col = typeof st === "object" && st.color ? st.color : flowColors[i % flowColors.length];

    // Chevron
    s.addShape(pptx.shapes.CHEVRON, {
      x, y: 1.88, w: 2.35, h: 1.05,
      fill: { color: col }, line: { color: col },
    });
    s.addText(title, {
      x, y: 1.88, w: 2.35, h: 0.58,
      fontSize: 10, bold: true, color: C.white, align: "center", valign: "middle",
    });
    if (detail) {
      s.addText(detail, {
        x, y: 2.42, w: 2.35, h: 0.46,
        fontSize: 8, color: "DDDDDD", align: "center", italic: true, valign: "middle",
      });
    }

    // Detail box below chevron
    s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: x + 0.08, y: 3.08, w: 2.18, h: 2.35,
      fill: { color: C.white }, line: { color: col, pt: 1 }, rectRadius: 0.07,
    });
    s.addShape(pptx.shapes.RECTANGLE, {
      x: x + 0.08, y: 3.08, w: 2.18, h: 0.3,
      fill: { color: col }, line: { color: col },
    });
    s.addText("Detail", {
      x: x + 0.08, y: 3.08, w: 2.18, h: 0.3,
      fontSize: 8.5, bold: true, color: C.white, align: "center", valign: "middle",
    });
    s.addText(detail || title, {
      x: x + 0.15, y: 3.42, w: 2.0, h: 1.88,
      fontSize: 9.5, color: C.black, valign: "top",
    });
  });

  // Image zones at bottom
  addImageZone(pptx, s, 0.3, 5.55, 4.0, 1.6, "HANDY SCREEN", "");
  addImageZone(pptx, s, 4.5, 5.55, 4.0, 1.6, "PC SCREEN", "");
  addImageZone(pptx, s, 8.7, 5.55, 4.3, 1.6, "SAMPLE OUTPUT", "");
}

// ═══════════════════════════════════════════════════════════
// SLIDE — OPERATION FLOW (numbered circle steps + descriptions)
// ═══════════════════════════════════════════════════════════
function buildOperationFlow(pptx, data, section) {
  const s = pptx.addSlide();
  const flowTitle = section.flow_title || "Explaining Operation Flow";
  addHeader(pptx, s, "Suggestion | " + flowTitle, null);

  s.addText("Summary operation flow", {
    x: 0.3, y: 1.28, w: 9, h: 0.38,
    fontSize: 13, bold: true, color: C.headerBlue,
  });
  s.addShape(pptx.shapes.LINE, { x: 0.3, y: 1.68, w: 9, h: 0, line: { color: C.headerBlue, pt: 1.5 } });

  const content = section.content;
  let flowItems = extractArray(content);

  const maxItems = Math.min(flowItems.length, 5);
  const itemH = Math.min(1.28, 5.2 / maxItems - 0.1);

  flowItems.slice(0, maxItems).forEach((f, i) => {
    const y = 1.88 + i * (itemH + 0.1);
    const title = typeof f === "string" ? f : f.title || f.name || f.t || `Step ${i + 1}`;
    const desc = typeof f === "object" ? f.description || f.detail || f.d || "" : "";

    // Numbered circle
    s.addShape(pptx.shapes.OVAL, {
      x: 0.3, y: y + 0.08, w: 0.42, h: 0.42,
      fill: { color: C.headerBlue }, line: { color: C.headerBlue },
    });
    s.addText(String(i + 1), {
      x: 0.3, y: y + 0.08, w: 0.42, h: 0.42,
      fontSize: 13, bold: true, color: C.white, align: "center", valign: "middle",
    });

    // Title + description
    s.addText(title, {
      x: 0.85, y, w: 8.5, h: 0.38,
      fontSize: 12, bold: true, color: C.black,
    });
    if (desc) {
      s.addText(desc, {
        x: 0.85, y: y + 0.38, w: 8.5, h: itemH - 0.38,
        fontSize: 10, color: C.gray, valign: "top",
      });
    }
  });

  // Screen flow zone right
  addImageZone(pptx, s, 9.8, 1.28, 3.2, 5.85, "SCREEN FLOW", "");
}

// ═══════════════════════════════════════════════════════════
// SLIDE — TIMELINE (chevron phases)
// ═══════════════════════════════════════════════════════════
function buildTimeline(pptx, data, section) {
  const s = pptx.addSlide();
  addHeader(pptx, s, "Timeline & Milestones", null);

  const phases = extractArray(section.content);
  const colors = [C.headerBlue, C.teal, C.orange, C.headerBlue, C.teal, C.orange, C.green, C.navy];
  const maxPerRow = 4;

  phases.forEach((phase, i) => {
    const row = Math.floor(i / maxPerRow);
    const col = i % maxPerRow;
    const x = 0.3 + col * 3.2;
    const y = 1.4 + row * 2.8;
    const color = colors[i % colors.length];
    const text = typeof phase === "string" ? phase : phase.name || phase.title || phase.phase || `Phase ${i + 1}`;
    const duration = typeof phase === "object" ? phase.duration || phase.period || "" : "";

    s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x, y, w: 2.9, h: 0.55,
      fill: { color }, rectRadius: 0.06,
    });
    s.addText(text, {
      x, y, w: 2.9, h: 0.55,
      fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle",
    });

    if (duration) {
      s.addText(duration, {
        x, y: y + 0.6, w: 2.9, h: 0.3,
        fontSize: 10, color: C.gray, align: "center",
      });
    }

    if (col < maxPerRow - 1 && i < phases.length - 1) {
      s.addText("\u25B6", {
        x: x + 2.9, y, w: 0.3, h: 0.55,
        fontSize: 14, color, align: "center", valign: "middle",
      });
    }
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE — TECH STACK (2-column grid with icons)
// ═══════════════════════════════════════════════════════════
function buildTechStack(pptx, data, section) {
  const s = pptx.addSlide();
  addHeader(pptx, s, "Technology Stack", null);

  const items = extractArray(section.content);
  if (items.length === 0 && typeof section.content === "object" && !Array.isArray(section.content)) {
    // Handle key-value object
    Object.entries(section.content).forEach(([k, v], i) => {
      const col = i % 2;
      const row = Math.floor(i / 2);
      const x = 0.3 + col * 6.5;
      const y = 1.4 + row * 0.85;

      s.addShape(pptx.shapes.OVAL, { x, y: y + 0.05, w: 0.4, h: 0.4, fill: { color: C.teal } });
      s.addText("\u2699", { x, y: y + 0.05, w: 0.4, h: 0.4, fontSize: 14, color: C.white, align: "center", valign: "middle" });
      s.addText(k, { x: x + 0.5, y, w: 5.5, h: 0.35, fontSize: 12, bold: true, color: C.black, valign: "middle" });
      s.addText(String(v), { x: x + 0.5, y: y + 0.35, w: 5.5, h: 0.3, fontSize: 10, color: C.gray });
    });
    return;
  }

  items.forEach((item, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.3 + col * 6.5;
    const y = 1.4 + row * 0.85;
    const text = typeof item === "string" ? item : item.name || item.technology || JSON.stringify(item);
    const desc = typeof item === "object" ? item.description || item.purpose || "" : "";

    s.addShape(pptx.shapes.OVAL, { x, y: y + 0.05, w: 0.4, h: 0.4, fill: { color: C.teal } });
    s.addText("\u2699", { x, y: y + 0.05, w: 0.4, h: 0.4, fontSize: 14, color: C.white, align: "center", valign: "middle" });
    s.addText(text, { x: x + 0.5, y, w: 5.5, h: 0.35, fontSize: 12, bold: true, color: C.black, valign: "middle" });
    if (desc) {
      s.addText(desc, { x: x + 0.5, y: y + 0.35, w: 5.5, h: 0.3, fontSize: 10, color: C.gray });
    }
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE — TEAM STRUCTURE (org chart cards)
// ═══════════════════════════════════════════════════════════
function buildTeamStructure(pptx, data, section) {
  const s = pptx.addSlide();
  addHeader(pptx, s, "Team Structure", null);

  const roles = extractArray(section.content);

  roles.slice(0, 9).forEach((role, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.5 + col * 4.2;
    const y = 1.5 + row * 1.8;
    const text = typeof role === "string" ? role : role.role || role.title || role.name || JSON.stringify(role);
    const desc = typeof role === "object" ? role.responsibilities || role.description || "" : "";

    // Card
    s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x, y, w: 3.8, h: 1.4,
      fill: { color: "F8F9FA" }, line: { color: C.headerBlue, pt: 1.5 }, rectRadius: 0.08,
    });
    // Title bar
    s.addShape(pptx.shapes.RECTANGLE, {
      x: x + 0.05, y: y + 0.05, w: 3.7, h: 0.4,
      fill: { color: C.headerBlue },
    });
    s.addText(text, {
      x: x + 0.05, y: y + 0.05, w: 3.7, h: 0.4,
      fontSize: 11, bold: true, color: C.white, align: "center", valign: "middle",
    });

    if (desc) {
      const descText = Array.isArray(desc) ? desc.join(", ") : desc;
      s.addText(descText, {
        x: x + 0.15, y: y + 0.5, w: 3.5, h: 0.8,
        fontSize: 9, color: C.gray, valign: "top",
      });
    }
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE — PRICING TABLE
// ═══════════════════════════════════════════════════════════
function buildPricing(pptx, data, section) {
  const s = pptx.addSlide();
  addHeader(pptx, s, "Pricing Estimate", null);

  const content = section.content;

  if (typeof content === "string") {
    s.addText(content, {
      x: 0.5, y: 1.4, w: 12.3, h: 5.5,
      fontSize: 13, color: "333333", valign: "top", lineSpacing: 22,
    });
    return;
  }

  let rows = [];
  if (Array.isArray(content)) rows = content;
  else if (content?.items || content?.line_items) rows = content.items || content.line_items;
  else rows = [content];

  const tableData = [
    [
      { text: "#", options: { fill: C.headerBlue, color: C.white, bold: true, fontSize: 10, align: "center" } },
      { text: "Item", options: { fill: C.headerBlue, color: C.white, bold: true, fontSize: 10 } },
      { text: "Description", options: { fill: C.headerBlue, color: C.white, bold: true, fontSize: 10 } },
      { text: "Estimate", options: { fill: C.headerBlue, color: C.white, bold: true, fontSize: 10, align: "right" } },
    ],
  ];

  rows.forEach((row, i) => {
    const item = typeof row === "string" ? row : row.item || row.name || row.service || "";
    const desc = typeof row === "object" ? row.description || row.detail || "" : "";
    const price = typeof row === "object" ? row.price || row.estimate || row.cost || "" : "";
    const bgColor = i % 2 === 0 ? "FFFFFF" : "F5F7FA";

    tableData.push([
      { text: String(i + 1), options: { fill: bgColor, fontSize: 10, align: "center" } },
      { text: String(item), options: { fill: bgColor, fontSize: 10 } },
      { text: String(desc), options: { fill: bgColor, fontSize: 9, color: C.gray } },
      { text: String(price), options: { fill: bgColor, fontSize: 10, align: "right", bold: true } },
    ]);
  });

  if (content?.total) {
    tableData.push([
      { text: "", options: { fill: C.navy } },
      { text: "TOTAL", options: { fill: C.navy, color: C.white, bold: true, fontSize: 11 } },
      { text: "", options: { fill: C.navy } },
      { text: String(content.total), options: { fill: C.navy, color: C.orange, bold: true, fontSize: 12, align: "right" } },
    ]);
  }

  s.addTable(tableData, {
    x: 0.3, y: 1.4, w: 12.7,
    border: { pt: 0.5, color: "CCCCCC" },
    colW: [0.6, 3.5, 5.5, 3.1],
    rowH: 0.45,
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE — TERMS & CONDITIONS
// ═══════════════════════════════════════════════════════════
function buildTerms(pptx, data, section) {
  const s = pptx.addSlide();
  addHeader(pptx, s, "Terms & Conditions", null);

  const content = section.content;
  let text = "";

  if (typeof content === "string") {
    text = content;
  } else if (Array.isArray(content)) {
    text = content.map((item, i) => {
      if (typeof item === "string") return `${i + 1}. ${item}`;
      return `${i + 1}. ${item.title || item.term || ""}: ${item.description || item.detail || ""}`;
    }).join("\n\n");
  } else {
    text = JSON.stringify(content, null, 2);
  }

  s.addText(text, {
    x: 0.5, y: 1.4, w: 12.3, h: 5.5,
    fontSize: 12, color: "333333", valign: "top", lineSpacing: 20,
    paraSpaceAfter: 6,
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE — FUNCTION LIST TABLE
// ═══════════════════════════════════════════════════════════
function buildFunctionList(pptx, data, section) {
  const s = pptx.addSlide();
  addHeader(pptx, s, "Function List", null);

  const content = section.content;
  let functions = extractArray(content);

  const tableData = [
    [
      { text: "#", options: { fill: C.headerBlue, color: C.white, bold: true, fontSize: 10, align: "center" } },
      { text: "Category", options: { fill: C.headerBlue, color: C.white, bold: true, fontSize: 10 } },
      { text: "Function", options: { fill: C.headerBlue, color: C.white, bold: true, fontSize: 10 } },
      { text: "Description", options: { fill: C.headerBlue, color: C.white, bold: true, fontSize: 10 } },
    ],
  ];

  functions.slice(0, 15).forEach((fn, i) => {
    const cat = typeof fn === "object" ? fn.category || fn.module || "" : "";
    const name = typeof fn === "string" ? fn : fn.name || fn.function || fn.title || "";
    const desc = typeof fn === "object" ? fn.description || fn.detail || "" : "";
    const bgColor = i % 2 === 0 ? "FFFFFF" : "F5F7FA";

    tableData.push([
      { text: String(i + 1), options: { fill: bgColor, fontSize: 9, align: "center" } },
      { text: String(cat), options: { fill: bgColor, fontSize: 9 } },
      { text: String(name), options: { fill: bgColor, fontSize: 9, bold: true } },
      { text: String(desc), options: { fill: bgColor, fontSize: 9, color: C.gray } },
    ]);
  });

  s.addTable(tableData, {
    x: 0.3, y: 1.4, w: 12.7,
    border: { pt: 0.5, color: "CCCCCC" },
    colW: [0.5, 2.5, 3.5, 6.2],
    rowH: 0.38,
  });
}

// ═══════════════════════════════════════════════════════════
// SLIDE — MAINTENANCE / SLA TABLE
// ═══════════════════════════════════════════════════════════
function buildMaintenance(pptx, data, section) {
  const s = pptx.addSlide();
  addHeader(pptx, s, "System Maintenance & Support", null);

  const content = section.content;
  let text = extractText(content, "");

  if (typeof content === "string" || !content) {
    s.addText(text || "Maintenance details to be confirmed.", {
      x: 0.5, y: 1.4, w: 12.3, h: 5.5,
      fontSize: 12, color: "333333", valign: "top", lineSpacing: 20,
    });
    return;
  }

  // Try to build table from structured data
  let items = extractArray(content);
  if (items.length > 0) {
    const tableData = [
      [
        { text: "Item", options: { fill: C.headerBlue, color: C.white, bold: true, fontSize: 10 } },
        { text: "Year 1", options: { fill: C.headerBlue, color: C.white, bold: true, fontSize: 10, align: "center" } },
        { text: "Year 2+", options: { fill: C.headerBlue, color: C.white, bold: true, fontSize: 10, align: "center" } },
        { text: "Notes", options: { fill: C.headerBlue, color: C.white, bold: true, fontSize: 10 } },
      ],
    ];

    items.forEach((item, i) => {
      const name = typeof item === "string" ? item : item.name || item.item || item.service || "";
      const y1 = typeof item === "object" ? item.year1 || item.warranty || "Included" : "";
      const y2 = typeof item === "object" ? item.year2 || item.renewal || "" : "";
      const notes = typeof item === "object" ? item.notes || item.description || "" : "";
      const bgColor = i % 2 === 0 ? "FFFFFF" : "F5F7FA";

      tableData.push([
        { text: String(name), options: { fill: bgColor, fontSize: 10 } },
        { text: String(y1), options: { fill: bgColor, fontSize: 10, align: "center" } },
        { text: String(y2), options: { fill: bgColor, fontSize: 10, align: "center" } },
        { text: String(notes), options: { fill: bgColor, fontSize: 9, color: C.gray } },
      ]);
    });

    s.addTable(tableData, {
      x: 0.3, y: 1.4, w: 12.7,
      border: { pt: 0.5, color: "CCCCCC" },
      colW: [4.0, 2.5, 2.5, 3.7],
      rowH: 0.45,
    });
  } else {
    // Fallback: render summary text
    s.addText(content.summary || JSON.stringify(content), {
      x: 0.5, y: 1.4, w: 12.3, h: 5.5,
      fontSize: 12, color: "333333", valign: "top", lineSpacing: 20,
    });
  }
}

// ═══════════════════════════════════════════════════════════
// SLIDE — THANK YOU
// ═══════════════════════════════════════════════════════════
function buildThankYouSlide(pptx, data) {
  const s = pptx.addSlide();

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 0, w: 13.3, h: 7.5,
    fill: { color: C.navy },
  });

  s.addText("Thank You", {
    x: 0, y: 2.0, w: 13.3, h: 1.2,
    fontSize: 48, bold: true, color: C.white, align: "center", valign: "middle",
  });

  s.addText("TOMAS TECH CO., LTD.", {
    x: 0, y: 3.5, w: 13.3, h: 0.6,
    fontSize: 20, color: C.orange, align: "center", valign: "middle",
  });

  s.addText(
    "7/1(3C) Udomsuk 46 Alley, Khwaeng Bang Na Nuea, Khet Bang Na, Bangkok 10260\n" +
    "Tel: +66-98-271-9741  |  Email: info@tomastc.com  |  www.tomastc.com",
    {
      x: 1.5, y: 4.5, w: 10.3, h: 1.0,
      fontSize: 11, color: C.white, align: "center", lineSpacing: 18,
    }
  );

  addFooter(pptx, s, null);
}

// ═══════════════════════════════════════════════════════════
// SLIDE — GENERIC FALLBACK
// ═══════════════════════════════════════════════════════════
function buildGenericSlide(pptx, data, section, slideNum) {
  const s = pptx.addSlide();
  const title = section.section_title || section.title || `Section ${slideNum}`;
  addHeader(pptx, s, title, slideNum);

  const text = extractText(section.content, "");
  s.addText(text.substring(0, 2000), {
    x: 0.5, y: 1.3, w: 12.3, h: 5.7,
    fontSize: 12, color: "333333", valign: "top", lineSpacing: 20,
    paraSpaceAfter: 6,
  });
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

  // Parse if string
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

  // 2. Process sections — smart keyword matching
  const sections = data.sections || [];
  let slideNum = 2;

  // Enhanced section builder mapping with priority order
  const sectionBuilders = [
    // Project purpose / overview (with chevron flow)
    { keywords: ["project purpose", "purpose", "overview", "proposal overview"], builder: buildProjectPurpose },
    // Executive summary
    { keywords: ["executive", "summary"], builder: buildExecutiveSummary },
    // Pain point / problems
    { keywords: ["pain point", "pain", "problem", "challenge", "current issue", "current situation"], builder: buildPainPoint },
    // Benefits (3-column)
    { keywords: ["benefit", "target", "improvement", "advantage"], builder: buildBenefitsOverview },
    // System outline (architecture diagram)
    { keywords: ["system outline", "outline", "architecture", "infrastructure", "system design"], builder: buildSystemOutline },
    // System flow diagram
    { keywords: ["system flow", "flow diagram", "process flow", "inbound", "outbound"], builder: buildSystemFlowDiagram },
    // Operation flow
    { keywords: ["operation flow", "operation", "explaining"], builder: buildOperationFlow },
    // Function list
    { keywords: ["function list", "function", "feature list"], builder: buildFunctionList },
    // Scope of work
    { keywords: ["scope"], builder: buildScopeOfWork },
    // Timeline
    { keywords: ["timeline", "milestone", "schedule", "go live", "phase"], builder: buildTimeline },
    // Tech stack
    { keywords: ["technology", "tech_stack", "stack", "tech"], builder: buildTechStack },
    // Team
    { keywords: ["team", "structure", "resource", "organization"], builder: buildTeamStructure },
    // Pricing
    { keywords: ["pricing", "price", "cost", "budget", "estimate", "quotation"], builder: buildPricing },
    // Maintenance / SLA
    { keywords: ["maintenance", "support", "sla", "service level", "warranty"], builder: buildMaintenance },
    // Terms
    { keywords: ["terms", "condition", "payment"], builder: buildTerms },
  ];

  sections.forEach((section) => {
    const title = (section.section_title || section.title || "").toLowerCase();
    let built = false;

    // Check for section divider markers
    if (section.type === "divider" || title.startsWith("[divider]")) {
      const d = pptx.addSlide();
      const divTitle = title.replace("[divider]", "").trim() || section.divider_title || "Section";
      addSectionDivider(pptx, d, divTitle.toUpperCase(), slideNum);
      slideNum++;
      return;
    }

    // Match against keyword patterns
    for (const { keywords, builder } of sectionBuilders) {
      if (keywords.some((kw) => title.includes(kw))) {
        builder(pptx, data, section);
        built = true;
        break;
      }
    }

    if (!built) {
      buildGenericSlide(pptx, data, section, slideNum);
    }
    slideNum++;
  });

  // Last: Thank you slide
  buildThankYouSlide(pptx, data);

  // Generate buffer
  const buffer = await pptx.write({ outputType: "nodebuffer" });
  return buffer;
}

// ═══════════════════════════════════════════════════════════
// API Handler
// ═══════════════════════════════════════════════════════════
module.exports = async function handler(req, res) {
  // CORS
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  try {
    const input = req.body;
    const proposalContent = input.proposal_content || input.output || input.text || input;

    if (!proposalContent) {
      return res.status(400).json({ error: "proposal_content is required" });
    }

    // Parse proposal content
    let proposalData = proposalContent;
    if (typeof proposalData === "string") {
      try {
        const cleaned = proposalData.replace(/```json\s*/g, "").replace(/```\s*/g, "").trim();
        const match = cleaned.match(/\{[\s\S]*\}/);
        if (match) proposalData = JSON.parse(match[0]);
      } catch (e) {
        // keep as string
      }
    }

    const requestId = input.request_id;
    const clientName = input.client_name || (typeof proposalData === "object" ? proposalData.client_name : null);

    // Generate PPTX
    const buffer = await generatePPTX(proposalData);

    // Generate filename
    const safeName = (clientName || "proposal").replace(/[^a-zA-Z0-9\u0E00-\u0E7F]/g, "_").substring(0, 30);
    const timestamp = new Date().toISOString().slice(0, 10).replace(/-/g, "");
    const filename = `TOMAS_TECH_Proposal_${safeName}_${timestamp}.pptx`;

    // Upload to Supabase Storage
    if (SUPABASE_KEY) {
      try {
        const supabase = createClient(SUPABASE_URL, SUPABASE_KEY);

        const { data: uploadData, error: uploadError } = await supabase.storage
          .from("proposals")
          .upload(`pptx/${filename}`, buffer, {
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
          await supabase
            .from("proposal_requests")
            .update({ pptx_url: urlData.publicUrl, status: "completed" })
            .eq("id", requestId);
        }

        return res.status(200).json({
          success: true,
          filename,
          pptx_url: urlData.publicUrl,
          request_id: requestId || null,
        });
      } catch (storageError) {
        console.error("Storage error:", storageError);
      }
    }

    // Fallback: return binary file
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    return res.status(200).send(buffer);

  } catch (error) {
    console.error("PPTX generation error:", error);
    return res.status(500).json({ error: "Failed to generate PPTX", details: error.message });
  }
};
