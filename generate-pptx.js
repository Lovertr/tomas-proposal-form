const pptxgen = require("pptxgenjs");
const { createClient } = require("@supabase/supabase-js");

// ── Supabase Config ──
const SUPABASE_URL = "https://hsmawbvxhlkssdowlbzx.supabase.co";
const SUPABASE_KEY = process.env.SUPABASE_SERVICE_KEY; // set in Vercel env vars

// ── Brand Colors (from real TOMAS TECH proposals) ──
const C = {
  headerBlue: "1F5BA8",
  navy: "1C3F7F",
  orange: "F7941D",
  teal: "2BB5B8",
  white: "FFFFFF",
  black: "000000",
  gray: "595959",
  ltgray: "D9D9D9",
  green: "70AD47",
};

// ── HELPERS ──
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

// ── SLIDE BUILDERS ──

function buildCoverSlide(pptx, data) {
  const s = pptx.addSlide();
  addPegasusLogo(pptx, s);

  // System name
  s.addText((data.proposal_title || data.service_type || "SYSTEM PROPOSAL").toUpperCase(), {
    x: 0.3, y: 0.7, w: 12.7, h: 1.0,
    fontSize: 34, bold: true, color: C.black, fontFace: "Calibri",
  });

  s.addShape(pptx.shapes.LINE, { x: 0.3, y: 2.0, w: 12.7, h: 0, line: { color: C.ltgray, pt: 1.5 } });
  s.addText("PROPOSAL FOR :", { x: 0.3, y: 2.1, w: 5, h: 0.45, fontSize: 18, color: C.gray });
  s.addShape(pptx.shapes.LINE, { x: 0.3, y: 2.62, w: 12.7, h: 0, line: { color: C.ltgray, pt: 2 } });

  // Client name
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

function buildExecutiveSummary(pptx, data, section) {
  const s = pptx.addSlide();
  addHeader(pptx, s, "Executive Summary", 2);

  const content = typeof section.content === "string"
    ? section.content
    : section.content?.summary || section.content?.description || JSON.stringify(section.content);

  s.addText(content, {
    x: 0.5, y: 1.3, w: 12.3, h: 5.5,
    fontSize: 13, color: "333333", valign: "top", lineSpacing: 22,
    paraSpaceAfter: 8,
  });
}

function buildScopeOfWork(pptx, data, section) {
  // Divider
  const d = pptx.addSlide();
  addSectionDivider(pptx, d, "SCOPE OF WORK", 3);

  const s = pptx.addSlide();
  addHeader(pptx, s, "Scope of Work - Details", 4);

  const content = section.content;
  let items = [];

  if (typeof content === "string") {
    items = content.split("\n").filter((l) => l.trim());
  } else if (content?.items) {
    items = content.items;
  } else if (content?.scope_items) {
    items = content.scope_items;
  } else if (Array.isArray(content)) {
    items = content;
  } else {
    items = [JSON.stringify(content)];
  }

  let yPos = 1.3;
  items.forEach((item, i) => {
    if (yPos > 6.5) {
      // Overflow → new slide
      const ns = pptx.addSlide();
      addHeader(pptx, ns, "Scope of Work - Details (cont.)", null);
      yPos = 1.3;
    }
    const text = typeof item === "string" ? item : item.title || item.name || JSON.stringify(item);
    const desc = typeof item === "object" ? item.description || item.detail || "" : "";

    // Numbered item with blue accent
    s.addShape(pptx.shapes.ROUNDED_RECTANGLE, {
      x: 0.3, y: yPos, w: 0.4, h: 0.4,
      fill: { color: C.headerBlue }, rectRadius: 0.05,
    });
    s.addText(String(i + 1), {
      x: 0.3, y: yPos, w: 0.4, h: 0.4,
      fontSize: 14, bold: true, color: C.white, align: "center", valign: "middle",
    });
    s.addText(text, {
      x: 0.85, y: yPos, w: 11.5, h: 0.4,
      fontSize: 13, bold: true, color: C.black, valign: "middle",
    });
    yPos += 0.45;

    if (desc) {
      s.addText(desc, {
        x: 0.85, y: yPos, w: 11.5, h: 0.35,
        fontSize: 11, color: C.gray, valign: "top",
      });
      yPos += 0.4;
    }
    yPos += 0.15;
  });
}

function buildTimeline(pptx, data, section) {
  const s = pptx.addSlide();
  addHeader(pptx, s, "Timeline & Milestones", null);

  const content = section.content;
  let phases = [];

  if (typeof content === "string") {
    phases = content.split("\n").filter((l) => l.trim());
  } else if (content?.phases) {
    phases = content.phases;
  } else if (content?.milestones) {
    phases = content.milestones;
  } else if (Array.isArray(content)) {
    phases = content;
  } else {
    phases = [content];
  }

  // Draw chevron timeline
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

    // Phase box
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

    // Arrow between boxes
    if (col < maxPerRow - 1 && i < phases.length - 1) {
      s.addText("\u25B6", {
        x: x + 2.9, y, w: 0.3, h: 0.55,
        fontSize: 14, color, align: "center", valign: "middle",
      });
    }
  });
}

function buildTechStack(pptx, data, section) {
  const s = pptx.addSlide();
  addHeader(pptx, s, "Technology Stack", null);

  const content = section.content;
  let items = [];

  if (typeof content === "string") {
    items = content.split("\n").filter((l) => l.trim());
  } else if (Array.isArray(content)) {
    items = content;
  } else if (content?.technologies || content?.stack) {
    items = content.technologies || content.stack;
  } else {
    items = Object.entries(content || {}).map(([k, v]) => `${k}: ${v}`);
  }

  // Grid layout - 2 columns
  items.forEach((item, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.3 + col * 6.5;
    const y = 1.4 + row * 0.85;
    const text = typeof item === "string" ? item : item.name || item.technology || JSON.stringify(item);
    const desc = typeof item === "object" ? item.description || item.purpose || "" : "";

    // Icon circle
    s.addShape(pptx.shapes.OVAL, {
      x, y: y + 0.05, w: 0.4, h: 0.4,
      fill: { color: C.teal },
    });
    s.addText("\u2699", {
      x, y: y + 0.05, w: 0.4, h: 0.4,
      fontSize: 14, color: C.white, align: "center", valign: "middle",
    });

    s.addText(text, {
      x: x + 0.5, y, w: 5.5, h: 0.35,
      fontSize: 12, bold: true, color: C.black, valign: "middle",
    });
    if (desc) {
      s.addText(desc, {
        x: x + 0.5, y: y + 0.35, w: 5.5, h: 0.3,
        fontSize: 10, color: C.gray,
      });
    }
  });
}

function buildTeamStructure(pptx, data, section) {
  const s = pptx.addSlide();
  addHeader(pptx, s, "Team Structure", null);

  const content = section.content;
  let roles = [];

  if (typeof content === "string") {
    roles = content.split("\n").filter((l) => l.trim());
  } else if (Array.isArray(content)) {
    roles = content;
  } else if (content?.team || content?.roles || content?.members) {
    roles = content.team || content.roles || content.members;
  } else {
    roles = [content];
  }

  // Org chart style
  roles.forEach((role, i) => {
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

  // Table format
  let rows = [];
  if (Array.isArray(content)) {
    rows = content;
  } else if (content?.items || content?.line_items) {
    rows = content.items || content.line_items;
  } else {
    rows = [content];
  }

  // Header row
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

  // Total row if available
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

function buildGenericSlide(pptx, data, section, slideNum) {
  const s = pptx.addSlide();
  const title = section.section_title || section.title || `Section ${slideNum}`;
  addHeader(pptx, s, title, slideNum);

  const content = section.content;
  let text = "";

  if (typeof content === "string") {
    text = content;
  } else if (content?.summary) {
    text = content.summary;
    if (content.details) text += "\n\n" + (typeof content.details === "string" ? content.details : JSON.stringify(content.details, null, 2));
  } else {
    text = JSON.stringify(content, null, 2);
  }

  s.addText(text.substring(0, 2000), {
    x: 0.5, y: 1.3, w: 12.3, h: 5.7,
    fontSize: 12, color: "333333", valign: "top", lineSpacing: 20,
    paraSpaceAfter: 6,
  });
}

// ── MAIN: Generate PPTX from proposal JSON ──
async function generatePPTX(proposalData) {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "TOMAS TECH CO., LTD.";
  pptx.company = "TOMAS TECH CO., LTD.";
  pptx.subject = proposalData.proposal_title || "Proposal";

  // Parse proposal if string
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

  const sectionBuilders = {
    executive: buildExecutiveSummary,
    summary: buildExecutiveSummary,
    scope: buildScopeOfWork,
    timeline: buildTimeline,
    milestone: buildTimeline,
    schedule: buildTimeline,
    technology: buildTechStack,
    tech_stack: buildTechStack,
    stack: buildTechStack,
    team: buildTeamStructure,
    structure: buildTeamStructure,
    resource: buildTeamStructure,
    pricing: buildPricing,
    price: buildPricing,
    cost: buildPricing,
    budget: buildPricing,
    estimate: buildPricing,
    terms: buildTerms,
    condition: buildTerms,
    warranty: buildTerms,
  };

  sections.forEach((section) => {
    const title = (section.section_title || section.title || "").toLowerCase();
    let built = false;

    for (const [keyword, builder] of Object.entries(sectionBuilders)) {
      if (title.includes(keyword)) {
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

// ── API Handler ──
module.exports = async function handler(req, res) {
  // CORS
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  try {
    const { proposal_content, request_id, client_name } = req.body;

    if (!proposal_content) {
      return res.status(400).json({ error: "proposal_content is required" });
    }

    // Parse proposal content
    let proposalData = proposal_content;
    if (typeof proposalData === "string") {
      try {
        const cleaned = proposalData.replace(/```json\s*/g, "").replace(/```\s*/g, "").trim();
        const match = cleaned.match(/\{[\s\S]*\}/);
        if (match) proposalData = JSON.parse(match[0]);
      } catch (e) {
        // keep as string
      }
    }

    // Generate PPTX
    const buffer = await generatePPTX(proposalData);

    // Generate filename
    const safeName = (client_name || "proposal").replace(/[^a-zA-Z0-9ก-๙]/g, "_").substring(0, 30);
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
          // Fallback: return file directly
          res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
          res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
          return res.status(200).send(buffer);
        }

        // Get public URL
        const { data: urlData } = supabase.storage.from("proposals").getPublicUrl(`pptx/${filename}`);

        // Update proposal_requests table if request_id provided
        if (request_id) {
          await supabase
            .from("proposal_requests")
            .update({ pptx_url: urlData.publicUrl, status: "completed" })
            .eq("id", request_id);
        }

        return res.status(200).json({
          success: true,
          filename,
          pptx_url: urlData.publicUrl,
          request_id: request_id || null,
        });
      } catch (storageError) {
        console.error("Storage error:", storageError);
      }
    }

    // Fallback: return binary file if no Supabase
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    return res.status(200).send(buffer);

  } catch (error) {
    console.error("PPTX generation error:", error);
    return res.status(500).json({ error: "Failed to generate PPTX", details: error.message });
  }
};
