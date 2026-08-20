// Vercel Serverless Function: /api/acknowledge-web
// Web-based acknowledge + resolve endpoint for LINE button clicks
// GET /api/acknowledge-web?alert_id=xxx&line_user_id=Uxxx&sensor_id=PT-01
// GET /api/acknowledge-web?action=resolve&alert_id=xxx&line_user_id=Uxxx  → show resolve form
// POST /api/acknowledge-web  → submit resolve with root_cause

const { createClient } = require("@supabase/supabase-js");

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_SERVICE_KEY
);

const N8N_NOTIFY_WEBHOOK = process.env.N8N_NOTIFY_WEBHOOK_URL;

module.exports = async function handler(req, res) {
  // Handle POST (resolve form submission)
  if (req.method === "POST") {
    return handleResolveSubmit(req, res);
  }

  if (req.method !== "GET") return res.status(405).send("Method not allowed");

  const { alert_id, line_user_id, sensor_id, action } = req.query;

  if (!alert_id || !line_user_id) {
    return res.status(400).send(renderPage("Error", "Missing parameters", "error"));
  }

  // Route: resolve form
  if (action === "resolve") {
    return showResolveForm(req, res);
  }

  // Route: acknowledge (default)
  return handleAcknowledge(req, res);
};

// ─── Acknowledge Handler ────────────────────────────────────────────
async function handleAcknowledge(req, res) {
  const { alert_id, line_user_id, sensor_id } = req.query;

  try {
    const { data: user } = await supabase
      .from("users").select("*").eq("line_user_id", line_user_id).single();
    if (!user) return res.status(404).send(renderPage("ไม่พบผู้ใช้", "LINE User ID ไม่ตรงกับในระบบ", "error"));

    const { data: alert } = await supabase
      .from("alerts").select("*, sensors(*)").eq("id", alert_id).single();
    if (!alert) return res.status(404).send(renderPage("ไม่พบการแจ้งเตือน", "Alert ID ไม่ถูกต้อง", "error"));

    if (alert.status === "acknowledged") {
      let ackName = 'ทีมงาน';
      if (alert.acknowledged_by) {
        const { data: ackUser } = await supabase.from("users").select("display_name").eq("id", alert.acknowledged_by).single();
        if (ackUser) ackName = ackUser.display_name;
      }
      return res.status(200).send(renderPage("รับทราบแล้ว", `การแจ้งเตือนนี้ถูก acknowledge แล้วโดย ${ackName}`, "info"));
    }

    if (alert.status === "resolved") {
      return res.status(200).send(renderPage("แก้ไขเสร็จแล้ว", "การแจ้งเตือนนี้ได้รับการแก้ไขเรียบร้อยแล้ว", "success"));
    }

    const now = new Date().toISOString();

    await supabase.from("alerts").update({
      status: "acknowledged", acknowledged_by: user.id, acknowledged_at: now,
    }).eq("id", alert_id);

    await supabase.from("alert_actions").insert({
      alert_id, user_id: user.id, action: "acknowledged",
      note: `${user.display_name} acknowledged via LINE button`,
    });

    // Notify n8n
    if (N8N_NOTIFY_WEBHOOK) {
      try {
        await fetch(N8N_NOTIFY_WEBHOOK, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            type: "acknowledge_broadcast", alert_id,
            sensor_id: alert.sensor_id, sensor_name: alert.sensors?.name,
            sensor_type: alert.sensors?.type, unit: alert.sensors?.unit,
            responder_name: user.display_name, responder_line_user_id: line_user_id,
            ack_time: now,
          }),
        });
      } catch (e) { console.error("n8n notify failed:", e.message); }
    }

    // Send LINE push with "ซ่อมสำเร็จ" URI button
    const lineToken = process.env.LINE_CHANNEL_ACCESS_TOKEN;
    if (lineToken) {
      const sid = sensor_id || alert.sensor_id;
      const sName = alert.sensors?.name || sid;
      const resolveUri = `https://tomas-proposal-form.vercel.app/api/acknowledge-web?action=resolve&alert_id=${alert_id}&line_user_id=${line_user_id}&sensor_id=${sid}`;
      try {
        const resolveMsg = {
          type: "flex",
          altText: `🔧 กำลังดำเนินการซ่อม ${sName}`,
          contents: {
            type: "bubble",
            header: {
              type: "box", layout: "vertical", backgroundColor: "#F7941D",
              contents: [{ type: "text", text: "🔧 กำลังดำเนินการ", color: "#FFFFFF", weight: "bold", size: "lg" }],
            },
            body: {
              type: "box", layout: "vertical", spacing: "sm",
              contents: [
                { type: "text", text: sName, weight: "bold", size: "md" },
                { type: "separator" },
                { type: "box", layout: "horizontal", contents: [
                  { type: "text", text: "Sensor", color: "#999999", size: "sm", flex: 2 },
                  { type: "text", text: sid, weight: "bold", size: "sm", flex: 3 },
                ]},
                { type: "text", text: "คุณได้ยืนยันเข้าหน้างานแล้ว\nกรุณากดเมื่อซ่อมเสร็จ", size: "sm", color: "#555555", wrap: true, margin: "md" },
              ],
            },
            footer: {
              type: "box", layout: "vertical",
              contents: [{
                type: "button", style: "primary", color: "#388E3C",
                action: { type: "uri", label: "✅ ซ่อมเสร็จแล้ว", uri: resolveUri },
              }],
            },
          },
        };
        const pushRes = await fetch("https://api.line.me/v2/bot/message/push", {
          method: "POST",
          headers: { "Content-Type": "application/json", Authorization: `Bearer ${lineToken}` },
          body: JSON.stringify({ to: line_user_id, messages: [resolveMsg] }),
        });
        if (!pushRes.ok) {
          console.error("LINE push resolve FAILED:", pushRes.status, await pushRes.text());
        }
      } catch (e) { console.error("LINE push error:", e.message); }
    }

    // Broadcast to other assigned users
    try {
      const sid = sensor_id || alert.sensor_id;
      const sName = alert.sensors?.name || sid;
      const ts = new Date().toLocaleString("th-TH", { timeZone: "Asia/Bangkok" });
      const lineToken2 = process.env.LINE_CHANNEL_ACCESS_TOKEN;
      if (lineToken2) {
        const { data: perms } = await supabase.from("user_sensor_permissions").select("user_id, users(line_user_id)").eq("sensor_id", sid);
        if (perms) {
          const otherIds = perms.map(p => p.users?.line_user_id).filter(id => id && id !== line_user_id);
          if (otherIds.length > 0) {
            const broadcastMsg = {
              type: "flex", altText: `👷 ${user.display_name} กำลังเข้าดูแล ${sName}`,
              contents: {
                type: "bubble",
                header: { type: "box", layout: "vertical", backgroundColor: "#1B6B93",
                  contents: [{ type: "text", text: "👷 มีผู้เข้าดูแลแล้ว", color: "#FFFFFF", weight: "bold", size: "lg" }] },
                body: { type: "box", layout: "vertical", spacing: "sm",
                  contents: [
                    { type: "text", text: sName, weight: "bold", size: "md" },
                    { type: "separator" },
                    { type: "box", layout: "horizontal", contents: [
                      { type: "text", text: "Sensor", color: "#999999", size: "sm", flex: 2 },
                      { type: "text", text: sid, weight: "bold", size: "sm", flex: 3 },
                    ]},
                    { type: "box", layout: "horizontal", contents: [
                      { type: "text", text: "ผู้เข้าดูแล", color: "#999999", size: "sm", flex: 2 },
                      { type: "text", text: user.display_name, weight: "bold", color: "#1B6B93", size: "sm", flex: 3 },
                    ]},
                    { type: "box", layout: "horizontal", contents: [
                      { type: "text", text: "เวลา", color: "#999999", size: "sm", flex: 2 },
                      { type: "text", text: ts, size: "sm", flex: 3 },
                    ]},
                  ],
                },
              },
            };
            const pushFn = otherIds.length === 1
              ? fetch("https://api.line.me/v2/bot/message/push", { method: "POST", headers: { "Content-Type": "application/json", Authorization: `Bearer ${lineToken2}` }, body: JSON.stringify({ to: otherIds[0], messages: [broadcastMsg] }) })
              : fetch("https://api.line.me/v2/bot/message/multicast", { method: "POST", headers: { "Content-Type": "application/json", Authorization: `Bearer ${lineToken2}` }, body: JSON.stringify({ to: otherIds, messages: [broadcastMsg] }) });
            await pushFn;
          }
        }
      }
    } catch (e) { console.error("Broadcast failed:", e.message); }

    return res.status(200).send(renderPage(
      "ยืนยันสำเร็จ ✅",
      `${user.display_name} ได้ยืนยันเข้าหน้างานเพื่อตรวจสอบ ${sensor_id || alert.sensor_id} แล้ว<br><br>ระบบจะแจ้งทีมงานทุกคนทราบ<br>กรุณากลับไปที่ LINE เพื่อกดซ่อมสำเร็จเมื่อเสร็จงาน`,
      "success"
    ));
  } catch (err) {
    console.error("Acknowledge-web error:", err);
    return res.status(500).send(renderPage("Error", err.message, "error"));
  }
}

// ─── Resolve Form ───────────────────────────────────────────────────
async function showResolveForm(req, res) {
  const { alert_id, line_user_id, sensor_id } = req.query;

  try {
    const { data: user } = await supabase.from("users").select("*").eq("line_user_id", line_user_id).single();
    if (!user) return res.status(404).send(renderPage("ไม่พบผู้ใช้", "LINE User ID ไม่ตรงกับในระบบ", "error"));

    const { data: alert } = await supabase.from("alerts").select("*, sensors(*)").eq("id", alert_id).single();
    if (!alert) return res.status(404).send(renderPage("ไม่พบการแจ้งเตือน", "Alert ID ไม่ถูกต้อง", "error"));

    if (alert.status === "resolved") {
      return res.status(200).send(renderPage("แก้ไขเสร็จแล้ว", "การแจ้งเตือนนี้ได้รับการแก้ไขเรียบร้อยแล้ว", "success"));
    }

    const sid = sensor_id || alert.sensor_id;
    const sName = alert.sensors?.name || sid;

    return res.status(200).send(renderResolveForm(alert_id, line_user_id, sid, sName, user.display_name));
  } catch (err) {
    console.error("Resolve form error:", err);
    return res.status(500).send(renderPage("Error", err.message, "error"));
  }
}

// ─── Resolve Submit (POST) ──────────────────────────────────────────
async function handleResolveSubmit(req, res) {
  // CORS
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");

  const { alert_id, line_user_id, root_cause } = req.body || {};

  if (!alert_id || !line_user_id) {
    return res.status(400).send(renderPage("Error", "Missing parameters", "error"));
  }

  try {
    const { data: user } = await supabase.from("users").select("*").eq("line_user_id", line_user_id).single();
    if (!user) return res.status(404).send(renderPage("ไม่พบผู้ใช้", "LINE User ID ไม่ตรงกับในระบบ", "error"));

    const { data: alert } = await supabase.from("alerts").select("*, sensors(*)").eq("id", alert_id).single();
    if (!alert) return res.status(404).send(renderPage("ไม่พบการแจ้งเตือน", "Alert ID ไม่ถูกต้อง", "error"));

    if (alert.status === "resolved") {
      return res.status(200).send(renderPage("แก้ไขเสร็จแล้ว", "การแจ้งเตือนนี้ได้รับการแก้ไขเรียบร้อยแล้ว", "success"));
    }

    const now = new Date().toISOString();
    const ackAt = alert.acknowledged_at || now;
    const responseMin = Math.round((new Date(now) - new Date(alert.created_at)) / 60000);

    // Update THIS alert
    await supabase.from("alerts").update({
      status: "resolved",
      resolved_by: user.id,
      resolved_at: now,
      root_cause: root_cause || null,
      response_time_min: responseMin,
    }).eq("id", alert_id);

    // ALSO resolve ALL other open/acknowledged alerts for the same sensor
    // This prevents stale alerts from keeping the sensor in ALARM state
    const { data: otherAlerts } = await supabase
      .from("alerts")
      .select("id")
      .eq("sensor_id", alert.sensor_id)
      .in("status", ["open", "acknowledged"])
      .neq("id", alert_id);

    if (otherAlerts && otherAlerts.length > 0) {
      const otherIds = otherAlerts.map(a => a.id);
      await supabase.from("alerts").update({
        status: "resolved",
        resolved_by: user.id,
        resolved_at: now,
        root_cause: root_cause ? `(auto) ${root_cause}` : "(auto) resolved with related alert",
      }).in("id", otherIds);

      // Log auto-resolve for each
      for (const oid of otherIds) {
        await supabase.from("alert_actions").insert({
          alert_id: oid, user_id: user.id, action: "resolved",
          note: `Auto-resolved: related alert ${alert_id} was resolved by ${user.display_name}`,
        });
      }
    }

    // Log action for the main alert
    await supabase.from("alert_actions").insert({
      alert_id, user_id: user.id, action: "resolved",
      note: `${user.display_name} resolved via LINE. ${root_cause ? 'สาเหตุ: ' + root_cause : 'ไม่ได้ระบุสาเหตุ'}`,
    });

    // Reset sensor to normal value so dashboard shows NORMAL
    const sensorConfig = alert.sensors;
    if (sensorConfig) {
      const normalVal = sensorConfig.threshold_low != null
        ? ((sensorConfig.threshold_low + sensorConfig.threshold_high) / 2)
        : ((sensorConfig.min_value || 0) + ((sensorConfig.threshold_high || 50) - (sensorConfig.min_value || 0)) * 0.6);
      const { error: readingErr } = await supabase.from("sensor_readings").insert({
        sensor_id: alert.sensor_id,
        value: parseFloat(normalVal.toFixed(1)),
        is_anomaly: false,
      });
      if (readingErr) console.error("Failed to insert normal reading:", readingErr.message);
    } else {
      // Fallback: insert a mid-range value even without sensor config
      console.warn("No sensor config found for", alert.sensor_id, "— inserting fallback normal reading");
      await supabase.from("sensor_readings").insert({
        sensor_id: alert.sensor_id,
        value: 0,
        is_anomaly: false,
      });
    }

    // Send LINE confirmation to resolver
    const lineToken = process.env.LINE_CHANNEL_ACCESS_TOKEN;
    if (lineToken) {
      const sid = alert.sensor_id;
      const sName = alert.sensors?.name || sid;
      try {
        const confirmMsg = {
          type: "flex", altText: `✅ แก้ไข ${sName} เสร็จแล้ว`,
          contents: {
            type: "bubble",
            header: { type: "box", layout: "vertical", backgroundColor: "#388E3C",
              contents: [{ type: "text", text: "✅ แก้ไขเสร็จสิ้น", color: "#FFFFFF", weight: "bold", size: "lg" }] },
            body: { type: "box", layout: "vertical", spacing: "sm",
              contents: [
                { type: "text", text: sName, weight: "bold", size: "md" },
                { type: "separator" },
                { type: "box", layout: "horizontal", contents: [
                  { type: "text", text: "Sensor", color: "#999999", size: "sm", flex: 2 },
                  { type: "text", text: sid, weight: "bold", size: "sm", flex: 3 },
                ]},
                { type: "box", layout: "horizontal", contents: [
                  { type: "text", text: "แก้ไขโดย", color: "#999999", size: "sm", flex: 2 },
                  { type: "text", text: user.display_name, weight: "bold", color: "#388E3C", size: "sm", flex: 3 },
                ]},
                { type: "box", layout: "horizontal", contents: [
                  { type: "text", text: "เวลาตอบสนอง", color: "#999999", size: "sm", flex: 2 },
                  { type: "text", text: `${responseMin} นาที`, size: "sm", flex: 3 },
                ]},
                ...(root_cause ? [{ type: "box", layout: "horizontal", margin: "sm", contents: [
                  { type: "text", text: "สาเหตุ", color: "#999999", size: "sm", flex: 2 },
                  { type: "text", text: root_cause, size: "sm", flex: 3, wrap: true },
                ]}] : []),
              ],
            },
          },
        };
        await fetch("https://api.line.me/v2/bot/message/push", {
          method: "POST",
          headers: { "Content-Type": "application/json", Authorization: `Bearer ${lineToken}` },
          body: JSON.stringify({ to: line_user_id, messages: [confirmMsg] }),
        });

        // Broadcast resolve to other users
        const { data: perms } = await supabase.from("user_sensor_permissions").select("user_id, users(line_user_id)").eq("sensor_id", sid);
        if (perms) {
          const otherIds = perms.map(p => p.users?.line_user_id).filter(id => id && id !== line_user_id);
          if (otherIds.length > 0) {
            const broadcastResolve = {
              type: "flex", altText: `✅ ${user.display_name} แก้ไข ${sName} เสร็จแล้ว`,
              contents: {
                type: "bubble",
                header: { type: "box", layout: "vertical", backgroundColor: "#388E3C",
                  contents: [{ type: "text", text: "✅ แก้ไขเสร็จสิ้น", color: "#FFFFFF", weight: "bold", size: "lg" }] },
                body: { type: "box", layout: "vertical", spacing: "sm",
                  contents: [
                    { type: "text", text: sName, weight: "bold", size: "md" },
                    { type: "separator" },
                    { type: "box", layout: "horizontal", contents: [
                      { type: "text", text: "แก้ไขโดย", color: "#999999", size: "sm", flex: 2 },
                      { type: "text", text: user.display_name, weight: "bold", color: "#388E3C", size: "sm", flex: 3 },
                    ]},
                    { type: "box", layout: "horizontal", contents: [
                      { type: "text", text: "เวลาตอบสนอง", color: "#999999", size: "sm", flex: 2 },
                      { type: "text", text: `${responseMin} นาที`, size: "sm", flex: 3 },
                    ]},
                    ...(root_cause ? [{ type: "box", layout: "horizontal", margin: "sm", contents: [
                      { type: "text", text: "สาเหตุ", color: "#999999", size: "sm", flex: 2 },
                      { type: "text", text: root_cause, size: "sm", flex: 3, wrap: true },
                    ]}] : []),
                  ],
                },
              },
            };
            if (otherIds.length === 1) {
              await fetch("https://api.line.me/v2/bot/message/push", { method: "POST", headers: { "Content-Type": "application/json", Authorization: `Bearer ${lineToken}` }, body: JSON.stringify({ to: otherIds[0], messages: [broadcastResolve] }) });
            } else {
              await fetch("https://api.line.me/v2/bot/message/multicast", { method: "POST", headers: { "Content-Type": "application/json", Authorization: `Bearer ${lineToken}` }, body: JSON.stringify({ to: otherIds, messages: [broadcastResolve] }) });
            }
          }
        }
      } catch (e) { console.error("Resolve LINE error:", e.message); }
    }

    return res.status(200).send(renderPage(
      "ซ่อมเสร็จแล้ว ✅",
      `${user.display_name} ได้ยืนยันซ่อม ${alert.sensor_id} เสร็จเรียบร้อย<br>เวลาตอบสนอง: ${responseMin} นาที${root_cause ? '<br>สาเหตุ: ' + root_cause : ''}`,
      "success"
    ));
  } catch (err) {
    console.error("Resolve error:", err);
    return res.status(500).send(renderPage("Error", err.message, "error"));
  }
}

// ─── Render Helpers ─────────────────────────────────────────────────

function renderPage(title, message, type) {
  const colors = { success: '#388E3C', error: '#D32F2F', info: '#1B6B93' };
  const icons = { success: '✅', error: '❌', info: 'ℹ️' };
  const color = colors[type] || colors.info;
  return `<!DOCTYPE html>
<html lang="th"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>${title} — CONSERTECH Sensor Monitor</title>
<style>
*{margin:0;padding:0;box-sizing:border-box;}
body{font-family:'Segoe UI',Tahoma,sans-serif;background:#EFF3F7;display:flex;align-items:center;justify-content:center;min-height:100vh;}
.card{background:white;border-radius:16px;padding:40px 32px;text-align:center;max-width:380px;width:90%;box-shadow:0 4px 20px rgba(0,0,0,0.1);}
.icon{font-size:48px;margin-bottom:16px;}
h1{color:${color};font-size:20px;margin-bottom:12px;}
p{color:#64748B;font-size:14px;line-height:1.6;}
.brand{margin-top:24px;padding-top:16px;border-top:1px solid #E2E8F0;font-size:10px;color:#94A3B8;letter-spacing:1px;}
.back{display:inline-block;margin-top:16px;padding:10px 24px;background:${color};color:white;text-decoration:none;border-radius:8px;font-size:13px;font-weight:600;}
</style></head>
<body><div class="card">
<div class="icon">${icons[type] || '📋'}</div>
<h1>${title}</h1>
<p>${message}</p>
<a class="back" href="https://tomas-proposal-form.vercel.app">กลับหน้า Dashboard</a>
<div class="brand">CONSERTECH CO., LTD.<br>Utility Monitoring & Alert System</div>
</div></body></html>`;
}

function renderResolveForm(alertId, lineUserId, sensorId, sensorName, userName) {
  return `<!DOCTYPE html>
<html lang="th"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>ยืนยันซ่อมสำเร็จ — CONSERTECH</title>
<style>
*{margin:0;padding:0;box-sizing:border-box;}
body{font-family:'Segoe UI',Tahoma,sans-serif;background:#EFF3F7;display:flex;align-items:center;justify-content:center;min-height:100vh;padding:16px;}
.card{background:white;border-radius:16px;padding:32px 24px;max-width:400px;width:100%;box-shadow:0 4px 20px rgba(0,0,0,0.1);}
.header{text-align:center;margin-bottom:20px;}
.header .icon{font-size:40px;margin-bottom:8px;}
.header h1{color:#388E3C;font-size:18px;}
.info{background:#F0F9F4;border-radius:8px;padding:12px;margin-bottom:16px;}
.info-row{display:flex;justify-content:space-between;margin-bottom:4px;font-size:13px;}
.info-row .label{color:#999;}
.info-row .val{font-weight:600;color:#333;}
label{display:block;font-size:13px;color:#555;margin-bottom:6px;font-weight:600;}
textarea{width:100%;border:1px solid #D1D5DB;border-radius:8px;padding:10px;font-size:14px;font-family:inherit;resize:vertical;min-height:80px;margin-bottom:16px;}
textarea:focus{outline:none;border-color:#388E3C;box-shadow:0 0 0 2px rgba(56,142,60,0.2);}
.submit-btn{width:100%;padding:12px;background:#388E3C;color:white;border:none;border-radius:8px;font-size:15px;font-weight:700;cursor:pointer;}
.submit-btn:hover{background:#2E7D32;}
.submit-btn:disabled{background:#999;cursor:not-allowed;}
.brand{text-align:center;margin-top:20px;padding-top:12px;border-top:1px solid #E2E8F0;font-size:10px;color:#94A3B8;letter-spacing:1px;}
</style></head>
<body>
<div class="card">
  <div class="header">
    <div class="icon">🔧</div>
    <h1>ยืนยันซ่อมสำเร็จ</h1>
  </div>
  <div class="info">
    <div class="info-row"><span class="label">Sensor</span><span class="val">${sensorId}</span></div>
    <div class="info-row"><span class="label">ชื่อ</span><span class="val">${sensorName}</span></div>
    <div class="info-row"><span class="label">ผู้ซ่อม</span><span class="val">${userName}</span></div>
  </div>
  <form id="resolveForm">
    <label for="root_cause">สาเหตุของปัญหา (ไม่บังคับ)</label>
    <textarea id="root_cause" name="root_cause" placeholder="เช่น วาล์วรั่ว, เซนเซอร์เสีย, สายหลุด..."></textarea>
    <button type="submit" class="submit-btn" id="submitBtn">✅ ยืนยันซ่อมสำเร็จ</button>
  </form>
  <div class="brand">CONSERTECH CO., LTD.<br>Utility Monitoring & Alert System</div>
</div>
<script>
document.getElementById('resolveForm').addEventListener('submit', async function(e) {
  e.preventDefault();
  const btn = document.getElementById('submitBtn');
  btn.disabled = true;
  btn.textContent = 'กำลังบันทึก...';
  try {
    const res = await fetch('/api/acknowledge-web', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        alert_id: '${alertId}',
        line_user_id: '${lineUserId}',
        root_cause: document.getElementById('root_cause').value.trim()
      })
    });
    const html = await res.text();
    document.open();
    document.write(html);
    document.close();
  } catch (err) {
    btn.disabled = false;
    btn.textContent = '✅ ยืนยันซ่อมสำเร็จ';
    alert('เกิดข้อผิดพลาด: ' + err.message);
  }
});
</script>
</body></html>`;
}
