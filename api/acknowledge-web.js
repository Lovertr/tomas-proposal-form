// Vercel Serverless Function: /api/acknowledge-web
// Web-based acknowledge endpoint for LINE button clicks
// GET /api/acknowledge-web?alert_id=xxx&line_user_id=Uxxx&sensor_id=PT-01

const { createClient } = require("@supabase/supabase-js");

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_SERVICE_KEY
);

const N8N_NOTIFY_WEBHOOK = process.env.N8N_NOTIFY_WEBHOOK_URL;

module.exports = async function handler(req, res) {
  if (req.method !== "GET") return res.status(405).send("Method not allowed");

  const { alert_id, line_user_id, sensor_id } = req.query;

  if (!alert_id || !line_user_id) {
    return res.status(400).send(renderPage("Error", "Missing parameters", "error"));
  }

  try {
    // Look up user
    const { data: user } = await supabase
      .from("users")
      .select("*")
      .eq("line_user_id", line_user_id)
      .single();

    if (!user) {
      return res.status(404).send(renderPage("ไม่พบผู้ใช้", "LINE User ID ไม่ตรงกับในระบบ", "error"));
    }

    // Get alert
    const { data: alert } = await supabase
      .from("alerts")
      .select("*, sensors(*)")
      .eq("id", alert_id)
      .single();

    if (!alert) {
      return res.status(404).send(renderPage("ไม่พบการแจ้งเตือน", "Alert ID ไม่ถูกต้อง", "error"));
    }

    if (alert.status === "acknowledged") {
      return res.status(200).send(renderPage(
        "รับทราบแล้ว",
        `การแจ้งเตือนนี้ถูก acknowledge แล้วโดย ${alert.acknowledged_by || 'ทีมงาน'}`,
        "info"
      ));
    }

    if (alert.status === "resolved") {
      return res.status(200).send(renderPage(
        "แก้ไขเสร็จแล้ว",
        "การแจ้งเตือนนี้ได้รับการแก้ไขเรียบร้อยแล้ว",
        "success"
      ));
    }

    const now = new Date().toISOString();

    // Acknowledge the alert
    await supabase.from("alerts").update({
      status: "acknowledged",
      acknowledged_by: user.id,
      acknowledged_at: now,
    }).eq("id", alert_id);

    // Log action
    await supabase.from("alert_actions").insert({
      alert_id,
      user_id: user.id,
      action: "acknowledged",
      note: `${user.display_name} acknowledged via LINE button`,
    });

    // Notify n8n for broadcast
    if (N8N_NOTIFY_WEBHOOK) {
      try {
        await fetch(N8N_NOTIFY_WEBHOOK, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            type: "acknowledge_broadcast",
            alert_id,
            sensor_id: alert.sensor_id,
            sensor_name: alert.sensors?.name,
            sensor_type: alert.sensors?.type,
            unit: alert.sensors?.unit,
            responder_name: user.display_name,
            responder_line_user_id: line_user_id,
            ack_time: now,
          }),
        });
      } catch (e) {
        console.error("n8n notify failed:", e.message);
      }
    }

    return res.status(200).send(renderPage(
      "ยืนยันสำเร็จ ✅",
      `${user.display_name} ได้ยืนยันเข้าหน้างานเพื่อตรวจสอบ ${sensor_id || alert.sensor_id} แล้ว<br><br>ระบบจะแจ้งทีมงานทุกคนทราบ`,
      "success"
    ));

  } catch (err) {
    console.error("Acknowledge-web error:", err);
    return res.status(500).send(renderPage("Error", err.message, "error"));
  }
};

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
