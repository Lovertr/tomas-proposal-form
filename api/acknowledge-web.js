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
      // Look up acknowledger's name
      let ackName = 'ทีมงาน';
      if (alert.acknowledged_by) {
        const { data: ackUser } = await supabase
          .from("users")
          .select("display_name")
          .eq("id", alert.acknowledged_by)
          .single();
        if (ackUser) ackName = ackUser.display_name;
      }
      return res.status(200).send(renderPage(
        "รับทราบแล้ว",
        `การแจ้งเตือนนี้ถูก acknowledge แล้วโดย ${ackName}`,
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

    // Send LINE push with "ซ่อมสำเร็จ" button to the acknowledger
    const lineToken = process.env.LINE_CHANNEL_ACCESS_TOKEN;
    console.log("=== ACK LINE PUSH ===", "token?", !!lineToken, "len:", lineToken ? lineToken.length : 0, "to:", line_user_id, "alert:", alert_id);
    if (lineToken) {
      const sid = sensor_id || alert.sensor_id;
      const sName = alert.sensors?.name || sid;
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
                action: { type: "postback", label: "✅ ซ่อมเสร็จแล้ว", data: `action=resolve&alert_id=${alert_id}` },
              }],
            },
          },
        };
        console.log("LINE push resolve button to:", line_user_id, "alert:", alert_id);
        const pushRes = await fetch("https://api.line.me/v2/bot/message/push", {
          method: "POST",
          headers: { "Content-Type": "application/json", Authorization: `Bearer ${lineToken}` },
          body: JSON.stringify({ to: line_user_id, messages: [resolveMsg] }),
        });
        if (!pushRes.ok) {
          const errBody = await pushRes.text();
          console.error("LINE push resolve FAILED:", pushRes.status, errBody);
        } else {
          console.log("LINE push resolve OK:", pushRes.status);
        }
      } catch (e) { console.error("LINE push error:", e.message); }
    }

    // Also broadcast acknowledge to other assigned users
    try {
      const sid = sensor_id || alert.sensor_id;
      const sName = alert.sensors?.name || sid;
      const ts = new Date().toLocaleString("th-TH", { timeZone: "Asia/Bangkok" });
      const { data: perms } = await supabase
        .from("user_sensor_permissions")
        .select("user_id, users(line_user_id)")
        .eq("sensor_id", sid);
      if (perms) {
        const otherIds = perms
          .map(p => p.users?.line_user_id)
          .filter(id => id && id !== line_user_id);
        if (otherIds.length > 0) {
          const broadcastMsg = {
            type: "flex",
            altText: `👷 ${user.display_name} กำลังเข้าดูแล ${sName}`,
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
            ? fetch("https://api.line.me/v2/bot/message/push", {
                method: "POST",
                headers: { "Content-Type": "application/json", Authorization: `Bearer ${lineToken}` },
                body: JSON.stringify({ to: otherIds[0], messages: [broadcastMsg] }),
              })
            : fetch("https://api.line.me/v2/bot/message/multicast", {
                method: "POST",
                headers: { "Content-Type": "application/json", Authorization: `Bearer ${lineToken}` },
                body: JSON.stringify({ to: otherIds, messages: [broadcastMsg] }),
              });
          await pushFn;
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
