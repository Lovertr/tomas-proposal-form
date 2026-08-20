// Vercel Serverless Function: /api/line-webhook
// LINE Messaging API webhook handler
// Handles: follow, postback (acknowledge/resolve)
//
// POST /api/line-webhook

const crypto = require("crypto");
const { createClient } = require("@supabase/supabase-js");

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_SERVICE_KEY
);

// ─── LINE API Helpers ────────────────────────────────────────────────
async function sendLineMessage(to, messages) {
  const token = process.env.LINE_CHANNEL_ACCESS_TOKEN;
  if (!token) return;
  await fetch("https://api.line.me/v2/bot/message/push", {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
    body: JSON.stringify({ to, messages }),
  });
}

async function multicastLineMessage(userIds, messages) {
  const token = process.env.LINE_CHANNEL_ACCESS_TOKEN;
  if (!token) return;
  if (userIds.length === 0) return;
  if (userIds.length === 1) return sendLineMessage(userIds[0], messages);
  await fetch("https://api.line.me/v2/bot/message/multicast", {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
    body: JSON.stringify({ to: userIds, messages }),
  });
}

async function replyLineMessage(replyToken, messages) {
  const token = process.env.LINE_CHANNEL_ACCESS_TOKEN;
  if (!token) return;
  await fetch("https://api.line.me/v2/bot/message/reply", {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: `Bearer ${token}` },
    body: JSON.stringify({ replyToken, messages }),
  });
}

// ─── Flex Message Builders ───────────────────────────────────────────

function buildAcknowledgeBroadcastFlex(sensorId, sensorName, responderName, timestamp) {
  return {
    type: "flex",
    altText: `👷 ${responderName} กำลังเข้าดูแล ${sensorName}`,
    contents: {
      type: "bubble",
      header: {
        type: "box",
        layout: "vertical",
        backgroundColor: "#1B6B93",
        contents: [
          { type: "text", text: "👷 มีผู้เข้าดูแลแล้ว", color: "#FFFFFF", weight: "bold", size: "lg" },
        ],
      },
      body: {
        type: "box",
        layout: "vertical",
        spacing: "sm",
        contents: [
          { type: "text", text: sensorName, weight: "bold", size: "md" },
          { type: "separator" },
          {
            type: "box",
            layout: "horizontal",
            contents: [
              { type: "text", text: "Sensor", color: "#999999", size: "sm", flex: 2 },
              { type: "text", text: sensorId, weight: "bold", size: "sm", flex: 3 },
            ],
          },
          {
            type: "box",
            layout: "horizontal",
            contents: [
              { type: "text", text: "ผู้เข้าดูแล", color: "#999999", size: "sm", flex: 2 },
              { type: "text", text: responderName, weight: "bold", color: "#1B6B93", size: "sm", flex: 3 },
            ],
          },
          {
            type: "box",
            layout: "horizontal",
            contents: [
              { type: "text", text: "เวลา", color: "#999999", size: "sm", flex: 2 },
              { type: "text", text: timestamp, size: "sm", flex: 3 },
            ],
          },
        ],
      },
    },
  };
}

function buildRepairInProgressFlex(alertId, sensorId, sensorName) {
  return {
    type: "flex",
    altText: `🔧 กำลังดำเนินการซ่อม ${sensorName}`,
    contents: {
      type: "bubble",
      header: {
        type: "box",
        layout: "vertical",
        backgroundColor: "#F7941D",
        contents: [
          { type: "text", text: "🔧 กำลังดำเนินการ", color: "#FFFFFF", weight: "bold", size: "lg" },
        ],
      },
      body: {
        type: "box",
        layout: "vertical",
        spacing: "sm",
        contents: [
          { type: "text", text: sensorName, weight: "bold", size: "md" },
          { type: "separator" },
          {
            type: "box",
            layout: "horizontal",
            contents: [
              { type: "text", text: "Sensor", color: "#999999", size: "sm", flex: 2 },
              { type: "text", text: sensorId, weight: "bold", size: "sm", flex: 3 },
            ],
          },
          {
            type: "text",
            text: "คุณได้ยืนยันเข้าหน้างานแล้ว\nกรุณากดเมื่อซ่อมเสร็จ",
            size: "sm",
            color: "#555555",
            wrap: true,
            margin: "md",
          },
        ],
      },
      footer: {
        type: "box",
        layout: "vertical",
        contents: [
          {
            type: "button",
            style: "primary",
            color: "#388E3C",
            action: {
              type: "postback",
              label: "✅ ซ่อมเสร็จแล้ว",
              data: `action=resolve&alert_id=${alertId}`,
            },
          },
        ],
      },
    },
  };
}

function buildResolvedBroadcastFlex(sensorId, sensorName, responderName, durationMin, timestamp) {
  return {
    type: "flex",
    altText: `✅ ${sensorId} กลับมาปกติแล้ว`,
    contents: {
      type: "bubble",
      header: {
        type: "box",
        layout: "vertical",
        backgroundColor: "#388E3C",
        contents: [
          { type: "text", text: "✅ กลับมาปกติแล้ว", color: "#FFFFFF", weight: "bold", size: "lg" },
        ],
      },
      body: {
        type: "box",
        layout: "vertical",
        spacing: "sm",
        contents: [
          { type: "text", text: sensorName, weight: "bold", size: "md" },
          { type: "separator" },
          {
            type: "box",
            layout: "horizontal",
            contents: [
              { type: "text", text: "Sensor", color: "#999999", size: "sm", flex: 2 },
              { type: "text", text: sensorId, weight: "bold", size: "sm", flex: 3 },
            ],
          },
          {
            type: "box",
            layout: "horizontal",
            contents: [
              { type: "text", text: "ซ่อมโดย", color: "#999999", size: "sm", flex: 2 },
              { type: "text", text: responderName, weight: "bold", color: "#388E3C", size: "sm", flex: 3 },
            ],
          },
          {
            type: "box",
            layout: "horizontal",
            contents: [
              { type: "text", text: "ใช้เวลา", color: "#999999", size: "sm", flex: 2 },
              { type: "text", text: `${durationMin} นาที`, weight: "bold", size: "sm", flex: 3 },
            ],
          },
          {
            type: "box",
            layout: "horizontal",
            contents: [
              { type: "text", text: "เวลา", color: "#999999", size: "sm", flex: 2 },
              { type: "text", text: timestamp, size: "sm", flex: 3 },
            ],
          },
        ],
      },
    },
  };
}

// ─── Helpers ─────────────────────────────────────────────────────────

function thaiTimestamp() {
  return new Date().toLocaleString("th-TH", { timeZone: "Asia/Bangkok" });
}

async function getAssignedLineUserIds(sensorId) {
  const { data } = await supabase
    .from("user_sensor_permissions")
    .select("user_id, users(line_user_id)")
    .eq("sensor_id", sensorId);
  if (!data) return [];
  return data.map((d) => d.users?.line_user_id).filter(Boolean);
}

async function getAssignedLineUserIdsExcept(sensorId, excludeLineUserId) {
  const all = await getAssignedLineUserIds(sensorId);
  return all.filter((id) => id !== excludeLineUserId);
}

// ─── Signature Verification ──────────────────────────────────────────

function verifySignature(body, signature) {
  const secret = process.env.LINE_CHANNEL_SECRET;
  if (!secret) return true; // Skip verification if not configured
  const hash = crypto.createHmac("SHA256", secret).update(body).digest("base64");
  return hash === signature;
}

// ─── Main Handler ────────────────────────────────────────────────────

module.exports = async function handler(req, res) {
  // CORS
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, X-Line-Signature");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  try {
    // Verify LINE signature
    const rawBody = typeof req.body === "string" ? req.body : JSON.stringify(req.body);
    const signature = req.headers["x-line-signature"];
    if (!verifySignature(rawBody, signature)) {
      return res.status(401).json({ error: "Invalid signature" });
    }

    const body = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
    const events = body.events || [];

    for (const event of events) {
      const userId = event.source?.userId;
      const replyToken = event.replyToken;

      // ─── Follow Event ──────────────────────────────────────
      if (event.type === "follow") {
        await replyLineMessage(replyToken, [
          {
            type: "text",
            text: "ยินดีต้อนรับเข้าสู่ระบบแจ้งเตือน CONSERTECH\nกรุณาลงทะเบียนที่ลิงก์ด้านล่าง",
          },
          {
            type: "flex",
            altText: "ลงทะเบียนใช้งานระบบ",
            contents: {
              type: "bubble",
              body: {
                type: "box",
                layout: "vertical",
                spacing: "md",
                contents: [
                  {
                    type: "text",
                    text: "CONSERTECH Sensor Monitor",
                    weight: "bold",
                    size: "md",
                    color: "#1B6B93",
                  },
                  {
                    type: "text",
                    text: "กรุณาลงทะเบียนเพื่อรับการแจ้งเตือนจากเซ็นเซอร์",
                    size: "sm",
                    color: "#555555",
                    wrap: true,
                  },
                ],
              },
              footer: {
                type: "box",
                layout: "vertical",
                contents: [
                  {
                    type: "button",
                    style: "primary",
                    color: "#1B6B93",
                    action: {
                      type: "uri",
                      label: "📋 ลงทะเบียน",
                      uri: "https://consetech-monitoring.vercel.app/register.html",
                    },
                  },
                ],
              },
            },
          },
        ]);
        continue;
      }

      // ─── Postback Event ────────────────────────────────────
      if (event.type === "postback") {
        const params = new URLSearchParams(event.postback.data);
        const action = params.get("action");
        const alertId = params.get("alert_id");

        if (!action || !alertId) continue;

        // Look up user
        const { data: user } = await supabase
          .from("users")
          .select("*")
          .eq("line_user_id", userId)
          .single();

        if (!user) {
          await replyLineMessage(replyToken, [
            { type: "text", text: "กรุณาลงทะเบียนก่อนใช้งาน\nhttps://consetech-monitoring.vercel.app/register.html" },
          ]);
          continue;
        }

        // ── Acknowledge ──────────────────────────────────────
        if (action === "acknowledge") {
          // Get alert
          const { data: alert } = await supabase
            .from("alerts")
            .select("*, sensors(*)")
            .eq("id", alertId)
            .single();

          if (!alert) {
            await replyLineMessage(replyToken, [
              { type: "text", text: "ไม่พบการแจ้งเตือนนี้ในระบบ" },
            ]);
            continue;
          }

          if (alert.status !== "open") {
            const statusMsg = alert.status === "acknowledged"
              ? "การแจ้งเตือนนี้มีผู้ยืนยันเข้าดูแลแล้ว"
              : "การแจ้งเตือนนี้ได้รับการแก้ไขแล้ว";
            await replyLineMessage(replyToken, [{ type: "text", text: statusMsg }]);
            continue;
          }

          const sensorId = alert.sensor_id;
          const sensorName = alert.sensors?.name || sensorId;

          // Reply FIRST (before DB ops) to ensure reply token is still valid
          await replyLineMessage(replyToken, [
            buildRepairInProgressFlex(alertId, sensorId, sensorName),
          ]);

          const now = new Date().toISOString();
          const timestamp = thaiTimestamp();

          // Update alert
          await supabase
            .from("alerts")
            .update({
              status: "acknowledged",
              acknowledged_by: user.id,
              acknowledged_at: now,
            })
            .eq("id", alertId);

          // Log action
          await supabase.from("alert_actions").insert({
            alert_id: alertId,
            user_id: user.id,
            action: "acknowledged",
            note: `${user.display_name} acknowledged via LINE postback`,
          });

          // Notify OTHER assigned users (via push, not reply)
          const otherUserIds = await getAssignedLineUserIdsExcept(sensorId, userId);
          if (otherUserIds.length > 0) {
            await multicastLineMessage(otherUserIds, [
              buildAcknowledgeBroadcastFlex(sensorId, sensorName, user.display_name, timestamp),
            ]);
          }

          continue;
        }

        // ── Resolve ──────────────────────────────────────────
        if (action === "resolve") {
          // Get alert
          const { data: alert } = await supabase
            .from("alerts")
            .select("*, sensors(*)")
            .eq("id", alertId)
            .single();

          if (!alert) {
            await replyLineMessage(replyToken, [
              { type: "text", text: "ไม่พบการแจ้งเตือนนี้ในระบบ" },
            ]);
            continue;
          }

          if (alert.status !== "acknowledged") {
            const statusMsg = alert.status === "open"
              ? "กรุณายืนยันเข้าหน้างานก่อน"
              : "การแจ้งเตือนนี้ได้รับการแก้ไขแล้ว";
            await replyLineMessage(replyToken, [{ type: "text", text: statusMsg }]);
            continue;
          }

          const sensorId = alert.sensor_id;
          const sensorName = alert.sensors?.name || sensorId;

          // Reply FIRST — ask for root cause
          await replyLineMessage(replyToken, [
            { type: "text", text: `✅ บันทึกเรียบร้อย — ${sensorName} (${sensorId}) กลับมาปกติแล้ว` },
            { type: "text", text: `📝 กรุณาพิมพ์สาเหตุของปัญหา/ความผิดปกติ\n(เช่น "วาล์วรั่ว", "เซ็นเซอร์เสื่อม" ฯลฯ)` },
          ]);

          const now = new Date().toISOString();
          const timestamp = thaiTimestamp();

          // Calculate repair duration
          const ackTime = new Date(alert.acknowledged_at || alert.created_at);
          const resolveTime = new Date(now);
          const durationMin = Math.round((resolveTime - ackTime) / 60000);

          // Update alert
          await supabase
            .from("alerts")
            .update({
              status: "resolved",
              resolved_by: user.id,
              resolved_at: now,
              response_time_min: durationMin,
            })
            .eq("id", alertId);

          // Log action
          await supabase.from("alert_actions").insert({
            alert_id: alertId,
            user_id: user.id,
            action: "resolved",
            note: `Resolved by ${user.display_name} in ${durationMin} min`,
          });

          // Notify ALL assigned users
          const allUserIds = await getAssignedLineUserIds(sensorId);
          if (allUserIds.length > 0) {
            await multicastLineMessage(allUserIds, [
              buildResolvedBroadcastFlex(sensorId, sensorName, user.display_name, durationMin, timestamp),
            ]);
          }

          continue;
        }
      }

      // ─── Message Event — capture root cause or fallback ───
      if (event.type === "message" && event.message?.type === "text") {
        const msgText = event.message.text.trim();

        // Look up user
        const { data: msgUser } = await supabase
          .from("users")
          .select("id, display_name")
          .eq("line_user_id", userId)
          .single();

        if (msgUser) {
          // Check for recently resolved alerts by this user that have no root_cause
          const { data: pendingAlert } = await supabase
            .from("alerts")
            .select("id, sensor_id")
            .eq("resolved_by", msgUser.id)
            .eq("status", "resolved")
            .is("root_cause", null)
            .order("resolved_at", { ascending: false })
            .limit(1)
            .single();

          if (pendingAlert) {
            // Save root cause
            await supabase
              .from("alerts")
              .update({ root_cause: msgText })
              .eq("id", pendingAlert.id);

            await replyLineMessage(replyToken, [
              { type: "text", text: `📝 บันทึกสาเหตุเรียบร้อยแล้ว\nSensor: ${pendingAlert.sensor_id}\nสาเหตุ: ${msgText}` },
            ]);
            continue;
          }
        }

        // Default fallback
        await replyLineMessage(replyToken, [
          {
            type: "text",
            text: "ระบบ CONSERTECH Sensor Monitor\nใช้สำหรับรับการแจ้งเตือนจากเซ็นเซอร์เท่านั้น\n\nหากยังไม่ได้ลงทะเบียน กรุณาลงทะเบียนที่:\nhttps://consetech-monitoring.vercel.app/register.html",
          },
        ]);
        continue;
      }
    }

    return res.status(200).json({ success: true });
  } catch (err) {
    console.error("LINE webhook error:", err);
    return res.status(500).json({ error: err.message });
  }
};
