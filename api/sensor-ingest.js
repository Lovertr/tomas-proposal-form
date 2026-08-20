// Vercel Serverless Function: /api/sensor-ingest
// Receives sensor data, stores in Supabase, checks thresholds, sends LINE alert
//
// POST /api/sensor-ingest
// Body: { sensor_id: "PS-01", value: 9.2 }
// or batch: { readings: [{ sensor_id: "PS-01", value: 9.2 }, { sensor_id: "TS-01", value: 88 }] }
// Header: x-api-key: <INGEST_API_KEY>

const { createClient } = require("@supabase/supabase-js");

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_SERVICE_KEY
);

// n8n webhook URL — fallback if LINE_CHANNEL_ACCESS_TOKEN not set
const N8N_ALERT_WEBHOOK = process.env.N8N_ALERT_WEBHOOK_URL;

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

// ─── System Config Cache ─────────────────────────────────────────────
let reAlertIntervalMin = 5; // default, overridden by system_config
let configCacheTime = 0;
const CONFIG_CACHE_TTL = 120000; // 2 minutes

async function loadReAlertInterval() {
  if (Date.now() - configCacheTime < CONFIG_CACHE_TTL) return;
  try {
    const { data } = await supabase.from("system_config").select("value").eq("key", "re_alert_interval").single();
    if (data && data.value != null) {
      var parsed = typeof data.value === 'string' ? JSON.parse(data.value) : data.value;
      reAlertIntervalMin = parseInt(parsed, 10) || 5;
    }
  } catch (e) { /* use default */ }
  configCacheTime = Date.now();
}

// ─── Sensor Thresholds Cache ─────────────────────────────────────────
let sensorsCache = null;
let cacheTime = 0;
const CACHE_TTL = 60000; // 60 seconds

async function getSensors() {
  if (sensorsCache && Date.now() - cacheTime < CACHE_TTL) return sensorsCache;
  const { data, error } = await supabase
    .from("sensors")
    .select("*")
    .eq("is_active", true);
  if (error) throw new Error(`Failed to load sensors: ${error.message}`);
  sensorsCache = {};
  data.forEach((s) => (sensorsCache[s.id] = s));
  cacheTime = Date.now();
  return sensorsCache;
}

function checkAnomaly(sensor, value) {
  if (value >= sensor.threshold_high) {
    const severity =
      value >= sensor.threshold_high + (sensor.max_value - sensor.threshold_high) * 0.5
        ? "CRITICAL"
        : "WARNING";
    return { is_anomaly: true, direction: "HIGH", threshold: sensor.threshold_high, severity };
  }
  if (value <= sensor.threshold_low) {
    const severity =
      value <= sensor.threshold_low - (sensor.threshold_low - sensor.min_value) * 0.5
        ? "CRITICAL"
        : "WARNING";
    return { is_anomaly: true, direction: "LOW", threshold: sensor.threshold_low, severity };
  }
  return { is_anomaly: false };
}

// ─── Flex Message Builder ────────────────────────────────────────────

function buildAlertFlex(alertId, sensor, value, anomaly, timestamp, lineUserId) {
  return {
    type: "flex",
    altText: `⚠️ ALARM — ${sensor.name}`,
    contents: {
      type: "bubble",
      header: {
        type: "box",
        layout: "vertical",
        backgroundColor: anomaly.severity === "CRITICAL" ? "#D32F2F" : "#F7941D",
        contents: [
          {
            type: "text",
            text: "⚠️ " + (anomaly.severity === "CRITICAL" ? "ALARM" : "WARNING"),
            color: "#FFFFFF",
            weight: "bold",
            size: "lg",
          },
        ],
      },
      body: {
        type: "box",
        layout: "vertical",
        spacing: "sm",
        contents: [
          { type: "text", text: sensor.name, weight: "bold", size: "md" },
          { type: "separator" },
          {
            type: "box",
            layout: "horizontal",
            contents: [
              { type: "text", text: "Sensor", color: "#999999", size: "sm", flex: 2 },
              { type: "text", text: sensor.id, weight: "bold", size: "sm", flex: 3 },
            ],
          },
          {
            type: "box",
            layout: "horizontal",
            contents: [
              { type: "text", text: "ค่าปัจจุบัน", color: "#999999", size: "sm", flex: 2 },
              {
                type: "text",
                text: `${value} ${sensor.unit}`,
                weight: "bold",
                color: "#D32F2F",
                size: "sm",
                flex: 3,
              },
            ],
          },
          {
            type: "box",
            layout: "horizontal",
            contents: [
              { type: "text", text: "Threshold", color: "#999999", size: "sm", flex: 2 },
              {
                type: "text",
                text: `${anomaly.direction} ${anomaly.threshold} ${sensor.unit}`,
                size: "sm",
                flex: 3,
              },
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
              label: "✅ ยืนยันเข้าหน้างาน",
              uri: `https://consetech-monitoring.vercel.app/api/acknowledge-web?alert_id=${alertId}&line_user_id=${lineUserId}&sensor_id=${sensor.id}`,
            },
          },
        ],
      },
    },
  };
}

// ─── n8n Fallback ────────────────────────────────────────────────────

async function triggerN8nAlert(alertData) {
  if (!N8N_ALERT_WEBHOOK) {
    console.warn("N8N_ALERT_WEBHOOK_URL not configured, skipping n8n notification");
    return;
  }
  try {
    await fetch(N8N_ALERT_WEBHOOK, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(alertData),
    });
  } catch (err) {
    console.error("Failed to trigger n8n alert:", err.message);
  }
}

// ─── Main Handler ────────────────────────────────────────────────────

module.exports = async function handler(req, res) {
  // CORS
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization, x-api-key");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  // Auth check
  const apiKey = req.headers["x-api-key"];
  if (process.env.INGEST_API_KEY && apiKey !== process.env.INGEST_API_KEY) {
    return res.status(401).json({ error: "Unauthorized" });
  }

  try {
    const sensors = await getSensors();
    await loadReAlertInterval();
    const body = req.body;

    // Normalize to array
    const readings = body.readings || [{ sensor_id: body.sensor_id, value: body.value }];
    const results = [];

    for (const reading of readings) {
      const { sensor_id, value } = reading;

      // Validate
      if (!sensor_id || value === undefined || value === null) {
        results.push({ sensor_id, error: "Missing sensor_id or value" });
        continue;
      }

      const sensor = sensors[sensor_id];
      if (!sensor) {
        results.push({ sensor_id, error: "Unknown sensor" });
        continue;
      }

      const numValue = parseFloat(value);
      const anomaly = checkAnomaly(sensor, numValue);

      // 1. Insert reading
      const { data: readingData, error: readingError } = await supabase
        .from("sensor_readings")
        .insert({
          sensor_id,
          value: numValue,
          is_anomaly: anomaly.is_anomaly,
        })
        .select("id")
        .single();

      if (readingError) {
        results.push({ sensor_id, error: readingError.message });
        continue;
      }

      const result = {
        sensor_id,
        value: numValue,
        is_anomaly: anomaly.is_anomaly,
        reading_id: readingData.id,
      };

      // 2. If anomaly -> check for existing open alert
      if (anomaly.is_anomaly) {
        const { data: existingAlert } = await supabase
          .from("alerts")
          .select("id, created_at, status")
          .eq("sensor_id", sensor_id)
          .in("status", ["open", "acknowledged"])
          .order("created_at", { ascending: false })
          .limit(1)
          .single();

        if (!existingAlert) {
          // Resolve cooldown: skip creating new alert if same sensor was resolved recently
          const cooldownMs = reAlertIntervalMin * 60 * 1000;
          const cooldownAgo = new Date(Date.now() - cooldownMs).toISOString();
          const { data: recentResolve } = await supabase
            .from("alert_actions")
            .select("id, alert_id")
            .eq("action", "resolved")
            .gte("created_at", cooldownAgo)
            .limit(5);

          // Check if any recent resolve was for this sensor
          let resolvedRecently = false;
          if (recentResolve && recentResolve.length > 0) {
            const resolvedAlertIds = recentResolve.map(r => r.alert_id);
            const { data: resolvedAlerts } = await supabase
              .from("alerts")
              .select("id")
              .eq("sensor_id", sensor_id)
              .in("id", resolvedAlertIds)
              .limit(1);
            resolvedRecently = resolvedAlerts && resolvedAlerts.length > 0;
          }

          if (resolvedRecently) {
            result.alert_status = "cooldown";
            result.cooldown_minutes = reAlertIntervalMin;
          } else {
          // Create new alert
          const { data: alertData, error: alertError } = await supabase
            .from("alerts")
            .insert({
              sensor_id,
              reading_id: readingData.id,
              value: numValue,
              threshold: anomaly.threshold,
              direction: anomaly.direction,
              severity: anomaly.severity,
              status: "open",
            })
            .select("id")
            .single();

          if (!alertError && alertData) {
            // Log action
            await supabase.from("alert_actions").insert({
              alert_id: alertData.id,
              action: "created",
              note: `Anomaly detected: ${numValue} ${sensor.unit} (${anomaly.direction} threshold: ${anomaly.threshold} ${sensor.unit})`,
            });

            const timestamp = new Date().toLocaleString("th-TH", { timeZone: "Asia/Bangkok" });

            // Send LINE notification to assigned users
            if (process.env.LINE_CHANNEL_ACCESS_TOKEN) {
              try {
                // Get assigned users for this sensor
                const { data: perms } = await supabase
                  .from("user_sensor_permissions")
                  .select("user_id, users(line_user_id)")
                  .eq("sensor_id", sensor_id);

                const lineUserIds = (perms || [])
                  .map((p) => p.users?.line_user_id)
                  .filter(Boolean);

                if (lineUserIds.length > 0) {
                  // Send individual messages (each user gets URI with their own line_user_id)
                  for (const uid of lineUserIds) {
                    const flexMsg = buildAlertFlex(alertData.id, sensor, numValue, anomaly, timestamp, uid);
                    await sendLineMessage(uid, [flexMsg]);
                  }

                  // Log notification
                  await supabase.from("alert_actions").insert({
                    alert_id: alertData.id,
                    action: "notified",
                    note: `LINE notification sent to ${lineUserIds.length} user(s)`,
                  });
                }
              } catch (lineErr) {
                console.error("LINE notification error:", lineErr.message);
              }
            } else {
              // Fallback: trigger n8n
              await triggerN8nAlert({
                alert_id: alertData.id,
                sensor_id,
                sensor_name: sensor.name,
                sensor_type: sensor.type,
                value: numValue,
                unit: sensor.unit,
                threshold: anomaly.threshold,
                direction: anomaly.direction,
                severity: anomaly.severity,
                timestamp: new Date().toISOString(),
              });
            }

            result.alert_id = alertData.id;
            result.alert_status = "new";
          }
          } // end: not resolvedRecently
        } else {
          result.alert_id = existingAlert.id;
          result.alert_status = "existing";

          // Re-alert: if alert is OPEN (not acknowledged) and older than configured interval
          if (existingAlert.status === "open") {
            const reAlertMs = reAlertIntervalMin * 60 * 1000;
            const alertAgeMs = Date.now() - new Date(existingAlert.created_at).getTime();
            if (alertAgeMs > reAlertMs) {
              // Check if reminder was sent recently (within last interval)
              const intervalAgo = new Date(Date.now() - reAlertMs).toISOString();
              const { data: recentReminder } = await supabase
                .from("alert_actions")
                .select("id")
                .eq("alert_id", existingAlert.id)
                .eq("action", "reminder")
                .gte("created_at", intervalAgo)
                .limit(1);

              if (!recentReminder || recentReminder.length === 0) {
                const alertMinutes = Math.round(alertAgeMs / 60000);
                const timestamp = new Date().toLocaleString("th-TH", { timeZone: "Asia/Bangkok" });

                // Send reminder to assigned users
                if (process.env.LINE_CHANNEL_ACCESS_TOKEN) {
                  try {
                    const { data: perms } = await supabase
                      .from("user_sensor_permissions")
                      .select("user_id, users(line_user_id)")
                      .eq("sensor_id", sensor_id);
                    const lineUserIds = (perms || []).map((p) => p.users?.line_user_id).filter(Boolean);

                    for (const uid of lineUserIds) {
                      const reminderFlex = {
                        type: "flex",
                        altText: `🔴 เตือนซ้ำ — ${sensor.name} ยังไม่มีผู้ดูแล (${alertMinutes} นาที)`,
                        contents: {
                          type: "bubble",
                          header: {
                            type: "box", layout: "vertical", backgroundColor: "#B71C1C",
                            contents: [{ type: "text", text: "🔴 เตือนซ้ำ — ยังไม่มีผู้รับผิดชอบ", color: "#FFFFFF", weight: "bold", size: "md", wrap: true }],
                          },
                          body: {
                            type: "box", layout: "vertical", spacing: "sm",
                            contents: [
                              { type: "text", text: sensor.name, weight: "bold", size: "md" },
                              { type: "separator" },
                              { type: "box", layout: "horizontal", contents: [
                                { type: "text", text: "Sensor", color: "#999999", size: "sm", flex: 2 },
                                { type: "text", text: sensor.id, weight: "bold", size: "sm", flex: 3 },
                              ]},
                              { type: "box", layout: "horizontal", contents: [
                                { type: "text", text: "ค่าปัจจุบัน", color: "#999999", size: "sm", flex: 2 },
                                { type: "text", text: `${numValue} ${sensor.unit}`, weight: "bold", color: "#D32F2F", size: "sm", flex: 3 },
                              ]},
                              { type: "box", layout: "horizontal", contents: [
                                { type: "text", text: "ระยะเวลา", color: "#999999", size: "sm", flex: 2 },
                                { type: "text", text: `${alertMinutes} นาที`, weight: "bold", color: "#B71C1C", size: "sm", flex: 3 },
                              ]},
                              { type: "text", text: "⚠️ Alert นี้ยังไม่มีผู้เข้าดูแล กรุณาตรวจสอบ!", size: "sm", color: "#B71C1C", wrap: true, margin: "md" },
                            ],
                          },
                          footer: {
                            type: "box", layout: "vertical",
                            contents: [{
                              type: "button", style: "primary", color: "#B71C1C",
                              action: {
                                type: "uri",
                                label: "✅ ยืนยันเข้าหน้างาน",
                                uri: `https://consetech-monitoring.vercel.app/api/acknowledge-web?alert_id=${existingAlert.id}&line_user_id=${uid}&sensor_id=${sensor.id}`,
                              },
                            }],
                          },
                        },
                      };
                      await sendLineMessage(uid, [reminderFlex]);
                    }

                    // Log reminder
                    await supabase.from("alert_actions").insert({
                      alert_id: existingAlert.id,
                      action: "reminder",
                      note: `Reminder sent — alert open for ${alertMinutes} minutes, value: ${numValue} ${sensor.unit}`,
                    });
                    result.reminder_sent = true;
                  } catch (reminderErr) {
                    console.error("Reminder error:", reminderErr.message);
                  }
                }
              }
            }
          }
        }
      }

      results.push(result);
    }

    return res.status(200).json({
      success: true,
      count: results.length,
      results,
    });
  } catch (err) {
    console.error("Ingest error:", err);
    return res.status(500).json({ error: err.message });
  }
};
