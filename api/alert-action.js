// Vercel Serverless Function: /api/alert-action
// Handles acknowledge and resolve actions from LINE Postback
//
// POST /api/alert-action
// Body: { action: "acknowledge"|"resolve", alert_id: "uuid", line_user_id: "U...", sensor_id: "PS-01" }

const { createClient } = require("@supabase/supabase-js");

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_SERVICE_KEY
);

const N8N_NOTIFY_WEBHOOK = process.env.N8N_NOTIFY_WEBHOOK_URL;

async function notifyN8n(payload) {
  if (!N8N_NOTIFY_WEBHOOK) return;
  try {
    await fetch(N8N_NOTIFY_WEBHOOK, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
  } catch (err) {
    console.error("Failed to notify n8n:", err.message);
  }
}

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  try {
    const { action, alert_id, line_user_id, sensor_id } = req.body;

    if (!action || !alert_id || !line_user_id) {
      return res.status(400).json({ error: "Missing action, alert_id, or line_user_id" });
    }

    // Look up user
    const { data: user } = await supabase
      .from("users")
      .select("*")
      .eq("line_user_id", line_user_id)
      .single();

    if (!user) {
      return res.status(404).json({ error: "User not found" });
    }

    // Get alert
    const { data: alert, error: alertError } = await supabase
      .from("alerts")
      .select("*, sensors(*)")
      .eq("id", alert_id)
      .single();

    if (alertError || !alert) {
      return res.status(404).json({ error: "Alert not found" });
    }

    const now = new Date().toISOString();

    if (action === "acknowledge") {
      if (alert.status !== "open") {
        return res.status(400).json({ error: "Alert is not open", current_status: alert.status });
      }

      // Update alert
      await supabase
        .from("alerts")
        .update({
          status: "acknowledged",
          acknowledged_by: user.id,
          acknowledged_at: now,
        })
        .eq("id", alert_id);

      // Log action
      await supabase.from("alert_actions").insert({
        alert_id,
        user_id: user.id,
        action: "acknowledged",
        note: `${user.display_name} is heading to site for ${sensor_id || alert.sensor_id}`,
      });

      // Notify others via n8n
      await notifyN8n({
        type: "acknowledge_broadcast",
        alert_id,
        sensor_id: alert.sensor_id,
        sensor_name: alert.sensors?.name,
        sensor_type: alert.sensors?.type,
        unit: alert.sensors?.unit,
        responder_name: user.display_name,
        responder_line_user_id: line_user_id,
        ack_time: now,
      });

      return res.status(200).json({
        success: true,
        action: "acknowledged",
        responder: user.display_name,
      });

    } else if (action === "resolve") {
      if (alert.status !== "acknowledged") {
        return res.status(400).json({ error: "Alert is not acknowledged", current_status: alert.status });
      }

      // Calculate response time
      const ackTime = new Date(alert.acknowledged_at || alert.created_at);
      const resolveTime = new Date(now);
      const responseTimeMin = Math.round((resolveTime - ackTime) / 60000);

      // Get current sensor value
      const { data: latestReading } = await supabase
        .from("sensor_readings")
        .select("value")
        .eq("sensor_id", alert.sensor_id)
        .order("recorded_at", { ascending: false })
        .limit(1)
        .single();

      // Update alert
      await supabase
        .from("alerts")
        .update({
          status: "resolved",
          resolved_by: user.id,
          resolved_at: now,
          response_time_min: responseTimeMin,
        })
        .eq("id", alert_id);

      // Log action
      await supabase.from("alert_actions").insert({
        alert_id,
        user_id: user.id,
        action: "resolved",
        note: `Resolved by ${user.display_name} in ${responseTimeMin} min`,
      });

      // Notify all via n8n
      await notifyN8n({
        type: "resolved_broadcast",
        alert_id,
        sensor_id: alert.sensor_id,
        sensor_name: alert.sensors?.name,
        sensor_type: alert.sensors?.type,
        unit: alert.sensors?.unit,
        responder_name: user.display_name,
        responder_line_user_id: line_user_id,
        duration: responseTimeMin,
        current_value: latestReading?.value,
        resolve_time: now,
      });

      return res.status(200).json({
        success: true,
        action: "resolved",
        responder: user.display_name,
        response_time_min: responseTimeMin,
      });

    } else {
      return res.status(400).json({ error: `Unknown action: ${action}` });
    }
  } catch (err) {
    console.error("Alert action error:", err);
    return res.status(500).json({ error: err.message });
  }
};
