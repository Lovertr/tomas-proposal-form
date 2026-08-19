// Vercel Serverless Function: /api/dashboard-data
// Provides real-time data for the monitoring dashboard
//
// GET /api/dashboard-data
// Query params:
//   ?type=latest      → latest reading per sensor
//   ?type=alerts      → open/acknowledged alerts
//   ?type=history     → recent readings for charts (default: last 1 hour)
//   ?type=summary     → daily stats
//   ?sensor_id=PS-01  → filter by sensor (optional)

const { createClient } = require("@supabase/supabase-js");

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY
);

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "GET") return res.status(405).json({ error: "Method not allowed" });

  try {
    const { type = "latest", sensor_id, hours = "1" } = req.query;

    switch (type) {
      case "latest": {
        const { data, error } = await supabase
          .from("latest_readings")
          .select("*");
        if (error) throw error;

        // Also get sensor config
        const { data: sensors } = await supabase.from("sensors").select("*").eq("is_active", true);

        return res.status(200).json({
          readings: data,
          sensors,
        });
      }

      case "alerts": {
        const { data, error } = await supabase
          .from("open_alerts")
          .select("*");
        if (error) throw error;
        return res.status(200).json({ alerts: data });
      }

      case "history": {
        const since = new Date(Date.now() - parseInt(hours) * 3600000).toISOString();
        let query = supabase
          .from("sensor_readings")
          .select("sensor_id, value, is_anomaly, recorded_at")
          .gte("recorded_at", since)
          .order("recorded_at", { ascending: true });

        if (sensor_id) query = query.eq("sensor_id", sensor_id);

        const { data, error } = await query.limit(1000);
        if (error) throw error;
        return res.status(200).json({ history: data });
      }

      case "summary": {
        const today = new Date().toISOString().split("T")[0];

        // Today's alerts count
        const { count: totalToday } = await supabase
          .from("alerts")
          .select("*", { count: "exact", head: true })
          .gte("created_at", today);

        const { count: resolvedToday } = await supabase
          .from("alerts")
          .select("*", { count: "exact", head: true })
          .gte("created_at", today)
          .eq("status", "resolved");

        const { count: openAlerts } = await supabase
          .from("alerts")
          .select("*", { count: "exact", head: true })
          .in("status", ["open", "acknowledged"]);

        // Avg response time today
        const { data: avgData } = await supabase
          .from("alerts")
          .select("response_time_min")
          .gte("created_at", today)
          .not("response_time_min", "is", null);

        const avgResponse = avgData?.length
          ? Math.round(avgData.reduce((s, r) => s + r.response_time_min, 0) / avgData.length)
          : 0;

        return res.status(200).json({
          total_alerts_today: totalToday || 0,
          resolved_today: resolvedToday || 0,
          open_alerts: openAlerts || 0,
          avg_response_min: avgResponse,
          active_sensors: 6,
        });
      }

      default:
        return res.status(400).json({ error: `Unknown type: ${type}` });
    }
  } catch (err) {
    console.error("Dashboard data error:", err);
    return res.status(500).json({ error: err.message });
  }
};
