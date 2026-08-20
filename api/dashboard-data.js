// Vercel Serverless Function: /api/dashboard-data
// Provides real-time data for the monitoring dashboard
//
// GET /api/dashboard-data
// Query params:
//   ?type=latest           → latest reading per sensor
//   ?type=alerts           → open/acknowledged alerts
//   ?type=history          → recent readings for charts (default: last 1 hour)
//   ?type=summary          → daily stats
//   ?type=alert-history    → all alerts with user names
//   ?type=resolve-all      → resolve all open/acknowledged alerts
//   ?type=sensor-detail    → alerts for a specific sensor with user info
//   ?type=clear-test-data  → delete all test data
//   ?type=users-summary    → quick user count
//   ?sensor_id=PS-01       → filter by sensor (optional)

const { createClient } = require("@supabase/supabase-js");

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_SERVICE_KEY
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
        const { data, error } = await supabase.from("latest_readings").select("*");
        if (error) throw error;

        // Also get sensor config
        const { data: sensors } = await supabase.from("sensors").select("*").eq("is_active", true);

        return res.status(200).json({
          readings: data,
          sensors,
        });
      }

      case "alerts": {
        const { data, error } = await supabase.from("open_alerts").select("*");
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

        // Active sensor count
        const { count: activeSensors } = await supabase
          .from("sensors")
          .select("*", { count: "exact", head: true })
          .eq("is_active", true);

        return res.status(200).json({
          total_alerts_today: totalToday || 0,
          resolved_today: resolvedToday || 0,
          open_alerts: openAlerts || 0,
          avg_response_min: avgResponse,
          active_sensors: activeSensors || 0,
        });
      }

      case "alert-history": {
        // Join users table for acknowledger and resolver names
        const { data, error } = await supabase
          .from("alerts")
          .select(`
            *,
            sensors(name, type, unit),
            acknowledger:users!alerts_acknowledged_by_fkey(display_name),
            resolver:users!alerts_resolved_by_fkey(display_name)
          `)
          .order("created_at", { ascending: false })
          .limit(100);

        if (error) throw error;

        // Flatten user names
        const alerts = (data || []).map((a) => ({
          ...a,
          acknowledged_by_name: a.acknowledger?.display_name || null,
          resolved_by_name: a.resolver?.display_name || null,
          acknowledger: undefined,
          resolver: undefined,
        }));

        return res.status(200).json({ alerts });
      }

      case "resolve-all": {
        // Resolve all open/acknowledged alerts (used by dashboard reset)
        const { data, error } = await supabase
          .from("alerts")
          .update({ status: "resolved", resolved_at: new Date().toISOString() })
          .in("status", ["open", "acknowledged"])
          .select("id");
        if (error) throw error;
        return res.status(200).json({ resolved: data?.length || 0 });
      }

      case "sensor-detail": {
        if (!sensor_id) {
          return res.status(400).json({ error: "sensor_id is required for sensor-detail" });
        }

        // Get sensor info
        const { data: sensorData, error: sensorError } = await supabase
          .from("sensors")
          .select("*")
          .eq("id", sensor_id)
          .single();

        if (sensorError) throw sensorError;

        // Get all alerts for this sensor with user info
        const { data: alertsData, error: alertsError } = await supabase
          .from("alerts")
          .select(`
            *,
            acknowledger:users!alerts_acknowledged_by_fkey(id, display_name, department),
            resolver:users!alerts_resolved_by_fkey(id, display_name, department)
          `)
          .eq("sensor_id", sensor_id)
          .order("created_at", { ascending: false })
          .limit(50);

        if (alertsError) throw alertsError;

        // Flatten
        const alerts = (alertsData || []).map((a) => ({
          ...a,
          acknowledged_by_name: a.acknowledger?.display_name || null,
          acknowledged_by_department: a.acknowledger?.department || null,
          resolved_by_name: a.resolver?.display_name || null,
          resolved_by_department: a.resolver?.department || null,
          acknowledger: undefined,
          resolver: undefined,
        }));

        // Get latest readings
        const { data: readings } = await supabase
          .from("sensor_readings")
          .select("value, is_anomaly, recorded_at")
          .eq("sensor_id", sensor_id)
          .order("recorded_at", { ascending: false })
          .limit(100);

        return res.status(200).json({
          sensor: sensorData,
          alerts,
          recent_readings: readings || [],
        });
      }

      case "clear-test-data": {
        // Delete all test data (alert_actions first due to FK, then alerts, then readings)
        const { count: actionsCount } = await supabase
          .from("alert_actions")
          .delete()
          .neq("id", 0) // match all rows
          .select("*", { count: "exact", head: true });

        const { count: alertsCount } = await supabase
          .from("alerts")
          .delete()
          .neq("id", "00000000-0000-0000-0000-000000000000") // match all rows
          .select("*", { count: "exact", head: true });

        const { count: readingsCount } = await supabase
          .from("sensor_readings")
          .delete()
          .neq("id", 0) // match all rows
          .select("*", { count: "exact", head: true });

        return res.status(200).json({
          success: true,
          deleted: {
            alert_actions: actionsCount || 0,
            alerts: alertsCount || 0,
            sensor_readings: readingsCount || 0,
          },
        });
      }

      case "users-summary": {
        const { count: totalUsers } = await supabase
          .from("users")
          .select("*", { count: "exact", head: true });

        const { count: activeUsers } = await supabase
          .from("users")
          .select("*", { count: "exact", head: true })
          .eq("is_active", true);

        return res.status(200).json({
          total_users: totalUsers || 0,
          active_users: activeUsers || 0,
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
