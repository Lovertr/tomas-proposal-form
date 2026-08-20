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
//   ?type=sensors-config   → all sensors with full config (GET)
//   ?type=system-config    → system settings (GET)
//   ?sensor_id=PS-01       → filter by sensor (optional)
//
// POST /api/dashboard-data
// Body: { type, ...data }
//   type=sensor-create     → create new sensor
//   type=sensor-update     → update sensor by id
//   type=sensor-delete     → delete sensor by id
//   type=system-config     → update system config
//   type=sensors-reorder   → update sort order

const { createClient } = require("@supabase/supabase-js");

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_SERVICE_KEY
);

module.exports = async function handler(req, res) {
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();

  try {
    // --- GET handlers ---
    if (req.method === "GET") {
      const { type = "latest", sensor_id, hours = "1" } = req.query;

      switch (type) {
        case "latest": {
          const { data, error } = await supabase.from("latest_readings").select("*");
          if (error) throw error;
          const { data: sensors } = await supabase
            .from("sensors")
            .select("*")
            .eq("is_active", true)
            .order("sort_order", { ascending: true });
          return res.status(200).json({ readings: data, sensors });
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
          const { data: avgData } = await supabase
            .from("alerts")
            .select("response_time_min")
            .gte("created_at", today)
            .not("response_time_min", "is", null);
          const avgResponse = avgData?.length
            ? Math.round(avgData.reduce((s, r) => s + r.response_time_min, 0) / avgData.length)
            : 0;
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
          const { data: sensorData, error: sensorError } = await supabase
            .from("sensors")
            .select("*")
            .eq("id", sensor_id)
            .single();
          if (sensorError) throw sensorError;
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
          const alerts = (alertsData || []).map((a) => ({
            ...a,
            acknowledged_by_name: a.acknowledger?.display_name || null,
            acknowledged_by_department: a.acknowledger?.department || null,
            resolved_by_name: a.resolver?.display_name || null,
            resolved_by_department: a.resolver?.department || null,
            acknowledger: undefined,
            resolver: undefined,
          }));
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
          const { count: actionsCount } = await supabase
            .from("alert_actions")
            .delete()
            .neq("id", 0)
            .select("*", { count: "exact", head: true });
          const { count: alertsCount } = await supabase
            .from("alerts")
            .delete()
            .neq("id", "00000000-0000-0000-0000-000000000000")
            .select("*", { count: "exact", head: true });
          const { count: readingsCount } = await supabase
            .from("sensor_readings")
            .delete()
            .neq("id", 0)
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

        // --- NEW: Get all sensors with full config ---
        case "sensors-config": {
          const { data, error } = await supabase
            .from("sensors")
            .select("*")
            .order("sort_order", { ascending: true });
          if (error) throw error;
          return res.status(200).json({ sensors: data });
        }

        // --- NEW: Get system config ---
        case "system-config": {
          const { data, error } = await supabase
            .from("system_config")
            .select("*");
          if (error) throw error;
          // Convert array to key-value object
          const config = {};
          (data || []).forEach((row) => {
            config[row.key] = row.value;
          });
          return res.status(200).json({ config });
        }

        default:
          return res.status(400).json({ error: `Unknown type: ${type}` });
      }
    }

    // --- POST handlers ---
    if (req.method === "POST") {
      const body = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
      const { type } = body;

      switch (type) {
        // --- Create new sensor ---
        case "sensor-create": {
          const { id, name, name_th, sensor_type, unit, signal_type, connection_protocol,
                  connection_config, polling_interval_sec, location,
                  min_value, max_value, normal_min, normal_max,
                  threshold_low, threshold_high, nominal, sort_order, card_color } = body;

          if (!id || !name || !sensor_type || !unit) {
            return res.status(400).json({ error: "id, name, sensor_type, unit are required" });
          }

          const { data, error } = await supabase.from("sensors").insert({
            id,
            name,
            name_th: name_th || null,
            type: sensor_type,
            unit,
            signal_type: signal_type || "4-20mA",
            connection_protocol: connection_protocol || "simulator",
            connection_config: connection_config || {},
            polling_interval_sec: polling_interval_sec || 30,
            location: location || "",
            min_value: min_value ?? 0,
            max_value: max_value ?? 100,
            normal_min: normal_min ?? 0,
            normal_max: normal_max ?? 100,
            threshold_low: threshold_low ?? null,
            threshold_high: threshold_high ?? null,
            nominal: nominal ?? 50,
            sort_order: sort_order ?? 99,
            card_color: card_color || null,
            is_active: true,
          }).select().single();

          if (error) throw error;

          // Also add permission for all existing users
          const { data: users } = await supabase.from("users").select("id, assigned_sensors");
          if (users) {
            for (const user of users) {
              const sensors = user.assigned_sensors || [];
              if (!sensors.includes(id)) {
                sensors.push(id);
                await supabase.from("users").update({ assigned_sensors: sensors }).eq("id", user.id);
              }
            }
            // Add to user_sensor_permissions
            for (const user of users) {
              await supabase.from("user_sensor_permissions").insert({
                user_id: user.id,
                sensor_id: id,
                can_acknowledge: true,
                can_resolve: true,
              }).onConflict("user_id,sensor_id").ignore();
            }
          }

          return res.status(200).json({ success: true, sensor: data });
        }

        // --- Update sensor ---
        case "sensor-update": {
          const { sensor_id: sid, ...updates } = body;
          if (!sid) return res.status(400).json({ error: "sensor_id is required" });

          // Map sensor_type to type for DB
          const dbUpdates = {};
          const allowedFields = [
            "name", "name_th", "unit", "signal_type", "connection_protocol",
            "connection_config", "polling_interval_sec", "location",
            "min_value", "max_value", "normal_min", "normal_max",
            "threshold_low", "threshold_high", "nominal", "sort_order", "is_active",
            "card_color",
          ];
          for (const key of allowedFields) {
            if (updates[key] !== undefined) dbUpdates[key] = updates[key];
          }
          if (updates.sensor_type) dbUpdates.type = updates.sensor_type;

          const { data, error } = await supabase
            .from("sensors")
            .update(dbUpdates)
            .eq("id", sid)
            .select()
            .single();

          if (error) throw error;
          return res.status(200).json({ success: true, sensor: data });
        }

        // --- Delete sensor ---
        case "sensor-delete": {
          const { sensor_id: delId } = body;
          if (!delId) return res.status(400).json({ error: "sensor_id is required" });

          // Soft delete — set is_active = false
          const { error } = await supabase
            .from("sensors")
            .update({ is_active: false })
            .eq("id", delId);

          if (error) throw error;
          return res.status(200).json({ success: true, deleted: delId });
        }

        // --- Update system config ---
        case "system-config": {
          const { config } = body;
          if (!config || typeof config !== "object") {
            return res.status(400).json({ error: "config object is required" });
          }

          for (const [key, value] of Object.entries(config)) {
            await supabase
              .from("system_config")
              .upsert({
                key,
                value: typeof value === "string" ? JSON.stringify(value) : value,
                updated_at: new Date().toISOString(),
              });
          }

          return res.status(200).json({ success: true });
        }

        // --- Reorder sensors ---
        case "sensors-reorder": {
          const { order } = body; // [{id: "PT-01", sort_order: 1}, ...]
          if (!Array.isArray(order)) {
            return res.status(400).json({ error: "order array is required" });
          }

          for (const item of order) {
            await supabase
              .from("sensors")
              .update({ sort_order: item.sort_order })
              .eq("id", item.id);
          }

          return res.status(200).json({ success: true });
        }

        default:
          return res.status(400).json({ error: `Unknown POST type: ${type}` });
      }
    }

    return res.status(405).json({ error: "Method not allowed" });
  } catch (err) {
    console.error("Dashboard data error:", err);
    return res.status(500).json({ error: err.message });
  }
};
