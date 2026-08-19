// Vercel Serverless Function: /api/sensor-ingest
// Receives sensor data, stores in Supabase, checks thresholds, triggers n8n alert workflow
//
// POST /api/sensor-ingest
// Body: { sensor_id: "PS-01", value: 9.2 }
// or batch: { readings: [{ sensor_id: "PS-01", value: 9.2 }, { sensor_id: "TS-01", value: 88 }] }

const { createClient } = require("@supabase/supabase-js");

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_SERVICE_KEY
);

// n8n webhook URL — only called when anomaly detected (low volume)
const N8N_ALERT_WEBHOOK = process.env.N8N_ALERT_WEBHOOK_URL;

// Sensor thresholds cache (loaded once per cold start)
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

async function triggerN8nAlert(alertData) {
  if (!N8N_ALERT_WEBHOOK) {
    console.warn("N8N_ALERT_WEBHOOK_URL not configured, skipping notification");
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
    // Don't throw — data is already saved, alert is logged
  }
}

module.exports = async function handler(req, res) {
  // CORS
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  // Auth check (simple API key)
  const apiKey = req.headers["x-api-key"];
  if (process.env.INGEST_API_KEY && apiKey !== process.env.INGEST_API_KEY) {
    return res.status(401).json({ error: "Unauthorized" });
  }

  try {
    const sensors = await getSensors();
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

      // 2. If anomaly → check if there's already an open alert for this sensor
      if (anomaly.is_anomaly) {
        const { data: existingAlert } = await supabase
          .from("alerts")
          .select("id")
          .eq("sensor_id", sensor_id)
          .in("status", ["open", "acknowledged"])
          .order("created_at", { ascending: false })
          .limit(1)
          .single();

        if (!existingAlert) {
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

            // Trigger n8n (low volume — only on NEW alerts)
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

            result.alert_id = alertData.id;
            result.alert_status = "new";
          }
        } else {
          result.alert_id = existingAlert.id;
          result.alert_status = "existing";
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
