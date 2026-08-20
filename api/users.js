// Vercel Serverless Function: /api/users
// User management API for Dashboard
//
// GET  /api/users          → List all users with sensor permissions
// POST /api/users          → Update user sensor permissions { user_id, sensors: [...] }
// PUT  /api/users          → Update user info { user_id, first_name, last_name, department, role, is_active }

const { createClient } = require("@supabase/supabase-js");

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_SERVICE_KEY
);

module.exports = async function handler(req, res) {
  // CORS
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, POST, PUT, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  if (req.method === "OPTIONS") return res.status(200).end();

  try {
    // ─── GET: List all users with sensor permissions ──────
    if (req.method === "GET") {
      // Get all users
      const { data: usersData, error: usersError } = await supabase
        .from("users")
        .select("id, line_user_id, display_name, first_name, last_name, department, role, is_active, created_at")
        .order("created_at", { ascending: true });

      if (usersError) throw usersError;

      // Get all sensor permissions
      const { data: permsData, error: permsError } = await supabase
        .from("user_sensor_permissions")
        .select("user_id, sensor_id");

      if (permsError) throw permsError;

      // Group sensor_ids by user_id
      const permsByUser = {};
      for (const perm of permsData || []) {
        if (!permsByUser[perm.user_id]) permsByUser[perm.user_id] = [];
        permsByUser[perm.user_id].push(perm.sensor_id);
      }

      // Merge
      const users = (usersData || []).map((u) => ({
        ...u,
        sensors: permsByUser[u.id] || [],
      }));

      return res.status(200).json({ users });
    }

    // ─── POST: Update user sensor permissions ─────────────
    if (req.method === "POST") {
      const { user_id, sensors } = req.body;

      if (!user_id || !Array.isArray(sensors)) {
        return res.status(400).json({ error: "Missing user_id or sensors array" });
      }

      // Delete existing permissions
      const { error: deleteError } = await supabase
        .from("user_sensor_permissions")
        .delete()
        .eq("user_id", user_id);

      if (deleteError) throw deleteError;

      // Insert new permissions
      if (sensors.length > 0) {
        const rows = sensors.map((sensor_id) => ({ user_id, sensor_id }));
        const { error: insertError } = await supabase
          .from("user_sensor_permissions")
          .insert(rows);

        if (insertError) throw insertError;
      }

      return res.status(200).json({
        success: true,
        message: "Permissions updated",
        user_id,
        sensors,
      });
    }

    // ─── PUT: Update user info ────────────────────────────
    if (req.method === "PUT") {
      const { user_id, first_name, last_name, department, role, is_active } = req.body;

      if (!user_id) {
        return res.status(400).json({ error: "Missing user_id" });
      }

      const updates = {};
      if (first_name !== undefined) updates.first_name = first_name;
      if (last_name !== undefined) updates.last_name = last_name;
      if (department !== undefined) updates.department = department;
      if (role !== undefined) updates.role = role;
      if (is_active !== undefined) updates.is_active = is_active;

      // Update display_name if name fields changed
      if (first_name !== undefined || last_name !== undefined) {
        // Need to fetch current values for the unchanged field
        if (first_name === undefined || last_name === undefined) {
          const { data: current } = await supabase
            .from("users")
            .select("first_name, last_name")
            .eq("id", user_id)
            .single();
          if (current) {
            const fn = first_name !== undefined ? first_name : current.first_name;
            const ln = last_name !== undefined ? last_name : current.last_name;
            updates.display_name = `${fn} ${ln}`;
          }
        } else {
          updates.display_name = `${first_name} ${last_name}`;
        }
      }

      if (Object.keys(updates).length === 0) {
        return res.status(400).json({ error: "No fields to update" });
      }

      const { data, error } = await supabase
        .from("users")
        .update(updates)
        .eq("id", user_id)
        .select("*")
        .single();

      if (error) throw error;

      return res.status(200).json({
        success: true,
        message: "User updated",
        user: data,
      });
    }

    return res.status(405).json({ error: "Method not allowed" });
  } catch (err) {
    console.error("Users API error:", err);
    return res.status(500).json({ error: err.message });
  }
};
