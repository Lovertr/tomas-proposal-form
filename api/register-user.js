// Vercel Serverless Function: /api/register-user
// POST endpoint for LIFF registration form
//
// POST /api/register-user
// Body: { line_user_id, first_name, last_name, department }

const { createClient } = require("@supabase/supabase-js");

const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_SERVICE_KEY
);

module.exports = async function handler(req, res) {
  // CORS
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  try {
    const { line_user_id, first_name, last_name, department } = req.body;

    if (!line_user_id || !first_name || !last_name) {
      return res.status(400).json({ error: "Missing required fields: line_user_id, first_name, last_name" });
    }

    const display_name = `${first_name} ${last_name}`;

    const { data, error } = await supabase
      .from("users")
      .upsert(
        {
          line_user_id,
          first_name,
          last_name,
          display_name,
          department: department || null,
        },
        { onConflict: "line_user_id" }
      )
      .select("*")
      .single();

    if (error) {
      console.error("Register user error:", error);
      return res.status(500).json({ error: error.message });
    }

    return res.status(200).json({
      success: true,
      message: "User registered successfully",
      user: data,
    });
  } catch (err) {
    console.error("Register user error:", err);
    return res.status(500).json({ error: err.message });
  }
};
