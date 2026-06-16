import { createClient } from "@supabase/supabase-js";

const supabaseUrl = "https://zedanvoikhrxlhnghyui.supabase.co";

const supabaseKey =
"eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InplZGFudm9pa2hyeGxobmdoeXVpIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODE1MjIyMTAsImV4cCI6MjA5NzA5ODIxMH0.xqWJ_f3ZBkF7fSa51wSu2ddIpb9x843aOcc3I9CH0fE";

export const supabase = createClient(
  supabaseUrl,
  supabaseKey
);