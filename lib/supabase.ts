import { createClient } from "@supabase/supabase-js";

export const supabase = createClient(
  "https://fcjtntvuuhmrqgafdsjl.supabase.co",
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZjanRudHZ1dWhtcnFnYWZkc2psIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzQ1MTI2OTYsImV4cCI6MjA5MDA4ODY5Nn0.dx8b_rkv7Lt-9K-xGq9-z9OnLsolFNnWJfoTTA8re7M"
);
