import { createClient } from "@supabase/supabase-js";
import { requireSupabaseStorageConfiguration } from "./config.js";

let storageClient: ReturnType<typeof createClient> | null = null;

export const supabaseStorage = () => {
  requireSupabaseStorageConfiguration();
  if (!storageClient) {
    storageClient = createClient(
      (process.env.SUPABASE_URL || process.env.VITE_SUPABASE_URL) as string,
      (process.env.SUPABASE_SECRET_KEY || process.env.SUPABASE_SERVICE_ROLE_KEY) as string,
      { auth: { persistSession: false, autoRefreshToken: false, detectSessionInUrl: false } },
    );
  }
  return storageClient.storage.from("hr-document-assets");
};
