-- Migration 024: surgical_photos and surgical_drawings
-- Supports Revela photo upload + Fabric.js / Servier Medical Art drawing tool

-- ── cr.surgical_photos ───────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS cr.surgical_photos (
  id                  uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  org_id              uuid NOT NULL,
  encounter_id        uuid NOT NULL REFERENCES cr.encounter(encounter_id) ON DELETE CASCADE,
  patient_id          uuid NOT NULL,
  storage_path        text NOT NULL,
  storage_bucket      text NOT NULL DEFAULT 'patient-photos',
  body_area           text NOT NULL,
  -- e.g. breast_left, breast_right, abdomen, face, pectoral, full_body, lipo_zone
  photo_type          text NOT NULL CHECK (photo_type IN ('before', 'after', 'intraop', 'reference')),
  view_angle          text,
  -- e.g. front, lateral_left, lateral_right, oblique, back
  notes               text,
  taken_at            timestamptz NOT NULL DEFAULT now(),
  uploaded_by         uuid REFERENCES auth.users(id),
  file_name           text,
  file_size_bytes     integer,
  mime_type           text,
  created_at          timestamptz NOT NULL DEFAULT now(),
  updated_at          timestamptz NOT NULL DEFAULT now()
);

-- ── cr.surgical_drawings ─────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS cr.surgical_drawings (
  id                  uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  org_id              uuid NOT NULL,
  encounter_id        uuid NOT NULL REFERENCES cr.encounter(encounter_id) ON DELETE CASCADE,
  patient_id          uuid NOT NULL,
  template_key        text NOT NULL,
  -- format: "breast:Front", "body_female:Back", "pectoral:Front", etc.
  template_source     text NOT NULL DEFAULT 'servier',
  -- 'servier' = Servier Medical Art CC BY 4.0 | 'custom' = PatientTrac custom SVG
  drawing_svg         text,
  -- Full SVG export of canvas (base template + annotations)
  drawing_json        text,
  -- Fabric.js JSON (editable state — allows re-opening and editing)
  notes               text,
  created_by          uuid REFERENCES auth.users(id),
  created_at          timestamptz NOT NULL DEFAULT now(),
  updated_at          timestamptz NOT NULL DEFAULT now(),
  UNIQUE (encounter_id, template_key)
  -- One drawing per encounter per template+view combination
);

-- ── Indexes ──────────────────────────────────────────────────────────────────
CREATE INDEX IF NOT EXISTS idx_surgical_photos_encounter    ON cr.surgical_photos(encounter_id);
CREATE INDEX IF NOT EXISTS idx_surgical_photos_patient      ON cr.surgical_photos(patient_id);
CREATE INDEX IF NOT EXISTS idx_surgical_photos_org          ON cr.surgical_photos(org_id);
CREATE INDEX IF NOT EXISTS idx_surgical_photos_type         ON cr.surgical_photos(photo_type);
CREATE INDEX IF NOT EXISTS idx_surgical_drawings_encounter  ON cr.surgical_drawings(encounter_id);
CREATE INDEX IF NOT EXISTS idx_surgical_drawings_patient    ON cr.surgical_drawings(patient_id);
CREATE INDEX IF NOT EXISTS idx_surgical_drawings_org        ON cr.surgical_drawings(org_id);

-- ── Row Level Security ────────────────────────────────────────────────────────
ALTER TABLE cr.surgical_photos   ENABLE ROW LEVEL SECURITY;
ALTER TABLE cr.surgical_drawings ENABLE ROW LEVEL SECURITY;

-- Photos: org isolation
CREATE POLICY "surgical_photos_org_isolation" ON cr.surgical_photos
  USING (org_id = (
    SELECT org_id FROM saas.org_members WHERE user_id = auth.uid() LIMIT 1
  ));

-- Drawings: org isolation
CREATE POLICY "surgical_drawings_org_isolation" ON cr.surgical_drawings
  USING (org_id = (
    SELECT org_id FROM saas.org_members WHERE user_id = auth.uid() LIMIT 1
  ));

-- ── Supabase Storage buckets (run via Supabase dashboard or CLI) ──────────────
-- Bucket 1: patient-photos  (PRIVATE — RLS enforced via signed URLs)
-- Bucket 2: revela-templates (PUBLIC READ — Servier Medical Art + custom SVGs)
--
-- After applying this migration, create the buckets in Supabase Storage:
--   supabase storage create-bucket patient-photos --private
--   supabase storage create-bucket revela-templates --public
--
-- Upload Servier Medical Art SVGs to revela-templates/:
--   female-breast-front.svg, female-breast-lateral.svg, female-breast-oblique.svg
--   female-body-front.svg, female-body-back.svg, female-body-lateral.svg
--   female-face-front.svg, female-face-lateral.svg
--   male-pectoral-front.svg
--   male-body-front.svg, male-body-back.svg, male-body-lateral.svg
--   male-face-front.svg, male-face-lateral.svg
--
-- Attribution required in UI: "Illustrations by Servier Medical Art
--   (smart.servier.com), CC BY 4.0"
