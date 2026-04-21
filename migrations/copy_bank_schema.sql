-- Copy Bank: shared outreach copy storage
-- Run once in the Supabase SQL editor

CREATE TABLE IF NOT EXISTS copy_bank_templates (
    key     TEXT PRIMARY KEY,   -- e.g. "US_PR", "UK_Marketing"
    content JSONB NOT NULL DEFAULT '{
        "email":    [{"subject":"","body":""}],
        "flyout":   [{"body":""}],
        "linkedin": [{"body":""}]
    }'::jsonb
    -- content.email    = array of {subject, body}  — multiple variations
    -- content.flyout   = array of {body}            — multiple variations
    -- content.linkedin = array of {body}            — ordered sequence steps
);

ALTER TABLE copy_bank_templates DISABLE ROW LEVEL SECURITY;
