-- DialX Landing Page — Supabase Schema
-- Run this in your Supabase SQL Editor

-- Waitlist table
CREATE TABLE IF NOT EXISTS waitlist (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  name TEXT NOT NULL,
  email TEXT NOT NULL UNIQUE,
  phone TEXT NOT NULL,
  city TEXT NOT NULL,
  user_type TEXT NOT NULL,
  biggest_challenge TEXT NOT NULL,
  consent BOOLEAN NOT NULL DEFAULT false,
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- Pharmacy partners table
CREATE TABLE IF NOT EXISTS pharmacy_partners (
  id UUID DEFAULT gen_random_uuid() PRIMARY KEY,
  pharmacy_name TEXT NOT NULL,
  owner_name TEXT NOT NULL,
  gstin TEXT,
  city TEXT NOT NULL,
  contact_number TEXT NOT NULL,
  email TEXT NOT NULL UNIQUE,
  daily_order_estimate TEXT NOT NULL,
  status TEXT DEFAULT 'pending',
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- Enable Row Level Security
ALTER TABLE waitlist ENABLE ROW LEVEL SECURITY;
ALTER TABLE pharmacy_partners ENABLE ROW LEVEL SECURITY;

-- Service role bypasses RLS; no public policies needed for server-side inserts

-- Indexes for performance
CREATE INDEX IF NOT EXISTS idx_waitlist_email ON waitlist(email);
CREATE INDEX IF NOT EXISTS idx_waitlist_created_at ON waitlist(created_at DESC);
CREATE INDEX IF NOT EXISTS idx_pharmacy_email ON pharmacy_partners(email);
CREATE INDEX IF NOT EXISTS idx_pharmacy_city ON pharmacy_partners(city);
