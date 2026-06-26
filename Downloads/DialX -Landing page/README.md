# DialX — Pre-Launch Landing Page

A premium, conversion-focused pre-launch landing page for **DialX**, India's emergency healthcare coordination platform.

Built with Next.js 15, React 19, TypeScript, Tailwind CSS, Shadcn/UI, Framer Motion, Supabase, and Resend.

## Features

- **15-section landing page** — Hero, trust indicators, problem/solution, features, roadmap, waitlist, pharmacy partners, FAQ, and more
- **Waitlist & pharmacy forms** — Validated with Zod, stored in Supabase, confirmation emails via Resend
- **SEO optimized** — Structured data, Open Graph, Twitter Cards, sitemap, robots.txt
- **Analytics** — Google Analytics 4, Google Tag Manager, Microsoft Clarity, scroll depth tracking
- **Accessible** — Semantic HTML, keyboard navigation, ARIA labels, focus states
- **Performance** — Optimized for Core Web Vitals, mobile-first responsive design

## Quick Start

### Prerequisites

- Node.js 18.17+
- npm or yarn
- Supabase account (optional for local dev)
- Resend account (optional for local dev)

### Installation

```bash
# Clone and install
npm install

# Copy environment variables
cp .env.example .env.local

# Start development server
npm run dev
```

Open [http://localhost:3000](http://localhost:3000).

Forms work in **dev mode** without Supabase/Resend — submissions are logged to the console.

## Environment Variables

| Variable | Description |
|----------|-------------|
| `NEXT_PUBLIC_SITE_URL` | Production site URL (e.g. `https://dialx.in`) |
| `NEXT_PUBLIC_SUPABASE_URL` | Supabase project URL |
| `NEXT_PUBLIC_SUPABASE_ANON_KEY` | Supabase anon key |
| `SUPABASE_SERVICE_ROLE_KEY` | Supabase service role key (server only) |
| `RESEND_API_KEY` | Resend API key for emails |
| `RESEND_FROM_EMAIL` | Sender email (e.g. `DialX <noreply@dialx.in>`) |
| `NEXT_PUBLIC_GA_MEASUREMENT_ID` | Google Analytics 4 ID |
| `NEXT_PUBLIC_GTM_ID` | Google Tag Manager container ID |
| `NEXT_PUBLIC_CLARITY_ID` | Microsoft Clarity project ID |
| `NEXT_PUBLIC_WHATSAPP_COMMUNITY_URL` | WhatsApp Community/Group invite link |
| `NEXT_PUBLIC_WHATSAPP_NUMBER` | Fallback WhatsApp number (country code, no +) |

## WhatsApp Community Setup

1. Create a WhatsApp Community or Group for early supporters
2. Copy the invite link (`https://chat.whatsapp.com/...`)
3. Add it to `NEXT_PUBLIC_WHATSAPP_COMMUNITY_URL` in `.env.local`
4. Optionally set `NEXT_PUBLIC_WHATSAPP_NUMBER` as fallback for direct messages

Users can join via the hero CTA, waitlist alternative card, floating button, or navbar — no form required.

## Supabase Setup

1. Create a new Supabase project at [supabase.com](https://supabase.com)
2. Run the SQL in `supabase/schema.sql` in the SQL Editor
3. Copy your project URL and keys to `.env.local`

## Resend Setup

1. Create an account at [resend.com](https://resend.com)
2. Verify your sending domain
3. Add your API key and from email to `.env.local`

## Deploy to Vercel

[![Deploy with Vercel](https://vercel.com/button)](https://vercel.com/new)

1. Push this repo to GitHub
2. Import the project in Vercel
3. Add all environment variables from `.env.example`
4. Deploy

Vercel will auto-detect Next.js and configure the build.

## Project Structure

```
src/
├── app/                    # Next.js App Router pages & API routes
│   ├── api/waitlist/       # Waitlist form endpoint
│   ├── api/pharmacy/       # Pharmacy partner endpoint
│   ├── contact/            # Contact page
│   ├── privacy/            # Privacy policy
│   ├── terms/              # Terms & conditions
│   └── success/waitlist/   # Success page
├── components/
│   ├── analytics/          # GA4, GTM, Clarity, scroll tracking
│   ├── layout/             # Navbar, Footer
│   ├── sections/           # Landing page sections
│   ├── shared/             # Reusable components
│   └── ui/                 # Shadcn/UI components
└── lib/                    # Utils, constants, validations, integrations
```

## Scripts

| Command | Description |
|---------|-------------|
| `npm run dev` | Start development server |
| `npm run build` | Production build |
| `npm run start` | Start production server |
| `npm run lint` | Run ESLint |

## License

Proprietary — DialX. All rights reserved.
