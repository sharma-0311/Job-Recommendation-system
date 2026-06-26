export const siteConfig = {
  name: "DialX",
  tagline: "Emergency Healthcare Coordination",
  description:
    "DialX is India's emergency healthcare coordination platform. When a family needs critical medicine fast, we connect patients, licensed pharmacies, caregivers, and delivery — in one transparent flow.",
  url: process.env.NEXT_PUBLIC_SITE_URL || "https://dialx.in",
  ogImage: "/og-image.svg",
  links: {
    email: "hello@dialx.in",
    support: "support@dialx.in",
    partners: "partners@dialx.in",
  },
  launchTimeline: "Q3 2026",
};

export const navLinks = [
  { href: "#problem", label: "Problem" },
  { href: "#why-dialx", label: "Why DialX" },
  { href: "#services", label: "Services" },
  { href: "#security", label: "Security" },
  { href: "#waitlist", label: "Join" },
  { href: "#faq", label: "FAQ" },
];

export const narrativeAnswers = [
  { question: "What is DialX?", answer: "Emergency healthcare coordination platform" },
  { question: "What problem?", answer: "Families lose critical time in medical crises" },
  { question: "How?", answer: "One request, full coordination flow" },
  { question: "Why trust?", answer: "Licensed partners, verified Rx, privacy-first" },
  { question: "Why now?", answer: "Early access & shape the launch" },
];

export const trustRibbonBadges = [
  { label: "Privacy First", icon: "Shield" },
  { label: "Secure Platform", icon: "Lock" },
  { label: "Licensed Pharmacy Network", icon: "Building2" },
  { label: "Prescription-Based Workflow", icon: "FileCheck" },
  { label: "Emergency Coordination", icon: "Zap" },
  { label: "Cloud Infrastructure", icon: "Cloud" },
  { label: "Real-Time Updates", icon: "Activity" },
  { label: "Family Notifications", icon: "Bell" },
  { label: "Built for India", icon: "MapPin" },
];

export const problemScenarios = [
  {
    title: "Multiple pharmacy visits",
    description:
      "Driving or sending someone from store to store — not knowing who has the medicine in stock.",
    icon: "Store",
  },
  {
    title: "Endless phone calls",
    description:
      "Calling pharmacy after pharmacy with no visibility into availability, pricing, or timelines.",
    icon: "Phone",
  },
  {
    title: "Scattered messaging",
    description:
      "Families coordinating through WhatsApp groups and calls — with no single source of truth.",
    icon: "MessageCircle",
  },
  {
    title: "Time lost in crisis",
    description:
      "Every minute spent searching adds stress and risk when someone you love needs medicine urgently.",
    icon: "Clock",
  },
  {
    title: "Zero transparency",
    description:
      "No way to track where things stand — from prescription to pharmacy to delivery.",
    icon: "EyeOff",
  },
  {
    title: "Caregivers alone",
    description:
      "One family member carries the entire coordination burden while others wait anxiously.",
    icon: "UserX",
  },
];

export const problemStats = [
  { value: "47%", label: "of families call 3+ pharmacies during emergencies" },
  { value: "2.5 hrs", label: "average time lost locating critical medicines" },
  { value: "68%", label: "report stress from uncoordinated care logistics" },
];

export const currentApproaches = [
  {
    approach: "Visiting pharmacies individually",
    limitation: "Works for one store — but no visibility across the network",
    icon: "Store",
  },
  {
    approach: "Phone calls",
    limitation: "Manual, repetitive, and no shared status for the family",
    icon: "Phone",
  },
  {
    approach: "Search engines",
    limitation: "Shows options — not real-time availability or coordination",
    icon: "Search",
  },
  {
    approach: "Messaging apps",
    limitation: "Great for communication — not built for healthcare workflows",
    icon: "MessageCircle",
  },
];

export const fragmentationPoints = [
  {
    title: "Each tool solves one piece",
    description:
      "Pharmacies fulfil. Delivery moves packages. Families communicate. None connect in real time during an emergency.",
    icon: "Puzzle",
  },
  {
    title: "Families become coordinators",
    description:
      "Loved ones absorb the logistics burden — calling, comparing, hoping — when they should be caring.",
    icon: "Users",
  },
  {
    title: "No shared visibility",
    description:
      "Caregivers can't see availability, confirmation, or delivery status in one place.",
    icon: "EyeOff",
  },
  {
    title: "The missing layer",
    description:
      "India doesn't lack pharmacies or delivery. It lacks coordinated emergency healthcare workflows.",
    icon: "Layers",
  },
];

export const dialxComparisons = [
  {
    traditional: "Multiple pharmacy visits",
    dialx: "One coordinated request",
    icon: "MapPin",
  },
  {
    traditional: "Manual phone calls",
    dialx: "Connected licensed pharmacy network",
    icon: "Network",
  },
  {
    traditional: "No visibility",
    dialx: "Real-time status at every step",
    icon: "Activity",
  },
  {
    traditional: "Fragmented communication",
    dialx: "Unified coordination hub",
    icon: "Radio",
  },
  {
    traditional: "No tracking",
    dialx: "End-to-end request tracking",
    icon: "Route",
  },
  {
    traditional: "Separate family updates",
    dialx: "Live notifications for caregivers",
    icon: "Bell",
  },
];

export const productCapabilities = [
  {
    problem: "Don't know which pharmacy has your medicine",
    solution: "Find nearby participating pharmacies",
    benefit: "Instant network visibility during emergencies",
    icon: "Search",
  },
  {
    problem: "No one orchestrating the process",
    solution: "Coordinate emergency requests",
    benefit: "One platform manages the entire workflow",
    icon: "Radio",
  },
  {
    problem: "Family left guessing",
    solution: "Track progress live",
    benefit: "Real-time updates from request to delivery",
    icon: "Activity",
  },
  {
    problem: "Caregivers out of the loop",
    solution: "Family notifications",
    benefit: "Every loved one stays informed automatically",
    icon: "Bell",
  },
  {
    problem: "Prescription handling concerns",
    solution: "Secure prescription handling",
    benefit: "Verified, encrypted, compliance-ready workflow",
    icon: "FileCheck",
  },
  {
    problem: "No record of past requests",
    solution: "Digital request history",
    benefit: "Full audit trail for families and compliance",
    icon: "History",
  },
  {
    problem: "Crisis with no support",
    solution: "Emergency support coordination",
    benefit: "Purpose-built for urgent healthcare logistics",
    icon: "LifeBuoy",
  },
  {
    problem: "Data privacy fears",
    solution: "Privacy & security by design",
    benefit: "Encrypted, never sold, patient-controlled data",
    icon: "ShieldCheck",
  },
];

export const healthcareServices = [
  {
    title: "Emergency Medicine Coordination",
    description: "Locate and coordinate critical medicines when every minute counts.",
    benefit: "Faster access during medical crises",
    icon: "Siren",
  },
  {
    title: "Prescription Request Coordination",
    description: "Digital prescription upload with verification before any action.",
    benefit: "Safe, compliant request handling",
    icon: "FileCheck",
  },
  {
    title: "Licensed Pharmacy Network",
    description: "Smart matching with verified, licensed pharmacy partners nearby.",
    benefit: "Trusted fulfilment, not grey-market risk",
    icon: "Building2",
  },
  {
    title: "Delivery Coordination",
    description: "Seamless handoff to delivery partners for time-critical transport.",
    benefit: "Medicine reaches the patient faster",
    icon: "Truck",
  },
  {
    title: "Family Notifications",
    description: "Automatic status updates for every caregiver in the loop.",
    benefit: "No more anxious phone tag",
    icon: "Bell",
  },
  {
    title: "Caregiver Coordination",
    description: "Remote family members stay connected to the same live workflow.",
    benefit: "Coordinate care from anywhere",
    icon: "HeartHandshake",
  },
  {
    title: "Real-Time Tracking",
    description: "See availability, confirmation, and delivery status live.",
    benefit: "Full transparency, zero black boxes",
    icon: "MapPinned",
  },
  {
    title: "Digital Request History",
    description: "Encrypted records of past coordination requests.",
    benefit: "Continuity of care and compliance readiness",
    icon: "FolderLock",
  },
  {
    title: "Emergency Alerts",
    description: "Priority notifications when urgent requests need attention.",
    benefit: "Critical requests never get lost",
    icon: "AlertTriangle",
  },
  {
    title: "AI-Assisted Coordination",
    description: "Intelligent routing and prioritization — coming soon.",
    benefit: "Faster matching as we scale",
    icon: "Sparkles",
    badge: "Future",
  },
];

export const securityFeatures = [
  { title: "Encrypted Data Handling", description: "Health data encrypted in transit and at rest.", icon: "Lock" },
  { title: "Secure Communication", description: "Protected channels between all platform participants.", icon: "Shield" },
  { title: "Patient Privacy", description: "Privacy-by-design. Your data is never sold.", icon: "Eye" },
  { title: "Prescription Verification", description: "Every prescription validated before coordination.", icon: "FileCheck" },
  { title: "Audit Logs", description: "Complete trail of coordination actions for compliance.", icon: "ScrollText" },
  { title: "Human-in-the-Loop", description: "Critical decisions reviewed by qualified personnel.", icon: "UserCheck" },
  { title: "Role-Based Access", description: "Granular permissions for patients, partners, and staff.", icon: "KeyRound" },
  { title: "Platform Monitoring", description: "Continuous security monitoring and incident response.", icon: "Monitor" },
  { title: "Secure Cloud Infrastructure", description: "Enterprise-grade hosting with redundancy and backups.", icon: "Cloud" },
];

export const audiences = [
  { title: "Patients & Families", description: "Coordinate emergency medicine access with full transparency.", icon: "Users" },
  { title: "Caregivers", description: "Stay informed remotely while loved ones receive critical care.", icon: "HeartHandshake" },
  { title: "Licensed Pharmacies", description: "Receive qualified emergency requests from verified patients.", icon: "Store" },
  { title: "Delivery Partners", description: "Purpose-built logistics for time-critical healthcare deliveries.", icon: "Truck" },
];

export const advisoryExperts = [
  { role: "Healthcare Professionals", icon: "Stethoscope" },
  { role: "Pharmacy Experts", icon: "Pill" },
  { role: "Compliance Advisors", icon: "Scale" },
  { role: "Legal Experts", icon: "Gavel" },
  { role: "Cloud Architects", icon: "Cloud" },
  { role: "AI Engineers", icon: "Brain" },
  { role: "Security Specialists", icon: "ShieldCheck" },
];

export const techStack = [
  { name: "React", category: "Frontend" },
  { name: "Next.js", category: "Framework" },
  { name: "TypeScript", category: "Language" },
  { name: "Tailwind CSS", category: "Styling" },
  { name: "Vercel", category: "Hosting" },
  { name: "Supabase", category: "Database" },
  { name: "Resend", category: "Email" },
];

export const solutionFlow = [
  { step: 1, title: "Request", description: "Family uploads prescription & need" },
  { step: 2, title: "Verify", description: "Prescription validated for safety" },
  { step: 3, title: "Match", description: "Nearby licensed pharmacies checked" },
  { step: 4, title: "Confirm", description: "Availability confirmed in real time" },
  { step: 5, title: "Coordinate", description: "Pickup or delivery arranged" },
  { step: 6, title: "Track", description: "Everyone stays informed" },
];

export const howItWorksSteps = [
  {
    number: "01",
    title: "Tell us what you need",
    description: "Upload a prescription and describe the urgency. DialX verifies it before anything moves forward.",
  },
  {
    number: "02",
    title: "We coordinate availability",
    description: "DialX reaches licensed pharmacies near you, confirms stock, and matches the fastest path.",
  },
  {
    number: "03",
    title: "Your family stays in the loop",
    description: "From confirmation to doorstep — every caregiver sees live status updates.",
  },
];

export const roadmapItems = [
  {
    phase: "Phase 1",
    title: "Waitlist & Partnerships",
    status: "current" as const,
    items: ["Early supporter waitlist", "Pharmacy partner onboarding", "User research & city prioritization"],
  },
  {
    phase: "Phase 2",
    title: "Beta Launch",
    status: "upcoming" as const,
    items: ["Pilot in select cities", "Core coordination features", "Family notifications & tracking"],
  },
  {
    phase: "Phase 3",
    title: "Full Launch",
    status: "upcoming" as const,
    items: ["Multi-city expansion", "AI-assisted matching", "Advanced analytics & compliance"],
  },
];

export const waitlistBenefits = [
  { title: "Early platform access", description: "Be among the first to use DialX when we launch." },
  { title: "Priority onboarding", description: "Skip the queue in your city." },
  { title: "Exclusive updates", description: "Behind-the-scenes progress and launch news." },
  { title: "Shape the platform", description: "Your feedback directly influences features and cities." },
  { title: "Early supporter recognition", description: "Founding community member status at launch." },
  { title: "Launch benefits", description: "Special perks reserved for early believers." },
];

export const faqs = [
  {
    question: "What exactly is DialX?",
    answer: "DialX is an emergency healthcare coordination platform. We connect families, licensed pharmacies, and delivery partners in one transparent flow. We are not a pharmacy, hospital, or telemedicine service.",
  },
  {
    question: "Is DialX an online pharmacy?",
    answer: "No. DialX never stocks, sells, or delivers medicines. Licensed pharmacies fulfil every order. We provide the coordination layer.",
  },
  {
    question: "How is my health data protected?",
    answer: "Encryption, privacy-by-design, role-based access, audit logs, and never selling your data. Built with healthcare compliance requirements in mind.",
  },
  {
    question: "Why should I join the waitlist now?",
    answer: "Early supporters get priority access, shape our launch cities, and join a community building something India urgently needs. Free and takes two minutes.",
  },
  {
    question: "How can pharmacies partner with DialX?",
    answer: "Licensed pharmacies can apply through our partner form. We verify credentials before onboarding in each city.",
  },
  {
    question: "When will DialX launch?",
    answer: "Targeting launch in 2–3 months. Waitlist members receive first access and exclusive updates.",
  },
];

export const heroStats = [
  { value: 2500, suffix: "+", label: "Early supporters" },
  { value: 150, suffix: "+", label: "Pharmacy partners" },
  { value: 12, suffix: "", label: "Cities in pipeline" },
];

export const userTypes = [
  "Patient",
  "Family Member / Caregiver",
  "Healthcare Professional",
  "Pharmacy Owner",
  "Other",
];

export const orderEstimates = [
  "1-10 orders/day",
  "11-25 orders/day",
  "26-50 orders/day",
  "50+ orders/day",
];

// Legacy export for trust section positioning
export const trustIndicators = securityFeatures.slice(0, 6).map((s) => ({
  title: s.title,
  description: s.description,
  icon: s.icon,
}));
