import type { Metadata } from "next";
import { siteConfig } from "@/lib/constants";

export const metadata: Metadata = {
  title: "Privacy Policy",
  description: `Privacy Policy for ${siteConfig.name} — how we collect, use, and protect your personal and health information.`,
  alternates: {
    canonical: `${siteConfig.url}/privacy`,
  },
};

export default function PrivacyPage() {
  return (
    <article className="section-padding pt-28">
      <div className="container-narrow max-w-3xl prose-policy">
        <h1 className="text-3xl font-semibold tracking-tight sm:text-4xl">
          Privacy Policy
        </h1>
        <p className="mt-4 text-muted-foreground">
          Last updated: June 26, 2026
        </p>

        <div className="mt-10 space-y-8 text-muted-foreground">
          <section>
            <h2 className="text-xl font-semibold text-foreground">
              1. Introduction
            </h2>
            <p className="mt-3 leading-relaxed">
              {siteConfig.name} (&ldquo;we,&rdquo; &ldquo;our,&rdquo; or
              &ldquo;us&rdquo;) is committed to protecting your privacy. This
              Privacy Policy explains how we collect, use, disclose, and
              safeguard your information when you visit our website or join our
              waitlist.
            </p>
          </section>

          <section>
            <h2 className="text-xl font-semibold text-foreground">
              2. Information We Collect
            </h2>
            <p className="mt-3 leading-relaxed">We may collect:</p>
            <ul className="mt-3 list-disc space-y-2 pl-6">
              <li>Name, email address, and phone number</li>
              <li>City and user type (patient, caregiver, pharmacy, etc.)</li>
              <li>Responses to survey questions about healthcare challenges</li>
              <li>Pharmacy business information for partnership applications</li>
              <li>Technical data: IP address, browser type, device information</li>
              <li>Usage data via analytics tools (Google Analytics, Clarity)</li>
            </ul>
          </section>

          <section>
            <h2 className="text-xl font-semibold text-foreground">
              3. How We Use Your Information
            </h2>
            <ul className="mt-3 list-disc space-y-2 pl-6">
              <li>To manage waitlist and partnership registrations</li>
              <li>To send product updates and launch notifications</li>
              <li>To improve our platform and user experience</li>
              <li>To conduct user research and validate product-market fit</li>
              <li>To comply with legal obligations</li>
            </ul>
          </section>

          <section>
            <h2 className="text-xl font-semibold text-foreground">
              4. Data Protection
            </h2>
            <p className="mt-3 leading-relaxed">
              We implement appropriate technical and organizational measures to
              protect your personal information. Health-related data will be
              handled with enhanced security measures when our platform launches.
              We do not sell your personal information to third parties.
            </p>
          </section>

          <section>
            <h2 className="text-xl font-semibold text-foreground">
              5. Your Rights
            </h2>
            <p className="mt-3 leading-relaxed">
              You have the right to access, correct, or delete your personal
              information. You may withdraw consent for marketing communications
              at any time by contacting us at {siteConfig.links.email}.
            </p>
          </section>

          <section>
            <h2 className="text-xl font-semibold text-foreground">
              6. Contact
            </h2>
            <p className="mt-3 leading-relaxed">
              For privacy-related inquiries, contact us at{" "}
              <a
                href={`mailto:${siteConfig.links.email}`}
                className="text-primary hover:underline"
              >
                {siteConfig.links.email}
              </a>
              .
            </p>
          </section>
        </div>
      </div>
    </article>
  );
}
