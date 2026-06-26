import type { Metadata } from "next";
import { siteConfig } from "@/lib/constants";

export const metadata: Metadata = {
  title: "Terms & Conditions",
  description: `Terms and Conditions for using the ${siteConfig.name} website and pre-launch services.`,
  alternates: {
    canonical: `${siteConfig.url}/terms`,
  },
};

export default function TermsPage() {
  return (
    <article className="section-padding pt-28">
      <div className="container-narrow max-w-3xl">
        <h1 className="text-3xl font-semibold tracking-tight sm:text-4xl">
          Terms &amp; Conditions
        </h1>
        <p className="mt-4 text-muted-foreground">
          Last updated: June 26, 2026
        </p>

        <div className="mt-10 space-y-8 text-muted-foreground">
          <section>
            <h2 className="text-xl font-semibold text-foreground">
              1. Acceptance of Terms
            </h2>
            <p className="mt-3 leading-relaxed">
              By accessing or using the {siteConfig.name} website, you agree to
              be bound by these Terms &amp; Conditions. If you do not agree, please
              do not use our services.
            </p>
          </section>

          <section>
            <h2 className="text-xl font-semibold text-foreground">
              2. Pre-Launch Status
            </h2>
            <p className="mt-3 leading-relaxed">
              {siteConfig.name} is currently in pre-launch phase. The platform
              described on this website is under development. Joining the
              waitlist does not guarantee access or create any binding service
              agreement.
            </p>
          </section>

          <section>
            <h2 className="text-xl font-semibold text-foreground">
              3. Not a Medical Service
            </h2>
            <p className="mt-3 leading-relaxed">
              {siteConfig.name} is a healthcare coordination platform, not a
              medical provider, pharmacy, or telemedicine service. We do not
              provide medical advice, diagnosis, or treatment. In medical
              emergencies, always contact emergency services (112/108) or visit
              the nearest hospital.
            </p>
          </section>

          <section>
            <h2 className="text-xl font-semibold text-foreground">
              4. Waitlist &amp; Partnerships
            </h2>
            <p className="mt-3 leading-relaxed">
              Waitlist registration and pharmacy partnership applications are
              subject to review. We reserve the right to accept or decline
              applications at our discretion. Information provided must be
              accurate and truthful.
            </p>
          </section>

          <section>
            <h2 className="text-xl font-semibold text-foreground">
              5. Intellectual Property
            </h2>
            <p className="mt-3 leading-relaxed">
              All content on this website, including text, graphics, logos, and
              software, is the property of {siteConfig.name} and protected by
              applicable intellectual property laws.
            </p>
          </section>

          <section>
            <h2 className="text-xl font-semibold text-foreground">
              6. Limitation of Liability
            </h2>
            <p className="mt-3 leading-relaxed">
              To the maximum extent permitted by law, {siteConfig.name} shall not
              be liable for any indirect, incidental, or consequential damages
              arising from your use of this website during the pre-launch phase.
            </p>
          </section>

          <section>
            <h2 className="text-xl font-semibold text-foreground">
              7. Contact
            </h2>
            <p className="mt-3 leading-relaxed">
              Questions about these terms? Contact us at{" "}
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
