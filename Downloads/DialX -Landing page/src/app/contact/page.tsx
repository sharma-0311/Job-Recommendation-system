import type { Metadata } from "next";
import Link from "next/link";
import { Mail, MapPin } from "lucide-react";
import { siteConfig } from "@/lib/constants";

export const metadata: Metadata = {
  title: "Contact Us",
  description: `Get in touch with the ${siteConfig.name} team for support, partnerships, or general inquiries.`,
  alternates: {
    canonical: `${siteConfig.url}/contact`,
  },
};

export default function ContactPage() {
  return (
    <div className="section-padding pt-28">
      <div className="container-narrow max-w-2xl">
        <h1 className="text-3xl font-semibold tracking-tight sm:text-4xl">
          Contact Us
        </h1>
        <p className="mt-4 text-lg text-muted-foreground">
          Have questions about DialX? We&apos;d love to hear from you.
        </p>

        <div className="mt-10 space-y-6">
          <div className="flex gap-4 rounded-xl border border-border bg-card p-6">
            <Mail className="size-5 shrink-0 text-primary" aria-hidden="true" />
            <div>
              <h2 className="font-semibold">General Inquiries</h2>
              <a
                href={`mailto:${siteConfig.links.email}`}
                className="mt-1 text-primary hover:underline"
              >
                {siteConfig.links.email}
              </a>
            </div>
          </div>

          <div className="flex gap-4 rounded-xl border border-border bg-card p-6">
            <Mail className="size-5 shrink-0 text-primary" aria-hidden="true" />
            <div>
              <h2 className="font-semibold">Partnerships</h2>
              <a
                href={`mailto:${siteConfig.links.partners}`}
                className="mt-1 text-primary hover:underline"
              >
                {siteConfig.links.partners}
              </a>
            </div>
          </div>

          <div className="flex gap-4 rounded-xl border border-border bg-card p-6">
            <Mail className="size-5 shrink-0 text-primary" aria-hidden="true" />
            <div>
              <h2 className="font-semibold">Support</h2>
              <a
                href={`mailto:${siteConfig.links.support}`}
                className="mt-1 text-primary hover:underline"
              >
                {siteConfig.links.support}
              </a>
            </div>
          </div>

          <div className="flex gap-4 rounded-xl border border-border bg-card p-6">
            <MapPin className="size-5 shrink-0 text-primary" aria-hidden="true" />
            <div>
              <h2 className="font-semibold">Location</h2>
              <p className="mt-1 text-muted-foreground">India</p>
            </div>
          </div>
        </div>

        <p className="mt-10 text-sm text-muted-foreground">
          Looking to join early access?{" "}
          <Link href="/#waitlist" className="text-primary hover:underline">
            Join the waitlist
          </Link>{" "}
          or{" "}
          <Link href="/#pharmacy-partner" className="text-primary hover:underline">
            apply as a pharmacy partner
          </Link>
          .
        </p>
      </div>
    </div>
  );
}
