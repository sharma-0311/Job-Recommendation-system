import type { Metadata } from "next";
import Link from "next/link";
import { CheckCircle2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { WhatsAppButton } from "@/components/shared/whatsapp-button";
import { siteConfig } from "@/lib/constants";

export const metadata: Metadata = {
  title: "Welcome to the Waitlist",
  description: `You've successfully joined the ${siteConfig.name} waitlist.`,
  robots: { index: false, follow: false },
};

export default function WaitlistSuccessPage() {
  return (
    <div className="flex min-h-[70vh] items-center justify-center section-padding pt-28">
      <div className="container-narrow max-w-md text-center">
        <div className="mx-auto mb-6 flex size-20 items-center justify-center rounded-full bg-trust/10">
          <CheckCircle2 className="size-10 text-trust" aria-hidden="true" />
        </div>
        <h1 className="text-3xl font-semibold">You&apos;re on the list!</h1>
        <p className="mt-4 text-muted-foreground">
          Thank you for joining the {siteConfig.name} waitlist. Check your email
          for a confirmation. We&apos;ll notify you with priority access when we
          launch.
        </p>
        <div className="mt-6 flex flex-col items-center gap-3 sm:flex-row sm:justify-center">
          <Button asChild>
            <Link href="/">Back to Home</Link>
          </Button>
          <WhatsAppButton
            location="success_page"
            label="Join WhatsApp Community"
            variant="whatsapp"
          />
        </div>
      </div>
    </div>
  );
}
