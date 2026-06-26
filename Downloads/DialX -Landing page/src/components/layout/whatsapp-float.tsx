"use client";

import { WhatsAppButton } from "@/components/shared/whatsapp-button";

export function WhatsAppFloat() {
  return (
    <div className="fixed bottom-6 right-6 z-40">
      <WhatsAppButton
        location="floating_button"
        label="Chat"
        size="lg"
        className="size-14 rounded-full p-0 shadow-lg hover:shadow-xl [&_svg]:size-6"
      />
      <span className="sr-only">Join DialX community on WhatsApp</span>
    </div>
  );
}
