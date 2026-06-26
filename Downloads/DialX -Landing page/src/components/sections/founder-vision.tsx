"use client";

import { motion } from "framer-motion";
import { Quote } from "lucide-react";

export function FounderVisionSection() {
  return (
    <section
      id="vision"
      className="section-padding"
      aria-labelledby="founder-vision-heading"
    >
      <div className="container-narrow max-w-3xl">
        <motion.div
          initial={{ opacity: 0, y: 16 }}
          whileInView={{ opacity: 1, y: 0 }}
          viewport={{ once: true }}
          className="relative rounded-2xl border border-border bg-card p-8 sm:p-10"
        >
          <Quote
            className="absolute top-6 right-6 size-10 text-primary/10"
            aria-hidden="true"
          />

          <p className="mb-3 text-sm font-semibold uppercase tracking-wider text-primary">
            From the Founding Team
          </p>
          <h2
            id="founder-vision-heading"
            className="text-xl font-semibold leading-relaxed sm:text-2xl"
          >
            &ldquo;We built DialX because no family should become the emergency
            coordinator — calling ten pharmacies at midnight, with no
            visibility and no one to help.&rdquo;
          </h2>
          <p className="mt-5 leading-relaxed text-muted-foreground">
            This started as a personal experience, not a market slide. We&apos;re
            building infrastructure India&apos;s healthcare system needs: a
            trusted, compliant coordination layer that works alongside existing
            pharmacies and delivery networks — not against them.
          </p>
          <p className="mt-4 text-sm font-medium text-foreground">
            — DialX Founding Team
          </p>
        </motion.div>
      </div>
    </section>
  );
}
