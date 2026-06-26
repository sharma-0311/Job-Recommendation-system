"use client";

import { motion } from "framer-motion";
import {
  Activity,
  ArrowRight,
  Bell,
  MapPin,
  Network,
  Radio,
  Route,
  X,
  Check,
} from "lucide-react";
import { SectionHeader } from "@/components/shared/section-header";
import { dialxComparisons } from "@/lib/constants";

const icons = [MapPin, Network, Activity, Radio, Route, Bell];

export function WhyDialXSection() {
  return (
    <section
      id="why-dialx"
      className="section-padding section-transition"
      aria-labelledby="why-dialx-heading"
    >
      <div className="container-narrow">
        <SectionHeader
          eyebrow="Why Choose DialX"
          title="Traditional process vs. coordinated emergency care"
          description="See how DialX transforms fragmented emergency logistics into one transparent, trackable workflow."
        />

        <div className="overflow-hidden rounded-2xl border border-border bg-card shadow-sm">
          <div className="grid grid-cols-2 border-b border-border bg-muted/50 text-center text-sm font-semibold">
            <div className="border-r border-border px-4 py-4 text-muted-foreground">
              Traditional Process
            </div>
            <div className="px-4 py-4 text-primary">DialX Platform</div>
          </div>

          {dialxComparisons.map((row, index) => {
            const Icon = icons[index] || ArrowRight;
            return (
              <motion.div
                key={row.traditional}
                initial={{ opacity: 0, x: -10 }}
                whileInView={{ opacity: 1, x: 0 }}
                viewport={{ once: true }}
                transition={{ delay: index * 0.06 }}
                className="grid grid-cols-2 border-b border-border last:border-b-0"
              >
                <div className="flex items-start gap-3 border-r border-border bg-destructive/5 px-4 py-5 sm:px-6">
                  <X className="mt-0.5 size-4 shrink-0 text-destructive/70" aria-hidden="true" />
                  <span className="text-sm text-muted-foreground sm:text-base">
                    {row.traditional}
                  </span>
                </div>
                <div className="flex items-start gap-3 bg-trust/5 px-4 py-5 sm:px-6">
                  <div className="flex size-8 shrink-0 items-center justify-center rounded-lg bg-trust/15 text-trust">
                    <Icon className="size-4" aria-hidden="true" />
                  </div>
                  <span className="text-sm font-medium text-foreground sm:text-base">
                    {row.dialx}
                  </span>
                  <Check className="ml-auto mt-0.5 size-4 shrink-0 text-trust" aria-hidden="true" />
                </div>
              </motion.div>
            );
          })}
        </div>
      </div>
    </section>
  );
}
