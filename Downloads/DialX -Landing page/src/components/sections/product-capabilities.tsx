"use client";

import { motion } from "framer-motion";
import {
  Activity,
  Bell,
  FileCheck,
  History,
  LifeBuoy,
  Radio,
  Search,
  ShieldCheck,
  ArrowDown,
  type LucideIcon,
} from "lucide-react";
import { SectionHeader } from "@/components/shared/section-header";
import { productCapabilities } from "@/lib/constants";

const iconMap: Record<string, LucideIcon> = {
  Search,
  Radio,
  Activity,
  Bell,
  FileCheck,
  History,
  LifeBuoy,
  ShieldCheck,
};

export function ProductCapabilitiesSection() {
  return (
    <section
      id="capabilities"
      className="section-padding section-transition bg-muted/40"
      aria-labelledby="capabilities-heading"
    >
      <div className="container-narrow">
        <SectionHeader
          eyebrow="Platform Capabilities"
          title="Every capability solves a real emergency problem"
          description="Not a feature list — a problem-to-solution map built for families in crisis."
        />

        <div className="grid gap-6 sm:grid-cols-2">
          {productCapabilities.map((cap, index) => {
            const Icon = iconMap[cap.icon] || Search;
            return (
              <motion.div
                key={cap.solution}
                initial={{ opacity: 0, y: 20 }}
                whileInView={{ opacity: 1, y: 0 }}
                viewport={{ once: true }}
                transition={{ delay: index * 0.05 }}
                whileHover={{ y: -4 }}
                className="group glass-card rounded-2xl p-6 transition-shadow hover:shadow-lg"
              >
                <div className="mb-4 flex size-11 items-center justify-center rounded-xl bg-primary/10 text-primary">
                  <Icon className="size-5" aria-hidden="true" />
                </div>

                <div className="space-y-3">
                  <div className="rounded-lg bg-destructive/5 px-3 py-2">
                    <p className="text-xs font-semibold uppercase tracking-wider text-destructive/80">
                      Problem
                    </p>
                    <p className="mt-0.5 text-sm text-muted-foreground">
                      {cap.problem}
                    </p>
                  </div>

                  <ArrowDown className="mx-auto size-4 text-primary/40" aria-hidden="true" />

                  <div className="rounded-lg bg-primary/5 px-3 py-2">
                    <p className="text-xs font-semibold uppercase tracking-wider text-primary">
                      Solution
                    </p>
                    <p className="mt-0.5 text-sm font-medium text-foreground">
                      {cap.solution}
                    </p>
                  </div>

                  <ArrowDown className="mx-auto size-4 text-trust/40" aria-hidden="true" />

                  <div className="rounded-lg bg-trust/5 px-3 py-2">
                    <p className="text-xs font-semibold uppercase tracking-wider text-trust">
                      Benefit
                    </p>
                    <p className="mt-0.5 text-sm text-muted-foreground">
                      {cap.benefit}
                    </p>
                  </div>
                </div>
              </motion.div>
            );
          })}
        </div>
      </div>
    </section>
  );
}
