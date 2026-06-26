"use client";

import { motion } from "framer-motion";
import { Check, Circle } from "lucide-react";
import { SectionHeader } from "@/components/shared/section-header";
import { roadmapItems } from "@/lib/constants";
import { cn } from "@/lib/utils";

export function RoadmapSection() {
  return (
    <section
      id="roadmap"
      className="section-padding section-transition"
      aria-labelledby="roadmap-heading"
    >
      <div className="container-narrow">
        <SectionHeader
          eyebrow="Roadmap"
          title="Our journey to launch"
          description="Transparent about where we are and where we're headed."
        />

        <div className="grid gap-6 lg:grid-cols-3">
          {roadmapItems.map((phase, index) => (
            <motion.div
              key={phase.phase}
              initial={{ opacity: 0, y: 16 }}
              whileInView={{ opacity: 1, y: 0 }}
              viewport={{ once: true }}
              transition={{ delay: index * 0.12 }}
              className={cn(
                "glass-card rounded-xl p-6",
                phase.status === "current" && "ring-2 ring-primary/20"
              )}
            >
              <div className="mb-4 flex items-center justify-between">
                <span className="text-sm font-medium text-primary">
                  {phase.phase}
                </span>
                <span
                  className={cn(
                    "rounded-full px-2.5 py-0.5 text-xs font-medium",
                    phase.status === "current"
                      ? "bg-primary/10 text-primary"
                      : "bg-muted text-muted-foreground"
                  )}
                >
                  {phase.status === "current" ? "Current" : "Upcoming"}
                </span>
              </div>
              <h3 className="text-lg font-semibold">{phase.title}</h3>
              <ul className="mt-4 space-y-2.5">
                {phase.items.map((item) => (
                  <li key={item} className="flex items-start gap-2 text-sm">
                    {phase.status === "current" ? (
                      <Check className="mt-0.5 size-4 shrink-0 text-trust" aria-hidden="true" />
                    ) : (
                      <Circle className="mt-0.5 size-4 shrink-0 text-muted-foreground" aria-hidden="true" />
                    )}
                    <span className="text-muted-foreground">{item}</span>
                  </li>
                ))}
              </ul>
            </motion.div>
          ))}
        </div>
      </div>
    </section>
  );
}
