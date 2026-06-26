"use client";

import { motion } from "framer-motion";
import {
  Layers,
  MessageCircle,
  Phone,
  Puzzle,
  Search,
  Store,
  Users,
  EyeOff,
  type LucideIcon,
} from "lucide-react";
import { SectionHeader } from "@/components/shared/section-header";
import { currentApproaches, fragmentationPoints } from "@/lib/constants";

const approachIcons: Record<string, LucideIcon> = {
  Store,
  Phone,
  Search,
  MessageCircle,
};

const gapIcons: Record<string, LucideIcon> = {
  Puzzle,
  Users,
  EyeOff,
  Layers,
};

export function FragmentationSection() {
  return (
    <section
      id="fragmentation"
      className="section-padding section-transition bg-muted/40"
      aria-labelledby="fragmentation-heading"
    >
      <div className="container-narrow">
        <SectionHeader
          eyebrow="Why Current Solutions Fall Short"
          title="Each approach solves part of the journey — not the whole emergency"
          description="Visiting pharmacies, making calls, searching online, and messaging family each help in their own way. None connect the full workflow when urgency strikes."
        />

        <div className="mb-14 grid gap-4 sm:grid-cols-2">
          {currentApproaches.map((item, index) => {
            const Icon = approachIcons[item.icon] || Search;
            return (
              <motion.div
                key={item.approach}
                initial={{ opacity: 0, x: index % 2 === 0 ? -12 : 12 }}
                whileInView={{ opacity: 1, x: 0 }}
                viewport={{ once: true }}
                transition={{ delay: index * 0.08 }}
                className="flex gap-4 rounded-xl border border-border bg-card p-5"
              >
                <div className="flex size-10 shrink-0 items-center justify-center rounded-lg bg-muted text-muted-foreground">
                  <Icon className="size-5" aria-hidden="true" />
                </div>
                <div>
                  <h3 className="font-semibold">{item.approach}</h3>
                  <p className="mt-1 text-sm text-muted-foreground">
                    {item.limitation}
                  </p>
                </div>
              </motion.div>
            );
          })}
        </div>

        <div className="grid gap-5 sm:grid-cols-2">
          {fragmentationPoints.map((point, index) => {
            const Icon = gapIcons[point.icon] || Layers;
            return (
              <motion.div
                key={point.title}
                initial={{ opacity: 0, y: 16 }}
                whileInView={{ opacity: 1, y: 0 }}
                viewport={{ once: true }}
                transition={{ delay: index * 0.08 }}
                className="glass-card rounded-xl p-6"
              >
                <Icon className="mb-3 size-5 text-primary" aria-hidden="true" />
                <h3 className="font-semibold">{point.title}</h3>
                <p className="mt-2 text-sm leading-relaxed text-muted-foreground">
                  {point.description}
                </p>
              </motion.div>
            );
          })}
        </div>

        <motion.p
          initial={{ opacity: 0 }}
          whileInView={{ opacity: 1 }}
          viewport={{ once: true }}
          className="mx-auto mt-12 max-w-2xl text-center text-lg font-medium text-foreground"
        >
          DialX doesn&apos;t replace pharmacies or delivery — it coordinates them
          into one emergency healthcare workflow.
        </motion.p>
      </div>
    </section>
  );
}
