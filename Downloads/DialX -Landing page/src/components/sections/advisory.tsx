"use client";

import { motion } from "framer-motion";
import {
  Brain,
  Cloud,
  Gavel,
  Pill,
  Scale,
  ShieldCheck,
  Stethoscope,
  type LucideIcon,
} from "lucide-react";
import { SectionHeader } from "@/components/shared/section-header";
import { advisoryExperts } from "@/lib/constants";

const iconMap: Record<string, LucideIcon> = {
  Stethoscope,
  Pill,
  Scale,
  Gavel,
  Cloud,
  Brain,
  ShieldCheck,
};

export function AdvisorySection() {
  return (
    <section
      className="section-padding section-transition"
      aria-labelledby="advisory-heading"
    >
      <div className="container-narrow">
        <SectionHeader
          eyebrow="Expert Guidance"
          title="Designed with domain expertise"
          description="DialX is being built with input from healthcare, pharmacy, compliance, legal, cloud, AI, and security professionals."
        />

        <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-4">
          {advisoryExperts.map((expert, index) => {
            const Icon = iconMap[expert.icon] || Stethoscope;
            return (
              <motion.div
                key={expert.role}
                initial={{ opacity: 0, y: 12 }}
                whileInView={{ opacity: 1, y: 0 }}
                viewport={{ once: true }}
                transition={{ delay: index * 0.06 }}
                className="glass-card rounded-xl p-5 text-center"
              >
                <div className="mx-auto mb-3 flex size-12 items-center justify-center rounded-full bg-primary/10 text-primary">
                  <Icon className="size-5" aria-hidden="true" />
                </div>
                <p className="text-sm font-medium">{expert.role}</p>
              </motion.div>
            );
          })}
        </div>

        <p className="mx-auto mt-8 max-w-2xl text-center text-sm text-muted-foreground">
          Advisory input informs our compliance, security, and product design.
          No third-party healthcare company logos or implied endorsements are
          displayed on this page.
        </p>
      </div>
    </section>
  );
}
