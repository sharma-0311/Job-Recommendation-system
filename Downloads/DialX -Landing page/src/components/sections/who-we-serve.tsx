"use client";

import { motion } from "framer-motion";
import {
  HeartHandshake,
  Store,
  Truck,
  Users,
  type LucideIcon,
} from "lucide-react";
import { SectionHeader } from "@/components/shared/section-header";
import { audiences } from "@/lib/constants";

const iconMap: Record<string, LucideIcon> = {
  Users,
  HeartHandshake,
  Store,
  Truck,
};

export function WhoWeServeSection() {
  return (
    <section
      id="who-we-serve"
      className="section-padding section-transition"
      aria-labelledby="who-we-serve-heading"
    >
      <div className="container-narrow">
        <SectionHeader
          eyebrow="Who We Serve"
          title="Built for everyone in the emergency care chain"
          description="DialX creates value for every stakeholder — patients, families, pharmacies, and delivery partners."
        />

        <div className="grid gap-6 sm:grid-cols-2">
          {audiences.map((audience, index) => {
            const Icon = iconMap[audience.icon] || Users;
            return (
              <motion.div
                key={audience.title}
                initial={{ opacity: 0, y: 16 }}
                whileInView={{ opacity: 1, y: 0 }}
                viewport={{ once: true }}
                transition={{ delay: index * 0.1 }}
                whileHover={{ scale: 1.02 }}
                className="flex gap-5 glass-card rounded-2xl p-6"
              >
                <div className="flex size-12 shrink-0 items-center justify-center rounded-xl bg-primary/10 text-primary">
                  <Icon className="size-6" aria-hidden="true" />
                </div>
                <div>
                  <h3 className="text-lg font-semibold">{audience.title}</h3>
                  <p className="mt-2 text-muted-foreground leading-relaxed">
                    {audience.description}
                  </p>
                </div>
              </motion.div>
            );
          })}
        </div>
      </div>
    </section>
  );
}
