"use client";

import { motion } from "framer-motion";
import {
  Activity,
  Bell,
  Building2,
  Cloud,
  FileCheck,
  Lock,
  MapPin,
  Shield,
  Zap,
  type LucideIcon,
} from "lucide-react";
import { trustRibbonBadges } from "@/lib/constants";

const iconMap: Record<string, LucideIcon> = {
  Shield,
  Lock,
  Building2,
  FileCheck,
  Zap,
  Cloud,
  Activity,
  Bell,
  MapPin,
};

export function TrustRibbon() {
  return (
    <section
      className="border-y border-border/60 bg-muted/30 py-6"
      aria-label="Trust indicators"
    >
      <div className="container-narrow overflow-hidden px-4 sm:px-6 lg:px-8">
        <div className="trust-marquee-wrap flex trust-marquee gap-3 sm:gap-4">
          {[...trustRibbonBadges, ...trustRibbonBadges].map((badge, i) => {
            const Icon = iconMap[badge.icon] || Shield;
            return (
              <div
                key={`${badge.label}-${i}`}
                className="flex shrink-0 items-center gap-2.5 rounded-full border border-border/80 bg-card/90 px-4 py-2 shadow-sm backdrop-blur-sm"
              >
                <div className="flex size-7 items-center justify-center rounded-full bg-trust/10 text-trust">
                  <Icon className="size-3.5" aria-hidden="true" />
                </div>
                <span className="whitespace-nowrap text-sm font-medium text-foreground">
                  {badge.label}
                </span>
              </div>
            );
          })}
        </div>

        {/* Static grid on larger screens */}
        <div className="mt-0 hidden flex-wrap justify-center gap-3 lg:flex">
          {trustRibbonBadges.map((badge, index) => {
            const Icon = iconMap[badge.icon] || Shield;
            return (
              <motion.div
                key={badge.label}
                initial={{ opacity: 0, y: 8 }}
                whileInView={{ opacity: 1, y: 0 }}
                viewport={{ once: true }}
                transition={{ delay: index * 0.04 }}
                className="flex items-center gap-2.5 rounded-full border border-border/80 bg-card px-4 py-2 shadow-sm"
              >
                <div className="flex size-7 items-center justify-center rounded-full bg-trust/10 text-trust">
                  <Icon className="size-3.5" aria-hidden="true" />
                </div>
                <span className="text-sm font-medium">{badge.label}</span>
              </motion.div>
            );
          })}
        </div>
      </div>
    </section>
  );
}
