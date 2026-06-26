"use client";

import { motion } from "framer-motion";
import {
  Cloud,
  Eye,
  FileCheck,
  KeyRound,
  Lock,
  Monitor,
  ScrollText,
  Shield,
  UserCheck,
  type LucideIcon,
} from "lucide-react";
import { SectionHeader } from "@/components/shared/section-header";
import { securityFeatures } from "@/lib/constants";

const iconMap: Record<string, LucideIcon> = {
  Lock,
  Shield,
  Eye,
  FileCheck,
  ScrollText,
  UserCheck,
  KeyRound,
  Monitor,
  Cloud,
};

export function SecurityComplianceSection() {
  return (
    <section
      id="security"
      className="section-padding section-transition bg-muted/40"
      aria-labelledby="security-heading"
    >
      <div className="container-narrow">
        <SectionHeader
          eyebrow="Security & Compliance"
          title="Enterprise-grade protection for healthcare data"
          description="Built with the security, privacy, and compliance standards that healthcare coordination demands."
        />

        <motion.div
          initial={{ opacity: 0, y: 12 }}
          whileInView={{ opacity: 1, y: 0 }}
          viewport={{ once: true }}
          className="mx-auto mb-12 max-w-3xl rounded-2xl border-2 border-trust/30 bg-trust/5 p-6 text-center sm:p-8"
        >
          <p className="text-lg font-semibold text-foreground">
            DialX is not a pharmacy. Not a hospital. Not telemedicine.
          </p>
          <p className="mt-3 text-muted-foreground leading-relaxed">
            We are the coordination layer — with encrypted data, verified
            prescriptions, licensed partners only, and full audit trails. No
            misleading affiliations. No implied endorsements.
          </p>
        </motion.div>

        <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-3">
          {securityFeatures.map((item, index) => {
            const Icon = iconMap[item.icon] || Shield;
            return (
              <motion.div
                key={item.title}
                initial={{ opacity: 0, y: 12 }}
                whileInView={{ opacity: 1, y: 0 }}
                viewport={{ once: true }}
                transition={{ delay: index * 0.04 }}
                className="glass-card rounded-xl p-5"
              >
                <div className="mb-3 flex size-10 items-center justify-center rounded-lg bg-trust/10 text-trust">
                  <Icon className="size-5" aria-hidden="true" />
                </div>
                <h3 className="font-semibold">{item.title}</h3>
                <p className="mt-1.5 text-sm leading-relaxed text-muted-foreground">
                  {item.description}
                </p>
              </motion.div>
            );
          })}
        </div>
      </div>
    </section>
  );
}
