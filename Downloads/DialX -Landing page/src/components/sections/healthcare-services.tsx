"use client";

import { motion } from "framer-motion";
import {
  AlertTriangle,
  Bell,
  Building2,
  FileCheck,
  FolderLock,
  HeartHandshake,
  MapPinned,
  Sparkles,
  Siren,
  Truck,
  type LucideIcon,
} from "lucide-react";
import { SectionHeader } from "@/components/shared/section-header";
import { healthcareServices } from "@/lib/constants";
import { cn } from "@/lib/utils";

const iconMap: Record<string, LucideIcon> = {
  Siren,
  FileCheck,
  Building2,
  Truck,
  Bell,
  HeartHandshake,
  MapPinned,
  FolderLock,
  AlertTriangle,
  Sparkles,
};

export function HealthcareServicesSection() {
  return (
    <section
      id="services"
      className="section-padding section-transition"
      aria-labelledby="services-heading"
    >
      <div className="container-narrow">
        <SectionHeader
          eyebrow="Healthcare Services We Coordinate"
          title="Practical services for real emergencies"
          description="DialX enables coordinated healthcare logistics — not just software features, but services families actually need during a crisis."
        />

        <div className="grid gap-5 sm:grid-cols-2 lg:grid-cols-3">
          {healthcareServices.map((service, index) => {
            const Icon = iconMap[service.icon] || Siren;
            return (
              <motion.div
                key={service.title}
                initial={{ opacity: 0, y: 16 }}
                whileInView={{ opacity: 1, y: 0 }}
                viewport={{ once: true }}
                transition={{ delay: index * 0.04 }}
                whileHover={{ y: -4 }}
                className={cn(
                  "group glass-card rounded-xl p-6 transition-all hover:shadow-md",
                  service.badge && "border-dashed"
                )}
              >
                {service.badge && (
                  <span className="mb-3 inline-block rounded-full bg-accent px-2.5 py-0.5 text-xs font-medium text-accent-foreground">
                    {service.badge}
                  </span>
                )}
                <div className="mb-4 flex size-11 items-center justify-center rounded-xl bg-primary/10 text-primary transition-colors group-hover:bg-primary/20">
                  <Icon className="size-5" aria-hidden="true" />
                </div>
                <h3 className="font-semibold">{service.title}</h3>
                <p className="mt-2 text-sm leading-relaxed text-muted-foreground">
                  {service.description}
                </p>
                <p className="mt-3 text-xs font-medium text-trust">
                  → {service.benefit}
                </p>
              </motion.div>
            );
          })}
        </div>
      </div>
    </section>
  );
}
