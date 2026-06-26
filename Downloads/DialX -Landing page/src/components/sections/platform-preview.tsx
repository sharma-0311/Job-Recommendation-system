"use client";

import { motion } from "framer-motion";
import {
  Bell,
  Building2,
  CheckCircle2,
  FileCheck,
  MapPin,
  Truck,
} from "lucide-react";
import { SectionHeader } from "@/components/shared/section-header";

const previewSteps = [
  { icon: FileCheck, label: "Prescription verified", status: "complete" },
  { icon: Building2, label: "Pharmacy matched", status: "complete" },
  { icon: CheckCircle2, label: "Stock confirmed", status: "active" },
  { icon: Truck, label: "Delivery coordinated", status: "pending" },
  { icon: Bell, label: "Family notified", status: "pending" },
];

export function PlatformPreviewSection() {
  return (
    <section
      id="preview"
      className="section-padding section-transition bg-muted/40"
      aria-labelledby="preview-heading"
    >
      <div className="container-narrow">
        <SectionHeader
          eyebrow="Platform Preview"
          title="See coordination in action"
          description="A glimpse of how DialX keeps every step visible — from request to delivery."
        />

        <motion.div
          initial={{ opacity: 0, y: 20 }}
          whileInView={{ opacity: 1, y: 0 }}
          viewport={{ once: true }}
          className="mx-auto max-w-3xl overflow-hidden rounded-2xl border border-border bg-card shadow-xl"
        >
          <div className="flex items-center gap-2 border-b border-border bg-muted/50 px-4 py-3">
            <div className="size-3 rounded-full bg-destructive/60" />
            <div className="size-3 rounded-full bg-yellow-400/60" />
            <div className="size-3 rounded-full bg-trust/60" />
            <span className="ml-2 text-xs text-muted-foreground">
              DialX Coordination Dashboard
            </span>
          </div>

          <div className="p-6 sm:p-8">
            <div className="mb-6 flex items-center justify-between">
              <div>
                <p className="text-sm text-muted-foreground">Emergency Request</p>
                <p className="text-lg font-semibold">Medicine coordination in progress</p>
              </div>
              <div className="flex items-center gap-1.5 rounded-full bg-trust/10 px-3 py-1 text-xs font-medium text-trust">
                <MapPin className="size-3" aria-hidden="true" />
                Mumbai
              </div>
            </div>

            <div className="space-y-4">
              {previewSteps.map((step, index) => (
                <motion.div
                  key={step.label}
                  initial={{ opacity: 0, x: -10 }}
                  whileInView={{ opacity: 1, x: 0 }}
                  viewport={{ once: true }}
                  transition={{ delay: index * 0.1 }}
                  className={`flex items-center gap-4 rounded-xl border p-4 ${
                    step.status === "active"
                      ? "border-primary/30 bg-primary/5"
                      : step.status === "complete"
                        ? "border-trust/20 bg-trust/5"
                        : "border-border bg-muted/30"
                  }`}
                >
                  <div
                    className={`flex size-10 items-center justify-center rounded-lg ${
                      step.status === "complete"
                        ? "bg-trust/15 text-trust"
                        : step.status === "active"
                          ? "bg-primary/15 text-primary"
                          : "bg-muted text-muted-foreground"
                    }`}
                  >
                    <step.icon className="size-5" aria-hidden="true" />
                  </div>
                  <div className="flex-1">
                    <p className="font-medium">{step.label}</p>
                    <p className="text-xs text-muted-foreground capitalize">
                      {step.status === "active"
                        ? "In progress..."
                        : step.status === "complete"
                          ? "Completed"
                          : "Waiting"}
                    </p>
                  </div>
                  {step.status === "complete" && (
                    <CheckCircle2 className="size-5 text-trust" aria-hidden="true" />
                  )}
                </motion.div>
              ))}
            </div>

            <p className="mt-6 text-center text-xs text-muted-foreground">
              Preview mockup — actual interface at launch
            </p>
          </div>
        </motion.div>
      </div>
    </section>
  );
}
