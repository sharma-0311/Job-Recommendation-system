"use client";

import { motion } from "framer-motion";
import {
  Clock,
  EyeOff,
  MessageCircle,
  Phone,
  Store,
  UserX,
  type LucideIcon,
} from "lucide-react";
import { SectionHeader } from "@/components/shared/section-header";
import { problemScenarios, problemStats } from "@/lib/constants";

const iconMap: Record<string, LucideIcon> = {
  Store,
  Phone,
  MessageCircle,
  Clock,
  EyeOff,
  UserX,
};

export function ProblemSection() {
  return (
    <section
      id="problem"
      className="section-padding section-transition"
      aria-labelledby="problem-heading"
    >
      <div className="container-narrow">
        <SectionHeader
          eyebrow="The Real Problem"
          title="When emergencies strike, families become the coordinators"
          description="A medical crisis is already overwhelming. Finding medicine shouldn't mean hours of phone calls, pharmacy visits, and anxious waiting — with no visibility into what's happening."
        />

        <motion.blockquote
          initial={{ opacity: 0, y: 16 }}
          whileInView={{ opacity: 1, y: 0 }}
          viewport={{ once: true }}
          className="mx-auto mb-12 max-w-3xl rounded-2xl border border-primary/20 bg-primary/5 p-6 text-center sm:p-8"
        >
          <p className="text-lg leading-relaxed text-foreground italic">
            &ldquo;It&apos;s late at night. The prescription is ready. Now you&apos;re
            calling pharmacy after pharmacy — praying someone has it in
            stock.&rdquo;
          </p>
        </motion.blockquote>

        <div className="mb-14 grid gap-4 sm:grid-cols-3">
          {problemStats.map((stat, index) => (
            <motion.div
              key={stat.label}
              initial={{ opacity: 0, y: 12 }}
              whileInView={{ opacity: 1, y: 0 }}
              viewport={{ once: true }}
              transition={{ delay: index * 0.08 }}
              className="glass-card rounded-xl p-6 text-center"
            >
              <p className="text-3xl font-bold text-foreground">{stat.value}</p>
              <p className="mt-2 text-sm leading-relaxed text-muted-foreground">
                {stat.label}
              </p>
            </motion.div>
          ))}
        </div>

        <div className="grid gap-5 sm:grid-cols-2 lg:grid-cols-3">
          {problemScenarios.map((item, index) => {
            const Icon = iconMap[item.icon] || Clock;
            return (
              <motion.div
                key={item.title}
                initial={{ opacity: 0, y: 20 }}
                whileInView={{ opacity: 1, y: 0 }}
                viewport={{ once: true }}
                transition={{ delay: index * 0.06 }}
                whileHover={{ y: -4 }}
                className="group glass-card rounded-xl p-6 transition-shadow hover:shadow-md"
              >
                <div className="mb-4 flex size-11 items-center justify-center rounded-xl bg-destructive/10 text-destructive transition-colors group-hover:bg-destructive/15">
                  <Icon className="size-5" aria-hidden="true" />
                </div>
                <h3 className="font-semibold">{item.title}</h3>
                <p className="mt-2 text-sm leading-relaxed text-muted-foreground">
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
