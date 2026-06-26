"use client";

import { motion } from "framer-motion";
import { SectionHeader } from "@/components/shared/section-header";
import { techStack } from "@/lib/constants";

export function TechnologySection() {
  return (
    <section
      id="technology"
      className="section-padding section-transition bg-muted/40"
      aria-labelledby="technology-heading"
    >
      <div className="container-narrow">
        <SectionHeader
          eyebrow="Technology"
          title="Powered by modern & secure technology"
          description="Built with production-grade tools we actually use — no fake logos, no competitor brands."
        />

        <div className="grid grid-cols-2 gap-4 sm:grid-cols-3 lg:grid-cols-4">
          {techStack.map((tech, index) => (
            <motion.div
              key={tech.name}
              initial={{ opacity: 0, scale: 0.95 }}
              whileInView={{ opacity: 1, scale: 1 }}
              viewport={{ once: true }}
              transition={{ delay: index * 0.05 }}
              whileHover={{ y: -4 }}
              className="glass-card flex flex-col items-center justify-center rounded-xl p-6 text-center transition-shadow hover:shadow-md"
            >
              <p className="text-lg font-semibold text-foreground">{tech.name}</p>
              <p className="mt-1 text-xs text-muted-foreground">{tech.category}</p>
            </motion.div>
          ))}
        </div>
      </div>
    </section>
  );
}
