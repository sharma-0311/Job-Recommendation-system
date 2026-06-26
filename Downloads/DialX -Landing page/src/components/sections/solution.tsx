"use client";

import { motion } from "framer-motion";
import {
  Bell,
  Building2,
  FileCheck,
  Heart,
  Truck,
  Upload,
  Users,
} from "lucide-react";
import { SectionHeader } from "@/components/shared/section-header";
import { howItWorksSteps, solutionFlow } from "@/lib/constants";

const flowIcons = [Heart, Upload, FileCheck, Building2, Truck, Bell];

export function SolutionSection() {
  return (
    <section
      id="solution"
      className="section-padding"
      aria-labelledby="solution-heading"
    >
      <div className="container-narrow">
        <SectionHeader
          eyebrow="How DialX Solves It"
          title="One request. One flow. Everyone connected."
          description="Instead of families acting as the coordinator, DialX orchestrates the entire path — from prescription to pharmacy to doorstep — with full visibility."
        />

        {/* Visual flow diagram */}
        <motion.div
          initial={{ opacity: 0, y: 20 }}
          whileInView={{ opacity: 1, y: 0 }}
          viewport={{ once: true }}
          className="mb-14 overflow-hidden rounded-2xl border border-border bg-card shadow-sm"
        >
          <div className="border-b border-border bg-muted/40 px-6 py-4">
            <p className="text-sm font-medium">The coordination flow</p>
            <p className="text-xs text-muted-foreground">
              Request → Verify → Match → Confirm → Coordinate → Track
            </p>
          </div>

          <div className="hidden p-8 lg:block">
            <CoordinationDiagram />
          </div>

          <div className="grid gap-3 p-4 sm:grid-cols-2 sm:p-6 lg:hidden">
            {solutionFlow.map((step, index) => {
              const Icon = flowIcons[index] || Users;
              return (
                <div
                  key={step.step}
                  className="flex items-center gap-3 rounded-lg border border-border bg-background p-4"
                >
                  <div className="flex size-9 shrink-0 items-center justify-center rounded-full bg-primary text-xs font-bold text-primary-foreground">
                    {step.step}
                  </div>
                  <div>
                    <p className="text-sm font-semibold">{step.title}</p>
                    <p className="text-xs text-muted-foreground">
                      {step.description}
                    </p>
                  </div>
                  <Icon className="ml-auto size-4 text-muted-foreground" aria-hidden="true" />
                </div>
              );
            })}
          </div>
        </motion.div>

        {/* Three steps */}
        <div id="how-it-works" className="grid gap-8 lg:grid-cols-3">
          {howItWorksSteps.map((step, index) => (
            <motion.div
              key={step.number}
              initial={{ opacity: 0, y: 20 }}
              whileInView={{ opacity: 1, y: 0 }}
              viewport={{ once: true }}
              transition={{ delay: index * 0.1 }}
              className="relative rounded-xl border border-border bg-card p-6"
            >
              <span className="text-4xl font-bold text-primary/20">
                {step.number}
              </span>
              <h3 className="mt-2 text-lg font-semibold">{step.title}</h3>
              <p className="mt-2 text-sm leading-relaxed text-muted-foreground">
                {step.description}
              </p>
            </motion.div>
          ))}
        </div>
      </div>
    </section>
  );
}

function CoordinationDiagram() {
  const nodes = [
    { x: 70, y: 150, label: "Family", sub: "Submits need", r: 28 },
    { x: 200, y: 80, label: "Verify", sub: "Prescription", r: 24 },
    { x: 320, y: 150, label: "DialX", sub: "Coordinates", r: 36, hub: true },
    { x: 440, y: 80, label: "Pharmacy", sub: "Licensed", r: 28 },
    { x: 440, y: 220, label: "Delivery", sub: "Arranged", r: 24 },
    { x: 560, y: 150, label: "Family", sub: "Tracks live", r: 28 },
  ];

  return (
    <svg
      viewBox="0 0 630 300"
      className="mx-auto w-full max-w-3xl"
      role="img"
      aria-label="Flow diagram: family request through DialX to pharmacy and delivery, with live tracking back to family"
    >
      <motion.path
        d="M70 150 L200 80 L320 150 L440 80 L560 150"
        fill="none"
        stroke="oklch(0.75 0.02 240)"
        strokeWidth="2"
        strokeDasharray="6 4"
        initial={{ pathLength: 0 }}
        whileInView={{ pathLength: 1 }}
        viewport={{ once: true }}
        transition={{ duration: 1.5 }}
      />
      <motion.path
        d="M320 150 L440 220 L560 150"
        fill="none"
        stroke="oklch(0.55 0.12 165)"
        strokeWidth="2"
        strokeDasharray="6 4"
        initial={{ pathLength: 0 }}
        whileInView={{ pathLength: 1 }}
        viewport={{ once: true }}
        transition={{ duration: 1.5, delay: 0.3 }}
      />

      {nodes.map((node, i) => (
        <motion.g
          key={`${node.label}-${i}`}
          initial={{ opacity: 0, scale: 0.85 }}
          whileInView={{ opacity: 1, scale: 1 }}
          viewport={{ once: true }}
          transition={{ delay: 0.15 + i * 0.1 }}
        >
          <circle
            cx={node.x}
            cy={node.y}
            r={node.r}
            fill={node.hub ? "oklch(0.45 0.14 250)" : "white"}
            stroke={node.hub ? "oklch(0.45 0.14 250)" : "oklch(0.75 0.02 240)"}
            strokeWidth="2"
          />
          <text
            x={node.x}
            y={node.y - 2}
            textAnchor="middle"
            fill={node.hub ? "white" : "oklch(0.18 0.02 250)"}
            fontSize={node.hub ? 12 : 10}
            fontWeight="600"
          >
            {node.label}
          </text>
          <text
            x={node.x}
            y={node.y + 12}
            textAnchor="middle"
            fill={node.hub ? "oklch(0.9 0 0)" : "oklch(0.5 0.02 250)"}
            fontSize="8"
          >
            {node.sub}
          </text>
        </motion.g>
      ))}

      <text
        x={315}
        y={285}
        textAnchor="middle"
        fill="oklch(0.5 0.02 250)"
        fontSize="10"
      >
        One transparent path — no blind calls, no guesswork
      </text>
    </svg>
  );
}
