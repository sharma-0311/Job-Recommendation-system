"use client";

import { motion } from "framer-motion";
import { ArrowRight } from "lucide-react";
import { Button } from "@/components/ui/button";
import { CountUp } from "@/components/shared/count-up";
import { getWhatsAppJoinUrl } from "@/lib/whatsapp";
import { heroStats, narrativeAnswers } from "@/lib/constants";
import { trackCTAClick } from "@/lib/analytics";

export function HeroSection() {
  return (
    <section
      className="relative overflow-hidden pt-28 pb-16 sm:pt-32 sm:pb-20 lg:pt-36 lg:pb-24"
      aria-labelledby="hero-heading"
    >
      <div className="absolute inset-0 -z-10 gradient-subtle" />

      <div className="container-narrow px-4 sm:px-6 lg:px-8">
        <motion.div
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.6 }}
          className="mx-auto max-w-3xl text-center"
        >
          <p className="mb-4 text-sm font-semibold uppercase tracking-wider text-primary">
            What is DialX?
          </p>

          <h1
            id="hero-heading"
            className="text-balance text-4xl font-semibold tracking-tight text-foreground sm:text-5xl lg:text-[3.25rem] lg:leading-tight"
          >
            India&apos;s emergency healthcare{" "}
            <span className="text-primary">coordination platform</span>
          </h1>

          <p className="mx-auto mt-6 max-w-2xl text-lg leading-relaxed text-muted-foreground sm:text-xl">
            When someone you love needs critical medicine urgently, DialX
            connects your family, licensed pharmacies, and delivery — in one
            transparent flow.{" "}
            <span className="text-foreground">
              We coordinate. We don&apos;t sell medicines.
            </span>
          </p>

          <div className="mt-8 flex flex-col items-center justify-center gap-3 sm:flex-row">
            <Button
              size="lg"
              asChild
              onClick={() => trackCTAClick("Join Waitlist", "hero")}
            >
              <a href="#waitlist">
                Become an Early Supporter
                <ArrowRight className="size-4" aria-hidden="true" />
              </a>
            </Button>
            <Button
              variant="outline"
              size="lg"
              asChild
              onClick={() => trackCTAClick("See How It Works", "hero")}
            >
              <a href="#solution">See How It Works</a>
            </Button>
          </div>

          <p className="mt-4 text-sm text-muted-foreground">
            Prefer chat?{" "}
            <a
              href={getWhatsAppJoinUrl()}
              target="_blank"
              rel="noopener noreferrer"
              className="font-medium text-primary underline-offset-4 hover:underline"
              onClick={() => trackCTAClick("Join on WhatsApp", "hero")}
            >
              Join on WhatsApp
            </a>{" "}
            — no form needed.
          </p>
        </motion.div>

        <motion.div
          initial={{ opacity: 0, y: 12 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.6, delay: 0.15 }}
          className="mx-auto mt-10 grid max-w-2xl grid-cols-3 gap-4"
        >
          {heroStats.map((stat) => (
            <div key={stat.label} className="glass-card rounded-xl p-4 text-center">
              <p className="text-2xl font-bold text-foreground">
                <CountUp value={stat.value} suffix={stat.suffix} />
              </p>
              <p className="mt-1 text-xs text-muted-foreground">{stat.label}</p>
            </div>
          ))}
        </motion.div>

        {/* 5-question answer strip */}
        <motion.div
          initial={{ opacity: 0, y: 16 }}
          animate={{ opacity: 1, y: 0 }}
          transition={{ duration: 0.6, delay: 0.25 }}
          className="mx-auto mt-14 max-w-4xl"
        >
          <p className="mb-4 text-center text-xs font-medium uppercase tracking-wider text-muted-foreground">
            The story in 30 seconds
          </p>
          <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-5">
            {narrativeAnswers.map((item) => (
              <div
                key={item.question}
                className="rounded-xl border border-border bg-card/80 p-4 text-left backdrop-blur-sm"
              >
                <p className="text-xs font-semibold text-primary">
                  {item.question}
                </p>
                <p className="mt-1.5 text-sm leading-snug text-foreground">
                  {item.answer}
                </p>
              </div>
            ))}
          </div>
        </motion.div>
      </div>
    </section>
  );
}
