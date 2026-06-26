"use client";

import { useEffect } from "react";
import { trackScrollDepth } from "@/lib/analytics";

const DEPTHS = [25, 50, 75, 100];

export function ScrollTracker() {
  useEffect(() => {
    const tracked = new Set<number>();

    const handleScroll = () => {
      const scrollTop = window.scrollY;
      const docHeight =
        document.documentElement.scrollHeight - window.innerHeight;
      const scrollPercent = Math.round((scrollTop / docHeight) * 100);

      DEPTHS.forEach((depth) => {
        if (scrollPercent >= depth && !tracked.has(depth)) {
          tracked.add(depth);
          trackScrollDepth(depth);
        }
      });
    };

    window.addEventListener("scroll", handleScroll, { passive: true });
    return () => window.removeEventListener("scroll", handleScroll);
  }, []);

  return null;
}
