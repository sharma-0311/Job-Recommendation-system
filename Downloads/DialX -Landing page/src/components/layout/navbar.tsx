"use client";

import { useEffect, useState } from "react";
import Link from "next/link";
import { Menu, X } from "lucide-react";
import { Button } from "@/components/ui/button";
import { WhatsAppButton } from "@/components/shared/whatsapp-button";
import { navLinks, siteConfig } from "@/lib/constants";
import { trackCTAClick } from "@/lib/analytics";
import { cn } from "@/lib/utils";

export function Navbar() {
  const [isScrolled, setIsScrolled] = useState(false);
  const [isMobileOpen, setIsMobileOpen] = useState(false);

  useEffect(() => {
    const handleScroll = () => setIsScrolled(window.scrollY > 20);
    window.addEventListener("scroll", handleScroll, { passive: true });
    return () => window.removeEventListener("scroll", handleScroll);
  }, []);

  useEffect(() => {
    document.body.style.overflow = isMobileOpen ? "hidden" : "";
    return () => {
      document.body.style.overflow = "";
    };
  }, [isMobileOpen]);

  const closeMobile = () => setIsMobileOpen(false);

  const handleWaitlistClick = () => {
    trackCTAClick("Join Waitlist", "navbar");
    closeMobile();
  };

  return (
    <header
      className={cn(
        "fixed top-0 right-0 left-0 z-50 transition-all duration-300",
        isScrolled || isMobileOpen
          ? "border-b border-border/60 bg-background/95 shadow-sm backdrop-blur-md"
          : "bg-background/80 backdrop-blur-sm"
      )}
    >
      <div className="container-narrow px-4 sm:px-6 lg:px-8">
        {/* Desktop & tablet bar */}
        <nav
          className="grid h-16 grid-cols-[auto_1fr_auto] items-center gap-4 lg:gap-8"
          aria-label="Main navigation"
        >
          {/* Logo */}
          <Link
            href="/"
            className="flex shrink-0 items-center gap-2.5 font-semibold text-foreground"
            aria-label={`${siteConfig.name} home`}
            onClick={closeMobile}
          >
            <span className="flex size-9 items-center justify-center rounded-lg bg-primary text-sm font-bold text-primary-foreground">
              DX
            </span>
            <span className="text-lg tracking-tight">{siteConfig.name}</span>
          </Link>

          {/* Center nav — desktop only */}
          <ul className="hidden items-center justify-center gap-1 lg:flex xl:gap-2">
            {navLinks.map((link) => (
              <li key={link.href}>
                <a
                  href={link.href}
                  className="rounded-md px-2.5 py-2 text-sm text-muted-foreground transition-colors hover:bg-muted hover:text-foreground xl:px-3"
                >
                  {link.label}
                </a>
              </li>
            ))}
          </ul>

          {/* Right actions — desktop */}
          <div className="hidden shrink-0 items-center gap-2 lg:flex">
            <WhatsAppButton
              location="navbar"
              label="WhatsApp"
              size="sm"
              iconOnly
              className="xl:hidden"
            />
            <WhatsAppButton
              location="navbar"
              label="WhatsApp"
              size="sm"
              className="hidden xl:inline-flex"
            />
            <div className="mx-1 h-6 w-px bg-border" aria-hidden="true" />
            <Button variant="ghost" size="sm" className="hidden xl:inline-flex" asChild>
              <a href="#pharmacy-partner">Partner</a>
            </Button>
            <Button size="sm" asChild onClick={handleWaitlistClick}>
              <a href="#waitlist">Join Early</a>
            </Button>
          </div>

          {/* Mobile menu toggle */}
          <button
            type="button"
            className="col-start-3 inline-flex size-10 items-center justify-center rounded-lg border border-border bg-background text-foreground transition-colors hover:bg-muted lg:hidden"
            onClick={() => setIsMobileOpen(!isMobileOpen)}
            aria-expanded={isMobileOpen}
            aria-controls="mobile-nav"
            aria-label={isMobileOpen ? "Close menu" : "Open menu"}
          >
            {isMobileOpen ? (
              <X className="size-5" aria-hidden="true" />
            ) : (
              <Menu className="size-5" aria-hidden="true" />
            )}
          </button>
        </nav>
      </div>

      {/* Mobile menu overlay */}
      {isMobileOpen && (
        <>
          <button
            type="button"
            className="fixed inset-0 top-16 z-40 bg-foreground/20 backdrop-blur-sm lg:hidden"
            onClick={closeMobile}
            aria-label="Close menu"
          />
          <div
            id="mobile-nav"
            className="relative z-50 border-t border-border bg-background lg:hidden"
          >
            <div className="container-narrow max-h-[calc(100vh-4rem)] overflow-y-auto px-4 py-6 sm:px-6">
              <p className="mb-3 text-xs font-semibold uppercase tracking-wider text-muted-foreground">
                Navigate
              </p>
              <ul className="grid gap-1">
                {navLinks.map((link) => (
                  <li key={link.href}>
                    <a
                      href={link.href}
                      className="flex items-center rounded-lg px-3 py-3 text-base font-medium text-foreground transition-colors hover:bg-muted"
                      onClick={closeMobile}
                    >
                      {link.label}
                    </a>
                  </li>
                ))}
              </ul>

              <div className="my-6 h-px bg-border" aria-hidden="true" />

              <p className="mb-3 text-xs font-semibold uppercase tracking-wider text-muted-foreground">
                Get started
              </p>
              <div className="flex flex-col gap-3">
                <Button size="lg" className="w-full" asChild onClick={handleWaitlistClick}>
                  <a href="#waitlist">Join Waitlist</a>
                </Button>
                <Button variant="outline" size="lg" className="w-full" asChild onClick={closeMobile}>
                  <a href="#pharmacy-partner">Partner With Us</a>
                </Button>
                <WhatsAppButton
                  location="navbar_mobile"
                  label="Join on WhatsApp"
                  size="lg"
                  fullWidth
                />
              </div>
            </div>
          </div>
        </>
      )}
    </header>
  );
}
