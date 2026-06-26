"use client";

import { Button } from "@/components/ui/button";
import { WhatsAppIcon } from "@/components/shared/whatsapp-icon";
import { getWhatsAppJoinUrl, whatsappConfig } from "@/lib/whatsapp";
import { trackCTAClick } from "@/lib/analytics";
import { cn } from "@/lib/utils";

interface WhatsAppButtonProps {
  location: string;
  label?: string;
  message?: string;
  variant?: "default" | "outline" | "ghost" | "whatsapp";
  size?: "default" | "sm" | "lg";
  className?: string;
  showIcon?: boolean;
  fullWidth?: boolean;
  iconOnly?: boolean;
  "aria-label"?: string;
}

export function WhatsAppButton({
  location,
  label = whatsappConfig.communityLabel,
  message,
  variant = "whatsapp",
  size = "default",
  className,
  showIcon = true,
  fullWidth = false,
  iconOnly = false,
  "aria-label": ariaLabel,
}: WhatsAppButtonProps) {
  const href = getWhatsAppJoinUrl(message);
  const accessibleLabel =
    ariaLabel || (iconOnly ? "Join on WhatsApp" : `${label} on WhatsApp (opens in new tab)`);

  const handleClick = () => {
    trackCTAClick(label, location);
  };

  if (variant === "whatsapp") {
    return (
      <a
        href={href}
        target="_blank"
        rel="noopener noreferrer"
        onClick={handleClick}
        className={cn(
          "inline-flex items-center justify-center gap-2 rounded-lg bg-[#25D366] px-5 py-2.5 text-sm font-medium text-white shadow-sm transition-all hover:bg-[#20BD5A] active:scale-[0.98] focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[#25D366] focus-visible:ring-offset-2",
          size === "lg" && "h-12 px-8 text-base",
          size === "sm" && "h-8 px-3 text-xs",
          iconOnly && "size-8 gap-0 p-0",
          fullWidth && "w-full",
          className
        )}
        aria-label={accessibleLabel}
      >
        {showIcon && <WhatsAppIcon />}
        {!iconOnly && label}
      </a>
    );
  }

  return (
    <Button
      variant={variant}
      size={size}
      className={cn(fullWidth && "w-full", className)}
      asChild
    >
      <a
        href={href}
        target="_blank"
        rel="noopener noreferrer"
        onClick={handleClick}
        aria-label={`${label} on WhatsApp (opens in new tab)`}
      >
        {showIcon && <WhatsAppIcon />}
        {label}
      </a>
    </Button>
  );
}

interface WhatsAppCommunityCardProps {
  location: string;
  className?: string;
}

export function WhatsAppCommunityCard({
  location,
  className,
}: WhatsAppCommunityCardProps) {
  return (
    <div
      className={cn(
        "rounded-2xl border border-[#25D366]/30 bg-[#25D366]/5 p-6 sm:p-8",
        className
      )}
    >
      <div className="flex items-start gap-4">
        <div className="flex size-12 shrink-0 items-center justify-center rounded-xl bg-[#25D366] text-white">
          <WhatsAppIcon className="size-6" />
        </div>
        <div className="flex-1">
          <h3 className="text-lg font-semibold text-foreground">
            Prefer WhatsApp? Join our community
          </h3>
          <p className="mt-2 text-sm text-muted-foreground leading-relaxed">
            {whatsappConfig.communityDescription}
          </p>
          <ul className="mt-4 space-y-2 text-sm text-muted-foreground">
            <li className="flex items-center gap-2">
              <span className="size-1.5 rounded-full bg-[#25D366]" />
              Launch updates &amp; early access alerts
            </li>
            <li className="flex items-center gap-2">
              <span className="size-1.5 rounded-full bg-[#25D366]" />
              Direct line to the founding team
            </li>
            <li className="flex items-center gap-2">
              <span className="size-1.5 rounded-full bg-[#25D366]" />
              No lengthy forms — join in one tap
            </li>
          </ul>
          <WhatsAppButton
            location={location}
            size="lg"
            className="mt-6"
            label="Join on WhatsApp"
          />
        </div>
      </div>
    </div>
  );
}
