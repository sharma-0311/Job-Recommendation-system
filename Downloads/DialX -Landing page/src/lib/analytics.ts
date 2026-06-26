declare global {
  interface Window {
    gtag?: (...args: unknown[]) => void;
    dataLayer?: unknown[];
    clarity?: (...args: unknown[]) => void;
  }
}

export type AnalyticsEvent =
  | "cta_click"
  | "form_start"
  | "form_submit"
  | "form_abandon"
  | "scroll_depth"
  | "waitlist_conversion"
  | "pharmacy_conversion";

interface EventParams {
  event_category?: string;
  event_label?: string;
  value?: number;
  [key: string]: unknown;
}

export function trackEvent(event: AnalyticsEvent, params?: EventParams) {
  if (typeof window === "undefined") return;

  if (window.gtag) {
    window.gtag("event", event, params);
  }

  if (window.dataLayer) {
    window.dataLayer.push({
      event,
      ...params,
    });
  }
}

export function trackCTAClick(label: string, location: string) {
  trackEvent("cta_click", {
    event_category: "engagement",
    event_label: label,
    cta_location: location,
  });
}

export function trackFormStart(formName: string) {
  trackEvent("form_start", {
    event_category: "form",
    event_label: formName,
  });
}

export function trackFormSubmit(formName: string) {
  trackEvent("form_submit", {
    event_category: "form",
    event_label: formName,
  });
}

export function trackScrollDepth(depth: number) {
  trackEvent("scroll_depth", {
    event_category: "engagement",
    value: depth,
  });
}
