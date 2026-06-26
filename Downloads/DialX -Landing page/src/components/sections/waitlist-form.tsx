"use client";

import { useState } from "react";
import { useForm, Controller } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { Loader2, CheckCircle2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { WhatsAppButton } from "@/components/shared/whatsapp-button";
import { Input } from "@/components/ui/input";
import { Textarea } from "@/components/ui/textarea";
import { Label } from "@/components/ui/label";
import { Checkbox } from "@/components/ui/checkbox";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { SectionHeader } from "@/components/shared/section-header";
import { WhatsAppCommunityCard } from "@/components/shared/whatsapp-button";
import { userTypes, waitlistBenefits } from "@/lib/constants";
import { waitlistSchema, type WaitlistFormData } from "@/lib/validations";
import { trackFormStart, trackFormSubmit } from "@/lib/analytics";

export function WaitlistSection() {
  const [isSuccess, setIsSuccess] = useState(false);
  const [serverError, setServerError] = useState<string | null>(null);

  const {
    register,
    handleSubmit,
    control,
    reset,
    formState: { errors, isSubmitting },
  } = useForm<WaitlistFormData>({
    resolver: zodResolver(waitlistSchema),
    defaultValues: {
      consent: undefined,
    },
  });

  const onSubmit = async (data: WaitlistFormData) => {
    setServerError(null);

    try {
      const response = await fetch("/api/waitlist", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(data),
      });

      const result = await response.json();

      if (!response.ok) {
        throw new Error(result.error || "Something went wrong");
      }

      trackFormSubmit("waitlist");
      setIsSuccess(true);
      reset();
    } catch (error) {
      setServerError(
        error instanceof Error ? error.message : "Failed to submit. Please try again."
      );
    }
  };

  if (isSuccess) {
    return (
      <section id="waitlist" className="section-padding gradient-subtle">
        <div className="container-narrow max-w-xl text-center">
          <div className="mx-auto mb-6 flex size-16 items-center justify-center rounded-full bg-trust/10">
            <CheckCircle2 className="size-8 text-trust" aria-hidden="true" />
          </div>
          <h2 className="text-2xl font-semibold">You&apos;re on the list!</h2>
          <p className="mt-3 text-muted-foreground">
            Thank you for joining the DialX waitlist. We&apos;ll send you
            priority access updates as we approach launch.
          </p>
          <div className="mt-6 flex flex-col items-center gap-3">
            <WhatsAppButton location="waitlist_success" label="Also join on WhatsApp" />
            <Button
              variant="outline"
              onClick={() => setIsSuccess(false)}
            >
              Submit another response
            </Button>
          </div>
        </div>
      </section>
    );
  }

  return (
    <section
      id="waitlist"
      className="section-padding gradient-subtle"
      aria-labelledby="waitlist-heading"
    >
      <div className="container-narrow max-w-4xl">
        <SectionHeader
          eyebrow="Why Join Today"
          title="Become an early supporter"
          description="We're launching in 2–3 months. The families and pharmacies who join now will shape where we go first — and get priority access when we do."
        />

        {/* Why sign up today */}
        <div className="mb-10 grid gap-4 sm:grid-cols-2 lg:grid-cols-3">
          {waitlistBenefits.map((reason) => (
            <div
              key={reason.title}
              className="glass-card rounded-xl p-5"
            >
              <p className="font-semibold text-foreground">{reason.title}</p>
              <p className="mt-1 text-sm leading-relaxed text-muted-foreground">
                {reason.description}
              </p>
            </div>
          ))}
        </div>

        <div className="grid gap-8 lg:grid-cols-2 lg:items-start">
        <form
          onSubmit={handleSubmit(onSubmit)}
          onFocus={() => trackFormStart("waitlist")}
          className="space-y-5 rounded-2xl border border-border bg-card p-6 shadow-sm sm:p-8"
          noValidate
        >
          <div className="grid gap-5 sm:grid-cols-2">
            <div className="space-y-2">
              <Label htmlFor="name">Full Name *</Label>
              <Input
                id="name"
                placeholder="Your full name"
                aria-invalid={!!errors.name}
                {...register("name")}
              />
              {errors.name && (
                <p className="text-sm text-destructive" role="alert">
                  {errors.name.message}
                </p>
              )}
            </div>

            <div className="space-y-2">
              <Label htmlFor="email">Email *</Label>
              <Input
                id="email"
                type="email"
                placeholder="you@example.com"
                aria-invalid={!!errors.email}
                {...register("email")}
              />
              {errors.email && (
                <p className="text-sm text-destructive" role="alert">
                  {errors.email.message}
                </p>
              )}
            </div>
          </div>

          <div className="grid gap-5 sm:grid-cols-2">
            <div className="space-y-2">
              <Label htmlFor="phone">Phone *</Label>
              <Input
                id="phone"
                type="tel"
                placeholder="10-digit mobile number"
                aria-invalid={!!errors.phone}
                {...register("phone")}
              />
              {errors.phone && (
                <p className="text-sm text-destructive" role="alert">
                  {errors.phone.message}
                </p>
              )}
            </div>

            <div className="space-y-2">
              <Label htmlFor="city">City *</Label>
              <Input
                id="city"
                placeholder="Your city"
                aria-invalid={!!errors.city}
                {...register("city")}
              />
              {errors.city && (
                <p className="text-sm text-destructive" role="alert">
                  {errors.city.message}
                </p>
              )}
            </div>
          </div>

          <div className="space-y-2">
            <Label htmlFor="userType">I am a *</Label>
            <Controller
              name="userType"
              control={control}
              render={({ field }) => (
                <Select onValueChange={field.onChange} value={field.value}>
                  <SelectTrigger id="userType" aria-invalid={!!errors.userType}>
                    <SelectValue placeholder="Select user type" />
                  </SelectTrigger>
                  <SelectContent>
                    {userTypes.map((type) => (
                      <SelectItem key={type} value={type}>
                        {type}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              )}
            />
            {errors.userType && (
              <p className="text-sm text-destructive" role="alert">
                {errors.userType.message}
              </p>
            )}
          </div>

          <div className="space-y-2">
            <Label htmlFor="biggestChallenge">
              What&apos;s your biggest challenge with emergency medicines? *
            </Label>
            <Textarea
              id="biggestChallenge"
              placeholder="Tell us about your experience..."
              rows={4}
              aria-invalid={!!errors.biggestChallenge}
              {...register("biggestChallenge")}
            />
            {errors.biggestChallenge && (
              <p className="text-sm text-destructive" role="alert">
                {errors.biggestChallenge.message}
              </p>
            )}
          </div>

          <div className="flex items-start gap-3">
            <Controller
              name="consent"
              control={control}
              render={({ field }) => (
                <Checkbox
                  id="consent"
                  checked={field.value === true}
                  onCheckedChange={(checked) =>
                    field.onChange(checked === true)
                  }
                  aria-invalid={!!errors.consent}
                />
              )}
            />
            <Label htmlFor="consent" className="text-sm leading-relaxed font-normal">
              I agree to receive updates about DialX and understand my data will
              be handled per the{" "}
              <a href="/privacy" className="text-primary underline">
                Privacy Policy
              </a>
              . *
            </Label>
          </div>
          {errors.consent && (
            <p className="text-sm text-destructive" role="alert">
              {errors.consent.message}
            </p>
          )}

          {serverError && (
            <p className="rounded-lg bg-destructive/10 p-3 text-sm text-destructive" role="alert">
              {serverError}
            </p>
          )}

          <Button
            type="submit"
            size="lg"
            className="w-full"
            disabled={isSubmitting}
          >
            {isSubmitting ? (
              <>
                <Loader2 className="size-4 animate-spin" aria-hidden="true" />
                Joining...
              </>
            ) : (
              "Join as Early Supporter"
            )}
          </Button>
        </form>

        <div className="flex flex-col gap-6">
          <div className="flex items-center gap-4 lg:flex-col lg:gap-2">
            <div className="h-px flex-1 bg-border lg:hidden" aria-hidden="true" />
            <span className="shrink-0 text-sm font-medium text-muted-foreground">
              or skip the form
            </span>
            <div className="h-px flex-1 bg-border lg:hidden" aria-hidden="true" />
          </div>
          <WhatsAppCommunityCard location="waitlist_section" />
        </div>
        </div>
      </div>
    </section>
  );
}
