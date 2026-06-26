"use client";

import { useState } from "react";
import { useForm, Controller } from "react-hook-form";
import { zodResolver } from "@hookform/resolvers/zod";
import { Loader2, CheckCircle2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { SectionHeader } from "@/components/shared/section-header";
import { WhatsAppButton } from "@/components/shared/whatsapp-button";
import { orderEstimates } from "@/lib/constants";
import { pharmacySchema, type PharmacyFormData } from "@/lib/validations";
import { trackFormStart, trackFormSubmit } from "@/lib/analytics";

export function PharmacyPartnerSection() {
  const [isSuccess, setIsSuccess] = useState(false);
  const [serverError, setServerError] = useState<string | null>(null);

  const {
    register,
    handleSubmit,
    control,
    reset,
    formState: { errors, isSubmitting },
  } = useForm<PharmacyFormData>({
    resolver: zodResolver(pharmacySchema),
  });

  const onSubmit = async (data: PharmacyFormData) => {
    setServerError(null);

    try {
      const response = await fetch("/api/pharmacy", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(data),
      });

      const result = await response.json();

      if (!response.ok) {
        throw new Error(result.error || "Something went wrong");
      }

      trackFormSubmit("pharmacy");
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
      <section id="pharmacy-partner" className="section-padding">
        <div className="container-narrow max-w-xl text-center">
          <div className="mx-auto mb-6 flex size-16 items-center justify-center rounded-full bg-trust/10">
            <CheckCircle2 className="size-8 text-trust" aria-hidden="true" />
          </div>
          <h2 className="text-2xl font-semibold">Application received!</h2>
          <p className="mt-3 text-muted-foreground">
            Our partnerships team will review your application and reach out
            within 3-5 business days.
          </p>
          <Button
            className="mt-6"
            variant="outline"
            onClick={() => setIsSuccess(false)}
          >
            Submit another application
          </Button>
        </div>
      </section>
    );
  }

  return (
    <section
      id="pharmacy-partner"
      className="section-padding"
      aria-labelledby="pharmacy-heading"
    >
      <div className="container-narrow max-w-2xl">
        <SectionHeader
          eyebrow="Pharmacy Partners"
          title="Partner with DialX"
          description="Join our network of licensed pharmacies and help families access emergency medicines faster."
        />

        <form
          onSubmit={handleSubmit(onSubmit)}
          onFocus={() => trackFormStart("pharmacy")}
          className="space-y-5 rounded-2xl border border-border bg-card p-6 shadow-sm sm:p-8"
          noValidate
        >
          <div className="grid gap-5 sm:grid-cols-2">
            <div className="space-y-2">
              <Label htmlFor="pharmacyName">Pharmacy Name *</Label>
              <Input
                id="pharmacyName"
                placeholder="Your pharmacy name"
                aria-invalid={!!errors.pharmacyName}
                {...register("pharmacyName")}
              />
              {errors.pharmacyName && (
                <p className="text-sm text-destructive" role="alert">
                  {errors.pharmacyName.message}
                </p>
              )}
            </div>

            <div className="space-y-2">
              <Label htmlFor="ownerName">Owner Name *</Label>
              <Input
                id="ownerName"
                placeholder="Owner / Manager name"
                aria-invalid={!!errors.ownerName}
                {...register("ownerName")}
              />
              {errors.ownerName && (
                <p className="text-sm text-destructive" role="alert">
                  {errors.ownerName.message}
                </p>
              )}
            </div>
          </div>

          <div className="grid gap-5 sm:grid-cols-2">
            <div className="space-y-2">
              <Label htmlFor="gstin">GSTIN (Optional)</Label>
              <Input
                id="gstin"
                placeholder="22AAAAA0000A1Z5"
                aria-invalid={!!errors.gstin}
                {...register("gstin")}
              />
              {errors.gstin && (
                <p className="text-sm text-destructive" role="alert">
                  {errors.gstin.message}
                </p>
              )}
            </div>

            <div className="space-y-2">
              <Label htmlFor="pharmacyCity">City *</Label>
              <Input
                id="pharmacyCity"
                placeholder="Pharmacy city"
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

          <div className="grid gap-5 sm:grid-cols-2">
            <div className="space-y-2">
              <Label htmlFor="contactNumber">Contact Number *</Label>
              <Input
                id="contactNumber"
                type="tel"
                placeholder="10-digit mobile number"
                aria-invalid={!!errors.contactNumber}
                {...register("contactNumber")}
              />
              {errors.contactNumber && (
                <p className="text-sm text-destructive" role="alert">
                  {errors.contactNumber.message}
                </p>
              )}
            </div>

            <div className="space-y-2">
              <Label htmlFor="pharmacyEmail">Email *</Label>
              <Input
                id="pharmacyEmail"
                type="email"
                placeholder="pharmacy@example.com"
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

          <div className="space-y-2">
            <Label htmlFor="dailyOrderEstimate">Daily Order Estimate *</Label>
            <Controller
              name="dailyOrderEstimate"
              control={control}
              render={({ field }) => (
                <Select onValueChange={field.onChange} value={field.value}>
                  <SelectTrigger
                    id="dailyOrderEstimate"
                    aria-invalid={!!errors.dailyOrderEstimate}
                  >
                    <SelectValue placeholder="Select daily order volume" />
                  </SelectTrigger>
                  <SelectContent>
                    {orderEstimates.map((estimate) => (
                      <SelectItem key={estimate} value={estimate}>
                        {estimate}
                      </SelectItem>
                    ))}
                  </SelectContent>
                </Select>
              )}
            />
            {errors.dailyOrderEstimate && (
              <p className="text-sm text-destructive" role="alert">
                {errors.dailyOrderEstimate.message}
              </p>
            )}
          </div>

          {serverError && (
            <p className="rounded-lg bg-destructive/10 p-3 text-sm text-destructive" role="alert">
              {serverError}
            </p>
          )}

          <Button
            type="submit"
            size="lg"
            variant="trust"
            className="w-full"
            disabled={isSubmitting}
          >
            {isSubmitting ? (
              <>
                <Loader2 className="size-4 animate-spin" aria-hidden="true" />
                Submitting...
              </>
            ) : (
              "Apply for Partnership"
            )}
          </Button>

          <div className="relative flex items-center gap-4 pt-2">
            <div className="h-px flex-1 bg-border" aria-hidden="true" />
            <span className="text-xs text-muted-foreground">or</span>
            <div className="h-px flex-1 bg-border" aria-hidden="true" />
          </div>

          <div className="rounded-xl border border-[#25D366]/20 bg-[#25D366]/5 p-4 text-center">
            <p className="text-sm text-muted-foreground">
              Prefer a quick chat? Reach our partnerships team on WhatsApp.
            </p>
            <WhatsAppButton
              location="pharmacy_section"
              label="Chat on WhatsApp"
              message="Hi DialX! I'm interested in becoming a pharmacy partner."
              className="mt-3"
            />
          </div>
        </form>
      </div>
    </section>
  );
}
