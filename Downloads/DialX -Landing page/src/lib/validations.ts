import { z } from "zod";
import { orderEstimates, userTypes } from "./constants";

export const waitlistSchema = z.object({
  name: z
    .string()
    .min(2, "Name must be at least 2 characters")
    .max(100, "Name is too long"),
  email: z.string().email("Please enter a valid email address"),
  phone: z
    .string()
    .regex(/^[6-9]\d{9}$/, "Please enter a valid 10-digit Indian mobile number"),
  city: z
    .string()
    .min(2, "City is required")
    .max(100, "City name is too long"),
  userType: z.enum(userTypes as [string, ...string[]], {
    errorMap: () => ({ message: "Please select a user type" }),
  }),
  biggestChallenge: z
    .string()
    .min(10, "Please share at least 10 characters")
    .max(1000, "Response is too long"),
  consent: z.literal(true, {
    errorMap: () => ({ message: "You must agree to continue" }),
  }),
});

export type WaitlistFormData = z.infer<typeof waitlistSchema>;

export const pharmacySchema = z.object({
  pharmacyName: z
    .string()
    .min(2, "Pharmacy name is required")
    .max(200, "Name is too long"),
  ownerName: z
    .string()
    .min(2, "Owner name is required")
    .max(100, "Name is too long"),
  gstin: z
    .string()
    .optional()
    .refine(
      (val) =>
        !val ||
        /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$/.test(val),
      { message: "Please enter a valid GSTIN" }
    ),
  city: z
    .string()
    .min(2, "City is required")
    .max(100, "City name is too long"),
  contactNumber: z
    .string()
    .regex(/^[6-9]\d{9}$/, "Please enter a valid 10-digit mobile number"),
  email: z.string().email("Please enter a valid email address"),
  dailyOrderEstimate: z.enum(orderEstimates as [string, ...string[]], {
    errorMap: () => ({ message: "Please select an estimate" }),
  }),
});

export type PharmacyFormData = z.infer<typeof pharmacySchema>;
