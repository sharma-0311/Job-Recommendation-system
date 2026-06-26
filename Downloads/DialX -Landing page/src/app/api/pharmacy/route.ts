import { NextResponse } from "next/server";
import { pharmacySchema } from "@/lib/validations";
import { getSupabaseAdmin, isSupabaseConfigured } from "@/lib/supabase";
import { isResendConfigured, sendPharmacyConfirmation } from "@/lib/resend";

export async function POST(request: Request) {
  try {
    const body = await request.json();
    const parsed = pharmacySchema.safeParse(body);

    if (!parsed.success) {
      return NextResponse.json(
        { error: parsed.error.errors[0]?.message || "Invalid input" },
        { status: 400 }
      );
    }

    const data = parsed.data;

    if (!isSupabaseConfigured()) {
      console.warn("Supabase not configured — logging pharmacy submission:", data.email);
      return NextResponse.json({
        success: true,
        message: "Pharmacy application received (dev mode)",
      });
    }

    const supabase = getSupabaseAdmin();

    const { error: dbError } = await supabase.from("pharmacy_partners").insert({
      pharmacy_name: data.pharmacyName,
      owner_name: data.ownerName,
      gstin: data.gstin || null,
      city: data.city,
      contact_number: data.contactNumber,
      email: data.email,
      daily_order_estimate: data.dailyOrderEstimate,
    });

    if (dbError) {
      if (dbError.code === "23505") {
        return NextResponse.json(
          { error: "This email has already submitted an application" },
          { status: 409 }
        );
      }
      console.error("Supabase error:", dbError);
      return NextResponse.json(
        { error: "Failed to save application" },
        { status: 500 }
      );
    }

    if (isResendConfigured()) {
      try {
        await sendPharmacyConfirmation(data.email, data.pharmacyName);
      } catch (emailError) {
        console.error("Email error:", emailError);
      }
    }

    return NextResponse.json({
      success: true,
      message: "Pharmacy application submitted successfully",
    });
  } catch (error) {
    console.error("Pharmacy API error:", error);
    return NextResponse.json(
      { error: "Internal server error" },
      { status: 500 }
    );
  }
}
