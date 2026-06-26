import { NextResponse } from "next/server";
import { waitlistSchema } from "@/lib/validations";
import { getSupabaseAdmin, isSupabaseConfigured } from "@/lib/supabase";
import { isResendConfigured, sendWaitlistConfirmation } from "@/lib/resend";

export async function POST(request: Request) {
  try {
    const body = await request.json();
    const parsed = waitlistSchema.safeParse(body);

    if (!parsed.success) {
      return NextResponse.json(
        { error: parsed.error.errors[0]?.message || "Invalid input" },
        { status: 400 }
      );
    }

    const data = parsed.data;

    if (!isSupabaseConfigured()) {
      console.warn("Supabase not configured — logging waitlist submission:", data.email);
      return NextResponse.json({
        success: true,
        message: "Waitlist submission received (dev mode)",
      });
    }

    const supabase = getSupabaseAdmin();

    const { error: dbError } = await supabase.from("waitlist").insert({
      name: data.name,
      email: data.email,
      phone: data.phone,
      city: data.city,
      user_type: data.userType,
      biggest_challenge: data.biggestChallenge,
      consent: data.consent,
    });

    if (dbError) {
      if (dbError.code === "23505") {
        return NextResponse.json(
          { error: "This email is already on the waitlist" },
          { status: 409 }
        );
      }
      console.error("Supabase error:", dbError);
      return NextResponse.json(
        { error: "Failed to save submission" },
        { status: 500 }
      );
    }

    if (isResendConfigured()) {
      try {
        await sendWaitlistConfirmation(data.email, data.name);
      } catch (emailError) {
        console.error("Email error:", emailError);
      }
    }

    return NextResponse.json({
      success: true,
      message: "Successfully joined the waitlist",
    });
  } catch (error) {
    console.error("Waitlist API error:", error);
    return NextResponse.json(
      { error: "Internal server error" },
      { status: 500 }
    );
  }
}
