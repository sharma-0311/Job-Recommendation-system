import { Resend } from "resend";

const resendApiKey = process.env.RESEND_API_KEY;
const fromEmail = process.env.RESEND_FROM_EMAIL || "DialX <noreply@dialx.in>";

export function getResend() {
  if (!resendApiKey) {
    throw new Error("RESEND_API_KEY is not configured");
  }
  return new Resend(resendApiKey);
}

export function isResendConfigured(): boolean {
  return Boolean(process.env.RESEND_API_KEY);
}

export async function sendWaitlistConfirmation(email: string, name: string) {
  const resend = getResend();

  await resend.emails.send({
    from: fromEmail,
    to: email,
    subject: "You're on the DialX waitlist",
    html: `
      <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; max-width: 560px; margin: 0 auto; padding: 40px 20px;">
        <h1 style="color: #1a365d; font-size: 24px; margin-bottom: 16px;">Welcome to DialX, ${name}!</h1>
        <p style="color: #4a5568; line-height: 1.6; margin-bottom: 16px;">
          Thank you for joining our waitlist. You're now among the first to experience India's emergency healthcare coordination platform.
        </p>
        <p style="color: #4a5568; line-height: 1.6; margin-bottom: 16px;">
          We'll keep you updated on our launch progress and give you priority access when we go live.
        </p>
        <p style="color: #718096; font-size: 14px; margin-top: 32px;">
          — The DialX Team
        </p>
      </div>
    `,
  });
}

export async function sendPharmacyConfirmation(email: string, pharmacyName: string) {
  const resend = getResend();

  await resend.emails.send({
    from: fromEmail,
    to: email,
    subject: "DialX Pharmacy Partnership Application Received",
    html: `
      <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif; max-width: 560px; margin: 0 auto; padding: 40px 20px;">
        <h1 style="color: #1a365d; font-size: 24px; margin-bottom: 16px;">Application Received</h1>
        <p style="color: #4a5568; line-height: 1.6; margin-bottom: 16px;">
          Thank you for your interest in partnering with DialX, ${pharmacyName}.
        </p>
        <p style="color: #4a5568; line-height: 1.6; margin-bottom: 16px;">
          Our partnerships team will review your application and reach out within 3-5 business days.
        </p>
        <p style="color: #718096; font-size: 14px; margin-top: 32px;">
          — The DialX Partnerships Team
        </p>
      </div>
    `,
  });
}
