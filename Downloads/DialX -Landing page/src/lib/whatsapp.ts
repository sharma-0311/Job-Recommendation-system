export const whatsappConfig = {
  /** WhatsApp Community or Group invite link — preferred for community join */
  communityUrl:
    process.env.NEXT_PUBLIC_WHATSAPP_COMMUNITY_URL ||
    "https://chat.whatsapp.com/your-invite-link",
  /** Business number with country code, e.g. 919876543210 */
  number: process.env.NEXT_PUBLIC_WHATSAPP_NUMBER || "916398407954",
  defaultMessage:
    "Hi DialX! I'd like to join the early access community and stay updated on launch.",
  communityLabel: "Join WhatsApp Community",
  communityDescription:
    "No forms needed. Get launch updates, product previews, and connect with other early supporters.",
};

export function getWhatsAppJoinUrl(message?: string): string {
  const { communityUrl, number, defaultMessage } = whatsappConfig;

  if (communityUrl && !communityUrl.includes("your-invite-link")) {
    return communityUrl;
  }

  const text = encodeURIComponent(message || defaultMessage);
  const cleanNumber = number.replace(/\D/g, "");
  return `https://wa.me/${cleanNumber}?text=${text}`;
}

export function isWhatsAppConfigured(): boolean {
  const { communityUrl, number } = whatsappConfig;
  return (
    (Boolean(communityUrl) && !communityUrl.includes("your-invite-link")) ||
    Boolean(number.replace(/\D/g, ""))
  );
}
