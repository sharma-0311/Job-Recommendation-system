import {
  generateFAQSchema,
  generateOrganizationSchema,
  generateWebSiteSchema,
} from "@/lib/seo";
import { HeroSection } from "@/components/sections/hero";
import { TrustRibbon } from "@/components/sections/trust-ribbon";
import { ProblemSection } from "@/components/sections/problem";
import { FragmentationSection } from "@/components/sections/fragmentation";
import { SolutionSection } from "@/components/sections/solution";
import { WhyDialXSection } from "@/components/sections/why-dialx";
import { ProductCapabilitiesSection } from "@/components/sections/product-capabilities";
import { HealthcareServicesSection } from "@/components/sections/healthcare-services";
import { SecurityComplianceSection } from "@/components/sections/security-compliance";
import { WhoWeServeSection } from "@/components/sections/who-we-serve";
import { PlatformPreviewSection } from "@/components/sections/platform-preview";
import { TechnologySection } from "@/components/sections/technology";
import { AdvisorySection } from "@/components/sections/advisory";
import { FounderVisionSection } from "@/components/sections/founder-vision";
import { RoadmapSection } from "@/components/sections/roadmap";
import { WaitlistSection } from "@/components/sections/waitlist-form";
import { PharmacyPartnerSection } from "@/components/sections/pharmacy-form";
import { FAQSection } from "@/components/sections/faq";
import { ClientOnly } from "@/components/shared/client-only";
import { FormSectionPlaceholder } from "@/components/shared/form-section-placeholder";

export default function HomePage() {
  const organizationSchema = generateOrganizationSchema();
  const websiteSchema = generateWebSiteSchema();
  const faqSchema = generateFAQSchema();

  return (
    <>
      <script
        type="application/ld+json"
        dangerouslySetInnerHTML={{
          __html: JSON.stringify(organizationSchema),
        }}
      />
      <script
        type="application/ld+json"
        dangerouslySetInnerHTML={{ __html: JSON.stringify(websiteSchema) }}
      />
      <script
        type="application/ld+json"
        dangerouslySetInnerHTML={{ __html: JSON.stringify(faqSchema) }}
      />

      <HeroSection />
      <TrustRibbon />
      <ProblemSection />
      <FragmentationSection />
      <SolutionSection />
      <WhyDialXSection />
      <ProductCapabilitiesSection />
      <HealthcareServicesSection />
      <SecurityComplianceSection />
      <WhoWeServeSection />
      <PlatformPreviewSection />
      <TechnologySection />
      <AdvisorySection />
      <FounderVisionSection />
      <RoadmapSection />
      <ClientOnly
        fallback={
          <FormSectionPlaceholder id="waitlist" className="section-padding gradient-subtle" />
        }
      >
        <WaitlistSection />
      </ClientOnly>
      <ClientOnly
        fallback={<FormSectionPlaceholder id="pharmacy-partner" />}
      >
        <PharmacyPartnerSection />
      </ClientOnly>
      <ClientOnly fallback={<FormSectionPlaceholder id="faq" />}>
        <FAQSection />
      </ClientOnly>
    </>
  );
}
