"use client";

import {
  Accordion,
  AccordionContent,
  AccordionItem,
  AccordionTrigger,
} from "@/components/ui/accordion";
import { SectionHeader } from "@/components/shared/section-header";
import { faqs } from "@/lib/constants";

export function FAQSection() {
  return (
    <section
      id="faq"
      className="section-padding"
      aria-labelledby="faq-heading"
    >
      <div className="container-narrow max-w-3xl">
        <SectionHeader
          eyebrow="Still Have Questions?"
          title="Straight answers, no jargon"
          description="The questions investors, families, and pharmacy partners ask us most."
        />

        <Accordion type="single" collapsible className="w-full">
          {faqs.map((faq, index) => (
            <AccordionItem key={index} value={`item-${index}`}>
              <AccordionTrigger className="text-left text-base">
                {faq.question}
              </AccordionTrigger>
              <AccordionContent>{faq.answer}</AccordionContent>
            </AccordionItem>
          ))}
        </Accordion>
      </div>
    </section>
  );
}
