interface FormSectionPlaceholderProps {
  id: string;
  className?: string;
}

export function FormSectionPlaceholder({
  id,
  className = "section-padding",
}: FormSectionPlaceholderProps) {
  return (
    <section id={id} className={className} aria-hidden="true">
      <div className="container-narrow mx-auto max-w-4xl animate-pulse space-y-6 px-4">
        <div className="mx-auto h-8 w-48 rounded-lg bg-muted" />
        <div className="mx-auto h-4 w-96 max-w-full rounded bg-muted" />
        <div className="h-64 rounded-2xl bg-muted/60" />
      </div>
    </section>
  );
}
