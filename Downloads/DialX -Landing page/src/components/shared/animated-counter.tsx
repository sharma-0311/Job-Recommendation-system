import { cn } from "@/lib/utils";

interface AnimatedCounterProps {
  value: number;
  suffix?: string;
  className?: string;
}

export function AnimatedCounter({
  value,
  suffix = "",
  className,
}: AnimatedCounterProps) {
  return (
    <span className={cn("tabular-nums", className)}>
      {value.toLocaleString("en-IN")}
      {suffix}
    </span>
  );
}
