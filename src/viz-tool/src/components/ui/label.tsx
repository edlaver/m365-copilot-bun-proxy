import type { LabelHTMLAttributes } from "react"

import { cn } from "@/lib/utils"

function Label({
  className,
  ...props
}: LabelHTMLAttributes<HTMLLabelElement>) {
  return (
    <label
      data-slot="label"
      className={cn(
        "text-[11px] font-medium uppercase tracking-[0.18em] text-muted-foreground",
        className
      )}
      {...props}
    />
  )
}

export { Label }
