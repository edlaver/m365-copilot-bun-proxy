import type { HTMLAttributes } from "react"

import { cn } from "@/lib/utils"

function Card({ className, ...props }: HTMLAttributes<HTMLDivElement>) {
  return (
    <div
      data-slot="card"
      className={cn(
        "rounded-2xl border border-border/70 bg-card/90 shadow-sm backdrop-blur",
        className
      )}
      {...props}
    />
  )
}

function CardHeader({ className, ...props }: HTMLAttributes<HTMLDivElement>) {
  return (
    <div
      data-slot="card-header"
      className={cn("flex items-start justify-between gap-4 border-b border-border/60 px-4 py-3", className)}
      {...props}
    />
  )
}

function CardTitle({ className, ...props }: HTMLAttributes<HTMLHeadingElement>) {
  return (
    <h2
      data-slot="card-title"
      className={cn("text-sm font-semibold tracking-tight", className)}
      {...props}
    />
  )
}

function CardDescription({
  className,
  ...props
}: HTMLAttributes<HTMLParagraphElement>) {
  return (
    <p
      data-slot="card-description"
      className={cn("text-xs text-muted-foreground", className)}
      {...props}
    />
  )
}

function CardContent({ className, ...props }: HTMLAttributes<HTMLDivElement>) {
  return (
    <div
      data-slot="card-content"
      className={cn("min-h-0 flex-1 p-0", className)}
      {...props}
    />
  )
}

export { Card, CardContent, CardDescription, CardHeader, CardTitle }
