import { useEffect, useState } from "react"
import { AlertCircle, LoaderCircle, RefreshCcw } from "lucide-react"

import { JsonPane } from "@/components/json-pane"
import { Button } from "@/components/ui/button"
import { Label } from "@/components/ui/label"
import {
  getFixtures,
  type RequestType,
  type TraceResponse,
  type TransformMode,
} from "@/lib/fixtures"
import { cn } from "@/lib/utils"

const requestTypeOptions: Array<{ label: string; value: RequestType }> = [
  { label: "chat/completions", value: "chat/completions" },
  { label: "responses", value: "responses" },
]

const transformModeOptions: Array<{ label: string; value: TransformMode }> = [
  { label: "simulated", value: "simulated" },
  { label: "mapped", value: "mapped" },
]

const emptyPaneText = JSON.stringify(
  {
    note: "Submit a request to populate this pane.",
  },
  null,
  2
)
const maxProxyRetryAttempts = 3

export function App() {
  const [transformMode, setTransformMode] = useState<TransformMode>("simulated")
  const [requestType, setRequestType] = useState<RequestType>("chat/completions")
  const [selectedFixtureId, setSelectedFixtureId] = useState("")
  const [pane1, setPane1] = useState("")
  const [pane2, setPane2] = useState(emptyPaneText)
  const [pane3, setPane3] = useState(emptyPaneText)
  const [pane4, setPane4] = useState(emptyPaneText)
  const [statusText, setStatusText] = useState("Ready")
  const [errorText, setErrorText] = useState<string | null>(null)
  const [isSubmitting, setIsSubmitting] = useState(false)

  const fixtures = getFixtures(requestType, transformMode)
  const selectedFixture =
    fixtures.find((fixture) => fixture.id === selectedFixtureId) ?? fixtures[0] ?? null

  useEffect(() => {
    if (!selectedFixture) {
      setSelectedFixtureId("")
      setPane1("{}")
      return
    }
    setSelectedFixtureId(selectedFixture.id)
    setPane1(selectedFixture.content)
  }, [selectedFixtureId, selectedFixture])

  async function handleSubmit() {
    setErrorText(null)
    setIsSubmitting(true)
    setStatusText("Submitting request")
    setPane2(emptyPaneText)
    setPane3(emptyPaneText)
    setPane4(emptyPaneText)

    let parsedBody: unknown
    try {
      parsedBody = JSON.parse(pane1)
    } catch (error) {
      setErrorText(`Pane 1 must contain valid JSON. ${String(error)}`)
      setStatusText("Invalid request JSON")
      setIsSubmitting(false)
      return
    }

    const traceId = crypto.randomUUID()
    const endpoint =
      requestType === "chat/completions"
        ? "/v1/chat/completions"
        : "/v1/responses"

    let responseText = ""
    let responseStatus = 0
    try {
      const response = await fetchWithProxyRetry(endpoint, {
        method: "POST",
        headers: {
          "content-type": "application/json",
          "x-m365-openai-transform-mode": transformMode,
          "x-m365-viz-trace-id": traceId,
        },
        body: JSON.stringify(parsedBody),
      })
      responseStatus = response.status
      responseText = await response.text()

      setStatusText(
        response.ok
          ? "Request completed, waiting for trace"
          : `Proxy returned ${response.status}, waiting for trace`
      )

      const trace = await waitForTrace(traceId)
      if (trace) {
        setPane2(formatJson(trace.pane2 ?? trace.error ?? { status: trace.status }))
        setPane3(formatJson(trace.pane3 ?? { note: "No transformed upstream request was captured." }))
        setPane4(formatJson(trace.pane4 ?? { note: "No upstream response was captured." }))
        setStatusText(
          `Trace ${trace.status} · transport ${trace.transport} · proxy ${responseStatus}`
        )
        if (trace.error) {
          setErrorText(formatInline(trace.error))
        }
        return
      }

      setPane2(formatJson(parseJsonOrWrap(responseText, responseStatus)))
      setErrorText("Trace data was not available before the timeout.")
      setStatusText(`Proxy ${responseStatus} · trace timeout`)
    } catch (error) {
      setPane2(formatJson(parseJsonOrWrap(responseText, responseStatus || 0)))
      setErrorText(`Request failed. ${String(error)}`)
      setStatusText("Request failed")
    } finally {
      setIsSubmitting(false)
    }
  }

  return (
    <div className="min-h-svh bg-[radial-gradient(circle_at_top_left,_rgba(88,124,255,0.14),_transparent_30%),radial-gradient(circle_at_top_right,_rgba(14,165,233,0.10),_transparent_28%),linear-gradient(180deg,_var(--color-background),color-mix(in_oklch,var(--color-background)_90%,var(--color-muted)_10%))]">
      <div className="mx-auto flex min-h-svh max-w-[1800px] flex-col gap-4 p-4 lg:p-6">
        <header className="rounded-2xl border border-border/70 bg-background/80 px-4 py-4 shadow-sm backdrop-blur">
          <div className="flex flex-col gap-4 xl:flex-row xl:items-end xl:justify-between">
            <div className="space-y-1">
              <p className="text-[11px] font-semibold uppercase tracking-[0.22em] text-muted-foreground">
                Proxy Viz Tool
              </p>
              <h1 className="text-xl font-semibold tracking-tight">
                Visualize proxy payload mapping across the full request lifecycle
              </h1>
              <p className="text-sm text-muted-foreground">
                Pane 1 is editable. Panes 2-4 are populated from the completed live trace.
              </p>
            </div>

            <div className="flex flex-col gap-3 xl:min-w-[840px]">
              <div className="grid gap-3 md:grid-cols-[180px_220px_minmax(0,1fr)_auto]">
                <ToolbarSelect
                  label="Transform"
                  value={transformMode}
                  onChange={(value) => setTransformMode(value as TransformMode)}
                  options={transformModeOptions}
                />
                <ToolbarSelect
                  label="Request Type"
                  value={requestType}
                  onChange={(value) => setRequestType(value as RequestType)}
                  options={requestTypeOptions}
                />
                <ToolbarSelect
                  label="Canned Request"
                  value={selectedFixtureId}
                  onChange={setSelectedFixtureId}
                  options={fixtures.map((fixture) => ({
                    label: fixture.label,
                    value: fixture.id,
                  }))}
                />
                <div className="flex items-end justify-end">
                  <Button
                    className="w-full md:w-auto"
                    size="lg"
                    onClick={() => void handleSubmit()}
                    disabled={isSubmitting}
                  >
                    {isSubmitting ? (
                      <LoaderCircle className="size-4 animate-spin" />
                    ) : (
                      <RefreshCcw className="size-4" />
                    )}
                    Submit
                  </Button>
                </div>
              </div>

              <div className="flex flex-wrap items-center gap-3 text-xs text-muted-foreground">
                <span className="rounded-full border border-border/70 bg-background/70 px-2.5 py-1">
                  {statusText}
                </span>
                {selectedFixture ? (
                  <span className="rounded-full border border-border/70 bg-background/70 px-2.5 py-1">
                    {selectedFixture.sourceFile}:{selectedFixture.sourceLine}
                  </span>
                ) : null}
                <span className="rounded-full border border-border/70 bg-background/70 px-2.5 py-1">
                  Press <kbd className="font-mono">d</kbd> to toggle theme
                </span>
              </div>

              {errorText ? (
                <div className="flex items-start gap-2 rounded-xl border border-destructive/30 bg-destructive/8 px-3 py-2 text-sm text-destructive">
                  <AlertCircle className="mt-0.5 size-4 shrink-0" />
                  <p>{errorText}</p>
                </div>
              ) : null}
            </div>
          </div>
        </header>

        <main className="grid min-h-0 flex-1 gap-4 lg:grid-cols-2 lg:grid-rows-2">
          <JsonPane
            title="1. Submitted Request"
            description="Editable OpenAI request body sent by the tool."
            value={pane1}
            onChange={setPane1}
          />
          <JsonPane
            title="2. Proxy Output"
            description="Final transformed response returned by the proxy after buffering."
            value={pane2}
            readOnly
          />
          <JsonPane
            title="3. Upstream Request"
            description="Payload produced by the proxy before it is sent upstream."
            value={pane3}
            readOnly
          />
          <JsonPane
            title="4. Upstream Response"
            description="Buffered raw upstream response before proxy transformation."
            value={pane4}
            readOnly
          />
        </main>
      </div>
    </div>
  )
}

type ToolbarOption = {
  label: string
  value: string
}

type ToolbarSelectProps = {
  label: string
  value: string
  onChange: (value: string) => void
  options: ToolbarOption[]
}

function ToolbarSelect({
  label,
  value,
  onChange,
  options,
}: ToolbarSelectProps) {
  return (
    <div className="space-y-1.5">
      <Label>{label}</Label>
      <select
        className={cn(
          "h-11 w-full rounded-xl border border-border/70 bg-background/85 px-3 text-sm shadow-sm outline-none transition focus:border-ring focus:ring-3 focus:ring-ring/25"
        )}
        value={value}
        onChange={(event) => onChange(event.target.value)}
      >
        {options.map((option) => (
          <option key={option.value} value={option.value}>
            {option.label}
          </option>
        ))}
      </select>
    </div>
  )
}

function formatJson(value: unknown): string {
  return JSON.stringify(value ?? null, null, 2)
}

function parseJsonOrWrap(rawText: string, status: number): unknown {
  if (!rawText.trim()) {
    return {
      status,
      note: "Response body was empty.",
    }
  }
  try {
    return JSON.parse(rawText) as unknown
  } catch {
    return {
      status,
      rawText,
    }
  }
}

function formatInline(value: unknown): string {
  if (typeof value === "string") {
    return value
  }
  return JSON.stringify(value)
}

async function waitForTrace(traceId: string): Promise<TraceResponse | null> {
  for (let attempt = 0; attempt < 24; attempt += 1) {
    let response: Response
    try {
      response = await fetchWithProxyRetry(`/__viz/traces/${traceId}`)
    } catch {
      return null
    }
    if (response.ok) {
      const trace = (await response.json()) as TraceResponse
      if (trace.status !== "pending" || attempt === 23) {
        return trace
      }
    }
    await delay(250)
  }
  return null
}

function delay(timeoutMs: number): Promise<void> {
  return new Promise((resolve) => {
    window.setTimeout(resolve, timeoutMs)
  })
}

async function fetchWithProxyRetry(
  input: RequestInfo | URL,
  init?: RequestInit
): Promise<Response> {
  let lastError: unknown
  for (let attempt = 1; attempt <= maxProxyRetryAttempts; attempt += 1) {
    try {
      return await fetch(input, init)
    } catch (error) {
      lastError = error
      if (attempt === maxProxyRetryAttempts) {
        break
      }
      await delay(250)
    }
  }
  throw lastError instanceof Error ? lastError : new Error("Proxy request failed.")
}

export default App
