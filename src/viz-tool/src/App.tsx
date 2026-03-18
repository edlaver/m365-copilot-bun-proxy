import { useEffect, useRef, useState } from "react"
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
const preferredDefaultModel = "m365-copilot"

export function App() {
  const [availableModels, setAvailableModels] = useState<string[]>([])
  const [selectedModel, setSelectedModel] = useState("")
  const [isLoadingModels, setIsLoadingModels] = useState(true)
  const [transformMode, setTransformMode] = useState<TransformMode>("mapped")
  const [requestType, setRequestType] = useState<RequestType>("chat/completions")
  const [selectedFixtureId, setSelectedFixtureId] = useState("")
  const [pane1, setPane1] = useState("")
  const [pane2, setPane2] = useState(emptyPaneText)
  const [pane3, setPane3] = useState(emptyPaneText)
  const [pane4, setPane4] = useState(emptyPaneText)
  const [pane4Data, setPane4Data] = useState<unknown>(null)
  const [selectedPane4FrameTypes, setSelectedPane4FrameTypes] = useState<number[]>([2])
  const [statusText, setStatusText] = useState("Ready")
  const [errorText, setErrorText] = useState<string | null>(null)
  const [isSubmitting, setIsSubmitting] = useState(false)
  const activeRequestControllerRef = useRef<AbortController | null>(null)

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

  useEffect(() => {
    const controller = new AbortController()

    async function loadModels() {
      setIsLoadingModels(true)
      try {
        const response = await fetchWithProxyRetry("/v1/models", {
          signal: controller.signal,
        })
        if (!response.ok) {
          throw new Error(`Model list request failed with ${response.status}.`)
        }

        const payload = (await response.json()) as {
          data?: Array<{ id?: unknown }>
        }
        const nextModels = Array.isArray(payload.data)
          ? payload.data
              .map((entry) => (typeof entry?.id === "string" ? entry.id : null))
              .filter((value): value is string => Boolean(value))
          : []

        if (nextModels.length === 0) {
          throw new Error("No models were returned by the proxy.")
        }

        setAvailableModels(nextModels)
        setSelectedModel((current) =>
          current && nextModels.includes(current)
            ? current
            : nextModels.includes(preferredDefaultModel)
              ? preferredDefaultModel
              : nextModels[0]
        )
      } catch (error) {
        if (controller.signal.aborted) {
          return
        }
        setAvailableModels([])
        setSelectedModel("")
        setErrorText(`Unable to load models. ${String(error)}`)
        setStatusText("Model list unavailable")
      } finally {
        if (!controller.signal.aborted) {
          setIsLoadingModels(false)
        }
      }
    }

    void loadModels()

    return () => {
      controller.abort()
    }
  }, [])

  useEffect(() => {
    setPane4(formatJson(filterPane4Data(pane4Data, selectedPane4FrameTypes)))
  }, [pane4Data, selectedPane4FrameTypes])

  function applyTrace(trace: TraceResponse, responseStatus: number | null) {
    if (trace.pane2 !== null) {
      setPane2(formatJson(trace.pane2))
    } else if (trace.error !== null) {
      setPane2(formatJson(trace.error))
    } else if (trace.status !== "pending") {
      setPane2(formatJson({ status: trace.status }))
    }

    if (trace.pane3 !== null) {
      setPane3(formatJson(trace.pane3))
    } else if (trace.status !== "pending") {
      setPane3(formatJson({ note: "No transformed upstream request was captured." }))
    }

    if (trace.pane4 !== null) {
      setPane4Data(trace.pane4)
    } else if (trace.status !== "pending") {
      setPane4Data({ note: "No upstream response was captured." })
    }

    setStatusText(buildTraceStatusText(trace, responseStatus))
    if (trace.error !== null) {
      setErrorText(formatInline(trace.error))
    }
  }

  async function handleSubmit() {
    setErrorText(null)
    setIsSubmitting(true)
    setStatusText("Submitting request")
    setPane2(emptyPaneText)
    setPane3(emptyPaneText)
    setPane4Data(null)
    setSelectedPane4FrameTypes([2])

    if (!selectedModel) {
      setErrorText("A model must be selected before submitting.")
      setStatusText("Model required")
      setIsSubmitting(false)
      return
    }

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
    const requestController = new AbortController()
    activeRequestControllerRef.current = requestController
    const endpoint =
      requestType === "chat/completions"
        ? "/v1/chat/completions"
        : "/v1/responses"

    let responseText = ""
    let responseStatus: number | null = null
    const tracePromise = waitForTrace(traceId, requestController.signal, (trace) => {
      applyTrace(trace, responseStatus)
    })
    try {
      const response = await fetchWithProxyRetry(endpoint, {
        method: "POST",
        signal: requestController.signal,
        headers: {
          "content-type": "application/json",
          "x-m365-openai-transform-mode": transformMode,
          "x-m365-viz-trace-id": traceId,
        },
        body: JSON.stringify(replaceFixtureTemplates(parsedBody, selectedModel)),
      })
      responseStatus = response.status
      responseText = await response.text()

      const trace = await tracePromise
      if (trace) {
        applyTrace(trace, responseStatus)
        if (trace.status === "pending") {
          if (trace.pane2 === null && trace.error === null) {
            setPane2(formatJson(parseJsonOrWrap(responseText, responseStatus ?? 0)))
          }
          setErrorText("Trace data was not available before the timeout.")
          setStatusText(`${buildTraceStatusText(trace, responseStatus)} · timeout`)
        }
        return
      }

      setPane2(formatJson(parseJsonOrWrap(responseText, responseStatus ?? 0)))
      setErrorText("Trace data was not available before the timeout.")
      setStatusText(`Proxy ${responseStatus ?? 0} · trace timeout`)
    } catch (error) {
      if (requestController.signal.aborted) {
        setErrorText("Request was cancelled.")
        setStatusText("Request cancelled")
        return
      }
      requestController.abort()
      setPane2(formatJson(parseJsonOrWrap(responseText, responseStatus ?? 0)))
      setErrorText(`Request failed. ${String(error)}`)
      setStatusText("Request failed")
    } finally {
      if (activeRequestControllerRef.current === requestController) {
        activeRequestControllerRef.current = null
      }
      setIsSubmitting(false)
    }
  }

  function handleCancel() {
    activeRequestControllerRef.current?.abort()
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
                Pane 1 is editable. Panes 2-4 update from the live proxy trace while the request runs.
              </p>
            </div>

            <div className="flex flex-col gap-3 xl:min-w-[840px]">
              <div className="grid gap-3 md:grid-cols-[220px_180px_220px_minmax(0,1fr)_auto]">
                <ToolbarSelect
                  label="Model"
                  value={selectedModel}
                  onChange={setSelectedModel}
                  options={
                    availableModels.length > 0
                      ? availableModels.map((model) => ({
                          label: model,
                          value: model,
                        }))
                      : [
                          {
                            label: isLoadingModels ? "Loading models..." : "No models available",
                            value: "",
                          },
                        ]
                  }
                  disabled={isLoadingModels || availableModels.length === 0}
                />
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
                    variant={isSubmitting ? "destructive" : "default"}
                    onClick={() => {
                      if (isSubmitting) {
                        handleCancel()
                        return
                      }
                      void handleSubmit()
                    }}
                  >
                    {isSubmitting ? (
                      <LoaderCircle className="size-4 animate-spin" />
                    ) : (
                      <RefreshCcw className="size-4" />
                    )}
                    {isSubmitting ? "Cancel" : "Submit"}
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
            actions={
              <div className="flex items-center gap-2">
                <span className="text-[11px] font-medium uppercase tracking-[0.18em] text-muted-foreground">
                  Frame type
                </span>
                {[1, 2, 3].map((frameType) => {
                  const pressed = selectedPane4FrameTypes.includes(frameType)
                  return (
                    <Button
                      key={frameType}
                      variant={pressed ? "default" : "outline"}
                      size="xs"
                      disabled={!hasPane4Frames(pane4Data)}
                      aria-pressed={pressed}
                      onClick={() => {
                        const nextTypes = pressed
                          ? selectedPane4FrameTypes.filter((value) => value !== frameType)
                          : [...selectedPane4FrameTypes, frameType].sort((a, b) => a - b)
                        setSelectedPane4FrameTypes(nextTypes)
                      }}
                    >
                      {frameType}
                    </Button>
                  )
                })}
              </div>
            }
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
  disabled?: boolean
}

function ToolbarSelect({
  label,
  value,
  onChange,
  options,
  disabled = false,
}: ToolbarSelectProps) {
  return (
    <div className="space-y-1.5">
      <Label>{label}</Label>
      <select
        className={cn(
          "h-11 w-full rounded-xl border border-border/70 bg-background/85 px-3 text-sm shadow-sm outline-none transition focus:border-ring focus:ring-3 focus:ring-ring/25"
        )}
        value={value}
        disabled={disabled}
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

function replaceFixtureTemplates(value: unknown, selectedModel: string): unknown {
  if (value === "{{model}}") {
    return selectedModel
  }
  if (Array.isArray(value)) {
    return value.map((item) => replaceFixtureTemplates(item, selectedModel))
  }
  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value).map(([key, nestedValue]) => [
        key,
        replaceFixtureTemplates(nestedValue, selectedModel),
      ])
    )
  }
  return value
}

function hasPane4Frames(value: unknown): boolean {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return false
  }
  const frames = (value as Record<string, unknown>).frames
  return Array.isArray(frames) && frames.length > 0
}

function filterPane4Data(value: unknown, selectedTypes: number[]): unknown {
  if (!hasPane4Frames(value)) {
    return value ?? { note: "No upstream response was captured." }
  }

  const typed = value as Record<string, unknown>
  const frames = typed.frames as unknown[]
  if (selectedTypes.length === 0) {
    return typed
  }

  return {
    ...typed,
    frames: frames.filter((frame) => {
      if (!frame || typeof frame !== "object" || Array.isArray(frame)) {
        return false
      }
      const type = (frame as Record<string, unknown>).type
      return typeof type === "number" && selectedTypes.includes(type)
    }),
  }
}

async function waitForTrace(
  traceId: string,
  signal: AbortSignal,
  onUpdate: (trace: TraceResponse) => void
): Promise<TraceResponse | null> {
  let lastTrace: TraceResponse | null = null
  let lastUpdatedAtUnix = -1

  for (let attempt = 0; attempt < 240; attempt += 1) {
    if (signal.aborted) {
      return lastTrace
    }

    let response: Response
    try {
      response = await fetch(`/__viz/traces/${traceId}`, { signal })
    } catch {
      return lastTrace
    }

    if (response.ok) {
      const trace = (await response.json()) as TraceResponse
      lastTrace = trace
      if (trace.updatedAtUnix !== lastUpdatedAtUnix) {
        lastUpdatedAtUnix = trace.updatedAtUnix
        onUpdate(trace)
      }
      if (trace.status !== "pending") {
        return trace
      }
    }

    await delay(250)
  }

  return lastTrace
}

function buildTraceStatusText(
  trace: TraceResponse,
  responseStatus: number | null
): string {
  const parts = [trace.status === "pending" ? "Tracing live" : `Trace ${trace.status}`]

  if (trace.transport) {
    parts.push(`transport ${trace.transport}`)
  }
  if (trace.upstreamStatusCode !== null) {
    parts.push(`upstream ${trace.upstreamStatusCode}`)
  }
  const resolvedProxyStatus = responseStatus ?? trace.proxyStatusCode
  if (resolvedProxyStatus !== null) {
    parts.push(`proxy ${resolvedProxyStatus}`)
  }

  return parts.join(" · ")
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
      if (init?.signal?.aborted) {
        throw error
      }
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
