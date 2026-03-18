/// <reference types="vite/client" />

export type RequestType = "chat/completions" | "responses"
export type TransformMode = "simulated" | "mapped"

export type FixtureManifestEntry = {
  id: string
  label: string
  requestType: RequestType
  transformModes: TransformMode[]
  sourceFile: string
  sourceLine: number
  fileName: string
}

const fixtureModules = import.meta.glob("../fixtures/generated/**/*.json", {
  eager: true,
  import: "default",
}) as Record<string, unknown>
const manualFixtureModules = import.meta.glob("../fixtures/manual/**/*.json", {
  eager: true,
  import: "default",
}) as Record<string, unknown>

const manifestModule = Object.entries(fixtureModules).find(([filePath]) =>
  filePath.endsWith("/manifest.json")
)?.[1] as FixtureManifestEntry[] | undefined
const manualManifestModule = Object.entries(manualFixtureModules).find(([filePath]) =>
  filePath.endsWith("/manifest.json")
)?.[1] as FixtureManifestEntry[] | undefined

if (!manifestModule) {
  throw new Error("Fixture manifest is missing.")
}

if (!manualManifestModule) {
  throw new Error("Manual fixture manifest is missing.")
}

const manifest = [...manualManifestModule, ...manifestModule]

const fixtureContentByFileName = new Map<string, string>()
for (const [filePath, moduleValue] of [
  ...Object.entries(manualFixtureModules),
  ...Object.entries(fixtureModules),
]) {
  if (filePath.endsWith("/manifest.json")) {
    continue
  }
  const fileName = filePath.split("/").at(-1)
  if (!fileName) {
    continue
  }
  fixtureContentByFileName.set(fileName, JSON.stringify(moduleValue, null, 2))
}

export function getFixtures(
  requestType: RequestType,
  _transformMode: TransformMode,
): Array<FixtureManifestEntry & { content: string }> {
  const selected = manifest.filter((fixture) => fixture.requestType === requestType)

  return selected.map((fixture) => ({
    ...fixture,
    content:
      fixtureContentByFileName.get(fixture.fileName) ??
      JSON.stringify({ error: `Missing fixture file '${fixture.fileName}'.` }, null, 2),
  }))
}

export type TraceResponse = {
  status: "pending" | "completed" | "failed"
  requestType: string
  transformMode: string
  transport: string
  proxyStatusCode: number | null
  upstreamStatusCode: number | null
  pane2: unknown
  pane3: unknown
  pane4: unknown
  error: unknown
  createdAtUnix: number
  updatedAtUnix: number
}
