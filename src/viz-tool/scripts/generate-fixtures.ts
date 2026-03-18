import { promises as fs } from "node:fs"
import path from "node:path"
import ts from "typescript"

type JsonValue = string | number | boolean | null | JsonValue[] | { [key: string]: JsonValue }

type FixtureManifestEntry = {
  id: string
  label: string
  requestType: "chat/completions" | "responses"
  transformModes: Array<"simulated" | "mapped">
  sourceFile: string
  sourceLine: number
  fileName: string
}

type TestContext = {
  name: string
}

const repoRoot = path.resolve(import.meta.dir, "..", "..", "..")
const testsDir = path.join(repoRoot, "tests")
const outputDir = path.join(repoRoot, "src", "viz-tool", "src", "fixtures", "generated")

async function main(): Promise<void> {
  const testFiles = await findUnitTestFiles(testsDir)
  const manifest: FixtureManifestEntry[] = []
  const fixtureCounts = new Map<string, number>()

  await fs.rm(outputDir, { recursive: true, force: true })
  await fs.mkdir(outputDir, { recursive: true })

  for (const filePath of testFiles) {
    const sourceText = await fs.readFile(filePath, "utf8")
    const sourceFile = ts.createSourceFile(
      filePath,
      sourceText,
      ts.ScriptTarget.Latest,
      true,
      ts.ScriptKind.TS,
    )
    const transformModes = inferTransformModes(sourceText)
    const scopeValues = new Map<string, JsonValue>()

    collectTopLevelValues(sourceFile, scopeValues)
    walk(sourceFile, [], (node, testStack) => {
      const request = extractRequest(node, sourceFile, scopeValues)
      if (!request) {
        return
      }

      const sourceLine =
        sourceFile.getLineAndCharacterOfPosition(node.getStart(sourceFile)).line + 1
      const baseName = path.basename(filePath, path.extname(filePath))
      const requestKey = `${baseName}-${request.requestType}`
      const index = (fixtureCounts.get(requestKey) ?? 0) + 1
      fixtureCounts.set(requestKey, index)

      const id = `${baseName}-${request.requestType.replace("/", "-")}-${index.toString().padStart(3, "0")}`
      const fileName = `${id}.json`
      const labelPrefix =
        testStack.length > 0 ? testStack[testStack.length - 1].name : "Unnamed test"
      const label = `${labelPrefix} (${path.basename(filePath)}:${sourceLine})`
      const requestDir = path.join(outputDir, request.requestType === "chat/completions" ? "chat-completions" : "responses")

      manifest.push({
        id,
        label,
        requestType: request.requestType,
        transformModes,
        sourceFile: normalizePath(path.relative(repoRoot, filePath)),
        sourceLine,
        fileName,
      })

      void fs
        .mkdir(requestDir, { recursive: true })
        .then(() =>
          fs.writeFile(
            path.join(requestDir, fileName),
            `${JSON.stringify(request.body, null, 2)}\n`,
            "utf8",
          ),
        )
    })
  }

  manifest.sort((a, b) => a.label.localeCompare(b.label))
  await fs.writeFile(
    path.join(outputDir, "manifest.json"),
    `${JSON.stringify(manifest, null, 2)}\n`,
    "utf8",
  )
}

async function findUnitTestFiles(dir: string): Promise<string[]> {
  const entries = await fs.readdir(dir, { withFileTypes: true })
  const files: string[] = []
  for (const entry of entries) {
    const entryPath = path.join(dir, entry.name)
    if (entry.isDirectory()) {
      files.push(...(await findUnitTestFiles(entryPath)))
      continue
    }
    if (!entry.name.endsWith(".test.ts")) {
      continue
    }
    if (entry.name.endsWith(".integration.test.ts")) {
      continue
    }
    files.push(entryPath)
  }
  return files.sort((a, b) => a.localeCompare(b))
}

function inferTransformModes(
  sourceText: string,
): Array<"simulated" | "mapped"> {
  const hasSimulated = sourceText.includes("OpenAiTransformModes.Simulated")
  const hasMapped = sourceText.includes("OpenAiTransformModes.Mapped")
  if (hasSimulated && !hasMapped) {
    return ["simulated"]
  }
  if (hasMapped && !hasSimulated) {
    return ["mapped"]
  }
  return ["simulated", "mapped"]
}

function collectTopLevelValues(
  sourceFile: ts.SourceFile,
  scopeValues: Map<string, JsonValue>,
): void {
  for (const statement of sourceFile.statements) {
    if (!ts.isVariableStatement(statement)) {
      continue
    }
    for (const declaration of statement.declarationList.declarations) {
      if (!ts.isIdentifier(declaration.name) || !declaration.initializer) {
        continue
      }
      const resolved = resolveExpressionValue(declaration.initializer, scopeValues)
      if (resolved !== undefined) {
        scopeValues.set(declaration.name.text, resolved)
      }
    }
  }
}

function walk(
  node: ts.Node,
  testStack: TestContext[],
  visit: (node: ts.Node, testStack: TestContext[]) => void,
): void {
  let nextStack = testStack
  if (ts.isCallExpression(node) && isTestBlock(node)) {
    const testName = getStringLiteralArgument(node.arguments[0]) ?? "Unnamed test"
    nextStack = [...testStack, { name: testName }]
  }

  visit(node, nextStack)
  node.forEachChild((child) => walk(child, nextStack, visit))
}

function isTestBlock(node: ts.CallExpression): boolean {
  if (!ts.isIdentifier(node.expression)) {
    return false
  }
  if (!["test", "it"].includes(node.expression.text)) {
    return false
  }
  return node.arguments.length > 0
}

function extractRequest(
  node: ts.Node,
  sourceFile: ts.SourceFile,
  scopeValues: Map<string, JsonValue>,
): { requestType: "chat/completions" | "responses"; body: JsonValue } | null {
  if (ts.isNewExpression(node) && ts.isIdentifier(node.expression) && node.expression.text === "Request") {
    const url = node.arguments?.[0]
    const init = node.arguments?.[1]
    const requestType = resolveRequestTypeFromUrl(url)
    if (!requestType || !init || !ts.isObjectLiteralExpression(init)) {
      return null
    }
    const bodyExpression = getPropertyExpression(init, "body")
    const requestBody = extractJsonStringifyArgument(bodyExpression, scopeValues)
    if (requestBody === undefined) {
      return null
    }
    return { requestType, body: requestBody }
  }

  if (ts.isCallExpression(node) && ts.isIdentifier(node.expression) && node.expression.text === "postJson") {
    const requestType = resolveRequestTypeFromUrl(node.arguments[0])
    if (!requestType || node.arguments.length < 2) {
      return null
    }
    const requestBody = resolveExpressionValue(node.arguments[1], scopeValues)
    if (requestBody === undefined) {
      return null
    }
    return { requestType, body: requestBody }
  }

  return null
}

function resolveRequestTypeFromUrl(
  expression: ts.Expression | undefined,
): "chat/completions" | "responses" | null {
  const value = getStringLiteralArgument(expression)
  if (!value) {
    return null
  }
  if (value.includes("/v1/chat/completions")) {
    return "chat/completions"
  }
  if (value.includes("/v1/responses")) {
    return "responses"
  }
  return null
}

function extractJsonStringifyArgument(
  expression: ts.Expression | undefined,
  scopeValues: Map<string, JsonValue>,
): JsonValue | undefined {
  const unwrapped = unwrapExpression(expression)
  if (!unwrapped || !ts.isCallExpression(unwrapped)) {
    return undefined
  }
  if (!ts.isPropertyAccessExpression(unwrapped.expression)) {
    return undefined
  }
  if (unwrapped.expression.expression.getText() !== "JSON") {
    return undefined
  }
  if (unwrapped.expression.name.text !== "stringify") {
    return undefined
  }
  return resolveExpressionValue(unwrapped.arguments[0], scopeValues)
}

function resolveExpressionValue(
  expression: ts.Expression | undefined,
  scopeValues: Map<string, JsonValue>,
): JsonValue | undefined {
  const unwrapped = unwrapExpression(expression)
  if (!unwrapped) {
    return undefined
  }

  if (ts.isObjectLiteralExpression(unwrapped)) {
    const result: Record<string, JsonValue> = {}
    for (const property of unwrapped.properties) {
      if (ts.isPropertyAssignment(property)) {
        const key = getPropertyName(property.name)
        if (!key) {
          return undefined
        }
        const value = resolveExpressionValue(property.initializer, scopeValues)
        if (value === undefined) {
          return undefined
        }
        result[key] = value
        continue
      }

      if (ts.isShorthandPropertyAssignment(property)) {
        const value = scopeValues.get(property.name.text)
        if (value === undefined) {
          return undefined
        }
        result[property.name.text] = value
        continue
      }

      return undefined
    }
    return result
  }

  if (ts.isArrayLiteralExpression(unwrapped)) {
    const values: JsonValue[] = []
    for (const element of unwrapped.elements) {
      if (ts.isSpreadElement(element)) {
        return undefined
      }
      const value = resolveExpressionValue(element, scopeValues)
      if (value === undefined) {
        return undefined
      }
      values.push(value)
    }
    return values
  }

  if (ts.isIdentifier(unwrapped)) {
    return scopeValues.get(unwrapped.text)
  }

  if (ts.isStringLiteralLike(unwrapped) || ts.isNoSubstitutionTemplateLiteral(unwrapped)) {
    return unwrapped.text
  }

  if (ts.isNumericLiteral(unwrapped)) {
    return Number(unwrapped.text)
  }

  if (unwrapped.kind === ts.SyntaxKind.TrueKeyword) {
    return true
  }
  if (unwrapped.kind === ts.SyntaxKind.FalseKeyword) {
    return false
  }
  if (unwrapped.kind === ts.SyntaxKind.NullKeyword) {
    return null
  }

  return undefined
}

function unwrapExpression(
  expression: ts.Expression | undefined,
): ts.Expression | undefined {
  let current = expression
  while (current) {
    if (ts.isParenthesizedExpression(current) || ts.isAsExpression(current) || ts.isSatisfiesExpression(current)) {
      current = current.expression
      continue
    }
    return current
  }
  return current
}

function getPropertyExpression(
  objectLiteral: ts.ObjectLiteralExpression,
  propertyName: string,
): ts.Expression | undefined {
  for (const property of objectLiteral.properties) {
    if (!ts.isPropertyAssignment(property)) {
      continue
    }
    const key = getPropertyName(property.name)
    if (key === propertyName) {
      return property.initializer
    }
  }
  return undefined
}

function getPropertyName(name: ts.PropertyName): string | null {
  if (ts.isIdentifier(name) || ts.isStringLiteralLike(name) || ts.isNumericLiteral(name)) {
    return name.text
  }
  return null
}

function getStringLiteralArgument(
  expression: ts.Node | undefined,
): string | null {
  if (!expression) {
    return null
  }
  if (ts.isStringLiteralLike(expression) || ts.isNoSubstitutionTemplateLiteral(expression)) {
    return expression.text
  }
  return null
}

function normalizePath(value: string): string {
  return value.replaceAll(path.sep, "/")
}

await main()
