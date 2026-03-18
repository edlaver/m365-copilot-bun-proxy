import Editor from "@monaco-editor/react"
import type { ReactNode } from "react"

import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "@/components/ui/card"

type JsonPaneProps = {
  title: string
  description: string
  value: string
  readOnly?: boolean
  onChange?: (value: string) => void
  actions?: ReactNode
}

export function JsonPane({
  title,
  description,
  value,
  readOnly = false,
  onChange,
  actions,
}: JsonPaneProps) {
  const monacoTheme =
    typeof document !== "undefined" &&
    document.documentElement.classList.contains("dark")
      ? "vs-dark"
      : "light"

  return (
    <Card className="flex min-h-0 flex-col overflow-hidden">
      <CardHeader>
        <div className="space-y-1">
          <CardTitle>{title}</CardTitle>
          <CardDescription>{description}</CardDescription>
        </div>
        {actions ? <div className="shrink-0">{actions}</div> : null}
      </CardHeader>
      <CardContent className="min-h-0 flex-1">
        <Editor
          height="100%"
          defaultLanguage="json"
          language="json"
          theme={monacoTheme}
          value={value}
          onChange={(nextValue) => {
            if (!onChange) {
              return
            }
            onChange(nextValue ?? "")
          }}
          options={{
            automaticLayout: true,
            fontFamily: "Geist Variable, ui-monospace, SFMono-Regular, monospace",
            fontLigatures: false,
            fontSize: 13,
            lineNumbersMinChars: 3,
            minimap: { enabled: false },
            padding: { top: 16, bottom: 16 },
            readOnly,
            renderLineHighlight: "line",
            scrollBeyondLastLine: false,
            tabSize: 2,
            wordWrap: "on",
          }}
        />
      </CardContent>
    </Card>
  )
}
