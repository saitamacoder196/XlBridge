# Slate Translation Prompt Template

Use this prompt when translating one batch (max 3 files) from Japanese to Vietnamese.

## Rules
- Translate only Japanese content in `value`.
- Keep unchanged:
- Technical IDs/functions: `CF_*`, `F_*`, `CS_*`
- VB/.NET tokens: `ByRef`, `ByVal`, `As String`, `As Integer`, ...
- Dates/timestamps, numeric IDs
- XML/HTML tags, escape markers like `\n`, `@(f)`
- Keep record order unchanged.
- Do not add/remove records.

## Input
Paste the JSON records from `translation_runs/manifests/batch_XXX.json`:

```json
[
  {"file": "共通処理実装状況一覧_0001.txt", "line_no": 6, "prefix": "[変更履歴]!A1", "value": "変更履歴"}
]
```

## Output (JSON only)

```json
[
  {"file": "共通処理実装状況一覧_0001.txt", "line_no": 6, "translated_value": "Lịch sử thay đổi"}
]
```
