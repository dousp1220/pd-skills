---
name: glazer-prospectus-refresh
description: Refresh a bond prospectus through PDRouter with Glazer, using one historical DOCX prospectus and one or more audit/financial reports, then downloading the highlighted annotated DOCX result.
metadata:
  author: PAODINGAI
  version: 1.0.0
  openclaw:
    requires:
      env: [PAODINGAI_API_KEY]
      bins: [node]
---

# Glazer Prospectus Refresh

Use this skill to refresh an existing bond prospectus through **PDRouter**. The workflow uploads one historical prospectus and one or more audit/financial reports, starts a Glazer refresh project, watches incremental progress events, and downloads the highlighted annotated DOCX result.

All service requests must go through PDRouter. PDRouter authenticates the caller, bills the request, and forwards it to Glazer. Do not call the Glazer upstream host directly from this skill.

## Configuration

```bash
export PAODINGAI_API_KEY="<your-pdrouter-api-key>"
export PAODINGAI_API_BASE_URL="https://platform.paodingai.com/"
```

Environment variables:

- `PAODINGAI_API_KEY`: Required. Bearer API key used when calling PDRouter.
- `PAODINGAI_API_BASE_URL`: Optional. PDRouter origin URL. Defaults to `https://platform.paodingai.com/`.

Treat credentials as confidential. Never echo API keys.

## Commands

Pass one `--period` per input file, in the same order as the files, plus `--new-period` for the new prospectus. If the user identified a historical prospectus, pass a single `--file-type prospectus`. The script assigns it to the prospectus `.docx` (by `募集`/`prospectus` in the name, or the only `.docx`) and treats remaining files as audit. Pass `--file-type` once per file, in file order, only when several `.docx` files are ambiguous. The script stamps `{index}__{basename}` onto uploads internally so the API can reassemble files even if two files share a name.

Never pass inline JSON on the command line. Use `--period` / `--new-period`, or `--periods-file`. Nested quotes break on Windows PowerShell.

```bash
node skills/glazer-prospectus-refresh/scripts/prime_document.mjs start \
  <historical-prospectus.docx> <audit-report.pdf> [audit-report-2.xlsx ...] \
  --period "2023年、2022年" \
  --period "2024年" \
  --new-period "2024年、2023年" \
  [--file-type prospectus] \
  [--output-dir <dir>] \
  [--json]

node skills/glazer-prospectus-refresh/scripts/prime_document.mjs watch <project_id> \
  [--after-event-id <id>] [--json]

node skills/glazer-prospectus-refresh/scripts/prime_document.mjs status <project_id> [--json]

node skills/glazer-prospectus-refresh/scripts/prime_document.mjs download <project_id> \
  [--output-dir <dir>] [--json]
```

If a JSON object is unavoidable, write it to a UTF-8 file and pass `--periods-file`. Do not put JSON on the command line.

### `start`

Validates local files, submits one multipart request to PDRouter, immediately prints an accepted JSON object containing `project_id`, then watches incremental events until a terminal state is reached. After `succeeded` or `export_failed`, it downloads the result next to the historical prospectus.

`start` already watches and downloads. While `start` is running:

- Do not launch a second `watch`.
- Do not call `start` again.

The first stdout line is only the accepted `project_id`. The process is not finished until it prints the final JSON object or exits.

Pass `--output-dir` only when the user explicitly specified an output directory. Never invent a workspace `output/` folder.

### `watch`

Resume incremental progress events for an existing `project_id` only after a local interruption. Pass the previous `max_event_id` as `--after-event-id`. The command returns the latest `max_event_id` so the next watch call can continue without repeating old events.

### `status`

Performs one foreground status request. Do not poll `/status` inside the `start`/`watch` loop; those commands already poll `/watch` every 20 seconds and include the current status.

If the user asks for a current snapshot, you may run `status` once even while a watcher is running.

### `download`

Downloads the terminal DOCX result after the project reaches `succeeded` or `export_failed`. The client verifies declared file size and SHA-256 when provided, writes the result atomically, and avoids overwriting an existing same-name file.

Prefer letting `start` download the result. Do not invent an `output/` directory.

If the user explicitly specified an output directory, pass that path as `--output-dir`.

A standalone `download` cannot infer the historical prospectus directory. If `start` did not write the file and you must run `download` separately, set `--output-dir` to the directory containing the historical prospectus, unless the user specified a different output directory.

If `start` already accepted a `project_id` and later failed during watch or download, resume with `watch` or `download`. Never create a second project with `start`.

## Execution Constraints

- Run start, watch, status, and download only through `scripts/prime_document.mjs`.
- Do not extract internal functions from the script, search it for validators, copy fragments, or write a temporary `preflight.mjs`.
- Do not inspect, truncate, or reconstruct the script source to "understand" validation. The script already owns local validation, PDRouter transport, progress watching, and download checksums.
- Route all HTTP through the script. Do not call PDRouter or Glazer APIs directly.
- If the user identified the historical prospectus, pass one `--file-type prospectus`. Remaining files default to audit even if the prospectus is not the first file. If no `--file-type` is passed, let the script classify basenames.
- Only the historical prospectus must be `.docx`. Audit/financial reports may be `.pdf`, `.doc`, `.docx`, `.xls`, or `.xlsx`. Never treat a PDF, `.doc`, Excel, or image as the historical prospectus, and never offer those files as a substitute prospectus.
- Do not pass `--output-dir` unless the user explicitly specified an output directory.
- If a standalone `download` is required because `start` did not write the result, set `--output-dir` to the directory of the historical prospectus. Never invent a workspace `output/` folder.
- Tool output is an implementation detail. After each command, interpret the returned status for the user; do not paste raw JSON or truncated source fragments.

## Mandatory Agent Preflight Validation

Before invoking `prime_document.mjs start`, the agent must perform a preliminary validation of every supplied file and every supplied period. Do not upload files or call PDRouter until all checks pass. If a check fails, explain the missing or ambiguous information in plain language and ask the user to correct it.

1. **Classify every file from the user's explicit type first.**
   - If the user already said a file is the historical prospectus or an audit/financial report, use that declared type.
   - In that case, only validate format: the historical prospectus must be `.docx`; an audit/financial report may be `.pdf`, `.doc`, `.docx`, `.xls`, or `.xlsx`. A `.docx` financial report is valid as audit, not as a substitute prospectus.
   - Do not reject a declared file because its basename lacks `募集`, `prospectus`, `审计`, `财务`, `财报`, `年报`, `audit`, or `annual`.
   - If the user identified the prospectus, pass one `--file-type prospectus`. The script picks the prospectus `.docx` even when it is not the first file. Remaining files default to audit. Pass `--file-type` in file order only when multiple `.docx` files are ambiguous.
   - Guess from basename only when the user did not specify a type.
   - Basename guessing uses `募集`/`prospectus` for prospectus and `审计`/`财务`/`财报`/`年报`/`audit`/`annual` for audit/financial reports.
   - If a file has no declared type and its basename is ambiguous, ask the user; do not call `start`.
2. **Require exactly one historical prospectus.**
   - Only this file must be `.docx`. PDF, `.doc`, `.xls`, `.xlsx`, and images are not accepted as the historical prospectus.
   - A `.docx` audit/financial report is allowed, but it does not replace the historical prospectus. There must still be exactly one historical prospectus `.docx`.
   - If there are zero or multiple prospectus files, ask the user to identify the single historical prospectus `.docx`.
   - If the workspace has no historical prospectus `.docx`, ask the user to provide that `.docx` or its path. Do not offer to use a PDF or a financial report as the prospectus, and do not present multiple-choice options that include invalid formats.
3. **Require at least one audit/financial report.**
   - Accepted report extensions are `.pdf`, `.doc`, `.docx`, `.xls`, and `.xlsx`.
   - Reject screenshots, images, text files, unsupported spreadsheets, and other formats before calling the script.
4. **Require a period for every uploaded file.**
   - Pass `--period` once per file, in the same order as the `start` file arguments.
   - Never silently map one file's period to another file.
   - A missing or empty period must be clarified with the user. A four-digit year such as `2024` is valid and means `2024年`.
5. **Require the new prospectus period.**
   - The new target period must be passed as `--new-period`.
   - It must be a human-readable period string, not a bare year number.
6. **Validate period content before submission.**
   - Period strings may contain annual periods, month periods, `及`/`和`/`与`, `月末`, `年度`, quarters, half years, dates, and year ranges such as `至`/`到`.
   - Reject malformed months, reversed month ranges such as `2024年4-3月`, invalid same-year ranges such as `2022年-2022年`, and non-contiguous years such as `2022年、2024年`. Year ranges are expanded inclusively before validation.
   - Year-plus-month lists such as `2024年1-3月、2023年、2022年` are valid.
   - The historical prospectus and new prospectus must each contain at least two periods.
   - At least one audit/financial report must be supplied.
   - An audit/financial report may contain a single period. The service expands one prior year in ascending order, for example `2023年` becomes `2022年、2023年`.
   - Any period newly added by the new prospectus must be covered by the supplied audit/financial reports after that expansion.

Only after this preflight passes should the agent invoke `prime_document.mjs start` with `--period` / `--new-period`. The script repeats these checks locally as a final guard; the agent must not rely on a server-side error to discover an obvious input problem.

## Input Rules

The user-facing prompt can be plain Chinese. The agent must transform it into `--period` values in file order plus `--new-period`.

Required files:

- Exactly one historical prospectus:
  - Format must be `.docx`. This is the only file with that hard format rule.
  - If the user declared the type, do not require the basename to contain `募集` or `prospectus`.
- At least one audit/financial report:
  - Formats: `.pdf`, `.doc`, `.docx`, `.xls`, `.xlsx`. A financial report may itself be `.docx`.
  - If the user declared the type, do not require the basename to contain `审计`, `财务`, `财报`, `年报`, `audit`, or `annual`.

Period strings:

- Prefer human-readable period strings. A four-digit year such as `2024` is accepted as `2024年`.
- Accepted examples:
  - `2022年、2023年、2024年3月`
  - `2024年1-3月、2023年、2022年`
  - `2024年、2023年`
  - `2023年、2022年`
  - `2025年及2026年`
  - `2026年3月末`
  - `2022年-2025年3月末`
  - `2025年-2022年`
  - `2024年度`
  - `2022年至2025年`
  - `2022、2023、2024年`
  - `2024年1月至3月`
  - `截至2024年3月31日`
  - `2024年上半年`
  - `2024年一季度`
- Pass `--period` once per file, matching the `start` file order, plus `--new-period`.
- Do not pass inline JSON. Two files may share the same name.

Business validation performed locally by this script before submission, with the upstream service as a second guard:

1. Historical prospectus period must contain at least two periods.
2. New prospectus period must contain at least two periods.
3. At least one audit/financial report must be supplied. A single audit period is allowed; the service expands one prior year in ascending order.
4. Periods newly introduced by the new prospectus must be covered by audit/financial reports.

Default business limits:

- No retroactive adjustment.
- No guarantor refresh.
- No chapter-specific refresh.

## User-visible Progress Reporting

`start` and `watch` write progress diagnostics to stderr. Convert watch events into formatted Chinese UI updates. Do not paste raw JSON, stderr dumps, or untranslated `type` values unless the user explicitly asked for them.

Each watch event has `stage`, `type`, and `message`. Format them in the assistant UI as a short list, one line per new event:

- Translate `type`: `start` → 开始, `finish` → 已完成, `log` → 进行中, `fail` → 失败
- Show the stage name and the event message
- Skip exact duplicate events
- Keep the latest status visible; do not hide failures

Recommended layout:

```
项目已受理，项目 ID：123。

进度
- [已完成] 文档版式/表格结构解析完成
- [进行中] 正在抽取财务数据
- [进行中] 正在生成新的 Word 文档
```

Also use these summaries when the corresponding phase appears:

- 项目已受理，项目 ID：123。
- 正在预处理文档并解析版式/表格结构。
- 正在抽取财务数据。
- 正在生成新的 Word 文档。
- 刷报已完成，开始下载结果。
- 导出未完全成功，仍将尝试下载已生成的文档。

If there are no new events for a while, send a concise heartbeat such as "刷报仍在进行中，正在等待下一次进度更新。"

## Error Handling

When validation fails, explain it in business language and ask the user for the missing input:

- Unsupported format: only the historical prospectus must be DOCX; audit/financial reports may be PDF/Word/Excel, including `.docx`. If there is no prospectus `.docx`, ask the user to provide one. Do not suggest using a PDF or a financial report as the prospectus.
- Ambiguous file name: ask the user to identify whether a file is the prospectus or an audit report. A `.docx` may be either, depending on the user's declaration or the basename. A declared prospectus that is not `.docx` is still invalid.
- Missing period: ask for each file's period and the new prospectus period.
- Period coverage failure: this should be caught before submission; explain which new periods are missing from audit/financial reports.
- Watch or download failed after accept: keep the `project_id`, resume with `watch` or `download`, and do not call `start` again.
- System failure: tell the user Glazer could not generate the refreshed document and suggest retrying later.

## Long-running Execution

Large prospectus refresh jobs can run for a long time. Run `start` or `watch` as the host-managed background task when available. A local interruption only stops the local watcher; it does not cancel the remote Glazer project. Resume with:

```bash
node skills/glazer-prospectus-refresh/scripts/prime_document.mjs watch <project_id> --after-event-id <last_max_event_id>
```
