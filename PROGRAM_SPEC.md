# Excel Copilot Translation System

This document explains how the Excel-based translation assistant works so that newcomers can understand the moving pieces without having to read the entire code base.

## Purpose and Capabilities

- Automate Japanese-to-English translation (and other target languages) of spreadsheet content directly inside Excel.
- Enrich translations by consulting bilingual reference materials so that terminology follows approved phrasing.
- Review existing translations for accuracy and style, flagging issues and suggesting fixes.
- Produce annotated outputs that include process notes, aligned reference pairs, and optional citation tables.

## High-Level Architecture

| Layer | Responsibilities | Key Modules |
| --- | --- | --- |
| Desktop UI (Flet) | Presents chat-style interface, collects user input, displays agent thoughts/action logs, streams progress, and keeps track of workbook/sheet context. | `desktop_app.py`, `excel_copilot/ui/chat.py`, `excel_copilot/ui/theme.py` |
| Worker Thread | Handles blocking operations: launches Playwright, builds the ReAct agent, loads Excel tool functions, executes user requests, and emits structured responses back to the UI. | `excel_copilot/ui/worker.py`, `excel_copilot/ui/messages.py` |
| Agent Layer | Maintains the ReAct conversation with the large language model (LLM), injects system prompts, and turns high-level instructions into tool calls. | `excel_copilot/agent/react_agent.py`, `excel_copilot/agent/prompts.py` |
| Tool Layer | Implements the concrete Excel automation tasks (translation with/without references, translation quality review, diff highlighting, shape operations, etc.). | `excel_copilot/tools/excel_tools.py`, `excel_copilot/tools/actions.py` |
| Core Services | Provide connectivity to Excel (via xlwings) and to the Microsoft 365 Copilot chat interface (via Playwright). Custom exceptions and configuration live here. | `excel_copilot/core/excel_manager.py`, `excel_copilot/core/browser_copilot_manager.py`, `excel_copilot/core/exceptions.py`, `excel_copilot/config.py` |

The translation tools rely on a shared `BrowserCopilotManager` instance to submit prompts to M365 Copilot and on `ExcelActions` to read/write spreadsheet data safely.

## Runtime Control Flow

1. **App startup (`desktop_app.py`)**  
   The Flet frontend boots, creates request/response queues, and launches `CopilotWorker` on a background thread. UI state transitions (`AppState`) keep the chat controls enabled/disabled appropriately.

2. **Worker initialization (`CopilotWorker._initialize`)**  
   - Spins up Playwright using paths/configuration from `excel_copilot/config.py`.  
   - Opens a persistent browser context and navigates to the Copilot web app (`BrowserCopilotManager.start`).  
   - Loads the allowed tool functions for the current mode and generates JSON schemas for the LLM (`create_tool_schema`).  
   - Builds the ReAct agent with the correct system prompt (`build_system_prompt`) and signals the UI that the worker is ready.

3. **Request loop (`CopilotWorker._main_loop`)**  
   The worker consumes `RequestMessage` objects from the queue:
   - `USER_INPUT` → formats the user instruction, dispatches it to the agent, and streams back `Thought`, `Action`, `Observation`, and `Final Answer` messages.
   - `STOP` → sets a cancellation event so long-running operations can halt gracefully.
   - `RESET_BROWSER` / `UPDATE_CONTEXT` / `QUIT` → adjust internal state or tear down resources.

4. **Agent execution (`ReActAgent._run`)**  
   The agent keeps a history of system/user/assistant messages, follows the Thought→Action→Observation loop, and dispatches tool calls by name. Observations are appended to the chat log and the cycle repeats until completion.

5. **Tool call handlers (`excel_copilot/tools/excel_tools.py`)**  
   Tool functions receive an `ExcelActions` instance plus the shared browser manager. They:
   - Read data from Excel ranges.
   - Build structured prompts for Copilot (ensuring cancellations and retries are respected).
  - Parse JSON responses, validate shapes, and write outputs back to Excel.
  - Return human-readable summaries that the agent reports in its final answer.

6. **UI rendering**  
   Messages returned by the worker are wrapped in `ChatMessage` components which apply consistent styling for thoughts, actions, observations, and final answers.

## Supported Modes

### 1. Translation (No References)

- Enabled when `CopilotMode.TRANSLATION` is selected.
- Agent system prompt restricts tool usage to `translate_range_without_references`, and enforces one translation column per source column (`excel_copilot/agent/prompts.py`).
- `translate_range_without_references` (wrapper around `translate_range_contents`) batches rows, submits a plain translation prompt, and writes only the translated text to the specified output range.

### 2. Translation with References

- Default mode (`CopilotMode.TRANSLATION_WITH_REFERENCES`).
- Requires at least one of `source_reference_urls` or `target_reference_urls`.
- `translate_range_with_references` enforces row-by-row batching (one item per request) so reference evidence can be captured precisely.
- Core pipeline inside `translate_range_contents`:
  1. **Reference ingestion**  
     - Combines URLs passed via `reference_urls`, `source_reference_urls`, and `target_reference_urls`.  
     - Normalizes entries by stripping quotes, resolving local file paths, and deduplicating. Invalid tokens generate warnings surfaced to the user.  
     - Reference directories are looked up in `_REFERENCE_FALLBACK_ROOTS`, which include the current working directory, the user’s Downloads folder, and any extra directories supplied through `COPILOT_REFERENCE_DIRS`.
  2. **Source sentence extraction**  
     - When references are available, constructs a multi-step prompt asking Copilot to locate up to ten quotations per `context_id` that match the source Japanese sentence. Non-body content (navigation, references) is filtered out.
  3. **Target pair extraction**  
     - Using the extracted Japanese sentences and the target-language URLs, requests aligned sentence pairs. Each pair must reuse the exact wording from the reference documents.
  4. **Translation prompt**  
     - For each row, builds a rich prompt instructing the LLM to translate the source while prioritizing terminology and phrasing found in `reference_pairs`.  
     - The prompt (updated in `excel_copilot/tools/excel_tools.py:1439`) tells the model to use reference sentences as the primary guide, borrowing wording when the meanings match, and documenting how references influenced the translation.
  5. **Response validation**  
     - Ensures every item contains `translated_text`, `process_notes_jp`, and `reference_pairs` arrays. Missing `process_notes_jp` entries are backfilled with a stock explanation when references were expected.  
     - Rejects outputs that remain Japanese when an English translation is requested.
  6. **Excel write-back**  
     - Writes translations, Japanese process notes, and formatted reference pairs into column blocks of width `_MIN_CONTEXT_BLOCK_WIDTH` (default 12 columns: translation, process notes, and up to ten reference-pair slots).  
     - Optionally mirrors translations back to the source cells when `overwrite_source` is true.  
     - Handles citation output ranges when provided, supporting single-column, per-cell, paired-column, or triplet layouts.
  7. **Progress & messaging**  
     - Logs per-row progress (cell previews) and aggregates warnings (e.g., truncated reference slots) that are returned to the agent.

### 3. Translation Quality Review

- Activated with `CopilotMode.REVIEW`.
- Tool: `check_translation_quality`.  
  - Reads Japanese source and translated ranges, builds review prompts, and expects a JSON array with status (`OK`/`REVISE`), notes, corrected text, and highlight markup.  
  - Writes results into status/issue/highlight columns (optionally corrected text).  
  - Enforces Japanese issue descriptions with the `Issue: ... / Suggestion: ...` format.  
  - Uses inline diff markers (`[DEL]...`, `[ADD]...`) in the highlighted text.

## Excel Automation (`ExcelActions`)

- Wraps `xlwings` to read/write ranges, adjust column widths, apply diff highlighting, and insert shapes (`excel_copilot/tools/actions.py`).  
- Provides utilities for normalized matrix dimensions, safe cell addressing, and logging progress back to the UI via a callback.
- Ensures Excel remains visible and responsive by keeping screen updating enabled and respecting Excel’s cell size limits (e.g., `32766` character maximum).

## Browser Automation (`BrowserCopilotManager`)

- Uses Playwright to drive the Microsoft 365 Copilot chat experience (`excel_copilot/core/browser_copilot_manager.py`). Key behaviours:
  - Launches a persistent profile (Edge/Chrome) with configurable headless/slow-motion options.
  - Navigates to `https://m365.cloud.microsoft/chat/` and ensures the page is ready before accepting prompts.
  - Manages prompt submission, copy-button detection, and output retrieval. Handles retries when the chat input fails to capture the full prompt text.
  - Supports session resets and cancellation by monitoring a `stop_event` and issuing UI-level stops if the user interrupts execution.
  - Copies the response text via clipboard or page DOM extraction, then returns plain text to the caller.

## Messaging Model

- **Requests** (`excel_copilot/ui/messages.py`) originate from the UI: user text, context updates (workbook name, sheet name, mode), stop/quit commands, and browser resets.
- **Responses** include status updates, agent thoughts/actions, Playwright logs, errors, and final summaries. The UI maps response categories to visual styles (`excel_copilot/ui/chat.py`).
- The worker serializes everything through queues to avoid thread-safety issues with Flet widgets.

## Configuration

`excel_copilot/config.py` centralizes environment-driven settings:

- Agent loop limits (`COPILOT_MAX_ITERATIONS`, `COPILOT_HISTORY_MAX_MESSAGES`).
- Playwright options (user data directory, channel preferences, headless mode, slow-mo delay, page timeouts, focus suppression offsets).
- Downloadable reference directories through `COPILOT_REFERENCE_DIRS`.
- Rich-text diff tuning via environment variables consumed in `excel_copilot/tools/actions.py`.

Most defaults focus on reliable desktop automation; they can be overridden via environment variables without code changes.

## Supporting Scripts

- `desktop_app.py` – Entry point that launches the Flet UI. Also contains an optional “autotest” scenario driven by environment variables.
- `move_any_translation.py` – One-off maintenance script that adjusts code blocks inside `excel_tools.py`.
- `update_prompt.py` – Patch script that updates the reference translation prompt section.

## Error Handling and Cancellation

- Tool functions raise `ToolExecutionError` when validation fails (e.g., mismatched output shapes, missing JSON keys), which propagate back to the agent for corrective action.
- Long-running loops check `_ensure_not_stopped()` helpers that watch the shared cancellation event, so STOP requests halt both Excel writes and Copilot prompts.
- The agent inspects tool observations and can issue retries with adjusted parameters before producing a final answer.

## Typical Translation-with-References Workflow

1. User enters an instruction identifying the source range, output columns, and reference URLs.
2. Agent converts the request into a `translate_range_with_references` call.
3. Tool loads source text from Excel, normalizes references, and extracts aligned quotations from the reference URLs.
4. Reference pairs are fed into the translation prompt, with explicit emphasis on reusing approved phrasing.
5. Copilot returns JSON containing translation, process notes (in Japanese, documenting how references guided choices), and the subset of reference pairs actually cited.
6. Tool writes translations, notes, and references into the output block, logging progress into the UI.
7. Final message lists any warnings (e.g., unused reference columns, coercions of local files) and confirms completion.

## Getting Started for New Contributors

- Run `desktop_app.py` to launch the UI and connect to Excel. Ensure Excel is open with the target workbook.
- Use the mode selector to switch between plain translation, reference-assisted translation, and review.
- Place reference documents either as HTTP(S) URLs or as files inside the working directory/Downloads folder. Local paths are auto-resolved to `file://` URIs.
- Monitor the UI’s Thought/Action/Observation stream to diagnose tool parameter issues; errors include detailed instructions for corrective action.

This architecture allows the AI agent to stay focused on orchestration while the deterministic Python layer guarantees that Excel edits remain well-structured and traceable.
