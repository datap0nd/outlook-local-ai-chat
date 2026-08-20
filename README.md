# MetoAI

A Windows-only AI chat add-in for classic Outlook in Microsoft Office
Professional Plus 2021. It installs locally, opens as a native Outlook sidebar,
and lets an OpenAI-compatible model request bounded read-only context from the
local Inbox and Sent Items. It opens and safely revises one linked unsent draft
for human review.

It does not use Microsoft 365 add-in deployment, Microsoft Graph, Entra ID, or
an external Outlook MCP server.

## Install

1. Close Outlook.
2. Download
   [OutlookLocalAIChatSetup.exe](https://github.com/datap0nd/outlook-local-ai-chat/releases/latest/download/OutlookLocalAIChatSetup.exe).
   This link tracks the **Latest** release, which is rebuilt automatically on
   every push to `main`.
3. Run the installer for your Windows account.
4. Start classic Outlook.
5. Choose **MetoAI > MetoAI** on the ribbon.
6. Open **Settings** and enter:
   - the OpenAI-compatible endpoint or base URL;
   - the API key;
   - **Allow insecure HTTP** only when a non-local endpoint uses plain HTTP.
7. Click **Refresh models** to load available model IDs from `GET /v1/models`,
   then choose or type a model that supports OpenAI-compatible chat tool calls.
8. Click **Check endpoint**. Save only after authentication, the selected
   model, and mailbox tool calling pass.
9. Optional: open **Writing style**, click **Analyze 15 sent emails**, review
   the generated drafting instructions, edit them, and enable the profile.

To update later, open **Settings** and click **Update MetoAI**. After a
confirmation it downloads the latest release installer, closes Outlook,
installs silently for your Windows account, and reopens Outlook.

Examples:

```text
https://ai.example.test/v1
https://ai.example.test/v1/chat/completions
http://127.0.0.1:1234/v1
```

HTTPS is recommended. Plain HTTP is accepted automatically for loopback
addresses such as `localhost` and `127.0.0.1`. For another HTTP host, enable
**Allow insecure HTTP** in Settings. That opt-in sends the API key, prompts, and
retrieved email context without transport encryption.

The first unsigned build may trigger a Windows SmartScreen warning. A trusted
code-signing certificate is required to remove that warning for normal company
distribution.

The chat sidebar renders in Microsoft Edge WebView2, which ships with
Windows 10/11 and Microsoft Edge. The embedded page is network-isolated by
CSP, never navigates anywhere, and inserts all model and mailbox text as
inert text nodes - model output is still never parsed as HTML.

## Model choice

MetoAI does not ship with a preset model list or a preferred default. After you
enter an endpoint and API key, use **Refresh models** in Settings to populate the
dropdown from `GET /v1/models`. You can also type any model ID manually. Every
dropdown entry is tagged **Vision** (reads email images) or **Text**
(filename-only), and the sidebar header shows the same tag for the saved model.

For email **images**, pick a model tagged **Vision**. Vision capability is
detected from the model ID: `vl` or `vision` in the name (for example
`qwen3-vl-30b`), multimodal Gemma generations (`gemma3`/`gemma-4` and later),
and common vision families such as LLaVA, Pixtral, MiniCPM-V, InternVL,
Moondream, and SmolVLM. MetoAI loads image attachments automatically and sends
them as multimodal input, capped at eight images per request. Multimodal Gemma
requires the server to load its vision projector; if the server rejects image
input, use a `vl` model instead. Text-only models get spreadsheet text and
image metadata only, and the chat will say so if you ask about an image.
Optional: enable **Auto-switch to vision for images** to temporarily use a
discovered vision model while keeping your everyday text model saved. Save after
**Refresh models** so auto-switch knows which vision models are available.

Embedding-only models are excluded from discovery. **Check endpoint** verifies
authentication with a lightweight `search_mailbox` tool-call probe. It does not
read mailbox data during the check. Model discovery allows up to 15 seconds;
the tool-call probe allows up to 90 seconds.

## Use

1. Open **MetoAI**. The chat appears as a sidebar inside Outlook.
   You can also right-click selected email messages and choose **Send to MetoAI**.
   The sidebar opens with `Selected: subject` at the top. Common `RE:`, `FW:`,
   and `FWD:` prefixes are hidden in that display.
2. To choose a bounded group first, enter `/search person or topic`. MetoAI
   searches locally and keeps the newest ten matching Inbox or Sent Items
   emails as the working set. No email body is sent during this command.
3. Review the listed subjects and send another `/search` to replace the set if
   it is wrong. Results appear in a collapsible working-set layer as ten
   distinct email cards with subject, sender, and date. Use `/search clear` to
   remove it. The layer collapses automatically when you send a normal AI
   prompt and can be reopened with **Show**.
4. Alternatively, Ctrl+click one to ten emails in Outlook, then choose
   **Add email**, right-click **Send to MetoAI**, or drag the selected messages
   onto the MetoAI pane. Multiple messages become the same locked working set.
5. Use the **+** menu or drag files from Windows Explorer to add external
   context: up to three documents and four images. Documents go through the
   same extractors as email attachments (PDF, Office, text formats), and
   images become vision input with a tray thumbnail. HTML files are read as
   inert text, not rendered. Each file may be up to 25 MB on disk; extracted
   text is bounded to 48,000 characters per document and 120,000 characters
   total. A file that exceeds a cap still appears in the tray as an amber
   warning chip explaining what was kept — oversized files are noted, and
   over-length text keeps its first 48,000 characters with a truncation
   notice the model can see.
6. Ask a normal mailbox question. When a working set exists, the model can read
   only those emails. Without one, it may perform one bounded mailbox search
   and load no more than ten unique email bodies for the request. Meeting
   invites and calendar items are readable like email — subject, body, time,
   location, and attachments — but MetoAI can never accept, decline, or
   schedule anything. When a body
   is loaded, MetoAI also reads up to ten **email attachments** per
   message: images (PNG, JPEG, GIF, BMP, WebP, TIFF), spreadsheets (XLSX,
   XLSM, XLSB, XLTX, XLTM, XLS, CSV, TSV — all worksheets, including
   binary BIFF12 workbooks), documents (PDF, PPTX, PPTM, PPSX, PPSM,
   POTX, DOCX, DOCM, DOTX, DOTM, PPT, DOC, RTF, and OpenDocument
   ODT/ODS/ODP), attached Outlook messages (MSG, OFT — subject, sender,
   and body), and text files (TXT, MD, LOG, JSON, XML, YAML, INI, HTML,
   EML). Unknown extensions are identified by content (image, OOXML or
   binary Office, OpenDocument, MSG, PDF, or plain text). PDF
   extraction reads the text layer, including CID-font PDFs from Word and
   Chrome — scanned PDFs yield a clear "no readable text" note. Legacy
   binary Office files get best-effort extraction. Every attachment is
   listed; anything unreadable is noted rather than silently skipped.
   Attachments up to 25 MB are read; extraction is streamed and bounded
   to 48,000 characters per attachment and 120,000 characters per
   message, with an explicit truncation notice when more content
   remains.
   Small inline images embedded in the body (64 KB or less) are treated
   as signature graphics and ignored, with a note in the tool result;
   pasted screenshots and photos are far larger and are always read.
   Attachments are decrypted locally through Outlook COM before reading.
7. The sidebar records which bounded context operations ran.
8. Ask explicitly, for example "create a reply draft" or "write an email."
   Local code recognizes that drafting intent and exposes one creation attempt
   for that request. The draft opens unsent in Outlook. You can also
   right-click an email and choose **MetoAI - Suggest a response**: the
   sidebar asks up to three quick questions (reply tone plus up to two
   model-suggested questions specific to that email), and your answers
   shape the reply draft. Skipping the questions goes straight to a
   draft. The composed request goes through the same drafting pipeline,
   so it still authorizes exactly one draft that opens unsent for review.
9. A mailbox question without explicit drafting language cannot expose draft
   creation. Loaded email text and model output cannot authorize it.
10. The same Outlook draft stays linked to that chat. Follow-up instructions such
   as "make it shorter" or "bold the deadline" update and redisplay that exact
   unsent item. No second draft is created.
11. Review, edit, address, and send the message using Outlook's normal editor.

Selecting an email is optional for mailbox questions. When one is selected, the
model receives its metadata and may request its body using the temporary
`selected` handle. A two-to-ten-email selection is stored as a locked working
set with `context1` through `context10` handles. The conversation and working set
remain in memory until cleared or Outlook closes. `/search clear` removes the
email working set but retains external files. **Clear** removes all context, and
**New** starts a new conversation with no retained context.

Settings is organized into four tabs: **Connection** (endpoint, model, API
key, updates), **Gemini** (Google sign-in with a responsible-use notice —
cloud processing under your account; follow your organization's
confidential-data policies), **Writing soul**, and **Support** (describe a
problem and MetoAI opens a pre-filled, unsent report email to the creator
with the recent diagnostic log — timestamps, operations, and error codes
only — for you to review and send yourself).

The **writing soul** is a small, editable portrait of how you write.
Analysis never runs automatically: it requires a click, reads at most 15
recent usable Sent Items messages, removes obvious quoted history, and
sends bounded samples to the configured AI endpoint. The result is visible
and editable before saving. A **soul strength** slider (10–100) controls
how strongly drafts follow your voice, and **hard draft rules** (one per
line) are followed exactly in every draft. Soul, strength, and rules apply
only to draft creation and revision, and only to wording, greeting,
cadence, and sign-off. They cannot alter any capability or security rule.

## Limits tab

Settings has a **Limits** tab. **Use recommended limits** stays ticked by
default and keeps the values this README describes. Untick it to adjust,
at your own risk: the reading-budget multiplier (email bodies, attachments,
and documents, up to x8), your message length (up to 16,000 characters),
answer length (up to 48,000), history turns (up to 24), tool rounds and
tool calls per round (up to 8 each). Every slider is hard-clamped in code -
the settings file cannot push a value past those bounds. Raising limits
sends more text to the model and can overflow a small local model's context
window. Capability guardrails are not adjustable from here or anywhere
else: the ten-email working set, one draft per request, and
never-send/never-save stay fixed. With Gemini sign-in the reading budgets
already scale x4 automatically; the larger of that and your multiplier
wins.

## Hard security boundary

The model is not given general Outlook access. Draft creation is a narrowly
scoped exception, not a general mutation permission.

- A request without a working set exposes three read-only tools:
  `search_mailbox`, `read_messages`, and `read_thread`. A request with a locked
  working set exposes only `read_messages`, and only for its ten approved
  handles.
- `create_draft` is added only when local code recognizes an explicit drafting
  instruction in the latest user-written prompt, such as "create a draft" or
  "write a reply." Model output and email content never enter that decision.
- Once one draft exists, `create_draft` disappears and only `update_draft` is
  eligible. Local code exposes it only for a recognized revision instruction,
  and it can mutate only that linked unsent item once per user request.
- The draft host requires `create_draft` to be the only tool call in its model
  response, validates strict arguments, and atomically consumes permission
  before creating anything.
- One chat can link at most one draft. A user request can make at most one draft
  creation or update attempt.
- The local hosts reject every other tool name and cap tool calls, tool rounds,
  result counts, message bodies, draft fields, and total returned context.
- Search results use temporary handles. The model cannot submit arbitrary COM
  objects, Outlook commands, or executable code.
- A reply draft must include the exact temporary handle for its source email.
  The local host rejects missing, expired, or invented handles and never falls
  back to the selected or latest mailbox item.
- The model client never receives the Outlook application object or draft
  service.
- Model output is length-limited text displayed in a Windows control. It is
  never evaluated, executed, or rendered as model-provided HTML. Local code removes Markdown
  emphasis markers and applies only native bold spans in the transcript.
- Only a one-request authorization derived locally from explicit user drafting
  intent can create the linked draft. Later revisions require both recognized
  revision intent and that local linked-draft session.
- The model-invoked mailbox host remains read-only. A separate draft host accepts
  only the bounded `create_draft` and `update_draft` operations.
- The draft path exposes no send, move, delete, schedule, BCC, arbitrary HTML,
  or mailbox traversal operation.
- Classic Outlook COM has no permission-manifest switch for sending. MetoAI
  instead hardcodes the absence of a send capability, keeps the Outlook object
  outside the model client, and verifies the source plus compiled assembly in CI.
- Drafts are saved and displayed as unsent Outlook items.
- CI fails if forbidden Outlook action calls are introduced.

These controls let model output select read-only context and, after explicit
local authorization, create one unsent draft. They prevent it from reaching an
email-send or source-mailbox mutation capability. They do not claim protection
against a compromised Windows account, modified add-in binary, vulnerabilities
in Outlook or .NET, or an administrator replacing installed files.

See [SECURITY.md](SECURITY.md) for the full threat model.

## Data flow

Every chat request initially sends the configured endpoint:

- selected email metadata, or metadata for up to ten working-set emails;
- up to 12 recent chat turns;
- the current prompt;
- up to three explicitly added bounded text files;
- the editable writing profile only when drafting is locally authorized and the
  profile is enabled.

The model may then request:

- one search with up to ten bounded result summaries from the primary Inbox
  and Sent Items when no working set is locked;
- no more than ten unique message bodies across the entire request;
- conversation messages only within that same request-wide ten-body limit;
- at most four tool calls per round and four context-retrieval rounds.

`/search` is handled locally before an LLM request is created. It returns at
most ten metadata matches and does not transmit bodies. A later normal prompt
sends the working-set metadata and exposes only the body-read tool for those
exact handles.

When the latest user prompt explicitly asks to create or open a draft, local
intent rules expose `create_draft` for that request. Its bounded arguments may
contain a new-message subject,
recipients, CC recipients, and body, or a reply body plus the exact temporary
handle of a searched or selected source message. The tool can only save and
display one unsent Outlook draft. While that draft is
linked, a recognized revision request exposes `update_draft` instead. Each
update supplies the complete bounded body and optional exact phrases to bold.
The local formatter HTML-encodes all text and inserts only fixed headings,
subheadings, lists, dividers, paragraphs, and `<strong>` markup. The model can
request these visual structures with a small text layout syntax, but raw HTML is
rejected. If a model returns Markdown emphasis markers, the shared local formatter
removes them and applies real bold formatting in both MetoAI and Outlook. Stray
formatting asterisks are removed. Arbitrary model HTML is never accepted. Neither
tool has a send operation.

Email bodies are sent only when the model requests them through an approved
read-only tool. The add-in does not index, upload, or transmit the entire
mailbox automatically.

The optional Settings check sends the API key to `GET /v1/models` when that route
is available, then submits a lightweight synthetic chat request that contains no
mailbox data. **Refresh models** loads the dropdown without running that probe.
A successful check proves that authentication, the entered model, and the
tool-call response shape work before the first real mailbox question.

Nothing is sent to Microsoft 365 by the add-in. Outlook itself continues to use
whatever mail server your organization configured.

## API compatibility

The endpoint must support:

```http
POST /v1/chat/completions
Authorization: Bearer YOUR_KEY
Content-Type: application/json
```

The request uses `model`, `messages`, `stream: false`, and standard
OpenAI-compatible function tools. The endpoint and selected model must support
chat-completions tool calling. The final response must provide
`choices[0].message.content` as text. `GET /v1/models` is optional. When
available, **Refresh models** or **Check endpoint** populates the editable model
list in Settings.

Cloud OpenAI-compatibility layers work too. For **Google Gemini**, set the
base URL to `https://generativelanguage.googleapis.com/v1beta/openai` and use
an API key from Google AI Studio as the key; Gemini models are detected as
vision-capable and support tool calling.

**Gemini with Google sign-in (no API key):** enable *Use Google Gemini with
browser sign-in* in Settings and click **Sign in with Google**. A browser
window opens on Google's own sign-in pages (standard OAuth installed-app flow
with PKCE and a loopback redirect — the same flow, OAuth client, and scopes
the open-source Gemini CLI uses, so a Google Workspace org that allows Gemini
CLI allows this identically). MetoAI never sees the password; it receives
tokens on 127.0.0.1, stores the refresh token encrypted for the current
Windows user (DPAPI), and calls Google's Code Assist `generateContent` API,
where enterprise Gemini licensing resolves from the account alone. Requests
are translated to Gemini's native format and back, so tool calling, vision,
and every hard guardrail (read-only mailbox, one-shot unsent drafts, no send
capability) apply unchanged. An existing Gemini CLI sign-in on the machine is
honored automatically as a fallback. If your organization's Gemini license
designates a Google Cloud project, enter that project id in the **Google
Cloud project** field in Settings (the same value for everyone in the
organization; the `GOOGLE_CLOUD_PROJECT` environment variable also works as
a fallback). The Gemini CLI's other environment settings are not needed:
`NODE_OPTIONS=--use-system-ca` exists because Node.js does not trust the
Windows certificate store by default, while MetoAI uses it natively (so
corporate TLS inspection just works), and `GEMINI_TELEMETRY_ENABLED`
controls CLI telemetry, which MetoAI does not have. Note that with any
cloud provider, email content leaves your machine — the local-endpoint
setup keeps everything on-device.

Context budgets are provider-aware: in Gemini mode every text budget is
multiplied by four — attachments carry up to 80,000 extracted characters
each (192,000 per message), email bodies up to 96,000 characters, inline
selected-email text 24,000, external documents 80,000 each (192,000
total), and tool results up to 480,000 characters — because Gemini models
have ~1M-token context windows. Local endpoints keep the standard
budgets, and capability caps (ten emails, ten attachments, three external
documents, tool rounds, the 25 MB intake) never scale.

Gemini responses are speed-tuned: internal "thinking" is disabled on
`gemini-2.5-flash` and `flash-lite` and floored at the model minimum on
`pro` (thinking is the dominant latency cost on ordinary mailbox
questions), responses stream into the chat token by token, the selected
email's bounded body is inlined into the first request so common questions
answer in a single model round instead of a tool round trip, and the
system prompt asks for concise answers. `gemini-2.5-flash` is the
recommended model: near-pro quality on mail tasks with far lower latency
and higher quotas.

If a request fails, the sidebar shows diagnostic identifiers such as:

```text
HTTP_401_UNAUTHORIZED
HTTP_400_BAD_REQUEST
NETWORK_CONNECT_FAILURE
NETWORK_NAME_RESOLUTION
TLS_SECURE_CHANNEL_FAILURE
AI_TIMEOUT
RESPONSE_INVALID_JSON
RESPONSE_MISSING_CONTENT
TOOL_ROUND_LIMIT
DRAFT_PERMISSION_NOT_AVAILABLE
DRAFT_UPDATE_NOT_AVAILABLE
DRAFT_ALREADY_LINKED
DRAFT_TOOL_MUST_BE_EXCLUSIVE
DRAFT_CREATION_FAILED
DRAFT_UPDATE_FAILED
TONE_SAMPLES_INSUFFICIENT
TONE_ANALYSIS_FAILED
EXTERNAL_CONTEXT_FAILED
OUTLOOK_COM_0x800...
```

For HTTP failures it also shows the provider error message/code, request ID when
present, and a bounded response excerpt. The local diagnostic log records the
operation, exception type, diagnostic code, and HRESULT without email content,
prompts, endpoint responses, or API keys.

For connection failures, the sidebar also shows the target host and port,
exception chain, HRESULT, `WebExceptionStatus`, and Windows socket/native error
when available. This distinguishes DNS, connection refusal, proxy, TLS, and
timeout failures before the endpoint returns an HTTP response.

## Remove

1. Close Outlook.
2. Open Windows **Installed apps** or **Apps & features**.
3. Uninstall **MetoAI**.

Endpoint settings remain under:

```text
%LOCALAPPDATA%\OutlookLocalAIChat
```

Delete that folder manually if you also want to remove the encrypted API key and
local diagnostic log.

## Build

Requirements:

- Windows 10 or newer
- Visual Studio 2022 Build Tools with .NET Framework 4.8 targeting pack
- Inno Setup 6

Build the assembly and tests:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\Restore-StrongNameKey.ps1
msbuild OutlookLocalAIChat.sln /m /p:Configuration=Release
tests\GuardrailTests\bin\Release\GuardrailTests.exe
powershell -ExecutionPolicy Bypass -File scripts\Test-Guardrails.ps1
```

The repository stores the stable strong-name key as Base64 so local and CI builds
use the same COM identity. A strong name is an assembly identity mechanism, not a
trusted publisher signature.

Build the installer:

```powershell
& "${env:ProgramFiles(x86)}\Inno Setup 6\ISCC.exe" `
  "installer\OutlookLocalAIChat.iss"
```

The installer is written to:

```text
artifacts\OutlookLocalAIChatSetup.exe
```

GitHub Actions builds, smoke-tests, and publishes the same single-file installer.
Every push to `main` updates the **Latest** GitHub release so the install link
above always points at the newest build.

## Compatibility

- Classic Outlook for Windows
- Microsoft Office Professional Plus 2021
- 32-bit or 64-bit Office on Windows
- .NET Framework 4.8
- OpenAI-compatible endpoint and model with tool-calling support
- HTTPS, loopback HTTP, or explicitly enabled remote HTTP

The new Outlook for Windows does not load COM add-ins.
