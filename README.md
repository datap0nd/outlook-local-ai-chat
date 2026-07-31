# MailAI

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
3. Run the installer for your Windows account.
4. Start classic Outlook.
5. Choose **MailAI > MailAI** on the ribbon.
6. Open **Settings** and enter:
   - the OpenAI-compatible endpoint or base URL;
   - the model name, with `qwen3.5-35b-a3b` recommended;
   - the API key.
   - **Allow insecure HTTP** only when a non-local endpoint uses plain HTTP.
7. Choose **Check endpoint**. Save only after authentication, the selected
   model, and mailbox tool calling pass.

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

## Model choice

The default is `qwen3.5-35b-a3b`. For mailbox search, context selection, and
drafting, it is the best balance of tool-use quality, response speed, and server
load among the commonly exposed options supported by this project.

- `qwen3.5-35b-a3b`: recommended balanced default.
- `gpt-oss-120b`: quality-first fallback when the server can absorb higher
  latency and load.
- `gpt-oss-20b`: speed-first fallback.
- Other editable model IDs remain available when they support OpenAI-compatible
  chat tool calls.

The add-in does not recommend Gauss or Gausso variants. It also excludes
embedding-only models from endpoint discovery. **Check endpoint** makes a
harmless configuration request and requires the selected model to return one of
the add-in's read-only mailbox tool calls. It does not execute that tool or read
email during the check.

## Use

1. Open **MailAI**. The chat appears as a sidebar inside Outlook.
   You can also right-click one email and choose **Send to MailAI**. The
   sidebar opens with `Selected: subject` at the top. Common `RE:`, `FW:`, and
   `FWD:` prefixes are hidden in that display.
2. Ask a mailbox question, such as "What did I agree to send this week?"
3. The model can search Inbox and Sent Items, inspect selected results, and load
   a conversation thread when needed.
4. The sidebar records which bounded context operations ran.
5. Ask explicitly, for example "create a reply draft" or "write an email."
   Local code recognizes that drafting intent and exposes one creation attempt
   for that request. The draft opens unsent in Outlook.
6. A mailbox question without explicit drafting language cannot expose draft
   creation. Loaded email text and model output cannot authorize it.
7. The same Outlook draft stays linked to that chat. Follow-up instructions such
   as "make it shorter" or "bold the deadline" update and redisplay that exact
   unsent item. No second draft is created.
8. Review, edit, address, and send the message using Outlook's normal editor.

Selecting an email is optional for mailbox questions. When one is selected, the
model receives its metadata and may request its body using the temporary
`selected` handle. The conversation remains in memory until **New chat** is
chosen or Outlook closes.

## Hard security boundary

The model is not given general Outlook access. Draft creation is a narrowly
scoped exception, not a general mutation permission.

- Every request exposes exactly three read-only tools: `search_mailbox`,
  `read_messages`, and `read_thread`.
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
- Model output is length-limited plain text displayed in a Windows control. It is
  never evaluated, executed, or rendered as HTML.
- Only a one-request authorization derived locally from explicit user drafting
  intent can create the linked draft. Later revisions require both recognized
  revision intent and that local linked-draft session.
- The model-invoked mailbox host remains read-only. A separate draft host accepts
  only the bounded `create_draft` and `update_draft` operations.
- The draft path exposes no send, move, delete, schedule, BCC, arbitrary HTML,
  or mailbox traversal operation.
- Classic Outlook COM has no permission-manifest switch for sending. MailAI
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

- selected email metadata, when an email is selected;
- up to 12 recent chat turns;
- the current prompt.

The model may then request:

- up to 20 bounded result summaries from the primary Inbox and Sent Items;
- up to four bounded message bodies per tool call;
- up to 12 messages from one Outlook conversation;
- at most four tool calls per round and four context-retrieval rounds.

When the latest user prompt explicitly asks to create or open a draft, local
intent rules expose `create_draft` for that request. Its bounded arguments may
contain a new-message subject,
recipients, CC recipients, and body, or a reply body plus the exact temporary
handle of a searched or selected source message. The tool can only save and
display one unsent Outlook draft. While that draft is
linked, a recognized revision request exposes `update_draft` instead. Each
update supplies the complete bounded plain-text body and optional exact phrases to bold. The local
formatter HTML-encodes all text and inserts only fixed `<strong>` and line-break
markup. If a model mistakenly returns paired Markdown bold markers, the local
formatter removes the markers and applies real Outlook bold formatting before
display. Arbitrary model HTML is never accepted. Neither tool has a send operation.

Email bodies are sent only when the model requests them through an approved
read-only tool. The add-in does not index, upload, or transmit the entire
mailbox automatically.

The optional Settings check sends the API key to `GET /v1/models` when that route
is available, then submits a synthetic chat request that contains no mailbox
data. A successful check proves that authentication, the entered model, and the
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
available, it populates the editable model list in Settings.

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
3. Uninstall **MailAI**.

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

## Compatibility

- Classic Outlook for Windows
- Microsoft Office Professional Plus 2021
- 32-bit or 64-bit Office on Windows
- .NET Framework 4.8
- OpenAI-compatible endpoint and model with tool-calling support
- HTTPS, loopback HTTP, or explicitly enabled remote HTTP

The new Outlook for Windows does not load COM add-ins.
