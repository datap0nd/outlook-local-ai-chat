# Security model

## Security objective

Untrusted email content, user prompts, conversation history, and model responses
must never reach an Outlook send or source-message mutation capability. Model
tool calls may select bounded read-only mailbox context. The only mutation this
add-in permits is creating one unsent draft after explicit, request-scoped local
authorization and updating that same locally linked unsent item after user
feedback.

## Capability separation

```text
Prompt + optional selected-message metadata
    |
    v
OpenAiCompatibleClient -> messages + read-only tool schema -> endpoint
    |
    v
allowlisted tool call -> MailboxToolHost -> bounded local Outlook reads
    |
    v
temporary handles + bounded untrusted text -> endpoint
    |
    v
bounded plain-text response -> Outlook custom task pane

User selects one-shot draft authorization before Send
    |
    v
request includes create_draft + exact reply handle -> DraftToolHost -> consume once
    |
    v
DraftService -> Save + Display one unsent Outlook draft

Later user feedback while the draft is linked
    |
    v
request includes update_draft -> DraftToolHost -> consume once
    |
    v
SafeDraftHtml -> remove Markdown markers + encode text + apply bold -> Save + Display same item
```

`OpenAiCompatibleClient` has no reference to the Outlook application object or
`DraftService`. The mailbox host remains separate from `DraftService`. The
dedicated draft host can reach only the internal draft service. Creation requires
an atomic one-shot authorization created from the local checkbox. Updates are
available only while the host retains exactly one linked draft and are limited
to one mutation attempt per user request.

## Enforced invariants

1. The model request schema always exposes exactly `search_mailbox`,
   `read_messages`, and `read_thread`. It exposes `create_draft` only when the
   local one-shot checkbox was selected for that request. Once a draft is linked,
   it exposes `update_draft` instead and never exposes both.
2. `MailboxToolHost` has one public dispatcher and rejects any tool name outside
   that compile-time allowlist.
3. Model-selected searches are limited to the primary Inbox and Sent Items.
   Results, body lengths, thread lengths, calls per round, and tool rounds are
   capped.
4. Search results receive temporary handles. Read operations accept only handles
   issued within the current request, plus the optional `selected` handle.
   Reply creation also requires one of those exact handles. Missing, expired,
   and fabricated handles are rejected without consuming draft permission, and
   the host never substitutes the selected or latest item.
5. The mailbox host has no reference to `DraftService`, and the endpoint client
   has no Outlook application object.
6. Response text is stripped of control characters and truncated before display
   or drafting.
7. A `RichTextBox` displays the response as literal text. No browser or HTML
   renderer is used.
8. `DraftService` and `DraftSession` are internal implementation types. The
   public draft host exposes state plus one guarded dispatcher and no send path.
9. `DraftToolHost` accepts only `create_draft` and `update_draft`, requires a
   draft operation to be the only tool call in that response, bounds every
   field, rejects unknown properties, and atomically consumes local permission
   before mutation.
10. A chat can link at most one draft. A request can make at most one creation
    or update attempt. Starting a new chat releases the COM link without deleting
    the unsent Outlook item.
11. Draft operations call Outlook save and display behavior only. Subject and
    recipient fields are bounded single-line text. Body input is plain text;
    local code removes paired Markdown bold markers, HTML-encodes the remaining
    text, and may add only fixed `<strong>` and `<br>` tags from those markers or
    at most 12 exact bold phrases. BCC and arbitrary HTML are not accepted.
12. Source scans fail on Outlook send, delete, move, Outbox, or send/receive
   capabilities.
13. Conversation history is held in memory and cleared by **New chat** or Outlook
    shutdown.

The system prompt reinforces these limits, but no security property depends on
the model obeying it.

## Secrets

The API key is encrypted with Windows Data Protection API using the current-user
scope. The encrypted value is stored in:

```text
%LOCALAPPDATA%\OutlookLocalAIChat\settings.json
```

Any process running as the same Windows user can potentially invoke DPAPI and
recover current-user secrets. This protects the key at rest from casual file
inspection, not from a compromised user session.

The key is sent in the Authorization header of the configured endpoint. HTTPS
protects that header and submitted mailbox context in transit. Loopback HTTP is
permitted automatically for local model servers. Non-local HTTP requires an
explicit persisted opt-in in Settings and displays a warning because the API
key, prompts, and retrieved email context are then sent without transport
encryption.

The Settings endpoint check may send the same Authorization header to
`GET /v1/models`. It then submits a synthetic tool-call request containing no
selected-message metadata, email bodies, or mailbox search results. The returned
tool call is validated but never executed.

## Logging

The diagnostic log records UTC time, operation name, exception type, diagnostic
code, and HRESULT category. It does not record email content, prompts, provider
response bodies, endpoints, request IDs, or API keys.

## Installation trust

The repository does not contain a code-signing certificate. Unsigned installers
and assemblies can trigger Windows warnings and may be blocked by corporate
application-control policy.

For organizational distribution:

1. Build in a controlled Windows pipeline.
2. Sign the DLL and installer with the organization's trusted code-signing
   certificate.
3. Publish hashes and retain build provenance.
4. Allowlist the publisher rather than a mutable file path.

## Out of scope

The design cannot guarantee safety if:

- the installed binary or registry entries are replaced;
- the Windows account is compromised;
- Outlook, .NET Framework, or Windows has an exploitable vulnerability;
- another Outlook add-in modifies the draft after creation;
- the configured AI endpoint mishandles or retains submitted data.

Review the endpoint provider's privacy, retention, and data-residency controls
before using real work email.
