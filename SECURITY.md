# Security model

## Security objective

Untrusted email content, user prompts, conversation history, and model responses
must never reach an Outlook send or source-message mutation capability. Model
tool calls may select bounded read-only mailbox context. The only mutation this
add-in permits is creating an unsent draft after a direct user click.

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

User clicks draft button
    |
    v
DraftService -> Save + Display unsent Outlook draft
```

`OpenAiCompatibleClient` has no reference to the Outlook application object or
`DraftService`. `DraftService` is instantiated only in the two draft-button event
handlers.

## Enforced invariants

1. The model request schema exposes exactly `search_mailbox`, `read_messages`,
   and `read_thread`.
2. `MailboxToolHost` has one public dispatcher and rejects any tool name outside
   that compile-time allowlist.
3. Model-selected searches are limited to the primary Inbox and Sent Items.
   Results, body lengths, thread lengths, calls per round, and tool rounds are
   capped.
4. Search results receive temporary handles. Read operations accept only handles
   issued within the current request, plus the optional `selected` handle.
5. The mailbox host has no reference to `DraftService`, and the endpoint client
   has no Outlook application object.
6. Response text is stripped of control characters and truncated before display
   or drafting.
7. A `RichTextBox` displays the response as literal text. No browser or HTML
   renderer is used.
8. `DraftService` has exactly two public operations:
   `CreateReplyDraft` and `CreateNewDraft`.
9. Draft operations call Outlook save and display behavior only.
10. Source scans fail on Outlook send, delete, move, Outbox, or send/receive
   capabilities.
11. Conversation history is held in memory and cleared by **New chat** or Outlook
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

The key is sent only in the HTTPS Authorization header of the configured endpoint.
Loopback HTTP is permitted for local model servers.

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
