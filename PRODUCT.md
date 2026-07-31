# Product

<!-- impeccable:product-schema 1 -->

## Platform

windows

## Stack

Delegated: C# on .NET Framework 4.8 as a classic Outlook COM add-in, with a
Windows Forms control hosted in an Outlook Custom Task Pane and a single-file
Windows installer. This targets
Microsoft Office Professional Plus 2021 on Windows without Microsoft 365 add-in
deployment.

## Users

The primary user works in classic Outlook on a managed Windows work PC and wants
to ask questions across their local mailbox, let the model retrieve only relevant
context, refine a response through conversation, and open an unsent draft for
final review.

## Product Purpose

The add-in provides a native chat sidebar inside Outlook. It sends the prompt,
recent in-memory conversation, and optional selected-message metadata to a
user-configured OpenAI-compatible endpoint. The model may request bounded
read-only searches, message bodies, and conversation threads from the primary
Inbox and Sent Items. After the user explicitly arms one request, the model may
create one unsent Outlook reply or new-message draft. That visible Outlook item
then remains linked to the chat so later feedback updates the same draft.

Success means installation is understandable, configuration takes one endpoint,
model name, and API key, and no model response can invoke an Outlook send action.

## Positioning

Model output can invoke only a compile-time allowlist of bounded mailbox read
tools. `create_draft` appears only when local code recognizes explicit drafting
intent in the latest user-written prompt.
After creation it is replaced by `update_draft`, which can modify only the one
locally linked item. The dedicated host exposes no send operation.

## Operating Context

- Microsoft Office Professional Plus 2021 with classic Outlook on Windows.
- Per-user local installation is preferred.
- The user opens MailAI from the ribbon or right-clicks one email and chooses
  **Send to MailAI**, then works in a right-docked Outlook Custom Task Pane.
  Selecting an email is optional, but makes that message available through a
  temporary read handle.
- Configuration is stored for the current Windows user. The API key is encrypted
  with Windows Data Protection API.
- Conversations are kept in memory and disappear when Outlook closes or the user
  starts a new chat.

## Capabilities and Constraints

- Search and read bounded context from the primary Inbox and Sent Items.
- Hold a text conversation about the mailbox, a selected message, or a retrieved
  conversation.
- Generate text suitable for a reply or a new message.
- Create and display at most one unsent Outlook draft per chat, then update that
  same item at most once per later user request.
- Bind reply drafts to the exact temporary handle returned for the searched or
  selected source message. Never fall back to the latest mailbox item.
- Never send, schedule, move, delete, mark, categorize, or modify the source email.
- Always expose only `search_mailbox`, `read_messages`, and `read_thread`.
  Conditionally expose `create_draft` for a locally recognized drafting request
  or `update_draft` for a locally recognized revision of the one linked draft,
  never both.
- Reject all other model tool calls and cap calls, rounds, results, and returned
  text.
- Never render model output as HTML or execute it as code. Draft formatting
  accepts only plain text plus exact bold phrases and is HTML-encoded locally.
  Paired Markdown bold markers are removed and converted to real local bold
  formatting if a compatible model returns them anyway.
- Support an OpenAI-compatible `/v1/chat/completions` endpoint.
- Recommend `qwen3.5-35b-a3b` as the balanced default while preserving editable
  model identifiers and quality-first or speed-first fallbacks.
- Verify authentication, optional `/v1/models` discovery, and actual read-only
  tool-call compatibility from Settings without loading mailbox data.
- Permit HTTPS endpoints and loopback HTTP endpoints for local model servers.
- Permit non-local HTTP only after an explicit persisted opt-in that warns the
  user that the API key and mailbox context will travel without encryption.
- Target both 32-bit and 64-bit Office from one installer when practical.
- A production installer should be code-signed. Signing credentials are not
  included in the repository.

## Brand Commitments

The user-facing product name is MailAI. The stable COM assembly, ProgID,
CLSID, settings path, installer filename, and repository name retain the
`OutlookLocalAIChat` technical identity so upgrades do not break. The UI should feel like a
restrained Windows productivity utility, not an AI showcase. Language must be
direct, calm, and explicit about what data is read and when a draft is created.

## Evidence on Hand

The product brief is the user's requested workflow. There are no approved company
logos, claims, screenshots, or signing certificates, and future work must not
fabricate them.

## Product Principles

- Capabilities, not prompts, define the security boundary.
- Nothing leaves the active conversation and model-selected bounded read context.
- Drafting always ends in Outlook's normal editor with the user in control.
- Local configuration should be inspectable, reversible, and per-user.
- Familiar Windows behavior is more important than decorative novelty.

## Accessibility & Inclusion

The chat sidebar must support keyboard-only operation, visible focus, system text
scaling, high-contrast-compatible colors, and plain-language error recovery.
