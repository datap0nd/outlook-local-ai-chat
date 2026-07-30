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
Inbox and Sent Items. The user can then create an Outlook reply draft or
blank-addressed new-message draft from text they explicitly choose.

Success means installation is understandable, configuration takes one endpoint,
model name, and API key, and no model response can invoke an Outlook send action.

## Positioning

Model output can invoke only a compile-time allowlist of bounded mailbox read
tools. It cannot invoke Outlook mutation or arbitrary COM capabilities. Only
user interface events can call a separate Outlook draft service, and that service
exposes create, save, and display operations but no send operation.

## Operating Context

- Microsoft Office Professional Plus 2021 with classic Outlook on Windows.
- Per-user local installation is preferred.
- The user presses an AI Chat ribbon button and works in a right-docked Outlook
  Custom Task Pane. Selecting an email is optional, but makes that message
  available through a temporary read handle.
- Configuration is stored for the current Windows user. The API key is encrypted
  with Windows Data Protection API.
- Conversations are kept in memory and disappear when Outlook closes or the user
  starts a new chat.

## Capabilities and Constraints

- Search and read bounded context from the primary Inbox and Sent Items.
- Hold a text conversation about the mailbox, a selected message, or a retrieved
  conversation.
- Generate text suitable for a reply or a new message.
- Create and display an unsent Outlook draft only after an explicit user click.
- Never send, schedule, move, delete, mark, categorize, or modify the source email.
- Expose only `search_mailbox`, `read_messages`, and `read_thread` tools.
- Reject all other model tool calls and cap calls, rounds, results, and returned
  text.
- Never render model output as HTML or execute it as code.
- Support an OpenAI-compatible `/v1/chat/completions` endpoint.
- Permit HTTPS endpoints and loopback HTTP endpoints for local model servers.
- Permit non-local HTTP only after an explicit persisted opt-in that warns the
  user that the API key and mailbox context will travel without encryption.
- Target both 32-bit and 64-bit Office from one installer when practical.
- A production installer should be code-signed. Signing credentials are not
  included in the repository.

## Brand Commitments

The product name is Outlook Local AI Chat. The interface should feel like a
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
