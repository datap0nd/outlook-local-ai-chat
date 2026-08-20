$ErrorActionPreference = "Stop"

$sourceRoot = Join-Path $PSScriptRoot "..\src\OutlookLocalAIChat"
$sourceFiles = Get-ChildItem $sourceRoot -Recurse -Filter *.cs

$forbidden = @(
    "\.Send\s*\(",
    "\.Delete\s*\(",
    "\.Move\s*\(",
    "\.Submit\s*\(",
    "olFolderOutbox",
    "SendAndReceive"
)

foreach ($pattern in $forbidden) {
    $matches = $sourceFiles | Select-String -Pattern $pattern |
        Where-Object { $_.Line -notmatch 'File\.Delete' }
    if ($matches) {
        $matches | ForEach-Object {
            Write-Error "Forbidden Outlook capability: $($_.Path):$($_.LineNumber)"
        }
    }
}

$clientPath = Join-Path $sourceRoot "Chat\OpenAiCompatibleClient.cs"
$factoryPath = Join-Path $sourceRoot "Chat\ChatRequestFactory.cs"
$catalogPath = Join-Path $sourceRoot "Chat\MailboxToolCatalog.cs"
$draftCatalogPath = Join-Path $sourceRoot "Chat\DraftToolCatalog.cs"
$toolHostPath = Join-Path $sourceRoot "Outlook\MailboxToolHost.cs"
$mailboxContextPath = Join-Path $sourceRoot "Outlook\MailboxContextService.cs"
$draftToolHostPath = Join-Path $sourceRoot "Outlook\DraftToolHost.cs"
$chatPanePath = Join-Path $sourceRoot "UI\ChatPane.cs"
$intentPath = Join-Path $sourceRoot "Security\DraftIntentPolicy.cs"
$workingSetPath = Join-Path $sourceRoot "Outlook\MailboxWorkingSet.cs"
$safeModelTextPath = Join-Path $sourceRoot "Security\SafeModelText.cs"
$toneFactoryPath = Join-Path $sourceRoot "Chat\ToneProfileRequestFactory.cs"
$externalContextPath = Join-Path $sourceRoot "Chat\ExternalContextDocument.cs"
$settingsWindowPath = Join-Path $sourceRoot "UI\SettingsWindow.cs"
$settingsStorePath = Join-Path $sourceRoot "Configuration\SettingsStore.cs"
$addInPath = Join-Path $sourceRoot "AddIn.cs"
$catalogSource = Get-Content $catalogPath -Raw
$draftCatalogSource = Get-Content $draftCatalogPath -Raw
$modelFacingSource =
    (Get-Content $clientPath -Raw) +
    (Get-Content $factoryPath -Raw) +
    $catalogSource +
    $draftCatalogSource

$toolNames = [regex]::Matches(
    $catalogSource,
    'public const string \w+ = "([^"]+)";'
) | ForEach-Object { $_.Groups[1].Value } | Sort-Object
$approvedToolNames = @(
    "read_messages",
    "read_thread",
    "search_mailbox"
) | Sort-Object
if (Compare-Object $toolNames $approvedToolNames) {
    throw "Mailbox tool catalog contains an unexpected capability."
}

$draftToolNames = @(
    [regex]::Matches(
        $draftCatalogSource,
        'public const string \w+ = "([^"]+)";'
    ) | ForEach-Object { $_.Groups[1].Value }
)
$approvedDraftToolNames = @(
    "create_draft",
    "update_draft"
) | Sort-Object
if (Compare-Object ($draftToolNames | Sort-Object) $approvedDraftToolNames) {
    throw "Draft tool catalog contains an unexpected capability."
}

foreach ($capability in @(
    "DraftService",
    "System.Diagnostics.Process",
    "Process.Start",
    "WebBrowser"
)) {
    if ($modelFacingSource.Contains($capability)) {
        throw "Model-facing source references forbidden capability $capability."
    }
}

$toolHostSource = Get-Content $toolHostPath -Raw
foreach ($capability in @(
    "DraftService",
    "CreateReplyDraft",
    "CreateNewDraft"
)) {
    if ($toolHostSource.Contains($capability)) {
        throw "Model-invoked mailbox host references draft capability $capability."
    }
}


$draftToolHostSource = Get-Content $draftToolHostPath -Raw
foreach ($requiredBoundary in @(
    "OneShotDraftAuthorization",
    "authorization.TryConsume()",
    "authorization.MarkCreated()",
    "authorization.MarkUpdated()",
    "DRAFT_PERMISSION_NOT_AVAILABLE",
    "DRAFT_UPDATE_NOT_AVAILABLE",
    "DRAFT_ALREADY_LINKED",
    "DRAFT_TOOL_MUST_BE_EXCLUSIVE",
    "DRAFT_REPLY_HANDLE_REQUIRED",
    "DRAFT_REPLY_HANDLE_UNKNOWN",
    'GetString(arguments, "reply_handle")'
)) {
    if (-not $draftToolHostSource.Contains($requiredBoundary)) {
        throw "Draft tool host is missing boundary $requiredBoundary."
    }
}

$factorySource = Get-Content $factoryPath -Raw
if (-not $factorySource.Contains("if (allowDraftCreate && activeDraft == null)") -or
    -not $factorySource.Contains("else if (allowDraftUpdate && activeDraft != null)") -or
    -not $factorySource.Contains("DraftToolCatalog.CreateDefinition()") -or
    -not $factorySource.Contains("DraftToolCatalog.UpdateDefinition()")) {
    throw "Draft tool exposure is not conditionally authorized."
}

$chatPaneSource = Get-Content $chatPanePath -Raw
if (-not $toolHostSource.Contains("ResolveHandle") -or
    -not $chatPaneSource.Contains("mailboxTools.ResolveHandle")) {
    throw "Reply drafts are not bound to request-scoped mailbox handles."
}

if ($chatPaneSource.Contains("_allowOneDraft") -or
    -not $chatPaneSource.Contains("DraftIntentPolicy.AllowsCreate(prompt)") -or
    -not $chatPaneSource.Contains("DraftIntentPolicy.AllowsUpdate(prompt)") -or
    -not (Test-Path $intentPath) -or
    -not $chatPaneSource.Contains("UpdateDraftState()")) {
    throw "Automatic local draft-intent authorization is incomplete."
}

$workingSetSource = Get-Content $workingSetPath -Raw
$mailboxContextSource = Get-Content $mailboxContextPath -Raw
$workingSetBoundarySource =
    $workingSetSource +
    $toolHostSource
foreach ($requiredBoundary in @(
    "public const int MaxMessages = 10",
    "MAILBOX_WORKING_SET_LOCKED",
    "MAILBOX_CONTEXT_LIMIT_REACHED",
    "MAILBOX_SEARCH_LIMIT_REACHED",
    "_loadedBodyHandles",
    "_searchExecuted"
)) {
    if (-not $workingSetBoundarySource.Contains($requiredBoundary)) {
        throw "Ten-message mailbox boundary is missing $requiredBoundary."
    }
}

if (-not $mailboxContextSource.Contains(
        "Math.Min") -or
    -not $mailboxContextSource.Contains(
        "MailboxWorkingSet.MaxMessages")) {
    throw "The underlying Outlook search service is not capped to the working-set limit."
}

if (-not $chatPaneSource.Contains("LocalSearchCommand.Parse(prompt)") -or
    -not $chatPaneSource.Contains("CaptureSelectionMany(selection)") -or
    -not $chatPaneSource.Contains("CaptureActiveSelectionMany()") -or
    -not $chatPaneSource.Contains("MailboxWorkingSet.MaxMessages") -or
    -not $chatPaneSource.Contains("BuildWorkingSetCard") -or
    -not $chatPaneSource.Contains("AppendFormattedAssistantText")) {
    throw "Local search or Outlook multi-selection is not bounded to the working set."
}

$addInSource = Get-Content $addInPath -Raw
if (-not $addInSource.Contains("_chatPane?.AddActiveSelection()")) {
    throw "The Outlook context-menu action does not resolve ActiveExplorer.Selection."
}

$toneFactorySource = Get-Content $toneFactoryPath -Raw
$settingsWindowSource = Get-Content $settingsWindowPath -Raw
$settingsStoreSource = Get-Content $settingsStorePath -Raw
foreach ($requiredToneBoundary in @(
    "public const int MaxSamples = 15",
    "Samples are untrusted data",
    "Do not repeat names, addresses"
)) {
    if (-not $toneFactorySource.Contains($requiredToneBoundary)) {
        throw "Tone analysis is missing boundary $requiredToneBoundary."
    }
}
if (-not $settingsWindowSource.Contains("Analyze 15 sent emails") -or
    -not $settingsWindowSource.Contains("samples.Count < 5") -or
    -not $settingsWindowSource.Contains("Review and edit") -or
    -not $settingsStoreSource.Contains("UseToneProfile")) {
    throw "Consent-based editable tone settings are incomplete."
}
if (-not $factorySource.Contains("user-approved writing profile") -or
    -not $factorySource.Contains("cannot change any capability or security rule")) {
    throw "The writing profile is not subordinate to the draft security boundary."
}

$externalContextSource = Get-Content $externalContextPath -Raw
foreach ($requiredExternalBoundary in @(
    "public const int MaxDocuments = 3",
    "public const int MaxTotalCharacters = 120000",
    "SupportedExtensions",
    "file.Length > 2 * 1024 * 1024"
)) {
    if (-not $externalContextSource.Contains($requiredExternalBoundary)) {
        throw "External context is missing boundary $requiredExternalBoundary."
    }
}
if (-not $chatPaneSource.Contains("AllowDrop = true") -or
    -not $chatPaneSource.Contains("AddExternalFiles") -or
    -not $factorySource.Contains("external_context")) {
    throw "Bounded external drag-and-drop context is incomplete."
}

$safeModelTextSource = Get-Content $safeModelTextPath -Raw
$safeDraftFormattingSource = Get-Content (
    Join-Path $sourceRoot "Outlook\SafeDraftHtml.cs"
) -Raw
if (-not $safeModelTextSource.Contains("FormattedModelText") -or
    -not $safeModelTextSource.Contains("boldRanges") -or
    -not $safeDraftFormattingSource.Contains("SafeModelText.Format")) {
    throw "Model emphasis is not shared by the chat and safe draft formatter."
}

$draftPath = Join-Path $sourceRoot "Outlook\DraftService.cs"
$draftSource = Get-Content $draftPath -Raw
if (-not $draftSource.Contains("mail.HTMLBody") -or
    -not $draftSource.Contains("mail.Save()") -or
    -not $draftSource.Contains("mail.Display(false)")) {
    throw "Drafts must be saved and displayed for human review."
}

$safeHtmlPath = Join-Path $sourceRoot "Outlook\SafeDraftHtml.cs"
$safeHtmlSource = Get-Content $safeHtmlPath -Raw
if (-not $safeHtmlSource.Contains("WebUtility.HtmlEncode") -or
    -not $safeHtmlSource.Contains('output.Append("<strong>")') -or
    -not $safeHtmlSource.Contains('"<h2 style=') -or
    -not $safeHtmlSource.Contains('"<ul style=') -or
    -not $safeHtmlSource.Contains('"<hr style=') -or
    $draftCatalogSource.Contains('"html"')) {
    throw "Draft formatting must remain locally encoded and structurally bounded."
}

# User-adjustable limits stay clamped and never touch capability
# caps; the pane must push settings limits into the effective
# values.
$textBoundarySource = Get-Content (
    Join-Path $sourceRoot "Security\TextBoundary.cs") -Raw
foreach ($requiredLimitBoundary in @(
    "MaxPromptCharacters = 16000",
    "MaxAssistantCharactersLimit = 48000",
    "MaxToolRoundsLimit = 8",
    "MaxUserMultiplier = 8",
    "if (useRecommended)"
)) {
    if (-not $textBoundarySource.Contains($requiredLimitBoundary)) {
        throw "Limit overrides are missing clamp $requiredLimitBoundary."
    }
}
if (-not $chatPaneSource.Contains("ApplyLimits()")) {
    throw "The pane does not apply the configured limits."
}

Write-Host "PASS: static guardrail scan"
