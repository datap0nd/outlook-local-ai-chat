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
    $matches = $sourceFiles | Select-String -Pattern $pattern
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
$draftToolHostPath = Join-Path $sourceRoot "Outlook\DraftToolHost.cs"
$chatPanePath = Join-Path $sourceRoot "UI\ChatPane.cs"
$intentPath = Join-Path $sourceRoot "Security\DraftIntentPolicy.cs"
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
    -not $safeHtmlSource.Contains('html.Append("<strong>")') -or
    $draftCatalogSource.Contains('"html"')) {
    throw "Draft formatting must remain locally encoded and structurally bounded."
}

Write-Host "PASS: static guardrail scan"
