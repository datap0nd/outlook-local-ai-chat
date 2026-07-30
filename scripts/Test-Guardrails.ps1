$ErrorActionPreference = "Stop"

$sourceRoot = Join-Path $PSScriptRoot "..\src\OutlookLocalAIChat"
$sourceFiles = Get-ChildItem $sourceRoot -Recurse -Filter *.cs

$forbidden = @(
    "\.Send\s*\(",
    "\.Delete\s*\(",
    "\.Move\s*\(",
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
$toolHostPath = Join-Path $sourceRoot "Outlook\MailboxToolHost.cs"
$catalogSource = Get-Content $catalogPath -Raw
$modelFacingSource =
    (Get-Content $clientPath -Raw) +
    (Get-Content $factoryPath -Raw) +
    $catalogSource

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

$draftPath = Join-Path $sourceRoot "Outlook\DraftService.cs"
$draftSource = Get-Content $draftPath -Raw
if (-not $draftSource.Contains("replyMail.Save()") -or
    -not $draftSource.Contains("replyMail.Display(false)")) {
    throw "Reply draft must be saved and displayed for human review."
}

Write-Host "PASS: static guardrail scan"
