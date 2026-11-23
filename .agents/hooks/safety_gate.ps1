# Hotel Master PreToolUse Safety Gate Hook (Pure ASCII for Windows PowerShell 5.1 compatibility)
$ErrorActionPreference = "SilentlyContinue"
try {
    $raw = [Console]::In.ReadToEnd()
    if ([string]::IsNullOrWhiteSpace($raw)) {
        Write-Output '{"decision":"allow"}'
        exit 0
    }
    $payload = $raw | ConvertFrom-Json
    $cmd = ""
    if ($payload.toolCall -and $payload.toolCall.args -and $payload.toolCall.args.CommandLine) {
        $cmd = [string]$payload.toolCall.args.CommandLine
    }

    # 1. Block blanket git add (. or -A or --all)
    if ($cmd -match '(?i)\bgit\s+add\s+(\.|-A|--all)(\s|;|$)') {
        Write-Output '{"decision":"deny","reason":"[Hotel Master Safety Gate] Blocked blanket git add (. / -A). Use explicit file allowlist per safe-deploy-streamlit skill."}'
        exit 0
    }

    # 2. Block destructive SQL statements against JinXu PMS (HT2000) tables
    if ($cmd -match '(?i)\b(DELETE\s+FROM|UPDATE|INSERT\s+INTO|DROP\s+TABLE|TRUNCATE\s+TABLE)\s+HT[A-Z0-9_]+') {
        Write-Output '{"decision":"deny","reason":"[Hotel Master Safety Gate] Blocked non-SELECT SQL against JinXu PMS (HT2000). PMS access is strictly READ-ONLY."}'
        exit 0
    }

    # 3. Block staging secrets.toml or credential JSONs into git
    if ($cmd -match '(?i)\bgit\s+add\b.*(secrets\.toml|rodersplusapp-|secrets_dump\.json)') {
        Write-Output '{"decision":"deny","reason":"[Hotel Master Safety Gate] Blocked staging credential or secrets file into git."}'
        exit 0
    }

    Write-Output '{"decision":"allow"}'
} catch {
    Write-Output '{"decision":"allow"}'
}
