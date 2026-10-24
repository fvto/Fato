$ErrorActionPreference = 'Stop'
$oldDashboardRoot = (Resolve-Path -LiteralPath 'C:\Users\User\Documents\MQAA-Record').Path.TrimEnd('\')
$documentsParent = (Resolve-Path -LiteralPath 'C:\Users\User\Documents').Path.TrimEnd('\')
$newDashboardRoot = [IO.Path]::GetFullPath('C:\Users\User\Documents\Chingluh Dashboard').TrimEnd('\')
if ((Split-Path -Parent $newDashboardRoot) -ne $documentsParent) { throw 'Destination is outside Documents' }
if (Test-Path -LiteralPath $newDashboardRoot) { throw 'Destination already exists; do not overwrite it' }
function Assert-ChildPath([string]$root, [string]$target) {
    $resolved = [IO.Path]::GetFullPath($target)
    if (-not $resolved.StartsWith($root + '\', [StringComparison]::OrdinalIgnoreCase)) { throw "Unsafe target: $resolved" }
    return $resolved
}
$rootNames = @('MO_BANG_DIEU_KHIEN.bat','MO_SERVER.bat','mqaa_launcher.py','dashboard_paths.py','dashboard_config.json','sfc_dashboard_server.py','sfc_controller.py','sfc_command_jobs.py','sfc_commands.py','sfc_excel.py','sfc_hfpa.py','sfc_http.py','sfc_session_cookies.js','sfc_memory_crawler.py','sfc_memory_knowledge_base.json')
$moveEntries = @()
foreach ($name in $rootNames) {
    $destinationName = if ($name -eq 'mqaa_launcher.py') { 'dashboard_launcher.py' } else { $name }
    $moveEntries += [pscustomobject]@{ source=$name; destination=$destinationName }
}
foreach ($name in @('templates','tests','Report\HFPAExports','scratch\backup_nike_20261024_093431','scratch\backup_sfc_command_20261024_094337')) {
    $moveEntries += [pscustomobject]@{ source=$name; destination=$name }
}
$scratchNames = @('crawl_sfc_links.py','inspect_flowchart.py','inspect_qmsmain.py','inspect_nav.py','inspect_itemclick.py','test_factory_click.py','test_fast_switch.py','test_switch_all_factories.py','test_sse_concept.py','test_stream_endpoint.py','test_nlp_parser.py','test_urls.py','test_cbo_detail.py','start_chrome_clean.py','flowchart_preview.png','qmsmain_preview.png','today_export_response.bin','migrate_dashboard_project.ps1')
foreach ($item in Get-ChildItem -LiteralPath (Join-Path $oldDashboardRoot 'scratch') -File) {
    if ($item.Name -match '^(dashboard_|diagnose_sfc_|inspect_date_|inspect_factory_|inspect_menu_|inspect_sfc_|narrow_live_scope_|run_live_|sfc_|test_http_|test_dashboard_|test_sfc_|verify_sfc_|verify_vh_)' -or $item.Name -in $scratchNames) {
        $moveEntries += [pscustomobject]@{ source=('scratch\'+$item.Name); destination=('scratch\'+$item.Name) }
    }
}
foreach ($entry in $moveEntries) {
    $sourcePath = Assert-ChildPath $oldDashboardRoot (Join-Path $oldDashboardRoot $entry.source)
    $destinationPath = Assert-ChildPath $newDashboardRoot (Join-Path $newDashboardRoot $entry.destination)
    if (-not (Test-Path -LiteralPath $sourcePath)) { throw "Missing source: $($entry.source)" }
    if (Test-Path -LiteralPath $destinationPath) { throw "Destination collision: $($entry.destination)" }
}
$exportRoot = Assert-ChildPath $oldDashboardRoot (Join-Path $oldDashboardRoot 'Report\HFPAExports')
$exportHashes = @(Get-ChildItem -LiteralPath $exportRoot -Recurse -File | Where-Object { $_.Extension -in @('.xls','.xlsx') } | ForEach-Object {
    [pscustomobject]@{ path=$_.FullName.Substring($oldDashboardRoot.Length+1); sha256=(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash }
})
$legacyHashes = @('auto_scan_all_21_models.py','generate_master_comparison.py','TU_DONG_QUET_MQAA.bat','MQAA_Master_DoiChieu_21_Forms_MoiNhat.xlsx','MQAA_Master_DoiChieu_21_Forms.xlsx') | ForEach-Object {
    $legacyPath = Join-Path $oldDashboardRoot $_
    [pscustomobject]@{ path=$_; sha256=(Get-FileHash -LiteralPath $legacyPath -Algorithm SHA256).Hash }
}
$liveStatus = Invoke-RestMethod 'http://localhost:5050/api/status'
if ($liveStatus.command_running -eq $true) { throw 'A Dashboard export is active; do not stop it' }
$serverProcessId = (Get-NetTCPConnection -LocalPort 5050 -State Listen | Select-Object -First 1).OwningProcess
$serverProcess = Get-CimInstance Win32_Process -Filter "ProcessId=$serverProcessId"
if ($serverProcess.CommandLine -notlike "*$oldDashboardRoot*sfc_dashboard_server.py*") { throw 'Unexpected server; do not stop it' }
Stop-Process -Id $serverProcessId
New-Item -ItemType Directory -Path $newDashboardRoot | Out-Null
$record = [ordered]@{ source_root=$oldDashboardRoot; destination_root=$newDashboardRoot; moved_at=(Get-Date -Format o); moves=$moveEntries; excel_files=$exportHashes; legacy_files=$legacyHashes; state='moving' }
$recordPath = Join-Path $newDashboardRoot 'migration_manifest.json'
$record | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $recordPath -Encoding utf8
foreach ($entry in $moveEntries) {
    $sourcePath = Assert-ChildPath $oldDashboardRoot (Join-Path $oldDashboardRoot $entry.source)
    $destinationPath = Assert-ChildPath $newDashboardRoot (Join-Path $newDashboardRoot $entry.destination)
    $parentPath = Split-Path -Parent $destinationPath
    if (-not (Test-Path -LiteralPath $parentPath)) { New-Item -ItemType Directory -Path $parentPath | Out-Null }
    Move-Item -LiteralPath $sourcePath -Destination $destinationPath
}
$newRules = Assert-ChildPath $newDashboardRoot (Join-Path $newDashboardRoot '.agents\rules')
New-Item -ItemType Directory -Path $newRules -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $oldDashboardRoot '.agents\rules\devexpress_and_audit_reconciliation.md') -Destination (Join-Path $newRules 'devexpress_and_audit_reconciliation.md')
# Python caches contain old source paths; regenerate them from moved source files.
$newTestCache = Assert-ChildPath $newDashboardRoot (Join-Path $newDashboardRoot 'tests\__pycache__')
if (Test-Path -LiteralPath $newTestCache) { Remove-Item -LiteralPath $newTestCache -Recurse -Force }
foreach ($cache in Get-ChildItem -LiteralPath (Join-Path $oldDashboardRoot '__pycache__') -File) {
    if ($cache.Name -match '^(sfc_|mqaa_launcher\.|dashboard_paths\.)') {
        $cachePath = Assert-ChildPath $oldDashboardRoot $cache.FullName
        Remove-Item -LiteralPath $cachePath -Force
    }
}
foreach ($item in $exportHashes) {
    $newFile = Assert-ChildPath $newDashboardRoot (Join-Path $newDashboardRoot $item.path)
    if ((Get-FileHash -LiteralPath $newFile -Algorithm SHA256).Hash -ne $item.sha256) { throw "Excel changed: $($item.path)" }
}
foreach ($item in $legacyHashes) {
    if ((Get-FileHash -LiteralPath (Join-Path $oldDashboardRoot $item.path) -Algorithm SHA256).Hash -ne $item.sha256) { throw "Legacy changed: $($item.path)" }
}
$record.state = 'moved_and_hash_verified'
$record | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $recordPath -Encoding utf8
Write-Output "Moved $($moveEntries.Count) code/support entries. Verified $($exportHashes.Count) Excel files unchanged. Legacy MQAA files retained."
