<# 
.SYNOPSIS
  Delegated Consents (oauth2PermissionGrants) → CSV + interaktiver HTML-Report
#>

param(
  [string]$AppNameContains
)

function Invoke-GraphPaged {
  param([Parameter(Mandatory)][string]$Uri)
  $headers = @{ "ConsistencyLevel" = "eventual" }
  $all = @(); $next = $Uri
  while ($null -ne $next) {
    $res = Invoke-MgGraphRequest -Method GET -Uri $next -Headers $headers
    if ($res.value) { $all += $res.value }
    $next = $res.'@odata.nextLink'
  }
  return $all
}
function Get-DirectoryObjectsByIds {
  param([string[]]$Ids)
  if (-not $Ids -or $Ids.Count -eq 0) { return @() }
  $results = @()
  for ($i=0; $i -lt $Ids.Count; $i+=1000) {
    $chunk = $Ids[$i..([Math]::Min($i+999, $Ids.Count-1))]
    $body = @{ ids = $chunk } | ConvertTo-Json
    $res = Invoke-MgGraphRequest -Method POST `
      -Uri "https://graph.microsoft.com/beta/directoryObjects/getByIds" `
      -Body $body -ContentType "application/json"
    if ($res.value) { $results += $res.value }
  }
  return $results
}

try { $null = Get-MgContext -ErrorAction Stop } catch {
  Connect-MgGraph -Scopes "Directory.Read.All","Application.Read.All" | Out-Null
}

$baseUri = "https://graph.microsoft.com/beta/oauth2PermissionGrants`?$count=true&$top=999"
$grants = Invoke-GraphPaged -Uri $baseUri
if (-not $grants) { Write-Host "Keine oauth2PermissionGrants gefunden."; return }

$clientIds   = [System.Collections.Generic.HashSet[string]]::new()
$resourceIds = [System.Collections.Generic.HashSet[string]]::new()
$userIds     = [System.Collections.Generic.HashSet[string]]::new()
foreach ($g in $grants) {
  if ($g.clientId)   { [void]$clientIds.Add($g.clientId) }
  if ($g.resourceId) { [void]$resourceIds.Add($g.resourceId) }
  if ($g.consentType -eq "Principal" -and $g.principalId) { [void]$userIds.Add($g.principalId) }
}

$dirObjs = Get-DirectoryObjectsByIds -Ids @($clientIds + $resourceIds + $userIds)
$spById = @{}; $userById = @{}
foreach ($o in $dirObjs) {
  switch ($o.'@odata.type') {
    "#microsoft.graph.servicePrincipal" { $spById[$o.id] = @{ Id=$o.id; AppId=$o.appId; Name=$o.displayName } }
    "#microsoft.graph.user"            { $userById[$o.id] = @{ Id=$o.id; UPN=$o.userPrincipalName; Name=$o.displayName } }
  }
}

if ($AppNameContains) {
  $grants = $grants | Where-Object { $spById[$_.clientId].Name -like "*$AppNameContains*" }
}
if (-not $grants) { Write-Host "Keine Grants nach Filter gefunden."; return }

$records = foreach ($g in $grants) {
  $client   = $spById[$g.clientId]
  $resource = $spById[$g.resourceId]
  $scopes = ($g.scope -split '\s+') | Where-Object { $_ }
  foreach ($s in $scopes) {
    [pscustomobject]@{
      GrantId         = $g.id
      ConsentType     = $g.consentType
      ClientId        = $g.clientId
      ClientName      = $client.Name
      ClientAppId     = $client.AppId
      ResourceId      = $g.resourceId
      ResourceName    = $resource.Name
      ResourceAppId   = $resource.AppId
      Scope           = $s
      PrincipalUserId = if ($g.consentType -eq "Principal") { $g.principalId }
      PrincipalUPN    = if ($g.consentType -eq "Principal" -and $userById.ContainsKey($g.principalId)) { $userById[$g.principalId].UPN }
      PrincipalName   = if ($g.consentType -eq "Principal" -and $userById.ContainsKey($g.principalId)) { $userById[$g.principalId].Name }
    }
  }
}

$criticalPatterns = @(
  '.*ReadWrite.*',
  '^Directory\.AccessAsUser\.All$',
  '^Mail\.ReadWrite(\.All)?$','^Calendars\.ReadWrite(\.All)?$','^Files\.ReadWrite(\.All)?$','^Contacts\.ReadWrite(\.All)?$',
  '^MailboxSettings\.ReadWrite$','^SecurityEvents\.ReadWrite\.All$','^PrivilegedAccess.*$','^Sites\.FullControl\.All$'
)
$lowImpactExact = @('User.Read','openid','profile','email','offline_access')
function Test-CriticalScope { param([string]$s) foreach ($p in $criticalPatterns) { if ($s -match $p) { return $true } } return $false }
function Is-LowImpactScopeSet { param([string[]]$scopes)
  if (-not $scopes -or $scopes.Count -eq 0) { return $false }
  foreach ($s in $scopes) { if ($lowImpactExact -notcontains $s) { return $false } }
  return $true
}

$apps = @()
$byClient = $records | Group-Object ClientId
foreach ($grp in $byClient) {
  $items       = $grp.Group
  $clientId    = $grp.Name
  $clientName  = ($items | Select-Object -First 1).ClientName
  $clientAppId = ($items | Select-Object -First 1).ClientAppId

  $hasAdmin = $items | Where-Object { $_.ConsentType -eq 'AllPrincipals' }
  $hasUser  = $items | Where-Object { $_.ConsentType -eq 'Principal'  }

  $allScopes = @(($items | Select-Object -Expand Scope) | Sort-Object -Unique)
  $isCritical = $false; foreach ($sc in $allScopes) { if (Test-CriticalScope $sc) { $isCritical = $true; break } }
  $status = if ($isCritical) { 'fail' } elseif (Is-LowImpactScopeSet $allScopes) { 'pass' } else { 'warn' }

  $adminDetails = @()
  foreach ($g in ($hasAdmin | Group-Object ResourceName, Scope)) {
    $x = $g.Group[0]
    $adminDetails += [pscustomobject]@{ resource = $x.ResourceName; scope = $x.Scope; type = 'Admin' }
  }
  $userDetails = @()
  foreach ($g in ($hasUser | Group-Object ResourceName, Scope)) {
    $x = $g.Group[0]
    $users = $g.Group | Where-Object { $_.PrincipalUPN } | Select-Object -Expand PrincipalUPN -Unique | Sort-Object
    $userDetails += [pscustomobject]@{ resource=$x.ResourceName; scope=$x.Scope; type='User'; users=@($users); userCount=$users.Count }
  }

  $userPermCount = ($userDetails | Select-Object -ExpandProperty scope -Unique).Count
  $userCount     = ($userDetails | ForEach-Object { $_.users } | Where-Object { $_ } | Select-Object -Unique).Count

  $apps += [pscustomobject]@{
    clientId       = $clientId
    clientName     = $clientName
    clientAppId    = $clientAppId
    status         = $status
    tags           = @(
      if ($hasAdmin) { "Admin Consent" }
      if ($hasUser)  { "User Consent"  }
      "user-perm: $userPermCount"
      "user-count: $userCount"
    )
    adminGrants    = $adminDetails
    userGrants     = $userDetails
    allScopes      = $allScopes
  }
}

$total = $apps.Count
$fail  = ($apps | Where-Object status -eq 'fail').Count
$pass  = ($apps | Where-Object status -eq 'pass').Count
$warn  = ($apps | Where-Object status -eq 'warn').Count

$outDir = (Get-Location).Path
$ts = Get-Date -Format "yyyyMMdd_HHmm"
$csvPath = Join-Path $outDir ("OAuth2_Consent_Apps_$ts.csv")
$apps |
  Select-Object clientName, clientAppId, status,
                @{n='tags';e={$_.tags -join ';'}},
                @{n='adminScopes';e={($_.adminGrants | ForEach-Object { "$($_.resource):$($_.scope)" }) -join ';'}},
                @{n='userScopes'; e={($_.userGrants  | ForEach-Object { "$($_.resource):$($_.scope)" }) -join ';'}},
                @{n='userCount';  e={($_.userGrants  | ForEach-Object { $_.users } | Select-Object -Unique).Count}} |
  Export-Csv -NoTypeInformation -Encoding UTF8 -Path $csvPath

$data = [pscustomobject]@{
  generatedAt = (Get-Date).ToString("s")
  totals      = @{ total=$total; pass=$pass; warn=$warn; fail=$fail }
  items       = $apps
}
$json = $data | ConvertTo-Json -Depth 12

$html = @'
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>OAuth2 Consent Report</title>
<style>

  :root{
    --brand:#7ad12b;         /* Green (approx) */
    --brand-2:#00c2a8;       /* Teal Accent */
    --ok:#7ad12b;            /* pass */
    --warn:#ffd166;          /* warm yellow for warn */
    --fail:#ff5a5f;          /* coral red for critical */
    --muted:#9aa3b2;
    --bg:#0a0f14;            /* very dark */
    --card:#0f151c;
    --text:#e9f0f3;
    --chip:#111b22;
    --line:#1e2a33;
  }
  html,body{margin:0;height:100%;background:var(--bg);color:var(--text);font:14px/1.45 system-ui,Segoe UI,Roboto,Arial,sans-serif}
  a{color:var(--brand)}
  header{position:sticky;top:0;background:linear-gradient(180deg,#0a0f14 60%,#0a0f14cc 100%);backdrop-filter:blur(6px);z-index:10;border-bottom:1px solid var(--line)}
  .wrap{max-width:1200px;margin:0 auto;padding:14px 16px}
  .row{display:flex;gap:8px;flex-wrap:wrap;align-items:center}
  input[type="search"],button{background:#0c1319;border:1px solid #1a2a33;color:var(--text);padding:8px 10px;border-radius:8px}
  input[type="search"]::placeholder{color:#7b8794}
  button{cursor:pointer}
  .chip{display:inline-flex;gap:8px;align-items:center;background:var(--chip);padding:6px 10px;border-radius:20px;border:1px solid #1a2a33}
  .chip b{font-weight:600}
  .chip.clickable{cursor:pointer}
  .chip.active{outline:2px solid var(--brand-2); box-shadow:0 0 0 2px #081118 inset}
  .status{width:10px;height:10px;border-radius:999px}
  .status.pass{background:var(--ok)} .status.warn{background:var(--warn)} .status.fail{background:var(--fail)}
  /* Liste mittig zentrieren */
  #list{max-width:1100px;margin:14px auto;display:grid;grid-template-columns:1fr;gap:10px}
  .card{background:var(--card);border:1px solid var(--line);border-radius:12px;overflow:hidden}
  .item{display:grid;grid-template-columns:28px 1fr auto;gap:10px;align-items:center;padding:12px;border-bottom:1px solid var(--line)}
  h4{margin:0;font-size:15px}
  .meta{color:#7b8794;font-size:12px}
  .tags{display:flex;gap:6px;flex-wrap:wrap;margin-top:6px}
  .tag{background:#0c1319;border:1px solid #1a2a33;padding:2px 8px;border-radius:999px;font-size:12px}
  .tag-admin{background:rgba(122,209,43,0.12);border-color:#2f4; color:#bdf8a7}
  .tag-user{background:rgba(0,194,168,0.12);border-color:#0bd; color:#a6fff3}
  .actions{display:flex;gap:6px}
  details{background:#0b1218}
  details[open] summary{border-bottom:1px dashed var(--line)}
  summary{list-style:none;padding:10px 12px;cursor:pointer;color:#cbd5e1}
  summary::-webkit-details-marker{display:none}
  .sec-title{font-size:12px;text-transform:uppercase;letter-spacing:.08em;color:#9aa3b2;margin:8px 0}
  .divider{height:1px;background:linear-gradient(90deg,transparent, #24424f 30%, #24424f 70%, transparent);margin:10px 0}
  .consent-row{padding:6px 0}
  .pill{border-radius:999px;padding:2px 8px;border:1px solid #1a2a33;background:#0c1319;font-size:12px}
  .footer{padding:20px;color:#9aa3b2;text-align:center;border-top:1px solid var(--line)}
  .footer .sig{display:inline-flex;gap:8px;align-items:center}
  .icon{width:16px;height:16px;vertical-align:-3px;fill:#0e76a8}
</style>
</head>
<body>
<header>
  <div class="wrap row">
    <strong style="font-size:16px">Entra ID – OAuth2 Consent Report</strong>

    <!-- Klickbare Filter-Chips -->
    <span id="chipTotal" class="chip clickable"><b id="tTotal">0</b> Enterprise Apps</span>
    <span id="chipFail"  class="chip clickable"><span class="status fail"></span><b id="tFail">0</b> ReadWrite / critical</span>
    <span id="chipPass"  class="chip clickable"><span class="status pass"></span><b id="tPass">0</b> Low Impact only</span>
    <span id="chipWarn"  class="chip clickable"><span class="status warn"></span><b id="tWarn">0</b> Other (no ReadWrite)</span>

    <div style="flex:1"></div>
    <input id="q" type="search" placeholder="Search enterprise app" />
    <button id="exportCsv" title="Export filtered apps to CSV">Export CSV</button>
  </div>
</header>

<div class="wrap" id="list"></div>
<div class="footer">
  <span class="sig">
    created by <strong>Jürgen Waldl</strong>
    <a href="https://www.linkedin.com/in/jürgen-waldl-6592837b/" target="_blank" rel="noopener" title="LinkedIn – Jürgen Waldl">
      <svg class="icon" viewBox="0 0 24 24" aria-hidden="true">
        <path d="M4.98 3.5C4.98 4.88 3.86 6 2.5 6S0 4.88 0 3.5 1.12 1 2.5 1 4.98 2.12 4.98 3.5zM0 8h5v16H0zM8 8h4.8v2.2h.07c.67-1.2 2.3-2.46 4.73-2.46 5.06 0 6 3.33 6 7.66V24h-5V16.4c0-1.8-.03-4.1-2.5-4.1-2.5 0-2.88 1.95-2.88 3.97V24H8z"/>
      </svg>
    </a>
  </span>
  <div>generated @ <span id="genAt"></span></div>
</div>

<script>
const report = __JSON__;

// --- State & Elements ---
const state = { q:"", filter:"all" };
const $q = document.getElementById("q");
const $list = document.getElementById("list");

// Chips
const $chipTotal = document.getElementById("chipTotal");
const $chipFail  = document.getElementById("chipFail");
const $chipPass  = document.getElementById("chipPass");
const $chipWarn  = document.getElementById("chipWarn");

// Headerzahlen
document.getElementById("tTotal").textContent = report.totals.total;
document.getElementById("tPass").textContent  = report.totals.pass;
document.getElementById("tWarn").textContent  = report.totals.warn;
document.getElementById("tFail").textContent  = report.totals.fail;
document.getElementById("genAt").textContent  = report.generatedAt;

// Suche
$q.addEventListener("input", ()=>{ state.q = $q.value.trim().toLowerCase(); render(); });
document.addEventListener("keydown", e=>{ if(e.key==="/" && (e.ctrlKey||e.metaKey)){ e.preventDefault(); $q.focus(); }});

// Filter-Chip-UI
function setFilter(f){ state.filter = f; updateActiveChip(); render(); }
function updateActiveChip(){
  [$chipTotal,$chipFail,$chipPass,$chipWarn].forEach(el=>el.classList.remove("active"));
  if(state.filter==="all")  $chipTotal.classList.add("active");
  if(state.filter==="fail") $chipFail.classList.add("active");
  if(state.filter==="pass") $chipPass.classList.add("active");
  if(state.filter==="warn") $chipWarn.classList.add("active");
}
$chipTotal.addEventListener("click", ()=> setFilter("all"));
$chipFail .addEventListener("click", ()=> setFilter("fail"));
$chipPass .addEventListener("click", ()=> setFilter("pass"));
$chipWarn .addEventListener("click", ()=> setFilter("warn"));
updateActiveChip();

// Utils
function escapeHtml(s){ return String(s).replace(/[&<>"']/g,m=>({ "&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#039;"}[m])) }
function asArray(x){ return Array.isArray(x) ? x : (x ? [x] : []); }
function tagHtml(t){
  const label = escapeHtml(t);
  if (/^admin consent$/i.test(t)) return `<span class="tag tag-admin">${label}</span>`;
  if (/^user consent$/i.test(t))  return `<span class="tag tag-user">${label}</span>`;
  return `<span class="tag">${label}</span>`;
}
function filterItems(items){
  let out = items;
  if(state.filter !== "all"){ out = out.filter(x => (x.status||"") === state.filter); }
  if(state.q){
    const q = state.q;
    out = out.filter(x => (x.clientName||"").toLowerCase().includes(q));
  }
  return out;
}

// Render
function render(){
  const items = filterItems(report.items);
  $list.innerHTML = items.map(item => {
    const tags = (item.tags||[]).map(tagHtml).join("");
    const adminRows = (item.adminGrants||[]).map(a=>`<div class="consent-row"><span class="pill">Admin</span> ${escapeHtml(a.resource)} • <b>${escapeHtml(a.scope)}</b></div>`).join("");
    const userRows  = (item.userGrants ||[]).map(u=>{
      const arr = asArray(u.users);
      return `<div class="consent-row"><span class="pill">User</span> ${escapeHtml(u.resource)} • <b>${escapeHtml(u.scope)}</b> • users: ${escapeHtml(arr.join(", "))}</div>`;
    }).join("");
    const hasAdmin = !!adminRows, hasUser = !!userRows;
    const sections = [
      hasAdmin ? `<div class="sec-title">Admin Consents</div>${adminRows}` : "",
      (hasAdmin && hasUser) ? `<div class="divider"></div>` : "",
      hasUser  ? `<div class="sec-title">User Consents</div>${userRows}`   : ""
    ].join("");
    const none = (!hasAdmin && !hasUser) ? "<div class='consent-row' style='color:#7b8794'>No delegated grants found.</div>" : "";
    const scopeArr = asArray(item.allScopes);
    return `
      <div class="card">
        <div class="item">
          <span class="status ${item.status}"></span>
          <div>
            <h4>${escapeHtml(item.clientName||"(no name)")}</h4>
            <div class="meta">${escapeHtml(item.clientAppId||"")} • ${scopeArr.length} scope(s)</div>
            <div class="tags">${tags}</div>
          </div>
          <div class="actions">
            <button onclick='copyJson(${JSON.stringify(item).replace(/</g,"\\u003c")})' title="Copy JSON">Copy</button>
          </div>
        </div>
        <details>
          <summary>Details</summary>
          <div style="padding:12px 14px">
            ${sections}${none}
          </div>
        </details>
      </div>`;
  }).join("");
}

function copyJson(obj){ navigator.clipboard.writeText(JSON.stringify(obj,null,2)); }

// Export CSV (aktuell gefiltert)
document.getElementById("exportCsv").addEventListener("click", () => {
  const items = filterItems(report.items);
  const headers = ["clientName","clientAppId","status","tags","adminScopes","userScopes","userCount"];
  const rows = items.map(x=>{
    const adminScopes = (x.adminGrants||[]).map(a=>`${a.resource}:${a.scope}`).join(";")
    const userScopes  = (x.userGrants ||[]).map(u=>`${u.resource}:${u.scope}`).join(";")
    const userCount   = (x.userGrants ||[]).flatMap(u => asArray(u.users)).filter((v,i,arr)=>arr.indexOf(v)===i).length;
    const tags = (x.tags||[]).join(";");
    return {clientName:x.clientName||"", clientAppId:x.clientAppId||"", status:x.status||"", tags, adminScopes, userScopes, userCount};
  });
  const csv = [headers.join(",")].concat(rows.map(r=>headers.map(h=>`"${String(r[h]).replace(/"/g,'""')}"`).join(","))).join("\n");
  const blob = new Blob([csv], {type:"text/csv;charset=utf-8;"});
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href = url; a.download = "oauth2-consents.csv"; a.click();
  URL.revokeObjectURL(url);
});

render();
</script>
</body>
</html>
'@

$escapedJson = $json.Replace('</script>','<\/script>')
$html = $html -replace '__JSON__', $escapedJson

$ts = Get-Date -Format "yyyyMMdd_HHmm"
$htmlPath = Join-Path $outDir ("OAuth2_Consent_Report_$ts.html")
Set-Content -Path $htmlPath -Encoding UTF8 -Value $html

Write-Host "`nFertig:" -ForegroundColor Cyan
Write-Host "  CSV : $csvPath"
Write-Host "  HTML: $htmlPath"
