

function Find-FirstFieldValue($obj, $fieldNames) {
    if ($null -eq $obj) { return $null }

    if ($obj -is [pscustomobject]) {
        foreach ($name in $fieldNames) {
            $prop = $obj.PSObject.Properties[$name]
            if ($prop -and $prop.Value -and [string]$prop.Value -ne "SEM RESULTADO") {
                return [string]$prop.Value
            }
        }

        foreach ($p in $obj.PSObject.Properties) {
            $found = Find-FirstFieldValue $p.Value $fieldNames
            if ($found) { return $found }
        }
    }
    elseif ($obj -is [System.Collections.IEnumerable] -and -not ($obj -is [string])) {
        foreach ($item in $obj) {
            $found = Find-FirstFieldValue $item $fieldNames
            if ($found) { return $found }
        }
    }

    return $null
}


function Convert-BirthDateToAge($birthValue) {
    if($null -eq $birthValue){ return $null }
    $s = ([string]$birthValue).Trim()
    if([string]::IsNullOrWhiteSpace($s) -or $s -eq "SEM RESULTADO"){ return $null }

    $formats = @(
        "dd/MM/yyyy",
        "d/M/yyyy",
        "yyyy-MM-dd",
        "yyyy-M-d",
        "yyyy-MM-dd HH:mm:ss",
        "yyyy-M-d H:mm:ss",
        "dd/MM/yyyy HH:mm:ss",
        "d/M/yyyy H:mm:ss"
    )

    $dataNasc = $null
    foreach($fmt in $formats){
        try {
            $dataNasc = [datetime]::ParseExact($s, $fmt, [System.Globalization.CultureInfo]::InvariantCulture)
            break
        } catch {}
    }

    if($null -eq $dataNasc){
        try {
            $dataNasc = [datetime]::Parse($s, [System.Globalization.CultureInfo]::GetCultureInfo("pt-BR"))
        } catch {}
    }

    if($null -eq $dataNasc){ return $null }

    $hoje = Get-Date
    $idade = $hoje.Year - $dataNasc.Year
    if(($hoje.Month -lt $dataNasc.Month) -or (($hoje.Month -eq $dataNasc.Month) -and ($hoje.Day -lt $dataNasc.Day))){
        $idade--
    }
    return $idade
}

function Set-OrderAge($order, $idade) {
    if($null -eq $order -or $null -eq $idade){ return }
    $idadeStr = [string]$idade
    if($order.PSObject.Properties.Name -contains "idade"){
        $order.idade = $idadeStr
    } else {
        $order | Add-Member -NotePropertyName idade -NotePropertyValue $idadeStr -Force
    }
}

function Is-OrderAgeMissing($order) {
    if($null -eq $order){ return $true }
    if($order.PSObject.Properties.Name -contains "idade"){
        $v = ([string]$order.idade).Trim()
        if([string]::IsNullOrWhiteSpace($v)){ return $true }
        if($v -match '^(sem idade|sem\s+resultado|null|n/a|-)$'){ return $true }
        $n = 0
        if([int]::TryParse($v, [ref]$n)){ return $false }
        return $true
    }
    return $true
}

function BuscarIdadeCPF($cpf) {
    $script:FDX_DEBUG = [pscustomobject]@{
        cpfOriginal = $cpf
        cpfLimpo = ""
        url = ""
        tokenConfigurado = $false
        response = $null
        idadeEncontrada = $null
        erro = ""
        nascimentoBruto = ""
    }

    if(-not $cpf){ return $null }

    $cpfLimpo = ([string]$cpf) -replace '[^\d]', ''
    $script:FDX_DEBUG.cpfLimpo = $cpfLimpo

    if([string]::IsNullOrWhiteSpace($cpfLimpo)){ return $null }
    if($FDX_TOKEN -eq "" -or $FDX_TOKEN -eq "COLOQUE_SEU_TOKEN_AQUI"){
        $script:FDX_DEBUG.erro = "Token não configurado"
        return $null
    }

    $script:FDX_DEBUG.tokenConfigurado = $true

    try {
        $url = "https://api.fdxapis.us/api.php?token=$FDX_TOKEN&cpf_simples=$cpfLimpo"
        $script:FDX_DEBUG.url = $url

        $resp = Invoke-RestMethod -Method Get -Uri $url -Headers @{
            "Accept" = "application/json"
            "User-Agent" = "Mozilla/5.0"
        }

        $script:FDX_DEBUG.response = $resp

        $idadeDireta = $null
        try {
            if($resp.response -and $resp.response.DADOS){
                foreach($campoIdade in @('IDADE','idade','ANOS','anos')){
                    if($resp.response.DADOS.PSObject.Properties.Name -contains $campoIdade -and $resp.response.DADOS.$campoIdade){
                        $tmpIdade = 0
                        if([int]::TryParse(([string]$resp.response.DADOS.$campoIdade), [ref]$tmpIdade)){
                            $idadeDireta = $tmpIdade
                            break
                        }
                    }
                }
            }
        } catch {}
        if($null -ne $idadeDireta){
            $script:FDX_DEBUG.idadeEncontrada = $idadeDireta
            return $idadeDireta
        }

        $nasc = $null
        try {
            if($resp.response -and $resp.response.DADOS -and $resp.response.DADOS.NASC){
                $nasc = [string]$resp.response.DADOS.NASC
            }
        } catch {}

        if([string]::IsNullOrWhiteSpace([string]$nasc)){
            $nasc = Find-FirstFieldValue $resp @("NASC", "NASCIMENTO", "DATA_NASCIMENTO", "DT_NASCIMENTO", "dataDeNascimento", "data_de_nascimento", "nascimento")
        }

        $script:FDX_DEBUG.nascimentoBruto = [string]$nasc

        $idade = Convert-BirthDateToAge $nasc
        if($null -ne $idade){
            $script:FDX_DEBUG.idadeEncontrada = $idade
            return $idade
        }

        $script:FDX_DEBUG.erro = "Nascimento não encontrado ou formato inválido"
        return $null
    } catch {
        $script:FDX_DEBUG.erro = $_.Exception.Message
        return $null
    }
}

$FDX_TOKEN = "33d3d68c55fb8fdf5c63fa03a55d4d1e"

Write-Host "Trevizio Bling Multi iniciando..." -ForegroundColor Cyan

$ErrorActionPreference = "Stop"
$baseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$port = if($env:PORT){ [int]$env:PORT } else { 8765 }
$hostName = if($env:RENDER -or $env:RENDER_EXTERNAL_URL){ "0.0.0.0" } else { "localhost" }
$publicBaseUrl = if($env:PUBLIC_BASE_URL){ $env:PUBLIC_BASE_URL.TrimEnd("/") } elseif($env:RENDER_EXTERNAL_URL){ $env:RENDER_EXTERNAL_URL.TrimEnd("/") } else { "http://localhost:$port" }

if($env:DATA_DIR){
    $dataDir = $env:DATA_DIR
}
elseif($env:APPDATA){
    $dataDir = Join-Path $env:APPDATA "TrevizioBlingMulti"
}
else {
    $dataDir = Join-Path $baseDir "data"
}

if (!(Test-Path $dataDir)) { New-Item -ItemType Directory -Path $dataDir -Force | Out-Null }

Write-Host ("Servidor rodando em {0}" -f $publicBaseUrl) -ForegroundColor Green
$accountsPath = Join-Path $dataDir "accounts.json"
$ordersPath   = Join-Path $dataDir "orders.json"
$agesPath     = Join-Path $dataDir "ages.json"
$pendingPath  = Join-Path $dataDir "pending.json"
$usersPath    = Join-Path $dataDir "users.json"
$sessionsPath = Join-Path $dataDir "sessions.json"

$SUPABASE_URL = $env:SUPABASE_URL
$SUPABASE_KEY = if($env:SUPABASE_SECRET_KEY){ $env:SUPABASE_SECRET_KEY } elseif($env:SUPABASE_SERVICE_ROLE_KEY){ $env:SUPABASE_SERVICE_ROLE_KEY } else { "" }

function Get-SupabaseHeaders {
    if([string]::IsNullOrWhiteSpace($SUPABASE_URL) -or [string]::IsNullOrWhiteSpace($SUPABASE_KEY)){
        throw "Supabase não configurado. Defina SUPABASE_URL e SUPABASE_SECRET_KEY no Render."
    }
    return @{
        "apikey" = $SUPABASE_KEY
        "Authorization" = $SUPABASE_KEY
        "Content-Type" = "application/json"
        "Accept" = "application/json"
        "Prefer" = "return=representation"
    }
}
function Invoke-SupabaseGet($pathAndQuery) {
    $url = "$SUPABASE_URL/rest/v1/$pathAndQuery"
    return Invoke-RestMethod -Method Get -Uri $url -Headers (Get-SupabaseHeaders)
}
function Invoke-SupabasePost($pathAndQuery, $bodyObj) {
    $url = "$SUPABASE_URL/rest/v1/$pathAndQuery"
    $body = ($bodyObj | ConvertTo-Json -Depth 30)
    return Invoke-RestMethod -Method Post -Uri $url -Headers (Get-SupabaseHeaders) -Body $body
}
function Invoke-SupabasePatch($pathAndQuery, $bodyObj) {
    $url = "$SUPABASE_URL/rest/v1/$pathAndQuery"
    $body = ($bodyObj | ConvertTo-Json -Depth 30)
    return Invoke-RestMethod -Method Patch -Uri $url -Headers (Get-SupabaseHeaders) -Body $body
}
function Invoke-SupabaseDelete($pathAndQuery) {
    $url = "$SUPABASE_URL/rest/v1/$pathAndQuery"
    return Invoke-RestMethod -Method Delete -Uri $url -Headers (Get-SupabaseHeaders)
}



function Read-JsonFile($path, $defaultObj) {
    if (Test-Path $path) { try { return (((Get-Content $path) -join "`n") | ConvertFrom-Json) } catch { return $defaultObj } }
    return $defaultObj
}
function Write-JsonFile($path, $obj) { $obj | ConvertTo-Json -Depth 60 | Set-Content -Encoding UTF8 $path }

function Convert-DbContaToApp($a) {
    return [pscustomobject]@{
        alias = $a.alias
        clientId = $a.client_id
        clientSecret = $a.client_secret
        redirectUri = "$publicBaseUrl/callback"
        access_token = $a.access_token
        refresh_token = $a.refresh_token
        expires_at = $(if($a.expires_at){$a.expires_at}else{$null})
    }
}
function Get-Accounts {
    $rows = Invoke-SupabaseGet "contas_bling?select=*"
    $out = @()
    foreach($a in @($rows)){ $out += (Convert-DbContaToApp $a) }
    return @($out)
}
function Save-Accounts($accounts) {
    foreach($a in @($accounts)){
        $payload = [pscustomobject]@{
            alias = $a.alias
            client_id = $a.clientId
            client_secret = $a.clientSecret
            access_token = $a.access_token
            refresh_token = $a.refresh_token
            expires_at = $(if($a.expires_at){$a.expires_at}else{$null})
            connected = [bool](-not [string]::IsNullOrWhiteSpace([string]$a.access_token))
            updated_at = (Get-Date).ToString("o")
        }
        $existing = @(Invoke-SupabaseGet ("contas_bling?select=id&alias=eq.{0}" -f [uri]::EscapeDataString([string]$a.alias)))
        if($existing.Count -gt 0){
            [void](Invoke-SupabasePatch ("contas_bling?alias=eq.{0}" -f [uri]::EscapeDataString([string]$a.alias)) $payload)
        } else {
            [void](Invoke-SupabasePost "contas_bling" $payload)
        }
    }
}
function Get-Orders {
    try {
        $rows = @(Invoke-SupabaseGet "notas_sincronizadas?select=conteudo&id=eq.1")
        if($rows.Count -gt 0 -and $null -ne $rows[0].conteudo -and -not [string]::IsNullOrWhiteSpace([string]$rows[0].conteudo)){
            $parsed = ([string]$rows[0].conteudo | ConvertFrom-Json)
            if($null -eq $parsed){ return @() }
            return @(Apply-AgesToOrders @($parsed))
        }
    } catch {
    }
    return @(Apply-AgesToOrders @(Read-JsonFile $ordersPath @()))
}
function Save-Orders($orders) {
    $safeOrders = @($orders)
    Write-JsonFile $ordersPath $safeOrders
    try {
        $payload = [pscustomobject]@{
            id = 1
            conteudo = ($safeOrders | ConvertTo-Json -Depth 60 -Compress)
            updated_at = (Get-Date).ToString("o")
        }
        $existing = @(Invoke-SupabaseGet "notas_sincronizadas?select=id&id=eq.1")
        if($existing.Count -gt 0){
            [void](Invoke-SupabasePatch "notas_sincronizadas?id=eq.1" $payload)
        } else {
            [void](Invoke-SupabasePost "notas_sincronizadas" $payload)
        }
    } catch {
    }
}
function Get-Ages { Read-JsonFile $agesPath ([pscustomobject]@{}) }

function Get-DefaultUsers {
    return @(
        [pscustomobject]@{ usuario="admin"; senha="mister4419"; tipo="admin"; nome="Administrador"; bloqueado=$false },
        [pscustomobject]@{ usuario="colaborador"; senha="colaborador"; tipo="colaborador"; nome="Colaborador"; bloqueado=$false }
    )
}
function Ensure-UsersFile {
    return Get-Users
}
function Get-Users {
    $rows = Invoke-SupabaseGet "usuarios?select=*"
    $out = @()
    foreach($u in @($rows)){
        $out += [pscustomobject]@{
            usuario = $u.usuario
            senha = $u.senha
            tipo = $u.tipo
            nome = $u.nome
            bloqueado = [bool]$u.bloqueado
        }
    }
    if($out.Count -eq 0){
        foreach($u in (Get-DefaultUsers)){
            [void](Invoke-SupabasePost "usuarios" ([pscustomobject]@{
                usuario = $u.usuario
                senha = $u.senha
                nome = $u.nome
                tipo = $u.tipo
                bloqueado = [bool]$u.bloqueado
            }))
        }
        return @(Get-Users)
    }
    return @($out)
}
function Save-Users($users) {
    # não usado mais com Supabase
}
function Get-Sessions {
    $s = Read-JsonFile $sessionsPath @()
    if ($s -is [string]) { return @() }
    if ($null -eq $s) { return @() }
    return @($s)
}
function Save-Sessions($sessions) {
    Write-JsonFile $sessionsPath @($sessions)
}
function Find-User($usuario) {
    $usuarioNorm = ([string]$usuario).Trim().ToLower()
    $rows = @(Invoke-SupabaseGet ("usuarios?select=*&usuario=eq.{0}" -f [uri]::EscapeDataString($usuarioNorm)))
    if($rows.Count -gt 0){
        $u = $rows[0]
        return [pscustomobject]@{
            usuario = $u.usuario
            senha = $u.senha
            tipo = $u.tipo
            nome = $u.nome
            bloqueado = [bool]$u.bloqueado
        }
    }
    return $null
}
function To-PublicUser($u) {
    return [pscustomobject]@{ usuario=$u.usuario; tipo=$u.tipo; nome=$u.nome; bloqueado=[bool]$u.bloqueado }
}
function New-LoginSession($u) {
    $token = [guid]::NewGuid().ToString('N')
    $sessions = @()
    foreach($s in (Get-Sessions)){
        if([string]$s.usuario -ne [string]$u.usuario){ $sessions += $s }
    }
    $sessions += [pscustomobject]@{
        token = $token
        usuario = $u.usuario
        tipo = $u.tipo
        nome = $u.nome
        createdAt = (Get-Date).ToString('o')
    }
    Save-Sessions $sessions
    return [pscustomobject]@{ token=$token; usuario=$u.usuario; tipo=$u.tipo; nome=$u.nome }
}
function Remove-SessionToken($token) {
    $new = @()
    foreach($s in (Get-Sessions)){
        if([string]$s.token -ne [string]$token){ $new += $s }
    }
    Save-Sessions $new
}
function Get-SessionFromRequest($req) {
    $token = $req.Headers['X-Auth-Token']
    if(-not $token){ return $null }
    foreach($s in (Get-Sessions)){
        if([string]$s.token -eq [string]$token){ return $s }
    }
    return $null
}
function Require-Admin($req) {
    $session = Get-SessionFromRequest $req
    if(-not $session){ throw 'Sessão inválida. Entre novamente.' }
    if([string]$session.tipo -ne 'admin'){ throw 'Somente admin pode fazer isso.' }
    return $session
}

function Require-Auth($req) {
    $session = Get-SessionFromRequest $req
    if(-not $session){ throw 'Sessão inválida. Entre novamente.' }
    return $session
}
function Add-AppUser($usuario, $senha, $tipo, $nome) {
    $usuario = ([string]$usuario).Trim().ToLower()
    $nome = ([string]$nome).Trim()
    if(-not $usuario){ throw 'Usuário é obrigatório.' }
    if(-not $senha){ throw 'Senha é obrigatória.' }
    if($tipo -ne 'admin' -and $tipo -ne 'colaborador'){ throw 'Tipo inválido.' }
    if(Find-User $usuario){ throw 'Já existe um usuário com esse nome.' }
    [void](Invoke-SupabasePost "usuarios" ([pscustomobject]@{
        usuario = $usuario
        senha = $senha
        tipo = $tipo
        nome = $(if($nome){$nome}else{$usuario})
        bloqueado = $false
    }))
}
function Update-AppUser($usuario, $senha, $tipo, $nome, $bloqueado) {
    $usuario = ([string]$usuario).Trim().ToLower()
    $existing = Find-User $usuario
    if(-not $existing){ throw 'Usuário não encontrado.' }

    $novoTipo = if($tipo){ $tipo } else { $existing.tipo }
    $novoNome = if($nome){ $nome } else { $existing.nome }
    $novoBloq = if($null -ne $bloqueado){ [bool]$bloqueado } else { [bool]$existing.bloqueado }

    if($usuario -eq 'admin' -and $novoTipo -ne 'admin'){ throw 'O admin principal deve continuar admin.' }
    if($usuario -eq 'admin' -and $novoBloq){ throw 'O admin principal não pode ser bloqueado.' }

    $payload = [ordered]@{
        tipo = $novoTipo
        nome = $novoNome
        bloqueado = $novoBloq
    }
    if($senha){ $payload.senha = $senha }

    [void](Invoke-SupabasePatch ("usuarios?usuario=eq.{0}" -f [uri]::EscapeDataString([string]$usuario)) ([pscustomobject]$payload))
}
function Delete-AppUser($usuario) {
    $usuario = ([string]$usuario).Trim().ToLower()
    if([string]$usuario -eq 'admin'){ throw 'O admin principal não pode ser excluído.' }
    $existing = Find-User $usuario
    if(-not $existing){ throw 'Usuário não encontrado.' }
    [void](Invoke-SupabaseDelete ("usuarios?usuario=eq.{0}" -f [uri]::EscapeDataString([string]$usuario)))
    $sessions = @((Get-Sessions) | Where-Object { [string]$_.usuario -ne [string]$usuario })
    Save-Sessions $sessions
}

function Get-BasicAuthHeader($clientId, $clientSecret) {
    $pair = "{0}:{1}" -f $clientId, $clientSecret
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($pair)
    "Basic " + [Convert]::ToBase64String($bytes)
}
function Send-Json($ctx, $obj, $status=200) {
    $json = $obj | ConvertTo-Json -Depth 60
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $ctx.Response.StatusCode = $status
    $ctx.Response.ContentType = "application/json; charset=utf-8"
    $ctx.Response.ContentLength64 = $bytes.Length
    $ctx.Response.OutputStream.Write($bytes,0,$bytes.Length)
    $ctx.Response.OutputStream.Close()
}
function Send-Text($ctx, $text, $ctype="text/plain; charset=utf-8", $status=200) {
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($text)
    $ctx.Response.StatusCode = $status
    $ctx.Response.ContentType = $ctype
    $ctx.Response.ContentLength64 = $bytes.Length
    $ctx.Response.OutputStream.Write($bytes,0,$bytes.Length)
    $ctx.Response.OutputStream.Close()
}

function Write-StreamJson($ctx, $obj) {
    $json = ($obj | ConvertTo-Json -Compress -Depth 20) + "`n"
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($json)
    $ctx.Response.OutputStream.Write($bytes,0,$bytes.Length)
    $ctx.Response.OutputStream.Flush()
}

function Read-BodyJson($req) {
    $sr = New-Object System.IO.StreamReader($req.InputStream, $req.ContentEncoding)
    $body = $sr.ReadToEnd()
    if ([string]::IsNullOrWhiteSpace($body)) { return $null }
    return ($body | ConvertFrom-Json)
}
function Find-Account($alias) {
    foreach($a in (Get-Accounts)){ if([string]$a.alias -eq [string]$alias){ return $a } }
    return $null
}
function Upsert-Account($alias, $clientId, $clientSecret) {
    if (-not $alias) { throw "Alias da loja é obrigatório." }
    $existing = Find-Account $alias
    $payload = [pscustomobject]@{
        alias = $alias
        client_id = $(if($clientId){$clientId}elseif($existing){$existing.clientId}else{""})
        client_secret = $(if($clientSecret){$clientSecret}elseif($existing){$existing.clientSecret}else{""})
        access_token = $(if($existing){$existing.access_token}else{""})
        refresh_token = $(if($existing){$existing.refresh_token}else{""})
        expires_at = $(if($existing -and $existing.expires_at){$existing.expires_at}else{$null})
        connected = $(if($existing){[bool](-not [string]::IsNullOrWhiteSpace([string]$existing.access_token))}else{$false})
        updated_at = (Get-Date).ToString("o")
    }
    if($existing){
        [void](Invoke-SupabasePatch ("contas_bling?alias=eq.{0}" -f [uri]::EscapeDataString([string]$alias)) $payload)
    } else {
        [void](Invoke-SupabasePost "contas_bling" $payload)
    }
}

function Delete-AccountByAlias($alias) {
    $alias = ([string]$alias).Trim()
    if([string]::IsNullOrWhiteSpace($alias)){ throw 'Informe a loja que deseja excluir.' }
    $existing = Find-Account $alias
    if(-not $existing){ throw 'Loja não encontrada.' }

    [void](Invoke-SupabaseDelete ("contas_bling?alias=eq.{0}" -f [uri]::EscapeDataString([string]$alias)))

    # Remove também as notas salvas dessa loja para ela sumir da lista/filtros.
    $orders = Get-Orders
    $updated = @()
    foreach($o in @($orders)){
        if([string]$o.alias -ne [string]$alias){ $updated += $o }
    }
    Save-Orders $updated
    return $updated.Count
}

function Refresh-TokenIfNeeded($alias) {
    $accounts = Get-Accounts
    $updated = @()
    foreach($a in $accounts){
        if($a.alias -eq $alias -and $a.refresh_token){
            $doRefresh = $true
            if($a.expires_at){
                try {
                    $exp = [DateTime]::Parse($a.expires_at)
                    if($exp -gt (Get-Date).AddMinutes(2)){ $doRefresh = $false }
                } catch {}
            }
            if($doRefresh){
                $headers = @{ "Authorization" = Get-BasicAuthHeader $a.clientId $a.clientSecret; "enable-jwt"="1" }
                $body = @{ grant_type="refresh_token"; refresh_token=$a.refresh_token }
                $resp = Invoke-RestMethod -Method Post -Uri "https://www.bling.com.br/Api/v3/oauth/token" -Headers $headers -Body $body -ContentType "application/x-www-form-urlencoded"
                $a.access_token = $resp.access_token
                $a.refresh_token = $resp.refresh_token
                $a.expires_at = (Get-Date).AddSeconds([int]$resp.expires_in).ToString("o")
            }
        }
        $updated += $a
    }
    Save-Accounts $updated
}
function Get-AuthHeader($alias) {
    Refresh-TokenIfNeeded $alias
    $a = Find-Account $alias
    if(-not $a -or -not $a.access_token){ throw "Conta sem token: $alias" }
    return @{ "Authorization" = "Bearer $($a.access_token)"; "enable-jwt"="1"; "Accept"="application/json" }
}
function Find-DeepValuesByFieldName($obj, $fieldName) {
    $results = New-Object System.Collections.ArrayList
    if($null -eq $obj){ return @() }

    if($obj -is [pscustomobject]) {
        $prop = $obj.PSObject.Properties[$fieldName]
        if($prop){
            if(($prop.Value -is [System.Collections.IEnumerable]) -and -not ($prop.Value -is [string])) {
                foreach($item in $prop.Value){ [void]$results.Add($item) }
            } else {
                [void]$results.Add($prop.Value)
            }
        }
        foreach($p in $obj.PSObject.Properties){
            $val = $p.Value
            if($val -is [pscustomobject] -or (($val -is [System.Collections.IEnumerable]) -and -not ($val -is [string]))){
                foreach($x in (Find-DeepValuesByFieldName $val $fieldName)) { [void]$results.Add($x) }
            }
        }
    }
    elseif(($obj -is [System.Collections.IEnumerable]) -and -not ($obj -is [string])) {
        foreach($item in $obj){
            foreach($x in (Find-DeepValuesByFieldName $item $fieldName)) { [void]$results.Add($x) }
        }
    }

    return @($results)
}

function Get-FdxDebugData($cpf) {
    $rawJson = "{}"
    $historicoTelefone = @()
    $ultimoTelefone = @()
    $phones = @()
    $erro = ""

    if(-not $cpf){
        return [pscustomobject]@{
            cpf = $cpf
            phones = @()
            historicoTelefone = @()
            ultimoTelefone = @()
            rawJson = "{}"
            erro = "CPF vazio"
        }
    }

    $cpfLimpo = ([string]$cpf) -replace '[^\d]', ''
    if([string]::IsNullOrWhiteSpace($cpfLimpo)){
        return [pscustomobject]@{
            cpf = $cpf
            phones = @()
            historicoTelefone = @()
            ultimoTelefone = @()
            rawJson = "{}"
            erro = "CPF inválido"
        }
    }

    if($FDX_TOKEN -eq "" -or $FDX_TOKEN -eq "COLOQUE_SEU_TOKEN_AQUI"){
        return [pscustomobject]@{
            cpf = $cpfLimpo
            phones = @()
            historicoTelefone = @()
            ultimoTelefone = @()
            rawJson = "{}"
            erro = "Token FDX não configurado"
        }
    }

    try {
        $url = "https://api.fdxapis.us/api.php?token=$FDX_TOKEN&cpf_simples=$cpfLimpo"
        $resp = Invoke-RestMethod -Method Get -Uri $url -Headers @{
            "Accept" = "application/json"
            "User-Agent" = "Mozilla/5.0"
        }

        $rawJson = ($resp | ConvertTo-Json -Depth 100)

        $historicoTelefone = @(Find-DeepValuesByFieldName $resp "HISTORICO_TELEFONE")
        $ultimoTelefone = @(Find-DeepValuesByFieldName $resp "ULTIMO_TELEFONE")

        foreach($entry in $historicoTelefone){
            if($entry -is [pscustomobject]){
                $ddd = ""
                $tel = ""
                if($entry.PSObject.Properties["DDD"]){ $ddd = [string]$entry.DDD }
                if($entry.PSObject.Properties["TELEFONE"]){ $tel = [string]$entry.TELEFONE }
                $num = (($ddd + $tel) -replace '[^\d]', '')
                if($num){ $phones += $num }
            } elseif($entry -is [string]) {
                $num = ($entry -replace '[^\d]', '')
                if($num){ $phones += $num }
            }
        }

        foreach($entry in $ultimoTelefone){
            if($entry -is [pscustomobject]){
                $ddd = ""
                $tel = ""
                if($entry.PSObject.Properties["DDD"]){ $ddd = [string]$entry.DDD }
                if($entry.PSObject.Properties["TELEFONE"]){ $tel = [string]$entry.TELEFONE }
                $num = (($ddd + $tel) -replace '[^\d]', '')
                if($num){ $phones += $num }
            } elseif($entry -is [string]) {
                $num = ($entry -replace '[^\d]', '')
                if($num){ $phones += $num }
            }
        }

        $phones = @($phones | Select-Object -Unique)

    } catch {
        $erro = $_.Exception.Message
        $rawJson = ('{"erro":"' + ($erro -replace '"','\"') + '"}')
    }

    return [pscustomobject]@{
        cpf = $cpfLimpo
        phones = @($phones)
        historicoTelefone = @($historicoTelefone)
        ultimoTelefone = @($ultimoTelefone)
        rawJson = $rawJson
        erro = $erro
    }
}

function Parse-OrderDate($value) {
    if(-not $value){ return $null }
    $s = [string]$value
    try { return [datetime]::Parse($s) } catch {}
    try { return [datetime]::ParseExact($s, "yyyy-MM-dd", $null) } catch {}
    try { return [datetime]::ParseExact($s, "yyyy-MM-ddTHH:mm:ss", $null) } catch {}
    return $null
}

function Find-PhonesInObject($obj) {
    $out = New-Object System.Collections.ArrayList
    if($null -eq $obj){ return @() }

    if($obj -is [pscustomobject]) {
        foreach($p in $obj.PSObject.Properties) {
            $name = [string]$p.Name
            $val = $p.Value

            if($name -match 'telefone|celular|whatsapp|fone') {
                if($val -is [string]) {
                    $limpo = ($val -replace '[^\d\+]', '')
                    if($limpo){ [void]$out.Add($limpo) }
                } elseif($val -is [pscustomobject] -or ($val -is [System.Collections.IEnumerable] -and -not ($val -is [string]))) {
                    foreach($x in (Find-PhonesInObject $val)) { [void]$out.Add($x) }
                }
            } else {
                foreach($x in (Find-PhonesInObject $val)) { [void]$out.Add($x) }
            }
        }
    }
    elseif(($obj -is [System.Collections.IEnumerable]) -and -not ($obj -is [string])) {
        foreach($item in $obj) {
            foreach($x in (Find-PhonesInObject $item)) { [void]$out.Add($x) }
        }
    }

    return @($out | Select-Object -Unique)
}

function Get-ContactDebugData($alias, $contactId) {
    $rawJson = "{}"
    $phones = @()
    $historicoTelefone = @()
    $ultimoTelefone = @()

    if(-not $contactId -or [string]::IsNullOrWhiteSpace([string]$contactId) -or [string]$contactId -eq "0"){
        return [pscustomobject]@{
            contactId = $contactId
            phones = @()
            historicoTelefone = @()
            ultimoTelefone = @()
            rawJson = $rawJson
        }
    }

    try {
        $headers = Get-AuthHeader $alias
        $url = "https://api.bling.com.br/Api/v3/contatos/$contactId"
        $resp = Invoke-RestMethod -Method Get -Uri $url -Headers $headers
        $raw = if($resp.data){ $resp.data } else { $resp }

        $phones = Find-PhonesInObject $raw

        # procura os blocos específicos no JSON retornado
        $hist = Find-FirstFieldValue $raw @("HISTORICO_TELEFONE")
        $ult = Find-FirstFieldValue $raw @("ULTIMO_TELEFONE")

        if($raw.PSObject.Properties["HISTORICO_TELEFONE"]){
            $historicoTelefone = @($raw.HISTORICO_TELEFONE)
        } elseif($raw.response -and $raw.response.PSObject.Properties["HISTORICO_TELEFONE"]) {
            $historicoTelefone = @($raw.response.HISTORICO_TELEFONE)
        } elseif($raw.resposta -and $raw.resposta.PSObject.Properties["HISTORICO_TELEFONE"]) {
            $historicoTelefone = @($raw.resposta.HISTORICO_TELEFONE)
        }

        if($raw.PSObject.Properties["ULTIMO_TELEFONE"]){
            $ultimoTelefone = @($raw.ULTIMO_TELEFONE)
        } elseif($raw.response -and $raw.response.PSObject.Properties["ULTIMO_TELEFONE"]) {
            $ultimoTelefone = @($raw.response.ULTIMO_TELEFONE)
        } elseif($raw.resposta -and $raw.resposta.PSObject.Properties["ULTIMO_TELEFONE"]) {
            $ultimoTelefone = @($raw.resposta.ULTIMO_TELEFONE)
        }

        $rawJson = ($raw | ConvertTo-Json -Depth 80)
    } catch {
        $rawJson = ('{"erro":"' + ($_.Exception.Message -replace '"','\"') + '"}')
    }

    return [pscustomobject]@{
        contactId = $contactId
        phones = @($phones | Select-Object -Unique)
        historicoTelefone = @($historicoTelefone)
        ultimoTelefone = @($ultimoTelefone)
        rawJson = $rawJson
    }
}

function Get-AgeValue($alias, $cpf, $cliente) {
    $ages = Get-Ages
    $cpfRaw = [string]$cpf
    $cpfLimpo = ($cpfRaw -replace '[^\d]', '')
    $clienteRaw = [string]$cliente
    $keys = @(
        "$alias|$cpfRaw|$clienteRaw",
        "$alias|$cpfLimpo|$clienteRaw",
        "$alias|$cpfRaw|",
        "$alias|$cpfLimpo|",
        "$cpfRaw|$clienteRaw",
        "$cpfLimpo|$clienteRaw",
        "$cpfLimpo"
    ) | Where-Object { $_ -and -not [string]::IsNullOrWhiteSpace([string]$_) }
    foreach($key in $keys){
        if($ages.PSObject.Properties.Name -contains $key){ return $ages.$key }
    }
    return $null
}

function Set-AgeValue($ages, $alias, $cpf, $cliente, $idade) {
    $cpfRaw = [string]$cpf
    $cpfLimpo = ($cpfRaw -replace '[^\d]', '')
    $clienteRaw = [string]$cliente
    $keys = @(
        "$alias|$cpfRaw|$clienteRaw",
        "$alias|$cpfLimpo|$clienteRaw",
        "$alias|$cpfRaw|",
        "$alias|$cpfLimpo|",
        "$cpfRaw|$clienteRaw",
        "$cpfLimpo|$clienteRaw",
        "$cpfLimpo"
    ) | Where-Object { $_ -and -not [string]::IsNullOrWhiteSpace([string]$_) }
    foreach($key in $keys){
        $ages | Add-Member -NotePropertyName $key -NotePropertyValue ([string]$idade) -Force
    }
}


function Update-AgeForMatchingOrders($orders, $alias, $cpf, $cliente, $idade) {
    $cpfRef = ([string]$cpf) -replace '[^\d]', ''
    $clienteRef = ([string]$cliente).Trim()
    $out = @()
    foreach($o in @($orders)){
        if($o){
            $cpfO = ([string]$o.cpf) -replace '[^\d]', ''
            $clienteO = ([string]$o.cliente).Trim()
            $aliasOk = ([string]$o.alias -eq [string]$alias)
            $cpfOk = ($cpfRef -and $cpfO -and $cpfO -eq $cpfRef)
            $clienteOk = ((-not $cpfRef) -and $clienteRef -and $clienteO -eq $clienteRef)
            if($aliasOk -and ($cpfOk -or $clienteOk)){
                Set-OrderAge $o $idade
            }
        }
        $out += $o
    }
    return @($out)
}

function Apply-AgesToOrders($orders) {
    $out = @()
    foreach($o in @($orders)){
        if($o){
            if(Is-OrderAgeMissing $o){
                $idadeSalva = Get-AgeValue $o.alias $o.cpf $o.cliente
                if($null -ne $idadeSalva -and -not [string]::IsNullOrWhiteSpace([string]$idadeSalva)){
                    Set-OrderAge $o $idadeSalva
                }
            }
        }
        $out += $o
    }
    return @($out)
}
function Join-AddressParts($obj) {
    if(-not $obj -or -not ($obj -is [pscustomobject])){ return "" }
    $parts=@()
    foreach($f in @("endereco","logradouro","rua")){
        if($obj.PSObject.Properties.Name -contains $f -and $obj.$f){ $parts += [string]$obj.$f; break }
    }
    foreach($f in @("numero")){
        if($obj.PSObject.Properties.Name -contains $f -and $obj.$f){ $parts += [string]$obj.$f; break }
    }
    foreach($f in @("complemento")){
        if($obj.PSObject.Properties.Name -contains $f -and $obj.$f){ $parts += [string]$obj.$f }
    }
    foreach($f in @("bairro")){
        if($obj.PSObject.Properties.Name -contains $f -and $obj.$f){ $parts += [string]$obj.$f; break }
    }
    $city=@()
    foreach($f in @("municipio","cidade")){
        if($obj.PSObject.Properties.Name -contains $f -and $obj.$f){ $city += [string]$obj.$f; break }
    }
    foreach($f in @("uf","estado")){
        if($obj.PSObject.Properties.Name -contains $f -and $obj.$f){ $city += [string]$obj.$f; break }
    }
    if($city.Count -gt 0){ $parts += ($city -join "/") }
    foreach($f in @("cep")){
        if($obj.PSObject.Properties.Name -contains $f -and $obj.$f){ $parts += "CEP " + [string]$obj.$f; break }
    }
    return (($parts | Where-Object { $_ -and -not [string]::IsNullOrWhiteSpace([string]$_) }) -join ", ")
}




function Normalize-CanalText($v) {
    if($null -eq $v){ return "" }
    $t = [string]$v
    $t = $t.ToUpperInvariant()
    $t = $t.Normalize([System.Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach($ch in $t.ToCharArray()){
        if([System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [System.Globalization.UnicodeCategory]::NonSpacingMark){ [void]$sb.Append($ch) }
    }
    return $sb.ToString()
}

function Get-PropValue($obj, [string[]]$names) {
    if($null -eq $obj){ return $null }
    foreach($n in $names){
        foreach($p in $obj.PSObject.Properties){
            if((Normalize-CanalText $p.Name) -eq (Normalize-CanalText $n)){ return $p.Value }
        }
    }
    return $null
}

function Add-CanalTextParts($obj, [System.Collections.ArrayList]$parts, [int]$depth=0, [bool]$parentRelevant=$false) {
    if($null -eq $obj -or $depth -gt 8){ return }

    if($obj -is [string] -or $obj.GetType().IsPrimitive){
        if($parentRelevant){
            $text = [string]$obj
            if(-not [string]::IsNullOrWhiteSpace($text)){ [void]$parts.Add($text) }
        }
        return
    }

    if($obj -is [System.Collections.IEnumerable] -and -not ($obj -is [string])){
        foreach($x in @($obj)){ Add-CanalTextParts $x $parts ($depth + 1) $parentRelevant }
        return
    }

    # Campos que realmente indicam canal/origem/integração. Não usamos o alias da conta,
    # porque a mesma loja pode vender em Mercado Livre e Shopee ao mesmo tempo.
    $canalKeys = @(
        'loja','lojaVirtual','loja_virtual','numeroLoja','numero_loja','idLoja','id_loja','lojaId','loja_id',
        'marketplace','origem','origemVenda','origem_venda','canal','canalVenda','canal_venda','ecommerce',
        'intermediador','plataforma','integracao','integração','tipoIntegracao','tipo_integracao',
        'store','seller','pedidoLoja','pedido_loja','numeroPedidoLoja','numero_pedido_loja',
        'observacoes','observação','observacao','informacoesAdicionais','informacoes_adicionais'
    )
    $subKeys = @('descricao','descrição','nome','codigo','código','id','tipo','canal','origem','numero','loja')

    foreach($prop in $obj.PSObject.Properties){
        $propName = Normalize-CanalText $prop.Name
        $isCanalKey = $false
        foreach($w in $canalKeys){
            $wn = Normalize-CanalText $w
            if($propName -eq $wn -or $propName -like "*$wn*"){ $isCanalKey = $true; break }
        }

        $isSubKey = $false
        foreach($w in $subKeys){
            $wn = Normalize-CanalText $w
            if($propName -eq $wn -or $propName -like "*$wn*"){ $isSubKey = $true; break }
        }

        if($isCanalKey -or ($parentRelevant -and $isSubKey)){
            Add-CanalTextParts $prop.Value $parts ($depth + 1) ($isCanalKey -or $parentRelevant)
        }
    }
}

function Get-CanalCodigoFromText($texto) {
    $txt = Normalize-CanalText $texto
    if([string]::IsNullOrWhiteSpace($txt)){ return 'OUTROS' }

    if(
        $txt -match 'SHOPEE' -or
        $txt -match '\bSHOPE\b' -or
        $txt -match '\bSHP\b'
    ){ return 'SHOPEE' }

    if(
        $txt -match 'MERCADO\s*LIVRE' -or
        $txt -match 'MERCADOLIVRE' -or
        $txt -match 'MERCADO\s*ENVIOS' -or
        $txt -match 'MERCADO\s*PAGO' -or
        $txt -match 'MERCADO\s*SHOPS' -or
        $txt -match 'MERCADOSHOPS' -or
        $txt -match '\bMELI\b' -or
        $txt -match '\bME\d{6,}\b' -or
        $txt -match '\bML\b' -or
        $txt -match '\bMLB[0-9]+' -or
        $txt -match '\bMSHOPS?\b'
    ){
        return 'MERCADO_LIVRE'
    }

    return 'OUTROS'
}

function Get-CanalCodigo($alias, $obj) {
    $parts = New-Object System.Collections.ArrayList
    # Importante: alias da conta não é usado para decidir canal.
    # Uma mesma conta/loja pode ter pedidos Shopee e Mercado Livre.
    Add-CanalTextParts $obj $parts 0
    return Get-CanalCodigoFromText (($parts | Where-Object { $_ -and -not [string]::IsNullOrWhiteSpace([string]$_) }) -join ' ')
}

$script:LOJA_CANAL_CACHE = @{}
function Get-LojaRawForChannel($alias, $lojaId) {
    if(-not $lojaId -or [string]::IsNullOrWhiteSpace([string]$lojaId) -or [string]$lojaId -eq '0'){ return $null }
    $cacheKey = "$alias|$lojaId"
    if($script:LOJA_CANAL_CACHE.ContainsKey($cacheKey)){ return $script:LOJA_CANAL_CACHE[$cacheKey] }

    $headers = Get-AuthHeader $alias
    $tentativas = @(
        "https://api.bling.com.br/Api/v3/lojas/$lojaId",
        "https://api.bling.com.br/Api/v3/lojas/virtuais/$lojaId",
        "https://api.bling.com.br/Api/v3/lojas?ids[]=$lojaId",
        "https://api.bling.com.br/Api/v3/lojas/virtuais?ids[]=$lojaId"
    )

    foreach($url in $tentativas){
        try{
            $resp = Invoke-RestMethod -Method Get -Uri $url -Headers $headers
            $raw = if($resp.data){ $resp.data } else { $resp }
            # Se a API devolver uma lista, usa apenas a loja do ID do pedido.
            if($raw -is [System.Collections.IEnumerable] -and -not ($raw -is [string])){
                $match = $null
                foreach($x in @($raw)){
                    $xid = Get-PropValue $x @('id','codigo','código','numero')
                    if($xid -and [string]$xid -eq [string]$lojaId){ $match = $x; break }
                }
                if($match){ $raw = $match }
                elseif(@($raw).Count -eq 1){ $raw = @($raw)[0] }
                else { continue }
            }
            $script:LOJA_CANAL_CACHE[$cacheKey] = $raw
            return $raw
        } catch {}
    }

    $script:LOJA_CANAL_CACHE[$cacheKey] = $null
    return $null
}

function Get-LojaIdFromOrder($obj) {
    if($null -eq $obj){ return "" }
    $loja = Get-PropValue $obj @('loja','lojaVirtual','loja_virtual','store')
    if($loja){
        if($loja -is [string] -or $loja.GetType().IsPrimitive){ return [string]$loja }
        $id = Get-PropValue $loja @('id','codigo','código','numero')
        if($id){ return [string]$id }
    }
    $id2 = Get-PropValue $obj @('lojaId','loja_id','idLoja','id_loja','numeroLoja','numero_loja')
    if($id2){ return [string]$id2 }
    return ""
}

function Get-CanalCodigoPedido($alias, $obj) {
    # 1) Tenta pelo próprio pedido/detalhe.
    $canal = Get-CanalCodigo $null $obj
    if($canal -ne 'OUTROS'){ return $canal }

    # 2) Se o pedido só trouxe ID da loja, busca a loja/integração no Bling e classifica pelo nome/tipo dela.
    $lojaId = Get-LojaIdFromOrder $obj
    if($lojaId){
        try{
            $lojaRaw = Get-LojaRawForChannel $alias $lojaId
            $canalLoja = Get-CanalCodigo $null $lojaRaw
            if($canalLoja -ne 'OUTROS'){ return $canalLoja }
        } catch {}
    }

    return 'OUTROS'
}

function Get-OrderRawForChannel($alias, $pedidoId) {
    if(-not $pedidoId){ return $null }
    $headers = Get-AuthHeader $alias
    $url = "https://api.bling.com.br/Api/v3/pedidos/vendas/$pedidoId"
    $resp = Invoke-RestMethod -Method Get -Uri $url -Headers $headers
    if($resp.data){ return $resp.data }
    return $resp
}

function Get-CanalNome($codigo) {
    if($codigo -eq 'MERCADO_LIVRE'){ return 'Mercado Livre' }
    if($codigo -eq 'SHOPEE'){ return 'Shopee' }
    return 'Outros'
}

function Normalize-Order($alias, $item) {
    $pedidoId = ""; $pedidoNumero = ""
    if($item.id){ $pedidoId = [string]$item.id }
    if($item.numero){ $pedidoNumero = [string]$item.numero } else { $pedidoNumero = $pedidoId }
    $cliente=""; $cpf=""; $data=""; $valor=""; $nota=""; $idade=""

    if($item.contato){
        if($item.contato.nome){ $cliente = [string]$item.contato.nome }
        if($item.contato.numeroDocumento){ $cpf = [string]$item.contato.numeroDocumento }
        if($item.contato.cpf){ $cpf = [string]$item.contato.cpf }
    }
    if($item.cliente){
        if($item.cliente.nome){ $cliente = [string]$item.cliente.nome }
        if($item.cliente.numeroDocumento){ $cpf = [string]$item.cliente.numeroDocumento }
    }

    $idade = Get-AgeValue $alias $cpf $cliente
    if($null -eq $idade){ $idade = "" }

    if($item.data){ $data = [string]$item.data }
    if($item.dataCriacao){ $data = [string]$item.dataCriacao }
    if($item.total){ $valor = [string]$item.total }
    if($item.valor){ $valor = [string]$item.valor }
    if($item.notaFiscal){
        if($item.notaFiscal.numero){ $nota = [string]$item.notaFiscal.numero }
        elseif($item.notaFiscal.id){ $nota = [string]$item.notaFiscal.id }
    }

    $lojaId = Get-LojaIdFromOrder $item
    $canal = Get-CanalCodigoPedido $alias $item

    [pscustomobject]@{
        alias = $alias
        pedidoId = $pedidoId
        pedidoNumero = $pedidoNumero
        canal = $canal
        canalNome = Get-CanalNome $canal
        lojaId = $lojaId
        data = $data
        cliente = $cliente
        cpf = $cpf
        idade = $idade
        valor = $valor
        nota = $nota
    }
}
function Sync-Accounts($aliases) {
    $all = @()
    $selected = @()
    if($aliases){ $selected = @($aliases | ForEach-Object { [string]$_ }) }
    foreach($a in (Get-Accounts)){
        if(-not $a.access_token){ continue }
        if($selected.Count -gt 0 -and -not ($selected -contains [string]$a.alias)){ continue }
        $headers = Get-AuthHeader $a.alias
        for($pagina=1; $pagina -le 3; $pagina++){
            $url = "https://api.bling.com.br/Api/v3/pedidos/vendas?pagina=$pagina"
            $resp = Invoke-RestMethod -Method Get -Uri $url -Headers $headers
            $items = @()
            if($resp.data){ $items = $resp.data } elseif($resp.itens){ $items = $resp.itens }
            foreach($it in $items){
                $ord = Normalize-Order $a.alias $it
                if($ord.canal -eq 'OUTROS' -and $ord.pedidoId){
                    try{
                        $detailRaw = Get-OrderRawForChannel $a.alias $ord.pedidoId
                        $canalDetail = Get-CanalCodigoPedido $a.alias $detailRaw
                        if($canalDetail -ne 'OUTROS'){
                            $ord.canal = $canalDetail
                            $ord.canalNome = Get-CanalNome $canalDetail
                        }
                    } catch {}
                }
                $all += $ord
            }
            if(-not $items -or $items.Count -lt 100){ break }
            Start-Sleep -Milliseconds 400
        }
    }

    if($selected.Count -gt 0){
        $keep = @()
        foreach($o in (Get-Orders)){
            if(-not ($selected -contains [string]$o.alias)){ $keep += $o }
        }
        Save-Orders (@($keep) + @($all))
    } else {
        Save-Orders $all
    }
    return $all
}

function Sync-All {
    return Sync-Accounts @()
}

function Format-AddressFromGeneral($e) {
    if($null -eq $e){ return "" }

    $parts = @()
    if($e.endereco){ $parts += [string]$e.endereco }
    if($e.numero){ $parts += [string]$e.numero }
    if($e.complemento){ $parts += [string]$e.complemento }
    if($e.bairro){ $parts += [string]$e.bairro }

    $cidadeUf = ""
    if($e.municipio){ $cidadeUf += [string]$e.municipio }
    if($e.uf){
        if($cidadeUf){ $cidadeUf += "/" }
        $cidadeUf += [string]$e.uf
    }
    if($cidadeUf){ $parts += $cidadeUf }

    if($e.cep){ $parts += "CEP: $([string]$e.cep)" }

    return (($parts | Where-Object { $_ -and -not [string]::IsNullOrWhiteSpace([string]$_) }) -join " - ")
}

function Format-AddressFromEtiqueta($e) {
    if($null -eq $e){ return "" }

    $parts = @()
    if($e.nome){ $parts += [string]$e.nome }
    if($e.endereco){ $parts += [string]$e.endereco }
    if($e.numero){ $parts += [string]$e.numero }
    if($e.complemento){ $parts += [string]$e.complemento }
    if($e.bairro){ $parts += [string]$e.bairro }

    $cidadeUf = ""
    if($e.municipio){ $cidadeUf += [string]$e.municipio }
    if($e.uf){
        if($cidadeUf){ $cidadeUf += "/" }
        $cidadeUf += [string]$e.uf
    }
    if($cidadeUf){ $parts += $cidadeUf }

    if($e.cep){ $parts += "CEP: $([string]$e.cep)" }

    return (($parts | Where-Object { $_ -and -not [string]::IsNullOrWhiteSpace([string]$_) }) -join " - ")
}

function Get-ContactAddress($alias, $contactId) {
    if(-not $contactId -or [string]::IsNullOrWhiteSpace([string]$contactId) -or [string]$contactId -eq "0"){
        return ""
    }
    try {
        $headers = Get-AuthHeader $alias
        $url = "https://api.bling.com.br/Api/v3/contatos/$contactId"
        $resp = Invoke-RestMethod -Method Get -Uri $url -Headers $headers
        $raw = if($resp.data){ $resp.data } else { $resp }

        if($raw.endereco -and $raw.endereco.geral){
            $e = $raw.endereco.geral
            $parts = @()
            if($e.endereco){ $parts += [string]$e.endereco }
            if($e.numero){ $parts += [string]$e.numero }
            if($e.bairro){ $parts += [string]$e.bairro }

            $city = @()
            if($e.municipio){ $city += [string]$e.municipio }
            if($e.uf){ $city += [string]$e.uf }
            if($city.Count -gt 0){ $parts += ($city -join "/") }

            if($e.cep){ $parts += ("CEP " + [string]$e.cep) }
            if($e.complemento){ $parts += [string]$e.complemento }

            if($parts.Count -gt 0){ return ($parts -join ", ") }
        }

        return ""
    } catch {
        return ""
    }
}


function Get-OrderDetail($alias, $pedidoId) {
    $headers = Get-AuthHeader $alias
    $url = "https://api.bling.com.br/Api/v3/pedidos/vendas/$pedidoId"
    $resp = Invoke-RestMethod -Method Get -Uri $url -Headers $headers
    $raw = if($resp.data){ $resp.data } else { $resp }

    $cliente=""; $cpf=""; $data=""; $valor=""; $nota=""; $endereco=""; $itemDescricao=""; $rawHint=""; $idade=""; $canal=""
    $contactId = ""

    if($raw.contato){
        if($raw.contato.id){ $contactId=[string]$raw.contato.id }
        if($raw.contato.nome){ $cliente=[string]$raw.contato.nome }
        if($raw.contato.numeroDocumento){ $cpf=[string]$raw.contato.numeroDocumento }
        if($raw.contato.cpf){ $cpf=[string]$raw.contato.cpf }
    }
    if(-not $cliente -and $raw.cliente -and $raw.cliente.nome){ $cliente=[string]$raw.cliente.nome }
    if(-not $cpf -and $raw.cliente -and $raw.cliente.numeroDocumento){ $cpf=[string]$raw.cliente.numeroDocumento }

    $idade = Get-AgeValue $alias $cpf $cliente
    if($null -eq $idade){ $idade = "" }
    $lojaId = Get-LojaIdFromOrder $raw
    $canal = Get-CanalCodigoPedido $alias $raw

    if($raw.data){ $data=[string]$raw.data }
    if($raw.total){ $valor=[string]$raw.total } elseif($raw.valor){ $valor=[string]$raw.valor }
    if($raw.notaFiscal){
        if($raw.notaFiscal.numero){ $nota=[string]$raw.notaFiscal.numero }
        elseif($raw.notaFiscal.id){ $nota=[string]$raw.notaFiscal.id }
    }

    # 1) tenta pelo pedido (Shopee costuma vir aqui)
    if($raw.transporte -and $raw.transporte.etiqueta){
        $parts=@()
        foreach($f in @("endereco","logradouro","rua")){ if($raw.transporte.etiqueta.PSObject.Properties.Name -contains $f -and $raw.transporte.etiqueta.$f){ $parts += [string]$raw.transporte.etiqueta.$f; break } }
        foreach($f in @("numero")){ if($raw.transporte.etiqueta.PSObject.Properties.Name -contains $f -and $raw.transporte.etiqueta.$f){ $parts += [string]$raw.transporte.etiqueta.$f; break } }
        foreach($f in @("bairro")){ if($raw.transporte.etiqueta.PSObject.Properties.Name -contains $f -and $raw.transporte.etiqueta.$f){ $parts += [string]$raw.transporte.etiqueta.$f; break } }
        $city=@()
        foreach($f in @("municipio","cidade")){ if($raw.transporte.etiqueta.PSObject.Properties.Name -contains $f -and $raw.transporte.etiqueta.$f){ $city += [string]$raw.transporte.etiqueta.$f; break } }
        foreach($f in @("uf","estado")){ if($raw.transporte.etiqueta.PSObject.Properties.Name -contains $f -and $raw.transporte.etiqueta.$f){ $city += [string]$raw.transporte.etiqueta.$f; break } }
        if($city.Count -gt 0){ $parts += ($city -join "/") }
        foreach($f in @("cep")){ if($raw.transporte.etiqueta.PSObject.Properties.Name -contains $f -and $raw.transporte.etiqueta.$f){ $parts += ("CEP " + [string]$raw.transporte.etiqueta.$f); break } }
        if($parts.Count -gt 0){ $endereco = ($parts -join ", ") }
    }

    # 2) fallback pelo contato (Mercado Livre)
    if(-not $endereco){
        $endereco = Get-ContactAddress $alias $contactId
    }

    if(-not $endereco){ $endereco = "-" }

    if($raw.itens -and $raw.itens.Count -gt 0){
        if($raw.itens[0].descricao){ $itemDescricao=[string]$raw.itens[0].descricao }
        elseif($raw.itens[0].descricaoDetalhada){ $itemDescricao=[string]$raw.itens[0].descricaoDetalhada }
        elseif($raw.itens[0].codigo){ $itemDescricao=[string]$raw.itens[0].codigo }
    }

    if(-not $itemDescricao){ $rawHint = "Item não veio no formato esperado." }
    if($endereco -eq "-"){
        if($rawHint){ $rawHint += " " }
        $rawHint += "Endereço não veio no pedido e também não foi localizado no contato do Bling."
    }

    return [pscustomobject]@{
        alias = $alias
        pedido = $pedidoId
        canal = $canal
        canalNome = Get-CanalNome $canal
        lojaId = $lojaId
        cliente = $cliente
        cpf = $cpf
        data = $data
        valor = $valor
        nota = $nota
        endereco = $endereco
        idade = $idade
        itemDescricao = $itemDescricao
        subtotal = $valor
        total = $valor
        rawHint = $rawHint
    }
}


$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://*:$port/")
$listener.Start()
if(-not ($env:RENDER -or $env:RENDER_EXTERNAL_URL) -and $IsWindows){ Start-Process $publicBaseUrl }

try {
    while($listener.IsListening){
        $ctx = $listener.GetContext()
        $req = $ctx.Request
        $path = $req.Url.AbsolutePath
        try {
            if($path -eq "/"){
                $html = (Get-Content (Join-Path $baseDir "index.html")) -join "`n"
                Send-Text $ctx $html "text/html; charset=utf-8"
                continue
            }
            if($path -eq "/api/auth/login" -and $req.HttpMethod -eq "POST"){
                $body = Read-BodyJson $req
                $usuario = ([string]$body.usuario).Trim().ToLower()
                $senha = [string]$body.senha
                $u = Find-User $usuario
                if(-not $u -or [string]$u.senha -ne $senha){ Send-Json $ctx @{error='Usuário ou senha inválidos.'} 401; continue }
                if([bool]$u.bloqueado){ Send-Json $ctx @{error='Usuário bloqueado.'} 403; continue }
                $session = New-LoginSession $u
                Send-Json $ctx @{ok=$true; session=$session; user=(To-PublicUser $u)}
                continue
            }
            if($path -eq "/api/auth/logout" -and $req.HttpMethod -eq "POST"){
                $token = $req.Headers['X-Auth-Token']
                if($token){ Remove-SessionToken $token }
                Send-Json $ctx @{ok=$true}
                continue
            }
            if($path -eq "/api/auth/me"){
                $session = Get-SessionFromRequest $req
                if(-not $session){ Send-Json $ctx @{error='Sessão inválida.'} 401; continue }
                $u = Find-User $session.usuario
                if(-not $u -or [bool]$u.bloqueado){ Send-Json $ctx @{error='Sessão inválida.'} 401; continue }
                Send-Json $ctx @{ok=$true; user=(To-PublicUser $u); session=$session}
                continue
            }
            if($path -eq "/api/users"){
                $null = Require-Admin $req
                $public = @()
                foreach($u in (Get-Users)){ $public += (To-PublicUser $u) }
                Send-Json $ctx @{users=$public}
                continue
            }
            if($path -eq "/api/users/add" -and $req.HttpMethod -eq "POST"){
                $null = Require-Admin $req
                $body = Read-BodyJson $req
                Add-AppUser ([string]$body.usuario).Trim() ([string]$body.senha) ([string]$body.tipo) ([string]$body.nome)
                $public = @(); foreach($u in (Get-Users)){ $public += (To-PublicUser $u) }
                Send-Json $ctx @{ok=$true; users=$public}
                continue
            }
            if($path -eq "/api/users/update" -and $req.HttpMethod -eq "POST"){
                $null = Require-Admin $req
                $body = Read-BodyJson $req
                $bloq = $null
                if($body.PSObject.Properties.Name -contains 'bloqueado'){ $bloq = [bool]$body.bloqueado }
                Update-AppUser ([string]$body.usuario).Trim() ([string]$body.senha) ([string]$body.tipo) ([string]$body.nome) $bloq
                $public = @(); foreach($u in (Get-Users)){ $public += (To-PublicUser $u) }
                Send-Json $ctx @{ok=$true; users=$public}
                continue
            }
            if($path -eq "/api/users/delete" -and $req.HttpMethod -eq "POST"){
                $null = Require-Admin $req
                $body = Read-BodyJson $req
                Delete-AppUser ([string]$body.usuario).Trim()
                $public = @(); foreach($u in (Get-Users)){ $public += (To-PublicUser $u) }
                Send-Json $ctx @{ok=$true; users=$public}
                continue
            }
            if($path -eq "/api/account" -and $req.HttpMethod -eq "POST"){
                $null = Require-Admin $req
                $body = Read-BodyJson $req
                Upsert-Account $body.alias $body.clientId $body.clientSecret
                Send-Json $ctx @{ok=$true}
                continue
            }
            if($path -eq "/api/account/delete" -and $req.HttpMethod -eq "POST"){
                $null = Require-Admin $req
                $body = Read-BodyJson $req
                $remaining = Delete-AccountByAlias $body.alias
                Send-Json $ctx @{ok=$true; remainingOrders=$remaining}
                continue
            }
            if($path -eq "/api/accounts"){
                $out = @()
                foreach($a in (Get-Accounts)){
                    $out += [pscustomobject]@{ alias=$a.alias; clientId=$a.clientId; connected=[bool]$a.access_token }
                }
                Send-Json $ctx @{accounts=$out}
                continue
            }
            if($path -eq "/api/login"){
                $alias = $req.QueryString["alias"]
                $a = Find-Account $alias
                if(-not $a){ throw "Conta não encontrada." }
                $pending = [pscustomobject]@{ alias=$alias; state=[guid]::NewGuid().ToString("N") }
                Write-JsonFile $pendingPath $pending
                $redir = [System.Uri]::EscapeDataString("$publicBaseUrl/callback")
                $loc = "https://www.bling.com.br/Api/v3/oauth/authorize?response_type=code&client_id=$($a.clientId)&state=$($pending.state)&redirect_uri=$redir"
                $ctx.Response.StatusCode = 302
                $ctx.Response.RedirectLocation = $loc
                $ctx.Response.Close()
                continue
            }
            if($path -eq "/callback"){
                $pending = Read-JsonFile $pendingPath ([pscustomobject]@{})
                $code = $req.QueryString["code"]
                $state = $req.QueryString["state"]
                if(-not $code){ throw "Bling não retornou authorization code." }
                if($pending.state -and $state -ne $pending.state){ throw "State inválido." }
                $a = Find-Account $pending.alias
                if(-not $a){ throw "Conta pendente não encontrada." }
                $headers = @{ "Authorization" = Get-BasicAuthHeader $a.clientId $a.clientSecret; "enable-jwt"="1" }
                $body = @{ grant_type="authorization_code"; code=$code }
                $resp = Invoke-RestMethod -Method Post -Uri "https://www.bling.com.br/Api/v3/oauth/token" -Headers $headers -Body $body -ContentType "application/x-www-form-urlencoded"

                $new = @()
                foreach($x in (Get-Accounts)){
                    if($x.alias -eq $a.alias){
                        $x.access_token = $resp.access_token
                        $x.refresh_token = $resp.refresh_token
                        $x.expires_at = (Get-Date).AddSeconds([int]$resp.expires_in).ToString("o")
                    }
                    $new += $x
                }
                Save-Accounts $new
                Send-Text $ctx "<html><body style='font-family:Segoe UI;background:#0b1220;color:#fff;padding:30px'><h2>Conta conectada com sucesso.</h2><p>Volte para o app e clique em Sincronizar todas.</p></body></html>" "text/html; charset=utf-8"
                continue
            }
            if($path -eq "/api/sync-all"){
                $null = Require-Admin $req
                $orders = Sync-All
                Send-Json $ctx @{ok=$true; count=$orders.Count}
                continue
            }
            if($path -eq "/api/sync-selected" -and $req.HttpMethod -eq "POST"){
                $null = Require-Auth $req
                $body = Read-BodyJson $req
                $aliases = @()
                if($body.aliases){ $aliases = @($body.aliases) }
                if($aliases.Count -eq 0){ throw "Selecione pelo menos uma loja para sincronizar." }
                $orders = Sync-Accounts $aliases
                Send-Json $ctx @{ok=$true; count=$orders.Count; aliases=$aliases}
                continue
            }
            if($path -eq "/api/orders"){
                $orders = Get-Orders
                Send-Json $ctx @{orders=$orders}
                continue
            }
            
            if($path -match "^/api/fdx-debug/(.+)$"){
                $cpf = [System.Uri]::UnescapeDataString($Matches[1])
                $idade = BuscarIdadeCPF $cpf
                $foneDbg = Get-FdxDebugData $cpf
                Send-Json $ctx @{
                    ok = $true
                    idade = $idade
                    debug = $script:FDX_DEBUG
                    phones = $foneDbg.phones
                    historicoTelefone = $foneDbg.historicoTelefone
                    ultimoTelefone = $foneDbg.ultimoTelefone
                    rawJson = $foneDbg.rawJson
                    erro = $foneDbg.erro
                }
                continue
            }


if($path -eq "/api/ages-by-period" -and $req.HttpMethod -eq "POST"){
    $body = Read-BodyJson $req
    $startDate = Parse-OrderDate $body.startDate
    $endDate = Parse-OrderDate $body.endDate
    if($null -eq $startDate -or $null -eq $endDate){
        Send-Json $ctx @{error="Período inválido"} 400
        continue
    }

    $endDate = $endDate.Date.AddDays(1).AddSeconds(-1)
    $orders = Get-Orders
    $aliasesFiltro = @()
    if($body.aliases){
        $aliasesFiltro = @($body.aliases | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }
    if($aliasesFiltro.Count -gt 0){
        $orders = @($orders | Where-Object { $aliasesFiltro -contains [string]$_.alias })
    }
    $ages = Get-Ages
    $updated = @()
    $processados = 0
    $encontrados = 0

    foreach($o in $orders){
        $od = Parse-OrderDate $o.data
        if($od -and $od -ge $startDate.Date -and $od -le $endDate){
            $processados++
        }
    }

    foreach($o in $orders){
        $od = Parse-OrderDate $o.data
        if($od -and $od -ge $startDate.Date -and $od -le $endDate){
            if((Is-OrderAgeMissing $o) -and $o.cpf){
                $idadeResp = BuscarIdadeCPF $o.cpf
                if($idadeResp -ne $null){
                    Set-OrderAge $o $idadeResp
                    Set-AgeValue $ages $o.alias $o.cpf $o.cliente $o.idade
                    $updated = @(Update-AgeForMatchingOrders $updated $o.alias $o.cpf $o.cliente $o.idade)
                    $encontrados++
                }
            }
        }
        $updated += $o
    }

    Write-JsonFile $agesPath $ages
    Save-Orders $updated
    Send-Json $ctx @{
        ok = $true
        processados = $processados
        encontrados = $encontrados
    }
    continue
}

if($path -eq "/api/ages-by-period-stream" -and $req.HttpMethod -eq "POST"){
    $ctx.Response.StatusCode = 200
    $ctx.Response.ContentType = "application/x-ndjson; charset=utf-8"
    try {
        $body = Read-BodyJson $req
        $startDate = Parse-OrderDate $body.startDate
        $endDate = Parse-OrderDate $body.endDate
        if($null -eq $startDate -or $null -eq $endDate){
            Write-StreamJson $ctx @{ type="error"; error="Período inválido" }
            $ctx.Response.OutputStream.Close()
            continue
        }

        $endDate = $endDate.Date.AddDays(1).AddSeconds(-1)
        $orders = Get-Orders
        $aliasesFiltro = @()
        if($body.aliases){
            $aliasesFiltro = @($body.aliases | ForEach-Object { [string]$_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        }
        if($aliasesFiltro.Count -gt 0){
            $orders = @($orders | Where-Object { $aliasesFiltro -contains [string]$_.alias })
        }
        $ages = Get-Ages
        $updated = @()
        $processados = 0
        $encontrados = 0
        $total = 0

        foreach($o in $orders){
            $od = Parse-OrderDate $o.data
            if($od -and $od -ge $startDate.Date -and $od -le $endDate){
                $total++
            }
        }

        if($total -eq 0){
            Write-StreamJson $ctx @{ type="progress"; percent=100; message="Nenhum pedido no período." }
            Write-StreamJson $ctx @{ type="done"; ok=$true; processados=0; encontrados=0 }
            $ctx.Response.OutputStream.Close()
            continue
        }

        $atual = 0
        Write-StreamJson $ctx @{ type="progress"; percent=0; message="Iniciando consulta..." }

        foreach($o in $orders){
            $od = Parse-OrderDate $o.data
            if($od -and $od -ge $startDate.Date -and $od -le $endDate){
                $atual++
                $processados++
                if((Is-OrderAgeMissing $o) -and $o.cpf){
                    $idadeResp = BuscarIdadeCPF $o.cpf
                    if($idadeResp -ne $null){
                        Set-OrderAge $o $idadeResp
                        Set-AgeValue $ages $o.alias $o.cpf $o.cliente $o.idade
                        $updated = @(Update-AgeForMatchingOrders $updated $o.alias $o.cpf $o.cliente $o.idade)
                        $encontrados++
                    }
                }
                $percent = [math]::Floor(($atual / $total) * 100)
                $nomeCliente = if($o.cliente){ [string]$o.cliente } else { "-" }
                Write-StreamJson $ctx @{
                    type="progress"
                    percent=$percent
                    message=("Processando {0} de {1}: {2}" -f $atual, $total, $nomeCliente)
                }
            }
            $updated += $o
        }

        Write-JsonFile $agesPath $ages
        Save-Orders $updated
        Write-StreamJson $ctx @{ type="done"; ok=$true; processados=$processados; encontrados=$encontrados }
    } catch {
        Write-StreamJson $ctx @{ type="error"; error=$_.Exception.Message }
    }
    $ctx.Response.OutputStream.Close()
    continue
}

if($path -match "^/api/contact-debug/(.+?)/(.+)$"){
    $alias = [System.Uri]::UnescapeDataString($Matches[1])
    $pedidoId = [System.Uri]::UnescapeDataString($Matches[2])

    $orders = Get-Orders
    $cpf = ""
    foreach($o in $orders){
        if($o.alias -eq $alias -and [string]$o.pedidoId -eq [string]$pedidoId){
            $cpf = [string]$o.cpf
            break
        }
    }

    if(-not $cpf){
        # fallback: tenta achar pelo detalhe do pedido
        try {
            $headers = Get-AuthHeader $alias
            $url = "https://api.bling.com.br/Api/v3/pedidos/vendas/$pedidoId"
            $resp = Invoke-RestMethod -Method Get -Uri $url -Headers $headers
            $raw = if($resp.data){ $resp.data } else { $resp }
            if($raw.contato -and $raw.contato.numeroDocumento){ $cpf = [string]$raw.contato.numeroDocumento }
            elseif($raw.contato -and $raw.contato.cpf){ $cpf = [string]$raw.contato.cpf }
            elseif($raw.cliente -and $raw.cliente.numeroDocumento){ $cpf = [string]$raw.cliente.numeroDocumento }
        } catch {}
    }

    $dbg = Get-FdxDebugData $cpf

    Send-Json $ctx @{
        ok = $true
        pedidoId = $pedidoId
        cpf = $dbg.cpf
        contactId = "-"
        phones = $dbg.phones
        historicoTelefone = $dbg.historicoTelefone
        ultimoTelefone = $dbg.ultimoTelefone
        rawJson = $dbg.rawJson
        erro = $dbg.erro
    }
    continue
}

if($path -match "^/api/order/(.+?)/(.+)$"){
                $alias = [System.Uri]::UnescapeDataString($Matches[1])
                $pedidoId = [System.Uri]::UnescapeDataString($Matches[2])
                $detail = Get-OrderDetail $alias $pedidoId
                Send-Json $ctx @{detail=$detail}
                continue
            }
            if($path -eq "/api/age" -and $req.HttpMethod -eq "POST"){
                $body = Read-BodyJson $req
                $ages = Get-Ages
                Set-AgeValue $ages $body.alias $body.cpf $body.cliente $body.idade
                Write-JsonFile $agesPath $ages

                $orders = Get-Orders
                $updated = @(Update-AgeForMatchingOrders $orders $body.alias $body.cpf $body.cliente $body.idade)
                Save-Orders $updated
                Send-Json $ctx @{ok=$true}
                continue
            }
            if($path -eq "/api/clear" -and $req.HttpMethod -eq "POST"){
                $null = Require-Admin $req
                Save-Orders @()
                foreach($p in @($agesPath,$pendingPath)){ if(Test-Path $p){ Remove-Item $p -Force } }
                Send-Json $ctx @{ok=$true}
                continue
            }
            Send-Json $ctx @{error="Rota não encontrada"} 404
        } catch {
            Send-Json $ctx @{error=$_.Exception.Message} 500
        }
    }
} finally {
    $listener.Stop()
    $listener.Close()
}
