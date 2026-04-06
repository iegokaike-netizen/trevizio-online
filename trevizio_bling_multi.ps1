

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

        $nasc = Find-FirstFieldValue $resp @("NASC", "NASCIMENTO")
        $script:FDX_DEBUG.nascimentoBruto = [string]$nasc

        if($nasc -and $nasc -ne "SEM RESULTADO"){
            try {
                $dataNasc = [datetime]::ParseExact([string]$nasc, "dd/MM/yyyy", $null)
                $hoje = Get-Date
                $idade = $hoje.Year - $dataNasc.Year

                if(($hoje.Month -lt $dataNasc.Month) -or (($hoje.Month -eq $dataNasc.Month) -and ($hoje.Day -lt $dataNasc.Day))){
                    $idade--
                }

                $script:FDX_DEBUG.idadeEncontrada = $idade
                return $idade
            } catch {
                $script:FDX_DEBUG.erro = $_.Exception.Message
                return $null
            }
        }

        return $null
    } catch {
        $script:FDX_DEBUG.erro = $_.Exception.Message
        return $null
    }
}

$FDX_TOKEN = "499490c3b21042fe891dafb2c067ef9d"

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



function Read-JsonFile($path, $defaultObj) {
    if (Test-Path $path) { try { return (((Get-Content $path) -join "`n") | ConvertFrom-Json) } catch { return $defaultObj } }
    return $defaultObj
}
function Write-JsonFile($path, $obj) { $obj | ConvertTo-Json -Depth 60 | Set-Content -Encoding UTF8 $path }
function Get-Accounts {
    $a = Read-JsonFile $accountsPath @()
    if ($a -is [string]) { return @() }
    if ($null -eq $a) { return @() }
    return @($a)
}
function Save-Accounts($accounts) { Write-JsonFile $accountsPath $accounts }
function Get-Ages { Read-JsonFile $agesPath ([pscustomobject]@{}) }
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
function Read-BodyJson($req) {
    $sr = New-Object System.IO.StreamReader($req.InputStream, $req.ContentEncoding)
    $body = $sr.ReadToEnd()
    if ([string]::IsNullOrWhiteSpace($body)) { return $null }
    return ($body | ConvertFrom-Json)
}
function Find-Account($alias) {
    foreach($a in (Get-Accounts)){ if($a.alias -eq $alias){ return $a } }
    return $null
}
function Upsert-Account($alias, $clientId, $clientSecret) {
    if (-not $alias) { throw "Alias da loja é obrigatório." }
    $accounts = Get-Accounts
    $found = $false
    $new = @()
    foreach($a in $accounts){
        if($a.alias -eq $alias){
            if($clientId){ $a.clientId = $clientId }
            if($clientSecret){ $a.clientSecret = $clientSecret }
            $new += $a
            $found = $true
        } else { $new += $a }
    }
    if(-not $found){
        $new += [pscustomobject]@{
            alias = $alias
            clientId = $clientId
            clientSecret = $clientSecret
            redirectUri = "$publicBaseUrl/callback"
            access_token = ""
            refresh_token = ""
            expires_at = ""
        }
    }
    Save-Accounts $new
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
    $key = "$alias|$cpf|$cliente"
    if($ages.PSObject.Properties.Name -contains $key){ return $ages.$key }
    return $null
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
function Find-AddressCandidates($obj) {
    $results = New-Object System.Collections.ArrayList
    if($null -eq $obj){ return @() }

    if($obj -is [pscustomobject]) {
        $candidate = Join-AddressParts $obj
        if($candidate){ [void]$results.Add($candidate) }

        foreach($propName in @("enderecoEntrega","endereco","entrega","etiqueta","destinatario","cliente","contato","geral","principal")){
            if($obj.PSObject.Properties.Name -contains $propName){
                $val = $obj.$propName
                if($val -is [pscustomobject]){
                    $candidate = Join-AddressParts $val
                    if($candidate){ [void]$results.Add($candidate) }
                }
            }
        }

        foreach($p in $obj.PSObject.Properties){
            $val = $p.Value
            if($val -is [pscustomobject] -or (($val -is [System.Collections.IEnumerable]) -and -not ($val -is [string]))){
                foreach($x in (Find-AddressCandidates $val)){
                    if($x){ [void]$results.Add($x) }
                }
            }
        }
    }
    elseif(($obj -is [System.Collections.IEnumerable]) -and -not ($obj -is [string])) {
        foreach($item in $obj){
            foreach($x in (Find-AddressCandidates $item)){
                if($x){ [void]$results.Add($x) }
            }
        }
    }

    return @($results | Select-Object -Unique)
}
function Get-BestAddress($obj) {
    $candidates = Find-AddressCandidates $obj
    foreach($c in $candidates){
        if($c -and $c.Length -ge 8){ return $c }
    }
    return ""
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

    [pscustomobject]@{
        alias = $alias
        pedidoId = $pedidoId
        pedidoNumero = $pedidoNumero
        data = $data
        cliente = $cliente
        cpf = $cpf
        idade = $idade
        valor = $valor
        nota = $nota
    }
}
function Sync-All {
    $all = @()
    foreach($a in (Get-Accounts)){
        if(-not $a.access_token){ continue }
        $headers = Get-AuthHeader $a.alias
        for($pagina=1; $pagina -le 3; $pagina++){
            $url = "https://api.bling.com.br/Api/v3/pedidos/vendas?pagina=$pagina"
            $resp = Invoke-RestMethod -Method Get -Uri $url -Headers $headers
            $items = @()
            if($resp.data){ $items = $resp.data } elseif($resp.itens){ $items = $resp.itens }
            foreach($it in $items){ $all += (Normalize-Order $a.alias $it) }
            if(-not $items -or $items.Count -lt 100){ break }
            Start-Sleep -Milliseconds 400
        }
    }
    Write-JsonFile $ordersPath $all
    return $all
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

        $addr = Get-BestAddress $raw
        if($addr){ return $addr }

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

    $cliente=""; $cpf=""; $data=""; $valor=""; $nota=""; $endereco=""; $itemDescricao=""; $rawHint=""; $idade=""
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

    if($raw.data){ $data=[string]$raw.data }
    if($raw.total){ $valor=[string]$raw.total } elseif($raw.valor){ $valor=[string]$raw.valor }
    if($raw.notaFiscal){
        if($raw.notaFiscal.numero){ $nota=[string]$raw.notaFiscal.numero }
        elseif($raw.notaFiscal.id){ $nota=[string]$raw.notaFiscal.id }
    }

    # 1) tenta achar no pedido inteiro (Shopee / Mercado Livre / outros)
    $endereco = Get-BestAddress $raw

    # 2) fallback pelo contato
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
            if($path -eq "/api/account" -and $req.HttpMethod -eq "POST"){
                $body = Read-BodyJson $req
                Upsert-Account $body.alias $body.clientId $body.clientSecret
                Send-Json $ctx @{ok=$true}
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
                $orders = Sync-All
                Send-Json $ctx @{ok=$true; count=$orders.Count}
                continue
            }
            if($path -eq "/api/orders"){
                $orders = Read-JsonFile $ordersPath @()
                Send-Json $ctx @{orders=$orders}
                continue
            }
            
            if($path -match "^/api/fdx-debug/(.+)$"){
                $cpf = [System.Uri]::UnescapeDataString($Matches[1])
                $idade = BuscarIdadeCPF $cpf
                Send-Json $ctx @{
                    ok = $true
                    idade = $idade
                    debug = $script:FDX_DEBUG
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
    $orders = Read-JsonFile $ordersPath @()
    $ages = Get-Ages
    $updated = @()
    $processados = 0
    $encontrados = 0

    foreach($o in $orders){
        $od = Parse-OrderDate $o.data
        if($od -and $od -ge $startDate.Date -and $od -le $endDate){
            $processados++
            if(((-not $o.idade) -or [string]::IsNullOrWhiteSpace([string]$o.idade)) -and $o.cpf){
                $idadeResp = BuscarIdadeCPF $o.cpf
                if($idadeResp -ne $null){
                    $o.idade = [string]$idadeResp
                    $key = "$($o.alias)|$($o.cpf)|$($o.cliente)"
                    $ages | Add-Member -NotePropertyName $key -NotePropertyValue $o.idade -Force
                    $encontrados++
                }
            }
        }
        $updated += $o
    }

    Write-JsonFile $agesPath $ages
    Write-JsonFile $ordersPath $updated
    Send-Json $ctx @{
        ok = $true
        processados = $processados
        encontrados = $encontrados
    }
    continue
}

if($path -match "^/api/contact-debug/(.+?)/(.+)$"){
    $alias = [System.Uri]::UnescapeDataString($Matches[1])
    $pedidoId = [System.Uri]::UnescapeDataString($Matches[2])

    $orders = Read-JsonFile $ordersPath @()
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
                $key = "$($body.alias)|$($body.cpf)|$($body.cliente)"
                $ages | Add-Member -NotePropertyName $key -NotePropertyValue $body.idade -Force
                Write-JsonFile $agesPath $ages

                $orders = Read-JsonFile $ordersPath @()
                $updated = @()
                foreach($o in $orders){
                    if($o.alias -eq $body.alias -and ((($o.cpf) -and $o.cpf -eq $body.cpf) -or ((-not $o.cpf) -and $o.cliente -eq $body.cliente))){
                        $o.idade = $body.idade
                    }
                    $updated += $o
                }
                Write-JsonFile $ordersPath $updated
                Send-Json $ctx @{ok=$true}
                continue
            }
            if($path -eq "/api/clear" -and $req.HttpMethod -eq "POST"){
                foreach($p in @($accountsPath,$ordersPath,$agesPath,$pendingPath)){ if(Test-Path $p){ Remove-Item $p -Force } }
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
