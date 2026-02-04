<#
.SYNOPSIS
  Rimuove tutte le versioni dei file (eccetto l’ultima) in una cartella OneDrive/SharePoint, opzionalmente ricorsivo.

.DESCRIPTION
  - Usa PnP.PowerShell e CSOM per caricare in modo affidabile le versioni dei file.
  - Mantiene SEMPRE la versione più recente; opzionalmente rimuove solo versioni più vecchie di N giorni.
  - Supporta -WhatIf e retry/backoff in caso di throttling (429/503).
  - Logging migliorato con livelli e log su file opzionale.
  - Supporta autenticazione con Entra ID App tramite -ClientId (e opzionale -TenantId).

.PARAMETER OneDriveSiteUrl
  URL completo del sito OneDrive personale (es. https://<tenant>-my.sharepoint.com/personal/<upn_sostituito_con_underscore>)

.PARAMETER FolderSiteRelativeUrl
  Percorso site-relative della cartella (senza slash iniziale), es. "Documents/Virtual Machines/Windows 10 VM".
  Nota: lo script converte automaticamente "Documenti/..." in "Documents/...".

.PARAMETER IncludeSubfolders
  Se specificato, elabora ricorsivamente le sottocartelle.

.PARAMETER OlderThanDays
  Se specificato (>0), rimuove solo versioni più vecchie di N giorni (mantiene comunque l’ultima).

.PARAMETER LogPath
  Percorso file per scrivere il log (opzionale). Se non specificato, il log va solo a console.

.PARAMETER ExcludeOneNote
  Esclude i file OneNote (.one, .onetoc2).

.PARAMETER ClientId
  (NUOVO) Client ID dell'app Entra ID da usare per Connect-PnPOnline -Interactive.
  Se omesso, verrà usato il flusso interattivo predefinito di PnP.PowerShell.

.PARAMETER TenantId
  (NUOVO) Tenant ID (GUID) del tenant Entra ID. Consigliato se usi -ClientId per connessioni determinate.
  Se omesso, PnP tenterà la risoluzione automatica.

.EXAMPLES
  .\Purge-OneDriveFileVersions.ps1 `
    -OneDriveSiteUrl "https://tenant365-my.sharepoint.com/personal/UPN_utente_dominio_com" `
    -FolderSiteRelativeUrl "Documents/Virtual Machines/Windows 20 VM" `
    -IncludeSubfolders -WhatIf -Verbose

  # Con autenticazione tramite App (consigliato con PnP recenti)
  .\Purge-OneDriveFileVersions.ps1 `
    -OneDriveSiteUrl "https://tenant365-my.sharepoint.com/personal/UPN_utente_dominio_com" `
    -FolderSiteRelativeUrl "Documents/Operations" `
    -ClientId "00000000-0000-0000-0000-000000000000" `
    -TenantId "11111111-1111-1111-1111-111111111111" `
    -IncludeSubfolders -OlderThanDays 60 -LogPath "C:\Temp\PurgeVersions.log" -Verbose
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $OneDriveSiteUrl,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string] $FolderSiteRelativeUrl,

    [switch] $IncludeSubfolders,

    [ValidateRange(0, 36500)]
    [int] $OlderThanDays = 0,

    [string] $LogPath,

    [switch] $ExcludeOneNote,

    # =========================
    # NUOVI PARAMETRI AUTH PNP
    # =========================
    [string] $ClientId,
    [string] $TenantId
)

#region Logging helpers
$script:LogLevels = @{
    "DEBUG" = 0
    "INFO"  = 1
    "WARN"  = 2
    "ERROR" = 3
}
$script:MinLogLevel = if ($PSBoundParameters.ContainsKey('Verbose')) { 0 } else { 1 }

function Write-Log {
    param(
        [Parameter(Mandatory=$true)][ValidateSet("DEBUG","INFO","WARN","ERROR")] [string] $Level,
        [Parameter(Mandatory=$true)][string] $Message
    )
    if ($script:LogLevels[$Level] -lt $script:MinLogLevel) { return }

    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss.fff")
    $line = "[$ts] [$Level] $Message"

    switch ($Level) {
        "DEBUG" { Write-Host $line -ForegroundColor DarkGray }
        "INFO"  { Write-Host $line -ForegroundColor Cyan }
        "WARN"  { Write-Host $line -ForegroundColor Yellow }
        "ERROR" { Write-Host $line -ForegroundColor Red }
    }

    if ($LogPath) {
        try   { Add-Content -Path $LogPath -Value $line -ErrorAction Stop }
        catch { Write-Host "[$ts] [WARN] Impossibile scrivere su log file '$LogPath': $($_.Exception.Message)" -ForegroundColor Yellow }
    }
}
#endregion

#region Retry helper
function Invoke-WithRetry {
    param(
        [Parameter(Mandatory=$true)] [scriptblock] $ScriptBlock,
        [int] $MaxRetries = 5
    )
    $attempt = 0
    while ($true) {
        try { return & $ScriptBlock }
        catch {
            $attempt++
            $msg = $_.Exception.Message
            if ($attempt -le $MaxRetries -and ($msg -match '429' -or $msg -match '503' -or $msg -match 'throttl')) {
                $delay = [Math]::Pow(2, $attempt)
                Write-Log -Level "WARN" -Message "Throttling/rate limit (tentativo $attempt/$MaxRetries): $msg. Backoff ${delay}s"
                Start-Sleep -Seconds $delay
                continue
            } else {
                Write-Log -Level "ERROR" -Message "Operazione fallita dopo $attempt tentativi: $msg"
                throw
            }
        }
    }
}
#endregion

#region URL helpers
function Get-SiteRootFromAbsoluteUrl {
    param([Parameter(Mandatory=$true)][string]$AbsoluteUrl)
    $uri = [uri]$AbsoluteUrl
    return $uri.AbsolutePath.TrimEnd('/')  # es. "/personal/UPN_utente_dominio_com""
}

function Convert-ServerRelativeToSiteRelative {
    param(
        [Parameter(Mandatory=$true)][string]$ServerRelativeUrl,
        [Parameter(Mandatory=$true)][string]$SiteRootPath
    )
    $srv  = $ServerRelativeUrl.TrimStart('/')
    $root = $SiteRootPath.TrimStart('/')

    if ($srv.StartsWith($root)) { return $srv.Substring($root.Length).TrimStart('/') }
    else                        { return $srv }
}

function Ensure-ServerRelativeFileUrl {
    param(
        [Parameter(Mandatory=$true)][string]$OneDriveSiteUrl,
        [Parameter(Mandatory=$true)][string]$FolderSiteRelativeUrl,
        [Parameter(Mandatory=$true)][string]$FileName
    )
    $siteRoot = Get-SiteRootFromAbsoluteUrl -AbsoluteUrl $OneDriveSiteUrl
    $siteRel  = $FolderSiteRelativeUrl.TrimStart('/')
    return "$siteRoot/$siteRel/$FileName"
}
#endregion

#region Normalizzazione input & connettività
Write-Log -Level "INFO" -Message "Avvio purga versioni in sito: $OneDriveSiteUrl"
if ($PSBoundParameters.ContainsKey('LogPath')) {
    Write-Log -Level "INFO" -Message "Scrittura log su file: $LogPath"
}

# Normalizza "Documenti" -> "Documents"
$FolderSiteRelativeUrl = $FolderSiteRelativeUrl.TrimStart('/')
if ($FolderSiteRelativeUrl -like 'Documenti/*') {
    Write-Log -Level "WARN" -Message "Rilevato prefisso 'Documenti/'. Converto automaticamente in 'Documents/'."
    $FolderSiteRelativeUrl = 'Documents/' + $FolderSiteRelativeUrl.Substring(10)
}
if ($FolderSiteRelativeUrl -eq 'Documenti') {
    Write-Log -Level "WARN" -Message "Rilevato 'Documenti' come library. Converto automaticamente in 'Documents'."
    $FolderSiteRelativeUrl = 'Documents'
}

# =========================
# CONNESSIONE PNP (AGGIORNATA)
# =========================
if ($PSBoundParameters.ContainsKey('ClientId') -and $ClientId) {
    # Percorso Auth con App Registrata
    if ($PSBoundParameters.ContainsKey('TenantId') -and $TenantId) {
        Write-Log -Level "INFO" -Message "Connessione a PnP con App (ClientId=$ClientId, TenantId=$TenantId) - Interactive"
        Invoke-WithRetry { Connect-PnPOnline -Url $OneDriveSiteUrl -ClientId $ClientId -Tenant $TenantId -Interactive -ErrorAction Stop } | Out-Null
    } else {
        Write-Log -Level "INFO" -Message "Connessione a PnP con App (ClientId=$ClientId) - Interactive"
        Invoke-WithRetry { Connect-PnPOnline -Url $OneDriveSiteUrl -ClientId $ClientId -Interactive -ErrorAction Stop } | Out-Null
    }
} else {
    # Fallback: connessione interattiva predefinita
    Write-Log -Level "INFO" -Message "Connessione a PnP: Interactive (senza ClientId esplicito)"
    Invoke-WithRetry { Connect-PnPOnline -Url $OneDriveSiteUrl -Interactive -ErrorAction Stop } | Out-Null
}

# Test connettività
try {
    $web = Get-PnPWeb -ErrorAction Stop
    Write-Log -Level "INFO" -Message "Connesso a: $($web.Url)"
} catch {
    throw "Connessione fallita a '$OneDriveSiteUrl'. Dettagli: $($_.Exception.Message)"
}

# Validazione cartella
try {
    Write-Log -Level "DEBUG" -Message "Verifico esistenza cartella site-relative: '$FolderSiteRelativeUrl'"
    $null = Get-PnPFolder -Url $FolderSiteRelativeUrl -ErrorAction Stop
    Write-Log -Level "INFO" -Message "Cartella trovata: '$FolderSiteRelativeUrl'"
} catch {
    throw "La cartella site-relative '$FolderSiteRelativeUrl' non esiste o non è accessibile. Verifica la library ('Documents') e i permessi. Dettagli: $($_.Exception.Message)"
}
#endregion

#region Contatori & opzioni
$script:CountFolders          = 0
$script:CountFiles            = 0
$script:CountFilesNoVersions  = 0
$script:CountVersionsRemoved  = 0
$script:CountErrors           = 0

$timeStart = Get-Date
#endregion

#region Core processing
function Process-Folder {
    param(
        [Parameter(Mandatory=$true)][string] $FolderSiteRelativeUrlLocal
    )

    $script:CountFolders++
    Write-Log -Level "INFO" -Message "Elaboro cartella: $FolderSiteRelativeUrlLocal"

    # FILES
    $files = Invoke-WithRetry {
        Get-PnPFolderItem -FolderSiteRelativeUrl $FolderSiteRelativeUrlLocal -ItemType File -ErrorAction Stop
    }

    foreach ($file in $files) {
        try {
            $script:CountFiles++

            if ($ExcludeOneNote -and ($file.Name -match '\.one(toc2)?$')) {
                Write-Log -Level "DEBUG" -Message "Escludo file OneNote: $($file.Name)"
                continue
            }

            # URL server-relative del file
            $fileUrl = $null
            if ($null -ne $file.ServerRelativeUrl -and $file.ServerRelativeUrl -match '^/') {
                $fileUrl = $file.ServerRelativeUrl
            } else {
                $fileUrl = Ensure-ServerRelativeFileUrl -OneDriveSiteUrl $OneDriveSiteUrl -FolderSiteRelativeUrl $FolderSiteRelativeUrlLocal -FileName $file.Name
            }
            Write-Log -Level "DEBUG" -Message "File: $($file.Name) | URL (server-relative): $fileUrl"

            # Materializza File + Versions
            $listItem = Invoke-WithRetry { Get-PnPFile -Url $fileUrl -AsListItem -ErrorAction Stop }
            Get-PnPProperty -ClientObject $listItem -Property File -ErrorAction Stop
            Get-PnPProperty -ClientObject $listItem.File -Property Versions -ErrorAction Stop

            $versions = $listItem.File.Versions
            if (-not $versions -or $versions.Count -le 1) {
                $script:CountFilesNoVersions++
                Write-Log -Level "DEBUG" -Message "Nessuna versione da rimuovere (0 o 1) per '$($file.Name)'."
                continue
            }

            # Ordina, tieni la più recente
            $sorted  = $versions | Sort-Object -Property Created -Descending
            $latest  = $sorted | Select-Object -First 1
            $toPurge = $sorted | Select-Object -Skip 1

            if ($OlderThanDays -gt 0) {
                $threshold = (Get-Date).AddDays(-$OlderThanDays)
                $toPurge = $toPurge | Where-Object { $_.Created -lt $threshold }
            }

            if (-not $toPurge -or $toPurge.Count -eq 0) {
                Write-Log -Level "DEBUG" -Message "Nessuna versione da rimuovere per '$($file.Name)' (filtro giorni/ultima)."
                continue
            }

            Write-Log -Level "INFO" -Message ("File '{0}': mantengo ultima (Created={1}, Label={2}); rimuovo {3} versione/i." -f `
                $file.Name, $latest.Created, $latest.VersionLabel, $toPurge.Count)

            foreach ($ver in $toPurge) {
                $identity = $ver.ID
                $label    = $ver.VersionLabel
                $created  = $ver.Created

                if ($PSCmdlet.ShouldProcess("Versione $identity ($label) di $($file.Name)", "Remove-PnPFileVersion")) {
                    try {
                        Invoke-WithRetry {
                            Remove-PnPFileVersion -Url $fileUrl -Identity $identity -Force -ErrorAction Stop
                        }
                        $script:CountVersionsRemoved++
                        Write-Log -Level "DEBUG" -Message "Rimossa versione Id=$identity (Label=$label, Created=$created) per '$($file.Name)'."
                    } catch {
                        $script:CountErrors++
                        Write-Log -Level "WARN" -Message "Fallita rimozione versione Id=$identity ($label) per '$($file.Name)': $($_.Exception.Message)"
                    }
                }
            }
        } catch {
            $script:CountErrors++
            Write-Log -Level "ERROR" -Message "Errore elaborando file '$($file.Name)': $($_.Exception.Message)"
        }
    }

    # SUBFOLDERS
    if ($IncludeSubfolders) {
        $subFolders = Invoke-WithRetry {
            Get-PnPFolderItem -FolderSiteRelativeUrl $FolderSiteRelativeUrlLocal -ItemType Folder -ErrorAction Stop
        }

        foreach ($sub in $subFolders) {
            try {
                $siteRoot  = Get-SiteRootFromAbsoluteUrl -AbsoluteUrl $OneDriveSiteUrl
                $subSiteRel = if ($null -ne $sub.ServerRelativeUrl) {
                    Convert-ServerRelativeToSiteRelative -ServerRelativeUrl $sub.ServerRelativeUrl -SiteRootPath $siteRoot
                } else {
                    ($FolderSiteRelativeUrlLocal.TrimEnd('/') + '/' + $sub.Name)
                }

                Process-Folder -FolderSiteRelativeUrlLocal $subSiteRel
            } catch {
                $script:CountErrors++
                Write-Log -Level "ERROR" -Message "Errore elaborando sottocartella '$($sub.Name)': $($_.Exception.Message)"
            }
        }
    }
}
#endregion

#region Avvio
try {
    Process-Folder -FolderSiteRelativeUrlLocal $FolderSiteRelativeUrl
} finally {
    $elapsed = (Get-Date) - $timeStart
    Write-Host ""
    Write-Log -Level "INFO" -Message "=== RIEPILOGO ==="
    Write-Log -Level "INFO" -Message ("Cartelle elaborate : {0}" -f $script:CountFolders)
    Write-Log -Level "INFO" -Message ("File analizzati    : {0}" -f $script:CountFiles)
    Write-Log -Level "INFO" -Message ("File senza versioni: {0}" -f $script:CountFilesNoVersions)
    Write-Log -Level "INFO" -Message ("Versioni rimosse   : {0}" -f $script:CountVersionsRemoved)
    Write-Log -Level "INFO" -Message ("Errori             : {0}" -f $script:CountErrors)
    Write-Log -Level "INFO" -Message ("Durata             : {0:hh\:mm\:ss}" -f $elapsed)
}
#endregion