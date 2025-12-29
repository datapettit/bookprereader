$ErrorActionPreference = 'Stop'
cls
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$SettingsPath = Join-Path $ScriptRoot 'settings.json'
$MaxInputCharacters = 4096
$SupportedModels = @('gpt-4o-mini-tts', 'tts-1', 'tts-1-hd')
$SupportedVoices = @('alloy', 'echo', 'fable', 'onyx', 'nova', 'shimmer')
$DefaultApiKey = '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'
$EnableClearHost = $false
$IsWindows = $false
if ($env:OS -eq 'Windows_NT') {
    $IsWindows = $true
} elseif ($PSVersionTable.PSEdition -eq 'Desktop') {
    $IsWindows = $true
}

function Write-Info {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Cyan
}

function Write-Warn {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Yellow
}

function Write-ErrorMessage {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Red
}

function Write-Success {
    param([string]$Message)
    Write-Host $Message -ForegroundColor Green
}

function Clear-HostSafe {
    if ($EnableClearHost) {
        Clear-Host
    }
}

function Get-DefaultSettings {
    return [ordered]@{
        WorkspaceFolder = 'C:\git\'
        Model = 'gpt-4o-mini-tts'
        ApiKey = $DefaultApiKey
    }
}

function Load-Settings {
    if (Test-Path $SettingsPath) {
        $loaded = Get-Content $SettingsPath -Raw | ConvertFrom-Json
        $defaults = Get-DefaultSettings
        foreach ($key in $defaults.Keys) {
            if (-not $loaded.PSObject.Properties.Name.Contains($key)) {
                $loaded | Add-Member -MemberType NoteProperty -Name $key -Value $defaults[$key]
            }
        }
        return $loaded
    }
    return Get-DefaultSettings
}

function Save-Settings {
    param([object]$Settings)

    $Settings | ConvertTo-Json -Depth 4 | Set-Content -Path $SettingsPath -Encoding UTF8
}

function Ensure-WorkspaceFolder {
    param([string]$Path)

    if (-not (Test-Path $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Resolve-WorkspaceFolder {
    param([string]$Path)

    $normalized = Normalize-InputPath -Path $Path
    if (-not $normalized) {
        return $normalized
    }
    if (-not [System.IO.Path]::IsPathRooted($normalized)) {
        $normalized = Join-Path (Get-Location) $normalized
    }
    return [System.IO.Path]::GetFullPath($normalized)
}

function Get-ApiKey {
    param([object]$Settings)
    $candidate = if ($env:OPENAI_API_KEY) { $env:OPENAI_API_KEY } else { $Settings.ApiKey }
    if ($candidate) {
        $candidate = $candidate.Trim()
    }
    if ([string]::IsNullOrWhiteSpace($candidate)) {
        throw 'OpenAI API key is missing. Set OPENAI_API_KEY or update it in Settings.'
    }
    return $candidate
}

function Select-Voice {
    Clear-HostSafe
    Write-Info 'Select a voice model:'
    for ($i = 0; $i -lt $SupportedVoices.Count; $i++) {
        Write-Host ("  {0}) {1}" -f ($i + 1), $SupportedVoices[$i])
    }
    while ($true) {
        $inputValue = (Read-Host 'Enter number or name').Trim()
        if ([string]::IsNullOrWhiteSpace($inputValue)) {
            continue
        }
        if ($inputValue -match '^\d+$') {
            $index = [int]$inputValue - 1
            if ($index -ge 0 -and $index -lt $SupportedVoices.Count) {
                Write-Success ("Selected voice: {0}" -f $SupportedVoices[$index])
                return $SupportedVoices[$index]
            }
        }
        $match = $SupportedVoices | Where-Object { $_ -eq $inputValue.ToLowerInvariant() }
        if ($match) {
            Write-Success ("Selected voice: {0}" -f $match)
            return $match
        }
        Write-Warn 'Invalid selection. Try again.'
    }
}

function Select-InputMethod {
    Clear-HostSafe
    Write-Info 'Input method:'
    Write-Host '  1) File upload'
    Write-Host '  2) Paste text'
    Write-Host '  3) Folder'
    while ($true) {
        $inputValue = (Read-Host 'Choose input type').Trim().ToLowerInvariant()
        switch ($inputValue) {
            '1' { return 'file' }
            '2' { return 'text' }
            '3' { return 'folder' }
            'file' { return 'file' }
            'text' { return 'text' }
            'folder' { return 'folder' }
        }
        Write-Warn 'Invalid selection. Enter 1, 2, 3, file, text, or folder.'
    }
}

function Get-WorkspaceFileCandidates {
    param([string]$WorkspaceFolder)

    $extensions = @('*.txt', '*.text', '*.rtf', '*.doc', '*.docx')
    return Get-ChildItem -Path $WorkspaceFolder -File -Include $extensions | Sort-Object Name
}

function Test-SupportedInputFile {
    param([string]$Path)

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    return $extension -in @('.txt', '.text', '.rtf', '.doc', '.docx')
}

function Resolve-InputFileSelection {
    param(
        [string]$Selection,
        [string]$WorkspaceFolder,
        [object[]]$Candidates
    )

    $inputValue = Normalize-InputPath -Path $Selection
    if ([string]::IsNullOrWhiteSpace($inputValue)) {
        return $null
    }

    if ($inputValue -match '^\d+$') {
        $index = [int]$inputValue - 1
        if ($index -ge 0 -and $index -lt $Candidates.Count) {
            return $Candidates[$index].FullName
        }
        return $null
    }

    $candidatePath = $inputValue
    if (-not [System.IO.Path]::IsPathRooted($candidatePath)) {
        $candidatePath = Join-Path $WorkspaceFolder $candidatePath
    }

    if (Test-Path -LiteralPath $candidatePath) {
        $item = Get-Item -LiteralPath $candidatePath
        if (-not $item.PSIsContainer -and (Test-SupportedInputFile -Path $item.FullName)) {
            return $item.FullName
        }
    }

    $fileName = [System.IO.Path]::GetFileName($inputValue)
    if ($fileName) {
        $matching = $Candidates | Where-Object { $_.Name -ieq $fileName }
        if ($matching.Count -eq 1) {
            return $matching[0].FullName
        }
    }

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($inputValue)
    if ($baseName) {
        $matching = $Candidates | Where-Object { $_.BaseName -ieq $baseName }
        if ($matching.Count -eq 1) {
            return $matching[0].FullName
        }
    }

    return $null
}

function Select-InputFile {
    param([object]$Settings)

    $workspaceFolder = Resolve-WorkspaceFolder -Path $Settings.WorkspaceFolder
    Ensure-WorkspaceFolder -Path $workspaceFolder

    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    } catch {
        Write-Warn 'File dialog not available. Falling back to manual path entry.'
    }

    while ($true) {
        $candidates = Get-WorkspaceFileCandidates -WorkspaceFolder $workspaceFolder
        Write-Info ("Working directory: {0}" -f $workspaceFolder)
        if ($candidates.Count -gt 0) {
            Write-Info 'Files available in working directory:'
            for ($i = 0; $i -lt $candidates.Count; $i++) {
                Write-Host ("  {0}) {1}" -f ($i + 1), $candidates[$i].Name)
            }
        } else {
            Write-Warn 'No supported files found in the working directory.'
        }

        if ([System.Type]::GetType('System.Windows.Forms.OpenFileDialog')) {
            Write-Host '  d) Open file dialog'
        }

        $selection = Read-Host 'Enter file number, file name, or full path'
        if ([string]::IsNullOrWhiteSpace($selection)) {
            continue
        }

        if ($selection.Trim().ToLowerInvariant() -eq 'd' -and [System.Type]::GetType('System.Windows.Forms.OpenFileDialog')) {
            $dialog = New-Object System.Windows.Forms.OpenFileDialog
            $dialog.Title = 'Select a text or Word document'
            $dialog.Filter = 'Text and Word Documents|*.txt;*.text;*.rtf;*.doc;*.docx|All Files|*.*'
            $dialog.Multiselect = $false
            $dialog.InitialDirectory = $workspaceFolder

            $result = $dialog.ShowDialog()
            if ($result -eq [System.Windows.Forms.DialogResult]::OK -and $dialog.FileName) {
                if (Test-SupportedInputFile -Path $dialog.FileName) {
                    return $dialog.FileName
                }
                Write-Warn 'Selected file type is not supported.'
            }
            continue
        }

        $resolved = Resolve-InputFileSelection -Selection $selection -WorkspaceFolder $workspaceFolder -Candidates $candidates
        if ($resolved) {
            return $resolved
        }

        Write-Warn 'Selection did not match a readable file. Try again.'
    }
}

function Select-InputFolder {
    param([object]$Settings)

    $workspaceFolder = Resolve-WorkspaceFolder -Path $Settings.WorkspaceFolder
    Ensure-WorkspaceFolder -Path $workspaceFolder
    $manualPath = Read-Host 'Enter the full path to the folder'
    $manualPath = Normalize-InputPath -Path $manualPath
    if ($manualPath -and -not [System.IO.Path]::IsPathRooted($manualPath)) {
        $manualPath = Join-Path $workspaceFolder $manualPath
    }
    if (-not (Test-Path -LiteralPath $manualPath)) {
        throw "Folder not found: $manualPath"
    }
    $item = Get-Item -LiteralPath $manualPath
    if (-not $item.PSIsContainer) {
        throw "Path is not a folder: $manualPath"
    }
    return $item.FullName
}

function Normalize-InputPath {
    param([string]$Path)

    if ($null -eq $Path) {
        return $Path
    }
    $trimmed = $Path.Trim()
    if ($trimmed.Length -ge 2) {
        if (($trimmed.StartsWith('"') -and $trimmed.EndsWith('"')) -or ($trimmed.StartsWith("'") -and $trimmed.EndsWith("'"))) {
            $trimmed = $trimmed.Substring(1, $trimmed.Length - 2)
        }
    }
    return $trimmed
}

function Get-TextFromDocx {
    param([string]$Path)

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        $entry = $zip.GetEntry('word/document.xml')
        if (-not $entry) {
            throw 'Invalid docx file: missing document.xml.'
        }
        $reader = New-Object System.IO.StreamReader($entry.Open())
        try {
            $xml = $reader.ReadToEnd()
        } finally {
            $reader.Close()
        }
    } finally {
        $zip.Dispose()
    }

    $xml = $xml -replace '</w:p>', "`n`n"
    $text = $xml -replace '<[^>]+>', ''
    $text = [System.Net.WebUtility]::HtmlDecode($text)
    return $text
}

function Get-TextFromRtf {
    param([string]$Path)

    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    $rtb = New-Object System.Windows.Forms.RichTextBox
    $rtb.LoadFile($Path)
    return $rtb.Text
}

function Get-TextFromWordDoc {
    param([string]$Path)

    if (-not $IsWindows) {
        throw 'Reading .doc files requires Microsoft Word on Windows.'
    }
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $document = $null
    try {
        $document = $word.Documents.Open($Path, $false, $true)
        return $document.Content.Text
    } finally {
        if ($document -and $document.Saved -ne $true) {
            $document.Close($false)
        }
        $word.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    }
}

function Get-TextFromFile {
    param([string]$Path)

    $extension = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    switch ($extension) {
        '.txt' { return Get-Content -Path $Path -Raw }
        '.text' { return Get-Content -Path $Path -Raw }
        '.rtf' { return Get-TextFromRtf -Path $Path }
        '.docx' { return Get-TextFromDocx -Path $Path }
        '.doc' { return Get-TextFromWordDoc -Path $Path }
        default { throw "Unsupported file extension: $extension" }
    }
}

function Get-InputFilesFromFolder {
    param([string]$FolderPath)

    $extensions = @('*.txt', '*.text', '*.rtf', '*.doc', '*.docx')
    $files = Get-ChildItem -Path $FolderPath -File -Recurse -Include $extensions | Sort-Object FullName
    return $files
}

function Get-ChapterNameFromFile {
    param([string]$FilePath)

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    $normalized = $baseName -replace '[^a-zA-Z0-9]', ''
    return $normalized
}

function Get-ChapterNameFromInput {
    param([string]$InputName)

    $normalized = $InputName -replace '[^a-zA-Z0-9]', ''
    return $normalized
}

function Read-PastedText {
    Write-Info 'Paste text below. Enter a single line with END to finish.'
    $lines = New-Object System.Collections.Generic.List[string]
    while ($true) {
        $line = Read-Host
        if ($line -eq 'END') {
            break
        }
        $lines.Add($line)
    }
    return ($lines -join "`n")
}

function Split-TextIntoChunks {
    param(
        [string]$Text,
        [int]$Limit
    )

    $chunks = New-Object System.Collections.Generic.List[string]
    $paragraphs = $Text -split "(?:\r?\n){2,}"
    $current = ''

    foreach ($paragraph in $paragraphs) {
        $clean = $paragraph.Trim()
        if (-not $clean) {
            continue
        }
        if ($current -and (($current.Length + $clean.Length + 2) -gt $Limit)) {
            $chunks.Add($current.Trim())
            $current = ''
        }

        if ($clean.Length -gt $Limit) {
            Write-Warn ("Paragraph exceeds limit ({0} chars). Splitting by sentences." -f $clean.Length)
            $sentences = $clean -split '(?<=[.!?])\s+'
            $buffer = ''
            foreach ($sentence in $sentences) {
                if ([string]::IsNullOrWhiteSpace($sentence)) {
                    continue
                }
                if ($buffer -and (($buffer.Length + $sentence.Length + 1) -gt $Limit)) {
                    $chunks.Add($buffer.Trim())
                    $buffer = ''
                }
                if ($sentence.Length -gt $Limit) {
                    $offset = 0
                    while ($offset -lt $sentence.Length) {
                        $sliceLength = [Math]::Min($Limit, $sentence.Length - $offset)
                        $chunks.Add($sentence.Substring($offset, $sliceLength))
                        $offset += $sliceLength
                    }
                } else {
                    if ($buffer) {
                        $buffer += " $sentence"
                    } else {
                        $buffer = $sentence
                    }
                }
            }
            if ($buffer) {
                $chunks.Add($buffer.Trim())
            }
            continue
        }

        if ($current) {
            $current += "`n`n$clean"
        } else {
            $current = $clean
        }
    }

    if ($current) {
        $chunks.Add($current.Trim())
    }

    return $chunks
}

function Normalize-TextForJson {
    param([string]$Text)

    if ($null -eq $Text) {
        return $Text
    }

    $normalized = $Text
    $normalized = $normalized.Replace([char]0x201C, '"')
    $normalized = $normalized.Replace([char]0x201D, '"')
    $normalized = $normalized.Replace([char]0x2018, "'")
    $normalized = $normalized.Replace([char]0x2019, "'")

    return $normalized
}

function Invoke-OpenAITts {
    param(
        [string]$Text,
        [string]$Voice,
        [string]$Model,
        [string]$ApiKey,
        [string]$OutputPath
    )

    $uri = 'https://api.openai.com/v1/audio/speech'
    $headers = @{
        "Authorization" = "Bearer $ApiKey"
        "Content-Type"  = "application/json"
    }
    $safeText = Normalize-TextForJson -Text $Text
    $bodyObject = @{
        model = $Model
        input = $safeText
        voice = $Voice
    }
    $body = $bodyObject | ConvertTo-Json -Depth 4

    if (Test-Path $OutputPath) {
        Remove-Item $OutputPath -Force
    }

    for ($attempt = 0; $attempt -le 2; $attempt++) {
        try {
            Write-Info 'OpenAI TTS request details (full):'
            Write-Info ("  URL: {0}" -f $uri)
            Write-Info ("  Headers (raw): {0}" -f ($headers | ConvertTo-Json -Depth 6))
            Write-Info ("  Headers (redacted): {0}" -f ((Get-RedactedHeadersForLog -Headers $headers) | ConvertTo-Json -Depth 6))
            Write-Info ("  Body: {0}" -f $body)
            Write-Info ("  Output path: {0}" -f $OutputPath)

            $response = Invoke-WebRequest -Method Post -Uri $uri -Headers $headers -Body $body -ContentType 'application/json' -OutFile $OutputPath -PassThru -ErrorAction Stop
            Write-Info 'OpenAI TTS response details (success):'
            Write-Info ("  Status: {0} {1}" -f $response.StatusCode, $response.StatusDescription)
            Write-Info ("  Headers: {0}" -f ($response.Headers | ConvertTo-Json -Depth 6))
            if ($response.Headers['x-request-id']) {
                Write-Info ("  OpenAI request id: {0}" -f $response.Headers['x-request-id'])
            }

            return
        } catch {
            Write-Error $_.Exception.ToString()
            Write-Error $_.ToString()
            $errorDetails = New-Object System.Collections.Generic.List[string]
            $errorDetails.Add('OpenAI TTS request failed.')
            $errorDetails.Add(("Request URL: {0}" -f $uri))
            $redactedHeaders = Get-RedactedHeadersForLog -Headers $headers
            $unredactedHeaders = Get-UnRedactedHeadersForLog -Headers $headers
            $errorDetails.Add(("Request raw headers: {0}" -f ($unredactedHeaders | ConvertTo-Json -Depth 6)))
            $errorDetails.Add(("Request redacted headers: {0}" -f ($redactedHeaders | ConvertTo-Json -Depth 6)))
            $errorDetails.Add(("Request body: {0}" -f $body))
            $errorDetails.Add(("Request output path: {0}" -f $OutputPath))

            
            $response = $_.Exception.Response
            if ($response) {
                $errorDetails.Add(("Response status: {0} {1}" -f [int]$response.StatusCode, $response.StatusDescription))
                $errorDetails.Add(("Response headers (raw): {0}" -f $response.Headers.ToString()))
                $errorDetails.Add(("Response headers (json): {0}" -f ($response.Headers | ConvertTo-Json -Depth 6)))
                if ($response.Headers['x-request-id']) {
                    $errorDetails.Add(("OpenAI request id: {0}" -f $response.Headers['x-request-id']))
                }
                if ($response.GetResponseStream()) {
                    $reader = New-Object System.IO.StreamReader($response.GetResponseStream())
                    try {
                        $details = $reader.ReadToEnd()
                        if ($details) {
                            $errorDetails.Add(("Response body: {0}" -f $details))
                        } else {
                            $errorDetails.Add('Response body: (empty)')
                        }
                    } finally {
                        $reader.Close()
                    }
                } else {
                    $errorDetails.Add('Response body: (no response stream)')
                }
            } else {
                $errorDetails.Add('Response: (none)')
            }
            $errorDetails.Add(("Exception type: {0}" -f $_.Exception.GetType().FullName))
            $errorDetails.Add(("Exception message: {0}" -f $_.Exception.Message))
            foreach ($line in $errorDetails) {
                Write-ErrorMessage $line
            }
            if ($attempt -lt 2) {
                Write-Warn 'OpenAI TTS request failed. Waiting 30 seconds before retry...'
                Start-Sleep -Seconds 30
                if (Test-Path $OutputPath) {
                    Remove-Item $OutputPath -Force
                }
            } else {
                throw 'OpenAI TTS request failed after 2 retries.'
            }
        }
    }
}

function Ensure-Ffmpeg {
    $existing = Get-Command ffmpeg -ErrorAction SilentlyContinue
    if ($existing) {
        return $existing.Source
    }

    $localRoot = Join-Path $ScriptRoot 'tools/ffmpeg'
    $localBinary = Join-Path $localRoot 'bin/ffmpeg.exe'
    if (Test-Path $localBinary) {
        return $localBinary
    }

    if (-not $IsWindows) {
        throw 'FFmpeg is required to merge MP3 files. Install ffmpeg and ensure it is in PATH.'
    }

    Write-Warn 'FFmpeg not found. Downloading a free build from gyan.dev...'
    $zipPath = Join-Path $ScriptRoot 'ffmpeg.zip'
    try {
        Invoke-WebRequest -Uri 'https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-essentials.zip' -OutFile $zipPath -ErrorAction Stop
        Expand-Archive -Path $zipPath -DestinationPath $ScriptRoot -Force
    } finally {
        if (Test-Path $zipPath) {
            Remove-Item $zipPath -Force
        }
    }

    $extracted = Get-ChildItem -Path $ScriptRoot -Directory | Where-Object { $_.Name -like 'ffmpeg-*' } | Select-Object -First 1
    if (-not $extracted) {
        throw 'Failed to extract FFmpeg.'
    }
    if (Test-Path $localRoot) {
        Remove-Item $localRoot -Recurse -Force
    }
    Move-Item -Path $extracted.FullName -Destination $localRoot

    if (-not (Test-Path $localBinary)) {
        throw 'FFmpeg binary not found after download.'
    }

    return $localBinary
}

function Get-FfprobePath {
    param([string]$FfmpegPath)

    $probeCommand = Get-Command ffprobe -ErrorAction SilentlyContinue
    if ($probeCommand) {
        return $probeCommand.Source
    }

    if ($FfmpegPath) {
        $ffmpegDir = Split-Path $FfmpegPath
        $probeCandidate = Join-Path $ffmpegDir 'ffprobe.exe'
        if (Test-Path $probeCandidate) {
            return $probeCandidate
        }
    }

    return $null
}

function Get-Mp3Duration {
    param(
        [string]$Path,
        [string]$FfprobePath
    )

    if (-not $FfprobePath) {
        return $null
    }

    try {
        $durationText = & $FfprobePath -v error -show_entries format=duration -of default=nw=1:nk=1 $Path 2>$null
        if ($durationText -and $durationText -match '^\d+(\.\d+)?$') {
            return [double]$durationText
        }
    } catch {
        return $null
    }

    return $null
}

function Write-Mp3Report {
    param(
        [string]$Path,
        [bool]$Success,
        [string]$FailureMessage,
        [string]$FfprobePath
    )

    if (-not $Success) {
        Write-ErrorMessage ("Final MP3 creation failed: {0}" -f $FailureMessage)
        Write-ErrorMessage ("Expected output path: {0}" -f $Path)
        return
    }

    if (-not (Test-Path $Path)) {
        Write-ErrorMessage ("Final MP3 was not found at: {0}" -f $Path)
        return
    }

    $item = Get-Item -Path $Path
    $duration = Get-Mp3Duration -Path $Path -FfprobePath $FfprobePath
    Write-Info 'Final MP3 details:'
    Write-Info ("  Name: {0}" -f $item.Name)
    Write-Info ("  Location: {0}" -f $item.FullName)
    Write-Info ("  Size: {0} bytes ({1:N2} MB)" -f $item.Length, ($item.Length / 1MB))
    if ($null -ne $duration) {
        Write-Info ("  Length: {0:N2} seconds" -f $duration)
    } else {
        Write-Warn '  Length: unavailable (ffprobe not found).'
    }
}

function Get-Mp3SortKey {
    param([string]$Path)

    $fileName = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    if ($fileName -match '(\d+)$') {
        return [int]$Matches[1]
    }
    return [int]::MaxValue
}

function Test-Mp3File {
    param(
        [string]$Path,
        [string]$FfmpegPath
    )

    if (-not (Test-Path $Path)) {
        Write-Warn ("MP3 missing: {0}" -f $Path)
        return $false
    }

    $length = (Get-Item $Path).Length
    Write-Info ("MP3 size check: {0} bytes ({1})" -f $length, $Path)
    if ($length -le 0) {
        Write-Warn ("MP3 has zero size and will be skipped: {0}" -f $Path)
        return $false
    }

    if (-not $FfmpegPath) {
        Write-Warn ("Skipping MP3 validation (ffmpeg missing): {0}" -f $Path)
        return $true
    }

    Write-Info ("Validating MP3 with ffmpeg: {0}" -f $Path)
    $process = Start-Process -FilePath $FfmpegPath -ArgumentList @('-v', 'error', '-i', $Path, '-f', 'null', '-') -NoNewWindow -Wait -PassThru
    if ($process.ExitCode -ne 0) {
        Write-Warn ("MP3 validation failed (ffmpeg exit {0}): {1}" -f $process.ExitCode, $Path)
        return $false
    }

    Write-Success ("MP3 validation passed: {0}" -f $Path)
    return $true
}

function Invoke-FfmpegConcat {
    param(
        [string]$FfmpegPath,
        [string]$ListPath,
        [string]$OutputPath,
        [bool]$Reencode
    )

    $arguments = @('-y', '-f', 'concat', '-safe', '0', '-i', $ListPath)
    if ($Reencode) {
        $arguments += @('-c:a', 'libmp3lame', '-q:a', '2')
    } else {
        $arguments += @('-c', 'copy')
    }
    $arguments += $OutputPath

    return Start-Process -FilePath $FfmpegPath -ArgumentList $arguments -NoNewWindow -Wait -PassThru
}

function Merge-Mp3Files {
    param(
        [string[]]$Files,
        [string]$OutputPath
    )

    $ffmpeg = Ensure-Ffmpeg
    Write-Info ("FFmpeg path: {0}" -f $ffmpeg)
    $outputFolder = Split-Path $OutputPath -Parent
    if ($outputFolder) {
        Ensure-WorkspaceFolder -Path $outputFolder
    }

    Write-Info 'Ordering MP3 chunks by numeric suffix...'
    $orderedFiles = $Files | Sort-Object { Get-Mp3SortKey -Path $_ }
    foreach ($file in $orderedFiles) {
        Write-Info ("  Ordered chunk: {0}" -f $file)
    }

    Write-Info 'Validating MP3 chunks before merge...'
    $validFiles = New-Object System.Collections.Generic.List[string]
    foreach ($file in $orderedFiles) {
        if (Test-Mp3File -Path $file -FfmpegPath $ffmpeg) {
            $validFiles.Add($file)
        } else {
            Write-Warn ("Skipping invalid MP3 chunk: {0}" -f $file)
        }
    }

    if ($validFiles.Count -eq 0) {
        throw 'No valid MP3 chunks remain after validation.'
    }

    $listPath = Join-Path $outputFolder ('ffmpeg-list-' + [Guid]::NewGuid().ToString('N') + '.txt')
    $content = $validFiles | ForEach-Object { "file '$($_.Replace("'", "''"))'" }
    Set-Content -Path $listPath -Value $content -Encoding UTF8

    if (Test-Path $OutputPath) {
        Remove-Item $OutputPath -Force
    }

    try {
        Write-Info ("Merging {0} chunks into {1}" -f $validFiles.Count, $OutputPath)
        $process = Invoke-FfmpegConcat -FfmpegPath $ffmpeg -ListPath $listPath -OutputPath $OutputPath -Reencode $false
        if ($process.ExitCode -ne 0) {
            Write-Warn ("FFmpeg stream copy merge failed (exit {0}). Retrying with re-encode..." -f $process.ExitCode)
            if (Test-Path $OutputPath) {
                Remove-Item $OutputPath -Force
            }
            $process = Invoke-FfmpegConcat -FfmpegPath $ffmpeg -ListPath $listPath -OutputPath $OutputPath -Reencode $true
            if ($process.ExitCode -ne 0) {
                throw ("FFmpeg merge failed with exit code {0}." -f $process.ExitCode)
            }
        }
        Write-Success ("Merge complete: {0}" -f $OutputPath)
    } finally {
        if (Test-Path $listPath) {
            Remove-Item $listPath -Force
        }
    }

    return $ffmpeg
}

function Choose-Model {
    param([object]$Settings)

    Clear-HostSafe
    Write-Info 'Select a TTS model:'
    for ($i = 0; $i -lt $SupportedModels.Count; $i++) {
        Write-Host ("  {0}) {1}" -f ($i + 1), $SupportedModels[$i])
    }
    while ($true) {
        $inputValue = (Read-Host 'Enter number or model name').Trim().ToLowerInvariant()
        if ([string]::IsNullOrWhiteSpace($inputValue)) {
            continue
        }
        if ($inputValue -match '^\d+$') {
            $index = [int]$inputValue - 1
            if ($index -ge 0 -and $index -lt $SupportedModels.Count) {
                $Settings.Model = $SupportedModels[$index]
                Write-Success ("Selected model: {0}" -f $Settings.Model)
                return
            }
        }
        $match = $SupportedModels | Where-Object { $_ -eq $inputValue }
        if ($match) {
            $Settings.Model = $match
            Write-Success ("Selected model: {0}" -f $Settings.Model)
            return
        }
        Write-Warn 'Invalid selection. Try again.'
    }
}

function Get-ApiKeyPreview {
    param([string]$ApiKey)

    if ([string]::IsNullOrWhiteSpace($ApiKey)) {
        return '(not set)'
    }
    $trimmed = $ApiKey.Trim()
    if ($trimmed.Length -le 8) {
        return ('*' * $trimmed.Length)
    }
    $prefix = $trimmed.Substring(0, 4)
    $suffix = $trimmed.Substring($trimmed.Length - 4, 4)
    return ("{0}...{1}" -f $prefix, $suffix)
}

function Get-RedactedHeadersForLog {
    param([hashtable]$Headers)

    $copy = @{}
    foreach ($key in $Headers.Keys) {
        if ($key -eq 'Authorization') {
            $copy[$key] = 'Bearer [REDACTED]'
        } else {
            $copy[$key] = $Headers[$key]
        }
    }
    return $copy
}

function Get-UnRedactedHeadersForLog {
    param([hashtable]$Headers)

    $copy = @{}
    foreach ($key in $Headers.Keys) {
            $copy[$key] = $Headers[$key]
    }
    return $copy
}

function Convert-ToTitleCase {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $Text
    }
    $culture = [System.Globalization.CultureInfo]::CurrentCulture
    return $culture.TextInfo.ToTitleCase($Text.ToLowerInvariant())
}

function Get-SceneSeparatorInfo {
    param([string]$Line)

    if ($null -eq $Line) {
        return $null
    }

    $trimmed = $Line.Trim()
    if ($trimmed -match '^(Scene|Chapter)$') {
        return @{
            Type = $Matches[1]
            Title = $null
            RawLine = $trimmed
        }
    }

    if ($trimmed -match '^(Scene|Chapter)\s*-\s*(.+)$') {
        return @{
            Type = $Matches[1]
            Title = $Matches[2].Trim()
            RawLine = $trimmed
        }
    }

    return $null
}

function Sanitize-FileName {
    param([string]$Name)

    if ($null -eq $Name) {
        return $Name
    }
    $sanitized = $Name -replace '[<>:"/\\|?*]', ''
    $sanitized = $sanitized -replace '\s+', ' '
    return $sanitized.Trim()
}

function Invoke-OpenAISceneTitle {
    param(
        [string]$Text,
        [string]$ApiKey,
        [string]$Type
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $null
    }

    $uri = 'https://api.openai.com/v1/chat/completions'
    $headers = @{
        "Authorization" = "Bearer $ApiKey"
        "Content-Type"  = "application/json"
    }
    $safeText = Normalize-TextForJson -Text $Text
    $prompt = "Create a short, punchy $Type title (3-8 words) based on the content. Respond with the title only."
    $bodyObject = @{
        model = 'gpt-4o-mini'
        messages = @(
            @{ role = 'system'; content = 'You create concise scene and chapter titles.' },
            @{ role = 'user'; content = "$prompt`n`nContent:`n$safeText" }
        )
        temperature = 0.6
        max_tokens = 20
    }
    $body = $bodyObject | ConvertTo-Json -Depth 6

    try {
        $response = Invoke-RestMethod -Method Post -Uri $uri -Headers $headers -Body $body -ContentType 'application/json' -ErrorAction Stop
        $title = $response.choices[0].message.content
        if ($title) {
            return $title.Trim().Trim('"')
        }
    } catch {
        Write-Warn ("Title generation failed: {0}" -f $_.Exception.Message)
        return $null
    }

    return $null
}

function Split-BookIntoScenes {
    param(
        [string]$FilePath,
        [object]$Settings
    )

    $Settings.WorkspaceFolder = Resolve-WorkspaceFolder -Path $Settings.WorkspaceFolder
    Ensure-WorkspaceFolder -Path $Settings.WorkspaceFolder
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FilePath)
    $outputFolder = Join-Path $Settings.WorkspaceFolder ("{0}-split" -f $baseName)
    Ensure-WorkspaceFolder -Path $outputFolder

    Write-Info ("Reading input file: {0}" -f $FilePath)
    $text = Get-TextFromFile -Path $FilePath
    if (-not $text) {
        throw 'No text found in the file.'
    }

    $lines = $text -split '\r?\n'
    $chunks = New-Object System.Collections.Generic.List[object]
    $currentLines = New-Object System.Collections.Generic.List[string]
    $currentInfo = $null

    foreach ($line in $lines) {
        $separatorInfo = Get-SceneSeparatorInfo -Line $line
        if ($separatorInfo) {
            if ($currentLines.Count -gt 0) {
                $chunks.Add([pscustomobject]@{
                        Info = $currentInfo
                        Content = ($currentLines -join "`n")
                    })
                $currentLines.Clear()
            }
            $currentInfo = $separatorInfo
            $currentLines.Add($line)
            continue
        }
        $currentLines.Add($line)
    }

    if ($currentLines.Count -gt 0) {
        $chunks.Add([pscustomobject]@{
                Info = $currentInfo
                Content = ($currentLines -join "`n")
            })
    }

    if ($chunks.Count -eq 0) {
        throw 'No content detected to split.'
    }

    $apiKey = $null
    try {
        $apiKey = Get-ApiKey -Settings $Settings
    } catch {
        $apiKey = $null
    }

    $index = 1
    foreach ($chunk in $chunks) {
        $info = $chunk.Info
        $type = if ($info -and $info.Type) { $info.Type } else { 'Scene' }
        $title = $null
        if ($info -and $info.Title) {
            $title = Convert-ToTitleCase -Text $info.Title
        } else {
            if ($apiKey) {
                $generated = Invoke-OpenAISceneTitle -Text $chunk.Content -ApiKey $apiKey -Type $type
                if ($generated) {
                    $title = Convert-ToTitleCase -Text $generated
                }
            }
        }

        $displayName = if ($title) { "$type - $title" } else { $type }
        $fileName = Sanitize-FileName -Name ("{0} - {1}.txt" -f $index, $displayName)
        $outputPath = Join-Path $outputFolder $fileName

        Set-Content -Path $outputPath -Value $chunk.Content -Encoding UTF8
        Write-Success ("Saved: {0}" -f $outputPath)
        $index++
    }

    Write-Success ("Split complete. Output folder: {0}" -f $outputFolder)
}

function Invoke-TtsForText {
    param(
        [string]$ChapterName,
        [string]$InputText,
        [string]$Voice,
        [object]$Settings
    )

    $Settings.WorkspaceFolder = Resolve-WorkspaceFolder -Path $Settings.WorkspaceFolder
    Ensure-WorkspaceFolder -Path $Settings.WorkspaceFolder

    $cleanName = Get-ChapterNameFromInput -InputName $ChapterName
    if (-not $cleanName) {
        throw 'Chapter name cannot be blank after cleaning.'
    }

    $text = $InputText.Trim()
    if (-not $text) {
        throw 'No text provided.'
    }

    Write-Info ("Chapter name set to: {0}" -f $cleanName)
    Write-Info 'Splitting text into chunks...'
    $chunks = Split-TextIntoChunks -Text $text -Limit $MaxInputCharacters
    if ($chunks.Count -eq 0) {
        throw 'No usable text detected after splitting.'
    }
    Write-Success ("Prepared {0} chunks for TTS." -f $chunks.Count)
    $apiKey = Get-ApiKey -Settings $Settings
    Write-Info 'API key detected for TTS requests.'
    Write-Warn $apiKey

    $outputFiles = New-Object System.Collections.Generic.List[string]
    $ffmpegPath = $null
    $mergeSucceeded = $false
    $mergeError = $null

    for ($i = 0; $i -lt $chunks.Count; $i++) {
        $index = $i + 1
        $chunkPath = Join-Path $Settings.WorkspaceFolder ("{0}_{1}.mp3" -f $cleanName, $index)
        Write-Info ("Creating audio chunk {0}/{1} -> {2}" -f $index, $chunks.Count, $chunkPath)
        Invoke-OpenAITts -Text $chunks[$i] -Voice $voice -Model $Settings.Model -ApiKey $apiKey -OutputPath $chunkPath
        Write-Success ("Chunk created: {0}" -f $chunkPath)
        $outputFiles.Add($chunkPath)
    }

    $finalPath = Join-Path $Settings.WorkspaceFolder ("{0}.mp3" -f $cleanName)
    if (Test-Path $finalPath) {
        Write-Info ("Removing existing output file before merge: {0}" -f $finalPath)
        Remove-Item $finalPath -Force
    }
    if ($outputFiles.Count -gt 1) {
        Write-Info 'Merging chunks into final MP3...'
        try {
            $ffmpegPath = Merge-Mp3Files -Files $outputFiles -OutputPath $finalPath
            $mergeSucceeded = $true
            foreach ($file in $outputFiles) {
                Write-Info ("Removing temporary chunk: {0}" -f $file)
                Remove-Item $file -Force
            }
        } catch {
            $mergeSucceeded = $false
            $mergeError = $_.Exception.Message
            Write-ErrorMessage ("Merge failed: {0}" -f $mergeError)
        }
    } else {
        if (Test-Path $finalPath) {
            Write-Info ("Removing existing output file: {0}" -f $finalPath)
            Remove-Item $finalPath -Force
        }
        Write-Info ("Moving single chunk to final output: {0}" -f $finalPath)
        Move-Item -Path $outputFiles[0] -Destination $finalPath
        $mergeSucceeded = $true
    }
    $ffprobePath = Get-FfprobePath -FfmpegPath $ffmpegPath
    Write-Mp3Report -Path $finalPath -Success $mergeSucceeded -FailureMessage $mergeError -FfprobePath $ffprobePath
}

$settings = Load-Settings
$settings.WorkspaceFolder = Resolve-WorkspaceFolder -Path $settings.WorkspaceFolder
Ensure-WorkspaceFolder -Path $settings.WorkspaceFolder
Save-Settings -Settings $settings
Write-Info ("Working directory initialized: {0}" -f $settings.WorkspaceFolder)

while ($true) {
    Clear-HostSafe
    Write-Info 'Main Menu'
    Write-Info ("Working directory: {0}" -f $settings.WorkspaceFolder)
    Write-Host '1) Create audio'
    Write-Host '2) Split book into chapters/scenes'
    Write-Host '9) Settings'
    Write-Host '0) Exit'
    $selection = (Read-Host 'Choose an option').Trim()

    switch ($selection) {
        '1' {
            $voice = Select-Voice
            $inputMethod = Select-InputMethod

            try {
                if ($inputMethod -eq 'file') {
                    Write-Info 'Input method selected: file upload'
                    $filePath = Select-InputFile -Settings $settings
                    $chapterName = Get-ChapterNameFromFile -FilePath $filePath
                    Write-Info ("Reading input file: {0}" -f $filePath)
                    $inputText = Get-TextFromFile -Path $filePath
                    Invoke-TtsForText -ChapterName $chapterName -InputText $inputText -Voice $voice -Settings $settings
                } elseif ($inputMethod -eq 'folder') {
                    Write-Info 'Input method selected: folder'
                    $folderPath = Select-InputFolder -Settings $settings
                    Write-Info ("Scanning folder: {0}" -f $folderPath)
                    $files = Get-InputFilesFromFolder -FolderPath $folderPath
                    if (-not $files -or $files.Count -eq 0) {
                        Write-Warn 'No supported files found in the selected folder.'
                        break
                    }
                    foreach ($file in $files) {
                        $chapterName = Get-ChapterNameFromFile -FilePath $file.FullName
                        if (-not $chapterName) {
                            Write-Warn ("Skipping file with empty chapter name after cleaning: {0}" -f $file.FullName)
                            continue
                        }
                        Write-Info ("Reading input file: {0}" -f $file.FullName)
                        $inputText = Get-TextFromFile -Path $file.FullName
                        Invoke-TtsForText -ChapterName $chapterName -InputText $inputText -Voice $voice -Settings $settings
                    }
                } else {
                    Write-Info 'Input method selected: paste text'
                    $chapterName = (Read-Host 'Chapter name').Trim()
                    if (-not $chapterName) {
                        Write-Warn 'Chapter name cannot be blank.'
                        break
                    }
                    $inputText = Read-PastedText
                    Invoke-TtsForText -ChapterName $chapterName -InputText $inputText -Voice $voice -Settings $settings
                }
            } catch {
                Write-ErrorMessage ("Error: {0}" -f $_.Exception.Message)
            }
            Read-Host 'Press Enter to return to the main menu'
        }
        '2' {
            try {
                $filePath = Select-InputFile -Settings $settings
                Split-BookIntoScenes -FilePath $filePath -Settings $settings
            } catch {
                Write-ErrorMessage ("Error: {0}" -f $_.Exception.Message)
            }
            Read-Host 'Press Enter to return to the main menu'
        }
        '9' {
            Clear-HostSafe
            Write-Info 'Settings:'
            Write-Host ("  Workspace folder: {0}" -f $settings.WorkspaceFolder)
            Write-Host ("  Model: {0}" -f $settings.Model)
            Write-Host ("  API key: {0}" -f (Get-ApiKeyPreview -ApiKey $settings.ApiKey))
            Write-Host '  a) Change workspace folder'
            Write-Host '  b) Change TTS model'
            Write-Host '  c) Change API key'
            $settingChoice = (Read-Host 'Choose a setting').Trim().ToLowerInvariant()
            switch ($settingChoice) {
                'a' {
                    $newPath = Read-Host 'Enter new workspace folder'
                    if ($newPath) {
                        $resolvedPath = Resolve-WorkspaceFolder -Path $newPath
                        Write-Info ("Updating workspace folder to: {0}" -f $resolvedPath)
                        $settings.WorkspaceFolder = $resolvedPath
                        Ensure-WorkspaceFolder -Path $settings.WorkspaceFolder
                        Save-Settings -Settings $settings
                        Write-Success 'Workspace folder updated.'
                    }
                }
                'b' {
                    Choose-Model -Settings $settings
                    Save-Settings -Settings $settings
                }
                'c' {
                    $newKey = Read-Host 'Enter new API key'
                    if ($newKey) {
                        Write-Info 'Updating API key in settings.'
                        $settings.ApiKey = $newKey
                        $env:OPENAI_API_KEY = $newKey
                        Save-Settings -Settings $settings
                        Write-Success 'API key updated.'
                    }
                }
                default { Write-Warn 'Unknown settings option.' }
            }
            Read-Host 'Press Enter to return to the main menu'
        }
        '0' { break }
        default { Write-Warn 'Invalid selection.' }
    }
}
