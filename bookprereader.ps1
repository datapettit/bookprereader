#Requires -Version 7.0

$ErrorActionPreference = 'Stop'

$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$SettingsPath = Join-Path $ScriptRoot 'settings.json'
$MaxInputCharacters = 4096
$SupportedModels = @('gpt-4o-mini-tts', 'tts-1', 'tts-1-hd')
$SupportedVoices = @('alloy', 'echo', 'fable', 'onyx', 'nova', 'shimmer')

function Get-DefaultSettings {
    return [ordered]@{
        WorkspaceFolder = 'C:\git\'
        Model = 'gpt-4o-mini-tts'
        ApiKey = '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'
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

function Get-ApiKey {
    param([object]$Settings)

    $candidate = if ($env:OPENAI_API_KEY) { $env:OPENAI_API_KEY } else { $Settings.ApiKey }
    if ([string]::IsNullOrWhiteSpace($candidate)) {
        throw 'OpenAI API key is missing. Set OPENAI_API_KEY or update it in Settings.'
    }
    return $candidate
}

function Select-Voice {
    Write-Host 'Select a voice model:'
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
                return $SupportedVoices[$index]
            }
        }
        $match = $SupportedVoices | Where-Object { $_ -eq $inputValue.ToLowerInvariant() }
        if ($match) {
            return $match
        }
        Write-Host 'Invalid selection. Try again.'
    }
}

function Select-InputMethod {
    Write-Host 'Input method:'
    Write-Host '  1) File upload'
    Write-Host '  2) Paste text'
    while ($true) {
        $inputValue = (Read-Host 'Choose file or text').Trim().ToLowerInvariant()
        switch ($inputValue) {
            '1' { return 'file' }
            '2' { return 'text' }
            'file' { return 'file' }
            'text' { return 'text' }
        }
        Write-Host 'Invalid selection. Enter 1, 2, file, or text.'
    }
}

function Select-InputFile {
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    } catch {
        Write-Host 'File dialog not available. Falling back to manual path entry.'
    }

    if ([System.Type]::GetType('System.Windows.Forms.OpenFileDialog')) {
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Title = 'Select a text or Word document'
        $dialog.Filter = 'Text and Word Documents|*.txt;*.text;*.rtf;*.doc;*.docx|All Files|*.*'
        $dialog.Multiselect = $false

        $result = $dialog.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK -and $dialog.FileName) {
            return $dialog.FileName
        }
    }

    $manualPath = Read-Host 'Enter the full path to the input file'
    if (-not (Test-Path $manualPath)) {
        throw "File not found: $manualPath"
    }
    return $manualPath
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

function Read-PastedText {
    Write-Host 'Paste text below. Enter a single line with END to finish.'
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
    $paragraphs = $Text -split "(\r?\n){2,}"
    $current = ''

    foreach ($paragraph in $paragraphs) {
        $clean = $paragraph.Trim()
        if (-not $clean) {
            continue
        }
        if ($clean.Length -gt $Limit) {
            $sentences = $clean -split '(?<=[.!?])\s+'
            foreach ($sentence in $sentences) {
                if ([string]::IsNullOrWhiteSpace($sentence)) {
                    continue
                }
                if (($current.Length + $sentence.Length + 2) -gt $Limit) {
                    if ($current) {
                        $chunks.Add($current.Trim())
                        $current = ''
                    }
                }
                if ($sentence.Length -gt $Limit) {
                    $offset = 0
                    while ($offset -lt $sentence.Length) {
                        $sliceLength = [Math]::Min($Limit, $sentence.Length - $offset)
                        $chunks.Add($sentence.Substring($offset, $sliceLength))
                        $offset += $sliceLength
                    }
                } else {
                    if ($current) {
                        $current += "`n`n$sentence"
                    } else {
                        $current = $sentence
                    }
                }
            }
            continue
        }

        if (($current.Length + $clean.Length + 2) -gt $Limit) {
            $chunks.Add($current.Trim())
            $current = $clean
        } else {
            if ($current) {
                $current += "`n`n$clean"
            } else {
                $current = $clean
            }
        }
    }

    if ($current) {
        $chunks.Add($current.Trim())
    }

    return $chunks
}

function Invoke-OpenAITts {
    param(
        [string]$Text,
        [string]$Voice,
        [string]$Model,
        [string]$ApiKey,
        [string]$OutputPath
    )

    $headers = @{ Authorization = "Bearer $ApiKey" }
    $body = @{
        model = $Model
        input = $Text
        voice = $Voice
        format = 'mp3'
    } | ConvertTo-Json -Depth 4

    if (Test-Path $OutputPath) {
        Remove-Item $OutputPath -Force
    }

    try {
        Invoke-WebRequest -Method Post -Uri 'https://api.openai.com/v1/audio/speech' -Headers $headers -Body $body -ContentType 'application/json' -OutFile $OutputPath -ErrorAction Stop
    } catch {
        if (Test-Path $OutputPath) {
            Remove-Item $OutputPath -Force
        }
        $response = $_.Exception.Response
        if ($response -and $response.GetResponseStream()) {
            $reader = New-Object System.IO.StreamReader($response.GetResponseStream())
            $details = $reader.ReadToEnd()
            throw "OpenAI TTS request failed: $details"
        }
        throw
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

    Write-Host 'FFmpeg not found. Downloading a free build from gyan.dev...'
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

function Merge-Mp3Files {
    param(
        [string[]]$Files,
        [string]$OutputPath
    )

    $ffmpeg = Ensure-Ffmpeg
    $listPath = Join-Path $ScriptRoot ('ffmpeg-list-' + [Guid]::NewGuid().ToString('N') + '.txt')
    $content = $Files | ForEach-Object { "file '$($_.Replace("'", "''"))'" }
    Set-Content -Path $listPath -Value $content -Encoding UTF8

    if (Test-Path $OutputPath) {
        Remove-Item $OutputPath -Force
    }

    try {
        & $ffmpeg -y -f concat -safe 0 -i $listPath -c copy $OutputPath | Out-Null
    } finally {
        if (Test-Path $listPath) {
            Remove-Item $listPath -Force
        }
    }
}

function Choose-Model {
    param([object]$Settings)

    Write-Host 'Select a TTS model:'
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
                return
            }
        }
        $match = $SupportedModels | Where-Object { $_ -eq $inputValue }
        if ($match) {
            $Settings.Model = $match
            return
        }
        Write-Host 'Invalid selection. Try again.'
    }
}

function Get-ApiKeyPreview {
    param([string]$ApiKey)

    if ([string]::IsNullOrWhiteSpace($ApiKey)) {
        return '(not set)'
    }
    if ($ApiKey.Length -le 8) {
        return $ApiKey
    }
    return $ApiKey.Substring(0, 8) + '...'
}

$settings = Load-Settings
Ensure-WorkspaceFolder -Path $settings.WorkspaceFolder
Save-Settings -Settings $settings

while ($true) {
    Write-Host ''
    Write-Host 'Main Menu'
    Write-Host '1) Create audio from file'
    Write-Host '9) Settings'
    Write-Host '0) Exit'
    $selection = (Read-Host 'Choose an option').Trim()

    switch ($selection) {
        '1' {
            $chapterName = (Read-Host 'Chapter name').Trim()
            if (-not $chapterName) {
                Write-Host 'Chapter name cannot be blank.'
                break
            }

            $voice = Select-Voice
            $inputMethod = Select-InputMethod

            if ($inputMethod -eq 'file') {
                $filePath = Select-InputFile
                $inputText = Get-TextFromFile -Path $filePath
            } else {
                $inputText = Read-PastedText
            }

            $inputText = $inputText.Trim()
            if (-not $inputText) {
                Write-Host 'No text provided.'
                break
            }

            try {
                $chunks = Split-TextIntoChunks -Text $inputText -Limit $MaxInputCharacters
                if ($chunks.Count -eq 0) {
                    Write-Host 'No usable text detected after splitting.'
                    break
                }
                $apiKey = Get-ApiKey -Settings $settings
                $outputFiles = New-Object System.Collections.Generic.List[string]

                for ($i = 0; $i -lt $chunks.Count; $i++) {
                    $index = $i + 1
                    $chunkPath = Join-Path $settings.WorkspaceFolder ("{0}_{1}.mp3" -f $chapterName, $index)
                    Write-Host ("Creating audio chunk {0}/{1}..." -f $index, $chunks.Count)
                    Invoke-OpenAITts -Text $chunks[$i] -Voice $voice -Model $settings.Model -ApiKey $apiKey -OutputPath $chunkPath
                    $outputFiles.Add($chunkPath)
                }

                $finalPath = Join-Path $settings.WorkspaceFolder ("{0}.mp3" -f $chapterName)
                if ($outputFiles.Count -gt 1) {
                    Write-Host 'Merging chunks into final MP3...'
                    Merge-Mp3Files -Files $outputFiles -OutputPath $finalPath
                    foreach ($file in $outputFiles) {
                        Remove-Item $file -Force
                    }
                } else {
                    if (Test-Path $finalPath) {
                        Remove-Item $finalPath -Force
                    }
                    Move-Item -Path $outputFiles[0] -Destination $finalPath
                }
                Write-Host "Saved MP3: $finalPath"
            } catch {
                Write-Host "Error: $($_.Exception.Message)"
            }
        }
        '9' {
            Write-Host 'Settings:'
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
                        $settings.WorkspaceFolder = $newPath
                        Ensure-WorkspaceFolder -Path $settings.WorkspaceFolder
                        Save-Settings -Settings $settings
                    }
                }
                'b' {
                    Choose-Model -Settings $settings
                    Save-Settings -Settings $settings
                }
                'c' {
                    $newKey = Read-Host 'Enter new API key'
                    if ($newKey) {
                        $settings.ApiKey = $newKey
                        $env:OPENAI_API_KEY = $newKey
                        Save-Settings -Settings $settings
                    }
                }
                default { Write-Host 'Unknown settings option.' }
            }
        }
        '0' { break }
        default { Write-Host 'Invalid selection.' }
    }
}
