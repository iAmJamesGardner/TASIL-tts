# ——————————————————————————————————————————
# 1) Global initialization (run once)
# ——————————————————————————————————————————
$script:i = 0
$script:Key = $env:SPEECH_KEY
$script:Region = $env:SPEECH_REGION
$script:StyleMap = @{}
$script:CharacterNames = @()
$script:voicesUri = "https://$script:Region.tts.speech.microsoft.com/cognitiveservices/voices/list"
$script:TokenEndpoint = "https://$script:Region.api.cognitive.microsoft.com/sts/v1.0/issueToken"
$script:TTSEndpoint = "https://$script:Region.tts.speech.microsoft.com/cognitiveservices/v1"
$script:getVoiceHeaders = @{
    "Ocp-Apim-Subscription-Key" = $script:key
}
$script:CommonHeaders = @{
    'Content-Type'             = 'application/ssml+xml'
    'X-Microsoft-OutputFormat' = 'audio-16khz-128kbitrate-mono-mp3'
}

# ——————————————————————————————————————————
# 2) Token Caching (fetch token only when needed)
# ——————————————————————————————————————————
$script:Token = $null

function Get-SpeechToken {
    if (-not $script:Token) {
        $script:Token = Invoke-RestMethod -Method Post -Uri $script:TokenEndpoint `
            -Headers @{ 'Ocp-Apim-Subscription-Key' = $script:Key } `
            -Body $null
    }
    return $script:Token
}

# ——————————————————————————————————————————
# 3) Get-SpeechVoices function to fetch the voice based on display name
# ——————————————————————————————————————————
function Get-SpeechVoices {
    param([string]$name)
    
    # Fetch the available voices list from Azure
    $results = Invoke-RestMethod -Uri $script:voicesUri -Headers $script:getVoiceHeaders -Method Get

    # Return the ShortName of the voice matching the DisplayName
    if ($name) { 
        return ($results | Where-Object { $_.DisplayName -ieq "$name" }).ShortName
    } else { 
        return $results 
    }
}

# ——————————————————————————————————————————
# 4) Load Voice Mapping from JSON and Cache it
# ——————————————————————————————————————————
$voiceMapping = Get-Content "characters.json" | ConvertFrom-Json
$script:VoiceMap = @{}
foreach ($entry in $voiceMapping) {
    $script:VoiceMap[$entry.Character] = Get-SpeechVoices $entry.Voice
    $script:CharacterNames += $entry.Character
    $script:StyleMap[$entry.Character] = $entry.AvailableStyles
}

# Validate if there are any duplicate voice assignments
function Validate-UniqueVoices {
    $duplicates = $script:VoiceMap.Values | Group-Object | Where-Object { $_.Count -gt 1 }
    if ($duplicates) {
        foreach ($dup in $duplicates) {
            $characters = ($script:VoiceMap.GetEnumerator() | Where-Object { $_.Value -eq $dup.Name } | ForEach-Object { $_.Key }) -join ', '
            Write-Error "Voice '$($dup.Name)' assigned to multiple speakers: $characters"
        }
        throw 'Duplicate voice assignments found – fix your $VoiceMap!'
    }
}
Validate-UniqueVoices

# ——————————————————————————————————————————
# 5) Core TTS Invoker Function (Updated with Token Usage)
# ——————————————————————————————————————————
function Invoke-TTS {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Text,
        [Parameter(Mandatory)][string]$VoiceName,
        [string]$Style,
        [string]$OutputFile = "TASIL-$script:i.mp3",
        [switch]$Play
    )

    # Ensure we have a valid token before making the request
    $script:Token = Get-SpeechToken

    $body =
    if ($Style) { "<mstts:express-as style='$Style'>$Text</mstts:express-as>" }
    else { $Text }

    $ssml = @"
<speak version='1.0' xmlns='https://www.w3.org/2001/10/synthesis'
       xmlns:mstts='https://www.w3.org/2001/mstts' xml:lang='en-US'>
  <voice name='$VoiceName'>$body</voice>
</speak>
"@

    # Set the Authorization header to use the token dynamically
    $h = $script:CommonHeaders.Clone()
    $h['Authorization'] = "Bearer $($script:Token)"  # Dynamically using the fetched token

    # Send the request to the TTS endpoint
    Invoke-RestMethod -Uri $script:TTSEndpoint -Method Post `
        -Headers $h -Body $ssml -OutFile $OutputFile

    $script:i++
    if ($Play) { Start-Process "$OutputFile" }
}


# ——————————————————————————————————————————
# 6) Dynamic Character TTS
# ——————————————————————————————————————————
function Invoke-CharacterTTS {
    param(
        [Parameter(Mandatory)][ValidateScript({@($script:CharacterNames) -contains $_})][string]$CharacterName,
        [string]$Text,
        [string]$Style,
        [string]$OutputFile = "TASIL-$script:i.mp3"
    )

    # Validate if the CharacterName exists in the available characters
    if ($CharacterName -notin $script:CharacterNames) {
        Write-Error "Invalid CharacterName! Available characters: $($script:CharacterNames -join ', ')"
        return
    }

    $voiceName = $script:VoiceMap[$CharacterName]
    if (-not $voiceName) {
        Write-Error "Voice not found for character: $CharacterName"
        return
    }

    # Validate if the style is available for the selected character
    if ($Style -notin $script:StyleMap[$CharacterName]) {
        Write-Error "Invalid Style! Available styles for $CharacterName\: $($script:StyleMap[$CharacterName] -join ', ')"
        return
    }
    

    Invoke-TTS -Text $Text `
        -VoiceName $voiceName `
        -Style $Style `
        -OutputFile $OutputFile
}

# ——————————————————————————————————————————
# 7) Parallel Processing for TTS
# ——————————————————————————————————————————
function Invoke-TTSParallel {
    param(
        [string]$Text,
        [string]$CharacterName,
        [string]$Style
    )
    Start-Job -ScriptBlock {
        Invoke-CharacterTTS -CharacterName $CharacterName -Text $Text -Style $Style
    }
}

# ——————————————————————————————————————————
# 8) Bystander Function for Random Voice
# ——————————————————————————————————————————
function Bystander {
    param($Text, $Style, $OutputFile = "TASIL-$script:i.mp3")
    Invoke-TTS -Text $Text `
        -VoiceName ((Get-SpeechVoices | ? { $_.locale -imatch 'en-US|en-GB|en-AU'}) | Get-Random).ShortName `
        -Style $Style `
        -OutputFile $OutputFile
}
