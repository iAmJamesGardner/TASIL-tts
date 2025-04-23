# ——————————————————————————————————————————
# 1) Global initialization (run once)
# ——————————————————————————————————————————
$script:i = 0
$script:Key = $env:SPEECH_KEY
$script:Region = $env:SPEECH_REGION
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

function Get-SpeechVoices($name) {
    $results = Invoke-RestMethod -Uri $script:voicesUri -Headers $script:getVoiceHeaders -Method Get
    if ($name) { ($results | ? { $_.DisplayName -ieq "$name" }).ShortName } else { $results }
}

function Get-SpeechToken {
    Invoke-RestMethod -Method Post -Uri $script:TokenEndpoint `
        -Headers @{ 'Ocp-Apim-Subscription-Key' = $script:Key } `
        -Body $null
}
$script:Token = Get-SpeechToken

# ——————————————————————————————————————————
# 2) Define your speaker→voice map in one spot
# ——————————————————————————————————————————
$script:VoiceMap = @{
    Narrator        = Get-SpeechVoices 'Adam Multilingual'
    CharlieWade     = Get-SpeechVoices 'Brandon Multilingual'
    ClaireWilson    = Get-SpeechVoices Sara
    DorisYoung      = Get-SpeechVoices Aria
    WendellJones    = Get-SpeechVoices Guy
    LadyWilson      = Get-SpeechVoices Nancy
    HaroldWilson    = Get-SpeechVoices 'Brian Multilingual'
    WendyWilson     = Get-SpeechVoices Amber
    GeraldWhite     = Get-SpeechVoices Andrew
    KevinWhite      = Get-SpeechVoices 'Ryan Multilingual'
    StephenThompson = Get-SpeechVoices 'Ollie Multilingual'
    SabrinaLee      = Get-SpeechVoices 'Emma Dragon HD Latest'
    IsaacCameron    = Get-SpeechVoices Ryan
    ElaineWilson    = Get-SpeechVoices 'Aria Dragon HD Latest'
    JacobWilson     = Get-SpeechVoices 'Alloy Dragon HD Latest'
    MrsLewis        = Get-SpeechVoices 'Ava Dragon HD Latest'
    LordWade        = Get-SpeechVoices Thomas
    MrJones         = Get-SpeechVoices Tony
    CaptainCooper   = Get-SpeechVoices Kai
    Nurse           = Get-SpeechVoices Ethan
    Laird           = Get-SpeechVoices Annette
}

function Validate-UniqueVoices {
    $dups = $script:VoiceMap.GetEnumerator() |
    Group-Object -Property Value |
    Where-Object Count -gt 1
    if ($dups) {
        $dups | ForEach-Object {
            $names = ($_.Group | ForEach-Object Key) -join ', '
            Write-Error "Voice '$($_.Name)' assigned to multiple speakers: $names"
        }
        throw 'Duplicate voice assignments found – fix your $VoiceMap!'
    }
}
Validate-UniqueVoices

# ——————————————————————————————————————————
# 3) Core TTS invoker
# ——————————————————————————————————————————
function Invoke-TTS {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Text,
        [Parameter(Mandatory)][string]$VoiceName,
        [string]$Style,
        [string]$OutputFile = "TASIL-$script:i.mp3"
    )
    $body =
    if ($Style) { "<mstts:express-as style='$Style'>$Text</mstts:express-as>" }
    else { $Text }

    $ssml = @"
<speak version='1.0' xmlns='https://www.w3.org/2001/10/synthesis'
       xmlns:mstts='https://www.w3.org/2001/mstts' xml:lang='en-US'>
  <voice name='$VoiceName'>$body</voice>
</speak>
"@

    $h = $script:CommonHeaders.Clone()
    $h['Authorization'] = "Bearer $($script:Token)"

    Invoke-RestMethod -Uri $script:TTSEndpoint -Method Post `
        -Headers $h -Body $ssml -OutFile $OutputFile

    $script:i++
    Start-Process ".\$OutputFile"
}

# ——————————————————————————————————————————
# 4) Thin wrappers per speaker
# ——————————————————————————————————————————
function MyNarrator {
    param($Text, $Style, $OutputFile = "TASIL-$script:i.mp3")
    Invoke-TTS -Text $Text `
        -VoiceName $script:VoiceMap['Narrator'] `
        -Style $Style `
        -OutputFile $OutputFile
}

function CharlieWade {
    param($Text, $Style, $OutputFile = "TASIL-$script:i.mp3")
    Invoke-TTS -Text $Text `
        -VoiceName $script:VoiceMap['CharlieWade'] `
        -Style $Style `
        -OutputFile $OutputFile
}

function ClaireWilson {
    param($Text, $Style, $OutputFile = "TASIL-$script:i.mp3")
    Invoke-TTS -Text $Text `
        -VoiceName $script:VoiceMap['ClaireWilson'] `
        -Style $Style `
        -OutputFile $OutputFile
}

function DorisYoung {
    param($Text, $Style, $OutputFile = "TASIL-$script:i.mp3")
    Invoke-TTS -Text $Text `
        -VoiceName $script:VoiceMap['DorisYoung'] `
        -Style $Style `
        -OutputFile $OutputFile
}

function WendellJones {
    param($Text, $Style, $OutputFile = "TASIL-$script:i.mp3")
    Invoke-TTS -Text $Text `
        -VoiceName $script:VoiceMap['WendellJones'] `
        -Style $Style `
        -OutputFile $OutputFile
}

function LadyWilson {
    param($Text, $Style, $OutputFile = "TASIL-$script:i.mp3")
    Invoke-TTS -Text $Text `
        -VoiceName $script:VoiceMap['LadyWilson'] `
        -Style $Style `
        -OutputFile $OutputFile

}
function HaroldWilson {
    param($Text, $Style, $OutputFile = "TASIL-$script:i.mp3")
    Invoke-TTS -Text $Text `
        -VoiceName $script:VoiceMap['HaroldWilson'] `
        -Style $Style `
        -OutputFile $OutputFile

}
function WendyWilson {
    param($Text, $Style, $OutputFile = "TASIL-$script:i.mp3")
    Invoke-TTS -Text $Text `
        -VoiceName $script:VoiceMap['WendyWilson'] `
        -Style $Style `
        -OutputFile $OutputFile

}
function GeraldWhite {
    param($Text, $Style, $OutputFile = "TASIL-$script:i.mp3")
    Invoke-TTS -Text $Text `
        -VoiceName $script:VoiceMap['GeraldWhite'] `
        -Style $Style `
        -OutputFile $OutputFile

}
function KevinWhite {
    param($Text, $Style, $OutputFile = "TASIL-$script:i.mp3")
    Invoke-TTS -Text $Text `
        -VoiceName $script:VoiceMap['KevinWhite'] `
        -Style $Style `
        -OutputFile $OutputFile

}
function StephenThompson {
    param($Text, $Style, $OutputFile = "TASIL-$script:i.mp3")
    Invoke-TTS -Text $Text `
        -VoiceName $script:VoiceMap['StephenThompson'] `
        -Style $Style `
        -OutputFile $OutputFile

}
function SabrinaLee {
    param($Text, $Style, $OutputFile = "TASIL-$script:i.mp3")
    Invoke-TTS -Text $Text `
        -VoiceName $script:VoiceMap['SabrinaLee'] `
        -Style $Style `
        -OutputFile $OutputFile

}
function IsaacCameron {
    param($Text, $Style, $OutputFile = "TASIL-$script:i.mp3")
    Invoke-TTS -Text $Text `
        -VoiceName $script:VoiceMap['IsaacCameron'] `
        -Style $Style `
        -OutputFile $OutputFile

}
function JaneWolfe {
    param($Text, $Style, $OutputFile = "TASIL-$script:i.mp3")
    Invoke-TTS -Text $Text `
        -VoiceName $script:VoiceMap['JaneWolfe'] `
        -Style $Style `
        -OutputFile $OutputFile

}
function ElaineWilson {
    param($Text, $Style, $OutputFile = "TASIL-$script:i.mp3")
    Invoke-TTS -Text $Text `
        -VoiceName $script:VoiceMap['ElaineWilson'] `
        -Style $Style `
        -OutputFile $OutputFile

}
function JacobWilson {
    param($Text, $Style, $OutputFile = "TASIL-$script:i.mp3")
    Invoke-TTS -Text $Text `
        -VoiceName $script:VoiceMap['JacobWilson'] `
        -Style $Style `
        -OutputFile $OutputFile

}
function MrsLewis {
    param($Text, $Style, $OutputFile = "TASIL-$script:i.mp3")
    Invoke-TTS -Text $Text `
        -VoiceName $script:VoiceMap['MrsLewis'] `
        -Style $Style `
        -OutputFile $OutputFile

}
function LordWade {
    param($Text, $Style, $OutputFile = "TASIL-$script:i.mp3")
    Invoke-TTS -Text $Text `
        -VoiceName $script:VoiceMap['LordWade'] `
        -Style $Style `
        -OutputFile $OutputFile

}
function MrJones {
    param($Text, $Style, $OutputFile = "TASIL-$script:i.mp3")
    Invoke-TTS -Text $Text `
        -VoiceName $script:VoiceMap['MrJones'] `
        -Style $Style `
        -OutputFile $OutputFile

}
function CaptainCooper {
    param($Text, $Style, $OutputFile = "TASIL-$script:i.mp3")
    Invoke-TTS -Text $Text `
        -VoiceName $script:VoiceMap['CaptainCooper'] `
        -Style $Style `
        -OutputFile $OutputFile

}
function Nurse {
    param($Text, $Style, $OutputFile = "TASIL-$script:i.mp3")
    Invoke-TTS -Text $Text `
        -VoiceName $script:VoiceMap['Nurse'] `
        -Style $Style `
        -OutputFile $OutputFile

}
function Laird {
    param($Text, $Style, $OutputFile = "TASIL-$script:i.mp3")
    Invoke-TTS -Text $Text `
        -VoiceName $script:VoiceMap['Laird'] `
        -Style $Style `
        -OutputFile $OutputFile

}