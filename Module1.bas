Attribute VB_Name = "Module1"
Public DirectX As New DirectX7
Public DD As DirectDraw7
Public DS As DirectSound

Public Type Grenade
X As Long
Y As Long
SV As Long
V As Long
NB As Long
Used As Boolean
End Type

Public Type Thing
X As Long
Y As Long
Wid As Long
Hgt As Long
End Type
Public Type Particle
Life As Long
X As Long
Y As Long
SV As Long
V As Long
End Type

Function ExModeActive() As Boolean
     Dim TestCoopRes As Long 'holds the return value of the test.

     TestCoopRes = DD.TestCooperativeLevel 'Tells DDraw to do the test

     If (TestCoopRes = DD_OK) Then
         ExModeActive = True 'everything is fine
     Else
         ExModeActive = False 'this computer doesn't support this mode
     End If
End Function


Public Sub LoadWave(File As String, ByRef Buffer As DirectSoundBuffer)

    Dim bufferDesc As DSBUFFERDESC
    Dim waveFormat As WAVEFORMATEX
    Set Buffer = Nothing
    
    bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
    waveFormat.nFormatTag = WAVE_FORMAT_PCM
    waveFormat.nChannels = 2    '
    waveFormat.lSamplesPerSec = 44100
    waveFormat.nBitsPerSample = 16
    waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
    waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign
    Set Buffer = DS.CreateSoundBufferFromFile(File, bufferDesc, waveFormat)
    'checks for any errors
    If Err.Number <> 0 Then
        MsgBox "Error with Sound!: Unable to find " + File, vbCritical
    Form1.EndIt
    End If

End Sub
