Attribute VB_Name = "fmod_init"
Option Explicit
Option Base 0

' new part
Public system As Long
Public Sound As Long
Public channel As Long
Public szFile(20481) As Byte
Public szFileTitle(4096) As Byte

Public brushBlack As Long
Public brushWhite As Long
Public brushGreen As Long
'Public brushRed As Long
'Public brushYellow As Long
'Public brush(75) As Long

Public Sub initFMOD()
    'start FMOD
    Dim result As FMOD_RESULT
    Dim version As Long
    
    ' Create the brushes we will be using
    brushBlack = CreateSolidBrush(RGB(0, 0, 0))
    brushWhite = CreateSolidBrush(RGB(255, 255, 255))
    brushGreen = CreateSolidBrush(RGB(0, 255, 0))
    'brushRed = CreateSolidBrush(RGB(255, 0, 0))
    'brushYellow = CreateSolidBrush(RGB(255, 255, 0))
    
    ' Maybe later
    'createSpectrumBrush &H0, &HFFFFFF

    ' Create a System object and initialize.
    result = FMOD_System_Create(system)
    ERRCHECK (result)

    result = FMOD_System_GetVersion(system, version)
    ERRCHECK (result)

    If version <> FMOD_VERSION Then
        MsgBox "Error!  You are using an old version of FMOD " & hex$(version) & ". " & _
               "This program requires " & hex$(FMOD_VERSION)
    End If

    result = FMOD_System_Init(system, 32, FMOD_INIT_NORMAL, 0)
    ERRCHECK (result)
End Sub

Public Sub finishFMOD()
    'Unload audio and close fmod
    Dim result As FMOD_RESULT
    
    ' Shut down
    If Sound Then
        result = FMOD_Sound_Release(Sound)
        ERRCHECK (result)
    End If
    
    If system Then

        result = FMOD_System_Close(system)
        ERRCHECK (result)
        
        result = FMOD_System_Release(system)
        ERRCHECK (result)
    End If
End Sub


Public Sub ERRCHECK(result As FMOD_RESULT)
    Dim msgResult As VbMsgBoxResult
    
    If result <> FMOD_OK Then
        msgResult = MsgBox("FMOD error! (" & result & ") " & FMOD_ErrorString(result))
    End If
    
    If msgResult Then
        'End
        Exit Sub
    End If
End Sub
