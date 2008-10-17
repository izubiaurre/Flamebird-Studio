VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbaliml6.ocx"
Object = "{E142732F-A852-11D4-B06C-00500427A693}#1.14#0"; "vbaltbar6.ocx"
Begin VB.Form frmPlayer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Flame Player"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmr 
      Interval        =   10
      Left            =   3600
      Top             =   0
   End
   Begin vbalIml6.vbalImageList ilPlayerDis 
      Left            =   600
      Top             =   360
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   16
      Size            =   5740
      Images          =   "frmPlayer.frx":0000
      Version         =   131072
      KeyCount        =   5
      Keys            =   "ÿÿÿÿCLOSE"
   End
   Begin vbalTBar6.cToolbar tbrPlayer 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   661
   End
   Begin vbalIml6.vbalImageList ilPlayer 
      Left            =   3480
      Top             =   480
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   16
      Size            =   5740
      Images          =   "frmPlayer.frx":168C
      Version         =   131072
      KeyCount        =   5
      Keys            =   "PLAYÿPAUSEÿSTOPÿREPEATÿCLOSE"
   End
   Begin vbalTBar6.cReBar Rebar 
      Left            =   1680
      Top             =   600
      _ExtentX        =   2355
      _ExtentY        =   688
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "No audio file loaded"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   4095
   End
End
Attribute VB_Name = "frmPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Flamebird MX
'Copyright (C) 2003-2007 Flamebird Team
'Contact:
'   JaViS:      javisarias@ gmail.com            (JaViS)
'   Danko:      lord_danko@users.sourceforge.net (Darío Cutillas)
'   Zubiaurre:  izubiaurre@users.sourceforge.net (Imanol Zubiaurre)
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.

Option Explicit

'Private Const MSG_INITFMOD_ERROR = "Could not initialize fmod library. Reason: "
'Private Const MSG_LOAD_ERRORLOADING = "Could not load the file. Reason: "
'
'Private Enum AUDIO_TYPE_CONSTANTS
'    MUSIC
'    Sound
'End Enum

Private Enum status
    ST_STOPPED
    ST_PLAYING
    ST_PAUSED
End Enum

'Private audiomode As AUDIO_TYPE_CONSTANTS

Private fmodLoaded As Boolean
Private audioHandle As Long
Private FilePath As String
Private m_st As status

'' new part
'Dim system As Long
'Dim Sound As Long
'Dim channel As Long
'Dim szFile(20481) As Byte
'Dim szFileTitle(4096) As Byte
'
'Dim brushBlack As Long
'Dim brushWhite As Long
'Dim brushGreen As Long

Private Const GRAPHICWINDOW_WIDTH As Long = 280
Private Const GRAPHICWINDOW_HEIGHT As Long = 75

Private Property Get st() As status
    st = m_st
End Property
Private Property Let st(newVal As status)
    m_st = newVal
    EnableDisableButtons
End Property

Private Sub EnableDisableButtons()
    With tbrPlayer
        Select Case st
        Case ST_PLAYING
            .ButtonEnabled("Pause") = True
            .ButtonEnabled("Stop") = True
            .ButtonEnabled("Play") = False
            .ButtonEnabled("Repeat") = False
        Case ST_STOPPED
            .ButtonEnabled("Pause") = False
            .ButtonEnabled("Stop") = False
            .ButtonEnabled("Play") = True
            .ButtonEnabled("Repeat") = True
        Case ST_PAUSED
            .ButtonEnabled("Pause") = False
            .ButtonEnabled("Stop") = True
            .ButtonEnabled("Play") = True
            .ButtonEnabled("Repeat") = False
        End Select
    End With
End Sub

Private Sub PlayAudio()
    Dim result As FMOD_RESULT
    Dim Isplaying As Long
    Dim paused As Long

    If channel Then
        Call FMOD_Channel_IsPlaying(channel, Isplaying)
        result = FMOD_Channel_GetPaused(channel, paused)
        ERRCHECK (result)
    End If

    If Sound Then 'And Isplaying = 0 Then
        If paused Then
            st = ST_PLAYING
            result = FMOD_Channel_SetPaused(channel, 0)
        Else
            result = FMOD_System_PlaySound(system, FMOD_CHANNEL_FREE, Sound, 0, channel)
            ERRCHECK (result)
        
            st = ST_PLAYING
        End If
    Else
        If channel Then
            Call FMOD_Channel_Stop(channel)
            channel = 0
        End If

        st = ST_STOPPED
    End If
End Sub

Private Sub PauseAudio()
    Dim result As FMOD_RESULT
    Dim paused As Long
    
    If channel Then
        result = FMOD_Channel_GetPaused(channel, paused)
        ERRCHECK (result)
        
        If paused Then
'            st = ST_PLAYING
'            result = FMOD_Channel_SetPaused(channel, 0)
        Else
            st = ST_PAUSED
            result = FMOD_Channel_SetPaused(channel, 1)
        End If
    End If
End Sub

Private Sub StopAudio()
    Dim result As FMOD_RESULT
    Dim Isplaying As Long
    
    If channel Then
        Call FMOD_Channel_IsPlaying(channel, Isplaying)
    End If
    
    If Sound And Isplaying = 0 Then
'        result = FMOD_System_PlaySound(system, FMOD_CHANNEL_FREE, Sound, 0, channel)
'        ERRCHECK (result)
'
'        st = ST_PLAYING
        If channel Then
            Call FMOD_Channel_Stop(channel)
            channel = 0
        End If
        
        st = ST_STOPPED
    Else
        If channel Then
            Call FMOD_Channel_Stop(channel)
            channel = 0
        End If
        
        st = ST_STOPPED
    End If
End Sub

Private Sub RepeatAudio()
    Dim result As FMOD_RESULT
    If tbrPlayer.ButtonChecked("Repeat") Then
        ERRCHECK (FMOD_Sound_SetMode(Sound, FMOD_LOOP_NORMAL))
        result = FMOD_Sound_SetLoopCount(Sound, -1)
        ERRCHECK (result)
    Else
        ERRCHECK (FMOD_Sound_SetMode(Sound, FMOD_LOOP_OFF))
        result = FMOD_Sound_SetLoopCount(Sound, 0)
        ERRCHECK (result)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    'Create toolbar
    With tbrPlayer
        .ImageSource = CTBExternalImageList
        .SetImageList ilPlayer.hIml, CTBImageListNormal
        .SetImageList ilPlayerDis.hIml, CTBImageListDisabled
        .DrawStyle = T_Style
        .CreateToolbar 16, True, True
        .AddButton "Play", 0, sButtonText:="Play", sKey:="Play", eButtonStyle:=CTBAutoSize
        .AddButton "Pause", 1, sButtonText:="Pause", sKey:="Pause", eButtonStyle:=CTBAutoSize
        .AddButton "Stop", 2, sButtonText:="Stop", sKey:="Stop", eButtonStyle:=CTBAutoSize
        .AddButton "Repeat", 3, , , "Repeat", CTBCheck + CTBAutoSize, "Repeat"
'        .AddButton eButtonStyle:=CTBSeparator
'        .AddButton "Close", 4, , , "", CTBAutoSize, "Close"
    End With
    
    'Create the rebar
    With Rebar
        If A_Bitmaps Then
            .BackgroundBitmap = App.Path & "\resources\backrebar" & A_Color & ".bmp"
        End If
        .CreateRebar Me.Hwnd
        .AddBandByHwnd tbrPlayer.Hwnd, , True, False
    End With

'    'start FMOD
'    Dim result As FMOD_RESULT
'    Dim version As Long
'
'    ' Create the brushes we will be using
'    brushBlack = CreateSolidBrush(RGB(0, 0, 0))
'    brushWhite = CreateSolidBrush(RGB(255, 255, 255))
'    brushGreen = CreateSolidBrush(RGB(0, 255, 0))
'
'    ' Create a System object and initialize.
'    result = FMOD_System_Create(system)
'    ERRCHECK (result)
'
'    result = FMOD_System_GetVersion(system, version)
'    ERRCHECK (result)
'
'    If version <> FMOD_VERSION Then
'        MsgBox "Error!  You are using an old version of FMOD " & hex$(version) & ". " & _
'               "This program requires " & hex$(FMOD_VERSION)
'    End If
'
'    result = FMOD_System_Init(system, 32, FMOD_INIT_NORMAL, 0)
'    ERRCHECK (result)

End Sub

Public Function Load(sFile As String) As Long
    Dim result As FMOD_RESULT
    
    On Error GoTo errhandler
    
    If Sound Then
        If channel Then
            Call FMOD_Channel_Stop(channel)
            channel = 0
        End If
        Call FMOD_Sound_Release(Sound)
        Sound = 0
    End If

    lblTitle = FSO.GetFileName(sFile)
    ' Create the stream
    result = FMOD_System_CreateStream(system, sFile, FMOD_2D Or FMOD_SOFTWARE, Sound)
    ERRCHECK (result)
    
    RepeatAudio
'            ERRCHECK (FMOD_Sound_SetMode(Sound, FMOD_LOOP_NORMAL))
'            result = FMOD_Sound_SetLoopCount(Sound, -1)
'            ERRCHECK (result)

    PlayAudio
    st = ST_PLAYING
errhandler:
    Load = -1
End Function

Private Sub Form_Unload(Cancel As Integer)
    StopAudio
'    'Unload audio and close fmod
'    Dim result As FMOD_RESULT
'
'    ' Shut down
'    If Sound Then
'        result = FMOD_Sound_Release(Sound)
'        ERRCHECK (result)
'    End If
'
'    If system Then
'
'        result = FMOD_System_Close(system)
'        ERRCHECK (result)
'
'        result = FMOD_System_Release(system)
'        ERRCHECK (result)
'    End If
End Sub

Private Sub tbrPlayer_ButtonClick(ByVal lButton As Long)
    Select Case tbrPlayer.ButtonKey(lButton)
        Case "Play"
            PlayAudio
        Case "Stop"
            StopAudio
        Case "Pause"
            PauseAudio
        Case "Repeat"
            RepeatAudio
'        Case "Close"
'            Form_Unload 1
    End Select
End Sub

Private Sub tmr_Timer()
    Dim result As FMOD_RESULT
    Dim Isplaying As Long
    Dim hdc As Long
    Dim hdcbuffer As Long
    Dim hdcmem As Long
    Dim hbmold As Long
    Dim hbmbuffer As Long
    Dim hbmoldbuffer As Long
    Dim Rectangle As RECT

    hdc = GetDC(Me.Hwnd)
    hdcbuffer = CreateCompatibleDC(hdc)
    hbmbuffer = CreateCompatibleBitmap(hdc, GRAPHICWINDOW_WIDTH, GRAPHICWINDOW_HEIGHT)
    hbmoldbuffer = SelectObject(hdcbuffer, hbmbuffer)
    hdcmem = CreateCompatibleDC(hdc)
    hbmold = SelectObject(hdcmem, 0)
    
    GetClientRect Me.Hwnd, Rectangle
    
    FillRect hdcbuffer, Rectangle, brushBlack
    
    If system Then
        DrawSpectrum (hdcbuffer)
        'DrawOscilliscope (hdcbuffer)
        
        result = FMOD_System_Update(system)
        ERRCHECK (result)
        If Not st = ST_STOPPED And Not tbrPlayer.ButtonChecked("Repeat") Then
            Call FMOD_Channel_IsPlaying(channel, Isplaying)
            ERRCHECK (result)
            Debug.Print Isplaying
            If Isplaying = 0 Then
                StopAudio
            End If
        End If
    End If
    
    BitBlt hdc, Rectangle.Left, Rectangle.Top, GRAPHICWINDOW_WIDTH, GRAPHICWINDOW_HEIGHT, hdcbuffer, Rectangle.Left, Rectangle.Top, vbSrcCopy

    SelectObject hdcmem, hbmold
    DeleteDC hdcmem
    SelectObject hdcbuffer, hbmoldbuffer
    DeleteObject hbmbuffer
    DeleteDC hdcbuffer
    ReleaseDC Me.Hwnd, hdc

End Sub

Private Sub DrawSpectrum(hdcbuffer As Long)
    Dim result As FMOD_RESULT
    Dim spectrum(512) As Single
    Dim count As Long
    Dim count2 As Long
    Dim Numchannels As Long
    Dim line As RECT
    Dim max As Single
    
    result = FMOD_System_GetSoftwareFormat(system, 0, 0, Numchannels, 0, 0, 0)
    ERRCHECK (result)

    '
    ' Draw Spectrum
    '
    For count = 0 To Numchannels - 1
        result = FMOD_System_GetSpectrum(system, spectrum(0), 512, count, FMOD_DSP_FFT_WINDOW_TRIANGLE)
        ERRCHECK (result)
        
        For count2 = 0 To 255
            If max < spectrum(count2) Then
                max = spectrum(count2)
            End If
        Next
        
        ' Draw the actual spectrum
        ' The upper band of frequencies at 44khz is pretty boring (ie 11-22khz), so we are only
        ' going to display the first 256 frequencies, or (0-11khz)
        For count2 = 0 To 255
            Dim Height As Single
            
            Height = spectrum(count2) / max * GRAPHICWINDOW_HEIGHT
            
            If Height >= GRAPHICWINDOW_HEIGHT Then
                Height = Height - 1#
            End If
            
            If Height < 0 Then
                Height = 0#
            End If
            
            Height = GRAPHICWINDOW_HEIGHT - Height
            
            line.Bottom = GRAPHICWINDOW_HEIGHT
            line.Top = Height
            line.Left = count2
            line.Right = count2 + 1
            
            'FillRect hdcbuffer, line, brush(Height)
            
            FillRect hdcbuffer, line, brushGreen
        Next
    Next
End Sub
'
'Private Sub ERRCHECK(result As FMOD_RESULT)
'    Dim msgResult As VbMsgBoxResult
'
'    If result <> FMOD_OK Then
'        msgResult = MsgBox("FMOD error! (" & result & ") " & FMOD_ErrorString(result))
'    End If
'
'    If msgResult Then
'        End
'    End If
'End Sub
