VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbaliml6.ocx"
Object = "{E142732F-A852-11D4-B06C-00500427A693}#1.14#0"; "vbaltbar6.ocx"
Begin VB.Form frmPlayer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Flame Player"
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmr 
      Interval        =   1000
      Left            =   3600
      Top             =   0
   End
   Begin vbalIml6.vbalImageList ilPlayerDis 
      Left            =   360
      Top             =   240
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
      Left            =   3000
      Top             =   360
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
      Top             =   480
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
      Top             =   480
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
'   JaViS:      javisarias@ gmail.com(JaViS)
'   Danko:      lord_danko@users.sourceforge.net (Darío Cutillas)
'   Izubiaurre: izubiaurre@users.sourceforge.net (Imanol Izubiaurre)
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

Private Const MSG_INITFMOD_ERROR = "Could not initialize fmod library. Reason: "
Private Const MSG_LOAD_ERRORLOADING = "Could not load the file. Reason: "

Private Enum AUDIO_TYPE_CONSTANTS
    MUSIC
    SOUND
End Enum

Private Enum status
    ST_STOPPED
    ST_PLAYING
    ST_PAUSED
End Enum

Private audiomode As AUDIO_TYPE_CONSTANTS

Private fmodLoaded As Boolean
Private audioHandle As Long
Private FilePath As String
Private m_st As status

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

Private Sub FreeAudio()
    If fmodLoaded And audioHandle <> 0 Then
        If audiomode = MUSIC Then
            FMUSIC_FreeSong audioHandle
            audioHandle = 0
        Else
            'Sound free
        End If
    End If
End Sub

Private Sub PlayAudio()
    Dim result As Boolean
    
    If fmodLoaded = True And audioHandle <> 0 Then
        If audiomode = MUSIC Then
            Select Case st
            Case ST_STOPPED
                FMUSIC_SetLooping audioHandle, tbrPlayer.ButtonChecked("Repeat")
                result = FMUSIC_PlaySong(audioHandle)
            Case ST_PAUSED
                result = FMUSIC_SetPaused(audioHandle, False)
            End Select
        Else
            'Play sound
        End If
        'Change status
        If result Then
            st = ST_PLAYING
        Else
            MsgBox "Error playing"
        End If
    End If
End Sub

Private Sub PauseAudio()
    If fmodLoaded = True And audioHandle <> 0 Then
        If audiomode = MUSIC Then
            If FMUSIC_SetPaused(audioHandle, True) Then
                st = ST_PAUSED
            End If
        Else
            'Pause sound
        End If
    End If
End Sub

Private Sub StopAudio()
    If fmodLoaded = True And audioHandle <> 0 Then
        If audiomode = MUSIC Then
            If FMUSIC_StopSong(audioHandle) Then
                st = ST_STOPPED
            End If
        Else
            'Stop sound
        End If
    End If
End Sub

Private Sub InitFMod()
    Dim result As Boolean
    result = FSOUND_Init(44100, 32, 0)
    fmodLoaded = False
    If result Then
        fmodLoaded = True
    Else
        'An error occured
        MsgBox MSG_INITFMOD_ERROR & FSOUND_GetErrorString(FSOUND_GetError)
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
        .AddButton eButtonStyle:=CTBSeparator
        .AddButton "Close", 4, , , "", CTBAutoSize, "Close"
    End With
    'Create the rebar
    With Rebar
        If A_Bitmaps Then
            .BackgroundBitmap = App.Path & "\resources\backrebar.bmp"
        End If
        .CreateRebar Me.hwnd
        .AddBandByHwnd tbrPlayer.hwnd, , True, False
    End With
End Sub

Public Function Load(sFile As String) As Long
    Dim lResult As Long
    
    'Initialize the library if necessary
    If fmodLoaded = False Then
        InitFMod
    End If
    
    lResult = 0
    If fmodLoaded = True Then 'Can load the song
        'Reset status
        FilePath = ""
        st = ST_STOPPED
        FreeAudio
        'Load the new song
        audioHandle = FMUSIC_LoadSong(sFile)
        If audioHandle <> 0 Then
            FilePath = sFile
            lblTitle = FSO.GetFileName(sFile)
            PlayAudio
            lResult = -1 'Loading succesfully
        Else
            'Something went wrong
            MsgBox MSG_LOAD_ERRORLOADING & FSOUND_GetErrorString(FSOUND_GetError)
        End If
    End If
    Load = lResult
End Function

Private Sub Form_Unload(Cancel As Integer)
    'Unload audio and close fmod
    FreeAudio
    FSOUND_Close
    fmodLoaded = False
End Sub

Private Sub tbrPlayer_ButtonClick(ByVal lButton As Long)
    If fmodLoaded = True And audioHandle <> 0 Then
        Select Case tbrPlayer.ButtonKey(lButton)
        Case "Play"
            PlayAudio
        Case "Stop"
            StopAudio
        Case "Pause"
            PauseAudio
        Case "Close"
            Unload Me
        End Select
    End If
End Sub

Private Sub tmr_Timer()
    If fmodLoaded And audioHandle <> 0 Then
        If Not st = ST_STOPPED And Not tbrPlayer.ButtonChecked("Repeat") Then
            If audiomode = MUSIC Then
                If FMUSIC_IsFinished(audioHandle) Then
                    StopAudio
                End If
            Else
                'Stop if necessary
            End If
        End If
    End If
End Sub
