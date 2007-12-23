VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbaliml6.ocx"
Object = "{E142732F-A852-11D4-B06C-00500427A693}#1.14#0"; "vbaltbar6.ocx"
Begin VB.Form frmWebBrowser 
   Caption         =   "Web Broser"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6615
   ControlBox      =   0   'False
   Icon            =   "frmWebBrowser.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   6615
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbURL 
      Height          =   315
      Left            =   1560
      TabIndex        =   1
      Top             =   4920
      Width           =   4095
   End
   Begin vbalIml6.vbalImageList iml 
      Left            =   5280
      Top             =   2160
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   16
      Size            =   5740
      Images          =   "frmWebBrowser.frx":038A
      Version         =   131072
      KeyCount        =   5
      Keys            =   "BACKÿFORWARDÿREFRESHÿSTOPÿGO"
   End
   Begin vbalTBar6.cReBar ReBar 
      Left            =   2160
      Top             =   0
      _ExtentX        =   1508
      _ExtentY        =   661
   End
   Begin vbalTBar6.cToolbar tbrWeb 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   4695
      ExtentX         =   8281
      ExtentY         =   6588
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmWebBrowser"
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

Private m_URL As String
Private m_cFlat As cFlatControl

Public Property Let URL(newURL As String)
    If newURL <> "" Then
        m_URL = newURL
        wb.Navigate newURL
    End If
End Property

Public Property Get URL() As String
    URL = m_URL
End Property

Private Sub cmbURL_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        tbrWeb.RaiseButtonClick (tbrWeb.ButtonIndex("GO"))
    End If
End Sub

Private Sub Form_Load()
    'Configure toolbar
    With tbrWeb
        .ImageSource = CTBExternalImageList
        .DrawStyle = T_Style
        .SetImageList iml.hIml, CTBImageListNormal
        .CreateToolbar 16, False, False, True, 16
        .AddButton "Go back", iml.ItemIndex("BACK") - 1, sKey:="GO_BACK"
        .AddButton "Go forward", iml.ItemIndex("FORWARD") - 1, sKey:="GO_FORWARD"
        .AddButton "Refresh", iml.ItemIndex("REFRESH") - 1, sKey:="REFRESH"
        .AddButton "Stop", iml.ItemIndex("STOP") - 1, sKey:="STOP"
        .AddControl cmbURL.hwnd, , "URL"
        .ControlStretch("URL") = True
        .AddButton "Go", iml.ItemIndex("GO") - 1, sKey:="GO"
    End With
    'Create the rebar
    With ReBar
        If A_Bitmaps Then
            .BackgroundBitmap = App.Path & "\resources\backrebar.bmp"
        End If
        .CreateRebar Me.hwnd
        .AddBandByHwnd tbrWeb.hwnd, , True, False
    End With
    ReBar.RebarSize
    
    Set m_cFlat = New cFlatControl
    m_cFlat.Attach cmbURL
End Sub

Private Sub Form_Resize()
    If frmMain.WindowState <> vbMinimized Then
        wb.Move 0, ScaleY(ReBar.RebarHeight, vbPixels, vbTwips), Me.ScaleWidth, Me.ScaleHeight - ScaleY(ReBar.RebarHeight, vbPixels, vbTwips)
        ReBar.RebarSize
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReBar.RemoveAllRebarBands
    Set m_cFlat = Nothing
End Sub

Private Sub tbrWeb_ButtonClick(ByVal lButton As Long)
    Select Case tbrWeb.ButtonKey(lButton)
    Case "GO_BACK"
        wb.GoBack
    Case "GO_FORWARD"
        wb.GoForward
    Case "STOP"
        wb.Stop
    Case "REFRESH"
        wb.Refresh
    Case "GO"
        wb.Navigate cmbURL.Text
    End Select
End Sub

Private Sub wb_NewWindow2(ppDisp As Object, Cancel As Boolean)
   Dim frmWB As frmWebBrowser
   Set frmWB = New frmWebBrowser

   frmWB.wb.RegisterAsBrowser = True

   Set ppDisp = frmWB.wb.Object
   frmWB.Visible = True
   
   frmMain.RefreshTabs
End Sub

