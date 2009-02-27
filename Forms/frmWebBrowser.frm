VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbaliml6.ocx"
Object = "{E142732F-A852-11D4-B06C-00500427A693}#1.14#0"; "vbaltbar6.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
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
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   3375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3735
      ExtentX         =   6588
      ExtentY         =   5953
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
      Location        =   ""
   End
   Begin vbalTBar6.cReBar ReBar 
      Left            =   3480
      Top             =   0
      _ExtentX        =   1720
      _ExtentY        =   873
   End
   Begin vbalIml6.vbalImageList iml 
      Left            =   5160
      Top             =   720
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   24
      Size            =   5740
      Images          =   "frmWebBrowser.frx":0442
      Version         =   131072
      KeyCount        =   5
      Keys            =   "ÿÿÿÿ"
   End
   Begin vbalTBar6.cToolbar tbrWeb 
      Height          =   375
      Left            =   0
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
   End
   Begin VB.ComboBox cmbURL 
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   4920
      Width           =   4095
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
'        .AddButton "Go back", iml.ItemIndex("BACK") - 1 ', sKey:="GO_BACK"
'        .AddButton "Go forward", iml.ItemIndex("FORWARD") - 1 ', sKey:="GO_FORWARD"
'        .AddButton "Refresh", iml.ItemIndex("REFRESH") - 1 ', sKey:="REFRESH"
'        .AddButton "Stop", iml.ItemIndex("STOP") - 1 ', sKey:="STOP"
'        .AddControl cmbURL.hWnd ', , "URL"
'        .ControlStretch("URL") = True
'        .AddButton "Go", iml.ItemIndex("GO") - 1 ', sKey:="GO"
        .AddButton "Go back", 1, sKey:="GO_BACK"
        .AddButton "Go forward", 2, sKey:="GO_FORWARD"
        .AddButton "Refresh", 3, sKey:="REFRESH"
        .AddButton "Stop", 4, sKey:="STOP"
        .AddControl cmbURL.Hwnd, , "URL"
        .ControlStretch("URL") = True
        .AddButton "Go", 0, sKey:="GO"
    End With
    'Create the rebar
    With rebar
        If A_Bitmaps Then
            .BackgroundBitmap = App.Path & "\resources\backrebar" & A_Color & ".bmp"
        End If
        .CreateRebar Me.Hwnd
        .AddBandByHwnd tbrWeb.Hwnd, , True, False
    End With
    rebar.RebarSize
    
    Set m_cFlat = New cFlatControl
    m_cFlat.Attach cmbURL
End Sub

Private Sub Form_Resize()
    If frmMain.WindowState <> vbMinimized Then
        wb.Move 0, ScaleY(rebar.RebarHeight, vbPixels, vbTwips), Me.ScaleWidth, Me.ScaleHeight - ScaleY(rebar.RebarHeight, vbPixels, vbTwips)
        rebar.RebarSize
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rebar.RemoveAllRebarBands
    Set m_cFlat = Nothing
End Sub

Private Sub tbrWeb_ButtonClick(ByVal lButton As Long)
On Error GoTo hay
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
        wb.Navigate cmbURL.text
    End Select
    
hay:
End Sub

Private Sub wb_NewWindow2(ppDisp As Object, Cancel As Boolean)
   Dim frmWB As frmWebBrowser
   Set frmWB = New frmWebBrowser

   frmWB.wb.RegisterAsBrowser = True

   Set ppDisp = frmWB.wb.Object
   frmWB.Visible = True
   
   frmMain.RefreshTabs
End Sub

