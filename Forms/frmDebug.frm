VERSION 5.00
Begin VB.Form frmDebug 
   Caption         =   "Debug"
   ClientHeight    =   1815
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4005
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDebug.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   4005
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOutput 
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
End
Attribute VB_Name = "frmDebug"
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

Implements ITDockMoveEvents

Private Function ITDockMoveEvents_DockChange(tDockAlign As AlignConstants, tDocked As Boolean) As Variant
       
End Function

Private Function ITDockMoveEvents_Move(Left As Integer, Top As Integer, Bottom As Integer, Right As Integer)
    Dim Width As Integer
    If frmMain.WindowState <> vbMinimized Then
        Width = IIf(Right < 0, Me.ScaleWidth, Right)
        txtOutput.Move Left, Top, Width, Bottom
    End If
End Function

