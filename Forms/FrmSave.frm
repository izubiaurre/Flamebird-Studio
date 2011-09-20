VERSION 5.00
Object = "{CA5A8E1E-C861-4345-8FF8-EF0A27CD4236}#1.1#0"; "vbalTreeView6.ocx"
Begin VB.Form frmSave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Save all documents"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   Icon            =   "FrmSave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   5
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdCloseAll 
      Caption         =   "Close All"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdSaveSel 
      Caption         =   "Save selected"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
   End
   Begin vbalTreeViewLib6.vbalTreeView DocList 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2778
      CheckBoxes      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Modified open documents to be saved:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2790
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Changes to unheked documents will be lost."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   3180
   End
End
Attribute VB_Name = "frmSave"
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

Private counter As Integer
Private m_FFFiles() As String 'Stores the path of the files of the File Forms

'Save selected Files
Private Sub cmdSaveSel_Click()
    Dim i As Integer
    Dim ff As IFileForm
    
    For i = 1 To counter
        If DocList.Nodes(i).Checked = True Then
            Set ff = FindFileForm(m_FFFiles(counter))
            SaveFileOfFileForm ff 'save the fileform
        End If
    Next
    
    frmMain.CancelUnload = False
    Unload Me
End Sub

Private Sub cmdCloseAll_Click()
    Dim i As Integer
    Dim ff As IFileForm
    
    For i = 1 To counter
        If DocList.Nodes(i).Checked = True Then
            Set ff = FindFileForm(m_FFFiles(counter))
            ff.CloseW 'Close the fileform
        End If
    Next
    
    frmMain.CancelUnload = False
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    frmMain.CancelUnload = True
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        cmdCancel_Click
    End If
End Sub

Private Sub Form_Load()
    Dim f As Form
    Dim ff As IFileForm
    Dim nod As cTreeViewNode

    frmMain.CancelUnload = True
    'Look for Dirty forms
    counter = 0
    For Each f In Forms
        If TypeOf f Is IFileForm Then
            Set ff = f
            If ff.IsDirty Then
                counter = counter + 1
                ReDim Preserve m_FFFiles(counter) As String
                m_FFFiles(counter) = ff.FilePath
                Set nod = DocList.Nodes.Add(, etvwChild, CStr(counter), ff.Title _
                & IIf(ff.FilePath <> "", "(" & ff.FilePath & ")", ""))
                nod.Checked = True
            End If
        End If
    Next
End Sub


