VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbaliml6.ocx"
Object = "{E910F8E1-8996-4EE9-90F1-3E7C64FA9829}#1.1#0"; "vbalistview6.ocx"
Begin VB.Form frmNewFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New file"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   Icon            =   "frmNewFile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUseWizard 
      Caption         =   "Use wizard"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   3600
      Width           =   975
   End
   Begin VB.CheckBox chkAddToProject 
      Appearance      =   0  'Flat
      Caption         =   "Add file to project"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   3600
      Width           =   975
   End
   Begin vbalIml6.vbalImageList imgList 
      Left            =   5400
      Top             =   3000
      _ExtentX        =   953
      _ExtentY        =   953
      IconSizeX       =   32
      IconSizeY       =   32
      ColourDepth     =   8
      Size            =   26472
      Images          =   "frmNewFile.frx":08CA
      Version         =   131072
      KeyCount        =   6
      Keys            =   "PROJECTÿSOURCEÿMAPÿFPGÿPALÿFNT"
   End
   Begin vbalListViewLib6.vbalListViewCtl lstTypes 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   4895
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MultiSelect     =   -1  'True
      LabelEdit       =   0   'False
      AutoArrange     =   0   'False
      BorderStyle     =   2
      CustomDraw      =   0   'False
      FlatScrollBar   =   -1  'True
      HeaderButtons   =   0   'False
      HeaderTrackSelect=   0   'False
      HideSelection   =   0   'False
      InfoTips        =   0   'False
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New file"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   660
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmNewFile.frx":7052
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   5805
   End
   Begin VB.Image Image1 
      Height          =   765
      Left            =   -2160
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8835
   End
End
Attribute VB_Name = "frmNewFile"
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

Private Type FileType
    name As String
    IconKey As String
    Enabled As Boolean
End Type

Private fileTypes(3) As FileType

Private m_Key As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim sFile As String
    Dim tStream As textStream
    
    On Error GoTo errhandler
    
    If Not lstTypes.SelectedItem Is Nothing Then
        m_Key = lstTypes.SelectedItem.Key
        
        modMenuActions.NewAddToProject = IIf(chkAddToProject.Value = 0, False, True)
        modMenuActions.newType = m_Key
    End If
    
    Unload Me
    
errhandler:
    If Err.Number > 0 Then ShowError ("frmNewFile.cmdOk_click()")
End Sub

Private Sub cmdUseWizard_Click()
    Unload Me
    frmCodeWizard.Show vbModal, Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        cmdCancel_Click
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim item As vbalListViewLib6.cListItem
    
    Image1.Picture = LoadPicture(App.Path & "\Resources\frmHeader.jpg")
    
    If Not openedProject Is Nothing Then
        chkAddToProject.Enabled = True
        chkAddToProject.Value = 1
    Else
        chkAddToProject.Value = 0
        chkAddToProject.Enabled = False
    End If
    
    'Define file types
    fileTypes(0).name = "Flamebird Project"
    fileTypes(0).IconKey = "PROJECT"
    fileTypes(1).name = "Bennu Source"
    fileTypes(1).IconKey = "SOURCE"
    fileTypes(2).name = "Bennu Map"
    fileTypes(2).IconKey = "MAP"
    fileTypes(3).name = "Bennu Fpg"
    fileTypes(3).IconKey = "FPG"
'    fileTypes(4).name = "Bennu Palette"
'    fileTypes(4).IconKey = "PAL"
    
    'Configure the list view
    lstTypes.ImageList = imgList
    For i = 0 To UBound(fileTypes)
        Set item = lstTypes.ListItems.Add(, fileTypes(i).IconKey, fileTypes(i).name, imgList.ItemIndex(fileTypes(i).IconKey) - 1)
    Next i

    If openedProject Is Nothing Then
        lstTypes.ListItems.item(1).Selected = True
    Else
        lstTypes.ListItems.item(2).Selected = True
    End If
End Sub


Private Sub lstTypes_ItemClick(item As vbalListViewLib6.cListItem)
    If item.Key = "FBP" = True Then
        chkAddToProject.Enabled = False
    Else
        If Not openedProject Is Nothing Then chkAddToProject.Enabled = True
    End If
    If item.Key = "SOURCE" = True Then
        cmdUseWizard.Enabled = True
    Else
        cmdUseWizard.Enabled = False
    End If
End Sub

Private Sub lstTypes_ItemDblClick(item As vbalListViewLib6.cListItem)
    cmdOk_Click
End Sub
