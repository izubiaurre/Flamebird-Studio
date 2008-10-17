VERSION 5.00
Begin VB.Form frmCategoriesEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Categories"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2475
   Icon            =   "frmListEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   2475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   3
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   2280
      Width           =   735
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   2280
      Width           =   735
   End
   Begin VB.ListBox lst 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      ItemData        =   "frmListEditor.frx":000C
      Left            =   0
      List            =   "frmListEditor.frx":000E
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label lblListName 
      Alignment       =   2  'Center
      Caption         =   "Categories"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmCategoriesEditor"
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
Dim AT As cTracker

Private Sub cmdDelete_Click()
    If lst.ListIndex >= 0 Then
        lst.RemoveItem lst.ListIndex
    End If
End Sub

Private Sub cmdEdit_Click()
    If lst.ListIndex >= 0 Then
        Dim text As String
        text = InputBox("Write the new name of the element", , lst.text)
        If text <> "" Then lst.List(lst.ListIndex) = text
    End If
End Sub

Private Sub cmdNew_Click()
    Dim text As String
    text = InputBox("Write the name of the new element")
    If text <> "" Then lst.AddItem text
End Sub

Private Sub Form_Load()
    Set AT = frmTodoList.AT
    Dim s As Variant
    For Each s In AT.CategoryCol
        lst.AddItem CStr(s)
    Next
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    AT.CategoryClear
    Dim i As Integer
    For i = 0 To lst.ListCount - 1
        AT.AddCategory lst.List(i)
    Next
    Set AT = Nothing
End Sub
