VERSION 5.00
Begin VB.Form frmBookmarkEditor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Bookmarks"
   ClientHeight    =   3690
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmBookmarkEditor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "List of Bookmarks"
      Height          =   3135
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4455
      Begin VB.TextBox txtNewName 
         Height          =   285
         Left            =   2280
         TabIndex        =   5
         Top             =   720
         Width           =   2055
      End
      Begin VB.ListBox lstBookmark 
         Height          =   2400
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lblNewName 
         Caption         =   "Name of the Bookmark"
         Height          =   255
         Left            =   2280
         TabIndex        =   4
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Accept"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   1215
   End
End
Attribute VB_Name = "frmBookmarkEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fDoc As frmDoc
Dim curBookmark As Long

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
    lstBookmark.RemoveItem curBookmark
    lstBookmark.AddItem txtNewName.text, curBookmark
    fDoc.bookmarkList_name(curBookmark + 2) = txtNewName.text
    fDoc.refreshBookmarkList
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        cmdCancel_Click
    End If
End Sub

Private Sub Form_Load()
    Dim i As Long
    Set fDoc = frmMain.ActiveForm

    For i = 2 To fDoc.getLastBookmarkIndex
        lstBookmark.AddItem fDoc.bookmarkList_name(i)
    Next i
End Sub

Private Sub lstBookmark_Click()
    txtNewName.text = lstBookmark.text
    curBookmark = lstBookmark.ListIndex
    Debug.Print curBookmark
End Sub

Private Sub OKButton_Click()
    Unload Me
End Sub

