VERSION 5.00
Object = "{9DC93C3A-4153-440A-88A7-A10AEDA3BAAA}#3.5#0"; "vbalDTab6.ocx"
Begin VB.Form frmTrackerManager 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tracker Manager"
   ClientHeight    =   10185
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8520
   Icon            =   "frmEditTracker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10185
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame grbColumns 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   240
      TabIndex        =   34
      Top             =   7680
      Visible         =   0   'False
      Width           =   7455
      Begin VB.CommandButton Command4 
         Height          =   460
         Left            =   3840
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmEditTracker.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1590
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command3 
         Height          =   460
         Left            =   3840
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmEditTracker.frx":0BE4
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1140
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Height          =   460
         Left            =   3840
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmEditTracker.frx":123E
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   690
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   460
         Left            =   3840
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmEditTracker.frx":1898
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   240
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ListBox lstVC 
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
         Left            =   0
         Style           =   1  'Checkbox
         TabIndex        =   36
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Visible &columns:"
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
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   1170
      End
   End
   Begin VB.Frame grbBehavior 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   240
      TabIndex        =   24
      Top             =   5160
      Visible         =   0   'False
      Width           =   7455
      Begin VB.CheckBox chkHideClosedItems 
         Caption         =   "Mark closed items as 'Hidden' after"
         Enabled         =   0   'False
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
         TabIndex        =   33
         Top             =   1815
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         MaxLength       =   3
         TabIndex        =   31
         Text            =   "7"
         Top             =   1800
         Width           =   375
      End
      Begin VB.TextBox txtOldItemsDays 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         MaxLength       =   3
         TabIndex        =   29
         Text            =   "30"
         Top             =   1320
         Width           =   375
      End
      Begin VB.CheckBox chkMarkOldItems 
         Caption         =   "Mark in red the date of items older than"
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
         TabIndex        =   28
         Top             =   1335
         Width           =   3135
      End
      Begin VB.CheckBox chkColorByPriority 
         Caption         =   "Color items by priority"
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
         TabIndex        =   27
         Top             =   840
         Width           =   3495
      End
      Begin VB.CheckBox chkAutoexpand 
         Caption         =   "Autoexpand selected items"
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
         Left            =   360
         TabIndex        =   26
         Top             =   360
         Width           =   2535
      End
      Begin VB.CheckBox chkShowDesc 
         Caption         =   "Show detailed description."
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
         TabIndex        =   25
         Top             =   0
         Width           =   2415
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "days of its closing date."
         Enabled         =   0   'False
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
         Left            =   3840
         TabIndex        =   32
         Top             =   1845
         Width           =   1725
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "days."
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
         Left            =   3840
         TabIndex        =   30
         Top             =   1365
         Width           =   405
      End
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
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
      Left            =   2400
      TabIndex        =   12
      Top             =   3870
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
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
      Left            =   1200
      TabIndex        =   11
      Top             =   3870
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
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
      Left            =   120
      TabIndex        =   10
      Top             =   3870
      Width           =   975
   End
   Begin VB.ComboBox cboTrackers 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5040
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   3840
      Width           =   2655
   End
   Begin VB.Frame grbGeneral 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2175
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   7455
      Begin VB.PictureBox picIcons 
         Height          =   375
         Left            =   480
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   461
         TabIndex        =   13
         Top             =   1680
         Width           =   6975
         Begin VB.CommandButton cmdMore 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   6720
            TabIndex        =   15
            Top             =   -120
            Width           =   255
         End
         Begin VB.CommandButton cmdLess 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   6480
            TabIndex        =   16
            Top             =   -120
            Width           =   255
         End
         Begin VB.Shape focus 
            BorderColor     =   &H00800000&
            Height          =   240
            Left            =   120
            Shape           =   1  'Square
            Top             =   0
            Width           =   240
         End
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   5
         Top             =   0
         Width           =   2415
      End
      Begin VB.TextBox txtDescription 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   0
         TabIndex        =   4
         Top             =   720
         Width           =   7455
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Icon:"
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
         Left            =   0
         TabIndex        =   14
         Top             =   1770
         Width           =   345
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         Caption         =   "&Name:"
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
         Left            =   0
         TabIndex        =   7
         Top             =   45
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "&Description:"
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
         Left            =   0
         TabIndex        =   6
         Top             =   480
         Width           =   855
      End
   End
   Begin vbalDTab6.vbalDTabControl tabConfig 
      Height          =   2895
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5106
      AllowScroll     =   0   'False
      TabAlign        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SelectedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowCloseButton =   0   'False
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   -120
      TabIndex        =   17
      Top             =   4200
      Width           =   8055
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
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
         Left            =   5520
         TabIndex        =   20
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "&Apply"
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
         Left            =   6720
         TabIndex        =   19
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "&Ok"
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
         Left            =   4320
         TabIndex        =   18
         Top             =   180
         Width           =   1095
      End
   End
   Begin VB.Frame grbNoTrackers 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   120
      TabIndex        =   21
      Top             =   1170
      Visible         =   0   'False
      Width           =   7455
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "You haven't added any tracker yet"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2295
         TabIndex        =   23
         Top             =   750
         Width           =   2955
      End
      Begin VB.Label Label1 
         Caption         =   "Use the add button bellow to add a new tracker"
         Height          =   555
         Left            =   2055
         TabIndex        =   22
         Top             =   1110
         Width           =   3375
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Selected tracker:"
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
      Left            =   3480
      TabIndex        =   8
      Top             =   3900
      Width           =   1230
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmEditTracker.frx":1EF2
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
      Left            =   480
      TabIndex        =   1
      Top             =   225
      Width           =   7605
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trackers"
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
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   750
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   765
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8835
   End
End
Attribute VB_Name = "frmTrackerManager"
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
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

Private Const MARGINX As Integer = 18
Private Const MARGINY As Integer = 2


Private f As frmTodoList 'Reference to the frmTodoList
Private colTrackers As cTrackerCollection, colOriginal As cTrackerCollection
Private AT As cTracker, ilst As vbalImageList 'Ilst is a reference to the image list in the main form

Private clickX, clickY, n As Integer 'For the focus square
Private bNoTrackers As Boolean

Private Property Let NoTrackers(newBool As Boolean)
    bNoTrackers = newBool
    If newBool = True Then 'No trackers
        'Disable some controls and Show the grbNotrackers
        tabConfig.ShowTabs = False
        cmdDelete.Enabled = False
        cmdCopy.Enabled = False
        grbNoTrackers.ZOrder 0 'Move to the foreground
        grbNoTrackers.Visible = True
    Else
        'Enable some controls and hide the grdbnotrackers
        tabConfig.ShowTabs = True
        cmdDelete.Enabled = True
        cmdCopy.Enabled = True
        grbNoTrackers.ZOrder 1 'Move to the background
        grbNoTrackers.Visible = False
    End If
End Property

Private Property Get NoTrackers() As Boolean
    NoTrackers = bNoTrackers
End Property

Private Sub ApplyChanges()
    Dim ot As cTracker 'Original tracker
    Dim mt As cTracker 'Modified tracker
    
    colTrackers.CopyIn colOriginal
        
    Set ot = Nothing
    Set mt = Nothing
End Sub

Private Sub SelectTracker(sKey As String)
    Dim i As Integer
    
    Set AT = colTrackers(sKey)
    
    With AT
        'general
        txtName = .name
        txtDescription = .description
        GetIconsFromIlst picIcons
        focus.Move MARGINX + .IconIndex * 16, MARGINY
        'behavior
        If .AutoExpandSelItems = True Then .ShowDescription = True
        chkShowDesc = Abs(CInt(AT.ShowDescription))
        chkAutoexpand = Abs(.AutoExpandSelItems)
        chkColorByPriority = Abs(.ColorItemsByPriority)
        chkMarkOldItems = Abs(.ColorOldItems)
        txtOldItemsDays = .OldItemsDays
        'Columns
        lstVC.Clear
        For i = 1 To f.grdTracker.Columns - 1
            lstVC.AddItem f.grdTracker.ColumnKey(i)
            lstVC.ItemData(i - 1) = CLng(f.grdTracker.ColumnTag(i))
            If AT.ColumnVisible(CLng(f.grdTracker.ColumnTag(i))) Then lstVC.Selected(i - 1) = True
        Next
    End With
End Sub

Private Sub GetIconsFromIlst(pBox As PictureBox)
    Dim hdc As Long
    Dim picIcons As New StdPicture, oldObject As Long
    
    Set picIcons = ilst.ImagePictureStrip(1, ilst.ImageCount)
    hdc = CreateCompatibleDC(0)
    oldObject = SelectObject(hdc, picIcons.Handle)
    
    pBox.AutoRedraw = True
    BitBlt pBox.hdc, 0, 0, picIcons.Height, picIcons.Width, hdc, -18, -MARGINY, vbSrcCopy
    pBox.Picture = pBox.Image
    pBox.AutoRedraw = False
    
    ' restore object from HDC and delete the HDC
    SelectObject hdc, oldObject
    DeleteDC hdc
    Set picIcons = Nothing
End Sub

'Fill the combo with the trackers and return false if there isn't any trackers
Private Function FillCombo() As Boolean
    Dim tr As cTracker
    
    With cboTrackers
        .Clear
        For Each tr In colTrackers
            .AddItem tr.name
        Next
        FillCombo = CBool(cboTrackers.ListCount)
    End With
End Function

'Selects the text of a textbox
Private Sub MakeSel(Control As TextBox)
    Control.SetFocus:    Control.SelStart = 0:    Control.SelLength = Len(Control.text)
End Sub

Private Sub cboTrackers_Click()
    SelectTracker colTrackers.KeyForName(cboTrackers.List(cboTrackers.ListIndex))
End Sub

Private Sub chkAutoexpand_Click()
    AT.AutoExpandSelItems = CBool(chkAutoexpand.Value)
End Sub

Private Sub chkColorByPriority_Click()
    AT.ColorItemsByPriority = CBool(chkColorByPriority.Value)
End Sub

Private Sub chkMarkOldItems_Click()
    AT.ColorOldItems = CBool(chkMarkOldItems)
    txtOldItemsDays.Enabled = CBool(chkMarkOldItems)
End Sub

Private Sub chkShowDesc_Click()
    If chkShowDesc = 1 Then
        AT.ShowDescription = True
        chkAutoexpand.Enabled = True
    Else
        AT.ShowDescription = False
        chkAutoexpand = 0
        chkAutoexpand.Enabled = False
    End If
End Sub

Private Sub cmdAdd_Click()
    Dim sName As String, newID As Integer
    Dim i As Integer
    
    'Get a valid name for the new tracker
    i = 1
    Do
        If colTrackers.IndexForName("<New Tracker> " & CStr(i)) = 0 Then
            sName = "<New Tracker> " & CStr(i)
        Else
            i = i + 1
        End If
    Loop Until sName <> ""
    
    colTrackers.Add sName 'Add the new tracker to the col width the default properties
    
    'Refresh the combo and select the new tracker
    FillCombo
    cboTrackers.text = sName
    MakeSel txtName
    
    NoTrackers = False
End Sub

Private Sub cmdApply_Click()
    ApplyChanges
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDelete_Click()
    Dim prevItem As Integer
    
    'Get the previous item of the selected
    prevItem = cboTrackers.ListIndex - 1
    If prevItem < 0 Then prevItem = 0
    
    colTrackers.Remove AT.Key
    FillCombo
    
    If colTrackers.count = 0 Then
        NoTrackers = True
    Else
        cboTrackers.ListIndex = prevItem
    End If
End Sub

Private Sub cmdLess_Gotfocus()
    picIcons.SetFocus
End Sub

Private Sub cmdMore_Gotfocus()
    picIcons.SetFocus
End Sub


Private Sub cmdOk_Click()
    ApplyChanges
    Unload Me
End Sub

Private Sub Form_Load()

    Image1.Picture = LoadPicture(App.Path & "/Resources/frmHeader.jpg")

    Set f = frmTodoList
    'Set colTrackers = f.colTrackers.Copy
    'Set colOriginal = f.colTrackers
    Set colTrackers = openedProject.colTrackers.Copy
    Set colOriginal = openedProject.colTrackers
    Set ilst = f.ilstTabs
    
    'Set the appropiated dimensions of the form and some controls
    Me.Width = 7800: Me.Height = 5160
    grbBehavior.Move grbGeneral.Left, grbGeneral.Top
    grbColumns.Move grbGeneral.Left, grbGeneral.Top
    
    'Tabs
    tabConfig.Tabs.Add Caption:="General"
    tabConfig.Tabs.Add Caption:="Behavior"
    tabConfig.Tabs.Add Caption:="Columns"
    
    If FillCombo Then ' if exists trackers
        cboTrackers.ListIndex = 0
        NoTrackers = False
        'Select the appropiate tracker
        cboTrackers.text = f.AT.name
    Else
        NoTrackers = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    f.Update
    Set f = Nothing
    Set ilst = Nothing
    Set colTrackers = Nothing
    Set colOriginal = Nothing
    Set AT = Nothing
End Sub

Private Sub lstVC_Click()
    AT.ColumnVisible(lstVC.ItemData(lstVC.ListIndex)) = lstVC.Selected(lstVC.ListIndex)
End Sub

Private Sub picIcons_Click()
    n = (clickX - MARGINX) \ 16
    If (n + 1) > ilst.ImageCount Then Exit Sub
    If (clickX - MARGINX) < 0 Then n = -1 'No icon
    focus.Move MARGINX + n * 16
    AT.IconIndex = n 'Change
End Sub

Private Sub picIcons_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clickX = X
    clickY = Y
End Sub

Private Sub tabConfig_TabClick(theTab As vbalDTab6.cTab, ByVal iButton As MouseButtonConstants, ByVal Shift As ShiftConstants, ByVal X As Single, ByVal Y As Single)
    grbGeneral.Visible = False
    grbBehavior.Visible = False
    grbColumns.Visible = False
    
    Select Case theTab.Caption
    Case "General"
        grbGeneral.Visible = True
    Case "Behavior"
        grbBehavior.Visible = True
    Case "Columns"
        grbColumns.Visible = True
    End Select
End Sub

Private Sub txtDescription_LostFocus()
    AT.description = txtDescription 'Change
End Sub

Private Sub txtName_LostFocus()
    Dim tID As Long
    
    'Check for trackers with the same name
    tID = colTrackers.IndexForName(txtName)
    If tID And tID <> AT.Id Then
        MsgBox "There is another tracker with the name '" & txtName & "'", vbOKOnly + vbCritical
       ' Cancel = True
        txtName.text = AT.name
        MakeSel txtName
    Else
        AT.name = txtName 'Change
        FillCombo
        cboTrackers.text = AT.name
    End If
End Sub

Private Sub txtOldItemsDays_LostFocus()
    Dim bError As Boolean
    If IsNumeric(txtOldItemsDays) Then
        If CInt(txtOldItemsDays) > 0 Then
            AT.OldItemsDays = CInt(txtOldItemsDays)
        Else
            bError = True
        End If
    Else
        bError = True
    End If
    
    If bError Then 'Invalid
        MsgBox "Please, enter a numeric value between 0 and 999 ", vbCritical, "Tracker Manager"
        txtOldItemsDays.text = AT.OldItemsDays
        MakeSel txtOldItemsDays
    End If
End Sub
