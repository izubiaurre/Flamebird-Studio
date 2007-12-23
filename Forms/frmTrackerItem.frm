VERSION 5.00
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.5#0"; "vbalsgrid6.ocx"
Object = "{9DC93C3A-4153-440A-88A7-A10AEDA3BAAA}#3.5#0"; "vbaldtab6.ocx"
Begin VB.Form frmTrackerItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tracker item"
   ClientHeight    =   6180
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8160
   Icon            =   "frmTrackerItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmTrackerItem.frx":058A
   ScaleHeight     =   6180
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame grbButtons 
      BorderStyle     =   0  'None
      Caption         =   "Buttons"
      Height          =   375
      Left            =   5850
      TabIndex        =   16
      Top             =   5760
      Width           =   2295
      Begin VB.CommandButton cmdOk 
         Caption         =   "&OK"
         Height          =   375
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   1200
         TabIndex        =   17
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Frame grbMonitoring 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   6480
      Width           =   7815
      Begin vbAcceleratorSGrid6.vbalGrid vbalGrid1 
         Height          =   1935
         Left            =   360
         TabIndex        =   19
         Top             =   120
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   3413
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Header          =   0   'False
         BorderStyle     =   2
         DisableIcons    =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Comments:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   780
      End
   End
   Begin VB.Frame grbGeneral 
      BorderStyle     =   0  'None
      Caption         =   "General"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   7935
      Begin VB.CheckBox chkHidden 
         Caption         =   "Hidden"
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
         Left            =   6120
         TabIndex        =   30
         Top             =   2250
         Width           =   975
      End
      Begin VB.CheckBox chkLocked 
         Caption         =   "Locked"
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
         Left            =   5040
         TabIndex        =   29
         Top             =   2280
         Width           =   975
      End
      Begin VB.CheckBox chkClosed 
         Caption         =   "Mark to close this item"
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
         Top             =   3600
         Width           =   2295
      End
      Begin VB.HScrollBar hs 
         Height          =   135
         LargeChange     =   10
         Left            =   1455
         Max             =   100
         TabIndex        =   26
         Top             =   3270
         Width           =   6255
      End
      Begin VB.ComboBox cboSubmittedBy 
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
         TabIndex        =   23
         Top             =   2760
         Width           =   2295
      End
      Begin VB.ComboBox cboPriority 
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
         ItemData        =   "frmTrackerItem.frx":0894
         Left            =   5040
         List            =   "frmTrackerItem.frx":08A7
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1200
         Width           =   2295
      End
      Begin VB.ComboBox cboModule 
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
         TabIndex        =   14
         Top             =   600
         Width           =   2295
      End
      Begin VB.ComboBox cboCategory 
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
         TabIndex        =   13
         Top             =   0
         Width           =   2295
      End
      Begin VB.ComboBox cboAssignedTo 
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
         Top             =   1785
         Width           =   2295
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
         Height          =   2175
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   7
         Top             =   840
         Width           =   3615
      End
      Begin VB.TextBox txtSummary 
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
         Left            =   840
         TabIndex        =   5
         Top             =   0
         Width           =   2775
      End
      Begin VB.Image imgEditDevelopers 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Index           =   1
         Left            =   7440
         MouseIcon       =   "frmTrackerItem.frx":08E4
         MousePointer    =   99  'Custom
         Picture         =   "frmTrackerItem.frx":0A50
         Top             =   2760
         Width           =   270
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Closed:"
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
         Left            =   5520
         TabIndex        =   36
         Top             =   3960
         Width           =   540
      End
      Begin VB.Label lblClosed 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "14-may-2004; 00:00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   6165
         TabIndex        =   35
         Top             =   3960
         Width           =   1680
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Created:"
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
         TabIndex        =   34
         Top             =   3960
         Width           =   615
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Modified:"
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
         Left            =   2760
         TabIndex        =   33
         Top             =   3960
         Width           =   645
      End
      Begin VB.Label lblCreated 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "14-may-2004; 00:00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   765
         TabIndex        =   32
         Top             =   3960
         Width           =   1680
      End
      Begin VB.Label lblModified 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "14-may-2004; 00:00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3525
         TabIndex        =   31
         Top             =   3960
         Width           =   1680
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
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
         Left            =   3960
         TabIndex        =   28
         Top             =   2280
         Width           =   510
      End
      Begin VB.Image imgEditDevelopers 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Index           =   0
         Left            =   7440
         MouseIcon       =   "frmTrackerItem.frx":0D94
         MousePointer    =   99  'Custom
         Picture         =   "frmTrackerItem.frx":0F00
         Top             =   1800
         Width           =   270
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Com&pleted:"
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
         TabIndex        =   25
         Top             =   3240
         Width           =   795
      End
      Begin VB.Label lblPercent 
         Alignment       =   1  'Right Justify
         Caption         =   "0%"
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
         Left            =   840
         TabIndex        =   24
         Top             =   3240
         Width           =   420
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Submitted by:"
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
         Left            =   3960
         TabIndex        =   22
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "&Category:"
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
         Left            =   3990
         TabIndex        =   12
         Top             =   60
         Width           =   705
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "&Module:"
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
         Left            =   3990
         TabIndex        =   11
         Top             =   600
         Width           =   555
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "&Priority:"
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
         Left            =   3990
         TabIndex        =   10
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "&Assigned To:"
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
         Left            =   3990
         TabIndex        =   8
         Top             =   1800
         Width           =   960
      End
      Begin VB.Image imgEditCategory 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   7440
         MouseIcon       =   "frmTrackerItem.frx":1244
         MousePointer    =   99  'Custom
         Picture         =   "frmTrackerItem.frx":13B0
         Top             =   15
         Width           =   270
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "&Detailed Description::"
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
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&Summary:"
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
         TabIndex        =   4
         Top             =   45
         Width           =   720
      End
   End
   Begin vbalDTab6.vbalDTabControl tabItems 
      Height          =   4845
      Left            =   0
      TabIndex        =   20
      Top             =   720
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   8546
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
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmTrackerItem.frx":16F4
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
      Left            =   480
      TabIndex        =   21
      Top             =   240
      Width           =   7455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Adding / Editing Tracker items"
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
      Width           =   2490
   End
   Begin VB.Image Image1 
      Height          =   765
      Left            =   0
      Picture         =   "frmTrackerItem.frx":17A4
      Top             =   0
      Width           =   8835
   End
End
Attribute VB_Name = "frmTrackerItem"
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

Private item As cTrackerItem
Private f As frmTodoList
Private devcol As cDeveloperCollection

Public bIsNew As Boolean

Private Sub RefreshModules()
    Dim s As Variant
    cboModule.Clear
    'cboModule.AddItem "<None>"
    For Each s In openedProject.Files
        cboModule.AddItem CStr(s)
    Next
End Sub

Private Sub RefreshCategories()
    Dim s As Variant
    cboCategory.Clear
    cboCategory.AddItem "<None>"
    For Each s In f.AT.CategoryCol
        cboCategory.AddItem CStr(s)
    Next
End Sub

Private Sub RefreshDevLists()
    Dim dev As cDeveloper
    
    cboAssignedTo.Clear
    cboSubmittedBy.Clear
    cboAssignedTo.AddItem "<None>"
    cboSubmittedBy.AddItem "<None>"
    For Each dev In devcol
        cboAssignedTo.AddItem dev.name
        cboSubmittedBy.AddItem dev.name
    Next
End Sub

'Fill the controls with the data of the item object
Private Sub LoadItem()
    Dim i As Integer, bCatFound As Boolean
    
    txtSummary = item.Summary
    txtDescription = item.DetailedDesc
    
    'Check if the category exists (it may have been deleted)
    bCatFound = False
    For i = 0 To cboCategory.ListCount - 1
        If cboCategory.List(i) = item.Category Then bCatFound = True: Exit For
    Next
    If bCatFound = False Then item.Category = ""
    
    cboCategory.text = IIf(item.Category = "", "<None>", item.Category)
    cboModule.text = item.module
    cboPriority.ListIndex = 5 - item.Priority
    
    'Check if the AssignedTo and the SubmittedBy have a valid developer
    If Not (openedProject.devcol.IDForName(item.AssignedTo) > 0) Then item.AssignedTo = ""
    If Not (openedProject.devcol.IDForName(item.SubmittedBy) > 0) Then item.SubmittedBy = ""
    
    cboAssignedTo.text = IIf(item.AssignedTo = "", "<None>", item.AssignedTo)
    cboSubmittedBy.text = IIf(item.SubmittedBy = "", "<None>", item.SubmittedBy)
    chkClosed.value = IIf(item.Closed, 1, 0)
    chkHidden.value = IIf(item.Hidden, 1, 0)
    chkLocked.value = IIf(item.Locked, 1, 0)
    hs.value = item.Completed
    lblCreated = format(item.DateCreated, "ddddd at ttttt")
    lblModified = format(item.DateModified, "ddddd at ttttt")
    lblClosed = format(item.DateClosing, "ddddd at ttttt")
End Sub

'Fill the item object based on the contents of the controls
Private Sub SaveItem()
    item.Summary = txtSummary
    item.DetailedDesc = txtDescription
    item.Category = IIf(cboCategory.text = "<None>", "", cboCategory.text)
    item.module = cboModule.text
    item.Priority = 5 - cboPriority.ListIndex
    item.AssignedTo = IIf(cboAssignedTo.text = "<None>", "", cboAssignedTo.text)
    item.SubmittedBy = IIf(cboSubmittedBy.text = "<None>", "", cboSubmittedBy.text)
    item.Completed = hs.value
    item.Closed = IIf(chkClosed.value = 0, False, True)
    item.Hidden = IIf(chkHidden.value = 0, False, True)
    item.Locked = IIf(chkLocked.value = 0, False, True)
End Sub


Private Sub cmdCancel_Click()
    If bIsNew Then f.AT.Remove (Hex(item.id))
    Unload Me
End Sub

Private Sub cmdOk_Click()
    SaveItem
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Height = 6555
    
    tabItems.Tabs.Add Caption:="General"
    'tabItems.Tabs.Add Caption:="Monitoring"
    'tabItems.Tabs.Add Caption:="Follow-up"
    
    Set f = frmTodoList
    
    Set devcol = openedProject.devcol
    RefreshDevLists
    RefreshCategories
    RefreshModules
    
    Set item = f.ai
    LoadItem 'Fill the controls
    
'    lblDateModified = Format(Date, "Medium Date") & "; " & Format(Time, "hh:mm:ss")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not UnloadMode = vbFormCode Then If bIsNew Then f.AT.Remove (Hex(item.id))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    f.Update
    Set item = Nothing
    Set devcol = Nothing
    Set f = Nothing
    bIsNew = False
End Sub

Private Sub hs_Change()
    lblPercent = CStr(hs.value) & "%"
End Sub

Private Sub hs_Scroll()
    lblPercent = CStr(hs.value) & "%"
End Sub

Private Sub imgEditCategory_Click()
    frmCategoriesEditor.Show 1
    RefreshCategories
End Sub

Private Sub imgEditDevelopers_Click(Index As Integer)
    frmDevelopersList.Show 1
    RefreshDevLists
End Sub
