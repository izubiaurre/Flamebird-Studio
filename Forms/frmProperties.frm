VERSION 5.00
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.5#0"; "vbalSGrid6.ocx"
Object = "{5ABC9E42-2956-4D74-82BD-044D57BB671A}#1.0#0"; "cssplit.ocx"
Begin VB.Form frmProperties 
   Caption         =   "Properties"
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3600
   Icon            =   "frmProperties.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin CSSplitter.TSplitter HSplitter 
      Height          =   75
      Left            =   0
      Top             =   1920
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   132
      BorderStyle     =   2
      Orientation     =   2
      MousePointer    =   7
   End
   Begin VB.PictureBox picDesc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   3375
      TabIndex        =   8
      Top             =   2160
      Width           =   3375
      Begin VB.Label lblDesc 
         Caption         =   "Description"
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
         Left            =   30
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblPropName 
         Caption         =   "PropertyName"
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
         Left            =   30
         TabIndex        =   9
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.PictureBox cmdLink 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   2415
      TabIndex        =   4
      Top             =   3240
      Visible         =   0   'False
      Width           =   2415
      Begin VB.PictureBox picLink 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1320
         ScaleHeight     =   375
         ScaleWidth      =   615
         TabIndex        =   5
         Top             =   120
         Width           =   615
         Begin VB.Label lblLink 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00D6C3BC&
            Caption         =   "..."
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   255
         End
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
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
         TabIndex        =   7
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.PictureBox picEditBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   2655
      TabIndex        =   2
      Top             =   3960
      Visible         =   0   'False
      Width           =   2655
      Begin VB.TextBox txtEditBox 
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         Height          =   195
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.ComboBox cboOption 
      Appearance      =   0  'Flat
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
      ItemData        =   "frmProperties.frx":058A
      Left            =   0
      List            =   "frmProperties.frx":0594
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   2880
      Visible         =   0   'False
      Width           =   1575
   End
   Begin vbAcceleratorSGrid6.vbalGrid grdProp 
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   3413
      GridLines       =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Header          =   0   'False
      HeaderButtons   =   0   'False
      BorderStyle     =   2
      ScrollBarStyle  =   1
      DisableIcons    =   -1  'True
      HotTrack        =   -1  'True
      SelectionAlphaBlend=   -1  'True
   End
End
Attribute VB_Name = "frmProperties"
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
Private Const CLR_FLAT_BUTTON As Long = &HD6C3BC
Private Const CLR_FLAT_BUTTON_PRESSED As Long = &HB59386
Private Const MIN_PICDESC_HEIGHT As Integer = 555
Private Const FORM_SPLIT_DIST As Integer = 100 ' Min separation between form and splitter

Implements ITDockMoveEvents

Private CurHeight As Integer ' The height set by the user
                             ' To that increasing the size of the form
                             ' is not exceeded
Private props As cProperties
Private m_cFlat As cFlatControl

Public Sub ClearProperties()
    props.Clear
End Sub
Public Sub AddProperty(ByVal Caption As String, ByVal Key As String, ByVal TypeOfProp As PropertyType, _
                    ByRef CallingObject As Object, ByVal CallBackFunction As String, ByVal Value As String, _
                    ByVal Editable As Boolean, Optional max As Integer, _
                    Optional min As Integer = 0, Optional CanBeEmpty As Boolean = False)
    props.Add Caption, Key, TypeOfProp, CallingObject, CallBackFunction, Value, Editable, max, min, CanBeEmpty
End Sub

Public Sub AddPropertyOption(ByVal Key As String, OptionName As String)
    props(Key).AddOption OptionName
End Sub

Public Sub AddPropertyDescription(ByVal Key As String, description As String)
    props(Key).description = description
End Sub

Private Sub cboOption_KeyPress(KeyAscii As Integer)
    'Accept
    If KeyAscii = 13 Then KeyAscii = 0: Call grdProp.EndEdit
    'Scape
    If KeyAscii = 27 Then KeyAscii = 0: Call grdProp.canceledit
End Sub

Private Sub Form_Load()
    Set props = New cProperties
    Set m_cFlat = New cFlatControl

    ' Sets special apparence of picProperties
    Dim PictureStyle As Long
    PictureStyle = GetWindowLong(picDesc.Hwnd, GWL_EXSTYLE)
    PictureStyle = PictureStyle Or WS_EX_STATICEDGE
    SetWindowLong picDesc.Hwnd, GWL_EXSTYLE, PictureStyle
    picDesc.Refresh
    
    ' Conects the controls with the splitter
    CurHeight = 800
    lblPropName.Caption = ""
    lblDesc.Caption = ""
    
    With grdProp
        .AddColumn: .AddColumn
        .DefaultRowHeight = ScaleY(cboOption.Height, 1, 3) ' The height must be of the combo
                                                           ' cause this can¡t be modified
        .StretchLastColumnToFit = True
        .Editable = True
        .RowMode = True     ' Selection by rows
        .HighlightBackColor = QBColor(7)
        .HighlightForeColor = QBColor(0)
    End With

    m_cFlat.OnFocusedRectColor = grdProp.GridLineColor  ' To be ok
    
    RefreshGrid
End Sub

Private Sub RefreshGrid()
    Dim Prop As Variant
    Dim cnt As Integer
    cnt = 1
    grdProp.Clear False     ' Delete the rows
    For Each Prop In props  ' For each property in the collection
        With grdProp
            .AddRow
            .CellText(cnt, 1) = Prop.name
            If Prop.TypeOfProp <> ptCombo Then
                .CellText(cnt, 2) = Prop.Value
            Else ' If it's a combo, value points to index, not to the text
                .CellText(cnt, 2) = Prop.OptionItem(Prop.Value)
            End If
            cnt = cnt + 1
        End With
    Next
    lblPropName.Caption = ""
    lblDesc.Caption = ""
End Sub

Public Sub RefreshProperties()
    Dim pf As IPropertiesForm
    
    props.Clear
    If Not frmMain.ActiveForm Is Nothing Then
        If TypeOf frmMain.ActiveForm Is IPropertiesForm Then
            Set pf = frmMain.ActiveForm
            Set props = pf.GetProperties
        End If
    End If
    
    If Not props Is Nothing Then
        RefreshGrid
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set props = Nothing
    Set m_cFlat = Nothing
End Sub

Private Sub grdProp_CancelEdit()
    cboOption.Visible = False
    picEditBox.Visible = False
    cmdLink.Visible = False
End Sub

Private Sub grdProp_PreCancelEdit(ByVal lRow As Long, ByVal lCol As Long, newValue As Variant, bStayInEditMode As Boolean)
    With props(lRow)
    Select Case .TypeOfProp
    Case ptCombo
        newValue = CInt(cboOption.ListIndex)
        ' We call the function callback, giving the new index and the associated text
        If (CallByName(.CallingObject, .CallBackFunction, VbMethod, newValue)) Then   ', .OptionItem(newValue))
            grdProp.CellText(lRow, lCol) = .OptionItem(newValue)
            .Value = newValue
        End If
    Case ptLink
        ' This proptery-type makes nothing, the work is managed
        ' by the function linked to the property
        Call CallByName(.CallingObject, .CallBackFunction, VbMethod, newValue)
    Case Else
        If .CanBeEmpty = False And Len(txtEditBox) = 0 Then 'Text=""
            MsgBox "'" & .name & "' can't be empty", vbCritical
            GoTo canceledit
        ElseIf .CanBeEmpty = True Or Len(txtEditBox) <> 0 Then
            ' Data validation
            If .TypeOfProp = ptNumeric Or .TypeOfProp = ptInteger Then ' It's number
                If IsNumeric(txtEditBox) Then
                    ' Limit values
                    If .IsLimited Then
                        If txtEditBox > .max Or txtEditBox < .min Then
                            MsgBox "'" & .name & "' must be between '" & .min & "' and '" & .max & "'", vbCritical
                            GoTo canceledit
                        End If
                    End If
                Else
                    MsgBox "'" & txtEditBox & "' is not a numeric value", vbCritical
                    GoTo canceledit
                End If
            End If
        End If
        newValue = txtEditBox
        If .TypeOfProp = ptInteger And newValue <> "" Then newValue = CInt(txtEditBox) ' If it is integer, converts it
        ' Call function callback, giving the new value
        If CallByName(.CallingObject, .CallBackFunction, VbMethod, newValue) Then
            grdProp.CellText(lRow, lCol) = newValue
            .Value = newValue
        End If
    End Select
    End With
    
    Exit Sub

canceledit:
    bStayInEditMode = True
    txtEditBox = props(lRow).Value
    Exit Sub
End Sub

Private Sub grdProp_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
    
    If props(lRow).Editable = False Or lCol <> 2 Then bCancel = True: Exit Sub ' Can't be edited

    ' gets the cell size
    Dim lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long
    grdProp.CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight
    
    Select Case props(lRow).TypeOfProp
    Case ptCombo 'Combo
        m_cFlat.Attach cboOption ' Flat style
        With cboOption
            ' Set background color for the combobox
            Set .font = grdProp.CellFont(lRow, lCol)
            If grdProp.CellBackColor(lRow, lCol) = -1 Then
                .BackColor = grdProp.BackColor
            Else
                .BackColor = grdProp.CellBackColor(lRow, lCol)
            End If
        
            ' Position of the combo
            .Move lLeft + grdProp.Left - 10, lTop + grdProp.Top + Screen.TwipsPerPixelY, lWidth
        
            ' fills it, shows and gets the focus
            Dim opt As Variant
            .Clear
            For Each opt In props(lRow)
                .AddItem opt
            Next
            .ListIndex = props(lRow).Value
            .Visible = True
            .SetFocus
        End With
    
    Case ptLink
        With cmdLink
            ' Button position
            lblCaption = props(lRow).Value
            .Move lLeft + grdProp.Left, lTop + grdProp.Top + Screen.TwipsPerPixelY + 8, lWidth - 10, lHeight - 30
            .Visible = True
        End With
        
    Case Else 'ptText, ptInteger y ptNumeric
        With txtEditBox
            ' if it is text type, and has char limits, limit the textbox
            If props(lRow).TypeOfProp = ptText And props(lRow).IsLimited = True Then
                .MaxLength = props(lRow).max
            Else
                .MaxLength = 0
            End If
            
            ' Set background color for the textbox and its container
            Set .font = grdProp.CellFont(lRow, lCol)
            If grdProp.CellBackColor(lRow, lCol) = -1 Then
                .BackColor = grdProp.BackColor
                picEditBox.BackColor = .BackColor
            Else
                .BackColor = grdProp.CellBackColor(lRow, lCol)
                picEditBox.BackColor = .BackColor
            End If
            
            ' Set textbox position (its container)
            picEditBox.Move lLeft + grdProp.Left, lTop + grdProp.Top + Screen.TwipsPerPixelY + 8, lWidth - 10, lHeight - 30
            
            ' Shows it, fills it, focuses it and selectes the text
            picEditBox.Visible = True
            picEditBox.ZOrder
            .text = props(lRow).Value ' Initial text
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.text)
        End With
    End Select
    
    ' If is a key pressed, we send it
    If iKeyAscii <> 0 Then SendKeys (Chr(iKeyAscii))
End Sub

Private Sub grdProp_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
    ' put the description
    lblPropName.Caption = props(lRow).name
    lblDesc.Caption = props(lRow).description
End Sub

Private Sub HSplitter_EndMoving()
    Dim ITop As Long, FTop As Long, IHeight As Long, FHeight As Long
    
    FHeight = HSplitter.Top - grdProp.Top
    If FHeight < 0 Then ' We have move up too much the splitter
        HSplitter.Top = grdProp.Top + FORM_SPLIT_DIST
        HSplitter_EndMoving
        Exit Sub
    End If
    grdProp.Height = FHeight

    ITop = picDesc.Top
    IHeight = picDesc.Height
    FTop = HSplitter.Top + HSplitter.Height
    FHeight = (ITop - FTop) + IHeight
    If FHeight < MIN_PICDESC_HEIGHT Then ' We have move down too much the splitter
        HSplitter.Top = (ITop - HSplitter.Height) + IHeight - MIN_PICDESC_HEIGHT
        HSplitter_EndMoving
        Exit Sub
    End If
    picDesc.Top = FTop
    picDesc.Height = FHeight
    CurHeight = FHeight
End Sub

Private Sub lblLink_Click()
    grdProp.canceledit  ' Cancel the edition
    'Function callback
    CallByName props(grdProp.SelectedRow).CallingObject, props(grdProp.SelectedRow).CallBackFunction, VbMethod
End Sub

Private Sub picDesc_Resize()
    ' Fixes lblDescription size
    lblDesc.Width = picDesc.ScaleWidth - lblDesc.Left - 10
    lblDesc.Height = picDesc.ScaleHeight - lblDesc.Top
End Sub

Private Sub picEditBox_Resize()
    txtEditBox.Move 65, (picEditBox.Height - txtEditBox.Height) / 2 - 40, picEditBox.ScaleWidth - 65
End Sub

Private Sub lblLink_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblLink.BackColor = CLR_FLAT_BUTTON_PRESSED
End Sub

Private Sub lblLink_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblLink.BackColor = CLR_FLAT_BUTTON
End Sub

Private Sub picLink_Resize()
    lblLink.Move 0, 0, picLink.ScaleWidth, picLink.ScaleHeight
End Sub

Private Sub cmdLink_Resize()
    With cmdLink
        Dim h As Long
        h = .Height '- 20 ' button height
        picLink.Move .Width - h, (.Height - h) / 2, h, h
        lblCaption.Move 65, (.Height - lblCaption.Height) / 2 - 25
    End With
End Sub

Private Function ITDockMoveEvents_DockChange(tDockAlign As AlignConstants, tDocked As Boolean) As Variant
       
End Function

Private Function ITDockMoveEvents_Move(Left As Integer, Top As Integer, Bottom As Integer, Right As Integer)
On Error Resume Next
    ' If the form is very little, don't resize the controls
    ' (picDesc can't be smaller than than its minimun size)
    If (Bottom - Top) < (MIN_PICDESC_HEIGHT + FORM_SPLIT_DIST) Then
        picDesc.Left = Left: picDesc.Width = Right: grdProp.Left = Left
        grdProp.Width = Right: HSplitter.Left = Left: HSplitter.Width = Width
    Else
        If (Bottom - Top) < CurHeight Then CurHeight = Bottom - Top - FORM_SPLIT_DIST
        picDesc.Move Left, Bottom - CurHeight + Top, Right, CurHeight
        HSplitter.Move Left, picDesc.Top - HSplitter.Height, Right
        grdProp.Move Left, Top, Right, HSplitter.Top
    End If
End Function

Private Sub txtEditBox_KeyPress(KeyAscii As Integer)
    'Accept
    If KeyAscii = 13 Then KeyAscii = 0: Call grdProp.EndEdit
    'Scape
    If KeyAscii = 27 Then KeyAscii = 0: Call grdProp.canceledit
End Sub
