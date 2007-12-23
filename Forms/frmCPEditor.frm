VERSION 5.00
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.5#0"; "vbalsgrid6.ocx"
Begin VB.Form frmCPEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control points editor"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   Icon            =   "frmCPEditor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInsertAt 
      Height          =   375
      Left            =   3480
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmCPEditor.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdDiscard 
      Cancel          =   -1  'True
      Caption         =   "&Discard"
      Height          =   375
      Left            =   4560
      TabIndex        =   18
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "&Accept"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   17
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export..."
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "&Import..."
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdInsert 
      Height          =   375
      Left            =   3480
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmCPEditor.frx":06CE
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdMoveUp 
      Height          =   375
      Left            =   3480
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmCPEditor.frx":08A2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2760
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.CommandButton cmdMoveDown 
      Height          =   375
      Left            =   3480
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmCPEditor.frx":0930
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3240
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.TextBox txtEdit 
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
      Height          =   285
      Left            =   600
      TabIndex        =   22
      Top             =   2040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdRemove 
      Height          =   375
      Left            =   3480
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmCPEditor.frx":09BF
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Predefined sets"
      Height          =   1335
      Left            =   4080
      TabIndex        =   21
      Top             =   2280
      Width           =   1335
      Begin VB.CommandButton cmdBottomRight 
         Height          =   375
         Left            =   960
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCPEditor.frx":0B93
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdBottom 
         Height          =   375
         Left            =   480
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCPEditor.frx":0ED7
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdBottomLeft 
         Height          =   375
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCPEditor.frx":121B
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdRight 
         Height          =   375
         Left            =   960
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCPEditor.frx":155F
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdCenter 
         Height          =   375
         Left            =   480
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCPEditor.frx":18A3
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdLeft 
         Height          =   375
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCPEditor.frx":1BE7
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdTopRight 
         Height          =   375
         Left            =   960
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCPEditor.frx":1F2B
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdTop 
         Height          =   375
         Left            =   480
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCPEditor.frx":226F
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton cmdTopLeft 
         Height          =   375
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmCPEditor.frx":25B3
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   375
      Left            =   3480
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmCPEditor.frx":28F7
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   840
      UseMaskColor    =   -1  'True
      Width           =   375
   End
   Begin vbAcceleratorSGrid6.vbalGrid grd 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4895
      GridLines       =   -1  'True
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
      HeaderButtons   =   0   'False
      HeaderDragReorderColumns=   0   'False
      HeaderHeight    =   17
      HeaderFlat      =   -1  'True
      BorderStyle     =   2
      DisableIcons    =   -1  'True
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Predefined sets:"
      Height          =   195
      Left            =   4080
      TabIndex        =   23
      Top             =   2040
      Width           =   1140
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   5400
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line1 
      X1              =   3960
      X2              =   3960
      Y1              =   840
      Y2              =   3600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Control points editor"
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
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   1725
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Here you can add, edit and remove control points."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   19
      Top             =   360
      Width           =   4965
   End
   Begin VB.Image Image1 
      Height          =   765
      Left            =   -2880
      Picture         =   "frmCPEditor.frx":2AE7
      Top             =   0
      Width           =   8835
   End
End
Attribute VB_Name = "frmCPEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Flamebird MX
'Copyright (C) 2003-2007 Flamebird Team
'Contact:
'   JaViS:      javisarias@ gmail.com(JaViS)
'   Danko:      lord_danko@users.sourceforge.net (Dar�o Cutillas)
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
Option Base 0

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Implements IGridCellOwnerDraw

Private Type T_CP
    X As Integer
    Y As Integer
End Type

Private m_map As cMap
Private cpoints(999) As T_CP
Private cpexists(999) As Boolean
Private lastcp As Integer
Private dirty As Boolean

Const MSG_SAVEFILE = "The control-points have been changed. Do you want to save changes?"

Public Sub SelectMap(m As cMap)
    Dim i As Integer
    Dim cp() As Integer
    
    Set m_map = m
    'Initialize to the invalid coord
    Erase cpoints
    Erase cpexists
    lastcp = -1
    If Not m_map Is Nothing Then
        'Fill CP with the map cp data
        For i = 0 To m_map.CPointsCount - 1
            cp() = m_map.ControlPoint(i)
            cpoints(i).X = cp(0)
            cpoints(i).Y = cp(1)
            cpexists(i) = True
            lastcp = lastcp + 1
        Next
    End If
    FillGrid
End Sub

Private Sub AddCp(ByVal bInsert As Boolean)
    
    Dim index As Long
    Dim i As Integer
    
    If lastcp < 999 Then
        With grd
        If bInsert = False Then 'New
            lastcp = lastcp + 1
            cpoints(lastcp).X = 0
            cpoints(lastcp).Y = 0
            cpexists(lastcp) = True
            'Add the row and start editing
            .AddRow .Rows
            .RowItemData(.Rows - 1) = lastcp
            .CellDetails .Rows - 1, 1, CStr(lastcp), DT_RIGHT
            .CellDetails .Rows - 1, 2, CStr(cpoints(lastcp).X), DT_RIGHT
            .CellDetails .Rows - 1, 3, CStr(cpoints(lastcp).Y), DT_RIGHT
            .SelectedRow = .Rows - 1
            .EndEdit
            .StartEdit .Rows - 1, 2
        Else 'Insert
            If .SelectedRow > 0 Then
                Dim tempIndex As Integer
                tempIndex = .CellText(.SelectedRow, 1)
                index = .RowItemData(.SelectedRow)
                'Move cpdata 1 position down
                CopyMemory ByVal VarPtr(cpoints(index + 1)), ByVal VarPtr(cpoints(index)), (999 - index) * 4
                'Move cpexists 1 position down
                CopyMemory ByVal VarPtr(cpexists(index + 1)), ByVal VarPtr(cpexists(index)), (999 - index) * 2
                lastcp = lastcp + 1
                'Assign the new item data
'                For i = .SelectedRow To .Rows - 1
''                    .RowItemData(i) = .RowItemData(i) + 1
''                    .CellDetails i, 1, CStr(.RowItemData(i)), DT_RIGHT
''                    .CellDetails i, 2, CStr(cpoints(.RowItemData(i)).X), DT_RIGHT
''                    .CellDetails i, 3, CStr(cpoints(.RowItemData(i)).Y), DT_RIGHT
'
'                    .CellText(i + 1, 2) = .CellText(i, 2)
'                    .CellText(i + 1, 3) = .CellText(i, 3)
'                Next
                'Create the new cp and start editing
'                For i = 0 To 5
'                    Debug.Print cpexists(i)
'                Next
                'Stop
                cpoints(index).X = 10
                cpoints(index).Y = 10
                cpexists(index) = True
                .AddRow .SelectedRow
                .EndEdit
                .SelectedRow = .SelectedRow - 1
                .RowItemData(.SelectedRow) = lastcp
                .CellDetails .SelectedRow, 1, CStr(tempIndex), DT_RIGHT
                .CellDetails .SelectedRow, 2, CStr(cpoints(index).X), DT_RIGHT
                .CellDetails .SelectedRow, 3, CStr(cpoints(index).Y), DT_RIGHT
                
                changeSameIndex .SelectedRow + 1, tempIndex
                
                .StartEdit .SelectedRow, 2
            End If
        End If
        End With
        dirty = True
    End If
End Sub

Private Sub FillGrid()
    Dim i As Integer
    With grd
        .Redraw = False
        .Clear
        .AddColumn sHeader:="Index"
        .AddColumn sHeader:="X"
        .AddColumn sHeader:="Y"
        .ColumnAlign(1) = ecgHdrTextALignCentre
        .ColumnAlign(2) = ecgHdrTextALignRight
        .ColumnAlign(3) = ecgHdrTextALignRight
        For i = 0 To 999
            If cpexists(i) Then
                .AddRow
                .RowItemData(i + 1) = i
                .CellDetails i + 1, 1, CStr(i), DT_RIGHT
                .CellDetails i + 1, 2, CStr(cpoints(i).X), DT_RIGHT
                .CellDetails i + 1, 3, CStr(cpoints(i).Y), DT_RIGHT
            End If
        Next
        .AddRow lItemData:=-1
        .Redraw = True
    End With
End Sub

Private Sub ConfigureGrid()
    With grd
    .Redraw = False
    'Grid lines
    .GridLines = False
    .GridLineMode = ecgGridFillControl
    'Display and behaviour settings
    .DefaultRowHeight = 15
    .HighlightSelectedIcons = False
    .RowMode = True
    .Editable = True
    .SingleClickEdit = False
    .SelectionOutline = False
    .DrawFocusRectangle = True
    .SelectionAlphaBlend = True
    .OwnerDrawImpl = Me
    .Redraw = True
    End With
End Sub

Private Sub cmdAccept_Click()
    Dim i As Integer
    
    For i = 0 To lastcp
    
        m_map.RemoveCPoint i
        
        If cpexists(i) Then
            m_map.NewCPoint cpoints(i).X, cpoints(i).Y
        Else
            m_map.NewCPoint -1, -1
        End If
        
    Next i
    Unload Me
End Sub

Private Sub cmdAdd_Click()
    AddCp (False)
End Sub

Private Sub cmdBottom_Click()
    With grd
        .CellText(.SelectedRow, 2) = m_map.Width / 2
        .CellText(.SelectedRow, 3) = m_map.Height
    End With
    dirty = True
End Sub

Private Sub cmdBottomLeft_Click()
    With grd
        .CellText(.SelectedRow, 2) = 0
        .CellText(.SelectedRow, 3) = m_map.Height
    End With
    dirty = True
End Sub

Private Sub cmdBottomRight_Click()
    With grd
        .CellText(.SelectedRow, 2) = m_map.Width
        .CellText(.SelectedRow, 3) = m_map.Height
    End With
    dirty = True
End Sub

Private Sub cmdCenter_Click()
    With grd
        .CellText(.SelectedRow, 2) = m_map.Width / 2
        .CellText(.SelectedRow, 3) = m_map.Height / 2
    End With
    dirty = True
End Sub

Private Sub cmdDiscard_Click()
    Unload Me
End Sub

Private Sub cmdExport_Click()
    
    Dim FN As String
    Dim i As Integer
    
    Dim myTextStream As textStream
    
    FN = ShowSaveDialog("cpt", "CP file (*.cpt)|*.cpt| All files (*.*)|(*.*)")

    If FN <> "" Then
    
        If FSO.FileExists(FN) Then
            Kill FN
        End If
        
        Set myTextStream = FSO.OpenTextFile(FN, ForWriting, True)
        
        myTextStream.WriteLine "CTRL-PTS"
        myTextStream.WriteLine CStr(lastcp + 1)
    
        For i = 1 To lastcp + 1
            myTextStream.WriteLine grd.CellText(i, 1)
            'myTextStream.Write " "
            myTextStream.WriteLine grd.CellText(i, 2)
            'myTextStream.Write " "
            myTextStream.WriteLine grd.CellText(i, 3)
        Next i
        
        myTextStream.Close

    End If
    
End Sub

Private Sub cmdImport_Click()

    Dim FN As String
    Dim fileCount As Integer
    Dim sFile() As String
    Dim sMagic As String
    Dim n_points As Integer, indexCP As Integer, X As Integer, Y As Integer, i As Integer
    Dim myTextStream As textStream

    fileCount = ShowOpenDialog(sFile, "CP file (*.cpt)|*.cpt| All files (*.*)|(*.*)")
    
    If fileCount > 0 Then

        Set myTextStream = FSO.OpenTextFile(sFile(0), ForReading)
        
        sMagic = myTextStream.ReadLine
        
        If sMagic <> "CTRL-PTS" Then
              MsgBox "Not Control-Point type file"
              myTextStream.Close
              Exit Sub
        End If
    
        n_points = myTextStream.ReadLine
        
        For i = 1 To n_points
            indexCP = myTextStream.ReadLine
            X = myTextStream.ReadLine
            Y = myTextStream.ReadLine
            With grd
                addCPAt indexCP, X, Y
                cpexists(indexCP) = True
            End With

        Next i
            
        myTextStream.Close
        lastcp = n_points - 1
        
        dirty = True
        
    End If
End Sub

Private Sub cmdInsert_Click()
    AddCp (True)
End Sub

Private Sub cmdInsertAt_Click()

    Dim res As String
    Dim iIndex As Integer
    
begin:
    
    res = InputBox("Enter the index of the CP. Value must be between 0 and 999, both included. Be care that the index is not in the grid yet.", , lastcp + 1)
    
    If res = "" Then
        Exit Sub
    ElseIf IsNumeric(res) Then
        iIndex = CInt(res)
    Else
        MsgBox "Index number incorrect. Please try again", , "Incorrect index"
        GoTo begin
    End If
    
    If 999 >= iIndex Or iIndex >= 0 Then
        addCPAt iIndex, 0, 0
        dirty = True
    ElseIf cpexists(iIndex) Then
        MsgBox "Index exists. Please try another index that doesn't exist", , "Incorrect index"
        GoTo begin
    Else
        MsgBox "Index number incorrect. Please try again", , "Incorrect index"
        GoTo begin
    End If
        
End Sub

Private Sub cmdLeft_Click()
    With grd
        .CellText(.SelectedRow, 2) = 0
        .CellText(.SelectedRow, 3) = m_map.Height / 2
    End With
    dirty = True
End Sub

Private Sub cmdMoveDown_Click()

    Dim tempX As Integer, tempY As Integer
    
    With grd
        If .SelectedRow < lastcp Then

            tempX = .CellText(.SelectedRow, 2)
            tempY = .CellText(.SelectedRow, 3)
            
            .CellText(.SelectedRow, 2) = .CellText(.SelectedRow + 1, 2)
            .CellText(.SelectedRow, 3) = .CellText(.SelectedRow + 1, 3)
            .CellText(.SelectedRow + 1, 2) = tempX
            .CellText(.SelectedRow + 1, 3) = tempY
            
            dirty = True
            
        End If
    End With
End Sub

Private Sub cmdMoveUp_Click()

    Dim tempX As Integer, tempY As Integer
    
    With grd
        If .SelectedRow > 1 Then

            tempX = .CellText(.SelectedRow, 2)
            tempY = .CellText(.SelectedRow, 3)
            
            .CellText(.SelectedRow, 2) = .CellText(.SelectedRow - 1, 2)
            .CellText(.SelectedRow, 3) = .CellText(.SelectedRow - 1, 3)
            .CellText(.SelectedRow - 1, 2) = tempX
            .CellText(.SelectedRow - 1, 3) = tempY
            
            dirty = True
            
        End If
    End With
        
End Sub

Private Sub cmdRemove_Click()
    With grd
        If .Rows > 1 Then
            cpexists(.CellText(.SelectedRow, 1)) = False
'            cpoints(.CellText(.SelectedRow, 1)).X = -1
'            cpoints(.CellText(.SelectedRow, 1)).Y = -1
            .RemoveRow (.SelectedRow)
            If .SelectedRow = .Rows Then
                lastcp = lastcp - 1
            End If
            dirty = True
        End If
    End With
End Sub

Private Sub cmdRight_Click()
    With grd
        .CellText(.SelectedRow, 2) = m_map.Width
        .CellText(.SelectedRow, 3) = m_map.Height / 2
    End With
    dirty = True
End Sub

Private Sub cmdTop_Click()
    With grd
        .CellText(.SelectedRow, 2) = m_map.Width / 2
        .CellText(.SelectedRow, 3) = "0"
    End With
    dirty = True
End Sub

Private Sub cmdTopLeft_Click()
    With grd
        .CellText(.SelectedRow, 2) = "0"
        .CellText(.SelectedRow, 3) = "0"
    End With
    dirty = True
End Sub

Private Sub cmdTopRight_Click()
    With grd
        .CellText(.SelectedRow, 2) = m_map.Width
        .CellText(.SelectedRow, 3) = "0"
    End With
    dirty = True
End Sub

Private Sub Form_Load()
    dirty = False
    ConfigureGrid
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim vbmsg As VbMsgBoxResult
    If dirty Then
        Select Case UnloadMode
            Case vbFormControlMenu:
                vbmsg = MsgBox(MSG_SAVEFILE, vbYesNoCancel)
            Case vbAppTaskManager:
                vbmsg = MsgBox(MSG_SAVEFILE, vbYesNoCancel)
            Case vbFormOwner:
                vbmsg = MsgBox(MSG_SAVEFILE, vbYesNoCancel)
        End Select
        If vbmsg = vbCancel Then
            Cancel = 1
        ElseIf vbmsg = vbYes Then
            Cancel = 1
            cmdAccept_Click
            Cancel = 0
        End If
    End If
End Sub

Private Sub grd_CancelEdit()
    txtEdit.Visible = False
End Sub


Private Sub grd_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
    Dim lLeft As Long, lHeight As Long, lTop As Long, lWidth As Long
    With grd
        If lCol = 1 Or .RowItemData(lRow) = -1 Then
            bCancel = True
        Else
            'Show the editor text box
            .CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight
            txtEdit.text = .CellText(lRow, lCol)
            txtEdit.Move .Left + Screen.TwipsPerPixelX + lLeft, .Top + 2 * Screen.TwipsPerPixelY + lTop + (.RowHeight(lRow) * Screen.TwipsPerPixelY - txtEdit.Height) \ 2, lWidth - 2 * Screen.TwipsPerPixelX
            txtEdit.Visible = True
            txtEdit.SetFocus
            txtEdit.SelStart = Len(txtEdit)
        End If
    End With
End Sub

Private Sub DrawNewFrame(lHdc As Long, lLeft As Long, lTop As Long, lBottom As Long)
    Dim hBr As Long, tr As RECT
    Dim lW As Long
    lW = grd.ColumnWidth(1) + grd.ColumnWidth(2) + grd.ColumnWidth(3)
    SetRect tr, lLeft, lTop, lW - lLeft, lBottom
    hBr = CreateSolidBrush(&H0&)
    FrameRect lHdc, tr, hBr
    Dim oldClr As Long
    oldClr = SetTextColor(lHdc, RGB(125, 125, 125))
    DrawTextA lHdc, "New Control Point", -1, tr, DT_CENTER
    SetTextColor lHdc, oldClr
    DeleteObject hBr
End Sub

Private Sub grd_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
    With grd
        If lRow = .Rows Then 'New CP Row
            AddCp (False)
        End If
    End With
End Sub

Private Sub IGridCellOwnerDraw_Draw( _
      cell As cGridCell, _
      ByVal lHdc As Long, _
      ByVal eDrawStage As ECGDrawStage, _
      ByVal lLeft As Long, ByVal lTop As Long, _
      ByVal lRight As Long, ByVal lBottom As Long, _
      bSkipDefault As Boolean _
   )
   If (eDrawStage = ecgBeforeIconAndText) Then
      If (grd.RowItemData(cell.Row) = -1) And cell.Column = 1 Then
            DrawNewFrame lHdc, lLeft, lTop, lBottom
            bSkipDefault = True
      End If
   End If
End Sub

Private Sub txtEdit_Change()

    Dim SCoord As String
    
    With grd
        SCoord = CStr(CInt(txtEdit.text))
        .CellDetails .SelectedRow, .SelectedCol, SCoord, DT_RIGHT
        
    End With
End Sub

Private Sub addCPAt(index As Integer, X As Integer, Y As Integer)
    With grd
        If lastcp < index Then
            lastcp = index
        End If
        cpoints(index).X = 0
        cpoints(index).Y = 0
        cpexists(index) = True
        
        .AddRow .Rows
        .RowItemData(.Rows - 1) = lastcp
        .CellDetails .Rows - 1, 1, CStr(index), DT_RIGHT
        .CellDetails .Rows - 1, 2, CStr(X), DT_RIGHT
        .CellDetails .Rows - 1, 3, CStr(Y), DT_RIGHT
        .SelectedRow = .Rows - 1
        .EndEdit
        .StartEdit .Rows - 1, 2
            
    End With
End Sub

Private Sub changeSameIndex(from As Integer, index As Integer)
    Dim i As Integer
    With grd
        For i = from To .Rows
            If CInt(.CellText(i, 1)) = index Then
                changeSameIndex i, index + 1
                .CellText(i, 1) = .CellText(i, 1) + 1
            End If
        Next i
    End With
End Sub
