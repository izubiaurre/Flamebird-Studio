VERSION 5.00
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.5#0"; "vbalsgrid6.ocx"
Begin VB.Form frmDevelopersList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Developer List"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8805
   Icon            =   "frmDevelopersList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   8805
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
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
      Left            =   5280
      TabIndex        =   7
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdDefaultDev 
      Caption         =   "Make default developer"
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
      TabIndex        =   6
      Top             =   4140
      Width           =   2775
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
      Left            =   1800
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   3855
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
      Left            =   7680
      TabIndex        =   4
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
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
      Left            =   6480
      TabIndex        =   3
      Top             =   4080
      Width           =   1095
   End
   Begin vbAcceleratorSGrid6.vbalGrid grdDev 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5530
      GridLines       =   -1  'True
      GridLineMode    =   1
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
      HeaderButtons   =   0   'False
      HeaderDragReorderColumns=   0   'False
      HeaderFlat      =   -1  'True
      BorderStyle     =   2
      Editable        =   -1  'True
      DisableIcons    =   -1  'True
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The developer's list"
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
      TabIndex        =   2
      Top             =   75
      Width           =   1650
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmDevelopersList.frx":058A
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
      Top             =   300
      Width           =   8085
   End
   Begin VB.Image Image1 
      Height          =   765
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8835
   End
End
Attribute VB_Name = "frmDevelopersList"
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

Private Const CLR_DEFAULTDEV = &HB4DD95

Private fntLink As StdFont
Private devcol As cDeveloperCollection
Private RowDefaultDev As Long

Implements IGridCellOwnerDraw

'Dibuja un frame del tamaño de las columnas y escribe el texto centrado
Private Sub DrawNewDevFrame(lHdc As Long, lLeft As Long, lTop As Long, lBottom As Long)
    Dim hBr As Long, tr As RECT
    Dim lW As Long
    lW = grdDev.ColumnWidth(1) + grdDev.ColumnWidth(2) + grdDev.ColumnWidth(3)
    SetRect tr, lLeft, lTop, lW - lLeft, lBottom
    hBr = CreateSolidBrush(&H0&)
    FrameRect lHdc, tr, hBr
    Dim oldClr As Long
    oldClr = SetTextColor(lHdc, RGB(125, 125, 125))
    DrawTextA lHdc, "Click here to add a new developer", -1, tr, DT_CENTER
    SetTextColor lHdc, oldClr
    DeleteObject hBr
End Sub

Private Sub cmdApply_Click()
    SaveDevelopers devcol
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDefaultDev_Click()
    Dim i As Integer, j As Integer
    With grdDev
    If .SelectedRow > 0 Then
        If .RowItemData(.SelectedRow) <> -1 Then  'Fila válida
            For i = 1 To .Rows - 1
                For j = 1 To .Columns
                    'Pintamos la fila seleccionada de un color y el resto de otro
                    .CellBackColor(i, j) = IIf(i = .SelectedRow, CLR_DEFAULTDEV, vbWhite)
                Next
            Next
            RowDefaultDev = .SelectedRow
            'devcol.DefaultDev = .CellText(.SelectedRow, 1)
        End If
    End If
   ' .SetFocus
    End With
End Sub

Private Sub cmdOk_Click()
    cmdApply_Click
    Unload Me
End Sub

Private Sub grdDev_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 And Not grdDev.RowItemData(grdDev.SelectedRow) = -1 Then 'Suprimir
        grdDev.RemoveRow grdDev.SelectedRow
    End If
End Sub

Private Sub grdDev_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
    Dim str As String, i As Integer, cnt As Integer
    With grdDev
        If lRow = .Rows Then 'Si es la celda Nuevo Desarrollador agregamos uno
            'Obtenemos un nombre válido
            cnt = 1
            Do
                str = "New Developer " & CStr(cnt)
                For i = 1 To .Rows - 1
                    If .CellText(i, 1) = str Then 'Existe el developer
                        str = ""
                        cnt = cnt + 1
                    End If
                Next
            Loop Until str <> ""
            
            .AddRow .Rows
            .CellDetails .Rows - 1, 1, str
            .CellDetails .Rows - 1, 3, "", oForecolor:=vbBlue, oFont:=fntLink
            .SelectedRow = .Rows - 1
            .EndEdit
            .StartEdit .Rows - 1, 1
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
      If (grdDev.RowItemData(cell.Row) = -1) And cell.Column = 1 Then
            DrawNewDevFrame lHdc, lLeft, lTop, lBottom
            bSkipDefault = True
      End If
   End If
End Sub

Private Sub ConfigureGrid()
    Dim iCol As Long
    
    With grdDev
        .Redraw = False
        
        ' Set grid lines
        .GridLines = False
        .GridLineMode = ecgGridFillControl
        
        ' Various display and behaviour settings
        .DefaultRowHeight = 15
        .HighlightSelectedIcons = False
        .RowMode = True
        .Editable = True
        .SingleClickEdit = False
        .StretchLastColumnToFit = True
        .SelectionOutline = False
        .DrawFocusRectangle = True
        .SelectionAlphaBlend = True
        .OwnerDrawImpl = Me
      
        'Fuente Link
        Set fntLink = .font
        fntLink.Underline = True
        
        .Redraw = True
    End With
End Sub

Private Sub addNewDeveloperRow()
    grdDev.AddRow lItemData:=-1
End Sub

Private Sub LoadDevelopers(devs As cDeveloperCollection)
    Dim dev As cDeveloper
    Dim i As Integer
    With grdDev
        'Añade los desarrolladores
        For Each dev In devs
            
            .AddRow
            .CellDetails .Rows, 1, dev.name
            .CellDetails .Rows, 2, dev.RealName
            .CellDetails .Rows, 3, dev.Mail, oFont:=fntLink, oForecolor:=vbBlue
        Next
        'Establece el Default Developer
        If devs.defaultDev <> "" Then
            For i = 1 To .Rows
                If .CellText(i, 1) = devs.defaultDev Then
                    grdDev.SelectedRow = i
                End If
            Next
        End If
    End With
    Set dev = Nothing
End Sub

Private Sub SaveDevelopers(devs As cDeveloperCollection)
    Dim i As Integer
    devs.Clear
    With grdDev
        For i = 1 To .Rows - 1
            devs.Add .CellText(i, 1), .CellText(i, 2), .CellText(i, 3)
        Next
        'Asignamos el DefaultDev en caso de que haya
        If RowDefaultDev > 0 And RowDefaultDev <= .Rows Then
            devcol.defaultDev = IIf(.CellBackColor(RowDefaultDev, 1) = CLR_DEFAULTDEV, .CellText(RowDefaultDev, 1), "")
        End If
    End With
End Sub

Private Sub Form_Load()
    Image1.Picture = LoadPicture(App.Path & "\Resources\frmHeader.jpg")
    With grdDev
        .AddColumn "DevName", "Developer name", , , 100
        .AddColumn "RealName", "Real name", , , 150
        .AddColumn "Mail", "E-mail"
        .StretchLastColumnToFit = True
    End With
    
    ConfigureGrid 'configuración
    Set devcol = openedProject.devcol
    
    LoadDevelopers devcol  'Cargamos la lista en el grid
    addNewDeveloperRow
    cmdDefaultDev_Click
End Sub

Private Sub grdDev_CancelEdit()
     txtEdit.Visible = False
End Sub

Private Sub grdDev_PreCancelEdit(ByVal lRow As Long, ByVal lCol As Long, newValue As Variant, bStayInEditMode As Boolean)
    Dim i As Long
    With grdDev
        If .RowItemData(lRow) <> -1 Then 'No es la fila Nuevo Desarollador
            If lCol = 1 And txtEdit = "" Then 'Developer name no puede estar vacio
                MsgBox "The 'Developer Name' field can't be empty", vbExclamation, "Developer List"
                txtEdit = .CellText(lRow, lCol)
            End If
            If lCol = 1 Then 'Si es el Dev Name, comprobamos que no esté repetido
                For i = 1 To .Rows - 1
                    If txtEdit = .CellText(i, 1) And i <> lRow Then
                        MsgBox "There is another developer using this name", vbExclamation, "Developer List"
                        txtEdit = .CellText(lRow, lCol)
                    End If
                Next
            End If
            .CellText(lRow, lCol) = txtEdit 'Actualiza el texto de la celda
        End If
    End With
End Sub

Private Sub grdDev_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
    Dim lLeft As Long, lHeight As Long, lTop As Long, lWidth As Long
    With grdDev
        If .RowItemData(lRow) = -1 Then 'Fila nuevo desarrollador
            bCancel = True
        Else 'Fila normal
            'Muestra el textbox donde sea necesário
            .CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight
            txtEdit.text = .CellText(lRow, lCol)
            txtEdit.Move .Left + Screen.TwipsPerPixelX + lLeft, .Top + 2 * Screen.TwipsPerPixelY + lTop + (.RowHeight(lRow) * Screen.TwipsPerPixelY - txtEdit.Height) \ 2, lWidth - 2 * Screen.TwipsPerPixelX
            txtEdit.Visible = True
            txtEdit.SetFocus
            txtEdit.SelStart = Len(txtEdit)
        End If
    End With
End Sub

Private Sub txtEdit_GotFocus()
    cmdOk.Default = False
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then grdDev.EndEdit: KeyAscii = 0
End Sub

Private Sub txtEdit_LostFocus()
    cmdOk.Default = True
End Sub
