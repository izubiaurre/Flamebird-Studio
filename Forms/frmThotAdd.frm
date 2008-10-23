VERSION 5.00
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.5#0"; "vbalsgrid6.ocx"
Object = "{9DC93C3A-4153-440A-88A7-A10AEDA3BAAA}#3.5#0"; "vbaldtab6.ocx"
Begin VB.Form frmThotAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add..."
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   11640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmThotAdd.frx":0000
   ScaleHeight     =   608
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   776
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   4875
      TabIndex        =   21
      Top             =   6840
      Width           =   4935
      Begin VB.CommandButton cmdOk 
         Caption         =   "Ok"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   0
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3480
         TabIndex        =   22
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.PictureBox picStruct 
      Height          =   5775
      Left            =   6360
      ScaleHeight     =   5715
      ScaleWidth      =   4875
      TabIndex        =   12
      Top             =   1200
      Width           =   4935
      Begin VB.ComboBox cmbStructType 
         Height          =   315
         Left            =   1440
         TabIndex        =   24
         Top             =   3000
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Caption         =   "Element List"
         Height          =   3255
         Left            =   120
         TabIndex        =   19
         Top             =   1800
         Width           =   4695
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1875
            Left            =   3600
            ScaleHeight     =   1875
            ScaleWidth      =   315
            TabIndex        =   35
            Top             =   780
            Width           =   315
            Begin VB.CommandButton cmdStructAdd 
               Caption         =   "+"
               Height          =   255
               Left            =   0
               TabIndex        =   40
               Top             =   0
               Width           =   255
            End
            Begin VB.CommandButton cmdStructRemove 
               Caption         =   "-"
               Height          =   255
               Left            =   0
               TabIndex        =   39
               Top             =   360
               Width           =   255
            End
            Begin VB.CommandButton cmdStructUp 
               Height          =   255
               Left            =   0
               Picture         =   "frmThotAdd.frx":0342
               Style           =   1  'Graphical
               TabIndex        =   38
               Top             =   720
               Width           =   255
            End
            Begin VB.CommandButton cmdStructDown 
               Height          =   255
               Left            =   0
               Picture         =   "frmThotAdd.frx":03D0
               Style           =   1  'Graphical
               TabIndex        =   37
               Top             =   1080
               Width           =   255
            End
            Begin VB.CommandButton cmdStructClear 
               Caption         =   "x"
               Height          =   255
               Left            =   0
               TabIndex        =   36
               Top             =   1440
               Width           =   255
            End
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   2640
            ScaleHeight     =   435
            ScaleWidth      =   1575
            TabIndex        =   28
            Top             =   240
            Width           =   1575
            Begin VB.CommandButton cmdStructRight 
               Caption         =   ">"
               Height          =   255
               Left            =   900
               TabIndex        =   31
               Top             =   60
               Width           =   375
            End
            Begin VB.TextBox txtStructElementList 
               Height          =   285
               Left            =   480
               TabIndex        =   30
               Top             =   60
               Width           =   375
            End
            Begin VB.CommandButton cmdStructLeft 
               Caption         =   "<"
               Height          =   255
               Left            =   60
               TabIndex        =   29
               Top             =   60
               Width           =   375
            End
         End
         Begin vbAcceleratorSGrid6.vbalGrid grdStructList 
            Height          =   1695
            Left            =   240
            TabIndex        =   20
            Top             =   840
            Width           =   3255
            _ExtentX        =   5741
            _ExtentY        =   2990
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
            DisableIcons    =   -1  'True
         End
      End
      Begin VB.TextBox txtStructNumbers 
         Height          =   285
         Left            =   1740
         TabIndex        =   18
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox TxtStructDesc 
         Height          =   615
         Left            =   960
         TabIndex        =   16
         Top             =   540
         Width           =   3195
      End
      Begin VB.TextBox txtStructName 
         Height          =   315
         Left            =   960
         TabIndex        =   14
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Numbers of Items:"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "Desc:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   855
      End
   End
   Begin vbalDTab6.vbalDTabControl tabCategories 
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      TabAlign        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin VB.PictureBox picProcFunc 
      Height          =   5595
      Left            =   0
      ScaleHeight     =   369
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   361
      TabIndex        =   0
      Top             =   1200
      Width           =   5475
      Begin VB.Frame fraType 
         Caption         =   "Type"
         Height          =   855
         Left            =   120
         TabIndex        =   8
         Top             =   1140
         Width           =   5235
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   240
            ScaleHeight     =   615
            ScaleWidth      =   1095
            TabIndex        =   25
            Top             =   180
            Width           =   1095
            Begin VB.OptionButton optFunction 
               Caption         =   "Function"
               Height          =   255
               Left            =   60
               TabIndex        =   27
               Top             =   300
               Width           =   975
            End
            Begin VB.OptionButton optProcess 
               Caption         =   "Process"
               Height          =   255
               Left            =   60
               TabIndex        =   26
               Top             =   60
               Width           =   1095
            End
         End
         Begin VB.ComboBox cmbReturn 
            Height          =   315
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   480
            Width           =   1695
         End
      End
      Begin VB.Frame fraLocal 
         Caption         =   "Local Variables"
         Height          =   1695
         Left            =   120
         TabIndex        =   6
         Top             =   3840
         Width           =   5295
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            FillColor       =   &H8000000F&
            ForeColor       =   &H80000008&
            Height          =   1455
            Left            =   3840
            ScaleHeight     =   1455
            ScaleWidth      =   1335
            TabIndex        =   43
            Top             =   240
            Width           =   1335
            Begin VB.CommandButton cmdLocalDel 
               Caption         =   "-"
               Height          =   255
               Left            =   0
               TabIndex        =   49
               Top             =   240
               Width           =   375
            End
            Begin VB.CommandButton cmdLocalAdd 
               Caption         =   "+"
               Height          =   255
               Left            =   0
               TabIndex        =   48
               Top             =   0
               Width           =   375
            End
            Begin VB.CommandButton cmdLocalClearAll 
               Caption         =   "x"
               Height          =   315
               Left            =   0
               TabIndex        =   47
               Top             =   480
               Width           =   375
            End
         End
         Begin VB.ComboBox cmbLocal 
            Height          =   315
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   600
            Visible         =   0   'False
            Width           =   1215
         End
         Begin vbAcceleratorSGrid6.vbalGrid grdLocal 
            Height          =   1215
            Left            =   240
            TabIndex        =   10
            Top             =   240
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   2143
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
            DisableIcons    =   -1  'True
         End
      End
      Begin VB.Frame fraParameters 
         Caption         =   "Parameters"
         Height          =   1695
         Left            =   120
         TabIndex        =   5
         Top             =   2040
         Width           =   5295
         Begin VB.PictureBox Picture4 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1395
            Left            =   3840
            ScaleHeight     =   1395
            ScaleWidth      =   1035
            TabIndex        =   32
            Top             =   240
            Width           =   1035
            Begin VB.CommandButton cmdRemoveParam 
               Caption         =   "-"
               Height          =   255
               Left            =   0
               TabIndex        =   46
               Top             =   240
               Width           =   375
            End
            Begin VB.CommandButton cmdAddParam 
               Caption         =   "+"
               Height          =   255
               Left            =   0
               TabIndex        =   45
               Top             =   0
               Width           =   375
            End
            Begin VB.CommandButton cmdClearAllParam 
               Caption         =   "x"
               Height          =   315
               Left            =   0
               TabIndex        =   44
               Top             =   960
               Width           =   375
            End
            Begin VB.CommandButton cmdUpParam 
               Height          =   255
               Left            =   0
               Picture         =   "frmThotAdd.frx":045F
               Style           =   1  'Graphical
               TabIndex        =   34
               Top             =   480
               Width           =   375
            End
            Begin VB.CommandButton cmdDownParam 
               Height          =   255
               Left            =   0
               Picture         =   "frmThotAdd.frx":04ED
               Style           =   1  'Graphical
               TabIndex        =   33
               Top             =   720
               Width           =   375
            End
         End
         Begin vbAcceleratorSGrid6.vbalGrid grdList 
            Height          =   1095
            Left            =   240
            TabIndex        =   7
            Top             =   240
            Width           =   3555
            _ExtentX        =   6271
            _ExtentY        =   1931
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
            DisableIcons    =   -1  'True
         End
      End
      Begin VB.TextBox txtDesc 
         Height          =   645
         Left            =   660
         TabIndex        =   4
         Top             =   480
         Width           =   4275
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   660
         TabIndex        =   2
         Top             =   120
         Width           =   2715
      End
      Begin VB.Label lblDesc 
         Caption         =   "Desc:"
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblName 
         Caption         =   "Name:"
         Height          =   255
         Left            =   60
         TabIndex        =   1
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Add processes, functions or structs to the code"
      Height          =   375
      Left            =   360
      TabIndex        =   42
      Top             =   240
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   870
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5565
   End
End
Attribute VB_Name = "frmThotAdd"
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
Option Base 0

Dim Table()
Private m_flat As cFlatControl

'Sets controls placement and size
Private Sub PlaceControls()
    tabCategories.Move 0, 60, 361, 430  ' 0, 920, 553, 4425
    Me.Width = 5535     '6100
    Me.Height = 7800    ' 6395
    'cmdCancel.Move 4320, 5400   ' 5380
    'cmdOk.Move 120, 5400        ' 3120
End Sub

Private Sub cmbLocal_Change()
    grdLocal.CellDetails grdLocal.SelectedRow, 1, cmbLocal.text, DT_RIGHT
End Sub

Private Sub cmbStructType_Click()
    grdStructList.CellDetails grdStructList.SelectedRow, 2, cmbStructType.text, DT_RIGHT
End Sub

Private Sub cmdAddParam_Click()
    grdList.AddRow
    'grdList.RowItemData(grdList.Rows) = InputBox("Insert name of parameter")
    grdList.CellDetails grdList.Rows, 1, CStr(InputBox("Insert name of parameter")), DT_RIGHT
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClearAllParam_Click()
    With grdList
        '.SelectedCol
        .SelectedRow = 1
         Do While .SelectedRow < .Rows
            .RemoveRow .SelectedRow
        Loop
        .RemoveRow 1
    End With
End Sub

Private Sub cmdClearLocal_Click()
    With grdLocal
        .SelectedRow = 1
        Do While .SelectedRow < .Rows
           .RemoveRow .SelectedRow
        Loop
        .RemoveRow 1
    End With
End Sub

Private Sub cmdDownParam_Click()

    Dim tempX As Integer ' tempY As Integer
    
    With grdList
        'If .SelectedRow < lastcp Then
        If .SelectedRow < .Rows Then

            tempX = CInt(.CellText(.SelectedRow, 1))
            'tempY = CInt(.CellText(.SelectedRow, 3))
            .CellText(.SelectedRow, 1) = CInt(.CellText(.SelectedRow + 1, 1))
            '.CellText(.SelectedRow, 3) = CInt(.CellText(.SelectedRow + 1, 3))
            .CellText(.SelectedRow + 1, 1) = tempX
            '.CellText(.SelectedRow + 1, 3) = tempY
            If .SelectedRow < .Rows Then
                .SelectedRow = .SelectedRow + 1
            End If
            

        End If
    End With
End Sub

Private Sub cmdLocalAdd_Click()
    grdLocal.AddRow
    cmbLocal.Visible = True
    cmbLocal.SetFocus
End Sub

Private Sub cmdLocalClearAll_Click()
    With grdLocal
    '.SelectedCol
        .SelectedRow = 1
         Do While .SelectedRow < .Rows
            .RemoveRow .SelectedRow
        Loop
        .RemoveRow 1
    End With
End Sub

Private Sub cmdLocalDel_Click()
    grdList.RemoveRow grdLocal.SelectedRow
End Sub

Private Sub cmdOk_Click()
    Dim Message As String, Message2 As String
    Dim i As Integer
    
    If tabCategories.SelectedTab.Key = "STRUCT" Then
        If (txtStructName.text = "") Then
            MsgBox "Please, write the name of  the struct"
            txtStructName.SetFocus
            Exit Sub
        End If
    
        Message2 = "// " & TxtStructDesc.text & vbCrLf
        Message2 = Message2 & "struct" & " " & txtStructName.text & "[ " & txtStructNumbers.text & " ]" & vbCrLf
        
        With grdStructList
            .SelectedRow = 1
            
            Do While .SelectedRow < .Rows
                  Message2 = Message2 & .CellText(.SelectedRow, 2) & " "
                  Message2 = Message2 & .CellText(.SelectedRow, 1) & " " & ";" & vbCrLf
                  'Message2 = Message2 & .CellText(.SelectedRow, 3) & " " & vbCrLf
                     .SelectedRow = .SelectedRow + 1
            Loop
     
            i = grdStructList.Rows
            'i = i - 1
            Message2 = Message2 & .CellText(i, 2) & " "
            Message2 = Message2 & .CellText(i, 1) & " " & vbCrLf & ";"
            'Message2 = Message2 & .CellText(1, 3) & " " & vbCrLf
            Message2 = Message2 & "end="
            .SelectedRow = 1
            Do While .SelectedRow < .Rows
                Message2 = Message2 & .CellText(.SelectedRow, 3) & ", "
                .SelectedRow = .SelectedRow + 1
            Loop
      i = grdStructList.Rows
     ' i = i - 1
      Message2 = Message2 & .CellText(i, 3) & ";"
      MsgBox Message2
        End With
    Else
    If (txtName.text = "") Then
        MsgBox "Please, write the name of  the function or process"
        txtName.SetFocus
        Exit Sub
    End If
    
    Message = txtName.text
    
    If optFunction Then
        Message = cmbReturn.text & " Function " & Message
    Else
        Message = "Process " & Message
    End If

         With grdList
         .SelectedRow = 1
         Do While .SelectedRow < .Rows
                  Message = Message & "(" & .CellText(.SelectedRow, 1) & ","
                  .SelectedRow = .SelectedRow + 1
         Loop
        Message = Message & .CellText(.SelectedRow, 1) & ")" & vbCrLf
         End With
         
'  With grdLocal
'    .SelectedRow = 1
'        Do While .SelectedRow < .Rows
'            If .CellText(.SelectedRow, 2) <> "" Then
'            Message = Message & .CellText(.SelectedRow, 1) & "  = " & .CellText(.SelectedRow, 2) & ";" & vbCrLf
'            End If
'            .SelectedRow = .SelectedRow + 1
'        Loop
'   End With
   
    If (txtDesc.text <> " ") Then
        Message = Message & " \\ " & txtDesc.text & vbCrLf & " Begin " & vbCrLf & vbCrLf & " End "
    Else
        Message = Message & " Begin " & vbCrLf & "End"
    End If
    
    MsgBox Message
 End If
End Sub

Private Sub cmdRemoveParam_Click()
    grdList.RemoveRow grdList.SelectedRow
End Sub

Private Sub cmdStructAdd_Click()
    Dim tempX As String, tempY As Integer
    
    If txtStructNumbers.text = "" Then
        MsgBox "Insert a number of items"
        txtStructNumbers.SetFocus
        Exit Sub
    End If
    
    grdStructList.AddRow
    'grdList.RowItemData(grdList.Rows) = InputBox("Insert name of parameter")
    grdStructList.CellDetails grdStructList.Rows, 1, CStr(InputBox("Insert name of parameter")), DT_RIGHT
    grdStructList.CellDetails grdStructList.Rows, 2, CStr(InputBox("Insert  type of parameter")), DT_RIGHT
    grdStructList.CellDetails grdStructList.Rows, 3, CStr(InputBox("Insert a value of the parameter", , 0)), DT_RIGHT
    tempY = CInt(grdStructList.CellText(grdStructList.SelectedRow + 1, 3))
    ReDim Preserve Table(CInt(txtStructNumbers.text), grdStructList.Rows)
    Table(CInt(txtStructNumbers.text), grdStructList.Rows) = tempY
End Sub

Private Sub cmdStructClear_Click()
With grdStructList
        .SelectedRow = 1
         Do While .SelectedRow < .Rows
            .RemoveRow .SelectedRow
        Loop
        .RemoveRow 1
    End With
ReDim Table(CInt(txtStructNumbers.text), 0)
End Sub

Private Sub cmdStructDown_Click()
 Dim tempX As String, tempY As Integer
    
    With grdStructList
        'If .SelectedRow < lastcp Then
        If .SelectedRow < .Rows Then

            tempX = CStr(.CellText(.SelectedRow, 1))
            tempY = CInt(.CellText(.SelectedRow, 3))
            .CellText(.SelectedRow, 1) = CStr(.CellText(.SelectedRow + 1, 1))
            .CellText(.SelectedRow, 3) = CInt(.CellText(.SelectedRow + 1, 3))
            .CellText(.SelectedRow + 1, 1) = tempX
            .CellText(.SelectedRow + 1, 3) = tempY
            If .SelectedRow < .Rows Then
                .SelectedRow = .SelectedRow + 1
            End If
            

        End If
    End With
End Sub

Private Sub cmdStructLeft_Click()
Dim i As Integer, j As Integer
j = 0
With grdStructList

    If CInt(txtStructElementList.text) > 0 Then
       txtStructElementList.text = CInt(txtStructElementList.text) - 1
       i = txtStructElementList.text
       Do While j < .Rows
       .CellDetails j, 3 = Table(i, j)
       j = j + 1
       Loop
    End If
 End With
    
End Sub

Private Sub cmdStructRemove_Click()
 grdStructList.RemoveRow grdStructList.SelectedRow
 ReDim Preserve Table(CInt(txtStructNumbers.text), grdStructList.Rows)
End Sub

Private Sub cmdStructRight_Click()
Dim i As Integer, j As Integer
j = 0
With grdStructList

    If CInt(txtStructElementList.text) < CInt(txtStructNumbers.text) Or CInt(txtStructNumbers.text) = 0 Then
        txtStructElementList.text = CInt(txtStructElementList.text) + 1
        i = txtStructElementList.text
       Do While j < .Rows
        .CellDetails j, 3 = Table(i, j)
        j = j + 1
       Loop
    End If
 End With
    
End Sub

Private Sub cmdStructUp_Click()
  Dim tempX As String, tempY As Integer
    With grdStructList
        If .SelectedRow > 1 Then

            tempX = CStr(.CellText(.SelectedRow, 1))
            tempY = CInt(.CellText(.SelectedRow, 3))
            
            .CellText(.SelectedRow, 1) = .CellText(.SelectedRow - 1, 1)
            .CellText(.SelectedRow, 3) = .CellText(.SelectedRow - 1, 3)
            .CellText(.SelectedRow - 1, 1) = tempX
             .CellText(.SelectedRow - 1, 3) = tempY
            If .SelectedRow > 2 Then
                .SelectedRow = .SelectedRow - 1
            End If
        End If
    End With
End Sub
Private Sub cmdUpParam_Click()
  Dim tempX As Integer 'tempY As Integer
    With grdList
        If .SelectedRow > 1 Then

            'tempX = CInt(.CellText(.SelectedRow, 1))
            'tempY = CInt(.CellText(.SelectedRow, 3))
            
            .CellText(.SelectedRow, 1) = .CellText(.SelectedRow - 1, 1)
            '.CellText(.SelectedRow, 3) = .CellText(.SelectedRow - 1, 3)
            .CellText(.SelectedRow - 1, 1) = tempX
            '.CellText(.SelectedRow - 1, 3) = tempY
            If .SelectedRow > 2 Then
                .SelectedRow = .SelectedRow - 1
            End If
        End If
    End With
End Sub

Private Sub Form_Load()

    On Error GoTo errhandler
    Dim mnuPreSets As cMenus
    Dim i As Integer
    Image1.Picture = LoadPicture(App.Path & "\Resources\frmHeader.jpg")

    PlaceControls

    'Set m_flat = New cFlatControl
    'm_flat.Attach picPredefSets
    'Set mnuPresets = New cMenus
    'mnuPreSets.CreateFromNothing Me.Hwnd

    'Create the tabs
    Dim nTab As cTab
    With tabCategories
        .ImageList = 0
        Set nTab = .Tabs.Add("PROCFUNC", , "ProcFunc")
        nTab.Panel = picProcFunc
        Set nTab = .Tabs.Add("STRUCT", , "Struct")
        nTab.Panel = picStruct

    End With
    
    With grdList
            .Redraw = False
            'Grid lines
            .GridLines = True
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
            '.OwnerDrawImpl = Me
            .Redraw = True
            .AddColumn , "Name", ecgHdrTextALignCentre
    End With
      With grdLocal
            .Redraw = False
            'Grid lines
            .GridLines = True
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
            '.OwnerDrawImpl = Me
            .Redraw = True
            .AddColumn , "Name", ecgHdrTextALignCentre
            .AddColumn , "Value", ecgHdrTextALignCentre
            'cmbReturn.AddItem
           ' cmbReturn.AddItem = ""
            For i = 1 To UBound(typeList)
                cmbReturn.AddItem typeList(i)
            Next i
            'cmbStructType.AddItem
            cmbStructType.AddItem ""
            For i = 1 To UBound(typeList)
                cmbStructType.AddItem typeList(i)
            Next i
            For i = 1 To UBound(localList)
                cmbLocal.AddItem localList(i)
            Next i
    End With
     With grdStructList
            .Redraw = False
            'Grid lines
            .GridLines = True
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
            '.OwnerDrawImpl = Me
            .Redraw = True
            .AddColumn , "Name", ecgHdrTextALignCentre
            .AddColumn , "Type", ecgHdrTextALignCentre
            .AddColumn , "Initial value", ecgHdrTextALignCentre
            End With
    optProcess.Value = True
    cmbReturn.Enabled = False
    Exit Sub
    
errhandler:
    If Err.Number > 0 Then ShowError ("frmThotAdd.Form_Load")
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        cmdCancel_Click
    End If
End Sub

Private Sub grdLocal_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
Dim lLeft As Long, lHeight As Long, lTop As Long, lWidth As Long
    With grdLocal
        If lCol = 1 Then
            cmbLocal.Visible = True
            cmbLocal.SetFocus
        End If
        If lCol = 2 Then
            'Show the editor text box
            .CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight
            .CellText(grdLocal.SelectedRow, 2) = InputBox("Insert value of local variable")
        End If
    End With
    
End Sub


Private Sub grdStructList_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
Dim lLeft As Long, lHeight As Long, lTop As Long, lWidth As Long
Dim i As Integer
    With grdStructList
        If lCol = 3 Then
            'Show the editor text box
            .CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight
            .CellText(grdStructList.SelectedRow, 3) = InputBox("Insert the value of the struct value")
        End If
        If lCol = 2 Then
            cmbStructType.Visible = True
            i = grdStructList.SelectedRow
            i = i - 1
            i = i * 200
            cmbStructType.Top = 3000 + i
            cmbStructType.SetFocus
        End If
    End With
    
End Sub

Private Sub optFunction_Click()
    cmbReturn.Enabled = True
    optProcess.Value = False
    fraLocal.Enabled = False
End Sub

Private Sub optProcess_Click()
    cmbReturn.Enabled = False
    optFunction.Value = False
    fraLocal.Enabled = True
End Sub

Private Sub txtStructNumbers_Change()
    ReDim Preserve Table(CInt(txtStructNumbers.text), grdStructList.Rows)
End Sub

