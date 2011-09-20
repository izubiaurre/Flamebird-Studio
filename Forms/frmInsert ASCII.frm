VERSION 5.00
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.5#0"; "vbalSGrid6.ocx"
Begin VB.Form frmInsertASCII 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert ASCII"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   7485
   Icon            =   "frmInsert ASCII.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton cmdInsert 
      Caption         =   "Insert"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   5040
      Width           =   1335
   End
   Begin vbAcceleratorSGrid6.vbalGrid grd 
      Height          =   4840
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7223
      _ExtentX        =   12753
      _ExtentY        =   8546
      GridLines       =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Header          =   0   'False
      HeaderButtons   =   0   'False
      HeaderDragReorderColumns=   0   'False
      HeaderFlat      =   -1  'True
      BorderStyle     =   2
      DisableIcons    =   -1  'True
      DefaultRowHeight=   30
   End
   Begin VB.Label lblDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Description of the ASCII char."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   5160
      Width           =   4455
   End
End
Attribute VB_Name = "frmInsertASCII"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0


Private Type ASCII
    Char As String
    num As Integer
    desc As String
    type As String
    shortkeys As String
End Type

Private ascii_table(255) As ASCII


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdInsert_Click()
    grd_DblClick grd.SelectedRow, grd.SelectedCol
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        cmdCancel_Click
    End If
End Sub

Private Sub Form_Load()
    ConfigureGrid
    FillGrid
    fillASCII
    FillData
End Sub

Private Sub FillGrid()
    Dim i As Integer
    Dim cont As Integer
    With grd
        .Redraw = False
        .Clear
        .AddColumn , "", , , 0  'empty
        .AddColumn , "0", , , 30
        .AddColumn , "1", , , 30
        .AddColumn , "2", , , 30
        .AddColumn , "3", , , 30
        .AddColumn , "4", , , 30
        .AddColumn , "5", , , 30
        .AddColumn , "6", , , 30
        .AddColumn , "7", , , 30
        .AddColumn , "8", , , 30
        .AddColumn , "9", , , 30
        .AddColumn , "A", , , 30
        .AddColumn , "B", , , 30
        .AddColumn , "C", , , 30
        .AddColumn , "D", , , 30
        .AddColumn , "E", , , 30
        .AddColumn , "F", , , 30
        '.ColumnAlign(1) = ecgHdrTextALignCentre

        .AddRow lItemData:=-1
        .AddRow lItemData:=-1
        .AddRow lItemData:=-1
        .AddRow lItemData:=-1
        .AddRow lItemData:=-1
        .AddRow lItemData:=-1
        .AddRow lItemData:=-1
        .AddRow lItemData:=-1
        .AddRow lItemData:=-1
        .AddRow lItemData:=-1
        .AddRow lItemData:=-1
        .AddRow lItemData:=-1
        .AddRow lItemData:=-1
        .AddRow lItemData:=-1
        .AddRow lItemData:=-1
        .AddRow lItemData:=-1
        .Redraw = True
    End With
End Sub

Private Sub ConfigureGrid()
    With grd
    .Redraw = False
    'Grid lines
    .GridLines = True
    .GridLineMode = ecgGridFillControl
    'Display and behaviour settings
    .DefaultRowHeight = 20
    .HighlightSelectedIcons = False
    .RowMode = False
    .Editable = True
    .SingleClickEdit = False
    .SelectionOutline = False
    .DrawFocusRectangle = True
    .SelectionAlphaBlend = True
    '.OwnerDrawImpl = Me
    .Redraw = True
    End With
End Sub

Private Sub FillData()
    Dim i As Double, j As Double
    With grd
        For i = 0 To 15
            For j = 0 To 16
                '.CellText(i + 1, j + 1) = CStr((i * 16) + j - 1)
                If (i * 16) + j - 1 > 0 Then
                    .CellText(i + 1, j + 1) = ascii_table((i * 16) + j - 1).Char
                End If
            Next j
        Next i
    End With
End Sub

Private Function fillASCII()
    Dim i As Double
    ' ***** num
    For i = 0 To 255
        ascii_table(i).num = i
    Next i
    ' ***** printable type or control type
    For i = 0 To 31
        ascii_table(i).type = "Control char"
    Next
    For i = 32 To 255
        ascii_table(i).type = "Printable"
    Next
    ascii_table(127).type = "Control char"
    ascii_table(129).type = "Control char"
    ascii_table(141).type = "Control char"
    ascii_table(143).type = "Control char"
    ascii_table(144).type = "Control char"
    ascii_table(157).type = "Control char"
    ' ***** characters
    For i = 0 To 255
        ascii_table(i).Char = Chr(i)
    Next i
    ' ***** desc
    For i = 0 To 255
        ascii_table(i).desc = CStr(Chr(i))
    Next i
    ascii_table(0).desc = "Null char"
    ascii_table(1).desc = "Header start"
    ascii_table(2).desc = "Start of text"
    ascii_table(3).desc = "End of text"
    ascii_table(4).desc = "End of transmision"
    ascii_table(5).desc = "Enquiry"
    ascii_table(6).desc = "Acknowledgement"
    ascii_table(7).desc = "Bell"
    ascii_table(8).desc = "Backstep"
    ascii_table(9).desc = "Horizontal tabulation"
    ascii_table(10).desc = "Line feed (LF)"
    ascii_table(11).desc = "Vertical tabulation"
    ascii_table(12).desc = "Form feed"
    ascii_table(13).desc = "Carriage return (CR)"
    ascii_table(14).desc = "Shift Out"
    ascii_table(15).desc = "Shift In"
    ascii_table(16).desc = "Data Link Escape"
    ascii_table(17).desc = "Device Control 1 ó oft. XON"
    ascii_table(18).desc = "Device Control 2"
    ascii_table(19).desc = "Device Control 3 ó oft. XOFF"
    ascii_table(20).desc = "Device Control 4"
    ascii_table(21).desc = "Negative Acknowledgement"
    ascii_table(22).desc = "Synchronous Idle"
    ascii_table(23).desc = "End of Trans. Block"
    ascii_table(24).desc = "Cancel"
    ascii_table(25).desc = "End of Medium"
    ascii_table(26).desc = "Substitute"
    ascii_table(27).desc = "Escape"
    ascii_table(28).desc = "File Separator"
    ascii_table(29).desc = "Group Separator"
    ascii_table(30).desc = "Record Separator"
    ascii_table(31).desc = "Unit Separator"
    ascii_table(32).desc = "Space"
    ascii_table(127).desc = "Delete"
End Function

'  ascii_taula[256]="","","","","","","","","    "," ","",""," ","","","",    //16
'              "","","","","","","","","",""," ","","","",""," ",    //32
'              "!","Æ","#","$","%","É","Æ","(",")","*","+",",","-",".","/","0",    //48
'              "1","2","3","4","5","6","7","8","9",":",";","<","=",">","?","@",    //64
'              "A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P",    //80
'              "Q","R","S","T","U","V","W","X","Y","Z","[","\","]","^","_","`",    //96
'              "a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p",    //112
'              "q","r","s","t","u","v","w","x","y","z","{","|","}","~","","Ä",    //128
'              "Å","Ç","É","Ñ","Ö","Ü","á","à","â","ä","ã","å","ç","é","è","ê",    //144
'              "ë","í","ì","î","ï","ñ","ó","ò","ô","ö","õ","ú","ù","û","ü","†",    //160
'              "°","¢","£","§","•","¶","ß","®","©","™","´","¨","≠","Æ","Ø","∞",    //176
'              "±","≤","≥","¥","µ","∂","∑","∏","π","∫","ª","º","Ω","æ","ø","¿",    //192
'              "¡","¬","√","ƒ","≈","∆","«","»","…"," ","À","Ã","Õ","Œ","œ","–",    //208
'              "—","“","”","‘","’","÷","◊","ÿ","Ÿ","⁄","€","‹","›","ﬁ","ﬂ","‡",    //224
'              "·","‚","„","‰","Â","Ê","Á","Ë","È","Í","Î","Ï","Ì","Ó","Ô","",    //240
'              "Ò","Ú","Û","Ù","ı","ˆ","˜","¯","˘","˙","˚","¸","˝","˛","ˇ"," ";    //256


Private Sub grd_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    If ascii_table(grd.SelectedRow * 16 + lCol - 18).type = "Printable" Then
        'insert char
        insertChar ascii_table(grd.SelectedRow * 16 + lCol - 18).Char
    End If
End Sub

Private Sub grd_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
    'MsgBox "click at " & ascii_table(grd.SelectedRow * 16 + lCol - 18).char
    With ascii_table(grd.SelectedRow * 16 + lCol - 18)
        lblDesc.Caption = " " & .num & ": " & .desc & " (" & .type & ")"
    End With
End Sub
