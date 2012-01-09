VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Flamebird MX"
   ClientHeight    =   7710
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   7785
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   514
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   519
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":000C
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0F0F0&
      Height          =   375
      Left            =   2520
      TabIndex        =   24
      Top             =   5880
      Width           =   4935
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "And all other people who has participated directly or indirectly"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0F0F0&
      Height          =   195
      Left            =   2520
      TabIndex        =   23
      Top             =   4440
      Width           =   4830
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Special thanks"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BDBDBD&
      Height          =   255
      Left            =   930
      TabIndex        =   22
      Top             =   3720
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dependencies"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BDBDBD&
      Height          =   255
      Left            =   960
      TabIndex        =   21
      Top             =   4920
      Width           =   1260
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GINO"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0F0F0&
      Height          =   195
      Left            =   2520
      TabIndex        =   20
      Top             =   4200
      Width           =   420
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sebastian Quiest"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0F0F0&
      Height          =   195
      Left            =   2520
      TabIndex        =   19
      Top             =   3960
      Width           =   1305
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Coptroner"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0F0F0&
      Height          =   195
      Left            =   2520
      TabIndex        =   18
      Top             =   3720
      Width           =   795
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Original FB2 Team"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BDBDBD&
      Height          =   240
      Left            =   135
      TabIndex        =   17
      Top             =   3240
      Width           =   2085
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "JaViS, Danko, Viator, BlueSteel"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0F0F0&
      Height          =   195
      Left            =   2520
      TabIndex        =   16
      Top             =   3240
      Width           =   2355
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Imanol Zubiaurre"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0F0F0&
      Height          =   195
      Left            =   2520
      TabIndex        =   15
      Top             =   2760
      Width           =   1320
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "izubiaurre@users.sourceforge.net"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   195
      Left            =   4710
      MouseIcon       =   "frmAbout.frx":00A4
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   2760
      Width           =   2625
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mantained by"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BDBDBD&
      Height          =   480
      Index           =   1
      Left            =   675
      TabIndex        =   13
      Top             =   2520
      Width           =   1545
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.sourceforge.net/projects/fbtwo"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   195
      Left            =   3960
      MouseIcon       =   "frmAbout.frx":0210
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   1560
      Width           =   3375
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Project page:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0F0F0&
      Height          =   195
      Left            =   2520
      TabIndex        =   11
      Top             =   1560
      Width           =   1005
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Darío Cutillas (Danko)"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0F0F0&
      Height          =   195
      Left            =   2520
      TabIndex        =   10
      Top             =   2520
      Width           =   1680
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Javier Arias (JaViS)"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0F0F0&
      Height          =   195
      Left            =   2520
      TabIndex        =   9
      Top             =   2040
      Width           =   1365
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "javis@users.sourceforge.net"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   195
      Left            =   5160
      MouseIcon       =   "frmAbout.frx":037C
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lord_danko@users.sourceforge.net"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   195
      Left            =   4605
      MouseIcon       =   "frmAbout.frx":04E8
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   2520
      Width           =   2730
   End
   Begin VB.Label lblDev 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FB2 Original idea"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BDBDBD&
      Height          =   255
      Left            =   690
      TabIndex        =   6
      Top             =   2040
      Width           =   1530
   End
   Begin VB.Label lblOk 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[OK]"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0F0F0&
      Height          =   315
      Index           =   1
      Left            =   6360
      TabIndex        =   5
      Top             =   7320
      Width           =   480
   End
   Begin VB.Label lblLin 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[LICENSE]"
      BeginProperty Font 
         Name            =   "Segoe UI Semibold"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0F0F0&
      Height          =   315
      Index           =   2
      Left            =   2880
      TabIndex        =   4
      Top             =   7320
      Width           =   1050
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Code Sense is based on the popular CodeMax control © 2000 WinMain Software."
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0F0F0&
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   5400
      Width           =   4575
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "The TabDock control used in this application was created by Marclei V Silva. http://www.spnorte.com"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0F0F0&
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   6360
      Width           =   4695
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "This product includes software developed by vbAccelerator (http://vbaccelerator.com/)."""
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F0F0F0&
      Height          =   495
      Left            =   2520
      TabIndex        =   1
      Top             =   4920
      Width           =   5055
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.1.00"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5745
      TabIndex        =   0
      Top             =   585
      Width           =   450
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00333333&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00333333&
      FillColor       =   &H00333333&
      Height          =   6495
      Left            =   0
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00666666&
      BorderColor     =   &H00666666&
      FillColor       =   &H00666666&
      FillStyle       =   0  'Solid
      Height          =   6495
      Left            =   2400
      Top             =   1200
      Width           =   5415
   End
End
Attribute VB_Name = "frmAbout"
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

Private Sub cmdOk_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    frmLicense.Show 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        cmdOk_Click
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblVersion.Caption = App.Major & "." & App.Minor & "." & App.Revision
    Picture = LoadPicture(App.Path & "/Resources/frmAbout.jpg")
End Sub

Private Sub Label10_Click()
    Unload Me
    NewWindowWeb "http://www.sourceforge.net/projects/fbtwo", "Flamebird mx project page"
End Sub

Private Sub Label11_Click()
    Clipboard.SetText "izubiaurre@users.sourceforge.net"
End Sub

Private Sub Label3_Click()
    Clipboard.SetText "lord_danko@users.sourceforge.net"
End Sub

Private Sub Label5_Click()
    Clipboard.SetText "javis@users.sourceforge.net"
End Sub

Private Sub lblLin_Click(Index As Integer)
    frmLicense.Show 1
End Sub

Private Sub lblOk_Click(Index As Integer)
    Unload Me
End Sub

