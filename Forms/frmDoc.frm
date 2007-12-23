VERSION 5.00
Object = "{665BF2B8-F41F-4EF4-A8D0-303FBFFC475E}#2.0#0"; "cmcs21.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbaliml6.ocx"
Object = "{E142732F-A852-11D4-B06C-00500427A693}#1.14#0"; "vbaltbar6.ocx"
Begin VB.Form frmDoc 
   Caption         =   "Form1"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5580
   ControlBox      =   0   'False
   Icon            =   "frmDoc.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3780
   ScaleWidth      =   5580
   WindowState     =   2  'Maximized
   Begin vbalTBar6.cReBar ReBar 
      Left            =   3120
      Top             =   120
      _ExtentX        =   1508
      _ExtentY        =   661
   End
   Begin vbalTBar6.cToolbar tbrSource 
      Height          =   375
      Left            =   480
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
   End
   Begin vbalIml6.vbalImageList ilSource 
      Left            =   1920
      Top             =   2640
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   16
      Size            =   9184
      Images          =   "frmDoc.frx":2B8A
      Version         =   131072
      KeyCount        =   8
      Keys            =   "ÿÿÿÿÿÿÿ"
   End
   Begin CodeSenseCtl.CodeSense cs 
      Height          =   3135
      Left            =   0
      OleObjectBlob   =   "frmDoc.frx":4F8A
      TabIndex        =   0
      Top             =   600
      Width           =   5535
   End
End
Attribute VB_Name = "frmDoc"
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
Option Base 1 'IMPORTANT!!!! TO CHECK LATER

'MSG Constants (for future multi-language support)
Private Const MSG_SAVE_FILEREADONLY = "This File is read-only. You must save to a different location."
Private Const MSG_SAVE_ERRORSAVING = "An error occurred when trying to save the file: "
Private Const MSG_SAVE_SUCCESS = "File saved succesfully!"
Private Const MSG_COMPILE_NOFENIXDIR = "Fenix directory has not been configured or does not exist"
Private Const MSG_COMPILE_NOTALREADYSAVED = "The file has not been saved yet. Save the file before compile"
Private Const MSG_RUN_DBCNOTFOUND = "DCB file not found. Compile first!"


Dim selRange As CodeSenseCtl.IRange
Private nextTipText() As String
Public rangoActual As CodeSense.range
Public mustRefresh As Boolean
Dim showingList As CodeSenseCtl.ICodeList 'determina si se esta mostrando la lista de autocompletado
Dim showingToolTip As CodeSenseCtl.ICodeTip

Private m_FilePath As String
Private m_IsDirty As Boolean 'This should be never set directly (use the IsDirty property)
Private m_Title As String 'Basicaly the caption of the form
Private m_addToProject As Boolean

Implements IFileForm

Private Sub Cs_Change(ByVal Control As CodeSenseCtl.ICodeSense)
    On Error Resume Next
    
    IsDirty = True
    '*********************************************************************************
    'analiza para determinar si hace falta refrescar el arbol de declaraciones
    '*********************************************************************************
    If (Not rangoActual Is Nothing) And mustRefresh = False Then
    
        'si esta en zona de declaracion (obvio que en este evento nos damos cuenta de que se cambio algo)
        If inDeclarationZone(rangoActual.StartLineNo) = True Then
        
            Dim token As CodeSenseCtl.cmTokenType
        
            ' We don't want to display a tip inside quoted or commented-out lines...
            token = Control.CurrentToken
        
            If ((cmTokenTypeText = token) Or (cmTokenTypeKeyword = token)) Then
                mustRefresh = True
            End If
            
        End If
    End If

End Sub

Private Function Cs_CodeList(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList) As Boolean
    Dim i As Long
    'agrega items a la lista de autocompletado
    For i = 1 To UBound(functionList)
        ListCtrl.AddItem functionList(i)
    Next i
    
    'agrega items a la lista de autocompletado
    For i = 1 To UBound(varList)
        ListCtrl.AddItem varList(i)
    Next i
    
    Dim wordIndex As Long
    
    wordIndex = ListCtrl.FindString(cs.CurrentWord)
    'si no se encuentra la palabra cancela
    If wordIndex = -1 Then
        ListCtrl.Destroy
        Cs_CodeList = False
        Exit Function
    End If
    
    ' Just for kicks, we'll select the first item by default...
    ListCtrl.SelectedItem = wordIndex
    
    ' Enable mouse hot-tracking
    ListCtrl.EnableHotTracking
    
    'define a la variable que contiene la lista que se esta mostrando
    Set showingList = ListCtrl
    
    ' Allow list view control to be displayed
    Cs_CodeList = True
    
End Function

Private Function Cs_CodeListCancel(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList) As Boolean
    Set showingList = Nothing
End Function

Private Function Cs_CodeListSelChange(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList, ByVal lItem As Long) As String
    ' Set the tooltip text...
    'Cs_CodeListSelChange = "This is function #" & lItem + 1
End Function

Private Function Cs_CodeListSelMade(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList) As Boolean
    Dim strItem As String
    Dim range As New CodeSenseCtl.range
    
    ' Determine which item was selected in the list
    strItem = ListCtrl.GetItemText(ListCtrl.SelectedItem)
    
    ' Get actual  selection
    Set range = cs.GetSel(True)
    
    'si no esta seleccionado mas de una linea
    If range.StartLineNo = range.EndLineNo And range.StartColNo = range.EndColNo Then
        'calcula la longitud de la nueva seleccion
        range.EndColNo = range.StartColNo
        range.StartColNo = range.StartColNo - Len(cs.CurrentWord)
    End If
    
    'selecciona la palabra que en la que este actualmente posicionado
    cs.SetSel range, False
    
    ' Replace current selection
    cs.ReplaceSel (strItem)

    ' Get new selection
    Set range = cs.GetSel(True)

    ' Update range to end of newly inserted text
    range.StartColNo = range.StartColNo + Len(strItem)
    range.EndColNo = range.StartColNo
    range.EndLineNo = range.StartLineNo

    ' Move cursor
    cs.SetSel range, True
    
    Set showingList = Nothing

    ' Don't prevent list view control from being hidden
    Cs_CodeListSelMade = False
    
End Function

Private Function Cs_CodeListSelWord(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList, ByVal lItem As Long) As Boolean
    ' Allow the CodeList control to automatically select the item in the
    ' list that most closely matches the current word.
    Cs_CodeListSelWord = True
End Function

Private Function Cs_CodeTip(ByVal Control As CodeSenseCtl.ICodeSense) As CodeSenseCtl.cmToolTipType
 Dim token As CodeSenseCtl.cmTokenType

    ' We don't want to display a tip inside quoted or commented-out lines...
    token = Control.CurrentToken

    'If ((cmTokenTypeText = token) Or (cmTokenTypeKeyword = token)) Then
        ' We want to use the tooltip control that automatically
        ' highlights the arguments in the function prototypes for us.
        ' This type also provides support for overloaded function
        ' prototypes.
        If UBound(nextTipText) > 1 Then
            Cs_CodeTip = cmToolTipTypeMultiFunc
        Else
            Cs_CodeTip = cmToolTipTypeFuncHighlight
        End If
    'Else
        ' Don't display a tooltip
        'Cs_CodeTip = cmToolTipTypeNone
    'End If

End Function

Private Function Cs_CodeTipCancel(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ToolTipCtrl As CodeSenseCtl.ICodeTip) As Boolean
Set showingToolTip = Nothing
End Function

Private Sub Cs_CodeTipInitialize(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ToolTipCtrl As CodeSenseCtl.ICodeTip)
    Dim tip As Variant
'    If UBound(nextTipText) > 1 Then
'        Dim tip As CodeSenseCtl.CodeTipMultiFunc
'    Else
'        Dim tip As CodeSenseCtl.CodeTipFuncHighlight
'    End If
    
    Set tip = ToolTipCtrl
    
    Set showingToolTip = ToolTipCtrl
       
    ' Default to first argument
    tip.Argument = 0

    ' Save the starting position for use with the CodeTip.  This is so we
    ' can destroy the tip window if the user moves the cursor back before
    ' or above the starting point.
    '
    Set selRange = Control.GetSel(True)
    selRange.EndColNo = selRange.EndColNo + 1

    If UBound(nextTipText) > 1 Then
        ' Set number of function overloads
        tip.FunctionCount = UBound(nextTipText) - 1

        ' Default to first function prototype
        tip.CurrentFunction = 0
        ' Set tooltip text to first function prototype
        tip.TipText = nextTipText(tip.CurrentFunction + 1)
    Else
        ' Set tooltip text to first function prototype
        tip.TipText = nextTipText(1)
    End If
    
End Sub

Private Sub Cs_CodeTipUpdate(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ToolTipCtrl As CodeSenseCtl.ICodeTip)
    Dim tip As Variant
    
    Set tip = ToolTipCtrl

    ' Destroy the tooltip window if the caret is moved above or before
    ' the starting point.
    Dim range As CodeSenseCtl.IRange
    Set range = Control.GetSel(True)

    If ((range.EndLineNo < selRange.EndLineNo) Or _
        ((range.EndColNo < selRange.EndColNo - 1) And _
        (range.EndLineNo <= selRange.EndLineNo))) Then

        ' Caret moved too far up / back
        tip.Destroy
        
        Set showingToolTip = Nothing
    Else

        ' Determine which argument to highlight
        Dim iArg, i As Integer
        Dim strLine As String

        iArg = 0
        i = selRange.EndLineNo

        While ((i <= range.EndLineNo) And (iArg <> -1))

            'Get next line from buffer
            strLine = Control.getLine(i)

            If (i = range.EndLineNo) Then
                ' Trim off any excess beyond cursor pos so the argument with the
                ' cursor in it will be highlighted.
                Dim iTrim As Variant
                iTrim = Len(strLine) + 1
                If (range.EndColNo < iTrim) Then
                    iTrim = range.EndColNo
                End If
                strLine = Left(strLine, iTrim)
            End If

            ' Parse arguments from current line
            Dim j As Integer
            j = 0
            While ((Len(strLine) <> 0) And (j <= Len(strLine)) And (iArg <> -1))
                If (Mid(strLine, j + 1, 1) = ",") Then
                    iArg = iArg + 1
                ElseIf (Mid(strLine, j + 1, 1) = ")") And j + 1 = Len(strLine) Then
                    iArg = -1
                End If
                j = j + 1
            Wend

            i = i + 1
        Wend
        

        If (-1 = iArg) Then
            tip.Destroy 'Right-paren found
            Set showingToolTip = Nothing
        Else
            tip.Argument = iArg
            
            ' Set tooltip text to current function prototype
                        
            If UBound(nextTipText) > 1 Then
                tip.TipText = nextTipText(tip.CurrentFunction + 1)
            Else
                tip.TipText = nextTipText(1)
            End If

        End If

    End If

End Sub

'************************************************************************
    'Devuelve el indice de la string en la lista variables declaradas
    'por el usuario
    'si no lo encuentra devuelve -1
    'pero se supone que nunca devuelve 0 XD
'************************************************************************
Private Function indexOnVarList(Word As String)
    Dim wordIndex As Long
    Dim i As Long
            
        wordIndex = -1
        
        For i = 1 To UBound(varList)
            If InStr(LCase(varList(i)), LCase(Word)) > 0 Then
                wordIndex = i
                Exit For
            End If
        Next i
    
    indexOnVarList = wordIndex
End Function


'************************************************************************
    'Devuelve el indice de la string en la lista de funciones
    'si no lo encuentra devuelve -1
    'pero se supone que nunca devuelve 0 XD
'************************************************************************
Private Function indexOnFunctionList(Word As String)
Dim wordIndex As Long
Dim i As Long
        
    wordIndex = -1
    
    For i = 1 To UBound(functionList)
        If InStr(LCase(functionList(i)), LCase(Word)) > 0 Then
            wordIndex = i
            Exit For
        End If
    Next i

    indexOnFunctionList = wordIndex
End Function

'Esta funcion trae el nombre de la funcion que esta dentro de la linea actual
Private Function getCurrentFunction() As String
    
    'primero limita la linea en donde se va a buscar
    Dim linea As String
    
    linea = cs.getLine(rangoActual.EndLineNo)
    
    'corta la linea en donde se ubica el cursor
    linea = Mid(linea, 1, rangoActual.EndColNo)
    
    '*******************************************
    '*********** OPTIMIZACION DE LA LINEA ******
    '*******************************************
        
    'reemplaza los caracteres no visibles por espacios
    linea = replace(linea, Chr(9), " ")
    linea = replace(linea, vbNewLine, " ")
    linea = replace(linea, vbCrLf, " ")
    linea = replace(linea, vbCr, " ")
    linea = replace(linea, vbLf, " ")
    linea = replace(linea, vbNullChar, " ")
    linea = replace(linea, vbBack, " ")
    linea = replace(linea, vbFormFeed, " ")
    linea = replace(linea, vbVerticalTab, " ")
                
    'se fija si hay funciones cerradas antes
    While InStrRev(linea, ")")
        linea = Left(linea, InStrRev(linea, ")") - 1)
        If InStr(linea, "(") Then
            linea = Mid(linea, 1, InStrRev(linea, "(") - 1)
        End If
    Wend
            
    'ahora corta la linea donde encuentre el ultimo "("
    If InStr(linea, "(") Then
        linea = Mid(linea, 1, InStrRev(linea, "(") - 1)
    Else
        'si no hay ( es que no hay  funcion
        linea = ""
        Exit Function
    End If
    
    ' toma la ultima palabra de la linea
    linea = getWordRev(linea)
    
    getCurrentFunction = linea
    
End Function



Private Function Cs_KeyPress(ByVal Control As CodeSenseCtl.ICodeSense, ByVal KeyAscii As Long, ByVal Shift As Long) As Boolean
On Error Resume Next

Dim i As Integer
    Dim token As CodeSenseCtl.cmTokenType
 
    'codigo para mostar el list de autocompletado
    ' We don't want to display a tip inside quoted or commented-out lines...
    token = Control.CurrentToken
    If ((cmTokenTypeText = token) Or (cmTokenTypeKeyword = token)) Then
    
        'si no se esta mostrando manda a crearla si se cumplen los requisitos
        If cs.CurrentWordLength > 2 Then
            If showingList Is Nothing And KeyAscii >= 65 And KeyAscii <= 122 And (indexOnFunctionList(cs.CurrentWord & Chr(KeyAscii)) > 0 Or indexOnVarList(cs.CurrentWord & Chr(KeyAscii)) > 0) Then
                cs.ExecuteCmd cmCmdCodeList
            End If
        End If
        'si se esta mostrando la lista de autompletado
        If Not showingList Is Nothing Then
            ' la destruye en los siguientes casos
            
            If KeyAscii < 65 Or KeyAscii > 122 Then
                showingList.Destroy
                Set showingList = Nothing
            End If
            
            If Not showingList Is Nothing Then
                'si no se encuentra la palabra cancela
                If indexOnFunctionList(cs.CurrentWord & Chr(KeyAscii)) = -1 And indexOnVarList(cs.CurrentWord & Chr(KeyAscii)) = -1 Then
                    showingList.Destroy
                    Set showingList = Nothing
                End If
            End If
            
        End If
        
        Dim funcion As String

        'codigo para mostrar el tooltip
        If (Asc("(") = KeyAscii) Then
            Dim linea As String
            ' toma la linea donde esta posicionado el cursor
            linea = Trim(cs.getLine(rangoActual.EndLineNo))
            ' la corta en donde esta el cursor
            linea = Mid(linea, 1, rangoActual.StartColNo)
            
            '*******************************************
            '*********** OPTIMIZACION DE LA LINEA ******
            '*******************************************
                
            'reemplaza los caracteres no visibles por espacios
            linea = replace(linea, Chr(9), " ")
            linea = replace(linea, vbNewLine, " ")
            linea = replace(linea, vbCrLf, " ")
            linea = replace(linea, vbCr, " ")
            linea = replace(linea, vbLf, " ")
            linea = replace(linea, vbNullChar, " ")
            linea = replace(linea, vbBack, " ")
            linea = replace(linea, vbFormFeed, " ")
            linea = replace(linea, vbVerticalTab, " ")
            
            showToolTip getWordRev(linea)
        End If
        
        
    Else
        'muestra el tooltip cuando se aprieta "," y no se estaba mostrando
        If (Asc(",") = KeyAscii) Then
            funcion = getCurrentFunction
            If funcion <> "" Then
                showToolTip funcion
            End If
        End If
    End If
End Function

'muestra el tooltip para la palabra del parametro
Private Sub showToolTip(palabra As String)
    Dim prototipo() As String
    Dim cantidad As String
    Dim i As Integer
    
    If indexOnFunctionList(palabra) > 0 And indexOnFunctionList(palabra) <= iFunctionCount Then
        With Ini
            .Path = App.Path & "\Help\functions.lan"
            .Section = palabra
            
            .Key = "QOfPrototipes"
            .Default = "0"
            cantidad = CInt(.value)
            
            If cantidad > 0 Then
                ReDim prototipo(1 To cantidad) As String
                
                                  
                i = 1
                While i <= cantidad
                    If i = 1 Then
                        .Key = "Prototipe"
                    Else
                        .Key = "Prototipe" & i
                    End If
                    .Default = ""
                    prototipo(i) = .value
                    i = i + 1
                Wend
            End If
            
        End With
               
        If cantidad > 0 Then
            ReDim nextTipText(1 To cantidad) As String
            i = 1
            While i <= cantidad
                nextTipText(i) = prototipo(i)
                i = i + 1
            Wend
            Me.cs.ExecuteCmd cmCmdCodeTip
        End If
    End If
    
    'si esta dentro de la lista pero es declarada por el usuario
    If indexOnFunctionList(palabra) > iFunctionCount Then
        ReDim nextTipText(1) As String
        nextTipText(1) = palabra & " (" & parameters(indexOnFunctionList(palabra) - iFunctionCount) & ")"
        Me.cs.ExecuteCmd cmCmdCodeTip
    End If
End Sub

Private Function Cs_KeyUp(ByVal Control As CodeSenseCtl.ICodeSense, ByVal KeyCode As Long, ByVal Shift As Long) As Boolean

    'mas comprovaciones para la lista de autocompletado
    If Not showingList Is Nothing Then
        If cs.CurrentWordLength < 3 Then
            showingList.Destroy
            Set showingList = Nothing
        End If
    End If
    
End Function

Private Function Cs_MouseDown(ByVal Control As CodeSenseCtl.ICodeSense, ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long) As Boolean
    On Error Resume Next
    If (Button = 2) Then
        frmMain.cMenu.PopupMenu "mnuEdit"
        'contextMenu.ShowPopupMenu X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY
    End If
End Function

Private Function Cs_RClick(ByVal Control As CodeSenseCtl.ICodeSense) As Boolean
    Cs_RClick = True
End Function

Private Sub Cs_RegisteredCmd(ByVal Control As CodeSenseCtl.ICodeSense, ByVal lCmd As CodeSenseCtl.cmCommand)
    ' comando registrado para mostrar ayuda
    If lCmd = 1000 Then
        Dim sword As String
        sword = cs.CurrentWord
        If sword <> "" Then
            ' ayuda local
            NewWindowWeb App.Path & "\help\fenix\func.php-func=" & UCase(sword) & ".htm", "HELP: " & UCase(sword), App.Path & "\help\fenix\func.php-frame=top.htm"
            
            ' ayuda en inet
            'NewWindowWeb "http://fenix.jlceb.com/func.php?func=" & UCase(sword)
        End If
    End If
End Sub

Private Sub Cs_SelChange(ByVal Control As CodeSenseCtl.ICodeSense)
    On Error Resume Next
    
    cs.HighlightedLine = -1
    
    Dim rangoTemp As CodeSense.range
    Set rangoTemp = cs.GetSel(True)
    
    If Not rangoActual Is Nothing Then
        'detectamos cambio de linea
        If rangoTemp.StartLineNo <> rangoActual.StartLineNo Then
            'esta variable indica si se debe refrescar al cambiar de linea
            If mustRefresh = True Then
                MakeProgramIndex IFileForm_FilePath
                mustRefresh = False
            End If
            
            ' si se estaba mostrando el tooltip lo eliminamos
            If Not showingToolTip Is Nothing Then
                showingToolTip.Destroy
                Set showingToolTip = Nothing
            End If
        End If
    End If
    
    Set rangoActual = cs.GetSel(True) 'ubica la posicion actual y la guarda en una var de alcance modular
End Sub

Private Sub Form_Activate()
    ' testeando
    If existTreeForFile(IFileForm_FilePath) = False Then
        mustRefresh = True
    End If
    MakeProgramIndex IFileForm_FilePath
    mustRefresh = False
End Sub

Private Sub Form_Load()
    On Error GoTo errhandler:
    
    With tbrSource
        .ImageSource = CTBExternalImageList
        .DrawStyle = T_Style
        .SetImageList ilSource.hIml, CTBImageListNormal
        .CreateToolbar 16, False, True, True, 16
        .AddButton "Toogle bookmark", 0, , , , CTBAutoSize, "ToogleBookmark"
        .AddButton "Next bookmark", 1, , , , CTBAutoSize, "NextBookmark"
        .AddButton "Previous bookmark", 2, , , , CTBAutoSize, "PreviousBookmark"
        .AddButton "Delete all bookmarks", 3, , , , CTBAutoSize, "DeleteBookmarks"
        .AddButton eButtonStyle:=CTBSeparator
        .AddButton "Shift right", 4, , , , CTBAutoSize, "ShiftRight"
        .AddButton "Shift left", 5, , , , CTBAutoSize, "ShiftLeft"
        .AddButton eButtonStyle:=CTBSeparator
        .AddButton "Comment", 6, , , , CTBAutoSize, "Comment"
        .AddButton "Uncomment", 7, , , , CTBAutoSize, "Uncomment"
    End With
    
    'Create the rebar
    With Rebar
        If A_Bitmaps Then
            .BackgroundBitmap = App.Path & "\resources\backrebar.bmp"
        End If
        .CreateRebar Me.hwnd
        .AddBandByHwnd tbrSource.hwnd, , True, False
    End With
    Rebar.RebarSize
    
    'cofigura el control de edicion
    cs.LineNumbering = True
    cs.LineNumberStart = 1
    cs.LineNumberStyle = cmDecimal
    cs.LineToolTips = True
    cs.BorderStyle = cmBorderStatic
    cs.EnableDragDrop = True
    cs.EnableCRLF = True
    cs.TabSize = 2
    cs.ColorSyntax = True
    cs.Language = "Fenix"
    cs.DisplayLeftMargin = True
    cs.AutoIndentMode = cmIndentPrevLine
    LoadCSConf cs
    
    ' registra comando para mostrar ayuda
    ' se ejecuta en el evento Cs_RegisteredCmd
    Dim g As New CodeSenseCtl.Globals
    Call g.RegisterCommand(1000, "ShowHelp", "Shows the help about the current word in the control.")
    'registra hotkey
    Dim h As New CodeSenseCtl.HotKey
    h.VirtKey1 = Chr(vbKeyF1)   'F8
    Call g.RegisterHotKey(h, 1000)

    Exit Sub
errhandler:
    If Err.Number > 0 Then ShowError ("frmDoc.FormLoad")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim msgRes As VbMsgBoxResult
    'Ask for saving if the document is dirty and we are not closing the entire application
    If IFileForm_IsDirty = True And UnloadMode <> vbFormMDIForm Then
        msgRes = MsgBox("The file '" & IFileForm_Title & "' is modified. " _
                    & "Save it?", vbYesNoCancel + vbQuestion, "Save")
        If msgRes = vbYes Then 'Save
            SaveFileOfFileForm Me
        ElseIf msgRes = vbCancel Then 'Cancel
            Cancel = 1
        End If
    End If
End Sub

Private Sub Form_Resize()
    If frmMain.WindowState <> vbMinimized Then
        Rebar.RebarSize
        cs.Move 0, ScaleY(Rebar.RebarHeight, vbPixels, vbTwips)
        cs.Width = Me.ScaleWidth
        cs.Height = Me.ScaleHeight - ScaleY(Rebar.RebarHeight, vbPixels, vbTwips)
    End If
End Sub
    
Private Sub Form_Terminate()
    'si no hay prg abiertos limpia el tree de declaraciones
    If DocExist() = False Then
        frmProgramInspector.tv_program.Nodes.Clear
    End If
End Sub

'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'STARTS INTERFACE IFILEFORM
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
Private Property Get IFileForm_AlreadySaved() As Boolean
    IFileForm_AlreadySaved = IIf(m_FilePath = "", False, True)
End Property

Private Function IFileForm_CloseW() As Long

End Function

Private Property Get IFileForm_FileName() As String
    IFileForm_FileName = FSO.GetFileName(m_FilePath)
End Property

Private Property Get IFileForm_FilePath() As String
    IFileForm_FilePath = m_FilePath
End Property

Private Function IFileForm_Identify() As EFileFormConstants
    IFileForm_Identify = FF_SOURCE
End Function

Private Property Get IFileForm_IsDirty() As Boolean
    IFileForm_IsDirty = m_IsDirty
End Property

Private Function IFileForm_Load(ByVal sFile As String) As Long
    Dim lResult As Long
    
    cs.OpenFile (sFile)
    'There is no way to determine if the cs.openfile fails so assume everything goes well
    'since we check for the existence of the file in the NewFileForm function this should work
    'well (any file is supposed to be able to be opened in text format)
    lResult = -1
    m_FilePath = sFile
    IsDirty = False
    
    IFileForm_Load = lResult
End Function

Private Property Get IFileForm_Title() As String
    Dim sTitle As String
    If IFileForm_FilePath = "" Then
        sTitle = m_Title
    Else
        sTitle = IFileForm_FileName
    End If
    IFileForm_Title = sTitle
End Property

Private Function IFileForm_NewW(ByVal iUntitledCount As Integer) As Long
    'Nothing special is needed, maby add some starting code
    m_Title = "Untitled " & iUntitledCount
    IsDirty = False
    m_addToProject = modMenuActions.NewAddToProject
    IFileForm_NewW = -1 'Succesful
End Function

Private Function IFileForm_Save(ByVal sFile As String) As Long
    Dim lResult As Long
    
    If FSO.FileExists(sFile) Then Kill sFile 'Delete the file if exists
    
    On Error GoTo errhandler
    cs.SaveFile sFile, False 'Save the file
    'HERE THERE SHOULD BE SOME KIND OF COMPROBATION FOR ERRORS AFTER SAVEFILE
    lResult = -1
    If (lResult) Then 'Save succesful
        'Add to project if necessary
        If IFileForm_AlreadySaved = False And m_addToProject = True Then addFileToProject sFile
    
        If m_FilePath <> sFile Then 'Show a success message only if the name changed
            MsgBox MSG_SAVE_SUCCESS, vbInformation
        End If
        m_FilePath = sFile
        IsDirty = False
    Else
        MsgBox MSG_SAVE_ERRORSAVING, vbCritical
    End If

errhandler:
    If Err.Number > 0 Then lResult = -1: Resume Next
    
End Function
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'END INTERFACE IFILEFORM
'·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-·-'
'This property is not part of the interface, just helps to reduce code
'by setting the caption of the form properly
Private Property Let IsDirty(ByVal newVal As Boolean)
    m_IsDirty = newVal
    'Put an * to the caption if dirty
    Caption = IFileForm_Title & IIf(newVal, " *", "")
    
    frmMain.RefreshTabs
End Property

Private Sub tbrSource_ButtonClick(ByVal lButton As Long)
    Dim sKey As String
    
    sKey = tbrSource.ButtonKey(lButton)
    Select Case sKey
    Case "ToogleBookmark"
        modMenuActions.mnuBookmarkToggle
    Case "NextBookmark"
        modMenuActions.mnuBookmarkNext
    Case "PreviousBookmark"
        modMenuActions.mnuBookmarkPrev
    Case "DeleteBookmarks"
        modMenuActions.mnuBookmarkDel
    Case "ShiftRight"
        modMenuActions.mnuAdvancedTab
    Case "ShiftLeft"
        modMenuActions.mnuAdvancedUnTab
    Case "Comment"
        modMenuActions.mnuAdvancedComment
    Case "Uncomment"
        modMenuActions.mnuAdvancedUnComment
    End Select
        
End Sub
