VERSION 5.00
Object = "{665BF2B8-F41F-4EF4-A8D0-303FBFFC475E}#2.0#0"; "cmcs21.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{E142732F-A852-11D4-B06C-00500427A693}#1.14#0"; "vbalTbar6.ocx"
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
   Begin vbalIml6.vbalImageList ilSource2 
      Left            =   4440
      Top             =   2640
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   32
      Size            =   12628
      Images          =   "frmDoc.frx":2B8A
      Version         =   131072
      KeyCount        =   11
      Keys            =   "����������"
   End
   Begin VB.ComboBox cmbBookmarkList 
      Height          =   315
      Left            =   3600
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin vbalTBar6.cReBar ReBar 
      Left            =   2520
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
      Size            =   12628
      Images          =   "frmDoc.frx":5CFE
      Version         =   131072
      KeyCount        =   11
      Keys            =   "����������"
   End
   Begin CodeSenseCtl.CodeSense cs 
      Height          =   3135
      Left            =   0
      OleObjectBlob   =   "frmDoc.frx":8E72
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
'   JaViS:      javisarias@ gmail.com            (JaViS)
'   Danko:      lord_danko@users.sourceforge.net (Dar�o Cutillas)
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
Option Base 1 'IMPORTANT!!!! TO CHECK LATER

'MSG Constants (for future multi-language support)
Private Const MSG_SAVE_FILEREADONLY = "This File is read-only. You must save to a different location."
Private Const MSG_SAVE_ERRORSAVING = "An error occurred when trying to save the file: "
Private Const MSG_SAVE_SUCCESS = "File saved succesfully!"
Private Const MSG_COMPILE_NOFENIXDIR = "Compiler directory has not been configured or does not exist"
Private Const MSG_COMPILE_NOTALREADYSAVED = "The file has not been saved yet. Save the file before compile"
Private Const MSG_RUN_DBCNOTFOUND = "DCB file not found. Compile first!"

'Public curPosition As Long
'Public prePosition As Long

Public codePosIndex As Long
Dim codePos() As Long


Dim selRange As CodeSenseCtl.IRange
Private nextTipText() As String
Public rangoActual As CodeSense.range
Public mustRefresh As Boolean
Dim showingList As CodeSenseCtl.ICodeList ' intellisense list
Dim showingToolTip As CodeSenseCtl.ICodeTip

Dim argID As String             ' Id of the parameter, the name
Dim codeListType As Boolean     ' Show codelist type: function, vars, process, ... list or param const list, global var const values, ...

Private WithEvents m_ContextMenu As cMenus
Attribute m_ContextMenu.VB_VarHelpID = -1

' struct to save and use the bookmark named list
Private Type bookmarkL
    line_number As Long
    name As String
End Type
Dim bookmarkList() As bookmarkL
Private curBookmark As Long
Private numLines As Long


Private m_FilePath As String
Private m_IsDirty As Boolean    'This should be never set directly (use the IsDirty property)
Private m_Title As String       'Basicly the caption of the form
Private m_addToProject As Boolean

Implements IFileForm

Private Sub cmbBookmarkList_Click()
    curBookmark = cmbBookmarkList.ListIndex
    cs.ExecuteCmd cmCmdGoToLine, bookmarkList(curBookmark + 2).line_number
    cs.HighlightedLine = bookmarkList(curBookmark + 2).line_number - 1
End Sub

Private Sub Cs_Change(ByVal Control As CodeSenseCtl.ICodeSense)
    On Error Resume Next

    ' this part controlls bookmark's state change
    testBookmarkListState (cs.LineCount - numLines)
    
    IsDirty = True
    '*********************************************************************************
    ' analyzes if we have to refresh the tree
    '*********************************************************************************
    If (Not rangoActual Is Nothing) And mustRefresh = False Then
    
        ' if we're in declaration zone
        If inDeclarationZone(rangoActual.StartLineNo) = True Then
        
            Dim token As CodeSenseCtl.cmTokenType
        
            ' We don't want to display a tip inside quoted or commented-out lines...
            token = Control.CurrentToken
        
            If ((cmTokenTypeText = token) Or (cmTokenTypeKeyword = token)) Then
                mustRefresh = True
            End If
            
        End If
    End If
    
    If cs.CurrentWordLength >= IS_Sensitive Then
        Cs_CodeList Control, showingList
        Control.ExecuteCmd cmCmdCodeList
'    Else
'        'If Not showingList Is Nothing Then
'        If Not showingToolTip Is Nothing Then
'        ' show the cmbParam
'        fillParamCmb
''            Cs_CodeList Control, showingList
''            Control.ExecuteCmd cmCmdCodeList
'        End If
    End If
    
End Sub

Private Function cs_Click(ByVal Control As CodeSenseCtl.ICodeSense) As Boolean
    RefreshStatusBar
End Function


Private Function Cs_CodeList(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList) As Boolean
    Dim i As Long
    Dim token As CodeSenseCtl.cmTokenType
    Dim curWord As String

    ' We don't want to display a tip inside quoted or commented-out lines...
    token = Control.CurrentToken

    If ((cmTokenTypeSingleLineComment = token) Or (cmTokenTypeMultiLineComment = token)) Then Exit Function
    
    
    If Not IS_Show Then
        Cs_CodeList = False
        Exit Function
    End If
    
    curWord = LCase(Control.CurrentWord)
    
    If IS_Sensitive > Len(curWord) Then
        ListCtrl.Destroy
        Cs_CodeList = False
        Exit Function
    End If
 
    ListCtrl.hImageList = frmProgramInspector.programImageList.hIml
    
    ' empty the list
    While ListCtrl.ItemCount > 0
        ListCtrl.DeleteItem (0)
    Wend
    

    
    If codeListType Then
         ' adds func/proc/vars/const... to the autocomplete list
         If Not inDeclarationZone(rangoActual.StartLineNo + 1) Then
         ' in declaration zone, functions, vars are forbidden
             If IS_LangDefFunc Then
                 For i = 1 To UBound(functionList)
                     If LCase(Left(functionList(i), Len(curWord))) = curWord Then
                         ListCtrl.AddItem functionList(i), 18 - 1, 1
                     End If
                 Next i
             End If
                 
             If IS_UserDefFunc Then
                 For i = 1 To UBound(userFunctionList)
                     If LCase(Left(userFunctionList(i), Len(curWord))) = curWord Then
                         ListCtrl.AddItem userFunctionList(i), 9 - 1, 1
                     End If
                 Next i
             End If
                 
             If IS_LangDefVar Then
                 For i = 1 To UBound(globalList)
                     If LCase(Left(globalList(i), Len(curWord))) = curWord Then
                         ListCtrl.AddItem globalList(i), 12 - 1, 2
                     End If
                 Next i
                 
                 For i = 1 To UBound(localList)
                     If LCase(Left(localList(i), Len(curWord))) = curWord Then
                         ListCtrl.AddItem localList(i), 13 - 1, 2
                     End If
                 Next i
                     
                 For i = 1 To UBound(globalStructList)
                     If LCase(Left(globalStructList(i), Len(curWord))) = curWord Then
                         ListCtrl.AddItem globalStructList(i), 16 - 1, 2
                     End If
                 Next i
                 
             '    For i = 1 To UBound(localStructList)
             '        ListCtrl.AddItem constList(i), 4, 2
             '    Next i
             End If
             
             If IS_UserDefVar Then
                 For i = 1 To UBound(varList)
                     If LCase(Left(varList(i), Len(curWord))) = curWord Then
                         ListCtrl.AddItem varList(i), 2 - 1, 2
                     End If
                 Next i
             End If
         End If
         
         If IS_LangDefConst Then
             For i = 1 To UBound(constList)
                 If LCase(Left(constList(i), Len(curWord))) = curWord Then
                     ListCtrl.AddItem constList(i), 11 - 1, 2
                 End If
             Next i
         End If
        
         If IS_UserDefConst Then
             For i = 1 To UBound(userConstList)
                 If LCase(Left(userConstList(i), Len(curWord))) = curWord Then
                     ListCtrl.AddItem userConstList(i), 1 - 1, 2
                 End If
             Next i
         End If
         
         If Not inConstDeclarationZone(rangoActual.StartLineNo + 1) And inDeclarationZone(rangoActual.StartLineNo + 1) Then
         ' In const declaration zone and out of declaration zone, types are forbidden
             For i = 1 To UBound(userTypeList)   ' user defined types,
                 If LCase(Left(userTypeList(i), Len(curWord))) = curWord Then
                     ListCtrl.AddItem userTypeList(i), 8 - 1, 2
                 End If
             Next i
             
             For i = 1 To UBound(typeList)   ' int, word, dword, byte, string, ...
                 If LCase(Left(typeList(i), Len(curWord))) = curWord Then
                     ListCtrl.AddItem typeList(i), 17 - 1, 2
                 End If
             Next i
         End If
    Else
        ' param const list
            
            Dim num As Integer
               
            With Ini
                .Path = App.Path & "/Help/constants.lan"
                .Section = argID
                
                .Key = "QOfConsts"
                .Default = "0"
                num = CInt(.Value)
                
                If num > 0 Then
                                      
                    i = 1
                    While i <= num
                        If i = 1 Then
                            .Key = "Const"
                        Else
                            .Key = "Const" & i
                        End If
                        'Debug.Print argID & "-" & i & " of " & num & ":" & .Value
                        ListCtrl.AddItem .Value
                        i = i + 1
                    Wend
                End If
            End With
    End If
    
    
    Dim wordIndex As Long

    wordIndex = ListCtrl.FindString(cs.CurrentWord)
    
    ' if doesn't match the word, destroyes it
    If wordIndex = -1 Then
        ListCtrl.Destroy
        Cs_CodeList = False
        Exit Function
    End If
        
   
    ' Just for kicks, we'll select the first item by default...
    ListCtrl.SelectedItem = wordIndex
    
    ' Enable mouse hot-tracking
    ListCtrl.EnableHotTracking
        
    ' defines the var that contains the list to be show
    Set showingList = ListCtrl
    
    ' Allow list view control to be displayed
    Cs_CodeList = True
    
    Control.ExecuteCmd cmCmdCodeList
    
End Function

Private Function Cs_CodeListCancel(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList) As Boolean
    Set showingList = Nothing
End Function

Private Function cs_CodeListChar(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList, ByVal wChar As Long, ByVal lKeyData As Long) As Boolean
    Cs_CodeList Control, ListCtrl
    RefreshStatusBar
End Function

Private Sub cs_CodeListHotTrack(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList, ByVal lItem As Long)
    RefreshStatusBar
End Sub

Private Function Cs_CodeListSelMade(ByVal Control As CodeSenseCtl.ICodeSense, ByVal ListCtrl As CodeSenseCtl.ICodeList) As Boolean
    Dim strItem As String
    Dim range As New CodeSenseCtl.range
    Dim isFunc As Boolean

    ' Determine which item was selected in the list
    strItem = ListCtrl.GetItemText(ListCtrl.SelectedItem)
    
    ' Get current selection
    Set range = cs.GetSel(True)
    
    ' if there's no more than one line selected
    If range.StartLineNo = range.EndLineNo And range.StartColNo = range.EndColNo Then
        ' calcule the long of the new sel
        range.EndColNo = range.StartColNo
        range.StartColNo = range.StartColNo - Len(cs.CurrentWord)
    End If
    
    If LCase(cs.CurrentWord) = "end" Then
        Exit Function
    End If
    
    cs.SetSel range, False
    
    isFunc = isReservedFunction(strItem)
        
    If isFunc Then ' Replace current selection
        cs.ReplaceSel (strItem & "(")
    ElseIf isUserDefinedFunction(strItem) Then
        cs.ReplaceSel (strItem & "(")
    Else
        cs.ReplaceSel (strItem)
    End If

    ' Get new selection
    Set range = cs.GetSel(True)

    ' Update range to end of newly inserted text
    range.StartColNo = range.StartColNo + Len(strItem) + IIf(isFunc Or isUserDefinedFunction(strItem), 1, 0)  ' if it's a function the add another position
        
    range.EndColNo = range.StartColNo
    range.EndLineNo = range.StartLineNo

    ' Move cursor
    cs.SetSel range, True
    
    Set showingList = Nothing
    
    '#TODO: only call this if the strItm was function, or a process
    showToolTip strItem
    
    ' Don't prevent list view control from being hidden
    Cs_CodeListSelMade = False
    
    codeListType = True
    
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
    
    tip.font.name = "Segoe UI"
    'tip.font.Size = 10
    
    
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
        Dim iArg As Integer
        Dim i As Integer
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
            tip.Destroy     ' Right-parentesis found
            Set showingToolTip = Nothing
        Else
            tip.Argument = iArg

            ' Set tooltip text to current function prototype
                        
            If UBound(nextTipText) > 1 Then
                tip.TipText = nextTipText(tip.CurrentFunction + 1)
            Else
                tip.TipText = nextTipText(1)
            End If
                
            argID = LCase(getParamName(iArg, tip.TipText))
            
            modMenuActions.mnuEditCodeCompletionHelp
            'fillParamCmb
        End If

    End If

End Sub

'Private Sub showConstCodeList(argID As String)
'    Dim num As Integer
'    Dim list As CodeSenseCtl.CodeList
'    Dim i
'
'    Set list = Nothing
'    Set list = New CodeSenseCtl.CodeList
'
'    With Ini
'        .Path = App.Path & "\Help\constants.lan"
'        .Section = argID
'
'        .Key = "QOfConsts"
'        .Default = "0"
'        num = CInt(.Value)
'
'        If num > 0 Then
'            'ReDim list(1 To num) As String
'
'            i = 1
'            While i <= num
'                If i = 1 Then
'
'                    .Key = "Const"
'                Else
'                    .Key = "Const" & i
'                End If
'                .Default = ""
'                'list.AddItem .Value
'                i = i + 1
'            Wend
'        End If
'    End With
'
'    Dim wordIndex As Long
'
'    wordIndex = list.FindString(cs.CurrentWord)
'    ' if doesn't match the word, destroyes it
'    If wordIndex = -1 Then
'        list.Destroy
'        Exit Sub
'    End If
'
'    ' Just for kicks, we'll select the first item by default...
'    list.SelectedItem = wordIndex
'
'    ' Enable mouse hot-tracking
'    list.EnableHotTracking
'
'    ' defines the var that contains the list to be show
'    Set showingList = list
'
'    cs.ExecuteCmd cmCmdCodeList
    
'End Sub

'*******************************************************************************
' Gets the parameter name to show it's constants list if it has
' *************************************************************

Private Function getParamName(ByVal numArg As Integer, func As String) As String
    Dim strName As String
    Dim Pos As Integer
    Dim i As Integer
    
    strName = Right(func, Len(func) - InStr(func, "("))

    If numArg > 0 Then
        For i = 1 To numArg
            strName = Right(strName, Len(strName) - InStr(2, strName, ","))
        Next i
    End If
    
    strName = Right(strName, Len(strName) - InStr(2, strName, " "))
    
    If InStr(strName, ",") > 0 Then
        strName = Left(strName, InStr(strName, " ") - 2)
    Else
        strName = Left(strName, InStr(strName, " ") - 1)
    End If
    
        Debug.Print " -- " & numArg & " -- " & strName
    
    getParamName = strName
    
End Function


'************************************************************************
    ' Returns the index of the string in the user declared variables list
    ' -1 if it hasn't found
    ' if it returns 0, something  has gone wrong
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
    ' Returns the index of the string in the function list
    ' -1 if it hasn't found
    ' if it returns 0, something  has gone wrong
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

' Returns the function name of the current line
Private Function getCurrentFunction() As String
    
    ' get the line, current line
    Dim linea As String
    
    linea = cs.getLine(rangoActual.EndLineNo)
    
    ' cut the line on the cursor por
    linea = Mid(linea, 1, rangoActual.EndColNo)
    
    '*******************************************
    '*********** Line clearing *****************
    '*******************************************
        
    ' replaces the not visible chars with spaces
    linea = replace(linea, Chr(9), " ")
    linea = replace(linea, vbNewLine, " ")
    linea = replace(linea, vbCrLf, " ")
    linea = replace(linea, vbCr, " ")
    linea = replace(linea, vbLf, " ")
    linea = replace(linea, vbNullChar, " ")
    linea = replace(linea, vbBack, " ")
    linea = replace(linea, vbFormFeed, " ")
    linea = replace(linea, vbVerticalTab, " ")
                
    ' look if there are closed function before
    While InStrRev(linea, ")")
        linea = Left(linea, InStrRev(linea, ")") - 1)
        If InStr(linea, "(") Then
            linea = Mid(linea, 1, InStrRev(linea, "(") - 1)
        End If
    Wend
            
    'now, cut the line on the last "("
    If InStr(linea, "(") Then
        linea = Mid(linea, 1, InStrRev(linea, "(") - 1)
    Else
        ' if there isn't, there's no function
        linea = ""
        Exit Function
    End If
    
    ' get the last word from the line
    linea = getWordRev(linea)
    
    getCurrentFunction = linea
    
End Function

Private Sub cs_DeleteLine(ByVal Control As CodeSenseCtl.ICodeSense, ByVal lLine As Long, ByVal lItemData As Long)
    RefreshStatusBar
    
    If existsBookmark(lLine) <> -1 Then
        delBookmark (existsBookmark(lLine))
        refreshBookmarkList
    End If
    updateBookmarks lLine, -1
End Sub

Private Function cs_KeyDown(ByVal Control As CodeSenseCtl.ICodeSense, ByVal KeyCode As Long, ByVal Shift As Long) As Boolean
    RefreshStatusBar
End Function

Private Function Cs_KeyPress(ByVal Control As CodeSenseCtl.ICodeSense, ByVal KeyAscii As Long, ByVal Shift As Long) As Boolean
Dim linea As String

On Error Resume Next
    
    Dim i As Integer
    
    RefreshStatusBar
    Dim token As CodeSenseCtl.cmTokenType
 
    ' code to show the intellisense list
    ' We don't want to display a tip inside quoted or commented-out lines...
    token = Control.CurrentToken
    If ((cmTokenTypeText = token) Or (cmTokenTypeKeyword = token)) Then
    
        ' show IntelliSense
        If cs.CurrentWordLength >= IS_Sensitive - 1 Then
        'If cs.CurrentWordLength > 0 Then
            If showingList Is Nothing And KeyAscii >= 65 And KeyAscii <= 122 And (indexOnFunctionList(cs.CurrentWord & Chr(KeyAscii)) > 0 Or indexOnVarList(cs.CurrentWord & Chr(KeyAscii)) > 0) Then
                cs.ExecuteCmd cmCmdCodeList
            End If
        End If
        ' in the case of it's showing the list
        If Not showingList Is Nothing Then
            ' destroy in the next cases

            If KeyAscii < 65 Or KeyAscii > 122 Then
                showingList.Destroy
                Set showingList = Nothing
            End If
            
            If Not showingList Is Nothing Then
                ' if we don't found the word, cancel
                If indexOnFunctionList(cs.CurrentWord & Chr(KeyAscii)) = -1 And indexOnVarList(cs.CurrentWord & Chr(KeyAscii)) = -1 Then
                    showingList.Destroy
                    Set showingList = Nothing
                End If
            End If
            
        End If
        
        Dim func As String

        ' code to show the list
        If (Asc("(") = KeyAscii) Then
            'Dim linea As String
            ' get the line of the cursor
            linea = Trim(cs.getLine(rangoActual.EndLineNo))
            ' cut where is the cursor
            linea = Mid(linea, 1, rangoActual.StartColNo)
            
            '*******************************************
            '*********** Line clearing *****************
            '*******************************************
                
            ' replaces the not visible chars with spaces
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
        
        If (Asc("=") = KeyAscii) Then
'             'Dim linea As String
'
'            ' get the line of the cursor
'            linea = Trim(cs.getLine(rangoActual.EndLineNo))
'            ' cut where is the cursor
'            linea = Mid(linea, 1, rangoActual.StartColNo)
'
'            '*******************************************
'            '*********** Line clearing *****************
'            '*******************************************
'
'            ' replaces the not visible chars with spaces
'            linea = replace(linea, Chr(9), " ")
'            linea = replace(linea, vbNewLine, " ")
'            linea = replace(linea, vbCrLf, " ")
'            linea = replace(linea, vbCr, " ")
'            linea = replace(linea, vbLf, " ")
'            linea = replace(linea, vbNullChar, " ")
'            linea = replace(linea, vbBack, " ")
'            linea = replace(linea, vbFormFeed, " ")
'            linea = replace(linea, vbVerticalTab, " ")
            
            argID = getLine
            
            codeListType = False
            
            cs.ExecuteCmd cmCmdCodeList
            
        End If
        
        If Not (cmTokenTypeSingleLineComment = token) And Not (cmTokenTypeMultiLineComment = token) Then
            ' shows the tooltip when it's pressed ","
            If (Asc(",") = KeyAscii) Then
                func = getCurrentFunction
                If func <> "" Then
                    showToolTip func
                End If
            End If
        End If
            
    Else    ' It's necessary of both cases
        If Not (cmTokenTypeSingleLineComment = token) And Not (cmTokenTypeMultiLineComment = token) Then
            
            ' shows the tooltip when it's pressed "," and it was no tooltip
            If (Asc(",") = KeyAscii) Then
                func = getCurrentFunction
                If func <> "" Then
                    showToolTip func
                End If
            End If
            
            '#TODO:
            ' show global variable's const list
            If (Asc("=") = KeyAscii) Then
'                 'Dim linea As String
'                ' get the line of the cursor
'                linea = Trim(cs.getLine(rangoActual.EndLineNo))
'                ' cut where is the cursor
'                linea = Mid(linea, 1, rangoActual.StartColNo)
'
'                '*******************************************
'                '*********** Line clearing *****************
'                '*******************************************
'
'                ' replaces the not visible chars with spaces
'                linea = replace(linea, Chr(9), " ")
'                linea = replace(linea, vbNewLine, " ")
'                linea = replace(linea, vbCrLf, " ")
'                linea = replace(linea, vbCr, " ")
'                linea = replace(linea, vbLf, " ")
'                linea = replace(linea, vbNullChar, " ")
'                linea = replace(linea, vbBack, " ")
'                linea = replace(linea, vbFormFeed, " ")
'                linea = replace(linea, vbVerticalTab, " ")
                
                argID = getLine
                
                codeListType = False
                
                cs.ExecuteCmd cmCmdCodeList

            End If
            
            '#TODO:
            ' maybe a little control if the sentence is well writen?
            If (Asc(";") = KeyAscii) Then
            
            End If
        End If
    End If
End Function
' returns the current line
Private Function getLine() As String
Dim linea As String
            
            'pos.StartLineNo
            linea = cs.getLine(rangoActual.StartLineNo)
            
            linea = clearLine(linea)
            
            Debug.Print "getLine" & linea
            
            linea = Trim(linea)
            
            ' if it's a composited sentence
            If InStrRev(linea, ";") <> 0 Then
                linea = Right$(linea, Len(linea) - InStr(linea, ";"))
            End If
            
            linea = Trim(linea)
            
            Debug.Print "getLine(;)" & linea
            
            getLine = linea
End Function



' shows the parameter tooltip
Private Sub showToolTip(Word As String)
    Dim prototipo() As String
    Dim howMany As Long
    Dim i As Integer
    
    If indexOnFunctionList(Word) > 0 And indexOnFunctionList(Word) <= iFunctionCount Then
        With Ini
            .Path = App.Path & "/Help/functions.lan"
            .Section = Word
            
            .Key = "QOfPrototipes"
            .Default = "0"
            howMany = CInt(.Value)
            
            If howMany > 0 Then
                ReDim prototipo(1 To howMany) As String
               
                i = 1
                While i <= howMany
                    If i = 1 Then
                        .Key = "Prototipe"
                    Else
                        .Key = "Prototipe" & i
                    End If
                    .Default = ""
                    prototipo(i) = .Value
                    i = i + 1
                Wend
            End If
            
        End With
               
        If howMany > 0 Then
            ReDim nextTipText(1 To howMany) As String
            i = 1
            While i <= howMany
                nextTipText(i) = prototipo(i)
                i = i + 1
            Wend
            Me.cs.ExecuteCmd cmCmdCodeTip
        End If
    End If
    
    ' is inside of the list but declared by the user
    If indexOnFunctionList(Word) > iFunctionCount Then
        ReDim nextTipText(1) As String
        nextTipText(1) = Word & " (" & parameters(indexOnFunctionList(Word) - iFunctionCount) & ")"
        Me.cs.ExecuteCmd cmCmdCodeTip
    End If
End Sub

Private Function Cs_KeyUp(ByVal Control As CodeSenseCtl.ICodeSense, ByVal KeyCode As Long, ByVal Shift As Long) As Boolean

    ' check-ins for the list
    If Not showingList Is Nothing Then
        If cs.CurrentWordLength < 0 Then
            showingList.Destroy
            Set showingList = Nothing
        End If
    End If
    
End Function

Private Function Cs_MouseDown(ByVal Control As CodeSenseCtl.ICodeSense, ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long) As Boolean
    Dim lParentIndex, iP2 As Long
    Dim s, sl, sw, n, c As Boolean
    ' s for selected
    '   sl for single line selected
    '   sw for single word selected
    ' n for nothing selected
    ' c for converteable selections

On Error Resume Next

    s = False
    n = False
    sl = False
    sw = False
    c = False
    
    If rangoActual.IsEmpty Then
        n = True
    Else
        s = True
        If rangoActual.StartLineNo = rangoActual.EndLineNo Then
            If cs.SelText = cs.CurrentWord Then
                Debug.Print cs.SelText & "..." & cs.CurrentWord
                If isBin(cs.SelText) Or isHex(cs.SelText) Or IsNumeric(cs.SelText) Then
                    c = True
                End If
                sw = True
            Else
                sl = True
            End If
        End If
    End If
    
    If (Button = 2) Then
        
        Set m_ContextMenu = Nothing
        Set m_ContextMenu = New cMenus
        m_ContextMenu.DrawStyle = M_Style
        Set m_ContextMenu.ImageList = frmMain.ImgList1.Object
        m_ContextMenu.CreateFromNothing Me.Hwnd
        
        lParentIndex = m_ContextMenu.AddItem(0, Key:="ContextMenu")
        With m_ContextMenu
            .AddItem lParentIndex, "Go to definition", , , "mnuNavigationGoToDefinition"
            .AddItem lParentIndex, "-"
            If s Then
                .AddItem lParentIndex, "Cut", "Ctrl+X", , "mnuEditCut", , , , 5
                .AddItem lParentIndex, "Copy", "Ctrl+C", , "mnuEditCopy", , , , 4
            End If
            If cs.CanPaste Then
                .AddItem lParentIndex, "Paste", "Ctrl+V", , "mnuEditPaste", , , , 6
            End If
            If s Or cs.CanPaste Then
                .AddItem lParentIndex, "-"
            End If
            If n Then
                .AddItem lParentIndex, "Select all", "Ctrl+A", , "mnuEditSelectAll", , , , 75
                .AddItem lParentIndex, "Select line", "Ctrl+Shift+L", , "mnuEditSelectLine", , , , 76
                .AddItem lParentIndex, "Select word", "Ctrl+Shift+W", , "mnuEditSelectWord", , , , 86
            Else
                .AddItem lParentIndex, "Deselect", , , "mnuEditDeselect"
            End If
            If n Then
                .AddItem lParentIndex, "-"
                .AddItem lParentIndex, "Duplicate line", "Ctrl+D", , "mnuEditDuplicateLine", , , , 83
                .AddItem lParentIndex, "Delete line", "Ctrl+R", , "mnuEditDeleteLine", , , , 84
                .AddItem lParentIndex, "Clear line", , , "mnuEditClearLine"
                .AddItem lParentIndex, "Up line      ^", "Ctrl+Shift+Up", , "mnuEditUpLine", , , , 87
                .AddItem lParentIndex, "Down line  v", "Ctrl+Shift+Down", , "mnuEditDownLine", , , , 88
            End If
            .AddItem lParentIndex, "-"
            .AddItem lParentIndex, "Shift line &left", "Tab", , "mnuEditTab", Image:=40
            .AddItem lParentIndex, "Shift line &right", "Shift+Tab", , "mnuEditUnTab", Image:=41
            .AddItem lParentIndex, "-"
            .AddItem lParentIndex, "Comment", "Ctrl+J", , "mnuEditComment", Image:=42
            .AddItem lParentIndex, "UnComment", "Ctrl+Shift+J", , "mnuEditUnComment", Image:=43
            .AddItem lParentIndex, "-"
            If s Then
                .AddItem lParentIndex, "UPPER CASE", "Ctrl+U", , "mnuEditUpperCase", Image:=60
                .AddItem lParentIndex, "lower case", "Ctrl+L", , "mnuEditLowerCase", Image:=61
                .AddItem lParentIndex, "Proper Case", , , "mnuEditFirstCase", Image:=94
                .AddItem lParentIndex, "Sentence case.", , , "mnuEditSentenceCase", Image:=93
                .AddItem lParentIndex, "iNVERSE cASE", , , "mnuEditChangeCase", Image:=92
                .AddItem lParentIndex, "-"
            End If
            If sw Then
                iP2 = .AddItem(lParentIndex, "Convert") 'Conversions
                    .AddItem iP2, "Bin -> Hex", , , "mnuConvertBinHex"
                    .AddItem iP2, "Bin -> Dec", , , "mnuConvertBinDec"
                    .AddItem iP2, "-"
                    .AddItem iP2, "Hex -> Bin", , , "mnuConvertHexBin"
                    .AddItem iP2, "Hex -> Dec", , , "mnuConvertHexDec"
                    .AddItem iP2, "-"
                    .AddItem iP2, "Dec -> Bin", , , "mnuConvertDecBin"
                    .AddItem iP2, "Dec -> Hex", , , "mnuConvertDecHex"
            End If
            If n Then
                .AddItem lParentIndex, "Code completion help", "Ctrl+Space", , "mnuEditCodeCompletionHelp"
            End If
            If n Or sw Then
                .AddItem lParentIndex, "-"
            End If
            .AddItem lParentIndex, "Search...", "Ctrl+F", , "mnuNavigationSearch", , , , 13
            If sw Or sl Then
                .AddItem lParentIndex, "Search next selected", "Ctrl+F3", , "mnuNavigationSearchNextWord", , , , 89
                .AddItem lParentIndex, "Search prev selected", "Ctrl+Shift+F3", , "mnuNavigationSearchPrevWord", , , , 90
            End If
            .AddItem lParentIndex, "-"
            .AddItem lParentIndex, "Replace...", "Ctrl+H", , "mnuNavigationReplace", Image:=62
            If sw And (cs.SelText = "[" Or cs.SelText = "]") Then
                .AddItem lParentIndex, "-"
                .AddItem lParentIndex, "Go to matching brace", "Ctrl+Shift+B", , "mnuNavigationGotoMatchBrace", , , , 85
            End If
        
            .PopupMenu "ContextMenu"
        End With
        
        'frmMain.cMenu.PopupMenu "mnuEdit"
        'frmMain.cMenu.PopupMenu "mnuNavigation"
        'contextMenu.ShowPopupMenu X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY
    End If
End Function

Private Function cs_MouseMove(ByVal Control As CodeSenseCtl.ICodeSense, ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long) As Boolean
    RefreshStatusBar
End Function

Private Function cs_MouseUp(ByVal Control As CodeSenseCtl.ICodeSense, ByVal Button As Long, ByVal Shift As Long, ByVal X As Long, ByVal Y As Long) As Boolean
    If rangoActual Is Nothing Then
        Exit Function
    End If
'    prePosition = curPosition
'    curPosition = rangoActual.StartLineNo + 1
'    codePosIndex = codePosIndex + 1
'    ReDim Preserve codePos(codePosIndex)
'    codePos(codePosIndex) = rangoActual.StartColNo
    setNewPos (rangoActual.StartLineNo)
End Function

Private Function Cs_RClick(ByVal Control As CodeSenseCtl.ICodeSense) As Boolean
    Cs_RClick = True
End Function

Private Sub Cs_RegisteredCmd(ByVal Control As CodeSenseCtl.ICodeSense, ByVal lCmd As CodeSenseCtl.cmCommand)
    ' register command to show the help
    Dim sword As String
    If lCmd = 1000 Then
        sword = cs.CurrentWord
        If sword <> "" Then
            ' local help
            NewWindowWeb App.Path & "/help/fenix/func.php-func=" & UCase(sword) & ".htm", "HELP: " & UCase(sword), App.Path & "/help/fenix/func.php-frame=top.htm"
            
            ' help on inet
            'NewWindowWeb "http://fenix.jlceb.com/func.php?func=" & UCase(sword)
        End If
    End If
End Sub

Private Function cs_Return(ByVal Control As CodeSenseCtl.ICodeSense) As Boolean
    updateBookmarks 1, 1
End Function

Private Sub Cs_SelChange(ByVal Control As CodeSenseCtl.ICodeSense)
    On Error Resume Next
    
    cs.HighlightedLine = -1
    
    Dim rangoTemp As CodeSense.range
    Set rangoTemp = cs.GetSel(True)
    
    If Not rangoActual Is Nothing Then
        ' line changing
        If rangoTemp.StartLineNo <> rangoActual.StartLineNo Then
            ' we have to refresh when changing the line
            If mustRefresh = True Then
                MakeProgramIndex IFileForm_FilePath
                mustRefresh = False
            End If
            
            ' kill the tooltip
            If Not showingToolTip Is Nothing Then
                showingToolTip.Destroy
                Set showingToolTip = Nothing
            End If
        End If
    End If
    
    Set rangoActual = cs.GetSel(True)
End Sub

Private Sub Form_Activate()
    ' testing
'    If existTreeForFile(IFileForm_FilePath) = False Then
'        mustRefresh = True
'    End If

    ' If this is very slow, try copy/pasting this part in IFileForm.Load
    mustRefresh = True
    MakeProgramIndex IFileForm_FilePath
    mustRefresh = False
    
    cs.EnableColumnSel = False
    'ReDim Preserve bookmarkList(1)
    enableDisableBookmarks
    
    codeListType = True
   
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler:
    
'    prePosition = 1
'    curPosition = 1
    
    codePosIndex = 1
    'ReDim Preserve codePos(2) As CodeSenseCtl.position
    'codePos(1).LineNo = 1
    MakeProgramIndex IFileForm_FilePath
    mustRefresh = False
    
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
        .AddButton eButtonStyle:=CTBSeparator
        .AddControl cmbBookmarkList.Hwnd, , "cmbBookmarkList"
        .AddButton "Edit Bookmarks", 8, , , , CTBAutoSize, "EditBookmarks"
        .AddButton eButtonStyle:=CTBSeparator
        .AddButton "Previous Position", 9, , , , CTBAutoSize, "PrevPos"
        .AddButton "Next Position", 10, , , , CTBAutoSize, "NextPos"
    End With
    
    'Create the rebar
    With ReBar
        If A_Bitmaps Then
            .BackgroundBitmap = App.Path & "\resources\backrebar" & A_Color & ".bmp"
        End If
        .CreateRebar Me.Hwnd
        .AddBandByHwnd tbrSource.Hwnd, , True, False
    End With
    ReBar.RebarSize
    
    ' configure the edition control
    cs.LineNumbering = True
    cs.LineNumberStart = 1
    cs.LineNumberStyle = cmDecimal
    cs.LineToolTips = True
    cs.BorderStyle = cmBorderStatic
    cs.EnableDragDrop = True
    cs.EnableCRLF = True
    cs.TabSize = 2
    cs.ColorSyntax = True
    cs.Language = "Bennu"
    cs.DisplayLeftMargin = True
    cs.AutoIndentMode = cmIndentPrevLine
    LoadCSConf cs
    
    ' register the command to show help
    ' execute the event Cs_RegisteredCmd
    Dim g As New CodeSenseCtl.Globals
    Dim h As New CodeSenseCtl.HotKey    'registers hotkey
    ' F1
    Call g.RegisterCommand(1000, "ShowHelp", "Shows the help about the current word in the control.")
    h.VirtKey1 = Chr(vbKeyF1)
    Call g.RegisterHotKey(h, 1000)
    ' Shift + F1
'    Call g.RegisterCommand(1001, "ShowWiki", "Shows the wiki-help about the current word in the control.")
'    h.VirtKey1 = Chr(vbKeyShift)
'    h.VirtKey2 = Chr(vbKeyF1)
'    Call g.RegisterHotKey(h, 1001)

    Exit Sub
ErrHandler:
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
        ReBar.RebarSize
        cs.Move 0, ScaleY(ReBar.RebarHeight, vbPixels, vbTwips)
        cs.Width = Me.ScaleWidth
        cs.Height = Me.ScaleHeight - ScaleY(ReBar.RebarHeight, vbPixels, vbTwips)
    End If
End Sub
    
Private Sub Form_Terminate()
    ' in the case that there's no opened prg, clear the declaration tree
    If DocExist() = False Then
        frmProgramInspector.tv_program.Nodes.Clear
    End If
End Sub

'�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-'
'STARTS INTERFACE IFILEFORM
'�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-'
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
    Dim sFileBMK As String
    'Dim fs As FileSystemObject
    Dim A As textStream
    Dim i As Long
    Dim str As String
    
    cs.OpenFile (sFile)
    'There is no way to determine if the cs.openfile fails so assume everything goes well
    'since we check for the existence of the file in the NewFileForm function this should work
    'well (any file is supposed to be able to be opened in text format)
    lResult = -1
    m_FilePath = sFile
    IsDirty = False
       
    ' prepare the bookmarks from file
    sFileBMK = Left(sFile, Len(sFile) - 3) & "bmk"
    ReDim bookmarkList(1)
    ReDim codePos(1)
    If FSO.FileExists(sFileBMK) Then
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Set A = FSO.OpenTextFile(sFileBMK, ForReading)
        i = 1
        While Not A.AtEndOfStream
            str = A.ReadLine
            'addBookmark (CLng(str))
            ReDim Preserve bookmarkList(UBound(bookmarkList) + 1)
            bookmarkList(i + 1).line_number = CLng(str)
            str = A.ReadLine
            bookmarkList(i + 1).name = str
            cs.SetBookmark bookmarkList(i + 1).line_number - 1, True
            i = i + 1
            cmbBookmarkList.Enabled = True
        Wend
        A.Close
        
        refreshBookmarkList
    Else
        cmbBookmarkList.Enabled = False
    End If
    
    
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
    ReDim bookmarkList(1)
    ReDim codePos(1)
End Function

Private Function IFileForm_Save(ByVal sFile As String) As Long
    Dim lResult As Long
    Dim sFileBMK As String
    'Dim fs As FileSystemObject
    Dim A As textStream
    Dim i As Long
    
    
    If FSO.FileExists(sFile) Then Kill sFile 'Delete the file if exists
    
    On Error GoTo ErrHandler
    cs.SaveFile sFile, False 'Save the file
    

    sFileBMK = Left(sFile, Len(sFile) - 3) & "bmk"
    If FSO.FileExists(sFileBMK) Then Kill sFileBMK
    
    If getLastBookmarkIndex > 1 Then
        Set FSO = CreateObject("Scripting.FileSystemObject")
        Set A = FSO.CreateTextFile(sFileBMK, True, False)
        For i = 2 To UBound(bookmarkList)
            A.WriteLine (bookmarkList(i).line_number)
            A.WriteLine (bookmarkList(i).name)
        Next i
        
        A.Close
    End If

    ' HERE THERE SHOULD BE SOME KIND OF COMPROBATION FOR ERRORS AFTER SAVEFILE
    lResult = -1
    If (lResult) Then 'Save succesful
        ' Add to project if necessary
        If IFileForm_AlreadySaved = False And m_addToProject = True Then addFileToProject sFile
    
        If m_FilePath <> sFile Then 'Show a success message only if the name changed
            MsgBox MSG_SAVE_SUCCESS, vbInformation
        End If
        m_FilePath = sFile
        IsDirty = False
    Else
        MsgBox MSG_SAVE_ERRORSAVING, vbCritical
    End If

ErrHandler:
    If Err.Number > 0 Then lResult = -1: Resume Next
    
End Function
'�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-'
'END INTERFACE IFILEFORM
'�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-�-'
'This property is not part of the interface, just helps to reduce code
'by setting the caption of the form properly
Private Property Let IsDirty(ByVal newVal As Boolean)
    m_IsDirty = newVal
    'Put an * to the caption if dirty
    Caption = IFileForm_Title & IIf(newVal, " *", "")
    
    frmMain.RefreshTabs
End Property

Private Sub m_ContextMenu_Click(ByVal Index As Long)
    Select Case m_ContextMenu.ItemKey(Index)
        Case "mnuNavigationGoToDefinition":     Call mnuNavigationGoToDefinition
        Case "mnuEditCut":                      Call mnuEditCut
        Case "mnuEditCopy":                     Call mnuEditCopy
        Case "mnuEditPaste":                    Call mnuEditPaste
        Case "mnuEditSelectAll":                Call mnuEditSelectAll
        Case "mnuEditSelectWord":               Call mnuEditSelectWord
        Case "mnuEditSelectLine":               Call mnuEditSelectLine
        Case "mnuEditDeselect":                 Call mnuEditDeselect
        Case "mnuEditClearLine":                Call mnuEditClearLine
        Case "mnuEditDuplicateLine":            Call mnuEditDuplicateLine
        Case "mnuEditDeleteLine":               Call mnuEditDeleteLine
        Case "mnuEditUpLine":                   Call mnuEditUpLine
        Case "mnuEditDownLine":                 Call mnuEditDownLine
        Case "mnuEditTab":                      Call mnuEditTab
        Case "mnuEditUnTab":                    Call mnuEditUnTab
        Case "mnuEditComment":                  Call mnuEditComment
        Case "mnuEditUnComment":                Call mnuEditUnComment
        Case "mnuEditUpperCase":                Call mnuEditUpperCase
        Case "mnuEditLowerCase":                Call mnuEditLowerCase
        Case "mnuEditChangeCase":               Call mnuEditChangeCase
        Case "mnuEditFirstCase":                Call mnuEditFirstCase
        Case "mnuEditSentenceCase":             Call mnuEditSentenceCase
        Case "mnuEditCodeCompletionHelp":       Call mnuEditCodeCompletionHelp
        
        Case "mnuConvertBinHex":                Call mnuConvertBinHex
        Case "mnuConvertBinDec":                Call mnuConvertBinDec
        Case "mnuConvertHexBin":                Call mnuConvertHexBin
        Case "mnuConvertHexDec":                Call mnuConvertHexDec
        Case "mnuConvertDecBin":                Call mnuConvertDecBin
        Case "mnuConvertDecHex":                Call mnuConvertDecHex
            
        Case "mnuNavigationSearch":             Call mnuNavigationSearch
        Case "mnuNavigationSearchNext":         Call mnuNavigationSearchNext
        Case "mnuNavigationSearchPrev":         Call mnuNavigationSearchPrev
        Case "mnuNavigationSearchNextWord":     Call mnuNavigationSearchNextWord
        Case "mnuNavigationSearchPrevWord":     Call mnuNavigationSearchPrevWord
        Case "mnuNavigationReplace":            Call mnuNavigationReplace
        Case "mnuNavigationGotoMatchBrace":     Call mnuNavigationGotoMatchBrace
    End Select
    
End Sub

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
        modMenuActions.mnuEditTab
    Case "ShiftLeft"
        modMenuActions.mnuEditUnTab
    Case "Comment"
        modMenuActions.mnuEditComment
    Case "Uncomment"
        modMenuActions.mnuEditUnComment
    Case "EditBookmarks"
        modMenuActions.mnuBookmarkEdit
    Case "PrevPos"
        modMenuActions.mnuNavigationPrevPosition
    Case "NextPos"
        modMenuActions.mnuNavigationNextPosition
    End Select
        
End Sub

Public Property Get bookmarkList_name(Index As Long) As String
    bookmarkList_name = bookmarkList(Index).name
End Property

Public Property Get bookmarkList_line_Number(Index As Long) As Long
    bookmarkList_line_Number = bookmarkList(Index).line_number
End Property

Public Property Let bookmarkList_name(Index As Long, name As String)
    bookmarkList(Index).name = name
End Property

Public Property Let bookmarkList_line_Number(Index As Long, line_number As Long)
    bookmarkList(Index).line_number = line_number
End Property

Public Function existsBookmark(ln As Long) As Long
    Dim i As Long
    For i = 1 To UBound(bookmarkList)
        If bookmarkList(i).line_number = ln Then
            existsBookmark = i
            Exit Function
        End If
    Next i
    existsBookmark = -1
End Function

Public Sub addBookmark(ln As Long)
    Dim lines As Long
    Dim i As Long

    Dim lastBookmark As Long
    lastBookmark = UBound(bookmarkList)
    ReDim Preserve bookmarkList(lastBookmark + 1) As bookmarkL
    bookmarkList(lastBookmark + 1).line_number = ln
    bookmarkList(lastBookmark + 1).name = "Bookmark" & ln
    
    enableDisableBookmarks
End Sub

Public Sub delBookmark(Index As Long)
    Dim lastBookmark As Long
    Dim i As Long
    lastBookmark = UBound(bookmarkList)
    ' refresh all the list
    If Index <> lastBookmark Then
        For i = Index To lastBookmark - 1
            bookmarkList(i).name = bookmarkList(i + 1).name
            bookmarkList(i).line_number = bookmarkList(i + 1).line_number
        Next i
    End If
    ReDim Preserve bookmarkList(lastBookmark - 1) As bookmarkL
    
    enableDisableBookmarks
        
End Sub

Public Sub refreshBookmarkList()
    Dim i As Long
    cmbBookmarkList.Clear
    For i = 2 To UBound(bookmarkList)
        cmbBookmarkList.AddItem bookmarkList(i).name
    Next i
    cmbBookmarkList.Refresh
End Sub

Public Sub delAllBookmark()
    ReDim bookmarkList(1)
    enableDisableBookmarks
End Sub

Public Function getLastBookmarkIndex() As Long
    On Error GoTo Err
    getLastBookmarkIndex = UBound(bookmarkList)
    Exit Function
Err:
    getLastBookmarkIndex = 1
    Exit Function
End Function

Public Sub updateBookmarks(l As Long, n As Long)
    Dim lastBookmark As Long
    Dim i As Long
    lastBookmark = UBound(bookmarkList)
    ' updates line number after change in the code

    For i = 2 To lastBookmark
        If bookmarkList(i).line_number > l Then
            bookmarkList(i).line_number = bookmarkList(i).line_number + n
        End If
    Next i
End Sub

Public Sub testBookmarkListState(diference As Long)
    Dim lastBookmark As Long
    Dim i As Long
    Dim delList() As Boolean           ' lisr of bookmark that must be deleted
    
    lastBookmark = UBound(bookmarkList)
    
    ReDim delList(2 + lastBookmark)
    For i = 2 To UBound(delList)    ' we init the list
        delList(i) = False
    Next i
    
    For i = 2 To lastBookmark
        If Not cs.GetBookmark(bookmarkList(i).line_number - 1) Then
            If Not cs.GetBookmark((bookmarkList(i).line_number - 1) + diference) Then
                'delete this bookmark
                delList(i) = True
            Else
                'update bookmark line number
                'updateBookmarks bookmarkList(i).line_number, diference
                bookmarkList_line_Number(i) = bookmarkList(i).line_number + diference
            End If
        End If
    Next i
    
    For i = 2 To UBound(delList)    ' we delete the elements marked
        If delList(i) Then
            delBookmark (i)
        End If
    Next i
    
    refreshBookmarkList
End Sub

Public Sub enableDisableBookmarks()
    If getLastBookmarkIndex > 1 Then    ' enable
        cmbBookmarkList.Enabled = True
        tbrSource.ButtonEnabled("EditBookmarks") = True
    Else                                ' disable
        cmbBookmarkList.Enabled = False
        tbrSource.ButtonEnabled("EditBookmarks") = False
    End If
End Sub


Public Sub RefreshStatusBar()
    If rangoActual Is Nothing Then
        Exit Sub
    End If
    numLines = cs.LineCount
    If rangoActual.StartLineNo = rangoActual.EndLineNo Then
        frmMain.StatusBar.PanelText("MAIN") = "Line: " & rangoActual.StartLineNo + 1 _
            & " of " & cs.LineCount & Chr(vbKeyTab) & "Sel: None"
    Else
        frmMain.StatusBar.PanelText("MAIN") = "Line: " & rangoActual.StartLineNo + 1 _
            & " of " & cs.LineCount & Chr(vbKeyTab) & "Sel: " _
            & rangoActual.StartLineNo + 1 & " to " & rangoActual.EndLineNo + 1 _
            & "   Len: " & cs.SelLengthLogical
    End If
    RefreshProcPos
End Sub

Public Sub RefreshProcPos()
    Dim Pos As Long
    If rangoActual Is Nothing Then
        Exit Sub
    End If
    If rangoActual.StartLineNo = rangoActual.EndLineNo Then
        Pos = rangoActual.StartLineNo
        frmProgramInspector.findCurProc (Pos + 1)
    Else
        ' don't highlight the current proc/func
        Pos = -1
    End If
End Sub


Public Sub setNewPos(line As Long)
    codePosIndex = codePosIndex + 1
    ReDim Preserve codePos(codePosIndex) As Long
    codePos(codePosIndex) = line
    refreshPosList
End Sub

Public Function getPos(Index As Long) As Long ' currently returns a line number
    getPos = codePos(codePosIndex)
End Function

Public Function uPos()
    uPos = UBound(codePos)
End Function

Public Sub refreshPosList()
On Error GoTo ErrHandler
    Dim i As Long
    For i = LBound(codePos) To UBound(codePos)
        If i = codePosIndex Then
            Debug.Print ">> " & codePos(i)
        Else
            Debug.Print "   " & codePos(i)
        End If
    Next i
ErrHandler:
    Exit Sub
End Sub
