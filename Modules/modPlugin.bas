Attribute VB_Name = "modPlugin"
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
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public Sub RegisterPlugins()
    Dim fileString As String
    Dim result As Variant
    fileString = Dir(App.Path & "\Plugins\")
    fileString = LCase(fileString)
    Do Until fileString = ""
        If Right(fileString, 4) = ".dll" Then
            
            SetSplashMessage "Registering Plug-in" & fileString
            DoEvents
            
            result = Register(App.Path & "\Plugins\" & fileString)
            
            Select Case result
                Case 1: MsgBox "File Could Not Be Loaded Into Memory Space"
                Case 2: MsgBox "Not A Valid ActiveX Component"
                Case 3: MsgBox "ActiveX Component Registration Failed"
                'Case 4: SetSplashMessage "stActiveXComponentRegistrationSuccessful"
                Case 5: MsgBox "ActiveX Component UnRegister Successful"
                Case 6: MsgBox "ActiveX Component UnRegistration Failed"
                Case 7: MsgBox "No File Provided"
            End Select
        End If
        fileString = Dir
        fileString = LCase(fileString)
        DoEvents
    Loop
End Sub

Public Sub LoadPlugins()

' This generic function will look for all plugins in a spesified directory.
' It will then query the plugin for identification and add the plugin
' to the main form.

Dim objTemp As Object
Dim sTemp As String
Dim sPlugin As String

'Now, we loop through all the plugin files and add them to the menus.
' In addition to this, we call a common function on the plugins that
' Identifies the plugins for us.
Dim s As String

s = Dir(App.Path & "\plugins\")
Do Until s = ""
  If Right(s, 4) = ".dll" Then
    sPlugin = Mid(s, 1, Len(s) - 4) & ".clsPluginInterface"
    Set objTemp = CreateObject(sPlugin)
    sTemp = objTemp.Identify ' Run the function on the plugin to get the identification
    'add the plugin to the form's menus.
    
    With frmMain.cMenu
      .AddItem .IndexForKey("mnuPlugins"), sTemp, , , sPlugin
    End With

    Set objTemp = Nothing
  End If
  s = Dir()
Loop

End Sub
Public Sub RunPlugin(sPlugin As String)

'On Error GoTo Error_H
    MsgBox sPlugin
    'Declare a clean object to use
    Dim objPlugIn As Object
    Dim strResponse As String
    ' Run the Plugin
    'Set objPlugIn = CreateObject(Combo1.Text)
    Set objPlugIn = CreateObject(sPlugin)
    strResponse = objPlugIn.Run(frmMain)
    'MsgBox FormX.Name
    'if the plug-in returns an error, let us know
    If strResponse <> vbNullString Then
        MsgBox strResponse
    End If
    
Exit Sub

Error_H:

MsgBox sPlugin & " - Error executing the plugin" & vbCrLf & Err.Description

End Sub

