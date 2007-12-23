VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFNT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const PALETTE_SIZE = 768
Private Const PALETTE_USELESS = 576

Private m_Palette() As Byte
Private m_FilePath As String
Private m_IsLoaded As Boolean
Private m_LastError As String

Private m_Info As Long

Private m_Maps As Collection
Private m_Chars(255) As T_CHAR

Private Type T_CHAR
    Width As Long
    Height As Long
    LineBase As Long
    offset As Long
End Type

Public Property Get Map(lIndex As Long) As cMap
    If Not (lIndex + 1) > m_Maps.Count Then
        Set Map = m_Maps.Item(lIndex + 1)
    Else
      '  MsgBox "Item " & CStr(i) & " doesn't exist", vbCritical
    End If
End Property

Public Property Get MapCount() As Long
    MapCount = m_Maps.Count
End Property

'-------------------------------------------------------------------------------------
'FUNCTION: Load()
'DESCRIPTION: Loads an FPG file
'RETURNS: -1 if no error, otherwise 0.
'-------------------------------------------------------------------------------------
Public Function Load(sFile As String) As Long
    Dim lFile As Long
    Dim i As Integer
    'MAP INFO
    Dim bMData() As Byte, lMLen As Long
    Dim magic As String * 8, paluseless(PALETTE_USELESS) As Byte
    Dim m As cMap
    Dim lSucceded As Long, lError As Long
    
    lFile = gzopen(sFile, "rb")
    
    If Not (lFile = vbNull) Then
        'Dispose
        m_FilePath = sFile
        
        gzread lFile, ByVal magic, Len(magic)
        'COMPROBACIÓN DE LA CABECERA
        If StrComp(Left(magic, 3), FNT_MAGIC) = 0 Then
            'Read the palette
            ReDim m_Palette(PALETTE_SIZE) As Byte
            If (gzread(lFile, m_Palette(0), PALETTE_SIZE) < PALETTE_SIZE) Then lError = -1
            If (gzread(lFile, paluseless(0), PALETTE_USELESS) < PALETTE_USELESS) Then lError = -1
            If lError Then SetLastError "No palete"
        Else
            lError = -1
            SetLastError "FILE INCOMPATIBLE"
        End If
        gzread lFile, m_Info, 4
        For i = 0 To 255
            gzread lFile, ByVal VarPtr(m_Chars(i).Width), 4
            gzread lFile, ByVal VarPtr(m_Chars(i).Height), 4
            gzread lFile, ByVal VarPtr(m_Chars(i).LineBase), 4
            gzread lFile, ByVal VarPtr(m_Chars(i).offset), 4
        Next
        For i = 0 To 255
            Set m = New cMap
            lMLen = m_Chars(i).Width * m_Chars(i).Height
            If Not lMLen = 0 Then
                ReDim bMData(lMLen - 1) As Byte
                gzseek lFile, m_Chars(i).offset, 0
                gzread lFile, bMData(0), lMLen
                If m.CreateFromStream8("Caracter", 0, _
                        m_Chars(i).Width, m_Chars(i).Height, bMData, m_Palette) = 0 Then
                    lError = -1
                End If
                If lError = 0 Then
                    m_Maps.Add m
                Else
                    SetLastError "Error u" & ": " & m.GetLastError
                    Exit For
                End If
            End If
            Set m = Nothing
        Next
        gzclose (lFile)
    Else 'Fichero no encontrado
        SetLastError "Not found"
    End If
    
    If lError Then m_IsLoaded = False
    Load = IIf(lError = 0, -1, 0)
End Function

Private Sub Class_Initialize()
    Set m_Maps = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_Maps = Nothing
End Sub
'---------------------------------------------------------------------
Private Sub SetLastError(sError As String)
    m_LastError = sError
End Sub

Public Function GetLastError() As String
    GetLastError = m_LastError
End Function
'--------------------------------------------------------------------

