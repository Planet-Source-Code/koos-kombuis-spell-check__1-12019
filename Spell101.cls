VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Spell101"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function RegOpenKey Lib _
"advapi32" Alias "RegOpenKeyA" (ByVal hKey _
As Long, ByVal lpSubKey As String, _
phkResult As Long) As Long

Private Declare Function RegQueryValueEx _
Lib "advapi32" Alias "RegQueryValueExA" _
(ByVal hKey As Long, ByVal lpValueName As _
String, lpReserved As Long, lptype As _
Long, lpData As Any, lpcbData As Long) _
As Long

Private Declare Function RegCloseKey& Lib _
"advapi32" (ByVal hKey&)

Private Const REG_SZ = 1
Private Const REG_EXPAND_SZ = 2
Private Const ERROR_SUCCESS = 0

Private Const HKEY_CLASSES_ROOT = &H80000000


Public Function GetRegString(hKey As Long, _
strSubKey As String, strValueName As _
String) As String

Dim strSetting As String
Dim lngDataLen As Long
Dim lngRes As Long
    If RegOpenKey(hKey, strSubKey, _
    lngRes) = ERROR_SUCCESS Then
       strSetting = Space(255)
       lngDataLen = Len(strSetting)
       If RegQueryValueEx(lngRes, _
       strValueName, ByVal 0, _
       REG_EXPAND_SZ, ByVal strSetting, _
       lngDataLen) = ERROR_SUCCESS Then
          If lngDataLen > 1 Then
          GetRegString = Left(strSetting, lngDataLen - 1)
       End If
    End If
    
    If RegCloseKey(lngRes) <> ERROR_SUCCESS Then
       MsgBox "RegCloseKey Failed: " & _
       strSubKey, vbCritical
    End If
    End If
End Function



Public Function MsSpellCheck(strText As String) As String
    Dim oWord As Object
    Dim strSelection As String
    Set oWord = CreateObject("Word.Basic")
    
    If IsAppPresent("Word.Document\CurVer", "") <> False Then
    
        oWord.AppMinimize
        MsSpellCheck = strText
        oWord.FileNewDefault
        oWord.EditSelectAll
        oWord.EditCut
        oWord.Insert strText
        oWord.StartOfDocument
        On Error Resume Next
        oWord.ToolsSpelling
        On Error GoTo 0
        oWord.EditSelectAll
        strSelection = oWord.Selection$
    
    
        If Mid(strSelection, Len(strSelection), 1) = Chr(13) Then
            strSelection = Mid(strSelection, 1, Len(strSelection) - 1)
        End If
    
    
        If Len(strSelection) > 1 Then
            MsSpellCheck = strSelection
        End If
        oWord.FileCloseAll 2
        oWord.AppClose
        Set oWord = Nothing
    Else
        MsSpellCheck = strText
    End If
End Function


Public Function FileExists(sFileName$) As Boolean
    On Error Resume Next
    FileExists = IIf(Dir(Trim(sFileName)) <> "", True, False)
End Function


Public Function IsAppPresent(strSubKey$, strValueName$) As Boolean
    IsAppPresent = CBool(Len(GetRegString(HKEY_CLASSES_ROOT, strSubKey, strValueName)))
End Function


