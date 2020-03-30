Attribute VB_Name = "Module2"
Option Explicit

' Constants for Registry top-level keys
Private Const HKEY_CLASSES_ROOT = &H80000000

' Return values
Private Const ERROR_SUCCESS = 0&

' Registry security attributes
Private Const SYNCHRONIZE = &H100000
Private Const KEY_NOTIFY = &H10
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_QUERY_VALUE = &H1
Private Const READ_CONTROL = &H20000
Private Const KEY_READ = ((READ_CONTROL Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

'Registry API Calls
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" _
        Alias "RegOpenKeyExA" _
        (ByVal hKey As Long, ByVal lpSubKey As String, _
        ByVal ulOptions As Long, ByVal samDesired As Long, _
        phkResult As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" _
        Alias "RegQueryValueExA" _
        (ByVal hKey As Long, ByVal lpszValueName As String, _
        ByVal lpdwReserved As Long, lpdwType As Long, _
        lpData As Any, lpcbData As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub LaunchURL(URL As String)

Dim KeyHandle As Long, ValueType As Long
Dim StringSize As Long, RegString() As Byte, myString As String
  
  If RegOpenKeyEx(HKEY_CLASSES_ROOT, "http\shell\open\command", 0, KEY_READ, KeyHandle) = ERROR_SUCCESS Then
    
    Call RegQueryValueEx(KeyHandle, "", 0, ValueType, StringSize, StringSize)
    ReDim RegString(0 To StringSize - 1)
    
    Call RegQueryValueEx(KeyHandle, "", 0, ValueType, RegString(0), StringSize)
    myString = Left$(StrConv(RegString, vbUnicode), UBound(RegString))
    
    Call RegCloseKey(KeyHandle)
    
    If Left$(myString, 1) = Chr$(34) Then
        'exe name is around quotes
        myString = Left$(myString, InStrRev(myString, """", , vbBinaryCompare) - 1)
        myString = Right$(myString, Len(myString) - 1)
    Else
        'no quotes - uses Dos 8.3 file name
        myString = Left$(myString, InStr(myString, " ") - 1)
    End If
    
    ShellExecute 0, "open", myString, URL, "", vbNormalFocus
  
  End If

End Sub


