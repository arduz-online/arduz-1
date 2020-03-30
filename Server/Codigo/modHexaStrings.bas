Attribute VB_Name = "modHexaStrings"

Option Explicit

Public Function hexMd52Asc(ByVal MD5 As String) As String
    Dim i As Long
    Dim L As String
    
    If Len(MD5) And &H1 Then MD5 = "0" & MD5
    
    For i = 1 To Len(MD5) \ 2
        L = mid$(MD5, (2 * i) - 1, 2)
        hexMd52Asc = hexMd52Asc & Chr$(hexHex2Dec(L))
    Next i
End Function

Public Function hexHex2Dec(ByVal hex As String) As Long
    hexHex2Dec = Val("&H" & hex)
End Function

Public Function txtOffset(ByVal Text As String, ByVal off As Integer) As String
    Dim i As Long
    Dim L As String
    
    For i = 1 To Len(Text)
        L = mid$(Text, i, 1)
        txtOffset = txtOffset & Chr$((Asc(L) + off) And &HFF)
    Next i
End Function
