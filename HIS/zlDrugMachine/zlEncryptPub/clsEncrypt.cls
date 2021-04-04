VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEncrypt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const cstBase64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

Private mclsAES As clsAES
Private mclsDES As clsDES

Public Property Get AES() As clsAES
    Set AES = mclsAES
End Property

Public Property Get DES() As clsDES
    Set DES = mclsDES
End Property

Public Function Base64Encode(strSource As String) As String
    Dim arrBase64() As String
    Dim arrB() As Byte, bTmp(2) As Byte, bT As Byte
    Dim i As Long, j As Long
    
    On Error Resume Next
    
    If UBound(arrBase64) = -1 Then
        arrBase64 = Split(StrConv(cstBase64, vbUnicode), vbNullChar)
    End If
    
    arrB = StrConv(strSource, vbFromUnicode)

    j = UBound(arrB)
    For i = 0 To j Step 3
        Erase bTmp
        bTmp(0) = arrB(i + 0)
        bTmp(1) = arrB(i + 1)
        bTmp(2) = arrB(i + 2)

        bT = (bTmp(0) And 252) / 4
        Base64Encode = Base64Encode & arrBase64(bT)

        bT = (bTmp(0) And 3) * 16
        bT = bT + bTmp(1) \ 16
        Base64Encode = Base64Encode & arrBase64(bT)

        bT = (bTmp(1) And 15) * 4
        bT = bT + bTmp(2) \ 64
        If i + 1 <= j Then
            Base64Encode = Base64Encode & arrBase64(bT)
        Else
            Base64Encode = Base64Encode & "="
        End If

        bT = bTmp(2) And 63
        If i + 2 <= j Then
            Base64Encode = Base64Encode & arrBase64(bT)
        Else
            Base64Encode = Base64Encode & "="
        End If
    Next
End Function

Public Function Base64Decode(strEncoded As String) As String '??
    Dim arrB() As Byte, bTmp(3) As Byte, bT As Long, bRet() As Byte
    Dim i As Long, j As Long
    
    On Error Resume Next
    
    arrB = StrConv(strEncoded, vbFromUnicode)
    j = InStr(strEncoded & "=", "=") - 2
    ReDim bRet(j - j \ 4 - 1)
    For i = 0 To j Step 4
        Erase bTmp
        bTmp(0) = (InStr(cstBase64, Chr(arrB(i))) - 1) And 63
        bTmp(1) = (InStr(cstBase64, Chr(arrB(i + 1))) - 1) And 63
        bTmp(2) = (InStr(cstBase64, Chr(arrB(i + 2))) - 1) And 63
        bTmp(3) = (InStr(cstBase64, Chr(arrB(i + 3))) - 1) And 63

        bT = bTmp(0) * 2 ^ 18 + bTmp(1) * 2 ^ 12 + bTmp(2) * 2 ^ 6 + bTmp(3)

        bRet((i \ 4) * 3) = bT \ 65536
        bRet((i \ 4) * 3 + 1) = (bT And 65280) \ 256
        bRet((i \ 4) * 3 + 2) = bT And 255
    Next
    Base64Decode = StrConv(bRet, vbUnicode)
End Function

Private Sub Class_Initialize()
    Set mclsAES = New clsAES
    Set mclsDES = New clsDES
End Sub

Private Sub Class_Terminate()
    Set mclsAES = Nothing
    Set mclsDES = Nothing
End Sub
