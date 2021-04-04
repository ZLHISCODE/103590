VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsErrInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type ERROBJ
    StrObjName As String
    LngObjIndex As Long
    LngRow As Long
    LngCol As Long
    OldColor As Long
    PicIndex As Long
End Type

Private objErrArray() As ERROBJ
Public IntErrType As Integer       '错误信息
Public StrErrInfo As String        '0-错误；1-警告

Public lngObjNum As Long
Public strErrID As String

Public Function AddErrObj(objErr As Object) As Boolean
    Dim objTmp As ERROBJ
    Dim strTmp As String
    
    '重庆七院XP报错，要先访问typename
    strTmp = TypeName(objErrArray)
    ReDim Preserve objErrArray(lngObjNum)
    lngObjNum = lngObjNum + 1

    objTmp.StrObjName = objErr.Name
    On Error Resume Next
    objTmp.LngObjIndex = objErr.Index
    If Err.Number <> 0 Then
        objTmp.LngObjIndex = -1
        Err = 0
    End If
    objTmp.OldColor = objErr.BackColor
    On Error GoTo 0
    
    If UCase(TypeName(objErr)) = UCase("VSFlexGrid") Then
        objTmp.LngRow = objErr.Row
        objTmp.LngCol = objErr.Col
    End If
    
    If gBlnNew And Not gfrmMecCol Is Nothing Then
        If objErr.Tag <> "" Then
            If UBound(Split(objErr.Tag, "|")) = 2 Then
                objTmp.PicIndex = Val(Split(objErr.Tag, "|")(2))
            End If
            
        End If
    End If
    
    objErrArray(UBound(objErrArray)) = objTmp
End Function

Private Sub Class_Initialize()
    lngObjNum = 0
End Sub

Public Function GetObjErr() As Variant
    GetObjErr = objErrArray()
End Function
