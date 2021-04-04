VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Format(strExp As String, StrFmt As String) As String
    Format = VBA.Format(strExp, StrFmt)
End Function

Public Function IIF(varExp As Boolean, varExpA As Variant, varExpB As Variant) As Variant
    IIF = VBA.IIF(varExp, varExpA, varExpB)
End Function

Public Function Val(strVal As String) As Double
    If IsNumeric(strVal) Then
        If strVal Like "*,*" Then strVal = CDbl(strVal)
    End If
    Val = VBA.Val(strVal)
End Function

