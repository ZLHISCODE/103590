Attribute VB_Name = "mdlPubString"
Option Explicit

'本模块保存 字符串处理有关的公共函数 ,此模块中的函数均以 PStr_开头

Public Function PStr_CutCode(ByRef strIn As String, strS As String, strE As String) As String
    '按指定的开始符，结束符，截取一段字符
    '成功返回截取的字符串
    Dim lngS As Long, lngE As Long
    lngE = 0: lngS = 0
    lngS = InStr(strIn, strS)
    If lngS > 0 Then lngE = InStr(lngS, strIn, strE)
    PStr_CutCode = ""
    If lngS > 0 And lngE > 0 Then
        PStr_CutCode = Mid(strIn, lngS, lngE - lngS + Len(strE))
        strIn = Mid(strIn, lngE + Len(strE))
    End If
    
End Function

