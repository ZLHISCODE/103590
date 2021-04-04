Attribute VB_Name = "mdlClipBoard"
'///////////////////////////////////////////////////////////////////////////////
'
'       模块：剪贴板操作
'       功能：剪贴板操作,复制文件目录到剪贴板
'       编写：祝庆
'       日期：2011年1月3日
'
'///////////////////////////////////////////////////////////////////////////////
Option Explicit

Public Function clipClear() As Boolean
'清空当前剪贴板
    Call EmptyClipboard
End Function

Public Function clipCopyFiles(File() As String) As Boolean
'复制多个文件到剪贴板
   On Error Resume Next
   Dim strData As String
   Dim df As DROPFILES
   Dim hGlobal As Long
   Dim lpGlobal As Long
   Dim i As Long
   strData = ""

   
   '清除剪贴版中现存的数据
   If OpenClipboard(0&) Then
        '清空当前剪贴板
        Call EmptyClipboard
        
        '判断文件数组是否为空
        If SafeArrayGetDim(File) = 0 Then Exit Function
        For i = LBound(File) To UBound(File)
            strData = strData & File(i) & vbNullChar
        Next
        
        hGlobal = GlobalAlloc(GHND, Len(df) + LenB(strData))
        
        If hGlobal Then
            lpGlobal = GlobalLock(hGlobal)
         
            df.pFiles = Len(df)
            Call CopyMemory(ByVal lpGlobal, df, Len(df))
            Call CopyMemory(ByVal (lpGlobal + Len(df)), ByVal strData, LenB(strData))
   
            Call GlobalUnlock(hGlobal)
         
            If SetClipboardData(CF_HDROP, hGlobal) Then
                clipCopyFiles = True
            End If

        End If
        
        Call CloseClipboard
    End If
End Function


