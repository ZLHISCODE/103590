Attribute VB_Name = "mdlClipBoard"
'///////////////////////////////////////////////////////////////////////////////
'
'       ģ�飺���������
'       ���ܣ����������,�����ļ�Ŀ¼��������
'       ��д��ף��
'       ���ڣ�2011��1��3��
'
'///////////////////////////////////////////////////////////////////////////////
Option Explicit

Public Function clipClear() As Boolean
'��յ�ǰ������
    Call EmptyClipboard
End Function

Public Function clipCopyFiles(File() As String) As Boolean
'���ƶ���ļ���������
   On Error Resume Next
   Dim strData As String
   Dim df As DROPFILES
   Dim hGlobal As Long
   Dim lpGlobal As Long
   Dim i As Long
   strData = ""

   
   '������������ִ������
   If OpenClipboard(0&) Then
        '��յ�ǰ������
        Call EmptyClipboard
        
        '�ж��ļ������Ƿ�Ϊ��
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


