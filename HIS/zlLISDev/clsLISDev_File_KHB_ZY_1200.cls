VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_KHB_ZY_1200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' �����Ժ KHB_ZY_1200

Private mdata As String
Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, ByVal dtStart As Date, _
                                Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
    Dim k As Integer, i As Integer
    Dim dt_���� As Date, str_�ļ��� As String, str_����ʱ�� As String, str_����걾
    Dim aResults() As String, strFilePath As String, objFile As New Scripting.FileSystemObject, strFileErs As String
    
    Dim blean As Boolean, yn As Boolean
    Dim sz_���, sz_��Ŀ
    Dim lngTime As Integer, lngmm As Integer
    Dim strz() As String
    Dim str_���� As Date
    Dim strz1() As String
    
    dt_���� = dtStart
    dtStart = Format(dtStart, "yyyy-mm-dd")
    dtEnd = Format(dtEnd, "yyyy-mm-dd")
    strFilePath = objFile.GetParentFolderName(strFile) 'ȡ·��
    
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    strFileErs = strFile
    str_�ļ��� = ""
   
    If Trim(strFileErs) <> "" Then
        strFileErs = strFile
        str_�ļ��� = strFile
    Else
       strFileErs = ""
    End If
    ReDim aResults(0)
    
    If str_�ļ��� <> "" Then
          mdata = ""
          Open str_�ļ��� For Binary As #1
          mdata = Input(LOF(1), #1)
          Close #1
          
          
           strz1 = Split(mdata, "--------" & Chr(13) & Chr(10))
           
           For i = 1 To UBound(strz1)
             If dtStart <= Format(Mid(Split(strz1(i), Chr(13) & Chr(10))(0), InStr(Split(strz1(1), Chr(13) & Chr(10))(0), ":") + 1), "yyyy-mm-dd") And Format(Mid(Split(strz1(i), Chr(13) & Chr(10))(0), InStr(Split(strz1(1), Chr(13) & Chr(10))(0), ":") + 1), "yyyy-mm-dd") <= dtEnd Then
                str_����ʱ�� = Format(Mid(Split(strz1(i), Chr(13) & Chr(10))(0), InStr(Split(strz1(1), Chr(13) & Chr(10))(0), ":") + 1), "yyyy-mm-dd")
                str_����걾 = str_����ʱ�� & "|" & Mid(Split(strz1(i), Chr(13) & Chr(10))(1), InStr(Split(strz1(i), Chr(13) & Chr(10))(1), ":") + 1) & "| |" & "ѪҺ|0"
                
                strz = Split(strz1(i), Chr(13) & Chr(10))
                
                For k = 2 To UBound(strz) - 1
                    sz_��� = Mid(strz(k), InStr(strz(k), "=") + 1)
                    sz_��Ŀ = Mid(strz(k), 1, InStr(strz(k), "=") - 1)
                    str_����걾 = str_����걾 & "|" & sz_��Ŀ & "|" & sz_���
                Next k
                
                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                 aResults(UBound(aResults)) = str_����걾
                
             End If
           Next i
    
         ResultFromFile = aResults
         str_�ļ��� = ""
    End If

End Function

Public Function GetStartCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸��ʼ��������
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    '�豸��׼�ӿ�
    '���ܣ���ȡ�豸ֹͣ��������
    GetEndCmd = ""
End Function











