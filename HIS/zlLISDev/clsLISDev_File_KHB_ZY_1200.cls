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

' 大足二院 KHB_ZY_1200

Private mdata As String
Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, ByVal dtStart As Date, _
                                Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
    Dim k As Integer, i As Integer
    Dim dt_日期 As Date, str_文件名 As String, str_检验时间 As String, str_检验标本
    Dim aResults() As String, strFilePath As String, objFile As New Scripting.FileSystemObject, strFileErs As String
    
    Dim blean As Boolean, yn As Boolean
    Dim sz_结果, sz_项目
    Dim lngTime As Integer, lngmm As Integer
    Dim strz() As String
    Dim str_日期 As Date
    Dim strz1() As String
    
    dt_日期 = dtStart
    dtStart = Format(dtStart, "yyyy-mm-dd")
    dtEnd = Format(dtEnd, "yyyy-mm-dd")
    strFilePath = objFile.GetParentFolderName(strFile) '取路径
    
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    
    strFileErs = strFile
    str_文件名 = ""
   
    If Trim(strFileErs) <> "" Then
        strFileErs = strFile
        str_文件名 = strFile
    Else
       strFileErs = ""
    End If
    ReDim aResults(0)
    
    If str_文件名 <> "" Then
          mdata = ""
          Open str_文件名 For Binary As #1
          mdata = Input(LOF(1), #1)
          Close #1
          
          
           strz1 = Split(mdata, "--------" & Chr(13) & Chr(10))
           
           For i = 1 To UBound(strz1)
             If dtStart <= Format(Mid(Split(strz1(i), Chr(13) & Chr(10))(0), InStr(Split(strz1(1), Chr(13) & Chr(10))(0), ":") + 1), "yyyy-mm-dd") And Format(Mid(Split(strz1(i), Chr(13) & Chr(10))(0), InStr(Split(strz1(1), Chr(13) & Chr(10))(0), ":") + 1), "yyyy-mm-dd") <= dtEnd Then
                str_检验时间 = Format(Mid(Split(strz1(i), Chr(13) & Chr(10))(0), InStr(Split(strz1(1), Chr(13) & Chr(10))(0), ":") + 1), "yyyy-mm-dd")
                str_检验标本 = str_检验时间 & "|" & Mid(Split(strz1(i), Chr(13) & Chr(10))(1), InStr(Split(strz1(i), Chr(13) & Chr(10))(1), ":") + 1) & "| |" & "血液|0"
                
                strz = Split(strz1(i), Chr(13) & Chr(10))
                
                For k = 2 To UBound(strz) - 1
                    sz_结果 = Mid(strz(k), InStr(strz(k), "=") + 1)
                    sz_项目 = Mid(strz(k), 1, InStr(strz(k), "=") - 1)
                    str_检验标本 = str_检验标本 & "|" & sz_项目 & "|" & sz_结果
                Next k
                
                If aResults(UBound(aResults)) <> "" Then ReDim Preserve aResults(UBound(aResults) + 1)
                 aResults(UBound(aResults)) = str_检验标本
                
             End If
           Next i
    
         ResultFromFile = aResults
         str_文件名 = ""
    End If

End Function

Public Function GetStartCmd() As String
    '设备标准接口
    '功能：获取设备开始发送命令
    GetStartCmd = ""
End Function

Public Function GetEndCmd() As String
    '设备标准接口
    '功能：获取设备停止发送命令
    GetEndCmd = ""
End Function












