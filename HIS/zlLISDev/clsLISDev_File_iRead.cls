VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_iRead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:iRead
'厂商:未知
'作者:陈东
'用户:妇儿医院

'文件名为iRead.txt,文件名固定
'

'文件内容如下
'1001,CRP500,<2.50,2014-11-05 19:23:17
'1002,CRP500,<2.50,2014-11-05 20:02:29
'1003,CRP500,<2.50,2014-11-05 20:08:14
'1004,CRP500,<2.50,2014-11-05 20:30:13
'100000159175,CRP500,<5.0,2014-11-06 11:56:22
'100000159175,HsCRP,<2.50,2014-11-06 11:56:22
'100000173653,CRP500,49.45,2014-11-08 11:13:14
'100000173653,HsCRP,>5.0,2014-11-08 11:13:14
'100000173655,CRP500,<5.0,2014-11-08 11:16:56
'100000173655,HsCRP,<2.50,2014-11-08 11:16:56
'1001,CRP500,<5.0,2014-11-08 11:17:56
'1001,HsCRP,<2.50,2014-11-08 11:17:56


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


Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNo As Long, _
    ByVal dtStart As Date, Optional ByVal dtEnd As Date = CDate("3000-12-31")) As String()
    '标准函数：从文件获取检验结果
    '   strFile：包含路径的文件名
    '   lngSampleNO：标本号。为-1表示选取所有时间范围内的标本
    '   dtStart：开始时间。如果指定了标本号，则只提取该参数当天的这个标本（dtEnd无效）
    '   dtEnd：结束时间。只有当选取多个标本（lngSampleNO=-1）时，该参数有效。如果不指定该参数值，则时间范围为>=dtStart。
    '   返回的检验结果数组(各解析程序必须按以下标准组织结果)
    '   元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果
    
        Dim aResults() As String, objFile As New Scripting.FileSystemObject
        Dim strChenel As String, strDate As String
        Dim strFilePath As String, strTmp As String
    
        Dim objStream As textStream
        Dim strLine As String, varItem As Variant, varTmp As Variant
        Dim strSampleNo As String, strFileName As String
        Dim lngRow As Long, varLine As Variant
    
 
        On Error GoTo errHandle
        '避免出现下标错误
100     ReDim aResults(0)
102     aResults(0) = ""
104     ResultFromFile = aResults

        
106     strFilePath = objFile.GetParentFolderName(strFile)
108     If Len(strFilePath) = 0 Then strFilePath = App.Path
110     If Right$(strFilePath, 1) = "\" Then strFilePath = Mid$(strFilePath, 1, Len(strFilePath) - 1)
        
112     strFileName = strFile
114     If objFile.FileExists(strFileName) Then
116         Set objStream = objFile.OpenTextFile(strFileName, ForReading)
        
118         Do Until objStream.AtEndOfStream
120             strLine = objStream.ReadLine
122             If InStr(strLine, ",") > 0 Then
124             varLine = Split(strLine, ",")

    
126                 strSampleNo = Trim(varLine(0))
128                 If Val(strSampleNo) > 0 And UBound(varLine) >= 3 Then
130                     strDate = Trim$("" & varLine(3))
132                     If Not IsDate(strDate) Then strDate = Format(Now, "yyyy-MM-dd")
                    
134                     strChenel = Trim$(varLine(1))
                        
136                     strTmp = Trim$(varLine(2))
                        
138                     If strChenel <> "" And strTmp <> "" Then
140                         strChenel = strChenel & "|" & strTmp
    
142                         If aResults(UBound(aResults)) <> "" And strDate <> "" Then
144                             ReDim Preserve aResults(UBound(aResults) + 1)
                            End If
146                         aResults(UBound(aResults)) = strDate & "|" & strSampleNo & "| |Boold|0|" & strChenel
                        End If
     
                    End If
                End If
148             lngRow = lngRow + 1
150             DoEvents
            Loop
152         ResultFromFile = aResults
        Else
154          WriteLog " iRead,文件" & strFileName & "不存在!", "", ""
        End If
        Exit Function
errHandle:
 
156     WriteLog " iRead," & Erl() & Err.Description, "", ""
End Function

