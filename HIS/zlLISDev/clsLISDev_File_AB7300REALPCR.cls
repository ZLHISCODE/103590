VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_AB7300REALPCR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:AB7300REALPCR
'作者:李骏驰
'用户:大理州人民医院

Private Type T标本
    T标本号 As String
    T项目数 As Integer
    T结果 As String
    T数据() As String
    T项目() As String
End Type

Private M标本() As T标本

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

Public Function ResultFromFile(ByVal strFile As String, ByVal lngSampleNO As Long, _
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
    
    Dim aResults() As String
    Dim strChenel As String, str操作员 As String, str日期 As String
    Dim i As Integer, str编号 As String, j As Integer, k As Integer
    Dim str标本号 As String, str结果 As String, blnReturn As Boolean
    Dim lngResultCount As Long, str格式1 As String, str格式2 As String
    Dim lngRtCount As Long, strFilePath As String
    Dim myArray() As String, strFilename As String
    Dim MyColor(2) As Long
    Dim objStream As textStream, objFile As New Scripting.FileSystemObject
    Dim strLine As String
    Dim objForm   As New frmGraph
    Dim varline As Variant
    
    

    On Error GoTo errHandle
    '避免出现下标错误
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    str日期 = Format(Now, "yyyy-MM-dd hh:mm:ss")

    MyColor(0) = vbRed
    MyColor(1) = vbYellow
    MyColor(2) = vbBlue

    str格式1 = Format(Now, "yyyymmdd")
    str格式2 = Format(Now, "yyyy-mm-dd")


    '图形数据
    Set objStream = objFile.OpenTextFile(strFilePath & "\" & str格式1 & ".csv", ForReading)
    lngResultCount = 0
    Do Until objStream.AtEndOfStream
        strLine = objStream.ReadLine
        If strLine <> "" Then
            varline = Split(strLine, ",")
            If varline(0) <> "Well" Then
                    ReDim Preserve myArray(lngResultCount)
                    myArray(lngResultCount) = varline(0)
                For i = 3 To UBound(varline)
                    myArray(lngResultCount) = myArray(lngResultCount) & "," & Round(varline(i) / 1000, 0)
                Next i
                lngResultCount = lngResultCount + 1
            End If
        End If
    Loop

    '标本数据
    Set objStream = objFile.OpenTextFile(strFilePath & "\" & str格式2 & ".csv", ForReading)
    lngRtCount = 0
    Do Until objStream.AtEndOfStream
            strLine = objStream.ReadLine
            If j > 29 Then
            varline = Split(strLine, ",")
            If varline(1) <> "" Then
            If CInt(Val(Mid(varline(1), 1, 3))) > 0 Then
                
                str编号 = (Asc(Mid(varline(0), 1, 1)) - 65) * 12 + Mid(varline(0), 2)
                str标本号 = CInt(Val(Mid(varline(1), 1, 3)))
                str结果 = Mid(varline(1), 4) & "|" & varline(4)
                For i = 0 To lngRtCount - 1
                    If M标本(i).T标本号 = str标本号 Then
                        M标本(i).T标本号 = str标本号
                        M标本(i).T结果 = M标本(i).T结果 & str结果
                        M标本(i).T项目数 = Val(M标本(i).T项目数) + 1
                        '标本图形数据
                        For k = 0 To lngResultCount - 1
                            If Trim(str编号) = Trim(Split(myArray(k), ",")(0)) Then
                                ReDim Preserve M标本(i).T项目(M标本(i).T项目数)
                                ReDim Preserve M标本(i).T数据(M标本(i).T项目数)
                                M标本(i).T项目(M标本(i).T项目数) = Mid(varline(1), 4)
                                M标本(i).T数据(M标本(i).T项目数) = Mid(myArray(k), InStr(1, myArray(k), ",") + 1)
                            End If
                        Next k
                        Exit For
                    End If
                Next i
                
                If i >= lngRtCount Then
                    ReDim Preserve M标本(lngRtCount)
                    M标本(lngRtCount).T标本号 = str标本号
                    M标本(lngRtCount).T结果 = str结果
                    M标本(lngRtCount).T项目数 = 0
                    '标本图形数据
                    For k = 0 To lngResultCount - 1
                        If Trim(str编号) = Trim(Split(myArray(k), ",")(0)) Then
                            ReDim Preserve M标本(lngRtCount).T数据(M标本(lngRtCount).T项目数)
                            ReDim Preserve M标本(lngRtCount).T项目(M标本(lngRtCount).T项目数)
                            M标本(lngRtCount).T项目(M标本(lngRtCount).T项目数) = Mid(varline(1), 4)
                            M标本(lngRtCount).T数据(M标本(lngRtCount).T项目数) = Mid(myArray(k), InStr(1, myArray(k), ",") + 1)
                        End If
                    Next k
                    lngRtCount = lngRtCount + 1
                End If
            End If
            End If
        End If
        j = j + 1
    Loop
    
    For i = 0 To lngRtCount - 1
        ReDim Preserve aResults(i)
        aResults(i) = str日期
        aResults(i) = aResults(i) & "|" & M标本(i).T标本号                '样本序号
        aResults(i) = aResults(i) & "|" & " "                             '检验人
        aResults(i) = aResults(i) & "|" & "血液"                          '标本
        aResults(i) = aResults(i) & "|0|"                                  '质控
        aResults(i) = aResults(i) & M标本(i).T结果                        '项目和结果
        '先画线
        objForm.Draw
        '再画图
        objForm.DrawGraph M标本(i).T数据, MyColor, M标本(i).T项目
        strFilename = strFilePath & "\" & M标本(i).T标本号 & ".bmp"
        '保存图像
        SavePicture objForm.Picture1.Image, strFilename
        aResults(i) = aResults(i) & vbCrLf & " ;103;1;" & strFilename
        '清空图片控件中的图像
        objForm.Picture1.Cls
        'frmGraph荧光.Draw
    Next i
    '释放资源占用
    Unload objForm
    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_AB7300REALPCR.ResultFromFile", strFile, Err.Number & " " & Err.Description
End Function








