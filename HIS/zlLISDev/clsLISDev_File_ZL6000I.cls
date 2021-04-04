VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_ZL6000I"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:血流变ZL6000I
'厂商:
'作者:董雨
'用户:长安区医院

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
    Dim cnAccess As New ADODB.Connection
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsVisc As ADODB.Recordset
    Dim lngResultCount As Long, strBh As String
    Dim blnUnion As Boolean
    Dim intLoop As Integer
    Dim str图形  As String

    On Error GoTo errHandle
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path

    Set cnAccess = New ADODB.Connection
    With cnAccess
        .ConnectionString = "Provider=MSDASQL.1;Persist Security Info=True;User ID=admin;Extended Properties=" _
                            & """" & "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}" & """"
        .Open
        If .State = adStateClosed Then Exit Function
        strSQL = "": blnUnion = False

        strSQL = "select * from testdata WHERE " & _
                 IIf(lngSampleNO = -1, "  日期 between #" & Format(dtStart, "yyyy-mm-dd") & "#" & _
                 " And #" & Format(dtEnd, "yyyy-mm-dd") & "#", _
                 " 序号 ='" & lngSampleNO & "'") & " ORDER BY 日期,序号"
        Set rsTmp = .Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            ReDim Preserve aResults(0)
            ResultFromFile = aResults
            Exit Function
        End If
        lngResultCount = -1

        ReDim Preserve aResults(0)
        rsTmp.MoveFirst
        Do Until rsTmp.EOF
            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)
            aResults(lngResultCount) = Format(Nvl(rsTmp!日期) & " " & Nvl(rsTmp!时间), "YYYY-MM-DD HH:NN:SS")    '检验时间
            '            aResults(lngResultCount) = Nvl(rsTmp!日期) & Nvl(rsTmp!日期) & " " & Nvl(rsTmp!时间)    '检验时间
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & Nvl(rsTmp!序号)             '样本序号
            aResults(lngResultCount) = aResults(lngResultCount) & "|" & Nvl(rsTmp!检验医生)         '检验人
            aResults(lngResultCount) = aResults(lngResultCount) & "|血液|0"
            '全血粘度曲线
            strSQL = "select * from Visc where TestDataID =" & rsTmp!TestDataID & " order by ShearRate"
            Set rsVisc = .Execute(strSQL)
            Do Until rsVisc.EOF
                aResults(lngResultCount) = aResults(lngResultCount) & "|" & rsVisc!ShearRate & "|" & Nvl(rsVisc!Visc)
                rsVisc.MoveNext
            Loop
            '其他结果
            aResults(lngResultCount) = aResults(lngResultCount) & "|全血粘度|" & Nvl(rsTmp!血浆粘度, " ") & "|血沉|" & Nvl(rsTmp!血沉, " ")
            aResults(lngResultCount) = aResults(lngResultCount) & "|压积|" & Nvl(rsTmp!压积, " ") & "|全血高切相对指数|" & Nvl(rsTmp!全血高切相对指数, " ")
            aResults(lngResultCount) = aResults(lngResultCount) & "|全血低切相对指数|" & Nvl(rsTmp!全血低切相对指数, " ") & "|血沉方程K值|" & Nvl(rsTmp!血沉方程K值, " ")
            aResults(lngResultCount) = aResults(lngResultCount) & "|红细胞聚集指数|" & Nvl(rsTmp!红细胞聚集指数, " ") & "|红细胞变形指数|" & Nvl(rsTmp!红细胞变形指数, " ")
            aResults(lngResultCount) = aResults(lngResultCount) & "|全血低切还原粘度|" & Nvl(rsTmp!全血低切还原粘度, " ") & "|全血高切还原粘度|" & Nvl(rsTmp!全血高切还原粘度, " ")
            aResults(lngResultCount) = aResults(lngResultCount) & "|红细胞变形指数TK|" & Nvl(rsTmp!红细胞变形指数TK, " ") & "|红细胞刚性指数|" & Nvl(rsTmp!红细胞刚性指数, " ")
            aResults(lngResultCount) = aResults(lngResultCount) & "|卡松粘度|" & Nvl(rsTmp!卡松粘度, " ") & "|纤维蛋白原|" & Nvl(rsTmp!纤维蛋白原, " ")
            aResults(lngResultCount) = aResults(lngResultCount) & "|血胆固醇|" & Nvl(rsTmp!血胆固醇, " ") & "|甘油三脂|" & Nvl(rsTmp!甘油三脂, " ")

            rsTmp.MoveNext
        Loop
    End With

    ResultFromFile = aResults
    Exit Function
errHandle:
    WriteLog "clsLISDev_File_ZL6000I", strFile, Err.Number & " " & Err.Description
End Function
