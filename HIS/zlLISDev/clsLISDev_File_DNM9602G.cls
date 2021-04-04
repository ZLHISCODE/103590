VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_DNM9602G"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'仪器:酶标仪DNM9602G
'医院:德阳人民医院(隆昌县人民医院)
'开发:李建飞
'时间:2009/10/16
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
    Dim strSQL As String, rsTmp As ADODB.Recordset, rsItem As ADODB.Recordset
    Dim lngResultCount As Long
    Dim blnUnion As Boolean
    Dim lngSampleID As Long
    Dim blnNew As Boolean           '是否新标本号
    Dim lngSampleNumber As Long     '标本ID
    Dim lngloop As Long
    Dim intLoop As Integer
    Dim strBatch As String          '记录批号
    Dim strResult As String         '结果
    Dim strOD As String             '原始结果(OD)
    Dim strCO As String             '判断值(Cut Off)
    Dim strItem As String           '项目（通道码)
    Dim strBh As String             '当前板号
    Dim strCalc As String           '计算值
    Dim dateVerify As Date          '检验时间
    Dim StrTime As String
    Dim lng_count As Long
    Dim arrTmp(95) As Integer
    Dim i As Integer, j As Integer, k As Integer
    '避免出现下标错误
    ReDim aResults(0)
    aResults(0) = ""
    lng_count = -1
    ResultFromFile = aResults
    
    '横向摆放标本
    For i = 0 To 7
        For j = 0 To 11
            arrTmp(i * 12 + j) = j * 8 + i
            'Debug.Print arrTmp(j)
        Next
    Next
    i = 0: j = 0
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function

        strSQL = "select * from 批量病人表 a , bzsz b where a.编号 = b.bh  and " & _
                 " Format(b.rq,'yyyy-MM-dd')>='" & Format(dtStart, "yyyy-MM-dd") & "'" & _
                 " And Format(b.rq,'yyyy-MM-dd')<='" & Format(dtEnd, "yyyy-MM-dd") & "'" & _
                " order by 试验编号,编号,孔号位置"
        Set rsTmp = cnAccess.Execute(strSQL)

        Do Until rsTmp.EOF
            If lngSampleNumber <> rsTmp("试验编号") Then
                lngSampleNumber = rsTmp("试验编号")
                intLoop = 1
                
                If k > 0 And i < 95 And j < 95 And rsTmp!PLH = strBatch Then
                    j = j + 1
                    i = i + 1
                Else
                    i = 0
                    j = 0
                End If
                strBatch = rsTmp!PLH
                k = 1
            End If
            If strBh <> rsTmp("编号") Then
                intLoop = 1
            End If


            '得到当前OD
            strSQL = "select distinct * from bzsj where bh = '" & rsTmp("bh") & "'"

            Set rsItem = cnAccess.Execute(strSQL)

            If rsItem.EOF = False Then
                strOD = Nvl(rsItem("SJ" & arrTmp(i)))               '(Val(Replace(rsTmp("标号"), "S", "")) - 1)
            End If

            '得到当前Cut Off值
            strItem = Nvl(rsTmp("xm" & intLoop & "c"))
            strSQL = "select * from xxsm where bh = '" & rsTmp("bh") & "' and xmbh = '" & strItem & "'"
            Set rsItem = cnAccess.Execute(strSQL)
            If rsItem.EOF = False Then
                strCO = Nvl(rsItem("co"))
            End If

            '计算值(公式: OD/Cut Off
            If Val(strCO) <> 0 Then
                strCalc = Round(Val(strOD) / Val(strCO), 3)
            Else
                strCalc = 0
            End If
            '计算时间
            dateVerify = Nvl(rsTmp("rq")) & " " & Nvl(rsTmp("time"))

            '计算结果
            strSQL = "select * from 计算结果 where bh = '" & rsTmp("bh") & "'"
            Set rsItem = cnAccess.Execute(strSQL)
            If rsItem.EOF = False Then
                strResult = Nvl(rsItem("a" & arrTmp(j)))              '(Val(Replace(rsTmp("标号"), "S", "")) - 1))
                
                If InStr(strResult, "-") > 0 Then
                    strResult = "阴性(-)"
                ElseIf InStr(strResult, "+") > 0 Then
                    strResult = "阳性(+)"
                Else
                    strResult = ""
                End If
            End If

            '得到当前的项目通道码
            strItem = Nvl(rsTmp("xm" & intLoop & "c"))
            strSQL = "select * from xmcs where xmbh = '" & strItem & "'"
            Set rsItem = cnAccess.Execute(strSQL)
            If rsItem.EOF = False Then
                strItem = Nvl(rsItem("xmmc1"))
            End If

            If strItem <> "" Then
                ReDim Preserve aResults(lng_count + 1)
                aResults(UBound(aResults)) = dateVerify & "|" & rsTmp("试验编号") & "| |血清|0|" & strItem & "|" & _
                                        strResult & "^" & strOD & "^" & strCO & "^" & strCalc
                WriteLog "解析 IMS 972", aResults(UBound(aResults)), ""
            End If
            lng_count = lng_count + 1
            intLoop = intLoop + 1
            strBh = rsTmp("编号")
            lngSampleNumber = rsTmp("试验编号")
            If rsTmp("rq") <> StrTime Then
                i = 0
                j = 0
            End If
            StrTime = rsTmp("rq")
            rsTmp.MoveNext
        Loop
        ResultFromFile = aResults
    End With
End Function


