VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_MVIS_2040"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:MVIS-2040 血流变分析
'厂商:重庆天海医疗设备有限公司
'作者:王振涛
'用户:山西盐湖人民医院
'时间：2013-5-28
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
    '返回的检验结果数组(各解析程序必须按以下标准组织结果)
    '   每组检验结果以||分隔,元素之间以|分隔
    '   第0个元素：检验时间
    '   第1个元素：样本序号
    '   第2个元素：检验人
    '   第3个元素：标本
    '   第4个元素：是否质控品
    '   从第5个元素开始为检验结果，每2个元素表示一个检验项目。
    '       如：第5i个元素为检验项目，第5i+1个元素为检验结果(微生物结果以"^"分为三段,分别是:耐药^方法^结果)
    
    '    2) 粘度特征曲线:图像名称;图像画法;座标数据;曲线及描点数据;坐标轴标题数据
    '                  其中  座标数据：Y长度,X长度|X座标-X座标显示的数字,....|Y座标-Y座标显示的数字,....
    '                  曲线及描点数据:粘度曲线1的高点和低点座标|粘度曲线2的高点和低点座标~低切点坐标,中切点坐标,高切点坐标
    '                  坐标轴标题数据:Y坐标标题文字,X坐标,Y座标~X坐标标题文字,X坐标,Y座标
    '       例:粘度特征曲线;2;20,200|20-20,40-40,60-60,80-80,100-100,120-120,140-140,160-160,180-180,200-200|2-2,4-4,6-6,8-8,10-10,12-12,14-14,16-16,18-18,20-20;
    '          9.25,10,4.4,150|6.5,10,3.65,150~10-8.989,60-4.803,150-4.05;
    '          VIS(mPa.s),25,20~SHR(1/S),195,1

    Dim aResults() As String
    Dim cnAccess As New ADODB.Connection
    Dim strFilePath As String, objFile As New Scripting.FileSystemObject
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim lngResultCount As Long, lngBh As Long
    Dim blnUnion As Boolean
    Dim i As Integer
    Dim strItems As String '字段
    Dim strImg As String
    Dim strLine_Wman As String, strLine_Man As String
    Dim arrLow(2) As Double, arrMid(2) As Double, arrHigh(2) As Double, arrNianDu(2) As Double
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        
        strSQL = " Select * From Test Where todayDate >= Cdate('" & Format(dtStart, "yyyy-mm-dd 00:00:00") & "') and  todayDate <= cdate('" & Format(dtEnd, "yyyy-mm-dd 23:59:59") & "')"
        
        If lngSampleNO <> -1 Then
           strSQL = strSQL & " and val(SN) = " & lngSampleNO
        End If
        strSQL = strSQL & " Order By todayDate,ID"
        
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        strItems = ""
        For i = 0 To rsTmp.Fields.count - 1
            strItems = strItems & "," & rsTmp.Fields(i).Name
        Next
        
        lngResultCount = -1
        Do While Not rsTmp.EOF
            If InStr(";123456789;", Mid(rsTmp!sn, 1, 1)) <= 0 Or Not IsNumeric(rsTmp!sn) Then
                '首位非具体数字就默认为质控标本
                lngBh = 999
            Else
                lngBh = rsTmp!sn
            End If
            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)
            aResults(lngResultCount) = Format(rsTmp!todayDate, "yyyy-MM-DD") & "|" & lngBh & "| " & "|血液|" & IIf(lngBh = 999, 1, 0)
            '下面开始逐项加入项目和值
            If Nvl(rsTmp!highShear, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|highShear|" & Round(rsTmp!highShear, 2)
            End If
            If Nvl(rsTmp!lowShear, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|lowShear|" & Round(rsTmp!lowShear, 2)
            End If
            If Nvl(rsTmp!plasm, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|plasm|" & Round(rsTmp!plasm, 2)
            End If
            If Nvl(rsTmp!esr, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|esr|" & Round(rsTmp!esr, 2)
            End If
            If Nvl(rsTmp!hct, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|hct|" & Round(rsTmp!hct, 2)
            End If
            If Nvl(rsTmp!middleShear1, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|middleShear1|" & Round(rsTmp!middleShear1, 2)
            End If
            If Nvl(rsTmp!middleShear10, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|middleShear10|" & Round(rsTmp!middleShear10, 2)
            End If
            If Nvl(rsTmp!middleShear30, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|middleShear30|" & Round(rsTmp!middleShear30, 2)
            End If
            If Nvl(rsTmp!middleShear50, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|middleShear50|" & Round(rsTmp!middleShear50, 2)
            End If
            If Nvl(rsTmp!middleShear100, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|middleShear100|" & Round(rsTmp!middleShear100, 2)
            End If
            If Nvl(rsTmp!reductiveHighShear, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|reductiveHighShear|" & Round(rsTmp!reductiveHighShear, 2)
            End If
            If Nvl(rsTmp!reductiveLowShear, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|reductiveLowShear|" & Round(rsTmp!reductiveLowShear, 2)
            End If
            If Nvl(rsTmp!arbc, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|arbc|" & Round(rsTmp!arbc, 2)
            End If
            If Nvl(rsTmp!tx, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|tx|" & Round(rsTmp!tx, 2)
            End If
            If Nvl(rsTmp!tk, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|tk|" & Round(rsTmp!tk, 2)
            End If
            If Nvl(rsTmp!esrK, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|esrK|" & Round(rsTmp!esrK, 2)
            End If
            If Nvl(rsTmp!hemoglobin, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|hemoglobin|" & Round(rsTmp!hemoglobin, 2)
            End If
            If Nvl(rsTmp!rbc, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|rbc|" & Round(rsTmp!rbc, 2)
            End If
            If Nvl(rsTmp!eft, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|eft|" & Round(rsTmp!eft, 2)
            End If
            If Nvl(rsTmp!yieldStress, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|yieldStress|" & Round(rsTmp!yieldStress, 2)
            End If
            If Nvl(rsTmp!cassonHighShear, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|cassonHighShear|" & Round(rsTmp!cassonHighShear, 2)
            End If
            If Nvl(rsTmp!cassonMiddleShear, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|cassonMiddleShear|" & Round(rsTmp!cassonMiddleShear, 2)
            End If
            If Nvl(rsTmp!cassonLowShear, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|cassonLowShear|" & Round(rsTmp!cassonLowShear, 2)
            End If
            If Nvl(rsTmp!pr, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|pr|" & Round(rsTmp!pr, 2)
            End If
            If Nvl(rsTmp!thrombusLength, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|thrombusLength|" & Round(rsTmp!thrombusLength, 2)
            End If
            If Nvl(rsTmp!thrombusWaterWeight, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|thrombusWaterWeight|" & Round(rsTmp!thrombusWaterWeight, 2)
            End If
            If Nvl(rsTmp!thrombusWeight, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|thrombusWeight|" & Round(rsTmp!thrombusWeight, 2)
            End If
            If Nvl(rsTmp!padt, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|padt|" & Round(rsTmp!padt, 2)
            End If
            If Nvl(rsTmp!cholesterol, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|cholesterol|" & Round(rsTmp!cholesterol, 2)
            End If
            If Nvl(rsTmp!tg, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|tg|" & Round(rsTmp!tg, 2)
            End If
            If Nvl(rsTmp!glu, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|glu|" & Round(rsTmp!glu, 2)
            End If
            If Nvl(rsTmp!hdlc, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|hdlc|" & Round(rsTmp!hdlc, 2)
            End If
            
           
            '分析图形(格式：图像名称;图像画法;座标数据;曲线及描点数据;坐标轴标题数据)
            '根据性别判断改变划线坐标值
'            If rsTmp!sex = "女" Then
'               strLine_Man = "11.98,10,5.78,150|9.35,10,4.85,150"
'               'strimg = "21.35,1,4.4,150|6.5,10,3.65,150"
'            Else
'               strLine_Wman = "12.23,10,6.16,150|9.48,10,5.12,150"
'            End If
            If rsTmp!sex = "女" Then
               strLine_Wman = "11.98,mpas,5.78,mpas,8.73,mpas,1.58,mpas|9.35,mpas,4.85,mpas,7.03,mpas,1.37,mpas"
               'strimg = "21.35,1,4.4,150|6.5,10,3.65,150"
            Else
               strLine_Man = "12.23,mpas,6.16,mpas,9.08,mpas,1.63,mpas|9.48,mpas,5.12,mpas,7.25,mpas,1.41,mpas"
            End If
'            '分析图形(格式：图像名称;图像画法;座标数据;曲线及描点数据;坐标轴标题数据)
            
            Dim varTmp() As String
            
            If "" & rsTmp!sex = "男" Then
                '下限
                varTmp = Split(Split(strLine_Man, "|")(0), ",")
                arrLow(0) = varTmp(0): arrMid(0) = varTmp(4): arrHigh(0) = varTmp(2): arrNianDu(0) = varTmp(6)
                '检验值
                arrLow(1) = Val(rsTmp!lowShear): arrMid(1) = Mid(Val(rsTmp!middleShear10), 1, 4): arrHigh(1) = Val(rsTmp!highShear): arrNianDu(1) = Val(rsTmp!plasm)
                '上限
                varTmp = Split(Split(strLine_Man, "|")(1), ",")
                arrLow(2) = varTmp(0): arrMid(2) = varTmp(4): arrHigh(2) = varTmp(2): arrNianDu(2) = varTmp(6)
            Else
                '下限
                varTmp = Split(Split(strLine_Wman, "|")(0), ",")
                arrLow(0) = varTmp(0): arrMid(0) = varTmp(4): arrHigh(0) = varTmp(2): arrNianDu(0) = varTmp(6)
                '检验值
                arrLow(1) = Val(rsTmp!lowShear): arrMid(1) = Mid(Val(rsTmp!middleShear10), 1, 4): arrHigh(1) = Val(rsTmp!highShear): arrNianDu(1) = Val(rsTmp!plasm)
                '上限
                varTmp = Split(Split(strLine_Wman, "|")(1), ",")
                arrLow(2) = varTmp(0): arrMid(2) = varTmp(4): arrHigh(2) = varTmp(2): arrNianDu(2) = varTmp(6)
            End If
            
            strImg = "粘度特征曲线;110;0;" & frmGraph.Draw_mvis(arrHigh, arrMid, arrLow, arrNianDu, App.Path, rsTmp!sn)
            If strImg <> "" Then
                aResults(UBound(aResults)) = aResults(UBound(aResults)) & vbCrLf & strImg
            End If
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
End Function



