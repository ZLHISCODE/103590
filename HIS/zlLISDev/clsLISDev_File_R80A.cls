VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLISDev_File_R80A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'仪器:血流变
'厂家：北京中勤世帝科学仪器公司
'医院：黑龙江省农垦总医院,南京红十字医院,河南南石医院
'2009-03-08 ZHQ
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
    Dim strimg As String
    
    
    strFilePath = objFile.GetParentFolderName(strFile)
    If Len(strFilePath) = 0 Then strFilePath = App.Path
    With cnAccess
        .ConnectionString = "DBQ=" & strFile & ";DefaultDir=" & strFilePath & ";Driver={Microsoft Access Driver (*.mdb)}"
        .Open
        If .State = adStateClosed Then Exit Function
        
        strSQL = " Select * From Checks Where CheckDate between cdate('" & Format(dtStart, "yyyy-mm-dd") & "') and cdate('" & Format(dtEnd, "yyyy-mm-dd") & "')"
        
        If lngSampleNo <> -1 Then
           strSQL = strSQL & " and val(CheckNum) = " & lngSampleNo
        End If
        strSQL = strSQL & " Order By CheckDate,CheckNum"
        
        Set rsTmp = cnAccess.Execute(strSQL)
        If rsTmp.EOF Then
            .Close
            Exit Function
        End If
        '---  增加河南南石医院兼容性处理，南石医院无“LowAlbumen”字段。
        '---  这里记录下所有字段，用于后面判断,有此字段，才执行相关语句，无此字段，则不执行。
        strItems = ""
        For i = 0 To rsTmp.Fields.count - 1
            strItems = strItems & "," & rsTmp.Fields(i).Name
        Next
        
        lngResultCount = -1
        Do While Not rsTmp.EOF
            If InStr(";123456789;", Mid(rsTmp!checknum, 1, 1)) <= 0 Or Not IsNumeric(rsTmp!checknum) Then
                '首位非具体数字就默认为质控标本
                lngBh = 999
            Else
                lngBh = rsTmp!checknum
            End If
            lngResultCount = lngResultCount + 1
            ReDim Preserve aResults(lngResultCount)
            aResults(lngResultCount) = Format(rsTmp!CheckDate, "yyyy-MM-DD") & "|" & lngBh & "| " & "|血液|" & IIf(lngBh = 999, 1, 0)
            '下面开始逐项加入项目和值
            If Nvl(rsTmp!BloodValueHighHigh, 0) <> 0 Then
                '全血粘度值(mPa/s): 通常是200
                aResults(lngResultCount) = aResults(lngResultCount) & "|1|" & rsTmp!BloodValueHighHigh
            End If
            If Nvl(rsTmp!BloodValueHigh, 0) <> 0 Then
                '全血粘度值(mPa/s): 通常是100
                aResults(lngResultCount) = aResults(lngResultCount) & "|2|" & rsTmp!BloodValueHigh
            End If
            If Nvl(rsTmp!BloodValueMiddle, 0) <> 0 Then
                '全血粘度值(mPa/s): 通常是5
                aResults(lngResultCount) = aResults(lngResultCount) & "|3|" & rsTmp!BloodValueMiddle
            End If
            If Nvl(rsTmp!BloodValueLow, 0) <> 0 Then
                '全血粘度值(mPa/s): 通常是1
                aResults(lngResultCount) = aResults(lngResultCount) & "|4|" & rsTmp!BloodValueLow
            End If
            If Nvl(rsTmp!PlasmaValue, 0) <> 0 Then
                '血浆粘度值(mPa/s): 通常是100
                aResults(lngResultCount) = aResults(lngResultCount) & "|5|" & rsTmp!PlasmaValue
            End If
            If Nvl(rsTmp!ESR, 0) <> 0 Then
                'ESR血沉
                aResults(lngResultCount) = aResults(lngResultCount) & "|6|" & rsTmp!ESR
            End If
            If Nvl(rsTmp!Press, 0) <> 0 Then
                '红细胞压积
                aResults(lngResultCount) = aResults(lngResultCount) & "|7|" & rsTmp!Press
            End If
            If Nvl(rsTmp!HighShearVisc, 0) <> 0 Then
                '全血高切还原粘度
                aResults(lngResultCount) = aResults(lngResultCount) & "|8|" & rsTmp!HighShearVisc
            End If
            If Nvl(rsTmp!LowShearVisc, 0) <> 0 Then
                '全血低切还原粘度
                aResults(lngResultCount) = aResults(lngResultCount) & "|9|" & rsTmp!LowShearVisc
            End If
            If Nvl(rsTmp!Rigidity, 0) <> 0 Then
                '红细胞刚性指数
                aResults(lngResultCount) = aResults(lngResultCount) & "|10|" & rsTmp!Rigidity
            End If
            If Nvl(rsTmp!Assemble, 0) <> 0 Then
                '红细胞聚集指数
                aResults(lngResultCount) = aResults(lngResultCount) & "|11|" & rsTmp!Assemble
            End If
            If Nvl(rsTmp!EquationK, 0) <> 0 Then
                '血沉方程K值
                aResults(lngResultCount) = aResults(lngResultCount) & "|12|" & rsTmp!EquationK
            End If
            If Nvl(rsTmp!HighShearOppositeVisc, 0) <> 0 Then
                '全血高切相对粘度
                aResults(lngResultCount) = aResults(lngResultCount) & "|13|" & rsTmp!HighShearOppositeVisc
            End If
            If Nvl(rsTmp!LowShearOppositeVisc, 0) <> 0 Then
                '全血低切相对粘度
                aResults(lngResultCount) = aResults(lngResultCount) & "|14|" & rsTmp!LowShearOppositeVisc
            End If
            If Nvl(rsTmp!TransformTK, 0) <> 0 Then
                '红细胞变形指数TK
                aResults(lngResultCount) = aResults(lngResultCount) & "|15|" & rsTmp!TransformTK
            End If
            If Nvl(rsTmp!Proteid, 0) <> 0 Then
                '纤维蛋白原
                aResults(lngResultCount) = aResults(lngResultCount) & "|16|" & rsTmp!Proteid
            End If
            If Nvl(rsTmp!Cholesterin, 0) <> 0 Then
                '血胆固醇
                aResults(lngResultCount) = aResults(lngResultCount) & "|17|" & rsTmp!Cholesterin
            End If
            If Nvl(rsTmp!Glyceride, 0) <> 0 Then
                '甘油三脂
                aResults(lngResultCount) = aResults(lngResultCount) & "|18|" & rsTmp!Glyceride
            End If
            If Nvl(rsTmp!Albumen, 0) <> 0 Then
                '高密脂蛋白
                aResults(lngResultCount) = aResults(lngResultCount) & "|19|" & rsTmp!Albumen
            End If
            If Nvl(rsTmp!Bloodsugar, 0) <> 0 Then
                '血糖
                aResults(lngResultCount) = aResults(lngResultCount) & "|20|" & rsTmp!Bloodsugar
            End If
            If Nvl(rsTmp!ThrombusLen, 0) <> 0 Then
                '体外血栓长度
                aResults(lngResultCount) = aResults(lngResultCount) & "|21|" & rsTmp!ThrombusLen
            End If
            If Nvl(rsTmp!ThrombusdampWeight, 0) <> 0 Then
                '体外血栓湿重
                aResults(lngResultCount) = aResults(lngResultCount) & "|22|" & rsTmp!ThrombusdampWeight
            End If
            If Nvl(rsTmp!ThrombusdryWeight, 0) <> 0 Then
                '体外血栓干重
                aResults(lngResultCount) = aResults(lngResultCount) & "|23|" & rsTmp!ThrombusdryWeight
            End If
            If Nvl(rsTmp!BeforePlateletNum, 0) <> 0 Then
                '粘附前血小板数
                aResults(lngResultCount) = aResults(lngResultCount) & "|24|" & rsTmp!BeforePlateletNum
            End If
            If Nvl(rsTmp!AfterPlateletNum, 0) <> 0 Then
                '粘附后血小板数
                aResults(lngResultCount) = aResults(lngResultCount) & "|25|" & rsTmp!AfterPlateletNum
            End If
            If Nvl(rsTmp!PlateletAdhere, 0) <> 0 Then
                '血小板粘附率
                aResults(lngResultCount) = aResults(lngResultCount) & "|26|" & rsTmp!PlateletAdhere
            End If
            If Nvl(rsTmp!Electrophoresis, 0) <> 0 Then
                '红细胞电泳
                aResults(lngResultCount) = aResults(lngResultCount) & "|27|" & rsTmp!Electrophoresis
            End If
            If Nvl(rsTmp!PlateletAssemble, 0) <> 0 Then
                '血小板聚集率
                aResults(lngResultCount) = aResults(lngResultCount) & "|28|" & rsTmp!PlateletAssemble
            End If
            If InStr(strItems & ",", ",LowAlbumen" & ",") > 0 Then
                If Nvl(rsTmp!LowAlbumen, 0) <> 0 Then
                    '低密脂蛋白
                    aResults(lngResultCount) = aResults(lngResultCount) & "|29|" & rsTmp!LowAlbumen
                End If
            End If
            '检测结果（临床参考）   Result(由于HIS无法回传，暂时取消)
            If Nvl(rsTmp!Result, 0) <> 0 Then
                aResults(lngResultCount) = aResults(lngResultCount) & "|30|" & rsTmp!Result
            End If
           
            '分析图形(格式：图像名称;图像画法;座标数据;曲线及描点数据;坐标轴标题数据)
            '根据性别判断改变划线坐标值
            If rsTmp!sex = False Then
               strimg = "9.25,10,4.4,150|6.5,10,3.65,150"
               'strimg = "21.35,1,4.4,150|6.5,10,3.65,150"
            Else
               strimg = "9.58,10,4.6,150|6.8,10,3.73,150"
            End If

'            '分析图形(格式：图像名称;图像画法;座标数据;曲线及描点数据;坐标轴标题数据)
'            aResults(lngResultCount) = aResults(lngResultCount) & vbCrLf & "粘度特征曲线;2;" & _
'                "30,200|3-3,10-10,30-30,100-100,200-200|5-5,10-10,15-15,20-20,25-25,30-30;" & _
'                Nvl(rsTmp!BloodValueHighHigh, 0) & "," & Nvl(rsTmp!BloodValueHigh, 0) & "," & Nvl(rsTmp!BloodValueMiddle, 0) & "," & Nvl(rsTmp!BloodValueLow, 0) & ";" & _
'                "粘度值(mPa.s),1,35~切变率(1/s),210,1"
            
            aResults(lngResultCount) = aResults(lngResultCount) & vbCrLf & "粘度特征曲线;2;" & _
                "20,230|3-3,10-10,30-30,100-100,200-200|5-5,10-10,15-15,20-20,25-25,30-30;" & _
                strimg & _
                "~5-" & Nvl(rsTmp!BloodValueLow, 0) & ",30-" & Nvl(rsTmp!BloodValueMiddle, 0) & ",200-" & Nvl(rsTmp!BloodValueHighHigh, 0) & _
                ";粘度值(mPa.s),35,20~切变率(1/s),210,1"
  
            
            rsTmp.MoveNext
        Loop
    End With
    ResultFromFile = aResults
End Function
