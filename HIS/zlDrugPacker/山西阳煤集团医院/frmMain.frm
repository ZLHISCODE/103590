VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HIS数据上传 v1.2"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6090
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6345
   ScaleWidth      =   6090
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdPara 
      Caption         =   "参数设置(&P)"
      Height          =   350
      Left            =   3000
      TabIndex        =   5
      Tag             =   "0"
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&E)"
      Height          =   350
      Left            =   4800
      TabIndex        =   4
      Top             =   5760
      Width           =   1100
   End
   Begin VB.Timer TimerTrans 
      Interval        =   60000
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "数据库设置(&D)"
      Height          =   350
      Left            =   1560
      TabIndex        =   3
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Frame fraH 
      Height          =   45
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Width           =   5800
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "开始上传(&S)"
      Height          =   350
      Left            =   120
      TabIndex        =   1
      Tag             =   "0"
      Top             =   5760
      Width           =   1335
   End
   Begin VB.ListBox lstLog 
      Height          =   5280
      ItemData        =   "frmMain.frx":030A
      Left            =   120
      List            =   "frmMain.frx":030C
      TabIndex        =   0
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOutConnect As Boolean   '外部数据库是否连接

Private mlng药房id As Long
Private mstr药房编码 As String
Private mlng轮询间隔 As Long
Private mint查询天数 As Integer
Private mstr剂型 As String
Private mstr开始时间 As String
Private mstr结束时间 As String
Private mstrUpdate As String        '要更新的数据：单据,NO|单据,NO

Private Function GetHisData() As Variant
    '获取HIS数据
    '从未发药品记录表取对应的NO
    '根据NO提取药品明细信息，并按频次分解成每一包的信息
    '但NO中药品有多个批次时，只处理一次
    
    Dim rsData As ADODB.Recordset
    Dim rsDataDrug As ADODB.Recordset
    Dim rsGetNext As ADODB.Recordset
    Dim n As Integer
    Dim strReturn As String
    Dim strLastTime As String
    Dim intCount As Integer
    Dim varReturn As Variant
    Dim str领药部门编码 As String
    Dim strPatiid As String
    Dim str分包设备编号 As String
    Dim str姓名     As String
    Dim strNO As String
    Dim str药房编码 As String
    Dim lng药品id As Long
    Dim int单据 As Integer
    Dim strDeptId As String     '慢病科室
    Dim intMBType As Integer    '慢病科室单据规则：0-不上传单据,1-只上传划价单,2-只上传收费单,3-所有单据都上传
    Dim intFMBType As Integer   '非慢病科室单据规则：0-不上传单据,1-只上传划价单,2-只上传收费单,3-所有单据都上传
    
    str分包设备编号 = "1"
    
    varReturn = Array()
    GetHisData = Array()
    
    On Error GoTo errHandle
    
    strDeptId = GetSetting("ZLSOFT", "公共模块\门诊药房包药机", "慢病科室", "")
    intMBType = Val(GetSetting("ZLSOFT", "公共模块\门诊药房包药机", "慢病科室单据性质", 0))
    intFMBType = Val(GetSetting("ZLSOFT", "公共模块\门诊药房包药机", "非慢病科室单据性质", 0))
    
    '如果参数值为不上传单据，则不执行查询
    If intMBType = 0 And intFMBType = 0 Then Exit Function
    
    '读取指定时间段的未上传处方信息
    gstrSql = "Select f.编码 as 领药部门编码, g.编码 药房编码, a.病人id,a.姓名,a.单据, a.No From 未发药品记录 A, 部门表 F, 部门表 G "
    
    '未关联药品收发记录表，注释掉
'    If mstr剂型 <> "" Then
'        gstrSql = gstrSql & " ,药品规格 C, 药品特性 D, Table(Cast(f_Str2list([4]) As zlTools.t_Strlist)) E "
'    End If

    gstrSql = gstrSql & " Where a.对方部门id = f.Id and a.库房id=g.id And a.单据 In (8, 9) And Nvl(a.是否上传, 0) = 0 And a.库房id = [1] And a.填制日期 Between [2] And [3] "

    '未关联药品收发记录表，注释掉
'    If mstr剂型 <> "" Then
'        gstrSql = gstrSql & " And b.药品id = c.药品id And c.药名id = d.药名id And d.药品剂型 = e.Column_Value"
'    End If
    
    '慢病科室单据上传规则
    If strDeptId = "" Then
        '如果没有选择慢病科室，则按所有科室都是非慢病科室规则处理
        If intFMBType = 1 Then
            '只划价单
            gstrSql = gstrSql & " And A.已收费 = 0 "
        ElseIf intFMBType = 2 Then
            '只收费单
            gstrSql = gstrSql & " And A.已收费 = 1 "
        ElseIf intFMBType = 3 Then
            '划价单，收费单都传，不加条件
        End If
    Else
        '选择了慢病科室， 按分别的规则处理
        If intMBType = 3 And intFMBType = 3 Then
            '如果所有单据都传则不加条件
        Else
            gstrSql = gstrSql & " And ("
            
            '慢病科室处理规则
            If intMBType = 1 Then
                '只划价单
                gstrSql = gstrSql & " (Instr([5], ',' || a.对方部门id || ',', 1) > 0 And 已收费 = 0) "
            ElseIf intMBType = 2 Then
                '只收费单
                gstrSql = gstrSql & " (Instr([5], ',' || a.对方部门id || ',', 1) > 0 And 已收费 = 1) "
            ElseIf intMBType = 3 Then
                '划价单，收费单都传
                gstrSql = gstrSql & " Instr([5], ',' || a.对方部门id || ',', 1) > 0  "
            End If
            
            gstrSql = gstrSql & " Or "
            
            '非慢病科室处理规则
            If intFMBType = 1 Then
                '只划价单
                gstrSql = gstrSql & " (Instr([5], ',' || a.对方部门id || ',', 1) = 0 And 已收费 = 0) "
            ElseIf intFMBType = 2 Then
                '只收费单
                gstrSql = gstrSql & " (Instr([5], ',' || a.对方部门id || ',', 1) = 0 And 已收费 = 1) "
            ElseIf intFMBType = 3 Then
                '划价单，收费单都传
                gstrSql = gstrSql & " Instr([5], ',' || a.对方部门id || ',', 1) = 0 "
            End If
             
            gstrSql = gstrSql & ")"
        End If
    End If
    
    gstrSql = gstrSql & " Order By f.编码, a.病人id, a.No "

    Set rsData = OpenSQLRecord(gstrSql, "HisTransData", mlng药房id, CDate(mstr开始时间), CDate(mstr结束时间), mstr剂型, strDeptId)
    
    If rsData.RecordCount = 0 Then Exit Function
    
    str领药部门编码 = rsData!领药部门编码
    strPatiid = rsData!病人id
    strNO = rsData!NO
    str药房编码 = rsData!药房编码
    
    Do While Not rsData.EOF
'        '按NO组织数据
'        If str领药部门编码 & "," & strPatiid & "," & strNO <> rsData!领药部门编码 & "," & rsData!病人id & "," & rsData!NO And strReturn <> "" Then
'            strReturn = str领药部门编码 & ";" & str姓名 & ";" & str分包设备编号 & ";" & strNO & "|" & strReturn
'
'            ReDim Preserve varReturn(UBound(varReturn) + 1)
'            varReturn(UBound(varReturn)) = strReturn
'
'            strReturn = ""
'        End If
                
        str领药部门编码 = rsData!领药部门编码
        strPatiid = rsData!病人id
        strNO = rsData!NO
        str药房编码 = rsData!药房编码
        int单据 = rsData!单据
        
        '初始数据
        lng药品id = 0
        strReturn = ""
        
        '读取药品信息，在指定NO的情况下必须要按药品ID排序
        gstrSql = " Select a.收发id, a.住院号, a.病人id,A.姓名, a.科室编码, a.科室名称, a.开单人, a.床号, a.用法, a.医生嘱托 ,a.药品编码, a.药品名称, a.规格, a.剂量系数, a.剂量单位,a.计算单位, a.服用数量,A.总给予量," & vbNewLine & _
        "                   a.首次时间, a.末次时间, a.开始执行时间, a.频率间隔, a.间隔单位, a.执行时间方案, Nvl(b.发送数次, 0) As 次数 ,a.批号,a.效期, a.发药数量, 整包装,a.执行频次,a.天数,a.药品id " & vbNewLine & _
        "            From (Select Distinct a.Id As 收发id, b.标识号 As 住院号, b.病人id, b.姓名, c.编码 As 科室编码, c.名称 As 科室名称, b.开单人, '' As 床号, a.用法,h.医生嘱托," & vbNewLine & _
        "                                  d.编码 As 药品编码, d.名称 As 药品名称, d.规格, e.剂量系数, d.计算单位,f.计算单位 As 剂量单位, h.单次用量/e.剂量系数  As 服用数量,round(H.总给予量,0) 总给予量,g.首次时间, g.末次时间," & vbNewLine & _
        "                                  h.开始执行时间, h.频率间隔, h.间隔单位, h.执行时间方案, h.相关id, g.发送号, a.实际数量 * Nvl(a.付数, 1) / e.门诊包装 As 发药数量," & vbNewLine & _
        "                                  Decode(Mod(a.实际数量 * Nvl(a.付数, 1), e.药库包装), 0, 1, 0) 整包装,h.执行频次,h.天数 ,a.批号,a.效期,a.药品id " & vbNewLine & _
        "                  From 药品收发记录 A, 门诊费用记录 B, 部门表 C, 收费项目目录 D, 药品规格 E, 诊疗项目目录 F, 病人医嘱发送 G, 病人医嘱记录 H" & vbNewLine & _
        "                  Where a.费用id = b.Id And b.开单部门id= c.Id And a.药品id = d.Id And b.记录状态 in (1,3) and a.药品id = e.药品id And e.药名id = f.Id And" & vbNewLine & _
        "                        b.医嘱序号 = g.医嘱id And b.No = g.No And b.医嘱序号 = h.Id And a.库房id = [1] And  a.单据 = [2] And a.No = [3]  ) A, 病人医嘱发送 B" & vbNewLine & _
        "             Where a.相关id = b.医嘱id(+) And a.发送号 = b.发送号(+) And a.用法='口服'" & vbNewLine & _
        "            Order By a.药品id "
        Set rsDataDrug = OpenSQLRecord(gstrSql, "HisTransData", mlng药房id, rsData!单据, rsData!NO)
        
        str姓名 = ""
        With rsDataDrug
            '循环整个单据的药品
            Do While Not .EOF
                
                If str姓名 = "" Then str姓名 = NVL(!姓名)
                
                If lng药品id <> !药品id Then
                    '对于药房分批药品，同一个NO可能有多个批次，只对一个批次的进行分解，如果药品ID相同则不处理
                    lng药品id = !药品id
                    
                    If Val(NVL(!频率间隔, 0)) = 0 Or NVL(!间隔单位, "") = "" Or NVL(!执行时间方案, "") = "" Then
                        intCount = 1
                    Else
                        intCount = Val(!次数)
                        If intCount = 0 Then
                            gstrSql = "Select Zl_Gettransexenumber([1],[2],[3],[4],[5],[6]) From Dual "
                            Set rsGetNext = OpenSQLRecord(gstrSql, "取下次执行时间", CDate(!开始执行时间), CDate(!首次时间), CDate(!末次时间), Val(!频率间隔), !间隔单位, !执行时间方案)
                            If Not rsGetNext.EOF Then
                                intCount = Val(rsGetNext.Fields(0).Value)
                            End If
                        End If
                        If intCount = 0 Then
                            intCount = 1
                        End If
                    End If
                    
                    For n = 1 To intCount
                        strReturn = IIf(strReturn = "", "", strReturn & "|")
                        strReturn = strReturn & !收发id
                        strReturn = strReturn & ";" & NVL(!住院号)
                        strReturn = strReturn & ";" & NVL(!病人id)
                        strReturn = strReturn & ";" & Replace(Replace(!姓名, ";", ""), "|", "")
                        strReturn = strReturn & ";" & !科室编码
                        strReturn = strReturn & ";" & Replace(Replace(!科室名称, ";", ""), "|", "")
                        strReturn = strReturn & ";" & Replace(Replace(!开单人, ";", ""), "|", "")
                        strReturn = strReturn & ";" & Replace(Replace(NVL(!床号, ""), ";", ""), "|", "")
                        strReturn = strReturn & ";" & Replace(Replace(NVL(!用法, ""), ";", ""), "|", "")
                        strReturn = strReturn & ";" & ""    '服用时间说明
                        strReturn = strReturn & ";" & !药品编码
                        strReturn = strReturn & ";" & Replace(Replace(!药品名称, ";", ""), "|", "")
                        strReturn = strReturn & ";" & Replace(Replace(!规格, ";", ""), "|", "")
                        strReturn = strReturn & ";" & NVL(!剂量系数, 1) * NVL(!服用数量, 1)
                        strReturn = strReturn & ";" & !剂量单位
                        strReturn = strReturn & ";" & !总给予量
                        
                        If n = 1 Then
                            strLastTime = Format(!首次时间, "YYYY-MM-DD HH:MM:SS")
                        Else
                            gstrSql = "Select Zl_Gettransexetime([1],[2],[3],[4],[5]) From Dual "
                            Set rsGetNext = OpenSQLRecord(gstrSql, "取下次执行时间", CDate(!开始执行时间), CDate(strLastTime), Val(!频率间隔), !间隔单位, !执行时间方案)
                            If Not rsGetNext.EOF Then
                                strLastTime = Format(rsGetNext.Fields(0).Value, "YYYY-MM-DD HH:MM:SS")
                            End If
                        End If
                        
                        strReturn = strReturn & ";" & strLastTime
                        strReturn = strReturn & ";" & "1"           '分包设备编号
                        strReturn = strReturn & ";" & "0"           '优先标记
                        strReturn = strReturn & ";" & "1"           '临嘱
                        strReturn = strReturn & ";" & Replace(Replace(!科室名称, ";", ""), "|", "")
                        strReturn = strReturn & ";" & !执行频次
                        strReturn = strReturn & ";" & Format(!首次时间, "YYYY-MM-DD HH:MM:SS")
                        strReturn = strReturn & ";" & Format(!天数, "0.0")
                        strReturn = strReturn & ";" & !执行时间方案
                        strReturn = strReturn & ";" & NVL(!批号)
                        strReturn = strReturn & ";" & NVL(!效期)
                        strReturn = strReturn & ";" & NVL(!服用数量, 1)
                    Next
                End If
                
                .MoveNext
            Loop
        End With
        
        If strReturn <> "" Then
            '按NO组织上传数据
            strReturn = str领药部门编码 & ";" & str姓名 & ";" & str分包设备编号 & ";" & strNO & ";" & int单据 & "|" & strReturn
            
            ReDim Preserve varReturn(UBound(varReturn) + 1)
            varReturn(UBound(varReturn)) = strReturn
            
'            '记录要更新的数据
'            If InStr(1, mstrUpdate, rsData!单据 & "," & rsData!NO) = 0 Then
'                mstrUpdate = IIf(mstrUpdate = "", "", mstrUpdate & "|") & rsData!单据 & "," & rsData!NO
'            End If
            
            Call OutputLog("" & Now & vbCrLf & strReturn)
        End If
                        
        rsData.MoveNext
        
'        If rsData.EOF And strReturn <> "" Then
'            '后面没有记录时，传递数据，并返回没有传递成功的收发ID
'            strReturn = str领药部门编码 & ";" & str姓名 & ";" & str分包设备编号 & ";" & strNO & "|" & strReturn
'
'            ReDim Preserve varReturn(UBound(varReturn) + 1)
'            varReturn(UBound(varReturn)) = strReturn
'
'        End If
    Loop
    
    GetHisData = varReturn
    
    Exit Function
    
errHandle:
'    If gobjComLib.ErrCenter() = 1 Then
'        Resume
'    End If
'    Call gobjComLib.SaveErrLog
    Call LogListItem(Err.Description)
    Call OutputLog("异常：" & Err.Description)
    Set varReturn = Nothing
End Function

Private Sub AutoTrans()
    Dim arrTrans As Variant
    Dim strReturn As String
    Dim i As Integer
    Dim int单据 As Integer
    Dim strNO As String
    Dim strTmp As String
    
    On Error GoTo errHandle
    
    '更新日期范围
    Call UpdateDateValue
    
    '获取HIS数据
    arrTrans = GetHisData()
       
    If UBound(arrTrans) = -1 Then
        LogListItem "本次无数据！" & Now
        Exit Sub
    End If
    
    mstrUpdate = ""
    
    Me.cmdStart.Enabled = False
    
    '分批上传数据
    For i = 0 To UBound(arrTrans)
        strReturn = TranToPacker(CStr(arrTrans(i)))
        If strReturn <> "" Then
            LogListItem "上传失败的收发ID：" & strReturn
        Else
            '记录提交成功的单据和单据号，后面需要更新上传标志
            strTmp = Left(arrTrans(i), InStr(arrTrans(i), "|") - 1)
            strNO = Split(strTmp, ";")(3)      'NO
            int单据 = Split(strTmp, ";")(4)     '单据
            If InStr(1, mstrUpdate, int单据 & "," & strNO) = 0 Then
                mstrUpdate = IIf(mstrUpdate = "", "", mstrUpdate & "|") & int单据 & "," & strNO
            End If
        End If
    Next
    
    '更新上传标志
    If mstrUpdate <> "" Then
         gstrSql = "Zl_未发药品记录_更新上传标志("
        '配药ID,打包
        gstrSql = gstrSql & mlng药房id
        '单据,NO
        gstrSql = gstrSql & ",'" & mstrUpdate & "'"
        gstrSql = gstrSql & ")"
        Call ExecuteProcedure(gstrSql, "更新上传标志")
    End If
    
    Me.cmdStart.Enabled = True
    
    LogListItem "本次上传数据完成！" & Now
    
    Exit Sub
errHandle:
    Me.cmdStart.Enabled = True
    LogListItem Err.Description
End Sub

Private Function TranToPacker(ByVal strData As String) As String
'功能： 传送药品自动分包数据
'参数： 分包数据字符串
'格式： 病区编码;库房组号;分包设备编号;NO;单据|收发ID1;病例号;...|收发ID2;病例号;...|收发ID3;病例号;...
'规则： 收发ID,病例号,病人ID,姓名,病区编码,病区名称,药师姓名,床号,服用方法,服药时间说明,
'       药品编码,药品名称,规格,剂量,剂量单位,服用数量,服用时间,分包设备编号,医嘱类型
'返回值：未成功传送的收发ID字符串
    Dim arrPrimary As Variant, arrSecondly As Variant, arrSecondlyVals As Variant
    Dim strInsert As String, strTmp As String, strID As String, strPageNO As String
    Dim i As Integer, j As Integer, intPageNO As Integer
    Dim rsInsert As New ADODB.Recordset
    Dim blnRollback As Boolean, blnInsert As Boolean, blnInserted As Boolean
    
    If gcnOutside Is Nothing Or gcnOutside.State = adStateClosed Then
        MsgBox "你未连接数据库，请先执行DBConnect()函数！", vbCritical, GSTR_MESSAGE
        TranToPacker = "NOT"
        Exit Function
    End If
    
    strTmp = Trim(strData)
    If strTmp = "" Then Exit Function
'     Exit Function
    arrPrimary = Split(Mid(strTmp, 1, InStr(1, strTmp, "|") - 1), ";")
    
    strTmp = Mid(strTmp, InStr(1, strTmp, "|") + 1)
    arrSecondly = Split(strTmp, "|")
    
''    取PageNO号
'    strTmp = "select convert(char(6),getdate(),12) + right('000000'+cast(isnull(max(substring(page_no,7,len(page_no))),0)+1 as varchar(4)),4) max_no " _
'           & "from dbo.atf_ypxx where convert(char(6),getdate(),12)=left(page_no,6)"
'    rsInsert.Open strTmp, gcnOutside
'    strPageNO = rsInsert!max_no
'    rsInsert.Close

    '取NO作为PageNO，山西阳煤门诊要求
    strPageNO = arrPrimary(3)
    
    '先传送表数据(从)
'    intPageNO = 1   '计数
'    intAbate = 0    '回滚数
    strInsert = "insert into dbo.atf_ypxx " _
              & "(DETAIL_SN,inpatient_no,p_id,name,ward_sn,ward_name,doctor,bed_no,comment,comm2,drug_code,drugname" _
              & ",specification,dosage,dos_unit,total,occ_time,atf_no,pri_flag,Mz_flag,dept_name,freq,start_times,days,script,lot,expiredate,amount,page_no) " & Chr(13)
    strTmp = ""
    For i = LBound(arrSecondly) To UBound(arrSecondly)
        '得到元素
        arrSecondlyVals = Split(arrSecondly(i), ";")
        '组织字符串
        strTmp = strTmp & "select "
        For j = LBound(arrSecondlyVals) To UBound(arrSecondlyVals)
            Select Case j
            Case 0
                strTmp = strTmp & "'" & arrSecondlyVals(j) & "'"
            Case 1 To 12, 14, 20, 16 To 18, 21 To 26
                strTmp = strTmp & ",'" & arrSecondlyVals(j) & "'"
            Case 13, 15, 19, 27
                strTmp = strTmp & "," & arrSecondlyVals(j)
            End Select
        Next
        strTmp = strTmp & ",'" & strPageNO & "'"
        strTmp = strTmp & " union all " & Chr(13)
        '判断下条记录是否为同一收发ID
        strID = arrSecondlyVals(0)
        If i = UBound(arrSecondly) Then
            blnInsert = True
        Else
            If Mid(arrSecondly(i + 1), 1, InStr(1, arrSecondly(i + 1), ";") - 1) = strID Then
                blnInsert = False
            Else
                blnInsert = True
            End If
        End If
        '是否执行Insert语句
        If blnInsert = True Then
            blnRollback = False
            strTmp = Left(strTmp, Len(strTmp) - 11)
            
            gcnOutside.BeginTrans
            On Error GoTo errRollback
            rsInsert.Open strInsert & strTmp, gcnOutside
            On Error GoTo 0
            If blnRollback = False Then
                gcnOutside.CommitTrans
                blnInserted = True
            Else
'                intPageNO = intPageNO - intAbate - 1
                '记录未提交的收发ID
                TranToPacker = TranToPacker & strID & ";"
            End If
            If rsInsert.State = adStateOpen Then rsInsert.Close
            strTmp = ""
'            intAbate = 0
        Else
            strTmp = strTmp & Chr(13)
            '记录多少条相同的
'            intAbate = intAbate + 1
        End If
'        intPageNO = intPageNO + 1
    Next
    If rsInsert.State = adStateOpen Then rsInsert.Close
    
    '先传送表数据(主)
    If blnInserted Then
        blnRollback = False
        strTmp = "insert into dbo.atf_yp_page_no (ward_sn,group_no,atf_no,submit_time,page_no,flag) " & Chr(13)
        strTmp = strTmp & "select "
        For i = LBound(arrPrimary) To UBound(arrPrimary)
            Select Case i
            Case 0 To 2
                strTmp = strTmp & "'" & arrPrimary(i) & "',"
'            Case 3
'                strTmp = strTmp & "getdate(),"
'            Case 4
'                strTmp = strTmp & "'" & strPageNO & "'"
            End Select
        Next
        strTmp = strTmp & "getdate(),'" & strPageNO & "',0"
        'strTmp = Left(strTmp, Len(strTmp) - 1)
        '提交数据
        gcnOutside.BeginTrans
        On Error GoTo errRollback
        rsInsert.Open strTmp, gcnOutside
        On Error GoTo 0
        If blnRollback = False Then
            gcnOutside.CommitTrans
        Else
            '如果主表数据失败，同事删除从表对应数据
            strTmp = "delete dbo.atf_ypxx where page_no='" & strPageNO & "'"
            On Error Resume Next
            If rsInsert.State = adStateOpen Then rsInsert.Close
            rsInsert.Open strTmp, gcnOutside
            If rsInsert.State = adStateOpen Then rsInsert.Close
            '返回所有收发ID字符串
            strID = "": TranToPacker = ""
            For i = LBound(arrSecondly) To UBound(arrSecondly)
                If Left(arrSecondly(i), InStr(1, arrSecondly(i), ";") - 1) <> strID Then
                    strID = Left(arrSecondly(i), InStr(1, arrSecondly(i), ";") - 1)
                    TranToPacker = TranToPacker & strID & ";"
                End If
            Next
        End If
    End If
    'If gcnOutside.State = adStateOpen Then gcnOutside.Close
    If Trim(TranToPacker) <> "" Then
        '返回收发ID字符串
        TranToPacker = Left(TranToPacker, Len(TranToPacker) - 1)
    End If
    
    Exit Function

errRollback:
    Call OutputLog("TranToPacker: " & Err.Description)
    gcnOutside.RollbackTrans
    blnRollback = True
    Resume Next
End Function


Private Sub cmdConnect_Click()
    frmOutsideLinkSet.Show
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPara_Click()
    frmPara.Show 1, Me
End Sub

Private Sub cmdStart_Click()
    If cmdStart.Tag = "0" Then
        cmdStart.Tag = "1"
        cmdStart.Caption = "停止上传(&S)"
        
        '开始上传
        TimerTrans.Enabled = True
        
        LogListItem "开始上传：" & Now
        
        cmdConnect.Enabled = False
        cmdPara.Enabled = False
    Else
        cmdStart.Tag = "0"
        cmdStart.Caption = "开始上传(&S)"
        
        '停止上传
        TimerTrans.Enabled = False
        
        LogListItem "停止上传" & Now
        
        cmdConnect.Enabled = True
        cmdPara.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    '初始化公共部件

'    Set gobjComLib = CreateObject("zl9ComLib.clsComLib")
'
'    If gobjComLib Is Nothing Then
'        MsgBox "初始化公共部件失败！", vbInformation, ""
'        Unload Me
'    End If
    
    '连接外部数据库
    mblnOutConnect = DBConnect
    
    '读取注册表参数
    mlng药房id = Val(GetSetting("ZLSOFT", "公共模块\门诊药房包药机", "药房ID"))
    mstr药房编码 = Val(GetSetting("ZLSOFT", "公共模块\门诊药房包药机", "药房编码"))
    mlng轮询间隔 = Val(GetSetting("ZLSOFT", "公共模块\门诊药房包药机", "轮询间隔", 60))
    mint查询天数 = Val(GetSetting("ZLSOFT", "公共模块\门诊药房包药机", "查询天数", 0))
    mstr剂型 = GetSetting("ZLSOFT", "公共模块\门诊药房包药机", "剂型", "")
    
    If mlng轮询间隔 > 60 Then
        mlng轮询间隔 = 60
    End If
    TimerTrans.Interval = mlng轮询间隔 * 1000
    
    '更新日期范围
    Call UpdateDateValue
    
End Sub

Private Sub UpdateDateValue()
    If mint查询天数 = 0 Then
        '默认是当天
        mstr开始时间 = Format(Currentdate, "YYYY-MM-DD")
        mstr结束时间 = Format(Currentdate, "YYYY-MM-DD 23:59:59")
    Else
        '指定天数内
        If mint查询天数 > 3 Then
            mint查询天数 = 3
        End If
        
        mstr开始时间 = Format(DateAdd("d", -mint查询天数, Currentdate), "YYYY-MM-DD")
        mstr结束时间 = Format(Currentdate, "YYYY-MM-DD 23:59:59")
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If cmdStart.Enabled = False Then Cancel = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Set gobjComLib = Nothing
End Sub

Private Sub TimerTrans_Timer()
    TimerTrans.Enabled = False
    
    On Error GoTo errHandle
    
    '检查连接
    If gcnOracle.State <> adStateOpen Then
        gcnOracle.Open
    End If
    If gcnOutside.State <> adStateOpen Then
        gcnOutside.Open
    End If
    
    DoEvents
    '调用自动上传程序
    Call AutoTrans
    DoEvents
    TimerTrans.Enabled = True
    
    Exit Sub
    
errHandle:
    Call LogListItem("异常：" & Err.Description)
    TimerTrans.Enabled = True
End Sub

Private Sub LogListItem(ByVal strLog As String)
    Const INT_MAX_LINES As Integer = 200

    Me.lstLog.AddItem strLog
    Me.lstLog.Selected(Me.lstLog.ListCount - 1) = True
    Me.lstLog.TopIndex = Me.lstLog.ListCount - 1
    If lstLog.ListCount >= INT_MAX_LINES Then lstLog.RemoveItem 0

End Sub
