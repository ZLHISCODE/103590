VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "护理数据升迁"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6780
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form21"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Caption         =   "运行时间设定(系统升级不受限)"
      ForeColor       =   &H8000000D&
      Height          =   1665
      Index           =   1
      Left            =   3690
      TabIndex        =   6
      Top             =   240
      Width           =   2955
      Begin MSComCtl2.DTPicker dtp开始时间 
         Height          =   315
         Left            =   1920
         TabIndex        =   8
         Top             =   300
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:mm"
         Format          =   101646339
         CurrentDate     =   40540
      End
      Begin MSComCtl2.DTPicker dtp开始时间1 
         Height          =   315
         Left            =   1920
         TabIndex        =   10
         Top             =   690
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:mm"
         Format          =   101384195
         CurrentDate     =   40540.0833333333
      End
      Begin MSComCtl2.DTPicker dtp结束时间 
         Height          =   315
         Left            =   1920
         TabIndex        =   12
         Top             =   1080
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "HH:mm"
         Format          =   101384195
         CurrentDate     =   40540.1666666667
      End
      Begin VB.Label lbl开始时间 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "数据升迁开始时间"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   315
         TabIndex        =   7
         Top             =   360
         Width           =   1440
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "打印解析开始时间"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   315
         TabIndex        =   9
         Top             =   750
         Width           =   1440
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "结束处理时间"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   675
         TabIndex        =   11
         Top             =   1140
         Width           =   1080
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "导入模式"
      Height          =   1245
      Index           =   0
      Left            =   150
      TabIndex        =   2
      Top             =   660
      Width           =   3375
      Begin VB.OptionButton opt 
         Caption         =   "升迁历史数据(后台)"
         Height          =   180
         Index           =   2
         Left            =   690
         TabIndex        =   5
         Top             =   900
         Width           =   2295
      End
      Begin VB.OptionButton opt 
         Caption         =   "系统升级"
         Height          =   180
         Index           =   1
         Left            =   690
         TabIndex        =   4
         Top             =   600
         Width           =   1845
      End
      Begin VB.OptionButton opt 
         Caption         =   "升级前准备"
         Height          =   180
         Index           =   0
         Left            =   690
         TabIndex        =   3
         Top             =   300
         Width           =   1845
      End
   End
   Begin VB.Timer tim 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   150
      Top             =   2940
   End
   Begin VB.CommandButton Command1 
      Caption         =   "后台运行"
      Height          =   350
      Left            =   5370
      TabIndex        =   14
      Top             =   2070
      Width           =   1100
   End
   Begin VB.ComboBox cbo病区 
      Height          =   300
      Left            =   750
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   2685
   End
   Begin VB.Label Label2 
      Caption         =   "请选择一种导入模式"
      ForeColor       =   &H000000FF&
      Height          =   525
      Left            =   240
      TabIndex        =   13
      Top             =   2040
      Width           =   3285
   End
   Begin VB.Label lbl病区 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "病区"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   300
      TabIndex        =   0
      Top             =   300
      Width           =   360
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mlng病人ID As Long
Dim strSQL As String
Dim rsTemp As New ADODB.Recordset
Dim rsPati As New ADODB.Recordset
Dim rsFile As New ADODB.Recordset
Dim rsDept As New ADODB.Recordset

'工作流程说明:
'升级前准备或升级后的历史数据升迁,程序只要启动后就会于每晚12点至凌晨6点之间进行数据升迁及打印解析工作,主要分两步走,第一步完成所有在院病人(新入院病人)历史数据的导入,第二步完成历史病人的数据导入工作


'--提取65txyy数据库上相关数据进行核对
'--64757,1|661573,1|386712,1|470068,1  这四个病人的护理数据分别是>300,>400,>900,>2000
'--select 病人ID,住院次数 from 病人信息 where 住院号=124631;
'select * from 病人信息 where 病人ID=661573;
'select * from 病案主页 where 病人ID=661573 and 主页ID=1;
'select count(*) from 病人护理记录 where 病人ID=661573 and 主页ID=1;
'select count(*) from 病人护理记录 A,病人护理内容 B Where A.病人ID=661573 and A.主页ID=1 And A.ID=B.记录ID ;
'select count(*) from 病人护理文件 where 病人ID=661573 and 主页ID=1;
'select A.文件ID,count(*) from 病人护理数据 A,病人护理文件 B where A.文件ID=B.ID And B.病人ID=661573 and B.主页ID=1 group by A.文件ID ;
'select A.ID AS 文件ID,count(*) from 病人护理文件 A,病人护理数据 B,病人护理明细 C Where A.病人ID=661573 and A.主页ID=1 And A.ID=B.文件ID And B.ID=C.记录ID Group by A.ID ;
'select * from 病人护理文件 where 病人ID=661573 and 主页ID=1;
'select count(*) from 病人护理记录 B,病人护理内容 C where B.ID=C.记录ID And B.病人ID=661573 And B.主页ID=1 And B.婴儿=0 and 项目分组 IN ('1)体温曲线项目','2)体温表格项目');



Private Sub DataUpgrade()
    Dim objFile As New FileSystemObject
    Dim objStream As TextStream
    Dim i As Integer, j As Integer, intStart As Integer, intEnd As Integer
    Dim datStart As Date, datEnd As Date
    Dim lng文件ID As Long
    Dim lng病人ID As Long, lng主页ID As Long, int婴儿 As Integer, lng格式ID As Long, lng保留 As Long
    Dim str归档人 As String, str归档时间 As String, strMsg As String
    Dim bln体温 As Boolean      '插入体温单数据后置为真
    Dim bln汇总 As Boolean      '是否插入汇总数据(病人变化或文件格式变化时)
    Dim blnFirst As Boolean     '第一步运行
    Dim blnError As Boolean     '发生错误
    Dim blnCommit As Boolean    '切换病区时需要提交事务
    Dim blnDataMoved As Boolean '数据转出标志
    Dim rsDate As New ADODB.Recordset
    Dim str时间 As String
    On Error GoTo errHand
    '文件格式相同的不重复产生数据,所有数据的开始日期为入院日期,结束日期为出院日期
    '待病人所有文件清单产生以后,再产生汇总数据,把与bln汇总相关的代码恢复即可,现在的模式为:如果找不到结束时间对应的护理文件,就以该格式文件最后一个科室为准产生
    Set objStream = objFile.OpenTextFile("C:\" & IIf(gintAutoRUN = 1, "AUTO", "") & "Data_LOG" & Format(Now, "yyyyMMddHHmmss") & ".txt", ForWriting, True)
    str时间 = "sysdate - 30"
    
    Command1.Enabled = False
    blnFirst = True
    gcnOracle.BeginTrans
    
    If Me.cbo病区.ListIndex = 0 Then
        intStart = 1
        intEnd = Me.cbo病区.ListCount - 1   '多加了"所有病区",因此从1开始循环,循环数次减去增加的所有病区
    Else
        intEnd = Me.cbo病区.ListIndex
        intStart = intEnd
    End If

redo:
    For i = intStart To intEnd
        Call WriteLog(objStream, String(50, "-"))
        Select Case gintMode
        Case 0  '升级前准备（当前在院病人及最近30天出院病人的历史住院的护理数据）
            strSQL = "" & _
                    "SELECT  /*+ RULE */ 病人ID,主页ID,婴儿 " & vbNewLine & _
                    "FROM (" & vbNewLine & _
                    "    SELECT  B.病人ID,B.主页ID,0 AS 婴儿" & vbNewLine & _
                    "    FROM 病人信息 C,病案主页 B," & vbNewLine & _
                    "        (SELECT A.病人ID" & vbNewLine & _
                    "        FROM 病人信息 A" & vbNewLine & _
                    "        WHERE A.在院=1 And A.当前病区ID=[1]" & vbNewLine & _
                    "        UNION" & vbNewLine & _
                    "        SELECT DISTINCT A.病人ID" & vbNewLine & _
                    "        FROM 病案主页 A" & vbNewLine & _
                    "        WHERE A.当前病区ID=[1] And A.出院日期>=" & str时间 & ") A" & vbNewLine & _
                    "    WHERE B.病人ID=C.病人ID AND B.主页ID<>C.住院次数 AND C.病人ID=A.病人ID" & vbNewLine
            strSQL = strSQL & _
                    "    UNION" & vbNewLine & _
                    "    SELECT B.病人ID,B.主页ID,B.序号 AS 婴儿" & vbNewLine & _
                    "    FROM 病人信息 C,病人新生儿记录 B," & vbNewLine & _
                    "        (SELECT A.病人ID" & vbNewLine & _
                    "        FROM 病人信息 A" & vbNewLine & _
                    "        WHERE A.在院=1 And A.当前病区ID=[1]" & vbNewLine & _
                    "        UNION" & vbNewLine & _
                    "        SELECT DISTINCT A.病人ID" & vbNewLine & _
                    "        FROM 病案主页 A" & vbNewLine & _
                    "        WHERE A.当前病区ID=[1] And A.出院日期>=" & str时间 & ") A" & vbNewLine & _
                    "    WHERE C.病人ID=B.病人ID AND C.住院次数<>B.主页ID AND C.病人ID=A.病人ID" & vbNewLine & _
                    "    MINUS" & vbNewLine & _
                    "    SELECT 病人ID,主页ID,婴儿 From 护理升迁记录" & vbNewLine & _
                    "    ) " & vbNewLine & _
                    "ORDER BY 病人ID,主页ID DESC ,婴儿"
        Case 1  '升级（升级当天仍在院病人及最近30天出院病人的所有护理数据）
            strSQL = "" & _
                    "SELECT  /*+ RULE */ 病人ID,主页ID,婴儿 " & vbNewLine & _
                    "FROM ( " & vbNewLine & _
                    "    SELECT  B.病人ID,B.主页ID,0 AS 婴儿" & vbNewLine & _
                    "    FROM 病案主页 B," & vbNewLine & _
                    "        (SELECT A.病人ID" & vbNewLine & _
                    "        FROM 病人信息 A" & vbNewLine & _
                    "        WHERE A.在院=1 And A.当前病区ID=[1]" & vbNewLine & _
                    "        UNION" & vbNewLine & _
                    "        SELECT DISTINCT A.病人ID" & vbNewLine & _
                    "        FROM 病案主页 A" & vbNewLine & _
                    "        WHERE A.当前病区ID=[1] ANd A.出院日期>=" & str时间 & ") A" & vbNewLine & _
                    "    WHERE B.病人ID=A.病人ID" & vbNewLine
            strSQL = strSQL & _
                    "    UNION" & vbNewLine & _
                    "    SELECT B.病人ID,B.主页ID,B.序号 AS 婴儿" & vbNewLine & _
                    "    FROM 病人新生儿记录 B," & vbNewLine & _
                    "        (SELECT A.病人ID" & vbNewLine & _
                    "        FROM 病人信息 A" & vbNewLine & _
                    "        WHERE A.在院=1 And A.当前病区ID=[1]" & vbNewLine & _
                    "        UNION" & vbNewLine & _
                    "        SELECT DISTINCT A.病人ID" & vbNewLine & _
                    "        FROM 病案主页 A" & vbNewLine & _
                    "        WHERE A.当前病区ID=[1] And A.出院日期>=" & str时间 & ") A" & vbNewLine & _
                    "    WHERE B.病人ID=A.病人ID" & vbNewLine & _
                    "    MINUS" & vbNewLine & _
                    "    SELECT 病人ID,主页ID,婴儿 From 护理升迁记录" & vbNewLine & _
                    "    ) " & vbNewLine & _
                    "ORDER BY 病人ID,主页ID DESC ,婴儿"
        Case 2  '后台转历史数据
            strSQL = "" & _
                    "SELECT /*+ RULE */  病人ID,主页ID,婴儿 " & vbNewLine & _
                    "FROM (" & _
                    "    SELECT A.病人ID,A.主页ID,0 AS 婴儿" & vbNewLine & _
                    "    FROM 病案主页 A" & vbNewLine & _
                    "    WHERE A.当前病区ID=[1]" & vbNewLine & _
                    "    UNION" & vbNewLine & _
                    "    SELECT A.病人ID,A.主页ID,A.序号 AS 婴儿" & vbNewLine & _
                    "    FROM 病人新生儿记录 A,病案主页 B" & vbNewLine & _
                    "    WHERE A.病人ID=B.病人ID AND A.主页ID=B.主页ID AND B.当前病区ID=[1]) " & vbNewLine & _
                    "    MINUS" & vbNewLine & _
                    "    SELECT 病人ID,主页ID,婴儿 From 护理升迁记录" & vbNewLine & _
                    "ORDER BY 病人ID,主页ID desc ,婴儿"
        End Select
        Set rsPati = OpenSQLRecord(strSQL, "提取指定病区所有病人", CLng(Me.cbo病区.ItemData(i)))
        Call WriteLog(objStream, "提取指定病区所有病人清单...完成,病区:" & CLng(Me.cbo病区.ItemData(i)) & ",住院人次:" & rsPati.RecordCount)
        
        With rsPati
            Do While Not .EOF
                '当病人发生变化时就产生最后一个文件的汇总数据
                If lng病人ID <> 0 Then   '只要有数据就有体温单与护理记录单
                    If Not blnError Then
                        'Call WriteLog(objStream, "病人ID=" & lng病人ID & ";主页ID=" & lng主页ID & ";婴儿=" & int婴儿 & ";文件格式ID=" & lng格式ID & "准备产生汇总数据")
                        If Not InsertCollect(lng格式ID, lng病人ID, lng主页ID, int婴儿, datStart, datEnd, lng文件ID) Then
                            blnError = True
                            strMsg = "正在处理[" & Me.cbo病区.List(i) & "]病人ID=" & lng病人ID & ";主页ID=" & lng主页ID & ";婴儿=" & int婴儿 & ";文件格式ID:" & lng格式ID & "  的汇总数据时发生错误"
                            objStream.WriteLine "发生错误,跳过该病人:" & vbCrLf & strMsg
                        Else
                            'Call WriteLog(objStream, "产生汇总数据...完成")
                        End If
                    End If
                    '插入升迁记录
                    gcnOracle.Execute "zl_护理升迁记录_Insert(" & lng病人ID & "," & lng主页ID & "," & int婴儿 & "," & IIf(blnError, "1", "0") & ",'" & Replace(Replace(Replace(Replace(strMsg, "'", ""), vbLf, ""), "[", ""), "]", "") & "')", , adCmdStoredProc
                End If
EXITDO:
                If blnCommit Or blnError Then
                    If blnError = False Then
                        gcnOracle.CommitTrans
                        blnError = False
                        blnCommit = False
                        gcnOracle.BeginTrans
                    Else
                        gcnOracle.RollbackTrans
                        blnError = False
                        blnCommit = False
                        gcnOracle.BeginTrans
                    End If
                End If
                
                strMsg = ""
                '新病人进行重新赋值
                bln体温 = False
                lng病人ID = !病人ID
                lng主页ID = !主页ID
                int婴儿 = !婴儿
                lng格式ID = 0
                Me.Caption = "[" & Me.cbo病区.List(i) & "] 进度:" & rsPati.AbsolutePosition & "/" & rsPati.RecordCount
                
                '检查是否归档
                strSQL = " Select   归档人,归档时间 From 病人护理记录 Where 病人ID=[1] And 主页ID=[2]"
                Set rsTemp = OpenSQLRecord(strSQL, "检查是否归档", lng病人ID, lng主页ID)
                If rsTemp.RecordCount <> 0 Then
                    str归档人 = NVL(rsTemp!归档人)
                    str归档时间 = Format(rsTemp!归档时间, "yyyy-MM-dd HH:mm:ss")
                End If
                
                '提取病人入出院时间
                gstrSQL = " Select   入院日期,nvl(出院日期,sysDate) as 出院日期,NVL(数据转出,0) AS 转出 From 病案主页 " & _
                          " Where 病人ID=[1] ANd 主页ID=[2] And [3]=0 " & _
                          " UNION " & _
                          " Select A.出生时间 AS 入院时间,nvl(B.出院日期,sysDate) as 出院日期,NVL(B.数据转出,0) AS 转出 From 病人新生儿记录 A,病案主页 B" & _
                          " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And A.病人ID=[1] And A.主页ID=[2] And A.序号=[3]"
                Set rsDate = OpenSQLRecord(gstrSQL, "提取病人入出院时间", lng病人ID, lng主页ID, int婴儿)
                blnDataMoved = rsDate!转出
                
                '按原规则提取病人的护理文件列表
                'Call WriteLog(objStream, "按原规则提取病人的护理文件列表")
                strSQL = "" & _
                        " SELECT   DISTINCT A.ID, A.编号, A.名称 AS 文件," & vbNewLine & _
                        "        A.开始,A.截止," & vbNewLine & _
                        "        A.科室ID, B.名称 AS 科室, 0 AS 护理级别,A.文件级别,保留" & vbNewLine & _
                        " FROM (" & vbNewLine & _
                        "        SELECT F.ID, F.编号, F.名称, R.开始, R.截止, R.科室ID, 保留,文件级别" & vbNewLine & _
                        "        FROM (" & vbNewLine & _
                        "        SELECT ID, 编号, 名称, 3 AS 文件级别, 通用, 0 AS 科室ID,保留 FROM 病历文件列表 WHERE 种类=3 AND 保留<0" & vbNewLine & _
                        "        UNION ALL" & vbNewLine & _
                        "        SELECT L.ID, L.编号, L.名称, F.报表 AS 文件级别, L.通用, A.科室ID,L.保留" & vbNewLine & _
                        "        FROM 病历文件列表 L, 病历页面格式 F, 病历应用科室 A" & vbNewLine & _
                        "        WHERE L.种类 = 3 AND L.保留 = 0 AND L.种类 = F.种类 AND L.编号 = F.编号 AND L.ID = A.文件ID(+)) F," & vbNewLine & _
                        "      (SELECT R.科室ID, NVL(MIN(R.护理级别),3) AS 护理级别, MIN(R.发生时间) AS 开始, MAX(R.发生时间) AS 截止" & vbNewLine & _
                        "      FROM 病人护理记录 R" & vbNewLine & _
                        "      WHERE R.病人来源 = 2 AND R.病人ID = [1] AND NVL(R.主页ID, 0) = [2] AND NVL(R.婴儿,0)=[3]" & vbNewLine & _
                        "      GROUP BY R.科室ID) R" & vbNewLine & _
                        "        WHERE (F.保留<0 OR F.通用 = 1 OR F.通用 = 2 AND R.科室ID IN (SELECT T.科室ID FROM 病区科室对应 T WHERE T.病区ID=F.科室ID)) AND F.文件级别 >= R.护理级别) A, 部门表 B" & vbNewLine & _
                        " WHERE A.科室ID = B.ID" & vbNewLine & _
                        " ORDER BY A.保留,A.文件级别,A.编号 DESC, TO_CHAR(A.开始, 'YYYY-MM-DD HH24:MI') || ' ～ ' || TO_CHAR(A.截止, 'YYYY-MM-DD HH24:MI')"
                If blnDataMoved Then
                    strSQL = Replace(strSQL, "病人护理记录", "H病人护理记录")
                    strSQL = Replace(strSQL, "病人护理内容", "H病人护理内容")
                End If
                Set rsFile = OpenSQLRecord(strSQL, "按原规则提取病人的护理文件列表", lng病人ID, lng主页ID, int婴儿)
                'Call WriteLog(objStream, "按原规则提取病人的护理文件列表...完成,记录数:" & rsFile.RecordCount)
                
'                bln汇总 = False
                '先产生病人所有护理文件(含体温单)的数据,然后再重新循环产生该病人的汇总数据
                Do While Not rsFile.EOF
                    
                    datStart = rsDate!入院日期
                    datEnd = rsDate!出院日期
                    
                    '当格式发生变化时就产生汇总数据
                    '体温单不产生汇总数据
'                    If bln汇总 Then
'                        If rsFile!保留 <> -1 And lng保留 <> -1 And lng格式ID <> rsFile!ID And lng格式ID <> 0 Then
'                            'Call WriteLog(objStream, "病人ID=" & lng病人ID & ";主页ID=" & lng主页ID & ";婴儿=" & int婴儿 & ";文件格式ID=" & lng格式ID & "准备产生汇总数据")
'                            If Not InsertCollect(lng格式ID, lng病人ID, lng主页ID, int婴儿, datStart, datEnd, lng文件ID) Then
'                                blnError = True
'                                strMsg = "正在处理[" & Me.cbo病区.List(i) & "]病人ID=" & lng病人ID & ";主页ID=" & lng主页ID & ";婴儿=" & int婴儿 & ";文件格式ID:" & lng格式ID & " 的汇总数据时发生错误"
'                                objStream.WriteLine "发生错误,跳过该病人:" & vbCrLf & strMsg
'                                Exit Do
'                            Else
'                                'Call WriteLog(objStream, "产生汇总数据...完成")
'                            End If
'                        End If
'                    End If
                    lng格式ID = rsFile!ID
                    lng保留 = rsFile!保留
                    
                    '先产生护理文件列表
                    If (rsFile!保留 = -1 And bln体温 = False) Or rsFile!保留 <> -1 Then
                        lng文件ID = GetNextId("病人护理文件")
                        strSQL = "insert into 病人护理文件(ID,科室ID,病人ID,主页ID,婴儿,格式ID,文件名称," & _
                                                          "开始时间,结束时间,续打ID,归档人,归档时间,创建人,创建时间)" & _
                                " Values (" & lng文件ID & "," & CLng(rsFile!科室ID) & "," & rsPati!病人ID & "," & rsPati!主页ID & "," & rsPati!婴儿 & "," & lng格式ID & ",'" & "[" & rsFile!科室 & "]" & rsFile!文件 & "'," & _
                                         "to_date('" & Format(datStart, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss'),NULL,NULL,'" & str归档人 & "'," & _
                                         "" & IIf(str归档时间 = "", "NULL", "to_date('" & str归档时间 & "','yyyy-MM-dd hh24:mi:ss')") & ",'ZLHIS',sysdate)"
                        gcnOracle.Execute strSQL
                    End If
                    
                    If rsFile!保留 = -1 Then   '体温单
                        If bln体温 = False Then
                            bln体温 = True
                            
                            '--3.1、数据来源全部处理为手工录入的数据，允许操作员使用新版程序时可直接更改，绑定变量有：病人ID，主页ID，婴儿，开始时间，结束时间
                            '--     体温单曲线项目数据与表格项目，也就是病人护理记录中的记录类型 IN (1,5)，终止版本为空的数据为最新的数据
                            '--     体温单中的上/下标、手术标记等内容，也就是病人护理记录中的记录类型 NOT IN （1，5）
                            'Call WriteLog(objStream, "病人ID=" & lng病人ID & ";主页ID=" & lng主页ID & ";婴儿=" & int婴儿 & ";文件格式ID=" & lng格式ID & "准备产生体温单主表数据")
                            '插入体温单主表数据(删掉条件: And A.科室ID=[1],体温单提取所有数据)
                            strSQL = "" & _
                                    " SELECT   A.ID,A.发生时间,A.最后版本,A.保存人,A.保存时间" & vbNewLine & _
                                    " FROM 病人护理记录 A,病人护理内容 B,体温记录项目 C" & vbNewLine & _
                                    " WHERE A.ID=B.记录ID ANd A.病人ID=[2] AND A.主页ID=[3] AND A.婴儿=[4] AND A.发生时间 BETWEEN [5] AND [6] And B.项目序号=C.项目序号" & vbNewLine & _
                                    " UNION" & vbNewLine & _
                                    " SELECT A.ID,A.发生时间,A.最后版本,A.保存人,A.保存时间" & vbNewLine & _
                                    " FROM 病人护理记录 A,病人护理内容 B" & vbNewLine & _
                                    " WHERE A.ID=B.记录ID ANd A.病人ID=[2] AND A.主页ID=[3] AND A.婴儿=[4] AND A.发生时间 BETWEEN [5] AND [6] And B.记录类型 NOT IN (1,5,9)"
                            strSQL = "" & _
                                    " INSERT INTO 病人护理数据(ID,文件ID,发生时间,最后版本,保存人,保存时间)" & vbNewLine & _
                                    " SELECT 病人护理数据_ID.Nextval,[7],A.发生时间,A.最后版本,A.保存人,A.保存时间" & vbNewLine & _
                                    " From (" & strSQL & ") A"
                            If blnDataMoved Then
                                strSQL = Replace(strSQL, "病人护理记录", "H病人护理记录")
                                strSQL = Replace(strSQL, "病人护理内容", "H病人护理内容")
                            End If
                            Set rsTemp = OpenSQLRecord(strSQL, "插入体温单主表数据", CLng(rsFile!科室ID), lng病人ID, lng主页ID, int婴儿, datStart, datEnd, lng文件ID)
                            'Call WriteLog(objStream, "病人ID=" & lng病人ID & ";主页ID=" & lng主页ID & ";婴儿=" & int婴儿 & ";文件格式ID=" & lng格式ID & "准备产生体温单主表数据...完成")
                            
                            '插入体温单明细表数据
                            'Call WriteLog(objStream, "病人ID=" & lng病人ID & ";主页ID=" & lng主页ID & ";婴儿=" & int婴儿 & ";文件格式ID=" & lng格式ID & "准备产生体温单明细表数据")
                            strSQL = "" & _
                                    " SELECT /*+ RULE */  A.ID, D.ID AS 记录ID, A.记录类型, A.项目分组, A.项目ID, A.项目序号, A.项目名称, A.项目类型, A.记录内容, A.项目单位, A.记录标记," & vbNewLine & _
                                    "      A.体温部位, A.记录组号, A.复试合格,9 AS 数据来源, A.未记说明, A.开始版本, A.终止版本, A.记录人,B.保存时间 AS 修改时间" & vbNewLine & _
                                    " FROM 病人护理内容 A,病人护理记录 B,体温记录项目 C,病人护理数据 D" & vbNewLine & _
                                    " WHERE A.记录ID=B.ID AND B.病人ID=[2] AND B.主页ID=[3] AND B.婴儿=[4] AND B.发生时间 BETWEEN [5] AND [6] And A.项目序号=C.项目序号" & vbNewLine & _
                                    " And B.发生时间=D.发生时间 And D.文件ID=[7]" & vbNewLine & _
                                    " UNION" & vbNewLine & _
                                    " SELECT A.ID, D.ID AS 记录ID, A.记录类型, A.项目分组, A.项目ID, A.项目序号, A.项目名称, A.项目类型, A.记录内容, A.项目单位, A.记录标记," & vbNewLine & _
                                    "     A.体温部位, A.记录组号, A.复试合格,9 AS 数据来源, A.未记说明, A.开始版本, A.终止版本, A.记录人,B.保存时间 AS 修改时间" & vbNewLine & _
                                    " FROM 病人护理内容 A,病人护理记录 B,病人护理数据 D" & vbNewLine & _
                                    " WHERE A.记录ID=B.ID AND B.病人ID=[2] AND B.主页ID=[3] AND B.婴儿=[4] AND B.发生时间 BETWEEN [5] AND [6] And A.记录类型 NOT IN (1,5,9)" & vbNewLine & _
                                    " And B.发生时间=D.发生时间 And D.文件ID=[7]"
                            strSQL = "" & _
                                    " INSERT INTO 病人护理明细(ID, 记录ID, 记录类型, 项目分组, 项目ID, 项目序号, 项目名称, 项目类型, 记录内容, 项目单位, 记录标记," & vbNewLine & _
                                    "                       体温部位, 记录组号, 复试合格,数据来源, 未记说明, 开始版本, 终止版本, 记录人,记录时间)" & vbNewLine & _
                                    " SELECT 病人护理明细_ID.Nextval, 记录ID, 记录类型, 项目分组, 项目ID, 项目序号, 项目名称, 项目类型, 记录内容, 项目单位, 记录标记," & vbNewLine & _
                                    "      体温部位, 记录组号, 复试合格, 数据来源, 未记说明, 开始版本, 终止版本, 记录人,修改时间" & vbNewLine & _
                                    " From (" & strSQL & ") "
                            If blnDataMoved Then
                                strSQL = Replace(strSQL, "病人护理记录", "H病人护理记录")
                                strSQL = Replace(strSQL, "病人护理内容", "H病人护理内容")
                            End If
                            Set rsTemp = OpenSQLRecord(strSQL, "插入体温单明细表数据", CLng(rsFile!科室ID), lng病人ID, lng主页ID, int婴儿, datStart, datEnd, lng文件ID)
                            'Call WriteLog(objStream, "病人ID=" & lng病人ID & ";主页ID=" & lng主页ID & ";婴儿=" & int婴儿 & ";文件格式ID=" & lng格式ID & "准备产生体温单明细表数据...完成")
                        End If
                    Else        '护理记录单
                        '--3.2、循环提取该病人指定的护理文件数据（根据病历文件格式），并产生到新版的病人护理数据、病人护理明细中，绑定变量有：文件ID、病人ID、主页ID、婴儿、开始时间、结束时间
                        '--     a)护理文件的汇总数据以前是虚拟产生的
                        '--老版程序护理记录单没有活动项目的概念,定义了哪些项目就只显示哪些项目,插入数据时不处理活动项目以及体温单特有的入出转,分娩,上下标等信息
                        'Call WriteLog(objStream, "病人ID=" & lng病人ID & ";主页ID=" & lng主页ID & ";婴儿=" & int婴儿 & ";文件格式ID=" & lng格式ID & "准备产生护理记录单主表数据" & IIf(rsFile!保留 = -1, "", "[" & rsFile!科室 & "]") & rsFile!文件)
                        '插入护理文件主表数据
                        strSQL = "" & _
                                " INSERT INTO 病人护理数据(ID,文件ID,发生时间,最后版本,保存人,保存时间)" & vbNewLine & _
                                " Select 病人护理数据_ID.Nextval,[8],发生时间,最后版本,保存人,保存时间" & vbNewLine & _
                                " FROM (" & vbNewLine & _
                                "   SELECT /*+ RULE */  DISTINCT ID,发生时间,最后版本,保存人,保存时间" & vbNewLine & _
                                "   FROM (" & vbNewLine & _
                                "       SELECT C.*" & vbNewLine & _
                                "       FROM" & vbNewLine & _
                                "         (SELECT * FROM 病历文件结构 WHERE 父ID=(SELECT DISTINCT ID FROM 病历文件结构 WHERE 对象类型=1 AND 对象序号=4 AND 文件ID=[2] )) A," & vbNewLine & _
                                "         护理记录项目 B,病人护理记录 C,病人护理内容 D" & vbNewLine & _
                                "       WHERE 文件ID=[2] AND 对象类型=4 AND 要素名称=B.项目名称" & vbNewLine & _
                                "       And C.科室ID=[1] AND C.病人ID=[3] AND C.主页ID=[4] AND NVL(C.婴儿,0)=[5]" & vbNewLine & _
                                "       AND C.发生时间 BETWEEN [6] AND [7]" & vbNewLine & _
                                "       AND D.记录ID=C.ID AND D.记录类型 IN (1,5) AND D.项目序号 =B.项目序号" & vbNewLine & _
                                "       UNION"
                        strSQL = strSQL & _
                                "       SELECT  C.*" & vbNewLine & _
                                "       FROM 病人护理记录 C,病人护理内容 D," & vbNewLine & _
                                "           (SELECT DISTINCT C.*" & vbNewLine & _
                                "           FROM" & vbNewLine & _
                                "               (SELECT * FROM 病历文件结构 WHERE 父ID=(SELECT DISTINCT ID FROM 病历文件结构 WHERE 对象类型=1 AND 对象序号=4 AND 文件ID=[2] )) A," & vbNewLine & _
                                "               护理记录项目 B,病人护理记录 C,病人护理内容 D" & vbNewLine & _
                                "           WHERE 文件ID=[2] AND 对象类型=4 AND 要素名称=B.项目名称" & vbNewLine & _
                                "           And C.科室ID=[1] AND C.病人ID=[3] AND C.主页ID=[4] AND NVL(C.婴儿,0)=[5]" & vbNewLine & _
                                "           AND C.发生时间 BETWEEN [6] AND [7]" & vbNewLine & _
                                "           AND D.记录ID=C.ID AND D.记录类型=1 AND D.项目序号 =B.项目序号) A" & vbNewLine & _
                                "       WHERE C.科室ID=[1] And C.病人ID=[3] AND C.主页ID=[4] AND NVL(C.婴儿,0)=[5]" & vbNewLine & _
                                "       AND C.发生时间 between [6] AND [7]" & vbNewLine & _
                                "       AND D.记录ID=C.ID AND D.记录类型=5 And C.ID=A.ID))"
                        If blnDataMoved Then
                            strSQL = Replace(strSQL, "病人护理记录", "H病人护理记录")
                            strSQL = Replace(strSQL, "病人护理内容", "H病人护理内容")
                        End If
                        Set rsTemp = OpenSQLRecord(strSQL, "插入护理文件主表数据", CLng(rsFile!科室ID), lng格式ID, lng病人ID, lng主页ID, int婴儿, datStart, datEnd, lng文件ID)
                        'Call WriteLog(objStream, "病人ID=" & lng病人ID & ";主页ID=" & lng主页ID & ";婴儿=" & int婴儿 & ";文件格式ID=" & lng格式ID & "准备产生护理记录单主表数据...完成")
                        
                        '插入护理文件明细表数据
                        'Call WriteLog(objStream, "病人ID=" & lng病人ID & ";主页ID=" & lng主页ID & ";婴儿=" & int婴儿 & ";文件格式ID=" & lng格式ID & "准备产生护理记录单明细表数据")
                        strSQL = "" & _
                                " SELECT /*+ RULE */ D.ID, Z.ID AS 记录ID, D.记录类型, D.项目分组, D.项目ID, D.项目序号, D.项目名称, D.项目类型, D.记录内容, D.项目单位, D.记录标记," & vbNewLine & _
                                "      D.体温部位, D.记录组号, D.复试合格,0 AS 数据来源, D.未记说明, D.开始版本, D.终止版本, D.记录人,C.保存时间 AS 修改时间" & vbNewLine & _
                                " FROM" & vbNewLine & _
                                "    (SELECT * FROM 病历文件结构 WHERE 父ID=(SELECT DISTINCT ID FROM 病历文件结构 WHERE 对象类型=1 AND 对象序号=4 AND 文件ID=[2] )) A," & vbNewLine & _
                                "    护理记录项目 B,病人护理记录 C,病人护理内容 D,病人护理数据 Z" & vbNewLine & _
                                " WHERE A.文件ID=[2] AND 对象类型=4 AND 要素名称=B.项目名称" & vbNewLine & _
                                " And C.科室ID=[1] AND C.病人ID=[3] AND C.主页ID=[4] AND NVL(C.婴儿,0)=[5]" & vbNewLine & _
                                " AND C.发生时间 BETWEEN [6] AND [7] And C.发生时间=Z.发生时间 And Z.文件ID=[8]" & vbNewLine & _
                                " AND D.记录ID=C.ID AND D.记录类型=1 AND D.项目序号 =B.项目序号" & vbNewLine & _
                                " UNION"
                        strSQL = strSQL & _
                                " SELECT  D.ID, Z.ID AS 记录ID, D.记录类型, D.项目分组, 项目ID, D.项目序号,D.项目名称, D.项目类型, D.记录内容, D.项目单位, D.记录标记," & vbNewLine & _
                                "     D.体温部位, D.记录组号, D.复试合格,0 AS 数据来源, D.未记说明, D.开始版本, D.终止版本, D.记录人,C.保存时间 AS 修改时间" & vbNewLine & _
                                " FROM 病人护理记录 C,病人护理内容 D," & vbNewLine & _
                                "    (SELECT DISTINCT C.*" & vbNewLine & _
                                "    FROM" & vbNewLine & _
                                "       (SELECT * FROM 病历文件结构 WHERE 父ID=(SELECT DISTINCT ID FROM 病历文件结构 WHERE 对象类型=1 AND 对象序号=4 AND 文件ID=[2] )) A," & vbNewLine & _
                                "    护理记录项目 B,病人护理记录 C,病人护理内容 D" & vbNewLine & _
                                "    WHERE A.文件ID=[2] AND 对象类型=4 AND 要素名称=B.项目名称" & vbNewLine & _
                                "    And C.科室ID=[1] AND C.病人ID=[3] AND C.主页ID=[4] AND NVL(C.婴儿,0)=[5]" & vbNewLine & _
                                "    AND C.发生时间 BETWEEN [6] AND [7]" & vbNewLine & _
                                "    AND D.记录ID=C.ID AND D.记录类型=1 AND D.项目序号 =B.项目序号) A,病人护理数据 Z" & vbNewLine & _
                                " WHERE C.科室ID=[1] And C.病人ID=[3] AND C.主页ID=[4] AND NVL(C.婴儿,0)=[5]" & vbNewLine & _
                                " AND C.发生时间 between [6] AND [7] And C.发生时间=Z.发生时间 And Z.文件ID=[8]" & vbNewLine & _
                                " AND D.记录ID=C.ID AND D.记录类型=5 And C.ID=A.ID"
                        strSQL = "" & _
                                " INSERT INTO 病人护理明细(ID, 记录ID, 记录类型, 项目分组, 项目ID, 项目序号, 项目名称, 项目类型, 记录内容, 项目单位, 记录标记," & vbNewLine & _
                                "                       体温部位, 记录组号, 复试合格,数据来源, 未记说明, 开始版本, 终止版本, 记录人,记录时间)" & vbNewLine & _
                                " SELECT 病人护理明细_ID.Nextval, 记录ID, 记录类型, 项目分组, DECODE(记录类型,5,NULL,项目ID), 项目序号, DECODE(记录类型,5,NULL,项目名称), 项目类型, 记录内容, 项目单位, 记录标记," & vbNewLine & _
                                "      体温部位, 记录组号, 复试合格, 数据来源, 未记说明, 开始版本, 终止版本, 记录人,修改时间" & vbNewLine & _
                                " FROM (" & strSQL & ")"
                        If blnDataMoved Then
                            strSQL = Replace(strSQL, "病人护理记录", "H病人护理记录")
                            strSQL = Replace(strSQL, "病人护理内容", "H病人护理内容")
                        End If
                        Set rsTemp = OpenSQLRecord(strSQL, "插入护理文件明细表数据", CLng(rsFile!科室ID), lng格式ID, lng病人ID, lng主页ID, int婴儿, datStart, datEnd, lng文件ID)
                        'Call WriteLog(objStream, "病人ID=" & lng病人ID & ";主页ID=" & lng主页ID & ";婴儿=" & int婴儿 & ";文件格式ID=" & lng格式ID & "准备产生护理记录单明细表数据...完成")
                        
                    End If
                    '下一个护理文件
                    rsFile.MoveNext
                    DoEvents
                Loop
                
                '自动执行时检查时间 , 前4小时升迁数据, 后2小时进行数据打印解析
                If gintAutoRUN = 1 Then
                    If Format(Now, "HH:mm") >= gstrNextTime Then
                        blnFirst = False    '避免再次进入下一次循环
                        GoTo todoPrint
                    End If
                End If
                
                '下一个病人
                .MoveNext
            Loop
            
            blnCommit = True '循环完了应该提交了
        End With
    Next
    
todoPrint:
    '最后一个病人的最后一个文件格式需要产生汇总数据
    If lng病人ID <> 0 Then
        'Call WriteLog(objStream, "病人ID=" & lng病人ID & ";主页ID=" & lng主页ID & ";婴儿=" & int婴儿 & ";文件格式ID=" & lng格式ID & "准备产生汇总数据")
        If Not InsertCollect(lng格式ID, lng病人ID, lng主页ID, int婴儿, datStart, datEnd, lng文件ID) Then
            blnError = True
            strMsg = "正在处理[" & Me.cbo病区.List(i) & "]病人ID=" & lng病人ID & ";主页ID=" & lng主页ID & ";婴儿=" & int婴儿 & ";文件格式ID=" & lng格式ID & "  的汇总数据时发生错误"
            objStream.WriteLine "发生错误,跳过该病人:" & vbCrLf & strMsg
        Else
            'Call WriteLog(objStream, "产生汇总数据...完成")
        End If
        
        '插入升迁记录
        gcnOracle.Execute "zl_护理升迁记录_Insert(" & lng病人ID & "," & lng主页ID & "," & int婴儿 & "," & IIf(blnError, "1", "0") & ",'" & Replace(strMsg, "'", "") & "')", , adCmdStoredProc
    End If
    
    If Not blnError Then
        gcnOracle.CommitTrans
    Else
        gcnOracle.RollbackTrans
    End If
    
    '如果超过4点,则不进入二次循环,上面已对blnFirst赋值为假
    If gintAutoRUN = 1 Then
        If blnFirst Then
            blnFirst = False
            gcnOracle.BeginTrans
            GoTo redo
        End If
    End If
    
    objStream.WriteLine Format(Now, "yyyy-MM-dd HH:mm:ss") & "数据升迁成功!"
    objStream.Close
    
    Me.Caption = "正在进行打印数据解析,请稍候..."
    Call DoPrintData
    
    If gintAutoRUN = 1 Then Unload Me
    Command1.Enabled = True
    Exit Sub
errHand:
    blnError = True
    strMsg = "正在处理[" & Me.cbo病区.List(i) & "]病人ID=" & lng病人ID & ";主页ID=" & lng主页ID & ";婴儿=" & int婴儿 & ";文件格式ID=" & lng格式ID & "  的护理数据时发生错误:" & Err.Description
    objStream.WriteLine "发生错误,跳过该病人:" & vbCrLf & strMsg
    GoTo EXITDO
End Sub

Private Function InsertCollect(ByVal lng格式ID As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal int婴儿 As Integer, _
    ByVal datStart As Date, ByVal datEnd As Date, ByVal lng文件ID As Long) As Boolean
    Dim datCur As Date
    Dim lngID As Long
    Dim strSQL As String
    Dim str日期 As String, lng对象序号 As Long, lng文件ID_Cur As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
'    '--一个病人一份文件只统计一次
'    '--3.3、已签名数据更新
'    strSQL = " SELECT distinct B.ID,Nvl(D.聘任技术职务,5)-1 AS 签名级别" & _
'             " From 病人护理文件 A,病人护理数据 B,病人护理明细 C,人员表 D" & _
'             " Where A.ID=[1] And A.ID=B.文件ID And B.ID=C.记录ID And C.记录类型=5 And B.保存人=D.姓名"
'    Set rsTemp = OpenSQLRecord(strSQL, "更新签名数据", lng文件ID)
'    Do While Not rsTemp.EOF
'        gcnOracle.Execute "Update 病人护理数据 Set 签名人=保存人,签名时间=sysdate,签名级别=" & rsTemp!签名级别 & " Where ID=" & rsTemp!ID
'        rsTemp.MoveNext
'    Loop
    
'    '--3.4、产生护理记录单的汇总数据,以前是展现数据时程序汇总的,绑定变量:文件ID,病人ID,主页ID,婴儿
'    '--根据护理文件格式定义提取汇总数据,程序循环插入主表,明细表
'    '--注意:小结名称做为项目名称,结束时间+对象序号做为发生时间,汇总量做为记录内容,日期+对象序号发生变化时,取病人护理记录_Nextval
'    '循环插入汇总数据
'    strSQL = "" & _
'            " SELECT A.日期,A.对象序号,A.小结名称,A.开始时间,A.结束时间,B.项目分组,B.项目序号,B.项目名称,B.项目单位,SUM(zl_to_number(B.记录内容,2)) AS 汇总量" & vbNewLine & _
'            " FROM" & vbNewLine & _
'            "    (SELECT" & vbNewLine & _
'            "        B.日期,A.对象序号,A.小结名称,B.日期||' '||DECODE(SIGN(LENGTH(A.开始时间)-8),-1,'0','')||A.开始时间 AS 开始时间," & vbNewLine & _
'            "        DECODE(SIGN(TO_NUMBER(SUBSTR(A.开始时间,1,INSTR(A.开始时间,':',1)-1))-TO_NUMBER(SUBSTR(A.结束时间,1,INSTR(A.结束时间,':',1)-1)))," & vbNewLine & _
'            "            1,TO_CHAR(TO_DATE(B.日期,'YYYY-MM-DD')+1,'YYYY-MM-DD'),B.日期)||' '||DECODE(SIGN(LENGTH(A.结束时间)-8),-1,'0','')||A.结束时间 AS 结束时间" & vbNewLine & _
'            "    FROM"
'    strSQL = strSQL & _
'            "        (SELECT 对象序号,小结名称," & vbNewLine & _
'            "        开始时间||DECODE(INSTR(开始时间,':',1),0,':00:00',':00') AS 开始时间," & vbNewLine & _
'            "        结束时间||DECODE(INSTR(结束时间,':',1),0,':59:59',':59') AS 结束时间" & vbNewLine & _
'            "        FROM (" & vbNewLine & _
'            "            SELECT 对象序号,SUBSTR(内容文本,1,INSTR(内容文本 ,',',1,1)-1) AS 小结名称," & vbNewLine & _
'            "                   replace(SUBSTR(内容文本,INSTR(内容文本 ,',',1,1)+1,INSTR(内容文本 ,',',1,2)-INSTR(内容文本 ,',',1,1)-1),'：',':') AS 开始时间," & vbNewLine & _
'            "                   replace(SUBSTR(内容文本,INSTR(内容文本 ,',',1,2)+1,LENGTH(内容文本)-INSTR(内容文本 ,',',1,2)),'：',':') AS 结束时间" & vbNewLine & _
'            "            FROM 病历文件结构" & vbNewLine & _
'            "            WHERE 父ID=(SELECT DISTINCT ID FROM 病历文件结构 WHERE 对象类型=1 AND 对象序号=5 AND 文件ID=[1]))) A," & vbNewLine & _
'            "        (SELECT DISTINCT TO_CHAR(发生时间,'YYYY-MM-DD') AS 日期 FROM 病人护理记录 A WHERE A.病人ID=[2] AND A.主页ID=[3] AND A.婴儿=[4] ) B) A," & vbNewLine & _
'            "    (SELECT A.发生时间,C.项目序号,C.项目名称,B.记录内容,C.项目单位" & vbNewLine & _
'            "    FROM 病人护理记录 A,病人护理内容 B,护理记录项目 C,病历文件结构 D" & vbNewLine & _
'            "    WHERE A.ID=B.记录ID AND A.病人ID=[2] AND A.主页ID=[3] AND A.婴儿=[4] AND A.发生时间 between [5] AND [6]" & vbNewLine & _
'            "    AND B.项目序号=C.项目序号 AND D.要素名称=C.项目名称 AND NVL(D.要素表示,0)=1 AND D.文件ID=[1]) B" & vbNewLine & _
'            "WHERE B.发生时间 BETWEEN TO_DATE(A.开始时间,'YYYY-MM-DD HH24:MI:SS') AND TO_DATE(A.结束时间,'YYYY-MM-DD HH24:MI:SS')" & vbNewLine & _
'            "GROUP BY A.日期,A.对象序号,A.小结名称,A.开始时间,A.结束时间,B.项目序号,B.项目名称,B.项目单位" & vbNewLine & _
'            "ORDER BY 日期,对象序号"
'    Set rsTemp = OpenSQLRecord(strSQL, "插入护理文件明细表数据", lng格式ID, lng病人ID, lng主页ID, int婴儿, datStart, datEnd)
'    With rsTemp
'        Do While Not .EOF
'            If str日期 <> !日期 Or lng对象序号 <> !对象序号 Then
'                str日期 = !日期
'                lng对象序号 = !对象序号
'
'                '根据结束时间取所在科室(有可能内一转内二)
'                datCur = !结束时间
'                strSQL = "" & _
'                        " SELECT A.科室ID,B.ID AS 文件ID" & vbNewLine & _
'                        " FROM 病人变动记录 A,病人护理文件 B" & vbNewLine & _
'                        " WHERE A.科室ID=B.科室ID And A.病人ID=B.病人ID And A.主页ID=B.主页ID " & vbNewLine & _
'                        " And B.格式ID=[5] And A.病人ID=[1] AND A.主页ID=[2] And B.婴儿=[3]" & vbNewLine & _
'                        " AND [4] BETWEEN A.开始时间 AND NVL(A.终止时间,SYSDATE)"
'                Set rsDept = OpenSQLRecord(strSQL, "提取结束时间病人所属科室", lng病人ID, lng主页ID, int婴儿, datCur, lng格式ID)
'                lngID = 0
'                lng文件ID_Cur = 0
'                'lng文件ID_Cur = lng文件ID      '不设置缺省值,是哪个文件的就产生到哪个文件中去
'
'                If rsDept.RecordCount <> 0 Then
'                    lng文件ID_Cur = rsDept!文件ID
'                End If
'
'                If lng文件ID_Cur <> 0 Then
'                    lngID = GetNextId("病人护理数据")
'                    '插入主记录
'                    strSQL = "" & _
'                            " INSERT INTO 病人护理数据(ID,文件ID,发生时间,最后版本,保存人,保存时间,汇总类别,汇总文本,汇总标记)" & vbNewLine & _
'                            " Values (" & lngID & "," & lng文件ID_Cur & ",to_date('" & Format(DateAdd("s", !对象序号, !结束时间), "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss')," & _
'                                     "NULL,'ZLHIS',sysdate," & -1 * !对象序号 & ",'" & !小结名称 & "',0)"
'                    gcnOracle.Execute strSQL
'                End If
'            End If
'
'            If lngID <> 0 Then
'                '插入明细数据
'                strSQL = "" & _
'                        " INSERT INTO 病人护理明细(ID, 记录ID, 记录类型, 项目分组, 项目ID, 项目序号, 项目名称, 项目类型, 记录内容, 项目单位, 记录标记," & vbNewLine & _
'                        "                       体温部位, 记录组号, 复试合格,数据来源, 未记说明, 开始版本, 终止版本, 记录人,记录时间)" & vbNewLine & _
'                        " Values (病人护理明细_ID.Nextval," & lngID & ",1,'" & !项目分组 & "',NULL," & !项目序号 & ",'" & !项目名称 & "',0,'" & !汇总量 & "','" & NVL(!项目单位) & "',0," & _
'                                 "NULL,NULL,0,0,NULL,1,NULL,'ZLHIS',sysdate)"
'                gcnOracle.Execute strSQL
'            End If
'
'            .MoveNext
'        Loop
'    End With
    
    InsertCollect = True
errHand:
    Exit Function
End Function

Private Sub WriteLog(ByVal objStream As TextStream, ByVal strLog As String)
    objStream.WriteLine "时间:" & Format(Now, "yyyy-MM-dd HH:mm:ss") & ";" & strLog
End Sub

Private Sub DoPrintData()
    Dim objFile As New FileSystemObject
    Dim objStream As TextStream
    Dim arrData
    Dim lngRows As Long, lngParent As Long
    Dim rsTemp As New ADODB.Recordset
    Dim rsFormat As New ADODB.Recordset
    On Error GoTo errHand
    '按文件循环，对所有相关病人的数据进行打印数据解析
    Set objStream = objFile.OpenTextFile("C:\" & IIf(gintAutoRUN = 1, "AUTO", "") & "PRINT_LOG" & Format(Now, "yyyyMMddHHmmss") & ".txt", ForWriting, True)
    
    strSQL = " Select ID,编号,名称 From 病历文件列表 Where 种类=3 And 保留<>-1 and 通用<>0 Order by 编号"
    Call OpenRecordset(rsFile, strSQL, "提取护理文件列表")
    
    With rsFile
        Do While Not .EOF
            objStream.WriteLine String(50, "-")
            objStream.WriteLine "文件：" & rsFile!ID & "于" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "开始进行打印解析"
            
            '读取页面格式（后面多处需要使用）
            '(纸张|纸向|高|宽|上边距|下边距|左边距|右边距|行高|固定行数|标题栏字体名|标题栏字体大小|标题文本|表上项字体名|表上项字体大小|表上项文本)
            strSQL = "" & _
                    " SELECT B.ID,A.种类,A.编号,B.名称," & vbNewLine & _
                    "       SUBSTR(A.格式,1,INSTR(A.格式,';',1,1)-1) AS PAGE ," & vbNewLine & _
                    "       SUBSTR(A.格式,INSTR(A.格式,';',1,1)+1,INSTR(A.格式,';',1,2)-INSTR(A.格式,';',1,1)-1) AS Orient," & vbNewLine & _
                    "       SUBSTR(A.格式,INSTR(A.格式,';',1,2)+1,INSTR(A.格式,';',1,3)-INSTR(A.格式,';',1,2)-1) AS HEIGHT ," & vbNewLine & _
                    "       SUBSTR(A.格式,INSTR(A.格式,';',1,3)+1,INSTR(A.格式,';',1,4)-INSTR(A.格式,';',1,3)-1) AS WIDTH ," & vbNewLine & _
                    "       SUBSTR(A.格式,INSTR(A.格式,';',1,4)+1,INSTR(A.格式,';',1,5)-INSTR(A.格式,';',1,4)-1) AS LEFT ," & vbNewLine & _
                    "       SUBSTR(A.格式,INSTR(A.格式,';',1,5)+1,INSTR(A.格式,';',1,6)-INSTR(A.格式,';',1,5)-1) AS RIGHT," & vbNewLine & _
                    "       SUBSTR(A.格式,INSTR(A.格式,';',1,6)+1,INSTR(A.格式,';',1,7)-INSTR(A.格式,';',1,6)-1) AS TOP," & vbNewLine & _
                    "       SUBSTR(A.格式,INSTR(A.格式,';',1,7)+1,DECODE(INSTR(A.格式,';',1,8),0,LENGTH(格式)+1,INSTR(A.格式,';',1,8))-INSTR(A.格式,';',1,7)-1) AS BOTTOM" & vbNewLine & _
                    " FROM 病历页面格式 A,病历文件列表 B " & _
                    " WHERE A.种类=B.种类 AND A.编号=B.编号 AND B.种类=3 AND B.保留<>-1 And B.ID=[1]" & vbNewLine & _
                    " ORDER BY 编号"
            Set rsFormat = OpenSQLRecord(strSQL, "提取护理文件格式", CLng(rsFile!ID))
            arrData = Split(rsFormat!Page & "," & rsFormat!orient & "," & rsFormat!Height & "," & rsFormat!Width & "," & rsFormat!Left & "," & rsFormat!Right & "," & rsFormat!Top & "," & rsFormat!Bottom, ",")
            
            '如果没有解析护理文件一页可显示多少行有效数据则先进行解析并更新
            strSQL = "select 父ID,内容文本,对象属性 from 病历文件结构 where 父ID=(select ID from 病历文件结构 where 文件ID=[1] and 对象序号=1 and 父ID is null)"
            Set rsTemp = OpenSQLRecord(strSQL, "提取当前文件有效数据行", CLng(rsFile!ID))
            lngParent = rsTemp!父ID
            rsTemp.Filter = "对象属性='有效数据行'"
            If rsTemp.RecordCount <> 0 Then
                lngRows = rsTemp!内容文本
            Else
                lngRows = frmPreview.ShowMe(Me, rsFile!ID, arrData)
                
                '产生数据供以后使用
                gcnOracle.Execute "insert into 病历文件结构(ID,文件ID,父ID,对象序号,对象类型,对象属性,内容文本,要素名称) select 病历文件结构_ID.Nextval," & rsFile!ID & "," & lngParent & ",13,4,'有效数据行'," & lngRows & ",'有效数据行' from dual"
            End If
            
            '按病人循环加载所有数据进行解析
            gcnOracle.BeginTrans
            Call frmPreview.AnaliseData(Me, rsFile!ID, arrData, objStream)
            gcnOracle.CommitTrans
            objStream.WriteLine "文件：" & rsFile!ID & "于" & Format(Now, "yyyy-MM-dd HH:mm:ss") & "完成打印解析！"
            
            If gintAutoRUN = 1 Then
                If Format(Now, "HH:mm") >= gstrEndTime Then
                    Exit Do
                End If
            End If
            
            .MoveNext
        Loop
    End With
    
    objStream.WriteLine Format(Now, "yyyy-MM-dd HH:mm:ss") & "完成"
    objStream.Close
    Unload frmPreview
    
    If gintAutoRUN = 0 Then MsgBox "护理数据打印解析完成！"
    Exit Sub
errHand:
    MsgBox Err.Description
    objStream.WriteLine Format(Now, "yyyy-MM-dd HH:mm:ss") & Err.Description
    objStream.Close
End Sub

Private Sub Command1_Click()
    If gintMode = -1 Then
        MsgBox "必须选择一种数据导入模式！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gstrStartTime = Format(Me.dtp开始时间.Value, "HH:mm")
    gstrNextTime = Format(Me.dtp开始时间1.Value, "HH:mm")
    gstrEndTime = Format(Me.dtp结束时间.Value, "HH:mm")
    SaveSetting "ZLSOFT", "私有模块\ZLHIS\护理数据升迁", "开始时间", gstrStartTime
    SaveSetting "ZLSOFT", "私有模块\ZLHIS\护理数据升迁", "开始时间1", gstrNextTime
    SaveSetting "ZLSOFT", "私有模块\ZLHIS\护理数据升迁", "结束时间", gstrEndTime
    
    If gintMode = 1 Then
        Call DataUpgrade
    Else
        Me.Hide
        tim.Enabled = True
    End If
End Sub

Private Sub Command2_Click()
    frmSet.Show 1, Me
End Sub

Private Sub Form_Activate()
    If gintAutoRUN = 1 Then Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Private Sub Form_Load()
    If gintAutoRUN = 1 Then
        If gintMode < 0 Then
            gintAutoRUN = 0
            MsgBox "自动运行需指定导入模式，现在进入手动模式！", vbInformation, gstrSysName
        ElseIf gintMode > 2 Then
            gintAutoRUN = 0
            MsgBox "指定的导入模式错误(0,1,2)，现在进入手动模式！", vbInformation, gstrSysName
        End If
    End If
    
    strSQL = "" & _
            " SELECT DISTINCT A.ID,A.编码,A.名称" & vbNewLine & _
            " FROM 部门表 A,部门性质说明 B" & vbNewLine & _
            " WHERE A.ID=B.部门ID AND B.服务对象 IN(1,2,3) AND B.工作性质='护理'" & vbNewLine & _
            " AND (A.撤档时间 IS NULL OR TRUNC(A.撤档时间)=TO_DATE('3000-01-01','YYYY-MM-DD'))" & vbNewLine & _
            " ORDER BY A.编码"
    Call OpenRecordset(rsTemp, strSQL, "提取本院所有病区")
    With rsTemp
        Me.cbo病区.Clear
        Me.cbo病区.AddItem "所有病区"
        Do While Not .EOF
            Me.cbo病区.AddItem !名称
            Me.cbo病区.ItemData(Me.cbo病区.NewIndex) = !ID
            .MoveNext
        Loop
        Me.cbo病区.ListIndex = 0
    End With
    
    gstrStartTime = GetSetting("ZLSOFT", "私有模块\ZLHIS\护理数据升迁", "开始时间", "00:00")
    gstrNextTime = GetSetting("ZLSOFT", "私有模块\ZLHIS\护理数据升迁", "开始时间1", "02:00")
    gstrEndTime = GetSetting("ZLSOFT", "私有模块\ZLHIS\护理数据升迁", "结束时间", "04:00")
    dtp开始时间.Value = gstrStartTime
    dtp开始时间1.Value = gstrNextTime
    dtp结束时间.Value = gstrEndTime
End Sub

Private Sub opt_Click(Index As Integer)
    Command1.Caption = "后台运行"
    Select Case Index
    Case 0
        Label2.Caption = "    当前在院病人及最近30天出院病人的历史护理数据(受时间设置影响)"
    Case 1
        Label2.Caption = "    升级当天仍在院病人及最近30天出院病人的所有护理数据"
        Command1.Caption = "数据升迁"
    Case 2
        Label2.Caption = "    所有历史病人护理数据(受时间设置影响)"
    End Select
    
    If gintAutoRUN = 1 Then Exit Sub
    gintMode = Index
End Sub

Private Sub tim_Timer()
    '数据升迁后会自动进行打印解析,所以此处只判断数据升迁的有效时间即可
    If gintMode = 1 Then Exit Sub   '升级模式直接运行
    If Not (Format(Now, "HH:mm") >= gstrStartTime And Format(Now, "HH:mm") <= gstrNextTime) Then Exit Sub
    If Not Command1.Enabled Then Exit Sub
    
    Call DataUpgrade
End Sub
