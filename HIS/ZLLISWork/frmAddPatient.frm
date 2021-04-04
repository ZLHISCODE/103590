VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "CO70B6~1.OCX"
Begin VB.Form frmAddPatient 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "病人标本增加"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   11730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra执行选项 
      Caption         =   "执行选项"
      Height          =   975
      Left            =   90
      TabIndex        =   25
      Top             =   6120
      Width           =   11565
      Begin MSComCtl2.DTPicker DTP 
         Height          =   285
         Left            =   1020
         TabIndex        =   49
         Top             =   570
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   245628931
         CurrentDate     =   39620
      End
      Begin VB.TextBox txt生成标本号 
         Height          =   270
         Left            =   6780
         TabIndex        =   47
         Top             =   570
         Width           =   1725
      End
      Begin VB.TextBox txt标本批号 
         Height          =   270
         Left            =   4050
         TabIndex        =   46
         Top             =   570
         Width           =   1725
      End
      Begin VB.ComboBox cbo检验仪器 
         Height          =   300
         Left            =   9630
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   210
         Width           =   1725
      End
      Begin VB.ComboBox cbo执行科室 
         Height          =   300
         Left            =   6780
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   210
         Width           =   1725
      End
      Begin VB.ComboBox cbo开单科室 
         Height          =   300
         Left            =   1020
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   210
         Width           =   2145
      End
      Begin VB.ComboBox cbo开单医生 
         Height          =   300
         Left            =   4050
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   210
         Width           =   1725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申请时间"
         Height          =   180
         Left            =   210
         TabIndex        =   48
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lbl标本批号 
         Caption         =   "标本批号"
         Height          =   195
         Left            =   3240
         TabIndex        =   37
         Top             =   600
         Width           =   825
      End
      Begin VB.Label lbl生成标本号 
         Caption         =   "标 本 号"
         Height          =   195
         Left            =   5940
         TabIndex        =   35
         Top             =   630
         Width           =   1035
      End
      Begin VB.Label lbl检验仪器 
         AutoSize        =   -1  'True
         Caption         =   "检验仪器"
         Height          =   180
         Left            =   8790
         TabIndex        =   33
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lbl执行科室 
         AutoSize        =   -1  'True
         Caption         =   "执行科室"
         Height          =   180
         Left            =   5925
         TabIndex        =   31
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lbl开单科室 
         AutoSize        =   -1  'True
         Caption         =   "开单科室"
         Height          =   180
         Left            =   210
         TabIndex        =   29
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lbl开单医生 
         AutoSize        =   -1  'True
         Caption         =   "开单医生"
         Height          =   180
         Left            =   3255
         TabIndex        =   27
         Top             =   270
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   10110
      TabIndex        =   14
      Top             =   7260
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确认(&O)"
      Height          =   350
      Left            =   8580
      TabIndex        =   13
      Top             =   7260
      Width           =   1100
   End
   Begin VB.Frame fra项目选择 
      Caption         =   "项目选择"
      Height          =   2805
      Left            =   90
      TabIndex        =   6
      Top             =   3270
      Width           =   11565
      Begin XtremeReportControl.ReportControl rptItemSelect 
         Height          =   2445
         Left            =   7170
         TabIndex        =   12
         Top             =   210
         Width           =   4305
         _Version        =   589884
         _ExtentX        =   7594
         _ExtentY        =   4313
         _StockProps     =   0
         BorderStyle     =   2
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         ShowItemsInGroups=   -1  'True
         AutoColumnSizing=   0   'False
      End
      Begin XtremeReportControl.ReportControl rptItemSource 
         Height          =   2445
         Left            =   2460
         TabIndex        =   7
         Top             =   210
         Width           =   3975
         _Version        =   589884
         _ExtentX        =   7011
         _ExtentY        =   4313
         _StockProps     =   0
         BorderStyle     =   2
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         ShowItemsInGroups=   -1  'True
         AutoColumnSizing=   0   'False
      End
      Begin VB.CommandButton cmd查找 
         Caption         =   "&P"
         Height          =   255
         Left            =   2010
         TabIndex        =   32
         Top             =   765
         Width           =   285
      End
      Begin VB.CommandButton cmdItemLeftAll 
         Caption         =   "<<"
         Height          =   375
         Left            =   6630
         TabIndex        =   24
         Top             =   2130
         Width           =   405
      End
      Begin VB.CommandButton cmdItemLeft 
         Caption         =   "<"
         Height          =   375
         Left            =   6630
         TabIndex        =   23
         Top             =   1590
         Width           =   405
      End
      Begin VB.CommandButton cmdItemRightAll 
         Caption         =   ">>"
         Height          =   375
         Left            =   6630
         TabIndex        =   22
         Top             =   1050
         Width           =   405
      End
      Begin VB.CommandButton cmdItemRight 
         Caption         =   ">"
         Height          =   375
         Left            =   6630
         TabIndex        =   21
         Top             =   540
         Width           =   405
      End
      Begin VB.OptionButton optIF 
         Caption         =   "单项"
         Height          =   225
         Index           =   1
         Left            =   960
         TabIndex        =   16
         Top             =   1200
         Width           =   765
      End
      Begin VB.OptionButton optIF 
         Caption         =   "组合"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   15
         Top             =   1200
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.TextBox txt查找 
         Height          =   270
         Left            =   750
         TabIndex        =   11
         Top             =   780
         Width           =   1275
      End
      Begin VB.ComboBox cbo类别 
         Height          =   300
         Left            =   750
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lbl查找 
         Caption         =   "查  找"
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   825
         Width           =   615
      End
      Begin VB.Label lbl类别 
         Caption         =   "类  别"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   420
         Width           =   555
      End
   End
   Begin VB.Frame fra病人选择 
      Caption         =   "病人选择"
      Height          =   3165
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   11565
      Begin XtremeReportControl.ReportControl rptListSelect 
         Height          =   2865
         Left            =   7170
         TabIndex        =   2
         Top             =   180
         Width           =   4275
         _Version        =   589884
         _ExtentX        =   7541
         _ExtentY        =   5054
         _StockProps     =   0
         BorderStyle     =   2
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         ShowItemsInGroups=   -1  'True
         AutoColumnSizing=   0   'False
      End
      Begin XtremeReportControl.ReportControl rptListSource 
         Height          =   2865
         Left            =   2460
         TabIndex        =   1
         Top             =   180
         Width           =   3975
         _Version        =   589884
         _ExtentX        =   7011
         _ExtentY        =   5054
         _StockProps     =   0
         BorderStyle     =   2
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         ShowItemsInGroups=   -1  'True
         AutoColumnSizing=   0   'False
      End
      Begin VB.ComboBox cbo仪器 
         Height          =   300
         Left            =   780
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   660
         Width           =   1575
      End
      Begin VB.CommandButton cmd过滤 
         Caption         =   "过滤(&F)"
         Height          =   350
         Left            =   1230
         TabIndex        =   44
         Top             =   2640
         Width           =   1100
      End
      Begin VB.ComboBox cbo科室 
         Height          =   300
         Left            =   780
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   270
         Width           =   1575
      End
      Begin VB.TextBox txt批号 
         Height          =   270
         Left            =   780
         TabIndex        =   40
         Top             =   1080
         Width           =   1545
      End
      Begin VB.TextBox txt标本号 
         Height          =   270
         Left            =   780
         TabIndex        =   38
         Top             =   1500
         Width           =   1545
      End
      Begin VB.CommandButton cmdPatientLeftAll 
         Caption         =   "<<"
         Height          =   375
         Left            =   6600
         TabIndex        =   20
         Top             =   2100
         Width           =   405
      End
      Begin VB.CommandButton cmdPatientLeft 
         Caption         =   "<"
         Height          =   375
         Left            =   6600
         TabIndex        =   19
         Top             =   1560
         Width           =   405
      End
      Begin VB.CommandButton cmdPatientRightAll 
         Caption         =   ">>"
         Height          =   375
         Left            =   6600
         TabIndex        =   18
         Top             =   1020
         Width           =   405
      End
      Begin VB.CommandButton cmdPatientRight 
         Caption         =   ">"
         Height          =   375
         Left            =   6600
         TabIndex        =   17
         Top             =   510
         Width           =   405
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   255
         Left            =   780
         TabIndex        =   4
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   245628929
         CurrentDate     =   39533
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   255
         Left            =   780
         TabIndex        =   5
         Top             =   2250
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   450
         _Version        =   393216
         Format          =   245628929
         CurrentDate     =   39533
      End
      Begin VB.Label lbl仪器 
         AutoSize        =   -1  'True
         Caption         =   "仪  器"
         Height          =   180
         Left            =   150
         TabIndex        =   36
         Top             =   720
         Width           =   540
      End
      Begin VB.Label lbl科室 
         Caption         =   "科  室"
         Height          =   195
         Left            =   150
         TabIndex        =   43
         Top             =   330
         Width           =   585
      End
      Begin VB.Label lbl批号 
         Caption         =   "批  号"
         Height          =   195
         Left            =   150
         TabIndex        =   41
         Top             =   1125
         Width           =   825
      End
      Begin VB.Label lbl标本号 
         Caption         =   "标本号"
         Height          =   195
         Left            =   150
         TabIndex        =   39
         Top             =   1545
         Width           =   555
      End
      Begin VB.Label lbl日期 
         Caption         =   "日  期"
         Height          =   195
         Left            =   150
         TabIndex        =   3
         Top             =   1950
         Width           =   585
      End
   End
End
Attribute VB_Name = "frmAddPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mlngMachineID As Long           '仪器ID
Dim mlngExecDeptID As Long          '执行科室ID

Private Enum mColP
    病人ID = 0: 标本序号: 姓名: 性别: 年龄: 婴儿: 挂号单: 门诊号: 住院号: 出生日期: 主页ID: 标识号: 床号: 病人科室: 申请人: 申请科室ID: 病人来源
End Enum
Private Enum mColI
    ID = 0: 编码: 名称: 标本: 项目类别
End Enum

Private Sub Option2_Click()

End Sub

Private Sub cbo检验仪器_Click()
    If Me.cbo检验仪器.ItemData(Me.cbo检验仪器.ListIndex) = -1 Then
        txt标本批号.Enabled = True
    Else
        txt标本批号.Text = ""
        txt标本批号.Enabled = False
    End If
End Sub

Private Sub cbo开单科室_Click()
    Dim lngApplyDept As Long
    Dim rsTmp As New ADODB.Recordset
    Dim lngKey As Long
    
    lngApplyDept = Me.cbo开单科室.ItemData(Me.cbo开单科室.ListIndex)
    lngKey = zldatabase.GetPara("frmAddPatient_开单医生", 100, 1208, -1)
    
    '读入对应科室下的人员
    gstrSql = "select distinct a.id,a.编号,a.姓名 from 人员表 a , 人员性质说明 b , 部门人员 c " & _
                 " where a.id = b.人员id and a.id = c.人员ID and  b.人员性质 in ('医生','护士') " & _
                 " and (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) "
    If lngApplyDept = 0 Then
        gstrSql = gstrSql & " order by a.编号"
    Else
        gstrSql = gstrSql & " and c.部门id = [1] order by a.编号 "
    End If
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, gstrSysName, lngApplyDept)
    
    With Me.cbo开单医生
        .Clear
        Do Until rsTmp.EOF
            .AddItem Nvl(rsTmp("编号")) & "-" & Nvl(rsTmp("姓名"))
            .ItemData(.NewIndex) = rsTmp("ID")
            If lngKey = rsTmp("ID") Then
                .ListIndex = .NewIndex
            End If
            rsTmp.MoveNext
        Loop
    End With
End Sub

Private Sub cbo科室_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim lngDept As Long
    Dim lngKey As Long
    
    lngKey = zldatabase.GetPara("frmAddPatient_选择仪器", 100, 1208, -1)
    gstrSql = "select ID,编码,名称  from 检验仪器 "
    
    lngDept = cbo科室.ItemData(cbo科室.ListIndex)
    If lngDept > 0 Then
        gstrSql = gstrSql & " Where 使用小组ID = [1] "
    End If
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngDept)
    With cbo仪器
        .Clear
        .AddItem "手工"
        .ItemData(.NewIndex) = -1
        Do Until rsTmp.EOF
            .AddItem Nvl(rsTmp("编码")) & "-" & Nvl(rsTmp("名称"))
            .ItemData(.NewIndex) = rsTmp("ID")
            If lngKey = rsTmp("ID") Then
                .ListIndex = .NewIndex
            End If
            rsTmp.MoveNext
        Loop
    End With
End Sub

Private Sub cbo执行科室_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim lngDept As Long
    Dim lngKey As Long
    
    lngKey = zldatabase.GetPara("frmAddPatient_检验仪器", 100, 1208, -1)
    gstrSql = "select ID,编码,名称  from 检验仪器 "
    If cbo执行科室.ListIndex >= 0 Then
        lngDept = cbo执行科室.ItemData(cbo执行科室.ListIndex)
    End If
    If lngDept > 0 Then
        gstrSql = gstrSql & " Where 使用小组ID = [1] "
    End If
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngDept)
    With cbo检验仪器
        .Clear
        .AddItem "手工"
        .ItemData(.NewIndex) = -1
        If lngKey = -1 Then .ListIndex = .NewIndex
        Do Until rsTmp.EOF
            .AddItem Nvl(rsTmp("编码")) & "-" & Nvl(rsTmp("名称"))
            .ItemData(.NewIndex) = rsTmp("ID")
            If lngKey = rsTmp("ID") Then
                .ListIndex = .NewIndex
            End If
            rsTmp.MoveNext
        Loop
        
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdItemLeft_Click()
    MoveItem 2, 1, False
End Sub

Private Sub cmdItemLeftAll_Click()
    MoveItem 2, 1, True
End Sub

Private Sub cmdItemRight_Click()
    MoveItem 2, 2, False
End Sub

Private Sub cmdItemRightAll_Click()
    MoveItem 2, 2, True
End Sub

Private Sub cmdOk_Click()
    If chkSaveData = True Then
        Call SaveData
    End If
End Sub

Private Sub cmdPatientLeft_Click()
    MoveItem 1, 1, False
End Sub

Private Sub cmdPatientLeftAll_Click()
    MoveItem 1, 1, True
End Sub

Private Sub cmdPatientRight_Click()
    MoveItem 1, 2, False
End Sub

Private Sub cmdPatientRightAll_Click()
    MoveItem 1, 2, True
End Sub

Private Sub cmd查找_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim strType As String
    Dim Record As ReportRecord
    Dim intLoop As Integer
    
    gstrSql = "Select Distinct A.ID, A.编码, B.名称, 标本部位, D.项目类别" & vbNewLine & _
                "From 诊疗项目目录 A, 诊疗项目别名 B, 检验报告项目 C, 检验项目 D" & vbNewLine & _
                "Where A.ID = B.诊疗项目id And A.类别 = 'C' And A.单独应用 = 1 And A.ID = C.诊疗项目id And " & vbNewLine & _
                " C.报告项目id = D.诊治项目id(+) And D.项目类别 In (1, 2,4) " & vbNewLine & _
                " And (A.撤档时间 Is Null Or To_Char(A.撤档时间, 'yyyy-mm-dd') = '3000-01-01') "
    
    If Me.cbo类别.ListIndex > 0 Then
        gstrSql = gstrSql & " And a.操作类型 = [2] "
        strType = Mid(Me.cbo类别.Text, InStr(Me.cbo类别, "-") + 1)
    End If
    
    If Trim(Me.txt查找.Text) <> "" Then
        gstrSql = gstrSql & " And (b.简码 like [3] or d.缩写 like [3]) "
    End If
    
    gstrSql = gstrSql & " And nvl(a.组合项目,0) = [4] "
   
    gstrSql = gstrSql & " Order by a.编码 "
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, IIf(Me.optIF(0).Value, 1, 0), strType, "%" & UCase(Me.txt查找) & "%", _
                            IIf(optIF(0).Value, 1, 0))
    rptItemSource.Records.DeleteAll:    rptItemSelect.Records.DeleteAll
    Do Until rsTmp.EOF
        Set Record = Me.rptItemSource.Records.Add
        For intLoop = 0 To Me.rptItemSource.Columns.Count
            Record.AddItem ""
        Next
        Record.Item(mColI.ID).Value = Nvl(rsTmp("ID"))
        Record.Item(mColI.编码).Value = Nvl(rsTmp("编码"))
        Record.Item(mColI.名称).Value = Nvl(rsTmp("名称"))
        Record.Item(mColI.标本).Value = Nvl(rsTmp("标本部位"))
        Record.Item(mColI.项目类别).Value = Nvl(rsTmp("项目类别"))
        rsTmp.MoveNext
    Loop
    rptItemSource.Populate: rptItemSelect.Populate
    Me.txt查找.Text = ""
End Sub

Private Sub cmd过滤_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim strWhere As String
    Dim lngDept As Long, lngMachine As Long
    Dim strBegingNO As String, strEndNO As String
    Dim astrItem() As String
    Dim Record As ReportRecord                                      '列表记录集
    
    On Error GoTo errH
    
    gstrSql = "Select to_Number(标本序号) as 标本序号,病人id, 婴儿,姓名, 性别, 年龄, 挂号单, 门诊号, 住院号, 出生日期, 主页id, 标识号, " & vbNewLine & _
              " 床号, 病人科室, 申请人, 申请科室id,病人来源, " & vbNewLine & _
              " Decode(仪器id, Null," & vbNewLine & _
              "                 To_Char(Trunc(标本序号 / 10000) + 1, '0000') || '-' || To_Char(Mod(标本序号, 10000), '0000')," & vbNewLine & _
              "                 标本序号) As 标本号显示 " & vbNewLine & _
              " From 检验标本记录 where 医嘱ID is not null and  核收时间 between [1] and [2] "
    '执行科室
    If cbo科室.Text <> "" Then
        lngDept = cbo科室.ItemData(cbo科室.ListIndex)
        If lngDept > 0 Then
            strWhere = " And 执行科室ID = [3] "
        End If
    End If
    '仪器
    If cbo仪器.Text <> "" Then
        lngMachine = cbo仪器.ItemData(cbo仪器.ListIndex)
        If lngMachine = -1 Then
            strWhere = strWhere & " And 仪器ID is null "
        ElseIf lngMachine > 0 Then
            strWhere = strWhere & " And 仪器ID = [4]"
        End If
    End If
    '标本号
    If Trim(txt标本号) <> "" Then
        txt标本号 = Replace(Replace(txt标本号, "～", "~"), "-", "~")
        varItem = Split(Trim(txt标本号.Text), ",")
        
        For lngLoop = 0 To UBound(varItem)
            astrItem = Split(varItem(lngLoop), "~")
            
            If UBound(astrItem) <= 0 Then
                strBegingNO = TransSampleNO(IIf(Val(Me.txt批号) <> 0, Val(Me.txt批号) & "-" & Val(varItem(lngLoop)), Val(varItem(lngLoop))))
                strEndNO = TransSampleNO(IIf(Val(Me.txt批号) <> 0, Val(Me.txt批号) & "-" & Val(varItem(lngLoop)), Val(varItem(lngLoop))))
            Else
                strBegingNO = TransSampleNO(IIf(Val(Me.txt批号) <> 0, Val(Me.txt批号) & "-" & Val(astrItem(0)), Val(astrItem(0))))
                strEndNO = TransSampleNO(IIf(Val(Me.txt批号) <> 0, Val(Me.txt批号) & "-" & Val(astrItem(1)), Val(astrItem(1))))
            End If
            If lngLoop = 0 Then
                strWhere = strWhere & " and (to_Number(标本序号) between " & Val(strBegingNO) & " and " & Val(strEndNO) & " "
            Else
                strWhere = strWhere & "  or to_Number(标本序号) between " & Val(strBegingNO) & " and " & Val(strEndNO) & " "
            End If
        Next
        If lngLoop >= 0 Then strWhere = strWhere & ")"
    ElseIf Trim(txt批号) <> "" Then
        strWhere = strWhere & " and to_Number(标本序号) between [5] and [6] "
        strBegingNO = TransSampleNO(Val(Me.txt批号) & "-0001")
        strEndNO = TransSampleNO(Val(Me.txt批号) & "-9999")
    End If
    
    gstrSql = gstrSql & strWhere & " Order by to_Number(标本序号) "
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, CDate(Format(dtpBegin.Value, "yyyy-mm-dd 00:00:00")), _
                CDate(Format(dtpEnd.Value, "yyyy-mm-dd 23:59:59")), lngDept, lngMachine, Val(strBegingNO), Val(strEndNO))
    Me.rptListSelect.Records.DeleteAll
    Me.rptListSource.Records.DeleteAll
    Me.rptListSource.Populate
    Me.rptListSelect.Populate
    Do Until rsTmp.EOF
        
        Set Record = Me.rptListSource.Records.Add
            For intLoop = 0 To Me.rptListSource.Columns.Count
                Record.AddItem ""
            Next
            Record.Item(mColP.病人ID).Value = Nvl(rsTmp("病人ID"))
            Record.Item(mColP.标本序号).Value = Nvl(rsTmp("标本序号"))
            Record.Item(mColP.标本序号).Caption = Nvl(rsTmp("标本号显示"))
            Record.Item(mColP.姓名).Value = Nvl(rsTmp("姓名"))
            Record.Item(mColP.性别).Value = Nvl(rsTmp("性别"))
            Record.Item(mColP.年龄).Value = Nvl(rsTmp("年龄"))
            Record.Item(mColP.婴儿).Value = Nvl(rsTmp("婴儿"))
            Record.Item(mColP.挂号单).Value = Nvl(rsTmp("挂号单"))
            Record.Item(mColP.门诊号).Value = Nvl(rsTmp("门诊号"))
            Record.Item(mColP.住院号).Value = Nvl(rsTmp("住院号"))
            Record.Item(mColP.出生日期).Value = Nvl(rsTmp("出生日期"))
            Record.Item(mColP.主页ID).Value = Nvl(rsTmp("主页ID"))
            Record.Item(mColP.标识号).Value = Nvl(rsTmp("标识号"))
            Record.Item(mColP.床号).Value = Nvl(rsTmp("床号"))
            Record.Item(mColP.病人科室).Value = Nvl(rsTmp("病人科室"))
            Record.Item(mColP.申请人).Value = Nvl(rsTmp("申请人"))
            Record.Item(mColP.申请科室ID).Value = Nvl(rsTmp("申请科室ID"))
            Record.Item(mColP.病人来源).Value = Nvl(rsTmp("病人来源"))
        rsTmp.MoveNext
    Loop
    Me.rptListSource.Populate
    Me.rptListSelect.Populate
    
    Exit Sub
errH:
    If errcenter() = 1 Then
        Resume
    End If
    Call saveerrlog
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strDate As String
    Dim Column As ReportColumn
    Dim lngKey As Long
    
    DTP.Value = Now
    
    '读入下拉框内容
    '==科室 和 执行科室
    lngKey = zldatabase.GetPara("frmAddPatient_选择科室", 100, 1208, -1)
    gstrSql = "select id,编码,名称 from 部门表 a , 部门性质说明 b" & vbNewLine & _
              "where a.id = b.部门id and 工作性质 = '检验'" & vbNewLine & _
              "order by 编码"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption)
    
    With cbo科室
        .AddItem "所有科室"
        .ItemData(.NewIndex) = 0
        Do Until rsTmp.EOF
            .AddItem Nvl(rsTmp("名称"))
            .ItemData(.NewIndex) = rsTmp("ID")
            If lngKey = rsTmp("ID") Then
                .ListIndex = .NewIndex
            End If
            rsTmp.MoveNext
        Loop
    End With
    
    lngKey = zldatabase.GetPara("frmAddPatient_执行科室", 100, 1208, -1)
    rsTmp.MoveFirst
    With cbo执行科室
        Do Until rsTmp.EOF
            .AddItem Nvl(rsTmp("名称"))
            .ItemData(.NewIndex) = rsTmp("ID")
            If lngKey = rsTmp("ID") Then
                .ListIndex = .NewIndex
            End If
            rsTmp.MoveNext
        Loop
    End With
        
    '==日期
    strDate = zldatabase.Currentdate
    Me.dtpBegin = strDate
    Me.dtpEnd = strDate
    
    '==类别
    lngKey = zldatabase.GetPara("frmAddPatient_选择类别", 100, 1208, -1)
    gstrSql = "select 编码,名称 from 诊疗检验类型 "
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With cbo类别
        .AddItem "所有类别"
        .ItemData(.NewIndex) = 0
        Do Until rsTmp.EOF
            .AddItem Nvl(rsTmp("名称"))
            .ItemData(.NewIndex) = rsTmp("编码")
            If lngKey = rsTmp("编码") Then
                .ListIndex = .NewIndex
            End If
            rsTmp.MoveNext
        Loop
    End With
    
    '==开单科室
    lngKey = zldatabase.GetPara("frmAddPatient_开单科室", 100, 1208, -1)
    gstrSql = "select distinct a.id,a.编码,a.名称 from 部门表 a , 部门性质说明 b " & _
                 " where a.id = b.部门id and b.工作性质 in ('检验','护理','临床')  order by a.编码"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With cbo开单科室
        Do Until rsTmp.EOF
            .AddItem Nvl(rsTmp("编码")) & "-" & Nvl(rsTmp("名称"))
            .ItemData(.NewIndex) = rsTmp("ID")
            If lngKey = rsTmp("ID") Then
                .ListIndex = .NewIndex
            End If
            rsTmp.MoveNext
        Loop
    End With
                
    '==检验仪器
    gstrSql = "select ID,编码,名称  from 检验仪器 "
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With cbo检验仪器
        .AddItem "手工"
        .ItemData(.NewIndex) = -1
        Do Until rsTmp.EOF
            .AddItem Nvl(rsTmp("编码")) & "-" & Nvl(rsTmp("名称"))
            .ItemData(.NewIndex) = rsTmp("ID")
            rsTmp.MoveNext
        Loop
    End With
    
    '==检验类别
    gstrSql = "select 编码,名称 from 诊疗检验类型 order by 编码"
    Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With cbo类别
        .Clear
        .AddItem "所有类别"
        .ItemData(.NewIndex) = 0
        Do Until rsTmp.EOF
            .AddItem Nvl(rsTmp("编码")) & "-" & Nvl(rsTmp("名称"))
            .ItemData(.NewIndex) = Nvl(rsTmp("编码"))
            rsTmp.MoveNext
        Loop
    End With
    
    With Me.rptListSource.Columns
        rptListSource.AllowColumnRemove = False
        rptListSource.ShowItemsInGroups = False
        
        With rptListSource.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
'        rptListSource.SetImageList ImgList
        Set Column = .Add(mColP.病人ID, "病人ID", 75, True): Column.Visible = False
        Set Column = .Add(mColP.标本序号, "标本序号", 50, True)
        Set Column = .Add(mColP.姓名, "姓名", 75, True)
        Set Column = .Add(mColP.性别, "性别", 45, True)
        Set Column = .Add(mColP.年龄, "年龄", 60, True)
        Set Column = .Add(mColP.婴儿, "婴儿", 45, True)
        Set Column = .Add(mColP.挂号单, "挂号单", 75, True): Column.Visible = False
        Set Column = .Add(mColP.门诊号, "门诊号", 75, True): Column.Visible = False
        Set Column = .Add(mColP.住院号, "住院号", 75, True): Column.Visible = False
        Set Column = .Add(mColP.出生日期, "出生日期", 75, True): Column.Visible = False
        Set Column = .Add(mColP.主页ID, "主页ID", 75, True): Column.Visible = False
        Set Column = .Add(mColP.标识号, "标识号", 75, True): Column.Visible = False
        Set Column = .Add(mColP.床号, "床号", 75, True): Column.Visible = False
        Set Column = .Add(mColP.病人科室, "病人科室", 75, True): Column.Visible = False
        Set Column = .Add(mColP.申请人, "申请人", 75, True): Column.Visible = False
        Set Column = .Add(mColP.申请科室ID, "申请科室ID", 75, True): Column.Visible = False
        Set Column = .Add(mColP.病人来源, "病人来源", 75, True): Column.Visible = False
    End With
    
    With Me.rptListSelect.Columns
        rptListSelect.AllowColumnRemove = False
        rptListSelect.ShowItemsInGroups = False
        
        With rptListSelect.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
'        rptListSource.SetImageList ImgList
        Set Column = .Add(mColP.病人ID, "病人ID", 75, True): Column.Visible = False
        Set Column = .Add(mColP.标本序号, "标本序号", 50, True)
        Set Column = .Add(mColP.姓名, "姓名", 75, True)
        Set Column = .Add(mColP.性别, "性别", 45, True)
        Set Column = .Add(mColP.年龄, "年龄", 60, True)
        Set Column = .Add(mColP.婴儿, "婴儿", 45, True)
        Set Column = .Add(mColP.挂号单, "挂号单", 75, True): Column.Visible = False
        Set Column = .Add(mColP.门诊号, "门诊号", 75, True): Column.Visible = False
        Set Column = .Add(mColP.住院号, "住院号", 75, True): Column.Visible = False
        Set Column = .Add(mColP.出生日期, "出生日期", 75, True): Column.Visible = False
        Set Column = .Add(mColP.主页ID, "主页ID", 75, True): Column.Visible = False
        Set Column = .Add(mColP.标识号, "标识号", 75, True): Column.Visible = False
        Set Column = .Add(mColP.床号, "床号", 75, True): Column.Visible = False
        Set Column = .Add(mColP.病人科室, "病人科室", 75, True): Column.Visible = False
        Set Column = .Add(mColP.申请人, "申请人", 75, True): Column.Visible = False
        Set Column = .Add(mColP.申请科室ID, "申请科室ID", 75, True): Column.Visible = False
        Set Column = .Add(mColP.病人来源, "病人来源", 75, True): Column.Visible = False
    End With
    
    With Me.rptItemSource.Columns
        rptItemSource.AllowColumnRemove = False
        rptItemSource.ShowItemsInGroups = False
        
        With rptItemSource.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
'        rptListSource.SetImageList ImgList
        Set Column = .Add(mColI.ID, "ID", 75, True): Column.Visible = False
        Set Column = .Add(mColI.编码, "编码", 75, True)
        Set Column = .Add(mColI.名称, "名称", 100, True)
        Set Column = .Add(mColI.标本, "标本", 60, True)
        Set Column = .Add(mColI.项目类别, "项目类别", 60, True): Column.Visible = False
    End With
    
    With Me.rptItemSelect.Columns
        rptItemSelect.AllowColumnRemove = False
        rptItemSelect.ShowItemsInGroups = False
        
        With rptItemSelect.PaintManager
            .ColumnStyle = xtpColumnShaded
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
'        rptListSource.SetImageList ImgList
        Set Column = .Add(mColI.ID, "ID", 75, True): Column.Visible = False
        Set Column = .Add(mColI.编码, "编码", 75, True)
        Set Column = .Add(mColI.名称, "名称", 75, True)
        Set Column = .Add(mColI.标本, "标本", 60, True)
        Set Column = .Add(mColI.项目类别, "项目类别", 60, True): Column.Visible = False
    End With
End Sub
Public Sub ShowMe(objfrm As Object, lngMachineID As Long, lngExecDeptID As Long)
    mlngMachineID = lngMachineID
    mlngExecDeptID = lngExecDeptID
    Me.Show vbModal, objfrm
End Sub

Private Function chkRepeat(lngID As Long, rptRows As ReportRows, strItemIndex As Integer) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '功能           检验重复
    '参数
    '返回           True=重复  False=不重复
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim intLoop As Integer
    For intLoop = 0 To rptRows.Count - 1
        If lngID = rptRows(intLoop).Record(strItemIndex).Value Then
            chkRepeat = True
            Exit Function
        End If
    Next
    chkrepeate = False
End Function
Private Sub MoveItem(intType As Integer, intMove As Integer, AllItem As Boolean)
    '功能               移动项目到指定列表框中
    '参数               intType 1=病人 2=项目
    '                   intMove 1= Left 2=Right
    '                   AllItem True=所有 False=单行
    Dim Record As ReportRecord
    Dim intLoop As Integer, intRow As Integer
    Dim lngKey As Long
    
    '病人
    If intType = 1 Then
        If intMove = 1 Then
            'Left
            If AllItem = False Then
                If Not Me.rptListSelect.FocusedRow Is Nothing Then
                    lngKey = Me.rptListSelect.FocusedRow.Record(mColP.病人ID).Value
                    If chkRepeat(lngKey, Me.rptListSource.Rows, CInt(mColP.病人ID)) = False Then
                        Set Record = Me.rptListSource.Records.Add
                        For intLoop = 0 To Me.rptListSelect.Columns.Count
                            Record.AddItem ""
                            Record.Item(intLoop).Value = Me.rptListSelect.FocusedRow.Record(intLoop).Value
                        Next
                        Me.rptListSelect.Records.RemoveAt (Me.rptListSelect.FocusedRow.Record.Index)
                    End If
                End If
            Else
                For intLoop = 0 To Me.rptListSelect.Records.Count - 1
                    lngKey = Me.rptListSelect.Records(intLoop).Item(mColP.病人ID).Value
                    If chkRepeat(lngKey, Me.rptListSource.Rows, CInt(mColP.病人ID)) = False Then
                        Set Record = Me.rptListSource.Records.Add
                        For intRow = 0 To Me.rptListSelect.Columns.Count
                            Record.AddItem ""
                            Record.Item(intRow).Value = Me.rptListSelect.Records(intLoop).Item(intRow).Value
                         Next
                    End If
                Next
                Me.rptListSelect.Records.DeleteAll
            End If
            Me.rptListSelect.Populate: Me.rptListSource.Populate
        Else
            'Right
            If AllItem = False Then
                If Not Me.rptListSource.FocusedRow Is Nothing Then
                    lngKey = Me.rptListSource.FocusedRow.Record(mColP.病人ID).Value
                    If chkRepeat(lngKey, Me.rptListSelect.Rows, CInt(mColP.病人ID)) = False Then
                        Set Record = Me.rptListSelect.Records.Add
                        For intLoop = 0 To Me.rptListSource.Columns.Count
                            Record.AddItem ""
                            Record.Item(intLoop).Value = Me.rptListSource.FocusedRow.Record(intLoop).Value
                        Next
                        Me.rptListSource.Records.RemoveAt (Me.rptListSource.FocusedRow.Record.Index)
                        
                    End If
                End If
            Else
                For intLoop = 0 To Me.rptListSource.Records.Count - 1
                    lngKey = Me.rptListSource.Records(intLoop).Item(mColP.病人ID).Value
                    If chkRepeat(lngKey, Me.rptListSelect.Rows, CInt(mColP.病人ID)) = False Then
                        Set Record = Me.rptListSelect.Records.Add
                        For intRow = 0 To Me.rptListSource.Columns.Count
                            Record.AddItem ""
                            Record.Item(intRow).Value = Me.rptListSource.Records(intLoop).Item(intRow).Value
                         Next
                    End If
                Next
                Me.rptListSource.Records.DeleteAll
            End If
            Me.rptListSelect.Populate: Me.rptListSource.Populate
        End If
    End If
    
    '项目
    If intType = 2 Then
        If intMove = 1 Then
            'Left
            If AllItem = False Then
                If Not Me.rptItemSelect.FocusedRow Is Nothing Then
                    lngKey = Me.rptItemSelect.FocusedRow.Record(mColI.ID).Value
                    If chkRepeat(lngKey, Me.rptItemSource.Rows, CInt(mColI.ID)) = False Then
                        Set Record = Me.rptItemSource.Records.Add
                        For intLoop = 0 To Me.rptItemSelect.Columns.Count
                            Record.AddItem ""
                            Record.Item(intLoop).Value = Me.rptItemSelect.FocusedRow.Record(intLoop).Value
                        Next
                        Me.rptItemSelect.Records.RemoveAt (Me.rptItemSelect.FocusedRow.Record.Index)
                    End If
                End If
            Else
                For intLoop = 0 To Me.rptItemSelect.Records.Count - 1
                    lngKey = Me.rptItemSelect.Records(intLoop).Item(mColI.ID).Value
                    If chkRepeat(lngKey, Me.rptItemSource.Rows, CInt(mColI.ID)) = False Then
                        Set Record = Me.rptItemSource.Records.Add
                        For intRow = 0 To Me.rptItemSelect.Columns.Count
                            Record.AddItem ""
                            Record.Item(intRow).Value = Me.rptItemSelect.Records(intLoop).Item(intRow).Value
                         Next
                    End If
                Next
                Me.rptItemSelect.Records.DeleteAll
            End If
            Me.rptItemSelect.Populate: Me.rptItemSource.Populate
        Else
            'Right
            If AllItem = False Then
                If Not Me.rptItemSource.FocusedRow Is Nothing Then
                    lngKey = Me.rptItemSource.FocusedRow.Record(mColI.ID).Value
                    If chkRepeat(lngKey, Me.rptItemSelect.Rows, CInt(mColI.ID)) = False Then
                        Set Record = Me.rptItemSelect.Records.Add
                        For intLoop = 0 To Me.rptItemSource.Columns.Count
                            Record.AddItem ""
                            Record.Item(intLoop).Value = Me.rptItemSource.FocusedRow.Record(intLoop).Value
                        Next
                        Me.rptItemSource.Records.RemoveAt (Me.rptItemSource.FocusedRow.Record.Index)
                        
                    End If
                End If
            Else
                For intLoop = 0 To Me.rptItemSource.Records.Count - 1
                    lngKey = Me.rptItemSource.Records(intLoop).Item(mColI.ID).Value
                    If chkRepeat(lngKey, Me.rptItemSelect.Rows, CInt(mColI.ID)) = False Then
                        Set Record = Me.rptItemSelect.Records.Add
                        For intRow = 0 To Me.rptItemSource.Columns.Count
                            Record.AddItem ""
                            Record.Item(intRow).Value = Me.rptItemSource.Records(intLoop).Item(intRow).Value
                         Next
                    End If
                Next
                Me.rptItemSource.Records.DeleteAll
            End If
            Me.rptItemSelect.Populate: Me.rptItemSource.Populate
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    zldatabase.SetPara "frmAddPatient_选择科室", Me.cbo科室.ItemData(Me.cbo科室.ListIndex), 100, 1208
    zldatabase.SetPara "frmAddPatient_选择仪器", Me.cbo仪器.ItemData(Me.cbo仪器.ListIndex), 100, 1208
    zldatabase.SetPara "frmAddPatient_选择类别", Me.cbo类别.ItemData(Me.cbo类别.ListIndex), 100, 1208
    zldatabase.SetPara "frmAddPatient_开单科室", Me.cbo开单科室.ItemData(Me.cbo开单科室.ListIndex), 100, 1208
    zldatabase.SetPara "frmAddPatient_开单医生", Me.cbo开单医生.ItemData(Me.cbo开单医生.ListIndex), 100, 1208
    zldatabase.SetPara "frmAddPatient_执行科室", Me.cbo执行科室.ItemData(Me.cbo执行科室.ListIndex), 100, 1208
    zldatabase.SetPara "frmAddPatient_检验仪器", Me.cbo检验仪器.ItemData(Me.cbo检验仪器.ListIndex), 100, 1208
End Sub

Private Sub rptItemSelect_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call cmdItemLeft_Click
End Sub

Private Sub rptItemSource_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call cmdItemRight_Click
End Sub

Private Sub rptListSelect_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call cmdPatientLeft_Click
End Sub

Private Sub rptListSource_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call cmdPatientRight_Click
End Sub

Private Sub SaveData()
    '功能                   保存数据
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL() As String
    Dim intLoop As Integer, intItem As Integer
    Dim lngTmpID As Long, lngAdviceID As Long
    Dim lngMaxSeq As Long, iSendSeq   As Integer
    Dim intPatientType As Integer, lngPatientID As Long, intPatientPage As Integer
    Dim intBaby As Integer, lngExecDept As Long, lngPatientDept As Long, lngSendNO As Long
    Dim strDate As String, strDoctor As String, strNO As String, strAdviceText As String
    Dim strSample As String, strSampleNO As String, strName As String, strSex As String
    Dim lngCapID As Long, lngSampleID As Long, strlngID As String
    Dim strAge As String, strBed As String, strItemIDs As String, strItemResults As String
    Dim intMicrobe As Integer
    Dim blnJumpRepeat As Boolean
    Dim blnGetNO As Boolean
    Dim intNo As Integer
    Dim lngDeviceID As Long
    Dim strTmpDate As String
    Dim blnNew As Boolean
    
    Me.MousePointer = 11
    zlCommFun.ShowFlash "正在生成标本请稍等。。。"
    
    ReDim strSQL(1 To 1)
    blnJumpRepeat = False
    
    With Me.cbo执行科室
        If .Text <> "" Then
            lngExecDept = .ItemData(.ListIndex)
        End If
    End With
    
    With Me.cbo开单科室
        If .Text <> "" Then
            lngPatientDetp = .ItemData(.ListIndex)
        End If
    End With
    
    With Me.cbo开单医生
        If .Text <> "" Then
            strDoctor = Mid(.Text, InStr(.Text, "-") + 1)
        End If
    End With
    
    With Me.cbo检验仪器
        If .Text <> "" Then
            lngDeviceID = .ItemData(.ListIndex)
        End If
    End With
    
    For intLoop = 0 To Me.rptListSelect.Records.Count - 1
        
        '==============================================新增病人医嘱==================================================
        strDate = "To_Date('" & Format(DTP, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        With Me.rptListSelect
            intPatientType = Val(.Records(intLoop).Item(mColP.病人来源).Value)
            intPatientType = IIf(intPatientType = 0, 3, intPatientType)
            lngPatientID = Val(.Records(intLoop).Item(mColP.病人ID).Value)
            intPatientPage = Val(.Records(intLoop).Item(mColP.主页ID).Value)
            intBaby = Val(.Records(intLoop).Item(mColP.婴儿).Value)
            strName = .Records(intLoop).Item(mColP.姓名).Value
            strSex = .Records(intLoop).Item(mColP.性别).Value
            strAge = .Records(intLoop).Item(mColP.年龄).Value
            strlngID = Val(.Records(intLoop).Item(mColP.标识号).Value)
            strBed = .Records(intLoop).Item(mColP.床号).Value
        End With

        lngAdviceID = zldatabase.GetNextId("病人医嘱记录")              '相关ID
        lngSendNO = zldatabase.GetNextNo(10)                            '发送号
        strNO = zldatabase.GetNextNo(IIf(PatientType = 2, 14, 13))      '单据号
        
        '=======取的单据号会重复，不知道原因先处理为如果发现重复重新取一次。临时处理
        intNo = 0
        blnGetNO = False
        Do Until blnGetNO = False
            intNo = intNo + 1
            gstrSql = "select " & gConst_病人医嘱发送_列名 & " from 病人医嘱发送 a where no = [1] "
            Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, strNO)
            If rsTmp.EOF = False Then blnGetNO = True
            If intNo >= 10 Then blnGetNO = True
        Loop
        '======================================================================
        '得到最大医嘱序号
        gstrSql = "select max(序号) as 序号 from 病人医嘱记录 where 病人id = [1] "
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.ControlBox, lngPatientID)
        If rsTmp.EOF = False Then
            lngMaxSeq = Val(Nvl(rsTmp("序号"), 0))
        Else
            lngMaxSeq = 0
        End If

        iSendSeq = 1
        For intItem = 0 To Me.rptItemSelect.Records.Count - 1
            '检验项目医嘱
            lngTmpID = zldatabase.GetNextId("病人医嘱记录")
            lngMaxSeq = lngMaxSeq + 1
            If intItem = 0 Then
                strAdviceText = Replace(rptItemSelect.Records(intItem).Item(mColI.名称).Value, "'", "''") & "(" & _
                    rptItemSelect.Records(intItem).Item(mColI.标本).Value & ")"
                strSample = rptItemSelect.Records(intItem).Item(mColI.标本).Value
                intMicrobe = rptItemSelect.Records(intItem).Item(mColI.项目类别).Value
            Else
                strAdviceText = Replace(rptItemSelect.Records(intItem).Item(mColI.名称).Value, "'", "''") & "," & strAdviceText
            End If
            strItemIDs = strItemIDs & "," & rptItemSelect.Records(intItem).Item(mColI.ID).Value
            strSQL(ReDimArray(strSQL)) = "ZL_病人医嘱记录_Insert(" & lngTmpID & "," & lngAdviceID & "," & lngMaxSeq & "," & intPatientType & _
                "," & lngPatientID & "," & IIf(intPatientPage = 0, "NULL", intPatientPage) & "," & IIf(intBaby = 0, "NULL", intBaby) & _
                ",1,1,'C'," & rptItemSelect.Records(intItem).Item(mColI.ID).Value & ",NULL,NULL,NULL,NULL,'" & _
                Replace(rptItemSelect.Records(intItem).Item(mColI.名称).Value, "'", "''") & "',NULL,'" & _
                rptItemSelect.Records(intItem).Item(mColI.标本).Value & "','一次性',NULL,NULL,NULL,NULL,0," & lngExecDept & ",4,0," & _
                strDate & ",NULL," & lngPatientDetp & "," & lngPatientDetp & ",'" & strDoctor & "'," & strDate & ")"

            iSendSeq = iSendSeq + 1
            strSQL(ReDimArray(strSQL)) = "ZL_病人医嘱发送_Insert(" & lngTmpID & "," & lngSendNO & "," & IIf(intPatientType = 2, 2, 1) & _
            ",'" & strNO & "'," & iSendSeq & ",NULL,NULL,NULL,Sysdate+1/(24*3600),0," & lngExecDept & ",0,0)"
        Next

        '采集方式医嘱
        gstrSql = "Select 用法id From 诊疗项目目录 A, 诊疗用法用量 B Where A.ID = B.项目id and a.id = [1] and b.性质 = 1 "
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, CLng(Me.rptItemSelect.Records(0).Item(mColI.ID).Value))
        '没有找到采集方式时退出
        If rsTmp.EOF = True Then MsgBox "没有定义标本采集方式，请到诊疗项目管理中设置。", vbInformation, gstrSysName: Exit Sub
        lngCapID = rsTmp("用法ID")
        lngMaxSeq = lngMaxSeq + 1

        strSQL(ReDimArray(strSQL)) = "ZL_病人医嘱记录_Insert(" & lngAdviceID & ",NULL," & lngMaxSeq & "," & intPatientType & _
                "," & lngPatientID & "," & IIf(intPatientPage = 0, "NULL", intPatientPage) & "," & IIf(intBaby = 0, "NULL", intBaby) & _
                ",1,1,'E'," & lngCapID & ",NULL,NULL,NULL,NULL,'" & _
                strAdviceText & "',NULL,'" & _
                strSample & "','一次性',NULL,NULL,NULL,NULL,2," & lngExecDept & ",3,0," & _
                strDate & ",NULL," & lngPatientDetp & "," & lngPatientDetp & ",'" & strDoctor & "'," & strDate & ")"

        iSendSeq = iSendSeq + 1
        strSQL(ReDimArray(strSQL)) = "ZL_病人医嘱发送_Insert(" & lngAdviceID & "," & lngSendNO & "," & IIf(intPatientType = 2, 2, 1) & _
            ",'" & strNO & "'," & iSendSeq & ",NULL,NULL,NULL,Sysdate+1/(24*3600),0," & lngExecDept & ",0,1)"
        '====================================================================================================================================

        '=================================================================标本信息===========================================================
        '标本号
        strSampleNO = Val(Me.txt生成标本号.Text) + intLoop
        If lngDeviceID = -1 Then
            strSampleNO = TransSampleNO(Val(txt标本批号.Text) & "-" & Val(strSampleNO))
        End If
        
        gstrSql = "Select Id,标本序号" & vbNewLine & _
                    "From 检验标本记录" & vbNewLine & _
                    "Where 核收时间 Between To_Date(To_Char([1], 'yyyy-MM-dd') || ' 00:00:00', 'yyyy-MM-dd HH24:mi:ss') And" & vbNewLine & _
                    "           To_Date(To_Char([1], 'yyyy-MM-dd') || ' 23:59:59', 'yyyy-MM-dd HH24:mi:ss') And 标本序号 = [2] And" & vbNewLine & _
                    "           Nvl(仪器id, 0) = Nvl([3], 0) " & IIf(blnEmergency = 1, " and nvl(标本类型,0) = [4]", "")
        strTmpDate = Format(zldatabase.Currentdate, "yyyy-mm-dd")
        Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, CDate(strTmpDate), strSampleNO, _
        Val(IIf(lngDeviceID = -1, 0, lngDeviceID)), 0)
        If rsTmp.EOF = False Then
            If blnJumpRepeat = False Then
                If MsgBox("发现有标本号重复是否续继？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
                    Exit Sub
                End If
                blnJumpRepeat = True
            End If
            lngSampleID = rsTmp("ID")
            blnNew = False
        Else
            lngSampleID = zldatabase.GetNextId("检验标本记录")
            blnNew = True
        End If
        
        
        strSQL(ReDimArray(strSQL)) = "ZL_检验标本记录_标本核收(" & lngSampleID & "," & lngAdviceID & ",'" & lngAdviceID & "',0,'" & strSampleNO & "'," & _
            strDate & ",NULL," & IIf(lngDeviceID = -1, "NULL", lngDeviceID) & "," & strDate & ",NULL,'" & UserInfo.姓名 & "'," & _
            strDate & "," & IIf(intMicrobe = 2, 1, "Null") & ",0,NULL,'" & strName & "','" & strSex & "','" & strAge & "','" & lngSendNO & "','" & strSample & "'," & _
            lngPatientDetp & ",'" & strDoctor & "'," & IIf(strlngID = 0, "NULL", strlngID) & ",'" & strBed & "'," & lngPatientDetp & ",'" & _
            strAdviceText & "',1," & lngPatientID & "," & cbo执行科室.ItemData(cbo执行科室.ListIndex) & ",'" & UserInfo.编号 & "','" & UserInfo.姓名 & "')"
            
        If blnNew = True Then
            If intMicrobe = 2 Then
                gstrSql = "Select Id, 原始结果, 检验结果, 结果标志, 结果参考, Rownum As 排列序号, 诊疗项目id" & vbNewLine & _
                        "From (Select d.细菌id As Id, '' As 原始结果, c.默认结果 As 检验结果, '' As 结果标志, '' As 结果参考, d.排列序号," & vbNewLine & _
                        "                           a.Id As 诊疗项目id" & vbNewLine & _
                        "            From 诊疗项目目录 a, 检验报告项目 d, 检验细菌 c" & vbNewLine & _
                        "            Where a.Id = d.诊疗项目id And d.细菌id = c.Id And a.Id In (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & vbNewLine & _
                        "            Order By a.编码, d.排列序号)"
            Else
                gstrSql = "Select Id, 原始结果, 检验结果, 结果标志, Rownum As 排列序号, 诊疗项目id,结果参考" & vbNewLine & _
                            "From (Select d.报告项目id As Id, '' As 原始结果, Decode(c.结果类型, 3, Nvl(c.默认值, '-'), 2, c.默认值, '') As 检验结果," & vbNewLine & _
                            "                           '' As 结果标志, d.排列序号, a.Id As 诊疗项目id," & vbNewLine & _
                            "                           Trim(Replace(Replace(' ' || Zlgetreference(d.报告项目id, a.标本部位, 0, Null), ' .', '0.'), '～.', '～0.')) As 结果参考" & vbNewLine & _
                            "            From 诊疗项目目录 a, 检验报告项目 d, 检验项目 c" & vbNewLine & _
                            "            Where a.Id = d.诊疗项目id And d.报告项目id = c.诊治项目id And d.细菌id Is Null And c.项目类别 <> 2 And" & vbNewLine & _
                            "                        a.Id In (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & vbNewLine & _
                            "            Order By a.编码, d.排列序号) "
            End If
            Set rsTmp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Mid(strItemIDs, 2))
            strItemResults = ""
            Do While Not rsTmp.EOF
                strItemResults = strItemResults & "|" & lngAdviceID & "^" & Nvl(rsTmp("ID")) & "^" & Nvl(rsTmp("检验结果")) & _
                    "^" & Nvl(rsTmp("结果标志")) & "^" & Nvl(rsTmp("结果参考")) & "^" & Nvl(rsTmp("诊疗项目ID")) & _
                    "^" & Nvl(rsTmp("排列序号"))
                rsTmp.MoveNext
            Loop
            
            strSQL(ReDimArray(strSQL)) = "Zl_检验普通结果_Write(" & lngSampleID & "," & IIf(lngDeviceID = -1, "NULL", lngDeviceID) & ",'" & _
                Mid(strItemResults, 2) & "',0," & IIf(intMicrobe = 2, 1, 0) & ")"
            strSQL(ReDimArray(strSQL)) = "Zl_重新计算结果_Cale(" & lngSampleID & ")"
        End If
        '====================================================================================================================================
    Next
    '开始执行
    On Error GoTo errH
'    gcnOracle.BeginTrans
    
    
    For intLoop = 1 To UBound(strSQL)
        Debug.Print strSQL(intLoop) & vbCrLf
        If strSQL(intLoop) <> "" Then zldatabase.ExecuteProcedure strSQL(intLoop), Me.Caption
    Next
    zlCommFun.StopFlash
'    gcnOracle.CommitTrans
    MsgBox "生成标本完成!", vbInformation, Me.Caption
    Me.MousePointer = 0
    Exit Sub
errH:
    zlCommFun.StopFlash
    Me.MousePointer = 0
'    gcnOracle.RollbackTrans
    If errcenter() = 1 Then Resume
    Call saveerrlog
End Sub
  
Private Sub txt标本号_GotFocus()
    txt标本号.SelStart = 0
    txt标本号.SelLength = Len(txt标本号)
End Sub

Private Sub txt标本号_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call cmd过滤_Click: Me.txt标本号.SelStart = 0: Me.txt标本号.SelLength = Len(Me.txt标本号)
    If InStr("1234567890~,-" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txt标本批号_GotFocus()
    txt标本批号.SelStart = 0
    txt标本批号.SelLength = Len(txt标本批号)
End Sub

Private Sub txt标本批号_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txt查找_GotFocus()
    Me.txt查找.SelStart = 0
    Me.txt查找.SelLength = Len(Me.txt查找.Text)
End Sub

Private Sub txt查找_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call cmd查找_Click
    End If
End Sub
Private Function chkSaveData() As Boolean
    '功能           保存检查
    
    If rptListSelect.Records.Count = 0 Then
        MsgBox "病人信息为空，请选择病人信息!", vbInformation, Me.Caption
        Exit Function
    End If
    
    If rptItemSelect.Records.Count = 0 Then
        MsgBox "项目信息为空，请选择项目信息!", vbInformation, Me.Caption
        Exit Function
    End If
    
    If cbo开单科室.Text = "" Then
        MsgBox "请选择开单科室!", vbInformation, Me.Caption
        Me.cbo开单科室.SetFocus
        Exit Function
    End If
    
    If cbo开单医生.Text = "" Then
        MsgBox "请选择开单医生!", vbInformation, Me.Caption
        Me.cbo开单医生.SetFocus
        Exit Function
    End If
    
    If cbo执行科室.Text = "" Then
        MsgBox "请选择执行科室!", vbInformation, Me.Caption
        Me.cbo执行科室.SetFocus
        Exit Function
    End If
    
    If cbo检验仪器.Text = "" Then
        MsgBox "请选择检验仪器!", vbInformation, Me.Caption
        Me.cbo检验仪器.SetFocus
        Exit Function
    End If
    
    If cbo检验仪器.ItemData(cbo检验仪器.ListIndex) = -1 Then
        If Trim(txt标本批号) = "" Or Trim(txt生成标本号) = "" Then
            MsgBox "手工项目必须输入批号和标本号!", vbInformation, Me.Caption
            Me.txt批号.SetFocus
            Exit Function
        End If
    Else
        If Trim(txt生成标本号) = "" Then
            MsgBox "请输入生成的标本号开始号!", vbInformation, Me.Caption
            Me.txt生成标本号.SetFocus
            Exit Function
        End If
    End If
    
    chkSaveData = True
    
    
End Function

Private Sub txt批号_GotFocus()
    txt批号.SelStart = 0
    txt批号.SelLength = Len(txt批号.Text)
End Sub

Private Sub txt批号_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub

Private Sub txt生成标本号_GotFocus()
    txt生成标本号.SelStart = 0
    txt生成标本号.SelLength = Len(txt生成标本号)
End Sub

Private Sub txt生成标本号_KeyPress(KeyAscii As Integer)
    If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
End Sub
