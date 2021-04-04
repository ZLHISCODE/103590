VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmLabSampleRegisterRefuse 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "标本拒收"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboRefuse 
      Height          =   300
      Left            =   1230
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2280
      Width           =   6345
   End
   Begin VB.PictureBox PicRecord 
      Height          =   1935
      Left            =   60
      ScaleHeight     =   1875
      ScaleWidth      =   7455
      TabIndex        =   6
      Top             =   300
      Width           =   7515
      Begin XtremeReportControl.ReportControl rptAlist 
         Height          =   1335
         Left            =   360
         TabIndex        =   7
         Top             =   60
         Width           =   4665
         _Version        =   589884
         _ExtentX        =   8229
         _ExtentY        =   2355
         _StockProps     =   0
         AllowColumnRemove=   0   'False
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
   End
   Begin VB.CheckBox chkHide 
      Caption         =   "隐藏未选中的医嘱"
      Height          =   225
      Left            =   5850
      TabIndex        =   5
      Top             =   68
      Value           =   1  'Checked
      Width           =   1875
   End
   Begin VB.CommandButton cmdRefuse 
      Caption         =   "拒收(&F)"
      Default         =   -1  'True
      Height          =   350
      Left            =   4440
      TabIndex        =   4
      Top             =   4140
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6060
      TabIndex        =   3
      Top             =   4140
      Width           =   1100
   End
   Begin VB.TextBox txt拒收 
      Height          =   1335
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2610
      Width           =   7545
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   1500
      Top             =   4020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampleRegisterRefuse.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampleRegisterRefuse.frx":006C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampleRegisterRefuse.frx":0606
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLabSampleRegisterRefuse.frx":0BA0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "填写拒收理由"
      Height          =   180
      Left            =   90
      TabIndex        =   1
      Top             =   2340
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "选择拒收标本"
      Height          =   180
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   1080
   End
End
Attribute VB_Name = "frmLabSampleRegisterRefuse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum mAcol                                  '医嘱列表
    ID
    选择
    图标
    已执行
    采集方式
    医嘱内容
    条码
    执行科室
    开嘱医生
    开嘱时间
    发送人
    发送时间
    标本
    采样时间
    试管颜色
    合并医嘱
    试管编码
    采样人
    采血量
    试管名称
    紧急
    病人来源
    申请科室
    婴儿
    别名
    相关ID
    医嘱id
    病人ID
    姓名
    性别
    年龄
    标识号
    床号
    病人科室
    接收时间
    诊疗项目ID
    执行状态
End Enum
Dim mRecords As ReportRecords
Public Sub ShowMe(Objfrm As Object, Recordset As ReportRecords)
    Set mRecords = Recordset
    Me.Show vbModal, Objfrm
End Sub

Private Sub cboRefuse_Click()
    Me.txt拒收.Text = Mid(Me.cboRefuse.Text, InStr(Me.cboRefuse.Text, "-") + 1)
End Sub

Private Sub chkHide_Click()
    Dim intLoop As Integer
    With Me.rptAlist
        If Me.chkHide.Value = 1 Then
            For intLoop = 0 To .Records.Count - 1
                If .Records(intLoop).Item(mAcol.选择).Checked = True Then
                    .Records(intLoop).Visible = True
                Else
                    .Records(intLoop).Visible = False
                End If
            Next
        Else
            If .Records.Count > 0 Then
                .Records(intLoop).Visible = True
            End If
        End If
        .Populate
    End With
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRefuse_Click()
    Dim blnSelect As Boolean
    Dim intLoop As Integer
    
    
    With Me.rptAlist
        If .Rows.Count > 0 Then
            For intLoop = 0 To .Rows.Count - 1
                If .Rows(intLoop).Record(mAcol.选择).Checked = True Then
                    blnSelect = True
                    Exit For
                End If
            Next
        End If
    End With
    
    '没有选择拒收医嘱
    If blnSelect = False Then
        MsgBox "请选择一个医嘱才能进行拒收!", vbInformation, Me.Caption
        Exit Sub
    End If
    
    '必须填写拒收理由
    If Trim(Me.txt拒收) = "" Then
        MsgBox "请填写拒收理由！", vbInformation, Me.Caption
        Me.txt拒收.SetFocus
        Exit Sub
    End If
    
    '开始拒收
    On Error GoTo errH
    
    gcnOracle.BeginTrans
    With Me.rptAlist
        For intLoop = 0 To .Rows.Count - 1
            If .Rows(intLoop).Record(mAcol.选择).Checked = True Then
                gstrSql = "Zl_检验标本记录_标本拒收(" & _
                            .Rows(intLoop).Record(mAcol.医嘱id).Value & ",'" & _
                            Me.txt拒收.Text & "','" & UserInfo.姓名 & "')"
                zlDatabase.ExecuteProcedure gstrSql, Me.Caption
            End If
        Next
    End With
    gcnOracle.CommitTrans
    Unload Me
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Sub

Private Sub Form_Load()
    Dim intLoop As Integer
    Dim lngLoop As Long
    Dim Record As ReportRecord
    Dim rsTmp As New ADODB.Recordset
    
    With Me.rptAlist
        .Top = 0
        .Left = 0
        .Width = Me.PicRecord.ScaleWidth
        .Height = Me.PicRecord.ScaleHeight
    End With
    
    rptAlist.SetImageList ImgList
    With rptAlist.PaintManager
        .ColumnStyle = xtpColumnShaded
        .GridLineColor = RGB(225, 225, 225)
        .NoGroupByText = "拖动列标题到这里,按该列分组..."
        .NoItemsText = "没有可显示的项目..."
        .VerticalGridStyle = xtpGridSolid
        .HideSelection = True
    End With
    With Me.rptAlist.Columns
        Set Column = .Add(mAcol.ID, "ID", 0, False): Column.Visible = False
        Set Column = .Add(mAcol.选择, "Check", 18, False): Column.Icon = 0
        Set Column = .Add(mAcol.图标, "", 18, False): Column.Icon = 3
        Set Column = .Add(mAcol.采集方式, "采集方式", 75, True)
        Set Column = .Add(mAcol.标本, "标本", 55, True)
        Set Column = .Add(mAcol.医嘱内容, "医嘱内容", 75, True)
        Set Column = .Add(mAcol.条码, "条码", 75, True)
        Set Column = .Add(mAcol.执行科室, "执行科室", 75, True)
        Set Column = .Add(mAcol.开嘱医生, "开嘱医生", 75, True)
        Set Column = .Add(mAcol.开嘱时间, "开嘱时间", 75, True)
        Set Column = .Add(mAcol.发送人, "发送人", 65, True)
        Set Column = .Add(mAcol.发送时间, "发送时间", 75, True)
        Set Column = .Add(mAcol.采样时间, "采样时间", 75, True)
        Set Column = .Add(mAcol.接收时间, "接收时间", 75, True)
        Set Column = .Add(mAcol.试管颜色, "颜色编码", 18, True): Column.Visible = False
        Set Column = .Add(mAcol.试管编码, "试管编码", 18, True): Column.Visible = False
        Set Column = .Add(mAcol.采样人, "采样人", 60, True)
        Set Column = .Add(mAcol.采血量, "采血量", 60, True): Column.Visible = False
        Set Column = .Add(mAcol.试管名称, "试管名称", 60, True): Column.Visible = False
        Set Column = .Add(mAcol.紧急, "紧急", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.病人来源, "病人来源", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.申请科室, "申请科室", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.婴儿, "婴儿", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.别名, "别名", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.相关ID, "相关ID", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.病人ID, "病人ID", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.姓名, "姓名", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.性别, "性别", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.年龄, "年龄", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.标识号, "标识号", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.床号, "床号", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.病人科室, "病人科室", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.诊疗项目ID, "诊疗项目Id", 50, True): Column.Visible = False
        Set Column = .Add(mAcol.医嘱id, "医嘱ID", 50, True): Column.Visible = False
    End With
    
    For lngLoop = 0 To mRecords.Count - 1
        Set Record = Me.rptAlist.Records.Add
        For intLoop = 0 To Me.rptAlist.Columns.Count + 1
            Record.AddItem ""
        Next
                
        Record(mAcol.ID).Value = mRecords(lngLoop).Item(mAcol.ID).Value
        Record(mAcol.选择).HasCheckbox = True
        Record(mAcol.选择).Checked = mRecords(lngLoop).Item(mAcol.选择).Checked
        Record(mAcol.图标).BackColor = mRecords(lngLoop).Item(mAcol.图标).BackColor
        Record(mAcol.采集方式).Value = mRecords(lngLoop).Item(mAcol.采集方式).Value
        Record(mAcol.医嘱内容).Value = mRecords(lngLoop).Item(mAcol.医嘱内容).Value
        Record(mAcol.条码).Value = mRecords(lngLoop).Item(mAcol.条码).Value
        Record(mAcol.执行科室).Value = mRecords(lngLoop).Item(mAcol.执行科室).Value
        Record(mAcol.开嘱医生).Value = mRecords(lngLoop).Item(mAcol.开嘱医生).Value
        Record(mAcol.开嘱时间).Value = mRecords(lngLoop).Item(mAcol.开嘱时间).Value
        Record(mAcol.发送人).Value = mRecords(lngLoop).Item(mAcol.发送人).Value
        Record(mAcol.发送时间).Value = mRecords(lngLoop).Item(mAcol.发送时间).Value
        Record(mAcol.试管颜色).Value = mRecords(lngLoop).Item(mAcol.试管颜色).Value
        Record(mAcol.试管编码).Value = mRecords(lngLoop).Item(mAcol.试管编码).Value
        Record(mAcol.标本).Value = mRecords(lngLoop).Item(mAcol.标本).Value
        Record(mAcol.采样时间).Value = mRecords(lngLoop).Item(mAcol.采样时间).Value
        Record(mAcol.采样人).Value = mRecords(lngLoop).Item(mAcol.采样人).Value
        Record(mAcol.采血量).Value = mRecords(lngLoop).Item(mAcol.采血量).Value
        Record(mAcol.试管名称).Value = mRecords(lngLoop).Item(mAcol.试管名称).Value
        Record(mAcol.紧急).Value = mRecords(lngLoop).Item(mAcol.紧急).Value
        Record(mAcol.病人来源).Value = mRecords(lngLoop).Item(mAcol.病人来源).Value
        Record(mAcol.婴儿).Value = mRecords(lngLoop).Item(mAcol.婴儿).Value
        Record(mAcol.别名).Value = mRecords(lngLoop).Item(mAcol.别名).Value
        Record(mAcol.相关ID).Value = mRecords(lngLoop).Item(mAcol.相关ID).Value
        
        Record(mAcol.病人ID).Value = mRecords(lngLoop).Item(mAcol.病人ID).Value
        Record(mAcol.姓名).Value = mRecords(lngLoop).Item(mAcol.姓名).Value
        Record(mAcol.性别).Value = mRecords(lngLoop).Item(mAcol.性别).Value
        Record(mAcol.年龄).Value = mRecords(lngLoop).Item(mAcol.年龄).Value
        Record(mAcol.标识号).Value = mRecords(lngLoop).Item(mAcol.标识号).Value
        Record(mAcol.床号).Value = mRecords(lngLoop).Item(mAcol.床号).Value
        Record(mAcol.病人科室).Value = mRecords(lngLoop).Item(mAcol.病人科室).Value
        Record(mAcol.接收时间).Value = mRecords(lngLoop).Item(mAcol.接收时间).Value
        Record(mAcol.诊疗项目ID).Value = mRecords(lngLoop).Item(mAcol.诊疗项目ID).Value
        Record(mAcol.医嘱id).Value = mRecords(lngLoop).Item(mAcol.医嘱id).Value
        
        For intLoop = 0 To Me.rptAlist.Columns.Count + 1
            Record(intLoop).ForeColor = mRecords(lngLoop).Item(mAcol.试管颜色).Value
        Next
        
    Next
    Call chkHide_Click
    Me.rptAlist.Populate
    
    On Error GoTo errH:
    
    gstrSql = "select 编码,名称 from 检验拒收理由"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    Do While Not rsTmp.EOF
        With Me.cboRefuse
            .AddItem Nvl(rsTmp("编码")) & "-" & Nvl(rsTmp("名称"))
        End With
        rsTmp.MoveNext
    Loop
        
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Picture1_Resize()
    
End Sub

Private Sub rptAlist_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Call chkHide_Click
End Sub

