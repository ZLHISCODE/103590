VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmPriceGradeEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "价格等级"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6765
   Icon            =   "frmPriceGradeEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   6765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picApplyBack 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Index           =   1
      Left            =   2190
      ScaleHeight     =   3015
      ScaleWidth      =   1500
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1500
      Begin MSComctlLib.ListView lvwApply 
         Height          =   2055
         Index           =   1
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   3625
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "名称"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.PictureBox picApplyBack 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Index           =   0
      Left            =   1350
      ScaleHeight     =   3015
      ScaleWidth      =   1380
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1290
      Width           =   1380
      Begin MSComctlLib.ListView lvwApply 
         Height          =   2025
         Index           =   0
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   3572
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "名称"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin XtremeSuiteControls.TabControl tbPageGradeApply 
      Height          =   3045
      Left            =   60
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   990
      Width           =   6645
      _Version        =   589884
      _ExtentX        =   11721
      _ExtentY        =   5371
      _StockProps     =   64
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   270
      TabIndex        =   14
      Top             =   4320
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5250
      TabIndex        =   13
      Top             =   4320
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4080
      TabIndex        =   12
      Top             =   4320
      Width           =   1100
   End
   Begin VB.Frame frmPriceGradeBaseInfo 
      Caption         =   "基本信息"
      Height          =   855
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6645
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         Height          =   280
         Index           =   2
         Left            =   4950
         TabIndex        =   6
         Top             =   360
         Width           =   1605
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         Height          =   280
         Index           =   1
         Left            =   2250
         TabIndex        =   4
         Top             =   360
         Width           =   1995
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         Height          =   280
         Index           =   0
         Left            =   690
         TabIndex        =   2
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "简码"
         Height          =   180
         Index           =   2
         Left            =   4560
         TabIndex        =   5
         Top             =   420
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "名称"
         Height          =   180
         Index           =   1
         Left            =   1860
         TabIndex        =   3
         Top             =   420
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "编码"
         Height          =   180
         Index           =   0
         Left            =   300
         TabIndex        =   1
         Top             =   420
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmPriceGradeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOk As Boolean
Private Enum TxtIndex
    Txt_编码 = 0
    Txt_名称 = 1
    Txt_简码 = 2
End Enum
Private Enum TabPageIndex
    Pg_NodeList = 0 '院区
    Pg_PatientType = 1 '医疗付款方式
End Enum
Private Enum FunType
    Fun_Add = 0 '新增
    Fun_Update = 1 '调整
    Fun_Delete = 2 '删除
    Fun_View = 3 '查看
End Enum
Private mbytFun As FunType '0-新增,1-调整,2-删除,3-查看
Private mstr价格等级 As String

Private mblnChanged As Boolean
Private mblnFirst As Boolean
Private mblnLoading As Boolean

Public Function ShowMe(frmParent As Form, ByVal bytFun As Byte, _
    Optional ByVal strIn价格等级 As String, _
    Optional ByRef strOut价格等级 As String) As Boolean
    '程序入口
    '入参：
    '   frmParent 调用窗口对象
    '   bytFun 操作类型：0-新增,1-调整,2-删除,3-查看
    '   strIn价格等级 查看、调整、删除时传入价格等级名称
    '出参：
    '   strOut价格等级 新增时返回价格等级名称，用于调用者定位
    mbytFun = bytFun
    mstr价格等级 = IIF(mbytFun = Fun_Add, "-", strIn价格等级)
    
    On Error Resume Next
    mblnOk = False
    If CheckDepend() = False Then Exit Function
    Me.Show 1, frmParent
    ShowMe = mblnOk
    strOut价格等级 = IIF(mbytFun = Fun_Add, mstr价格等级, "")
End Function

Private Function CheckDepend() As Boolean
    '功能:数据加载前检查
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    Err = 0: On Error GoTo ErrHandler
    If Not (mbytFun = Fun_Update Or mbytFun = Fun_Delete) Then CheckDepend = True: Exit Function
    
    '已经停用的，不允许调整/删除
    strSQL = "Select 1 From 收费价格等级 Where 名称 = [1] And Nvl(撤档时间, To_Date('3000-01-01','yyyy-mm-dd')) < To_Date('3000-01-01','yyyy-mm-dd')"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查价格等级是否被停用", mstr价格等级)
    If Not rsTemp.EOF Then
        MsgBox "当前价格等级已停用，不允许" & IIF(mbytFun = Fun_Update, "调整", "删除") & "。", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If mbytFun = Fun_Delete Then
        '如果当前价格等级已经调价（即已经存在调价记录），则不允许删除。
        strSQL = "Select 1 From 收费价目 Where 价格等级 = [1] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查价格等级在收费价目中是否使用", mstr价格等级)
        If Not rsTemp.EOF Then
            MsgBox "当前价格等级已在收费价目中使用，不允许删除！", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    CheckDepend = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmdCancel_Click()
    Err = 0: On Error GoTo ErrHandler
    If mblnChanged Then
        If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Sub
        End If
    End If
    Unload Me
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdHelp_Click()
    Err = 0: On Error GoTo ErrHandler
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Err = 0: On Error GoTo ErrHandler
    If mbytFun = Fun_View Then Unload Me: Exit Sub
    
    cmdOK.Enabled = False
    If IsValied() = False Then cmdOK.Enabled = True: Exit Sub
    If SaveData() = False Then cmdOK.Enabled = True: Exit Sub
    
    If mbytFun = Fun_Add Then
        mstr价格等级 = Trim(txtEdit(Txt_名称).Text)
    End If
    mblnOk = True
    Unload Me
    Exit Sub
ErrHandler:
    cmdOK.Enabled = True
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function IsValied() As Boolean
    '功能:数据检查
    '返回:检查通过返回True,否则返回False
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer, k As Integer
    Dim strTemp As String
    
    Err = 0: On Error GoTo ErrHandler
    If mbytFun = Fun_View Then IsValied = True: Exit Function
    If CheckDepend() = False Then Exit Function
    If mbytFun = Fun_Delete Then
        If MsgBox("你确认要删除名称为“" & mstr价格等级 & "”的价格等级吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Exit Function
        End If
        IsValied = True: Exit Function
    End If

    If zlControl.TxtCheckInput(txtEdit(Txt_编码), "编码", , False) = False Then Exit Function
    If zlControl.TxtCheckInput(txtEdit(Txt_名称), "名称", , False) = False Then Exit Function
    If zlControl.TxtCheckInput(txtEdit(Txt_简码), "简码") = False Then Exit Function
    
    If mbytFun = Fun_Update And mstr价格等级 <> Trim(txtEdit(Txt_名称).Text) Then
        '如果当前价格等级已经调价（即已经存在调价记录），则不允许更改名称。
        strSQL = "Select 1 From 收费价目 Where 价格等级 = [1] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查价格等级在收费价目中是否使用", mstr价格等级)
        If Not rsTemp.EOF Then
            MsgBox "当前价格等级已在收费价目中使用，不允许更改名称！", vbInformation + vbOKOnly, gstrSysName
            If txtEdit(Txt_名称).Visible And txtEdit(Txt_名称).Enabled Then txtEdit(Txt_名称).SetFocus
            Exit Function
        End If
    End If
    
    '编码唯一
    strSQL = "Select 1 From 收费价格等级 Where 编码 = [1] And 名称 <> [2] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "编码唯一检查", Trim(txtEdit(Txt_编码).Text), mstr价格等级)
    If Not rsTemp.EOF Then
        MsgBox "编码为“" & Trim(txtEdit(Txt_编码).Text) & "”的价格等级已存在！", vbInformation + vbOKOnly, gstrSysName
        If txtEdit(Txt_编码).Visible And txtEdit(Txt_编码).Enabled Then txtEdit(Txt_编码).SetFocus
        Exit Function
    End If
    
    '名称唯一
    strSQL = "Select 1 From 收费价格等级 Where 名称 = [1] And 名称 <> [2] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "名称唯一检查", Trim(txtEdit(Txt_名称).Text), mstr价格等级)
    If Not rsTemp.EOF Then
        MsgBox "名称为“" & Trim(txtEdit(Txt_名称).Text) & "”的价格等级已存在！", vbInformation + vbOKOnly, gstrSysName
        If txtEdit(Txt_名称).Visible And txtEdit(Txt_名称).Enabled Then txtEdit(Txt_名称).SetFocus
        Exit Function
    End If
    
    '一个站点，不能设置多个有效的等级
    strTemp = ""
    For i = 1 To lvwApply(Pg_NodeList).ListItems.Count
        If lvwApply(Pg_NodeList).ListItems(i).Checked Then
            strTemp = strTemp & "|" & lvwApply(Pg_NodeList).ListItems(i).Tag
        End If
    Next
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    strSQL = "Select /*+cardinality(B,10)*/c.名称 As 站点, a.价格等级" & vbNewLine & _
            " From 收费价格等级 D, 收费价格等级应用 A, Table(f_Str2list([1], '|')) B, Zlnodelist C" & vbNewLine & _
            " Where d.名称 = a.价格等级 And a.站点 = b.Column_Value And a.站点 = c.编号 And a.性质 = 0" & vbNewLine & _
            "       And (d.撤档时间 Is Null Or d.撤档时间 = To_Date('3000-01-01','yyyy-mm-dd')) And a.价格等级 <> [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "价格等级应用检查", strTemp, mstr价格等级)
    If Not rsTemp.EOF Then
        strTemp = ""
        Do While Not rsTemp.EOF
            strTemp = strTemp & vbCrLf & Nvl(rsTemp!站点) & "：" & Nvl(rsTemp!价格等级)
            rsTemp.MoveNext
        Loop
        If MsgBox("由于一个院区只能设置一个有效的价格等级，而你当前选择的" & _
            "以下院区已设置其它有效的价格等级。如果继续操作，将会清除这些院区的其它有效价格等级，" & _
            "然后应用当前价格等级，是否继续？" & vbCrLf & strTemp, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            If tbPageGradeApply.Item(Pg_NodeList).Selected = False Then tbPageGradeApply.Item(Pg_NodeList).Selected = True
            If lvwApply(Pg_NodeList).Visible And lvwApply(Pg_NodeList).Enabled Then lvwApply(Pg_NodeList).SetFocus
            Exit Function
        End If
    End If
    
    '一个医疗付款方式，不能设置多个有效的等级
    strTemp = ""
    For i = 1 To lvwApply(Pg_PatientType).ListItems.Count
        If lvwApply(Pg_PatientType).ListItems(i).Checked = True Then
            strTemp = strTemp & "|" & lvwApply(Pg_PatientType).ListItems(i).Text
        End If
    Next
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    strSQL = "Select /*+cardinality(B,10)*/a.医疗付款方式, a.价格等级" & vbNewLine & _
            " From 收费价格等级 D, 收费价格等级应用 A, Table(f_Str2list([1], '|')) B" & vbNewLine & _
            " Where d.名称 = a.价格等级 And a.医疗付款方式 = b.Column_Value And a.性质 = 1" & vbNewLine & _
            "       And (d.撤档时间 Is Null Or d.撤档时间 = To_Date('3000-01-01','yyyy-mm-dd')) And a.价格等级 <> [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "价格等级应用检查", strTemp, mstr价格等级)
    If Not rsTemp.EOF Then
        strTemp = ""
        Do While Not rsTemp.EOF
            strTemp = strTemp & vbCrLf & Nvl(rsTemp!医疗付款方式) & "：" & Nvl(rsTemp!价格等级)
            rsTemp.MoveNext
        Loop
        If MsgBox("由于一个医疗付款方式只能设置一个有效的价格等级，而你当前选择的" & _
            "以下医疗付款方式已设置其它有效的价格等级。如果继续操作，将会清除这些医疗付款方式的其它有效价格等级，" & _
            "然后应用当前价格等级，是否继续？" & vbCrLf & strTemp, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            If tbPageGradeApply.Item(Pg_PatientType).Selected = False Then tbPageGradeApply.Item(Pg_PatientType).Selected = True
            If lvwApply(Pg_PatientType).Visible And lvwApply(Pg_PatientType).Enabled Then lvwApply(Pg_PatientType).SetFocus
            Exit Function
        End If
    End If
    
    IsValied = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveData() As Boolean
    '功能:保存数据
    '返回:保存成功返回True,否则返回False
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim str应用站点 As String, str应用医疗付款方式 As String
    Dim i As Integer
    
    Err = 0: On Error GoTo ErrHandler
    If mbytFun = Fun_View Then SaveData = True: Exit Function
    If mbytFun = Fun_Add Or mbytFun = Fun_Update Then
        For i = 1 To lvwApply(Pg_NodeList).ListItems.Count
            If lvwApply(Pg_NodeList).ListItems(i).Checked = True Then
                str应用站点 = str应用站点 & "|" & lvwApply(Pg_NodeList).ListItems(i).Tag
            End If
        Next
        If str应用站点 <> "" Then str应用站点 = Mid(str应用站点, 2)
        
        For i = 1 To lvwApply(Pg_PatientType).ListItems.Count
            If lvwApply(Pg_PatientType).ListItems(i).Checked = True Then
                str应用医疗付款方式 = str应用医疗付款方式 & "|" & lvwApply(Pg_PatientType).ListItems(i).Text
            End If
        Next
        If str应用医疗付款方式 <> "" Then str应用医疗付款方式 = Mid(str应用医疗付款方式, 2)
    End If
    
    Select Case mbytFun
    Case Fun_Add
        'Zl_收费价格等级_Insert(
        strSQL = "Zl_收费价格等级_Insert("
        '  编码_In             In 收费价格等级.编码%Type,
        strSQL = strSQL & "'" & Trim(txtEdit(Txt_编码).Text) & "',"
        '  名称_In             In 收费价格等级.名称%Type,
        strSQL = strSQL & "'" & Trim(txtEdit(Txt_名称).Text) & "',"
        '  简码_In             In 收费价格等级.简码%Type,
        strSQL = strSQL & "'" & Trim(txtEdit(Txt_简码).Text) & "',"
        '  是否适用药品_In     In 收费价格等级.是否适用药品%Type := 0,
        strSQL = strSQL & "" & 0 & ","
        '  是否适用卫材_In     In 收费价格等级.是否适用卫材%Type := 0,
        strSQL = strSQL & "" & 0 & ","
        '  是否适用普通项目_In In 收费价格等级.是否适用普通项目%Type := 1,
        strSQL = strSQL & "" & 1 & ","
        '  应用站点_In         In Varchar2, --应用于的站点编号，多个用单竖线"|"分隔，如：01|02|...
        strSQL = strSQL & "'" & str应用站点 & "',"
        '  应用医疗付款方式_In In Varchar2 --应用于的医疗付款方式，多个用单竖线"|"分隔，如：公费医疗|自费医疗|...
        strSQL = strSQL & "'" & str应用医疗付款方式 & "')"
    Case Fun_Update
        'Zl_收费价格等级_Update(
        strSQL = "Zl_收费价格等级_Update("
        '  原名称_In           In 收费价格等级.名称%Type,
        strSQL = strSQL & "'" & mstr价格等级 & "',"
        '  编码_In             In 收费价格等级.编码%Type,
        strSQL = strSQL & "'" & Trim(txtEdit(Txt_编码).Text) & "',"
        '  名称_In             In 收费价格等级.名称%Type,
        strSQL = strSQL & "'" & Trim(txtEdit(Txt_名称).Text) & "',"
        '  简码_In             In 收费价格等级.简码%Type,
        strSQL = strSQL & "'" & Trim(txtEdit(Txt_简码).Text) & "',"
        '  是否适用药品_In     In 收费价格等级.是否适用药品%Type := 0,
        strSQL = strSQL & "" & 0 & ","
        '  是否适用卫材_In     In 收费价格等级.是否适用卫材%Type := 0,
        strSQL = strSQL & "" & 0 & ","
        '  是否适用普通项目_In In 收费价格等级.是否适用普通项目%Type := 1,
        strSQL = strSQL & "" & 1 & ","
        '  应用站点_In         In Varchar2, --应用于的站点编号，多个用单竖线"|"分隔，如：01|02|...
        strSQL = strSQL & "'" & str应用站点 & "',"
        '  应用医疗付款方式_In In Varchar2 --应用于的医疗付款方式，多个用单竖线"|"分隔，如：公费医疗|自费医疗|...
        strSQL = strSQL & "'" & str应用医疗付款方式 & "')"
    Case Fun_Delete
        'Zl_收费价格等级_Delete(
        strSQL = "Zl_收费价格等级_Delete("
        '  名称_In In 收费价格等级.名称%Type
        strSQL = strSQL & "'" & mstr价格等级 & "')"
    End Select
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Activate()
    Err = 0: On Error GoTo ErrHandler
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If Trim(txtEdit(Txt_编码).Text) <> "" Then
        If txtEdit(Txt_名称).Visible And txtEdit(Txt_名称).Enabled Then txtEdit(Txt_名称).SetFocus
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Err = 0: On Error GoTo ErrHandler
    
    mblnFirst = True: mblnLoading = True
    If InitPage() = False Then Unload Me: Exit Sub
    If GetDefineSize() = False Then Unload Me: Exit Sub
    If InitData() = False Then Unload Me: Exit Sub
    If LoadData() = False Then Unload Me: Exit Sub
    
    If Not (mbytFun = Fun_Add Or mbytFun = Fun_Update) Then
        Call ZlSetEnabled(Me.Controls, False)
        Call ZlSetEnabledBackColor(Me.Controls)
    End If
    
    Me.Caption = Choose(mbytFun + 1, "新增", "调整", "删除", "查看") & "价格等级"
    If mbytFun = Fun_View Then
        cmdOK.Visible = False
        cmdCancel.Caption = cmdOK.Caption
    End If
    mblnLoading = False
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function InitPage() As Boolean
    '功能:初始化页签控件
    Err = 0: On Error GoTo ErrHandler
    With tbPageGradeApply
        .RemoveAll
        .InsertItem Pg_NodeList, "院区", picApplyBack(Pg_NodeList).hwnd, 0
        .InsertItem Pg_PatientType, "医疗付款方式", picApplyBack(Pg_PatientType).hwnd, 0

         With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .Layout = xtpTabLayoutAutoSize
            .StaticFrame = True
            .ClientFrame = xtpTabFrameBorder
        End With
        .Item(Pg_NodeList).Selected = True
    End With
    InitPage = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitData() As Boolean
    '初始化界面基础数据
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim k As Integer, objListItem As ListItem
    
    Err = 0: On Error GoTo ErrHandler
    If mbytFun = Fun_Add Or mbytFun = Fun_Update Then
        If mbytFun = Fun_Add Then
            txtEdit(Txt_编码).Text = zlDatabase.GetMax("收费价格等级", "编码", 2)
        End If
        
        lvwApply(Pg_NodeList).ListItems.Clear
        lvwApply(Pg_PatientType).ListItems.Clear
        strSQL = "Select 1 As 类型, 编号 As 编码, 名称 From Zlnodelist" & vbNewLine & _
                " Union All" & vbNewLine & _
                " Select 2 As 类型, 编码, 名称 From 医疗付款方式" & vbNewLine & _
                " Order By 类型, 编码"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "读取基础数据")
        If rsTemp.RecordCount = 0 Then
            tbPageGradeApply.Item(Pg_PatientType).Visible = False
        Else
            '1.院区,2.医疗付款方式
            For k = 0 To 1
                rsTemp.Filter = "类型=" & IIF(k = 0, 1, 2)
                If rsTemp.RecordCount = 0 Then
                    tbPageGradeApply.Item(k).Visible = False
                    If k = Pg_NodeList Then tbPageGradeApply.Item(Pg_PatientType).Selected = True
                Else
                    Do While Not rsTemp.EOF
                        Set objListItem = lvwApply(k).ListItems.Add(, "K" & Nvl(rsTemp!编码), Nvl(rsTemp!名称))
                        objListItem.Tag = Nvl(rsTemp!编码)
                        rsTemp.MoveNext
                    Loop
                End If
            Next
        End If
    End If
    InitData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LoadData() As Boolean
    '加载数据
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer, k As Integer, blnFind As Boolean
    Dim objListItem As ListItem
    
    Err = 0: On Error GoTo ErrHandler
    If mbytFun = Fun_Add Then LoadData = True: Exit Function
    
    strSQL = "Select 编码, 名称, 简码, 是否适用药品, 是否适用卫材, 是否适用普通项目" & vbNewLine & _
            " From 收费价格等级" & vbNewLine & _
            " Where 名称 = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "收费价格等级", mstr价格等级)
    If rsTemp.EOF Then
        MsgBox "价格等级 " & mstr价格等级 & " 不存在，可能已被他人删除。请刷新后查看...", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    txtEdit(Txt_编码).Text = Nvl(rsTemp!编码)
    txtEdit(Txt_名称).Text = Nvl(rsTemp!名称)
    txtEdit(Txt_简码).Text = Nvl(rsTemp!简码)
    
    strSQL = "Select Nvl(a.性质, 0) As 性质, " & vbNewLine & _
            "        Decode(Nvl(a.性质, 0), 0, b.编号, c.编码) As 编码," & vbNewLine & _
            "        Decode(Nvl(a.性质, 0), 0, b.名称, c.名称) As 名称" & vbNewLine & _
            " From 收费价格等级应用 A, Zlnodelist B, 医疗付款方式 C" & vbNewLine & _
            " Where a.站点 = b.编号(+) And a.医疗付款方式 = c.名称(+) And a.价格等级 = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "收费价格等级", mstr价格等级)
    If rsTemp.EOF Then
        If mbytFun <> Fun_Update Then tbPageGradeApply.Item(Pg_PatientType).Visible = False
    Else
        For k = 0 To 1
            '0-院区,1-医疗付款方式
            rsTemp.Filter = "性质=" & IIF(k = 0, 0, 1)
            If rsTemp.RecordCount = 0 Then
                If mbytFun <> Fun_Update Then
                    tbPageGradeApply.Item(k).Visible = False
                    If k = Pg_NodeList Then tbPageGradeApply.Item(Pg_PatientType).Selected = True
                End If
            Else
                Do While Not rsTemp.EOF
                    blnFind = False
                    For i = 1 To lvwApply(k).ListItems.Count
                        Set objListItem = lvwApply(k).ListItems(i)
                        If objListItem.Tag = Nvl(rsTemp!编码) Then
                            objListItem.Checked = True
                            blnFind = True: Exit For
                        End If
                    Next
                    If blnFind = False Then
                        Set objListItem = lvwApply(k).ListItems.Add(, "K" & Nvl(rsTemp!编码), Nvl(rsTemp!名称))
                        objListItem.Tag = Nvl(rsTemp!编码)
                        objListItem.Checked = True
                    End If
                    rsTemp.MoveNext
                Loop
            End If
        Next
    End If
    LoadData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetDefineSize() As Boolean
'功能：得到数据库的表字段的长度
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    Err = 0: On Error GoTo ErrHandler
    strSQL = "Select 编码, 名称, 简码 From 收费价格等级 Where Rownum < 0"
    Call zlDatabase.OpenRecordset(rsTemp, strSQL, "收费价格等级编辑")
    
    txtEdit(Txt_编码).MaxLength = rsTemp.Fields("编码").DefinedSize
    txtEdit(Txt_名称).MaxLength = rsTemp.Fields("名称").DefinedSize
    txtEdit(Txt_简码).MaxLength = rsTemp.Fields("简码").DefinedSize
    
    GetDefineSize = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub lvwApply_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    Err = 0: On Error GoTo ErrHandler
    If mblnLoading Then Exit Sub
    mblnChanged = True
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwApply_KeyPress(Index As Integer, KeyAscii As Integer)
    Err = 0: On Error GoTo ErrHandler
    If KeyAscii = vbKeyReturn Then
        If tbPageGradeApply.Selected.Index < tbPageGradeApply.ItemCount - 1 Then
            tbPageGradeApply(tbPageGradeApply.Selected.Index + 1).Selected = True
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub picApplyBack_Resize(Index As Integer)
    On Error Resume Next
    With lvwApply(Index)
        .Left = 0
        .Top = 0
        .Width = picApplyBack(Index).ScaleWidth
        .Height = picApplyBack(Index).ScaleHeight
    End With
End Sub

Private Sub tbPageGradeApply_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    On Error Resume Next
    If lvwApply(Item.Index).Visible And lvwApply(Item.Index).Enabled Then lvwApply(Item.Index).SetFocus
End Sub

Private Sub txtEdit_Change(Index As Integer)
    Err = 0: On Error GoTo ErrHandler
    If mblnLoading Then Exit Sub
    mblnChanged = True
    If Index = Txt_名称 Then
        txtEdit(Txt_简码).Text = zlStr.GetCodeByVB(txtEdit(Txt_名称).Text)
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    Err = 0: On Error GoTo ErrHandler
    zlControl.TxtSelAll txtEdit(Index)
    If Index = Txt_名称 Then zlCommFun.OpenIme True
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Err = 0: On Error GoTo ErrHandler
    If InStr("'}|,""/", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    ElseIf KeyAscii = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    Err = 0: On Error GoTo ErrHandler
    zlCommFun.OpenIme False
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ZlSetEnabled(ByVal objControls As Object, ByVal blnEnabled As Boolean)
    '设置控件可用状态
    Dim i As Integer
    
    On Error Resume Next
    For i = 0 To objControls.Count - 1
        If UCase(objControls(i).Name) <> UCase("cmdHelp") _
            And UCase(objControls(i).Name) <> UCase("cmdOk") _
            And UCase(objControls(i).Name) <> UCase("cmdCancel") _
            And UCase(TypeName(objControls(i))) <> UCase("Label") _
            And UCase(TypeName(objControls(i))) <> UCase("Frame") _
            And UCase(TypeName(objControls(i))) <> UCase("TabControl") _
            And UCase(TypeName(objControls(i))) <> UCase("PictureBox") _
            And UCase(TypeName(objControls(i))) <> UCase("VSFlexGrid") Then
            objControls(i).Enabled = blnEnabled
        End If
    Next
End Sub

Private Sub ZlSetEnabledBackColor(ByVal objControls As Object)
    '设置控件可用状态与不可用状态的背景颜色
    Dim i As Integer
    
    On Error Resume Next
    For i = 0 To objControls.Count - 1
        If UCase(TypeName(objControls(i))) = UCase("TextBox") _
            Or UCase(TypeName(objControls(i))) = UCase("ComboBox") Then
            objControls(i).BackColor = IIF(objControls(i).Enabled, vbWindowBackground, vbButtonFace)
        End If
    Next
End Sub

