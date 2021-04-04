VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#7.0#0"; "zlIDKind.ocx"
Begin VB.Form Frm药品发药查找 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤设置"
   ClientHeight    =   2970
   ClientLeft      =   3255
   ClientTop       =   4680
   ClientWidth     =   6795
   Icon            =   "Frm药品发药查找.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   6795
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5460
      TabIndex        =   25
      Top             =   2490
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4230
      TabIndex        =   24
      Top             =   2490
      Width           =   1100
   End
   Begin VB.Frame fra附加条件 
      Caption         =   "附加条件"
      Enabled         =   0   'False
      Height          =   1155
      Left            =   120
      TabIndex        =   16
      Top             =   3120
      Width           =   6555
      Begin VB.ComboBox Cbo发药人 
         Height          =   300
         Left            =   4230
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   240
         Width           =   2085
      End
      Begin VB.ComboBox Cbo填制人 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   270
         Width           =   2085
      End
      Begin VB.CommandButton Cmd药品 
         Caption         =   "…"
         Enabled         =   0   'False
         Height          =   285
         Left            =   6000
         TabIndex        =   23
         Top             =   660
         Width           =   285
      End
      Begin VB.TextBox Txt药品 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   22
         Top             =   660
         Width           =   4725
      End
      Begin VB.CheckBox Chk药品 
         Caption         =   "药品(&P)"
         Height          =   210
         Left            =   270
         TabIndex        =   21
         Top             =   720
         Width           =   990
      End
      Begin VB.Label Lbl发药人 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "发药人"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3600
         TabIndex        =   19
         Top             =   330
         Width           =   540
      End
      Begin VB.Label Lbl填制人 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "填制人"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   17
         Top             =   330
         Width           =   540
      End
   End
   Begin VB.Frame Fra基本条件 
      Caption         =   "基本条件"
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   6555
      Begin VB.CheckBox chkSend 
         Caption         =   "离院带药"
         Height          =   180
         Index           =   1
         Left            =   5280
         TabIndex        =   29
         Top             =   1860
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox chkSend 
         Caption         =   "院内用药"
         Height          =   180
         Index           =   0
         Left            =   4200
         TabIndex        =   28
         Top             =   1860
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox txt医保号 
         Height          =   300
         Left            =   930
         TabIndex        =   26
         Top             =   1800
         Width           =   2085
      End
      Begin VB.TextBox txt就诊卡 
         Height          =   300
         Left            =   4200
         MaxLength       =   20
         TabIndex        =   11
         Top             =   1050
         Width           =   2085
      End
      Begin VB.TextBox txt住院号 
         Height          =   300
         Left            =   930
         MaxLength       =   20
         TabIndex        =   13
         Top             =   1440
         Width           =   2085
      End
      Begin VB.TextBox Txt姓名 
         Height          =   300
         Left            =   930
         MaxLength       =   12
         TabIndex        =   10
         Top             =   1050
         Width           =   2085
      End
      Begin VB.ComboBox Cbo科室 
         Height          =   276
         Left            =   4200
         TabIndex        =   15
         Text            =   "Cbo科室"
         Top             =   1440
         Width           =   2085
      End
      Begin MSComCtl2.DTPicker Dtp开始Date 
         Height          =   300
         Left            =   930
         TabIndex        =   2
         Top             =   270
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   85524483
         CurrentDate     =   37007
      End
      Begin VB.TextBox Txt结束NO 
         Height          =   300
         Left            =   4200
         MaxLength       =   8
         TabIndex        =   8
         Top             =   660
         Width           =   2085
      End
      Begin VB.TextBox Txt开始NO 
         Height          =   300
         Left            =   930
         MaxLength       =   8
         TabIndex        =   6
         Top             =   660
         Width           =   2085
      End
      Begin MSComCtl2.DTPicker Dtp结束Date 
         Height          =   300
         Left            =   4200
         TabIndex        =   4
         Top             =   270
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   85524483
         CurrentDate     =   37007
      End
      Begin zlIDKind.IDKindNew IDKNType 
         Height          =   300
         Left            =   3240
         TabIndex        =   31
         Top             =   1050
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   529
         ShowSortName    =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   9
         FontName        =   "宋体"
         IDKind          =   -1
         ShowPropertySet =   -1  'True
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         BackColor       =   -2147483633
      End
      Begin VB.Label lbl发药类型 
         AutoSize        =   -1  'True
         Caption         =   "发药类型"
         Height          =   180
         Left            =   3360
         TabIndex        =   30
         Top             =   1860
         Width           =   720
      End
      Begin VB.Label lbl医保号 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "医保号"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   27
         Top             =   1860
         Width           =   540
      End
      Begin VB.Label lbl住院号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "标识号"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   12
         Top             =   1500
         Width           =   540
      End
      Begin VB.Label Lbl姓名 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   510
         TabIndex        =   9
         Top             =   1110
         Width           =   360
      End
      Begin VB.Label Lbl科室 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "科室"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3780
         TabIndex        =   14
         Top             =   1500
         Width           =   360
      End
      Begin VB.Label Lbl结束Date 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "结束日期"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3420
         TabIndex        =   3
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Lbl开始Date 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "开始日期"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   150
         TabIndex        =   1
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Lbl结束NO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "结束NO"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3600
         TabIndex        =   7
         Top             =   720
         Width           =   540
      End
      Begin VB.Label Lbl开始NO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "开始NO"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   5
         Top             =   720
         Width           =   540
      End
   End
End
Attribute VB_Name = "Frm药品发药查找"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--窗体高度--
Private Const DblNormalHeight As Double = 3330
Private Const DblAdvanceHeight As Double = 4845
Private FrmObj As Form
Private mstrPrivs As String                             '权限，用来检查是否有附加条件选择的权限，以决定窗体的大小

'--本程序使用--
Private BlnStartUp As Boolean                           '启动成功
Private strReturn As String
Private BlnState As Boolean                             '状态(此状态决定窗体高度及输出的SQL)
Private mbln就诊卡 As Boolean

Private mobjSquareCard As Object             '一卡通接口

Private mlng病人ID As Long

'--外部传入参数--
Private lng药房ID As Long                               '库房ID
Private Int单据 As Integer                              '单据
Private IntOper As Integer

Private Type Type_SQLCondition
    date开始日期 As Date
    date结束日期 As Date
    str开始NO As String
    str结束NO As String
    str姓名 As String
    str就诊卡 As String
    str标识号 As String
    lng科室ID As Long
    str填制人 As String
    str审核人 As String
    lng药品id As Long
    str医保号 As String
End Type

Private SQLCondition As Type_SQLCondition

Private mint离院带药 As Integer

Private Sub Cbo科室_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim str工作性质 As String
    
    str工作性质 = "A,D"
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Cbo科室.ListCount = 0 Then Exit Sub
    
    If Cbo科室.ListIndex >= 0 Then
        If Val(Cbo科室.Tag) = Cbo科室.ItemData(Cbo科室.ListIndex) Then
            Exit Sub
        End If
    End If
    
    If Select部门选择器(Me, Cbo科室, Trim(Cbo科室.Text), str工作性质, , "1,2,3") = False Then
        Exit Sub
    End If
    If Cbo科室.ListIndex >= 0 Then
        Cbo科室.Tag = Cbo科室.ItemData(Cbo科室.ListIndex)
    End If
End Sub

Private Sub Cbo科室_KeyPress(KeyAscii As Integer)
    '屏蔽输入单引号
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Cbo科室_Validate(Cancel As Boolean)
    If Cbo科室.ListCount > 0 Then
        If Cbo科室.ListIndex = -1 Then
            MsgBox "请选择一个药库或者药房！", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub

Private Sub Chk药品_Click()
    Txt药品.Enabled = IIf(Chk药品.Value = 1, True, False)
    Cmd药品.Enabled = Txt药品.Enabled
    If Txt药品.Enabled Then Txt药品.SetFocus
End Sub

Private Sub CmdCancel_Click()
    strReturn = ""
    Unload Me
End Sub

Private Sub InitIDKindNew()
    Call IDKNType.zlInit(Me, glngSys, 1341, gcnOracle, gstrDbUser, mobjSquareCard, "", txt就诊卡, , True)
End Sub

Private Sub cmdOk_Click()
    If CheckData = False Then Exit Sub
    Call GetSQL
    
    FrmObj.int模式 = IIf(BlnState, -1, 1)
    Unload Me
End Sub

Private Sub cmd药品_Click()
    Dim RecReturn As New ADODB.Recordset
    
'    With Frm药品选择器
'        Set RecReturn = .ShowME(Me, 1, lng药房ID, , , False)
'    End With
    
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(1, "药品处方发药", lng药房ID, lng药房ID)
    End If
    Set RecReturn = frmSelector.ShowMe(Me, 0, 1, , , , lng药房ID, , , , False, , , , False)
    
    With RecReturn
        If .EOF Then Exit Sub
        Txt药品.Tag = !药品ID
        Txt药品 = "[" & !药品编码 & "]" & IIf(IsNull(!通用名), "", !通用名)
    End With
End Sub

Private Sub Form_Load()
    Dim intDays As Integer
    Dim dateCurDate As Date
    
    BlnStartUp = False
    BlnState = (IntOper = 6)
    strReturn = ""
    
    intDays = Val(zldatabase.GetPara("查询天数", glngSys, 1341, 1))
    intDays = intDays - 1
    
    dateCurDate = Sys.Currentdate()
    Me.Dtp开始Date.Value = Format(DateAdd("d", -1 * intDays, dateCurDate), "yyyy-MM-dd 00:00:00")
    Me.Dtp结束Date.Value = Format(dateCurDate, "yyyy-MM-dd 23:59:59")
    
    Select Case IntOper
    Case 1
        Me.Caption = "查找未配药处方单据"
    Case 2
        Me.Caption = "查找已配药处方单据"
    Case 3
        Me.Caption = "查找未发药处方单据"
    Case 4
        Me.Caption = "查找超时未发药处方单据"
    Case 5
        Me.Caption = "查找已发药处方单据"
    End Select
    
    If DependOnCheck = False Then Exit Sub
    If glngSys \ 100 <> 1 Then
        Lbl科室.Visible = False
        Cbo科室.Visible = False
    End If
    
    If IntOper <> 5 Then
        Me.Dtp开始Date.Enabled = zlStr.IsHavePrivs(mstrPrivs, "修改过滤日期")
        Me.Dtp结束Date.Enabled = Me.Dtp开始Date.Enabled
    End If
    
    If Not IsInString(mstrPrivs, "允许查询所有时间范围单据", ";") Then
        Dtp开始Date.Enabled = False
        Dtp结束Date.Enabled = False
    End If
    
    BlnStartUp = True
    
    Call zlfuncCard_Ini(mobjSquareCard, Me, 1341)
    Call InitIDKindNew
    
    On Error Resume Next
    If mbln就诊卡 Then txt就诊卡.SetFocus
End Sub

Private Sub Form_Activate()
    If BlnStartUp = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ReleaseSelectorRS
End Sub

Private Sub IDKNType_ItemClick(index As Integer, objCard As zlIDKind.Card)
    If objCard.卡号密文规则 <> "" Then
        txt就诊卡.PasswordChar = "*"
    Else
        txt就诊卡.PasswordChar = ""
    End If
End Sub

Private Sub IDKNType_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    txt就诊卡.Text = objPatiInfor.卡号
End Sub

Private Sub Txt结束NO_GotFocus()
    GetFocus Txt结束NO
End Sub

Private Sub Txt结束NO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intYear As Integer, strYear As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(Txt结束NO) = "" Then Exit Sub
    '--如果不满八位,则按规则产生--
    Me.Txt结束NO = GetFullNO(UCase(LTrim(Me.Txt结束NO)), 13)
End Sub

Private Sub txt就诊卡_Change()
    If txt就诊卡.Text <> "" And Len(txt就诊卡.Text) = 18 And Not mobjSquareCard Is Nothing And IDKNType.GetCurCard.名称 = "二代身份证" Then
        If mobjSquareCard.zlGetPatiID("身份证", UCase(txt就诊卡.Text), False, mlng病人ID) = False Then mlng病人ID = 0
    End If
End Sub

Private Sub txt就诊卡_KeyPress(KeyAscii As Integer)
    '去掉磁卡的其他的特殊字符
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
End Sub


Private Sub Txt开始NO_GotFocus()
    GetFocus Txt开始NO
End Sub

Private Sub Txt开始NO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(Txt开始NO) = "" Then Exit Sub
    '--如果不满八位,则按规则产生--
    Me.Txt开始NO = GetFullNO(UCase(LTrim(Me.Txt开始NO)), 13)
End Sub



Private Sub Txt姓名_GotFocus()
    GetFocus Txt姓名
End Sub

Private Sub Txt药品_GotFocus()
    GetFocus Txt药品
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.Height = IIf((BlnState And zlStr.IsHavePrivs(mstrPrivs, "过滤附加条件")), DblAdvanceHeight, DblNormalHeight)
    fra附加条件.Enabled = (BlnState And zlStr.IsHavePrivs(mstrPrivs, "过滤附加条件"))
    
    With fra附加条件
        .Top = IIf((BlnState And zlStr.IsHavePrivs(mstrPrivs, "过滤附加条件")), Fra基本条件.Top + Fra基本条件.Height + 80, CmdOK.Top + CmdOK.Height + 180)
    End With
    With CmdOK
        .Top = IIf((BlnState And zlStr.IsHavePrivs(mstrPrivs, "过滤附加条件")), fra附加条件.Top + fra附加条件.Height + 80, Fra基本条件.Top + Fra基本条件.Height + 80)
    End With
    With CmdCancel
        .Top = CmdOK.Top
    End With
    
    If fra附加条件.Enabled = True Then
        Me.Cbo填制人.Enabled = zlStr.IsHavePrivs(mstrPrivs, "医生查询")
    End If
End Sub

Private Function DependOnCheck() As Boolean
    Dim RecTmp As ADODB.Recordset

    '检查依赖数据是否完整
    DependOnCheck = False
    
    On Error GoTo errHandle
    Cbo科室.Clear
    
    If glngSys \ 100 = 1 Then
        '根据当前部门的服务对象，提取科室
        gstrSQL = " Select 编码||'-'||名称 科室,ID From 部门表 " & _
                 " Where ID in (Select 部门ID From 部门性质说明 Where 工作性质 In ('临床','手术') And 服务对象 IN(1,2,3))" & _
                 " And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) " & _
                 " Order By 编码||'-'||名称 "
        Set RecTmp = zldatabase.OpenSQLRecord(gstrSQL, "DependOnCheck")

        If RecTmp.EOF Then
            MsgBox "请初始化部门表（部门管理）！", vbInformation, gstrSysName
            Exit Function
        End If
        Me.Cbo科室.AddItem "所有"
        Do While Not RecTmp.EOF
            Cbo科室.AddItem RecTmp!科室
            Cbo科室.ItemData(Cbo科室.NewIndex) = RecTmp!Id
            RecTmp.MoveNext
        Loop
        Cbo科室.ListIndex = 0
    End If
        
    If IntOper = 6 Then
        '添加填制人
        Cbo填制人.Clear
        Cbo填制人.AddItem "所有"

        gstrSQL = " Select distinct 姓名 填制人 From 人员表" & _
                 " Where ID IN (" & _
                 " Select 人员ID From 人员性质说明" & _
                 " Where 人员性质='医生')" & _
                 " And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) "
         
        Set RecTmp = zldatabase.OpenSQLRecord(gstrSQL, "DependOnCheck")
        
        Do While Not RecTmp.EOF
            Cbo填制人.AddItem Trim(RecTmp!填制人)
            RecTmp.MoveNext
        Loop
        Cbo填制人.ListIndex = 0
        
        '添加发药人
        Cbo发药人.Clear
        Cbo发药人.AddItem "所有"

        gstrSQL = " Select distinct 姓名 审核人 From 人员表" & _
                 " Where ID IN (" & _
                 " Select 人员ID From 人员性质说明" & _
                 " Where 人员性质='药房发药人')" & _
                 " And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) "
        Set RecTmp = zldatabase.OpenSQLRecord(gstrSQL, "DependOnCheck")

        Do While Not RecTmp.EOF
            Cbo发药人.AddItem Trim(RecTmp!审核人)
            RecTmp.MoveNext
        Loop
        Cbo发药人.ListIndex = 0
    End If

    DependOnCheck = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckData() As Boolean
    '检查数据正确性
    CheckData = False
    
    Txt开始NO = UCase(Trim(Txt开始NO))
    Txt结束NO = UCase(Trim(Txt结束NO))
    Txt姓名 = UCase(Trim(Txt姓名))
    
    If BlnState Then
'        Cbo填制人 = Trim(Cbo填制人)
'        Cbo发药人 = Trim(Cbo发药人)
        If Chk药品.Value = 1 Then
            If Txt药品.Tag = 0 Then
                MsgBox "请输入药品信息！", vbInformation, gstrSysName
                Txt药品.SetFocus
                Exit Function
            End If
        End If
    End If
    
    CheckData = True
End Function

Private Function GetSQL()
    '根据用户输入产生SQL
    strReturn = ""
    
    If BlnState = False Then
        strReturn = " And A.填制日期 Between [1] And [2] "
    Else
        strReturn = " And A.审核日期 Between [1] And [2] "
    End If
    
    If Txt开始NO <> "" Or Txt结束NO <> "" Then
        If Txt开始NO <> "" And Txt结束NO <> "" Then
            strReturn = strReturn & " And A.NO Between [3] And [4] "
        Else
            If Txt开始NO <> "" Then
                strReturn = strReturn & " And A.NO = [3] "
            Else
                strReturn = strReturn & " And A.NO = [4] "
            End If
        End If
    End If
    
    If BlnState = False Then
        If Txt姓名 <> "" Then strReturn = strReturn & " And Upper(A.姓名) Like [5] "
        If txt住院号 <> "" Then strReturn = strReturn & " And Upper(DECODE(A.单据,8,A.门诊号,A.住院号)) Like [7] "
        If Cbo科室.ListIndex <> 0 And glngSys \ 100 = 1 Then strReturn = strReturn & " And C.对方部门ID+0=[8] "
    Else
        If Txt姓名 <> "" Then strReturn = strReturn & " And Upper(H.姓名) Like [5] "
        If txt住院号 <> "" Then strReturn = strReturn & " And Upper(H.标识号) Like [7] "
        If Cbo科室.ListIndex <> 0 And glngSys \ 100 = 1 Then strReturn = strReturn & " And A.对方部门ID+0=[8] "
    End If
    If Trim(txt就诊卡.Text) <> "" Then
        mbln就诊卡 = True
        If BlnState = False Then
            strReturn = strReturn & " And Upper(A.就诊卡号) = [6] "
        Else
            strReturn = strReturn & " And Upper(B.就诊卡号) = [6] "
        End If
    End If
    
    SQLCondition.date开始日期 = CDate(Format(Me.Dtp开始Date, "yyyy-MM-dd hh:mm:ss"))
    SQLCondition.date结束日期 = CDate(Format(Me.Dtp结束Date, "yyyy-MM-dd hh:mm:ss"))
    SQLCondition.str开始NO = Txt开始NO
    SQLCondition.str结束NO = Txt结束NO
    SQLCondition.str姓名 = IIf(Txt姓名 = "", "", Txt姓名 & "%")
    SQLCondition.str就诊卡 = txt就诊卡.Text & IIf(txt就诊卡.Text = "", "", "|" & IDKNType.GetCurCard.名称 & "," & IDKNType.GetCurCard.接口序号 & IIf(mlng病人ID <> 0, "|" & mlng病人ID, ""))
    SQLCondition.str标识号 = IIf(txt住院号 = "", "", txt住院号 & "%")
    SQLCondition.lng科室ID = Cbo科室.ItemData(Cbo科室.ListIndex)
    SQLCondition.str医保号 = UCase(Trim(txt医保号.Text))
    
    '发药类型
    If chkSend(0).Value = 1 And chkSend(1).Value = 1 Then
        mint离院带药 = 0
    ElseIf chkSend(0).Value = 1 Then
        mint离院带药 = 1
    ElseIf chkSend(1).Value = 1 Then
        mint离院带药 = 2
    End If
    
    If BlnState = False Then Exit Function
    
    If Cbo填制人.ListIndex <> 0 Then strReturn = strReturn & " And Trim(A.填制人) Like [9] "
    If Cbo发药人.ListIndex <> 0 Then strReturn = strReturn & " And Trim(A.审核人) Like [10] "
    If Val(Txt药品.Tag) <> 0 Then strReturn = strReturn & " And A.药品ID+0=[11] "
    
    SQLCondition.str填制人 = ""
    SQLCondition.str审核人 = ""
    SQLCondition.lng药品id = 0
    
    If Cbo填制人.Text <> "" And Cbo填制人.Text <> "所有" Then SQLCondition.str填制人 = Cbo填制人.Text
    If Cbo发药人.Text <> "" And Cbo发药人.Text <> "所有" Then SQLCondition.str审核人 = Cbo发药人.Text
    SQLCondition.lng药品id = Val(Txt药品.Tag)
End Function

Public Function ShowMe(ByVal FrmMain As Form, ByVal In_Lng药房ID As Long, _
    ByVal In_Int操作模式 As Integer, ByVal In_权限 As String, bln就诊卡 As Boolean, _
    ByRef date开始日期 As Date, _
    ByRef date结束日期 As Date, _
    ByRef str开始NO As String, _
    ByRef str结束NO As String, _
    ByRef str姓名 As String, _
    ByRef str就诊卡 As String, _
    ByRef str标识号 As String, _
    ByRef lng科室ID As Long, _
    ByRef str填制人 As String, _
    ByRef str审核人 As String, _
    ByRef lng药品id As Long, _
    ByRef str医保号 As String, _
    ByRef int离院带药 As Integer) As String
    
    lng药房ID = In_Lng药房ID
    IntOper = In_Int操作模式
    mstrPrivs = In_权限
    mbln就诊卡 = bln就诊卡
    
    Set FrmObj = FrmMain
    With Me
        .Show 1, FrmMain
    End With
    
    bln就诊卡 = mbln就诊卡
    
    date开始日期 = SQLCondition.date开始日期
    date结束日期 = SQLCondition.date结束日期
    str开始NO = SQLCondition.str开始NO
    str结束NO = SQLCondition.str结束NO
    str姓名 = SQLCondition.str姓名
    str就诊卡 = SQLCondition.str就诊卡
    str标识号 = SQLCondition.str标识号
    lng科室ID = SQLCondition.lng科室ID
    str填制人 = SQLCondition.str填制人
    str审核人 = SQLCondition.str审核人
    lng药品id = SQLCondition.lng药品id
    str医保号 = SQLCondition.str医保号
    int离院带药 = mint离院带药
    
    ShowMe = strReturn
End Function

Private Sub Txt药品_Validate(Cancel As Boolean)
    Txt药品 = Trim(Txt药品)
    If Txt药品 = "" Then
        Txt药品.Tag = 0
        Exit Sub
    End If
    
    Dim RecReturn As New ADODB.Recordset
    Dim sngLeft As Single, sngTop As Single
    
    If InStr(1, Txt药品, "[") <> 0 And InStr(1, Txt药品, "]") <> 0 Then Txt药品.Text = Mid(Txt药品.Text, 2, InStr(1, Txt药品, "]") - 2)
    sngLeft = Me.Left + Txt药品.Left + fra附加条件.Left + 50
    sngTop = Me.Top + (Me.Height - Me.ScaleHeight) + Txt药品.Top + fra附加条件.Top + Txt药品.Height - 100
    If DblFrmHeight + sngTop > Screen.Height Then sngTop = sngTop - DblFrmHeight - Txt药品.Height + 50
'    With Frm药品多选选择器
'        Set RecReturn = .ShowME(Me, 1, lng药房ID, , , Txt药品.Text, sngLeft, sngTop, False)
'        If RecReturn.EOF Then Cancel = True: Exit Sub
'    End With
    
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(1, "药品处方发药", lng药房ID, lng药房ID)
    End If
    Set RecReturn = frmSelector.ShowMe(Me, 1, 1, UCase(Txt药品.Text), sngLeft, sngTop, lng药房ID, , , , False, , , , False)
    
    If RecReturn.EOF Then Cancel = True: Exit Sub
    Txt药品.Tag = RecReturn!药品ID
    Txt药品 = "[" & RecReturn!药品编码 & "]" & IIf(IsNull(RecReturn!通用名), "", RecReturn!通用名)
End Sub
