VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPriceFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤设置"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra收费类别 
      Caption         =   "收费类别"
      Height          =   2295
      Left            =   30
      TabIndex        =   25
      Top             =   3180
      Width           =   5775
      Begin MSComctlLib.ListView lvwType 
         Height          =   2025
         Left            =   60
         TabIndex        =   13
         Top             =   210
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   3572
         View            =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.CommandButton cmdDef 
      Caption         =   "缺省(&D)"
      Height          =   350
      Left            =   5970
      TabIndex        =   16
      Top             =   1455
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   3120
      Left            =   30
      TabIndex        =   17
      Top             =   0
      Width           =   5790
      Begin VB.OptionButton opt病人 
         Caption         =   "门诊病人和住院病人"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   2
         Left            =   3720
         TabIndex        =   12
         Top             =   2790
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton opt病人 
         Caption         =   "住院病人"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   2220
         TabIndex        =   11
         Top             =   2790
         Width           =   1020
      End
      Begin VB.OptionButton opt病人 
         Caption         =   "门诊病人"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   840
         TabIndex        =   10
         Top             =   2790
         Width           =   1020
      End
      Begin VB.TextBox txt姓名 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   3600
         MaxLength       =   64
         TabIndex        =   6
         Top             =   1500
         Width           =   2070
      End
      Begin VB.TextBox txtPatient 
         Height          =   300
         IMEMode         =   1  'ON
         Left            =   1440
         MaxLength       =   64
         TabIndex        =   9
         Top             =   2340
         Width           =   4215
      End
      Begin VB.CheckBox chk收费 
         Caption         =   "仅已收费单据"
         Height          =   255
         Left            =   3600
         TabIndex        =   2
         Top             =   713
         Width           =   1440
      End
      Begin VB.ComboBox cbo费别 
         Height          =   300
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1932
         Width           =   2070
      End
      Begin VB.ComboBox cbo科室 
         Height          =   300
         Left            =   855
         TabIndex        =   5
         Text            =   "cbo科室"
         Top             =   1500
         Width           =   2070
      End
      Begin VB.ComboBox cbo操作员 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   855
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1920
         Width           =   2070
      End
      Begin VB.TextBox txtNOBegin 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   855
         MaxLength       =   8
         TabIndex        =   3
         Top             =   1098
         Width           =   2070
      End
      Begin VB.TextBox txtNoEnd 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3600
         MaxLength       =   8
         TabIndex        =   4
         Top             =   1098
         Width           =   2070
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   855
         TabIndex        =   1
         Top             =   690
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   146800643
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   855
         TabIndex        =   0
         Top             =   270
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   146800643
         CurrentDate     =   36588
      End
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   300
         Left            =   840
         TabIndex        =   26
         Top             =   2340
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   529
         Appearance      =   2
         IDKindStr       =   "就|就诊卡|0|0|0|0|0|;医|医保号|0|0|0|0|0|;身|身份证号|0|0|0|0|0|;IC|IC卡号|1|0|0|0|0|;门|门诊号|0|0|0|0|0|"
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
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         BackColor       =   -2147483633
      End
      Begin VB.Label lblFil 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "病人来源"
         Height          =   180
         Left            =   60
         TabIndex        =   29
         Top             =   2790
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         Height          =   180
         Left            =   3180
         TabIndex        =   28
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份识别"
         Height          =   180
         Left            =   60
         TabIndex        =   27
         Top             =   2400
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "费别"
         Height          =   180
         Left            =   3180
         TabIndex        =   24
         Top             =   1980
         Width           =   360
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开始时间"
         Height          =   180
         Left            =   60
         TabIndex        =   23
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "结束时间"
         Height          =   180
         Left            =   60
         TabIndex        =   22
         Top             =   750
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   3180
         TabIndex        =   21
         Top             =   1155
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据号"
         Height          =   180
         Left            =   240
         TabIndex        =   20
         Top             =   1155
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开单科室"
         Height          =   180
         Left            =   60
         TabIndex        =   19
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label lbl操作员 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "划价人"
         Height          =   180
         Left            =   240
         TabIndex        =   18
         Top             =   1980
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5970
      TabIndex        =   15
      Top             =   645
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5970
      TabIndex        =   14
      Top             =   225
      Width           =   1100
   End
End
Attribute VB_Name = "frmPriceFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mstrFilter As String
Public mblnDateMoved As Boolean
Public mstrPrivs As String
Public mstr收费类别 As String


Public mlngPrePatient As Long
Private mblnKeyReturn As Boolean
Private mblnNotClick As Boolean
Private mblnUnChange  As Boolean
Private mrsInfo As ADODB.Recordset
Private mblnOlnyBJYB As Boolean
Private mrsDept As ADODB.Recordset

Private Sub cbo操作员_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii >= 32 Then
        lngIdx = zlControl.CboMatchIndex(cbo操作员.hWnd, KeyAscii)
        If lngIdx = -1 And cbo操作员.ListCount > 0 Then lngIdx = 0
        cbo操作员.ListIndex = lngIdx
    End If
End Sub

Private Sub cbo费别_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If KeyAscii >= 32 Then
        lngIdx = zlControl.CboMatchIndex(cbo费别.hWnd, KeyAscii)
        If lngIdx = -1 And cbo费别.ListCount > 0 Then lngIdx = 0
        cbo费别.ListIndex = lngIdx
    End If
End Sub

Private Sub cbo科室_Click()
    Dim lng开单科室ID As Long
    
    On Error GoTo errHandler
    If cbo科室.ListIndex <> -1 Then lng开单科室ID = cbo科室.ItemData(cbo科室.ListIndex)
    If Val(cbo科室.Tag) = lng开单科室ID Then Exit Sub
    cbo科室.Tag = lng开单科室ID
        
    '定位医生
    If gbyt科室医生 = 1 Then
        If cbo科室.ListIndex <> -1 Then
            Call FillPerson(lng开单科室ID)
        Else
            cbo操作员.Clear
        End If
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FillPerson(Optional ByVal lng科室ID As Long)
'功能：根据指定的开单科室ID读取并填写医生列表,但不缺省医生
    Dim rsTmp As ADODB.Recordset
    Dim bln仅操作员部门 As Boolean
    
    cbo操作员.Clear
    cbo操作员.AddItem "所有划价人"
    bln仅操作员部门 = zlStr.IsHavePrivs(mstrPrivs, "所有科室") = False And gblnUserIsClinic '113577
    Set rsTmp = GetPersonnel("药房发药人,医生,护士", True, bln仅操作员部门, lng科室ID)
    Do While Not rsTmp.EOF
        cbo操作员.AddItem rsTmp!简码 & "-" & rsTmp!姓名
        If rsTmp!ID = UserInfo.ID Then cbo操作员.ListIndex = cbo操作员.NewIndex
        rsTmp.MoveNext
    Loop
    If cbo操作员.ListIndex < 0 And cbo操作员.ListCount > 0 Then cbo操作员.ListIndex = 0
End Sub

Private Sub cbo科室_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
'    If KeyAscii >= 32 Then
'        lngIdx = zlControl.CboMatchIndex(cbo科室.hWnd, KeyAscii)
'        If lngIdx = -1 And cbo科室.ListCount > 0 Then lngIdx = 0
'        cbo科室.ListIndex = lngIdx
'    End If
    
    If KeyAscii <> 13 Then Exit Sub
    
    If cbo科室.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    
    If mrsDept Is Nothing Then Set mrsDept = GetDepartments("'临床','手术'", gint病人来源 & ",3")
    If zlSelectDept(Me, 1120, cbo科室, mrsDept, cbo科室.Text, True, _
        IIf(zlStr.IsHavePrivs(mstrPrivs, "所有科室"), "所有科室", "")) = False Then
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
End Sub

Private Sub cbo科室_Validate(Cancel As Boolean)
    If cbo科室.ListIndex >= 0 Then Exit Sub
    If cbo科室.ListIndex < 0 And cbo科室.ListCount <> 0 Then cbo科室.ListIndex = 0
End Sub

Private Sub cmdCancel_Click()
    gblnOK = False
    Hide
End Sub

Private Sub cmdDef_Click()
    Form_Load
End Sub

Private Sub cmdOK_Click()
    If txtNOBegin.Text <> "" And txtNoEnd.Text <> "" Then
        If txtNoEnd.Text < txtNOBegin.Text Then
            MsgBox "结束单据号不能小于开始单据号！", vbInformation, gstrSysName
            txtNoEnd.SetFocus: Exit Sub
        End If
    End If
    Call zlGet收费类别
    If mstr收费类别 = "" Then
        MsgBox "必须选择一个类别进行查找,请选择类别!", vbInformation + vbOKOnly, gstrSysName
        If lvwType.Enabled Then lvwType.SetFocus
        Exit Sub
    End If
    Call MakeFilter
    
    gblnOK = True
    Hide
End Sub

Private Sub dtpEnd_Change()
    dtpBegin.MaxDate = dtpEnd.Value
End Sub

Private Sub Form_Activate()
    dtpBegin.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.ActiveControl Is cbo科室 Then Exit Sub
    
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr(1, "'[]", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If Me.ActiveControl Is cbo科室 Then Exit Sub
    
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim Curdate As Date, i As Integer, lngOldID As Long
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim bln仅操作员部门 As Boolean
    On Error GoTo errH
    
    gblnOK = False
    
    txtNOBegin.Text = ""
    txtNoEnd.Text = ""
    txtPatient.Text = ""
    
    chk收费.Value = 0
    Call InitIDKind
    '设置初始值
    Curdate = zlDatabase.Currentdate
    dtpBegin.MaxDate = Format(Curdate, "yyyy-MM-dd 23:59:59")
    dtpBegin.Value = Format(Curdate, "yyyy-MM-dd 00:00:00")
    dtpEnd.Value = dtpBegin.MaxDate
    
    '操作员
    Call FillPerson
    cbo.SetListWidth cbo操作员.hWnd, cbo操作员.Width * 3 / 2
    
    '可选费别
    cbo费别.Clear
    cbo费别.AddItem "所有费别"
    cbo费别.ListIndex = 0
    
    strSQL = "Select 编码,名称,简码,Nvl(缺省标志,0) as 缺省 From 费别 Where Nvl(服务对象,3) IN(1,3) Order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cbo费别.AddItem rsTmp!编码 & "-" & rsTmp!名称
            rsTmp.MoveNext
        Next
    End If
    
    '开单科室
    bln仅操作员部门 = zlStr.IsHavePrivs(mstrPrivs, "所有科室") = False And gblnUserIsClinic '113577
    cbo科室.Clear: cbo科室.Tag = ""
    If bln仅操作员部门 = False Then cbo科室.AddItem "所有科室"
    Set mrsDept = GetDepartments("'临床','手术','治疗','检查','检验','产科'", gint病人来源 & ",3", bln仅操作员部门)
    For i = 1 To mrsDept.RecordCount
        If lngOldID <> mrsDept!ID Then
            cbo科室.AddItem mrsDept!编码 & "-" & mrsDept!名称
            cbo科室.ItemData(cbo科室.NewIndex) = mrsDept!ID
            lngOldID = mrsDept!ID
        End If
        mrsDept.MoveNext
    Next
    If cbo科室.ListIndex < 0 And cbo科室.ListCount > 0 Then cbo科室.ListIndex = 0
    cbo.SetListWidth cbo科室.hWnd, cbo科室.Width * 3 / 2
    
    Dim str收费类别 As String
    str收费类别 = zlDatabase.GetPara("上次过滤收费类别", glngSys, 1120, "", Array(lvwType, fra收费类别), InStr(1, mstrPrivs, "参数设置") > 0)
    
    strSQL = "Select 编码,名称 as 类别 from 收费项目类别  Order by 序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Dim objList As ListItem
    With lvwType
        .ListItems.Clear
        Do While Not rsTmp.EOF
            Set objList = .ListItems.Add(, "K" & Nvl(rsTmp!编码), Nvl(rsTmp!类别))
            If str收费类别 = "" Then
                objList.Checked = True
            Else
                objList.Checked = InStr(1, "," & str收费类别 & ",", "," & rsTmp!编码 & ",") > 0
            End If
            rsTmp.MoveNext
        Loop
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim objList As ListItem
    Dim str收费类别 As String
    str收费类别 = ""
    With lvwType
        For Each objList In .ListItems
                If objList.Checked Then
                    str收费类别 = str收费类别 & "," & Mid(objList.Key, 2)
                End If
        Next
    End With
    If str收费类别 <> "" Then str收费类别 = Mid(str收费类别, 2)
    zlDatabase.SetPara "上次过滤收费类别", str收费类别, glngSys, 1120, InStr(1, mstrPrivs, "参数设置") > 0
    If Not mrsDept Is Nothing Then Set mrsDept = Nothing
End Sub
Private Function zlGet收费类别() As String
    Dim objList As ListItem
    mstr收费类别 = ""
    With lvwType
        For Each objList In .ListItems
                If objList.Checked Then
                    mstr收费类别 = mstr收费类别 & "," & Mid(objList.Key, 2)
                End If
        Next
    End With
    If mstr收费类别 <> "" Then mstr收费类别 = Mid(mstr收费类别, 2)

End Function
Private Sub txtNOBegin_Change()
    txtNoEnd.Enabled = Not (Trim(txtNOBegin.Text) = "")
    If Trim(txtNOBegin.Text = "") Then txtNoEnd.Text = ""
End Sub

Private Sub txtNOBegin_GotFocus()
    zlControl.TxtSelAll txtNOBegin
End Sub

Private Sub txtNOBegin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
        '46516
    zlControl.TxtCheckKeyPress txtNOBegin, KeyAscii, m文本式
End Sub

Private Sub txtNOBegin_LostFocus()
    If txtNOBegin.Text <> "" Then txtNOBegin.Text = GetFullNO(txtNOBegin.Text, 13)
End Sub

Private Sub txtNOEnd_LostFocus()
 
    If txtNoEnd.Text <> "" Then txtNoEnd.Text = GetFullNO(txtNoEnd.Text, 13)
End Sub

Private Sub txtNoEnd_GotFocus()
    zlControl.TxtSelAll txtNoEnd
End Sub
Private Sub txtNoEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '46516
    zlControl.TxtCheckKeyPress txtNoEnd, KeyAscii, m文本式
End Sub

Private Sub MakeFilter()

    mstrFilter = " And 登记时间 Between [1] And [2]"
    
    If chk收费.Value = 1 Then
        mblnDateMoved = zlDatabase.DateMoved(Format(IIf(dtpBegin.Value < dtpEnd.Value, dtpBegin.Value, dtpEnd.Value), dtpBegin.CustomFormat), , , Me.Caption)
    Else
        mblnDateMoved = False
    End If
    
    If txtNOBegin.Text <> "" And txtNoEnd.Text <> "" Then
        mstrFilter = mstrFilter & " And NO Between [3] And [4]"
    ElseIf txtNOBegin.Text <> "" Then
        mstrFilter = mstrFilter & " And NO=[3]"
    End If
    
    If cbo操作员.ListIndex <> 0 Then
        mstrFilter = mstrFilter & " And 划价人||''=[5]"
    End If
    
    If txtPatient.Text <> "" And mlngPrePatient <> 0 And Not mrsInfo Is Nothing Then
        If Val(Nvl(mrsInfo!ID)) = mlngPrePatient Then
            mstrFilter = mstrFilter & " And 病人ID=[11]"
        End If
    End If
    
    If txt姓名.Text <> "" Then
        If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(txt姓名.Text, 1))) > 0 Then
            mstrFilter = mstrFilter & " And Upper(姓名) Like [6]"
        Else
            mstrFilter = mstrFilter & " And 姓名 Like [6]"
        End If
    End If
    
    
    If cbo费别.ListIndex <> 0 Then
        mstrFilter = mstrFilter & " And 费别=[7]"
    End If
    
    If cbo科室.ListIndex = 0 Then
        '第一个可能是所有科室
        If Val(cbo科室.ItemData(cbo科室.ListIndex)) > 0 Then
            mstrFilter = mstrFilter & " And 开单部门ID+0=[8]"
        End If
    ElseIf cbo科室.ListIndex > 0 Then
        mstrFilter = mstrFilter & " And 开单部门ID+0=[8]"
    End If
    If mstr收费类别 <> "" Then
        mstrFilter = mstrFilter & " And instr( [10], ','||收费类别||',')>0 "
    End If
End Sub

 


'初始化IDKIND
Private Function InitIDKind() As Boolean
    Dim objCard As Card
    Dim lngCardID As Long
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
    lngCardID = Val(zlDatabase.GetPara("缺省医疗卡类别", glngSys, glngModul, 0))
    If lngCardID <> 0 Then
        IDKind.DefaultCardType = lngCardID
    End If
    Set objCard = IDKind.GetfaultCard
    If IDKind.Cards.按缺省卡查找 And Not objCard Is Nothing Then
        gobjSquare.bln缺省卡号密文 = objCard.卡号密文规则 <> ""
        gobjSquare.int缺省卡号长度 = objCard.卡号长度
        Set gobjSquare.objDefaultCard = objCard
       
    Else
        gobjSquare.bln缺省卡号密文 = IDKind.Cards.加密显示
        gobjSquare.int缺省卡号长度 = 100
    End If
End Function
'获取默认IDKind索引
Private Function IDKindDefaultKind() As Long
    Dim lngIndex As Long
    'IDkind的默认Kind
    If IDKind.DefaultCardType = "" Then
        lngIndex = -1
    Else
        If IsNumeric(IDKind.DefaultCardType) Then
           lngIndex = IDKind.GetKindIndex(IDKind.GetfaultCard.名称)
        Else
           lngIndex = IDKind.GetKindIndex(IDKind.DefaultCardType)
        End If
    End If
    IDKindDefaultKind = lngIndex
End Function

 
'控件名称是否匹配
Private Function IsCardType(ByVal IDKindCtl As IDKindNew, ByVal strCardName As String) As Boolean
    If IDKindCtl Is Nothing Then Exit Function
    If UCase(TypeName(IDKindCtl)) <> "IDKINDNEW" Then Exit Function
    Select Case strCardName
     Case "姓名", "姓名或就诊卡"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "姓名*"
     Case "身份证", "身份证号", "二代身份证"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "*身份证*"
     Case "IC卡号", "IC卡"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "IC卡*"
     Case "医保号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "医保号"
     Case "门诊号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "门诊号"
     Case Else
            If IDKindCtl.GetCurCard Is Nothing Then Exit Function
            If Not IsNumeric(strCardName) Or Val(strCardName) <= 0 Then
                  IsCardType = strCardName = IDKindCtl.GetCurCard.名称
            Else
                If IDKindCtl.GetCurCard.接口序号 <= 0 Then Exit Function
                IsCardType = IDKindCtl.GetCurCard.接口序号 = Val(strCardName)
            End If
     End Select
End Function
                
Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
    '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
    Set gobjSquare.objCurCard = objCard
    '第7位后,就只能用索引,不然取不到数
     
    txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    
    If IDKind.GetCardNoLen <> 0 Then
        txtPatient.MaxLength = IDKind.GetCardNoLen
    Else
        txtPatient.MaxLength = 64
    End If
    '需要清除信息,避免刷卡后,再切换,造成密文显示失去意义
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    If mlngPrePatient Then txtPatient.PasswordChar = ""
    zlControl.TxtSelAll txtPatient
End Sub
Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    If txtPatient.Locked Then Exit Sub
    
    If objCard.名称 Like "IC卡*" And objCard.系统 Then
'        '系统IC卡
'        If Not mobjICCard Is Nothing Then
'           txtPatient.Text = mobjICCard.Read_Card()
'           If txtPatient.Text <> "" Then
'                mblnUnChange = True
'                Call txtPatient_Validate(False)
'                mblnUnChange = False
'           End If
'        End If
        Exit Sub
    End If
    
    lng卡类别ID = objCard.接口序号
    
    If lng卡类别ID <= 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:读卡接口
    '    '入参:frmMain-调用的父窗口
    '    '       lngModule-调用的模块号
    '    '       strExpand-扩展参数,暂无用
    '    '       blnOlnyCardNO-仅仅读取卡号
    '    '出参:strOutCardNO-返回的卡号
    '    '       strOutPatiInforXML-(病人信息返回.XML串)
    '    '返回:函数返回    True:调用成功,False:调用失败\
    If gobjSquare.objSquareCard.zlReadCard(Me, glngModul, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    
    If txtPatient.Text <> "" Then
        mblnUnChange = True
        Call txtPatient_Validate(False)
        mblnUnChange = False
    End If
    
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    Dim lngPreIDKind As Long, intIndex As Integer
    Dim dtDate As Date
    Dim blnNew As Boolean

    If txtPatient.Locked Or txtPatient.Text <> "" Then Exit Sub 'Or Not Me.ActiveControl Is txtPatient
    mblnNotClick = True

    intIndex = IDKind.GetKindIndex(objCard.名称)
    lngPreIDKind = IDKind.IDKind
    If intIndex > 0 Then IDKind.IDKind = intIndex
     
    If IsCardType(IDKind, "身份证") Then
        txtPatient.Text = objPatiInfor.身份证号
    Else
        txtPatient.Text = objPatiInfor.卡号
    End If
    Call txtPatient_KeyPress(vbKeyReturn)
    IDKind.IDKind = lngPreIDKind
    mblnNotClick = False
End Sub
Private Sub GetPatient(ByVal strInput As String, Optional blnCard As Boolean)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取病人信息
    '入参：blnCard=是否就诊卡刷卡
    '出参：
    '返回：
    '编制：刘兴洪
    '日期：2010-07-16 14:24:14
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strTemp As String
    Dim blnSame As Boolean, blnCancel As Boolean
    Dim cur余额 As Currency, curMoney As Currency
    Dim i As Integer, strPati As String
    Dim vRect As RECT, str非在院 As String
    Dim strSQL As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim strTmp As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean
    
    On Error GoTo errH
    mlngPrePatient = 0
    strSQL = ""
    
    If (blnCard Or IDKind.IDKind = IDKindDefaultKind) And InStr("-+*", Left(strInput, 1)) = 0 Then  '103563
        If IDKind.Cards.按缺省卡查找 And Not IDKind.GetfaultCard Is Nothing Then
            lng卡类别ID = IDKind.GetfaultCard.接口序号
        ElseIf IDKind.GetCurCard.接口序号 > 0 Then
            lng卡类别ID = IDKind.GetCurCard.接口序号
        Else
            lng卡类别ID = -1
        End If
        
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
        If lng病人ID <= 0 Then lng病人ID = 0
        strInput = "-" & lng病人ID
        blnHavePassWord = True
        strSQL = strSQL & " And B.病人ID=[2] " & str非在院
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then
        '门诊号
        strSQL = strSQL & " And B.门诊号=[2]" & str非在院
        '75087,冉俊明,2014-7-29,门诊病人收费时,不需要输入完整的门诊号,只需要输入门诊号的最后顺序号即能找到当天就诊的病人信息、费用
        strInput = "*" & zlCommFun.GetFullNO(Mid(strInput, 2), 3)
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then
        '病人ID
        strSQL = strSQL & " And B.病人ID=[2]" & str非在院
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号(病人在院)
        strSQL = strSQL & " And B.住院号=[2]" & str非在院
    Else
        Select Case IDKind.GetCurCard.名称
            Case "姓名", "姓名或就诊卡"
                '姓名
                blnSame = False
                If Not mrsInfo Is Nothing Then
                    If txtPatient.Text = mrsInfo!姓名 Then blnSame = True
                End If
                
                If Not blnSame Then
                    If (Not gblnSeekName) Or (gblnSeekName And Len(strInput) < 2) Then
                        txtPatient.Text = ""
                        Set mrsInfo = Nothing: Exit Sub
                    Else
                       strSQL = strSQL & " And  B.姓名 Like [3]"
                       
                    End If
                Else
                    strSQL = strSQL & " And B.病人ID=[2]"
                    strInput = "-" & Val(mrsInfo!病人ID)
                End If
            Case "医保号"
                strInput = UCase(strInput)
                If mblnOlnyBJYB And zlCommFun.ActualLen(strInput) >= 9 Then
                    '仅北京医保才有效:见问题:问题:26982
                    strSQL = strSQL & " And B.医保号 like [3] " & str非在院
                    strTemp = Left(strInput, 9) & "%"
                Else
                    strSQL = strSQL & " And B.医保号=[1]" & str非在院
                End If
            Case "身份证号", "身份证", "二代身份证"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("身份证", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strSQL = strSQL & " And B.病人ID=[2]" & str非在院
                strInput = "-" & lng病人ID
                ' strSQL = strSQL & " And B.身份证号=[1] " & str非在院
            Case "IC卡号", "IC卡"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("IC卡", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strSQL = strSQL & " And B.病人ID=[2]" & str非在院
                strInput = "-" & lng病人ID
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And B.门诊号=[1]" & str非在院
                '75087,冉俊明,2014-7-29,门诊病人收费时,不需要输入完整的门诊号,只需要输入门诊号的最后顺序号即能找到当天就诊的病人信息、费用
                strInput = zlCommFun.GetFullNO(strInput, 3)
            Case "住院号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strSQL = strSQL & " And B.住院号=[1]" & str非在院
            Case Else
                '其他类别的,获取相关的病人ID
                If Val(IDKind.GetCurCard.接口序号) >= 0 Then
                    lng卡类别ID = Val(IDKind.GetCurCard.接口序号)
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                    If lng病人ID = 0 Then lng病人ID = 0
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(IDKind.GetCurCard.名称, strInput, False, lng病人ID, _
                        strPassWord, strErrMsg) = False Then lng病人ID = 0
                End If
                If lng病人ID <= 0 Then lng病人ID = 0
                strSQL = strSQL & " And B.病人ID=[2]" & str非在院
                strInput = "-" & lng病人ID
                blnHavePassWord = True
        End Select
    End If
    strTmp = strSQL
    strSQL = "    " & vbNewLine & " Select distinct  B.病人id As ID, Decode(sign(nvl(ylkxx.病人id,0)),0,'','√') as 三方账户, B.病人id,B.姓名, B.性别, B.年龄, B.门诊号, B.出生日期, B.身份证号, B.家庭地址, B.工作单位,"
    strSQL = strSQL & vbNewLine & "      A.名称 险类名称"
    strSQL = strSQL & vbNewLine & " From 病人信息 B, 保险类别 A,医疗卡类别 YLK,病人医疗卡信息 YLKXX"
    strSQL = strSQL & vbNewLine & " Where B.险类 = A.序号(+) and b.病人id=ylkxx.病人id(+) and ylkxx.状态(+)=0 and  ylkxx.卡类别id=ylk.id(+)  and ylk.是否自制(+)=0 And B.停用时间 Is Null   "
    strSQL = strSQL & vbNewLine & strTmp
     
    On Error GoTo errH
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strInput, CStr(Mid(strInput, 2)), strInput & "%")
'     vRect = zlcontrol.GetControlRect(txtPatient.hWnd)
'     Set mrsInfo = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "病人查找", 1, "√", "请选择病人", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput, CStr(Mid(strInput, 2)), strInput & "%")
    If Not mrsInfo Is Nothing Then ' And Not blnCancel
       If mrsInfo.RecordCount = 0 Then
            Set mrsInfo = Nothing
            txtPatient.Text = ""
            Exit Sub
        End If
        
        If mrsInfo!ID = 0 Then '没有找到病人信息
            Set mrsInfo = Nothing
            txtPatient.Text = ""
            Exit Sub
        Else '获取到病人信息
        
          txtPatient.Text = Nvl(mrsInfo!姓名)
          Me.txtPatient.Tag = Nvl(mrsInfo!ID)
          mlngPrePatient = Val(Nvl(mrsInfo!ID))
         
        End If
    Else '取消选择
        txtPatient.Text = ""
        Set mrsInfo = Nothing: Exit Sub
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Private Sub txtPatient_Change()
    txtPatient.Tag = "": mlngPrePatient = 0
    If Me.ActiveControl Is txtPatient Then
        IDKind.SetAutoReadCard txtPatient.Text = ""
    End If
   
End Sub


Private Sub txtPatient_GotFocus()
    Call zlControl.TxtSelAll(txtPatient)
    Call zlCommFun.OpenIme(True)
    If txtPatient.Text = "" And ActiveControl Is txtPatient Then IDKind.SetAutoReadCard True
End Sub


Private Sub txtPatient_LostFocus()
    IDKind.SetAutoReadCard False
End Sub

Private Sub txtPatient_Validate(Cancel As Boolean)
    If mblnKeyReturn = False Then
        Call txtPatient_KeyPress(13)
    Else
        mblnKeyReturn = False
    End If
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
  Dim lngID As Long, lngUnit As Long, i As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim strSQL As String, curTotal As Currency
    Dim blnCard As Boolean, blnICCard As Boolean
    
    On Error GoTo errH
    If txtPatient.Locked Then Exit Sub
    mblnKeyReturn = KeyAscii = 13
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    
    If IsCardType(IDKind, "姓名") Then
        '103563,只要输入的第一个字符是“-+*”，后面是全数字，都认为不是刷卡
        If Not (InStr("-+*", Left(txtPatient.Text, 1)) > 0 And IsNumeric(Mid(txtPatient.Text, 2))) Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        End If
    ElseIf IsCardType(IDKind, "门诊号") Or IsCardType(IDKind, "住院号") Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        End If
    End If
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        ElseIf IsNumeric(txtPatient.Tag) Then
            KeyAscii = 0
            'If txtPatient.Tag <> "" Then
            '刷新病人信息:"-病人ID"
            If Val(txtPatient.Tag) <> 0 Then
                Call zlCommFun.PressKey(vbKeyTab)
                Exit Sub
            End If
            Call GetPatient(txtPatient.Tag, False)
            Exit Sub
        End If
        KeyAscii = 0
        If IsCardType(IDKind, "IC卡号") Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
        Call GetPatient(txtPatient.Text, blnCard)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog '
End Sub

