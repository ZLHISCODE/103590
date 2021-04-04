VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Begin VB.Form frmDistFilter 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "分诊过滤"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4080
      TabIndex        =   18
      Top             =   2550
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5205
      TabIndex        =   19
      Top             =   2565
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   2400
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   6345
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   315
         Left            =   1020
         TabIndex        =   22
         Top             =   1890
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   556
         Appearance      =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   12
         FontName        =   "宋体"
         IDKind          =   -1
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         BackColor       =   -2147483633
      End
      Begin VB.TextBox txtValue 
         Height          =   300
         Left            =   1560
         TabIndex        =   17
         ToolTipText     =   "定位F3"
         Top             =   1890
         Width           =   4515
      End
      Begin VB.TextBox txtFactEnd 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3975
         TabIndex        =   12
         Top             =   1094
         Width           =   2085
      End
      Begin VB.TextBox txtFactBegin 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1005
         TabIndex        =   10
         Top             =   1094
         Width           =   2085
      End
      Begin VB.TextBox txtNoEnd 
         Enabled         =   0   'False
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3975
         MaxLength       =   8
         TabIndex        =   8
         Top             =   682
         Width           =   2085
      End
      Begin VB.TextBox txtNOBegin 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1005
         MaxLength       =   8
         TabIndex        =   6
         Top             =   682
         Width           =   2085
      End
      Begin VB.ComboBox cbo操作员 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3975
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1485
         Width           =   2085
      End
      Begin VB.ComboBox cbo科室 
         Height          =   300
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1506
         Width           =   2085
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   300
         Left            =   3975
         TabIndex        =   4
         Top             =   270
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   140378115
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   1005
         TabIndex        =   2
         Top             =   270
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   140378115
         CurrentDate     =   36588
      End
      Begin VB.Label lbl病人 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "病人查找"
         Height          =   180
         Left            =   210
         TabIndex        =   21
         Top             =   1950
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "票据号"
         Height          =   180
         Left            =   405
         TabIndex        =   9
         Top             =   1155
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   3420
         TabIndex        =   11
         Top             =   1155
         Width           =   180
      End
      Begin VB.Label lbl操作员 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "挂号员"
         Height          =   180
         Left            =   3390
         TabIndex        =   15
         Top             =   1545
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "科室"
         Height          =   180
         Left            =   585
         TabIndex        =   13
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "挂号时间"
         Height          =   180
         Left            =   225
         TabIndex        =   1
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "单据号"
         Height          =   180
         Left            =   405
         TabIndex        =   5
         Top             =   735
         Width           =   540
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   3420
         TabIndex        =   7
         Top             =   735
         Width           =   180
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Left            =   3420
         TabIndex        =   3
         Top             =   330
         Width           =   180
      End
   End
   Begin VB.CommandButton cmdDef 
      Caption         =   "缺省(&D)"
      Height          =   350
      Left            =   150
      TabIndex        =   20
      Top             =   2580
      Width           =   1100
   End
   Begin VB.Menu mnuIDKind 
      Caption         =   "身份类别"
      Visible         =   0   'False
      Begin VB.Menu mnuIDKinds 
         Caption         =   ""
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmDistFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Public mlngModul As Long
Public mstrFilter As String
Public mstrSectName As String   '用来指定当前默认的科室
Private mstrPrivs As String
Private mrsDept As ADODB.Recordset  '记录临床科室
Private mrs挂号员 As ADODB.Recordset
Private mcllFiter As Variant       '条件信息
Private mblnOk As Boolean
Private mlngPrePatient As Long
Private mrsInfo As ADODB.Recordset
Private mblnKeyReturn As Boolean
Private mblnOlnyBJYB As Boolean
'-----------------------------------------------------

Public Function zlShowMe(ByVal frmMain As Form, ByVal lngModule As Long, _
    ByRef cllFilter As Variant, ByVal strPrivs As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：程序入口,获取相关条件设置
    '入参：frmMain-主窗体
    '         lngModule-模块号
    '出参：cllFilter-返回相关的条件信息
    '返回：
    '编制：刘兴洪
    '日期：2010-06-02 15:25:35
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    mlngModul = lngModule: Set mcllFiter = cllFilter: mblnOk = False
    mstrPrivs = strPrivs
    Me.Show 1, frmMain
    If mblnOk Then Set cllFilter = mcllFiter
    zlShowMe = mblnOk
End Function

Private Function LoadData() As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：加载基础数据
    '编制：刘兴洪
    '日期：2010-06-02 15:59:42
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim str挂号员 As String, lng科室ID As Long, i  As Long, strTmp As String
    
    If mrs挂号员 Is Nothing Then
        Set mrs挂号员 = GetPersonnel("门诊挂号员", True)
    ElseIf mrs挂号员.State <> 1 Then
        Set mrs挂号员 = GetPersonnel("门诊挂号员", True)
    End If
    If Not mcllFiter Is Nothing Then
        str挂号员 = Trim(mcllFiter("挂号员"))
        lng科室ID = Val(mcllFiter("科室"))
    End If
    '挂号员
    cbo操作员.Clear
    cbo操作员.AddItem "所有挂号员"
    cbo操作员.ListIndex = 0
    If mrs挂号员.RecordCount > 0 Then
        Call mrs挂号员.MoveFirst
        For i = 1 To mrs挂号员.RecordCount
            cbo操作员.AddItem mrs挂号员!简码 & "-" & mrs挂号员!姓名
            If str挂号员 = Nvl(mrs挂号员!姓名) Then cbo操作员.ListIndex = cbo操作员.NewIndex
            mrs挂号员.MoveNext
        Next
    End If
   '读取门诊临床科室，如果已经读取就不再读取
    '143274:李南春,2019/7/26，如果操作员不具有“所有科室”权限，只显示操作员所属科室
    strTmp = Get分诊科室(glngSys, mlngModul, mstrPrivs)
    If strTmp = "" Then strTmp = UserInfo.部门ID
    
    If mrsDept Is Nothing Then
        Set mrsDept = GetDepartments("'临床'", "1,3")
    ElseIf mrsDept.State <> 1 Then
        Set mrsDept = GetDepartments("'临床'", "1,3")
    End If
    
    cbo科室.Clear
    cbo科室.AddItem "所有科室"
    cbo科室.ListIndex = 0
    With mrsDept
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            If InStr(1, "," & strTmp & ",", "," & !id & ",") > 0 Then
                cbo科室.AddItem !编码 & "-" & !名称
                cbo科室.ItemData(cbo科室.NewIndex) = !id
                If lng科室ID = Val(Nvl(!id)) Then cbo科室.ListIndex = cbo科室.NewIndex
            End If
            .MoveNext
        Loop
    End With
    LoadData = True
End Function
Private Sub cbo操作员_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo操作员.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo操作员.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo操作员.ListIndex = lngIdx
    If cbo操作员.ListIndex = -1 And cbo操作员.ListCount <> 0 Then cbo操作员.ListIndex = 0
End Sub

Private Sub cbo科室_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    If SendMessage(cbo科室.Hwnd, CB_GETDROPPEDSTATE, 0, 0) = 0 And KeyAscii <> 27 And KeyAscii <> 13 Then Call zlCommFun.PressKey(vbKeyF4)
    lngIdx = MatchIndex(cbo科室.Hwnd, KeyAscii)
    If lngIdx <> -2 Then cbo科室.ListIndex = lngIdx
    If cbo科室.ListIndex = -1 And cbo科室.ListCount <> 0 Then cbo科室.ListIndex = 0
End Sub

Private Sub cmdCancel_Click()
    mblnOk = False: Unload Me
End Sub

Private Sub cmdDef_Click()
    Dim Curdate As Date
    txtNOBegin.Text = ""
    txtNoEnd.Text = ""
    txtFactBegin.Text = ""
    txtFactEnd.Text = ""
    txtValue.Text = ""
    '当天内
    Curdate = zlDatabase.Currentdate
    dtpBegin.Value = Format(DateAdd("D", -1 * IIf(gSysPara.Sy_Reg.bytNODaysGeneral > gSysPara.Sy_Reg.bytNoDayseMergency, gSysPara.Sy_Reg.bytNODaysGeneral, gSysPara.Sy_Reg.bytNoDayseMergency), Curdate), "yyyy-MM-dd 00:00:00")
    dtpEnd.Value = Format(Curdate, "yyyy-MM-dd 23:59:59")
    mstrFilter = "  And A.发生时间 Between [1] And [2]"
    Set mcllFiter = Nothing
    Call InitCllData
    Call LoadData
End Sub

Private Sub cmdOK_Click()
    If Not IsNull(dtpEnd.Value) Then
        If dtpEnd.Value < dtpBegin.Value Then
            MsgBox "结束时间不能小于开始时间！", vbInformation, gstrSysName
            dtpEnd.SetFocus: Exit Sub
        End If
    End If
    If txtNOBegin.Text <> "" And txtNoEnd.Text <> "" Then
        If txtNoEnd.Text < txtNOBegin.Text Then
            MsgBox "结束单据号不能小于开始单据号！", vbInformation, gstrSysName
            txtNoEnd.SetFocus: Exit Sub
        End If
    End If
    If txtFactBegin.Text <> "" And txtFactEnd.Text <> "" Then
        If txtFactEnd.Text < txtFactBegin.Text Then
            MsgBox "结束票据号不能小于开始票据号！", vbInformation, gstrSysName
            txtFactEnd.SetFocus: Exit Sub
        End If
    End If
    If MakeFilter = False Then Exit Sub
    mblnOk = True: Unload Me
End Sub

Private Sub Form_Activate()
    dtpBegin.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Not ActiveControl Is txtValue Then Call zlCommFun.PressKey(vbKeyTab)
    If KeyCode = vbKeyF3 Then Call txtValue.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '问题号:30346
    If InStr(1, "《》？；：‘|｛｝【】<>?:;|'{}[]", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = 13 And Not ActiveControl Is txtValue Then KeyAscii = 0
End Sub

Public Sub Form_Load()
    Dim Curdate As Date, i As Long
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    
    txtNOBegin.Text = ""
    txtNoEnd.Text = ""
    txtFactBegin.Text = ""
    txtFactEnd.Text = ""
    txtValue.Text = ""
    
    '当天内
    Curdate = zlDatabase.Currentdate
    dtpBegin.Value = Format(DateAdd("D", -1 * IIf(gSysPara.Sy_Reg.bytNODaysGeneral > gSysPara.Sy_Reg.bytNoDayseMergency, gSysPara.Sy_Reg.bytNODaysGeneral, gSysPara.Sy_Reg.bytNoDayseMergency), Curdate), "yyyy-MM-dd 00:00:00")
    dtpEnd.Value = Format(Curdate, "yyyy-MM-dd 23:59:59")
    mstrFilter = "  And A.发生时间 Between [1] And [2]"
    Call InitIDKind
    Call LoadData
    Call InitCllData
End Sub

'初始化IDKIND
Private Function InitIDKind() As Boolean
    Dim objCard As Card, rsTmp As ADODB.Recordset
    Dim lngCardID As Long, strSQL As String
    Call IDKind.zlInit(Me, glngSys, mlngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtValue)
    lngCardID = Val(zlDatabase.GetPara("缺省医疗卡类别", glngSys, mlngModul, 0))
    '72936:刘尔旋,2014-05-13,缺省发卡类型被停用后报错的问题
    If lngCardID <> 0 Then
        strSQL = "Select 1 From 医疗卡类别 Where ID=[1] And Nvl(是否启用,0)=1"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngCardID)
        If Not rsTmp.EOF Then IDKind.DefaultCardType = lngCardID
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

Private Sub Form_Unload(Cancel As Integer)
    Set mrsInfo = Nothing
    mlngPrePatient = 0
    IDKind.SetAutoReadCard False
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand
    Dim strOutPatiInforXML As String
    
    If objCard.名称 Like "IC卡*" And objCard.系统 Then
'        If mobjICCard Is Nothing Then
'            Set mobjICCard = CreateObject("zlICCard.clsICCard")
'            Set mobjICCard.gcnOracle = gcnOracle
'        End If
'        If mobjICCard Is Nothing Then Exit Sub
'        txtValue.Text = mobjICCard.Read_Card()
'        If txtValue.Text <> "" Then
'            Call FindPati(objCard, True, txtValue.Text)
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
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModul, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtValue.Text = strOutCardNO
    If txtValue.Text <> "" Then
        Call FindPati(objCard, True, txtValue.Text)
    End If
End Sub

Private Sub IDKind_ItemClick(index As Integer, objCard As zlIDKind.Card)
    Set gobjSquare.objCurCard = objCard
    If txtValue.Text <> "" Then txtValue.Text = ""
    If txtValue.Enabled And txtValue.Visible Then txtValue.SetFocus
    zlControl.TxtSelAll txtValue
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtValue.Locked Then Exit Sub
    txtValue.Text = objPatiInfor.卡号
    Call FindPati(objCard, True, txtValue.Text)
End Sub

Private Sub txtFactBegin_GotFocus()
    zlControl.TxtSelAll txtFactBegin
End Sub

Private Sub txtFactBegin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFactEnd_GotFocus()
    zlControl.TxtSelAll txtFactEnd
End Sub

Private Sub txtFactEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtFactBegin_Change()
    txtFactEnd.Enabled = Not (Trim(txtFactBegin.Text) = "")
    If Trim(txtFactBegin.Text = "") Then txtFactEnd.Text = ""
End Sub

Private Sub txtNOBegin_Change()
    txtNoEnd.Enabled = Not (Trim(txtNOBegin.Text) = "")
    If Trim(txtNOBegin.Text = "") Then txtNoEnd.Text = ""
End Sub

Private Sub txtNOBegin_GotFocus()
    zlControl.TxtSelAll txtNOBegin
End Sub

Private Sub txtNOBegin_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '46512
    zlControl.TxtCheckKeyPress txtNOBegin, KeyAscii, m文本式

End Sub

Private Sub txtNOBegin_LostFocus()
    If txtNOBegin.Text <> "" Then txtNOBegin.Text = GetFullNO(txtNOBegin.Text, 12)
End Sub

Private Sub txtNOEnd_LostFocus()
    If txtNoEnd.Text <> "" Then txtNoEnd.Text = GetFullNO(txtNoEnd.Text, 12)
End Sub

Private Sub txtNoEnd_GotFocus()
    zlControl.TxtSelAll txtNoEnd
End Sub


Private Sub txtNoEnd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    '46512
   zlControl.TxtCheckKeyPress txtNoEnd, KeyAscii, m文本式
End Sub

Private Function MakeFilter() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定的过滤条件
    '编制:刘兴洪
    '日期:2011-10-21 15:23:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strTmp As String, strSQLtmp As String
    Dim lng病人ID As Long, lng卡类别ID As Long, strErrMsg As String, strPassWord As String
    Dim strKind As String
    Set mcllFiter = New Collection
    mstrFilter = " And A.发生时间 Between [1] And [2]"
    mcllFiter.Add Array(Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS"), Format(dtpEnd.Value, "yyyy-mm-dd HH:MM:SS")), "挂号时间"
    mcllFiter.Add Array(Trim(txtNOBegin.Text), Trim(txtNoEnd)), "挂号NO"
    mcllFiter.Add Array(Trim(txtFactBegin.Text), Trim(txtFactEnd)), "发票号"
    If cbo操作员.ListIndex > 0 Then
        mcllFiter.Add NeedName(cbo操作员.Text), "挂号员"
    Else
        mcllFiter.Add "", "挂号员"
    End If
    mcllFiter.Add "", "科室"
    mcllFiter.Add "", "门诊号": mcllFiter.Add "", "就诊卡号"
    mcllFiter.Add "", "医保号": mcllFiter.Add "", "病人姓名"
    mcllFiter.Add Val(IDKind.IDKind), "KIND"
    mcllFiter.Add "", "病人ID"
    strKind = IDKind.GetCurCard.名称

    mcllFiter.Add Trim(txtValue.Text), "_" & strKind
    If txtNOBegin.Text <> "" And txtNoEnd.Text <> "" Then
        mstrFilter = mstrFilter & " And A.NO Between [3] And [4]"
    ElseIf txtNOBegin.Text <> "" Then
        mstrFilter = mstrFilter & " And A.NO=[3]"
    End If
    
    If cbo操作员.ListIndex > 0 Then mstrFilter = mstrFilter & " And A.操作员姓名||''=[9]"
    If Trim(txtValue.Text) <> "" Then
        If mlngPrePatient <> 0 Then
            mstrFilter = mstrFilter & " And A.病人ID=[12]"
            mcllFiter.Remove "病人ID": mcllFiter.Add mlngPrePatient, "病人ID"
        Else
            Select Case strKind
            Case "门诊号"
                mstrFilter = mstrFilter & " And A.门诊号 = [11]"
                mcllFiter.Remove "门诊号": mcllFiter.Add Trim(txtValue.Text), "门诊号"
            Case "姓名", "姓名或就诊卡"
                If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Left(txtValue.Text, 1))) > 0 Then
                    mstrFilter = mstrFilter & " And Upper(A.姓名) Like [8]"
                Else
                    mstrFilter = mstrFilter & " And A.姓名 Like [8]"
                End If
                mcllFiter.Remove "病人姓名": mcllFiter.Add Trim(txtValue.Text), "病人姓名"
            Case "医保号"
                mstrFilter = mstrFilter & " And B.医保号=[13]"
                mcllFiter.Remove "医保号": mcllFiter.Add Trim(txtValue.Text), "医保号"
            Case Else
                '其他类别的,获取相关的病人ID
                '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
                '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
                '第7位后,就只能用索引,不然取不到数
                lng卡类别ID = Val(IDKind.GetCurCard.接口序号)
                
                If lng卡类别ID <> 0 Then
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, Trim(txtValue.Text), True, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(strKind, Trim(txtValue.Text), True, lng病人ID, _
                        strPassWord, strErrMsg) = False Then lng病人ID = 0
                End If
                If lng病人ID = 0 Then
                    If strErrMsg = "" Then
                        MsgBox "未找到满足条件的病人", vbInformation + vbOKOnly, gstrSysName
                        If txtValue.Enabled And txtValue.Visible Then txtValue.SetFocus
                        zlControl.TxtSelAll txtValue
                        Exit Function
                    End If
                End If
                mstrFilter = mstrFilter & " And A.病人ID=[12]"
                mcllFiter.Remove "病人ID": mcllFiter.Add lng病人ID, "病人ID"
            End Select
        End If
    End If
    
    strSQL = ""
    If (txtFactBegin.Text <> "" And txtFactEnd.Text <> "") Or (txtFactBegin.Text <> "" And txtFactEnd.Text = "") Then
        '无需根据票据号判断,直接根据单据的发生时间判断
        strSQLtmp = IIf(txtFactEnd.Text = "", " =[5] ", " Between [5] And [6] ")
        strSQL = "Select A.NO" & _
        " From 票据打印内容 A,票据使用明细 B" & _
        " Where A.数据性质=4 And A.ID=B.打印ID And B.性质=1" & _
        " And B.号码 " & strSQLtmp
    End If
    If strSQL <> "" Then mstrFilter = mstrFilter & " And A.NO IN(" & strSQL & ")"
    '挂号科室(执行科室)
    If cbo科室.ListIndex > 0 Then
        mstrFilter = mstrFilter & " And A.执行部门ID+0=[7]"
        mcllFiter.Remove "科室"
        mcllFiter.Add cbo科室.ItemData(cbo科室.ListIndex), "科室"
    End If
    mcllFiter.Add mstrFilter, "条件"
    MakeFilter = True
End Function

Private Sub txtValue_Change()
    txtValue.Tag = "": mlngPrePatient = 0
    If Me.ActiveControl Is txtValue Then
        'If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtValue.Text = "")
       ' If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtValue.Text = "")
        IDKind.SetAutoReadCard txtValue.Text = ""
    End If
End Sub

Private Sub txtValue_GotFocus()
    Call zlControl.TxtSelAll(txtValue)
    Call zlCommFun.OpenIme(True)
    If txtValue.Text = "" And ActiveControl Is txtValue Then
'        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtValue.Text = "")
'        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtValue.Text = "")
        IDKind.SetAutoReadCard txtValue.Text = ""
    End If
End Sub

Private Sub txtvalue_LostFocus()
    Call zlCommFun.OpenIme
    IDKind.SetAutoReadCard False
End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
    Dim lngID As Long, lngUnit As Long, i As Long
    Dim rsTmp As ADODB.Recordset, strInfo As String
    Dim strSQL As String, curTotal As Currency
    Dim blnCard As Boolean, blnICCard As Boolean

    On Error GoTo errH
    If txtValue.Locked Then Exit Sub
    mblnKeyReturn = KeyAscii = 13
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub

    If IDKind.GetCurCard.名称 Like "姓名*" Then
        blnCard = zlCommFun.InputIsCard(txtValue, KeyAscii, IDKind.ShowPassText)
    ElseIf IDKind.IDKind = IDKind.GetKindIndex("门诊号") Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not (IsNumeric(Chr(KeyAscii)) Or Chr(KeyAscii) = "-") Then KeyAscii = 0: Exit Sub
        End If
        txtValue.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
    End If
    If blnCard And Len(txtValue.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtValue.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtValue.Text = txtValue.Text & Chr(KeyAscii)
            txtValue.SelStart = Len(txtValue.Text)
        End If
        KeyAscii = 0
        Call FindPati(IDKind.GetCurCard, blnCard, txtValue.Text)
    End If
    If Me.ActiveControl Is txtValue And mblnKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetPatient(ByVal objCard As Card, ByVal strInput As String, Optional blnCard As Boolean)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取病人信息
    '入参：blnCard=是否就诊卡刷卡
    '编制：刘兴洪
    '日期：2010-07-16 14:24:14
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strTemp As String
    Dim blnSame As Boolean, blnCancel As Boolean
    Dim cur余额 As Currency, curMoney As Currency
    Dim i As Integer, strPati As String
    Dim vRect As RECT
    Dim strSQL As String, lng卡类别ID As Long, strPassWord As String, strErrMsg As String
    Dim strTmp As String
    Dim lng病人ID As Long, blnHavePassWord As Boolean

    On Error GoTo errH

    strSQL = ""
    If blnCard = True And objCard.名称 Like "姓名*" Then    '刷卡
        If IDKind.Cards.按缺省卡查找 And Not IDKind.GetfaultCard Is Nothing Then
            lng卡类别ID = IDKind.GetfaultCard.接口序号
        Else
            lng卡类别ID = "-1"
        End If
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        strInput = "-" & lng病人ID
        blnHavePassWord = True
        strSQL = strSQL & " And B.病人ID=[2] "
        
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then
        '门诊号
        strSQL = strSQL & " And B.门诊号=[2]"
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then
        '病人ID
        strSQL = strSQL & " And B.病人ID=[2]"
    Else
        Select Case objCard.名称
        Case "姓名", "姓名或就诊卡"
            txtValue.Tag = strInput
            Set mrsInfo = Nothing: Exit Sub
            zlCommFun.PressKey vbKeyTab
        Case "医保号"
            strInput = UCase(strInput)
            If mblnOlnyBJYB And zlCommFun.ActualLen(strInput) >= 9 Then
                '仅北京医保才有效:见问题:问题:26982
                strSQL = strSQL & " And B.医保号 like [3] "
                strTemp = Left(strInput, 9) & "%"
            Else
                strSQL = strSQL & " And B.医保号=[1]"
            End If
        Case "身份证号", "身份证", "二代身份证"
            strInput = UCase(strInput)
            If gobjSquare.objSquareCard.zlGetPatiID("身份证", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
            strSQL = strSQL & " And B.病人ID=[2]"
            strInput = "-" & lng病人ID
        Case "IC卡号", "IC卡"
            strInput = UCase(strInput)
            If gobjSquare.objSquareCard.zlGetPatiID("IC卡", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
            strSQL = strSQL & " And B.病人ID=[2]"
            strInput = "-" & lng病人ID
        Case "门诊号"
            If Not IsNumeric(strInput) Then strInput = "0"
            strSQL = strSQL & " And B.门诊号=[1]"
        Case Else
            '其他类别的,获取相关的病人ID
            If Val(objCard.接口序号) > 0 Then
                lng卡类别ID = Val(objCard.接口序号)
                If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                If lng病人ID = 0 Then lng病人ID = 0
            Else
                If gobjSquare.objSquareCard.zlGetPatiID(objCard.名称, strInput, False, lng病人ID, _
                                                        strPassWord, strErrMsg) = False Then lng病人ID = 0
            End If
            If lng病人ID <= 0 Then lng病人ID = 0
            strSQL = strSQL & " And B.病人ID=[2]"
            strInput = "-" & lng病人ID
            blnHavePassWord = True
        End Select
    End If
    strSQL = "" & _
    "   Select distinct  B.病人id As ID, Decode(sign(nvl(X.病人id,0)),0,'','√') as 三方账户,  " & _
    "           B.病人id,B.姓名, B.性别, B.年龄, B.门诊号, B.出生日期, B.身份证号, B.家庭地址, B.工作单位," & _
    "            A.名称 险类名称" & _
    "   From 病人信息 B, 保险类别 A,医疗卡类别 Y,病人医疗卡信息 X" & _
    "   Where B.险类 = A.序号(+) and b.病人id=X.病人id(+)  " & _
    "               And X.状态(+)=0 and  X.卡类别id=Y.id(+)  and Y.是否自制(+)=0 And B.停用时间 Is Null   " & _
                    strSQL
    On Error GoTo errH
    vRect = zlControl.GetControlRect(txtValue.Hwnd)
    Set mrsInfo = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "病人查找", 1, "√", "请选择病人", False, False, True, vRect.Left, vRect.Top, txtValue.Height, blnCancel, False, True, strInput, CStr(Mid(strInput, 2)), strInput & "%", dtpBegin.Value, dtpEnd.Value)
    
    If blnCancel Or mrsInfo Is Nothing Then
        Set mrsInfo = Nothing: txtValue.Text = "": Exit Sub
    End If
    
    If mrsInfo!id = 0 Then    '没有找到病人信息
        Set mrsInfo = Nothing: txtValue.Text = "": Exit Sub
    End If
    
    txtValue.MaxLength = zlGetPatiInforMaxLen.intPatiName
    txtValue.Text = Nvl(mrsInfo!姓名)
    Me.txtValue.Tag = Nvl(mrsInfo!id)
    mlngPrePatient = Val(Nvl(mrsInfo!id))
    zlCommFun.PressKey vbKeyTab
    Exit Sub
    
NotFoundPati:
    Set mrsInfo = Nothing: txtValue.Text = "": Exit Sub
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查找病人
    '编制:刘兴洪
    '日期:2012-08-29 17:53:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnICCard As Boolean, blnIDCard As Boolean
   '读取病人信息
    Call GetPatient(objCard, txtValue.Text, blnCard)
End Sub

Private Sub InitCllData()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：初始化集合数据
    '编制：刘兴洪
    '日期：2010-06-02 15:44:19
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    If mcllFiter Is Nothing Then
        Set mcllFiter = New Collection
        mcllFiter.Add Array(Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS"), Format(dtpEnd.Value, "yyyy-mm-dd HH:MM:SS")), "挂号时间"
        mcllFiter.Add Array("", ""), "挂号NO"
        mcllFiter.Add Array("", ""), "发票号"
        mcllFiter.Add "", "挂号员"
        mcllFiter.Add "", "科室"
        mcllFiter.Add "", "门诊号": mcllFiter.Add "", "就诊卡号"
        mcllFiter.Add "", "医保号": mcllFiter.Add "", "病人姓名"
        mcllFiter.Add 0, "KIND"
        mcllFiter.Add mstrFilter, "条件"
        Exit Sub
    End If
    '恢复默认数据
    txtNOBegin.Text = mcllFiter("挂号NO")(0):    txtNoEnd.Text = mcllFiter("挂号NO")(1)
    txtFactBegin.Text = mcllFiter("发票号")(0):    txtFactEnd.Text = mcllFiter("发票号")(1)
    dtpBegin.Value = CDate(mcllFiter("挂号时间")(0)):    dtpEnd.Value = CDate(mcllFiter("挂号时间")(1))
    mstrFilter = CStr(mcllFiter("条件"))

    '集何中可能不存在,所以不加载值
    Err = 0: On Error Resume Next
    If mcllFiter(Trim(IDKind.GetCurCard.名称)) <> "" Then
        '初始化
        txtValue.Text = mcllFiter("_" & Trim(IDKind.GetCurCard.名称))
    End If
End Sub
