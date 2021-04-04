VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.10#0"; "zlIDKind.ocx"
Begin VB.Form frm卫材发放过滤 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1515
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11805
   LinkTopic       =   "Form1"
   ScaleHeight     =   1515
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra 
      Height          =   1380
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11865
      Begin VB.ComboBox cbo发料部门 
         Height          =   300
         Left            =   900
         TabIndex        =   16
         Text            =   "cbo发料部门"
         Top             =   975
         Width           =   4575
      End
      Begin VB.PictureBox picFilter 
         BorderStyle     =   0  'None
         Height          =   825
         Left            =   5850
         ScaleHeight     =   825
         ScaleWidth      =   5940
         TabIndex        =   7
         Top             =   120
         Width           =   5940
         Begin VB.CommandButton cmdIC 
            Caption         =   "读卡"
            Height          =   300
            Left            =   5280
            TabIndex        =   11
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox txtPati 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2820
            TabIndex        =   10
            Top             =   30
            Width           =   2520
         End
         Begin VB.CommandButton cmd科室 
            Caption         =   "…"
            Height          =   255
            Left            =   5040
            TabIndex        =   9
            Top             =   398
            Width           =   285
         End
         Begin VB.TextBox txt科室 
            Height          =   300
            Left            =   2820
            TabIndex        =   8
            Top             =   375
            Width           =   2520
         End
         Begin MSComctlLib.TabStrip tbsType 
            Height          =   255
            Left            =   780
            TabIndex        =   12
            Top             =   405
            Width           =   2205
            _ExtentX        =   3889
            _ExtentY        =   450
            MultiRow        =   -1  'True
            Style           =   2
            HotTracking     =   -1  'True
            Separators      =   -1  'True
            _Version        =   393216
            BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
               NumTabs         =   3
               BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "临床"
                  Key             =   "T1"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "医技"
                  Key             =   "T2"
                  ImageVarType    =   2
               EndProperty
               BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
                  Caption         =   "病区"
                  Key             =   "T3"
                  ImageVarType    =   2
               EndProperty
            EndProperty
         End
         Begin zlIDKind.IDKindNew IDKNType 
            Height          =   375
            Left            =   1800
            TabIndex        =   13
            Top             =   0
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   661
            ShowSortName    =   0   'False
            IDKindStr       =   "住|住院号|0|0|0|0|0|;床|床号|0|0|0|0|0|;姓|姓名|0|0|0|0|0|;病|病人id|0|0|0|0|0|;门|门诊号|0|0|0|0|0|;IC|IC卡号|1|0|0|0|0|"
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
            DefaultCardType =   "0"
            AutoSize        =   -1  'True
            AllowAutoICCard =   -1  'True
            AllowAutoCommCard=   0   'False
            BackColor       =   -2147483644
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "病人信息"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   840
            TabIndex        =   15
            Top             =   90
            Width           =   720
         End
         Begin VB.Label lbl部门类型 
            AutoSize        =   -1  'True
            Caption         =   "部门类型"
            Height          =   180
            Left            =   0
            TabIndex        =   14
            Top             =   435
            Width           =   720
         End
      End
      Begin VB.CommandButton cmd刷新 
         Caption         =   "刷新(&R)"
         Height          =   350
         Left            =   10665
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   960
         Width           =   1100
      End
      Begin VB.TextBox txtEDIT 
         Height          =   300
         Index           =   0
         Left            =   900
         MaxLength       =   8
         TabIndex        =   5
         Top             =   585
         Width           =   2085
      End
      Begin VB.TextBox txtEDIT 
         Height          =   300
         Index           =   1
         Left            =   3405
         MaxLength       =   8
         TabIndex        =   4
         Top             =   585
         Width           =   2085
      End
      Begin VB.Frame fra 
         BorderStyle     =   0  'None
         Height          =   420
         Index           =   1
         Left            =   5850
         TabIndex        =   1
         Top             =   885
         Width           =   4800
         Begin VB.CheckBox chkType 
            Caption         =   "住院"
            Height          =   240
            Index           =   1
            Left            =   960
            TabIndex        =   3
            Top             =   150
            Value           =   1  'Checked
            Width           =   2145
         End
         Begin VB.CheckBox chkType 
            Caption         =   "门诊"
            Height          =   240
            Index           =   0
            Left            =   105
            TabIndex        =   2
            Top             =   150
            Value           =   1  'Checked
            Width           =   885
         End
      End
      Begin MSComCtl2.DTPicker Dtp开始Date 
         Height          =   300
         Left            =   900
         TabIndex        =   17
         Top             =   210
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   80740355
         CurrentDate     =   37007
      End
      Begin MSComCtl2.DTPicker Dtp结束Date 
         Height          =   300
         Left            =   3405
         TabIndex        =   18
         Top             =   210
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   80740355
         CurrentDate     =   37007
      End
      Begin VB.Label lblCon 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "发料部门"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   4
         Left            =   135
         TabIndex        =   23
         Top             =   1020
         Width           =   720
      End
      Begin VB.Label lblCon 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "～"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   3
         Left            =   3120
         TabIndex        =   22
         Top             =   645
         Width           =   180
      End
      Begin VB.Label lblCon 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "单据范围"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   645
         Width           =   720
      End
      Begin VB.Label lblCon 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "时间范围"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   270
         Width           =   720
      End
      Begin VB.Label lblCon 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "～"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   3120
         TabIndex        =   19
         Top             =   270
         Width           =   180
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm卫材发放过滤.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm卫材发放过滤.frx":031A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm卫材发放过滤"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mArrFilter As Variant
Private mintType As Integer
Private mstrPrivs As String
Private mlngModule As Long
Private mblnCard As Boolean     '是否刷的是就诊卡
Private mobjcard As Card
Private mintOld输入模式 As Integer

Private Enum mFindType
    住院号 = 0
    床号 = 1
    姓名 = 2
    病人ID = 3
    门诊号 = 4
    IC卡号 = 5
End Enum

Private Enum mtxtIdx
    idx_开始NO = 0
    idx_结束NO = 1
End Enum

Private mblnDrop As Boolean                     '在KeyDown中判断下拉列表是否弹出

Private Const CB_GETDROPPEDSTATE = &H157
Private Const CB_SHOWDROPDOWN = &H14F

Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1

'--------------------------------------------------------------------------------------------------------
'药品发药传入
Private mblnTrans As Boolean            'True表示从药品处方发药窗口调用
Private mstrNo  As String               '单据号，仅用于定位
Private mlng库房id As Long              '发药库房ID，一般和发料部门一致
Private mstrDrugStartDate As String     '药品单据开始时间
Private mstrDrugEndDate As String       '药品单据结束时间
Private mlng病人id As Long
Private mlngPre部门ID As Long
Private mblnNoClick As Boolean

'--------------------------------------------------------------------------------------------------------
Public Event zlRefreshCon(ByVal arrFilter As Variant)
Public Event zlPopupMenus(ByVal x As Long, ByVal Y As Long)

Private Sub InitIDKindNew()
    Dim int输入模式 As Integer
    Dim strTemp As String

    strTemp = "住|住院号|0;床|床号|0;姓|姓名|0;病|病人id|0;门|门诊号|0;IC|IC卡号|1"
    Me.IDKNType.IDKindStr = strTemp
    Call IDKNType.zlInit(Me, glngSys, mlngModule, gcnOracle, gstrDBUser, gobjSquareCard, strTemp, txtPati)
    IDKNType.SetAutoReadCard True
    Me.IDKNType.IDKind = 0
    
End Sub

Private Sub chkType_Click(Index As Integer)
    Call cmd刷新_Click
End Sub

Private Sub IDKNType_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    Set mobjcard = objCard
    mintType = Index - 1
    mintOld输入模式 = mintType
    
    txtPati.Text = ""
    txtPati.MaxLength = objCard.卡号长度
    If objCard.卡号密文规则 <> "" Then
        txtPati.PasswordChar = "*"
    Else
        txtPati.PasswordChar = ""
    End If
    
End Sub

Private Sub IDKNType_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    txtPati.Text = objPatiInfor.卡号
    If txtPati.Text <> "" Then Call txtPati_KeyPress(vbKeyReturn)
End Sub


Private Sub mobjICCard_ShowICCardInfo(ByVal strNo As String)
    If Not txtPati.Locked And txtPati.Text = "" And Me.ActiveControl Is txtPati And strNo <> "" Then
        txtPati.Text = strNo

        If txtPati.Text = "" Then
            Call mobjICCard.SetEnabled(False)
        Else
'            Me.PatiTittle = 6

            Call txtPati_KeyDown(vbKeyReturn, 0)
        End If
    End If
End Sub
Private Function CheckDepend() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:检查数据依赖性
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New Recordset
    Dim lng发料部门ID As Long
    
    On Error GoTo ErrHandle
    CheckDepend = False
    
    gstrSQL = "" & _
        "   SELECT DISTINCT a.id, a.简码 || '-' || a.名称 As 名称 " & _
        "   FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " & _
        "   Where c.工作性质 = b.名称 And (a.站点=[2] or a.站点 is null) " & _
        "       AND b.编码 ='W' " & _
        "       AND a.id = c.部门id " & _
        "       AND TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'" & _
        IIf(InStr(mstrPrivs, "所有部门") <> 0, "", " And a.ID IN (Select 部门ID From 部门人员 Where 人员ID=[1])") & _
        " Order by a.简码 || '-' || a.名称"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取相应的库房", UserInfo.Id, gstrNodeNo)
    
    If rsTemp.EOF Then
        rsTemp.Close
        Exit Function
    End If
    
    '如果是药品窗口传入，设置发料部门与药品发药部门一致
    If mblnTrans Then
        If mlng库房id <> UserInfo.部门ID Then
            lng发料部门ID = mlng库房id
        Else
            lng发料部门ID = UserInfo.部门ID
        End If
    End If
    '装入发料部门数据
    With cbo发料部门
        .Clear
        mblnNoClick = True
        Do While Not rsTemp.EOF
            .AddItem rsTemp!名称
            .ItemData(.NewIndex) = rsTemp!Id
            If rsTemp!Id = lng发料部门ID Then
                .ListIndex = .NewIndex
                mlngPre部门ID = lng发料部门ID
            End If
            rsTemp.MoveNext
        Loop
        If .ListIndex = -1 Then .ListIndex = 0: mlngPre部门ID = .ItemData(.ListIndex)
        mblnNoClick = False
        rsTemp.Close
    End With
    CheckDepend = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetFilter() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取条件信息
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-04-30 11:52:50
    '-----------------------------------------------------------------------------------------------------------
    Dim cllFilter As Collection, strReg As String
    Dim int收费处方 As Integer
    Dim lng病人id As Long
    Dim strCard As String

    
    strReg = Trim(zlDatabase.GetPara("查询业务类型", glngSys, mlngModule, ""))
    If strReg = "" Then strReg = "24,25,26"
    
    strCard = Split(Split(IDKNType.IDKindStr, ";")(mintType), "|")(1)
    '基本查询条件
    Set cllFilter = New Collection
    
    int收费处方 = Val(zlDatabase.GetPara("收费处方显示方式", glngSys, mlngModule, 0))
    
    If int收费处方 < 0 Or int收费处方 > 2 Then
        int收费处方 = 0
    End If
    
    cllFilter.Add int收费处方, "收费处方"
    
    cllFilter.Add txtPati.Text, "内容"
    
    If cbo发料部门.ListIndex < 0 Then
        cllFilter.Add 0, "发料部门ID"
    Else
        cllFilter.Add cbo发料部门.ItemData(cbo发料部门.ListIndex), "发料部门ID"
    End If
    cllFilter.Add Array(Format(Dtp开始Date.Value, "yyyy-mm-dd HH:MM:SS"), Format(Dtp结束Date.Value, "yyyy-mm-dd HH:MM:SS")), "日期范围"
    cllFilter.Add strReg, "单据"
    If Trim(txt科室.Tag) = "" Then
        cllFilter.Add "", "开单科室ID"
    Else
        cllFilter.Add Trim(txt科室.Tag), "开单科室ID"
    End If
    
    If tbsType.SelectedItem Is Nothing Then
        cllFilter.Add 0, "部门类型"
    Else
        cllFilter.Add tbsType.SelectedItem.Index - 1, "部门类型"
    End If
   
    cllFilter.Add Array(Trim(txtEDIT(mtxtIdx.idx_开始NO)), Trim(txtEDIT(mtxtIdx.idx_结束NO))), "单据号"
    If strCard = "住院号" Then
        cllFilter.Add Val(txtPati.Text), "住院号"
    Else
        cllFilter.Add 0, "住院号"
    End If
    If strCard = "姓名" Then
        If mblnCard = True Then
            cllFilter.Add Trim(txtPati.Text), "就诊卡号"
            cllFilter.Add "", "姓名"
        Else
            cllFilter.Add Trim(txtPati.Text), "姓名"
        End If
    Else
        cllFilter.Add "", "姓名"
    End If
    
    If strCard = "床号" Then
        cllFilter.Add Trim(txtPati.Text), "床号"
    Else
        cllFilter.Add "", "床号"
    End If
    If strCard = "病人id" Then
        cllFilter.Add Val(txtPati.Text), "病人ID"
        cllFilter.Add 0, "IC卡号"
    ElseIf strCard = "IC卡" Then
        If Not gobjSquareCard Is Nothing Then
            Call gobjSquareCard.zlGetPatiID("IC卡", txtPati.Text, True, lng病人id)
        End If
        cllFilter.Add lng病人id, "病人ID"
        If txtPati.Text <> "" Then
            cllFilter.Add 1, "IC卡号"
        Else
            cllFilter.Add 0, "IC卡号"
        End If
    Else
        '银行卡
        If Not gobjSquareCard Is Nothing And strCard <> "姓名" And strCard <> "床号" And strCard <> "住院号" And strCard <> "门诊号" Then
            If gobjSquareCard.zlGetPatiID(mobjcard.接口序号, txtPati.Text, False, lng病人id) = False And txtPati.Text <> "" Then lng病人id = -1
        End If
        cllFilter.Add lng病人id, "病人ID"
        cllFilter.Add 0, "IC卡号"
    End If
    
    If strCard = "门诊号" Then
        cllFilter.Add Val(txtPati.Text), "门诊号"
    Else
        cllFilter.Add "", "门诊号"
    End If
    
    If Not (strCard = "姓名" And mblnCard) Then
        cllFilter.Add "", "就诊卡号"
    End If
    
    If (chkType(0).Value = 1 And chkType(1).Value = 1) Or (chkType(0).Value = 0 And chkType(1).Value = 0) Then
        '0-所有
        cllFilter.Add 0, "请求类型"
    ElseIf chkType(0).Value = 1 Then
        '1-门诊及记帐
        cllFilter.Add 1, "请求类型"
    ElseIf chkType(1).Value = 1 Then
        '2-住院记帐单
        cllFilter.Add 2, "请求类型"
    End If
    
    'zlDatabase.OpenSQLRecord(gstrsql, Me.Caption, _
        Val(mArrFilter("发料部门ID")), _
        CDate(mArrFilter("日期范围")(0)), CDate(mArrFilter("日期范围")(1)), _
        CStr("," & mArrFilter("单据") & ","), _
        Val(mArrFilter("开单科室ID")), _
        CStr(mArrFilter("单据号")(0)), CStr(mArrFilter("单据号")(1)), _
        Val(mArrFilter("病人ID")), Val(mArrFilter("住院号")), _
        CStr(mArrFilter("姓名")))
        
    Set mArrFilter = cllFilter
    
End Function

Private Sub cbo发料部门_Click()
    If cbo发料部门.ListIndex < 0 Then Exit Sub
    If mblnNoClick = True Then Exit Sub
    
    If mlngPre部门ID <> cbo发料部门.ItemData(cbo发料部门.ListIndex) Then
        mlngPre部门ID = cbo发料部门.ItemData(cbo发料部门.ListIndex)
        Call cmd刷新_Click
    End If
End Sub

Private Sub cbo发料部门_KeyDown(KeyCode As Integer, Shift As Integer)
'    mblnDrop = False
'    If KeyCode = 13 Then mblnDrop = SendMessage(cbo发料部门.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 1
'
''    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab

    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If cbo发料部门.ListIndex >= 0 Then
        If mlngPre部门ID <> cbo发料部门.ItemData(cbo发料部门.ListIndex) Then
            mlngPre部门ID = cbo发料部门.ItemData(cbo发料部门.ListIndex)
            Call cmd刷新_Click
        End If
    End If
    
    If Select部门选择器(Me, cbo发料部门, Trim(cbo发料部门.Text), "W", False) = False Then
        DoEvents
        
        cbo发料部门.SetFocus
        Exit Sub
    End If
    
    If cbo发料部门.ListIndex >= 0 Then
        If mlngPre部门ID <> cbo发料部门.ItemData(cbo发料部门.ListIndex) Then
            mlngPre部门ID = cbo发料部门.ItemData(cbo发料部门.ListIndex)
            Call cmd刷新_Click
        End If
    End If
End Sub

Private Sub cbo发料部门_KeyPress(KeyAscii As Integer)
'    Dim i As Long, intIdx As Integer
'    Dim strText As String, strResult As String, strFilter As String
'
'    If KeyAscii = 13 Then
'        strText = UCase(cbo发料部门.Text)
'        If cbo发料部门.ListIndex <> -1 Then
'            '弹出列表时,又在文本框输入了内容
'            If strText <> cbo发料部门.List(cbo发料部门.ListIndex) Then Call zlControl.CboSetIndex(cbo发料部门.hwnd, -1)
'        End If
'        If strText = "" Then
'            cbo发料部门.ListIndex = -1
'        ElseIf cbo发料部门.ListIndex = -1 Then
'            intIdx = -1
'
'            For i = 1 To cbo发料部门.ListCount - 1
'                If Mid(cbo发料部门.List(i), 1, InStr(1, cbo发料部门.List(i), "-") - 1) = strText _
'                    Or Mid(cbo发料部门.List(i), InStr(1, cbo发料部门.List(i), "-")) = strText Then
'                    intIdx = i
'                    Exit For
'                End If
'            Next
'
'            If intIdx = -1 Then
'                For i = 1 To cbo发料部门.ListCount - 1
'                    If UCase(cbo发料部门.List(i)) Like strText & "*" Then
'                        intIdx = i
'                    End If
'                Next
'            End If
'
'            cbo发料部门.ListIndex = intIdx
'            SendMessage cbo发料部门.hwnd, CB_SHOWDROPDOWN, True, 0
'        ElseIf Not mblnDrop Then
'            '回车光标经过
'            Call cbo发料部门_Click
'            Exit Sub
'        End If
'        If cbo发料部门.ListIndex = -1 Then
'            cbo发料部门.ListIndex = 0
'        Else
'            If intIdx <> -1 And mblnDrop Then
'                '弹出回车-强行激活Click
'                Call cbo发料部门_Click
'            ElseIf intIdx <> cbo发料部门.ListIndex And intIdx <> -1 Then
'                '弹出让选择-自动激活Click
'                cbo发料部门.SetFocus
'                Exit Sub
'            ElseIf intIdx <> -1 Then
'                '一次性输中-强行激活Click
'                Call cbo发料部门_Click
'            End If
'        End If
'    End If
End Sub


Private Sub cbo发料部门_Validate(Cancel As Boolean)
'    Dim i As Long
'    Dim blnTmp As Boolean
'
'    If cbo发料部门.Text = "" Then
'        cbo发料部门.ListIndex = 0
'    Else
'        For i = 0 To cbo发料部门.ListCount - 1
'            If cbo发料部门.Text = cbo发料部门.List(i) Then
'                blnTmp = True
'                Exit For
'            End If
'        Next
'
'        If blnTmp = False Then
'            cbo发料部门.ListIndex = 0
'        End If
'    End If
End Sub

Private Sub chkType_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub cmdIC_Click()
    Dim strOutXML As String
    Dim strTemp As String
    Dim strCard As String
    
    strCard = Split(Split(IDKNType.IDKindStr, ";")(mintType), "|")(1)
    If strCard = "IC卡号" Then
        If mobjICCard Is Nothing Then
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            Set mobjICCard.gcnOracle = gcnOracle
        End If
        If Not mobjICCard Is Nothing Then
            txtPati.Text = mobjICCard.Read_Card()
            If txtPati.Text <> "" Then Call cmd刷新_Click
        End If
    Else
        If Not gobjSquareCard Is Nothing Then
            Call gobjSquareCard.zlReadCard(Me, mlngModule, mobjcard.接口序号, True, "", strTemp, strOutXML)
            txtPati.Text = strTemp
            If txtPati.Text <> "" Then Call txtPati_KeyPress(vbKeyReturn)
        End If
    End If
End Sub

Private Sub cmd科室_Click()
    If Select部门类型(txt科室, "") = False Then Exit Sub
    Call InitData
End Sub

Private Sub cmd刷新_Click()
    Call GetFilter
    RaiseEvent zlRefreshCon(mArrFilter)
End Sub
Private Sub InitData()
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化数据
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-05-01 21:55:55
    '-----------------------------------------------------------------------------------------------------------
   
    Dtp结束Date.MaxDate = Format(sys.Currentdate, "yyyy-mm-dd") & " 23:59:59"
    If mblnTrans Then
        Dtp开始Date.Value = CDate(mstrDrugStartDate)
        Dtp结束Date.Value = CDate(mstrDrugEndDate)
    Else
         Dtp结束Date.Value = Dtp结束Date.MaxDate
         Dtp开始Date.Value = Format(DateAdd("d", -7, sys.Currentdate), "yyyy-mm-dd") & " 00:00:00"
    End If
    Dtp开始Date.MaxDate = Dtp结束Date.MaxDate
'    txtEDIT(mtxtIdx.idx_开始NO) = mstrNo

    
End Sub


Private Sub Dtp结束Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab

End Sub

Private Sub Dtp开始Date_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    mstrPrivs = gstrPrivs:    mlngModule = glngModul: mblnCard = False
    
    Set mobjICCard = New clsICCard
    Call mobjICCard.SetParent(Me.hwnd)
    Set mobjICCard.gcnOracle = gcnOracle
    
    Call InitIDKindNew
    
    Call CheckDepend
    
    cmdIC.Visible = False
End Sub

Private Sub Form_Resize()
    Dim sngTemp As Single
    
    On Error Resume Next
    
    With fra(0)
        .Top = ScaleTop
        .Height = ScaleHeight
        .Left = ScaleLeft
        .Width = ScaleWidth
    End With
    
    With picFilter
        .Width = IIf(ScaleWidth - .Left - 50 < 0, 0, ScaleWidth - .Left - 50)
        cmd刷新.Left = .Left + .Width - cmd刷新.Width - 50
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '卸载一卡通接口
    gstrCardType = ""
    Set gobjSquareCard = Nothing
    
    '卸载IC卡刷卡接口
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
End Sub


Private Sub picFilter_Resize()
    err = 0: On Error Resume Next
    With txtPati
        .Width = picFilter.ScaleWidth - .Left
        txt科室.Width = .Width
        cmd科室.Left = .Left + .Width - cmd科室.Width - 10
        cmdIC.Left = picFilter.Width - cmdIC.Width - 20
    End With
End Sub
Private Sub lblPatiInputType_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton Then Exit Sub
    RaiseEvent zlPopupMenus(x, Y)
End Sub
Public Property Get PatiTittle() As Integer
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取病人信息的相关标题
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-01 12:49:27
    '-----------------------------------------------------------------------------------------------------------
    PatiTittle = mintType
End Property

Public Property Get PatiCardID() As Long
    '获取消费卡的类别ID
    If mintType > 5 Then
        PatiCardID = mobjcard.接口序号
    Else
        PatiCardID = 0
    End If
End Property
'Public Property Let PatiTittle(ByVal vNewValue As Integer)
'
'    mintType = vNewValue
'
'    If mintType <= 5 Then
'        Me.lblPatiInputType.Caption = Decode(mintType, mFindType.住院号, "住院号↓", mFindType.姓名, "姓  名↓", mFindType.床号, " 床  号↓", mFindType.病人ID, "病人ID↓", mFindType.门诊号, "门诊号↓", mFindType.IC卡号, "IC卡号")
'    Else
'        '银行卡
'        If gstrCardType <> "" Then
'            Me.lblPatiInputType.Caption = Split(Split(gstrCardType, ";")(mintType - 6), "|")(1) & "↓"
'        End If
'    End If
'
'    '明确为就诊卡的输入框:
'    '以前有个渠道提出来的，就诊卡中含特殊符号，要屏蔽：" :：;；?？"。
'    '对汉字，将输入框的ImeMode设置为Disable就可以了。
'    txtPati.IMEMode = 0
'    cmdIC.Visible = False
'    If mintType = mFindType.病人ID Or mintType = mFindType.门诊号 Or mintType = mFindType.姓名 Or mintType = mFindType.住院号 Then
'        txtPati.MaxLength = 18
'    ElseIf mintType = mFindType.IC卡号 Then
'        cmdIC.Visible = True
'    ElseIf mintType > 5 Then
'        '银行卡
'        txtPati.Tag = Split(gstrCardType, ";")(mintType - 6)
'        txtPati.MaxLength = Val(Split(txtPati.Tag, "|")(gCardFormat.卡号长度))
'        cmdIC.Visible = (Val(Split(txtPati.Tag, "|")(gCardFormat.刷卡标志)) = 1)
'    Else
'        txtPati.MaxLength = 0
'    End If
'
'End Property
Private Function Select部门类型(ByVal objCtl As Control, ByVal strSearch As String) As Boolean
    Dim blnCancel As Boolean, strKey As String, strTittle As String, lngH As Long, strTemp As String
    Dim vRect As RECT
    Dim rsTemp  As ADODB.Recordset
    Dim rsCount As ADODB.Recordset
    Dim strSelectSql As String
    Dim strName As String
    Dim int类型 As Integer '0-所有;1-门诊;2-住院
    
    On Error GoTo ErrHandle
    If (chkType(0).Value = 1 And chkType(1).Value = 1) Or (chkType(0).Value = 0 And chkType(1).Value = 0) Then
        '0-所有
        int类型 = 0
    ElseIf chkType(0).Value = 1 Then
        '1-门诊及记帐
        int类型 = 1
    ElseIf chkType(1).Value = 1 Then
        '2-住院记帐单
        int类型 = 2
    End If
    
    strKey = GetMatchingSting(UCase(strSearch), False)
    
    strTittle = "部门选择器"
    vRect = zlControl.GetControlRect(objCtl.hwnd)
    lngH = objCtl.Height
    
    If frm卫材发放管理_New.tbPage.Selected.Index = 4 Then
        If tbsType.SelectedItem.Index - 1 = 0 Then
            gstrSQL = "" & _
                " Select ID, 编码,名称 From 部门表 " & _
                " Where ID in (Select 部门ID From 部门性质说明 Where 工作性质='临床')" & _
                "     And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) " & _
                "     And (站点=[2] or 站点 is null) "
        ElseIf tbsType.SelectedItem.Index - 1 = 1 Then
            gstrSQL = "" & _
                " Select ID, 编码,名称,简码 From 部门表 " & _
                " Where ID in (Select 部门ID From 部门性质说明 Where 工作性质 In ('检查','检验','治疗','手术'))" & _
                "     And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) " & _
                "     And (站点=[2] or 站点 is null) "
        Else
            gstrSQL = "" & _
                " Select ID, 编码,名称,简码 From 部门表 " & _
                " Where ID in (Select 部门ID From 部门性质说明 Where 工作性质='护理')" & _
                "     And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) " & _
                "     And (站点=[2] or 站点 is null) "
        End If
        If strSearch <> "" Then
            gstrSQL = gstrSQL & _
                "     And ( 编码 like [1] or 名称 like [1] or 简码 like [1] )"
        End If
        gstrSQL = gstrSQL & vbCrLf & " Order by 编码"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取部门科室", strKey, gstrNodeNo)
        
        With rsTemp
            If .EOF Then
                MsgBox "没有设置该类部门！（部门管理）", vbInformation, gstrSysName
                Exit Function
            End If
        End With
        Set rsTemp = zlDatabase.ShowSQLMultiSelect(Me, gstrSQL, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, True, strKey, gstrNodeNo)
    Else
        If tbsType.SelectedItem.Index - 1 = 0 Then
            gstrSQL = "" & _
                "Select Distinct A.ID,A.编码,A.名称 " & _
                "From 部门表 A, 部门性质说明 B, 未发药品记录 C, 门诊费用记录 D " & _
                "Where (A.站点 = [6] Or A.站点 Is Null) And B.工作性质 ='临床' And A.ID = B.部门id " & _
                "   And (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) And C.库房id = [2] " & _
                "   And instr([3],','||C.单据||',')>0 And C.填制日期 Between [4] And [5] And C.NO = D.NO And C.库房id = D.执行部门id " & _
                "   And A.ID = D.开单部门id And D.病人科室id = D.开单部门id " & _
                IIf(strSearch = "", "", " and (A.编码 like upper([1]) or A.名称 like [1] or A.简码 like upper([1]))")
               
            If int类型 = 1 Then
                '仅门诊，使用门诊费用记录
                gstrSQL = gstrSQL & " Order By A.编码 "
            ElseIf int类型 = 2 Then
                '仅住院，使用住院费用记录
                gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
                gstrSQL = gstrSQL & " Order By A.编码 "
            Else
                '所有，联合门诊、住院费用记录
                gstrSQL = gstrSQL & " Union " & Replace(gstrSQL, "门诊费用记录", "住院费用记录")
                gstrSQL = gstrSQL & " Order By 编码 "
            End If
        ElseIf tbsType.SelectedItem.Index - 1 = 1 Then
            gstrSQL = "" & _
                "Select Distinct A.ID,A.编码,A.名称 " & _
                "From 部门表 A, 部门性质说明 B, 未发药品记录 C, 门诊费用记录 D " & _
                "Where (A.站点 = [6] Or A.站点 Is Null) And B.工作性质 In ('检查','检验','治疗','手术') " & _
                "   And A.ID = B.部门id And (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) " & _
                "   And C.库房id = [2] And instr([3],','||C.单据||',')>0 And C.填制日期 Between [4] And [5] And C.NO = D.NO " & _
                "   And C.库房id = D.执行部门id And A.ID = D.开单部门id And D.病人科室id <> D.开单部门id " & _
                IIf(strSearch = "", "", " and (A.编码 like upper([1]) or A.名称 like [1] or A.简码 like upper([1]))")
                
            If int类型 = 1 Then
                '仅门诊，使用门诊费用记录
                gstrSQL = gstrSQL & " Order By A.编码 "
            ElseIf int类型 = 2 Then
                '仅住院，使用住院费用记录
                gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
                gstrSQL = gstrSQL & " Order By A.编码 "
            Else
                '所有，联合门诊、住院费用记录
                gstrSQL = gstrSQL & " Union " & Replace(gstrSQL, "门诊费用记录", "住院费用记录")
                gstrSQL = gstrSQL & " Order By 编码 "
            End If
        Else
            '以病区为条件时，仅门诊类型时不提取数据
            If int类型 = 1 Then
                Exit Function
            End If
                
            gstrSQL = "" & _
                "Select Distinct A.ID,A.编码,A.名称 " & _
                "From 部门表 A, 部门性质说明 B, 未发药品记录 C, 住院费用记录 D " & _
                "Where (A.站点 = [6] Or A.站点 Is Null) And B.工作性质 = '护理' And A.ID = B.部门id " & _
                "   And (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01', 'yyyy-MM-dd')) And C.库房id = [2] " & _
                "   And instr([3],','||C.单据||',')>0 And C.填制日期 Between [4] And [5] And C.NO = D.NO " & _
                "   And C.库房id = D.执行部门id And A.ID = D.病人病区id " & _
                IIf(strSearch = "", "", " and (A.编码 like upper([1]) or A.名称 like [1] or A.简码 like upper([1]))")
                
            If zlDatabase.GetPara("病区发料方式", glngSys, mlngModule) = "" Then
                gstrSQL = gstrSQL & " And D.病人科室id = D.开单部门id "
            End If
            
            gstrSQL = gstrSQL & " Order By A.编码 "
        End If
    
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strKey, Val(mArrFilter("发料部门ID")), CStr("," & mArrFilter("单据") & ","), CDate(mArrFilter("日期范围")(0)), CDate(mArrFilter("日期范围")(1)), gstrNodeNo)
        
        With rsTemp
            If .EOF Then
                Exit Function
            End If
           
            Do While Not .EOF
                gstrSQL = "Select Distinct A.药品id " & _
                    " From 药品收发记录 A, 未发药品记录 B, 门诊费用记录 C " & _
                    " Where A.单据 = B.单据 And A.NO = B.NO And a.库房id = b.库房id And A.审核人 Is Null And A.NO = C.NO And B.库房id = C.执行部门id " & _
                    " And B.库房id = [2] And instr([3],','||B.单据||',')>0 And B.填制日期 Between [4] And [5] "
                    
                If tbsType.SelectedItem.Index - 1 = 0 Then
                    gstrSQL = gstrSQL & " And C.开单部门id = [1] And C.病人科室id=C.开单部门id "
                    
                    If int类型 = 2 Then
                        '仅住院时，使用住院费用记录
                        gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
                    ElseIf int类型 = 0 Then
                        '所有时，联合门诊、住院费用记录
                        gstrSQL = gstrSQL & " Union " & Replace(gstrSQL, "门诊费用记录", "住院费用记录")
                    End If
                ElseIf tbsType.SelectedItem.Index - 1 = 1 Then
                    gstrSQL = gstrSQL & " And C.开单部门id = [1] And C.病人科室id<>C.开单部门id "
                    
                    If int类型 = 2 Then
                        '仅住院时，使用住院费用记录
                        gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
                    ElseIf int类型 = 0 Then
                        '所有时，联合门诊、住院费用记录
                        gstrSQL = gstrSQL & " Union " & Replace(gstrSQL, "门诊费用记录", "住院费用记录")
                    End If
                Else
                    '以病区为条件时，仅门诊类型时不提取数据
                    If int类型 = 1 Then
                        Exit Function
                    End If
            
                    If zlDatabase.GetPara("病区发料方式", glngSys, mlngModule) = "" Then
                        gstrSQL = gstrSQL & " And C.病人病区id = [1] And C.病人科室id=C.开单部门id "
                    Else
                        gstrSQL = gstrSQL & " And C.病人病区id = [1] "
                    End If
                    gstrSQL = Replace(gstrSQL, "门诊费用记录", "住院费用记录")
                End If
                
                gstrSQL = "Select Count(Distinct 药品id) As 药品 From (" & gstrSQL & ")"

                Set rsCount = zlDatabase.OpenSQLRecord(gstrSQL, "取部门科室", CLng(!Id), Val(mArrFilter("发料部门ID")), CStr("," & mArrFilter("单据") & ","), CDate(mArrFilter("日期范围")(0)), CDate(mArrFilter("日期范围")(1)))
                
                strName = !名称 & "(" & rsCount!药品 & "种卫材待发）"
                strSelectSql = IIf(strSelectSql = "", "", strSelectSql & " Union All ") & "Select " & !Id & " As ID," & !编码 & " As 编码," & "'" & strName & "'" & " As 名称  From Dual "
                
                .MoveNext
            Loop
        End With
        
        Set rsTemp = zlDatabase.ShowSQLMultiSelect(Me, strSelectSql, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, True)
    End If

    If blnCancel = True Then
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    
    If rsTemp Is Nothing Then
        ShowMsgBox "没有满足条件的科室,请检查!"
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    If objCtl.Enabled Then objCtl.SetFocus
    With rsTemp
        objCtl.Tag = ""
        Do While Not .EOF
            strTemp = strTemp & "," & NVL(rsTemp!名称)
            objCtl.Tag = objCtl.Tag & "," & NVL(rsTemp!Id)
            .MoveNext
        Loop
    End With
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    strKey = objCtl.Tag
    objCtl.Text = strTemp
    objCtl.Tag = strKey
    OS.PressKey vbKeyTab
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub tbsType_Click()
    txt科室.Text = ""
    txt科室.Tag = ""
End Sub

Private Sub tbsType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub txtEDIT_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    Dim intYear As Integer, strYear As String
    Dim strType As String
    Dim intType As Integer
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtEDIT(Index)) = "" Then Exit Sub
    '--如果不满八位,则按规则产生--
    Me.txtEDIT(Index) = UCase(LTrim(Me.txtEDIT(Index)))
    If Len(txtEDIT(Index)) < 8 Then
        strType = Trim(zlDatabase.GetPara("查询业务类型", glngSys, mlngModule, ""))
        
        If strType = "" Or strType = "0,0,0" Or InStr(1, strType, "25") > 0 Or InStr(1, strType, "26") > 0 Then
            intType = 14
        Else
            intType = 13
        End If
        
        txtEDIT(Index).Text = zlCommFun.GetFullNO(txtEDIT(Index).Text, intType, cbo发料部门.ItemData(cbo发料部门.ListIndex))
    End If
    OS.PressKey (vbKeyTab)
End Sub

Private Sub txtPati_Change()
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtPati.Text = "" And Me.ActiveControl Is txtPati)
End Sub

Private Sub txtPati_GotFocus()
    If Not mobjICCard Is Nothing And txtPati.Text = "" Then
        Call mobjICCard.SetEnabled(True)
    End If
End Sub

Private Sub txtPati_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strCard As String
    
    strCard = Split(Split(IDKNType.IDKindStr, ";")(mintType), "|")(1)
    If KeyCode = vbKeyReturn Then
        If strCard = "姓名" Then
            '混合显示,刷就诊卡
            Call cmd刷新_Click
        ElseIf strCard = "IC卡号" Then
            'IC卡
            Call cmd刷新_Click
        ElseIf strCard = "住院号" Then
            OS.PressKey vbKeyTab
        ElseIf strCard = "病人id" Then
            OS.PressKey vbKeyTab
        ElseIf strCard = "床号" Then
            OS.PressKey vbKeyTab
        ElseIf strCard = "门诊号" Then
            OS.PressKey vbKeyTab
        Else
            '银行卡
            Call cmd刷新_Click
        End If
    End If
End Sub
Private Sub txtPati_KeyPress(KeyAscii As Integer)
    Dim strCard As String
    mblnCard = False

    strCard = Split(Split(IDKNType.IDKindStr, ";")(mintType), "|")(1)
    If strCard <> "姓名" And strCard <> "IC卡" And strCard <> "住院号" And strCard <> "病人id" And strCard <> "床号" And strCard <> "门诊号" Then
        '其他消费卡
        If Len(txtPati.Text) = txtPati.MaxLength - 1 And KeyAscii <> 8 Then
            txtPati.Text = txtPati.Text & Chr(KeyAscii)
            txtPati.SelStart = Len(txtPati.Text)
            KeyAscii = 0

            cmd刷新_Click
        End If
    End If
End Sub

Private Sub txtPati_LostFocus()
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (False)
End Sub

Private Sub txt科室_Change()
    txt科室.Tag = ""
End Sub
Public Sub Set发料窗口条件(ByVal blnTran As Boolean, ByVal strNo As String, ByVal strStartDate As String, ByVal strEndDate As String, ByVal lng病人id As Long, ByVal lng库房ID As Long)
    '-----------------------------------------------------------------------------------------------------------
    '功能:设置相关的发药窗口传入的条件
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-01 22:09:07
    '-----------------------------------------------------------------------------------------------------------
    mstrDrugStartDate = strStartDate: mstrDrugEndDate = strEndDate: mblnTrans = blnTran
    mlng库房id = lng库房ID: mlng病人id = lng病人id: mstrNo = strNo
    
    Call InitData
    
    If mlng病人id <> 0 Then
        Me.txtPati.Text = mlng病人id
        mintType = 4
        Me.IDKNType.IDKind = 4
        
    End If
End Sub
Private Sub txt科室_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txt科室.Tag <> "" Then OS.PressKey vbKeyTab: Exit Sub
    If Select部门类型(txt科室, Trim(txt科室.Text)) = False Then
        DoEvents
        txt科室.SetFocus
    Else
        DoEvents
        cmd刷新.SetFocus
    End If
End Sub

Public Property Get GetFilterCon() As Variant
    Call GetFilter
    Set GetFilterCon = mArrFilter
End Property

Public Property Get CheckDept() As Boolean
    CheckDept = cbo发料部门.ListCount <> 0
End Property




