VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frm卫材销帐 
   Caption         =   "卫材退料销帐"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   11760
   Icon            =   "frm卫材销帐.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   11760
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picList 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2520
      Left            =   90
      ScaleHeight     =   2520
      ScaleWidth      =   5070
      TabIndex        =   19
      Top             =   2295
      Width           =   5070
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   1605
         Left            =   15
         TabIndex        =   20
         Top             =   30
         Width           =   4290
         _Version        =   589884
         _ExtentX        =   7567
         _ExtentY        =   2831
         _StockProps     =   64
      End
   End
   Begin VB.CommandButton cmdVerify 
      Caption         =   "退料销帐(&V)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   12135
      TabIndex        =   11
      ToolTipText     =   "热键：F2"
      Top             =   90
      Width           =   1335
   End
   Begin VB.CommandButton CmdHelp 
      Caption         =   "帮助(&H)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   12135
      TabIndex        =   10
      Top             =   825
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&E)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   12135
      TabIndex        =   9
      ToolTipText     =   "热键：F2"
      Top             =   465
      Width           =   1335
   End
   Begin VB.Frame fraCondition 
      Height          =   1125
      Left            =   30
      TabIndex        =   1
      Top             =   75
      Width           =   11985
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
         Height          =   330
         Left            =   9540
         TabIndex        =   17
         ToolTipText     =   "输入住院号、病人ID、床号(指定了病区时)"
         Top             =   165
         Width           =   2355
      End
      Begin VB.ComboBox cbo申请人 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8250
         TabIndex        =   15
         Text            =   "cbo申请人"
         Top             =   645
         Width           =   2310
      End
      Begin VB.ComboBox cbo科室 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4425
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   645
         Width           =   2790
      End
      Begin VB.OptionButton opt科室 
         Caption         =   "医技科室(&W)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2670
         TabIndex        =   7
         Top             =   690
         Width           =   1575
      End
      Begin VB.OptionButton opt科室 
         Caption         =   "病区(&T)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1470
         TabIndex        =   6
         Top             =   690
         Value           =   -1  'True
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker Dtp开始时间 
         Height          =   315
         Left            =   1350
         TabIndex        =   2
         Top             =   195
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
         Format          =   121831427
         CurrentDate     =   36985
      End
      Begin MSComCtl2.DTPicker Dtp结束时间 
         Height          =   315
         Left            =   4455
         TabIndex        =   3
         Top             =   195
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
         Format          =   121831427
         CurrentDate     =   36985
      End
      Begin VB.CheckBox chk申请期间 
         Caption         =   "审请期间(&S)"
         Height          =   195
         Left            =   60
         TabIndex        =   0
         Top             =   240
         Width           =   1290
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "刷新(&R)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10710
         TabIndex        =   8
         ToolTipText     =   "热键：F2"
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label lbl时间 
         AutoSize        =   -1  'True
         Caption         =   "申请期间(&S)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label lblPatiInputType 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "住院号↓"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   8685
         TabIndex        =   18
         Top             =   225
         Width           =   840
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "病人信息"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   7725
         TabIndex        =   16
         Top             =   210
         Width           =   840
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "申请人"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   7530
         TabIndex        =   14
         Top             =   705
         Width           =   630
      End
      Begin VB.Label Lbl科室 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "申请科室"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   465
         TabIndex        =   13
         Top             =   705
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "至"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4200
         TabIndex        =   5
         Top             =   255
         Width           =   210
      End
   End
   Begin VB.Menu mnuPati 
      Caption         =   "病人"
      Visible         =   0   'False
      Begin VB.Menu mnuPatiItem 
         Caption         =   "住院号(&0)"
         Index           =   0
      End
      Begin VB.Menu mnuPatiItem 
         Caption         =   "ID(&1)"
         Index           =   1
      End
      Begin VB.Menu mnuPatiItem 
         Caption         =   "床号(&2)"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frm卫材销帐"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'接口参数
Private mintUnit As String              '0-散装单位,1-包装单位
Private mint金额保留位数 As Integer
Private mlng发料部门ID As Long
Private mArrFilter As Variant   '过滤条件
Private mstrPrivs As String
Private mlngModule As Long
'其它变量
Private mblnDrop As Boolean                     '在KeyDown中判断下拉列表是否弹出
Private Const CB_GETDROPPEDSTATE = &H157
Private Const CB_SHOWDROPDOWN = &H14F
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private mfrm未审 As frm卫材销帐_未审核
Private mfrm已审 As frm卫材销帐_已审核
Private mstr开始申请时间 As String, mstr结束申请时间 As String
Private mstr开始审核时间 As String, mstr结束审核时间 As String

Private Enum mPage
    pag_未审 = 0
    pag_已审 = 1
End Enum

Private mobjPlugIn As Object             '外挂接口对象

Public Property Get In_PlugIn() As Object
    Set In_PlugIn = mobjPlugIn
End Property
Public Property Set In_PlugIn(ByVal objVal As Object)
    Set mobjPlugIn = objVal
End Property
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
       
    '基本查询条件
    Set cllFilter = New Collection
    If mlng发料部门ID < 0 Then
        cllFilter.Add 0, "发料部门ID"
    Else
        cllFilter.Add mlng发料部门ID, "发料部门ID"
    End If
    If cbo科室.ListIndex <= 0 Then
        cllFilter.Add 0, "申请科室ID"
    Else
        cllFilter.Add cbo科室.ItemData(cbo科室.ListIndex), "申请科室ID"
    End If
    
    cllFilter.Add Array("1949-01-01 00:00:00", "1949-01-01 23:59:59"), "日期范围"
    Select Case Val(tbPage.Selected.Tag)
    Case mPage.pag_未审
        If chk申请期间.Value = 1 Then
            cllFilter.Remove "日期范围"
            cllFilter.Add Array(Format(dtp开始时间.Value, "yyyy-mm-dd HH:MM:SS"), Format(dtp结束时间.Value, "yyyy-mm-dd HH:MM:SS")), "日期范围"
            mstr开始申请时间 = Format(dtp开始时间.Value, "yyyy-mm-dd HH:MM:SS")
            mstr结束申请时间 = Format(dtp结束时间.Value, "yyyy-mm-dd HH:MM:SS")
        End If
    Case mPage.pag_已审
        cllFilter.Add Array(Format(dtp开始时间.Value, "yyyy-mm-dd HH:MM:SS"), Format(dtp结束时间.Value, "yyyy-mm-dd HH:MM:SS")), "审核日期"
        mstr开始审核时间 = Format(dtp开始时间.Value, "yyyy-mm-dd HH:MM:SS")
        mstr结束审核时间 = Format(dtp结束时间.Value, "yyyy-mm-dd HH:MM:SS")
    End Select
    
    If cbo申请人.ListIndex = 0 Then
        cllFilter.Add "", "申请人"
    Else
        cllFilter.Add NeedName(cbo申请人.Text), "申请人"
    End If
    cllFilter.Add "", "病人姓名"
    
    If Trim(txtPati.Text) <> "" Then
        If Val(lblPatiInputType.Tag) = 0 Then
            cllFilter.Add Val(txtPati.Tag), "住院号"
        Else
            cllFilter.Add 0, "住院号"
        End If
        
        If Val(lblPatiInputType.Tag) = 1 Then
            cllFilter.Add Val(txtPati.Tag), "病人ID"
        Else
            cllFilter.Add 0, "病人ID"
        End If
        If Val(lblPatiInputType.Tag) = 2 Then
            cllFilter.Add Trim(txtPati.Tag), "床号"
        Else
            cllFilter.Add "", "床号"
        End If
    Else
        cllFilter.Add 0, "住院号"
        cllFilter.Add 0, "病人ID"
        cllFilter.Add "", "床号"
    End If
 
   ' Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取批次明细", _
          Val(mArrFilter("发料部门id")), Val(mArrFilter("申请科室id")), _
          CDate(mArrFilter("时间范围")(0)), CDate(mArrFilter("时间范围")(1)), _
          CDate(mArrFilter("发料时间")(0)), CDate(mArrFilter("发料时间")(1)), _
          Trim(mArrFilter("申请人")), Trim(mArrFilter("病人姓名")), _
          Val(mArrFilter("住院号")), Val(mArrFilter("病人ID")))
        
    
    Set mArrFilter = cllFilter
    
End Function

Private Sub 获取申请人()
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取相关部门的申请人
    '入参:int部门类型
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-03 23:25:54
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    
    Dim strWhere As String
    
    On Error GoTo ErrHandle
    If cbo科室.ListIndex > 0 Then strWhere = " And B.部门id = [1] "
 
    gstrSQL = "" & _
        "   Select Distinct A.ID, A.简码||'-'||A.姓名 As 姓名 " & _
        "   From 人员表 A, 部门人员 B " & _
        "   Where A.ID = B.人员id And (a.站点=[2] or a.站点 is null) " & strWhere & _
        "           And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) " & _
        "   Order By 姓名"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取部门人员", Val(cbo科室.ItemData(cbo科室.ListIndex)), gstrNodeNo)
        
    cbo申请人.Clear
    cbo申请人.AddItem "所有申请人"
    cbo申请人.ItemData(cbo申请人.NewIndex) = 0
    Do While Not rsTemp.EOF
        cbo申请人.AddItem rsTemp!姓名
        cbo申请人.ItemData(cbo申请人.NewIndex) = rsTemp!Id
        rsTemp.MoveNext
    Loop
    cbo申请人.ListIndex = 0
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub 获取发料部门名称()
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取发料部门名称
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-05-03 23:29:52
    '-----------------------------------------------------------------------------------------------------------

    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select 名称 From 部门表 Where ID = [1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取库房名称", mlng发料部门ID)
    
    If Not rsTemp.EOF Then
        Me.Caption = Me.Caption & "(当前库房：" & rsTemp!名称 & ")"
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function ShowList(frmMain As Form, strPirvs As String, lngModule As Long, ByVal lng发料部门ID As Long, ByVal intUnit As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:外调程序入库
    '入参:lng发料部门ID-发料部门ID
    '     intUnit-显示单位(0-散装单位,1-包装单位)
    '     int金额保留位数-金额保留位数
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-03 23:46:24
    '-----------------------------------------------------------------------------------------------------------
    mstrPrivs = strPirvs: mlngModule = lngModule
    mlng发料部门ID = lng发料部门ID
    mintUnit = intUnit
    'mint金额保留位数 = int金额保留位数
    Me.Show vbModal, frmMain
    ShowList = True
End Function

Private Sub cbo科室_Click()
    If cbo科室.ListIndex = -1 Then Exit Sub
    If Val(cbo科室.Tag) <> cbo科室.ItemData(cbo科室.ListIndex) Then
        cbo科室.Tag = cbo科室.ItemData(cbo科室.ListIndex)
        Call 获取申请人
    End If
End Sub

Private Sub cbo申请人_Click()
    Exit Sub
End Sub

Private Sub cbo申请人_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnDrop = False
    If KeyCode = 13 Then mblnDrop = SendMessage(cbo申请人.hwnd, CB_GETDROPPEDSTATE, 0, 0) = 1
End Sub

Private Sub cbo申请人_KeyPress(KeyAscii As Integer)
    Dim i As Long, intIdx As Integer
    Dim strText As String, strResult As String, strFilter As String

    If KeyAscii = 13 Then
        strText = UCase(cbo申请人.Text)
        If cbo申请人.ListIndex <> -1 Then
            '弹出列表时,又在文本框输入了内容
            If strText <> cbo申请人.List(cbo申请人.ListIndex) Then Call zlControl.CboSetIndex(cbo申请人.hwnd, -1)
        End If
        If strText = "" Then
            cbo申请人.ListIndex = -1
        ElseIf cbo申请人.ListIndex = -1 Then
            intIdx = -1

            For i = 1 To cbo申请人.ListCount - 1
                If Mid(cbo申请人.List(i), 1, InStr(1, cbo申请人.List(i), "-") - 1) = strText _
                    Or Mid(cbo申请人.List(i), InStr(1, cbo申请人.List(i), "-")) = strText Then
                    intIdx = i
                    Exit For
                End If
            Next

            If intIdx = -1 Then
                For i = 1 To cbo申请人.ListCount - 1
                    If UCase(cbo申请人.List(i)) Like strText & "*" Then
                        intIdx = i
                    End If
                Next
            End If

            cbo申请人.ListIndex = intIdx
            SendMessage cbo申请人.hwnd, CB_SHOWDROPDOWN, True, 0
        ElseIf Not mblnDrop Then
            '回车光标经过
            Call cbo申请人_Click
            Exit Sub
        End If
        If cbo申请人.ListIndex = -1 Then
            cbo申请人.ListIndex = 0
        Else
            If intIdx <> -1 And mblnDrop Then
                '弹出回车-强行激活Click
                Call cbo申请人_Click
            ElseIf intIdx <> cbo申请人.ListIndex And intIdx <> -1 Then
                '弹出让选择-自动激活Click
                cbo申请人.SetFocus
                Exit Sub
            ElseIf intIdx <> -1 Then
                '一次性输中-强行激活Click
                Call cbo申请人_Click
            End If
        End If
    End If
End Sub
Private Function NeedName(strList As String) As String
    NeedName = Mid(strList, InStr(strList, "-") + 1)
End Function

Private Sub chk申请期间_Click()
    dtp开始时间.Enabled = chk申请期间.Value = 1
    dtp结束时间.Enabled = chk申请期间.Value = 1
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name)
End Sub
Private Sub IniDate()
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化相关日期的默认值
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-05-03 23:51:57
    '-----------------------------------------------------------------------------------------------------------
    Dim dtCurrDate As Date
    dtCurrDate = sys.Currentdate
    dtp开始时间.MaxDate = CDate(Format(dtCurrDate, "yyyy-MM-dd 23:59:59"))
    dtp结束时间.MaxDate = dtp开始时间.MaxDate
    dtp开始时间.Value = CDate(Format(DateAdd("D", -1, dtCurrDate), "yyyy-MM-dd 00:00:00"))
    dtp结束时间.Value = CDate(Format(dtCurrDate, "yyyy-MM-dd 23:59:59"))
    mstr开始申请时间 = Format(dtp开始时间.Value, "yyyy-mm-dd HH:MM:SS")
    mstr结束申请时间 = Format(dtp结束时间.Value, "yyyy-mm-dd HH:MM:SS")
    mstr开始审核时间 = mstr开始申请时间
    mstr结束审核时间 = mstr结束申请时间
    
End Sub
Private Sub 获取部门数据(ByVal int部门类型 As Integer)
    '-----------------------------------------------------------------------------------------------------------
    '功能:获取部门数据
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-05-03 23:52:53
    '-----------------------------------------------------------------------------------------------------------

    'int部门类型：0-病区；1-医技科室
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    Select Case int部门类型
        Case 0
            gstrSQL = " Select 编码||'-'||名称 科室,ID From 部门表 " & _
             " Where ID in (Select 部门ID From 部门性质说明 Where 工作性质='护理' And 服务对象 IN(2,3))" & _
             "     And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) And (站点=[1] or 站点 is null) " & _
             " Order By 编码||'-'||名称 "
        Case 1
            gstrSQL = " Select 编码||'-'||名称 科室,ID From 部门表 " & _
             " Where ID in (Select 部门ID From 部门性质说明 Where 工作性质 In ('检查','检验','治疗','手术') And 服务对象 IN(2,3))" & _
             "     And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','yyyy-MM-dd')) And (站点=[1] or 站点 is null) " & _
             " Order By 编码||'-'||名称 "
    End Select
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取科室", gstrNodeNo)
    
    cbo科室.Clear
    
    If int部门类型 = 0 Then
        cbo科室.AddItem "所有病区"
        cbo科室.ItemData(cbo科室.NewIndex) = 0
    Else
        cbo科室.AddItem "所有科室"
        cbo科室.ItemData(cbo科室.NewIndex) = 0
    End If
    
    Do While Not rsTemp.EOF
        cbo科室.AddItem rsTemp!科室
        cbo科室.ItemData(cbo科室.NewIndex) = rsTemp!Id
        rsTemp.MoveNext
    Loop
    
    cbo科室.ListIndex = 0
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub IniDept()
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化部门特性
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2008-05-03 23:53:39
    '-----------------------------------------------------------------------------------------------------------
    If Lbl科室.Tag = "" Then
        Lbl科室.Tag = "-1"
        opt科室_Click (0)
    End If
End Sub
Private Function FullData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:填充数据
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-03 23:57:13
    '-----------------------------------------------------------------------------------------------------------
    Select Case Val(tbPage.Selected.Tag)
    Case mPage.pag_未审
        FullData = mfrm未审.zlRefreshData(Me, mstrPrivs, mlngModule, mintUnit, mArrFilter)
    Case mPage.pag_已审
        FullData = mfrm已审.zlRefreshData(Me, mstrPrivs, mlngModule, mintUnit, mArrFilter)
    End Select
End Function
Private Sub cmdRefresh_Click()
    Call GetFilter
    Call FullData
End Sub

Private Sub cmdVerify_Click()
    Set mfrm未审.In_PlugIn = mobjPlugIn
    If mfrm未审.zlVerifyData = False Then Exit Sub
    Call cmdRefresh_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Load()
    cbo科室.Tag = "-1"
    Call IniDate
    Call IniDept
    Call 获取发料部门名称
    Call InitPage
'    Call GetFilter
End Sub

Private Sub Form_Resize()
    Dim lngTmp As Long
    
    If WindowState = 1 Then Exit Sub
    On Error Resume Next
    If Me.Width < 13860 Then Me.Width = 13860
    If Me.Height < 8790 Then Me.Height = 8790
    
    cmdVerify.Left = Me.ScaleWidth - cmdVerify.Width - 50
    cmdExit.Left = cmdVerify.Left
    cmdHelp.Left = cmdVerify.Left
    fraCondition.Left = Me.ScaleLeft + 20
    fraCondition.Width = Me.ScaleWidth - cmdVerify.Width - 100
    With picList
        .Top = fraCondition.Height + fraCondition.Top + 50
        .Height = Me.ScaleHeight - .Top - 50
        .Width = Me.ScaleWidth - .Left - 50
    End With
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mfrm未审 Is Nothing Then
        Unload mfrm未审
        Set mfrm未审 = Nothing
    End If
    
    If Not mfrm已审 Is Nothing Then
        Unload mfrm已审
        Set mfrm已审 = Nothing
    End If
End Sub

Private Sub lblPatiInputType_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 1 Then
        PopupMenu mnuPati, 2, lblPatiInputType.Left + lblPatiInputType.Width - 30, lblPatiInputType.Top
    End If
End Sub

Private Sub mnuPatiItem_Click(Index As Integer)
    Select Case Index
        Case 0
            lblPatiInputType.Caption = "住院号↓"
            lblPatiInputType.Tag = 0
            txtPati.Text = ""
            txtPati.Tag = ""
        Case 1
            lblPatiInputType.Caption = "ID↓"
            lblPatiInputType.Tag = 1
            txtPati.Text = ""
            txtPati.Tag = ""
        Case 2
            lblPatiInputType.Caption = "床号↓"
            lblPatiInputType.Tag = 2
            txtPati.Text = ""
            txtPati.Tag = ""
    End Select
End Sub
Private Sub opt科室_Click(Index As Integer)
    If Val(Lbl科室.Tag) <> Index Then
        If Index = 1 Then
            mnuPatiItem(2).Enabled = False
            If Val(lblPatiInputType.Tag) = 2 Then
                Call mnuPatiItem_Click(0)
            End If
        Else
            mnuPatiItem(2).Enabled = True
        End If
        
        Call 获取部门数据(Index)
        Lbl科室.Tag = Index
    End If
End Sub

Private Sub picList_Resize()
    With tbPage
        .Top = picList.ScaleTop
        .Height = picList.ScaleHeight
        .Width = picList.ScaleWidth
        .Left = picList.ScaleLeft
    End With

End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Call SetInitDate(Val(Item.Tag))
End Sub
Private Sub SetInitDate(ByVal intType As Integer)
    '功能:设置时间显示
    Select Case intType
    Case mPage.pag_已审
        cmdVerify.Enabled = False
         lbl时间.Caption = "审核期间(&S)"
        chk申请期间.Visible = False
        lbl时间.Visible = True
        dtp开始时间.Value = CDate(mstr开始审核时间)
        dtp结束时间.Value = CDate(mstr结束审核时间)
        dtp开始时间.Enabled = True
        dtp结束时间.Enabled = True
    Case Else
        lbl时间.Caption = "申请期间(&S)"
        cmdVerify.Enabled = True
        chk申请期间.Visible = True
        dtp开始时间.Enabled = chk申请期间.Value = 1: dtp结束时间.Enabled = chk申请期间.Value = 1
        lbl时间.Visible = False
        dtp开始时间.Value = CDate(mstr开始申请时间)
        dtp结束时间.Value = CDate(mstr结束申请时间)
        dtp开始时间.Enabled = chk申请期间.Value = 1
        dtp结束时间.Enabled = chk申请期间.Value = 1
    End Select
End Sub

Private Sub txtPati_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim str标识号 As String
    
    On Error GoTo ErrHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    txtPati.Text = Trim(txtPati.Text)
    
    If txtPati.Text = "" Then Exit Sub
    
    If Val(lblPatiInputType.Tag) = 0 Then
        If InStr(1, txtPati.Text, "-") > 0 Then
            str标识号 = Mid(txtPati.Text, 1, InStr(1, txtPati.Text, "-") - 1)
        Else
            str标识号 = txtPati.Text
        End If
        gstrSQL = "Select Distinct 姓名,住院号 As 标识 From 病人信息 Where 住院号 = [1] "
    ElseIf Val(lblPatiInputType.Tag) = 1 Then
        gstrSQL = "Select Distinct 姓名,病人ID As 标识 From 病人信息 Where 病人ID = [2] "
    Else
        If cbo科室.ListIndex = 0 Then
            MsgBox "请选择病区！"
            Exit Sub
        End If
        str标识号 = txtPati.Text
        gstrSQL = "Select A.姓名,B.床号 As 标识 From 病人信息 A, 床位状况记录 B Where A.病人id = B.病人id And 病区id = [3] And B.床号 = [1] "
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取病人姓名", str标识号, Val(txtPati.Text), Val(cbo科室.ItemData(cbo科室.ListIndex)))
    If rsTemp.RecordCount > 0 Then
        txtPati.Text = rsTemp!标识 & "-" & rsTemp!姓名
        txtPati.Tag = rsTemp!标识
        
        cmdRefresh_Click
    Else
        '通过查找一个不存在的数据来清空表格数据
        txtPati.Text = "-1"
        txtPati.Tag = "-1"
        cmdRefresh_Click
        
        txtPati.Text = ""
        txtPati.Tag = ""
        txtPati.SetFocus
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub txtPati_KeyPress(KeyAscii As Integer)
    If Val(lblPatiInputType.Tag) = 0 Or Val(lblPatiInputType.Tag) = 1 Then
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyEscape Then Exit Sub
        KeyAscii = 0
    End If
End Sub
   
Private Sub InitPage()
    '------------------------------------------------------------------------------
    '功能:初始化页面控件
    '返回:
    '编制:刘兴宏
    '日期:2007/08/18
    '------------------------------------------------------------------------------
    Dim i As Long
    Dim objItem As TabControlItem
 
    
    Set mfrm未审 = New frm卫材销帐_未审核
    Set objItem = tbPage.InsertItem(mPage.pag_未审, "未审核", mfrm未审.hwnd, 0)
    objItem.Tag = mPage.pag_未审
    Set mfrm已审 = New frm卫材销帐_已审核
    Set objItem = tbPage.InsertItem(mPage.pag_已审, "已审核", mfrm已审.hwnd, 0)
    objItem.Tag = mPage.pag_已审
    With tbPage
        .Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
    End With
    Call SetInitDate(mPage.pag_未审)
    Call GetFilter
    Call FullData
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

