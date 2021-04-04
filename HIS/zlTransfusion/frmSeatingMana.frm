VERSION 5.00
Begin VB.Form frmSeatingMana 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "座位管理"
   ClientHeight    =   3885
   ClientLeft      =   2355
   ClientTop       =   2385
   ClientWidth     =   5085
   Icon            =   "frmSeatingMana.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleMode       =   0  'User
   ScaleWidth      =   5085
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkAgainAdd 
      Caption         =   "连续增加(&A)"
      Height          =   270
      Left            =   450
      TabIndex        =   20
      Top             =   3375
      Width           =   1665
   End
   Begin VB.Frame fraOne 
      Height          =   3165
      Left            =   150
      TabIndex        =   2
      Top             =   45
      Width           =   4755
      Begin VB.TextBox txt分类 
         Height          =   300
         Left            =   1005
         MaxLength       =   30
         TabIndex        =   8
         Top             =   540
         Width           =   1500
      End
      Begin VB.TextBox txt呼叫器 
         Height          =   300
         Left            =   1005
         MaxLength       =   30
         TabIndex        =   5
         Top             =   2670
         Width           =   3510
      End
      Begin VB.OptionButton Opt类型 
         Caption         =   "床位"
         Height          =   240
         Index           =   1
         Left            =   3810
         TabIndex        =   12
         Top             =   915
         Width           =   675
      End
      Begin VB.OptionButton Opt类型 
         Caption         =   "坐位"
         Height          =   240
         Index           =   0
         Left            =   3105
         TabIndex        =   11
         Top             =   915
         Value           =   -1  'True
         Width           =   675
      End
      Begin VB.CommandButton cmdPopu 
         Caption         =   "…"
         Height          =   240
         Left            =   4200
         TabIndex        =   7
         Top             =   1230
         Width           =   270
      End
      Begin VB.TextBox txt收费项目 
         Height          =   300
         Left            =   1005
         MaxLength       =   103
         TabIndex        =   13
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txt编号 
         Height          =   300
         Left            =   2985
         MaxLength       =   30
         TabIndex        =   9
         Top             =   540
         Width           =   1500
      End
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   1005
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   870
         Width           =   1500
      End
      Begin VB.TextBox txt备注 
         Height          =   705
         Left            =   1005
         MaxLength       =   100
         TabIndex        =   4
         Top             =   1890
         Width           =   3495
      End
      Begin VB.TextBox txtDept 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1005
         TabIndex        =   6
         Top             =   210
         Width           =   3495
      End
      Begin VB.TextBox txt收费标准 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Left            =   1005
         TabIndex        =   3
         Top             =   1545
         Width           =   3495
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "分类"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   23
         Top             =   600
         Width           =   720
      End
      Begin VB.Label lbl编号 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "呼叫器号"
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Top             =   2700
         Width           =   720
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "类型"
         Height          =   180
         Left            =   2445
         TabIndex        =   21
         Top             =   945
         Width           =   495
      End
      Begin VB.Label lbl编号 
         Alignment       =   1  'Right Justify
         Caption         =   "编号"
         Height          =   180
         Index           =   0
         Left            =   2445
         TabIndex        =   19
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lbl等级 
         Alignment       =   1  'Right Justify
         Caption         =   "状态"
         Height          =   180
         Left            =   465
         TabIndex        =   18
         Top             =   915
         Width           =   495
      End
      Begin VB.Label lbl收费项目 
         Alignment       =   1  'Right Justify
         Caption         =   "收费项目"
         Height          =   180
         Left            =   165
         TabIndex        =   17
         Top             =   1212
         Width           =   795
      End
      Begin VB.Label lbl备注 
         Alignment       =   1  'Right Justify
         Caption         =   "备注"
         Height          =   180
         Left            =   165
         TabIndex        =   16
         Top             =   1890
         Width           =   795
      End
      Begin VB.Label lbl科室 
         Alignment       =   1  'Right Justify
         Caption         =   "科室"
         Height          =   180
         Left            =   465
         TabIndex        =   15
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "收费标准"
         Height          =   180
         Left            =   165
         TabIndex        =   14
         Top             =   1536
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdCance 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3525
      TabIndex        =   1
      Top             =   3345
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2265
      TabIndex        =   0
      Top             =   3345
      Width           =   1100
   End
End
Attribute VB_Name = "frmSeatingMana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintType As Integer '调用模式 0-新增 1-修改
Private mSeating As Seating
Private mstr科室名称 As String
Private mblnOk As Boolean '修改，增加是否成功
Private mSelectTxt As String '选择的收费项目文本
Private mblnShow As Boolean '是否已显示，用于连续增加座位时
Private mSeatings As Seatings
Private mfrmMain As frmDockSeat

Public Function SeatingMana(ByVal intType As Integer, ByVal curSeatings As Seatings, ByVal int类别 As Integer, ByVal StrKey As String, ByVal frmParent As Form, Optional strType As String) As Boolean
    'intType: 0-增加 1-修改
    'curSeatings : 座位记录集
    'int类别 : 要增加或修改的座位类型
    'strKey : 如果是修改方式，传入要修改的座位的编号,增加方式可传空串
    '
    mblnOk = False
    Set mSeatings = curSeatings
    mintType = intType
    Set mSeating = New Seating
    
    If intType = 0 Then
        mSeating.编号 = mSeatings.GetNextNo(int类别)
        mSeating.分类 = strType
        If mSeating.分类 = "" Then mSeating.分类 = "普通座位"
        
    Else
        Set mSeating = mSeatings.Item(StrKey)
    End If
    mSeating.类别 = int类别
    mstr科室名称 = mSeatings.科室名称
    
    If (intType = 0 And Not mblnShow) Or intType = 1 Then
        Set mfrmMain = frmParent
        frmSeatingMana.Show vbModal, frmParent
    Else
        Call initForm
    End If

    SeatingMana = True
    
End Function

Private Sub cmdCance_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    
    mSeating.收费项目 = txt收费项目
    mSeating.备注 = txt备注

    mSeating.收费细目ID = txt收费项目.Tag
    mSeating.状态 = cboType.ItemData(cboType.ListIndex)
    mSeating.类型 = IIf(Opt类型(0).Value = True, 0, 1)
    mSeating.呼叫器编号 = "" & txt呼叫器
    
    mblnOk = True
    If mintType = 0 Then
        mSeating.编号 = txt编号
        mSeating.分类 = txt分类
        If mSeating.分类 = "" Then mSeating.分类 = "普通座位"
        If mSeating.分类 <> "普通座位" Then
            mSeating.类别 = 1
        Else
            mSeating.类别 = 0
        End If
        With mSeating
        Call mSeatings.Add(0, 0, 0, "", "", .编号, .类别, .状态, _
                         IIf(IsNull(.现价), 0, .现价), IIf(IsNull(.收费细目ID), 0, .收费细目ID), "", IIf(IsNull(.备注), "", .备注), .类型, .分类, .呼叫器编号, .类别 & "_" & .编号)
        End With
        
        Call SeatingMana(mintType, mSeatings, mSeating.类别, "", Me, mSeating.分类)
        If chkAgainAdd.Value = 0 Then
            Unload Me
        Else
            'Call mfrmMain.RefreshMain
        End If
    Else
        Dim strReturn As String
        
        If mSeating.分类 <> "普通座位" Then
            mSeating.类别 = 1
        Else
            mSeating.类别 = 0
        End If
        
        With mSeating
        strReturn = .Update(mSeatings.科室ID, .收费细目ID, .状态, .收费项目, .现价, .备注, .类型, .呼叫器编号)
        If strReturn <> "" Then
            MsgBox "保存数据时出现错误：" & strReturn, vbInformation, gstrSysName
        End If
        End With
        Unload Me
    End If
End Sub

Private Sub cmdPopu_Click()
    Call ShowSelectWindow(0)
End Sub

Private Sub Form_Load()
    'WAIT:座位管理界面
    Call initForm
    mblnShow = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mfrmMain = Nothing
    mblnShow = False
End Sub


Private Sub txt收费项目_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    
    If (txt收费项目 <> mSelectTxt) Or (txt收费项目.Tag > 0) And (InStr(txt收费项目, "]") - InStr(txt收费项目, "[") < 1) Then
        If Trim(txt收费项目) <> "" Then
            Call ShowSelectWindow(1)
        Else
            txt收费项目.Tag = 0
            zlCommFun.PressKey (vbKeyTab)
        End If
    Else
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub ShowSelectWindow(ByVal intLoadType As Integer)
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim blnCanel As Boolean, vRect As RECT, strTXT As String
    On Error GoTo hErr
    strTXT = UCase(Trim(txt收费项目.Text))
    If intLoadType = 0 Then
        strSQL = "Select A.ID, A.编码, A.名称, A.计算单位, B.现价, A.费用类型, Decode(A.服务对象, 1, '门诊', '门诊和住院') As 服务对象," & vbNewLine & _
                "       A.执行科室" & vbNewLine & _
                "From (Select 现价, 收费细目id,价格等级 From 收费价目 Where 终止日期 Is Null Or 终止日期 = To_Date('3000-01-01', 'YYYY-MM-DD')) B," & vbNewLine & _
                "     收费项目目录 A" & vbNewLine & _
                "Where A.ID = B.收费细目id And Mod(A.服务对象, 2) = 1 And" & vbNewLine & _
                "      (A.站点='" & zl9ComLib.gstrNodeNo & "' Or A.站点 is Null) And " & vbNewLine & _
                "      (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And Nvl(A.是否变价, 0) = 0 And" & vbNewLine & _
                "      A.类别 = 'J'" & GetPriceGradeSQL(gstr药品价格等级, gstr卫材价格等级, gstr普通项目价格等级, "A", "B", "1", "2", "3")
    Else
        If InStr(txt收费项目.Text, "]") - InStr(txt收费项目.Text, "[") > 1 Then
            strTXT = Mid(txt收费项目, InStr(strTXT, "[") + 1, InStr(strTXT, "]") - 2)
        End If
        
        strSQL = "Select A.ID, A.编码, A.名称, A.计算单位, B.现价, A.费用类型, Decode(A.服务对象, 1, '门诊', '门诊和住院') As 服务对象," & vbNewLine & _
                "       A.执行科室" & vbNewLine & _
                "From 收费项目别名 C," & vbNewLine & _
                "     (Select 现价, 收费细目id,价格等级 From 收费价目 Where 终止日期 Is Null Or 终止日期 = To_Date('3000-01-01', 'YYYY-MM-DD')) B," & vbNewLine & _
                "     收费项目目录 A" & vbNewLine & _
                "Where A.ID = C.收费细目id And A.ID = B.收费细目id And Mod(A.服务对象, 2) = 1 And" & vbNewLine & _
                "      (A.站点='" & zl9ComLib.gstrNodeNo & "' Or A.站点 is Null) And " & vbNewLine & _
                "      (C.简码 Like '%" & strTXT & "%' Or A.名称 Like '%" & strTXT & "%' Or A.编码 Like '%" & strTXT & "%') And C.码类 = 1 And" & vbNewLine & _
                "      (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And Nvl(A.是否变价, 0) = 0 And" & vbNewLine & _
                "      A.类别 = 'J'" & GetPriceGradeSQL(gstr药品价格等级, gstr卫材价格等级, gstr普通项目价格等级, "A", "B", "1", "2", "3")

    End If
    
    vRect = ZLControl.GetControlRect(txt收费项目.hwnd)
    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "指定收费项目", False, "", "选择收费项目", False, False, True, _
                                         vRect.Left, vRect.Top, txt收费项目.Height, blnCanel, True, True, gstr药品价格等级, gstr卫材价格等级, gstr普通项目价格等级)
                                         
    If Not blnCanel And Not rsTmp Is Nothing Then
        txt收费项目 = Replace("[" & zlCommFun.NVL(rsTmp.Fields("编码")) & "] " & zlCommFun.NVL(rsTmp.Fields("名称")), "[]", "")
        txt收费项目.Tag = zlCommFun.NVL(rsTmp.Fields("ID"), 0)
        txt收费标准 = Format(zlCommFun.NVL(rsTmp.Fields("现价"), 0), "0.00")
        mSelectTxt = txt收费项目
        zlCommFun.PressKey (vbKeyTab)
    Else
        txt收费项目 = ""
        mSelectTxt = txt收费项目
        txt收费项目.Tag = 0
        txt收费标准 = "0.00"
        txt收费项目.SetFocus
    End If
    Exit Sub
hErr:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub initForm()
    Dim str类别 As String
    txt编号 = mSeating.编号
    txt分类 = mSeating.分类
    txtDept = mstr科室名称
    
    cboType.Clear
    
    cboType.AddItem "0-在编", 0
    cboType.ItemData(0) = 0
    cboType.AddItem "2-修缮", 1
    cboType.ItemData(1) = 2
    
'    Select Case mSeating.类别
'    Case 0
'        str类别 = "普通座位"
'    Case 1
'        str类别 = "加座"
'    Case 2
'        str类别 = "特殊药品座位"
'    Case 3
'        str类别 = "VIP座位"
'    End Select
    
    If mintType = 0 Then
        Me.Caption = "座位管理 - 增加"
        txt备注 = ""
        If chkAgainAdd.Value = 0 Then
            '连续增加,不清除收费项目
            txt收费标准 = "0.00"
            txt收费项目 = ""
            txt收费项目.Tag = 0
            
            Opt类型(0).Value = True: Opt类型(1).Value = False
        End If
        cboType.ListIndex = 0
        txt编号.Enabled = True
        txt分类.Enabled = True
        chkAgainAdd.Enabled = True
        
    Else
        txt编号.Enabled = False
        Me.Caption = "座位管理 - 修改"
        txt备注 = mSeating.备注
        txt收费标准 = Format(mSeating.现价, "0.00")
        txt收费项目.Tag = mSeating.收费细目ID
        txt收费项目 = mSeating.收费项目
        cboType.ListIndex = IIf(mSeating.状态 = 0, 0, 1)
        
        If mSeating.类型 = 0 Then
            Opt类型(0).Value = True: Opt类型(1).Value = False
        Else
            Opt类型(0).Value = False: Opt类型(1).Value = True
        End If
        txt呼叫器 = mSeating.呼叫器编号
        txt分类 = mSeating.分类: txt分类.Enabled = False
        If txt分类 = "" Then txt分类 = "普通座位"
        chkAgainAdd.Enabled = False
    End If
End Sub
