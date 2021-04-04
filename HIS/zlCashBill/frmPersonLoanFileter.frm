VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPersonLoanFileter 
   BorderStyle     =   0  'None
   Caption         =   "过滤条件"
   ClientHeight    =   3870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   4050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmd刷新 
      Caption         =   "过滤(&F)"
      Height          =   390
      Left            =   2700
      TabIndex        =   28
      Top             =   3105
      Width           =   1050
   End
   Begin VB.PictureBox picRequisition 
      BorderStyle     =   0  'None
      Height          =   2940
      Index           =   0
      Left            =   75
      ScaleHeight     =   2940
      ScaleWidth      =   3855
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   90
      Width           =   3855
      Begin VB.CheckBox chkDate 
         Caption         =   "已确认的借款"
         Height          =   375
         Index           =   1
         Left            =   75
         TabIndex        =   4
         Top             =   825
         Width           =   1665
      End
      Begin VB.CheckBox chkDate 
         Caption         =   "待确认的借款"
         Height          =   375
         Index           =   0
         Left            =   75
         TabIndex        =   0
         Top             =   0
         Value           =   1  'Checked
         Width           =   1665
      End
      Begin VB.TextBox txtEdit 
         Height          =   330
         Index           =   0
         Left            =   615
         TabIndex        =   13
         Top             =   2490
         Width           =   3105
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   315
         Index           =   0
         Left            =   615
         TabIndex        =   1
         Top             =   375
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127926275
         CurrentDate     =   37007
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   315
         Index           =   0
         Left            =   2430
         TabIndex        =   3
         Top             =   375
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127926275
         CurrentDate     =   37007
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   315
         Index           =   1
         Left            =   615
         TabIndex        =   5
         Top             =   1185
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127926275
         CurrentDate     =   37007
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   315
         Index           =   1
         Left            =   2430
         TabIndex        =   7
         Top             =   1185
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127926275
         CurrentDate     =   37007
      End
      Begin MSComCtl2.DTPicker dtpStartDate 
         Height          =   315
         Index           =   2
         Left            =   615
         TabIndex        =   9
         Top             =   1935
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127926275
         CurrentDate     =   37007
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   315
         Index           =   2
         Left            =   2430
         TabIndex        =   11
         Top             =   1935
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127926275
         CurrentDate     =   37007
      End
      Begin VB.CheckBox chkDate 
         Caption         =   "已取消确认的借款"
         Height          =   375
         Index           =   2
         Left            =   75
         TabIndex        =   8
         Top             =   1575
         Width           =   2025
      End
      Begin VB.Label lbl至 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   180
         Index           =   0
         Left            =   2070
         TabIndex        =   2
         Top             =   435
         Width           =   180
      End
      Begin VB.Label lbl至 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   180
         Index           =   1
         Left            =   2070
         TabIndex        =   6
         Top             =   1245
         Width           =   180
      End
      Begin VB.Label lbl至 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   180
         Index           =   2
         Left            =   2070
         TabIndex        =   10
         Top             =   1995
         Width           =   180
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "借出人"
         Height          =   180
         Index           =   0
         Left            =   0
         TabIndex        =   12
         Top             =   2565
         Width           =   540
      End
   End
   Begin VB.PictureBox picRequisition 
      BorderStyle     =   0  'None
      Height          =   2865
      Index           =   1
      Left            =   75
      ScaleHeight     =   2865
      ScaleWidth      =   3885
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   180
      Width           =   3885
      Begin VB.TextBox txtEdit 
         Height          =   330
         Index           =   1
         Left            =   615
         TabIndex        =   27
         Top             =   2415
         Width           =   3105
      End
      Begin VB.CheckBox chkOutDate 
         Caption         =   "按申请日期查找"
         Height          =   375
         Index           =   0
         Left            =   75
         TabIndex        =   14
         Top             =   0
         Width           =   1665
      End
      Begin VB.CheckBox chkOutDate 
         Caption         =   "按借出日期查找"
         Height          =   375
         Index           =   1
         Left            =   75
         TabIndex        =   18
         Top             =   825
         Value           =   1  'Checked
         Width           =   1665
      End
      Begin VB.CheckBox chkOutDate 
         Caption         =   "按取消日期查找"
         Height          =   375
         Index           =   2
         Left            =   75
         TabIndex        =   22
         Top             =   1575
         Width           =   1665
      End
      Begin MSComCtl2.DTPicker dtpOutStartDate 
         Height          =   315
         Index           =   0
         Left            =   615
         TabIndex        =   15
         Top             =   375
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127926275
         CurrentDate     =   37007
      End
      Begin MSComCtl2.DTPicker dtpOutEndDate 
         Height          =   315
         Index           =   0
         Left            =   2430
         TabIndex        =   17
         Top             =   375
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127926275
         CurrentDate     =   37007
      End
      Begin MSComCtl2.DTPicker dtpOutStartDate 
         Height          =   315
         Index           =   1
         Left            =   615
         TabIndex        =   19
         Top             =   1185
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127926275
         CurrentDate     =   37007
      End
      Begin MSComCtl2.DTPicker dtpOutEndDate 
         Height          =   315
         Index           =   1
         Left            =   2430
         TabIndex        =   21
         Top             =   1185
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127926275
         CurrentDate     =   37007
      End
      Begin MSComCtl2.DTPicker dtpOutStartDate 
         Height          =   315
         Index           =   2
         Left            =   615
         TabIndex        =   23
         Top             =   1935
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127926275
         CurrentDate     =   37007
      End
      Begin MSComCtl2.DTPicker dtpOutEndDate 
         Height          =   315
         Index           =   2
         Left            =   2430
         TabIndex        =   25
         Top             =   1935
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127926275
         CurrentDate     =   37007
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "借款人"
         Height          =   180
         Index           =   1
         Left            =   0
         TabIndex        =   26
         Top             =   2490
         Width           =   540
      End
      Begin VB.Label lbl至 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   180
         Index           =   5
         Left            =   2070
         TabIndex        =   24
         Top             =   1995
         Width           =   180
      End
      Begin VB.Label lbl至 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   180
         Index           =   4
         Left            =   2070
         TabIndex        =   20
         Top             =   1245
         Width           =   180
      End
      Begin VB.Label lbl至 
         AutoSize        =   -1  'True
         Caption         =   "至"
         Height          =   180
         Index           =   3
         Left            =   2070
         TabIndex        =   16
         Top             =   435
         Width           =   180
      End
   End
End
Attribute VB_Name = "frmPersonLoanFileter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Option Explicit
Private mArrFilter As Variant
Private mblnRequisition As Boolean   'true-我的借款记录,false-我的借出记录
Private mstrPrivs As String
Private mlngModule As Long
Private Enum mTxtIdx
    idx_借款人 = 1
    idx_借出人 = 0
End Enum
Private mblnRequisitionChange As Boolean   '改变了我的借款记录条件
Private mblnOutPayChange As Boolean   '改变了我的借出记录条件


'--------------------------------------------------------------------------------------------------------
Public Event zlRefreshCon(ByVal arrFilter As Variant, ByVal blnRequisition As Boolean)

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
    
    If chkDate(0).Value = 1 Then
        cllFilter.Add Array(Format(dtpStartDate(0).Value, "yyyy-mm-dd") & " 00:00:00", Format(dtpEndDate(0).Value, "yyyy-mm-dd") & " 23:59:59"), "申请时间"
    Else
        cllFilter.Add Array("1901-01-01", "1901-01-01"), "申请时间"
    End If
    
    If chkDate(1).Value = 1 Then
        cllFilter.Add Array(Format(dtpStartDate(1).Value, "yyyy-mm-dd") & " 00:00:00", Format(dtpEndDate(1).Value, "yyyy-mm-dd") & " 23:59:59"), "借出时间"
    Else
        cllFilter.Add Array("1901-01-01", "1901-01-01"), "借出时间"
    End If
    
    If chkDate(2).Value = 1 Then
        cllFilter.Add Array(Format(dtpStartDate(2).Value, "yyyy-mm-dd") & " 00:00:00", Format(dtpEndDate(2).Value, "yyyy-mm-dd") & " 23:59:59"), "取消时间"
    Else
        cllFilter.Add Array("1901-01-01", "1901-01-01"), "取消时间"
    End If
    
    
    If chkOutDate(0).Value = 1 Then
        cllFilter.Add Array(Format(dtpOutStartDate(0).Value, "yyyy-mm-dd") & " 00:00:00", Format(dtpOutEndDate(0).Value, "yyyy-mm-dd") & " 23:59:59"), "借出-申请时间"
    Else
        cllFilter.Add Array("1901-01-01", "1901-01-01"), "借出-申请时间"
    End If
    
    If chkOutDate(1).Value = 1 Then
        cllFilter.Add Array(Format(dtpOutStartDate(1).Value, "yyyy-mm-dd") & " 00:00:00", Format(dtpOutEndDate(1).Value, "yyyy-mm-dd") & " 23:59:59"), "借出-借出时间"
    Else
        cllFilter.Add Array("1901-01-01", "1901-01-01"), "借出-借出时间"
    End If
    
    If chkOutDate(2).Value = 1 Then
        cllFilter.Add Array(Format(dtpOutStartDate(2).Value, "yyyy-mm-dd") & " 00:00:00", Format(dtpOutEndDate(2).Value, "yyyy-mm-dd") & " 23:59:59"), "借出-取消时间"
    Else
        cllFilter.Add Array("1901-01-01", "1901-01-01"), "借出-取消时间"
    End If
    cllFilter.Add Trim(txtEdit(mTxtIdx.idx_借款人)), "借款人"
    cllFilter.Add Trim(txtEdit(mTxtIdx.idx_借出人)), "借出人"
    Set mArrFilter = cllFilter
    
End Function

 
Private Sub cmd刷新_Click()
    Call GetFilter
    RaiseEvent zlRefreshCon(mArrFilter, blnRequisition)
    If blnRequisition Then
        mblnRequisitionChange = False
    Else
        mblnOutPayChange = False
    End If
End Sub

Private Sub InitData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据
    '入参:intType:0-我的借款记录条件;1-我的借出记录条件
    '编制:刘兴洪
    '日期:2009-09-09 14:41:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    dtpEndDate(0).MaxDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    dtpEndDate(1).MaxDate = dtpEndDate(0).MaxDate
    dtpEndDate(2).MaxDate = dtpEndDate(0).MaxDate

    dtpEndDate(0).Value = dtpEndDate(0).MaxDate
    dtpEndDate(1).Value = dtpEndDate(0).MaxDate
    dtpEndDate(2).Value = dtpEndDate(0).MaxDate
    
    dtpStartDate(0).Value = Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-mm-dd")
    dtpStartDate(1).Value = dtpStartDate(0).Value
    dtpStartDate(2).Value = dtpStartDate(0).Value



    dtpOutEndDate(0).MaxDate = dtpEndDate(0).MaxDate
    dtpOutEndDate(1).MaxDate = dtpOutEndDate(0).MaxDate
    dtpOutEndDate(2).MaxDate = dtpOutEndDate(0).MaxDate

    dtpOutEndDate(0).Value = dtpOutEndDate(0).MaxDate
    dtpOutEndDate(1).Value = dtpOutEndDate(0).MaxDate
    dtpOutEndDate(2).Value = dtpOutEndDate(0).MaxDate
    
    dtpOutStartDate(0).Value = dtpStartDate(0).Value
    dtpOutStartDate(1).Value = dtpOutStartDate(0).Value
    dtpOutStartDate(2).Value = dtpOutStartDate(0).Value
End Sub
 
Private Sub chkDate_Click(Index As Integer)
    dtpStartDate(Index).Enabled = chkDate(Index).Value = 1
    dtpEndDate(Index).Enabled = chkDate(Index).Value = 1
    mblnRequisitionChange = True
End Sub

Private Sub chkDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub dtpEndDate_Change(Index As Integer)
     If dtpEndDate(Index).Value > dtpStartDate(Index).MaxDate Then dtpEndDate(Index).Value = dtpStartDate(Index).MaxDate
    
    If dtpEndDate(Index).Value < dtpStartDate(Index).Value Then
        dtpStartDate(Index).Value = dtpEndDate(Index).Value
    End If
End Sub
Private Sub dtpEndDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub dtpOutEndDate_Change(Index As Integer)
    If dtpOutEndDate(Index).Value > dtpOutStartDate(Index).MaxDate Then dtpOutEndDate(Index).Value = dtpOutStartDate(Index).MaxDate
    If dtpOutEndDate(Index).Value < dtpOutStartDate(Index).Value Then
        dtpOutStartDate(Index).Value = dtpOutEndDate(Index).Value
    End If
End Sub

Private Sub dtpOutStartDate_Change(Index As Integer)
    mblnOutPayChange = True
    If dtpOutStartDate(Index).Value > dtpOutEndDate(Index).MaxDate Then dtpOutStartDate(Index).Value = dtpOutEndDate(Index).MaxDate
    If dtpOutEndDate(Index).Value < dtpOutStartDate(Index).Value Then
        dtpOutEndDate(Index).Value = dtpOutStartDate(Index).Value
    End If
End Sub

Private Sub dtpStartDate_Change(Index As Integer)
    mblnRequisitionChange = True
    If dtpStartDate(Index).Value > dtpOutEndDate(Index).MaxDate Then dtpStartDate(Index).Value = dtpEndDate(Index).MaxDate
    If dtpEndDate(Index).Value < dtpStartDate(Index).Value Then
        dtpEndDate(Index).Value = dtpStartDate(Index).Value
    End If
End Sub

Private Sub dtpStartDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    mstrPrivs = gstrPrivs: mlngModule = glngModul
    mblnRequisitionChange = True: mblnOutPayChange = True
End Sub

Private Sub Form_Resize()
        cmd刷新.Left = Me.ScaleLeft + ScaleWidth - cmd刷新.Width - 50
End Sub

Private Function Select人员选择器(ByVal objCtl As Control, ByVal strSearch As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:部门选择器
    '入参::objCtl-指定控件
    '     strSearch-要搜索的条件
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-05-01 14:18:58
    '-----------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, strKey As String, strTittle As String, lngH As Long, strTemp As String
    Dim vRect As RECT
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '功能：
    '参数：
    '     frmParent=显示的父窗体
    '     strSQL=数据来源,不同风格的选择器对SQL中的字段有不同要求
    '     bytStyle=选择器风格
    '       为0时:列表风格:ID,…
    '       为1时:树形风格:ID,上级ID,编码,名称(如果bln末级，则需要末级字段)
    '       为2时:双表风格:ID,上级ID,编码,名称,末级…；ListView只显示末级=1的项目
    '     strTitle=选择器功能命名,也用于个性化区分
    '     bln末级=当树形选择器(bytStyle=1)时,是否只能选择末级为1的项目
    '     strSeek=当bytStyle<>2时有效,缺省定位的项目。
    '             bytStyle=0时,以ID和上级ID之后的第一个字段为准。
    '             bytStyle=1时,可以是编码或名称
    '     strNote=选择器的说明文字
    '     blnShowSub=当选择一个非根结点时,是否显示所有下级子树中的项目(项目多时较慢)
    '     blnShowRoot=当选择根结点时,是否显示所有项目(项目多时较慢)
    '     blnNoneWin,X,Y,txtH=处理成非窗体风格,X,Y,txtH表示调用界面输入框的坐标(相对于屏幕)和高度
    '     Cancel=返回参数,表示是否取消,主要用于blnNoneWin=True时
    '     blnMultiOne=当bytStyle=0时,是否将对多行相同记录当作一行判断
    '     blnSearch=是否显示行号,并可以输入行号定位
    '返回：取消=Nothing,选择=SQL源的单行记录集
    '说明：
    '     1.ID和上级ID可以为字符型数据
    '     2.末级等字段不要带空值
    '应用：可用于各个程序中数据量不是很大的选择器,输入匹配列表等。
    
    strTittle = "人员选择器"
    vRect = zlcontrol.GetControlRect(objCtl.hwnd)
    lngH = objCtl.Height
    
  
    gstrSQL = "" & _
    "   Select Distinct B.ID, B.编号, B.姓名, B.别名, B.简码, B.出生日期, B.性别, B.办公室电话 " & _
    "   From 人员性质说明 A, 人员表 B " & _
    "   Where A.人员id = B.ID And A.人员性质 In ('门诊挂号员', '门诊收费员', '预交收款员', '住院结帐员') " & _
    "         and (b.编号 like upper([1]) or b.姓名 like [1] or b.简码 like upper([1]) or b.别名 like [1]) " & _
    "   Order By b.编号"
    
    strKey = GetMatchingSting(strSearch, False)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, strKey)
 
    If blnCancel = True Then
       zlcontrol.ControlSetFocus objCtl, True
        Exit Function
    End If
    
    If rsTemp Is Nothing Then
        ShowMsgbox "没有满足条件的人员信息,请检查!"
       zlcontrol.ControlSetFocus objCtl, True
        Exit Function
    End If
    zlcontrol.ControlSetFocus objCtl, True
    objCtl.Text = Nvl(rsTemp!姓名)
    objCtl.Tag = Nvl(rsTemp!姓名)
    zlCommFun.PressKey vbKeyTab
    Select人员选择器 = True
End Function
Public Sub Init条件()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关条件
    '编制:刘兴洪
    '日期:2009-09-09 14:59:04
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call InitData
End Sub
Public Property Get GetFilterCon() As Variant
    Call GetFilter
    Set GetFilterCon = mArrFilter
    
End Property

Private Sub txtEdit_Change(Index As Integer)
    txtEdit(Index).Tag = ""
    If Index = mTxtIdx.idx_借款人 Then
        mblnOutPayChange = True
    Else
        mblnRequisitionChange = True
    End If
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txtEdit(Index).Tag <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If txtEdit(Index).Text = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Select人员选择器(txtEdit(Index), Trim(txtEdit(Index).Text)) = False Then Exit Sub
End Sub
Private Sub chkOutDate_Click(Index As Integer)
    dtpOutStartDate(Index).Enabled = chkOutDate(Index).Value = 1
    dtpOutEndDate(Index).Enabled = chkOutDate(Index).Value = 1
    mblnOutPayChange = True
End Sub

Private Sub chkOutDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub dtpOutEndDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Private Sub dtpOutStartDate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
Public Property Get blnRequisition() As Boolean
    blnRequisition = mblnRequisition
End Property
Public Property Let blnRequisition(ByVal vNewValue As Boolean)
    mblnRequisition = vNewValue
    picRequisition(0).Visible = mblnRequisition
    picRequisition(1).Visible = Not mblnRequisition
End Property

Public Property Get IsMyRequistionConChange() As Boolean
   '条件发生了改变
   IsMyRequistionConChange = mblnRequisitionChange
End Property

Public Property Let IsMyRequistionConChange(ByVal vNewValue As Boolean)
    mblnRequisitionChange = vNewValue
End Property

Public Property Get IsMyOutPayConChange() As Boolean
   '条件发生了改变
   IsMyOutPayConChange = mblnOutPayChange
End Property

Public Property Let IsMyOutPayConChange(ByVal vNewValue As Boolean)
    mblnOutPayChange = vNewValue
End Property
Public Sub ReActionFilter(ByVal blnRequisition As Boolean)
    '重新缴活过滤
    mblnRequisition = blnRequisition
    cmd刷新_Click
End Sub
