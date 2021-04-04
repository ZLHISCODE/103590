VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmExpensesInfo 
   Caption         =   "诊疗信息"
   ClientHeight    =   9900
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExpensesInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9900
   ScaleWidth      =   13695
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraPati 
      Caption         =   "病人信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   90
      TabIndex        =   3
      Top             =   0
      Width           =   11295
      Begin VB.ComboBox cboShow 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         ItemData        =   "frmExpensesInfo.frx":6852
         Left            =   840
         List            =   "frmExpensesInfo.frx":6854
         Style           =   2  'Dropdown List
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   690
         Width           =   4815
      End
      Begin VB.Label lblInformation 
         AutoSize        =   -1  'True
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   7
         Left            =   7200
         TabIndex        =   18
         Top             =   720
         Width           =   90
      End
      Begin VB.Label lblInformation 
         AutoSize        =   -1  'True
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   4
         Left            =   9960
         TabIndex        =   17
         Top             =   360
         Width           =   90
      End
      Begin VB.Label lblInformation 
         AutoSize        =   -1  'True
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   5
         Left            =   7260
         TabIndex        =   16
         Top             =   360
         Width           =   90
      End
      Begin VB.Label lblInformation 
         AutoSize        =   -1  'True
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   3
         Left            =   4830
         TabIndex        =   15
         Top             =   360
         Width           =   90
      End
      Begin VB.Label lblCaption 
         Caption         =   "申请人："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   6660
         TabIndex        =   14
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblCaption 
         Caption         =   "申请科室："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   9000
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblCaption 
         Caption         =   "年龄："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   4230
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblCaption 
         Caption         =   "姓名："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblCaption 
         Caption         =   "性别："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2340
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblCaption 
         Caption         =   "历次："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         Caption         =   "标本类型："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   7
         Left            =   6300
         TabIndex        =   8
         Top             =   720
         Width           =   900
      End
      Begin VB.Label lblInformation 
         AutoSize        =   -1  'True
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   6
         Left            =   6840
         TabIndex        =   7
         Top             =   720
         Width           =   90
      End
      Begin VB.Label lblInformation 
         AutoSize        =   -1  'True
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   2
         Left            =   2940
         TabIndex        =   6
         Top             =   360
         Width           =   90
      End
      Begin VB.Label lblInformation 
         AutoSize        =   -1  'True
         Caption         =   "#"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   180
         Index           =   1
         Left            =   840
         TabIndex        =   5
         Top             =   360
         Width           =   90
      End
   End
   Begin VB.PictureBox picRefresh 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   10896
      Picture         =   "frmExpensesInfo.frx":6856
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "刷新(F5)"
      Top             =   312
      Width           =   480
   End
   Begin VB.PictureBox PicWindows 
      BorderStyle     =   0  'None
      Height          =   276
      Left            =   11496
      ScaleHeight     =   270
      ScaleWidth      =   510
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   516
   End
   Begin XtremeSuiteControls.TabControl TabCtlWindow 
      Height          =   5895
      Left            =   180
      TabIndex        =   0
      Top             =   1380
      Width           =   10545
      _Version        =   589884
      _ExtentX        =   18606
      _ExtentY        =   10393
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmExpensesInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const p医嘱附费管理 As Integer = 1257                       '病人费用模块授权
Private Const p门诊医嘱下达 As Integer = 1252                       '门诊医嘱下达
Private Const p住院医嘱下达 As Integer = 1253                       '住院医嘱下达
Private Const p门诊病历管理 As Integer = 1250                       '门诊病历
Private Const p住院病历管理 As Integer = 1251
Private Const p新版病历管理 As Integer = 2250                       '新版病历
Private Const p新版病历管理_门诊 As Integer = 2251                  '新版病历（门诊）
Private Const p新版病历管理_住院 As Integer = 2252                  '新版病历（住院）

Private mcolSubForm As Collection                                   '卸载子窗体

Private mclsExpenses As Object                                          '费用对象
Private mclsOutAdvices As Object                                        '门诊医嘱对象
Private mclsInAdvices As Object                                         '住院医嘱对象
Private mclsOutEPRs As Object                                           '门诊病历
Private mclsInEPRs As Object                                            '住院病历
Private mobjKernel As Object                                            '医嘱部件
Private mclsEMR As Object                                               '新版电子病历
Private mobjRichEPR As Object                                           '病历核心部件

Private mlngSapmeID As Long                                             '标本ID
Private mrsInfo As New ADODB.Recordset                                  '查出来的基本信息
Private mblnLoadfrm As Boolean                                          '是否加载完成



Private Sub cboShow_Click()
        Call RefreshTab(TabCtlWindow.Selected.Index)
End Sub

Private Sub Form_Activate()
    gobjHisComLib.InitCommon gcnHisOracle
    gobjHisComLib.RegCheck
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 116 Then       'F5
        picRefresh_Click
    End If
End Sub

Private Sub Form_Load()
    Dim lngSysNo As Long
    Dim intIndex As Integer
    Dim strPrivs As String
    Dim strSQL As String, strTmp As String
    Dim rsTmp As ADODB.Recordset

    On Error GoTo Form_Load_Error

    mblnLoadfrm = False

    lngSysNo = 100

    '初始化核心部件
    Set mobjKernel = CreateObject("zlCISKernel.clsCISKernel")
    Set mobjRichEPR = CreateObject("zlRichEPR.cRichEPR")

    Call mobjKernel.InitCISKernel(gcnHisOracle, Me, lngSysNo, "")
    Call mobjRichEPR.InitRichEPR(gcnHisOracle, Me, lngSysNo, False)



    With Me.TabCtlWindow
        Set .Icons = frmPubIcons.imgPublic.Icons
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.Color = xtpTabColorOffice2003
        .PaintManager.OneNoteColors = False
        .PaintManager.BoldSelected = True

        strPrivs = GetPrivFunc(Sel_His_DB, lngSysNo, p医嘱附费管理)  '没有医嘱附费管理权限时不显示
        .InsertItem(0, "费用查询", PicWindows.hWnd, 1).Tag = IIf(strPrivs <> "", "费用查询", "")
        '        .Item(0).Visible = IIf(strPrivs <> "", True, False)
        .Item(0).Visible = False

        strPrivs = GetPrivFunc(Sel_His_DB, lngSysNo, p门诊医嘱下达)
        .InsertItem(1, "门诊医嘱", PicWindows.hWnd, 1).Tag = IIf(strPrivs <> "", "门诊医嘱", "")
        .Item(1).Visible = IIf(strPrivs <> "", True, False)

        strPrivs = GetPrivFunc(Sel_His_DB, lngSysNo, p住院医嘱下达)
        .InsertItem(2, "住院医嘱", PicWindows.hWnd, 1).Tag = IIf(strPrivs <> "", "住院医嘱", "")
        .Item(2).Visible = IIf(strPrivs <> "", True, False)

        strPrivs = GetPrivFunc(Sel_His_DB, lngSysNo, p门诊病历管理)
        .InsertItem(3, "门诊病历", PicWindows.hWnd, 1).Tag = IIf(strPrivs <> "", "门诊病历", "")
        .Item(3).Visible = IIf(strPrivs <> "", True, False)

        strPrivs = GetPrivFunc(Sel_His_DB, lngSysNo, p住院病历管理)
        .InsertItem(4, "住院病历", PicWindows.hWnd, 1).Tag = IIf(strPrivs <> "", "住院病历", "")
        .Item(4).Visible = IIf(strPrivs <> "", True, False)

        If mrsInfo("病人来源") & "" = 2 Then
            strPrivs = GetPrivFunc(Sel_His_DB, lngSysNo, p新版病历管理_住院)
        Else
            strPrivs = GetPrivFunc(Sel_His_DB, lngSysNo, p新版病历管理_门诊)
        End If

        '处理新版电子病历部件
        On Error Resume Next
        If Not gobjEmr.IsInited Or gobjEmr.IsOffline Then
            '没链接服务器不加载
            .InsertItem(5, "电子病历", PicWindows.hWnd, 1).Tag = IIf(strPrivs <> "", "电子病历", "")
            .Item(5).Visible = False
        Else
            Set mclsEMR = CreateObject("zlRichEMR.clsDockEMR")

            Err.Clear: On Error GoTo Form_Load_Error
            If mclsEMR Is Nothing Then
                strPrivs = ""
            Else

            End If
            If mcolSubForm Is Nothing Then
                Set mcolSubForm = New Collection
                If Not mclsEMR Is Nothing Then
                    mcolSubForm.Add mclsEMR.zlGetForm, "_电子病历"
                End If
            End If
            .InsertItem(5, "电子病历", PicWindows.hWnd, 1).Tag = IIf(strPrivs <> "", "电子病历", "")
            .Item(5).Visible = IIf(strPrivs <> "", True, False)
        End If
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.ShowIcons = True

        lblInformation(1).Caption = mrsInfo("姓名")
        lblInformation(2).Caption = mrsInfo("性别")
        lblInformation(3).Caption = mrsInfo("年龄")
        lblInformation(4).Caption = mrsInfo("申请科室")
        lblInformation(5).Caption = mrsInfo("申请人")
        lblInformation(7).Caption = mrsInfo("标本类型")
        If mrsInfo("病人来源") & "" = 2 Then
            strSQL = "Select rownum as 序号,病人id,主页ID,NVL(病人性质,0) 病人性质,当前病区id,住院号,To_Char(入院日期,'YYYY-MM-DD HH24:MI') as 入院日期 From 病案主页 Where 主页ID<>0 And 病人ID=[1] Order by 主页ID Desc"
            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "中联信息", Val(mrsInfo("病人ID") & ""))
        Else
            strSQL = "Select A.ID,A.NO,A.发生时间 as 时间,B.名称 as 科室,a.执行人,a.接收时间 From 病人挂号记录 A,部门表 B" & _
                   " Where A.执行部门ID=B.ID And A.病人ID=[1] And A.发生时间<=[2] And A.记录性质=1 And A.记录状态=1 Order by A.发生时间 Desc,a.接收时间 Desc"
            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "中联信息", Val(mrsInfo("病人ID") & ""), Now)


        End If
        If rsTmp.RecordCount = 0 Then
            Exit Sub
        End If
        cboShow.Clear
        Do While Not rsTmp.EOF
            If mrsInfo("病人来源") & "" = 2 Then
                strTmp = "第 " & rsTmp!主页ID & " 次"    '&  Decode(rsTmp!病人性质, 1, "(门诊留观)", 2, "(住院留观)", "")
                cboShow.AddItem strTmp
                cboShow.ItemData(cboShow.NewIndex) = rsTmp!主页ID

            Else
                strTmp = Format(rsTmp!时间, "YYMMdd") & "/" & rsTmp!科室 & "/" & rsTmp!执行人
                cboShow.AddItem strTmp
                cboShow.ItemData(cboShow.NewIndex) = rsTmp!ID
            End If
            rsTmp.MoveNext
        Loop

        cboShow.ListIndex = 0

        ' Call cboShow_Click


        '只显示门诊或住院

        With Me.TabCtlWindow
            If mrsInfo("病人来源") & "" = 2 Then
                TabCtlWindow.Item(1).Visible = False
                .Item(2).Visible = True
                .Item(3).Visible = False
                .Item(4).Visible = True
            Else
                .Item(1).Visible = True
                .Item(2).Visible = False
                .Item(3).Visible = True
                .Item(4).Visible = False
            End If
        End With

        mblnLoadfrm = True

        '默认加载第一个没有隐藏的页面
        For intIndex = 0 To 5
            If .Item(intIndex).Visible = True Then
                .Item(intIndex).Selected = True
                Call RefreshTab(intIndex)
                Exit For
            End If
        Next

    End With



    Exit Sub
Form_Load_Error:
    Call WriteErrLog("zl9LisInsideComm", "frmExpensesInfo", "执行(Form_Load)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
    Err.Clear

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With fraPati
        .Top = 50
        .Left = 50
        .Width = Me.ScaleWidth - 100
    End With
    
    
    With Me.TabCtlWindow
        .Top = fraPati.Top + fraPati.Height
        .Left = 50
        .Width = Me.ScaleWidth - 100
        .Height = Me.ScaleHeight - fraPati.Height - 100
    End With
    
'    With Me.picRefresh
'        .Top = 100
'        .Left = Me.ScaleWidth - .Width - 100
'
'
'    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call mobjKernel.InitCISKernel(gcnLisOracle, Me, 100, "")
    If Not gobjEmr Is Nothing Then
        Call gobjEmr.CloseForms
    End If
    
    Set mcolSubForm = Nothing
    Set mclsExpenses = Nothing
    Set mclsInAdvices = Nothing
    Set mclsOutAdvices = Nothing
    Set mclsOutEPRs = Nothing
    Set mclsInEPRs = Nothing
    Set mclsEMR = Nothing
    Set mobjKernel = Nothing
    Set mobjRichEPR = Nothing
    Set mrsInfo = Nothing
    mblnLoadfrm = False
    TabCtlWindow.RemoveAll
End Sub

Private Sub picRefresh_Click()
    Call RefreshTab(Me.TabCtlWindow.Selected.Index)
End Sub

Private Sub picRefresh_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.picRefresh.BorderStyle = 1
End Sub

Private Sub picRefresh_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.picRefresh.BorderStyle = 0
End Sub

Private Sub RefreshTab(intIndex As Integer)
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset
          Dim lngMainID As Long, lngDeptID As Long
          Dim strData As String
          Dim lngSysNo As Long

1         On Error GoTo RefreshTab_Error

2         If mblnLoadfrm = False Then Exit Sub

3         mblnLoadfrm = False

4         strData = cboShow.ItemData(cboShow.ListIndex)

          '没有记录时退出
5         If mrsInfo.RecordCount <= 0 Then Exit Sub

6         lngSysNo = 100

          '只显示门诊或住院
7         With Me.TabCtlWindow
8             If mrsInfo("病人来源") & "" = 2 Then
9                 TabCtlWindow.Item(1).Visible = False
10                .Item(2).Visible = True
11                .Item(3).Visible = False
12                .Item(4).Visible = True
13            Else
14                .Item(1).Visible = True
15                .Item(2).Visible = False
16                .Item(3).Visible = True
17                .Item(4).Visible = False
18            End If
19        End With

20        Select Case intIndex

          Case 0                                                                  '费用
21            If mcolSubForm Is Nothing Then
22                Set mcolSubForm = New Collection
23            End If
24            If mclsExpenses Is Nothing Then
25                Set mclsExpenses = CreateObject("zlCISKernel.clsDockExpense")
26                mcolSubForm.Add mclsExpenses.zlGetForm, "_费用"             '得到子窗体
27            End If
28            With Me.TabCtlWindow
29                If .Item(intIndex).Handle = PicWindows.hWnd Then
30                    .RemoveItem (intIndex)
31                    .InsertItem(intIndex, "费用查询", mcolSubForm("_费用").hWnd, 0).Tag = "费用查询"
32                End If
33            End With

34            strSQL = "select a.id as 医嘱ID, b.发送号,b.执行部门ID from 病人医嘱记录 a,病人医嘱发送 b " & vbCrLf & _
                     " Where a.ID = b.医嘱id And a.相关id = [1] "
35            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "查看诊疗信息", Val(mrsInfo("申请ID") & ""))
36            If rsTmp.EOF = False Then
37                mclsExpenses.zlRefresh rsTmp("执行部门ID"), rsTmp("医嘱ID"), rsTmp("发送号")
38            End If
39        Case 1
40            If mcolSubForm Is Nothing Then
41                Set mcolSubForm = New Collection
42            End If
43            If mclsOutAdvices Is Nothing Then
44                Set mclsOutAdvices = CreateObject("zlCISKernel.clsDockOutAdvices")
45                mcolSubForm.Add mclsOutAdvices.zlGetForm, "_门诊医嘱"
46            End If
              '第一次打开时再加载
47            With Me.TabCtlWindow
48                If .Item(intIndex).Handle = PicWindows.hWnd Then

49                    .RemoveItem (intIndex)
50                    .InsertItem(intIndex, "门诊医嘱", mcolSubForm("_门诊医嘱").hWnd, 1).Tag = "门诊医嘱"
51                    .Item(intIndex).Selected = True
52                End If
53                strSQL = "select d.no 挂号单  from 病人挂号记录 d " & vbCrLf & _
                         " Where d.id = [1] "
54                Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "查看诊疗信息", strData)

55                mclsOutAdvices.zlRefresh Val(mrsInfo("病人ID") & ""), rsTmp("挂号单") & "", False
56                TabCtlWindow.Item(intIndex).Selected = True
57            End With
58        Case 2
59            If mcolSubForm Is Nothing Then
60                Set mcolSubForm = New Collection
61            End If
62            If mclsInAdvices Is Nothing Then
63                Set mclsInAdvices = CreateObject("zlCISKernel.clsDockInAdvices")
64                mcolSubForm.Add mclsInAdvices.zlGetForm, "_住院医嘱"
65            End If
              '第一次打开时再加载
66            With Me.TabCtlWindow
67                If .Item(intIndex).Handle = PicWindows.hWnd Then

68                    .RemoveItem (intIndex)
69                    .InsertItem(intIndex, "住院医嘱", mcolSubForm("_住院医嘱").hWnd, 1).Tag = "住院医嘱"
70                    .Item(intIndex).Selected = True
71                End If
72            End With
73            strSQL = "Select a.入院科室id 病人科室ID  ,a.当前病区id 病区ID From 病案主页 a where a.病人id =[1] and a.主页id =[2]"
74            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "查看诊疗信息", Val(mrsInfo("病人ID") & ""), strData)
75            If rsTmp.EOF = False Then
76                mclsInAdvices.zlRefresh Val(mrsInfo("病人ID") & ""), strData, Val(rsTmp("病区ID") & ""), _
                                          Val(rsTmp("病人科室ID") & ""), 0
77            End If
78            TabCtlWindow.Item(intIndex).Selected = True
79        Case 3
80            If mcolSubForm Is Nothing Then
81                Set mcolSubForm = New Collection
82            End If
83            If mclsOutEPRs Is Nothing Then
84                Set mclsOutEPRs = CreateObject("zlRichEPR.cDockOutEPRs")
85                mcolSubForm.Add mclsOutEPRs.zlGetForm, "_门诊病历"
86            End If
              '第一次打开时再加载
87            With Me.TabCtlWindow
88                If .Item(intIndex).Handle = PicWindows.hWnd Then

89                    .RemoveItem (intIndex)
90                    .InsertItem(intIndex, "门诊病历", mcolSubForm("_门诊病历").hWnd, 1).Tag = "门诊病历"
91                    .Item(intIndex).Selected = True
92                End If
93            End With
94            strSQL = "select a.id as 医嘱ID, b.发送号,b.执行部门ID,c.病区ID,a.病人科室ID,d.id 挂号ID from 病人医嘱记录 a,病人医嘱发送 b,病区科室对应 c,病人挂号记录 d " & vbCrLf & _
                     " Where a.ID = b.医嘱id and a.病人科室ID = 科室ID(+) and a.挂号单 = d.no And d.id = [1] "
95            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "查看诊疗信息", strData)
96            If rsTmp.EOF = False Then
97                mclsOutEPRs.zlRefresh Val(mrsInfo("病人ID") & ""), rsTmp("挂号ID"), Val(rsTmp("病人科室ID") & ""), False
98            Else
99                mclsOutEPRs.zlRefresh 0, 0, 0, False
100           End If
101           TabCtlWindow.Item(intIndex).Selected = True
102       Case 4
103           If mcolSubForm Is Nothing Then
104               Set mcolSubForm = New Collection
105           End If
106           If mclsInEPRs Is Nothing Then
107               Set mclsInEPRs = CreateObject("zlRichEPR.cDockInEPRs")
108               mcolSubForm.Add mclsInEPRs.zlGetForm, "_住院病历"
109           End If
              '第一次打开时再加载
110           With Me.TabCtlWindow
111               If .Item(intIndex).Handle = PicWindows.hWnd Then
112                   .RemoveItem (intIndex)
113                   .InsertItem(intIndex, "住院病历", mcolSubForm("_住院病历").hWnd, 1).Tag = "住院病历"
114                   .Item(intIndex).Selected = True
115               End If
116           End With
117           strSQL = "Select a.入院科室id 病人科室ID ,a.当前病区id  From 病案主页 a where a.病人id =[1] and a.主页id =[2]"
118           Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "查看诊疗信息", Val(mrsInfo("病人ID") & ""), strData)
119           If rsTmp.EOF = False Then
120               mclsInEPRs.zlRefresh Val(mrsInfo("病人ID") & ""), strData, Val(rsTmp("病人科室ID") & ""), False
121           Else
122               mclsInEPRs.zlRefresh 0, 0, 0, False
123           End If
124           TabCtlWindow.Item(intIndex).Selected = True
125       Case 5
126           If mcolSubForm Is Nothing Then
127               Set mcolSubForm = New Collection

128           End If
129           If mclsEMR Is Nothing Then
130               Set mclsEMR = CreateObject("zlRichEMR.clsDockEMR")
131               mcolSubForm.Add mclsEMR.zlGetForm, "_电子病历"
132           End If
133           If Not mclsEMR Is Nothing Then
134               If Not mclsEMR.Init(gobjEmr, gcnHisOracle, lngSysNo) Then
135                   Set mclsEMR = Nothing
136               End If
137           End If
              '第一次打开时再加载
138           With Me.TabCtlWindow
139               If .Item(intIndex).Handle = PicWindows.hWnd Then
140                   .RemoveItem (intIndex)
141                   .InsertItem(intIndex, "电子病历", mcolSubForm("_电子病历").hWnd, 1).Tag = "电子病历"
142                   .Item(intIndex).Selected = True
143               End If
144           End With
145           If mrsInfo("病人来源") & "" = 2 Then
146               strSQL = "Select a.入院科室id 病人科室ID ,a.当前病区id  From 病案主页 a where a.病人id =[2] and a.主页id =[1]"
147           Else
148               strSQL = "select a.id as 医嘱ID, b.发送号,b.执行部门ID,c.病区ID,a.病人科室ID,d.id 挂号ID from 病人医嘱记录 a,病人医嘱发送 b,病区科室对应 c,病人挂号记录 d " & vbCrLf & _
                         " Where a.ID = b.医嘱id and a.病人科室ID = 科室ID(+) and a.挂号单 = d.no And d.id = [1] "
149           End If
150           Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "查看诊疗信息", strData, Val(mrsInfo("病人ID") & ""))

151           If rsTmp.RecordCount > 0 Then
152               mclsEMR.zlRefresh Val(mrsInfo("病人ID") & ""), strData, Val(rsTmp("病人科室ID") & ""), 0, IIf(mrsInfo("病人来源") & "" = 2, 2, 1)
153           End If
154           TabCtlWindow.Item(intIndex).Selected = True
155       End Select

156       mblnLoadfrm = True

157       Exit Sub
RefreshTab_Error:
158       mblnLoadfrm = True

160       Call WriteErrLog("zl9LisInsideComm", "frmExpensesInfo", "执行(RefreshTab)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
161       Err.Clear
End Sub
'Public Function zlRefresh(ByVal lngPatiID As Long, ByVal lngBillId As Long, ByVal lngDeptID As Long, Optional ByVal bnEdit As Boolean, _
 '                            Optional ByVal blnMoved As Boolean, Optional ByVal blnForce As Boolean, Optional ByVal lngAdviceID As Long) As Long
'    '功能:调用刷新指定病人的病历内容，并根据情况提供编辑功能
'    '参数:  lngPatiId-病人id;
'    '       lngBillId-挂号id;
'    '       lngDeptId-当前操作部门，注意不是病人本次就诊科室；
'    '       blnEdit-是否允许编辑，通常当前操作部门不是病人本次就诊科室，就应该不允许编辑。
'    '       blnMoved-数据是否被转储
'    '       lngAdviceID 医嘱ID－目前只有手术模块调用传用
'    zlRefresh = frmOutEPRs.zlRefresh(lngPatiID, lngBillId, lngDeptID, bnEdit, blnForce, blnMoved, lngAdviceID)
'End Function
Public Sub ShowMe(lngSapmeID, objEMR As Object, parfrom As Object)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    gobjHisComLib.InitCommon gcnHisOracle
    gobjHisComLib.RegCheck
    Set gobjEmr = objEMR

    mlngSapmeID = lngSapmeID

    strSQL = "select id,HIS病人ID 病人ID,病人来源,申请ID,医嘱ID,门诊号,住院号,主页ID,挂号单,病人科室编码,病区编码,申请科室编码,姓名, decode(性别,1,'男',2,'女','未知') 性别,年龄,申请人,申请科室,标本类型 from 检验申请组合 where 标本id = [1] Order By 医嘱id "

    Set mrsInfo = ComOpenSQL(Sel_Lis_DB, strSQL, "", mlngSapmeID)

    If mrsInfo.RecordCount <= 0 Then
        Unload Me
        MsgBox "没有找到当前标本的诊疗信息,请检查!", vbInformation, "查看诊疗"
        Exit Sub
    Else
        If mrsInfo("病人来源") & "" = "" Or mrsInfo("医嘱ID") & "" = "" Then
            MsgBox "手工申请病人，不能查看诊疗信息", vbInformation, "查看诊疗"
            Exit Sub
        End If
    End If
    If mrsInfo("病人来源") & "" = 2 Then
        strSQL = "Select rownum as 序号,病人id,主页ID,NVL(病人性质,0) 病人性质,当前病区id,住院号,To_Char(入院日期,'YYYY-MM-DD HH24:MI') as 入院日期 From 病案主页 Where 主页ID<>0 And 病人ID=[1] Order by 主页ID Desc"
        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "中联信息", Val(mrsInfo("病人ID") & ""))
    Else
        strSQL = "Select A.ID,A.NO,A.发生时间 as 时间,B.名称 as 科室,a.执行人,a.接收时间 From 病人挂号记录 A,部门表 B" & _
               " Where A.执行部门ID=B.ID And A.病人ID=[1] And A.发生时间<=[2] And A.记录性质=1 And A.记录状态=1 Order by A.发生时间 Desc,a.接收时间 Desc"
        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "中联信息", Val(mrsInfo("病人ID") & ""), Now)


    End If
    If rsTmp.RecordCount = 0 Then
        Unload Me
        MsgBox "该病人没有找到诊疗信息!", vbInformation, "查看诊疗"
        Exit Sub
    End If
    Me.Show


End Sub

Private Sub TabCtlWindow_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
        Call RefreshTab(Item.Index)
End Sub
