VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmTriageManager 
   BorderStyle     =   0  'None
   Caption         =   "门诊分诊管理"
   ClientHeight    =   6600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10170
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picMain 
      BorderStyle     =   0  'None
      Height          =   3030
      Left            =   600
      ScaleHeight     =   3030
      ScaleWidth      =   3735
      TabIndex        =   5
      Top             =   600
      Width           =   3735
      Begin MSComctlLib.ListView lvwMain 
         Height          =   1770
         Left            =   -105
         TabIndex        =   6
         Tag             =   "可变化的"
         Top             =   1455
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   3122
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "img161"
         SmallIcons      =   "img161"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin MSComctlLib.ListView LvwYY 
         Height          =   1770
         Left            =   1695
         TabIndex        =   7
         Tag             =   "可变化的"
         Top             =   705
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   3122
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "img161"
         SmallIcons      =   "img161"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   4875
         Left            =   840
         TabIndex        =   8
         Top             =   240
         Width           =   2535
         _Version        =   589884
         _ExtentX        =   4471
         _ExtentY        =   8599
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picHZPati 
      BorderStyle     =   0  'None
      Height          =   2430
      Left            =   330
      ScaleHeight     =   2430
      ScaleWidth      =   2895
      TabIndex        =   3
      Top             =   4065
      Width           =   2895
      Begin MSComctlLib.ListView lvwHZPati 
         Height          =   1770
         Left            =   60
         TabIndex        =   4
         Tag             =   "可变化的"
         Top             =   30
         Width           =   2010
         _ExtentX        =   3545
         _ExtentY        =   3122
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "img161"
         SmallIcons      =   "img161"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.PictureBox picDept 
      BorderStyle     =   0  'None
      Height          =   5250
      Left            =   4815
      ScaleHeight     =   5250
      ScaleWidth      =   4485
      TabIndex        =   0
      Top             =   585
      Width           =   4485
      Begin MSComctlLib.ListView lvwRoom 
         Height          =   4230
         Left            =   0
         TabIndex        =   1
         Top             =   360
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   7461
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "img161"
         SmallIcons      =   "img161"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "诊室名称"
            Object.Width           =   2381
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "状态"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "候诊"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "在诊"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "当天已诊"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "科室"
            Object.Width           =   2293
         EndProperty
      End
      Begin XtremeSuiteControls.ShortcutCaption stcTittl 
         Height          =   315
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   2385
         _Version        =   589884
         _ExtentX        =   4207
         _ExtentY        =   556
         _StockProps     =   6
         Caption         =   "当前各诊室状态"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList img161 
      Left            =   885
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483637
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriageManager.frx":0000
            Key             =   "ry"
            Object.Tag             =   "ry"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriageManager.frx":059A
            Key             =   "yf"
            Object.Tag             =   "yf"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriageManager.frx":0934
            Key             =   "zz"
            Object.Tag             =   "zz"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriageManager.frx":0ECE
            Key             =   "yz"
            Object.Tag             =   "yz"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriageManager.frx":1468
            Key             =   "bm"
            Object.Tag             =   "bm"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriageManager.frx":1A02
            Key             =   "WomanStop"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriageManager.frx":1F9C
            Key             =   "ManStop"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriageManager.frx":2536
            Key             =   "WomanSign_in"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriageManager.frx":2AD0
            Key             =   "ManSign_in"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriageManager.frx":306A
            Key             =   "rySign_in"
            Object.Tag             =   "rySign_in"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlPaneIcons 
      Left            =   3960
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriageManager.frx":3604
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTriageManager.frx":3958
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmTriageManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String, mlngModul As Long, mintFindKeys As Integer
Private mbytViewScrop(0 To 3) As Byte  '0-显示已分诊病人;1-显示已接诊病人;2-显示已完成病人;3-显示不就诊病人
Private mbyt候诊排序方式 As Byte  '候诊病人的排序方式,0-科室编码,号码,单据号;1-科室编码,号码,挂号时间;
Private mstr分诊科室   As String               '以逗号分隔的分诊科室id串,空表示所有科室,0表示没有选择任何科室
Private mfrmMain As Form     '调用的父窗口
Private mlngOutModeMC As Long    '本地医保设置的外挂式医保险类
Private Const STR_COMP = "|',~" '分隔字符串
Private Type Type_SQLCondition
    Default As Boolean          '是否是缺省进入，此时没有条件值,缺省值在mstrFilter中
    DateB As Date
    DateE As Date
    NOB As String
    NOE As String
    FactB As String
    FactE As String
    DeptID As Long
    Patient As String
    Operator As String
    门诊号 As String
    就诊卡号 As String
    医保号 As String
End Type
Private mSQLCondition As Type_SQLCondition
'-----------------------------------------------------------------------------------
'消息相关变量
Private mobjMsgModule As clsMipModule
Attribute mobjMsgModule.VB_VarHelpID = -1

Private mcllFilter As Collection

Private mint有效天数 As Integer
Private mlngPre病人ID As Long   '上次病人ID
Private mbytIDKind As Byte  '0-门诊号;1-姓名;2-挂号单;3-就诊卡号;4-医保号
Private mlngDefaultCardID As Long '默认卡类别

Private Const conPane_PatiList = 1
Private Const conPane_Room = 2
Private Const conPane_PatiHZ = 3

Private Enum midx
    idx_排队队列 = 0
    idx_预约队列 = 1
End Enum

Public Event zlPopuMenu(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event zlShowInfor(strShowInfor As String)
Public Event zlQueueAsk(intType As Integer, strNO As String, lng病人ID As Long, Cancel As Boolean)
'intType:intType:1-分诊;2-换号;3-病人不就诊;4-病人待诊;5-病人完成就诊;6-恢复就诊;7-回诊
'strNO:-单据号
' 0-复诊,1-直呼,2-弃号,3-暂停,4-完成就诊,5-广播;7回诊

Private mcbsThis As Object
Private mobjPublicPatient As Object
Private Enum EnmCol
    Enm挂号单 = 0
    Enm病人状态 = 1
    Enm科室 = 2
    Enm号类 = 3
    Enm挂号项目 = 4
    Enm门诊号 = 5
    Enm姓名 = 6
    Enm性别 = 7
    Enm年龄 = 8
    Enm诊室 = 9
    Enm医生 = 10
    Enm发生时间 = 11
    Enm挂号时间 = 12
    Enm号序 = 13
    Enm医保号 = 14
    Enm摘要 = 15
    Enm预约 = 16
    Enm呼叫 = 17
    Enm呼叫人 = 18
    Enm呼叫诊室 = 19
    Enm呼叫时间 = 20
End Enum
'-----------------------------------------------------------------------------------
Private mstrRegistIdsed As String '已经刷新的挂号ID,多个用逗号分离,消息发送时有效


Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.id
    Case conPane_PatiList
        Item.Handle = picMain.Hwnd
    Case conPane_Room
        Item.Handle = picDept.Hwnd
    Case conPane_PatiHZ
        Item.Handle = picHZPati.Hwnd
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If mfrmMain Is Nothing Then Exit Sub
    Call mfrmMain.ActiveIDKindKey
End Sub

Private Sub InitPancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:区哉设置
    '编制:刘兴洪
    '日期:2009-09-14 18:06:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim sngWidth As Single, strReg As String, panThis As Pane
    Dim panLeft As Pane
    
    Set panLeft = dkpMan.CreatePane(conPane_PatiList, 200, 580, DockLeftOf, Nothing)
    panLeft.Title = "病人列表": panLeft.Tag = conPane_PatiList
    panLeft.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panLeft.Handle = picMain.Hwnd
    
    Set panThis = dkpMan.CreatePane(conPane_Room, 250, 580, DockRightOf, panLeft)
    panThis.Title = "诊室情况"
    panThis.Tag = conPane_Room
    panThis.Handle = picDept.Hwnd
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    
    Set panThis = dkpMan.CreatePane(conPane_PatiHZ, 200, 580, DockBottomOf, panLeft)
    panThis.Title = "回诊病人列表": panThis.Tag = panThis
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panThis.Handle = picHZPati.Hwnd
    
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.HideClient = True
    
    zlRestoreDockPanceToReg Me, dkpMan, "区域"
End Sub

 

Private Sub picMain_Resize()
    Err = 0: On Error Resume Next
    With picMain
        tbPage.Left = .ScaleLeft
        tbPage.Width = .ScaleWidth
        tbPage.Top = .ScaleTop
        tbPage.Height = .ScaleHeight
    End With
End Sub
Private Sub picHZPati_Resize()
    Err = 0: On Error Resume Next
    With picHZPati
        lvwHZPati.Left = .ScaleLeft
        lvwHZPati.Width = .ScaleWidth
        lvwHZPati.Top = .ScaleTop
        lvwHZPati.Height = .ScaleHeight
    End With
End Sub
Private Sub picDept_Resize()
    Err = 0: On Error Resume Next
    With picDept
        lvwRoom.Left = .ScaleLeft
        lvwRoom.Width = .ScaleWidth
        stcTittl.Top = .ScaleTop: stcTittl.Left = .ScaleLeft
        stcTittl.Width = .ScaleWidth
        lvwRoom.Top = stcTittl.Top + stcTittl.Height
        lvwRoom.Height = .ScaleHeight - lvwRoom.Top
    End With
End Sub
Public Sub zlInitVar(ByVal frmMain As Form, ByVal byt候诊排序方式 As Byte)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：设置常用变量
    '入参：frmMain-调用的父窗体
    '出参：
    '返回：
    '编制：刘兴洪
    '日期：2010-06-01 17:21:15
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    mbyt候诊排序方式 = byt候诊排序方式
    Set mfrmMain = frmMain
End Sub
Public Sub zlExcuteReport(ByVal lngSys As Long, ByVal strReportNO As String)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：执行报表
    '编制：刘兴洪
    '日期：2010-06-01 15:53:17
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    If Not lvwMain.SelectedItem Is Nothing Then
         With lvwMain.SelectedItem
             Call ReportOpen(gcnOracle, lngSys, strReportNO, Me, _
                 "NO=" & .Text, "门诊号=" & .SubItems(EnmCol.Enm门诊号), _
                 "医生=" & .SubItems(EnmCol.Enm医生), "执行科室=" & .ListSubItems(2).Tag)
         End With
     Else
         Call ReportOpen(gcnOracle, lngSys, strReportNO, Me)
     End If
End Sub

Public Sub zlExc签道(ByVal bln取消签到 As Boolean, Optional ByVal blnClick As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行签道
    '编制:刘兴洪
    '日期:2010-12-08 10:56:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objLvw As ListView, blnHZ As Boolean '是否回诊
    Dim blnTriage As Boolean, lng病人ID  As Long, lngExeState As Long
    Dim lngID As Long, strTittle As String
    Dim bln已分诊 As Boolean, strNO As String, bln是否预约 As Boolean
    Dim str发生时间 As String
    Dim strDoc As String, strRoom As String
    Dim bln分诊台签到排队 As Boolean
    
    If tbPage.Item(midx.idx_排队队列).Selected Then
        Set objLvw = lvwMain
        If objLvw.SelectedItem Is Nothing Then Exit Sub
        If Val(Split(objLvw.SelectedItem.Tag, "|")(8)) = 0 Then
            '转诊病人都是已分诊病人
            bln已分诊 = True
        End If
        bln分诊台签到排队 = objLvw.SelectedItem.ListSubItems(4).Tag = 1
    End If
    
    If tbPage.Item(midx.idx_预约队列).Selected Then
        Set objLvw = LvwYY
        If objLvw.SelectedItem Is Nothing Then Exit Sub
        bln是否预约 = True
        bln分诊台签到排队 = objLvw.SelectedItem.ListSubItems(4).Tag = 1
    End If
    
    lng病人ID = Val(Split(objLvw.SelectedItem.Tag, "|")(0))
    lngExeState = Val(Split(objLvw.SelectedItem.Tag, "|")(6)) '执行状态
    
    blnTriage = (lngExeState = 0)
    lngID = Val(objLvw.SelectedItem.ListSubItems(1).Tag)
    strDoc = lvwMain.SelectedItem.SubItems(EnmCol.Enm医生)
    strNO = Trim(objLvw.SelectedItem.Text)
    str发生时间 = objLvw.SelectedItem.SubItems(EnmCol.Enm发生时间)
    Err = 0: On Error GoTo Errhand:
    If lngID = 0 Then Exit Sub
    
    If blnTriage Then
        If objLvw.SelectedItem.SubItems(EnmCol.Enm诊室) <> "" Then bln已分诊 = True
    End If
    
    '125454:李南春，2018/5/18，手动签到才需要提示，自动签到的时候不能直接确定签到都直接返回false但不提示
    '95637:李南春,2016/7/18,签到需检查当前号别是否在排队中，或者当天有其他号别处于排队中
    If Not bln取消签到 Then
        If Check签到(bln是否预约, lng病人ID, lngID, str发生时间, blnClick, bln分诊台签到排队) = False Then Exit Sub
    End If
    
    If Not bln已分诊 And Not bln取消签到 Then
        '未分诊部分
        zlExecuteTriage mfrmMain, True: Exit Sub
        objLvw.SelectedItem.ListSubItems(3).Tag = IIf(bln取消签到, 0, 1)
        objLvw.SelectedItem.Icon = "rySign_in"
        objLvw.SelectedItem.SmallIcon = "rySign_in"
        Exit Sub
    End If
    
    strRoom = objLvw.SelectedItem.SubItems(EnmCol.Enm诊室)
    
    If ExcPlugInFun(IIf(bln取消签到, 14, 4), lngID, strDoc, strRoom) = False Then Exit Sub
    
    
    'intType:---操作类型_IN:0-签到;1-取消签到/取消医生标记回诊;2-医生标记回诊/取消分诊台回诊签道;3-分诊台回诊签道
    If zl签到或取消(Not bln取消签到, lngID, strNO) = False Then
        strTittle = IIf(bln取消签到, "取消签到失败!", "签到失败!")
        RaiseEvent zlShowInfor(strTittle)
        ShowMsgbox strTittle
        Exit Sub
    End If
    If bln取消签到 = False Then '问题:38165
        Call zlPrintBill(lngID)
        '77412:李南春，2014/9/3,门诊病人条码打印
        Call zlPrintBarcode
    End If
        
    strTittle = IIf(bln取消签到, "取消签到成功!", "签到成功!")
    RaiseEvent zlShowInfor(strTittle)
    ShowMsgbox strTittle
    objLvw.SelectedItem.ListSubItems(3).Tag = IIf(bln取消签到, 0, 1)
    If bln取消签到 Then
        If bln已分诊 Then
            objLvw.SelectedItem.Icon = "yf"
            objLvw.SelectedItem.SmallIcon = "yf"
        Else
            objLvw.SelectedItem.Icon = "ry"
            objLvw.SelectedItem.SmallIcon = "ry"
        End If
    Else
        objLvw.SelectedItem.Icon = "rySign_in"
        objLvw.SelectedItem.SmallIcon = "rySign_in"
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Function zl签到或取消(bln签到 As Boolean, ByVal lng挂号ID As Long, ByVal strNO As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查单据中输入的负数数量及退回科室是否正确
    '入参:int签到类型 0-正常签到；1-回诊签到;2-转诊签到;3-换号签到;4-回诊重新签到;5-其他业务重新签到
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2011-01-16 13:56:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, strSQL As String
    On Error GoTo errHandle
    If bln签到 Then
        'Zl_病人挂号记录_签到
        '  Id_In     病人挂号记录.ID%Type,
        '  仅签到_In Integer:=0
        strSQL = "Zl_病人挂号记录_签到(" & lng挂号ID & "," & 0 & ",'" & zl_Get预约方式ByID(lng挂号ID) & "')" '问题号:48350
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        'intType:intType:1-分诊;2-换号;3-病人不就诊;4-病人待诊;5-病人完成就诊;6-恢复就诊,7-回诊
        RaiseEvent zlQueueAsk(1, strNO, mlngPre病人ID, blnCancel)
        If blnCancel Then Exit Function
        zl签到或取消 = True
        Exit Function
    End If
    'Zl_病人挂号记录_取消签到
    '  Id_In           病人挂号记录.ID%Type,
    '  仅改签到标志_In Integer:=0
    strSQL = "Zl_病人挂号记录_取消签到(" & lng挂号ID & "," & 0 & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    zl签到或取消 = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 
Public Sub zlExcuteFunction()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：执行相关功能
    '编制：刘兴洪
    '日期：2010-05-31 16:42:44
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim blnTriage As Boolean, lng病人ID  As Long, lngExeState As Long
    Dim objLvw As ListView
        
    'ListSubItem(3).tag:A.记录标志:0表示就诊病人,1-表示签到的病人,2-表示需要回诊就诊的病人; 3-表示已回诊但还未接收的病人;
    'ListSubItem(4).tag:A.记录标志:0-表示不启用分诊台签到排队,1-表示启用分诊台签到排队;
    
    If Val(zlDatabase.GetPara("免挂号模式", glngSys)) = 1 Then Exit Sub
    If Not (Me.ActiveControl Is lvwMain Or Me.ActiveControl Is LvwYY Or Me.ActiveControl Is lvwHZPati) Then Exit Sub
    Set objLvw = Me.ActiveControl
    If objLvw.SelectedItem Is Nothing Then
        blnTriage = False: lng病人ID = 0
    Else
        lng病人ID = Val(Split(objLvw.SelectedItem.Tag, "|")(0))
        '!病人ID & "|" & !险类 & "|" & !就诊卡号 & "|" & !卡验证码 & "|" & !ID & "|" & !号别 & "|" & !执行状态
        lngExeState = Val(Split(objLvw.SelectedItem.Tag, "|")(6))
        blnTriage = (lngExeState = 0) And lng病人ID <> 0
    End If
    If blnTriage Then
        If Val(objLvw.SelectedItem.ListSubItems(4).Tag) = 1 Then
            'ListSubItem(3).tag:A.记录标志:表示0表示就诊病人,1-表示签到的病人,2-表示需要回诊就诊的病人; 3-表示已回诊但还未接收的病人;
            If Val(objLvw.SelectedItem.ListSubItems(3).Tag) = 0 Then
                GoTo EdPati:
            End If
        End If
        zlExecuteTriage mfrmMain: Exit Sub
    End If
EdPati:
    
    If lng病人ID = 0 Then
         zlExcuteEditPati mfrmMain: Exit Sub
    End If
End Sub

Public Sub zlSubPrint(bytMode As Byte)
    '功能:进行打印,预览和输出到EXCEL
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    On Error GoTo errHandle
    Dim objPrint As New zlPrintLvw
    If Me.ActiveControl Is lvwHZPati Then
        objPrint.Title.Text = "回诊病人清册"
        Set objPrint.Body.objData = lvwHZPati
    
    ElseIf tbPage.Item(midx.idx_排队队列).Selected Then
        objPrint.Title.Text = Me.Caption
        Set objPrint.Body.objData = lvwMain
    Else
        objPrint.Title.Text = "预约病人清册"
        Set objPrint.Body.objData = LvwYY
    End If
    objPrint.BelowAppItems.Add "打印人：" & UserInfo.姓名
    objPrint.BelowAppItems.Add "打印时间：" & Format(zlDatabase.Currentdate, "yyyy年MM月dd日")
    If bytMode = 1 Then
        Select Case zlPrintAsk(objPrint)
        Case 1
            zlPrintOrViewLvw objPrint, 1
        Case 2
            zlPrintOrViewLvw objPrint, 2
        Case 3
            zlPrintOrViewLvw objPrint, 3
        End Select
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Sub zlExecuteTriage(ByVal frmMain As Object, _
    Optional bln签到 As Boolean = False, _
    Optional blnAppointment As Boolean = False)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：执行分诊
    '入参: blnAppointment-是否对预约病人进行分诊
    '编制：刘兴洪
    '日期：2010-05-31 14:54:48
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim strRoom As String, strDate As String
    Dim strDoctor As String, strSQL As String, rsTemp As ADODB.Recordset
    Dim i As Long, blnCancel As Boolean, lngID As Long
    Dim cllPro As Collection
    Dim lng记录性质 As Long, bln预约 As Boolean, lng病人ID As Long
    Dim strNO As String
     
    On Error GoTo errH
    
    If tbPage.Item(midx.idx_排队队列).Selected Then
        '没选择就退出
        If lvwMain.SelectedItem Is Nothing Then Exit Sub
    Else
        If LvwYY.SelectedItem Is Nothing Then Exit Sub
    End If
    
    lng记录性质 = IIf(tbPage.Item(midx.idx_排队队列).Selected, 1, 2)
    
    If tbPage.Item(midx.idx_排队队列).Selected Then
        strNO = lvwMain.SelectedItem.Text
        strDoctor = lvwMain.SelectedItem.SubItems(EnmCol.Enm医生)
        lngID = Val(lvwMain.SelectedItem.ListSubItems(1).Tag)
        lng病人ID = Val(Split(lvwMain.SelectedItem.Tag, "|")(0))
    Else
        strNO = LvwYY.SelectedItem.Text
        strDoctor = LvwYY.SelectedItem.SubItems(EnmCol.Enm医生)
        lngID = Val(LvwYY.SelectedItem.ListSubItems(1).Tag)
        lng病人ID = Val(Split(lvwMain.SelectedItem.Tag, "|")(0))
        bln预约 = True
    End If
    
    ReadRoom
    strRoom = STR_COMP
    If frmDistRoom.cmb.ListCount > 0 Then
        frmDistRoom.cmb.ListIndex = 0
        For i = 0 To frmDistRoom.cmb.ListCount - 1
            If frmDistRoom.cmb.List(i) Like "*-" & Trim(strDoctor) Then
                frmDistRoom.cmb.ListIndex = i
                Exit For
            End If
        Next
    End If
    frmDistRoom.ShowMe strRoom, Me, bln签到, lngID
    If strRoom = STR_COMP Then
        RaiseEvent zlShowInfor("用户取消!") '选择了"取消"了
        Exit Sub
    End If
    'NO_IN       病人挂号记录.NO%TYPE:=NULL,
    '病人ID_IN   病人挂号记录.病人id%TYPE:=NULL,
    '诊室_IN     病人挂号记录.诊室%TYPE:=NULL
    '
    Set cllPro = New Collection
    strDoctor = Trim(Split(strRoom, STR_COMP)(1))
    strDate = "To_date('" & Split(strRoom, STR_COMP)(2) & "','yyyy-mm-dd hh24:mi:ss')"
    strRoom = Split(strRoom, STR_COMP)(0)
    
    '111121:焦博，2017/7/17,重复调用"Zl_病人挂号记录_签到"
    '问题号:48350
    strSQL = "ZL_病人挂号记录_更新诊室 ('" & strNO & "'," & mlngPre病人ID & ",'" & strRoom & "','" & strDoctor & "'," & strDate & ",'1','" & zl_Get预约方式ByNo(strNO) & "')"
    zlAddArray cllPro, strSQL
    
    If bln签到 Then
        'Zl_病人挂号记录_签到
        '  Id_In     病人挂号记录.ID%Type,
        '  仅签到_In Integer:=0
        '问题号:48350
        strSQL = "Zl_病人挂号记录_签到(" & lngID & "," & 0 & ",'" & zl_Get预约方式ByID(lngID) & "')"
        zlAddArray cllPro, strSQL
    End If
    
    Err = 0: On Error GoTo Errhand:
    zlExecuteProcedureArrAy cllPro, Me.Caption
    
    If bln签到 = False Then
      '分诊触发分诊消息
      Call SendMsgModule(strNO)
    End If
    
    RaiseEvent zlQueueAsk(1, strNO, mlngPre病人ID, blnCancel)
    'intType:intType:1-分诊;2-换号;3-病人不就诊;4-病人待诊;5-病人完成就诊;6-恢复就诊
    ' 0-复诊,1-直呼,2-弃号,3-暂停,4-完成就诊,5-广播
    Err = 0: On Error GoTo errH:
     '显示出队列号
     strSQL = " Select A.排队号码,B.姓名 From 排队叫号队列 A,病人挂号记录 B Where a.业务ID= B.ID and  A.业务id = [1] And A.业务类型 = 0 and b.记录性质=[2] and b.记录状态=1  "
     Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID, lng记录性质)
     strDate = ""
     If Not rsTemp.EOF Then
        strDate = Nvl(rsTemp!姓名) & " 的排队号码为:" & Nvl(rsTemp!排队号码)
     End If
     '77412:李南春，2014/9/3,门诊病人条码打印
    Call zlPrintBarcode
    Call zlRefreshData
    If bln签到 Then  '问题:38165
        Call zlPrintBill(lngID)
    End If
     RaiseEvent zlShowInfor(strDate) '显示号别
     Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Exit Sub
Errhand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Public Sub zlExcuteChangeNum(ByVal frmMain As Form)
    '病人换号
    Dim blnCancel As Boolean, strNO As String, lngID As Long
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    
    strNO = Trim(lvwMain.SelectedItem.Text)
    lngID = Split(lvwMain.SelectedItem.Tag, "|")(4)
    
    If InStr(lvwMain.SelectedItem.Tag, "|") > 0 Then
        If gbytRegistMode = 0 Then
            If frmChangeNum.ShowMe(lngID, Me) Then
                RaiseEvent zlQueueAsk(2, strNO, mlngPre病人ID, blnCancel)
                'intType:intType:1-分诊;2-换号;3-病人不就诊;4-病人待诊;5-病人完成就诊;6-恢复就诊
                ' 0-复诊,1-直呼,2-弃号,3-暂停,4-完成就诊,5-广播
                
                '换号触发换号消息
                Call SendMsgModule(strNO)
                Call zlRefreshData
            End If
        Else
            If Sys.Currentdate < gdatRegistTime Then
                If frmChangeNum.ShowMe(lngID, Me) Then
                    RaiseEvent zlQueueAsk(2, strNO, mlngPre病人ID, blnCancel)
                    'intType:intType:1-分诊;2-换号;3-病人不就诊;4-病人待诊;5-病人完成就诊;6-恢复就诊
                    ' 0-复诊,1-直呼,2-弃号,3-暂停,4-完成就诊,5-广播
                    
                    '换号触发换号消息
                    Call SendMsgModule(strNO)
                    Call zlRefreshData
                End If
            Else
                If frmChangeNumNew.ShowMe(lngID, Me) Then
                    RaiseEvent zlQueueAsk(2, strNO, mlngPre病人ID, blnCancel)
                    'intType:intType:1-分诊;2-换号;3-病人不就诊;4-病人待诊;5-病人完成就诊;6-恢复就诊
                    ' 0-复诊,1-直呼,2-弃号,3-暂停,4-完成就诊,5-广播
                    
                    '换号触发换号消息
                    Call SendMsgModule(strNO)
                    Call zlRefreshData
                End If
            End If
        End If
    End If
End Sub
Public Sub zlExcuteEditPati(ByVal frmMain As Form)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：编辑病人档案信息
    '编制：刘兴洪
    '日期：2010-05-31 15:41:06
    '说明：
    '------------------------------------------------------------------------------------------------------------------------

    On Error GoTo errH
    Dim lng病人ID As Long, lng险类 As Long, bln入院 As Boolean
    Dim i As Long
    Dim strNO As String
    Dim str验证密码 As String
    Dim str就诊卡号 As String
    
    If tbPage.Item(midx.idx_排队队列).Selected Then
        '没选择就退出
        If lvwMain.SelectedItem Is Nothing Then Exit Sub
    Else
        If LvwYY.SelectedItem Is Nothing Then Exit Sub
    End If
    
   If tbPage.Item(midx.idx_排队队列).Selected Then
        lng病人ID = CLng(Split(lvwMain.SelectedItem.Tag, "|")(0))
        lng险类 = CLng(Split(lvwMain.SelectedItem.Tag, "|")(1))
        strNO = lvwMain.SelectedItem.Text
        str就诊卡号 = Split(lvwMain.SelectedItem.Tag, "|")(2)
        str验证密码 = Split(lvwMain.SelectedItem.Tag, "|")(3)
    Else
        lng病人ID = CLng(Split(LvwYY.SelectedItem.Tag, "|")(0))
        lng险类 = CLng(Split(LvwYY.SelectedItem.Tag, "|")(1))
        strNO = LvwYY.SelectedItem.Text
        str就诊卡号 = Split(LvwYY.SelectedItem.Tag, "|")(2)
        str验证密码 = Split(LvwYY.SelectedItem.Tag, "|")(3)
    End If


    With frmDistRoomPatiEdit
        .mstrNo = strNO
        .mlng病人ID = lng病人ID
        .mlng险类 = lng险类
        .mstrPrivs = mstrPrivs
        .mlngModul = mlngModul
        '外挂式医保没有险类
        If lng险类 = 0 And mlngOutModeMC > 0 Then
            .mlngOutModeMC = mlngOutModeMC
        Else
            .mlngOutModeMC = 0
        End If
        .m就诊卡号 = str就诊卡号
        .m验证密码 = str验证密码
        .InitData
        .Init过敏药物
        '79912:李南春,2014/11/20,入院病人不能在分诊台修改信息
        Call LoadPatientInfo(lng病人ID, bln入院)
        If bln入院 Then
            MsgBox "该病人已入院,请至『病人入院管理』修改信息。", vbInformation, gstrSysName
            Exit Sub
        End If
        If lng病人ID <= 0 Then
            If Not .GetRegBillID() Then
                MsgBox "无法获取挂号ID", vbInformation, gstrSysName
                 Unload frmDistRoomPatiEdit
                Exit Sub
            End If
        End If
        '67070:刘尔旋,2013-11-04,读取病人体征信息
        .UCPatiVitalSigns.LoadPatiVitalSigns .mlng病人ID, .mlng挂号ID
        Call .SetPatiBaseInforEnabled
        If mfrmMain Is Nothing Then
            .Show 1, Me
        Else
            .Show 1, mfrmMain
        End If
    End With
    
    '重新刷新
    If gblnOk Then zlRefreshData (True)

    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Sub zlExcutePatiLeave(ByVal frmMain As Form)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：病人不就诊
    '编制：刘兴洪
    '日期：2010-05-31 15:46:16
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Call Set病人挂号状态(-1)
End Sub
Public Sub zlExcutePatiWait(ByVal frmMain As Form)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：病人待诊
    '编制：刘兴洪
    '日期：2010-05-31 15:50:34
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
     Call Set病人挂号状态(0)
End Sub

Public Sub zlExcutePatiCancelOver(ByVal frmMain As Form)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：取消完成就诊
    '编制：刘兴洪
    '日期：2010-05-31 15:56:58
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim strMsgbox As String, strSQL As String, lng病人ID As Long, lngExeState As Long
    Dim blnCancel As Boolean, lngID As Long
    
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    If InStr(mstrPrivs, "完成就诊") = 0 Then Exit Sub
    lng病人ID = Val(Split(lvwMain.SelectedItem.Tag, "|")(0))
    lngExeState = Val(Split(lvwMain.SelectedItem.Tag, "|")(6))
    If lngExeState <> 1 Then Exit Sub
    If lng病人ID = 0 Then MsgBox "不存在的病人！", vbInformation, gstrSysName: Exit Sub
    
    strMsgbox = "正常情况下，应该由医生在必要时决定是否取消接诊完成，" & vbCrLf & _
                "除非是有护士在分诊台直接标记的病人接诊完成，否则不能进行该操作！" & vbCrLf & vbCrLf & _
                "真的要取消完成吗？"
    If MsgBox(strMsgbox, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    If ExcPlugInFun(13, Val(lvwMain.SelectedItem.ListSubItems(1).Tag)) = False Then Exit Sub


    Err = 0: On Error GoTo errHandle
    gcnOracle.BeginTrans
    strSQL = "zl_病人接诊完成_Cancel(" & Split(lvwMain.SelectedItem.Tag, "|")(0) & ",'" & lvwMain.SelectedItem.Text & "',1)"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    'intType:1-分诊;2-换号;3-病人不就诊;4-病人待诊;5-病人完成就诊;6-恢复就诊,7-回诊
   RaiseEvent zlQueueAsk(6, Trim(lvwMain.SelectedItem.Text), mlngPre病人ID, blnCancel)
    ' 0-复诊,1-直呼,2-弃号,3-暂停,4-完成就诊,5-广播
    If blnCancel = True Then
        gcnOracle.RollbackTrans: Exit Sub
    End If
    gcnOracle.CommitTrans
    
    Exit Sub
errHandle:
     gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub Set病人挂号状态(ByVal lngState As Long)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：设置病人挂号状态
    '入参：lngState : -1- 病人不就诊
    '                         0-病人待诊
    '出参：
    '返回：
    '编制：刘兴洪
    '日期：2010-06-03 15:24:48
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, blnCancel As Boolean
    blnCancel = False

    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    
    If ExcPlugInFun(IIf(lngState = -1, 3, 6), Val(lvwMain.SelectedItem.ListSubItems(1).Tag)) = False Then Exit Sub

    On Error GoTo errH
    gcnOracle.BeginTrans
    strSQL = "Zl_病人挂号记录_状态 ('" & lvwMain.SelectedItem.Text & "'," & lngState & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    RaiseEvent zlQueueAsk(IIf(lngState = -1, 3, 4), Trim(lvwMain.SelectedItem.Text), mlngPre病人ID, blnCancel)
    'intType:intType:1-分诊;2-换号;3-病人不就诊;4-病人待诊;5-病人完成就诊;6-恢复就诊
    ' 0-复诊,1-直呼,2-弃号,3-暂停,4-完成就诊,5-广播
    If blnCancel = True Then
        gcnOracle.RollbackTrans: Exit Sub
    End If
    
    gcnOracle.CommitTrans
    MsgBox "操作成功!", vbInformation, gstrSysName

    Call zlRefreshData
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub ReadRoom()
    On Error GoTo ErrHead
    Dim rsTmp As ADODB.Recordset
    Dim objListItem As ListItem
    Dim i As Long, lngSel As Long
    Dim strNO As String, strSQL As String
    Dim lng记录性质 As Long
    Dim strTmp As String
    Dim blnBusy As Boolean
    Dim lng转诊科室ID As Long
    '没有选择就退出
    If tbPage.Item(midx.idx_排队队列).Selected Then
        '没选择就退出
        If lvwMain.SelectedItem Is Nothing Then Exit Sub
        strNO = lvwMain.SelectedItem.Text
        
    Else
        If LvwYY.SelectedItem Is Nothing Then Exit Sub
        strNO = LvwYY.SelectedItem.Text
    End If
    
    lng记录性质 = IIf(tbPage.Item(midx.idx_排队队列).Selected, 1, 2)
    
    frmDistRoom.lvwMain.ListItems.Clear
    frmDistRoom.cmb.Clear
   
    
    
    '先增加医生
    '95637：李南春，2016/7/17，转诊签到,以转诊科室获取医生信息
    strSQL = _
        " Select c.编号,c.姓名,Nvl(d.转诊科室ID,0) as 转诊科室ID From 人员性质说明 a, 部门人员 b ,人员表 c,病人挂号记录 d" & vbCrLf & _
        " Where b.人员id=c.id And b.人员id=a.人员id And a.人员性质='医生' " & vbCrLf & _
        " And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & vbNewLine & _
        " And b.部门id=Nvl(d.转诊科室ID,d.执行部门ID) And d.记录性质=[2] and d.记录状态=1 and  d.NO=[1] And (c.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or c.撤档时间 Is Null)"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, lng记录性质)
    frmDistRoom.cmb.AddItem "无"
    If rsTmp.RecordCount > 0 Then
        lng转诊科室ID = Val(Nvl(rsTmp!转诊科室ID))
        For i = 1 To rsTmp.RecordCount
            frmDistRoom.cmb.AddItem zlCommFun.Nvl(rsTmp!编号) & "-" & zlCommFun.Nvl(rsTmp!姓名)
            rsTmp.MoveNext
        Next
    End If

    '调整为当前医生
    If gbytRegistMode = 0 Then
        strSQL = "Select A.医生姓名,B.登记时间,sysdate 当前时间 From 挂号安排 A,病人挂号记录 B Where A.号码=B.号别 And B.NO=[1] and b.记录性质=[2] and b.记录状态=1"
    Else
        If Sys.Currentdate < gdatRegistTime Then
            strSQL = "Select A.医生姓名,B.登记时间,sysdate 当前时间 From 挂号安排 A,病人挂号记录 B Where A.号码=B.号别 And B.NO=[1] and b.记录性质=[2] and b.记录状态=1"
        Else
            strSQL = "Select A.医生姓名,B.登记时间,sysdate 当前时间 From 临床出诊记录 A,病人挂号记录 B Where A.ID=B.出诊记录ID And B.NO=[1] and b.记录性质=[2] and b.记录状态=1"
        End If
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, lng记录性质)
    If rsTmp.RecordCount > 0 Then
        For i = 0 To frmDistRoom.cmb.ListCount - 1
            If frmDistRoom.cmb.List(i) Like "*-" & zlCommFun.Nvl(rsTmp!医生姓名) Then
                frmDistRoom.cmb.ListIndex = i
                Exit For
            End If
        Next
        frmDistRoom.dtpBegin.MaxDate = rsTmp!当前时间
        frmDistRoom.dtpBegin.MinDate = rsTmp!登记时间
        frmDistRoom.dtpBegin.Value = rsTmp!当前时间
    End If

    '读出所有的该号类的诊室供选择
    '79694:李南春,2014/11/25,根据参数读取门诊诊室
    blnBusy = Val(zlDatabase.GetPara("诊室忙时允许分诊", glngSys, mlngModul, 0)) = 1
    If lng转诊科室ID <> 0 Then
        '95637:李南春,2016/7/18,发生了转诊，只能以转诊科室去确定诊室
        If gbytRegistMode = 0 Then
            strSQL = _
                " Select Distinct b.编码, b.名称, b.位置" & vbNewLine & _
                " From 挂号安排诊室 a, 门诊诊室 b, 挂号安排 c" & vbNewLine & _
                " Where a.门诊诊室 = b.名称 And a.号表id = c.Id And c.ID IN (Select ID From 挂号安排 Where 科室ID=[3]) " & _
                IIf(blnBusy, " ", " And b.缺省标志=0 ") & _
                " Order By B.编码 "
        Else
            If Sys.Currentdate < gdatRegistTime Then
                strSQL = _
                    " Select Distinct b.编码, b.名称, b.位置" & vbNewLine & _
                    " From 挂号安排诊室 a, 门诊诊室 b, 挂号安排 c" & vbNewLine & _
                    " Where a.门诊诊室 = b.名称 And a.号表id = c.Id And c.ID IN (Select ID From 挂号安排 Where 科室ID=[3]) " & _
                    IIf(blnBusy, " ", " And b.缺省标志=0 ") & _
                    " Order By B.编码"
            Else
                strSQL = _
                    " Select Distinct b.编码, b.名称, b.位置" & vbNewLine & _
                    " From 门诊诊室适用科室 a, 门诊诊室 b" & vbNewLine & _
                    " Where a.诊室id = b.id And a.科室ID=[3] " & _
                    IIf(blnBusy, " ", " And b.缺省标志=0 ") & _
                    " Order By B.编码"
            End If
        End If
    Else
        If gbytRegistMode = 0 Then
            strSQL = _
                " Select b.编码, b.名称, b.位置" & vbNewLine & _
                " From 挂号安排诊室 a, 门诊诊室 b, 挂号安排 c, 病人挂号记录 d" & vbNewLine & _
                " Where a.门诊诊室 = b.名称 And a.号表id = c.Id And c.号码 = d.号别 And d.No = [1] " & _
                IIf(blnBusy, " ", " And b.缺省标志=0 ") & " and d.记录性质=[2] and d.记录状态=1" & _
                    " Order By B.编码"
        Else
            If Sys.Currentdate < gdatRegistTime Then
                strSQL = _
                    " Select b.编码, b.名称, b.位置" & vbNewLine & _
                    " From 挂号安排诊室 a, 门诊诊室 b, 挂号安排 c, 病人挂号记录 d" & vbNewLine & _
                    " Where a.门诊诊室 = b.名称 And a.号表id = c.Id And c.号码 = d.号别 And d.No = [1] " & _
                    IIf(blnBusy, " ", " And b.缺省标志=0 ") & " and d.记录性质=[2] and d.记录状态=1" & _
                    " Order By B.编码"
            Else
                strSQL = "Select 分诊方式 From 临床出诊记录 A,病人挂号记录 B Where B.NO = [1] And B.记录性质 = [2] And A.ID = B.出诊记录ID"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, lng记录性质)
                If rsTmp.EOF Then
                    strSQL = _
                        " Select b.编码, b.名称, b.位置" & vbNewLine & _
                        " From 门诊诊室适用科室 a, 门诊诊室 b, 病人挂号记录 d" & vbNewLine & _
                        " Where a.诊室id = b.id And a.科室id = d.执行部门ID And d.No = [1] " & _
                        IIf(blnBusy, " ", " And b.缺省标志=0 ") & " and d.记录性质=[2] and d.记录状态=1"
                Else
                    If Val(Nvl(rsTmp!分诊方式)) = 0 Then
                        strSQL = _
                            " Select b.编码, b.名称, b.位置" & vbNewLine & _
                            " From 门诊诊室适用科室 a, 门诊诊室 b, 病人挂号记录 d" & vbNewLine & _
                            " Where a.诊室id = b.id And a.科室id = d.执行部门ID And d.No = [1] " & _
                            IIf(blnBusy, " ", " And b.缺省标志=0 ") & " and d.记录性质=[2] and d.记录状态=1"
                    Else
                        strSQL = _
                            " Select b.编码, b.名称, b.位置" & vbNewLine & _
                            " From 临床出诊诊室记录 a, 门诊诊室 b, 病人挂号记录 d" & vbNewLine & _
                            " Where a.诊室id = b.id And a.记录id = d.出诊记录id And d.No = [1] " & _
                            IIf(blnBusy, " ", " And b.缺省标志=0 ") & " and d.记录性质=[2] and d.记录状态=1"
                    End If
                End If
            End If
        End If
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO, lng记录性质, lng转诊科室ID)
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        For i = 1 To rsTmp.RecordCount
            Set objListItem = frmDistRoom.lvwMain.ListItems.Add(, , zlCommFun.Nvl(rsTmp!名称))
            objListItem.SubItems(1) = zlCommFun.Nvl(rsTmp!位置)
            
            If tbPage.Item(midx.idx_排队队列).Selected Then
                strTmp = Me.lvwMain.SelectedItem.SubItems(EnmCol.Enm诊室)
            Else
                strTmp = Me.LvwYY.SelectedItem.SubItems(EnmCol.Enm诊室)
            End If
            If rsTmp!名称 = strTmp Then
                objListItem.Selected = True
                objListItem.EnsureVisible
                lngSel = i
            End If
            rsTmp.MoveNext
        Next
        If lngSel = 0 Then
            frmDistRoom.lvwMain.ListItems(1).Selected = True
            frmDistRoom.lvwMain.ListItems(1).EnsureVisible
        End If
    End If

    Exit Sub
ErrHead:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadPatientInfo(ByVal lng病人ID As Long, Optional ByRef bln入院 As Boolean = False)
    On Error GoTo errH
    '功能:读取病人信息
    Dim str过敏 As String, strSQL As String
    Dim i As Integer
    Dim lngTmp As Long
    Dim rsTmp As ADODB.Recordset
    Dim strNO As String
    Dim strName As String
    Dim strSex  As String, strAge As String
    Dim lng记录性质 As String
    If tbPage.Item(midx.idx_排队队列).Selected Then
        strNO = lvwMain.SelectedItem.Text
        strName = lvwMain.SelectedItem.SubItems(EnmCol.Enm姓名)
        strSex = lvwMain.SelectedItem.SubItems(EnmCol.Enm性别)
        strAge = lvwMain.SelectedItem.SubItems(EnmCol.Enm年龄)
    Else
        strNO = LvwYY.SelectedItem.Text
        strName = LvwYY.SelectedItem.SubItems(EnmCol.Enm姓名)
        strSex = LvwYY.SelectedItem.SubItems(EnmCol.Enm性别)
        strAge = LvwYY.SelectedItem.SubItems(EnmCol.Enm年龄)
    End If
    
    lng记录性质 = IIf(tbPage.Item(midx.idx_排队队列).Selected, 1, 2)
    
    With frmDistRoomPatiEdit
        .txtPatient.MaxLength = GetColumnLength("病人信息", "姓名")
        .txt年龄.MaxLength = GetColumnLength("病人信息", "年龄")
        .txt门诊号.MaxLength = GetColumnLength("病人信息", "门诊号")
        .padd家庭地址.MaxLength = GetColumnLength("病人信息", "家庭地址")
        .padd户口地址.MaxLength = GetColumnLength("病人信息", "户口地址")
        .mstr姓名 = "": .mstr性别 = "": .mstr年龄 = "": .mstr出生日期 = ""
        .mbln医嘱业务 = False
        If lng病人ID <= 0 Then
            .mbytType = 1  '创建一个新的病人信息
            .txt门诊号.Text = zlDatabase.GetNextNo(3)
            .txtPatient.Text = strName
            For i = 0 To .cbo性别.ListCount - 1
                If .cbo性别.List(i) Like "*" & Trim(strSex) Then
                    .cbo性别.ListIndex = i
                    Exit For
                End If
            Next
            Call LoadOldData(strAge, .txt年龄, .cbo年龄单位)
            Exit Sub
        End If
        '79912:李南春,2014/11/20,入院病人不能在分诊台修改信息
        strSQL = _
            "Select A.*,D.姓名 as 挂号姓名,D.性别 as 挂号性别,D.年龄 as 挂号年龄,Decode(B.病人ID,NULL,0,1) As 病案,C.医疗类别,To_Char(C.就诊时间,'YYYY-MM-DD HH24:MI:SS') As  就诊时间,D.ID As 挂号ID " & vbCrLf & _
            " From 病人信息 A,门诊病案记录 B,就诊登记记录 C,病人挂号记录 D" & vbCrLf & _
            " Where A.病人ID=B.病人ID(+) And A.病人ID=[1] " & vbCrLf & _
            "　And D.NO=[2] And D.登记时间=C.就诊时间(+) And D.病人ID=C.病人ID(+) and d.记录性质=[3] and d.记录状态=1 "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, strNO, lng记录性质)
        
        If rsTmp.RecordCount > 0 Then
            If Val(Nvl(rsTmp!当前科室id)) <> 0 Then bln入院 = True: Exit Sub
            .ClearFace
            '对于有信息的病人,也缺省为不修改信息,不建病案
            '对于有病案的病人,只能为修改病案信息
            If rsTmp!病案 = 1 Then
                .mbytType = 3 '只更新病人信息
            Else
                .mbytType = 2 '无病案,建立一个新的病案信息
            End If
            .mstr出生日期 = Format(rsTmp!出生日期, "YYYY-MM-DD")
            If lng病人ID = 0 Then
                .mbln医嘱业务 = False
            Else
                 .mbln医嘱业务 = zlExistOperationData(lng病人ID, strNO, Val(Nvl(rsTmp!挂号ID)))
            End If
            .mstr出生日期 = Format(rsTmp!出生日期, "YYYY-MM-DD")
            .mstr病人_姓名 = Nvl(rsTmp!姓名)
            .mstr病人_性别 = Nvl(rsTmp!性别)
            .mstr病人_年龄 = Nvl(rsTmp!年龄)
            .mstr姓名 = Nvl(rsTmp!挂号姓名)
            .mstr性别 = Nvl(rsTmp!挂号性别)
            .mstr年龄 = Nvl(rsTmp!挂号年龄)
            
            If IsNull(rsTmp!门诊号) Then
                .txt门诊号.Text = zlDatabase.GetNextNo(3)     '有病案者一定有门诊号
            Else
                .txt门诊号.Text = rsTmp!门诊号
            End If
            If .mbln医嘱业务 And InStr(.mstrPrivsPubPatient, ";基本信息调整;") = 0 Then
                .txtPatient.Text = zlCommFun.Nvl(rsTmp!挂号姓名)
                .cbo性别.ListIndex = cbo.FindIndex(.cbo性别, zlCommFun.Nvl(rsTmp!挂号性别), True)
                Call LoadOldData("" & rsTmp!挂号年龄, .txt年龄, .cbo年龄单位)
            Else
                .txtPatient.Text = zlCommFun.Nvl(rsTmp!姓名)
                .cbo性别.ListIndex = cbo.FindIndex(.cbo性别, zlCommFun.Nvl(rsTmp!性别), True)
                Call LoadOldData("" & rsTmp!年龄, .txt年龄, .cbo年龄单位)
            End If
            '74428：李南春，2014-7-8，病人姓名显示颜色处理
            Call SetPatiColor(.txtPatient, Nvl(rsTmp!病人类型), IIf(IsNull(rsTmp!险类), .ForeColor, vbRed))
            .mblnChange = False
            .txt出生日期.Text = Format(IIf(IsNull(rsTmp!出生日期), "____-__-__", rsTmp!出生日期), "YYYY-MM-DD")
            .mblnChange = True

            If Not IsNull(rsTmp!出生日期) Then
                If .mbln医嘱业务 = False Then
                    .txt年龄.Text = ReCalcOld(CDate(.txt出生日期.Text), .cbo年龄单位, lng病人ID) '根据出生日期重算年龄
                    If CDate(.txt出生日期.Text) - CDate(rsTmp!出生日期) <> 0 Then .txt出生时间.Text = Format(rsTmp!出生日期, "HH:MM")
                End If
            Else
                .txt出生时间.Text = "__:__"
                .mblnChange = False
                  If .mbln医嘱业务 = False Then
                    .txt出生日期.Text = ReCalcBirth(.txt年龄.Text, .cbo年龄单位.Text)
                End If
                .mblnChange = True
            End If


            If .mlngOutModeMC > 0 Then
                If .mlngOutModeMC = 920 Then
                    .txtPatiMCNO(0).MaxLength = 12
                Else
                    .txtPatiMCNO(0).MaxLength = 30
                End If
                .txtPatiMCNO(0).ToolTipText = "最大长度" & .txtPatiMCNO(0).MaxLength & "位"
                .txtPatiMCNO(1).MaxLength = .txtPatiMCNO(0).MaxLength

                .txtPatiMCNO(0).Text = "" & rsTmp!医保号    '自动截断超过最大长度的字符
                .txtPatiMCNO(0).Tag = .txtPatiMCNO(0).Text
                .txtPatiMCNO(1).Text = .txtPatiMCNO(0).Text

                If Not IsNull(rsTmp!医疗类别) Then
                    For i = 0 To .cbo医疗类别.ListCount - 1
                        lngTmp = InStr(1, .cbo医疗类别.List(i), "-")
                        If lngTmp > 1 Then
                            If Mid(.cbo医疗类别.List(i), 1, lngTmp - 1) = rsTmp!医疗类别 Then
                                .cbo医疗类别.ListIndex = i: Exit For
                            End If
                        End If
                    Next
                    .cbo医疗类别.Tag = "" & rsTmp!就诊时间
                End If
            ElseIf .mlng险类 > 0 Then
                 
                .mstr医保号 = "" & rsTmp!医保号
                 
            End If
            

            .cbo费别.ListIndex = cbo.FindIndex(.cbo费别, zlCommFun.Nvl(rsTmp!费别), True)
            .cbo付款方式.ListIndex = cbo.FindIndex(.cbo付款方式, zlCommFun.Nvl(rsTmp!医疗付款方式), True)
            .cbo国籍.ListIndex = cbo.FindIndex(.cbo国籍, zlCommFun.Nvl(rsTmp!国籍), True)
            .cbo民族.ListIndex = cbo.FindIndex(.cbo民族, zlCommFun.Nvl(rsTmp!民族), True)
            .cbo婚姻.ListIndex = cbo.FindIndex(.cbo婚姻, zlCommFun.Nvl(rsTmp!婚姻状况), True)
            .cbo职业.ListIndex = cbo.FindIndex(.cbo职业, zlCommFun.Nvl(rsTmp!职业), True)
            .txt身份证号.Text = zlCommFun.Nvl(rsTmp!身份证号)
            .txt单位名称.Text = zlCommFun.Nvl(rsTmp!工作单位)
            .txt单位名称.Tag = zlCommFun.Nvl(rsTmp!合同单位ID)
            .txt单位电话.Text = zlCommFun.Nvl(rsTmp!单位电话)
            .txt单位邮编.Text = zlCommFun.Nvl(rsTmp!单位邮编)
            .cbo家庭地址.Text = zlCommFun.Nvl(rsTmp!家庭地址)
            '89242:李南春,2015/12/7,使用结构化地址
            Call zlReadAddrInfo(.padd家庭地址, Val(Nvl(rsTmp!病人ID)), Val(Nvl(rsTmp!主页ID)), 3, .cbo家庭地址.Text)
            If .padd家庭地址.Value = "" Then Call zlLoadDefaultAddr(.padd家庭地址)
            .txt家庭电话.Text = zlCommFun.Nvl(rsTmp!家庭电话)
            .txt家庭邮编.Text = zlCommFun.Nvl(rsTmp!家庭地址邮编)
            .txt户口地址.Text = zlCommFun.Nvl(rsTmp!户口地址)
            Call zlReadAddrInfo(.padd户口地址, Val(Nvl(rsTmp!病人ID)), Val(Nvl(rsTmp!主页ID)), 4, .txt户口地址.Text)
            If .padd户口地址.Value = "" Then Call zlLoadDefaultAddr(.padd户口地址)
            .txt户口邮编.Text = zlCommFun.Nvl(rsTmp!户口地址邮编)
            .txtEdit(0).Text = zlCommFun.Nvl(rsTmp!监护人)
            .mlng挂号ID = Val(Nvl(rsTmp!挂号ID))
            '过敏药物
            str过敏 = Get过敏药物(rsTmp!病人ID)
            If str过敏 <> "" Then
                If UBound(Split(str过敏, ";")) + 1 > .msh过敏.Rows - 1 Then .msh过敏.Rows = UBound(Split(str过敏, ";")) + 2
                For i = 0 To UBound(Split(str过敏, ";"))
                    .msh过敏.RowData(i + 1) = Val(Split(Split(str过敏, ";")(i), "|")(0))
                    .msh过敏.TextMatrix(i + 1, 0) = Split(Split(str过敏, ";")(i), "|")(1)
                    .msh过敏.TextMatrix(i + 1, 1) = Split(Split(str过敏, ";")(i), "|")(2)
                Next
            End If
        Else
            .mbytType = 1 '创建一个新的病人信息
            .txt门诊号.Text = zlDatabase.GetNextNo(3)

            .txtPatient.Text = strName
            For i = 0 To .cbo性别.ListCount - 1
                If .cbo性别.List(i) Like "*" & Trim(strSex) Then
                    .cbo性别.ListIndex = i
                    Exit For
                End If
            Next
            Call LoadOldData(strAge, .txt年龄, .cbo年龄单位)
        End If
        strSQL = "Select 就诊ID,信息名,信息值 From 病人信息从表 Where 病人ID=[1] And 就诊ID=[2] "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, .mlng病人ID, .mlng挂号ID)
        If rsTmp.EOF Then Exit Sub
        'idx_监护人 = 0
        'idx_身高 = 1
        'idx_体重 = 2
        'idx_体温 = 3
        rsTmp.Filter = "信息名='身高'"
        If rsTmp.RecordCount > 0 Then
            .txtEdit(1).Text = Nvl(rsTmp!信息值)
        End If
        rsTmp.Filter = "信息名='体重'"
        If rsTmp.RecordCount > 0 Then
            .txtEdit(2).Text = Nvl(rsTmp!信息值)
        End If
        rsTmp.Filter = "信息名='体温'"
        If rsTmp.RecordCount > 0 Then
            .txtEdit(3).Text = Nvl(rsTmp!信息值)
        End If
        
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
 
Private Sub InitVariateAndPara()
    '------------------------------------------------------------------------------------------------------------------------
    '功能：初始货变量
    '编制：刘兴洪
    '日期：2010-05-31 16:58:25
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim arrTmp As Variant, i As Long
    '将用来记录病人ID进行初始化
    mlngPre病人ID = 0: mlngOutModeMC = 0
    arrTmp = Split(GetSetting("ZLSOFT", "公共全局", "本地支持的医保", ""), ",")
    For i = 0 To UBound(arrTmp)
        If IsNumeric(arrTmp(i)) Then
            If CheckMCOutMode(arrTmp(i)) Then mlngOutModeMC = Val(arrTmp(i)): Exit For
        End If
    Next
    mlngDefaultCardID = Val(zlDatabase.GetPara("缺省医疗卡类别", glngSys, mlngModul, 0))
    
    frmDistRoomPatiEdit.mblnStructAdress = Val(zlDatabase.GetPara(251, glngSys)) <> 0 '病人地址结构化录入
    frmDistRoomPatiEdit.mblnShowTown = Val(zlDatabase.GetPara(252, glngSys)) <> 0 '乡镇地址结构化录入
End Sub
Private Sub Form_Load()
    Dim strText As String, i As Integer
    On Error GoTo errHandle
    mstrPrivs = gstrPrivs: mlngModul = glngModul
    Call InitPancel
    Call InitPage
    lvwMain.View = lvwReport
    lvwMain.ColumnHeaders.Clear
    '医保号,摘要:21101
    '74898:李南春,2015/4/9,标记病人的呼叫状态
    zlControl.LvwSelectColumns lvwMain, "挂号单,1100,0,2;病人状态,900,0,1;科室,1200,0,1;号类,600,0,1;挂号项目,1350,0,1;门诊号,800,0,1;姓名,960,0,1;性别,400,0,1;年龄,600,0,1;诊室,1600,0,1;医生,960,0,1;发生时间,2000,0,1;挂号时间,2000,0,1;序号,600,0,1;医保号,600,0,1;摘要,2000,0,1;预约,800,0,1;呼叫,400,0,1;呼叫人,960,0,1;呼叫诊室,1600,0,1;呼叫时间,2000,0,1", True
    zlControl.LvwSelectColumns LvwYY, "预约单,1100,0,2;病人状态,900,0,1;科室,1200,0,1;号类,600,0,1;挂号项目,1350,0,1;门诊号,800,0,1;姓名,960,0,1;性别,400,0,1;年龄,600,0,1;诊室,1600,0,1;医生,960,0,1;预约时间,2000,0,1;挂号时间,2000,0,1;序号,600,0,1;医保号,600,0,1;摘要,2000,0,1", True
    zlControl.LvwSelectColumns lvwHZPati, "挂号单,1100,0,2;病人状态,900,0,1;科室,1200,0,1;号类,600,0,1;挂号项目,1350,0,1;门诊号,800,0,1;姓名,960,0,1;性别,400,0,1;年龄,600,0,1;诊室,1600,0,1;医生,960,0,1;发生时间,2000,0,1;挂号时间,2000,0,1;序号,600,0,1;医保号,600,0,1;摘要,2000,0,1;预约,800,0,1", True
    Call InitVariateAndPara
'    '进入时刷新一次
'    Call zlRefreshData     问题108110,多次调用刷新分诊列表
    strText = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\zl9RegEvent\" & Me.Name & "\ListView", lvwMain.Name & "名称")
    For i = 1 To lvwMain.ColumnHeaders.Count
        '如果添加了列，则不恢复个性化
        If InStr(strText, lvwMain.ColumnHeaders(i).Text) = 0 Then lvwMain.Tag = "": Exit For
        '如果减少了列，也不恢复个性化
        strText = Replace(strText, lvwMain.ColumnHeaders(i).Text, "")
    Next
    strText = Replace(strText, ",", "")
    If strText <> "" Then lvwMain.Tag = ""
    Call RestoreWinState(Me, App.ProductName)
    lvwMain.Tag = "可变化的"
    
    If CreatePlugInOK(glngModul) Then
        gblnPlugin = True
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function CheckMCOutMode(ByVal strMCCode As String) As Boolean
    Dim rsTmp As ADODB.Recordset, strSQL As String

    strSQL = "Select 1 From 保险类别 Where 外挂=1 And 序号=[1]"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strMCCode)

    CheckMCOutMode = rsTmp.RecordCount > 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    zlSaveDockPanceToReg Me, dkpMan, "区域"
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub lvwMain_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwMain.Sorted = True
    If lvwMain.SortKey = ColumnHeader.index - 1 Then
        If lvwMain.SortOrder = lvwAscending Then
            lvwMain.SortOrder = lvwDescending
        Else
            lvwMain.SortOrder = lvwAscending
        End If
    Else
        lvwMain.SortKey = ColumnHeader.index - 1
    End If
End Sub

Private Sub lvwYY_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LvwYY.Sorted = True
    If LvwYY.SortKey = ColumnHeader.index - 1 Then
        If LvwYY.SortOrder = lvwAscending Then
            LvwYY.SortOrder = lvwDescending
        Else
            LvwYY.SortOrder = lvwAscending
        End If
    Else
        LvwYY.SortKey = ColumnHeader.index - 1
    End If
End Sub
Private Sub lvwMain_DblClick()
    If Not lvwMain.SelectedItem Is Nothing Then
    
        Call zlExcuteFunction
    End If
End Sub

Private Sub lvwYY_DblClick()
    If Not LvwYY.SelectedItem Is Nothing Then
         Call zlExcuteFunction
    End If
End Sub

Private Sub lvwMain_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
    Dim lng病人ID As Long, lngExeState As Long, i As Long, j As Long
    Dim strSQL As String, strFilter As String, dteTmp As Date
    Dim objListItem As ListItem


    '有错误退出
    Err = 0: On Error GoTo errHandle
    If IsEmpty(Item.Tag) Then Exit Sub
    If TypeName(Item.Tag) <> "String" Then Exit Sub
    If InStr(1, Item.Tag, "|") < 1 Then Exit Sub

    lvwMain.Tag = Item.Text

    '根据是否已经建立病案(存在病人id)、执行状态，决定是否可分诊、换号、建立病案、完成接诊等系列操作
    lng病人ID = Val(Split(Item.Tag, "|")(0))
    lngExeState = Val(Split(Item.Tag, "|")(6))
    mlngPre病人ID = lng病人ID

    RaiseEvent zlShowInfor("单据号:" & Item.Text & _
        "  病人:" & Item.SubItems(EnmCol.Enm姓名) & _
        "  诊室:" & IIf(Item.SubItems(EnmCol.Enm诊室) = "", "未分诊", Item.SubItems(EnmCol.Enm诊室)) & _
        "  医生:" & IIf(Item.SubItems(EnmCol.Enm医生) = "", "未指定", Item.SubItems(EnmCol.Enm医生)))
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub lvwYY_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
    Dim lng病人ID As Long, lngExeState As Long, i As Long, j As Long
    Dim strSQL As String, strFilter As String, dteTmp As Date
    Dim objListItem As ListItem
    '有错误退出
    Err = 0: On Error GoTo errHandle
    If IsEmpty(Item.Tag) Then Exit Sub
    If TypeName(Item.Tag) <> "String" Then Exit Sub
    If InStr(1, Item.Tag, "|") < 1 Then Exit Sub

    LvwYY.Tag = Item.Text

    '根据是否已经建立病案(存在病人id)、执行状态，决定是否可分诊、换号、建立病案、完成接诊等系列操作
    lng病人ID = Val(Split(Item.Tag, "|")(0))
    lngExeState = Val(Split(Item.Tag, "|")(6))
    mlngPre病人ID = lng病人ID

    RaiseEvent zlShowInfor("预约单据:" & Item.Text & _
        "  病人:" & Item.SubItems(EnmCol.Enm姓名) & _
        "  诊室:" & IIf(Item.SubItems(EnmCol.Enm诊室) = "", "未分诊", Item.SubItems(EnmCol.Enm诊室)) & _
        "  医生:" & IIf(Item.SubItems(EnmCol.Enm医生) = "", "未指定", Item.SubItems(EnmCol.Enm医生)))
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim bln分诊台签到排队 As Boolean
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    bln分诊台签到排队 = Val(lvwMain.SelectedItem.ListSubItems(4).Tag) = 1
    If Button = 1 And IsNumeric(Trim(lvwMain.SelectedItem.SubItems(EnmCol.Enm门诊号))) Then
        If bln分诊台签到排队 Then
            '    'ListSubItem(3).tag:A.记录标志:表示0表示就诊病人,1-表示签到的病人,2-表示需要回诊就诊的病人; 3-表示已回诊但还未接收的病人;
            If Val(lvwMain.SelectedItem.ListSubItems(3).Tag) = 0 And Val(Split(lvwMain.SelectedItem.Tag, "|")(6)) = 0 Then
                Exit Sub
            End If
        End If
        Set lvwMain.DragIcon = lvwMain.SelectedItem.CreateDragImage
        lvwMain.Drag 1
    End If
End Sub
Private Sub lvwMain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And Shift = 0 Then
        Call ReadRoom
        RaiseEvent zlPopuMenu(Button, Shift, X, Y)
    End If
End Sub

Private Sub lvwRoom_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwRoom.Sorted = True
    If lvwRoom.SortKey = ColumnHeader.index - 1 Then
        If lvwRoom.SortOrder = lvwAscending Then
            lvwRoom.SortOrder = lvwDescending
        Else
            lvwRoom.SortOrder = lvwAscending
        End If
    Else
        lvwRoom.SortKey = ColumnHeader.index - 1
    End If
End Sub

Private Sub lvwRoom_DragDrop(Source As Control, X As Single, Y As Single)
    On Error GoTo errH
    Dim strRoom As String, strDate As String
    Dim strDoctor As String, strSQL As String
    Dim i As Long
    Dim blnCancel As Boolean
    If Source Is lvwMain And Not lvwRoom.DropHighlight Is Nothing Then
        Set lvwRoom.SelectedItem = lvwRoom.DropHighlight
        Set lvwRoom.DropHighlight = Nothing
        ReadRoom
        strRoom = STR_COMP
        If frmDistRoom.cmb.ListCount > 0 Then
            frmDistRoom.cmb.ListIndex = 0
            For i = 0 To frmDistRoom.cmb.ListCount - 1
                If frmDistRoom.cmb.List(i) Like "*-" & lvwMain.SelectedItem.SubItems(EnmCol.Enm医生) Then
                    frmDistRoom.cmb.ListIndex = i
                    Exit For
                End If
            Next
        End If
        If frmDistRoom.lvwMain.ListItems.Count > 0 Then
            For i = 1 To frmDistRoom.lvwMain.ListItems.Count
                If frmDistRoom.lvwMain.ListItems(i).Text = lvwRoom.SelectedItem.Text Then
                    frmDistRoom.lvwMain.ListItems(i).Selected = True
                    frmDistRoom.lvwMain.ListItems(i).EnsureVisible
                    Exit For
                End If
            Next
        End If
        frmDistRoom.ShowMe strRoom, Me
        If strRoom = STR_COMP Then
            RaiseEvent zlShowInfor("用户取消！"): Exit Sub   '选择了"取消"了
        End If
        'NO_IN       病人挂号记录.NO%TYPE:=NULL,
        '病人ID_IN   病人挂号记录.病人id%TYPE:=NULL,
        '诊室_IN     病人挂号记录.诊室%TYPE:=NULL
        '
        strDoctor = Trim(Split(strRoom, STR_COMP)(1))
        strDate = "To_date('" & Split(strRoom, STR_COMP)(2) & "','yyyy-mm-dd hh24:mi:ss')"
        strRoom = Split(strRoom, STR_COMP)(0)
        '问题号:48350
        strSQL = "ZL_病人挂号记录_更新诊室 ('" & lvwMain.SelectedItem.Text & "'," & mlngPre病人ID & ",'" & strRoom & "','" & strDoctor & "'," & strDate & ",'','" & zl_Get预约方式ByNo(lvwMain.SelectedItem.Text) & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        RaiseEvent zlQueueAsk(2, Trim(lvwMain.SelectedItem.Text), mlngPre病人ID, blnCancel)
        'intType:intType:1-分诊;2-换号;3-病人不就诊;4-病人待诊;5-病人完成就诊;6-恢复就诊
        ' 0-复诊,1-直呼,2-弃号,3-暂停,4-完成就诊,5-广播
        
        '77412:李南春，2014/9/3,门诊病人条码打印
        Call zlPrintBarcode

        Call zlRefreshData
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwRoom_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    On Error GoTo errH
    Dim objOver As ListItem

    '没选择就退出
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    If Source Is lvwMain Then
        Set objOver = lvwRoom.HitTest(X, Y)
        If Not objOver Is Nothing Then
            If objOver.ForeColor <> RGB(255, 0, 0) And Trim(lvwMain.SelectedItem.SubItems(EnmCol.Enm门诊号)) <> "" And objOver.SubItems(5) Like "*" & lvwMain.SelectedItem.SubItems(EnmCol.Enm科室) & "*" Then
                Set lvwRoom.DropHighlight = objOver
            Else
                Set lvwRoom.DropHighlight = Nothing
            End If
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Sub zlRefreshData(Optional blnFilter As Boolean = False, _
    Optional strFindValue As String = "", Optional bytReadType As Byte = 0, Optional objCard As Card, Optional ByVal blnAuto签到 As Boolean)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：重新刷新数据
    '入参：blnFilter-是否过滤
    '          bytReadType-读取类型(0-不区分;1-刷卡;2-读取身份证;3-读取IC卡)
    '编制：刘兴洪
    '日期：2010-06-02 09:43:08
    '------------------------------------------------------------------------------------------------------------------------
    Call ShowBills(blnFilter, strFindValue, bytReadType, objCard, blnAuto签到)
End Sub

Private Sub ShowBills(blnFilter As Boolean, Optional strFindValue As String = "", _
    Optional bytReadType As Byte = 0, Optional objCard As Card, Optional ByVal blnAuto签到 As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取分诊数据
    '入参：blnFilter-是否过滤
    '          bytReadType-读取类型(0-不区分;1-刷卡;2-读取身份证;3-读取IC卡)
    '     blbAppointment-是否刷新预约队列
    '编制:刘兴洪
    '日期:2011-11-21 10:50:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strValue As String, bytType As Byte
    mstrRegistIdsed = ""
    
    On Error GoTo errHandle
    Screen.MousePointer = 11
    
    If GetFilterCons(strFindValue, objCard, bytReadType, strValue, bytType) = False Then Screen.MousePointer = 0: Exit Sub
        
    '加载预约数据
    Call ShowBillsAppointment(blnFilter, strValue, bytType, objCard, blnAuto签到)
    
    '加载挂号数据
    Call ShowBillRegister(blnFilter, strValue, bytType, objCard, blnAuto签到)
    
    '加载回诊数据
    Call ShowBillRegisterHZ(blnFilter, strValue, bytType, objCard)
    
    '加载诊室
    Call LoadRooms
    
    Screen.MousePointer = 0
    Exit Sub
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Public Sub zlSetViewScrop(ByVal index As Integer, ByVal bytValue As Byte, Optional blnRefrashData As Boolean = False)
    '设置相关的显示病人
    'mbytViewScrop(0 To 3) As Byte  '0-显示已分诊病人;1-显示已接诊病人;2-显示已完成病人;3-显示不就诊病人
    mbytViewScrop(index) = bytValue
     
    If blnRefrashData Then Call zlRefreshData
End Sub

'设置过滤条件
Public Sub zlSetFilterCons(ByVal ArrFilter As Variant)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设条件
    '编制:刘兴洪
    '日期:2009-09-15 11:19:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mcllFilter = ArrFilter
End Sub

Public Sub zlSetobjMsgModule(ByVal objMsgModule As Object)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设条件
    '编制:刘兴洪
    '日期:2009-09-15 11:19:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mobjMsgModule = objMsgModule
End Sub
 

'------------------------------------------------------------------------------------------------------
'相关属性设置:
'    分诊科室;有效天数
Public Property Get zl分诊科室() As String
    zl分诊科室 = mstr分诊科室
End Property

Public Property Let zl分诊科室(ByVal vNewValue As String)
    mstr分诊科室 = vNewValue
End Property

Public Property Get zl有效天数() As Integer
    zl有效天数 = mint有效天数
End Property

Public Property Let zl有效天数(ByVal vNewValue As Integer)
    mint有效天数 = vNewValue
End Property

Public Property Get zlintFindKeys() As Integer
    zlintFindKeys = mintFindKeys
End Property
Public Property Let zlintFindKeys(ByVal vNewValue As Integer)
    mintFindKeys = vNewValue
End Property

Public Property Get zlIsHaveData() As Boolean
    If Me.ActiveControl Is Me.lvwHZPati Then
        zlIsHaveData = lvwHZPati.ListItems.Count <> 0
    ElseIf tbPage.Item(midx.idx_排队队列).Selected Then
        zlIsHaveData = lvwMain.ListItems.Count <> 0
    Else
        zlIsHaveData = LvwYY.ListItems.Count <> 0
    End If
End Property
Public Property Get zlIsTriage() As Boolean
    Dim lng病人ID As Long, lngExeState As Long
    
    If tbPage.Item(midx.idx_排队队列).Selected Then
        '是否能分诊
        If lvwMain.SelectedItem Is Nothing Or Me.ActiveControl Is Me.lvwHZPati Then
            zlIsTriage = False
        Else
            lng病人ID = Val(Split(lvwMain.SelectedItem.Tag, "|")(0))
            lngExeState = Val(Split(lvwMain.SelectedItem.Tag, "|")(6))
            'ListSubItem(3).tag:A.记录标志:表示0表示就诊病人,1-表示签到的病人,2-表示需要回诊就诊的病人; 3-表示已回诊但还未接收的病人;
            'ListSubItems(4).Tag:0-不启用分诊台签到排队;1-启用分诊台签到排队
            If Val(lvwMain.SelectedItem.ListSubItems(3).Tag) = 1 Then
                '已经签道
                 zlIsTriage = (lngExeState = 0)
            Else '未签到
                 zlIsTriage = (lngExeState = 0) And Not Val(lvwMain.SelectedItem.ListSubItems(4).Tag) = 1
            End If
        End If
    Else
        '是否能分诊
        If LvwYY.SelectedItem Is Nothing Or Me.ActiveControl Is Me.lvwHZPati Then
            zlIsTriage = False
        Else
            lng病人ID = Val(Split(LvwYY.SelectedItem.Tag, "|")(0))
            lngExeState = Val(Split(LvwYY.SelectedItem.Tag, "|")(6))
            'ListSubItem(3).tag:A.记录标志:表示0表示就诊病人,1-表示签到的病人,2-表示需要回诊就诊的病人; 3-表示已回诊但还未接收的病人;
            'ListSubItems(4).Tag:0-不启用分诊台签到排队;1-启用分诊台签到排队
            If Val(LvwYY.SelectedItem.ListSubItems(3).Tag) = 1 Then
                '已经签道
                 zlIsTriage = (lngExeState = 0)
            Else '未签到
                 zlIsTriage = (lngExeState = 0) And Not Val(LvwYY.SelectedItem.ListSubItems(4).Tag) = 1
            End If
        End If
    End If
End Property
Public Property Get zlIsPatiLeave() As Boolean
    '病人允许不就诊
    Dim lng病人ID As Long, lngExeState As Long
    If lvwMain.SelectedItem Is Nothing Or Me.ActiveControl Is Me.lvwHZPati Then
        lngExeState = 0
    Else
        lngExeState = Val(Split(lvwMain.SelectedItem.Tag, "|")(6))
    End If
    zlIsPatiLeave = (lngExeState = 0)
End Property

Public Property Get zlIsPatiWait() As Boolean
    '病人是否允许待诊
    Dim lng病人ID As Long, lngExeState As Long
    If lvwMain.SelectedItem Is Nothing Or Me.ActiveControl Is Me.lvwHZPati Then
        lngExeState = 0
    Else
        lngExeState = Val(Split(lvwMain.SelectedItem.Tag, "|")(6))
    End If
     zlIsPatiWait = (lngExeState = -1)
End Property
Public Property Get zlIsPatiFinish() As Boolean
    '病人是否允许完成就诊
    Dim lng病人ID As Long, lngExeState As Long
    If lvwMain.SelectedItem Is Nothing Or Me.ActiveControl Is Me.lvwHZPati Or tbPage.Item(midx.idx_预约队列).Selected Then
        zlIsPatiFinish = False
    Else
        lng病人ID = Val(Split(lvwMain.SelectedItem.Tag, "|")(0))
        lngExeState = Val(Split(lvwMain.SelectedItem.Tag, "|")(6))
         zlIsPatiFinish = (lngExeState = 0 Or lngExeState = 2)
    End If
End Property

Public Property Get zlIsPatiReDo() As Boolean
    '是否允许病人恢复就诊
    Dim lng病人ID As Long, lngExeState As Long
    If lvwMain.SelectedItem Is Nothing Or Me.ActiveControl Is Me.lvwHZPati Or tbPage.Item(midx.idx_预约队列).Selected Then
        zlIsPatiReDo = False
    Else
        lng病人ID = Val(Split(lvwMain.SelectedItem.Tag, "|")(0))
        lngExeState = Val(Split(lvwMain.SelectedItem.Tag, "|")(6))
         zlIsPatiReDo = (lngExeState = 1)
    End If
End Property

Public Property Get zlIs允许签道(Optional ByRef bytQueue As Byte) As Boolean
    '是否允许病人签道
    'mbln分诊台签到排队: 0-挂号立即排队,1-分诊台签到排队
    'bytQueue0-正常签到，1-重新签到
    Dim lng病人ID As Long, lngExeState As Long, lngTrunState As Long
    If Me.ActiveControl Is Me.lvwHZPati Then
        '回诊病人
        zlIs允许签道 = False
    '63789,刘尔旋,2014-01-09,允许预约病人签到
    ElseIf tbPage.Item(midx.idx_排队队列).Selected Then
        If lvwMain.SelectedItem Is Nothing Then
            zlIs允许签道 = False
        Else
            lng病人ID = Val(Split(lvwMain.SelectedItem.Tag, "|")(0))
            lngExeState = Val(Split(lvwMain.SelectedItem.Tag, "|")(6))
            lngTrunState = Val(Split(lvwMain.SelectedItem.Tag, "|")(8))
            'ListSubItem(3).tag:A.记录标志:表示0表示就诊病人,1-表示签到的病人,2-表示需要回诊就诊的病人; 3-表示已回诊但还未接收的病人;
            
            '95637:李南春,2016/7/17,支持分诊台签到排队模式的换号，转诊签到以及重新签到
            zlIs允许签道 = (lngExeState = 0 Or lngExeState = 2 And lngTrunState = 0) And (Val(lvwMain.SelectedItem.ListSubItems(3).Tag) = 0 Or Val(lvwMain.SelectedItem.ListSubItems(3).Tag) = 1)
            bytQueue = Val(lvwMain.SelectedItem.ListSubItems(3).Tag)
            '!病人ID & "|" & !险类 & "|" & !就诊卡号 & "|" & !卡验证码 & "|" & !ID & "|" & !号别 & "|" & !执行状态 & "|" & !签到类型 & "|" & !转诊状态
        End If
    ElseIf tbPage.Item(midx.idx_预约队列).Selected Then
        If LvwYY.SelectedItem Is Nothing Then
            zlIs允许签道 = False
        Else
            lng病人ID = Val(Split(LvwYY.SelectedItem.Tag, "|")(0))
            lngExeState = Val(Split(LvwYY.SelectedItem.Tag, "|")(6))
            'ListSubItem(3).tag:A.记录标志:表示0表示就诊病人,1-表示签到的病人,2-表示需要回诊就诊的病人; 3-表示已回诊但还未接收的病人;
            'ListSubItems(4).Tag:0-不启用分诊台签到排队;1-启用分诊台签到排队
            zlIs允许签道 = (lngExeState = 0) And Val(LvwYY.SelectedItem.ListSubItems(3).Tag) = 0 And Val(LvwYY.SelectedItem.ListSubItems(4).Tag) = 1
            '!病人ID & "|" & !险类 & "|" & !就诊卡号 & "|" & !卡验证码 & "|" & !ID & "|" & !号别 & "|" & !执行状态
        End If
    End If
End Property

Public Property Get zlIs允许取消签道() As Boolean
    '是否允许病人签到
    Dim lng病人ID As Long, lngExeState As Long, lngTrunState As Long
    If tbPage.Item(midx.idx_排队队列).Selected Then
        If lvwMain.SelectedItem Is Nothing Or Me.ActiveControl Is Me.lvwHZPati Then
            zlIs允许取消签道 = False
        Else
            lng病人ID = Val(Split(lvwMain.SelectedItem.Tag, "|")(0))
            lngExeState = Val(Split(lvwMain.SelectedItem.Tag, "|")(6))
            lngTrunState = Val(Split(lvwMain.SelectedItem.Tag, "|")(8))
            'ListSubItem(3).tag:A.记录标志:表示0表示就诊病人,1-表示签到的病人,2-表示需要回诊就诊的病人; 3-表示已回诊但还未接收的病人;
            '未接诊或转诊待接收
            zlIs允许取消签道 = (lngExeState = 0 Or lngExeState = 2 And lngTrunState = 0) And Val(lvwMain.SelectedItem.ListSubItems(3).Tag) = 1
            '!病人ID & "|" & !险类 & "|" & !就诊卡号 & "|" & !卡验证码 & "|" & !ID & "|" & !号别 & "|" & !执行状态
        End If
    ElseIf tbPage.Item(midx.idx_预约队列).Selected Then
        If LvwYY.SelectedItem Is Nothing Or Me.ActiveControl Is Me.lvwHZPati Then
            zlIs允许取消签道 = False
        Else
            lng病人ID = Val(Split(LvwYY.SelectedItem.Tag, "|")(0))
            lngExeState = Val(Split(LvwYY.SelectedItem.Tag, "|")(6))
            'ListSubItem(3).tag:A.记录标志:表示0表示就诊病人,1-表示签到的病人,2-表示需要回诊就诊的病人; 3-表示已回诊但还未接收的病人;
            'ListSubItems(4).Tag:0-不启用分诊台签到排队;1-启用分诊台签到排队
            zlIs允许取消签道 = (lngExeState = 0) And Val(LvwYY.SelectedItem.ListSubItems(3).Tag) = 1 And Val(LvwYY.SelectedItem.ListSubItems(4).Tag) = 1
            '!病人ID & "|" & !险类 & "|" & !就诊卡号 & "|" & !卡验证码 & "|" & !ID & "|" & !号别 & "|" & !执行状态
        End If
    End If
End Property
Public Property Get zlIsRegistData() As Boolean
    If Me.ActiveControl Is Me.lvwHZPati Then
        zlIsRegistData = Not Me.lvwHZPati.SelectedItem Is Nothing
        Exit Property
    End If
    If Me.ActiveControl Is Me.LvwYY Then
        zlIsRegistData = Not Me.LvwYY.SelectedItem Is Nothing
        Exit Property
    End If
    zlIsRegistData = Not Me.lvwMain.SelectedItem Is Nothing
End Property

Public Property Get zlIs允许回诊(Optional ByRef bytQueue As Byte) As Boolean
    '是否允许回诊
    Dim lng病人ID As Long, lngExeState As Long
    If Me.ActiveControl Is Me.lvwHZPati And Not Me.lvwHZPati.SelectedItem Is Nothing Then
        '回诊病人
        'ListSubItem(3).tag:A.记录标志:表示0表示就诊病人,1-表示签到的病人,2-表示需要回诊就诊的病人; 3-表示已回诊但还未接收的病人;
        zlIs允许回诊 = Val(lvwHZPati.SelectedItem.ListSubItems(3).Tag) = 2 Or Val(lvwHZPati.SelectedItem.ListSubItems(3).Tag) = 3
        bytQueue = IIf(Val(lvwHZPati.SelectedItem.ListSubItems(3).Tag) = 3, 1, 0)
    Else
        zlIs允许回诊 = False
    End If
End Property
Public Property Get zlIs允许取消回诊() As Boolean
    '是否允许病人签道
    Dim lng病人ID As Long, lngExeState As Long
    If Not Me.ActiveControl Is Me.lvwHZPati Or lvwHZPati.SelectedItem Is Nothing Then
        zlIs允许取消回诊 = False
    Else
        lng病人ID = Val(Split(lvwHZPati.SelectedItem.Tag, "|")(0))
        lngExeState = Val(Split(lvwHZPati.SelectedItem.Tag, "|")(6))
        'ListSubItem(3).tag:A.记录标志:表示0表示就诊病人,1-表示签到的病人,2-表示需要回诊就诊的病人; 3-表示已回诊但还未接收的病人;
        zlIs允许取消回诊 = Val(lvwHZPati.SelectedItem.ListSubItems(3).Tag) = 3
        '!病人ID & "|" & !险类 & "|" & !就诊卡号 & "|" & !卡验证码 & "|" & !ID & "|" & !号别 & "|" & !执行状态
    End If
End Property

Public Property Get zlGet病人ID() As Long
    If lvwMain.SelectedItem Is Nothing Then zlGet病人ID = 0: Exit Property
    zlGet病人ID = Val(Split(lvwMain.SelectedItem.Tag, "|")(0))
End Property
 Public Property Get zlGet挂号NO() As String
    If lvwMain.SelectedItem Is Nothing Then zlGet挂号NO = "": Exit Property
    zlGet挂号NO = lvwMain.SelectedItem.Text
End Property
 Public Property Get zlGet挂号医生() As String
    If lvwMain.SelectedItem Is Nothing Then zlGet挂号医生 = "": Exit Property
    zlGet挂号医生 = lvwMain.SelectedItem.SubItems(EnmCol.Enm医生)
End Property
 Public Property Get zlGet挂号诊室() As String
    If lvwMain.SelectedItem Is Nothing Then zlGet挂号诊室 = "": Exit Property
    zlGet挂号诊室 = lvwMain.SelectedItem.SubItems(EnmCol.Enm诊室)
End Property
 Public Property Get zlGet挂号执行状态() As Integer
    If lvwMain.SelectedItem Is Nothing Then zlGet挂号执行状态 = 0: Exit Property
    zlGet挂号执行状态 = Val(Split(lvwMain.SelectedItem.Tag, "|")(6))
End Property

Public Property Get zlGet挂号ID() As Long
    If lvwMain.SelectedItem Is Nothing Then zlGet挂号ID = 0: Exit Property
    zlGet挂号ID = Val(lvwMain.SelectedItem.ListSubItems(1).Tag)
End Property
 
Private Sub lvwHZPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvwHZPati.Sorted = True
    If lvwHZPati.SortKey = ColumnHeader.index - 1 Then
        If lvwHZPati.SortOrder = lvwAscending Then
            lvwHZPati.SortOrder = lvwDescending
        Else
            lvwHZPati.SortOrder = lvwAscending
        End If
    Else
        lvwHZPati.SortKey = ColumnHeader.index - 1
    End If
End Sub
Private Sub lvwHZPati_DblClick()
    If Not lvwHZPati.SelectedItem Is Nothing Then
        Call zlExcuteFunction
    End If
End Sub

Private Sub lvwHZPati_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim rsTmp As New ADODB.Recordset, rsTmp1 As New ADODB.Recordset
    Dim lng病人ID As Long, lngExeState As Long, i As Long, j As Long
    Dim strSQL As String, strFilter As String, dteTmp As Date
    Dim objListItem As ListItem


    '有错误退出
    Err = 0: On Error GoTo errHandle
    If IsEmpty(Item.Tag) Then Exit Sub
    If TypeName(Item.Tag) <> "String" Then Exit Sub
    If InStr(1, Item.Tag, "|") < 1 Then Exit Sub

    lvwHZPati.Tag = Item.Text

    '根据是否已经建立病案(存在病人id)、执行状态，决定是否可分诊、换号、建立病案、完成接诊等系列操作
    lng病人ID = Val(Split(Item.Tag, "|")(0))
    lngExeState = Val(Split(Item.Tag, "|")(6))
    mlngPre病人ID = lng病人ID

    RaiseEvent zlShowInfor("单据号:" & Item.Text & _
        "  病人:" & Item.SubItems(EnmCol.Enm姓名) & _
        "  诊室:" & IIf(Item.SubItems(EnmCol.Enm诊室) = "", "未分诊", Item.SubItems(EnmCol.Enm诊室)) & _
        "  医生:" & IIf(Item.SubItems(EnmCol.Enm医生) = "", "未指定", Item.SubItems(EnmCol.Enm医生)))
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub lvwHZPati_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And Shift = 0 Then
        Call ReadRoom
        RaiseEvent zlPopuMenu(Button, Shift, X, Y)
    End If
End Sub
Public Sub zlExc回诊(ByVal bln取消回诊 As Boolean, Optional ByVal blnClick As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行签道
    '编制:刘兴洪
    '日期:2010-12-08 10:56:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnTriage As Boolean, lng病人ID  As Long, lngExeState As Long
    Dim lngID As Long, strTittle As String, strSQL As String
    Dim strNO As String, bln分诊台签到排队 As Boolean
    
    If Not Me.ActiveControl Is lvwHZPati Then
        Exit Sub
    End If
    If lvwHZPati.SelectedItem Is Nothing Then Exit Sub
    bln分诊台签到排队 = lvwHZPati.SelectedItem.ListSubItems(4).Tag = 1
    lng病人ID = Val(Split(lvwHZPati.SelectedItem.Tag, "|")(0))
    lngExeState = Val(Split(lvwHZPati.SelectedItem.Tag, "|")(6))
    blnTriage = (lngExeState = 0)
    lngID = Val(lvwHZPati.SelectedItem.ListSubItems(1).Tag)
    strNO = Trim(lvwHZPati.SelectedItem.Text)
    Err = 0: On Error GoTo Errhand:
    If lngID = 0 Then Exit Sub
    If ExcPlugInFun(IIf(bln取消回诊, 15, 5), lngID) = False Then Exit Sub
    
    If Not bln取消回诊 Then '回诊签道
        '95637:李南春,2016/7/18,签到需检查当前号别是否在排队中，或者当天有其他号别处于排队中
        If Check签到(False, lng病人ID, lngID, , blnClick, bln分诊台签到排队) = False Then Exit Sub
        If frmDistRoomHz.ShowMe(mfrmMain, mlngModul, mstrPrivs, strNO) = False Then Exit Sub
        strTittle = IIf(bln取消回诊, "取消回诊成功!", "病人回诊成功!")
        ShowMsgbox strTittle
        RaiseEvent zlShowInfor(strTittle)
        Call ShowBills(False, "")
        Exit Sub
    End If
    'Zl_病人挂号记录_取消回诊
    strSQL = "Zl_病人挂号记录_取消回诊("
    '  Id_In     病人挂号记录.ID%Type,
    strSQL = strSQL & "" & lngID & ","
    '  需回诊_In Integer:=0
    strSQL = strSQL & "0)"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    strTittle = IIf(bln取消回诊, "取消回诊成功!", "病人回诊成功!")
    ShowMsgbox strTittle
    RaiseEvent zlShowInfor(strTittle)
    Call ShowBills(False, "")
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub zlPrintBill(ByVal lngID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印指定的排队单
    '编制:刘兴洪
    '日期:2011-05-24 15:57:41
    '问题:38165
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnPrint As Boolean
     blnPrint = True
     If InStr(1, mstrPrivs, ";分诊排队单;") = 0 Then Exit Sub
     
     Select Case Val(zlDatabase.GetPara("排队单打印", glngSys, mlngModul))
     Case 0
         blnPrint = False
     Case 2
         If MsgBox("你是否要打印排队单吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then blnPrint = False
     End Select
     If blnPrint Then Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1113", Me, "挂号ID=" & lngID, 2)
End Sub
Public Sub zlRePrintBill()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重打排队单
    '编制:刘兴洪
    '日期:2011-05-24 16:36:20
    '问题:38165
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, strNO As String, strSQL As String, rsTemp As ADODB.Recordset
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    lngID = Val(lvwMain.SelectedItem.ListSubItems(1).Tag)
    strNO = Trim(lvwMain.SelectedItem.Text)
    If lngID = 0 Then Exit Sub
    strSQL = "Select  1 From 排队叫号队列 Where 业务类型=0 and 业务ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
    If Not rsTemp.EOF Then
        Call zlPrintBill(lngID)
    Else
        MsgBox "该病人未生成排队队列，不能打印排队单!", vbInformation + vbOKOnly, gstrSysName
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function InitPage() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载页
    '编制:李光福
    '日期:2013-05-02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem
    Dim strTemp As String
    On Error GoTo errHandle
    tbPage.RemoveAll
    
    Set ObjItem = tbPage.InsertItem(midx.idx_排队队列, "挂号病人", Me.lvwMain.Hwnd, 0)
    ObjItem.Tag = midx.idx_排队队列
    Set ObjItem = tbPage.InsertItem(midx.idx_预约队列, "预约病人", Me.LvwYY.Hwnd, 0)
    ObjItem.Tag = midx.idx_预约队列
     With tbPage
         
        tbPage.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = False
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionBottom
    End With
    InitPage = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function

Private Sub ShowBillsAppointment(blnFilter As Boolean, Optional strValue As String = "", _
    Optional bytType As Byte = 0, Optional objCard As Card, Optional ByVal blnAuto签到 As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取分诊数据
    '入参：strValue:过滤条件
    '      bytType:0-不进行查找;1-病人ID;2-门诊号;3-按姓名模糊查找;4-挂号单;5-医保号
    '编制:刘兴洪
    '日期:2011-11-21 10:50:39
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim rsTmp As ADODB.Recordset, strSQL As String, strFilter As String
    Dim objList As ListItem
    Dim i As Long, j As Long, blnUnOutJoin As Boolean
    Dim strNO As String, strTmp As String
    Dim str门诊号 As String, str姓名 As String, str就诊卡号 As String, str医保号 As String, str挂号单开始号 As String
    Dim str病人ID As String
    Dim lng提前分诊小时 As Long, str号别 As String, lngPatiID As Long
    Dim bln分诊台签到排队 As Boolean
    On Error GoTo errHandle
    '预约挂号记录的刷新
    If LvwYY.SelectedItem Is Nothing Then
        strNO = ""
    Else
        strNO = LvwYY.SelectedItem.Text
    End If

    LockWindowUpdate LvwYY.Hwnd
    LvwYY.ListItems.Clear
    LvwYY.Sorted = False

    If blnFilter Then
        strFilter = mcllFilter("条件")
    Else
         '问题号:51223
        lng提前分诊小时 = CLng(zlDatabase.GetPara("提前N小时分诊", glngSys, mlngModul, 0))
        strFilter = " And A.发生时间 Between Trunc(sysdate)-" & mint有效天数 & " And sysdate + 1/24 * " & lng提前分诊小时   'gbytNODay :27600
    End If
    
    str就诊卡号 = CStr(mcllFilter("就诊卡号"))
    str姓名 = CStr(mcllFilter("病人姓名"))
    str门诊号 = CStr(mcllFilter("门诊号"))
    str医保号 = CStr(mcllFilter("医保号"))
    str挂号单开始号 = CStr(mcllFilter("挂号NO")(0))
    str病人ID = Val(mcllFilter("病人ID"))
    If str姓名 <> "" Or str就诊卡号 <> "" Or str门诊号 <> "" Or str医保号 <> "" Or Nvl(str病人ID) <> 0 Then blnUnOutJoin = True
    If str姓名 <> "" Then str姓名 = str姓名 & "%"
    
    If strValue <> "" Then
        Select Case bytType '0-不进行查找;1-病人ID;2-门诊号;3-按姓名模糊查找;4-挂号单;5-医保号
            Case 0  '不进行查找:按缺省过滤条件过滤
            Case 1  '病人ID
                str病人ID = Val(strValue)
                strFilter = strFilter & " And A.病人ID=[12]"
                blnUnOutJoin = True
            Case 2 '门诊号
                str门诊号 = strValue
                strFilter = strFilter & " And A.门诊号 = [11]"
                blnUnOutJoin = True
            Case 3  '按姓名模糊查找
                str姓名 = strValue
                strFilter = strFilter & " And A.姓名 Like [8]"
            Case 4 '挂号单
                str挂号单开始号 = strValue
                strFilter = strFilter & " And A.NO=[3]"
            Case 5 '医保号
                str医保号 = strValue
                strFilter = strFilter & " And B.医保号=[13]"
                blnUnOutJoin = True
        End Select
    End If
    
    'A.记录标志:表示0表示就诊病人,1-表示签到的病人,2-表示需要回诊就诊的病人; 3-表示已回诊但还未接收的病人;
    strFilter = strFilter & IIf(mstr分诊科室 <> "", " And Instr(','||[10]||',',','||A.执行部门id||',')>0", "") & _
            " And (Nvl(A.执行状态,0) = 0 And A.诊室 Is Null" & _
            IIf(mbytViewScrop(0) = 1, " Or nvl(A.执行状态,0) = 0 And A.诊室 Is Not Null", "") & _
            IIf(mbytViewScrop(1) = 1, " Or A.执行状态 = 2", "") & _
            IIf(mbytViewScrop(2) = 1, " Or A.执行状态 = 1", "") & _
            IIf(mbytViewScrop(3) = 1, " Or A.执行状态 = -1", "") & _
            " ) "
    '问题:43012
    'mbyt候诊排序方式和:Decode(A.预约,1, nvl(A.接收时间,A.登记时间),A.登记时间)
    
    
    If gbytRegistMode = 0 Then
        strSQL = _
            "Select A.诊室,A.ID,A.号别,C.号类,D.名称 as 挂号项目," & vbCrLf & _
            "      A.执行部门ID,E.名称 as 执行部门名称,A.NO,NVL(A.病人ID, 0) 病人ID,A.姓名," & vbCrLf & _
            "      NVL(B.门诊号, 0) 门诊号,B.就诊卡号,B.卡验证码,A.性别,A.年龄," & vbCrLf & _
            "      A.发生时间 as 发生时间," & vbCrLf & _
            "      decode(A.预约,1,nvl(A.接收时间,A.登记时间),A.登记时间) as 登记时间, " & _
            "      NVL(B.险类, 0) 险类,A.执行人," & _
            "      nvl(A.执行状态,0) as 执行状态,A.号序,A.摘要,decode(A.预约,1,'√','') as 预约,B.医保号,A.记录标志" & vbCrLf & _
            "  From 病人挂号记录 a,病人信息 b,挂号安排 c,收费项目目录 d,部门表 e " & vbCrLf & _
            " Where a.病人id=b.病人id " & IIf(blnUnOutJoin, "", "(+)") & " and ((nvl(A.执行状态,0)=2 and nvl(A.记录标志,0) in (0,1))   or Nvl(A.执行状态,0)<>2 ) " & _
            "           And a.执行部门id=e.id And a.号别=c.号码 And c.项目id=d.ID" & vbCrLf & strFilter & _
            "           And (E.站点='" & gstrNodeNo & "' Or E.站点 is Null) and a.记录性质=2 and a.记录状态=1" & vbNewLine & _
            "  "
    Else
        If Sys.Currentdate < gdatRegistTime Then
            strSQL = _
                "Select A.诊室,A.ID,A.号别,C.号类,D.名称 as 挂号项目," & vbCrLf & _
                "      A.执行部门ID,E.名称 as 执行部门名称,A.NO,NVL(A.病人ID, 0) 病人ID,A.姓名," & vbCrLf & _
                "      NVL(B.门诊号, 0) 门诊号,B.就诊卡号,B.卡验证码,A.性别,A.年龄," & vbCrLf & _
                "      A.发生时间 as 发生时间," & vbCrLf & _
                "      decode(A.预约,1,nvl(A.接收时间,A.登记时间),A.登记时间) as 登记时间, " & _
                "      NVL(B.险类, 0) 险类,A.执行人," & _
                "      nvl(A.执行状态,0) as 执行状态,A.号序,A.摘要,decode(A.预约,1,'√','') as 预约,B.医保号,A.记录标志" & vbCrLf & _
                "  From 病人挂号记录 a,病人信息 b,挂号安排 c,收费项目目录 d,部门表 e " & vbCrLf & _
                " Where a.病人id=b.病人id " & IIf(blnUnOutJoin, "", "(+)") & " and ((nvl(A.执行状态,0)=2 and nvl(A.记录标志,0) in (0,1))   or Nvl(A.执行状态,0)<>2 ) " & _
                "           And a.执行部门id=e.id And a.号别=c.号码 And c.项目id=d.ID" & vbCrLf & strFilter & _
                "           And (E.站点='" & gstrNodeNo & "' Or E.站点 is Null) and a.记录性质=2 and a.记录状态=1" & vbNewLine & _
                "  "
        Else
            strSQL = _
                "Select A.诊室,A.ID,A.号别,C.号类,D.名称 as 挂号项目," & vbCrLf & _
                "      A.执行部门ID,E.名称 as 执行部门名称,A.NO,NVL(A.病人ID, 0) 病人ID,A.姓名," & vbCrLf & _
                "      NVL(B.门诊号, 0) 门诊号,B.就诊卡号,B.卡验证码,A.性别,A.年龄," & vbCrLf & _
                "      A.发生时间 as 发生时间," & vbCrLf & _
                "      decode(A.预约,1,nvl(A.接收时间,A.登记时间),A.登记时间) as 登记时间, " & _
                "      NVL(B.险类, 0) 险类,A.执行人," & _
                "      nvl(A.执行状态,0) as 执行状态,A.号序,A.摘要,decode(A.预约,1,'√','') as 预约,B.医保号,A.记录标志" & vbCrLf & _
                "  From 病人挂号记录 a,病人信息 b,临床出诊号源 c,临床出诊记录 c1,收费项目目录 d,部门表 e " & vbCrLf & _
                " Where a.病人id=b.病人id " & IIf(blnUnOutJoin, "", "(+)") & " And ((nvl(A.执行状态,0)=2 and nvl(A.记录标志,0) in (0,1))   or Nvl(A.执行状态,0)<>2 ) " & _
                "           And a.执行部门id=e.id And a.出诊记录id=c1.id And c1.号源id=c.id And c.项目id=d.ID" & vbCrLf & strFilter & _
                "           And (E.站点='" & gstrNodeNo & "' Or E.站点 is Null) and a.记录性质=2 and a.记录状态=1" & vbNewLine & _
                "  "
        End If
    End If
     '50427
     Select Case mbyt候诊排序方式
     Case 0  '科室编码,号码,NO
        strSQL = strSQL & vbCrLf & " Order By e.编码,c.号码,a.NO "
     Case 1 '科室编码,号码,挂号时间
        strSQL = strSQL & vbCrLf & _
        " Order By e.编码,c.号码, Decode(A.预约,1,nvl(A.接收时间,A.登记时间),A.登记时间)"
     Case 2 '科室编码,号码,发生时间
        strSQL = strSQL & vbCrLf & " Order By e.编码,c.号码, A.发生时间,A.登记时间 " '问题号：51665
     End Select
     
     Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
        CDate(mcllFilter("挂号时间")(0)), CDate(mcllFilter("挂号时间")(1)), _
        str挂号单开始号, CStr(mcllFilter("挂号NO")(1)), _
        CStr(mcllFilter("发票号")(0)), CStr(mcllFilter("发票号")(1)), _
        Val(mcllFilter("科室")), _
        str姓名, CStr(mcllFilter("挂号员")), mstr分诊科室, _
        str门诊号, str病人ID, str医保号)

    With rsTmp
        If .RecordCount > 0 Then
            .MoveFirst
            str号别 = Nvl(!号别): lngPatiID = !病人ID
        End If
        Do While Not .EOF
            mstrRegistIdsed = mstrRegistIdsed & "," & Nvl(!id)
            Set objList = LvwYY.ListItems.Add(, , !NO, "ry", "ry")
            objList.SubItems(EnmCol.Enm号类) = zlCommFun.Nvl(!号类)
            objList.SubItems(EnmCol.Enm科室) = zlCommFun.Nvl(!执行部门名称)
            objList.SubItems(EnmCol.Enm挂号项目) = zlCommFun.Nvl(!挂号项目)
            objList.SubItems(EnmCol.Enm姓名) = zlCommFun.Nvl(!姓名)
            objList.SubItems(EnmCol.Enm门诊号) = IIf(!门诊号 = 0, "", CStr(!门诊号))
            objList.SubItems(EnmCol.Enm性别) = zlCommFun.Nvl(!性别)
            objList.SubItems(EnmCol.Enm年龄) = zlCommFun.Nvl(!年龄)
            objList.SubItems(EnmCol.Enm诊室) = zlCommFun.Nvl(!诊室)
            objList.SubItems(EnmCol.Enm医生) = zlCommFun.Nvl(!执行人)
            objList.SubItems(EnmCol.Enm发生时间) = Format(!发生时间, "YYYY-MM-DD HH:MM:SS") '51774
            objList.SubItems(EnmCol.Enm挂号时间) = Format(!登记时间, "YYYY-MM-DD HH:MM:SS")
            objList.SubItems(EnmCol.Enm号序) = "" & !号序
            objList.SubItems(EnmCol.Enm医保号) = Nvl(!医保号)
            objList.SubItems(EnmCol.Enm摘要) = Nvl(!摘要)
            objList.SubItems(EnmCol.Enm病人状态) = IIf(Nvl(!执行状态, 0) = 1, "已完成", IIf(Nvl(!执行状态, 0) = 2, "已接诊", IIf(Nvl(!执行状态, 0) = -1, "不就诊", _
                                                       IIf(zlCommFun.Nvl(!诊室) <> "", "已分诊", "待分诊"))))
            '95637：李南春，2016/7/17，预约签到都当作正常签到
            objList.Tag = !病人ID & "|" & !险类 & "|" & !就诊卡号 & "|" & !卡验证码 & "|" & !id & "|" & !号别 & "|" & !执行状态 & "|" & Nvl(!记录标志, 0)
            objList.ListSubItems(1).Tag = Nvl(!id)
            objList.ListSubItems(2).Tag = !执行部门id
            objList.ListSubItems(3).Tag = Nvl(!记录标志)
            bln分诊台签到排队 = Val(zlDatabase.GetPara("分诊台签到排队", glngSys, mlngModul, 0, , , , Val(!执行部门id))) = 1
            objList.ListSubItems(4).Tag = IIf(bln分诊台签到排队, 1, 0)
              
            If str号别 <> Nvl(!号别) Or lngPatiID <> !病人ID Then blnAuto签到 = False
            '0-等待接诊,1-完成就诊,2-正在就诊,-1标记为不就诊
            Select Case Nvl(!执行状态, 0)
            Case 0
                If Not (IsNull(!诊室) Or !病人ID = 0) Then
                    objList.Icon = "yf": objList.SmallIcon = "yf"
                    If Val(Nvl(!记录标志)) = 1 Or bln分诊台签到排队 = False Then objList.ForeColor = &H8000000C
                ElseIf zlCommFun.Nvl(!号类) = "专家" Then
                    objList.ForeColor = RGB(0, 0, 255)      '兰色
                End If
                'A.记录标志:表示0表示就诊病人,1-表示签到的病人,2-表示需要回诊就诊的病人; 3-表示已回诊但还未接收的病人;
                If Val(Nvl(!记录标志)) = 1 Then
                    objList.Icon = "rySign_in": objList.SmallIcon = "rySign_in"
                End If
            Case 2
                objList.Icon = "zz": objList.SmallIcon = "zz"
                objList.ForeColor = RGB(255, 192, 0)        '黄色
            Case 1
                objList.Icon = "yz": objList.SmallIcon = "yz"
                objList.ForeColor = RGB(255, 0, 0)          '红色
            Case -1
                objList.ForeColor = &HC000&                 '绿色
            End Select
            For j = 1 To LvwYY.ColumnHeaders.Count - 1
                objList.ListSubItems(j).ForeColor = objList.ForeColor
            Next
            .MoveNext
        Loop
        '95637:李南春,2016/7/18,如果只有一种类型的挂号单，直接签到
        If tbPage.Item(midx.idx_预约队列).Selected And blnAuto签到 Then
            Call ShowBillsAppointment(blnFilter, strValue, bytType, objCard)         '刷新列表后退出
            If zlIs允许签道() Then Screen.MousePointer = 0: Call zlExc签道(False)
            Screen.MousePointer = 0
            Exit Sub
        End If
    End With

    If Me.LvwYY.ListItems.Count > 0 Then
        LvwYY.ListItems(1).Selected = True
        For i = 1 To LvwYY.ListItems.Count
            If LvwYY.ListItems(i).Text = strNO Then
                LvwYY.ListItems(i).Selected = True
                LvwYY.Drag 0
                LvwYY.Drag 2
                Exit For
            End If
        Next
        LvwYY.SelectedItem.EnsureVisible
        'lvwYY_ItemClick LvwYY.SelectedItem
    Else
        lvwRoom.ListItems.Clear
    End If
    LockWindowUpdate 0
    LvwYY.Refresh
    Exit Sub
errHandle:
    LockWindowUpdate 0
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
    Resume
    End If
    Call SaveErrLog
End Sub

Public Property Get zlGetRegistIDsed() As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取界面上的挂号列表
    '返回:挂号列表,多个用逗号分离
    '编制:刘兴洪
    '日期:2014-03-11 16:01:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    zlGetRegistIDsed = mstrRegistIdsed
End Property
 

 Public Sub SendMsgModule(ByVal strNO As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:消息发送处理
    '入参: strNO-挂号单号
    '编制:刘兴洪
    '日期:2014-03-11 11:59:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    '0-挂号,1-预约,2-接收,3-取消预约 ,4-退号 预约有两种模式:0-挂号,此时预约要收费,1-预约,不收费
    If mobjMsgModule Is Nothing Then Exit Sub
    If mobjMsgModule.IsConnect = False Then Exit Sub



    strSQL = "" & _
    " Select A.id ,A.姓名,nvl(A.门诊号,B.门诊号) as 门诊号,A.病人Id,b.身份证号,A.NO,A.执行部门ID,C.名称 as 执行部门名称,A.诊室,A.执行人  " & _
    " From 病人挂号记录 A,病人信息 B,部门表 C  " & _
    " where A.No=[1] and a.记录状态 =1 And a.记录性质=1 and a.病人ID=b.病人id and a.执行部门id=c.id"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
    If rsTemp.EOF Then Exit Sub
    
    '3.1.1.  ZLHIS_REGIST_002 -门诊分诊通知
    '节点名称    属性    含义    重复    类型    缺省值  值域描述
    '<patient_info>
    '    <patient_id>病人ID</patient_id>
    '    <patient_name>病人姓名</patient_name>
    '    <identity_card>身份证号</identity_card>
    '    <out_number>门诊号</out_number>
    '</patient_info>
    '<register_info>
    '    <register_id>挂号id</register_id>
    '    <register_no>挂号单号</register_no>
    '    <register_dept_id>挂号科室id</register_dept_id>
    '    <register_dept_title>挂号科室</register_dept_title>
    '    <register_room>挂号诊室</register_room>
    '    <register_doctor>挂号医生</register_doctor>
    '</register_info>
    zlXML.ClearXmlText
 
    Call zlXML.AppendNode("patient_info")
        Call zlXML.appendData("patient_id", Val(Nvl(rsTemp!病人ID)))
        Call zlXML.appendData("patient_name", Nvl(rsTemp!姓名))
        Call zlXML.appendData("identity_card", Nvl(rsTemp!身份证号))
        Call zlXML.appendData("out_number", Nvl(rsTemp!门诊号))
    Call zlXML.AppendNode("patient_info", True)
    
    Call zlXML.AppendNode("triage_info")
        Call zlXML.appendData("register_id", Val(Nvl(rsTemp!id)))
        Call zlXML.appendData("register_no", strNO)
        Call zlXML.appendData("register_dept_id", Val(Nvl(rsTemp!执行部门id)))
        Call zlXML.appendData("register_dept_title", Nvl(rsTemp!执行部门名称))
        Call zlXML.appendData("register_doctor", Nvl(rsTemp!执行人))
        Call zlXML.appendData("triage_room", Nvl(rsTemp!诊室))
    Call zlXML.AppendNode("triage_info", True)
    Call mobjMsgModule.CommitMessage("ZLHIS_REGIST_002", zlXML.XmlText)
    zlXML.ClearXmlText
 End Sub
 
 Public Sub zlModiyPatiBaseInfo(ByVal frmMain As Form)
    '------------------------------------------------------------------------------------------------------------------------
    '功能：调整病人基本信息
    '入参：frmMain-父窗体
    '编制：李南春
    '日期：2014-07-03
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long
    Dim lng挂号ID As Long
    Dim strInfo As String
    On Error GoTo Errhand
    
    lng病人ID = 0: lng挂号ID = 0
    
    If Me.ActiveControl Is lvwHZPati Then
        If Not lvwHZPati.SelectedItem Is Nothing Then
            lng病人ID = CLng(Val(Split(lvwHZPati.SelectedItem.Tag, "|")(0)))
            lng挂号ID = CLng(Val(Split(lvwHZPati.SelectedItem.Tag, "|")(4)))
        End If
    ElseIf tbPage.Item(midx.idx_排队队列).Selected And Not lvwMain.SelectedItem Is Nothing Then
        lng病人ID = CLng(Val(Split(lvwMain.SelectedItem.Tag, "|")(0)))
        lng挂号ID = CLng(Val(Split(lvwMain.SelectedItem.Tag, "|")(4)))
    ElseIf tbPage.Item(midx.idx_预约队列).Selected And Not LvwYY.SelectedItem Is Nothing Then
        lng病人ID = CLng(Val(Split(LvwYY.SelectedItem.Tag, "|")(0)))
        lng挂号ID = CLng(Val(Split(LvwYY.SelectedItem.Tag, "|")(4)))
    End If
    
    If mobjPublicPatient Is Nothing Then
        On Error Resume Next
        Set mobjPublicPatient = CreateObject("zlPublicPatient.clsPublicPatient")
        Err.Clear: On Error GoTo 0
    End If
    If Not mobjPublicPatient Is Nothing Then
        If mobjPublicPatient.zlInitCommon(gcnOracle, glngSys, gstrDBUser) Then
            If mobjPublicPatient.ModipatiBaseInfo(Me, "门诊分诊", lng病人ID, lng挂号ID, 1) Then
                '重新刷新
                zlRefreshData (True)
            End If
            Exit Sub
        End If
    End If
    MsgBox "创建病人信息公共部件(zlPublicPatient.clsPublicPatient)失败！", vbExclamation, gstrSysName
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub zlPrintBarcode()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:为指定的病人打印条码
    '编制:李南春
    '日期:2014/9/2 09:43
    '问题:77412
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnPrint As Boolean, lng病人ID As Long, strNO As String
    Dim objLvw As ListView
    blnPrint = True
    If InStr(1, mstrPrivs, ";条码打印;") = 0 Then Exit Sub
    '没选择就退出
    If tbPage.Item(midx.idx_排队队列).Selected Then
        Set objLvw = lvwMain
        If objLvw.SelectedItem Is Nothing Then Exit Sub
    End If
    If tbPage.Item(midx.idx_预约队列).Selected Then
        Set objLvw = LvwYY
        If objLvw.SelectedItem Is Nothing Then Exit Sub
    End If
    lng病人ID = CLng(Val(Split(objLvw.SelectedItem.Tag, "|")(0)))
    strNO = Trim(objLvw.SelectedItem.Text)
    
    Select Case Val(zlDatabase.GetPara("条码打印方式", glngSys, mlngModul))
    Case 0
         blnPrint = False
    Case 2
         If MsgBox("你是否要打印病人条码?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then blnPrint = False
    End Select
    If blnPrint Then Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1113_1", Me, "病人ID=" & lng病人ID, "NO=" & strNO, "PrintEmpty=0", 2)
End Sub

Private Function Check签到(ByVal bln预约 As Boolean, ByVal lng病人ID As Long, ByVal lng挂号ID As Long, _
                Optional ByVal str发生时间 As String, Optional ByVal blnNeedMsg As Boolean, _
                Optional bln分诊台签到排队 As Boolean) As Boolean
    '功能：检查当前挂号单是否允许签到
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strMsg As String
    '分诊台来签到的病人都应该是要来就诊的病人，所有检查当天的排队信息
    On Error GoTo Errhand
    '更改诊室：
        '没有排队生成当天的队列
        '如果小于前时间生成当天的队列
        '预约在分诊的时候生成队列
    If bln预约 Then
        If CDate(Format(str发生时间, "YYYY-MM-DD")) > zlDatabase.Currentdate Then
            strMsg = "预约签到只针对今天以前的单据，如果要提前签到，请到门诊挂号管理处提前接收!"
            If blnNeedMsg Then
                MsgBox strMsg, vbInformation, gstrSysName
            ElseIf bln分诊台签到排队 Then
                RaiseEvent zlShowInfor(strMsg)
            End If
            Exit Function
        End If
    End If
    strSQL = "Select 业务ID,排队号码 From 排队叫号队列 Where 病人ID= [1] And  Trunc(排队时间) < sysdate And 业务类型 = 0 And 排队状态 IN (0,1,7)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查当前病人的排队队列", lng病人ID)
    If rsTemp.RecordCount = 0 Then
    ElseIf rsTemp.RecordCount > 1 Or (rsTemp.RecordCount = 1 And rsTemp!业务ID <> lng挂号ID) Then
        If bln分诊台签到排队 Then '挂号立即排队下还检查签到，一定是重新签到
            If blnNeedMsg Then
                If MsgBox("病人有其他挂号项目正处于排队中，此时签到将取消排队，是否继续签到?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Function
                End If
            Else
                strMsg = "病人有其他挂号项目正处于排队中，不能自动完成签到!"
                RaiseEvent zlShowInfor(strMsg)
                Exit Function
            End If
        End If
    ElseIf rsTemp!业务ID = lng挂号ID Then
        strMsg = "病人正处于排队中，无需重新签到!"
        If blnNeedMsg Then
            MsgBox strMsg, vbInformation, gstrSysName
        ElseIf bln分诊台签到排队 Then
            RaiseEvent zlShowInfor(strMsg)
        End If
        Exit Function
    End If
    Check签到 = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetFilterCons(ByVal strFindValue As String, ByVal objCard As Card, ByVal bytReadType As Byte, _
                               ByRef strFindValue_Out As String, ByRef bytType_Out As Byte) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取相关的查找条件
    '入参:bytReadType-读取类型(0-不区分;1-刷卡;2-读取身份证;3-读取IC卡)
    '出参:bytType_Out-0-不进行查找;1-病人ID;2-门诊号;3-按姓名模糊查找;4-挂号单;5-医保号
    '返回:成功返回true,否则返回Fale
    '编制:刘兴洪
    '日期:2018-02-08 16:09:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str门诊号 As String, str姓名 As String, str病人ID As String
    Dim strPassWord As String, strErrMsg As String, lng病人ID As Long
    On Error GoTo errHandle
    
    strFindValue_Out = "": bytType_Out = 0
    If strFindValue = "" Then GetFilterCons = True: Exit Function
    
    If bytReadType = 1 Then
        '读卡或刷卡
        If objCard.名称 = "姓名" Or objCard.名称 Like "*姓*名*" Then
             If gobjSquare.objSquareCard.zlGetPatiID(mlngDefaultCardID, strFindValue, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
        Else
            If gobjSquare.objSquareCard.zlGetPatiID(objCard.接口序号, strFindValue, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
        End If
        
        strFindValue_Out = lng病人ID: bytType_Out = 1
        GetFilterCons = True
        Exit Function
    ElseIf bytReadType = 2 Or objCard.名称 = "身份证号" Or objCard.名称 = "二代身份证" Then '读取身份证
            If gobjSquare.objSquareCard.zlGetPatiID("身份证号", strFindValue, False, lng病人ID, _
                strPassWord, strErrMsg) = False Then lng病人ID = 0
        strFindValue_Out = lng病人ID: bytType_Out = 1
        GetFilterCons = True
        Exit Function

    ElseIf bytReadType = 3 Or objCard.名称 = "IC卡号" Then '读取IC卡
        If gobjSquare.objSquareCard.zlGetPatiID("IC卡号", strFindValue, False, lng病人ID, _
            strPassWord, strErrMsg) = False Then lng病人ID = 0
        strFindValue_Out = lng病人ID: bytType_Out = 1
        GetFilterCons = True
        Exit Function

    ElseIf (Left(strFindValue, 1) = "-" And IsNumeric(Mid(strFindValue, 2))) Then
        str病人ID = Val(Mid(strFindValue, 2))
        strFindValue_Out = str病人ID: bytType_Out = 1
        GetFilterCons = True
        Exit Function
    ElseIf (Left(strFindValue, 1) = "*" And IsNumeric(Mid(strFindValue, 2))) Or objCard.名称 = "门诊号" Then
        str门诊号 = IIf(Left(strFindValue, 1) = "*", Val(Mid(strFindValue, 2)), Val(strFindValue))
        strFindValue_Out = str门诊号: bytType_Out = 2
        GetFilterCons = True
        Exit Function
    Else
       Select Case objCard.名称
       Case "姓名"
            str姓名 = strFindValue & "%"
            strFindValue_Out = str姓名: bytType_Out = 3
            GetFilterCons = True
            Exit Function
       Case "挂号单"
            strFindValue_Out = strFindValue: bytType_Out = 4
            GetFilterCons = True
            Exit Function
       Case "医保号"
            strFindValue_Out = strFindValue: bytType_Out = 5
            GetFilterCons = True
            Exit Function
       Case Else
            '其他类别的,获取相关的病人ID
            '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
            '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
            '第7位后,就只能用索引,不然取不到数
            If objCard.接口序号 <> 0 Then
                If gobjSquare.objSquareCard.zlGetPatiID(objCard.接口序号, strFindValue, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
            Else
                If gobjSquare.objSquareCard.zlGetPatiID(objCard.名称, strFindValue, False, lng病人ID, _
                    strPassWord, strErrMsg) = False Then Exit Function
            End If
            strFindValue_Out = lng病人ID: bytType_Out = 1
            GetFilterCons = True
            Exit Function
       End Select
    End If

    GetFilterCons = True
    Exit Function
errHandle:
  Screen.MousePointer = 0
  If ErrCenter = 1 Then
    Resume
End If
  Call SaveErrLog
End Function

Private Sub ShowBillRegister(blnFilter As Boolean, Optional strValue As String = "", _
    Optional bytType As Byte = 0, Optional objCard As Card, Optional ByVal blnAuto签到 As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取挂号数据
    '入参：strValue:过滤条件
    '      bytType:0-不进行查找;1-病人ID;2-门诊号;3-按姓名模糊查找;4-挂号单;5-医保号
    '编制:刘兴洪
    '日期:2018-2-8 10:50:39
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim objList As ListItem, blnUnOutJoin As Boolean
    Dim i As Long, j As Long, strFilter As String
    Dim str门诊号 As String, str姓名 As String, str就诊卡号 As String, str医保号 As String, str挂号单开始号 As String
    Dim strNO As String, strTmp As String
    Dim str号别 As String, lngPatiID As Long
    Dim lng提前分诊小时 As Long, str病人ID As String
    Dim bln分诊台签到排队 As Boolean
    On Error GoTo errHandle

    '挂号记录的刷新
    If lvwMain.SelectedItem Is Nothing Then
        strNO = ""
    Else
        strNO = lvwMain.SelectedItem.Text
    End If
   
    LockWindowUpdate lvwMain.Hwnd
    lvwMain.ListItems.Clear
    lvwMain.Sorted = False
    
    If blnFilter Then
        strFilter = mcllFilter("条件")
    Else
         '问题号:51223
        lng提前分诊小时 = CLng(zlDatabase.GetPara("提前N小时分诊", glngSys, mlngModul, 0))
        strFilter = " And A.发生时间 Between Trunc(sysdate)-" & mint有效天数 & " And sysdate + 1/24 * " & lng提前分诊小时   'gbytNODay :27600
    End If
    
    str就诊卡号 = CStr(mcllFilter("就诊卡号"))
    str姓名 = CStr(mcllFilter("病人姓名"))
    str门诊号 = CStr(mcllFilter("门诊号"))
    str医保号 = CStr(mcllFilter("医保号"))
    str挂号单开始号 = CStr(mcllFilter("挂号NO")(0))
    str病人ID = Val(mcllFilter("病人ID"))
    If str姓名 <> "" Or str就诊卡号 <> "" Or str门诊号 <> "" Or str医保号 <> "" Or Val(str病人ID) <> 0 Then blnUnOutJoin = True
    If str姓名 <> "" Then str姓名 = str姓名 & "%"
    
    If strValue <> "" Then
        Select Case bytType '0-不进行查找;1-病人ID;2-门诊号;3-按姓名模糊查找;4-挂号单;5-医保号
        Case 0  '不进行查找:按缺省过滤条件过滤
        Case 1  '病人ID
            str病人ID = Val(strValue)
            strFilter = strFilter & " And A.病人ID=[12]"
            blnUnOutJoin = True
        Case 2 '门诊号
            str门诊号 = strValue
            strFilter = strFilter & " And A.门诊号 = [11]"
            blnUnOutJoin = True
        Case 3  '按姓名模糊查找
            str姓名 = strValue
            strFilter = strFilter & " And A.姓名 Like [8]"
        Case 4 '挂号单
            str挂号单开始号 = strValue
            strFilter = strFilter & " And A.NO=[3]"
        Case 5 '医保号
            str医保号 = strValue
            strFilter = strFilter & " And B.医保号=[13]"
            blnUnOutJoin = True
        End Select
    End If
    
    'A.记录标志:表示0表示就诊病人,1-表示签到的病人,2-表示需要回诊就诊的病人; 3-表示已回诊但还未接收的病人;
    strFilter = strFilter & IIf(mstr分诊科室 <> "", " And Instr(','||[10]||',',','||A.执行部门id||',')>0", "") & _
            " And (Nvl(A.执行状态,0) = 0 And A.诊室 Is Null" & _
            IIf(mbytViewScrop(0) = 1, " Or nvl(A.执行状态,0) = 0 And A.诊室 Is Not Null", "") & _
            IIf(mbytViewScrop(1) = 1, " Or A.执行状态 = 2", "") & _
            IIf(mbytViewScrop(2) = 1, " Or A.执行状态 = 1", "") & _
            IIf(mbytViewScrop(3) = 1, " Or A.执行状态 = -1", "") & _
            " ) "
    '问题:43012
    'mbyt候诊排序方式和:Decode(A.预约,1, nvl(A.接收时间,A.登记时间),A.登记时间)
    
    '74898:李南春,2015/4/9,标记病人的呼叫状态
    '95637:李南春,2016/7/17 签到类型 0 -正常签到；1-回诊签到；2-转诊签到；4-换号签到；5-重新签到
    '      转诊号码显示转诊科室，转诊医生，转诊诊室
    If gbytRegistMode = 0 Then
        strSQL = _
            "Select decode(A.转诊科室ID,Null,A.诊室,A.转诊诊室) as 诊室,A.ID,A.号别,C.号类,D.名称 as 挂号项目," & vbCrLf & _
            "      Nvl(A.转诊科室ID,A.执行部门ID) as 执行部门ID,E.名称 as 执行部门名称,A.NO,NVL(A.病人ID, 0) 病人ID,A.姓名," & vbCrLf & _
            "      NVL(B.门诊号, 0) 门诊号,B.就诊卡号,B.卡验证码,A.性别,A.年龄," & vbCrLf & _
            "      A.发生时间 as 发生时间," & vbCrLf & _
            "      decode(A.预约,1,nvl(A.接收时间,A.登记时间),A.登记时间) as 登记时间, " & _
            "      NVL(B.险类, 0) 险类,decode(A.转诊科室ID,Null,A.执行人,A.转诊医生) as 执行人," & _
            "      f.呼叫医生 As 呼叫人, f.诊室 As 呼叫诊室, f.呼叫时间, " & _
            "      nvl(A.执行状态,0) as 执行状态,A.号序,A.摘要,decode(A.预约,1,'√','') as 预约,B.医保号,A.记录标志,f.排队状态" & vbCrLf & _
            "      ,Nvl(A.转诊状态, 10) as 转诊状态 " & vbCrLf & _
            "  From 病人挂号记录 a,病人信息 b,挂号安排 c,收费项目目录 d,部门表 e,排队叫号队列 f " & vbCrLf & _
            " Where a.病人id=b.病人id  " & IIf(blnUnOutJoin, "", "(+)") & " And a.ID=f.业务id(+) and ((nvl(A.执行状态,0)=2 and nvl(A.记录标志,0) in (0,1))   or Nvl(A.执行状态,0)<>2 ) " & _
            "           And Nvl(A.转诊科室ID,A.执行部门ID)=e.id And a.号别=c.号码 And c.项目id=d.ID" & vbCrLf & strFilter & _
            "           And (E.站点='" & gstrNodeNo & "' Or E.站点 is Null) and a.记录性质=1 and a.记录状态=1" & vbNewLine & _
            "  "
    Else
        If Sys.Currentdate < gdatRegistTime Then
            strSQL = _
                "Select decode(A.转诊科室ID,Null,A.诊室,A.转诊诊室) as 诊室,A.ID,A.号别,C.号类,D.名称 as 挂号项目," & vbCrLf & _
                "      Nvl(A.转诊科室ID,A.执行部门ID) as 执行部门ID,E.名称 as 执行部门名称,A.NO,NVL(A.病人ID, 0) 病人ID,A.姓名," & vbCrLf & _
                "      NVL(B.门诊号, 0) 门诊号,B.就诊卡号,B.卡验证码,A.性别,A.年龄," & vbCrLf & _
                "      A.发生时间 as 发生时间," & vbCrLf & _
                "      decode(A.预约,1,nvl(A.接收时间,A.登记时间),A.登记时间) as 登记时间, " & _
                "      NVL(B.险类, 0) 险类,decode(A.转诊科室ID,Null,A.执行人,A.转诊医生) as 执行人," & _
                "      f.呼叫医生 As 呼叫人, f.诊室 As 呼叫诊室, f.呼叫时间, " & _
                "      nvl(A.执行状态,0) as 执行状态,A.号序,A.摘要,decode(A.预约,1,'√','') as 预约,B.医保号,A.记录标志,f.排队状态" & vbCrLf & _
                "      ,Nvl(A.转诊状态, 10) as 转诊状态 " & vbCrLf & _
                "  From 病人挂号记录 a,病人信息 b,挂号安排 c,收费项目目录 d,部门表 e,排队叫号队列 f " & vbCrLf & _
                " Where a.病人id=b.病人id  " & IIf(blnUnOutJoin, "", "(+)") & " And a.ID=f.业务id(+) and ((nvl(A.执行状态,0)=2 and nvl(A.记录标志,0) in (0,1))   or Nvl(A.执行状态,0)<>2 ) " & _
                "           And Nvl(A.转诊科室ID,A.执行部门ID)=e.id And a.号别=c.号码 And c.项目id=d.ID" & vbCrLf & strFilter & _
                "           And (E.站点='" & gstrNodeNo & "' Or E.站点 is Null) and a.记录性质=1 and a.记录状态=1" & vbNewLine & _
                "  "
        Else
            strSQL = _
                "Select decode(A.转诊科室ID,Null,A.诊室,A.转诊诊室) as 诊室,A.ID,A.号别,C.号类,D.名称 as 挂号项目," & vbCrLf & _
                "      Nvl(A.转诊科室ID,A.执行部门ID) as 执行部门ID,E.名称 as 执行部门名称,A.NO,NVL(A.病人ID, 0) 病人ID,A.姓名," & vbCrLf & _
                "      NVL(B.门诊号, 0) 门诊号,B.就诊卡号,B.卡验证码,A.性别,A.年龄," & vbCrLf & _
                "      A.发生时间 as 发生时间," & vbCrLf & _
                "      decode(A.预约,1,nvl(A.接收时间,A.登记时间),A.登记时间) as 登记时间, " & _
                "      NVL(B.险类, 0) 险类,decode(A.转诊科室ID,Null,A.执行人,A.转诊医生) as 执行人," & _
                "      f.呼叫医生 As 呼叫人, f.诊室 As 呼叫诊室, f.呼叫时间, " & _
                "      nvl(A.执行状态,0) as 执行状态,A.号序,A.摘要,decode(A.预约,1,'√','') as 预约,B.医保号,A.记录标志,f.排队状态" & vbCrLf & _
                "      ,Nvl(A.转诊状态, 10) as 转诊状态 " & vbCrLf & _
                "  From 病人挂号记录 a,病人信息 b,临床出诊号源 c,临床出诊记录 c1,收费项目目录 d,部门表 e,排队叫号队列 f " & vbCrLf & _
                " Where a.病人id=b.病人id  " & IIf(blnUnOutJoin, "", "(+)") & " And a.ID=f.业务id(+) and ((nvl(A.执行状态,0)=2 and nvl(A.记录标志,0) in (0,1))   or Nvl(A.执行状态,0)<>2 ) " & _
                "           And Nvl(A.转诊科室ID,A.执行部门ID)=e.id And a.出诊记录id=c1.id And c1.号源id=c.id And c.项目id=d.ID" & vbCrLf & strFilter & _
                "           And (E.站点='" & gstrNodeNo & "' Or E.站点 is Null) and a.记录性质=1 and a.记录状态=1" & vbNewLine & _
                "  "
        End If
    End If
     '50427
     Select Case mbyt候诊排序方式
     Case 0  '科室编码,号码,NO
        strSQL = strSQL & vbCrLf & " Order By e.编码,c.号码,a.NO "
     Case 1 '科室编码,号码,挂号时间
        strSQL = strSQL & vbCrLf & _
        " Order By e.编码,c.号码, Decode(A.预约,1,nvl(A.接收时间,A.登记时间),A.登记时间)"
     Case 2 '科室编码,号码,发生时间
        strSQL = strSQL & vbCrLf & " Order By e.编码,c.号码, A.发生时间,A.登记时间 " '问题号：51665
     End Select
     
     Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
        CDate(mcllFilter("挂号时间")(0)), CDate(mcllFilter("挂号时间")(1)), _
        str挂号单开始号, CStr(mcllFilter("挂号NO")(1)), _
        CStr(mcllFilter("发票号")(0)), CStr(mcllFilter("发票号")(1)), _
        Val(mcllFilter("科室")), _
        str姓名, CStr(mcllFilter("挂号员")), mstr分诊科室, _
        str门诊号, str病人ID, str医保号)

    With rsTmp
        If .RecordCount > 0 Then
            .MoveFirst
            str号别 = Nvl(!号别): lngPatiID = !病人ID
        End If
        Do While Not .EOF
            mstrRegistIdsed = mstrRegistIdsed & "," & Nvl(!id)
            Set objList = lvwMain.ListItems.Add(, , !NO, "ry", "ry")
            objList.SubItems(EnmCol.Enm号类) = zlCommFun.Nvl(!号类)
            objList.SubItems(EnmCol.Enm科室) = zlCommFun.Nvl(!执行部门名称)
            objList.SubItems(EnmCol.Enm挂号项目) = zlCommFun.Nvl(!挂号项目)
            objList.SubItems(EnmCol.Enm姓名) = zlCommFun.Nvl(!姓名)
            objList.SubItems(EnmCol.Enm门诊号) = IIf(!门诊号 = 0, "", CStr(!门诊号))
            objList.SubItems(EnmCol.Enm性别) = zlCommFun.Nvl(!性别)
            objList.SubItems(EnmCol.Enm年龄) = zlCommFun.Nvl(!年龄)
            objList.SubItems(EnmCol.Enm诊室) = zlCommFun.Nvl(!诊室)
            objList.SubItems(EnmCol.Enm医生) = zlCommFun.Nvl(!执行人)
            objList.SubItems(EnmCol.Enm发生时间) = Format(!发生时间, "YYYY-MM-DD HH:MM:SS") '51774
            objList.SubItems(EnmCol.Enm挂号时间) = Format(!登记时间, "YYYY-MM-DD HH:MM:SS")
            objList.SubItems(EnmCol.Enm号序) = "" & !号序
            objList.SubItems(EnmCol.Enm医保号) = Nvl(!医保号)
            objList.SubItems(EnmCol.Enm摘要) = Nvl(!摘要)
            objList.SubItems(EnmCol.Enm预约) = Nvl(!预约)
            objList.SubItems(EnmCol.Enm呼叫人) = Nvl(!呼叫人)
            objList.SubItems(EnmCol.Enm呼叫诊室) = Nvl(!呼叫诊室)
            objList.SubItems(EnmCol.Enm呼叫时间) = Nvl(!呼叫时间)
            objList.SubItems(EnmCol.Enm病人状态) = IIf(Nvl(!执行状态, 0) = 1, "已完成", IIf(Nvl(!执行状态, 0) = 2, "已接诊", IIf(Nvl(!执行状态, 0) = -1, "不就诊", _
                                                       IIf(zlCommFun.Nvl(!诊室) <> "", "已分诊", "待分诊"))))
            '74898:李南春,2015/4/9,标记病人的呼叫状态
            objList.SubItems(EnmCol.Enm呼叫) = IIf(Nvl(!排队状态) = 1 Or Nvl(!排队状态) = 7, "√", "")
            objList.Tag = !病人ID & "|" & !险类 & "|" & !就诊卡号 & "|" & !卡验证码 & "|" & !id & "|" & !号别 & "|" & !执行状态 & "|" & Nvl(!记录标志, 0) & "|" & Nvl(!转诊状态, 10)
            objList.ListSubItems(1).Tag = Nvl(!id)
            objList.ListSubItems(2).Tag = !执行部门id
            objList.ListSubItems(3).Tag = Nvl(!记录标志)
            bln分诊台签到排队 = Val(zlDatabase.GetPara("分诊台签到排队", glngSys, mlngModul, 0, , , , Val(!执行部门id))) = 1
            objList.ListSubItems(4).Tag = IIf(bln分诊台签到排队, 1, 0)
              
            If str号别 <> Nvl(!号别) Or lngPatiID <> !病人ID Then blnAuto签到 = False
            '0-等待接诊,1-完成就诊,2-正在就诊,-1标记为不就诊
            Select Case Nvl(!执行状态, 0)
            Case 0
                If Not (IsNull(!诊室) Or !病人ID = 0) Then
                    objList.Icon = "yf": objList.SmallIcon = "yf"
                    If Val(Nvl(!记录标志)) = 1 Or bln分诊台签到排队 = False Then objList.ForeColor = &H8000000C
                ElseIf zlCommFun.Nvl(!号类) = "专家" Then
                    objList.ForeColor = RGB(0, 0, 255)      '兰色
                End If
                'A.记录标志:表示0表示就诊病人,1-表示签到的病人,2-表示需要回诊就诊的病人; 3-表示已回诊但还未接收的病人;
                If Val(Nvl(!记录标志)) = 1 Then
                    objList.Icon = "rySign_in": objList.SmallIcon = "rySign_in"
                End If
            Case 2
                objList.Icon = "zz": objList.SmallIcon = "zz"
                objList.ForeColor = RGB(255, 192, 0)        '黄色
            Case 1
                objList.Icon = "yz": objList.SmallIcon = "yz"
                objList.ForeColor = RGB(255, 0, 0)          '红色
            Case -1
                objList.ForeColor = &HC000&                 '绿色
            End Select
            
            For j = 1 To objList.ListSubItems.Count - 1
                objList.ListSubItems(j).ForeColor = objList.ForeColor
            Next
            .MoveNext
        Loop
        '95637:李南春,2016/7/18,如果只有一种类型的挂号单，直接签到
        If tbPage.Item(midx.idx_排队队列).Selected And blnAuto签到 Then
            Call ShowBillRegister(blnFilter, strValue, bytType, objCard)         '刷新列表后退出
            If zlIs允许签道() Then Screen.MousePointer = 0: Call zlExc签道(False)
            Screen.MousePointer = 0
            Exit Sub
        End If
    End With
    
    If Me.lvwMain.ListItems.Count > 0 Then
        lvwMain.ListItems(1).Selected = True
        For i = 1 To lvwMain.ListItems.Count
            If lvwMain.ListItems(i).Text = strNO Then
                lvwMain.ListItems(i).Selected = True
                lvwMain.Drag 0
                lvwMain.Drag 2
                Exit For
            End If
        Next
        lvwMain.SelectedItem.EnsureVisible
        lvwMain_ItemClick lvwMain.SelectedItem
    Else
        lvwRoom.ListItems.Clear
    End If
    LockWindowUpdate 0
    lvwMain.Refresh
    
    Exit Sub
errHandle:
    LockWindowUpdate 0
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ShowBillRegisterHZ(blnFilter As Boolean, Optional strValue As String = "", _
    Optional bytType As Byte = 0, Optional objCard As Card)
     '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取回诊数据
    '入参：strValue:过滤条件
    '      bytType:0-不进行查找;1-病人ID;2-门诊号;3-按姓名模糊查找;4-挂号单;5-医保号
    '编制:刘兴洪
    '日期:2018-2-8 10:50:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strIco As String, objList As ListItem
    Dim i As Long, j As Long, blnUnOutJoin As Boolean
    Dim str门诊号 As String, str姓名 As String, str就诊卡号 As String, str医保号 As String, str挂号单开始号 As String
    Dim strNO As String, str病人ID As String
    Dim strHzWhere As String '回诊病人信息条件
    Dim blnBusy As Boolean, strFilter As String
    Dim str号别 As String, lngPatiID As Long
    Dim lng提前分诊小时 As Long
    Dim bln分诊台签到排队 As Boolean
    
    On Error GoTo errHandle
    '加载回诊病人信息
    If lvwHZPati.SelectedItem Is Nothing Then
        strNO = ""
    Else
        strNO = lvwHZPati.SelectedItem.Text
    End If
    LockWindowUpdate lvwHZPati.Hwnd
    lvwHZPati.ListItems.Clear
    lvwHZPati.Sorted = False
    
    If blnFilter Then
        strFilter = mcllFilter("条件")
    Else
         '问题号:51223
        lng提前分诊小时 = CLng(zlDatabase.GetPara("提前N小时分诊", glngSys, mlngModul, 0))
        strFilter = " And A.发生时间 Between Trunc(sysdate)-" & mint有效天数 & " And sysdate + 1/24 * " & lng提前分诊小时   'gbytNODay :27600
    End If
    
    str就诊卡号 = CStr(mcllFilter("就诊卡号"))
    str姓名 = CStr(mcllFilter("病人姓名"))
    str门诊号 = CStr(mcllFilter("门诊号"))
    str医保号 = CStr(mcllFilter("医保号"))
    str挂号单开始号 = CStr(mcllFilter("挂号NO")(0))
    str病人ID = Val(mcllFilter("病人ID"))
    If str姓名 <> "" Or str就诊卡号 <> "" Or str门诊号 <> "" Or str医保号 <> "" Or Val(str病人ID) <> 0 Then blnUnOutJoin = True
    If str姓名 <> "" Then str姓名 = str姓名 & "%"
    
    If strValue <> "" Then
        Select Case bytType '0-不进行查找;1-病人ID;2-门诊号;3-按姓名模糊查找;4-挂号单;5-医保号
        Case 0  '不进行查找:按缺省过滤条件过滤
        Case 1  '病人ID
            str病人ID = Val(strValue)
            strFilter = strFilter & " And A.病人ID=[12]"
            blnUnOutJoin = True
        Case 2 '门诊号
            str门诊号 = strValue
            strFilter = strFilter & " And A.门诊号 = [11]"
            blnUnOutJoin = True
        Case 3  '按姓名模糊查找
            str姓名 = strValue
            strFilter = strFilter & " And A.姓名 Like [8]"
        Case 4 '挂号单
            str挂号单开始号 = strValue
            strFilter = strFilter & " And A.NO=[3]"
        Case 5 '医保号
            str医保号 = strValue
            strFilter = strFilter & " And B.医保号=[13]"
            blnUnOutJoin = True
        End Select
    End If
    
    strHzWhere = strFilter & IIf(mstr分诊科室 <> "", " And Instr(','||[10]||',',','||A.执行部门id||',')>0", "")
    'A.记录标志:表示0表示就诊病人,1-表示签到的病人,2-表示需要回诊就诊的病人; 3-表示已回诊但还未接收的病人;
    If gbytRegistMode = 0 Then
        strSQL = _
            "Select A.诊室,A.ID,A.号别,C.号类,D.名称 as 挂号项目," & vbCrLf & _
            "      A.执行部门ID,E.名称 as 执行部门名称,A.NO,NVL(A.病人ID, 0) 病人ID,A.姓名," & vbCrLf & _
            "      NVL(B.门诊号, 0) 门诊号,B.就诊卡号,B.卡验证码,A.性别,A.年龄," & vbCrLf & _
            "      A.发生时间 as 发生时间," & vbCrLf & _
            "      Decode(A.预约,1, nvl(A.接收时间,A.登记时间),A.登记时间) as 登记时间,NVL(B.险类, 0) 险类, " & _
            "       A.执行人,nvl(A.执行状态,0) as 执行状态,A.记录标志,A.号序,A.摘要,decode(A.预约,1,'√','') as 预约,B.医保号" & vbCrLf & _
            "  From 病人挂号记录 a,病人信息 b,挂号安排 c,收费项目目录 d,部门表 e" & vbCrLf & _
            " Where a.病人id=b.病人id " & IIf(blnUnOutJoin, "", "(+)") & _
            "           And (A.执行状态=2 and nvl(A.记录标志,0) in (2,3) )" & _
            "           And a.执行部门id=e.id And a.号别=c.号码 And c.项目id=d.ID" & vbCrLf & strHzWhere & _
            " And (E.站点='" & gstrNodeNo & "' Or E.站点 is Null) and a.记录性质=1 and a.记录状态=1 " & vbNewLine & _
           "  "
    Else
        If Sys.Currentdate < gdatRegistTime Then
            strSQL = _
             "Select A.诊室,A.ID,A.号别,C.号类,D.名称 as 挂号项目," & vbCrLf & _
             "      A.执行部门ID,E.名称 as 执行部门名称,A.NO,NVL(A.病人ID, 0) 病人ID,A.姓名," & vbCrLf & _
             "      NVL(B.门诊号, 0) 门诊号,B.就诊卡号,B.卡验证码,A.性别,A.年龄," & vbCrLf & _
             "      A.发生时间 as 发生时间," & vbCrLf & _
             "      Decode(A.预约,1, nvl(A.接收时间,A.登记时间),A.登记时间) as 登记时间,NVL(B.险类, 0) 险类, " & _
             "       A.执行人,nvl(A.执行状态,0) as 执行状态,A.记录标志,A.号序,A.摘要,decode(A.预约,1,'√','') as 预约,B.医保号" & vbCrLf & _
             "  From 病人挂号记录 a,病人信息 b,挂号安排 c,收费项目目录 d,部门表 e" & vbCrLf & _
             " Where a.病人id=b.病人id " & IIf(blnUnOutJoin, "", "(+)") & _
             "           And (A.执行状态=2 and nvl(A.记录标志,0) in (2,3) )" & _
             "           And a.执行部门id=e.id And a.号别=c.号码 And c.项目id=d.ID" & vbCrLf & strHzWhere & _
             " And (E.站点='" & gstrNodeNo & "' Or E.站点 is Null) and a.记录性质=1 and a.记录状态=1 " & vbNewLine & _
            "  "
        Else
            strSQL = _
                "Select A.诊室,A.ID,A.号别,C.号类,D.名称 as 挂号项目," & vbCrLf & _
                "      A.执行部门ID,E.名称 as 执行部门名称,A.NO,NVL(A.病人ID, 0) 病人ID,A.姓名," & vbCrLf & _
                "      NVL(B.门诊号, 0) 门诊号,B.就诊卡号,B.卡验证码,A.性别,A.年龄," & vbCrLf & _
                "      A.发生时间 as 发生时间," & vbCrLf & _
                "      Decode(A.预约,1, nvl(A.接收时间,A.登记时间),A.登记时间) as 登记时间,NVL(B.险类, 0) 险类, " & _
                "       A.执行人,nvl(A.执行状态,0) as 执行状态,A.记录标志,A.号序,A.摘要,decode(A.预约,1,'√','') as 预约,B.医保号" & vbCrLf & _
                "  From 病人挂号记录 a,病人信息 b,临床出诊号源 c,临床出诊记录 c1,收费项目目录 d,部门表 e" & vbCrLf & _
                " Where a.病人id=b.病人id " & IIf(blnUnOutJoin, "", "(+)") & _
                "           And (A.执行状态=2 and nvl(A.记录标志,0) in (2,3) )" & _
                "           And a.执行部门id=e.id And a.出诊记录id=c1.id And c1.号源id = c.id And c.项目id=d.ID" & vbCrLf & strHzWhere & _
                " And (E.站点='" & gstrNodeNo & "' Or E.站点 is Null) and a.记录性质=1 and a.记录状态=1 " & vbNewLine & _
               "  "
        End If
    End If
    '问题:43012
    '50427
     Select Case mbyt候诊排序方式
     Case 0  '科室编码,号码,NO
        strSQL = strSQL & vbCrLf & " Order By e.编码,c.号码,a.NO "
     Case 1 '科室编码,号码,挂号时间
        strSQL = strSQL & vbCrLf & _
        " Order By e.编码,c.号码, Decode(A.预约,1,nvl(A.接收时间,A.登记时间),A.登记时间)"
     Case 2 '科室编码,号码,发生时间
        strSQL = strSQL & vbCrLf & " Order By e.编码,c.号码, A.发生时间,A.登记时间 " '问题号:51665
     End Select
     
    'mbyt候诊排序方式和:Decode(A.预约,1, nvl(A.接收时间,A.登记时间),A.登记时间)
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, _
        CDate(mcllFilter("挂号时间")(0)), CDate(mcllFilter("挂号时间")(1)), _
        str挂号单开始号, CStr(mcllFilter("挂号NO")(1)), _
        CStr(mcllFilter("发票号")(0)), CStr(mcllFilter("发票号")(1)), _
        Val(mcllFilter("科室")), _
        str姓名, CStr(mcllFilter("挂号员")), mstr分诊科室, _
        str门诊号, str病人ID, str医保号)
   With rsTmp
        If .RecordCount > 0 Then .MoveFirst
        Do While Not .EOF
            mstrRegistIdsed = mstrRegistIdsed & "," & Nvl(!id)
            '0表示就诊病人,1-表示签到的病人,2-表示需要回诊就诊的病人; 3-表示已回诊但还未接收的病人;
            If Val(Nvl(rsTmp!记录标志)) = 2 Then
                strIco = "ManStop"
                If Nvl(!性别) Like "*女*" Then strIco = "WomanStop"
            Else
                strIco = "ManSign_in"
                If Nvl(!性别) Like "*女*" Then strIco = "WomanSign_in"
            End If
            Set objList = lvwHZPati.ListItems.Add(, , !NO, strIco, strIco)
            objList.SubItems(EnmCol.Enm号类) = zlCommFun.Nvl(!号类)
            objList.SubItems(EnmCol.Enm科室) = zlCommFun.Nvl(!执行部门名称)
            objList.SubItems(EnmCol.Enm挂号项目) = zlCommFun.Nvl(!挂号项目)
            objList.SubItems(EnmCol.Enm姓名) = zlCommFun.Nvl(!姓名)
            objList.SubItems(EnmCol.Enm门诊号) = IIf(!门诊号 = 0, "", CStr(!门诊号))
            objList.SubItems(EnmCol.Enm性别) = zlCommFun.Nvl(!性别)
            objList.SubItems(EnmCol.Enm年龄) = zlCommFun.Nvl(!年龄)
            objList.SubItems(EnmCol.Enm诊室) = zlCommFun.Nvl(!诊室)
            objList.SubItems(EnmCol.Enm医生) = zlCommFun.Nvl(!执行人)
            objList.SubItems(EnmCol.Enm发生时间) = Format(!发生时间, "YYYY-MM-DD HH:MM:SS") ''51774
            objList.SubItems(EnmCol.Enm挂号时间) = Format(!登记时间, "YYYY-MM-DD HH:MM:SS")
            objList.SubItems(EnmCol.Enm号序) = "" & !号序
            objList.SubItems(EnmCol.Enm医保号) = Nvl(!医保号)
            objList.SubItems(EnmCol.Enm摘要) = Nvl(!摘要)
            objList.SubItems(EnmCol.Enm预约) = Nvl(!预约)
            objList.SubItems(EnmCol.Enm病人状态) = IIf(Nvl(!执行状态, 0) = 1, "已完成", IIf(Nvl(!执行状态, 0) = 2, "已接诊", IIf(Nvl(!执行状态, 0) = -1, "不就诊", _
                                                       IIf(zlCommFun.Nvl(!诊室) <> "", "已分诊", "待分诊"))))
            objList.Tag = !病人ID & "|" & !险类 & "|" & !就诊卡号 & "|" & !卡验证码 & "|" & !id & "|" & !号别 & "|" & !执行状态 & "|" & Nvl(!记录标志, 0)
            objList.ListSubItems(1).Tag = Nvl(!id)
            objList.ListSubItems(2).Tag = !执行部门id
            objList.ListSubItems(3).Tag = Nvl(!记录标志)
            bln分诊台签到排队 = Val(zlDatabase.GetPara("分诊台签到排队", glngSys, mlngModul, 0, , , , Val(!执行部门id))) = 1
            objList.ListSubItems(4).Tag = IIf(bln分诊台签到排队, 1, 0)
            'objList.ForeColor = RGB(255, 192, 0)        '黄色
            For j = 1 To lvwHZPati.ColumnHeaders.Count - 1
                objList.ListSubItems(j).ForeColor = objList.ForeColor
            Next
            .MoveNext
        Loop
    End With
    If mstrRegistIdsed <> "" Then mstrRegistIdsed = Mid(mstrRegistIdsed, 2)
    
    If Me.lvwHZPati.ListItems.Count > 0 Then
        lvwHZPati.ListItems(1).Selected = True
        For i = 1 To lvwHZPati.ListItems.Count
            If lvwHZPati.ListItems(i).Text = strNO Then
                lvwHZPati.ListItems(i).Selected = True
                lvwHZPati.Drag 0
                lvwHZPati.Drag 2
                Exit For
            End If
        Next
        lvwHZPati.SelectedItem.EnsureVisible
        lvwHZPati_ItemClick lvwHZPati.SelectedItem
    End If
    LockWindowUpdate 0
    lvwHZPati.Refresh
    Exit Sub
errHandle:
    LockWindowUpdate 0
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadRooms()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim objList As ListItem, i As Integer, j As Integer
    Dim strTmp As String
    Dim blnBusy As Boolean
    On Error GoTo errHandle

    '79694:李南春,2014/11/25,根据参数读取门诊诊室
    blnBusy = Val(zlDatabase.GetPara("诊室忙时允许分诊", glngSys, mlngModul, 0)) = 1
    '诊室情况刷新
    If gbytRegistMode = 0 Then
        strSQL = "Select Distinct R.编码, R.名称, R.缺省标志 As 忙闲状态, T.候诊, T.在诊, T.当日已诊" & vbNewLine & _
                "From 门诊诊室 R, 挂号安排诊室 S, 挂号安排 P," & vbNewLine & _
                "     (Select 诊室, Sum(Decode(执行状态, Null, 1, 0, 1, 0)) As 候诊, Sum(Decode(执行状态, 2, 1, 0)) As 在诊," & vbNewLine & _
                "              Sum(Decode(执行状态, 1, Decode(Sign(Trunc(Sysdate) - 执行时间), 1, 0, 1), 0)) As 当日已诊" & vbNewLine & _
                "       From 病人挂号记录" & vbNewLine & _
                "       Where 发生时间 > Sysdate - " & mint有效天数 & " And 诊室 Is Not Null and 记录性质=1 and 记录状态=1  " & vbNewLine & _
                "       Group By 诊室) T" & vbNewLine & _
                "Where R.名称 = S.门诊诊室 And S.号表id = P.ID And R.名称 = T.诊室(+) " & _
                IIf(blnBusy, " ", " And R.缺省标志=0 ") & _
                " And (R.站点='" & gstrNodeNo & "' Or R.站点 is Null)" & vbNewLine & _
                IIf(mstr分诊科室 <> "", " And Instr(','||[1]||',',','||P.科室id||',')>0", "") & vbNewLine & _
                "Order By R.编码"
    Else
        If Sys.Currentdate < gdatRegistTime Then
            strSQL = "Select Distinct R.编码, R.名称, R.缺省标志 As 忙闲状态, T.候诊, T.在诊, T.当日已诊" & vbNewLine & _
                    "From 门诊诊室 R, 挂号安排诊室 S, 挂号安排 P," & vbNewLine & _
                    "     (Select 诊室, Sum(Decode(执行状态, Null, 1, 0, 1, 0)) As 候诊, Sum(Decode(执行状态, 2, 1, 0)) As 在诊," & vbNewLine & _
                    "              Sum(Decode(执行状态, 1, Decode(Sign(Trunc(Sysdate) - 执行时间), 1, 0, 1), 0)) As 当日已诊" & vbNewLine & _
                    "       From 病人挂号记录" & vbNewLine & _
                    "       Where 发生时间 > Sysdate - " & mint有效天数 & " And 诊室 Is Not Null and 记录性质=1 and 记录状态=1  " & vbNewLine & _
                    "       Group By 诊室) T" & vbNewLine & _
                    "Where R.名称 = S.门诊诊室 And S.号表id = P.ID And R.名称 = T.诊室(+) " & _
                    IIf(blnBusy, " ", " And R.缺省标志=0 ") & _
                    " And (R.站点='" & gstrNodeNo & "' Or R.站点 is Null)" & vbNewLine & _
                    IIf(mstr分诊科室 <> "", " And Instr(','||[1]||',',','||P.科室id||',')>0", "") & vbNewLine & _
                    "Order By R.编码"
        Else
            strSQL = "Select Distinct R.编码, R.名称, R.缺省标志 As 忙闲状态, T.候诊, T.在诊, T.当日已诊" & vbNewLine & _
                    "From 门诊诊室 R, 临床出诊诊室记录 S, 临床出诊记录 P," & vbNewLine & _
                    "     (Select 诊室, Sum(Decode(执行状态, Null, 1, 0, 1, 0)) As 候诊, Sum(Decode(执行状态, 2, 1, 0)) As 在诊," & vbNewLine & _
                    "              Sum(Decode(执行状态, 1, Decode(Sign(Trunc(Sysdate) - 执行时间), 1, 0, 1), 0)) As 当日已诊" & vbNewLine & _
                    "       From 病人挂号记录" & vbNewLine & _
                    "       Where 发生时间 > Sysdate - " & mint有效天数 & " And 诊室 Is Not Null and 记录性质=1 and 记录状态=1  " & vbNewLine & _
                    "       Group By 诊室) T" & vbNewLine & _
                    "Where R.id = S.诊室id And S.记录id = P.ID And R.名称 = T.诊室(+) " & _
                    IIf(blnBusy, " ", " And R.缺省标志=0 ") & _
                    " And (R.站点='" & gstrNodeNo & "' Or R.站点 is Null)" & vbNewLine & _
                    IIf(mstr分诊科室 <> "", " And Instr(','||[1]||',',','||P.科室id||',')>0", "") & vbNewLine & _
                    "Order By R.编码"
        End If
    End If
    'gbytNODay: 问题:27600

    LockWindowUpdate lvwRoom.Hwnd
    lvwRoom.ListItems.Clear
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr分诊科室)
    With rsTmp
        If .RecordCount > 0 Then rsTmp.MoveFirst
        Do While Not .EOF
            Set objList = lvwRoom.ListItems.Add(, "K" & !编码, !名称, "bm", "bm")
            objList.SubItems(1) = IIf(!忙闲状态 <> 0, "忙", "")
            objList.SubItems(2) = Format(!候诊, "0;0; ; ")
            objList.SubItems(3) = Format(!在诊, "0;0; ; ")
            objList.SubItems(4) = Format(!当日已诊, "0;0; ; ")
            objList.SubItems(5) = ""
            If !忙闲状态 <> 0 Then
                objList.ForeColor = RGB(255, 0, 0)
                For j = 1 To Me.lvwRoom.ColumnHeaders.Count - 1
                    objList.ListSubItems(j).ForeColor = objList.ForeColor
                Next
            End If
            If InStr(1, strTmp & ",", "," & !名称 & "") = 0 Then
                strTmp = strTmp & "," & !名称 & ""
            End If
            rsTmp.MoveNext
        Loop
    End With
    strTmp = Mid(strTmp, 2)
    If gbytRegistMode = 0 Then
        strSQL = "Select Distinct S.门诊诊室,D.名称" & _
                " From 挂号安排诊室 S,挂号安排 P,部门表 D" & _
                " Where S.号表id=P.ID And P.科室id=D.ID And Instr(','||[1]||',',','||S.门诊诊室||',')>0" & _
                " And (D.站点='" & gstrNodeNo & "' Or D.站点 is Null)" & vbNewLine & _
                IIf(mstr分诊科室 <> "", " And Instr(','||[2]||',',','||P.科室id||',')>0", "")
    Else
        If Sys.Currentdate < gdatRegistTime Then
            strSQL = "Select Distinct S.门诊诊室,D.名称" & _
                    " From 挂号安排诊室 S,挂号安排 P,部门表 D" & _
                    " Where S.号表id=P.ID And P.科室id=D.ID And Instr(','||[1]||',',','||S.门诊诊室||',')>0" & _
                    " And (D.站点='" & gstrNodeNo & "' Or D.站点 is Null)" & vbNewLine & _
                    IIf(mstr分诊科室 <> "", " And Instr(','||[2]||',',','||P.科室id||',')>0", "")
        Else
            strSQL = "Select Distinct S.名称 As 门诊诊室,D.名称" & _
                    " From 门诊诊室 S,临床出诊诊室记录 S1,临床出诊记录 P,临床出诊号源 E,部门表 D" & _
                    " Where S1.记录id=P.ID And P.号源ID=E.ID And E.科室id=D.ID And S.ID=S1.诊室ID And Instr(','||[1]||',',','||S.名称||',')>0" & _
                    " And (D.站点='" & gstrNodeNo & "' Or D.站点 is Null)" & vbNewLine & _
                    IIf(mstr分诊科室 <> "", " And Instr(','||[2]||',',','||E.科室id||',')>0", "")
        End If
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strTmp, mstr分诊科室)

    '补充可分配到该诊室的科室
    For Each objList In Me.lvwRoom.ListItems
        strTmp = ""
        rsTmp.Filter = "门诊诊室='" & objList.Text & "'"
        Do While Not rsTmp.EOF
            If InStr(1, strTmp & ";", ";" & rsTmp!名称 & ";") = 0 Then strTmp = strTmp & ";" & rsTmp!名称
            rsTmp.MoveNext
        Loop
        If strTmp <> "" Then strTmp = Mid(strTmp, 2)
        objList.SubItems(5) = strTmp
    Next
    LockWindowUpdate 0
    Exit Sub
errHandle:
    LockWindowUpdate 0
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

