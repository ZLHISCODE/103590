VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmCooperateUnitsReg 
   BackColor       =   &H8000000D&
   Caption         =   "合约单位安排控制"
   ClientHeight    =   10230
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   13245
   LinkTopic       =   "Form1"
   ScaleHeight     =   10230
   ScaleWidth      =   13245
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraLine 
      Height          =   88
      Left            =   1200
      TabIndex        =   17
      Top             =   2280
      Width           =   11535
   End
   Begin VB.PictureBox picList 
      BorderStyle     =   0  'None
      Height          =   1530
      Left            =   120
      ScaleHeight     =   1530
      ScaleWidth      =   12720
      TabIndex        =   0
      Top             =   240
      Width           =   12720
      Begin VB.Frame fraInfo 
         Caption         =   "基本信息"
         Height          =   1260
         Left            =   240
         TabIndex        =   1
         Top             =   120
         Width           =   12255
         Begin VB.CommandButton cmdOK 
            Caption         =   "确定(&O)"
            Height          =   350
            Left            =   9360
            TabIndex        =   19
            Top             =   720
            Width           =   1100
         End
         Begin VB.CommandButton cmdCancel 
            Caption         =   "取消(&C)"
            Height          =   350
            Left            =   10575
            TabIndex        =   18
            Top             =   720
            Width           =   1100
         End
         Begin VB.TextBox txt号别 
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   720
            MaxLength       =   5
            TabIndex        =   8
            Top             =   307
            Width           =   960
         End
         Begin VB.ComboBox cboItem 
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3360
            TabIndex        =   7
            Text            =   "cboItem"
            Top             =   705
            Width           =   2235
         End
         Begin VB.ComboBox cboDoctor 
            Enabled         =   0   'False
            Height          =   300
            Left            =   6720
            TabIndex        =   6
            Top             =   705
            Width           =   2115
         End
         Begin VB.ComboBox cbo科室 
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   720
            TabIndex        =   5
            Text            =   "cbo科室"
            Top             =   705
            Width           =   2115
         End
         Begin VB.CheckBox chk病案 
            Caption         =   "挂号时必须建病案"
            Enabled         =   0   'False
            Height          =   195
            Left            =   5160
            TabIndex        =   4
            Top             =   360
            Width           =   1845
         End
         Begin VB.CheckBox chk序号控制 
            Caption         =   "序号控制"
            Enabled         =   0   'False
            Height          =   255
            Left            =   1800
            TabIndex        =   3
            Top             =   330
            Width           =   1095
         End
         Begin VB.ComboBox cbo号类 
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3360
            TabIndex        =   2
            Text            =   "cbo号类"
            Top             =   307
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "号别"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   120
            TabIndex        =   13
            Top             =   367
            Width           =   390
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "科室"
            Height          =   180
            Left            =   120
            TabIndex        =   12
            Top             =   765
            Width           =   360
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "项目"
            Height          =   180
            Left            =   3000
            TabIndex        =   11
            Top             =   765
            Width           =   360
         End
         Begin VB.Label lbl医生 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "院内医生"
            Height          =   180
            Left            =   5940
            TabIndex        =   10
            Top             =   765
            Width           =   720
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "号类"
            Height          =   180
            Left            =   3000
            TabIndex        =   9
            Top             =   367
            Width           =   360
         End
      End
   End
   Begin VB.PictureBox picContent 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   11160
      Left            =   5040
      ScaleHeight     =   11160
      ScaleWidth      =   8055
      TabIndex        =   15
      Top             =   2520
      Width           =   8055
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   1605
         Left            =   0
         TabIndex        =   16
         Top             =   -15
         Width           =   4290
         _Version        =   589884
         _ExtentX        =   7567
         _ExtentY        =   2831
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   14
      Top             =   9870
      Width           =   13245
      _ExtentX        =   23363
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCooperateUnitsReg.frx":0000
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   18283
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
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
Attribute VB_Name = "frmCooperateUnitsReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngModule As Long, mstrPrivs As String

Private Const conPane_Info = 1
Private Const conPane_Plan = 3
Private mlngPriItem As Long
Private mlng安排ID              As Long
Private mrs限号                 As ADODB.Recordset
Private mrs安排                 As ADODB.Recordset
Private mstr排班                As String '周日|全日||周一|白天||…………
Private mbln序号控制            As Boolean
Private mblnUnload As Boolean
Private mbln时段                As Boolean '如果安排设置了时段则严格按照时段来分配
Private mrs时间段               As ADODB.Recordset
Private mstrKey       As String
Private mrsSource     As ADODB.Recordset
Private mrsUnitsReg   As ADODB.Recordset
Private WithEvents mfrmReg       As frmCooperateReg
Attribute mfrmReg.VB_VarHelpID = -1
Private WithEvents mfrmRegNoTime As frmCooperateRegArrange
Attribute mfrmRegNoTime.VB_VarHelpID = -1
Private mblnOk      As Boolean
 
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case conPane_Info
            Item.Handle = picList.hWnd
        Case conPane_Plan
            Item.Handle = picContent
    End Select
End Sub
 
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If mbln序号控制 Then
        If Not mfrmReg Is Nothing Then
            If mfrmReg.SaveData() = False Then Exit Sub
        End If
    Else
        If Not mfrmRegNoTime Is Nothing Then
             If mfrmRegNoTime.SaveData() = False Then Exit Sub
        End If
    End If
    mblnOk = True
    Unload Me
End Sub

Private Sub Form_Resize()
      Err.Number = 0
     On Error Resume Next
     With Me.picList
         .Left = Me.ScaleLeft
         .Top = Me.ScaleTop
         .Width = Me.ScaleWidth
     End With
     With fraLine
         .Left = Me.ScaleLeft
         .Top = picList.Top + picList.Height
         .Width = Me.ScaleWidth
     End With
     With Me.picContent
         .Left = Me.ScaleLeft
         .Top = picList.Top + picList.Height - 18 * Screen.TwipsPerPixelY
         .Height = ScaleHeight - .Top - Me.stbThis.Height
         .Width = ScaleWidth
     End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mfrmReg Is Nothing Then Unload mfrmReg
    If Not mfrmRegNoTime Is Nothing Then Unload mfrmRegNoTime
    Set mfrmReg = Nothing
    Set mfrmRegNoTime = Nothing
     
End Sub

Public Function zlShowMe(ByVal lng安排ID As Long, ByVal lngModule As Long, ByVal strPrivs As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口
    '返回:设置成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-10-29 14:19:07
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
    mlng安排ID = lng安排ID
    If InitData() = False Then Exit Function
    InitPage
    Me.Show 1
    zlShowMe = mblnOk
End Function

Private Sub mfrmReg_frmUnload(ByVal blnCancel As Boolean)
    mblnOk = Not blnCancel
    Unload Me
End Sub

Private Sub mfrmRegNoTime_frmUnload(ByVal blnCancel As Boolean)
   mblnOk = Not blnCancel
   Unload Me
End Sub

Private Sub picContent_Resize()
    Err = 0: On Error Resume Next

    With picContent
        tbPage.Top = .ScaleTop
        tbPage.Left = .ScaleLeft
        'tbPage.Top = .ScaleTop
        tbPage.Width = .ScaleWidth
        tbPage.Height = .ScaleHeight
    End With

End Sub

Private Sub Form_Activate()
    Me.Icon = frmRegistPlan.Icon
    If mblnUnload Then mblnUnload = False: Unload Me
End Sub

'Private Sub InitPancel()
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:区哉设置
'    '编制:
'    '日期:2009-09-14 18:06:29
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim sngWidth As Single
'    Dim strReg   As String
'    Dim panThis  As Pane
'    Dim panInfo  As Pane
'
'    Set panInfo = dkpMan.CreatePane(conPane_Info, 900, 100, DockTopOf, Nothing)
'    'panThis.Title = "挂号安排信息"
'    panInfo.Handle = picList.hWnd
'    panInfo.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
'    panInfo.Tag = conPane_Info
'    panInfo.MaxTrackSize.Height = 100
'    panInfo.MinTrackSize.Height = 100
'    Set panThis = dkpMan.CreatePane(conPane_Plan, 160, 600, DockBottomOf, panInfo)
'    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoCaption Or PaneNoHideable
'    panThis.Title = "合作单位"
'    panThis.Tag = conPane_Plan
'
'    panThis.Handle = picContent.hWnd
'    'dkpMan.Options.ThemedFloatingFrames = True
'    dkpMan.Options.ThemedFloatingFrames = False
'    dkpMan.Options.HideClient = False
'    dkpMan.Options.UseSplitterTracker = False '实时拖动
'    dkpMan.Options.AlphaDockingContext = True
'   ' panThis.MaxTrackSize.Height = 600
'     panThis.MinTrackSize.Height = 600
'
'    '    Set panThis = dkpMan.CreatePane(conPane_Plan, 250, 580, DockBottomOf, panThis)
'    '    panThis.Title = ""
'    '    panThis.Tag = conPane_Plan
'    '    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
'    '    panThis.Handle = picPage.hWnd
'    '    dkpMan.Options.ThemedFloatingFrames = True
'    '    dkpMan.Options.HideClient = True
'    ' zlRestoreDockPanceToReg Me, dkpMan, "区域"
'
'End Sub

Private Sub InitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化页面控件
    '编制:刘兴洪
    '日期:2009-09-09 11:01:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem, objForm As Object

    Err = 0: On Error GoTo Errhand:
    
    tbPage.RemoveAll
    If mbln序号控制 Then
    Set mfrmReg = New frmCooperateReg
        mfrmReg.frmInit mlng安排ID
        Set ObjItem = tbPage.InsertItem(1, "", mfrmReg.hWnd, 0)
        ObjItem.Tag = 1
    Else
        Set mfrmRegNoTime = New frmCooperateRegArrange

        mfrmRegNoTime.frmInit mlng安排ID, mlngModule, mstrPrivs
        Set ObjItem = tbPage.InsertItem(2, "", mfrmRegNoTime.hWnd, 0)
        ObjItem.Tag = 2
   End If
    With tbPage
        tbPage.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        ' .PaintManager.Layout = xtpTabLayoutCompressed
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
        .PaintManager.ClientFrame = xtpTabFrameBorder
        .PaintManager.Position = xtpTabPositionTop
    End With
Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
 
'------------------------------------------------------------------------
'页面调用过程与方法
'------------------------------------------------------------------------
Public Function InitData() As Boolean
    Dim strSQL As String
    Dim lng安排ID       As Long
    Dim i       As Long
    Dim strTemp As String
    Dim rsTmp   As ADODB.Recordset
    If mlng安排ID = -1 Then Exit Function
    lng安排ID = mlng安排ID
    On Error GoTo Hd
    strSQL = "Select count(0) as 单位 From 挂号合作单位  Where Rownum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If Val(Nvl(rsTmp!单位)) = 0 Then MsgBox "没有设置【挂号合作单位】,请到数据字典中设置!", vbOKOnly, Me.Caption: Exit Function
    strSQL = " " & _
    "   Select A.Id as 安排ID,0 as 计划ID,A.号类,  A.号码,  A.科室id,  A.项目id, A.医生姓名,  A.医生id," & _
    "          A.周日,  A.周一,  A.周二,  A.周三,  A.周四,  A.周五,  A.周六,nvl(A.默认时段间隔,5) As 默认时段间隔, " & _
    "           A.病案必须,  A.分诊方式,  A.序号控制,  A.开始时间,  A.终止时间,B.名称 As 项目,D.名称 As 科室 " & _
    "   From 挂号安排 A,收费项目目录 B,部门表 D " & _
    "   Where A.项目id=b.Id(+) And A.科室id =d.Id(+) " & _
    "         And A.Id=[1]"
    
    Set mrs安排 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID)
         
    If mrs安排.EOF Then
        ShowMsgbox "未找到指定的号别,请检查!"
        Exit Function

    End If
        
    mstr排班 = ""

    For i = 0 To 6
        strTemp = Switch(i = 0, "日", i = 1, "一", i = 2, "二", i = 3, "三", i = 4, "四", i = 5, "五", True, "六")

        If Nvl(mrs安排("周" & strTemp)) <> "" Then
            If mstr排班 <> "" Then mstr排班 = mstr排班 & "||"
            mstr排班 = mstr排班 & "周" & strTemp & "|" & Nvl(mrs安排("周" & strTemp))
        End If
    Next

    If mstr排班 = "" Then Exit Function
'    strSQL = "Select 限制项目,限号数,  限约数,限制项目 as 星期 From  挂号安排限制 where 安排ID=[1]  Order BY 限制项目      "
'    Set mrs限号 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID)
    cbo号类.Text = Nvl(mrs安排!号类)
    txt号别.Tag = Nvl(mrs安排!安排Id)
      
    txt号别.Text = Nvl(mrs安排!号码)
    cbo科室.Text = Nvl(mrs安排!科室)
    cboItem.Text = Nvl(mrs安排!项目)
    cboDoctor.Text = Nvl(mrs安排!医生姓名)
    chk病案.Value = IIf(Val(Nvl(mrs安排!病案必须)) = 1, 1, 0)
    chk序号控制.Value = IIf(Val(Nvl(mrs安排!序号控制)) = 1, 1, 0):  chk序号控制.Tag = chk序号控制.Value
    mbln序号控制 = IIf(Val(Nvl(mrs安排!序号控制)) = 1, True, False)
    InitData = True
    Exit Function

Hd:

    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function

