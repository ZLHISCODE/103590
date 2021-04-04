VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdvicePrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人医嘱单打印"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6885
   Icon            =   "frmAdvicePrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraPrint 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4755
      Index           =   0
      Left            =   150
      TabIndex        =   4
      Tag             =   "常规打印"
      Top             =   720
      Visible         =   0   'False
      Width           =   6600
      Begin VB.Frame fraClear 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   3465
         TabIndex        =   24
         Top             =   4320
         Width           =   2385
         Begin VB.TextBox txtClearPage 
            Height          =   270
            Left            =   945
            MaxLength       =   3
            TabIndex        =   26
            Top             =   45
            Width           =   510
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "清除(&D)"
            Height          =   350
            Left            =   1485
            TabIndex        =   25
            Top             =   0
            Width           =   800
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "清除起始页"
            Height          =   180
            Left            =   0
            TabIndex        =   27
            Top             =   80
            Width           =   900
         End
      End
      Begin VB.CheckBox chkSeqPage 
         Caption         =   "重打“待续打”页"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   4020
         TabIndex        =   22
         Top             =   585
         Visible         =   0   'False
         Width           =   1770
      End
      Begin VB.Image imgIcon 
         Height          =   360
         Index           =   2
         Left            =   2655
         MouseIcon       =   "frmAdvicePrint.frx":058A
         Picture         =   "frmAdvicePrint.frx":0B14
         Top             =   450
         Width           =   360
      End
      Begin VB.Image imgIcon 
         Height          =   360
         Index           =   1
         Left            =   1395
         MouseIcon       =   "frmAdvicePrint.frx":11FE
         Picture         =   "frmAdvicePrint.frx":1788
         Top             =   450
         Width           =   360
      End
      Begin VB.Image imgIcon 
         DragIcon        =   "frmAdvicePrint.frx":1E72
         Height          =   360
         Index           =   0
         Left            =   195
         MouseIcon       =   "frmAdvicePrint.frx":255C
         Picture         =   "frmAdvicePrint.frx":2AE6
         Top             =   450
         Width           =   360
      End
      Begin VB.Label lblPrint 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAdvicePrint.frx":31D0
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   210
         TabIndex        =   12
         Top             =   90
         Width           =   3600
      End
      Begin VB.Label lblStopPrint 
         AutoSize        =   -1  'True
         Caption         =   "提醒："
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   100
         TabIndex        =   11
         Top             =   4500
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.Label lblPrintIcoInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "已打印       待续打        未打印"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   0
         Left            =   600
         TabIndex        =   17
         Top             =   585
         Width           =   2970
      End
   End
   Begin VB.CommandButton cmdPrintSet 
      Caption         =   "打印设置"
      Height          =   315
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   1000
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "刷新(&F)"
      Height          =   315
      Left            =   5355
      TabIndex        =   28
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdClsLastPrint 
      Caption         =   "清除上次打印(&C)"
      Height          =   350
      Left            =   1800
      TabIndex        =   23
      Top             =   5730
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   4950
      ScaleHeight     =   315
      ScaleWidth      =   375
      TabIndex        =   20
      Top             =   30
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.VScrollBar vsc 
      Height          =   1815
      Left            =   6090
      SmallChange     =   50
      TabIndex        =   19
      Top             =   2130
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.PictureBox chkH 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   810
      Left            =   5895
      ScaleHeight     =   810
      ScaleWidth      =   1320
      TabIndex        =   13
      Top             =   -75
      Visible         =   0   'False
      Width           =   1320
      Begin VB.PictureBox picPaper 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   900
         Index           =   0
         Left            =   1245
         ScaleHeight     =   900
         ScaleWidth      =   705
         TabIndex        =   15
         Top             =   525
         Visible         =   0   'False
         Width           =   700
      End
      Begin VB.PictureBox picPaperB 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   900
         Index           =   0
         Left            =   1365
         ScaleHeight     =   900
         ScaleWidth      =   705
         TabIndex        =   14
         Top             =   585
         Visible         =   0   'False
         Width           =   700
      End
      Begin VB.Image imgChk 
         Height          =   240
         Index           =   0
         Left            =   495
         Picture         =   "frmAdvicePrint.frx":321A
         Top             =   420
         Width           =   240
      End
      Begin VB.Image imgIco 
         Height          =   240
         Index           =   0
         Left            =   150
         Picture         =   "frmAdvicePrint.frx":3C1C
         Top             =   465
         Width           =   240
      End
      Begin VB.Label lblNum 
         AutoSize        =   -1  'True
         Caption         =   "1"
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
         Index           =   0
         Left            =   210
         TabIndex        =   16
         Top             =   855
         Visible         =   0   'False
         Width           =   105
      End
   End
   Begin VB.OptionButton optReport 
      Caption         =   "长期医嘱单"
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   1275
   End
   Begin VB.OptionButton optReport 
      Caption         =   "临时医嘱单"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   1
      Left            =   1455
      TabIndex        =   1
      Top             =   120
      Width           =   1275
   End
   Begin VB.ComboBox cboBaby 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3510
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   75
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "退出(&X)"
      Height          =   350
      Left            =   5940
      TabIndex        =   8
      Top             =   5730
      Width           =   800
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "预览(&V)"
      Height          =   350
      Left            =   945
      TabIndex        =   7
      Top             =   5730
      Width           =   800
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "打印(&P)"
      Height          =   350
      Left            =   100
      TabIndex        =   6
      Top             =   5730
      Width           =   800
   End
   Begin MSComctlLib.TabStrip tbsMain 
      Height          =   5200
      Left            =   105
      TabIndex        =   3
      Top             =   400
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   9181
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "常规打印"
            Key             =   "常规打印"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "医嘱套打"
            Key             =   "医嘱套打"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraPrint 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4080
      Index           =   1
      Left            =   150
      TabIndex        =   5
      Tag             =   "医嘱套打"
      Top             =   960
      Visible         =   0   'False
      Width           =   6600
      Begin VB.Image imgIcon 
         DragIcon        =   "frmAdvicePrint.frx":41A6
         Height          =   360
         Index           =   3
         Left            =   210
         MouseIcon       =   "frmAdvicePrint.frx":4890
         Picture         =   "frmAdvicePrint.frx":4E1A
         Top             =   525
         Width           =   360
      End
      Begin VB.Label lblPrintIcoInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "套打"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   600
         TabIndex        =   18
         Top             =   675
         Width           =   360
      End
      Begin VB.Label lblInSidePrint 
         BackStyle       =   0  'Transparent
         Caption         =   "医嘱套打指在已打印的医嘱单上对校对/停止/确认停进行套打。请单击下面的图片选择要套打的医嘱单页号。"
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   210
         TabIndex        =   10
         Top             =   135
         Width           =   4320
      End
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "长期医嘱单：共13页。"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   3630
      TabIndex        =   21
      Top             =   5820
      Width           =   1800
   End
   Begin VB.Label lblBaby 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "医嘱"
      Height          =   180
      Left            =   3090
      TabIndex        =   9
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "frmAdvicePrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'接口参数：
Private mfrmParent As Object
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mstrDefKey As String '缺省定位到的打印功能

'模块变量
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private mrsPrint As ADODB.Recordset
Private mintPrintCount As Integer
Attribute mintPrintCount.VB_VarHelpID = -1
Private mblnTrans As Boolean '用于事件之件的事务嵌套控制
Private mlngPrintType As Long   '医嘱打印模式   1-新开后打印，0-校对后打印

Private mlngRows临嘱 As Long    '在1页纸上打印的总行数
Private mlngRows长嘱 As Long    '
Private mintMid As Integer      '重打页和应打页的临界页，mintMid  （未打页/续打页）
Private mbln续打页 As Boolean   '是否有待续打页
Private mstrTurnPages As String '所有的换页打印的页号，格式 "2,3,6,8,9"
Private mlngPrintedMaxPage As Long ' 已经打印过的医嘱的最大页号
Private mlngPage重整前 As Long '最近一次重整前打印过的最大页号

Private mintPageCount As Integer        '总页数  常规打印的页数，只要进入窗体这个数字是固定的。
Private mintStopPageCount As Integer    '总页数  停嘱打印的页数，只变小，执行套打后会
Private mdat重整时间 As Date
Private mint中药分行长嘱 As Integer
Private mint中药分行临嘱 As Integer

Private Enum mCtlID
    opt医嘱_长嘱 = 0
    opt医嘱_临嘱 = 1
    
    fra界面_连打 = 0
    fra界面_套打 = 1
    
    opt位置_长嘱 = 0
    opt位置_临嘱 = 1
    opt位置_两者 = 2
    
    lbl图标说明_连打 = 0
    lbl图标说明_套打 = 1
    
    img已打 = 0
    img续打 = 1
    img未打 = 2
    img套打 = 3
    
    pic连打_容器 = 1
    pic连打_纸面 = 2
    
    pic套打_容器 = 3
    pic套打_纸面 = 4
    
End Enum

Public Sub ShowMe(frmParent As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long, Optional ByVal strDefKey As String)
    Set mfrmParent = frmParent
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mstrDefKey = strDefKey
    
    On Error Resume Next
    Me.Show 1, frmParent
End Sub

Private Sub cmdPrintSet_Click()
'报表打印设置
    Dim strReport As String
    strReport = IIF(optReport(opt医嘱_长嘱).value, "ZL1_INSIDE_1254_1", "ZL1_INSIDE_1254_2")
    Call mobjReport.ReportPrintSet(gcnOracle, glngSys, strReport, Me)
End Sub

Private Sub cmdRefresh_Click()
'刷新
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lngTmp As Long
    Dim i As Long
    Dim lng期效 As Long
    Dim lng婴儿 As Long
    Dim arrSQL As Variant
    
    On Error GoTo errH
 
    Set mrsPrint = Nothing
    mbln续打页 = False
    mintPageCount = 0
    mintStopPageCount = 0
 
    lng婴儿 = cboBaby.ListCount - 1
    lng期效 = IIF(optReport(opt医嘱_长嘱).value, 0, 1)
    
    arrSQL = Array()
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_病人医嘱打印_Delete(" & mlng病人ID & "," & mlng主页ID & "," & lng婴儿 & "," & lng期效 & ",null,null,3)"
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_病人医嘱打印_Insert(" & mlng病人ID & "," & mlng主页ID & "," & lng婴儿 & "," & lng期效 & "," & IIF(lng期效 = 0, mlngRows长嘱, mlngRows临嘱) & ")"
    
    '提交数据
    If Not mblnTrans Then gcnOracle.BeginTrans: mblnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zldatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    If mblnTrans Then gcnOracle.CommitTrans: mblnTrans = False
    
    '判断是否还要添加纸张
    strSQL = "select max(a.页号) as 页数 from 病人医嘱打印 a where a.病人id=[1] and a.主页id=[2]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
    lngTmp = Val(rsTmp!页数 & "") * 2
    If Val(picPaper(0).Tag) < lngTmp Then
        For i = Val(picPaper(0).Tag) + 1 To lngTmp
            Call LoadPaper(0, i)
            Call LoadPaper(1, i)
        Next
    End If
    Call tbsMain_Click
    Exit Sub
errH:
    If mblnTrans Then gcnOracle.RollbackTrans: mblnTrans = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim arrBaby As Variant, strBaby As String
    Dim blnPriv As Boolean, i As Long
 
    mblnTrans = False
    '设置报表权限
    '长期医嘱单
    blnPriv = False
    If InStr(UserInfo.性质, "医生") > 0 Then
        If InStr(GetInsidePrivs(p住院医嘱下达), "长期医嘱单") > 0 Then blnPriv = True
    End If
    If Not blnPriv And InStr(UserInfo.性质, "护士") > 0 Then
        If InStr(GetInsidePrivs(p住院医嘱发送), "长期医嘱单") > 0 Then blnPriv = True
    End If
    If Not blnPriv Then
        optReport(opt医嘱_临嘱).value = True
        optReport(opt医嘱_长嘱).Enabled = False
    End If
    
    '临时医嘱单
    blnPriv = False
    If InStr(UserInfo.性质, "医生") > 0 Then
        If InStr(GetInsidePrivs(p住院医嘱下达), "临时医嘱单") > 0 Then blnPriv = True
    End If
    If Not blnPriv And InStr(UserInfo.性质, "护士") > 0 Then
        If InStr(GetInsidePrivs(p住院医嘱发送), "临时医嘱单") > 0 Then blnPriv = True
    End If
    If Not blnPriv Then
        optReport(opt医嘱_长嘱).value = True
        optReport(opt医嘱_临嘱).Enabled = False
    End If
    
    '例外情况：两个报表应至少有一个有权限
    If Not optReport(opt医嘱_长嘱).Enabled And Not optReport(opt医嘱_临嘱).Enabled Then
        Unload Me: Exit Sub
    End If
    
    '初始化婴儿选择
    cboBaby.AddItem "病人医嘱"
    Call Cbo.SetIndex(cboBaby.hwnd, 0)
    
    strBaby = GetBabyRegList(mlng病人ID, mlng主页ID)
    
    If strBaby <> "" Then
        arrBaby = Split(strBaby, "<Split>")
        For i = 0 To UBound(arrBaby)
            cboBaby.AddItem "婴儿 " & i + 1 & IIF(arrBaby(i) <> "", "：" & arrBaby(i), "")
        Next
    Else
        lblBaby.Visible = False
        cboBaby.Visible = False
    End If
    Call Cbo.SetListWidth(cboBaby.hwnd, cboBaby.Width * 1.55)
    
    If mobjReport Is Nothing Then Set mobjReport = New clsReport
    
    mdat重整时间 = GetRsRedoDate(mlng病人ID, mlng主页ID)
    
    '医嘱打印模式
    mlngPrintType = Val(zldatabase.GetPara("医嘱单打印模式", glngSys, p住院医嘱下达))
    
    mlngRows临嘱 = GetReportRows(glngSys, "ZL1_INSIDE_1254_2")
    mlngRows长嘱 = GetReportRows(glngSys, "ZL1_INSIDE_1254_1")
    
    mint中药分行长嘱 = Val(zldatabase.GetPara("长嘱单中药医嘱单行显示字数", glngSys, p住院医嘱发送))
    mint中药分行临嘱 = Val(zldatabase.GetPara("临嘱单中药医嘱单行显示字数", glngSys, p住院医嘱发送))
    
    Call Insert打印记录
    
    Call LoadAllPaper
    
    '刷新界面数据
    If mstrDefKey <> "" And tbsMain.SelectedItem.Key <> mstrDefKey Then
        tbsMain.Tag = "NoneClick"
        For i = 1 To tbsMain.Tabs.Count
            If tbsMain.Tabs(i).Key = mstrDefKey Then
                tbsMain.Tabs(i).Selected = True
                Exit For
            End If
        Next
        tbsMain.Tag = ""
    End If
     
    Call tbsMain_Click
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    
    On Error Resume Next
    
    Me.Height = 6600
    
    With tbsMain
        .Left = 100
        .Top = 400
        .Width = Me.ScaleWidth - 200
        .Height = 5200
    End With
    
    cmdPreview.Top = Me.ScaleHeight - 470
    cmdPrint.Top = cmdPreview.Top
    cmdCancel.Top = cmdPreview.Top
    cmdCancel.Left = tbsMain.Width + tbsMain.Left - cmdCancel.Width
    cmdClsLastPrint.Top = cmdPreview.Top
    lblTotal.Top = cmdPreview.Top + 100
    fraClear.Visible = cmdClsLastPrint.Visible
    For i = 0 To 2
        fraPrint(i).Top = 750
        fraPrint(i).Left = 150
        fraPrint(i).Width = tbsMain.Width - 400
        fraPrint(i).Height = tbsMain.Height - 430
    Next
    
    
    For i = 0 To 3
        imgIcon(i).Top = 530
    Next
    
    imgIcon(img套打).Left = imgIcon(img已打).Left
    
    lblPrintIcoInfo(lbl图标说明_连打).Top = 670
    lblPrintIcoInfo(lbl图标说明_套打).Top = 670
    lblStopPrint.Top = 4500
    lblStopPrint.Left = 100
    fraClear.Top = lblStopPrint.Top - 80
    fraClear.Left = tbsMain.Width - fraClear.Width - 350
 
    lblPrint.Left = lblInSidePrint.Left
    lblInSidePrint.Top = lblPrint.Top
    cmdRefresh.Top = cboBaby.Top
    cmdRefresh.Left = cmdCancel.Left + cmdCancel.Width - cmdRefresh.Width
    
    cmdPrintSet.Top = cmdRefresh.Top
    cmdPrintSet.Left = cmdRefresh.Left - cmdPrintSet.Width - 20
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnTrans Then gcnOracle.RollbackTrans: mblnTrans = False '以防万一
    Set mrsPrint = Nothing
    mbln续打页 = False
    Call UnLoadPaper
    mintPageCount = 0
    mintStopPageCount = 0
End Sub

Private Sub cboBaby_Click()
    Call RefreshFace
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPreview_Click()
    Call AdvicePrint(1)
    Call RefreshFace
End Sub

Private Sub cmdPrint_Click()
    Call AdvicePrint(2)
    Call RefreshFace
End Sub

Private Sub optReport_Click(Index As Integer)
    If Not Visible Then Exit Sub
    Call tbsMain_Click
End Sub

Private Sub tbsMain_Click()
    Dim i As Long
    
    If tbsMain.Tag = "NoneClick" Then Exit Sub
    
    For i = 0 To fraPrint.UBound
        fraPrint(i).Visible = fraPrint(i).Tag = tbsMain.SelectedItem.Key
        If fraPrint(i).Tag = tbsMain.SelectedItem.Key Then
            fraPrint(i).ZOrder
            If i = fra界面_连打 Then picContainer(pic连打_容器).ZOrder
            If i = fra界面_套打 Then picContainer(pic套打_容器).ZOrder
        End If
    Next
    Call RefreshFace
    picContainer(pic套打_纸面).Top = 0
    picContainer(pic连打_纸面).Top = 0
    vsc.value = 0
    cmdRefresh.Enabled = tbsMain.SelectedItem.Key = "常规打印"
End Sub

Private Sub RefreshFace()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim i As Long, j As Long
    Dim intIco As Integer '0 - 已打印 1 - 待续打 2 － 待打印
    Dim intIndex As Integer
    Dim lngTmp As Long
    
    On Error GoTo errH
    
    If tbsMain.SelectedItem.Key = "常规打印" Then
        strSQL = "select m.页号,sum(m.打印) as 打印,sum(m.未打印) as 未打印,count(1) as 行数" & vbNewLine & _
            "from (select a.页号,decode(a.打印时间,null,0,1) as 打印,decode(a.打印时间,null,1,0) as 未打印" & vbNewLine & _
            "from 病人医嘱打印 a where a.病人id=[1] and a.主页id=[2] and nvl(a.婴儿,0)=[3] and a.期效=[4] and 行号>0) m" & vbNewLine & _
            "group by m.页号 order by m.页号"
        Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, cboBaby.ListIndex, IIF(optReport(opt医嘱_长嘱).value, 0, 1))
        
        mintMid = 0
        mlngPrintedMaxPage = 0
        mstrTurnPages = ""
        mbln续打页 = False
        mintPageCount = rsTmp.RecordCount
        chkSeqPage.Visible = False
        For i = 1 To rsTmp.RecordCount

            If Val(rsTmp!打印 & "") > 0 And Val(rsTmp!未打印 & "") = 0 Then
                intIco = 0
                mlngPrintedMaxPage = Val(rsTmp!页号 & "")
            ElseIf Val(rsTmp!打印 & "") > 0 And Val(rsTmp!未打印 & "") > 0 Then
                If 0 = mintMid Then
                    mintMid = Val(rsTmp!页号 & "")
                    mbln续打页 = True
                    chkSeqPage.Visible = True
                End If
                intIco = 1
                mlngPrintedMaxPage = Val(rsTmp!页号 & "")
            ElseIf Val(rsTmp!打印 & "") = 0 And Val(rsTmp!未打印 & "") > 0 Then
                If 0 = mintMid Then mintMid = Val(rsTmp!页号 & "")
                intIco = 2
            End If
            
            '最后一页不算做换打页
            If i <> rsTmp.RecordCount Then
                If Val(rsTmp!行数 & "") < IIF(optReport(0).value, mlngRows长嘱, mlngRows临嘱) Then
                    mstrTurnPages = mstrTurnPages & "," & rsTmp!页号
                End If
            End If
      
            Set imgIco(i).Picture = imgIcon(intIco).Picture
            imgIco(i).ToolTipText = IIF(intIco = 0, "已打印", IIF(intIco = 1, "待续打", "未打印"))
            imgChk(i).Visible = IIF(intIco = 0, False, True)
            picPaper(i).Visible = True
            picPaperB(i).Visible = True
            rsTmp.MoveNext
        Next
        
        '判断是否应该显示清除上次打印按钮
        rsTmp.Filter = "打印>0"
        cmdClsLastPrint.Visible = Not rsTmp.EOF
        For i = mintPageCount + 1 To Val(picPaper(0).Tag)
            imgChk(i).Visible = False
            picPaper(i).Visible = False
            picPaperB(i).Visible = False
        Next
        
        If mstrTurnPages <> "" Then mstrTurnPages = Mid(mstrTurnPages, 2)
        
        Set rsTmp = GetStopedAdvice(True)
        
        If mlngPrintType = 1 Then
            If optReport(opt医嘱_长嘱).value Then
                lblStopPrint.Caption = "有校对/停止/确认停止的医嘱需要打印。"
            Else
                lblStopPrint.Caption = "有校对的医嘱需要打印。"
            End If
        Else
            lblStopPrint.Caption = "有确认停止的医嘱需要打印。"
        End If
            
        lblStopPrint.Visible = rsTmp.RecordCount > 0
         
        cmdPreview.Enabled = mintPageCount > 0
        cmdPrint.Enabled = mintPageCount > 0
        
        lngTmp = IntEx(mintPageCount / 21)
        If lngTmp = 0 Then lngTmp = 1
 
        picContainer(pic连打_纸面).Height = lngTmp * 3450
        vsc.Visible = lngTmp > 1
        If lngTmp > 1 Then
            vsc.Max = (lngTmp - 1) * 3450 / Screen.TwipsPerPixelY
        End If
        
        If mintPageCount = 0 Then
            lblTotal.Caption = IIF(optReport(opt医嘱_长嘱).value, "长期", "临时") & "医嘱单：无。"
        Else
            lblTotal.Caption = IIF(optReport(opt医嘱_长嘱).value, "长期", "临时") & "医嘱单：共" & mintPageCount & "页。"
        End If
        lblTotal.Visible = True
        If mlngPrintedMaxPage <> 0 Then txtClearPage.Text = mlngPrintedMaxPage
    ElseIf tbsMain.SelectedItem.Key = "医嘱套打" Then
        
        Set rsTmp = GetStopedAdvice(False)
        
        mintStopPageCount = rsTmp.RecordCount
        
        For i = 1 To rsTmp.RecordCount
            lblNum(i + 1000).Caption = Val(rsTmp!页号 & "")
            lblNum(i + 1000).ToolTipText = "第" & Val(rsTmp!页号 & "") & "页"
            picPaper(i + 1000).Visible = True
            picPaperB(i + 1000).Visible = True
            imgChk(i + 1000).Visible = False
            rsTmp.MoveNext
        Next
        
        For i = mintStopPageCount + 1 To Val(picPaperB(0).Tag)
            imgChk(i + 1000).Visible = False
            picPaper(i + 1000).Visible = False
            picPaperB(i + 1000).Visible = False
        Next
        
        cmdPrint.Enabled = mintStopPageCount <> 0
        cmdPreview.Enabled = mintStopPageCount <> 0
        
        lngTmp = IntEx(mintStopPageCount / 21)
        If lngTmp = 0 Then lngTmp = 1
        picContainer(pic套打_纸面).Height = lngTmp * 3450
        
        vsc.Visible = lngTmp > 1
        If lngTmp > 1 Then
            vsc.Max = (lngTmp - 1) * 3450 / Screen.TwipsPerPixelY
        End If
        If mintStopPageCount = 0 Then
            lblTotal.Caption = "套打医嘱单：无。"
        Else
            lblTotal.Caption = "套打医嘱单：共" & mintStopPageCount & "页。"
        End If
        lblTotal.Visible = True
        cmdClsLastPrint.Visible = False
    ElseIf tbsMain.SelectedItem.Key = "打印选项" Then
        cmdPrint.Enabled = False
        cmdPreview.Enabled = False
        vsc.Visible = False
        lblTotal.Visible = False
        cmdClsLastPrint.Visible = False
    End If
    fraClear.Visible = cmdClsLastPrint.Visible
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function GetStopedAdvice(ByVal blnOnlyCheckExists As Boolean) As ADODB.Recordset
'功能：获取当前病人需要停嘱打印的记录集
'参数：blnOnlyCheckExists-只检查是否存在停嘱打印
    Dim strSQL As String
    
    If optReport(opt医嘱_长嘱).value Then
        If blnOnlyCheckExists Then
            strSQL = "Select 1 From 病人医嘱打印 A, 病人医嘱记录 B Where A.医嘱id = B.ID And A.期效 = 0 And A.病人id = [1] And A.主页id = [2] And Nvl(A.婴儿, 0) = [3] And a.打印时间 is not null and (B.确认停嘱时间 Is Not Null And" & vbNewLine & _
                "     Not Exists (Select 1 From 病人医嘱打印 S Where S.医嘱id = A.医嘱id And S.打印标记 = 2) " & _
                IIF(mlngPrintType = 1, "Or B.执行终止时间 Is Not Null And Not exists (Select 1 From 病人医嘱打印 S Where S.医嘱id = A.医嘱id And S.打印标记 in (1,2))  or b.校对时间 is not null and not exists (Select 1 From 病人医嘱打印 S Where S.医嘱id = A.医嘱id And S.打印标记 in(1,2,3))", "") & ") And Rownum<2"
        Else
            strSQL = _
                "Select Distinct 页号" & vbNewLine & _
                "From (Select A.医嘱id, Max(A.页号) As 页号" & vbNewLine & _
                "       From 病人医嘱打印 A, 病人医嘱记录 B" & vbNewLine & _
                "       Where A.医嘱id = B.ID And A.期效 = 0 And A.病人id = [1] And A.主页id = [2] And Nvl(A.婴儿, 0) = [3] And a.打印时间 is not null And (B.确认停嘱时间 Is Not Null And" & vbNewLine & _
                "             Not Exists (Select 1 From 病人医嘱打印 S Where S.医嘱id = A.医嘱id And S.打印标记 = 2) " & _
                IIF(mlngPrintType = 1, "Or B.执行终止时间 Is Not Null And Not exists(Select 1 From 病人医嘱打印 S Where S.医嘱id = A.医嘱id And S.打印标记  in (1,2))  or b.校对时间 is not null and not exists (Select 1 From 病人医嘱打印 S Where S.医嘱id = A.医嘱id And S.打印标记 in(1,2,3)) ", "") & ")" & vbNewLine & _
                "       Group By A.医嘱id)" & vbNewLine & _
                "Order By 页号"
        End If
    
    Else
        If blnOnlyCheckExists Then
            strSQL = "Select 1 From 病人医嘱打印 A, 病人医嘱记录 B Where A.医嘱id = B.ID And A.期效 = 1 And A.病人id = [1] And A.主页id = [2] And Nvl(A.婴儿, 0) = [3] And a.打印时间 is not null " & vbNewLine & _
                IIF(mlngPrintType = 1, " and b.校对时间 is not null and not exists (Select 1 From 病人医嘱打印 S Where S.医嘱id = A.医嘱id And S.打印标记 in(1,2,3))", " and 1=0") & " And Rownum<2"
        Else
            strSQL = _
                "Select Distinct 页号" & vbNewLine & _
                "From (Select A.医嘱id, Max(A.页号) As 页号" & vbNewLine & _
                "       From 病人医嘱打印 A, 病人医嘱记录 B" & vbNewLine & _
                "       Where A.医嘱id = B.ID And A.期效 = 1 And A.病人id = [1] And A.主页id = [2] And Nvl(A.婴儿, 0) = [3] And a.打印时间 is not null " & vbNewLine & _
                IIF(mlngPrintType = 1, " and b.校对时间 is not null and not exists (Select 1 From 病人医嘱打印 S Where S.医嘱id = A.医嘱id And S.打印标记 in(1,2,3)) ", " and 1=0") & vbNewLine & _
                "       Group By A.医嘱id)" & vbNewLine & _
                "Order By 页号"
        End If
        
        If mlngPrintType = 0 Then
            strSQL = "Select 1 as 页号 From dual where 0=1"
        End If
        
    End If
    On Error GoTo errH
    
    Set GetStopedAdvice = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, cboBaby.ListIndex)

    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub AdvicePrint(ByVal intMode As Integer)
'功能：执行医嘱单打印或以预览
'参数：intMode=1-预览,2-打印
    Dim lngBegin As Long, lngEnd As Long
    Dim lng行号 As Long, strReport As String
    Dim colSegment As Collection
    Dim col常规打印 As Collection
    Dim strSQL As String, i As Long, j As Long
    Dim rsTmp As ADODB.Recordset
    Dim intIndex As Long
    Dim varArr As Variant
    Dim int中药分行 As Integer
    
    '确定具体的报表编号
    strReport = IIF(optReport(opt医嘱_长嘱).value, "ZL1_INSIDE_1254_1", "ZL1_INSIDE_1254_2")
    
    If optReport(opt医嘱_长嘱).value Then
        int中药分行 = IIF(mint中药分行长嘱 = 0, 0, 1)
    Else
        int中药分行 = IIF(mint中药分行临嘱 = 0, 0, 1)
    End If
    
    On Error GoTo errH
    
    If tbsMain.SelectedItem.Key = "常规打印" Then '医嘱续打
        '只有在打印过的医嘱界面才能进行跳选，未打的只能连续选择
        '根据选择情况自动对页号分段
        Set colSegment = New Collection
        lngBegin = 0: lngEnd = 0
        
        For i = 1 To mintPageCount
            If imgChk(i).Visible Then
                If lngBegin = 0 Then
                    lngBegin = i: lngEnd = i
                ElseIf i = lngEnd + 1 Then
                    lngEnd = i
                Else
                    colSegment.Add lngBegin & "-" & lngEnd
                    lngBegin = i: lngEnd = i
                End If
            End If
        Next
        
        If lngBegin <> 0 And lngEnd <> 0 Then
            colSegment.Add lngBegin & "-" & lngEnd
        End If
        
        If colSegment.Count = 0 Then
            MsgBox "请选择需要打印的医嘱单页号范围。", vbInformation, gstrSysName
            Exit Sub
        ElseIf intMode = 1 And colSegment.Count > 1 Then
            MsgBox "请一次只选择一个或连续的一段页号范围进行预览。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '换页处理，可能存在换页打印的情况，再生成一次打印段
        If mstrTurnPages <> "" Then
            Set col常规打印 = New Collection
            For i = 1 To colSegment.Count '分段调用打印
            
                lngBegin = Split(colSegment(i), "-")(0)
                lngEnd = Split(colSegment(i), "-")(1)
                
                varArr = Split(mstrTurnPages, ",")
                For j = 0 To UBound(varArr)
                    If lngBegin <= Val(varArr(j)) And Val(varArr(j)) <= lngEnd Then
                        col常规打印.Add lngBegin & "-" & Val(varArr(j))
                        lngBegin = Val(varArr(j)) + 1
                    End If
                Next
                
                If lngBegin <= lngEnd Then col常规打印.Add lngBegin & "-" & lngEnd
            Next
            Set colSegment = col常规打印
        End If
        
        For i = 1 To colSegment.Count '分段调用打印
        
            mintPrintCount = 0 '用于防止预览时多次重复打印
            
            lng行号 = 0
            lngBegin = Split(colSegment(i), "-")(0)
            lngEnd = Split(colSegment(i), "-")(1)
            
            '续打处理，只会处理一次
            If mintMid = lngBegin Then
                If mbln续打页 Then '续打页，本次按续打规则处理，计算行号
                    strSQL = "select max(行号)+1 as 行号 from 病人医嘱打印 where 打印时间 is not null and 病人id=[1] and 主页id=[2] and nvl(婴儿,0)=[3] and 期效=[4] and 页号=[5]"
                    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, cboBaby.ListIndex, IIF(optReport(0).value, 0, 1), lngBegin)
                    If Not rsTmp.EOF Then
                        lng行号 = Val(rsTmp!行号 & "")
                        If chkSeqPage.value = 1 And chkSeqPage.Visible Then lng行号 = 0
                    End If
                End If
            End If
            Call mobjReport.ReportOpen(gcnOracle, glngSys, strReport, mfrmParent, _
                "病人ID=" & mlng病人ID, "主页ID=" & mlng主页ID, "婴儿=" & cboBaby.ListIndex, "打印模式=" & mlngPrintType, "停嘱打印=0", "起始行号=" & lng行号, _
                "StartPageNum=" & lngBegin, "起始页号=" & lngBegin, "结束页号=" & lngEnd, "中药分行=" & int中药分行, "PressWorkFirst=" & IIF(lng行号 <> 0, 1, 0), intMode)
        Next
    ElseIf tbsMain.SelectedItem.Key = "医嘱套打" Then
        '根据选择情况自动对页号分段
        Set colSegment = New Collection
        lngBegin = 0: lngEnd = 0
    
        For i = 1 To mintStopPageCount
            intIndex = 1000 + i
            If imgChk(intIndex).Visible Then
                If lngBegin = 0 Then
                    lngBegin = Val(lblNum(intIndex).Caption)
                    lngEnd = lngBegin
                ElseIf Val(lblNum(intIndex).Caption) = lngEnd + 1 Then
                    lngEnd = Val(lblNum(intIndex).Caption)
                Else
                    colSegment.Add lngBegin & "-" & lngEnd
                    lngBegin = Val(lblNum(intIndex).Caption)
                    lngEnd = lngBegin
                End If
            End If
        Next
        
        If lngBegin <> 0 Then colSegment.Add lngBegin & "-" & lngEnd

        If colSegment.Count = 0 Then
            MsgBox "请选择需要套打的医嘱单页号范围。", vbInformation, gstrSysName
            Exit Sub
        ElseIf intMode = 1 And colSegment.Count > 1 Then
            MsgBox "请一次只选择一个或连续的一段页号范围进行预览。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        For i = 1 To colSegment.Count '分页号段调用套打
            mintPrintCount = 0 '用于防止预览时多次重复打印
            lngBegin = Split(colSegment(i), "-")(0): lngEnd = Split(colSegment(i), "-")(1)
            Call mobjReport.ReportOpen(gcnOracle, glngSys, strReport, mfrmParent, _
                "病人ID=" & mlng病人ID, "主页ID=" & mlng主页ID, "婴儿=" & cboBaby.ListIndex, "打印模式=" & mlngPrintType, "停嘱打印=1", "起始行号=1", _
                "StartPageNum=" & lngBegin, "起始页号=" & lngBegin, "结束页号=" & lngEnd, "中药分行=" & int中药分行, "PressWork=1", intMode)
        Next
    End If
    Exit Sub
errH:
    If mblnTrans Then gcnOracle.RollbackTrans: mblnTrans = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mobjReport_BeforePrint(ByVal ReportNum As String, ByVal TotalPages As Integer, Cancel As Boolean, arrBill As Variant)
'功能：开始打印事件，初始化医嘱打印信息记录集
    
    If tbsMain.SelectedItem.Key = "常规打印" Then
        '预览时多次重复打印检查
        If mintPrintCount > 0 Then
            MsgBox "已经打印过了，要想重新打印，请使用重打功能。", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
        
        '预览时部份打印检查
        If TotalPages < 0 Then
            MsgBox "为保证有效进行续打，请选择对全部页面进行打印。", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
        
        mintPrintCount = mintPrintCount + 1
        
        Set mrsPrint = New ADODB.Recordset
        mrsPrint.Fields.Append "医嘱ID", adBigInt
        mrsPrint.Fields.Append "页号", adBigInt
        mrsPrint.Fields.Append "行号", adBigInt
        mrsPrint.CursorLocation = adUseClient
        mrsPrint.LockType = adLockOptimistic
        mrsPrint.CursorType = adOpenStatic
        mrsPrint.Open
    ElseIf tbsMain.SelectedItem.Key = "医嘱套打" Then
        '预览时多次重复打印检查
        If mintPrintCount > 0 Then
            MsgBox "已经打印过了，不能重复打印。", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
        
        '预览时部份打印检查
        If TotalPages < 0 Then
            MsgBox "为保证有效进行套打，请选择对全部页面进行打印。", vbInformation, gstrSysName
            Cancel = True: Exit Sub
        End If
        
        mintPrintCount = mintPrintCount + 1
        
        Set mrsPrint = New ADODB.Recordset
        mrsPrint.Fields.Append "医嘱ID", adBigInt
        mrsPrint.Fields.Append "页号", adBigInt
        mrsPrint.Fields.Append "行号", adBigInt
        mrsPrint.CursorLocation = adUseClient
        mrsPrint.LockType = adLockOptimistic
        mrsPrint.CursorType = adOpenStatic
        mrsPrint.Open
    End If
End Sub

Private Sub mobjReport_AfterPrint(ByVal ReportNum As String)
'功能：结束打印事件，写入病人医嘱打印数据
    Dim curDate As Date, strSQL As String
    
    If tbsMain.SelectedItem.Key = "常规打印" Then
        '产生医嘱打印位置记录
        curDate = zldatabase.Currentdate
        mrsPrint.Filter = 0
        If Not mrsPrint.EOF Then
            On Error GoTo errH
            
            If Not mblnTrans Then gcnOracle.BeginTrans: mblnTrans = True
            Do While Not mrsPrint.EOF
                strSQL = "zl_病人医嘱打印_Update(" & ZVal(mrsPrint!医嘱ID) & "," & mrsPrint!页号 & "," & mrsPrint!行号 & "," & _
                    mlng病人ID & "," & mlng主页ID & "," & cboBaby.ListIndex & "," & IIF(optReport(0).value, 0, 1) & "," & _
                    "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),'" & UserInfo.姓名 & "'," & mlngPrintType & ")"
                Call zldatabase.ExecuteProcedure(strSQL, Me.Caption)
                mrsPrint.MoveNext
                
            Loop
            If mblnTrans Then gcnOracle.CommitTrans: mblnTrans = False
        End If
    ElseIf tbsMain.SelectedItem.Key = "医嘱套打" Then
        '标记医嘱停嘱时间已套打标志
        mrsPrint.Filter = 0
        If Not mrsPrint.EOF Then
            On Error GoTo errH
            If Not mblnTrans Then gcnOracle.BeginTrans: mblnTrans = True
            Do While Not mrsPrint.EOF
                strSQL = "Zl_病人医嘱打印_Update(" & mrsPrint!医嘱ID & "," & mrsPrint!页号 & "," & mrsPrint!行号 & ",null,null,null,null,null,null," & mlngPrintType & ",1)"
                Call zldatabase.ExecuteProcedure(strSQL, Me.Caption)
                mrsPrint.MoveNext
            Loop
            If mblnTrans Then gcnOracle.CommitTrans: mblnTrans = False
        End If
    End If
 
    Set mrsPrint = Nothing
    Exit Sub
errH:
    If mblnTrans Then gcnOracle.RollbackTrans: mblnTrans = False
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mobjReport_PrintSheetRow(ByVal ReportNum As String, Sheet As Object, ByVal Page As Integer, ByVal Row As Long, ByVal ID As Long)
'功能：报表数据打印事件，记录医嘱打印行数据
'说明：当表格行无数据要打印时，是不会激活该事件的
    If tbsMain.SelectedItem.Key = "常规打印" Then
        If Page >= 1 And Row >= 1 Then
            'mrsPrint.Filter = "医嘱ID=" & ID 'NULL会返回为0
            mrsPrint.Filter = "医嘱ID=" & ID & " and 页号 =" & Page & " and 行号=" & Row
            If mrsPrint.EOF Then
                mrsPrint.AddNew
                mrsPrint!医嘱ID = ID
                mrsPrint!页号 = Page
                mrsPrint!行号 = Row
            End If
            mrsPrint.Update
        End If
    ElseIf tbsMain.SelectedItem.Key = "医嘱套打" Then
        If ID > 0 And Page >= 1 And Row >= 1 Then
            mrsPrint.Filter = "医嘱ID=" & ID
            If mrsPrint.EOF Then
                mrsPrint.AddNew
                mrsPrint!医嘱ID = ID
                mrsPrint!页号 = Page
                mrsPrint!行号 = Row
                mrsPrint.Update
            End If
        End If
    End If
End Sub

Private Function GetReportRows(ByVal lngSys As Long, ByVal strReport As String, Optional ByVal intFormat As Integer = 1) As Long
'功能：获取指定报表中主要任意表格的可打印数据行数
'参数：lngSys=系统编号，为0表示共享报表
'      strReport=报表编号
'      intFormat=报表格式号,缺省为1
'返回：0表示没有任意表格
'说明：
'  1.如果报表中存在多个任意表格，则以最大的一个作为主要表格。
'  2.如果表格分栏，则可打印行数是指分栏之后的总行数。
    Dim rsTable As ADODB.Recordset
    Dim rsColumn As ADODB.Recordset
    Dim strSQL As String, i As Long, j
    Dim blnHead As Boolean, blnBody As Boolean
    Dim lngBodyH As Long, lngHeadH As Long
    
    On Error GoTo errH
    
    strSQL = "Select A.ID as 报表ID,B.ID,B.W,B.H,B.行高,B.分栏" & _
        " From zlReports A,zlRPTItems B" & _
        " Where A.ID=B.报表ID And B.类型=4 And Nvl(A.系统,0)=[1] And A.编号=[2] And B.格式号=[3]" & _
        " Order by B.W*B.H Desc"
    Set rsTable = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lngSys, strReport, intFormat)
    If rsTable.EOF Then Exit Function
    
    strSQL = "Select 序号,表头,内容 From zlRPTItems Where 报表ID=[1] And 格式号=[2] And 上级ID=[3] And 类型=6 Order by 序号"
    Set rsColumn = zldatabase.OpenSQLRecord(strSQL, Me.Caption, Val(rsTable!报表ID), intFormat, Val(rsTable!ID))
    If rsColumn.EOF Then Exit Function
    
    '以下代码参照自定义报表中的方法编写
    '----------------------------------
    '求出表头高度:以第一列为准
    For i = 0 To UBound(Split(rsColumn!表头, "|"))
        lngHeadH = lngHeadH + Val(Split(Split(rsColumn!表头, "|")(i), "^")(1))
    Next
    
    '求出表体高度
    blnHead = False: blnBody = False
    rsColumn.MoveFirst
    Do While Not rsColumn.EOF
        i = UBound(Split(rsColumn!表头, "|"))
        If i > 0 Then
            blnHead = True
        ElseIf i = 0 Then
            blnHead = blnHead Or (Split(Split(rsColumn!表头, "|")(i), "^")(2) <> "#")
        End If
        blnBody = blnBody Or Not IsNull(rsColumn!内容)
        rsColumn.MoveNext
    Loop
    If Not blnHead And blnBody Then '仅有表体
        lngBodyH = rsTable!H
    Else
        If rsTable!H - lngHeadH + 15 < 0 Then
            lngBodyH = 0
        Else
            lngBodyH = rsTable!H - lngHeadH + 15
        End If
    End If
    
    '求出行数
    GetReportRows = Int(lngBodyH / rsTable!行高) * NVL(rsTable!分栏, 1)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Insert打印记录()
'功能：生成将要打印的医嘱记录和要进行停嘱打印的医嘱。临嘱/长嘱，病人医嘱和婴儿是分开的，单独产生。
    Dim arrSQL As Variant
    Dim lngRows As Long
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo errH
    
    arrSQL = Array()
    
    '病人和婴儿都要判断，在求医嘱最后序号时要考虑重整的情况，临时医嘱单不考虑转换页和打印重整标记的情况
    '判断是否要生成常规打印的记录，两层循环，j 表示期效，i 表示婴儿序号i=0时表示病人
    For j = 0 To 1
        lngRows = IIF(j = 0, mlngRows长嘱, mlngRows临嘱)
        For i = 0 To cboBaby.ListCount - 1
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病人医嘱打印_Delete(" & mlng病人ID & "," & mlng主页ID & "," & i & "," & j & ",null,null,3)"
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_病人医嘱打印_Insert(" & mlng病人ID & "," & mlng主页ID & "," & i & "," & j & "," & lngRows & ")"
        Next
    Next
    
    '提交数据
    If Not mblnTrans Then gcnOracle.BeginTrans: mblnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zldatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    If mblnTrans Then gcnOracle.CommitTrans: mblnTrans = False
        
    Exit Sub
errH:
    If mblnTrans Then gcnOracle.RollbackTrans: mblnTrans = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
 
Private Sub LoadAllPaper()
'功能：加载容器，所有图片纸张
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim intTmp As Integer, i As Integer
    
    On Error GoTo errH
    
    For i = 1 To 4
        Load picContainer(i)
        picContainer(i).Width = 6160
        picContainer(i).Height = 3400
    Next

    Set picContainer(1).Container = Me
    Set picContainer(3).Container = Me

    Set picContainer(2).Container = picContainer(1)
    Set picContainer(4).Container = picContainer(3)
    
    picContainer(1).Top = 1720
    picContainer(1).Left = 350
    picContainer(3).Top = 1720
    picContainer(3).Left = 350
    
    picContainer(2).Top = 0
    picContainer(4).Top = 0
    picContainer(2).Left = 0
    picContainer(4).Left = 0
    
    For i = 1 To 4
        picContainer(i).Visible = True
        picContainer(i).ZOrder 0
    Next
    
    vsc.Left = 6520
    vsc.Height = 3300
    vsc.Top = 1820
    vsc.Width = 200
    vsc.ZOrder 0

    strSQL = "select max(a.页号) as 页数 from 病人医嘱打印 a where a.病人id=[1] and a.主页id=[2]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
 
    intTmp = Val(rsTmp!页数 & "") * 2
    picPaper(0).Tag = intTmp
    picPaperB(0).Tag = intTmp
    
    For i = 1 To intTmp
        Call LoadPaper(0, i)
        Call LoadPaper(1, i)
    Next
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadPaper(ByVal intCt As Integer, ByVal intNum As Integer)
'功能：加载图纸张，目前支持最多页数 999页
'参数：intCt容器，0－连续打印fraPrint(0)，2－停嘱打印fraPrint(1)；intNum 页号
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim intRow As Integer
    Dim intIndex As Integer
    
    On Error GoTo errH
    
    intIndex = intNum + 1000 * intCt
    
    intRow = 1 + (intNum - 1) \ 7
    
    lngLeft = ((intNum - 1) Mod 7) * (picPaper(0).Width + 200)
    lngTop = (intRow - 1) * (picPaper(0).Height + 250)
    
    '主图片
    Load picPaper(intIndex)
    Load picPaperB(intIndex)
 
    '背景图和容器图片
    Set picPaperB(intIndex).Container = picContainer(2 + intCt * 2)
    Set picPaper(intIndex).Container = picContainer(2 + intCt * 2)
 
    picPaper(intIndex).Left = lngLeft
    picPaper(intIndex).Top = lngTop
    picPaper(intIndex).Width = picPaper(0).Width
    picPaper(intIndex).Height = picPaper(0).Height
    picPaper(intIndex).BackColor = picPaper(0).BackColor
    picPaper(intIndex).Visible = False
    picPaper(intIndex).ZOrder 0

    picPaperB(intIndex).Left = picPaper(intIndex).Left + 50
    picPaperB(intIndex).Top = picPaper(intIndex).Top + 50
    picPaperB(intIndex).Width = picPaper(0).Width
    picPaperB(intIndex).Height = picPaper(0).Height
    picPaperB(intIndex).BackColor = picPaperB(0).BackColor
    picPaperB(intIndex).Visible = False
    
    '纸上的图标
    Load imgIco(intIndex)
    Set imgIco(intIndex).Container = picPaper(intIndex)
    Set imgIco(intIndex).Picture = imgIcon(0).Picture
    imgIco(intIndex).Left = (picPaper(intIndex).Width - imgIco(intIndex).Width) / 2
    imgIco(intIndex).Top = 260
    imgIco(intIndex).Visible = True
    imgIco(intIndex).ZOrder 1
    
    Load lblNum(intIndex)
    Set lblNum(intIndex).Container = picPaper(intIndex)
    lblNum(intIndex).Visible = True
    lblNum(intIndex).Caption = intNum
    lblNum(intIndex).ToolTipText = "第" & intNum & "页"
    lblNum(intIndex).FontSize = lblNum(0).FontSize
    lblNum(intIndex).Left = (picPaper(intIndex).Width - lblNum(intIndex).Width) / 2
    lblNum(intIndex).Top = imgIco(intIndex).Height + imgIco(intIndex).Top + 10
    lblNum(intIndex).BackColor = picPaper(0).BackColor
    
    '勾选图片，程序中控件可见性
    Load imgChk(intIndex)
    Set imgChk(intIndex).Container = picPaper(intIndex)
    Set imgChk(intIndex).Picture = imgChk(0).Picture '固定
    imgChk(intIndex).Width = 240
    imgChk(intIndex).Height = 240
    imgChk(intIndex).Left = picPaper(0).Width - imgChk(intIndex).Width
    imgChk(intIndex).Top = -10
    imgChk(intIndex).Visible = False
    imgChk(intIndex).ZOrder 1
    
    Exit Sub
errH:
    If 1 = 2 Then
        Resume
    End If
    err.Clear
End Sub
   
Private Sub UnLoadPaper()
    Dim i As Integer
    
    On Error Resume Next
    
    '先卸载容器内的制件再卸载容器
    For i = 1 To Val(picPaper(0).Tag)
        Unload imgChk(i)
        Unload imgIco(i)
        Unload lblNum(i)
        Unload picPaperB(i)
        Unload picPaper(i)
    Next
    
    For i = 1 To Val(picPaperB(0).Tag)
        Unload imgChk(i + 1000)
        Unload imgIco(i + 1000)
        Unload lblNum(i + 1000)
        Unload picPaperB(i + 1000)
        Unload picPaper(i + 1000)
    Next
    
    For i = 1 To 4
        Unload picContainer(i)
    Next
    
    err.Clear
End Sub

Private Sub imgIco_Click(Index As Integer)
    Call picPaper_Click(Index)
End Sub

Private Sub lblNum_Click(Index As Integer)
    Call picPaper_Click(Index)
End Sub

Private Sub imgChk_Click(Index As Integer)
    Call picPaper_Click(Index)
End Sub

Private Sub picPaper_Click(Index As Integer)
    Dim blnTmp As Boolean
    Dim i As Integer
    
    blnTmp = imgChk(Index).Visible
    imgChk(Index).Visible = Not blnTmp
    
    If Not (Index > 1000 Or mintMid = 0 Or mintMid > Index) Then
        If blnTmp Then
            For i = Index + 1 To mintPageCount
                imgChk(i).Visible = imgChk(Index).Visible
            Next
        Else
            For i = mintMid To Index - 1
                imgChk(i).Visible = imgChk(Index).Visible
            Next
        End If
    End If
    
    If mbln续打页 And Index < 1000 Then
        If mintMid = 1 Then
            chkSeqPage.Visible = imgChk(mintMid).Visible
        Else
            chkSeqPage.Visible = imgChk(mintMid).Visible And Not imgChk(mintMid - 1).Visible
        End If
    End If
End Sub

Private Sub vsc_Change()
    Call vsc_Scroll
End Sub

Private Sub vsc_Scroll()
    If tbsMain.SelectedItem.Key = "常规打印" Then
        picContainer(pic连打_纸面).Top = (-1) * vsc.value * Screen.TwipsPerPixelY
    Else
        picContainer(pic套打_纸面).Top = (-1) * vsc.value * Screen.TwipsPerPixelY
    End If
End Sub

Private Sub cmdClsLastPrint_Click()
'功能：清除打印记录
    Call ClearPrintRs(True)
End Sub

Private Sub txtClearPage_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call txtClearPage_Validate(False)
        cmdClear.SetFocus
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtClearPage_Validate(Cancel As Boolean)
    Dim lngTmp As Long
    lngTmp = Val(txtClearPage.Text)
    If lngTmp = 0 Then
        txtClearPage.Text = 1
    ElseIf lngTmp > mlngPrintedMaxPage Then
        txtClearPage.Text = mlngPrintedMaxPage
    Else
        txtClearPage.Text = lngTmp
    End If
End Sub

Private Sub cmdClear_Click()
    Call ClearPrintRs(False)
End Sub

Private Sub ClearPrintRs(ByVal bln上次打印 As Boolean)
'功能：从某页开始清除打印
'参数：bln上次打印 －true 清除上次打印，false 清除指定页
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim rsTmpOther As ADODB.Recordset
    Dim str位置 As String, str打印人 As String
    Dim lngTmp As Long, strTmp As String, str打印时间 As String
    Dim arrSQL As Variant
    Dim lngRows As Long
    Dim i As Long
    Dim lng页号 As Long
    
    If MsgBox("确实要清除已打印的医嘱记录吗，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
    
    arrSQL = Array()
    
    On Error GoTo errH
    
    lngRows = IIF(optReport(opt医嘱_长嘱).value, mlngRows长嘱, mlngRows临嘱)
    
    If bln上次打印 Then
        If optReport(opt医嘱_长嘱).value Then
            '进行重整时间判断
            strSQL = "select max(打印时间) as 时间 from 病人医嘱打印 Where 病人id=[1] And 主页id=[2] And Nvl(婴儿,0)=[3] And 期效=0"
            Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, cboBaby.ListIndex)
            If Not IsNull(rsTmp!时间) Then
                If mdat重整时间 > rsTmp!时间 Then
                    MsgBox "上次打印在重整之前，要先回退重整才能清除打印。", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        End If
        
        strSQL = "Zl_病人医嘱打印_Delete(" & mlng病人ID & "," & mlng主页ID & "," & cboBaby.ListIndex & "," & IIF(optReport(opt医嘱_长嘱).value, 0, 1) & ",null,null,1)"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
         
        '生成记录
        strSQL = "Zl_病人医嘱打印_Insert(" & mlng病人ID & "," & mlng主页ID & "," & cboBaby.ListIndex & "," & IIF(optReport(opt医嘱_长嘱).value, 0, 1) & "," & lngRows & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
    Else
        If mdat重整时间 <> CDate("1900-01-01") And optReport(opt医嘱_长嘱).value Then
            strSQL = "select max(页号) as 页号 from 病人医嘱打印 a " & _
                " where a.病人id=[1] and a.主页id=[2] and nvl(a.婴儿,0)=[3] and a.期效=[4] and 打印时间<[5]"
            Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, cboBaby.ListIndex, 0, mdat重整时间)
            lng页号 = Val(rsTmp!页号 & "")
        End If
    
        If Val(txtClearPage.Text) <= lng页号 Then
            If MsgBox("清除打印的医嘱单中包含了重整前打过的内容，若要清除请先回退重整操作，选 是 则最清最近一次重整之后打印的内容，选 否 则不清除，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        
        strSQL = "Zl_病人医嘱打印_Delete(" & mlng病人ID & "," & mlng主页ID & "," & cboBaby.ListIndex & "," & IIF(optReport(opt医嘱_长嘱).value, 0, 1) & "," & Val(txtClearPage.Text) & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
        
        '生成记录
        strSQL = "Zl_病人医嘱打印_Insert(" & mlng病人ID & "," & mlng主页ID & "," & cboBaby.ListIndex & "," & IIF(optReport(opt医嘱_长嘱).value, 0, 1) & "," & lngRows & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
        
        If Val(txtClearPage.Text) > 1 Then
            strSQL = "Select a.打印时间, a.打印人" & vbNewLine & _
                " From 病人医嘱打印 A,病人医嘱记录 b Where a.医嘱id=b.id and b.诊疗类别 in ('5','6') and a.病人id =[1] And a.主页id =[2]" & _
                " and a.婴儿=[3] And a.期效 =[4] And a.页号 =[5] and a.行号=[6]"
            Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, cboBaby.ListIndex, IIF(optReport(opt医嘱_长嘱).value, 0, 1), _
                Val(txtClearPage.Text) - 1, lngRows)
            If Not rsTmp.EOF Then
                str打印人 = rsTmp!打印人 & ""
                strTmp = Format(rsTmp!打印时间, "yyyy-MM-dd HH:mm:ss")
                strTmp = "To_Date('" & strTmp & "','YYYY-MM-DD HH24:MI:SS')"
                str打印时间 = strTmp
            End If
        End If
    End If
    
    '提交数据
    If Not mblnTrans Then gcnOracle.BeginTrans: mblnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zldatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    If mblnTrans Then gcnOracle.CommitTrans: mblnTrans = False
    
    If str打印人 <> "" Then
        strSQL = "Select a.医嘱id,a.页号,a.行号 From 病人医嘱打印 A Where a.打印时间 is null and a.病人id=[1] And a.主页id=[2] and a.婴儿=[3] And a.期效=[4] And a.页号=[5]"
        Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, cboBaby.ListIndex, IIF(optReport(opt医嘱_长嘱).value, 0, 1), Val(txtClearPage.Text) - 1)
        If Not rsTmp.EOF Then
            arrSQL = Array()
            For i = 1 To rsTmp.RecordCount
                strSQL = "zl_病人医嘱打印_Update(" & ZVal(rsTmp!医嘱ID) & "," & rsTmp!页号 & "," & rsTmp!行号 & "," & _
                    mlng病人ID & "," & mlng主页ID & "," & cboBaby.ListIndex & "," & IIF(optReport(opt医嘱_长嘱).value, 0, 1) & "," & _
                    str打印时间 & ",'" & str打印人 & "'," & mlngPrintType & ")"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
                rsTmp.MoveNext
            Next
            If Not mblnTrans Then gcnOracle.BeginTrans: mblnTrans = True
            For i = 0 To UBound(arrSQL)
                Call zldatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
            Next
            If mblnTrans Then gcnOracle.CommitTrans: mblnTrans = False
        End If
    End If
    
    Call RefreshFace
    
    Exit Sub
errH:
    If mblnTrans Then gcnOracle.RollbackTrans: mblnTrans = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
