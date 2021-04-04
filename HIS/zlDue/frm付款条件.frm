VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm付款条件 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "付款条件设置"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8940
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd取消 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7635
      TabIndex        =   50
      Top             =   885
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7635
      TabIndex        =   49
      Top             =   450
      Width           =   1100
   End
   Begin TabDlg.SSTab sstFilter 
      Height          =   6495
      Left            =   75
      TabIndex        =   52
      Top             =   105
      Width           =   7425
      _ExtentX        =   13097
      _ExtentY        =   11456
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "常规(&0)"
      TabPicture(0)   =   "frm付款条件.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra范围"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chkType(4)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkType(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkType(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkType(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkType(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "高级(&1)"
      TabPicture(1)   =   "frm付款条件.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra"
      Tab(1).ControlCount=   1
      Begin VB.CheckBox chkType 
         Caption         =   "设备(&S)"
         Height          =   195
         Index           =   2
         Left            =   3060
         TabIndex        =   24
         Tag             =   "4"
         Top             =   5985
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkType 
         Caption         =   "物资(&M)"
         Height          =   195
         Index           =   1
         Left            =   1845
         TabIndex        =   23
         Tag             =   "2"
         Top             =   5985
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkType 
         Caption         =   "药品(&D)"
         Height          =   195
         Index           =   0
         Left            =   630
         TabIndex        =   22
         Tag             =   "1"
         Top             =   5985
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkType 
         Caption         =   "其他(&Q)"
         Height          =   195
         Index           =   3
         Left            =   4245
         TabIndex        =   25
         Tag             =   "4"
         Top             =   5985
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkType 
         Caption         =   "卫材(&W)"
         Height          =   195
         Index           =   4
         Left            =   5460
         TabIndex        =   26
         Tag             =   "4"
         Top             =   5985
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.Frame fra 
         Caption         =   "其他条件"
         Height          =   4785
         Left            =   -74916
         TabIndex        =   54
         Top             =   420
         Width           =   7215
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   11
            Left            =   960
            TabIndex        =   57
            Tag             =   "开始随货单号"
            Top             =   3480
            Width           =   2800
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   12
            Left            =   4125
            TabIndex        =   56
            Tag             =   "结束随货单号"
            Top             =   3480
            Width           =   2800
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   10
            Left            =   960
            TabIndex        =   48
            Tag             =   "审核人"
            Top             =   4260
            Width           =   5985
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   9
            Left            =   960
            TabIndex        =   46
            Tag             =   "填制人"
            Top             =   3870
            Width           =   5985
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   8
            Left            =   4125
            TabIndex        =   44
            Tag             =   "结束发票号"
            Top             =   2328
            Width           =   2800
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   7
            Left            =   960
            TabIndex        =   42
            Tag             =   "开始发票号"
            Top             =   2328
            Width           =   2800
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   6
            Left            =   4125
            TabIndex        =   40
            Tag             =   "结束入库单据号"
            Top             =   1944
            Width           =   2800
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   5
            Left            =   960
            TabIndex        =   38
            Tag             =   "开始入库单据号"
            Top             =   1944
            Width           =   2800
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   4
            Left            =   4125
            TabIndex        =   36
            Tag             =   "结束批号"
            Top             =   1545
            Width           =   2800
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   3
            Left            =   960
            TabIndex        =   34
            Tag             =   "开始批号"
            Top             =   1545
            Width           =   2800
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   2
            Left            =   960
            TabIndex        =   32
            Tag             =   "产地"
            Top             =   1140
            Width           =   5985
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   1
            Left            =   960
            TabIndex        =   30
            Tag             =   "规格"
            Top             =   780
            Width           =   5985
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   0
            Left            =   960
            TabIndex        =   28
            Tag             =   "品名"
            Top             =   396
            Width           =   5985
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "随货单号"
            Height          =   180
            Index           =   11
            Left            =   192
            TabIndex        =   59
            Top             =   3555
            Width           =   720
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   0
            Left            =   3840
            TabIndex        =   58
            Top             =   3540
            Width           =   180
         End
         Begin VB.Label lblEdit 
            Caption         =   $"frm付款条件.frx":0038
            Height          =   600
            Left            =   960
            TabIndex        =   55
            Top             =   2745
            Width           =   5865
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "审核人"
            Height          =   180
            Index           =   9
            Left            =   372
            TabIndex        =   47
            Top             =   4350
            Width           =   540
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "填制人"
            Height          =   180
            Index           =   8
            Left            =   372
            TabIndex        =   45
            Top             =   3930
            Width           =   540
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   5
            Left            =   3840
            TabIndex        =   43
            Top             =   2385
            Width           =   180
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "发票号"
            Height          =   180
            Index           =   7
            Left            =   372
            TabIndex        =   41
            Top             =   2400
            Width           =   540
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   2
            Left            =   3840
            TabIndex        =   39
            Top             =   2010
            Width           =   180
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "入库单号"
            Height          =   180
            Index           =   6
            Left            =   192
            TabIndex        =   37
            Top             =   2016
            Width           =   720
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   1
            Left            =   3840
            TabIndex        =   35
            Top             =   1605
            Width           =   180
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "批号"
            Height          =   180
            Index           =   5
            Left            =   552
            TabIndex        =   33
            Top             =   1596
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "产地"
            Height          =   180
            Index           =   4
            Left            =   552
            TabIndex        =   31
            Top             =   1200
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "规格"
            Height          =   180
            Index           =   3
            Left            =   552
            TabIndex        =   29
            Top             =   840
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "品名"
            Height          =   180
            Index           =   2
            Left            =   552
            TabIndex        =   27
            Top             =   456
            Width           =   360
         End
      End
      Begin VB.Frame fra范围 
         Height          =   5370
         Left            =   240
         TabIndex        =   53
         Top             =   432
         Width           =   6900
         Begin VB.CheckBox chkStorage 
            Caption         =   "按所有药品库存数量小于发票数量(&S)"
            Height          =   270
            Left            =   564
            TabIndex        =   13
            Top             =   2100
            Width           =   4104
         End
         Begin VB.ComboBox cboStock 
            Height          =   300
            Left            =   1035
            TabIndex        =   21
            Top             =   4830
            Width           =   2460
         End
         Begin VB.TextBox txt随货单 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   810
            Left            =   870
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   19
            Top             =   3885
            Width           =   5745
         End
         Begin VB.CheckBox chk随货单 
            Caption         =   "按随货单号查找(&F)"
            Height          =   384
            Left            =   570
            TabIndex        =   17
            Top             =   3555
            Width           =   1845
         End
         Begin MSComCtl2.DTPicker dtp开始时间 
            Height          =   312
            Index           =   0
            Left            =   1632
            TabIndex        =   10
            Top             =   1704
            Width           =   1608
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   314441731
            CurrentDate     =   36263
         End
         Begin VB.CheckBox chk发票日期 
            Caption         =   "按应付单据的发票日期查找(&R)"
            Height          =   270
            Left            =   564
            TabIndex        =   8
            Top             =   1356
            Width           =   4104
         End
         Begin VB.TextBox txt发票号 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   810
            Left            =   888
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   2685
            Width           =   5745
         End
         Begin VB.CommandButton Cmd供应商 
            Caption         =   "…"
            Height          =   264
            Left            =   6435
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   348
            Width           =   255
         End
         Begin VB.CheckBox chk审核 
            Caption         =   "按应付单据的审核日期查找(&V)"
            Height          =   270
            Left            =   576
            TabIndex        =   3
            Top             =   720
            Width           =   4104
         End
         Begin MSComCtl2.DTPicker dtp开始时间 
            Height          =   312
            Index           =   1
            Left            =   1632
            TabIndex        =   5
            Top             =   996
            Width           =   1608
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   314572803
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp结束时间 
            Height          =   312
            Index           =   1
            Left            =   3540
            TabIndex        =   7
            Top             =   996
            Width           =   1608
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   314572803
            CurrentDate     =   36263
         End
         Begin VB.CheckBox chk发票号 
            Caption         =   "按发票号查找(&F)"
            Height          =   384
            Left            =   564
            TabIndex        =   14
            Top             =   2370
            Width           =   1668
         End
         Begin MSComCtl2.DTPicker dtp结束时间 
            Height          =   312
            Index           =   0
            Left            =   3540
            TabIndex        =   12
            Top             =   1704
            Width           =   1608
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   314572803
            CurrentDate     =   36263
         End
         Begin VB.TextBox txt供应商 
            Height          =   300
            Left            =   924
            MaxLength       =   50
            TabIndex        =   1
            Top             =   324
            Width           =   5520
         End
         Begin VB.Label lblStock 
            AutoSize        =   -1  'True
            Caption         =   "按库房付款"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   4905
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   ":如果存在多张随货单据号，请用逗分隔。"
            Height          =   180
            Index           =   10
            Left            =   2400
            TabIndex        =   18
            Top             =   3300
            Width           =   3330
         End
         Begin VB.Label lbl时间 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "发票日期"
            Height          =   180
            Index           =   0
            Left            =   864
            TabIndex        =   9
            Top             =   1764
            Width           =   720
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   4
            Left            =   3300
            TabIndex        =   11
            Top             =   1764
            Width           =   180
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   ":如果存在多张发票号，请用逗分隔。"
            Height          =   180
            Index           =   0
            Left            =   2205
            TabIndex        =   15
            Top             =   2450
            Width           =   2970
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "供应商(&G)"
            Height          =   180
            Index           =   1
            Left            =   72
            TabIndex        =   0
            Top             =   384
            Width           =   828
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   3
            Left            =   3300
            TabIndex        =   6
            Top             =   1056
            Width           =   180
         End
         Begin VB.Label lbl时间 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "审核日期"
            Height          =   180
            Index           =   1
            Left            =   852
            TabIndex        =   4
            Top             =   1056
            Width           =   720
         End
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   3090
      Left            =   1560
      TabIndex        =   51
      Top             =   4965
      Visible         =   0   'False
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   5450
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   32768
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frm付款条件"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnAdvance As Boolean '是否展开
Private mlng供应商ID As Long
Private mstrSelectTag As String     '当前选择的对象
Private mstrFind As String
Private mstrPrivs As String
Private mblnOK As Boolean
Private mcllFilter As Collection
Private mrsStock As ADODB.Recordset
Private mlngModule As Long
Private mblnNoClick As Boolean

Public Sub 权限设置()
    If Check相关权限(mstrPrivs, "药品") = False Then
        chkType(0).Enabled = False
        chkType(0).Value = 0
    End If
    If Check相关权限(mstrPrivs, "物资") = False Then
        chkType(1).Enabled = False
        chkType(1).Value = 0
    End If
    
    If Check相关权限(mstrPrivs, "设备") = False Then
        chkType(2).Enabled = False
        chkType(2).Value = 0
    End If
    If Check相关权限(mstrPrivs, "其他") = False Then
        chkType(3).Enabled = False
        chkType(3).Value = 0
    End If
    If Check相关权限(mstrPrivs, "卫材") = False Then
        chkType(4).Enabled = False
        chkType(4).Value = 0
    End If
End Sub

Public Function ShowFind(ByVal FrmMain As Form, ByVal lng供应商ID As Long, ByVal strPrivs As String, ByRef cllFilter As Collection, Optional int标记 As Integer = 0) As Boolean
    '--------------------------------------------------------------
    '功能：获取检索所需的SQL语句
    '参数：FrmMain-调用窗体
    '       lng供应商ID-供应商ID
    '       strPrivs-权限串
    '返回：设置了条件返回true,否则返回False
    '说明：
    '--------------------------------------------------------------
    mstrFind = ""
    mstrPrivs = strPrivs
    '问题27930 by lesfeng 2010-03-23
    If int标记 = 1 Then Me.Caption = "标记付款条件设置"
    If CheckCompete = False Then Exit Function
    mlng供应商ID = lng供应商ID
    Me.Show vbModal, FrmMain
    Set cllFilter = mcllFilter
    ShowFind = mblnOK
End Function
 

Private Sub cboStock_Click()
   If mblnNoClick Then Exit Sub
    If cboStock.ListIndex >= 0 Then cboStock.Tag = cboStock.ItemData(cboStock.ListIndex)
End Sub

Private Sub cboStock_KeyPress(KeyAscii As Integer)
    '问题:33640
    If KeyAscii <> 13 Then Exit Sub
    If cboStock.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    If mrsStock Is Nothing Then InitStockData
    If zlSelectDept(Me, mlngModule, cboStock, mrsStock, cboStock.Text, True, "所有部门") = False Then KeyAscii = 0: Exit Sub
    If cboStock.ListIndex >= 0 Then
        cboStock.Tag = cboStock.ItemData(cboStock.ListIndex)
    End If
End Sub

Private Sub cboStock_LostFocus()
    Dim i As Long
    If cboStock.ListCount = 0 Then Exit Sub
    If cboStock.ListIndex < 0 Then
        For i = 0 To cboStock.ListCount - 1
            If Val(cboStock.Tag) = cboStock.ItemData(i) Then
                mblnNoClick = True
                cboStock.ListIndex = i: Exit For
            End If
        Next
    End If
    mblnNoClick = False
End Sub

Private Sub chkStorage_Click()
    Dim i As Integer
    If chkStorage.Value = 1 Then
        chkType(0).Value = 1
        For i = 1 To chkType.Count - 1
            chkType(i).Value = 0
        Next
    End If
End Sub

Private Sub chkType_Click(Index As Integer)
    If Index = 0 Then
        If chkType(Index) = 0 Then
            chkStorage.Value = 0
        End If
    Else
        If chkType(Index) = 1 Then
            chkStorage.Value = IIf(chkType(Index).Value = 1, 0, 1) And chkType(0).Value = 1
        End If
    End If
End Sub

Private Sub chkType_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Index = 4 Then
            If cmd确定.Enabled Then cmd确定.SetFocus
        Else
            zlCommFun.PressKey (vbKeyTab)
        End If
    End If
End Sub

Private Sub chk发票号_Click()
    txt发票号.Enabled = chk发票号.Value = 1
    txt发票号.BackColor = IIf(txt发票号.Enabled, vbWhite, Me.BackColor)
End Sub

Private Sub chk发票号_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub chk随货单_Click()
    txt随货单.Enabled = chk随货单.Value = 1
    txt随货单.BackColor = IIf(txt随货单.Enabled, vbWhite, Me.BackColor)
End Sub

Private Sub chk随货单_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub chk发票日期_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk审核_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub chk发票日期_Click()
    dtp开始时间(0).Enabled = IIf(chk发票日期.Value = 1, True, False)
    dtp结束时间(0).Enabled = IIf(chk发票日期.Value = 1, True, False)
End Sub

Private Sub chk审核_Click()
    dtp开始时间(1).Enabled = IIf(chk审核.Value = 1, True, False)
    dtp结束时间(1).Enabled = IIf(chk审核.Value = 1, True, False)
End Sub
 
Private Sub Cmd供应商_Click()
    Dim strTemp As String
    
    strTemp = frm供应商选择.SelDept(mstrPrivs)
    If strTemp = "" Then Exit Sub
    txt供应商.Text = Mid(strTemp, InStr(strTemp, ",") + 1)
    txt供应商.Tag = Left(strTemp, InStr(strTemp, ",") - 1)
    Unload frm供应商选择
    If chk审核.Enabled And chk审核.Visible Then chk审核.SetFocus
End Sub

Private Sub Cmd取消_Click()
    mblnOK = False
    Unload Me
End Sub

Private Function CheckValied() As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能:检查输入数据的合法性
    '入参:
    '出参:
    '返回: 合法返回true,否则返回False
    '修改人:刘兴宏
    '修改时间:2007/2/28
    '------------------------------------------------------------------------------------------------------
    Dim i As Long
 
    If chk发票号.Value = 1 And Trim(Replace(txt发票号.Text, vbCrLf, "")) = "" Then
        ShowMsgbox "未输入发票号,请检查!"
        sstFilter.Tab = 0
        If chk发票号.Enabled Then Me.chk发票号.SetFocus
        Exit Function
    End If
    
    If chk随货单.Value = 1 And Trim(Replace(txt随货单.Text, vbCrLf, "")) = "" Then
        ShowMsgbox "未输入随货单号,请检查!"
        sstFilter.Tab = 0
        If chk随货单.Enabled Then Me.chk随货单.SetFocus
        Exit Function
    End If
    
    If Check供应商 = False Then Exit Function
    
    For i = 0 To txtEdit.UBound
        If InStr(1, txtEdit(i).Text, "'") > 0 Then
            ShowMsgbox txtEdit(i).Tag & "含用非法字符"
            sstFilter.Tab = 1
            If txtEdit(i).Enabled Then txtEdit(i).SetFocus
            Exit Function
        End If
        If InStr(1, txtEdit(i).Tag, "开始") > 0 Then
            If txtEdit(i).Text > txtEdit(i + 1).Text And txtEdit(i + 1).Text <> "" Then
                ShowMsgbox txtEdit(i).Tag & "大于" & txtEdit(i + 1).Tag
                sstFilter.Tab = 1
                If txtEdit(i).Enabled Then txtEdit(i).SetFocus
                Exit Function
            End If
        End If
    Next
    Dim blnHaving As Boolean
    blnHaving = False
    For i = 0 To chkType.UBound
        If chkType(i).Value = 1 Then
            blnHaving = True
            Exit For
        End If
    Next
    If blnHaving = False Then
        ShowMsgbox "未选择你要查找的类别,请检查"
        sstFilter.Tab = 0
        If chkType(0).Enabled Then Me.chkType(0).SetFocus
        Exit Function
    End If
    
    CheckValied = True
End Function

Private Sub Cmd确定_Click()
    Dim strTemp As String
    Dim i As Long
    mstrFind = ""
    '生成SQL查询条件语句
    Dim intTemp As Integer
    Dim strFind As String
    
    If CheckValied = False Then Exit Sub
    
    Set mcllFilter = New Collection
    mcllFilter.Add Val(txt供应商.Tag), "供应商ID"
    mcllFilter.Add Array("1901-01-01 00:00:00", "1901-01-01 00:00:00"), "审核日期"
    mcllFilter.Add Array("1901-01-01 00:00:00", "1901-01-01 00:00:00"), "发票日期"
    mcllFilter.Add "", "发票号列表"
    mcllFilter.Add "", "随货单号列表"
    mcllFilter.Add "", "系统标识"
    mcllFilter.Add "", "品名"
    mcllFilter.Add "", "规格"
    mcllFilter.Add "", "产地"
    mcllFilter.Add Array("", ""), "批号"
    mcllFilter.Add Array("", ""), "入库单号"
    mcllFilter.Add Array("", ""), "发票号"
    mcllFilter.Add Array("", ""), "随货单号"
    mcllFilter.Add "", "填制人"
    mcllFilter.Add "", "审核人"
    mcllFilter.Add 0, "库房"
    mcllFilter.Add "0", "按所有药品库存数量小于发票数量"
    
    If chk发票日期.Value = 1 And chk审核.Value = 1 Then
        mstrFind = " And ( ([alias]审核日期 Between [2] And [3]) or ([alias]发票日期 Between [4] And [5])) "
        mcllFilter.Remove "审核日期"
        mcllFilter.Remove "发票日期"
        mcllFilter.Add Array(Format(dtp开始时间(1), "yyyy-mm-dd") & " 00:00:00", Format(dtp结束时间(1), "yyyy-mm-dd") & " 23:59:59"), "审核日期"
        mcllFilter.Add Array(Format(dtp开始时间(0), "yyyy-mm-dd") & " 00:00:00", Format(dtp结束时间(0), "yyyy-mm-dd") & " 23:59:59"), "发票日期"
    ElseIf chk审核.Value = 1 Then
        mstrFind = " And ( [alias]审核日期 Between [2] And [3]) "
        mcllFilter.Remove "审核日期"
        mcllFilter.Add Array(Format(dtp开始时间(1), "yyyy-mm-dd") & " 00:00:00", Format(dtp结束时间(1), "yyyy-mm-dd") & " 23:59:59"), "审核日期"
    ElseIf chk发票日期.Value = 1 Then
        mstrFind = " And ([alias]发票日期 Between [4] And [5]) "
        mcllFilter.Remove "发票日期"
        mcllFilter.Add Array(Format(dtp开始时间(0), "yyyy-mm-dd") & " 00:00:00", Format(dtp结束时间(0), "yyyy-mm-dd") & " 23:59:59"), "发票日期"
    End If
    
    '按所有药品库存数量小于发票数量
    mcllFilter.Remove "按所有药品库存数量小于发票数量"
    mcllFilter.Add chkStorage.Value, "按所有药品库存数量小于发票数量"
    
    If chk发票号.Value = 1 Then
        mstrFind = mstrFind & " And Exists (Select /*+ cardinality(J, 10)*/ 1 " & vbCrLf & _
                              "             From Table(Cast(f_Str2list([6]) As zlTools.t_Strlist)) J " & vbCrLf & _
                              "             Where Instr(',' || [alias]发票号 || ',', ',' || Column_Value || ',') > 0) "
        
        mcllFilter.Remove "发票号列表"
        mcllFilter.Add Replace(txt发票号.Text, vbCrLf, ""), "发票号列表"
    End If
    If chk随货单.Value = 1 Then
        mstrFind = mstrFind & " And exists(Select /*+ cardinality(J, 10)*/ 1 " & vbCrLf & _
                              "            From Table(Cast(f_Str2list([7]) As zlTools.t_Strlist)) J " & vbCrLf & _
                              "            Where J.Column_Value = [alias]随货单号) "
        mcllFilter.Remove "随货单号列表"
        mcllFilter.Add Replace(txt随货单.Text, vbCrLf, ""), "随货单号列表"
    End If
    '确定相关类别
    Dim blnAll As Boolean
    blnAll = True
    strTemp = ""
    For i = 0 To chkType.UBound
        If chkType(i).Value = 1 Then
            strTemp = strTemp & "," & i + 1
        Else
          blnAll = False
        End If
    Next
    strTemp = Mid(strTemp, 2)
    If blnAll = False Then
        '1——药品应付款   2——物资应付款   3——设备应付款   4——其他,5--卫生材料
        mstrFind = mstrFind & _
                " And exists(Select /*+ cardinality(J, 10)*/ 1 " & _
                "            From Table(Cast(f_Num2list([8]) As zlTools.t_Numlist)) J " & _
                "            Where J.Column_Value = [alias]系统标识 )"
        mcllFilter.Remove "系统标识"
        mcllFilter.Add strTemp, "系统标识"
    End If
    If cboStock.ListIndex >= 0 Then
        If cboStock.ItemData(cboStock.ListIndex) <> 0 Then
            mstrFind = mstrFind & " And [alias]库房ID=[23]"
            mcllFilter.Remove "库房"
            mcllFilter.Add cboStock.ItemData(cboStock.ListIndex), "库房"
        End If
    End If
    
    '扩展查询条件
    If mblnAdvance = False Then
        GoTo EndSub:
    End If
    '------------------------------------------------------------------------------------------------------------
    '品名
    If Trim(txtEdit(0).Text) <> "" Then
        strTemp = GetMatchingSting(Trim(txtEdit(0).Text), False)
        mcllFilter.Remove "品名"
        mcllFilter.Add strTemp, "品名"
        
        strFind = " And [alias]品名 like [9]"
        If zlCommFun.IsCharAlpha(Trim(txtEdit(0).Text)) Then          '01,11.输入全是字母时只匹配简码
            '0-拼音码,1-五笔码,2-两者
            '.int简码方式 = Val(zlDatabase.GetPara("简码方式", , , True))
            If gSystemPara.int简码方式 = 1 Then
                '五笔码查询
                If Mid(gSystemPara.Para_输入方式, 2, 1) = "1" Then strFind = " And zltools.zlWBCode([alias]品名) Like upper([9]) "
            ElseIf gSystemPara.int简码方式 = 0 Then
                If Mid(gSystemPara.Para_输入方式, 2, 1) = "1" Then strFind = " And zltools.zlspellcode([alias]品名) Like Upper([9]) "
            Else
                If Mid(gSystemPara.Para_输入方式, 2, 1) = "1" Then strFind = " And (zltools.zlWBCode([alias]品名) Like Upper([9]) or zltools.zlspellcode([alias]品名) Like upper([9]) "
            End If
        End If
        mstrFind = mstrFind & strFind
    End If
    '------------------------------------------------------------------------------------------------------------
    '规格
    If Trim(txtEdit(1).Text) <> "" Then
        mstrFind = mstrFind & " And [alias]规格 like [10]"
        strTemp = GetMatchingSting(Trim(txtEdit(1).Text), False)
        mcllFilter.Remove "规格"
        mcllFilter.Add strTemp, "规格"
    End If
    '产地
    If Trim(txtEdit(2).Text) <> "" Then
        mstrFind = mstrFind & " And [alias]规格 like [11]"
        strTemp = GetMatchingSting(Trim(txtEdit(2).Text), False)
        mcllFilter.Remove "产地"
        mcllFilter.Add strTemp, "产地"
    End If
    '批号
    If Trim(txtEdit(3).Text) <> "" And Trim(txtEdit(4).Text) <> "" Then
        mstrFind = mstrFind & " And ([alias]批号 between [12] and [13])"
    ElseIf Trim(txtEdit(3).Text) <> "" Then
        mstrFind = mstrFind & " And [alias]批号 >= [12] "
    ElseIf Trim(txtEdit(4).Text) <> "" Then
        mstrFind = mstrFind & " And [alias]批号 <= [13] "
    End If
    mcllFilter.Remove "批号"
    mcllFilter.Add Array(Trim(txtEdit(3).Text), Trim(txtEdit(4).Text)), "批号"
    '入库单号
    If Trim(txtEdit(5).Text) <> "" And Trim(txtEdit(6).Text) <> "" Then
        mstrFind = mstrFind & " And ([alias]入库单据号 between [14] and [15])"
    ElseIf Trim(txtEdit(5).Text) <> "" Then
        mstrFind = mstrFind & " And [alias]入库单据号 >= [14] "
    ElseIf Trim(txtEdit(6).Text) <> "" Then
        mstrFind = mstrFind & " And [alias]入库单据号 <= [15] "
    End If
    mcllFilter.Remove "入库单号"
    mcllFilter.Add Array(Trim(txtEdit(5).Text), Trim(txtEdit(6).Text)), "入库单号"
    '发票号
    If Trim(txtEdit(7).Text) <> "" And Trim(txtEdit(8).Text) <> "" Then
        mstrFind = mstrFind & " And ([alias]发票号 between [16] and [17])"
    ElseIf Trim(txtEdit(7).Text) <> "" Then
        mstrFind = mstrFind & " And [alias]发票号 >= [16] "
    ElseIf Trim(txtEdit(8).Text) <> "" Then
        mstrFind = mstrFind & " And [alias]发票号 <= [17] "
    End If
    mcllFilter.Remove "发票号"
    mcllFilter.Add Array(Trim(txtEdit(7).Text), Trim(txtEdit(8).Text)), "发票号"
    '随货单号
    If Trim(txtEdit(11).Text) <> "" And Trim(txtEdit(12).Text) <> "" Then
        mstrFind = mstrFind & " And ([alias]随货单号 between [18] and [19])"
    ElseIf Trim(txtEdit(11).Text) <> "" Then
        mstrFind = mstrFind & " And [alias]随货单号 >= [18] "
    ElseIf Trim(txtEdit(12).Text) <> "" Then
        mstrFind = mstrFind & " And [alias]随货单号 <= [19] "
    End If
    mcllFilter.Remove "随货单号"
    mcllFilter.Add Array(Trim(txtEdit(11).Text), Trim(txtEdit(12).Text)), "随货单号"
    '填制人:
    If Trim(txtEdit(9).Text) <> "" Then
        mstrFind = mstrFind & " And [alias]填制人 like [20] "
        mcllFilter.Remove "填制人"
        mcllFilter.Add Trim(txtEdit(9).Text), "填制人"
    End If
    '审核人:
    If Trim(txtEdit(10).Text) <> "" Then
        mstrFind = mstrFind & " And [alias]审核人 like [21] "
        mcllFilter.Remove "审核人"
        mcllFilter.Add Trim(txtEdit(10).Text), "审核人"
    End If
EndSub:
    mcllFilter.Add mstrFind, "过滤"
    
    mblnOK = True
    Unload Me
End Sub

Private Sub dtp结束时间_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
     End If
End Sub

Private Sub dtp开始时间_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
     End If
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    mlngModule = 1323
    '功能:权限控制:2008-08-18 14:41:40
    Call 权限设置
    '问题27930 by lesfeng 2010-03-23
    Call setInitDate
    
    Call InitStockData  '33640
    
    Me.txt供应商.Tag = 0
    On Error GoTo errHandle
    If mlng供应商ID <> 0 Then
        gstrSQL = "Select 编码,名称 from 供应商 where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng供应商ID)
        If rsTemp.EOF = False Then
            txt供应商.Text = "[" & Nvl(rsTemp!编码) & "]" & Nvl(rsTemp!名称)
            txt供应商.Tag = mlng供应商ID
        End If
    End If
    '打开记录集
    sstFilter.Tab = 0
    mblnAdvance = False
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Function CheckCompete() As Boolean
    '--------------------------------------------------------------
    '功能：检查是否有供应商数据
    '参数：
    '返回：是否有供应商数据
    '说明：
    '--------------------------------------------------------------
    Dim rsCompete As New Recordset
    
    CheckCompete = False
    Err = 0
    On Error GoTo ErrHand:
    gstrSQL = "Select id From 供应商 Where (撤档时间 is null or To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01') " & zl_获取站点限制 & "  and  末级=1 and rownum<=2 "
    zlDatabase.OpenRecordset rsCompete, gstrSQL, "检查供应商"
    With rsCompete
        If .EOF Then
            .Close
            ShowMsgbox "供应商信息不全，请在供应商管理中设置供应商信息！"
            Exit Function
        End If
    End With
    CheckCompete = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub Form_Unload(Cancel As Integer)
    If mshSelect.Visible = True Then
        mshSelect.Visible = False
        Select Case mstrSelectTag
            Case "Provider"
                txt供应商.SetFocus
            Case "Booker"
                txtEdit(9).SetFocus
            Case "Verify"
                txtEdit(10).SetFocus
        End Select
        Cancel = True
    End If
End Sub

Private Sub mshSelect_DblClick()
    mshSelect_KeyPress 13
End Sub

Private Sub mshSelect_KeyPress(KeyAscii As Integer)
    With mshSelect
        If KeyAscii = 13 Then
            Select Case mstrSelectTag
                Case "Provider"
                    txt供应商.Text = "[" & .TextMatrix(.Row, 1) & "]" & .TextMatrix(.Row, 2)
                    txt供应商.Tag = .TextMatrix(.Row, 0)
                    If chk审核.Enabled And chk审核.Visible Then chk审核.SetFocus
                Case "9"
                    txtEdit(9) = .TextMatrix(.Row, 2)
                    If txtEdit(10).Enabled And txtEdit(10).Visible Then txtEdit(10).SetFocus
                Case "10"
                    txtEdit(10) = .TextMatrix(.Row, 2)
                    If cmd确定.Enabled Then cmd确定.SetFocus
            End Select
            .Visible = False
            Exit Sub
        End If
    End With
End Sub

Private Sub mshSelect_LostFocus()
    mshSelect.Visible = False
End Sub

Private Sub sstFilter_Click(PreviousTab As Integer)
    With sstFilter
        If .Tab = 1 Then
            mblnAdvance = True
        End If
    End With
End Sub

Private Sub sstFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 Or KeyCode = 13 Then
        If sstFilter.Tab = 0 Then
          If txt供应商.Enabled Then txt供应商.SetFocus
        Else
            If txtEdit(0).Enabled Then txtEdit(0).SetFocus
        End If
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    If InStr(1, txtEdit(Index).Tag, "人") > 0 Then
         SetTxtGotFocus txtEdit(Index), True
    Else
        zlControl.TxtSelAll txtEdit(Index)
    End If
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If InStr(1, txtEdit(Index).Tag, "人") <> 0 Then
            Call SelectPerson(Index, KeyCode)
        Else
            zlCommFun.PressKey (vbKeyTab)
        End If
    End If
End Sub

Private Function SelectPerson(ByVal intIndex As Integer, ByRef KeyCode As Integer) As Boolean
    '功能:选择相关的人员信息
    Dim rsTemp As New ADODB.Recordset
    Dim strKey As String
    If Trim(txtEdit(intIndex).Text) = "" Then
        zlCommFun.PressKey vbKeyTab
        Exit Function
    End If
    txtEdit(intIndex).Text = UCase(txtEdit(intIndex).Text)
    strKey = txtEdit(intIndex).Text
    strKey = GetMatchingSting(strKey)
    
    On Error GoTo errHandle
    gstrSQL = "" & _
        "   Select 编号,简码,姓名 " & _
        "   From 人员表 " & _
        "   Where (姓名 like [1] or 编号 like [1] or 简码 like [1] )  " & zl_获取站点限制 & "" & _
        "       and (撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)" & _
        "   order by 编号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取" & txtEdit(intIndex).Tag, strKey)
       
    With rsTemp
        If .EOF Then
            MsgBox "输入值无效！", vbInformation, gstrSysName
            KeyCode = 0
            txtEdit(intIndex).SelStart = 0
            txtEdit(intIndex).SelLength = Len(txtEdit(intIndex).Text)
            
            Exit Function
        End If
        Dim sngHeight As Single
        
        If .RecordCount > 1 Then
            mstrSelectTag = intIndex
            
            Set mshSelect.Recordset = rsTemp
            sngHeight = sstFilter.Top + fra.Top + txtEdit(intIndex).Top
            If sngHeight > mshSelect.Rows * (mshSelect.RowHeight(0) + 30) + 200 Then
                mshSelect.Height = mshSelect.Rows * (mshSelect.RowHeight(0) + 30) + 200
            Else
                mshSelect.Height = sngHeight
            End If
            With mshSelect
                .Top = sstFilter.Top + fra.Top + txtEdit(intIndex).Top - .Height
                .Left = sstFilter.Left + fra.Left + txtEdit(intIndex).Left
                .Visible = True
                .SetFocus
                .ColWidth(0) = 800
                .ColWidth(1) = 1500
                .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                .Row = 1
                .Col = 0
                .ColSel = .Cols - 1
                .ZOrder 0
                Exit Function
            End With
        Else
            txtEdit(intIndex).Text = IIf(IsNull(!姓名), "", !姓名)
        End If
    End With
    zlCommFun.PressKey vbKeyTab
    Exit Function
    
errHandle:
    If ErrCenter = 1 Then Resume
End Function

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m文本式
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    ImeLanguage False
End Sub

Private Sub txt发票号_GotFocus()
    zlControl.TxtSelAll txt发票号
    zlCommFun.OpenIme False
End Sub

Private Sub txt发票号_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txt发票号_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt发票号, KeyAscii, m文本式
End Sub

Private Function Check供应商() As Boolean
    '---------------------------------------------------------------------------------------------
    '功能:检查相关的供应商
    '返回:成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/11/05
    '---------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, i As Long
    
    Dim strVarr As Variant, strTemp As String, str开始发票号 As String, str结束发票号 As String
    Err = 0: On Error GoTo ErrHand:
    
    If Val(txt供应商.Tag) <> 0 Then
        Check供应商 = True
        Exit Function
    End If
    
    If chk发票号.Value = 1 Then
        If chk审核.Value = 0 Then
            ShowMsgbox "由于未选择相关的审核时间或供应商，为了提高性能，请务必选择一个条件（审核日期或供应商）！!"
            Exit Function
        End If
        
        strVarr = Split(Replace(txt发票号.Text, vbCrLf, ""), ",")
        strTemp = ""
        For i = 0 To UBound(strVarr)
            If Trim(strVarr(i)) <> "" Then
                strTemp = strTemp & "," & strVarr(i)
            End If
        Next
        strTemp = Mid(strTemp, 2)
        If strTemp = "" Then
            ShowMsgbox "未输入相关的发票号,请检查!"
            Exit Function
        End If
        str开始发票号 = strTemp
         
        If InStr(1, strTemp, ",") <> 0 Then
            gstrSQL = "" & _
                "   Select distinct M.id,M.编码,M.名称,M.末级,M.简码,M.许可证号,M.许可证效期,M.执照号,M.执照效期," & _
                "           M.税务登记号,M.地址,M.开户银行,M.帐号,M.联系人,M.建档时间,M.类型,M.信用期 " & _
                "   From 应付记录 A,供应商 M" & _
                "   Where A.单位ID=M.ID and a.审核日期 between [3] and [4] " & vbCrLf & _
                "         And Exists (Select /*+ cardinality(J, 10)*/ 1 " & vbCrLf & _
                "                     From Table(Cast(f_Str2list(A.发票号) As zlTools.t_Strlist)) J " & vbCrLf & _
                "                     Where exists(Select /*+ cardinality(M, 10)*/ 1 " & vbCrLf & _
                "                                  From Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) M " & vbCrLf & _
                "                                  Where j.Column_Value=m.Column_Value)) "
        Else
            gstrSQL = "" & _
                "   Select  distinct M.id,M.编码,M.名称,M.末级,M.简码,M.许可证号,M.许可证效期,M.执照号,M.执照效期," & _
                "           M.税务登记号,M.地址,M.开户银行,M.帐号,M.联系人,M.建档时间,M.类型,M.信用期 " & _
                "   From 应付记录 A,供应商 M  " & _
                "   Where  A.单位ID=M.ID and a.审核日期 between [3] and [4] " & _
                "         And Exists (Select /*+ cardinality(J, 10)*/ 1 " & vbCrLf & _
                "                     From Table(Cast(f_Str2list(A.发票号) As zlTools.t_Strlist)) J " & vbCrLf & _
                "                     Where j.Column_Value=[1])"
                
        End If
    ElseIf Trim(txtEdit(7).Text) <> "" Or txtEdit(8).Text <> "" Then
        If chk审核.Value = 0 Then
            ShowMsgbox "由于未选择相关的审核时间或供应商，为了提高性能，请务必选择一个条件（审核日期或供应商）！!"
            Exit Function
        End If
        strTemp = ""
        str开始发票号 = Trim(txtEdit(7).Text)
        str结束发票号 = Trim(txtEdit(8).Text)
        
        If Trim(txtEdit(7).Text) <> "" And Trim(txtEdit(8).Text) <> "" Then
            strTemp = "  And Exists (Select /*+ cardinality(J, 10)*/ 1 " & vbCrLf & _
                      "              From Table(Cast(f_Str2list(A.发票号) As zlTools.t_Strlist)) J " & vbCrLf & _
                      "              Where j.Column_Value>=[1]  and j.Column_Value<=[2])"
        ElseIf Trim(txtEdit(7).Text) = "" And Trim(txtEdit(8).Text) <> "" Then
'            strTemp = "  And A.发票号<=[2] "
            strTemp = "  And Exists (Select /*+ cardinality(J, 10)*/ 1 " & vbCrLf & _
                      "              From Table(Cast(f_Str2list(A.发票号) As zlTools.t_Strlist)) J " & vbCrLf & _
                      "              Where  j.Column_Value<=[2])"
        Else
            strTemp = "  And Exists (Select /*+ cardinality(J, 10)*/ 1 " & vbCrLf & _
                      "              From Table(Cast(f_Str2list(A.发票号) As zlTools.t_Strlist)) J " & vbCrLf & _
                      "              Where  j.Column_Value>=[1])"
        End If
        
        gstrSQL = "" & _
            "   Select  distinct M.id,M.编码,M.名称,M.末级,M.简码,M.许可证号,M.许可证效期,M.执照号,M.执照效期," & _
            "           M.税务登记号,M.地址,M.开户银行,M.帐号,M.联系人,M.建档时间,M.类型,M.信用期 " & _
            "   From  应付记录 A,供应商 M" & _
            "   Where  A.单位ID=M.ID And a.审核日期 between [3] and [4]  " & strTemp
    Else
        If Val(txt供应商.Tag) = 0 Then
            ShowMsgbox "供应商未选择,不能继续!"
            sstFilter.Tab = 0
            If txt供应商.Enabled Then txt供应商.SetFocus
        Else
            Check供应商 = True
        End If
        Exit Function
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str开始发票号, str结束发票号, _
           CDate(Format(dtp开始时间(1).Value, "yyyy-mm-dd") & " 00:00:00"), CDate(Format(dtp结束时间(1).Value, "yyyy-mm-dd") & " 23:59:59"))
    If rsTemp.EOF = True Then
        ShowMsgbox "无此发票号的供应商,请检查!"
        Exit Function
    End If
    txt供应商.Text = Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称)
    txt供应商.Tag = Nvl(rsTemp!ID)
    Check供应商 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub txt供应商_Change()
    txt供应商.Tag = ""
End Sub

Private Sub txt供应商_GotFocus()
    SetTxtGotFocus txt供应商, True
End Sub

Private Function Select供应商(ByVal strKey As String) As Boolean
    '----------------------------------------------------------------------------------------
    '功能:选择供应商
    '参数:strKey-选择供应商
    '返回:选择成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/11/5
    '----------------------------------------------------------------------------------------
    Dim str权限 As String
    Dim rsTemp As ADODB.Recordset
    Dim blnCancel As Boolean
    Err = 0: On Error GoTo ErrHand:
    If strKey <> "" Then
        strKey = GetMatchingSting(strKey)
    End If
      
    str权限 = " and " & Get分类权限(mstrPrivs)
    gstrSQL = "" & _
        "   Select id, 编码,名称,末级,简码,许可证号,许可证效期,执照号,执照效期,税务登记号,地址,开户银行,帐号,联系人,建档时间,类型,信用期" & _
        "   From 供应商 " & _
        "   where   末级=1 " & zl_获取站点限制 & "  " & _
        "           and  (撤档时间 is null or 撤档时间>=to_date('3000-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss')) " & _
        "           and (编码 like [1] or 名称 like [1] or 简码 like [1])  " & str权限
    'ShowSelect:
    '功能：多功能选择器
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
    Dim lngX As Long, lngY As Long, lngH As Long
    lngX = Me.Left + txt供应商.Left + Screen.TwipsPerPixelX
    lngY = Me.Top + Me.Height - Me.ScaleHeight + txt供应商.Top
    lngH = txt供应商.Height
    
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "供应商选择", False, "", "选择供应商", False, True, True, lngX, lngY, lngH, blnCancel, False, True, strKey)
    If blnCancel Then Exit Function
    If rsTemp Is Nothing Then
        ShowMsgbox "不存在符何条件的供应商,请检查!"
        Exit Function
    End If
    If rsTemp.State <> 1 Then Exit Function
    txt供应商 = "[" & Nvl(rsTemp!编码) & "]" & Nvl(rsTemp!名称)
    txt供应商.Tag = Nvl(rsTemp!ID)
    Select供应商 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub txt供应商_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As New Recordset
    Dim strTemp As String
    Dim str权限 As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub '
    If Val(txt供应商.Tag) = 0 Then
        If Trim(txt供应商.Text) <> "" Then
            If Select供应商(Trim(txt供应商.Text)) = False Then
                Exit Sub
            End If
        End If
    End If
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt供应商_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
   
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt随货单_GotFocus()
    zlControl.TxtSelAll txt随货单
    zlCommFun.OpenIme False
End Sub

Private Sub txt随货单_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt随货单_KeyPress(KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txt随货单, KeyAscii, m文本式)
End Sub
'问题27930 by lesfeng 2010-03-23
Private Sub setInitDate()
    Dim arrHead As Variant
    Dim strMonth As String
    Dim intGetEndMonth As Integer
    Dim intGetBeginMonth As Integer
    Dim blnMonth As Boolean
    Dim dtTempDate As Date
    Dim dtTemp As Date
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    strMonth = zlDatabase.GetPara("设置付款时间", glngSys, 1323)
    If InStr(1, strMonth, "-") > 0 Then
        arrHead = Split(strMonth, "-")
        blnMonth = Val(arrHead(0)) = 1
        intGetEndMonth = Val(arrHead(1))
        intGetBeginMonth = Val(arrHead(2))
    Else
        blnMonth = False
        intGetEndMonth = 0
        intGetBeginMonth = 0
    End If
    On Error GoTo errHandle
    If blnMonth Then
        dtTempDate = DateAdd("m", -intGetEndMonth, zlDatabase.Currentdate)
        dtTemp = CDate(Format(dtTempDate, "yyyy-MM") & "-01")
        strSQL = "select to_date('" & Format(dtTemp, "yyyy-mm-dd") & "','yyyy-mm-dd') -1/24/60/60 as dtdate from dual"
        zlDatabase.OpenRecordset rsTemp, strSQL, Me.Caption
        If Not rsTemp.EOF Then
            Me.dtp结束时间(0) = IIf(IsNull(rsTemp!dtdate), zlDatabase.Currentdate, rsTemp!dtdate)
            If intGetBeginMonth = 0 Then
                Me.dtp开始时间(0) = Me.dtp结束时间(0)
            Else
                Me.dtp开始时间(0) = DateAdd("m", -intGetBeginMonth, dtTemp)
            End If
        Else
            Me.dtp结束时间(0) = zlDatabase.Currentdate
            Me.dtp开始时间(0) = DateAdd("d", -7, Me.dtp结束时间(0))
        End If
        rsTemp.Close
    Else
        Me.dtp结束时间(0) = zlDatabase.Currentdate
        Me.dtp开始时间(0) = DateAdd("d", -7, Me.dtp结束时间(0))
    End If
    Me.dtp结束时间(1) = Me.dtp结束时间(0)
    Me.dtp开始时间(1) = Me.dtp开始时间(0)
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub
Private Sub InitStockData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载库房数据
    '编制:刘兴洪
    '日期:2010-11-02 16:13:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strStock As String
    strStock = "HIJKLMN"
    strStock = strStock & "V"   '卫材库和制剂室(K)
    strStock = strStock & "RS"  '物资库房和供应室
    strStock = strStock & "T"   '设备科
    On Error GoTo errHandle
    gstrSQL = "" & _
    "   SELECT DISTINCT a.id,A.编码, a.名称,A.简码 " & _
    "   FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " & _
    "   Where (a.站点 = '" & gstrNodeNo & "' Or a.站点 is Null) And c.工作性质 = b.名称 " & _
    "           AND Instr([1],b.编码,1) > 0 " & _
    "           AND a.id = c.部门id " & _
    "           AND TO_CHAR(a.撤档时间, 'yyyy-MM-dd') = '3000-01-01'" & _
    "   Order by 编码"
    Set mrsStock = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strStock)
    With mrsStock
        cboStock.Clear
        cboStock.AddItem "所有库房"
        cboStock.ListIndex = cboStock.NewIndex
        Do While Not .EOF
            cboStock.AddItem Nvl(!编码) & IIf(Nvl(!编码) = "", "", "-") & Nvl(!名称)
            cboStock.ItemData(cboStock.NewIndex) = Val(Nvl(!ID))
            If cboStock.ListIndex < 0 Then cboStock.ListIndex = cboStock.NewIndex
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


