VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmPurchaseSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤"
   ClientHeight    =   5130
   ClientLeft      =   3150
   ClientTop       =   3165
   ClientWidth     =   7515
   Icon            =   "frmPurchaseSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   2055
      Left            =   1920
      TabIndex        =   39
      Top             =   4680
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3625
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
   Begin TabDlg.SSTab sstFilter 
      Height          =   4695
      Left            =   120
      TabIndex        =   40
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   8281
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "范围(&R)"
      TabPicture(0)   =   "frmPurchaseSearch.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra范围"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "附加条件(&D)"
      TabPicture(1)   =   "frmPurchaseSearch.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra附加条件"
      Tab(1).ControlCount=   1
      Begin VB.Frame fra附加条件 
         Height          =   3810
         Left            =   -74760
         TabIndex        =   48
         Top             =   600
         Width           =   5505
         Begin VB.TextBox txt条码 
            Height          =   300
            Left            =   1380
            TabIndex        =   50
            Top             =   2640
            Width           =   3765
         End
         Begin VB.CheckBox Chk供应商 
            Caption         =   "供应商"
            Height          =   300
            Left            =   480
            TabIndex        =   19
            Top             =   660
            Width           =   1110
         End
         Begin VB.TextBox txt供应商 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   20
            Top             =   660
            Width           =   3255
         End
         Begin VB.CommandButton Cmd供应商 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4890
            TabIndex        =   21
            Top             =   660
            Width           =   255
         End
         Begin VB.CommandButton Cmd材料 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4890
            TabIndex        =   18
            Top             =   240
            Width           =   255
         End
         Begin VB.TextBox Txt材料 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   17
            Top             =   240
            Width           =   3255
         End
         Begin VB.CheckBox Chk材料 
            Caption         =   "卫生材料"
            Height          =   300
            Left            =   480
            TabIndex        =   16
            Top             =   240
            Width           =   1035
         End
         Begin VB.TextBox Txt填制人 
            Height          =   300
            Left            =   1380
            MaxLength       =   8
            TabIndex        =   30
            Top             =   1860
            Width           =   1365
         End
         Begin VB.TextBox Txt审核人 
            Height          =   300
            Left            =   3780
            MaxLength       =   8
            TabIndex        =   32
            Top             =   1860
            Width           =   1365
         End
         Begin VB.TextBox Txt开始发票号 
            Height          =   300
            Left            =   1380
            TabIndex        =   34
            Top             =   2250
            Width           =   1365
         End
         Begin VB.TextBox Txt结束发票号 
            Height          =   300
            Left            =   3780
            TabIndex        =   36
            Top             =   2250
            Width           =   1365
         End
         Begin VB.CheckBox Chk生产商 
            Caption         =   "生产商"
            Height          =   300
            Left            =   480
            TabIndex        =   22
            Top             =   1050
            Width           =   1155
         End
         Begin VB.TextBox txt生产商 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            TabIndex        =   23
            Top             =   1050
            Width           =   3255
         End
         Begin VB.CommandButton Cmd生产商 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4890
            TabIndex        =   24
            Top             =   1050
            Width           =   255
         End
         Begin MSComCtl2.DTPicker dtp开始时间 
            Height          =   315
            Index           =   2
            Left            =   1650
            TabIndex        =   26
            Top             =   1455
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   119996419
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp结束时间 
            Height          =   315
            Index           =   2
            Left            =   3540
            TabIndex        =   28
            Top             =   1455
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   119996419
            CurrentDate     =   36263
         End
         Begin VB.CheckBox chk生产日期 
            Caption         =   "生产日期"
            Height          =   300
            Left            =   480
            TabIndex        =   25
            Top             =   1462
            Width           =   1095
         End
         Begin VB.Label lbl条码 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "条  码"
            Height          =   180
            Left            =   750
            TabIndex        =   49
            Top             =   2700
            Width           =   540
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   4
            Left            =   3300
            TabIndex        =   27
            Top             =   1522
            Width           =   180
         End
         Begin VB.Label Lbl填制人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "填制人"
            Height          =   180
            Left            =   750
            TabIndex        =   29
            Top             =   1920
            Width           =   540
         End
         Begin VB.Label Lbl审核人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "审核人"
            Height          =   180
            Left            =   3120
            TabIndex        =   31
            Top             =   1920
            Width           =   540
         End
         Begin VB.Label Lbl发票号 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "发票号"
            Height          =   180
            Left            =   750
            TabIndex        =   33
            Top             =   2310
            Width           =   540
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   2
            Left            =   3240
            TabIndex        =   35
            Top             =   2310
            Width           =   180
         End
      End
      Begin VB.Frame fra范围 
         Height          =   3810
         Left            =   240
         TabIndex        =   41
         Top             =   600
         Width           =   5520
         Begin VB.CheckBox chk无发票 
            Caption         =   "无发票"
            Height          =   180
            Left            =   2760
            TabIndex        =   15
            Top             =   3480
            Width           =   1095
         End
         Begin VB.CheckBox chk有发票 
            Caption         =   "有发票"
            Height          =   180
            Left            =   720
            TabIndex        =   14
            Top             =   3480
            Width           =   1095
         End
         Begin VB.CheckBox chkYesVerifyBack 
            Caption         =   "已审核退库"
            Enabled         =   0   'False
            Height          =   180
            Left            =   2760
            TabIndex        =   13
            Top             =   3080
            Width           =   1215
         End
         Begin VB.CheckBox chkNOVerifyBack 
            Caption         =   "未审核退库"
            Height          =   180
            Left            =   720
            TabIndex        =   12
            Top             =   3080
            Width           =   1215
         End
         Begin VB.CheckBox chkNot高值耗材 
            Caption         =   "非高值耗材单据"
            Enabled         =   0   'False
            Height          =   180
            Left            =   2760
            TabIndex        =   9
            Top             =   2280
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.CheckBox chk高值耗材 
            Caption         =   "高值耗材单据"
            Enabled         =   0   'False
            Height          =   180
            Left            =   720
            TabIndex        =   8
            Top             =   2280
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chk包含财务审核 
            Caption         =   "包含财务审核"
            Height          =   180
            Left            =   2760
            TabIndex        =   11
            Top             =   2680
            Value           =   1  'Checked
            Width           =   1425
         End
         Begin VB.TextBox txt开始No 
            Height          =   300
            Left            =   840
            MaxLength       =   8
            TabIndex        =   0
            Top             =   360
            Width           =   1605
         End
         Begin VB.TextBox txt结束NO 
            Height          =   300
            Left            =   2970
            MaxLength       =   8
            TabIndex        =   1
            Top             =   360
            Width           =   1605
         End
         Begin VB.CheckBox chk填制 
            Caption         =   "未审核单据"
            Height          =   300
            Left            =   480
            TabIndex        =   2
            Top             =   840
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chk审核 
            Caption         =   "已审核单据"
            Height          =   300
            Left            =   480
            TabIndex        =   5
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CheckBox chkStrike 
            Caption         =   "包含冲销"
            Enabled         =   0   'False
            Height          =   180
            Left            =   720
            TabIndex        =   10
            Top             =   2680
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker dtp开始时间 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   3
            Top             =   1080
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   120913923
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp结束时间 
            Height          =   315
            Index           =   0
            Left            =   3585
            TabIndex        =   4
            Top             =   1080
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   86900739
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp开始时间 
            Height          =   315
            Index           =   1
            Left            =   1680
            TabIndex        =   6
            Top             =   1845
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   86900739
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp结束时间 
            Height          =   315
            Index           =   1
            Left            =   3585
            TabIndex        =   7
            Top             =   1845
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   86507523
            CurrentDate     =   36263
         End
         Begin VB.Label LblNO 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "No"
            Height          =   180
            Left            =   480
            TabIndex        =   47
            Top             =   420
            Width           =   180
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   1
            Left            =   2640
            TabIndex        =   46
            Top             =   420
            Width           =   180
         End
         Begin VB.Label lbl时间 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "审核日期"
            Height          =   180
            Index           =   1
            Left            =   900
            TabIndex        =   45
            Top             =   1905
            Width           =   720
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   3
            Left            =   3345
            TabIndex        =   44
            Top             =   1905
            Width           =   180
         End
         Begin VB.Label lbl时间 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "填制日期"
            Height          =   180
            Index           =   0
            Left            =   900
            TabIndex        =   43
            Top             =   1140
            Width           =   720
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   0
            Left            =   3345
            TabIndex        =   42
            Top             =   1140
            Width           =   180
         End
      End
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6330
      TabIndex        =   38
      Top             =   930
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6330
      TabIndex        =   37
      Top             =   435
      Width           =   1100
   End
End
Attribute VB_Name = "FrmPurchaseSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFind As String  '查找字符串
Private BlnAdvance As Boolean '是否展开
Private mdatStart As Date   '开始时间
Private mdatEnd As Date     '结束时间
Private mdatVerifyStart As Date
Private mdatVerifyEnd As Date
Private mfrmMain As Form    '父窗体
Private mstrSelectTag As String     '当前选择的对象
Private mstrOthers(0 To 13) As String ' 0-记录状态(计划类型),1-开始单号,2-结束单号,3-材料id,4-对方部门id(或入出类别id或编制方法(计划用)),5-填制人,6-审核人,7-供应商ID,8-生产商,9-开始生产日期,10-结束生产日期,11-开始发票号,12-结束发票号,13-条码信息
Public lng材料ID As Long
Private mstr高值耗材 As String      '用来记录高值耗材是否被选择
Private mint有发票 As Integer
Private mint无发票 As Integer

Public Function GetSearch(ByVal frmMain As Form, _
        ByRef datStart As Date, ByRef datEnd As Date, _
        ByRef datVerifyStart As Date, ByRef datVerifyEnd As Date, _
        ByRef strOthers() As String, ByRef str高值耗材 As String, ByRef intNo发票 As Integer, _
        ByRef intYes发票 As Integer) As String
    mstrFind = ""
    mstrSelectTag = ""
    Set mfrmMain = frmMain
    If Not CheckCompete Then Exit Function
    
    Me.Show vbModal, mfrmMain
    GetSearch = mstrFind
    datStart = mdatStart
    datEnd = mdatEnd
    datVerifyStart = mdatVerifyStart
    datVerifyEnd = mdatVerifyEnd
    strOthers = mstrOthers
    str高值耗材 = mstr高值耗材
    intNo发票 = mint无发票
    intYes发票 = mint有发票
End Function


Private Sub chkStrike_Click()
    chk包含财务审核.Enabled = chkStrike.Value = 1
End Sub

Private Sub chkStrike_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub
Private Sub chk包含财务审核_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub Chk供应商_Click()
    txt供应商.Enabled = IIf(Chk供应商.Value = 1, True, False)
    Cmd供应商.Enabled = IIf(Chk供应商.Value = 1, True, False)
    
End Sub

Private Sub Chk供应商_KeyDown(KeyCode As Integer, Shift As Integer)
    If Chk供应商.Value = 1 Then
        txt供应商.SetFocus
    Else
        Chk生产商.SetFocus
    End If
End Sub


Private Sub chk审核_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then
        If chk审核.Value = 0 Then
            cmd确定.SetFocus
        Else
            SendKeys vbTab
        End If
    End If
    
End Sub

Private Sub chk生产日期_Click()
    dtp开始时间(2).Enabled = chk生产日期.Value = 1
    dtp结束时间(2).Enabled = dtp开始时间(2).Enabled
End Sub

Private Sub chk生产日期_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub Chk生产商_Click()
    Me.txt生产商.Enabled = IIf(Chk生产商.Value = 1, True, False)
    Cmd生产商.Enabled = IIf(Chk生产商.Value = 1, True, False)
End Sub

Private Sub Chk生产商_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        
        If Chk生产商.Value = 1 Then
            txt生产商.SetFocus
        
        Else
            Txt填制人.SetFocus
        End If
    End If
End Sub

Private Sub chk填制_Click()
    dtp开始时间(0).Enabled = IIf(chk填制.Value = 1, True, False)
    dtp结束时间(0).Enabled = IIf(chk填制.Value = 1, True, False)
    chkNOVerifyBack.Enabled = IIf(chk填制.Value = 1, True, False)
    If chk填制.Value = 0 Then chkNOVerifyBack.Value = 0
End Sub

Private Sub chk审核_Click()
    dtp开始时间(1).Enabled = IIf(chk审核.Value = 1, True, False)
    dtp结束时间(1).Enabled = IIf(chk审核.Value = 1, True, False)
    chkStrike.Enabled = IIf(chk审核.Value = 1, True, False)
    chk包含财务审核.Enabled = chkStrike.Value = 1
    chkYesVerifyBack.Enabled = IIf(chk审核.Value = 1, True, False)
    If chk审核.Value = 0 Then chkYesVerifyBack.Value = 0
End Sub

Private Sub chk填制_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Chk材料_Click()
    Txt材料.Enabled = IIf(Chk材料.Value = 1, True, False)
    Cmd材料.Enabled = IIf(Chk材料.Value = 1, True, False)
End Sub

Private Sub Chk材料_GotFocus()
    sstFilter.Tab = 1
    Chk材料.SetFocus
End Sub

Private Sub Chk材料_KeyDown(KeyCode As Integer, Shift As Integer)
    If Chk材料.Value = 1 Then
        Txt材料.SetFocus
    ElseIf Chk供应商.Visible = True Then
        Chk供应商.SetFocus
    End If
End Sub



Private Sub Cmd供应商_Click()
    Dim rsTemp As New Recordset
    Dim blnCancel As Boolean
    Dim vRect As RECT
    
    vRect = zlControl.GetControlRect(txt供应商.hwnd)
    
    gstrSQL = "" & _
        "   Select id,上级ID,编码,简码,名称,末级 " & _
        "   From 供应商 " & _
        "   where (substr(类型,5,1)=1 And (站点=[1] or 站点 is null) Or Nvl(末级,0)=0) " & _
        "   Start with 上级ID is null connect by prior ID =上级ID order by level,ID"
    
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
    
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 2, "供应商选择", True, "", "请选择符合材料的供应商", True, True, True, vRect.Left - 15, vRect.Top, txt供应商.Height, blnCancel, False, False, gstrNodeNo)
        
    If rsTemp Is Nothing Or blnCancel Then Exit Sub
    If rsTemp.State <> 1 Then Exit Sub
    
    With rsTemp
        txt供应商.Text = zlStr.NVL(!名称)
        txt供应商.Tag = zlStr.NVL(!Id)
    End With
End Sub

Private Sub Cmd取消_Click()
    mstrFind = ""
    Dim i As Integer
    For i = 0 To 13
        mstrOthers(i) = ""
    Next
    Unload Me
End Sub

Private Sub cmd确定_Click()
    Dim 未审核子条件 As String
    Dim 已审核子条件 As String
    
    mint有发票 = 0
    mint无发票 = 0
    '检查数据
    If Chk材料.Value = 1 Then
        If Txt材料.Tag = 0 Then
            MsgBox "请选择需查询的卫材信息！", vbInformation, gstrSysName
            Me.Txt材料.SetFocus
            Exit Sub
        End If
    End If
    If Chk供应商.Value = 1 Then
        If txt供应商.Tag = 0 Then
            MsgBox "请选择需查询的卫材供应商信息！", vbInformation, gstrSysName
            Me.txt供应商.SetFocus
            Exit Sub
        End If
    End If
    If Chk生产商.Value = 1 Then
        If txt生产商.Tag = 0 Then
            MsgBox "请选择需查询的卫材生产商信息！", vbInformation, gstrSysName
            Me.txt生产商.SetFocus
            Exit Sub
        End If
    End If
    
    If chk填制.Value = 0 And chk审核.Value = 0 Then
        MsgBox "对不起，必须选择一个填制日期或者审核日期!", vbInformation, gstrSysName
        chk填制.SetFocus
        Exit Sub
    End If
    Dim i As Integer
    For i = 0 To 13
        mstrOthers(i) = ""
    Next
    
    mstrFind = ""
    '基本查询条件
    '参数范围:[1]-库房id,[2]:开始填制日期,[3]结束填制日期,[4]开始审核日期,[5] 结束审核日期,[6]-记录状态,[7]开始单据号,[8]结束单据号,[9]材料id,[10]对方部门id,[11]填制人,[12]审核人[13]-供应商ID,[14]-生产商,[15]-开始生产日期,[16]-结束生产日期,[17]-开始发票号,[18]-结束发票号
    mdatStart = Format("1901 - 01 - 01", "yyyy-mm-dd")
    mdatEnd = Format("1901-01-01", "yyyy-mm-dd")
    mdatVerifyStart = Format("1901-01-01", "yyyy-mm-dd")
    mdatVerifyEnd = Format("1901-01-01", "yyyy-mm-dd")
    mstrOthers(0) = IIf(chkStrike.Value = 1, "0", "1")
    
    '未审核下的子条件
    If chkNOVerifyBack.Value = 0 Then '不勾选未审核退库，只显示入库的
        未审核子条件 = 未审核子条件 & " and nvl(a.发药方式,0)=0 "
    End If
    '已审核下的子条件
    If chkStrike.Value = 1 Then '含冲销
        已审核子条件 = IIf(chk包含财务审核, "", " And nvl(A.费用ID,0)=0 ")
    Else
        已审核子条件 = "and a.记录状态 =[6]"
    End If
    If chkYesVerifyBack.Value = 0 Then  '不勾选已审核退库，只显示入库的
        已审核子条件 = 已审核子条件 & " and nvl(a.发药方式,0)=0 "
    End If
    
    If chk填制.Value = 1 And chk审核.Value = 1 Then '包含已审核和未审核单据
    
        mstrFind = " And ((A.填制日期 Between [2] And [3] and A.审核日期 is null " & 未审核子条件 & " )  or (A.审核日期 Between [4] And [5] " & 已审核子条件 & "))"
        
        mdatStart = Format(dtp开始时间(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp结束时间(0), "yyyy-mm-dd")
                
        mdatVerifyStart = Format(dtp开始时间(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp结束时间(1), "yyyy-mm-dd")
        
    ElseIf chk审核.Value = 1 Then

        mstrFind = " And A.审核日期 Between [4] And [5] " & 已审核子条件

        mdatVerifyStart = Format(dtp开始时间(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp结束时间(1), "yyyy-mm-dd")
    ElseIf chk填制.Value = 1 Then
        mstrFind = " And (A.填制日期 Between [2] And [3] and A.审核日期 is null " & 未审核子条件 & ")  "
        mdatStart = Format(dtp开始时间(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp结束时间(0), "yyyy-mm-dd")
    End If
    
    '发票
    If chk有发票.Value = 1 And chk无发票.Value = 0 Then
        mstrFind = mstrFind & " And e.发票号 is not null "
        mint有发票 = 1
        mint无发票 = 0
    ElseIf chk无发票.Value = 1 And chk有发票.Value = 0 Then
        mstrFind = mstrFind & " And e.发票号 is null "
        mint有发票 = 0
        mint无发票 = 1
    End If
    
    If chk高值耗材.Value = 1 And chkNot高值耗材.Value = 0 Then '高值耗材
        mstr高值耗材 = " and  (a.费用id > 1 or d.高值材料=1) "
    ElseIf chk高值耗材.Value = 0 And chkNot高值耗材.Value = 1 Then '非高值耗材
        mstr高值耗材 = " and (d.高值材料=0 or d.高值材料 is null) " '等于1的是财务审核的单据
    Else
        mstr高值耗材 = ""
    End If
    
    Dim intYear As Integer, strYear As String
    
    If Len(txt开始No) < 8 And Len(txt开始No) > 0 Then
        Me.txt开始No = UCase(LTrim(Me.txt开始No))
        intYear = Format(sys.Currentdate, "YYYY") - 1990
        strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
        If Len(txt开始No) < 8 Then Me.txt开始No = strYear & String(7 - Len(txt开始No), "0") & Me.txt开始No
    End If
    If Len(txt结束NO) < 8 And Len(txt结束NO) > 0 Then
        Me.txt结束NO = UCase(LTrim(Me.txt结束NO))
        intYear = Format(sys.Currentdate, "YYYY") - 1990
        strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
        If Len(txt结束NO) < 8 Then Me.txt结束NO = strYear & String(7 - Len(txt结束NO), "0") & Me.txt结束NO
    End If
    
    mstrOthers(1) = Trim(Me.txt开始No.Text)
    mstrOthers(2) = Trim(Me.txt结束NO.Text)

    If Me.txt开始No <> "" And Me.txt结束NO <> "" Then mstrFind = mstrFind & " And A.No >= [7] And A.No <=[8] "
    If Me.txt开始No <> "" And Me.txt结束NO = "" Then mstrFind = mstrFind & " And A.No >= [7] "
    If Me.txt开始No = "" And Me.txt结束NO <> "" Then mstrFind = mstrFind & " And A.No <=[8] "
    
    '扩展查询条件
    
    If BlnAdvance = False Then
        Unload Me
        Exit Sub
    End If
    ' 0-记录状态(计划类型),1-开始单号,2-结束单号,3-材料id,4-对方部门id(或入出类别id),5-填制人,
    ' 6-审核人,7-供应商ID,8-生产商,9-开始生产日期,10-结束生产日期,11-开始发票号,12-结束发票号
    
    If Chk材料.Value = 1 Then
        lng材料ID = Txt材料.Tag
        mstrFind = mstrFind & " And A.药品ID+0=[9]"
        mstrOthers(3) = Txt材料.Tag
    End If
    If Me.Txt填制人 <> "" Then
        mstrFind = mstrFind & " And A.填制人 like '" & Me.Txt填制人 & "%'"
        mstrOthers(5) = Trim(Me.Txt填制人) & "%"
    End If
    If Me.Txt审核人 <> "" Then
        mstrFind = mstrFind & " And A.审核人 like [12]"
        mstrOthers(6) = Trim(Me.Txt审核人) & "%"
    End If
    
    If Chk供应商.Value = 1 Then
        mstrFind = mstrFind & " And A.供药单位ID+0=[13]"
        mstrOthers(7) = txt供应商.Tag
    End If
    If Chk生产商.Value = 1 Then
        mstrFind = mstrFind & " And A.产地=[14]"
        mstrOthers(8) = txt生产商.Text
    End If
    If chk生产日期.Value = 1 Then
        mstrFind = mstrFind & " And A.生产日期 Between [15] And [16] "
        mstrOthers(9) = Format(dtp开始时间(2), "yyyy-mm-dd")
        mstrOthers(10) = Format(dtp结束时间(2), "yyyy-mm-dd")
    End If
    mstrOthers(11) = Trim(Txt开始发票号.Text)
    mstrOthers(12) = Trim(Txt结束发票号.Text)
    If Trim(Txt开始发票号.Text) <> "" Or Trim(Txt结束发票号.Text) <> "" Then
         mstrFind = mstrFind & "   And Exists(Select 1 From 应付记录 D Where a.Id=d.收发ID And  D.系统标识=5 And D.记录性质=0 "
        If Me.Txt开始发票号 <> "" And Me.Txt结束发票号 <> "" Then mstrFind = mstrFind & " And d.发票号 >= [17] And d.发票号 <=[18] "
        If Me.Txt开始发票号 <> "" And Me.Txt结束发票号 = "" Then mstrFind = mstrFind & " And d.发票号 >=[17] "
        If Me.Txt开始发票号 = "" And Me.Txt结束发票号 <> "" Then mstrFind = mstrFind & " And d.发票号 <=[18] "
        mstrFind = mstrFind & ")"
    End If
    If gblnCode = True And Trim(txt条码.Text) <> "" Then
        mstrOthers(13) = UCase(Trim(txt条码.Text))
        mstrFind = mstrFind & " And (A.商品条码 Like [19] Or A.内部条码 Like [19])"
    End If
    
    Unload Me
End Sub

Private Sub Cmd生产商_Click()
    Dim rsTemp As New Recordset
    Dim blnCancel As Boolean
    Dim vRect As RECT
    
    vRect = zlControl.GetControlRect(txt生产商.hwnd)
    
    gstrSQL = "Select rownum as id,null as 上级id,编码,名称,简码,1 as 末级 From 材料生产商 " & _
              "Where (站点=[1] or 站点 is null) "
    
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 1, "卫材生产商选择", True, "", "选择卫材生产商或厂牌", False, False, True, vRect.Left - 15, vRect.Top, txt生产商.Height, blnCancel, False, False, gstrNodeNo)
    
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
    
    If rsTemp Is Nothing Then
        txt生产商.SetFocus
        Exit Sub
    End If
    If rsTemp.State <> 1 Then
        txt生产商.SetFocus
        Exit Sub
    End If
    With rsTemp
        txt生产商.Tag = 1
        txt生产商.Text = zlStr.NVL(!名称)
        chk生产日期.SetFocus
    End With
End Sub

Private Sub Cmd材料_Click()
    Dim RecReturn As Recordset
    
    Set RecReturn = Frm材料选择器.ShowMe(Me, 1, 0, mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex))
    If RecReturn.RecordCount = 0 Then Exit Sub
    Txt材料 = "[" & RecReturn!编码 & "]" & RecReturn!名称
    Txt材料.Tag = RecReturn!材料ID
        
    If Chk供应商.Visible = True Then
        Chk供应商.SetFocus
    End If
End Sub

Private Sub dtp结束时间_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub dtp开始时间_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then Me.dtp结束时间(Index).SetFocus
End Sub


Private Sub Form_Load()
    Me.dtp结束时间(0) = sys.Currentdate
    Me.dtp结束时间(1) = Me.dtp结束时间(0)
    
    Me.dtp开始时间(0) = DateAdd("d", -7, Me.dtp结束时间(0))
    Me.dtp开始时间(1) = Me.dtp开始时间(0)
    
    
    Me.dtp开始时间(2) = DateAdd("d", -7, Me.dtp结束时间(0))
    Me.dtp结束时间(2) = Me.dtp结束时间(0)
    
    lbl条码.Visible = gblnCode
    txt条码.Visible = gblnCode
    
    Me.txt供应商.Tag = 0
    Me.Txt材料.Tag = 0
    Me.txt生产商.Tag = 0
    lng材料ID = 0
    chk包含财务审核.Enabled = False
    chk高值耗材.Enabled = True
    chkNot高值耗材.Enabled = True
    '打开记录集
    sstFilter.Tab = 0
    BlnAdvance = False
    
End Sub

Private Function CheckCompete() As Boolean
    Dim rsCompete As New Recordset
    
    On Error GoTo ErrHandle
    CheckCompete = False
    
    gstrSQL = "" & _
        "   Select id,上级ID,编码,简码,末级,名称 " & _
        "   From 供应商 " & _
        "   Where 名称 is Not NULL " & _
        "       And  (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null) " & _
        "       And (substr(类型,5,1)=1 And (站点=[1] or 站点 is null) Or Nvl(末级,0)=0) " & _
        "   Start with 上级ID is NULL Connect by prior id=上级id"
    Set rsCompete = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, gstrNodeNo)
    
    With rsCompete
        If .EOF Then
            .Close
            MsgBox "卫材供应商信息不全，请在供药单位管理中设置卫材供应商信息！", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    gstrSQL = "Select 编码,名称,简码 From 材料生产商 where (站点=[1] Or 站点 is null) "
    Set rsCompete = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "-材料生产商", gstrNodeNo)
    With rsCompete
        If .EOF Then
            MsgBox "卫材生产商信息不全,请在字典管理中设置卫材生产商信息！", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    CheckCompete = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    If mshSelect.Visible = True Then
        mshSelect.Visible = False
        Select Case mstrSelectTag
            Case "Provider"
                txt供应商.SetFocus
                txt供应商.SelStart = 0
                txt供应商.SelLength = Len(txt供应商.Text)
        
            Case "Maker"
                txt生产商.SetFocus
                txt生产商.SelStart = 0
                txt生产商.SelLength = Len(txt生产商.Text)
            
            Case "Booker"
                Txt填制人.SetFocus
                Txt填制人.SelStart = 0
                Txt填制人.SelLength = Len(Txt填制人.Text)
            Case "Verify"
                Txt审核人.SetFocus
                Txt审核人.SelStart = 0
                Txt审核人.SelLength = Len(Txt审核人.Text)
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
                    txt供应商.Text = .TextMatrix(.Row, 3)
                    txt供应商.Tag = .TextMatrix(.Row, 0)
                    Chk生产商.SetFocus
                Case "Maker"
                    txt生产商.Text = .TextMatrix(.Row, 1)
                    txt生产商.Tag = 1
                    chk生产日期.SetFocus
                Case "Booker"
                    Txt填制人 = .TextMatrix(.Row, 2)
                    Txt审核人.SetFocus
                Case "Verify"
                    Txt审核人 = .TextMatrix(.Row, 2)
                    Txt开始发票号.SetFocus
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
            BlnAdvance = True
        End If
    End With
End Sub

Private Sub sstFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 Or KeyCode = 13 Then
        If sstFilter.Tab = 0 Then
            txt开始No.SetFocus
        Else
            Chk材料.SetFocus
        End If
    End If
End Sub

Private Sub sstFilter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 9 Or KeyAscii = 13 Then
        If sstFilter.Tab = 0 Then
            txt开始No.SetFocus
        Else
            Chk材料.SetFocus
        End If
    End If
    
End Sub

Private Sub txt供应商_GotFocus()
'    Tvw.Visible = False
End Sub

Private Sub txt供应商_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim recTmp As New Recordset
    
    On Error GoTo ErrHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    If LTrim(RTrim(txt供应商)) <> "" Then
        txt供应商 = UCase(txt供应商)
        
        gstrSQL = "" & _
            "   Select id,编码,简码,名称 " & _
            "   From 供应商 " & _
            "   where  末级=1 And (站点=[2] or 站点 is null) " & _
            "           And (substr(类型,5,1)=1 Or Nvl(末级,0)=0) " & _
            "           And (编码 like [1] or 简码 like [1] or 名称 like [1])"
        
        Set recTmp = zlDatabase.OpenSQLRecord(gstrSQL, "卫材供应商", IIf(gstrMatchMethod = "0", "%", "") & txt供应商.Text & "%", gstrNodeNo)
        With recTmp
            If .EOF Then
                MsgBox "输入值无效！", vbInformation, gstrSysName
                KeyCode = 0
                txt供应商.Tag = 0
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Provider"
                Set mshSelect.Recordset = recTmp
                With mshSelect
                    .Top = sstFilter.Top + fra附加条件.Top + txt供应商.Top + txt供应商.Height
                    .Left = sstFilter.Left + fra附加条件.Left + txt供应商.Left
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 0
                    .ColWidth(1) = 800
                    .ColWidth(2) = 800
                    .ColWidth(3) = .Width - .ColWidth(1) - .ColWidth(2)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                End With
            Else
                txt供应商 = !名称
                txt供应商.Tag = !Id
            End If
        End With
    End If
    
    If Chk生产商.Value = 1 Then
        txt生产商.SetFocus
    Else
        Chk生产商.SetFocus
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt供应商_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt结束NO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng库房ID As Long
    
    lng库房ID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    If KeyCode = vbKeyReturn Then
        If Len(txt结束NO) < 8 And Len(txt结束NO) > 0 Then
            txt结束NO.Text = zlCommFun.GetFullNO(txt结束NO.Text, 68, lng库房ID)
        End If
        OS.PressKey (vbKeyTab)
    End If

End Sub

Private Sub txt结束NO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt结束发票号_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Me.cmd确定.SetFocus
End Sub

Private Sub txt开始No_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng库房ID As Long
    
    lng库房ID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    If KeyCode = vbKeyReturn Then
        If Len(txt开始No) < 8 And Len(txt开始No) > 0 Then
            txt开始No.Text = zlCommFun.GetFullNO(txt开始No.Text, 68, lng库房ID)
        End If
        OS.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txt开始No_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt开始发票号_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Txt结束发票号.SetFocus
End Sub

Private Sub Txt审核人_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyReturn Then Txt开始发票号.SetFocus
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt审核人.Text) = "" Then
            Txt开始发票号.SetFocus
            Exit Sub
        End If
        Txt审核人.Text = UCase(Txt审核人.Text)
        
        gstrSQL = "" & _
            "   Select 编号,简码,姓名 " & _
            "   From 人员表 " & _
            "   Where (姓名 like [1] or 编号 like [1] or 简码 like [1] ) " & _
            "       and (站点=[2] or 站点 is null) " & _
            "       and (撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)" & _
            "   order by 编号"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取审核人", IIf(gstrMatchMethod = "0", "%", "") & Me.Txt审核人 & "%", gstrNodeNo)
        
        With rsTemp
            
            If .EOF Then
                MsgBox "输入值无效！", vbInformation, gstrSysName
                KeyCode = 0
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Verify"
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = sstFilter.Top + fra附加条件.Top + Txt审核人.Top - .Height ' + Txt审核人.Height
                    .Left = sstFilter.Left + fra附加条件.Left + Txt审核人.Left + Txt审核人.Width - .Width
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 800
                    .ColWidth(1) = 800
                    .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                Txt审核人 = IIf(IsNull(!姓名), "", !姓名)
                Txt开始发票号.SetFocus
            End If
        End With
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    
End Sub

Private Sub Txt审核人_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt生产商_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strKey As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    If KeyCode = vbKeyReturn Then
        If Me.txt生产商 = "" Then Exit Sub
        If Trim(txt生产商) = "" Then Exit Sub
        
        strKey = GetMatchingSting(txt生产商.Text, False)
        gstrSQL = "" & _
            "   Select 编码,名称,简码 " & _
            "   From 材料生产商 " & _
            "   Where (站点=[2] or 站点 is null) " & _
            "         and (名称 like [1] or 编码 like upper([1]) or 简码 like upper([1]))"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strKey, gstrNodeNo)
        
        With rsTemp
            If .EOF Then
                MsgBox "输入值无效！", vbInformation, gstrSysName
                KeyCode = 0
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Maker"
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = sstFilter.Top + fra附加条件.Top + txt生产商.Top + txt生产商.Height
                    .Left = sstFilter.Left + fra附加条件.Left + txt生产商.Left
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 800
                    .ColWidth(1) = 800
                    .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                    
                End With
            Else
                txt生产商 = IIf(IsNull(!名称), "", !名称)
                
                txt生产商.Tag = 1
                chk生产日期.SetFocus
            End If
        End With
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt填制人_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt填制人.Text) = "" Then
            Txt审核人.SetFocus
            Exit Sub
        End If
        Txt填制人.Text = UCase(Txt填制人.Text)
        
        gstrSQL = "" & _
            "   Select 编号,简码,姓名 " & _
            "   From 人员表 " & _
            "   Where (姓名 like [1] or 编号 like [1] or 简码 like [1] ) " & _
            "       and (站点=[2] or 站点 is null) " & _
            "       and (撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)" & _
            "   order by 编号"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, IIf(gstrMatchMethod = "0", "%", "") & Me.Txt填制人 & "%", gstrNodeNo)
        
        With rsTemp
            If .EOF Then
                MsgBox "输入值无效！", vbInformation, gstrSysName
                KeyCode = 0
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Booker"
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = sstFilter.Top + fra附加条件.Top + Txt填制人.Top - .Height '+ Txt填制人.Height
                    .Left = sstFilter.Left + fra附加条件.Left + Txt填制人.Left
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 800
                    .ColWidth(1) = 800
                    .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    Exit Sub
                End With
            Else
                Txt填制人 = IIf(IsNull(!姓名), "", !姓名)
                Me.Txt审核人.SetFocus
            End If
        End With
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt材料_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim RecReturn As Recordset
    Dim strKey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(Txt材料.Text) = "" Then Exit Sub
    sngLeft = Me.Left + sstFilter.Left + fra附加条件.Left + Txt材料.Left
    sngTop = Me.Top + sstFilter.Top + fra附加条件.Top + Txt材料.Top + Txt材料.Height + Me.Height - Me.ScaleHeight '  50
    If sngTop + 3630 > Screen.Height Then
        sngTop = sngTop - Txt材料.Height - 3630
    End If
    
    strKey = Trim(Txt材料.Text)
    If Mid(strKey, 1, 1) = "[" Then
        If InStr(2, strKey, "]") <> 0 Then
            strKey = Mid(strKey, 2, InStr(2, strKey, "]") - 2)
        Else
            strKey = Mid(strKey, 2)
        End If
    End If
    
    Set RecReturn = FrmMulitSel.ShowSelect(Me, 1, , mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), strKey, sngLeft, sngTop, Txt材料.Width, Txt材料.Height)
    If RecReturn.RecordCount = 0 Then Exit Sub
    Txt材料 = "[" & RecReturn!编码 & "]" & RecReturn!名称
    Txt材料.Tag = RecReturn!材料ID
    
    If Chk供应商.Visible = True Then
        If Chk供应商.Value = 1 Then
            txt供应商.SetFocus
        Else
            Chk供应商.SetFocus
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

