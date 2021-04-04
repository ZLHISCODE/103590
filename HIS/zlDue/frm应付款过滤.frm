VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm应付款过滤 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "应付款过滤设置"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd取消 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6255
      TabIndex        =   29
      Top             =   840
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6270
      TabIndex        =   28
      Top             =   405
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   2535
      Left            =   750
      TabIndex        =   30
      Top             =   4485
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4471
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
      Height          =   3960
      Left            =   105
      TabIndex        =   31
      Top             =   135
      Width           =   5970
      _ExtentX        =   10530
      _ExtentY        =   6985
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "范围(&R)"
      TabPicture(0)   =   "frm应付款过滤.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra范围"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chkDept(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkDept(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkDept(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkDept(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkDept(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "附加条件(&F)"
      TabPicture(1)   =   "frm应付款过滤.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra附加条件"
      Tab(1).ControlCount=   1
      Begin VB.CheckBox chkDept 
         Caption         =   "卫材(&W)"
         Height          =   195
         Index           =   4
         Left            =   4785
         TabIndex        =   34
         Tag             =   "4"
         Top             =   3450
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkDept 
         Caption         =   "其他(&Q)"
         Height          =   195
         Index           =   3
         Left            =   3720
         TabIndex        =   18
         Tag             =   "4"
         Top             =   3450
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkDept 
         Caption         =   "药品(&D)"
         Height          =   195
         Index           =   0
         Left            =   450
         TabIndex        =   15
         Tag             =   "1"
         Top             =   3450
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkDept 
         Caption         =   "物资(&M)"
         Height          =   195
         Index           =   1
         Left            =   1545
         TabIndex        =   16
         Tag             =   "2"
         Top             =   3450
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkDept 
         Caption         =   "设备(&S)"
         Height          =   195
         Index           =   2
         Left            =   2685
         TabIndex        =   17
         Tag             =   "4"
         Top             =   3450
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.Frame fra范围 
         Height          =   2685
         Left            =   240
         TabIndex        =   33
         Top             =   600
         Width           =   5520
         Begin VB.CheckBox chkStrike 
            Caption         =   "包含冲销(&K)"
            Enabled         =   0   'False
            Height          =   300
            Left            =   720
            TabIndex        =   14
            Top             =   2280
            Width           =   1440
         End
         Begin VB.CheckBox chk审核 
            Caption         =   "已审核单据(&V)"
            Height          =   270
            Left            =   480
            TabIndex        =   9
            Top             =   1560
            Width           =   2070
         End
         Begin VB.CheckBox chk填制 
            Caption         =   "未审核单据(&W)"
            Height          =   240
            Left            =   480
            TabIndex        =   4
            Top             =   840
            Value           =   1  'Checked
            Width           =   1725
         End
         Begin VB.TextBox txt结束NO 
            Height          =   300
            Left            =   2970
            MaxLength       =   8
            TabIndex        =   3
            Top             =   360
            Width           =   1605
         End
         Begin VB.TextBox txt开始No 
            Height          =   300
            Left            =   840
            MaxLength       =   8
            TabIndex        =   1
            Top             =   360
            Width           =   1605
         End
         Begin MSComCtl2.DTPicker dtp开始时间 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   6
            Top             =   1080
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   315949059
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp结束时间 
            Height          =   315
            Index           =   0
            Left            =   3585
            TabIndex        =   8
            Top             =   1080
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   315949059
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp开始时间 
            Height          =   315
            Index           =   1
            Left            =   1680
            TabIndex        =   11
            Top             =   1845
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   315949059
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp结束时间 
            Height          =   315
            Index           =   1
            Left            =   3585
            TabIndex        =   13
            Top             =   1845
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   315949059
            CurrentDate     =   36263
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   0
            Left            =   3345
            TabIndex        =   7
            Top             =   1140
            Width           =   180
         End
         Begin VB.Label lbl时间 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "填制日期(&N)"
            Height          =   180
            Index           =   0
            Left            =   630
            TabIndex        =   5
            Top             =   1140
            Width           =   990
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   3
            Left            =   3345
            TabIndex        =   12
            Top             =   1905
            Width           =   180
         End
         Begin VB.Label lbl时间 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "审核日期(&E)"
            Height          =   180
            Index           =   1
            Left            =   630
            TabIndex        =   10
            Top             =   1905
            Width           =   990
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   1
            Left            =   2640
            TabIndex        =   2
            Top             =   420
            Width           =   180
         End
         Begin VB.Label LblNO 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "N&o"
            Height          =   180
            Left            =   480
            TabIndex        =   0
            Top             =   420
            Width           =   180
         End
      End
      Begin VB.Frame fra附加条件 
         Height          =   2715
         Left            =   -74760
         TabIndex        =   32
         Top             =   585
         Width           =   5505
         Begin VB.ComboBox cboStore 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   940
            Width           =   3615
         End
         Begin VB.CheckBox chkStore 
            Caption         =   "库房(&W)"
            Height          =   300
            Left            =   435
            TabIndex        =   22
            Top             =   940
            Width           =   975
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   25
            Top             =   1320
            Width           =   3570
         End
         Begin VB.TextBox Txt审核人 
            Height          =   300
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   27
            Top             =   2070
            Width           =   1365
         End
         Begin VB.TextBox Txt填制人 
            Height          =   300
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   26
            Top             =   1695
            Width           =   1365
         End
         Begin VB.CommandButton Cmd供应商 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4890
            TabIndex        =   21
            Top             =   540
            Width           =   255
         End
         Begin VB.TextBox txt供应商 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   20
            Top             =   540
            Width           =   3375
         End
         Begin VB.CheckBox Chk供应商 
            Caption         =   "供应商(&P)"
            Height          =   300
            Left            =   435
            TabIndex        =   19
            Top             =   540
            Width           =   1215
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "随货单号(&S)"
            Height          =   180
            Left            =   480
            TabIndex        =   24
            Top             =   1380
            Width           =   990
         End
         Begin VB.Label Lbl填制人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "填制人(&T)"
            Height          =   180
            Left            =   660
            TabIndex        =   36
            Top             =   1770
            Width           =   810
         End
         Begin VB.Label Lbl审核人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "审核人(&V)"
            Height          =   180
            Left            =   660
            TabIndex        =   35
            Top             =   2145
            Width           =   810
         End
      End
   End
End
Attribute VB_Name = "frm应付款过滤"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFind As String  '查找字符串
Private mblnAdvance As Boolean '是否展开
Private mdtStart As Date   '开始时间
Private mdtEnd As Date     '结束时间
Private mdtVerifyStart As Date
Private mdtVerifyEnd As Date
Private mstrSelectTag As String     '当前选择的对象
Private mstr类型 As String
Private mstrPrivs As String
Private mcllFilter As Collection

Public Function GetSearch(ByVal FrmMain As Form, ByVal strPrivs As String, _
        ByRef dtStart As Date, ByRef dtEnd As Date, _
        ByRef dtVerifyStart As Date, ByRef dtVerifyEnd As Date, ByRef str类型 As String, ByRef cllFilter As Collection) As String
    '--------------------------------------------------------------
    '功能：获取检索所需的SQL语句
    '参数：FrmMain----------调用窗体
    '      dtStart---------起始日期
    '      dtEnd-----------结束日期
    '      dtVerifyStart---审核起始日期
    '      dtVerifyEnd-----审核结束日期
    '返回：SQL语句
    '说明：
    '--------------------------------------------------------------
    mstrFind = "": mstrPrivs = strPrivs
    If Not CheckCompete Then Exit Function
    Call 权限设置
    Me.Show vbModal, FrmMain
    GetSearch = mstrFind
    dtStart = mdtStart
    dtEnd = mdtEnd
    dtVerifyStart = mdtVerifyStart
    dtVerifyEnd = mdtVerifyEnd
    str类型 = mstr类型
    Set cllFilter = mcllFilter
End Function

Public Sub 权限设置()
    If Check相关权限(gstrPrivs, "药品") = False Then
        chkDept(0).Enabled = False
        chkDept(0).Value = 0
    End If
    If Check相关权限(gstrPrivs, "物资") = False Then
        chkDept(1).Enabled = False
        chkDept(1).Value = 0
    End If
    
    If Check相关权限(gstrPrivs, "设备") = False Then
        chkDept(2).Enabled = False
        chkDept(2).Value = 0
    End If
    If Check相关权限(gstrPrivs, "其他") = False Then
        chkDept(3).Enabled = False
        chkDept(3).Value = 0
    End If
    If Check相关权限(gstrPrivs, "卫材") = False Then
        chkDept(4).Enabled = False
        chkDept(4).Value = 0
    End If
End Sub

Private Sub chkDept_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlcommfun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub chkStore_Click()
    cboStore.Enabled = chkStore.Value = 1
End Sub

Private Sub chkStrike_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlcommfun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub Chk供应商_Click()
    txt供应商.Enabled = IIf(Chk供应商.Value = 1, True, False)
    Cmd供应商.Enabled = IIf(Chk供应商.Value = 1, True, False)
End Sub

Private Sub Chk供应商_GotFocus()
    If sstFilter.Tab = 0 Then
        sstFilter.Tab = 1
    End If
    Chk供应商.SetFocus
End Sub

Private Sub Chk供应商_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlcommfun.PressKey (vbKeyTab)
    End If

End Sub

Private Sub chk审核_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlcommfun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub chk填制_Click()
    dtp开始时间(0).Enabled = IIf(chk填制.Value = 1, True, False)
    dtp结束时间(0).Enabled = IIf(chk填制.Value = 1, True, False)
End Sub

Private Sub chk审核_Click()
    dtp开始时间(1).Enabled = IIf(chk审核.Value = 1, True, False)
    dtp结束时间(1).Enabled = IIf(chk审核.Value = 1, True, False)
    chkStrike.Enabled = IIf(chk审核.Value = 1, True, False)
End Sub

Private Sub chk填制_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlcommfun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub Cmd供应商_Click()
    Dim strTemp As String
    
    strTemp = frm供应商选择.SelDept(mstrPrivs)
    If strTemp = "" Then Exit Sub
    txt供应商.Tag = Left(strTemp, InStr(strTemp, ",") - 1)
    txt供应商.Text = Mid(strTemp, InStr(strTemp, ",") + 1)
    Unload frm供应商选择
End Sub

Private Sub Cmd取消_Click()
    mstrFind = ""
    Unload Me
End Sub

Private Sub Cmd确定_Click()
    '检查数据
    If Chk供应商.Value = 1 Then
        '问题29757 by lesfeng 2010-05-10
        If Val(txt供应商.Tag) = 0 Then
            MsgBox "请选择需查询的供应商信息！", vbInformation, gstrSysName
            Me.txt供应商.SetFocus
            Exit Sub
        End If
    End If
    
    If chk填制.Value = 0 And chk审核.Value = 0 Then
        MsgBox "对不起，必须选择一个填制日期或者审核日期!", vbInformation, gstrSysName
        chk填制.SetFocus
        Exit Sub
    End If

    mstrFind = ""
    '生成SQL查询条件语句
    Dim intTemp As Integer
    
    'by lesfeng 2010-1-7 性能优化
    Set mcllFilter = New Collection
    mcllFilter.Add Array("1901-01-01 00:00:00", "1901-01-01 00:00:00"), "填制日期"
    mcllFilter.Add Array("1901-01-01 00:00:00", "1901-01-01 00:00:00"), "审核日期"
    mcllFilter.Add Array("", ""), "单据号"
    mcllFilter.Add "", "随货单号"
    mcllFilter.Add "", "供应商id"
    mcllFilter.Add "", "库房ID"
    mcllFilter.Add "", "填制人"
    mcllFilter.Add "", "审核人"
            
    If chk填制.Value = 1 And chk审核.Value = 1 Then
        If chkStrike.Value = 1 Then
            mstrFind = " And ((A.填制日期 Between [1] And [2]) or (A.审核日期 Between [3] And [4]))"
        Else
            mstrFind = " And ((A.填制日期 Between [1] And [2]) or (A.审核日期 Between [3] And [4])) and a.记录状态 =1 "
        End If
        
        mcllFilter.Remove "填制日期"
        mcllFilter.Remove "审核日期"
        mcllFilter.Add Array(Format(dtp开始时间(0), "yyyy-mm-dd") & " 00:00:00", Format(dtp结束时间(0), "yyyy-mm-dd") & " 23:59:59"), "填制日期"
        mcllFilter.Add Array(Format(dtp开始时间(1), "yyyy-mm-dd") & " 00:00:00", Format(dtp结束时间(1), "yyyy-mm-dd") & " 23:59:59"), "审核日期"
        
        mdtStart = Format(dtp开始时间(0), "yyyy-mm-dd")
        mdtEnd = Format(dtp结束时间(0), "yyyy-mm-dd")
                
        mdtVerifyStart = Format(dtp开始时间(1), "yyyy-mm-dd")
        mdtVerifyEnd = Format(dtp结束时间(1), "yyyy-mm-dd")
        
    ElseIf chk审核.Value = 1 Then
        If chkStrike.Value = 1 Then
            mstrFind = " And A.审核日期 Between [3] And [4] "
        Else
            mstrFind = " And A.审核日期 Between [3] And [4] and a.记录状态 =1 "
        End If
        mcllFilter.Remove "审核日期"
        mcllFilter.Add Array(Format(dtp开始时间(1), "yyyy-mm-dd") & " 00:00:00", Format(dtp结束时间(1), "yyyy-mm-dd") & " 23:59:59"), "审核日期"
        
        mdtVerifyStart = Format(dtp开始时间(1), "yyyy-mm-dd")
        mdtVerifyEnd = Format(dtp结束时间(1), "yyyy-mm-dd")
        mdtStart = Format("1901-01-01", "yyyy-mm-dd")
        mdtEnd = Format("1901-01-01", "yyyy-mm-dd")
    ElseIf chk填制.Value = 1 Then
        mstrFind = " And (A.填制日期 Between [1] And [2]) and 审核日期 is null "
        mcllFilter.Remove "填制日期"
        mcllFilter.Add Array(Format(dtp开始时间(0), "yyyy-mm-dd") & " 00:00:00", Format(dtp结束时间(0), "yyyy-mm-dd") & " 23:59:59"), "填制日期"
        
        mdtStart = Format(dtp开始时间(0), "yyyy-mm-dd")
        mdtEnd = Format(dtp结束时间(0), "yyyy-mm-dd")
        
        mdtVerifyStart = Format("1901-01-01", "yyyy-mm-dd")
        mdtVerifyEnd = Format("1901-01-01", "yyyy-mm-dd")
    End If
    
    Dim intYear As Integer, strYear As String
    
    If Len(txt开始No) < 8 And Len(txt开始No) > 0 Then
        Me.txt开始No = UCase(LTrim(Me.txt开始No))
        intYear = Format(zlDatabase.Currentdate, "YYYY") - 1990
        strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
        If Len(txt开始No) < 8 Then Me.txt开始No = strYear & String(7 - Len(txt开始No), "0") & Me.txt开始No
    End If
    If Len(txt结束NO) < 8 And Len(txt结束NO) > 0 Then
        Me.txt结束NO = UCase(LTrim(Me.txt结束NO))
        intYear = Format(zlDatabase.Currentdate, "YYYY") - 1990
        strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
        If Len(txt结束NO) < 8 Then Me.txt结束NO = strYear & String(7 - Len(txt结束NO), "0") & Me.txt结束NO
    End If
    
    If Me.txt开始No <> "" And Me.txt结束NO <> "" Then mstrFind = mstrFind & " And A.No >= [5] And A.No <= [6]"
    If Me.txt开始No <> "" And Me.txt结束NO = "" Then mstrFind = mstrFind & " And A.No >= [5]"
    If Me.txt开始No = "" And Me.txt结束NO <> "" Then mstrFind = mstrFind & " And A.No <= [6]"
    
    mcllFilter.Remove "单据号"
    mcllFilter.Add Array(Trim(Me.txt开始No), Trim(Me.txt结束NO)), "单据号"
 
    Dim strTemp As String
    
    Dim intIndex As Integer
    For intIndex = 0 To 4
        If chkDept(intIndex).Value = 1 Then
            strTemp = strTemp & "1"
        Else
            strTemp = strTemp & "0"
        End If
    Next
    mstr类型 = strTemp ' Bin2Dec(strTemp)
    
    '扩展查询条件
    If mblnAdvance = False Then
        Unload Me
        Exit Sub
    End If
    If Trim(txtEdit.Text) <> "" Then mstrFind = mstrFind & " And A.随货单号 like [7]"
    
    mcllFilter.Remove "随货单号"
    mcllFilter.Add GetMatchingSting(txtEdit), "随货单号"
    
    If Chk供应商.Value = 1 Then
        mstrFind = mstrFind & " and a.单位ID = [8]"
        mcllFilter.Remove "供应商id"
        mcllFilter.Add txt供应商.Tag, "供应商id"
    End If
    
    If Me.Txt填制人 <> "" Then mstrFind = mstrFind & " And A.填制人 like [9]"
    If Me.Txt审核人 <> "" Then mstrFind = mstrFind & " And A.审核人 like [10]"
    
    If chkStore.Value = 1 Then
        mstrFind = mstrFind & " and a.库房ID = [11] "
        mcllFilter.Remove "库房ID"
        If cboStore.ListIndex = -1 Then
            mcllFilter.Add "", "库房ID"
        Else
            mcllFilter.Add cboStore.ItemData(cboStore.ListIndex), "库房ID"
        End If
    End If
    
    mcllFilter.Remove "填制人"
    mcllFilter.Add GetMatchingSting(Txt填制人), "填制人"
    mcllFilter.Remove "审核人"
    mcllFilter.Add GetMatchingSting(Txt审核人), "审核人"
    Unload Me
End Sub

Private Sub dtp结束时间_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        SendKeys vbTab
    End If
End Sub

Private Sub dtp开始时间_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then
        zlcommfun.PressKey (vbKeyTab)
     End If
End Sub

Private Sub Form_Load()
    Me.dtp结束时间(0) = zlDatabase.Currentdate
    Me.dtp结束时间(1) = Me.dtp结束时间(0)
    
    Me.dtp开始时间(0) = DateAdd("d", -7, Me.dtp结束时间(0))
    Me.dtp开始时间(1) = Me.dtp开始时间(0)
    
    Me.txt供应商.Tag = 0
    '打开记录集
    sstFilter.Tab = 0
    mblnAdvance = False
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
    gstrSQL = "Select id From 供应商 Where (撤档时间 is null or To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01') and  末级=1 " & zl_获取站点限制 & " and rownum<=2 "
    zlDatabase.OpenRecordset rsCompete, gstrSQL, Me.Caption
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
                txt供应商.SelStart = 0
                txt供应商.SelLength = Len(txt供应商.Text)
            
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
                    Txt填制人.SetFocus
                Case "Booker"
                    Txt填制人 = .TextMatrix(.Row, 2)
                    Txt审核人.SetFocus
                Case "Verify"
                    Txt审核人 = .TextMatrix(.Row, 2)
                    cmd确定.SetFocus
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
    Dim rsTmp As ADODB.Recordset

    With sstFilter
        If .Tab = 1 Then
            mblnAdvance = True
        End If
        
        cboStore.Clear
        
        Set rsTmp = GetStoreInfo("'H', 'I', 'J', 'K', 'L', 'M', 'R', 'T', 'V', 'S'")
        If rsTmp Is Nothing Then Exit Sub
        If rsTmp.State <> adStateOpen Then Exit Sub
        
        If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
        Do While rsTmp.EOF = False
            cboStore.AddItem "[" & rsTmp!编码 & "]" & rsTmp!名称
            cboStore.ItemData(cboStore.NewIndex) = rsTmp!ID
            
            rsTmp.MoveNext
        Loop
        rsTmp.Close
        
    End With
End Sub

Private Sub sstFilter_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 Or KeyCode = 13 Then
        If sstFilter.Tab = 0 Then
            txt开始No.SetFocus
        Else
            Chk供应商.SetFocus
        End If
    End If
End Sub

Private Sub sstFilter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 9 Or KeyAscii = 13 Then
        If sstFilter.Tab = 0 Then
            txt开始No.SetFocus
        Else
            Chk供应商.SetFocus
        End If
    End If
End Sub
 
Private Sub txtEdit_GotFocus()
    zlControl.TxtSelAll txtEdit
    zlcommfun.OpenIme False
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txtEdit, KeyAscii, m文本式)
End Sub

Private Sub txt供应商_GotFocus()
    SetTxtGotFocus txt供应商, True
End Sub

Private Sub txt供应商_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As New Recordset
    Dim strTemp As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    On Error GoTo errHandle
    If LTrim(RTrim(txt供应商)) <> "" Then
        txt供应商 = UCase(txt供应商)
        strTemp = GetMatchingSting(txt供应商)
        Dim str权限 As String
        
        str权限 = " and " & Get分类权限(gstrPrivs)
            
        gstrSQL = "" & _
            "   Select id,编码,简码,名称 " & _
            "   From 供应商 " & _
            "   Where  (撤档时间 is null or  To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01') " & _
            "       " & zl_获取站点限制 & " And 末级=1" & _
            "     And (编码 like [1] or 简码 like [1] or 名称 like [1]) " & str权限
              
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strTemp)
        If rsTemp.EOF Then
            MsgBox "无此条件的供应商！", vbInformation, gstrSysName
            KeyCode = 0
            txt供应商.Tag = 0
            txt供应商.SelStart = 0
            txt供应商.SelLength = Len(txt供应商.Text)
            
            Exit Sub
        End If
        If rsTemp.RecordCount > 1 Then
            mstrSelectTag = "Provider"
            Set mshSelect.Recordset = rsTemp
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
            txt供应商 = rsTemp!名称
            txt供应商.Tag = rsTemp!ID
        End If
    
    End If
    Txt填制人.SetFocus
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub txt供应商_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt结束NO_GotFocus()
    SetTxtGotFocus txt结束NO, False
End Sub

Private Sub txt结束NO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Len(txt结束NO) < 8 And Len(txt结束NO) > 0 Then
            Dim intYear  As Integer, strYear As String
            Me.txt结束NO = UCase(LTrim(Me.txt结束NO))
            intYear = Format(zlDatabase.Currentdate, "YYYY") - 1990
            strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
            If Len(txt结束NO) < 8 Then Me.txt结束NO = strYear & String(7 - Len(txt结束NO), "0") & Me.txt结束NO
        End If
        zlcommfun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txt结束NO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt开始No_GotFocus()
      SetTxtGotFocus txt开始No, False
End Sub

Private Sub txt开始No_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Len(txt开始No) < 8 And Len(txt开始No) > 0 Then
            Dim intYear  As Integer, strYear As String
            Me.txt开始No = UCase(LTrim(Me.txt开始No))
            intYear = Format(zlDatabase.Currentdate, "YYYY") - 1990
            strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
            If Len(txt开始No) < 8 Then Me.txt开始No = strYear & String(7 - Len(txt开始No), "0") & Me.txt开始No
        End If
        zlcommfun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txt开始No_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt审核人_GotFocus()
    SetTxtGotFocus Txt审核人, True
End Sub

Private Sub Txt审核人_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then cmd确定.SetFocus
    
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt审核人.Text) = "" Then
            cmd确定.SetFocus
            Exit Sub
        End If
        Txt审核人.Text = UCase(Txt审核人.Text)
        
        gstrSQL = "" & _
             "   Select 编号,简码,姓名 " & _
             "   From 人员表 " & _
             "   Where (姓名 like [1] or 编号 like [1] or 简码 like [1] )  " & zl_获取站点限制 & " " & _
             "       and (撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)" & _
             "   order by 编号"
         Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取审核人", IIf(gstrMatchMethod = "0", "%", "") & Me.Txt审核人 & "%")
        
        With rsTemp
            
            If .EOF Then
                MsgBox "输入值无效！", vbInformation, gstrSysName
                KeyCode = 0
                Txt审核人.SelStart = 0
                Txt审核人.SelLength = Len(Txt审核人.Text)
                
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Verify"
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = sstFilter.Top + fra附加条件.Top + Txt审核人.Top + Txt审核人.Height
                    .Left = sstFilter.Left + fra附加条件.Left + Txt审核人.Left
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
                cmd确定.SetFocus
            End If
        End With
    End If
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Txt审核人_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt审核人_LostFocus()
    ImeLanguage False
End Sub

Private Sub Txt填制人_GotFocus()
    SetTxtGotFocus Txt填制人, True
End Sub

Private Sub Txt填制人_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyReturn Then Me.Txt审核人.SetFocus
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt填制人.Text) = "" Then
            Txt审核人.SetFocus
            Exit Sub
        End If
        
        Txt填制人.Text = UCase(Txt填制人.Text)
        gstrSQL = "" & _
             "   Select 编号,简码,姓名 " & _
             "   From 人员表 " & _
             "   Where (姓名 like [1] or 编号 like [1] or 简码 like [1] ) " & zl_获取站点限制 & "" & _
             "       and (撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)" & _
             "   order by 编号"
         Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取填制人", IIf(gstrMatchMethod = "0", "%", "") & Me.Txt填制人 & "%")
          
        With rsTemp
            If .EOF Then
                MsgBox "输入值无效！", vbInformation, gstrSysName
                KeyCode = 0
                Txt填制人.SelStart = 0
                Txt填制人.SelLength = Len(Txt填制人.Text)
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Booker"
                Set mshSelect.Recordset = rsTemp
                With mshSelect
                    .Top = sstFilter.Top + fra附加条件.Top + Txt填制人.Top + Txt填制人.Height
                    .Left = sstFilter.Left + fra附加条件.Left + Txt填制人.Left
                    If .Height > ScaleHeight - .Top Then
                        .Height = ScaleHeight - .Top - 20
                    Else
                        .Height = 2535
                    End If
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
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt填制人_LostFocus()
    ImeLanguage False
End Sub
