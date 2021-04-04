VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmOtherInputSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤"
   ClientHeight    =   4200
   ClientLeft      =   3150
   ClientTop       =   3165
   ClientWidth     =   7515
   Icon            =   "frmOtherInputSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   2055
      Left            =   1320
      TabIndex        =   27
      Top             =   4080
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
      Height          =   3975
      Left            =   120
      TabIndex        =   28
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7011
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "范围(&R)"
      TabPicture(0)   =   "frmOtherInputSearch.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra范围"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "附加条件(&D)"
      TabPicture(1)   =   "frmOtherInputSearch.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra附加条件"
      Tab(1).ControlCount=   1
      Begin VB.Frame fra附加条件 
         Height          =   3300
         Left            =   -74760
         TabIndex        =   36
         Top             =   600
         Width           =   5505
         Begin VB.TextBox txt条码 
            Height          =   300
            Left            =   1650
            TabIndex        =   37
            Top             =   2760
            Width           =   3525
         End
         Begin VB.CheckBox chk生产日期 
            Caption         =   "生产日期"
            Height          =   300
            Left            =   480
            TabIndex        =   17
            Top             =   1875
            Width           =   1095
         End
         Begin VB.CheckBox Chk类别 
            Caption         =   "类别"
            Height          =   300
            Left            =   480
            TabIndex        =   15
            Top             =   1350
            Width           =   960
         End
         Begin VB.CommandButton Cmd材料 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4890
            TabIndex        =   11
            Top             =   360
            Width           =   255
         End
         Begin VB.TextBox Txt材料 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            MaxLength       =   50
            TabIndex        =   10
            Top             =   360
            Width           =   3255
         End
         Begin VB.CheckBox Chk材料 
            Caption         =   "卫生材料"
            Height          =   315
            Left            =   480
            TabIndex        =   9
            Top             =   360
            Width           =   1065
         End
         Begin VB.TextBox Txt填制人 
            Height          =   300
            Left            =   1650
            MaxLength       =   8
            TabIndex        =   22
            Top             =   2355
            Width           =   1365
         End
         Begin VB.TextBox Txt审核人 
            Height          =   300
            Left            =   3780
            MaxLength       =   8
            TabIndex        =   24
            Top             =   2355
            Width           =   1365
         End
         Begin VB.ComboBox Cbo类别 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   1350
            Width           =   3495
         End
         Begin VB.CheckBox Chk生产商 
            Caption         =   "生产商"
            Height          =   300
            Left            =   480
            TabIndex        =   12
            Top             =   855
            Width           =   1155
         End
         Begin VB.TextBox txt生产商 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1650
            TabIndex        =   13
            Top             =   855
            Width           =   3255
         End
         Begin VB.CommandButton Cmd生产商 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4890
            TabIndex        =   14
            Top             =   855
            Width           =   255
         End
         Begin MSComCtl2.DTPicker dtp开始时间 
            Height          =   315
            Index           =   2
            Left            =   1650
            TabIndex        =   18
            Top             =   1875
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   162791427
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp结束时间 
            Height          =   315
            Index           =   2
            Left            =   3540
            TabIndex        =   20
            Top             =   1875
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   162791427
            CurrentDate     =   36263
         End
         Begin VB.Label lbl条码 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "条  码"
            Height          =   180
            Left            =   750
            TabIndex        =   38
            Top             =   2820
            Width           =   540
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   4
            Left            =   3300
            TabIndex        =   19
            Top             =   1935
            Width           =   180
         End
         Begin VB.Label Lbl填制人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "填制人"
            Height          =   180
            Left            =   750
            TabIndex        =   21
            Top             =   2415
            Width           =   540
         End
         Begin VB.Label Lbl审核人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "审核人"
            Height          =   180
            Left            =   3120
            TabIndex        =   23
            Top             =   2415
            Width           =   540
         End
      End
      Begin VB.Frame fra范围 
         Height          =   2850
         Left            =   240
         TabIndex        =   29
         Top             =   600
         Width           =   5520
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
            Height          =   300
            Left            =   720
            TabIndex        =   8
            Top             =   2280
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
            Format          =   162791427
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
            Format          =   162791427
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
            Format          =   162791427
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
            Format          =   162791427
            CurrentDate     =   36263
         End
         Begin VB.Label LblNO 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "No"
            Height          =   180
            Left            =   480
            TabIndex        =   35
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
            TabIndex        =   34
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
            TabIndex        =   33
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   30
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
      TabIndex        =   26
      Top             =   930
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6330
      TabIndex        =   25
      Top             =   435
      Width           =   1100
   End
End
Attribute VB_Name = "FrmOtherInputSearch"
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
Public lng材料ID As Long
Private mstrSelectTag As String     '当前选择的对象
Private mstrOthers(0 To 13) As String ' 0-记录状态(计划类型),1-开始单号,2-结束单号,3-材料id,4-对方部门id(或入出类别id或编制方法(计划用)),5-填制人,6-审核人,7-供应商ID,8-生产商,9-开始生产日期,10-结束生产日期,11-开始发票号,12-结束发票号,13-条码信息

Public Function GetSearch(ByVal frmMain As Form, _
        ByRef datStart As Date, ByRef datEnd As Date, _
        ByRef datVerifyStart As Date, ByRef datVerifyEnd As Date, _
        ByRef strOthers() As String) As String
        
    mstrFind = ""
    Set mfrmMain = frmMain
    If Not CheckCompete Then Exit Function
    
    Me.Show vbModal, mfrmMain
    GetSearch = mstrFind
    datStart = mdatStart
    datEnd = mdatEnd
    datVerifyStart = mdatVerifyStart
    datVerifyEnd = mdatVerifyEnd
    strOthers = mstrOthers
End Function

Private Sub Cbo类别_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub chkStrike_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 9 Or KeyCode = 13 Then
        cmd确定.SetFocus
    End If
    
End Sub

Private Sub chkStrike_KeyPress(KeyAscii As Integer)
    If KeyAscii = 9 Or KeyAscii = 13 Then
        cmd确定.SetFocus
    End If
End Sub



Private Sub Chk类别_Click()
    Cbo类别.Enabled = IIf(Chk类别.Value = 1, True, False)
End Sub

Private Sub Chk类别_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Chk类别.Value = 1 Then
        Cbo类别.SetFocus
    Else
        Txt填制人.SetFocus
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
        ElseIf Chk类别.Visible = True Then
            Chk类别.SetFocus
        Else
            Txt填制人.SetFocus
        End If
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
    Else
        Chk生产商.SetFocus
    End If
End Sub




Private Sub Cmd取消_Click()
    Dim i As Integer
    For i = 0 To 13
        mstrOthers(i) = ""
    Next
    
    mstrFind = ""
    Unload Me
End Sub

Private Sub cmd确定_Click()
    '检查数据
    If Chk材料.Value = 1 Then
        If Txt材料.Tag = 0 Then
            MsgBox "请选择需查询的卫材信息！", vbInformation, gstrSysName
            Me.Txt材料.SetFocus
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
    mdatStart = Format("1901-01- 01", "yyyy-mm-dd")
    mdatEnd = Format("1901-01-01", "yyyy-mm-dd")
    mdatVerifyStart = Format("1901-01-01", "yyyy-mm-dd")
    mdatVerifyEnd = Format("1901-01-01", "yyyy-mm-dd")
    mstrOthers(0) = IIf(chkStrike.Value = 1, "0", "1")
      
    If chk填制.Value = 1 And chk审核.Value = 1 Then
        If chkStrike.Value = 1 Then
            mstrFind = " And ((A.填制日期 Between [2] And [3] and 审核日期 is null) " _
                    & " or (A.审核日期 Between [4] And [5]))"
        Else
            mstrFind = " And ((A.填制日期 Between [2] And [3] and 审核日期 is null) " _
                    & " or (A.审核日期 Between [4] And [5] and a.记录状态 =[6]))  "
        End If
        mdatStart = Format(dtp开始时间(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp结束时间(0), "yyyy-mm-dd")
                
        mdatVerifyStart = Format(dtp开始时间(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp结束时间(1), "yyyy-mm-dd")
        
    ElseIf chk审核.Value = 1 Then
        If chkStrike.Value = 1 Then
            mstrFind = " And A.审核日期 Between [4] And [5] "
        Else
            mstrFind = " And A.审核日期 Between [4] And [5] and a.记录状态 =[6] "
            
        End If
        mdatVerifyStart = Format(dtp开始时间(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp结束时间(1), "yyyy-mm-dd")
    ElseIf chk填制.Value = 1 Then
        mstrFind = " And (A.填制日期 Between [2] And To_Date('" & Format(dtp结束时间(0), "YYYY-mm-dd") & "23:59:59 ','YYYY-MM-DD HH24:MI:SS')) and 审核日期 is null "
        mdatStart = Format(dtp开始时间(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp结束时间(0), "yyyy-mm-dd")
    End If
        
    
    
    Dim intYear As Integer, strYear As String
    
    If Len(txt开始No) < 8 And Len(txt开始No) > 0 Then
        Me.txt开始No = UCase(LTrim(Me.txt开始No))
        intYear = Format(Sys.Currentdate, "YYYY") - 1990
        strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
        If Len(txt开始No) < 8 Then Me.txt开始No = strYear & String(7 - Len(txt开始No), "0") & Me.txt开始No
    End If
    If Len(txt结束NO) < 8 And Len(txt结束NO) > 0 Then
        Me.txt结束NO = UCase(LTrim(Me.txt结束NO))
        intYear = Format(Sys.Currentdate, "YYYY") - 1990
        strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
        If Len(txt结束NO) < 8 Then Me.txt结束NO = strYear & String(7 - Len(txt结束NO), "0") & Me.txt结束NO
    End If
    
    mstrOthers(1) = Trim(Me.txt开始No.Text)
    mstrOthers(2) = Trim(Me.txt结束NO.Text)

    If Me.txt开始No <> "" And Me.txt结束NO <> "" Then mstrFind = mstrFind & " And A.No >= [7] And A.No <=[8]"
    If Me.txt开始No <> "" And Me.txt结束NO = "" Then mstrFind = mstrFind & " And A.No >= [7]"
    If Me.txt开始No = "" And Me.txt结束NO <> "" Then mstrFind = mstrFind & " And A.No <= [8]"
    
    '扩展查询条件
    If BlnAdvance = False Then
        Unload Me
        Exit Sub
    End If
    
    ' 0-记录状态(计划类型),1-开始单号,2-结束单号,3-材料id,4-对方部门id(或入出类别id),5-填制人,
    ' 6-审核人,7-供应商ID,8-生产商,9-开始生产日期,10-结束生产日期,11-开始发票号,12-结束发票号
     '参数范围:[1]-库房id,[2]:开始填制日期,[3]结束填制日期,[4]开始审核日期,[5] 结束审核日期,[6]-记录状态,[7]开始单据号,
     '[8]结束单据号,[9]材料id,[10]对方部门id,[11]填制人,[12]审核人[13]-供应商ID,[14]-生产商,
     '[15]-开始生产日期,[16]-结束生产日期,[17]-开始发票号,[18]-结束发票号
  
  
    If Chk材料.Value = 1 Then
        lng材料ID = Txt材料.Tag
        mstrFind = mstrFind & " And A.药品ID=[9]"
        mstrOthers(3) = Txt材料.Tag
    End If
    If Chk类别.Value = 1 Then
        mstrFind = mstrFind & " And A.入出类别ID=[10]"
        mstrOthers(4) = Cbo类别.ItemData(Cbo类别.ListIndex)
    End If
    
    If Me.Txt填制人 <> "" Then
        mstrFind = mstrFind & " And A.填制人 like '" & Me.Txt填制人 & "%'"
        mstrOthers(5) = Trim(Me.Txt填制人) & "%"
    End If
    If Me.Txt审核人 <> "" Then
        mstrFind = mstrFind & " And A.审核人 like [12]"
        mstrOthers(6) = Trim(Me.Txt审核人) & "%"
    End If
    
    If Chk生产商.Value = 1 Then
        mstrFind = mstrFind & " And A.产地=[14]"
        mstrOthers(8) = txt生产商.Text
    End If
    
    If chk生产日期.Value = 1 Then
        mstrFind = " And A.生产日期 Between [15] And [16] "
        mstrOthers(9) = Format(dtp开始时间(2), "yyyy-mm-dd")
        mstrOthers(10) = Format(dtp结束时间(2), "yyyy-mm-dd")
    End If
    
    If gblnCode = True And Trim(txt条码.Text) <> "" Then
        mstrOthers(13) = UCase(Trim(txt条码.Text))
        mstrFind = mstrFind & " And (A.商品条码 Like [19] Or A.内部条码 Like [19])"
    End If
    
    Unload Me
End Sub

Private Sub Cmd生产商_Click()
    Dim rsTemp As New Recordset
    Dim vRect As RECT
    Dim blnCancel As Boolean
    
    vRect = zlControl.GetControlRect(txt生产商.hwnd)
    
    gstrSQL = "Select rownum as id,null as 上级id,编码,名称,简码,1 as 末级 From 材料生产商 " & _
              "Where (站点 = [1] or 站点 is null) "
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
    If rsTemp Is Nothing Then Exit Sub
    If rsTemp.State <> 1 Then Exit Sub
    With rsTemp
        txt生产商.Tag = 1
        txt生产商.Text = zlStr.Nvl(!名称)
    End With
End Sub

Private Sub Cmd材料_Click()
    Dim RecReturn As Recordset
    
    Set RecReturn = Frm材料选择器.ShowMe(Me, 1, 0, mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex))
    If RecReturn.RecordCount = 0 Then Exit Sub
    Txt材料 = "[" & RecReturn!编码 & "]" & RecReturn!名称
    Txt材料.Tag = RecReturn!材料ID
    
    Chk生产商.SetFocus
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
    Me.dtp结束时间(0) = Sys.Currentdate
    Me.dtp结束时间(1) = Me.dtp结束时间(0)
    
    Me.dtp开始时间(0) = DateAdd("d", -7, Me.dtp结束时间(0))
    Me.dtp开始时间(1) = Me.dtp开始时间(0)
    
    Me.dtp开始时间(2) = DateAdd("d", -7, Me.dtp结束时间(0))
    Me.dtp结束时间(2) = Me.dtp结束时间(0)
    
    lbl条码.Visible = gblnCode
    txt条码.Visible = gblnCode
    
    Me.Txt材料.Tag = 0
    Me.txt生产商.Tag = 0
    lng材料ID = 0
    
    '打开记录集
    sstFilter.Tab = 0
    BlnAdvance = False
    
End Sub

Private Function CheckCompete() As Boolean
    Dim rsTemp As New Recordset
    
    On Error GoTo ErrHandle
    CheckCompete = False
    
    gstrSQL = "Select 编码,名称,简码 From 材料生产商 where (站点 = [1] or 站点 is null) "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "-材料生产商", gstrNodeNo)
    With rsTemp
        If .EOF Then
            MsgBox "卫材生产商信息不全,请在字典管理中设置卫材生产商信息！", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    With rsTemp
        gstrSQL = "SELECT B.Id, b.名称 " & _
                  "FROM 药品单据性质 A, 药品入出类别 B " & _
                  "Where A.类别id = B.ID AND A.单据 = 32 "
        zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
        
        If .EOF Then
            MsgBox "卫材其他入库没有设置相应的入出类别，请检查卫材入出分类！", vbInformation, gstrSysName
            Exit Function
        End If
        .MoveFirst
        Do While Not .EOF
            Cbo类别.AddItem .Fields(1)
            Cbo类别.ItemData(Cbo类别.NewIndex) = .Fields(0)
            .MoveNext
        Loop
        Cbo类别.ListIndex = 0
        .Close
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
                Case "Maker"
                    txt生产商.Text = .TextMatrix(.Row, 1)
                    txt生产商.Tag = 1
                    Chk类别.SetFocus
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

Private Sub txt结束NO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng库房ID As Long

    lng库房ID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    If KeyCode = vbKeyReturn Then
        If Len(txt结束NO) < 8 And Len(txt结束NO) > 0 Then
            txt结束NO.Text = zlCommFun.GetFullNO(txt结束NO.Text, 70, lng库房ID)
        End If
        OS.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txt结束NO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt开始No_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng库房ID As Long
    
    lng库房ID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    If KeyCode = vbKeyReturn Then
        If Len(txt开始No) < 8 And Len(txt开始No) > 0 Then
            txt开始No.Text = zlCommFun.GetFullNO(txt开始No.Text, 70, lng库房ID)
        End If
        OS.PressKey (vbKeyTab)
    End If

End Sub

Private Sub txt开始No_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt审核人_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then cmd确定.SetFocus
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt审核人.Text) = "" Then
            cmd确定.SetFocus
            Exit Sub
        End If
        Txt审核人.Text = UCase(Txt审核人.Text)
        
        gstrSQL = "" & _
            "   Select 编号,简码,姓名 " & _
            "   From 人员表 " & _
            "   Where (姓名 like [1] or 编号 like [1] or 简码 like [1] ) And (站点 = [2] or 站点 is null) " & _
            "       and (撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) " & _
            "   order by 编号"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取审核人", IIf(gstrMatchMethod = "0", "%", "") & Me.Txt审核人 & "%", gstrNodeNo)
            
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
                cmd确定.SetFocus
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
    
    On Error GoTo ErrHandle
    If KeyCode = vbKeyReturn Then
        If Me.txt生产商 = "" Then Exit Sub
        If Trim(txt生产商) = "" Then Exit Sub
        txt生产商 = UCase(txt生产商)
    
        Dim rsTemp As New ADODB.Recordset
        
        gstrSQL = "Select 编码,名称,简码 From 材料生产商 Where upper(名称) like [1] or Upper(编码) like [1] or Upper(简码) like [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "材料生产商", IIf(gstrMatchMethod = "0", "%", "") & Me.txt生产商 & "%")
        
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
            End If
        End With
        
        If Chk类别.Visible = True Then
            If Chk类别.Value = 1 Then
                Cbo类别.SetFocus
            Else
                Chk类别.SetFocus
            End If
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub Txt填制人_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then Me.Txt审核人.SetFocus
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
            "   Where (姓名 like [1] or 编号 like [1] or 简码 like [1] ) And (站点 = [2] or 站点 is null) " & _
            "       and (撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取填制人", IIf(gstrMatchMethod = "0", "%", "") & Me.Txt填制人 & "%", gstrNodeNo)
        
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
                    .Top = sstFilter.Top + fra附加条件.Top + Txt填制人.Top - .Height ' + Txt填制人.Height
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
    
    Chk生产商.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

