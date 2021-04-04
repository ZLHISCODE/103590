VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmTransferSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤"
   ClientHeight    =   4260
   ClientLeft      =   3150
   ClientTop       =   3165
   ClientWidth     =   7560
   Icon            =   "frmTransferSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   2190
      Left            =   6075
      TabIndex        =   31
      Top             =   2250
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3863
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
      Left            =   105
      TabIndex        =   25
      Top             =   135
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7011
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "范围(&R)"
      TabPicture(0)   =   "frmTransferSearch.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra范围"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "附加条件(&D)"
      TabPicture(1)   =   "frmTransferSearch.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra附加条件"
      Tab(1).ControlCount=   1
      Begin VB.Frame fra附加条件 
         Height          =   2850
         Left            =   -74760
         TabIndex        =   30
         Top             =   600
         Width           =   5520
         Begin VB.TextBox txt条码 
            Height          =   300
            Left            =   1530
            TabIndex        =   33
            Top             =   2400
            Width           =   3765
         End
         Begin VB.CommandButton cmdDept 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4900
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   960
            Width           =   270
         End
         Begin VB.TextBox txtDept 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1530
            TabIndex        =   13
            Top             =   960
            Width           =   3375
         End
         Begin VB.CheckBox Chk移入库房 
            Caption         =   "移入库房"
            Height          =   420
            Left            =   420
            TabIndex        =   12
            Top             =   945
            Width           =   1110
         End
         Begin VB.CommandButton Cmd材料 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4900
            TabIndex        =   11
            Top             =   420
            Width           =   255
         End
         Begin VB.TextBox Txt材料 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1530
            MaxLength       =   50
            ScrollBars      =   3  'Both
            TabIndex        =   10
            Top             =   420
            Width           =   3375
         End
         Begin VB.CheckBox Chk材料 
            Caption         =   "卫生材料"
            Height          =   300
            Left            =   420
            TabIndex        =   9
            Top             =   420
            Width           =   1035
         End
         Begin VB.TextBox Txt填制人 
            Height          =   300
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   16
            Top             =   1500
            Width           =   1845
         End
         Begin VB.TextBox Txt审核人 
            Height          =   300
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   17
            Top             =   1980
            Width           =   1845
         End
         Begin VB.ComboBox Cbo移入库房 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1530
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   980
            Width           =   3615
         End
         Begin VB.Label lbl条码 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "条  码"
            Height          =   180
            Left            =   570
            TabIndex        =   34
            Top             =   2460
            Width           =   540
         End
         Begin VB.Label LblEnterStock 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "领料部门(&D)"
            Height          =   180
            Left            =   480
            TabIndex        =   32
            Top             =   1005
            Width           =   990
         End
         Begin VB.Label Lbl填制人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "填制人"
            Height          =   180
            Left            =   570
            TabIndex        =   23
            Top             =   1560
            Width           =   540
         End
         Begin VB.Label Lbl审核人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "审核人"
            Height          =   180
            Left            =   570
            TabIndex        =   24
            Top             =   2040
            Width           =   540
         End
      End
      Begin VB.Frame fra范围 
         Height          =   2850
         Left            =   240
         TabIndex        =   26
         Top             =   600
         Width           =   5520
         Begin VB.CheckBox chkNoStrike 
            Caption         =   "未审核冲销"
            Height          =   300
            Left            =   2040
            TabIndex        =   36
            Top             =   2280
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CheckBox chkYesStrike 
            Caption         =   "已审核冲销"
            Enabled         =   0   'False
            Height          =   300
            Left            =   3585
            TabIndex        =   35
            Top             =   2280
            Visible         =   0   'False
            Width           =   1335
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
            Format          =   169738243
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
            Format          =   169738243
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
            Format          =   169738243
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
            Format          =   169738243
            CurrentDate     =   36263
         End
         Begin VB.Label LblNO 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "No"
            Height          =   180
            Left            =   480
            TabIndex        =   20
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
            TabIndex        =   29
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
            TabIndex        =   22
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
            TabIndex        =   28
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
            TabIndex        =   21
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
            TabIndex        =   27
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
      TabIndex        =   19
      Top             =   930
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6330
      TabIndex        =   18
      Top             =   420
      Width           =   1100
   End
End
Attribute VB_Name = "FrmTransferSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrFind As String  '查找字符串
Private BlnAdvance As Boolean '是否展开
Private mlngMode As Long    '单据类型
Private mdatStart As Date   '开始时间
Private mdatEnd As Date     '结束时间
Private mdatVerifyStart As Date
Private mdatVerifyEnd As Date
Private mfrmMain As Form    '父窗体
Private mstrSelectTag As String     '当前选择的对象
Private mstrOthers(0 To 13) As String ' 0-记录状态(计划类型),1-开始单号,2-结束单号,3-材料id,4-对方部门id(或入出类别id或编制方法(计划用)),5-填制人,6-审核人,7-供应商ID,8-生产商,9-开始生产日期,10-结束生产日期,11-开始发票号,12-结束发票号,13-条码信息
Private mint冲销申请 As Integer '0-不需要申请 1-需要申请
Private mstrPrivs As String

Public Function GetSearch(ByVal frmMain As Form, ByVal lngMode As Long, _
        ByRef datStart As Date, ByRef datEnd As Date, _
        ByRef datVerifyStart As Date, ByRef datVerifyEnd As Date, _
        ByRef strPrivs As String, _
        ByRef strOthers() As String) As String
        'strOthers():返回相关参数值:0-记录状态(计划类型),1-开始单号,2-结束单号,3-材料id,4-对方部门id(或入出类别id或编制方法(计划用)),5-填制人,6-审核人)
        
    mstrFind = ""
    mlngMode = lngMode
    mstrPrivs = strPrivs
    Set mfrmMain = frmMain
    Me.Show vbModal, mfrmMain
    
    GetSearch = mstrFind
    datStart = mdatStart
    datEnd = mdatEnd
    datVerifyStart = mdatVerifyStart
    datVerifyEnd = mdatVerifyEnd
    strOthers = mstrOthers
    
End Function

Private Sub Cbo移入库房_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub chkStrike_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmd确定.SetFocus
    End If
    
End Sub

Private Sub chk审核_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If chk审核.Value = 1 Then
            SendKeys vbTab
        Else
            cmd确定.SetFocus
        End If
    End If
    
End Sub

Private Sub chk填制_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
    
End Sub

Private Sub Chk材料_GotFocus()
    If sstFilter.Tab = 0 Then
        sstFilter.Tab = 1
        Chk材料.SetFocus
    End If
End Sub

Private Sub Chk材料_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub Chk材料_Click()
    Txt材料.Enabled = IIf(Chk材料.Value = 1, True, False)
    Cmd材料.Enabled = IIf(Chk材料.Value = 1, True, False)
End Sub

Private Sub Chk移入库房_click()
    If mlngMode = 1718 Then
        Cbo移入库房.Enabled = IIf(Chk移入库房.Value = 1, True, False)
    Else
        txtDept.Enabled = IIf(Chk移入库房.Value = 1, True, False)
        cmdDept.Enabled = IIf(Chk移入库房.Value = 1, True, False)
    End If
End Sub

Private Sub Chk移入库房_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    OS.PressKey vbKeyTab
'    If Chk移入库房.Value = 1 Then
'        Cbo移入库房.SetFocus
'    Else
'        Txt填制人.SetFocus
'    End If
End Sub
Private Sub chk填制_Click()
    dtp开始时间(0).Enabled = IIf(chk填制.Value = 1, True, False)
    dtp结束时间(0).Enabled = IIf(chk填制.Value = 1, True, False)
    chkNoStrike.Enabled = IIf(chk填制.Value = 1, True, False)
End Sub

Private Sub chk审核_Click()
    dtp开始时间(1).Enabled = IIf(chk审核.Value = 1, True, False)
    dtp结束时间(1).Enabled = IIf(chk审核.Value = 1, True, False)
    chkStrike.Enabled = IIf(chk审核.Value = 1, True, False)
    chkYesStrike.Enabled = IIf(chk审核.Value = 1, True, False)
End Sub

Private Sub cmdDept_Click()
    If getDept("") = False Then
        Exit Sub
    End If
    If Txt填制人.Enabled Then Txt填制人.SetFocus
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
            MsgBox "请选择需查询的材料信息！", vbInformation, gstrSysName
            Me.Txt材料.SetFocus
            Exit Sub
        End If
    End If
    
    If chk填制.Value = 0 And chk审核.Value = 0 Then
        MsgBox "必须选择一个填制日期或者审核日期!", vbInformation, gstrSysName
        chk填制.SetFocus
        Exit Sub
    End If
    mstrFind = ""
    '基本查询条件
    
    mdatStart = Format("1901-01-01", "yyyy-mm-dd")
    mdatEnd = Format("1901-01-01", "yyyy-mm-dd")
    mdatVerifyStart = Format("1901-01-01", "yyyy-mm-dd")
    mdatVerifyEnd = Format("1901-01-01", "yyyy-mm-dd")
    Dim i As Integer
    For i = 0 To 13
        mstrOthers(i) = ""
    Next
    
    mstrOthers(0) = IIf(chkStrike.Value = 1, "0", "1")
    
    '2-开始填制日期,3-结束填制日期
    
    If chk填制.Value = 1 And chk审核.Value = 1 Then
        If mlngMode <> 1716 Then '材料移库
            If chkStrike.Value = 1 Then
                mstrFind = " And ((A.填制日期 Between [2] And [3] and A.审核人 is null) " _
                        & " or (A.审核日期 Between [4] And [5]))"
            Else
                mstrFind = " And ((A.填制日期 Between [2] And [3] and A.审核人 is null) " _
                        & " or (A.审核日期 Between [4] And [5] and a.记录状态 =[6]))  "
            End If
        Else
            If chkStrike.Value = 1 Then
                mstrFind = " And ((A.填制日期 Between [2] And [3] and A.审核人 is null) " _
                    & " or (A.审核日期 Between [4] And [5]))"
            Else
                If chkNoStrike.Value = 1 And chkYesStrike.Value = 1 Then
                    mstrFind = " And ((A.填制日期 Between [2] And [3] and A.审核人 is null) " _
                                & " or (A.审核日期 Between [4] And [5]))"
                ElseIf chkNoStrike.Value = 1 And chkYesStrike.Value = 0 Then
                    mstrFind = " and (A.记录状态=2 or mod(A.记录状态,3)=2) And A.填制日期 Between [2] And [3] and A.审核人 is null  "
                ElseIf chkNoStrike.Value = 0 And chkYesStrike.Value = 1 Then
                    mstrFind = " and (A.记录状态=2 or mod(A.记录状态,3)=2) And ((A.填制日期 Between [2] And [3] and A.审核人 is null) " _
                                & " or (A.审核日期 Between [4] And [5])) and A.审核人 is not null "
                Else
                    mstrFind = " And ((A.填制日期 Between [2] And [3] and A.审核人 is null) " _
                                & " or (A.审核日期 Between [4] And [5])) and a.记录状态 =1 "
                End If
            End If
        End If
        mdatStart = Format(dtp开始时间(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp结束时间(0), "yyyy-mm-dd")
        mdatVerifyStart = Format(dtp开始时间(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp结束时间(1), "yyyy-mm-dd")
    ElseIf chk审核.Value = 1 Then
        If mlngMode <> 1716 Then
            If chkStrike.Value = 1 Then
                mstrFind = " And A.审核日期 Between [4] And [5] "
            Else
                mstrFind = " And A.审核日期 Between [4] And [5] and a.记录状态 =[6] "
            End If
        Else
            If chkStrike.Value = 1 Then
                mstrFind = " And A.审核日期 Between [4] And [5] "
            Else
                If chkYesStrike.Value = 1 Then
                    mstrFind = " and (A.记录状态=2 or mod(A.记录状态,3)=2) And A.审核日期 Between [4] And [5] "
                Else
                    mstrFind = " And A.审核日期 Between [4] And [5] and a.记录状态 =1"
                End If
            End If
        End If
        mdatVerifyStart = Format(dtp开始时间(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp结束时间(1), "yyyy-mm-dd")
    ElseIf chk填制.Value = 1 Then
        If mlngMode <> 1716 Then
            mstrFind = " And (A.填制日期 Between [2] And [3] and 审核日期 is null ) "
        Else
            If chkNoStrike.Value = 1 Then
                mstrFind = " and (A.记录状态=2 or mod(A.记录状态,3)=2) and (A.填制日期 Between [2] And [3]) and 审核日期 is null "
            Else
                mstrFind = " And (A.填制日期 Between [2] And [3]) and 审核日期 is null "
            End If
        End If
        mdatStart = Format(dtp开始时间(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp结束时间(0), "yyyy-mm-dd")
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
    If Me.txt开始No <> "" And Me.txt结束NO = "" Then mstrFind = mstrFind & " And A.No >=[7] "
    If Me.txt开始No = "" And Me.txt结束NO <> "" Then mstrFind = mstrFind & " And A.No <=[8] "
    
    '扩展查询条件
    
    If BlnAdvance = False Then
        Unload Me
        Exit Sub
    End If
    
    
    If Chk材料.Value = 1 Then
        mstrFind = mstrFind & " And A.药品ID=[9]"
        mstrOthers(3) = Val(Txt材料.Tag)
    End If
    
    If Chk移入库房.Value = 1 Then
        If txtDept.Visible = True Then
            mstrOthers(4) = txtDept.Tag
        Else
            mstrOthers(4) = Cbo移入库房.ItemData(Cbo移入库房.ListIndex)
        End If
    End If
    If mlngMode = 1718 Then
        If Chk移入库房.Value = 1 Then
            mstrFind = mstrFind & " And A.入出类别ID=[10]"
        End If
    Else
        If Chk移入库房.Value = 1 Then
            mstrFind = mstrFind & " And A.对方部门ID=[10]"
        End If
        
    End If
    If Me.Txt填制人 <> "" Then
        mstrFind = mstrFind & " And A.填制人 like [11] "
        mstrOthers(5) = Txt填制人.Text & "%"
    End If
    
    If Me.Txt审核人 <> "" Then
        mstrFind = mstrFind & " And A.审核人 like [12]"
        mstrOthers(6) = Me.Txt审核人 & "%"
    End If
    
    If gblnCode = True And Trim(txt条码.Text) <> "" Then
        mstrOthers(13) = UCase(Trim(txt条码.Text))
        mstrFind = mstrFind & " And (A.商品条码 Like [19] Or A.内部条码 Like [19])"
    End If
    
    Unload Me
End Sub

Private Sub Cmd材料_Click()
    Dim RecReturn As Recordset
    
    Set RecReturn = Frm材料选择器.ShowMe(Me, 1, 0, _
        mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), _
        True, True, False, False, True, 0, True, False, "", False, 0, False, mstrPrivs)
    If RecReturn.RecordCount = 0 Then Exit Sub
    Txt材料 = "[" & RecReturn!编码 & "]" & RecReturn!名称
    Txt材料.Tag = RecReturn!材料ID
    
    If Chk移入库房.Visible = True Then
        Chk移入库房.SetFocus
    Else
        Txt填制人.SetFocus
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
    
    Dim intLop As Integer
    
    Me.dtp结束时间(0) = sys.Currentdate
    Me.dtp结束时间(1) = Me.dtp结束时间(0)
    Me.dtp开始时间(0) = DateAdd("d", -7, Me.dtp结束时间(0))
    Me.dtp开始时间(1) = Me.dtp开始时间(0)
    
    Me.Txt材料.Tag = 0
    sstFilter.Tab = 0
    
    lbl条码.Visible = gblnCode
    txt条码.Visible = gblnCode
    
    Select Case mlngMode
        Case 1715   '差价调整
            Chk移入库房.Caption = "库房"
            lbl条码.Visible = False
            txt条码.Visible = False
        Case 1716 '移库管理
            mint冲销申请 = IIf(Val(zlDatabase.GetPara("冲销申请", glngSys, mlngMode, "0")) = 1, 1, 0)
            Chk移入库房.Caption = "发出库房"
            If mint冲销申请 = 1 Then
                chkStrike.Visible = False
                chkNoStrike.Visible = True
                chkYesStrike.Visible = True
            Else
                chkStrike.Visible = True
                chkNoStrike.Visible = False
                chkYesStrike.Visible = False
            End If
        Case 1717   '卫材领用
            Chk移入库房.Caption = "领用部门"
        Case 1718   '其他出库管理
            Chk移入库房.Caption = "入出类别"
        Case 1719   '盘点管理
            lbl条码.Visible = False
            txt条码.Visible = False
    End Select
    
    '打开记录集
    
    BlnAdvance = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mshSelect.Visible = True Then
        mshSelect.Visible = False
        Select Case mstrSelectTag
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

Private Sub sstFilter_Click(PreviousTab As Integer)
    Dim rsTemp As New Recordset
    Dim strStock As String
    
    On Error GoTo ErrHandle
    With sstFilter
        If .Tab = 1 Then
            BlnAdvance = True
            txtDept.Visible = False
            cmdDept.Visible = False
            
            Cbo移入库房.Visible = True
            If Cbo移入库房.ListCount < 1 Then
                Select Case mlngMode
                    Case 1716
                        txtDept.Visible = True
                        cmdDept.Visible = True
                        Cbo移入库房.Visible = False
                        Exit Sub
'                        strStock = "V,W,K,12"
'                        gstrSQL = "" & _
'                        "   SELECT /*+ Rule*/ DISTINCT a.id, a.名称 " & _
'                        "   FROM 部门性质说明 c, 部门性质分类 b, 部门表 a, Table(cast(f_Str2List([2]) as zlTools.t_StrList)) D " & _
'                        "   Where c.工作性质 = b.名称 " & _
'                        "       AND b.编码=D.Column_Value " & _
'                        "       AND a.id = c.部门id " & _
'                        "       AND a.撤档时间 = to_date('3000-01-01','yyyy-MM-dd')"
                    Case 1717
                        txtDept.Visible = True
                        cmdDept.Visible = True
                        Cbo移入库房.Visible = False
                        Exit Sub
                        'strStock = "O"
'                        If Check普通科室 = True Then
'                            gstrSQL = "" & _
'                            "   SELECT DISTINCT a.id, a.名称 " & _
'                            "   FROM  部门表 a " & _
'                            "   Where a.撤档时间 = to_date('3000-01-01','yyyy-MM-dd') " & _
'                            "       and a.ID in (Select 部门ID From 部门人员 Where 人员ID=[1])"
'                        Else
'                            gstrSQL = "" & _
'                            "   SELECT DISTINCT a.id, a.名称 " & _
'                            "   FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " & _
'                            "   Where c.工作性质 = b.名称 " & _
'                            "       AND a.id = c.部门id " & _
'                            "       AND a.撤档时间 = to_date('3000-01-01','yyyy-MM-dd')"
'                            '"               AND b.编码 in " & strStock
'                        End If
                    Case 1718
                        gstrSQL = "" & _
                            "   SELECT b.Id,b.名称 " & _
                            "   FROM 药品单据性质 A, 药品入出类别 B " & _
                            "   Where A.类别id = B.ID AND A.单据 = 36 "
                    Case 1715, 1719
                        If Chk移入库房.Visible = True Then
                            Chk移入库房.Visible = False
                            Cbo移入库房.Visible = False
                        End If
                        LblEnterStock.Visible = False
                        Exit Sub
                    Case Else
                End Select
                
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UserInfo.Id, strStock)
                With Cbo移入库房
                    Do While Not rsTemp.EOF
                        .AddItem rsTemp.Fields(1)
                        .ItemData(.NewIndex) = rsTemp.Fields(0)
                        rsTemp.MoveNext
                    Loop
                    If .ListCount > 0 Then .ListIndex = 0
                End With
                rsTemp.Close
            End If
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtDept_Change()
    txtDept.Tag = ""
End Sub

Private Sub txtDept_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If txtDept.Tag = "" Then
            If getDept(Trim(txtDept.Text)) = False Then
                Exit Sub
            End If
        End If
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub txt结束NO_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim IntBill As Integer
    Dim lng库房ID As Long
    Dim strNo As String
    Dim rsTemp As New ADODB.Recordset
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    On Error GoTo ErrHandle
    Select Case mlngMode
        Case 1715         '卫材库存差价调整'
            IntBill = 71        '卫材库存差价调整
        Case 1716, 1722         '卫材申领管理'
            If UCase(mfrmMain.Name) = UCase("frmRequestStuffList") Then
                '表示申领单
                '因为不好确定库房,不能分出单据
                gstrSQL = "Select 项目序号 From 号码控制表 where 项目序号=72 and 编号规则<>2"
                zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
                If rsTemp.RecordCount <> 0 Then
                    GoTo NO:
                End If
                rsTemp.Close
                OS.PressKey (vbKeyTab)
                Exit Sub
            End If
            IntBill = 72        '卫材库房转移
        Case 1717         '卫材领用管理'
            IntBill = 73        '部门领用卫材
        Case 1718         '卫材其他出库管理'
            IntBill = 74        '卫材其他出库
        Case 1719         '卫材盘点管理'
            If mfrmMain.TabShow.Tab = 0 Then
                IntBill = 76        '库存卫材盘点
            Else
                IntBill = 75        '库存卫材盘点
            End If
        Case Else
            IntBill = 0
    End Select
    
    If IntBill = 0 Then
NO:
        If Len(txt结束NO) < 8 And Len(txt结束NO) > 0 Then
            Dim intYear  As Integer, strYear As String
            Me.txt结束NO = UCase(LTrim(Me.txt结束NO))
            intYear = Format(sys.Currentdate, "YYYY") - 1990
            strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
            If Len(txt结束NO) < 8 Then Me.txt结束NO = strYear & String(7 - Len(txt结束NO), "0") & Me.txt结束NO
        End If
        OS.PressKey (vbKeyTab)
    Else
        lng库房ID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
        If KeyCode = vbKeyReturn Then
            If Len(txt结束NO) < 8 And Len(txt结束NO) > 0 Then
                txt结束NO.Text = zlCommFun.GetFullNO(txt结束NO.Text, IntBill, lng库房ID)
            End If
            OS.PressKey (vbKeyTab)
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt结束NO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub


Private Sub txt开始No_KeyDown(KeyCode As Integer, Shift As Integer)
     
    Dim IntBill As Integer
    Dim lng库房ID As Long
    Dim strNo As String
    Dim rsTemp As New ADODB.Recordset
    If KeyCode <> vbKeyReturn Then Exit Sub
    On Error GoTo ErrHandle
    
    Select Case mlngMode
        Case 1715         '卫材库存差价调整'
            IntBill = 71        '卫材库存差价调整
        Case 1716               '移库'
            If UCase(mfrmMain.Name) = UCase("frmRequestStuffList") Then
                '表示申领单
                '因为不好确定库房,不能分出单据
                gstrSQL = "Select 项目序号 From 号码控制表 where 项目序号=72 and 编号规则<>2"
                zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
                If rsTemp.RecordCount <> 0 Then
                    GoTo NO:
                End If
                rsTemp.Close
                OS.PressKey (vbKeyTab)
                Exit Sub
            End If
            IntBill = 72        '卫材库房转移
        Case 1722               '卫材申领管理
        Case 1717         '卫材领用管理'
            IntBill = 73        '部门领用卫材
        Case 1718         '卫材其他出库管理'
            IntBill = 74        '卫材其他出库
        Case 1719         '卫材盘点管理'
            If mfrmMain.TabShow.Tab = 0 Then
                IntBill = 76        '库存卫材盘点
            Else
                IntBill = 75        '库存卫材盘点
            End If
        Case Else
            IntBill = 0
    End Select
    
    If IntBill = 0 Then
NO:
        If Len(txt开始No) < 8 And Len(txt开始No) > 0 Then
            Dim intYear  As Integer, strYear As String
            
            Me.txt开始No = UCase(LTrim(Me.txt开始No))
            intYear = Format(sys.Currentdate, "YYYY") - 1990
            strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
            If Len(txt开始No) < 8 Then Me.txt开始No = strYear & String(7 - Len(txt开始No), "0") & Me.txt开始No
        End If
        OS.PressKey (vbKeyTab)
    Else
        lng库房ID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
        If KeyCode = vbKeyReturn Then
            If Len(txt开始No) < 8 And Len(txt开始No) > 0 Then
                txt开始No.Text = zlCommFun.GetFullNO(txt开始No.Text, IntBill, lng库房ID)
            End If
            OS.PressKey (vbKeyTab)
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt开始No_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub


Private Sub Txt审核人_KeyDown(KeyCode As Integer, Shift As Integer)
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
            "   Where (姓名 like [1] or 编号 like [1] or 简码 like [1] ) And (站点=[2] or 站点 is null) " & _
            "       and (撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)" & _
            "   order by 编号"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, GetMatchingSting(Me.Txt审核人), gstrNodeNo)
        
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
                    If .Top + .Height > Me.ScaleHeight Then .Top = .Top - .Height - Txt审核人.Height
                    .Left = sstFilter.Left + fra附加条件.Left + Txt审核人.Left
                    .Visible = True
                    .SetFocus
                    .ColWidth(0) = 800
                    .ColWidth(1) = 800
                    .ColWidth(2) = .Width - .ColWidth(1) - .ColWidth(0)
                    .Row = 1
                    .Col = 0
                    .ColSel = .Cols - 1
                    .ZOrder 0
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
            "   Where (姓名 like [1] or 编号 like [1] or 简码 like [1] ) And (站点=[2] or 站点 is null) " & _
            "       and (撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)" & _
            "   order by 编号"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, GetMatchingSting(Me.Txt填制人), gstrNodeNo)
        
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
                    If .Top + .Height > Me.ScaleHeight Then .Top = .Top - .Height - Txt填制人.Height
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

Private Sub Txt填制人_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt材料_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim RecReturn As Recordset
    Dim strKey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(Txt材料.Text) = "" Then Exit Sub
    sngLeft = Me.Left + fra附加条件.Left + Txt材料.Left
    sngTop = Me.Top + fra附加条件.Top + Txt材料.Top + Txt材料.Height + Me.Height - Me.ScaleHeight '  50
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
    
    Set RecReturn = FrmMulitSel.ShowSelect(Me, 1, , mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex), _
        strKey, sngLeft, sngTop, Txt材料.Width, Txt材料.Height, _
        True, True, False, False, True, _
        0, True, "", False, 0, _
        False, mstrPrivs)
    If RecReturn.RecordCount = 0 Then Exit Sub
    Txt材料 = "[" & RecReturn!编码 & "]" & RecReturn!名称
    Txt材料.Tag = RecReturn!材料ID
    
    If Chk移入库房.Visible = True Then
        Chk移入库房.SetFocus
    Else
        Txt填制人.SetFocus
    End If
    
End Sub

Private Sub Txt材料_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub mshSelect_DblClick()
    mshSelect_KeyPress 13
End Sub

Private Sub mshSelect_KeyPress(KeyAscii As Integer)
    With mshSelect
        If KeyAscii = 13 Then
            Select Case mstrSelectTag
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

Private Function getDept(ByVal strKey As String) As Boolean
    Dim rsTemp As New Recordset
    Dim strSeach As String
    Dim vRect As RECT
    Dim lngH As Long
    Dim strStock As String
    Dim blnCancel As Boolean
    Dim strWhere As String
    
    strSeach = strKey
    
    strWhere = ""
    If strSeach <> "" Then
        strSeach = GetMatchingSting(strSeach)
        strWhere = "           and (a.编码 like [1] or a.名称 like [1] or a.简码 like [1]) "
    End If
    
    Select Case mlngMode
    Case 1716
        strStock = "V,W,K,12"
        gstrSQL = "" & _
        "   SELECT /*+ Rule*/ DISTINCT a.id,a.编码, a.名称,a.简码,a.位置" & _
        "   FROM 部门性质说明 c, 部门性质分类 b, 部门表 a, Table(cast(f_Str2List([3]) as zlTools.t_StrList)) D " & _
        "   Where c.工作性质 = b.名称 And (a.站点=[2] or a.站点 is null) " & _
        "       AND b.编码=D.Column_value " & _
        "       AND a.id = c.部门id " & _
        "       AND (TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01' or a.撤档时间 is null)" & _
            strWhere
    Case 1717
        'strStock = "O"
        If Check普通科室 = True Then
            gstrSQL = "" & _
                "  SELECT DISTINCT a.id,a.编码,a.名称,a.简码,a.位置 " & _
            "      FROM 部门表 a " & _
            "      Where ( TO_CHAR(a.撤档时间, 'yyyy-MM-dd') = '3000-01-01' or a.撤档时间 is null ) And (a.站点=[2] or a.站点 is null) " & _
            "           and a.ID in (Select 部门ID From 部门人员 Where 人员ID=[4] ) " & _
                strWhere
        Else
            gstrSQL = "" & _
            "   SELECT DISTINCT a.id,a.编码, a.名称,a.简码,a.位置 " & _
            "   FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " & _
            "   Where c.工作性质 = b.名称 And (a.站点=[2] or a.站点 is null) " & _
            "       AND a.id = c.部门id " & _
            "       AND (TO_CHAR (a.撤档时间, 'yyyy-MM-dd') = '3000-01-01' or a.撤档时间 is null)" & _
                strWhere
        End If
    End Select
    
    vRect = zlControl.GetControlRect(txtDept.hwnd)
    lngH = txtDept.Height
    
    'If strkey = "" Then Exit Function
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "部门选择", False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, strSeach, gstrNodeNo, strStock, UserInfo.Id)

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
        If txtDept.Enabled Then txtDept.SetFocus
        Exit Function
    End If
    If rsTemp.State <> 1 Then
        If txtDept.Enabled Then txtDept.SetFocus
        Exit Function
    End If
    
    Me.txtDept = zlStr.Nvl(rsTemp!编码) & "-" & zlStr.Nvl(rsTemp!名称)
    Me.txtDept.Tag = zlStr.Nvl(rsTemp!Id)
    getDept = True
    
End Function
