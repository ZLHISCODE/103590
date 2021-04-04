VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmStuffPlanSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤"
   ClientHeight    =   4260
   ClientLeft      =   3150
   ClientTop       =   3165
   ClientWidth     =   7410
   Icon            =   "frmStuffPlanSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   1815
      Left            =   2160
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   3240
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3201
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
      TabIndex        =   21
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7011
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "范围(&R)"
      TabPicture(0)   =   "frmStuffPlanSearch.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra范围"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "附加条件(&D)"
      TabPicture(1)   =   "frmStuffPlanSearch.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra附加条件"
      Tab(1).ControlCount=   1
      Begin VB.Frame fra附加条件 
         Height          =   2850
         Left            =   -74760
         TabIndex        =   26
         Top             =   600
         Width           =   5520
         Begin VB.ComboBox cbo编制方法 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1770
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   1020
            Width           =   3615
         End
         Begin VB.CheckBox chk编制方法 
            Caption         =   "编制方法"
            Height          =   420
            Left            =   480
            TabIndex        =   10
            Top             =   960
            Width           =   1110
         End
         Begin VB.CheckBox Chk计划类型 
            Caption         =   "计划类型"
            Height          =   420
            Left            =   480
            TabIndex        =   8
            Top             =   420
            Width           =   1110
         End
         Begin VB.TextBox Txt填制人 
            Height          =   300
            Left            =   1770
            MaxLength       =   8
            TabIndex        =   12
            Top             =   1740
            Width           =   1845
         End
         Begin VB.TextBox Txt审核人 
            Height          =   300
            Left            =   1770
            MaxLength       =   8
            TabIndex        =   13
            Top             =   2220
            Width           =   1845
         End
         Begin VB.ComboBox Cbo计划类型 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1770
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   480
            Width           =   3615
         End
         Begin VB.Label Lbl填制人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "填制人"
            Height          =   180
            Left            =   690
            TabIndex        =   19
            Top             =   1800
            Width           =   540
         End
         Begin VB.Label Lbl审核人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "审核人"
            Height          =   180
            Left            =   690
            TabIndex        =   20
            Top             =   2280
            Width           =   540
         End
      End
      Begin VB.Frame fra范围 
         Height          =   2850
         Left            =   240
         TabIndex        =   22
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
            Format          =   162267139
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
            Format          =   162267139
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
            Format          =   162267139
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
            Format          =   162267139
            CurrentDate     =   36263
         End
         Begin VB.Label LblNO 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "No"
            Height          =   180
            Left            =   480
            TabIndex        =   16
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
            TabIndex        =   25
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
            TabIndex        =   18
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
            TabIndex        =   24
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
            TabIndex        =   17
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
            TabIndex        =   23
            Top             =   1140
            Width           =   180
         End
      End
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6210
      TabIndex        =   15
      Top             =   930
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6210
      TabIndex        =   14
      Top             =   435
      Width           =   1100
   End
End
Attribute VB_Name = "FrmStuffPlanSearch"
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
Private mstrOthers(0 To 12) As String ' 0-记录状态(计划类型),1-开始单号,2-结束单号,3-材料id,4-对方部门id(或入出类别id或编制方法(计划用)),5-填制人,6-审核人,7-供应商ID,8-生产商,9-开始生产日期,10-结束生产日期,11-开始发票号,12-结束发票号

Public Function GetSearch(ByVal frmMain As Form, _
        ByRef datStart As Date, ByRef datEnd As Date, _
        ByRef datVerifyStart As Date, ByRef datVerifyEnd As Date, _
        ByRef strOthers() As String) As String
    mstrFind = ""
    Set mfrmMain = frmMain
    
    Me.Show vbModal, mfrmMain
    GetSearch = mstrFind
    datStart = mdatStart
    datEnd = mdatEnd
    datVerifyStart = mdatVerifyStart
    datVerifyEnd = mdatVerifyEnd
    strOthers = mstrOthers
    
End Function


Private Sub cbo编制方法_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Cbo计划类型_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chk编制方法_Click()
    cbo编制方法.Enabled = IIf(chk编制方法.Value = 1, True, False)
End Sub

Private Sub chk编制方法_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub Chk计划类型_Click()
    Cbo计划类型.Enabled = IIf(Chk计划类型.Value = 1, True, False)
End Sub

Private Sub Chk计划类型_GotFocus()
    If sstFilter.Tab = 0 Then
        sstFilter.Tab = 1
        Chk计划类型.SetFocus
    End If
    
End Sub

Private Sub Chk计划类型_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
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

Private Sub chk填制_Click()
    dtp开始时间(0).Enabled = IIf(chk填制.Value = 1, True, False)
    dtp结束时间(0).Enabled = IIf(chk填制.Value = 1, True, False)
    
End Sub

Private Sub chk审核_Click()
    dtp开始时间(1).Enabled = IIf(chk审核.Value = 1, True, False)
    dtp结束时间(1).Enabled = IIf(chk审核.Value = 1, True, False)
End Sub


Private Sub chk填制_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub



Private Sub Cmd取消_Click()
    mstrFind = ""
    Unload Me
End Sub

Private Sub cmd确定_Click()
    '检查数据
    
    If chk填制.Value = 0 And chk审核.Value = 0 Then
        MsgBox "对不起，必须选择一个填制日期或者审核日期!", vbInformation, gstrSysName
        chk填制.SetFocus
        Exit Sub
    End If
    
    mstrFind = ""
    '基本查询条件
    
    Dim i As Integer
    For i = 0 To 12
        mstrOthers(i) = ""
    Next
    
    mdatVerifyStart = Format("1901-01-01", "yyyy-mm-dd")
    mdatVerifyEnd = Format("1901-01-01", "yyyy-mm-dd")
    mdatStart = Format("1901 - 01 - 01", "yyyy-mm-dd")
    mdatEnd = Format("1901-01-01", "yyyy-mm-dd")

    If chk填制.Value = 1 And chk审核.Value = 1 Then
        mstrFind = " And ((A.编制日期 Between [2] And [3] and 审核日期 is null)  or (A.审核日期 Between [4] And [5]))"
                    
        mdatStart = Format(dtp开始时间(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp结束时间(0), "yyyy-mm-dd")
                
        mdatVerifyStart = Format(dtp开始时间(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp结束时间(1), "yyyy-mm-dd")
        
    ElseIf chk审核.Value = 1 Then
        mstrFind = " And A.审核日期 Between [4] And [5] "
        mdatVerifyStart = Format(dtp开始时间(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp结束时间(1), "yyyy-mm-dd")
    ElseIf chk填制.Value = 1 Then
        mstrFind = " And (A.编制日期 Between [2] And [3]) and 审核日期 is null "
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
    
    If Chk计划类型.Value = 1 Then
        mstrFind = mstrFind & " And A.计划类型=[6]"
         mstrOthers(0) = Cbo计划类型.ListIndex + 1
    End If
    ' 0-记录状态(计划类型),1-开始单号,2-结束单号,3-材料id,4-对方部门id(或入出类别id或编制方式),5-填制人,
    ' 6-审核人,7-供应商ID,8-生产商,9-开始生产日期,10-结束生产日期,11-开始发票号,12-结束发票号

    If chk编制方法.Value = 1 Then
        mstrFind = mstrFind & " And A.编制方法=[10] "
         mstrOthers(4) = cbo编制方法.ListIndex + 1
    End If
    
    If Me.Txt填制人 <> "" Then
        mstrFind = mstrFind & " And A.编制人 like [11]"
        mstrOthers(5) = Txt填制人.Text & "%'"
    End If
    
    If Me.Txt审核人 <> "" Then
        mstrFind = mstrFind & " And A.审核人 like [12]"
        mstrOthers(6) = Txt审核人.Text & "%'"
    End If
    
    Unload Me
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
    
    Me.dtp结束时间(0) = Sys.Currentdate
    Me.dtp结束时间(1) = Me.dtp结束时间(0)
    Me.dtp开始时间(0) = DateAdd("d", -7, Me.dtp结束时间(0))
    Me.dtp开始时间(1) = Me.dtp开始时间(0)
    
    sstFilter.Tab = 0
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
    With sstFilter
        If .Tab = 1 Then
            BlnAdvance = True
            If Cbo计划类型.ListCount < 1 Then
                With Cbo计划类型
                    .Clear
                    .AddItem "月度计划", 0
                    .AddItem "季度计划", 1
                    .AddItem "年度计划", 2
                    .ListIndex = 0
                End With
                
                With cbo编制方法
                    .Clear
                    .AddItem "往年同期线形参照法", 0
                    .AddItem "临近期间平均参照法", 1
                    .AddItem "材料储备定额参照法", 2
                    .ListIndex = 0
                End With
            End If
        End If
    End With
End Sub

Private Sub txt结束NO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng库房ID As Long
    
    lng库房ID = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    If KeyCode = vbKeyReturn Then
        If Len(txt结束NO) < 8 And Len(txt结束NO) > 0 Then
            txt结束NO.Text = zlCommFun.GetFullNO(txt结束NO.Text, 77, lng库房ID)
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
            txt开始No.Text = zlCommFun.GetFullNO(txt开始No.Text, 77, lng库房ID)
        End If
        OS.PressKey (vbKeyTab)
    End If
End Sub

Private Sub Txt审核人_GotFocus()
    Txt审核人.SelStart = 0
    Txt审核人.SelLength = Len(Txt审核人.Text)
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
            "   Where (姓名 like [1] or 编号 like [1] or 简码 like [1] ) And (站点=[2] or 站点 is null) " & _
            "       and (撤档时间 >= To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)" & _
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
                    .Top = sstFilter.Top + fra附加条件.Top + Txt审核人.Top + Txt审核人.Height
                    If .Top + .Height > Me.ScaleHeight Then .Top = .Top - Txt审核人.Height - .Height
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
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt填制人_GotFocus()
    Txt填制人.SelStart = 0
    Txt填制人.SelLength = Len(Txt填制人.Text)
End Sub

Private Sub Txt填制人_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyReturn Then Me.Txt审核人.SetFocus
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
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, IIf(gstrMatchMethod = "0", "%", "") & Me.Txt填制人 & "%", gstrNodeNo)
        
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

