VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmDrugQualitySearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤"
   ClientHeight    =   4260
   ClientLeft      =   3150
   ClientTop       =   3165
   ClientWidth     =   7410
   Icon            =   "frmDrugQualitySearch.frx":0000
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
      Height          =   2175
      Left            =   720
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3360
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3836
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
      TabPicture(0)   =   "frmDrugQualitySearch.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra范围"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "附加条件(&D)"
      TabPicture(1)   =   "frmDrugQualitySearch.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra附加条件"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fra附加条件 
         Height          =   2850
         Left            =   -74760
         TabIndex        =   25
         Top             =   600
         Width           =   5520
         Begin VB.CheckBox Chk供应商 
            Caption         =   "供药单位(&S)"
            Height          =   300
            Left            =   480
            TabIndex        =   9
            Top             =   960
            Width           =   1350
         End
         Begin VB.TextBox txt供应商 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1890
            MaxLength       =   50
            TabIndex        =   10
            Top             =   960
            Width           =   3255
         End
         Begin VB.CommandButton Cmd供应商 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   300
            Left            =   5130
            TabIndex        =   11
            Top             =   960
            Width           =   255
         End
         Begin VB.CommandButton Cmd药品 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   300
            Left            =   5130
            TabIndex        =   8
            Top             =   480
            Width           =   255
         End
         Begin VB.TextBox Txt药品 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1890
            MaxLength       =   50
            TabIndex        =   7
            Top             =   480
            Width           =   3255
         End
         Begin VB.CheckBox Chk药品 
            Caption         =   "药品(&P)"
            Height          =   300
            Left            =   480
            TabIndex        =   6
            Top             =   480
            Width           =   990
         End
         Begin VB.TextBox Txt填制人 
            Height          =   300
            Left            =   1890
            MaxLength       =   8
            TabIndex        =   12
            Top             =   1740
            Width           =   1845
         End
         Begin VB.TextBox Txt审核人 
            Height          =   300
            Left            =   1890
            MaxLength       =   8
            TabIndex        =   13
            Top             =   2220
            Width           =   1845
         End
         Begin VB.Label Lbl填制人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "登记人"
            Height          =   180
            Left            =   1170
            TabIndex        =   18
            Top             =   1800
            Width           =   540
         End
         Begin VB.Label Lbl审核人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "处理人"
            Height          =   180
            Left            =   1170
            TabIndex        =   19
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
         Begin VB.CheckBox chk填制 
            Caption         =   "未处理单据"
            Height          =   300
            Left            =   480
            TabIndex        =   0
            Top             =   720
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chk审核 
            Caption         =   "已处理单据"
            Height          =   300
            Left            =   480
            TabIndex        =   3
            Top             =   1560
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker dtp开始时间 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   1
            Top             =   960
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   57671683
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp结束时间 
            Height          =   315
            Index           =   0
            Left            =   3585
            TabIndex        =   2
            Top             =   960
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   57671683
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp开始时间 
            Height          =   315
            Index           =   1
            Left            =   1680
            TabIndex        =   4
            Top             =   1845
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   57671683
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp结束时间 
            Height          =   315
            Index           =   1
            Left            =   3585
            TabIndex        =   5
            Top             =   1845
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   57671683
            CurrentDate     =   36263
         End
         Begin VB.Label lbl时间 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "处理日期"
            Height          =   180
            Index           =   1
            Left            =   900
            TabIndex        =   17
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
            Caption         =   "登记日期"
            Height          =   180
            Index           =   0
            Left            =   900
            TabIndex        =   16
            Top             =   1020
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
            Top             =   1020
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
Attribute VB_Name = "FrmDrugQualitySearch"
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
Private mlng库房ID As Long  '库房id

Private Type Type_SQLCondition
    date填制时间开始 As Date
    date填制时间结束 As Date
    date审核时间开始 As Date
    date审核时间结束 As Date
    lng药品 As Long
    lng供应商 As Long
    str填制人 As String
    str审核人 As String
End Type

Private SQLCondition As Type_SQLCondition
Public Function GetSearch(ByVal FrmMain As Form, _
        ByRef datStart As Date, ByRef datEnd As Date, _
        ByRef datVerifyStart As Date, ByRef datVerifyEnd As Date, _
        ByRef date填制时间开始 As Date, _
        ByRef date填制时间结束 As Date, _
        ByRef date审核时间开始 As Date, _
        ByRef date审核时间结束 As Date, _
        ByRef lng药品 As Long, _
        ByRef lng供应商 As Long, _
        ByRef str填制人 As String, _
        ByRef str审核人 As String, _
        ByVal lng库房ID As Long) As String
    mstrFind = ""
    mlng库房ID = lng库房ID
    Set mfrmMain = FrmMain
    
    Me.Show vbModal, mfrmMain
    GetSearch = mstrFind
    datStart = mdatStart
    datEnd = mdatEnd
    datVerifyStart = mdatVerifyStart
    datVerifyEnd = mdatVerifyEnd

    date填制时间开始 = SQLCondition.date填制时间开始
    date填制时间结束 = SQLCondition.date填制时间结束
    date审核时间开始 = SQLCondition.date审核时间开始
    date审核时间结束 = SQLCondition.date审核时间结束
    lng药品 = SQLCondition.lng药品
    lng供应商 = SQLCondition.lng供应商
    str审核人 = SQLCondition.str审核人
    str填制人 = SQLCondition.str填制人
    
End Function



Private Sub Chk供应商_Click()
    txt供应商.Enabled = IIf(Chk供应商.Value = 1, True, False)
    Cmd供应商.Enabled = txt供应商.Enabled
End Sub

Private Sub Chk供应商_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
End Sub

Private Sub chk审核_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
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

Private Sub Chk药品_Click()
    Txt药品.Enabled = IIf(Chk药品.Value = 1, True, False)
    cmd药品.Enabled = Txt药品.Enabled
    
End Sub

Private Sub Chk药品_GotFocus()
    If sstFilter.Tab = 0 Then
        sstFilter.Tab = 1
        Chk药品.SetFocus
    End If
    
    
End Sub

Private Sub Chk药品_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys vbTab
    
End Sub

Private Sub Cmd供应商_Click()
    Dim rsProvider As New Recordset
    Dim vRect As RECT
    
    vRect = zlControl.GetControlRect(txt供应商.hWnd)
    
    On Error GoTo errHandle
    gstrSQL = "Select id,上级ID,末级,编码,简码,名称 " & _
              "From 供应商 Where (站点 = [1] Or 站点 is Null) " & _
              " And (substr(类型,1,1)=1 Or Nvl(末级,0)=0) " & _
              "Start with 上级ID is null connect by prior ID =上级ID Order by level,ID"
'    Set rsProvider = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "-供应商", gstrNodeNo)
    Set rsProvider = zlDataBase.ShowSQLSelect(Me, gstrSQL, 1, "供应商", True, "", "", False, False, _
            True, vRect.Left, vRect.Top, 300, False, False, True, gstrNodeNo)
    
    If rsProvider.State = 0 Then
        Exit Sub
    End If
    
    If rsProvider.EOF Then
        rsProvider.Close
        Exit Sub
    End If
 
    txt供应商.Tag = rsProvider!id
    txt供应商.Text = rsProvider!名称
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Cmd取消_Click()
    mstrFind = ""
    Unload Me
End Sub

Private Sub Cmd确定_Click()
    '检查数据
    
    If chk填制.Value = 0 And chk审核.Value = 0 Then
        MsgBox "对不起，必须选择一个登记日期或者处理日期!", vbInformation, gstrSysName
        chk填制.SetFocus
        Exit Sub
    End If
    If Chk药品.Value = 1 Then
        If Txt药品.Tag = 0 Then
            MsgBox "请选择需查询的药品信息！", vbInformation, gstrSysName
            Me.Txt药品.SetFocus
            Exit Sub
        End If
    End If
    If Chk供应商.Value = 1 Then
        If txt供应商.Tag = 0 Then
            MsgBox "请选择需查询的药品供应商信息！", vbInformation, gstrSysName
            Me.txt供应商.SetFocus
            Exit Sub
        End If
    End If
    
    
    mstrFind = ""
    '基本查询条件
    Dim i As Integer
    
    If chk填制.Value = 1 And chk审核.Value = 1 Then
        mstrFind = " And ((A.登记时间 Between [1] And [2] and 处理时间 is null) " _
                    & " or (A.处理时间 Between [3] And [4]))"
        mdatStart = Format(dtp开始时间(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp结束时间(0), "yyyy-mm-dd")
                
        mdatVerifyStart = Format(dtp开始时间(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp结束时间(1), "yyyy-mm-dd")
        
    ElseIf chk审核.Value = 1 Then
        mstrFind = " And A.处理时间 Between [3] And [4] "
        mdatVerifyStart = Format(dtp开始时间(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp结束时间(1), "yyyy-mm-dd")
        mdatStart = Format("1901 - 01 - 01", "yyyy-mm-dd")
        mdatEnd = Format("1901-01-01", "yyyy-mm-dd")
    ElseIf chk填制.Value = 1 Then
        mstrFind = " And (A.登记时间 Between [1] And [2]) and 处理时间 is null "
        mdatStart = Format(dtp开始时间(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp结束时间(0), "yyyy-mm-dd")
        
        mdatVerifyStart = Format("1901-01-01", "yyyy-mm-dd")
        mdatVerifyEnd = Format("1901-01-01", "yyyy-mm-dd")
    End If
    
    SQLCondition.date填制时间开始 = CDate(Format(dtp开始时间(0), "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date填制时间结束 = CDate(Format(dtp结束时间(0), "yyyy-mm-dd") & " 23:59:59")
    SQLCondition.date审核时间开始 = CDate(Format(dtp开始时间(1), "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date审核时间结束 = CDate(Format(dtp结束时间(1), "yyyy-mm-dd") & " 23:59:59")
    
    '扩展查询条件
    
    If BlnAdvance = False Then
        Unload Me
        Exit Sub
    End If
    
    If Chk药品.Value = 1 Then mstrFind = mstrFind & " And A.药品id=[5] "
    If Chk供应商.Value = 1 Then mstrFind = mstrFind & " And A.供药单位id +0 =[6]"
    If Me.Txt审核人 <> "" Then mstrFind = mstrFind & " And A.处理人 like [8] "
    If Me.Txt填制人 <> "" Then mstrFind = mstrFind & " And A.登记人 like [7] "
    
    SQLCondition.lng药品 = Val(Txt药品.Tag)
    SQLCondition.lng供应商 = Val(txt供应商.Tag)
    SQLCondition.str审核人 = Me.Txt审核人 & "%"
    SQLCondition.str填制人 = Me.Txt填制人 & "%"
        
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
    Txt药品.Tag = 0
    txt供应商.Tag = 0
    
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
    Call ReleaseSelectorRS
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
            chk填制.SetFocus
        Else
            Chk药品.SetFocus
        End If
    End If
    
End Sub

Private Sub txt供应商_GotFocus()
    txt供应商.SelStart = 0
    txt供应商.SelLength = Len(txt供应商.Text)
End Sub

Private Sub Txt审核人_GotFocus()
    Txt审核人.SelStart = 0
    Txt审核人.SelLength = Len(Txt审核人.Text)
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
        
        gstrSQL = "Select 编号,简码,姓名 From 人员表 " & _
                  "Where (站点 = [3] Or 站点 is Null) And (upper(姓名) like [1] or Upper(编号) like [1] or Upper(简码) like [2]) " & _
                  "  And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[取审核人]", _
                        IIf(gstrMatchMethod = "0", "%", "") & Me.Txt审核人 & "%", _
                        Me.Txt审核人 & "%", gstrNodeNo)
        
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
                    .Height = Me.ScaleHeight - sstFilter.Top - fra附加条件.Top - Txt审核人.Top - Txt审核人.Height - 50
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
'    If KeyCode = vbKeyReturn Then Me.Txt审核人.SetFocus
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt填制人.Text) = "" Then
            Txt审核人.SetFocus
            Exit Sub
        End If
        Txt填制人.Text = UCase(Txt填制人.Text)
        
        gstrSQL = "Select 编号,简码,姓名 From 人员表 " & _
                  "Where (站点 = [3] Or 站点 is Null) And (upper(姓名) like [1] or Upper(编号) like [1] or Upper(简码) like [2]) " & _
                  "  And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) "
        Set rsTemp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption & "[取填制人]", _
                        IIf(gstrMatchMethod = "0", "%", "") & Me.Txt填制人 & "%", _
                        Me.Txt填制人 & "%", gstrNodeNo)
        
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
                    .Height = Me.ScaleHeight - sstFilter.Top - fra附加条件.Top - Txt填制人.Top - Txt填制人.Height - 50
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
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt药品_GotFocus()
    Txt药品.SelStart = 0
    Txt药品.SelLength = Len(Txt药品.Text)
End Sub

Private Sub Txt药品_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim RecReturn As Recordset
    Dim strkey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(Txt药品.Text) = "" Then Exit Sub
    sngLeft = Me.Left + sstFilter.Left + fra附加条件.Left + Txt药品.Left
    sngTop = Me.Top + sstFilter.Top + fra附加条件.Top + Txt药品.Top + Txt药品.Height + Me.Height - Me.ScaleHeight  '  50
    If sngTop + 3630 > Screen.Height Then
        sngTop = sngTop - Txt药品.Height - 3630
    End If
    
    strkey = Trim(Txt药品.Text)
    If Mid(strkey, 1, 1) = "[" Then
        If InStr(2, strkey, "]") <> 0 Then
            strkey = Mid(strkey, 2, InStr(2, strkey, "]") - 2)
        Else
            strkey = Mid(strkey, 2)
        End If
    End If
    
    Call SetSelectorRS(1, "药品质量管理", mlng库房ID, mlng库房ID, mlng库房ID, , , True)
    
'    Set RecReturn = Frm药品多选选择器.ShowME(Me, 1, , , UserInfo.部门ID, strkey, sngLeft, sngTop)
    Set RecReturn = frmSelector.showMe(Me, 1, 1, strkey, sngLeft, sngTop, mlng库房ID, mlng库房ID, mlng库房ID, , , , , 2, False)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    
    If gint药品名称显示 = 1 Then
        Txt药品.Text = "[" & RecReturn!药品编码 & "]" & IIf(IsNull(RecReturn!商品名), RecReturn!通用名, RecReturn!商品名)
    Else
        Txt药品.Text = "[" & RecReturn!药品编码 & "]" & RecReturn!通用名
    End If

    Txt药品.Tag = RecReturn!药品id
    
    If Chk供应商.Visible = True Then
        If Chk供应商.Value = 1 Then
            txt供应商.SetFocus
        Else
            Chk供应商.SetFocus
        End If
    End If
End Sub

Private Sub Cmd药品_Click()
    Dim RecReturn As Recordset
    
    Call SetSelectorRS(1, "药品质量管理", mlng库房ID, mlng库房ID, mlng库房ID, , , True)

'    Set RecReturn = Frm药品选择器.ShowME(Me, 1, , , UserInfo.部门ID)
    Set RecReturn = frmSelector.showMe(Me, 0, 1, , , , mlng库房ID, mlng库房ID, mlng库房ID, , , , , 2, False)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    
    If gint药品名称显示 = 1 Then
        Txt药品.Text = "[" & RecReturn!药品编码 & "]" & IIf(IsNull(RecReturn!商品名), RecReturn!通用名, RecReturn!商品名)
    Else
        Txt药品.Text = "[" & RecReturn!药品编码 & "]" & RecReturn!通用名
    End If
    
    Txt药品.Tag = RecReturn!药品id
    
    If Chk供应商.Visible = True Then
        If Chk供应商.Value = 1 Then
            txt供应商.SetFocus
        Else
            Chk供应商.SetFocus
        End If
    End If
End Sub


Private Sub txt供应商_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RecTmp As New Recordset
    Dim vRect As RECT, blnCancel As Boolean
    
    On Error GoTo errHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    If LTrim(RTrim(txt供应商)) <> "" Then
        txt供应商 = UCase(txt供应商)
        vRect = zlControl.GetControlRect(txt供应商.hWnd)
        
        gstrSQL = "Select id,编码,简码,名称 From 供应商 " & _
                  "Where (站点 = [2] Or 站点 is Null) " & _
                  "  And 末级=1 And (substr(类型,1,1)=1 Or Nvl(末级,0)=0)" & _
                  "  And (编码 like [1] or 简码 like [1] or 名称 like [1] )"
'        Set RecTmp = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, IIf(gstrMatchMethod = "0", "%", "") & txt供应商 & "%", gstrNodeNo)
        Set RecTmp = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "产地", False, "", "", False, False, _
                        True, vRect.Left, vRect.Top, 300, blnCancel, False, True, IIf(gstrMatchMethod = "0", "%", "") & txt供应商 & "%", gstrNodeNo)
        
        
        If blnCancel Then txt供应商.SetFocus: Exit Sub
        
        If RecTmp.State = 0 Then
            MsgBox "输入值无效！", vbInformation, gstrSysName
            KeyCode = 0
            txt供应商.Tag = 0
            txt供应商.SelStart = 0
            txt供应商.SelLength = Len(txt供应商.Text)
            Exit Sub
        End If
        
        txt供应商 = RecTmp!名称
        txt供应商.Tag = RecTmp!id
                  
'        With RecTmp
'            If .EOF Then
'                MsgBox "输入值无效！", vbInformation, gstrSysName
'                KeyCode = 0
'                txt供应商.Tag = 0
'                txt供应商.SelStart = 0
'                txt供应商.SelLength = Len(txt供应商.Text)
'
'                Exit Sub
'            End If
'            If .RecordCount > 1 Then
'                mstrSelectTag = "Provider"
'                Set mshSelect.Recordset = RecTmp
'                With mshSelect
'                    .Top = sstFilter.Top + fra附加条件.Top + txt供应商.Top + txt供应商.Height
'                    .Left = sstFilter.Left + fra附加条件.Left + txt供应商.Left
'                    .Visible = True
'                    .SetFocus
'                    .ColWidth(0) = 0
'                    .ColWidth(1) = 800
'                    .ColWidth(2) = 800
'
'                    .ColWidth(3) = .Width - .ColWidth(1) - .ColWidth(2)
'                    .Row = 1
'                    .Col = 0
'                    .ColSel = .Cols - 1
'                    Exit Sub
'
'                End With
'            Else
'                txt供应商 = !名称
'                txt供应商.Tag = !id
'
'            End If
'        End With
    End If
    
    Txt填制人.SetFocus
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


