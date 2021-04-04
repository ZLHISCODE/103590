VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmDrugPaymentSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤"
   ClientHeight    =   4200
   ClientLeft      =   3150
   ClientTop       =   3165
   ClientWidth     =   7515
   Icon            =   "frmDrugPaymentSearch.frx":0000
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
      Height          =   2535
      Left            =   2760
      TabIndex        =   17
      Top             =   3600
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
      Height          =   3975
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   7011
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "范围(&R)"
      TabPicture(0)   =   "frmDrugPaymentSearch.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra范围"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "附加条件(&D)"
      TabPicture(1)   =   "frmDrugPaymentSearch.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra附加条件"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fra附加条件 
         Height          =   2715
         Left            =   -74760
         TabIndex        =   26
         Top             =   600
         Width           =   5505
         Begin VB.CheckBox Chk供应商 
            Caption         =   "供应商"
            Height          =   300
            Left            =   480
            TabIndex        =   9
            Top             =   540
            Width           =   870
         End
         Begin VB.TextBox txt供应商 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   10
            Top             =   540
            Width           =   3375
         End
         Begin VB.CommandButton Cmd供应商 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4890
            TabIndex        =   11
            Top             =   540
            Width           =   255
         End
         Begin VB.TextBox Txt填制人 
            Height          =   300
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   12
            Top             =   1260
            Width           =   1365
         End
         Begin VB.TextBox Txt审核人 
            Height          =   300
            Left            =   1500
            MaxLength       =   8
            TabIndex        =   13
            Top             =   1860
            Width           =   1365
         End
         Begin VB.Label Lbl填制人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "填制人"
            Height          =   180
            Left            =   750
            TabIndex        =   28
            Top             =   1320
            Width           =   540
         End
         Begin VB.Label Lbl审核人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "审核人"
            Height          =   180
            Left            =   750
            TabIndex        =   27
            Top             =   1920
            Width           =   540
         End
      End
      Begin VB.Frame fra范围 
         Height          =   2850
         Left            =   240
         TabIndex        =   19
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
            Format          =   108396547
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
            Format          =   108396547
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
            Format          =   108396547
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
            Format          =   108396547
            CurrentDate     =   36263
         End
         Begin VB.Label LblNO 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "No"
            Height          =   180
            Left            =   480
            TabIndex        =   25
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
            TabIndex        =   24
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
            TabIndex        =   23
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
            TabIndex        =   22
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
            TabIndex        =   20
            Top             =   1140
            Width           =   180
         End
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   6330
      TabIndex        =   16
      Top             =   3720
      Width           =   1100
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6330
      TabIndex        =   15
      Top             =   930
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6330
      TabIndex        =   14
      Top             =   435
      Width           =   1100
   End
End
Attribute VB_Name = "FrmDrugPaymentSearch"
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
Private mstrSelectTag As String     '当前选择的对象

Public Function GetSearch(ByVal FrmMain As Form, _
        ByRef datStart As Date, ByRef datEnd As Date, _
        ByRef datVerifyStart As Date, ByRef datVerifyEnd As Date) As String
    mstrFind = ""
    If Not CheckCompete Then Exit Function
    
    Me.Show vbModal, FrmMain
    GetSearch = mstrFind
    datStart = mdatStart
    datEnd = mdatEnd
    datVerifyStart = mdatVerifyStart
    datVerifyEnd = mdatVerifyEnd
    
End Function


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
    If KeyCode = 13 Then
        SendKeys vbTab
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


Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub Cmd供应商_Click()
    Dim rsProvider As New Recordset
    
    On Error GoTo errHandle
    gstrSQL = "Select id,上级ID,末级,编码,简码,名称 From 药品供应商 Where " & _
                " To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01'  start with 上级ID is null connect by prior ID =上级ID order by level,ID"
    Call SQLTest(App.Title, Me.Caption, gstrSQL)
    Set rsProvider = zldatabase.OpenSQLRecord(gstrSQL, "Cmd供应商_Click")
    Call SQLTest
    
    If rsProvider.EOF Then
        rsProvider.Close
        Exit Sub
    End If
    With FrmSelect
        Set .TreeRec = rsProvider
        .StrNode = "所有药品供应商"
        .lngMode = 0
        .Show 1, Me
        If .BlnSuccess = True Then
            txt供应商.Tag = .CurrentID
            txt供应商.Text = .CurrentName
        End If
    End With
    Unload FrmSelect
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
    If Chk供应商.Value = 1 Then
        If txt供应商.Tag = 0 Then
            MsgBox "请选择需查询的药品供应商信息！", vbInformation, gstrSysName
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
    '基本查询条件
    Dim i As Integer
    
    If chk填制.Value = 1 And chk审核.Value = 1 Then
        If chkStrike.Value = 1 Then
            mstrFind = " And ((A.填制日期 Between To_Date('" & Format(dtp开始时间(0), "yyyy-mm-dd") & "00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp结束时间(0), "yyyy-mm-dd") & "23:59:59','YYYY-MM-DD HH24:MI:SS')) " _
                    & " or (A.审核日期 Between To_Date('" & Format(dtp开始时间(1), "yyyy-mm-dd") & "00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp结束时间(1), "yyyy-mm-dd") & "23:59:59','YYYY-MM-DD HH24:MI:SS')))"
        Else
            mstrFind = " And ((A.填制日期 Between To_Date('" & Format(dtp开始时间(0), "yyyy-mm-dd") & "00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp结束时间(0), "yyyy-mm-dd") & "23:59:59','YYYY-MM-DD HH24:MI:SS')) " _
                    & " or (A.审核日期 Between To_Date('" & Format(dtp开始时间(1), "yyyy-mm-dd") & "00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp结束时间(1), "yyyy-mm-dd") & "23:59:59','YYYY-MM-DD HH24:MI:SS'))) and a.记录状态 =1 "
        End If
        mdatStart = Format(dtp开始时间(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp结束时间(0), "yyyy-mm-dd")
                
        mdatVerifyStart = Format(dtp开始时间(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp结束时间(1), "yyyy-mm-dd")
        
    ElseIf chk审核.Value = 1 Then
        If chkStrike.Value = 1 Then
            mstrFind = " And A.审核日期 Between To_Date('" & Format(dtp开始时间(1), "yyyy-mm-dd") & "00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp结束时间(1), "yyyy-mm-dd") & "23:59:59','YYYY-MM-DD HH24:MI:SS') "
        Else
            mstrFind = " And A.审核日期 Between To_Date('" & Format(dtp开始时间(1), "yyyy-mm-dd") & "00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp结束时间(1), "yyyy-mm-dd") & "23:59:59','YYYY-MM-DD HH24:MI:SS') and a.记录状态 =1 "
            
        End If
        mdatVerifyStart = Format(dtp开始时间(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp结束时间(1), "yyyy-mm-dd")
        mdatStart = Format("1901 - 01 - 01", "yyyy-mm-dd")
        mdatEnd = Format("1901-01-01", "yyyy-mm-dd")
    ElseIf chk填制.Value = 1 Then
        mstrFind = " And (A.填制日期 Between To_Date('" & Format(dtp开始时间(0), "yyyy-mm-dd") & "00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp结束时间(0), "YYYY-mm-dd") & "23:59:59 ','YYYY-MM-DD HH24:MI:SS')) and 审核日期 is null "
        mdatStart = Format(dtp开始时间(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp结束时间(0), "yyyy-mm-dd")
        
        mdatVerifyStart = Format("1901-01-01", "yyyy-mm-dd")
        mdatVerifyEnd = Format("1901-01-01", "yyyy-mm-dd")
    End If
    
    Dim intYear As Integer, strYear As String
    
    If Len(txt开始No) < 8 And Len(txt开始No) > 0 Then
        Me.txt开始No = UCase(LTrim(Me.txt开始No))
        intYear = Format(zldatabase.Currentdate, "YYYY") - 1990
        strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
        If Len(txt开始No) < 8 Then Me.txt开始No = strYear & String(7 - Len(txt开始No), "0") & Me.txt开始No
    End If
    If Len(txt结束NO) < 8 And Len(txt结束NO) > 0 Then
        Me.txt结束NO = UCase(LTrim(Me.txt结束NO))
        intYear = Format(zldatabase.Currentdate, "YYYY") - 1990
        strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
        If Len(txt结束NO) < 8 Then Me.txt结束NO = strYear & String(7 - Len(txt结束NO), "0") & Me.txt结束NO
    End If
    
    If Me.txt开始No <> "" And Me.txt结束NO <> "" Then mstrFind = mstrFind & " And A.No >= '" & Me.txt开始No & "' And A.No <='" & Me.txt结束NO & "'"
    If Me.txt开始No <> "" And Me.txt结束NO = "" Then mstrFind = mstrFind & " And A.No >= '" & Me.txt开始No & "'"
    If Me.txt开始No = "" And Me.txt结束NO <> "" Then mstrFind = mstrFind & " And A.No <= '" & Me.txt结束NO & "'"
    
    '扩展查询条件
    
    If BlnAdvance = False Then
        Unload Me
        Exit Sub
    End If
    
    If Chk供应商.Value = 1 Then
        mstrFind = mstrFind & " and a.单位id=" & txt供应商.Tag
    End If
    
    If Me.Txt审核人 <> "" Then mstrFind = mstrFind & " And A.审核人 like '" & Me.Txt审核人 & "%'"
    If Me.Txt填制人 <> "" Then mstrFind = mstrFind & " And A.填制人 like '" & Me.Txt填制人 & "%'"
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
    Me.dtp结束时间(0) = zldatabase.Currentdate
    Me.dtp结束时间(1) = Me.dtp结束时间(0)
    
    Me.dtp开始时间(0) = DateAdd("m", -1, Me.dtp结束时间(0))
    Me.dtp开始时间(1) = Me.dtp开始时间(0)
    
    Me.txt供应商.Tag = 0
    '打开记录集
    sstFilter.Tab = 0
    BlnAdvance = False
    
End Sub

Private Function CheckCompete() As Boolean
    Dim rsCompete As New Recordset
    
    On Error GoTo errHandle
    CheckCompete = False
    With rsCompete
        If .State = 1 Then .Close
        gstrSQL = "Select id,上级ID,编码,简码,末级,名称 From 药品供应商 Where 名称 is Not NULL And " & _
              " To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' Start with 上级ID is NULL Connect by prior id=上级id"
        Call SQLTest(App.Title, "过滤", gstrSQL)
        Set rsCompete = zldatabase.OpenSQLRecord(gstrSQL, "cmd产地_Click")
        Call SQLTest
        
        If .EOF Then
            .Close
            MsgBox "药品供应商信息不全，请在供药单位管理中设置药品供应商信息！", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    CheckCompete = True
    Exit Function
errHandle:
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

Private Sub txt供应商_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RecTmp As New Recordset
    
    On Error GoTo errHandle
    If KeyCode <> vbKeyReturn Then Exit Sub
    If LTrim(RTrim(txt供应商)) <> "" Then
        txt供应商 = UCase(txt供应商)
        With RecTmp
            gstrSQL = "Select id,编码,简码,名称 From 药品供应商 Where " & _
                  " To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' And 末级=1" & _
                  " And (编码 like '" & IIf(gstrMatchMethod = "0", "%", "") & txt供应商 & "%' or 简码 like '" & IIf(gstrMatchMethod = "0", "%", "") & txt供应商 & "%' or 名称 like '" & IIf(gstrMatchMethod = "0", "%", "") & txt供应商 & "%')"
            Call OpenRecordset(RecTmp, "药品供应商")
            If .EOF Then
                MsgBox "输入值无效！", vbInformation, gstrSysName
                KeyCode = 0
                txt供应商.Tag = 0
                txt供应商.SelStart = 0
                txt供应商.SelLength = Len(txt供应商.Text)
                
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Provider"
                Set mshSelect.Recordset = RecTmp
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
    
    Txt填制人.SetFocus
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt供应商_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt结束NO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Len(txt结束NO) < 8 And Len(txt结束NO) > 0 Then
            Dim intYear  As Integer, strYear As String
            Me.txt结束NO = UCase(LTrim(Me.txt结束NO))
            intYear = Format(zldatabase.Currentdate, "YYYY") - 1990
            strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
            If Len(txt结束NO) < 8 Then Me.txt结束NO = strYear & String(7 - Len(txt结束NO), "0") & Me.txt结束NO
        End If
        zlcommfun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txt结束NO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub txt开始No_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Len(txt开始No) < 8 And Len(txt开始No) > 0 Then
            Dim intYear  As Integer, strYear As String
            Me.txt开始No = UCase(LTrim(Me.txt开始No))
            intYear = Format(zldatabase.Currentdate, "YYYY") - 1990
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
        With rsTemp
            gstrSQL = "Select 编号,简码,姓名 From 人员表 Where upper(姓名) like '" & IIf(gstrMatchMethod = "0", "%", "") & Me.Txt审核人 & "%' or Upper(编号) like '" & IIf(gstrMatchMethod = "0", "%", "") & Me.Txt审核人 & "%' or Upper(简码) like '" & Me.Txt审核人 & "%'"
            Call OpenRecordset(rsTemp, "取审核人")
            
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
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt审核人_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt填制人_GotFocus()
    Txt填制人.SelStart = 0
    Txt填制人.SelLength = Len(Txt填制人.Text)
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
        With rsTemp
            gstrSQL = "Select 编号,简码,姓名 From 人员表 Where upper(姓名) like '" & IIf(gstrMatchMethod = "0", "%", "") & Me.Txt填制人 & "%' or Upper(编号) like '" & IIf(gstrMatchMethod = "0", "%", "") & Me.Txt填制人 & "%' or Upper(简码) like '" & Me.Txt填制人 & "%'"
            Call OpenRecordset(rsTemp, "取填制人")
            
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
