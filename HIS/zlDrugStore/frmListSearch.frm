VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmListSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤"
   ClientHeight    =   4260
   ClientLeft      =   3156
   ClientTop       =   3168
   ClientWidth     =   7692
   Icon            =   "frmListSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   7692
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSelect 
      Height          =   2535
      Left            =   960
      TabIndex        =   29
      Top             =   3090
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7853
      _ExtentY        =   4466
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
      TabIndex        =   23
      Top             =   120
      Width           =   6015
      _ExtentX        =   10605
      _ExtentY        =   7006
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "范围(&R)"
      TabPicture(0)   =   "frmListSearch.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra范围"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "附加条件(&D)"
      TabPicture(1)   =   "frmListSearch.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra附加条件"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame fra附加条件 
         Height          =   2850
         Left            =   -74760
         TabIndex        =   28
         Top             =   600
         Width           =   5520
         Begin VB.CheckBox Chk移入库房 
            Caption         =   "移出库房"
            Height          =   420
            Left            =   360
            TabIndex        =   16
            Top             =   900
            Width           =   1110
         End
         Begin VB.CommandButton Cmd药品 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   300
            Left            =   4890
            TabIndex        =   22
            Top             =   420
            Width           =   255
         End
         Begin VB.TextBox Txt药品 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1530
            MaxLength       =   50
            TabIndex        =   15
            Top             =   420
            Width           =   3375
         End
         Begin VB.CheckBox Chk药品 
            Caption         =   "药品"
            Height          =   300
            Left            =   360
            TabIndex        =   14
            Top             =   420
            Width           =   990
         End
         Begin VB.TextBox Txt填制人 
            Height          =   300
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   19
            Top             =   1500
            Width           =   1845
         End
         Begin VB.TextBox Txt审核人 
            Height          =   300
            Left            =   1530
            MaxLength       =   8
            TabIndex        =   21
            Top             =   1980
            Width           =   1845
         End
         Begin VB.ComboBox Cbo移入库房 
            Enabled         =   0   'False
            Height          =   276
            Left            =   1530
            TabIndex        =   17
            Text            =   "Cbo移入库房"
            Top             =   960
            Width           =   3615
         End
         Begin VB.Label Lbl填制人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "填制人"
            Height          =   180
            Left            =   570
            TabIndex        =   18
            Top             =   1560
            Width           =   540
         End
         Begin VB.Label Lbl审核人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "审核人"
            Height          =   180
            Left            =   570
            TabIndex        =   20
            Top             =   2040
            Width           =   540
         End
      End
      Begin VB.Frame fra范围 
         Height          =   2850
         Left            =   240
         TabIndex        =   24
         Top             =   600
         Width           =   5520
         Begin VB.TextBox txt开始No 
            Height          =   300
            Left            =   840
            MaxLength       =   8
            TabIndex        =   1
            Top             =   360
            Width           =   1605
         End
         Begin VB.TextBox txt结束NO 
            Height          =   300
            Left            =   2970
            MaxLength       =   8
            TabIndex        =   2
            Top             =   360
            Width           =   1605
         End
         Begin VB.CheckBox chk填制 
            Caption         =   "未审核单据"
            Height          =   300
            Left            =   480
            TabIndex        =   3
            Top             =   840
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox chk审核 
            Caption         =   "已审核单据"
            Height          =   300
            Left            =   480
            TabIndex        =   7
            Top             =   1560
            Width           =   1215
         End
         Begin VB.CheckBox chkStrike 
            Caption         =   "包含冲销"
            Enabled         =   0   'False
            Height          =   300
            Left            =   720
            TabIndex        =   11
            Top             =   2280
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker dtp开始时间 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   5
            Top             =   1080
            Width           =   1605
            _ExtentX        =   2836
            _ExtentY        =   550
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   104333315
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp结束时间 
            Height          =   315
            Index           =   0
            Left            =   3585
            TabIndex        =   6
            Top             =   1080
            Width           =   1605
            _ExtentX        =   2836
            _ExtentY        =   550
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   104333315
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp开始时间 
            Height          =   315
            Index           =   1
            Left            =   1680
            TabIndex        =   9
            Top             =   1845
            Width           =   1605
            _ExtentX        =   2836
            _ExtentY        =   550
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   104333315
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp结束时间 
            Height          =   315
            Index           =   1
            Left            =   3585
            TabIndex        =   10
            Top             =   1845
            Width           =   1605
            _ExtentX        =   2836
            _ExtentY        =   550
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   104333315
            CurrentDate     =   36263
         End
         Begin VB.Label LblNO 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "No"
            Height          =   180
            Left            =   480
            TabIndex        =   0
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
            TabIndex        =   27
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
            TabIndex        =   8
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
            TabIndex        =   26
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
            TabIndex        =   4
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
            TabIndex        =   25
            Top             =   1140
            Width           =   180
         End
      End
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6450
      TabIndex        =   13
      Top             =   930
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6450
      TabIndex        =   12
      Top             =   435
      Width           =   1100
   End
End
Attribute VB_Name = "FrmListSearch"
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
Private mlng库房id As Long  '库房id

Private Type Type_SQLCondition
    strNO开始 As String
    strNO结束 As String
    date填制时间开始 As Date
    date填制时间结束 As Date
    date审核时间开始 As Date
    date审核时间结束 As Date
    lng药品 As Long
    lng移出库房 As Long
    str填制人 As String
    str审核人 As String
End Type

Private SQLCondition As Type_SQLCondition
Public Function GetSearch(ByVal FrmMain As Form, ByVal lngMode As Long, ByVal lng库房id As Long, _
        ByRef datStart As Date, ByRef datEnd As Date, _
        ByRef datVerifyStart As Date, ByRef datVerifyEnd As Date, _
        ByRef strNO开始 As String, _
        ByRef strNO结束 As String, _
        ByRef date填制时间开始 As Date, _
        ByRef date填制时间结束 As Date, _
        ByRef date审核时间开始 As Date, _
        ByRef date审核时间结束 As Date, _
        ByRef lng药品 As Long, _
        ByRef lng移出库房 As Long, _
        ByRef str填制人 As String, _
        ByRef str审核人 As String) As String
    mstrFind = ""
    mlngMode = lngMode
    mlng库房id = lng库房id
    Set mfrmMain = FrmMain
    
    Me.Show vbModal, mfrmMain
    GetSearch = mstrFind
    datStart = mdatStart
    datEnd = mdatEnd
    datVerifyStart = mdatVerifyStart
    datVerifyEnd = mdatVerifyEnd
    
    strNO开始 = SQLCondition.strNO开始
    strNO结束 = SQLCondition.strNO结束
    date填制时间开始 = SQLCondition.date填制时间开始
    date填制时间结束 = SQLCondition.date填制时间结束
    date审核时间开始 = SQLCondition.date审核时间开始
    date审核时间结束 = SQLCondition.date审核时间结束
    lng药品 = SQLCondition.lng药品
    lng移出库房 = SQLCondition.lng移出库房
    str审核人 = SQLCondition.str审核人
    str填制人 = SQLCondition.str填制人

End Function

Private Sub Cbo移入库房_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim str工作性质 As String
    
    str工作性质 = "H,I,J,K,L,M,N"
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Cbo移入库房.ListCount = 0 Then Exit Sub
    
    If Cbo移入库房.ListIndex >= 0 Then
        If Val(Cbo移入库房.Tag) = Cbo移入库房.ItemData(Cbo移入库房.ListIndex) Then
            Exit Sub
        End If
    End If
    
    If Select部门选择器(Me, Cbo移入库房, Trim(Cbo移入库房.Text), str工作性质, , "0,1,2,3") = False Then
        Exit Sub
    End If
    If Cbo移入库房.ListIndex >= 0 Then
        Cbo移入库房.Tag = Cbo移入库房.ItemData(Cbo移入库房.ListIndex)
    End If
End Sub

Private Sub Cbo移入库房_KeyPress(KeyAscii As Integer)
    '屏蔽输入单引号
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Cbo移入库房_Validate(Cancel As Boolean)
    If Cbo移入库房.ListCount > 0 Then
        If Cbo移入库房.ListIndex = -1 Then
            MsgBox "请选择一个药库或者药房！", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub

Private Sub Chk移入库房_click()
    Cbo移入库房.Enabled = IIf(Chk移入库房.Value = 1, True, False)
End Sub

Private Sub Chk移入库房_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Chk移入库房.Value = 1 Then
        Cbo移入库房.SetFocus
    Else
        Txt填制人.SetFocus
    End If
End Sub
Private Sub chk填制_Click()
    Dtp开始时间(0).Enabled = IIf(chk填制.Value = 1, True, False)
    Dtp结束时间(0).Enabled = IIf(chk填制.Value = 1, True, False)
    
End Sub

Private Sub chk审核_Click()
    Dtp开始时间(1).Enabled = IIf(chk审核.Value = 1, True, False)
    Dtp结束时间(1).Enabled = IIf(chk审核.Value = 1, True, False)
    chkStrike.Enabled = IIf(chk审核.Value = 1, True, False)
    
End Sub

Private Sub Chk药品_Click()
    txt药品.Enabled = IIf(Chk药品.Value = 1, True, False)
    cmd药品.Enabled = IIf(Chk药品.Value = 1, True, False)
End Sub



Private Sub Cmd取消_Click()
    mstrFind = ""
    Unload Me
End Sub

Private Sub Cmd确定_Click()
    Dim lng库房id As Long
    Dim intNO As Integer, strNo As String
    
    '初始准备
    intNO = Switch(mlngMode = 1343, 26, mlngMode = 1344, 23)
    lng库房id = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    '检查数据
    If Chk药品.Value = 1 Then
        If txt药品.Tag = 0 Then
            MsgBox "请选择需查询的药品信息！", vbInformation, gstrSysName
            Me.txt药品.SetFocus
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
'            mstrFind = " And ((A.填制日期 Between To_Date('" & Format(dtp开始时间(0), "yyyy-mm-dd") & "00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp结束时间(0), "yyyy-mm-dd") & "23:59:59','YYYY-MM-DD HH24:MI:SS')) " _
'                    & " or (A.审核日期 Between To_Date('" & Format(dtp开始时间(1), "yyyy-mm-dd") & "00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp结束时间(1), "yyyy-mm-dd") & "23:59:59','YYYY-MM-DD HH24:MI:SS')))"
            mstrFind = " And ((A.填制日期 Between [3] And [4] and 审核日期 is null) " _
                    & " or (A.审核日期 Between [5] And [6]))"
        Else
'            mstrFind = " And ((A.填制日期 Between To_Date('" & Format(dtp开始时间(0), "yyyy-mm-dd") & "00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp结束时间(0), "yyyy-mm-dd") & "23:59:59','YYYY-MM-DD HH24:MI:SS')) " _
'                    & " or (A.审核日期 Between To_Date('" & Format(dtp开始时间(1), "yyyy-mm-dd") & "00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp结束时间(1), "yyyy-mm-dd") & "23:59:59','YYYY-MM-DD HH24:MI:SS'))) and a.记录状态 =1 "
            mstrFind = " And ((A.填制日期 Between [3] And [4] and 审核日期 is null) " _
                    & " or (A.审核日期 Between [5] And [6])) and (a.记录状态 =1 or mod(A.记录状态,3)=0) "
        End If
        
        mdatStart = Format(Dtp开始时间(0), "yyyy-mm-dd")
        mdatEnd = Format(Dtp结束时间(0), "yyyy-mm-dd")
                
        mdatVerifyStart = Format(Dtp开始时间(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(Dtp结束时间(1), "yyyy-mm-dd")
        
    ElseIf chk审核.Value = 1 Then
        If chkStrike.Value = 1 Then
'            mstrFind = " And A.审核日期 Between To_Date('" & Format(dtp开始时间(1), "yyyy-mm-dd") & "00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp结束时间(1), "yyyy-mm-dd") & "23:59:59','YYYY-MM-DD HH24:MI:SS') "
            mstrFind = " and (A.记录状态=2 or A.记录状态=1 or mod(A.记录状态,3)=2 or mod(A.记录状态,3)=0) And A.审核日期 Between [5] And [6] "
        Else
'            mstrFind = " And A.审核日期 Between To_Date('" & Format(dtp开始时间(1), "yyyy-mm-dd") & "00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp结束时间(1), "yyyy-mm-dd") & "23:59:59','YYYY-MM-DD HH24:MI:SS') and a.记录状态 =1 "
            mstrFind = " and (a.记录状态 =1 or mod(A.记录状态,3)=0) And A.审核日期 Between [5] And [6] "
        End If
        
        mdatVerifyStart = Format(Dtp开始时间(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(Dtp结束时间(1), "yyyy-mm-dd")
        mdatStart = Format("1901 - 01 - 01", "yyyy-mm-dd")
        mdatEnd = Format("1901-01-01", "yyyy-mm-dd")
    ElseIf chk填制.Value = 1 Then
'        mstrFind = " And (A.填制日期 Between To_Date('" & Format(dtp开始时间(0), "yyyy-mm-dd") & "00:00:00','YYYY-MM-DD HH24:MI:SS') And To_Date('" & Format(dtp结束时间(0), "YYYY-mm-dd") & "23:59:59 ','YYYY-MM-DD HH24:MI:SS')) and 审核日期 is null "
        mstrFind = " And (A.填制日期 Between [3] And [4]) and 审核日期 is null "
            
        mdatStart = Format(Dtp开始时间(0), "yyyy-mm-dd")
        mdatEnd = Format(Dtp结束时间(0), "yyyy-mm-dd")
        
        mdatVerifyStart = Format("1901-01-01", "yyyy-mm-dd")
        mdatVerifyEnd = Format("1901-01-01", "yyyy-mm-dd")
    End If
    
    Dim intYear As Integer, strYear As String
    
    If Len(txt开始NO) < 8 And Len(txt开始NO) > 0 Then
        txt开始NO.Text = GetFullNO(txt开始NO.Text, intNO, lng库房id)
    End If
    If Len(txt结束NO) < 8 And Len(txt结束NO) > 0 Then
        txt结束NO.Text = GetFullNO(txt结束NO.Text, intNO, lng库房id)
    End If
    
'    If Me.txt开始No <> "" And Me.txt结束NO <> "" Then mstrFind = mstrFind & " And A.No >= '" & Me.txt开始No & "' And A.No <='" & Me.txt结束NO & "'"
    If Me.txt开始NO <> "" And Me.txt结束NO <> "" Then mstrFind = mstrFind & " And A.No >= [1] And A.No <=[2] "
    
    
'    If Me.txt开始No <> "" And Me.txt结束NO = "" Then mstrFind = mstrFind & " And A.No >= '" & Me.txt开始No & "'"
    If Me.txt开始NO <> "" And Me.txt结束NO = "" Then mstrFind = mstrFind & " And A.No >= [1] "
       
'    If Me.txt开始No = "" And Me.txt结束NO <> "" Then mstrFind = mstrFind & " And A.No <= '" & Me.txt结束NO & "'"
    If Me.txt开始NO = "" And Me.txt结束NO <> "" Then mstrFind = mstrFind & " And A.No <= [2] "
        
    SQLCondition.strNO开始 = Me.txt开始NO
    SQLCondition.strNO结束 = Me.txt结束NO
    SQLCondition.date填制时间开始 = CDate(Format(Dtp开始时间(0), "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date填制时间结束 = CDate(Format(Dtp结束时间(0), "yyyy-mm-dd") & " 23:59:59")
    SQLCondition.date审核时间开始 = CDate(Format(Dtp开始时间(1), "yyyy-mm-dd") & " 00:00:00")
    SQLCondition.date审核时间结束 = CDate(Format(Dtp结束时间(1), "yyyy-mm-dd") & " 23:59:59")
        
    '扩展查询条件
    
    If BlnAdvance = False Then
        Unload Me
        Exit Sub
    End If
    
    
    If Chk药品.Value = 1 Then
'        mstrFind = mstrFind & " And A.药品ID=" & Txt药品.Tag
        mstrFind = mstrFind & " And A.药品ID + 0=[7] "
    End If
    
    If mlngMode = 1343 Then
'        If Chk移入库房.Value = 1 Then mstrFind = mstrFind & " And A.对方部门ID=" & Cbo移入库房.ItemData(Cbo移入库房.ListIndex)
        If Chk移入库房.Value = 1 Then
            mstrFind = mstrFind & " And A.对方部门ID + 0=[8] "
        End If
    End If
'    If Me.Txt审核人 <> "" Then mstrFind = mstrFind & " And A.审核人 like '" & Me.Txt审核人 & "%'"
    If Me.Txt审核人 <> "" Then
        mstrFind = mstrFind & " And A.审核人 like [10] "
    End If
        
'    If Me.Txt填制人 <> "" Then mstrFind = mstrFind & " And A.填制人 like '" & Me.Txt填制人 & "%'"
    If Me.Txt填制人 <> "" Then
        mstrFind = mstrFind & " And A.填制人 like [9] "
    End If
    
    SQLCondition.lng药品 = Val(txt药品.Tag)
    If Cbo移入库房.Visible Then
        SQLCondition.lng移出库房 = Cbo移入库房.ItemData(Cbo移入库房.ListIndex)
    End If
    SQLCondition.str审核人 = Me.Txt审核人 & "%"
    SQLCondition.str填制人 = Me.Txt填制人 & "%"
    
    
    Unload Me
    
End Sub


Private Sub cmd药品_Click()
    Dim RecReturn As Recordset
    
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(1, "药品申领管理", mlng库房id, mlng库房id, mlng库房id, , , True)
    End If
    Set RecReturn = frmSelector.ShowMe(Me, 0, 1, , , , mlng库房id, mlng库房id, mlng库房id, , , , , 2, False)
    
'    Set RecReturn = Frm药品选择器.ShowME(Me, 1, 0, mlng库房id, mlng库房id)
    If RecReturn.RecordCount = 0 Then Exit Sub
    If gint药品名称显示 = 1 Then
        txt药品.Text = "[" & RecReturn!药品编码 & "]" & IIf(IsNull(RecReturn!商品名), RecReturn!通用名, RecReturn!商品名)
    Else
        txt药品.Text = "[" & RecReturn!药品编码 & "]" & RecReturn!通用名
    End If
    txt药品.Tag = RecReturn!药品ID
    
    If Chk移入库房.Visible = True Then
        Chk移入库房.SetFocus
    Else
        Txt填制人.SetFocus
    End If
End Sub

Private Sub Dtp结束时间_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If BlnAdvance Then
            Chk药品.SetFocus
        Else
            cmd确定.SetFocus
        End If
    End If
End Sub

Private Sub Dtp开始时间_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyReturn Then Me.Dtp结束时间(index).SetFocus
End Sub

Private Sub Form_Load()
    Dim intLop As Integer
    
    Me.Dtp结束时间(0) = zldatabase.Currentdate
    Me.Dtp结束时间(1) = Me.Dtp结束时间(0)
    Me.Dtp开始时间(0) = DateAdd("d", -7, Me.Dtp结束时间(0))
    Me.Dtp开始时间(1) = Me.Dtp开始时间(0)
    
    Me.txt药品.Tag = 0
    sstFilter.Tab = 0
    
    Select Case mlngMode
        Case 1304
            Chk移入库房.Caption = "移入库房"
        Case 1305
            Chk移入库房.Caption = "领用部门"
        Case 1306
            Chk移入库房.Caption = "入出类别"
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
    Call ReleaseSelectorRS
End Sub

Private Sub sstFilter_Click(PreviousTab As Integer)
    Dim rsDepartment As New Recordset
    Dim strStock As String
    
    On Error GoTo errHandle
    With sstFilter
        If .Tab = 1 Then
            BlnAdvance = True
            If Cbo移入库房.ListCount < 1 Then
                Select Case mlngMode
                    Case 1343
                        strStock = "HIJKLMN"
                        gstrSQL = "SELECT DISTINCT a.id, a.名称 " _
                             & "FROM 部门性质说明 c, 部门性质分类 b, 部门表 a " _
                            & "Where (a.站点 = '" & gstrNodeNo & "' Or a.站点 is Null) And c.工作性质 = b.名称 " _
                              & "AND Instr([1],b.编码,1) > 0 " _
                             & " AND a.id = c.部门id " _
                              & "AND a.撤档时间 = to_date('3000-01-01','yyyy-MM-dd')"
                    Case 1344
                        If Chk移入库房.Visible = True Then
                            Chk移入库房.Visible = False
                            Cbo移入库房.Visible = False
                            
                        End If
                        Exit Sub
                        
                End Select

                Set rsDepartment = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, strStock)
                
                With Cbo移入库房
                    Do While Not rsDepartment.EOF
                        .AddItem rsDepartment.Fields(1)
                        .ItemData(.NewIndex) = rsDepartment.Fields(0)
                        rsDepartment.MoveNext
                    Loop
                    .ListIndex = 0
                End With
                rsDepartment.Close
            End If
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Txt结束NO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng库房id As Long
    Dim intNO As Integer, strNo As String
    
    '初始准备
    intNO = Switch(mlngMode = 1343, 26, mlngMode = 1344, 23)
    lng库房id = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    If KeyCode = vbKeyReturn Then
        If Len(txt结束NO) < 8 And Len(txt结束NO) > 0 Then
            txt结束NO.Text = GetFullNO(txt结束NO.Text, intNO, lng库房id)
        End If
    End If
End Sub

Private Sub txt结束NO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub


Private Sub Txt开始NO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lng库房id As Long
    Dim intNO As Integer, strNo As String
    
    '初始准备
    intNO = Switch(mlngMode = 1343, 26, mlngMode = 1344, 23)
    lng库房id = mfrmMain.cboStock.ItemData(mfrmMain.cboStock.ListIndex)
    If KeyCode = vbKeyReturn Then
        If Len(txt开始NO) < 8 And Len(txt开始NO) > 0 Then
            txt开始NO.Text = GetFullNO(txt开始NO.Text, intNO, lng库房id)
        End If
        Me.txt结束NO.SetFocus
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
    
    Dim rstemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt审核人.Text) = "" Then
            cmd确定.SetFocus
            Exit Sub
        End If
        Txt审核人.Text = UCase(Txt审核人.Text)
        
        gstrSQL = "Select 编号,简码,姓名 From 人员表 Where (站点 = '" & gstrNodeNo & "' Or 站点 is Null) And (upper(姓名) like [1] or Upper(编号) like [1] or Upper(简码) like [2]) " & _
                " And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) "
        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[取审核人]", IIf(gstrMatchMethod = "0", "%", "") & Me.Txt审核人 & "%", Me.Txt审核人 & "%")
            
        With rstemp
            If .EOF Then
                MsgBox "输入值无效！", vbInformation, gstrSysName
                KeyCode = 0
                Txt审核人.SelStart = 0
                Txt审核人.SelLength = Len(Txt审核人.Text)
                
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Verify"
                Set mshSelect.Recordset = rstemp
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

Private Sub Txt审核人_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub


Private Sub Txt填制人_GotFocus()
    Txt填制人.SelStart = 0
    Txt填制人.SelLength = Len(Txt填制人.Text)
End Sub

Private Sub Txt填制人_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyReturn Then Me.Txt审核人.SetFocus
    Dim rstemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    If KeyCode = vbKeyReturn Then
        If Trim(Txt填制人.Text) = "" Then
            Txt审核人.SetFocus
            Exit Sub
        End If
        Txt填制人.Text = UCase(Txt填制人.Text)
        
        gstrSQL = "Select 编号,简码,姓名 From 人员表 Where (站点 = '" & gstrNodeNo & "' Or 站点 is Null) And (upper(姓名) like [1] or Upper(编号) like [1] or Upper(简码) like [2]) " & _
                " And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null) "
        Set rstemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[提取填制人]", IIf(gstrMatchMethod = "0", "%", "") & Me.Txt填制人 & "%", Me.Txt填制人 & "%")
        
        With rstemp
            If .EOF Then
                MsgBox "输入值无效！", vbInformation, gstrSysName
                KeyCode = 0
                Txt填制人.SelStart = 0
                Txt填制人.SelLength = Len(Txt填制人.Text)
                
                Exit Sub
            End If
            If .RecordCount > 1 Then
                mstrSelectTag = "Booker"
                Set mshSelect.Recordset = rstemp
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

Private Sub Txt填制人_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt药品_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim RecReturn As Recordset
    Dim strKey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txt药品.Text) = "" Then Exit Sub
    sngLeft = Me.Left + fra附加条件.Left + txt药品.Left
    sngTop = Me.Top + fra附加条件.Top + txt药品.Top + txt药品.Height + Me.Height - Me.ScaleHeight '  50
    If sngTop + 3630 > Screen.Height Then
        sngTop = sngTop - txt药品.Height - 3630
    End If
    
    strKey = Trim(txt药品.Text)
'    Set RecReturn = Frm药品多选选择器.ShowME(Me, 1, , mlng库房id, mlng库房id, strkey, sngLeft, sngTop)
    
    
    Call SetSelectorRS(1, "药品申领管理", mlng库房id, mlng库房id, mlng库房id, , , True)

    Set RecReturn = frmSelector.ShowMe(Me, 1, 1, strKey, sngLeft, sngTop, mlng库房id, mlng库房id, mlng库房id, , , , , 2, False)
    
    If RecReturn.RecordCount = 0 Then Exit Sub
    If gint药品名称显示 = 1 Then
        txt药品.Text = "[" & RecReturn!药品编码 & "]" & IIf(IsNull(RecReturn!商品名), RecReturn!通用名, RecReturn!商品名)
    Else
        txt药品.Text = "[" & RecReturn!药品编码 & "]" & RecReturn!通用名
    End If
    txt药品.Tag = RecReturn!药品ID
    
    If Chk移入库房.Visible = True Then
        Chk移入库房.SetFocus
    Else
        Txt填制人.SetFocus
    End If
    
End Sub

Private Sub Txt药品_KeyPress(KeyAscii As Integer)
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

