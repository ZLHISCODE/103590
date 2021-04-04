VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
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
      TabIndex        =   32
      Top             =   4245
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
      TabIndex        =   33
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
      Tab(0).Control(1)=   "chkDept(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkDept(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "chkDept(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkDept(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "chkDept(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "附加条件(&D)"
      TabPicture(1)   =   "frmDrugPaymentSearch.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra附加条件"
      Tab(1).ControlCount=   1
      Begin VB.CheckBox chkDept 
         Caption         =   "卫材(&W)"
         Height          =   195
         Index           =   4
         Left            =   4830
         TabIndex        =   13
         Tag             =   "4"
         Top             =   3615
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkDept 
         Caption         =   "设备(&S)"
         Height          =   195
         Index           =   2
         Left            =   2550
         TabIndex        =   11
         Tag             =   "4"
         Top             =   3615
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkDept 
         Caption         =   "物资(&M)"
         Height          =   195
         Index           =   1
         Left            =   1380
         TabIndex        =   10
         Tag             =   "2"
         Top             =   3615
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkDept 
         Caption         =   "药品(&D)"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   9
         Tag             =   "1"
         Top             =   3615
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkDept 
         Caption         =   "其他(&Q)"
         Height          =   195
         Index           =   3
         Left            =   3675
         TabIndex        =   12
         Tag             =   "4"
         Top             =   3615
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.Frame fra附加条件 
         Height          =   3225
         Left            =   -74760
         TabIndex        =   41
         Top             =   510
         Width           =   5505
         Begin VB.ComboBox cboStore 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1245
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   750
            Width           =   3970
         End
         Begin VB.CheckBox chkStore 
            Caption         =   "库房"
            Height          =   300
            Left            =   330
            TabIndex        =   17
            Top             =   750
            Width           =   800
         End
         Begin VB.CommandButton Cmd供应商 
            Caption         =   "…"
            Enabled         =   0   'False
            Height          =   270
            Left            =   4935
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   390
            Width           =   255
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   0
            Left            =   1245
            TabIndex        =   20
            Tag             =   "品名"
            Top             =   1125
            Width           =   3945
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   1
            Left            =   1245
            TabIndex        =   22
            Tag             =   "开始发票号"
            Top             =   1500
            Width           =   3945
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   2
            Left            =   1245
            TabIndex        =   24
            Tag             =   "结束发票号"
            Top             =   1875
            Width           =   3945
         End
         Begin VB.CheckBox Chk供应商 
            Caption         =   "供应商"
            Height          =   300
            Left            =   330
            TabIndex        =   14
            Top             =   375
            Width           =   870
         End
         Begin VB.TextBox txt供应商 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1245
            MaxLength       =   50
            TabIndex        =   15
            Top             =   375
            Width           =   3945
         End
         Begin VB.TextBox Txt填制人 
            Height          =   300
            Left            =   1245
            MaxLength       =   8
            TabIndex        =   26
            Top             =   2250
            Width           =   3945
         End
         Begin VB.TextBox Txt审核人 
            Height          =   300
            Left            =   1245
            MaxLength       =   8
            TabIndex        =   28
            Top             =   2625
            Width           =   3945
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "品名"
            Height          =   180
            Index           =   2
            Left            =   840
            TabIndex        =   19
            Top             =   1185
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "开始发票号"
            Height          =   180
            Index           =   7
            Left            =   300
            TabIndex        =   21
            Top             =   1560
            Width           =   900
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "结束发票号"
            Height          =   180
            Index           =   5
            Left            =   300
            TabIndex        =   23
            Top             =   1935
            Width           =   900
         End
         Begin VB.Label Lbl填制人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "填制人"
            Height          =   180
            Left            =   660
            TabIndex        =   25
            Top             =   2310
            Width           =   540
         End
         Begin VB.Label Lbl审核人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "审核人"
            Height          =   180
            Left            =   660
            TabIndex        =   27
            Top             =   2685
            Width           =   540
         End
      End
      Begin VB.Frame fra范围 
         Height          =   2850
         Left            =   240
         TabIndex        =   34
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
            Format          =   269221891
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
            Format          =   269221891
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
            Format          =   269221891
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
            Format          =   80281603
            CurrentDate     =   36263
         End
         Begin VB.Label LblNO 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "No"
            Height          =   180
            Left            =   480
            TabIndex        =   40
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
            TabIndex        =   39
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
            TabIndex        =   38
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
            TabIndex        =   37
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
            TabIndex        =   36
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
            TabIndex        =   35
            Top             =   1140
            Width           =   180
         End
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   6330
      TabIndex        =   31
      Top             =   3720
      Width           =   1100
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6330
      TabIndex        =   30
      Top             =   930
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6330
      TabIndex        =   29
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
Private mblnAdvance As Boolean '是否展开
Private mdatStart As Date   '开始时间
Private mdatEnd As Date     '结束时间
Private mdatVerifyStart As Date
Private mdatVerifyEnd As Date
Private mstrSelectTag As String     '当前选择的对象
Private mstrPrivs As String
Private mstr类型 As String
Private mstrOthers(0 To 9) As String '0-记录状态,1-开始单号,2-结束单号,3-供应商ID,4-审核人,5-填制人,6-开始发票号,7-结束发票号,8-品名,9-库房ID

Public Function GetSearch(ByVal FrmMain As Form, ByVal strPrivs As String, _
        ByRef datStart As Date, ByRef datEnd As Date, _
        ByRef datVerifyStart As Date, ByRef datVerifyEnd As Date, ByRef str类型 As String, ByRef strOthers() As String) As String
'--------------------------------------------------------------
'功能：获取检索所需的SQL语句
'参数：FrmMain----------调用窗体
'      datStart---------起始日期
'      datEnd-----------结束日期
'      datVerifyStart---审核起始日期
'      datVerifyEnd-----审核结束日期
'      strOthers-附加条件设置(0-记录状态,1-开始单号,2-结束单号,3-供应商ID,4-审核人,5-填制人,6-开始发票号,7-结束发票号,8-品名)
'返回：SQL语句
'说明：
'--------------------------------------------------------------
    mstrFind = "": mstrPrivs = strPrivs
    
    If Not CheckCompete Then Exit Function
    Call 权限设置
    Me.Show vbModal, FrmMain
    GetSearch = mstrFind
    datStart = mdatStart
    datEnd = mdatEnd
    datVerifyStart = mdatVerifyStart
    datVerifyEnd = mdatVerifyEnd
    str类型 = mstr类型
    strOthers = mstrOthers
End Function

Private Sub chkDept_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlcommfun.PressKey (vbKeyTab)
    End If
End Sub


Private Sub chkStore_Click()
    cboStore.Enabled = chkStore.Value = 1
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


Private Sub cmdHelp_Click()
       ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub Cmd供应商_Click()
    Dim strTemp As String
    
    strTemp = frm供应商选择.SelDept(mstrPrivs)
    If strTemp = "" Then Exit Sub
    txt供应商.Text = Mid(strTemp, InStr(strTemp, ",") + 1)
    txt供应商.Tag = Left(strTemp, InStr(strTemp, ",") - 1)
    Unload frm供应商选择
End Sub

Private Sub Cmd取消_Click()
    mstrFind = ""
    Unload Me
End Sub

Private Sub Cmd确定_Click()
    Dim strFind As String, strKey As String
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
    'by lesfeng 2009-12-2 性能优化
    'mstrOthers(0 To 8) '0-记录状态,1-开始单号,2-结束单号,3-供应商ID,4-审核人,5-填制人,6-开始发票号,7-结束发票号,8-品名
    '参数范围: 1-开始填制日期,2-结束填制日期
    '          3-开始审核日期,4-结束审核日期
    '          5-开始单号,6-结束单号,7-供应商ID,8-审核人,9-填制人,10-开始发票号,11-结束发票号,12-品名,13-库房ID
    
    If chk填制.Value = 1 And chk审核.Value = 1 Then
        If chkStrike.Value = 1 Then
           mstrFind = " And ((A.填制日期 Between [1] And [2]) or (A.审核日期 Between [3] And [4]))"
        Else
           mstrFind = " And ((A.填制日期 Between [1] And [2]) or (A.审核日期 Between [3] And [4])) And A.记录状态 =1"
        End If
        mdatStart = Format(dtp开始时间(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp结束时间(0), "yyyy-mm-dd")
                
        mdatVerifyStart = Format(dtp开始时间(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp结束时间(1), "yyyy-mm-dd")
        
    ElseIf chk审核.Value = 1 Then
        If chkStrike.Value = 1 Then
            mstrFind = " And A.审核日期 Between [3] And [4]"
        Else
            mstrFind = " And A.审核日期 Between [3] And [4] And A.记录状态 =1"
        End If
        mdatVerifyStart = Format(dtp开始时间(1), "yyyy-mm-dd")
        mdatVerifyEnd = Format(dtp结束时间(1), "yyyy-mm-dd")
        mdatStart = Format("1901 - 01 - 01", "yyyy-mm-dd")
        mdatEnd = Format("1901-01-01", "yyyy-mm-dd")
    ElseIf chk填制.Value = 1 Then
        mstrFind = " And (A.填制日期 Between [1] And [2]) And A.审核日期 is null "
        mdatStart = Format(dtp开始时间(0), "yyyy-mm-dd")
        mdatEnd = Format(dtp结束时间(0), "yyyy-mm-dd")
        mdatVerifyStart = Format("1901-01-01", "yyyy-mm-dd")
        mdatVerifyEnd = Format("1901-01-01", "yyyy-mm-dd")
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
    
    mstrOthers(1) = Me.txt开始No
    mstrOthers(2) = Me.txt结束NO
    
    If Me.txt开始No <> "" And Me.txt结束NO <> "" Then mstrFind = mstrFind & " And A.No >= [5] And A.No <= [6] "
    If Me.txt开始No <> "" And Me.txt结束NO = "" Then mstrFind = mstrFind & " And A.No >= [5] "
    If Me.txt开始No = "" And Me.txt结束NO <> "" Then mstrFind = mstrFind & " And A.No <= [6] "
    
    '扩展查询条件
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
    
    If mblnAdvance = False Then
        Unload Me
        Exit Sub
    End If
    
    mstrOthers(3) = Val(txt供应商.Tag)
    If Chk供应商.Value = 1 Then
        mstrFind = mstrFind & " and a.单位id= [7] "
'        mstrOthers(3) = Val(txt供应商.Tag)
    End If
    
    If Me.Txt审核人 <> "" Then
        mstrFind = mstrFind & " And A.审核人 like [8] "
        mstrOthers(4) = IIf(gstrMatchMethod = "0", "%", "") & Me.Txt审核人 & "%"
    End If
    If Me.Txt填制人 <> "" Then
        mstrFind = mstrFind & " And A.填制人 like [9] "
        mstrOthers(5) = IIf(gstrMatchMethod = "0", "%", "") & Me.Txt填制人 & "%"
    End If
    
    strTemp = ""
    
    If Me.txtEdit(1).Text <> "" And Me.txtEdit(2) <> "" Then
        strTemp = strTemp & " And 发票号 >= [10] And 发票号 <= [11] "
        mstrOthers(6) = Me.txtEdit(1)
        mstrOthers(7) = Me.txtEdit(2)
    End If
    If Me.txtEdit(1) <> "" And Me.txtEdit(2) = "" Then
        strTemp = strTemp & " And 发票号 >= [10] "
        mstrOthers(6) = Me.txtEdit(1)
    End If
    If Me.txtEdit(1) = "" And Me.txtEdit(2) <> "" Then
        strTemp = strTemp & " And 发票号 <= [11] "
        mstrOthers(7) = Me.txtEdit(2)
    End If
    If Trim(Me.txtEdit(0)) <> "" Then
        strKey = GetMatchingSting(Trim(Me.txtEdit(0)), True)
        strFind = " And 品名 like [12]"
        mstrOthers(8) = GetMatchingSting(Me.txtEdit(0).Text, False)
        If zlcommfun.IsCharAlpha(Trim(txtEdit(0).Text)) Then          '01,11.输入全是字母时只匹配简码
            '0-拼音码,1-五笔码,2-两者
            If gSystemPara.int简码方式 = 1 Then
                '五笔码查询
                If Mid(gSystemPara.Para_输入方式, 2, 1) = "1" Then strFind = " And zltools.zlWBCode(品名) Like upper([12]) "
            ElseIf gSystemPara.int简码方式 = 0 Then
                If Mid(gSystemPara.Para_输入方式, 2, 1) = "1" Then strFind = " And zltools.zlspellcode(品名) Like upper([12]) "
            Else
                If Mid(gSystemPara.Para_输入方式, 2, 1) = "1" Then strFind = " And (zltools.zlWBCode(品名) Like upper([12]) or zltools.zlspellcode(品名) Like upper([12])"
            End If
            
        End If
        strTemp = strTemp & strFind
    End If
    
    If chkStore.Value = 1 Then
        strTemp = strTemp & " And 库房ID = [13] "
        If cboStore.ListIndex = -1 Then
            mstrOthers(9) = ""
        Else
            mstrOthers(9) = cboStore.ItemData(cboStore.ListIndex)
        End If
    End If
    
    If strTemp <> "" Then
        mstrFind = mstrFind & " And  Exists (Select 1 From 应付记录 Where a.付款序号=付款序号 " & strTemp & ") "
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
    Me.dtp结束时间(0) = zlDatabase.Currentdate
    Me.dtp结束时间(1) = Me.dtp结束时间(0)
    
    Me.dtp开始时间(0) = DateAdd("d", -15, Me.dtp结束时间(0))
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
    gstrSQL = "" & _
        "   Select id,上级ID,编码,简码,末级,名称 " & _
        "   From 供应商 " & _
        "   Where 名称 is Not NULL " & zl_获取站点限制 & " " & _
        "       And (撤档时间 is null or 撤档时间>= to_date('3000-01-01','yyyy-mm-dd')) " & _
        "   Start with 上级ID is NULL Connect by prior id=上级id"
    On Error GoTo errHandle
    zlDatabase.OpenRecordset rsCompete, gstrSQL, "过滤"
    
    With rsCompete
        If .EOF Then
            .Close
            MsgBox "供应商信息不全，请在供应商管理中设置供应商信息！", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    
    CheckCompete = True
    Exit Function
    
errHandle:
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

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlcommfun.PressKey vbKeyTab
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txtEdit(0), KeyAscii, m文本式)
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
    Dim vRect As RECT
    Err = 0: On Error GoTo ErrHand:
    If strKey <> "" Then
        strKey = GetMatchingSting(strKey)
    End If
      
    str权限 = " and " & Get分类权限(mstrPrivs)
    gstrSQL = "" & _
        "   Select id, 编码,名称,末级,简码,许可证号,许可证效期,执照号,执照效期,税务登记号,地址,开户银行,帐号,联系人,建档时间,类型,信用期" & _
        "   From 供应商 " & _
        "   where 末级=1 " & zl_获取站点限制 & " " & _
        "          and  (撤档时间 is null or 撤档时间>=to_date('3000-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss')) " & _
        "          and (编码 like [1] or 名称 like [1] or 简码 like [1])  " & str权限
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
    Dim lngH As Long
    vRect = zlControl.GetControlRect(txt供应商.hwnd)
 
    lngH = txt供应商.Height
    
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "供应商选择", False, "", "选择供应商", False, True, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, True, strKey)
    If blnCancel Then Exit Function
    If rsTemp Is Nothing Then
        ShowMsgbox "不存在符何条件的供应商,请检查!"
        Exit Function
    End If
    If rsTemp.State <> 1 Then Exit Function
    txt供应商 = Nvl(rsTemp!编码) & "-" & Nvl(rsTemp!名称)
    txt供应商.Tag = Nvl(rsTemp!ID)
    Select供应商 = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub txt供应商_Change()
    txt供应商.Tag = ""
End Sub

Private Sub txt供应商_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As New Recordset
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If txt供应商.Tag <> "" Then
        zlcommfun.PressKey vbKeyTab
        Exit Sub
    End If
    If Select供应商(txt供应商.Text) = False Then
        If txt供应商.Enabled And txt供应商.Visible Then txt供应商.SetFocus
        Exit Sub
    End If
    zlcommfun.PressKey vbKeyTab
End Sub

Private Sub txt供应商_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub
'问题29757 by lesfeng 2010-05-10
Private Sub txt供应商_Validate(Cancel As Boolean)
    Dim rsTemp As New Recordset
    
    If txt供应商.Tag <> "" Then Exit Sub
    
    If Select供应商(txt供应商.Text) = False Then
        Exit Sub
    End If

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

Private Sub Txt审核人_Change()
    Txt审核人.Tag = ""
End Sub

Private Sub Txt审核人_GotFocus()
    Txt审核人.SelStart = 0
    Txt审核人.SelLength = Len(Txt审核人.Text)
End Sub

Private Sub Txt审核人_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Txt审核人.Text <> "" And Txt审核人.Tag = "" Then
        Dim lng人员ID As Long
        
        If MulitSelectPersion(Me, Txt审核人, Trim(Txt审核人.Text), 0, lng人员ID) = False Then
            If Txt审核人.Enabled Then Txt审核人.SetFocus
            Exit Sub
        End If
        Txt审核人.Tag = lng人员ID
    End If
    zlcommfun.PressKey vbKeyTab
End Sub

Private Sub Txt审核人_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Txt填制人_Change()
    Txt填制人.Tag = ""
End Sub

Private Sub Txt填制人_GotFocus()
    Txt填制人.SelStart = 0
    Txt填制人.SelLength = Len(Txt填制人.Text)
End Sub

Private Sub Txt填制人_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Txt填制人.Text <> "" And Txt填制人.Tag = "" Then
        Dim lng人员ID As Long
        
        If MulitSelectPersion(Me, Txt填制人, Trim(Txt填制人.Text), 0, lng人员ID) = False Then
            If Txt填制人.Enabled Then Txt填制人.SetFocus
            Exit Sub
        End If
        Txt填制人.Tag = lng人员ID
    End If
    zlcommfun.PressKey vbKeyTab

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Public Sub 权限设置()
    
    If Check相关权限(mstrPrivs, "药品") = False Then
        chkDept(0).Enabled = False
        chkDept(0).Value = 0
    End If
    If Check相关权限(mstrPrivs, "物资") = False Then
        chkDept(1).Enabled = False
        chkDept(1).Value = 0
    End If
    
    If Check相关权限(mstrPrivs, "设备") = False Then
        chkDept(2).Enabled = False
        chkDept(2).Value = 0
    End If
    If Check相关权限(mstrPrivs, "其他") = False Then
        chkDept(3).Enabled = False
        chkDept(3).Value = 0
    End If
    If Check相关权限(mstrPrivs, "卫材") = False Then
        chkDept(4).Enabled = False
        chkDept(4).Value = 0
    End If
End Sub

