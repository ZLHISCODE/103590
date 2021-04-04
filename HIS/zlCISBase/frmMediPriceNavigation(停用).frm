VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMediPriceNavigation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "调价选项"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4950
   Icon            =   "frmMediPriceNavigation.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCanc 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3480
      Picture         =   "frmMediPriceNavigation.frx":000C
      TabIndex        =   11
      Top             =   3600
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2400
      Picture         =   "frmMediPriceNavigation.frx":0156
      TabIndex        =   10
      Top             =   3600
      Width           =   1100
   End
   Begin VB.Frame fra辅助选项 
      Caption         =   "辅助选项（成本价调价相关）"
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   4695
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshProvider 
         Height          =   1695
         Left            =   120
         TabIndex        =   16
         Top             =   2280
         Visible         =   0   'False
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   2990
         _Version        =   393216
         FixedCols       =   0
         GridColor       =   32768
         FocusRect       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CheckBox chk加成率 
         Caption         =   "指定加成率"
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   1125
         Width           =   1215
      End
      Begin VB.CheckBox chk供应商 
         Caption         =   "指定供应商"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.CheckBox chk应付记录 
         Caption         =   "产生成本价调价带来的应付款修正记录"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   3495
      End
      Begin VB.TextBox txt加成率 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   270
         Left            =   1440
         TabIndex        =   7
         Text            =   "15.0000"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txt供应商 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1440
         TabIndex        =   6
         Top             =   360
         Width           =   2655
      End
      Begin VB.CommandButton cmd供应商 
         Caption         =   "…"
         Enabled         =   0   'False
         Height          =   270
         Left            =   4080
         TabIndex        =   5
         Top             =   350
         Width           =   375
      End
      Begin VB.Label lblComment加成率 
         Caption         =   "（指定加成率，则统一默认按该加成率计算成本价；不指定，则默认显示实际加成率）"
         ForeColor       =   &H00FF0000&
         Height          =   540
         Left            =   240
         TabIndex        =   15
         Top             =   1440
         Width           =   4260
      End
      Begin VB.Label lblComment供应商 
         AutoSize        =   -1  'True
         Caption         =   "（指定供应商，则只调整该供应商的库存药品成本价）"
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   4320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   180
         Left            =   2415
         TabIndex        =   8
         Top             =   1125
         Width           =   90
      End
   End
   Begin VB.Frame fra内容 
      Caption         =   "调价内容"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.OptionButton opt调价 
         Caption         =   "调售价及成本价"
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton opt调价 
         Caption         =   "仅调成本价"
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton opt调价 
         Caption         =   "调售价"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmMediPriceNavigation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mblnSelect As Boolean
Private mint调价 As Integer             '0-调售价;1-调成本价;2-调售价及成本价
Private mlng供应商ID As Long
Private mdbl加成率 As Double
Private mbln应付记录 As Boolean         '0-不产生应付记录;1-产生应付记录
Private mstrPrivs As String
Private Sub SetForm(ByVal int调价 As Integer)
    If int调价 = 0 Then
        fra辅助选项.Visible = False
        cmdOk.Top = fra内容.Top + fra内容.Height + 200
        cmdCanc.Top = cmdOk.Top
    Else
        fra辅助选项.Visible = True
        cmdOk.Top = fra辅助选项.Top + fra辅助选项.Height + 200
        cmdCanc.Top = cmdOk.Top
    End If
    Me.Height = cmdOk.Top + cmdOk.Height + 800
    
    If InStr(1, mstrPrivs, "售价管理") = 0 Then
        opt调价(0).Visible = False
        opt调价(2).Visible = False
        opt调价(1).Left = opt调价(0).Left
    End If
End Sub

Private Sub chk供应商_Click()
    If chk供应商.Value = 1 Then
        txt供应商.Enabled = True
        cmd供应商.Enabled = True
        chk应付记录.Enabled = True
    Else
        txt供应商.Enabled = False
        cmd供应商.Enabled = False
        chk应付记录.Value = 0
        chk应付记录.Enabled = False
    End If
End Sub

Private Sub chk加成率_Click()
    If chk加成率.Value = 1 Then
        txt加成率.Enabled = True
        If Val(Trim(txt加成率.Text)) = 0 Then
            txt加成率.Text = "15.0000"
        End If
    Else
        txt加成率.Enabled = False
    End If
End Sub

Private Sub cmdCanc_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If opt调价(0).Value Then
        mint调价 = 0
    ElseIf opt调价(1).Value Then
        mint调价 = 1
    Else
        mint调价 = 2
    End If
    
    If fra辅助选项.Visible Then
        If chk供应商.Value = 1 Then
            If Val(Split(txt供应商.Tag, "|")(0)) = 0 Then
                MsgBox "请选择供应商。", vbInformation, gstrSysName
                txt供应商.SetFocus
                Exit Sub
            End If
        End If
                
        mlng供应商ID = IIf(chk供应商.Value = 1, Val(Split(txt供应商.Tag, "|")(0)), 0)
        mdbl加成率 = IIf(chk加成率.Value = 1, Val(Trim(txt加成率.Text)), 0)
        mbln应付记录 = (chk应付记录.Enabled And chk应付记录.Value = 1)
    End If
    
    mblnSelect = True
    Unload Me
End Sub

Public Function GetCondition(frmMain As Form, ByVal strPrivs As String, ByRef int调价 As Integer, ByRef lng供应商ID As Long, ByRef dbl加成率 As Double, ByRef bln应付记录 As Boolean) As Boolean
    mblnSelect = False
    mstrPrivs = strPrivs
    Set mfrmMain = frmMain
    Me.Show vbModal, frmMain
    GetCondition = mblnSelect
    
    If mblnSelect = False Then Exit Function
    
    int调价 = mint调价
    lng供应商ID = mlng供应商ID
    dbl加成率 = mdbl加成率
    bln应付记录 = mbln应付记录
End Function


Private Sub cmd供应商_Click()
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSql = "Select 编码,名称,简码,id" & _
        " From 供应商" & _
        " where 末级=1 And substr(类型,1,1) = '1' And (撤档时间 is null or 撤档时间=to_date('3000-01-01','YYYY-MM-DD')) " & _
        " Order By 编码 "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "取供应商信息")
    If rsTemp.EOF Then
        MsgBox "请初始化供应商（字典管理）！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With Me.mshProvider
        .Left = chk供应商.Left
        .Top = txt供应商.Top + txt供应商.Height
        .Clear
        Set .DataSource = rsTemp
        .ColWidth(0) = 800: .ColWidth(1) = 2500: .ColWidth(2) = 800: .ColWidth(3) = 0
        .Row = 1: .ColSel = .Cols - 1
        .ZOrder 0: .Visible = True: .SetFocus
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    Me.txt供应商.Tag = "|"
    Call SetForm(0)
End Sub

Private Sub mshProvider_DblClick()
    With Me.mshProvider
        Me.txt供应商.Text = .TextMatrix(.Row, 1)
        Me.txt供应商.Tag = .TextMatrix(.Row, 3) & "|" & .TextMatrix(.Row, 1)
        .Visible = False
    End With
    Me.txt供应商.SetFocus
End Sub


Private Sub mshProvider_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call mshProvider_DblClick
End Sub


Private Sub mshProvider_LostFocus()
    Me.mshProvider.Visible = False
End Sub


Private Sub opt调价_Click(Index As Integer)
    SetForm (Index)
End Sub


Private Sub txt供应商_GotFocus()
    Me.txt供应商.SelStart = 0: Me.txt供应商.SelLength = Len(Me.txt供应商.Text)
End Sub


Private Sub txt供应商_KeyPress(KeyAscii As Integer)
    Dim strTmp As String
    Dim rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
        
    strTmp = UCase(Trim(Me.txt供应商.Text))
    
    If strTmp = "" Then
        Me.txt供应商.Tag = "|"
        Exit Sub
    ElseIf strTmp = Split(Me.txt供应商.Tag, "|")(1) Then
        Exit Sub
    End If
    
    gstrSql = "Select 编码,名称,简码,id" & _
            " From 供应商" & _
            " where (编码 Like [1] " & _
            "       Or 名称 Like [2] " & _
            "       Or 简码 Like [2])" & _
            " And 末级=1 And substr(类型,1,1) = '1' And (撤档时间 is null or 撤档时间=to_date('3000-01-01','YYYY-MM-DD')) " & _
            " Order By 编码 "
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, strTmp & "%", gstrMatch & strTmp & "%")
    
    With rsTemp
        If .EOF Then
            MsgBox "没有找到匹配的供应商，请在供应商管理中增加供应商！", vbInformation, gstrSysName
            Me.txt供应商.Text = Split(Me.txt供应商.Tag, "|")(1)
            Me.txt供应商.SelStart = 0: Me.txt供应商.SelLength = Len(Me.txt供应商.Text)
            Exit Sub
        End If
        
        If .RecordCount = 1 Then
            Me.txt供应商.Text = Trim(rsTemp!名称): Me.txt供应商.Tag = rsTemp!ID & "|" & rsTemp!名称
            Exit Sub
        Else
            With Me.mshProvider
                .Left = Me.chk供应商.Left
                .Top = Me.txt供应商.Top + Me.txt供应商.Height
                .Clear
                Set .DataSource = rsTemp
                .ColWidth(0) = 800: .ColWidth(1) = 2500: .ColWidth(2) = 800: .ColWidth(3) = 0
                .Row = 1: .ColSel = .Cols - 1
                .ZOrder 0: .Visible = True: .SetFocus
            End With
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub txt供应商_Validate(Cancel As Boolean)
    If Me.txt供应商.Text = "" Then
        Me.txt供应商.Tag = "|"
    ElseIf Me.txt供应商.Text <> Split(Me.txt供应商.Tag, "|")(1) Then
        txt供应商_KeyPress (vbKeyReturn)
    End If
End Sub


Private Sub txt加成率_GotFocus()
    txt加成率.SelStart = 0
    txt加成率.SelLength = Len(txt加成率)
End Sub

Private Sub txt加成率_KeyPress(KeyAscii As Integer)
    If Not (Chr(KeyAscii) >= 0 And Chr(KeyAscii) <= 9 Or KeyAscii = vbKeyBack Or KeyAscii = 46) Then KeyAscii = 0
End Sub


