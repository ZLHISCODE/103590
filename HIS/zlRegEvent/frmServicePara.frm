VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServicePara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "挂号参数设置"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6900
   Icon            =   "frmServicePara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraBespeak 
      Caption         =   "预约挂号单"
      Height          =   705
      Left            =   135
      TabIndex        =   16
      Top             =   870
      Width           =   6675
      Begin VB.OptionButton optPrintBespeak 
         Caption         =   "选择是否打印"
         Height          =   180
         Index           =   2
         Left            =   3885
         TabIndex        =   6
         Top             =   300
         Width           =   1380
      End
      Begin VB.OptionButton optPrintBespeak 
         Caption         =   "不打印"
         Height          =   180
         Index           =   0
         Left            =   675
         TabIndex        =   4
         Top             =   315
         Width           =   900
      End
      Begin VB.OptionButton optPrintBespeak 
         Caption         =   "自动打印"
         Height          =   180
         Index           =   1
         Left            =   2130
         TabIndex        =   5
         Top             =   300
         Value           =   -1  'True
         Width           =   1020
      End
   End
   Begin VB.ComboBox cboDefaultStyle 
      Height          =   300
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   435
      Width           =   2160
   End
   Begin VB.TextBox txtAuto 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   1665
      TabIndex        =   2
      Text            =   "0"
      Top             =   150
      Width           =   435
   End
   Begin VB.CheckBox chkAuto 
      Caption         =   "自动刷新间隔:      秒"
      Height          =   300
      Left            =   195
      TabIndex        =   1
      Top             =   105
      Width           =   2700
   End
   Begin VB.CommandButton cmdDeviceSetup 
      Caption         =   "设备配置(&S)"
      Height          =   330
      Left            =   135
      TabIndex        =   14
      Top             =   4680
      Width           =   1425
   End
   Begin VB.Frame fraPrintSet 
      Caption         =   "打印设置"
      Height          =   720
      Left            =   135
      TabIndex        =   13
      Top             =   1755
      Width           =   6675
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "预约挂号单打印设置"
         Height          =   345
         Index           =   2
         Left            =   4575
         TabIndex        =   9
         Top             =   240
         Width           =   1890
      End
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "挂号凭条打印设置"
         Height          =   345
         Index           =   1
         Left            =   2385
         TabIndex        =   8
         Top             =   240
         Width           =   1890
      End
      Begin VB.CommandButton cmdPrintSet 
         Caption         =   "挂号票据打印设置"
         Height          =   345
         Index           =   0
         Left            =   195
         TabIndex        =   7
         Top             =   240
         Width           =   1890
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4335
      TabIndex        =   11
      Top             =   4665
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5520
      TabIndex        =   12
      Top             =   4665
      Width           =   1100
   End
   Begin VB.Frame fraTitle 
      Caption         =   "共用挂号票据"
      Height          =   1845
      Left            =   135
      TabIndex        =   0
      Top             =   2670
      Width           =   6675
      Begin MSComctlLib.ListView lvwBill 
         Height          =   1455
         Left            =   150
         TabIndex        =   10
         Top             =   240
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483630
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "领用人"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "领用日期"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "号码范围"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "剩余"
            Object.Width           =   1499
         EndProperty
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "缺省预约方式"
      Height          =   180
      Left            =   195
      TabIndex        =   15
      Top             =   495
      Width           =   1080
   End
End
Attribute VB_Name = "frmServicePara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub zlShowMe(ByVal frmMain As Object, ByVal lngModul As Long)
    Me.Show vbModal, frmMain
End Sub

Private Sub chkAuto_Click()
    If chkAuto.Value = 1 Then
        txtAuto.Enabled = True
    Else
        txtAuto.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, 100, 1111)
End Sub

Private Sub cmdOk_Click()
    Dim strTMP As String, i As Integer
    '共用挂号票据批次
    strTMP = "0"
    For i = 1 To lvwBill.ListItems.Count
        If lvwBill.ListItems(i).Checked Then strTMP = Mid(lvwBill.ListItems(i).Key, 2)
    Next
    zlDatabase.SetPara "共用挂号票据批次", strTMP, glngSys, 1111, True
    zlDatabase.SetPara "缺省预约方式", NeedName(cboDefaultStyle.Text), glngSys, 9000, True
    If chkAuto.Value = 1 Then
        zlDatabase.SetPara "刷新方式", "1" & "|" & Val(txtAuto.Text), glngSys, 1115, True
    Else
        zlDatabase.SetPara "刷新方式", 0, glngSys, 1115, True
    End If
    For i = 0 To Me.optPrintBespeak.UBound
        If optPrintBespeak(i).Value Then
            zlDatabase.SetPara "预约挂号单打印方式", i, glngSys, 9000, True
            Exit For
        End If
    Next
    Unload Me
End Sub

Private Sub cmdPrintSet_Click(index As Integer)
    Select Case index
    Case 0
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111", Me)
    Case 1
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1802", Me)
    Case 2
        Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1111_1", Me)
    End Select
End Sub

Private Sub Form_Load()
    Dim strTMP As String, i As Integer
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim intIndex As Integer
    strTMP = zlDatabase.GetPara("缺省预约方式", glngSys, 9000, "", Array(cboDefaultStyle), True)
    strSQL = "Select 编码,名称,缺省标志 From 预约方式 Order By 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    cboDefaultStyle.Clear
    Do While Not rsTmp.EOF
        cboDefaultStyle.AddItem rsTmp!编码 & "-" & rsTmp!名称
        If strTMP = Nvl(rsTmp!名称) Then intIndex = cboDefaultStyle.NewIndex
        If Val(Nvl(rsTmp!缺省标志)) = 1 Then cboDefaultStyle.ListIndex = cboDefaultStyle.NewIndex
        rsTmp.MoveNext
    Loop
    If cboDefaultStyle.ListCount <> 0 And intIndex <> 0 Then cboDefaultStyle.ListIndex = intIndex
    strTMP = zlDatabase.GetPara("刷新方式", glngSys, 1115, "0", Array(chkAuto, txtAuto), True) & "|"
    chkAuto.Value = IIf(Split(strTMP, "|")(0) = "1", 1, 0)
    If chkAuto.Value = 0 Then
        txtAuto.Text = "0"
        txtAuto.Enabled = False
    Else
        txtAuto.Text = Val(Split(strTMP, "|")(1))
        txtAuto.Enabled = True
    End If
    Call LoadFactList
    i = Val(zlDatabase.GetPara("预约挂号单打印方式", glngSys, 9000, 1, Array(optPrintBespeak(0), optPrintBespeak(1), optPrintBespeak(2)), True))
    If i <= optPrintBespeak.UBound Then optPrintBespeak(i).Value = True
End Sub

Private Function LoadFactList() As Boolean
'功能：读取可用公用挂号票据
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Integer, lngTmp As Long
    Dim objItem As ListItem
    Dim blnBill As Boolean
    
    On Error GoTo errH
    lngTmp = zlDatabase.GetPara("共用挂号票据批次", glngSys, 1111, 0, Array(lvwBill), True)
    Set rsTmp = GetShareInvoiceGroupID
    
    For i = 1 To rsTmp.RecordCount
        Set objItem = lvwBill.ListItems.Add(, "_" & rsTmp!ID, rsTmp!领用人)
        objItem.SubItems(1) = Format(rsTmp!登记时间, "yyyy-MM-dd")
        objItem.SubItems(2) = rsTmp!开始号码 & "," & rsTmp!终止号码
        objItem.SubItems(3) = rsTmp!剩余数量
        If rsTmp!ID = lngTmp Then
            objItem.Checked = True
            objItem.Selected = True
            blnBill = True
        End If
        rsTmp.MoveNext
    Next
    
    If Not blnBill Then
        zlDatabase.SetPara "共用挂号票据批次", "0", glngSys, 9000, True
    End If
    
    LoadFactList = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetShareInvoiceGroupID() As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定票种的共用票据批次
    '编制:刘兴洪
    '日期:2011-04-29 10:24:48
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    On Error GoTo errH
    
    strSQL = "" & _
    "   Select A.ID,A.使用类别,A.领用人,A.登记时间,A.开始号码,A.终止号码,A.剩余数量 " & _
    "   From 票据领用记录 A,人员表 B" & vbNewLine & _
    "   Where A.票种=4 And A.使用方式=2 And A.剩余数量>0 And A.领用人=B.姓名" & _
    "           And (B.站点='" & gstrNodeNo & "' Or B.站点 is Null)" & vbNewLine & _
    "   Order by 使用类别,剩余数量 Desc"
    
    Set GetShareInvoiceGroupID = zlDatabase.OpenSQLRecord(strSQL, App.ProductName)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
