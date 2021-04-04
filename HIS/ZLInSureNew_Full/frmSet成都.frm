VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BILLEDIT.OCX"
Begin VB.Form frmSet成都 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医疗保险接口配置"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   ControlBox      =   0   'False
   Icon            =   "frmSet成都.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin ZL9BillEdit.BillEdit Bill 
      Height          =   1365
      Left            =   900
      TabIndex        =   14
      Top             =   2100
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2408
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.CheckBox chkhisCharge 
      Caption         =   "HIS收费"
      Height          =   200
      Left            =   2640
      TabIndex        =   13
      Top             =   2050
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox txtInterCode 
      Height          =   300
      Left            =   4500
      MaxLength       =   6
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "713"
      Top             =   1965
      Visible         =   0   'False
      Width           =   960
   End
   Begin MSComCtl2.UpDown UDCard 
      Height          =   315
      Left            =   2190
      TabIndex        =   2
      Top             =   1965
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   556
      _Version        =   393216
      Value           =   30
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtCard"
      BuddyDispid     =   196613
      OrigLeft        =   2415
      OrigTop         =   1965
      OrigRight       =   2655
      OrigBottom      =   2280
      Max             =   30
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtCard 
      Alignment       =   2  'Center
      Height          =   300
      Left            =   1710
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "30"
      Top             =   1980
      Width           =   480
   End
   Begin VB.CommandButton cmdODBC 
      Caption         =   "数据源(&D)"
      Height          =   350
      Left            =   225
      TabIndex        =   6
      Top             =   2520
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   2640
      TabIndex        =   4
      Top             =   2520
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3840
      TabIndex        =   5
      Top             =   2520
      Width           =   1100
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "测试(&T)"
      Height          =   350
      Left            =   1425
      TabIndex        =   3
      Top             =   2520
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   -210
      TabIndex        =   9
      Top             =   2340
      Width           =   5850
   End
   Begin VB.TextBox txt连接串 
      Height          =   720
      Left            =   915
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1095
      Width           =   4575
   End
   Begin VB.Label Lbl收费类别对照 
      AutoSize        =   -1  'True
      Caption         =   "收费类别对照"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   930
      TabIndex        =   15
      Top             =   1890
      Width           =   1080
   End
   Begin VB.Label lblInterCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "医保内码"
      Height          =   180
      Left            =   3720
      TabIndex        =   12
      Top             =   2040
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lblCard 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "卡号长度"
      Height          =   180
      Left            =   930
      TabIndex        =   10
      Top             =   2040
      Width           =   720
   End
   Begin VB.Label lbl连接串 
      AutoSize        =   -1  'True
      Caption         =   "连接串"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   930
      TabIndex        =   8
      Top             =   885
      Width           =   540
   End
   Begin VB.Label lblNote 
      Caption         =   "    设置到医疗保险数据服务器的连接串；为保证设置有效，这时医疗保险数据服务器必须可用。"
      Height          =   390
      Left            =   930
      TabIndex        =   7
      Top             =   225
      Width           =   4500
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   210
      Picture         =   "frmSet成都.frx":030A
      Top             =   180
      Width           =   240
   End
End
Attribute VB_Name = "frmSet成都"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mint险类 As Integer
Dim mblnOK As Boolean
Dim str费用项目 As String   '收费类别对应医保的费用项目

Private Sub Bill_cboClick(ListIndex As Long)
    Bill.TextMatrix(Bill.Row, Bill.COL) = Bill.CboText
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdODBC_Click()
    On Error Resume Next
    Shell "ODBCAD32", vbNormalFocus
    If Err.Number <> 0 Then
        MsgBox "不能进入ODBC数据源管理器，请检查系统是否正确安装！", vbInformation, gstrSysName
    End If
    Err.Clear
End Sub

Private Sub cmdOK_Click()
    Select Case mint险类
        Case TYPE_成都市
            SaveSetting "ZLSOFT", "公共模块\zl9Insure", UCase("ConnectionStrINg"), Trim(txt连接串.Text)
            SaveSetting "ZLSOFT", "公共模块\zl9Insure", UCase("CardNOLength"), txtCard.Text
        Case TYPE_成都南充
            If Not CheckItem Then
                If MsgBox("有部分收费类别未设置对应的医保收费项目，继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            End If
            Call Combinate
            SaveSetting "ZLSOFT", "公共模块\zl9Insure", UCase("LCConnectionString"), Trim(txt连接串.Text)
            SaveSetting "ZLSOFT", "公共模块\zl9Insure", UCase("LCItem"), str费用项目
        Case TYPE_成都莲合
            SaveSetting "ZLSOFT", "公共模块\zl9Insure", UCase("CardNOLength"), txtCard.Text
            SaveSetting "ZLSOFT", "公共模块\zl9Insure", UCase("LHConnectionStrINg"), Trim(txt连接串.Text)
            SaveSetting "ZLSOFT", "公共模块\zl9Insure", UCase("intercode"), txtInterCode.Text
            SaveSetting "ZLSOFT", "公共模块\zl9Insure", UCase("HIS收费"), chkhisCharge.Value
        Case TYPE_开县
            '20050124
            SaveSetting "ZLSOFT", "公共模块\zl9Insure", UCase("LHConnectionStrINg"), Trim(txt连接串.Text)
            SaveSetting "ZLSOFT", "公共模块\zl9Insure", UCase("intercode"), txtInterCode.Text
            SaveSetting "ZLSOFT", "公共模块\zl9Insure", UCase("HIS收费"), chkhisCharge.Value
    End Select
    mblnOK = True
    Unload Me
End Sub

Private Sub cmdTest_Click()
    Dim cnInsure As New ADODB.Connection
    Err = 0
    On Error Resume Next
    With cnInsure
        If .State = adStateOpen Then .Close
        .ConnectionString = Trim(Me.txt连接串.Text)
        .Open
        If Err <> 0 Then
            MsgBox "测试不成功，请检查医保数据服务器是否可用，以及数据源是否正确配置！", vbExclamation, gstrSysName
            Exit Sub
        End If
        .Close
        If txtInterCode.Visible = True Then
            If txtInterCode.Text = "" Then
                MsgBox "医保内码不能为空，请重输！", vbExclamation, gstrSysName
                Exit Sub
            End If
            If IsNumeric(txtInterCode.Text) = False Then
                MsgBox "医保内码必须为数字型，请重输！", vbExclamation, gstrSysName
                txtInterCode.SelStart = 0
                txtInterCode.SelLength = Len(txtInterCode.Text)
                txtInterCode.SetFocus
                Exit Sub
            End If
        End If
        
        MsgBox "测试成功，与医保数据服务器正常连接！", vbInformation, gstrSysName
        Me.cmdOK.Enabled = True
    End With
End Sub

Private Sub txtCard_Change()
    If txtCard.Locked Then Exit Sub
    Me.cmdOK.Enabled = True
End Sub

Private Sub txtInterCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        zlCommFun.PressKey (vbKeyTab)
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0: Beep: Exit Sub
    End If
End Sub

Private Sub txt连接串_KeyPress(KeyAscii As Integer)
    Me.cmdOK.Enabled = False
End Sub

Public Function ShowSet(ByVal int险类 As Integer) As Boolean
'功能：得到参数配置信息
    Dim rsTemp As New ADODB.Recordset
    mblnOK = False
    mint险类 = int险类
    
    Lbl收费类别对照.Visible = False
    Bill.Visible = False
    If int险类 <> TYPE_成都南充 Then
        Frame1.Top = txtCard.Top + txtCard.Height + 100
        Me.Height = 3380
    Else
        Frame1.Top = Bill.Top + Bill.Height + 100
        Me.Height = 4560
    End If
    Call AdjustCons
    
    Select Case int险类
        Case TYPE_成都市
            txt连接串.Text = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("ConnectionStrINg"), "dsn=cnnSyb;uID=face;pwd=facepass")
            txtCard.Text = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("CardNOLength"), 20)
        Case TYPE_成都南充
            Lbl收费类别对照.Visible = True
            Bill.Visible = True
            txtCard.Visible = False
            lblCard.Visible = False
            UDCard.Visible = False
            lblInterCode.Visible = False
            txtInterCode.Visible = False
            chkhisCharge.Visible = False
            txt连接串.Text = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("LCConnectionStrINg"), "dsn=lcyb;uid=hisuser;pwd=hiscdgk;")
            str费用项目 = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("LCItem"), "")
            
            '初始化表格
            Call InitBill
            
            '装入费用项目
            gstrSQL = "Select 类别 From 收费类别 Order By 编码"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取收费类别")
            
            '装入设定值
            Call LoadSet(rsTemp)
        Case TYPE_成都莲合
            lblCard.Caption = "主键长度"
            UDCard.Visible = False
            lblInterCode.Visible = True
            txtInterCode.Visible = True
            chkhisCharge.Visible = True
            txtCard.Text = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("CardNOLength"), 10)
            txt连接串.Text = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("LHConnectionStrINg"), "dsn=lhyb;uid=sa;pwd=;")
            txtInterCode.Text = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("intercode"), 713)
            chkhisCharge.Value = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("HIS收费"), 0)
            txtCard.Locked = False
        Case TYPE_开县
            txtCard.Visible = False
            lblCard.Visible = False
            UDCard.Visible = False
            lblInterCode.Visible = True
            txtInterCode.Visible = True
            chkhisCharge.Visible = True
            txt连接串.Text = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("LHConnectionStrINg"), "dsn=lhyb;uid=sa;pwd=;")
            txtInterCode.Text = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("intercode"), 713)
            chkhisCharge.Value = GetSetting("ZLSOFT", "公共模块\zl9Insure", UCase("HIS收费"), 0)
    End Select
    frmSet成都.Show vbModal
    
    ShowSet = mblnOK
End Function

Private Sub AdjustCons()
    With cmdODBC
        .Top = Frame1.Top + 200
    End With
    cmdCancel.Top = cmdODBC.Top
    cmdOK.Top = cmdODBC.Top
    cmdTest.Top = cmdODBC.Top
End Sub

Private Sub InitBill()
    With Bill
        .AllowAddRow = False
        .Active = True
        .ClearBill
        .Cols = 2
        
        .TextMatrix(0, 0) = "收费类别"
        .TextMatrix(0, 1) = "医保费用项目"
        
        .ColData(0) = 0
        .ColData(1) = 3
        
        .ColWidth(0) = 1200
        .ColWidth(1) = 2500
        
        .PrimaryCol = 1
        .LocateCol = 1
    End With
End Sub

Private Sub LoadSet(ByVal rsTemp As ADODB.Recordset)
    Dim arrItem, intItem As Integer, strItem As String
    
    Bill.Rows = rsTemp.RecordCount + 1
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    
    '装入收费类别
    For intItem = 1 To rsTemp.RecordCount
        Bill.TextMatrix(intItem, 0) = rsTemp!类别
        rsTemp.MoveNext
    Next
    
    '装入医保费用项目
    arrItem = Split(gstr费用项目, gstrSplit大类)
    For intItem = 0 To UBound(arrItem)
        Bill.AddItem arrItem(intItem)
    Next
    
    '装入设定正确的费用项目
    arrItem = Split(str费用项目, gstrSplit大类)
    For intItem = 0 To UBound(arrItem)
        strItem = Split(arrItem(intItem), gstrSplit小类)(1)
        '检查所设定的医保费用项目是否是正确的
        If InStr(1, gstrSplit大类 & gstr费用项目 & gstrSplit大类, gstrSplit大类 & strItem & gstrSplit大类) <> 0 Then
            rsTemp.MoveFirst
            rsTemp.Find "类别='" & Split(arrItem(intItem), gstrSplit小类)(0) & "'"  '找到其对应的收费类别
            If Not rsTemp.EOF Then Bill.TextMatrix(rsTemp.AbsolutePosition, 1) = strItem
        End If
    Next
End Sub

Private Sub Combinate()
    Dim intItem  As Integer
    '将设定的内容组合成串
    str费用项目 = ""
    For intItem = 1 To Bill.Rows - 1
        str费用项目 = str费用项目 & gstrSplit大类 & Bill.TextMatrix(intItem, 0) & gstrSplit小类 & Bill.TextMatrix(intItem, 1)
    Next
    str费用项目 = Mid(str费用项目, 2)
End Sub

Private Function CheckItem() As Boolean
    Dim intRow As Integer
    CheckItem = False
    For intRow = 1 To Bill.Rows - 1
        If Trim(Bill.TextMatrix(intRow, 1)) = "" Then Exit Function
    Next
    CheckItem = True
End Function
