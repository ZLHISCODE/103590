VERSION 5.00
Begin VB.Form frmTechnicSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   Icon            =   "frmTechnicSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Caption         =   "缩略图显示设置"
      Height          =   690
      Left            =   120
      TabIndex        =   37
      Top             =   6450
      Width           =   5475
      Begin VB.TextBox TxtShowPhotoNumber 
         Height          =   315
         Left            =   1740
         TabIndex        =   20
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "最大显示缩略图数"
         Height          =   180
         Left            =   120
         TabIndex        =   19
         Top             =   300
         Width           =   1440
      End
   End
   Begin VB.Frame fraAction 
      Caption         =   " 执行设置 "
      Height          =   1875
      Left            =   120
      TabIndex        =   24
      Top             =   4560
      Width           =   5460
      Begin VB.CheckBox chkEmergencyPrint 
         Caption         =   "紧急审核打印"
         Height          =   255
         Left            =   2940
         TabIndex        =   41
         Top             =   1560
         Width           =   1695
      End
      Begin VB.CheckBox chkIgnorePosi 
         Caption         =   "忽略结果的阴阳性"
         Height          =   180
         Left            =   2940
         TabIndex        =   40
         Top             =   1330
         Width           =   2235
      End
      Begin VB.CheckBox chkBatchInput 
         Caption         =   "连续输入申请"
         Height          =   225
         Left            =   2940
         TabIndex        =   39
         Top             =   390
         Width           =   2475
      End
      Begin VB.CheckBox chkSample 
         Caption         =   "申请登记后直接检查"
         Height          =   225
         Left            =   2940
         TabIndex        =   38
         Top             =   135
         Width           =   2475
      End
      Begin VB.CheckBox chkView 
         Caption         =   "填写报告时打开观片站"
         Height          =   180
         Left            =   2940
         TabIndex        =   18
         Top             =   1115
         Width           =   2235
      End
      Begin VB.CheckBox chkFinish 
         Caption         =   "允许未收费病人完成执行"
         Height          =   195
         Left            =   2940
         TabIndex        =   17
         Top             =   885
         Width           =   2280
      End
      Begin VB.CheckBox chkActLog 
         Caption         =   "允许其他人代行执行记录"
         Height          =   195
         Left            =   2940
         TabIndex        =   16
         Top             =   645
         Width           =   2280
      End
      Begin VB.ListBox lstRoom 
         Enabled         =   0   'False
         Height          =   690
         ItemData        =   "frmTechnicSetup.frx":000C
         Left            =   255
         List            =   "frmTechnicSetup.frx":000E
         Style           =   1  'Checkbox
         TabIndex        =   15
         Top             =   555
         Width           =   2535
      End
      Begin VB.CheckBox chkRoom 
         Caption         =   "指定病人的执行间范围"
         Height          =   195
         Left            =   255
         TabIndex        =   14
         Top             =   285
         Width           =   2100
      End
   End
   Begin VB.Frame fraExpence 
      Caption         =   " 计费设置 "
      Height          =   4470
      Left            =   120
      TabIndex        =   25
      Top             =   45
      Width           =   5460
      Begin VB.Frame fraLine 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   780
         TabIndex        =   35
         Top             =   4320
         Width           =   465
      End
      Begin VB.TextBox txtRefresh 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   780
         MaxLength       =   4
         TabIndex        =   12
         Text            =   "0"
         Top             =   4140
         Width           =   465
      End
      Begin VB.Frame Frame2 
         Caption         =   " 药房设置 "
         Height          =   2505
         Left            =   195
         TabIndex        =   28
         Top             =   1560
         Width           =   3525
         Begin VB.ComboBox cbo住成药 
            Height          =   300
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   1710
            Width           =   2190
         End
         Begin VB.ComboBox cbo住西药 
            Height          =   300
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1365
            Width           =   2190
         End
         Begin VB.ComboBox cbo住中药 
            Height          =   300
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   2055
            Width           =   2190
         End
         Begin VB.ComboBox cbo门成药 
            Height          =   300
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   615
            Width           =   2190
         End
         Begin VB.ComboBox cbo门西药 
            Height          =   300
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   270
            Width           =   2190
         End
         Begin VB.ComboBox cbo门中药 
            Height          =   300
            Left            =   1155
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   960
            Width           =   2190
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "住院成药房"
            Height          =   180
            Left            =   165
            TabIndex        =   34
            Top             =   1770
            Width           =   900
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "住院西药房"
            Height          =   180
            Left            =   165
            TabIndex        =   33
            Top             =   1425
            Width           =   900
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "住院中药房"
            Height          =   180
            Left            =   165
            TabIndex        =   32
            Top             =   2115
            Width           =   900
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "门诊成药房"
            Height          =   180
            Left            =   165
            TabIndex        =   31
            Top             =   675
            Width           =   900
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "门诊西药房"
            Height          =   180
            Left            =   165
            TabIndex        =   30
            Top             =   330
            Width           =   900
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "门诊中药房"
            Height          =   180
            Left            =   165
            TabIndex        =   29
            Top             =   1020
            Width           =   900
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " 药品单位 "
         Height          =   615
         Left            =   195
         TabIndex        =   27
         Top             =   870
         Width           =   3525
         Begin VB.OptionButton opt药品单位 
            Caption         =   "售价单位"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   0
            Left            =   465
            TabIndex        =   4
            Top             =   285
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton opt药品单位 
            Caption         =   "门诊/住院单位"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   1
            Left            =   1560
            TabIndex        =   5
            Top             =   285
            Width           =   1470
         End
      End
      Begin VB.CheckBox chk药房 
         Caption         =   "显示其它药房库存"
         Height          =   195
         Left            =   1995
         TabIndex        =   2
         Top             =   285
         Width           =   1770
      End
      Begin VB.CheckBox chk药库 
         Caption         =   "显示其它药库库存"
         Height          =   195
         Left            =   1995
         TabIndex        =   3
         Top             =   570
         Width           =   1770
      End
      Begin VB.CheckBox chkTime 
         Caption         =   "变价允许输入数次"
         Height          =   195
         Left            =   195
         TabIndex        =   0
         Top             =   285
         Width           =   1740
      End
      Begin VB.CheckBox chkPay 
         Caption         =   "中药可以输入付数"
         Height          =   195
         Left            =   195
         TabIndex        =   1
         Top             =   570
         Value           =   1  'Checked
         Width           =   1740
      End
      Begin VB.ListBox lst收费类别 
         Height          =   3420
         Left            =   3930
         Style           =   1  'Checkbox
         TabIndex        =   13
         ToolTipText     =   "请复选允许使用的收费类别"
         Top             =   645
         Width           =   1350
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "每隔      秒自动刷新病人清单"
         Height          =   180
         Left            =   390
         TabIndex        =   36
         Top             =   4155
         Width           =   2520
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "输入类别:"
         Height          =   180
         Left            =   3945
         TabIndex        =   26
         Top             =   420
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3060
      TabIndex        =   21
      Top             =   7215
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4155
      TabIndex        =   22
      Top             =   7215
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   420
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   7215
      Width           =   1100
   End
End
Attribute VB_Name = "frmTechnicSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Public mlng科室ID As Long 'IN:当前执行科室ID
Public mblnOK As Boolean

Private Sub chkRoom_Click()
    lstRoom.Enabled = chkRoom.Value = 1
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim strPar As String, i As Long
    
    '执行间范围
    strPar = ""
    If chkRoom.Value = 1 Then
        For i = 0 To lstRoom.ListCount - 1
            If lstRoom.Selected(i) Then
                strPar = strPar & "|" & lstRoom.List(i)
            End If
        Next
        If strPar = "" Then
            MsgBox "请至少选择一个执行间。", vbInformation, gstrSysName
            lstRoom.SetFocus: Exit Sub
        End If
    End If
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\科室" & mlng科室ID, "执行间范围", Mid(strPar, 2)
        
    '其它
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "中药付数", chkPay.Value
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "变价数次", chkTime.Value
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "显示其它药房库存", chk药房.Value
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "显示其它药库库存", chk药库.Value
    
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "医技刷新间隔", Val(txtRefresh.Text)
    
    '药品单位
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "药品单位", IIf(opt药品单位(0).Value, 0, 1)
    
    '缺省药房
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "门诊缺省西药房", cbo门西药.ItemData(cbo门西药.ListIndex)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "门诊缺省成药房", cbo门成药.ItemData(cbo门成药.ListIndex)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "门诊缺省中药房", cbo门中药.ItemData(cbo门中药.ListIndex)
    
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "住院缺省西药房", cbo住西药.ItemData(cbo住西药.ListIndex)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "住院缺省成药房", cbo住成药.ItemData(cbo住成药.ListIndex)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "住院缺省中药房", cbo住中药.ItemData(cbo住中药.ListIndex)
    
    '收费类别
    strPar = ""
    For i = lst收费类别.ListCount - 1 To 0 Step -1
        If lst收费类别.Selected(i) Then strPar = strPar & "'" & Chr(lst收费类别.ItemData(i)) & "',"
    Next
    If strPar <> "" Then strPar = Left(strPar, Len(strPar) - 1)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "收费类别", strPar
    
    '是否允许代行执行记录
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "代行执行记录", chkActLog.Value

    '是否允许完成未收费病人的项目
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "未收费完成", chkFinish.Value
    
    '显示图像数
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "显示图像数", CLng(Val(TxtShowPhotoNumber.Text))
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "报告时观片", chkView.Value
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "连续登记申请", chkBatchInput.Value
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "登记直接检查", chkSample.Value
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "忽略结果阴阳性", chkIgnorePosi.Value
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "紧急审核时打印", chkEmergencyPrint.Value
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Call cmdHelp_Click
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim objCbo As ComboBox, lng药房ID As Long
    Dim strSQL As String, strPar As String, i As Long
    
    mblnOK = False
    
    chkPay.Value = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "中药付数", 1))
    chkTime.Value = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "变价数次", 0))
    chk药房.Value = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "显示其它药房库存", 0))
    chk药库.Value = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "显示其它药库库存", 0))
    
    txtRefresh.Text = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "医技刷新间隔", 0))
        
    '药品单位
    i = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "药品单位", 0))
    opt药品单位(IIf(i = 0, 0, 1)).Value = True
    
    '缺省药房
    cbo门西药.AddItem "手工选择": cbo门西药.ListIndex = 0
    cbo门成药.AddItem "手工选择": cbo门成药.ListIndex = 0
    cbo门中药.AddItem "手工选择": cbo门中药.ListIndex = 0
    cbo住西药.AddItem "手工选择": cbo住西药.ListIndex = 0
    cbo住成药.AddItem "手工选择": cbo住成药.ListIndex = 0
    cbo住中药.AddItem "手工选择": cbo住中药.ListIndex = 0
    strSQL = _
        "Select Distinct A.ID,A.编码,A.名称,B.工作性质,B.服务对象" & _
        " From 部门表 A,部门性质说明 B " & _
        " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And B.部门ID=A.ID And B.服务对象 IN(1,2,3)" & _
        " And B.工作性质 in('西药房','成药房','中药房')" & _
        " Order by A.编码"
    Call OpenRecord(rsTmp, strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        If rsTmp!工作性质 = "西药房" Then
            Set objCbo = IIf(rsTmp!服务对象 = 1, cbo门西药, IIf(rsTmp!服务对象 = 2, cbo住西药, Nothing))
        End If
        If rsTmp!工作性质 = "成药房" Then
            Set objCbo = IIf(rsTmp!服务对象 = 1, cbo门成药, IIf(rsTmp!服务对象 = 2, cbo住成药, Nothing))
        End If
        If rsTmp!工作性质 = "中药房" Then
            Set objCbo = IIf(rsTmp!服务对象 = 1, cbo门中药, IIf(rsTmp!服务对象 = 2, cbo住中药, Nothing))
        End If
        If objCbo Is Nothing Then
            If rsTmp!工作性质 = "西药房" Then
                cbo门西药.AddItem rsTmp!名称
                cbo门西药.ItemData(cbo门西药.NewIndex) = rsTmp!ID
                cbo住西药.AddItem rsTmp!名称
                cbo住西药.ItemData(cbo住西药.NewIndex) = rsTmp!ID
            ElseIf rsTmp!工作性质 = "成药房" Then
                cbo门成药.AddItem rsTmp!名称
                cbo门成药.ItemData(cbo门成药.NewIndex) = rsTmp!ID
                cbo住成药.AddItem rsTmp!名称
                cbo住成药.ItemData(cbo住成药.NewIndex) = rsTmp!ID
            ElseIf rsTmp!工作性质 = "中药房" Then
                cbo门中药.AddItem rsTmp!名称
                cbo门中药.ItemData(cbo门中药.NewIndex) = rsTmp!ID
                cbo住中药.AddItem rsTmp!名称
                cbo住中药.ItemData(cbo住中药.NewIndex) = rsTmp!ID
            End If
        Else
            objCbo.AddItem rsTmp!名称
            objCbo.ItemData(objCbo.NewIndex) = rsTmp!ID
        End If
        rsTmp.MoveNext
    Next
    lng药房ID = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "门诊缺省西药房", 0))
    Call FindCboIndex(cbo门西药, lng药房ID, True)
    lng药房ID = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "门诊缺省成药房", 0))
    Call FindCboIndex(cbo门成药, lng药房ID, True)
    lng药房ID = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "门诊缺省中药房", 0))
    Call FindCboIndex(cbo门中药, lng药房ID, True)
    lng药房ID = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "住院缺省西药房", 0))
    Call FindCboIndex(cbo住西药, lng药房ID, True)
    lng药房ID = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "住院缺省成药房", 0))
    Call FindCboIndex(cbo住成药, lng药房ID, True)
    lng药房ID = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "住院缺省中药房", 0))
    Call FindCboIndex(cbo住中药, lng药房ID, True)
    
    '收费类别
    strSQL = "Select 编码,名称 as 类别 From 收费项目类别 Where 编码<>'1' Order by 序号"
    Call OpenRecord(rsTmp, strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        lst收费类别.AddItem rsTmp!类别
        lst收费类别.ItemData(lst收费类别.NewIndex) = Asc(rsTmp!编码)
        rsTmp.MoveNext
    Loop
    strPar = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "收费类别", "")
    If strPar = "" Then
        For i = 0 To lst收费类别.ListCount - 1
            lst收费类别.Selected(i) = True
        Next
    Else
        For i = 0 To lst收费类别.ListCount - 1
            If InStr(strPar, Chr(lst收费类别.ItemData(i))) Then lst收费类别.Selected(i) = True
        Next
    End If
    If lst收费类别.ListCount > 0 Then lst收费类别.TopIndex = 0: lst收费类别.ListIndex = 0
    
    '是否允许代行执行记录
    chkActLog.Value = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "代行执行记录", 0))
    
    '是否允许完成未收费病人的项目
    chkFinish.Value = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "未收费完成", 0))
        
    '执行房间
    strPar = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\科室" & mlng科室ID, "执行间范围", "")
    chkRoom.Value = IIf(strPar = "", 0, 1)
    strSQL = "Select 执行间 From 医技执行房间 Where 科室ID=[1]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, mlng科室ID)
    Do While Not rsTmp.EOF
        lstRoom.AddItem rsTmp!执行间
        If InStr("|" & strPar & "|", "|" & rsTmp!执行间 & "|") > 0 Then
            lstRoom.Selected(lstRoom.NewIndex) = True
        End If
        rsTmp.MoveNext
    Loop
    If lstRoom.ListCount > 0 Then
        lstRoom.TopIndex = 0
        lstRoom.ListIndex = 0
    Else
        chkRoom.Value = 0
        chkRoom.Enabled = False
    End If
    
    '显示图像数
    TxtShowPhotoNumber = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "显示图像数", 20))
    
    chkView.Value = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "报告时观片", 0))
    chkBatchInput.Value = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "连续登记申请", 0))
    chkSample.Value = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "登记直接检查", 0))
    chkIgnorePosi.Value = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "忽略结果阴阳性", 0))
    chkEmergencyPrint.Value = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "紧急审核时打印", 0))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlng科室ID = 0
End Sub

Private Sub lst收费类别_ItemCheck(Item As Integer)
    If lst收费类别.SelCount = 0 And Not lst收费类别.Selected(Item) Then
        lst收费类别.Selected(Item) = True
    End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    
End Sub

Private Sub txtRefresh_GotFocus()
    Call zlControl.TxtSelAll(txtRefresh)
End Sub

Private Sub txtRefresh_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub TxtShowPhotoNumber_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
