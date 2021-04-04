VERSION 5.00
Begin VB.Form frmExpenseSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医嘱附费选项"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7500
   Icon            =   "frmExpenseSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   405
      Left            =   6180
      TabIndex        =   16
      Top             =   4155
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   405
      Left            =   5085
      TabIndex        =   15
      Top             =   4155
      Width           =   1100
   End
   Begin VB.Frame fraExpence 
      Height          =   3930
      Left            =   90
      TabIndex        =   17
      Top             =   45
      Width           =   7245
      Begin VB.OptionButton opt缺省科室 
         Caption         =   "病人科室"
         Height          =   195
         Index           =   1
         Left            =   5205
         TabIndex        =   32
         Top             =   3600
         Width           =   1065
      End
      Begin VB.OptionButton opt缺省科室 
         Caption         =   "医技科室"
         Height          =   195
         Index           =   0
         Left            =   4065
         TabIndex        =   31
         Top             =   3585
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.ComboBox cboSendMateria 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   3525
         Width           =   1860
      End
      Begin VB.ListBox lst收费类别 
         ForeColor       =   &H80000012&
         Height          =   3000
         Left            =   5700
         Style           =   1  'Checkbox
         TabIndex        =   14
         ToolTipText     =   "请复选允许使用的收费类别"
         Top             =   450
         Width           =   1440
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
      Begin VB.CheckBox chkTime 
         Caption         =   "变价允许输入数次"
         Height          =   195
         Left            =   195
         TabIndex        =   0
         Top             =   285
         Width           =   1740
      End
      Begin VB.CheckBox chk药库 
         Caption         =   "显示其它药库库存"
         Height          =   195
         Left            =   195
         TabIndex        =   3
         Top             =   1170
         Width           =   1770
      End
      Begin VB.CheckBox chk药房 
         Caption         =   "显示其它药房库存"
         Height          =   195
         Left            =   195
         TabIndex        =   2
         Top             =   885
         Width           =   1770
      End
      Begin VB.Frame Frame3 
         Caption         =   " 药品单位 "
         Height          =   1215
         Left            =   3120
         TabIndex        =   25
         Top             =   240
         Width           =   2445
         Begin VB.OptionButton opt药品单位 
            Caption         =   "门诊/住院单位"
            Height          =   180
            Index           =   1
            Left            =   360
            TabIndex        =   5
            Top             =   720
            Width           =   1470
         End
         Begin VB.OptionButton opt药品单位 
            Caption         =   "售价单位"
            Height          =   180
            Index           =   0
            Left            =   345
            TabIndex        =   4
            Top             =   330
            Value           =   -1  'True
            Width           =   1020
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   " 药房设置 "
         Height          =   1875
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   5445
         Begin VB.ComboBox cbo门发料部门 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   1290
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   1320
            Width           =   1350
         End
         Begin VB.ComboBox cbo住发料部门 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   3915
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1320
            Width           =   1350
         End
         Begin VB.ComboBox cbo门中药 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   1275
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   960
            Width           =   1350
         End
         Begin VB.ComboBox cbo门西药 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   1275
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   270
            Width           =   1350
         End
         Begin VB.ComboBox cbo门成药 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   1275
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   615
            Width           =   1350
         End
         Begin VB.ComboBox cbo住中药 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   3915
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   975
            Width           =   1350
         End
         Begin VB.ComboBox cbo住西药 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   3915
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   285
            Width           =   1350
         End
         Begin VB.ComboBox cbo住成药 
            ForeColor       =   &H80000012&
            Height          =   300
            Left            =   3915
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   630
            Width           =   1350
         End
         Begin VB.Label lbl门发料部门 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "门诊发料部门"
            Height          =   180
            Left            =   120
            TabIndex        =   28
            Top             =   1380
            Width           =   1080
         End
         Begin VB.Label lbl住发料部门 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "住院发料部门"
            Height          =   180
            Left            =   2745
            TabIndex        =   27
            Top             =   1380
            Width           =   1080
         End
         Begin VB.Label lbl门中药 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "门诊中药房"
            Height          =   180
            Left            =   285
            TabIndex        =   24
            Top             =   1020
            Width           =   900
         End
         Begin VB.Label lbl门西药 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "门诊西药房"
            Height          =   180
            Left            =   285
            TabIndex        =   23
            Top             =   330
            Width           =   900
         End
         Begin VB.Label lbl门成药 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "门诊成药房"
            Height          =   180
            Left            =   285
            TabIndex        =   22
            Top             =   675
            Width           =   900
         End
         Begin VB.Label lbl住中药 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "住院中药房"
            Height          =   180
            Left            =   2925
            TabIndex        =   21
            Top             =   1035
            Width           =   900
         End
         Begin VB.Label lbl住西药 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "住院西药房"
            Height          =   180
            Left            =   2925
            TabIndex        =   20
            Top             =   345
            Width           =   900
         End
         Begin VB.Label lbl住成药 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "住院成药房"
            Height          =   180
            Left            =   2925
            TabIndex        =   19
            Top             =   690
            Width           =   900
         End
      End
      Begin VB.Label lblFee 
         AutoSize        =   -1  'True
         Caption         =   "补费缺省科室"
         Height          =   180
         Left            =   2910
         TabIndex        =   30
         Top             =   3585
         Width           =   1080
      End
      Begin VB.Label lbl发药 
         Caption         =   "记帐之后"
         Height          =   255
         Left            =   105
         TabIndex        =   33
         Top             =   3555
         Width           =   735
      End
      Begin VB.Label lbl收费类别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "输入类别:"
         Height          =   180
         Left            =   5745
         TabIndex        =   26
         Top             =   225
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmExpenseSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明

Public mMainPrivs As String
Public mblnOK As Boolean
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub
Private Sub cmdOK_Click()
    Dim strPar As String, i As Long, blnSetup As Boolean
    
    blnSetup = InStr(GetInsidePrivs(p医嘱附费管理), "附费选项设置") > 0
    

    If cbo门西药.ListIndex = -1 Then
        MsgBox "请指定缺省的门诊西药房。", vbInformation, gstrSysName
        cbo门西药.SetFocus: Exit Sub
    End If
    If cbo门成药.ListIndex = -1 Then
        MsgBox "请指定缺省的门诊成药房。", vbInformation, gstrSysName
        cbo门成药.SetFocus: Exit Sub
    End If
    If cbo门中药.ListIndex = -1 Then
        MsgBox "请指定缺省的门诊中药房。", vbInformation, gstrSysName
        cbo门中药.SetFocus: Exit Sub
    End If
    If cbo住西药.ListIndex = -1 Then
        MsgBox "请指定缺省的住院西药房。", vbInformation, gstrSysName
        cbo住西药.SetFocus: Exit Sub
    End If
    If cbo住成药.ListIndex = -1 Then
        MsgBox "请指定缺省的住院成药房。", vbInformation, gstrSysName
        cbo住成药.SetFocus: Exit Sub
    End If
    If cbo住中药.ListIndex = -1 Then
        MsgBox "请指定缺省的住院中药房。", vbInformation, gstrSysName
        cbo住中药.SetFocus: Exit Sub
    End If
    
    '其它
    Call zlDatabase.SetPara("中药输入付数", chkPay.Value, glngSys, p医嘱附费管理, blnSetup)
    Call zlDatabase.SetPara("变价输入数次", chkTime.Value, glngSys, p医嘱附费管理, blnSetup)
    '问题:51762
    Call zlDatabase.SetPara("显示其它药库库存", chk药库.Value, glngSys, p医嘱附费管理, blnSetup)
    Call zlDatabase.SetPara("显示其它药房库存", chk药房.Value, glngSys, p医嘱附费管理, blnSetup)
    Call zlDatabase.SetPara("补费缺省科室", IIF(opt缺省科室(0).Value, 0, 1), glngSys, p医嘱附费管理, blnSetup)
        
    '药品单位
    Call zlDatabase.SetPara("药品单位", IIF(opt药品单位(0).Value, 0, 1), glngSys, p医嘱附费管理, blnSetup)
    '发药方式:25490
    Call zlDatabase.SetPara("记帐后发药", cboSendMateria.ListIndex, glngSys, p医嘱附费管理, blnSetup)
    '缺省药房
    Call zlDatabase.SetPara("门诊缺省西药房", cbo门西药.ItemData(cbo门西药.ListIndex), glngSys, p医嘱附费管理, blnSetup)
    Call zlDatabase.SetPara("门诊缺省成药房", cbo门成药.ItemData(cbo门成药.ListIndex), glngSys, p医嘱附费管理, blnSetup)
    Call zlDatabase.SetPara("门诊缺省中药房", cbo门中药.ItemData(cbo门中药.ListIndex), glngSys, p医嘱附费管理, blnSetup)
    Call zlDatabase.SetPara("门诊缺省发料部门", cbo门发料部门.ItemData(cbo门发料部门.ListIndex), glngSys, p医嘱附费管理, blnSetup)
    Call zlDatabase.SetPara("住院缺省西药房", cbo住西药.ItemData(cbo住西药.ListIndex), glngSys, p医嘱附费管理, blnSetup)
    Call zlDatabase.SetPara("住院缺省成药房", cbo住成药.ItemData(cbo住成药.ListIndex), glngSys, p医嘱附费管理, blnSetup)
    Call zlDatabase.SetPara("住院缺省中药房", cbo住中药.ItemData(cbo住中药.ListIndex), glngSys, p医嘱附费管理, blnSetup)
    Call zlDatabase.SetPara("住院缺省发料部门", cbo住发料部门.ItemData(cbo住发料部门.ListIndex), glngSys, p医嘱附费管理, blnSetup)
    
    '收费类别
    strPar = ""
    For i = lst收费类别.ListCount - 1 To 0 Step -1
        If lst收费类别.Selected(i) Then strPar = strPar & "'" & Chr(lst收费类别.ItemData(i)) & "',"
    Next
    If strPar <> "" Then strPar = Left(strPar, Len(strPar) - 1)
    Call zlDatabase.SetPara("收费类别", Replace(strPar, "'", "''"), glngSys, p医嘱附费管理, blnSetup)
    
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
    Dim objCbo As ComboBox, strPar As String
    Dim strSQL As String, i As Long
    Dim blnSetup As Boolean
    
    On Error GoTo errH
    mblnOK = False
    blnSetup = InStr(GetInsidePrivs(p医嘱附费管理), "附费选项设置") > 0
    
    chkPay.Value = Val(zlDatabase.GetPara("中药输入付数", glngSys, p医嘱附费管理, , Array(chkPay), blnSetup))
    chkTime.Value = Val(zlDatabase.GetPara("变价输入数次", glngSys, p医嘱附费管理, , Array(chkTime), blnSetup))
    chk药房.Value = Val(zlDatabase.GetPara("显示其它药房库存", glngSys, p医嘱附费管理, , Array(chk药房), blnSetup))
    chk药库.Value = Val(zlDatabase.GetPara("显示其它药库库存", glngSys, p医嘱附费管理, , Array(chk药库), blnSetup))
    '问题:36060
    If Val(zlDatabase.GetPara("补费缺省科室", glngSys, p医嘱附费管理, , Array(opt缺省科室(0), opt缺省科室(1), lblFee), blnSetup)) = 0 Then
        opt缺省科室(0).Value = True
    Else
        opt缺省科室(1).Value = True
    End If
      

    '药品单位
    i = Val(zlDatabase.GetPara("药品单位", glngSys, p医嘱附费管理, , Array(opt药品单位(0), opt药品单位(1)), blnSetup))
    opt药品单位(IIF(i = 0, 0, 1)).Value = True
    
    '25490
    cboSendMateria.Clear
    cboSendMateria.AddItem "不发药"
    cboSendMateria.AddItem "自动发药"
    cboSendMateria.AddItem "提示发药"
    i = Val(zlDatabase.GetPara("记帐后发药", glngSys, p医嘱附费管理, 0, Array(lbl发药, cboSendMateria), blnSetup))
    If i > cboSendMateria.ListCount Then i = 0
    cboSendMateria.ListIndex = i
    
    
    '缺省药房
    cbo门西药.AddItem "手工选择"
    cbo门成药.AddItem "手工选择"
    cbo门中药.AddItem "手工选择"
    cbo住西药.AddItem "手工选择"
    cbo住成药.AddItem "手工选择"
    cbo住中药.AddItem "手工选择"
    cbo门发料部门.AddItem "手工选择"
    cbo住发料部门.AddItem "手工选择"
    strSQL = _
        "Select Distinct A.ID,A.编码,A.名称,B.工作性质,B.服务对象" & _
        " From 部门表 A,部门性质说明 B " & _
        " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And B.部门ID=A.ID And B.服务对象 IN(1,2,3)" & _
        " And B.工作性质 in('西药房','成药房','中药房','发料部门')" & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " Order by A.编码"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        If rsTmp!工作性质 = "西药房" Then
            Set objCbo = IIF(rsTmp!服务对象 = 1, cbo门西药, IIF(rsTmp!服务对象 = 2, cbo住西药, Nothing))
        End If
        If rsTmp!工作性质 = "成药房" Then
            Set objCbo = IIF(rsTmp!服务对象 = 1, cbo门成药, IIF(rsTmp!服务对象 = 2, cbo住成药, Nothing))
        End If
        If rsTmp!工作性质 = "中药房" Then
            Set objCbo = IIF(rsTmp!服务对象 = 1, cbo门中药, IIF(rsTmp!服务对象 = 2, cbo住中药, Nothing))
        End If
        If rsTmp!工作性质 = "发料部门" Then
            Set objCbo = IIF(rsTmp!服务对象 = 1, cbo门发料部门, IIF(rsTmp!服务对象 = 2, cbo住发料部门, Nothing))
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
            ElseIf rsTmp!工作性质 = "发料部门" Then
                cbo门发料部门.AddItem rsTmp!名称
                cbo门发料部门.ItemData(cbo门发料部门.NewIndex) = rsTmp!ID
                cbo住发料部门.AddItem rsTmp!名称
                cbo住发料部门.ItemData(cbo住发料部门.NewIndex) = rsTmp!ID
            End If
        Else
            objCbo.AddItem rsTmp!名称
            objCbo.ItemData(objCbo.NewIndex) = rsTmp!ID
        End If
        rsTmp.MoveNext
    Next
    strPar = zlDatabase.GetPara("门诊缺省西药房", glngSys, p医嘱附费管理, , Array(lbl门西药, cbo门西药), blnSetup)
    For i = 0 To cbo门西药.ListCount - 1
        If cbo门西药.ItemData(i) = Val(strPar) Then cbo门西药.ListIndex = i: Exit For
    Next
    strPar = zlDatabase.GetPara("门诊缺省成药房", glngSys, p医嘱附费管理, , Array(lbl门成药, cbo门成药), blnSetup)
    For i = 0 To cbo门成药.ListCount - 1
        If cbo门成药.ItemData(i) = Val(strPar) Then cbo门成药.ListIndex = i: Exit For
    Next
    strPar = zlDatabase.GetPara("门诊缺省中药房", glngSys, p医嘱附费管理, , Array(lbl门中药, cbo门中药), blnSetup)
    For i = 0 To cbo门中药.ListCount - 1
        If cbo门中药.ItemData(i) = Val(strPar) Then cbo门中药.ListIndex = i: Exit For
    Next
    strPar = zlDatabase.GetPara("门诊缺省发料部门", glngSys, p医嘱附费管理, , Array(lbl门发料部门, cbo门发料部门), blnSetup)
    For i = 0 To cbo门发料部门.ListCount - 1
        If cbo门发料部门.ItemData(i) = Val(strPar) Then cbo门发料部门.ListIndex = i: Exit For
    Next
    
    
    strPar = zlDatabase.GetPara("住院缺省西药房", glngSys, p医嘱附费管理, , Array(lbl住西药, cbo住西药), blnSetup)
    For i = 0 To cbo住西药.ListCount - 1
        If cbo住西药.ItemData(i) = Val(strPar) Then cbo住西药.ListIndex = i: Exit For
    Next
    strPar = zlDatabase.GetPara("住院缺省成药房", glngSys, p医嘱附费管理, , Array(lbl住成药, cbo住成药), blnSetup)
    For i = 0 To cbo住成药.ListCount - 1
        If cbo住成药.ItemData(i) = Val(strPar) Then cbo住成药.ListIndex = i: Exit For
    Next
    strPar = zlDatabase.GetPara("住院缺省中药房", glngSys, p医嘱附费管理, , Array(lbl住中药, cbo住中药), blnSetup)
    For i = 0 To cbo住中药.ListCount - 1
        If cbo住中药.ItemData(i) = Val(strPar) Then cbo住中药.ListIndex = i: Exit For
    Next
    strPar = zlDatabase.GetPara("住院缺省发料部门", glngSys, p医嘱附费管理, , Array(lbl住发料部门, cbo住发料部门), blnSetup)
    For i = 0 To cbo住发料部门.ListCount - 1
        If cbo住发料部门.ItemData(i) = Val(strPar) Then cbo住发料部门.ListIndex = i: Exit For
    Next
    
    
    '收费类别
    strSQL = "Select 编码,名称 as 类别 From 收费项目类别 Where 编码<>'1' Order by 序号"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        lst收费类别.AddItem rsTmp!类别
        lst收费类别.ItemData(lst收费类别.NewIndex) = Asc(rsTmp!编码)
        rsTmp.MoveNext
    Loop
    strPar = zlDatabase.GetPara("收费类别", glngSys, p医嘱附费管理, "", Array(lbl收费类别, lst收费类别), blnSetup)
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
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mMainPrivs = ""
End Sub

Private Sub lst收费类别_ItemCheck(Item As Integer)
    If lst收费类别.SelCount = 0 And Not lst收费类别.Selected(Item) Then
        lst收费类别.Selected(Item) = True
    End If
End Sub
