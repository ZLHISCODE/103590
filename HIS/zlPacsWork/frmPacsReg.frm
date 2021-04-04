VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPACSReg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "开始影像检查"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   Icon            =   "frmPacsReg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkUnicode 
      Caption         =   "同一患者的检查号在本科室统一编号"
      Height          =   210
      Left            =   120
      TabIndex        =   24
      Top             =   4080
      Width           =   4080
   End
   Begin VB.Frame fraMatch 
      Caption         =   "对于提前进行的检查，按下列项目匹配检查图像"
      Height          =   645
      Left            =   -30
      TabIndex        =   43
      Top             =   4395
      Width           =   6135
      Begin VB.OptionButton optMatch 
         Caption         =   "检查号"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   25
         ToolTipText     =   "按检查号将病人和接收的影像进行匹配"
         Top             =   300
         Width           =   1395
      End
      Begin VB.OptionButton optMatch 
         Caption         =   "门诊/住院号"
         Height          =   195
         Index           =   1
         Left            =   2550
         TabIndex        =   26
         ToolTipText     =   "按病人标识号将病人和接收的影像进行匹配"
         Top             =   300
         Width           =   1395
      End
      Begin VB.OptionButton optMatch 
         Caption         =   "检查标识号"
         Height          =   195
         Index           =   2
         Left            =   4620
         TabIndex        =   27
         ToolTipText     =   "按检查标识号将病人和接收的影像进行匹配"
         Top             =   300
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      Height          =   945
      Left            =   -100
      TabIndex        =   32
      Top             =   720
      Width           =   6250
      Begin VB.TextBox txtDept 
         Height          =   300
         Left            =   3780
         Locked          =   -1  'True
         TabIndex        =   41
         Top             =   510
         Width           =   2085
      End
      Begin VB.TextBox txtPatID 
         Height          =   300
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   510
         Width           =   1455
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   0
         Left            =   1275
         MaxLength       =   10
         TabIndex        =   1
         Top             =   120
         Width           =   4590
      End
      Begin VB.Label Label7 
         Caption         =   "申请科室"
         Height          =   255
         Left            =   2940
         TabIndex        =   40
         Top             =   570
         Width           =   765
      End
      Begin VB.Label lblPatID 
         Caption         =   "门诊号"
         Height          =   225
         Left            =   420
         TabIndex        =   38
         Top             =   570
         Width           =   555
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "检查类别(&T)"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   180
         Width           =   990
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   6255
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4795
      TabIndex        =   29
      Top             =   5145
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3540
      TabIndex        =   28
      Top             =   5145
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   150
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   5130
      Width           =   1100
   End
   Begin VB.Frame Frame2 
      Height          =   2350
      Left            =   -100
      TabIndex        =   34
      Top             =   1680
      Width           =   6250
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   8
         Left            =   1275
         MaxLength       =   20
         TabIndex        =   13
         Top             =   1230
         Width           =   1455
      End
      Begin VB.ComboBox cboRoom 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   60
         Width           =   4620
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   7
         Left            =   1275
         MaxLength       =   30
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   450
         Width           =   1455
      End
      Begin VB.ComboBox cboSex 
         Height          =   300
         ItemData        =   "frmPacsReg.frx":000C
         Left            =   3795
         List            =   "frmPacsReg.frx":0019
         TabIndex        =   15
         Text            =   "cboSex"
         Top             =   1230
         Width           =   735
      End
      Begin MSComCtl2.DTPicker DTBirth 
         Height          =   300
         Left            =   1275
         TabIndex        =   42
         Top             =   1230
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   80478211
         CurrentDate     =   38156
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   1
         Left            =   3795
         MaxLength       =   30
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   450
         Width           =   2055
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   4
         Left            =   5280
         MaxLength       =   4
         TabIndex        =   17
         Top             =   1230
         Width           =   570
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   2
         Left            =   1275
         MaxLength       =   20
         TabIndex        =   9
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   3
         Left            =   3795
         MaxLength       =   30
         TabIndex        =   11
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   6
         Left            =   3795
         MaxLength       =   3
         TabIndex        =   21
         Top             =   1620
         Width           =   1815
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   5
         Left            =   1275
         MaxLength       =   3
         TabIndex        =   19
         Top             =   1620
         Width           =   1215
      End
      Begin VB.CheckBox chk病理 
         Caption         =   "病理检查(&C)"
         Height          =   255
         Left            =   1275
         TabIndex        =   22
         Top             =   1995
         Width           =   1455
      End
      Begin VB.CheckBox chk胶片 
         Caption         =   "发放胶片(&F)"
         Height          =   255
         Left            =   3795
         TabIndex        =   23
         Top             =   1995
         Width           =   1335
      End
      Begin VB.Label lblRoom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "执行间(&R)"
         Height          =   180
         Left            =   420
         TabIndex        =   2
         Top             =   120
         Width           =   810
      End
      Begin VB.Label Label6 
         Caption         =   "检查号(&U)"
         Height          =   255
         Left            =   420
         TabIndex        =   4
         Top             =   510
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "KG"
         Height          =   180
         Left            =   5680
         TabIndex        =   37
         Top             =   1680
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "CM"
         Height          =   180
         Left            =   2520
         TabIndex        =   36
         Top             =   1680
         Width           =   180
      End
      Begin VB.Label Label3 
         Height          =   135
         Left            =   0
         TabIndex        =   35
         Top             =   -20
         Width           =   6255
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "检查设备(&D)"
         Height          =   180
         Index           =   8
         Left            =   2760
         TabIndex        =   6
         Top             =   510
         Width           =   990
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "年龄(&A)"
         Height          =   180
         Index           =   6
         Left            =   4600
         TabIndex        =   16
         Top             =   1290
         Width           =   630
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "姓名(&N)"
         Height          =   180
         Index           =   1
         Left            =   600
         TabIndex        =   8
         Top             =   900
         Width           =   630
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "英文名(&E)"
         Height          =   180
         Index           =   4
         Left            =   2940
         TabIndex        =   10
         Top             =   900
         Width           =   810
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "性别(&S)"
         Height          =   180
         Index           =   5
         Left            =   3120
         TabIndex        =   14
         Top             =   1290
         Width           =   630
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "电话(&B)"
         Height          =   180
         Index           =   2
         Left            =   600
         TabIndex        =   12
         Top             =   1290
         Width           =   630
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "体重(&W)"
         Height          =   180
         Index           =   3
         Left            =   3120
         TabIndex        =   20
         Top             =   1680
         Width           =   630
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "身高(&H)"
         Height          =   180
         Index           =   7
         Left            =   600
         TabIndex        =   18
         Top             =   1680
         Width           =   630
      End
   End
   Begin VB.Label Label1 
      Caption         =   $"frmPacsReg.frx":0028
      Height          =   615
      Left            =   840
      TabIndex        =   31
      Top             =   120
      Width           =   4815
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "frmPacsReg.frx":00BA
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmPACSReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private AdviceID As Long, SendNO As Long, mlngPatientID As Long
Private iReturn As Integer, blnModi As Boolean
Private aDevices() As String

Public Function ShowMe(objParent As Object, ByVal lngAdviceID As Long, ByVal lngSendNO As Long) As Integer
    '返回：0＝取消、1=开始检查、2＝修改检查信息
    AdviceID = lngAdviceID: SendNO = lngSendNO
    
    blnModi = False
    Me.Show vbModal, objParent
    ShowMe = iReturn
End Function

Private Sub cboRoom_Click()
    On Error Resume Next
    txtItem(1) = aDevices(cboRoom.ListIndex)
End Sub

Private Sub cboRoom_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboSex_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkUnicode_Click()
    If Not blnModi Then
        txtItem(7).Text = Next检查号(Me.txtItem(0), mlngPatientID, AdviceID, SendNO, chkUnicode.Value = 1)
    End If
End Sub

Private Sub chk病理_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk胶片_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo DBError
    If Len(Trim(txtItem(2))) = 0 Then
        MsgBox "请输入姓名！", vbInformation, gstrSysName
        txtItem(2).SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(txtItem(2).Text), vbFromUnicode)) > txtItem(2).MaxLength Then
        MsgBox "姓名超长（最多" & txtItem(2).MaxLength & "个字符或" & CInt(txtItem(2).MaxLength / 2) & "个汉字）！", vbInformation, gstrSysName
        txtItem(2).SetFocus: Exit Sub
    End If
    If Len(Trim(txtItem(3))) = 0 Then
        MsgBox "请输入英文名！", vbInformation, gstrSysName
        txtItem(3).SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(txtItem(3).Text), vbFromUnicode)) > txtItem(3).MaxLength Then
        MsgBox "英文名超长（最多" & txtItem(3).MaxLength & "个字符或" & CInt(txtItem(3).MaxLength / 2) & "个汉字）！", vbInformation, gstrSysName
        txtItem(3).SetFocus: Exit Sub
    End If
    
    '判断检查号是否重复
    strSQL = "Select 姓名,性别,年龄 From 影像检查记录 Where 影像类别=[1] And 检查号=[2]"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, txtItem(0), txtItem(7))
    If Not rsTmp.EOF Then
        If MsgBox("当前检查号与下列患者重复！是否继续？" & Chr(10) & Chr(13) & "患者信息：" & Nvl(rsTmp(0)) & " " & Nvl(rsTmp(1)) & " " & Nvl(rsTmp(2)), vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            txtItem(7).SetFocus: Exit Sub
        End If
    End If
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "当前执行间", Me.cboRoom.Text
    strSQL = ""
    For i = 0 To cboRoom.ListCount - 1
        strSQL = strSQL & "||" & cboRoom.List(i) & "|" & aDevices(i)
    Next
    If Len(strSQL) > 0 Then strSQL = Mid(strSQL, 3)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "检查设备", strSQL
    For i = 0 To optMatch.Count - 1
        If optMatch(i).Value Then Exit For
    Next
    If i > optMatch.Count - 1 Then i = 0
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\开始检查", "影像匹配方式", i
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\开始检查", "按执行科室编号", chkUnicode.Value
    
    gcnOracle.BeginTrans
    strSQL = "ZL_影像检查_BEGIN('" & cboRoom.Text & "'," & txtItem(7).Text & "," & AdviceID & "," & SendNO & ",'" & txtItem(0) & "','" & _
        Trim(txtItem(2)) & "','" & Trim(txtItem(3)) & "','" & Trim(cboSex.Text) & "','" & _
        txtItem(4) & "'," & IIf(IsNull(DTBirth.Value), "Null", "to_Date('" & Format(DTBirth.Value, "yyyy-MM-dd") & "','YYYY-MM-DD')") & ",'" & txtItem(5) & "','" & txtItem(6) & "'," & _
        Me.chk病理.Value & "," & Me.chk胶片.Value & ",'" & Trim(txtItem(1)) & "'," & _
        IIf(blnModi, 1, 0) & ",'" & txtItem(8) & "')"
    ExecuteProc strSQL, Me.Caption
        
    '查找提前进行的检查
    strSQL = "Select A.检查UID As ID From 影像临时记录 a " & _
        " Where a.检查号=[1] And a.影像类别=[2]"
    If optMatch(0).Value Then '检查号
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, txtItem(7).Text, txtItem(0).Text)
    End If
    If optMatch(1).Value Then '门诊/住院号
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, txtPatID.Text, txtItem(0).Text)
    End If
    If optMatch(2).Value Then '检查标识号（医嘱ID）
        Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, AdviceID, txtItem(0).Text)
    End If
    If rsTmp.RecordCount = 1 Then '将图像和检查自动匹配
        strSQL = "ZL_影像检查_SET(" & AdviceID & "," & SendNO & ",'" & _
            rsTmp("ID") & "')"
        ExecuteProc strSQL, Me.Caption
    End If
    gcnOracle.CommitTrans
    
    iReturn = IIf(blnModi, 2, 1)
    Unload Me
    Exit Sub
DBError:
    gcnOracle.RollbackTrans
    Call ErrCenter
    txtItem(7).SetFocus
    Call SaveErrLog
End Sub

Private Sub DTBirth_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Activate()
    If Me.Tag = "Unload" Then
        Me.Tag = ""
        Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    Call cmdCancel_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim strExeRoom As String
    
    iReturn = 0
    
    On Error GoTo DBError
    
    strSQL = "Select Nvl(A.执行部门ID,0) As 执行部门ID,Nvl(E.姓名,C.姓名) As 姓名,Nvl(E.年龄,C.年龄) As 年龄," & _
        "Nvl(E.性别,C.性别) As 性别,Nvl(E.出生日期,C.出生日期) As 出生日期," & _
        "Nvl(D.影像类别,' ') As 影像类别,E.病理检查,E.发放胶片," & _
        "Nvl(D.可行病检,0) As 可行病检,Nvl(D.可发胶片,0) As 可发胶片," & _
        "E.检查号,E.英文名,E.身高,E.体重,E.检查设备,A.执行间,Nvl(A.执行状态,0) As 执行状态,B.病人ID," & _
        "Nvl(E.联系电话,C.联系人电话) As 联系人电话,B.病人来源,Decode(B.病人来源,2,C.住院号,C.门诊号) As 标识号,F.名称 As 申请科室 " & _
        "From 病人医嘱发送 A,病人医嘱记录 B,病人信息 C,影像检查项目 D,影像检查记录 E,部门表 F " & _
        "Where A.医嘱ID=B.ID And B.病人ID=C.病人ID And B.诊疗项目ID=D.诊疗项目ID(+) " & _
        "And A.医嘱ID=E.医嘱ID(+) And A.发送号=E.发送号(+) And B.开嘱科室ID=F.ID " & _
        "And A.医嘱ID= [1] And A.发送号=[2] "
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, AdviceID, SendNO)
    
    If rsTmp.EOF Then
        MsgBox "不能正确读取执行项目信息。", vbInformation, gstrSysName
        Me.Tag = "Unload": Exit Sub
    End If
    If rsTmp("执行状态") = 1 Or rsTmp("执行状态") = 2 Then
        MsgBox "该检查已被其他人" & IIf(rsTmp("执行状态") = 1, "执行完成。", "拒绝执行。"), vbInformation, gstrSysName
        Me.Tag = "Unload": Exit Sub
    End If
    
    mlngPatientID = Nvl(rsTmp!病人ID, 0)
    Me.txtItem(0) = rsTmp("影像类别")
    Me.lblPatID.Caption = IIf(Nvl(rsTmp("病人来源"), 0) = 2, "住院号", "门诊号")
    Me.txtPatID = Nvl(rsTmp("标识号"))
    Me.txtDept = Nvl(rsTmp("申请科室"))
    Me.txtItem(2) = rsTmp("姓名")
    Me.txtItem(4) = Nvl(rsTmp("年龄")): Me.cboSex.Text = Nvl(rsTmp("性别"), " ")
    If IsNull(rsTmp!出生日期) Then
        DTBirth.Value = Empty
    Else
        DTBirth.Value = rsTmp!出生日期
    End If
    Me.txtItem(8) = Nvl(rsTmp("联系人电话"))
    chk病理.Value = Nvl(rsTmp!病理检查, 0)
    Select Case rsTmp("可行病检")
        Case 0, 1
            chk病理.Value = rsTmp("可行病检"): chk病理.Enabled = False
        Case Else
            chk病理.Enabled = True
    End Select
    
    chk胶片.Value = Nvl(rsTmp!发放胶片, 0)
    Select Case rsTmp("可发胶片")
        Case 0, 1
            chk胶片.Value = rsTmp("可发胶片"): chk胶片.Enabled = False
        Case Else
            chk胶片.Enabled = True
    End Select
    txtItem(1).Text = Nvl(rsTmp!检查设备)
    txtItem(3).Text = Nvl(rsTmp!英文名, UCase(Replace(zlCommFun.mGetFullPY(Trim(txtItem(2))), vbCrLf, "")))
    txtItem(5).Text = Nvl(rsTmp!身高)
    txtItem(6).Text = Nvl(rsTmp!体重)
    If Not IsNull(rsTmp!检查号) Then
        txtItem(7).Text = rsTmp!检查号
        blnModi = True
    Else
        txtItem(7).Text = Next检查号(rsTmp!影像类别, Nvl(rsTmp!病人ID, 0), AdviceID, SendNO, chkUnicode.Value = 1)
    End If
    
    '执行间内容
    strExeRoom = Nvl(rsTmp("执行间"))
    If Len(Trim(strExeRoom)) = 0 Then '取默认本地执行间
        strExeRoom = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName, "当前执行间")
    End If
    If rsTmp("执行部门ID") = 0 Then
        strSQL = "Select * From 医技执行房间"
    Else
        strSQL = "Select * From 医技执行房间 Where 科室ID=[1]"
    End If
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption, Val(rsTmp("执行部门ID")))
    cboRoom.Clear
    If rsTmp.EOF Then
        cboRoom.AddItem "": cboRoom.ListIndex = 0
    Else
        Do While Not rsTmp.EOF
            cboRoom.AddItem rsTmp!执行间
            rsTmp.MoveNext
        Loop
    End If
    InitDevice
    
    '影像匹配设置
    i = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\开始检查", "影像匹配方式", 0))
    optMatch(i).Value = True
    chkUnicode.Value = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\开始检查", "按执行科室编号", 0))
    
    If blnModi Then
        Me.Caption = "修改影像信息"
    Else
        Me.Caption = "开始影像检查"
    End If
    On Error Resume Next
    cboRoom.ListIndex = 0
    cboRoom.Text = strExeRoom
    On Error GoTo DBError
    
    Exit Sub
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub optMatch_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtItem_GotFocus(Index As Integer)
    With Me.txtItem(Index)
        .SelStart = 0: .SelLength = .MaxLength
    End With
    Select Case Index
        Case 1, 2
            Call zlCommFun.OpenIme(True)
        Case Else
            Call zlCommFun.OpenIme(False)
    End Select
End Sub

Private Sub txtItem_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call zlCommFun.PressKey(vbKeyTab)
        '得到拼音
        If Trim(txtItem(3)) = "" Then
            txtItem(3).Text = UCase(Replace(zlCommFun.mGetFullPY(Trim(txtItem(2).Text)), vbCrLf, ""))
        End If
    End If
End Sub

Private Sub txtItem_KeyPress(Index As Integer, KeyAscii As Integer)
    If ifEditKey(KeyAscii, False) Then Exit Sub
    
    If LenB(StrConv(Trim(txtItem(Index).Text), vbFromUnicode)) >= txtItem(Index).MaxLength Then
        KeyAscii = 0
        Exit Sub
    End If
    Select Case Index
        Case 5, 6, 7
            If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then KeyAscii = 0
    End Select
End Sub

Private Sub txtItem_LostFocus(Index As Integer)
    Select Case Index
        Case 1, 2
            Call zlCommFun.OpenIme(False)
            If Index = 1 Then aDevices(cboRoom.ListIndex) = txtItem(1)
    End Select
End Sub

'判断是否为编辑键
Private Function ifEditKey(ByVal KeyAscii As Integer, Optional ByVal AllowSubtract As Boolean = True) As Boolean
    If KeyAscii = vbKeyBack Or (KeyAscii = vbKeyInsert And AllowSubtract) Or KeyAscii = vbKeyDelete Or _
      KeyAscii = vbKeyHome Or KeyAscii = vbKeyEnd Or KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or _
      KeyAscii = vbKeyEscape Or KeyAscii = vbKeyReturn Then
        ifEditKey = True
    Else
        ifEditKey = False
    End If
End Function

Private Function Next检查号(str类别 As String, ByVal lngPatientID As Long, Optional ByVal lngAdviceID As Long = 0, Optional ByVal lngSendNO As Long = 0, Optional ByVal blnUnicode As Boolean = False) As Double
    Dim rsCtrl As New ADODB.Recordset
    Dim strSQL As String, lngNO As Double
    Dim lngExeDept As Long

ReStart:
    Err = 0
    On Error GoTo errH
    
    If Not blnUnicode Then '按类别编号
        strSQL = "Select 检查号 From 影像检查记录 A,病人医嘱记录 B" & _
            " Where A.医嘱ID=B.ID And B.病人ID=[1] And A.影像类别=[2] Order By B.停嘱时间 Desc"
        Set rsCtrl = OpenSQLRecord(strSQL, Me.Caption, lngPatientID, str类别)
        If Not rsCtrl.EOF Then
            lngNO = Val(Nvl(rsCtrl("检查号"), 0))
        Else
            strSQL = "Select * From 影像检查类别 Where 编码=[1]"
            Set rsCtrl = OpenSQLRecord(strSQL, Me.Caption, str类别)
            If rsCtrl.EOF Then Exit Function
            
            lngNO = Val(Nvl(rsCtrl("最大号码"), 0)) + 1
        End If
    Else '按执行科室编号
        strSQL = "Select A.执行部门ID,B.病人ID From 病人医嘱发送 A,病人医嘱记录 B" & _
            " Where A.医嘱ID=B.ID And A.医嘱ID=[1] And A.发送号=[2]"
        Set rsCtrl = OpenSQLRecord(strSQL, Me.Caption, lngAdviceID, lngSendNO)
        If rsCtrl.EOF Then
            Next检查号 = 0
            Exit Function
        End If
        
        lngExeDept = Nvl(rsCtrl(0), 0)
        strSQL = "Select A.检查号 From 影像检查记录 A,病人医嘱记录 B" & _
            " Where A.医嘱ID=B.ID+0 And B.执行科室ID+0=[1] And B.病人ID=[2] Order By B.停嘱时间 Desc"
        Set rsCtrl = OpenSQLRecord(strSQL, Me.Caption, lngExeDept, lngPatientID)
        If rsCtrl.EOF Then '取类别的最大号码
'            strSQL = "Select * From 影像检查类别 Where 编码=[1]"
'            Set rsCtrl = OpenSQLRecord(strSQL, Me.Caption, str类别)
            strSQL = "SELECT DISTINCT C.编码,Nvl(C.最大号码,0) FROM 影像检查项目 A,诊疗执行科室 B,影像检查类别 C" & _
                " WHERE A.诊疗项目ID=B.诊疗项目ID+0 AND A.影像类别||''=C.编码 AND B.执行科室ID=[1] ORDER BY Nvl(C.最大号码,0) DESC"
            Set rsCtrl = OpenSQLRecord(strSQL, Me.Caption, lngExeDept)
            If rsCtrl.EOF Then
                Next检查号 = 0
                Exit Function
            End If
            
            lngNO = Val(Nvl(rsCtrl(1), 0)) + 1
        Else
            lngNO = Val(Nvl(rsCtrl("检查号"), 0))
        End If
    End If
    
    Next检查号 = lngNO
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub InitDevice()
    Dim i As Integer, iPos As Integer
    Dim strDevices As String, aTmpArray() As String, aTmpArray1() As String
    On Error Resume Next
    
    strDevices = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "检查设备")
    If Len(Trim(strDevices)) = 0 Then
        ReDim aTmpArray(1, 0) As String
    Else
        aTmpArray1 = Split(strDevices, "||")
        ReDim aTmpArray(1, UBound(aTmpArray1)) As String
        For i = 0 To UBound(aTmpArray1)
            iPos = InStr(aTmpArray1(i), "|")
            If iPos = 0 Then
                aTmpArray(0, i) = ""
                aTmpArray(1, i) = aTmpArray1(i)
            Else
                aTmpArray(0, i) = Mid(aTmpArray1(i), 1, iPos - 1)
                aTmpArray(1, i) = Mid(aTmpArray1(i), iPos + 1)
            End If
        Next
    End If
    
    ReDim aDevices(cboRoom.ListCount - 1) As String
    For i = 0 To cboRoom.ListCount - 1
        iPos = GetIndex(aTmpArray, cboRoom.List(i))
        aDevices(i) = aTmpArray(1, iPos)
    Next
End Sub

Private Function GetIndex(aSeekArray() As String, ByVal vSeekValue As Variant) As Integer
    Dim i As Integer
    For i = 0 To UBound(aSeekArray, 2)
        If aSeekArray(0, i) = vSeekValue Then Exit For
    Next
    If i > UBound(aSeekArray, 2) Then
        GetIndex = 0
    Else
        GetIndex = i
    End If
End Function
