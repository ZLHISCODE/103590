VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTechnicPlan 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "执行报到"
   ClientHeight    =   4095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   Icon            =   "frmTechnicPlan.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboRoom 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1170
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   945
      Width           =   4620
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4725
      TabIndex        =   13
      Top             =   3600
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3540
      TabIndex        =   12
      Top             =   3600
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   285
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3585
      Width           =   1100
   End
   Begin VB.Frame fraDetail 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2115
      Left            =   -75
      TabIndex        =   23
      Top             =   1290
      Width           =   6250
      Begin VB.TextBox txtItem 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   120
         Width           =   1185
      End
      Begin VB.ComboBox cboSex 
         Height          =   300
         IMEMode         =   3  'DISABLE
         ItemData        =   "frmTechnicPlan.frx":000C
         Left            =   3570
         List            =   "frmTechnicPlan.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   915
         Width           =   930
      End
      Begin MSComCtl2.DTPicker dtpBirth 
         Height          =   300
         Left            =   1245
         TabIndex        =   5
         Top             =   915
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   98762755
         CurrentDate     =   38156
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   1
         Left            =   3570
         MaxLength       =   30
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   120
         Width           =   2295
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   4
         Left            =   5310
         MaxLength       =   10
         TabIndex        =   7
         Top             =   915
         Width           =   555
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Index           =   2
         Left            =   1245
         MaxLength       =   20
         TabIndex        =   3
         Top             =   510
         Width           =   1185
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   3570
         MaxLength       =   30
         TabIndex        =   4
         Top             =   525
         Width           =   2295
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   6
         Left            =   3570
         MaxLength       =   3
         TabIndex        =   9
         Top             =   1320
         Width           =   915
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   1245
         MaxLength       =   3
         TabIndex        =   8
         Top             =   1320
         Width           =   1185
      End
      Begin VB.CheckBox chk病理 
         Caption         =   "病理检查(&C)"
         Height          =   225
         Left            =   1245
         TabIndex        =   10
         Top             =   1740
         Width           =   1290
      End
      Begin VB.CheckBox chk胶片 
         Caption         =   "发放胶片(&F)"
         Height          =   225
         Left            =   3570
         TabIndex        =   11
         Top             =   1740
         Width           =   1290
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检查号(&H)"
         Height          =   180
         Index           =   0
         Left            =   420
         TabIndex        =   31
         Top             =   180
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "kg"
         Height          =   180
         Left            =   4560
         TabIndex        =   25
         Top             =   1380
         Width           =   180
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "cm"
         Height          =   180
         Left            =   2475
         TabIndex        =   24
         Top             =   1380
         Width           =   180
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "检查设备(&D)"
         Height          =   180
         Index           =   8
         Left            =   2535
         TabIndex        =   15
         Top             =   180
         Width           =   990
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄(&A)"
         Height          =   180
         Index           =   6
         Left            =   4635
         TabIndex        =   20
         Top             =   975
         Width           =   630
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名(&N)"
         Height          =   180
         Index           =   1
         Left            =   600
         TabIndex        =   16
         Top             =   570
         Width           =   630
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "英文名(&E)"
         Height          =   180
         Index           =   4
         Left            =   2715
         TabIndex        =   18
         Top             =   555
         Width           =   810
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别(&S)"
         Height          =   180
         Index           =   5
         Left            =   2895
         TabIndex        =   19
         Top             =   975
         Width           =   630
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生日期(&B)"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   975
         Width           =   990
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "体重(&W)"
         Height          =   180
         Index           =   3
         Left            =   2895
         TabIndex        =   22
         Top             =   1380
         Width           =   630
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身高(&H)"
         Height          =   180
         Index           =   7
         Left            =   600
         TabIndex        =   21
         Top             =   1380
         Width           =   630
      End
   End
   Begin VB.Frame fraSplit2 
      Height          =   120
      Left            =   0
      TabIndex        =   30
      Top             =   3345
      Width           =   6420
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      X1              =   0
      X2              =   8000
      Y1              =   765
      Y2              =   765
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   8000
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Label lblItemDetail 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C00000&
      Height          =   345
      Left            =   1185
      TabIndex        =   29
      Top             =   405
      Width           =   4575
   End
   Begin VB.Label lblRoom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "执行间(&R)"
      Height          =   180
      Left            =   330
      TabIndex        =   28
      Top             =   1005
      Width           =   810
   End
   Begin VB.Label lblPati 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "病人："
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   645
      TabIndex        =   27
      Top             =   165
      Width           =   540
   End
   Begin VB.Label lblItemTit 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "项目："
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   645
      TabIndex        =   26
      Top             =   420
      Width           =   540
   End
End
Attribute VB_Name = "frmTechnicPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng医嘱ID As Long
Private mlng发送号 As Long
Private mlng执行科室ID As Long
Private mrsRoom As ADODB.Recordset
Private mblnOK As Boolean
Private mbln执行报到时结算 As Boolean
Private mlng卡类别ID  As Long
Private mlng病人ID As Long
Private mstrPrivs As String
Private mobjSquareCard As Object      '卡结算对象

Public Function ShowMe(objParent As Object, ByVal lng医嘱ID As Long, ByVal lng发送号 As Long, Optional ByVal lng执行科室ID As Long, Optional ByVal lng卡类别ID As Long, Optional ByVal lng病人ID As Long, Optional ByVal strPrivs As String, Optional ByRef objSquareCard As Object) As Boolean
    mlng医嘱ID = lng医嘱ID
    mlng发送号 = lng发送号
    mlng执行科室ID = lng执行科室ID
    mlng卡类别ID = lng卡类别ID
    mlng病人ID = lng病人ID
    mstrPrivs = strPrivs
    Set mobjSquareCard = objSquareCard
    
    On Local Error Resume Next
    Me.Show 1, objParent
    On Error GoTo 0
    
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String
    Dim blnTrans As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim str医嘱IDs As String
    Dim i As Long
     
    '检查输入内容
    If fraDetail.Visible Then
        If zlCommFun.ActualLen(txtItem(1).Text) > txtItem(1).MaxLength Then
            MsgBox "检查设备最多允许输入 " & txtItem(1).MaxLength \ 2 & " 个汉字或 " & txtItem(1).MaxLength & " 个字符，请检查。", vbInformation, gstrSysName
            txtItem(1).SetFocus: Exit Sub
        End If
        If zlCommFun.ActualLen(txtItem(2).Text) > txtItem(2).MaxLength Then
            MsgBox "姓名最多允许输入 " & txtItem(2).MaxLength \ 2 & " 个汉字或 " & txtItem(2).MaxLength & " 个字符，请检查。", vbInformation, gstrSysName
            txtItem(2).SetFocus: Exit Sub
        End If
        If zlCommFun.ActualLen(txtItem(4).Text) > txtItem(4).MaxLength Then
            MsgBox "年龄最多允许输入 " & txtItem(4).MaxLength \ 2 & " 个汉字或 " & txtItem(4).MaxLength & " 个字符，请检查。", vbInformation, gstrSysName
            txtItem(4).SetFocus: Exit Sub
        End If
        If Trim(txtItem(1).Text) = "" Then
            MsgBox "请输入检查设备。", vbInformation, gstrSysName
            txtItem(1).SetFocus: Exit Sub
        End If
        If Trim(txtItem(2).Text) = "" Then
            MsgBox "请输入病人姓名。", vbInformation, gstrSysName
            txtItem(2).SetFocus: Exit Sub
        End If
    End If
    
    On Error GoTo errH
    '门诊一卡通,病人执行报到前必须先收费或先记帐审核,不传单据号，根据医嘱ID读取所有未收费单据或未审核的记帐单
    If mbln执行报到时结算 Then
        '获取整组医嘱的ID串
        strSQL = "select a.ID from 病人医嘱记录 a where a.相关id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng医嘱ID)
        str医嘱IDs = mlng医嘱ID
        For i = 1 To rsTmp.RecordCount
            str医嘱IDs = str医嘱IDs & "," & rsTmp!ID
            rsTmp.MoveNext
        Next
        If Not mobjSquareCard Is Nothing Then
            If mobjSquareCard.zlSquareAffirm(Me, p医技工作站, mstrPrivs, mlng病人ID, mlng卡类别ID, False, , , str医嘱IDs) = False Then
                Exit Sub
            End If
        Else
            MsgBox "一卡通部件初始化失败，请检查部件。", vbInformation, Me.Caption
            Exit Sub
        End If
    End If
    gcnOracle.BeginTrans: blnTrans = True
    If mlng执行科室ID <> 0 Then
        strSQL = "Zl_病人医嘱发送_科室变更(" & mlng医嘱ID & "," & mlng发送号 & "," & mlng执行科室ID & ")"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
    
    If fraDetail.Visible And Not IsNull(dtpBirth.Value) Then
        strSQL = "To_Date('" & Format(dtpBirth.Value, "yyyy-MM-dd") & "','YYYY-MM-DD')"
    Else
        strSQL = "NULL"
    End If
    strSQL = "ZL_病人医嘱执行_Plan(" & mlng医嘱ID & "," & mlng发送号 & ",1," & _
        "'" & cboRoom.Text & "','" & lblItemDetail.Tag & "'," & ZVal(txtItem(0).Text) & ",'" & txtItem(2).Text & "'," & _
        "'" & txtItem(3).Text & "','" & zlCommFun.GetNeedName(cboSex.Text) & "','" & txtItem(4).Text & "'," & _
        strSQL & "," & ZVal(txtItem(5).Text) & "," & ZVal(txtItem(6).Text) & "," & _
        chk病理.Value & "," & chk胶片.Value & ",'" & txtItem(1).Text & "')"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    gcnOracle.CommitTrans: blnTrans = False
    On Error GoTo 0
    
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub dtpBirth_Change()
    If Not IsNull(dtpBirth.Value) And IsNumeric(txtItem(4).Text) Then
        txtItem(4).Text = CInt(Format(zlDatabase.Currentdate, "yyyy")) - CInt(Format(dtpBirth.Value, "yyyy"))
        If Format(zlDatabase.Currentdate, "MMdd") < Format(dtpBirth.Value, "MMdd") Then
            txtItem(4).Text = CInt(txtItem(4).Text) - 1
        End If
        If CInt(txtItem(4).Text) < 0 Then txtItem(4).Text = ""
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Call cmdHelp_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = 13 Then
        KeyAscii = 0
        Call ZLCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    mblnOK = False
    On Error GoTo errH
    '性别字典
    cboSex.AddItem " "
    cboSex.ListIndex = 0
    strSQL = "Select 编码,名称,简码,缺省标志 From 性别 Order by 编码"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        cboSex.AddItem rsTmp!编码 & "-" & rsTmp!名称
        If Nvl(rsTmp!缺省标志, 0) = 1 Then
            cboSex.ListIndex = cboSex.NewIndex
        End If
        rsTmp.MoveNext
    Next
    
    '执行项目内容
    strSQL = _
        "Select A.执行部门ID,A.执行间,B.医嘱内容 as 内容,D.检查号," & _
        " Nvl(D.姓名,C.姓名) as 姓名,Nvl(D.性别,C.性别) as 性别,Nvl(D.年龄,C.年龄) as 年龄," & _
        " Nvl(D.出生日期,C.出生日期) as 出生日期,D.英文名,D.身高,D.体重," & _
        " D.病理检查,D.发放胶片,D.检查设备," & _
        " F.名称 as 类别名称,E.影像类别,E.可行病检,E.可发胶片" & _
        " From 病人医嘱发送 A,病人医嘱记录 B,病人信息 C," & _
            " 影像检查记录 D,影像检查项目 E,影像检查类别 F" & _
        " Where A.医嘱ID=B.ID And B.病人ID=C.病人ID" & _
        " And B.诊疗项目ID=E.诊疗项目ID(+) And E.影像类别=F.编码(+)" & _
        " And A.医嘱ID=D.医嘱ID(+) And A.发送号=D.发送号(+)" & _
        " And A.医嘱ID=[1] And A.发送号=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng医嘱ID, mlng发送号)
    If rsTmp.EOF Then
        MsgBox "不能正确读取执行项目信息。", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    If IsNull(rsTmp!影像类别) Then
        '普通执行项目
        fraDetail.Visible = False
        Me.Height = Me.Height - fraDetail.Height + 300
        fraSplit2.Top = fraSplit2.Top - fraDetail.Height + 300
        cmdHelp.Top = cmdHelp.Top - fraDetail.Height + 300
        cmdOK.Top = cmdOK.Top - fraDetail.Height + 300
        cmdCancel.Top = cmdCancel.Top - fraDetail.Height + 300
        
        lblRoom.Top = lblRoom.Top + 100
        cboRoom.Top = cboRoom.Top + 100
        lblRoom.Left = lblRoom.Left + 500
        cboRoom.Left = cboRoom.Left + 500
        cboRoom.Width = cboRoom.Width - 1000
        
        lblPati.Caption = "病人：" & rsTmp!姓名
        lblItemDetail.Caption = Nvl(rsTmp!内容)
    Else
        '影像检查项目
        lblPati.Caption = "病人：" & rsTmp!姓名 & "  影像类别：" & rsTmp!影像类别 & "-" & rsTmp!类别名称
        lblItemTit.Caption = "影像检查："
        lblItemDetail.Caption = Nvl(rsTmp!内容)
        lblItemDetail.Tag = rsTmp!影像类别
        
        If Not IsNull(rsTmp!检查号) Then
            txtItem(0).Text = rsTmp!检查号
        Else
            txtItem(0).Text = Next检查号(rsTmp!影像类别)
        End If
        txtItem(1).Text = Nvl(rsTmp!检查设备)
        txtItem(2).Text = rsTmp!姓名
        If IsNull(rsTmp!英文名) Then
            txtItem(3).Text = ZLCommFun.SpellCode(rsTmp!姓名)
        Else
            txtItem(3).Text = rsTmp!英文名
        End If
        If IsNull(rsTmp!出生日期) Then
            dtpBirth.Value = Empty
        Else
            dtpBirth.Value = rsTmp!出生日期
        End If
        If Not IsNull(rsTmp!性别) Then
             Cbo.SeekIndex cboSex, rsTmp!性别, True
        End If
        txtItem(4).Text = Nvl(rsTmp!年龄)
        txtItem(5).Text = Nvl(rsTmp!身高)
        txtItem(6).Text = Nvl(rsTmp!体重)
        
        chk病理.Value = IIf(Nvl(rsTmp!病理检查, 0) = 0, 0, 1)
        If Nvl(rsTmp!可行病检, 0) = 0 Then
            chk病理.Enabled = False
            chk病理.Value = 0
        ElseIf Nvl(rsTmp!可行病检, 0) = 1 Then
            chk病理.Enabled = False
            chk病理.Value = 1
        End If
        
        chk胶片.Value = IIf(Nvl(rsTmp!发放胶片, 0) = 0, 0, 1)
        If Nvl(rsTmp!可发胶片, 0) = 0 Then
            chk胶片.Enabled = False
            chk胶片.Value = 0
        ElseIf Nvl(rsTmp!可发胶片, 0) = 1 Then
            chk胶片.Enabled = False
            chk胶片.Value = 1
        End If
    End If
    mbln执行报到时结算 = Val(zlDatabase.GetPara("执行报到时收费或记账审核", glngSys, p医技工作站)) = 1
    '执行间内容
    strSQL = "Select 科室ID,执行间,当前分配,检查设备,简码 From 医技执行房间 Where 科室ID=[1]"
    'Set mrsRoom = New ADODB.Recordset
    Set mrsRoom = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(mlng执行科室ID <> 0, mlng执行科室ID, Val(rsTmp!执行部门ID)))
    If mrsRoom.EOF And IsNull(rsTmp!影像类别) Then
        MsgBox "当前科室还没有设置执行间，请先设置。", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
    
    For i = 1 To mrsRoom.RecordCount
        cboRoom.AddItem mrsRoom!执行间
        mrsRoom.MoveNext
    Next
    If cboRoom.ListCount > 0 Then
        cboRoom.ListIndex = 0
    End If
    If Not IsNull(rsTmp!执行间) Then
        Call Cbo.SeekIndex(cboRoom, rsTmp!执行间, True)
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mlng医嘱ID = 0
    mlng发送号 = 0
    Set mrsRoom = Nothing
    Set mobjSquareCard = Nothing
End Sub

Private Sub txtItem_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtItem(Index))
End Sub

Private Sub txtItem_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 5 Or Index = 6 Then
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    ElseIf Index = 3 Then
        If Not Between(KeyAscii, 32, 128) Then KeyAscii = 0
    End If
End Sub

Private Sub txtItem_Validate(Index As Integer, Cancel As Boolean)
    If Index = 5 Or Index = 6 Then
        If Val(txtItem(Index).Text) = 0 Then
            txtItem(Index).Text = ""
        End If
    End If
End Sub

Public Function Next检查号(str类别 As String) As Double
    Dim rsCtrl As New ADODB.Recordset
    Dim strSQL As String, dblNO As Double

ReStart:
    err = 0
    On Error GoTo errH
    With rsCtrl
        If .State = 1 Then .Close
        strSQL = "Select 排列,最大号码,编码,名称,简码 From 影像检查类别 Where 编码='" & str类别 & "'"
        Call SQLTest(App.ProductName, Me.Caption, strSQL)
        .CursorLocation = adUseClient
        .Open strSQL, gcnOracle, adOpenKeyset, adLockOptimistic
        Call SQLTest
        If .EOF Then Exit Function
        
        dblNO = Val(Nvl(!最大号码, 0)) + 1
        
        On Error Resume Next
        .Update "最大号码", dblNO
        If err <> 0 Then
            .CancelUpdate
            GoTo ReStart
        End If
        Next检查号 = dblNO
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
