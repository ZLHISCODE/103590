VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRegist 
   BackColor       =   &H80000005&
   Caption         =   "用户注册管理"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8025
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   Picture         =   "frmRegist.frx":0000
   ScaleHeight     =   7425
   ScaleWidth      =   8025
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdSpecReg 
      Caption         =   "查看其他授权项(&O)"
      Height          =   345
      Left            =   5040
      TabIndex        =   15
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1095
      TabIndex        =   14
      Tag             =   "查找"
      Top             =   3300
      Width           =   4170
   End
   Begin VSFlex8Ctl.VSFlexGrid vsInfo 
      Height          =   1275
      Left            =   735
      TabIndex        =   0
      Top             =   555
      Width           =   6360
      _cx             =   1980181458
      _cy             =   1980172489
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmRegist.frx":04F9
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComctlLib.ProgressBar pgbRegist 
      Height          =   75
      Left            =   375
      TabIndex        =   1
      Top             =   2415
      Visible         =   0   'False
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   132
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdVerify 
      Caption         =   "校验(&V)"
      Height          =   350
      Left            =   3945
      TabIndex        =   4
      Top             =   2550
      Width           =   1100
   End
   Begin VB.OptionButton optGrade 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "功能"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   2
      Left            =   6885
      TabIndex        =   9
      Top             =   3345
      Width           =   675
   End
   Begin VB.OptionButton optGrade 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "程序"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   6105
      TabIndex        =   8
      Top             =   3345
      Value           =   -1  'True
      Width           =   675
   End
   Begin VB.OptionButton optGrade 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "系统"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   5325
      TabIndex        =   7
      Top             =   3345
      Width           =   675
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "还原(&C)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   6300
      TabIndex        =   6
      Top             =   2550
      Width           =   1100
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "应用(&A)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5115
      TabIndex        =   5
      Top             =   2550
      Width           =   1100
   End
   Begin VB.CommandButton cmdRegist 
      Caption         =   "重新注册(&R)…"
      Height          =   350
      Left            =   375
      TabIndex        =   2
      Top             =   2550
      Width           =   1440
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgFunc 
      Height          =   3555
      Left            =   375
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3705
      Width           =   7170
      _cx             =   12647
      _cy             =   6271
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   16777215
      ForeColorFixed  =   -2147483630
      BackColorSel    =   13811126
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Label lblFind 
      BackStyle       =   0  'Transparent
      Caption         =   "查找(&Z)"
      Height          =   255
      Left            =   375
      TabIndex        =   13
      Top             =   3345
      Width           =   690
   End
   Begin VB.Label lblRegist 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "正在注册，请稍等..."
      Height          =   210
      Left            =   1905
      TabIndex        =   3
      Top             =   2655
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户注册管理"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   255
      TabIndex        =   12
      Top             =   120
      Width           =   1440
   End
   Begin VB.Image imgMain 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   150
      Picture         =   "frmRegist.frx":05FB
      Top             =   570
      Width           =   480
   End
   Begin VB.Label lblRegFunc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "已安装系统的应用授权："
      Height          =   180
      Left            =   375
      TabIndex        =   11
      Top             =   3060
      Width           =   1980
   End
End
Attribute VB_Name = "frmRegist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const conIdent = 4  '程序和功能授权记录的缩进空格

Dim strSQL As String
Dim lngCount As Long
'---------------------------------------------
Dim mstrRegCode As String      '暂时保存了注册码
Dim mblnIsCancel As Boolean
Dim mintIndex As Integer
Dim mintCount As Integer        '定位时记录上一次定位的位置


Private Sub cmdApply_Click()
    Dim blnAudit As Boolean
    Dim strRegError As String
    
    err = 0: On Error GoTo errHand
     
    Me.MousePointer = vbHourglass
    
    gcnOracle.Execute "call zltools.p_Reg_Apply()", , adCmdText
    
    Me.Tag = ""
    Me.cmdApply.Enabled = False
    Me.cmdCancel.Enabled = False
    
    '再次调用验证，以保证信息正确性
    strRegError = gobjRegister.zlRegCheck(False)
    Me.MousePointer = vbDefault
    
    If strRegError = "" Then
        SaveSetting "ZLSOFT", "注册信息", "单位名称", gobjRegister.zlRegInfo("单位名称", , -1)
        MsgBox "注册授权信息已经应用！", vbInformation, gstrSysName
    Else
        MsgBox strRegError, vbExclamation, gstrSysName
    End If
    Exit Sub
errHand:
    MsgBox "应用失败，请检查文件的正确性！" & vbNewLine & err.Description, vbExclamation, gstrSysName
End Sub

Private Sub cmdCancel_Click()
    Call zlRefGrant
    Me.Tag = ""
    Me.vfgFunc.SetFocus
    Me.cmdApply.Enabled = False
    Me.cmdCancel.Enabled = False
    MsgBox "注册授权信息已还原！", vbInformation, gstrSysName
End Sub

Private Sub cmdRegist_Click()
    Dim strFile As String, strRegError As String
    Dim rsFile As New ADODB.Recordset
    Dim blnApplyEnabled As Boolean, blnCancelEnabled As Boolean, blnVerifyEnabled As Boolean
    Dim i As Integer, blnNotPrompt As Boolean
    
    With frmMDIMain.DlgMain
        .FileName = ""
        .DialogTitle = "选择注册授权文件"
        .Filter = "(注册授权文件)|*.zcr"
        .ShowOpen
        If .FileName = "" Then Exit Sub
        strFile = .FileName
    End With
        
    Me.cmdRegist.Enabled = False
    
    '记录按钮原来的enabled属性
    blnApplyEnabled = Me.cmdApply.Enabled
    blnCancelEnabled = Me.cmdCancel.Enabled
    blnVerifyEnabled = Me.cmdVerify.Enabled
    
    '禁用按钮
    Me.cmdApply.Enabled = False
    Me.cmdCancel.Enabled = False
    Me.cmdVerify.Enabled = False
    For i = 0 To optGrade.UBound
        If Not optGrade(i).value Then optGrade(i).Enabled = False
    Next
    For i = 0 To optGrade.UBound
        If optGrade(i).value Then optGrade(i).Enabled = False
    Next
    
    err = 0: On Error GoTo errHand
        
    lblRegist.Visible = True
    Me.MousePointer = vbHourglass
    
    If gobjRegister.zlRegBuild(strFile, pgbRegist) = False Then
        lblRegist.Visible = False
        Me.MousePointer = vbDefault
        
        blnNotPrompt = True
        GoTo errHand
    End If
    
    lblRegist.Visible = False
    Me.MousePointer = vbDefault
    
    Me.cmdRegist.Enabled = True
    
    '还原按钮的enabled属性
    Me.cmdApply.Enabled = blnApplyEnabled
    Me.cmdCancel.Enabled = blnCancelEnabled
    Me.cmdVerify.Enabled = blnVerifyEnabled
    '启用控件
    optGrade(0).Enabled = True
    optGrade(1).Enabled = True
    optGrade(2).Enabled = True
    
    strRegError = gobjRegister.zlRegCheck(True)
    If strRegError = "" Then
        Call zlRefGrant(True)
        Me.Tag = "修改"
        cmdApply.Enabled = True
        cmdCancel.Enabled = True
    Else
        Call zlRefGrant
        MsgBox strRegError & vbCrLf & "系统已经自动还原！", vbExclamation, gstrSysName
        Me.Tag = ""
        Me.cmdApply.Enabled = False
        Me.cmdCancel.Enabled = False
    End If
    Me.vfgFunc.SetFocus
    Exit Sub

errHand:
    Me.cmdRegist.Enabled = True
    Me.cmdApply.Enabled = False
    
    '还原按钮的enabled属性
    Me.cmdCancel.Enabled = blnCancelEnabled
    Me.cmdVerify.Enabled = blnVerifyEnabled
    '启用控件
    optGrade(0).Enabled = True
    optGrade(1).Enabled = True
    optGrade(2).Enabled = True
    
    If Not blnNotPrompt Then MsgBox "注册授权文件时出现错误，请检查！" & vbNewLine & err.Description, vbExclamation, Me.Caption
End Sub

Private Sub cmdSpecReg_Click()
    Dim objPopup As CommandBar
    Dim objControl As CommandBarControl
    Dim rsTemp As ADODB.Recordset
    Dim vRect As RECT, strSQL As String
    Dim blnFirst As Boolean
    
    On Error GoTo errHandle
    
    Set objPopup = gcbsMain.Add("Popup", xtpBarPopup)
    
    strSQL = "Select Item, Prog, Text From Table(Cast(zltools.f_Reg_Info(" & IIf(cmdApply.Enabled, 1, 0) & ") As zlTools.t_Reg_Rowset))"
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
    
    With objPopup.Controls
        rsTemp.Filter = "Item='移动护士站授权性质'"
        If Not rsTemp.EOF Then
            Set objControl = .Add(xtpControlButton, 0, rsTemp!Item & "：" & Decode(Val(Nvl(rsTemp!Text)), 1, "正式", 2, "试用", 3, "测试"))
        End If
        rsTemp.Filter = "Item='移动护士站授权期限'"
        If Not rsTemp.EOF Then
            Set objControl = .Add(xtpControlButton, 0, rsTemp!Item & "：" & Decode(Val(Nvl(rsTemp!Text)), 0, "无限制", rsTemp!Text & "天"))
        End If
        rsTemp.Filter = "Item='移动护士站设备数量'"
        If Not rsTemp.EOF Then
            Set objControl = .Add(xtpControlButton, 0, rsTemp!Item & "：" & Decode(Val(Nvl(rsTemp!Text)), 0, "无限制", rsTemp!Text & "台"))
        End If
        rsTemp.Filter = "Item='移动医生站授权性质'"
        If Not rsTemp.EOF Then
            Set objControl = .Add(xtpControlButton, 0, rsTemp!Item & "：" & Decode(Val(Nvl(rsTemp!Text)), 1, "正式", 2, "试用", 3, "测试"))
        End If
        rsTemp.Filter = "Item='移动医生站授权期限'"
        If Not rsTemp.EOF Then
            Set objControl = .Add(xtpControlButton, 0, rsTemp!Item & "：" & Decode(Val(Nvl(rsTemp!Text)), 0, "无限制", rsTemp!Text & "天"))
        End If
        rsTemp.Filter = "Item='移动医生站设备数量'"
        If Not rsTemp.EOF Then
            Set objControl = .Add(xtpControlButton, 0, rsTemp!Item & "：" & Decode(Val(Nvl(rsTemp!Text)), 0, "无限制", rsTemp!Text & "台"))
        End If
    End With
    
    rsTemp.Filter = "Prog=-1"
    If Not rsTemp.EOF Then
        blnFirst = True
        With objPopup.Controls
            rsTemp.MoveFirst
            Do While Not rsTemp.EOF
                Set objControl = .Add(xtpControlButton, 0, rsTemp!Item & "：" & rsTemp!Text)
                If blnFirst Then objControl.BeginGroup = True
                blnFirst = False
                rsTemp.MoveNext
            Loop
        End With
    End If
        
    If objPopup.Controls.Count > 0 Then
        GetWindowRect Me.hwnd, vRect
        objPopup.ShowPopup , vRect.Left * Screen.TwipsPerPixelX + cmdSpecReg.Left, vRect.Top * Screen.TwipsPerPixelY + cmdSpecReg.Top + cmdSpecReg.Height
    Else
        MsgBox "无其他特定授权项目内容。", vbInformation, Me.Caption
    End If
    
    Exit Sub
errHandle:
    MsgBox err.Number & ":" & err.Description, vbExclamation, Me.Caption
End Sub

Private Sub cmdVerify_Click()
    Dim strRegError As String
    Me.MousePointer = vbHourglass
    strRegError = gobjRegister.zlRegCheck(IIf(Me.Tag = "修改", True, False))
    Me.MousePointer = vbDefault
    If strRegError = "" Then
        MsgBox "当前注册授权文件正确有效！", vbInformation, gstrSysName
    Else
        MsgBox strRegError, vbExclamation, gstrSysName
    End If
End Sub

Private Sub Form_Activate()
    Call zlRefGrant
End Sub

Private Sub Form_Deactivate()
    If Tag = "修改" Then
        If MsgBox("已经修改了注册信息，如果不保存，将被自动还原。" & vbCr & "是否保存？", vbQuestion + vbYesNo) = vbYes Then
            Call cmdApply_Click
        Else
            Call cmdCancel_Click
        End If
    End If
    
End Sub

Private Sub Form_Load()
    '搜索框初始化
    txtFind.Text = "请输入编号或关键字"
    txtFind.ForeColor = vbGrayText
    mintCount = -1
    
    mblnIsCancel = False
End Sub

Private Sub Form_Resize()
    err = 0: On Error Resume Next
    Me.vfgFunc.Height = Me.ScaleHeight - Me.vfgFunc.Top - 150
    
End Sub

Private Sub optGrade_Click(Index As Integer)
    With Me.vfgFunc
        .Redraw = flexRDNone
        For lngCount = .FixedRows To .Rows - 1
            Select Case Index
            Case 0
                .RowHidden(lngCount) = (Val(.TextMatrix(lngCount, 2)) > -2)
            Case 1
                .RowHidden(lngCount) = (Val(.TextMatrix(lngCount, 2)) > -1)
            Case 2
                .RowHidden(lngCount) = False
            End Select
        Next
        .Redraw = flexRDDirect
    End With
End Sub
    
'--------------------------------------------------
'功能：按数据库或文件刷新授权信息
'参数：blnTemp-是否按文件刷新
'--------------------------------------------------
Private Sub zlRefGrant(Optional blnTemp As Boolean)
    Dim rsTemp As New ADODB.Recordset
    Dim intKind As Integer, intLimit As Integer, intStation As Integer
    Dim strUnitName As String, i As Integer
    
    On Error GoTo errHandle
    '授权信息
    With vsInfo
        strUnitName = gobjRegister.zlRegInfo("单位名称", blnTemp, -1)
        .TextMatrix(0, 1) = Replace(strUnitName, ";", vbCrLf)
        If strUnitName <> "" Then
            i = UBound(Split(strUnitName, ";")) + 1
        Else
            i = 1
        End If
        .Height = .rowHeight(1) * (5.5 + i - 1)
        .rowHeight(0) = .rowHeight(1) * i
        .Cell(flexcpAlignment, 0, 0, 0, 0) = flexAlignRightTop
        .Cell(flexcpAlignment, 0, 2, 0, 2) = flexAlignRightTop
        .Cell(flexcpAlignment, 0, 1, 0, 1) = flexAlignLeftTop
        .Cell(flexcpAlignment, 0, 3, 0, 3) = flexAlignLeftTop
        cmdSpecReg.Top = .Top + .Height + 30
        
        Select Case Val(gobjRegister.zlRegInfo("授权性质", blnTemp))
            Case 1: intKind = 1: .TextMatrix(1, 1) = "正式版本"
            Case 2: intKind = 2: .TextMatrix(1, 1) = "试用版本"
            Case Else: intKind = 3: .TextMatrix(1, 1) = "测试版本"
        End Select
        
        .TextMatrix(2, 1) = "无限制"
        If intKind <> 1 Then
            intLimit = Val(gobjRegister.zlRegInfo("使用期限", blnTemp))
            If intLimit > 0 Then .TextMatrix(2, 1) = "限用" & intLimit & "天"
        End If
        
        intStation = Val(gobjRegister.zlRegInfo("授权站点", blnTemp))
        If intStation = 0 Then
            .TextMatrix(3, 1) = "无限制"
        Else
            .TextMatrix(3, 1) = "不超过" & intStation & "站点"
        End If
        .TextMatrix(4, 1) = gobjRegister.zlRegInfo("授权日期", blnTemp)
        
        'PACS/LIS授权
        .TextMatrix(0, 3) = gobjRegister.zlRegInfo("影像DICOM设备数量", blnTemp)
        If .TextMatrix(0, 3) = "" Then .TextMatrix(0, 3) = "无限制"
    
        .TextMatrix(1, 3) = gobjRegister.zlRegInfo("影像视频设备数量", blnTemp)
        If .TextMatrix(1, 3) = "" Then .TextMatrix(1, 3) = "无限制"
    
        .TextMatrix(2, 3) = gobjRegister.zlRegInfo("影像胶片打印机数量", blnTemp)
        If .TextMatrix(2, 3) = "" Then .TextMatrix(2, 3) = "无限制"
    
        .TextMatrix(3, 3) = gobjRegister.zlRegInfo("影像观片站数量", blnTemp)
        If .TextMatrix(3, 3) = "" Then .TextMatrix(3, 3) = "无限制"
    
        .TextMatrix(4, 3) = gobjRegister.zlRegInfo("检验仪器数量", blnTemp)
        If .TextMatrix(4, 3) = "" Then .TextMatrix(4, 3) = "无限制"
    
    End With
    
    '授权功能
    If blnTemp Then
        strSQL = "Select Distinct r.系统, 0 As 序号, -2 As 功能, r.系统 || '-' || u.名称 As 内容" & _
                " From zlRegFile r, zlSystems u, (Select Min(编号) As 编号 From zlSystems Group By Trunc(编号 / 100)) s" & _
                " Where r.系统 = Trunc(u.编号 / 100) And u.编号 = s.编号 And r.项目 = '授权功能' And r.功能 = '基本'" & _
                " Union All" & _
                " Select Distinct r.系统, r.序号, -1 As 功能, '" & Space(conIdent) & "' || r.序号 || '-' || p.标题 As 内容" & _
                " From zlRegFile r, zlPrograms p, (Select Min(编号) As 编号 From zlSystems Group By Trunc(编号 / 100)) s,zlRPTGroups g" & _
                " Where r.系统 = Trunc(p.系统 / 100) And r.序号 = p.序号 And p.系统 = s.编号 And r.项目 = '授权功能'" & _
                "   And p.系统=g.系统(+) And p.序号=g.程序ID(+) And (r.功能 = '基本' Or g.程序ID is Not Null)" & _
                " Union All" & _
                " Select r.系统, r.序号, Nvl(f.排列, 0) As 功能, '" & Space(conIdent * 2) & "' || f.功能 As 内容" & _
                " From zlRegFile r, zlProgfuncs f, (Select Min(编号) As 编号 From zlSystems Group By Trunc(编号 / 100)) s" & _
                " Where r.系统 = Trunc(f.系统 / 100) And r.序号 = f.序号 And r.功能 = f.功能 And f.系统 = s.编号 And r.项目 = '授权功能' And r.功能 <> '基本'" & _
                " Order By 系统, 序号, 功能"
    Else
        strSQL = "Select Distinct r.系统, 0 As 序号, -2 As 功能, r.系统 || '-' || u.名称 As 内容" & _
                " From zlRegFunc r, zlSystems u, (Select Min(编号) As 编号 From zlSystems Group By Trunc(编号 / 100)) s" & _
                " Where r.系统 = Trunc(u.编号 / 100) And u.编号 = s.编号 And r.功能 = '基本'" & _
                " Union All" & _
                " Select Distinct r.系统, r.序号, -1 As 功能, '" & Space(conIdent) & "' || r.序号 || '-' || p.标题 As 内容" & _
                " From zlRegFunc r, zlPrograms p, (Select Min(编号) As 编号 From zlSystems Group By Trunc(编号 / 100)) s,zlRPTGroups g" & _
                " Where r.系统 = Trunc(p.系统 / 100) And r.序号 = p.序号 And p.系统 = s.编号" & _
                "   And p.系统=g.系统(+) And p.序号=g.程序ID(+) And (r.功能 = '基本' Or g.程序ID is Not Null)" & _
                " Union All" & _
                " Select r.系统, r.序号, Nvl(f.排列, 0) As 功能, '" & Space(conIdent * 2) & "' || f.功能 As 内容" & _
                " From zlRegFunc r, zlProgfuncs f, (Select Min(编号) As 编号 From zlSystems Group By Trunc(编号 / 100)) s" & _
                " Where r.系统 = Trunc(f.系统 / 100) And r.序号 = f.序号 And r.功能 = f.功能 And f.系统 = s.编号 And r.功能 <> '基本'" & _
                " Order By 系统, 序号, 功能"
    End If
    If rsTemp.State = adStateOpen Then rsTemp.Close
    rsTemp.Open strSQL, gcnOracle, adOpenKeyset, adLockReadOnly
    
    With Me.vfgFunc
        .Clear
        Set .DataSource = rsTemp
        .ColWidth(0) = 0: .ColHidden(0) = True
        .ColWidth(1) = 0: .ColHidden(1) = True
        .ColWidth(2) = 0: .ColHidden(2) = True
    End With
    Me.optGrade(1).value = True
    Call optGrade_Click(1)
    Exit Sub
errHandle:
    MsgBox "刷新授权信息时出现错误，显示的授权信息可能不正确！" & vbNewLine & err.Description, vbExclamation, Me.Caption
End Sub

'--------------------------------------------------
'按管理工具规范，必须提供的函数
'--------------------------------------------------
Public Function SupportPrint() As Boolean
    '返回本窗口是否支持打印，供主窗口调用
    SupportPrint = True
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
    '供主窗口调用，实现具体的打印工作
    '如果没有可打印的，就留下一个空的接口
    
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    objPrint.Title.Text = "用户注册信息"
    
    objRow.Add vsInfo.TextMatrix(0, 0) & vsInfo.TextMatrix(0, 1)
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印时间：" & Format(date, "yyyy年MM月dd日")
    Set objPrint.Body = Me.vfgFunc
    objPrint.BelowAppRows.Add objRow
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub txtFind_Change()
    mintCount = -1
End Sub

Private Sub txtFind_GotFocus()
    If txtFind.Text <> "" And txtFind.ForeColor = vbGrayText Then
        txtFind.Text = ""
        txtFind.ForeColor = vbBlack
    Else
        txtFind.SelStart = 0
        txtFind.SelLength = Len(txtFind.Text)
    End If
End Sub

Private Sub txtFind_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim intRow As Integer
    Dim blnFindTag As Boolean
    
    If KeyCode = vbKeyReturn And txtFind.Text <> "" Then
        txtFind.Text = Replace(txtFind.Text, " ", "")
        With vfgFunc
            blnFindTag = False
            For intRow = mintCount + 1 To vfgFunc.Rows - 1
                If .RowHidden(intRow) = False And InStr(.TextMatrix(intRow, 3), txtFind.Text) > 0 Then blnFindTag = True: Exit For
            Next
            If blnFindTag Then .Row = intRow: .ShowCell intRow, 3: mintCount = intRow
            If intRow = .Rows Then
                If mintCount = -1 Then
                    Call MsgBox("未找到与“" & txtFind.Text & "”匹配的项目，请重新输入编号或关键字。", vbInformation, gstrSysName)
                    txtFind.Text = "": txtFind.SetFocus
                Else
                    mintCount = -1
                End If
            End If
        End With
    End If
End Sub

Private Sub txtFind_LostFocus()
    If txtFind.Text = "" Then
        txtFind.Text = "请输入编号或关键字"
        txtFind.ForeColor = vbGrayText
    End If
End Sub


