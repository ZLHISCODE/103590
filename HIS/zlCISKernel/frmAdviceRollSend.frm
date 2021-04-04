VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmAdviceRollSend 
   AutoRedraw      =   -1  'True
   Caption         =   "超期发送收回"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9540
   Icon            =   "frmAdviceRollSend.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   9540
   Begin VB.Frame fraSetup 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   9315
      Begin VB.Frame fraBaby 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   6120
         TabIndex        =   9
         Top             =   50
         Visible         =   0   'False
         Width           =   3195
         Begin VB.OptionButton optBaby 
            Caption         =   "婴儿医嘱"
            Height          =   180
            Index           =   2
            Left            =   2175
            TabIndex        =   12
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton optBaby 
            Caption         =   "所有医嘱"
            Height          =   180
            Index           =   0
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton optBaby 
            Caption         =   "病人医嘱"
            Height          =   180
            Index           =   1
            Left            =   1080
            TabIndex        =   10
            Top             =   0
            Width           =   1020
         End
      End
   End
   Begin VB.Frame fraInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   60
      TabIndex        =   5
      Top             =   525
      Width           =   9435
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C0FFFF&
         Height          =   180
         Left            =   120
         TabIndex        =   6
         Top             =   60
         Width           =   90
      End
   End
   Begin VB.TextBox txtPer 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   180
      Left            =   7290
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "100%"
      Top             =   6255
      Visible         =   0   'False
      Width           =   405
   End
   Begin MSComctlLib.ProgressBar psb 
      Height          =   270
      Left            =   2115
      TabIndex        =   1
      Top             =   6210
      Visible         =   0   'False
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   6150
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmAdviceRollSend.frx":058A
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13917
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   1376
            MinWidth        =   2
            Picture         =   "frmAdviceRollSend.frx":0E1E
            Text            =   "通过"
            TextSave        =   "通过"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Visible         =   0   'False
            Object.Width           =   1376
            MinWidth        =   2
            Picture         =   "frmAdviceRollSend.frx":1408
            Text            =   "疑问"
            TextSave        =   "疑问"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComCtl3.CoolBar cbr 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   714
      BandCount       =   1
      _CBWidth        =   9540
      _CBHeight       =   405
      _Version        =   "6.7.9782"
      Child1          =   "tbr"
      MinHeight1      =   345
      Width1          =   3525
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbr 
         Height          =   345
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   9420
         _ExtentX        =   16616
         _ExtentY        =   609
         ButtonWidth     =   1349
         ButtonHeight    =   609
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "全选"
               Key             =   "全选"
               Description     =   "全选"
               Object.ToolTipText     =   "全选(Ctrl+A)"
               Object.Tag             =   "全选"
               ImageKey        =   "全选"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "全清"
               Key             =   "全清"
               Description     =   "全清"
               Object.ToolTipText     =   "全清(Ctrl+R)"
               Object.Tag             =   "全清"
               ImageKey        =   "全清"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "收回"
               Key             =   "收回"
               Description     =   "收回"
               Object.ToolTipText     =   "超期收回选择的医嘱(Ctrl+E)"
               Object.Tag             =   "收回"
               ImageKey        =   "收回"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "重置"
               Key             =   "重置"
               Description     =   "重置"
               Object.ToolTipText     =   "重新设置条件并产生发送清单(F12)"
               Object.Tag             =   "重置"
               ImageKey        =   "重置"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Description     =   "帮助"
               Object.ToolTipText     =   "帮助(F1)"
               Object.Tag             =   "帮助"
               ImageKey        =   "帮助"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Description     =   "退出"
               Object.ToolTipText     =   "退出(ALT+X)"
               Object.Tag             =   "退出"
               ImageKey        =   "退出"
            EndProperty
         EndProperty
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   4935
      Left            =   0
      TabIndex        =   7
      Top             =   1185
      Width           =   9540
      _cx             =   16828
      _cy             =   8705
      Appearance      =   1
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16771802
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAdviceRollSend.frx":19F2
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
      OwnerDraw       =   1
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
End
Attribute VB_Name = "frmAdviceRollSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mMainPrivs As String 'IN
Private mlng病区ID As Long 'IN:用于记录主界面的病区及上次发送病区
Private mlng病人ID As Long 'IN
Private mlng主页ID As Long 'IN
Private mblnAutoRoll As Boolean 'In,停止后自动收回
Private mblnOnePati As Boolean  '新护士站调用(单病人模式)，或停止确认操作时调用超期收回

Private mblnRoll As Boolean 'OUT:是否成功收回过。
Private mblnAdjustNum As Boolean '是否有超期收回调整的权限（调整数量）
Private mbln只显示当前病区医嘱  As Boolean

Private mrsBill As ADODB.Recordset
Private mbln超期负数 As Boolean
Private mblnFirst As Boolean
Private mblnReturn As Boolean
Private mstr科室IDs As String   '当前病区对应的科室IDs+当前病区ID
Private mint医嘱处理范围 As Integer    '医嘱处理范围   0-所有医嘱,1-病人医嘱,2-婴儿医嘱
Private mblnFirstLoad As Boolean
Private mlng医护科室ID As Long
Private mlng婴儿病区ID As Long
Private mblnLimit As Boolean '本次发送给药途径计算是否以结束时间限制
Private mbln销帐申请 As Boolean '非打包的输液单在配液之后是否可以进行销帐申请

Private Const COL_选择 = 0
Private Const COL_科室 = 1
Private Const COL_姓名 = 2
Private Const COL_住院号 = 3
Private Const COL_床号 = 4
Private Const COL_婴儿 = 5
Private Const col_医嘱内容 = 6
Private Const COL_规格 = 7
Private Const COL_总量 = 8
Private Const COL_单位 = 9
Private Const COL_频率 = 10
Private Const COL_用法 = 11
Private Const COL_执行时间 = 12
Private Const COL_上次执行 = 13
Private Const COL_终止时间 = 14
Private Const COL_执行科室 = 15
Private Const COL_病人ID = 16
Private Const COL_主页ID = 17
Private Const COL_险类 = 18
Private Const COL_ID = 19
Private Const COL_相关ID = 20
Private Const COL_诊疗类别 = 21
Private Const COL_药品ID = 22
Private Const COL_病人科室ID = 23
Private Const COL_开嘱科室ID = 24
Private Const COL_开嘱医生 = 25
Private Const COL_执行科室ID = 26
Private Const COL_次数 = 27
Private Const COL_计算量 = 28
Private Const COL_单量 = 29
Private Const COL_剂量系数 = 30
Private Const COL_住院包装 = 31
Private Const COL_可否分零 = 32
Private Const COL_执行性质 = 33
Private Const COL_上次 = 34 '收回后应该的上次执行时间
Private Const COL_操作类型 = 35 '输液药品医嘱的判定
Private Const COL_执行分类 = 36
Private Const COL_病人性质 = 37

Private Property Let Progress(ByVal vNewValue As Single)
'vNewValue=0-100
    If vNewValue = 0 Then
        psb.value = 0: txtPer.Text = ""
        psb.Visible = False: txtPer.Visible = False
    Else
        psb.value = vNewValue
        txtPer.Text = CInt(psb.value) & "%"
        psb.Visible = True: txtPer.Visible = True
        txtPer.Refresh
    End If
End Property


Public Function ShowMe(frmParent As Object, ByVal MainPrivs As String, _
    ByVal lng病区ID As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
     ByVal blnOnePati As Boolean, ByVal blnAutoRoll As Boolean, Optional ByVal lng医护科室ID As Long, Optional ByVal lng婴儿病区ID As Long) As Boolean
'参数：
'       blnOnePati=单病人模式
    mMainPrivs = MainPrivs
    
    mlng病人ID = lng病人ID
    mlng病区ID = lng病区ID
    mlng主页ID = lng主页ID
    mlng医护科室ID = lng医护科室ID
    mlng婴儿病区ID = lng婴儿病区ID
    mblnOnePati = blnOnePati
    mblnAutoRoll = blnAutoRoll
        
    Me.Show 1, frmParent
    ShowMe = mblnRoll
End Function

Private Sub Form_Activate()
    Dim blnAutoRoll As Boolean
    
    If mblnFirst Then
        mblnFirst = False
        '单病人模式
        If mblnOnePati Then
            Call LoadAdviceRoll(mlng病人ID, mlng主页ID)
            tbr.Buttons("重置").Visible = False
                    
            If mblnAutoRoll Then
                Call tbr_ButtonClick(tbr.Buttons("收回"))
            End If
        Else
            If Not ResetSend Then Unload Me: Exit Sub
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        Call tbr_ButtonClick(tbr.Buttons("全选"))
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        Call tbr_ButtonClick(tbr.Buttons("全清"))
    ElseIf KeyCode = vbKeyE And Shift = vbCtrlMask Then
        Call tbr_ButtonClick(tbr.Buttons("收回"))
    ElseIf KeyCode = vbKeyF12 And Shift = 0 Then
        Call tbr_ButtonClick(tbr.Buttons("重置"))
    ElseIf KeyCode = vbKeyF1 And Shift = 0 Then
        Call tbr_ButtonClick(tbr.Buttons("帮助"))
    ElseIf KeyCode = vbKeyX And Shift = vbAltMask Then
        Call tbr_ButtonClick(tbr.Buttons("退出"))
    End If
End Sub

Private Sub Form_Load()

    '设置公共按钮图标
    Set tbr.HotImageList = frmIcons.imgColor
    Set tbr.ImageList = frmIcons.imgGray
    tbr.Buttons("全选").Image = "全选"
    tbr.Buttons("全清").Image = "全清"
    tbr.Buttons("收回").Image = "执行"
    tbr.Buttons("重置").Image = "重置"
    tbr.Buttons("帮助").Image = "帮助"
    tbr.Buttons("退出").Image = "退出"
    tbr.ButtonHeight = 500
    mblnFirstLoad = True
        
    Call InitAdviceTable
    Call RestoreWinState(Me, App.ProductName)
    
    mblnRoll = False
    mblnFirst = True
    mbln超期负数 = Val(zlDatabase.GetPara("超期收回产生负数费用", glngSys, p住院医嘱发送)) = 1
    mblnAdjustNum = InStr(GetInsidePrivs(p住院医嘱发送), "超期收回调整") > 0
    mbln只显示当前病区医嘱 = Val(zlDatabase.GetPara("只显示当前病区的医嘱", glngSys, p住院医嘱发送, "0")) = 1
    mblnLimit = Val(zlDatabase.GetPara("药嘱发送限制结束时间", glngSys, p住院医嘱发送, 0)) = 1
    mstr科室IDs = Get科室IDs(mlng病区ID)
    mbln销帐申请 = Val(zlDatabase.GetPara("配液输液单配药后允许销帐申请", glngSys, 1345, 0)) = 1
    
End Sub

Private Sub InitBillSet()
'功能：初始化医嘱记帐单据生成记录集
    Set mrsBill = New ADODB.Recordset
    
    mrsBill.Fields.Append "Key", adVarChar, 100
    mrsBill.Fields.Append "NO", adVarChar, 8
    mrsBill.CursorLocation = adUseClient
    mrsBill.LockType = adLockOptimistic
    mrsBill.CursorType = adOpenStatic
    mrsBill.Open
End Sub

Private Sub Form_Resize()
    Dim lngW As Long
    Dim i As Long
    
    On Error Resume Next
    
    fraInfo.Top = cbr.Height
    fraInfo.Left = 0
    fraInfo.Width = Me.ScaleWidth
    
    fraSetup.Top = fraInfo.Top + fraInfo.Height
    fraSetup.Left = 0
    fraSetup.Width = Me.ScaleWidth
    
    fraBaby.Left = fraSetup.Width - fraBaby.Width
    
    vsAdvice.Left = 0
    vsAdvice.Top = IIF(fraSetup.Visible, fraSetup.Top + fraSetup.Height, fraInfo.Top + fraInfo.Height)
    vsAdvice.Width = Me.ScaleWidth
    vsAdvice.Height = Me.ScaleHeight - fraInfo.Height - cbr.Height - stbThis.Height
    
    psb.Top = Me.ScaleHeight - stbThis.Height + 60
    psb.Left = stbThis.Panels(1).Width + 90
    
    For i = 1 To stbThis.Panels.Count
        If i <> 2 And stbThis.Panels(i).Visible Then
            lngW = lngW + (stbThis.Panels(i).Width + 60)
        End If
    Next
    psb.Width = Me.ScaleWidth - lngW - txtPer.Width - 500
    
    txtPer.Left = psb.Left + psb.Width
    txtPer.Top = psb.Top + (psb.Height - txtPer.Height) / 2
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    mMainPrivs = ""
    mlng病区ID = 0
    mlng病人ID = 0
    mstr科室IDs = ""
    mlng主页ID = 0
    mblnLimit = False
    Set mrsBill = Nothing
End Sub

Private Sub InitAdviceTable()
'功能：初始化清单格式
    Dim arrHead As Variant, strHead As String, i As Long
    
    strHead = ",300,4;科室,850,1;姓名,750,1;住院号,750,1;床号,500,4;婴儿,550,1;" & _
        "医嘱内容,2000,1;规格,2000,1;收回量,700,7;单位,450,1;频率,1000,1;用法,1000,1;" & _
        "执行时间,1000,1;上次执行,1530,1;终止时间,1530,1;执行科室,850,1;" & _
        "病人ID;主页ID;险类;ID;相关ID;诊疗类别;药品ID;病人科室ID;开嘱科室ID;开嘱医生;执行科室ID;" & _
        "次数;计算量;单量;剂量系数;住院包装;可否分零;执行性质;上次;操作类型;执行分类;病人性质"
    arrHead = Split(strHead, ";")
    With vsAdvice
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = 4
        .FrozenCols = COL_选择 + 1 - .FixedCols
        .ColDataType(COL_选择) = flexDTBoolean
        .RowHeight(0) = 320
    End With
End Sub

Private Function ResetSend() As Boolean
'功能：重置发送条件
    With frmAdviceRollSendCond
        .mMainPrivs = mMainPrivs
        .mlng病区ID = mlng病区ID
        If mlng婴儿病区ID <> 0 Then
            If mlng婴儿病区ID = mlng医护科室ID Then
                .mlng病区ID = mlng婴儿病区ID
            End If
        End If
        .mlng病人ID = mlng病人ID
        .Show 1, Me
        If .mblnOK Then
            mlng病区ID = .mlng病区ID
            mstr科室IDs = Get科室IDs(mlng病区ID)
            mlng医护科室ID = mlng病区ID
            Call LoadAdviceRoll(.mstr病人IDs, .mstr主页IDs)
        End If
        ResetSend = .mblnOK
    End With
End Function

Private Sub optBaby_Click(Index As Integer)
    mint医嘱处理范围 = Index
    '单病人模式
    If Not mblnFirstLoad Then
        If mblnOnePati Then
            Call LoadAdviceRoll(mlng病人ID, mlng主页ID)
        Else
            Call LoadAdviceRoll(frmAdviceRollSendCond.mstr病人IDs, frmAdviceRollSendCond.mstr主页IDs)
        End If
    End If
End Sub

Private Sub tbr_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim str医嘱IDs As String, i As Long, strMsg As String, strMsgAll As String
    
    Select Case Button.Key
        Case "全选"
            With vsAdvice
                For i = .FixedRows To .Rows - 1
                    If .RowHidden(i) = False And Val(.TextMatrix(i, COL_选择)) = 0 Then
                        If RowCanRoll(i, strMsg) Then
                            .TextMatrix(i, COL_选择) = 1
                            Call RowSelectSame(i)
                        Else
                            strMsgAll = strMsgAll & vbCrLf & .TextMatrix(i, col_医嘱内容) & ":" & strMsg
                        End If
                    End If
                Next
                If strMsgAll <> "" Then
                    MsgBox "以下医嘱不能超期收回：" & strMsgAll, vbInformation, gstrSysName
                End If
            End With
        Case "全清"
            vsAdvice.Cell(flexcpText, vsAdvice.FixedRows, COL_选择, vsAdvice.Rows - 1, COL_选择) = 0
        Case "收回"
            With vsAdvice
                For i = .FixedRows To .Rows - 1
                    If Val(.TextMatrix(i, COL_选择)) <> 0 And Val(.TextMatrix(i, COL_ID)) <> 0 Then
                        str医嘱IDs = str医嘱IDs & "," & Val(.TextMatrix(i, COL_ID))
                    End If
                Next
                If str医嘱IDs = "" Then
                    MsgBox "请至少选择一行要收回的医嘱。", vbInformation, gstrSysName
                    Exit Sub
                Else
                    str医嘱IDs = Mid(str医嘱IDs, 2)
                    
                    '对要收回医嘱的发送费用进行结帐检查
                    If Not CheckRollMoneyBalance(str医嘱IDs) Then Exit Sub
                    
                    '检查并提示，收费对照为一天只收一次，或一次发送只收一次等的收费项目
                    Call CheckRollPriceItem(str医嘱IDs)
                End If
            End With
            If mblnAutoRoll Then
                If RollAdvice(UBound(Split(str医嘱IDs, ",")) + 1) Then mblnRoll = True: Unload Me
            Else
                If MsgBox("确实要对当前选择的医嘱执行收回操作吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    If RollAdvice(UBound(Split(str医嘱IDs, ",")) + 1) Then mblnRoll = True: Unload Me
                End If
            End If
        Case "重置"
            Call ResetSend
        Case "帮助"
            ShowHelp App.ProductName, Me.hwnd, Me.Name
        Case "退出"
            Unload Me
    End Select
End Sub

Private Sub CheckRollPriceItem(ByVal str医嘱IDs As String)
'功能：检查并提示，收费对照为一天只收一次，或一次发送只收一次等的收费项目如果存在  医嘱执行计价数据，则允许收回，否则禁止
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strTmp As String, i As Long
       
    strSQL = "Select  /*+ rule*/Distinct c.名称 as 收费名称,e.名称 as 诊疗名称" & vbNewLine & _
        "From 病人医嘱计价 A,Table(f_Num2list([1])) B,收费项目目录 C,病人医嘱记录 D,诊疗项目目录 E" & vbNewLine & _
        "Where a.医嘱id = b.Column_Value And Nvl(a.收费方式, 0) <> 0 and a.收费细目id=c.id And a.医嘱id = d.id And d.诊疗项目id = e.id"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "超期收回", str医嘱IDs)
    For i = 1 To rsTmp.RecordCount
        strTmp = strTmp & vbCrLf & rsTmp!诊疗名称 & "：" & rsTmp!收费名称
        If i > 9 Then
            strTmp = strTmp & "......"
            Exit For
        End If
        rsTmp.MoveNext
    Next
    If strTmp <> "" Then
        strSQL = "select Column_Value from Table(f_Num2list([1]))" & vbNewLine & _
            "minus" & vbNewLine & _
            "Select 医嘱id From 医嘱执行计价 Where 医嘱id In (select Column_Value from Table(f_Num2list([1]))) Group By 医嘱id"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "超期收回", str医嘱IDs)
        If rsTmp.RecordCount <> 0 Then
            MsgBox "检查发现要收回医嘱的费用存在以下一天只收一次或一次发送只收一次的项目：" & vbCrLf & _
                strTmp & vbCrLf & "由于无法明确收回数量，它们将不会被自动收回，你可以使用销帐申请来收回。", vbInformation, gstrSysName
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function CheckRollMoneyBalance(ByVal str医嘱IDs As String) As Boolean
'功能：对要收回医嘱的发送费用进行结帐检查
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    CheckRollMoneyBalance = True
    If gbytBillOpt = 0 Then Exit Function
    
    '取医嘱最近发送的记帐NO
    strSQL = "Select Column_Value From Table(f_Num2list([1]))"
    strSQL = _
        " Select B.姓名,A.医嘱ID,Decode(Instr(',4,5,6,7,',B.诊疗类别),0,C.名称,B.医嘱内容) as 医嘱内容,Max(A.NO) as NO" & _
        " From 病人医嘱发送 A,病人医嘱记录 B,诊疗项目目录 C,(" & strSQL & ") X" & _
        " Where A.医嘱ID=B.ID And B.诊疗项目ID=C.ID And A.记录性质=2 And A.医嘱ID=X.Column_Value" & _
        " Group by B.姓名,A.医嘱ID,Decode(Instr(',4,5,6,7,',B.诊疗类别),0,C.名称,B.医嘱内容)"
    
    '取这些NO的结帐情况(非划价未销帐)
    strSQL = "Select B.姓名,B.医嘱ID,B.医嘱内容,A.NO,Nvl(A.价格父号,A.序号) as 序号,Sum(Nvl(A.结帐金额,0)) as 结帐金额" & _
        " From 住院费用记录 A,(" & strSQL & ") B" & _
        " Where A.NO=B.NO And A.医嘱序号=B.医嘱ID And A.记录性质 IN(2,12) And A.记录状态=1" & _
        " Group by B.姓名,B.医嘱ID,B.医嘱内容,A.NO,Nvl(A.价格父号,A.序号) Having Sum(Nvl(A.结帐金额,0))<>0"
    strSQL = "Select /*+ Rule*/ 姓名,医嘱ID,医嘱内容 From (" & strSQL & ") Group by 姓名,医嘱ID,医嘱内容"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "超期收回", str医嘱IDs)
    
    strSQL = ""
    Do While Not rsTmp.EOF
        If UBound(Split(strSQL, vbCrLf)) > 10 Then
            strSQL = strSQL & vbCrLf & "… …"
            Exit Do
        Else
            strSQL = strSQL & vbCrLf & "●" & rsTmp!姓名 & "：" & rsTmp!医嘱内容
        End If
        rsTmp.MoveNext
    Loop
    
    If strSQL <> "" Then
        If gbytBillOpt = 1 Then
            If MsgBox("要收回的下列医嘱的最近发送费用存在已结帐的情况：" & vbCrLf & strSQL & vbCrLf & vbCrLf & "确实要执行收回操作吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                CheckRollMoneyBalance = False
            End If
        ElseIf gbytBillOpt = 2 Then
            MsgBox "要收回的下列医嘱的最近发送费用存在已结帐的情况：" & vbCrLf & strSQL & vbCrLf & vbCrLf & "不能执行收回操作。", vbInformation, gstrSysName
            CheckRollMoneyBalance = False
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub RowSelectSame(ByVal lngRow As Long, Optional lngBegin As Long, Optional lngEnd As Long)
'功能：根据可见行的选择状态,将相关医嘱一并选择
    Dim lngS组ID As Long, lngO组ID As Long, i As Long
    
    With vsAdvice
        lngBegin = lngRow: lngEnd = lngRow
        lngS组ID = IIF(Val(.TextMatrix(lngRow, COL_相关ID)) <> 0, Val(.TextMatrix(lngRow, COL_相关ID)), Val(.TextMatrix(lngRow, COL_ID)))
        For i = lngRow + 1 To .Rows - 1
            lngO组ID = IIF(Val(.TextMatrix(i, COL_相关ID)) <> 0, Val(.TextMatrix(i, COL_相关ID)), Val(.TextMatrix(i, COL_ID)))
            If lngO组ID = lngS组ID Then
                .TextMatrix(i, COL_选择) = .TextMatrix(lngRow, COL_选择)
                lngEnd = i
            Else
                Exit For
            End If
        Next
        For i = lngRow - 1 To .FixedRows Step -1
            lngO组ID = IIF(Val(.TextMatrix(i, COL_相关ID)) <> 0, Val(.TextMatrix(i, COL_相关ID)), Val(.TextMatrix(i, COL_ID)))
            If lngO组ID = lngS组ID Then
                .TextMatrix(i, COL_选择) = .TextMatrix(lngRow, COL_选择)
                lngBegin = i
            Else
                Exit For
            End If
        Next
    End With
End Sub

Private Sub GetGroupRow(ByVal lngRow As Long, Optional lngBegin As Long, Optional lngEnd As Long)
'功能：根据当前医嘱行返回一组医嘱的行范围
    Dim lngS组ID As Long, lngO组ID As Long, i As Long
    
    With vsAdvice
        lngBegin = lngRow: lngEnd = lngRow
        lngS组ID = IIF(Val(.TextMatrix(lngRow, COL_相关ID)) <> 0, Val(.TextMatrix(lngRow, COL_相关ID)), Val(.TextMatrix(lngRow, COL_ID)))
        For i = lngRow + 1 To .Rows - 1
            lngO组ID = IIF(Val(.TextMatrix(i, COL_相关ID)) <> 0, Val(.TextMatrix(i, COL_相关ID)), Val(.TextMatrix(i, COL_ID)))
            If lngO组ID = lngS组ID Then
                lngEnd = i
            Else
                Exit For
            End If
        Next
        For i = lngRow - 1 To .FixedRows Step -1
            lngO组ID = IIF(Val(.TextMatrix(i, COL_相关ID)) <> 0, Val(.TextMatrix(i, COL_相关ID)), Val(.TextMatrix(i, COL_ID)))
            If lngO组ID = lngS组ID Then
                lngBegin = i
            Else
                Exit For
            End If
        Next
    End With
End Sub

Private Sub vsAdvice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = COL_选择 Then Call RowSelectSame(Row)
End Sub

Private Sub vsAdvice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewCol = COL_总量 Then
        If Not CellEditable(NewRow, NewCol) Then
            vsAdvice.FocusRect = flexFocusLight
        Else
            vsAdvice.FocusRect = flexFocusHeavy
        End If
    Else
        vsAdvice.FocusRect = flexFocusLight
    End If
End Sub

Private Sub vsAdvice_AfterUserFreeze()
    With vsAdvice
        If .FrozenCols < COL_选择 + 1 - .FixedCols Then
            .FrozenCols = COL_选择 + 1 - .FixedCols
        End If
    End With
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    With vsAdvice
        If Col = col_医嘱内容 Or Col = COL_规格 Then
            If Not .ColHidden(COL_规格) Then
                .AutoSize col_医嘱内容, COL_规格
            Else
                .AutoSize col_医嘱内容
            End If
            .RowHeight(0) = 320
        ElseIf Row = -1 Then
            lngW = Me.TextWidth(.TextMatrix(.FixedRows - 1, Col) & "A")
            If .ColWidth(Col) < lngW Then
                .ColWidth(Col) = lngW
            ElseIf .ColWidth(Col) > .Width * 0.5 Then
                .ColWidth(Col) = .Width * 0.5
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = COL_选择 Then Cancel = True
End Sub

Private Function RowIn一并给药(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'功能：判断指定行是否在一并给药的范围中,如果是,同时返回行号范围
'说明：返回的行号范围不包括给药途径的行号
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If .TextMatrix(lngRow, COL_诊疗类别) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, COL_诊疗类别)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_相关ID)) = Val(.TextMatrix(lngRow, COL_相关ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowIn一并给药 = blnTmp
    End With
End Function

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'说明：1.OwnerDraw要设置为Over(画出单元所有内容)
'      2.Cell的GridLine从上下左右向内都是从第1根线开始
'      3.Cell的Border从左上是从第2根线开始,右下是从第1根线开始
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        lngLeft = COL_频率: lngRight = COL_用法
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = COL_科室: lngRight = COL_婴儿
        End If
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        
        If Not RowIn一并给药(Row, lngBegin, lngEnd) Then Exit Sub
        
        vRect.Left = Left '擦除左边表格线
        vRect.Right = Right - 1 '保留右边表格线
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '首行保留文字内容
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 2 '底行保留下边线(本窗体中用到下边线粗为2)
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, OS.SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    Dim i As Long
    With vsAdvice
        If KeyAscii = 13 Then
            KeyAscii = 0
            For i = .Row + 1 To .Rows - 1
                If Not .RowHidden(i) Then
                    .Row = i: Exit For
                End If
            Next
            If i > .Rows - 1 And Not .RowHidden(.FixedRows) Then .Row = .FixedRows
            Call .ShowCell(.Row, .Col)
        End If
    End With
End Sub

Private Function AcceptInput(ByVal Row As Long, ByVal Col As Long) As Boolean
'功能：检查并接收输入总量
    Dim lngBegin As Long, lngEnd As Long, i As Long
    Dim dblOnce As Double, dblModify As Double
    Dim lng次数 As Long, lngMin次数 As Long
    
    AcceptInput = False
    With vsAdvice
        If Val(.EditText) = Val(.TextMatrix(Row, Col)) Then AcceptInput = True: Exit Function
        
        '检查输入有效性
        If Val(.TextMatrix(Row, COL_相关ID)) <> 0 And InStr(",5,6,", "," & .TextMatrix(Row, COL_诊疗类别) & ",") > 0 Then
            If CheckAdvcieComPound(Val(.TextMatrix(Row, COL_相关ID))) Then
                MsgBox "输液配药的记录不允许修改收回量。", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
            End If
        End If
        If Not IsNumeric(.EditText) Or Val(.EditText) < 0 Or Val(.EditText) > LONG_MAX Then
            MsgBox "输入错误，不是大于等于零的数字或输入数值过大！", vbInformation, gstrSysName
            .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
        End If
        If Val(.EditText) > Val(.TextMatrix(Row, COL_计算量)) Then
            MsgBox "收回量不能大于 " & .TextMatrix(Row, COL_计算量) & .TextMatrix(Row, COL_单位) & "。", vbInformation, gstrSysName
            .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
        End If
        If .TextMatrix(Row, COL_诊疗类别) = "E" And Val(.TextMatrix(Row, COL_ID)) = Val(.TextMatrix(Row - 1, COL_相关ID)) _
            And InStr(",E,7,", .TextMatrix(Row - 1, COL_诊疗类别)) > 0 Then
            If Val(.EditText) <> Int(.EditText) Then
                MsgBox "中药配方收回付数应为整数。", vbInformation, gstrSysName
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col)): Exit Function
            End If
        End If
        
        '接收当前输入值
        .EditText = FormatEx(.EditText, 5)
        If InStr(",5,6,", .TextMatrix(Row, COL_诊疗类别)) > 0 Then
            '药品要管分零特性及给药次数
            If Val(.TextMatrix(Row, COL_可否分零)) = 0 Then
                '可分零
            ElseIf Val(.TextMatrix(Row, COL_可否分零)) = 1 Or Val(.TextMatrix(Row, COL_可否分零)) < 0 Then
                '不分零:整数住院包装,按比输入的收回值少处理
                .EditText = Int(Val(.EditText))
            ElseIf Val(.TextMatrix(Row, COL_可否分零)) = 2 Then
                '一次性:足单次用量的整数住院包装,按比输入的收回值少处理
                dblOnce = IntEx(Val(.TextMatrix(Row, COL_单量)) / Val(.TextMatrix(Row, COL_剂量系数)) / Val(.TextMatrix(Row, COL_住院包装)))
                .EditText = Int(Val(.EditText) / dblOnce) * dblOnce
            End If
        End If
        .TextMatrix(Row, Col) = .EditText
        .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
        .Cell(flexcpFontBold, Row, Col) = Val(.TextMatrix(Row, Col)) <> Val(.TextMatrix(Row, COL_计算量)) '标记为修改过
        If Val(.TextMatrix(Row, Col)) = 0 Then
            .TextMatrix(Row, COL_选择) = 0
        Else
            If RowCanRoll(Row) Then
                .TextMatrix(Row, COL_选择) = 1
            Else
                .TextMatrix(Row, COL_选择) = 0
            End If
        End If
        Call RowSelectSame(Row, lngBegin, lngEnd)
        
        '计算相关输入值
        If InStr(",5,6,", .TextMatrix(Row, COL_诊疗类别)) > 0 Then
            '给药途径
            lngMin次数 = LONG_MAX
            For i = lngBegin To lngEnd
                If InStr(",5,6,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                    If Val(.TextMatrix(i, COL_总量)) = Val(.TextMatrix(i, COL_计算量)) Then
                        lng次数 = 0 '未变动的,恢复原次数
                    Else
                        '求本次修改少收回的量可执行的次数,一并给药以最小的为准
                        dblModify = Val(.TextMatrix(i, COL_计算量)) - Val(.TextMatrix(i, COL_总量))
                        If Val(.TextMatrix(i, COL_可否分零)) = 0 Then
                            '可分零,按比输入的收回值少处理
                            lng次数 = Int(dblModify * Val(.TextMatrix(i, COL_住院包装)) * Val(.TextMatrix(i, COL_剂量系数)) / Val(.TextMatrix(i, COL_单量)))
                        ElseIf Val(.TextMatrix(i, COL_可否分零)) = 1 Or Val(.TextMatrix(i, COL_可否分零)) < 0 Then
                            '不分零:整数住院包装,反过来根据实际收回量计算可用次数
                            lng次数 = IntEx(Val(.TextMatrix(i, COL_总量)) * Val(.TextMatrix(i, COL_住院包装)) * Val(.TextMatrix(i, COL_剂量系数)) / Val(.TextMatrix(i, COL_单量)))
                            lng次数 = Val(.TextMatrix(i, COL_次数)) - lng次数
                        ElseIf Val(.TextMatrix(i, COL_可否分零)) = 2 Then
                            '一次性:足单次用量的整数住院包装,按比输入的收回值少处理
                            lng次数 = Int(dblModify / IntEx(Val(.TextMatrix(i, COL_单量)) / Val(.TextMatrix(i, COL_剂量系数)) / Val(.TextMatrix(i, COL_住院包装))))
                        End If
                    End If
                    If lng次数 < 0 Then lng次数 = 0
                    If lng次数 < lngMin次数 Then lngMin次数 = lng次数
                ElseIf .TextMatrix(i, COL_诊疗类别) = "E" Then
                    If lngMin次数 <> LONG_MAX Then
                        If Val(.TextMatrix(i, COL_次数)) - lngMin次数 >= 0 Then
                            .TextMatrix(i, COL_总量) = Val(.TextMatrix(i, COL_次数)) - lngMin次数
                            .Cell(flexcpData, i, COL_总量) = .TextMatrix(i, COL_总量)
                        End If
                    End If
                End If
            Next
        Else
            '中药配方，以及非药品相关:同步与当前行输入相同
            For i = lngBegin To lngEnd
                If i <> Row Then
                    .TextMatrix(i, COL_总量) = .TextMatrix(Row, COL_总量)
                    .Cell(flexcpData, i, COL_总量) = .TextMatrix(i, COL_总量)
                End If
            Next
        End If
    End With
    AcceptInput = True
End Function

Private Sub vsAdvice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsAdvice
        If mblnReturn Then mblnReturn = False
        If KeyAscii = 13 Then
            If Col = COL_总量 Then
                KeyAscii = 0
                mblnReturn = True
                If Not AcceptInput(Row, Col) Then Exit Sub
                '定位到一下行
                Call vsAdvice.FinishEditing(False)
                Call vsAdvice_KeyPress(13)
            End If
        Else
            If Col = COL_总量 Then
                If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0: Exit Sub
                End If
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsAdvice.EditSelStart = 0
    vsAdvice.EditSelLength = zlCommFun.ActualLen(vsAdvice.EditText)
End Sub

Private Function CellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    CellEditable = True
    
    If lngCol = COL_总量 Then
        If Not mblnAdjustNum Then
            CellEditable = False
        End If
    ElseIf lngCol <> COL_选择 Then
        CellEditable = False
    End If
End Function

Private Sub vsAdvice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strMsg As String
    
    If Not CellEditable(Row, Col) Then
        Cancel = True
    Else
        If Col = COL_总量 Then
            vsAdvice.EditMaxLength = 10
        Else
            vsAdvice.EditMaxLength = 0
        End If
        
        If Col = COL_选择 And Val(vsAdvice.TextMatrix(Row, Col)) = 0 Then
            If Not RowCanRoll(Row, strMsg) Then
                Cancel = True
                MsgBox strMsg, vbInformation, gstrSysName
            End If
        End If
    End If
End Sub

Private Function LoadAdviceRoll(ByVal str病人IDs As String, ByVal str主页IDs As String) As Boolean
'功能：读取指定病人的超期发送医嘱清单,包含药品及非药医嘱
'参数：str病人IDs=包含病人ID的字符串
    Dim rsAdvice As New ADODB.Recordset
    Dim rsDrug As New ADODB.Recordset
    Dim rsSend As New ADODB.Recordset
    Dim strSQL As String, str持续性 As String, lng药品ID As Long
    Dim str科室 As String, lng病人数 As Long, lng病人ID As Long
    Dim strPause As String, lng次数 As Long, dbl总量 As Double, dbl总量All As Double
    Dim arr分解时间 As Variant, str分解时间 As String, str上次时间 As String
    Dim datBegin As Date, lngRow As Long, i As Long, j As Long, k As Long
    Dim lngDel组ID As Long
    Dim int可否分零 As Integer, strUnRoll As String
    
    Screen.MousePointer = 11
    lblInfo.Caption = "正在读取数据...."
    
    vsAdvice.Rows = vsAdvice.FixedRows
    vsAdvice.ColHidden(COL_规格) = True
    vsAdvice.ColHidden(COL_科室) = True
    vsAdvice.ColHidden(COL_婴儿) = True
    Me.Refresh
    
    If DeptIsWoman(0, Get科室IDs(mlng病区ID)) Then
        If mblnFirstLoad Then
            fraSetup.Visible = True
            fraBaby.Visible = True
            '医嘱处理范围
            mint医嘱处理范围 = Val(zlDatabase.GetPara("医嘱处理范围", glngSys, p住院医嘱发送, "0"))
            optBaby(mint医嘱处理范围).value = True
            mblnFirstLoad = False
        End If
    Else
        mblnFirstLoad = True
        fraSetup.Visible = False
        optBaby(0).value = True
    End If
    Call Form_Resize
    
    strUnRoll = zlDatabase.GetPara("发药后不收回", glngSys, p住院医嘱发送)
    
    '不包含护理等级,术前术后医嘱和叮嘱这种不发送的医嘱(给药途径,配方煎法,用法也可能为叮嘱)
    '应该不包含检查手术(临嘱)
    '注意"持续性"长嘱终止这天不发送
    '中药用法即使叮嘱也固定要读出来(以输入收回付数)
    str持续性 = "(A.执行时间方案 is NULL And (Nvl(A.频率次数,0)=0 Or Nvl(A.频率间隔,0)=0 Or A.频率间隔 is NULL))"
    
    For k = 0 To UBound(Split(str病人IDs, ","))
        strSQL = "Select A.ID,A.相关ID,Nvl(A.相关ID,A.ID) as 组ID,Nvl(X.序号,A.序号) as 组号," & _
            " D.名称 as 科室,A.病人ID,A.主页ID,B.险类,A.收费细目ID,A.姓名,B.住院号,B.出院病床 as 床号,B.入院日期," & _
            " A.婴儿,A.医嘱内容,A.诊疗类别,A.诊疗项目ID,A.病人科室ID,A.开嘱科室ID,A.开嘱医生,A.总给予量,A.单次用量," & _
            " A.执行频次 as 频率,E.计算单位,E.名称 as 诊疗项目,Nvl(F.名称,Decode(Nvl(A.执行性质,0),5,'-')) as 执行科室,A.执行科室ID,A.执行性质," & _
            " A.开始执行时间,A.执行时间方案,A.上次执行时间,A.执行终止时间,A.频率次数,A.频率间隔,A.间隔单位," & _
            " A.可否分零,Decode(Instr(',5,6,',A.诊疗类别),0,NULL,G.名称) as 给药途径,A.首次用量,e.操作类型,e.执行分类,b.病人性质" & _
            " From 病人医嘱记录 A,病人医嘱记录 X,病案主页 B,病人信息 C,部门表 D,诊疗项目目录 E,部门表 F,诊疗项目目录 G" & _
            " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID" & _
            " And A.病人ID=C.病人ID And B.出院科室ID=D.ID And A.诊疗项目ID=E.ID" & _
            " And A.执行科室ID=F.ID(+) And A.病人ID=[1] And A.主页ID=[2]" & _
            " And (Nvl(A.执行性质,0)<>0 Or A.诊疗类别='E' And E.操作类型='4')" & _
            " And Not(A.诊疗类别='H' And E.操作类型='1' And E.执行频率=2) And Not(A.诊疗类别='Z' And E.操作类型 In('4','14'))" & _
            " And Nvl(A.医嘱期效,0)=0 And A.相关ID=X.ID(+) And X.诊疗项目ID = G.ID(+)" & _
            " And ((Not " & str持续性 & " And A.执行终止时间<A.上次执行时间)" & _
            " Or (" & str持续性 & " And Trunc(A.执行终止时间)<Trunc(A.上次执行时间)+1))" & _
            " And A.开始执行时间 is Not NULL And Nvl(A.医嘱状态,0)<>-1" & _
            " And Nvl(A.执行标记,0)<>-1 And A.病人来源<>3 And NVL(a.执行频次,'无')<>'必要时' And NVL(a.执行频次,'无')<>'需要时'" & _
            IIF(mbln只显示当前病区医嘱, " And instr(',' || [3] || ',',',' || Decode(NVL(A.婴儿,0),0,a.病人科室ID,NVL(b.婴儿科室ID,a.病人科室ID)) || ',')>0 ", "") & _
            Decode(mint医嘱处理范围, 1, " And nvl(a.婴儿,0) = 0 ", 2, " And nvl(a.婴儿,0) <> 0 ", "") & _
            " And (B.婴儿科室ID is null or B.婴儿科室ID is not null and B.婴儿病区ID=[4] and NVL(A.婴儿,0)<>0 or B.婴儿科室ID is not null and B.婴儿病区ID<>[4] and NVL(A.婴儿,0)=0)" & _
            " Order by D.编码,LPAD(B.出院病床,10,' '),A.婴儿,组号,组ID,A.序号"
        On Error GoTo errH
        Set rsAdvice = zlDatabase.OpenSQLRecord(strSQL, "超期收回", Val(Split(str病人IDs, ",")(k)), Val(Split(str主页IDs, ",")(k)), mstr科室IDs, mlng医护科室ID)
        
        '计算并显示收回清单
        '----------------------------------------------------------------------------------------------------------
        If Not rsAdvice.EOF Then
            With vsAdvice
                .Redraw = flexRDNone
                For i = 1 To rsAdvice.RecordCount
                    If NVL(rsAdvice!相关ID, rsAdvice!ID) = lngDel组ID Then
                        GoTo NextLoop '一组医嘱后续不用再处理
                    Else
                        lngDel组ID = 0
                    End If
                    
                    '加入当前行
                    .Rows = .Rows + 1: lngRow = .Rows - 1
                    
                    '隐藏相关行(虽然不含检查手术,但也处理了)
                    If rsAdvice!诊疗类别 = "7" Then
                        .RowHidden(lngRow) = True '单味中药行
                    ElseIf rsAdvice!诊疗类别 = "E" And NVL(rsAdvice!相关ID, 0) = Val(.TextMatrix(lngRow - 1, COL_相关ID)) And NVL(rsAdvice!相关ID, 0) <> 0 Then
                        .RowHidden(lngRow) = True '配方煎法行
                    ElseIf rsAdvice!诊疗类别 = "E" And rsAdvice!ID = Val(.TextMatrix(lngRow - 1, COL_相关ID)) _
                        And InStr(",5,6,", .TextMatrix(lngRow - 1, COL_诊疗类别)) > 0 Then
                        .RowHidden(lngRow) = True '给药途径
                    ElseIf InStr(",D,F,G,E,", rsAdvice!诊疗类别) > 0 And Not IsNull(rsAdvice!相关ID) Then
                        .RowHidden(lngRow) = True '检查部位,附加手术,手术麻醉,输血途径
                    End If
                    
                    '一般列赋值
                    '---------------------------------------------------------------
                    If NVL(rsAdvice!婴儿, 0) = 0 Then
                        .TextMatrix(lngRow, COL_婴儿) = "病人"
                    Else
                        .TextMatrix(lngRow, COL_婴儿) = "婴儿" & rsAdvice!婴儿
                        .ColHidden(COL_婴儿) = False '有婴儿医嘱时才显示
                    End If
                    
                    .TextMatrix(lngRow, COL_科室) = rsAdvice!科室
                    If InStr(str科室 & ",", "," & rsAdvice!科室 & ",") = 0 Then
                        If str科室 <> "" Then .ColHidden(COL_科室) = False
                        str科室 = str科室 & "," & rsAdvice!科室
                    End If
                    
                    .TextMatrix(lngRow, COL_病人ID) = rsAdvice!病人ID
                    .TextMatrix(lngRow, COL_主页ID) = rsAdvice!主页ID
                    .TextMatrix(lngRow, COL_险类) = NVL(rsAdvice!险类)
                    .TextMatrix(lngRow, COL_姓名) = rsAdvice!姓名
                    .TextMatrix(lngRow, COL_住院号) = NVL(rsAdvice!住院号)
                    .TextMatrix(lngRow, COL_床号) = NVL(rsAdvice!床号)
                    .TextMatrix(lngRow, COL_ID) = rsAdvice!ID
                    .TextMatrix(lngRow, COL_相关ID) = NVL(rsAdvice!相关ID)
                    .TextMatrix(lngRow, COL_诊疗类别) = rsAdvice!诊疗类别
                    .TextMatrix(lngRow, col_医嘱内容) = NVL(rsAdvice!医嘱内容)
                    .TextMatrix(lngRow, COL_单位) = NVL(rsAdvice!计算单位)
                    .TextMatrix(lngRow, COL_频率) = NVL(rsAdvice!频率)
                    .TextMatrix(lngRow, COL_执行时间) = NVL(rsAdvice!执行时间方案)
                    .TextMatrix(lngRow, COL_上次执行) = Format(NVL(rsAdvice!上次执行时间), "yyyy-MM-dd HH:mm")
                    .TextMatrix(lngRow, COL_终止时间) = Format(NVL(rsAdvice!执行终止时间), "yyyy-MM-dd HH:mm")
                    .TextMatrix(lngRow, COL_执行科室) = NVL(rsAdvice!执行科室)
                    .TextMatrix(lngRow, COL_执行科室ID) = NVL(rsAdvice!执行科室ID, 0)
                    .TextMatrix(lngRow, COL_执行性质) = NVL(rsAdvice!执行性质, 0)
                    
                    .TextMatrix(lngRow, COL_病人科室ID) = NVL(rsAdvice!病人科室id, 0)
                    .TextMatrix(lngRow, COL_开嘱科室ID) = NVL(rsAdvice!开嘱科室id, 0)
                    .TextMatrix(lngRow, COL_开嘱医生) = NVL(rsAdvice!开嘱医生)
                    .TextMatrix(lngRow, COL_操作类型) = NVL(rsAdvice!操作类型)
                    .TextMatrix(lngRow, COL_执行分类) = NVL(rsAdvice!执行分类)
                    .TextMatrix(lngRow, COL_病人性质) = NVL(rsAdvice!病人性质)
                    
                    '计算收回次数(要管暂停时段,不然可能多收回)
                    '---------------------------------------------------------------
                    lng次数 = 0: str分解时间 = "": str上次时间 = ""
                    strPause = GetAdvicePause(rsAdvice!ID)
                    If IsNull(rsAdvice!执行时间方案) And (NVL(rsAdvice!频率次数, 0) = 0 Or NVL(rsAdvice!频率间隔, 0) = 0 Or IsNull(rsAdvice!间隔单位)) Then
                        '"持续性"的长嘱
                        Call Calc持续性长嘱次数(rsAdvice!开始执行时间, rsAdvice!上次执行时间, "", "", strPause, "", "", str分解时间)
                        arr分解时间 = Split(str分解时间, ",")
                        For j = 0 To UBound(arr分解时间)
                            If Format(arr分解时间(j), "yyyy-MM-dd") <= Format(rsAdvice!执行终止时间, "yyyy-MM-dd") Then
                                str上次时间 = Format(arr分解时间(j), "yyyy-MM-dd HH:mm:ss")
                            Else
                                lng次数 = lng次数 + 1
                            End If
                        Next
                    Else
                        '"可选频率"长嘱
                        str分解时间 = Calc段内分解时间(rsAdvice!开始执行时间, rsAdvice!上次执行时间, strPause, NVL(rsAdvice!执行时间方案), rsAdvice!频率次数, rsAdvice!频率间隔, rsAdvice!间隔单位, rsAdvice!开始执行时间)
                        arr分解时间 = Split(str分解时间, ",")
                        For j = 0 To UBound(arr分解时间)
                            If arr分解时间(j) <= Format(rsAdvice!执行终止时间, "yyyy-MM-dd HH:mm:ss") Then
                                str上次时间 = Format(arr分解时间(j), "yyyy-MM-dd HH:mm:ss")
                            Else
                                lng次数 = lng次数 + 1
                            End If
                        Next
                    End If
                    If lng次数 = 0 Then '不用收回的情况
                        lngDel组ID = NVL(rsAdvice!相关ID, rsAdvice!ID)
                        .RemoveItem lngRow: GoTo NextLoop
                    End If
                    
                    '计算收回总量
                    '---------------------------------------------------------------
                    If rsAdvice!诊疗类别 = "7" Then
                        '如果以前的配方长嘱输有总量，相当于每次的单量
                        .TextMatrix(lngRow, COL_总量) = lng次数 * NVL(rsAdvice!总给予量, 1)
                    ElseIf InStr(",5,6,", rsAdvice!诊疗类别) > 0 Then
                        '西，中成药
                        '------------------
                        '读取原药品规格(自备药无对应费用,从药品目录取一个规格)
                        lng药品ID = 0
                        If Not IsNull(rsAdvice!收费细目ID) Then
                            lng药品ID = rsAdvice!收费细目ID
                        Else
                            '最近发送的药品费用中的药品ID:药品肯定填写了发送记录
                            '药品只有一个收入项目(价格父号=NULL),不排除该费用已被人为操作(例如：划价单被删除)
                            lng药品ID = GetLastSendMediCineID(Val(rsAdvice!ID), CDate(rsAdvice!上次执行时间), Val(rsAdvice!病人性质 & ""))
                        End If
                        '无发送或无对应费用的药品,也不收回(如自备药，或划价单被删除)
                        If lng药品ID = 0 Then
                            lngDel组ID = NVL(rsAdvice!相关ID, rsAdvice!ID)
                            .RemoveItem lngRow: GoTo NextLoop
                        End If
                                                
                        '已经发送过,一定有规格信息
                        strSQL = "Select A.药品ID,A.剂量系数,A.住院包装,A.住院单位," & _
                            " A.药房分批,A.住院可否分零 As 可否分零,Nvl(C.名称,B.名称) as 名称,B.规格,B.产地,A.发药类型" & _
                            " From 药品规格 A,收费项目目录 B,收费项目别名 C" & _
                            " Where A.药品ID=B.ID And A.药名ID=[1] And A.药品ID=[2]" & _
                            " And B.ID=C.收费细目ID(+) And C.码类(+)=1 And C.性质(+)=[3] And Rownum=1"
                        Set rsDrug = zlDatabase.OpenSQLRecord(strSQL, "超期收回", Val(rsAdvice!诊疗项目ID), lng药品ID, IIF(gbyt药品名称显示 = 0, 1, 3))
                        
                        '一但发药后就不再收回，如果有多次发送的,只检查最近一次是否发药(因为分别判断需重算收回次数，太复杂，一般这种按病人发药都是一起发)
                        If Not IsNull(rsDrug!发药类型) Then
                            If InStr("," & strUnRoll & ",", "," & rsDrug!发药类型 & ",") > 0 Then
                                If CheckMedicineSended(Val(rsAdvice!ID), CDate(rsAdvice!上次执行时间)) Then
                                    lngDel组ID = NVL(rsAdvice!相关ID, rsAdvice!ID)
                                    .RemoveItem lngRow: GoTo NextLoop
                                End If
                            End If
                        End If
                        
                        int可否分零 = NVL(rsAdvice!可否分零, NVL(rsDrug!可否分零, 0))
                        
                        .TextMatrix(lngRow, COL_药品ID) = rsDrug!药品ID '记录用于保存时排序
                        .Cell(flexcpData, lngRow, COL_药品ID) = Val(NVL(rsDrug!药房分批, 0))
                        
                        .TextMatrix(lngRow, COL_单位) = rsDrug!住院单位
                        .TextMatrix(lngRow, COL_规格) = rsDrug!名称 & IIF(Not IsNull(rsDrug!产地), "(" & rsDrug!产地 & ")", "") & IIF(Not IsNull(rsDrug!规格), " " & rsDrug!规格, "")
                        
                        '按不分零特性计算收回总量(住院单位)
                        dbl总量 = 0
                        If int可否分零 = 0 Then
                            '可分零
                            dbl总量 = NVL(rsAdvice!单次用量, 0) * lng次数 / rsDrug!剂量系数 / rsDrug!住院包装
                            If str上次时间 = "" And NVL(rsAdvice!首次用量, 0) <> 0 Then
                                '如果上次时间为空则包括首次
                                dbl总量 = dbl总量 + (NVL(rsAdvice!首次用量, 0) - NVL(rsAdvice!单次用量, 0)) / rsDrug!剂量系数 / rsDrug!住院包装
                            End If
                        ElseIf int可否分零 = 1 Then
                            '不分零:按少退,不足一个住院单位不退,足的按小于的整数退
                            dbl总量 = Int(NVL(rsAdvice!单次用量, 0) * lng次数 / rsDrug!剂量系数 / rsDrug!住院包装)
                            If str上次时间 = "" And NVL(rsAdvice!首次用量, 0) <> 0 Then
                                '如果上次时间为空则包括首次
                                dbl总量 = dbl总量 + (NVL(rsAdvice!首次用量, 0) - NVL(rsAdvice!单次用量, 0)) / rsDrug!剂量系数 / rsDrug!住院包装
                            End If
                        ElseIf int可否分零 = 2 Then
                            '一次性(即时失效)
                            dbl总量 = lng次数 * IntEx(NVL(rsAdvice!单次用量, 0) / rsDrug!剂量系数 / rsDrug!住院包装)
                            If str上次时间 = "" And NVL(rsAdvice!首次用量, 0) <> 0 Then
                                '如果上次时间为空则包括首次
                                dbl总量 = dbl总量 + (NVL(rsAdvice!首次用量, 0) - NVL(rsAdvice!单次用量, 0)) / rsDrug!剂量系数 / rsDrug!住院包装
                            End If
                        ElseIf int可否分零 < 0 Then
                            'N天内分零有效:收回量=上次发送量-收回后应发量
                            
                            '应上次发送末次时间=当前上次执行时间
                            If str上次时间 <> "" Then
                                '停止终止时间在多次发送之间
                                strSQL = "Select Min(首次时间) as 首次时间,Max(末次时间) as 末次时间" & _
                                    " From 病人医嘱发送 Where 医嘱ID=[1] And [2]<=末次时间"
                                Set rsSend = zlDatabase.OpenSQLRecord(strSQL, "超期收回", Val(rsAdvice!ID), CDate(str上次时间))
                            Else
                                strSQL = "Select 首次时间,末次时间 From 病人医嘱发送 Where 医嘱ID=[1]" & _
                                    " And 发送号=(Select Max(发送号) From 病人医嘱发送 Where 医嘱ID=[1])"
                                Set rsSend = zlDatabase.OpenSQLRecord(strSQL, "超期收回", Val(rsAdvice!ID))
                            End If
                            
                            '计算上次发送的总量:返回的分解时间次数应与给药途径的"发送数次"相同
                            datBegin = Calc本周期开始时间(rsAdvice!开始执行时间, rsSend!首次时间, rsAdvice!频率间隔, rsAdvice!间隔单位)
                            str分解时间 = Calc段内分解时间(datBegin, rsSend!末次时间, strPause, NVL(rsAdvice!执行时间方案), rsAdvice!频率次数, rsAdvice!频率间隔, rsAdvice!间隔单位, rsAdvice!开始执行时间)
                            dbl总量All = Calc发送药品总量(rsAdvice!开始执行时间, 0, str分解时间, _
                                    NVL(rsAdvice!单次用量, 0), rsDrug!剂量系数, rsDrug!住院包装, int可否分零, _
                                    CDate("3000-01-01"), strPause, NVL(rsAdvice!执行时间方案), _
                                    rsAdvice!频率次数, rsAdvice!频率间隔, rsAdvice!间隔单位, mblnLimit, NVL(rsAdvice!首次用量, 0))
                            If str上次时间 <> "" Then
                                str分解时间 = Calc段内分解时间(datBegin, CDate(str上次时间), strPause, NVL(rsAdvice!执行时间方案), rsAdvice!频率次数, rsAdvice!频率间隔, rsAdvice!间隔单位, rsAdvice!开始执行时间)
                                dbl总量 = Calc发送药品总量(rsAdvice!开始执行时间, 0, str分解时间, _
                                        NVL(rsAdvice!单次用量, 0), rsDrug!剂量系数, rsDrug!住院包装, int可否分零, _
                                        CDate("3000-01-01"), strPause, NVL(rsAdvice!执行时间方案), _
                                        rsAdvice!频率次数, rsAdvice!频率间隔, rsAdvice!间隔单位, mblnLimit, NVL(rsAdvice!首次用量, 0))
                                dbl总量 = dbl总量All - dbl总量
                            Else
                                '为空表示全部收回的情况
                                dbl总量 = dbl总量All
                            End If
                        End If
                        .TextMatrix(lngRow, COL_总量) = FormatEx(dbl总量, 5)
                                                
                        '药品其它信息
                        .TextMatrix(lngRow, COL_剂量系数) = rsDrug!剂量系数
                        .TextMatrix(lngRow, COL_住院包装) = rsDrug!住院包装
                        .TextMatrix(lngRow, COL_可否分零) = int可否分零
                        
                        '显示药品给药途径
                        .TextMatrix(lngRow, COL_用法) = "" & rsAdvice!给药途径
                        
                        .ColHidden(COL_规格) = gbln药品按规格下医嘱
                    Else
                        '非药医嘱
                        '------------------
                        .TextMatrix(lngRow, COL_总量) = lng次数 * NVL(rsAdvice!单次用量, 1)
                        If str上次时间 = "" And NVL(rsAdvice!首次用量, 0) <> 0 Then
                            '如果上次时间为空则包括首次
                            dbl总量 = dbl总量 + (NVL(rsAdvice!首次用量, 0) - NVL(rsAdvice!单次用量, 0))
                        End If
                        
                        '中药配方单位
                        If rsAdvice!诊疗类别 = "E" And rsAdvice!ID = Val(.TextMatrix(lngRow - 1, COL_相关ID)) _
                            And InStr(",E,7,", .TextMatrix(lngRow - 1, COL_诊疗类别)) > 0 Then
                            .TextMatrix(lngRow, COL_单位) = "付"
                        End If
                    End If
                    
                    .TextMatrix(lngRow, COL_单量) = NVL(rsAdvice!单次用量, 0)   '中药存储的是频率
                    .TextMatrix(lngRow, COL_计算量) = .TextMatrix(lngRow, COL_总量)
                    .TextMatrix(lngRow, COL_次数) = lng次数
                    .TextMatrix(lngRow, COL_上次) = str上次时间 '可能为空,如全部收回的情况
                    .Cell(flexcpData, lngRow, COL_总量) = .TextMatrix(lngRow, COL_总量) '用于输入恢复
                    
                    '其它处理
                    '---------------------------------------------------------------
                    '病人计数及分隔
                    If rsAdvice!病人ID <> lng病人ID Then
                        lng病人数 = lng病人数 + 1
                        If lng病人ID <> 0 Then
                            For j = lngRow - 1 To .FixedRows Step -1
                                If Not .RowHidden(j) Then
                                    .CellBorderRange j, .FixedCols, j, .Cols - 1, vbBlack, 0, 0, 0, 2, 0, 0
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                    lng病人ID = rsAdvice!病人ID

NextLoop:           '---------------------------------------------------------------
                    Progress = i / rsAdvice.RecordCount * 100
                    rsAdvice.MoveNext
                Next
            End With
        End If
    Next
    
    lblInfo.Caption = "共有" & IIF(str科室 = "", " ", "(" & Mid(str科室, 2) & ") ") & lng病人数 & " 个病人的医嘱"
    With vsAdvice
        .RowHeight(0) = 320
        If Not .ColHidden(COL_规格) Then
            .AutoSize col_医嘱内容, COL_规格
        Else
            .AutoSize col_医嘱内容
        End If
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        
        .Col = .FixedCols
        For i = .FixedRows To .Rows - 1
            If Not .RowHidden(i) Then
                .Row = i: Exit For
            End If
        Next
        Call .ShowCell(.Row, .Col)
        .Redraw = flexRDDirect
        
        '只有一行时，选中
        k = 0
        For i = .FixedRows To .Rows - 1
            If Not .RowHidden(i) Then k = k + 1
            If k > 1 Then Exit For
        Next
        If k = 1 Or mblnOnePati Then    '单病人调用模式，全选
            If Val(.TextMatrix(.Rows - 1, COL_ID)) <> 0 Then Call tbr_ButtonClick(tbr.Buttons("全选"))
        End If
        If mblnAdjustNum Then
            .Cell(flexcpBackColor, .FixedRows, COL_总量, .Rows - 1, COL_总量) = COLEditBackColor       '浅绿
        End If
    End With
    Progress = 0: Screen.MousePointer = 0
    LoadAdviceRoll = True
    Exit Function
errH:
    vsAdvice.Redraw = flexRDDirect
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        vsAdvice.Redraw = flexRDNone: Resume
    End If
    Call SaveErrLog
    lblInfo.Caption = "": Progress = 0
End Function

Private Function RowCanRoll(ByVal lngRow As Long, Optional strMsg As String) As Boolean
'功能：判断指定行是否允许收回(一组医嘱一起判断)
'参数：strMsg=返回不允许收回的原因提示
    Dim lngBegin As Long, lngEnd As Long, i As Long
    
    strMsg = "": RowCanRoll = True
    
    With vsAdvice
        If mbln超期负数 Then
            If InStr(",5,6,", .TextMatrix(lngRow, COL_诊疗类别)) > 0 Then
                If Val(.Cell(flexcpData, lngRow, COL_药品ID)) = 1 Then
                    strMsg = "分批管理的药品不允许负数方式记帐，该医嘱不能收回。"
                    RowCanRoll = False: Exit Function
                End If
            End If
        End If
        
        Call GetGroupRow(lngRow, lngBegin, lngEnd)
        For i = lngBegin To lngEnd
            If Val(.TextMatrix(i, COL_险类)) <> 0 And Val(.TextMatrix(i, COL_总量)) > 0 Then
                If mbln超期负数 Then
                    If Not gclsInsure.GetCapability(support负数记帐, Val(.TextMatrix(i, COL_病人ID)), Val(.TextMatrix(i, COL_险类))) Then
                        strMsg = "该医保病人所属险类不允许负数方式记帐，该医嘱不能收回。"
                        RowCanRoll = False: Exit Function
                    End If
                Else
                    If Not gclsInsure.GetCapability(support允许部份冲销单据, Val(.TextMatrix(i, COL_病人ID)), Val(.TextMatrix(i, COL_险类))) Then
                        strMsg = "该医保病人所属险类不允许部份冲销费用，该医嘱不能收回。"
                        RowCanRoll = False: Exit Function
                    End If
                End If
            End If
        Next
    End With
End Function

Private Sub GetCurBillSet(ByVal strKey As String, strNO As String)
'功能：获取当前记帐单据的NO
    mrsBill.Filter = "Key='" & strKey & "'"
    If mrsBill.EOF Then
        mrsBill.AddNew
        mrsBill!Key = strKey
        mrsBill!NO = zlDatabase.GetNextNo(14)
        mrsBill.Update
    End If
    strNO = mrsBill!NO
End Sub

Public Function RollAdvice(ByVal lngCount As Long) As Boolean
'功能：处理医嘱发送(这个过程中记帐报警)
'参数：lngCount=已选择的行数
'说明：逐个病人发送提交
    Dim arrSQL() As Variant
    Dim strSQL As String, strNOKey As String
    Dim curDate As Date, blnTran As Boolean
    Dim strNO As String, strTmp As String
    Dim i As Long, j As Long, k As Long
    Dim int配方数 As Integer, dbl总量 As Double, dbl剂量系数 As Double, dbl住院包装 As Double
    Dim str医嘱IDs As String, str中药医嘱IDs As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo errH
    Screen.MousePointer = 11
    
    If mbln超期负数 Then Call InitBillSet
    curDate = zlDatabase.Currentdate
    arrSQL = Array()
    int配方数 = 1
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, COL_选择)) <> 0 And Val(.TextMatrix(i, COL_执行性质)) <> 0 Then '排开叮嘱
                If mbln超期负数 Then
                    dbl剂量系数 = Val(.TextMatrix(i, COL_剂量系数))
                    If dbl剂量系数 = 0 Then dbl剂量系数 = 1
                    If InStr(",7,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                        '界面显示的付数，需要乘以单量（即每日n剂）
                        dbl总量 = Format(Val(.TextMatrix(i, COL_总量)) * Val(.TextMatrix(i, COL_单量)) / dbl剂量系数, "0.00000")
                    Else
                        dbl住院包装 = Val(.TextMatrix(i, COL_住院包装))
                        If dbl住院包装 = 0 Then dbl住院包装 = 1
                        '还原为售价单位来计算(发送记录是售价单位)
                        dbl总量 = Format(Val(.TextMatrix(i, COL_总量)) * dbl住院包装 * dbl剂量系数, "0.00000")
                    End If
                    
                    If CheckAllPrice(Val(.TextMatrix(i, COL_ID)), dbl总量, Val(.TextMatrix(i, COL_病人性质))) Then
                        strNO = "调整划价单"
                    Else
                        '产生单据号分配关键字:与发送时的分号规则相同
                        '-----------------------------------------------------------------------------------------
                        If InStr(",5,6,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                            '中西成药按"病人(病人ID,主页ID)_病人科室ID_开嘱科室ID_开嘱医生_执行科室ID"分号。
                            strNOKey = "中西成药_" & Val(.TextMatrix(i, COL_病人ID)) & "_" & Val(.TextMatrix(i, COL_主页ID)) & "_" & _
                                Val(.TextMatrix(i, COL_病人科室ID)) & "_" & Val(.TextMatrix(i, COL_开嘱科室ID)) & "_" & _
                                .TextMatrix(i, COL_开嘱医生) & "_" & Val(.TextMatrix(i, COL_执行科室ID))
                        ElseIf InStr(",4,M,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                            '材料按"病人(病人ID,主页ID)_病人科室ID_开嘱科室ID_开嘱医生_执行科室ID"分号。
                            strNOKey = "材料医嘱_" & Val(.TextMatrix(i, COL_病人ID)) & "_" & Val(.TextMatrix(i, COL_主页ID)) & "_" & _
                                Val(.TextMatrix(i, COL_病人科室ID)) & "_" & Val(.TextMatrix(i, COL_开嘱科室ID)) & "_" & _
                                .TextMatrix(i, COL_开嘱医生) & "_" & Val(.TextMatrix(i, COL_执行科室ID))
                        ElseIf .TextMatrix(i, COL_诊疗类别) = "7" Then
                            '一个配方中的所有草药分配一个独立单据号
                            strNOKey = "中药配方_" & Val(.TextMatrix(i, COL_病人ID)) & "_" & Val(.TextMatrix(i, COL_主页ID)) & "_" & int配方数
                        ElseIf Val(.TextMatrix(i, COL_相关ID)) <> 0 And .TextMatrix(i, COL_诊疗类别) = "C" Then
                            '一并采集的检验组合分配相同的单据号，标本采集方法分配单独的单据号
                            strNOKey = "一并采集_" & Val(.TextMatrix(i, COL_相关ID))
                        ElseIf Val(.TextMatrix(i, COL_相关ID)) <> 0 And InStr(",F,D,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                            '检查部位和附加手术与主要医嘱分配相同单据号，手术麻醉分配单独的单据号。
                            strNOKey = "非药医嘱_" & Val(.TextMatrix(i, COL_相关ID))
                        Else
                            '其它非药医嘱每条医嘱一个独立单据号(包括给药途径，配方煎法、用法，采集方式，麻醉方式，输血医嘱/输血途径)
                            strNOKey = "非药医嘱_" & Val(.TextMatrix(i, COL_ID))
                        End If
                            
                        Call GetCurBillSet(strNOKey, strNO)
                    End If
                End If
                '能产生配液记录的组医嘱：给药方式为输液，执行药房为配制中心
                If gstr输液配置中心 <> "" And Not mbln销帐申请 Then
                    If .TextMatrix(i, COL_诊疗类别) = "E" And .TextMatrix(i, COL_操作类型) = "2" And .TextMatrix(i, COL_执行分类) = "1" Then
                        For j = i - 1 To .FixedRows Step -1
                            If Val(.TextMatrix(i, COL_ID)) = Val(.TextMatrix(j, COL_相关ID)) Then
                                If InStr("," & gstr输液配置中心 & ",", "," & Val(.TextMatrix(j, COL_执行科室ID)) & ",") > 0 Then
                                    str医嘱IDs = str医嘱IDs & "," & Val(.TextMatrix(i, COL_ID)): Exit For
                                End If
                            Else
                                Exit For
                            End If
                        Next
                    End If
                End If
                
                If InStr(",7,", .TextMatrix(i, COL_诊疗类别)) > 0 Then
                    str中药医嘱IDs = str中药医嘱IDs & "," & Val(.TextMatrix(i, COL_ID))
                End If
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                
                If .TextMatrix(i, COL_上次) = "" Then
                    strTmp = "NULL"
                Else
                    strTmp = "To_Date('" & .TextMatrix(i, COL_上次) & "','YYYY-MM-DD HH24:MI:SS')"
                End If
                
                arrSQL(UBound(arrSQL)) = _
                    IIF(Val(.TextMatrix(i, COL_药品ID)) = 0, "999999999", Val(.TextMatrix(i, COL_药品ID))) & ":" & _
                    "ZL_病人医嘱记录_收回(" & Val(.TextMatrix(i, COL_总量)) & "," & Val(.TextMatrix(i, COL_ID)) & "," & strTmp & "," & _
                    "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                    IIF(mbln超期负数, "'" & strNO & "'", "NULL") & ")"
                
                '计算中药配方数
                If .TextMatrix(i, COL_诊疗类别) = "E" And Val(.TextMatrix(i, COL_ID)) = Val(.TextMatrix(i - 1, COL_相关ID)) _
                    And InStr(",E,7,", .TextMatrix(i - 1, COL_诊疗类别)) > 0 Then '中药用法
                    int配方数 = int配方数 + 1
                End If
                
                '---------------------------------
                k = k + 1
                Progress = k / (lngCount * 2) * 100
            End If
        Next
        
        If str医嘱IDs <> "" Then
            str医嘱IDs = Mid(str医嘱IDs, 2)
            If Drug配液(str医嘱IDs) Then
                If MsgBox("本次收回的输液医嘱的药品中包含已经配液的记录，不允许收回，是否继续收回其它未配液的记录？", vbQuestion + vbYesNo + vbDefaultButton2, "超期收回") = vbNo Then
                    Progress = 0: Screen.MousePointer = 0
                    RollAdvice = False
                    Exit Function
                End If
            End If
        End If
        
        If str中药医嘱IDs <> "" Then
            str中药医嘱IDs = Mid(str中药医嘱IDs, 2)
            strSQL = "select 1 from 住院费用记录 a where a.记录状态 In (0, 1, 3) And Nvl(a.执行状态,0)<>0" & _
                " and a.医嘱序号 in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) and rownum<2"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str中药医嘱IDs)
            If Not rsTmp.EOF Then
                If MsgBox("本次收回的中草药中存在已经发药的，只能收回未发药的部分，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, "超期收回") = vbNo Then
                    Progress = 0: Screen.MousePointer = 0
                    RollAdvice = False
                    Exit Function
                End If
            End If
        End If
        
        '按药品ID排序(如果有),给药途径及非药排在后面
        For i = 0 To UBound(arrSQL) - 1
            For j = i + 1 To UBound(arrSQL)
                If Val(Left(arrSQL(j), InStr(arrSQL(j), ":") - 1)) < Val(Left(arrSQL(i), InStr(arrSQL(i), ":") - 1)) Then
                    strTmp = arrSQL(j)
                    arrSQL(j) = arrSQL(i)
                    arrSQL(i) = strTmp
                End If
            Next
        Next
                
        '提交数据
        gcnOracle.BeginTrans: blnTran = True
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(Mid(arrSQL(i), InStr(arrSQL(i), ":") + 1), "超期收回")
            '---------------------------------
            k = k + 1
            Progress = k / (lngCount * 2) * 100
        Next
        gcnOracle.CommitTrans: blnTran = False
        
        '提交成功,删除已收回行
        .Redraw = flexRDNone
        For i = .Rows - 1 To .FixedRows Step -1
            If Val(.TextMatrix(i, COL_选择)) <> 0 Then
                .RemoveItem i
            End If
        Next
        If .Rows = .FixedRows Then .Rows = .FixedRows + 1
        For i = .FixedRows To .Rows + 1
            If Not .RowHidden(i) Then
                .Row = i: Exit For
            End If
        Next
        .ShowCell .Row, .Col
        .Redraw = flexRDDirect
    End With
    Progress = 0: Screen.MousePointer = 0
    RollAdvice = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTran Then gcnOracle.RollbackTrans
    If err.Number <> 0 Then '如医保上传失败退出没有错误
        If ErrCenter() = 1 Then Resume
        Call SaveErrLog
    End If
    Progress = 0
End Function

Private Function Drug配液(ByVal str医嘱IDs As String) As Boolean
'功能：医嘱中是否存在已经被配液的
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    strSQL = "Select 1 from 病人医嘱记录 A,输液配药记录 B Where a.Id=b.医嘱id And (b.操作状态 In (4,5,6,7,8) AND NVL(B.是否打包,0) = 0) And b.执行时间>a.执行终止时间" & _
        " and a.Id In (Select Column_Value From Table(Cast(f_Num2list([1]) As Zltools.t_Numlist))) And Rownum<2"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str医嘱IDs)
    Drug配液 = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckAllPrice(ByVal lng医嘱ID As Long, ByVal dbl收回总量 As Double, ByVal lng病人性质 As Long) As Boolean
'功能：检查收回次数的医嘱对应费用是否全是未审核的划价单，以便确定直接修改划价单，无需取新的单据号
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim lng收回数 As Long, lng发送数 As Long
    
    CheckAllPrice = False
    
    '发送数次是按售价单位存储的
    strSQL = "Select Sum(a.发送数次) 发送数次" & vbNewLine & _
            "From 病人医嘱发送 A" & vbNewLine & _
            "Where a.医嘱id = [1] And a.记录性质 = 2 And Not Exists" & vbNewLine & _
            " (Select 1 From " & IIF(lng病人性质 = 1, "门诊", "住院") & "费用记录 B Where a.医嘱id = b.医嘱序号 And a.No = b.No And b.记录性质 = 2 And 记录状态 <> 0)"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
    If rsTmp.RecordCount > 0 Then
        If dbl收回总量 <= Val("" & rsTmp!发送数次) Then CheckAllPrice = True
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsAdvice_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = COL_总量 And Not mblnReturn Then
        vsAdvice.Refresh    '如果有弹出提示，不刷新的话，一并给药通过Drawcell被擦除的单元格会再次显示
        If Not AcceptInput(Row, Col) Then
            Cancel = True
        End If
    End If
End Sub

Private Function CheckAdvcieComPound(ByVal lng医嘱ID As Long) As Boolean
'功能：根据医嘱ID，判断是否是输液配药的药品
    Dim strSQL As String, rsTmp As Recordset
    
    strSQL = "Select 1 from 输液配药记录 Where 医嘱ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng医嘱ID)
    CheckAdvcieComPound = rsTmp.RecordCount > 0
End Function
