VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmLisPrinter 
   BorderStyle     =   0  'None
   ClientHeight    =   7230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10410
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   10410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrReturn 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   210
      Top             =   6660
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      ScaleHeight     =   1275
      ScaleWidth      =   4515
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4680
      Width           =   4575
      Begin VB.HScrollBar hsb 
         Height          =   330
         Left            =   240
         MousePointer    =   99  'Custom
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   900
         Width           =   3135
      End
      Begin VB.VScrollBar vsb 
         Height          =   975
         Left            =   4200
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   120
         Width           =   330
      End
      Begin zl9NewQuery.ctlQueryItem QueryItem 
         Height          =   735
         Left            =   1080
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   120
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1296
      End
      Begin VB.PictureBox picBack1 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   8.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3360
         ScaleHeight     =   735
         ScaleWidth      =   975
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame fraMain 
      Height          =   1735
      Left            =   30
      TabIndex        =   3
      Top             =   -30
      Width           =   10005
      Begin VB.Frame fratak 
         Height          =   105
         Left            =   0
         TabIndex        =   8
         Top             =   960
         Width           =   9675
      End
      Begin VB.TextBox TxtID 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   2730
         TabIndex        =   0
         Top             =   150
         Width           =   6855
      End
      Begin VB.Label lbl提示 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   6780
         TabIndex        =   9
         Top             =   1170
         Width           =   165
      End
      Begin VB.Label lblID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请扫描条码："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   2700
      End
      Begin VB.Image imgTitle 
         Height          =   750
         Left            =   30
         Picture         =   "frmLisPrinter.frx":0000
         Stretch         =   -1  'True
         Top             =   150
         Width           =   2655
      End
      Begin VB.Label lbl年龄 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   525
         Left            =   7320
         TabIndex        =   6
         Top             =   1170
         Width           =   1365
      End
      Begin VB.Label lbl性别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   525
         Left            =   3720
         TabIndex        =   5
         Top             =   1170
         Width           =   1620
      End
      Begin VB.Label lbl姓名 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   26.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   525
         Left            =   120
         TabIndex        =   4
         Top             =   1170
         Width           =   1620
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid msfMain 
      Height          =   2835
      Left            =   30
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1770
      Width           =   9690
      _cx             =   17092
      _cy             =   5001
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   15199202
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16633516
      ForeColorSel    =   16711680
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   16761024
      GridColorFixed  =   16761024
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   25
      Cols            =   30
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   450
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   2
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
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin zl9NewQuery.ctlButton ctlClear 
      Height          =   540
      Left            =   8370
      TabIndex        =   2
      Top             =   5790
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   953
      Caption         =   "清除"
      BackColor       =   16777215
      FontSize        =   21.75
      FontBold        =   -1  'True
      AutoSize        =   0   'False
      ButtonHeight    =   420
   End
   Begin zl9NewQuery.ctlButton ctlReturn 
      Height          =   540
      Left            =   5280
      TabIndex        =   10
      Top             =   5880
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   953
      Caption         =   "返回"
      BackColor       =   16777215
      FontSize        =   21.75
      FontBold        =   -1  'True
      AutoSize        =   0   'False
      ButtonHeight    =   420
   End
   Begin VB.Shape shp 
      BorderColor     =   &H00FF0000&
      Height          =   1575
      Left            =   120
      Top             =   4560
      Width           =   4935
   End
End
Attribute VB_Name = "frmLisPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mInputType As Integer               '输入类型:0=条码;1=住院;2=门诊;3=就诊卡;4=IC卡;5=病人ID
Private mstrSource As String                '病人来源
Private mstrNO As String                    '单据编码
Private Enum mCol
    检验项目 = 0
    检验人
    检验时间
    审核人
    审核时间
    状态
    说明
    打印次数
    标本ID
    医嘱id
    病人id
End Enum

Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mblnFist As Boolean
Private mvarPageNo As Long
Private mvarSvrDept As String           '保存增加医生的科室
Private mvarSvrDuty As String           '保存增加医生的职务
Private mlngHelpPage As Long
Private mintPrintDelayed As Integer     '打印延时时间
Private mintClear As Integer            '情况数据时间
Private mintBack As Integer             '时候打印报告后返回主页

Private mvarLeftStart As Single
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private mintReturn As Integer

Private Sub ctlClear_CommandClick()
    Dim intRow As Integer, intCol As Integer
     '清空记录
    Me.lbl姓名 = "姓名:"
    Me.lbl年龄 = "年龄:"
    Me.lbl性别 = "性别:"
    Me.lbl提示 = ""
    With Me.msfMain
        For intRow = 1 To .Rows - 1
            For intCol = 0 To .Cols - 1
                .TextMatrix(intRow, intCol) = ""
            Next
        Next
    End With
    Me.TxtID.Text = ""
    tmrReturn.Enabled = False
    Me.TxtID.SetFocus
End Sub

'Private Sub ctlRead_CommandClick()
'    If mobjICCard Is Nothing Then
'        Set mobjICCard = CreateObject("zlICCard.clsICCard")
'        Set mobjICCard.gcnOracle = gcnOracle
'    End If
'    If Not mobjICCard Is Nothing Then
'        TxtID.Text = mobjICCard.Read_Card()
'        If TxtID.Text <> "" Then
'            Call TxtID_KeyPress(vbKeyReturn)
'            mblnICCard = True
'        End If
'    End If
'End Sub

Private Sub ctlReturn_CommandClick()
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim lngPageNum As Long
    If mblnFist = False Then Exit Sub
    mblnFist = False
  
    DoEvents
    mlngHelpPage = Val(GetPara("自助打印帮助页面"))
    ctlReturn.Visible = Val(GetPara("自助打印显示返回按钮"))
    mintPrintDelayed = Val(GetPara("检验打印延时", 0))
    
    If mlngHelpPage > 0 Then
        Call Form_Resize
        Call LoadPageItemList(mlngHelpPage)
        Call CalcVsb
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    Dim intLoop As Integer
    
    mblnFist = True
    Me.ctlClear.ShowPicture = False
    'Me.ctlPrinter.ShowPicture = False
    Me.ctlReturn.ShowPicture = False
    
    With Me.msfMain
'        .ColWidth(mCol.检验项目) = 4500
'        .TextMatrix(0, mCol.检验项目) = "检验项目"
'
'        .ColWidth(mCol.检验人) = 2000
'        .TextMatrix(0, mCol.检验人) = "核收人"
'
'        .ColWidth(mCol.检验时间) = 2500
'        .TextMatrix(0, mCol.检验时间) = "核收时间"
'
'        .ColWidth(mCol.审核人) = 2000
'        .TextMatrix(0, mCol.审核人) = "审核人"
'
'        .ColWidth(mCol.审核时间) = 2500
'        .TextMatrix(0, mCol.审核时间) = "审核时间"
'
'        .ColWidth(mCol.状态) = 2000
'        .TextMatrix(0, mCol.状态) = "状态"
'
'        .ColWidth(mCol.说明) = 4000
'        .TextMatrix(0, mCol.说明) = "说明"
        
        .ColWidth(mCol.检验项目) = 6000
        .TextMatrix(0, mCol.检验项目) = "检验项目"
        
        .ColWidth(mCol.检验人) = 0
        .TextMatrix(0, mCol.检验人) = "核收人"

        .ColWidth(mCol.检验时间) = 0
        .TextMatrix(0, mCol.检验时间) = "核收时间"

        .ColWidth(mCol.审核人) = 0
        .TextMatrix(0, mCol.审核人) = "审核人"

        .ColWidth(mCol.审核时间) = 0
        .TextMatrix(0, mCol.审核时间) = "审核时间"
        
        .ColWidth(mCol.状态) = 5000
        .TextMatrix(0, mCol.状态) = "状态"
        
        .ColWidth(mCol.说明) = 8500
        .TextMatrix(0, mCol.说明) = "说明"
        
        .ColWidth(mCol.打印次数) = 0
        .TextMatrix(0, mCol.打印次数) = "打印次数"
        
        .ColWidth(mCol.病人id) = 0
        .TextMatrix(0, mCol.病人id) = "病人ID"
        
        .ColWidth(mCol.医嘱id) = 0
        .TextMatrix(0, mCol.医嘱id) = "医嘱ID"
        
        .ColWidth(mCol.标本ID) = 0
        .TextMatrix(0, mCol.标本ID) = "标本ID"
    End With
    
    Set mobjIDCard = New clsIDCard
    Set mobjICCard = New clsICCard
    Call mobjIDCard.SetParent(Me.hwnd)
    Call mobjICCard.SetParent(Me.hwnd)
    
    mInputType = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\frmPrinterSetup", "查找方式", 0)
    mstrSource = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\frmPrinterSetup", "病人来源", "0,0,0")
    mstrNO = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\frmPrinterSetup", "诊疗单据", "")
    mintClear = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\frmPrinterSetup", "打印报告后清空", 0)
    mintBack = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\frmPrinterSetup", "打印报告后返回主页", 0)
    tmrReturn.Enabled = False
    Select Case mInputType
    
        Case 0              '条码
            Me.lblID.Caption = "请扫描条码："
        Case 1              '住院
            Me.lblID.Caption = "请扫描条码："
        Case 2              '门诊
            Me.lblID.Caption = "请扫描条码："
        Case 3              '就诊卡
            Me.lblID.Caption = "请刷就诊卡："
        Case 4              'IC卡
            Me.lblID.Caption = "请刷IC卡："
        Case 5              '病人ID
            Me.lblID.Caption = "请扫描条码："
    
    End Select
    Call Form_Resize
    Call InitSysPar
End Sub

Private Sub Form_Resize()
    With Me.fraMain
        .Width = Me.Width - 100
    End With
    
    With Me.fratak
        .Width = fraMain.Width - 100
    End With

    With Me.TxtID
        .Width = fraMain.Width - .Left - 100
    End With
    
    With Me.msfMain
        .Width = Me.Width - 100
        .Height = Me.Height - .Top - 1000
    End With
    
    With Me.ctlClear
        .Left = Me.Width - .Width - 300
        .Top = Me.Height - .Height - 250
    End With
    
'    With Me.ctlPrinter
'        .Left = Me.ctlClear.Left - .Width - 300
'        .Top = Me.ctlClear.Top
'    End With
    
    With Me.ctlReturn
        .Left = Me.Left + 900
        .Top = Me.ctlClear.Top
    End With
    
    With Me.lbl提示
        .Left = 10920
    End With
    
    picBack.Enabled = mlngHelpPage > 0
    picBack1.Enabled = mlngHelpPage > 0
    hsb.Enabled = mlngHelpPage > 0
    vsb.Enabled = mlngHelpPage > 0

    shp.Visible = mlngHelpPage > 0
    picBack.Visible = mlngHelpPage > 0
    picBack1.Visible = mlngHelpPage > 0
    hsb.Visible = mlngHelpPage > 0
    vsb.Visible = mlngHelpPage > 0
    
    If mlngHelpPage > 0 Then
        With Me.msfMain
            .Width = Me.Width - 100
            .Height = Me.Height - .Top - 4000
        End With
        
        
        QueryItem.Width = Screen.Width - 2010 - 45
        Call ResizeControl(shp, 15, Me.msfMain.Top + Me.msfMain.Height + 30, Me.ScaleWidth - 30, Me.ScaleHeight - (Me.msfMain.Top + Me.msfMain.Height + 30) - (Me.ScaleHeight - ctlClear.Top) - 100)
        
        Call ResizeControl(picBack, 45, Me.shp.Top + 30, Me.shp.Width - 60, Me.shp.Height - 60)
        Call ResizeControl(QueryItem, picBack.Left, 30, QueryItem.Width, QueryItem.Height)
        
        mvarLeftStart = QueryItem.Left
        
        Call ResizeControl(vsb, picBack.ScaleWidth - vsb.Width + 60, 0, vsb.Width, picBack.ScaleHeight - hsb.Height + 60)
        Call ResizeControl(hsb, 0, picBack.ScaleHeight - hsb.Height + 60, picBack.ScaleWidth - vsb.Width + 60, hsb.Height)
        picBack1.Left = vsb.Left
        picBack1.Top = hsb.Top
        
        Call CalcVsb
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mobjIDCard Is Nothing Then
        Call mobjIDCard.SetEnabled(False)
        Set mobjIDCard = Nothing
    End If
    If Not mobjICCard Is Nothing Then
        Call mobjICCard.SetEnabled(False)
        Set mobjICCard = Nothing
    End If
End Sub

Private Sub mobjICCard_ShowICCardInfo(ByVal strCardNO As String)
    If Not TxtID.Locked And TxtID.Text = "" And Me.ActiveControl Is TxtID Then
        TxtID.Text = strCardNO
        Call TxtID_KeyPress(vbKeyReturn)
    End If
End Sub

Private Sub msfMain_Click()
    Me.TxtID.SetFocus
End Sub

Private Sub msfMain_SelChange()
    With Me.msfMain
        If .TextMatrix(.Row, mCol.状态) = "不可打印" Then
            .ForeColorSel = &HC0&
        Else
            .ForeColorSel = &HFF0000
        End If
    End With
End Sub

Private Sub tmrReturn_Timer()
    ''定时清空
    If mintReturn - 1 < 0 Then
        Call ctlClear_CommandClick
        If mintBack = 1 Then
            ctlReturn_CommandClick
        End If
    Else
        mintReturn = mintReturn - 1
        lbl提示.Caption = mintReturn & "秒后,将清除病人信息!"
        If mintBack = 1 Then
        
             lbl提示.Caption = lbl提示.Caption & "并返回首页"
        End If
    End If
End Sub

Private Sub TxtID_Change()
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (TxtID.Text = "" And Me.ActiveControl Is TxtID)
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (TxtID.Text = "" And Me.ActiveControl Is TxtID)
End Sub

Private Sub TxtID_GotFocus()
    Me.TxtID.SelStart = 0
    Me.TxtID.SelLength = Len(Me.TxtID)
    If Not mobjIDCard Is Nothing And TxtID.Text = "" And Not TxtID.Locked Then mobjIDCard.SetEnabled (True)
    If Not mobjICCard Is Nothing And TxtID.Text = "" And Not TxtID.Locked Then mobjICCard.SetEnabled (True)
End Sub

Private Sub TxtID_KeyPress(KeyAscii As Integer)
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim intRow As Integer, intCol As Integer
    Dim blnPrinter As Boolean           '是否有可以打印的
    Dim blnCard As Boolean
    Dim strSource As String
    Dim strAdvice As Long               '医嘱id
    Dim strTimeHorizon As String
    Dim intDays As Integer
    Dim intPrintend As Integer           '已完成打印
    Dim intNotPrint As Integer           '未打印
    Dim intPrinting As Integer           '打印中
    
    
    If CheckIsInclude(UCase(Chr(KeyAscii)), "'‘’;；:：?？|,，.。""") = True Then KeyAscii = 0
    Call zlCommFun.InputIsCard(TxtID, KeyAscii, glngSys)
    
    
    TxtID.Text = ReplaseSpecial(TxtID.Text)
    '是否刷卡完成
    
    blnCard = KeyAscii <> 8 And Len(TxtID.Text) = gbytCardNOLen - 1 And TxtID.SelLength <> Len(TxtID.Text)
    If KeyAscii = 13 Then blnCard = True
    If mInputType = 3 Then
        '就诊卡只要输入的位数够就执行
        If blnCard = False Then Exit Sub
        If KeyAscii <> 13 Then
            Me.TxtID = Me.TxtID & Chr(KeyAscii)
        End If
        KeyAscii = 0
    Else
        If KeyAscii <> 13 Then Exit Sub
    End If
    
    strTimeHorizon = GetPara("报告日期范围", "0")
    If Split(strTimeHorizon, "-")(0) = "1" Then
        intDays = Val(Split(strTimeHorizon, "-")(1))
    Else
        intDays = 30
    End If
    
    If Me.TxtID = "" Then Exit Sub
    
    blnPrinter = False
    
    On Error GoTo errH
    
    strSQL = "Select /*+ rule */" & vbNewLine & _
            " A.医嘱内容 As 检验项目, " & vbNewLine & _
            "         Decode(B.医嘱id, Null, '1-未发送', Decode(B.采样人, Null, '2-未采样', Decode(B.接收人, Null, '3-已采样', '4-已接收'))) As 状态," & vbNewLine & _
            "         '' As 检验人, '' As 检验时间, '' As 审核人, '' As 审核时间, '' As 打印次数, " & vbNewLine & _
            "         e.姓名,e.性别,e.年龄,a.相关id as 医嘱ID,a.病人ID,null as 标本ID " & vbNewLine & _
            "From 病人医嘱记录 A, 病人医嘱发送 B, 部门表 D, 病人信息 E,病历单据应用 F,病历文件列表 G " & vbNewLine & _
            "Where A.Id = B.医嘱id And A.开嘱科室id = D.Id And A.病人id = E.病人id And nvl(B.执行状态,0) = 0 And A.诊疗类别 = 'C' " & vbNewLine & _
            " And a.诊疗项目id = f.诊疗项目ID and f.病历文件id = g.id  And decode(a.病人来源,3,1,a.病人来源) = f.应用场合 " & vbNewLine & _
            " and a.病人来源 in (Select * From Table(Cast(f_Num2list([2]) As zlTools.t_Numlist))) " & vbNewLine & _
            " And g.编号 In (Select * From Table(Cast(f_Num2list([3]) As zlTools.t_Numlist))) " & vbNewLine & _
            " And a.开嘱时间 + 0 between [4] and [5] " & vbNewLine & _
            "查询条件1" & vbNewLine

            strSQL = strSQL & " Union All" & vbNewLine & _
            "Select /*+ rule */" & vbNewLine & _
            "Distinct A.检验项目, Decode(样本状态, 1, '5-已核收', Decode(Sign(Nvl(打印次数, 0)), 1, '7-已打印', '6-已审核')) As 状态, A.检验人," & vbNewLine & _
            "         To_Char(A.检验时间, 'YYYY-MM-DD HH24:MI:SS') As 检验时间, A.审核人, To_Char(A.审核时间, 'YYYY-MM-DD HH24:MI:SS') As 审核时间," & vbNewLine & _
            "         Decode(Nvl(A.打印次数, 0), 0, '', '√') As 打印次数,e.姓名,e.性别,e.年龄, " & vbNewLine & _
            "         a.医嘱ID,a.病人ID,g.标本ID " & vbNewLine & _
            "From 检验标本记录 A, 病人医嘱记录 B, 病人医嘱发送 D, 部门表 F, 病人信息 E, 检验项目分布 G,病历单据应用 H,病历文件列表 I " & vbNewLine & _
            "Where A.Id = G.标本id And G.医嘱id = B.相关id And B.相关id = D.医嘱id And B.开嘱科室id = F.Id And A.病人id = E.病人id" & vbNewLine & _
            " And b.诊疗项目id = h.诊疗项目ID And H.病历文件id = I.id And decode(b.病人来源,3,1,b.病人来源) = H.应用场合 " & vbNewLine & _
            " And b.病人来源 in (Select * From Table(Cast(f_Num2list([2]) As zlTools.t_Numlist))) " & vbNewLine & _
            " And  I.编号 In (Select * From Table(Cast(f_Num2list([3]) As zlTools.t_Numlist))) " & vbNewLine & _
            " And a.核收时间 + 0 between [4] and [5] " & vbNewLine & _
            "查询条件2 " & vbNewLine & _
            " order by 状态 "


    '输入类型:0=条码;1=住院;2=门诊;3=就诊卡;4=IC卡;5=病人ID
    Select Case mInputType
    
        Case 0              '条码
            strSQL = Replace$(strSQL, "查询条件1", " And b.样本条码 = [1] ")
            strSQL = Replace$(strSQL, "查询条件2", " And d.样本条码 = [1] ")
        Case 1              '住院
            strSQL = Replace$(strSQL, "查询条件1", " And e.住院号 = [1] ")
            strSQL = Replace$(strSQL, "查询条件2", " And e.住院号 = [1] ")
        Case 2              '门诊
            strSQL = Replace$(strSQL, "查询条件1", " And e.门诊号 = [1] ")
            strSQL = Replace$(strSQL, "查询条件2", " And e.门诊号 = [1] ")
        Case 3              '就诊卡
            strSQL = Replace$(strSQL, "查询条件1", " And e.就诊卡号 = [1] ")
            strSQL = Replace$(strSQL, "查询条件2", " And e.就诊卡号 = [1] ")
        Case 4              'IC卡
            strSQL = Replace$(strSQL, "查询条件1", " And e.IC卡号 = [1] ")
            strSQL = Replace$(strSQL, "查询条件2", " And e.IC卡号 = [1] ")
        Case 5              '病人ID
            strSQL = Replace$(strSQL, "查询条件1", " And e.病人ID = [1] ")
            strSQL = Replace$(strSQL, "查询条件2", " And e.病人ID = [1] ")
    End Select
    strSource = IIf(Split(mstrSource, ",")(0) = 1, "1,3", 0)
    strSource = strSource & "," & IIf(Split(mstrSource, ",")(1) = 1, "2", 0)
    strSource = strSource & "," & IIf(Split(mstrSource, ",")(2) = 1, "4", 0)
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Me.TxtID, strSource, mstrNO, CDate(Format(Now - intDays, "yyyy-mm-dd 00:00:00")), CDate(Format(Now, "yyyy-mm-dd 23:59:59")))
    
    '清空记录
    Call ctlClear_CommandClick
    
    '写入记录
    If rsTmp.RecordCount > Me.msfMain.Rows - 1 Then
        Me.msfMain.Rows = rsTmp.RecordCount + 1
    End If
    intRow = 0
    If rsTmp.RecordCount > 0 Then
        Me.lbl姓名 = "姓名:" & Nvl(rsTmp("姓名"))
        Me.lbl年龄 = "年龄:" & Nvl(rsTmp("年龄"))
        Me.lbl性别 = "性别:" & Nvl(rsTmp("性别"))
    End If
    Do While Not rsTmp.EOF
        If strAdvice <> Nvl(rsTmp("医嘱id")) And Nvl(rsTmp("医嘱id")) <> "" Then
            intRow = intRow + 1
            Me.msfMain.TextMatrix(intRow, mCol.检验项目) = Nvl(rsTmp("检验项目"), "")
            
            Me.msfMain.TextMatrix(intRow, mCol.检验人) = Nvl(rsTmp("检验人"), "")
            Me.msfMain.TextMatrix(intRow, mCol.检验时间) = Nvl(rsTmp("检验时间"), "")
            Me.msfMain.TextMatrix(intRow, mCol.审核人) = Nvl(rsTmp("审核人"), "")
            Me.msfMain.TextMatrix(intRow, mCol.审核时间) = Nvl(rsTmp("审核时间"), "")
            With Me.msfMain
                Select Case Nvl(rsTmp("状态"), "")
                    Case "1-未发送", "2-未采样", "3-已采样", "4-已接收", "5-已核收"
                        Me.msfMain.TextMatrix(intRow, mCol.状态) = "不可打印"
                        .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = vbRed
                        .Cell(flexcpFontBold, intRow, 0, intRow, .Cols - 1) = True
                        intNotPrint = intNotPrint + 1
                    Case "7-已打印"
                        Me.msfMain.TextMatrix(intRow, mCol.状态) = "不可打印"
                        .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = vbRed
                        .Cell(flexcpFontBold, intRow, 0, intRow, .Cols - 1) = True
                    Case "6-已审核"
                        Me.msfMain.TextMatrix(intRow, mCol.状态) = "可以打印"
                        .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = vbBlack
                        .Cell(flexcpFontBold, intRow, 0, intRow, .Cols - 1) = False
                        intPrinting = intPrinting + 1
                End Select
            End With
            Select Case Nvl(rsTmp("状态"), "")
                Case "1-未发送"
                    Me.msfMain.TextMatrix(intRow, mCol.说明) = "医嘱未发送！"
                Case "2-未采样"
                    Me.msfMain.TextMatrix(intRow, mCol.说明) = "标本未采样！"
                Case "3-已采样"
                    Me.msfMain.TextMatrix(intRow, mCol.说明) = "标本等待检验..."
                Case "4-已接收"
                    Me.msfMain.TextMatrix(intRow, mCol.说明) = "标本正在检验，请稍后再来！"
                Case "5-已核收"
                    Me.msfMain.TextMatrix(intRow, mCol.说明) = "标本正在检验，请稍后再来！"
                Case "6-已审核"
                    
                Case "7-已打印"
                    Me.msfMain.TextMatrix(intRow, mCol.说明) = "已打印不能再打印！"
            End Select
            Me.msfMain.TextMatrix(intRow, mCol.打印次数) = Nvl(rsTmp("打印次数"), "")
            If Nvl(rsTmp("打印次数"), "") = "" Then
                blnPrinter = True
            End If
            Me.msfMain.TextMatrix(intRow, mCol.医嘱id) = Nvl(rsTmp("医嘱ID"), "")
            Me.msfMain.TextMatrix(intRow, mCol.病人id) = Nvl(rsTmp("病人ID"), "")
            Me.msfMain.TextMatrix(intRow, mCol.标本ID) = Nvl(rsTmp("标本ID"), "")
        End If
        strAdvice = Nvl(rsTmp("医嘱id"))
        
        rsTmp.MoveNext
    Loop
    Me.lbl提示.Caption = "正在打印报告" & intPrinting & "张，检验中报告" & intNotPrint & "张未打印！"
    Call msfMain_SelChange
    Me.TxtID.Text = ""
    Me.TxtID.SetFocus
    
    '打印报告
    If blnPrinter Then
        With Me.msfMain
            For intRow = 1 To .Rows - 1
                If .TextMatrix(intRow, mCol.审核人) <> "" And .TextMatrix(intRow, mCol.打印次数) = "" Then
                    .TextMatrix(intRow, mCol.状态) = "正在打印…"
                    '有审核并没有打印过时才进行打印
                    If ReportPrint(Val(.TextMatrix(intRow, mCol.医嘱id)), Val(.TextMatrix(intRow, mCol.标本ID)), Val(.TextMatrix(intRow, mCol.病人id)), True) = True Then
                        .TextMatrix(intRow, mCol.打印次数) = Val(.TextMatrix(intRow, mCol.打印次数)) + 1
                        .TextMatrix(intRow, mCol.状态) = "已打印"
                        .TextMatrix(intRow, mCol.说明) = "已打印不能再打印！"
                        .Cell(flexcpForeColor, intRow, 0, intRow, .Cols - 1) = vbRed
                        .Cell(flexcpFontBold, intRow, 0, intRow, .Cols - 1) = True
'                        intPrinting = intPrinting - 1
                        Me.lbl提示.Caption = "本次共打印报告" & intPrinting & "张，检验中报告" & intNotPrint & "张未打印！"
'                        Me.lbl提示.Caption = "本次共打印报告" & intPrinting & "张。"
                    Else
                        .TextMatrix(intRow, mCol.状态) = "打印失败"
                    End If
                    If mintPrintDelayed > 0 Then
                        Call Sleep(mintPrintDelayed * 1000)
                    End If
                End If
            Next
            Me.lbl提示.Caption = "打印完成！请注意取走所有报告！"
            Call Sleep(2 * 1000)
            Me.lbl提示.Caption = ""
        End With
        Call msfMain_SelChange
        Me.TxtID.SetFocus
    End If
    If mintClear > 0 Then
        If rsTmp.RecordCount > 0 Then
            tmrReturn.Interval = 1000
            mintReturn = mintClear
            tmrReturn.Enabled = True
        Else
            If mintBack = 1 Then
                tmrReturn.Interval = 1000
                mintReturn = 5
                tmrReturn.Enabled = True
            End If
        End If
    Else
        If mintBack = 1 Then
            tmrReturn.Interval = 1000
            mintReturn = 5
            tmrReturn.Enabled = True
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Function ReadImageData(lngKeyID As Long) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim DrawIndex As Integer
    Dim StrTime As Date
    On Error GoTo errH
    StrTime = Now
    gstrSQL = "select id ,标本ID,图像类型 from 检验图像结果 where 标本id = [1] order by ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKeyID)

    
    Do Until rsTmp.EOF
        If Dir(App.Path & "\" & rsTmp("ID") & ".cht") = "" Then
             Call LoadImageData(App.Path, rsTmp("ID"))
        End If
        DrawIndex = DrawIndex + 1
        rsTmp.MoveNext
    Loop
    ReadImageData = True
    Exit Function
    
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Function ReportPrint(ByVal lng医嘱ID As Long, ByVal lngKey As Long, ByVal lng病人ID As Long, ByVal blnPrint As Boolean) As Boolean
    '单个报告打印
    
    Dim strReportCode As String
    Dim strReportParaNo As String
    Dim bytReportParaMode As Byte
    Dim rsTmp As New ADODB.Recordset
    Dim blnCurrMoved As Boolean
    Dim strSQL As String
    Dim strChart(1 To 9) As String
    Dim intLoop As Integer
    
    ReportPrint = False
    Me.MousePointer = 11
    zlCommFun.ShowFlash "正在打印请等待...", Me
    
    '生成图形供自定义报表调用
    ReadImageData lngKey
    strSQL = "select id from 检验图像结果 where 标本id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)
    intLoop = 1
    Do Until rsTmp.EOF
        strChart(intLoop) = App.Path & "\" & rsTmp("ID") & ".cht"
        intLoop = intLoop + 1
        rsTmp.MoveNext
    Loop
    
    
    If GetReportCode(lng医嘱ID, lng发送号, strReportCode, strReportParaNo, bytReportParaMode, blnCurrMoved) Then
        Call ReportOpen(gcnOracle, glngSys, strReportCode, Me, "NO=" & strReportParaNo, "性质=" & bytReportParaMode, "医嘱ID=" & lng医嘱ID, _
                        "病人ID=" & lng病人ID, "标本ID=" & lngKey, "多个医嘱=" & lng医嘱ID, "多个标本=" & lngKey, _
                        "图形1=" & strChart(1), "图形2=" & strChart(2), "图形3=" & strChart(3), "图形4=" & strChart(4), _
                        "图形5=" & strChart(5), "图形6=" & strChart(6), "图形7=" & strChart(7), "图形8=" & strChart(8), _
                        "图形9=" & strChart(9), IIf(blnPrint, 2, 1))
    End If
    
    
    On Error GoTo errH

    gstrSQL = " select id from 检验标本记录 where 医嘱id = [1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng医嘱ID)
    Do Until rsTmp.EOF
        strSQL = "ZL_检验标本记录_标本质控(" & rsTmp("ID") & ",'',1)"
        zlDatabase.ExecuteProcedure strSQL, gstrSysName
        rsTmp.MoveNext
    Loop
    
    Me.MousePointer = 0
    zlCommFun.StopFlash
    ReportPrint = True
    On Error Resume Next
    '删除图形文件
    For intLoop = 1 To 9
        Kill strChart(intLoop)
    Next
    
    Exit Function
errH:
    Me.MousePointer = 0
    zlCommFun.StopFlash
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Public Function GetReportCode(ByVal lng医嘱ID As Long, ByVal lng发送号 As Long, ByRef strCode As String, ByRef strNo As String, ByRef bytMode As Byte, Optional ByVal DataMoved As Boolean = False) As Boolean
    '--------------------------------------------------------------------------------------------------------
    '功能;
    '--------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    If lng医嘱ID = 0 And lng发送号 = 0 Then Exit Function
    
'    strSQL = "SELECT DISTINCT 'ZLCISBILL'||Trim(To_Char(C.编号,'00000'))||'-2' AS 报表编号," & _
                       "A.NO," & _
                       "A.记录性质 " & _
                "FROM 病人医嘱发送 A,病历文件列表 C,病人医嘱记录 D,病历单据应用 E " & _
                "Where E.病历文件id = C.ID " & _
                        "AND D.诊疗项目ID=E.诊疗项目ID " & _
                      "AND A.医嘱ID=D.ID AND E.应用场合=Decode(D.病人来源,2,2,4,4,1) " & _
                      " AND D.相关id= [1] "
                      
    strSQL = "Select Distinct 'ZLCISBILL' || Trim(To_Char(C.编号, '00000')) || '-2' As 报表编号, A.NO, A.记录性质, F.ID, F.编码" & vbNewLine & _
            "From 病人医嘱发送 A, 病历文件列表 C, 病人医嘱记录 D, 病历单据应用 E, 诊疗项目目录 F" & vbNewLine & _
            "Where E.病历文件id = C.ID And D.诊疗项目id = E.诊疗项目id And D.诊疗项目id = F.ID And A.医嘱id = D.ID And" & vbNewLine & _
            "      E.应用场合 = Decode(D.病人来源, 2, 2, 4, 4, 1) And D.相关id = [1] " & vbNewLine & _
            "Order By F.编码 "
                          
    If DataMoved Then
        strSQL = Replace(strSQL, "病人医嘱发送", "H病人医嘱发送")
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If

'    strSQL = "SELECT DISTINCT 'ZLCISBILL'||Trim(To_Char(C.编号,'00000'))||'-2' AS 报表编号," & _
'                       "A.NO," & _
'                       "A.记录性质 " & _
'                "FROM 病历单据应用 A,病历文件目录 C,病人医嘱记录 D,病人医嘱发送 B " & _
'                "Where A.病历文件id = C.ID " & _
'                      "AND A.诊疗项目id=D.诊疗项目ID " & _
'                      "AND B.病人ID=D.病人ID " & _
'                      "AND NVL(B.主页ID,0)=NVL(D.主页ID,0) " & _
'                      "AND B.文件id=C.ID " & _
'                      "AND D.相关id=" & lng医嘱id & " " & _
'                      "AND A.发送号=" & lng发送号

    Set rs = zlDatabase.OpenSQLRecord(strSQL, "mdlLISWork", lng医嘱ID, lng发送号)
                      
    
    If rs.BOF = False Then
        strCode = zlCommFun.Nvl(rs("报表编号"))
        strNo = zlCommFun.Nvl(rs("NO"))
        bytMode = zlCommFun.Nvl(rs("记录性质"), 1)
    End If
    
    GetReportCode = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
    ' 2007-08-17 增加一卡通支持
    Dim lngPreIDKind As Long
    If Not TxtID.Locked And TxtID.Text = "" And Me.ActiveControl Is TxtID Then
        TxtID.Text = strID
        Call TxtID_KeyPress(vbKeyReturn)
    End If
End Sub



Private Sub LoadPageItemList(ByVal PageNo As Long)
'功能:加载页面的每一查询项目
'参数:PageNo            页面序号
'说明:这是查询内容显示的主体部份,显示查询内容
    Dim FileName As String
    Dim W As Single
    Dim H As Single
    Dim vFont As New StdFont
    Dim i As Long
    Dim j As Long
    Dim strTmp As String
    Dim rs As New ADODB.Recordset
    Dim vNextY As Single
    Dim vNextX As Single
    Dim objDraw As ctlQueryItem
    Dim vWidth As Single
    Dim vHeight As Single
    Dim vTmp As Single
    Dim vTmp1 As Single
    Dim vMaxWidth As Single
    Dim vVisible As Boolean
    Dim strText As String
    
    On Error GoTo errHand
    i = 1
    vNextY = 60 + (i - 1) * 600
    vNextX = 120
    vMaxWidth = 120
            
    ShowFlatFlash "请稍候，正在生成页面...", Me
    DoEvents
    
    Set objDraw = QueryItem
    objDraw.ClientVisible = False
    Call objDraw.ClearAllPageItem
    
    '读取页面的背景及广告条幅
'    Set gRs = OpenRecord(gRs, "select B.类型,B.名称 from 咨询页面目录 A,咨询图片元素 B where A.宣传标语=B.序号 and A.页面序号=" & PageNo)
'    If gRs.BOF = False Then FrameDefault.AdviceMovie = IIf(IsNull(gRs!名称), "", App.Path & "\图形\" & gRs!名称 & IIf(gRs!类型 <> 2, ".pic", ".swf"))
                    
    '开始生成自定义查询页面
    gstrSQL = "select 页面序号,段落序号,标题文本,标题图标,标题隐藏,标题位置,标题字体,返回页首,段落类型,段落字体,插表序号,插表位置,插图序号,插图位置 from 咨询段落目录 where 页面序号=[1] order by 段落序号"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo)
    If gRs.BOF = False Then
        While Not gRs.EOF
            strTmp = IIf(IsNull(gRs!标题字体), "宋体;12;0;0;0", gRs!标题字体)
            vFont.Name = Split(strTmp, ";")(0)
            vFont.Size = Val(Split(strTmp, ";")(1))
            vFont.Bold = Val(Split(strTmp, ";")(2))
            vFont.Italic = Val(Split(strTmp, ";")(3))
                                    
            FileName = ""
            '1.加载标题内容及标题图标
            vVisible = IIf(IsNull(gRs!标题隐藏), 1, gRs!标题隐藏)
            
            gstrSQL = "select 名称 from 咨询图片元素 where 序号=[1]"
            Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(IIf(IsNull(gRs!标题图标), 0, gRs!标题图标)))
            If rs.BOF = False Then
                FileName = GetFileName(IIf(IsNull(gRs!标题图标), 0, gRs!标题图标), W, H)
            End If
            Call objDraw.AddPageItemTitle(i, vNextY, IIf(IsNull(gRs!标题文本), "", gRs!标题文本), Val(Split(strTmp, ";")(4)), vFont, FileName, PageNo, IIf(IsNull(gRs!段落序号), 0, gRs!段落序号), vWidth, vHeight, Not vVisible, IIf(IsNull(gRs!标题位置), 0, gRs!标题位置))
                                                                                    
            If Not vVisible = True Then vNextY = vNextY + vHeight + 150
            
            
            Select Case zlCommFun.Nvl(gRs("段落类型").Value, 0)
            '----------------------------------------------------------------------------------------------------------
            Case 0          '纯文本内容
                strTmp = IIf(IsNull(gRs!段落字体), "宋体;12;0;0;0", gRs!段落字体)
                j = objDraw.NextTxtIndex
                
                vWidth = QueryItem.Width - 330
                
                strText = Sys.ReadLob(glngSys, 29, PageNo & "," & gRs("段落序号").Value, "", 1)
                
                Call objDraw.AddPageItemTxt(j, vNextX, vNextY, strText & Chr(13) & Chr(10), strTmp, vWidth, vHeight)
                vNextY = vNextY + vHeight + 150
            '----------------------------------------------------------------------------------------------------------
            Case 1          '纯表格内容
                vHeight = 0
                Call InsertGrid(objDraw, IIf(IsNull(gRs!插表序号), 0, gRs!插表序号), vNextX, vNextY, vWidth, vHeight)
                If vHeight > 0 Then vNextY = vNextY + vHeight + 150
            '----------------------------------------------------------------------------------------------------------
            Case 2          '纯图形内容
                FileName = GetFileName(IIf(IsNull(gRs!插图序号), 0, gRs!插图序号), W, H)
                Call objDraw.AddPageItemPic(objDraw.NextPicIndex, 0, vNextY, FileName, vWidth, vHeight, W, H)
                vNextY = vNextY + vHeight + 150
            '----------------------------------------------------------------------------------------------------------
            Case 3          '纯链接内容
                gstrSQL = "select C.页面名称||decode(B.标题文本,NULL,'','：'||B.标题文本) as 标题文本,A.链接页面,A.页内段号 from 咨询段落链接 A,咨询段落目录 B,咨询页面目录 C Where A.链接页面=C.页面序号 and A.链接页面=B.页面序号(+) and A.页内段号=B.段落序号(+) and A.页面序号 = [1] And A.段落序号 = [2]"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo, Val(IIf(IsNull(gRs!段落序号), 0, gRs!段落序号)))
                If rs.BOF = False Then
                    While Not rs.EOF
                        Call objDraw.AddPageItemConnect(objDraw.NextConnectIndex, vNextX + 150, vNextY, IIf(IsNull(rs!标题文本), "", rs!标题文本), IIf(IsNull(rs!链接页面), 0, rs!链接页面), IIf(IsNull(rs!页内段号), 0, rs!页内段号), vWidth, vHeight)
                        vNextY = vNextY + 300
                        rs.MoveNext
                    Wend
                    vNextY = vNextY + 150
                Else
                    '检查是否连接到ZLHIS的人员
                    gstrSQL = "select B.姓名,A.链接页面,A.页内段号 from 咨询段落链接 A,人员表 B Where A.页内段号=B.id And (b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or b.撤档时间 Is Null) and A.页面序号 = [1] And A.段落序号 = [2]"
                    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo, Val(IIf(IsNull(gRs!段落序号), 0, gRs!段落序号)))
                    If rs.BOF = False Then
                        While Not rs.EOF
                            Call objDraw.AddPageItemConnect(objDraw.NextConnectIndex, vNextX, vNextY, IIf(IsNull(rs!姓名), "", rs!姓名), IIf(IsNull(rs!链接页面), 0, rs!链接页面), IIf(IsNull(rs!页内段号), 0, rs!页内段号), vWidth, vHeight)
                            vNextY = vNextY + 300
                            rs.MoveNext
                        Wend
                        vNextY = vNextY + 150
                    End If
                End If

            '----------------------------------------------------------------------------------------------------------
            Case 4          '文本和表格
                strTmp = IIf(IsNull(gRs!段落字体), "宋体;12;0;0;0", gRs!段落字体)
                j = objDraw.NextTxtIndex
                strText = Sys.ReadLob(glngSys, 29, PageNo & "," & gRs("段落序号").Value, "", 1)
                Select Case IIf(IsNull(gRs!插表位置), 0, gRs!插表位置)
                Case 0
                    vHeight = 0
                    Call InsertGrid(objDraw, IIf(IsNull(gRs!插表序号), 0, gRs!插表序号), 0, vNextY, vTmp1, vTmp)
                    vWidth = QueryItem.Width - vTmp1 - 120 - 120
                    Call objDraw.AddPageItemTxt(j, vNextX + vTmp1 + 60, vNextY, strText & Chr(13) & Chr(10), strTmp, vWidth, vHeight)
                Case 1
                    Call objDraw.AddPageItemTxt(j, vNextX, vNextY, strText & Chr(13) & Chr(10), strTmp, vTmp1, vHeight)
                    Call InsertGrid(objDraw, IIf(IsNull(gRs!插表序号), 0, gRs!插表序号), 1, vNextY, vWidth, vTmp)
                End Select
                vNextY = vNextY + IIf(vTmp > vHeight, vTmp, vHeight) + 150
            '----------------------------------------------------------------------------------------------------------
            Case 5          '文本和图形
                FileName = GetFileName(IIf(IsNull(gRs!插图序号), 0, gRs!插图序号), W, H)
                strText = Sys.ReadLob(glngSys, 29, PageNo & "," & gRs("段落序号").Value, "", 1)
                strTmp = IIf(IsNull(gRs!段落字体), "宋体;12;0;0;0", gRs!段落字体)
                j = objDraw.NextTxtIndex
                Select Case IIf(IsNull(gRs!插图位置), 0, gRs!插图位置)
                Case 0
                    Call objDraw.AddPageItemPic(objDraw.NextPicIndex, 0, vNextY, FileName, vTmp1, vTmp, W, H)
                    vWidth = QueryItem.Width - vTmp1 - 120 - 120
                    Call objDraw.AddPageItemTxt(j, vNextX + vTmp1 + 60, vNextY, strText & Chr(13) & Chr(10), strTmp, vWidth, vHeight)
                Case 1
                    Call objDraw.AddPageItemPic(objDraw.NextPicIndex, 1, vNextY, FileName, vWidth, vTmp, W, H)
                    vTmp1 = QueryItem.Width - vWidth - 60 - 90
                    Call objDraw.AddPageItemTxt(j, vNextX, vNextY, strText & Chr(13) & Chr(10), strTmp, vTmp1, vHeight)
                End Select
                vNextY = vNextY + IIf(vTmp > vHeight, vTmp, vHeight) + 150
            End Select
                        
            '8.设置返回页首标志
            If IIf(IsNull(gRs!返回页首), 0, gRs!返回页首) = 1 Then
                vHeight = 0
                Call objDraw.AddReturnFlag(vNextX, vNextY, vHeight)
                If vHeight > 0 Then vNextY = vNextY + vHeight + 150
            End If
            
            i = i + 1
            gRs.MoveNext
        Wend
    End If
        
    Call objDraw.ResizePage(QueryItem.Width, vNextY)
    QueryItem.Height = vNextY
    'Call FrameDefault.InitNavigator(FrameDefault.ClientWidth, vNextY)
    
    '获取背景并画出页面背景
    gstrSQL = "select B.类型,B.名称,B.宽度,B.高度 from 咨询页面目录 A,咨询图片元素 B where A.页面背景=B.序号 and A.页面序号=[1]"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo)
    If gRs.BOF = False Then
        Call objDraw.BackPicture(IIf(IsNull(gRs!名称), "", App.Path & "\图形\" & gRs!名称 & IIf(gRs!类型 <> 2, ".pic", ".swf")), IIf(IsNull(gRs!宽度), 0, gRs!宽度) * Screen.TwipsPerPixelX, IIf(IsNull(gRs!高度), 0, gRs!高度) * Screen.TwipsPerPixelY)
    End If
            
    Call objDraw.InitLoad
    objDraw.ClientVisible = True
    
    StopFlatFlash
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub hsb_Change()
    QueryItem.Left = mvarLeftStart - hsb.Value * 360
    If QueryItem.Left + QueryItem.Width < picBack.Left + picBack.Width - vsb.Width Then
        QueryItem.Left = picBack.Left + picBack.Width - QueryItem.Width - vsb.Width
    End If
    If QueryItem.Left > 0 Then QueryItem.Left = picBack.Width - vsb.Width - QueryItem.Width
End Sub

Private Sub hsb_KeyDown(KeyCode As Integer, Shift As Integer)
    Call picBack_KeyDown(KeyCode, Shift)
End Sub

Private Sub picBack_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If vsb.Enabled Then vsb.Value = IIf(vsb.Value < vsb.Max, vsb.Value + 1, vsb.Max)
    End If

    If KeyCode = vbKeyUp Then
        If vsb.Enabled Then vsb.Value = IIf(vsb.Value > 0, vsb.Value - 1, 0)
    End If

    If KeyCode = vbKeyRight Then
        If hsb.Enabled Then hsb.Value = IIf(hsb.Value < hsb.Max, hsb.Value + 1, hsb.Max)
    End If

    If KeyCode = vbKeyLeft Then
        If hsb.Enabled Then hsb.Value = IIf(hsb.Value > 0, hsb.Value - 1, 0)
    End If
End Sub

Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call InitInternal
End Sub

Private Sub picBack_Paint()
    Call RaisEffect(picBack, -1)
End Sub

Private Sub picBack1_KeyDown(KeyCode As Integer, Shift As Integer)
    Call picBack_KeyDown(KeyCode, Shift)
End Sub

Private Sub picBack1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call InitInternal
End Sub

Private Sub QueryItem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call InitInternal
End Sub

Private Sub vsb_Change()
    QueryItem.Top = 0 - vsb.Value * 360
    If QueryItem.Top + QueryItem.Height < picBack.Height - hsb.Height Then
        QueryItem.Top = picBack.Top + picBack.Height - hsb.Height - QueryItem.Height
    End If
    If QueryItem.Top > 30 Then QueryItem.Top = picBack.Height - hsb.Height - QueryItem.Height
    
End Sub

Private Sub CalcVsb()
    vsb.Max = 0 - Int(0 - (QueryItem.Height - picBack.ScaleHeight + hsb.Height + 45) / 360)
    If vsb.Max > 0 Then
        vsb.Enabled = True
        vsb.Visible = True
        vsb.SmallChange = 1
        vsb.LargeChange = 1
        vsb.Value = 0
        hsb.Width = picBack.Width - hsb.Width
    Else
        vsb.Enabled = False
        vsb.Visible = False
        hsb.Width = picBack.Width
    End If
    
    hsb.Max = 0 - Int(0 - (QueryItem.Width - picBack.ScaleWidth + vsb.Width + 45) / 360)
    If hsb.Max > 0 Then
        hsb.Enabled = True
        hsb.Visible = True
        hsb.SmallChange = 1
        hsb.LargeChange = 1
        hsb.Value = 0
        vsb.Height = picBack.Height - hsb.Height
    Else
        hsb.Enabled = False
        hsb.Visible = False
        vsb.Height = picBack.Height
    End If
End Sub

Private Sub vsb_KeyDown(KeyCode As Integer, Shift As Integer)
    Call picBack_KeyDown(KeyCode, Shift)
End Sub

