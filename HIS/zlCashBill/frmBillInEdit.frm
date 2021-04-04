VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmBillInEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "票据入库编辑"
   ClientHeight    =   6600
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8805
   Icon            =   "frmBillInEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   420
      Left            =   7455
      TabIndex        =   22
      Top             =   5580
      Width           =   1200
   End
   Begin VB.Frame fraUse 
      Caption         =   "入库基本信息"
      Height          =   2490
      Left            =   135
      TabIndex        =   21
      Top             =   390
      Width           =   6990
      Begin VB.ComboBox cbo类别 
         Height          =   300
         Left            =   1110
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   390
         Width           =   2670
      End
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   360
         Index           =   7
         Left            =   4605
         MaxLength       =   20
         TabIndex        =   15
         Top             =   1785
         Width           =   2265
      End
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   360
         Index           =   6
         Left            =   1110
         MaxLength       =   20
         TabIndex        =   13
         Top             =   1785
         Width           =   2655
      End
      Begin VB.TextBox txtEdit 
         Height          =   330
         Index           =   0
         Left            =   4935
         MaxLength       =   20
         TabIndex        =   3
         Top             =   375
         Width           =   1950
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         IMEMode         =   2  'OFF
         Index           =   4
         Left            =   4605
         MaxLength       =   20
         TabIndex        =   9
         Top             =   855
         Width           =   2295
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         IMEMode         =   2  'OFF
         Index           =   3
         Left            =   4215
         MaxLength       =   2
         TabIndex        =   8
         Top             =   855
         Width           =   375
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         IMEMode         =   2  'OFF
         Index           =   2
         Left            =   1515
         MaxLength       =   20
         TabIndex        =   6
         Top             =   855
         Width           =   2295
      End
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   360
         Index           =   5
         Left            =   1110
         MaxLength       =   20
         TabIndex        =   11
         Top             =   1320
         Width           =   5775
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         IMEMode         =   2  'OFF
         Index           =   1
         Left            =   1125
         MaxLength       =   2
         TabIndex        =   5
         Top             =   855
         Width           =   375
      End
      Begin VB.Label lblUserType 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "使用类别(&L)"
         Height          =   180
         Left            =   120
         TabIndex        =   0
         Top             =   450
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入库批次"
         Height          =   180
         Index           =   7
         Left            =   4200
         TabIndex        =   2
         Top             =   450
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "～"
         Height          =   240
         Index           =   5
         Left            =   3945
         TabIndex        =   7
         Top             =   945
         Width           =   240
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "号码范围(&B)"
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   4
         Top             =   945
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "登记时间"
         Height          =   180
         Index           =   3
         Left            =   3870
         TabIndex        =   14
         Top             =   1875
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "登记人"
         Height          =   180
         Index           =   2
         Left            =   540
         TabIndex        =   12
         Top             =   1875
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "备注(&G)"
         Height          =   180
         Index           =   0
         Left            =   480
         TabIndex        =   10
         Top             =   1410
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      Height          =   6735
      Left            =   7215
      TabIndex        =   20
      Top             =   -15
      Width           =   30
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   420
      Left            =   7455
      TabIndex        =   19
      Top             =   690
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   420
      Left            =   7455
      TabIndex        =   18
      Top             =   210
      Width           =   1200
   End
   Begin VSFlex8Ctl.VSFlexGrid vsMemo 
      Height          =   3150
      Left            =   150
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   3330
      Width           =   6945
      _cx             =   12250
      _cy             =   5556
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483643
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   2
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmBillInEdit.frx":058A
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   1
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
   Begin VB.Label Label2 
      Caption         =   "详细情况"
      Height          =   255
      Left            =   135
      TabIndex        =   16
      Top             =   3090
      Width           =   975
   End
End
Attribute VB_Name = "frmBillInEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum BillInEditType
    Ed_增加 = 0
    Ed_修改 = 1
    Ed_查看 = 2
End Enum
Private mstrPrivs As String, mlngModule As Long
Private mEditType As BillInEditType '编辑类型
Private mblnChange As Boolean     '为真时表示已改变了
Private mstr票据长度 As String '表示各种票据的号码长度，各位分别为1-收费,2-预交,3-结帐,4-挂号,5-就诊卡  77777
Private mlng长度 As Long       '当前票据种类的长度
Private mintSucceed As Integer
Private mlng入库ID  As Long
Private mstrDrawBill As String, mstrDrawNum As String '领用分段信息
Private mstrDamnifyBill As String, mlngDamnifyNum As Long  '领用分段信息,报损数量合计
Private mint票种 As gBillType  '票种
Private mblnIsBIll As Boolean '当前票种是否为票据
Private mblnFirst As Boolean
Private mstr类别 As String '缺省传入类别
Private mstrPreType(1 To 6) As String
Private mcllCardProperty As Collection
Private mblnNotClick As Boolean
Private Enum mTxtIdx
    idx_批次 = 0
    idx_开始前缀 = 1
    idx_开始号码 = 2
    idx_终止前缀 = 3
    idx_终止号码 = 4
    idx_备注 = 5
    idx_登记人 = 6
    idx_登记时间 = 7
End Enum
Public Function zlBillEdit(ByVal frmMain As Form, ByVal int票种 As gBillType, ByVal EditType As BillInEditType, ByVal strPrivs As String, _
    ByVal lngModule As Long, Optional ByVal lng入库ID As Long = 0, Optional str类别 As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入库,票据入库查询或操作功能(包含增加和修改)
    '入参:frmMain-调用主窗体
    '       BillEditType-单据操作类型
    '       strPrivs-权限串
    '       lngModule-模块号
    '       lng入库ID-修改时,转入的入库批次.
    '       str类别:使用类别名称(27559)
    '出参:
    '返回:操作一张以上成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-11-16 10:29:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String
    mstr类别 = str类别: mint票种 = int票种: mblnIsBIll = CurrentIsBill(int票种)
    mstrPreType(mint票种) = mstr类别
    mEditType = EditType: mstrPrivs = strPrivs: mlngModule = lngModule: mlng入库ID = lng入库ID
    mstr票据长度 = zlDatabase.GetPara(20, glngSys, , "7|7|7|7|7")
    mintSucceed = False
    If mEditType = Ed_查看 Then
        cmdOK.Caption = "退出(&O)"
        cmdCancel.Visible = False
    End If
    strTemp = Decode(mint票种, gBillType.收费收据, "收费收据", gBillType.预交收据, "预交收据", gBillType.结帐收据, "结帐收据", _
        gBillType.挂号收据, "挂号收据", gBillType.就诊卡, "就诊卡", gBillType.消费卡, "消费卡", "就诊卡")
    Me.Caption = strTemp & "入库"
    fraUse.Caption = "『" & strTemp & "』入库基本信息"
    Me.Show 1, frmMain
    zlBillEdit = mintSucceed > 0
End Function
Private Function LoadCardData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载卡片数据
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-11-16 10:35:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, lngLen As Long
    Dim i As Long, blnFind As Boolean
    
    If mint票种 = gBillType.就诊卡 Or mint票种 = gBillType.消费卡 Then
         mlng长度 = mcllCardProperty(cbo类别.ListIndex + 1)(0)
    Else
        mlng长度 = Split(mstr票据长度, "|")(mint票种 - 1)
    End If
    If UserInfo.姓名 = "" Then
        MsgBox "你还未设置人员的对照关系，请与系统管理员联系，设置后才能使用本功能。", vbExclamation, gstrSysName
        Exit Function
    End If
    
    Call ClearData  '清除控件数据
    Err = 0: On Error GoTo errHandle
    If mEditType = Ed_增加 Then
        If mint票种 = gBillType.就诊卡 Or mint票种 = gBillType.消费卡 Then
            txtEdit(mTxtIdx.idx_开始前缀).Text = mcllCardProperty(cbo类别.ListIndex + 1)(1)
        End If
        txtEdit(mTxtIdx.idx_登记人) = UserInfo.姓名
        txtEdit(mTxtIdx.idx_登记时间) = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
        
        If mint票种 = gBillType.消费卡 Or mint票种 = gBillType.消费卡 Then
            Call Set前缀(mcllCardProperty(cbo类别.ListIndex + 1)(1))
        End If
        LoadCardData = True
        Exit Function
    End If
    
    If mint票种 = gBillType.消费卡 Then
        gstrSQL = _
            "Select Id, 接口编号 As 使用类别, 前缀文本, 开始卡号 As 开始号码, 终止卡号 As 终止号码, 入库数量, 剩余数量, 备注, 登记人, 登记时间, 批次  " & _
            "From 消费卡入库记录 " & _
            "Where Id=[1]"
    Else
        gstrSQL = _
            "Select Id, 使用类别, 前缀文本, 开始号码, 终止号码, 入库数量, 剩余数量, 备注, 登记人, 登记时间, 批次  " & _
            "From 票据入库记录 " & _
            "Where Id=[1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng入库ID)
    If rsTemp.RecordCount = 0 Then
        MsgBox "注意:" & vbCrLf & "    该批次的入库" & IIf(mblnIsBIll, "票据", "卡片") & "已经被他人删除，请检查！", vbOKOnly + vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    
    With cbo类别
        blnFind = False
        For i = 0 To .ListCount - 1
            If mint票种 = gBillType.预交收据 Then
                 If .ItemData(i) = Val(NVL(rsTemp!使用类别)) + 1 Then
                    .ListIndex = i: blnFind = True: Exit For
                 End If
            ElseIf mint票种 = gBillType.就诊卡 Or mint票种 = gBillType.消费卡 Then
                 If .ItemData(i) = Val(NVL(rsTemp!使用类别)) Then
                    .ListIndex = i: blnFind = True: Exit For
                 End If
            Else
                If .List(i) = NVL(rsTemp!使用类别) Then
                   .ListIndex = i: blnFind = True: Exit For
                End If
            End If
        Next
        '58071
        If blnFind = False And Not (mint票种 = gBillType.就诊卡 Or mint票种 = gBillType.消费卡) Then
            .AddItem NVL(rsTemp!使用类别, " ")
            .ListIndex = .NewIndex
        End If
        .Tag = NVL(rsTemp!使用类别, " ")
        .Enabled = NVL(rsTemp!入库数量) = NVL(rsTemp!剩余数量)
    End With
    
    txtEdit(mTxtIdx.idx_批次).Text = NVL(rsTemp!批次)
    txtEdit(mTxtIdx.idx_批次).Tag = NVL(rsTemp!批次)
    txtEdit(mTxtIdx.idx_开始前缀).Text = NVL(rsTemp!前缀文本)
    lngLen = Len(Trim(txtEdit(mTxtIdx.idx_开始前缀).Text))
    txtEdit(mTxtIdx.idx_开始号码).Text = Mid(NVL(rsTemp!开始号码), lngLen + 1)
    txtEdit(mTxtIdx.idx_开始号码).Tag = txtEdit(mTxtIdx.idx_开始号码).Text
    txtEdit(mTxtIdx.idx_终止前缀).Text = NVL(rsTemp!前缀文本)
    txtEdit(mTxtIdx.idx_终止号码).Text = Mid(NVL(rsTemp!终止号码), lngLen + 1)
    txtEdit(mTxtIdx.idx_终止号码).Tag = txtEdit(mTxtIdx.idx_终止号码).Text
    txtEdit(mTxtIdx.idx_备注).Text = NVL(rsTemp!备注)
    txtEdit(mTxtIdx.idx_登记人).Text = NVL(rsTemp!登记人)
    txtEdit(mTxtIdx.idx_登记时间).Text = Format(rsTemp!登记时间, "yyyy-mm-dd HH:MM:SS")
    
    '加载详细其他信息
    vsMemo.Tag = Val(NVL(rsTemp!入库数量)) & "-" & Val(NVL(rsTemp!剩余数量))
    If mint票种 = gBillType.消费卡 Then
        gstrSQL = _
            "Select A.登记时间,A.开始卡号 As 开始号码,A.终止卡号 As 终止号码 " & _
            "From 消费卡领用记录 A " & _
            "Where A.批次=[1] " & _
            "Order By 登记时间"
    Else
        gstrSQL = _
            "Select A.登记时间,A.开始号码,A.终止号码 " & _
            "From 票据领用记录 A " & _
            "Where A.批次=[1] " & _
            "Order By 登记时间"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng入库ID)
    With rsTemp
        mstrDrawNum = "0"
        Do While Not .EOF
            If NVL(rsTemp!开始号码) = NVL(rsTemp!终止号码) Then
                mstrDrawBill = mstrDrawBill & "," & NVL(rsTemp!开始号码)
            Else
                mstrDrawBill = mstrDrawBill & "," & NVL(rsTemp!开始号码) & "-" & NVL(rsTemp!终止号码)
            End If
            'mstrDrawNum = mlngDrawNum + GetBillNum(Mid(Nvl(rsTemp!开始号码), lngLen + 1), Mid(Nvl(rsTemp!终止号码), lngLen + 1))
            '问题号:54259
            '77390:李南春,2014/9/3 09:33:32,计算票据数量
             mstrDrawNum = zlStr.ExpressValue(GetBillNum(Mid(NVL(rsTemp!开始号码), lngLen + 1), Mid(NVL(rsTemp!终止号码), lngLen + 1)) & "+" & mstrDrawNum)
            .MoveNext
        Loop
        If mstrDrawBill <> "" Then mstrDrawBill = Mid(mstrDrawBill, 2)
    End With
    
    '报损信息
    If mint票种 = gBillType.消费卡 Then
        gstrSQL = _
            "Select A.终止卡号 As 终止号码, A.开始卡号 As 开始号码,A.报损时间,A.数量 " & _
            "From 消费卡报损记录 A " & _
            "Where 入库ID=[1] " & _
            "Order By 开始号码,报损时间"
    Else
        gstrSQL = _
            "Select A.终止号码, A.开始号码,A.报损时间,A.数量 " & _
            "From 票据报损记录 A " & _
            "Where 入库ID=[1] " & _
            "Order By 开始号码,报损时间"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng入库ID)
    mstrDamnifyBill = ""
     With rsTemp
        mlngDamnifyNum = 0
        Do While Not .EOF
            If NVL(rsTemp!开始号码) = NVL(rsTemp!终止号码) Then
                mstrDamnifyBill = mstrDamnifyBill & "," & NVL(rsTemp!开始号码)
            ElseIf NVL(rsTemp!开始号码) = "" And NVL(rsTemp!终止号码) <> "" Then
                mstrDamnifyBill = mstrDamnifyBill & "," & NVL(rsTemp!终止号码)
            ElseIf NVL(rsTemp!开始号码) <> "" And NVL(rsTemp!终止号码) = "" Then
                mstrDamnifyBill = mstrDamnifyBill & "," & NVL(rsTemp!开始号码)
            Else
                mstrDamnifyBill = mstrDamnifyBill & "," & NVL(rsTemp!开始号码) & "-" & NVL(rsTemp!终止号码)
            End If
            mlngDamnifyNum = mlngDamnifyNum + Val(NVL(rsTemp!数量))
            .MoveNext
        Loop
        If mstrDamnifyBill <> "" Then mstrDamnifyBill = Mid(mstrDamnifyBill, 2)
    End With
    Call SetCtlEnable
    Call SetMemo
    If mint票种 = gBillType.就诊卡 Or mint票种 = gBillType.消费卡 Then
        Call Set前缀(mcllCardProperty(cbo类别.ListIndex + 1)(1))
    End If
    LoadCardData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub SetCtlEnable()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件可编辑属性
    '编制:刘兴洪
    '日期:2010-11-17 16:03:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Select Case mEditType
    Case Ed_增加
        txtEdit(mTxtIdx.idx_开始前缀).Enabled = True
        txtEdit(mTxtIdx.idx_终止前缀).Enabled = True
        txtEdit(mTxtIdx.idx_开始号码).Enabled = True
        txtEdit(mTxtIdx.idx_终止号码).Enabled = True
        txtEdit(mTxtIdx.idx_备注).Enabled = True
    Case Ed_修改
       'If mlng长度 > 2 Then
            txtEdit(mTxtIdx.idx_开始前缀).Enabled = True
            txtEdit(mTxtIdx.idx_终止前缀).Enabled = True
        ' End If
        txtEdit(mTxtIdx.idx_开始号码).Enabled = True
        txtEdit(mTxtIdx.idx_终止号码).Enabled = True
        txtEdit(mTxtIdx.idx_备注).Enabled = True
        If mstrDamnifyBill <> "" Or mstrDrawBill <> "" Then
            '不能更改前缀
            txtEdit(mTxtIdx.idx_开始前缀).Enabled = False: txtEdit(mTxtIdx.idx_终止前缀).Enabled = False:
        End If
    Case Else
        For i = 0 To txtEdit.UBound
            txtEdit(i).Enabled = False
        Next
    End Select
End Sub


Private Sub SetMemo()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置说明信息
    '编制:刘兴洪
    '日期:2010-11-16 10:55:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, sngY As Single, intTittleFontSize As Integer, intTextFontSize As Integer
    Dim strTmp As String, strTemp As String, strText As String, i As Long
    Dim varTemp As Variant
    
    With vsMemo
        .Redraw = flexRDNone
        .Clear
        lngRow = 1
        '-----------------------------------------------------------------------
        '入库票据处理
        .Cell(flexcpFontBold, lngRow, 0, lngRow, .Cols - 1) = True  '初体显示
        .Cell(flexcpFontSize, lngRow, 0, lngRow, .Cols - 1) = intTittleFontSize  '初体显示
        .Cell(flexcpText, lngRow, 0, lngRow, .Cols - 1) = "入库:"
        
        lngRow = lngRow + 1
        .Cell(flexcpFontBold, lngRow, 0, lngRow, .Cols - 1) = False  '
        .Cell(flexcpFontSize, lngRow, 0, lngRow, .Cols - 1) = intTextFontSize
        .Cell(flexcpText, lngRow, 1, lngRow, .Cols - 1) = IIf(mblnIsBIll, "票号范围:", "卡号范围:") & Trim(txtEdit(mTxtIdx.idx_开始前缀).Text) & Trim(txtEdit(mTxtIdx.idx_开始号码)) & "至" & Trim(txtEdit(mTxtIdx.idx_开始前缀).Text) & Trim(txtEdit(mTxtIdx.idx_终止号码))
        '问题号:54259
        strTmp = "0"
        If mEditType = Ed_查看 Then
            strText = Val(Split(vsMemo.Tag & "-", "-")(0))
        Else
            strTmp = GetBillNum(Trim(txtEdit(mTxtIdx.idx_开始号码)), Trim(txtEdit(mTxtIdx.idx_终止号码)), strTemp)
            strText = strTmp
            If strTemp <> "" Then
                strText = strTemp
            End If
        End If
        
        If Not IsNumeric(strText) Then
            .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbRed
        Else
            .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = .ForeColor
        End If
        lngRow = lngRow + 1
        .Cell(flexcpFontBold, lngRow, 0, lngRow, .Cols - 1) = False  '
        .Cell(flexcpFontSize, lngRow, 0, lngRow, .Cols - 1) = intTextFontSize
        .Cell(flexcpText, lngRow, 1, lngRow, .Cols - 1) = "入库张数:" & strText & "张"
        
        lngRow = lngRow + 1
        If mEditType = Ed_增加 Then GoTo goEnd:
        varTemp = Split(vsMemo.Tag & "-", "-")
        strText = Val(varTemp(1))
        '问题号:54259
        If strTmp <> "0" Then    '修改时,可能剩余张数要发生变化
'            lngTemp = lngTemp - (Val(varTemp(0)) - Val(varTemp(1)))
'            strText = lngTemp
            '77390:李南春,2014/9/3 09:33:32,计算票据数量
            strTmp = GetBillNum(GetBillNum(varTemp(1), varTemp(0)), strTmp)
            strText = strTmp
        End If
        .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = .ForeColor
        .Cell(flexcpFontBold, lngRow, 0, lngRow, .Cols - 1) = False  '
        .Cell(flexcpFontSize, lngRow, 0, lngRow, .Cols - 1) = intTextFontSize
        .Cell(flexcpText, lngRow, 1, lngRow, .Cols - 1) = "剩余张数:" & strText & "张"
        If Val(strText) < 0 Then
            .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbRed
        Else
            .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = .ForeColor
        End If
        '-----------------------------------------------------------------------
        '2.领用票据处理
        lngRow = lngRow + 1
        .Cell(flexcpFontBold, lngRow, 0, lngRow, .Cols - 1) = True  '初体显示
        .Cell(flexcpFontSize, lngRow, 0, lngRow, .Cols - 1) = intTittleFontSize  '初体显示
        .Cell(flexcpText, lngRow, 0, lngRow, .Cols - 1) = "领用:"
        
        lngRow = lngRow + 1
        .Cell(flexcpFontBold, lngRow, 0, lngRow, .Cols - 1) = False  '
        .Cell(flexcpFontSize, lngRow, 0, lngRow, .Cols - 1) = intTextFontSize
        .Cell(flexcpText, lngRow, 1, lngRow, .Cols - 1) = IIf(mblnIsBIll, "领用票据:", "领用卡号:") & mstrDrawBill
        
        lngRow = lngRow + 1
        .Cell(flexcpFontBold, lngRow, 0, lngRow, .Cols - 1) = False  '
        .Cell(flexcpFontSize, lngRow, 0, lngRow, .Cols - 1) = intTextFontSize
        .Cell(flexcpText, lngRow, 1, lngRow, .Cols - 1) = "领用张数:" & mstrDrawNum & "张" '问题号:54259
              
      '-----------------------------------------------------------------------
        '3.报损票据处理
        lngRow = lngRow + 1
        .Cell(flexcpFontBold, lngRow, 0, lngRow, .Cols - 1) = True  '初体显示
        .Cell(flexcpFontSize, lngRow, 0, lngRow, .Cols - 1) = intTittleFontSize  '初体显示
        .Cell(flexcpText, lngRow, 0, lngRow, .Cols - 1) = "报损:"
        
        lngRow = lngRow + 1
        .Cell(flexcpFontBold, lngRow, 0, lngRow, .Cols - 1) = False
        .Cell(flexcpFontSize, lngRow, 0, lngRow, .Cols - 1) = intTextFontSize
        .Cell(flexcpText, lngRow, 1, lngRow, .Cols - 1) = IIf(mblnIsBIll, "报损票据:", "报损卡号:") & mstrDamnifyBill
        
        lngRow = lngRow + 1
        .Cell(flexcpFontBold, lngRow, 0, lngRow, .Cols - 1) = False  '
        .Cell(flexcpFontSize, lngRow, 0, lngRow, .Cols - 1) = intTextFontSize
        .Cell(flexcpText, lngRow, 1, lngRow, .Cols - 1) = "报损张数:" & mlngDamnifyNum & "张"
goEnd:
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 1, 1
        .MergeCells = flexMergeFree
        For i = 0 To .Rows - 1
            .MergeRow(i) = True
        Next
        For i = 0 To .Cols - 1
            .MergeCol(i) = True
        Next
        .Redraw = flexRDBuffered
    End With
End Sub

Private Function GetBillNum(ByVal str开始号码 As String, ByVal str终卡号码 As String, Optional ByRef strErrMsg As String = "") As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取票据张数
    '入参:str开始号码-必须为数字;
    '       str终卡号码-必须为数字
    '出参:strErrMsg-返回错误的计算信息
    '返回:票据总张数
    '编制:刘兴洪
    '日期:2010-11-16 11:06:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo errHandle
    strErrMsg = ""
'    If (str开始号码 <> "" And str终卡号码 = "") Or (str开始号码 = "" And str终卡号码 <> "") Then
'        GetBillNum = 1: Exit Function
'    End If
'    GetBillNum = CDec(str终卡号码) - CDec(str开始号码) + 1

    GetBillNum = zlStr.ExpressValue(str终卡号码 & "-" & str开始号码) + 1
    Exit Function
errHandle:
    strErrMsg = "计算错误或超出了计算范围"
    GetBillNum = "0"
End Function


Private Sub ClearData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除控件数据
    '编制:刘兴洪
    '日期:2010-11-16 10:35:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    mstrDrawBill = "": mstrDamnifyBill = ""

    For i = 0 To txtEdit.UBound
        txtEdit(i).Text = ""
        If i = mTxtIdx.idx_开始前缀 Or i = mTxtIdx.idx_终止前缀 Then
            Call txtEdit_Change(i)  '问题:38021
        End If
        If txtEdit(i).Enabled = False Then
            txtEdit(i).BackColor = Me.BackColor
        Else
            txtEdit(i).BackColor = &H80000005
        End If
    Next
    
    vsMemo.Clear
    vsMemo.Rows = 11
End Sub

Private Sub cbo类别_Click()
    If mint票种 = gBillType.就诊卡 Or mint票种 = gBillType.消费卡 Then
        mlng长度 = mcllCardProperty(cbo类别.ListIndex + 1)(0)
        Call Set前缀(mcllCardProperty(cbo类别.ListIndex + 1)(1))
        If mlng长度 < 3 Then
            txtEdit(mTxtIdx.idx_开始前缀).Text = "": txtEdit(mTxtIdx.idx_开始前缀).Enabled = False
            txtEdit(mTxtIdx.idx_终止前缀).Enabled = False
        End If
        txtEdit(mTxtIdx.idx_开始号码).MaxLength = mlng长度 - zlCommFun.ActualLen(txtEdit(mTxtIdx.idx_开始前缀).Text)
        txtEdit(mTxtIdx.idx_终止号码).MaxLength = txtEdit(mTxtIdx.idx_开始号码).MaxLength
    End If
End Sub

Private Sub Set前缀(Optional ByVal str前缀 As String = "")
    Me.txtEdit(mTxtIdx.idx_开始前缀).Enabled = str前缀 <> ""
    Me.txtEdit(mTxtIdx.idx_终止前缀).Enabled = Me.txtEdit(mTxtIdx.idx_开始前缀).Enabled
    Me.txtEdit(mTxtIdx.idx_开始前缀).BackColor = Me.txtEdit(mTxtIdx.idx_开始号码).BackColor
    Me.txtEdit(mTxtIdx.idx_终止前缀).BackColor = Me.txtEdit(mTxtIdx.idx_开始号码).BackColor
    If str前缀 = "" And mlng长度 > 2 Then Exit Sub
    Me.txtEdit(mTxtIdx.idx_开始前缀).Text = UCase(str前缀)
    Me.txtEdit(mTxtIdx.idx_开始前缀).BackColor = Me.BackColor
    Me.txtEdit(mTxtIdx.idx_终止前缀).BackColor = Me.BackColor
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    If LoadCombox = False Then Unload Me: Exit Sub
    mblnFirst = False
    Call SetCtlEnable
    If LoadCardData = False Then Unload Me: Exit Sub
    If zlControl.IsCtrlSetFocus(txtEdit(mTxtIdx.idx_开始前缀)) Then
        txtEdit(mTxtIdx.idx_开始前缀).SetFocus
    Else
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_开始号码)
    End If
    
    mblnChange = False
End Sub

Private Sub Form_Load()
    mblnFirst = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查输入的数据是否合法
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-11-16 15:04:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str最小号码 As String, str最大号码 As String, varTemp As Variant, varData As Variant
    Dim str开始号码 As String, str结束号码 As String, i As Long, strTemp As String
    Dim str类别 As String, str入库数量 As String
    Dim rsTemp As ADODB.Recordset
    Dim byt发卡控制 As Byte, blnDefult As Boolean
    Dim strName1 As String, strName2 As String
    
    On Error GoTo errHandle
    strName1 = IIf(mblnIsBIll, "票据", "卡片")
    strName2 = IIf(mblnIsBIll, "号码", "卡号")
    '问题号:54259
    If Len(GetBillNum(Trim(txtEdit(mTxtIdx.idx_开始号码)), Trim(txtEdit(mTxtIdx.idx_终止号码)))) > 25 Then
        ShowMsgbox "注意" & vbCrLf & "    入库" & strName1 & "数量位数不得超过" & 25 & "位!"
        Exit Function
    End If
    
    If zlCommFun.ActualLen(Trim(txtEdit(mTxtIdx.idx_备注))) > 200 Then
        ShowMsgbox "注意" & vbCrLf & "    备注最多只能输入200个字符或100个汉字,请检查!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_备注): Exit Function
    End If
    If zlCommFun.ActualLen(Trim(txtEdit(mTxtIdx.idx_开始前缀))) > 2 Then
        ShowMsgbox "注意" & vbCrLf & "   " & strName2 & "前缀最多只能输入2个字符或1个汉字,请检查!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_开始前缀): Exit Function
    End If
    If InStr(1, txtEdit(mTxtIdx.idx_备注), "'") > 0 Then
        ShowMsgbox "注意" & vbCrLf & "    备注中含有非法字符单引号,请检查!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_备注): Exit Function
    End If
    If Trim(txtEdit(mTxtIdx.idx_开始号码).Text) = "" Then
        ShowMsgbox "注意" & vbCrLf & "    " & strName2 & "范围中的开始" & strName2 & "必须输入,请检查!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_开始号码): Exit Function
    End If
    If Trim(txtEdit(mTxtIdx.idx_终止号码).Text) = "" Then
        ShowMsgbox "注意" & vbCrLf & "    " & strName2 & "范围中的结束" & strName2 & "必须输入,请检查!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_终止号码): Exit Function
    End If
    If Not IsNumeric(txtEdit(mTxtIdx.idx_开始号码).Text) Then
        ShowMsgbox "注意" & vbCrLf & "    " & strName2 & "范围中的开始" & strName2 & "必须输入数字,请检查!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_开始号码): Exit Function
    End If
    If Not IsNumeric(txtEdit(mTxtIdx.idx_终止号码).Text) Then
        ShowMsgbox "注意" & vbCrLf & "    " & strName2 & "范围中的结束" & strName2 & "必须输入数字,请检查!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_终止号码): Exit Function
    End If
    '103428:李南春，2017/2/15，检查卡号长度
    If zlCommFun.ActualLen(txtEdit(mTxtIdx.idx_开始前缀) & txtEdit(mTxtIdx.idx_开始号码).Text) <> mlng长度 Then
        If mint票种 = gBillType.就诊卡 Or mint票种 = gBillType.消费卡 Then
            byt发卡控制 = mcllCardProperty(cbo类别.ListIndex + 1)(3)
            Select Case byt发卡控制
                Case 0
                    ShowMsgbox "注意" & vbCrLf & "    卡号范围中的开始卡号长度不对(应为" & mlng长度 & "位),请检查!"
                Case 2
                    ShowMsgbox "注意" & vbCrLf & "    卡号范围中的开始卡号长度未达到最大位数,是否继续？", True, blnDefult
                    If Not blnDefult Then byt发卡控制 = 0
            End Select
        Else
            ShowMsgbox "注意" & vbCrLf & "    号码范围中的开始号码长度不对(应为" & mlng长度 & "位),请检查!"
            byt发卡控制 = 0
        End If
        If byt发卡控制 = 0 Then
            zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_开始号码): Exit Function
        End If
    End If
    If zlCommFun.ActualLen(txtEdit(mTxtIdx.idx_终止前缀) & txtEdit(mTxtIdx.idx_终止号码).Text) <> zlCommFun.ActualLen(txtEdit(mTxtIdx.idx_开始前缀) & txtEdit(mTxtIdx.idx_开始号码).Text) Then
        ShowMsgbox "注意" & vbCrLf & "    " & strName2 & "范围中的结束" & strName2 & "与开始" & strName2 & "的长度不一致,请检查!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_终止号码): Exit Function
    End If
    If txtEdit(mTxtIdx.idx_终止号码).Text < txtEdit(mTxtIdx.idx_开始号码) Then
        ShowMsgbox "注意" & vbCrLf & "    " & strName2 & "范围中的结束" & strName2 & "小于了开始" & strName2 & ",请检查!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_终止号码): Exit Function
    End If
    
    If zlIsOnlyNum(Trim(txtEdit(mTxtIdx.idx_开始号码))) = False Then
        MsgBox "开始" & strName2 & "中含有非数字字符，字母只能作为前缀。", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_开始号码): Exit Function
    End If
    
    If zlIsOnlyNum(Trim(txtEdit(mTxtIdx.idx_终止号码))) = False Then
        MsgBox "终止" & strName2 & "中含有非数字字符，字母只能作为前缀。", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_终止号码): Exit Function
    End If
    
    '105916:李南春,2017/4/10，检查票据号码
    If txtEdit(mTxtIdx.idx_开始号码).Text = String(mlng长度, "0") And txtEdit(mTxtIdx.idx_终止号码).Text = String(mlng长度, "9") Then
        MsgBox "不能使用" & String(mlng长度, "0") & "-" & String(mlng长度, "9") & "的" & strName2 & "范围。", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_终止号码): Exit Function
    End If
    
    str入库数量 = GetBillNum(Trim(txtEdit(mTxtIdx.idx_开始号码)), Trim(txtEdit(mTxtIdx.idx_终止号码)))
    If InStr(str入库数量, "E") > 0 Or Len(str入库数量) > 10 Then '数量太大已经变成科学计算法
        MsgBox "入库的" & strName1 & "数量不能超过 10000000000 张，请分批入库。", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_开始号码): Exit Function
    End If
    
    '检查是否已经使用过,使用过的票据不能改变其长度
    If mEditType = Ed_修改 And (mstrDrawBill <> "" Or mstrDamnifyBill <> "") Then
        If Len(txtEdit(mTxtIdx.idx_开始号码).Text) <> Len(txtEdit(mTxtIdx.idx_开始号码).Tag) Then
            MsgBox "这张入库的" & strName1 & "已经被使用过," & strName2 & "长度不能改变," & vbCrLf & _
                strName2 & "长度应该是" & Len(txtEdit(mTxtIdx.idx_开始前缀).Text & txtEdit(mTxtIdx.idx_开始号码).Tag) & "位。", vbExclamation, gstrSysName
            zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_终止号码): Exit Function
        End If
        
        varData = Split(mstrDrawBill, ",")
        For i = 0 To UBound(varData)
            If InStr(varData(i), "-") > 0 Then
                varTemp = Split(varData(i), "-")
                If str最小号码 = "" Or str最小号码 > varTemp(0) Then
                    str最小号码 = varTemp(0)
                End If
                If str最大号码 = "" Or str最大号码 < varTemp(1) Then
                    str最大号码 = varTemp(1)
                End If
            Else
                If str最小号码 = "" Or str最小号码 > varData(i) Then
                    str最小号码 = varData(i)
                End If
                If str最大号码 = "" Or str最大号码 < varData(i) Then
                    str最大号码 = varData(i)
                End If
            End If
        Next
        varData = Split(mstrDamnifyBill, ",")
        For i = 0 To UBound(varData)
            If InStr(varData(i), "-") > 0 Then
                varTemp = Split(varData(i), "-")
                If str最小号码 = "" Or str最小号码 > varTemp(0) Then
                    str最小号码 = varTemp(0)
                End If
                If str最大号码 = "" Or str最大号码 < varTemp(1) Then
                    str最大号码 = varTemp(1)
                End If
            Else
                If str最小号码 = "" Or str最小号码 > varData(i) Then
                    str最小号码 = varData(i)
                End If
                If str最大号码 = "" Or str最大号码 < varData(i) Then
                    str最大号码 = varData(i)
                End If
            End If
        Next
        
        If txtEdit(mTxtIdx.idx_开始前缀).Text & txtEdit(mTxtIdx.idx_开始号码).Text > str最小号码 Then
            MsgBox "这张入库的" & strName1 & "已经使用，" & vbCrLf & "开始" & strName2 & "只可以小于" & str最小号码 & "。", vbExclamation, gstrSysName
            zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_开始号码): Exit Function
        End If
        If txtEdit(mTxtIdx.idx_终止前缀).Text & txtEdit(mTxtIdx.idx_终止号码).Text < str最大号码 Then
            MsgBox "这张入库的" & strName1 & "已经使用，" & vbCrLf & strName2 & "已经用到" & str最大号码 & ",终止" & strName2 & "必须大于它。", vbExclamation, gstrSysName
            zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_终止号码): Exit Function
        End If
    End If
    
    '检查是否有使用类别
    If mint票种 = 1 Or mint票种 = 3 Then
        If cbo类别.ListIndex < 0 Then
            MsgBox "注意:" & vbCrLf & "    使用类别没有选择,请选择!", vbInformation + vbOKOnly, gstrSysName
            zlControl.ControlSetFocus cbo类别: Exit Function
            Exit Function
        End If
    End If
    
    str类别 = Get使用类别(mint票种)
    
    '检查批次是否重复
    If mEditType = Ed_增加 Then
        If Trim(txtEdit(idx_批次).Text) <> "" Then
            If batchValied = False Then
                ShowMsgbox "存在与本次入库批次相同的入库记录，不能继续，请检查！"
                Exit Function
            End If
        End If
    ElseIf mEditType = Ed_修改 Then
        If Trim(txtEdit(idx_批次).Text) = "" Then
            ShowMsgbox "未输入入库批次，不能继续，请检查！"
            Exit Function
        ElseIf Trim(txtEdit(mTxtIdx.idx_批次).Tag) <> Trim(txtEdit(idx_批次).Text) Then
            If batchValied = False Then
                ShowMsgbox "存在与本次入库批次相同的入库记录，不能继续，请检查！"
                Exit Function
            End If
        End If
    End If
    
    '检查是否已经领用并且使用类别与当前修改的不一致时
    If mEditType = Ed_修改 And str类别 <> Trim(cbo类别.Tag) Then
        If mint票种 = gBillType.消费卡 Then
            gstrSQL = _
                "Select b.名称 As 使用类别 " & _
                "From 消费卡领用记录 A,消费卡类别目录 B " & _
                "Where Nvl(a.接口编号,0)=b.编号(+) And a.批次=[1] And Nvl(a.接口编号,0)<>Nvl([3],0) And Nvl(a.剩余数量,0) >0 And Rownum < 2 "
        ElseIf mint票种 = gBillType.就诊卡 Then
            gstrSQL = _
                "Select b.名称 As 使用类别 " & _
                "From 票据领用记录 A,医疗卡类别 B " & _
                "Where To_Number(Nvl(a.使用类别,0))=b.ID(+) And a.批次=[1] And a.票种=[2] " & _
                "      And Nvl(a.使用类别,'LXH')<>Nvl([3],'LXH') And Nvl(a.剩余数量,0) >0 And Rownum < 2 "
        Else
            gstrSQL = _
                "Select " & IIf(mint票种 = gBillType.预交收据, "Decode(使用类别,'2','住院预交','门诊预交') As 使用类别 ", "使用类别 ") & _
                "From 票据领用记录 " & _
                "Where 批次=[1] And 票种=[2] And Nvl(使用类别,'LXH')<>Nvl([3],'LXH') And Nvl(剩余数量,0) >0 And Rownum < 2 "
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng入库ID, mint票种, str类别)
        If rsTemp.EOF = False Then
            If rsTemp.RecordCount >= 0 Then
                If MsgBox("注意:" & vbCrLf & "     你将原使用类别为『" & IIf(Trim(NVL(rsTemp!使用类别)) = "", "不区分类别", NVL(rsTemp!使用类别)) & "』调整为" & vbCrLf & _
                                  "    『" & IIf(Trim(cbo类别.Text) = "", "不区分类别", cbo类别.Text) & "』的入库记录已经被领用, " & vbCrLf & _
                                  "     是否将领用的" & strName1 & "一起调整? ", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    zlControl.ControlSetFocus cbo类别: Exit Function
                End If
            End If
        End If
    End If
    
    '判断入库是否重复
    str开始号码 = Trim(txtEdit(mTxtIdx.idx_开始前缀).Text) & Trim(txtEdit(mTxtIdx.idx_开始号码).Text)
    str结束号码 = Trim(txtEdit(mTxtIdx.idx_终止前缀).Text) & Trim(txtEdit(mTxtIdx.idx_终止号码).Text)
    '115348:李南春，2017/10/24，医疗卡要分类别检查，不同医疗卡可能卡号会有重复
    If mint票种 = gBillType.消费卡 Then
        gstrSQL = _
            "Select ID,nvl(剩余数量,0) 剩余数量 " & _
            "From 消费卡入库记录 " & _
            "Where ID<>[3] And Nvl(接口编号,0)=Nvl([5],0) " & _
            "      And (([1] Between 开始卡号 And 终止卡号) Or ([2] Between 开始卡号 And 终止卡号))  " & _
            "      And Length(开始卡号)=Length([1]) And 批次=[6]"
    Else
        gstrSQL = _
            "Select ID,nvl(剩余数量,0) 剩余数量 " & _
            "From 票据入库记录 " & _
            "Where ID<>[3] And 票种=[4] And nvl(使用类别,'LXH')=nvl([5],'LXH') " & _
            "      And (([1] between 开始号码 and  终止号码) or  ([2] between 开始号码  and 终止号码)) " & _
            "      And Length(开始号码)=Length([1]) And 批次=[6]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str开始号码, str结束号码, mlng入库ID, mint票种, str类别, Trim(txtEdit(idx_批次).Text))
    If rsTemp.RecordCount > 0 Then
        If mblnIsBIll Then
            If MsgBox("存在与本次入库重叠的入库记录" & IIf(Val(NVL(rsTemp!剩余数量)) > 0, "，并且还有未使用完的票据。", "。") & vbCrLf & _
                "你还需要继续吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        Else
            ShowMsgbox "存在与本次入库卡号重叠的入库记录，不能继续，请检查！"
            Exit Function
        End If
    End If
    
    '102996:李南春,2016/11/23,医疗发票电子化管理
    If (mEditType = Ed_增加 Or mEditType = Ed_修改) And gblnBillPrint Then
        Dim strErrMsg As String, strExpended As String
        On Error Resume Next
        If gobjBillPrint.zlBillInCheckValied(mEditType + 1, mint票种, str类别, mlng入库ID, str开始号码, str结束号码, strExpended) = False Then
            zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_开始号码): Exit Function
        End If
        Err = 0: On Error GoTo errHandle
    End If
    
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function batchValied() As Boolean
    '检查批次是否重复
    Dim str接口编号 As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    If mint票种 = gBillType.消费卡 Then
        str接口编号 = Get使用类别(mint票种)
        gstrSQL = "Select 1 From 消费卡入库记录 Where 接口编号 = [1] And 批次 = [2] And ID <> 批次 And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str接口编号, Trim(txtEdit(idx_批次).Text))
    Else
        gstrSQL = "Select 1 From 票据入库记录 Where 票种 = [1] And 批次 = [2] And ID <> 批次 And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint票种, Trim(txtEdit(idx_批次).Text))
    End If
    If rsTemp.EOF = False Then
        Exit Function
    End If
    batchValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Get使用类别(ByVal int票种 As gBillType) As String
    '获取使用类别
    Dim str类别 As String
    
    On Error GoTo errHandle
    Select Case int票种
    Case gBillType.收费收据, gBillType.结帐收据
        '收费和结帐
        str类别 = Trim(cbo类别.Text)
    Case gBillType.预交收据
        str类别 = cbo类别.ItemData(cbo类别.ListIndex) - 1
        If Val(str类别) = 0 Then str类别 = ""
    Case gBillType.就诊卡, gBillType.消费卡
        str类别 = cbo类别.ItemData(cbo类别.ListIndex)
        If Val(str类别) = 0 Then str类别 = ""
    Case Else
        str类别 = ""
    End Select
    Get使用类别 = str类别
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存数据
    '返回:数据保存成功,返回true,否则返回为False
    '编制:刘兴洪
    '日期:2010-11-16 15:15:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '问题号:54259
    Dim lngID As Long, str入库数量 As String, str剩余数量 As String
    Dim varTemp As Variant, str类别 As String
    Dim blnTrans As Boolean, strErrMsg As String, strExpended As String
    
    On Error GoTo errHandle
    
    str入库数量 = GetBillNum(Trim(txtEdit(mTxtIdx.idx_开始号码)), Trim(txtEdit(mTxtIdx.idx_终止号码)))
    str剩余数量 = "0"
    If Len(str入库数量) <= 0 Then
       If Val(str入库数量) <= 0 Then Exit Function
    End If
    str类别 = Get使用类别(mint票种)
        
    If mEditType = Ed_增加 Then
        If mint票种 = gBillType.消费卡 Then
            lngID = zlDatabase.GetNextId("消费卡入库记录")
        Else
            lngID = zlDatabase.GetNextId("票据入库记录")
        End If
        str剩余数量 = str入库数量
    Else
        lngID = mlng入库ID
        '77390:李南春,2014/9/3 09:33:32,计算票据数量
        varTemp = Split(vsMemo.Tag & "-", "-")
        str剩余数量 = GetBillNum(varTemp(1), varTemp(0))
        If Val(str剩余数量) < 0 Then str剩余数量 = "0"
        
        str剩余数量 = GetBillNum(str剩余数量, str入库数量)
        If Val(str剩余数量) < 0 Then Exit Function
    End If
    
    If mint票种 = gBillType.消费卡 Then
        ' Zl_消费卡入库记录_Insert
        gstrSQL = "Zl_消费卡入库记录_Insert("
        '  Id_In       In 消费卡入库记录.ID%Type,
        gstrSQL = gstrSQL & "" & lngID & ","
        '  接口编号_In In 消费卡入库记录.接口编号%Type,
        gstrSQL = gstrSQL & "" & Val(str类别) & ","
        '  前缀文本_In In 消费卡入库记录.前缀文本%Type,
        gstrSQL = gstrSQL & "'" & Trim(txtEdit(mTxtIdx.idx_开始前缀)) & "',"
        '  开始卡号_In In 消费卡入库记录.开始卡号%Type,
        gstrSQL = gstrSQL & "'" & Trim(txtEdit(mTxtIdx.idx_开始前缀)) & Trim(txtEdit(mTxtIdx.idx_开始号码)) & "',"
        '  终止卡号_In In 消费卡入库记录.终止卡号%Type,
        gstrSQL = gstrSQL & "'" & Trim(txtEdit(mTxtIdx.idx_终止前缀)) & Trim(txtEdit(mTxtIdx.idx_终止号码)) & "',"
        '  入库数量_In In 消费卡入库记录.入库数量%Type,
        gstrSQL = gstrSQL & "'" & str入库数量 & "',"
        '  剩余数量_In In 消费卡入库记录.剩余数量%Type,
        gstrSQL = gstrSQL & "'" & str剩余数量 & "',"
        '  备注_In     In 消费卡入库记录.备注%Type,
        gstrSQL = gstrSQL & "'" & Trim(txtEdit(mTxtIdx.idx_备注)) & "',"
        '  登记人_In   In 消费卡入库记录.登记人%Type,
        gstrSQL = gstrSQL & "'" & UserInfo.姓名 & "',"
        '  修改标志_In Integer:=0
        gstrSQL = gstrSQL & "" & IIf(mEditType = Ed_增加, 0, 1) & ","
        '  批次_In     In 票据入库记录.批次%Type := Null
        gstrSQL = gstrSQL & "" & IIf(Trim(txtEdit(idx_批次).Text) = "", lngID, Trim(txtEdit(idx_批次).Text)) & ")"
    Else
        ' Zl_票据入库记录_Insert
        gstrSQL = "Zl_票据入库记录_Insert("
        '  Id_In       In 票据入库记录.ID%Type,
        gstrSQL = gstrSQL & "" & lngID & ","
        '  票种_In     In 票据入库记录.票种%Type,
        gstrSQL = gstrSQL & "" & mint票种 & ","
        '  使用类别_In In 票据入库记录.使用类别%Type,
        gstrSQL = gstrSQL & "" & IIf(str类别 = "", "NULL", "'" & str类别 & "'") & ","
        '  前缀文本_In In 票据入库记录.前缀文本%Type,
        gstrSQL = gstrSQL & "'" & Trim(txtEdit(mTxtIdx.idx_开始前缀)) & "',"
        '  开始号码_In In 票据入库记录.开始号码%Type,
        gstrSQL = gstrSQL & "'" & Trim(txtEdit(mTxtIdx.idx_开始前缀)) & Trim(txtEdit(mTxtIdx.idx_开始号码)) & "',"
        '  终止号码_In In 票据入库记录.终止号码%Type,
        gstrSQL = gstrSQL & "'" & Trim(txtEdit(mTxtIdx.idx_终止前缀)) & Trim(txtEdit(mTxtIdx.idx_终止号码)) & "',"
        '  入库数量_In In 票据入库记录.入库数量%Type,
        gstrSQL = gstrSQL & "'" & str入库数量 & "',"
        '  剩余数量_In In 票据入库记录.剩余数量%Type,
        gstrSQL = gstrSQL & "'" & str剩余数量 & "',"
        '  备注_In     In 票据入库记录.备注%Type,
        gstrSQL = gstrSQL & "'" & Trim(txtEdit(mTxtIdx.idx_备注)) & "',"
        '  登记人_In   In 票据入库记录.登记人%Type,
        gstrSQL = gstrSQL & "'" & UserInfo.姓名 & "',"
        '  修改标志_In Integer:=0
        gstrSQL = gstrSQL & "" & IIf(mEditType = Ed_增加, 0, 1) & ","
         '  批次_In     In 票据入库记录.批次%Type := Null
        gstrSQL = gstrSQL & "" & IIf(Trim(txtEdit(idx_批次).Text) = "", lngID, Trim(txtEdit(idx_批次).Text)) & ")"
    End If
    
    '102996:李南春,2016/11/23,医疗发票电子化管理
    gcnOracle.BeginTrans: blnTrans = True
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    If gblnBillPrint Then
        On Error Resume Next
        If gobjBillPrint.zlBillIn(mEditType + 1, mint票种, str类别, lngID, strExpended) = False Then
            gcnOracle.RollbackTrans: blnTrans = False
            zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_开始号码): Exit Function
        End If
        Err = 0: On Error GoTo errHandle
    End If
    gcnOracle.CommitTrans: blnTrans = False
    SaveData = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cmdOK_Click()
    If mEditType = Ed_查看 Then
        mblnChange = False
        Unload Me: Exit Sub
    End If
    If isValied() = False Then Exit Sub
    If SaveData() = False Then Exit Sub
    mintSucceed = mintSucceed + 1
    If mEditType = Ed_修改 Then
        mblnChange = False
        Unload Me
        Exit Sub
    End If
    Call ClearData: mblnChange = False
    zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_开始前缀)
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If mEditType = Ed_查看 Then Exit Sub
    
    mblnChange = True
    If Index = mTxtIdx.idx_开始前缀 And txtEdit(mTxtIdx.idx_开始前缀).Text <> txtEdit(mTxtIdx.idx_终止前缀).Text Then
        txtEdit(mTxtIdx.idx_终止前缀).Text = txtEdit(mTxtIdx.idx_开始前缀).Text
    End If
    If Index = mTxtIdx.idx_终止前缀 And txtEdit(mTxtIdx.idx_开始前缀).Text <> txtEdit(mTxtIdx.idx_终止前缀).Text Then
        txtEdit(mTxtIdx.idx_开始前缀).Text = txtEdit(mTxtIdx.idx_终止前缀).Text
    End If
    If Index = mTxtIdx.idx_开始前缀 Or Index = mTxtIdx.idx_终止前缀 Then
        txtEdit(mTxtIdx.idx_开始号码).MaxLength = mlng长度 - zlCommFun.ActualLen(txtEdit(mTxtIdx.idx_开始前缀).Text)
        txtEdit(mTxtIdx.idx_终止号码).MaxLength = txtEdit(mTxtIdx.idx_开始号码).MaxLength
    End If
    If Index = mTxtIdx.idx_开始号码 Or Index = mTxtIdx.idx_终止号码 Then
        Call SetMemo
    End If
End Sub
Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    If idx_备注 = Index Then
        zlCommFun.OpenIme True
    Else
        zlCommFun.OpenIme False
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    If Index = mTxtIdx.idx_开始前缀 Or Index = mTxtIdx.idx_终止前缀 Then
        txtEdit(Index).Text = UCase(txtEdit(Index).Text)
    End If
    txtEdit(Index).Text = Trim(txtEdit(Index).Text)
    If idx_备注 = Index Then zlCommFun.OpenIme False
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Index = mTxtIdx.idx_开始前缀 Or Index = mTxtIdx.idx_终止前缀 Then
        If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
            KeyAscii = KeyAscii - Asc("a") + Asc("A")
        End If
    ElseIf Index = mTxtIdx.idx_批次 Then
        If InStr("'[]，。‘：；,.'［］", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    Else
        If Index <> mTxtIdx.idx_备注 Then
            If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
                KeyAscii = 0
            End If
        Else
            zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m文本式
        End If
    End If
End Sub

Private Function LoadCombox() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载Combox数据
    '返回:加载成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-04-27 10:22:29
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim str类别 As String
    
    On Error GoTo errHandle
    str类别 = mstrPreType(mint票种)
    lblUserType.Caption = "使用类别(&L)"
    lbl(6).Caption = "号码范围(&B)"
    Select Case mint票种
    Case gBillType.收费收据, gBillType.结帐收据
        strSQL = "Select 编码,名称,简码,缺省标志 From 票据使用类别 "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        With cbo类别
            .Clear
            Do While Not rsTemp.EOF
                .AddItem NVL(rsTemp!名称)
                .ItemData(.NewIndex) = 1
                If Val(NVL(rsTemp!缺省标志)) = 1 And .ListIndex < 0 Then .ListIndex = .NewIndex
                If mstr类别 = NVL(rsTemp!名称) Then .ListIndex = .NewIndex
                rsTemp.MoveNext
            Loop
            .AddItem " "    '允许设置为空
            If mstr类别 = "" Then .ListIndex = .NewIndex
            If .ListIndex < 0 Then .ListIndex = 0
            .Visible = True: lblUserType.Visible = True
        End With
  Case gBillType.预交收据
        mblnNotClick = True
        With cbo类别
            .Clear
            If InStr(1, mstrPrivs, ";预交门诊票据;") > 0 Then
                .AddItem "门诊预交": .ItemData(.NewIndex) = 2
                If Val(str类别) = 2 Then .ListIndex = .NewIndex
            End If
            If InStr(1, mstrPrivs, ";预交住院票据;") > 0 Then
                .AddItem "住院预交": .ItemData(.NewIndex) = 3
                If Val(str类别) = 3 Then .ListIndex = .NewIndex
            End If
            '58071
            If InStr(1, mstrPrivs, ";预交住院票据;") > 0 And InStr(1, mstrPrivs, ";预交门诊票据;") > 0 Then
                .AddItem " "
                .ItemData(.NewIndex) = 1
            End If
            If .ListIndex < 0 And .ListCount > 0 Then .ListIndex = 0
        End With
        mblnNotClick = False
    Case gBillType.就诊卡
        '103428:李南春，2017/2/15，检查卡号长度
        strSQL = "Select ID,编码,名称,缺省标志,卡号长度,卡号密文,前缀文本,发卡控制 From 医疗卡类别 where nvl(是否启用,0) >=1 Order by 编码 "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        mblnNotClick = True
        Set mcllCardProperty = New Collection
        With cbo类别
            .Clear
            Do While Not rsTemp.EOF
                .AddItem NVL(rsTemp!编码) & "-" & NVL(rsTemp!名称)
                .ItemData(.NewIndex) = Val(NVL(rsTemp!ID))
                mcllCardProperty.Add Array(Val(NVL(rsTemp!卡号长度)), CStr(NVL(rsTemp!前缀文本)), CStr(NVL(rsTemp!卡号密文)), Val(NVL(rsTemp!发卡控制))), "K" & Val(NVL(rsTemp!ID))
                If Val(NVL(rsTemp!缺省标志)) = 1 And .ListIndex < 0 Then .ListIndex = .NewIndex
                If Val(str类别) = Val(NVL(rsTemp!ID)) Then .ListIndex = .NewIndex
                rsTemp.MoveNext
            Loop
            If .ListIndex < 0 And .ListCount > 0 Then .ListIndex = 0
            lblUserType.Caption = "卡类别(&L)"
            lbl(6).Caption = "卡号范围(&B)"
            mblnNotClick = False
        End With
    Case gBillType.消费卡
        strSQL = "Select 编号,名称,卡号长度,是否密文 As 卡号密文,前缀文本 From 消费卡类别目录 Where Nvl(启用,0) >=1 Order By 编号 "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
        mblnNotClick = True
        Set mcllCardProperty = New Collection
        With cbo类别
            .Clear
            Do While Not rsTemp.EOF
                .AddItem NVL(rsTemp!编号) & "-" & NVL(rsTemp!名称)
                .ItemData(.NewIndex) = Val(NVL(rsTemp!编号))
                mcllCardProperty.Add Array(Val(NVL(rsTemp!卡号长度)), CStr(NVL(rsTemp!前缀文本)), CStr(NVL(rsTemp!卡号密文)), 0), "K" & Val(NVL(rsTemp!编号))
                If Val(str类别) = Val(NVL(rsTemp!编号)) Then .ListIndex = .NewIndex
                rsTemp.MoveNext
            Loop
            If .ListIndex < 0 And .ListCount > 0 Then .ListIndex = 0
            lblUserType.Caption = "卡类别(&L)"
            lbl(6).Caption = "卡号范围(&B)"
            mblnNotClick = False
        End With
    Case Else
            cbo类别.Visible = False: lblUserType.Visible = False
    End Select
    LoadCombox = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

