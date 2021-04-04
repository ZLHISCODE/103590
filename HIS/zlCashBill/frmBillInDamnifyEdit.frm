VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmBillInDamnifyEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "票据报损"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9600
   Icon            =   "frmBillInDamnifyEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Height          =   7350
      Left            =   7980
      TabIndex        =   28
      Top             =   -60
      Width           =   30
   End
   Begin VB.Frame fra 
      Caption         =   "本次报损情况"
      Height          =   5145
      Left            =   165
      TabIndex        =   11
      Top             =   1695
      Width           =   7680
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   360
         Index           =   10
         Left            =   1035
         MaxLength       =   20
         TabIndex        =   22
         Top             =   4500
         Width           =   2145
      End
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   360
         Index           =   9
         Left            =   5415
         MaxLength       =   20
         TabIndex        =   24
         Top             =   4500
         Width           =   2145
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         Index           =   8
         Left            =   1050
         MaxLength       =   200
         TabIndex        =   20
         Top             =   4065
         Width           =   6525
      End
      Begin VB.CommandButton cmdRemove 
         Cancel          =   -1  'True
         Caption         =   "删除(&R)"
         Height          =   375
         Left            =   6720
         TabIndex        =   17
         Top             =   285
         Width           =   825
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "增加(&N)"
         Height          =   375
         Left            =   5835
         TabIndex        =   16
         Top             =   285
         Width           =   840
      End
      Begin VB.TextBox txtEdit 
         Height          =   360
         Index           =   7
         Left            =   3660
         MaxLength       =   20
         TabIndex        =   15
         Top             =   300
         Width           =   2145
      End
      Begin VSFlex8Ctl.VSFlexGrid vsMemo 
         Height          =   3210
         Left            =   105
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   765
         Width           =   7470
         _cx             =   13176
         _cy             =   5662
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
         GridColor       =   -2147483648
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmBillInDamnifyEdit.frx":058A
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
      Begin VB.TextBox txtEdit 
         Height          =   360
         Index           =   6
         Left            =   1185
         MaxLength       =   20
         TabIndex        =   13
         Top             =   285
         Width           =   2145
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "报损人"
         Height          =   180
         Left            =   465
         TabIndex        =   21
         Top             =   4590
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "报损时间"
         Height          =   180
         Index           =   2
         Left            =   4590
         TabIndex        =   23
         Top             =   4590
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "报损原因(&M)"
         Height          =   180
         Left            =   60
         TabIndex        =   19
         Top             =   4155
         Width           =   990
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "～"
         Height          =   240
         Index           =   1
         Left            =   3405
         TabIndex        =   14
         Top             =   390
         Width           =   240
      End
      Begin VB.Label lblE 
         AutoSize        =   -1  'True
         Caption         =   "报损票号(&B)"
         Height          =   180
         Left            =   195
         TabIndex        =   12
         Top             =   390
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   420
      Left            =   8145
      TabIndex        =   27
      Top             =   6450
      Width           =   1200
   End
   Begin VB.Frame fraUse 
      Caption         =   "入库基本信息"
      Height          =   1380
      Left            =   150
      TabIndex        =   0
      Top             =   210
      Width           =   7710
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   360
         Index           =   1
         Left            =   3180
         MaxLength       =   2
         TabIndex        =   4
         Top             =   375
         Width           =   375
      End
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   360
         Index           =   5
         Left            =   1125
         MaxLength       =   20
         TabIndex        =   10
         Top             =   870
         Width           =   6300
      End
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   360
         Index           =   2
         Left            =   3570
         MaxLength       =   20
         TabIndex        =   5
         Top             =   375
         Width           =   1530
      End
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   360
         Index           =   3
         Left            =   5520
         MaxLength       =   2
         TabIndex        =   7
         Top             =   375
         Width           =   375
      End
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   360
         Index           =   4
         Left            =   5910
         MaxLength       =   20
         TabIndex        =   8
         Top             =   375
         Width           =   1530
      End
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   360
         Index           =   0
         Left            =   1125
         MaxLength       =   20
         TabIndex        =   2
         Top             =   390
         Width           =   915
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "备注"
         Height          =   180
         Index           =   0
         Left            =   660
         TabIndex        =   9
         Top             =   960
         Width           =   360
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "号码范围"
         Height          =   180
         Index           =   6
         Left            =   2355
         TabIndex        =   3
         Top             =   465
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "～"
         Height          =   240
         Index           =   5
         Left            =   5250
         TabIndex        =   6
         Top             =   435
         Width           =   240
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "入库批次"
         Height          =   180
         Index           =   7
         Left            =   285
         TabIndex        =   1
         Top             =   480
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   420
      Left            =   8145
      TabIndex        =   26
      Top             =   780
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   420
      Left            =   8145
      TabIndex        =   25
      Top             =   285
      Width           =   1200
   End
End
Attribute VB_Name = "frmBillInDamnifyEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum EditDamnifyType
    EdS_报损 = 0
    EdS_查看 = 2
End Enum
Private mstrPrivs As String, mlngModule As Long
Private mEditType As EditDamnifyType '编辑类型
Private mblnChange As Boolean     '为真时表示已改变了
Private mintSucceed As Integer
Private mlng长度 As Long
Private mlng入库ID  As Long, mint票种 As Integer '票种
Private mlng报损ID As Long
Private mblnFirst As Boolean
Private Enum mTxtIdx
    idx_批次 = 0
    idx_开始前缀 = 1
    idx_开始号码 = 2
    idx_终止前缀 = 3
    idx_终止号码 = 4
    idx_备注 = 5
    idx_报损开始 = 6
    idx_报损结束 = 7
    idx_报损原因 = 8
    idx_报损时间 = 9
    idx_报损人 = 10
End Enum

Public Function zlBillEdit(ByVal frmMain As Form, ByVal EditType As EditDamnifyType, ByVal strPrivs As String, _
    ByVal lngModule As Long, ByVal int票种 As gBillType, ByVal lng入库ID As Long, Optional lng报损ID As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口,票据入库报损功能(包含增加和查看)
    '入参:frmMain-调用主窗体
    '       BillEditType-单据操作类型
    '       strPrivs-权限串
    '       lngModule-模块号
    '       lng入库ID-报损指定批次的入库
    '       lng报损ID-修改或查看时的报损ID值.
    '出参:
    '返回:操作一张以上成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-11-16 10:29:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mEditType = EditType: mstrPrivs = strPrivs: mlngModule = lngModule
    mint票种 = int票种: mlng入库ID = lng入库ID: mlng报损ID = lng报损ID
    mintSucceed = False
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
    
    If UserInfo.姓名 = "" Then
        MsgBox "你还未设置人员的对照关系，请与系统管理员联系，设置后才能使用本功能。", vbExclamation, gstrSysName
        Exit Function
    End If
    Call ClearData  '清除控件数据
    Err = 0: On Error GoTo errHandle
    If mEditType <> EdS_报损 Then
        If mint票种 = gBillType.消费卡 Then
            gstrSQL = _
                "Select Id, 入库id, 开始卡号 As 开始号码, 终止卡号 As 终止号码, 数量, 报损原因, 报损人, 报损时间 " & _
                "From 消费卡报损记录 where id=[1]"
        Else
            gstrSQL = _
                "Select Id, 入库id, 开始号码, 终止号码, 数量, 报损原因, 报损人, 报损时间 " & _
                "From 票据报损记录 where id=[1]"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng报损ID)
        If rsTemp.RecordCount = 0 Then
            MsgBox "注意:" & vbCrLf & "    该批次的报损单据可能已经被他人删除，请检查！", vbOKOnly + vbInformation + vbDefaultButton1, gstrSysName
            Exit Function
        End If
        txtEdit(mTxtIdx.idx_报损人) = Nvl(rsTemp!报损人)
        txtEdit(mTxtIdx.idx_报损原因) = Nvl(rsTemp!报损原因)
        txtEdit(mTxtIdx.idx_报损时间) = Format(rsTemp!报损时间, "yyyy-mm-dd HH:MM:SS")
        With vsMemo
            .Clear 1
            .Rows = 2
            .TextMatrix(1, .ColIndex("序号")) = 1
            If Nvl(rsTemp!开始号码) <> Nvl(rsTemp!终止号码) And Nvl(rsTemp!终止号码) <> "" Then
                .TextMatrix(1, .ColIndex("报损票号")) = Nvl(rsTemp!开始号码) & "-" & Nvl(rsTemp!终止号码)
            Else
                .TextMatrix(1, .ColIndex("报损票号")) = Nvl(rsTemp!开始号码)
            End If
            .TextMatrix(1, .ColIndex("报损数量")) = Nvl(rsTemp!数量)
        End With
        mlng入库ID = Val(Nvl(rsTemp!入库ID))
    End If
    
    If mint票种 = gBillType.消费卡 Then
        gstrSQL = _
            "Select Id, 前缀文本, 开始卡号 As 开始号码, 终止卡号 As 终止号码, 入库数量, 剩余数量, 备注, 登记人, 登记时间  " & _
            "From 消费卡入库记录 " & _
            "Where Id=[1]"
    Else
        gstrSQL = _
            "Select Id, 前缀文本, 开始号码, 终止号码, 入库数量, 剩余数量, 备注, 登记人, 登记时间  " & _
            "From 票据入库记录 " & _
            "Where Id=[1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng入库ID)
    If rsTemp.RecordCount = 0 Then
        MsgBox "注意:" & vbCrLf & "    该批次的入库记录已经被他人删除，请检查！", vbOKOnly + vbInformation + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    txtEdit(mTxtIdx.idx_批次).Text = Nvl(rsTemp!ID)
    txtEdit(mTxtIdx.idx_开始前缀).Text = Nvl(rsTemp!前缀文本)
    lngLen = Len(Trim(txtEdit(mTxtIdx.idx_开始前缀).Text))
    txtEdit(mTxtIdx.idx_开始号码).Text = Mid(Nvl(rsTemp!开始号码), lngLen + 1)
    txtEdit(mTxtIdx.idx_开始号码).Tag = txtEdit(mTxtIdx.idx_开始号码).Text
    txtEdit(mTxtIdx.idx_终止前缀).Text = Nvl(rsTemp!前缀文本)
    txtEdit(mTxtIdx.idx_终止号码).Text = Mid(Nvl(rsTemp!终止号码), lngLen + 1)
    txtEdit(mTxtIdx.idx_终止号码).Tag = txtEdit(mTxtIdx.idx_终止号码).Text
    txtEdit(mTxtIdx.idx_报损开始).MaxLength = Len(Nvl(rsTemp!开始号码))
    txtEdit(mTxtIdx.idx_报损结束).MaxLength = txtEdit(mTxtIdx.idx_报损开始).MaxLength
    If mEditType = Ed_增加 Then
        txtEdit(mTxtIdx.idx_报损人) = UserInfo.姓名
        txtEdit(mTxtIdx.idx_报损时间) = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
        LoadCardData = True
        Exit Function
    End If
    Call RefreshNo
    LoadCardData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetBillNum(ByVal str开始号码 As String, ByVal str终卡号码 As String, Optional ByRef strErrMsg As String = "") As Long
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
    If (str开始号码 = "" And str终卡号码 <> "") Or (str终卡号码 = "" And str开始号码 <> "") Then
        GetBillNum = 1: Exit Function
    End If
    GetBillNum = CDec(str终卡号码) - CDec(str开始号码) + 1
    Exit Function
errHandle:
    strErrMsg = "计算错误或超出了计算范围"
    GetBillNum = 0
End Function

Private Sub ClearData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除控件数据
    '编制:刘兴洪
    '日期:2010-11-16 10:35:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer
    For i = 0 To txtEdit.UBound
        txtEdit(i).Text = ""
        If txtEdit(i).Enabled = False Then
            txtEdit(i).BackColor = Me.BackColor
        Else
            txtEdit(i).BackColor = &H80000005
        End If
    Next
    vsMemo.Clear 1
    vsMemo.Rows = 2
End Sub

Private Sub cmdAdd_Click()
    '增加数据
    Dim i As Long, lngRow As Long, str开始票号 As String, str结束票号 As String
    Dim lng前缀 As Long, lng数量 As Long
    
    On Error GoTo errHandle
    If CheckInputValied = False Then Exit Sub
    With vsMemo
        .Redraw = flexRDNone
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("序号")) = i
            If Trim(.TextMatrix(i, .ColIndex("报损票号"))) = "" Then
                lngRow = i: Exit For
            End If
        Next
        If lngRow = 0 Then
            .Rows = .Rows + 1
            lngRow = .Rows - 1
            .TextMatrix(lngRow, .ColIndex("序号")) = lngRow
        End If
        str开始票号 = Trim(txtEdit(mTxtIdx.idx_报损开始))
        str结束票号 = Trim(txtEdit(mTxtIdx.idx_报损结束))
        lng前缀 = Len(Trim(txtEdit(mTxtIdx.idx_开始前缀)))
        If str开始票号 = str结束票号 Then
            .TextMatrix(lngRow, .ColIndex("报损票号")) = str开始票号
            .Cell(flexcpData, lngRow, .ColIndex("报损票号")) = Mid(str开始票号, lng前缀 + 1)
            .TextMatrix(lngRow, .ColIndex("报损数量")) = 1
        Else
            lng数量 = GetBillNum(Mid(str开始票号, lng前缀 + 1), Mid(str结束票号, lng前缀 + 1))
            
            .TextMatrix(lngRow, .ColIndex("报损票号")) = str开始票号 & IIf(str开始票号 = "" Or str结束票号 = "", "", "-") & str结束票号
            .Cell(flexcpData, lngRow, .ColIndex("报损票号")) = Mid(str开始票号, lng前缀 + 1) & "-" & Mid(str结束票号, lng前缀 + 1)
            .TextMatrix(lngRow, .ColIndex("报损数量")) = lng数量
        End If
        .Row = lngRow
        .Redraw = flexRDBuffered
    End With
    txtEdit(mTxtIdx.idx_报损结束).Text = "": txtEdit(mTxtIdx.idx_报损开始).Text = ""
    zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_报损开始)
    Exit Sub
errHandle:
    vsMemo.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdRemove_Click()
    Dim lngRow As Long
    On Error GoTo errHandle
    '删除添加的报损单
    With vsMemo
        .Redraw = flexRDNone
        lngRow = .Row
        If lngRow < .Rows - 1 Then
            .Row = lngRow + 1
            .RemoveItem lngRow
        ElseIf lngRow = .Rows - 1 And lngRow = 1 Then
            .Cell(flexcpText, lngRow, 0, lngRow, .Cols - 1) = ""
            .Cell(flexcpData, lngRow, 0, lngRow, .Cols - 1) = ""
        Else
            .Row = lngRow - 1
            .RemoveItem lngRow
        End If
        .Redraw = flexRDBuffered
        Call RefreshNo
    End With
    Exit Sub
errHandle:
    vsMemo.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub RefreshNo()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:刷新序号
    '编制:刘兴洪
    '日期:2010-11-17 11:58:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    With vsMemo
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("序号")) = i
        Next
    End With
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If LoadCardData = False Then Unload Me: Exit Sub
    Call SetCtrlEnable
    zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_报损开始)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If InStr("'[]，。‘：；,.'［］", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub SetCtrlEnable()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的Enable属性
    '编制:刘兴洪
    '日期:2010-11-17 17:18:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    For i = 0 To txtEdit.UBound
        If mEditType = EdS_查看 Then
            txtEdit(i).Enabled = False
        End If
        If txtEdit(i).Enabled = False Then
            txtEdit(i).BackColor = Me.BackColor
        Else
            txtEdit(i).BackColor = &H80000005
        End If
    Next
End Sub

Private Sub Form_Load()
    Dim blnBill As Boolean
    
    mblnFirst = True
    blnBill = CurrentIsBill(mint票种)
    lbl(6).Caption = IIf(blnBill, "号码范围", "卡号范围")
    lblE.Caption = IIf(blnBill, "报损票号(&B)", "报损卡号(&B)")
    vsMemo.TextMatrix(0, vsMemo.ColIndex("报损票号")) = IIf(blnBill, "报损票号", "报损卡号")
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

Private Function CheckInputValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查输入的报损开始号或结束号是否合法
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-11-16 17:48:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strStartNo As String, strEndNo As String, strTemp As String
    Dim lngLen As Integer, str报损 As String, str领用 As String, rsTemp As ADODB.Recordset
    Dim strName As String
    
    On Error GoTo errHandle
    strName = IIf(CurrentIsBill(mint票种), "号码", "卡号")
    If Trim(txtEdit(mTxtIdx.idx_报损开始).Text) = "" And Trim(txtEdit(mTxtIdx.idx_报损结束).Text) = "" Then
        ShowMsgbox "注意" & vbCrLf & "    报损范围中的开始" & strName & "或结束" & strName & "必须输入,请检查!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_报损开始): Exit Function
    End If
    lngLen = Len(txtEdit(mTxtIdx.idx_开始前缀))
    strTemp = Mid(txtEdit(mTxtIdx.idx_报损开始).Text, lngLen + 1)
    If strTemp <> "" Then
        If zlIsOnlyNum(strTemp) = False Then
            MsgBox "报损范围中的开始" & strName & "中含有非数字字符，字母只能作为前缀。", vbExclamation, gstrSysName
            zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_报损开始): Exit Function
        End If
    End If
    strTemp = Mid(txtEdit(mTxtIdx.idx_报损结束).Text, lngLen + 1)
    If strTemp <> "" Then
        If zlIsOnlyNum(strTemp) = False Then
                MsgBox "报损范围中的终止" & strName & "中含有非数字字符，字母只能作为前缀。", vbExclamation, gstrSysName
                zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_报损结束): Exit Function
        End If
    End If
    mlng长度 = zlCommFun.ActualLen(txtEdit(mTxtIdx.idx_开始前缀) & txtEdit(mTxtIdx.idx_开始号码))
    If txtEdit(mTxtIdx.idx_报损开始).Text <> "" Then
        If zlCommFun.ActualLen(txtEdit(mTxtIdx.idx_报损开始).Text) <> mlng长度 Then
            ShowMsgbox "注意" & vbCrLf & "    报损范围中的开始" & strName & "长度不对(应为" & mlng长度 & "位()),请检查!"
            zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_报损开始): Exit Function
        End If
    End If
    If txtEdit(mTxtIdx.idx_报损结束).Text <> "" Then
        If zlCommFun.ActualLen(txtEdit(mTxtIdx.idx_报损结束).Text) <> mlng长度 Then
            ShowMsgbox "注意" & vbCrLf & "    报损范围中的结束" & strName & "长度不对(应为" & mlng长度 & "位),请检查!"
            zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_报损结束): Exit Function
        End If
        If txtEdit(mTxtIdx.idx_报损结束).Text < txtEdit(mTxtIdx.idx_报损开始) _
            And Trim(txtEdit(mTxtIdx.idx_报损结束).Text) <> "" And txtEdit(mTxtIdx.idx_报损开始) <> "" Then
            ShowMsgbox "注意" & vbCrLf & "    报损范围中的结束" & strName & "小于了开始" & strName & ",请检查!"
            zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_报损结束): Exit Function
        End If
    End If
    '检查是否包含在网格中了
    Dim varTemp As Variant
    With vsMemo
        strStartNo = Trim(txtEdit(mTxtIdx.idx_报损开始))
        strEndNo = Trim(txtEdit(mTxtIdx.idx_报损结束))
        For i = 1 To .Rows - 1
             If .TextMatrix(i, .ColIndex("报损票号")) <> "" Then
                varTemp = Split(.TextMatrix(i, .ColIndex("报损票号")) & "-", "-")
                If varTemp(1) = "" Then varTemp(1) = varTemp(0)
                If varTemp(0) <> "" And varTemp(1) <> "" Then
                    If strStartNo >= varTemp(0) And strStartNo <= varTemp(1) Then
                        ShowMsgbox "注意" & vbCrLf & "    报损范围中的开始" & strName & "已经包含在了第" & i & "行中了,请检查!"
                        .Row = i
                        If .RowIsVisible(.Row) = False Then Call .ShowCell(.Row, .Col)
                        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_报损开始): Exit Function
                    End If
                    
                    If strEndNo >= varTemp(0) And strEndNo <= varTemp(1) And strEndNo <> "" Then
                        ShowMsgbox "注意" & vbCrLf & "    报损范围中的结束" & strName & "已经包含在了第" & i & "行中了,请检查!"
                        .Row = i
                        If .RowIsVisible(.Row) = False Then Call .ShowCell(.Row, .Col)
                        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_报损结束): Exit Function
                    End If
                ElseIf varTemp(0) <> "" Then
                    If strStartNo = varTemp(0) Then
                        ShowMsgbox "注意" & vbCrLf & "    报损范围中的开始" & strName & "已经包含在了第" & i & "行中了,请检查!"
                        .Row = i
                        If .RowIsVisible(.Row) = False Then Call .ShowCell(.Row, .Col)
                        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_报损开始): Exit Function
                    End If
                    If strEndNo = varTemp(0) Then
                        ShowMsgbox "注意" & vbCrLf & "    报损范围中的结束" & strName & "已经包含在了第" & i & "行中了,请检查!"
                        .Row = i
                        If .RowIsVisible(.Row) = False Then Call .ShowCell(.Row, .Col)
                        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_报损结束): Exit Function
                    End If
                ElseIf varTemp(1) <> "" Then
                    If strStartNo = varTemp(1) Then
                        ShowMsgbox "注意" & vbCrLf & "    报损范围中的开始" & strName & "已经包含在了第" & i & "行中了,请检查!"
                        .Row = i
                        If .RowIsVisible(.Row) = False Then Call .ShowCell(.Row, .Col)
                        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_报损开始): Exit Function
                    End If
                    If strEndNo = varTemp(1) Then
                        ShowMsgbox "注意" & vbCrLf & "    报损范围中的结束" & strName & "已经包含在了第" & i & "行中了,请检查!"
                        .Row = i
                        If .RowIsVisible(.Row) = False Then Call .ShowCell(.Row, .Col)
                        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_报损结束): Exit Function
                    End If
                End If
             End If
        Next
    End With
    
    '检查是否存在使用的情况
    If mint票种 = gBillType.消费卡 Then
        gstrSQL = "" & _
        "   Select 1 as 类别,开始卡号 As 开始号码,终止卡号 As 终止号码 " & _
        "   From 消费卡报损记录 " & _
        "   Where (([1] between 开始卡号 and 终止卡号) or ([2] between 开始卡号 and 终止卡号)) and 入库ID=[3] " & _
        "   Union ALL " & _
        "   Select 2 as 类别,开始卡号 As 开始号码, 终止卡号 As 终止号码　" & _
        "   From 消费卡领用记录 " & _
        "   Where (([1] between 开始卡号 and 终止卡号) or ([2] between 开始卡号 and 终止卡号)) and 批次=[3]  "
    Else
        gstrSQL = "" & _
        "   Select 1 as 类别,开始号码,终止号码 " & _
        "   From 票据报损记录 " & _
        "   Where (([1]  between 开始号码  and 终止号码  ) or ([2] between 开始号码  and 终止号码  )) and 入库ID=[3] " & _
        "   Union ALL " & _
        "   Select 2 as 类别,开始号码, 终止号码　" & _
        "   From 票据领用记录 " & _
        "   Where (([1]  between 开始号码  and 终止号码  ) or ([2] between 开始号码  and 终止号码  )) and 批次=[3] and 票种=[4]  "
    End If
    If strStartNo = "" Then strStartNo = strEndNo
    If strEndNo = "" Then strEndNo = strStartNo
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strStartNo, strEndNo, mlng入库ID, mint票种)
    
    If Not rsTemp.EOF Then
        str报损 = "": str领用 = ""
        Do While Not rsTemp.EOF
            If Nvl(rsTemp!终止号码) = Nvl(rsTemp!开始号码) Then
                strTemp = Nvl(rsTemp!开始号码)
            Else
                strTemp = Nvl(rsTemp!开始号码) & "-" & Nvl(rsTemp!终止号码)
            End If
            If rsTemp!类别 = 1 Then
               If Len(str报损) <= 50 Then
                    str报损 = str报损 & vbCrLf & strTemp
               Else
                  If InStr(1, str报损, "...") = 0 Then str报损 = str报损 & vbCrLf & "..."
               End If
            Else
               If Len(str领用) <= 50 Then
                    str领用 = str领用 & vbCrLf & strTemp
               Else
                  If InStr(1, str领用, "...") = 0 Then str领用 = str领用 & vbCrLf & "..."
               End If
            End If
            rsTemp.MoveNext
        Loop
        If str报损 <> "" Then
            ShowMsgbox "注意:" & vbCrLf & "    当前报损范围中的" & strName & "已经被报损,已报损的" & strName & "如下:" & vbCrLf & str报损
            zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_报损开始): Exit Function
            Exit Function
        End If
        If str领用 <> "" Then
            ShowMsgbox "注意:" & vbCrLf & "    当前报损范围中的" & strName & "已经被领用,已领用的" & strName & "如下:" & vbCrLf & str领用
            zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_报损开始): Exit Function
            Exit Function
        End If
    End If
    
    CheckInputValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查输入的数据是否合法
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-11-16 15:04:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, varData As Variant, i As Long, str报损 As String, str领用 As String, strTemp As String
    Dim rsTemp As ADODB.Recordset, blnHaveData As Boolean '是否存在数据
    Dim blnBill As Boolean
    
    On Error GoTo errHandle
    blnBill = CurrentIsBill(mint票种)
    If zlCommFun.ActualLen(Trim(txtEdit(mTxtIdx.idx_报损原因))) > 200 Then
        ShowMsgbox "注意" & vbCrLf & "    报损原因最多只能输入200个字符或100个汉字,请检查!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_报损原因): Exit Function
    End If
    
    If InStr(1, txtEdit(mTxtIdx.idx_报损原因), "'") > 0 Then
        ShowMsgbox "注意" & vbCrLf & "    报损原因中含有非法字符单引号,请检查!"
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_报损原因): Exit Function
    End If
    
    If txtEdit(mTxtIdx.idx_开始号码).Text = String("0", mlng长度) And txtEdit(mTxtIdx.idx_终止前缀).Text = String("9", mlng长度) Then
        MsgBox "不能使用" & String("0", mlng长度) & "-" & String("9", mlng长度) & "的" & IIf(blnBill, "票号", "卡号") & "范围。", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(mTxtIdx.idx_终止号码): Exit Function
    End If
    
    '检查是否该号码已经使用或报损，已经领用的，不能再报损耗，已经报损了的,也不能报损
    With vsMemo
        blnHaveData = False
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("报损票号")) <> "" Then
                blnHaveData = True
                varTemp = Split(.TextMatrix(i, .ColIndex("报损票号")) & "-", "-")
                If varTemp(1) = "" Then varTemp(1) = varTemp(0)
                '1.检查报损情况
                gstrSQL = "" & _
                "   Select 1 as 类别,开始号码,终止号码 " & _
                "   From 票据报损记录 " & _
                "   Where (开始号码>=[1]  and 终止号码 <=[1]) or (开始号码>=[2] and 终止号码<=[2] ) and 入库ID=[3] " & _
                "   Union ALL " & _
                "   Select 2 as 类别,开始号码, 终止号码　" & _
                "   From 票据领用记录 " & _
                "   Where (开始号码>=[1]  and 终止号码 <=[1]) or (开始号码>=[2] and 终止号码<=[2] ) and 批次=[3] and 票种=[4] " & _
                "   "
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CStr(varTemp(0)), CStr(varTemp(1)), mlng入库ID, mint票种)
                If Not rsTemp.EOF Then
                    str报损 = "": str领用 = ""
                    Do While Not .EOF
                        If Nvl(rsTemp!终止号码) = Nvl(rsTemp!开始号码) Then
                            strTemp = Nvl(rsTemp!开始号码)
                        Else
                            strTemp = Nvl(rsTemp!开始号码) & "-" & Nvl(rsTemp!终止号码)
                        End If
                        If rsTemp!类别 = 1 Then
                           If Len(str报损) <= 50 Then
                                str报损 = str报损 & vbCrLf & strTemp
                           Else
                              If InStr(1, str报损, "...") = 0 Then str报损 = str报损 & vbCrLf & "..."
                           End If
                        Else
                           If Len(str领用) <= 50 Then
                                str领用 = str领用 & vbCrLf & strTemp
                           Else
                              If InStr(1, str领用, "...") = 0 Then str领用 = str领用 & vbCrLf & "..."
                           End If
                        End If
                        rsTemp.MoveNext
                    Loop
                    If str报损 <> "" Then
                        ShowMsgbox "注意:" & vbCrLf & "    在第" & i + 1 & "行记录中包含了已经包含报损的" & IIf(blnBill, "票号", "卡号") & _
                            ",已报损的" & IIf(blnBill, "票号", "卡号") & "如下:" & vbCrLf & str报损
                        Exit Function
                    End If
                    If str领用 <> "" Then
                        ShowMsgbox "注意:" & vbCrLf & "    在第" & i + 1 & "行记录中包含了已经领用的" & IIf(blnBill, "票号", "卡号") & _
                            ",已领用的" & IIf(blnBill, "票号", "卡号") & "如下:" & vbCrLf & str领用
                        Exit Function
                    End If
                End If
            End If
        Next
    End With
    If Not blnHaveData Then
        ShowMsgbox "注意:" & vbCrLf & "    你没有选择要报损的" & IIf(blnBill, "票据", "卡片") & "，不能继续！"
        Exit Function
    End If
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存数据
    '返回:数据保存成功,返回true,否则返回为False
    '编制:刘兴洪
    '日期:2010-11-16 15:15:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long, lng报损数量 As Long, strDate As String
    Dim i As Long, cllPro As Collection, varTemp As Variant, varData As Variant
    
    On Error GoTo errHandle
    Set cllPro = New Collection
    With vsMemo
        strDate = "to_Date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd HH24:mi:ss')"
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("报损票号")) <> "" Then
                varTemp = Split(.TextMatrix(i, .ColIndex("报损票号")) & "-", "-")
                If varTemp(1) = "" Then varTemp(1) = varTemp(0)
                lng报损数量 = Val(.TextMatrix(i, .ColIndex("报损数量")))
                If mint票种 = gBillType.消费卡 Then
                    '    Zl_消费卡报损记录_Insert
                    gstrSQL = "Zl_消费卡报损记录_Insert("
                    '      入库id_In   In 消费卡报损记录.入库id%Type,
                    gstrSQL = gstrSQL & "" & mlng入库ID & ","
                    '      开始卡号_In In 消费卡报损记录.开始卡号%Type,
                    gstrSQL = gstrSQL & "'" & varTemp(0) & "',"
                    '      终止卡号_In In 消费卡报损记录.终止卡号%Type,
                    gstrSQL = gstrSQL & "'" & varTemp(1) & "',"
                    '      数量_In     In 消费卡报损记录.数量%Type,
                    gstrSQL = gstrSQL & "" & lng报损数量 & ","
                    '      报损原因_In In 消费卡报损记录.报损原因%Type,
                    gstrSQL = gstrSQL & "" & _
                        IIf(Trim(txtEdit(mTxtIdx.idx_报损原因).Text) = "", "NULL", _
                        "'" & Trim(txtEdit(mTxtIdx.idx_报损原因).Text) & "'") & ","
                    '      报损人_In   In 消费卡报损记录.报损人%Type,
                    gstrSQL = gstrSQL & "'" & UserInfo.姓名 & "',"
                    '      报损时间_In In 消费卡报损记录.报损时间%Type
                    gstrSQL = gstrSQL & "" & strDate & ")"
                Else
                    '    Zl_票据报损记录_Insert
                    gstrSQL = "Zl_票据报损记录_Insert("
                    '      入库id_In   In 票据报损记录.入库id%Type,
                    gstrSQL = gstrSQL & "" & mlng入库ID & ","
                    '      开始号码_In In 票据报损记录.开始号码%Type,
                    gstrSQL = gstrSQL & "'" & varTemp(0) & "',"
                    '      终止号码_In In 票据报损记录.终止号码%Type,
                    gstrSQL = gstrSQL & "'" & varTemp(1) & "',"
                    '      数量_In     In 票据报损记录.数量%Type,
                    gstrSQL = gstrSQL & "" & lng报损数量 & ","
                    '      报损原因_In In 票据报损记录.报损原因%Type,
                    gstrSQL = gstrSQL & "" & IIf(Trim(txtEdit(mTxtIdx.idx_报损原因)) = "", "NULL", "'" & Trim(txtEdit(mTxtIdx.idx_报损原因)) & "'") & ","
                    '      报损人_In   In 票据报损记录.报损人%Type,
                    gstrSQL = gstrSQL & "'" & UserInfo.姓名 & "',"
                    '      报损时间_In In 票据报损记录.报损时间%Type
                    gstrSQL = gstrSQL & "" & strDate & ")"
                End If
                AddArray cllPro, gstrSQL
            End If
        Next
    End With
    ExecuteProcedureArrAy cllPro, Me.Caption
    SaveData = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
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
    mblnChange = False
    Unload Me
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If mEditType = Ed_查看 Then Exit Sub
    mblnChange = True
End Sub
Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    If idx_报损原因 = Index Then zlCommFun.OpenIme True
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    Dim lngLen As Long, lng前缀Len As Long, strTemp As String, str前缀 As String
    Dim strNum As String
    Dim strChr As String
    Dim i As Long
    
    If Index = mTxtIdx.idx_报损开始 Or Index = mTxtIdx.idx_报损结束 Then
        '长度不对时，需要补位
        strTemp = Trim(txtEdit(Index).Text)
        If strTemp = "" Then Exit Sub
        lngLen = Len(txtEdit(mTxtIdx.idx_开始号码))
        str前缀 = Trim(txtEdit(mTxtIdx.idx_开始前缀))
        lng前缀Len = Len(str前缀)
        If Len(txtEdit(Index)) < lngLen Then
            If zlIsOnlyNum(strTemp) Then
                strTemp = str前缀 & zlStr.Lpad(strTemp, lngLen, "0", True)
            ElseIf UCase(Mid(strTemp, 1, lng前缀Len)) = str前缀 Then
                  strTemp = str前缀 & zlStr.Lpad(Mid(strTemp, lng前缀Len + 1), lngLen, "0", True)
            Else
                strNum = ""
                For i = 1 To Len(strTemp)
                    strChr = Mid(strTemp, i, 1)
                    If InStr(1, "0123456789", strChr) > 0 Then
                        strNum = strNum & strChr
                    End If
                Next
                strTemp = str前缀 & zlStr.Lpad(strNum, lngLen, "0", True)
            End If
        ElseIf UCase(Mid(strTemp, 1, lng前缀Len)) = str前缀 Then
                strTemp = str前缀 & Right(Mid(strTemp, lng前缀Len + 1), lngLen)
        Else
                strTemp = Left(strTemp, lng前缀Len + lngLen)
        End If
        txtEdit(Index).Text = UCase(strTemp)
    End If
    txtEdit(Index).Text = Trim(txtEdit(Index).Text)
    If idx_报损原因 = Index Then zlCommFun.OpenIme False
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m文本式
End Sub
