VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmSquareSendCardTypeEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "消费卡类别编辑"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7800
   Icon            =   "frmSquareSendCardTypeEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   7800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6015
      TabIndex        =   48
      Top             =   6600
      Width           =   1100
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   6375
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin VB.Frame fra 
         Height          =   855
         Index           =   11
         Left            =   75
         TabIndex        =   12
         Top             =   1710
         Width           =   5625
         Begin VB.CheckBox chkEdit 
            Caption         =   "体检(&X)"
            Height          =   180
            Index           =   13
            Left            =   3870
            TabIndex        =   16
            Top             =   360
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "住院(&Z)"
            Height          =   180
            Index           =   12
            Left            =   2460
            TabIndex        =   15
            Top             =   360
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "门诊(&M)"
            Height          =   180
            Index           =   11
            Left            =   1140
            TabIndex        =   14
            Top             =   360
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "使用场合："
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   900
         End
      End
      Begin VB.Frame fra 
         Caption         =   "缺省限制类别"
         Height          =   1200
         Index           =   13
         Left            =   75
         TabIndex        =   26
         Top             =   2670
         Width           =   7665
         Begin VSFlex8Ctl.VSFlexGrid vsf限制类别 
            Height          =   945
            Left            =   30
            TabIndex        =   27
            Top             =   210
            Width           =   7575
            _cx             =   13361
            _cy             =   1667
            Appearance      =   2
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
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483643
            GridColorFixed  =   -2147483643
            TreeColor       =   -2147483643
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   0
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   0
            Cols            =   0
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
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
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   3735
         MaxLength       =   3
         TabIndex        =   8
         Tag             =   "卡号长度"
         Top             =   780
         Width           =   1935
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   3735
         MaxLength       =   100
         TabIndex        =   4
         Tag             =   "名称"
         Top             =   225
         Width           =   1935
      End
      Begin VB.CheckBox chkEdit 
         Caption         =   "启用消费卡(&S)"
         Height          =   180
         Index           =   0
         Left            =   3735
         TabIndex        =   11
         Top             =   1395
         Value           =   1  'Checked
         Width           =   1530
      End
      Begin VB.Frame fra 
         Height          =   1140
         Index           =   14
         Left            =   75
         TabIndex        =   28
         Top             =   3870
         Width           =   7665
         Begin VB.CheckBox chkEdit 
            Caption         =   "刷卡"
            Height          =   180
            Index           =   8
            Left            =   1215
            TabIndex        =   30
            Top             =   270
            Value           =   1  'Checked
            Width           =   885
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "扫描卡"
            Height          =   180
            Index           =   9
            Left            =   2250
            TabIndex        =   31
            Top             =   270
            Width           =   960
         End
         Begin VB.OptionButton opt键盘控制 
            Caption         =   "使用字符软键盘"
            Height          =   180
            Index           =   2
            Left            =   5040
            TabIndex        =   35
            Top             =   780
            Width           =   2055
         End
         Begin VB.OptionButton opt键盘控制 
            Caption         =   "使用数字软键盘"
            Height          =   180
            Index           =   1
            Left            =   3090
            TabIndex        =   34
            Top             =   780
            Width           =   1800
         End
         Begin VB.OptionButton opt键盘控制 
            Caption         =   "禁止使用软键盘"
            Height          =   180
            Index           =   0
            Left            =   1170
            TabIndex        =   33
            Top             =   780
            Value           =   -1  'True
            Width           =   1650
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000A&
            X1              =   0
            X2              =   7635
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "读卡性质："
            Height          =   180
            Index           =   5
            Left            =   90
            TabIndex        =   29
            Top             =   270
            Width           =   900
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "键盘控制："
            Height          =   180
            Index           =   8
            Left            =   90
            TabIndex        =   32
            Top             =   750
            Width           =   900
         End
      End
      Begin VB.Frame fra 
         Caption         =   "消费卡属性"
         Height          =   2445
         Index           =   12
         Left            =   5790
         TabIndex        =   17
         Top             =   120
         Width           =   1950
         Begin VB.CheckBox chkEdit 
            Caption         =   "允许换卡(&6)"
            Height          =   180
            Index           =   7
            Left            =   150
            TabIndex        =   23
            Top             =   1625
            Width           =   1335
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "允许余额退款(&8)"
            Height          =   180
            Index           =   3
            Left            =   150
            TabIndex        =   25
            Top             =   2160
            Width           =   1695
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "特定病人(&5)"
            Height          =   180
            Index           =   2
            Left            =   150
            TabIndex        =   22
            Top             =   1360
            Width           =   1335
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "允许补卡(&7)"
            Enabled         =   0   'False
            Height          =   180
            Index           =   1
            Left            =   150
            TabIndex        =   24
            Top             =   1890
            Width           =   1320
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "部分退款(&2)"
            Height          =   180
            Index           =   5
            Left            =   150
            TabIndex        =   19
            Top             =   565
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "卡号密文(&4)"
            Height          =   180
            Index           =   4
            Left            =   150
            TabIndex        =   20
            Top             =   830
            Width           =   1335
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "允许退现(&1)"
            Height          =   180
            Index           =   6
            Left            =   150
            TabIndex        =   18
            Top             =   300
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.CheckBox chkEdit 
            Caption         =   "严格控制(&3)"
            Height          =   180
            Index           =   10
            Left            =   150
            TabIndex        =   21
            Top             =   1095
            Width           =   1320
         End
      End
      Begin VB.Frame fra 
         Caption         =   "密码输入设置"
         Height          =   1290
         Index           =   15
         Left            =   75
         TabIndex        =   36
         Top             =   5085
         Width           =   7665
         Begin VB.Frame fra 
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   16
            Left            =   1335
            TabIndex        =   44
            Top             =   863
            Width           =   3915
            Begin VB.OptionButton optRule 
               Caption         =   "输入字符不限制"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   45
               Top             =   30
               Value           =   -1  'True
               Width           =   1560
            End
            Begin VB.OptionButton optRule 
               Caption         =   "输入字符只能为数字"
               Height          =   180
               Index           =   1
               Left            =   1620
               TabIndex        =   46
               Top             =   30
               Width           =   2070
            End
         End
         Begin VB.TextBox txtEdit 
            Enabled         =   0   'False
            Height          =   270
            Index           =   4
            Left            =   5565
            MaxLength       =   2
            TabIndex        =   42
            Text            =   "0"
            Top             =   375
            Width           =   300
         End
         Begin VB.OptionButton optPassInput 
            Caption         =   "固定输入10位"
            Height          =   210
            Index           =   1
            Left            =   2955
            TabIndex        =   40
            Top             =   405
            Width           =   1545
         End
         Begin VB.TextBox txtEdit 
            Height          =   270
            Index           =   5
            Left            =   435
            MaxLength       =   2
            TabIndex        =   38
            Text            =   "10"
            Top             =   330
            Width           =   300
         End
         Begin VB.OptionButton optPassInput 
            Caption         =   "输入不固定"
            Height          =   210
            Index           =   0
            Left            =   1335
            TabIndex        =   39
            Top             =   390
            Width           =   1380
         End
         Begin VB.OptionButton optPassInput 
            Caption         =   "必须输入    位密码以上"
            Height          =   210
            Index           =   2
            Left            =   4545
            TabIndex        =   41
            Top             =   405
            Width           =   2295
         End
         Begin VB.Line Line2 
            BorderColor     =   &H8000000A&
            X1              =   0
            X2              =   7635
            Y1              =   750
            Y2              =   750
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "长度    位："
            Height          =   180
            Index           =   4
            Left            =   60
            TabIndex        =   37
            Top             =   375
            Width           =   1080
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "密码规则："
            Height          =   180
            Index           =   9
            Left            =   240
            TabIndex        =   43
            Top             =   900
            Width           =   900
         End
      End
      Begin VB.ComboBox cbo结算方式 
         Height          =   300
         Left            =   1065
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1335
         Width           =   1545
      End
      Begin VB.TextBox txtEdit 
         Height          =   315
         Index           =   1
         Left            =   1065
         MaxLength       =   2
         TabIndex        =   6
         Tag             =   "前缀文本"
         Top             =   765
         Width           =   1545
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1065
         MaxLength       =   10
         TabIndex        =   2
         Tag             =   "编码"
         Top             =   225
         Width           =   1545
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "结算方式(&J)"
         Height          =   180
         Index           =   7
         Left            =   60
         TabIndex        =   9
         Top             =   1395
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "卡号长度(&L)"
         Height          =   180
         Index           =   6
         Left            =   2715
         TabIndex        =   7
         Top             =   840
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "前缀文本(&T)"
         Height          =   180
         Index           =   3
         Left            =   60
         TabIndex        =   5
         Top             =   840
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "名称(&N)"
         Height          =   180
         Index           =   2
         Left            =   3075
         TabIndex        =   3
         Top             =   285
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "编码(&B)"
         Height          =   180
         Index           =   0
         Left            =   420
         TabIndex        =   1
         Top             =   285
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4875
      TabIndex        =   47
      Top             =   6600
      Width           =   1100
   End
End
Attribute VB_Name = "frmSquareSendCardTypeEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------------------------------------------------------------------
'入口参数
Public Enum gSendCardEdit
    Card_增加 = 0
    Card_修改 = 1
    Card_删除 = 2
    Card_停用 = 3
    Card_启用 = 4
    Card_查看 = 5
End Enum
Private mlngModule As Long
Private mstrPrivs As String
Private mEditType As gSendCardEdit
Private mlngCardTypeID As Long
'-----------------------------------------------------------------------------------------
Private mintSucces As Integer
Private mblnFirst As Boolean
Private Enum mtxtIdx
     idx_编号 = 0
     idx_名称 = 2
     idx_前缀文本 = 1
     idx_卡号长度 = 3
     idx_密码长度 = 5
     idx_密码位数 = 4
End Enum

Private Enum mchkIdx
    idx_启用 = 0
    idx_退现 = 6
    idx_全退 = 5
    idx_密文 = 4
    idx_严格控制 = 10
    idx_特定病人 = 2
    idx_换卡 = 7
    idx_补卡 = 1
    idx_余额退款 = 3
    
    idx_刷卡 = 8
    idx_扫描卡 = 9
    
    idx_门诊 = 11
    idx_住院 = 12
    idx_体检 = 13
End Enum

Private Type Ty_CardType
    lng卡号长度 As Long
    bln固定 As Boolean
    bln已发卡 As Boolean '是否已经发过卡
    str前缀文本 As String
End Type
Private mCardType As Ty_CardType
Private mblnNotClick As Boolean
Private mblnChange As Boolean

Public Function zlEditSendCard(ByVal frmMain As Object, ByVal lngModule As Long, ByVal strPrivs As String, _
    ByVal EditType As gSendCardEdit, Optional lngCardTypeID As Long = 0) As Boolean
    '功能:医疗卡类别编辑
    '入参:EditType-编辑类型
    '        lngCardTypeID-增加时为0
    '出参:
    '返回:只要成功一次,返回true,否则返回Flase
    mEditType = EditType: mlngModule = lngModule: mstrPrivs = strPrivs
    mlngCardTypeID = lngCardTypeID
    
    On Error Resume Next
    mintSucces = 0
    Me.Show 1, frmMain
    zlEditSendCard = mintSucces > 0
End Function

Private Sub Form_Load()
    Dim ty_Temp As Ty_CardType
    
    mblnFirst = True
    mCardType = ty_Temp '自定义Type初始化
    
    If InitData() = False Then Unload Me: Exit Sub
    If LoadCardData() = False Then Unload Me: Exit Sub
    Call SetCtrlEnable
    
    If mEditType = dt_查看 Then
        cmdOK.Visible = False
    End If
    mblnChange = False
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    If mEditType = Card_增加 Then
        zlControl.ControlSetFocus txtEdit(mtxtIdx.idx_名称)
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("|,'～~;", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnFirst Or mblnChange = False Then Exit Sub
    If mEditType = Card_增加 Or mEditType = gEd_修改 Then
        If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", _
            vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Cancel = 1
    End If
End Sub
 
 Private Function InitData() As Boolean
    '功能:初始化数据
    '返回:初始化成功，返回true,否则返回False
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    If Not (mEditType = Card_增加 Or mEditType = Card_修改) Then InitData = True: Exit Function
    
    If mEditType = Card_增加 Then
        txtEdit(mtxtIdx.idx_编号).Text = zlDatabase.GetMax("消费卡类别目录", "编号", txtEdit(mtxtIdx.idx_编号).MaxLength)
    End If
    
    strSQL = "Select 名称 From 结算方式 Where 性质 = 8 And Nvl(应付款, 0) = 0"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With cbo结算方式
        .Clear
        Do While Not rsTemp.EOF
            If NVL(rsTemp!名称) <> "" Then .AddItem NVL(rsTemp!名称)
            rsTemp.MoveNext
        Loop
    End With
    
    Set rsTemp = zlGet收费类别()
    With vsf限制类别
        .Clear
        Do While Not rsTemp.EOF
           ZL_vsGrid_AddCell vsf限制类别, NVL(rsTemp!编码) & "-" & NVL(rsTemp!名称), NVL(rsTemp!名称), True
           rsTemp.MoveNext
        Loop
        ZL_vsGrid_AutoSetGridRowAndCol vsf限制类别
    End With
    
    txtEdit(mtxtIdx.idx_编号).MaxLength = 6
    txtEdit(mtxtIdx.idx_名称).MaxLength = 50
    txtEdit(mtxtIdx.idx_前缀文本).MaxLength = 2
    txtEdit(mtxtIdx.idx_卡号长度).MaxLength = 2
    InitData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
 End Function
 
Private Function LoadCardData() As Boolean
    '功能:加载卡片数据
    '返回:加载成功，返回true，否则返回False
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim rs卡类别 As ADODB.Recordset
    Dim strValue As String, intIndx As Integer
    Dim i As Long, j As Long
    
    On Error GoTo errHandle
    If mEditType = Card_增加 Then LoadCardData = True: Exit Function
    
    Set rs卡类别 = zlGet消费卡接口(, True)
    rs卡类别.Filter = "编号=" & mlngCardTypeID
    If rs卡类别.EOF Then
        MsgBox "未找到消费卡类别信息，可能已经被他人删除！", vbInformation, gstrSysName
        Exit Function
    End If
    
    txtEdit(mtxtIdx.idx_编号).Text = NVL(rs卡类别!编号)
    txtEdit(mtxtIdx.idx_名称).Text = NVL(rs卡类别!名称)
    txtEdit(mtxtIdx.idx_前缀文本).Text = NVL(rs卡类别!前缀文本)
    txtEdit(mtxtIdx.idx_卡号长度).Text = IIf(Val(NVL(rs卡类别!卡号长度)) = 0, 1, Val(NVL(rs卡类别!卡号长度)))
    
    cbo.SeekIndex cbo结算方式, NVL(rs卡类别!结算方式)
    If cbo结算方式.ListIndex < 0 Then
        cbo结算方式.AddItem NVL(rs卡类别!结算方式)
        cbo结算方式.ListIndex = cbo结算方式.NewIndex
    End If
    chkEdit(mchkIdx.idx_启用).value = IIf(Val(NVL(rs卡类别!启用)) = 1, 1, 0)
    
    chkEdit(mchkIdx.idx_退现).value = IIf(Val(NVL(rs卡类别!是否退现)) = 1, 1, 0)
    chkEdit(mchkIdx.idx_全退).value = IIf(Val(NVL(rs卡类别!是否全退)) = 1, 0, 1)
    chkEdit(mchkIdx.idx_密文).value = IIf(Val(NVL(rs卡类别!是否密文)) = 1, 1, 0)
    chkEdit(mchkIdx.idx_严格控制).value = IIf(Val(NVL(rs卡类别!是否严格控制)) = 1, 1, 0)
    chkEdit(mchkIdx.idx_特定病人).value = IIf(Val(NVL(rs卡类别!是否特定病人)) = 1, 1, 0)
    chkEdit(mchkIdx.idx_换卡).value = IIf(Val(NVL(rs卡类别!是否允许换卡)) = 1, 1, 0)
    chkEdit(mchkIdx.idx_补卡).value = IIf(Val(NVL(rs卡类别!是否允许补卡)) = 1, 1, 0)
    chkEdit(mchkIdx.idx_余额退款).value = IIf(Val(NVL(rs卡类别!是否允许余额退款)) = 1, 1, 0)
    
    strValue = NVL(rs卡类别!应用场合, "000")
    chkEdit(mchkIdx.idx_门诊).value = IIf(Val(Mid(strValue, 1, 1)) = 0, vbChecked, vbUnchecked)
    chkEdit(mchkIdx.idx_住院).value = IIf(Val(Mid(strValue, 2, 1)) = 0, vbChecked, vbUnchecked)
    chkEdit(mchkIdx.idx_体检).value = IIf(Val(Mid(strValue, 3, 1)) = 0, vbChecked, vbUnchecked)
    
    With vsf限制类别
        strValue = NVL(rs卡类别!限制类别)
        .Tag = strValue
        For i = 0 To .Rows - 1
            For j = 0 To .Cols - 1
                If .TextMatrix(i, j) <> "" Then
                    If InStr("," & strValue & ",", "," & .Cell(flexcpData, i, j) & ",") > 0 Then
                        .Cell(flexcpChecked, i, j) = 1
                    Else
                        .Cell(flexcpChecked, i, j) = 2
                    End If
                End If
            Next
        Next
    End With
    
    strValue = NVL(rs卡类别!读卡性质, "10")
    chkEdit(mchkIdx.idx_刷卡).value = Val(Mid(strValue, 1, 1))
    chkEdit(mchkIdx.idx_扫描卡).value = Val(Mid(strValue, 2, 1))
    
    intIndx = Val(NVL(rs卡类别!键盘控制方式))
    If intIndx < 0 Or intIndx > 2 Then intIndx = 0
    opt键盘控制(intIndx).value = True
    
    txtEdit(mtxtIdx.idx_密码长度).Text = Val(NVL(rs卡类别!密码长度))
    Select Case Val(NVL(rs卡类别!密码长度限制))
    Case 0
        optPassInput(0).value = True
    Case 1
        optPassInput(1).value = True
    Case Else '负数
        optPassInput(2).value = True
        txtEdit(mtxtIdx.idx_密码位数).Text = Abs(Val(NVL(rs卡类别!密码长度限制)))
    End Select
    intIndx = Val(NVL(rs卡类别!密码规则))
    If intIndx < 0 Or intIndx > 1 Then intIndx = 0
    optRule(intIndx).value = True
    
    With mCardType
        .lng卡号长度 = Val(NVL(rs卡类别!卡号长度))
        .bln固定 = Val(NVL(rs卡类别!系统)) = 1
        .str前缀文本 = NVL(rs卡类别!前缀文本)
        
        strSQL = "Select 1 From 消费卡信息 Where 接口编号=[1] And Rownum < 2"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngCardTypeID)
        .bln已发卡 = Not rsTemp.EOF
    End With
    
    LoadCardData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetCtrlEnable()
    '功能:设置控件的编辑属性
    '编制:刘兴洪
    Dim i As Long, blnEdit As Boolean
    
    On Error GoTo ErrHandler
    blnEdit = (mEditType = Card_增加 Or mEditType = Card_修改)
    For i = 0 To txtEdit.UBound
        Select Case i
        Case mtxtIdx.idx_编号
            txtEdit(i).Enabled = mEditType = Card_增加
        Case mtxtIdx.idx_名称
            txtEdit(i).Enabled = blnEdit And Not mCardType.bln固定
        Case mtxtIdx.idx_密码位数
            txtEdit(i).Enabled = False
        Case Else
            txtEdit(i).Enabled = blnEdit
        End Select
    Next
    
    For i = 0 To chkEdit.UBound
        chkEdit(i).Enabled = blnEdit
    Next
    Call chkEdit_Click(mchkIdx.idx_特定病人)
    
    cbo结算方式.Enabled = blnEdit
    vsf限制类别.Enabled = blnEdit
    vsf限制类别.Editable = IIf(blnEdit, flexEDKbdMouse, flexEDNone)
    
    optPassInput(0).Enabled = blnEdit
    optPassInput(1).Enabled = blnEdit
    optPassInput(2).Enabled = blnEdit
    
    optRule(0).Enabled = blnEdit
    optRule(1).Enabled = blnEdit
    
    Call SetEnabledBackColor(Me)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo结算方式_Change()
    mblnChange = True
End Sub

Private Sub cbo结算方式_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub lvw限制类别_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    mblnChange = True
End Sub

Private Sub lvw限制类别_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optRule_Click(Index As Integer)
    mblnChange = True
End Sub

Private Sub opt键盘控制_Click(Index As Integer)
    mblnChange = True
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = mtxtIdx.idx_密码长度 Then
        optPassInput(1).Caption = "固定输入" & Val(txtEdit(Index).Text) & "位"
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    If Index = mtxtIdx.idx_名称 Then
        zlCommFun.OpenIme True
    End If
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then txtEdit(Index).Text = ""
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = mtxtIdx.idx_卡号长度 Or Index = mtxtIdx.idx_编号 _
        Or Index = mtxtIdx.idx_密码位数 Or Index = mtxtIdx.idx_密码长度 Then
        zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m数字式
    ElseIf Index = mtxtIdx.idx_前缀文本 Then
        If zlStr.IsCharChinese(Chr(KeyAscii)) Then KeyAscii = 0
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    If Index = mtxtIdx.idx_名称 Then
        zlCommFun.OpenIme False
    ElseIf Index = mtxtIdx.idx_前缀文本 Then
        txtEdit(Index).Text = UCase(txtEdit(Index).Text)
    End If
End Sub

Private Sub txtEdit_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        glngTXTProc = GetWindowLong(txtEdit(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtEdit(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtEdit_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Call SetWindowLong(txtEdit(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub chkEdit_Click(Index As Integer)
    Dim blnEnabled As Boolean
    
    mblnChange = True
    
    '至少保留一项
    Select Case Index
    Case mchkIdx.idx_门诊, mchkIdx.idx_住院, mchkIdx.idx_体检
        CheckCheckboxValue Array(chkEdit(mchkIdx.idx_门诊), chkEdit(mchkIdx.idx_住院), chkEdit(mchkIdx.idx_体检))
    Case mchkIdx.idx_刷卡, mchkIdx.idx_扫描卡
        'CheckCheckboxValue Array(chkEdit(mchkIdx.idx_刷卡), chkEdit(mchkIdx.idx_扫描卡))
        If chkEdit(Index).value = vbUnchecked Then
            If Index = mchkIdx.idx_刷卡 Then
                chkEdit(mchkIdx.idx_扫描卡).value = vbChecked
            Else
                chkEdit(mchkIdx.idx_刷卡).value = vbChecked
            End If
        End If
    Case mchkIdx.idx_特定病人
        blnEnabled = chkEdit(mchkIdx.idx_特定病人).value
        chkEdit(mchkIdx.idx_补卡).Enabled = blnEnabled
    End Select
End Sub

Private Sub CheckCheckboxValue(ByVal varCheckBox As Variant)
    '设置一组CkeckBox，必须保证其中一个是勾选的
    Dim i As Integer, blnChecked As Boolean
    
    On Error GoTo ErrHandler
    For i = 0 To UBound(varCheckBox)
        If varCheckBox(i).value Then
            blnChecked = True: Exit For
        End If
    Next
    
    If blnChecked = False Then
        varCheckBox(0).value = vbChecked
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub optPassInput_Click(Index As Integer)
    mblnChange = True
    
    txtEdit(mtxtIdx.idx_密码位数).Enabled = optPassInput(2).value
    zl_SetCtlBackColor txtEdit(mtxtIdx.idx_密码位数), Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ErrHandler
    If isValied() = False Then Exit Sub
    If SaveData() = False Then Exit Sub
    
    mintSucces = mintSucces + 1
    mblnChange = False
    Unload Me
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
 
Private Function isValied() As Boolean
    '功能:检查数据的有效性
    '返回:数据有效，返回true,否则返回False
    Dim i As Long, strTemp As String
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    If zlControl.FormCheckInput(Me) = False Then Exit Function
    If zlControl.TxtCheckInput(txtEdit(mtxtIdx.idx_编号), "编号") = False Then Exit Function
    If zlControl.TxtCheckInput(txtEdit(mtxtIdx.idx_名称), "名称") = False Then Exit Function
    If zlControl.TxtCheckInput(txtEdit(mtxtIdx.idx_前缀文本), "前缀文本", , True) = False Then Exit Function
    If zlControl.TxtCheckInput(txtEdit(mtxtIdx.idx_卡号长度), "卡号长度") = False Then Exit Function
    
    If zlStr.IsCharChinese(txtEdit(mtxtIdx.idx_前缀文本)) Then
        ShowMsgbox "前缀文本不能包含汉字！"
        zlControl.ControlSetFocus txtEdit(mtxtIdx.idx_前缀文本)
        Exit Function
    End If
    
    If Val(txtEdit(mtxtIdx.idx_卡号长度).Text) < 1 Then
        ShowMsgbox "卡号长度必须大于等于1位！"
        zlControl.ControlSetFocus txtEdit(mtxtIdx.idx_卡号长度)
        Exit Function
    End If
    
    If zlCommFun.ActualLen(Trim(txtEdit(idx_前缀文本))) + Val(txtEdit(mtxtIdx.idx_卡号长度).Text) > 20 Then
        ShowMsgbox "卡号的最大长度(前缀+卡号长度)不能大于20位，请检查！"
        zlControl.ControlSetFocus txtEdit(mtxtIdx.idx_卡号长度)
        Exit Function
    End If
    
    If mCardType.bln已发卡 Then
        If Val(txtEdit(idx_卡号长度).Text) + Len(Trim(txtEdit(idx_前缀文本))) < mCardType.lng卡号长度 + Len(NVL(mCardType.str前缀文本)) Then
            ShowMsgbox "由于发生了发卡信息,所以消费卡前缀文本及卡号长度不能减小,请检查！"
            zlControl.ControlSetFocus txtEdit(mtxtIdx.idx_卡号长度)
            Exit Function
        End If
    End If
    
    If cbo结算方式.ListIndex < 0 Then
        ShowMsgbox "结算方式必须选择！"
        zlControl.ControlSetFocus cbo结算方式
        Exit Function
    End If
    
    strSQL = _
        "Select 名称 From 医疗卡类别 Where 结算方式 = [2]" & vbNewLine & _
        "Union All" & vbNewLine & _
        "Select 名称 From 消费卡类别目录 Where 编号 <> [1] And 结算方式 = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngCardTypeID, cbo结算方式.Text)
    If Not rsTemp.EOF Then
        ShowMsgbox "结算方式『" & cbo结算方式.Text & "』已被" & NVL(rsTemp!名称) & "使用，" & _
                   "重复使用会造成财务扎帐紊乱，请重新选定一种结算方式！"
        zlControl.ControlSetFocus cbo结算方式
        Exit Function
    End If
    
    If Val(txtEdit(mtxtIdx.idx_密码长度).Text) = 0 Then
        ShowMsgbox "密码长度不能设置为零！"
        zlControl.ControlSetFocus txtEdit(mtxtIdx.idx_密码长度)
        Exit Function
    End If
    If Val(txtEdit(mtxtIdx.idx_密码长度).Text) > 50 Then
        ShowMsgbox "密码长度不能大于50位！"
        zlControl.ControlSetFocus txtEdit(mtxtIdx.idx_密码长度)
        Exit Function
    End If
    If optPassInput(2).value Then
        If Val(txtEdit(mtxtIdx.idx_密码长度).Text) < Val(txtEdit(mtxtIdx.idx_密码位数).Text) Then
            ShowMsgbox "必须输入的密码长度不能大于总的密码长度(" & Val(txtEdit(mtxtIdx.idx_密码长度).Text) & ")位！"
            zlControl.ControlSetFocus txtEdit(mtxtIdx.idx_密码位数)
            Exit Function
        End If
    End If
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveData() As Boolean
    '功能:保存数据
    '返回:保存成功,返回true,否则返回False
    Dim strSQL As String
    Dim strValue As String

    On Error GoTo errHandle
    'Zl_消费卡类别目录_Update
    strSQL = "Zl_消费卡类别目录_Update("
    '  编码_In         In 消费卡类别目录.编号%Type,
    strSQL = strSQL & "'" & Trim(txtEdit(mtxtIdx.idx_编号).Text) & "',"
    '  名称_In         In 消费卡类别目录.名称%Type,
    strSQL = strSQL & "'" & Trim(txtEdit(mtxtIdx.idx_名称).Text) & "',"
    '  结算方式_In     In 消费卡类别目录.结算方式%Type,
    strSQL = strSQL & "'" & cbo结算方式.Text & "',"
    '  前缀文本_In     In 消费卡类别目录.前缀文本%Type,
    strSQL = strSQL & "'" & Trim(txtEdit(mtxtIdx.idx_前缀文本).Text) & "',"
    '  卡号长度_In     In 消费卡类别目录.卡号长度%Type,
    strSQL = strSQL & "" & Val(txtEdit(mtxtIdx.idx_卡号长度).Text) & ","
    '  卡号密文_In     In 消费卡类别目录.是否密文%Type,
    strSQL = strSQL & "" & IIf(chkEdit(idx_密文).value = vbChecked, "1", "0") & ","
    '  是否退现_In     In 消费卡类别目录.是否全退%Type,
    strSQL = strSQL & "" & IIf(chkEdit(idx_退现).value = vbChecked, "1", "0") & ","
    '  是否全退_In     In 消费卡类别目录.是否全退%Type,
    strSQL = strSQL & "" & IIf(chkEdit(idx_全退).value = vbChecked, "0", "1") & ","
    '  启用_In         In 消费卡类别目录.启用%Type,
    strSQL = strSQL & "" & IIf(chkEdit(idx_启用).value = vbChecked, "1", "0") & ","
    '  密码长度_In     In 消费卡类别目录.密码长度%Type,
    strSQL = strSQL & "" & Val(txtEdit(mtxtIdx.idx_密码长度).Text) & ","
    '  密码长度限制_In In 消费卡类别目录.密码长度限制%Type,
    If optPassInput(0).value Then
        strSQL = strSQL & "" & 0 & ","
    ElseIf optPassInput(1).value Then
        strSQL = strSQL & "" & 1 & ","
    Else
        strSQL = strSQL & "" & -1 * Val(txtEdit(mtxtIdx.idx_密码位数).Text) & ","
    End If
    '  密码规则_In     In 消费卡类别目录.密码规则%Type,
    If optRule(0).value Then
        strSQL = strSQL & "" & 0 & ","
    Else
        strSQL = strSQL & "" & 1 & ","
    End If
    '  操作方式_In     In Integer := 0
    strSQL = strSQL & "" & IIf(mEditType = Card_增加, 0, 1) & ","
    '  读卡性质_In         In 消费卡类别目录.读卡性质%Type,
    strValue = IIf(chkEdit(mchkIdx.idx_刷卡).value = vbChecked, "1", "0")
    strValue = strValue & IIf(chkEdit(mchkIdx.idx_扫描卡).value = 1, "1", "0")
    strSQL = strSQL & "'" & strValue & "',"
    '  键盘控制方式_In     In 消费卡类别目录.键盘控制方式%Type,
    If opt键盘控制(0).value Then
        strSQL = strSQL & "" & 0 & ","
    ElseIf opt键盘控制(1).value Then
        strSQL = strSQL & "" & 1 & ","
    ElseIf opt键盘控制(2).value Then
        strSQL = strSQL & "" & 2 & ","
    End If
    '  限制类别_In         In 消费卡类别目录.限制类别%Type,
    strSQL = strSQL & "'" & Get限制类别() & "',"
    '  是否严格控制_In     In 消费卡类别目录.是否严格控制%Type,
    strSQL = strSQL & "" & IIf(chkEdit(idx_严格控制).value = vbChecked, "1", "0") & ","
    '  是否特定病人_In     In 消费卡类别目录.是否特定病人%Type,
    strSQL = strSQL & "" & IIf(chkEdit(idx_特定病人).value = vbChecked, "1", "0") & ","
    '  是否允许换卡_In     In 消费卡类别目录.是否允许换卡%Type,
    strSQL = strSQL & "" & IIf(chkEdit(idx_换卡).value = vbChecked, "1", "0") & ","
    '  是否允许补卡_In     In 消费卡类别目录.是否允许补卡%Type,
    strSQL = strSQL & "" & IIf(chkEdit(idx_特定病人).value = vbChecked And chkEdit(idx_补卡).value = vbChecked, "1", "0") & ","
    '  是否允许余额退款_In In 消费卡类别目录.是否允许余额退款%Type,
    strSQL = strSQL & "" & IIf(chkEdit(idx_余额退款).value = vbChecked, "1", "0") & ","
    '  应用场合_In         In 消费卡类别目录.应用场合%Type
    strValue = IIf(chkEdit(mchkIdx.idx_门诊).value = vbChecked, "0", "1")
    strValue = strValue & IIf(chkEdit(mchkIdx.idx_住院).value = vbChecked, "0", "1")
    strValue = strValue & IIf(chkEdit(mchkIdx.idx_体检).value = vbChecked, "0", "1")
    strSQL = strSQL & "'" & strValue & "')"
    
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Get限制类别() As String
    '获取限制类别
    Dim strType As String, i As Long, j As Long
    
    On Error GoTo ErrHandler
    With vsf限制类别
         For i = 0 To .Rows - 1
            For j = 0 To .Cols - 1
                If Abs(Val(.Cell(flexcpChecked, i, j))) = 1 Then
                    strType = strType & "," & .Cell(flexcpData, i, j)
                End If
            Next
         Next
         If strType <> "" Then strType = Mid(strType, 2)
    End With
    Get限制类别 = strType
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsf限制类别_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    mblnChange = True
End Sub

Private Sub vsf限制类别_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Row < 0 Or Col < 0 Then Exit Sub
    If vsf限制类别.TextMatrix(Row, Col) = "" Then Cancel = True
End Sub

Private Sub vsf限制类别_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewRow < 0 Or NewCol < 0 Then Exit Sub
    If vsf限制类别.TextMatrix(NewRow, NewCol) = "" Then Cancel = True
End Sub

Private Sub vsf限制类别_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub
