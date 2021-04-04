VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmSquareRefundBalance 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "余额退款 - 消费卡"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8580
   Icon            =   "frmSquareRefundBalance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra退款面板 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   30
      TabIndex        =   5
      Top             =   3390
      Width           =   8505
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   1785
         Index           =   2
         Left            =   2610
         ScaleHeight     =   1755
         ScaleWidth      =   5865
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   60
         Width           =   5895
         Begin VB.TextBox txt 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            IMEMode         =   3  'DISABLE
            Index           =   4
            Left            =   930
            TabIndex        =   20
            Tag             =   "1"
            Top             =   960
            Width           =   4860
         End
         Begin VB.TextBox txt 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   930
            TabIndex        =   22
            Tag             =   "1"
            Top             =   1380
            Width           =   4860
         End
         Begin VB.TextBox txt 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   3
            Left            =   930
            TabIndex        =   18
            Tag             =   "1"
            Top             =   540
            Width           =   4860
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   1950
            TabIndex        =   14
            Top             =   112
            Width           =   1410
         End
         Begin VB.TextBox txt 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Index           =   2
            Left            =   4200
            Locked          =   -1  'True
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   120
            Width           =   1590
         End
         Begin VB.ComboBox cbo支付方式 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   930
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   112
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "帐  号"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   6
            Left            =   270
            TabIndex        =   19
            Top             =   1012
            Width           =   630
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "结算号码"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   7
            Left            =   60
            TabIndex        =   21
            Top             =   1440
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "开户行"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   5
            Left            =   270
            TabIndex        =   17
            Top             =   592
            Width           =   630
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "退 款"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Index           =   3
            Left            =   330
            TabIndex        =   12
            Top             =   165
            Width           =   570
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "找 补"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Index           =   4
            Left            =   3600
            TabIndex        =   15
            Top             =   165
            Width           =   570
         End
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   1785
         Index           =   1
         Left            =   0
         ScaleHeight     =   1755
         ScaleWidth      =   2535
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   60
         Width           =   2565
         Begin XtremeSuiteControls.ShortcutCaption ShortcutCaption1 
            Height          =   315
            Left            =   15
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   870
            Width           =   2505
            _Version        =   589884
            _ExtentX        =   4419
            _ExtentY        =   556
            _StockProps     =   6
            Caption         =   "退费合计"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   9
            Left            =   1710
            TabIndex        =   10
            Top             =   1335
            Width           =   660
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   2
            Left            =   1725
            TabIndex        =   8
            Top             =   450
            Width           =   660
         End
         Begin XtremeSuiteControls.ShortcutCaption lbl退款合计 
            Height          =   315
            Left            =   15
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   15
            Width           =   2505
            _Version        =   589884
            _ExtentX        =   4419
            _ExtentY        =   556
            _StockProps     =   6
            Caption         =   "当前未退"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
         End
      End
   End
   Begin VB.Frame fraButtons 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   795
      Left            =   30
      TabIndex        =   26
      Top             =   5250
      Width           =   8505
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   350
         Left            =   210
         TabIndex        =   25
         Top             =   210
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   5790
         TabIndex        =   23
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   6990
         TabIndex        =   24
         Top             =   240
         Width           =   1100
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "本次误差：0.00"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   8
         Left            =   4020
         TabIndex        =   28
         Top             =   325
         Visible         =   0   'False
         Width           =   1395
      End
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3345
      Index           =   0
      Left            =   30
      ScaleHeight     =   3315
      ScaleWidth      =   8475
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   8505
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   795
         TabIndex        =   2
         Top             =   120
         Width           =   2055
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfBlance 
         Height          =   2715
         Left            =   60
         TabIndex        =   4
         Top             =   510
         Width           =   8325
         _cx             =   14684
         _cy             =   4789
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
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483644
         GridColorFixed  =   -2147483648
         TreeColor       =   -2147483643
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmSquareRefundBalance.frx":000C
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
         Editable        =   2
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
      Begin VB.Label lblPatiInfo 
         AutoSize        =   -1  'True
         Height          =   180
         Left            =   3000
         TabIndex        =   27
         Top             =   180
         Width           =   90
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "余额：500.00"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Index           =   1
         Left            =   6990
         TabIndex        =   3
         Tag             =   "余额："
         Top             =   165
         Width           =   1395
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "卡号(&N)"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmSquareRefundBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'入口参数
Private mfrmMain As Form, mlngModule As Long, mstrPrivs As String
Private mlng卡类别 As Long, mlng卡ID As Long

'模块变量
Private mblnFirst As Boolean, mintSucces As Integer
Private mblnNotClick As Boolean

Private Enum mLableIndex
    lbl_余额 = 1
    lbl_当前未退 = 2
    lbl_退款合计 = 9
    lbl_找补 = 4
    lbl_误差 = 8
End Enum
Private Enum mTextIndex
    txt_卡号 = 0
    txt_金额 = 1
    txt_找补 = 2
    txt_开户行 = 3
    txt_帐号 = 4
    txt_结算号码 = 5
End Enum
Private Enum mPictureIndex
    pic_余额明细 = 0
    pic_缴款合计 = 1
    pic_缴款信息 = 2
End Enum

Private Type Ty_CardType
    str卡名称 As String
    str卡号前缀 As String
    lng卡号长度 As Long
    bln卡号密文 As Boolean
    bln严格控制 As Boolean
    str限制类别 As String
    bln特定病人 As Boolean
    lng共用批次 As Long
    lng领用ID As Long
End Type
Private mCardType As Ty_CardType

'支付相关
Private mobjPayCards As Cards
Private mlngPre支付方式 As Long
Private Type TY_PayMoney
    dbl退款合计 As Double
    dbl当前未退 As Double
    dbl本次误差 As Double
    str原结算序号 As String
    
    lng卡类别ID As Long
    str刷卡卡号 As String
    str刷卡密码 As String
    str交易流水号 As String
    str交易说明 As String
End Type
Private mCurCardPay As TY_PayMoney '本次卡支付
Private mBytMoney As Byte '分币处理规则

Public Function ShowMe(ByVal frmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String, _
    ByVal lng卡类别 As Long) As Boolean
    '程序入口
    '入参：
    '   frmMain - 父窗口
    '   lngModule - 模块号
    '   strPrivs - 权限串
    '   lng卡类别 As Long - 消费卡类别
    '返回：操作成功返回True,否则返回False
    Set mfrmMain = frmMain: mlngModule = lngModule: mstrPrivs = strPrivs:
    mlng卡类别 = lng卡类别
    mlng卡ID = 0
    
    mintSucces = 0
    On Error Resume Next
    Me.Show 1, frmMain
    ShowMe = mintSucces > 0
End Function

Private Function CardIsValid(ByVal bytMode As Byte, Optional ByVal lng卡ID As Long, _
    Optional ByVal str卡号 As String, Optional ByVal blnSaveAfter As Boolean) As Boolean
    '检查卡信息
    '入参：
    '   bytMode 0-按消费卡ID加载，1-按消费卡卡号加载
    '   blnSaveAfter 是否保存数据前检查
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strWhere As String, dbl余额 As Double, dbl失效面额 As Double
    
    On Error GoTo ErrHandler
    If bytMode = 1 Then
        strWhere = " And a.卡号 = [2] And a.接口编号 = [3]" & vbNewLine & _
                   " And a.序号 = (Select Max(序号) From 消费卡信息 Where 卡号 = a.卡号 And 接口编号 = a.接口编号)"
    Else
        strWhere = " And a.Id = [1]"
    End If
    
    strSQL = _
        "Select a.ID, a.可否充值, a.卡号, a.序号,to_char(a.有效期,'yyyy-mm-dd hh24:mi:ss') as 有效期, " & vbNewLine & _
        "       (Select Max(序号) From 消费卡信息 B Where a.卡号 = b.卡号 And a.接口编号 = b.接口编号) As 最大序号," & vbNewLine & _
        "       To_Char(a.回收时间, 'yyyy-mm-dd hh24:mi:ss') As 回收时间, " & vbNewLine & _
        "       To_Char(a.停用日期, 'yyyy-mm-dd hh24:mi:ss') As 停用日期, a.余额," & vbNewLine & _
        "       b.姓名, b.性别, b.年龄" & vbNewLine & _
        "From 消费卡信息 A,病人信息 B" & vbNewLine & _
        "Where a.病人ID = b.病人ID(+) " & strWhere
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng卡ID, str卡号, mlng卡类别)
    
    If rsTemp.EOF Then
        ShowMsgbox "未找到相关的" & mCardType.str卡名称 & "信息，可能已经被他人删除！"
        Exit Function
    End If
    
    str卡号 = Nvl(rsTemp!卡号)
    '检查卡号是否合法
    If Val(Nvl(rsTemp!序号)) < Val(Nvl(rsTemp!最大序号)) Then
        ShowMsgbox "不能对历史卡号进行余额退款(卡号为:" & str卡号 & ")！"
        Exit Function
    End If
    
    If Nvl(rsTemp!回收时间, "3000-01-01") < "3000-01-01" Then
        ShowMsgbox "卡号为:" & str卡号 & " 的" & mCardType.str卡名称 & "已被回收，不能再余额退款！"
        Exit Function
    End If
    
    '停用的也可以回收和取消回收
    If Nvl(rsTemp!停用日期, "3000-01-01") < "3000-01-01" Then
        ShowMsgbox "卡号为:" & str卡号 & " 的" & mCardType.str卡名称 & "已经停止使用，不能再余额退款！"
        Exit Function
    End If
    
    dbl余额 = Val(Nvl(rsTemp!余额))
    dbl失效面额 = 0
    '检查效期
    If Nvl(rsTemp!有效期, "3000-01-01 00:00:00") < Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") Then
        If Val(Nvl(rsTemp!可否充值)) = 1 Then
            '允许充值的，到期的，不能退款
            dbl失效面额 = zlGet失效面额(Val(Nvl(rsTemp!id)))
            dbl余额 = dbl余额 - dbl失效面额
            If dbl余额 <= 0 Then dbl余额 = 0
        End If
    End If
    
    If dbl余额 <= 0 Then
        ShowMsgbox "卡号为:" & str卡号 & " 的" & mCardType.str卡名称 & "当前无余额，不能进行余额退款！"
        Exit Function
    End If
    If blnSaveAfter Then CardIsValid = True: Exit Function
    
    If bytMode = 1 Then
        mlng卡ID = Val(Nvl(rsTemp!id))
    Else
        txt(txt_卡号).Text = str卡号
    End If
    lbl(lbl_余额).Caption = lbl(lbl_余额).Tag & Format(Val(Nvl(rsTemp!余额)), "0.00") & _
        IIf(dbl失效面额 > 0, "(已失效:" & Format(dbl失效面额, "0.00") & ")", "")
    
'    If nvl(rsTemp!姓名) = "" Then
'        lblPatiInfo.Caption = ""
'    Else
'        lblPatiInfo.Caption = "姓名：" & nvl(rsTemp!姓名) & " 性别：" & nvl(rsTemp!性别) & " 年龄：" & nvl(rsTemp!年龄)
'    End If
    
    CardIsValid = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cbo支付方式_Click()
    On Error GoTo ErrHandler
    If mblnNotClick Then Exit Sub
    If mlngPre支付方式 = cbo支付方式.ItemData(cbo支付方式.ListIndex) Then Exit Sub
    mlngPre支付方式 = cbo支付方式.ItemData(cbo支付方式.ListIndex)
    
    txt(txt_金额).Text = ""
    Call SetControlProperty
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetControlProperty(Optional ByVal blnLoadDefault As Boolean)
    '设置控件属性
    '入参:
    '   blnLoadDefault-是否加载缺省值
    Dim objCard As Card
    Dim blnEnabled As Boolean
    Dim dblTemp As Double, dblMoney As Double
    
    On Error GoTo ErrHandler
    Set objCard = GetCurCard()
    
    '支票、一卡通和老版一卡通允许输入缴款单位
    '1-现金结算方式,2-其他非医保结算,3-医保个人帐户,4-医保各类统筹,5-代收款项,6-费用折扣,7-一卡通结算,8-结算卡结算
    blnEnabled = InStr(",2,7,8,", "," & objCard.结算性质 & ",") > 0
    txt(txt_开户行).Enabled = objCard.结算性质 <> 1
    txt(txt_帐号).Enabled = objCard.结算性质 <> 1
    txt(txt_结算号码).Enabled = objCard.结算性质 <> 1
    If objCard.结算性质 = 1 Then
        txt(txt_开户行).Text = ""
        txt(txt_帐号).Text = ""
        txt(txt_结算号码).Text = ""
        
        dblMoney = CentMoney(mCurCardPay.dbl当前未退, mBytMoney)
    Else
        dblMoney = RoundEx(mCurCardPay.dbl当前未退, 2)
    End If
    mCurCardPay.dbl本次误差 = mCurCardPay.dbl当前未退 - dblMoney
    
    Call zl_SetCtlBackColor(Array(txt(txt_开户行), txt(txt_帐号), txt(txt_结算号码)), Me)
                
    '缺省金额的设置
    txt(txt_金额).Locked = False
    If objCard.接口序号 > 0 Then '三方结算
        txt(txt_金额).Text = Format(dblMoney, "0.00")
        txt(txt_金额).Locked = True
    ElseIf objCard.结算性质 = 1 Then '现金处理
        txt(txt_金额).Text = Format(dblMoney, "0.00")
    Else
        txt(txt_金额).Text = Format(dblMoney, "0.00")
        txt(txt_金额).Locked = True
    End If
    lbl(lbl_误差).Caption = FormatEx(mCurCardPay.dbl本次误差, 6, , , 2)
    lbl(lbl_误差).Visible = Val(lbl(lbl_误差).Caption) <> 0
    lbl(lbl_误差).Caption = "本次误差：" & lbl(lbl_误差).Caption
    lbl(lbl_当前未退).Caption = Format(mCurCardPay.dbl当前未退 - mCurCardPay.dbl本次误差, "0.00")
    
    '计算找补
    Call SetLblCaption
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbo支付方式_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub cmdOK_Click()
    Dim lngRow As Long
    Dim objCard As Card, lng卡类别ID As Long
    Dim str退款信息 As String, lng结算序号 As Long
    Dim dblDelMoney As Double, str交易序号 As String
    
    On Error GoTo ErrHandler
    If mlng卡ID = 0 Then
        ShowMsgbox "请正确录入卡号！"
        zlControl.ControlSetFocus txt(txt_卡号)
        Exit Sub
    End If
    If CardIsValid(0, mlng卡ID, , True) = False Then Exit Sub
    If Check缴款情况 = False Then Exit Sub
    
    lng结算序号 = zlDatabase.GetNextId("病人卡结算记录")
    '三方卡退回部分
    With vsfBlance
        For lngRow = 1 To .Rows - 1
            lng卡类别ID = Val(.Cell(flexcpData, lngRow, .ColIndex("卡号")))
            dblDelMoney = Val(.Cell(flexcpData, lngRow, .ColIndex("退款金额")))
            If lng卡类别ID = 0 Then Exit For
            If Val(.Cell(flexcpChecked, lngRow, .ColIndex("退现"))) <> 1 And dblDelMoney > 0 Then
                Set objCard = GetCurCard(lng卡类别ID)
                str交易序号 = .Cell(flexcpData, lngRow, .ColIndex("结算方式"))
                
                mCurCardPay.str原结算序号 = ""
                mCurCardPay.lng卡类别ID = lng卡类别ID
                mCurCardPay.str刷卡卡号 = .TextMatrix(lngRow, .ColIndex("卡号"))
                mCurCardPay.str刷卡密码 = ""
                mCurCardPay.str交易流水号 = .TextMatrix(lngRow, .ColIndex("交易流水号"))
                mCurCardPay.str交易说明 = .TextMatrix(lngRow, .ColIndex("交易说明"))
                If CheckThreeSwapIsValied(objCard, dblDelMoney, str交易序号) = False Then GoTo ErrCheckDelAll
                If SaveData(objCard, dblDelMoney, str交易序号, lng结算序号) = False Then GoTo ErrCheckDelAll
                
                str退款信息 = str退款信息 & vbCrLf & _
                    .TextMatrix(lngRow, .ColIndex("结算方式")) & ":" & Format(dblDelMoney, "0.00")
            End If
        Next
    End With
    
    '退现部分，可能支持转账及代扣
    str交易序号 = ""
    With vsfBlance
        For lngRow = 1 To .Rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("退款金额"))) > 0 Then
                If Val(.Cell(flexcpData, lngRow, .ColIndex("卡号"))) = 0 Then
                    str交易序号 = str交易序号 & "," & .Cell(flexcpData, lngRow, .ColIndex("结算方式"))
                ElseIf Val(.Cell(flexcpChecked, lngRow, .ColIndex("退现"))) = 1 Then
                    str交易序号 = str交易序号 & "," & .Cell(flexcpData, lngRow, .ColIndex("结算方式"))
                End If
            End If
        Next
        If str交易序号 <> "" Then str交易序号 = Mid(str交易序号, 2)
    End With
    Set objCard = GetCurCard()
    mCurCardPay.str原结算序号 = ""
    mCurCardPay.lng卡类别ID = IIf(objCard.接口序号 > 0, objCard.接口序号, 0)
    mCurCardPay.str刷卡卡号 = ""
    mCurCardPay.str刷卡密码 = ""
    mCurCardPay.str交易流水号 = ""
    mCurCardPay.str交易说明 = ""
    If CheckThreeSwapIsValied(objCard, mCurCardPay.dbl当前未退, str交易序号) = False Then GoTo ErrCheckDelAll
    If SaveData(objCard, mCurCardPay.dbl当前未退, str交易序号, lng结算序号, mCurCardPay.dbl本次误差) = False Then GoTo ErrCheckDelAll
    
    mintSucces = mintSucces + 1
    Unload Me
    Exit Sub
ErrCheckDelAll:
    '如果中途失败，则需要刷新界面
    If str退款信息 <> "" Then
        If MsgBox("已成功退款部分如下：" & str退款信息 & vbCrLf & "是否对剩下未成功部分继续退费？", _
            vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            If CardIsValid(0, mlng卡ID) Then
                If LoadCardData(mlng卡ID) = False Then
                    mlng卡ID = 0: Call ClearData
                    zlControl.ControlSetFocus txt(txt_卡号): Exit Sub
                End If
            Else
                mlng卡ID = 0: Call ClearData
                zlControl.ControlSetFocus txt(txt_卡号): Exit Sub
            End If
        Else
            mintSucces = mintSucces + 1
            Unload Me
        End If
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SaveData(ByVal objCard As Card, ByVal dblDelMoney As Double, _
     ByVal str交易序号 As String, ByVal lng结算序号 As Long, _
     Optional ByVal dbl误差费 As Double) As Boolean
    '保存数据
    '入参：
    '   str交易序号 - 多个用逗号分隔
    Dim strSQL As String, blnTrain As Boolean
    
    On Error GoTo ErrHandler
    'Zl_消费卡信息_余额退款
    strSQL = "Zl_消费卡信息_余额退款("
    '  消费卡id_In   消费卡信息.Id%Type,
    strSQL = strSQL & "" & mlng卡ID & ","
    '  交易序号_In   Varchar2,
    strSQL = strSQL & "'" & str交易序号 & "',"
    '  结算方式_In   帐户缴款余额.结算方式%Type,
    strSQL = strSQL & "'" & objCard.结算方式 & "',"
    '  退款金额_In   病人卡结算记录.实收金额%Type,
    strSQL = strSQL & "" & dblDelMoney & ","
    '  误差金额_In   病人卡结算记录.实收金额%Type,
    strSQL = strSQL & "" & dbl误差费 & ","
    '  退款时间_In   消费卡信息.回收时间%Type,
    strSQL = strSQL & "To_Date('" & Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss") & "','yyyy-mm-dd hh24:mi:ss'),"
    '  操作员编号_In 病人卡结算记录.操作员编号%Type,
    strSQL = strSQL & "'" & UserInfo.编号 & "',"
    '  操作员姓名_In 病人卡结算记录.操作员姓名%Type,
    strSQL = strSQL & "'" & UserInfo.姓名 & "',"
    '  结算序号_In   病人卡结算记录.结算序号%Type,
    strSQL = strSQL & "" & lng结算序号 & ","
    '  开户行_In       病人卡结算记录.单位开户行%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_开户行).Text) & "',"
    '  帐号_In         病人卡结算记录.单位帐号%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_帐号).Text) & "',"
    '  结算号码_In   病人卡结算记录.结算号码%Type := Null,
    strSQL = strSQL & "'" & Trim(txt(txt_结算号码).Text) & "',"
    '  卡类别id_In   病人卡结算记录.卡类别id%Type := Null,
    strSQL = strSQL & "" & IIf(mCurCardPay.lng卡类别ID = 0, "NULL", mCurCardPay.lng卡类别ID) & ","
    '  结算卡号_In   病人卡结算记录.结算卡号%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng卡类别ID = 0, "NULL", "'" & mCurCardPay.str刷卡卡号 & "'") & ","
    '  交易流水号_In 病人卡结算记录.交易流水号%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng卡类别ID = 0, "NULL", "'" & mCurCardPay.str交易流水号 & "'") & ","
    '  交易说明_In   病人卡结算记录.交易说明%Type := Null,
    strSQL = strSQL & IIf(mCurCardPay.lng卡类别ID = 0, "NULL", "'" & mCurCardPay.str交易说明 & "'") & ","
    '  缴款_In       病人卡结算记录.缴款%Type := Null,
    strSQL = strSQL & "" & IIf(objCard.结算性质 = 1, -1 * Round(Val(txt(txt_金额).Text), 4), "NULL") & ","
    '  找补_In       病人卡结算记录.找补%Type := Null
    strSQL = strSQL & "" & IIf(objCard.结算性质 = 1, -1 * Round(Val(txt(txt_找补).Tag), 4), "NULL") & ")"

    blnTrain = True
    gcnOracle.BeginTrans
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    '三方卡结算
    If objCard.接口序号 > 0 Then
        If ExecuteThreeSwapPay(objCard, lng结算序号, dblDelMoney) = False Then Exit Function
    Else
        gcnOracle.CommitTrans
    End If
    blnTrain = False
    SaveData = True
    Exit Function
ErrHandler:
    If blnTrain Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetLblCaption()
    '设置找补的显示
    Dim dbl找补 As Double
    
    On Error GoTo ErrHandler
    dbl找补 = RoundEx(Val(txt(txt_金额).Text) - (mCurCardPay.dbl当前未退 - mCurCardPay.dbl本次误差), 6)
    txt(txt_找补).Tag = dbl找补
    txt(txt_找补).Text = Format(-1 * dbl找补, "0.00")
    lbl(lbl_找补).ForeColor = IIf(dbl找补 <= 0, vbBlack, vbRed)
    txt(txt_找补).ForeColor = IIf(dbl找补 <= 0, vbBlack, vbRed)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function LoadCardData(ByVal lng卡ID As Long) As Boolean
    '加载缴款明细数据到控件
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim lngRow As Long
    Dim objCard As Card
    
    On Error GoTo ErrHandler
    If lng卡ID = 0 Then Exit Function
    
    '升级以前的数据不允许余额退款
    strSQL = _
        "Select a.结算方式, a.卡类别id, a.卡号, a.交易流水号, a.交易说明," & vbNewLine & _
        "       Sum(a.余额) As 余额, a.扣率, Sum(a.实际缴款) As 退款金额," & vbNewLine & _
        "       f_List2str(Cast(Collect(To_Char(交易序号)) As t_Strlist)) As 交易序号" & vbNewLine & _
        "From 帐户缴款余额 A" & vbNewLine & _
        "Where a.性质 = 1 And Nvl(a.有效期, Sysdate) >= Sysdate And a.消费卡id = [1] And a.交易序号 > 0" & vbNewLine & _
        "Group By a.结算方式, a.卡类别id, a.卡号, a.交易流水号, a.交易说明, a.扣率" & vbNewLine & _
        "Order By a.卡类别id, 结算方式"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng卡ID)

    If rsTemp.EOF Then
        ShowMsgbox "当前卡无可退余额！"
        Exit Function
    End If
    
    With vsfBlance
        .Redraw = flexRDNone
        .Clear 1
        .Rows = rsTemp.RecordCount + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("结算方式")) = Nvl(rsTemp!结算方式)
            .Cell(flexcpData, lngRow, .ColIndex("结算方式")) = Nvl(rsTemp!交易序号)
            .TextMatrix(lngRow, .ColIndex("余额")) = Format(Nvl(rsTemp!余额), "0.00")
            .TextMatrix(lngRow, .ColIndex("扣率")) = Format(Nvl(rsTemp!扣率), "0.00") & "%"
            .TextMatrix(lngRow, .ColIndex("退款金额")) = Format(Val(Nvl(rsTemp!退款金额)), "0.00")
            .Cell(flexcpData, lngRow, .ColIndex("退款金额")) = Val(Nvl(rsTemp!退款金额))
            .TextMatrix(lngRow, .ColIndex("卡号")) = Nvl(rsTemp!卡号)
            .Cell(flexcpData, lngRow, .ColIndex("卡号")) = Val(Nvl(rsTemp!卡类别id))
            .TextMatrix(lngRow, .ColIndex("交易流水号")) = Nvl(rsTemp!交易流水号)
            .TextMatrix(lngRow, .ColIndex("交易说明")) = Nvl(rsTemp!交易说明)
            If Val(Nvl(rsTemp!卡类别id)) > 0 Then
                Set objCard = GetCurCard(Val(Nvl(rsTemp!卡类别id)))
                If objCard.是否退现 And objCard.是否缺省退现 Then
                    .Cell(flexcpChecked, lngRow, .ColIndex("退现")) = 1
                Else
                    .Cell(flexcpChecked, lngRow, .ColIndex("退现")) = 2
                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = vbBlue
                End If
            End If
            
            rsTemp.MoveNext
            lngRow = lngRow + 1
        Loop
        .Redraw = flexRDBuffered
    End With
    
    Call Calc退款金额(True)

    LoadCardData = True
    Exit Function
ErrHandler:
    vsfBlance.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("':：;；?？" & Chr(22), Chr(KeyAscii)) > 0 Then 'Chr(22):Ctrl+V
        KeyAscii = 0: Exit Sub
    End If
End Sub

Private Sub Form_Load()
    mblnFirst = True
    
    Call ClearData
    If InitData() = False Then Unload Me: Exit Sub
    If Load支付方式() = False Then Unload Me: Exit Sub
    
    If mlng卡ID <> 0 Then
        If CardIsValid(0, mlng卡ID) Then
            If LoadCardData(mlng卡ID) = False Then
                mlng卡ID = 0
                txt(txt_卡号).Text = ""
            End If
        Else
            mlng卡ID = 0
        End If
    End If
    
    pic(pic_余额明细).AutoRedraw = True: zlControl.PicShowFlat pic(pic_余额明细)
    pic(pic_缴款合计).AutoRedraw = True: zlControl.PicShowFlat pic(pic_缴款合计)
    pic(pic_缴款信息).AutoRedraw = True: zlControl.PicShowFlat pic(pic_缴款信息)
    cbo.SetListWidth cbo支付方式, cbo支付方式.Width * 2
    
    txt(txt_找补).BackColor = Me.BackColor
    Me.Caption = "余额退款 - " & mCardType.str卡名称
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    On Error Resume Next
    '焦点定位
    zlControl.ControlSetFocus txt(txt_卡号)
End Sub

Private Function InitData() As Boolean
    '初始化模块变量
    Dim rsTemp As New ADODB.Recordset
    Dim ty_Temp As Ty_CardType
    Dim strValue As String
    
    On Error GoTo ErrHandler
    Set rsTemp = zlGet消费卡接口()
    rsTemp.Filter = "编号=" & mlng卡类别
    If rsTemp.EOF Then
        ShowMsgbox "未发现卡类别信息，不能继续！"
        Exit Function
    End If
    
    '消费卡分币处理方式
    strValue = zlDatabase.GetPara(14, glngSys, , 0)
    mBytMoney = Val(IIf(Len(strValue) = 1, strValue, Mid(strValue, 4, 1)))
    
    mCardType = ty_Temp '自定义Type初始化
    With mCardType
        .str卡名称 = Nvl(rsTemp!名称)
        .str卡号前缀 = Nvl(rsTemp!前缀文本)
        .lng卡号长度 = Val(Nvl(rsTemp!卡号长度))
        .bln卡号密文 = Val(Nvl(rsTemp!是否密文)) = 1
        .bln严格控制 = Val(Nvl(rsTemp!是否严格控制)) = 1
        .str限制类别 = Nvl(rsTemp!限制类别)
        .bln特定病人 = Val(Nvl(rsTemp!是否特定病人)) = 1
    End With
    
    If Init支付方式() = False Then Exit Function
    
    InitData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txt_Change(Index As Integer)
    If mblnNotClick Then Exit Sub
    
    On Error GoTo ErrHandler
    Select Case Index
    Case txt_金额
        Call SetLblCaption
    End Select
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index <> txt_卡号 Then
        If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
    Case txt_卡号
        Call zlControl.TxtCheckKeyPress(txt(Index), KeyAscii, m文本式)
        If InStr(1, "'~～|`-'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
        Call BrushCard(txt(Index), KeyAscii)
    Case txt_金额
        Call zlControl.TxtCheckKeyPress(txt(Index), KeyAscii, m金额式)
    Case Else
        Call zlControl.TxtCheckKeyPress(txt(Index), KeyAscii, m文本式)
    End Select
End Sub

Private Sub BrushCard(ByVal objEdit As Object, KeyAscii As Integer)
    '刷卡
    Static sngBegin As Single
    Dim sngNow As Single
    Dim blnCard As Boolean
    Dim lng卡张数 As Long
    
    On Error GoTo ErrHandler
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    blnCard = zlCommFun.InputIsCard(objEdit, KeyAscii, False)
    If blnCard And Len(objEdit.Text) = mCardType.lng卡号长度 - 1 And KeyAscii <> 8 _
        Or KeyAscii = 13 And Trim(objEdit.Text) <> "" Then '达到卡号长度或回车查找卡信息
        
        If KeyAscii <> 13 Then
            objEdit.Text = objEdit.Text & Chr(KeyAscii)
            objEdit.SelStart = Len(objEdit.Text)
        End If
        KeyAscii = 0
        
        If CardIsValid(1, , objEdit.Text) Then
            If LoadCardData(mlng卡ID) = False Then
                mlng卡ID = 0: Call ClearData
                zlControl.TxtSelAll objEdit: Exit Sub
            End If
        Else
            mlng卡ID = 0: Call ClearData
            zlControl.TxtSelAll objEdit: Exit Sub
        End If
        zlCommFun.PressKey vbKeyTab
    ElseIf KeyAscii = 13 And Trim(objEdit.Text) = "" Then
        zlCommFun.PressKey vbKeyTab
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        '安全刷卡检测
        If KeyAscii <> 0 And KeyAscii > 32 Then
            sngNow = timer
            If objEdit.Text = "" Then
                sngBegin = sngNow
            ElseIf Format((sngNow - sngBegin) / (Len(objEdit) + 1), "0.000") >= 0.04 Then '>0.007>=0.01
                objEdit.Text = Chr(KeyAscii)
                objEdit.SelStart = 1
                KeyAscii = 0
                sngBegin = sngNow
            End If
        End If
    End If
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Function Init支付方式() As Boolean
    '初始化支付方式
    '说明：
    '   只加入现金、支票和三方卡的结算方式
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim i As Long, objCards As Cards, objCard As Card
    Dim lngKey As Long
    
    On Error GoTo ErrHandler
    Set mobjPayCards = New Cards
    
    Set rsTemp = Get结算方式("消费卡")
    If Not gobjSquare Is Nothing Then
        ' zlGetCards(ByVal BytType As Byte)
            '入参:bytType-  0-所有医疗卡;
        '                        1-启用的医疗卡,
        '                        2-所有存在三方账户的三方卡
        '                        3-启用的三方账户的医疗卡
       Set objCards = gobjSquare.objSquareCard.zlGetCards(0)
    End If
    
    With rsTemp
        .Filter = 0
        If .RecordCount <> 0 Then .MoveFirst
        lngKey = 1
        Do While Not .EOF
            For i = 1 To objCards.count
                If objCards(i).结算方式 = Nvl(rsTemp!名称) Then blnFind = True: Exit For
            Next
            If Not blnFind Then
                If (Val(Nvl(rsTemp!性质)) = 1 Or Val(Nvl(rsTemp!性质)) = 2) _
                    And Val(Nvl(rsTemp!应付款)) = 0 Then
                    Set objCard = New Card
                    objCard.短名 = Mid(Nvl(!名称), 1, 1)
                    objCard.接口编码 = Nvl(!编码)
                    objCard.接口程序名 = ""
                    objCard.接口序号 = -1 * lngKey
                    objCard.结算方式 = Nvl(!名称)
                    objCard.名称 = Nvl(!名称)
                    objCard.启用 = True
                    objCard.缺省标志 = Val(Nvl(rsTemp!缺省)) = 1
                    objCard.启用 = True
                    objCard.结算性质 = Val(!性质)
                    
                    mobjPayCards.Add objCard, "K" & lngKey
                    lngKey = lngKey + 1
              End If
            End If
            .MoveNext
        Loop
    End With
    
    '加三方卡
    For i = 1 To objCards.count
        rsTemp.Filter = "名称='" & objCards(i).结算方式 & "'"
        If Not rsTemp.EOF And objCards(i).启用 And Not objCards(i).消费卡 Then
            mobjPayCards.Add objCards(i), "K" & lngKey
            lngKey = lngKey + 1
        End If
    Next
    
    If mobjPayCards.count = 0 Then
        ShowMsgbox "消费卡场合没有可用的结算方式，请先到【结算方式管理】中设置。"
        Exit Function
    End If
    Init支付方式 = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Load支付方式() As Boolean
    '加载支付方式
    '说明:
    '   缺省结算方式的规则，优先顺序如下：
    '   1.结算方式应用中设置的缺省项
    '   2.性质为"1-现金结算方式"的结算方式
    Dim objCard As Card, i As Long
    Dim str结算方式 As String
    
    On Error GoTo ErrHandler
    mlngPre支付方式 = 0

    mblnNotClick = True
    With cbo支付方式
        .Clear
        For i = 1 To mobjPayCards.count
            Set objCard = mobjPayCards(i)
            If objCard.启用 And Not objCard.消费卡 And InStr(str结算方式 & "|", "|" & objCard.结算方式 & "|") = 0 Then
                '三方账户的支付方式显示为医疗卡名称，其它显示结算方式
                If objCard.接口序号 > 0 Then
                    If objCard.是否转帐及代扣 Then
                        .AddItem objCard.名称
                        .ItemData(.NewIndex) = i
                    End If
                Else
                    .AddItem objCard.结算方式
                    .ItemData(.NewIndex) = i
                End If
                
                str结算方式 = str结算方式 & "|" & objCard.结算方式
            End If
            
            '设置缺省值
            If objCard.缺省标志 And .ListIndex < 0 Then .ListIndex = .NewIndex
            If objCard.结算性质 = 1 And .ListIndex < 0 Then .ListIndex = .NewIndex
        Next
            
        If .ListCount > 0 And .ListIndex < 0 Then .ListIndex = 0
    End With
    mblnNotClick = False
    Call cbo支付方式_Click
    Load支付方式 = True
    Exit Function
ErrHandler:
    mblnNotClick = False
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetCurCard(Optional ByVal lng医疗卡ID As Long) As Card
    '获取当前支付卡卡对象，或通过卡类别ID获取卡对象
    Dim intIndex As Integer
    Dim objCard As Card
    
    On Error GoTo ErrHandler
    If lng医疗卡ID = 0 Then
        If cbo支付方式.ListIndex <> -1 Then
            intIndex = cbo支付方式.ItemData(cbo支付方式.ListIndex)
            If intIndex <= 0 Then Exit Function
            Set objCard = mobjPayCards(intIndex)
        End If
    Else
        If Not mobjPayCards Is Nothing Then
            For Each objCard In mobjPayCards
                If objCard.接口序号 = lng医疗卡ID Then Exit For
            Next
        End If
    End If
    If objCard Is Nothing Then Set objCard = New Card
    Set GetCurCard = objCard
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Check缴款情况() As Boolean
    '功能:检查缴款情况
    Dim objCard As Card
    Dim strTitle As String
    
    On Error GoTo ErrHandler
    If mCurCardPay.dbl退款合计 <= 0 Then
        ShowMsgbox "当前卡无可退余额，不能进行余额退款！"
        Exit Function
    End If
    
    If cbo支付方式.ListIndex = -1 Then
        ShowMsgbox "当前退款方式未选择，请检查！"
        zlControl.ControlSetFocus cbo支付方式
        Exit Function
    End If
    
    If zlDblIsValid(Trim(txt(txt_金额).Text), 16, True, False, txt(txt_金额).hWnd, strTitle) = False Then Exit Function
    
    Set objCard = GetCurCard()
    If objCard.结算性质 <> 1 Then
        If RoundEx(mCurCardPay.dbl当前未退 - mCurCardPay.dbl本次误差, 6) = 0 Then
            ShowMsgbox "当前未退款金额为零，不能使用非现金结算方式！"
            zlControl.ControlSetFocus cbo支付方式
            Exit Function
        End If
        
        If Val(txt(txt_金额).Text) = 0 Then
            ShowMsgbox "未输入退款金额，请检查！"
            zlControl.ControlSetFocus txt(txt_金额)
            Exit Function
        End If
    End If
    If Val(txt(txt_金额).Text) <> 0 Then
        If Val(txt(txt_金额).Text) < RoundEx(mCurCardPay.dbl当前未退 - mCurCardPay.dbl本次误差, 6) Then
            ShowMsgbox "退款金额(" & Format(Val(txt(txt_金额).Text), "0.00") & ")不足本次未退金额(" & _
                Format(RoundEx(mCurCardPay.dbl当前未退 - mCurCardPay.dbl本次误差, 6), "0.00") & ")，请检查！"
            zlControl.ControlSetFocus txt(txt_金额)
            Exit Function
        End If
        
        If objCard.结算性质 <> 1 And Val(txt(txt_金额).Text) > RoundEx(mCurCardPay.dbl当前未退 - mCurCardPay.dbl本次误差, 6) Then
            ShowMsgbox "退款金额(" & Format(Val(txt(txt_金额).Text), "0.00") & ")大于了本次未退金额(" & _
                Format(RoundEx(mCurCardPay.dbl当前未退 - mCurCardPay.dbl本次误差, 6), "0.00") & ")，请检查！"
            zlControl.ControlSetFocus txt(txt_金额)
            Exit Function
        End If
    End If
    
    If zlCommFun.StrIsValid(Trim(txt(txt_开户行).Text), 50, txt(txt_开户行).hWnd, "开户行") = False Then Exit Function
    If zlCommFun.StrIsValid(Trim(txt(txt_帐号).Text), 20, txt(txt_帐号).hWnd, "帐号") = False Then Exit Function
    If zlCommFun.StrIsValid(Trim(txt(txt_结算号码).Text), 30, txt(txt_结算号码).hWnd, "结算号码") = False Then Exit Function
    Check缴款情况 = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckThreeSwapIsValied(ByVal objCard As Card, ByVal dblDelMoney As Double, _
    ByVal str交易序号 As String) As Boolean
    '功能:三方卡刷卡验证
    '入参:objCard-当前卡
    '    str交易序号 - 多个用逗号分隔，用于获取原结算序号
    '返回:刷卡成功,返回true,否则返回False
    Dim strXMLExpend As String, strBalanceIDs As String
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    If objCard.接口序号 <= 0 Then CheckThreeSwapIsValied = True: Exit Function
    If dblDelMoney = 0 Then CheckThreeSwapIsValied = True: Exit Function
    
    If objCard.启用 = False Then
        ShowMsgbox objCard.名称 & "未启用，因此不能退回，你可以选择退现！"
        Exit Function
    End If
    
    '获取原结算序号，以及全退检查
    mCurCardPay.str原结算序号 = ""
    strSQL = _
        "Select /*+cardinality(j,10)*/Nvl(Sum(a.实收金额), 0) As 缴款合计," & vbNewLine & _
        "       f_List2str(Cast(Collect(Distinct To_Char(b.结算序号)) As t_Strlist)) As 结算序号" & vbNewLine & _
        "From 病人卡结算记录 A, 病人卡结算记录 B, Table(f_Num2list([1])) J" & vbNewLine & _
        "Where a.结算序号 = b.结算序号 And b.交易序号 = j.Column_Value And a.记录性质 In (1, 2)"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str交易序号)
    If rsTemp.EOF = False Then
        mCurCardPay.str原结算序号 = Nvl(rsTemp!结算序号)
        If objCard.是否转帐及代扣 = False And objCard.是否全退 Then
            If Val(Nvl(rsTemp!缴款合计)) <> dblDelMoney Then
                ShowMsgbox objCard.名称 & "不支持部分退，因此不能退回，你可以选择退现！" & _
                    "(原支付金额：" & FormatEx(Val(Nvl(rsTemp!缴款合计)), 2) & _
                    "，现退款金额：" & FormatEx(dblDelMoney, 2) & ")"
                Exit Function
            End If
        End If
    End If
    
    If objCard.是否转帐及代扣 Then
        '   zlBrushCard(frmMain As Object, _
        ByVal lngModule As Long, _
        ByVal rsClassMoney As ADODB.Recordset, _
        ByVal lngCardTypeID As Long, _
        ByVal bln消费卡 As Boolean, _
        ByVal strPatiName As String, ByVal strSex As String, _
        ByVal strOld As String, ByRef dbl金额 As Double, _
        Optional ByRef strCardNo As String, _
        Optional ByRef strPassWord As String, _
        Optional ByRef bln退费 As Boolean = False, _
        Optional ByRef blnShowPatiInfor As Boolean = False, _
        Optional ByRef bln退现 As Boolean = False, _
        Optional ByVal bln余额不足禁止 As Boolean = True, _
        Optional ByRef varSquareBalance As Variant, _
        Optional ByVal bln转预交 As Boolean = False, _
        Optional ByVal blnAllPay As Boolean = False, _
        Optional ByVal strXmlIn As String = "") As Boolean
        '       strXmlIn-三方卡调用XML入参,目前格式如下:
        '       <IN>
        '           <CZLX>0</CZLX>    //操作类型,0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡
        '       </IN>
        If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, Nothing, _
            objCard.接口序号, False, "", "", "", -1 * dblDelMoney, mCurCardPay.str刷卡卡号, mCurCardPay.str刷卡密码, _
            False, False, False, True, Nothing, False, True, "<IN><CZLX>1</CZLX></IN>") = False Then Exit Function
        
        '调用转帐接口
        'zlTransferAccountsCheck 转帐检查接口
        '参数名  参数类型    入/出   备注
        'frmMain Object  In  调用的主窗体
        'lngModule   Long    In  HIS调用模块号
        'lngCardTypeID   Long    In  卡类别ID
        'strCardNo   String  In  卡号
        'dblDelMoney    Double  In  转帐金额(代扣时为负数)
        'strBalanceID    String  In  原支付结算序号,费用补充记录.结算序号或病人预交记录.结算序号
        'strXMLExpend String In   XML串:
        '                            <IN>
        '                                <CZLX>操作类型</CZLX> //0或NULL:补结算业务;1-补结算退费业务；
        '                                       2-结帐业务;3-结帐退费业务；4-门诊退费业务；5-消费卡管理退费业务
        '                            </IN>
        '                    Out  XML串:
        '                            <OUT>
        '                               <ERRMSG>错误信息</ERRMSG >
        '                            </OUT>
        '    Boolean 函数返回    检查的数据合法,返回True:否则返回False
        '说明:
        '１. 在三方转帐时的一些合法性检查，避免在转帐时弹出对话框之类的等待造成死锁或其它现象的发生。
        '２. 不存在检测的需要返回为True，否则不能完成转帐功能的调用。
        '构造XML串
        strXMLExpend = "<IN><CZLX>5</CZLX></IN>"
        If gobjSquare.objSquareCard.zlTransferAccountsCheck(Me, mlngModule, objCard.接口序号, _
            mCurCardPay.str刷卡卡号, dblDelMoney, mCurCardPay.str原结算序号, strXMLExpend) = False Then
            Call ShowThreeSwapErrMsg(0, strXMLExpend)
            Exit Function
        End If
    Else
        'zlReturnCheck(frmMain As Object, ByVal lngModule As Long, _
            ByVal lngCardTypeID As Long, bln消费卡 As Boolean, ByVal strCardNo As String, _
            ByVal strBalanceIDs As String, _
            ByVal dblDelMoney As Double, ByVal strSwapNo As String, _
            ByVal strSwapMemo As String, ByRef strXMLExpend As String) As Boolean
            '---------------------------------------------------------------------------------------------------------------------------------------------
            '功能:帐户回退交易前的检查
            '入参:frmMain-调用的主窗体
            '       lngModule-调用的模块号
            '       lngCardTypeID-卡类别ID
            '       strCardNo-卡号
            '       strBalanceIDs   String  In  本次支付所涉及的结算ID 格式:收费类型|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
            '                                   收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款,6-保险补充结算,7-消费卡收款；
            '                                           类型=7时，ID为病人卡结算记录.结算序号
            '       dblDelMoney-退款金额
            '       strSwapNo-交易流水号(退款时检查)
            '       strSwapMemo-交易说明(退款时传入)
            '       strXMLExpend    XML IN  可选参数(扩展用):
            '        <TFDATA> //退费数据
            '          <YCTF>1</YCTF> //是否异常重退:1-异常重退;0-退费 此节点可能没有
            '          <TFLIST> //退费列表
            '            <NO></NO> // 退费单据
            '            <TFITEM> //退费项
            '              <SerialNum></SerialNum> //序号
            '              …
            '            </TFITEM>
            '          </TFLIST>
            '          ....
            '        </TFDATA >
            '返回:退款合法,返回true,否则返回Flase
        strBalanceIDs = "7|" & mCurCardPay.str原结算序号
        If gobjSquare.objSquareCard.zlReturncheck(Me, mlngModule, objCard.接口序号, _
            objCard.消费卡, mCurCardPay.str刷卡卡号, strBalanceIDs, dblDelMoney, _
            mCurCardPay.str交易流水号, mCurCardPay.str交易说明, strXMLExpend) = False Then Exit Function
    
        If objCard.是否退款验卡 Then
           '弹出刷卡界面
            'zlBrushCard(frmMain As Object, _
            ByVal lngModule As Long, _
            ByVal rsClassMoney As ADODB.Recordset, _
            ByVal lngCardTypeID As Long, _
            ByVal bln消费卡 As Boolean, _
            ByVal strPatiName As String, ByVal strSex As String, _
            ByVal strOld As String, ByVal dbl金额 As Double, _
            Optional ByRef strCardNo As String, _
            Optional ByRef strPassWord As String, _
            Optional ByRef bln退费 As Boolean = False, _
            Optional ByRef blnShowPatiInfor As Boolean = False, _
            Optional ByRef bln退现 As Boolean = False, _
            Optional ByVal bln余额不足禁止 As Boolean = True, _
            Optional ByRef varSquareBalance As Variant, _
            Optional ByVal bln转预交 As Boolean = False, _
            Optional ByVal blnAllPay As Boolean = False, _
            Optional ByVal strXmlIn As String = "") As Boolean
            '       strXmlIn-三方卡调用XML入参,目前格式如下:
            '       <IN>
            '           <CZLX>0</CZLX>    //操作类型,0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡
            '       </IN>
            If gobjSquare.objSquareCard.zlBrushCard(Me, mlngModule, Nothing, _
                objCard.接口序号, False, "", "", "", dblDelMoney, mCurCardPay.str刷卡卡号, mCurCardPay.str刷卡密码, _
                True, False, False, True, Nothing, False, True, "<IN><CZLX>2</CZLX></IN>") = False Then Exit Function
        End If
    End If
    CheckThreeSwapIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteThreeSwapPay(ByVal objCard As Card, ByVal lng结算序号 As Long, _
    ByVal dblDelMoney As Double) As Boolean
    '功能:一卡通支付(三方接口)
    '入参:
    '   objCard-当前卡
    '   dblDelMoney-本次支付金额
    '出参:
    '返回:执行成功,返回true,否则返回False
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim strSwapExtendInfor As String, strTemp As String
    Dim strXMLExpend As String
    
    On Error GoTo ErrHandler
    If objCard.接口序号 <= 0 Then ExecuteThreeSwapPay = True: Exit Function
    If dblDelMoney = 0 Then ExecuteThreeSwapPay = True: Exit Function
    
    Set cllUpdate = New Collection
    Set cllThreeSwap = New Collection

    If objCard.是否转帐及代扣 Then
        'zlTransferAccountsMoney
        '参数名  参数类型    入/出   备注
        'frmMain Object  In  调用的主窗体
        'lngModule   Long    In  HIS调用模块号
        'lngCardTypeID   Long    In  卡类别ID
        'strCardNo   String  In  卡号
        'strBalanceID    String  In  结算ID 本次支付结算序号,费用补充记录.结算序号或病人预交记录.结算序号或病人卡结算记录.结算序号
        'dblDelMoney    Double  In  转帐金额
        'strSwapGlideNO  String  Out 交易流水号
        'strSwapMemo String  Out 交易说明
        'strSwapExtendInfor  String  In 退费业务时，传入本次退费的冲销ID:
        '                               格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
        '                               收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款,6-保险补充结算,7-消费卡管理收款(ID为结算序号)
        '                           Out 交易扩展信息: 格式为:项目名称1|项目内容2||…||项目名称n|项目内容n
        'strXMLExpend String In   XML串:
        '                            <IN>
        '                                <CZLX>操作类型</CZLX> //0或NULL:补结算业务;1-补结算退费业务；
        '                                       2-结帐业务;3-结帐退费业务；4-门诊退费业务；5-消费卡管理退费业务
        '                            </IN>
        '                    Out  XML串:
        '                            <OUT>
        '                               <ERRMSG>错误信息</ERRMSG >
        '                            </OUT>
        '    Boolean 函数返回    True:调用成功,False:调用失败
        '说明:
        '１. 在医保补充结算时进行的三方转帐时调用。
        '２. 一般来说，成功转帐后，都应该打印相关的结算票据，可以放在此接口进行处理.
        '３. 在转帐成功后，返回交易流水号和相关交易说明；如果存在其他交易信息，可以放在扩展信息中返回.
        '构造XML串
        strXMLExpend = "<IN><CZLX>5</CZLX></IN>"
        strSwapExtendInfor = "7|" & mCurCardPay.str原结算序号: strTemp = strSwapExtendInfor
        If gobjSquare.objSquareCard.zlTransferAccountsMoney(Me, mlngModule, objCard.接口序号, _
            mCurCardPay.str刷卡卡号, lng结算序号, dblDelMoney, _
            mCurCardPay.str交易流水号, mCurCardPay.str交易说明, strSwapExtendInfor, strXMLExpend) = False Then
            gcnOracle.RollbackTrans: Call ShowThreeSwapErrMsg(1, strXMLExpend)
            Exit Function
        End If
        gcnOracle.CommitTrans: ExecuteThreeSwapPay = True
        
        Call zlAddUpdateSwapSQL(False, lng结算序号, objCard.接口序号, objCard.消费卡, _
            mCurCardPay.str刷卡卡号, mCurCardPay.str交易流水号, mCurCardPay.str交易说明, cllUpdate, 1, 0, 1)
        zlExecuteProcedureArrAy cllUpdate, Me.Caption
        If strTemp <> strSwapExtendInfor Then
            Call zlAddThreeSwapSQLToCollection(False, lng结算序号, objCard.接口序号, objCard.消费卡, _
                mCurCardPay.str刷卡卡号, strSwapExtendInfor, cllThreeSwap, 0, 1)
            zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
        End If
    Else
        'zlReturnMoney(frmMain As Object, ByVal lngModule As Long, _
            ByVal lngCardTypeID As Long, ByVal strCardNo As String, ByVal strBalanceIDs As String, _
            ByVal dblDelMoney As Double, _
            ByRef strSwapGlideNO As String, ByRef strSwapMemo As String, _
            ByRef strSwapExtendInfor As String) As Boolean
        '---------------------------------------------------------------------------------------------------------------------------------------------
        '功能:帐户扣款回退交易
        '入参:frmMain-调用的主窗体
        '       lngModule-调用的模块号
        '       lngCardTypeID-卡类别ID:医疗卡类别.ID
        '       strCardNo-卡号
        '       strBalanceIDs-本次支付所涉及的结算ID(这是原结帐ID):
        '                           格式:收费类型(|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
        '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款,6-保险补充结算,7-消费卡收款
        '       dblDelMoney-退款金额
        '       strSwapNo-交易流水号(扣款时的交易流水号)
        '       strSwapMemo-交易说明(扣款时的交易说明)
        '       strSwapExtendInfor-出入，本次退费的冲销ID：
        '                           格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
        '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款,6-保险补充结算,7-消费卡收款
        '       strSwapExtendInfor-传出，交易的扩展信息
        '           格式为:项目名称1|项目内容2||…||项目名称n|项目内容n 每个项目中不能包含|字符
        strSwapExtendInfor = "7|" & lng结算序号
        If gobjSquare.objSquareCard.zlReturnMoney(Me, mlngModule, objCard.接口序号, objCard.消费卡, _
            mCurCardPay.str刷卡卡号, "7|" & mCurCardPay.str原结算序号, dblDelMoney, _
            mCurCardPay.str交易流水号, mCurCardPay.str交易说明, strSwapExtendInfor) = False Then
            gcnOracle.RollbackTrans: Exit Function
        End If
        gcnOracle.CommitTrans: ExecuteThreeSwapPay = True
        
        Call zlAddUpdateSwapSQL(False, lng结算序号, objCard.接口序号, objCard.消费卡, _
            mCurCardPay.str刷卡卡号, mCurCardPay.str交易流水号, mCurCardPay.str交易说明, cllUpdate, 1, 0, 1)
        zlExecuteProcedureArrAy cllUpdate, Me.Caption
        If strTemp <> strSwapExtendInfor Then
            Call zlAddThreeSwapSQLToCollection(False, lng结算序号, objCard.接口序号, objCard.消费卡, _
                mCurCardPay.str刷卡卡号, strSwapExtendInfor, cllThreeSwap, 0, 1)
            zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
        End If
    End If
    ExecuteThreeSwapPay = True
    Exit Function
ErrHandler:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowThreeSwapErrMsg(ByVal bytType As Byte, ByVal strXMLErrMsg As String)
    '功能:三方转账检查与代扣业务出错提示
    '参数:
    '   bytType:0-转账检查,1-转账交易
    '   strXMLErrMsg:格式如下
    '            <OUT>
    '               <ERRMSG>错误信息</ERRMSG >
    '            </OUT>
    Dim strValue As String
    
    On Error GoTo errHandle
    '解析错误信息
    If strXMLErrMsg <> "" Then
        If zlXML.OpenXMLDocument(strXMLErrMsg) = False Then strValue = ""
        If zlXML.GetSingleNodeValue("OUT/ERRMSG", strValue) = False Then strValue = ""
        Call zlXML.CloseXMLDocument
    End If
    '提示错误信息
    If Trim(strValue) = "" Then
        If bytType = 0 Then
            strValue = vbCrLf & "交易检查失败！"
        Else
            strValue = vbCrLf & "交易失败！"
        End If
    End If
    MsgBox strValue, vbExclamation, gstrSysName
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function CheckThreeBalanceToCash(ByVal objCard As Card) As Boolean
    '三方卡退现检查
    Dim str操作员 As String
    
    On Error GoTo errHandle
    If Not (objCard.接口序号 > 0 And Not objCard.消费卡) Then CheckThreeBalanceToCash = True: Exit Function
    If objCard.是否退现 Then CheckThreeBalanceToCash = True: Exit Function
    
    If zlstr.IsHavePrivs(GetPrivFunc(glngSys, 1151), "三方退款强制退现") Then
        If MsgBox(objCard.名称 & "不支持退现，你确定要将其强制退现吗？", _
            vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    Else
        str操作员 = zlDatabase.UserIdentifyByUser(Me, objCard.名称 & "强制退现，权限验证：", _
            glngSys, mlngModule, "三方退款强制退现", , True)
        If str操作员 = "" Then Exit Function
    End If
    CheckThreeBalanceToCash = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub vsfBlance_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsfBlance
        If Col = .ColIndex("退现") Then
            If Val(.Cell(flexcpChecked, Row, .ColIndex("退现"))) = 1 Then '退现
                .Cell(flexcpForeColor, Row, 0, Row, .Cols - 1) = .ForeColor
                .ForeColorSel = .ForeColor
            Else
                .Cell(flexcpForeColor, Row, 0, Row, .Cols - 1) = vbBlue
                .ForeColorSel = vbBlue
            End If
        
            Call Calc退款金额
        End If
    End With
End Sub

Private Sub vsfBlance_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    On Error Resume Next
    vsfBlance.ForeColorSel = vsfBlance.Cell(flexcpForeColor, NewRow, 0)
End Sub

Private Sub vsfBlance_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lng医疗卡ID As Long
    With vsfBlance
        If Col <> .ColIndex("退现") Then Cancel = True: Exit Sub
        lng医疗卡ID = Val(.Cell(flexcpData, Row, .ColIndex("卡号")))
        If lng医疗卡ID <= 0 Then Cancel = True: Exit Sub
    End With
End Sub

Private Sub vsfBlance_GotFocus()
    If vsfBlance.Row < vsfBlance.FixedRows And vsfBlance.Rows > 1 Then
        vsfBlance.Row = 1
    End If
End Sub

Private Sub vsfBlance_KeyDown(KeyCode As Integer, Shift As Integer)
    If vsfBlance.Col <> vsfBlance.ColIndex("退现") Then
        If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub vsfBlance_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lng医疗卡ID As Long, objCard As Card
    
    On Error GoTo errHandle
    With vsfBlance
        If Col <> .ColIndex("退现") Then Exit Sub
        lng医疗卡ID = Val(.Cell(flexcpData, Row, .ColIndex("卡号")))
        If lng医疗卡ID <= 0 Then Exit Sub
        
        '退现检查
        If Val(.TextMatrix(Row, .ColIndex("退款金额"))) = 0 Or Abs(Val(.EditText)) <> 1 Then Exit Sub
        Set objCard = GetCurCard(lng医疗卡ID)
        If objCard.是否退现 Then Exit Sub
        If CheckThreeBalanceToCash(objCard) = False Then Cancel = True: Exit Sub
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Calc退款金额(Optional ByVal bln计算合计 As Boolean)
    '计算退款金额
    Dim lngRow As Long
    
    On Error GoTo ErrHandler
    If bln计算合计 Then
        mCurCardPay.dbl退款合计 = 0
        mCurCardPay.dbl当前未退 = 0
        With vsfBlance
            For lngRow = 1 To .Rows - 1
                mCurCardPay.dbl退款合计 = mCurCardPay.dbl退款合计 + Val(.TextMatrix(lngRow, .ColIndex("退款金额")))
            Next
            mCurCardPay.dbl退款合计 = Round(mCurCardPay.dbl退款合计, 6)
        End With
    End If
    
    mCurCardPay.dbl当前未退 = 0
    mCurCardPay.dbl本次误差 = 0
    With vsfBlance
        For lngRow = 1 To .Rows - 1
            If Val(.Cell(flexcpData, lngRow, .ColIndex("卡号"))) > 0 Then
                If Val(.Cell(flexcpChecked, lngRow, .ColIndex("退现"))) = 1 Then
                    mCurCardPay.dbl当前未退 = mCurCardPay.dbl当前未退 + Val(.Cell(flexcpData, lngRow, .ColIndex("退款金额")))
                End If
            Else
                mCurCardPay.dbl当前未退 = mCurCardPay.dbl当前未退 + Val(.Cell(flexcpData, lngRow, .ColIndex("退款金额")))
            End If
        Next
        mCurCardPay.dbl当前未退 = Round(mCurCardPay.dbl当前未退, 6)
    End With
    
    lbl(lbl_退款合计).Caption = FormatEx(mCurCardPay.dbl退款合计, 6, , , 2)
    
    Call SetControlProperty
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ClearData()
    '清除数据
    Dim tyPayMoney As TY_PayMoney
    
    On Error GoTo ErrHandler
    mCurCardPay = tyPayMoney
    
    lbl(lbl_余额).Caption = lbl(lbl_余额).Tag & "0.00"
    lblPatiInfo.Caption = ""
    
    vsfBlance.Clear 1
    vsfBlance.Rows = 1
    
    lbl(lbl_退款合计).Caption = "0.00"
    lbl(lbl_当前未退).Caption = "0.00"
    
    txt(txt_金额).Text = ""
    Call SetControlProperty
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
