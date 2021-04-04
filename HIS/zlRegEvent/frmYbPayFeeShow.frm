VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmYbPayFeeShow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "医保病人缴款"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7620
   ControlBox      =   0   'False
   Icon            =   "frmYbPayFeeShow.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   7620
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame3 
      Caption         =   "本次结算信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3330
      Left            =   240
      TabIndex        =   0
      Top             =   1245
      Width           =   4245
      Begin VB.TextBox txt实收合计 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00108000&
         Height          =   450
         Left            =   1005
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "0.00"
         ToolTipText     =   "本次应缴合计=累计实缴金额-累计个人帐户支付-累计冲预交额"
         Top             =   2640
         Width           =   2985
      End
      Begin VSFlex8Ctl.VSFlexGrid vsData 
         Height          =   2130
         Left            =   210
         TabIndex        =   1
         Top             =   420
         Width           =   3825
         _cx             =   6747
         _cy             =   3757
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
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
         GridColor       =   -2147483631
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   350
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmYbPayFeeShow.frx":014A
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
         ExplorerBar     =   7
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "合计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   210
         TabIndex        =   2
         Top             =   2715
         Width           =   660
      End
   End
   Begin VB.Frame fraJK 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   1905
      Left            =   4800
      TabIndex        =   17
      Top             =   1920
      Width           =   2595
      Begin VB.TextBox txt本次应缴 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00108000&
         Height          =   450
         Left            =   840
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "0.00"
         ToolTipText     =   "本次应缴合计=累计实缴金额-累计个人帐户支付-累计冲预交额"
         Top             =   15
         Width           =   1575
      End
      Begin VB.TextBox txt找补 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   1230
         Width           =   1575
      End
      Begin VB.TextBox txt缴款 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   840
         MaxLength       =   10
         TabIndex        =   7
         Text            =   "0.00"
         Top             =   615
         Width           =   1575
      End
      Begin VB.Label lbl应缴 
         AutoSize        =   -1  'True
         Caption         =   "应缴"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   0
         TabIndex        =   4
         Top             =   90
         Width           =   660
      End
      Begin VB.Label lbl找补 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "找补"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   19
         Top             =   1305
         Width           =   690
      End
      Begin VB.Label lbl缴款 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "缴款"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   0
         TabIndex        =   6
         Top             =   690
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5685
      TabIndex        =   8
      ToolTipText     =   "热键:F2"
      Top             =   4005
      Width           =   1350
   End
   Begin VB.PictureBox picTotal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   -765
      ScaleHeight     =   930
      ScaleWidth      =   4620
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   5685
      Width           =   4620
      Begin VB.Label lbl合计 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   840
         Left            =   1395
         TabIndex        =   15
         Top             =   -15
         Width           =   1410
      End
      Begin VB.Label lblSum 
         BackStyle       =   0  'Transparent
         Caption         =   "合计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   15
         TabIndex        =   14
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   120
      Left            =   -630
      TabIndex        =   12
      Top             =   5235
      Width           =   7995
   End
   Begin VB.Frame Frame1 
      Height          =   120
      Left            =   -45
      TabIndex        =   10
      Top             =   885
      Width           =   7995
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5730
      TabIndex        =   9
      Top             =   5700
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label lblPati 
      AutoSize        =   -1  'True
      Caption         =   "病人姓名"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   210
      Left            =   1005
      TabIndex        =   16
      Top             =   645
      Width           =   900
   End
   Begin VB.Label lblTittle 
      AutoSize        =   -1  'True
      Caption         =   "以下为医保病人应该缴纳的应缴款项,请注意收款!"
      Height          =   180
      Left            =   1035
      TabIndex        =   11
      Top             =   300
      Width           =   3960
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   135
      Picture         =   "frmYbPayFeeShow.frx":01A3
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmYbPayFeeShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng病人ID As Long, mintInsure As Integer
Private mcur实收总额 As Currency, mcur应缴额 As Currency, mcur找补 As Currency, mcur缴款额 As Currency
Private mstr挂号NO As String, mstr就诊卡NO    As String
Private mobjInsure As Object, mblnLED As Boolean
Private mstr姓名 As String, mstr性别 As String, mstr年龄 As String
Private mblnOk  As Boolean
Private mblnFirst As Boolean
Private Function zlget缴款情况() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取本次结算情况
    '编制:刘兴洪
    '日期:2009-12-16 15:17:00
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp  As ADODB.Recordset, lngRow As Long
    Dim strSQL As String
    On Error GoTo Hd
    strSQL = "" & _
    "   Select  A.结算方式,Sum(A.冲预交) As 金额 " & _
    "   From 病人预交记录 A, Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) b" & _
    "   Where A.NO=b.Column_Value and A.记录性质=4 " & _
    "   Group by A.结算方式 "
    
    strSQL = strSQL & " UNION ALL " & _
    "   Select 结算方式,Sum(A.冲预交) As 金额 " & _
    "   From 病人预交记录 A, Table(Cast(f_Str2list([2]) As zlTools.t_Strlist)) b" & _
    "   Where A.NO=b.Column_Value and A.记录性质= 5" & _
    "   Group by A.结算方式"
    
     strSQL = "" & _
     "   Select /*+ rule */  A.结算方式,Sum(nvl(A.金额,0)) As 金额 " & _
    "   From (" & strSQL & ") A" & _
    "   Group by A.结算方式"
    
    mstr就诊卡NO = IIf(mstr就诊卡NO = "", "-()4243_Js2~~~", mstr就诊卡NO)
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr挂号NO, mstr就诊卡NO)
    With vsData
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("结算方式")) = Nvl(rsTemp!结算方式)
            .TextMatrix(lngRow, .ColIndex("金额")) = Format(Val(Nvl(rsTemp!金额)), "0.00")
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    zlget缴款情况 = True
Exit Function
Hd:
    If ErrCenter() = 1 Then Resume
    SaveErrLog
End Function

Public Function zlShowPayWindows(ByVal frmMain As Form, ByVal objInsure As Object, _
    ByVal blnLED As Boolean, ByVal str姓名 As String, ByVal str性别 As String, ByVal str年龄 As String, _
    ByVal lng病人ID As Long, ByVal intInsure As Integer, ByVal str挂号NO As String, ByVal str就诊卡NO As String, _
    cur实收总额 As Currency, cur应缴额 As Currency, cur缴款额 As Currency, cur找补 As Currency) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进入缴款金额界面,目前只有医保病原人有效
    '入参:objInsure-医保对象
    '      str挂号NO-挂号单号
    '      str就诊卡NO-就诊卡NO
    '出参:cur实收总额 , cur应缴额 , cur找补
    '返回:
    '编制:刘兴洪
    '日期:2009-12-02 15:13:53
    '---------------------------------------------------------------------------------------------------------------------------------------------
     Set mobjInsure = objInsure
    mlng病人ID = lng病人ID: mcur实收总额 = cur实收总额: mcur应缴额 = cur应缴额: mcur找补 = cur找补
    mstr姓名 = str姓名: mstr性别 = str性别: mstr年龄 = str年龄
    mstr挂号NO = str挂号NO: mstr就诊卡NO = str就诊卡NO: mcur缴款额 = 0: mcur找补 = 0
    mblnOk = False: mblnLED = blnLED
    mblnLED = mblnLED
    Me.Show 1, frmMain
    zlShowPayWindows = mblnOk
    cur缴款额 = mcur缴款额: cur找补 = mcur找补
End Function

Private Sub cmdCancel_Click()
    mblnOk = False: mcur找补 = 0
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If Val(txt缴款.Text) <> 0 Then
        If Val(txt找补.Text) < 0 Then
            MsgBox "缴款金额不足。", vbInformation, gstrSysName
            Call zlControl.TxtSelAll(txt缴款): Exit Sub
        End If
    End If
    mcur缴款额 = Val(txt缴款.Text)
    mcur找补 = Val(txt找补.Text)
    mblnOk = True
    Unload Me
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    Call zlget缴款情况
    lblPati.Caption = "病人姓名:" & mstr姓名 & Space(4) & "性别:" & mstr性别 & Space(4) & "年龄:" & mstr年龄
    txt本次应缴.Text = Format(mcur应缴额, "0.00")
    lbl合计.Caption = Format(mcur实收总额, "0.00")
    txt实收合计.Text = lbl合计.Caption
    txt缴款.Text = "0.00"
    txt找补.Text = "0.00"
    If txt缴款.Enabled And txt缴款.Visible Then txt缴款.SetFocus
End Sub
Private Sub Form_Load()
    mblnFirst = True
End Sub
Private Sub txt缴款_Change()
    Dim cur应缴 As Currency
    If Val(txt缴款.Text) = 0 Then
        txt找补.Text = "0.00"
    Else
        txt找补.Text = Format(Val(txt缴款.Text) - mcur应缴额, "0.00")
    End If
End Sub
Private Sub txt缴款_GotFocus()
    Call zlControl.TxtSelAll(txt缴款)
    If Val(txt缴款.Text) = 0 And Me.ActiveControl Is txt缴款 Then
        txt缴款.Text = ""
    End If
    '语音提示
    If mblnLED Then
        zl9LedVoice.Speak "#21 " & Format(mcur应缴额, "0.00")
    End If
End Sub
Private Sub txt缴款_KeyPress(KeyAscii As Integer)
    Dim cur应缴 As Currency
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txt缴款.Text = "" Then
            If mcur实收总额 = 0 Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        End If
        If txt缴款.Text = "" Then Exit Sub
        
        If Val(txt缴款.Text) <> 0 Then
            If Val(txt找补.Text) < 0 Then
                MsgBox "缴款金额不足。", vbInformation, gstrSysName
                Call zlControl.TxtSelAll(txt缴款): Exit Sub
            End If
        End If
        
        Call zlCommFun.PressKey(vbKeyTab)
        'LED显示
        If mblnLED And Val(txt找补.Text) >= 0 Then
            zl9LedVoice.DispCharge Format(mcur应缴额, "0.00"), Val(txt缴款.Text), Val(txt找补.Text)
            zl9LedVoice.Speak "#22 " & txt缴款.Text
            zl9LedVoice.Speak "#23 " & txt找补.Text
            zl9LedVoice.Speak "#3"
        End If
    Else
        If KeyAscii = Asc(".") And InStr(txt缴款.Text, ".") > 0 Then KeyAscii = 0:  Exit Sub
        If InStr("0123456789." & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub
Private Sub txt缴款_LostFocus()
    txt缴款.Text = Format(Val(txt缴款.Text), "0.00")
End Sub

Private Sub txt缴款_Validate(Cancel As Boolean)
    txt缴款.Text = Format(Val(txt缴款.Text), "0.00")
End Sub
 
