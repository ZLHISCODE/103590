VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmReviewReason 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "免责声明"
   ClientHeight    =   6630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10365
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VSFlex8Ctl.VSFlexGrid vsReason 
      Height          =   2535
      Left            =   240
      TabIndex        =   7
      Top             =   720
      Width           =   9255
      _cx             =   16325
      _cy             =   4471
      Appearance      =   3
      BorderStyle     =   0
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
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   500
      ColWidthMin     =   2000
      ColWidthMax     =   8000
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
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
   Begin VB.PictureBox picBottom 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   -360
      ScaleHeight     =   735
      ScaleWidth      =   10365
      TabIndex        =   4
      Top             =   5880
      Width           =   10365
      Begin VB.PictureBox picBtn 
         Appearance      =   0  'Flat
         BackColor       =   &H00D48A00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   0
         Left            =   4440
         ScaleHeight     =   420
         ScaleWidth      =   1095
         TabIndex        =   5
         Top             =   120
         Width           =   1100
         Begin VB.Label lblBtn 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "确定"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Left            =   360
            TabIndex        =   6
            Top             =   120
            Width           =   450
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         X1              =   0
         X2              =   9960
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Timer tmrTime 
      Interval        =   50
      Left            =   9960
      Top             =   6240
   End
   Begin VB.PictureBox picTop 
      BackColor       =   &H00D48A00&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   9975
      TabIndex        =   0
      Top             =   120
      Width           =   9975
      Begin VB.PictureBox picClosed 
         Appearance      =   0  'Flat
         BackColor       =   &H00D48A00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   500
         Left            =   9480
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   1
         Top             =   0
         Width           =   500
         Begin VB.Label lblClose 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "×"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   300
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   300
         End
      End
      Begin VB.Label lblFrmName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用药理由"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   900
      End
   End
   Begin VB.Line linScope 
      Index           =   3
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   10560
   End
   Begin VB.Line linScope 
      Index           =   1
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   10560
   End
   Begin VB.Line linScope 
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   10560
   End
   Begin VB.Line linScope 
      Index           =   2
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   10560
   End
End
Attribute VB_Name = "frmReviewReason"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mMoveX As Long, mMoveY As Long  '记录窗体移动前，窗体左上角与鼠标指针位置间的纵横距离
Private mudtRect As RECT
Private mudtRectClose As RECT
Private mudtPoint As POINTAPI
Private mblnMoveStart As Boolean '判断移动是否开始
Private mblnMove As Boolean

Private mrsAdvice As ADODB.Recordset
Private mbytStyle  As Byte          '0-填写拒绝理由;1-显示未审核通过原因
Private Enum E_COL
    COL_医嘱内容 = 0
    COL_审查详情 = 1
    COL_用药理由 = 2
End Enum

Public Function ShowMe(ByRef objfrmMain As Object, ByVal rsRet As ADODB.Recordset, Optional ByRef rsUnReason As ADODB.Recordset, _
    Optional ByVal bytStyle As Byte) As Boolean
'功能:
    Set mrsAdvice = rsRet
    mbytStyle = bytStyle
    Me.Show 1, objfrmMain
    mrsAdvice.Filter = ""
    Set rsUnReason = mrsAdvice
    Set mrsAdvice = Nothing
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then
        KeyAscii = 0
    ElseIf KeyAscii = Asc("`") Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim strTxt As String
    Dim i As Long
    
    picTop.BackColor = conCOLOR_TITLE_BAR
    If mbytStyle = 0 Then
        lblFrmName.Caption = "拒绝药师审核理由"
    Else
        lblFrmName.Caption = "审方结果"
    End If
    Call InitReason
    With vsReason
        .Rows = mrsAdvice.RecordCount + 1
        .WordWrap = True
        mrsAdvice.Filter = ""
        For i = 1 To mrsAdvice.RecordCount
            .Cell(flexcpData, i, COL_医嘱内容) = mrsAdvice!医嘱ID & ""
            .TextMatrix(i, COL_医嘱内容) = mrsAdvice!医嘱内容 & ""
            .TextMatrix(i, COL_审查详情) = mrsAdvice!审查详情 & ""
            .Cell(flexcpForeColor, i, COL_审查详情) = vbRed
            .TextMatrix(i, COL_用药理由) = ""
            mrsAdvice.MoveNext
        Next
        '对齐样式
        .Cell(flexcpAlignment, 1, COL_医嘱内容, .Rows - 1, COL_用药理由) = flexAlignLeftCenter
    End With
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picTop.Move 15, 15, Me.ScaleWidth - 30, 495
    picBottom.Move 15, Me.ScaleHeight - 750, Me.ScaleWidth - 30, 735
    vsReason.Move 15, picTop.Height + picTop.Top, Me.ScaleWidth - 30, Me.ScaleHeight - picTop.Height - picBottom.Height - 30
    picBottom.BackColor = vbWhite
    'Left
    With linScope(0)
        .X1 = 0: .X2 = 0: .Y1 = 0: .Y2 = Me.ScaleHeight
        .BorderColor = conCOLOR_TITLE_BAR
        '&H00808080&
        '&H80000010& '按钮阴影
    End With
    'bottom
    With linScope(1)
        .X1 = 0: .X2 = Me.ScaleWidth: .Y1 = Me.ScaleHeight - 15: .Y2 = Me.ScaleHeight - 15
        .BorderColor = conCOLOR_TITLE_BAR
    End With
    'right
    With linScope(2)
        .X1 = Me.ScaleWidth - 15: .X2 = Me.ScaleWidth - 15: .Y1 = 0: .Y2 = Me.ScaleHeight - 15
        .BorderColor = conCOLOR_TITLE_BAR
    End With
    'Top
    With linScope(3)
        .X1 = 0: .X2 = Me.ScaleWidth - 15: .Y1 = 0: .Y2 = 0
        .BorderColor = conCOLOR_TITLE_BAR
    End With
End Sub

Private Sub lblBtn_Click()
    Dim i As Long
    Dim strRefuse As String
    Dim strOut As String
    If mbytStyle = 0 Then
        With vsReason
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, COL_用药理由)) <> "" Then
                    strRefuse = strRefuse & "," & "{""order_id"":""" & .Cell(flexcpData, i, COL_医嘱内容) & """,""dr_refuse_comment"":""" & Trim(.TextMatrix(i, COL_用药理由)) & """}"
                    mrsAdvice.Filter = "医嘱ID =" & .Cell(flexcpData, i, COL_医嘱内容)
                    If Not mrsAdvice.EOF Then mrsAdvice.Delete
                End If
            Next
            If strRefuse <> "" Then '理由回传服务
                strRefuse = "[" & Mid(strRefuse, 2) & "]"
                Call sys.NewSystemSvr("药师处方审查", "回写医生拒绝理由", strRefuse, strOut)
            End If
        End With
    End If
    Unload Me
End Sub

Private Sub lblClose_Click()
    lblBtn_Click
End Sub

Private Sub picBottom_Resize()
    On Error Resume Next
    picBtn(0).Move picBottom.Width / 2 - picBtn(0).Width / 2, picBottom.Height / 2 - picBtn(0).Height / 2
    With Line1
        .X1 = 120: .Y1 = 0
        .X2 = picBottom.ScaleWidth - 240: .Y2 = 0
        .BorderColor = conCOLOR_TITLE_BAR
    End With
End Sub

Private Sub picClosed_Click()
    Unload Me
End Sub

Private Sub picClosed_Resize()
    On Error Resume Next
    lblClose.Move picClosed.ScaleWidth / 2 + lblClose.Width / 2, (picClosed.ScaleHeight - lblClose.Height) / 2
End Sub

Private Sub picTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mblnMove Then
        mMoveX = mudtPoint.X - mudtRect.Left
        mMoveY = mudtPoint.Y - mudtRect.Top
        mblnMoveStart = True
    End If
End Sub

Private Sub picTop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRet As Long
    If mblnMoveStart Then
        lngRet = MoveWindow(Me.hWnd, mudtPoint.X - mMoveX, mudtPoint.Y - mMoveY, mudtRect.Right - mudtRect.Left, mudtRect.Bottom - mudtRect.Top, -1)
    End If
End Sub

Private Sub picTop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call GetWindowRect(Me.hWnd, mudtRect)
    Call GetWindowRect(picClosed.hWnd, mudtRectClose)
    mblnMoveStart = False
End Sub

Private Sub picTop_Resize()
    On Error Resume Next
    lblFrmName.Move 120, picTop.ScaleHeight / 2 - lblFrmName.Height / 2
    picClosed.Move picTop.ScaleWidth - picTop.Height, 0, picTop.Height, picTop.Height
End Sub

Private Sub tmrTime_Timer()
    Dim lngRet As Long
    If tmrTime.Tag = "" Then
        Call GetWindowRect(Me.hWnd, mudtRect)
        Call GetWindowRect(picClosed.hWnd, mudtRectClose)
        tmrTime.Tag = "1" '首次记录窗体位置
    End If
    lngRet = GetCursorPos(mudtPoint)
    '判断鼠标指针是否位于窗体拖动区
    If PtInRect(mudtRect, mudtPoint.X, mudtPoint.Y) Then
       mblnMove = True
    Else
       mblnMove = False
    End If
    If PtInRect(mudtRectClose, mudtPoint.X, mudtPoint.Y) Then
        picClosed.BackColor = "&H" & Hex(RGB(212, 64, 39))  '红色
    Else
        picClosed.BackColor = picTop.BackColor
    End If
End Sub

Private Sub vsReason_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsReason
        '拒绝理由允许编辑
        If COL_用药理由 = Col Then
            Cancel = False
        Else
            Cancel = True
        End If
    End With
End Sub

Private Sub InitReason()
'功能: 初始化未审核医嘱列表
    Dim strCol As String, arrHead As Variant
    Dim i As Long
    If mbytStyle = 0 Then
        strCol = "医嘱内容,3500,4;审查详情,3000,4;用药理由,2000,4"
    Else
        strCol = "医嘱内容,3500,4;审查详情,5000,4;用药理由"
    End If
    arrHead = Split(strCol, ";")
    With vsReason
        .Redraw = flexRDNone
        .Clear
        .FixedRows = 1: .FixedCols = 0
        .Cols = UBound(arrHead) + 1
        .Rows = .FixedRows

        .Editable = flexEDKbdMouse
        .AllowUserResizing = flexResizeColumns

        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, i) = Split(arrHead(i), ",")(0)

            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(i) = False
                .ColWidth(i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(i) = True
                .ColWidth(i) = 0
            End If
        Next
        .Redraw = True
    End With
End Sub
