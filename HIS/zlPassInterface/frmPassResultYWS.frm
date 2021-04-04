VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPassResultYWS 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   4590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer tmrTime 
      Interval        =   50
      Left            =   1200
      Top             =   3240
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   0
      ScaleHeight     =   3975
      ScaleWidth      =   7335
      TabIndex        =   1
      Top             =   480
      Width           =   7335
      Begin RichTextLib.RichTextBox rtfInfo 
         Height          =   1335
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   2355
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         MousePointer    =   99
         Appearance      =   0
         TextRTF         =   $"frmPassResultYWS.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox picTop 
      BackColor       =   &H00D48A00&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   7335
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.PictureBox picClosed 
         Appearance      =   0  'Flat
         BackColor       =   &H00D48A00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   500
         Left            =   6720
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   4
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
            TabIndex        =   5
            Top             =   120
            Width           =   300
         End
      End
      Begin VB.Label lblFrmName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "审查详情"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   210
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   840
      End
   End
   Begin VB.Line linScope 
      BorderColor     =   &H80000010&
      Index           =   3
      X1              =   0
      X2              =   12720
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line linScope 
      Index           =   2
      X1              =   10560
      X2              =   10560
      Y1              =   0
      Y2              =   10560
   End
   Begin VB.Line linScope 
      Index           =   1
      X1              =   240
      X2              =   13800
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Line linScope 
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   8040
   End
End
Attribute VB_Name = "frmPassResultYWS"
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

'-------------------------------------------------------------------------------
Private mrsMsg      As ADODB.Recordset

Public Function ShowMe(rsMsg As ADODB.Recordset) As Boolean
'功能:显示审查结果
'参数:
    Set mrsMsg = rsMsg
    Me.Show 1
End Function

Private Sub Form_Load()
    Call LoadMsg
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picTop.Move 15, 15, Me.ScaleWidth - 30, 500
    picMain.Move 120, picTop.Height + picTop.Top, Me.ScaleWidth - 150, Me.ScaleHeight - Me.picTop.Height - 60
    
    'Left
    With linScope(0)
        .X1 = 0: .X2 = 0: .Y1 = 0: .Y2 = Me.ScaleHeight
        .BorderColor = &H80000010
        '&H00808080&
        '&H80000010& '按钮阴影
    End With
    'bottom
    With linScope(1)
        .X1 = 0: .X2 = Me.ScaleWidth: .Y1 = Me.ScaleHeight - 15: .Y2 = Me.ScaleHeight - 15
        .BorderColor = &H80000010
    End With
    'right
    With linScope(2)
        .X1 = Me.ScaleWidth - 15: .X2 = Me.ScaleWidth - 15: .Y1 = 0: .Y2 = Me.ScaleHeight - 15
        .BorderColor = &H80000010
    End With
    'Top
    With linScope(3)
        .X1 = 0: .X2 = Me.ScaleWidth - 15: .Y1 = 0: .Y2 = 0
        .BorderColor = &H80000010
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsMsg = Nothing
End Sub

Private Sub lblClose_Click()
    Unload Me
End Sub

Private Sub picClosed_Click()
    Call lblClose_Click
End Sub

Private Sub picClosed_Resize()
    On Error Resume Next
    lblClose.Move picClosed.ScaleWidth / 2 + lblClose.Width / 2, (picClosed.ScaleHeight - lblClose.Height) / 2
End Sub

Private Sub picTop_Resize()
    lblFrmName.Move 120, picTop.ScaleHeight / 2 - lblFrmName.Height / 2
    picClosed.Move picTop.ScaleWidth - picClosed.Width, picTop.ScaleHeight / 2 - picClosed.Height / 2
End Sub

Private Sub picMain_Resize()
    rtfInfo.Move 0, 0, picMain.ScaleWidth, picMain.ScaleHeight
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

Private Sub LoadMsg()
'功能:加载审查详情
    Dim strText As String
    Dim strItemPos As String
    Dim i As Long
     rtfInfo.Text = ""
    strText = vbNewLine
    mrsMsg.Filter = ""
    For i = 1 To mrsMsg.RecordCount
        strItemPos = strItemPos & ";" & Len(strText) & "," & Len(mrsMsg!Title)
        strText = strText & mrsMsg!Title & vbNewLine & vbNewLine & "【详细问题】" & vbNewLine & mrsMsg!Detail & vbNewLine & vbNewLine
        mrsMsg.MoveNext
    Next
    rtfInfo.Text = strText
    Call SetMsgFont(Mid(strItemPos, 2)) '设置字体
    rtfInfo.Visible = True
End Sub

Private Sub SetMsgFont(ByVal strTilePos As String)
'功能：对RichTextBox进行字体设置
'参数 strTilePos 记录标题行位置格式为，标题1起始位置，标题长度;标题2起始位置，标题长度
    Dim arrTmp As Variant, i As Long
    
    On Error Resume Next
    
    If Len(Trim(strTilePos)) = 0 Then Exit Sub
    arrTmp = Split(Trim(strTilePos), ";")

    With rtfInfo
        For i = LBound(arrTmp) To UBound(arrTmp)
            .SelStart = Split(arrTmp(i), ",")(0)
            .SelLength = Split(arrTmp(i), ",")(1)
            .SelFontSize = 12
            .SelFontName = "宋体"
            .SelBold = True
            .SelUnderline = False
            .SelColor = vbBlack
            .SelLength = 0
        Next

        .SelStart = 0 '光标移动到开始
    End With
End Sub
