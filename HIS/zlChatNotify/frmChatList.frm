VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmChatList 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "未读消息"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   10005
   ClientWidth     =   3600
   Icon            =   "frmChatList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   Begin VSFlex8Ctl.VSFlexGrid vsInfo 
      Height          =   975
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   3375
      _cx             =   5953
      _cy             =   1720
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
      BackColorSel    =   -2147483633
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   0
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
   Begin VB.Timer tmrTime 
      Interval        =   50
      Left            =   1680
      Top             =   1560
   End
   Begin VB.PictureBox picTop 
      Appearance      =   0  'Flat
      BackColor       =   &H00D48A00&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   500
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   3600
      TabIndex        =   0
      Top             =   0
      Width           =   3600
      Begin VB.PictureBox picClosed 
         Appearance      =   0  'Flat
         BackColor       =   &H00D48A00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   2400
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   1
         Top             =   0
         Width           =   480
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
            Left            =   90
            TabIndex        =   2
            Top             =   120
            Width           =   300
         End
      End
      Begin VB.Label lblFrmName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "讨论消息"
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
         Left            =   0
         TabIndex        =   3
         Top             =   120
         Width           =   900
      End
   End
   Begin VB.Timer tmrShow 
      Interval        =   1
      Left            =   2280
      Top             =   1560
   End
   Begin VB.Line linScope 
      Index           =   0
      X1              =   600
      X2              =   600
      Y1              =   600
      Y2              =   8640
   End
   Begin VB.Line linScope 
      Index           =   1
      X1              =   960
      X2              =   960
      Y1              =   240
      Y2              =   8280
   End
   Begin VB.Line linScope 
      Index           =   2
      X1              =   480
      X2              =   480
      Y1              =   120
      Y2              =   8160
   End
   Begin VB.Line linScope 
      Index           =   3
      X1              =   360
      X2              =   360
      Y1              =   120
      Y2              =   8160
   End
End
Attribute VB_Name = "frmChatList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mudtRect As RECT
Private mudtRectClose As RECT
Private mudtPoint As POINTAPI

Private Const CON_MIN_ROW_HEIGHT As Long = 360
Private Const CON_MAX_ROW_NUM As Long = 10



Public Function ShowMe(ByVal bytFunc As Byte) As Boolean
'参数:
'   bytFunc=0 显示未读讨论清单;=1 默认打开第一个未读讨论
    Dim udtPoint As POINTAPI
    Dim lngLeft As Long
    Dim lngPos As Long
    
    If grsList Is Nothing Then Exit Function
    If grsList.RecordCount = 0 Then Exit Function
    
    If grsList.RecordCount = 1 And bytFunc = 1 Then
        If OpenChatRoom(grsList!Url & "", grsList!Subject & "") Then
            '消息读取成功后,将状态更新为接收状态
            If gobjMain.UpdateChatState(Val(grsList!Id & ""), 1) Then
                grsList.Filter = "ID=" & Val(grsList!Id & "")
                If Not grsList.EOF Then grsList.Delete: grsList.Filter = ""
                If grsList.RecordCount = 0 Then
                    Call gfrmMain.SetIcon(2)
                End If
                If Me.Visible Then
                    If Not LoadList Then Exit Function
                End If
            End If
        End If
    Else
        Me.Show 0
        Call GetCursorPos(udtPoint)
        lngLeft = udtPoint.X * Screen.TwipsPerPixelX '鼠标位置LEFT
        Me.Top = Screen.Height
        lngPos = Screen.Width - Me.Width
        Me.Left = IIf(lngPos > lngLeft, lngLeft, lngPos)
        On Error Resume Next
        tmrShow.Enabled = True
        On Error GoTo 0
        Call Form_Resize
    End If
    
    ShowMe = True
End Function

Public Function RefreshList() As Boolean
'功能:刷新数据
    Dim lngRow As Long
    lngRow = vsInfo.Rows
    If Not LoadList Then Exit Function
    If lngRow < CON_MAX_ROW_NUM Then
        Me.Top = Me.Top - CON_MIN_ROW_HEIGHT
    End If
    Call Form_Resize
    RefreshList = True
End Function

Private Function LoadList() As Boolean
    Dim i As Long
    grsList.Filter = ""
    If grsList.RecordCount = 0 Then Unload Me: Exit Function
    With vsInfo
        .Redraw = flexRDNone
        .Clear
        .Rows = IIf(grsList.RecordCount > CON_MAX_ROW_NUM, CON_MAX_ROW_NUM, grsList.RecordCount)
        .Cols = 1
        .ColWidth(0) = 3000
        .RowHeightMin = CON_MIN_ROW_HEIGHT
        .ExtendLastCol = True
        For i = 0 To .Rows - 1
            .TextMatrix(i, 0) = grsList!Subject & ""
            .Cell(flexcpData, i, 0) = grsList!Url & ""
            .RowData(i) = Val(grsList!Id)
            grsList.MoveNext
        Next
        .Redraw = flexRDDirect
        grsList.Filter = ""
    End With
    LoadList = True
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gblnShow = False
End Sub

Private Sub lblClose_Click()
    Unload Me
End Sub

Private Sub picClosed_Click()
    Unload Me
End Sub

Private Sub picClosed_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call GetWindowRect(picClosed.hWnd, mudtRectClose)
End Sub

Private Sub picTop_Resize()
    Dim lngW As Long
    
    On Error Resume Next
    lngW = picTop.Height - 20
    lblFrmName.Move 120, picTop.ScaleHeight / 2 - lblFrmName.Height / 2
    picClosed.Move picTop.ScaleWidth - picTop.Height, 0, lngW, lngW
End Sub

Private Sub picClosed_Resize()
    On Error Resume Next
    lblClose.Move picClosed.ScaleWidth / 2 + lblClose.Width / 2, (picClosed.ScaleHeight - lblClose.Height) / 2
End Sub

Private Sub tmrShow_Timer()
    If Me.Top > Screen.Height - Me.Height - 600 Then  '粗略估计任务栏高度600
        Me.Top = Me.Top - 60
    Else
        tmrShow.Enabled = False
    End If
     
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    Call LoadList
    '界面颜色设置
    picTop.BackColor = conCOLOR_TITLE_BAR
    For i = linScope.LBound To linScope.UBound
        linScope(i).BorderColor = conCOLOR_TITLE_BAR
    Next
    
    gblnShow = True
End Sub

Private Sub Form_Resize()
    Dim lngFormW As Long
    Dim lngTopH As Long
    
    On Error Resume Next
    
    With Me
        lngTopH = 500
        lngFormW = 3600
        .Width = lngFormW
        .Height = vsInfo.Rows * CON_MIN_ROW_HEIGHT + lngTopH + 240
        picTop.Move 15, 15, Me.ScaleWidth - 30, lngTopH
        vsInfo.Move 120, lngTopH + 120, .ScaleWidth - 240, .ScaleHeight - lngTopH - 240
    End With
    
    'Left
    With linScope(0)
        .X1 = 0: .X2 = 0: .Y1 = 0: .Y2 = Me.ScaleHeight
        '&H00808080&
        '&H80000010& '按钮阴影
    End With
    'bottom
    With linScope(1)
        .X1 = 0: .X2 = Me.ScaleWidth: .Y1 = Me.ScaleHeight - 15: .Y2 = Me.ScaleHeight - 15
    End With
    'right
    With linScope(2)
        .X1 = Me.ScaleWidth - 15: .X2 = Me.ScaleWidth - 15: .Y1 = 0: .Y2 = Me.ScaleHeight - 15
    End With
    'Top
    With linScope(3)
        .X1 = 0: .X2 = Me.ScaleWidth - 15: .Y1 = 0: .Y2 = 0
    End With
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
    If PtInRect(mudtRectClose, mudtPoint.X, mudtPoint.Y) Then
        picClosed.BackColor = "&H" & Hex(RGB(212, 64, 39))  '红色
    Else
        picClosed.BackColor = picTop.BackColor
    End If
End Sub

Private Sub vsInfo_Click()
    Dim lngRow As Long
    Dim lngCol As Long
    Dim dblId As Double
    Dim lngRows As Long
    
    With vsInfo
        lngRow = .MouseRow: lngCol = .MouseCol
        If lngRow < 0 Or lngCol < 0 Then Exit Sub
        dblId = Val(.RowData(lngRow))
        If dblId = 0 Then Exit Sub
        If OpenChatRoom(CStr(.Cell(flexcpData, lngRow, 0)), .TextMatrix(lngRow, 0)) Then
            '消息读取成功后,将状态更新为接收状态
            If gobjMain.UpdateChatState(dblId, 1) Then
                grsList.Filter = "ID=" & dblId
                If Not grsList.EOF Then grsList.Delete: grsList.Filter = ""
                If grsList.RecordCount = 0 Then Call gfrmMain.SetIcon(2): Unload Me: Exit Sub
                lngRows = .Rows
                If LoadList Then
                    If lngRows > .Rows Then
                        Call Form_Resize
                        Me.Top = Me.Top + CON_MIN_ROW_HEIGHT
                        Call Form_Resize
                    End If
                Else
                    Exit Sub
                End If
            End If
        End If
        
    End With
End Sub

Private Sub vsInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long
    
    With vsInfo
        lngRow = .MouseRow
        If lngRow = -1 Then Exit Sub
        .Row = lngRow
    End With
End Sub
