VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPassResultZL 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   ClientHeight    =   9900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13410
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9900
   ScaleWidth      =   13410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Timer tmrTime 
      Interval        =   50
      Left            =   11160
      Top             =   480
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   7935
      Left            =   360
      ScaleHeight     =   7935
      ScaleWidth      =   12855
      TabIndex        =   2
      Top             =   840
      Width           =   12855
      Begin VSFlex8Ctl.VSFlexGrid vsInfo 
         Height          =   6975
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   10575
         _cx             =   18653
         _cy             =   12303
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         MouseIcon       =   "frmPassResultZL.frx":0000
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483641
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483643
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   0
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   2
         SelectionMode   =   0
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   300
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   500
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPassResultZL.frx":08DA
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
         AutoSizeMouse   =   0   'False
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
      Begin VB.Line linSplit 
         BorderColor     =   &H80000011&
         Index           =   0
         X1              =   240
         X2              =   13080
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line linSplit 
         BorderColor     =   &H00808000&
         Index           =   1
         X1              =   -120
         X2              =   12840
         Y1              =   7920
         Y2              =   7920
      End
   End
   Begin VB.PictureBox picBottom 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   -120
      ScaleHeight     =   615
      ScaleWidth      =   13455
      TabIndex        =   1
      Top             =   9240
      Width           =   13455
      Begin VB.PictureBox picBtn 
         Appearance      =   0  'Flat
         BackColor       =   &H00D48A00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   2
         Left            =   3000
         ScaleHeight     =   420
         ScaleWidth      =   1095
         TabIndex        =   12
         Top             =   120
         Width           =   1100
         Begin VB.Label lblBtn 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "������Ϣ"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   2
            Left            =   120
            TabIndex        =   13
            Top             =   120
            Width           =   900
         End
      End
      Begin VB.PictureBox picBtn 
         Appearance      =   0  'Flat
         BackColor       =   &H00D48A00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   1
         Left            =   6000
         ScaleHeight     =   420
         ScaleWidth      =   1095
         TabIndex        =   7
         Top             =   120
         Width           =   1100
         Begin VB.Label lblBtn 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   1
            Left            =   360
            TabIndex        =   9
            Top             =   120
            Width           =   450
         End
      End
      Begin VB.PictureBox picBtn 
         Appearance      =   0  'Flat
         BackColor       =   &H00D48A00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   0
         Left            =   4680
         ScaleHeight     =   420
         ScaleWidth      =   1095
         TabIndex        =   6
         Top             =   120
         Width           =   1100
         Begin VB.Label lblBtn 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "�޸�"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   210
            Index           =   0
            Left            =   360
            TabIndex        =   8
            Top             =   120
            Width           =   450
         End
      End
      Begin VB.Label lblDeclare 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E39F22&
         Height          =   180
         Left            =   480
         TabIndex        =   10
         Top             =   240
         Width           =   780
      End
   End
   Begin VB.PictureBox picTop 
      BackColor       =   &H00D48A00&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      ScaleHeight     =   495
      ScaleWidth      =   12975
      TabIndex        =   0
      Top             =   0
      Width           =   12975
      Begin VB.PictureBox picClosed 
         Appearance      =   0  'Flat
         BackColor       =   &H00D48A00&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   500
         Left            =   12480
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
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
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
         Caption         =   "�������"
         BeginProperty Font 
            Name            =   "����"
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
      BorderColor     =   &H80000010&
      Index           =   3
      X1              =   0
      X2              =   12720
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line linScope 
      Index           =   2
      X1              =   13320
      X2              =   13320
      Y1              =   0
      Y2              =   10560
   End
   Begin VB.Line linScope 
      Index           =   1
      X1              =   -240
      X2              =   13320
      Y1              =   10560
      Y2              =   10560
   End
   Begin VB.Line linScope 
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   8040
   End
End
Attribute VB_Name = "frmPassResultZL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mMoveX As Long, mMoveY As Long  '��¼�����ƶ�ǰ���������Ͻ������ָ��λ�ü���ݺ����
Private mudtRect As RECT
Private mudtRectClose As RECT
Private mudtPoint As POINTAPI
Private mblnMoveStart As Boolean '�ж��ƶ��Ƿ�ʼ
Private mblnMove As Boolean

'-------------------------------------------------------------------------------
Private mrsMsg      As ADODB.Recordset
Private mfrmDrug    As frmPassDrug

Private mbytResult  As Byte          '1-�޸Ĵ���;2-������;3-�޸�������Ϣ
Private mbytModel   As Byte          '0-ҽ���༭;1-ҽ�����;2-��ʾ���������Ϣ
Private mblnHaveOut  As Boolean      'T-����Ժ��ִ����ҩ
Private mstrFontUnderLine As String   '����»�����  �к�|��1
Private mbytOpen    As Byte

Public Enum E_COLINDEX
    COL_��ʾ = 0
    COL_���� = 1
    COL_˵���� = 2
End Enum

Public Function ShowMe(frmParent As Object, rsMsg As ADODB.Recordset, ByVal bytModel As Byte, _
    Optional ByRef bytResult As Byte, Optional ByVal blnIsHaveOut As Boolean) As Boolean
'����:��ʾ�����
'����:
'   bytResult=1-�޸Ĵ���;2-������
    If bytModel = 2 Then
        Set mrsMsg = zlDatabase.CopyNewRec(rsMsg)
    Else
        Set mrsMsg = rsMsg
    End If
    mbytResult = 0
    mbytModel = bytModel
    mblnHaveOut = blnIsHaveOut
    mbytOpen = 1
    Me.Show 1, frmParent
    bytResult = mbytResult
End Function

Private Sub Form_Load()
    Dim blnOK As Boolean
    picTop.BackColor = conCOLOR_TITLE_BAR
    picBtn(2).Visible = False
    If mbytModel = 0 Then
        lblBtn(1) = "����"
        picBtn(0).Visible = True
        picBtn(1).Visible = True
        picBtn(0).BackColor = &HD48A00
        picBtn(1).BackColor = &HD48A00
        picBtn(2).BackColor = &HD48A00
        
        mrsMsg.Filter = "Category=0 And Light = 4" '�ڵ�(����ܿ�ҩƷ)��ֹ����
        If mrsMsg.RecordCount > 0 Then
            If picBtn(1).Enabled Then
                picBtn(1).Enabled = False
                picBtn(1).BackColor = "&H" & Hex(RGB(144, 158, 149))
            End If
        Else
            mrsMsg.Filter = "Category=0 And Light = 2"
            If mrsMsg.RecordCount > 0 Then
                If gbytBlackLamp = 1 Then  '�����´������ҩ
                    picBtn(1).Enabled = True
                Else
                    If gbytOutBlackLamp = 1 And mblnHaveOut Then '�������´�Ժ�����
                        picBtn(1).Enabled = True
                    Else
                        picBtn(1).Enabled = False
                        picBtn(1).BackColor = "&H" & Hex(RGB(144, 158, 149))
                    End If
                End If
            Else
                picBtn(1).Enabled = True
            End If
        End If
        
        mrsMsg.Filter = "Category=1" '���������ֹ����
        If mrsMsg.RecordCount > 0 Then
            If picBtn(1).Enabled Then
                picBtn(1).Enabled = False
                picBtn(1).BackColor = "&H" & Hex(RGB(144, 158, 149))
            End If
            picBtn(2).Visible = True: picBtn(2).Enabled = True
        End If
    Else
        lblBtn(1).Caption = "�ر�"
        picBtn(0).Visible = False
        picBtn(1).Visible = True
    End If
    
    Call LoadMsg
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picTop.Move 15, 15, Me.ScaleWidth - 30, 500
    picBottom.Move 15, Me.ScaleHeight - 915, Me.ScaleWidth - 30, 900
    picMain.Move 240, picTop.Height + picTop.Top, Me.ScaleWidth - 300, Me.ScaleHeight - Me.picBottom.Height - Me.picTop.Height - 60
    
    'Left
    With linScope(0)
        .X1 = 0: .X2 = 0: .Y1 = 0: .Y2 = Me.ScaleHeight
        .BorderColor = conCOLOR_TITLE_BAR
        '&H00808080&
        '&H80000010& '��ť��Ӱ
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

Private Sub Form_Unload(Cancel As Integer)
    mbytOpen = 0
End Sub

Public Function IsOpen() As Boolean
    IsOpen = mbytOpen = 1
End Function

Private Sub lblBtn_Click(Index As Integer)
    Dim strMsg As String
    Dim i As Long
    
    If mbytModel = 0 Then
        mbytResult = Index + 1
    Else
        mbytResult = 0
    End If
    Unload Me
End Sub

Private Sub lblClose_Click()
    Call picBtn_Click(0)
    Unload Me
End Sub

Private Sub lblDeclare_Click()
    Call frmDeclare.Show(vbModal, Me)
End Sub

Private Sub lblDeclare_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblDeclare.Font.Underline = True
End Sub

Private Sub picBottom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblDeclare.Font.Underline = False
    If picBtn(0).BackColor <> conCOLOR_BULE And picBtn(0).Enabled Then picBtn(0).BackColor = conCOLOR_BULE
    If picBtn(1).BackColor <> conCOLOR_BULE And picBtn(1).Enabled Then picBtn(1).BackColor = conCOLOR_BULE
    If picBtn(2).BackColor <> conCOLOR_BULE And picBtn(2).Enabled Then picBtn(2).BackColor = conCOLOR_BULE
End Sub

Private Sub picBtn_Click(Index As Integer)
    lblBtn_Click Index
End Sub

Private Sub picBtn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    picBtn(Index).BackColor = conCOLOR_BULELIGHT
End Sub

Private Sub picBtn_Resize(Index As Integer)
    lblBtn(Index).Move picBtn(Index).ScaleWidth / 2 + lblBtn(Index).Width / 2, picBtn(Index).ScaleHeight / 2 - lblBtn(Index) / 2
End Sub

Private Sub picClosed_Click()
    Call lblClose_Click
End Sub

Private Sub picClosed_Resize()
    On Error Resume Next
    lblClose.Move picClosed.ScaleWidth / 2 + lblClose.Width / 2, (picClosed.ScaleHeight - lblClose.Height) / 2
End Sub

Private Sub picTop_Resize()
    On Error Resume Next
    lblFrmName.Move 120, picTop.ScaleHeight / 2 - lblFrmName.Height / 2
    picClosed.Move picTop.ScaleWidth - picClosed.Width, picTop.ScaleHeight / 2 - picClosed.Height / 2
End Sub

Private Sub picMain_Resize()
    With linSplit(0)
        .X1 = 0: .X2 = picMain.ScaleWidth
        .Y1 = 0: .Y2 = 0
        .BorderColor = vbWhite
    End With
    
    With linSplit(1)
        .X1 = 0: .X2 = picMain.ScaleWidth
        .Y1 = picMain.ScaleHeight - 15: .Y2 = picMain.ScaleHeight - 15
    End With
    vsInfo.Move 0, 240, picMain.ScaleWidth, picMain.ScaleHeight - 255
End Sub

Private Sub picBottom_Resize()
    On Error Resume Next
    If mbytModel = 0 Then
        picBtn(0).Move picBottom.ScaleWidth / 2 - picBtn(0).Width - 60, picBottom.ScaleHeight / 2 - picBtn(0).Height / 2
        picBtn(1).Move picBottom.ScaleWidth / 2 + 60, picBottom.ScaleHeight / 2 - picBtn(1).Height / 2
    ElseIf mbytModel = 1 Then
        picBtn(1).Move picBottom.ScaleWidth / 2 - picBtn(1).Width / 2, picBottom.ScaleHeight / 2 - picBtn(1).Height / 2
    End If
    picBtn(2).Move picBtn(0).Left - picBtn(0).Width * 2, picBtn(0).Top, picBtn(0).Width, picBtn(0).Height
    lblDeclare.Move 240, (picBottom.ScaleHeight - lblDeclare.Height) / 2
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
    Dim udtRect As RECT
    
    If tmrTime.Tag = "" Then
        Call GetWindowRect(Me.hWnd, mudtRect)
        Call GetWindowRect(picClosed.hWnd, mudtRectClose)
        tmrTime.Tag = "1" '�״μ�¼����λ��
    End If
    lngRet = GetCursorPos(mudtPoint)
    '�ж����ָ���Ƿ�λ�ڴ����϶���
    If PtInRect(mudtRect, mudtPoint.X, mudtPoint.Y) Then
       mblnMove = True
    Else
       mblnMove = False
    End If
    If PtInRect(mudtRectClose, mudtPoint.X, mudtPoint.Y) Then
        picClosed.BackColor = "&H" & Hex(RGB(212, 64, 39))  '��ɫ
    Else
        picClosed.BackColor = picTop.BackColor
    End If
End Sub

Private Sub LoadMsg()
'����:�����������
    Dim intLight As Integer
    Dim intLightMax As Integer
    Dim i As Long
    Dim lngRow As Long
    Dim strType As String
    
    With vsInfo
        .Redraw = flexRDNone
        .Rows = 0
        .Cols = 3
        .ColWidth(COL_��ʾ) = 300
        .ColWidth(COL_����) = 10500
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(COL_˵����) = flexAlignLeftCenter
        
        .RowHeightMin = 220
        .AutoResize = True
        .AllowUserResizing = flexResizeRows
        .AutoSizeMode = flexAutoSizeRowHeight
        .WordWrap = True
        '����
        mrsMsg.Filter = "Category=1"
        For i = 1 To mrsMsg.RecordCount
            If i = 1 Then
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, COL_����) = "�����á�"
                .Cell(flexcpFontBold, .Rows - 1, COL_����) = True
                .Cell(flexcpForeColor, .Rows - 1, COL_����) = RGB(203, 0, 0)
                .Cell(flexcpFontSize, .Rows - 1, COL_����) = 14
                .Rows = .Rows + 1
            End If
            .Rows = .Rows + 1
            If strType <> mrsMsg!Type & "" Then
                .TextMatrix(.Rows - 1, COL_����) = mrsMsg!Type & ":"
                .Cell(flexcpFontBold, .Rows - 1, COL_����) = True
                .Rows = .Rows + 1
            End If
            
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, COL_����) = mrsMsg!describ & Space(4) & mrsMsg!remaks
            strType = mrsMsg!Type & ""
            .Rows = .Rows + 1
            mrsMsg.MoveNext
        Next
        
        'ҩ�����
        If mbytModel = 2 Then
            mrsMsg.Filter = "Category=0"
        Else
            mrsMsg.Filter = "Category=0 And Tag = 0"
        End If
        mrsMsg.Sort = "WarnLevel DESC, Type ASC"
        If mbytModel = 2 Then intLightMax = 10
        
        For i = 1 To mrsMsg.RecordCount
            If mrsMsg!Light = 4 And intLight <> 4 Then
                .Rows = .Rows + 1
                If intLightMax < mrsMsg!WarnLevel Then
                    Call gobjFrm.SetLight("��"): intLightMax = mrsMsg!WarnLevel  '��ֹ
                End If
                
                .Cell(flexcpPictureAlignment, .Rows - 1, COL_��ʾ) = flexPicAlignLeftCenter
                .Cell(flexcpPicture, .Rows - 1, COL_��ʾ) = frmIcons.imgPass.ListImages("��_4").Picture
                
                .TextMatrix(.Rows - 1, COL_����) = "����ֹ��"
                .Cell(flexcpFontBold, .Rows - 1, COL_����) = True
                .Cell(flexcpForeColor, .Rows - 1, COL_����) = vbBlack
                .Cell(flexcpFontSize, .Rows - 1, COL_����) = 14
                .Rows = .Rows + 1
                intLight = 4
                strType = ""
            ElseIf mrsMsg!Light = 2 And intLight <> 2 Then
                .Rows = .Rows + 1
                If intLightMax < mrsMsg!WarnLevel Then
                    Call gobjFrm.SetLight("��"): intLightMax = mrsMsg!WarnLevel  '����
                End If
                
                .Cell(flexcpPictureAlignment, .Rows - 1, COL_��ʾ) = flexPicAlignLeftCenter
                .Cell(flexcpPicture, .Rows - 1, COL_��ʾ) = frmIcons.imgPass.ListImages("��_4").Picture
                
                .TextMatrix(.Rows - 1, COL_����) = "�����á�"
                .Cell(flexcpFontBold, .Rows - 1, COL_����) = True
                .Cell(flexcpForeColor, .Rows - 1, COL_����) = RGB(203, 0, 0)
                .Cell(flexcpFontSize, .Rows - 1, COL_����) = 14
                .Rows = .Rows + 1
                intLight = 2
                strType = ""
            ElseIf mrsMsg!Light = 1 And intLight <> 1 Then
                If .Rows >= 0 Then .Rows = .Rows + 1
                If intLightMax < mrsMsg!WarnLevel Then
                    Call gobjFrm.SetLight("��"): intLightMax = mrsMsg!WarnLevel
                End If

                .Cell(flexcpPictureAlignment, .Rows - 1, COL_��ʾ) = flexPicAlignLeftCenter
                .Cell(flexcpPicture, .Rows - 1, COL_��ʾ) = frmIcons.imgPass.ListImages("��_4").Picture
                
                .TextMatrix(.Rows - 1, COL_����) = "�����á�"
                .Cell(flexcpFontBold, .Rows - 1, COL_����) = True
                .Cell(flexcpForeColor, .Rows - 1, COL_����) = RGB(239, 90, 0)
                .Cell(flexcpFontSize, .Rows - 1, COL_����) = 14
                .Rows = .Rows + 1
                intLight = 1
                strType = ""
            ElseIf mrsMsg!Light = 3 And intLight <> 3 Then
                If .Rows >= 0 Then .Rows = .Rows + 1
                If intLightMax < mrsMsg!WarnLevel Then
                    Call gobjFrm.SetLight("��"): intLightMax = mrsMsg!WarnLevel
                End If
                
                .Cell(flexcpPictureAlignment, .Rows - 1, COL_��ʾ) = flexPicAlignLeftCenter
                .Cell(flexcpPicture, .Rows - 1, COL_��ʾ) = frmIcons.imgPass.ListImages("��_4").Picture
                
                .TextMatrix(.Rows - 1, COL_����) = "��ע���"
                .Cell(flexcpFontBold, .Rows - 1, COL_����) = True
                .Cell(flexcpForeColor, .Rows - 1, COL_����) = RGB(255, 192, 0)
                .Cell(flexcpFontSize, .Rows - 1, COL_����) = 14
                .Rows = .Rows + 1
                intLight = 1
                strType = ""
            End If
            .Rows = .Rows + 1
            If strType <> mrsMsg!Type & "" Then
                .TextMatrix(.Rows - 1, COL_����) = mrsMsg!Type & ":"
                .Cell(flexcpFontBold, .Rows - 1, COL_����) = True
                .Rows = .Rows + 1
            End If
            .TextMatrix(.Rows - 1, COL_����) = mrsMsg!describ & Space(4) & mrsMsg!remaks
            If mrsMsg!DrugCode <> "" Then
                .TextMatrix(.Rows - 1, COL_˵����) = "���鿴˵���顿"
                .Cell(flexcpData, .Rows - 1, COL_˵����) = mrsMsg!DrugCode & ""
            End If
            strType = mrsMsg!Type & ""
            .Rows = .Rows + 1
            mrsMsg.MoveNext
        Next
        .Cell(flexcpForeColor, 0, COL_˵����, .Rows - 1, COL_˵����) = conCOLOR_BULE
        .Redraw = flexRDDirect
        .AutoSize COL_����, COL_����, , 45
    End With
End Sub

Private Sub DelFontUnderLine()
    Dim arrTmp As Variant
    Dim lngColor As Long
    
    With vsInfo
        arrTmp = Split(mstrFontUnderLine, ",")
        If UBound(arrTmp) >= 2 Then
            lngColor = Val(arrTmp(2))
        Else
            lngColor = vbBlack
        End If
        .Cell(flexcpForeColor, arrTmp(0), arrTmp(1)) = lngColor
        .Cell(flexcpFontUnderline, arrTmp(0), arrTmp(1)) = False
        mstrFontUnderLine = ""
    End With
End Sub

Private Sub vsInfo_Click()
    Dim strDrugCode As String
    
    With vsInfo
        If .Col = COL_˵���� Then
            If CStr(.Cell(flexcpData, .Row, .Col)) <> "" Then
                strDrugCode = CStr(.Cell(flexcpData, .Row, .Col))
                Call GetDrugInstructions(Me, mfrmDrug, 1, strDrugCode)
            End If
        End If
    End With
End Sub

Private Sub vsInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long, lngCol As Long
    Dim lngColor As Long
    
    With vsInfo
        If .Enabled Then .SetFocus
        lngRow = .MouseRow: lngCol = .MouseCol
        If lngRow < 0 Or lngCol < 0 Then Exit Sub
        .MousePointer = flexDefault
        If mstrFontUnderLine <> "" Then
            DelFontUnderLine
        End If
        If lngCol = COL_˵���� And .TextMatrix(lngRow, lngCol) = "���鿴˵���顿" Then
           .Cell(flexcpFontUnderline, lngRow, lngCol) = True
           .Cell(flexcpForeColor, lngRow, lngCol) = vbBlue
           mstrFontUnderLine = lngRow & "," & lngCol & "," & conCOLOR_TITLE_BAR
           .MousePointer = flexCustom
        End If
    End With
End Sub
