VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPassResult 
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
      Top             =   600
   End
   Begin VB.PictureBox picMain 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   7935
      Left            =   240
      ScaleHeight     =   7935
      ScaleWidth      =   12855
      TabIndex        =   2
      Top             =   1080
      Width           =   12855
      Begin VSFlex8Ctl.VSFlexGrid vsInfo 
         Height          =   2055
         Left            =   0
         TabIndex        =   7
         Top             =   120
         Width           =   7575
         _cx             =   13361
         _cy             =   3625
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
         MouseIcon       =   "frmPassResult.frx":0000
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
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   10
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   500
         RowHeightMax    =   0
         ColWidthMin     =   500
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
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
         X1              =   0
         X2              =   12840
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
      Left            =   -360
      ScaleHeight     =   615
      ScaleWidth      =   13695
      TabIndex        =   1
      Top             =   9240
      Width           =   13695
      Begin VB.CommandButton cmd 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   4920
         TabIndex        =   5
         Top             =   113
         Width           =   1100
      End
      Begin VB.CommandButton cmd 
         Caption         =   "�޸�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   3600
         TabIndex        =   4
         Top             =   120
         Width           =   1100
      End
   End
   Begin VB.PictureBox picTop 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      ScaleHeight     =   495
      ScaleWidth      =   12975
      TabIndex        =   0
      Top             =   0
      Width           =   12975
      Begin VB.Image imgClose 
         Height          =   240
         Left            =   12600
         Picture         =   "frmPassResult.frx":08DA
         Stretch         =   -1  'True
         ToolTipText     =   "�ر�"
         Top             =   120
         Width           =   240
      End
      Begin VB.Label lblFrmName 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "�������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   840
      End
   End
   Begin VB.Label lblPati 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "����ID: 1232323  ����:���� ������:110"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   3885
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
Attribute VB_Name = "frmPassResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mMoveX As Long, mMoveY As Long  '��¼�����ƶ�ǰ���������Ͻ������ָ��λ�ü���ݺ����
Private mudtRect As RECT
Private mudtPoint As POINTAPI
Private mblnMoveStart As Boolean '�ж��ƶ��Ƿ�ʼ
Private mblnMove As Boolean
'-------------------------------------------------------------------------------
Private mrsMsg      As ADODB.Recordset
Private mstrPati    As String
Private mbytResult  As Byte          '1-�޸Ĵ���;2-������
Private mbytModel   As Byte          '0-ҽ���༭;1-ҽ�����
Private mblnHaveOut  As Boolean      'T-����Ժ��ִ����ҩ
Private mstrFontUnderLine As String   '����»�����  �к�|�к�

Private Enum E_COL
    COL_��� = 0
    COL_���� = 1
    COL_��ʾ = 2
    COL_���� = 3
End Enum

Public Function ShowMe(frmParent As Object, rsMsg As ADODB.Recordset, ByVal strPati As String, ByRef bytResult As Byte, _
    ByVal bytModel As Byte, ByVal blnIsHaveOut As Boolean) As Boolean
'����:��ʾ�����
'����:
'   bytResult=1-�޸Ĵ���;2-������
'   blnIsHaveOut-סԺ�༭���� ��Ժ��ҩ
    Set mrsMsg = rsMsg
    mstrPati = strPati
    mbytResult = 0
    mbytModel = bytModel
    mblnHaveOut = blnIsHaveOut
    Me.Show 1, frmParent
    bytResult = mbytResult
End Function

Private Sub cmd_Click(Index As Integer)
    If mbytModel = 0 Then
        If Index = 1 Then
            If MsgBox("��鷢�ֽ�����ҩ����ȷ��Ҫ������", vbOKCancel + vbQuestion + vbDefaultButton2, gstrSysName) = vbCancel Then
                Exit Sub
            End If
        End If
        mbytResult = Index + 1
    Else
        mbytResult = 0
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Dim blnOK As Boolean
    mstrFontUnderLine = ""
    If mbytModel = 0 Then
        cmd(1).Caption = "����"
        cmd(0).Visible = True
        cmd(1).Visible = True
        mrsMsg.Filter = "Severity=8"    '���ɵȼ�
        blnOK = mrsMsg.RecordCount > 0 And mbytModel = 0
        If blnOK Then
            If gbytBlackLamp = 1 Then  '�����´������ҩ
                cmd(1).Enabled = True
            Else
                If gbytOutBlackLamp = 1 And mblnHaveOut Then '�������´�Ժ�����
                    cmd(1).Enabled = True
                Else
                    cmd(1).Enabled = False
                End If
            End If
        Else
            cmd(1).Enabled = True
        End If
    Else
        cmd(1).Caption = "�ر�"
        cmd(0).Visible = False
        cmd(1).Visible = True
    End If
    
    Call LoadMsg
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picTop.Move 15, 15, Me.ScaleWidth - 30, 500
    picBottom.Move 15, Me.ScaleHeight - 715, Me.ScaleWidth - 30, 700
    lblPati.Move 240, picTop.Top + picTop.Height + 120
    picMain.Move 240, picTop.Height + 500, Me.ScaleWidth - 300, Me.ScaleHeight - Me.picBottom.Height - Me.picTop.Height - 600
    
    'Left
    With linScope(0)
        .X1 = 0: .X2 = 0: .Y1 = 0: .Y2 = Me.ScaleHeight
        .BorderColor = &H80000010
        '&H00808080&
        '&H80000010& '��ť��Ӱ
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

Private Sub imgClose_Click()
    Call cmd_Click(0)
    Unload Me
End Sub

Private Sub picTop_Resize()
    lblFrmName.Move 120, picTop.ScaleHeight / 2 - lblFrmName.Height / 2
    imgClose.Move picTop.ScaleWidth - (imgClose.Width + 120), picTop.ScaleHeight / 2 - imgClose.Height / 2
End Sub

Private Sub picMain_Resize()
    With linSplit(0)
        .X1 = 0: .X2 = picMain.ScaleWidth
        .Y1 = 0: .Y2 = 0
    End With
    
    With linSplit(1)
        .X1 = 0: .X2 = picMain.ScaleWidth
        .Y1 = picMain.ScaleHeight - 15: .Y2 = picMain.ScaleHeight - 15
    End With
    vsInfo.Move 0, linSplit(0).Y1 + 75, picMain.ScaleWidth, linSplit(1).Y1 - linSplit(0).Y1 - 135
End Sub

Private Sub picBottom_Resize()
    If mbytModel = 0 Then
        cmd(0).Move picBottom.ScaleWidth / 2 - cmd(0).Width - 60, 0
        cmd(1).Move picBottom.ScaleWidth / 2 + 60, 0
    ElseIf mbytModel = 1 Then
        cmd(1).Move picBottom.ScaleWidth / 2 - cmd(1).Width / 2, 0
    End If
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
    mblnMoveStart = False
End Sub

Private Sub tmrTime_Timer()
    Dim lngRet As Long

    lngRet = GetCursorPos(mudtPoint)
    '�ж����ָ���Ƿ�λ�ڴ����϶���
    If PtInRect(mudtRect, mudtPoint.X, mudtPoint.Y) Then
       mblnMove = True
    Else
       mblnMove = False
    End If
End Sub

Private Sub LoadMsg()
'����:�����������
    Dim strText As String
    Dim i As Long
    Dim strItem As String       '��¼�ı�����λ��
    Dim strItemPos As String
    
    lblPati.Caption = mstrPati
    
    mrsMsg.Filter = ""

    With vsInfo
        .Cols = 4
        .Rows = mrsMsg.RecordCount
        .ExtendLastCol = True
        .ColWidth(COL_���) = 500
        .ColWidth(COL_����) = 3000
        .ColWidth(COL_��ʾ) = 600
        .ColWidth(COL_����) = 8000
        .FocusRect = flexFocusNone
         For i = 0 To mrsMsg.RecordCount - 1
            .TextMatrix(i, COL_���) = (i + 1) & ":"
            .TextMatrix(i, COL_����) = mrsMsg!DrugName
            .Cell(flexcpData, i, COL_����) = mrsMsg!drugID & ""
            .TextMatrix(i, COL_��ʾ) = "[" & mrsMsg!severity & "] ,"
            .Cell(flexcpAlignment, i, COL_��ʾ) = flexAlignLeftBottom
            .TextMatrix(i, COL_����) = mrsMsg!message & " " & mrsMsg!Advice
            .Cell(flexcpForeColor, i, COL_����) = vbRed
            mrsMsg.MoveNext
         Next
         .ColAlignment(COL_��ʾ) = flexAlignCenterCenter
     End With
End Sub

Private Sub vsInfo_Click()
    Dim lngRow As Long, lngCol As Long
    With vsInfo
        lngRow = .MouseRow: lngCol = .MouseCol
        If lngRow < 0 Or lngCol < 0 Then Exit Sub
        If lngCol = COL_���� Then
            If .Cell(flexcpFontUnderline, lngRow, lngCol) And Val(.Cell(flexcpData, lngRow, COL_����)) <> 0 Then
                .Col = COL_���
                Call HZYY_DrugInstructions(Val(.Cell(flexcpData, lngRow, COL_����)))
            End If
        End If
    End With
End Sub

Private Sub vsInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long, lngCol As Long
    Dim arrTmp As Variant
    
    With vsInfo
        lngRow = .MouseRow: lngCol = .MouseCol
        .MousePointer = flexDefault
        If mstrFontUnderLine <> "" Then
            arrTmp = Split(mstrFontUnderLine, "|")
            If lngRow = arrTmp(0) And lngCol = arrTmp(1) Then .MousePointer = flexCustom: Exit Sub
            .Cell(flexcpFontUnderline, arrTmp(0), arrTmp(1)) = False
            mstrFontUnderLine = ""
        End If
        If lngCol < 0 Or lngRow < 0 Then Exit Sub
        If lngCol = COL_���� Then
            .Cell(flexcpFontUnderline, lngRow, lngCol) = True
            .MousePointer = flexCustom
            mstrFontUnderLine = lngRow & "|" & lngCol
        End If
    End With
End Sub
