VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMsgCallSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��������"
   ClientHeight    =   4245
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   10155
   Icon            =   "frmMsgCallSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   10155
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3075
      Left            =   105
      ScaleHeight     =   3075
      ScaleWidth      =   9405
      TabIndex        =   4
      Top             =   315
      Width           =   9405
      Begin VSFlex8Ctl.VSFlexGrid vsMsgList 
         Height          =   2670
         Left            =   180
         TabIndex        =   5
         Top             =   75
         Width           =   8865
         _cx             =   15637
         _cy             =   4710
         Appearance      =   2
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   99
         MouseIcon       =   "frmMsgCallSetup.frx":06EA
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483643
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
         OwnerDraw       =   1
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
         BackColorFrozen =   -2147483643
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.Frame fraBell 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   225
      Begin VB.Image imgBell 
         Height          =   240
         Left            =   0
         Picture         =   "frmMsgCallSetup.frx":0FC4
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   300
      Left            =   8430
      TabIndex        =   1
      Top             =   3540
      Width           =   1000
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   300
      Left            =   6945
      TabIndex        =   0
      Top             =   3540
      Width           =   1000
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   5010
      Top             =   -270
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imgFile 
      Height          =   240
      Left            =   1500
      Picture         =   "frmMsgCallSetup.frx":7816
      Top             =   105
      Width           =   240
   End
   Begin VB.Label lblW 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   1710
      TabIndex        =   6
      Top             =   60
      Width           =   90
   End
   Begin VB.Label lblDetail 
      Caption         =   "�ı���ʽ˵��������[����]��[סԺ��]�ֶΣ�ʹ����VBScript���ݵı��ʽ����Ϣ���ݽ��б༭���ֶ�����ʹ�÷�����""[]""�����ʾ��"
      Height          =   390
      Left            =   -45
      TabIndex        =   3
      Top             =   3690
      Width           =   6585
   End
End
Attribute VB_Name = "frmMsgCallSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum Col
    COL_�������� = 1
    COL_����
    COL_״̬
    COL_��ʾ��ʽ
    COL_��������
    COL_��������
End Enum

Private mobjVBA As Object
Private mobjScript As clsScript
Private mobjVoice As Object
Private mstr��Ϣ�� As String ' = "�¿���Ϣ,��ͣ��Ϣ,�·���Ϣ,������Ϣ,Σ��ֵ��Ϣ,��Һ�ܾ���Ϣ,����������Ϣ"
Private mint���� As Integer '0-����ҽ������վ��1��סԺҽ������վ��2��סԺ��ʿ����վ��3���ϰ�ҽ������վ
Private mrsPars As ADODB.Recordset

Public Function ShowMe(objFrm As Object, ByVal intType As Integer) As Boolean
    mint���� = intType
    Me.Show 1, objFrm
End Function

Private Sub InitMsgListTable()
'���ܣ���ʼ��������ݣ����ڴ�����Ի����ûָ�֮ǰ
    Dim strHead As String, i As Integer
    Dim arrHead As Variant, arrCol As Variant
 
    strHead = "��������,1200,1;����,540,4;״̬,540,4;��ʾ��ʽ,800,4;��������,5000,1;��������,900,4"
    arrHead = Split(strHead, ";")
    With vsMsgList
        .Clear
        .FixedRows = 1: .FixedCols = 1
        .Rows = 2: .Cols = .FixedCols + UBound(arrHead) + 1
        For i = 0 To UBound(arrHead)
            .FixedAlignment(.FixedCols + i) = 4
            arrCol = Split(arrHead(i), ",")
            .TextMatrix(0, .FixedCols + i) = arrCol(0)
             .Cell(flexcpText, 0, .FixedCols + i) = arrCol(0)
            
            If UBound(arrCol) > 0 Then
                .ColWidth(.FixedCols + i) = Val(arrCol(1))
                .ColAlignment(.FixedCols + i) = Val(arrCol(2))
                .ColHidden(.FixedCols + i) = False
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
        .ColWidth(0) = 0
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Long
    
    On Error GoTo errH
    
    With vsMsgList
        .Redraw = flexRDNone
        For i = .FixedRows To .Rows - 1
            If Not ChangePars(i) Then
                .Redraw = flexRDDirect
                Exit Sub
            End If
        Next
    End With
    
    mrsPars.Filter = "�޸�=1"
    If Not mrsPars.EOF Then
        For i = 1 To mrsPars.RecordCount
            Call zlDatabase.SetPara(mrsPars!������ & "", mrsPars!�ֲ���ֵ & "", glngSys, Val(mrsPars!ģ�� & ""))
            mrsPars.MoveNext
        Next
    End If
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim varTmp As Variant
    Dim strTmp As String
    Dim i As Long, lngģ�� As Long
    Dim objMsg As New clsCISMsg
    
    Call InitMsgListTable
    
    Call objMsg.InitRsMsgPar(mrsPars)
    
    If mint���� = 0 Then
        lngģ�� = p����ҽ��վ
    ElseIf mint���� = 1 Then
        lngģ�� = pסԺҽ��վ
    ElseIf mint���� = 2 Then
        lngģ�� = pסԺ��ʿվ
    ElseIf mint���� = 3 Then
        lngģ�� = pҽ������վ
    End If
    
    strTmp = objMsg.Get��Ϣ���(mint����)
    varTmp = Split(strTmp, ",")
    For i = 0 To UBound(varTmp)
        Call objMsg.AddDataToRsMsgPar(mrsPars, lngģ��, i + 1, varTmp(i) & "��������", varTmp(i))
    Next
    
    mrsPars.Filter = 0
    mrsPars.Sort = "���"
    mrsPars.MoveFirst
    With vsMsgList
        .Rows = UBound(varTmp) + 2
        For i = 1 To mrsPars.RecordCount
            .TextMatrix(i, COL_��������) = mrsPars!��������
            .TextMatrix(i, COL_״̬) = IIf(1 = Val(mrsPars!״̬ & ""), "����", "�ر�")
            .TextMatrix(i, COL_��ʾ��ʽ) = IIf(1 = Val(mrsPars!��ʾ��ʽ & ""), "��ʾ", "�ʶ�")
            .TextMatrix(i, COL_��������) = mrsPars!���� & ""
            .TextMatrix(i, COL_��������) = Val(mrsPars!���� & "")
            Set .Cell(flexcpPicture, i, COL_����) = imgBell.Picture
            .Cell(flexcpPictureAlignment, i, COL_����) = 4
            mrsPars.MoveNext
        Next
    End With
    
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    Me.Height = 3700
    picMain.BackColor = &H8000000A
    picMain.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 700
    lblW.Visible = False
    vsMsgList.Move 0, 0, picMain.ScaleWidth - 30, picMain.ScaleHeight - 30
    
    cmdCancel.Top = Me.ScaleHeight - 500
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 100
    
    cmdOK.Top = cmdCancel.Top
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 100
    
    lblDetail.Top = cmdCancel.Top
    lblDetail.Left = picMain.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjVBA = Nothing
    Set mobjScript = Nothing
    Set mobjVoice = Nothing
    Set mrsPars = Nothing
End Sub

Private Sub imgBell_Click()
'���ܣ�����
    Dim str���� As String
    Dim strTmp As String
    Dim lngRow As Long
    Dim objMsg As New clsCISMsg

    With vsMsgList
        If Not .Col = COL_���� Then Exit Sub
        lngRow = Val(fraBell.Tag)
        str���� = .TextMatrix(lngRow, COL_��������)
        If .TextMatrix(lngRow, COL_��ʾ��ʽ) = "��ʾ" Then
            If str���� = "" Then
                MsgBox "��ѡ��һ����Ƶ�ļ�(*.wav)��", vbInformation, gstrSysName
                Exit Sub
            Else
                If UCase(Right(str����, 4)) <> ".WAV" Then
                    MsgBox "��ѡ��һ����ȷ��ʽ����Ƶ�ļ�(*.wav)��", vbInformation, gstrSysName
                    Exit Sub
                ElseIf Not Dir(str����) <> "" Then
                    MsgBox "δ���ļ�[" & str���� & "]�����顣", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
            Call sndPlaySound(str����, 1)
        Else
            If str���� = "" Then
                MsgBox "�붨��һ���ʶ��ı���", vbInformation, gstrSysName
            End If
            
            If mobjVoice Is Nothing Then
                Set mobjVoice = CreateObject("SAPI.SpVoice")
                Call objMsg.CreateScript(mobjVBA, mobjScript)
            End If
            
            strTmp = Check�ı�(str����)
            If strTmp <> "" Then
                MsgBox .TextMatrix(lngRow, COL_��������) & "�ı���ʽ���δͨ������" & strTmp & "��", vbInformation, gstrSysName
                Exit Sub
            End If
                
            str���� = Get�����ı�(str����)
            mobjVoice.Speak str����, 1
        End If
    End With
End Sub

Private Sub vsMsgList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow = -1 Then Exit Sub
    With vsMsgList
        If NewCol = COL_�������� Then
            .Editable = flexEDKbdMouse
            If .TextMatrix(NewRow, COL_��ʾ��ʽ) = "��ʾ" Then
                .ComboList = "..."
                Set .CellButtonPicture = imgFile.Picture
            Else
                .ComboList = ""
            End If
            .FocusRect = flexFocusLight
        ElseIf NewCol = COL_�������� Then
            .Editable = flexEDKbdMouse
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .Editable = flexEDNone
            .FocusRect = flexFocusNone
            .ComboList = ""
        End If
    End With
End Sub

Private Sub vsMsgList_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
'���ܣ��༭���
    Dim strFileDir As String
    On Error GoTo errH
    If Col = COL_�������� And vsMsgList.TextMatrix(Row, COL_��ʾ��ʽ) = "��ʾ" Then
        vsMsgList.Redraw = flexRDNone
        With dlgFile
            .CancelError = False
            .Flags = cdlOFNHideReadOnly
            .Filter = "(*.wav)|*.wav"
            .FilterIndex = 2
            .ShowOpen
            strFileDir = .FileName
            If strFileDir = "" Then Exit Sub
        End With
        vsMsgList.TextMatrix(Row, COL_��������) = strFileDir
        vsMsgList.Redraw = flexRDDirect
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
 
Private Sub vsMsgList_Click()
    vsMsgList.Redraw = flexRDNone
    Call imgBell_Click
    vsMsgList.Redraw = flexRDDirect
End Sub

Private Sub vsMsgList_DblClick()
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strTmp As String
    
    With vsMsgList
        lngRow = .Row
        lngCol = .Col
        If lngRow >= .FixedRows Then
            If lngCol = COL_��ʾ��ʽ Then
                If .TextMatrix(lngRow, lngCol) = "��ʾ" Then
                    .TextMatrix(lngRow, lngCol) = "�ʶ�"
                Else
                    .TextMatrix(lngRow, lngCol) = "��ʾ"
                End If
                .TextMatrix(lngRow, COL_��������) = ""
            ElseIf lngCol = COL_״̬ Then
                If .TextMatrix(lngRow, lngCol) = "����" Then
                    .TextMatrix(lngRow, lngCol) = "�ر�"
                Else
                    .TextMatrix(lngRow, lngCol) = "����"
                End If
            End If
        End If
    End With
End Sub

Private Sub vsMsgList_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
'      2.Cell��GridLine�������������ڶ��Ǵӵ�1���߿�ʼ
'      3.Cell��Border�������Ǵӵ�2���߿�ʼ,�����Ǵӵ�1���߿�ʼ
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim lngColor As Long, lngclrg, k As Long, n As Long
    Dim r1 As Integer, g1 As Integer, b1 As Integer
    Dim r2 As Integer, g2 As Integer, b2 As Integer
    Dim rg As Integer, gg As Integer, bg As Integer
    
    Dim lngFontW As Long
    Dim lng��ID As Long
    Dim strContent As String
    
    Dim vRect As RECT, vRect1 As RECT, vRect2 As RECT
    With vsMsgList
        '�����б���ɫ����
        If .FixedRows - 1 = Row Then
            '��ȡ���ο�
            vRect.Left = Left
            vRect.Top = Top
            vRect.Right = Right - 1
            vRect.Bottom = Bottom - 1
            'draw frame
            lngColor = SetBkColor(hDC, 0)
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, k
    
            ' get colors
            r1 = 250: g1 = 250: b1 = 250   '������ʼ
            r2 = 229: g2 = 229: b2 = 229   '������ֹ
            ' show color
            vRect2 = vRect
            vRect2.Bottom = vRect.Bottom - (vRect.Bottom - vRect.Top) / 2
            vRect1 = vRect2
    
            For k = vRect2.Top To vRect2.Bottom
                rg = r1 + (k - vRect2.Top) * (r2 - r1) / (vRect2.Bottom - vRect2.Top)
                gg = g1 + (k - vRect2.Top) * (g2 - g1) / (vRect2.Bottom - vRect2.Top)
                bg = b1 + (k - vRect2.Top) * (b2 - b1) / (vRect2.Bottom - vRect2.Top)
                lngclrg = RGB(rg, gg, bg)
                SetBkColor hDC, lngclrg
                vRect1.Top = k
                ExtTextOut hDC, vRect1.Left, vRect1.Top, ETO_OPAQUE, vRect1, " ", 1, k
            Next
            ' get colors
            r1 = 229: g1 = 229: b1 = 229   '������ʼ
            r2 = 250: g2 = 250: b2 = 250   '������ֹ
            ' show color
            vRect2 = vRect
            vRect2.Top = vRect.Top + (vRect.Bottom - vRect.Top) / 2
            vRect1 = vRect2
            For k = vRect2.Top To vRect2.Bottom
                rg = r1 + (k - vRect2.Top) * (r2 - r1) / (vRect2.Bottom - vRect2.Top)
                gg = g1 + (k - vRect2.Top) * (g2 - g1) / (vRect2.Bottom - vRect2.Top)
                bg = b1 + (k - vRect2.Top) * (b2 - b1) / (vRect2.Bottom - vRect2.Top)
                lngclrg = RGB(rg, gg, bg)
                SetBkColor hDC, lngclrg
                vRect1.Top = k
                ExtTextOut hDC, vRect1.Left, vRect1.Top, ETO_OPAQUE, vRect1, " ", 1, k
            Next
    
            SetBkColor hDC, lngColor
            '����Ԫ������浽��������
            strContent = .Cell(flexcpText, Row, Col)
            lblW.Caption = strContent: lblW.AutoSize = True
            vRect1.Top = vRect.Top + (vRect.Bottom - vRect.Top) / 2 - (lblW.Height / 2) / Screen.TwipsPerPixelY

            vRect1.Left = vRect.Left + (vRect.Right - vRect.Left) / 2 - (lblW.Width / 2) / Screen.TwipsPerPixelX
    
            TextOut hDC, vRect1.Left, vRect1.Top, strContent, LenB(StrConv(strContent, vbFromUnicode))
        End If
    End With
End Sub

Private Sub vsMsgList_KeyPress(KeyAscii As Integer)
    With vsMsgList
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call EnterNextCell(.Row, .Col)
        Else
             If .Col = COL_�������� Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsMsgList_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" 'ʹ��ť״̬��������״̬
                End If
            End If
        End If
    End With
End Sub

Private Sub vsMsgList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call EnterNextCell(Row, Col)
        Exit Sub
    End If
    If Col = COL_�������� Then
        If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0: Exit Sub
        End If
    ElseIf Col = COL_�������� And vsMsgList.TextMatrix(Row, COL_��ʾ��ʽ) = "��ʾ" Then
        KeyAscii = 0
        Exit Sub
    End If
End Sub

Private Sub vsMsgList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngRow As Long
    Dim strTag As String
    Dim lngCol As Long
 
    lngRow = vsMsgList.MouseRow
    If Button = 0 And lngRow > 0 Then
        If vsMsgList.MouseCol = COL_���� Then
            If Val(fraBell.Tag) <> lngRow Then
                With vsMsgList
                    fraBell.Visible = False
                    fraBell.Tag = lngRow
                    If lngRow = .Row Then
                        fraBell.BackColor = .BackColorSel
                    Else
                        fraBell.BackColor = .BackColor
                    End If
                    fraBell.Height = .RowHeight(lngRow) - 20
                    If fraBell.Height > 250 Then fraBell.Height = 250
                    
                    fraBell.Top = .Top + .RowPos(lngRow) + .RowHeight(lngRow) - fraBell.Height + 10
                    If fraBell.Top + fraBell.Height > .Top + .Height Then Exit Sub
                    
                    fraBell.Left = .Left + .ColPos(COL_����) + (.ColWidth(COL_����) - fraBell.Width) / 2
                    
                    fraBell.Visible = True
                End With
            End If
        Else
            If fraBell.Visible Then
                fraBell.Tag = ""
                fraBell.Visible = False
            End If
        End If
    End If
    
    With vsMsgList
        lngRow = .MouseRow
        lngCol = .MouseCol
        strTag = lngRow & "<T>" & lngCol
        If lngRow > 0 And .Tag <> strTag Then
            If lngCol = COL_��ʾ��ʽ Or lngCol = COL_״̬ Then
                Call ClearListSel
                .Cell(flexcpForeColor, lngRow, lngCol) = .ForeColorSel
                .Cell(flexcpBackColor, lngRow, lngCol) = vbWhite
                .CellBorderRange lngRow, lngCol, lngRow, lngCol, 1, 1, 1, 1, 1, 1, 1
                .Tag = strTag
                .ToolTipText = .Cell(flexcpText, lngRow, lngCol)
            Else
                Call ClearListSel
            End If
        End If
        If lngRow > 0 And lngCol > 0 Then .ToolTipText = .Cell(flexcpText, lngRow, lngCol)
        
        If lngCol = COL_���� Then
            .MousePointer = 99
        Else
            .MousePointer = 0
        End If
    End With
End Sub

Private Sub ClearListSel()
    Dim lngCol As Long
    Dim lngRow As Long
    
    On Error Resume Next
    With vsMsgList
        If .Tag <> "" Then
            lngRow = Split(.Tag, "<T>")(0)
            lngCol = Split(.Tag, "<T>")(1)
            .Row = .FixedRows - 1
            .Cell(flexcpForeColor, lngRow, lngCol) = .ForeColor
            .Cell(flexcpBackColor, lngRow, lngCol) = .BackColor
            .CellBorderRange lngRow, lngCol, lngRow, lngCol, 0, 0, 0, 0, 0, 0, 0
            .ToolTipText = ""
            .Tag = ""
        End If
    End With
End Sub

Private Sub EnterNextCell(ByVal lngRow As Long, ByVal lngCol As Long)
'���ܣ���λ����һ����������ĵ�Ԫ��
    Dim i As Long, j As Long
    Dim blnDo As Boolean
    
    With vsMsgList
        '����һ��Ԫ��ʼѭ������
        For i = lngRow To .Rows - 1
            For j = IIf(i = lngRow, lngCol + 1, COL_��ʾ��ʽ) To .Cols - 1
                If j >= COL_�������� Then
                    Exit For
                End If
            Next
            If j <= .Cols - 1 Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
            blnDo = True
        End If
        .ShowCell .Row, .Col
        If Not blnDo Then Call zlCommFun.PressKey(vbKeyTab)
    End With
End Sub

Private Function Check�ı�(ByVal strText As String) As String
'���ܣ����ҽ�������Ƿ���ȷ
'���أ�������Ϣ
'      strPreview=Ԥ��ҽ������Ч��
    Dim intLeft As Integer, intRight As Integer
    Dim strTmp As String, strPar As String
    Dim strMsg As String, i As Long
    Dim objVBA As Object, strEval As String
    Dim objScript As New clsScript
    
    If Trim(strText) = "" Then Exit Function
    If zlCommFun.ActualLen(strText) > 100 Then
        strMsg = "�ı���������̫����ֻ����100���ַ���50�����֡�"
        GoTo EndLine
    End If
        
    '���������
    For i = 1 To Len(strText)
        If Mid(strText, i, 1) = "[" Then
            intLeft = intLeft + 1
        ElseIf Mid(strText, i, 1) = "]" Then
            intRight = intRight + 1
            If intLeft <> intRight Then
                strMsg = """[""��""]""���Ų���ԡ�"
                GoTo EndLine
            End If
        End If
    Next
    If intLeft = 0 And intRight = 0 Then Exit Function
    If intLeft <> intRight Then
        strMsg = """[""��""]""���Ų���ԡ�"
        GoTo EndLine
    End If
    
    '����ֶ�����
    strTmp = strText
    Do While InStr(strTmp, "[") > 0
        strTmp = Mid(strTmp, InStr(strTmp, "[") + 1)
        strPar = Trim(Left(strTmp, InStr(strTmp, "]") - 1))
                        
        If strPar = "" Then
            strMsg = """[]""����֮��û����д�ֶ�����"
            GoTo EndLine
        End If
        If "[����]" <> "[" & strPar & "]" And "[סԺ��]" = "[" & strPar & "]" Then
            strMsg = "ʹ���˲����ڵ�""[" & strPar & "]""�ֶΡ�"
            GoTo EndLine
        End If
    Loop
    
    'ִ�в���
    On Error Resume Next
    Set objVBA = CreateObject("ScriptControl")
    If objVBA Is Nothing Then
        strMsg = "Microsoft Script Control δ��ȷ��װ(msscript.ocx)������ִ�м�顣�����°�װ�ͻ��˳���"
        GoTo EndLine
    End If
    err.Clear: On Error GoTo 0
    objVBA.Language = "VBScript"
    objVBA.AddObject "clsScript", objScript, True
    strEval = Replace(strText, "[", """")
    strEval = Replace(strEval, "]", """")
    On Error Resume Next
    Call objVBA.Eval(strEval)
    If objVBA.Error.Number <> 0 Then
        strMsg = objVBA.Error.Description
        objVBA.Error.Clear
    End If
EndLine:
    Check�ı� = strMsg
End Function

Private Function CheckFile(ByVal strFile As String) As String
'���ܣ���Ƶ�ļ����
    Dim strMsg As String
    
    If UCase(Right(strFile, 4)) <> ".WAV" Then
        strMsg = "��ѡ��һ����ȷ��ʽ����Ƶ�ļ�(*.wav)��"
        MsgBox "��ѡ��һ����ȷ��ʽ����Ƶ�ļ�(*.wav)��", vbInformation, gstrSysName
    ElseIf Not Dir(strFile) <> "" Then
        strMsg = "δ���ļ�[" & strFile & "]�����顣"
    End If
End Function

Private Function Get�����ı�(ByVal strText As String) As String
'���ܣ���ȡ������ı�
    Dim str���� As String, strסԺ�� As String
    Dim strVal As String
    
    str���� = "1��"
    strסԺ�� = "201608010008��"
    strVal = strText
    strVal = Replace(strVal, "[����]", """" & str���� & """")
    strVal = Replace(strVal, "[סԺ��]", """" & strסԺ�� & """")
    
    On Error Resume Next
    strVal = mobjVBA.Eval(strVal)
    If mobjVBA.Error.Number <> 0 Then
        err.Clear
        strVal = ""
    End If
    Get�����ı� = strVal
End Function

Private Function ChangePars(ByVal lngRow As Long) As Boolean
'���ܣ������任
    Dim strTmp As String
    
    With vsMsgList
        If Not IsNumeric(.TextMatrix(lngRow, COL_��������)) Then
            MsgBox .TextMatrix(lngRow, COL_��������) & "������������Ϊ���֡�", vbInformation, gstrSysName
            Exit Function
        End If
        
        If Val(.TextMatrix(lngRow, COL_��������)) > 6 Then
            MsgBox .TextMatrix(lngRow, COL_��������) & "�����������ֻ����Ϊ6�Ρ�", vbInformation, gstrSysName
            Exit Function
        End If
        
        If .TextMatrix(lngRow, COL_��ʾ��ʽ) = "��ʾ" Then
            If .TextMatrix(lngRow, COL_��������) = "" Then
                MsgBox .TextMatrix(lngRow, COL_��������) & "δ������ʾ��Ƶ�ļ���", vbInformation, gstrSysName
                Exit Function
            Else
                strTmp = CheckFile(.TextMatrix(lngRow, COL_��������))
                If strTmp <> "" Then
                    MsgBox .TextMatrix(lngRow, COL_��������) & "�ļ����ã���" & strTmp & "��", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Else
            If .TextMatrix(lngRow, COL_��������) = "" Then
                MsgBox .TextMatrix(lngRow, COL_��������) & "δ�������ı���ʽ��", vbInformation, gstrSysName
                Exit Function
            Else
                strTmp = Check�ı�(.TextMatrix(lngRow, COL_��������))
                If strTmp <> "" Then
                    MsgBox .TextMatrix(lngRow, COL_��������) & "�ı���ʽ���δͨ������" & strTmp & "��", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        
        mrsPars.Filter = "��������='" & .TextMatrix(lngRow, COL_��������) & "'"
        mrsPars!״̬ = IIf("����" = .TextMatrix(lngRow, COL_״̬), 1, 0)
        mrsPars!��ʾ��ʽ = IIf("��ʾ" = .TextMatrix(lngRow, COL_��ʾ��ʽ), 1, 0)
        mrsPars!���� = .TextMatrix(lngRow, COL_��������)
        mrsPars!���� = Val(.TextMatrix(lngRow, COL_��������))
        mrsPars!�ֲ���ֵ = mrsPars!״̬ & "<sTab>" & mrsPars!��ʾ��ʽ & "<sTab>" & mrsPars!���� & "<sTab>" & mrsPars!����
        
        If mrsPars!�ֲ���ֵ <> mrsPars!ԭ����ֵ Then
            mrsPars!�޸� = 1
        End If
        
        mrsPars.Update
    End With
    ChangePars = True
End Function
