VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmClinicWorkTimeOther 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ʱ��θ�������"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   3630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdȡ�� 
      Caption         =   "�˳�(&X)"
      Height          =   350
      Left            =   2460
      TabIndex        =   6
      Top             =   4590
      Width           =   1100
   End
   Begin VB.OptionButton optʱ�� 
      Caption         =   "�ֶ�ʱ����"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   4
      Top             =   540
      Width           =   1380
   End
   Begin VB.OptionButton optʱ�� 
      Caption         =   "ƽ��ʱ����"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   3
      Top             =   180
      Value           =   -1  'True
      Width           =   1395
   End
   Begin VB.TextBox txtTimeOut 
      Height          =   300
      Left            =   1500
      TabIndex        =   1
      Text            =   "10"
      Top             =   135
      Width           =   450
   End
   Begin VB.CommandButton cmdRecalc 
      Caption         =   "��������(&F)"
      Height          =   350
      Left            =   990
      TabIndex        =   0
      ToolTipText     =   "������¼���ʱ��"
      Top             =   4590
      Width           =   1260
   End
   Begin MSComCtl2.UpDown udTime 
      Height          =   300
      Left            =   1980
      TabIndex        =   2
      Top             =   135
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   529
      _Version        =   393216
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txtTimeOut"
      BuddyDispid     =   196611
      OrigLeft        =   2025
      OrigTop         =   105
      OrigRight       =   2280
      OrigBottom      =   450
      Max             =   1440
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VSFlex8Ctl.VSFlexGrid vsTime 
      Height          =   3600
      Left            =   30
      TabIndex        =   7
      Top             =   840
      Width           =   3555
      _cx             =   6271
      _cy             =   6350
      Appearance      =   0
      BorderStyle     =   1
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
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmClinicWorkTimeOther.frx":0000
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
   Begin VB.Label lbl�� 
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   180
      Left            =   2235
      TabIndex        =   5
      Top             =   195
      Width           =   180
   End
End
Attribute VB_Name = "frmClinicWorkTimeOther"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type TYTime
    �ϰ�ʱ�� As String
    �°�ʱ�� As String
End Type
Private Enum mʱ������
    mƽ��ʱ���� = 0
    m�ֶ�ʱ���� = 1
End Enum
Private Type TyWorkTime
    ���� As TYTime
    ���� As TYTime
End Type
Private mWorkTime As TyWorkTime
Private mintȱʡ��� As Integer
Private mdtStartTime As Date
Private mdtEndTime As Date
Private mstr��Ϣʱ�� As String
Private mblnOk As Boolean

Private mvarTimes As Variant
'VarTimes
'       "ʱ����":Array("ʱ����",10)
'       "�ֶμ��":Array("�ֶμ��",ʱ��1(��:8:00��9:00),2;ʱ��2,���2;....)
Public Function ShowMe(ByVal frmMain As Object, ByVal intȱʡ��� As Integer, _
    ByVal datStartTime As Date, ByVal datEndTime As Date, ByVal str��Ϣʱ�� As String, ByRef varTimes As Variant) As Boolean
    '����:�������
    '��Σ�
    '   intȱʡ��� - ȱʡ�İ�"ƽ��ʱ����"�ļ��������
    '   datStartTime - ��ʼʱ��
    '   datEndTime - ��ֹʱ��
    '   str��Ϣʱ�� - ��ʽ����ʼʱ��1����ֹʱ��1; ��ʼʱ��2����ֹʱ��2;��.��
    '                 ��ʼʱ�����ֹʱ���ʽΪ: HH24:MM.���磺12:00��14:00;17:30��18:00
    mintȱʡ��� = intȱʡ���
    mdtStartTime = datStartTime: mdtEndTime = datEndTime: mstr��Ϣʱ�� = str��Ϣʱ��
    mblnOk = False: mvarTimes = Empty
    
    Err = 0: On Error Resume Next
    Me.Show 1, frmMain
    varTimes = mvarTimes
    ShowMe = mblnOk
End Function

Private Sub cmdRecalc_Click()
    Dim strTemp As String, i As Integer

    Err = 0: On Error GoTo errHandler
    If Val(txtTimeOut.Text) > 60 Or Val(txtTimeOut.Text) < 0 Then
        MsgBox "ƽ��ʱ�������ܴ���60���ӻ�С��0���ӣ�", vbInformation, gstrSysName
        txtTimeOut.Text = mintȱʡ���
        Exit Sub
    End If
    
    If optʱ��(mƽ��ʱ����).Value Then
        mvarTimes = Array("ʱ����", Val(txtTimeOut.Text))
    Else
        With vsTime
            strTemp = ""
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, .ColIndex("ʱ��̶�"))) <> "" _
                    And Val(.TextMatrix(i, .ColIndex("ʱ����"))) > 0 Then
                    strTemp = strTemp & ";" & .Cell(flexcpData, i, .ColIndex("ʱ��̶�"))
                    strTemp = strTemp & "," & Val(.TextMatrix(i, .ColIndex("ʱ����")))
                End If
            Next
        End With
        If strTemp = "" Then
            MsgBox "δ����ʱ���������飡", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        Else
            strTemp = Mid(strTemp, 2)
            mvarTimes = Array("�ֶμ��", strTemp)
        End If
    End If
    mblnOk = True: Unload Me
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Err = 0: On Error GoTo errHandler
    Call InitData
    optʱ��(m�ֶ�ʱ����).Value = True
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    cmdRecalc.Top = ScaleHeight - cmdRecalc.Height - 100
    cmdȡ��.Top = cmdRecalc.Top
    cmdȡ��.Left = ScaleWidth - cmdȡ��.Width - 50
    
    cmdRecalc.Left = cmdȡ��.Left - cmdRecalc.Width - 50
    With vsTime
        .Left = Me.ScaleLeft
        .Height = cmdRecalc.Top - .Top - 50
        .Width = Me.ScaleWidth
    End With
End Sub

Private Sub InitData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����Ϣ
    '����:���˺�
    '����:2012-07-10 17:25:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, varData As Variant
    Dim dtCurStart As Date, dtCurEnd As Date, lngRow As Integer
    Dim varTimes As Variant, dtStart As Date, dtEnd As Date
    Dim i As Integer
    
    On Error GoTo errHandler
    With vsTime
        .Clear 1
        .Rows = 1
        .Editable = flexEDKbdMouse
        lngRow = 1
        dtCurStart = mdtStartTime
        If mstr��Ϣʱ�� = "" Then
            dtCurEnd = DateAdd("h", 1, dtCurStart)
            Do While DateDiff("n", dtCurEnd, mdtEndTime) >= 0
                If DateDiff("n", dtCurEnd, mdtEndTime) < 0 Then dtCurEnd = mdtEndTime
                .Rows = .Rows + 1
                .TextMatrix(lngRow, .ColIndex("ʱ��̶�")) = Format(dtCurStart, "HH:MM") & "��" & Format(dtCurEnd, "HH:MM")
                .Cell(flexcpData, lngRow, .ColIndex("ʱ��̶�")) = dtCurStart & "��" & dtCurEnd
                .TextMatrix(lngRow, .ColIndex("ʱ����")) = txtTimeOut.Text
                lngRow = lngRow + 1
                dtCurStart = dtCurEnd: dtCurEnd = DateAdd("h", 1, dtCurStart)
            Loop
            If DateDiff("n", dtCurStart, mdtEndTime) > 0 Then
                .Rows = .Rows + 1
                .TextMatrix(lngRow, .ColIndex("ʱ��̶�")) = Format(dtCurStart, "HH:MM") & "��" & Format(mdtEndTime, "HH:MM")
                .Cell(flexcpData, lngRow, .ColIndex("ʱ��̶�")) = dtCurStart & "��" & mdtEndTime
                .TextMatrix(lngRow, .ColIndex("ʱ����")) = txtTimeOut.Text
            End If
        Else
            varTimes = Split(mstr��Ϣʱ��, ";")
            For i = 0 To UBound(varTimes)
                '�����Ϣʱ�εĿ�ʼʱ��С�ڵ�ǰʱ�εĿ�ʼʱ�䣬���ʾ�ǵڶ��죬��Ϣʱ�εĿ�ʼʱ�����ֹʱ�䶼Ҫ��һ��
                dtStart = CDate(Format(mdtStartTime, "yyyy-mm-dd ") & Split(varTimes(i), "-")(0))
                dtEnd = CDate(Format(mdtStartTime, "yyyy-mm-dd ") & Split(varTimes(i), "-")(1))
                If DateDiff("n", dtStart, dtCurStart) > 0 Then dtStart = DateAdd("d", 1, dtStart): dtEnd = DateAdd("d", 1, dtEnd)
                '��Ϣʱ�ε���ֹʱ��С����Ϣʱ�εĿ�ʼʱ�䣬����Ϣʱ�ε���ֹʱ���һ��
                If DateDiff("n", dtEnd, dtStart) > 0 Then dtEnd = DateAdd("d", 1, dtEnd)
                
                dtCurEnd = DateAdd("h", 1, dtCurStart)
                Do While DateDiff("n", dtCurEnd, dtStart) >= 0
                    If DateDiff("n", dtCurEnd, dtStart) < 0 Then dtCurEnd = dtStart
                    .Rows = .Rows + 1
                    .TextMatrix(lngRow, .ColIndex("ʱ��̶�")) = Format(dtCurStart, "HH:MM") & "��" & Format(dtCurEnd, "HH:MM")
                    .Cell(flexcpData, lngRow, .ColIndex("ʱ��̶�")) = dtCurStart & "��" & dtCurEnd
                    .TextMatrix(lngRow, .ColIndex("ʱ����")) = txtTimeOut.Text
                    lngRow = lngRow + 1
                    dtCurStart = dtCurEnd: dtCurEnd = DateAdd("h", 1, dtCurStart)
                Loop
                If DateDiff("n", dtCurStart, dtStart) > 0 Then
                    .Rows = .Rows + 1
                    .TextMatrix(lngRow, .ColIndex("ʱ��̶�")) = Format(dtCurStart, "HH:MM") & "��" & Format(dtStart, "HH:MM")
                    .Cell(flexcpData, lngRow, .ColIndex("ʱ��̶�")) = dtCurStart & "��" & dtStart
                    .TextMatrix(lngRow, .ColIndex("ʱ����")) = txtTimeOut.Text
                End If
                dtCurStart = dtEnd
            Next
            dtStart = mdtEndTime
            Do While DateDiff("n", dtCurEnd, mdtEndTime) >= 0
                If DateDiff("n", dtCurEnd, mdtEndTime) < 0 Then dtCurEnd = mdtEndTime
                .Rows = .Rows + 1
                .TextMatrix(lngRow, .ColIndex("ʱ��̶�")) = Format(dtCurStart, "HH:MM") & "��" & Format(dtCurEnd, "HH:MM")
                .Cell(flexcpData, lngRow, .ColIndex("ʱ��̶�")) = dtCurStart & "��" & dtCurEnd
                .TextMatrix(lngRow, .ColIndex("ʱ����")) = txtTimeOut.Text
                lngRow = lngRow + 1
                dtCurStart = dtCurEnd: dtCurEnd = DateAdd("h", 1, dtCurStart)
            Loop
            If DateDiff("n", dtCurStart, mdtEndTime) > 0 Then
                .Rows = .Rows + 1
                .TextMatrix(lngRow, .ColIndex("ʱ��̶�")) = Format(dtCurStart, "HH:MM") & "��" & Format(mdtEndTime, "HH:MM")
                .Cell(flexcpData, lngRow, .ColIndex("ʱ��̶�")) = dtCurStart & "��" & mdtEndTime
                .TextMatrix(lngRow, .ColIndex("ʱ����")) = txtTimeOut.Text
            End If
        End If
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub optʱ��_Click(index As Integer)
    Err = 0: On Error GoTo errHandler
   Call SetControlEnable
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsTime_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Err = 0: On Error GoTo errHandler
     With vsTime
        If Val(.Cell(flexcpText, Row, .ColIndex("ʱ����"))) > 60 Or Val(.Cell(flexcpText, Row, .ColIndex("ʱ����"))) < 0 Then
            MsgBox "ʱ�������ܴ���60���ӻ�С��0���ӣ�", vbInformation, gstrSysName
           .Cell(flexcpText, Row, .ColIndex("ʱ����")) = ""
        End If
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsTime_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsTime
        Select Case Col
        Case .ColIndex("ʱ����")
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub SetControlEnable()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ��Ƿ����
    '����:����
    '����:2012-07-11 09:52:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If optʱ��(mƽ��ʱ����).Value = True Then
        vsTime.Enabled = False
        txtTimeOut.Enabled = True
        udTime.Enabled = True
    ElseIf optʱ��(m�ֶ�ʱ����).Value = True Then
        vsTime.Enabled = True
        txtTimeOut.Enabled = False
        udTime.Enabled = False
    End If
End Sub

Private Sub vsTime_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then KeyAscii = 0
End Sub
