VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRegistPlanTimeOther 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "ʱ��θ�������"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdȡ�� 
      Caption         =   "�˳�(&X)"
      Height          =   350
      Left            =   2460
      TabIndex        =   7
      Top             =   4590
      Width           =   1100
   End
   Begin VB.OptionButton optʱ�� 
      Caption         =   "�ֶ�ʱ����"
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   5
      Top             =   540
      Width           =   2850
   End
   Begin VB.OptionButton optʱ�� 
      Caption         =   "ƽ��ʱ����"
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   4
      Top             =   180
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.TextBox txtTimeOut 
      Height          =   300
      Left            =   1530
      TabIndex        =   2
      Text            =   "10"
      Top             =   135
      Width           =   450
   End
   Begin VB.CommandButton cmdRecalc 
      Caption         =   "��������(&F)"
      Height          =   350
      Left            =   990
      TabIndex        =   1
      ToolTipText     =   "������¼���ʱ��"
      Top             =   4590
      Width           =   1260
   End
   Begin VSFlex8Ctl.VSFlexGrid vsTime 
      Height          =   3600
      Left            =   60
      TabIndex        =   0
      Top             =   900
      Width           =   3705
      _cx             =   6535
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
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483634
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
      FormatString    =   $"frmRegistPlanTimeOther.frx":0000
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
   Begin MSComCtl2.UpDown udTime 
      Height          =   300
      Left            =   1980
      TabIndex        =   3
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
   Begin VB.Label lbl�� 
      AutoSize        =   -1  'True
      Caption         =   "��"
      Height          =   180
      Left            =   2265
      TabIndex        =   6
      Top             =   195
      Width           =   180
   End
End
Attribute VB_Name = "frmRegistPlanTimeOther"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type TYTime
    �ϰ�ʱ�� As String
    �°�ʱ�� As String
End Type
Private Type TyWorkTime
    ���� As TYTime
    ���� As TYTime
End Type
Private Enum mʱ������
    mƽ��ʱ���� = 0
    m�ֶ�ʱ���� = 1
End Enum
Private mWorkTime As TyWorkTime
Private mrsʱ�� As ADODB.Recordset
'VarTiems
'       "ʱ����"
'       "�ֶμ��":ʱ��(��:8:00��9:00),2;ʱ��2,���;....
Public Event zlRefreshCon(ByVal varTimes As Variant)
Public Function zlShowMe(ByVal frmMain As Object, ByVal str���� As String, ByVal intȱʡ��� As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������(Ŀǰ��ʱӦ��)
    '����:���˺�
    '����:2012-07-10 18:35:09
    '˵��:
    '   21001   22001   �����������У��    IDCardCheck
    '---------------------------------------------------------------------------------------------------------------------------------------------
    
    Call zlInitVar(str����, intȱʡ���)
     Me.Show 1, frmMain
End Function
Public Sub zlInitVar(ByVal str���� As String, ByVal intȱʡ��� As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����ر���ֵ
    '���:str����-ʱ�䰲��,����:����;����
    '        intȱʡ���-ȱʡ��ʱ����
    '����:���˺�
    '����:2012-07-10 17:21:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str��ʼʱ�� As String, str����ʱ�� As String
    Dim i As Long, lngRow As Long
    Dim bln���� As Boolean
    Dim dtDate As Date
    Dim blnȫ�� As Boolean
    If mrsʱ�� Is Nothing Then InitData
    mrsʱ��.Filter = "ʱ���='" & str���� & "'"
    txtTimeOut.Text = intȱʡ���
    If Not mrsʱ��.EOF Then
        str��ʼʱ�� = mrsʱ��!��ʼʱ��
        str����ʱ�� = mrsʱ��!��ֹʱ��
        If mrsʱ��!ʱ��� = "ȫ��" Then
            bln���� = str��ʼʱ�� >= str����ʱ��
            str����ʱ�� = IIf(bln���� = False, str����ʱ��, mWorkTime.����.�°�ʱ��)
        End If
    End If
    With vsTime
        .Clear 1
        .Rows = 2
        If str��ʼʱ�� = "" Then str��ʼʱ�� = mWorkTime.����.�ϰ�ʱ��
        If str����ʱ�� = "" Then str����ʱ�� = mWorkTime.����.�°�ʱ��
        If str��ʼʱ�� > str����ʱ�� Then
            str��ʼʱ�� = "2000-01-01 " & str��ʼʱ��
            str����ʱ�� = "2000-01-02 " & str����ʱ��
        End If
        lngRow = 1
        Do While True
            dtDate = Format(CDate(str��ʼʱ��), "yyyy-mm-dd HH:00:00")
            dtDate = dtDate + 1 / 24
            If dtDate > CDate(str����ʱ��) Then dtDate = CDate(str����ʱ��)
            .TextMatrix(lngRow, .ColIndex("ʱ��̶�")) = Format(str��ʼʱ��, "HH:MM") & "��" & Format(dtDate, "HH:MM")
            If str��ʼʱ�� < CDate(mWorkTime.����.�°�ʱ��) Or str��ʼʱ�� >= CDate(mWorkTime.����.�ϰ�ʱ��) Then
                .TextMatrix(lngRow, .ColIndex("ʱ����")) = txtTimeOut.Text
            End If
            If dtDate >= str����ʱ�� Then Exit Do
            str��ʼʱ�� = Format(dtDate, "yyyy-mm-dd HH:MM:SS")
            .Rows = .Rows + 1
            lngRow = lngRow + 1
        Loop
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub cmdRecalc_Click()
    Dim cllTime As New Collection
    Dim strTemp As String, i As Long
    cllTime.Add "", "ʱ����"
    cllTime.Add "", "�ֶμ��"
    If optʱ��(0).Value Then
        cllTime.Remove "ʱ����"
        cllTime.Add txtTimeOut.Text, "ʱ����"
        RaiseEvent zlRefreshCon(cllTime)
        Exit Sub
    End If
    With vsTime
        strTemp = ""
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("ʱ��̶�"))) <> "" _
                And Val(.TextMatrix(i, .ColIndex("ʱ����"))) >= 0 Then
                strTemp = strTemp & ";" & Trim(.TextMatrix(i, .ColIndex("ʱ��̶�")))
                strTemp = strTemp & "," & Val(.TextMatrix(i, .ColIndex("ʱ����")))
            End If
        Next
    End With
    If strTemp = "" Then
        MsgBox "δ����ʱ����,����", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    strTemp = Mid(strTemp, 2)
    cllTime.Remove "�ֶμ��"
    cllTime.Add strTemp, "�ֶμ��"
    RaiseEvent zlRefreshCon(cllTime)
    cmdȡ��_Click
End Sub

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    optʱ��(m�ֶ�ʱ����).Value = True
    Call InitData
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
    End With
End Sub
Private Sub InitData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ����Ϣ
    '����:���˺�
    '����:2012-07-10 17:25:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTemp As String, varData As Variant
    Dim strSQL As String
    
    On Error GoTo errHandle
    strTemp = zlDatabase.GetPara("�������°�ʱ��", glngSys, , "07:00:00 AND 12:00:00")
    varData = Split(UCase(strTemp & " AND "), "AND")
    If IsDate(Trim(varData(0))) Then
        mWorkTime.����.�ϰ�ʱ�� = Trim(varData(0))
    Else
        mWorkTime.����.�ϰ�ʱ�� = "07:00:00"
    End If
    If IsDate(Trim(varData(1))) Then
        mWorkTime.����.�°�ʱ�� = Trim(varData(1))
    Else
        mWorkTime.����.�°�ʱ�� = "12:00:00"
    End If
    strTemp = zlDatabase.GetPara("�������°�ʱ��", glngSys, , "14:00:00 AND 18:00:00")
    varData = Split(UCase(strTemp & " AND "), "AND")
    If IsDate(Trim(varData(0))) Then
        mWorkTime.����.�ϰ�ʱ�� = Trim(varData(0))
    Else
        mWorkTime.����.�ϰ�ʱ�� = "14:00:00"
    End If
    If IsDate(Trim(varData(1))) Then
        mWorkTime.����.�°�ʱ�� = Trim(varData(1))
    Else
        mWorkTime.����.�°�ʱ�� = "18:00:00"
    End If
    strSQL = "Select ʱ���,to_char(��ʼʱ��,'hh24:mi:ss') as ��ʼʱ��,to_char(��ֹʱ��,'hh24:mi:ss') as ��ֹʱ�� from ʱ���"
    Set mrsʱ�� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsTime
        .Editable = flexEDKbdMouse
    End With
    Call optʱ��_Click(0)
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub optʱ��_Click(index As Integer)
   Call SetControlEnable
End Sub


Private Sub vsTime_AfterEdit(ByVal Row As Long, ByVal Col As Long)
     With vsTime
        If Val(.Cell(flexcpText, Row, .ColIndex("ʱ����"))) > 60 Or Val(.Cell(flexcpText, Row, .ColIndex("ʱ����"))) < 0 Then
            MsgBox "ʱ�������ܴ���60���ӻ�С��0���ӣ�", vbInformation, gstrSysName
           .Cell(flexcpText, Row, .ColIndex("ʱ����")) = ""
        End If
    End With
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
    If KeyAscii > 57 Or KeyAscii < 48 And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
