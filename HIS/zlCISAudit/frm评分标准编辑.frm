VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frm���ֱ�׼�༭ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ֱ�׼�༭"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8085
   Icon            =   "frm���ֱ�׼�༭.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   8085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraSource 
      Caption         =   "����Դ"
      Height          =   795
      Left            =   5190
      TabIndex        =   28
      Top             =   1005
      Visible         =   0   'False
      Width           =   2280
      Begin VB.OptionButton optSource 
         Caption         =   "��׼��"
         Height          =   195
         Index           =   0
         Left            =   780
         TabIndex        =   30
         Top             =   210
         Value           =   -1  'True
         Width           =   1320
      End
      Begin VB.OptionButton optSource 
         Caption         =   "EMR��"
         Height          =   180
         Index           =   1
         Left            =   780
         TabIndex        =   29
         Top             =   525
         Width           =   1125
      End
   End
   Begin VB.ComboBox cmb����ȼ� 
      Height          =   300
      Left            =   6240
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   4185
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6540
      TabIndex        =   22
      Top             =   4860
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5205
      TabIndex        =   21
      Top             =   4860
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   330
      TabIndex        =   25
      Top             =   4860
      Width           =   1100
   End
   Begin VB.OptionButton opt¼�뷽ʽ 
      Caption         =   "���ȼ�(&D)"
      Height          =   210
      Index           =   1
      Left            =   1530
      TabIndex        =   16
      Top             =   4230
      Width           =   1185
   End
   Begin VB.OptionButton opt¼�뷽ʽ 
      Caption         =   "������(&S)"
      Height          =   210
      Index           =   0
      Left            =   1530
      TabIndex        =   11
      Top             =   3795
      Value           =   -1  'True
      Width           =   1185
   End
   Begin VB.TextBox txt��׼��ֵ 
      Height          =   300
      Left            =   3420
      MaxLength       =   7
      TabIndex        =   13
      Top             =   3750
      Width           =   1185
   End
   Begin VB.ComboBox cmbȱ�ݵȼ� 
      Enabled         =   0   'False
      Height          =   300
      ItemData        =   "frm���ֱ�׼�༭.frx":000C
      Left            =   3420
      List            =   "frm���ֱ�׼�༭.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   4185
      Width           =   1185
   End
   Begin VB.ComboBox cmb���ֵ�λ 
      Height          =   300
      Left            =   6240
      TabIndex        =   15
      Top             =   3750
      Width           =   1185
   End
   Begin VB.TextBox txt���� 
      Height          =   705
      Left            =   1530
      MaxLength       =   2000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1085
      Width           =   5940
   End
   Begin VB.TextBox txt�������� 
      BackColor       =   &H8000000F&
      Height          =   300
      Left            =   1530
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   1
      Top             =   215
      Width           =   5940
   End
   Begin VB.TextBox txt���� 
      Height          =   705
      Left            =   1530
      MaxLength       =   50
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Top             =   1085
      Width           =   5940
   End
   Begin VB.CommandButton cmdXM 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   7155
      TabIndex        =   24
      Top             =   685
      Width           =   285
   End
   Begin VB.TextBox txt�ж�����_NotCheck 
      Height          =   1320
      Left            =   1530
      MaxLength       =   2000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   1920
      Width           =   5940
   End
   Begin VB.CommandButton cmdCheck 
      Height          =   300
      Left            =   7470
      Picture         =   "frm���ֱ�׼�༭.frx":0010
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1910
      Width           =   300
   End
   Begin zl9CISAudit.tipPopup tipPopup1 
      Height          =   420
      Left            =   1890
      Top             =   4500
      Width           =   3795
      _extentx        =   6694
      _extenty        =   741
      font            =   "frm���ֱ�׼�༭.frx":00EE
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   2175
      Left            =   1515
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   5940
      _cx             =   10477
      _cy             =   3836
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   0
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
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm���ֱ�׼�༭.frx":0116
      ScrollTrack     =   -1  'True
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
      Ellipsis        =   1
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   1
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
   Begin VB.TextBox txt�ϼ���Ŀ 
      BackColor       =   &H8000000F&
      Height          =   300
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   650
      Width           =   5940
   End
   Begin VB.Label labNote 
      AutoSize        =   -1  'True
      Caption         =   "ע��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   1500
      TabIndex        =   27
      Top             =   3360
      Width           =   390
   End
   Begin VB.Label lab����ID 
      AutoSize        =   -1  'True
      Caption         =   "[����ID]��[��ҳID]Ϊϵͳ�������ֱ����ϵͳ�еĲ���ID����ҳID��"
      Height          =   180
      Left            =   1950
      TabIndex        =   26
      Top             =   3360
      Width           =   5580
   End
   Begin VB.Label lab����ȼ� 
      Caption         =   "����ȼ�(&T)"
      Height          =   210
      Left            =   5205
      TabIndex        =   19
      Top             =   4260
      Width           =   1005
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8200
      Y1              =   4650
      Y2              =   4650
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   0
      X2              =   8200
      Y1              =   4680
      Y2              =   4665
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "¼�뷽ʽ(&I)"
      Height          =   180
      Left            =   450
      TabIndex        =   10
      Top             =   3810
      Width           =   990
   End
   Begin VB.Label lblFS2 
      Caption         =   "���ֵ�λ(&W)"
      Height          =   210
      Left            =   5205
      TabIndex        =   14
      Top             =   3795
      Width           =   1005
   End
   Begin VB.Label lblDJ 
      Caption         =   "�ȼ�(&G)"
      Enabled         =   0   'False
      Height          =   210
      Left            =   2730
      TabIndex        =   17
      Top             =   4230
      Width           =   1005
   End
   Begin VB.Label lblFS1 
      Caption         =   "����(&F)"
      Height          =   210
      Left            =   2730
      TabIndex        =   12
      Top             =   3795
      Width           =   1005
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ŀ����(&M)"
      Height          =   180
      Left            =   465
      TabIndex        =   6
      Top             =   1965
      Width           =   990
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������(&N)"
      Height          =   180
      Left            =   465
      TabIndex        =   0
      Top             =   270
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ϼ���Ŀ(&X)"
      Height          =   180
      Left            =   465
      TabIndex        =   2
      Top             =   675
      Width           =   990
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ŀ����(&B)"
      Height          =   180
      Left            =   465
      TabIndex        =   4
      Top             =   1080
      Width           =   990
   End
End
Attribute VB_Name = "frm���ֱ�׼�༭"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////
'       ���ܣ��ܹ��������ֱ�׼������Ŀ������Ŀ�����ܹ��޸����е����ֱ�׼��
'       ����ΰ  2005/1/6
'       ע������÷����Ѿ�ʹ�ã������������κ����ֱ�׼���������޸ģ�Ҳ����������
'///////////////////////////////////////////////////////////////////////////////////////

Option Explicit

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private m_lngID                 As Long     '��ǰ�༭�����ֱ�׼��ID��
Private m_lng����ID             As Long
Private m_lng�ϼ�ID             As Long
Private m_strEditMode           As String
Private m_lngOldRow             As Long
Private m_lngCurRow             As Long
Private m_blnModed              As Boolean
Private zlCheck                 As New clsCheck

Public Property Get Moded() As Boolean
   Moded = m_blnModed
End Property

Public Property Let Moded(ByVal blnModed As Boolean)
    m_blnModed = blnModed
End Property

'==============================================================================
'=���ܣ������ӿں��������ڴ����ʼ������:ID '��ʽΪ���룬��ID���ڣ�����IDֵǰ�ڵ���롣
'==============================================================================
Public Sub ShowForm(��ʽ As String, ����ID As Long, Optional �ϼ�ID As Long = 0, Optional ID As Long = 0, Optional blnUsed As Boolean = True)
    Dim rsTemp          As ADODB.Recordset
    On Error GoTo errH
    txt����.Locked = Not blnUsed
    txt����.Locked = Not blnUsed
    cmdXM.Enabled = blnUsed
    zlCheck.Sys_System Me
    
    m_blnModed = False
    m_lng����ID = ����ID  '��ѡ����
    m_lng�ϼ�ID = �ϼ�ID
    m_lngID = ID          'Ϊ0��ʾ����
    m_lngCurRow = -1
    If m_lng����ID < 1 Then
        Unload Me
        MsgBox "����ѡ��һ�����ַ���������㻹û��¼�뷽������¼�룡", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Call FillCmbs
    Call FillInitFixData
    
    m_strEditMode = ��ʽ
    If m_strEditMode = "����" Then
        Me.Caption = "����" & IIf(�ϼ�ID = 0, "��Ŀ", "��׼")
    ElseIf m_strEditMode = "����" Then
        Me.Caption = "����" & IIf(�ϼ�ID = 0, "��Ŀ", "��׼")
    Else
        Me.Caption = "�޸�" & IIf(�ϼ�ID = 0, "��Ŀ", "��׼")
        FillInitData
        txt�ϼ���Ŀ.TabIndex = 0
    End If
    
    '����Ѿ������¼���Ŀ���Ͳ�����ѡ���ϼ���Ŀ�ˣ�
    gstrSQL = "select count(*) from �������ֱ�׼��ͼ where ����='��' and ID = [1] and ����ID = [2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lngID, m_lng����ID)
    If rsTemp.Fields(0) > 0 And m_strEditMode = "�޸�" Then
        '�����¼���Ŀ
        cmdXM.Enabled = False
    End If
    rsTemp.Close

    Me.Show 1
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�������ʾ
'==============================================================================
Private Sub ShowSet()
    Dim blnһ����Ŀ         As Boolean

    On Error GoTo errH
    
    If gobjEmr Is Nothing Then
        fraSource.Visible = False
        optSource(0).Value = True
    Else
        fraSource.Visible = True
        optSource(0).Value = False
        optSource(1).Value = True
    End If
    
    If Len(txt�ϼ���Ŀ) > 0 Then
        blnһ����Ŀ = False
    Else
        blnһ����Ŀ = True
    End If
    
    If blnһ����Ŀ Then
        fraSource.Visible = False '��Ŀ����ָ������Դ
        lbl����.Visible = True
        lbl����.Caption = "��Ŀ����(&B)"
        lbl����.Caption = "��Ŀ����(&M)"
        txt����.Visible = True
        txt����.Move txt����.Left, lbl����.Top, txt����.Width, 1500
        txt�ж�����_NotCheck.Visible = False
        lab����ID.Visible = False
        labNote.Visible = False
        cmdCheck.Visible = False
    Else
        cmdCheck.Visible = True
        lbl����.Caption = "��׼����(&B)"
        lbl����.Caption = "�ж�����(&M)"
        txt����.Visible = False
        txt����.Visible = True
        txt�ж�����_NotCheck.Visible = True
        lab����ID.Visible = True
        labNote.Visible = True
        txt����.Move txt����.Left, txt����.Top, IIf(gobjEmr Is Nothing, txt����.Width, txt����.Width - fraSource.Width - 100), txt����.Height
        txt����.Text = ""
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ���������ID����̶����ݣ��緽��ID���ϼ�ID
'==============================================================================
Private Sub FillInitFixData()
    Dim rs As ADODB.Recordset

    On Error GoTo errH
    
    gstrSQL = "select ���� from �������ַ��� where ID = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lng����ID)
    If Not rs.EOF Then
        txt�������� = IIf(IsNull(rs.Fields("����")), "", rs.Fields("����"))
    Else
        Unload Me
        MsgBox "����ѡ�����ַ�����", vbInformation, "����ID����"
        Exit Sub
    End If
    gstrSQL = "select ���� from �������ֱ�׼ where ID = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lng�ϼ�ID)
    
    If Not rs.EOF Then
        txt�ϼ���Ŀ = IIf(IsNull(rs.Fields("����")), "", rs.Fields("����"))
    End If
    rs.Close
    
    Call ShowSet
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�����ID�����ʼ����
'==============================================================================
Private Sub FillInitData()
    Dim rs              As ADODB.Recordset
    On Error GoTo errH
    
    gstrSQL = "select A.ID,A.�ϼ�ID,A.����ID,A.����,A.����,A.��׼��ֵ,A.ȱ�ݵȼ�,A.���ֵ�λ,A.�ϼ����,A.���,A.�ж�����,B.���� as �ϼ���Ŀ,A.����ȼ�,A.����Դ from �������ֱ�׼ A,�������ֱ�׼ B where A.�ϼ�ID=B.ID(+) and A.ID = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lngID)
    If Not rs.EOF Then
        txt�ϼ���Ŀ.Text = IIf(IsNull(rs.Fields("�ϼ���Ŀ")), "", rs.Fields("�ϼ���Ŀ"))
        txt����.Text = IIf(IsNull(rs.Fields("����")), "", rs.Fields("����"))
        txt����.Text = IIf(IsNull(rs.Fields("����")), "", rs.Fields("����"))
        txt�ж�����_NotCheck.Text = IIf(IsNull(rs.Fields("�ж�����")), "", rs.Fields("�ж�����"))
        If rs!����Դ = 0 Then
            optSource(0).Value = True
            optSource(1).Value = False
        Else
            optSource(0).Value = False
            optSource(1).Value = True
        End If
        If IsNull(rs.Fields("ȱ�ݵȼ�")) Then
            txt��׼��ֵ = IIf(IsNull(rs.Fields("��׼��ֵ")), 0, IIf(rs.Fields("��׼��ֵ") < 1, Format(rs.Fields("��׼��ֵ"), "0.0"), rs.Fields("��׼��ֵ")))
            cmb���ֵ�λ.Text = IIf(IsNull(rs.Fields("���ֵ�λ")), "", rs.Fields("���ֵ�λ"))
            Set¼�뷽ʽ 0
        Else
            Select Case rs.Fields("ȱ�ݵȼ�")
                Case "��"
                    cmbȱ�ݵȼ�.ListIndex = 0
                Case "��"
                    cmbȱ�ݵȼ�.ListIndex = 1
                Case "��"
                    cmbȱ�ݵȼ�.ListIndex = 2
                Case "��"
                    cmbȱ�ݵȼ�.ListIndex = 3
            End Select
            Select Case rs.Fields("����ȼ�")
                Case "��"
                    cmb����ȼ�.ListIndex = 0
                Case "��"
                    cmb����ȼ�.ListIndex = 1
                Case "��", ""
                    cmb����ȼ�.ListIndex = 2
            End Select
            Set¼�뷽ʽ 1
        End If
        zlControl.TxtSelAll txt�ϼ���Ŀ
    Else
        Unload Me
        MsgBox "��ʼ�����ݴ���û�з��ָ������ֱ�׼�������ԡ�", vbOKOnly + vbInformation, "��������"
        Exit Sub
    End If
    Call ShowSet
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�����ѡ���������
'==============================================================================
Private Sub FillCmbs()
    On Error GoTo errH
    
    With cmbȱ�ݵȼ�
        .AddItem "�׼�"
        .AddItem "�Ҽ�"
        .AddItem "����"
        .AddItem "������"
        .ListIndex = 1
    End With
    With cmb���ֵ�λ
        .AddItem ""
        .AddItem "��"
        .AddItem "��"
        .AddItem "��"
    End With
    With cmb����ȼ�
        .AddItem "�Ҽ�"
        .AddItem "����"
        .AddItem "���ϸ�"
        .ListIndex = 1
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ����ֵ�λ��ʾ
'==============================================================================
Private Sub cmb���ֵ�λ_GotFocus()
    On Error GoTo errH
    Call zlCommFun.OpenIme(True)
    ShowTips cmb���ֵ�λ, "���������ֵ�λ��", "���ֵ�λ"
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmbȱ�ݵȼ�_Click()
    cmb����ȼ�.Enabled = (cmbȱ�ݵȼ�.Text = "������")
End Sub

'==============================================================================
'=���ܣ�ȱ�ݵȼ���ʾ
'==============================================================================
Private Sub cmbȱ�ݵȼ�_GotFocus()
    On Error GoTo errH
    
    ShowTips cmbȱ�ݵȼ�, "����ȱ�ݵȼ��趨��", "ȱ�ݵȼ�"
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�ȱ�ݵȼ����س�ȷ���ͱ�������
'==============================================================================
Private Sub cmbȱ�ݵȼ�_KeyPress(KeyAscii As Integer)
    On Error GoTo errH
    If KeyAscii = vbKeyReturn Then
        Call cmdOk_Click
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmbȱ�ݵȼ�_Validate(Cancel As Boolean)
    cmb����ȼ�.Enabled = (cmbȱ�ݵȼ�.Text = "������")
End Sub

'==============================================================================
'=���ܣ�ȡ������
'==============================================================================
Private Sub cmdCancel_Click()
    On Error GoTo errH
    m_blnModed = False
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �ж����ݼ��
'==============================================================================
Private Sub cmdCheck_Click()
    On Error GoTo errH
    Call CheckAuditSql_IN(Trim(txt�ж�����_NotCheck.Text), True, IIf(optSource(0).Value, 0, 1))
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ��������
'==============================================================================
Private Sub cmdHelp_Click()
    On Error GoTo errH
    ShowHelp App.ProductName, Me.hWnd, Me.Name, 3
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�ȷ����������
'==============================================================================
Private Sub cmdOk_Click()
Dim blnBigXM        As Boolean
Dim strT As String, str�ȼ� As String, strAudit As String, str����ȼ� As String, intSource As Integer
    
    On Error GoTo errH
    
    '���Ϸ���
    If Not IsValid() Then Exit Sub
    intSource = IIf(optSource(0).Value, 0, 1)
    If m_lng�ϼ�ID <= 0 Then
       '�ϼ���ĿΪ�գ���ʾΪһ����Ŀ
        blnBigXM = True
    Else
        blnBigXM = False
    End If
    Select Case cmbȱ�ݵȼ�.Text
        Case "�׼�"
            str�ȼ� = "��"
        Case "�Ҽ�"
            str�ȼ� = "��"
        Case "����"
            str�ȼ� = "��"
        Case "������"
            str�ȼ� = "��"
    End Select
    If str�ȼ� = "��" Then
        Select Case cmb����ȼ�.Text
            Case "�Ҽ�"
                str����ȼ� = "��"
            Case "����"
                str����ȼ� = "��"
            Case "���ϸ�"
                str����ȼ� = "��"
        End Select
    End If
    '�ж����ݱ���ʱ�赥����˫��
    strAudit = Replace(txt�ж�����_NotCheck.Text, "'", "''")
    If m_strEditMode = "����" Or m_strEditMode = "����" Then
        If cmdCheck.Visible Then
            If Not CheckAuditSql_IN(Trim(txt�ж�����_NotCheck.Text), False, IIf(optSource(0).Value, 0, 1)) Then Exit Sub
        End If
        If blnBigXM Then
            strT = "ZL_�������ֱ�׼_Insert"
            gstrSQL = strT & _
                    "(" & zlDatabase.GetNextId("�������ֱ�׼") & "," & IIf(m_lng�ϼ�ID = 0, "Null", CStr(m_lng�ϼ�ID)) & "," & m_lng����ID & ",'" & txt���� & "','" & txt���� & _
                    "'," & IIf(txt��׼��ֵ.Enabled = False, "null", Val(txt��׼��ֵ)) & ",'" & IIf(opt¼�뷽ʽ(0).Value, "", str�ȼ�) & "','" & cmb���ֵ�λ.Text & "'," & m_lngID & ",'" & strAudit & "','" & str����ȼ� & "'," & intSource & ")"
        Else
            strT = "ZL_�������ֱ�׼_Insert"
            gstrSQL = strT & _
                    "(" & zlDatabase.GetNextId("�������ֱ�׼") & "," & IIf(m_lng�ϼ�ID = 0, "Null", CStr(m_lng�ϼ�ID)) & "," & m_lng����ID & ",'" & txt���� & "','" & txt���� & _
                    "'," & IIf(txt��׼��ֵ.Enabled = False, "null", Val(txt��׼��ֵ)) & ",'" & IIf(opt¼�뷽ʽ(0).Value, "", str�ȼ�) & "','" & cmb���ֵ�λ.Text & "',0,'" & strAudit & "','" & str����ȼ� & "'," & intSource & ")"
        End If
    Else
        If cmdCheck.Visible Then
            If Not CheckAuditSql_IN(Trim(txt�ж�����_NotCheck.Text), False, IIf(optSource(0).Value, 0, 1)) Then Exit Sub
        End If
        strT = "ZL_�������ֱ�׼_Update"
        gstrSQL = strT & _
                "(" & CStr(m_lngID) & "," & IIf(m_lng�ϼ�ID = 0, "Null", CStr(m_lng�ϼ�ID)) & "," & m_lng����ID & ",'" & txt���� & "','" & txt���� & _
                "'," & IIf(txt��׼��ֵ.Enabled = False, "null", Val(txt��׼��ֵ)) & ",'" & IIf(opt¼�뷽ʽ(0).Value, "", str�ȼ�) & "','" & cmb���ֵ�λ.Text & "','" & strAudit & "','" & str����ȼ� & "'," & intSource & ")"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    m_blnModed = True
    '�ֹ�ˢ��
    Call frm���ֱ�׼ά��.DataLoad
    zlCheck.Msg_OK "���ֱ�׼����ɹ���"
    If m_strEditMode = "����" Then
        txt����.Text = ""
        txt����.Text = ""
        txt��׼��ֵ.Text = ""
        cmb���ֵ�λ.Text = ""
        txt�ж�����_NotCheck.Text = ""
        opt¼�뷽ʽ(0).Value = True
        If txt����.Visible = True Then
            txt����.SetFocus
        Else
            txt����.SetFocus
        End If
    Else
        Unload Me
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ����ݼ����ȷ���
'=���أ���Ч����True,����ΪFalse
'=˵����ͬһ��������Ŀ���Ʋ����ظ�
'==============================================================================
Private Function IsValid() As Boolean
    Dim blnһ����Ŀ         As Boolean
    
    On Error GoTo errH
    
    IsValid = False
    If Len(txt�ϼ���Ŀ) > 0 Then
        blnһ����Ŀ = False
    Else
        blnһ����Ŀ = True
    End If
    If blnһ����Ŀ And m_strEditMode = "����" Then
        Dim rsTmp As New ADODB.Recordset
        gstrSQL = "select ���� from �������ֱ�׼ where �ϼ�ID is null and ����ID = [1] And ���� = [2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lng����ID, Trim(txt����.Text))
        If Not rsTmp.EOF Then
            zlCheck.Msg_OK "ͬһ�����е���Ŀ���Ʋ����ظ���"
            zlControl.TxtSelAll txt����: txt����.SetFocus
            Exit Function
        End If
    End If
    '����StrIsValid������ȷ���ַ�����ʽ��ȷ��ע�⣺����ʹ�õ���lenBֵ����Ӧ���ݱ����е�ֵ��
    If zlCommFun.StrIsValid(txt����.Text, txt����.MaxLength * 2) = False Then
        zlCheck.Msg_OK "���������ƣ�"
        zlControl.TxtSelAll txt����: txt����.SetFocus
        Exit Function
    End If
    If zlCommFun.StrIsValid(txt����.Text, txt����.MaxLength * 2) = False Then
        zlCheck.Msg_OK "������������"
        zlControl.TxtSelAll txt����: txt����.SetFocus
        Exit Function
    End If
    If zlCommFun.StrIsValid(cmb���ֵ�λ.Text, 8) = False Then
        zlCheck.Msg_OK "���ֵ�λ���ȳ���4�����֣�������¼�롣"
        cmb���ֵ�λ.SetFocus
        Exit Function
    End If
    If Len(Trim(txt����)) = 0 And Len(Trim(txt����)) = 0 Then
        zlCheck.Msg_OK "���ƺ���������ͬʱΪ�գ�������¼�롣"
        If txt����.Visible = True Then
            zlControl.TxtSelAll txt����: txt����.SetFocus
        Else
            zlControl.TxtSelAll txt����: txt����.SetFocus
        End If
        Exit Function
    End If
    If opt¼�뷽ʽ(0).Value And Len(Trim(txt��׼��ֵ)) = 0 Then
        zlCheck.Msg_OK "�������׼��ֵ��"
        zlControl.TxtSelAll txt��׼��ֵ: txt��׼��ֵ.SetFocus
        Exit Function
    End If
    If Len(Trim(txt��׼��ֵ)) > 0 Then
        If Not IsNumeric(txt��׼��ֵ) Then
            zlCheck.Msg_OK "�������׼��ֵ��"
            zlControl.TxtSelAll txt��׼��ֵ: txt��׼��ֵ.SetFocus
            Exit Function
        End If
        If Val(txt��׼��ֵ.Text) > 9999# Then
            zlCheck.Msg_OK "����ı�׼��ֵ̫��"
            zlControl.TxtSelAll txt��׼��ֵ: txt��׼��ֵ.SetFocus
            Exit Function
        End If
    End If
    
    '��ǰ����վû�а�װ�°没���������ȴ���޸�����ԴΪEMR����ж����ݣ�Ӧ��ֹ
    If gobjEmr Is Nothing And optSource(1).Value Then
        zlCheck.Msg_OK "��ǰ����վδ��װ�����������ֹ�޸���Ҫ��EMR��ִ�е��ж����ݣ�"
        Exit Function
    End If
    
    IsValid = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=���ܣ���̬װ��һ����Ŀ
'==============================================================================
Private Sub cmdXM_Click()
    Dim rsTemp              As ADODB.Recordset
    
    On Error GoTo errH

    If cmdXM.Tag = "��" Then
        cmdXM.Tag = ""
        Grid.Visible = False
        Grid_SelChange
        Exit Sub
    Else
        cmdXM.Tag = "��"
    End If
    
    With Grid
        .Clear
        .Redraw = flexRDNone
        If m_strEditMode = "�޸�" Then
            '����Ǳ༭ģʽ��һ����Ŀ��Ҫ�ų�������
            gstrSQL = "select ID,��׼��ֵ,����,���� from �������ֱ�׼ Where �ϼ�ID is null and ID <> [1] and ����ID = [2]"
        Else
            gstrSQL = "select ID,��׼��ֵ,����,���� from �������ֱ�׼ Where �ϼ�ID is null and ����ID = [2]"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lngID, m_lng����ID)
        
        .FocusRect = flexFocusSolid
        '��������
        .Cols = 4
        .Rows = rsTemp.RecordCount + 2
        Dim i As Long
        .Cell(flexcpText, 0, 0) = "ID"
        .Cell(flexcpText, 0, 1) = "����"
        .Cell(flexcpText, 0, 2) = "��׼��ֵ"
        .Cell(flexcpText, 0, 3) = "����"
        .Cell(flexcpText, 1, 0) = "0"
        .Cell(flexcpText, 1, 1) = "<��>"
        .Cell(flexcpText, 1, 2) = "<��>"
        .Cell(flexcpText, 1, 3) = "<��>"
        i = 2
        Do Until rsTemp.EOF
            If m_lng�ϼ�ID = rsTemp.Fields("ID") Then m_lngCurRow = i
            .Cell(flexcpText, i, 0) = IIf(IsNull(rsTemp.Fields("ID")), "", rsTemp.Fields("ID"))
            .Cell(flexcpText, i, 1) = IIf(IsNull(rsTemp.Fields("����")), "", rsTemp.Fields("����"))
            .Cell(flexcpText, i, 2) = IIf(IsNull(rsTemp.Fields("��׼��ֵ")), "", Format(rsTemp.Fields("��׼��ֵ"), "####��"))
            .Cell(flexcpText, i, 3) = IIf(IsNull(rsTemp.Fields("����")), "", IIf(Len(rsTemp.Fields("����")) > 30, Left(rsTemp.Fields("����"), 27) + "...", rsTemp.Fields("����")))
            rsTemp.MoveNext
            i = i + 1
        Loop
        '�Զ�����
        .WordWrap = True
        '�и�����
        .RowHeightMin = 250
        .RowHeightMax = 300
        .ColAlignment(.ColIndex("ID")) = flexAlignLeftCenter
        .ColAlignment(.ColIndex("��׼��ֵ")) = flexAlignCenterCenter
        .ColAlignment(.ColIndex("����")) = flexAlignLeftTop
        .ColAlignment(.ColIndex("����")) = flexAlignLeftTop
        '�������
        .ColWidth(.ColIndex("ID")) = 0
        .ColWidth(.ColIndex("����")) = 1600
        .ColWidth(.ColIndex("��׼��ֵ")) = 800
        .ColWidth(.ColIndex("����")) = 4000
        '���������
        .ColWidthMax = 4000
        '�Զ���Ӧ�иߡ��п�
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 1
        .AutoSize 3
        .SelectionMode = flexSelectionByRow
        .AllowBigSelection = False
        
        'ѡ����ǰ����
        If m_lngCurRow > 0 And m_lngCurRow < i Then
            .Row = m_lngCurRow
            .ShowCell m_lngCurRow, 2
        Else
            m_lngCurRow = 1
            .Row = 1
            .ShowCell m_lngCurRow, 2
        End If
        .ZOrder 0
        .Redraw = flexRDBuffered
        .Visible = True
        If .Visible = True Then .SetFocus
    End With

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ���Ŀ�õ�����ʱ��ʾTips��ʾ
'==============================================================================
Private Sub cmdXM_GotFocus()
    On Error GoTo errH
    ShowTips cmdXM, "�����س�������ѡ������������ѡ���ϼ���Ŀ��", "�ϼ���Ŀ"
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ���Ŀ��F1��ʾTips��ʾ
'==============================================================================
Private Sub cmdXM_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo errH
    If KeyCode = vbKeyF1 Then
        ShowTips cmdXM, "ѡ���ϼ���Ŀ�����ð�ť��", "ѡ���ϼ���Ŀ"
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�ҳ��ؼ���ʼ��
'==============================================================================
Private Sub Form_Initialize()
    On Error GoTo errH
    Call InitCommonControls
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�ҳ���ʼ��
'==============================================================================
Private Sub Form_Load()
    On Error GoTo errH
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ���Ŀѡ��״̬
'==============================================================================
Private Sub Grid_Click()
    On Error GoTo errH
    Call Grid_SelChange
    If cmdXM.Tag = "��" Then
        cmdXM.Tag = ""
        Grid.Visible = False
        Exit Sub
    Else
        cmdXM.Tag = "��"
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�Grid�õ�����ʱ��ʾTips��ʾ
'==============================================================================
Private Sub Grid_GotFocus()
    On Error GoTo errH
    ShowTips cmdXM, "�����س���ѡ�����ϼ���Ŀ��", "ѡ����Ŀ"
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�Grid����ȷ��
'==============================================================================
Private Sub Grid_KeyPress(KeyAscii As Integer)
    On Error GoTo errH
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If txt����.Visible = True Then
            zlControl.TxtSelAll txt����
            txt����.SetFocus
        Else
            zlControl.TxtSelAll txt����
            txt����.SetFocus
        End If
        Grid.Visible = False
        cmdXM.Tag = ""
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�Grid�����кż�ID
'==============================================================================
Private Sub Grid_SelChange()
    Dim m_lngID As Long

    On Error GoTo errH
    
    m_lng�ϼ�ID = 0
    m_lngCurRow = Grid.Row
    If m_lngCurRow <= 0 Then Exit Sub
    m_lngCurRow = Grid.Row     '��ȡ�к�
    m_lng�ϼ�ID = Grid.Cell(flexcpText, m_lngCurRow, 0)
    txt�ϼ���Ŀ.Text = IIf(m_lng�ϼ�ID = 0, "", Grid.Cell(flexcpText, m_lngCurRow, 1))
    m_lngOldRow = m_lngCurRow
    Call ShowSet
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub optSource_Click(Index As Integer)
    If Index = 0 Then
        lab����ID.Caption = "[����ID]��[��ҳID]Ϊϵͳ�������ֱ����ϵͳ�еĲ���ID����ҳID��"
        txt�ж�����_NotCheck.Tag = 0 '���ڷŴ�༭����
    Else
        lab����ID.Caption = "[MID]��[ALIDIN]Ϊϵͳ�������ֱ����EMR�еĲ���ID����ԺID"
        lab����ID.ToolTipText = "ʹ��EMR�е�ID,��[MID]��[ALIDIN]�ⶼ��Ҫʹ��HextoRawת�����Ա��õ�������"
        txt�ж�����_NotCheck.Tag = 1 '���ڷŴ�༭����
    End If
End Sub
'==============================================================================
'=���ܣ�¼�뷽ʽ�趨
'==============================================================================
Private Sub opt¼�뷽ʽ_Click(Index As Integer)
    On Error GoTo errH
    Set¼�뷽ʽ Index
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�����¼�뷽ʽ���ÿؼ�״̬
'==============================================================================
Private Sub Set¼�뷽ʽ(i As Integer)
    On Error GoTo errH
    If i = 0 Then
        cmbȱ�ݵȼ�.Enabled = False
        lblDJ.Enabled = False
        txt��׼��ֵ.Enabled = True
        lblFS1.Enabled = True
        lblFS2.Enabled = True
        cmb���ֵ�λ.Enabled = True
        cmb����ȼ�.Enabled = False
    Else
        cmbȱ�ݵȼ�.Enabled = True
        lblDJ.Enabled = True
        txt��׼��ֵ.Enabled = False
        lblFS1.Enabled = False
        lblFS2.Enabled = False
        cmb���ֵ�λ.Enabled = False
        cmb����ȼ�.Enabled = (cmbȱ�ݵȼ�.Text = "������")
    End If
    opt¼�뷽ʽ(i).Value = True
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�¼�뷽ʽ�õ�������ʾTips��ʾ
'==============================================================================
Private Sub opt¼�뷽ʽ_GotFocus(Index As Integer)
    On Error GoTo errH
    If Index = 0 Then
        ShowTips opt¼�뷽ʽ(0), "���մ�ֵķ�ʽ�������֣��˴��ṩ����ı�׼������һ���������������ӷ���۷��ɷ�����ȷ����", "������"
    Else
        ShowTips opt¼�뷽ʽ(1), "����ĳЩ��Ҫ�����������ϸ���ô��������ֱ�Ӷ�λ���Ҽ����򡰱���������Ҫ������ѡ����ȱ�ݵȼ���������������ʾ�������ϸ��������������ϸ񣬲�����ȼ�������", "���ȼ�"
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�txt��׼��ֵ �õ�������ʾTips��ʾ
'==============================================================================
Private Sub txt��׼��ֵ_GotFocus()
    On Error GoTo errH
    ShowTips txt��׼��ֵ, "������������", "��׼��ֵ"
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�txt�������� �õ�������ʾTips��ʾ
'==============================================================================
Private Sub txt��������_GotFocus()
    On Error GoTo errH
    ShowTips txt��������, "���ַ��������ơ������ﲻ�����޸ġ�", "��������"
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�txt���� �õ�������ʾTips��ʾ
'==============================================================================
Private Sub txt����_GotFocus()
    On Error GoTo errH
    ShowTips txt����, "����¼����������Ŀ������������������Ŀ�����ơ�", "����"
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�txt���� �õ�������ʾTips��ʾ
'==============================================================================
Private Sub txt����_GotFocus()
    On Error GoTo errH
    zlControl.TxtSelAll txt����
    ShowTips txt����, "���������������Ŀ������������������������Ŀ�����ơ�", "����"
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�txt�ϼ���Ŀ �õ�������ʾTips��ʾ
'==============================================================================
Private Sub txt�ϼ���Ŀ_GotFocus()
    On Error GoTo errH
    zlControl.TxtSelAll txt�ϼ���Ŀ
    ShowTips txt�ϼ���Ŀ, "�������Ҫ������������Ŀ������ѡ�������ϼ���Ŀ��", "�ϼ���Ŀ"
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ��ؼ�ͨ��Tips��ʾ
'==============================================================================
Private Sub ShowTips(ctl As Control, str���� As String, Optional str���� As String = "��ʾ��Ϣ", Optional lngʱ�� As Long = 2500)
    Dim X As Single, Y As Single
    On Error GoTo errH
    X = (ctl.Left + ctl.Width / 2) / Screen.TwipsPerPixelX
    Y = (ctl.Top + ctl.Height) / Screen.TwipsPerPixelY
    If Len(str����) > 0 Then
        tipPopup1.Hide
        tipPopup1.StandardIcon = IDI_INFORMATION
        tipPopup1.ShowCloseButton = True
        tipPopup1.TimeOut = lngʱ��
        tipPopup1.Title = str����
        tipPopup1.Text = str����
        tipPopup1.Show Me.hWnd, X, Y
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



