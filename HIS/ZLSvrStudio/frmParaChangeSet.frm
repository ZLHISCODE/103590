VERSION 5.00
Begin VB.Form frmParaChangeSet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�ı��������"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8745
   Icon            =   "frmParaChangeSet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   8745
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdClear 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7500
      TabIndex        =   31
      Top             =   5400
      Width           =   1100
   End
   Begin VB.Frame fra������Ϣ 
      Caption         =   "�䶯��Ϣ"
      Height          =   1050
      Left            =   0
      TabIndex        =   23
      Top             =   4260
      Width           =   8640
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   7
         Left            =   915
         TabIndex        =   25
         Top             =   255
         Width           =   7605
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000012&
         Height          =   180
         Index           =   6
         Left            =   5025
         TabIndex        =   29
         Top             =   675
         Width           =   3495
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000012&
         Height          =   180
         Index           =   5
         Left            =   915
         TabIndex        =   27
         Top             =   675
         Width           =   2415
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�䶯ԭ��"
         Height          =   180
         Index           =   11
         Left            =   135
         TabIndex        =   24
         Top             =   315
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�䶯ʱ��"
         Height          =   180
         Index           =   10
         Left            =   4245
         TabIndex        =   28
         Top             =   675
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�䶯��"
         Height          =   180
         Index           =   9
         Left            =   330
         TabIndex        =   26
         Top             =   675
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "��Ȩ��ʽ�䶯"
      Height          =   1065
      Left            =   5040
      TabIndex        =   18
      Top             =   3105
      Width           =   3600
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000012&
         Height          =   180
         Index           =   3
         Left            =   1275
         TabIndex        =   20
         Text            =   "��Ҫ��Ȩ"
         Top             =   315
         Width           =   1515
      End
      Begin VB.ComboBox cboEdit 
         Height          =   300
         Index           =   1
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   630
         Width           =   1995
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ԭ��Ȩ��ʽ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   8
         Left            =   120
         TabIndex        =   19
         Top             =   315
         Width           =   1170
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����Ȩ��ʽ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   7
         Left            =   105
         TabIndex        =   21
         Top             =   675
         Width           =   1170
      End
   End
   Begin VB.Frame fra�����䶯 
      Caption         =   "�������ͱ䶯"
      Height          =   2925
      Left            =   5040
      TabIndex        =   8
      Top             =   105
      Width           =   3585
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000012&
         Height          =   180
         Index           =   4
         Left            =   1275
         TabIndex        =   10
         Text            =   "����ģ��"
         Top             =   315
         Width           =   1350
      End
      Begin VB.ComboBox cboEdit 
         Height          =   300
         Index           =   0
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   600
         Width           =   2070
      End
      Begin VB.Label lblMemo 
         AutoSize        =   -1  'True
         Caption         =   "��������˵����"
         ForeColor       =   &H80000011&
         Height          =   180
         Index           =   4
         Left            =   75
         TabIndex        =   13
         Top             =   1020
         Width           =   1260
      End
      Begin VB.Label lblMemo 
         Caption         =   "  ����˽��ģ���ʾ��Ը�ģ��ֲ����û�����վ��Ĳ�����"
         ForeColor       =   &H80000011&
         Height          =   420
         Index           =   3
         Left            =   60
         TabIndex        =   17
         Top             =   2460
         Width           =   3420
      End
      Begin VB.Label lblMemo 
         Caption         =   "  ˽��ģ���ʾ��Ը�ģ��ֲ����û�������վ��Ĳ�����"
         ForeColor       =   &H80000011&
         Height          =   420
         Index           =   2
         Left            =   60
         TabIndex        =   16
         Top             =   2040
         Width           =   3420
      End
      Begin VB.Label lblMemo 
         Caption         =   "  ��������ģ���ʾ��Ը�ģ�鲻�ֲ����û�����Ҫ�ֻ����Ĳ�����"
         ForeColor       =   &H80000011&
         Height          =   420
         Index           =   1
         Left            =   60
         TabIndex        =   15
         Top             =   1635
         Width           =   3420
      End
      Begin VB.Label lblMemo 
         Caption         =   "  ����ģ���ʾ��Ը�ģ�鲻�ֲ����û����ֻ����Ĳ�����"
         ForeColor       =   &H80000011&
         Height          =   420
         Index           =   0
         Left            =   60
         TabIndex        =   14
         Top             =   1245
         Width           =   3420
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�ֲ������ͣ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   11
         Top             =   675
         Width           =   1170
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ԭ�������ͣ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   5
         Left            =   120
         TabIndex        =   9
         Top             =   315
         Width           =   1170
      End
   End
   Begin VB.Frame fra���� 
      Caption         =   "����������Ϣ"
      Height          =   4050
      Left            =   0
      TabIndex        =   30
      Top             =   105
      Width           =   4980
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000012&
         Height          =   180
         Index           =   2
         Left            =   1095
         TabIndex        =   5
         Top             =   1125
         Width           =   3810
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000012&
         Height          =   180
         Index           =   1
         Left            =   1095
         TabIndex        =   3
         Tag             =   "������"
         Top             =   765
         Width           =   960
      End
      Begin VB.TextBox txtEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000012&
         Height          =   180
         Index           =   0
         Left            =   1095
         TabIndex        =   1
         Tag             =   "ģ��"
         Top             =   375
         Width           =   3810
      End
      Begin VB.Label lblEdit 
         Appearance      =   0  'Flat
         Caption         =   "   ddddddddddd"
         Height          =   2145
         Index           =   4
         Left            =   75
         TabIndex        =   7
         Tag             =   "����˵��"
         Top             =   1815
         Width           =   4815
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����˵����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   3
         Left            =   75
         TabIndex        =   6
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�������ƣ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   90
         TabIndex        =   4
         Top             =   1125
         Width           =   975
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�����ţ�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   270
         TabIndex        =   2
         Top             =   765
         Width           =   780
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ģ�飺"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   450
         TabIndex        =   0
         Top             =   375
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   6345
      TabIndex        =   32
      Top             =   5400
      Width           =   1100
   End
End
Attribute VB_Name = "frmParaChangeSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng����ID As Long
Private mblnOK As Boolean
Private mblnChange As Boolean
Private mstrUserName As String
Private mblnFirst As Boolean
Private mblnNotClick As Boolean
Private Enum mTxt_idx
    idx_ģ�� = 0
    idx_������ = 1
    idx_�������� = 2
    idx_ԭ��Ȩ��ʽ = 3
    idx_ԭ�������� = 4
    idx_�䶯�� = 5
    idx_�䶯ʱ�� = 6
    idx_�䶯ԭ�� = 7
End Enum
Public Function ShowEdit(ByVal frmMain As Form, ByVal lng����id As Long, ByVal strUserName As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʾ�༭����
    '���:frmMain-������
    '     lng����ID-����ֵ
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-02-19 14:16:42
    '-----------------------------------------------------------------------------------------------------------
    mlng����ID = lng����id: mblnChange = False: mblnOK = False: mstrUserName = strUserName: mblnFirst = True
    Me.Show 1
    ShowEdit = mblnOK
End Function
Private Function GetParaType(ByVal lngģ�� As Long, ByVal int˽�� As Integer, ByVal int���� As Integer) As String
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ��������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-02-17 16:44:21
    '-----------------------------------------------------------------------------------------------------------
    If lngģ�� = 0 Then
        '����ģ��,֤��ֻ����������:����ȫ�ֺ�˽��ȫ��
        GetParaType = IIf(int˽�� = 0, "����ȫ��", "˽��ȫ��")
        Exit Function
    End If
    '��ģ��Ĵ���
    If int���� = 0 Then
        '���Ǳ��������,ֻ����������:����ģ���˽��ģ��
         GetParaType = IIf(int˽�� = 0, "����ģ��", "˽��ģ��")
         Exit Function
    End If
    '�Ա�����ģ����д���Ҳ���������:
    GetParaType = IIf(int˽�� = 0, "��������ģ��", "����˽��ģ��")
End Function
Private Function LoadParaInfor() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:���ز�����Ϣ
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-02-19 14:19:26
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strCurDate As String
    
    Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.b_Public.Get_Current_Date")
    strCurDate = Format(rsTemp!����, "yyyy-mm-dd HH:MM:SS")
    
    Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Parameter", mlng����ID)
    'ID, ϵͳ,ģ��,˽��,������,������, ����ֵ, ȱʡֵ, ����˵��, ����, ��Ȩ, �̶�,ģ������
    If rsTemp.EOF Then
        MsgBox "����δ�ҵ��������Ѿ�������ɾ��,����!", vbOKOnly, gstrSysName
        Exit Function
    End If
    mblnNotClick = True

    txtEdit(idx_ģ��) = Nvl(rsTemp!ϵͳ) & "-" & Nvl(rsTemp!ģ��)
    txtEdit(idx_������) = Nvl(rsTemp!������)
    txtEdit(idx_��������) = Nvl(rsTemp!������)
    lblEdit(4) = Nvl(rsTemp!Ӱ�����˵��)
    txtEdit(idx_�䶯��) = mstrUserName
    txtEdit(idx_�䶯ʱ��) = strCurDate
    txtEdit(idx_ԭ��������) = GetParaType(Val(Nvl(rsTemp!ģ��)), Val(Nvl(rsTemp!˽��)), Val(Nvl(rsTemp!����)))
    txtEdit(idx_ԭ��Ȩ��ʽ) = IIf(Val(Nvl(rsTemp!��Ȩ)) = 0, "����Ҫ��Ȩ", "��Ҫ��Ȩ")
    txtEdit(idx_ԭ��Ȩ��ʽ).Tag = Val(Nvl(rsTemp!��Ȩ))
    If Val(Nvl(rsTemp!�̶�)) = 1 Then
        MsgBox "�ò���Ϊϵͳ�̶����������ܵ�����", vbOKOnly, gstrSysName
        mblnNotClick = False
        Exit Function
    End If
    mblnNotClick = False
    LoadParaInfor = True
End Function
Private Function SaveData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:������صı䶯����
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-02-19 15:01:37
    '-----------------------------------------------------------------------------------------------------------
    err = 0: On Error GoTo errHand:
    Dim int˽�� As Integer, int���� As Integer, int��Ȩ As Integer
    Select Case cboEdit(0).Text
    Case "����ȫ��", "˽��ȫ��"
        MsgBox "ע��:" & vbCrLf & _
               "   �������ܱ䶯Ϊ����ȫ�ֺ�˽��ȫ��,���飡", vbOKOnly, gstrSysName
        Exit Function
    Case "����ģ��"
        int˽�� = 0: int���� = 0
    Case "˽��ģ��"
        int˽�� = 1: int���� = 0
    Case "��������ģ��"
         int˽�� = 0: int���� = 1
    Case "����˽��ģ��"
         int˽�� = 1: int���� = 1
    Case ""
        MsgBox "ע��:" & vbCrLf & _
               "   �������ܱ䶯Ϊ��,���飡", vbOKOnly, gstrSysName
        Exit Function
    End Select
    If cboEdit(1).Text = "����Ҫ��Ȩ" Then
        int��Ȩ = 0
    Else
        int��Ȩ = 1
    End If
    SaveData = False
    'zl_Parameters_Change
    gstrSQL = "zl_Parameters_Change("
    '  ����id_In   zlParameters.ID%Type,
    gstrSQL = gstrSQL & "" & mlng����ID & ","
    '  ˽��_In     zlParameters.˽��%Type,
    gstrSQL = gstrSQL & "" & int˽�� & ","
    '  ����_In     zlParameters.����%Type,
    gstrSQL = gstrSQL & "" & int���� & ","
    '  ��Ȩ_In     zlParameters.��Ȩ%Type,
    gstrSQL = gstrSQL & "" & int��Ȩ & ","
    '  �䶯��_In   Zlparachangedlog.�䶯��%Type,
    gstrSQL = gstrSQL & "'" & txtEdit(mTxt_idx.idx_�䶯��).Text & "',"
    '  �䶯ԭ��_In Zlparachangedlog.�䶯ԭ��%Type
    gstrSQL = gstrSQL & "'" & txtEdit(mTxt_idx.idx_�䶯ԭ��).Text & "')"
    ExecuteProcedure gstrSQL, Me.Caption
    SaveData = True
    Exit Function
errHand:
    MsgBox "ע��:" & vbCrLf & _
           "   ��������ʱ�������󣬴�����Ϣ���£�" & vbCrLf & _
           "������Ϣ:" & err.Number & "-" & err.Description, vbOKOnly, gstrSysName
End Function
Private Sub InitCombox()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ��Combox��Ϣ
    '����:���˺�
    '����:2009-02-19 14:43:55
    '-----------------------------------------------------------------------------------------------------------
    mblnNotClick = True
    With cboEdit(0)
        .AddItem "����ģ��"
        .ItemData(.NewIndex) = 2
        If Trim(txtEdit(mTxt_idx.idx_ԭ��������).Text) = "����ģ��" Then .ListIndex = .NewIndex
        .AddItem "˽��ģ��"
        .ItemData(.NewIndex) = 0
        If Trim(txtEdit(mTxt_idx.idx_ԭ��������).Text) = "˽��ģ��" Then .ListIndex = .NewIndex
        .AddItem "��������ģ��"
        .ItemData(.NewIndex) = 1
        If Trim(txtEdit(mTxt_idx.idx_ԭ��������).Text) = "��������ģ��" Then .ListIndex = .NewIndex
        .AddItem "����˽��ģ��"
        .ItemData(.NewIndex) = 0
        If Trim(txtEdit(mTxt_idx.idx_ԭ��������).Text) = "����˽��ģ��" Then .ListIndex = .NewIndex
    End With
    
    With cboEdit(1)
        .AddItem "����Ҫ��Ȩ"
        If Trim(txtEdit(mTxt_idx.idx_ԭ��Ȩ��ʽ).Text) = "����Ҫ��Ȩ" Then .ListIndex = .NewIndex
        .AddItem "��Ҫ��Ȩ"
        If Trim(txtEdit(mTxt_idx.idx_ԭ��Ȩ��ʽ).Text) = "��Ҫ��Ȩ" Then .ListIndex = .NewIndex
    End With
    mblnNotClick = False
End Sub
Private Function IsValied() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:�����������ݵķ���
    '����:�Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-02-19 14:55:14
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Parameter", mlng����ID)
    'ID, ϵͳ,ģ��,˽��,������,������, ����ֵ, ȱʡֵ, ����˵��, ����, ��Ȩ, �̶�,ģ������
    If rsTemp.EOF Then
        MsgBox "����δ�ҵ��������Ѿ�������ɾ��,����!", vbOKOnly, gstrSysName
        Exit Function
    End If
    If Val(Nvl(rsTemp!�̶�)) = 1 Then
        MsgBox "�ò���Ϊϵͳ�̶����������ܵ�����", vbOKOnly, gstrSysName
        Exit Function
    End If
    If ActualLen(txtEdit(mTxt_idx.idx_�䶯ԭ��).Text) > 200 Then
        MsgBox "�䶯ԭ�����������200���ַ���100�����֣����ܵ�����", vbOKOnly, gstrSysName
        txtEdit(mTxt_idx.idx_�䶯ԭ��).SetFocus
        Exit Function
    End If
    If InStr(1, txtEdit(mTxt_idx.idx_�䶯ԭ��).Text, "'") > 0 Then
        MsgBox "�䶯ԭ���зǷ��ַ������ţ����飡", vbOKOnly, gstrSysName
        txtEdit(mTxt_idx.idx_�䶯ԭ��).SetFocus
        Exit Function
    End If
    IsValied = True
End Function

Private Sub SetCtlEnbaled()
    '-----------------------------------------------------------------------------------------------------------
    '����:������ؿؼ�����
    '����:���˺�
    '����:2009-02-19 14:49:28
    '-----------------------------------------------------------------------------------------------------------
    Dim blnOk As Boolean
    mblnNotClick = True
    With cboEdit(0)
        Select Case .ItemData(.ListIndex)
        Case 1   '���Ըı���Ȩ
            cboEdit(1).Enabled = True
        Case 2   'ǿ��Ϊ��Ȩ
            cboEdit(1).Enabled = False
            cboEdit(1).ListIndex = 1
        Case Else   '������Ȩ
            cboEdit(1).Enabled = False
            cboEdit(1).ListIndex = 0
        End Select
    End With
    mblnNotClick = False
    
    blnOk = Trim(txtEdit(mTxt_idx.idx_ԭ��������)) <> Trim(cboEdit(0).Text)
    blnOk = blnOk Or Trim(txtEdit(mTxt_idx.idx_ԭ��Ȩ��ʽ)) <> Trim(cboEdit(1).Text)
    cmdOK.Enabled = blnOk
End Sub
Private Sub cboEdit_Click(Index As Integer)
    If mblnNotClick Then Exit Sub
    Call SetCtlEnbaled
End Sub

Private Sub cmdClear_Click()
    Unload Me
End Sub
Private Sub cmdOK_Click()
    If IsValied = False Then Exit Sub
    If SaveData = False Then Exit Sub
    mblnOK = True
    Unload Me
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    Call InitCombox
    If LoadParaInfor() = False Then
        Unload Me: Exit Sub
    End If
    Call SetCtlEnbaled
    mblnChange = False
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    SendKeys "{tab}"
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If mblnNotClick = True Then Exit Sub

    mblnChange = True
    Call SetCtlEnbaled
End Sub

Private Sub txtEdit_Click(Index As Integer)
    Call SetCtlEnbaled
    mblnChange = True
End Sub
