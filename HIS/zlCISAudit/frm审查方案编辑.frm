VERSION 5.00
Begin VB.Form frm��鷽���༭ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���ַ����༭"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5670
   Icon            =   "frm��鷽���༭.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txt˵�� 
      Height          =   900
      Left            =   1185
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1395
      Width           =   3885
   End
   Begin VB.TextBox txt�ֶ��� 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1185
      MaxLength       =   10
      TabIndex        =   2
      ToolTipText     =   "�緽���ܷ�>=90Ϊ�ϸ�"
      Top             =   975
      Width           =   2040
   End
   Begin zl9CISAudit.tipPopup tipPopup1 
      Height          =   420
      Left            =   1890
      Top             =   2610
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4080
      TabIndex        =   6
      Top             =   2970
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2775
      TabIndex        =   5
      Top             =   2970
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   465
      TabIndex        =   7
      Top             =   2970
      Width           =   1100
   End
   Begin VB.TextBox txt�ܷ� 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1185
      MaxLength       =   10
      TabIndex        =   1
      Text            =   "100"
      ToolTipText     =   "�����۷�ʱ��߷�"
      Top             =   576
      Width           =   2040
   End
   Begin VB.CheckBox chkѡ�� 
      Caption         =   "ѡ��(&S)"
      Height          =   285
      Left            =   1185
      TabIndex        =   4
      Top             =   2340
      Width           =   1095
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1185
      LinkTimeout     =   25
      MaxLength       =   25
      TabIndex        =   0
      Top             =   135
      Width           =   3870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "˵��(&B)"
      Height          =   180
      Left            =   495
      TabIndex        =   13
      Top             =   1410
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      Height          =   180
      Left            =   3300
      TabIndex        =   12
      Top             =   1020
      Width           =   180
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   11520
      Y1              =   2745
      Y2              =   2745
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      X1              =   15
      X2              =   11520
      Y1              =   2775
      Y2              =   2775
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      Height          =   180
      Left            =   3300
      TabIndex        =   11
      Top             =   636
      Width           =   180
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ܷ�(&M)"
      Height          =   180
      Left            =   480
      TabIndex        =   10
      Top             =   636
      Width           =   630
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�ֶ���(&F)"
      Height          =   180
      Left            =   300
      TabIndex        =   9
      Top             =   1020
      Width           =   810
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����(&N)"
      Height          =   180
      Left            =   480
      TabIndex        =   8
      Top             =   195
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   900
      Left            =   15
      Picture         =   "frm��鷽���༭.frx":000C
      Top             =   1740
      Width           =   900
   End
End
Attribute VB_Name = "frm��鷽���༭"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private m_lngID                 As Long     '��ǰ�༭�����ַ�����ID��
Private m_strEditMode           As String   '���ڱ༭ģʽ��Add.���� Mod.�޸�
Private m_blnModed              As Boolean
Private zlCheck                 As New clsCheck

Public Property Get Moded() As Boolean
   Moded = m_blnModed
End Property

Public Property Let Moded(ByVal blnModed As Boolean)
    m_blnModed = blnModed
End Property

'==============================================================================
'=���ܣ� �����ӿں��������ڴ����ʼ������:ID��ID��ֵΪ0��ʾ��ӱ�׼��
'==============================================================================
Public Sub ShowForm(Optional ID As Long = 0)
    On Error GoTo ErrH
    
    m_lngID = ID          'Ϊ0��ʾ����
    m_blnModed = False
    
    If ID <= 0 Then
        m_strEditMode = "Add"
        Me.Caption = "��������"
    Else
        m_strEditMode = "Mod"
        Me.Caption = "�޸ķ���"
        FillInitData
    End If
    Me.Show 1
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ����ID�����ʼ����
'==============================================================================
Private Sub FillInitData()
    Dim rs      As ADODB.Recordset
    
    On Error GoTo ErrH
    
    gstrSQL = "Select ����,�ܷ�,�ֶ���,����ʱ��,ͣ��ʱ��,˵�� From ������鷽�� where ID = [1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, m_lngID)
    If Not rs.EOF Then
        txt���� = IIf(IsNull(rs.Fields("����")), "", rs.Fields("����"))
        txt�ܷ� = NVL(rs.Fields("�ܷ�"), 0)
        txt�ֶ��� = NVL(rs.Fields("�ֶ���"), 0)
        txt˵�� = NVL(rs.Fields("˵��"))
         
        If NVL(rs.Fields("����ʱ��")) = "" Then
            chkѡ��.Value = vbUnchecked
        Else
            chkѡ��.Value = vbChecked
        End If
        zlControl.TxtSelAll txt����

    Else
        Unload Me
        MsgBox "��ʼ�����ݴ���û�з��ָ������ַ����������ԡ�", vbOKOnly + vbInformation, "��������"
        Exit Sub
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


'==============================================================================
'=���ܣ� ѡ����ʾ
'==============================================================================
Private Sub chkѡ��_GotFocus()
    On Error GoTo ErrH
    ShowTips chkѡ��, "���ֻ����һ��������ѡ�á�", "ѡ��"
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ѡ�ûس��൱��ȷ��
'==============================================================================
Private Sub chkѡ��_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrH
    If KeyAscii = vbKeyReturn Then
        Call CmdOK_Click
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ȡ���༭
'==============================================================================
Private Sub CmdCancel_Click()
    On Error GoTo ErrH
    Moded = False
    Unload Me
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� �������
'==============================================================================
Private Sub cmdHelp_Click()
    On Error GoTo ErrH
    ShowHelp App.ProductName, Me.hWnd, Me.Name, 3
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�ȷ����������
'==============================================================================
Private Sub CmdOK_Click()
    Dim strT As String
    
    On Error GoTo ErrH
    
    If m_strEditMode = "Add" Then
        strT = "ZL_������鷽��_Insert"
        gstrSQL = strT & _
                "(" & zlDatabase.GetNextId("������鷽��") & ",'" & txt���� & "'," & txt�ܷ�.Text & "," & txt�ֶ���.Text & "," & CStr(IIf(chkѡ��.Value = vbChecked, 1, 0)) & _
                ",'" & txt˵��.Text & "'" & _
                ")"
    Else
        strT = "ZL_������鷽��_Update"
        gstrSQL = strT & _
                       "(" & m_lngID & ",'" & txt���� & "'," & txt�ܷ�.Text & "," & txt�ֶ���.Text & "," & CStr(IIf(chkѡ��.Value = vbChecked, 1, 0)) & _
                       ",'" & txt˵��.Text & "'" & _
                       ")"
    End If
    '�������Ϸ���
    If IsValid() = False Then Exit Sub
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Moded = True
    MsgBox "���ַ�������ɹ���", vbOKOnly + vbInformation, gstrSysName
    Unload Me
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�����������Ŀ�������Ƿ���Ч
'=���أ���Ч����True,����ΪFalse
'==============================================================================
Private Function IsValid() As Boolean
    On Error GoTo ErrH
    '�����ֶμ��
    IsValid = False
    '����StrIsValid������ȷ���ַ�����ʽ��ȷ��ע�⣺����ʹ�õ���lenBֵ����Ӧ���ݱ����е�ֵ��
    If Len(Trim(txt����)) = 0 Then
        MsgBox "�����뷽�����ƣ�", vbInformation, gstrSysName
        zlControl.TxtSelAll txt����: txt����.SetFocus
        Exit Function
    End If
    If zlCommFun.StrIsValid(txt����.Text, txt����.MaxLength * 2) = False Then
        MsgBox "���������̫�������������룡", vbInformation, gstrSysName
        zlControl.TxtSelAll txt����: txt����.SetFocus
        Exit Function
    End If
    If Len(Trim(txt�ܷ�)) = 0 Then
        MsgBox "�����뷽���ܷ֣�", vbInformation, gstrSysName
        zlControl.TxtSelAll txt�ܷ�: txt�ܷ�.SetFocus
        Exit Function
    End If

    If Len(Trim(txt�ܷ�)) > 0 Then
        If Not IsNumeric(txt�ܷ�) Then
            MsgBox "��������Ч�ķ����ܷ�ֵ��", vbInformation, gstrSysName
            zlControl.TxtSelAll txt�ܷ�: txt�ܷ�.SetFocus
            Exit Function
        End If
        If Val(txt�ܷ�.Text) > 9999# Then
            MsgBox "�����ܷ������������̫��", vbInformation, gstrSysName
            zlControl.TxtSelAll txt�ܷ�: txt�ܷ�.SetFocus
            Exit Function
        End If
    End If

    If Len(Trim(txt�ֶ���)) > 0 Then
        If Not IsNumeric(txt�ֶ���) Then
            MsgBox "��������ȷ�ķֶ�����ֵ��", vbInformation, gstrSysName
            zlControl.TxtSelAll txt�ֶ���: txt�ֶ���.SetFocus
            Exit Function
        End If
        If Val(txt�ֶ���.Text) > 9999# Or Val(txt�ֶ���.Text) > Val(txt�ܷ�.Text) Then
            MsgBox "�ֶ��������������̫��", vbInformation, gstrSysName
            zlControl.TxtSelAll txt�ֶ���: txt�ֶ���.SetFocus
            Exit Function
        End If
        If Val(txt�ֶ���.Text) < 0 Then
            MsgBox "�ֶ��������������̫С��", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    IsValid = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=���ܣ�ҳ���ʼ��
'==============================================================================
Private Sub Form_Initialize()
    On Error GoTo ErrH
    Call InitCommonControls
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�ҳ���ʼ��
'==============================================================================
Private Sub Form_Load()
    On Error GoTo ErrH
    zlCheck.Sys_System Me
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Sub

'==============================================================================
'=���ܣ����ƶ� ' | ¼��Ŀ���
'==============================================================================
Private Sub txt����_Change()
    On Error GoTo ErrH
    If InStr(txt����, "'") <> 0 Then txt���� = Replace(txt����, "'", "")
    If InStr(txt����, "|") <> 0 Then txt���� = Replace(txt����, "|", "")
    txt����.SelStart = Len(txt����)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ����Ƶõ�������ʾ
'==============================================================================
Private Sub txt����_GotFocus()
    On Error GoTo ErrH
    zlControl.TxtSelAll txt����
    Call zlCommFun.OpenIme(True)
    ShowTips txt����, "���뷽�����ƣ�������25���ַ��ڡ�", "��������"
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�����ʧȥ������
'==============================================================================
Private Sub txt����_LostFocus()
    On Error GoTo ErrH
    Call zlCommFun.OpenIme(False)
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ�����ȡ����ʾ
'==============================================================================
Private Sub txt����_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error GoTo ErrH
    tipPopup1.Hide
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ���ֵ�õ�������ʾ
'==============================================================================
Private Sub txt�ֶ���_GotFocus()
    On Error GoTo ErrH
    zlControl.TxtSelAll txt�ֶ���
    ShowTips txt�ֶ���, "�׼������ֶ���", "�ֶ���ֵ", 5000
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ���ֵ����F1��ʾ
'==============================================================================
Private Sub txt�ֶ���_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrH
    If KeyCode = vbKeyF1 Then
        ShowTips txt�ֶ���, "��:�ֶ���>=90Ϊ�ϸ�,����Ϊ���ϸ񣬱�����Ժ��24Сʱδ��д��Ժ������"
    End If
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


'==============================================================================
'=���ܣ� �ֵܷõ�������ʾ
'==============================================================================
Private Sub txt�ܷ�_GotFocus()
    On Error GoTo ErrH
    zlControl.TxtSelAll txt�ܷ�
    ShowTips txt�ܷ�, "���뷽����׼�ܷ֣�Ĭ��Ϊ100�֡�", "�����ܷ�"
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=���ܣ� ��ʾ
'==============================================================================
Private Sub ShowTips(ctl As Control, str���� As String, Optional str���� As String = "��ʾ��Ϣ", Optional lngʱ�� As Long = 2500, Optional ������ʾ As Boolean = False)
    Dim X       As Single
    Dim Y       As Single
    On Error GoTo ErrH
    
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
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
