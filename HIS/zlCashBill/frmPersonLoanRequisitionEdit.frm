VERSION 5.00
Begin VB.Form frmPersonLoanRequisitionEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�������"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7590
   Icon            =   "frmPersonLoanRequisitionEdit.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   7590
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   420
      Left            =   6225
      TabIndex        =   18
      Top             =   4365
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   420
      Left            =   4875
      TabIndex        =   17
      Top             =   4365
      Width           =   1200
   End
   Begin VB.Frame fraSplit 
      Height          =   120
      Index           =   1
      Left            =   -45
      TabIndex        =   21
      Top             =   3990
      Width           =   7830
   End
   Begin VB.Frame fraSplit 
      Height          =   120
      Index           =   0
      Left            =   30
      TabIndex        =   20
      Top             =   900
      Width           =   7830
   End
   Begin VB.Frame fra���� 
      BorderStyle     =   0  'None
      Height          =   2940
      Left            =   105
      TabIndex        =   19
      Top             =   1035
      Width           =   7530
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   6
         Left            =   870
         TabIndex        =   4
         Top             =   810
         Width           =   2490
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   5
         Left            =   870
         TabIndex        =   16
         Top             =   2430
         Width           =   6375
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   4
         Left            =   870
         TabIndex        =   14
         Top             =   2025
         Width           =   2490
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   870
         TabIndex        =   8
         Top             =   1230
         Width           =   6330
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   4695
         TabIndex        =   12
         Top             =   1620
         Width           =   2490
      End
      Begin VB.ComboBox cbo����� 
         Height          =   300
         Left            =   870
         TabIndex        =   10
         Text            =   "cbo�����"
         Top             =   1605
         Width           =   2490
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   4695
         TabIndex        =   6
         Top             =   765
         Width           =   2490
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   870
         TabIndex        =   2
         Top             =   390
         Width           =   2490
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�����"
         Height          =   180
         Index           =   7
         Left            =   75
         TabIndex        =   3
         Top             =   870
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ȡ��ԭ��"
         Height          =   180
         Index           =   6
         Left            =   75
         TabIndex        =   15
         Top             =   2520
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ȡ��ʱ��"
         Height          =   180
         Index           =   5
         Left            =   75
         TabIndex        =   13
         Top             =   2085
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��ע"
         Height          =   180
         Index           =   4
         Left            =   435
         TabIndex        =   7
         Top             =   1260
         Width           =   360
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "���ʱ��"
         Height          =   180
         Index           =   3
         Left            =   3855
         TabIndex        =   11
         Top             =   1680
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�����"
         Height          =   240
         Index           =   2
         Left            =   255
         TabIndex        =   9
         Top             =   1650
         Width           =   540
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��"
         Height          =   180
         Index           =   1
         Left            =   3855
         TabIndex        =   5
         Top             =   825
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�����"
         Height          =   180
         Index           =   0
         Left            =   255
         TabIndex        =   1
         Top             =   450
         Width           =   540
      End
   End
   Begin VB.Label lbl 
      Caption         =   $"frmPersonLoanRequisitionEdit.frx":058A
      Height          =   885
      Left            =   1035
      TabIndex        =   0
      Top             =   135
      Width           =   6435
   End
   Begin VB.Image img 
      Height          =   720
      Left            =   135
      Picture         =   "frmPersonLoanRequisitionEdit.frx":0675
      Top             =   150
      Width           =   720
   End
End
Attribute VB_Name = "frmPersonLoanRequisitionEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnFirst As Boolean, mblnChange As Boolean, mlngID As Long
Private mlngModule As Long, mstrPrivs As String
Private mblnSucceed As Boolean '������־:true,�ɹ�,����False
Private mrs��Ա As ADODB.Recordset

Public Enum gEditTypeLoan
    FN_���� = 0
    FN_�޸� = 1
    FN_��� = 2
    FN_ȡ����� = 3
    FN_��ѯ = 4
End Enum

Private mEditType As gEditTypeLoan
Private Enum mIdxTxt
    idx_����� = 0
    idx_����ʱ�� = 1
    idx_���ʱ�� = 2
    idx_��ע = 3
    idx_ȡ��ʱ�� = 4
    idx_ȡ��ԭ�� = 5
    idx_����� = 6
End Enum
Public Function ShowEdit(ByVal frmMain As Form, ByVal EditType As gEditTypeLoan, ByVal strPrivs As String, ByVal lngModule As Long, Optional lngID As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����༭���:��ʾ��ʾ������Ȳ���
    '���:frmMain-������
    '     EditType-�༭����
    '     strPrivs-Ȩ�޴�
    '     lngModule��ģ���
    '     lngID-������ʱ����Ч,������ID
    '����:
    '����:�����ɹ�������ture,���򷵻�False
    '����:���˺�
    '����:2009-09-08 11:54:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs: mEditType = EditType: mblnSucceed = False: mlngID = lngID
    Me.Show 1, frmMain
    ShowEdit = mblnSucceed
End Function

Private Function CheckDepend() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����ص�������ϵ
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-09-08 12:00:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    gstrSQL = "" & _
    "   Select Distinct B.ID, B.���, B.����, B.����, B.����, B.��������, B.�Ա�, B.�칫�ҵ绰 " & _
    "   From ��Ա����˵�� A, ��Ա�� B " & _
    "   Where A.��Աid = B.ID And A.��Ա���� In ('����Һ�Ա', '�����շ�Ա', 'Ԥ���տ�Ա', 'סԺ����Ա') " & _
    "   Order By ���"
    Set mrs��Ա = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If mrs��Ա.RecordCount = 0 Then
        MsgBox "ע�⣺" & vbCrLf & _
               "����û��һ����ԱΪ������Һ�Ա�������շ�Ա��Ԥ���տ�Ա��סԺ����Ա��" & vbCrLf & _
               " ���ڡ���Ա���������ã�", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If mEditType = FN_���� Or mEditType = FN_�޸� Then
        With cbo�����
            .Clear: mrs��Ա.MoveFirst
            Do While Not mrs��Ա.EOF
                If Nvl(mrs��Ա!����) <> UserInfo.���� Then
                    .AddItem Nvl(mrs��Ա!����)
                    .ItemData(.NewIndex) = Nvl(Val(mrs��Ա!ID))
                End If
                mrs��Ա.MoveNext
            Loop
        End With
    End If
    CheckDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub SetCtrolEnabled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿؼ��ı༭����
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-09-08 14:51:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
 
    txtEdit(mIdxTxt.idx_�����).Enabled = False
    txtEdit(mIdxTxt.idx_����ʱ��).Enabled = False
    txtEdit(mIdxTxt.idx_���ʱ��).Enabled = False
    txtEdit(mIdxTxt.idx_��ע).Enabled = False
    txtEdit(mIdxTxt.idx_ȡ��ʱ��).Enabled = False
    txtEdit(mIdxTxt.idx_ȡ��ԭ��).Enabled = False
    txtEdit(mIdxTxt.idx_�����).Enabled = False
    cbo�����.Enabled = False
    Select Case mEditType
    Case gEditTypeLoan.FN_����, gEditTypeLoan.FN_�޸�
        txtEdit(mIdxTxt.idx_��ע).Enabled = True
        txtEdit(mIdxTxt.idx_�����).Enabled = True
        cbo�����.Enabled = True
    Case gEditTypeLoan.FN_ȡ�����
        txtEdit(mIdxTxt.idx_ȡ��ԭ��).Enabled = True
    Case Else
    End Select
End Sub
Private Sub ClearCtrl()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '����:���˺�
    '����:2009-09-08 14:27:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim ctl As Control
    For Each ctl In Me.Controls
        Select Case UCase(TypeName(ctl))
        Case "TEXTBOX"
            ctl.Text = ""
            If mEditType = FN_���� Then
                If ctl.Index = mIdxTxt.idx_����� Then ctl.Text = UserInfo.����
                If ctl.Index = mIdxTxt.idx_����ʱ�� Then ctl.Text = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
            End If
        Case Else
        End Select
        Call zlSetCrlEnbled(ctl, ctl.Enabled)
    Next
End Sub

Private Function LoadDataToCtrol() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ݵ��ؼ�
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-09-08 14:26:36
    '---------------------------------------------------------------------------------------------------------------------------------------------\
    Dim rsTemp As ADODB.Recordset, i As Long
    
    Err = 0: On Error GoTo ErrHand:
    Call ClearCtrl
    If mEditType = FN_���� Then LoadDataToCtrol = True: Exit Function
    
  
    gstrSQL = " " & _
    "    Select Id, �����, ��ע, �����, to_char(����ʱ��,'yyyy-mm-dd hh24:mi:ss') as ����ʱ�� ,  " & _
    "           �����, to_char(���ʱ��,'yyyy-mm-dd hh24:mi:ss') as ���ʱ��, " & _
    "           to_char(ȡ��ʱ��,'yyyy-mm-dd hh24:mi:ss') as ȡ��ʱ��, ȡ��ԭ�� " & _
    "    From ��Ա����¼ " & _
    "    Where ID=[1] " & _
    "    Order by �����,����ʱ��"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngID)
    If rsTemp.RecordCount = 0 Then
        MsgBox "ע�⣺" & vbCrLf & _
               "    �ý���¼�����Ѿ�������ɾ�������ܼ�������!", vbOKOnly + vbDefaultButton1 + vbInformation, gstrSysName
        Exit Function
    End If
    
    txtEdit(mIdxTxt.idx_�����).Text = Nvl(rsTemp!�����)
    txtEdit(mIdxTxt.idx_����ʱ��).Text = Nvl(rsTemp!����ʱ��)
    txtEdit(mIdxTxt.idx_��ע).Text = Nvl(rsTemp!��ע)
    txtEdit(mIdxTxt.idx_�����).Text = Format(Val(Nvl(rsTemp!�����)), "####0.00;-###0.00;;")
    
    If mEditType = FN_�޸� Then
        If Nvl(rsTemp!���ʱ��) <> "" Then
            MsgBox "ע�⣺" & vbCrLf & _
                   "    �ý���¼�Ѿ������˽���������ٽ����޸Ĳ���!", vbOKOnly + vbDefaultButton1 + vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If mEditType = FN_��� Then
        If Nvl(rsTemp!���ʱ��) <> "" Then
            MsgBox "ע�⣺" & vbCrLf & _
                   "    �ý���¼�Ѿ������˽���������ٽ��н������!", vbOKOnly + vbDefaultButton1 + vbInformation, gstrSysName
            Exit Function
        End If
        txtEdit(mIdxTxt.idx_���ʱ��).Text = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    Else
        txtEdit(mIdxTxt.idx_���ʱ��).Text = Nvl(rsTemp!���ʱ��)
    End If
    
    If mEditType = FN_ȡ����� Then
        If Trim(Nvl(rsTemp!���ʱ��)) = "" Then
            MsgBox "ע�⣺" & vbCrLf & _
                   "    �ý���¼��δȷ�Ͻ�������ܽ���ȡ���������!", vbOKOnly + vbDefaultButton1 + vbInformation, gstrSysName
            Exit Function
        End If
        If Trim(Nvl(rsTemp!ȡ��ʱ��)) <> "" Then
            MsgBox "ע�⣺" & vbCrLf & _
                   "    �ý���¼�Ѿ�������ȡ���������ٽ���ȡ���������!", vbOKOnly + vbDefaultButton1 + vbInformation, gstrSysName
            Exit Function
        End If
        txtEdit(mIdxTxt.idx_ȡ��ʱ��).Text = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    Else
        txtEdit(mIdxTxt.idx_ȡ��ʱ��).Text = Nvl(rsTemp!ȡ��ʱ��)
    End If
    With cbo�����
        .ListIndex = -1
        For i = 0 To .ListCount - 1
            If .List(i) = Nvl(rsTemp!�����) Then .ListIndex = i: Exit For
        Next
        If .ListIndex < 0 Then
            .AddItem Nvl(rsTemp!�����)
            .ListIndex = .NewIndex
        End If
    End With
    
    txtEdit(mIdxTxt.idx_ȡ��ԭ��).Text = Nvl(rsTemp!ȡ��ԭ��)
    LoadDataToCtrol = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub SetDefaultInputLen()
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    gstrSQL = "Select ��ע,ȡ��ԭ��  From ��Ա����¼ where id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, -1)
    txtEdit(mIdxTxt.idx_��ע).MaxLength = rsTemp.Fields("��ע").DefinedSize
    txtEdit(mIdxTxt.idx_ȡ��ԭ��).MaxLength = rsTemp.Fields("ȡ��ԭ��").DefinedSize
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������Ƿ�Ϸ�
    '����:�Ϸ�������true,���򷵻�False
    '����:���˺�
    '����:2009-09-08 15:25:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim ctl As Object
    
    Err = 0: On Error GoTo ErrHand:
    
    Set ctl = txtEdit(mIdxTxt.idx_��ע)
    If zlCommFun.ActualLen(ctl.Text) > ctl.MaxLength Then
        MsgBox "��עֻ������" & ctl.MaxLength & " ���ַ���" & ctl.MaxLength \ 2 & "������,����!", vbInformation + vbOKOnly, gstrSysName
        Call zlcontrol.ControlSetFocus(ctl, True)
        Exit Function
    End If
    Set ctl = txtEdit(mIdxTxt.idx_ȡ��ԭ��)
    If zlCommFun.ActualLen(ctl.Text) > ctl.MaxLength Then
        MsgBox "ȡ��ԭ��ֻ������" & ctl.MaxLength & " ���ַ���" & ctl.MaxLength \ 2 & "������,����!", vbInformation + vbOKOnly, gstrSysName
        Call zlcontrol.ControlSetFocus(ctl, True)
        Exit Function
    End If
    
    Set ctl = txtEdit(mIdxTxt.idx_�����)
    If Val(ctl.Text) > 10 ^ 12 - 1 Then
        MsgBox "��������С��" & CStr(10 ^ 12 - 1) & " ,����!", vbInformation + vbOKOnly, gstrSysName
        Call zlcontrol.ControlSetFocus(ctl, True)
        Exit Function
    End If
    If Val(ctl.Text) <= 0 Then
        MsgBox "������������� ,����!", vbInformation + vbOKOnly, gstrSysName
        Call zlcontrol.ControlSetFocus(ctl, True)
        Exit Function
    End If
        
    Set ctl = txtEdit(mIdxTxt.idx_�����)
    If ctl.Text = cbo�����.Text Then
        MsgBox "�������������ͬһ��,���ܼ���!", vbInformation + vbOKOnly, gstrSysName
        Call zlcontrol.ControlSetFocus(cbo�����, True)
        Exit Function
    End If
    If Trim(ctl.Text) = "" Then
        MsgBox "δ������Ա�Ķ��չ�ϵ,����ϵͳ����Ա��ϵ!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If Trim(cbo�����.Text) = "" Or cbo�����.ListIndex < 0 Then
        MsgBox "�����δѡ��,���ܼ���!", vbInformation + vbOKOnly, gstrSysName
        Call zlcontrol.ControlSetFocus(cbo�����, True)
        Exit Function
    End If
    
    isValied = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-09-08 15:57:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngID As Long
    Err = 0: On Error GoTo ErrHand:
    If mEditType = FN_�޸� Then
        lngID = mlngID
    Else
        lngID = zlDatabase.GetNextId("��Ա����¼")
    End If
    'Zl_��Ա����¼_Insert
    gstrSQL = IIf(mEditType = FN_�޸�, "Zl_��Ա����¼_Update(", "Zl_��Ա����¼_Insert(")
    '  Id_In       In ��Ա����¼.ID%Type,
    gstrSQL = gstrSQL & "" & lngID & ","
    '  �����_In In ��Ա����¼.�����%Type,
    gstrSQL = gstrSQL & "" & Val(txtEdit(mIdxTxt.idx_�����)) & ","
    '  ��ע_In     In ��Ա����¼.��ע%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtEdit(mIdxTxt.idx_��ע)) & "',"
    '  �����_In   In ��Ա����¼.�����%Type,
    gstrSQL = gstrSQL & "'" & Trim(txtEdit(mIdxTxt.idx_�����)) & "',"
    '  ����ʱ��_In In ��Ա����¼.����ʱ��%Type,
    gstrSQL = gstrSQL & "to_date('" & Trim(txtEdit(mIdxTxt.idx_����ʱ��)) & "','yyyy-mm-dd hh24:mi:ss'),"
    '  �����_In   In ��Ա����¼.�����%Type
    gstrSQL = gstrSQL & "'" & Trim(cbo�����.Text) & "')"
    
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    mlngID = lngID
    SaveData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function SaveLoanOut(ByVal lngID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-09-08 16:28:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    'Zl_��Ա����¼_���(Id_In In ��Ա����¼.ID%Type) Is
    Err = 0: On Error GoTo ErrHand:
    gstrSQL = "Zl_��Ա����¼_���(" & lngID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    SaveLoanOut = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Function SaveCancelLoanOut() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2009-09-08 16:28:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo ErrHand:
    ' Zl_��Ա����¼_ȡ�����
    gstrSQL = "Zl_��Ա����¼_ȡ�����("
    '  Id_In       In ��Ա����¼.ID%Type,
    gstrSQL = gstrSQL & "" & mlngID & ","
    '  ȡ��ԭ��_In In ��Ա����¼.ȡ��ԭ��%Type,
    gstrSQL = gstrSQL & "'" & txtEdit(mIdxTxt.idx_ȡ��ԭ��).Text & "',"
    '  ȡ��ʱ��_In In ��Ա����¼.ȡ��ʱ��%Type
    gstrSQL = gstrSQL & "to_date('" & txtEdit(mIdxTxt.idx_ȡ��ʱ��).Text & "','yyyy-mm-dd hh24:mi:ss'))"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    SaveCancelLoanOut = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub cbo�����_GotFocus()
    zlcontrol.TxtSelAll cbo�����
    zlCommFun.OpenIme False
End Sub

Private Sub cbo�����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If cbo�����.ListIndex >= 0 Then zlCommFun.PressKey vbKeyTab: Exit Sub
    'ѡ��:
    If Select��Աѡ����(cbo�����, Trim(cbo�����.Text)) Then Exit Sub
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If mEditType = FN_��ѯ Then Unload Me: Exit Sub
    
    If isValied = False Then Exit Sub
    
    If mEditType = FN_��� Then
        If SaveLoanOut(mlngID) = False Then Exit Sub
        If IIf(Val(zlDatabase.GetPara("�����ӡ", glngSys, mlngModule, "0")) = 1, 1, 0) = 1 Then
            '��ӡ
            If InStr(mstrPrivs, "��") <> 0 Then
                ReportOpen gcnOracle, glngSys, "zl1_bill_1502", Me, "ID=" & mlngID, 2
            End If
        End If
        mblnChange = False: mblnSucceed = True: Unload Me: Exit Sub
    End If
    If mEditType = FN_ȡ����� Then
        If SaveCancelLoanOut = False Then Exit Sub
        If IIf(Val(zlDatabase.GetPara("�����ӡ", glngSys, mlngModule, "0")) = 1, 1, 0) = 1 Then
            '��ӡ
            If InStr(mstrPrivs, "��") <> 0 Then
                ReportOpen gcnOracle, glngSys, "zl1_bill_1502", Me, "ID=" & mlngID, 2
            End If
        End If
        mblnChange = False: mblnSucceed = True: Unload Me: Exit Sub
    End If
    
    If SaveData = False Then Exit Sub
    
    If IIf(Val(zlDatabase.GetPara("�����ӡ", glngSys, mlngModule, "0")) = 1, 1, 0) = 1 Then
        '��ӡ
        If InStr(mstrPrivs, "��") <> 0 Then
            ReportOpen gcnOracle, glngSys, "zl1_bill_1502", Me, "ID=" & mlngID, 2
        End If
    End If
    
    mblnSucceed = True: mblnChange = False
    If mEditType = FN_�޸� Then Unload Me: Exit Sub
    Call ClearCtrl
    zlcontrol.ControlSetFocus txtEdit(mIdxTxt.idx_�����), True
    mlngID = 0
    mblnSucceed = True: mblnChange = False
End Sub


Private Sub cbo�����_Change()
    mblnChange = True
    
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If CheckDepend = False Then Unload Me: Exit Sub
    Call SetDefaultInputLen
    mblnChange = False
    Call SetCtrolEnabled
    '��������
    If LoadDataToCtrol = False Then
        mblnChange = False
        Unload Me: Exit Sub
    End If
    If mEditType = FN_�޸� Or mEditType = FN_���� Then
        zlcontrol.ControlSetFocus txtEdit(mIdxTxt.idx_�����), True
        cbo�����.SelLength = 0
    ElseIf mEditType = FN_ȡ����� Then
        zlcontrol.ControlSetFocus txtEdit(mIdxTxt.idx_ȡ��ԭ��), True
    Else
        zlcontrol.ControlSetFocus cmdOK, True
    End If
    mblnChange = False
End Sub
Private Sub Form_Load()
    mblnFirst = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Then
        If MsgBox("���ݿ����Ѹı䣬��δ���̣���Ҫ�˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If Index <> mIdxTxt.idx_����� Then
        mblnChange = True
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlcontrol.TxtSelAll txtEdit(Index)
    Select Case Index
    Case mIdxTxt.idx_��ע, mIdxTxt.idx_ȡ��ԭ��
        zlCommFun.OpenIme True
    Case Else
        zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab: Exit Sub
    mblnChange = True
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = mIdxTxt.idx_����� Then
        zlcontrol.TxtCheckKeyPress txtEdit(Index), KeyAscii, m���ʽ
    Else
        zlcontrol.TxtCheckKeyPress txtEdit(Index), KeyAscii, m�ı�ʽ
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
     zlCommFun.OpenIme False
     If Index = mIdxTxt.idx_����� Then
        txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), "###0.00;-###0.00;;")
     End If
End Sub

Private Function Select��Աѡ����(ByVal objCtl As Control, ByVal strSearch As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����ѡ����
    '���::objCtl-ָ���ؼ�
    '     strSearch-Ҫ����������
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-05-01 14:18:58
    '-----------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, strKey As String, strTittle As String, lngH As Long, strTemp As String
    Dim vRect As RECT
    Dim rsTemp  As ADODB.Recordset
    Dim i As Long
    'zlDatabase.ShowSelect
    '���ܣ�
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    
    strTittle = "��Աѡ����"
    vRect = zlcontrol.GetControlRect(objCtl.hwnd)
    lngH = objCtl.Height
    
  
    gstrSQL = "" & _
    "   Select Distinct B.ID, B.���, B.����, B.����, B.����, B.��������, B.�Ա�, B.�칫�ҵ绰 " & _
    "   From ��Ա����˵�� A, ��Ա�� B " & _
    "   Where A.��Աid = B.ID And A.��Ա���� In ('����Һ�Ա', '�����շ�Ա', 'Ԥ���տ�Ա', 'סԺ����Ա') and B.id <>[2] " & _
    "         and (b.��� like upper([1]) or b.���� like [1] or b.���� like upper([1]) or b.���� like [1]) " & _
    "   Order By b.���"
    
    strKey = GetMatchingSting(strSearch, False)
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, strKey, UserInfo.ID)
 
    If blnCancel = True Then
       zlcontrol.ControlSetFocus objCtl, True
        Exit Function
    End If
    
    If rsTemp Is Nothing Then
        ShowMsgbox "û��������������Ա��Ϣ,����!"
       zlcontrol.ControlSetFocus objCtl, True
        Exit Function
    End If
    zlcontrol.ControlSetFocus objCtl, True
    Dim blnHaveData As Boolean
    With objCtl
        For i = 0 To .ListCount - 1
            If Nvl(rsTemp!����) = .List(i) Then
                .ListIndex = i: Exit For
                blnHaveData = True
            End If
        Next
        If blnHaveData = False Then
            .AddItem Nvl(rsTemp!����)
            .ItemData(.NewIndex) = Val(Nvl(rsTemp!ID)): .ListIndex = .NewIndex
        End If
    End With
    zlCommFun.PressKey vbKeyTab
    Select��Աѡ���� = True
End Function

