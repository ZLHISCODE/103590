VERSION 5.00
Begin VB.Form frmDepositNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����Ԥ��תסԺԤ��"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5760
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txtPrePay 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2190
      TabIndex        =   9
      Text            =   "0.00"
      Top             =   870
      Width           =   2355
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4230
      TabIndex        =   13
      Top             =   1770
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2940
      TabIndex        =   12
      Top             =   1770
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   330
      TabIndex        =   11
      Top             =   1770
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   45
      Left            =   30
      TabIndex        =   14
      Top             =   540
      Width           =   5775
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "סԺ�ţ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   210
      Index           =   4
      Left            =   3840
      TabIndex        =   6
      Top             =   180
      Width           =   840
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2145"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   14
      Left            =   4650
      TabIndex        =   7
      Top             =   180
      Width           =   420
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "12��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   13
      Left            =   3120
      TabIndex        =   5
      Top             =   180
      Width           =   420
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���䣺"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   210
      Index           =   3
      Left            =   2580
      TabIndex        =   4
      Top             =   180
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   12
      Left            =   2100
      TabIndex        =   3
      Top             =   180
      Width           =   210
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�Ա�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   210
      Index           =   2
      Left            =   1500
      TabIndex        =   2
      Top             =   180
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����С"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Index           =   11
      Left            =   780
      TabIndex        =   1
      Top             =   180
      Width           =   630
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ˣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   630
   End
   Begin VB.Label lblPrePay 
      AutoSize        =   -1  'True
      Caption         =   "���(&T):"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1050
      TabIndex        =   8
      Top             =   900
      Width           =   1020
   End
   Begin VB.Label lblremark 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "˵��:"
      Height          =   180
      Left            =   2580
      TabIndex        =   10
      Top             =   1320
      Width           =   465
   End
End
Attribute VB_Name = "frmDepositNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mdblPrePay As Double
Private mstrRemark As String
Private mblnOK As Boolean

Private Enum idx_Lable
    lblName = 1
    txtName = 11
    lblSex = 2
    txtSex = 12
    lblAge = 3
    txtAge = 13
    lblInNumber = 4
    txtInNumber = 14
End Enum

Private mpatiInfo As clsPatientInfo '������Ϣ

Public Function ShowMe(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '����:�����ﲡ������Ԥ��תסԺԤ��
    '���:
    '   lng����ID - ����ID
    '   lng��ҳID - ��ҳID
    '����:�ɹ�����True,���򷵻�False
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mblnOK = False
    
    On Error Resume Next
    Me.Show vbModal
    ShowMe = mblnOK

End Function

Private Sub Form_Load()
    Dim strData As String
    
    zlCommFun.ShowFlash "���ڻ�ȡ��ת�������Ԥ������Ժ�...", Me
    If GetBillData(mlng����ID, strData) = False Then GoTo ErrExit:
    If InitData(mlng����ID, mlng��ҳID) = False Then GoTo ErrExit:
    If InitFace() = False Then GoTo ErrExit:
    zlCommFun.StopFlash
    Exit Sub
ErrExit:
    zlCommFun.StopFlash
    Unload Me: Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    mlng����ID = 0
    mlng��ҳID = 0
    mdblPrePay = 0
    mstrRemark = ""

    Set mpatiInfo = Nothing

End Sub

Private Sub cmdOK_Click()
    Dim strJsonIn As String
    Dim strData As String
    Dim blnTrans As Boolean
    
    On Error GoTo ErrHander
    cmdOk.Enabled = False
    zlCommFun.ShowFlash "���ڽ�������Ԥ�����תסԺ�������Ժ�...", Me
    
    If CheckPrePayValid = False Then cmdOk.Enabled = True: Exit Sub
    '����Ԥ������
    gcnOracle.BeginTrans: blnTrans = True
    If SaveDate() = False Then gcnOracle.RollbackTrans: cmdOk.Enabled = True: Exit Sub
    
    '����Ԥ�����תסԺȷ��
    '����    ����             ����      ˵��                ��������        ��ע
    '        pid              ����ID                         Number(18)       �ǿ�
    '        prepaid_payment  Ԥ����                        Number(18,2)    �ǿ�
    '
    '���    ����             ����      ˵��                ��������        ��ע
    '        result           ִ�н��  1-�ɹ���-1-ʧ��     Number(1)       �ǿ�
    '        errmsg           ������Ϣ  ʧ��ʱ���ش�����Ϣ  Varchar2(200)
    strJsonIn = "{""head"":{""bizno"":""RJ005"",""sysno"":""ZLDAYROOM"",""time"":"""",""action_no"":"""",""tarno"":""03""}"
    strJsonIn = "{""input"":" & strJsonIn & ",""pid"":" & mlng����ID & ",""prepaid_payment"":" & mdblPrePay & "}}"
    
    Call Sys.NewSystemSvr("������ϵͳ", "����Ԥ�����תסԺȷ��", strJsonIn, strData)

    If strData = "" Then strData = "{}"
    If Val(zlStr.JSONParse("result", strData)) <> 1 Then
        gcnOracle.RollbackTrans
        MsgBox zlStr.JSONParse("errmsg", strData), vbInformation, gstrSysName
        zlCommFun.StopFlash: mblnOK = True
        Exit Sub
    End If
    gcnOracle.CommitTrans: blnTrans = False
    zlCommFun.StopFlash
    cmdOk.Enabled = True
    mblnOK = True
    Unload Me
    Exit Sub
ErrHander:
    zlCommFun.StopFlash
    cmdOk.Enabled = True
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Function InitFace() As Boolean
    '��ʼ������
    On Error GoTo ErrHandler
    
    lbl(txtName).Caption = mpatiInfo.����
    lbl(txtSex).Caption = mpatiInfo.�Ա�
    lbl(txtAge).Caption = mpatiInfo.����
    lbl(txtInNumber).Caption = mpatiInfo.סԺ��
    
    txtPrePay.Text = Format(mdblPrePay, "0.00")
    txtPrePay.Tag = Nvl(mdblPrePay)
    If Nvl(mstrRemark) = "" Then
        lblremark.Visible = False
    Else
        lblremark.Caption = "˵��:" & mstrRemark
        If LenB(lblremark.Caption) > 50 Then lblremark.Caption = MidB(lblremark.Caption, 1, 50) & "����"
    End If
    Call SetPatiControl
    InitFace = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitData(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '��ʼ������
    On Error GoTo ErrHandler
    '��ȡ������Ϣ
     If GetPatiInfo(lng����ID, lng��ҳID, mpatiInfo) = False Then
        MsgBox "δ�ҵ�������Ϣ��", vbInformation, gstrSysName
        Exit Function
    End If
        
    InitData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetPatiControl()
    '���ò�����Ϣ�ؼ�λ��
    Dim sngSplit As Single
    
    sngSplit = 200
    On Error Resume Next
    lbl(txtName).Left = lbl(lblName).Left + lbl(lblName).Width
    
    lbl(lblSex).Left = lbl(txtName).Left + lbl(txtName).Width + sngSplit
    lbl(txtSex).Left = lbl(lblSex).Left + lbl(lblSex).Width
    
    lbl(lblAge).Left = lbl(txtSex).Left + lbl(txtSex).Width + sngSplit
    lbl(txtAge).Left = lbl(lblAge).Left + lbl(lblAge).Width
    
    lbl(lblInNumber).Left = lbl(txtAge).Left + lbl(txtAge).Width + sngSplit
    lbl(txtInNumber).Left = lbl(lblInNumber).Left + lbl(lblInNumber).Width
End Sub

Private Function GetBillData(ByVal lng����ID As Long, ByRef strData As String) As Boolean
    'ͨ�������ȡ����
    Dim strJsonIn As String
    
    On Error GoTo ErrHandler
    
    '�������������Ԥ�����תסԺ������
    '    ����    ����               ����      ˵��                  ��������        ��ע
    '            pid                ����ID                          Number(18)      �ǿ�
    '
    '    ���    ����               ����       ˵��                 ��������        ��ע
    '            result             ִ�н��   1-�ɹ���-1-ʧ��      Number(1)       �ǿ�
    '            errmsg             ������Ϣ   ʧ��ʱ���ش�����Ϣ   Varchar2(200)
    '            prepaid_payment    Ԥ����                          Number(18,2)    �ǿ�
    '            remark             ��ע       ��ע��Ϣ����:���������ֽ�ת�롱������ZLHIS������Ԥ����¼.ժҪ����
    '                                                               VARCHAR2(50)
                         
    strJsonIn = "{""head"":{""bizno"":""RJ004"",""sysno"":""ZLDAYROOM"",""time"":"""",""action_no"":"""",""tarno"":""03""}"
    strJsonIn = "{""input"":" & strJsonIn & ",""pid"":" & lng����ID & "}}"
    Call Sys.NewSystemSvr("������ϵͳ", "����Ԥ�����תסԺ", strJsonIn, strData)
    If strData = "" Then strData = "{}"
    If Val(zlStr.JSONParse("result", strData)) <> 1 Then
        MsgBox "��ȡ������Ԥ�����תסԺ��Ϣʱ����" & vbCrLf & _
            zlStr.JSONParse("errmsg", strData), vbInformation, gstrSysName
        Exit Function
    End If
    mstrRemark = Nvl(zlStr.JSONParse("remark", strData))
    mdblPrePay = Val(zlStr.JSONParse("prepaid_payment", strData))
    If Nvl(mdblPrePay) = 0 Then
        MsgBox "�ò����޿�������Ԥ��תסԺԤ����", vbOKOnly + vbInformation, gstrSysName
        Exit Function
    End If
    
    GetBillData = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveDate() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�Ե�ǰ�����Ԥ����ݴ���
    '   lng����ID - ����ID
    '   lng��ҳID - ��ҳID
    '����:�ɹ�����True,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNO As String, strSQL As String
    Dim lngԤ��ID As Long
    
    strNO = zlDatabase.GetNextNo(11)
    lngԤ��ID = zlDatabase.GetNextId("����Ԥ����¼")
    
    'Zl_����Ԥ����¼_Insert_S
    strSQL = "Zl_����Ԥ����¼_Insert_S("
    '  Id_In         ����Ԥ����¼.ID%Type,
    strSQL = strSQL & "" & lngԤ��ID & ","
    '  ���ݺ�_In     ����Ԥ����¼.NO%Type,
    strSQL = strSQL & "'" & strNO & "',"
    '  Ʊ�ݺ�_In     Ʊ��ʹ����ϸ.����%Type,
    strSQL = strSQL & "NULL,"
    '  ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "" & ZVal(mlng����ID) & ","
    '  ��ҳid_In     ����Ԥ����¼.��ҳid%Type,
    strSQL = strSQL & "" & ZVal(mlng��ҳID) & ","
    '  ����_In         ����Ԥ����¼.����%Type,
    strSQL = strSQL & "'" & mpatiInfo.���� & "',"
    '  �Ա�_In         ����Ԥ����¼.�Ա�%Type,
    strSQL = strSQL & "'" & mpatiInfo.�Ա� & "',"
    '  ����_In         ����Ԥ����¼.����%Type,
    strSQL = strSQL & "'" & mpatiInfo.���� & "',"
    '  �����_In       ����Ԥ����¼.�����%Type,
    strSQL = strSQL & "NULL,"
    '  סԺ��_In       ����Ԥ����¼.סԺ��%Type,
    strSQL = strSQL & ZVal(mpatiInfo.סԺ��) & ","
    '  ���ʽ����_In ����Ԥ����¼.���ʽ����%Type,
    strSQL = strSQL & "'" & mpatiInfo.ҽ�Ƹ��ʽ & "',"
    '  ����id_In     ����Ԥ����¼.����id%Type,
    strSQL = strSQL & "NULL,"
    '  ���_In       ����Ԥ����¼.���%Type,
    strSQL = strSQL & "" & mdblPrePay & ","
    '  ���㷽ʽ_In   ����Ԥ����¼.���㷽ʽ%Type,
    strSQL = strSQL & "'" & "�ֽ�" & "',"
    '  �������_In   ����Ԥ����¼.�������%Type,
    strSQL = strSQL & "NULL,"
    '  �ɿλ_In   ����Ԥ����¼.�ɿλ%Type,
    strSQL = strSQL & "NULL,"
    '  ��λ������_In ����Ԥ����¼.��λ������%Type,
    strSQL = strSQL & "NULL,"
    '  ��λ�ʺ�_In   ����Ԥ����¼.��λ�ʺ�%Type,
    strSQL = strSQL & "NULL,"
    '  ժҪ_In       ����Ԥ����¼.ժҪ%Type,
    strSQL = strSQL & "'" & mstrRemark & "',"
    '  ����Ա���_In ����Ԥ����¼.����Ա���%Type,
    strSQL = strSQL & "'" & UserInfo.��� & "',"
    '  ����Ա����_In ����Ԥ����¼.����Ա����%Type,
    strSQL = strSQL & "'" & UserInfo.���� & "',"
    '  ����id_In     Ʊ��ʹ����ϸ.����id%Type,
    strSQL = strSQL & "NULL,"
    '  Ԥ�����_In   ����Ԥ����¼.Ԥ�����%Type := Null,
    strSQL = strSQL & " 2)"
    On Error GoTo errH
    
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    SaveDate = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txtPrePay_Change()
    If Val(txtPrePay.Text) <> 0 Then mdblPrePay = Val(txtPrePay.Text)
End Sub

Private Sub txtPrePay_GotFocus()
    zlControl.TxtSelAll txtPrePay
End Sub

Private Sub txtPrePay_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtPrePay_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txtPrePay, KeyAscii, m���ʽ
End Sub

Private Function CheckPrePayValid() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������Ԥ�������Ч��
    '����:�ɹ�����True,���򷵻�False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    
    If Val(txtPrePay.Text) > Val(txtPrePay.Tag) Then
        MsgBox "�ò�������Ԥ�����תסԺ����ܳ���" & txtPrePay.Tag & "!", vbOKOnly + vbInformation, gstrSysName
        txtPrePay.Text = Format(Val(txtPrePay.Tag), "0.00")
        txtPrePay.SetFocus: zlControl.TxtSelAll txtPrePay: Exit Function
    
    ElseIf Val(txtPrePay.Text) <= 0 Then
        MsgBox "����Ԥ�����תסԺ�����Ч,������0��" & txtPrePay.Tag & "֮��!", vbOKOnly + vbInformation, gstrSysName
        txtPrePay.Text = Format(Val(txtPrePay.Tag), "0.00")
        txtPrePay.SetFocus: zlControl.TxtSelAll txtPrePay: Exit Function
    End If
    
    CheckPrePayValid = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetPatiInfo(ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
                          ByRef PatiPageInfo As clsPatientInfo) As Boolean
    '���ܣ����ݲ���id����ҳid��ȡ������Ϣ�Ͳ�����ҳ�е���Ϣ
    '��Σ�lng����id-����id��lng��ҳid-��ҳid
    '���Σ�PatiPageInfo-������ҳ�е���Ϣ
    '���أ���ȡ�ɹ�����true,���򷵻�false
    Dim str����id As String
    
    On Error GoTo errHandle
  
    '��ȡָ��סԺ����סԺ����Ϣ
    str����id = lng����ID & ":" & lng��ҳID
    Call GetPatiPageInforByID(str����id, PatiPageInfo, False)
    If PatiPageInfo.����ID = 0 Then Exit Function
      
    GetPatiInfo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
