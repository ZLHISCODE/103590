VERSION 5.00
Begin VB.Form frmGroupEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������Ϣ"
   ClientHeight    =   6150
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10170
   Icon            =   "frmGroupEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   10170
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   165
      TabIndex        =   21
      Top             =   5640
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   8895
      TabIndex        =   20
      Top             =   5625
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7695
      TabIndex        =   19
      Top             =   5625
      Width           =   1100
   End
   Begin VB.PictureBox picBack 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   -15
      ScaleHeight     =   5415
      ScaleWidth      =   10125
      TabIndex        =   22
      Top             =   0
      Width           =   10125
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   9
         Left            =   5580
         TabIndex        =   15
         Top             =   1830
         Width           =   4410
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   8
         Left            =   1230
         TabIndex        =   13
         Top             =   1860
         Width           =   2940
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   10
         Left            =   1230
         TabIndex        =   17
         Top             =   2250
         Width           =   8775
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   2
         Left            =   3585
         TabIndex        =   5
         Top             =   480
         Width           =   6435
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   7
         Left            =   1230
         TabIndex        =   9
         Top             =   1485
         Width           =   2940
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1230
         TabIndex        =   1
         Top             =   120
         Width           =   8790
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   3
         Left            =   1230
         TabIndex        =   7
         Top             =   1125
         Width           =   2940
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1230
         MaxLength       =   18
         TabIndex        =   3
         Top             =   480
         Width           =   1590
      End
      Begin VB.TextBox txt 
         Height          =   2745
         Index           =   4
         Left            =   225
         TabIndex        =   18
         Top             =   2610
         Width           =   9780
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   5
         Left            =   5580
         TabIndex        =   11
         Top             =   1470
         Width           =   4410
      End
      Begin VB.Label lbl���֤�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&U)"
         Height          =   180
         Left            =   195
         TabIndex        =   2
         Top             =   540
         Width           =   630
      End
      Begin VB.Label lbl��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ʺ�(&Z)"
         Height          =   180
         Left            =   4455
         TabIndex        =   14
         Top             =   1875
         Width           =   990
      End
      Begin VB.Label lblְҵ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&B)"
         Height          =   180
         Left            =   195
         TabIndex        =   12
         Top             =   1935
         Width           =   990
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&N)"
         Height          =   180
         Left            =   195
         TabIndex        =   0
         Top             =   195
         Width           =   630
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����(&I)"
         Height          =   180
         Left            =   2895
         TabIndex        =   4
         Top             =   555
         Width           =   630
      End
      Begin VB.Label lblѧ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ�绰(&T)"
         Height          =   180
         Left            =   195
         TabIndex        =   8
         Top             =   1575
         Width           =   990
      End
      Begin VB.Label lvl����״�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ��ַ(&A)"
         Height          =   180
         Left            =   195
         TabIndex        =   16
         Top             =   2295
         Width           =   990
      End
      Begin VB.Label lbl��ͥ��ַ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ��(&L)"
         Height          =   180
         Left            =   195
         TabIndex        =   6
         Top             =   1200
         Width           =   810
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�����ʼ�(&E)"
         Height          =   180
         Index           =   7
         Left            =   4455
         TabIndex        =   10
         Top             =   1560
         Width           =   990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   150
         X2              =   10320
         Y1              =   930
         Y2              =   930
      End
   End
End
Attribute VB_Name = "frmGroupEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'���������弶��������**************************************************************************************************
Private mblnStartUp As Boolean                          '����������־
Private mblnOK As Boolean
Private mfrmMain As Object
Private mvarParam As Variant
Private mblnDataChange As Boolean
Private mrsGroup As New ADODB.Recordset

'�������Զ�����̻���************************************************************************************************
Private Property Let EditChanged(ByVal vData As Boolean)
    '------------------------------------------------------------------------------------------------------------------
    '����:
    'ֵ��:
    '------------------------------------------------------------------------------------------------------------------
    mblnDataChange = vData

End Property

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long

    On Error Resume Next

    For lngLoop = 0 To txt.UBound
        txt(lngLoop).Text = ""
        txt(lngLoop).Tag = ""
    Next

    On Error GoTo 0

    Call InitData

    EditChanged = True


End Function

Public Function ShowEdit(ByVal frmMain As Object, ByRef rsGroup As ADODB.Recordset, Optional ByVal blnModify As Boolean = True) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʾ�༭���壬������ô���Ľӿں���
    '����:  frmMain         ���ô������
    '       lngKey          ԤԼ�Ǽ�id
    '����:  True
    '       False
    '------------------------------------------------------------------------------------------------------------------
    mblnStartUp = True
    mblnOK = False
        
    'mvarParam = Split(strParam, "'")
    Set mrsGroup = rsGroup

    Set mfrmMain = frmMain

    If InitData = False Then Exit Function
    If ReadData = False Then Exit Function
    
    If Trim(txt(1).Text) = "" Then txt(1).Text = GetNextCode("��Լ��λ", "����", "�ϼ�id IS NULL")
        
'    If blnModify Then
'        txt(0).Text = mvarParam(1)
'        txt(3).Text = mvarParam(2)
'        txt(7).Text = mvarParam(3)
'    End If
        
    EditChanged = False

    Me.Show 1, frmMain
        
    ShowEdit = mblnOK
    If mblnOK Then Set rsGroup = mrsGroup
    
End Function

Private Function ReadData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ȡ����
    '����:  lngKey      ����������
    '����:  True        ��ȡ�ɹ�
    '       False       ��ȡʧ��
    '------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHand

    If mrsGroup.BOF = False Then
        txt(0).Text = zlCommFun.NVL(mrsGroup("����").Value)
        txt(1).Text = zlCommFun.NVL(mrsGroup("����").Value)
        txt(2).Text = zlCommFun.NVL(mrsGroup("����").Value)
        txt(3).Text = zlCommFun.NVL(mrsGroup("��ϵ��").Value)
        txt(7).Text = zlCommFun.NVL(mrsGroup("�绰").Value)
        txt(5).Text = zlCommFun.NVL(mrsGroup("�����ʼ�").Value)
        txt(8).Text = zlCommFun.NVL(mrsGroup("��������").Value)
        txt(9).Text = zlCommFun.NVL(mrsGroup("�ʺ�").Value)
        txt(10).Text = zlCommFun.NVL(mrsGroup("��ַ").Value)
        txt(4).Text = zlCommFun.NVL(mrsGroup("˵��").Value)
    End If

    ReadData = True

    Exit Function

errHand:
    If ErrCenter = 1 Then Resume

End Function

Private Function InitData() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  ��ʼ������
    '����:  True        ��ʼ���ɹ�
    '       False       ��ʼ��ʧ��
    '------------------------------------------------------------------------------------------------------------------

    On Error GoTo errHand

    '����������볤��

    txt(0).MaxLength = GetMaxLength("��Լ��λ", "����")
    txt(1).MaxLength = GetMaxLength("��Լ��λ", "����")
    txt(2).MaxLength = GetMaxLength("��Լ��λ", "����")
    txt(3).MaxLength = GetMaxLength("��Լ��λ", "��ϵ��")
    txt(7).MaxLength = GetMaxLength("��Լ��λ", "�绰")
    txt(5).MaxLength = GetMaxLength("��Լ��λ", "�����ʼ�")
    txt(8).MaxLength = GetMaxLength("��Լ��λ", "��������")
    txt(9).MaxLength = GetMaxLength("��Լ��λ", "�ʺ�")
    txt(10).MaxLength = GetMaxLength("��Լ��λ", "��ַ")
    txt(4).MaxLength = GetMaxLength("��Լ��λ", "˵��")
    
'    If mblnModify = False Then
'        txt(0).Locked = True
'        txt(1).Locked = True
'        txt(2).Locked = True
'        txt(3).Locked = True
'        txt(4).Locked = True
'        txt(5).Locked = True
'        txt(7).Locked = True
'        txt(8).Locked = True
'        txt(9).Locked = True
'        txt(10).Locked = True
'
'        mnuFileSave.Visible = False
'        mnuFileRestore.Visible = False
'
'        mnuFile_1.Visible = False
'
'        tbrThis.Buttons("����").Visible = False
'        tbrThis.Buttons("����").Visible = False
'        tbrThis.Buttons("Split_1").Visible = False
'        tbrThis.Buttons("Split_2").Visible = False
'    End If
    
    InitData = True

    Exit Function

errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ValidEdit() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  У�����ݵ���Ч��
    '����:  True        ������Ч
    '       False       ������Ч
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long


    ValidEdit = True

End Function

'Private Function SaveEdit(ByRef lngKey As Long) As Boolean
'    '------------------------------------------------------------------------------------------------------------------
'    '����:  ��������
'    '����:  True        ����ɹ�
'    '       False       ����ʧ��
'    '------------------------------------------------------------------------------------------------------------------
'    Dim blnTran As Boolean
'    Dim lngLoop As Long
'    Dim strSQL() As String
'    Dim rsPati As New ADODB.Recordset
'
'    On Error GoTo errHand
'
'    ReDim Preserve strSQL(1 To 1)
'
'    gstrSQL = "SELECT * FROM ��Լ��λ WHERE ID=" & lngKey
'    Call OpenRecord(rsPati, gstrSQL, Me.Caption)
'    If rsPati.BOF = False Then
'        '����
''        ID_IN IN ��Լ��λ.ID%TYPE,
''        �ϼ�ID_IN IN ��Լ��λ.�ϼ�ID%TYPE,
''        ����_IN IN ��Լ��λ.����%TYPE,
''        ����_IN IN ��Լ��λ.����%TYPE,
''        ����_IN IN ��Լ��λ.����%TYPE,
''        ��ַ_IN IN ��Լ��λ.��ַ%TYPE := NULL,
''        �绰_IN IN ��Լ��λ.�绰%TYPE := NULL,
''        ��������_IN IN ��Լ��λ.��������%TYPE := NULL,
''        �ʺ�_IN IN ��Լ��λ.�ʺ�%TYPE := NULL,
''        ��ϵ��_IN IN ��Լ��λ.��ϵ��%TYPE := NULL,
''        ԭ����_IN IN PLS_INTEGER,
''        �����ʼ�_IN IN ��Լ��λ.�����ʼ�%TYPE := NULL,
''        ˵��_IN IN ��Լ��λ.˵��%TYPE := NULL
'
'        gstrSQL = "zl_��Լ��λ_Update(" & lngKey & "," & _
'                                        IIf(IsNull(rsPati("�ϼ�ID").Value), "NULL", rsPati("�ϼ�ID").Value) & ",'" & _
'                                        txt(1).Text & "','" & _
'                                        txt(0).Text & "','" & _
'                                        txt(2).Text & "','" & _
'                                        txt(10).Text & "','" & _
'                                        txt(7).Text & "','" & _
'                                        txt(8).Text & "','" & _
'                                        txt(9).Text & "','" & _
'                                        txt(3).Text & "',0,'" & txt(5).Text & "','" & txt(4).Text & "')"
'        strSQL(ReDimArray(strSQL)) = gstrSQL
'    Else
'        '������
'    '    ID_IN IN ��Լ��λ.ID%TYPE,
'    '    �ϼ�ID_IN IN ��Լ��λ.�ϼ�ID%TYPE,
'    '    ����_IN IN ��Լ��λ.����%TYPE,
'    '    ����_IN IN ��Լ��λ.����%TYPE,
'    '    ����_IN IN ��Լ��λ.����%TYPE := NULL,
'    '    ��ַ_IN IN ��Լ��λ.��ַ%TYPE := NULL,
'    '    �绰_IN IN ��Լ��λ.�绰%TYPE := NULL,
'    '    ��������_IN IN ��Լ��λ.��������%TYPE := NULL,
'    '    �ʺ�_IN IN ��Լ��λ.�ʺ�%TYPE := NULL,
'    '    ��ϵ��_IN IN ��Լ��λ.��ϵ��%TYPE := NULL,
'    '    ĩ��_IN IN ��Լ��λ.ĩ��%TYPE := 1,
''        �����ʼ�_IN IN ��Լ��λ.�����ʼ�%TYPE := NULL,
''        ˵��_IN IN ��Լ��λ.˵��%TYPE := NULL
'        lngKey = zlDatabase.GetNextId("��Լ��λ")
'        gstrSQL = "zl_��Լ��λ_Insert(" & lngKey & "," & _
'                                        "NULL,'" & _
'                                        txt(1).Text & "','" & _
'                                        txt(0).Text & "','" & _
'                                        txt(2).Text & "','" & _
'                                        txt(10).Text & "','" & _
'                                        txt(7).Text & "','" & _
'                                        txt(8).Text & "','" & _
'                                        txt(9).Text & "','" & _
'                                        txt(3).Text & "',1,'" & txt(5).Text & "','" & txt(4).Text & "')"
'        strSQL(ReDimArray(strSQL)) = gstrSQL
'    End If
'
'    blnTran = True
'    gcnOracle.BeginTrans
'    For lngLoop = 1 To UBound(strSQL)
'        If strSQL(lngLoop) <> "" Then Call ExecuteProc(strSQL(lngLoop), Me.Caption)
'    Next
'    gcnOracle.CommitTrans
'    blnTran = False
'
'    SaveEdit = True
'
'    Exit Function
'
'errHand:
'
'    If ErrCenter = 1 Then Resume
'    If blnTran Then gcnOracle.RollbackTrans
'
'End Function


'���������弰��ؼ����¼�����******************************************************************************************

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    
    If ValidEdit = False Then Exit Sub

    mblnOK = True
    
    mrsGroup("����").Value = txt(0).Text
    mrsGroup("����").Value = txt(1).Text
    mrsGroup("����").Value = txt(2).Text
    mrsGroup("��ϵ��").Value = txt(3).Text
    mrsGroup("�绰").Value = txt(7).Text
    mrsGroup("�����ʼ�").Value = txt(5).Text
    mrsGroup("��������").Value = txt(8).Text
    mrsGroup("�ʺ�").Value = txt(9).Text
    mrsGroup("��ַ").Value = txt(10).Text
    mrsGroup("˵��").Value = txt(4).Text
    
    EditChanged = False
    Unload Me

    
End Sub


Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If mblnDataChange Then
        Cancel = (MsgBox("���ݱ��뱣������Ч���Ƿ񲻱�����˳���", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
        If Cancel Then Exit Sub
    End If
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub cmdHelp_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub txt_Change(Index As Integer)
    EditChanged = True
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
    Select Case Index
    Case 0, 3, 4, 10
        zlCommFun.OpenIme True
    End Select
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        zlCommFun.PressKey vbKeyTab
        
    Else
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        If Index = 2 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    Select Case Index
    Case 0, 3, 4, 10
        zlCommFun.OpenIme False
    End Select
    
    If Index = 0 Then
        If InStr(txt(Index).Text, "'") = 0 Then txt(2).Text = zlGetSymbol(txt(Index).Text)
    End If
End Sub

Private Sub txt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txt(Index).Locked Then
        glngTXTProc = GetWindowLong(txt(Index).hWnd, GWL_WNDPROC)
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txt_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And txt(Index).Locked Then
        Call SetWindowLong(txt(Index).hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub


