VERSION 5.00
Begin VB.Form frmҽ�����༭ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ�����༭"
   ClientHeight    =   3735
   ClientLeft      =   2760
   ClientTop       =   3645
   ClientWidth     =   5640
   Icon            =   "frmҽ�����༭.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chkNoLink 
      Caption         =   "������ֹ��������(&S)"
      Height          =   225
      Left            =   1215
      TabIndex        =   15
      Top             =   3270
      Width           =   2025
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "��"
      Enabled         =   0   'False
      Height          =   315
      Left            =   3660
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2160
      Width           =   315
   End
   Begin VB.CheckBox chk��ֹ 
      Caption         =   "��ϵͳ�н�ֹʹ��(&S)"
      Height          =   225
      Left            =   1215
      TabIndex        =   10
      Top             =   2910
      Width           =   2025
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "���ж��ҽ������(&R)"
      Height          =   225
      Left            =   1215
      TabIndex        =   9
      Top             =   2571
      Width           =   2025
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   4320
      TabIndex        =   13
      Top             =   2745
      Width           =   1100
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   3
      Left            =   1215
      MaxLength       =   20
      TabIndex        =   7
      Top             =   2157
      Width           =   2430
   End
   Begin VB.TextBox txtEdit 
      Height          =   1080
      Index           =   2
      Left            =   1215
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   963
      Width           =   2790
   End
   Begin VB.Frame Frame1 
      Height          =   4020
      Left            =   4155
      TabIndex        =   14
      Top             =   -195
      Width           =   30
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   0
      Left            =   1215
      MaxLength       =   3
      TabIndex        =   1
      Top             =   135
      Width           =   555
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   1
      Left            =   1215
      MaxLength       =   20
      TabIndex        =   3
      Top             =   549
      Width           =   2775
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4320
      TabIndex        =   12
      Top             =   660
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4320
      TabIndex        =   11
      Top             =   150
      Width           =   1100
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "ҽԺ���(&B)"
      Height          =   180
      Index           =   3
      Left            =   210
      TabIndex        =   6
      Top             =   2217
      Width           =   990
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "Ӧ��˵��(&E)"
      Height          =   180
      Index           =   2
      Left            =   195
      TabIndex        =   4
      Top             =   1020
      Width           =   990
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "ҽ�����(&S)"
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   195
      Width           =   990
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "ҽ������(&N)"
      Height          =   180
      Index           =   1
      Left            =   195
      TabIndex        =   2
      Top             =   609
      Width           =   990
   End
End
Attribute VB_Name = "frmҽ�����༭"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enum�༭
    Text��� = 0
    Text���� = 1
    Text˵�� = 2
    TextҽԺ���� = 3
End Enum

Dim mstr��� As String         '��ǰ�༭��ҽ��������
Dim mblnOK As Boolean
Dim mblnChange As Boolean     '�Ƿ�ı���

Private Sub chkNoLink_Click()
On Error GoTo ErrH
    SaveSetting "ZLSOFT", "����ģ��\ҽ��\" & mstr���, "NoLink", chkNoLink.Value
    Exit Sub
ErrH:
    Err.Clear
End Sub

Private Sub cmdGet_Click()
    Dim strReturn As String
    
    If mstr��� = "10" Then
        strReturn = ҽԺ����_����
        If strReturn <> "" Then
            TxtEdit(TextҽԺ����) = strReturn
            mblnChange = True
        End If
    ElseIf mstr��� = TYPE_��Ϫũҽ Then
        If Not ҽ����ʼ��_��Ϫũҽ Then Exit Sub
        Call ���ýӿ�_׼��_��Ϫũҽ(gstrFunc��Ϫũҽ_GetHospitalInfo, "")
        If Not ���ýӿ�_��Ϫũҽ() Then Exit Sub
        '������������ҽԺ����
        strReturn = gstrOutput_��Ϫũҽ
        TxtEdit(TextҽԺ����).Text = Split(Split(strReturn, "&")(2), "=")(1)
    End If
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    If IsValid() = False Then Exit Sub
    
    MousePointer = vbHourglass
    If Saveҽ�����() = False Then
        MousePointer = vbDefault
        Exit Sub
    End If
    MousePointer = vbDefault
    
    mblnOK = True
    mblnChange = False
    
    Unload Me
End Sub

Private Function IsValid() As Boolean
'����:���������й�ҽ�����������Ƿ���Ч
'����:
'����ֵ:��Ч����True,����ΪFalse
    Dim lngIndex As Integer
    Dim strTemp As String
    For lngIndex = Text��� To TextҽԺ����
        If zlCommFun.StrIsValid(Trim(TxtEdit(lngIndex).Text), TxtEdit(lngIndex).MaxLength) = False Then
            TxtEdit(lngIndex).SetFocus
            zlControl.TxtSelAll TxtEdit(lngIndex)
            Exit Function
        End If
        
        If lngIndex = Text��� Or lngIndex = Text���� Then
            If Len(Trim(TxtEdit(lngIndex).Text)) = 0 Then
                TxtEdit(lngIndex).Text = ""
                MsgBox "��Ż����ƶ�����Ϊ�ա�", vbExclamation, gstrSysName
                TxtEdit(lngIndex).SetFocus
                Exit Function
            End If
        End If
    Next
    
    If TxtEdit(Text���).Enabled = True Then
        If IsNumeric(TxtEdit(Text���)) = False Or Val(TxtEdit(Text���).Text) <= 900 Then
            MsgBox "���ֻ���Ǵ���900��������", vbExclamation, gstrSysName
            zlControl.TxtSelAll TxtEdit(Text���)
            TxtEdit(Text���).SetFocus
            Exit Function
        End If
    End If
    
    IsValid = True
End Function

Private Function Saveҽ�����() As Boolean
'����:����༭�����ݵ�ҽ��������
'����:
'����ֵ:�ɹ�����True,����ΪFalse
    Dim lng��� As Long
    Dim lst As ListItem
    
    gcnOracle.BeginTrans
    On Error GoTo errHandle
    
    If mstr��� = "" Then       '����һ����¼
        lng��� = Val(TxtEdit(Text���).Text)
        gstrSQL = "zl_�������_Insert(" & lng��� & _
            ",'" & TxtEdit(Text����).Text & "','" & TxtEdit(Text˵��).Text & _
            "','" & TxtEdit(TextҽԺ����).Text & "'," & chk����.Value & "," & chk��ֹ.Value & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        If chk����.Value = 0 Then
            '������ҽ��Ԥ�Ⱦ���������
            gstrSQL = "zl_��������Ŀ¼_Insert(" & lng��� & ",0,'1','" & TxtEdit(Text����).Text & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
    Else    '�޸�
        gstrSQL = "zl_�������_Update(" & mstr��� & _
            ",'" & TxtEdit(Text����).Text & "','" & TxtEdit(Text˵��).Text & _
            "','" & TxtEdit(TextҽԺ����).Text & "'," & chk��ֹ.Value & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    
    gcnOracle.CommitTrans
    
    '��������������Ӧ�ĵ���
    If mstr��� = "" Then
        '����
        Set lst = frmҽ�����.lvwKind_S.ListItems.Add(, "K" & lng���, " ", "Common", "Common")
        lst.Selected = True
        lst.EnsureVisible
    Else
        '�޸�
        Set lst = frmҽ�����.lvwKind_S.SelectedItem
    End If
    lst.Text = TxtEdit(Text����).Text
    lst.SubItems(1) = TxtEdit(Text���).Text
    lst.SubItems(2) = TxtEdit(TextҽԺ����).Text
    lst.SubItems(3) = TxtEdit(Text˵��).Text
    lst.Tag = chk����.Value
    lst.Ghosted = (chk��ֹ.Value = 1)
    
    Saveҽ����� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
    Call SaveErrLog
End Function

Public Function �༭ҽ�����(ByVal str��� As String) As Boolean
'����:��������õ�ҽ���������ڽ���ͨѶ�ĳ���
'����:str���           ��ǰ�༭��ҽ�����ĵ����
'����ֵ:�༭�ɹ�����True,����ΪFalse
    Dim rsҽ����� As New ADODB.Recordset
    Dim i As Integer
    
    mstr��� = str���
    If str��� = "10" Then
        If ҽ����ʼ��_���� = True Then
            cmdGet.Enabled = True
        End If
    End If
    
    mblnOK = False
    
    rsҽ�����.CursorLocation = adUseClient
    
    If str��� <> "" Then
        gstrSQL = "Select ���,����,˵��,ҽԺ����,��������,�Ƿ��ֹ" & _
            " From �������  Where ���=[1]"
        Set rsҽ����� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(str���))
        
        TxtEdit(Text���).Text = rsҽ�����("���")
        TxtEdit(Text����).Text = rsҽ�����("����")
        TxtEdit(Text˵��).Text = IIf(IsNull(rsҽ�����("˵��")), "", rsҽ�����("˵��"))
        TxtEdit(TextҽԺ����).Text = IIf(IsNull(rsҽ�����("ҽԺ����")), "", rsҽ�����("ҽԺ����"))
        chk����.Value = IIf(rsҽ�����("��������") = 1, 1, 0)
        chk��ֹ.Value = IIf(rsҽ�����("�Ƿ��ֹ") = 1, 1, 0)
        
        lblEdit(Text���).Enabled = False
        TxtEdit(Text���).Enabled = False
        chk����.Enabled = False
    Else
        TxtEdit(Text���).Text = zlDatabase.GetMax("�������", "���", 3)
        If Val(TxtEdit(Text���).Text) < 901 Then TxtEdit(Text���).Text = 901
    End If
    
    mblnChange = False
    
    chkNoLink.Enabled = (mstr��� = TYPE_�Ͻ�)
    frmҽ�����༭.Show vbModal
    �༭ҽ����� = mblnOK
End Function

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll TxtEdit(Index)
    Select Case Index
        Case Text����, Text˵��
          zlCommFun.OpenIme True
        Case Text���, TextҽԺ����
            zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0 'ʹ֮����
        zlCommFun.PressKey (vbKeyTab)
    Else
        If Index = Text��� Then
            KeyAscii = asc(UCase(Chr(KeyAscii)))
            If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    zlCommFun.OpenIme False
End Sub

Private Sub chk����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub chk��ֹ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub


