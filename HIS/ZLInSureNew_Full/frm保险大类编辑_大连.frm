VERSION 5.00
Begin VB.Form frm���մ���༭_���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "���մ���༭"
   ClientHeight    =   4980
   ClientLeft      =   2760
   ClientTop       =   3645
   ClientWidth     =   4530
   Icon            =   "frm���մ���༭_����.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.ComboBox cmb������� 
      Height          =   300
      ItemData        =   "frm���մ���༭_����.frx":000C
      Left            =   1170
      List            =   "frm���մ���༭_����.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1335
      Width           =   1425
   End
   Begin VB.CheckBox chkҽ�� 
      Caption         =   "ҽ����Ŀ(&I)"
      Height          =   225
      Left            =   1170
      TabIndex        =   8
      Top             =   1770
      Value           =   1  'Checked
      Width           =   1305
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   315
      TabIndex        =   18
      Top             =   4470
      Width           =   1100
   End
   Begin VB.Frame frmRule 
      Caption         =   "ͳ��֧���������"
      Height          =   1500
      Left            =   285
      TabIndex        =   13
      Top             =   2820
      Width           =   4080
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   5
         Left            =   1860
         MaxLength       =   16
         TabIndex        =   22
         Top             =   990
         Width           =   1320
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   4
         Left            =   1845
         MaxLength       =   16
         TabIndex        =   19
         Top             =   630
         Width           =   1320
      End
      Begin VB.TextBox txtEdit 
         Alignment       =   1  'Right Justify
         Height          =   300
         Index           =   3
         Left            =   1845
         MaxLength       =   16
         TabIndex        =   15
         Top             =   285
         Width           =   1320
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "�����                Ԫ"
         Height          =   330
         Left            =   810
         TabIndex        =   23
         Top             =   990
         Width           =   2775
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "�����               Ԫ"
         Height          =   180
         Left            =   1125
         TabIndex        =   21
         Top             =   1050
         Width           =   2250
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "סԺͳ��֧������               %"
         Height          =   180
         Index           =   4
         Left            =   390
         TabIndex        =   20
         Top             =   690
         Width           =   2880
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����ͳ��֧������               %"
         Height          =   180
         Index           =   3
         Left            =   390
         TabIndex        =   14
         Top             =   345
         Width           =   2880
      End
   End
   Begin VB.Frame fraKind 
      Caption         =   "����"
      Height          =   630
      Left            =   285
      TabIndex        =   9
      Top             =   2070
      Width           =   4095
      Begin VB.OptionButton opt���� 
         Caption         =   "����(&F)"
         Height          =   180
         Index           =   3
         Left            =   2640
         TabIndex        =   12
         Top             =   315
         Width           =   945
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "ҽ��(&D)"
         Height          =   180
         Index           =   2
         Left            =   1425
         TabIndex        =   11
         Top             =   315
         Width           =   945
      End
      Begin VB.OptionButton opt���� 
         Caption         =   "ҩƷ(&M)"
         Height          =   180
         Index           =   1
         Left            =   225
         TabIndex        =   10
         Top             =   315
         Value           =   -1  'True
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2130
      TabIndex        =   16
      Top             =   4470
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3330
      TabIndex        =   17
      Top             =   4470
      Width           =   1100
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   2
      Left            =   1170
      MaxLength       =   10
      TabIndex        =   5
      Top             =   937
      Width           =   1425
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   1
      Left            =   1170
      MaxLength       =   40
      TabIndex        =   3
      Top             =   536
      Width           =   3195
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   0
      Left            =   1170
      MaxLength       =   6
      TabIndex        =   1
      Top             =   135
      Width           =   1425
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      Caption         =   "�������(&F)"
      Height          =   180
      Left            =   135
      TabIndex        =   6
      Top             =   1398
      Width           =   990
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����(&U)"
      Height          =   180
      Index           =   0
      Left            =   495
      TabIndex        =   0
      Top             =   195
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����(&S)"
      Height          =   180
      Index           =   2
      Left            =   495
      TabIndex        =   4
      Top             =   997
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����(&N)"
      Height          =   180
      Index           =   1
      Left            =   495
      TabIndex        =   2
      Top             =   596
      Width           =   630
   End
End
Attribute VB_Name = "frm���մ���༭_����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enum�༭
    text���� = 0
    Text���� = 1
    Text���� = 2
    Text���� = 3
    TextסԺ = 4
    Text���� = 5
    
    CheckҩƷ = 1
    Checkҽ�� = 2
    Check���� = 3
    
    Check���� = 1
    CheckסԺ�� = 2
End Enum

Dim mlng���� As Long
Dim mstrID As String         '��ǰ�༭��ҽ������ID
Dim mblnOK As Boolean
Dim mblnChange As Boolean     '�Ƿ�ı���

Private Sub Check1_Click()

End Sub

Private Sub chk����_Click()
    txtEdit(Text����).Enabled = chk����.Value = 1
    txtEdit(Text����).Enabled = chk����.Value <> 1
    txtEdit(TextסԺ).Enabled = chk����.Value <> 1
End Sub

Private Sub chkҽ��_Click()
    mblnChange = True
End Sub

Private Sub chkҽ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub cmb�������_Click()
    mblnChange = True
    Select Case cmb�������.ListIndex
    Case 0  '����
        lblEdit(3).Enabled = True
        txtEdit(Text����).Enabled = True
        lblEdit(4).Enabled = False
        txtEdit(TextסԺ).Enabled = False
    Case 1  'סԺ
        lblEdit(3).Enabled = False
        txtEdit(Text����).Enabled = False
        lblEdit(4).Enabled = True
        txtEdit(TextסԺ).Enabled = True
    Case 2  '����
        lblEdit(3).Enabled = True
        txtEdit(Text����).Enabled = True
        lblEdit(4).Enabled = True
        txtEdit(TextסԺ).Enabled = True
    End Select
End Sub

Private Sub cmb�������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
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
    Dim lngIndex As Long
    
    If IsValid() = False Then Exit Sub
    If Save��Ŀ() = False Then Exit Sub
    
    If mstrID = "" Then
        '��������
        txtEdit(text����).Text = zlDatabase.GetMax("����֧������", "����", 6, " where ����=" & mlng����)
        For lngIndex = Text���� To TextסԺ
            txtEdit(lngIndex).Text = ""
        Next
        chkҽ��.Value = 1
        mblnChange = False
        txtEdit(text����).SetFocus
    Else
        mblnChange = False
        Unload Me
    End If
End Sub

Private Function Save��Ŀ() As Boolean
    Dim lngID As Long, lng���� As Long, lng�㷨 As Long
    Dim dblͳ��ȶ� As Double, dbl��׼���� As Double, dbl��׼���� As Double, dblסԺ�ȶ� As Double
    Dim dbl����ȶ� As Double
    Dim lngIndex As Long, lst As ListItem
    
    On Error GoTo errHandle
    
    For lngIndex = 1 To 3
        If opt����(lngIndex).Value = True Then
            lng���� = lngIndex
            Exit For
        End If
    Next
    dblͳ��ȶ� = 0
    Select Case Me.cmb�������.ListIndex
    Case 0
        dblͳ��ȶ� = Val(txtEdit(Text����).Text)
        dbl����ȶ� = Val(txtEdit(Text����).Text)
        dblסԺ�ȶ� = 0
    Case 1
        dbl����ȶ� = 0
        dblסԺ�ȶ� = Val(txtEdit(TextסԺ).Text)
    Case Else
        dblͳ��ȶ� = Val(txtEdit(Text����).Text)
        dbl����ȶ� = Val(txtEdit(Text����).Text)
        dblסԺ�ȶ� = Val(txtEdit(TextסԺ).Text)
    End Select
    
    dbl��׼���� = Val(txtEdit(Text����).Text)
    dbl��׼���� = 0
    If chk����.Value = 1 Then
        lng�㷨 = 2
    Else
        lng�㷨 = 1
    End If
    
    'zl_����֧������_UPDATE (
    '   ID_IN,����_IN,����_IN,����_IN,����_IN,�㷨_IN,����ȶ�_IN,סԺ�ȶ�_IN,��׼����_IN,��׼����_IN,
    '   �������_IN,�Ƿ�ҽ��_IN
    
    If mstrID = "" Then
        '����
        lngID = zlDatabase.GetNextID("����֧������")
        gstrSQL = "zl_����֧������_INSERT(" & lngID & "," & mlng���� & ",'" & Trim(txtEdit(text����).Text) & "','" & _
                Trim(txtEdit(Text����).Text) & "','" & Trim(txtEdit(Text����).Text) & "'," & lng���� & "," & lng�㷨 & "," & _
                 dbl����ȶ� & "," & dblסԺ�ȶ� & "," & dbl��׼���� & "," & dbl��׼���� & "," & GetTextFromCombo(cmb�������, False) & "," & chkҽ��.Value & ")"
    Else
        gstrSQL = "zl_����֧������_Update(" & mstrID & ",'" & Trim(txtEdit(text����).Text) & "','" & _
                Trim(txtEdit(Text����).Text) & "','" & Trim(txtEdit(Text����).Text) & "'," & lng���� & "," & lng�㷨 & "," & _
                  dbl����ȶ� & "," & dblסԺ�ȶ� & "," & dbl��׼���� & "," & dbl��׼���� & "," & GetTextFromCombo(cmb�������, False) & "," & chkҽ��.Value & ")"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '����������
    If mstrID = "" Then
        Set lst = frm���մ���_����.lvwItem.ListItems.Add(, "K" & lngID, txtEdit(text����), "Class", "Class")
    Else
        Set lst = frm���մ���_����.lvwItem.SelectedItem
        lst.Text = Trim(txtEdit(text����).Text)
    End If
    lst.SubItems(1) = Trim(txtEdit(Text����).Text)
    lst.SubItems(2) = Trim(txtEdit(Text����).Text)
    lst.SubItems(3) = Choose(lng����, "ҩƷ", "ҽ��", "����")
    lst.SubItems(4) = "�ܶ����"
    lst.SubItems(5) = Mid(cmb�������.Text, 3)
    lst.SubItems(6) = IIf(chkҽ��.Value = 1, "��", "��")
    lst.Tag = dblͳ��ȶ� & ";" & dblסԺ�ȶ�
    
    Save��Ŀ = True
    mblnOK = True
    Exit Function

errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function IsValid() As Boolean
'����:���������й�ҽ�����������Ƿ���Ч
'����:
'����ֵ:��Ч����True,����ΪFalse
    Dim lngIndex As Integer
    For lngIndex = text���� To TextסԺ
        If txtEdit(lngIndex).Enabled = True Then
            If zlCommFun.StrIsValid(Trim(txtEdit(lngIndex).Text), txtEdit(lngIndex).MaxLength) = False Then
                txtEdit(lngIndex).SetFocus
                zlControl.TxtSelAll txtEdit(lngIndex)
                Exit Function
            End If
            
            If lngIndex = text���� Or lngIndex = Text���� Then
                If Len(Trim(txtEdit(lngIndex).Text)) = 0 Then
                    txtEdit(lngIndex).Text = ""
                    MsgBox "��������ƶ�����Ϊ�ա�", vbExclamation, gstrSysName
                    txtEdit(lngIndex).SetFocus
                    Exit Function
                End If
            End If
            
            If lngIndex >= Text���� And chk����.Value <> 1 Then
                
                If IsNumeric(txtEdit(lngIndex).Text) = False Then
                    MsgBox "������Ϸ�����ֵ��", vbInformation, gstrSysName
                    zlControl.TxtSelAll txtEdit(lngIndex)
                    txtEdit(lngIndex).SetFocus
                    Exit Function
                End If
                        
                If Val(txtEdit(lngIndex).Text) < 0 Then
                    MsgBox "��ֵ����С��0��", vbInformation, gstrSysName
                    zlControl.TxtSelAll txtEdit(lngIndex)
                    txtEdit(lngIndex).SetFocus
                    Exit Function
                End If
                
                If lngIndex = Text���� Or lngIndex = TextסԺ Then
                    If Val(txtEdit(lngIndex).Text) > 100 Then
                        MsgBox IIf(lngIndex = Text����, "����ͳ��", "סԺͳ��") & "֧���������ܳ���100��", vbInformation, gstrSysName
                        zlControl.TxtSelAll txtEdit(lngIndex)
                        txtEdit(lngIndex).SetFocus
                        Exit Function
                    End If
                Else
                    If Val(txtEdit(lngIndex).Text) > 10000 Then
                        MsgBox "��ֵ���ܳ���10000��", vbInformation, gstrSysName
                        zlControl.TxtSelAll txtEdit(lngIndex)
                        txtEdit(lngIndex).SetFocus
                        Exit Function
                    End If
                End If
            End If
        End If
        If lngIndex = Text���� And chk����.Value = 1 Then
            
            If IsNumeric(txtEdit(lngIndex).Text) = False Then
                MsgBox "������Ϸ�����ֵ��", vbInformation, gstrSysName
                zlControl.TxtSelAll txtEdit(lngIndex)
                txtEdit(lngIndex).SetFocus
                Exit Function
            End If
                    
            If Val(txtEdit(lngIndex).Text) < 0 Then
                MsgBox "��ֵ����С��0��", vbInformation, gstrSysName
                zlControl.TxtSelAll txtEdit(lngIndex)
                txtEdit(lngIndex).SetFocus
                Exit Function
            End If
            
            If Val(txtEdit(lngIndex).Text) > 100000 Then
                MsgBox "��ֵ���ܳ���100000��", vbInformation, gstrSysName
                zlControl.TxtSelAll txtEdit(lngIndex)
                txtEdit(lngIndex).SetFocus
                Exit Function
            End If
        End If
    
    Next
    
    '����׼��������׼��������
    If chkҽ��.Value = 0 Then
        If MsgBox("���������������ҽ������Ӱ�쵽������������ҽ����Ŀ��" & vbCrLf & "�Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            chkҽ��.SetFocus
            Exit Function
        End If
    End If
    
    IsValid = True
End Function


Private Sub opt����_Click(Index As Integer)
    mblnChange = True
End Sub

Private Sub opt����_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = Text���� Then
        txtEdit(Text����).Text = zlCommFun.SpellCode(txtEdit(Text����).Text)
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    Select Case Index
        Case Text����
          zlCommFun.OpenIme True
        Case Else
            zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0 'ʹ֮����
        zlCommFun.PressKey (vbKeyTab)
    Else
        If Index = text���� Then
            KeyAscii = asc(UCase(Chr(KeyAscii)))
            If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
        ElseIf Index = Text���� Or Index = TextסԺ Or Index = Text���� Then
            zlControl.TxtCheckKeyPress txtEdit(Index), KeyAscii, m���ʽ
        End If
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    zlCommFun.OpenIme False
    If Index >= Text���� And Index <= TextסԺ Then
        txtEdit(Index).Text = Format(Val(txtEdit(Index).Text), "0.00")
    End If
End Sub

Public Function �༭ҽ������(ByVal lng���� As Long, ByVal strID As String) As Boolean
'����:��������õ�ҽ���������ڽ���ͨѶ�ĳ���
'����:str���           ��ǰ�༭��ҽ�����ĵ����
'����ֵ:�༭�ɹ�����True,����ΪFalse
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    
    mblnOK = False
    mlng���� = lng����
    mstrID = strID
    
    cmb�������.AddItem "1.���ﲡ��"
    cmb�������.AddItem "2.סԺ����"
    cmb�������.AddItem "3.���в���"
    cmb�������.ListIndex = 2
    
    rsTemp.CursorLocation = adUseClient
    If mstrID <> "" Then
        '�޸�ҽ������
        gstrSQL = "select ����,����,����,nvl(����,1) as ����,nvl(�㷨,1) as �㷨 " & _
                  ",ͳ��ȶ�,סԺ�ȶ�,��׼����,��׼����,�Ƿ�ҽ��,nvl(�������,3) as ������� " & _
                  "from ����֧������ where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CLng(mstrID))
        
        If rsTemp.RecordCount = 0 Then
            MsgBox "�ñ��մ����Ѿ���ɾ������ˢ�¡�", vbInformation, gstrSysName
            Exit Function
        End If
        txtEdit(text����).Text = rsTemp("����")
        txtEdit(Text����).Text = rsTemp("����")
        txtEdit(Text����).Text = IIf(IsNull(rsTemp("����")), "", rsTemp("����"))
        Call SetComboByText(cmb�������, rsTemp("�������"), False)
        '�ܺ�ȫ���� 2003-12-17
        '�޸�ʱ������������
        opt����(rsTemp("����")).Value = True
        chkҽ��.Value = IIf(rsTemp("�Ƿ�ҽ��") = 1, 1, 0)
        '1-����������Ŀ
        txtEdit(Text����).Text = Format(rsTemp("ͳ��ȶ�"), "0.00")
        txtEdit(TextסԺ).Text = Format(rsTemp("סԺ�ȶ�"), "0.00")
        '1-��׼����
        txtEdit(Text����).Text = Format(rsTemp("��׼����"), "###0.00;-####0.00; ;")
        If Val(txtEdit(Text����)) <> 0 Then
            chk����.Value = 1
        Else
            chk����.Value = 0
        End If
        chk����_Click
    Else
        '����ҽ������
        txtEdit(text����).Text = zlDatabase.GetMax("����֧������", "����", 6, " where ����=" & mlng����)
    End If
    
    
    mblnChange = False
    frm���մ���༭_����.Show vbModal, frm���մ���
    �༭ҽ������ = mblnOK
End Function



