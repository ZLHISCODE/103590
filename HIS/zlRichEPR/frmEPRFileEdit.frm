VERSION 5.00
Begin VB.Form frmEPRFileEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����ļ�����"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   Icon            =   "frmEPRFileEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fraEditType 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   765
      TabIndex        =   21
      Top             =   3195
      Width           =   4905
      Begin VB.CheckBox chk�����ݲ��� 
         Caption         =   "�����ݲ���"
         Height          =   180
         Left            =   1780
         TabIndex        =   24
         Top             =   112
         Visible         =   0   'False
         Width           =   1400
      End
      Begin VB.OptionButton optEditType 
         Caption         =   "���ʽ�����༭��"
         Height          =   225
         Index           =   1
         Left            =   0
         TabIndex        =   23
         Top             =   360
         Width           =   1845
      End
      Begin VB.OptionButton optEditType 
         Caption         =   "ȫ�Ĳ����༭��"
         Height          =   225
         Index           =   0
         Left            =   15
         TabIndex        =   22
         Top             =   90
         Value           =   -1  'True
         Width           =   1665
      End
   End
   Begin VB.ComboBox cbo�ȼ� 
      Height          =   300
      Left            =   1455
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2025
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.CheckBox chkCopy 
      Caption         =   "����(&V)"
      Height          =   195
      Left            =   3210
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   473
      Width           =   2670
   End
   Begin VB.ComboBox cboKind 
      Height          =   300
      Left            =   1455
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   420
      Width           =   1425
   End
   Begin VB.OptionButton optPage 
      Caption         =   "ʹ�ù���ҳ��(&2)"
      Height          =   180
      Index           =   1
      Left            =   780
      TabIndex        =   12
      Top             =   2910
      Width           =   1725
   End
   Begin VB.TextBox txtPageName 
      Height          =   300
      Left            =   3210
      TabIndex        =   14
      Top             =   2460
      Width           =   2490
   End
   Begin VB.TextBox txtPageNo 
      Enabled         =   0   'False
      Height          =   300
      Left            =   2535
      TabIndex        =   13
      Top             =   2460
      Width           =   645
   End
   Begin VB.OptionButton optPage 
      Caption         =   "ʹ���½�ҳ��(&1)"
      Height          =   180
      Index           =   0
      Left            =   780
      TabIndex        =   11
      Top             =   2505
      Value           =   -1  'True
      Width           =   1725
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4600
      TabIndex        =   17
      Top             =   4155
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3480
      TabIndex        =   16
      Top             =   4155
      Width           =   1100
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   1
      Left            =   -45
      TabIndex        =   20
      Top             =   4035
      Width           =   6390
   End
   Begin VB.ComboBox cboPage 
      Enabled         =   0   'False
      Height          =   300
      Left            =   2535
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2850
      Width           =   3165
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Index           =   0
      Left            =   765
      TabIndex        =   18
      Top             =   840
      Width           =   5325
   End
   Begin VB.TextBox txt��� 
      Height          =   300
      Left            =   1455
      TabIndex        =   4
      Top             =   975
      Width           =   645
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   3210
      TabIndex        =   6
      Top             =   975
      Width           =   2490
   End
   Begin VB.TextBox txt˵�� 
      Height          =   540
      Left            =   1455
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   1380
      Width           =   4245
   End
   Begin VB.Label lbl�ȼ� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&A)                 �����ϻ���ȼ��Ĳ��ˡ�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   780
      TabIndex        =   9
      Top             =   2085
      Visible         =   0   'False
      Width           =   4140
   End
   Begin VB.Label lblKind 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&K)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   795
      TabIndex        =   0
      Top             =   480
      Width           =   630
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "�Բ����ļ�������������ָ������ʽ��ӡ�����ҳ�档"
      Height          =   180
      Left            =   780
      TabIndex        =   19
      Top             =   120
      Width           =   4335
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   180
      Picture         =   "frmEPRFileEdit.frx":058A
      Top             =   150
      Width           =   480
   End
   Begin VB.Label lbl��� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   780
      TabIndex        =   3
      Top             =   1035
      Width           =   630
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2505
      TabIndex        =   5
      Top             =   1035
      Width           =   630
   End
   Begin VB.Label lbl˵�� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "˵��(&M)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   780
      TabIndex        =   7
      Top             =   1440
      Width           =   630
   End
End
Attribute VB_Name = "frmEPRFileEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'˵����
'   1���ϼ�����ͨ��������ShowMe�������������塢�༭�ļ�ID,�༭״̬����Ϣ���ݽ��뱾����
'   2���༭״̬����Me.tag��ţ��ֱ�Ϊ"����"��"�޸�"�����ϼ�����ͨ��ShowMe����
'---------------------------------------------------
Private mlngFileID As Long          '���༭�����ڸ������ӵ��ļ�ID���޸ġ�����ʱ���ϼ�����ͨ��ShowMe���ݽ���,����ʱΪ0.
Private mblnSpecial As Boolean      '�����ļ��Ƿ����ⲡ��
Private mblnSpecialWave As Boolean  '������ļ��Ƿ���ר�����µ�
Private mblnOK As Boolean

Public Function ShowMe(ByVal frmParent As Object, ByVal strKinds As String, ByVal blnAdd As Boolean, Optional ByVal lngFileID As Long) As Long
    '---------------------------------------------------
    '���ܣ��ϼ�������ñ�����ģ����ݲ���������ʾ����
    '---------------------------------------------------
Dim rsTemp As New ADODB.Recordset
Dim lngCount As Long
    With Me.cbo�ȼ�
        .Clear
        .AddItem "0-�ؼ�����"
        .AddItem "1-һ������"
        .AddItem "2-��������"
        .AddItem "3-��������"
        .ListIndex = .ListCount - 1
    End With
    
    If InStr(1, "," & strKinds, ",1") > 0 Then Me.cboKind.AddItem "1-���ﲡ��"
    If InStr(1, "," & strKinds, ",2") > 0 Then Me.cboKind.AddItem "2-סԺ����"
    If InStr(1, "," & strKinds, ",3") > 0 Then Me.cboKind.AddItem "3-�����¼"
    If InStr(1, "," & strKinds, ",4") > 0 Then Me.cboKind.AddItem "4-������"
    If InStr(1, "," & strKinds, ",5") > 0 Then Me.cboKind.AddItem "5-����֤������"
    If InStr(1, "," & strKinds, ",6") > 0 Then Me.cboKind.AddItem "6-֪���ļ�"
    If Me.cboKind.ListCount <= 1 Then Me.cboKind.Enabled = False
    
    If blnAdd Then
        Me.Tag = "����": mlngFileID = 0
    Else
        Me.Tag = "�޸�": mlngFileID = lngFileID
    End If
    
    mblnSpecialWave = False
    '��ǰ���ݶ�ȡ
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select l.����, l.���, l.����, l.˵��, l.����,l.����, f.��� As ҳ���, f.���� As ҳ����, Nvl(f.����, 0) As �ȼ�" & _
            " From �����ļ��б� l, ����ҳ���ʽ f" & _
            " Where l.���� = f.����(+) And l.ҳ�� = f.���(+) And l.Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngFileID)
    With rsTemp
        Me.txt���.MaxLength = .Fields("���").DefinedSize
        Me.txt����.MaxLength = .Fields("����").DefinedSize
        Me.txt˵��.MaxLength = .Fields("˵��").DefinedSize
        Me.txtPageNo.MaxLength = .Fields("ҳ���").DefinedSize
        Me.txtPageName.MaxLength = .Fields("ҳ����").DefinedSize
        If .RecordCount > 0 Then
            mblnSpecial = (NVL(!����, 0) < 0 Or NVL(!����, 0) = 2)
            mblnSpecialWave = NVL(!����, 0) < 0 And NVL(!����, 0) = 3 And NVL(!����) = "1"
            Me.cbo�ȼ�.ListIndex = !�ȼ�
            If Me.Tag = "����" Then
                Me.txtPageNo.Text = !ҳ���: Me.txtPageName.Text = !ҳ����
                '���ⲡ���������ڸ���
                Me.chkCopy.Tag = lngFileID: Me.chkCopy.Caption = "����(&V)" & !����
                If mblnSpecial Then Me.chkCopy.Value = vbUnchecked: Me.chkCopy.Visible = False
                If mblnSpecialWave Then Me.chkCopy.Value = vbChecked: Me.chkCopy.Visible = True: Me.chkCopy.Enabled = False
            Else
                Me.txt���.Text = !���: Me.txt����.Text = !����: Me.txt˵��.Text = "" & !˵��
                Me.txtPageNo.Text = !ҳ���: Me.txtPageName.Text = !ҳ����
                Me.chkCopy.Value = vbUnchecked: Me.chkCopy.Visible = False
                Me.cboKind.Enabled = False: optEditType(0).Enabled = False: optEditType(1).Enabled = False
                If NVL(!����, 0) < 0 Then optEditType(0).Value = False: optEditType(1).Value = False
                If NVL(!����, 0) = 0 Or NVL(!����, 0) = 1 Then optEditType(0).Value = True: optEditType(1).Value = False
                
            End If
            Me.cboKind.Tag = !����
            For lngCount = 0 To Me.cboKind.ListCount - 1
                If Val(Left(Me.cboKind.List(lngCount), 1)) = !���� Then
                    Me.cboKind.ListIndex = lngCount
                    Exit For
                End If
            Next
            If !��� = "" & !ҳ��� Or Val(Me.cboKind.Tag) = 3 Then
                Me.optPage(0).Value = True
            Else
                Me.optPage(1).Value = True
                For lngCount = 0 To Me.cboPage.ListCount - 1
                    If Val(Me.cboPage.List(lngCount)) = Val("" & !ҳ���) Then
                        Me.cboPage.ListIndex = lngCount
                        Exit For
                    End If
                Next
            End If
            If Me.Tag = "�޸�" Then
                If NVL(!����, 0) = 2 Then
                    optEditType(0).Value = False: optEditType(1).Value = True: optPage(1).Enabled = False: cboPage.Enabled = False
                ElseIf NVL(!����, 0) = 3 Then
                    chk�����ݲ���.Value = 1
                End If
            End If
        Else
            If Me.Tag = "����" Then
                Me.cboKind.ListIndex = 0
            Else
                MsgBox "ָ���ļ���ʧ��(���ܱ������û�ɾ��)", vbInformation, gstrSysName
                ShowMe = 0: Unload Me: Exit Function
            End If
        End If
    End With
    
    '��ʾ����
    Me.Show vbModal, frmParent
    If mblnOK = False Then ShowMe = 0: Unload Me: Exit Function
    ShowMe = mlngFileID
    Unload Me: Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cboKind_Click()
Dim intKind As Integer
Dim rsTemp As New ADODB.Recordset
    intKind = Left(Me.cboKind.Text, 1)
    
    If Me.Tag = "����" Then
        gstrSQL = "Select nvl(max(���),'" & String(Me.txt���.MaxLength, "0") & "') as ��� From �����ļ��б� Where ���� = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, intKind)
        Me.txt���.Text = Format(Val(rsTemp!���) + 1, String(Me.txt���.MaxLength, "0"))
        
        If Val(Me.cboKind.Tag) = intKind Then
            Me.chkCopy.Enabled = Not mblnSpecialWave
        Else
            Me.chkCopy.Value = vbUnchecked: Me.chkCopy.Enabled = False
        End If
    End If
    
    If intKind = 3 Then
        Me.lbl�ȼ�.Visible = True: Me.cbo�ȼ�.Visible = True
    Else
        Me.lbl�ȼ�.Visible = False: Me.cbo�ȼ�.Visible = False
    End If
    chk�����ݲ���.Visible = intKind = 1
    chk�����ݲ���.Enabled = optEditType(0).Value = True And Me.Tag = "����"
    
    Me.cboPage.Clear
    Select Case intKind
    Case 2, 4   '2-סԺ����;4-������
        gstrSQL = "Select f.���, f.����, Count(l.ID) As ʹ��" & _
                " From ����ҳ���ʽ f, �����ļ��б� l" & _
                " Where f.���� = l.���� And f.��� = l.ҳ�� And l.���� Between 0 And 1 And f.���� = [1]" & _
                " Group By f.���, f.����" & _
                " Order By ���"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, intKind)
        With rsTemp
            If Me.Tag = "����" Then
                Do While Not .EOF
                    Me.cboPage.AddItem !��� & "-" & !����
                    .MoveNext
                Loop
                If Me.cboPage.ListCount = 0 Then
                    Me.optPage(0).Value = True: Me.optPage(0).Enabled = False
                    Me.optPage(1).Value = False: Me.optPage(1).Enabled = False
                Else
                    Me.optPage(0).Enabled = True: Me.optPage(1).Enabled = True
                    Me.cboPage.ListIndex = 0
                End If
            Else
                Do While Not .EOF
                    If !��� <> Trim(Me.txt���.Text) Then
                        Me.cboPage.AddItem !��� & "-" & !����
                    Else
                        Me.txtPageNo.Text = !���: Me.txtPageName.Text = !����
                        If !ʹ�� > 1 Then
                            Me.cboPage.AddItem !��� & "-" & !����
                            Me.cboPage.ListIndex = Me.cboPage.NewIndex
                        Else
                            Me.optPage(0).Value = True
                        End If
                    End If
                    .MoveNext
                Loop
                If mblnSpecial Then
                    Me.optPage(0).Value = True: Me.optPage(0).Enabled = False
                    Me.optPage(1).Value = False: Me.optPage(1).Enabled = False
                    Me.txtPageNo.Text = Me.txt���.Text: Me.txtPageNo.Enabled = False
                    Me.txtPageName.Text = Me.txt����.Text: Me.txtPageName.Enabled = False
                ElseIf Me.cboPage.ListCount = 0 Or mblnSpecial Then
                    Me.optPage(0).Value = True: Me.optPage(0).Enabled = False
                    Me.optPage(1).Value = False: Me.optPage(1).Enabled = False
                Else
                    Me.optPage(0).Enabled = True: Me.optPage(1).Enabled = True
                End If
            End If
        End With
    Case Else
        Me.optPage(0).Value = True: Me.optPage(0).Enabled = False
        Me.optPage(1).Value = False: Me.optPage(1).Enabled = False
        Me.txtPageNo.Text = Me.txt���.Text: Me.txtPageNo.Enabled = False
        Me.txtPageName.Text = Me.txt����.Text: Me.txtPageName.Enabled = False
    End Select
    
    If Me.Tag = "����" Then '����ʱ�Ի����¼����
        If intKind = 3 Then '�����¼
            optEditType(0).Enabled = False: optEditType(1).Enabled = False
        ElseIf intKind = 5 Then  '�����걨��
            optEditType(0).Enabled = True: optEditType(1).Enabled = True
        Else
            optEditType(0).Enabled = True: optEditType(1).Enabled = True
        End If
    End If
    
    optEditType(0).Value = True
End Sub

Private Sub cboKind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboPage_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo�ȼ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chkCopy_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False: Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim intType As Integer
    
    If Trim(Me.txt���.Text) = "" Then MsgBox "�������ţ�", vbInformation, gstrSysName: Me.txt���.SetFocus: Exit Sub
    If Len(Me.txt���.Text) < Me.txt���.MaxLength Then MsgBox "��ų��Ȳ��㣡", vbInformation, gstrSysName: Me.txt���.SetFocus: Exit Sub
    If Trim(Me.txt����.Text) = "" Then MsgBox "���������ƣ�", vbInformation, gstrSysName: Me.txt����.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > Me.txt����.MaxLength Then
        MsgBox "���Ƴ��������" & Me.txt����.MaxLength & "���ַ���ȳ��ĺ��֣���", vbInformation, gstrSysName: Me.txt����.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt˵��.Text), vbFromUnicode)) > Me.txt˵��.MaxLength Then
        MsgBox "˵�����������" & Me.txt˵��.MaxLength & "���ַ���ȳ��ĺ��֣���", vbInformation, gstrSysName: Me.txt˵��.SetFocus: Exit Sub
    End If
    If Me.optPage(0).Value Then
        If Trim(Me.txtPageName.Text) = "" Then MsgBox "������ҳ�����ƣ�", vbInformation, gstrSysName: Me.txtPageName.SetFocus: Exit Sub
        If LenB(StrConv(Trim(Me.txtPageName.Text), vbFromUnicode)) > Me.txtPageName.MaxLength Then
            MsgBox "ҳ�����Ƴ��������" & Me.txt����.MaxLength & "���ַ���ȳ��ĺ��֣���", vbInformation, gstrSysName: Me.txtPageName.SetFocus: Exit Sub
        End If
    Else
        If Me.cboPage.ListIndex = -1 Then MsgBox "��ѡ����ҳ�棡", vbInformation, gstrSysName: Me.cboPage.SetFocus: Exit Sub
    End If
    
    '���ݱ���
    If Me.Tag = "����" Then
        mlngFileID = zlDatabase.GetNextId("�����ļ��б�")
        gstrSQL = mlngFileID & "," & Val(Left(Me.cboKind.Text, 1))
    Else
        gstrSQL = mlngFileID
    End If
    gstrSQL = gstrSQL & ",'" & Trim(Me.txt���.Text) & "','" & Trim(Me.txt����.Text) & "','" & Replace(Me.txt˵��, Chr(vbKeyReturn), "") & "'"
    If Me.optPage(0).Value Then
        gstrSQL = gstrSQL & ",'" & Trim(Me.txtPageNo.Text) & "','" & Trim(Me.txtPageName.Text) & "'"
    Else
        gstrSQL = gstrSQL & ",'" & Left(Me.cboPage.Text, Me.txt���.MaxLength) & "','" & Trim(Mid(Me.cboPage.Text, Me.txt���.MaxLength + 2)) & "'"
    End If
    If Val(Left(Me.cboKind.Text, 1)) <> 3 Then
        gstrSQL = gstrSQL & ",0"
    Else
        gstrSQL = gstrSQL & "," & Me.cbo�ȼ�.ListIndex
    End If
    
    If Me.Tag = "����" Then
        If mblnSpecialWave = False Then '����ר�����µ�
            If optEditType(1).Value Then
                intType = 2
            Else
                intType = IIf(chk�����ݲ���.Visible And chk�����ݲ���.Value = 1, 3, 0)
            End If
            gstrSQL = "Zl_�����ļ��б�_Insert(" & gstrSQL & "," & IIf(Me.chkCopy.Value = vbChecked, Val(Me.chkCopy.Tag), 0) & "," & intType & ")"
        Else
            intType = -1
            gstrSQL = "Zl_�����ļ��б�_Insert(" & gstrSQL & "," & IIf(Me.chkCopy.Value = vbChecked, Val(Me.chkCopy.Tag), 0) & "," & intType & ",'1')"
        End If
    Else
        gstrSQL = "Zl_�����ļ��б�_Modify(" & gstrSQL & ")"
    End If
    
    Err = 0: On Error GoTo errHand
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    mblnOK = True: Me.Hide: Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Me.txt���.SetFocus
End Sub

Private Sub optEditType_Click(Index As Integer)
    If Index = 1 Then
        optPage(0).Value = True: optPage(1).Value = False
        optPage(1).Enabled = False: cboPage.Enabled = False: chkCopy.Enabled = False
        chk�����ݲ���.Enabled = False
    Else
        If Val(cboKind.Text) = 2 Or Val(cboKind.Text) = 4 Then
            optPage(1).Enabled = True: cboPage.Enabled = True
        End If
        chkCopy.Enabled = chkCopy.Tag <> "" And Val(cboKind.Tag) = Val(cboKind.Text) And Not mblnSpecialWave
        chk�����ݲ���.Enabled = True
    End If
End Sub

Private Sub optPage_Click(Index As Integer)
    If Me.optPage(0).Value Then
        Me.txtPageName.Enabled = True: Me.cboPage.Enabled = False
        Me.txtPageNo.Text = Me.txt���.Text: Me.txtPageName.Text = Me.txt����.Text
        If Me.txtPageName.Visible Then Me.txtPageName.SetFocus
    Else
        Me.txtPageName.Enabled = False: Me.cboPage.Enabled = True
        If Me.cboPage.Visible Then Me.cboPage.SetFocus
    End If
End Sub

Private Sub optPage_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txtPageName_Change()
    ValidControlText txtPageName
End Sub

Private Sub txtPageName_GotFocus()
    Me.txtPageName.SelStart = 0: Me.txtPageName.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtPageName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt���_Change()
    ValidControlText txt���
    Me.txtPageNo.Text = Me.txt���.Text
End Sub

Private Sub txt���_GotFocus()
    Me.txt���.SelStart = 0: Me.txt���.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt���_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt����_Change()
    ValidControlText txt����
    Me.txtPageName.Text = Me.txt����.Text
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt˵��_Change()
    ValidControlText txt˵��
End Sub

Private Sub txt˵��_GotFocus()
    Me.txt˵��.SelStart = 0: Me.txt˵��.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt˵��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt˵��_LostFocus()
    Me.txt˵��.Text = Replace(Me.txt˵��, Chr(vbKeyReturn), "")
    Call zlCommFun.OpenIme(False)
End Sub
