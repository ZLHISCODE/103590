VERSION 5.00
Begin VB.Form frmMassRuleEdit 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8175
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5190
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox txt���� 
      Height          =   300
      Index           =   5
      Left            =   4305
      MaxLength       =   5
      TabIndex        =   20
      Top             =   1590
      Width           =   690
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Index           =   4
      Left            =   4305
      MaxLength       =   5
      TabIndex        =   18
      Top             =   1410
      Width           =   690
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Index           =   3
      Left            =   4305
      MaxLength       =   5
      TabIndex        =   16
      Top             =   1245
      Width           =   690
   End
   Begin VB.TextBox txt���� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   630
      MaxLength       =   60
      TabIndex        =   1
      Text            =   "���ÿ��ƹ���"
      Top             =   165
      Width           =   2235
   End
   Begin VB.Frame fraLine 
      Height          =   15
      Left            =   -30
      TabIndex        =   24
      Top             =   2850
      Width           =   5385
   End
   Begin VB.TextBox txt˵�� 
      Height          =   720
      Left            =   630
      MaxLength       =   60
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   1935
      Width           =   4365
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Index           =   2
      Left            =   4305
      MaxLength       =   5
      TabIndex        =   14
      Top             =   1095
      Width           =   690
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Index           =   1
      Left            =   4305
      MaxLength       =   5
      TabIndex        =   12
      Top             =   930
      Width           =   690
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Index           =   0
      Left            =   4305
      MaxLength       =   5
      TabIndex        =   10
      Top             =   607
      Width           =   690
   End
   Begin VB.ComboBox cbo��ʽ 
      Height          =   300
      Left            =   630
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1491
      Width           =   2235
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   630
      MaxLength       =   13
      TabIndex        =   3
      Top             =   607
      Width           =   780
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   630
      MaxLength       =   60
      TabIndex        =   5
      Top             =   1049
      Width           =   2235
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "h����׼��"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   5
      Left            =   3420
      TabIndex        =   19
      Top             =   1650
      Width           =   810
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "k����׼��"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   4
      Left            =   3420
      TabIndex        =   17
      Top             =   1470
      Width           =   810
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "P��ʧ����"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   3
      Left            =   3420
      TabIndex        =   15
      Top             =   1305
      Width           =   810
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������ز�����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3405
      TabIndex        =   8
      Top             =   285
      Width           =   1260
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   0
      Top             =   225
      Width           =   360
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "M���ⶨֵ"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   2
      Left            =   3420
      TabIndex        =   13
      Top             =   1155
      Width           =   810
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "X����׼��"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   3420
      TabIndex        =   11
      Top             =   990
      Width           =   810
   End
   Begin VB.Label lbl˵�� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "˵��"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      TabIndex        =   21
      Top             =   1980
      Width           =   360
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "N���ⶨֵ"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   3420
      TabIndex        =   9
      Top             =   660
      Width           =   810
   End
   Begin VB.Image imgNote 
      Height          =   240
      Left            =   195
      Picture         =   "frmMassRuleEdit.frx":0000
      Top             =   2925
      Width           =   240
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmMassRuleEdit.frx":058A
      ForeColor       =   &H00008000&
      Height          =   4680
      Left            =   465
      TabIndex        =   23
      Top             =   2985
      Width           =   4500
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl��ʽ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ʽ"
      Height          =   180
      Left            =   180
      TabIndex        =   6
      Top             =   1551
      Width           =   360
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      TabIndex        =   2
      Top             =   667
      Width           =   360
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      TabIndex        =   4
      Top             =   1109
      Width           =   360
   End
End
Attribute VB_Name = "frmMassRuleEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mPar       '����ö��
    n = 0: x: m: p: k: h
End Enum

Private mlngItemID As Long          '��ǰ��ʾ����Ŀid

Dim objItem As ListItem
Dim lngCount As Long

'--------------------------------------------
'����Ϊ���幫������
'--------------------------------------------

Public Function zlRefresh(lngItemID As Long) As Boolean
    '���ܣ�������Ŀidˢ�µ�ǰ��ʾ����
    Dim rsTemp As New ADODB.Recordset
    
    mlngItemID = lngItemID
    
    '�����ǰ��Ŀ����ʾ
    
    Me.txt����.Text = "": Me.txt����.Text = "": Me.txt˵��.Text = ""
    Me.txt����.Text = "": Me.cbo��ʽ.Clear
    For lngCount = 0 To Me.txt����.UBound
        Me.lbl����(lngCount).Visible = False
        Me.txt����(lngCount).Visible = False: Me.txt����(lngCount).Text = ""
    Next
    If lngItemID = 0 Then zlRefresh = True: Exit Function


    '��ȡָ����Ŀ����Ϣ
    Err = 0: On Error GoTo ErrHand

    gstrSql = "Select ����, ����, ����, ˵��, ��ʽ, N, X, M, P, K, H From �����ʿع��� Where Id = [1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
    With rsTemp
        Me.txt����.MaxLength = .Fields("����").DefinedSize
        Me.txt����.MaxLength = .Fields("����").DefinedSize
        Me.txt˵��.MaxLength = .Fields("˵��").DefinedSize
        If .RecordCount > 0 Then
            Me.txt����.Tag = Val("" & !����)
            Select Case Val("" & !����)
            Case 1
                Me.txt����.Text = "���ÿ��ƹ���"
                Me.cbo��ʽ.AddItem "N-Xs": Me.cbo��ʽ.AddItem "R-Xs": Me.cbo��ʽ.AddItem "N-T": Me.cbo��ʽ.AddItem "N-X": Me.cbo��ʽ.AddItem "(M of N)Xs"
            Case 2
                Me.txt����.Text = "������ƽ��޹���"
                Me.cbo��ʽ.AddItem "N-P": Me.cbo��ʽ.AddItem "X-P": Me.cbo��ʽ.AddItem "R-P"
            Case 3
                Me.txt����.Text = "�ۻ��͹���"
                Me.cbo��ʽ.AddItem "CS(k:h)"
            End Select
            Me.txt����.Text = "" & !����
            Me.txt����.Text = "" & !����
            Me.txt˵��.Text = "" & !˵��
            Me.cbo��ʽ.ListIndex = Val("" & !��ʽ)
            Me.txt����(mPar.n).Text = Val("" & !n)
            Me.txt����(mPar.x).Text = Replace(Replace(" 0" & !x, " 0.", "0."), " 0", "")
            Me.txt����(mPar.m).Text = Val("" & !m)
            Me.txt����(mPar.p).Text = Replace(Replace(" 0" & !p, " 0.", "0."), " 0", "")
            Me.txt����(mPar.k).Text = Replace(Replace(" 0" & !k, " 0.", "0."), " 0", "")
            Me.txt����(mPar.h).Text = Replace(Replace(" 0" & !h, " 0.", "0."), " 0", "")
        End If
    End With
        
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditStart(blnAdd As Boolean, lngItemID As Long) As Boolean
    '���ܣ���ʼ��Ŀ�༭
    '������ blnAdd-�Ƿ����ӣ�����Ϊ�޸�
    '       lngItemId-���ӵĲ�����Ŀ������ָ���༭����Ŀ
    Dim rsTemp As New ADODB.Recordset
    
    If blnAdd Then
        Err = 0: On Error GoTo ErrHand
        gstrSql = "Select Nvl(Max(To_Number(����)), 0) As ����, Nvl(Max(Length(����)), 0) As ���� From �����ʿع��� Where ���� = 1"
        
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "zlEditStart")
'            Call SQLTest
        With rsTemp
            If !���� <> 0 And !���� <= Me.txt����.MaxLength Then
                Me.txt����.Text = Format(Val(!����) + 1, String(!����, "0"))
            Else
                Me.txt����.Text = Format(Val(!����) + 1, String(Me.txt����.MaxLength, "0"))
            End If
        End With
        
        Me.txt����.Text = "": Me.txt˵��.Text = ""
        For lngCount = 0 To Me.txt����.UBound
            Me.txt����(lngCount).Text = ""
        Next
        If Val(Me.txt����.Tag) <> 1 Then
            Me.txt����.Tag = 1
            Me.txt����.Text = "���ÿ��ƹ���"
            With Me.cbo��ʽ
                .Clear
                .AddItem "N-Xs": .AddItem "R-Xs": .AddItem "N-T": .AddItem "N-X": .AddItem "(M of N)Xs"
                .ListIndex = 0
            End With
        End If
    End If

    Me.Tag = IIf(blnAdd, "����", "�޸�")
    Me.Enabled = True: Me.BackColor = RGB(250, 250, 250)
    
    Me.txt����.SetFocus
    zlEditStart = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditStart = False: Exit Function
End Function

Public Sub zlEditCancel()
    '���ܣ��������ڽ��еı༭
    Me.Tag = ""
    Me.Enabled = False: Me.BackColor = &H8000000F
    Call Me.zlRefresh(mlngItemID)
End Sub

Public Function zlEditSave() As Long
    '���ܣ��������ڽ��еı༭,���������ڱ༭��Ŀid,����ʧ�ܷ���0
    Dim lngNewId As Long, strLists As String
    
    'һ�����Լ��
    If Me.cbo��ʽ.ListIndex = -1 Then
        MsgBox "������ʽδ���ã�", vbInformation, gstrSysName
        Me.cbo��ʽ.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Trim(Me.txt����.Text) = "" Then
        MsgBox "��������룡", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Val(Me.txt����.Text) > Val(String(Me.txt����.MaxLength, "9")) Then
        MsgBox "����̫��", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Trim(Me.txt����.Text) = "" Then
        MsgBox "���������ƣ�", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > Me.txt����.MaxLength Then
        MsgBox "���Ƴ��������" & Me.txt����.MaxLength & "���ַ���ȳ����֣���", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt˵��.Text), vbFromUnicode)) > Me.txt˵��.MaxLength Then
        MsgBox "˵�����������" & Me.txt˵��.MaxLength & "���ַ���ȳ����֣���", vbInformation, gstrSysName
        Me.txt˵��.SetFocus: zlEditSave = 0: Exit Function
    End If
    For lngCount = 0 To Me.txt����.UBound
        If Me.txt����(lngCount).Visible Then
            If Val(Trim(Me.txt����(lngCount).Text)) = 0 Then
                MsgBox "����Լ��" & Me.lbl����(lngCount).Caption & "��������ָ����", vbInformation, gstrSysName
                Me.txt����(lngCount).SetFocus: zlEditSave = 0: Exit Function
            End If
            Me.txt����(lngCount).Text = Val(Trim(Me.txt����(lngCount).Text))
        Else
            Me.txt����(lngCount).Text = 0
        End If
    Next
    If Me.txt����(mPar.x).Visible Then
        If Val(Val(Me.txt����(mPar.x).Text) * 10) <> Int(Val(Val(Me.txt����(mPar.x).Text) * 10)) Then
            MsgBox "X��������̫�ߣ�", vbInformation, gstrSysName
            Me.txt����(mPar.x).SetFocus: zlEditSave = 0: Exit Function
        End If
    End If
    If Me.txt����(mPar.n).Visible And Me.txt����(mPar.m).Visible Then
        If Val(Me.txt����(mPar.n).Text) <= Val(Me.txt����(mPar.m).Text) Then
            MsgBox "N�����������M������", vbInformation, gstrSysName
            Me.txt����(mPar.n).SetFocus: zlEditSave = 0: Exit Function
        End If
    End If
    
    '���ݱ��������֯
    gstrSql = "'" & Trim(Me.txt����.Text) & "','" & Trim(Me.txt����.Text) & "'," & Me.cbo��ʽ.ListIndex
    gstrSql = gstrSql & "," & IIf(Me.cbo��ʽ.ListIndex = 1, 2, Val(Me.txt����(mPar.n).Text)) & "," & Val(Me.txt����(mPar.x).Text) & "," & Val(Me.txt����(mPar.m).Text)
    gstrSql = gstrSql & ",'" & Trim(Me.txt˵��.Text) & "'"
    lngNewId = mlngItemID
    If Me.Tag = "����" Then
        lngNewId = zldatabase.GetNextId("�����ʿع���")
        gstrSql = "Zl_�����ʿع���_Edit(1," & lngNewId & "," & gstrSql & ")"
    Else
        gstrSql = "Zl_�����ʿع���_Edit(2," & lngNewId & "," & gstrSql & ")"
    End If
    
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    If Me.Tag = "����" Then mlngItemID = lngNewId
    
    Me.Tag = ""
    Me.Enabled = False: Me.BackColor = &H8000000F
    
    zlEditSave = mlngItemID: Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function

'--------------------------------------------
'����Ϊ����ؼ���Ӧ�¼�
'--------------------------------------------

Private Sub cbo��ʽ_Click()
    Dim intCount As Integer, intVisible As Integer
    
    For intCount = Me.txt����.LBound To Me.txt����.UBound
        Me.lbl����(intCount).Visible = False
        Me.txt����(intCount).Visible = False
    Next
    
    Select Case Val(Me.txt����.Tag)
    Case 1
        Select Case Me.cbo��ʽ.ListIndex
        Case 0 '"N-Xs"
            Me.lbl����(mPar.n).Visible = True: Me.txt����(mPar.n).Visible = True
            Me.lbl����(mPar.x).Visible = True: Me.txt����(mPar.x).Visible = True
        Case 1 '"R-Xs"
            Me.lbl����(mPar.x).Visible = True: Me.txt����(mPar.x).Visible = True
        Case 2 '"N-T"
            Me.lbl����(mPar.n).Visible = True: Me.txt����(mPar.n).Visible = True
        Case 3 '"N-X"
            Me.lbl����(mPar.n).Visible = True: Me.txt����(mPar.n).Visible = True
        Case 4 '"(M of N)Xs"
            Me.lbl����(mPar.n).Visible = True: Me.txt����(mPar.n).Visible = True
            Me.lbl����(mPar.x).Visible = True: Me.txt����(mPar.x).Visible = True
            Me.lbl����(mPar.m).Visible = True: Me.txt����(mPar.m).Visible = True
        End Select
    Case 2
        Select Case Me.cbo��ʽ.ListIndex
        Case 0 '"N-P"
            Me.lbl����(mPar.n).Visible = True: Me.txt����(mPar.n).Visible = True
            Me.lbl����(mPar.p).Visible = True: Me.txt����(mPar.p).Visible = True
        Case 1 '"X-P"
            Me.lbl����(mPar.p).Visible = True: Me.txt����(mPar.p).Visible = True
        Case 2 '"R-P"
            Me.lbl����(mPar.p).Visible = True: Me.txt����(mPar.p).Visible = True
        End Select
    Case 3
        Me.lbl����(mPar.k).Visible = True: Me.txt����(mPar.k).Visible = True
        Me.lbl����(mPar.h).Visible = True: Me.txt����(mPar.h).Visible = True
    End Select
    
    intVisible = 0
    For intCount = Me.txt����.LBound To Me.txt����.UBound
        If Me.txt����(intCount).Visible Then
            Me.lbl����(intCount).Top = Me.lbl����.Top + (Me.lbl����.Top - Me.lbl����.Top) * intVisible
            Me.txt����(intCount).Top = Me.txt����.Top + (Me.txt����.Top - Me.txt����.Top) * intVisible
            intVisible = intVisible + 1
        End If
    Next
End Sub

Private Sub cbo��ʽ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    mlngItemID = 0
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt����_GotFocus(Index As Integer)
    Me.txt����(Index).SelStart = 0: Me.txt����(Index).SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt����_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        Select Case Index
        Case mPar.x, mPar.p, mPar.k, mPar.h
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then Exit Sub
        Case Else
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
        End Select
    End Select
    KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt˵��_GotFocus()
    Me.txt˵��.SelStart = 0: Me.txt˵��.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt˵��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
