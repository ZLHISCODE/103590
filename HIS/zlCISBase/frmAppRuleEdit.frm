VERSION 5.00
Begin VB.Form frmAppRuleEdit 
   BorderStyle     =   0  'None
   Caption         =   "�����ʿع���"
   ClientHeight    =   5055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8265
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   5055
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CheckBox chkUse 
      Caption         =   "�Ƿ�ʹ��"
      Height          =   180
      Left            =   6960
      TabIndex        =   23
      ToolTipText     =   "�ڼ���ʱ���Ƿ�ʹ�ô˹���"
      Top             =   75
      Width           =   1065
   End
   Begin VB.TextBox txt��ʾ 
      Enabled         =   0   'False
      Height          =   780
      Index           =   1
      Left            =   480
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   4140
      Width           =   7575
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "��������������ʾ:"
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   21
      Top             =   3900
      Width           =   2640
   End
   Begin VB.TextBox txt��ǹ��� 
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      Left            =   4395
      MaxLength       =   20
      TabIndex        =   20
      Top             =   3510
      Width           =   1860
   End
   Begin VB.ComboBox cbo��Ǽ� 
      Height          =   300
      Index           =   1
      Left            =   1245
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   3510
      Width           =   2025
   End
   Begin VB.CheckBox chk��ˮƽ 
      Caption         =   "���ˮƽ"
      Enabled         =   0   'False
      Height          =   240
      Left            =   6960
      TabIndex        =   7
      Top             =   690
      Width           =   1065
   End
   Begin VB.TextBox txt��ʾ 
      Enabled         =   0   'False
      Height          =   780
      Index           =   0
      Left            =   480
      MaxLength       =   500
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   2340
      Width           =   7575
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "��������������ʾ:"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   14
      Top             =   2100
      Width           =   2640
   End
   Begin VB.TextBox txt��ǹ��� 
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   4395
      MaxLength       =   20
      TabIndex        =   13
      Top             =   1695
      Width           =   1860
   End
   Begin VB.ComboBox cbo��Ǽ� 
      Height          =   300
      Index           =   0
      Left            =   1245
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1695
      Width           =   2025
   End
   Begin VB.ComboBox cbo����Χ 
      Height          =   300
      Left            =   6960
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1005
      Width           =   1050
   End
   Begin VB.ComboBox cbo���� 
      Height          =   300
      Left            =   915
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1005
      Width           =   2370
   End
   Begin VB.TextBox txt�ж� 
      Height          =   555
      Left            =   480
      MaxLength       =   80
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   375
      Width           =   6225
   End
   Begin VB.ComboBox cbo���� 
      Height          =   300
      Left            =   4395
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1005
      Width           =   2340
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "3.������(N)�жϹ���Ĵ���:"
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
      Left            =   135
      TabIndex        =   16
      Top             =   3225
      Width           =   2580
   End
   Begin VB.Label lbl��ǹ��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ǹ���                     (�ɲ��뵱ǰ����һ��)"
      Height          =   180
      Index           =   1
      Left            =   3615
      TabIndex        =   19
      Top             =   3570
      Width           =   4410
   End
   Begin VB.Label lbl��Ǽ� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ǵȼ�"
      Height          =   180
      Index           =   1
      Left            =   480
      TabIndex        =   17
      Top             =   3570
      Width           =   720
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2.����(Y)�жϹ���Ĵ���:"
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
      Left            =   135
      TabIndex        =   9
      Top             =   1425
      Width           =   2385
   End
   Begin VB.Label lbl������Ϣ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "1.�ж��������Ӧ����:"
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
      Left            =   135
      TabIndex        =   0
      Top             =   105
      Width           =   2070
   End
   Begin VB.Label lbl��ǹ��� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ǹ���                     (�ɲ��뵱ǰ����һ��)"
      Height          =   180
      Index           =   0
      Left            =   3615
      TabIndex        =   12
      Top             =   1755
      Width           =   4410
   End
   Begin VB.Label lbl��Ǽ� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ǵȼ�"
      Height          =   180
      Index           =   0
      Left            =   480
      TabIndex        =   10
      Top             =   1755
      Width           =   720
   End
   Begin VB.Label lbl����Χ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�걾��Χ:"
      Height          =   180
      Left            =   6960
      TabIndex        =   6
      Top             =   435
      Width           =   810
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      Height          =   180
      Left            =   480
      TabIndex        =   2
      Top             =   1065
      Width           =   360
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��Ӧ����"
      Height          =   180
      Left            =   3615
      TabIndex        =   4
      Top             =   1065
      Width           =   720
   End
End
Attribute VB_Name = "frmAppRuleEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngRuleId As Long          '��ǰ��ʾ�Ĺ���id
Private mlngParent As Long          '��ǰ������ϼ�id
Private mlngDevId As Long           '��ǰ��ʾ������id
Private mlngGroupID As Long         '��ǰ��ʾ�ķ���ID

Dim lngCount As Long

'--------------------------------------------
'����Ϊ���幫������
'--------------------------------------------
Public Function zlRefresh(lngRuleId As Long) As Boolean
    '���ܣ�������Ŀidˢ�µ�ǰ��ʾ����
    Dim rsTemp As New ADODB.Recordset
    mlngRuleId = lngRuleId: mlngParent = 0
    
    '�����ǰ��Ŀ����ʾ
    Me.txt�ж�.Text = "":
    Me.cbo����.Clear: Me.cbo����.ListIndex = -1
    Me.cbo����.Clear: Me.cbo����.ListIndex = -1
    Me.cbo����Χ.ListIndex = 0: Me.chk��ˮƽ.Value = vbUnchecked
    Me.cbo��Ǽ�(0).ListIndex = 0: Me.txt��ǹ���(0).Text = ""
    Me.chk����(0).Value = vbChecked: Me.txt��ʾ(0).Text = ""
    Me.cbo��Ǽ�(1).ListIndex = 0: Me.txt��ǹ���(1).Text = ""
    Me.chk����(1).Value = vbChecked: Me.txt��ʾ(1).Text = ""
    Me.chkUse = vbUnchecked
    If lngRuleId = 0 Then zlRefresh = True: Exit Function
    
    '��ȡָ����Ŀ����Ϣ
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select R.�ϼ�id, D.�ʿ�ˮƽ��, R.�ж�, R.����id, R.����, R.����Χ, R.��ˮƽ, R.Y��Ǽ�, R.Y����, R.Y����, R.Y��ʾ," & vbNewLine & _
            "       R.N��Ǽ�, R.N����, R.N����, R.N��ʾ, R.�Ƿ�ʹ�� " & vbNewLine & _
            "From ������������ R, �������� D" & vbNewLine & _
            "Where R.����id = D.ID And R.ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngRuleId)
    With rsTemp
        Me.chk��ˮƽ.Tag = 0
        If .RecordCount > 0 Then
            mlngParent = Val("" & !�ϼ�ID)
            Me.chk��ˮƽ.Tag = Val("" & !�ʿ�ˮƽ��)
            If Val("" & !�Ƿ�ʹ��) = 1 Then Me.chkUse = vbChecked
            
            Me.txt�ж�.Text = "" & !�ж�
            Select Case "" & !����
            Case "0": Me.cbo����.AddItem "0-��ʼ����": Me.cbo����.ListIndex = 0
            Case "Y": Me.cbo����.AddItem "Y-��һ������ʱִ�еĹ���": Me.cbo����.ListIndex = 0
            Case "N": Me.cbo����.AddItem "N-��һ������ʱִ�еĹ���": Me.cbo����.ListIndex = 0
            Case "1": Me.cbo����.AddItem "1-���ӹ���": Me.cbo����.ListIndex = 0
            End Select
            Me.cbo����.Tag = "" & !����id
            For lngCount = 0 To Me.cbo����Χ.ListCount - 1
                If lngCount = Val("" & !����Χ) - 1 Then Me.cbo����Χ.ListIndex = lngCount: Exit For
            Next
            If Val("" & !��ˮƽ) = 1 Then Me.chk��ˮƽ.Value = vbChecked
            
            For lngCount = 0 To Me.cbo��Ǽ�(0).ListCount - 1
                If lngCount = Val("" & !Y��Ǽ�) Then Me.cbo��Ǽ�(0).ListIndex = lngCount: Exit For
            Next
            Me.txt��ǹ���(0).Text = "" & !Y����
            Me.chk����(0).Value = IIf(Val("" & !Y����) = 0, vbUnchecked, vbChecked)
            Me.txt��ʾ(0).Text = "" & !Y��ʾ
            
            For lngCount = 0 To Me.cbo��Ǽ�(1).ListCount - 1
                If lngCount = Val("" & !n��Ǽ�) Then Me.cbo��Ǽ�(1).ListIndex = lngCount: Exit For
            Next
            Me.txt��ǹ���(1).Text = "" & !N����
            Me.chk����(1).Value = IIf(Val("" & !N����) = 0, vbUnchecked, vbChecked)
            Me.txt��ʾ(1).Text = "" & !N��ʾ
            
            If "" & !���� = "1" Then
                Me.chk����(0).Value = vbChecked: Me.chk����(0).Enabled = False
                Me.chk����(1).Value = vbChecked: Me.chk����(1).Enabled = False
            Else
                Me.chk����(0).Enabled = True
                Me.chk����(1).Enabled = True
            End If
        End If
    End With
    
    'Ŀǰֻ�����ù�����ɶಽ����򣬼�����ƽ��޺��ۻ��͹���ֻ����Ϊ���ӹ�����ֻ����ÿ��ˮƽ��>1ʱѡ�������ƽ��޹���
    If Left(Me.cbo����.Text, 1) = "1" Then
        If Val(Me.chk��ˮƽ.Tag) > 1 Then
            gstrSql = "Select ID, RPad(����, 200, ' ') || ��ˮƽ || ',' || N As ���� From �����ʿع��� Order By ����, ����"
        Else
            gstrSql = "Select ID, RPad(����, 200, ' ') || ��ˮƽ || ',' || N As ���� From �����ʿع��� Where ���� In (1, 3) Order By ����, ����"
        End If
    Else
        gstrSql = "Select ID, RPad(����, 200, ' ') || ��ˮƽ || ',' || N As ���� From �����ʿع��� Where ���� = 1 Order By ����, ����"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Me.cbo����.Clear
        Do While Not .EOF
            Me.cbo����.AddItem "" & !����
            Me.cbo����.ItemData(Me.cbo����.NewIndex) = Val("" & !ID)
            If Val("" & !ID) = Val(Me.cbo����.Tag) Then Me.cbo����.ListIndex = Me.cbo����.NewIndex
            .MoveNext
        Loop
    End With
    
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditStart(blnAdd As Boolean, lngRuleId As Long, lngDevId As Long, lngGroupID As Long, Optional blnSingle As Boolean) As Boolean
    '���ܣ���ʼ��Ŀ�༭
    '������ blnAdd-�Ƿ����ӣ�����Ϊ�޸�
    '       lngRuleId-����ʱ,Ϊ��ǰ�����ϼ�����û���ϼ�ʱΪ0���޸�ʱΪ��ǰ�����ID
    '       lngDevId-��ǰ������Ʒ�������豸id
    '       blnSingle-������ʱ��Ч��ָ�������Ӷ������ǵ�����
    Dim rsTemp As New ADODB.Recordset
    Dim strKind As String
    
    mlngDevId = lngDevId
    mlngGroupID = lngGroupID
    Err = 0: On Error GoTo ErrHand
    
    '�豸ˮƽ��
    If blnAdd Then
        Me.chk��ˮƽ.Tag = 0
        gstrSql = "Select �ʿ�ˮƽ�� From �������� Where ID = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngDevId)
        With rsTemp
            If .RecordCount > 0 Then Me.chk��ˮƽ.Tag = Val("" & !�ʿ�ˮƽ��)
        End With
    
        Me.cbo����.Clear
        If blnSingle Then
            Me.cbo����.AddItem "1-���ӹ���"
        Else
            If lngRuleId = 0 Then
                gstrSql = "Select Decode(Nvl(Count(*), 0), 0, 1, 0) As ��� From ������������ Where ����id = [1] And ��Ŀid =[2] And ���� = '0'"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngDevId, lngGroupID)
                With rsTemp
                    If .RecordCount > 0 Then
                        If !��� Then Me.cbo����.AddItem "0-��ʼ����"
                    End If
                End With
            Else
                gstrSql = "Select Decode(P.Y����, 1, 0, Decode(Y����, 0, 1, 0)) As Y��, Decode(P.N����, 1, 0, Decode(N����, 0, 1, 0)) As N��" & vbNewLine & _
                        "From (Select Nvl(Y����, 0) As Y����, Nvl(N����, 0) As N���� From ������������ Where ID = [1]) P," & vbNewLine & _
                        "     (Select Nvl(Sum(Decode(����, 'Y', 1, 0)), 0) As Y����, Nvl(Sum(Decode(����, 'N', 1, 0)), 0) As N����" & vbNewLine & _
                        "       From ������������" & vbNewLine & _
                        "       Where �ϼ�id = [1]) C"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngRuleId)
                With rsTemp
                    If .RecordCount > 0 Then
                        If !Y�� Then Me.cbo����.AddItem "Y-��һ������ʱִ�еĹ���"
                        If !N�� Then Me.cbo����.AddItem "N-��һ������ʱִ�еĹ���"
                    End If
                End With
            End If
            If Me.cbo����.ListCount = 0 Then
                MsgBox "�Ѿ�������Ӧ�Ĺ������һ���Ѿ�������", vbInformation, gstrSysName
                zlEditStart = False: Exit Function
            End If
        End If
        Me.cbo����.ListIndex = 0
        
        'Ŀǰֻ�����ù�����ɶಽ����򣬼�����ƽ��޺��ۻ��͹���ֻ����Ϊ���ӹ�����ֻ����ÿ��ˮƽ��>1ʱѡ�������ƽ��޹���
        If Left(Me.cbo����.Text, 1) = "1" Then
'            If Val(Me.chk��ˮƽ.Tag) > 1 Then
'                gstrSql = "Select ID, RPad(����, 200, ' ') || ��ˮƽ || ',' || N As ���� From �����ʿع��� Where ���� In (2, 3) Order By ����, ����"
'            Else
'                gstrSql = "Select ID, RPad(����, 200, ' ') || ��ˮƽ || ',' || N As ���� From �����ʿع��� Where ���� = 3 Order By ����, ����"
'            End If

            If Val(Me.chk��ˮƽ.Tag) > 1 Then
                'ÿ��ˮƽ>1ʱ������ѡ�����й�����Ϊ���ӹ���
                gstrSql = "Select ID, RPad(����, 200, ' ') || ��ˮƽ || ',' || N As ���� From �����ʿع��� Order By ����, ����"
            Else
                'ÿ��ˮƽ=1ʱ������ѡ�����ʿع�����ۻ��͹�����Ϊ���ӹ���
                gstrSql = "Select ID, RPad(����, 200, ' ') || ��ˮƽ || ',' || N As ���� From �����ʿع��� Where ���� In (1, 3) Order By ����, ����"
            End If
        Else
            gstrSql = "Select ID, RPad(����, 200, ' ') || ��ˮƽ || ',' || N As ���� From �����ʿع��� Where ���� = 1 Order By ����, ����"
        End If
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
        With rsTemp
            Me.cbo����.Clear
            Do While Not .EOF
                Me.cbo����.AddItem "" & !����
                Me.cbo����.ItemData(Me.cbo����.NewIndex) = Val("" & !ID)
                .MoveNext
            Loop
        End With
        If Me.cbo����.ListCount = 0 Then
            MsgBox "�����ȳ�ʼ�������ʿع���", vbInformation, gstrSysName
            zlEditStart = False: Exit Function
        Else
            Me.cbo����.ListIndex = 0
        End If
        mlngParent = lngRuleId
    
        Me.txt�ж�.Text = ""
        Me.cbo����Χ.ListIndex = 0
        Me.cbo��Ǽ�(0).ListIndex = 0: Me.txt��ǹ���(0).Text = ""
        Me.chk����(0).Value = vbChecked: Me.txt��ʾ(0).Text = ""
        Me.cbo��Ǽ�(1).ListIndex = 0: Me.txt��ǹ���(1).Text = ""
        Me.chk����(1).Value = vbChecked: Me.txt��ʾ(1).Text = ""
        If blnSingle Then
            Me.chk����(0).Enabled = False
            Me.chk����(1).Enabled = False
        Else
            Me.chk����(0).Enabled = True
            Me.chk����(1).Enabled = True
        End If
        Me.chkUse.Value = vbChecked
    Else
        strKind = Left(Me.cbo����.Text, 1)
        Me.cbo����.Clear
        Select Case strKind
        Case "0": Me.cbo����.AddItem "0-��ʼ����": Me.cbo����.ListIndex = 0
        Case "Y", "N"
            gstrSql = "Select Decode(P.Y����, 1, 0, Decode(Y����, 0, 1, 0)) As Y��, Decode(P.N����, 1, 0, Decode(N����, 0, 1, 0)) As N��" & vbNewLine & _
                    "From (Select Nvl(Y����, 0) As Y����, Nvl(N����, 0) As N���� From ������������ Where ID = [1]) P," & vbNewLine & _
                    "     (Select Nvl(Sum(Decode(����, 'Y', 1, 0)), 0) As Y����, Nvl(Sum(Decode(����, 'N', 1, 0)), 0) As N����" & vbNewLine & _
                    "       From ������������" & vbNewLine & _
                    "       Where �ϼ�id = [1]) C"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngParent)
            With rsTemp
                If .RecordCount > 0 Then
                    'ʵ��ֻ������һ����ɴ���
                    If !Y�� Then Me.cbo����.AddItem "Y-��һ������ʱִ�еĹ���"
                    If !N�� Then Me.cbo����.AddItem "N-��һ������ʱִ�еĹ���"
                End If
            End With
            If strKind = "Y" Then
                Me.cbo����.AddItem "Y-��һ������ʱִ�еĹ���": Me.cbo����.ListIndex = Me.cbo����.NewIndex
            Else
                Me.cbo����.AddItem "N-��һ������ʱִ�еĹ���": Me.cbo����.ListIndex = Me.cbo����.NewIndex
            End If
        Case "1": Me.cbo����.AddItem "1-���ӹ���": Me.cbo����.ListIndex = 0
        End Select
    End If
    
    Me.Enabled = True: Me.Tag = IIf(blnAdd, "����", "�޸�")
    Me.BackColor = RGB(250, 250, 250)
    Me.chk��ˮƽ.BackColor = Me.BackColor
    Me.chk����(0).BackColor = Me.BackColor
    Me.chk����(1).BackColor = Me.BackColor
    
    Me.txt�ж�.SetFocus
    zlEditStart = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditStart = False: Exit Function
End Function

Public Sub zlEditCancel()
    '���ܣ��������ڽ��еı༭
    Me.Enabled = False: Me.Tag = ""
    Me.BackColor = &H8000000F
    Me.chk��ˮƽ.BackColor = Me.BackColor
    Me.chk����(0).BackColor = Me.BackColor
    Me.chk����(1).BackColor = Me.BackColor
    
    Call Me.zlRefresh(mlngRuleId)
End Sub

Public Function zlEditSave() As Long
    '���ܣ��������ڽ��еı༭,���������ڱ༭��Ŀid,����ʧ�ܷ���0
    Dim lngNewId As Long, blnMatch As Boolean

    'һ�����Լ��
    If Trim(Me.txt�ж�.Text) = "" Then
        MsgBox "�������жϣ�", vbInformation, gstrSysName
        Me.txt�ж�.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt�ж�.Text), vbFromUnicode)) > Me.txt�ж�.MaxLength Then
        MsgBox "�жϳ��������" & Me.txt�ж�.MaxLength & "���ַ���ȳ����֣���", vbInformation, gstrSysName
        Me.txt�ж�.SetFocus: zlEditSave = 0: Exit Function
    End If
    Me.txt�ж�.Text = Replace(Me.txt�ж�.Text, vbCrLf, "")
    Me.txt�ж�.Text = Replace(Me.txt�ж�.Text, vbCr, "")
    Me.txt�ж�.Text = Replace(Me.txt�ж�.Text, vbLf, "")
    
    If Me.cbo����.ListIndex = -1 Then
        MsgBox "��ָ�����ʣ�", vbInformation, gstrSysName
        Me.txt�ж�.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Me.cbo����.ListIndex = -1 Then
        MsgBox "��ָ������", vbInformation, gstrSysName
        Me.cbo����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Left(Me.cbo����.Text, 1) <> "1" Then
        blnMatch = False
        If Me.chk��ˮƽ.Value = vbChecked Then
            If Val(Me.cbo����Χ.Tag) > Me.cbo����Χ.ListIndex * Val(Me.chk��ˮƽ.Tag) And Val(Me.cbo����Χ.Tag) <= (Me.cbo����Χ.ListIndex + 1) * Val(Me.chk��ˮƽ.Tag) Then blnMatch = True
        Else
            If Val(Me.cbo����Χ.Tag) = (Me.cbo����Χ.ListIndex + 1) Then blnMatch = True
        End If
        If blnMatch = False Then
            MsgBox "����Χ��Ҫ�͹���Ҫ��ı걾��Χƥ�䣡", vbInformation, gstrSysName
            Call chk��ˮƽ_Click
            Me.cbo����Χ.SetFocus: zlEditSave = 0: Exit Function
        End If
    Else
        If Val(Me.cbo����Χ.Tag) > (Me.cbo����Χ.ListIndex + 1) * IIf(Me.chk��ˮƽ.Value = vbChecked, Val(Me.chk��ˮƽ.Tag), 1) Then
            MsgBox "����Χ��Ҫ�͹���Ҫ��ı걾��Χƥ�䣡", vbInformation, gstrSysName
            Me.cbo����Χ.SetFocus: zlEditSave = 0: Exit Function
        End If
    End If
    
    For lngCount = 0 To 1
        If Me.txt��ǹ���(lngCount).Enabled Then
            If Trim(Me.txt��ǹ���(lngCount).Text = "") Then
                MsgBox "����Ǿ����ʧ��ʱ����Ҫָ����ǹ���", vbInformation, gstrSysName
                Me.txt��ǹ���(lngCount).SetFocus: zlEditSave = 0: Exit Function
            End If
            If LenB(StrConv(Trim(Me.txt��ǹ���(lngCount).Text), vbFromUnicode)) > Me.txt��ǹ���(lngCount).MaxLength Then
                MsgBox "��ǹ��򳬳������" & Me.txt��ǹ���(lngCount).MaxLength & "���ַ���ȳ����֣���", vbInformation, gstrSysName
                Me.txt��ǹ���(lngCount).SetFocus: zlEditSave = 0: Exit Function
            End If
        End If
        If Me.txt��ʾ(lngCount).Enabled Then
            If Trim(Me.txt��ʾ(lngCount).Text = "") Then
                MsgBox "������ʱ����Ҫ��д��ʾ���ݣ�", vbInformation, gstrSysName
                Me.txt��ʾ(lngCount).SetFocus: zlEditSave = 0: Exit Function
            End If
            If LenB(StrConv(Trim(Me.txt��ʾ(lngCount).Text), vbFromUnicode)) > Me.txt��ʾ(lngCount).MaxLength Then
                MsgBox "��ʾ���������" & Me.txt��ʾ(lngCount).MaxLength & "���ַ���ȳ����֣���", vbInformation, gstrSysName
                Me.txt��ʾ(lngCount).SetFocus: zlEditSave = 0: Exit Function
            End If
        End If
        Me.txt��ʾ(lngCount).Text = Replace(Me.txt��ʾ(lngCount).Text, vbCrLf, "")
        Me.txt��ʾ(lngCount).Text = Replace(Me.txt��ʾ(lngCount).Text, vbCr, "")
        Me.txt��ʾ(lngCount).Text = Replace(Me.txt��ʾ(lngCount).Text, vbLf, "")
    Next
    
    '���ݱ��������֯
    gstrSql = "'" & Trim(Me.txt�ж�.Text) & "','" & Left(Me.cbo����.Text, 1) & "'," & Me.cbo����.ItemData(Me.cbo����.ListIndex)
    gstrSql = gstrSql & "," & Me.cbo����Χ.ListIndex + 1 & "," & IIf(Me.chk��ˮƽ.Value = vbChecked, 1, 0)
    For lngCount = 0 To 1
        gstrSql = gstrSql & "," & Me.cbo��Ǽ�(lngCount).ListIndex
        If Me.cbo��Ǽ�(lngCount).ListIndex > 0 Then
            gstrSql = gstrSql & ",'" & Trim(Me.txt��ǹ���(lngCount).Text) & "'"
        Else
            gstrSql = gstrSql & ",''"
        End If
        If Me.chk����(lngCount).Value = vbChecked Then
            gstrSql = gstrSql & ",1,'" & Trim(Me.txt��ʾ(lngCount).Text) & "'"
        Else
            gstrSql = gstrSql & ",0,''"
        End If
        
    Next
    
    gstrSql = gstrSql & "," & IIf(Me.chkUse.Value = vbChecked, 1, 0)
    
    If Me.Tag = "����" Then
        lngNewId = zlDatabase.GetNextId("������������")
        gstrSql = "Zl_������������_Edit(1," & lngNewId & "," & mlngParent & "," & mlngDevId & "," & mlngGroupID & "," & gstrSql & ")"
    Else
        gstrSql = "Zl_������������_Edit(2," & mlngRuleId & "," & mlngParent & "," & mlngDevId & "," & mlngGroupID & "," & gstrSql & ")"
    End If
    
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    If Me.Tag = "����" Then mlngRuleId = lngNewId
    Me.Enabled = False: Me.Tag = ""
    Me.BackColor = &H8000000F
    Me.chk��ˮƽ.BackColor = Me.BackColor
    Me.chk����(0).BackColor = Me.BackColor
    Me.chk����(1).BackColor = Me.BackColor
    
    zlEditSave = mlngRuleId: Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function

'--------------------------------------------
'����Ϊ����ؼ���Ӧ�¼�
'--------------------------------------------
Private Sub cbo��Ǽ�_Click(Index As Integer)
    If Me.cbo��Ǽ�(Index).ListIndex <= 0 Then
        Me.txt��ǹ���(Index).Enabled = False
    Else
        Me.txt��ǹ���(Index).Enabled = True
    End If
End Sub

Private Sub cbo��Ǽ�_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo��Ǽ�_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cbo����_Click()
    Dim strRule As String
    
    '�༭״̬���Զ�����Υ���Ĺ���
    If Me.Tag <> "" Then
        strRule = Trim(Left(Me.cbo����.Text, 100))
        Me.txt��ǹ���(0).Text = strRule
        Me.txt��ǹ���(1).Text = strRule
    End If
    
    '��¼��ǰ����Ҫ��ļ�����,�ж϶�ˮƽ����
    Me.cbo����Χ.Tag = Split(Trim(Mid(Me.cbo����.Text, 200)), ",")(1)
    If Val(Split(Trim(Mid(Me.cbo����.Text, 200)), ",")(0)) = 0 Or Val(Me.chk��ˮƽ.Tag) <= 1 Then
        Me.chk��ˮƽ.Value = vbUnchecked: Me.chk��ˮƽ.Enabled = False
    Else
        Me.chk��ˮƽ.Enabled = True
    End If
    
    Call chk��ˮƽ_Click
End Sub

Private Sub cbo����_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cbo����Χ_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo����Χ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cbo����_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub chk��ˮƽ_Click()
    Dim lngBatch As Long
    If Val(Me.cbo����Χ.Tag) <> 0 Then
        If Me.chk��ˮƽ.Value = vbUnchecked Then
            lngBatch = Val(Me.cbo����Χ.Tag) - 1
        Else
            lngBatch = Int(Val(Me.cbo����Χ.Tag) / Val(Me.chk��ˮƽ.Tag) + 0.9) - 1
        End If
        If lngBatch < 0 Then lngBatch = 0
        Me.cbo����Χ.ListIndex = lngBatch
    End If
End Sub

Private Sub chk��ˮƽ_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub chk��ˮƽ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub chk����_Click(Index As Integer)
    If Me.chk����(Index).Value = vbChecked Then
        Me.txt��ʾ(Index).Enabled = True
    Else
        Me.txt��ʾ(Index).Enabled = False
    End If
End Sub

Private Sub chk����_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub chk����_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub Form_Load()
    mlngRuleId = 0: mlngDevId = 0: mlngGroupID = 0
        
    Me.cbo����Χ.AddItem "��ǰ��"
    For lngCount = 2 To 31
        Me.cbo����Χ.AddItem "��" & lngCount & "��"
    Next
    
    Me.cbo��Ǽ�(0).AddItem "0-�����"
    Me.cbo��Ǽ�(0).AddItem "1-���Ϊ����"
    Me.cbo��Ǽ�(0).AddItem "2-���Ϊʧ��"
    
    Me.cbo��Ǽ�(1).AddItem "0-�����"
    Me.cbo��Ǽ�(1).AddItem "1-���Ϊ����"
    Me.cbo��Ǽ�(1).AddItem "2-���Ϊʧ��"
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    Me.cbo����Χ.Left = Me.ScaleWidth - Me.cbo����Χ.Width - 150
    Me.lbl����Χ.Left = Me.cbo����Χ.Left
    Me.chk��ˮƽ.Left = Me.cbo����Χ.Left
    Me.txt�ж�.Width = Me.cbo����Χ.Left - Me.txt�ж�.Left - 300
    Me.cbo����.Width = Me.txt�ж�.Left + Me.txt�ж�.Width - Me.cbo����.Left
    
    Me.txt��ʾ(0).Width = Me.ScaleWidth - Me.txt��ʾ(0).Left - 150
    Me.txt��ʾ(1).Width = Me.ScaleWidth - Me.txt��ʾ(1).Left - 150
    
    Me.chkUse.Left = Me.chk��ˮƽ.Left
End Sub

Private Sub txt��ǹ���_GotFocus(Index As Integer)
    Me.txt��ǹ���(Index).SelStart = 0: Me.txt��ǹ���(Index).SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt��ǹ���_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt�ж�_GotFocus()
    Me.txt�ж�.SelStart = 0: Me.txt�ж�.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt�ж�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt��ʾ_GotFocus(Index As Integer)
    Me.txt��ʾ(Index).SelStart = 0: Me.txt��ʾ(Index).SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt��ʾ_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
