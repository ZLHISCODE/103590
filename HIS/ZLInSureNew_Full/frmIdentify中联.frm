VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIdentify���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ���������ʶ��"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   Icon            =   "frmIdentify����.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3930
      TabIndex        =   39
      Top             =   5070
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5340
      TabIndex        =   40
      Top             =   5070
      Width           =   1100
   End
   Begin VB.Frame fra���� 
      Caption         =   "�����ʻ����"
      Height          =   1305
      Index           =   1
      Left            =   150
      TabIndex        =   30
      Top             =   3570
      Width           =   6795
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   10
         Left            =   5130
         MaxLength       =   14
         TabIndex        =   38
         Top             =   720
         Width           =   1395
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   9
         Left            =   1590
         MaxLength       =   14
         TabIndex        =   36
         Top             =   720
         Width           =   1395
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   8
         Left            =   5130
         MaxLength       =   14
         TabIndex        =   34
         Top             =   330
         Width           =   1395
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   7
         Left            =   1590
         MaxLength       =   14
         TabIndex        =   32
         Top             =   330
         Width           =   1395
      End
      Begin VB.Label lbl��ʾ 
         AutoSize        =   -1  'True
         Caption         =   "����ͳ���ۼ�(&G)"
         Height          =   180
         Index           =   9
         Left            =   180
         TabIndex        =   35
         Top             =   780
         Width           =   1350
      End
      Begin VB.Label lbl��ʾ 
         AutoSize        =   -1  'True
         Caption         =   "�ʻ�֧���ۼ�(&W)"
         Height          =   180
         Index           =   8
         Left            =   3690
         TabIndex        =   33
         Top             =   390
         Width           =   1350
      End
      Begin VB.Label lbl��ʾ 
         AutoSize        =   -1  'True
         Caption         =   "�ʻ������ۼ�(&A)"
         Height          =   180
         Index           =   7
         Left            =   180
         TabIndex        =   31
         Top             =   390
         Width           =   1350
      End
      Begin VB.Label lbl��ʾ 
         AutoSize        =   -1  'True
         Caption         =   "ͳ�ﱨ���ۼ�(&P)"
         Height          =   180
         Index           =   10
         Left            =   3690
         TabIndex        =   37
         Top             =   780
         Width           =   1350
      End
   End
   Begin VB.Frame fra���� 
      Caption         =   "���˻�����Ϣ"
      Height          =   3195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   6795
      Begin VB.ComboBox Cbo��ǰ״̬ 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2280
         Width           =   1455
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   12
         Left            =   4440
         MaxLength       =   18
         TabIndex        =   17
         Top             =   1515
         Width           =   2085
      End
      Begin VB.ComboBox cmb�Ա� 
         Height          =   300
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1125
         Width           =   2085
      End
      Begin VB.ComboBox cmb���� 
         Height          =   300
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   330
         Width           =   2085
      End
      Begin VB.TextBox txtEdit 
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   4440
         MaxLength       =   26
         TabIndex        =   26
         Top             =   2310
         Width           =   2085
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "��"
         Height          =   240
         Index           =   2
         Left            =   6240
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   1935
         Width           =   255
      End
      Begin MSComCtl2.DTPicker dtp���� 
         Height          =   300
         Left            =   1320
         TabIndex        =   15
         Top             =   1515
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   87031811
         CurrentDate     =   36526
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   11
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   11
         Top             =   1125
         Width           =   1455
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "��"
         Height          =   240
         Index           =   1
         Left            =   6240
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2730
         Width           =   255
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "��"
         Height          =   240
         Index           =   0
         Left            =   2490
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   6
         Left            =   4440
         MaxLength       =   2
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   1320
         MaxLength       =   8
         TabIndex        =   19
         Top             =   1905
         Width           =   1455
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   7
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   2
         Top             =   330
         Width           =   1455
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   4
         Left            =   4440
         MaxLength       =   8
         TabIndex        =   21
         Top             =   1905
         Width           =   2085
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   5
         Left            =   4440
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   2700
         Width           =   2085
      End
      Begin VB.Label lbl��ʾ 
         AutoSize        =   -1  'True
         Caption         =   "��ǰ״̬(&K)"
         Height          =   180
         Index           =   16
         Left            =   240
         TabIndex        =   23
         Top             =   2340
         Width           =   990
      End
      Begin VB.Label lbl��ʾ 
         AutoSize        =   -1  'True
         Caption         =   "��������(&B)"
         Height          =   180
         Index           =   15
         Left            =   240
         TabIndex        =   14
         Top             =   1575
         Width           =   990
      End
      Begin VB.Label lbl��ʾ 
         AutoSize        =   -1  'True
         Caption         =   "�Ա�(&X)"
         Height          =   180
         Index           =   14
         Left            =   3720
         TabIndex        =   12
         Top             =   1185
         Width           =   630
      End
      Begin VB.Label lbl��ʾ 
         AutoSize        =   -1  'True
         Caption         =   "ҽ������(&R)"
         Height          =   180
         Index           =   13
         Left            =   3360
         TabIndex        =   4
         Top             =   390
         Width           =   990
      End
      Begin VB.Label lbl��ʾ 
         AutoSize        =   -1  'True
         Caption         =   "���֤��(&I)"
         Height          =   180
         Index           =   12
         Left            =   3360
         TabIndex        =   16
         Top             =   1575
         Width           =   990
      End
      Begin VB.Label lbl��ʾ 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   11
         Left            =   600
         TabIndex        =   10
         Top             =   1185
         Width           =   630
      End
      Begin VB.Label lbl��ʾ 
         AutoSize        =   -1  'True
         Caption         =   "סԺ����(&S)"
         Height          =   180
         Index           =   6
         Left            =   3360
         TabIndex        =   8
         Top             =   780
         Width           =   990
      End
      Begin VB.Label lbl��ʾ 
         AutoSize        =   -1  'True
         Caption         =   "����(&F)"
         Height          =   180
         Index           =   5
         Left            =   3720
         TabIndex        =   27
         Top             =   2760
         Width           =   630
      End
      Begin VB.Label lbl��ʾ 
         AutoSize        =   -1  'True
         Caption         =   "��λ����(&U)"
         Height          =   180
         Index           =   4
         Left            =   3360
         TabIndex        =   20
         Top             =   1965
         Width           =   990
      End
      Begin VB.Label lbl��ʾ 
         AutoSize        =   -1  'True
         Caption         =   "��Ա���(&E)"
         Height          =   180
         Index           =   3
         Left            =   240
         TabIndex        =   18
         Top             =   1965
         Width           =   990
      End
      Begin VB.Label lbl��ʾ 
         AutoSize        =   -1  'True
         Caption         =   "����֤��(&Z)"
         Enabled         =   0   'False
         Height          =   180
         Index           =   2
         Left            =   3360
         TabIndex        =   25
         Top             =   2370
         Width           =   990
      End
      Begin VB.Label lbl��ʾ 
         AutoSize        =   -1  'True
         Caption         =   "ҽ����(&Y)"
         Height          =   180
         Index           =   1
         Left            =   420
         TabIndex        =   6
         Top             =   780
         Width           =   810
      End
      Begin VB.Label lbl��ʾ 
         AutoSize        =   -1  'True
         Caption         =   "����(&D)"
         Height          =   180
         Index           =   0
         Left            =   600
         TabIndex        =   1
         Top             =   390
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmIdentify����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum �ı�Enum
    Text���� = 0
    Textҽ���� = 1
    Text����֤�� = 2
    Text��Ա��� = 3
    Text���˵�λ = 4
    Text���� = 5
    TextסԺ���� = 6
    Text�ʻ������ۼ� = 7
    Text�ʻ�֧���ۼ� = 8
    Text����ͳ���ۼ� = 9
    Textͳ�ﱨ���ۼ� = 10
    Text���� = 11
    Text���֤�� = 12
End Enum

Private Enum ѡ��Enum
    Select���� = 0
    Select���� = 1
    Select��λ = 2
End Enum

Dim mstrIdentify As String
Dim mbytType As Byte
Dim mlng����ID As Long
Dim mint���� As Integer

Public Function ShowCard(Optional bytType As Byte, Optional lng����ID As Long, Optional ByVal int���� As Integer) As String
'���ܣ�����ҽ�����˵������Ϣ
'������bytType-ʶ�����ͣ�0-���1-סԺ
'���أ�
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�,23��������
    Dim rsTemp As New ADODB.Recordset
    mbytType = bytType
    mlng����ID = lng����ID
    mint���� = int����
    mstrIdentify = ""
    
    cmb�Ա�.Clear
    gstrSQL = "select ����,���� from �Ա� order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    Do Until rsTemp.EOF
        cmb�Ա�.AddItem rsTemp("����") & "." & rsTemp("����")
        rsTemp.MoveNext
    Loop
    
    cmb����.Clear
    gstrSQL = "select A.��������,B.���,B.����,B.���� from ������� A,��������Ŀ¼ B where A.���=[1] and A.���=b.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint����)
    
    If rsTemp("��������") = 0 Then
        lbl��ʾ(13).Visible = False
        cmb����.Visible = False
        cmb����.AddItem "1.����" '������
    End If
    Do Until rsTemp.EOF
        cmb����.AddItem rsTemp("����") & "." & rsTemp("����")
        cmb����.ItemData(cmb����.NewIndex) = rsTemp("���")
        rsTemp.MoveNext
    Loop
    cmb����.ListIndex = 0
    
    '1-��ְ;2-����;3-����
    Cbo��ǰ״̬.Clear
    Cbo��ǰ״̬.AddItem "��ְ"
    Cbo��ǰ״̬.ItemData(Cbo��ǰ״̬.NewIndex) = 1
    Cbo��ǰ״̬.AddItem "����"
    Cbo��ǰ״̬.ItemData(Cbo��ǰ״̬.NewIndex) = 2
    Cbo��ǰ״̬.AddItem "����"
    Cbo��ǰ״̬.ItemData(Cbo��ǰ״̬.NewIndex) = 3
    Cbo��ǰ״̬.ListIndex = 0
        
    dtp����.MaxDate = zlDatabase.Currentdate
    frmIdentify����.Show vbModal
    ShowCard = mstrIdentify
End Function

Private Sub Cbo��ǰ״̬_Click()
    TxtEdit(Text����֤��).Enabled = (Cbo��ǰ״̬.ListIndex <> 0)
    lbl��ʾ(Text����֤��).Enabled = (Cbo��ǰ״̬.ListIndex <> 0)
End Sub

Private Sub Cbo��ǰ״̬_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub cmb����_Click()
    Dim lng���ų��� As Long, lng����֤���� As Long
    Dim rsTemp As New ADODB.Recordset
    
    'ȱʡֵ
    lng���ų��� = 20
    lng����֤���� = 26
    
    gstrSQL = "select ������,����ֵ from ���ղ��� where ����=[1] and (���� is null or ����=[2])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint����, CInt(cmb����.ItemData(cmb����.ListIndex)))
    
    Do Until rsTemp.EOF
        Select Case rsTemp("������")
            Case "���ų���"
                If IsNull(rsTemp("����ֵ")) = False Then lng���ų��� = Val(rsTemp("����ֵ"))
            Case "����֤����"
                If IsNull(rsTemp("����ֵ")) = False Then lng����֤���� = Val(rsTemp("����ֵ"))
        End Select
        rsTemp.MoveNext
    Loop
    
    TxtEdit(Text����).MaxLength = lng���ų���
    TxtEdit(Text����֤��).MaxLength = lng����֤����
End Sub

Private Sub cmdCancel_Click()
    mstrIdentify = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim strIdentify As String, strAddition As String
    Dim lng����ID As Long, lng���� As Long
    
    '���������ݵ���ȷ��
    If IsValid() = False Then
        Exit Sub
    End If
    
    '�õ��������
    If cmb����.Visible = False Then
        lng���� = 0
    Else
        If cmb����.ListIndex < 0 Then
            MsgBox "��ѡ��������ҽ�����ġ�", vbInformation, gstrSysName
            cmb����.SetFocus
            Exit Sub
        End If
        lng���� = cmb����.ItemData(cmb����.ListIndex)
    End If
    
    '��鲡��״̬
    gstrSQL = "select nvl(��ǰ״̬,0) as ״̬ from �����ʻ� where ����=[1] and ����=[2] and ҽ����=[3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint����, lng����, CStr(Trim(TxtEdit(Textҽ����).Text)))
    
    If rsTemp.RecordCount > 0 Then
        If rsTemp("״̬") > 0 Then
            MsgBox "�ò����Ѿ���Ժ������ͨ�������֤��", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '�����ַ���
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�,23��������
    strIdentify = Trim(TxtEdit(Text����).Text)                         '0����
    strIdentify = strIdentify & ";" & Trim(TxtEdit(Textҽ����).Text)   '1ҽ����
    strIdentify = strIdentify & ";"                                    '2����
    strIdentify = strIdentify & ";" & Trim(TxtEdit(Text����).Text)     '3����
    strIdentify = strIdentify & ";" & Replace(GetTextFromCombo(cmb�Ա�, True), "'", "") '4�Ա�
    strIdentify = strIdentify & ";" & Format(dtp����.Value, "yyyy-MM-dd") '5��������
    strIdentify = strIdentify & ";" & Trim(TxtEdit(Text���֤��).Text)    '6���֤
    strIdentify = strIdentify & ";" & Trim(TxtEdit(Text���˵�λ).Text) & "(" & Trim(TxtEdit(Text���˵�λ).Text) & ")"  '7.��λ����(����)
    strAddition = ";" & lng����                                 '8.���Ĵ���
    strAddition = strAddition & ";"                             '9.˳���
    strAddition = strAddition & ";" & Trim(TxtEdit(Text��Ա���).Text)       '10��Ա���
    strAddition = strAddition & ";" & Val(TxtEdit(Text�ʻ������ۼ�).Text) - Val(TxtEdit(Text�ʻ�֧���ۼ�).Text)  '11�ʻ����
    strAddition = strAddition & ";0"                            '12��ǰ״̬
    strAddition = strAddition & ";" & TxtEdit(Text����).Tag     '13����ID
    strAddition = strAddition & ";" & Cbo��ǰ״̬.ItemData(Cbo��ǰ״̬.ListIndex) '14��ְ(1,2,3)
    strAddition = strAddition & ";" & Trim(TxtEdit(Text����֤��).Text) '15����֤��
    strAddition = strAddition & ";" & DateDiff("yyyy", dtp����.Value, dtp����.MaxDate) '16�����
    strAddition = strAddition & ";"                             '17�Ҷȼ�
    strAddition = strAddition & ";" & Val(TxtEdit(Text�ʻ������ۼ�).Text)       '18�ʻ������ۼ�
    strAddition = strAddition & ";" & Val(TxtEdit(Text�ʻ�֧���ۼ�).Text)       '19�ʻ�֧���ۼ�
    strAddition = strAddition & ";" & Val(TxtEdit(Text����ͳ���ۼ�).Text)       '20����ͳ���ۼ�
    strAddition = strAddition & ";" & Val(TxtEdit(Textͳ�ﱨ���ۼ�).Text)       '21ͳ�ﱨ���ۼ�
    strAddition = strAddition & ";" & Int(Val(TxtEdit(TextסԺ����).Text))      '22סԺ�����ۼ�
    strAddition = strAddition & ";"                                             '23�������� (1����������)
    
    lng����ID = BuildPatiInfo(mbytType, strIdentify & strAddition, mlng����ID, mint����)
    '���ظ�ʽ:�м���벡��ID
    If lng����ID > 0 Then
        mstrIdentify = strIdentify & ";" & lng����ID & strAddition
    End If
    Unload Me
End Sub

Private Function IsValid() As Boolean
'���ܣ�������ݵ���ȷ��
    Dim lngIndex As Long
    
    For lngIndex = TxtEdit.LBound To TxtEdit.UBound
        If TxtEdit(lngIndex).Enabled = True Then
            If zlCommFun.StrIsValid(TxtEdit(lngIndex), TxtEdit(lngIndex).MaxLength) = False Then
                zlControl.TxtSelAll TxtEdit(lngIndex)
                TxtEdit(lngIndex).SetFocus
                Exit Function
            End If
        End If
        
        If lngIndex >= TextסԺ���� And lngIndex <= Textͳ�ﱨ���ۼ� Then
            If IsNumeric(TxtEdit(lngIndex).Text) = False Then
                MsgBox "������Ϸ�����ֵ��", vbInformation, gstrSysName
                zlControl.TxtSelAll TxtEdit(lngIndex)
                TxtEdit(lngIndex).SetFocus
                Exit Function
            End If
            
            
            If lngIndex = TextסԺ���� Then
                If Val(TxtEdit(lngIndex).Text) < 0 Or Val(TxtEdit(lngIndex).Text) > 100 Then
                    MsgBox "סԺ��������С��0���Ҳ��ܳ���100��", vbInformation, gstrSysName
                    zlControl.TxtSelAll TxtEdit(TextסԺ����)
                    TxtEdit(TextסԺ����).SetFocus
                    Exit Function
                End If
            Else
                If Val(TxtEdit(lngIndex).Text) < 0 Or Val(TxtEdit(lngIndex).Text) > 1000000 Then
                    MsgBox "����С��0���Ҳ��ܳ���100��", vbInformation, gstrSysName
                    zlControl.TxtSelAll TxtEdit(lngIndex)
                    TxtEdit(lngIndex).SetFocus
                    Exit Function
                End If
            End If
        End If
        If (lngIndex = Text���� Or lngIndex = Textҽ���� Or lngIndex = Text����) And Trim(TxtEdit(lngIndex).Text) = "" Then
            MsgBox "���š�ҽ���š�����������Ϊ�ա�", vbInformation, gstrSysName
            zlControl.TxtSelAll TxtEdit(lngIndex)
            TxtEdit(lngIndex).SetFocus
            Exit Function
        End If
    Next
    
    
    If Val(TxtEdit(Text�ʻ������ۼ�).Text) < Val(TxtEdit(Text�ʻ�֧���ۼ�).Text) Then
        MsgBox "�ʻ��ۼ�֧�����ܳ����ʻ��ۼ����ӡ�", vbInformation, gstrSysName
        zlControl.TxtSelAll TxtEdit(Text�ʻ�֧���ۼ�)
        TxtEdit(Text�ʻ�֧���ۼ�).SetFocus
        Exit Function
    End If
    
    If Val(TxtEdit(Text����ͳ���ۼ�).Text) < Val(TxtEdit(Textͳ�ﱨ���ۼ�).Text) Then
        MsgBox "ͳ�ﱨ���ۼƲ��ܳ�������ͳ���ۼơ�", vbInformation, gstrSysName
        zlControl.TxtSelAll TxtEdit(Textͳ�ﱨ���ۼ�)
        TxtEdit(Textͳ�ﱨ���ۼ�).SetFocus
        Exit Function
    End If
    
    IsValid = True
End Function

Private Sub cmdSelect_Click(Index As Integer)
    Dim rsTemp As ADODB.Recordset
    
    Select Case Index
        Case Select����
            gstrSQL = " Select A.����ID as ID,A.����,A.ҽ����,B.����,B.�Ա�,B.��������,B.���֤��,C.��� as ����ID " & _
                    " ,A.��Ա���,A.��λ����,A.����ID,D.���� as ����,A.��ְ as ��ְID,A.����֤��" & _
                    " From �����ʻ� A,������Ϣ B,��������Ŀ¼ C,���ղ��� D" & _
                    "  where A.����ID=B.����ID and A.����=" & mint���� & _
                    "  and A.����=C.���� and A.����=C.��� and A.����ID=D.ID(+)"
            
            Call Get�ʻ����
            zlControl.TxtSelAll TxtEdit(Text����)
            TxtEdit(Text����).SetFocus
        Case Select��λ
            Set rsTemp = frmPubSel.ShowSelect(Me, _
                    " Select ID,�ϼ�ID,ĩ��,����,����,��ַ,�绰,��������,�ʺ�,��ϵ�� From ��Լ��λ" & _
                    " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID", _
                    2, "������λ", , TxtEdit(Text���˵�λ).Text)
            If Not rsTemp Is Nothing Then
                TxtEdit(Text���˵�λ).Text = rsTemp("����")
                zlControl.TxtSelAll TxtEdit(Text���˵�λ)
            End If
            TxtEdit(Text���˵�λ).SetFocus
        
        Case Select����
            gstrSQL = " Select A.ID,A.����,A.����,A.����,decode(A.���,1,'���Բ�',2,'���ֲ�','��ͨ��') as ��� " & _
                    " From ���ղ��� A where A.����=" & mint����
            
            Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 0, "ҽ������", , TxtEdit(Text����).Text)
            If Not rsTemp Is Nothing Then
                TxtEdit(Text����).Text = rsTemp("����")
                TxtEdit(Text����).Tag = rsTemp("ID")
                zlControl.TxtSelAll TxtEdit(Text����)
            End If
            TxtEdit(Text����).SetFocus
    End Select
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll TxtEdit(Index)
    Select Case Index
        Case Text����
            zlCommFun.OpenIme True
        Case Else
            zlCommFun.OpenIme False
    End Select
End Sub

Private Sub dtp����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = Text���� Then
        If KeyCode = vbKeyDelete Then
            TxtEdit(Text����).Text = ""
            TxtEdit(Text����).Tag = ""
        End If
    End If
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strCode As String
    Dim str���� As String
    Dim rsTemp As New ADODB.Recordset
    
    If Index = Text���� Then
        If Len(TxtEdit(Text����).Text) = TxtEdit(Text����).MaxLength Or KeyAscii = vbKeyReturn Then
            strCode = Replace(Trim(TxtEdit(Text����).Text), "'", "")
            
            If IsNumeric(strCode) And IsNumeric(Left(strCode, 1)) Then 'ˢ��
                str���� = " and A.����='" & strCode & "' and A.����=" & cmb����.ItemData(cmb����.ListIndex)
            ElseIf (Left(strCode, 1) = "A" Or Left(strCode, 1) = "-") And IsNumeric(Mid(strCode, 2)) Then '����ID
                str���� = " and A.����ID=" & Mid(strCode, 2)
            ElseIf (Left(strCode, 1) = "B" Or Left(strCode, 1) = "+") And IsNumeric(Mid(strCode, 2)) Then 'סԺ��(��ס(��)Ժ�Ĳ���)
                str���� = " and B.סԺ��=" & Mid(strCode, 2)
            ElseIf (Left(strCode, 1) = "D" Or Left(strCode, 1) = "*") And IsNumeric(Mid(strCode, 2)) Then '�����(�������ﲡ��)
                str���� = " and B.�����=" & Mid(strCode, 2)
            Else '��������
                str���� = " and B.����='" & strCode & "'"
            End If
        
            gstrSQL = " Select A.����ID as ID,A.����,A.ҽ����,B.����,B.�Ա�,B.��������,B.���֤��,C.��� as ����ID " & _
                    " ,A.��Ա���,A.��λ����,A.����ID,D.���� as ����,A.��ְ as ��ְID,A.����֤��" & _
                    " From �����ʻ� A,������Ϣ B,��������Ŀ¼ C,���ղ��� D" & _
                    "  where A.����ID=B.����ID and A.����=" & mint���� & _
                    "  and A.����=C.���� and A.����=C.��� and A.����ID=D.ID(+)" & str����
            
            Call Get�ʻ����
        End If
    End If
    
    If KeyAscii = vbKeyReturn Then
        
        KeyAscii = 0  '��������
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub cmb����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub cmb�Ա�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    If Index >= TextסԺ���� And Index <= Textͳ�ﱨ���ۼ� Then
        If Index = TextסԺ���� Then
            TxtEdit(Index).Text = Format(Val(TxtEdit(Index).Text), "#0;0;0;0")
        Else
            TxtEdit(Index).Text = Format(Val(TxtEdit(Index).Text), "######0.00;0.00;0.00;0.00")
        End If
    End If

End Sub

Private Sub Get�ʻ����()
'���Ѿ����ڵļ�¼�ж����ʻ���Ϣ
    Dim rs�ʻ� As ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim lngIndex As Long
    
    
    Set rs�ʻ� = frmPubSel.ShowSelect(Me, gstrSQL, 0, "�����ʻ�", , TxtEdit(Text����).Text, "", False, True)
    If Not rs�ʻ� Is Nothing Then
    
        TxtEdit(Text����).Text = rs�ʻ�("����")
        '�������õ�����
        TxtEdit(Textҽ����).Text = IIf(IsNull(rs�ʻ�("ҽ����")), "", rs�ʻ�("ҽ����"))
        TxtEdit(Text����).Text = IIf(IsNull(rs�ʻ�("����")), "", rs�ʻ�("����"))
        TxtEdit(Text���֤��).Text = IIf(IsNull(rs�ʻ�("���֤��")), "", rs�ʻ�("���֤��"))
        TxtEdit(Text��Ա���).Text = IIf(IsNull(rs�ʻ�("��Ա���")), "", rs�ʻ�("��Ա���"))
        TxtEdit(Text���˵�λ).Text = IIf(IsNull(rs�ʻ�("��λ����")), "", rs�ʻ�("��λ����"))
        TxtEdit(Text����).Text = IIf(IsNull(rs�ʻ�("����")), "", rs�ʻ�("����"))
        TxtEdit(Text����).Tag = IIf(IsNull(rs�ʻ�("����ID")), "", rs�ʻ�("����ID"))
        
        Call SetComboByText(cmb�Ա�, IIf(IsNull(rs�ʻ�("�Ա�")), "", rs�ʻ�("�Ա�")), True)
        Cbo��ǰ״̬.ListIndex = rs�ʻ�("��ְID") - 1
        TxtEdit(Text����֤��).Text = ""
        If Cbo��ǰ״̬.ListIndex <> 0 Then
            TxtEdit(Text����֤��).Text = IIf(IsNull(rs�ʻ�("����֤��")), "", rs�ʻ�("����֤��"))
        End If
        If IsNull(rs�ʻ�("��������")) = False Then
            dtp����.Value = rs�ʻ�("��������")
        End If
        
        For lngIndex = 0 To cmb����.ListCount - 1
            If cmb����.ItemData(lngIndex) = rs�ʻ�("����ID") Then
                cmb����.ListIndex = lngIndex
                Exit For
            End If
        Next
        
        '�ٶ����ʻ������Ϣ
        gstrSQL = "select * from �ʻ������Ϣ where ����=[1]" & _
            " and ����ID=[2] and ���=[3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mint����, CLng(rs�ʻ�("ID")), Year(dtp����.MaxDate))
        
        If rsTemp.EOF = False Then
            '�����ʻ����
            TxtEdit(TextסԺ����).Text = Format(rsTemp("סԺ�����ۼ�"), "#0;0;0;0")
            TxtEdit(Text�ʻ������ۼ�).Text = Format(rsTemp("�ʻ������ۼ�"), "######0.00;0.00;0.00;0.00")
            TxtEdit(Text�ʻ�֧���ۼ�).Text = Format(rsTemp("�ʻ�֧���ۼ�"), "######0.00;0.00;0.00;0.00")
            TxtEdit(Text����ͳ���ۼ�).Text = Format(rsTemp("����ͳ���ۼ�"), "######0.00;0.00;0.00;0.00")
            TxtEdit(Textͳ�ﱨ���ۼ�).Text = Format(rsTemp("ͳ�ﱨ���ۼ�"), "######0.00;0.00;0.00;0.00")
        Else
            TxtEdit(TextסԺ����).Text = "0"
            TxtEdit(Text�ʻ������ۼ�).Text = "0.00"
            TxtEdit(Text�ʻ�֧���ۼ�).Text = "0.00"
            TxtEdit(Text����ͳ���ۼ�).Text = "0.00"
            TxtEdit(Textͳ�ﱨ���ۼ�).Text = "0.00"
        End If
        
    End If
End Sub

