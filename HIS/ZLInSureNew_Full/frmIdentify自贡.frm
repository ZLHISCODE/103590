VERSION 5.00
Begin VB.Form frmIdentify�Թ� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����֤"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   Icon            =   "frmIdentify�Թ�.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.TextBox txt���� 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1770
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   930
      Width           =   1665
   End
   Begin VB.CheckBox chk������Ա 
      Caption         =   "������Ա"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.CommandButton cmd�޸����� 
      Caption         =   "������(&M)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5280
      TabIndex        =   27
      Top             =   4440
      Width           =   1305
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5310
      TabIndex        =   26
      Top             =   1140
      Width           =   1305
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5310
      TabIndex        =   25
      Top             =   570
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      Height          =   6405
      Left            =   4950
      TabIndex        =   28
      Top             =   -240
      Width           =   30
   End
   Begin VB.TextBox txtҽ������ 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1770
      TabIndex        =   24
      Top             =   4590
      Width           =   2895
   End
   Begin VB.TextBox txt�ʻ���� 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1770
      TabIndex        =   22
      Top             =   4140
      Width           =   2895
   End
   Begin VB.TextBox txtְ����� 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1770
      TabIndex        =   20
      Top             =   3690
      Width           =   2895
   End
   Begin VB.TextBox txt�������� 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1770
      TabIndex        =   18
      Top             =   4140
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txt��״̬ 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1770
      TabIndex        =   16
      Top             =   3690
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.TextBox txt�������� 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1770
      TabIndex        =   14
      Top             =   3210
      Width           =   2865
   End
   Begin VB.TextBox txt�Ա� 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1770
      TabIndex        =   10
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txt���֤�� 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1770
      TabIndex        =   12
      Top             =   2730
      Width           =   2865
   End
   Begin VB.TextBox txt���� 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1770
      TabIndex        =   8
      Top             =   1830
      Width           =   1995
   End
   Begin VB.TextBox txt���� 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1770
      TabIndex        =   6
      Top             =   1380
      Width           =   2865
   End
   Begin VB.TextBox txtҽ����� 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1770
      TabIndex        =   2
      Top             =   480
      Width           =   2865
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&P)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   810
      TabIndex        =   3
      Top             =   975
      Width           =   840
   End
   Begin VB.Label lblҽ������ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ������(&Z)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   23
      Top             =   4635
      Width           =   1320
   End
   Begin VB.Label lbl�ʻ���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�ʻ����(&E)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   21
      Top             =   4185
      Width           =   1320
   End
   Begin VB.Label lblְ����� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ְ�����(&F)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   19
      Top             =   3735
      Width           =   1320
   End
   Begin VB.Label lbl�������� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������(&D)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   17
      Top             =   4185
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label lbl��״̬ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��״̬(&T)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   570
      TabIndex        =   15
      Top             =   3735
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label lbl�������� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������(&B)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   13
      Top             =   3270
      Width           =   1320
   End
   Begin VB.Label lbl���֤�� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���֤��(&I)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   11
      Top             =   2790
      Width           =   1320
   End
   Begin VB.Label lbl�Ա� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�Ա�(&S)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   810
      TabIndex        =   9
      Top             =   2325
      Width           =   840
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&X)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   810
      TabIndex        =   7
      Top             =   1875
      Width           =   840
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&K)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   810
      TabIndex        =   5
      Top             =   1425
      Width           =   840
   End
   Begin VB.Label lblҽ����� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ�����(&Y)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   1
      Top             =   525
      Width           =   1320
   End
End
Attribute VB_Name = "frmIdentify�Թ�"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrReturn As String
Private mbytType As Byte
Private mlng����ID As Long
Private mbln�ж���Ժ As Boolean
Private mbln�޸����� As Boolean
Private mstrPass As String      '���浱ǰ����

Private Function IsValid() As Boolean
'���ܣ��ж�IC���Ƿ�Ϸ�
    Dim rsTemp As New ADODB.Recordset
    Dim str��Ч�� As String
    Dim bln����ҽ�� As Boolean
    
    If Me.txt����.Text = "" Then
        MsgBox "ҽ�����˵���ݻ�δȷ�ϣ����ȶ�����", vbInformation, gstrSysName
        Exit Function
    End If
    
    'G�� �ж�ְ���Ƿ���סԺ���ж�IC����InpatientFlag����סԺ���㲻���д��жϣ�
    If mbln�ж���Ժ = True Then
        gstrSQL = "Select Nvl(��ǰ״̬,0) AS ״̬ From �����ʻ� Where ����=[1] And ҽ����=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�жϵ�ǰ�����Ƿ���Ժ", TYPE_�Ĵ��Թ�, CStr(txtҽ�����.Text))
        If rsTemp.RecordCount <> 0 Then
            If rsTemp!״̬ = 1 Then
                MsgBox "��ǰ����Ŀǰ����Ժ�������ƣ��������ٴ���Ժ��", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    IsValid = True
End Function

Public Function GetPatient(ByVal bytType As Byte, Optional lng����ID As Long, _
    Optional ByVal bln�ж���Ժ As Boolean, Optional ByVal bln�޸����� As Boolean = False) As String
    
    mstrReturn = ""
    mbytType = bytType
    mlng����ID = lng����ID
    mbln�ж���Ժ = bln�ж���Ժ
    mbln�޸����� = bln�޸�����
    
    frmIdentify�Թ�.Show vbModal
    lng����ID = mlng����ID
    GetPatient = mstrReturn
End Function

Private Sub chk������Ա_Click()
    txtҽ�����.Enabled = (chk������Ա.Value = 1)
    If txtҽ�����.Enabled Then
        txtҽ�����.SetFocus
    Else
        txt����.SetFocus
    End If
End Sub

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    '�����ӹ����嵥
    'IsValid:�Ա�Ҫ״̬���м��
    
    Dim rsTemp As New ADODB.Recordset
    Dim rsSelected As New ADODB.Recordset
    Dim rs���� As New ADODB.Recordset
    Dim strIdentify As String, strAddition As String, strBirthday As String
    Dim lng���� As Long, str���� As String
    Dim int��ǰ״̬ As Integer
    Dim datToday As Date
    
    '����������Ϣ��
    If Not IsValid() Then Exit Sub
    
    'ȡ�ò��˵ĵ�ǰ״̬
    gstrSQL = "Select Nvl(��ǰ״̬,0) AS ��ǰ״̬ From �����ʻ� Where ҽ����=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�ò��˵ĵ�ǰ״̬", CStr(Trim(txtҽ�����.Text)), TYPE_�Ĵ�üɽ)
    int��ǰ״̬ = 0
    If rsTemp.RecordCount <> 0 Then
        int��ǰ״̬ = rsTemp!��ǰ״̬
    End If
    
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�,23�������� (1����������)
    strIdentify = TrimStr(txtҽ�����.Text)                               '0����
    strIdentify = strIdentify & ";" & TrimStr(txtҽ�����.Text)   '1ҽ����
    strIdentify = strIdentify & ";"                               '2����
    strIdentify = strIdentify & ";" & TrimStr(txt����.Text)       '3����
    strIdentify = strIdentify & ";" & txt�Ա�.Text    '4�Ա�
    
    strBirthday = TrimStr(txt��������.Text)
    datToday = zlDatabase.Currentdate
    If strBirthday = "" Then
        strBirthday = Format(datToday, "yyyy-MM-dd")
    Else
        strBirthday = Mid(strBirthday, 1, 4) & "-" & Mid(strBirthday, 5, 2) & "-" & Mid(strBirthday, 7, 2)
    End If
    strIdentify = strIdentify & ";" & strBirthday              '5��������
    strIdentify = strIdentify & ";" & TrimStr(txt���֤��.Text)    '6���֤
    strIdentify = strIdentify & ";" & TrimStr(txt��������.Tag) & "(" & TrimStr(txt��������.Tag) & ")"   '7.��λ����(����)
    
    '�õ�ԭסԺ����
    If mbytType <> 1 Then
        gstrSQL = "Select Nvl(����ID,0) ����ID From �����ʻ� Where ����=[1] And ҽ����=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�õ�ԭסԺ����", TYPE_�Ĵ��Թ�, CStr(TrimStr(txtҽ�����.Text)))
        If Not rsTemp.EOF Then
            lng���� = rsTemp!����ID
        End If
    End If

    strAddition = ";" & Val(txtҽ������.Tag)                    '8.���Ĵ���
    strAddition = strAddition & ";"                             '9.˳���
    strAddition = strAddition & ";" & Mid(txtְ�����.Text, 1, InStr(1, txtְ�����.Text, "-") - 1)   '10��Ա���
    strAddition = strAddition & ";" & Val(txt�ʻ����.Text)   '11�ʻ����
    strAddition = strAddition & ";" & int��ǰ״̬             '12��ǰ״̬
    strAddition = strAddition & ";" & IIf(lng���� > 0, lng����, "") '13����ID

    strAddition = strAddition & ";" & Val(txtְ�����.Tag)
    strAddition = strAddition & ";"                             '15����֤��
    strAddition = strAddition & ";" & DateDiff("yyyy", CDate(strBirthday), datToday) '16�����
    strAddition = strAddition & ";"                             '17�Ҷȼ�
    strAddition = strAddition & ";" & Val(txt�ʻ����.Text)     '18�ʻ������ۼ�
    strAddition = strAddition & ";"                             '19�ʻ�֧���ۼ�
    strAddition = strAddition & ";"                             '20����ͳ���ۼ�
    strAddition = strAddition & ";"                             '21ͳ�ﱨ���ۼ�
    strAddition = strAddition & ";"                             '22סԺ�����ۼ�
    strAddition = strAddition & ";"                             '23�������� (1����������)
    
    mlng����ID = BuildPatiInfo(mbytType, strIdentify & strAddition, mlng����ID, TYPE_�Ĵ��Թ�)
    '���ظ�ʽ:�м���벡��ID
    mstrReturn = strIdentify & ";" & mlng����ID & strAddition
    
    If mbytType = 1 Then
'        gstrSQL = "zl_������Ϣ_INSERT(" & TYPE_�Ĵ��Թ� & "," & mlng����ID & ",'" & str���� & "')"
'        gcn�Թ�.Execute gstrSQL, , adCmdStoredProc
    End If
    
    Unload Me
End Sub

Private Sub cmd�޸�����_Click()
    Dim strPass As String
    Dim StrInput As String, strOutput As String
    
    strPass = frm�޸�����.ChangePassword("")
    If strPass = "" Then Exit Sub
    
    '�����޸����뺯��
    StrInput = Me.txt����.Text & "|" & strPass
    If Not ���ýӿ�_�Թ�(ҵ������_�Թ�.�޸�����, StrInput, strOutput) Then Exit Sub
    
    '���µ�ǰ�����ϵ�����
    Me.txt����.Text = strPass
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    cmd�޸�����.Visible = mbln�޸�����
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    '�����ӹ����嵥
    'Get������Ա:��������Ա�Ļ�����Ϣ�������ӿڵķ��ظ�ʽ��֯����
    
    Dim StrInput As String
    Dim strOutput As String
    Dim arrOutput
    Dim rsTemp As New ADODB.Recordset
    
    '������������ֵ����
    Const cintҽ����� As Integer = 0
    Const cint���� As Integer = 1
    Const cint���� As Integer = 2
    Const cint�Ա� As Integer = 3
    Const cint���֤�� As Integer = 4
    Const cint�������� As Integer = 5
    Const cint��״̬ As Integer = 6
    Const cint�������� As Integer = 7
    Const cint��λ���� As Integer = 8
    Const cintְ����� As Integer = 9
    Const cint�ʻ���� As Integer = 10
    Const cint���Ĵ��� As Integer = 11
    Const cintҽ����� As Integer = 12
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
'    If chk������Ա.Value = 0 Then
        StrInput = Me.txt����.Text
        If Not ���ýӿ�_�Թ�(ҵ������_�Թ�.����, StrInput, strOutput) Then Exit Sub
'    Else
'        strInput = Trim(txtҽ�����.Text)
'        strOutPut = Get������Ա(strInput)
'        If strOutPut = "" Then
'            txtҽ�����.SetFocus
'            Exit Sub
'        End If
'    End If
    'ҽ�����|����|����|�Ա�|���֤��|��������|��״̬(���)|����������Ϣ|���˵�λ����
    '|ְ����ݣ�0x-��ְ��1x-����, 05��11Ϊһ���Խɷ�,7x�����Ҽ��˲о���|�����˻����|���Ĵ���|ҽ�����
    arrOutput = Split(strOutput, "|")
    
    '��ȡҽ�����ĵ����������
    gstrSQL = "Select ���,���� From ��������Ŀ¼ Where ����=[1] And ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҽ�����ĵ����������", CStr(arrOutput(cint���Ĵ���)), TYPE_�Ĵ��Թ�)
    If rsTemp.RecordCount = 0 Then
        MsgBox "�����ڸ�ҽ�����ģ����Ĵ���Ϊ��" & arrOutput(cint���Ĵ���), vbInformation, gstrSysName
        Exit Sub
    End If
    txtҽ������.Text = rsTemp!����
    txtҽ������.Tag = Val(rsTemp!���)
    Me.txtҽ�����.Tag = arrOutput(cint���Ĵ���)            'ҽ�����.tag�������Ĵ���
    
    Me.txtҽ�����.Text = arrOutput(cintҽ�����)
    Me.txt����.Text = arrOutput(cint����)
    Me.txt����.Text = arrOutput(cint����)
    Me.txt�Ա�.Text = arrOutput(cint�Ա�)
    Me.txt���֤��.Text = arrOutput(cint���֤��)
    Me.txt��������.Text = arrOutput(cint��������)
    Me.txt��״̬.Text = arrOutput(cint��״̬)
    Me.txt��������.Text = arrOutput(cint��������)
    Me.txt��������.Tag = arrOutput(cint��λ����)            '��λ���뱣���txt�������ص�Tag��
    
    If arrOutput(cintְ�����) Like "5*" Then
        Me.txtְ�����.Text = arrOutput(cintְ�����) & "-" & "����"
        Me.txtְ�����.Tag = 3
    ElseIf arrOutput(cintְ�����) Like "0*" Then
        Me.txtְ�����.Text = arrOutput(cintְ�����) & "-" & "��ְ"
        Me.txtְ�����.Tag = 1
    ElseIf arrOutput(cintְ�����) Like "7*" Then
        Me.txtְ�����.Text = arrOutput(cintְ�����) & "-" & "�����Ҽ��˲о���"
        Me.txtְ�����.Tag = 4
    Else
        Me.txtְ�����.Text = arrOutput(cintְ�����) & "-" & "����"
        Me.txtְ�����.Tag = 2
    End If
    
    Me.txt�ʻ����.Text = Format(Val(arrOutput(cint�ʻ����)), "#####0.00")
End Sub

Private Sub txtҽ�����_GotFocus()
    Call zlControl.TxtSelAll(txtҽ�����)
End Sub
