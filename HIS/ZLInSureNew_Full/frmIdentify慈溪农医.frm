VERSION 5.00
Object = "{E5918DE2-9E1A-472B-96C6-5AE5994F9138}#1.0#0"; "ReadBarComm.dll"
Begin VB.Form frmIdentify��Ϫũҽ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����֤"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   Icon            =   "frmIdentify��Ϫũҽ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CheckBox chk���俨�� 
      Caption         =   "�ֹ����뿨��"
      Height          =   375
      Left            =   1200
      TabIndex        =   40
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CheckBox Chk�º���� 
      Caption         =   "�º����"
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
      Height          =   255
      Left            =   1440
      TabIndex        =   39
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CheckBox chkС����ҩ 
      Alignment       =   1  'Right Justify
      Caption         =   "С����ҩ"
      Height          =   255
      Left            =   600
      TabIndex        =   34
      Top             =   4170
      Width           =   1335
   End
   Begin VB.TextBox txt�е����� 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Left            =   6090
      MaxLength       =   3
      TabIndex        =   31
      Top             =   3390
      Width           =   2025
   End
   Begin VB.ComboBox cboҽ����� 
      Height          =   300
      Left            =   1410
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   3390
      Width           =   2025
   End
   Begin VB.TextBox txt������Ϣ 
      Height          =   300
      Left            =   1410
      TabIndex        =   33
      Top             =   3780
      Width           =   6735
   End
   Begin VB.CommandButton cmdȡ�� 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6960
      TabIndex        =   37
      Top             =   4680
      Width           =   1100
   End
   Begin VB.CommandButton cmdȷ�� 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5700
      TabIndex        =   36
      Top             =   4680
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   30
      Left            =   30
      TabIndex        =   35
      Top             =   4500
      Width           =   8835
   End
   Begin VB.TextBox txt��״̬ 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1410
      TabIndex        =   13
      Top             =   3000
      Width           =   2025
   End
   Begin VB.TextBox txt�������� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   6090
      TabIndex        =   17
      Top             =   1050
      Width           =   2025
   End
   Begin VB.TextBox txt�ʻ���� 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Left            =   6090
      TabIndex        =   27
      Top             =   3000
      Width           =   2025
   End
   Begin VB.TextBox txt�ʻ��ۼ�֧�� 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Left            =   6090
      TabIndex        =   25
      Top             =   2610
      Width           =   2025
   End
   Begin VB.TextBox txtסԺ�ۼ�ʵ�� 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Left            =   6090
      TabIndex        =   23
      Top             =   2220
      Width           =   2025
   End
   Begin VB.TextBox txt���������ۼƱ��� 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Left            =   6090
      TabIndex        =   21
      Top             =   1830
      Width           =   2025
   End
   Begin VB.TextBox txt�����ۼƱ��� 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   300
      Left            =   6090
      TabIndex        =   19
      Top             =   1440
      Width           =   2025
   End
   Begin VB.TextBox txt���ִ��� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   6090
      TabIndex        =   15
      Top             =   660
      Width           =   1395
   End
   Begin VB.TextBox txt�������� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1410
      TabIndex        =   11
      Top             =   2610
      Width           =   2025
   End
   Begin VB.TextBox txt���֤�� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1410
      TabIndex        =   9
      Top             =   2220
      Width           =   2025
   End
   Begin VB.TextBox txt�Ա� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1410
      TabIndex        =   7
      Top             =   1830
      Width           =   555
   End
   Begin VB.TextBox txt���� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1410
      TabIndex        =   5
      Top             =   1440
      Width           =   1185
   End
   Begin VB.TextBox txtҽ��֤�� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1410
      TabIndex        =   3
      Top             =   1050
      Width           =   1575
   End
   Begin VB.TextBox txt��֤���� 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1410
      MaxLength       =   30
      TabIndex        =   1
      Top             =   660
      Width           =   2835
   End
   Begin READBARCOMMLibCtl.ReadBar2Comm ReadCard 
      Height          =   375
      Left            =   1320
      OleObjectBlob   =   "frmIdentify��Ϫũҽ.frx":000C
      TabIndex        =   38
      Top             =   5400
      Width           =   5175
   End
   Begin VB.Label lbl�е����� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�е�����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5280
      TabIndex        =   30
      Top             =   3450
      Width           =   720
   End
   Begin VB.Label lblҽ����� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ�����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   600
      TabIndex        =   28
      Top             =   3450
      Width           =   720
   End
   Begin VB.Label lbl������Ϣ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������Ϣ"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   600
      TabIndex        =   32
      Top             =   3840
      Width           =   720
   End
   Begin VB.Label lbl��״̬ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��״̬"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   780
      TabIndex        =   12
      Top             =   3060
      Width           =   540
   End
   Begin VB.Label lbl�������� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���ⲡ������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4920
      TabIndex        =   16
      Top             =   1110
      Width           =   1080
   End
   Begin VB.Label lbl�ʻ���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�ʻ����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   5280
      TabIndex        =   26
      Top             =   3060
      Width           =   720
   End
   Begin VB.Label lbl�ʻ��ۼ�֧�� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�ʻ��ۼ�֧��"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4920
      TabIndex        =   24
      Top             =   2670
      Width           =   1080
   End
   Begin VB.Label lblסԺ�ۼ�ʵ�� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "סԺ�ۼ�ʵ��"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4920
      TabIndex        =   22
      Top             =   2280
      Width           =   1080
   End
   Begin VB.Label lbl���������ۼƱ��� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���������ۼƱ���"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4560
      TabIndex        =   20
      Top             =   1890
      Width           =   1440
   End
   Begin VB.Label lbl�����ۼƱ��� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�����ۼƱ���"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4920
      TabIndex        =   18
      Top             =   1500
      Width           =   1080
   End
   Begin VB.Label lbl���ִ��� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���ⲡ�ִ���"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4920
      TabIndex        =   14
      Top             =   720
      Width           =   1080
   End
   Begin VB.Label lbl�������� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   600
      TabIndex        =   10
      Top             =   2670
      Width           =   720
   End
   Begin VB.Label lbl���֤�� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���֤��"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   600
      TabIndex        =   8
      Top             =   2280
      Width           =   720
   End
   Begin VB.Label lbl�Ա� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�Ա�"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   960
      TabIndex        =   6
      Top             =   1890
      Width           =   360
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   960
      TabIndex        =   4
      Top             =   1500
      Width           =   360
   End
   Begin VB.Label lblҽ��֤�� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ��֤��"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   600
      TabIndex        =   2
      Top             =   1110
      Width           =   720
   End
   Begin VB.Label lbl��֤���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��֤����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   720
   End
End
Attribute VB_Name = "frmIdentify��Ϫũҽ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType As Byte            '0-�����շѣ�1-��Ժ�Ǽǣ�2-������������סԺ,3-�Һ�
Private mlng����ID As Long
Private mstrReturn As String

Private Sub cboҽ�����_Click()
    Me.txt�е�����.Enabled = False
    If cboҽ�����.ItemData(cboҽ�����.ListIndex) = 22 Then
        '��ͨ�¹�
        Me.txt�е�����.Enabled = True
        Me.txt�е�����.SetFocus
    End If
End Sub

Private Sub cboҽ�����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Chk�º����_Click()
Dim strDate As String
strDate = Format(zlDatabase.Currentdate, "yymmddhhmmss")
    If Chk�º����.Value = 1 Then
       txt��֤����.Text = UserInfo.�û��� + strDate
       txtҽ��֤��.Text = "�º󲹱�" + strDate
       txt����.Enabled = True
       txt�Ա�.Enabled = True
       txt���֤��.Enabled = True
       txt��������.Enabled = True
       txt�����ۼƱ���.Text = 0
       txt���������ۼƱ���.Text = 0
       txt�ʻ����.Text = 0
      chk���俨��.Enabled = False
     Else
        txt��֤����.Text = ""
       txtҽ��֤��.Text = ""
       txt����.Enabled = False
       txt�Ա�.Enabled = False
       txt���֤��.Enabled = False
       txt��������.Enabled = False
       txt�����ۼƱ���.Text = 0
       txt���������ۼƱ���.Text = 0
       txt�ʻ����.Text = 0
        chk���俨��.Enabled = True
    End If
End Sub

Private Sub cmd����_Click()

End Sub

Private Sub chk���俨��_Click()
If chk���俨��.Value = 1 Then
   txt��֤����.Enabled = True
Else
   txt��֤����.Enabled = False
End If

End Sub

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub cmdȷ��_Click()
    Dim lng����ID As Long
    Dim str�������� As String
    Dim strIdentify As String, strAddition As String
    Dim rsTemp As New ADODB.Recordset
    
    If Trim(txt��֤����.Text) = "" Then
        MsgBox "�������", vbInformation, gstrSysName
        txt��֤����.SetFocus
        Exit Sub
    End If
    If Trim(txt����.Text) = "" Then
        MsgBox "��û�л�ȡ��ҽ�����˵������Ϣ������ͨ����֤��", vbInformation, gstrSysName
        txt��֤����.SetFocus
        Exit Sub
    End If
    If Me.cboҽ�����.ItemData(Me.cboҽ�����.ListIndex) = 22 Then
        If Val(txt�е�����.Text) < 0 Then
            MsgBox "�е���������С���㣡", vbInformation, gstrSysName
            txt�е�����.SetFocus
            Exit Sub
        End If
        If Val(txt�е�����.Text) > 100 Then
            MsgBox "�е��������ܴ���һ�٣�", vbInformation, gstrSysName
            txt�е�����.SetFocus
            Exit Sub
        End If
    End If
  ' txt������Ϣ.Tag = 124
    If mbytType <> 3 Then
     
    
       If Val(txt������Ϣ.Tag) = 0 Then
           ' MsgBox "�����벡�˵ļ�����Ϣ��", vbInformation, gstrSysName
           'txt������Ϣ.SetFocus
         ' Exit Sub
          txt������Ϣ.Tag = 999
      End If
    
        
    End If
    
    If mbytType <> 2 Then
        '��鲡��״̬
        gstrSQL = "select nvl(��ǰ״̬,0) as ״̬ from �����ʻ� where ����=[1] and ҽ����=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_��Ϫũҽ, txtҽ��֤��)
        If rsTemp.RecordCount > 0 Then
            If rsTemp("״̬") > 0 Then
                MsgBox "�ò����Ѿ���Ժ������ͨ�������֤��", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
    Else
        '�����������סԺ�ģ�ֻ��ˢ����ʾһ�����ݶ��ѣ�������
        Unload Me
        Exit Sub
    End If
    
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�
    strIdentify = txt��֤����.Text                              '0����
    strIdentify = strIdentify & ";" & txtҽ��֤��.Text          '1ҽ����
    strIdentify = strIdentify & ";"                             '2����
    strIdentify = strIdentify & ";" & txt����.Text              '3����
    strIdentify = strIdentify & ";" & txt�Ա�.Text              '4�Ա�
    strIdentify = strIdentify & ";" & txt��������.Text          '5��������
    strIdentify = strIdentify & ";" & txt���֤��.Text          '6���֤
    strIdentify = strIdentify & ";"                             '7.��λ����(����)
    strAddition = ";0"                                          '8.���Ĵ���
    strAddition = strAddition & ";"                             '9.˳���
    strAddition = strAddition & ";" '10��Ա���
 
  strAddition = strAddition & ";" & Val(txt�ʻ����.Text)     '11�ʻ����
    strAddition = strAddition & ";0"                            '12��ǰ״̬
    strAddition = strAddition & ";" & Val(txt������Ϣ.Tag)                 '13����ID
    strAddition = strAddition & ";1"                            '14��ְ(1,2,3)
    strAddition = strAddition & ";" & Val(txt�е�����.Text)     '15����֤��
    strAddition = strAddition & ";"                             '16�����
    strAddition = strAddition & ";"                             '17�Ҷȼ�
    strAddition = strAddition & ";" & Val(txt�ʻ����.Text) + Val(txt�ʻ��ۼ�֧��.Text)   '18�ʻ������ۼ�
    strAddition = strAddition & ";" & Val(txt�ʻ��ۼ�֧��.Text)                           '19�ʻ�֧���ۼ�
    strAddition = strAddition & ";0"                            '20���깤���ܶ�
    strAddition = strAddition & ";"                             '21סԺ�����ۼ�
    
    mlng����ID = BuildPatiInfo(0, strIdentify & strAddition, mlng����ID, TYPE_��Ϫũҽ)
    '���ظ�ʽ:�м���벡��ID
    If mlng����ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng����ID & strAddition
    End If
    
    With gComInfo_��Ϫũҽ
        .ҽ��֤�� = txtҽ��֤��.Text
        .���˱�� = txt��֤����.Text
        .ҵ������ = Me.cboҽ�����.ItemData(Me.cboҽ�����.ListIndex)
    End With
    If Chk�º����.Value = 1 Then
       gComInfo_��Ϫũҽ.�������� = "�º󲹱�"
    Else
       gComInfo_��Ϫũҽ.�������� = "ʵʱ����"
    End If
    
    '���±����ʻ������Ϣ��ҵ�����ͣ�
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_��Ϫũҽ & ",'С����ҩ','" & chkС����ҩ.Value & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҵ������")
    
    Unload Me

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strMsg As String
    Dim IntPort As Integer
    Dim lngReturn As Long
    Dim rsTemp As New ADODB.Recordset
    
    
    With Me.cboҽ�����
        .Clear
        '��������ҵ��
        If mbytType = 0 Or mbytType = 2 Or mbytType = 3 Then
             .AddItem "ũ������"
            .ItemData(.NewIndex) = 11
            .AddItem "��������"
            .ItemData(.NewIndex) = 12
            Chk�º����.Visible = True
            
        End If
        '����סԺҵ��
        If mbytType = 1 Or mbytType = 2 Then
            .AddItem "��ͨסԺ"
            .ItemData(.NewIndex) = 21
            .AddItem "��ͨ�¹�"
            .ItemData(.NewIndex) = 22
            .AddItem "�󲡾���"
            .ItemData(.NewIndex) = 23
            .AddItem "�Ѳ�"
            .ItemData(.NewIndex) = 24
            .AddItem "����"
            .ItemData(.NewIndex) = 25
            
            chkС����ҩ.Enabled = True
        End If
        .ListIndex = 0
    End With
    
    '�Һſ��Բ������뼲���벢��֢��Ϣ
    If mbytType = 3 Then
        txt������Ϣ.Enabled = False
    End If
    
    'ȡIC�˿ں�
    gstrSQL = "Select ����ֵ From ���ղ��� Where ����=[1] ANd ������='IC�˿ں�'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡIC�˿ں�", TYPE_��Ϫũҽ)
    If rsTemp.RecordCount = 0 Then
        IntPort = 1
    Else
        IntPort = Nvl(rsTemp!����ֵ, 1)
    End If
    
    '��ʼ����������
    lngReturn = ReadCard.OpenPort(IntPort)
    Select Case lngReturn
    Case 0, 1
        strMsg = ""
    Case -1
        strMsg = "�򿪴���ʧ��(һ�����ڶ˿ںŲ����ڻ�ռ��)"
    Case -2
        strMsg = "��ȡ����״̬ʧ��"
    Case -3
        strMsg = "���ö˿�״̬ʧ��"
    Case -4
        strMsg = "�����������߳�ʧ��(ϵͳ��Դ����)"
    Case -5
        strMsg = "�ڲ�����"
    End Select
    If strMsg <> "" Then
        MsgBox strMsg, vbInformation, gstrSysName
        Exit Sub
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ReadCard.ClosePort
End Sub

Private Sub ReadCard_OnComm(ByVal strData As String, ByVal lValidData As Long)
    Me.txt��֤����.Text = Mid(strData, 1, 20)
    If Trim(txt��֤����.Text) <> "" Then Call txt��֤����_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub txt�е�����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt������Ϣ_GotFocus()
    Call zlControl.TxtSelAll(txt������Ϣ)
End Sub

Private Sub txt������Ϣ_KeyPress(KeyAscii As Integer)
    Dim strLike As String
    Dim StrInput As String
    Dim str�Ա� As String
    Dim blnCancel As Boolean
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errHandle
    If KeyAscii <> vbKeyReturn Then Exit Sub
    
    If txt������Ϣ.Text = lbl������Ϣ.Tag And txt������Ϣ.Text <> "" Then
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf txt������Ϣ.Text = "" Then
        txt������Ϣ.Tag = "": lbl������Ϣ.Tag = ""
        Call zlCommFun.PressKey(vbKeyTab) '��������
    Else
        strLike = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = 0, "%", "")
        StrInput = UCase(txt������Ϣ.Text)
        str�Ա� = txt�Ա�.Text
        If str�Ա� = "��" Then
            str�Ա� = " And (A.�Ա�����='��' Or A.�Ա����� is NULL)"
        ElseIf str�Ա� = "Ů" Then
            str�Ա� = " And (A.�Ա�����='Ů' Or A.�Ա����� is NULL)"
        Else
            str�Ա� = ""
        End If
        gstrSQL = "Select A.ID,A.����,A.����,A.����,A.����,A.˵��,A.�Ա�����,B.���" & _
            " From ��������Ŀ¼ A,����������� B" & _
            " Where A.���=B.���� And A.��� Not IN('B','Z')" & _
            " And (A.���� Like '" & StrInput & "%'" & _
            " Or Upper(A.����) Like '" & strLike & StrInput & "%'" & _
            " Or Upper(A.����) Like '" & strLike & StrInput & "%'" & _
            " Or Upper(A.����) Like '" & strLike & StrInput & "%')" & _
            " And Rownum<=100" & str�Ա� & _
            " Order by A.���,A.����"
        Set rsTemp = zlDatabase.ShowSelect(Me, gstrSQL, 0, "��������Input", , , , , , True, _
            txt������Ϣ.Left + Me.Left, _
            txt������Ϣ.Top + Me.Top, txt������Ϣ.Height, blnCancel, , True)
        If Not rsTemp Is Nothing Then
            txt������Ϣ.Tag = rsTemp!ID
            txt������Ϣ.Text = "(" & rsTemp!���� & ")" & rsTemp!����
            lbl������Ϣ.Tag = txt������Ϣ.Text '���ڻָ���ʾ
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            If Not blnCancel Then
                MsgBox "û���ҵ�ƥ��ļ������롣", vbInformation, gstrSysName
            End If
            If lbl������Ϣ.Tag <> "" Then txt������Ϣ.Text = lbl������Ϣ.Tag
            txt������Ϣ.SetFocus
        End If
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub txt��֤����_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strReturn As String
    Dim str���￨�� As String   '03-25Ӧ���ܼ�
    Dim strũ������ As String   '03-25Ӧ���ܼ�
   
    
    Dim rsItem As New ADODB.Recordset '03-25Ӧ���ܼ�
    Dim arrReturn
    Const Returncode As Integer = 0         '�������0��ʾ�ɹ�
    Const Returninfo As Integer = 1         '��Ӧ�Ĵ�����ʾ
    Const Hzylhm As Integer = 2             '����ҽ�ƺ���
    Const Cyxm As Integer = 3               '��Ա����
    Const Cyxb As Integer = 4               '��Ա�Ա�
    Const Sfzhm As Integer = 5              '���֤����
    Const Jtdz As Integer = 6               '��ͥ��ַ
    Const Csrq As Integer = 7               '��������
    Const Kzt As Integer = 8                '��״̬
    Const Tsbzdm As Integer = 9             '���ⲡ�ִ���
    Const Tsbzmc As Integer = 10            '���ⲡ������
    Const Mzljsb As Integer = 11            '�����ۼ�ʵ��
    Const Tsmzljsb As Integer = 12          '���������ۼ�ʵ��
    Const Zyljsb As Integer = 13            'סԺ�ۼ�ʵ��
    Const Zhljzf As Integer = 14            '�ʻ��ۼ�֧��
    Const Zhye As Integer = 15              '�ʻ����
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    '03-25 Ӧ���ܼ�
   
    
    If Trim(txt��֤����.Text) = "" Then Exit Sub
    Call ���ýӿ�_׼��_��Ϫũҽ(gstrFunc��Ϫũҽ_GetPersonalInfo, "Kzhm=" & txt��֤����.Text)
    If Not ���ýӿ�_��Ϫũҽ Then Exit Sub
    
    strReturn = gstrOutput_��Ϫũҽ
    arrReturn = Split(strReturn, "&")
    Me.txtҽ��֤��.Text = Trim(Split(arrReturn(Hzylhm), "=")(1))
    Me.txt����.Text = Trim(Split(arrReturn(Cyxm), "=")(1))
    Me.txt�Ա�.Text = Trim(Split(arrReturn(Cyxb), "=")(1))
    Me.txt���֤��.Text = Trim(Split(arrReturn(Sfzhm), "=")(1))
    Me.txt��������.Text = Format(Trim(Split(arrReturn(Csrq), "=")(1)), "yyyy-MM-dd")
    Me.txt���ִ���.Text = Trim(Split(arrReturn(Tsbzdm), "=")(1))
    Me.txt��������.Text = Trim(Split(arrReturn(Tsbzmc), "=")(1))
    Me.txt�����ۼƱ���.Text = Trim(Split(arrReturn(Mzljsb), "=")(1))
    Me.txt���������ۼƱ���.Text = Trim(Split(arrReturn(Tsmzljsb), "=")(1))
    Me.txtסԺ�ۼ�ʵ��.Text = Trim(Split(arrReturn(Zyljsb), "=")(1))
    Me.txt�ʻ��ۼ�֧��.Text = Trim(Split(arrReturn(Zhljzf), "=")(1))
    Me.txt�ʻ����.Text = Trim(Split(arrReturn(Zhye), "=")(1))
    Me.txt��״̬.Text = Trim(Split(arrReturn(Kzt), "=")(1))
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Public Function GetPatient(Optional bytType As Byte, Optional lng����ID As Long = 0) As String
    mbytType = bytType
    mlng����ID = lng����ID
    mstrReturn = ""
    Me.Show 1
    lng����ID = mlng����ID
    GetPatient = mstrReturn
End Function

