VERSION 5.00
Begin VB.Form frmIdentify���� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�����֤"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7860
   Icon            =   "frmidentify����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.CheckBox chk�޿����� 
      Caption         =   "�޿�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   780
      TabIndex        =   0
      Top             =   180
      Width           =   3585
   End
   Begin VB.ComboBox cboҵ������ 
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
      Left            =   2235
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   5025
      Width           =   3690
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
      Left            =   2235
      TabIndex        =   12
      Top             =   2760
      Width           =   3705
   End
   Begin VB.CommandButton cmd������Ϣ 
      Caption         =   "��"
      Height          =   345
      Left            =   5580
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5490
      Width           =   330
   End
   Begin VB.TextBox txt������Ϣ 
      Alignment       =   1  'Right Justify
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
      Left            =   2235
      TabIndex        =   25
      Top             =   5490
      Width           =   3360
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
      Left            =   6360
      TabIndex        =   29
      Top             =   1200
      Width           =   1320
   End
   Begin VB.CommandButton cmdOK 
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
      Left            =   6360
      TabIndex        =   28
      Top             =   675
      Width           =   1320
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "����(&R)"
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
      Left            =   6360
      TabIndex        =   21
      Top             =   5310
      Width           =   1320
   End
   Begin VB.Frame frame1 
      Height          =   6255
      Left            =   6135
      TabIndex        =   27
      Top             =   -105
      Width           =   45
   End
   Begin VB.TextBox txtҽ��֤�� 
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
      Left            =   2235
      TabIndex        =   2
      Top             =   570
      Width           =   3705
   End
   Begin VB.TextBox txt���������� 
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
      Height          =   360
      Left            =   2235
      TabIndex        =   20
      Top             =   4575
      Width           =   3705
   End
   Begin VB.TextBox txt���������� 
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
      Height          =   360
      Left            =   2235
      TabIndex        =   18
      Top             =   4125
      Width           =   3705
   End
   Begin VB.TextBox txt��λ���� 
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
      Left            =   2235
      TabIndex        =   16
      Top             =   3660
      Width           =   3705
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
      Left            =   2235
      TabIndex        =   14
      Top             =   3210
      Width           =   3705
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
      Height          =   360
      Left            =   2235
      TabIndex        =   10
      Top             =   2310
      Width           =   1290
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
      Height          =   360
      Left            =   2235
      TabIndex        =   8
      Top             =   1875
      Width           =   3705
   End
   Begin VB.TextBox txt����Ա 
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
      Left            =   2235
      TabIndex        =   6
      Top             =   1440
      Width           =   3705
   End
   Begin VB.TextBox txt�����ʺ� 
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
      Left            =   2235
      TabIndex        =   4
      Top             =   1005
      Width           =   3705
   End
   Begin VB.Label lblҵ������ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҵ������(&U)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   750
      TabIndex        =   22
      Top             =   5085
      Width           =   1425
   End
   Begin VB.Label lbl�������� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������(&H)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   750
      TabIndex        =   11
      Top             =   2820
      Width           =   1425
   End
   Begin VB.Label lbl������Ϣ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&I)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1260
      TabIndex        =   24
      Top             =   5550
      Width           =   915
   End
   Begin VB.Label lblҽ��֤�� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ��֤��(&A)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   750
      TabIndex        =   1
      Top             =   630
      Width           =   1425
   End
   Begin VB.Label lbl���������� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����������(&Y)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   240
      TabIndex        =   19
      Top             =   4635
      Width           =   1935
   End
   Begin VB.Label lbl���������� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����������(&L)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   240
      TabIndex        =   17
      Top             =   4185
      Width           =   1935
   End
   Begin VB.Label lbl��λ���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��λ����(&K)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   750
      TabIndex        =   15
      Top             =   3720
      Width           =   1425
   End
   Begin VB.Label lbl���֤�� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���֤��(&J)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   750
      TabIndex        =   13
      Top             =   3270
      Width           =   1425
   End
   Begin VB.Label lbl�Ա� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�Ա�(&G)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1260
      TabIndex        =   9
      Top             =   2370
      Width           =   915
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&F)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1260
      TabIndex        =   7
      Top             =   1935
      Width           =   915
   End
   Begin VB.Label lbl����Ա 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����Ա(&D)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   990
      TabIndex        =   5
      Top             =   1500
      Width           =   1170
   End
   Begin VB.Label lbl�����ʺ� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�����ʺ�(&S)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   750
      TabIndex        =   3
      Top             =   1065
      Width           =   1425
   End
End
Attribute VB_Name = "frmIdentify����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType As Byte
Private mlng����ID As Long
Private mstrReturn As String

Public Function GetPatient(ByVal bytType As Byte, Optional ByVal lng����ID As Long = 0) As String
    mstrReturn = ""
    mlng����ID = lng����ID
    mbytType = bytType
    Me.Show 1
    
    GetPatient = mstrReturn
End Function

Private Sub chk�޿�����_Click()
    txtҽ��֤��.Enabled = (chk�޿�����.Value = 1)
End Sub

Private Sub cmdOK_Click()
    Dim lngPatient As Long
    Dim strIdentify As String
    Dim strAddition As String
    Dim str����֤�� As String
    Dim rsTemp As New ADODB.Recordset
    
    If Trim(txtҽ��֤��.Text) = "" Then
        MsgBox "��δ������", vbInformation, gstrSysName
        Exit Sub
    End If
    
    'Modified by ZYB 2006-04-12������������ҽ������2006-04-06�·����ļ�Ҫ���޸ģ�����סԺ�������ϴ�������Ϣ
'    If mbytType = 1 Then
        If Val(txt������Ϣ.Tag) = 0 Then
            MsgBox "����ѡ����Ժ���֣�", vbInformation, gstrSysName
            txt������Ϣ.SetFocus
            Exit Sub
        End If
'    End If
    
    '��鲡��״̬
    gstrSQL = "select ����ID,nvl(��ǰ״̬,0) as ״̬,˳���,�Ҷȼ�,����֤�� from �����ʻ� where ����=[1] and ҽ����=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_����, txtҽ��֤��.Text)
    If rsTemp.RecordCount > 0 Then
        If rsTemp!״̬ = 1 Then
            MsgBox "��ǰ������Ժ�������ٴ�ͨ�������֤��", vbInformation, gstrSysName
            Exit Sub
        End If
        str����֤�� = Nvl(rsTemp!����֤��)
    End If
    
    '����������Ϣ
    '�����ַ���
    '�������˵�����Ϣ�������ʽ��
    '0����;1ҽ����;2����;3����;4�Ա�;5��������;6���֤;7��λ����(����);
    '8.���Ĵ���;9.˳���;10��Ա���;11�ʻ����;12��ǰ״̬;13����ID;14��ְ(0,1);15����֤��;16�����;17�Ҷȼ�
    '18�ʻ������ۼ�,19�ʻ�֧���ۼ�,20����ͳ���ۼ�,21ͳ�ﱨ���ۼ�,22סԺ�����ۼ�
    
    '�Ҷȼ�����¼������Ժ�Ĵ�����ȱʡ��һ��Ϊ�գ���Ϊ�ĵ�Ҫ��ÿ��ҽ����Ժ��סԺ�Ų�����ͬ������ֻ��ͨ��������Ժ�����������������Ժû��Ӱ��
    strIdentify = txtҽ��֤��.Text                              '0����
    strIdentify = strIdentify & ";" & txt�����ʺ�.Text          '1ҽ����
    strIdentify = strIdentify & ";"                             '2����
    strIdentify = strIdentify & ";" & txt����.Text              '3����
    strIdentify = strIdentify & ";" & txt�Ա�.Text              '4�Ա�
    strIdentify = strIdentify & ";" & txt��������.Text          '5��������
    strIdentify = strIdentify & ";"                             '6���֤
    strIdentify = strIdentify & ";" & txt��λ����               '7.��λ����(����)
    strAddition = ";0"                                          '8.���Ĵ���
    strAddition = strAddition & ";"                             '9.˳���
    strAddition = strAddition & ";"                             '10��Ա���
    strAddition = strAddition & ";" & Val(txt����������.Text) + Val(txt����������.Text)    '11�ʻ����
    strAddition = strAddition & ";0"                            '12��ǰ״̬
    strAddition = strAddition & ";" & Val(txt������Ϣ.Tag)      '13����ID
    strAddition = strAddition & ";1"                            '14��ְ(1,2,3)
    strAddition = strAddition & ";" & str����֤��               '15����֤��
    strAddition = strAddition & ";"                             '16�����
    strAddition = strAddition & ";" & chk�޿�����.Value         '17�Ҷȼ�
    strAddition = strAddition & ";" & Val(txt����������.Text) + Val(txt����������.Text)     '18�ʻ������ۼ�
    strAddition = strAddition & ";0"                            '19�ʻ�֧���ۼ�
    strAddition = strAddition & ";"                             '20���깤���ܶ�
    strAddition = strAddition & ";" & Val(IC_Data_����.�������)      '21סԺ�����ۼ�

    mlng����ID = BuildPatiInfo(0, strIdentify & strAddition, mlng����ID, TYPE_����)
    '���ظ�ʽ:�м���벡��ID
    If mlng����ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng����ID & strAddition
    Else
        Exit Sub
    End If
    
    IC_Data_����.mstrҵ������ = cboҵ������.ItemData(cboҵ������.ListIndex)
    If Val(IC_Data_����.mstrҵ������) = 0 Then        '��ͨʱ�����⴦����ͨ����Ϊ11����ͨסԺΪ21
        If mbytType = 0 Then
            IC_Data_����.mstrҵ������ = "11"
        Else
            IC_Data_����.mstrҵ������ = "21"
        End If
    End If
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_���� & ",'ҵ������','''" & IC_Data_����.mstrҵ������ & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "����ҵ������")
    
    '�����û�IC������������
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_���� & ",'IC','''" & IC_Data_����.IC������ & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�����û�IC������������")
    '���������������뱾��������
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_���� & ",'�����ʻ����','''" & Val(txt����������.Text) & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�����û�IC������������")
    gstrSQL = "zl_�����ʻ�_������Ϣ(" & mlng����ID & "," & TYPE_���� & ",'�����ʻ����','''" & Val(txt����������.Text) & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "�����û�IC������������")
    
    IC_Data_����.mlng����ID = Val(txt������Ϣ.Tag)
    Unload Me
End Sub

Private Sub cmdRead_Click()
    '������ȷ��ɺ��Զ�����ȷ����
    If Not ReadIC_����(IIf(chk�޿�����.Value = 1, txtҽ��֤��.Text, "")) Then Exit Sub
    
    txtҽ��֤��.Text = IC_Data_����.ҽ��֤��
    txt�����ʺ�.Text = IC_Data_����.�ʺ�
    txt����Ա.Text = IIf(IC_Data_����.����Ա = "1", "��", "��")
    txt����.Text = IC_Data_����.����
    txt�Ա�.Text = IC_Data_����.�Ա�
    txt��������.Text = IC_Data_����.��������
    txt���֤��.Text = IC_Data_����.��ݺ�
    txt��λ����.Text = IC_Data_����.��λ����
    txt����������.Text = Format(IC_Data_����.��ת��� - IC_Data_����.��������ʹ���ۼ�, "#0.00")
    txt����������.Text = Format(IC_Data_����.����ʵ�ʲ��� - IC_Data_����.���ʵ���ʹ���ۼ�, "#0.00")
    txt������Ϣ.SetFocus
End Sub

Private Sub cmd������Ϣ_Click()
    Dim rs���� As ADODB.Recordset
        
    gstrSQL = " Select A.JBBM AS ID,A.JBBZDM AS ����,A.JBMC AS ����,A.PYJM AS ���� " & _
            " From SIM_JBDA A "
    Set rs���� = New ADODB.Recordset
    rs����.Open gstrSQL, gcn����
    If rs����.RecordCount > 0 Then
        If frmListSel.ShowSelect(TYPE_����, rs����, "ID", "ҽ������ѡ��", "��ѡ��ҽ�����֣�") = True Then
            txt������Ϣ.Tag = rs����!ID
            txt������Ϣ.Text = "(" & rs����!���� & ")" & rs����!����
            lbl������Ϣ.Tag = txt������Ϣ.Text '���ڻָ���ʾ
        End If
    End If
    cmdOK.SetFocus
End Sub

Private Sub cmdȡ��_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.cboҵ������.Clear
    Me.cboҵ������.AddItem "��ͨ"           '����Ϊ11��סԺΪ21
    Me.cboҵ������.ItemData(Me.cboҵ������.NewIndex) = 0
    If mbytType = 0 Then
        Me.cboҵ������.AddItem "���ⲡ������"
        Me.cboҵ������.ItemData(Me.cboҵ������.NewIndex) = 12
    ElseIf mbytType = 1 Then
        Me.cboҵ������.AddItem "���ⲡ��סԺ"
        Me.cboҵ������.ItemData(Me.cboҵ������.NewIndex) = 32
        Me.cboҵ������.AddItem "��ͥ����"
        Me.cboҵ������.ItemData(Me.cboҵ������.NewIndex) = 31
    End If
    Me.cboҵ������.ListIndex = 0
    Me.cboҵ������.Enabled = (mbytType = 0 Or mbytType = 1)
End Sub

Private Sub txt������Ϣ_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset, strText As String
    Dim blnReturn As Boolean, lst As ListItem
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt������Ϣ.Text = "" And txt������Ϣ.Tag <> "" Then Exit Sub
    
    On Error GoTo errHandle
    strText = txt������Ϣ.Text
    If InStr(1, strText, "(") <> 0 Then
        If InStr(1, strText, ")") <> 0 Then
            strText = Mid(strText, 2, InStr(1, strText, ")") - 2)
        End If
    End If
        
    gstrSQL = "Select A.JBBM AS ID,A.JBBZDM AS ����,A.JBMC AS ����,A.PYJM AS ����" & _
             "   FROM SIM_JBDA A WHERE 1=1 And (" & _
                zlCommFun.GetLike("A", "JBBZDM", strText) & " or " & zlCommFun.GetLike("A", "JBMC", strText) & " or " & zlCommFun.GetLike("A", "PYJM", strText) & ")"
    If rsTemp.State = 1 Then rsTemp.Close
    rsTemp.Open gstrSQL, gcn����
    If rsTemp.RecordCount = 0 Then
        MsgBox "�����ڸò��֣����������룡", vbInformation, gstrSysName
        txt������Ϣ.Text = lbl������Ϣ.Tag
        zlControl.TxtSelAll txt������Ϣ
        Exit Sub
    Else
        '����ѡ����
        If rsTemp.RecordCount > 1 Then
            '�����ֶδ���3�ģ���ʹֻ��һ����¼�ѸöԻ�����ʾ�������Ա����û��õ��������Ϣ
            blnReturn = frmListSel.ShowSelect(TYPE_����, rsTemp, "ID", "ҽ������ѡ��", "��ѡ��ҽ�����֣�")
        Else
            blnReturn = True
        End If
    End If
    
    If blnReturn = False Then
        '��¼����û�п�ѡ�������
        txt������Ϣ.Text = lbl������Ϣ.Tag
        zlControl.TxtSelAll txt������Ϣ
        Exit Sub
    Else
        '�϶����м�¼����
        txt������Ϣ.Tag = rsTemp!ID
        txt������Ϣ.Text = "(" & rsTemp!���� & ")" & rsTemp!����
        lbl������Ϣ.Tag = txt������Ϣ.Text '���ڻָ���ʾ
    End If
    
    cmdOK.SetFocus
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
