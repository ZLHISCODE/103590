VERSION 5.00
Begin VB.Form frmDegreeCard 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������Ϣ"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   Icon            =   "frmDegreeCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fra��Ժ��Ϣ 
      ForeColor       =   &H00C00000&
      Height          =   1335
      Left            =   75
      TabIndex        =   77
      Top             =   30
      Width           =   8730
      Begin VB.TextBox txt���� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3090
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   885
         Width           =   1170
      End
      Begin VB.TextBox txt��λ�ȼ� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   885
         Width           =   1170
      End
      Begin VB.TextBox txt��Ժʱ�� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7335
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   885
         Width           =   1110
      End
      Begin VB.TextBox txt��Ժʱ�� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5250
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   885
         Width           =   1065
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7335
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   225
         Width           =   1110
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5235
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   1065
      End
      Begin VB.TextBox txtסԺ�� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3090
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   1170
      End
      Begin VB.TextBox txt�ѱ� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7335
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   555
         Width           =   1110
      End
      Begin VB.TextBox txt�Ա� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3090
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   555
         Width           =   1170
      End
      Begin VB.TextBox txt����ID 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   225
         Width           =   1170
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   555
         Width           =   1170
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5250
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   555
         Width           =   1065
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   2700
         TabIndex        =   18
         Top             =   945
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�ȼ�"
         Height          =   180
         Left            =   345
         TabIndex        =   16
         Top             =   945
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժʱ��"
         Height          =   180
         Left            =   6585
         TabIndex        =   22
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժʱ��"
         Height          =   180
         Left            =   4470
         TabIndex        =   20
         Top             =   945
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   6930
         TabIndex        =   6
         Top             =   285
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   4845
         TabIndex        =   4
         Top             =   285
         Width           =   360
      End
      Begin VB.Label lblסԺ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��"
         Height          =   180
         Left            =   2505
         TabIndex        =   2
         Top             =   285
         Width           =   540
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   705
         TabIndex        =   8
         Top             =   615
         Width           =   360
      End
      Begin VB.Label lbl�Ա� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   2700
         TabIndex        =   10
         Top             =   615
         Width           =   360
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   4830
         TabIndex        =   12
         Top             =   615
         Width           =   360
      End
      Begin VB.Label lbl�ѱ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ѱ�"
         Height          =   180
         Left            =   6930
         TabIndex        =   14
         Top             =   630
         Width           =   360
      End
      Begin VB.Label lbl����ID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ID"
         ForeColor       =   &H80000007&
         Height          =   180
         Left            =   510
         TabIndex        =   0
         Top             =   285
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&X)"
      Height          =   350
      Left            =   7620
      TabIndex        =   74
      Top             =   5370
      Width           =   1100
   End
   Begin VB.Frame fra������Ϣ 
      ForeColor       =   &H00C00000&
      Height          =   3345
      Left            =   75
      TabIndex        =   75
      Top             =   1380
      Width           =   8745
      Begin VB.TextBox txtҽ�Ƹ��� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txt�������� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7335
         Locked          =   -1  'True
         TabIndex        =   39
         Top             =   570
         Width           =   1140
      End
      Begin VB.TextBox txt��ϵ�˹�ϵ 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   53
         Top             =   1890
         Width           =   2000
      End
      Begin VB.TextBox txt��� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   37
         Top             =   570
         Width           =   1170
      End
      Begin VB.TextBox txtְҵ 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3105
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   570
         Width           =   1170
      End
      Begin VB.TextBox txt����״�� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   570
         Width           =   1170
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3105
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txtѧ�� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7335
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   240
         Width           =   1140
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   240
         Width           =   1170
      End
      Begin VB.TextBox txt�����ص� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   43
         Top             =   900
         Width           =   3225
      End
      Begin VB.TextBox txt��ͥ��ַ 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   1230
         Width           =   2790
      End
      Begin VB.TextBox txt�����ʱ� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   47
         Top             =   1230
         Width           =   1170
      End
      Begin VB.TextBox txt��ϵ������ 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   51
         Top             =   1560
         Width           =   1170
      End
      Begin VB.TextBox txt��ϵ�˵�ַ 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   55
         Top             =   1890
         Width           =   3225
      End
      Begin VB.TextBox txt��ϵ�˵绰 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   57
         Top             =   2220
         Width           =   2000
      End
      Begin VB.TextBox txt������λ 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   59
         Top             =   2220
         Width           =   3225
      End
      Begin VB.TextBox txt��λ�绰 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   61
         Top             =   2550
         Width           =   2000
      End
      Begin VB.TextBox txt��λ�ʱ� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   63
         Top             =   2550
         Width           =   1170
      End
      Begin VB.TextBox txt��λ������ 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   2880
         Width           =   3135
      End
      Begin VB.TextBox txt��λ�ʺ� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   2880
         Width           =   3225
      End
      Begin VB.TextBox txt��ͥ�绰 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   49
         Top             =   1560
         Width           =   2000
      End
      Begin VB.TextBox txt���֤�� 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1125
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   41
         Top             =   900
         Width           =   3150
      End
      Begin VB.Label lblҽ�Ƹ��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ�Ƹ���"
         Height          =   180
         Left            =   345
         TabIndex        =   24
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lbl�������� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   6570
         TabIndex        =   38
         Top             =   630
         Width           =   720
      End
      Begin VB.Label lbl�����ص� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ص�"
         Height          =   180
         Left            =   4470
         TabIndex        =   42
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lbl���֤�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��"
         Height          =   180
         Left            =   345
         TabIndex        =   40
         Top             =   960
         Width           =   720
      End
      Begin VB.Label lbl��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         Height          =   180
         Left            =   4830
         TabIndex        =   36
         Top             =   630
         Width           =   360
      End
      Begin VB.Label lblְҵ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ְҵ"
         Height          =   180
         Left            =   2685
         TabIndex        =   34
         Top             =   630
         Width           =   360
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   4830
         TabIndex        =   28
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   2685
         TabIndex        =   26
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lblѧ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ѧ��"
         Height          =   180
         Left            =   6930
         TabIndex        =   30
         Top             =   300
         Width           =   360
      End
      Begin VB.Label lvl����״�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����״��"
         Height          =   180
         Left            =   345
         TabIndex        =   32
         Top             =   630
         Width           =   720
      End
      Begin VB.Label lbl��ͥ��ַ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͥ��ַ"
         Height          =   180
         Left            =   345
         TabIndex        =   44
         Top             =   1290
         Width           =   720
      End
      Begin VB.Label lbl��ͥ�绰 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͥ�绰"
         Height          =   180
         Left            =   345
         TabIndex        =   48
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label lbl�����ʱ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͥ��ַ�ʱ�"
         Height          =   180
         Left            =   4110
         TabIndex        =   46
         Top             =   1290
         Width           =   1080
      End
      Begin VB.Label lbl��ϵ������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ������"
         Height          =   180
         Left            =   4290
         TabIndex        =   50
         Top             =   1635
         Width           =   900
      End
      Begin VB.Label lbl��ϵ�˹�ϵ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ�˹�ϵ"
         Height          =   180
         Left            =   165
         TabIndex        =   52
         Top             =   1950
         Width           =   900
      End
      Begin VB.Label lbl��ϵ�˵�ַ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ�˵�ַ"
         Height          =   180
         Left            =   4290
         TabIndex        =   54
         Top             =   1950
         Width           =   900
      End
      Begin VB.Label lbl��ϵ�˵绰 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ�˵绰"
         Height          =   180
         Left            =   165
         TabIndex        =   56
         Top             =   2280
         Width           =   900
      End
      Begin VB.Label lbl������λ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������λ"
         Height          =   180
         Left            =   4470
         TabIndex        =   58
         Top             =   2280
         Width           =   720
      End
      Begin VB.Label lbl��λ�绰 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�绰"
         Height          =   180
         Left            =   345
         TabIndex        =   60
         Top             =   2610
         Width           =   720
      End
      Begin VB.Label lbl��λ�ʱ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�ʱ�"
         Height          =   180
         Left            =   4470
         TabIndex        =   62
         Top             =   2610
         Width           =   720
      End
      Begin VB.Label lbl��λ������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ������"
         Height          =   180
         Left            =   165
         TabIndex        =   64
         Top             =   2940
         Width           =   900
      End
      Begin VB.Label lbl��λ�ʺ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�ʺ�"
         Height          =   180
         Left            =   4470
         TabIndex        =   66
         Top             =   2940
         Width           =   720
      End
   End
   Begin VB.Frame fra������Ϣ 
      ForeColor       =   &H00C00000&
      Height          =   705
      Left            =   75
      TabIndex        =   76
      Top             =   4530
      Width           =   8745
      Begin VB.TextBox txt������� 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   270
         Width           =   1155
      End
      Begin VB.TextBox txtԤ����� 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   270
         Width           =   1155
      End
      Begin VB.TextBox txt������ 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   73
         Top             =   270
         Width           =   1080
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "δ�����"
         Height          =   180
         Left            =   2370
         TabIndex        =   70
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ԥ�����"
         Height          =   180
         Left            =   375
         TabIndex        =   68
         Top             =   330
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   4635
         TabIndex        =   72
         Top             =   330
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmDegreeCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������

Public Function ShowInfo(ByVal frmMain As Form, ByVal lngKey As Long) As Boolean
    If ReadCard(lngKey) = False Then
        MsgBox "������ȷ��ȡ������Ϣ,����ϵͳ����Ա��ϵ��", vbInformation, gUserInfo.Name
        Exit Function
    End If
    Me.Show 1, frmMain
    ShowInfo = True
End Function

Private Function ReadCard(ByVal lng����ID As Long) As Boolean
          '���ܣ���ȡָ��������Ϣ,����ʾ�ڽ�����
          Dim rsTmp As New ADODB.Recordset
          Dim strSQL As String
          Dim lng��ҳID As Long

1         On Error GoTo ReadCard_Error

2         strSQL = "Select A.*," & _
                  "DECODE(A.��ǰ����id,NULL,��������,(SELECT ���� FROM ���ű� WHERE ID=A.��ǰ����id)) AS ���� " & _
                  "From ������Ϣ A Where A.����ID=[1] "
          
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption, lng����ID)
                      
4         If rsTmp.EOF Then Exit Function
5         If rsTmp.RecordCount <> 1 Then Exit Function
          
          'סԺ��Ϣ
6         lng��ҳID = Val(NVL(rsTmp("��ҳID")))
7         txt����ID.Text = lng����ID
8         txt����.Text = rsTmp!����
          
          '������Ϣ
9         txt�Ա�.Text = NVL(rsTmp("�Ա�"))
10        txt����.Text = NVL(rsTmp("����"))
11        txtҽ�Ƹ���.Text = NVL(rsTmp("ҽ�Ƹ��ʽ"))
12        txt����.Text = NVL(rsTmp("����"))
13        txt����.Text = NVL(rsTmp("����"))
14        txtѧ��.Text = NVL(rsTmp("ѧ��"))
15        txt����״��.Text = NVL(rsTmp("����״��"))
16        txtְҵ.Text = NVL(rsTmp("ְҵ"))
17        txt���.Text = NVL(rsTmp("���"))
18        txt��������.Text = Format(NVL(rsTmp("��������")), "yyyy-mm-dd")
19        txt���֤��.Text = NVL(rsTmp("���֤��"))
20        txt�����ص�.Text = NVL(rsTmp("�����ص�"))
21        txt��ͥ��ַ.Text = NVL(rsTmp("��ͥ��ַ"))
22        txt��ͥ�绰.Text = NVL(rsTmp("��ͥ�绰"))
23        txt�����ʱ�.Text = NVL(rsTmp("��ͥ��ַ�ʱ�"))
24        txt��ϵ������.Text = NVL(rsTmp("��ϵ������"))
25        txt��ϵ�˹�ϵ.Text = NVL(rsTmp("��ϵ�˹�ϵ"))
26        txt��ϵ�˵�ַ.Text = NVL(rsTmp("��ϵ�˵�ַ"))
27        txt��ϵ�˵绰.Text = NVL(rsTmp("��ϵ�˵绰"))
28        txt������λ.Text = NVL(rsTmp("������λ"))
29        txt��λ�绰.Text = NVL(rsTmp("��λ�绰"))
30        txt��λ�ʱ�.Text = NVL(rsTmp("��λ�ʱ�"))
31        txt��λ������.Text = NVL(rsTmp("��λ������"))
32        txt��λ�ʺ�.Text = NVL(rsTmp("��λ�ʺ�"))
33        txt����.Text = NVL(rsTmp("����"))
          
34        If NVL(rsTmp("��ǰ����id"), 0) > 0 Then
35            txtסԺ��.Text = NVL(rsTmp("סԺ��"))
36            txt����.Text = NVL(rsTmp("��ǰ����"))
37            txt��Ժʱ��.Text = Format(NVL(rsTmp("��Ժʱ��")), "yyyy-MM-dd HH:mm")
38            txt��Ժʱ��.Text = Format(NVL(rsTmp("��Ժʱ��")), "yyyy-MM-dd HH:mm")
          
39            strSQL = "Select B.�ѱ� AS סԺ�ѱ� From ������Ϣ A,������ҳ B " & _
                      "Where A.����ID=B.����ID And A.��ҳid=B.��ҳID And A.����ID=[1]"
40            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption, lng����ID)
41            If rsTmp.BOF = False Then txt�ѱ�.Text = NVL(rsTmp("סԺ�ѱ�"))
42        Else
43            lblסԺ��.Caption = "�����"
44            txtסԺ��.Text = NVL(rsTmp("�����"))
45        End If
          
46        strSQL = "Select a.����ID,a.����,a.Ԥ�����,a.�������   From ������� a Where ����=1 And ����ID= [1] "
47        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption, lng����ID)
          
48        If Not rsTmp.EOF Then
49            txt�������.Text = Format(NVL(rsTmp("�������")), "0.00")
50            txtԤ�����.Text = Format(NVL(rsTmp("Ԥ�����")), "0.00")
51        End If
          
          '������Ϣ
52        strSQL = "Select Zl_Patientsurety([1],[2]) As ������ From Dual"
53        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption, lng����ID, lng��ҳID)
54        If Not rsTmp.EOF Then
55            txt������.Text = Format(NVL(rsTmp("������")), "0.00")
56        End If
          
57        ReadCard = True


58        Exit Function
ReadCard_Error:
59        Call WriteErrLog("zl9LisInsideComm", "frmDegreeCard", "ִ��(ReadCard)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
60        Err.Clear
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then PressKey vbKeyTab
End Sub

