VERSION 5.00
Begin VB.Form frmDegreeCard 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������Ϣ"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8895
   Icon            =   "frmDegreeCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame fra��Ժ��Ϣ 
      ForeColor       =   &H00C00000&
      Height          =   1335
      Left            =   75
      TabIndex        =   80
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
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   780
      TabIndex        =   77
      Top             =   5460
      Width           =   1100
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&X)"
      Height          =   350
      Left            =   7050
      TabIndex        =   76
      Top             =   5460
      Width           =   1100
   End
   Begin VB.Frame fra������Ϣ 
      ForeColor       =   &H00C00000&
      Height          =   3345
      Left            =   75
      TabIndex        =   78
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
      TabIndex        =   79
      Top             =   4530
      Width           =   8745
      Begin VB.TextBox txt������� 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox txtԤ����� 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   240
         Width           =   1155
      End
      Begin VB.TextBox txt������ 
         Alignment       =   1  'Right Justify
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   7380
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   75
         Top             =   240
         Width           =   1080
      End
      Begin VB.TextBox txt������ 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   5265
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   73
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "δ�����"
         Height          =   180
         Left            =   2370
         TabIndex        =   70
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ԥ�����"
         Height          =   180
         Left            =   375
         TabIndex        =   68
         Top             =   300
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   6765
         TabIndex        =   74
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   4695
         TabIndex        =   72
         Top             =   300
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
        MsgBox "������ȷ��ȡ������Ϣ,����ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    Me.Show 1, frmMain
            
    ShowInfo = True
    
End Function

Private Function ReadCard(ByVal lng����ID As Long) As Boolean

    '���ܣ���ȡָ��������Ϣ,����ʾ�ڽ�����
    
    Dim rsTmp As New ADODB.Recordset
    Dim strsql As String
    
    On Error GoTo ErrHand
        
    strsql = "Select A.*," & _
            "DECODE(A.��ǰ����id,NULL,��������,(SELECT ���� FROM ���ű� WHERE ID=A.��ǰ����id)) AS ���� " & _
            "From ������Ϣ A Where A.����ID=[1] "
    
    Set rsTmp = zldatabase.OpenSQLRecord(strsql, Me.Caption, lng����ID)
    
    If rsTmp.EOF Then Exit Function
    If rsTmp.RecordCount <> 1 Then Exit Function
    
    'סԺ��Ϣ
    txt����ID.Text = lng����ID
    txt����.Text = rsTmp!����
    
    '������Ϣ
    txt�Ա�.Text = zlCommFun.Nvl(rsTmp("�Ա�"))
    txt����.Text = zlCommFun.Nvl(rsTmp("����"))
    txtҽ�Ƹ���.Text = zlCommFun.Nvl(rsTmp("ҽ�Ƹ��ʽ"))
    txt����.Text = zlCommFun.Nvl(rsTmp("����"))
    txt����.Text = zlCommFun.Nvl(rsTmp("����"))
    txtѧ��.Text = zlCommFun.Nvl(rsTmp("ѧ��"))
    txt����״��.Text = zlCommFun.Nvl(rsTmp("����״��"))
    txtְҵ.Text = zlCommFun.Nvl(rsTmp("ְҵ"))
    txt���.Text = zlCommFun.Nvl(rsTmp("���"))
    txt��������.Text = Format(zlCommFun.Nvl(rsTmp("��������")), "yyyy-mm-dd")
    txt���֤��.Text = zlCommFun.Nvl(rsTmp("���֤��"))
    txt�����ص�.Text = zlCommFun.Nvl(rsTmp("�����ص�"))
    txt��ͥ��ַ.Text = zlCommFun.Nvl(rsTmp("��ͥ��ַ"))
    txt��ͥ�绰.Text = zlCommFun.Nvl(rsTmp("��ͥ�绰"))
    txt�����ʱ�.Text = zlCommFun.Nvl(rsTmp("��ͥ��ַ�ʱ�"))
    txt��ϵ������.Text = zlCommFun.Nvl(rsTmp("��ϵ������"))
    txt��ϵ�˹�ϵ.Text = zlCommFun.Nvl(rsTmp("��ϵ�˹�ϵ"))
    txt��ϵ�˵�ַ.Text = zlCommFun.Nvl(rsTmp("��ϵ�˵�ַ"))
    txt��ϵ�˵绰.Text = zlCommFun.Nvl(rsTmp("��ϵ�˵绰"))
    txt������λ.Text = zlCommFun.Nvl(rsTmp("������λ"))
    txt��λ�绰.Text = zlCommFun.Nvl(rsTmp("��λ�绰"))
    txt��λ�ʱ�.Text = zlCommFun.Nvl(rsTmp("��λ�ʱ�"))
    txt��λ������.Text = zlCommFun.Nvl(rsTmp("��λ������"))
    txt��λ�ʺ�.Text = zlCommFun.Nvl(rsTmp("��λ�ʺ�"))
    txt����.Text = zlCommFun.Nvl(rsTmp("����"))
    
    '������Ϣ
    txt������.Text = zlCommFun.Nvl(rsTmp("������"))
    txt������.Text = zlCommFun.Nvl(rsTmp("������"))

    If zlCommFun.Nvl(rsTmp("��ǰ����id"), 0) > 0 Then
        txtסԺ��.Text = zlCommFun.Nvl(rsTmp("סԺ��"))
        txt����.Text = zlCommFun.Nvl(rsTmp("��ǰ����"))
        txt��Ժʱ��.Text = Format(zlCommFun.Nvl(rsTmp("��Ժʱ��")), "yyyy-MM-dd HH:mm")
        txt��Ժʱ��.Text = Format(zlCommFun.Nvl(rsTmp("��Ժʱ��")), "yyyy-MM-dd HH:mm")
    
        strsql = "Select B.�ѱ� AS סԺ�ѱ� From ������Ϣ A,������ҳ B " & _
                "Where A.����ID=B.����ID And A.��ҳid=B.��ҳID And A.����ID=[1]"
        Set rsTmp = zldatabase.OpenSQLRecord(strsql, Me.Caption, lng����ID)
        If rsTmp.BOF = False Then txt�ѱ�.Text = zlCommFun.Nvl(rsTmp("סԺ�ѱ�"))
        
    Else
        lblסԺ��.Caption = "�����"
        txtסԺ��.Text = zlCommFun.Nvl(rsTmp("�����"))
    End If
        
        
    strsql = "Select " & gConst_�������_���� & "  From ������� a Where ����=1 And ����ID= [1] "
    Set rsTmp = zldatabase.OpenSQLRecord(strsql, Me.Caption, lng����ID)
    
    If Not rsTmp.EOF Then
        txt�������.Text = Format(zlCommFun.Nvl(rsTmp("�������")), "0.00")
        txtԤ�����.Text = Format(zlCommFun.Nvl(rsTmp("Ԥ�����")), "0.00")
    End If
    
    ReadCard = True
    
    Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

