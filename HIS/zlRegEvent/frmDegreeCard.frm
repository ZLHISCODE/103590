VERSION 5.00
Begin VB.Form frmDegreeCard 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������Ϣ��Ƭ"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8850
   Icon            =   "frmDegreeCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   8850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picPati 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1260
      Left            =   75
      ScaleHeight     =   1260
      ScaleWidth      =   8730
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   75
      Width           =   8730
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   1170
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   1170
      End
      Begin VB.TextBox txt����ID 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   90
         Width           =   1170
      End
      Begin VB.TextBox txt�Ա� 
         Height          =   300
         Left            =   2985
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   480
         Width           =   1170
      End
      Begin VB.TextBox txt�ѱ� 
         Height          =   300
         Left            =   7350
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   480
         Width           =   1170
      End
      Begin VB.TextBox txt����� 
         Height          =   300
         Left            =   2985
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   90
         Width           =   1170
      End
      Begin VB.TextBox txtסԺ�� 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   90
         Width           =   1170
      End
      Begin VB.TextBox txt���� 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   7350
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   90
         Width           =   1170
      End
      Begin VB.TextBox txt���� 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   870
         Width           =   1170
      End
      Begin VB.TextBox txt���� 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   2985
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   870
         Width           =   1170
      End
      Begin VB.TextBox txt��Ժʱ�� 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   870
         Width           =   1170
      End
      Begin VB.TextBox txt��Ժʱ�� 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   7350
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   870
         Width           =   1170
      End
      Begin VB.Label lbl����ID 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ID"
         ForeColor       =   &H80000007&
         Height          =   180
         Left            =   525
         TabIndex        =   73
         Top             =   150
         Width           =   540
      End
      Begin VB.Label lbl�ѱ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ѱ�"
         Height          =   180
         Left            =   6945
         TabIndex        =   72
         Top             =   555
         Width           =   360
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   4845
         TabIndex        =   71
         Top             =   540
         Width           =   360
      End
      Begin VB.Label lbl�Ա� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         Height          =   180
         Left            =   2595
         TabIndex        =   70
         Top             =   540
         Width           =   360
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   720
         TabIndex        =   69
         Top             =   540
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����"
         Height          =   180
         Left            =   2415
         TabIndex        =   68
         Top             =   150
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "סԺ��"
         Height          =   180
         Left            =   4665
         TabIndex        =   67
         Top             =   150
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   6945
         TabIndex        =   66
         Top             =   150
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰ����"
         Height          =   180
         Left            =   360
         TabIndex        =   65
         Top             =   930
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   2595
         TabIndex        =   64
         Top             =   930
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժʱ��"
         Height          =   180
         Left            =   4485
         TabIndex        =   63
         Top             =   930
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Ժʱ��"
         Height          =   180
         Left            =   6600
         TabIndex        =   62
         Top             =   945
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&X)"
      Height          =   350
      Left            =   6960
      TabIndex        =   36
      Top             =   5880
      Width           =   1100
   End
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4275
      Left            =   75
      ScaleHeight     =   4275
      ScaleWidth      =   8730
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   1380
      Width           =   8730
      Begin VB.TextBox txt���ڵ�ַ 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   76
         Top             =   3960
         Width           =   3015
      End
      Begin VB.TextBox txt�����ʱ� 
         Height          =   300
         Left            =   5280
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   75
         Top             =   3960
         Width           =   1170
      End
      Begin VB.TextBox txtҽ�Ƹ��� 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   105
         Width           =   1170
      End
      Begin VB.TextBox txt���֤�� 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   18
         TabIndex        =   20
         Top             =   885
         Width           =   3015
      End
      Begin VB.TextBox txt��ͥ�绰 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   24
         Top             =   1665
         Width           =   2000
      End
      Begin VB.TextBox txt��λ�ʺ� 
         Height          =   300
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   3225
         Width           =   3255
      End
      Begin VB.TextBox txt��λ������ 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   3225
         Width           =   3015
      End
      Begin VB.TextBox txt��λ�ʱ� 
         Height          =   300
         Left            =   5280
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   31
         Top             =   2835
         Width           =   1170
      End
      Begin VB.TextBox txt��λ�绰 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   30
         Top             =   2835
         Width           =   2000
      End
      Begin VB.TextBox txt������λ 
         Height          =   300
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   2445
         Width           =   3255
      End
      Begin VB.TextBox txt��ϵ�˵绰 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   28
         Top             =   2445
         Width           =   2000
      End
      Begin VB.TextBox txt��ϵ�˵�ַ 
         Height          =   300
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   2055
         Width           =   3255
      End
      Begin VB.TextBox txt��ϵ������ 
         Height          =   300
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1665
         Width           =   1170
      End
      Begin VB.TextBox txt��ͥ�ʱ� 
         Height          =   300
         Left            =   5280
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   23
         Top             =   1275
         Width           =   1170
      End
      Begin VB.TextBox txt��ͥ��ַ 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1275
         Width           =   3015
      End
      Begin VB.TextBox txt�����ص� 
         Height          =   300
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   885
         Width           =   3255
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   105
         Width           =   1170
      End
      Begin VB.TextBox txtѧ�� 
         Height          =   300
         Left            =   7350
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   105
         Width           =   1170
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   2985
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   105
         Width           =   1170
      End
      Begin VB.TextBox txt����״�� 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   495
         Width           =   1170
      End
      Begin VB.TextBox txtְҵ 
         Height          =   300
         Left            =   2985
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   495
         Width           =   1170
      End
      Begin VB.TextBox txt��� 
         Height          =   300
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   495
         Width           =   1170
      End
      Begin VB.TextBox txt��ϵ�˹�ϵ 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   2055
         Width           =   2000
      End
      Begin VB.TextBox txt�������� 
         Height          =   300
         Left            =   7350
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   495
         Width           =   1170
      End
      Begin VB.TextBox txt������ 
         Height          =   300
         Left            =   1140
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   34
         Top             =   3615
         Width           =   1980
      End
      Begin VB.TextBox txt������ 
         Height          =   300
         Left            =   5280
         Locked          =   -1  'True
         MaxLength       =   12
         TabIndex        =   35
         Top             =   3615
         Width           =   1170
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ڵ�ַ"
         Height          =   180
         Left            =   360
         TabIndex        =   78
         Top             =   4020
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ʱ�"
         Height          =   180
         Left            =   4440
         TabIndex        =   77
         Top             =   4005
         Width           =   720
      End
      Begin VB.Label lblҽ�Ƹ��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ҽ�Ƹ���"
         Height          =   180
         Left            =   360
         TabIndex        =   74
         Top             =   165
         Width           =   720
      End
      Begin VB.Label lbl��λ�ʺ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�ʺ�"
         Height          =   180
         Left            =   4485
         TabIndex        =   60
         Top             =   3285
         Width           =   720
      End
      Begin VB.Label lbl��λ������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ������"
         Height          =   180
         Left            =   180
         TabIndex        =   59
         Top             =   3285
         Width           =   900
      End
      Begin VB.Label lbl��λ�ʱ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�ʱ�"
         Height          =   180
         Left            =   4485
         TabIndex        =   58
         Top             =   2895
         Width           =   720
      End
      Begin VB.Label lbl��λ�绰 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��λ�绰"
         Height          =   180
         Left            =   360
         TabIndex        =   57
         Top             =   2895
         Width           =   720
      End
      Begin VB.Label lbl������λ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������λ"
         Height          =   180
         Left            =   4485
         TabIndex        =   56
         Top             =   2505
         Width           =   720
      End
      Begin VB.Label lbl��ϵ�˵绰 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ�˵绰"
         Height          =   180
         Left            =   180
         TabIndex        =   55
         Top             =   2505
         Width           =   900
      End
      Begin VB.Label lbl��ϵ�˵�ַ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ�˵�ַ"
         Height          =   180
         Left            =   4305
         TabIndex        =   54
         Top             =   2115
         Width           =   900
      End
      Begin VB.Label lbl��ϵ�˹�ϵ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ�˹�ϵ"
         Height          =   180
         Left            =   180
         TabIndex        =   53
         Top             =   2115
         Width           =   900
      End
      Begin VB.Label lbl��ϵ������ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ϵ������"
         Height          =   180
         Left            =   4305
         TabIndex        =   52
         Top             =   1725
         Width           =   900
      End
      Begin VB.Label lbl�����ʱ� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͥ�ʱ�"
         Height          =   180
         Left            =   4485
         TabIndex        =   51
         Top             =   1335
         Width           =   720
      End
      Begin VB.Label lbl��ͥ�绰 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͥ�绰"
         Height          =   180
         Left            =   360
         TabIndex        =   50
         Top             =   1725
         Width           =   720
      End
      Begin VB.Label lbl��ͥ��ַ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͥ��ַ"
         Height          =   180
         Left            =   360
         TabIndex        =   49
         Top             =   1335
         Width           =   720
      End
      Begin VB.Label lvl����״�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����״��"
         Height          =   180
         Left            =   360
         TabIndex        =   48
         Top             =   555
         Width           =   720
      End
      Begin VB.Label lblѧ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ѧ��"
         Height          =   180
         Left            =   6945
         TabIndex        =   47
         Top             =   165
         Width           =   360
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   2595
         TabIndex        =   46
         Top             =   165
         Width           =   360
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   4845
         TabIndex        =   45
         Top             =   165
         Width           =   360
      End
      Begin VB.Label lblְҵ 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ְҵ"
         Height          =   180
         Left            =   2595
         TabIndex        =   44
         Top             =   555
         Width           =   360
      End
      Begin VB.Label lbl��� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���"
         Height          =   180
         Left            =   4845
         TabIndex        =   43
         Top             =   555
         Width           =   360
      End
      Begin VB.Label lbl���֤�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���֤��"
         Height          =   180
         Left            =   360
         TabIndex        =   42
         Top             =   945
         Width           =   720
      End
      Begin VB.Label lbl�����ص� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�����ص�"
         Height          =   180
         Left            =   4485
         TabIndex        =   41
         Top             =   945
         Width           =   720
      End
      Begin VB.Label lbl�������� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         Height          =   180
         Left            =   6585
         TabIndex        =   40
         Top             =   555
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   540
         TabIndex        =   39
         Top             =   3675
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������"
         Height          =   180
         Left            =   4665
         TabIndex        =   38
         Top             =   3675
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
Public mlng����ID As Long 'Ҫ�޸Ļ�鿴�Ĳ���ID
Private mblnUnload As Boolean

Private Function ReadCard(lngID As Long) As Boolean
'���ܣ���ȡָ��������Ϣ,����ʾ�ڽ�����
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
        
    strSQL = "Select A.*,B.���� as ��ǰ����,C.���� as ��ǰ����" & _
        " From ������Ϣ A,���ű� B,���ű� C" & _
        " Where A.��ǰ����ID=B.ID(+) And A.��ǰ����ID=C.ID(+) And A.����ID=[1]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, lngID)
    
    If rsTmp.EOF Then Exit Function
    If rsTmp.RecordCount <> 1 Then Exit Function
    
    txt����ID.Text = lngID
    txt����.Text = rsTmp!����
    
    txt�����.Text = IIf(IsNull(rsTmp!�����), "", rsTmp!�����)
    txtסԺ��.Text = IIf(IsNull(rsTmp!סԺ��), "", rsTmp!סԺ��)
    txt����.Text = IIf(IsNull(rsTmp!��ǰ����), "", rsTmp!��ǰ����)
    txt����.Text = IIf(IsNull(rsTmp!��ǰ����), "", rsTmp!��ǰ����)
    txt����.Text = IIf(IsNull(rsTmp!��ǰ����), "", rsTmp!��ǰ����)
    txt��Ժʱ��.Text = Format(IIf(IsNull(rsTmp!��Ժʱ��), "", rsTmp!��Ժʱ��), "yyyy-MM-dd")
    txt��Ժʱ��.Text = Format(IIf(IsNull(rsTmp!��Ժʱ��), "", rsTmp!��Ժʱ��), "yyyy-MM-dd")
    
    txt�Ա�.Text = IIf(IsNull(rsTmp!�Ա�), "", rsTmp!�Ա�)
    txt����.Text = IIf(IsNull(rsTmp!����), "", rsTmp!����)
    txt�ѱ�.Text = IIf(IsNull(rsTmp!�ѱ�), "", rsTmp!�ѱ�)
    txtҽ�Ƹ���.Text = IIf(IsNull(rsTmp!ҽ�Ƹ��ʽ), "", rsTmp!ҽ�Ƹ��ʽ)
    txt����.Text = IIf(IsNull(rsTmp!����), "", rsTmp!����)
    txt����.Text = IIf(IsNull(rsTmp!����), "", rsTmp!����)
    txtѧ��.Text = IIf(IsNull(rsTmp!ѧ��), "", rsTmp!ѧ��)
    txt����״��.Text = IIf(IsNull(rsTmp!����״��), "", rsTmp!����״��)
    txtְҵ.Text = IIf(IsNull(rsTmp!ְҵ), "", rsTmp!ְҵ)
    txt���.Text = IIf(IsNull(rsTmp!���), "", rsTmp!���)
    txt��������.Text = Format(IIf(IsNull(rsTmp!��������), "", rsTmp!��������), "yyyy-MM-dd")
    txt���֤��.Text = IIf(IsNull(rsTmp!���֤��), "", rsTmp!���֤��)
    txt�����ص�.Text = IIf(IsNull(rsTmp!�����ص�), "", rsTmp!�����ص�)
    txt��ͥ��ַ.Text = IIf(IsNull(rsTmp!��ͥ��ַ), "", rsTmp!��ͥ��ַ)
    txt��ͥ�绰.Text = IIf(IsNull(rsTmp!��ͥ�绰), "", rsTmp!��ͥ�绰)
    txt�����ʱ�.Text = IIf(IsNull(rsTmp!��ͥ��ַ�ʱ�), "", rsTmp!��ͥ��ַ�ʱ�)
    txt��ϵ������.Text = IIf(IsNull(rsTmp!��ϵ������), "", rsTmp!��ϵ������)
    txt��ϵ�˹�ϵ.Text = IIf(IsNull(rsTmp!��ϵ�˹�ϵ), "", rsTmp!��ϵ�˹�ϵ)
    txt��ϵ�˵�ַ.Text = IIf(IsNull(rsTmp!��ϵ�˵�ַ), "", rsTmp!��ϵ�˵�ַ)
    txt��ϵ�˵绰.Text = IIf(IsNull(rsTmp!��ϵ�˵绰), "", rsTmp!��ϵ�˵绰)
    txt������λ.Text = IIf(IsNull(rsTmp!������λ), "", rsTmp!������λ)
    txt��λ�绰.Text = IIf(IsNull(rsTmp!��λ�绰), "", rsTmp!��λ�绰)
    txt��λ�ʱ�.Text = IIf(IsNull(rsTmp!��λ�ʱ�), "", rsTmp!��λ�ʱ�)
    txt��λ������.Text = IIf(IsNull(rsTmp!��λ������), "", rsTmp!��λ������)
    txt��λ�ʺ�.Text = IIf(IsNull(rsTmp!��λ�ʺ�), "", rsTmp!��λ�ʺ�)
    txt������.Text = IIf(IsNull(rsTmp!������), "", rsTmp!������)
    txt������.Text = IIf(IsNull(rsTmp!������), "", rsTmp!������)
    txt���ڵ�ַ.Text = IIf(IsNull(rsTmp!���ڵ�ַ), "", rsTmp!���ڵ�ַ)
    txt��ͥ�ʱ�.Text = IIf(IsNull(rsTmp!��ͥ��ַ�ʱ�), "", rsTmp!��ͥ��ַ�ʱ�)
    '74428�����ϴ���2014-7-8������������ʾ��ɫ����
    Call SetPatiColor(txt����, Nvl(rsTmp!��������), IIf(IsNull(rsTmp!����), Me.ForeColor, vbRed))
    
    ReadCard = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub



Private Sub Form_Activate()
    If mblnUnload Then Unload Me: Exit Sub
    cmdExit.SetFocus
End Sub

Private Sub Form_Load()
    mblnUnload = False
    If Not ReadCard(mlng����ID) Then
        MsgBox "������ȷ��ȡ������Ϣ,����ϵͳ����Ա��ϵ��", vbExclamation, gstrSysName
        mblnUnload = True
    End If
    
    zlcontrol.PicShowFlat picInfo, -1, , taCenterAlign
    zlcontrol.PicShowFlat picPati, -1, , taCenterAlign
End Sub
