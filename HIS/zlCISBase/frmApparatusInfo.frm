VERSION 5.00
Begin VB.Form frmApparatusInfo 
   Appearance      =   0  'Flat
   BorderStyle     =   0  'None
   Caption         =   "����������Ϣ"
   ClientHeight    =   4290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8430
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   8430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '����ȱʡ
   Begin VB.CheckBox chk���� 
      Alignment       =   1  'Right Justify
      Caption         =   "����ʱָ������"
      Height          =   195
      Left            =   4140
      TabIndex        =   50
      ToolTipText     =   "���ϱ�ʾ��ʹ�ü�ʦ����վ��[��������]����ʱ��Ҫָ���̺źͱ��ţ�������ָ����"
      Top             =   135
      Width           =   1575
   End
   Begin VB.CheckBox chkLog 
      Caption         =   "�����ʿ�ͼ"
      Height          =   210
      Left            =   7125
      TabIndex        =   49
      ToolTipText     =   "һ��PCR�����Ų��ö����ʿ�ͼ"
      Top             =   3885
      Width           =   1200
   End
   Begin VB.Frame fraø�� 
      Caption         =   "ø��������"
      Height          =   1110
      Left            =   120
      TabIndex        =   38
      Top             =   2130
      Width           =   8160
      Begin VB.TextBox txtø�� 
         Height          =   300
         Index           =   4
         Left            =   5385
         MaxLength       =   40
         TabIndex        =   47
         Top             =   690
         Width           =   2610
      End
      Begin VB.TextBox txtø�� 
         Height          =   300
         Index           =   3
         Left            =   1335
         MaxLength       =   40
         TabIndex        =   45
         Top             =   690
         Width           =   2610
      End
      Begin VB.TextBox txtø�� 
         Height          =   300
         Index           =   2
         Left            =   7395
         MaxLength       =   40
         TabIndex        =   43
         Top             =   300
         Width           =   600
      End
      Begin VB.TextBox txtø�� 
         Height          =   300
         Index           =   1
         Left            =   4470
         MaxLength       =   40
         TabIndex        =   41
         Top             =   300
         Width           =   1800
      End
      Begin VB.TextBox txtø�� 
         Height          =   300
         Index           =   0
         Left            =   810
         MaxLength       =   40
         TabIndex        =   39
         Top             =   300
         Width           =   2500
      End
      Begin VB.Label lblø�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�հ���ʽ(&X)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   4
         Left            =   4275
         TabIndex        =   48
         Top             =   735
         Width           =   990
      End
      Begin VB.Label lblø�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���巽ʽ(&F)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   46
         Top             =   735
         Width           =   990
      End
      Begin VB.Label lblø�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���ʱ��(&J)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   6345
         TabIndex        =   44
         Top             =   345
         Width           =   990
      End
      Begin VB.Label lblø�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���Ƶ��(&L)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   3390
         TabIndex        =   42
         Top             =   345
         Width           =   990
      End
      Begin VB.Label lblø�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����(&B)"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   40
         Top             =   345
         Width           =   630
      End
   End
   Begin VB.ComboBox cbo������� 
      Height          =   300
      Left            =   3870
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   900
      Width           =   1860
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   -45
      TabIndex        =   36
      Top             =   3315
      Width           =   8535
   End
   Begin VB.ComboBox cboУ׼����Դ 
      Height          =   300
      ItemData        =   "frmApparatusInfo.frx":0000
      Left            =   4770
      List            =   "frmApparatusInfo.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   3855
      Width           =   2325
   End
   Begin VB.ComboBox cbo�Լ���Դ 
      Height          =   300
      ItemData        =   "frmApparatusInfo.frx":0004
      Left            =   1380
      List            =   "frmApparatusInfo.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   3855
      Width           =   1950
   End
   Begin VB.TextBox txt����QC�� 
      Height          =   300
      Left            =   6930
      MaxLength       =   8
      TabIndex        =   31
      Top             =   3420
      Width           =   1290
   End
   Begin VB.TextBox txt�ʿ�ˮƽ 
      Height          =   300
      Left            =   4785
      MaxLength       =   1
      TabIndex        =   29
      Top             =   3435
      Width           =   405
   End
   Begin VB.ComboBox cbo���ڵ�λ 
      Height          =   300
      ItemData        =   "frmApparatusInfo.frx":0008
      Left            =   2100
      List            =   "frmApparatusInfo.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   27
      Top             =   3435
      Width           =   645
   End
   Begin VB.TextBox txt�ʿ����� 
      Height          =   300
      Left            =   1665
      MaxLength       =   5
      TabIndex        =   26
      Top             =   3435
      Width           =   420
   End
   Begin VB.ComboBox cboУ��λ 
      Height          =   300
      ItemData        =   "frmApparatusInfo.frx":001E
      Left            =   6990
      List            =   "frmApparatusInfo.frx":0020
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   1710
      Width           =   1290
   End
   Begin VB.ComboBox cboֹͣλ 
      Height          =   300
      ItemData        =   "frmApparatusInfo.frx":0022
      Left            =   6990
      List            =   "frmApparatusInfo.frx":0024
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   1290
      Width           =   1290
   End
   Begin VB.ComboBox cbo����λ 
      Height          =   300
      ItemData        =   "frmApparatusInfo.frx":0026
      Left            =   6990
      List            =   "frmApparatusInfo.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   885
      Width           =   1290
   End
   Begin VB.ComboBox cbo������ 
      Height          =   300
      ItemData        =   "frmApparatusInfo.frx":002A
      Left            =   6990
      List            =   "frmApparatusInfo.frx":002C
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   480
      Width           =   1290
   End
   Begin VB.ComboBox cboͨѶ�� 
      Height          =   300
      ItemData        =   "frmApparatusInfo.frx":002E
      Left            =   6990
      List            =   "frmApparatusInfo.frx":0030
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   75
      Width           =   1290
   End
   Begin VB.TextBox txtͨѶ������ 
      Height          =   300
      Left            =   1380
      MaxLength       =   40
      TabIndex        =   14
      Top             =   1755
      Width           =   4335
   End
   Begin VB.TextBox txt���Ӽ���� 
      Height          =   300
      Left            =   1380
      MaxLength       =   40
      TabIndex        =   10
      Top             =   1335
      Width           =   1695
   End
   Begin VB.ComboBox cboʹ��С�� 
      Height          =   300
      ItemData        =   "frmApparatusInfo.frx":0032
      Left            =   3870
      List            =   "frmApparatusInfo.frx":0034
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   1335
      Width           =   1860
   End
   Begin VB.ComboBox cbo�������� 
      Height          =   300
      ItemData        =   "frmApparatusInfo.frx":0036
      Left            =   855
      List            =   "frmApparatusInfo.frx":0038
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   930
      Width           =   2235
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   855
      MaxLength       =   10
      TabIndex        =   1
      Top             =   120
      Width           =   1320
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   855
      MaxLength       =   20
      TabIndex        =   3
      Top             =   525
      Width           =   2610
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   4350
      MaxLength       =   10
      TabIndex        =   5
      Top             =   525
      Width           =   1365
   End
   Begin VB.Label lbl������� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���(&G)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3225
      TabIndex        =   37
      Top             =   990
      Width           =   630
   End
   Begin VB.Label lblУ׼����Դ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ĭ��У׼����Դ"
      Height          =   180
      Left            =   3435
      TabIndex        =   34
      Top             =   3915
      Width           =   1260
   End
   Begin VB.Label lbl�Լ���Դ 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ĭ���Լ���Դ"
      Height          =   180
      Left            =   180
      TabIndex        =   32
      Top             =   3915
      Width           =   1080
   End
   Begin VB.Label lbl����QC�� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����QC��"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6075
      TabIndex        =   30
      Top             =   3480
      Width           =   720
   End
   Begin VB.Label lbl�ʿ�ˮƽ 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ÿ�����     ��ˮƽ"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4020
      TabIndex        =   28
      Top             =   3495
      Width           =   1710
   End
   Begin VB.Label lbl�ʿ����� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�ʿ�Ҫ��: ����ÿ             ����һ���ʿ�,"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      TabIndex        =   25
      Top             =   3495
      Width           =   3780
   End
   Begin VB.Label lblУ��λ 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "У��λ(&5)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6135
      TabIndex        =   23
      Top             =   1770
      Width           =   810
   End
   Begin VB.Label lblֹͣλ 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ֹͣλ(&4)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6135
      TabIndex        =   21
      Top             =   1350
      Width           =   810
   End
   Begin VB.Label lbl����λ 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����λ(&3)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6135
      TabIndex        =   19
      Top             =   945
      Width           =   810
   End
   Begin VB.Label lbl������ 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������(&2)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6135
      TabIndex        =   17
      Top             =   540
      Width           =   810
   End
   Begin VB.Label lblͨѶ�� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ͨѶ��(&1)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   6135
      TabIndex        =   15
      Top             =   135
      Width           =   810
   End
   Begin VB.Label lblͨѶ������ 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ͨѶ������(&P)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      TabIndex        =   13
      Top             =   1800
      Width           =   1170
   End
   Begin VB.Label lbl���Ӽ���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���Ӽ����(&T)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      TabIndex        =   9
      Top             =   1395
      Width           =   1170
   End
   Begin VB.Label lblʹ��С�� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ʹ��(&U)"
      Height          =   180
      Left            =   3225
      TabIndex        =   11
      Top             =   1395
      Width           =   630
   End
   Begin VB.Label lbl�������� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&K)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      TabIndex        =   6
      Top             =   990
      Width           =   630
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   630
   End
   Begin VB.Label lbl��Ŀ���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      TabIndex        =   2
      Top             =   585
      Width           =   630
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����(&S)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3660
      TabIndex        =   4
      Top             =   585
      Width           =   630
   End
End
Attribute VB_Name = "frmApparatusInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLngAptId As Long          '��ǰ��ʾ������id

Dim objNode As Node
Dim strTemp As String, aryTemp() As String
Dim lngCount As Long

'--------------------------------------------
'����Ϊ���幫������
'--------------------------------------------
Public Function zlRefresh(lngAptId As Long) As Boolean
    '���ܣ�������Ŀidˢ�µ�ǰ��ʾ����
    Dim rsTemp As New ADODB.Recordset, intIndex As Integer
    mLngAptId = lngAptId
    
    '�����ǰ��Ŀ����ʾ
    Me.txt����.Text = "": Me.txt����.Text = "": Me.txt����.Text = ""
    Me.cbo��������.ListIndex = -1: Me.cbo�������.ListIndex = -1 'Me.chk΢����.Value = 0
    Me.txt���Ӽ����.Text = "": Me.cboʹ��С��.ListIndex = -1: Me.txtͨѶ������.Text = ""
    Me.cboͨѶ��.ListIndex = -1: Me.cbo������.ListIndex = -1
    Me.cbo����λ.ListIndex = -1: Me.cboֹͣλ.ListIndex = -1: Me.cboУ��λ.ListIndex = -1
    Me.txt�ʿ�����.Text = "": Me.cbo���ڵ�λ.ListIndex = -1:  Me.txt�ʿ�ˮƽ.Text = 0
    Me.txt����QC��.Text = "": Me.cbo�Լ���Դ.ListIndex = -1: Me.cboУ׼����Դ.ListIndex = -1
    For intIndex = 0 To 4
        Me.txtø��(intIndex).Text = ""
    Next
    If lngAptId = 0 Then zlRefresh = True: Exit Function
    
    '��ȡָ����Ŀ����Ϣ
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select A.����, A.����, A.����, A.��������, A.΢����, A.���Ӽ����, A.ͨѶ������, A.ͨѶ�˿�, A.������, A.����, A.ֹͣλ," & vbNewLine & _
            "       A.У��λ, A.ʹ��С��id, D.���� As ʹ��С��, A.�ʿ�����, A.���ڵ�λ, A.�ʿ�ˮƽ��, A.Qc��, A.�Լ���Դ," & vbNewLine & _
            "       A.У׼����Դ,A.����,A.���Ƶ��,A.���ʱ��,A.���巽ʽ,A.�հ���ʽ,A.�����ʿ�ͼ,A.����ʱָ������ " & vbNewLine & _
            "From �������� A, ���ű� D" & vbNewLine & _
            "Where A.ʹ��С��id = D.ID(+) And A.ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngAptId)
    With rsTemp
        If .RecordCount > 0 Then
            Me.txt����.Text = "" & !����: Me.txt����.Text = "" & !����: Me.txt����.Text = "" & !����
            For lngCount = 0 To Me.cbo��������.ListCount - 1
                If Mid(Me.cbo��������.List(lngCount), InStr(1, Me.cbo��������.List(lngCount), "-") + 1) = "" & !�������� Then
                    Me.cbo��������.ListIndex = lngCount: Exit For
                End If
            Next
            'Me.chk΢����.Value = IIf(Val("" & !΢����) = 1, 1, 0)
            For intIndex = 0 To cbo�������.ListCount - 1
                If Val(cbo�������.List(intIndex)) = Val("" & !΢����) Then
                    cbo�������.ListIndex = intIndex
                    Exit For
                End If
            Next
            
            Me.txtø��(0).Text = "" & !����
            Me.txtø��(1).Text = "" & !���Ƶ��
            Me.txtø��(2).Text = "" & !���ʱ��
            Me.txtø��(3).Text = "" & !���巽ʽ
            Me.txtø��(4).Text = "" & !�հ���ʽ
            Me.chkLog.Value = Val("" & !�����ʿ�ͼ)
            Me.chk����.Value = Val("" & !����ʱָ������)
            
            Me.txt���Ӽ����.Text = "" & !���Ӽ����
            For lngCount = 0 To Me.cboʹ��С��.ListCount - 1
                If Me.cboʹ��С��.ItemData(lngCount) = Val("" & !ʹ��С��id) Then
                    Me.cboʹ��С��.ListIndex = lngCount: Exit For
                End If
            Next
            Me.txtͨѶ������.Text = "" & !ͨѶ������
                        
            For lngCount = 0 To Me.cboͨѶ��.ListCount - 1
                If Me.cboͨѶ��.List(lngCount) = "" & !ͨѶ�˿� Then Me.cboͨѶ��.ListIndex = lngCount: Exit For
            Next
            For lngCount = 0 To Me.cbo������.ListCount - 1
                If Me.cbo������.List(lngCount) = "" & !������ Then Me.cbo������.ListIndex = lngCount: Exit For
            Next
            For lngCount = 0 To Me.cbo����λ.ListCount - 1
                If Me.cbo����λ.List(lngCount) = "" & !���� Then Me.cbo����λ.ListIndex = lngCount: Exit For
            Next
            For lngCount = 0 To Me.cboֹͣλ.ListCount - 1
                If Me.cboֹͣλ.List(lngCount) = "" & !ֹͣλ Then Me.cboֹͣλ.ListIndex = lngCount: Exit For
            Next
            For lngCount = 0 To Me.cboУ��λ.ListCount - 1
                If InStr(1, Me.cboУ��λ.List(lngCount), "-") = 0 Then
                    If Me.cboУ��λ.List(lngCount) = "" & !У��λ Then Me.cboУ��λ.ListIndex = lngCount: Exit For
                Else
                    If Left(Me.cboУ��λ.List(lngCount), 1) = "" & !У��λ Then Me.cboУ��λ.ListIndex = lngCount: Exit For
                End If
            Next
            
            Me.txt�ʿ�����.Text = Val("" & !�ʿ�����)
            If "" & !���ڵ�λ <> "��" Then
                Me.cbo���ڵ�λ.ListIndex = 0
            Else
                Me.cbo���ڵ�λ.ListIndex = 1
            End If
            Me.txt�ʿ�ˮƽ.Text = Val("" & !�ʿ�ˮƽ��)
            Me.txt����QC��.Text = "" & !Qc��
            For lngCount = 0 To Me.cbo�Լ���Դ.ListCount - 1
                If Me.cbo�Լ���Դ.List(lngCount) = "" & !�Լ���Դ Then Me.cbo�Լ���Դ.ListIndex = lngCount: Exit For
            Next
            For lngCount = 0 To Me.cboУ׼����Դ.ListCount - 1
                If Me.cboУ׼����Դ.List(lngCount) = "" & !У׼����Դ Then Me.cboУ׼����Դ.ListIndex = lngCount: Exit For
            Next
        End If
    End With
        
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditStart(blnAdd As Boolean, lngAptId As Long) As Boolean
    '���ܣ���ʼ��Ŀ�༭
    '������ blnAdd-�Ƿ����ӣ�����Ϊ�޸�
    '       lngAptId-���ӵĲ�����Ŀ������ָ���༭����Ŀ
    Dim rsTemp As New ADODB.Recordset, i As Integer
    If Me.cbo��������.ListCount = 0 Then
        MsgBox "�������ֵ��г�ʼ�����������͡���", vbInformation, gstrSysName
        zlEditStart = False: Exit Function
    End If
    
    If blnAdd Then
        Err = 0: On Error GoTo ErrHand
        gstrSql = "Select Nvl(Max(����),'000') As ���� From ��������"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "zlEditStart")
        With rsTemp
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.ProductName, Me.Caption, gstrSql)
            
'            Call SQLTest
            'Me.txt����.Text = Right(String(10, "0") & Val(!����) + 1, Len(!����))
            Me.txt����.Text = zlCommFun.IncStr(IIf("" & !���� = "", "000", "" & !����))
        End With
        
        '���������Ĭ��ֵ
        Me.txt����.Text = "": Me.txt����.Text = ""
        Me.cbo��������.ListIndex = 0: Me.cbo�������.ListIndex = 0 'Me.chk΢����.Value = 0
        Me.txt���Ӽ����.Text = "": Me.cboʹ��С��.ListIndex = -1: Me.txtͨѶ������.Text = ""
        Me.cboͨѶ��.ListIndex = 0: Me.cbo������.ListIndex = 5
        Me.cbo����λ.ListIndex = 4: Me.cboֹͣλ.ListIndex = 0: Me.cboУ��λ.ListIndex = 3
        Me.txt�ʿ�����.Text = 1: Me.cbo���ڵ�λ.ListIndex = 0: Me.txt�ʿ�ˮƽ.Text = 1
        Me.txt����QC��.Text = "": Me.cbo�Լ���Դ.ListIndex = -1: Me.cboУ׼����Դ.ListIndex = -1
        For i = 0 To 4
            Me.txtø��(i).Text = "": Me.txtø��(i).Tag = ""
        Next
    Else
        If Val(Me.txt�ʿ�����.Text) = 0 Then Me.txt�ʿ�����.Text = 1
        If Me.cbo���ڵ�λ.ListIndex = -1 Then Me.cbo���ڵ�λ.ListIndex = 0
        If Val(Me.txt�ʿ�ˮƽ.Text) = 0 Then Me.txt�ʿ�ˮƽ.Text = 1
        For i = 0 To 4
            Me.txtø��(i).Tag = Me.txtø��(i).Text
        Next
    End If

    mLngAptId = lngAptId
    Me.Enabled = True: Me.Tag = IIf(blnAdd, "����", "�޸�")
    Me.BackColor = RGB(250, 250, 250): Me.fraø��.BackColor = Me.BackColor
    Me.txt����.SetFocus
    zlEditStart = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditStart = False: Exit Function
End Function

Public Sub zlEditCancel()
    '���ܣ��������ڽ��еı༭
    Me.Enabled = False: Me.Tag = ""
    Me.BackColor = Me.fraLine.BackColor: Me.fraø��.BackColor = Me.BackColor
    Call Me.zlRefresh(mLngAptId)
End Sub

Public Function zlEditSave() As Long
    '���ܣ��������ڽ��еı༭,���������ڱ༭��Ŀid,����ʧ�ܷ���0
    Dim lngNewId As Long
    
    'һ�����Լ��
    If Trim(Me.txt����.Text) = "" Then
        MsgBox "��������룡", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > 3 Then
        MsgBox "����ĳ��ȳ��������3���ַ�����", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Trim(Me.txt����.Text) = "" Then
        MsgBox "���������ƣ�", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Me.cbo��������.ListIndex = -1 Then
        MsgBox "��ѡ���������ͣ�", vbInformation, gstrSysName
        Me.cbo��������.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt����.Text), vbFromUnicode)) > Me.txt����.MaxLength Then
        MsgBox "���Ƴ��������" & Me.txt����.MaxLength & "���ַ���ȳ����֣���", vbInformation, gstrSysName
        Me.txt����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt���Ӽ����.Text), vbFromUnicode)) > Me.txt���Ӽ����.MaxLength Then
        MsgBox "���Ӽ�������������" & Me.txt���Ӽ����.MaxLength & "���ַ���ȳ����֣���", vbInformation, gstrSysName
        Me.txt���Ӽ����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txtͨѶ������.Text), vbFromUnicode)) > Me.txtͨѶ������.MaxLength Then
        MsgBox "ͨѶ���������������" & Me.txtͨѶ������.MaxLength & "���ַ���ȳ����֣���", vbInformation, gstrSysName
        Me.txtͨѶ������.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Val(Me.txt�ʿ�����.Text) <= 0 Or Val(Me.txt�ʿ�����.Text) > 365 Then
        MsgBox "�����ú����ʿ����ڣ�", vbInformation, gstrSysName
        Me.txt�ʿ�����.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Me.cbo���ڵ�λ.ListIndex = -1 Then
        MsgBox "�������ʿ����ڵ�λ��", vbInformation, gstrSysName
        Me.cbo���ڵ�λ.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Val(Me.txt�ʿ�ˮƽ.Text) <= 0 Or Val(Me.txt�ʿ�ˮƽ.Text) > 9 Then
        MsgBox "�����ú����ʿ�ˮƽ����1��9����", vbInformation, gstrSysName
        Me.txt�ʿ�ˮƽ.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt����QC��.Text), vbFromUnicode)) > Me.txt����QC��.MaxLength Then
        MsgBox "����QC�볬�������" & Me.txt����QC��.MaxLength & "���ַ�����", vbInformation, gstrSysName
        Me.txt����QC��.SetFocus: zlEditSave = 0: Exit Function
    End If
    
    '���ݱ��������֯
    If Me.Tag = "����" Then
        lngNewId = zlDatabase.GetNextId("��������")
    Else
        lngNewId = mLngAptId
    End If

    gstrSql = lngNewId & ",'" & Replace(Trim(Me.txt����.Text), "'", "") & "','" & Replace(Trim(Me.txt����.Text), "'", "") & "','" & Replace(Trim(Me.txt����.Text), "'", "") & "'"
    gstrSql = gstrSql & ",'" & Replace(Trim(Me.txt���Ӽ����.Text), "'", "") & "','" & Replace(Trim(Me.txtͨѶ������.Text), "'", "") & "','" & Me.cboͨѶ��.Text & "'"
    gstrSql = gstrSql & "," & Val(Me.cbo������.Text) & "," & Val(Me.cbo����λ.Text) & "," & Val(Me.cboֹͣλ.Text)
    If InStr(1, Me.cboУ��λ.Text, "-") = 0 Then
        gstrSql = gstrSql & ",'" & Me.cboУ��λ.Text & "'"
    Else
        gstrSql = gstrSql & ",'" & Left(Me.cboУ��λ.Text, 1) & "'"
    End If
    gstrSql = gstrSql & ",'" & Mid(Me.cbo��������.Text, InStr(1, Me.cbo��������.Text, "-") + 1) & "'"
    gstrSql = gstrSql & "," & Val(Me.cbo�������.Text)
    If Me.cboʹ��С��.ListIndex = -1 Then
        gstrSql = gstrSql & ",Null"
    Else
        gstrSql = gstrSql & "," & Me.cboʹ��С��.ItemData(Me.cboʹ��С��.ListIndex)
    End If
    gstrSql = gstrSql & "," & Val(Me.txt�ʿ�����.Text) & ",'" & Trim(Me.cbo���ڵ�λ.Text) & "'," & Val(Me.txt�ʿ�ˮƽ.Text)
    gstrSql = gstrSql & ",'" & Replace(Trim(Me.txt����QC��.Text), "'", "") & "','" & Trim(Me.cbo�Լ���Դ.Text) & "','" & Trim(Me.cboУ׼����Դ.Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txtø��(0).Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txtø��(1).Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txtø��(2).Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txtø��(3).Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txtø��(4).Text) & "'"
    gstrSql = gstrSql & "," & Me.chkLog.Value
    gstrSql = gstrSql & "," & Me.chk����.Value
    If Me.Tag = "����" Then
        gstrSql = "Zl_��������_Insert(" & gstrSql & ")"
    Else
        gstrSql = "Zl_��������_Update(" & gstrSql & ")"
    End If
    
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    If Me.Tag = "����" Then mLngAptId = lngNewId
    Me.Enabled = False: Me.Tag = ""
    Me.BackColor = Me.fraLine.BackColor: Me.fraø��.BackColor = Me.BackColor
    zlEditSave = mLngAptId: Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function

'--------------------------------------------
'����Ϊ����ؼ���Ӧ�¼�
'--------------------------------------------

Private Sub cbo������_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboʹ��С��_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cboʹ��С��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo�Լ���Դ_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo�Լ���Դ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo����λ_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo����λ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboֹͣλ_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cboֹͣλ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboͨѶ��_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cboͨѶ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboУ��λ_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cboУ��λ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboУ׼����Դ_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cboУ׼����Դ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo�������_Click()
    Dim i As Integer
    If cbo�������.ListIndex = 2 Then
        For i = 0 To 4
            Me.txtø��(i).Enabled = True: Me.txtø��(i).Text = Me.txtø��(i).Tag
        Next
    Else
        For i = 0 To 4
            Me.txtø��(i).Enabled = False:  If Me.txtø��(i).Text <> "" Then Me.txtø��(i).Tag = Me.txtø��(i).Text: Me.txtø��(i).Text = ""
        Next
    End If
End Sub

Private Sub cbo�������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo��������_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo��������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo���ڵ�λ_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo���ڵ�λ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

'Private Sub chk΢����_GotFocus()
'    Call zlCommFun.OpenIme(False)
'End Sub
'
'Private Sub chk΢����_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
'End Sub
    
Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    Dim aryTemp() As String
    Err = 0: On Error GoTo ErrHand
    '�ֶγ�������
    gstrSql = "Select A.����, A.����, A.����, A.��������, A.΢����, A.���Ӽ����, A.ͨѶ������, A.ͨѶ�˿�, A.������, A.����, A.ֹͣλ," & vbNewLine & _
            "       A.У��λ, A.ʹ��С��id, D.���� As ʹ��С��, A.QC��" & vbNewLine & _
            "From �������� A, ���ű� D" & vbNewLine & _
            "Where A.ʹ��С��id = D.ID(+) And A.ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mLngAptId)
    With rsTemp
'        Me.txt����.MaxLength = .Fields("����").DefinedSize
        Me.txt����.MaxLength = .Fields("����").DefinedSize
        Me.txt����.MaxLength = .Fields("����").DefinedSize
        Me.txt���Ӽ����.MaxLength = .Fields("���Ӽ����").DefinedSize
        Me.txtͨѶ������.MaxLength = .Fields("ͨѶ������").DefinedSize
        Me.txt����QC��.MaxLength = .Fields("QC��").DefinedSize
    End With
    
    '���Ƽ�������
    gstrSql = "Select ����,���� From ���Ƽ�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mLngAptId)
    With rsTemp
        Me.cbo��������.Clear
        Do While Not .EOF
            Me.cbo��������.AddItem Trim(!����) & "-" & Trim(!����)
            .MoveNext
        Loop
        If Me.cbo��������.ListCount > 0 Then Me.cbo��������.ListIndex = 0
    End With
    '�����������
    aryTemp = Split("0-��ͨ����;1-΢������;2-ø����", ";")
    For lngCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo�������.AddItem aryTemp(lngCount)
    Next
    Me.cbo�������.ListIndex = 0
    
    '����ִ�в���
    gstrSql = "Select Id, ����, ����" & vbNewLine & _
            "From ���ű� d, ��������˵�� p" & vbNewLine & _
            "Where d.Id = p.����id And p.�������� = '����' And Instr(',1,2,3,', ','||p.�������||',') > 0 And" & vbNewLine & _
            "           (To_Char(d.����ʱ��, 'YYYY-MM-DD') = '3000-01-01' Or d.����ʱ�� Is Null)"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Do While Not .EOF
            Me.cboʹ��С��.AddItem !���� & "-" & !����
            Me.cboʹ��С��.ItemData(Me.cboʹ��С��.NewIndex) = !ID
            .MoveNext
        Loop
    End With
    
    '�Լ���Դ��У׼����Դ
    gstrSql = "Select ���� From �ʿ��Լ���Դ Order By ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Do While Not .EOF
            Me.cbo�Լ���Դ.AddItem "" & !����
            Me.cboУ׼����Դ.AddItem "" & !����
            .MoveNext
        Loop
    End With
    
    '�����̶�����װ��
    For lngCount = 1 To 50: Me.cboͨѶ��.AddItem "COM" & lngCount: Next
    Me.cboͨѶ��.ListIndex = 0

    aryTemp = Split("110|300|600|1200|2400|4800|9600|14400|19200|28800|38400|56000|128000|256000", "|")
    For lngCount = LBound(aryTemp) To UBound(aryTemp): Me.cbo������.AddItem aryTemp(lngCount): Next
    Me.cbo������.ListIndex = 0

    aryTemp = Split("4|5|6|7|8", "|")
    For lngCount = LBound(aryTemp) To UBound(aryTemp): Me.cbo����λ.AddItem aryTemp(lngCount): Next
    Me.cbo����λ.ListIndex = 0

    aryTemp = Split("1|1.5|2", "|")
    For lngCount = LBound(aryTemp) To UBound(aryTemp): Me.cboֹͣλ.AddItem aryTemp(lngCount): Next
    Me.cboֹͣλ.ListIndex = 0

    aryTemp = Split("E-ż��|M-���|N-ȱʡ|None|O-����|S-�ո�", "|")
    For lngCount = LBound(aryTemp) To UBound(aryTemp): Me.cboУ��λ.AddItem aryTemp(lngCount): Next
    Me.cboУ��λ.ListIndex = 0
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
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
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 1000
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt���Ӽ����_GotFocus()
    Me.txt���Ӽ����.SelStart = 0: Me.txt���Ӽ����.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt���Ӽ����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Me.txt����.Text = MoveSpecialChar(Me.txt����.Text)
        Me.txt����.Text = zlStr.GetCodeByORCL(Me.txt����.Text, False, Me.txt����.MaxLength)
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End If
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt����_LostFocus()
    Me.txt����.Text = zlStr.GetCodeByORCL(Me.txt����.Text, False, Me.txt����.MaxLength)
End Sub

Private Sub txtͨѶ������_GotFocus()
    Me.txtͨѶ������.SelStart = 0: Me.txtͨѶ������.SelLength = 1000
End Sub

Private Sub txtͨѶ������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt����QC��_GotFocus()
    Me.txt����QC��.SelStart = 0: Me.txt����QC��.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt����QC��_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt�ʿ�ˮƽ_GotFocus()
    Me.txt�ʿ�ˮƽ.SelStart = 0: Me.txt�ʿ�ˮƽ.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt�ʿ�ˮƽ_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii > Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt�ʿ�����_GotFocus()
    Me.txt�ʿ�����.SelStart = 0: Me.txt�ʿ�����.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt�ʿ�����_KeyPress(KeyAscii As Integer)
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
