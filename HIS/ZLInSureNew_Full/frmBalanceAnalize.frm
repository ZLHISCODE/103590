VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmBalanceAnalize 
   Caption         =   "ҽ������ָ��ͳ�Ʊ�"
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10215
   Icon            =   "frmBalanceAnalize.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6960
   ScaleWidth      =   10215
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdͳ�� 
      Caption         =   "ͳ��"
      Height          =   350
      Left            =   5100
      TabIndex        =   4
      Top             =   90
      Width           =   1100
   End
   Begin MSComCtl2.DTPicker dtp��ʼ���� 
      Height          =   315
      Left            =   1110
      TabIndex        =   1
      Top             =   90
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   95092739
      CurrentDate     =   38785
   End
   Begin VB.CommandButton cmd�˳� 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&X)"
      Default         =   -1  'True
      Height          =   350
      Left            =   8730
      TabIndex        =   6
      Top             =   6480
      Width           =   1100
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "���&EXCEL"
      Height          =   350
      Left            =   150
      TabIndex        =   7
      Top             =   6480
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshList 
      Height          =   5865
      Left            =   0
      TabIndex        =   5
      Top             =   510
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   10345
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   250
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      MouseIcon       =   "frmBalanceAnalize.frx":0ECA
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComCtl2.DTPicker dtp�������� 
      Height          =   315
      Left            =   3630
      TabIndex        =   3
      Top             =   90
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   95092739
      CurrentDate     =   38785
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
      Left            =   2760
      TabIndex        =   2
      Top             =   150
      Width           =   720
   End
   Begin VB.Label lbl��ʼ���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ʼ����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      TabIndex        =   0
      Top             =   157
      Width           =   720
   End
End
Attribute VB_Name = "frmBalanceAnalize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintInsure As Integer
'ͳ��ָ��ʱ�䷶Χ�ڣ����в��˵ķ������

Public Sub ShowME(ByVal intinsure As Integer)
    mintInsure = intinsure
    Me.Show 1
End Sub

Private Sub cmdͳ��_Click()
    Dim rsTmp As New ADODB.Recordset
    Call InitTable
    
    'ͳ�Ʋ��˽�������
    gstrSQL = "" & _
             " Select B.����ID,A.����ID,B.����,B.סԺ��, " & _
             "        To_char(C.��Ժ����,'yyyy-MM-dd') As ��Ժ����, " & _
             "        to_char(C.��Ժ����,'yyyy-MM-dd') As ��Ժ����,C.סԺ���� As סԺ����, " & _
             "        trim(to_char(A.�����ܶ�,'9000990.00')) AS �����ܶ�,trim(to_char(A.ҩƷ��,'9000990.00')) AS ҩƷ��,trim(to_char(A.��Ŀ¼��ҩƷ��,'9000990.00')) AS ��Ŀ¼��ҩƷ��, " & _
             "        trim(to_char(A.ҩƷ��/A.�����ܶ�*100,'9990.00'))||'%' As ҩƷ����, " & _
             "        trim(to_char(A.��Ŀ¼��ҩƷ��/A.�����ܶ�*100,'9990.00'))||'%' As ��Ŀ¼��ҩƷ���� " & _
             " From ( " & _
             "      Select A.����ID,A.��ҳID,B.ID As ����ID, " & _
             "             sum(Nvl(A.ʵ�ս��,0)) �����ܶ�,"
    gstrSQL = gstrSQL & "Sum(DECODE(A.�շ����,'5',Nvl(A.ʵ�ս��,0),'6',Nvl(A.ʵ�ս��,0),'7',Nvl(A.ʵ�ս��,0),0)) As ҩƷ��, " & _
             "             Sum(DECODE(Nvl(A.ͳ����,0),0,DECODE(A.�շ����,'5',Nvl(A.ʵ�ս��,0),'6',Nvl(A.ʵ�ս��,0),'7',Nvl(A.ʵ�ս��,0),0),0)) As ��Ŀ¼��ҩƷ�� " & _
             "      From סԺ���ü�¼ A,���˽��ʼ�¼ B " & _
             "      Where A.����ID=B.ID And Nvl(A.ʵ�ս��,0)<>0 And Nvl(A.��¼״̬,0)<>0 And Nvl(A.���ӱ�־,0)<>9 " & _
             "      And B.�շ�ʱ�� Between to_date('" & Format(dtp��ʼ����.Value, "yyyy-MM-dd") & "','yyyy-MM-dd') And to_date('" & Format(dtp��������.Value, "yyyy-MM-dd") & "','yyyy-MM-dd') " & _
             "      Having sum(Nvl(A.ʵ�ս��,0))<>0 " & _
             "      Group By A.����ID,A.��ҳID,B.Id " & _
             " ) A,������Ϣ B,������ҳ C,�����ʻ� D " & _
             " Where A.����ID =B.����ID And B.����ID=C.����ID And A.��ҳID=C.��ҳID And A.����ID=D.����ID And D.����=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "ͳ�Ʋ��˽�������", mintInsure)
    If rsTmp.RecordCount = 0 Then Exit Sub
    
    With mshList
        Set .DataSource = rsTmp
        .ColWidth(0) = 1000
        .ColWidth(1) = 0
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
        .ColWidth(4) = 1200
        .ColWidth(5) = 1200
        .ColWidth(6) = 800
        .ColWidth(7) = 1200
        .ColWidth(8) = 1200
        .ColWidth(9) = 1200
        .ColWidth(10) = 1000
        .ColWidth(11) = 1000
        .ColAlignment(3) = 1
        .ColAlignment(4) = 1
        .ColAlignment(5) = 1
        .ColAlignment(7) = 7
        .ColAlignment(8) = 7
        .ColAlignment(9) = 7
        .ColAlignment(10) = 7
        .ColAlignment(11) = 7
    End With
End Sub

Private Sub Form_Load()
    Me.dtp��������.Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    Me.dtp��ʼ����.Value = Format(DateAdd("m", -1, zlDatabase.Currentdate), "yyyy-MM-dd")
    
    Call InitTable
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    With cmd�˳�
        .Left = Me.ScaleWidth - .Width - 150
        .Top = Me.ScaleHeight - .Height - 150
    End With
    cmdExcel.Top = cmd�˳�.Top
    
    With mshList
        .Height = cmd�˳�.Top - 600
        .Width = Me.ScaleWidth
    End With
End Sub

Private Sub InitTable()
    With mshList
        .Clear
        .Rows = 2
        .Cols = 12
        .TextMatrix(0, 0) = "����ID"
        .TextMatrix(0, 1) = "����ID"
        .TextMatrix(0, 2) = "����"
        .TextMatrix(0, 3) = "סԺ��"
        .TextMatrix(0, 4) = "��Ժ����"
        .TextMatrix(0, 5) = "��Ժ����"
        .TextMatrix(0, 6) = "סԺ����"
        .TextMatrix(0, 7) = "�����ܶ�"
        .TextMatrix(0, 8) = "ҩƷ��"
        .TextMatrix(0, 9) = "��Ŀ¼��ҩƷ��"
        .TextMatrix(0, 10) = "ҩƷ����"
        .TextMatrix(0, 11) = "��Ŀ¼��ҩƷ����"
    End With
End Sub
