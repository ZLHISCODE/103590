VERSION 5.00
Begin VB.Form frm��ͨ_��Ŀ��Ϣ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Ŀ��Ϣ"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   Icon            =   "frm��ͨ_��Ŀ��Ϣ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3885
      TabIndex        =   14
      Top             =   2205
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2640
      TabIndex        =   13
      Top             =   2205
      Width           =   1100
   End
   Begin VB.ComboBox cboҽ������ 
      Height          =   300
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   12
      Top             =   2040
      Width           =   5715
   End
   Begin VB.Label lbl�շ�ϸĿ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�շ�ϸĿ"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   8
      Top             =   1290
      Width           =   720
   End
   Begin VB.Label lbl�շ�ϸĿ 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�շ�ϸĿ"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   1020
      TabIndex        =   9
      Top             =   1290
      Width           =   3930
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
      Index           =   0
      Left            =   180
      TabIndex        =   6
      Top             =   600
      Width           =   720
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
      Index           =   0
      Left            =   2400
      TabIndex        =   2
      Top             =   300
      Width           =   360
   End
   Begin VB.Label lbl������Ϣ 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ������"
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   1
      Left            =   1020
      TabIndex        =   7
      Top             =   600
      Width           =   3930
   End
   Begin VB.Label lbl��ʶ�� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ʶ��"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   1020
      TabIndex        =   1
      Top             =   300
      Width           =   540
   End
   Begin VB.Label lbl��ʶ�� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��ʶ��"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   300
      Width           =   540
   End
   Begin VB.Label lbl�Ա� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�Ա�"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   4650
      TabIndex        =   5
      Top             =   300
      Width           =   360
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
      Index           =   0
      Left            =   4140
      TabIndex        =   4
      Top             =   300
      Width           =   360
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   2910
      TabIndex        =   3
      Top             =   300
      Width           =   360
   End
   Begin VB.Label lblҽ������ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ҽ������"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      TabIndex        =   10
      Top             =   1620
      Width           =   720
   End
End
Attribute VB_Name = "frm��ͨ_��Ŀ��Ϣ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintType As Integer 'intType-��������(0-ҽ��,1-�����շ�,2-סԺ����)
Private mlng����ID As Long
Private mlngϸĿID As Long
Private mstrժҪ As String
Private mstr��ע As String
Private mbln�в�ҩ As Boolean

Public Function ShowME(ByVal intType As Integer, ByVal lng����ID As Long, ByVal lngϸĿID As Long, _
    ByVal strժҪ As String, ByVal str��ע As String, ByVal bln�в�ҩ As Boolean) As String
    'bln�в�ҩ=true,�������в�ҩ����ѡ��ҩƷ�ķ������ͣ������Ƿ�Ŀ¼��ҩƷ����ѡ���Ǵ󲡻���������ҩ
    mintType = intType
    mlng����ID = lng����ID
    mlngϸĿID = lngϸĿID
    mstrժҪ = Trim(UCase(strժҪ))
    mstr��ע = str��ע
    mbln�в�ҩ = bln�в�ҩ
    Me.Show 1
    ShowME = mstrժҪ
End Function

Private Sub cmdCancel_Click()
    If mstrժҪ = "" Then mstrժҪ = IIf(mbln�в�ҩ, "uzy03", "��ͨ")
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Select Case cboҽ������.ListIndex
    Case 0
        mstrժҪ = IIf(mbln�в�ҩ, "uzy01", "��ͨ")
    Case 1
        mstrժҪ = IIf(mbln�в�ҩ, "uzy02", "����")
    Case 2
        mstrժҪ = IIf(mbln�в�ҩ, "uzy03", "������Ҫ���󲡣�")
    End Select
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim rsTemp As New ADODB.Recordset
    
    '��ȡ������Ϣ
    gstrSQL = "Select ����,�Ա�,�����,סԺ�� From ������Ϣ Where ����ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ", mlng����ID)
    If mintType = 1 Then
        Me.lbl��ʶ��(1).Caption = Nvl(rsTemp!�����)
    Else
        Me.lbl��ʶ��(1).Caption = Nvl(rsTemp!סԺ��)
    End If
    Me.lbl����(1).Caption = Nvl(rsTemp!����)
    Me.lbl�Ա�(1).Caption = Nvl(rsTemp!�Ա�)
    
    '��ʾ���ݻ�ҽ������Ҫ��Ϣ�����Ϊ�գ���ʾΪ��ǰ����
    Me.lbl������Ϣ(1).Caption = IIf(Trim(mstr��ע) = "", "��ǰ���ʵ���", mstr��ע)
    
    '��ȡ�շ�ϸĿ��Ϣ
    gstrSQL = "Select 'Ʒ��:('||����||')'||����||' ���:'||Nvl(���,'') AS Ʒ�� From �շ�ϸĿ Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�շ�ϸĿ��Ϣ", mlngϸĿID)
    Me.lbl�շ�ϸĿ(1).Caption = Nvl(rsTemp!Ʒ��)
    
    If mbln�в�ҩ Then
        Me.cboҽ������.AddItem "����-uzy01"
        Me.cboҽ������.AddItem "����-uzy02"
        Me.cboҽ������.AddItem "��ҽ��-uzy03"
    Else
        Me.lblҽ������.Caption = "��ҩ��ʶ"
        Me.cboҽ������.AddItem "��ͨ-0"
        Me.cboҽ������.AddItem "����-1"
        Me.cboҽ������.AddItem "������Ҫ���󲡣�-2"
    End If
    Me.cboҽ������.ListIndex = 0
    
End Sub

Private Sub FindCbo()
    '������ǰ�趨��ժҪ����λ��ǰҽ������
    If mstrժҪ = "" Then Exit Sub
    Select Case mstrժҪ
    Case "UZY02", "����"
        cboҽ������.ListIndex = 1
    Case "UZY03", "������Ҫ���󲡣�"
        cboҽ������.ListIndex = 2
    End Select
End Sub
