VERSION 5.00
Begin VB.Form frm�޼�����_�山 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "�޼�������Ϣ����"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4635
      TabIndex        =   6
      Top             =   3315
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   3210
      TabIndex        =   5
      Top             =   3300
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   0
      Left            =   -1095
      TabIndex        =   8
      Top             =   3150
      Width           =   9300
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   1
      Left            =   -45
      TabIndex        =   7
      Top             =   630
      Width           =   6135
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "����ͨ��"
      Height          =   240
      Left            =   1260
      TabIndex        =   0
      Top             =   2370
      Value           =   1  'Checked
      Width           =   1080
   End
   Begin VB.TextBox txtEdit 
      Height          =   315
      Index           =   0
      Left            =   1260
      TabIndex        =   2
      Top             =   2693
      Width           =   1530
   End
   Begin VB.TextBox txtEdit 
      Height          =   315
      Index           =   1
      Left            =   4245
      TabIndex        =   4
      Top             =   2693
      Width           =   1575
   End
   Begin VB.Image img 
      Height          =   555
      Left            =   120
      Picture         =   "frm�޼�����_�山.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "ҽ����Ŀ�۸����200.00Ԫ��ȷ����ص�������Ϣ��"
      Height          =   225
      Index           =   0
      Left            =   945
      TabIndex        =   18
      Top             =   135
      Width           =   4965
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�շ���Ŀ����"
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   855
      Width           =   1080
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   0
      Left            =   1260
      TabIndex        =   16
      Top             =   795
      Width           =   1065
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "�շ���Ŀ����"
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   1245
      Width           =   1080
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   1
      Left            =   1260
      TabIndex        =   14
      Top             =   1185
      Width           =   4560
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "ҽ����Ŀ����"
      Height          =   180
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   1635
      Width           =   1080
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   2
      Left            =   1260
      TabIndex        =   12
      Top             =   1575
      Width           =   1065
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "ҽ����Ŀ����"
      Height          =   180
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   3
      Left            =   1260
      TabIndex        =   10
      Top             =   1980
      Width           =   4560
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����ҽ��"
      Height          =   180
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   2760
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "����ҽ��ְ��"
      Height          =   180
      Index           =   6
      Left            =   3105
      TabIndex        =   3
      Top             =   2760
      Width           =   1080
   End
   Begin VB.Label lblInfor 
      Height          =   150
      Left            =   975
      TabIndex        =   9
      Top             =   390
      Width           =   5040
   End
End
Attribute VB_Name = "frm�޼�����_�山"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrVerify As String
Dim mlng����ID As Long
 
Dim mstrCode As String
Private Sub chk����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
 
End Sub

Private Sub cmdCancel_Click()
    mstrVerify = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
   Dim strSQL As String
    mstrVerify = chk����.Value & "||" & txtEdit(0).Text & "||" & txtEdit(1).Text
    '--ZL_ҽ����ϸ����_INSERT(
    '    ����ID_IN IN ҽ����ϸ����.����ID%TYPE,
    '    ������_IN IN ҽ����ϸ����.������%TYPE,
    '    ������ְ��_IN IN ҽ����ϸ����.������ְ��%TYPE,
    '    ������־_IN IN ҽ����ϸ����.������־%TYPE
    '   ������,������
    '    �˵�������ˮ��
    strSQL = "ZL_ҽ����ϸ����_INSERT(" & _
         mlng����ID & "," & _
         "'" & txtEdit(0).Text & "'," & _
         "'" & txtEdit(1).Text & "'," & _
         IIf(chk����.Value = 1, 2, 0) & ",'" & _
         g�������_�����山.������ & "','" & _
         g�������_�����山.������ & "'," & _
         "NULL," & _
         IIf(mstrCode = "", "NULL", "'" & mstrCode & "'") & ")"
    
    Call SQLTest(App.ProductName, Me.Caption, strSQL)
    gcnOracle_CQYB.Execute strSQL, , adCmdStoredProc
    Call SQLTest
    Unload Me
End Sub

Private Sub Form_Load()
    lbl(0).Caption = "ҽ����Ŀ�۸����" & InitInfor_�����山.�����޼� & "Ԫ��ȷ����ص�������Ϣ��"
    Me.cmdOK.Enabled = True
End Sub
Public Function Get������Ϣ(lng����ID As Long, strCode As String) As String
    Dim rsTemp As New ADODB.Recordset
        
    gstrSQL = " select A.ID,NO,���, b.����,b.����,c.��Ŀ����, c.��Ŀ���� " & _
             " from ������ü�¼ a, �շ�ϸĿ b,����֧����Ŀ c " & _
             " where a.�շ�ϸĿid=b.id and a.�շ�ϸĿid=c.�շ�ϸĿid and a.id=[1]" & _
             " UNION " & _
             " select A.ID,NO,���, b.����,b.����,c.��Ŀ����, c.��Ŀ���� " & _
             " from סԺ���ü�¼ a, �շ�ϸĿ b,����֧����Ŀ c " & _
             " where a.�շ�ϸĿid=b.id and a.�շ�ϸĿid=c.�շ�ϸĿid and a.id=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ϸ��¼��Ϣ!", lng����ID)
    
    If rsTemp.EOF Then
        Exit Function
    End If
    
    lblInfor.Caption = "NO:" & Nvl(rsTemp!NO) & "   �к�:" & Nvl(rsTemp!���, 0)
    
    lblEdit(0).Caption = Nvl(rsTemp!����)
    lblEdit(1).Caption = Nvl(rsTemp!����)
    lblEdit(2).Caption = Nvl(rsTemp!��Ŀ����)
    lblEdit(3).Caption = Nvl(rsTemp!��Ŀ����)
    mlng����ID = lng����ID
    mstrCode = strCode
    Me.Show 1
    Get������Ϣ = mstrVerify
End Function

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
