VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm����תҽ�� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "����תҽ��"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   Icon            =   "frm����תҽ��.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   8130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   60
      TabIndex        =   13
      Top             =   4620
      Width           =   8025
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
      Left            =   6570
      TabIndex        =   12
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox txt������ 
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
      Left            =   1080
      TabIndex        =   10
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmd�˳� 
      Cancel          =   -1  'True
      Caption         =   "�˳�(&X)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6660
      TabIndex        =   8
      Top             =   4770
      Width           =   1215
   End
   Begin VB.CommandButton cmd�ϴ� 
      Caption         =   "�ϴ�(&S)"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5280
      TabIndex        =   7
      Top             =   4770
      Width           =   1215
   End
   Begin VB.TextBox txt�ܷ��� 
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
      Left            =   6570
      TabIndex        =   5
      Top             =   270
      Width           =   1455
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
      Left            =   3510
      TabIndex        =   3
      Top             =   270
      Width           =   1455
   End
   Begin VB.TextBox txtNO 
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
      Left            =   600
      TabIndex        =   1
      Top             =   270
      Width           =   1455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshDetail 
      Height          =   3195
      Left            =   120
      TabIndex        =   6
      Top             =   750
      Width           =   7905
      _ExtentX        =   13944
      _ExtentY        =   5636
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
      MouseIcon       =   "frm����תҽ��.frx":000C
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lbl�������� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
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
      Left            =   5490
      TabIndex        =   11
      Top             =   4140
      Width           =   960
   End
   Begin VB.Label lbl������ 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "������"
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
      Left            =   240
      TabIndex        =   9
      Top             =   4140
      Width           =   720
   End
   Begin VB.Label lbl�ܷ��� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�ܷ���"
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
      Left            =   5730
      TabIndex        =   4
      Top             =   330
      Width           =   720
   End
   Begin VB.Label lbl���� 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "����"
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
      Left            =   2970
      TabIndex        =   2
      Top             =   330
      Width           =   480
   End
   Begin VB.Label lblNO 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "NO"
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
      Left            =   300
      TabIndex        =   0
      Top             =   330
      Width           =   240
   End
End
Attribute VB_Name = "frm����תҽ��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintInsure As Integer
Private Enum Col_T
    ��Ŀ����
    ҽ������
    ���
    ��λ
    ����
    ʵ�ս��
    ����
End Enum

Public Sub ShowME(ByVal intinsure As Integer)
    mintInsure = intinsure
    Me.Show 1
End Sub

Private Sub CMD�ϴ�_CLICK()
    Dim lng����ID As Long
    Dim str�ſ����� As String
    On Error GoTo errHand
    
    If Val(txtNO.Tag) = 0 Then
        MsgBox "����ȷ�ϵ��ݺţ�", vbInformation, gstrSysName
        Exit Sub
    End If
    lng����ID = Val(txt����.Tag)
    
    str�ſ����� = ��ݱ�ʶ_����(0, lng����ID)
    If str�ſ����� = "" Then Exit Sub
    str�ſ����� = Split(str�ſ�����, ";")(0)
    gcnOracle.Execute "Update ������ü�¼ Set ����ID=" & lng����ID & " Where ����ID=" & Val(txtNO.Tag)
    If ����Һ�_����(Val(txtNO.Tag), lng����ID) Then
        gstrSQL = " zl_����תҽ��_Insert(" & lng����ID & "," & Val(txtNO.Tag) & ",'" & str�ſ����� & "','" & gstrUserName & "','" & txtNO.Text & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "����תҽ��")
        
        Call InitMsh
        txtNO.SetFocus
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmd�˳�_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call InitMsh
End Sub

Private Sub InitMsh()
    With mshDetail
        .Clear
        .Rows = 2
        .Cols = ����
        
        .TextMatrix(0, ��Ŀ����) = "��Ŀ����"
        .TextMatrix(0, ҽ������) = "ҽ������"
        .TextMatrix(0, ���) = "���"
        .TextMatrix(0, ��λ) = "��λ"
        .TextMatrix(0, ����) = "����"
        .TextMatrix(0, ʵ�ս��) = "ʵ�ս��"
        
        .ColWidth(��Ŀ����) = 2000
        .ColWidth(ҽ������) = 1500
        .ColWidth(���) = 1200
        .ColWidth(��λ) = 800
        .ColWidth(����) = 1000
        .ColWidth(ʵ�ս��) = 1200
    End With
    
    txt������.Text = ""
    txt��������.Text = ""
    txt����.Text = ""
    txt�ܷ���.Text = ""
End Sub

Private Sub txtNO_GotFocus()
    txtNO.SelStart = 0
    txtNO.SelLength = Len(txtNO.Text)
End Sub

Private Sub txtNO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim StrInput As String
    Dim datCurr As Date
    Dim dblMoney As Double
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHand
    If KeyCode <> vbKeyReturn Then Exit Sub
    StrInput = Trim(txtNO.Text)
    If Len(StrInput) < 4 Then
        datCurr = zlDatabase.Currentdate()
        txtNO.Text = PreFixNO & Format(CDate(Format(datCurr, "YYYY-MM-dd")) - CDate(Format(datCurr, "YYYY") & "-01-01") + 1, "000") & Format(StrInput, "0000") '����˳����
    Else
        txtNO.Text = GetFullNO(StrInput)
    End If
    
    '��ȡ������ϸ
    gstrSQL = " Select A.����ID,A.����ID,A.����,A.������,A.�Ǽ�ʱ��,B.���� AS ��Ŀ����,C.��Ŀ����,B.���,B.���㵥λ,A.����*A.���� AS ����,A.ʵ�ս��" & _
              " From ������ü�¼ A,�շ�ϸĿ B,����֧����Ŀ C " & _
              " Where A.�շ�ϸĿID=B.ID And B.ID=C.�շ�ϸĿID(+) And C.����(+)=[1]" & _
              " And Mod(A.��¼����,10)=1 And Nvl(A.ʵ�ս��,0)<>0 And Nvl(A.���ӱ�־,0)<>9 And Nvl(A.��¼״̬,0)<>0 And A.NO=[2]" & _
              " Order by B.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������ϸ", mintInsure, CStr(txtNO.Text))
    If rsTemp.RecordCount = 0 Then
        MsgBox "û���ҵ��õ��ݣ����������룡", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '��鲡��Ԥ����¼���������ҽ��֧����ʽ�������ٴν���ҽ������
    If ISInsure(rsTemp!����ID) Then
        MsgBox "���ϴ���ҽ�����������ٴν���ҽ�����㣡", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With rsTemp
        txt����.Text = !����
        txt������.Text = Nvl(!������)
        txt��������.Text = Format(!�Ǽ�ʱ��, "yyyy-MM-dd")
        txtNO.Tag = !����ID
        txt����.Tag = Nvl(!����ID, 0)
        
        Do While Not .EOF
            mshDetail.TextMatrix(.AbsolutePosition, ��Ŀ����) = !��Ŀ����
            mshDetail.TextMatrix(.AbsolutePosition, ҽ������) = Nvl(!��Ŀ����)
            mshDetail.TextMatrix(.AbsolutePosition, ���) = Nvl(!���)
            mshDetail.TextMatrix(.AbsolutePosition, ��λ) = Nvl(!���㵥λ)
            mshDetail.TextMatrix(.AbsolutePosition, ����) = Format(!����, "#0.00")
            mshDetail.TextMatrix(.AbsolutePosition, ʵ�ս��) = Format(!ʵ�ս��, "#0.00")
            
            dblMoney = dblMoney + !ʵ�ս��
            mshDetail.Rows = mshDetail.Rows + 1
            .MoveNext
        Loop
        mshDetail.Rows = mshDetail.Rows - 1
        txt�ܷ���.Text = Format(dblMoney, "#0.00")
    End With
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Function ISInsure(ByVal lng����ID As Long) As String
    '�������ҽ���Ľ��㷽ʽ˵���ѽ��й�ҽ������
    Dim rsTemp As New ADODB.Recordset
    gstrSQL = " select ���㷽ʽ,Nvl(��Ԥ��,0) AS ��� " & _
              " from ����Ԥ����¼ A,���㷽ʽ B " & _
              " where A.����ID = [1] and A.��¼����=3 and A.��¼״̬=1 " & _
              " And A.���㷽ʽ=B.���� And B.���� IN (3,4)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡHIS��Ԥ����¼���", lng����ID)
    ISInsure = (rsTemp.RecordCount <> 0)
End Function
