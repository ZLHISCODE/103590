VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm������Ŀ����_���� 
   Caption         =   "������Ŀ����_����"
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "frm������Ŀ����_����.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7080
   ScaleWidth      =   11880
   StartUpPosition =   1  '����������
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   11820
      TabIndex        =   1
      Top             =   5985
      Width           =   11880
      Begin VB.CommandButton cmdȫ�� 
         Caption         =   "ȫ��(&A)"
         Height          =   350
         Left            =   0
         TabIndex        =   6
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmd���� 
         Caption         =   "����(&O)"
         Height          =   350
         Left            =   7320
         TabIndex        =   5
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdȫ�Է� 
         Caption         =   "ȫ�Է�(&C)"
         Height          =   350
         Left            =   1200
         TabIndex        =   4
         Top             =   120
         Width           =   1100
      End
      Begin VB.TextBox txtסԺ�� 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2052
            SubFormatType   =   0
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   300
         Left            =   3360
         TabIndex        =   3
         Top             =   150
         Width           =   1455
      End
      Begin VB.CommandButton cmd��־ 
         Caption         =   "��־(&L)"
         Height          =   345
         Left            =   6090
         TabIndex        =   2
         Top             =   120
         Width           =   1200
      End
      Begin VB.Label lblNote 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   555
         Width           =   4755
      End
      Begin VB.Label Label1 
         Caption         =   "!�ر�˵������'��'��ʾ����,��'��'��ʾ�Է�,�ձ�ʾ������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   5160
         TabIndex        =   8
         Top             =   615
         Width           =   7935
      End
      Begin VB.Label lblסԺ�� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "סԺ��"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   2715
         TabIndex        =   7
         Top             =   210
         Width           =   540
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4920
      Top             =   2190
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfDetail 
      Height          =   5715
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   10081
      _Version        =   393216
      FixedCols       =   0
      FocusRect       =   0
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frm������Ŀ����_����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mintInsure As Integer

Public Sub ShowSelect(ByVal intinsure As Integer)
    On Error Resume Next
    mintInsure = intinsure
    mlng����ID = 0
    mlng��ҳID = 0
    Me.Show 1
End Sub

'���ڣ���Ҫÿһ����ϸ����Ҫ����ҩƷ����
'���ԣ�ԭ�շ�ϸĿID�����շ�ID

Private Sub Cmd����_Click()
    Dim strҽ������ As String
    Dim lngRow As Long, lngRows As Long
    Dim rsTemp As New ADODB.Recordset
    Dim str�������� As String
    
    Dim strTableD       As String
    Dim strWhereD       As String
    Dim i               As Integer
    Dim sFileName       As String
    
    
    If mlng����ID = 0 Then
        MsgBox "����ȷ������!", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gcnOracle.BeginTrans
    
    '�����־
    
    '��¼�޸�ǰ��־
    ' �����(�÷ֺ�";"����)
    strTableD = "����ҩƷ�շ�"
    ' �������(�÷ֺ�";"����)
    strWhereD = "����ID='" & mlng����ID & "' And ��ҳID = '" & mlng��ҳID & "'"
    ' ��¼�޸�ǰ������
    sFileName = EditFormerWriteFileA(strTableD, strWhereD)
    
    lngRows = msfDetail.Rows - 1
    For lngRow = 1 To lngRows
        '��������ҩƷ�շѱ�
        If msfDetail.TextMatrix(lngRow, 0) <> "" Then
            If msfDetail.RowData(lngRow) <> 0 Then
                gstrSQL = "ZL_����ҩƷ�շ�_Update(" & mlng����ID & "," & mlng��ҳID & "," & msfDetail.RowData(lngRow) & "," & IIf(msfDetail.TextMatrix(lngRow, 0) = "��", 1, 0) & ",'" & gstrUserName & "')"
                gcnOracle.Execute gstrSQL, , adCmdStoredProc
            End If
        End If
    Next
    
    '��ȡ��������ҩƷ�շ�,��������
    gstrSQL = " Select A.ID,A.NO,A.��¼����,A.��¼״̬,A.���,A.�շ����,C.��Ŀ���� AS ҽ������,B.��־" & _
              " From ���˷��ü�¼ A,����ҩƷ�շ� B,����֧����Ŀ C" & _
              " Where Nvl(A.ʵ�ս��,0)<>0 And Nvl(A.����,0)<>0 And Nvl(A.���ӱ�־,0)<>9 And Nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.��¼״̬,0)<>0 And Nvl(A.Ӥ����,0)=0 " & _
              " And A.����ID=B.����ID And A.��ҳID=B.��ҳID And A.ID=B.����ID And A.�շ�ϸĿID+0=C.�շ�ϸĿID And C.����=[3]" & _
              " And B.����ID=[1] And B.��ҳID=[2]" & _
              " Order by A.�շ�ϸĿID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������ҩƷ�շ�,��������", mlng����ID, mlng��ҳID, mintInsure)
    Do While Not rsTemp.EOF
        str�������� = ""
        If rsTemp!��־ <> 0 Then    '�Է� �� ���� ״̬δ���� �������д���
            If rsTemp!��־ = 1 Then
                strҽ������ = rsTemp!ҽ������
            Else
               '�Է���ҩ���룺810851900099,20110121����ǿ�޸ģ�ҽ�����������±���
                '�Է��г�ҩ���룺820851900099
                '�Է��в�ҩ���룺829000900099
                If rsTemp!�շ���� = "5" Then
                strҽ������ = "810851900099"
                ElseIf rsTemp!�շ���� = "6" Then
                strҽ������ = "820851900099"
                Else
                strҽ������ = "829000900099"
                End If
                str�������� = "ҩƷȫ�Է�"
            End If
            gstrSQL = "zl_���˷��ü�¼_�ϴ�('" & rsTemp!NO & "'," & rsTemp!��� & "," & rsTemp!��¼���� & "," & rsTemp!��¼״̬ & "," & _
                      "'" & strҽ������ & "'," & IIf(str�������� = "", "NULL", "'" & str�������� & "'") & ",0)"
            gcnOracle.Execute gstrSQL, , adCmdStoredProc
        End If
        
        rsTemp.MoveNext
    Loop
    
     '��¼�޸ĺ���־
    Call EditFormerWriteFileA(strTableD, strWhereD, sFileName)
    '�����޸���־
    AddLog "ҽ������", "����ҩƷ�շ�", DBConnLTEdit, , sFileName, CStr(mlng����ID), CStr(mlng��ҳID), , "����ҩƷ�շ�", , True
    
    gcnOracle.CommitTrans
    
    '�����񣬵ȴ�����������
    mlng����ID = 0
    mlng��ҳID = 0
    Me.txtסԺ��.Text = ""
    Me.txtסԺ��.Tag = ""
    Me.lblNote.Caption = ""
    
    
     msfDetail.Rows = 2: msfDetail.Cols = 10
    msfDetail.TextMatrix(0, 0) = "�򹴱���"
    msfDetail.TextMatrix(0, 1) = "������Ŀ"
    msfDetail.TextMatrix(0, 2) = "���"
    msfDetail.TextMatrix(0, 3) = "�޲�����Ϣ"
    msfDetail.TextMatrix(0, 4) = "��������"
    msfDetail.TextMatrix(0, 5) = "���ݺ�"
    msfDetail.TextMatrix(0, 6) = "����"
    msfDetail.TextMatrix(0, 7) = "����"
    msfDetail.TextMatrix(0, 8) = "���"
    msfDetail.TextMatrix(0, 9) = "������"
    
    
    msfDetail.ColWidth(0) = 1000
    msfDetail.ColWidth(1) = 2400
    msfDetail.ColWidth(2) = 1500
    msfDetail.ColWidth(3) = 3000
    msfDetail.ColWidth(4) = 1000
    msfDetail.ColWidth(5) = 800
    msfDetail.ColWidth(6) = 800
    msfDetail.ColWidth(7) = 800
    msfDetail.ColWidth(8) = 1000
    msfDetail.ColWidth(9) = 800
  
    
    MsgBox "���³ɹ���", vbInformation, gstrSysName
   
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Sub

Private Sub cmdȫ��_Click()
    Dim lngRow As Long, lngRows As Long
    lngRows = msfDetail.Rows - 1
    For lngRow = 1 To lngRows
        msfDetail.TextMatrix(lngRow, 0) = "��"
    Next
End Sub

Private Sub cmdȫ�Է�_Click()
    Dim lngRow As Long, lngRows As Long
    lngRows = msfDetail.Rows - 1
    For lngRow = 1 To lngRows
        msfDetail.TextMatrix(lngRow, 0) = "��"
    Next
End Sub

Private Sub RefreshData()
  ' On Error GoTo errHand
    Dim rsTemp As New ADODB.Recordset
    '���ò����Ƿ����������ҩ
    
    gstrSQL = " Select a.Id,to_char(a.�Ǽ�ʱ��,'yyyy-mm-dd') As ��������,a.No As ���ݺ�,a.����, Substr(c.���,1,Instr(���, '��'))  as ���,a.���� As ����,a.��׼���� As ����,nvl(a.ʵ�ս��,0) As ���,a.����Ա���� As ������,a.�շ�ϸĿid, c.���� As ������Ŀ,c.����, b.��Ŀ����, c.˵�� As �޲�����Ϣ," & _
             " Nvl(d.��־, 1) As ����" & _
              " From ���˷��ü�¼ A,����֧����Ŀ B,�շ�ϸĿ C,����ҩƷ�շ� D" & _
              " Where Nvl(A.ʵ�ս��,0)<>0 And Nvl(A.����,0)<>0 And Nvl(A.���ӱ�־,0)<>9 And Nvl(A.�Ƿ��ϴ�,0)=0 And Nvl(A.��¼״̬,0)<>0 And Nvl(A.Ӥ����,0)=0 " & _
              " And C.��� IN ('5','6','7') And C.˵�� Is Not NULL And A.�շ�ϸĿID=C.ID And C.ID=B.�շ�ϸĿID And B.����=[3]" & _
              " And A.����ID=D.����ID(+) And A.��ҳID=D.��ҳID(+) And A.�շ�ϸĿID=c.ID(+)  and  a.Id=d.����id(+) " & _
              " And A.����ID=[1] And A.��ҳID=[2]" & _
              " Order by C.����,a.�Ǽ�ʱ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ò����Ƿ����������ҩ", mlng����ID, mlng��ҳID, mintInsure)
    '����ǿ20100521�������б����ʾ�У���������,���ݺ�,����,����,���,������,���
    With rsTemp
        Do While Not .EOF
            msfDetail.TextMatrix(.AbsolutePosition, 0) = ""
            msfDetail.TextMatrix(.AbsolutePosition, 1) = !������Ŀ
            msfDetail.TextMatrix(.AbsolutePosition, 2) = !���
            msfDetail.TextMatrix(.AbsolutePosition, 3) = Nvl(!�޲�����Ϣ)
            msfDetail.TextMatrix(.AbsolutePosition, 4) = !��������
            msfDetail.TextMatrix(.AbsolutePosition, 5) = !���ݺ�
            msfDetail.TextMatrix(.AbsolutePosition, 6) = !����
            msfDetail.TextMatrix(.AbsolutePosition, 7) = !����
            msfDetail.TextMatrix(.AbsolutePosition, 8) = !���
            msfDetail.TextMatrix(.AbsolutePosition, 9) = !������
            

         '   msfDetail.TextMatrix(.AbsolutePosition, 0) = IIf(!���� = 1, "��", "��"),����ǿ�޸�Ĭ��Ϊ��
                       '������ϸID
            msfDetail.RowData(.AbsolutePosition) = !ID
            msfDetail.Rows = msfDetail.Rows + 1
            .MoveNext
        Loop
    End With
    Exit Sub
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Load()
    
     
   msfDetail.Rows = 2: msfDetail.Cols = 10
    msfDetail.TextMatrix(0, 0) = "�򹴱���"
    msfDetail.TextMatrix(0, 1) = "������Ŀ"
    msfDetail.TextMatrix(0, 2) = "���"
    msfDetail.TextMatrix(0, 3) = "�޲�����Ϣ"
    msfDetail.TextMatrix(0, 4) = "��������"
    msfDetail.TextMatrix(0, 5) = "���ݺ�"
    msfDetail.TextMatrix(0, 6) = "����"
    msfDetail.TextMatrix(0, 7) = "����"
    msfDetail.TextMatrix(0, 8) = "���"
    msfDetail.TextMatrix(0, 9) = "������"
    
    
    msfDetail.ColWidth(0) = 1000
    msfDetail.ColWidth(1) = 2400
    msfDetail.ColWidth(2) = 1500
    msfDetail.ColWidth(3) = 3000
    msfDetail.ColWidth(4) = 1000
    msfDetail.ColWidth(5) = 800
    msfDetail.ColWidth(6) = 800
    msfDetail.ColWidth(7) = 800
    msfDetail.ColWidth(8) = 1000
    msfDetail.ColWidth(9) = 800
   
   
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    msfDetail.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - Picture1.ScaleHeight
End Sub

Private Sub msfDetail_DblClick()
    If msfDetail.TextMatrix(msfDetail.Row, 0) = "��" Then
        msfDetail.TextMatrix(msfDetail.Row, 0) = "��"
    ElseIf msfDetail.TextMatrix(msfDetail.Row, 0) = "��" Then
        msfDetail.TextMatrix(msfDetail.Row, 0) = ""
    ElseIf msfDetail.TextMatrix(msfDetail.Row, 0) = "" Then
        msfDetail.TextMatrix(msfDetail.Row, 0) = "��"
    End If
End Sub

Private Sub msfDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeySpace Then Exit Sub
    If msfDetail.TextMatrix(msfDetail.Row, 0) = "��" Then
        msfDetail.TextMatrix(msfDetail.Row, 0) = "��"
    ElseIf msfDetail.TextMatrix(msfDetail.Row, 0) = "��" Then
        msfDetail.TextMatrix(msfDetail.Row, 0) = ""
    ElseIf msfDetail.TextMatrix(msfDetail.Row, 0) = "" Then
        msfDetail.TextMatrix(msfDetail.Row, 0) = "��"
    End If
End Sub

Private Sub txtסԺ��_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
      
    msfDetail.Rows = 2: msfDetail.Cols = 10
    msfDetail.TextMatrix(0, 0) = "�򹴱���"
    msfDetail.TextMatrix(0, 1) = "������Ŀ"
    msfDetail.TextMatrix(0, 2) = "���"
    msfDetail.TextMatrix(0, 3) = "�޲�����Ϣ"
    msfDetail.TextMatrix(0, 4) = "��������"
    msfDetail.TextMatrix(0, 5) = "���ݺ�"
    msfDetail.TextMatrix(0, 6) = "����"
    msfDetail.TextMatrix(0, 7) = "����"
    msfDetail.TextMatrix(0, 8) = "���"
    msfDetail.TextMatrix(0, 9) = "������"
    
    
    msfDetail.ColWidth(0) = 1000
    msfDetail.ColWidth(1) = 2400
    msfDetail.ColWidth(2) = 1500
    msfDetail.ColWidth(3) = 3000
    msfDetail.ColWidth(4) = 1000
    msfDetail.ColWidth(5) = 800
    msfDetail.ColWidth(6) = 800
    msfDetail.ColWidth(7) = 800
    msfDetail.ColWidth(8) = 1000
    msfDetail.ColWidth(9) = 800
    msfDetail.ColAlignment(0) = 3
    msfDetail.ColAlignment(3) = 1
    msfDetail.ColAlignment(2) = 1
    msfDetail.ColAlignment(6) = 3
    
    If Trim(txtסԺ��.Text) = "" Then
        txtסԺ��.Tag = ""
        Exit Sub
    End If
    

    
    gstrSQL = " Select A.����ID,A.סԺ���� AS ��ҳID,A.����,A.�Ա�,B.���� AS ���� " & _
              " From ������Ϣ A,���ű� B" & _
              " Where A.��ǰ����ID=B.ID(+) And A.סԺ��=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ������Ϣ", CStr(txtסԺ��.Text))
    If rsTemp.RecordCount = 0 Or IsNull(rsTemp!��ҳID) = True Then
        MsgBox "û���ҵ��ò��ˣ�", vbInformation, gstrSysName
        txtסԺ��.Tag = ""
        txtסԺ��.SetFocus
        
        Exit Sub
    End If
    
    Me.lblNote.Caption = rsTemp!���� & " " & rsTemp!���� & " " & rsTemp!�Ա�
    Me.txtסԺ��.Tag = rsTemp!����ID & "|" & rsTemp!��ҳID
    mlng����ID = rsTemp!����ID
    mlng��ҳID = rsTemp!��ҳID
    
    Call RefreshData
    
    '�����Ա����Ϊ��ʿ����ֻ�ܲ鿴��ʿ���ڲ�������Ա
    cmd����.Enabled = True
    gstrSQL = "select Count(1) from ��Ա����˵�� where ��Ա���� = '��ʿ' and ��ԱID = [1]"
    If Val(zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UserInfo.ID).Fields(0)) = 1 Then
        '��Ա����Ϊ��ʿ����Ҫ���Ȩ��
               
        gstrSQL = "select Count(1) from ������Ϣ a , ������Ա b where a.��ǰ����id = b.����id And ����id = [1] " & _
        " And ��ǰ����id in(select ����ID FROM ������Ա WHERE  ��ԱID=[2] )"
        If Val(zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID, UserInfo.ID).Fields(0)) = 0 Then
            cmd����.Enabled = False
            MsgBox "�˲����ѳ�Ժ�����ڱ����ң���˶ԣ�����ѳ�Ժ���볷����Ժ������������ת�ƣ�����ϵҽ���ƣ�", vbInformation, gstrSysName
        End If
    End If

    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmd��־_Click()
On Error GoTo ErrH
    If mlng����ID = 0 Then Exit Sub
    With frmҽ��������־
        .strģ�� = "ҽ������"
        .str���� = "����ҩƷ�շ�"
        .str����1 = mlng����ID
        .str����2 = mlng��ҳID
        .Show vbModal
    End With
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

