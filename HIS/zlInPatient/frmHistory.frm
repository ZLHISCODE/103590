VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmHistory 
   AutoRedraw      =   -1  'True
   Caption         =   "���˱䶯��¼"
   ClientHeight    =   3855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10605
   Icon            =   "frmHistory.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   10605
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   10605
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3300
      Width           =   10605
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "�˳�(&X)"
         Height          =   350
         Left            =   8280
         TabIndex        =   2
         Top             =   105
         Width           =   1575
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh 
      Height          =   3240
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   5715
      _Version        =   393216
      BackColor       =   16777215
      FixedCols       =   0
      RowHeightMin    =   250
      BackColorBkg    =   16777215
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mlng����ID As Long
Public mlng��ҳID As Long

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Integer
    
    Call RestoreWinState(Me, App.ProductName)
    
    msh.RowHeight(0) = 320
    
    On Error GoTo errH
    
    strSQL = _
        " Select B.���� as ����,C.���� as ����,A.����,D.���� as ��λ�ȼ�," & _
        " E.���� as ����ȼ�,A.���λ�ʿ as ��ʿ,A.����ҽʦ as סԺҽʦ,A.����ҽʦ as ����ҽ��,A.����ҽʦ as ����ҽ��,A.���� as ��ǰ����,A.����Ա���� as ��ʼ����Ա," & _
        " Decode(A.��ʼԭ��,1,'��Ժ',2,'��ס',3,'ת��',4,'����',5,'������λ�ȼ�',6,'��������ȼ�',7,'����סԺҽʦ',8,'������ʿ',9,'����תסԺ',10,'Ԥ��Ժ',11,'��������ҽʦ',12,'��������ҽʦ',13,'��������',14 ,'����ҽ��С��', 15, '��������') as ��ʼԭ��," & _
        " To_Char(A.��ʼʱ��,'YYYY-MM-DD HH24:MI:SS') as ��ʼʱ��,A.��ֹ��Ա as ��ֹ����Ա," & _
        " Decode(A.��ֹԭ��,1,'��Ժ',2,'��ס',3,'ת��',4,'����',5,'������λ�ȼ�',6,'��������ȼ�',7,'����סԺҽʦ',8,'������ʿ',9,'����תסԺ',10,'Ԥ��Ժ',11,'��������ҽʦ',12,'��������ҽʦ',13,'��������',14 ,'����ҽ��С��', 15, '��������') as ��ֹԭ��," & _
        " To_Char(A.��ֹʱ��,'YYYY-MM-DD HH24:MI:SS') as ��ֹʱ��" & _
        " " & _
        " From ���˱䶯��¼ A,���ű� B,���ű� C,�շ���ĿĿ¼ D,�շ���ĿĿ¼ E" & _
        " Where A.����ID=B.ID And A.����ID=C.ID" & _
        " And A.��λ�ȼ�ID=D.ID(+) And A.����ȼ�ID=E.ID(+)" & _
        " And A.����ID=[1] And A.��ҳID=[2]" & _
        " And A.��ʼʱ�� is Not NULL" & _
        " Order by A.��ֹʱ��,A.��ʼʱ��,A.����"

    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, mlng��ҳID)
    
    If Not rsTmp.EOF Then
        Set msh.DataSource = rsTmp
        
        For i = 0 To msh.Cols - 1
            If InStr(1, ",����,��ʿ,סԺҽʦ,����ҽʦ,����ҽʦ,��ʼ����Ա,��ʼʱ��,��ֹ����Ա,��ֹʱ��,", "," & msh.TextMatrix(0, i) & ",") = 0 Then
                msh.colAlignment(i) = 1
            Else
                msh.colAlignment(i) = 4
            End If
        Next
    End If
    Call SetGridWidth(msh, Me)
    
    RestoreFlexState msh, App.ProductName & "\" & Me.Name
    For i = 0 To msh.Cols - 1
        msh.ColAlignmentFixed(i) = 4
    Next
    msh.Row = 1: msh.Col = 0: msh.ColSel = msh.Cols - 1
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    msh.Left = 0
    msh.Top = 0
    msh.width = Me.ScaleWidth
    msh.Height = Me.ScaleHeight - picCmd.Height
    cmdExit.Left = Me.ScaleWidth - cmdExit.width * 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub
