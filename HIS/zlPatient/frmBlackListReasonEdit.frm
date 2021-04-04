VERSION 5.00
Begin VB.Form frmBlackListReasonEdit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "������Ϊ����ԭ��༭"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4335
      TabIndex        =   8
      Top             =   2220
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   3090
      TabIndex        =   6
      Top             =   2235
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "������Ϣ"
      Height          =   1695
      Left            =   210
      TabIndex        =   7
      Top             =   300
      Width           =   5250
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   840
         MaxLength       =   2
         TabIndex        =   1
         Tag             =   "����"
         Top             =   345
         Width           =   1500
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   840
         MaxLength       =   10
         TabIndex        =   3
         Tag             =   "����"
         Top             =   705
         Width           =   3675
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   840
         MaxLength       =   4
         TabIndex        =   5
         Tag             =   "����"
         Top             =   1110
         Width           =   1455
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "���(&U)"
         Height          =   210
         Index           =   1
         Left            =   180
         TabIndex        =   0
         Top             =   405
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "ԭ��(&N)"
         Height          =   180
         Index           =   2
         Left            =   180
         TabIndex        =   2
         Top             =   765
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&S)"
         Height          =   240
         Index           =   3
         Left            =   180
         TabIndex        =   4
         Top             =   1140
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmBlackListReasonEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Enum gReasonEdit
    EM_Rsn_���� = 0
    EM_Rsn_�޸�
    EM_Rsn_ɾ��
End Enum
Private mbytEditType As gReasonEdit
Private mfrmMain As Object
Private mstrCode As String
Private mblnChange As Boolean     '�Ƿ�ı���
Private mintSuccess As Integer
Private mblnFirst As Boolean
Private mblnUnload As Boolean


Public Function zlShowEdit(ByVal frmMain As Object, ByVal bytEditType As gReasonEdit, Optional strCode As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�༭������Ϊ���
    '���:frmMain-���õ�������
    '    bytEditType-�༭���:0-����;1-�޸�;2-�鿴
    '     strCode-����,����ʱ�����
    '����:�༭�ɹ�����True,����ΪFalse
    '����:���˺�
    '����:2018-11-08 17:01:16
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbytEditType = bytEditType: mintSuccess = 0
    Set mfrmMain = frmMain: mstrCode = strCode: mblnFirst = True
    mblnUnload = False
    
    Me.Show 1, frmMain
    zlShowEdit = mintSuccess > 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetInputDefineSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������ؿؼ����볤�ȣ��õ����ݿ�ı��ֶεĳ��ȣ�
    '����:���˺�
    '����:2018-11-09 17:06:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = "SELECT ����,����,����,�Ƿ�̶� FROM ���ò�����Ϊԭ�� Where Rownum<0 "
    Call zlDatabase.OpenRecordset(rsTemp, strSQL, "���ò�����Ϊԭ��")
    
    txtEdit(1).MaxLength = rsTemp.Fields("����").DefinedSize
    txtEdit(2).MaxLength = rsTemp.Fields("����").DefinedSize
    txtEdit(3).MaxLength = rsTemp.Fields("����").DefinedSize
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

 Private Sub SetCtrolEnabled()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ÿ��Ƶ�enable����
    '����:���˺�
    '����:2018-11-13 18:39:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEdit As Boolean, i As Long
    On Error GoTo errHandle
    
    blnEdit = (mbytEditType = EM_Ty_���� Or mbytEditType = EM_Ty_�޸�)
    txtEdit(1).Enabled = mbytEditType = EM_Ty_����
    txtEdit(2).Enabled = blnEdit
    txtEdit(3).Enabled = blnEdit
    
    For i = 1 To txtEdit.UBound
        txtEdit(i).BackColor = IIf(txtEdit(i).Enabled, &H80000005, &H8000000F)
    Next
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Sub

Private Function ReadData(ByVal strCode As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݱ����ȡ����
    '���:strCode-��ǰ����
    '����:
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-11-09 17:03:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, strSQL As String

    On Error GoTo errHandle
    If mbytEditType = EM_Rsn_���� Then
        '����
        txtEdit(1).Text = zlDatabase.GetMax("���ò�����Ϊԭ��", "����", txtEdit(1).MaxLength)
        Call SetCtrolEnabled
        ReadData = True
        Exit Function
    End If
     
    strSQL = "" & _
    "   SELECT ����,����,����,�Ƿ�̶� FROM ���ò�����Ϊԭ��  Where ����=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strCode)
    If rsTemp.EOF Then
        MsgBox "δ�ҵ�����Ϊ��" & strCode & "���Ĳ�����Ϊԭ�����ݣ�����!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    txtEdit(1).Text = Nvl(rsTemp!����)
    txtEdit(2).Text = Nvl(rsTemp!����)
    txtEdit(3).Text = Nvl(rsTemp!����)
    Call SetCtrolEnabled
    ReadData = True
      
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
 

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    If IsValid() = False Then Exit Sub
    If SaveData() = False Then Exit Sub
    
    mintSuccess = mintSuccess + 1
    If mbytEditType <> 0 Then
        mblnChange = False: Unload Me
        Exit Sub
    End If
    
    mstrCode = ""
    txtEdit(2).Text = ""
    txtEdit(3).Text = ""
    txtEdit(1).Text = zlDatabase.GetMax("���ò�����Ϊԭ��", "����", txtEdit(1).MaxLength)
    '�������ϴεĲ���
    
    mblnChange = False
    If txtEdit(2).Enabled And txtEdit(2).Visible Then txtEdit(2).SetFocus
End Sub

Private Function IsValid() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������й������Ƿ���Ч
    '����:��Ч����True,����ΪFalse
    '����:���˺�
    '����:2018-11-09 17:22:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer, varTemp As Variant, varData As Variant, blnHaveData As Boolean
    Dim strTemp As String
    
    On Error GoTo errHandle
    For i = 1 To 3
        txtEdit(i).Text = Trim(txtEdit(i).Text)
        strTemp = txtEdit(i).Text
        If LenB(StrConv(strTemp, vbFromUnicode)) > txtEdit(i).MaxLength Then
            MsgBox txtEdit(i).Tag & "���ܳ���" & Int(txtEdit(i).MaxLength / 2) & "������" & "��" & txtEdit(i).MaxLength & "����ĸ��", vbExclamation, gstrSysName
            zlControl.ControlSetFocus txtEdit(i)
            Exit Function
        End If
        If InStr(strTemp, "'") > 0 Then
            MsgBox txtEdit(i).Tag & "�к��зǷ��ַ���", vbExclamation, gstrSysName
            zlControl.ControlSetFocus txtEdit(i)
            Exit Function
        End If
    Next
    txtEdit(1).Text = Trim(txtEdit(1).Text)
    
    If Len(txtEdit(1).Text) = 0 Or Trim(txtEdit(1).Text) = "" Then
        MsgBox "���벻��Ϊ�ա�", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(1)
        Exit Function
    End If
    
    If Len(Trim(txtEdit(2).Text)) = 0 Then
        MsgBox "���Ʋ���Ϊ�ա�", vbExclamation, gstrSysName
        zlControl.ControlSetFocus txtEdit(2)
        Exit Function
    End If
    IsValid = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
Private Function SaveData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2018-11-09 17:23:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Long, cllPro As Collection, strSQL As String, blnDel As Boolean
    Dim blnTran As Boolean, strTemp As String, str���� As String, strRule As String
    On Error GoTo errHandle
    Set cllPro = New Collection
    
    '    Zl_���ò�����Ϊԭ��_Update
    strSQL = "Zl_���ò�����Ϊԭ��_Update("
    '  ����_In     Number, 0-����;1-�޸�
    strSQL = strSQL & "" & IIf(mbytEditType = EM_Rsn_����, 0, 1) & ","
    '  ����_In     ���ò�����Ϊԭ��.����%Type,
    strSQL = strSQL & "'" & txtEdit(1).Text & "',"
    '  ����_In     ���ò�����Ϊԭ��.����%Type,
    strSQL = strSQL & "'" & txtEdit(2).Text & "',"
    '  ����_In     ���ò�����Ϊԭ��.����%Type,
    strSQL = strSQL & "'" & txtEdit(3).Text & "',"
    '  �Ƿ�̶�_In ���ò�����Ϊԭ��.�Ƿ�̶�%Type := 0,
    strSQL = strSQL & "0)"
    
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    'zlAddArray cllPro, strSQL
    
    'blnTran = True
    'zlExecuteProcedureArrAy cllPro, Me.Caption
    SaveData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
 
Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If mblnUnload Then Unload Me: Exit Sub
    
    If txtEdit(2).Enabled And txtEdit(2).Visible Then txtEdit(2).SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    
    
    Call SetInputDefineSize '����ȱʡ�����볤��
    
    mblnUnload = Not ReadData(mstrCode) '��ȡ����
    
    mblnChange = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub
     

Private Sub txtEdit_Change(Index As Integer)
    mblnChange = True
    If Index = 2 Then
        txtEdit(3).Text = zlStr.GetCodeByVB(txtEdit(2).Text)
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    If Index = 2 Then zlCommFun.OpenIme True
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("'}|,""/", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    zlCommFun.OpenIme False
End Sub

