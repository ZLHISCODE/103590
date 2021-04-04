VERSION 5.00
Begin VB.Form frmNurseFileChange 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ļ���ʽ���"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4620
   Icon            =   "frmNurseFileChange.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCanCel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3000
      TabIndex        =   10
      Top             =   1935
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   45
      TabIndex        =   8
      Top             =   1740
      Width           =   4545
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1725
      TabIndex        =   9
      Top             =   1935
      Width           =   1155
   End
   Begin VB.TextBox txtOldName 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1275
      MaxLength       =   50
      TabIndex        =   3
      Top             =   570
      Width           =   2895
   End
   Begin VB.TextBox txtOldFormat 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1275
      MaxLength       =   50
      TabIndex        =   1
      Top             =   225
      Width           =   2895
   End
   Begin VB.ComboBox cboFormat 
      Height          =   300
      Left            =   1275
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   930
      Width           =   2895
   End
   Begin VB.TextBox txtNewName 
      Height          =   285
      Left            =   1275
      MaxLength       =   50
      TabIndex        =   7
      Top             =   1305
      Width           =   2895
   End
   Begin VB.Label lblOldName 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���ļ�����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   315
      TabIndex        =   2
      Top             =   615
      Width           =   900
   End
   Begin VB.Label lblOldFormat 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�ɸ�ʽ"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   675
      TabIndex        =   0
      Top             =   270
      Width           =   540
   End
   Begin VB.Label lblNewForamt 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�¸�ʽ"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   675
      TabIndex        =   4
      Top             =   990
      Width           =   540
   End
   Begin VB.Label lblNewName 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���ļ�����"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   315
      TabIndex        =   6
      Top             =   1350
      Width           =   900
   End
End
Attribute VB_Name = "frmNurseFileChange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOK As Boolean
Private mlngFileID As Long
Private mlngFormatID As Long
Private mblnWave As Boolean 'TRUE�����µ� FALSE����¼��
Private mstrDept As String '��������

Public Function ShowEditor(ByVal mfrmParent As Form, ByVal lngFileID As Long) As Boolean
    mblnOK = False: mblnWave = False
    mlngFileID = lngFileID
    Me.Show 1, mfrmParent
    ShowEditor = mblnOK
End Function

Private Sub cboFormat_Click()
    txtNewName.Text = Split(cboFormat.Text, "-")(1)
    If mstrDept <> "" Then txtNewName.Text = "[" & mstrDept & "]" & txtNewName.Text
End Sub

Private Sub cmdCanCel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
'���ܣ�����ļ���ʽ���
    Dim blnShow As Boolean
    On Error GoTo ErrHand
    
    If mlngFormatID = cboFormat.ItemData(cboFormat.ListIndex) Then
        MsgBox "�滻�ļ��ĸ�ʽ���ܺ�֮ǰ�ĸ�ʽ��ͬ��������ѡ��", vbInformation, gstrSysName
        cboFormat.SetFocus
        Exit Sub
    End If
    If txtNewName.Text = "" Then
        MsgBox "�������ļ����ƣ�", vbInformation, gstrSysName
        txtNewName.SetFocus
        Exit Sub
    End If
    If LenB(StrConv(txtNewName.Text, vbFromUnicode)) > 50 Then
        MsgBox "�ļ����Ƴ����������50���ַ���25�����֣�", vbInformation, gstrSysName
        txtNewName.SetFocus
        Exit Sub
    End If
    If MsgBox("�ò���������Ҫһ��ʱ�䣬�������Ƿ������", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Sub
    Screen.MousePointer = 11
    zlCommFun.ShowFlash "���ڸ����ļ���ʽ���������ĵȴ�....", Me
    blnShow = True
    gstrSQL = "Zl_���˻����ļ�_Repalce(" & mlngFileID & "," & Val(cboFormat.ItemData(cboFormat.ListIndex)) & ",'" & Trim(txtNewName.Text) & "'," & IIf(mblnWave = True, 0, 1) & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "Zl_���˻����ļ�_Repalce")
    zlCommFun.StopFlash
    Screen.MousePointer = 0
    
    mblnOK = True
    Unload Me
    Exit Sub
ErrHand:
    Screen.MousePointer = 0
    If blnShow = True Then zlCommFun.StopFlash
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 39 Then KeyCode = 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHand
    '��ȡ��ǰ�ļ���Ϣ
    gstrSQL = " Select B.���� ��ʽ����,B.���,B.����,B.ID ��ʽID,A.�ļ�����,A.����ID,A.��ҳID,A.Ӥ��,C.���� ��������" & _
          " From ���˻����ļ� A,�����ļ��б� B,���ű� C" & _
          " Where A.��ʽID=B.ID And A.ID=[1] And B.����=3 And A.����ID=C.ID(+)"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ǰ�ļ���Ϣ", mlngFileID)
    txtOldFormat.Text = rsTemp!��� & "-" & rsTemp!��ʽ����
    txtOldName.Text = rsTemp!�ļ�����
    mlngFormatID = rsTemp!��ʽID
    mblnWave = (Val(NVL(rsTemp!����, 0)) = -1)
    mstrDept = NVL(rsTemp!��������)
    '��ȡ��Ӧ�ڵ�ǰ�����ļ�¼�������µ�
    gstrSQL = _
        " Select ID, ����, ���, ��� || '-' || ���� As ��ʽ" & vbNewLine & _
        " From �����ļ��б�" & vbNewLine & _
        " Where ���� = 3  " & IIf(mblnWave = True, " And ���� =-1", " And ���� <> 1 And  ���� <> -1") & " And (ͨ�� = 1 Or (ͨ�� = 2 And ID In (Select �ļ�id From ����Ӧ�ÿ��� Where ����id = [1])))" & vbNewLine & _
        " Order By ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�ļ��б�", glng����ID)
    With rsTemp
        cboFormat.Clear
        Do While Not .EOF
            If !ID <> mlngFormatID Then
                cboFormat.AddItem !��ʽ
                cboFormat.ItemData(cboFormat.NewIndex) = !ID
            End If
        .MoveNext
        Loop
    End With
    If cboFormat.ListCount > 0 Then
        cboFormat.ListIndex = 0
    Else 'û�п���ѡ��ļ�¼�������µ�
        On Error Resume Next
        MsgBox "��[�����ļ��б�]��û���ҵ������ڱ�������������ʽ" & IIf(mblnWave, "", "") & "�ļ���", vbInformation, gstrSysName
        Unload Me
    End If
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

