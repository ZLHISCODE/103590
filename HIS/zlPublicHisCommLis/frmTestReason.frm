VERSION 5.00
Begin VB.Form frmTestReason 
   BackColor       =   &H00FEEDE9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ò���������"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7170
   BeginProperty Font 
      Name            =   "����"
      Size            =   7.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTestReason.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
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
      Left            =   6015
      TabIndex        =   5
      Top             =   4170
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
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
      Left            =   4665
      TabIndex        =   4
      Top             =   4170
      Width           =   1100
   End
   Begin VB.PictureBox picTestReason 
      Appearance      =   0  'Flat
      BackColor       =   &H00FEEDE9&
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
      Height          =   3930
      Left            =   60
      ScaleHeight     =   3900
      ScaleWidth      =   7020
      TabIndex        =   0
      Top             =   90
      Width           =   7050
      Begin VB.TextBox txtUser 
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
         Left            =   855
         MaxLength       =   20
         TabIndex        =   3
         Top             =   3435
         Width           =   1920
      End
      Begin VB.TextBox txtTestReason 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2820
         Left            =   75
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   525
         Width           =   6870
      End
      Begin VB.ComboBox cboTestReason 
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
         ItemData        =   "frmTestReason.frx":6852
         Left            =   1920
         List            =   "frmTestReason.frx":6854
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   90
         Width           =   5010
      End
      Begin VB.Label lblUser 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�ò���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   75
         TabIndex        =   7
         Top             =   3495
         Width           =   720
      End
      Begin VB.Label lblTestReason 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ѡ���ò�����:"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   90
         TabIndex        =   6
         Top             =   150
         Width           =   1800
      End
   End
End
Attribute VB_Name = "frmTestReason"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnfrmIfShow As Boolean                   '�����Ƿ���ʾ���
Private mstrIDs As String                          '�����������ID
Private mrsTestReason As ADODB.Recordset           'ѡ���ò�����
Private mblnOk As Boolean                          '�Ƿ񱣴�ɹ�
Private mobjLisInsideCom As New clsPublicHisCommLis

Private Sub cboTestReason_Click()
          
1         On Error GoTo cboTestReason_Click_Error
          
2         With cboTestReason
3             If .ListIndex < 0 Then txtTestReason.Text = "": Exit Sub
4             If .Text = "" Then txtTestReason.Text = "": Exit Sub
              
5             If Not mrsTestReason Is Nothing Then
6                 mrsTestReason.Filter = "���� = '" & .Text & "'"
7                 If mrsTestReason.RecordCount > 0 Then
8                     txtTestReason.Text = txtTestReason.Text & mrsTestReason("����") & ""
9                 End If
10            End If
11        End With
          
          
12        Exit Sub
cboTestReason_Click_Error:
13        Call WriteErrLog("zlPublicHisCommLis", "frmTestReason", "ִ��(cboTestReason_Click)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
14        Err.Clear
End Sub

Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Public Function ShowMe(ByVal objFrm As Object, ByVal strIDs As String, ByVal strUser As String) As Boolean
    mstrIDs = strIDs
    txtUser.Text = strUser
    
    Me.Show vbModal, objFrm
    
    ShowMe = mblnOk
End Function

Private Sub GetPerson(ByVal strCode As String)
'��ȡ������Ϣ
    Dim rsTemp As Recordset
    Dim strVal As String
    
    On Error GoTo GetPerson_Error
    
    Set rsTemp = mobjLisInsideCom.GetPersonTwo(strCode)
    
    strVal = SeletItemFromRsOld(Me, rsTemp, "")
    
    If strVal = "" Then
        txtUser.SetFocus
    Else
        txtUser = Split(strVal, ",")(1)
        txtUser.Tag = Split(strVal, ",")(1)
        cmdOK.SetFocus
    End If
    
    
    Exit Sub
GetPerson_Error:
    Call WriteErrLog("zlPublicHisCommLis", "frmTestReason", "ִ��(GetPerson)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
    Err.Clear
End Sub

Private Sub cmdOK_Click()
    If SaveData = True Then
        mblnOk = True
        Unload Me
    End If
End Sub

Private Function SaveData() As Boolean
          '��������
          Dim strSQL() As String
          Dim arrIDs() As String
          Dim i As Integer
          Dim blnTran As Boolean
          
1         On Error GoTo SaveData_Error
          
2         If Trim(txtTestReason.Text) = "" Then
3             MsgBox "������д�ò����ɣ�", vbInformation, gSysInfo.AppName
4             txtTestReason.SetFocus
5             Exit Function
6         End If
          
7         If Trim(txtUser.Text) = "" Then
8             MsgBox "������д�ò��ˣ�", vbInformation, gSysInfo.AppName
9             txtUser.SetFocus
10            Exit Function
11        End If
          
12        arrIDs = Str2Array(mstrIDs, ",", 4000)
          
13        ReDim strSQL(UBound(arrIDs))
          
14        For i = 0 To UBound(arrIDs)
15            strSQL(i) = "Zl_�ò�������ϸ_Edit('" & arrIDs(i) & "','" & Trim(txtUser.Text) & "','" & Trim(txtTestReason.Text) & "')"
16        Next
          
17        gcnLisOracle.BeginTrans
18        blnTran = True
19        For i = 0 To UBound(strSQL)
20            Call ComExecuteProc(Sel_Lis_DB, strSQL(i), "�����ò���������")
21        Next
22        gcnLisOracle.CommitTrans
23        blnTran = False
          
24        SaveData = True
          
          
25        Exit Function
SaveData_Error:
26        If blnTran Then gcnLisOracle.RollbackTrans
27        Call WriteErrLog("zlPublicHisCommLis", "frmTestReason", "ִ��(SaveData)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
28        Err.Clear
End Function

Private Sub Form_Activate()
    If mblnfrmIfShow = False Then
        mblnfrmIfShow = True
        Call InitData
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnfrmIfShow = False
    mstrIDs = ""
    Set mrsTestReason = Nothing
    Set mobjLisInsideCom = Nothing
End Sub

Private Sub InitData()
          '��ѯ�ò���������
          Dim strSQL As String
          
1         On Error GoTo InitData_Error
          
2         strSQL = "Select ����, ����, ����, ���� From �����ֵ�� Where ���� = '�ò���������'"
3         Set mrsTestReason = ComOpenSQL(Sel_Lis_DB, strSQL, "�ò���������")
          
4         With cboTestReason
5             .AddItem ""
6             .ItemData(.NewIndex) = 0
              
7             Do While Not mrsTestReason.EOF
8                 .AddItem mrsTestReason("����") & ""
9                 .ItemData(.NewIndex) = mrsTestReason("����") & ""
10                mrsTestReason.MoveNext
11            Loop
              
12            .ListIndex = 0
13        End With
          
          
14        Exit Sub
InitData_Error:
15        Call WriteErrLog("zlPublicHisCommLis", "frmTestReason", "ִ��(InitData)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
16        Err.Clear
End Sub

Private Sub txtTestReason_KeyPress(KeyAscii As Integer)
    If InStr("'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call GetPerson(Trim(txtUser.Text))
    Else
        If InStr("'", Chr(KeyAscii)) > 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

