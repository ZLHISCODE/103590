VERSION 5.00
Begin VB.Form frmӦ���λ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ӧ���λ����"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
   Icon            =   "frmӦ���λ.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   1
      Left            =   1650
      MaxLength       =   100
      TabIndex        =   2
      Top             =   450
      Width           =   2355
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   -300
      TabIndex        =   9
      Top             =   1800
      Width           =   5505
   End
   Begin VB.CommandButton cmd�ϼ� 
      Caption         =   "��"
      Enabled         =   0   'False
      Height          =   240
      Left            =   3720
      TabIndex        =   6
      Top             =   1380
      Width           =   255
   End
   Begin VB.OptionButton opt��λ 
      Caption         =   "�����ݺŶ�λ(&N)"
      Height          =   285
      Index           =   1
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Value           =   -1  'True
      Width           =   1725
   End
   Begin VB.OptionButton opt��λ 
      Caption         =   "��ҩƷ��Ӧ�̶�λ(&S)"
      Height          =   285
      Index           =   0
      Left            =   210
      TabIndex        =   3
      Top             =   1020
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   2760
      TabIndex        =   8
      Top             =   2040
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   1410
      TabIndex        =   7
      Top             =   2040
      Width           =   1100
   End
   Begin VB.TextBox txtEdit 
      Enabled         =   0   'False
      Height          =   300
      Index           =   0
      Left            =   1650
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1350
      Width           =   2355
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��ⵥ�ݺ�(&M)"
      Height          =   180
      Index           =   1
      Left            =   450
      TabIndex        =   1
      Top             =   540
      Width           =   1170
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "��Ӧ��(&U)"
      Enabled         =   0   'False
      Height          =   180
      Index           =   0
      Left            =   810
      TabIndex        =   4
      Top             =   1440
      Width           =   810
   End
End
Attribute VB_Name = "frmӦ���λ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mblnOK As Boolean
Dim mstr���ݺ� As String
Dim mstr��Ӧ��ID As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngIndex As Long
    
    For lngIndex = txtEdit.LBound To txtEdit.UBound
        If txtEdit(lngIndex).Enabled = True Then
            If StrIsValid(txtEdit(lngIndex).Text, txtEdit(lngIndex).MaxLength) = False Then
                txtEdit(lngIndex).SetFocus
                Exit Sub
            End If
            
            Select Case lngIndex
                Case 0
                    mstr��Ӧ��ID = txtEdit(lngIndex).Tag
                Case 1
                    mstr���ݺ� = UCase(Trim(txtEdit(lngIndex).Text))
            End Select
        End If
    Next
    
    If mstr���ݺ� = "" And mstr��Ӧ��ID = "" Then
        MsgBox "�����붨λ������", vbInformation, gstrSysName
        Exit Sub
    End If
    
    mblnOK = True
    Unload Me
End Sub

Private Sub cmd�ϼ�_Click()
    Dim rs��Ӧ�� As New ADODB.Recordset
    
    gstrSQL = "Select id,�ϼ�ID,ĩ��,����,����,���� From ҩƷ��Ӧ�� Where " & _
                " nvl(����ʱ��,to_date('3000-01-01','yyyy-MM-dd'))=to_date('3000-01-01','yyyy-MM-dd') " & _
                " start with �ϼ�ID is null connect by prior ID =�ϼ�ID order by level,ID"
    Call OpenRecordset(rs��Ӧ��, Me.Caption)
    
    If rs��Ӧ��.EOF Then
        rs��Ӧ��.Close
        Exit Sub
    End If
    With FrmSelect
        Set .TreeRec = rs��Ӧ��
        .StrNode = "����ҩƷ��Ӧ��"
        .lngMode = 0
        .Show 1, Me
        If .BlnSuccess = True Then
            txtEdit(0).Tag = .CurrentID
            txtEdit(0).Text = .CurrentName
        End If
    End With
    Unload FrmSelect
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlcommfun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mblnOK = False
End Sub

Public Function Get��λ����(str���ݺ� As String, str��Ӧ��ID As String) As Boolean
    
    frmӦ���λ.Show vbModal, frmӦ�����ѯ
    
    Get��λ���� = mblnOK
    If mblnOK = True Then
        str���ݺ� = mstr���ݺ�
        str��Ӧ��ID = mstr��Ӧ��ID
    End If
End Function

Public Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
'����ַ����Ƿ��зǷ��ַ�������ṩ���ȣ��Գ��ȵĺϷ���Ҳ����⡣
    If InStr(strInput, "'") > 0 Then
        MsgBox "���������ݺ��зǷ��ַ���", vbExclamation, gstrSysName
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "���������ݲ��ܳ���" & Int(intMax / 2) & "������" & "��" & intMax & "����ĸ��", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    StrIsValid = True
End Function

Private Sub opt��λ_Click(Index As Integer)
    txtEdit(0).Enabled = opt��λ(0).Value
    lbl(0).Enabled = opt��λ(0).Value
    cmd�ϼ�.Enabled = opt��λ(0).Value
    
    txtEdit(1).Enabled = opt��λ(1).Value
    lbl(1).Enabled = opt��λ(1).Value
    
    txtEdit(Index).SetFocus
End Sub
