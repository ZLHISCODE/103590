VERSION 5.00
Begin VB.Form frmIllSortEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������༭"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   Icon            =   "frmIllSortEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   4
      Left            =   1110
      MaxLength       =   50
      TabIndex        =   8
      Top             =   1464
      Width           =   3885
   End
   Begin VB.CommandButton cmd�ϼ� 
      Caption         =   "��"
      Height          =   240
      Left            =   4710
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1920
      Width           =   255
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   0
      Left            =   1110
      MaxLength       =   6
      TabIndex        =   1
      Top             =   195
      Width           =   1395
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "������Чֻ��������(&S)"
      Height          =   195
      Left            =   2850
      TabIndex        =   2
      Top             =   248
      Width           =   2205
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   3
      Left            =   1110
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1890
      Width           =   3885
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   180
      TabIndex        =   14
      Top             =   2520
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   -150
      TabIndex        =   15
      Top             =   2310
      Width           =   5445
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   3870
      TabIndex        =   13
      Top             =   2520
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   2610
      TabIndex        =   12
      Top             =   2520
      Width           =   1100
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   2
      Left            =   1110
      MaxLength       =   20
      TabIndex        =   6
      Top             =   1041
      Width           =   3885
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   1
      Left            =   1110
      MaxLength       =   150
      TabIndex        =   4
      Top             =   618
      Width           =   3885
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "���뷶Χ(&R)"
      Height          =   180
      Index           =   4
      Left            =   60
      TabIndex        =   7
      Top             =   1530
      Width           =   990
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "���(&T)"
      Height          =   180
      Index           =   0
      Left            =   420
      TabIndex        =   0
      Top             =   255
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "�ϼ�(&D)"
      Height          =   180
      Index           =   3
      Left            =   420
      TabIndex        =   9
      Top             =   1950
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����(&J)"
      Height          =   180
      Index           =   2
      Left            =   420
      TabIndex        =   5
      Top             =   1101
      Width           =   630
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "����(&N)"
      Height          =   180
      Index           =   1
      Left            =   420
      TabIndex        =   3
      Top             =   678
      Width           =   630
   End
End
Attribute VB_Name = "frmIllSortEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstrID As String             '��ǰ�༭����ĿID
Dim mstr�ϼ���ĿID As String     '��ǰ�༭���ϼ���ĿID
Dim mstr������� As String

Dim mblnChange As Boolean  '���޸�

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    If IsValid() = False Then Exit Sub
    If Save��Ŀ() = False Then Exit Sub
    
    If mstrID <> "" Then
        mblnChange = False
        Unload Me
        Exit Sub
    End If
    mstrID = ""
    txtEdit(0).Text = ""
    txtEdit(1).Text = ""
    txtEdit(2).Text = ""
    txtEdit(4).Text = ""
    chk����.Value = 0
    txtEdit(0).SetFocus
    mblnChange = False
'    Unload Me
End Sub

Private Function IsValid() As Boolean
'����:��������������������Ƿ���Ч
'����:
'����ֵ:��Ч����True,����ΪFalse
    Dim i As Integer
    Dim strTemp As String
    For i = 0 To 4
        If i <> 3 Then
            strTemp = Trim(txtEdit(i).Text)
            If zlCommFun.StrIsValid(Trim(txtEdit(i).Text), txtEdit(i).MaxLength) = False Then
                txtEdit(i).SetFocus
                zlControl.TxtSelAll txtEdit(i)
                Exit Function
            End If
        End If
    Next
    
    If Not IsNumeric(txtEdit(0).Text) Then
        MsgBox "��������������", vbInformation, gstrSysName
        txtEdit(0).SetFocus
        zlControl.TxtSelAll txtEdit(0)
        Exit Function
    End If
    If Val(txtEdit(0).Text) <= 0 Or Val(txtEdit(0).Text) > 999999 Then
        MsgBox "��Ų���С�ڻ�����㣬��ҪС��1000000��", vbInformation, gstrSysName
        txtEdit(0).SetFocus
        zlControl.TxtSelAll txtEdit(0)
        Exit Function
    End If
    If Val(txtEdit(0).Text) <> Int(txtEdit(0).Text) Then
        MsgBox "��������������", vbInformation, gstrSysName
        txtEdit(0).SetFocus
        zlControl.TxtSelAll txtEdit(0)
        Exit Function
    End If
    If Len(Trim(txtEdit(1).Text)) = 0 Then
        MsgBox "���Ʋ���Ϊ�ա�", vbExclamation, gstrSysName
        txtEdit(1).Text = ""
        txtEdit(1).SetFocus
        Exit Function
    End If
    IsValid = True
End Function

Private Function Save��Ŀ() As Boolean
'����:����༭�����ݵ�����������
'����:
'����ֵ:�ɹ�����True,����ΪFalse
    Dim lng����id As Long
    Dim nodTemp As Node
    On Error GoTo ErrHandle
    
    If mstrID = "" Then       '����һ����¼
        lng����id = zlDatabase.GetNextId("�����������")
        
        gstrSQL = "ZL_�����������_INSERT(" & lng����id & ",'" & mstr�ϼ���ĿID & "'," & txtEdit(0).Text & _
                ",'" & txtEdit(1).Text & "','" & UCase(txtEdit(2).Text) & "','" & txtEdit(4).Text & "','" & mstr������� & "'," & IIF(chk����.Value = 1, 0, 1) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Else    '�޸�
        lng����id = mstrID
        gstrSQL = "ZL_�����������_UPDATE(" & lng����id & ",'" & mstr�ϼ���ĿID & "'," & txtEdit(0).Text & _
                ",'" & txtEdit(1).Text & "','" & UCase(txtEdit(2).Text) & "','" & txtEdit(4).Text & "'," & IIF(chk����.Value = 1, 0, 1) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    
    '���¹�����
    With frmIllManage.tvwMain_S
        If mstrID = "" Then
            '��������
            If mstr�ϼ���ĿID = "" Then
                Set nodTemp = .Nodes.Add(, , "K" & lng����id, "��" & txtEdit(0).Text & "��" & Trim(txtEdit(1).Text), "Root", "Root")
            Else
                Set nodTemp = .Nodes.Add("K" & mstr�ϼ���ĿID, tvwChild, "K" & lng����id, "��" & txtEdit(0).Text & "��" & Trim(txtEdit(1).Text), "Root", "Root")
            End If
        Else
            '�޸ķ���
            Set nodTemp = .Nodes("K" & lng����id)
            nodTemp.Text = "��" & txtEdit(0).Text & "��" & Trim(txtEdit(1).Text)
            
            If mstr�ϼ���ĿID = "" Then
                If Not nodTemp.Parent Is Nothing Then
                    '�ı������
                    Call frmIllManage.FillTree
                End If
            Else
                If Not nodTemp.Parent Is .Nodes("K" & mstr�ϼ���ĿID) Then
                    '�ı������
                    Set nodTemp.Parent = .Nodes("K" & mstr�ϼ���ĿID)
                End If
            End If
        End If
        .Nodes("K" & lng����id).EnsureVisible
    End With
        
    Save��Ŀ = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function �����༭(ByVal str�ϼ���Ŀ As String, ByVal str�ϼ���ĿID As String, _
    ByVal str������� As String, Optional ByVal strID As String = "") As Boolean
'����:��������õı����������ڽ���ͨѶ�ĳ���
'����:str�ϼ���Ŀ     �ϼ�������������
'     str�ϼ���ĿID   �ϼ���������ID
'     str�������     ������������
'     strID           ���������ĵ�ID
'����ֵ:�༭�ɹ�����True,����ΪFalse
    
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    
    mstr������� = str�������
    
    mstrID = strID
    
    On Error GoTo ErrHandle
    If strID <> "" Then
        rsTemp.CursorLocation = adUseClient
        
        gstrSQL = "select A.ID,A.�ϼ�ID,A.����,A.����,A.���뷶Χ,A.���,A.�Ƿ���,B.��� as �ϼ����,B.���� as �ϼ����� " & _
                " from ����������� A,����������� B " & _
                " where B.ID(+)=A.�ϼ�ID and A.ID=[1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(strID))
                
        txtEdit(0).Text = rsTemp("���")
        txtEdit(1).Text = Trim(rsTemp("����"))
        txtEdit(2).Text = IIF(IsNull(rsTemp("����")), "", rsTemp("����"))
        txtEdit(4).Text = IIF(IsNull(rsTemp("���뷶Χ")), "", rsTemp("���뷶Χ"))
        chk����.Value = IIF(rsTemp("�Ƿ���") = 1, 0, 1)
        mstr�ϼ���ĿID = IIF(IsNull(rsTemp("�ϼ�ID")), "", rsTemp("�ϼ�ID"))
        
        If IsNull(rsTemp("�ϼ�����")) Then
            txtEdit(3).Text = "��"
        Else
            txtEdit(3).Text = "��" & rsTemp("�ϼ����") & "��" & Trim(rsTemp("�ϼ�����"))
        End If
        
    Else
        mstr�ϼ���ĿID = str�ϼ���ĿID
        txtEdit(3).Text = str�ϼ���Ŀ
    End If
    
    mblnChange = False
    frmIllSortEdit.Show vbModal
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmd�ϼ�_Click()
    Dim blnRe As Boolean
    Dim str���� As String
    Dim str�ϼ�ID As String
    Dim str���뷶Χ As String
    
    str�ϼ�ID = mstr�ϼ���ĿID
    str���� = txtEdit(3).Text
    blnRe = frmClassSel.ShowTree(str�ϼ�ID, str����, str���뷶Χ, mstr�������, mstrID)
    '�ɹ�����
    If blnRe Then
        '�µı����Ŀ��
        mstr�ϼ���ĿID = str�ϼ�ID
        txtEdit(3).Text = str����
        mblnChange = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk����_Click()
    mblnChange = True
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If Index = 1 Then
        txtEdit(2).Text = zlStr.GetCodeByVB(txtEdit(1).Text)
    ElseIf Index = 2 Then
        txtEdit(2).Text = UCase(txtEdit(2).Text)
    End If
    mblnChange = True
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Then
        'Ҫ���������ƣ����Բ����й��ַ�
        If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    ElseIf Index = 0 Then
        '���ֻ������������
        If InStr("0123456789" & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    ElseIf Index = 4 Then
        If KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
            KeyAscii = KeyAscii - Asc("a") + Asc("A")
        End If
        'ֻ��ȡ��Щ��ĸ
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789.,-" & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    If Index = 1 Then
        zlCommFun.OpenIme True
    End If
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    If Index = 2 Or Index = 0 Or Index = 4 Then
        zlCommFun.OpenIme False
    End If
End Sub
