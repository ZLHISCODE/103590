VERSION 5.00
Begin VB.Form frmҽ��������� 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҽ�����ı༭"
   ClientHeight    =   2250
   ClientLeft      =   2760
   ClientTop       =   3645
   ClientWidth     =   5595
   Icon            =   "frmҽ���������.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   4260
      TabIndex        =   9
      Top             =   1710
      Width           =   1100
   End
   Begin VB.Frame fra���� 
      Caption         =   "ҽ������"
      Height          =   1905
      Left            =   210
      TabIndex        =   0
      Top             =   150
      Width           =   3855
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   4
         Top             =   780
         Width           =   1755
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1935
         MaxLength       =   5
         TabIndex        =   2
         Top             =   360
         Width           =   765
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   2
         Left            =   1935
         MaxLength       =   20
         TabIndex        =   6
         Top             =   1215
         Width           =   1740
      End
      Begin VB.Image imgװ�� 
         Height          =   240
         Left            =   240
         Picture         =   "frmҽ���������.frx":000C
         Top             =   540
         Width           =   240
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "���ı���(&D)"
         Height          =   180
         Index           =   1
         Left            =   915
         TabIndex        =   3
         Top             =   840
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�������(&S)"
         Height          =   180
         Index           =   0
         Left            =   900
         TabIndex        =   1
         Top             =   420
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��������(&N)"
         Height          =   180
         Index           =   2
         Left            =   930
         TabIndex        =   5
         Top             =   1275
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   4260
      TabIndex        =   8
      Top             =   750
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4260
      TabIndex        =   7
      Top             =   240
      Width           =   1100
   End
End
Attribute VB_Name = "frmҽ���������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enum�༭
    Text��� = 0
    text���� = 1
    Text���� = 2
End Enum

Dim mlng���� As Long           '��ǰ�༭�ı�����������
Dim mstr��� As String         '��ǰ�༭�ı����������
Dim mblnOK As Boolean
Dim mblnChange As Boolean     '�Ƿ�ı���

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    If IsValid() = False Then Exit Sub
    
    MousePointer = vbHourglass
    If Save��������() = False Then
        MousePointer = vbDefault
        Exit Sub
    End If
    MousePointer = vbDefault
    
    mblnOK = True
    mblnChange = False
    
    Unload Me
End Sub

Private Function IsValid() As Boolean
'����:���������йر������ĵ������Ƿ���Ч
'����:
'����ֵ:��Ч����True,����ΪFalse
    Dim lngIndex As Integer
    Dim strTemp As String
    For lngIndex = Text��� To Text����
        If zlCommFun.StrIsValid(Trim(TxtEdit(lngIndex).Text), TxtEdit(lngIndex).MaxLength) = False Then
            TxtEdit(lngIndex).SetFocus
            zlControl.TxtSelAll TxtEdit(lngIndex)
            Exit Function
        End If
        
        If Len(Trim(TxtEdit(lngIndex).Text)) = 0 Then
            TxtEdit(lngIndex).Text = ""
            MsgBox "��Ż����ƶ�����Ϊ�ա�", vbExclamation, gstrSysName
            TxtEdit(lngIndex).SetFocus
            Exit Function
        End If
    Next
    
    If TxtEdit(Text���).Enabled = True Then
        If IsNumeric(TxtEdit(Text���)) = False Or Val(TxtEdit(Text���).Text) <= 0 Then
            MsgBox "���ֻ���Ǵ���900��������", vbExclamation, gstrSysName
            zlControl.TxtSelAll TxtEdit(Text���)
            TxtEdit(Text���).SetFocus
            Exit Function
        End If
    End If
    
    IsValid = True
End Function

Private Function Save��������() As Boolean
'����:����༭�����ݵ��������ı���
'����:
'����ֵ:�ɹ�����True,����ΪFalse
    Dim lng��� As Long
    Dim lst As ListItem
    
    On Error GoTo errHandle
    
    If mstr��� = "" Then     '����һ����¼
        lng��� = TxtEdit(Text���).Text
        gstrSQL = "zl_��������Ŀ¼_Insert(" & mlng���� & "," & lng��� & ",'" & TxtEdit(text����).Text & "','" & TxtEdit(Text����).Text & "')"
    Else                      '�޸�
        gstrSQL = "zl_��������Ŀ¼_Update(" & mlng���� & "," & mstr��� & ",'" & TxtEdit(text����).Text & "','" & TxtEdit(Text����).Text & "')"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '��������������Ӧ�ĵ���
    With frmҽ�����.cmb����
        If mstr��� <> "" Then
            .RemoveItem .ListIndex
            lng��� = Val(mstr���)
        End If
        '����
        .AddItem TxtEdit(text����).Text & "." & TxtEdit(Text����).Text
        .ItemData(.NewIndex) = lng���
        .ListIndex = .NewIndex
    End With
    
    Save�������� = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function �༭��������(ByVal lng���� As Long, ByVal str��� As String) As Boolean
'����:��������õı������Ĺ����ڽ���ͨѶ�ĳ���
'����::lng����           ��ǰ�༭�ı������ĵĵ�����
'      str���           ��ǰ�༭�ı������ĵĵ����
'����ֵ:�༭�ɹ�����True,����ΪFalse
    Dim rs�������� As New ADODB.Recordset
    Dim lng��� As Long
    
    mlng���� = lng����
    mstr��� = str���
    mblnOK = False
    
    rs��������.CursorLocation = adUseClient
    
    If str��� <> "" Then
        gstrSQL = "Select ����,���� From ��������Ŀ¼  Where ���=[1] and ����=[2]"
        Set rs�������� = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, CInt(str���), mlng����)
        
        TxtEdit(Text���).Text = str���
        TxtEdit(text����).Text = rs��������("����")
        TxtEdit(Text����).Text = rs��������("����")
        
        lblEdit(Text���).Enabled = False
        TxtEdit(Text���).Enabled = False
    Else
        lng��� = Val(zlDatabase.GetMax("��������Ŀ¼", "���", 5, " where ����=" & mlng����))
        If lng��� < 1 Then lng��� = 0
        TxtEdit(Text���).Text = lng���
    End If
    
    mblnChange = False
    frmҽ���������.Show vbModal
    �༭�������� = mblnOK
End Function

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll TxtEdit(Index)
    Select Case Index
        Case Text����
          zlCommFun.OpenIme True
        Case Text���, text����
            zlCommFun.OpenIme False
    End Select
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0 'ʹ֮����
        zlCommFun.PressKey (vbKeyTab)
    Else
        If Index = Text��� Then
            KeyAscii = asc(UCase(Chr(KeyAscii)))
            If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Or KeyAscii = 13) Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    zlCommFun.OpenIme False
End Sub
